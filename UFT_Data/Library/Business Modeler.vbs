Option Explicit
Dim iTime
iTime = 240

'#############################################################################################################
'																		BMIDE Function List
'#############################################################################################################
'List of Functions:
'0.   Fn_SISW_BMIDE_GetObject()
'1.   Fn_BMIDE_ExtensionTreeOperations(sAction,sNodeName)
'2.	  Fn_BMIDE_BusinessObjectTreeOperations(sAction,sNodeName)
'3.   Fn_BMIDE_NewBusinessObjectCreate()
'4.   Fn_BMIDE_TabOperations(strTabSection,strAction,strTabName)
'5.   Fn_BMIDE_NewLOVCreate()
'6.   Fn_BMIDE_LOVAttachmentOperations(strAction,strProject,strProperty,strCondition,bOverride)
'7.   Fn_BMIDE_SubLovOperations(sAction,sNavPath,sLovAttachName,sLovAttachCond)
'8.   Fn_BMIDE_LOVTableOperations(strAction,sNavPath,strValue,strDescription,strDisplayName,strCondition)
'9.   Fn_LaunchBMIDE()
'10. Fn_BMIDEWorkspace()
'11. Fn_BMIDE_LoadEnvXML()
'12. Fn_BMIDE_RandNoGenerate()
'13. Fn_BMIDE_MenuOperation()
'14. Fn_BMIDE_ResetPerspective()
'15. Fn_BMIDE_CreateNewProject()
'16. Fn_BMIDE_DeleteObject()
'17. Fn_BMIDE_LOVLocalizationOperations()
'18. Fn_BMIDE_CreateNewCustomProperties()
'19. Fn_BMIDE_CreateNewForm()
'20. Fn_BMIDE_PersistentPropertiesOperation()
'21. Fn_BMIDE_CreateAlternateIDRule()
'22. Fn_BMIDE_AddLOVValueErrorMsgVerify()
'23. Fn_BMIDE_SubLOVErrorMessageVerify()
'24. Fn_BMIDE_ImportProject()
'25. Fn_BMIDE_CloseDialogs()
'26. Fn_BMIDE_ErrorMessageVerify()
'27. Fn_BMIDE_NewLOVErrorMsgVerify()    - Eliminated. Not Used anywhere. By Sushma Pagare [13-Jun-13]
'28. Fn_BMIDE_SetView()
'29. Fn_BMIDE_AddServerConnectionProfile()
'30. Fn_BMIDE_DeployProject()
'31. Fn_BMIDE_LocalizationOperations()
'32. Fn_BMIDE_ExitBMIDE()
'33. Fn_BMIDE_TreeIndexIdentification()
'34. Fn_BMIDE_ServicesOperations()
'35. Fn_BMIDE_InnerTabOperations()
'36. Fn_BMIDE_OperationInputPropertyOperations()
'37. Fn_BMIDE_RuntimePropertiesOperation()
'38. Fn_BMIDE_DisplayRuleOperations()
'39. Fn_BMIDE_PropertyConstantsOperations()
'40. Fn_BMIDE_NamingRuleAttachesOperations()
'41. Fn_BMIDE_CreateNamingRule()
'42. Fn_BMIDE_PrefixErrorMsgVerify()
'43. Fn_BMIDE_AddDatasetReference()
'44. Fn_BMIDE_DatasetToolActionOperations()
'45. Fn_BMIDE_CreateDataset()
'46. Fn_BMIDE_DeepCopyRuleOperations()
'47. Fn_BMIDE_CreateGlobalConstant()
'48. Fn_BMIDE_CreateIDContext()
'49. Fn_BMIDE_NewBussinessContext()
'50. Fn_BMIDE_NewProjectErrorVerify()
'51. Fn_BMIDE_CreateBusinessObjectConstant()
'52. Fn_BMIDE_CreateCondition()
'53. Fn_BMIDE_CreateAliasIdRule()
'54. Fn_BMIDE_ProjectPropertiesOperation()
'55. Fn_BMIDE_TeamcenterRepositoryConnection()
'56. Fn_BMIDE_CreateExtensionDefination()
'57. Fn_BMIDE_ExtensionAvailabilityOperations()
'58. Fn_BMIDE_ExtensionParameterOperation()
'59. Fn_BMIDE_SaveDataModel()
'60. Fn_BMIDE_CreateLibrary()
'61. Fn_BMIDE_DeleteObjectErrorMsgVerify()
'62. Fn_BMIDE_CreateNoteType()
'63. Fn_BMIDE_CreateNewClass()
'64. Fn_BMIDE_CreateServiceLibrary()
'65. Fn_BMIDE_OperationsTreeOperation()
'66. Fn_BMIDE_AddExtensionRule()
'67. Fn_BMIDE_PreConditionTableOperations()
'68. Fn_BMIDE_PostActionTableOperations()
'69. Fn_BMIDE_BusinessObjectConstantTableOperation()
'70. Fn_BMIDE_RestartWorkspace()
'71. Fn_BMIDE_AttributeTableOperatons()
'72. Fn_BMIDE_CreateTemplate()
'73. Fn_BMIDE_CreatePackage()
'74. Fn_BMIDE_ToolbarButtonClick()
'75. Fn_BMIDE_FindObjects()
'76. Fn_BMIDE_RelationPropertiesOperation()
'77. Fn_BMIDE_GRMRuleOperations()
'78. Fn_BMIDE_MasterAlternateIDRulesTableOperations()
'79. Fn_BMIDE_SupplementalAlternateIDRulesTableOperations()
'80. Fn_BMIDE_CompoundPropertiesOperation()
'81. Fn_BMIDE_NewLOVCreateExt()
'82. Fn_BMIDE_ServerConnectionProfileOperations()
'83. Fn_BMIDE_ErrorWindowMsgVerify()
'84. Fn_BMIDE_VerifyLogs()
'85. Fn_BMIDE_CreateOperationalDataProject()
'86. Fn_BMIDE_CreateClassicChange()
'87. Fn_BMIDE_AccessorTableOperation()
'88. Fn_BMIDE_OpsProjectSelect()
'89. Fn_BMIDE_CloseProject()
'90. Fn_BMIDE_DeployProjectWithDetailVerifications()
'91. Fn_CommandOperations()
'92. Fn_BMIDE_ColorOperations()
'93. Fn_BMIDE_IncorporateLatestOperationalDataChanges()
'94 Fn_BMIDE_ApplicationExtensionRule()
'95 Fn_BMIDE_ApplicationExtensionPoint()
'96 Fn_BMIDE_RevisionNamingRuleAttach()
'97 Fn_BMIDE_CreateRevisionNamingRule()
'98 Fn_BMIDE_CreateGlobalConstantExt()
'99 Fn_BMIDE_CreateNewTool()
'100 Fn_BMIDE_CreateNewStorageMedia()
'101 Fn_BMIDE_GlobalConstantTableOperations()
'102 Fn_BMIDE_CreateBusinessObjectConstantExt()
'103 Fn_BMIDE_SetPerspective()
'104 Fn_BMIDE_PatternTableOperations()
'105 Fn_BMIDE_NamingRuleChangeIDTableOperations()
'106 Fn_BMIDE_NamingRuleRevIDTableOperations()
'107 Fn_BMIDE_RestartWarmServers()
'108 Fn_BMIDE_RevisionNamingRuleAttachmentsOparation()
'109 Fn_BMIDE_GetTemplatePrefix()
'110 Fn_BMIDE_CreateDependentTemplateProject()
'111 Fn_BMIDE_DeleteProject()
'112 Fn_CreateNewIRDC()
'113 Fn_BMIDE_CreateNewFunctionality()
'114 Fn_BMIDE_DeepCopyRuleOperationsExt()
'115 Fn_BMIDE_OperationInputPropertyTableOperations()
'116 Fn_SISW_BMIDE_ConvertBusinessObject()
'117 Fn_SISW_BMIDE_NamingRuleAttachmentsOperations()
'118 Fn_SISW_BMIDE_NewEventType()
'119 Fn_SISW_BMIDE_NewEventTypeMapping()
'120 Fn_SISW_BMIDE_CreateAuditDefinationProperty()
'121 Fn_SISW_BMIDE_NewAuditDefinition()
'122 Fn_SISW_BMIDE_JavaTreeGetItemPath()
'122 Fn_SISW_BMIDE_AuditExtensionsTableOperations()
'123 Fn_SISW_BMIDE_JavaTreeGetItemPathExt()
'124 Fn_SISW_BMIDE_AuditTypeMappingTableOperations()
'125 Fn_SISW_BMIDE_LOVAttachesOfPropertyTreetableOperations()
'126 Fn_SISW_BMIDE_HelpContentsTreeOperations()
'127 Fn_SISW_BMIDE_ConfigureBusinessObjectsToDisplayOperations()
'128 Fn_SISW_BMIDE_FilterConfigurationOperations()
'129 Fn_SISW_BMIDE_CreateBatchLOV()
'130 Fn_SISW_BMIDE_BuildQueryClauseOperations()
'131 Fn_SISW_BMIDE_NewDynamicLOVQueryCriteriaTableOperations()
'132 Fn_SISW_BMIDE_NewDynamicLOVOperations()
'133 Fn_SISW_BMIDE_CloseAllDialogsOperations()
'134 Fn_SISW_BMIDE_ToolbarButtonOperations()
'135 Fn_SISW_BMIDE_FilterAttributesTableOperations()
'136 Fn_SISW_BMIDE_AddNewModelElementOperations()
'137 Fn_SISW_BMIDE_DynamicLOVTestResultsTableOperations()
'138 Fn_SISW_BMIDEUIimprove_ToolMenuConfig_Basics()
'139 Fn_SISW_BMIDE_HTMLReportOperations()
'140 Fn_BMIDE_PushTemplatetoReferenceDir()
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'
''Function Name		 	:	Fn_SISW_BMIDE_GetObject
'
''Description		    :  	Function to get Object hierarchy

''Parameters		    :	1. sObjectName : Object Handle name
								
''Return Value		    :  	Object \ Nothing
'
''Examples		     	:	Fn_SISW_BMIDE_GetObject("NewEventType")

'History:
'	Developer Name			Date			Rev. No.		Reviewer		Changes Done	
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'	Avinash Jagdale 	 11-Aug-2012		1.0				
'-----------------------------------------------------------------------------------------------------------------------------------
'	Ashwini Kumar		 25-Sept-2013		2.0								Externalized object hierarchies
'-----------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_BMIDE_GetObject(sObjectName)
	Dim sObjectXMLPath
	sObjectXMLPath = Fn_GetEnvValue("User", "AutomationDir") & "\TestData\AutomationXML\ObjectXML\Business Modeler.xml"
	Set Fn_SISW_BMIDE_GetObject = Fn_SISW_Setup_GetObjectFromXML(sObjectXMLPath, sObjectName)
End Function
'#############################################################################################################

'Function Name		:				Fn_BMIDE_ExtensionTreeOperations()
'Description			:				Operation Related to Extension Tree
'Return Value			: 				TRUE \ FALSE\Index(InCase of GetIndex)
'Pre-requisite			:		 		BMIDE Prespective is Open.
'Examples				:			
'											Call Fn_BMIDE_ExtensionTreeOperations("Select","AutoTest1:LOV")
'											Call Fn_BMIDE_ExtensionTreeOperations("Expand","AutoTest1:LOV")
'											Call Fn_BMIDE_ExtensionTreeOperations("Collapse","AutoTest1:LOV")
'											Call Fn_BMIDE_ExtensionTreeOperations("Exist","AutoTest1:LOV:BillCodes")
'											Call Fn_BMIDE_ExtensionTreeOperations("GetIndex","AutoTest1:LOV:BillCodes")
'											Call Fn_BMIDE_ExtensionTreeOperations("PopupMenuSelect","AutoTest1:LOV","Search Conditions")
'											Call Fn_BMIDE_ExtensionTreeOperations("ExpandAll","AutoTest1:LOV","")
'History:
'			Developer Name		Date		Rev. No.	Changes Done			Reviewer					Build
'		---------------------------------------------------------------------------------------------------------------------------------------------
'			Rupali 				19/10/2010       1.0				Harshal					Harshal									Tc8.3(2010091600a)
'			pranav 				18/11/2010       1.0				Modified Case "PopupMenuSelect" 		Tc8.3(2010091600a)
'			Sandeep 		06/12/2010       1.0				Added Case "ExpandAll" 
'		---------------------------------------------------------------------------------------------------------------------------------------------
Function Fn_BMIDE_ExtensionTreeOperations(sAction,sNodeName,sPopupMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_ExtensionTreeOperations"
	Dim objJavaTreeExt, intNodeCount, intCount, sTreeItem, iLen, iCounter, iIndex, iTotal, sResult, arr
	Dim aMenuList,aNodeName,StrMenu,sName,WshShell
	Select Case sAction
		'----------------------------------------------------------------------- For selecting single node -------------------------------------------------------------------------
		Case "Select"
                    Call Fn_JavaTree_Select("Fn_BMIDE_ExtensionTreeOperations", JavaWindow("Business Modeler"), "Extension Tree",sNodeName)
					Fn_BMIDE_ExtensionTreeOperations = TRUE
		'----------------------------------------------------------------------- For expanding a particular  node-------------------------------------------------------------------------
		Case "Expand"
                    Call Fn_UI_JavaTree_Expand("Fn_BMIDE_ExtensionTreeOperations",JavaWindow("Business Modeler"),"Extension Tree",sNodeName)
					Fn_BMIDE_ExtensionTreeOperations = TRUE
		'----------------------------------------------------------------------- For collapssing a particular  node-------------------------------------------------------------------------
		Case "Collapse"
                    Call Fn_UI_JavaTree_Collapse("Fn_BMIDE_ExtensionTreeOperations", JavaWindow("Business Modeler"),"Extension Tree",sNodeName)
					Fn_BMIDE_ExtensionTreeOperations = TRUE
		'----------------------------------------------------------------------- For Checking existance of a particular  node-------------------------------------------------------------------------
		Case "Exist","ExistExt"
	            If sAction="ExistExt" Then
					sNodeName=Fn_SISW_BMIDE_JavaTreeGetItemPathExt(JavaWindow("Business Modeler").JavaTree("Extension Tree"),sNodeName)
					Fn_BMIDE_ExtensionTreeOperations = sNodeName				
				ElseIf sAction="Exist" Then
					Set objJavaTreeExt = Fn_UI_ObjectCreate( "Fn_BMIDE_ExtensionTreeOperations",  JavaWindow("Business Modeler").JavaTree("Extension Tree"))
					intNodeCount = objJavaTreeExt.GetROProperty ("items count") 
					For intCount = 0 to intNodeCount - 1
						sTreeItem = objJavaTreeExt.GetItem(intCount)
						If Trim(lcase(sTreeItem)) = Trim(Lcase(sNodeName)) Then
							Fn_BMIDE_ExtensionTreeOperations = TRUE	
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[" + sNodeName + "] of JavaTree exist")
							Exit For
						End If
					Next
					If Cstr(intCount) = intNodeCount Then
						Fn_BMIDE_ExtensionTreeOperations = FALSE						
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[" + sNodeName + "] of JavaTree does not exist")
						Exit Function
					End If
				End if
		'----------------------------------------------------------------------- Get Index value of a particular node-------------------------------------------------------------------------
		Case "GetIndex"
				bFlag = False
				For intCount=0 to JavaWindow("Business Modeler").JavaTree("Extension Tree").GetROProperty ("items count")-1
					sTreeItem = JavaWindow("Business Modeler").JavaTree("Extension Tree").GetItem (intCount)
					If Trim(lcase(sTreeItem)) = Trim(Lcase(sNodeName)) Then
						Fn_BMIDE_ExtensionTreeOperations = intCount
						bFlag = True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The index of the given node is "&intCount)
						Exit For
					End If
				Next
				If bFlag = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The given node does not exist")
					Fn_BMIDE_ExtensionTreeOperations = FALSE
					Exit Function
                 End If
		'----------------------------------------------------------------------- PopUp Menu Select.-------------------------------------------------------------------------		
		Case "PopupMenuSelect","PopupMenuSelectExt"

					If sAction="PopupMenuSelectExt" Then
						sNodeName=Fn_SISW_BMIDE_JavaTreeGetItemPathExt(JavaWindow("Business Modeler").JavaTree("Extension Tree"),sNodeName)
					End If
					
					'Build the Popup menu to be selected
					aMenuList = split(sPopupMenu, ":",-1,1)
					intCount = Ubound(aMenuList)
					Call Fn_JavaTree_Select("Fn_BMIDE_ExtensionTreeOperations", JavaWindow("Business Modeler"), "Extension Tree",sNodeName)
					Call Fn_UI_JavaTree_OpenContextMenu("Fn_BMIDE_ExtensionTreeOperations", JavaWindow("Business Modeler"), "Extension Tree",sNodeName)

					Wait(2)
					'Select Menu action
					Select Case intCount
						Case "0"
							 StrMenu = JavaWindow("Business Modeler").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0))
						Case "1"
							StrMenu = JavaWindow("Business Modeler").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1))
						Case "2"
							StrMenu = JavaWindow("Business Modeler").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1),aMenuList(2))
						Case Else
							Fn_BMIDE_ExtensionTreeOperations = FALSE
							Exit Function
					End Select

				    If JavaWindow("Business Modeler").WinMenu("ContextMenu").GetItemProperty(StrMenu,"enabled") Then
						JavaWindow("Business Modeler").WinMenu("ContextMenu").Select StrMenu
					Else 
						Fn_BMIDE_ExtensionTreeOperations = FALSE	
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "ContextMenu is disable.")
						Exit Function
					End If 

					Fn_BMIDE_ExtensionTreeOperations = TRUE		

		Case "PopupMenuExist"
		
					'Build the Popup menu to be selected
					aMenuList = split(sPopupMenu, ":",-1,1)
					intCount = Ubound(aMenuList)
					Call Fn_JavaTree_Select("Fn_BMIDE_ExtensionTreeOperations", JavaWindow("Business Modeler"), "Extension Tree",sNodeName)
					Call Fn_UI_JavaTree_OpenContextMenu("Fn_BMIDE_ExtensionTreeOperations", JavaWindow("Business Modeler"), "Extension Tree",sNodeName)

					Wait(2)
					'Select Menu action
					Select Case intCount
						Case "0"
							 StrMenu = JavaWindow("Business Modeler").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0))
						Case "1"
							StrMenu = JavaWindow("Business Modeler").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1))
						Case "2"
							StrMenu = JavaWindow("Business Modeler").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1),aMenuList(2))
						Case Else
							Fn_BMIDE_ExtensionTreeOperations = FALSE
							Exit Function
					End Select
					sResult= JavaWindow("Business Modeler").WinMenu("ContextMenu").GetItemProperty(StrMenu,"Exists")
					'Added by Sandeep
					Set WshShell = CreateObject("WScript.Shell")
					wait(1)
					WshShell.SendKeys "{ESC}"
					wait(2)
					Set WshShell =Nothing 
					' - - - - - - - - - - - - - - - -
                    If sResult= true Then
                    Else 
						Fn_BMIDE_ExtensionTreeOperations = FALSE	
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "ContextMenu is disable.")
						Exit Function
					End If 

					Fn_BMIDE_ExtensionTreeOperations = TRUE	
							
		Case "DoubleClick","DoubleClickExt"

					If sAction="DoubleClickExt" Then
						sNodeName=Fn_SISW_BMIDE_JavaTreeGetItemPathExt(JavaWindow("Business Modeler").JavaTree("Extension Tree"),sNodeName)
					End If
					JavaWindow("Business Modeler").JavaTree("Extension Tree").Activate sNodeName
					Fn_BMIDE_ExtensionTreeOperations = TRUE

		Case "ExpandAll"
				Call Fn_BMIDE_ExtensionTreeOperations("PopupMenuSelect",sNodeName,"Navigate:Expand Selection")
				Fn_BMIDE_ExtensionTreeOperations = TRUE

		Case "CollapseAll"
				Call Fn_BMIDE_ExtensionTreeOperations("PopupMenuSelect",sNodeName,"Navigate:Collapse Selection")
				Fn_BMIDE_ExtensionTreeOperations = TRUE

		Case Else
						Fn_BMIDE_ExtensionTreeOperations = FALSE
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_BMIDE_ExtensionTreeOperations function failed")
						Exit Function

End Select
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" Sucessfully completed on Node [" + sNodeName + "] of JavaTree of function Fn_BMIDE_ExtensionTreeOperations")
	Set objJavaWindowCat = nothing
	Set objJavaTreeCat = nothing
End Function
'#############################################################################################################
'Function Name		:				Fn_BMIDE_BusinessObjectTreeOperations()
'Description			:				Operation Related to Business Object Tree
'Return Value			: 				TRUE \ FALSE
'Pre-requisite			:		 		BMIDE Prespective is Open.
'Examples				:			
'								Call Fn_BMIDE_BusinessObjectTreeOperations("Select","AutoTest1:BusinessObject","")
'								Call Fn_BMIDE_BusinessObjectTreeOperations("Expand","AutoTest1:BusinessObject","")
'								Call Fn_BMIDE_BusinessObjectTreeOperations("Collapse","AutoTest1:BusinessObject","")
'								Call Fn_BMIDE_BusinessObjectTreeOperations("Exist","AutoTest1:BusinessObject","")
'								Call Fn_BMIDE_BusinessObjectTreeOperations("PopupMenuSelect","AutoTest1:BusinessObject:POM_object","Open")
'								Call Fn_BMIDE_BusinessObjectTreeOperations("ExpandAll","Trial1","")
'								"PopupMenuState":- This Case Return True If PopUp Menu is Enabled and False If PopUp Menu is Disabled
'								Call Fn_BMIDE_BusinessObjectTreeOperations("PopupMenuState","DemoBatch24811:BusinessObject:POM_object","New Business Object...")
'								Call Fn_BMIDE_BusinessObjectTreeOperations("PopupMenuState","DemoBatch24811:BusinessObject:POM_object","Organize:Move to Extension File...")
'								Call Fn_BMIDE_BusinessObjectTreeOperations("PopupMenuExist","DemoBatch24811:BusinessObject:POM_object","Organize:Move to Extension File...")
'History:
'			Developer Name		Date		Rev. No.	Changes Done			Reviewer						Build
'		---------------------------------------------------------------------------------------------------------------------------------------------
'			Rupali 				19/10/2010       1.0				Harshal					Harshal									Tc8.3(2010091600a)
'			pranav 				18/11/2010       1.0				Modified Case   "PopupMenuSelect" 		Tc8.3(2010091600a)
'			pranav 				22/11/2010       1.0				Added Case    "DoubleClick" 		
'			Sandeep 		26/11/2010       1.0				Added Case    "ExpandAll" 		
'			Sandeep 		14/1/2011       1.0				Added Case    "ExistBussinessObject" 		
'			Sandeep 		25/6/2013       1.0				Added Case    "PopupMenuExist" 		
'		---------------------------------------------------------------------------------------------------------------------------------------------
'#############################################################################################################

Function Fn_BMIDE_BusinessObjectTreeOperations(sAction,sNodeName,sPopupMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_BusinessObjectTreeOperations"
	Dim objJavaTreeExt, intNodeCount, intCount, sTreeItem, iLen, iCounter, iIndex, iTotal, sResult, arr
	Dim WshShell

	Select Case sAction
		'----------------------------------------------------------------------- For selecting single node -------------------------------------------------------------------------
		Case "Select"
                    Call Fn_JavaTree_Select("Fn_BMIDE_BusinessObjectTreeOperations", JavaWindow("Business Modeler"), "BusinessObjectTree",sNodeName)
					Fn_BMIDE_BusinessObjectTreeOperations = TRUE
		'----------------------------------------------------------------------- For expanding a particular  node-------------------------------------------------------------------------
		Case "Expand"
                    Call Fn_UI_JavaTree_Expand("Fn_BMIDE_BusinessObjectTreeOperations",JavaWindow("Business Modeler"),"BusinessObjectTree",sNodeName)
					Fn_BMIDE_BusinessObjectTreeOperations = TRUE
		'----------------------------------------------------------------------- For collapssing a particular  node-------------------------------------------------------------------------
		Case "Collapse"
                    Call Fn_UI_JavaTree_Collapse("Fn_BMIDE_BusinessObjectTreeOperations", JavaWindow("Business Modeler"),"BusinessObjectTree",sNodeName)
					Fn_BMIDE_BusinessObjectTreeOperations = TRUE
		'----------------------------------------------------------------------- For Checking existance of a particular  node-------------------------------------------------------------------------
		Case "Exist"
				Set objJavaTreeExt = Fn_UI_ObjectCreate( "Fn_BMIDE_BusinessObjectTreeOperations",  JavaWindow("Business Modeler").JavaTree("BusinessObjectTree"))
					intNodeCount = objJavaTreeExt.GetROProperty ("items count") 
					For intCount = 0 to intNodeCount - 1
						sTreeItem = objJavaTreeExt.GetItem(intCount)
						If Trim(lcase(sTreeItem)) = Trim(Lcase(sNodeName)) Then
							Fn_BMIDE_BusinessObjectTreeOperations = TRUE	
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[" + sNodeName + "] of JavaTree exist")
							Exit For
						End If
					Next
					If Cstr(intCount) = intNodeCount Then
						Fn_BMIDE_BusinessObjectTreeOperations = FALSE						
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[" + sNodeName + "] of JavaTree does not exist")
						Exit Function
					End If
		Case "ExistBusinessObject"
				Set objJavaTreeExt = Fn_UI_ObjectCreate( "Fn_BMIDE_BusinessObjectTreeOperations",  JavaWindow("Business Modeler").JavaTree("BusinessObjectTree"))
					intNodeCount = objJavaTreeExt.GetROProperty ("items count") 
					For intCount = 0 to intNodeCount - 1


						If intCount>4 Then
							Fn_BMIDE_BusinessObjectTreeOperations = FALSE						
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[" + sNodeName + "] of JavaTree does not exist")
							Exit Function
						End If
						If Trim(lcase(sTreeItem)) = Trim(Lcase(sNodeName)) Then
							Fn_BMIDE_BusinessObjectTreeOperations = TRUE	
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[" + sNodeName + "] of JavaTree exist")
							Exit For
						End If
					Next
					If Cstr(intCount) = intNodeCount Then
						Fn_BMIDE_BusinessObjectTreeOperations = FALSE						
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[" + sNodeName + "] of JavaTree does not exist")
						Exit Function
					End If

		'----------------------------------------------------------------------- PopUp Menu Select.-------------------------------------------------------------------------		
		Case "PopupMenuSelect"
					Dim aMenuList,aNodeName,StrMenu,sName
					'Build the Popup menu to be selected
					aMenuList = split(sPopupMenu, ":",-1,1)
					intCount = Ubound(aMenuList)
					Call Fn_JavaTree_Select("Fn_BMIDE_BusinessObjectTreeOperations", JavaWindow("Business Modeler"), "BusinessObjectTree",sNodeName)
					Call Fn_UI_JavaTree_OpenContextMenu("Fn_BMIDE_BusinessObjectTreeOperations", JavaWindow("Business Modeler"), "BusinessObjectTree",sNodeName)

                    'Select Menu action
					Select Case intCount
						Case "0"
							 StrMenu = JavaWindow("Business Modeler").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0))
						Case "1"
							StrMenu = JavaWindow("Business Modeler").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1))
						Case "2"
							StrMenu = JavaWindow("Business Modeler").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1),aMenuList(2))
						Case Else
							Fn_BMIDE_ExtensionTreeOperations = FALSE
							Exit Function
					End Select
				    If JavaWindow("Business Modeler").WinMenu("ContextMenu").GetItemProperty(StrMenu,"enabled") Then
						JavaWindow("Business Modeler").WinMenu("ContextMenu").Select StrMenu
					Else 
						Fn_BMIDE_BusinessObjectTreeOperations = FALSE	
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "ContextMenu is disable.")
						Exit Function
					End If 
					Fn_BMIDE_BusinessObjectTreeOperations = TRUE
		Case "DoubleClick"
					JavaWindow("Business Modeler").JavaTree("BusinessObjectTree").Activate sNodeName
					Fn_BMIDE_BusinessObjectTreeOperations = TRUE
		Case "ExpandAll"
				Call Fn_BMIDE_BusinessObjectTreeOperations("PopupMenuSelect",sNodeName,"Navigate:Expand Selection")
				Fn_BMIDE_BusinessObjectTreeOperations = TRUE
		Case "CollapseAll"
				Call Fn_BMIDE_BusinessObjectTreeOperations("PopupMenuSelect",sNodeName,"Navigate:Collapse Selection")
				Fn_BMIDE_BusinessObjectTreeOperations = TRUE
		Case "PopupMenuState"
					'Build the Popup menu to be selected
					aMenuList = split(sPopupMenu, ":",-1,1)
					intCount = Ubound(aMenuList)
					Call Fn_JavaTree_Select("Fn_BMIDE_BusinessObjectTreeOperations", JavaWindow("Business Modeler"), "BusinessObjectTree",sNodeName)
					Call Fn_UI_JavaTree_OpenContextMenu("Fn_BMIDE_BusinessObjectTreeOperations", JavaWindow("Business Modeler"), "BusinessObjectTree",sNodeName)

                    'Select Menu action
					Select Case intCount
						Case "0"
							 StrMenu = JavaWindow("Business Modeler").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0))
						Case "1"
							StrMenu = JavaWindow("Business Modeler").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1))
						Case "2"
							StrMenu = JavaWindow("Business Modeler").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1),aMenuList(2))
						Case Else
							Fn_BMIDE_ExtensionTreeOperations = FALSE
							Exit Function
					End Select
					
					Set WshShell = CreateObject("WScript.Shell")
				    If JavaWindow("Business Modeler").WinMenu("ContextMenu").GetItemProperty(StrMenu,"enabled") Then
						wait(2)
						WshShell.SendKeys "{ESC}"
						Fn_BMIDE_BusinessObjectTreeOperations = TRUE	
					Else 
						wait(2)
						WshShell.SendKeys "{ESC}"
						Fn_BMIDE_BusinessObjectTreeOperations = FALSE	
					End If
					set WshShell =Nothing
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "PopupMenuExist"
					'Build the Popup menu to be selected
					aMenuList = split(sPopupMenu, ":",-1,1)
					intCount = Ubound(aMenuList)
					Call Fn_JavaTree_Select("Fn_BMIDE_BusinessObjectTreeOperations", JavaWindow("Business Modeler"), "BusinessObjectTree",sNodeName)
					Call Fn_UI_JavaTree_OpenContextMenu("Fn_BMIDE_BusinessObjectTreeOperations", JavaWindow("Business Modeler"), "BusinessObjectTree",sNodeName)

                    'Select Menu action
					Select Case intCount
						Case "0"
							 StrMenu = JavaWindow("Business Modeler").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0))
						Case "1"
							StrMenu = JavaWindow("Business Modeler").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1))
						Case "2"
							StrMenu = JavaWindow("Business Modeler").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1),aMenuList(2))
						Case Else
							Fn_BMIDE_BusinessObjectTreeOperations = FALSE
							Exit Function
					End Select

					Set WshShell = CreateObject("WScript.Shell")
				    If  JavaWindow("Business Modeler").WinMenu("ContextMenu").GetItemProperty(StrMenu,"Exists") Then
						wait(2)
						WshShell.SendKeys "{ESC}"
						Fn_BMIDE_BusinessObjectTreeOperations = TRUE	
					Else 
						wait(2)
						WshShell.SendKeys "{ESC}"
						Fn_BMIDE_BusinessObjectTreeOperations = FALSE	
					End If
					set WshShell =Nothing		
		Case Else
						Fn_BMIDE_BusinessObjectTreeOperations = FALSE
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_BMIDE_BusinessObjectTreeOperations function failed")
						Exit Function
End Select
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" Sucessfully completed on Node [" + sNodeName + "] of JavaTree of function Fn_BMIDE_BusinessObjectTreeOperations")
	Set objJavaWindowCat = nothing
	Set objJavaTreeCat = nothing
End Function
'-------------------------------------------------------------------Function Used to Create New Bussiness Object----------------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_NewBusinessObjectCreate

'Description			 :	Function Used to Create New Bussiness Object

'Parameters			   :	1.strProject: Projetct Name
										'2.strName:Name Of Object
										'3.strDisplayName: Display Name Of object
										'4.strParent:Parent type
										'5.bAdvance:Advance Option
										'6.bPriObject:CreatePrimaryBusinessObject Option
										'7.bUninstantiable: Uninstantiable Option
										'8.strProperties=Properties

'Return Value		   : 	True Or False

'Pre-requisite			:	Should be Log In BMIDE

'Examples				: Call Fn_BMIDE_NewBusinessObjectCreate("","MyTestObject","Test Object","","This Is Test Object","Off","","","")

'										to add Custom properties 
										' to add 1 propety
'										strprop= "pname1:pDispName1:pDesc1:String:32:"""":"""":"""":"""":"""":"""":"""""
										' to add  2 propetyies
'										strprop= "pname1~pname2:pDispName1~pDispName2:pDesc1~pDesc2:String~String:32~32:""""~"""":""""~"""":""""~"""":""""~"""":""""~"""":""""~"""":""""~"""":""""~"""""

'											properties are seperated by '  : '   
											
'										Fn_BMIDE_NewBusinessObjectCreate("","MyTestObject1","Test Object1","","This Is Test Object","","","",strprop)

'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done								Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				16/11/2010			           1.0																								Sunny R
'													Pranav Ingle										   			 19/11/2010			           1.0								Modified 												Sunny R
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function Fn_BMIDE_NewBusinessObjectCreate(strProject,strName,strDisplayName,strParent,strDesc,bAdvance,bPriObject,bUninstantiable,strProperties)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_NewBusinessObjectCreate"
   Dim ObjNewBsnWindow
   Dim bFlag,strPrototype,sName,arrReferenceClass,strDefultDisplayName
   Dim  iCounter,bReturn,arrMainString,arrName,arrDisplayName,arrDescription,arrAttributeType,arrStringLength,arrSetNull ,arrInitialValue,arrLowerBound,arrUpperBound,arrArray,arrKeys
  'Function Return False
   Fn_BMIDE_NewBusinessObjectCreate=False
   'Checking Existance of NewBusinessObject window
   If Fn_UI_ObjectExist("Fn_BMIDE_NewBusinessObjectCreate", JavaWindow("Business Modeler").JavaWindow("NewBusinessObject"))=True Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Pass:NewBusinessObject Window is  Exist ")
  Else
	   'If NewBusinessObject window not exist then function will Exit
       Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail:NewBusinessObject Window is not Exist ")
	   Exit Function
   End If
	'Creating Object of NewBusinessObject window
	Set ObjNewBsnWindow=Fn_UI_ObjectCreate("Fn_BMIDE_NewBusinessObjectCreate",JavaWindow("Business Modeler").JavaWindow("NewBusinessObject"))
	If strProject<>"" Then
		'Verifying Project is exist in Project List Or Not
		bFlag=Fn_UI_ListItemExist("Fn_BMIDE_NewBusinessObjectCreate", ObjNewBsnWindow, "Project",strProject)
		If bFlag=False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail:["+strProject+"] is not present in Project List")
			Set ObjNewBsnWindow=Nothing
			Exit Function
		End If
		'Selecting Project from Project List
		Call Fn_List_Select("Fn_BMIDE_NewBusinessObjectCreate", ObjNewBsnWindow, "Project",strProject)
	End If
	If strName<>"" Then
		strPrototype= Fn_Edit_Box_GetValue("Fn_BMIDE_NewBusinessObjectCreate",ObjNewBsnWindow,"Name")
		strName=strPrototype+strName
		'Setting Name to New Business Object
        Call Fn_Edit_Box("Fn_BMIDE_NewBusinessObjectCreate",ObjNewBsnWindow,"Name",strName)
		'Commented code for Design change introduced due to PR : 6898697
'		strDefultDisplayName=Fn_Edit_Box_GetValue("Fn_BMIDE_NewBusinessObjectCreate",ObjNewBsnWindow,"DisplayName")
'		'Verifying Display Name Get Populated Automatically with Correct Value
'		If Trim(strDefultDisplayName)<>Trim(strName) Then
'			Set ObjNewBsnWindow=Nothing
'			Exit Function
'		End If
		If strDisplayName="" Then
			strDisplayName=strName
		End If
	End If

	If strDisplayName<>"" Then
		'Setting Display Name to New Business Object
        Call Fn_Edit_Box("Fn_BMIDE_NewBusinessObjectCreate",ObjNewBsnWindow,"DisplayName",strDisplayName)
	End If
	If strParent<>"" Then
		'Setting Parent to New Business Object
        Call Fn_Edit_Box("Fn_BMIDE_NewBusinessObjectCreate",ObjNewBsnWindow,"Parent",strParent)
	End If
	If strDesc<>"" Then
		'Setting Description to New Business Object
        Call Fn_Edit_Box("Fn_BMIDE_NewBusinessObjectCreate",ObjNewBsnWindow,"Description",strDesc)
	End If
	If bAdvance<>"" Then
		'Setting Advance Option
       ' Call Fn_CheckBox_Set("Fn_BMIDE_NewBusinessObjectCreate", ObjNewBsnWindow, "Advanced", bAdvance)
		If bPriObject<>"" Then
			'Setting CreatePrimaryBusinessObject Option
			Call Fn_CheckBox_Set("Fn_BMIDE_NewBusinessObjectCreate", ObjNewBsnWindow, "CreatePrimaryBusinessObj", bPriObject)
		End If
		If bUninstantiable<>"" Then
			'Setting Uninstantiable Option
			Call Fn_CheckBox_Set("Fn_BMIDE_NewBusinessObjectCreate", ObjNewBsnWindow, "Uninstantiable", bUninstantiable)
		End If
	End If

	If strProperties<>"" Then
        arrMainString = Split(strProperties, ":")
		arrName = Split(arrMainString(0),"~")
		arrDisplayName = Split(arrMainString(1),"~")
		arrDescription = Split(arrMainString(2),"~")
		arrAttributeType = Split(arrMainString(3),"~")
		arrStringLength = Split(arrMainString(4),"~")
		arrReferenceClass = Split(arrMainString(5),"~")
        arrSetNull = Split(arrMainString(6),"~")
		arrInitialValue = Split(arrMainString(7),"~")
		arrLowerBound = Split(arrMainString(8),"~")
		arrUpperBound = Split(arrMainString(9),"~")
		arrArray = Split(arrMainString(10),"~")
		arrKeys = Split(arrMainString(11),"~")
		For iCounter = 0 to UBound(arrName) 
			Call Fn_Button_Click("Fn_BMIDE_NewBusinessObjectCreate", JavaWindow("Business Modeler").JavaWindow("NewBusinessObject"), "Add")
			bReturn= Fn_BMIDE_CreateNewCustomProperties(arrName(iCounter),arrDisplayName(iCounter),arrDescription(iCounter),arrAttributeType(iCounter),arrStringLength(iCounter),arrReferenceClass(iCounter),arrSetNull(iCounter),arrInitialValue(iCounter),arrLowerBound(iCounter),arrUpperBound(iCounter),arrArray(iCounter),arrKeys(iCounter))
			If bReturn= true Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Pass:Successfully Added New property of Name ["+arrName(iCounter)+"]")
			Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail :failed to Add New property of Name ["+arrName(iCounter)+"]")
					Fn_BMIDE_NewBusinessObjectCreate= false
					Exit Function
			End If
		Next
		'Need To write code to Add Properties
	End If
	If Fn_UI_ObjectExist("Fn_BMIDE_NewBusinessObjectCreate",ObjNewBsnWindow.JavaButton("Next"))=True Then
		'Clicking On Next Button
'		Call Fn_Button_Click("Fn_BMIDE_NewBusinessObjectCreate", ObjNewBsnWindow, "Next")
		JavaWindow("Business Modeler").JavaWindow("NewBusinessObject").JavaButton("Next").Object.click
        wait 1
		If strDisplayName<>"" Then
			'Setting Display Name to New Business Object
			Call Fn_Edit_Box("Fn_BMIDE_NewBusinessObjectCreate",ObjNewBsnWindow,"DisplayName",strDisplayName+"Revision")
		End If
	End If
	'Clicking On Finish Button create New Business Object
'	Call Fn_Button_Click("Fn_BMIDE_NewBusinessObjectCreate", ObjNewBsnWindow, "Finish")
	JavaWindow("Business Modeler").JavaWindow("NewBusinessObject").JavaButton("Finish").Object.click
	wait 1
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Pass:Successfully Create New Business Object of Display Name ["+strDisplayName+"]")
	'Function returns True after creating New Business Object
	Fn_BMIDE_NewBusinessObjectCreate=True
	'Releasing Object of NewBusinessObject window
	Set ObjNewBsnWindow=Nothing
End Function

'-------------------------------------------------------------------Function Used to Perform Operations On Tab----------------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_TabOperations

'Description			 :	Function Used to Perform Operations On JavaTab Which presents on BMIDE perspective

'Parameters			   :	1.strTabSection: Section Of Tab 
										'2.strAction:Action to Perform
										'3.strTabName: Tab Name on which have to perform the operations

'Return Value		   : 	True Or False

'Pre-requisite			:	Should be Log In BMIDE

'Examples				: 	Call Fn_BMIDE_TabOperations("UpperLeft","Activate","Business Objects")
'										Call Fn_BMIDE_TabOperations("Main","Activate","T2_Testobject")
'										Call Fn_BMIDE_TabOperations("UpperLeft","Close","Business Objects")
'										Call Fn_BMIDE_TabOperations("ParentTab","VerifyActivate","Business Objects")			 -	Pranav/	 26-Jun-2013
'										bReturn=Fn_BMIDE_TabOperations("UpperLeft","Exist","Business Objects~Classes")
'										bReturn=Fn_BMIDE_TabOperations("UpperLeft","Exist","BMIDE")
				
'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				16/11/2010			           1.0																			Sunny R
'													Sandeep N										   				26/11/2010			           1.0							Case "Close"	   						Sunny R
'													Pranav Ingle										   			26/06/2013			           1.1							Case "VerifyActivate"				Sandeep
'													Sandeep N										   				28/06/2013			           1.2							Case "Exist"	   						Avinash J
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_BMIDE_TabOperations(strTabSection,strAction,strTabName)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_TabOperations"
   'Variable Declarations
   Dim ObjBusinessModelerWindow
   Dim strExePath, strWorkspace
   Dim bReturn
	'This Select Case to Define Section of Tab
   Select Case strTabSection
	 	'This Case to Perform Operations On Upper Lefts Corner Tabs
	 	Case "LowerLeft"	'Fn_BMIDE_TabOperations("UpperLeft","Activate","Classes")
				Set ObjBusinessModelerWindow=Fn_UI_ObjectCreate("Fn_BMIDE_TabOperations",JavaWindow("Business Modeler"))
				strTab="EOCTab"
		'This Case to Perform Operations On Lower Left Corner Tabs
		Case "UpperLeft"	'Fn_BMIDE_TabOperations("LowerLeft","Activate","Outline")
				Set ObjBusinessModelerWindow=Fn_UI_ObjectCreate("Fn_BMIDE_TabOperations",JavaWindow("Business Modeler"))
				strTab="BCNTab"
		'This Case to Perform Operations On Main window Tabs
		Case "Main"
				Set ObjBusinessModelerWindow=Fn_UI_ObjectCreate("Fn_BMIDE_TabOperations",JavaWindow("Business Modeler"))
				strTab="MainTab"	'Fn_BMIDE_TabOperations("Main","Activate","Item")
		Case "ParentTab"
				Set ObjBusinessModelerWindow=Fn_UI_ObjectCreate("Fn_BMIDE_TabOperations",JavaWindow("Business Modeler"))
				strTab="ParentTab"	'Fn_BMIDE_TabOperations("ParentTab","VerifyActivate","Item")
		Case Else
			Fn_BMIDE_TabOperations=False
   End Select
   
   'Need to handle launching of BMIDE - as it gets closed abruptly durinf batch execution
    If ObjBusinessModelerWindow.Exist(5) = FALSE Then
   		strExePath = Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\BMIDEConfig\BMIDE_Config.xml", "BMIDEExecutable")
		strWorkspace = Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\BMIDEConfig\BMIDE_Config.xml", "BMIDEWorkspacePath")
   		bReturn =  Fn_LaunchBMIDE(strExePath, strWorkspace)
   		wait(5)
    End If
   
   Select Case strAction
	 	'This Case to Activate Tabs
		 Case "Activate"	'Fn_BMIDE_TabOperations("Main","Activate","Item")
				Call Fn_UI_JavaTab_Select("Fn_BMIDE_TabOperations",ObjBusinessModelerWindow,strTab, strTabName)
                Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Successfully Activted the Tab [" + strTabName +"]")
				Fn_BMIDE_TabOperations=True
		Case "Close"	'Fn_BMIDE_TabOperations("UpperLeft","Close","Business Objects")
				 'Call Fn_UI_JavaTab_Select("Fn_BMIDE_TabOperations",ObjBusinessModelerWindow,strTab, strTabName)
				 JavaWindow("Business Modeler").JavaTab(strTab).CloseTab strTabName
                Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Successfully Activted the Tab [" + strTabName +"]")
				Fn_BMIDE_TabOperations=True
		Case "VerifyActivate"
				crrActiveTab=Fn_UI_Object_GetROProperty("Fn_BMIDE_TabOperations",ObjBusinessModelerWindow.JavaTab(strTab),"value")
				If Trim(crrActiveTab)=Trim(strTabName) Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Successfully verified Tab [" + strTabName +"] is currently activated")
					Fn_BMIDE_TabOperations=True
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fail to verify Tab [" + strTabName +"] is currently activated")
					Fn_BMIDE_TabOperations=False
				End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to check existance of tab
        Case "Exist"
					Set ObjTab=ObjBusinessModelerWindow.JavaTab(strTab)
					aTabName=Split(strTabName,"~")
                    For iCounter=0 to ubound(aTabName)
						bFlag=False
						For iCount=0 to cint(ObjTab.GetROProperty("items count"))-1
							If trim(ObjTab.Object.getItem(iCount).text)=trim(aTabName(iCounter)) Then
								bFlag=True
								Exit for
							End If
						Next
						If bFlag=False Then
							Exit for
						End If
					Next
					If bFlag=True Then
						Fn_BMIDE_TabOperations=True
					Else
						Fn_BMIDE_TabOperations=False
					End If
		Case Else
			Fn_BMIDE_TabOperations=False
   End Select
   'releasing Object of Business Modeler Java Window
   	Set ObjBusinessModelerWindow=Nothing
End Function 

'-------------------------------------------------------------------Function Used to Create New LOV----------------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_NewLOVCreate

'Description			 :	Function Used to Create New List Of Value

'Parameters			   :	1.strProject: Projetct Name
										'2.strName:Name Of LOV
										'3.strDesc: LOV Description
										'4.strType:LOV type
										'5.bUsage:Usage Option
										'6.strReference:LOV Reference
										'7.strLower: Lower Boundry of LOV
										'8.strUpper:Upper Boundry of LOV
										'9.bCascdingView:Cascding View Option
										'10.arrProperties: Properties

'Return Value		   : 	True Or False

'Pre-requisite			:	Should be Log In BMIDE

'Examples				: abc="LA~LA~Los Engelis~isTrue:USA~United States~Unite States~"
'									  Call Fn_BMIDE_NewLOVCreate("","TestLOV4","First Test LOV Object7","ListOfValuesString","Exhaustive","","","","OFF",abc)

'										Fn_BMIDE_NewLOVCreate("","TestLOV4","First Test LOV Object7","ListOfValuesFilter","Exhaustive","Based On LOV~D3MyLOV1_07759","","","","","A~C")
'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				16/11/2010			           1.0																					Sunny R
'													Pranav Ingle										   			 22/11/2010			           1.0									Modified								Sunny R
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_BMIDE_NewLOVCreate(strProject,strName,strDesc,strType,bUsage,strReference,strLower,strUpper,bCascdingView,arrProperties,strValues)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_NewLOVCreate"
   Dim ObjNewLOVWindow
   Dim bFlag,strPrototype,arrPropValues,arrPropPairs,iCounter,iCount
  'Function Return False
   Fn_BMIDE_NewLOVCreate=False
   'Checking Existance of NewLOVObject window
   If Fn_UI_ObjectExist("Fn_BMIDE_NewLOVCreate",JavaWindow("Business Modeler").JavaWindow("NewLOV"))=False Then
	   'If NewLOVObject window not exist then function will Exit
       Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail:NewLOVObject Window is not Exist ")
	   Exit Function
   End If
	'Creating Object of NewLOVObject window
	Set  ObjNewLOVWindow=Fn_UI_ObjectCreate("Fn_BMIDE_NewLOVCreate",JavaWindow("Business Modeler").JavaWindow("NewLOV"))

	Call Fn_UI_JavaRadioButton_SetON("Fn_BMIDE_NewLOVCreate",ObjNewLOVWindow, "ClassicLOV")
	Call Fn_Button_Click("Fn_BMIDE_NewLOVCreate", ObjNewLOVWindow, "Next")

	If strProject<>"" Then
		'Verifying Project is exist in Project List Or Not
		bFlag=Fn_UI_ListItemExist("Fn_BMIDE_NewLOVCreate", ObjNewLOVWindow, "Project",strProject)
		If bFlag=False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail:["+strProject+"] is not present in Project List")
			Set ObjNewLOVWindow=Nothing
			Exit Function
		End If
		'Selecting Project from Project List
		Call Fn_List_Select("Fn_BMIDE_NewLOVCreate", ObjNewLOVWindow, "Project",strProject)
	End If
	If strName<>"" Then
		strPrototype= Fn_Edit_Box_GetValue("Fn_BMIDE_NewLOVCreate",ObjNewLOVWindow,"Name")
		strName=strPrototype+strName
		'Setting Name to New LOV Object
        Call Fn_Edit_Box("Fn_BMIDE_NewLOVCreate",ObjNewLOVWindow,"Name",strName)
	End If
	If strDesc<>"" Then
		'Setting Description to New LOV Object
        Call Fn_Edit_Box("Fn_BMIDE_NewLOVCreate",ObjNewLOVWindow,"Description",strDesc)
	End If
	If strType<>"" Then
		'Setting Type to New LOV Object
        'Call Fn_Edit_Box("Fn_BMIDE_NewLOVCreate",ObjNewLOVWindow,"Type",strType)
		Call Fn_Button_Click("Fn_BMIDE_NewLOVCreate", ObjNewLOVWindow,"BrowseLOVType")
		Call Fn_Edit_Box("Fn_BMIDE_NewLOVCreate",JavaWindow("Business Modeler").JavaWindow("Find Business Object"),"Project",strType)
		Call Fn_Button_Click("Fn_BMIDE_NewLOVCreate",JavaWindow("Business Modeler").JavaWindow("Find Business Object"),"OK")
	End If
	If bUsage<>"" Then
		Call Fn_UI_Object_SetTOProperty("Fn_BMIDE_NewLOVCreate",ObjNewLOVWindow.JavaRadioButton("Usage"),"attached text",bUsage)
		'Setting Usage to New LOV Object
        Call Fn_UI_JavaRadioButton_SetON("Fn_BMIDE_NewLOVCreate",ObjNewLOVWindow, "Usage")
	End If
	If strReference<>"" Then
		arrPropValues = Split(strReference,"~")
		If UBound(arrPropValues) > 0 Then
				Call Fn_Edit_Box("Fn_BMIDE_NewLOVCreate",ObjNewLOVWindow,"BasedOnLOV",arrPropValues(1))
				arrPropPairs = Split(strValues,"~")
				For iCounter = 0 to UBound(arrPropPairs)
		'				JavaWindow("Business Modeler").JavaWindow("NewLOV").JavaTree("Reference").SetItemState "A",micChecked 
						ObjNewLOVWindow.JavaTree("Reference").SetItemState arrPropPairs(iCounter),micChecked
				Next
				Call Fn_List_Select("Fn_BMIDE_NewLOVCreate",ObjNewLOVWindow,"Show","Selected")
				'JavaWindow("Business Modeler").JavaWindow("NewLOV").JavaList("Show").Select "Selected"
		Else
			'Setting Reference to New LOV Object
			Call Fn_Edit_Box("Fn_BMIDE_NewLOVCreate",ObjNewLOVWindow,"Reference",strReference)
		End If
	End If
	If strLower<>"" Then
		'Setting Lower Range Of  New LOV Object
        Call Fn_Edit_Box("Fn_BMIDE_NewLOVCreate",ObjNewLOVWindow,"Lower",strLower)
	End If
	If strUpper<>"" Then
		'Setting Upper Range Of New LOV Object
        Call Fn_Edit_Box("Fn_BMIDE_NewLOVCreate",ObjNewLOVWindow,"Upper",strUpper)
	End If
	If bCascdingView<>"" Then
		'Setting Cascading Option
        Call Fn_CheckBox_Set("Fn_BMIDE_NewLOVCreate", ObjNewLOVWindow, "ShowCascadingView", bCascdingView)
	End If
	If arrProperties<>"" Then
        arrPropValues=Split(arrProperties,":")	
		For iCounter=0 To Ubound( arrPropValues)
			Call Fn_Button_Click("Fn_BMIDE_NewLOVCreate", ObjNewLOVWindow, "Add")
			 If Fn_UI_ObjectExist("Fn_BMIDE_NewLOVCreate",JavaWindow("Business Modeler").JavaWindow("AddLOVValue"))=True Then
			arrPropPairs=Split(arrPropValues(iCounter),"~")
			For iCount=0 To Ubound(arrPropPairs)
				If arrPropPairs(0)<>"" Then
					 Call Fn_Edit_Box("Fn_BMIDE_NewLOVCreate",JavaWindow("Business Modeler").JavaWindow("AddLOVValue"),"Value",arrPropPairs(0))				
				End If
				If arrPropPairs(1)<>"" Then
					'Call Fn_Edit_Box("Fn_BMIDE_NewLOVCreate",JavaWindow("Business Modeler").JavaWindow("AddLOVValue"),"ValueDisplayName",arrPropPairs(1))
					JavaWindow("Business Modeler").JavaWindow("AddLOVValue").JavaEdit("ValueDisplayName").Set arrPropPairs(1)
				End If
				If arrPropPairs(2)<>"" Then
					Call Fn_Edit_Box("Fn_BMIDE_NewLOVCreate",JavaWindow("Business Modeler").JavaWindow("AddLOVValue"),"Description",arrPropPairs(2))			
				End If
				If arrPropPairs(3)<>"" Then
					If (trim(Fn_SISW_UI_JavaEdit_Operations("Fn_BMIDE_NewLOVCreate", "GetText", JavaWindow("Business Modeler").JavaWindow("AddLOVValue"),"Condition", "")) <> trim(arrPropPairs(3))) Then
						Call Fn_Edit_Box("Fn_BMIDE_NewLOVCreate",JavaWindow("Business Modeler").JavaWindow("AddLOVValue"),"Condition",arrPropPairs(3))				
					End  If
				End If
				Call Fn_Button_Click("Fn_BMIDE_NewLOVCreate", JavaWindow("Business Modeler").JavaWindow("AddLOVValue"), "Finish")
				Exit For
			Next
			End If
		Next
	End If
    'Clicking On Finish Button create New LOV Object
	Call Fn_Button_Click("Fn_BMIDE_NewLOVCreate", ObjNewLOVWindow, "Finish")
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Pass:Successfully Create New LOV Object of Display Name ["+strName+"]")
	'Function returns True after creating New LOV Object
	Fn_BMIDE_NewLOVCreate=True
	'Releasing Object of NewLOVObject window
	Set ObjNewLOVWindow=Nothing
End Function

'-------------------------------------------------------------------Function Used to Perform Operations Attachments Table For LOV----------------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_LOVAttachmentOperations

'Description			 :	Function Used to Perform Operations Attachments Table For LOV

'Parameters			   :	1.strAction: Action to Perform 
										'2.strProject:Project Name
										'3.strProperty: Property Name
										'4.strCondition:Condition
										'5.bOverride: Override Option

'Return Value		   : 	True Or False

'Pre-requisite			:	Should be Log In BMIDE

'Examples				: 	'1. Call Fn_BMIDE_LOVAttachmentOperations("Attach","","AbsOccName.absocc_attr_name","","")
'										'2. Call Fn_BMIDE_LOVAttachmentOperations("Verify","","AbsOccName.abs_attr_name","","")
'										 3. Call Fn_BMIDE_LOVAttachmentOperations("EditAttachDescription","","AbsOccName.absocc_attr_value~occ_name","","")
'										4.Call Fn_BMIDE_LOVAttachmentOperations("EditAttach","","AbsOccName.absocc_attr_value~occ_name","","")
'											"PropertyName on which have to perform Attach Description ~ Attach Description Property"

'										5. Fn_BMIDE_LOVAttachmentOperations("EditVerify","","ItemRevision.object_desc","eng~chn","English~Chinese")
'											only for this case 		strCondition  - LOV values seperated by ~
'																					bOverride-   LOV  Description  (LOV description should be in sequence as LOV values)

'										6. Fn_BMIDE_LOVAttachmentOperations("Detach","","L2_identifier.l2_p1","","")
'										7.Fn_BMIDE_LOVAttachmentOperations("EditPropertyVerify","","d3Type:d3Material","","")
'										8.Call Fn_BMIDE_LOVAttachmentOperations("Select","","AbsOccName.absocc_attr_name","","")
'										9.Call Fn_BMIDE_LOVAttachmentOperations("EditAttachExt","","absocc_attr_value:abc_value~occ_name","","")
'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				16/11/2010			           1.0																						Sunny R
'													Sandeep N										   				17/11/2010			           1.0								'EditAttachDescription'	   		   Sunny R 
'													Pranav Ingle										   			18/11/2010			           1.0								'EditVerify'						   		   Sunny R 
'													Pranav Ingle										   			18/11/2010			           1.0								'Detach'							   		   Sunny R 
'													Sandeep N										   			23/11/2010			           1.0								'EditAttach'							   		   Sunny R 
'													Sandeep N										   			23/11/2010			           1.0								'Expand'							   		   Sunny R 
'													Sandeep N										   			29/11/2010			           1.0								'EditPropertyVerify'							   		   Sunny R 
'													Sandeep N										   			29/11/2010			           1.0								'EditAttachExt'							   		   Sunny R 
'													Nitish B													20/07/2015					   1.0								Case "Attach"							VivekA																								
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_BMIDE_LOVAttachmentOperations(strAction,strProject,strProperty,strCondition,bOverride)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_LOVAttachmentOperations"
  Dim bFlag,iItemCount,iCounter,iCount,strProp,arrProperty,arrIntrProperty,arrValue,arrDescription,strValueName,strDescValue
   Dim ObjAttachWindow,WshShell
   Fn_BMIDE_LOVAttachmentOperations=False
   bFlag=False
   Select Case strAction
		 	Case "Detach"
					'Selecting Property For Editing From LOV attachments tree 
					 Call Fn_JavaTree_Select("Fn_BMIDE_LOVAttachmentOperations", JavaWindow("Business Modeler"), "LOVAttachments",strProperty)
					'Clicking on Detach Button 
					Call Fn_Button_Click("Fn_BMIDE_LOVAttachmentOperations", JavaWindow("Business Modeler"),"DetachLOVAttachment")
					 Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Pass:Successfully Detach the attachment of Property [ "+strProperty+" ]")        
					Fn_BMIDE_LOVAttachmentOperations= true
			Case "Expand"
					'Selecting Property For Editing From LOV attachments tree 
					 Call Fn_UI_JavaTree_Expand("Fn_BMIDE_LOVAttachmentOperations", JavaWindow("Business Modeler"), "LOVAttachments",strProperty)
					 Fn_BMIDE_LOVAttachmentOperations=True
			Case "Select"
					'To Select LOV attachments tree Node
					 Call Fn_JavaTree_Select("Fn_BMIDE_LOVAttachmentOperations", JavaWindow("Business Modeler"), "LOVAttachments",strProperty)
					 Fn_BMIDE_LOVAttachmentOperations=True
		 	Case "Attach"
				 If Fn_UI_ObjectExist("Fn_BMIDE_LOVAttachmentOperations",JavaWindow("Business Modeler").JavaWindow("PropertyAttachment"))=False Then
					Call Fn_Button_Click("Fn_BMIDE_LOVAttachmentOperations",JavaWindow("Business Modeler"),"Attach")
				 End If
				Set ObjAttachWindow=Fn_UI_ObjectCreate("Fn_BMIDE_LOVAttachmentOperations",JavaWindow("Business Modeler").JavaWindow("PropertyAttachment"))
				If strProject<>"" Then
					bFlag=Fn_UI_ListItemExist("Fn_BMIDE_LOVAttachmentOperations", ObjAttachWindow,"Project",strProject)
					If bFlag=False Then
						Set ObjAttachWindow=Nothing
                        Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail: [ "+strProject+" ] is not Exist in Projects List")        
						Exit Function
					End If
					Call Fn_List_Select("Fn_BMIDE_LOVAttachmentOperations", ObjAttachWindow,"Project",strProject)
				End If
				
				If strProperty<>"" Then  'TC112-2015070100-VivekA-Porting-Changed code as per discussion with Samir Thosar
					arrProperty=Split(strProperty,".")
					Call Fn_Button_Click("Fn_BMIDE_LOVAttachmentOperations",JavaWindow("Business Modeler").JavaWindow("PropertyAttachment"),"Browse")	
					Call Fn_Edit_Box("Fn_BMIDE_LOVAttachmentOperations",JavaWindow("Business Modeler").JavaWindow("PropertyAttachment2"),"Project",arrProperty(0))
					Call Fn_Edit_Box("Fn_BMIDE_LOVAttachmentOperations",JavaWindow("Business Modeler").JavaWindow("PropertyAttachment2"),"Properties",arrProperty(1))
					Call Fn_Button_Click("Fn_BMIDE_LOVAttachmentOperations",JavaWindow("Business Modeler").JavaWindow("PropertyAttachment2"),"OK")	
				End If
				
				If strCondition<>"" Then
					Call Fn_UI_Object_SetTOProperty("Fn_BMIDE_LOVAttachmentOperations",JavaWindow("Business Modeler").JavaButton("BrowseCondition"),"enabled","1")
					Call Fn_Button_Click("Fn_BMIDE_LOVAttachmentOperations",JavaWindow("Business Modeler"),"BrowseCondition")
					Call Fn_UI_EditBox_Type("Fn_BMIDE_LOVAttachmentOperations",JavaWindow("Business Modeler").JavaWindow("FindAttachmentCondition"),"Condition",strCondition)
					Call Fn_Button_Click("Fn_BMIDE_LOVAttachmentOperations",JavaWindow("Business Modeler").JavaWindow("FindAttachmentCondition"),"OK")
				End If
				If  bOverride<>"" Then
					Call Fn_CheckBox_Set("Fn_BMIDE_LOVAttachmentOperations",ObjAttachWindow, "Override", bOverride)
				End If
				Call Fn_Button_Click("Fn_BMIDE_LOVAttachmentOperations",ObjAttachWindow,"Finish")
				Fn_BMIDE_LOVAttachmentOperations=True
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Pass:Successfully attach the attachment of Property [ "+strProperty+" ]")        
			Case "Verify"
				iItemCount=JavaWindow("Business Modeler").JavaTree("LOVAttachments").GetROProperty("items count")
				For iCounter=0 To iItemCount-1
					strProp=JavaWindow("Business Modeler").JavaTree("LOVAttachments").GetItem(iCounter)
					If  strProp=strProperty Then
						Fn_BMIDE_LOVAttachmentOperations=True
						Exit For
					End If
				Next
            Case "EditAttachDescription"
				'Spliting Property Parameter 
				arrProperty=Split(strProperty,"~")
				'Selecting Property For Editing From LOV attachments tree 
				Call Fn_JavaTree_Select("Fn_BMIDE_LOVAttachmentOperations", JavaWindow("Business Modeler"), "LOVAttachments",arrProperty(0))
				'Clicking on Edit Button to Edit Attach Description
				'Call Fn_Button_Click("Fn_BMIDE_LOVAttachmentOperations", JavaWindow("Business Modeler"),"EditAttachments")
				Call Fn_SISW_UI_JavaButton_Operations("Fn_BMIDE_LOVAttachmentOperations", "Object.click", JavaWindow("Business Modeler"), "EditAttachments")
				'Checking Existance of InterdependentLOVAttachment Window
				 If Fn_UI_ObjectExist("Fn_BMIDE_LOVAttachmentOperations",JavaWindow("Business Modeler").JavaWindow("InterdependentLOVAttachment"))=False Then
					Exit Function
				 End If
				 arrIntrProperty=Split(arrProperty(0),".")
				 'Selecting Property From InterdependentLOVTree
				Call Fn_JavaTree_Select("Fn_BMIDE_LOVAttachmentOperations",JavaWindow("Business Modeler").JavaWindow("InterdependentLOVAttachment"), "InterdependentTypeTree",arrIntrProperty(1))
				'Clicking on AttachDescription button to Attach Property Description
				Call Fn_Button_Click("Fn_BMIDE_LOVAttachmentOperations",JavaWindow("Business Modeler").JavaWindow("InterdependentLOVAttachment"),"AttachDescription")
				'Checking Existance of AttachmentProperties Window
				 If Fn_UI_ObjectExist("Fn_BMIDE_LOVAttachmentOperations",JavaWindow("Business Modeler").JavaWindow("AttachmentProperties"))=False Then
					Exit Function
				 End If
				 'Setting Description
				Call Fn_Edit_Box("Fn_BMIDE_LOVAttachmentOperations",JavaWindow("Business Modeler").JavaWindow("AttachmentProperties"),"AttachmentsProperty",arrProperty(1))
				'Clicking OK button to set Description
				Call Fn_Button_Click("Fn_BMIDE_LOVAttachmentOperations",JavaWindow("Business Modeler").JavaWindow("AttachmentProperties"),"OK")
				'Clicking On Finish button To Finish Attachment Description Process
				Call Fn_Button_Click("Fn_BMIDE_LOVAttachmentOperations",JavaWindow("Business Modeler").JavaWindow("InterdependentLOVAttachment"),"Finish")
				Fn_BMIDE_LOVAttachmentOperations=True

				Case "EditAttach"
						'Spliting Property Parameter 
						arrProperty=Split(strProperty,"~")
						'Selecting Property For Editing From LOV attachments tree 
						Call Fn_JavaTree_Select("Fn_BMIDE_LOVAttachmentOperations", JavaWindow("Business Modeler"), "LOVAttachments",arrProperty(0))
						'Clicking on Edit Button to Edit Attach Description
						Call Fn_Button_Click("Fn_BMIDE_LOVAttachmentOperations", JavaWindow("Business Modeler"),"EditAttachments")
						'Checking Existance of InterdependentLOVAttachment Window
						 If Fn_UI_ObjectExist("Fn_BMIDE_LOVAttachmentOperations",JavaWindow("Business Modeler").JavaWindow("InterdependentLOVAttachment"))=False Then
							Exit Function
						 End If
						 arrIntrProperty=Split(arrProperty(0),".")
						 'Selecting Property From InterdependentLOVTree
						Call Fn_JavaTree_Select("Fn_BMIDE_LOVAttachmentOperations",JavaWindow("Business Modeler").JavaWindow("InterdependentLOVAttachment"), "InterdependentTypeTree",arrIntrProperty(1))
						'Clicking on AttachDescription button to Attach Property Description
						Call Fn_Button_Click("Fn_BMIDE_LOVAttachmentOperations",JavaWindow("Business Modeler").JavaWindow("InterdependentLOVAttachment"),"Attach")
						'Checking Existance of AttachmentProperties Window
						 If Fn_UI_ObjectExist("Fn_BMIDE_LOVAttachmentOperations",JavaWindow("Business Modeler").JavaWindow("AttachmentProperties"))=False Then
							Exit Function
						 End If
						 'Setting Description
						Call Fn_Edit_Box("Fn_BMIDE_LOVAttachmentOperations",JavaWindow("Business Modeler").JavaWindow("AttachmentProperties"),"AttachmentsProperty",arrProperty(1))
						'Clicking OK button to set Description
						Call Fn_Button_Click("Fn_BMIDE_LOVAttachmentOperations",JavaWindow("Business Modeler").JavaWindow("AttachmentProperties"),"OK")
						'Clicking On Finish button To Finish Attachment Description Process
						Call Fn_Button_Click("Fn_BMIDE_LOVAttachmentOperations",JavaWindow("Business Modeler").JavaWindow("InterdependentLOVAttachment"),"Finish")
						Fn_BMIDE_LOVAttachmentOperations=True

				Case "EditAttachExt"
						arrProperty=Split(strProperty,"~")
						Call Fn_Button_Click("Fn_BMIDE_LOVAttachmentOperations", JavaWindow("Business Modeler"),"EditAttachments")
						 'Selecting Property From InterdependentLOVTree
						Call Fn_JavaTree_Select("Fn_BMIDE_LOVAttachmentOperations",JavaWindow("Business Modeler").JavaWindow("InterdependentLOVAttachment"), "InterdependentTypeTree",arrProperty(0))
						'Clicking on AttachDescription button to Attach Property Description
						Call Fn_Button_Click("Fn_BMIDE_LOVAttachmentOperations",JavaWindow("Business Modeler").JavaWindow("InterdependentLOVAttachment"),"Attach")
						'Checking Existance of AttachmentProperties Window
						 If Fn_UI_ObjectExist("Fn_BMIDE_LOVAttachmentOperations",JavaWindow("Business Modeler").JavaWindow("AttachmentProperties"))=False Then
							Exit Function
						 End If
						 'Setting Description
						Call Fn_Edit_Box("Fn_BMIDE_LOVAttachmentOperations",JavaWindow("Business Modeler").JavaWindow("AttachmentProperties"),"AttachmentsProperty",arrProperty(1))
						'Clicking OK button to set Description
						Call Fn_Button_Click("Fn_BMIDE_LOVAttachmentOperations",JavaWindow("Business Modeler").JavaWindow("AttachmentProperties"),"OK")
						'Clicking On Finish button To Finish Attachment Description Process
						Call Fn_Button_Click("Fn_BMIDE_LOVAttachmentOperations",JavaWindow("Business Modeler").JavaWindow("InterdependentLOVAttachment"),"Finish")
						Fn_BMIDE_LOVAttachmentOperations=True

				Case "EditVerify"
					'Spliting Property Parameter 	
					arrProperty=Split(strProperty,".")
					'Spliting LOV value and LOV Description Parameter 											
					arrValue=Split(strCondition,"~")			
					arrDescription=Split(bOverride,"~")
					'Selecting Property For Editing From LOV attachments tree 
					Call Fn_JavaTree_Select("Fn_BMIDE_LOVAttachmentOperations", JavaWindow("Business Modeler"), "LOVAttachments",arrProperty(0)+"."+arrProperty(1))
					'Clicking on Edit Button to Edit Attach Description
					Call Fn_Button_Click("Fn_BMIDE_LOVAttachmentOperations", JavaWindow("Business Modeler"),"EditAttachments")
					'Checking Existance of InterdependentLOVAttachment Window
					 If Fn_UI_ObjectExist("Fn_BMIDE_LOVAttachmentOperations",JavaWindow("Business Modeler").JavaWindow("InterdependentLOVAttachment"))=False Then
						 Fn_BMIDE_LOVAttachmentOperations= False
						Exit Function
					 End If

					'Selecting Property For Editing From LOV attachments tree 
					Call Fn_JavaTree_Select("Fn_BMIDE_LOVAttachmentOperations", JavaWindow("Business Modeler").JavaWindow("InterdependentLOVAttachment"), "InterdependentTypeTree",arrProperty(1))
					
					iItemCount=Fn_UI_Object_GetROProperty("Fn_BMIDE_LOVAttachmentOperations",JavaWindow("Business Modeler").JavaWindow("InterdependentLOVAttachment").JavaTree("InterdependentLOVTree") ,"items count")
					For iCount=0 To ubound(arrValue)
							bFlag=False
							For iCounter=0 To iItemCount-1
									strValueName=JavaWindow("Business Modeler").JavaWindow("InterdependentLOVAttachment").JavaTree("InterdependentLOVTree").GetItem(iCounter)
									If strValueName=arrValue(iCount) Then
											strDescValue=JavaWindow("Business Modeler").JavaWindow("InterdependentLOVAttachment").JavaTree("InterdependentLOVTree").GetColumnValue(arrValue(iCount),"Description")
											If  arrDescription(iCount)=strDescValue Then
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"pass : property  [ "+arrDescription(iCount) +" ] is Exist for LOV value ["+arrValue(iCount)+"]")  
													bFlag=True
													Exit For
											End If
									End If
							Next
							If bFlag=False Then
									Exit For
							End If
					Next
					'Clicking On Finish button To Finish Attachment Description Process
					Call Fn_Button_Click("Fn_BMIDE_LOVAttachmentOperations",JavaWindow("Business Modeler").JavaWindow("InterdependentLOVAttachment"),"Finish")
					If bFlag=True Then
							Fn_BMIDE_LOVAttachmentOperations=True
					End If
			Case "EditPropertyVerify"
				'To use this case user have to select Object Name.Property Explicitly Using Select Case
					Call Fn_Button_Click("Fn_BMIDE_LOVAttachmentOperations", JavaWindow("Business Modeler"),"EditAttachments")
					iItemCount=Fn_UI_Object_GetROProperty("Fn_BMIDE_LOVAttachmentOperations",JavaWindow("Business Modeler").JavaWindow("InterdependentLOVAttachment").JavaTree("InterdependentTypeTree") ,"items count")
					For iCounter=0 To iItemCount-1
						strValueName=JavaWindow("Business Modeler").JavaWindow("InterdependentLOVAttachment").JavaTree("InterdependentTypeTree").GetItem(iCounter)
						If Trim(strProperty)=Trim(strValueName) Then
							Fn_BMIDE_LOVAttachmentOperations=True
							Exit For
						End If
					Next
					Call Fn_Button_Click("Fn_BMIDE_LOVAttachmentOperations",JavaWindow("Business Modeler").JavaWindow("InterdependentLOVAttachment"),"Cancel")
			Case Else        
				Fn_BMIDE_LOVAttachmentOperations=False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail:Incorrect Case name")
   End Select
   Set ObjAttachWindow=Nothing
End Function


'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_SubLovOperations

'Description			 :	Function Used to perform operations on Sub LOV

'Parameters			   :   '1.sAction = Action to Select
'										2. sNavPath = LOVTable tree Path
'										3. sLovAttachName = name of LOV to be attach
'										4. sLovAttachCond = Condition of Lov  to be attached with	

'Return Value		   : 	True Or False

'Pre-requisite			:	LOV Tab Should be appear on srceen

'Examples				:	 '  Fn_BMIDE_SubLovOperations("Select","A:D4MY_LOV2","","")
										'	Fn_BMIDE_SubLovOperations("Expand","A:D4MY_LOV2","","")
										'	Fn_BMIDE_SubLovOperations("VerifyNode","A:D4MY_LOV2","","")
										'	Fn_BMIDE_SubLovOperations("ExpandAll","","","")
										'	Fn_BMIDE_SubLovOperations("CollapseAll","","","")
										'	Fn_BMIDE_SubLovOperations("RemSubLov","A:D4MY_LOV2","","")
										'	Fn_BMIDE_SubLovOperations("AddSubLov","A","D4MY_LOV2","isTrue")
'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Pranav Ingle						   							22-Nov-2010											1.0																						Sunny
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------	

Function Fn_BMIDE_SubLovOperations(sAction,sNavPath,sLovAttachName,sLovAttachCond)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_SubLovOperations"
	 Dim objBMIDE,intNodeCount, intCount, sTreeItem,objLovAttach
	 ' Create objet of BMIDE
	 Set objBMIDE=JavaWindow("Business Modeler")
	Set objLovAttach = objBMIDE.JavaWindow("SubLOVAttachment")
		 ' Set ShowCascadingView checkbox on 
		Call Fn_CheckBox_Set("Fn_BMIDE_SubLovOperations", objBMIDE, "ShowCascadingView", "ON") 
	
	 Select Case sAction

	   	 Case "VerifyNode"
	                intNodeCount = objBMIDE.JavaTree("LOVTable").GetROProperty("items count")		
					For intCount = 0 to intNodeCount - 1
							sTreeItem = objBMIDE.JavaTree("LOVTable").GetItem(intCount)
							If Trim(lcase(sTreeItem)) = Trim(Lcase(sNavPath)) Then
									Fn_BMIDE_SubLovOperations=True
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Pass: Successfully Verified Node "+ sNavPath)
									Exit For
							End If
					Next
					If Cint(intCount) = Cint(intNodeCount) Then
							Fn_BMIDE_SubLovOperations = FALSE
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail: Failed to Verify Node "+ sNavPath)
					End If
					
		Case "Select"
		            Call Fn_JavaTree_Select("Fn_BMIDE_SubLovOperations", objBMIDE, "LOVTable",sNavPath)
					Fn_BMIDE_SubLovOperations=True

		Case "Expand"
		            Call Fn_UI_JavaTree_Expand("Fn_BMIDE_SubLovOperations", objBMIDE, "LOVTable",sNavPath)
					Fn_BMIDE_SubLovOperations=True

		Case "ExpandAll"
					Call Fn_Button_Click("Fn_BMIDE_SubLovOperations",objBMIDE,"ExpandAllLOVValue")
                    Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Pass: Successfully Expand All  Node")
					Fn_BMIDE_SubLovOperations=True

		Case "CollapseAll"
					Call Fn_Button_Click("Fn_BMIDE_SubLovOperations",objBMIDE,"CollapseAllLOVValue")
                    Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Pass: Successfully Collapse Node ")
					Fn_BMIDE_SubLovOperations=True

		Case "AddSubLov"
					If sLovAttachName<> "" Then
                            ' Select item from LOVTable tree
							Call Fn_JavaTree_Select("Fn_BMIDE_SubLovOperations", objBMIDE, "LOVTable",sNavPath)
                           ' Click Add Sub Lov Button
							Call Fn_Button_Click("Fn_BMIDE_SubLovOperations",objBMIDE,"AddSubLOV")
                            ' Set Sub Lov Name
							'Call Fn_Edit_Box("Fn_BMIDE_SubLovOperations", objLovAttach,"LOV",sLovAttachName)
							Call Fn_Button_Click("Fn_BMIDE_SubLovOperations", objLovAttach,"BrowseLOV")
							Call Fn_Edit_Box("Fn_BMIDE_SubLovOperations",JavaWindow("Business Modeler").JavaWindow("FindLOV"),"LOVName",sLovAttachName)
							Call Fn_Button_Click("Fn_BMIDE_SubLovOperations",JavaWindow("Business Modeler").JavaWindow("FindLOV"),"OK")
							wait 1
													
							'Set Condition
							If  sLovAttachCond<>"" Then
									Call Fn_Edit_Box("Fn_BMIDE_SubLovOperations", objLovAttach,"Condition",sLovAttachCond)
							End If
							' Click Finish Button
							Call Fn_Button_Click("Fn_BMIDE_SubLovOperations",objLovAttach,"Finish")
							Fn_BMIDE_SubLovOperations=True
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Pass: Successfully Add Sub Lov")
					Else	
							Fn_BMIDE_SubLovOperations=False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail: Failed to Add Sub Lov")
					End If
					
		Case "RemSubLov"
					 ' Select item from LOVTable tree
					Call Fn_JavaTree_Select("Fn_BMIDE_SubLovOperations", objBMIDE, "LOVTable",sNavPath)

					Call Fn_Button_Click("Fn_BMIDE_SubLovOperations",objBMIDE,"RemSubLOV")
					Fn_BMIDE_SubLovOperations=True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Pass: Successfully remove Sub Lov")
		Case Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail : Please Select Correct Case")
					 Fn_BMIDE_SubLovOperations=false
	 End Select
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Pass: Function Completed " +sAction + " Successfully ")

Set objBMIDE = Nothing
End Function

'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_LOVTableOperations

'Description			 :	Function Used to perform operations on LOV

'Parameters			   :   '1.sAction = Action to Select
'										2. sNavPath = LOVTable tree Path
'										3. strValue = value of lov 
'										4.,strDescription = Description of Lov
'								'		5.,strDisplayName = Display name Lov
'										6 strCondition= Condition of Lov 

'Return Value		   : 	True Or False

'Pre-requisite			:	LOV Tab Should be appear on srceen

'Examples				:	 		Fn_BMIDE_LOVTableOperations("Add","","pl","pl","pl","isTrue")
'												Fn_BMIDE_LOVTableOperations("Edit","pl","pl","p123l","pl","isTrue")
'												Fn_BMIDE_LOVTableOperations("Remove","pl","","","","")
'												Fn_BMIDE_LOVTableOperations("Verify","","pl~sd","p123l~sd","","")
'												Fn_BMIDE_LOVTableOperations("CheckUncheck","Gold","","","","")
'												Fn_BMIDE_LOVTableOperations("MoveUp","Item Path","","","","Number of times Move Up the Item")
'												Fn_BMIDE_LOVTableOperations("MoveUp","D445566","","","","3")
'												Fn_BMIDE_LOVTableOperations("MoveDown","Item Path","","","","Number of times Move Down the Item")
'												Fn_BMIDE_LOVTableOperations("MoveDowm","D445566","","","","3")
'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done											Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Pranav Ingle						   						22-Nov-2010							1.0																												Sunny
'													Sandeep N						   						  29-Nov-2010							1.0							Case "CheckUncheck"									  Sunny
'													Sandeep N						   						  12-Jan-2011							1.0							Case "MoveUp","MoveDowm"					 Sunny
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------	

Public Function Fn_BMIDE_LOVTableOperations(strAction,sNavPath,strValue,strDescription,strDisplayName,strCondition)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_LOVTableOperations"
	Dim iItemCount,arrValue,arrDescription,iCount,iCounter,strDescValue,bFlag,strValueName,objBMIDE, objAddLov
	Set objBMIDE=JavaWindow("Business Modeler")
	
   Fn_BMIDE_LOVTableOperations=False
   Select Case strAction
		 	Case "Verify"
					arrValue=Split(strValue,"~")
					arrDescription=Split(strDescription,"~")
					iItemCount=Fn_UI_Object_GetROProperty("Fn_BMIDE_LOVTableOperations",JavaWindow("Business Modeler").JavaTree("LOVTable"),"items count")
					For iCount=0 To ubound(arrValue)
							bFlag=False
							For iCounter=0 To iItemCount-1
									strValueName=JavaWindow("Business Modeler").JavaTree("LOVTable").GetItem(iCounter)
									If strValueName=arrValue(iCount) Then
											strDescValue=JavaWindow("Business Modeler").JavaTree("LOVTable").GetColumnValue(arrValue(iCount),"Description")
											If  arrDescription(iCount)=strDescValue Then
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
							Fn_BMIDE_LOVTableOperations=True
					End If
			Case "Add"
					Set objAddLov = objBMIDE.JavaWindow("AddLOVValue")
					Call Fn_Button_Click("Fn_BMIDE_LOVTableOperations",objBMIDE,"AddLOVValue")
					Call Fn_Edit_Box("Fn_BMIDE_LOVTableOperations", objAddLov,"Value",strValue)
					If strDisplayName <> "" Then
							Call Fn_Edit_Box("Fn_BMIDE_LOVTableOperations", objAddLov,"ValueDisplayName",strDisplayName)
					End If
					If  strDescription<> "" Then
							Call Fn_Edit_Box("Fn_BMIDE_LOVTableOperations", objAddLov,"Description",strDescription)
					End If
					If strCondition <> "" Then
							Call Fn_Edit_Box("Fn_BMIDE_LOVTableOperations", objAddLov,"Condition",strCondition)
					End If
					Call Fn_Button_Click("Fn_BMIDE_LOVTableOperations",objAddLov,"Finish")
					Fn_BMIDE_LOVTableOperations=True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Pass: Successfully Add Lov")
					Set objAddLov  = Nothing

			Case "Edit","Edit_DoubleClick"

                    Set objAddLov = objBMIDE.JavaWindow("AddLOVValue")
					Call Fn_JavaTree_Select("Fn_BMIDE_SubLovOperations", objBMIDE, "LOVTable",sNavPath)
					If strAction="Edit_DoubleClick" Then
						Call Fn_JavaTree_Node_Activate("Fn_BMIDE_SubLovOperations",objBMIDE,"LOVTable",sNavPath)
						wait(1)
					Else
						Call Fn_Button_Click("Fn_BMIDE_LOVTableOperations",objBMIDE,"EditLOVValue")
						wait(2)
					End If

                    Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_BMIDE_LOVTableOperations",objAddLov,"title","Modify LOV Value")
					If strValue <> "" Then
						Call Fn_Edit_Box("Fn_BMIDE_LOVTableOperations", objAddLov,"Value",strValue)
					End If
					If strDisplayName <> "" Then
							Call Fn_Edit_Box("Fn_BMIDE_LOVTableOperations", objAddLov,"ValueDisplayName",strDisplayName)
					End If
					If  strDescription<> "" Then
							Call Fn_Edit_Box("Fn_BMIDE_LOVTableOperations", objAddLov,"Description",strDescription)
					End If
					If strCondition <> "" Then
							Call Fn_Edit_Box("Fn_BMIDE_LOVTableOperations", objAddLov,"Condition",strCondition)
					End If
					Call Fn_Button_Click("Fn_BMIDE_LOVTableOperations",objAddLov,"Finish")
					Fn_BMIDE_LOVTableOperations=True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Pass: Successfully Edit Lov")
					Set objAddLov  = Nothing

			Case "Remove"		
					Call Fn_JavaTree_Select("Fn_BMIDE_SubLovOperations", objBMIDE, "LOVTable",sNavPath)
					Call Fn_Button_Click("Fn_BMIDE_LOVTableOperations",objBMIDE,"RemoveLOVValue")
					Fn_BMIDE_LOVTableOperations=True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Pass: Successfully Remove Lov")

		Case "CheckUncheck"		
					Call Fn_JavaTree_Select("Fn_BMIDE_SubLovOperations", objBMIDE, "LOVTable",sNavPath)
					objBMIDE.JavaTree("LOVTable").PressKey " "
					Fn_BMIDE_LOVTableOperations=True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Pass: Check-Uncheck the value")
		Case "MoveUp"
					If strCondition="" Then
						strCondition=1
					End If
					For iCounter=1 To strCondition
						Call Fn_JavaTree_Select("Fn_BMIDE_SubLovOperations", objBMIDE, "LOVTable",sNavPath)
						Call Fn_Button_Click("Fn_BMIDE_LOVTableOperations",objBMIDE,"MoveUpLOV")
					Next
					Fn_BMIDE_LOVTableOperations=True
		Case "MoveDown"
					If strCondition="" Then
						strCondition=1
					End If
					For iCounter=1 To strCondition
						Call Fn_JavaTree_Select("Fn_BMIDE_SubLovOperations", objBMIDE, "LOVTable",sNavPath)
						Call Fn_Button_Click("Fn_BMIDE_LOVTableOperations",objBMIDE,"MoveDownLOV")
					Next
					Fn_BMIDE_LOVTableOperations=True
   End Select
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Pass: Function Completed " +strAction + " Successfully ")
	Set objBMIDE=Nothing
End Function
'-------------------------------------------------------------------Function Used to Launch the BMIDE Application-------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_LaunchBMIDE

'Description			 :	Function Used to Launch the BMIDE Application

'Return Value		   : 	True Or False

'Parameters				:	1. strWorkspcPath : WorkSpace Path

'Pre-requisite			:	BMIDE_Config.xml should be properly filled

'Examples				: 	'strExePath=Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\BMIDE_Config.xml", "BMIDEExecutable")
										'Call Fn_LaunchBMIDE(strExePath,"")
										'Call Fn_LaunchBMIDE(strExePath,"D:\Siemens\Teamcenter8\bmide\workspace\8000.3.0")

'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done															Reviewer
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				17/11/2010			           1.0																																									Sunny R
''													Sandeep N										   				13/01/2011			           1.0				Added Parameter "strBMIDEExecutablePath"				Sunny R
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\BMIDE_Config.xml", "BMIDEExecutable")
Public Function Fn_LaunchBMIDE(strBMIDEExecutablePath,strWorkspcPath)
	GBL_FAILED_FUNCTION_NAME="Fn_LaunchBMIDE"
	On Error resume next
	Dim sFolderPath,arrBatFileName,sDriveLetter,sNavFolder,objShell
	sFolderPath = strBMIDEExecutablePath
	If JavaWindow("Business Modeler").Exist(15)=True Then						  							
			Fn_LaunchBMIDE = True																				
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "BMIDE Application Is already Exist ")
	Else

			arrBatFileName = Split(sFolderPath,"\")
			sBatFileName = arrBatFileName(UBound(arrBatFileName))
			sFolderPath = Left(sFolderPath, (Len(sFolderPath)-(Len(sBatFileName)+1)))
			sDriveLetter = Split(sFolderPath, ":", -1, 1)
			sNavFolder = "cd " & sFolderPath 
			
			Set objShell = CreateObject("WScript.Shell")
			objShell.Run "%comspec% /c " & sDriveLetter(0) & ":" & "&" & sNavFolder & "&" & sBatFileName, 2, True
			Set objShell = Nothing

			bReturn=Fn_BMIDEWorkspace(strWorkspcPath)
			If bReturn=True Then
				If JavaWindow("Business Modeler").Exist(iTime)=True Then
						JavaWindow("Business Modeler").Maximize
						Fn_LaunchBMIDE = True  																						
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Invoked BMIDE Application from [" + strEXEPath + "]")
				Else
					 Fn_LaunchBMIDE = False
					 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Invoke BMIDE Application from [" + strEXEPath + "]")
					 Exit Function
				End If
			End If
	End If
End Function

'-------------------------------------------------------------------Function Used to Fill workspace Information-------------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDEWorkspace

'Description			 :	Function Used to Fill workspace Information

'Return Value		   : 	True Or False

'Pre-requisite			:	BMIDE_Config.xml should be properly filled

'Examples				: 	'Call Fn_BMIDEWorkspace("D:\Siemens\Teamcenter8\bmide\workspace\8000.3.0")
										'Call Fn_BMIDEWorkspace("")

'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				17/11/2010			           1.0																						Sunny R
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_BMIDEWorkspace(strWorkspace)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDEWorkspace"
    On Error resume next
    Dim strWorkspacePath
	If strWorkspace="" Then	
		strWorkspacePath = Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\BMIDEConfig\BMIDE_Config.xml", "BMIDEWorkspacePath")
	Else
		strWorkspacePath=strWorkspace
	End If
	If  strWorkspacePath="" Then
		Fn_BMIDEWorkspace=False
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Workspacepath is blank")
		Exit Function
	End If
	If JavaWindow("Eclipse Launcher").Exist(iTime) Then
		JavaWindow("Eclipse Launcher").JavaList("Workspace").Select strWorkspacePath
		JavaWindow("Eclipse Launcher").JavaButton("OK").Click
	ElseIf Dialog("WorkspaceLauncher").Exist(iTime)=True Then
		Dialog("WorkspaceLauncher").WinEdit("Workspace").Set strWorkspacePath
		Dialog("WorkspaceLauncher").WinButton("OK").Click
		Fn_BMIDEWorkspace=True
	ElseIf JavaWindow("TEM").Exist(iTime)=True Then
		JavaWindow("TEM").JavaList("Workspace").Select strWorkspacePath
		JavaWindow("TEM").JavaButton("OK").Click
		Fn_BMIDEWorkspace=True
	Else
		Fn_BMIDEWorkspace=False
	End If
End Function
'-------------------------------------------------------------------Function Used to Load BMIDE_Config xml--------------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_LoadEnvXML

'Description			 :	Function Used to Load BMIDE_Config xm

'Return Value		   : 	True Or False

'Pre-requisite			:	Environment path should be set proper

'Examples				: 	'Call Fn_BMIDE_LoadEnvXML()

'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done										Reviewer
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				17/11/2010			           1.0																										Sunny R
'													Sandeep N										   				18/11/2010			           1.0								Add Load File From Env Call				  Sunny R
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function Fn_BMIDE_LoadEnvXML()
	Dim sAutoDir
	sAutoDir = Fn_GetEnvValue("User", "AutomationDir")
	Environment.LoadFromFile(sAutoDir + "\TestData\BMIDEConfig\BMIDE_Config.xml")
	Fn_BMIDE_LoadEnvXML = True
End Function
'-------------------------------------------------------------------Function Used to Create Random Number-------------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_RandNoGenerate

'Description			 :	Function Used to Create Random Number

'Return Value		   : 	Random Number

'Parameters     		:	1. iLength : Length Of Random Number

'Examples				: 	'Call Fn_BMIDE_RandNoGenerate(2)
										'Call Fn_BMIDE_RandNoGenerate(3)
										'Call Fn_BMIDE_RandNoGenerate(4)
										'Call Fn_BMIDE_RandNoGenerate(5)
										'Call Fn_BMIDE_RandNoGenerate(6)
										'Call Fn_BMIDE_RandNoGenerate(7)

'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				17/11/2010			           1.0																						Sunny R
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_BMIDE_RandNoGenerate(iLength)
	Dim iNumber,iStartNumber
    Randomize
	iStartNumber="9"
	For iCount=1 To iLength-1
		iStartNumber=Cstr(iStartNumber)+"0"
	Next
	 iNumber = Int((iStartNumber * Rnd) + 1)
	 If Len(Cstr(iNumber)) < iLength Then
			Fn_BMIDE_RandNoGenerate = "0" + Cstr(iNumber)
	 Else
			Fn_BMIDE_RandNoGenerate = Cstr(iNumber)
	 End If
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Generated Random Number is [" + CStr(iNumber) + "]")		
End Function
'-------------------------------------------------------------------Function Used to Perform Operations On BMIDE Eclips Menus-------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_MenuOperation

''Description			:1. Function Used to Perform Operations On BMIDE Eclips Menus

'Return Value		   : 	True or False

'Parameters     		:	1. StrAction : Action Name
										'2. StrMenuLabel : Menu Name

'Pre-requisite			:	BMIDE Prespective is Open.

'Examples				: 'strMenu=Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\BMIDE_Menu.xml", "Close")
									  'Call Fn_BMIDE_MenuOperation("Select", strMenu)
									  'strMenu=Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\BMIDE_Menu.xml", "CloseAll")
									  'Call Fn_BMIDE_MenuOperation("Select", strMenu)
									  'Call Fn_BMIDE_MenuOperation("Select", "File:New:Project...")
									  'Calll Fn_BMIDE_MenuOperation("Select", "File:Import...")
									  'Imp Note :- All Menu Should Be Take From BMIDE_Menu.xml
									  'Fn_BMIDE_MenuOperation("KeyPress", "Project:Properties")

'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				17/11/2010			           1.0																						Sunny R
'													Sandeep N										   				25/11/2010			           1.0								Added Function Call					Sunny R
'													Sandeep N										   				12/01/2011			           1.0								Added Case "KeyPress"l					Sunny R
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_BMIDE_MenuOperation(StrAction, StrMenuLabel)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_MenuOperation"
   Dim ArrMenu,NumObjects
	Select Case StrAction
		'.---------------------------------------This case is used to select the menu ----------------------------------------------
		Case "Select"
            Fn_BMIDE_MenuOperation = Fn_UI_JavaMenu_Select("Fn_BMIDE_MenuOperation",JavaWindow("Business Modeler"),StrMenuLabel)
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Select  [" + StrMenuLabel + "] Menu")
		Case "KeyPress"
			'Split Menu String
			ArrMenu=Split(StrMenuLabel,":") 
			NumObjects = ubound(ArrMenu)

			'This is a Special case to operate menu by KeyPress method - Few Tc menus are not getting selected by traditional way
			Select Case NumObjects
				Case "1"
						JavaWindow("Business Modeler").PressKey Left(ArrMenu(0), 1), micAlt
						JavaWindow("Business Modeler").PressKey Left(ArrMenu(1), 1)
				Case "2"
						JavaWindow("Business Modeler").PressKey Left(ArrMenu(0), 1), micAlt
						JavaWindow("Business Modeler").PressKey Left(ArrMenu(1), 1)
						JavaWindow("Business Modeler").PressKey Left(ArrMenu(2), 1)
				End Select
				Fn_BMIDE_MenuOperation =True
         Case "SelectExt"
            Fn_BMIDE_MenuOperation = Fn_UI_JavaMenu_Select("Fn_BMIDE_MenuOperation",JavaWindow("BMIDEDefaultWindow"),StrMenuLabel)
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Select  [" + StrMenuLabel + "] Menu")
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[" + StrAction + "] Is Wrong Action Name")
			Fn_BMIDE_MenuOperation=False
	End Select
End Function
'-------------------------------------------------------------------Function Used to Reset Perspective---------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_ResetPerspective

''Description			:	Function Used to Reset Perspective

'Return Value		   : 	True or False

'Pre-requisite			:	BMIDE Prespective is Open.

'Examples				: 	'Call Fn_BMIDE_ResetPerspective()

'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				17/11/2010			           1.0																						Sunny R
'													Sandeep N										   				17/11/2010			           1.0							Added Env Menu Path 			   Sunny R
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_BMIDE_ResetPerspective()
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_ResetPerspective"
   Dim strMenuName
   Fn_BMIDE_ResetPerspective=False
   'Calling Window:Reset Perspective... Menu
   strMenuName=Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\BMIDEConfig\BMIDE_Menu.xml", "ResetPerspective")
   Call Fn_BMIDE_MenuOperation("Select", strMenuName)
   'Checking Existance of ResetPerspective window
	If Fn_UI_ObjectExist("Fn_BMIDE_ResetPerspective", JavaWindow("Business Modeler").JavaWindow("ResetPerspective"))=True Then
		'Clicking on OK button to Reset Perspective
		Call Fn_Button_Click("Fn_BMIDE_ResetPerspective", JavaWindow("Business Modeler").JavaWindow("ResetPerspective"), "Yes")
		'Function Returns True After reseting Perspective
		Fn_BMIDE_ResetPerspective=True
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Reset Perspective")
	End If
End Function
'-------------------------------------------------------------------Function Used to Create New Project----------------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_CreateNewProject

'Description			 :	Function Used to Create New Project

'Parameters			   :	1.strProjectName: Name of Project To Be Created
										'2.bDefaultOpt:Template Location Option
										'3.strDescription: Project Description
										'4.strPrefix:Project Prefix
										'5.strTempDirectory:Template Directory Path
										'6.strDepdTemplate:Dependant template Names
										'7.strLanguage:Language to Select

'Return Value		   : 	True Or False

'Pre-requisite			:	Should be Log In BMIDE

'Examples				: 	'1. Call Fn_BMIDE_CreateNewProject("Temp","True","","D3","D:\Siemens\Teamcenter8\bmide\templates","","")
										'Imp Note : For this Function All parameters come from environment(First Take Values From Env File And Pass to Function)
										'Pass Full Language Names by Collan Separeted (:)
										'strLanguage="cs_CZ - Czech (Czech Republic):zh_CN - Chinese (China)"
										'Call Fn_BMIDE_CreateNewProject("Temp","True","","D3","D:\Siemens\Teamcenter8\bmide\templates","",strLanguage)
										'strDepdTemplate :-Dependant Templates ( : ) sepearated
										'Call Fn_BMIDE_CreateNewProject("Temp","True","TestProject","D3","C:\Siemens\Teamcenter8\bmide\templates","Foundation:Change Management:Aerospace and Defense Change Management","")
'History					 :		
'													Developer Name												Date						Rev. No.						Changes Done														Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				18/11/2010			           1.0																																					Sunny R
'													Sandeep N										   				22/11/2010			           1.0						Added Code For Laguage Selection					  Sunny R
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public function Fn_BMIDE_CreateNewProject(strProjectName,bDefaultOpt,strDescription,strPrefix,strTempDirectory,strDepdTemplate,strLanguage)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_CreateNewProject"
   'Variable declaration
   Dim strMenu,iCounter,arrLanguage,iItemCount,iCount,LangName,bFlag,arrDepTemp,iRowCount,strTempName
   Dim ObjNewProjectWindow
   'Setting Function equals to False
   Fn_BMIDE_CreateNewProject=False
   'Checking Existance of NewProject window
	If Fn_UI_ObjectExist("Fn_BMIDE_CreateNewProject",JavaWindow("Business Modeler").JavaWindow("NewProject"))=False Then
		'Taking Menu Name from Environmet File
		strMenu=Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\BMIDEConfig\BMIDE_Menu.xml", "NewProject")
		'Calling File:New:Project... Menu to open Project Dialog
        Call Fn_BMIDE_MenuOperation("Select", strMenu)
	End If
	'Creating Object Of NewProject window
	Set ObjNewProjectWindow=Fn_UI_ObjectCreate("Fn_BMIDE_CreateNewProject", JavaWindow("Business Modeler").JavaWindow("NewProject"))
	'Expanding Business Modeler IDE Node
    Call Fn_UI_JavaTree_Expand("Fn_BMIDE_CreateNewProject", ObjNewProjectWindow, "WizardsTree","Business Modeler IDE")
	'Selecting Project
	Call Fn_JavaTree_Select("Fn_BMIDE_CreateNewProject", ObjNewProjectWindow,"WizardsTree","Business Modeler IDE:New Business Modeler IDE Template Project")
	'Clicking On Next button to Go Next Wizard
	Call Fn_Button_Click("Fn_BMIDE_CreateNewProject", ObjNewProjectWindow, "Next")
	If strPrefix<>"" Then
		'Setting Prefix To project
		Call Fn_Edit_Box("Fn_BMIDE_CreateNewProject",ObjNewProjectWindow,"Prefix",strPrefix)
	End If
	If strProjectName<>"" Then
		'Seeting project Name
		Call Fn_Edit_Box("Fn_BMIDE_CreateNewProject",ObjNewProjectWindow,"ProjectName",strProjectName)
	End If
	If bDefaultOpt=Cstr(True) Then
'		If strLocation<>"" Then
'			'Setting Project Location
'			Call Fn_Edit_Box("Fn_BMIDE_CreateNewProject",ObjNewProjectWindow,"UseDefaultLocation",strLocation)
'		End If
	End If
	If strDescription<>"" Then
		Call Fn_Edit_Box("Fn_BMIDE_CreateNewProject",ObjNewProjectWindow,"TemplateDesc",strProjectName)
	End If
	If strTempDirectory<>"" Then
		'Setting Template Directory to project
		Call Fn_Edit_Box("Fn_BMIDE_CreateNewProject",ObjNewProjectWindow,"TemplatesDirectoryPath",strTempDirectory)
	End If
	'SelectingDependanat Template Derectory
	'Clicking On Next button to Go Next Wizard
	Call Fn_Button_Click("Fn_BMIDE_CreateNewProject", ObjNewProjectWindow, "Next")
	If strDepdTemplate<>"" Then
		arrDepTemp=Split(strDepdTemplate,":")
		iRowCount=Fn_UI_Object_GetROProperty("Fn_BMIDE_CreateNewProject",ObjNewProjectWindow.JavaTable("DependentTemplates"),"rows")
		For iCount=0 To Ubound(arrDepTemp)
			bFlag=False
			For iCounter=0 To iRowCount-1
				strTempName=ObjNewProjectWindow.JavaTable("DependentTemplates").GetCellData(iCounter,"Template display name")
				If Trim(strTempName)=arrDepTemp(iCount) Then
					If arrDepTemp(iCount)<>"Foundation" Then
						ObjNewProjectWindow.JavaTable("DependentTemplates").SelectRow iCounter
						ObjNewProjectWindow.JavaTable("DependentTemplates").PressKey " "
					End If
					bFlag=True
					Exit For
				End If
			Next
			If bFlag=False Then
				Call Fn_Button_Click("Fn_BMIDE_CreateNewProject", ObjNewProjectWindow, "Cancel")
				Set ObjNewProjectWindow=Nothing
				Exit Function
			End If
		Next
	End If
	'Clicking On Next button to Go Next Wizard
	Call Fn_Button_Click("Fn_BMIDE_CreateNewProject", ObjNewProjectWindow, "Next")
	If strLanguage<>"" Then
		arrLanguage=Split(strLanguage,":")
		For iCounter=0 To Ubound(arrLanguage)
			iItemCount=Fn_UI_Object_GetROProperty("Fn_BMIDE_CreateNewProject",ObjNewProjectWindow.JavaTable("LocaleTable"), "rows")
			For iCount=0 To iItemCount-1
				bFlag=False
				LangName=ObjNewProjectWindow.JavaTable("LocaleTable").GetCellData(iCount,"Locale")
				If Trim(LangName)=Trim(arrLanguage(iCounter)) Then
					ObjNewProjectWindow.JavaTable("LocaleTable").SelectCell iCount,"Locale"
					wait(1)
					ObjNewProjectWindow.JavaTable("LocaleTable").PressKey " "
					bFlag=True
					Exit For
				End If
			Next
			If bFlag=False Then
				Exit For
			End If
		Next
		If bFlag=False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"False : Wrong Language Name pass by User")
			Set ObjNewProjectWindow=Nothing
			Exit Function
		End If
	End If
	'Clicking On Finish  button to Go to Create Project
	Call Fn_Button_Click("Fn_BMIDE_CreateNewProject", ObjNewProjectWindow, "Finish")
	For iCounter=0 to 20
		If Not JavaWindow("Business Modeler").JavaWindow("NewProject").JavaWindow("ProgressInformation").Exist(35) Then
			Fn_BMIDE_CreateNewProject=True
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Pass : Successfully Created a Project")
			Exit For
		End If
		wait(7)
	Next
	Set ObjNewProjectWindow=Nothing
End Function
'-------------------------------------------------------------------Function Used to Delete Existing Object----------------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_DeleteObject

'Description			 :	Function Used to Delete Existing Object

'Imp Note			   :	For This Function Delete Dialog open From PopUp Menu So user Have to Call PupUp Menu Call Explicitely

'Return Value		   : 	True Or False

'Pre-requisite			:	Delete Object Dialog Should be Appeared on Srceen

'Examples				: 	'1. Call Fn_BMIDE_DeleteObject()
										'Imp Note : For This Function Delete Dialog open From PopUp Menu So user Have to Call PupUp Menu Call Explicitely
'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				18/11/2010			           1.0																						Sunny R
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_BMIDE_DeleteObject()
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_DeleteObject"
   'Variable Declaration
   Dim bFlag
   'Initially Function Returns False
   Fn_BMIDE_DeleteObject=False
   'Chaking Existance Of DeleteObject Dioalog
   If Fn_UI_ObjectExist("Fn_BMIDE_DeleteObject",JavaWindow("Business Modeler").JavaWindow("Deleteobject"))=False Then
	   'If Delete Object Not Exist Then Function Will Exit From Next Statement
		Exit Function
   End If
   'Clicking On Finish Button to Delete Object
   bFlag=Fn_Button_Click("Fn_BMIDE_DeleteObject",JavaWindow("Business Modeler").JavaWindow("Deleteobject"),"Finish")
   If bFlag=True Then
	   'Function Returns True after Deleting the Object
	   Fn_BMIDE_DeleteObject=True
   End If
End Function

'-------------------------------------------------------------------Function Used to Perform Operation on LOV Localization----------------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_LOVLocalizationOperations

'Description			 :	Function Used to Perform Operation on LOV Localization

'Parameters			   :	1.strAction: Action Name
										'2.strValue:Value of LOV
										'3.strLocale: Locale Name
										'4.strValueLocalization:value Localization
										'5.strLocDesc:Localization Description
										'6.strStatus:Localization Status

'Return Value		   : 	True Or False

'Pre-requisite			:	Should be Log In BMIDE

'Examples				: 	'1. Call Fn_BMIDE_LOVLocalizationOperations("Add","CA","en_US","TestLoc","Test Localization","Approved")
										
'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				19/11/2010			           1.0																						Sunny R
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_BMIDE_LOVLocalizationOperations(strAction,strValue,strLocale,strValueLocalization,strLocDesc,strStatus)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_LOVLocalizationOperations"
   'Variable Declaration
   Dim ObjLOVLocWindow
   Dim iRowCount,strCellData,bFlag
   Fn_BMIDE_LOVLocalizationOperations=False
   bFlag=False
   If  strValue<>"" Then
	   'Selecting Value from LOV Table
	   Call Fn_JavaTree_Select("Fn_BMIDE_LOVLocalizationOperations", JavaWindow("Business Modeler"), "LOVTable",strValue)
   End If
   'Clicking on Localization button to add Localization
	Call Fn_Button_Click("Fn_BMIDE_LOVLocalizationOperations",JavaWindow("Business Modeler"),"LOVLocalization")
	wait(2)
	'Creating object of LOVLocalization Window
	Set ObjLOVLocWindow=Fn_UI_ObjectCreate("Fn_BMIDE_LOVLocalizationOperations",JavaWindow("Business Modeler").JavaWindow("LOVLocalization"))
	'Taking Row count of LOVValueLocalization Table
	iRowCount=Fn_UI_Object_GetROProperty("Fn_BMIDE_LOVLocalizationOperations",ObjLOVLocWindow.JavaTable("LOVValueLocalization"), "rows")
	For iCount=0 To iRowCount-1
		'Taking Cell data from LOVValueLocalization table
		strCellData=ObjLOVLocWindow.JavaTable("LOVValueLocalization").GetCellData(iCount,"LOV Value")
		If strValue=strCellData Then
			'Selecting Cell
			JavaWindow("Business Modeler").JavaWindow("LOVLocalization").JavaTable("LOVValueLocalization").SelectCell iCount,"Locale"
			bFlag=True
			Exit For
		End If
	Next
	If bFlag=False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL:[" + strValue +"] Is Not Exist in LOVValueLocalization Table")
		Set ObjLOVLocWindow=Nothing
		Exit Function
	End If
	Select Case strAction
	'case to Add Localization
		Case "Add"
			'Clicking on Add Button to Add Localization
			Call Fn_Button_Click("Fn_BMIDE_LOVLocalizationOperations",ObjLOVLocWindow,"Add")
			wait(2)
			If strLocale<>"" Then
                    Call Fn_List_Select("Fn_BMIDE_LOVLocalizationOperations", JavaWindow("Business Modeler").JavaWindow("Localization"),"Locale",strLocale)
			End If
			If strValueLocalization<>"" Then
				'Setting Value Localization
                Call Fn_Edit_Box("Fn_BMIDE_LOVLocalizationOperations",JavaWindow("Business Modeler").JavaWindow("Localization"),"ValueLocalization",strValueLocalization)
			End If
			If strLocDesc<>"" Then
				'Setting Value Localization Description
                Call Fn_Edit_Box("Fn_BMIDE_LOVLocalizationOperations",JavaWindow("Business Modeler").JavaWindow("Localization"),"LocalizationDesc",strLocDesc)
			End If
			If strStatus<>"" Then
    				'Selecting Status from Status List
                    Call Fn_List_Select("Fn_BMIDE_LOVLocalizationOperations", JavaWindow("Business Modeler").JavaWindow("Localization"),"Status",strStatus)
    		End If
			'Clicking On Finish button to finish localization add operation
			Call Fn_Button_Click("Fn_BMIDE_LOVLocalizationOperations",JavaWindow("Business Modeler").JavaWindow("Localization"),"Finish")
			Call Fn_Button_Click("Fn_BMIDE_LOVLocalizationOperations",ObjLOVLocWindow,"Finish")
			Fn_BMIDE_LOVLocalizationOperations=True
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Pass: Successfully Added Localization [" + strLocale +"]")
	End Select
	'Releasing LOVLocalization Window Object
	Set ObjLOVLocWindow=Nothing
End Function

'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_CreateNewCustomProperties

'Description			 :	Function Used to Create New Bussiness Object

'Parameters			   :   '1.strName= String property name
'										2.  strDisplayName= String Display name of property
'										3. strDescription = String Descrption of Property
'										4.,strAttributeType = Attribute type of Property 
'										5. strStringLength = length of of Property 
'										6. chkSetInitialValueToNull= 'ON'  if have to set initial value to null
'										7. strInitialValue = String Initail value of Property 
'										8.strLowerBound = Sting lowe bound of Property 
'										9.strUpperBound = String Upper bound of Property 
'										10.chkArray = 'ON'  if want to set array value 
'										11.chkKeys = 'ON' if have to set  key properties ON
'Return Value		   : 	True Or False


'Examples				:	 Fn_BMIDE_CreateNewCustomProperties("p1","pDisp1","pDesc1","String","32","","","","","","","")

'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Pranav Ingle						   					19-Nov-2010								1.0																						Sunny
'													Sandeep N												21-Nov-2010								1.0																						Sunny
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------	
Function Fn_BMIDE_CreateNewCustomProperties(strName,strDisplayName,strDescription,strAttributeType,strStringLength,strReferenceClass,chkSetInitialValueToNull,strInitialValue,strLowerBound,strUpperBound,chkArray,chkKeys)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_CreateNewCustomProperties"
	Dim objDialog,intCounter,strPrototype,sName,arrKeys,arrChkArray
	Set objDialog = JavaWindow("Business Modeler").JavaWindow("NewCustomProperty")

	Fn_BMIDE_CreateNewCustomProperties = false
	'Checking Existance of Custom property window
   If Fn_UI_ObjectExist("Fn_BMIDE_CreateNewCustomProperties", objDialog )=False Then
	   'If New Custom Properties window not exist then function will Exit
       Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail: New Custom Properties Window is not Exist ")
	   Exit Function
   End If
	If strName<>"" Then
		strPrototype= Fn_Edit_Box_GetValue("Fn_BMIDE_CreateNewCustomProperties",objDialog,"Name")
		strName=strPrototype+strName
		'Setting name to new custom property
        Call Fn_Edit_Box("Fn_BMIDE_CreateNewCustomProperties",objDialog ,"Name",strName)
	End If
	If strDescription<>"" Then
		'Setting Description to new custom property
        Call Fn_Edit_Box("Fn_BMIDE_CreateNewCustomProperties",objDialog ,"Description",strDescription)
	End If
	If strDisplayName<>"" Then
		'Setting Display Name to new custom property
        Call Fn_Edit_Box("Fn_BMIDE_CreateNewCustomProperties",objDialog ,"DisplayName",strDisplayName)
	End If
	If strAttributeType<> ""  Then
			'Setting Attribute type  to new custom property
			Call Fn_List_Select("Fn_BMIDE_CreateNewCustomProperties",objDialog,"AttributeType",strAttributeType)
	End If
	If strStringLength<>"" Then
		'Setting String Length to new custom property
        Call Fn_Edit_Box("Fn_BMIDE_CreateNewCustomProperties",objDialog ,"StringLength",strStringLength)
	End If

	If strReferenceClass<>"" Then
		'Setting Reference Class
		Call Fn_Edit_Box("Fn_BMIDE_CreateNewCustomProperties",objDialog ,"ReferenceClass",strReferenceClass)
	End If
	If chkSetInitialValueToNull<>"" Then
        Call Fn_CheckBox_Set("Fn_BMIDE_CreateNewCustomProperties", objDialog, "SetInitialValuetoNULL", chkSetInitialValueToNull)
	End If
	If Lcase(chkSetInitialValueToNull)<>"on" Then
			If strInitialValue<>"" Then
				Call Fn_Edit_Box("Fn_BMIDE_CreateNewCustomProperties",objDialog ,"InitialValue",strInitialValue)
			End If
		End If
	If strLowerBound<>"" Then
		Call Fn_Edit_Box("Fn_BMIDE_CreateNewCustomProperties",objDialog ,"LowerBound",strLowerBound)
	End If
	If strUpperBound<>"" Then
		Call Fn_Edit_Box("Fn_BMIDE_CreateNewCustomProperties",objDialog ,"UpperBound",strUpperBound)
	End If
	If chkArray<>"" Then
		arrChkArray=Split(chkArray,",")
		For intCounter=0 To Ubound(arrChkArray)
            Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_BMIDE_CreateNewCustomProperties",objDialog.JavaCheckBox("ArrayKeys"),"attached text",arrChkArray(intCounter))
			Call Fn_CheckBox_Set("Fn_BMIDE_CreateNewCustomProperties", objDialog, "ArrayKeys", "ON")
		Next
	End If
	If chkKeys<>"" Then
		arrKeys=Split(chkKeys,",")
		For intCounter=0 To Ubound(arrKeys)
            Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_BMIDE_CreateNewCustomProperties",objDialog.JavaCheckBox("Keys"),"attached text",arrKeys(intCounter))
			Call Fn_CheckBox_Set("Fn_BMIDE_CreateNewCustomProperties", objDialog, "Keys", "ON")
		Next
	End If

	' Click on Finnish Button
	Call Fn_Button_Click("Fn_BMIDE_CreateNewCustomProperties",objDialog,"Finish")
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Pass: Function Completed Successfully")
	Fn_BMIDE_CreateNewCustomProperties = True

Set objDialog = Nothing
End Function
'-------------------------------------------------------------------Function Used to Create New Bussiness Form----------------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_CreateNewForm

'Description			 :	Function Used to Create New Bussiness Form

'Parameters			   :	1.strFormName: Form Name
										'2.strDisplayName: Display Name Of Form
										'3.strParent:Parent type Of Object
										'4.strDesc: Form Description
										'5.bAdvance:Advance Option
										'6.strStorageClass:Storage Class Type
										'7.strClassName:Class name
										'8.strClassParent:Class Parent
										'9.strProperties: Properties to Add

'Return Value		   : 	True Or False

'Pre-requisite			:	Should be Log In BMIDE

'Examples				: Properties="Demo~Test:~:DemoPropety~TestProp:String~String:32~32:~:~ON:~:~:~:Array?,Unlimited~:Transient?,Nulls Allowed?~Nulls Allowed?"
'									  Call Fn_BMIDE_CreateNewForm("D2Test","","Form","Test Form","On","Use new class","D2Demo","POM_object",Properties)
									'Note 2 Properties Separated By (:) Colan
									'Number Of Properties Separated By (~)
									'Means If want two Add 3 New Properties Then there Names Are "Test1~Test2~Test3:DisplayTest1~DisplayTest2~DisplayTest3:........"
									'"Property Name~Property Name:Display Name~Display Name:Desciption~Desciption:Attribute Type:String Length:Reference Class:Set Initial Value to Null:Initial Value:Lower Bound:Upper Bound:Name Of Array Keys To Set ON(Separated by Comma(,)):Name Of Keys to Set ON(Separated by Comma(,))"
'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done								Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				16/11/2010			           1.0																								Sunny R
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_BMIDE_CreateNewForm(strFormName,strDisplayName,strParent,strDesc,bAdvance,strStorageClass,strClassName,strClassParent,strProperties)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_CreateNewForm"
	Dim ObjNewFormWindow
	Dim strPrefix,strName,arrRefClass
	Dim  iCounter,bReturn,arrMainString,arrName,arrDisplayName,arrDescription,arrAttributeType,arrStringLength,arrSetNull ,arrInitialValue,arrLowerBound,arrUpperBound,arrArray,arrKeys
	Fn_BMIDE_CreateNewForm=False
	If Fn_UI_ObjectExist("Fn_BMIDE_CreateNewForm", JavaWindow("Business Modeler").JavaWindow("CreateNewForm"))=False Then
	   'If NewBusinessObject window not exist then function will Exit
       Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail:New Form Create Window is not Exist ")
	   Exit Function
    End If
	'Creating Object of New Form window
	Set ObjNewFormWindow=Fn_UI_ObjectCreate("Fn_BMIDE_CreateNewForm",JavaWindow("Business Modeler").JavaWindow("CreateNewForm"))
	If  strFormName<>"" Then
		strPrefix= Fn_Edit_Box_GetValue("Fn_BMIDE_CreateNewForm",ObjNewFormWindow,"FormName")
		strFormName=strPrefix+strFormName
		'Setting Name to New Business Form
        Call Fn_Edit_Box("Fn_BMIDE_CreateNewForm",ObjNewFormWindow,"FormName",strFormName)
	End If
	If strDisplayName<>"" Then
		'Setting Display Name to New Business Form
        Call Fn_Edit_Box("Fn_BMIDE_CreateNewForm",ObjNewFormWindow,"DisplayName",strDisplayName)
	End If
	If strParent<>"" Then
		'Setting Display Name to New Business Form
        Call Fn_Edit_Box("Fn_BMIDE_CreateNewForm",ObjNewFormWindow,"Parent",strParent)
	End If
	If strDesc<>"" Then
		'Setting Display Name to New Business Form
        Call Fn_Edit_Box("Fn_BMIDE_CreateNewForm",ObjNewFormWindow,"Description",strDesc)
	End If
	If bAdvance<>"" Then
		'Setting Advanced option
		Call Fn_CheckBox_Set("Fn_BMIDE_CreateNewForm", ObjNewFormWindow, "Advanced", bAdvance)
		If strStorageClass<>"" Then
			'Selecting Storage Class
            Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_BMIDE_CreateNewForm",ObjNewFormWindow.JavaRadioButton("UseNewClass"),"attached text",strStorageClass)
			Call Fn_UI_JavaRadioButton_SetON("Fn_BMIDE_CreateNewForm",ObjNewFormWindow, "UseNewClass")
		End If
		If strClassName<>"" Then
			'Setting Class Name
			Call Fn_Edit_Box("Fn_BMIDE_CreateNewForm",ObjNewFormWindow,"ClassName",strClassName)
		End If
		If strClassParent<>"" Then
			'Setting Class Parent
			Call Fn_Edit_Box("Fn_BMIDE_CreateNewForm",ObjNewFormWindow,"ClassParent",strClassParent)
		End If
	End If
	'Setting Properties To Form
		If strProperties<>"" Then
			arrMainString = Split(strProperties, ":")
			arrName = Split(arrMainString(0),"~")
			arrDisplayName = Split(arrMainString(1),"~")
			arrDescription = Split(arrMainString(2),"~")
			arrAttributeType = Split(arrMainString(3),"~")
			arrStringLength = Split(arrMainString(4),"~")
			arrRefClass = Split(arrMainString(5),"~")
			arrSetNull = Split(arrMainString(6),"~")
			arrInitialValue = Split(arrMainString(7),"~")
			arrLowerBound = Split(arrMainString(8),"~")
			arrUpperBound = Split(arrMainString(9),"~")
			arrArray = Split(arrMainString(10),"~")
			arrKeys = Split(arrMainString(11),"~")

			For iCounter = 0 to UBound(arrName) 
				Call Fn_Button_Click("Fn_BMIDE_CreateNewForm", ObjNewFormWindow, "Add")
				bReturn= Fn_BMIDE_CreateNewCustomProperties(arrName(iCounter),arrDisplayName(iCounter),arrDescription(iCounter),arrAttributeType(iCounter),arrStringLength(iCounter),arrRefClass(iCounter),arrSetNull(iCounter),arrInitialValue(iCounter),arrLowerBound(iCounter),arrUpperBound(iCounter),arrArray(iCounter),arrKeys(iCounter))
				If bReturn= true Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Pass:Successfully Added New property of Name ["+arrName(iCounter)+"]")
				Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail :failed to Add New property of Name ["+arrName(iCounter)+"]")
						Fn_BMIDE_CreateNewForm= false
						Exit Function
				End If
			  Next
		End If

		Call Fn_Button_Click("Fn_BMIDE_CreateNewForm", ObjNewFormWindow, "Finish")
		Fn_BMIDE_CreateNewForm= True
		Set ObjNewFormWindow=Nothing
		
End Function
'------------------------------------------------------------'Function Used to Add Persistent Properties to the Object-----------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_PersistentPropertiesOperation

'Description			 :	Function Used to Add Persistent Properties to the Object

'Parameters			   :   '1.strAction: Action Name
'										1.strName= String property name
'										2.  strDisplayName= String Display name of property
'										3. strDescription = String Descrption of Property
'										4.,strAttributeType = Attribute type of Property 
'										5. strStringLength = length of of Property 
'										6. chkSetInitialValueToNull= 'ON'  if have to set initial value to null
'										7. strInitialValue = String Initail value of Property 
'										8.strLowerBound = Sting lowe bound of Property 
'										9.strUpperBound = String Upper bound of Property 
'										10.chkArray =Array Key CheckBox Names "Array?,Unlimited" OR "Array?,5" 'Comma Separated
'										11.chkKeys = Keys CheckBox Names "Transient?,Nulls Allowed?" 'Comma Separated
'										12.chkGetter=Getter CheckBox Names  "Overrridable?,Published?"
'										14.chkSetter=Setter CheckBox Names  "Overrridable?,Published?"

'Return Value		   : 	True Or False


'Examples				:	 Fn_BMIDE_PersistentPropertiesOperation("Add","TestDemo","","Test Demo","String","32","","","","","","","","","")
'										Fn_BMIDE_PersistentPropertiesOperation("SelectProperty","current_name","","","","","","","","","","","","","")
'										 Fn_BMIDE_PersistentPropertiesOperation("Edit","s3ToEdit","NewDispName","New Disc","","36","","","","","","","","","")
' 										Fn_BMIDE_PersistentPropertiesOperation("VerifyStorageType","S3bm_View1","","","UntypedReference","","","","","","","","","","")
'										Fn_BMIDE_PersistentPropertiesOperation("Remove","current_name","","","","","","","","","","","","","")
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N												21-Nov-2010								1.0																				 Sunny
'													Sandeep N												02-Dec-2010								1.0							Case "SelectProperty"				  Sunny
'													Sandeep N												31-Mar-2011								1.0							Case "VerifyStorageType"
'																																																			Case "Edit"								Sunny
'													Pranav Ingle											 20-Jan-2012							  1.1						   Case "Remove"						Sandeep
'													Pranav Ingle											 15-Feb-2012							  1.2						   Modified Case "Add"				   Sandeep
'																																																	For chkArray parameter		
'													Sandeep N												3-Jan-2013								1.3							modified Case "SelectProperty" : Replace ActivateRow method with SelectCell
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_BMIDE_PersistentPropertiesOperation(strAction,strName,strDisplayName,strDescription,strAttributeType,strStringLength,strReferenceClass,chkSetInitialValueToNull,strInitialValue,strLowerBound,strUpperBound,chkArray,chkKeys,chkGetter,chkSetter)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_PersistentPropertiesOperation"
	'Variable Declaration
	Dim ObjCustPropWindow,intCounter,strPrototype,sName,arrKeys,arrChkArray,arrGetter,arrSetter,bFlag,intRowCount,strPropName,arrArrayStatus,arrKeysStatus
	Dim crrStorageType
   'Function Returns False
	Fn_BMIDE_PersistentPropertiesOperation=False
	bFlag=False
	'Clicking On Properties Tab
	Call Fn_BMIDE_InnerTabOperations("Activate","Properties")
	Select Case strAction
		Case "Add", "AddFormProperty"
				'Clicking On Add Button To Add Persistent Properties 
				Call Fn_Button_Click("Fn_BMIDE_PersistentPropertiesOperation", JavaWindow("Business Modeler"), "AddPropeties")
				wait(2)
				'Checking Existance of NewCustomProperty Window
				If Fn_UI_ObjectExist("Fn_BMIDE_PersistentPropertiesOperation", JavaWindow("Business Modeler").JavaWindow("NewCustomProperty"))=False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail: NewCustomProperty Dialog Is Not Exist")
					Exit Function
				End If
				'Creating Object Of NewCustomProperty window
				Set ObjCustPropWindow=Fn_UI_ObjectCreate("Fn_BMIDE_PersistentPropertiesOperation", JavaWindow("Business Modeler").JavaWindow("NewCustomProperty"))
				'Selecting Persistent option
				Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_BMIDE_PersistentPropertiesOperation",ObjCustPropWindow.JavaRadioButton("PropertyType"),"attached text","Persistent")
				Call Fn_UI_JavaRadioButton_SetON("Fn_BMIDE_PersistentPropertiesOperation",ObjCustPropWindow, "PropertyType")
				'Clicking On Next Button
				Call Fn_Button_Click("Fn_BMIDE_PersistentPropertiesOperation", ObjCustPropWindow, "Next")
				If strName<>"" Then
					strPrototype= Fn_Edit_Box_GetValue("Fn_BMIDE_PersistentPropertiesOperation",ObjCustPropWindow,"Name")
					strName=strPrototype+strName
					'Setting name to new custom property
					Call Fn_Edit_Box("Fn_BMIDE_PersistentPropertiesOperation",ObjCustPropWindow ,"Name",strName)
					If strDisplayName="" Then
						strDisplayName=strName
					End If
				End If
				If strDescription<>"" Then
					'Setting Description to new custom property
					Call Fn_Edit_Box("Fn_BMIDE_PersistentPropertiesOperation",ObjCustPropWindow ,"Description",strDescription)
				End If
				If strDisplayName<>"" Then
					'Setting Display Name to new custom property
					Call Fn_Edit_Box("Fn_BMIDE_PersistentPropertiesOperation",ObjCustPropWindow ,"DisplayName",strDisplayName)
				End If
				If strAttributeType<> ""  Then
						'Setting Attribute type  to new custom property
						Call Fn_List_Select("Fn_BMIDE_PersistentPropertiesOperation",ObjCustPropWindow,"AttributeType",strAttributeType)
				End If
				If strStringLength<>"" Then
					'Setting String Length to new custom property
					Call Fn_Edit_Box("Fn_BMIDE_PersistentPropertiesOperation",ObjCustPropWindow ,"StringLength",strStringLength)
				End If
				If strReferenceClass<>"" Then
					'Setting Reference Class
					Call Fn_Edit_Box("Fn_BMIDE_PersistentPropertiesOperation",ObjCustPropWindow ,"ReferenceClass",strReferenceClass)
				End If
				If chkSetInitialValueToNull<>"" Then
					'Setting Initial Value To Null
					Call Fn_CheckBox_Set("Fn_BMIDE_PersistentPropertiesOperation", ObjCustPropWindow, "SetInitialValuetoNULL", chkSetInitialValueToNull)
					If Lcase(chkSetInitialValueToNull)<>"on" Then
						If strInitialValue<>"" Then
							Call Fn_Edit_Box("Fn_BMIDE_PersistentPropertiesOperation",ObjCustPropWindow ,"InitialValue",strInitialValue)
						End If
					End If
				End If
				If strLowerBound<>"" Then
					'Setting Lower Bound
					Call Fn_Edit_Box("Fn_BMIDE_PersistentPropertiesOperation",ObjCustPropWindow ,"LowerBound",strLowerBound)
				End If
				If strUpperBound<>"" Then
					'Setting Upper Bound
					Call Fn_Edit_Box("Fn_BMIDE_PersistentPropertiesOperation",ObjCustPropWindow ,"UpperBound",strUpperBound)
				End If
				If chkArray<>"" Then
					'Selecting Array Keys
					arrChkArray=Split(chkArray,",")
					Call Fn_CheckBox_Set("Fn_BMIDE_PersistentPropertiesOperation", ObjCustPropWindow, "ArrayKeys", "On")
					If Ubound(arrChkArray)=1 Then
						If arrChkArray(1)="Unlimited" Then
							Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_BMIDE_PersistentPropertiesOperation",ObjCustPropWindow.JavaCheckBox("Keys"),"attached text",arrChkArray(1))
							Call Fn_CheckBox_Set("Fn_BMIDE_PersistentPropertiesOperation", ObjCustPropWindow, "Keys", "On")
						ElseIf IsNumeric(arrChkArray(1)) =True Then
							Call Fn_Edit_Box("Fn_BMIDE_PersistentPropertiesOperation",ObjCustPropWindow ,"MaxLength",arrChkArray(1))
							If Fn_UI_ObjectExist("Fn_BMIDE_PersistentPropertiesOperation", ObjCustPropWindow.Dialog("AttributeKeysNotification"))=True Then
								ObjCustPropWindow.Dialog("AttributeKeysNotification").WinButton("OK").Click
							End If
						End If
					End If
				End If
				If chkKeys<>"" Then
					'Selecting Keys
					arrKeys=Split(chkKeys,",")
					For intCounter=0 To Ubound(arrKeys)
						arrKeysStatus=Split(arrKeys(intCounter),":")
						Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_BMIDE_PersistentPropertiesOperation",ObjCustPropWindow.JavaCheckBox("Keys"),"attached text",arrKeysStatus(0))
						If arrKeysStatus(1)<>"" Then
							Call Fn_CheckBox_Set("Fn_BMIDE_PersistentPropertiesOperation", ObjCustPropWindow, "Keys", arrKeysStatus(1))
							If Fn_UI_ObjectExist("Fn_BMIDE_PersistentPropertiesOperation", ObjCustPropWindow.Dialog("AttributeKeysNotification"))=True Then
								JavaWindow("Business Modeler").JavaWindow("NewCustomProperty").Dialog("AttributeKeysNotification").WinButton("OK").Click
							End If
						End If
						
					Next
				End If
				
				If chkGetter<>"" Then
					'Selecting Getter And Setter Check box		
					bFlag=Fn_Button_Click("Fn_BMIDE_PersistentPropertiesOperation", ObjCustPropWindow, "Next")		
					arrGetter=Split(chkGetter,",")
					For intCounter=0 To Ubound(arrGetterSetter)
						JavaWindow("Business Modeler").JavaWindow("NewCustomProperty").JavaCheckBox("Getter").SetTOProperty "index",0
						Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_BMIDE_PersistentPropertiesOperation",ObjCustPropWindow.JavaCheckBox("Getter"),"attached text",arrGetter(intCounter))
						Call Fn_CheckBox_Set("Fn_BMIDE_PersistentPropertiesOperation", ObjCustPropWindow, "Getter", "ON")
					Next
				
				End If
				If chkSetter<>"" Then
					If bFlag=False Then
						Call Fn_Button_Click("Fn_BMIDE_PersistentPropertiesOperation", ObjCustPropWindow, "Next")
					End If
						arrSetter=Split(chkSetter,",")
						For intCounter=0 To Ubound(arrGetterSetter)
							JavaWindow("Business Modeler").JavaWindow("NewCustomProperty").JavaCheckBox("Getter").SetTOProperty "index",1
							Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_BMIDE_PersistentPropertiesOperation",ObjCustPropWindow.JavaCheckBox("Getter"),"attached text",arrSetter(intCounter))
							Call Fn_CheckBox_Set("Fn_BMIDE_PersistentPropertiesOperation", ObjCustPropWindow, "Getter", "ON")
					Next
				End If
				If strAction <> "AddFormProperty" Then
					If Fn_UI_Object_SetTOProperty_ExistCheck("Fn_BMIDE_RuntimePropertiesOperation",ObjCustPropWindow.JavaCheckBox("DescriptorOption"),"attached text","Show this property during creation of a Business Object.") = True Then
						Call Fn_CheckBox_Set("Fn_BMIDE_RuntimePropertiesOperation", ObjCustPropWindow, "DescriptorOption", "OFF")
					End If
				End If
				' Click on Finnish Button
				Call Fn_Button_Click("Fn_BMIDE_PersistentPropertiesOperation",ObjCustPropWindow,"Finish")
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Pass: Function Completed Successfully")
				Fn_BMIDE_PersistentPropertiesOperation = True

	Case "SelectProperty","DoubleClickProperty"
				intRowCount=Fn_UI_Object_GetROProperty("Fn_BMIDE_PersistentPropertiesOperation",JavaWindow("Business Modeler").JavaTable("PropertiesTable"), "rows")
				For intCounter=0 To intRowCount-1
					strPropName=JavaWindow("Business Modeler").JavaTable("PropertiesTable").GetCellData(intCounter,"Property Name")
					If Trim(strPropName)=Trim(strName) Then
								If  strAction = "SelectProperty" Then
										JavaWindow("Business Modeler").JavaTable("PropertiesTable").SelectCell intCounter,0
								Else
										JavaWindow("Business Modeler").JavaTable("PropertiesTable").ActivateRow  intCounter
							   End If
						Fn_BMIDE_PersistentPropertiesOperation=True
						Exit For
					End If
				Next
	Case "VerifyProperty"
				intRowCount=Fn_UI_Object_GetROProperty("Fn_BMIDE_PersistentPropertiesOperation",JavaWindow("Business Modeler").JavaTable("PropertiesTable"), "rows")
				For intCounter=0 To intRowCount-1
					strPropName=JavaWindow("Business Modeler").JavaTable("PropertiesTable").GetCellData(intCounter,"Property Name")
					If Trim(strPropName)=Trim(strName) Then
                				Fn_BMIDE_PersistentPropertiesOperation=True
						Exit For
					End If
		      Next
		Case "Edit","Edit_DoubleClick"
				Call Fn_BMIDE_PersistentPropertiesOperation("SelectProperty",strName,"","","","","","","","","","","","","")
				'Clicking On Add Button To Add Persistent Properties 
				If strAction="Edit_DoubleClick" Then
					JavaWindow("Business Modeler").JavaTable("PropertiesTable").ClickCell 1,1
					Call Fn_BMIDE_PersistentPropertiesOperation("DoubleClickProperty",strName,"","","","","","","","","","","","","")
				Else
					Call Fn_Button_Click("Fn_BMIDE_PersistentPropertiesOperation", JavaWindow("Business Modeler"), "EditProperties")
				End If
				wait(2)
				JavaWindow("Business Modeler").JavaWindow("NewCustomProperty").SetTOProperty "title","Modify Property"
				'Checking Existance of NewCustomProperty Window
				If Fn_UI_ObjectExist("Fn_BMIDE_PersistentPropertiesOperation", JavaWindow("Business Modeler").JavaWindow("NewCustomProperty"))=False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail: NewCustomProperty Dialog Is Not Exist")
					Exit Function
				End If
				'Creating Object Of NewCustomProperty window
				Set ObjCustPropWindow=Fn_UI_ObjectCreate("Fn_BMIDE_PersistentPropertiesOperation", JavaWindow("Business Modeler").JavaWindow("NewCustomProperty"))
				If strDescription<>"" Then
					'Setting Description to new custom property
					Call Fn_Edit_Box("Fn_BMIDE_PersistentPropertiesOperation",ObjCustPropWindow ,"Description",strDescription)
				End If
				If strDisplayName<>"" Then
					'Setting Display Name to new custom property
					Call Fn_Edit_Box("Fn_BMIDE_PersistentPropertiesOperation",ObjCustPropWindow ,"DisplayName",strDisplayName)
				End If
				If strAttributeType<> ""  Then
						'Setting Attribute type  to new custom property
						Call Fn_List_Select("Fn_BMIDE_PersistentPropertiesOperation",ObjCustPropWindow,"AttributeType",strAttributeType)
				End If
				If strStringLength<>"" Then
					'Setting String Length to new custom property
                    Call Fn_Edit_Box("Fn_BMIDE_PersistentPropertiesOperation",ObjCustPropWindow ,"StringLength","")
					Call Fn_UI_EditBox_Type("Fn_BMIDE_PersistentPropertiesOperation",ObjCustPropWindow ,"StringLength",strStringLength)
				End If
				If strReferenceClass<>"" Then
					'Setting Reference Class
					Call Fn_Edit_Box("Fn_BMIDE_PersistentPropertiesOperation",ObjCustPropWindow ,"ReferenceClass",strReferenceClass)
				End If
				If chkSetInitialValueToNull<>"" Then
					'Setting Initial Value To Null
					Call Fn_CheckBox_Set("Fn_BMIDE_PersistentPropertiesOperation", ObjCustPropWindow, "SetInitialValuetoNULL", chkSetInitialValueToNull)
					If Lcase(chkSetInitialValueToNull)<>"on" Then
						If strInitialValue<>"" Then
							Call Fn_Edit_Box("Fn_BMIDE_PersistentPropertiesOperation",ObjCustPropWindow ,"InitialValue",strInitialValue)
						End If
					End If
				End If
				If strLowerBound<>"" Then
					'Setting Lower Bound
					Call Fn_Edit_Box("Fn_BMIDE_PersistentPropertiesOperation",ObjCustPropWindow ,"LowerBound",strLowerBound)
				End If
				If strUpperBound<>"" Then
					'Setting Upper Bound
					Call Fn_Edit_Box("Fn_BMIDE_PersistentPropertiesOperation",ObjCustPropWindow ,"UpperBound",strUpperBound)
				End If
				If chkArray<>"" Then
					'Selecting Array Keys
					arrChkArray=Split(chkArray,",")
					For intCounter=0 To Ubound(arrChkArray)
						arrArrayStatus=Split(arrChkArray(intCounter),":")

						Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_BMIDE_PersistentPropertiesOperation",ObjCustPropWindow.JavaCheckBox("ArrayKeys"),"attached text",arrArrayStatus(0))
						If arrArrayStatus(1)<>"" Then
							Call Fn_CheckBox_Set("Fn_BMIDE_PersistentPropertiesOperation", ObjCustPropWindow, "ArrayKeys", arrArrayStatus(1))
						End If
						
					Next
				End If
				If chkKeys<>"" Then
					'Selecting Keys
					arrKeys=Split(chkKeys,",")
					For intCounter=0 To Ubound(arrKeys)
						arrKeysStatus=Split(arrKeys(intCounter),":")
						Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_BMIDE_PersistentPropertiesOperation",ObjCustPropWindow.JavaCheckBox("Keys"),"attached text",arrKeysStatus(0))
						If arrKeysStatus(1)<>"" Then
							Call Fn_CheckBox_Set("Fn_BMIDE_PersistentPropertiesOperation", ObjCustPropWindow, "Keys", arrKeysStatus(1))
							If Fn_UI_ObjectExist("Fn_BMIDE_PersistentPropertiesOperation", ObjCustPropWindow.Dialog("AttributeKeysNotification"))=True Then
								JavaWindow("Business Modeler").JavaWindow("NewCustomProperty").Dialog("AttributeKeysNotification").WinButton("OK").Click
							End If
						End If
						
					Next

				End If
				
				If chkGetter<>"" Then
					'Selecting Getter And Setter Check box		
					bFlag=Fn_Button_Click("Fn_BMIDE_PersistentPropertiesOperation", ObjCustPropWindow, "Next")		
					arrGetter=Split(chkGetter,",")
					For intCounter=0 To Ubound(arrGetterSetter)
						JavaWindow("Business Modeler").JavaWindow("NewCustomProperty").JavaCheckBox("Getter").SetTOProperty "index",0
						Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_BMIDE_PersistentPropertiesOperation",ObjCustPropWindow.JavaCheckBox("Getter"),"attached text",arrGetter(intCounter))
						Call Fn_CheckBox_Set("Fn_BMIDE_PersistentPropertiesOperation", ObjCustPropWindow, "Getter", "ON")
					Next
				
				End If
				If chkSetter<>"" Then
					If bFlag=False Then
						Call Fn_Button_Click("Fn_BMIDE_PersistentPropertiesOperation", ObjCustPropWindow, "Next")
					End If
						arrSetter=Split(chkSetter,",")
						For intCounter=0 To Ubound(arrGetterSetter)
							JavaWindow("Business Modeler").JavaWindow("NewCustomProperty").JavaCheckBox("Getter").SetTOProperty "index",1
							Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_BMIDE_PersistentPropertiesOperation",ObjCustPropWindow.JavaCheckBox("Getter"),"attached text",arrSetter(intCounter))
							Call Fn_CheckBox_Set("Fn_BMIDE_PersistentPropertiesOperation", ObjCustPropWindow, "Getter", "ON")
					Next
				End If
				' Click on Finnish Button
				Call Fn_Button_Click("Fn_BMIDE_PersistentPropertiesOperation",ObjCustPropWindow,"Finish")
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Pass: Function Completed Successfully")
				Fn_BMIDE_PersistentPropertiesOperation = True

		Case "VerifyStorageType"
				intRowCount=Fn_UI_Object_GetROProperty("Fn_BMIDE_PersistentPropertiesOperation",JavaWindow("Business Modeler").JavaTable("PropertiesTable"), "rows")
				For intCounter=0 To intRowCount-1
					strPropName=JavaWindow("Business Modeler").JavaTable("PropertiesTable").GetCellData(intCounter,"Property Name")
					If Trim(strPropName)=Trim(strName) Then
						crrStorageType=JavaWindow("Business Modeler").JavaTable("PropertiesTable").GetCellData(intCounter,"Storage Type")
						If Trim(crrStorageType)=Trim(strAttributeType) Then
							Fn_BMIDE_PersistentPropertiesOperation=True
							Exit For
						End IF
					End If
				Next

		Case "Remove"
				bFlag=Fn_BMIDE_PersistentPropertiesOperation("SelectProperty",strName,"","","","","","","","","","","","","")
				If bFlag=True Then
					Call Fn_Button_Click("Fn_BMIDE_PersistentPropertiesOperation",JavaWindow("Business Modeler"),"RemoveChangeID")
					Call Fn_BMIDE_DeleteObject()
					Fn_BMIDE_PersistentPropertiesOperation=True
				Else
					Exit Function
				End If		
	End Select

Set ObjCustPropWindow = Nothing
End Function

'------------------------------------------------------------'Function Used to Create Alternate ID Rules-----------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_CreateAlternateIDRule

'Description			 :	Function Used to Create Alternate ID Rules

'Parameters			   :   '1.strIdentifierContextButtonName: One Button Name Which Appear Against Identifier Context Edit Box.
'																											Because If Want to Use Existing Context Then Pass "Browse"
'																											If Want to Create New Context Then Pass "New"
'																											This Parameter is must. 
'										2.strContextName: Context name.
'																			If want to use Existing Then Pass name of Existing Context Name.
'																			if Want to Create New Then Give New Name (whatever user wants)
'										3.  strContextDesplayName:Display Name of Newly Created Context identifier
'										4. strContextDesc: Descrption of Context Identifier which creating new
'										5.,strIdentifierType :Identifier Type
'										6. strMasterRule :Master Rule
'										7. strMasterRuleDesc: Master Rule Descriotionl
'										8. strSupplimentRule: Supplimentory Rule
'										9.strSupplimentRuleDesc: Supplimentory Rule Description

'Return Value		   : 	True Or False


'Examples				:	 Fn_BMIDE_CreateAlternateIDRule("New","Demo","","Demo Identifier Context","Identifier","My Rule","My Demo Rule","My Supliment Rule","My Demo Supliment Rule")
'										Fn_BMIDE_CreateAlternateIDRule("Browse","D3wq","","","Identifier","My Rule","My Demo Rule","My Supliment Rule","My Demo Supliment Rule")

'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N												22-Nov-2010								1.0																						Sunny
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------	
Public Function Fn_BMIDE_CreateAlternateIDRule(strIdentifierContextButtonName,strContextName,strContextDesplayName,strContextDesc,strIdentifierType,strMasterRule,strMasterRuleDesc,strSupplimentRule,strSupplimentRuleDesc)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_CreateAlternateIDRule"
	Dim ObjAlternateIDDialog
	Dim strPrefix
	Fn_BMIDE_CreateAlternateIDRule=False
	If Fn_UI_ObjectExist("Fn_BMIDE_CreateAlternateIDRule",JavaWindow("Business Modeler").JavaWindow("NewAlternateIdRule"))=True Then
        Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Pass: NewAlternateIdRule Dialog Is Already Open")	
	Else 
		'Selecting "Alternate Id Rules" Tab From Inner Tabs
		Call Fn_UI_JavaTab_Select("Fn_BMIDE_CreateAlternateIDRule",JavaWindow("Business Modeler"),"InnerTab", "Alternate ID Rules")
		'Clicking On "AddAlternateIDRule" Buttom To Invoke "NewAlternateIdRule" Dialog
		Call Fn_Button_Click("Fn_BMIDE_CreateAlternateIDRule", JavaWindow("Business Modeler"), "AddAlternateIDRule")
	End If
	'Creating Object Of "NewAlternateIdRule" Window
	Set ObjAlternateIDDialog=Fn_UI_ObjectCreate("Fn_BMIDE_CreateAlternateIDRule",JavaWindow("Business Modeler").JavaWindow("NewAlternateIdRule"))
	'If User Want to Use Existing Identifier Context Then Have to Pass strIdentifierContextButtonName Parameter to "Browse"
	If strIdentifierContextButtonName="Browse" Then
		'Setting strContextName
		Call Fn_Edit_Box("Fn_BMIDE_CreateAlternateIDRule",ObjAlternateIDDialog,"IdentifierContext",strContextName)
	'If User Want to Create New Identifier Context Then Have to Pass strIdentifierContextButtonName Parameter to "New"
	ElseIf strIdentifierContextButtonName="New" Then
		'Clicking on New Button To Create New Identifier Context
		Call Fn_Button_Click("Fn_BMIDE_CreateAlternateIDRule", ObjAlternateIDDialog, "New")
		'Taking Prefix from Name Edit Box
		strPrefix=Fn_Edit_Box_GetValue("Fn_BMIDE_CreateAlternateIDRule",JavaWindow("Business Modeler").JavaWindow("NewIdentifierContext"),"Name")
		strContextName=strPrefix+strContextName
		'Setting strContextName
		Call Fn_Edit_Box("Fn_BMIDE_CreateAlternateIDRule",JavaWindow("Business Modeler").JavaWindow("NewIdentifierContext"),"Name",strContextName)
		If strContextDesplayName<>"" Then
			'Setting Display Name Of Identifier Context
			Call Fn_Edit_Box("Fn_BMIDE_CreateAlternateIDRule",JavaWindow("Business Modeler").JavaWindow("NewIdentifierContext"),"DisplayName",strContextDesplayName)
		End If
		If strContextDesc<>"" Then
			'Setting Description Name Of Identifier Context
			Call Fn_Edit_Box("Fn_BMIDE_CreateAlternateIDRule",JavaWindow("Business Modeler").JavaWindow("NewIdentifierContext"),"Description",strContextDesc)
		End If
		'Clicking On Finish Button To Create New Identifier Context
		Call Fn_Button_Click("Fn_BMIDE_CreateAlternateIDRule", JavaWindow("Business Modeler").JavaWindow("NewIdentifierContext"), "Finish")
	Else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail: Wrong Button Name Pass to Set Identifier Context")	
	End If
	If strIdentifierType<>"" Then
		'Setting Identifier type
		Call Fn_Edit_Box("Fn_BMIDE_CreateAlternateIDRule",ObjAlternateIDDialog,"IdentifierType",strIdentifierType)
	End If
	'Creting Master Rule
	If strMasterRule<>"" Then
		'Setting Master Rule
		Call Fn_Edit_Box("Fn_BMIDE_CreateAlternateIDRule",ObjAlternateIDDialog,"Rule",strMasterRule)
	End If
	If strMasterRuleDesc<>"" Then
		'Setting Master Rule Description
		Call Fn_Edit_Box("Fn_BMIDE_CreateAlternateIDRule",ObjAlternateIDDialog,"Description",strMasterRuleDesc)
	End If
	'Clicking on Next Button To Set Supplimemtory Rule And Description
	Call Fn_Button_Click("Fn_BMIDE_CreateAlternateIDRule",ObjAlternateIDDialog, "Next")
	'Creting Master Rule
	If strSupplimentRule<>"" Then
		'Setting Master Rule
		Call Fn_Edit_Box("Fn_BMIDE_CreateAlternateIDRule",ObjAlternateIDDialog,"Rule",strSupplimentRule)
	End If
	If strSupplimentRuleDesc<>"" Then
		'Setting Master Rule Description
		Call Fn_Edit_Box("Fn_BMIDE_CreateAlternateIDRule",ObjAlternateIDDialog,"Description",strSupplimentRuleDesc)
	End If
	Fn_BMIDE_CreateAlternateIDRule=True
	'Clicking on Finish Button To Create New Alternate ID Rule
	Call Fn_Button_Click("Fn_BMIDE_CreateAlternateIDRule", ObjAlternateIDDialog, "Finish")
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Pass: Successfully Created Alternate Id Rule")
	'Releasing Object Of Alternate Id Dialog
	Set ObjAlternateIDDialog=Nothing
End Function
'-------------------------------------------------------------------Function Used to Verify Error Messages which comes while Adding LOV Values-------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_AddLOVValueErrorMsgVerify

'Description			 :	Function Used to Verify Error Messages which comes while Adding LOV Values

'Parameters			   :	1.strName: Name Of LOV
										'2.strDesc:LOV Description
										'3.strType: LOV type
										'4.bUsage:Usage Option
										'5.bUsage:Usage Option
										'6.strValue:Value To Add for LOV
										'7.strDisplayName:Value Display Name
										'8.strValueDesc:Value Description
										'9.strCondition:Condition
										'10.strErrorMsg: Possible Error Message

'Return Value		   : 	True Or False

'Pre-requisite			:	Should be Log In BMIDE

'Examples				: strmsg= "Invalid ""Value"" field. LOV value ""001"" contains leading zero(es)."
'									  strmsg= "Invalid ""Value"" field. ""001B"" must be a number."
'									  Call Fn_BMIDE_AddLOVValueErrorMsgVerify("Test","TestDescription","ListOfValuesInteger","","001","","Test Desc","isTrue",strmsg)
'									 Call Fn_BMIDE_AddLOVValueErrorMsgVerify("Test","TestDescription","ListOfValuesInteger","","001B","","Test Desc","isTrue",strmsg)

'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				24/11/2010			           1.0																					Sunny R
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_BMIDE_AddLOVValueErrorMsgVerify(strName,strDesc,strType,bUsage,strValue,strDisplayName,strValueDesc,strCondition,strErrorMsg)
GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_AddLOVValueErrorMsgVerify"
   Dim ObjNewLOVWindow
   Dim bFlag,strPrototype,strErrorMessage
   GBL_EXPECTED_MESSAGE=strErrorMsg
  'Function Return False
  bFlag=False
   Fn_BMIDE_AddLOVValueErrorMsgVerify=False
   'Checking Existance of NewLOVObject window
 If Not JavaWindow("Business Modeler").JavaWindow("AddLOVValue").Exist(8) Then
	 
	   If Fn_UI_ObjectExist("Fn_BMIDE_AddLOVValueErrorMsgVerify",JavaWindow("Business Modeler").JavaWindow("NewLOV"))=False Then
		   'If NewLOVObject window not exist then function will Exit
		   Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail:NewLOVObject Window is not Exist ")
		   Exit Function
	   End If
		'Creating Object of NewLOVObject window
		Set  ObjNewLOVWindow=Fn_UI_ObjectCreate("Fn_BMIDE_AddLOVValueErrorMsgVerify",JavaWindow("Business Modeler").JavaWindow("NewLOV"))
		If strName<>"" Then
			strPrototype= Fn_Edit_Box_GetValue("Fn_BMIDE_AddLOVValueErrorMsgVerify",ObjNewLOVWindow,"Name")
			strName=strPrototype+strName
			'Setting Name to New LOV Object
			Call Fn_Edit_Box("Fn_BMIDE_AddLOVValueErrorMsgVerify",ObjNewLOVWindow,"Name",strName)
		End If
		If strDesc<>"" Then
			'Setting Description to New LOV Object
			Call Fn_Edit_Box("Fn_BMIDE_AddLOVValueErrorMsgVerify",ObjNewLOVWindow,"Description",strDesc)
		End If
		If strType<>"" Then
			'Setting Type to New LOV Object
			Call Fn_Edit_Box("Fn_BMIDE_AddLOVValueErrorMsgVerify",ObjNewLOVWindow,"Type",strType)
		End If
		If bUsage<>"" Then
			Call Fn_UI_Object_SetTOProperty("Fn_BMIDE_AddLOVValueErrorMsgVerify",ObjNewLOVWindow.JavaRadioButton("Usage"),"attached text",bUsage)
			'Setting Usage to New LOV Object
			Call Fn_UI_JavaRadioButton_SetON("Fn_BMIDE_AddLOVValueErrorMsgVerify",ObjNewLOVWindow, "Usage")
		End If
		Call Fn_Button_Click("Fn_BMIDE_AddLOVValueErrorMsgVerify", ObjNewLOVWindow, "Add")
 End If
	If strValue<>"" Then
		'Setting LOV Value
		 Call Fn_Edit_Box("Fn_BMIDE_AddLOVValueErrorMsgVerify",JavaWindow("Business Modeler").JavaWindow("AddLOVValue"),"Value",strValue)				
	End If
	If strDisplayName<>"" Then
		'Setting LOV Value Display Name
		Call Fn_Edit_Box("Fn_BMIDE_AddLOVValueErrorMsgVerify",JavaWindow("Business Modeler").JavaWindow("AddLOVValue"),"ValueDisplayName",strDisplayName)
	End If
	If strValueDesc<>"" Then
		'Setting Value Description
		Call Fn_Edit_Box("Fn_BMIDE_AddLOVValueErrorMsgVerify",JavaWindow("Business Modeler").JavaWindow("AddLOVValue"),"Description",strValueDesc)			
	End If
	If strCondition<>"" Then
		'Setting LOV Condition
		Call Fn_Edit_Box("Fn_BMIDE_AddLOVValueErrorMsgVerify",JavaWindow("Business Modeler").JavaWindow("AddLOVValue"),"Condition",strCondition)				
	End If
	If strErrorMsg<>"" Then
		strErrorMessage=Fn_Edit_Box_GetValue("Fn_BMIDE_AddLOVValueErrorMsgVerify",JavaWindow("Business Modeler").JavaWindow("AddLOVValue"),"Create")
        If Instr(strErrorMessage,strErrorMsg) > 0 Then
			bFlag=True
		Else
			GBL_ACTUAL_MESSAGE=strErrorMessage
		End If
	End If
	Call Fn_Button_Click("Fn_BMIDE_AddLOVValueErrorMsgVerify", JavaWindow("Business Modeler").JavaWindow("AddLOVValue"), "Cancel")
	If bFlag=True Then
		Fn_BMIDE_AddLOVValueErrorMsgVerify=True
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Pass:Successfully Verify Error Mesaage")
	End If
	If Fn_UI_ObjectExist("Fn_BMIDE_AddLOVValueErrorMsgVerify",JavaWindow("Business Modeler").JavaWindow("NewLOV"))=True Then
		'Clicking On Finish Button create New LOV Object
		Call Fn_Button_Click("Fn_BMIDE_AddLOVValueErrorMsgVerify", ObjNewLOVWindow, "Cancel")
	End If
    'Releasing Object of NewLOVObject window
	Set ObjNewLOVWindow=Nothing
End Function
'-------------------------------------------------------------------Function Used to Verify Error Messages which comes while Adding Sub LOVs----------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_SubLOVErrorMessageVerify

'Description			 :	Function Used to Verify Error Messages which comes while Adding Sub LOVs

'Parameters			   :	1.strValue: value name on which have to add SubLOV to Verify error Message
										'2.strLOV:LOV Name
										'3.strCondition: Condition
										'4.strErrorMsg:expected Error Message

'Return Value		   : 	True Or False

'Pre-requisite			:	Should be Log In BMIDE

'Examples				:	strErrMsg= "Cyclic reference of LOV ""D3CA_Cities_61132"" in value hierarchy."
'										Call Fn_BMIDE_SubLOVErrorMessageVerify("Los Angeles","D3CA_Cities_61132","isTrue",strErrMsg)
'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				24/11/2010			           1.0																					Sunny R
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_BMIDE_SubLOVErrorMessageVerify(strValue,strLOV,strCondition,strErrorMsg)
		GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_SubLOVErrorMessageVerify"
		Dim bFlag,strErrorMessage,WshShell
		GBL_EXPECTED_MESSAGE=strErrorMsg
		bFlag=False
		Fn_BMIDE_SubLOVErrorMessageVerify=False
        Call Fn_CheckBox_Set("Fn_BMIDE_SubLOVErrorMessageVerify", JavaWindow("Business Modeler"), "ShowCascadingView", "ON") 
		If strValue<>"" Then
			' Select item from LOVTable tree
			Call Fn_JavaTree_Select("Fn_BMIDE_SubLOVErrorMessageVerify", JavaWindow("Business Modeler"), "LOVTable",strValue)
		End If
	   ' Click Add Sub Lov Button
		Call Fn_Button_Click("Fn_BMIDE_SubLOVErrorMessageVerify",JavaWindow("Business Modeler"),"AddSubLOV")
		If strLOV<> "" Then
			' Set Sub Lov Name
			Call Fn_Edit_Box("Fn_BMIDE_SubLOVErrorMessageVerify", JavaWindow("Business Modeler").JavaWindow("SubLOVAttachment"),"LOV",strLOV)
				wait 1
				Set WshShell = CreateObject("WScript.Shell")
				WshShell.SendKeys "{ESC}"
				wait 1,500
				Set WshShell = nothing
		End If
		'Set Condition
		If  strCondition<>"" Then
				Call Fn_Edit_Box("Fn_BMIDE_SubLOVErrorMessageVerify", JavaWindow("Business Modeler").JavaWindow("SubLOVAttachment"),"Condition",strCondition)
		End If
		If strErrorMsg<>"" Then
            strErrorMessage=Fn_Edit_Box_GetValue("Fn_BMIDE_AddLOVValueErrorMsgVerify", JavaWindow("Business Modeler").JavaWindow("SubLOVAttachment"),"LOVSelection")
			'Matching The Error Message With Existing Error Message
			If Trim(strErrorMessage)=Trim(strErrorMsg) Then
				bFlag=True
			Else
				GBL_ACTUAL_MESSAGE=strErrorMessage
			End If
		End If
		If bFlag=True Then
			Fn_BMIDE_SubLOVErrorMessageVerify=True
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Pass: Successfully verified Error Message")
		End If
		' Click Finish Button
		Call Fn_Button_Click("Fn_BMIDE_SubLOVErrorMessageVerify",JavaWindow("Business Modeler").JavaWindow("SubLOVAttachment"),"Cancel")
End Function
'-------------------------------------------------------------------Function Used to Import External Project----------------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_ImportProject

'Description			 :	Function Used to Import External Project

'Parameters			   :	1.strProjectContentPath: Path of Project which have to Import
										'2.strActiveFile:Active File Name
										'3.strDependentTempDirectoryPath: Dependent Template Directory Location(Path)
										'4.strCodeOPPath:Code output path where want to save Code Output

'Return Value		   : 	True Or False

'Pre-requisite			:	Should be Log In BMIDE

'Examples				: 	'1. Call Fn_BMIDE_ImportProject("C:\Siemens\Teamcenter8\bmide\workspace\8000.3.0\Trial1","","C:\Siemens\Teamcenter8\bmide\templates","")

'History					 :		
'													Developer Name												Date						Rev. No.						Changes Done														Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				24/11/2010			           1.0																																					Sunny R
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_BMIDE_ImportProject(strProjectContentPath,strActiveFile,strDependentTempDirectoryPath,strCodeOPPath)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_ImportProject"
   'Variable Declaration
	Dim ObjImpProjectDialog
	Dim strMenu
	Fn_BMIDE_ImportProject=False
	'Checking Existance of "ImportProject" window
   If Fn_UI_ObjectExist("Fn_BMIDE_ImportProject", JavaWindow("Business Modeler").JavaWindow("ImportProject"))=False Then
       strMenu=Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\BMIDEConfig\BMIDE_Menu.xml", "ImportProject")
	   'Calling Import... Menu
       Call Fn_BMIDE_MenuOperation("Select", strMenu)
   End If
   'Creating Object of "ImportProject" window
	Set ObjImpProjectDialog=Fn_UI_ObjectCreate("Fn_BMIDE_ImportProject", JavaWindow("Business Modeler").JavaWindow("ImportProject"))
	'Expanding Business Modeler IDE Node
    Call Fn_UI_JavaTree_Expand("Fn_BMIDE_ImportProject", ObjImpProjectDialog, "ImportSourceTree","Business Modeler IDE")
	'Selecting Project
	Call Fn_JavaTree_Select("Fn_BMIDE_ImportProject", ObjImpProjectDialog,"ImportSourceTree","Business Modeler IDE:Import a Business Modeler IDE Template Project")
	'Clicking on Next button
	Call Fn_Button_Click("Fn_BMIDE_ImportProject", ObjImpProjectDialog, "Next")
	If strProjectContentPath<>"" Then
		'Setting Project Content path
		'Call Fn_Edit_Box("Fn_BMIDE_ImportProject",ObjImpProjectDialog,"ProjectContentsPath",strProjectContentPath)
		Call Fn_SISW_UI_JavaEdit_Operations("Fn_BMIDE_ImportProject", "SetExt", ObjImpProjectDialog, "ProjectContentsPath", strProjectContentPath)
	End If
	If strActiveFile<>"" Then
		'Selecting Active File
		Call Fn_List_Select("Fn_BMIDE_ImportProject", ObjImpProjectDialog, "ActiveFileList",strActiveFile)
	End If
	If strDependentTempDirectoryPath<>"" Then
		'Setting Dependent Template Directory Path
		Call Fn_Edit_Box("Fn_BMIDE_ImportProject",ObjImpProjectDialog,"DependentTemplatesDirectory",strDependentTempDirectoryPath)
	End If
	If strCodeOPPath<>"" Then
		'Setting Code Output Path
		Call Fn_Edit_Box("Fn_BMIDE_ImportProject",ObjImpProjectDialog,"CodeOutputLocation",strCodeOPPath)
	End If
	'Clicking On Finish Button to Finish Import project operation
	Call Fn_Button_Click("Fn_BMIDE_ImportProject", ObjImpProjectDialog, "Finish")
	If JavaWindow("Business Modeler").Exist(iTime)=True Then
		Fn_BMIDE_ImportProject=True
	End If
	'Releasing Object of Projecct Dialog
	Set ObjImpProjectDialog=Nothing
End Function
'-------------------------------------------------------------------Function Used to Close Unwanted Dialog which are Exist on Window----------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_CloseDialogs

'Description			 :	Function Used to Close Unwanted Dialog which are Exist on Window

'Return Value		   : 	True Or False

'Pre-requisite			:	Should be Log In BMIDE

'Examples				: 	'1. Call Fn_BMIDE_CloseDialogs()

'History					 :		
'													Developer Name												Date						Rev. No.						Changes Done																Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				25/11/2010			           1.0																																													Sunny R
'													Sandeep N										   				26/11/2010			           1.0							Added Call For Close All Menu O/P							  Sunny R
'													Sandeep N										   				30/11/2010			           1.0							Added Code For Setting Index Of Trees					  Sunny R
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_BMIDE_CloseDialogs()
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_CloseDialogs"
   Dim ObjChild,ObjModelerChild,strTabName
   Dim bFlag,iCount,strButtonLabel,strMenu,BReturn,strProjectName
   Dim strExePath, strWorkspace
   Dim objViewConfiguration, objNewProject, objBmideGenericWindowToCloseDialogs, objPropertiesOfObject, objNewOperationInputProperty, objAddLovValue, objNewLOV 
   Fn_BMIDE_CloseDialogs=False
   bFlag=False
   BReturn=False
   
     If JavaWindow("Business Modeler").JavaWindow("DataModelMergeDialog").Exist(2)=True Then
       JavaWindow("Business Modeler").JavaWindow("DataModelMergeDialog").JavaButton("Abort").Click
    End If    
'If JavaWindow("Business Modeler").GetROProperty("enabled")=0 Then
	'If JavaWindow("Business Modeler").Exist(3) = False Then
		Set ObjChild=Description.Create()
		ObjChild("Class Name").value="JavaButton"
		'ObjChild("label").value="Cancel"
		Set ObjModelerChild=JavaWindow("Business Modeler").ChildObjects(ObjChild)
		For iCount=ObjModelerChild.Count-1 To 0 STEP -1
				strButtonLabel=ObjModelerChild(iCount).GetROProperty("label")
				If LCase(strButtonLabel)="cancel" OR  LCase(strButtonLabel)="close" OR  LCase(strButtonLabel)="ok" Then
					ObjModelerChild(iCount).Click
					bFlag=True
				End If
		Next
	'Else
		bFlag=True
	'End If 
	
	'Need to handle launching of BMIDE - as it gets closed abruptly during batch execution
    If JavaWindow("Business Modeler").Exist(5) = FALSE Then
   		strExePath = Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\BMIDEConfig\BMIDE_Config.xml", "BMIDEExecutable")
		strWorkspace = Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\BMIDEConfig\BMIDE_Config.xml", "BMIDEWorkspacePath")
   		BReturn =  Fn_LaunchBMIDE(strExePath, strWorkspace)
   		wait(5)
    End If
    
   
	'To close new project window
	Set objNewProject=Fn_SISW_BMIDE_GetObject("NewProject")
	If objNewProject.Exist(5) Then 
		objNewProject.Close
   End If
   
	Set objViewConfiguration=Fn_SISW_BMIDE_GetObject("ViewConfiguration")
   If objViewConfiguration.Exist(1) Then
		objViewConfiguration.Close	
		If objNewProject.Exist(1) Then 
			objNewProject.Close
		End If
   End If
    
   'To close any Generic Jwindow,Property attachment dialog,add LOV Value window
	Set objBmideGenericWindowToCloseDialogs = Fn_SISW_BMIDE_GetObject("BmideGenericWindowToCloseDialogs")
    If objBmideGenericWindowToCloseDialogs.Exist(1) Then
    	objBmideGenericWindowToCloseDialogs.Close
    End If
       
	'To close NewOperationInputProperty dialog
   Set objNewOperationInputProperty = Fn_SISW_BMIDE_GetObject("NewOperationInputProperty")
   If objNewOperationInputProperty.Exist(1) Then
   	objNewOperationInputProperty.Close
   End If
    
   'To close classic LOV>add LOV
   Set objAddLovValue=Fn_SISW_BMIDE_GetObject("AddLOVValue")
   Set objNewLOV=Fn_SISW_BMIDE_GetObject("NewLOV")
   If objAddLovValue.Exist(1) Then
		objAddLovValue.Close	
		If objNewLOV.Exist(1) Then 
			objNewLOV.Close
   	End If
   End If
   
    
	Call Fn_SISW_BMIDE_ToolbarButtonOperations("Select", 1, "Advanced", "")
	wait(1)
	If CInt(JavaWindow("Business Modeler").JavaTab("MainTab").GetROProperty("items count"))=2 Then
		If JavaWindow("Business Modeler").JavaTab("MainTab").GetROProperty("value")="Help" Then
			Call Fn_BMIDE_TabOperations("Main","Close","Help")
		End If
		If JavaWindow("Business Modeler").JavaTab("MainTab").GetROProperty("value")="Welcome" Then
			Call Fn_BMIDE_TabOperations("Main","Close","Welcome")
		End If
	ElseIf CInt(JavaWindow("Business Modeler").JavaTab("MainTab").GetROProperty("items count"))=1 Then
		If JavaWindow("Business Modeler").JavaTab("MainTab").GetROProperty("value")="Help" Then
			Call Fn_BMIDE_TabOperations("Main","Close","Help")
		ElseIf JavaWindow("Business Modeler").JavaTab("MainTab").GetROProperty("value")="Welcome" Then
			Call Fn_BMIDE_TabOperations("Main","Close","Welcome")
		ElseIf JavaWindow("Business Modeler").JavaTab("MainTab").GetROProperty("value")="BMIDE Assistant" Then
			Call Fn_BMIDE_TabOperations("Main","Close","BMIDE Assistant")
		End If
	End If

	If CInt(JavaWindow("Business Modeler").JavaTab("MainTab").GetROProperty("items count"))>0 Then
		If JavaWindow("Business Modeler").JavaTab("MainTab").GetROProperty("value")="Help" Then
			Call Fn_BMIDE_TabOperations("Main","Close","Help")
		ElseIf JavaWindow("Business Modeler").JavaTab("MainTab").GetROProperty("value")="Welcome" Then
			Call Fn_BMIDE_TabOperations("Main","Close","Welcome")
		ElseIf JavaWindow("Business Modeler").JavaTab("MainTab").GetROProperty("value")="BMIDE Assistant" Then
			Call Fn_BMIDE_TabOperations("Main","Close","BMIDE Assistant")
		End If
	End If

	If CInt(JavaWindow("Business Modeler").JavaTab("MainTab").GetROProperty("items count"))>0 Then
		strTabName=JavaWindow("Business Modeler").JavaTab("MainTab").GetROProperty("value")
		JavaWindow("Business Modeler").JavaTab("MainTab").Click 5,5,"LEFT"
		Call Fn_BMIDE_TabOperations("Main","Activate",strTabName)
        strMenu=Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\BMIDEConfig\BMIDE_Menu.xml", "CloseAll")
		Call Fn_BMIDE_MenuOperation("Select", strMenu)
	End If
	Call Fn_BMIDE_TabOperations("UpperLeft","Activate","Business Objects")
	Call Fn_BMIDE_TabOperations("LowerLeft","Activate","Extensions")
	Call Fn_BMIDE_TreeIndexIdentification()

	If bFlag=True Then
		Fn_BMIDE_CloseDialogs=True
	End If
	
	Set ObjModelerChild=Nothing
	Set ObjChild=Nothing
End Function
'-------------------------------------------------------------------Function Used to Verify Error Message----------------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_ErrorMessageVerify

'Description			 :	Function Used to Verify Error Message

'Parameters			   :	1.strDialogName: Error Dialog Name
										'2.strErrorMsg:Expected Error Message
										'3.strButtton: Button Name

'Return Value		   : 	True Or False

'Pre-requisite			:	Error Message Should be Appear on Screen

'Examples				: 	Call Fn_BMIDE_ErrorMessageVerify("Operation not allowed.","LOV Value Gold is referenced by Filter LOV: D3Material_Metal;","OK")
'										Call Fn_BMIDE_ErrorMessageVerify("Operation not allowed.","LOV Value Plastic has sub LOV attached","OK")

'History					 :		
'													Developer Name												Date						Rev. No.						Changes Done														Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				26/11/2010			           1.0																																					Sunny R
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_BMIDE_ErrorMessageVerify(strDialogName,strErrorMsg,strButtton)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_ErrorMessageVerify"
   'Variablr Declaration
   Dim strMsg
   Fn_BMIDE_ErrorMessageVerify=False
   GBL_EXPECTED_MESSAGE=strErrorMsg
   'Setting Dialog Title
   Call Fn_UI_Object_SetTOProperty("Fn_BMIDE_ErrorMessageVerify",JavaWindow("Business Modeler").Dialog("ErrorDialog"),"text",strDialogName)
   'Checking Existance Of Error Dialog
   If Fn_UI_ObjectExist("Fn_BMIDE_ErrorMessageVerify", JavaWindow("Business Modeler").Dialog("ErrorDialog"))=True Then
	   'Retriving Error Message which Apprers on Dialog
	  strMsg=Fn_UI_Object_GetROProperty("Fn_BMIDE_ErrorMessageVerify",JavaWindow("Business Modeler").Dialog("ErrorDialog").Static("ErrorMessage"),"text")
	  'Verifying Error Message Match With Expected Error Message
		If InStr(1,strMsg,strErrorMsg)>=1 Then
		  'Function Returns true
		  Fn_BMIDE_ErrorMessageVerify=True
		Else
			GBL_ACTUAL_MESSAGE=strMsg
		End If
	  'Clicking On strButtton Button
	  Call Fn_UI_WinButton_Click("Fn_BMIDE_ErrorMessageVerify", JavaWindow("Business Modeler").Dialog("ErrorDialog"), strButtton,"","","")
   End If   
End Function

'-------------------------------------------------------------------Function Used to Set Specific View---------------------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_SetView

'Description			 :	Function Used to Set Specific View

'Parameters			   :	1.strViewName:View Name Which Have to Set (Full Node Name)

'Return Value		   : 	True Or False

'Pre-requisite			:	Should be Log In BMIDE

'Examples				: Call Fn_BMIDE_SetView("Business Modeler IDE:Business Objects")

'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				26/11/2010			           1.0																					Sunny R
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_BMIDE_SetView(strViewName)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_SetView"
   'Variable Declaration 
   Dim ObjShowViewDialog
   Dim strMenu,arrViewName
   Fn_BMIDE_SetView=False
   'Checking Existance Of ShowView Dialog
   If Fn_UI_ObjectExist("Fn_BMIDE_SetView", JavaWindow("Business Modeler").JavaWindow("ShowView"))=False Then
       strMenu=Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\BMIDEConfig\BMIDE_Menu.xml", "ShowView")
	   'Calling "Window:Show View:Other..." Menu
	  Call Fn_BMIDE_MenuOperation("Select", strMenu)
   End If
   'Creating object of ShowView Dialog
	Set ObjShowViewDialog=Fn_UI_ObjectCreate("Fn_BMIDE_SetView", JavaWindow("Business Modeler").JavaWindow("ShowView"))
	arrViewName=Split(strViewName,":")
	'Expanding Root Node 
    Call Fn_UI_JavaTree_Expand("Fn_BMIDE_SetView", ObjShowViewDialog, "ViewTree",arrViewName(0))
	'Selecting View 
	Call Fn_JavaTree_Select("Fn_BMIDE_SetView", ObjShowViewDialog, "ViewTree",strViewName)
	'clicking On OK button to Set View
    Call Fn_Button_Click("Fn_BMIDE_SetView", ObjShowViewDialog, "OK")
	Fn_BMIDE_SetView=True
	'Releasing Object Of ShowView Dialog
	Set ObjShowViewDialog=Nothing
End Function
'-------------------------------------------------------------------Function Used to Set Specific View---------------------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_SetView

'Description			 :	Function Used to Set Specific View

'Parameters			   :	1.strViewName:View Name Which Have to Set (Full Node Name)

'Return Value		   : 	True Or False

'Pre-requisite			:	Should be Log In BMIDE

'Examples				: Call Fn_BMIDE_SetView("Business Modeler IDE:Business Objects")

'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				26/11/2010			           1.0																					Sunny R
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Public Function Fn_BMIDE_SetPerspective(strPerspectiveName)
'   'Variable Declaration 
'   Dim ObjShowViewDialog
'   Dim strMenu,arrViewName
'   Fn_BMIDE_SetView=False
'   'Checking Existance Of ShowView Dialog
'   If Fn_UI_ObjectExist("Fn_BMIDE_SetView", JavaWindow("Business Modeler").JavaWindow("ShowView"))=False Then
'       strMenu=Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\BMIDE_Menu.xml", "ShowView")
'	   'Calling "Window:Show View:Other..." Menu
'	  Call Fn_BMIDE_MenuOperation("Select", strMenu)
'   End If
'   'Creating object of ShowView Dialog
'	Set ObjShowViewDialog=Fn_UI_ObjectCreate("Fn_BMIDE_SetView", JavaWindow("Business Modeler").JavaWindow("ShowView"))
'	arrViewName=Split(strViewName,":")
'	'Expanding Root Node 
'    Call Fn_UI_JavaTree_Expand("Fn_BMIDE_SetView", ObjShowViewDialog, "ViewTree",arrViewName(0))
'	'Selecting View 
'	Call Fn_JavaTree_Select("Fn_BMIDE_SetView", ObjShowViewDialog, "ViewTree",strViewName)
'	'clicking On OK button to Set View
'    Call Fn_Button_Click("Fn_BMIDE_SetView", ObjShowViewDialog, "OK")
'	Fn_BMIDE_SetView=True
'	'Releasing Object Of ShowView Dialog
'	Set ObjShowViewDialog=Nothing
'End Function
'-------------------------------------------------------------------Function Used to Create New Server Connection Profile---------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_AddServerConnectionProfile

''Description			:1. Function Used to Create New Server Connection Profile

'Return Value		   : 	True or False

'Parameters     		:	1. strProfileName : Server Connection Profile Name
										'2. strProtocol : Protocol Name for Connection
										'3. strHost: Host Name
										'4. strPort: Port Number
										'5. strAppName: Application Name
										'6. strUserID: User ID
										'7. strGroup: User Group Name
										'8. strRole: User Role

'Pre-requisite			:	BMIDE Prespective should be Open.

'Examples				: Call Fn_BMIDE_AddServerConnectionProfile("DemoProfile","HTTP","pnv6s106","7001","tc","AutoTestDBA","dba","DBA")
'									  Call Fn_BMIDE_AddServerConnectionProfile("TestProfile","IIOP","pnv6s106","8001","tc","AutoTestDBA","dba","DBA")

'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				29/11/2010			           1.0																						Sunny R
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_BMIDE_AddServerConnectionProfile(strProfileName,strProtocol,strHost,strPort,strAppName,strUserID,strGroup,strRole)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_AddServerConnectionProfile"
   'Variable Declaration
   Dim strMenu,bFlag
   Dim ObjConnectionWnd
   bFlag=False
	Fn_BMIDE_AddServerConnectionProfile=False
	'Checking Existance Of "Preferences" Dialog
   If Fn_UI_ObjectExist("Fn_BMIDE_AddServerConnectionProfile",JavaWindow("Business Modeler").JavaWindow("Preferences"))=False Then
	   'Calling Window:Preference Menu to open Preference Dialog
		strMenu=Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\BMIDEConfig\BMIDE_Menu.xml", "Preferences")
		Call Fn_BMIDE_MenuOperation("Select", strMenu)
   End If
   'Expanding "Teamcenter" Node From "PreferenceTree"
	Call Fn_UI_JavaTree_Expand("Fn_BMIDE_AddServerConnectionProfile",JavaWindow("Business Modeler").JavaWindow("Preferences"),"PreferenceTree","Teamcenter")
	'Selecting "Teamcenter:Server Connection Profiles" node 
	Call Fn_JavaTree_Select("Fn_BMIDE_AddServerConnectionProfile",JavaWindow("Business Modeler").JavaWindow("Preferences"),"PreferenceTree","Teamcenter:Server Connection Profiles")
	bFlag=Fn_UI_ListItemExist("Fn_BMIDE_AddServerConnectionProfile", JavaWindow("Business Modeler").JavaWindow("Preferences"),"ListOfProfiles",strProfileName)
	If bFlag=True Then
        Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Profile [" + strProfileName +"] Is Already Exist In Profile List")
		Call Fn_Button_Click("Fn_BMIDE_AddServerConnectionProfile", JavaWindow("Business Modeler").JavaWindow("Preferences"), "OK")
		Fn_BMIDE_AddServerConnectionProfile=True
		Exit Function
	End If
	'Clicking "Add" Button To open "TeamcenterRepositoryConnection" Dialog
	Call Fn_Button_Click("Fn_BMIDE_AddServerConnectionProfile", JavaWindow("Business Modeler").JavaWindow("Preferences"), "Add")
	'Creating Object  of "TeamcenterRepositoryConnection" Dialog
	Set ObjConnectionWnd=Fn_UI_ObjectCreate("Fn_BMIDE_AddServerConnectionProfile",JavaWindow("Business Modeler").JavaWindow("TeamcenterRepositoryConnection"))
	If strProfileName<>"" Then
		'Setting Profile Name
	   Call Fn_Edit_Box("Fn_BMIDE_AddServerConnectionProfile",ObjConnectionWnd,"ServerConnectionProfile",strProfileName)
	End If
	If strProtocol<>"" Then
		'Selecting Protocol
	   Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_BMIDE_AddServerConnectionProfile",ObjConnectionWnd.JavaRadioButton("Protocol"),"attached text",strProtocol)
	   Call Fn_UI_JavaRadioButton_SetON("Fn_BMIDE_AddServerConnectionProfile",ObjConnectionWnd, "Protocol")
	End If
	 If strHost<>"" Then
	   Call Fn_Edit_Box("Fn_BMIDE_AddServerConnectionProfile",ObjConnectionWnd,"Host",strHost)
	End If
	If strPort<>"" Then
		'Setting Host Name
	   Call Fn_Edit_Box("Fn_BMIDE_AddServerConnectionProfile",ObjConnectionWnd,"Port",strPort)
	End If
	If strAppName<>"" Then
		'Setting Application Name
	   Call Fn_Edit_Box("Fn_BMIDE_AddServerConnectionProfile",ObjConnectionWnd,"ApplicationName",strAppName)
	End If
	If strUserID<>"" Then
		'Setting User ID
	   Call Fn_Edit_Box("Fn_BMIDE_AddServerConnectionProfile",ObjConnectionWnd,"UserID",strUserID)
	End If
	If strGroup<>"" Then
		'Setting Group
	   Call Fn_Edit_Box("Fn_BMIDE_AddServerConnectionProfile",ObjConnectionWnd,"Group",strGroup)
	End If
	If strRole<>"" Then
		'Setting Role
	   Call Fn_Edit_Box("Fn_BMIDE_AddServerConnectionProfile",ObjConnectionWnd,"Role",strRole)
	End If
	'Clicking On Finish Button to Create New Profile
	Call Fn_Button_Click("Fn_BMIDE_AddServerConnectionProfile",ObjConnectionWnd, "Finish")
	'Verifyng Profile Is Created Or Not
	bFlag=Fn_UI_ListItemExist("Fn_BMIDE_AddServerConnectionProfile", JavaWindow("Business Modeler").JavaWindow("Preferences"), "ListOfProfiles",strProfileName)
	If bFlag=True Then
		'Function Returns True
        Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS:Successfully Created Profile Of Name [" + strProfileName +"]")				
		Fn_BMIDE_AddServerConnectionProfile=True
	Else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail: Failed to Create Profile")				
	End If
	Call Fn_Button_Click("Fn_BMIDE_AddServerConnectionProfile", JavaWindow("Business Modeler").JavaWindow("Preferences"), "OK")
	'Releasing Object Connection Dialog
	Set ObjConnectionWnd=Nothing
End Function
'-------------------------------------------------------------------Function Used to Deploy Project------------------------------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_DeployProject

''Description			:1.Function Used to Deploy Project

'Return Value		   : 	True or False

'Parameters     		:	1. StrProjectName : Project Name For Deploy
										'2. StrProfile : Profile
										'3. StrUserID: User ID ()
										'4. StrPassword: Password
										'5. StrGroup: User Group Name
										'6. StrRole: User Role

'Pre-requisite			:	BMIDE Prespective should be Open.

'Examples				: Call Fn_BMIDE_DeployProject("Trail","DemoProfile","","AutoTestDBA","dba","DBA")
'									Call Fn_BMIDE_DeployProject("TestProject","DemoProfile","","AutoTestDBA","dba","DBA")

'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done									Reviewer
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				29/11/2010			           1.0																							Sunny R
'													Sandeep N										   				24/08/2011			           1.1					Added Code to handle Backup Dialog		 Sunny R
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_BMIDE_DeployProject(StrProjectName,StrProfile,StrUserID,StrPassword,StrGroup,StrRole)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_DeployProject"
   Dim strMenu,i
   Dim ObjDiployWnd
   Fn_BMIDE_DeployProject=False
	If Fn_UI_ObjectExist("Fn_BMIDE_DeployProject", JavaWindow("Business Modeler").JavaWindow("Deploy"))=False Then
			'Checking Existance Of "NewProject" Dialog
			 If Fn_UI_ObjectExist("Fn_BMIDE_DeployProject",JavaWindow("Business Modeler").JavaWindow("NewProject"))=False Then
			   'Calling File:New:Other... Menu to open New Project  Dialog
				strMenu=Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\BMIDEConfig\BMIDE_Menu.xml", "FileNewOther")
				Call Fn_BMIDE_MenuOperation("Select", strMenu)
		   End If
			'Expanding ""Business Modeler IDE"" Node From "WizardsTree"
			Call Fn_UI_JavaTree_Expand("Fn_BMIDE_DeployProject",JavaWindow("Business Modeler").JavaWindow("NewProject"),"WizardsTree","Business Modeler IDE")
			'Selecting "Business Modeler IDE:Deployment"" node 
			Call Fn_JavaTree_Select("Fn_BMIDE_DeployProject",JavaWindow("Business Modeler").JavaWindow("NewProject"),"WizardsTree","Business Modeler IDE:Deployment")
			'Clicking "Next" Button To open "Deploy" Dialog
			Call Fn_Button_Click("Fn_BMIDE_DeployProject", JavaWindow("Business Modeler").JavaWindow("NewProject"), "Next")
	End If
	For i=0 to 7			'//			Added code by Priyanka B for wait till Deploy Window Appears 				// Date : 19-Feb-2013
			If Fn_UI_ObjectExist("Fn_BMIDE_DeployProject",JavaWindow("Business Modeler").JavaWindow("Deploy")) = False Then
				wait(20)
			Else
				Exit For
			End If
		 Next
	'Creating Object Of Deploy Window
	Set ObjDiployWnd=Fn_UI_ObjectCreate("Fn_BMIDE_DeployProject", JavaWindow("Business Modeler").JavaWindow("Deploy"))

	If ObjDiployWnd.JavaCheckBox("BackupDuringShutdownOfBMIDE").Exist(3) Then
		Call Fn_Button_Click("Fn_BMIDE_DeployProject", ObjDiployWnd, "Next")
	End If
	
	If StrProjectName<>"" Then
		'Selecting Project For Deploy
		Call Fn_List_Select("Fn_BMIDE_DeployProject", ObjDiployWnd,"Project",StrProjectName)
	End If
	If StrProfile<>"" Then
		'Selecting Server Profile For Deploy
		Call Fn_List_Select("Fn_BMIDE_DeployProject", ObjDiployWnd,"ServerProfile",StrProfile)
	End If
	If ObjDiployWnd.JavaEdit("UserID").Exist(2) Then
		If ObjDiployWnd.JavaEdit("UserID").GetROProperty("enabled") = "1" Then
			If StrUserID<>"" Then
				'Setting user ID
				Call Fn_Edit_Box("Fn_BMIDE_DeployProjectWithDetailVerifications",ObjDiployWnd,"UserID",StrUserID)
			Else
				arrUser = Split(Environment.Value("TcUserDBA"),":")
				Call Fn_Edit_Box("Fn_BMIDE_DeployProjectWithDetailVerifications",ObjDiployWnd,"UserID",arrUser(0))
			End If
		End If
	End If
	If  Fn_UI_ObjectExist("Fn_BMIDE_DeployProject", JavaWindow("Business Modeler").JavaWindow("Deploy").JavaEdit("Password"))=True Then
		If StrPassword<>"" Then
			'Setting Password
			 Call Fn_Edit_Box("Fn_BMIDE_DeployProject",ObjDiployWnd,"Password",StrPassword)
		End If
	Else
		'- - - - - - - Added code to Select [ Generate Client Cache and Generate Server Cache ] option
		Call Fn_CheckBox_Set("Fn_BMIDE_DeployProject", ObjDiployWnd, "GenerateClientCache", "ON")
		Call Fn_CheckBox_Set("Fn_BMIDE_DeployProject", ObjDiployWnd, "GenerateServerCache", "ON")
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Call Fn_Button_Click("Fn_BMIDE_DeployProject",ObjDiployWnd, "Finish")
'		  For i=0 to 30
'			If Fn_UI_ObjectExist("Fn_BMIDE_DeployProject",JavaWindow("Business Modeler").JavaWindow("Deploying to Teamcenter")) = True Then
'				wait(20)
'			Else
'				Exit For
'			End If
'		 Next
			Do while JavaWindow("Business Modeler").JavaWindow("Deploying to Teamcenter").Exist(2)
				wait 10
			Loop
			If  Fn_UI_ObjectExist("Fn_BMIDE_DeployProject",JavaWindow("Business Modeler").JavaWindow("Deployment Complete")) = True Then
				Call Fn_Button_Click("Fn_BMIDE_DeployProject",JavaWindow("Business Modeler").JavaWindow("Deployment Complete"), "OK")
			End If
		Fn_BMIDE_DeployProject=True
		Set ObjDiployWnd=Nothing
		Exit Function
	End If
	If StrGroup<>"" Then
		'Setting Password
		 Call Fn_Edit_Box("Fn_BMIDE_DeployProject",ObjDiployWnd,"Group",StrGroup)
	End If
	If StrRole<>"" Then
		'Setting Password
		 Call Fn_Edit_Box("Fn_BMIDE_DeployProject",ObjDiployWnd,"Role",StrRole)
	End If
	If Fn_UI_ObjectExist("Fn_BMIDE_DeployProject", ObjDiployWnd.JavaButton("Connect"))=True Then
	'Clicking "Connect" Button To Connect To The Host
	Call Fn_Button_Click("Fn_BMIDE_DeployProject",ObjDiployWnd, "Connect")
	End If
	JavaWindow("Business Modeler").JavaWindow("Deploy").JavaButton("Finish").WaitProperty "enabled",1,60000
	'Clicking "Finish" Button
	'- - - - - - - Added code to Select [ Generate Client Cache and Generate Server Cache ] option
	Call Fn_CheckBox_Set("Fn_BMIDE_DeployProject", ObjDiployWnd, "GenerateClientCache", "ON")
	Call Fn_CheckBox_Set("Fn_BMIDE_DeployProject", ObjDiployWnd, "GenerateServerCache", "ON")
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	Call Fn_Button_Click("Fn_BMIDE_DeployProject",ObjDiployWnd, "Finish")
	'Checking Existance Of Save And Deploy dialog
	If Fn_UI_ObjectExist("Fn_BMIDE_DeployProject",JavaWindow("Business Modeler").JavaWindow("ConfirmSaveAndDeployment"))=True Then
		'Clicking "Connect" Button To Connect To The Host
		Call Fn_Button_Click("Fn_BMIDE_DeployProject",JavaWindow("Business Modeler").JavaWindow("ConfirmSaveAndDeployment"), "OK")
	End If
'    For i=0 to 30
'		If Fn_UI_ObjectExist("Fn_BMIDE_DeployProject",JavaWindow("Business Modeler").JavaWindow("Deploying to Teamcenter")) = True Then
'			wait(20)
'		Else
'			Exit For
'		End If
'     Next
	Do while JavaWindow("Business Modeler").JavaWindow("Deploying to Teamcenter").Exist(2)
		wait 10
	Loop
	If  Fn_UI_ObjectExist("Fn_BMIDE_DeployProject",JavaWindow("Business Modeler").JavaWindow("Deployment Complete")) = True Then
		Call Fn_Button_Click("Fn_BMIDE_DeployProject",JavaWindow("Business Modeler").JavaWindow("Deployment Complete"), "OK")
	ElseIf Fn_UI_ObjectExist("Fn_BMIDE_DeployProject",JavaWindow("DeploymentComplete")) = True Then
		Call Fn_Button_Click("Fn_BMIDE_DeployProject",JavaWindow("DeploymentComplete"), "OK")
	End If
	Fn_BMIDE_DeployProject=True
	'Releasing Object Of Deploy Eindow
	Set ObjDiployWnd=Nothing
End Function 

'-------------------------------------------------------------------Function Used to Perform Operations On Localization table-----------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_LocalizationOperations

''Description			:1.Function Used to Perform Operations On Localization table

'Return Value		   : 	True or Table Data (Content Name,Locale,Status) Or False

'Parameters     		:	1. strAction : Action Name
										'2. strContent : Content Name (value Localization)
										'3. strLocale: Locale Name
										'4. strStatus: Loclization Status
										'5. strErrorMsg: Error message 

'Pre-requisite			:	BMIDE Prespective & Localization Table should be Open.

'Examples				: Call Fn_BMIDE_LocalizationOperations("GetContent","","en_US","","")
'									  Call Fn_BMIDE_LocalizationOperations("GetContent","","ja_JP","","")
'									  Call Fn_BMIDE_LocalizationOperations("SelectRow","Date Last Backup","","","")
'									Call Fn_BMIDE_LocalizationOperations("Override","Date Last Backup_Ovr","","","") 
'									Call Fn_BMIDE_LocalizationOperations("Edit","Date Last Backup_Ovr","","","") 
'									Imp Note : - Localization Table Is present On 2 Tabs "Main" and Properties .
'															This Function Work On both Tab but user have to activate that tab On which he have to perform operation.
'															To Activate Inner Tabs use Function  Fn_BMIDE_InnerTabOperations()
'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				30/11/2010			           1.0																						Sunny R
'													Sandeep N										   				01/12/2010			           1.0					Added 2 Cases "SelectRow" 			Sunny R
'																																																	"Override"
'													Sandeep N										   				02/12/2010			           1.0					Added 1 Cases "Edit" 			Sunny R
'													Sandeep N										   				08/02/2013			           1.1					Added code to select Localization tab as its design chsange in TC 10.1
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_BMIDE_LocalizationOperations(strAction,strContent,strLocale,strStatus,strErrorMsg)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_LocalizationOperations"
   Dim iCounter,strLocaleName,intRowCount,bFlag,strContentName
   bFlag=False
   Fn_BMIDE_LocalizationOperations=False

   Select Case strAction
			Case "GetContent" 'This Case Return Content Name
               'Call Fn_UI_JavaTab_Select("",JavaWindow("Business Modeler"),"MainInnerTab", "Localization")
                Call Fn_UI_JavaTab_Select("",JavaWindow("Business Modeler"),"MainInnerTab2", "Localization")
				'Taking Row Count From "Localization" Table
						intRowCount=Fn_UI_Object_GetROProperty("Fn_BMIDE_LocalizationOperations",JavaWindow("Business Modeler").JavaTable("Localization"), "rows")
						For iCounter=0 To intRowCount-1
							'Taking Locale From Locale Column
							strLocaleName=JavaWindow("Business Modeler").JavaTable("Localization").GetCellData(iCounter,"Locale")
							If Trim(strLocale)=Trim(strLocaleName) Then
								'Function Returns Content Of Locale
								Fn_BMIDE_LocalizationOperations=JavaWindow("Business Modeler").JavaTable("Localization").GetCellData(iCounter,"Content")
								Exit For
							End If
						Next
			 Case "SelectRow" 'This Case use to Select Row of Localization table "strContent" is compulsory parameter for this case
						'Activating Localization tab
						Call Fn_UI_JavaTab_Select("",JavaWindow("Business Modeler"),"PropertiesInnerTab", "Localization")
						'Taking Row Count From "Localization" Table
						intRowCount=Fn_UI_Object_GetROProperty("Fn_BMIDE_LocalizationOperations",JavaWindow("Business Modeler").JavaTable("Localization"), "rows")
						For iCounter=0 To intRowCount-1
							'Taking Locale From Locale Column
							strContentName=JavaWindow("Business Modeler").JavaTable("Localization").GetCellData(iCounter,"Content")
							If Trim(strContentName)=Trim(strContent) Then
                                JavaWindow("Business Modeler").JavaTable("Localization").SelectCell iCounter,0
								bFlag=True
								Exit For
							End If
						Next 
						If bFlag=True Then
							Fn_BMIDE_LocalizationOperations=True
						End If
						
			Case "Override" 'Overriding Localization
                'Activating Localization tab
				Call Fn_UI_JavaTab_Select("",JavaWindow("Business Modeler"),"PropertiesInnerTab", "Localization")
				'Call Fn_Button_Click("Fn_BMIDE_LocalizationOperations", JavaWindow("Business Modeler"), "OverrideLocalization")
				If Fn_UI_ObjectExist("Fn_BMIDE_LocalizationOperations", JavaWindow("Business Modeler").JavaButton("OverrideLocalization"))=True Then
					JavaWindow("Business Modeler").JavaButton("OverrideLocalization").Object.click
					If Fn_UI_ObjectExist("Fn_BMIDE_LocalizationOperations", JavaWindow("Business Modeler").JavaWindow("Localization"))=False Then
						JavaWindow("Business Modeler").JavaButton("OverrideLocalization").Object.click
					End If
					If strContent<>"" Then
						'Setting Localization value(Content)
						Call Fn_Edit_Box("Fn_BMIDE_LocalizationOperations",JavaWindow("Business Modeler").JavaWindow("Localization"),"ValueLocalization",strContent)
					End If
					If strLocale<>"" Then
						'Selecting Locale From Locale List
						Call Fn_List_Select("Fn_BMIDE_LocalizationOperations",JavaWindow("Business Modeler").JavaWindow("Localization"), "Locale",strLocale)
					End If
					If strStatus<>"" Then
						'Selecting Locale From Status List
						Call Fn_List_Select("Fn_BMIDE_LocalizationOperations",JavaWindow("Business Modeler").JavaWindow("Localization"), "Status",strStatus)
					End If
					Call Fn_Button_Click("Fn_BMIDE_LocalizationOperations", JavaWindow("Business Modeler").JavaWindow("Localization"), "Finish")
				Else
					Call Fn_BMIDE_LocalizationOperations("Edit",strContent,strLocale,strStatus,"")
				End If
				Fn_BMIDE_LocalizationOperations=True
			 Case "Edit" 'Overriding Localization
                'Activating Localization tab
				Call Fn_UI_JavaTab_Select("",JavaWindow("Business Modeler"),"PropertiesInnerTab", "Localization")
				If Fn_UI_ObjectExist("Fn_BMIDE_LocalizationOperations", JavaWindow("Business Modeler").JavaWindow("Localization"))=False Then
					JavaWindow("Business Modeler").JavaButton("EditLocalization").Object.click
				End If
				If strContent<>"" Then
					'Setting Localization value(Content)
					Call Fn_Edit_Box("Fn_BMIDE_LocalizationOperations",JavaWindow("Business Modeler").JavaWindow("Localization"),"ValueLocalization",strContent)
				End If
				If strLocale<>"" Then
					'Selecting Locale From Locale List
					Call Fn_List_Select("Fn_BMIDE_LocalizationOperations",JavaWindow("Business Modeler").JavaWindow("Localization"), "Locale",strLocale)
				End If
				If strStatus<>"" Then
					'Selecting Locale From Status List
					Call Fn_List_Select("Fn_BMIDE_LocalizationOperations",JavaWindow("Business Modeler").JavaWindow("Localization"), "Status",strStatus)
				End If
				Call Fn_Button_Click("Fn_BMIDE_LocalizationOperations", JavaWindow("Business Modeler").JavaWindow("Localization"), "Finish")
				Fn_BMIDE_LocalizationOperations=True
			Case Else
						Fn_BMIDE_LocalizationOperations=False
   End Select
End Function
'-------------------------------------------------------------------Function Used to Exit BMIDE------------------------------------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_ExitBMIDE

'Description			 :	Function Used to Exit BMIDE

'Parameters			   :	1.strProjects: Project Name For Save (Separeted By Colan :)

'Return Value		   : 	True Or False

'Pre-requisite			:	Should be Log In BMIDE

'Examples				: 	'Call Fn_BMIDE_ExitBMIDE("sandeep:TestProject")
										'Call Fn_BMIDE_ExitBMIDE("")

'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				30/11/2010			           1.0																						Sunny R
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_BMIDE_ExitBMIDE(strProjects)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_ExitBMIDE"
   'Variable Declaration
   Dim strMenu,bFlag,iCounter,iRowCount,iCount,arrProjects,strPrjName
   Dim ObjSaveDataModelWnd
   bFlag=True
   Fn_BMIDE_ExitBMIDE=False
	'Checking Existance of "Business Modeler" main window
   If Fn_UI_ObjectExist("Fn_BMIDE_ExitBMIDE", JavaWindow("Business Modeler"))=True Then
		'Checking Existance Of "SaveDataModel" Window
	   If  Fn_UI_ObjectExist("Fn_BMIDE_ExitBMIDE", JavaWindow("Business Modeler").JavaWindow("SaveDataModel"))=False Then
		   strMenu=Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\BMIDEConfig\BMIDE_Menu.xml", "Exit")
		   'Selecting Menu File:Exit
			Call Fn_BMIDE_MenuOperation("Select", strMenu)
	   End If
	   'Checking Existance Of "SaveDataModel" Window
	   If Fn_UI_ObjectExist("Fn_BMIDE_ExitBMIDE", JavaWindow("Business Modeler").JavaWindow("SaveDataModel"))=True Then
			'Creating Object of "SaveDataModel" Window
			Set ObjSaveDataModelWnd=Fn_UI_ObjectCreate("Fn_BMIDE_ExitBMIDE", JavaWindow("Business Modeler").JavaWindow("SaveDataModel"))
			If strProjects<>"" Then
				'First Deselecting All Project 
				Call Fn_Button_Click("Fn_BMIDE_ExitBMIDE", ObjSaveDataModelWnd, "DeselectAll")
				'Taking Row Count Of Project Table
				iRowCount=Fn_UI_Object_GetROProperty("Fn_BMIDE_ExitBMIDE",ObjSaveDataModelWnd.JavaTable("Project"), "rows")
					arrProjects=Split(strProjects,":")
					For iCounter=0 To Ubound(arrProjects)
						bFlag=False
						For iCount=0 To iRowCount-1
							'Taking Row Data (Project name) From project Table
							strPrjName=ObjSaveDataModelWnd.JavaTable("Project").GetCellData(iCount,0)
							If Trim(arrProjects(iCounter))=Trim(strPrjName) Then
								'Selecting Row
								ObjSaveDataModelWnd.JavaTable("Project").SelectCell iCount,0
								ObjSaveDataModelWnd.JavaTable("Project").PressKey " "
								bFlag=True
								Exit For
							End If
						Next
					Next
				End If	
				'Releasing Object of "SaveDataModel" Window
				Set ObjSaveDataModelWnd=Nothing
	   End If
   End If

	If bFlag=True Then
		'Clicking On OK Button to Exit BMIDE
		Call Fn_Button_Click("Fn_BMIDE_ExitBMIDE", JavaWindow("Business Modeler").JavaWindow("SaveDataModel"), "OK")
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Pass :Successfully Exited BMIDE")
		'Function Returns True
		Fn_BMIDE_ExitBMIDE=True
	Else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail : Invalid Project names ["+strProjects+"] pass by user thats why BMIDE Exit Operation Canceled")
		Call Fn_Button_Click("Fn_BMIDE_ExitBMIDE", JavaWindow("Business Modeler").JavaWindow("SaveDataModel"), "Cancel")
	End If
End Function
'-------------------------------------------------------------------Function Used to Identify which Tree is working as Business Tree And which Working as Extension Tree----------------------------------------------
'Function Name		:	Fn_BMIDE_TreeIndexIdentification

'Description			 :	Function Used to Identify which Tree is working as Business Tree And which Working as Extension Tree
'										Because in BMIDE sometimes BusinessObject Tree Highlight to ExtensionTree and wise versa

'Return Value		   : 	True Or False

'Pre-requisite			:	Should be Log In BMIDE

'Examples				: 	'1. Call Fn_BMIDE_TreeIndexIdentification()

'History					 :		
'													Developer Name												Date						Rev. No.						Changes Done																Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				30/11/2010			           1.0																																											  Sunny R
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_BMIDE_TreeIndexIdentification()
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_TreeIndexIdentification"
   Dim strProjectName,BReturn,iCurrBOIndx,iCurrETIndx
   Fn_BMIDE_TreeIndexIdentification=False
   'Taking Project Name From Environment File
'   	strProjectName=Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\BMIDEConfig\BMIDE_Config.xml", "BMIDEProjectName")
''	strProjectName=JavaWindow("Business Modeler").JavaTree("BusinessObjectTree").GetItem(0)
'	Call Fn_BMIDE_BusinessObjectTreeOperations("Expand",strProjectName,"")
'	BReturn=Fn_BMIDE_BusinessObjectTreeOperations("ExistBusinessObject",strProjectName+":BusinessObject","")
'	If BReturn=False Then
'		BReturn=Fn_BMIDE_BusinessObjectTreeOperations("ExistBusinessObject",strProjectName+":POM_object","")
'	End If
'	If BReturn=False Then
'		BReturn=Fn_BMIDE_BusinessObjectTreeOperations("ExistBusinessObject",strProjectName+":extensions","")
'	End If
'	If BReturn=False Then
'		iCurrBOIndx=JavaWindow("Business Modeler").JavaTree("BusinessObjectTree").GetTOProperty("index")
'		iCurrETIndx=JavaWindow("Business Modeler").JavaTree("Extension Tree").GetTOProperty("index")
'
'		JavaWindow("Business Modeler").JavaTree("BusinessObjectTree").SetTOProperty "index",iCurrETIndx
'		JavaWindow("Business Modeler").JavaTree("Extension Tree").SetTOProperty "index",iCurrBOIndx
'	End If
	Fn_BMIDE_TreeIndexIdentification=True
End Function

'--------------------------------------------------Function Used to Start\Stop Remote Services---------------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_ServicesOperations

'Description			 :	Function Used to Start\Stop Remote Services

'Return Value		   : 	True Or False

'Pre-requisite			:	

'Examples				:   'StrServer=Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\EnvVar_Ext.xml", "TcServer")
										'StrUser=Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\BMIDE_Config.xml", "ServerUser")
										'StrUser=Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\BMIDE_Config.xml", "ServerUser")
										'StrServices=Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\BMIDE_Config.xml", "Services")
										'Call Fn_BMIDE_ServicesOperations(StrExeLoc,StrServer,StrUser,StrServices)
										'Call Fn_BMIDE_ServicesOperations("autoadmin:Password123","TeamcenterServerManager_MYDB:IISADMIN")

'History					 :		
'													Developer Name												Date						Rev. No.						Changes Done																Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				01/12/2010			           1.0																																											  Sunny R
'													Sandeep N										   				09/12/2010			           1.0																																											  Sunny R
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_BMIDE_ServicesOperations(strServerName,strUserName,strServiceNames)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_ServicesOperations"
	Const ForReading = 1
	Const ForWriting = 2
	
	Dim arrUser,arrServices,iCounter
	Dim restartBatchFilePath
	Dim cmdRestart
	Dim objFSO, objShell
	Dim currDir, batchName
	Dim cmdParam
	Dim restartLogFile
	Dim retVal, stringText
	Dim MyURL,req,sFileName
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	
	batchName = "Result.txt"
	currDir = Environment.Value("sPath") + "\Utilities\BMIDE\"
	'restartLogFile = "C:\psexec.txt"
	
'	Environment.Value("sPath") = Fn_GetEnvValue("User", "AutomationDir")
'	strAppServerType = Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\BMIDEConfig\BMIDE_Config.xml", "AppServerType") 
'	retVal=False
'
'	If strAppServerType = "IIS" Then
'		restartBatchFilePath = Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\BMIDEConfig\BMIDE_Config.xml", "RestartBatchFilePath") 
'		arrUser=Split(strUserName,":")
'
		If objFSO.FileExists(currDir & batchName) Then
		objFSO.DeleteFile currDir & batchName, True
		End If
'
'		Set objFile = objFSO.CreateTextFile(currDir & batchName , True) 
'		
'		'Adding code to bat file to clean TEMP folder for executing client m\c
'		cmdParam = "cd /d %TEMP%"
'		objFile.WriteLine cmdParam
'		cmdParam =  "for /d %%D in (*) do rd /s /q " + "%%D"
'		objFile.WriteLine cmdParam
'		cmdParam = "del /f /q *"
'		objFile.WriteLine cmdParam
'		
'		cmdParam = "cd /d " + Environment.Value("sPath") + "\Utilities\BMIDE"
'		objFile.WriteLine cmdParam
		'if instr(arrUser(0), "auto") > 0 Then
		'	cmdParam = "PsExec.exe -u " + strServerName + "\" + arrUser(0) + " -p " + arrUser(1) + " \\" + strServerName + " -s /accepteula cmd.exe /c " + restartBatchFilePath + " > " + restartLogFile
		'else
		'	cmdParam = "PsExec.exe -u " + "plm\" + arrUser(0) + " -p " + arrUser(1) + " \\" + strServerName + " -s /accepteula cmd.exe /c " + restartBatchFilePath + " > " + restartLogFile
		'end if
		'objFile.WriteLine cmdParam
		
		'cmdParam = "taskkill /s " + strServerName + " /u plm\" + arrUser(0) + " /p " + arrUser(1) + " /FI ""IMAGENAME eq PSEXESVC.exe"""
		'objFile.WriteLine cmdParam
		
'		cmdParam = "schtasks.exe /Create /S \\" + strServerName + " /TN restart_tcservices /XML " + Environment.Value("sPath") + "\Utilities\BMIDE\restart_tcservices.xml"
'		objFile.WriteLine cmdParam
'		cmdParam = "ping localhost -n 5 > nul"
'		objFile.WriteLine cmdParam
'		cmdParam = "schtasks.exe /Run /S \\" + strServerName + " /TN restart_tcservices > " + restartLogFile
'		objFile.WriteLine cmdParam
'		cmdParam = "ping localhost -n 80 > nul"
'		objFile.WriteLine cmdParam
'		cmdParam = "schtasks.exe /Delete /S \\" + strServerName + " /TN restart_tcservices /F"
'		objFile.WriteLine cmdParam
'		cmdParam = "ping localhost -n 5 > nul"
'		objFile.WriteLine cmdParam
'
'		objFile.Close
'
'		Set objShell = CreateObject("WScript.Shell")
'		objShell.Run "%comspec% /c " + Environment.Value("sPath") +"\Utilities\BMIDE\" + batchName ,2,True 
'		Set objShell = Nothing
'
'        wait(100)
	
MyURL = "http://"&strServerName&":7654/restartServices"
Set req = CreateObject("MSXML2.XMLHTTP.6.0")
req.Open "GET", myURL, False
req.Send
	
Set objLogFile = objFSO.CreateTextFile(currDir & batchName,true)
	objLogFile.WriteLine req.ResponseText
		
		If objFSO.FileExists(currDir & batchName) Then 
			Set objFile = objFSO.OpenTextFile(currDir & batchName, ForReading) 
			Do Until objFile.AtEndOfStream
				stringText = objFile.ReadLine
				If InStr(LCase(stringText), LCase("nRESTART SUCCESS")) > 0 Then
					retVal = True
					Exit Do
				End If
			Loop
		End If

wait(200)

	Set objFile = Nothing
	Set objFSO = Nothing 
	'End If  
	Fn_BMIDE_ServicesOperations = retVal

End Function
'-------------------------------------------------------------------Function Used to Perform Operations On Inner JavaTab Which presents on Main Tab------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_InnerTabOperations

'Description			 :	Function Used to Perform Operations On Inner JavaTab Which presents on Main Tab

'Parameters			   :   '1.strAction:Action to Perform
										'2.strTabName: Tab Name on which have to perform the operations

'Return Value		   : 	True Or False

'Pre-requisite			:	Should be Log In BMIDE

'Examples				: 	Call Fn_BMIDE_InnerTabOperations("Activate","Properties")
'										Call Fn_BMIDE_InnerTabOperations("Activate","Deep Copy Rules")

'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done								Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				02/12/2010			           1.0																								Sunny R
'													Sandeep N										   				18/01/2012			           1.1					Added Case "Deep Copy Rules"		   Sunny R
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_BMIDE_InnerTabOperations(strAction,strTabName)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_InnerTabOperations"
  Dim crrActiveTab
  Select Case strAction
	 	'This Case to Activate Tabs
		 Case "Activate"	'Fn_BMIDE_InnerTabOperations("Activate","Properties")
				Call Fn_UI_JavaTab_Select("Fn_BMIDE_InnerTabOperations",JavaWindow("Business Modeler"),"InnerTab", strTabName)
                Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Successfully Activted the Tab [" + strTabName +"]")
				Fn_BMIDE_InnerTabOperations=True

		Case "VerifyActivate"
				crrActiveTab=Fn_UI_Object_GetROProperty("Fn_BMIDE_InnerTabOperations",JavaWindow("Business Modeler").JavaTab("InnerTab"),"value")
				If Trim(crrActiveTab)=Trim(strTabName) Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Successfully verified Tab [" + strTabName +"] is currently activated")
					Fn_BMIDE_InnerTabOperations=True
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fail to verify Tab [" + strTabName +"] is currently activated")
					Fn_BMIDE_InnerTabOperations=False
				End If
		Case Else
			Fn_BMIDE_InnerTabOperations=False
   End Select
End Function


'-------------------------------------------------------------------Function Used to Add Or Modify Operation Input Property From Create Descriptor Tab------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_OperationInputPropertyOperations

'Description			 :	Function Used to Add Or Modify Operation Input Property From Create Descriptor Tab

'Parameters			   :   '1.strAction:Action to Perform
										'2.strPropName: Property name
										'3.bRequired :Required option
										'4.bVisible :Bisible Option
										'5.strDesc : Prperty Description
										'6.strUsage : Usage Name
										'7strObjectType: Compound Object Type
										'8.strObjectConstant :Compound Object Constants
										'9.strErrMsg: Error Message

'Return Value		   : 	True Or False

'Pre-requisite			:	Should be Log In BMIDE

'Examples				: 	Call Fn_BMIDE_OperationInputPropertyOperations("Add","d2errr","Off","On","test","None","","","")
'										Call Fn_BMIDE_OperationInputPropertyOperations("Edit","d2errr","On","Off","Edited Description test","","","","")
'										In "Verify" Case :- Pass Values which have to verify
'										Call Fn_BMIDE_OperationInputPropertyOperations("Verify","item_id","On","On","item id of the object","None","","","")

'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				02/12/2010			           1.0																				Sunny R
'													Sandeep N										   				04/12/2010			           1.0							Added Case "Verify"			  Sunny R
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_BMIDE_OperationInputPropertyOperations(strAction,strPropName,bRequired,bVisible,strDesc,strUsage,strObjectType,strObjectConstant,strErrMsg)
GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_OperationInputPropertyOperations"
   'Declaring Variables
   Dim intCounter,intRowCount,strPropertyName,ObjDilog,bFlag
   Dim statusOfRequired,statusOfVisible,statusOfDesc,statusOfUsage,statusOfType,statusOfConstant
   'Clicking On CreateDescriptor Tab
	Call Fn_BMIDE_InnerTabOperations("Activate","Operation Descriptor")
	Fn_BMIDE_OperationInputPropertyOperations=False
	bFlag=False
   Select Case strAction
		 	Case "Add" 'Case To Add New OperationInput Property From Business Object
				'Clicking On "AddOperationInputProperty" to Open "New operationInput Property" Dialog
				Call Fn_Button_Click("Fn_BMIDE_OperationInputPropertyOperations", JavaWindow("Business Modeler"), "AddOperationInputProperty")
				'Creting object "New operationInput Property" Dialog
				Set ObjDilog=Fn_UI_ObjectCreate("Fn_BMIDE_OperationInputPropertyOperations", JavaWindow("Business Modeler").JavaWindow("NewOperationInputProperty"))
				Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_BMIDE_OperationInputPropertyOperations",ObjDilog.JavaRadioButton("AddPropertyFromBusinessObject"),"attached text","Add a Property from Business Object")
				'Selecting "Add a Property from Business Object" to Add properties from Business Objects
				Call Fn_UI_JavaRadioButton_SetON("Fn_BMIDE_OperationInputPropertyOperations",ObjDilog, "AddPropertyFromBusinessObject")
				'Checking Existance of Next button
				If Fn_UI_ObjectExist("Fn_BMIDE_OperationInputPropertyOperations",ObjDilog.JavaButton("Next"))=True Then
					'Clicking on Next button
					Call Fn_Button_Click("Fn_BMIDE_OperationInputPropertyOperations", ObjDilog,"Next")
				End If
				If strPropName<>"" Then
					'Setting Property Name
					'Call Fn_UI_EditBox_Type("Fn_BMIDE_OperationInputPropertyOperations",ObjDilog,"PropertyName",strPropName)
					Call Fn_Button_Click("Fn_BMIDE_OperationInputPropertyOperations", ObjDilog,"PropertyNameBrowse")
					Call Fn_Edit_Box("Fn_BMIDE_OperationInputPropertyOperations",JavaWindow("Business Modeler").JavaWindow("AttachmentProperties"),"AttachmentsProperty",strPropName)
                    wait 2
					bFlag=False
					'For intCounter=0 to Cint(JavaWindow("Business Modeler").JavaWindow("AttachmentProperties").JavaTable("PropertiesTable").GetROProperty("rows"))-1
					For intCounter=0 to Cint(Fn_UI_Object_GetROProperty("",JavaWindow("Business Modeler").JavaWindow("AttachmentProperties").JavaTable("PropertiesTable"),"rows"))-1
						If JavaWindow("Business Modeler").JavaWindow("AttachmentProperties").JavaTable("PropertiesTable").GetCellData(intCounter,0)=strPropName Then
							JavaWindow("Business Modeler").JavaWindow("AttachmentProperties").JavaTable("PropertiesTable").SelectCell intCounter,0
							bFlag=True
							Exit for
						End If
					Next
					If bFlag=False Then
						 Set ObjDilog=Nothing
						Exit function
					End If
					Call Fn_Button_Click("Fn_BMIDE_OperationInputPropertyOperations", JavaWindow("Business Modeler").JavaWindow("AttachmentProperties"),"OK")
				End If
				If bRequired<>"" Then
					'Setting Status Of Required Check Box
					Call Fn_CheckBox_Set("Fn_BMIDE_OperationInputPropertyOperations", ObjDilog, "Required", bRequired)
				End If
				If bVisible<>"" Then
					'Setting Status Of Visible Check Box
					Call Fn_CheckBox_Set("Fn_BMIDE_OperationInputPropertyOperations", ObjDilog, "Visible", bVisible)
				End If
				If strDesc<>"" Then
					'Setting Description
					Call Fn_Edit_Box("Fn_BMIDE_OperationInputPropertyOperations",ObjDilog,"Description",strDesc)
				End If
				If strUsage<>"" Then
					Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_BMIDE_OperationInputPropertyOperations",ObjDilog.JavaRadioButton("Usage"),"attached text",strUsage)
					'Selecting usage Type
					Call Fn_UI_JavaRadioButton_SetON("Fn_BMIDE_OperationInputPropertyOperations",ObjDilog, "Usage")
					If strUsage="Type" Then
						'Selecting Type
						Call Fn_Edit_Box("Fn_BMIDE_OperationInputPropertyOperations",ObjDilog,"CompoundObjectType",strObjectType)
					 ElseIf strUsage="Constant" Then
   						'Selecting Constant
						Call Fn_Edit_Box("Fn_BMIDE_OperationInputPropertyOperations",ObjDilog,"CompoundObjectConstant",strObjectConstant)
					End If	
				End If
				Call Fn_Button_Click("Fn_BMIDE_OperationInputPropertyOperations", ObjDilog, "Finish")
				Fn_BMIDE_OperationInputPropertyOperations=True

			Case "Edit"		'Case To Modify OperationInput Property
                intRowCount=Fn_UI_Object_GetROProperty("Fn_BMIDE_OperationInputPropertyOperations",JavaWindow("Business Modeler").JavaTable("OperationInputPropertyTable"), "rows")
				For intCounter=0 To intRowCount-1
					strPropertyName=JavaWindow("Business Modeler").JavaTable("OperationInputPropertyTable").GetCellData(intCounter,"Name")
					If Trim(strPropName)=Trim(strPropertyName) Then
						JavaWindow("Business Modeler").JavaTable("OperationInputPropertyTable").SelectCell intCounter,0
						bFlag=True
						Exit For
					End If
				Next
				If bFlag=False Then
					Exit Function
				End If
				Call Fn_Button_Click("Fn_BMIDE_OperationInputPropertyOperations", JavaWindow("Business Modeler"), "EditOperationInputProperty")
				Set ObjDilog=Fn_UI_ObjectCreate("Fn_BMIDE_OperationInputPropertyOperations", JavaWindow("Business Modeler").JavaWindow("NewOperationInputProperty"))
'				If strPropName<>"" Then
'					'Setting Property Name
'					Call Fn_Edit_Box("Fn_BMIDE_OperationInputPropertyOperations",ObjDilog,"PropertyName",strPropName)
'				End If

                If bRequired<>"" Then
					'Setting Status Of Required Check Box
					Call Fn_CheckBox_Set("Fn_BMIDE_OperationInputPropertyOperations", ObjDilog, "Required", bRequired)
				End If

                If bVisible<>"" Then
					'Setting Status Of Visible Check Box
					Call Fn_CheckBox_Set("Fn_BMIDE_OperationInputPropertyOperations", ObjDilog, "Visible", bVisible)
				End If

				
				If strDesc<>"" Then
					'Setting Description
					Call Fn_Edit_Box("Fn_BMIDE_OperationInputPropertyOperations",ObjDilog,"Description",strDesc)
				End If
				If strUsage<>"" Then
					Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_BMIDE_OperationInputPropertyOperations",ObjDilog.JavaRadioButton("Usage"),"attached text",strUsage)
					'Selecting usage Type
					Call Fn_UI_JavaRadioButton_SetON("Fn_BMIDE_OperationInputPropertyOperations",ObjDilog, "Usage")
					If strUsage="Type" Then
						'Selecting Type
						Call Fn_Edit_Box("Fn_BMIDE_OperationInputPropertyOperations",ObjDilog,"CompoundObjectType",strObjectType)
					 ElseIf strUsage="Constant" Then
   						'Selecting Constant
						Call Fn_Edit_Box("Fn_BMIDE_OperationInputPropertyOperations",ObjDilog,"CompoundObjectConstant",strObjectConstant)
					End If
				End If
				Call Fn_Button_Click("Fn_BMIDE_OperationInputPropertyOperations", ObjDilog, "Finish")
				Fn_BMIDE_OperationInputPropertyOperations=True

		Case "Verify"
				intRowCount=Fn_UI_Object_GetROProperty("Fn_BMIDE_OperationInputPropertyOperations",JavaWindow("Business Modeler").JavaTable("OperationInputPropertyTable"), "rows")
				For intCounter=0 To intRowCount-1
					strPropertyName=JavaWindow("Business Modeler").JavaTable("OperationInputPropertyTable").GetCellData(intCounter,"Name")
					If Trim(strPropName)=Trim(strPropertyName) Then
						JavaWindow("Business Modeler").JavaTable("OperationInputPropertyTable").SelectCell intCounter,0
						bFlag=True
						Exit For
					End If
				Next
				If bFlag=False Then
                    Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail:IInvalid Property Name Passed by User [" + strPropName +"]")
					Exit Function
				End If
				'Clicking On "EditOperationInputProperty" Button to Open "ModifyOperationInputProperty" window
				Call Fn_Button_Click("Fn_BMIDE_OperationInputPropertyOperations", JavaWindow("Business Modeler"), "EditOperationInputProperty")
				'Creating Object Of "NewOperationInputProperty" Window
				Set ObjDilog=Fn_UI_ObjectCreate("Fn_BMIDE_OperationInputPropertyOperations", JavaWindow("Business Modeler").JavaWindow("NewOperationInputProperty"))
				bFlag=True
				If bRequired<>"" Then
					bFlag=False
                    statusOfRequired=Fn_UI_Object_GetROProperty("Fn_BMIDE_OperationInputPropertyOperations",ObjDilog.JavaCheckBox("Required"), "value")
					If statusOfRequired="0" Then
						statusOfRequired="off"
					ElseIf statusOfRequired="1" Then
						statusOfRequired="on"
					End If
					If LCase(bRequired)=LCase(statusOfRequired) Then
						bFlag=True
					Else
						Call Fn_Button_Click("Fn_BMIDE_OperationInputPropertyOperations", ObjDilog, "Cancel")
						Set ObjDilog=Nothing
						Exit Function
					End If
				End If
				If bVisible<>"" Then
					bFlag=False
                    statusOfVisible=Fn_UI_Object_GetROProperty("Fn_BMIDE_OperationInputPropertyOperations",ObjDilog.JavaCheckBox("Visible"), "value")
					If statusOfVisible="0" Then
						statusOfVisible="off"
					ElseIf statusOfVisible="1" Then
						statusOfVisible="on"
					End If
					If LCase(bVisible)=LCase(statusOfVisible) Then
						bFlag=True
					Else
						Call Fn_Button_Click("Fn_BMIDE_OperationInputPropertyOperations", ObjDilog, "Cancel")
						Set ObjDilog=Nothing
						Exit Function
					End If
				End If
				If strDesc<>"" Then
					'Verifying Description
					statusOfDesc=Fn_UI_Object_GetROProperty("Fn_BMIDE_OperationInputPropertyOperations",ObjDilog.JavaEdit("Description"), "value")
					If Trim(statusOfDesc)=Trim(strDesc) Then
						bFlag=True
					Else
						Call Fn_Button_Click("Fn_BMIDE_OperationInputPropertyOperations", ObjDilog, "Cancel")
						Set ObjDilog=Nothing
						Exit Function
					End If
				End If
				If strUsage<>"" Then
					Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_BMIDE_OperationInputPropertyOperations",ObjDilog.JavaRadioButton("Usage"),"attached text",strUsage)
					'Retriving Value of usage Type
					statusOfUsage=Fn_UI_Object_GetROProperty("Fn_BMIDE_OperationInputPropertyOperations",ObjDilog.JavaRadioButton("Usage"), "value")
					If statusOfUsage="1" Then
						bFlag=True
					Else
						Call Fn_Button_Click("Fn_BMIDE_OperationInputPropertyOperations", ObjDilog, "Cancel")
						Set ObjDilog=Nothing
						Exit Function
					End If
				End If
				If strObjectType<>"" Then
					'Verifying Type
					statusOfType=Fn_UI_Object_GetROProperty("Fn_BMIDE_OperationInputPropertyOperations",ObjDilog.JavaEdit("CompoundObjectType"), "value")
					If Trim(statusOfType)=Trim(strObjectType) Then
						bFlag=True
					Else
						Call Fn_Button_Click("Fn_BMIDE_OperationInputPropertyOperations", ObjDilog, "Cancel")
						Set ObjDilog=Nothing
						Exit Function
					End If
				End If
				If strObjectConstant<>"" Then
					'Verifying Constants
					statusOfConstant=Fn_UI_Object_GetROProperty("Fn_BMIDE_OperationInputPropertyOperations",ObjDilog.JavaEdit("CompoundObjectConstant"), "value")
					If Trim(statusOfConstant)=Trim(strObjectConstant) Then
						bFlag=True
					Else
						Call Fn_Button_Click("Fn_BMIDE_OperationInputPropertyOperations", ObjDilog, "Cancel")
						Set ObjDilog=Nothing
						Exit Function
					End If
				End If
				If bFlag=True Then
					Fn_BMIDE_OperationInputPropertyOperations=True
				End If
				'Clicking On Cancel Button To Close Modify OperationInputProperty Dialog
				Call Fn_Button_Click("Fn_BMIDE_OperationInputPropertyOperations", ObjDilog, "Cancel")

   End Select
   Set ObjDilog=Nothing
End Function

'------------------------------------------------------------'Function Used to Add Runtime Properties to the Object-----------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_RuntimePropertiesOperation

'Description			 :	Function Used to Add Runtime Properties to the Object

'Parameters			   :   '1.strAction= Action Name
'										1.strName= String property name
'										2.  strDisplayName= String Display name of property
'										3. strAttributeType = Attribute type of Property 
'										4.,strStringLength = length of of Property 
'										5. strReferenceObject = Reference Business Object
'										6. strDescription= Runtime Property Description
'										7.chkArray =Array Key CheckBox Names "Array?,Unlimited" 'Comma Separated

'Return Value		   : 	True Or False


'Examples				:	 Fn_BMIDE_RuntimePropertiesOperation("Add","Test","Test Diplay Name","String","32","","Test Runtime Property Description","")

'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N												02-Dec-2010								1.0																							Sunny
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_BMIDE_RuntimePropertiesOperation(strAction,strName,strDisplayName,strAttributeType,strStringLength,strReferenceObject,strDescription,chkArray)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_RuntimePropertiesOperation"
	'Variable Declaration
	Dim ObjCustPropWindow,intCounter,strPrototype,arrChkArray
   'Function Returns False
	Fn_BMIDE_RuntimePropertiesOperation=False
	'Clicking On Properties Tab
	Call Fn_BMIDE_InnerTabOperations("Activate","Properties")

	Select Case strAction
		Case "Add", "AddFormProperty"
				'Clicking On Add Button To Add Runtime Properties 
				Call Fn_Button_Click("Fn_BMIDE_RuntimePropertiesOperation", JavaWindow("Business Modeler"), "AddPropeties")
				wait(2)
				'Checking Existance of NewCustomProperty Window
				If Fn_UI_ObjectExist("Fn_BMIDE_RuntimePropertiesOperation", JavaWindow("Business Modeler").JavaWindow("NewCustomProperty"))=False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail: NewCustomProperty Dialog Is Not Exist")
					Exit Function
				End If
				'Creating Object Of NewCustomProperty window
				Set ObjCustPropWindow=Fn_UI_ObjectCreate("Fn_BMIDE_RuntimePropertiesOperation", JavaWindow("Business Modeler").JavaWindow("NewCustomProperty"))
				'Selecting Runtime option
				Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_BMIDE_RuntimePropertiesOperation",ObjCustPropWindow.JavaRadioButton("PropertyType"),"attached text","Runtime")
				Call Fn_UI_JavaRadioButton_SetON("Fn_BMIDE_RuntimePropertiesOperation",ObjCustPropWindow, "PropertyType")
				'Clicking On Next Button
				Call Fn_Button_Click("Fn_BMIDE_RuntimePropertiesOperation", ObjCustPropWindow, "Next")
				If strName<>"" Then
					strPrototype= Fn_Edit_Box_GetValue("Fn_BMIDE_RuntimePropertiesOperation",ObjCustPropWindow,"Name")
					strName=strPrototype+strName
					'Setting name to new custom property
					Call Fn_Edit_Box("Fn_BMIDE_RuntimePropertiesOperation",ObjCustPropWindow ,"Name",strName)
					If strDisplayName="" Then
						strDisplayName=strName
					End If
				End If
				If strDescription<>"" Then
					'Setting Description to new custom property
					Call Fn_Edit_Box("Fn_BMIDE_RuntimePropertiesOperation",ObjCustPropWindow ,"Description",strDescription)
				End If
				If strDisplayName<>"" Then
					'Setting Display Name to new custom property
					Call Fn_Edit_Box("Fn_BMIDE_RuntimePropertiesOperation",ObjCustPropWindow ,"DisplayName",strDisplayName)
				End If
				If strAttributeType<> ""  Then
						'Setting Attribute type  to new custom property
						Call Fn_List_Select("Fn_BMIDE_RuntimePropertiesOperation",ObjCustPropWindow,"AttributeType",strAttributeType)
				End If
				If strStringLength<>"" Then
					'Setting String Length to new custom property
					Call Fn_Edit_Box("Fn_BMIDE_RuntimePropertiesOperation",ObjCustPropWindow ,"StringLength",strStringLength)
				End If
				If strReferenceObject<>"" Then
					'Setting Reference Class
					Call Fn_Edit_Box("Fn_BMIDE_RuntimePropertiesOperation",ObjCustPropWindow ,"ReferenceClass",strReferenceObject)
				End If
	            	If chkArray<>"" Then
					'Selecting Array Keys
					arrChkArray=Split(chkArray,",")
					For intCounter=0 To Ubound(arrChkArray)
						Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_BMIDE_RuntimePropertiesOperation",ObjCustPropWindow.JavaCheckBox("ArrayKeys"),"attached text",arrChkArray(intCounter))
						Call Fn_CheckBox_Set("Fn_BMIDE_RuntimePropertiesOperation", ObjCustPropWindow, "ArrayKeys", "ON")
					Next
				End If

				If strAction <> "AddFormProperty" Then
					If Fn_UI_Object_SetTOProperty_ExistCheck("Fn_BMIDE_RuntimePropertiesOperation",ObjCustPropWindow.JavaCheckBox("DescriptorOption"),"attached text","Show this property during creation of a Business Object.") = True Then
						Call Fn_CheckBox_Set("Fn_BMIDE_RuntimePropertiesOperation", ObjCustPropWindow, "DescriptorOption", "OFF")
					End If
				End If
				' Click on Finnish Button
				Call Fn_Button_Click("Fn_BMIDE_RuntimePropertiesOperation",ObjCustPropWindow,"Finish")
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Pass: Function Completed Successfully")
				Fn_BMIDE_RuntimePropertiesOperation = True
				Set ObjCustPropWindow=Nothing
	End Select
End Function

'-------------------------------------------------------------------Function Used to Perform Operations Bussiness Object Display Rule-------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_DisplayRuleOperations

'Description			 :	Function Used to Perform Operations Bussiness Object Display Rule

'Parameters			   :   '1.strAction:Action to Perform
										'2.strOrganisation:Organization Node Path
										'3.strProject: Project Name For Connection
										'4.strProfileName: Profile Name
										'5:strPassword : Password For Connection
										'6:strGroup: Group name
										'7strRole:User Role
										'8.strCondition: Condition
										'9.bPropagate: Propagate To sun Bussiness Object option

'Return Value		   : 	True Or False

'Pre-requisite			:	Should be Log In BMIDE

'Examples				: 	Call Fn_BMIDE_DisplayRuleOperations("Add","Organization:dba:DBA","Temp","","AutoTestDBA","dba","DBA","isTrue","ON")
'																																		"Organization:Accessor Type"
'										Call Fn_BMIDE_DisplayRuleOperations("Select","AutoGrp1:Group","","","","","","","")
'										Call Fn_BMIDE_DisplayRuleOperations("Remove","AutoGrp1:Group","","","","","","","")
'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				03/12/2010			           1.0																			  Sunny R
'													Sandeep N										   				13/01/2011			           1.0								Case "Select"						 Sunny R
'													Pranav Ingle										   			 06/01/2012 		            1.1								 Case "Verify"							Sandeep
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_BMIDE_DisplayRuleOperations(strAction,strOrganisation,strProject,strProfileName,strPassword,strGroup,strRole,strCondition,bPropagate)
GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_DisplayRuleOperations"
   Dim iRowCount,iCounter,arrRowValue(1),strFullRowVal,arrOrgAss,strExpVal,bFlag
   Dim ObjDispalyRuleDialog,ObjConnectionDialog,ObjSrchOrgDialog
   'Clicking On Display Rule Tab
	Call Fn_BMIDE_InnerTabOperations("Activate","Display Rules")
	'Creating Objects Of Windows
	Set ObjDispalyRuleDialog=JavaWindow("Business Modeler").JavaWindow("DisplayRule")
	Set ObjConnectionDialog= JavaWindow("Business Modeler").JavaWindow("TeamcenterRepositoryConnection")
	Set ObjSrchOrgDialog=JavaWindow("Business Modeler").JavaWindow("SearchOrganization")
	Fn_BMIDE_DisplayRuleOperations=False
	Select Case strAction
			Case "Add" 'Case to Add Display Rule
				'Clicking on Add Display Rule Button
				Call Fn_Button_Click("Fn_BMIDE_DisplayRuleOperations",JavaWindow("Business Modeler"),"AddHideBusinessObjectRules")
				If strOrganisation<>"" Then
						'Selecting Organization
						Call Fn_Button_Click("Fn_BMIDE_DisplayRuleOperations",ObjDispalyRuleDialog,"Browse")
						'Checking Existance Of "TeamcenterRepositoryConnection" Dialog
						If Fn_UI_ObjectExist("Fn_BMIDE_DisplayRuleOperations", ObjConnectionDialog)=True Then						
							If strProject<>"" Then
                                'Selecting Project From Project List
								Call Fn_List_Select("Fn_BMIDE_DisplayRuleOperations", ObjConnectionDialog, "Project",strProject)
							End If
							If Fn_UI_Object_GetROProperty("Fn_BMIDE_DisplayRuleOperations",ObjConnectionDialog.JavaList("ServerProfile"), "enabled")=1 Then
								If strProfileName<>"" Then
									'Selecting Profile From Profile List
									Call Fn_List_Select("Fn_BMIDE_DisplayRuleOperations", ObjConnectionDialog, "ServerProfile",strProfileName)
								End If
							End If
							If  Fn_UI_ObjectExist("Fn_BMIDE_DisplayRuleOperations", ObjConnectionDialog.JavaEdit("Password"))=True Then
								If strPassword<>"" Then
									'Setting password 
									Call Fn_Edit_Box("Fn_BMIDE_DisplayRuleOperations",ObjConnectionDialog,"Password",strPassword)
								End If
							End If
							If  Fn_UI_ObjectExist("Fn_BMIDE_DisplayRuleOperations", ObjConnectionDialog.JavaEdit("Group"))=True Then
								If strGroup<>"" Then
									'Setting Group
									Call Fn_Edit_Box("Fn_BMIDE_DisplayRuleOperations",ObjConnectionDialog,"Group",strGroup)
								End If
							End If
							If  Fn_UI_ObjectExist("Fn_BMIDE_DisplayRuleOperations", ObjConnectionDialog.JavaEdit("Role"))=True Then
								If strRole<>"" Then
									'Setting Role
									Call Fn_Edit_Box("Fn_BMIDE_DisplayRuleOperations",ObjConnectionDialog,"Role",strRole)
								End If
							End If
                            'Clicking "Connect" Button To Connect To The Host
							If Fn_UI_ObjectExist("Fn_BMIDE_DeployProject",ObjConnectionDialog.JavaButton("Connect"))=True Then
								Call Fn_Button_Click("Fn_BMIDE_DeployProject",ObjConnectionDialog, "Connect")
							End If
							ObjConnectionDialog.JavaButton("Finish").WaitProperty "enabled","1",iTime
							'Clicking "Finish" Button
							Call Fn_Button_Click("Fn_BMIDE_DeployProject",ObjConnectionDialog, "Finish")
							End If
							
						'Checking Existance of Search Organisation Dialog
						If  Fn_UI_ObjectExist("Fn_BMIDE_DisplayRuleOperations", ObjSrchOrgDialog)=True Then
							'Selecting Organisation From OrganizationTree 
							Call Fn_JavaTree_Select("Fn_BMIDE_DisplayRuleOperations", ObjSrchOrgDialog, "OrganizationTree", strOrganisation)
                            Call Fn_JavaTree_Node_Activate("Fn_BMIDE_DisplayRuleOperations", ObjSrchOrgDialog,"OrganizationTree", strOrganisation)
							wait 1
							Call Fn_Button_Click("Fn_BMIDE_DeployProject",ObjSrchOrgDialog, "Finish")
						End If				
				End If
				If strCondition<>"" Then
					'Setting Condition
					ObjDispalyRuleDialog.JavaEdit("Condition").Object.setText strCondition
				End If
				If bPropagate<>"" Then
					'Setting Propagate
                    Call Fn_CheckBox_Set("Fn_BMIDE_DisplayRuleOperations", ObjDispalyRuleDialog, "PropagateBusinessObject", bPropagate)
				End If
				Call Fn_Button_Click("Fn_BMIDE_DeployProject",ObjDispalyRuleDialog, "Finish")
				Fn_BMIDE_DisplayRuleOperations=True

		Case "Select","Verify"
			
			bFlag=False
			iRowCount=Fn_UI_Object_GetROProperty("Fn_BMIDE_DisplayRuleOperations",JavaWindow("Business Modeler").JavaTable("HideBusinessObjectRules"),"rows")
			For iCounter=0 To iRowCount-1
				arrRowValue(0)=JavaWindow("Business Modeler").JavaTable("HideBusinessObjectRules").GetCellData(iCounter,"Accessor Name")
				arrRowValue(1)=JavaWindow("Business Modeler").JavaTable("HideBusinessObjectRules").GetCellData(iCounter,"Accessor Type")

				strFullRowVal=arrRowValue(0)+arrRowValue(1)
				arrOrgAss=Split(strOrganisation,":")
				strExpVal=arrOrgAss(0)+arrOrgAss(1)
				If Trim(strFullRowVal)=Trim(strExpVal) Then
					If strAction="Select" Then
						JavaWindow("Business Modeler").JavaTable("HideBusinessObjectRules").ActivateRow iCounter
					End If
					bFlag=True
					Exit For
				End If
			Next
			If bFlag=True Then
				Fn_BMIDE_DisplayRuleOperations=True
			End If

			Case "Remove"
				bFlag=Fn_BMIDE_DisplayRuleOperations("Select",strOrganisation,"","","","","","","")
				If bFlag=True Then
					Call Fn_Button_Click("Fn_BMIDE_DeployProject",JavaWindow("Business Modeler"), "RemoveHideBusinessObjectRules")
					If JavaWindow("Business Modeler").JavaWindow("ConfirmRemoveDisplayRule").Exist(6) Then
						Call Fn_Button_Click("Fn_BMIDE_DeployProject",JavaWindow("Business Modeler").JavaWindow("ConfirmRemoveDisplayRule"), "Yes")
					End If
					Fn_BMIDE_DisplayRuleOperations=True
				End If

	End Select
	'Releasing Objects
	Set ObjDispalyRuleDialog=Nothing
	Set ObjConnectionDialog=Nothing
	Set ObjSrchOrgDialog=Nothing
End Function

'------------------------------------------------------------Function Used to Perform Operations on "Property Constant" table Which appears on Properties Tab---------------------------------------------------------
'Function Name		:	Fn_BMIDE_PropertyConstantsOperations

'Description			 :	Function Used to Perform Operations on "Property Constant" table Which appears on Properties Tab

'Parameters			   :   '1.strAction: Action Name
'										 2.strName : Constant Name
'										 3.strValue : New Value
'										 4.strColumnName: Column Name
'										 5.strExpectedValue: Expected Value

'Return Value		   : 	True Or False


'Examples				:	' "Edit" Case For Modification Of Property Constants Value 
'										strValue Parameter is Varies - For CheckBox Control Its "ON" Or "OFF"
'																								 - For JavaList And EditBox "Value Name"
'										Fn_BMIDE_PropertyConstantsOperations("Edit","Visible","On","","")
'										"Verify" Case to verify values against Constant Name 
'										Fn_BMIDE_PropertyConstantsOperations("Verify","Visible","","Value","false")
'										Fn_BMIDE_PropertyConstantsOperations("Verify","Visible","","Template","foundation")
'										Fn_BMIDE_PropertyConstantsOperations("Select","Visible","","","")
'										VerifyFromODTab :Case to verify values from Operation Descriptor Tab
'										Fn_BMIDE_PropertyConstantsOperations("VerifyFromODTab","Visible","","Value","true")
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N													3-Dec-2010								1.0																						Sunny
'													Sandeep N													18-Feb-2011								1.1																						Sunny
'													Sandeep N													09-Feb-2012								1.2					Added Case "VerifyFromODTab"						Sunny
'													Sandeep N													08-Feb-2013								1.3					Modified case "Edit",Select replace ActivateRow call with SelectCell						Priyanka B
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------	
'strColumnName,strExpectedValue : - This 2 Parameter use for Verify Case
'strColumnName : - Column Name in which have to verify Value eg:- "Value" Or "Template"
'strExpectedValue : - Expected value 
'Parameters:  1.strAction :- Action Name 2.strName :- Constant Name 3.strValue :- New Value 4.strColumnName : - Column Name In Which have to Verify value 5. strExpectedValue :- Expected value																														
Public Function Fn_BMIDE_PropertyConstantsOperations(strAction,strName,strValue,strColumnName,strExpectedValue)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_PropertyConstantsOperations"
   'strValue Parameter is Varies - For CheckBox Control Its "ON" Or "OFF"
   '														 - For JavaList And EditBox "Value Name"
   'Variable Declaration
   Dim intRowCount,bFlag,strType,iCounter,strContantName,strTableData
   Dim ObjPropConstDialog
   Dim iCount,WshShell
   Fn_BMIDE_PropertyConstantsOperations=False
   bFlag=False
   'Activating Properties tab
   Call Fn_BMIDE_InnerTabOperations("Activate","Properties")
   Select Case strAction
				   Case "Edit"
					   'taking Row Count of "PropertyConstantsTable"
 						intRowCount=Fn_UI_Object_GetROProperty("Fn_BMIDE_PropertyConstantsOperations",JavaWindow("Business Modeler").JavaTable("PropertyConstantsTable"),"rows")
						For iCounter=0 To intRowCount-1
								'taking 1 by 1 Constants
								strContantName=JavaWindow("Business Modeler").JavaTable("PropertyConstantsTable").GetCellData(iCounter,"Name")
								'Verifying User Passed Constant Name and Table Constant Name are matched Or Not
								'Added SendKey work around for this table as QTP methods not working from 0914 build
								If Trim(strContantName)=Trim(strName) Then
									'If Trim(strContantName)="Visible" Then
										JavaWindow("Business Modeler").JavaTable("PropertyConstantsTable").SelectCell 0,0
										For iCount=0 To iCounter-1
                                            Set WshShell = CreateObject("WScript.Shell")
											WshShell.SendKeys "{DOWN}"
											wait(1)
										Next
										Set WshShell =Nothing
										bFlag=True
									'Else
										'JavaWindow("Business Modeler").JavaTable("PropertyConstantsTable").ActivateRow iCounter
										'bFlag=True
										Exit For
									'End If
								End If
						Next
						If bFlag=False Then
                            Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail: Invalid Constant Name ["+strName+"] pass by user")
							Exit Function
						End If
						'Clicking on "EditPropertyConstant" Button to open Property Constants Dialog
						Call Fn_Button_Click("Fn_BMIDE_PropertyConstantsOperations", JavaWindow("Business Modeler"), "EditPropertyConstant")
						'Verifying Existance Of "PropertyConstant" Dialog
						If Fn_UI_ObjectExist("Fn_BMIDE_PropertyConstantsOperations",JavaWindow("Business Modeler").JavaWindow("PropertyConstant"))=True Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Pass:Successfully Invoked the PropertyConstant Dialog")
						Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail: Failed to Invoked the PropertyConstant Dialog")
							Exit Function
						End If
						'Creating Object Of "PropertyConstant" Window
						Set ObjPropConstDialog=Fn_UI_ObjectCreate("Fn_BMIDE_PropertyConstantsOperations",JavaWindow("Business Modeler").JavaWindow("PropertyConstant"))
						If strValue<>"" Then
							'Retriving "Type" Of Constant
							strType=Fn_Edit_Box_GetValue("Fn_BMIDE_PropertyConstantsOperations",ObjPropConstDialog,"Type")                          
							'Performing Operation On Existing Control
							'There Are Different Controls Appear for value eg: "CheckBox"  "JavaList"  "EditBox"
							Select Case Trim(strType)
								Case "Boolean" 'Case For CheckBox
										'Call Fn_CheckBox_Set("Fn_BMIDE_PropertyConstantsOperations", ObjPropConstDialog, "Value", strValue)
										Select Case lcase(strValue)
											Case "on"
												Call Fn_UI_JavaRadioButton_SetON("Fn_BMIDE_PropertyConstantsOperations", ObjPropConstDialog, "True")
											Case "off"
												Call Fn_UI_JavaRadioButton_SetON("Fn_BMIDE_PropertyConstantsOperations", ObjPropConstDialog, "False")
										End Select
										
										If strName="Required" Then
											If ObjPropConstDialog.JavaWindow("WhatWouldYouLikeToDo").Exist(5) Then
												Call Fn_Button_Click("Fn_BMIDE_PropertyConstantsOperations", ObjPropConstDialog.JavaWindow("WhatWouldYouLikeToDo"), "OK")
											End If
										End If
								Case "List" 'Case For JavaList
										Call Fn_List_Select("Fn_BMIDE_PropertyConstantsOperations",ObjPropConstDialog,"Value",strValue)
								Case "String" 'Case For Editbox
		                                Call Fn_Edit_Box("Fn_BMIDE_PropertyConstantsOperations",ObjPropConstDialog,"Value",strValue)
							End Select
						End If
						'Clicking On "Finish" Button To Finish Constant Property Modification Operation
						Call Fn_Button_Click("Fn_BMIDE_PropertyConstantsOperations", ObjPropConstDialog, "Finish")
						Fn_BMIDE_PropertyConstantsOperations=True
						Set ObjPropConstDialog=Nothing

					Case "Verify"
						'taking Row Count of "PropertyConstantsTable"
 						intRowCount=Fn_UI_Object_GetROProperty("Fn_BMIDE_PropertyConstantsOperations",JavaWindow("Business Modeler").JavaTable("PropertyConstantsTable"),"rows")
						For iCounter=0 To intRowCount-1
								'taking 1 by 1 Constants
								strContantName=JavaWindow("Business Modeler").JavaTable("PropertyConstantsTable").GetCellData(iCounter,"Name")
								'Verifying User Passed Constant Name and Table Constant Name are matched Or Not
								If Trim(strContantName)=Trim(strName) Then
									strTableData=JavaWindow("Business Modeler").JavaTable("PropertyConstantsTable").GetCellData(iCounter,strColumnName)
									If LCase(Trim(strTableData))=LCase(Trim(strExpectedValue)) Then
										bFlag=True
										Exit For
									End If									
								End If
						Next
						If bFlag=False Then
                            Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail: Invalid Constant Name ["+strName+"] pass by user")
							Exit Function
						Else
							Fn_BMIDE_PropertyConstantsOperations=True
						End If

				 Case "Select"
					   'taking Row Count of "PropertyConstantsTable"
' 						intRowCount=Fn_UI_Object_GetROProperty("Fn_BMIDE_PropertyConstantsOperations",JavaWindow("Business Modeler").JavaTable("PropertyConstantsTable"),"rows")
'						For iCounter=0 To intRowCount-1
'								'taking 1 by 1 Constants
'								strContantName=JavaWindow("Business Modeler").JavaTable("PropertyConstantsTable").GetCellData(iCounter,"Name")
'								'Verifying User Passed Constant Name and Table Constant Name are matched Or Not
'								If Trim(strContantName)=Trim(strName) Then
'									JavaWindow("Business Modeler").JavaTable("PropertyConstantsTable").ActivateRow iCounter
'									bFlag=True
'									Exit For
'								End If
'						Next
						intRowCount=Fn_UI_Object_GetROProperty("Fn_BMIDE_PropertyConstantsOperations",JavaWindow("Business Modeler").JavaTable("PropertyConstantsTable"),"rows")
						For iCounter=0 To intRowCount-1
								'taking 1 by 1 Constants
								strContantName=JavaWindow("Business Modeler").JavaTable("PropertyConstantsTable").GetCellData(iCounter,"Name")
								'Verifying User Passed Constant Name and Table Constant Name are matched Or Not
								'Added SendKey work around for this table as QTP methods not working from 0914 build
								If Trim(strContantName)=Trim(strName) Then
									'If Trim(strContantName)="Visible" Then
										JavaWindow("Business Modeler").JavaTable("PropertyConstantsTable").SelectCell 0,0
										For iCount=0 To iCounter-1
                                            Set WshShell = CreateObject("WScript.Shell")
											WshShell.SendKeys "{DOWN}"
											wait(1)
										Next
										Set WshShell =Nothing
										bFlag=True
									'Else
										'JavaWindow("Business Modeler").JavaTable("PropertyConstantsTable").ActivateRow iCounter
										'bFlag=True
										Exit For
									'End If
								End If
						Next
						If bFlag=False Then
                            Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail: Invalid Constant Name ["+strName+"] pass by user")
							Exit Function
						End If
						Fn_BMIDE_PropertyConstantsOperations=True
			'"VerifyFromODTab" : - To verify property constants from Operationa Descriptor Tab	
			Case "VerifyFromODTab"
				'Activating Operation Descriptor tab
				Call Fn_BMIDE_InnerTabOperations("Activate","Operation Descriptor")
				'taking Row Count of "PropertyConstantsTable"
				intRowCount=Fn_UI_Object_GetROProperty("Fn_BMIDE_PropertyConstantsOperations",JavaWindow("Business Modeler").JavaTable("PropertyConstantsTable"),"rows")
				For iCounter=0 To intRowCount-1
						'taking 1 by 1 Constants
						strContantName=JavaWindow("Business Modeler").JavaTable("PropertyConstantsTable").GetCellData(iCounter,"Name")
						'Verifying User Passed Constant Name and Table Constant Name are matched Or Not
						If Trim(strContantName)=Trim(strName) Then
							strTableData=JavaWindow("Business Modeler").JavaTable("PropertyConstantsTable").GetCellData(iCounter,strColumnName)
							If LCase(Trim(strTableData))=LCase(Trim(strExpectedValue)) Then
								bFlag=True
								Exit For
							End If									
						End If
				Next
				If bFlag=False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail: Invalid Constant Name ["+strName+"] pass by user")
					Exit Function
				Else
					Fn_BMIDE_PropertyConstantsOperations=True
				End If
   End Select
End Function

'------------------------------------------------------------Function Used to Perform Operations on "Naming Rule Attaches" table Which appears on Properties Tab---------------------------------------------------------
'Function Name		:	Fn_BMIDE_NamingRuleAttachesOperations

'Description			 :	Function Used to Perform Operations on "Naming Rule Attaches" table Which appears on Properties Tab

'Parameters			   :   '1.strAction: Action Name
'										 2.strNamingRule : Naming Rule
'										 3.strCase : Case
'										 4.strCondition: Condition
'										 5.bOverride: Override Option
'										6.strColumnName: Column Name
'										7.strExpectedValue: Expected Value

'Return Value		   : 	True Or False


'Examples				:	Fn_BMIDE_NamingRuleAttachesOperations("Add","T2test","Upper","isTrue","On","","")

'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N													4-Dec-2010								1.0																						Sunny
'													Sandeep N													29-Nov-2011								1.1			Added Case "GetCount" & "RemoveAll"																		Sunny
'													Sandeep N													14-Jan-2013								1.2			Modified Case "Select" & "RemoveAll","Remove"																					Sunny
'Replaced JavaWindow("Business Modeler").JavaTable("NamingRuleAttachTable").ActivateRow intCounter with JavaWindow("Business Modeler").JavaTable("NamingRuleAttachTable").SelectCell intCounter, 0
'													Sandeep N													08-Feb-2013								1.3			Added code to select Naming Rule Attaches tab as its design chsange in TC 10.1
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------	
'strColumnName,strExpectedValue : - This 2 Parameter use for Verify Case
'strColumnName : - Column Name in which have to verify Value eg:- "Value" Or "Template"
'strExpectedValue : - Expected value 
Public Function Fn_BMIDE_NamingRuleAttachesOperations(strAction,strNamingRule,strCase,strCondition,bOverride,strColumnName,strExpectedValue)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_NamingRuleAttachesOperations"
   'Variable Declaration
   Dim ObjNamingRuleDialog,Icnt,IRowCount,WshShell
   Fn_BMIDE_NamingRuleAttachesOperations=False
   'Activating Properties tab
   Call Fn_BMIDE_InnerTabOperations("Activate","Properties")
    'Activating Naming Rule Attaches tab
   Call Fn_UI_JavaTab_Select("",JavaWindow("Business Modeler"),"PropertiesInnerTab", "Naming Rule Attaches")
	Select Case strAction
			Case "Add" 'Case To Add New Naming Rule
				'Clicking On "AddNamingRuleAttaches" to Attach Naming Rule
				Call Fn_Button_Click("Fn_BMIDE_NamingRuleAttachesOperations", JavaWindow("Business Modeler"), "AddNamingRuleAttaches")
				'Verifying Existance Of "AttachNamingRule" Dialog
				If Fn_UI_ObjectExist("Fn_BMIDE_NamingRuleAttachesOperations", JavaWindow("Business Modeler").JavaWindow("AttachNamingRule"))=True Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Pass:Successfully Invoked the AttachNamingRule Dialog")
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail: Failed to Invoked the AttachNamingRule Dialog")
					Exit Function
				End If
				'Creating Object Of "AttachNamingRule" Dialog
				Set ObjNamingRuleDialog=Fn_UI_ObjectCreate("Fn_BMIDE_NamingRuleAttachesOperations",JavaWindow("Business Modeler").JavaWindow("AttachNamingRule"))
				If strNamingRule<>"" Then
					'Setting Naming Rule
					Call Fn_Button_Click("Fn_BMIDE_NamingRuleAttachesOperations",ObjNamingRuleDialog, "Browse")                  
					Call Fn_Edit_Box("Fn_BMIDE_NamingRuleAttachesOperations", JavaWindow("Business Modeler").JavaWindow("Find Business Object"),"Criteria",strNamingRule)     'By Nitish B
					IRowCount= Fn_SISW_UI_JavaTable_Operations("Fn_BMIDE_NamingRuleAttachesOperations", "GetRowCount", JavaWindow("Business Modeler").JavaWindow("Find Business Object") , "Table", "", "", "", "", "", "", "")
					For Icnt = 0 To IRowCount-1 
						If Fn_SISW_UI_JavaTable_Operations("Fn_BMIDE_NamingRuleAttachesOperations", "GetCellData", JavaWindow("Business Modeler").JavaWindow("Find Business Object").JavaTable("Table") , "", "GetCellData", "", Icnt, "Name", "", "", "")= strNamingRule Then
							Call Fn_SISW_UI_JavaTable_Operations("Fn_BMIDE_NamingRuleAttachesOperations", "SelectRow", JavaWindow("Business Modeler").JavaWindow("Find Business Object").JavaTable("Table") , "", "", "",Icnt , "Name", "", "", "")
							Call Fn_Button_Click("Fn_BMIDE_NamingRuleAttachesOperations", JavaWindow("Business Modeler").JavaWindow("Find Business Object"), "OK")
							Exit For
						End If
					Next
				End If
				If strCase<>"" Then
					'Selecting Case
					Call Fn_List_Select("Fn_BMIDE_NamingRuleAttachesOperations",ObjNamingRuleDialog,"Case",strCase)
				End If
				If strCondition<>"" Then
					'Setting Condition
					Call Fn_Edit_Box("Fn_BMIDE_NamingRuleAttachesOperations",ObjNamingRuleDialog,"Condition",strCondition)
					wait 1
					Set WshShell = CreateObject("WScript.Shell")
					WshShell.SendKeys "{ESC}"
					wait 1,500
					Set WshShell = nothing
				End If
				If bOverride<>"" Then
					'Setting Status of Override Option
					Call Fn_CheckBox_Set("Fn_BMIDE_NamingRuleAttachesOperations",ObjNamingRuleDialog, "Override", bOverride)
				End If	
				'Clicking on Finish Button To Add Naming Rule
				Call Fn_Button_Click("Fn_BMIDE_NamingRuleAttachesOperations",ObjNamingRuleDialog, "Finish")
				Fn_BMIDE_NamingRuleAttachesOperations=True

			Case "Remove","Select"
				intRowCount=Fn_UI_Object_GetROProperty("Fn_BMIDE_NamingRuleAttachesOperations",JavaWindow("Business Modeler").JavaTable("NamingRuleAttachTable"), "rows")
				For intCounter=0 To intRowCount-1
					strPropName=JavaWindow("Business Modeler").JavaTable("NamingRuleAttachTable").GetCellData(intCounter, "Naming Rule")
					If Trim(strNamingRule)=Trim(strPropName) Then
						JavaWindow("Business Modeler").JavaTable("NamingRuleAttachTable").SelectCell intCounter,0
						Fn_BMIDE_NamingRuleAttachesOperations=True
						Exit For
					End If
				Next 

				If strAction= "Remove" Then
					Call Fn_Button_Click("Fn_BMIDE_NamingRuleAttachesOperations", JavaWindow("Business Modeler"), "RemoveNamingRuleAttches")
				End If

				Fn_BMIDE_NamingRuleAttachesOperations=True
				'Releasing object Of "AttachNamingRule" Dialog
				Set ObjNamingRuleDialog=Nothing

				Case "GetCount"
					Fn_BMIDE_NamingRuleAttachesOperations=Fn_UI_Object_GetROProperty("Fn_BMIDE_NamingRuleAttachesOperations",JavaWindow("Business Modeler").JavaTable("NamingRuleAttachTable"), "rows")

				Case "RemoveAll"
					intRowCount=Fn_UI_Object_GetROProperty("Fn_BMIDE_NamingRuleAttachesOperations",JavaWindow("Business Modeler").JavaTable("NamingRuleAttachTable"), "rows")
					For intCounter=0 To intRowCount-1
                        	JavaWindow("Business Modeler").JavaTable("NamingRuleAttachTable").SelectCell intCounter,0
							Call Fn_Button_Click("Fn_BMIDE_NamingRuleAttachesOperations", JavaWindow("Business Modeler"), "RemoveNamingRuleAttches")
							Fn_BMIDE_NamingRuleAttachesOperations=True
					Next 
	End Select
End Function
'------------------------------------------------------------Function Used to Create New Naming Rule---------------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_CreateNamingRule

'Description			 :	Function Used to Create New Naming Rule

'Parameters			   :   '1.strName: Naming Rulw Name
'										 2.strPattern : Pattern
'										 3.strDescription : Pattern Description
'										 4.bCounters: Generate counter Option
'										 5.strInitialValue: Initial value of Counter
'										6.strMaxValue: Maximum value of Counter

'Return Value		   : 	True Or False


'Examples				:	'Example For Adding Single Pattern
'										Fn_BMIDE_CreateNamingRule("DemoRule","{LOV:Activity Category}[RULE:T2test]","FirstDemoRulePattern","OFF","","")
'										'Example For Adding Multiple Patter . They Are Separeted by Tilda (~)\
'										Fn_BMIDE_CreateNamingRule("DemoRule1","{LOV:Activity Category}[RULE:T2test]~{LOV:Activity Category}[RULE:T2DemoRule]","FirstDemoRulePattern~ScndDemoRulePattern","OFF~OFF","","")

'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N													4-Dec-2010								1.0																						Sunny
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------	
'To Add Multiple Pattern Use Seperator Tilda (~)
Public Function Fn_BMIDE_CreateNamingRule(strName,strPattern,strDescription,bCounters,strInitialValue,strMaxValue)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_CreateNamingRule"
   'Variable Declaration
   Dim strPrefix,arrPattern,iCounter,arrCounters,arrDescription,arrInitialValue,arrMaxValue
   Dim ObjRulesDialog,ObjPatternDialog
   Fn_BMIDE_CreateNamingRule=False
	'Checking Existance of "NamingRules" window
   If Fn_UI_ObjectExist("Fn_BMIDE_CreateNamingRule",JavaWindow("Business Modeler").JavaWindow("NamingRules"))=True Then
	   'Creating Object Of "NamingRules" window
	   Set ObjRulesDialog=Fn_UI_ObjectCreate("Fn_BMIDE_CreateNamingRule",JavaWindow("Business Modeler").JavaWindow("NamingRules"))
	Else
	    Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail: Naming Rule Dialog is Not Exist On Screen")
		Exit Function
   End If
   If strName<>"" Then
		strPrefix=Fn_Edit_Box_GetValue("Fn_BMIDE_CreateNamingRule",ObjRulesDialog,"Name")
		strName=strPrefix+strName
		'Setting Naming Rule Name
        Call Fn_Edit_Box("Fn_BMIDE_CreateNamingRule",ObjRulesDialog,"Name",strName)
   End If
   If strPattern<>"" Then
	   'Creating Object Of "NamingRulePattern" Window
	   Set ObjPatternDialog=JavaWindow("Business Modeler").JavaWindow("NamingRulePattern")
	   arrPattern=Split(strPattern,"~")
		For iCounter=0 To Ubound(arrPattern)
			'Clicking On "Add" Button To Add Naming Rule Pattern
			Call Fn_Button_Click("Fn_BMIDE_CreateNamingRule",ObjRulesDialog, "Add")
			'Setting Naming Rule Pattern
			Call Fn_Edit_Box("Fn_BMIDE_CreateNamingRule",ObjPatternDialog,"Pattern",arrPattern(iCounter))
			If strDescription<>"" Then
				arrDescription=Split(strDescription,"~")
				If arrDescription(iCounter)<>"" Then
					'Setting Naming Rule Pattern Description
					Call Fn_Edit_Box("Fn_BMIDE_CreateNamingRule",ObjPatternDialog,"Description",arrDescription(iCounter))
				End If
			End If
			If bCounters<>"" Then
				arrCounters=Split(bCounters,"~")
				If UCase(arrCounters(iCounter))="ON" Then
					'Setting Stautus Of Generate Counter Option
					Call Fn_CheckBox_Set("Fn_BMIDE_CreateNamingRule",ObjPatternDialog,"GenerateCounters",arrCounters(iCounter))
					If strInitialValue<>"" Then
						arrInitialValue=Split(strInitialValue,"~")
						If arrInitialValue(iCounter)<>"" Then
							'Setting Initial Value
							Call Fn_Edit_Box("Fn_BMIDE_CreateNamingRule",ObjPatternDialog,"InitialValue",arrInitialValue(iCounter))
						End If
					End If
					If strMaxValue<>"" Then
						arrMaxValue=Split(strMaxValue,"~")
						If arrMaxValue(iCounter)<>"" Then
							'Setting Maximum Value
							Call Fn_Edit_Box("Fn_BMIDE_CreateNamingRule",ObjPatternDialog,"MaximumValue",arrMaxValue(iCounter))
						End If
					End If
				End If
			End If
		'Clicking On Finish Button
		Call Fn_Button_Click("Fn_BMIDE_CreateNamingRule", ObjPatternDialog, "Finish")
		Next
		'Releasing Object Of "NamingRulePattern" Window
		Set ObjPatternDialog=Nothing
   End If
   'Clicking Finish Button To Crete New Naming Rule
   Call Fn_Button_Click("Fn_BMIDE_CreateNamingRule",ObjRulesDialog, "Finish")
	Fn_BMIDE_CreateNamingRule=True
   'Releasing Object Of "NamingRule" Window
   Set ObjRulesDialog=Nothing
End Function

'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_PrefixErrorMsgVerify

'Description			 :	Function Used to verify Error message typing Invalid prefix /  to verify predefined prefix 

'Parameters			   :	1. strAction - Action to be performed 
'										 2. strLabel - label of window
'										3. strErrorLabel - label of Error Description 
'										4. strName - name to Enter in name Field
'											Note  - For lov  strName should be lov name and Type of LOV seperated by " : " 
'										5. strErrorMsg - Error message to verified	
'										6. btnFinish - pass va;ue if have to click on Finish Button otherwise ""
'Return Value		   : 	True Or False

'Example				: Fn_BMIDE_PrefixErrorMsgVerify("Backspace","New Business Object","Business Object","",  " Invalid ""Name:"" field. It must begin with the prefix ""D3"".","Cancel")
									'Fn_BMIDE_PrefixErrorMsgVerify("InvalidPrefix","New Business Object","Business Object", "B3",  " Invalid ""Name:"" field. It must begin with the prefix ""D3"".","Cancel")
									'Fn_BMIDE_PrefixErrorMsgVerify("Backspace","Rename","Rename Object","",  " Invalid ""New Name:"" field. It must begin with the prefix ""D3"".","")
									'Fn_BMIDE_PrefixErrorMsgVerify("VerifyPrefix","Rename","Rename Object", "D3", "","")
'									Fn_BMIDE_PrefixErrorMsgVerify("Rename","Rename","Rename Object", "D3_Item123", "","Cancel")				
'									Fn_BMIDE_PrefixErrorMsgVerify("VerifyPrefix","New LOV...","LOV","D3:ListOfValuesDate","","")
'									Fn_BMIDE_PrefixErrorMsgVerify("Backspace","New LOV...","LOV"," :ListOfValuesDate",  " Invalid ""Name:"" field. It must begin with the prefix ""D3"".","")
'									 Fn_BMIDE_PrefixErrorMsgVerify("InvalidPrefix","New LOV...","LOV","BB3:ListOfValuesDate",  " Invalid ""Name:"" field. It must begin with the prefix ""D3"".","")
'									Fn_BMIDE_PrefixErrorMsgVerify("InvalidPrefix_Ext","New LOV...","LOV","Id:A3Test",  " Invalid ""Name:"" field. It must begin with the prefix ""D3"".","")
'									Fn_BMIDE_PrefixErrorMsgVerify("Backspace_Ext","New LOV...","LOV"," Id",  " Invalid ""Name:"" field. It must begin with the prefix ""D3"".","")
								
'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Pranav Ingle										   			 06/12/2010			           1.0									Created									Sunny R
'													Sandeep			 									   			 24/08/2011			           1.1									Trim Error Msgs						Sunny R
'													Sandeep			 									   			 12/09/2012			           1.2									added case : InvalidPrefix_Ext						Avinash J
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function Fn_BMIDE_PrefixErrorMsgVerify(strAction,strLabel,strErrorLabel,strName,strErrorMsg,btnClick)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_PrefixErrorMsgVerify"
   Dim ObjDialog, objDesc, WshShell, ErrMsg,objType,arrName
   GBL_EXPECTED_MESSAGE=strErrorMsg
	Set WshShell = CreateObject("WScript.Shell")
	Fn_BMIDE_PrefixErrorMsgVerify = False
	' Set Object of BMIDE Window  new Business Object
	Set ObjDialog = JavaWindow("Business Modeler").JavaWindow("NewBusinessObject")
	' Set Title of window to given strLabel
	ObjDialog.SetTOProperty "title",strLabel
	If strLabel = "Rename"  Then
			ObjDialog.JavaEdit("Name").SetTOProperty "attached text","New Name:" 
			If  Fn_UI_ObjectExist("",ObjDialog.JavaButton("Next")) = true Then
					if ObjDialog.JavaButton("Next").GetROProperty("enabled") = 1 Then  
							Call Fn_Button_Click("Fn_BMIDE_PrefixErrorMsgVerify",ObjDialog,"Next")
					End if
			End If
	End If
	' set object of  Error Description 
	Set objDesc = ObjDialog.JavaEdit("BusinessObject")
	' Set  Attached text  of  javaedit  to given strErrorLabel
	objDesc.SetTOProperty "attached text",strErrorLabel
	Select Case strAction
			Case "Backspace"
					arrName = Split(strName,":")
					If UBound(arrName)= 1 Then
								Set objType = ObjDialog.JavaEdit("Parent")
								' Set  Attached text  of  javaedit  to given strErrorLabel
								objType.SetTOProperty "attached text","Type:"
								objType.Set arrName(1)
					End If
					ObjDialog.JavaEdit("Name").Activate
					WshShell.SendKeys"{BACKSPACE}"
					WshShell.SendKeys"{BACKSPACE}"
'					ObjDialog.JavaEdit("Name").Set ""
					ErrMsg = Fn_Edit_Box_GetValue("",ObjDialog,"BusinessObject") 						
					If  Trim(strErrorMsg) = Trim(ErrMsg) Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Pass : Sucessfully Verified Error messge")
							Fn_BMIDE_PrefixErrorMsgVerify = True
					Else
							GBL_ACTUAL_MESSAGE=ErrMsg
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail : Failed to Verify Error messge")
							Exit Function
					End If
			Case "InvalidPrefix"
					arrName = Split(strName,":")
					If UBound(arrName)= 1 Then
								Set objType = ObjDialog.JavaEdit("Parent")
								' Set  Attached text  of  javaedit  to given strErrorLabel
								objType.SetTOProperty "attached text","Type:"
								objType.Set arrName(1)
					End If
					Call Fn_Edit_Box("Fn_BMIDE_PrefixErrorMsgVerify",ObjDialog,"Name", arrName(0))
					ErrMsg = Fn_Edit_Box_GetValue("",ObjDialog,"BusinessObject")	
					If  Trim(strErrorMsg) = Trim(ErrMsg) Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Pass : Sucessfully Verified Error messge")
							Fn_BMIDE_PrefixErrorMsgVerify = True
					Else
							GBL_ACTUAL_MESSAGE=ErrMsg
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail : Failed to Verify Error messge")
							Exit Function
					End If
			Case "Rename"
					Call Fn_Edit_Box("Fn_BMIDE_PrefixErrorMsgVerify",ObjDialog,"Name", strName)
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Pass : Sucessfully Enter Name to Object")
					If ErrMsg <> "" Then
							ErrMsg = Fn_Edit_Box_GetValue("",ObjDialog,"BusinessObject")
							If   Trim(strErrorMsg) <>  Trim(ErrMsg) Then
									GBL_ACTUAL_MESSAGE=ErrMsg
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail : Failed Verify  Message")
									Exit Function
							End If
					End If
					Fn_BMIDE_PrefixErrorMsgVerify = True
			Case "VerifyPrefix"
					arrName = Split(strName,":")
					If UBound(arrName)= 1 Then
								Set objType = ObjDialog.JavaEdit("Parent")
								' Set  Attached text  of  javaedit  to given strErrorLabel
								objType.SetTOProperty "attached text","Type:"
								objType.Set arrName(1)
					End If
					If  arrName(0) =  Fn_Edit_Box_GetValue("",ObjDialog,"Name")	Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Pass : Sucessfully Verified prefix")
							Fn_BMIDE_PrefixErrorMsgVerify = True
					Else	
							GBL_ACTUAL_MESSAGE=Fn_Edit_Box_GetValue("",ObjDialog,"Name")
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail : Failed to Verify prefix")
							Exit Function
					End If
			Case "InvalidPrefix_Ext"
					arrName = Split(strName,":")
					If UBound(arrName)= 1 Then
								Set objType = ObjDialog.JavaEdit("Parent")
								' Set  Attached text  of  javaedit  to given strErrorLabel
								objType.SetTOProperty "attached text",arrName(0)+":"
								objType.Set arrName(1)
					else
						objType.Set arrName(0)
					End If
                    ErrMsg = Fn_Edit_Box_GetValue("",ObjDialog,"BusinessObject")	
					If  Trim(strErrorMsg) = Trim(ErrMsg) Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Pass : Sucessfully Verified Error messge")
							Fn_BMIDE_PrefixErrorMsgVerify = True
					Else
							GBL_ACTUAL_MESSAGE=ErrMsg
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail : Failed to Verify Error messge")
							Exit Function
					End If
			Case "Backspace_Ext"
				Set objType = ObjDialog.JavaEdit("Parent")
				objType.SetTOProperty "attached text",strName+":"
				objType.Activate
				WshShell.SendKeys"{BACKSPACE}"
				WshShell.SendKeys"{BACKSPACE}"
				ErrMsg = Fn_Edit_Box_GetValue("",ObjDialog,"BusinessObject") 						
				If  Trim(strErrorMsg) = Trim(ErrMsg) Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Pass : Sucessfully Verified Error messge")
						Fn_BMIDE_PrefixErrorMsgVerify = True
				Else
						GBL_ACTUAL_MESSAGE=ErrMsg
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail : Failed to Verify Error messge")
						Exit Function
				End If
	End Select

	If btnClick <> "" Then
			Call Fn_Button_Click("Fn_BMIDE_PrefixErrorMsgVerify",ObjDialog,btnClick)
	End If
End Function

'-------------------------------------------------------------------Function Used Add Dataset References------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_DatasetToolActionOperations

'Description			 :	Function Used Perform Operations On Dataset Tool Action

'Parameters			   :	1.strReference: Reference
										'2.strFileType: File Type
										'3.strFormat: File Format

'Return Value		   : 	True Or False

'Pre-requisite			:	"Add Dataset References" Should Open

'Examples				:	Call Fn_BMIDE_AddDatasetReference(TestRef1,TestReference1,BINARY)

'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done								Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				06/12/2010			           1.0																								Sunny R
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_BMIDE_AddDatasetReference(strReference,strFileType,strFormat)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_AddDatasetReference"
   'Variable Declaration
	Dim ObjDataSetRef
	Dim strPrifix
	Fn_BMIDE_AddDatasetReference=False
	'Creating Object Of "AddDatasetReference" Window
   Set ObjDataSetRef=Fn_UI_ObjectCreate("Fn_BMIDE_AddDatasetReference",JavaWindow("Business Modeler").JavaWindow("AddDatasetReference"))
	If  strReference<>"" Then
		strPrifix=Fn_Edit_Box_GetValue("Fn_BMIDE_CreateDataset",ObjDataSetRef,"Reference")
		strReference=strPrifix+strReference
		'Setting Dataset Reference
        Call Fn_Edit_Box("Fn_BMIDE_AddDatasetReference",ObjDataSetRef,"Reference",strReference)
	End If
	If  strFileType<>"" Then
		'Setting File Type
		Call Fn_Edit_Box("Fn_BMIDE_AddDatasetReference",ObjDataSetRef,"FileType",strFileType)
	End If
	If  strFormat<>"" Then
		'Selecting Dataset Format
		Call Fn_List_Select("Fn_BMIDE_AddDatasetReference", ObjDataSetRef, "Format",strFormat)
	End If
	'Clicking On "Finish" button To Add Dataset Reference
	Call Fn_Button_Click("Fn_BMIDE_AddDatasetReference", ObjDataSetRef, "Finish")
	Fn_BMIDE_AddDatasetReference=True
	'Releasing Object Of "AddDatasetReference" Window
	Set ObjDataSetRef=Nothing
End Function
'-------------------------------------------------------------------Function Used Perform Operations On Dataset Tool Action----------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_DatasetToolActionOperations

'Description			 :	Function Used Perform Operations On Dataset Tool Action

'Parameters			   :	1.strAction: Action Name
										'2.strTools: Tool Name
										'3.strOperation:Operation type
										'4.strReferences: References
										'5.strParameter:Parameters


'Return Value		   : 	True Or False

'Pre-requisite			:	"Dataset Tool Action" Dialog Shoul Open

'Examples				:	Call Fn_BMIDE_DatasetToolActionOperations("Image Editor","Open","T2MyRef;On!T2MyRef;Off!T2MyRef2;On","$ACTION!$GROUP!$USER")

'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done								Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				06/12/2010			           1.0																								Sunny R
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'strReferences="Reference Name;Explore Option ! Reference Name;Explore Option"
'strReferences= Main Set Separeted by ( ! ) and Subset Separeted by ( ; )
'strParameter="Parameter ! Parameter"
'strParameter= Main Set Separeted by ( ! )
Public Function Fn_BMIDE_DatasetToolActionOperations(strAction,strTools,strOperation,strReferences,strParameter)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_DatasetToolActionOperations"
   'Variable Declaration
   Dim arrRef,arrRefValue,iCounter,iCounter1,arrParameter
   Dim ObjToolActionDialog
   Fn_BMIDE_DatasetToolActionOperations=False
	'Creating Object Of "DatasetToolAction" Window
	Set ObjToolActionDialog=Fn_UI_ObjectCreate("Fn_BMIDE_DatasetToolActionOperations",JavaWindow("Business Modeler").JavaWindow("DatasetToolAction"))
	Select Case strAction
	Case "Add"

		If strTools<>"" Then
			'Setting Tool
	'		Call Fn_Edit_Box("Fn_BMIDE_DatasetToolActionOperations",ObjToolActionDialog,"Tools",strTools)
			Call Fn_UI_EditBox_Type("Fn_BMIDE_DatasetToolActionOperations",ObjToolActionDialog,"Tools",strTools)
		End If
		If strOperation<>"" Then
			'Setting Operation
			Call Fn_List_Select("Fn_BMIDE_DatasetToolActionOperations",ObjToolActionDialog, "Operations",strOperation)
		End If
		If strReferences<>"" Then
			'Adding References
			arrRef=Split(strReferences,"!")
			For iCounter=0 To Ubound(arrRef)
				If arrRef(iCounter)<>"" Then
						Call Fn_Button_Click("Fn_BMIDE_DatasetToolActionOperations", ObjToolActionDialog, "AddReference")
						arrRefValue=Split(arrRef(iCounter),";")
						If arrRefValue(0)<>"" Then
							Call Fn_List_Select("Fn_BMIDE_DatasetToolActionOperations",JavaWindow("Business Modeler").JavaWindow("AddReference"),"ReferenceName",arrRefValue(0))
						End If
						If arrRefValue(1)<>"" Then
							Call Fn_CheckBox_Set("Fn_BMIDE_DatasetToolActionOperations", JavaWindow("Business Modeler").JavaWindow("AddReference"), "Export", arrRefValue(1))
						End If
						Call Fn_Button_Click("Fn_BMIDE_DatasetToolActionOperations", JavaWindow("Business Modeler").JavaWindow("AddReference"), "Finish")
				End If
			Next
		End If
		If strParameter<>"" Then
			'Adding Parameters
			arrParameter=Split(strParameter,"!")
			For iCounter1=0 To Ubound(arrParameter)
					If arrParameter(iCounter1)<>"" Then
						Call Fn_Button_Click("Fn_BMIDE_DatasetToolActionOperations", ObjToolActionDialog, "AddParameters")
						Call Fn_List_Select("Fn_BMIDE_DatasetToolActionOperations", JavaWindow("Business Modeler").JavaWindow("AddDatasetParameter"),"Parameter",arrParameter(iCounter1))
						Call Fn_Button_Click("Fn_BMIDE_DatasetToolActionOperations",JavaWindow("Business Modeler").JavaWindow("AddDatasetParameter"), "Finish")
					End If
			Next
		End If
		Call Fn_Button_Click("Fn_BMIDE_DatasetToolActionOperations", ObjToolActionDialog, "Finish")
		Fn_BMIDE_DatasetToolActionOperations=True

	End Select
	Set ObjToolActionDialog=Nothing
End Function

'-------------------------------------------------------------------Function Used to Create New Dataset---------------------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_CreateDataset

'Description			 :	Function Used to Create New Dataset

'Parameters			   :	1.sName: New Dataset Name
										'2.sDispName: Display Name Of Dataset
										'3.sParent:Parent type Of Object
										'4.sDesc: Dataset Description
										'5.bAdvance:Advance Option
										'6.bPrimaryObj:Storage Class Type
										'7.sToolsForEdit:Tools For Edit
										'8.sToolsForView:Tools For View
										'9.sReferences: References
										'10.sToolAction: Tools for Action

'Return Value		   : 	True Or False

'Pre-requisite			:	New Dataset Creation Dialog Should be Appear on Screen

'Examples				: strEditTools="IExplore:MSExcel:MSWord"
									'strViewTool="IExplore:MSExcel:MSWord"
									'strReferences="TestRef1:TestReference1:BINARY~TestRef2:TestReference2:TEXT"
									'strActionTool="MSExcel:Open:T2TestRef1;ON!T2TestRef2;OFF:$T2TestRef1!$TestRef2"
									'Call Fn_BMIDE_CreateDataset("DemoDetaset","","Dataset","Demo Dataset For Test","","",strEditTools,strViewTool,strReferences,strActionTool)

'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done								Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				06/12/2010			           1.0																								Sunny R
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'sToolsForEdit :- If want to pass multile values . Separate Them by ( : )
'Eg:-"MSExcel : MSWord : Image Editor"
'sToolsForView:-If want to pass multile values . Separate Them by ( : )
'Eg:- "MSExcel:MSWord:Image Editor"
'sReferences:- Separate Set by (~) and its Values by ( : )
'"Reference : FileType  :Format ~ Reference : FileType : Format ~ Reference : FileType : Format"
'Eg:-"T2TestRef1:TestRef:BINARY~T2TestRef2:TestRef2:TEXT~T2TestRef3:TestRef3:BINARY"
'sToolAction :- Separate Set by (~) and its Values by ( : ) 
'"Tools : Operations : ReferenceName ; Explore Opt ! ReferenceName ;Explore Opt : Parameter ! Parameter ~ Tools : Operations : ReferenceName ; Explore Opt ! ReferenceName ;Explore Opt : Parameter ! Parameter"
'Eg:-"Image Editor:Open:T2TestRef2;On!T2TestRef1;OFF:$T2TestRef2!$T2TestRef1"

Public Function Fn_BMIDE_CreateDataset(sName,sDispName,sParent,sDesc,bAdvance,bPrimaryObj,sToolsForEdit,sToolsForView,sReferences,sToolAction)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_CreateDataset"
   'variable Declaration
	Dim ObjDatasetDialog
	Dim strPrefix,arrToolsForEdit,iCounter,arrToolsForView,arrRefSet,arrRefValue,arrActionSet,arrActionValue,iObjectCount,iCount,bFlag,strObjectName
	Fn_BMIDE_CreateDataset=False
	If Fn_UI_ObjectExist("Fn_BMIDE_CreateDataset", JavaWindow("Business Modeler").JavaWindow("NewDataset"))=True Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Pass: New Dataset Dialog Is Exist On Screen")
		Set ObjDatasetDialog=Fn_UI_ObjectCreate("Fn_BMIDE_CreateDataset",JavaWindow("Business Modeler").JavaWindow("NewDataset"))
	Else
		Exit Function
	    Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: New Dataset Dialog Is Not Exist On Screen")
	End If
	'Taking Project Prifix
	strPrefix=Fn_Edit_Box_GetValue("Fn_BMIDE_CreateDataset",ObjDatasetDialog,"Name")
	sName=strPrefix+sName
	'Setting Name To New Dataset
	Call Fn_Edit_Box("Fn_BMIDE_CreateDataset",ObjDatasetDialog,"Name",sName)
	If sDispName<>"" Then
		'Setting Display Name To New Dataset
		Call Fn_Edit_Box("Fn_BMIDE_CreateDataset",ObjDatasetDialog,"DisplayName",sDispName)
	End If
	If sParent<>"" Then
		'Setting Parent
		Call Fn_Edit_Box("Fn_BMIDE_CreateDataset",ObjDatasetDialog,"Parent",sParent)
	End If
	If sDesc<>"" Then
		'Setting Description For Dataset
		Call Fn_Edit_Box("Fn_BMIDE_CreateDataset",ObjDatasetDialog,"Description",sDesc)
	End If
	If bAdvance<>"" Then
		'Setting Status Of Advance Option
		'Call Fn_CheckBox_Set("Fn_BMIDE_CreateDataset",ObjDatasetDialog, "Advanced", bAdvance)
		If bPrimaryObj<>"" Then
			'Setting Status Of  "Create Primary Business Object" Option
			Call Fn_CheckBox_Set("Fn_BMIDE_CreateDataset",ObjDatasetDialog, "CreatePrimaryBusinessObj", bPrimaryObj)
		End If
	End If
	'Adding Tools For Edit
	If sToolsForEdit<>"" Then
		arrToolsForEdit=Split(sToolsForEdit,":")
		For iCounter=0 To Ubound(arrToolsForEdit)
				Call Fn_Button_Click("Fn_BMIDE_CreateDataset",ObjDatasetDialog,"AddToolsForEdit")
				'Setting Tool Object
				Call Fn_Edit_Box("Fn_BMIDE_CreateDataset",JavaWindow("Business Modeler").JavaWindow("Find Business Object"),"Project",arrToolsForEdit(iCounter))
				Call Fn_Button_Click("Fn_BMIDE_CreateDataset",JavaWindow("Business Modeler").JavaWindow("Find Business Object"),"OK")
		Next
	End If
	iCounter=0
	'Adding Tools For View
	If sToolsForView<>"" Then
		arrToolsForView=Split(sToolsForView,":")
		For iCounter=0 To Ubound(arrToolsForView)
				Call Fn_Button_Click("Fn_BMIDE_CreateDataset",ObjDatasetDialog,"AddToolsForView")
'				bFlag=False
'				'Setting Tool Object
'				iObjectCount=Fn_UI_Object_GetROProperty("Fn_BMIDE_CreateDataset",JavaWindow("Business Modeler").JavaWindow("Find Business Object").JavaTable("Table"), "rows")
'				For iCount=0 To iObjectCount-1
'					strObjectName=JavaWindow("Business Modeler").JavaWindow("Find Business Object").JavaTable("Table").GetCellData(iCount,0)
'					If Trim(arrToolsForView(iCounter))=Trim(strObjectName) Then
'						JavaWindow("Business Modeler").JavaWindow("Find Business Object").JavaTable("Table").SelectCell iCount,0
'						bFlag=True
'						Exit For
'					End If
'				Next
'				If bFlag=False Then
'					Call Fn_Button_Click("Fn_BMIDE_CreateDataset",JavaWindow("Business Modeler").JavaWindow("Find Business Object"),"Cancel")
'					Call Fn_Button_Click("Fn_BMIDE_CreateDataset",ObjDatasetDialog,"Cancel")
'					Set ObjDatasetDialog=Nothing
'					Exit Function
'				End If
				Call Fn_Edit_Box("Fn_BMIDE_CreateDataset",JavaWindow("Business Modeler").JavaWindow("Find Business Object"),"Project",arrToolsForView(iCounter))
				Call Fn_Button_Click("Fn_BMIDE_CreateDataset",JavaWindow("Business Modeler").JavaWindow("Find Business Object"),"OK")
		Next
	End If
	Call Fn_Button_Click("Fn_BMIDE_CreateDataset",ObjDatasetDialog,"Next")
	iCounter=0
	'Adding Dataset References
	If sReferences<>"" Then
		arrRefSet=Split(sReferences,"~")
		For iCounter=0 To Ubound(arrRefSet)
			arrRefValue=Split(arrRefSet(iCounter),":")
			Call Fn_Button_Click("Fn_BMIDE_CreateDataset",ObjDatasetDialog,"AddReferences")
			Call Fn_BMIDE_AddDatasetReference(arrRefValue(0),arrRefValue(1),arrRefValue(2))
		Next
	End If
	Call Fn_Button_Click("Fn_BMIDE_CreateDataset",ObjDatasetDialog,"Next")
	iCounter=0
	'Adding Dataset Tool Action
	If sToolAction<>"" Then
		arrActionSet=Split(sToolAction,"~")
		For iCounter=0 To Ubound(arrActionSet)
			arrActionValue=Split(arrActionSet(iCounter),":")
			Call Fn_Button_Click("Fn_BMIDE_CreateDataset",ObjDatasetDialog,"AddToolAction")
			Call Fn_BMIDE_DatasetToolActionOperations("Add",arrActionValue(0),arrActionValue(1),arrActionValue(2),arrActionValue(3))
		Next
	End If
	Call Fn_Button_Click("Fn_BMIDE_CreateDataset",ObjDatasetDialog,"Finish")
	Fn_BMIDE_CreateDataset=True
	Set ObjDatasetDialog=Nothing
End Function

'-------------------------------------------------------------------Function Used to Perform Operations On Deep Copy Rules Table----------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_DeepCopyRuleOperations

'Description			 :	Function Used to Perform Operations On Deep Copy Rules Table

'Parameters			   :   '1.strAction:Action to Perform
										'2.strOperationType: Operation Type
										'3.strRelationType: Relation Type
										'4.strObjectType: Object Type
										'5.strCondition: Condition
										'6.strActionType: Action Type
										'7.bTargetPri: Target Primary Option
										'8.bCopyPropOnRel: Copy Properties On Relation Option
										'9.bRequired: Required Option
										'10.bSecured: Secured Option

'Return Value		   : 	True Or False

'Pre-requisite			:	Should be Log In BMIDE

'Examples				: 	Call Fn_BMIDE_DeepCopyRuleOperations("Add","SaveAs","3DMarkup","AbsOccDataQualifier","isTrue","CopyAsObject","On","On","Off","On")
'										In "Verify" Case Pass All Expected Values which have to Verify
'										Call Fn_BMIDE_DeepCopyRuleOperations("Verify","SaveAs","3DMarkup","AbsOccDataQualifier","isTrue","CopyAsObject","On","On","Off","On")
'										' In Case "Select" Pass strOperationType,strRelationType,strObjectType,strCondition,strActionType parameters Compulsary
'										Call Fn_BMIDE_DeepCopyRuleOperations("Select","Revise","TC_Attaches","Match All","isTrue","CopyAsObject","","","","")
'										For Remove Case PreRequisite is Row Must Selected 
'										Call Fn_BMIDE_DeepCopyRuleOperations("Remove","","","","","","","","","")
'
'										'For Case "TemplateExistanace" use bTargetPri Parameter to Pass Template Name
'										Fn_BMIDE_DeepCopyRuleOperations("TemplateExistanace","SaveAs","IMAN_Reference","Match All","isTrue","NoCopy","regliveupdatebmidetemplate","","","")
'
'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				07/12/2010			           1.0																				Sunny R
'													Sandeep N										   				08/12/2010			           1.0							Added Case "Select"			Sunny R
'													Priyanka B										   				19/01/2011			           1.0							Added Case "Edit"			Sandeep N
'													Sandeep N										   			  20/01/2011			           1.0							Added Case "Remove"			Sandeep N
'													Sandeep N										   			  30/11/2011			           1.0							Added Case "TemplateExistanace"			Sandeep N
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_BMIDE_DeepCopyRuleOperations(strAction,strOperationType,strRelationType,strObjectType,strCondition,strActionType,bTargetPri,bCopyPropOnRel,bRequired,bSecured)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_DeepCopyRuleOperations"
	'Variable Declaration
	Dim ObjDeepCopyRule
	Dim iRowCount,bFlag,currOpType,currRelType,currObjType,currCondition,currAction,currTarget,currCopyOnRel,currReq,currSecured
	Dim iCounter,arrRowValue(4),strFullRowVal,strExpVal
	Dim relationColumn,iColCount,iCount
	bFlag=False
	'Function Returns False
	Fn_BMIDE_DeepCopyRuleOperations=False
	'Activating "Deep Copy Rules Tab"
	Call Fn_BMIDE_InnerTabOperations("Activate","Deep Copy Rules")
	Select Case strAction
		'"Add" Case to Add New Deep Copy Rule
		Case "Add"
			'Clicking On "AddDeepCopyRule" Button to Invoke "DeepCopyRule" Window
			Call Fn_Button_Click("Fn_BMIDE_DeepCopyRuleOperations", JavaWindow("Business Modeler"), "AddDeepCopyRule")
			'Checking Existance of "DeepCopyRule" Window
			If Fn_UI_ObjectExist("Fn_BMIDE_DeepCopyRuleOperations",JavaWindow("Business Modeler").JavaWindow("DeepCopyRule") )=True Then
				'Creating Object Of "DeepCopyRule" Window
				Set ObjDeepCopyRule=Fn_UI_ObjectCreate("Fn_BMIDE_DeepCopyRuleOperations",JavaWindow("Business Modeler").JavaWindow("DeepCopyRule") )
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Successfully Invoked the DeepCopyRule Dialog")
			Else
				Exit Function
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Invok the DeepCopyRule Dialog")				
			End If
			If strOperationType<>"" Then
				'Selecting Operation Type 
				Call Fn_List_Select("Fn_BMIDE_DeepCopyRuleOperations", ObjDeepCopyRule, "OperationType",strOperationType)
			End If
			If strRelationType<>"" Then
				'Setting Relation Type
				'Call Fn_Edit_Box("Fn_BMIDE_DeepCopyRuleOperations",ObjDeepCopyRule,"RelationType",strRelationType)
				Call Fn_Button_Click("Fn_BMIDE_DeepCopyRuleOperations",ObjDeepCopyRule, "RelationBrowse")  
				Call Fn_Edit_Box("Fn_BMIDE_DeepCopyRuleOperations", JavaWindow("Business Modeler").JavaWindow("Find Business Object"),"Project",strRelationType)  
				Call Fn_SISW_UI_JavaTable_Operations("Fn_BMIDE_DeepCopyRuleOperations", "SelectCell", JavaWindow("Business Modeler").JavaWindow("Find Business Object").JavaTable("Table") , "", "", "",strRelationType , 0, "", "", "")
				Call Fn_Button_Click("Fn_BMIDE_DeepCopyRuleOperations", JavaWindow("Business Modeler").JavaWindow("Find Business Object"), "OK")
			End If
			If strObjectType<>"" Then
				'Setting Object Type
				'Call Fn_Edit_Box("Fn_BMIDE_DeepCopyRuleOperations",ObjDeepCopyRule,"ObjectType",strObjectType)
				Call Fn_Button_Click("Fn_BMIDE_DeepCopyRuleOperations",ObjDeepCopyRule, "AttachObjectBrowse")  
				Call Fn_Edit_Box("Fn_BMIDE_DeepCopyRuleOperations", JavaWindow("Business Modeler").JavaWindow("Find Business Object"),"Project",strObjectType)  
				Call Fn_SISW_UI_JavaTable_Operations("Fn_BMIDE_DeepCopyRuleOperations", "SelectCell", JavaWindow("Business Modeler").JavaWindow("Find Business Object").JavaTable("Table") , "", "", "",strObjectType , 0, "", "", "")
				Call Fn_Button_Click("Fn_BMIDE_DeepCopyRuleOperations", JavaWindow("Business Modeler").JavaWindow("Find Business Object"), "OK")
			End If
			If strCondition<>"" Then
				'Setting Condition
				Call Fn_Edit_Box("Fn_BMIDE_DeepCopyRuleOperations",ObjDeepCopyRule,"Condition",strCondition)
			End If
			If strActionType<>"" Then
				'Selecting Action
				Call Fn_List_Select("Fn_BMIDE_DeepCopyRuleOperations", ObjDeepCopyRule, "Action",strActionType)
			End If
			If bTargetPri<>"" Then
				'Setting Status Of Target Primary Option
				Call Fn_CheckBox_Set("Fn_BMIDE_DeepCopyRuleOperations", ObjDeepCopyRule, "TargetPrimary", bTargetPri)
			End If
			If bCopyPropOnRel<>"" Then
				'Setting Status Of "Copy Properties On Relation" Option
				Call Fn_CheckBox_Set("Fn_BMIDE_DeepCopyRuleOperations", ObjDeepCopyRule, "CopyPropertiesOnRelation", bCopyPropOnRel)
			End If
			If bRequired<>"" Then
				'Setting Status Of "Required" Option
				Call Fn_CheckBox_Set("Fn_BMIDE_DeepCopyRuleOperations", ObjDeepCopyRule, "Required", bRequired)
			End If
			If bSecured<>"" Then
				'Setting Status Of "Secured" option
				Call Fn_CheckBox_Set("Fn_BMIDE_DeepCopyRuleOperations", ObjDeepCopyRule, "Secured", bSecured)
			End If
			'Clicking On Finish Button To Create "Deep Copy Rule"
			Call Fn_Button_Click("Fn_BMIDE_DeepCopyRuleOperations", ObjDeepCopyRule, "Finish")
			'Function Returns True
			Fn_BMIDE_DeepCopyRuleOperations=True
			'Releasing Object Of "DeepCopyRule" Dialog
			Set ObjDeepCopyRule=Nothing
		'"Verify" case to Verify Existing Deep Copy Rules Value
		Case "Verify"
			'Taking Row Count Of "DeepCopyRulesTable" Table
			iRowCount=Fn_UI_Object_GetROProperty("Fn_BMIDE_DeepCopyRuleOperations",JavaWindow("Business Modeler").JavaTable("DeepCopyRulesTable"),"rows")
			'Selecting Last Row Of "DeepCopyRulesTable" Table
'			JavaWindow("Business Modeler").JavaTable("DeepCopyRulesTable").ActivateRow iRowCount-1
			JavaWindow("Business Modeler").JavaTable("DeepCopyRulesTable").SelectCell iRowCount-1,0
			wait 2
			'Clicking On "EditDeepCopyRule" button To Open "DeepCopyRule" window
			Call Fn_Button_Click("Fn_BMIDE_DeepCopyRuleOperations", JavaWindow("Business Modeler"), "EditDeepCopyRule")
			'Checking Existance of "DeepCopyRule" Dialog
			If Fn_UI_ObjectExist("Fn_BMIDE_DeepCopyRuleOperations",JavaWindow("Business Modeler").JavaWindow("DeepCopyRule") )=True Then
				'Creating Object Of "DeepCopyRule" Dialog
				Set ObjDeepCopyRule=Fn_UI_ObjectCreate("Fn_BMIDE_DeepCopyRuleOperations",JavaWindow("Business Modeler").JavaWindow("DeepCopyRule") )
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Successfully Invoked the DeepCopyRule Dialog")
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Invok the DeepCopyRule Dialog")				
				Exit Function
			End If
			'Initially Setting bFlag=True
			bFlag=True
			If strOperationType<>"" Then
				bFlag=False
				'Taking Current Value Of Operation Type
				currOpType=Fn_UI_Object_GetROProperty("Fn_BMIDE_DeepCopyRuleOperations",ObjDeepCopyRule.JavaList("OperationType"),"value")
				'Verifying Current Value with Expected Value
				If Trim(currOpType)=Trim(strOperationType) Then
					bFlag=True
				End If
			End If
			If bFlag=False Then
				Call Fn_Button_Click("Fn_BMIDE_DeepCopyRuleOperations", ObjDeepCopyRule, "Cancel")
				Exit Function
			End If
	
			If strRelationType<>"" Then
				bFlag=False
				'Taking Current Value Of Relation Type
				currRelType=Fn_UI_Object_GetROProperty("Fn_BMIDE_DeepCopyRuleOperations",ObjDeepCopyRule.JavaEdit("RelationType"),"value")
				'Verifying Current Value with Expected Value
				 If Trim(currRelType)=Trim(strRelationType) Then
					bFlag=True
				End If
			End If
			If bFlag=False Then
				Call Fn_Button_Click("Fn_BMIDE_DeepCopyRuleOperations", ObjDeepCopyRule, "Cancel")
				Exit Function
			End If
	
			If strObjectType<>"" Then
				bFlag=False
				'Taking Current Object Type 
				currObjType=Fn_UI_Object_GetROProperty("Fn_BMIDE_DeepCopyRuleOperations",ObjDeepCopyRule.JavaEdit("ObjectType"),"value")
				'Verifying Current Value with Expected Value
				 If Trim(currObjType)=Trim(currObjType) Then
					bFlag=True
				End If
			End If
			If bFlag=False Then
				Call Fn_Button_Click("Fn_BMIDE_DeepCopyRuleOperations", ObjDeepCopyRule, "Cancel")
				Exit Function
			End If
	
			If strCondition<>"" Then
				bFlag=False
				'Taking Current Condition 
				currCondition=Fn_UI_Object_GetROProperty("Fn_BMIDE_DeepCopyRuleOperations",ObjDeepCopyRule.JavaEdit("Condition"),"value")
				'Verifying Current Value with Expected Value
				 If Trim(currCondition)=Trim(strCondition) Then
					bFlag=True
				End If
			End If
			If bFlag=False Then
				Call Fn_Button_Click("Fn_BMIDE_DeepCopyRuleOperations", ObjDeepCopyRule, "Cancel")
				Exit Function
			End If
	
			If strActionType<>"" Then
				bFlag=False
				'Taking Current Action
				currAction=Fn_UI_Object_GetROProperty("Fn_BMIDE_DeepCopyRuleOperations",ObjDeepCopyRule.JavaList("Action"),"value")
				'Verifying Current Value with Expected Value
				If Trim(currAction)=Trim(strActionType) Then
					bFlag=True
				End If
			End If
			If bFlag=False Then
				Call Fn_Button_Click("Fn_BMIDE_DeepCopyRuleOperations", ObjDeepCopyRule, "Cancel")
				Exit Function
			End If
	
			If bTargetPri<>"" Then
				If UCase(bTargetPri)="ON" Then
					bTargetPri="1"
				ElseIf UCase(bTargetPri)="OFF" Then
					bTargetPri="0"
				End If
				bFlag=False
				'Taking Current Status Of "TargetPrimary" Option
				currTarget=Fn_UI_Object_GetROProperty("Fn_BMIDE_DeepCopyRuleOperations",ObjDeepCopyRule.JavaCheckBox("TargetPrimary"),"value")
				'Verifying Current Value with Expected Value
				If Trim(currTarget)=Trim(bTargetPri) Then
					bFlag=True
				End If
			End If
			If bFlag=False Then
				Call Fn_Button_Click("Fn_BMIDE_DeepCopyRuleOperations", ObjDeepCopyRule, "Cancel")
				Exit Function
			End If
	
			If bCopyPropOnRel<>"" Then
				If UCase(bCopyPropOnRel)="ON" Then
					bCopyPropOnRel="1"
				ElseIf UCase(bCopyPropOnRel)="OFF" Then
					bCopyPropOnRel="0"
				End If
				bFlag=False
				''Taking Current Status Of "CopyPropertiesOnRelation" Option
				currCopyOnRel=Fn_UI_Object_GetROProperty("Fn_BMIDE_DeepCopyRuleOperations",ObjDeepCopyRule.JavaCheckBox("CopyPropertiesOnRelation"),"value")
				'Verifying Current Value with Expected Value
				If Trim(currCopyOnRel)=Trim(bCopyPropOnRel) Then
					bFlag=True
				End If
			End If
			If bFlag=False Then
				Call Fn_Button_Click("Fn_BMIDE_DeepCopyRuleOperations", ObjDeepCopyRule, "Cancel")
				Exit Function
			End If
	
			If bRequired<>"" Then
				If UCase(bRequired)="ON" Then
					bRequired="1"
				ElseIf UCase(bRequired)="OFF" Then
					bRequired="0"
				End If
				bFlag=False
				''Taking Current Status Of "Required" Option
				currReq=Fn_UI_Object_GetROProperty("Fn_BMIDE_DeepCopyRuleOperations",ObjDeepCopyRule.JavaCheckBox("Required"),"value")
				'Verifying Current Value with Expected Value
				If Trim(currReq)=Trim(bRequired) Then
					bFlag=True
				End If
			End If
			If bFlag=False Then
				Call Fn_Button_Click("Fn_BMIDE_DeepCopyRuleOperations", ObjDeepCopyRule, "Cancel")
				Exit Function
			End If
	
			If bSecured<>"" Then
				If UCase(bSecured)="ON" Then
					bSecured="1"
				ElseIf UCase(bSecured)="OFF" Then
					bSecured="0"
				End If
				bFlag=False
				'Taking Current Status Of "Secured" Option
				currSecured=Fn_UI_Object_GetROProperty("Fn_BMIDE_DeepCopyRuleOperations",ObjDeepCopyRule.JavaCheckBox("Secured"),"value")
				'Verifying Current Value with Expected Value
				If Trim(currSecured)=Trim(bSecured) Then
					bFlag=True
				End If
			End If
			If bFlag=False Then
				Call Fn_Button_Click("Fn_BMIDE_DeepCopyRuleOperations", ObjDeepCopyRule, "Cancel")
				Exit Function
			End If
			Call Fn_Button_Click("Fn_BMIDE_DeepCopyRuleOperations", ObjDeepCopyRule, "Finish")
			'Function Returns True
			Fn_BMIDE_DeepCopyRuleOperations=True
			'Releasing Object Of "DeepCopyRule" Dialog
			Set ObjDeepCopyRule=Nothing
		Case "Select"
			iColCount=JavaWindow("Business Modeler").JavaTable("DeepCopyRulesTable").GetROProperty("cols")
			For iCount=0 To iColCount-1
				relationColumn=JavaWindow("Business Modeler").JavaTable("DeepCopyRulesTable").GetColumnName(iCount)
				If Trim(relationColumn)=Trim("Relation Type/Reference Property") Then
					bFlag=True
					Exit For
				End If
			Next
			If bFlag=True Then
				relationColumn="Relation Type/Reference Property"
			Else
				relationColumn="Relation Type"
			End If
			bFlag=False
			iRowCount=Fn_UI_Object_GetROProperty("Fn_BMIDE_DeepCopyRuleOperations",JavaWindow("Business Modeler").JavaTable("DeepCopyRulesTable"),"rows")
			For iCounter=0 To iRowCount-1
				arrRowValue(0)=JavaWindow("Business Modeler").JavaTable("DeepCopyRulesTable").GetCellData(iCounter,"Operation")
				arrRowValue(1)=JavaWindow("Business Modeler").JavaTable("DeepCopyRulesTable").GetCellData(iCounter,relationColumn)
				arrRowValue(2)=JavaWindow("Business Modeler").JavaTable("DeepCopyRulesTable").GetCellData(iCounter,"Attached Business Object")
				arrRowValue(3)=JavaWindow("Business Modeler").JavaTable("DeepCopyRulesTable").GetCellData(iCounter,"Condition")
				arrRowValue(4)=JavaWindow("Business Modeler").JavaTable("DeepCopyRulesTable").GetCellData(iCounter,"Action")
				strFullRowVal=arrRowValue(0)+arrRowValue(1)+arrRowValue(2)+arrRowValue(3)+arrRowValue(4)
				strExpVal=strOperationType+strRelationType+strObjectType+strCondition+strActionType
				If Trim(strFullRowVal)=Trim(strExpVal) Then
'					JavaWindow("Business Modeler").JavaTable("DeepCopyRulesTable").ActivateRow iCounter
					JavaWindow("Business Modeler").JavaTable("DeepCopyRulesTable").SelectCell iCounter,0
					wait 2
					bFlag=True
					Exit For
				End If
			Next
			If bFlag=True Then
				Fn_BMIDE_DeepCopyRuleOperations=True
			End If

        Case "Edit"
			'Clicking On "EditDeepCopyRule" button To Open "DeepCopyRule" window
			Call Fn_Button_Click("Fn_BMIDE_DeepCopyRuleOperations", JavaWindow("Business Modeler"), "EditDeepCopyRule")
			'Checking Existance of "DeepCopyRule" Dialog
			If Fn_UI_ObjectExist("Fn_BMIDE_DeepCopyRuleOperations",JavaWindow("Business Modeler").JavaWindow("DeepCopyRule") )=True Then
				'Creating Object Of "DeepCopyRule" Dialog
				Set ObjDeepCopyRule=Fn_UI_ObjectCreate("Fn_BMIDE_DeepCopyRuleOperations",JavaWindow("Business Modeler").JavaWindow("DeepCopyRule") )
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Successfully Invoked the DeepCopyRule Dialog")
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Invoke the DeepCopyRule Dialog")				
				Exit Function
			End If

			If strOperationType<>"" Then
				'Selecting Operation Type 
				Call Fn_List_Select("Fn_BMIDE_DeepCopyRuleOperations", ObjDeepCopyRule, "OperationType",strOperationType)
			End If
			If strRelationType<>"" Then
				'Setting Relation Type
				Call Fn_Edit_Box("Fn_BMIDE_DeepCopyRuleOperations",ObjDeepCopyRule,"RelationType",strRelationType)
			End If
			If strObjectType<>"" Then
				'Setting Object Type
				Call Fn_Edit_Box("Fn_BMIDE_DeepCopyRuleOperations",ObjDeepCopyRule,"ObjectType",strObjectType)
			End If
			If strCondition<>"" Then
				'Setting Condition
				Call Fn_Edit_Box("Fn_BMIDE_DeepCopyRuleOperations",ObjDeepCopyRule,"Condition",strCondition)
			End If
			If strActionType<>"" Then
				'Selecting Action
				Call Fn_List_Select("Fn_BMIDE_DeepCopyRuleOperations", ObjDeepCopyRule, "Action",strActionType)
			End If
			If bTargetPri<>"" Then
				'Setting Status Of Target Primary Option
				Call Fn_CheckBox_Set("Fn_BMIDE_DeepCopyRuleOperations", ObjDeepCopyRule, "TargetPrimary", bTargetPri)
			End If
			If bCopyPropOnRel<>"" Then
				'Setting Status Of "Copy Properties On Relation" Option
				Call Fn_CheckBox_Set("Fn_BMIDE_DeepCopyRuleOperations", ObjDeepCopyRule, "CopyPropertiesOnRelation", bCopyPropOnRel)
			End If
			If bRequired<>"" Then
				'Setting Status Of "Required" Option
				Call Fn_CheckBox_Set("Fn_BMIDE_DeepCopyRuleOperations", ObjDeepCopyRule, "Required", bRequired)
			End If
			If bSecured<>"" Then
				'Setting Status Of "Secured" option
				Call Fn_CheckBox_Set("Fn_BMIDE_DeepCopyRuleOperations", ObjDeepCopyRule, "Secured", bSecured)
			End If

			'Clicking On Finish Button To Create "Deep Copy Rule"
			Call Fn_Button_Click("Fn_BMIDE_DeepCopyRuleOperations", ObjDeepCopyRule, "Finish")

			'Function Returns True
			Fn_BMIDE_DeepCopyRuleOperations=True

			'Releasing Object Of "DeepCopyRule" Dialog
			Set ObjDeepCopyRule=Nothing
		Case "Remove"
			Call Fn_Button_Click("Fn_BMIDE_DeepCopyRuleOperations", JavaWindow("Business Modeler"), "RemoveDeepCopyRule")
			Fn_BMIDE_DeepCopyRuleOperations=True

		Case "TemplateExistanace"
			Dim arrRowValue1(5)

			iRowCount=Fn_UI_Object_GetROProperty("Fn_BMIDE_DeepCopyRuleOperations",JavaWindow("Business Modeler").JavaTable("DeepCopyRulesTable"),"rows")
			For iCounter=0 To iRowCount-1
				bFlag=False
				arrRowValue1(0)=JavaWindow("Business Modeler").JavaTable("DeepCopyRulesTable").GetCellData(iCounter,"Operation")
				arrRowValue1(1)=JavaWindow("Business Modeler").JavaTable("DeepCopyRulesTable").GetCellData(iCounter,"Relation Type/Reference Property")
				arrRowValue1(2)=JavaWindow("Business Modeler").JavaTable("DeepCopyRulesTable").GetCellData(iCounter,"Attached Business Object")
				arrRowValue1(3)=JavaWindow("Business Modeler").JavaTable("DeepCopyRulesTable").GetCellData(iCounter,"Condition")
				arrRowValue1(4)=JavaWindow("Business Modeler").JavaTable("DeepCopyRulesTable").GetCellData(iCounter,"Action")
				arrRowValue1(5)=JavaWindow("Business Modeler").JavaTable("DeepCopyRulesTable").GetCellData(iCounter,"Template")

				strFullRowVal=arrRowValue1(0)+arrRowValue1(1)+arrRowValue1(2)+arrRowValue1(3)+arrRowValue1(4)+arrRowValue1(5)
				strExpVal=strOperationType+strRelationType+strObjectType+strCondition+strActionType+bTargetPri

				If Trim(strFullRowVal)=Trim(strExpVal) Then
					bFlag=True
					Exit For
				End If
			Next
			If bFlag=True Then
				Fn_BMIDE_DeepCopyRuleOperations=True
			Else
				Fn_BMIDE_DeepCopyRuleOperations=False
			End If

	End Select
End Function

'-------------------------------------------------------------------Function Used to Create New Global Constant----------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_CreateGlobalConstant

'Description			 :	Function Used to Create New Global Constant

'Parameters			   :	1.strName: New Global Constant Name
										'2.strDesc: Global Constant Description
										'3.strDataType:Data Type
										'4.bMultiValued: Is Multi Valued Option
										'5.strDefaultValue:Default Values
										'6.strValues:Values
										'7.bAttachmentSecured:Attachment Secured Option

'Return Value		   : 	True Or False

'Pre-requisite			:	New Global COnstant Dialog Should be Appear on Screen

'Examples				: 	Call Fn_BMIDE_CreateGlobalConstant("Demo","Demo Global Constant","String","Off","Test","","")
'										Call Fn_BMIDE_CreateGlobalConstant("Demo1","Demo1 Global Constant","String","ON","DVal1:DVal2:DVal3","","ON")
'										Call Fn_BMIDE_CreateGlobalConstant("Demo2","Demo2 Global Constant","Boolean","","true","","")
'										Call Fn_BMIDE_CreateGlobalConstant("Demo3","Demo3 Global Constant","List","","Val1","Val1:On~Val2:Off~Val3:On","")
'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done								Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				07/12/2010			           1.0																								Sunny R
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_BMIDE_CreateGlobalConstant(strName,strDesc,strDataType,bMultiValued,strDefaultValue,strValues,bAttachmentSecured)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_CreateGlobalConstant"
   'Variable Declaration
   Dim strPrefix,arrDefaultValues,iCounter,arrValues,arrValueSet
   Dim ObjGlobalConstDialog
	Fn_BMIDE_CreateGlobalConstant=False
	'Checking Existance Of "NewGlobalConstants" Window
	If Fn_UI_ObjectExist("Fn_BMIDE_CreateGlobalConstant",JavaWindow("Business Modeler").JavaWindow("NewGlobalConstants"))=True Then
		'Creating Object Of "NewGlobalConstants" Window
		Set ObjGlobalConstDialog=Fn_UI_ObjectCreate("Fn_BMIDE_CreateGlobalConstant",JavaWindow("Business Modeler").JavaWindow("NewGlobalConstants"))
	Else
	    Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail: NewGlobalConstants Dialog Is Not Exist")
		Exit Function
	End If
	'Retriving Prefix From "Name" Edit Box
	strPrefix=Fn_Edit_Box_GetValue("Fn_BMIDE_CreateGlobalConstant",ObjGlobalConstDialog,"Name")
	'Attaching Prefix To Name
	strName=strPrefix+strName
	'Setting Name To Global Constant
	Call Fn_Edit_Box("Fn_BMIDE_CreateGlobalConstant",ObjGlobalConstDialog,"Name",strName)
	If strDesc<>"" Then
		'Setting Description To New Global Constant
		Call Fn_Edit_Box("Fn_BMIDE_CreateGlobalConstant",ObjGlobalConstDialog,"Description",strDesc)
	End If
	If strDataType<>"" Then
		'Selecting Data Type
		Call Fn_List_Select("Fn_BMIDE_CreateGlobalConstant",ObjGlobalConstDialog,"DataType",strDataType)
	End If
	If LCase(strDataType)="string" Then
		If Trim(UCase(bMultiValued))="ON" Then
			'Setting Status Of "Is Multi Valued?" Option To ON
			Call Fn_CheckBox_Set("Fn_BMIDE_CreateGlobalConstant",ObjGlobalConstDialog,"IsMultiValued","On")
			If strDefaultValue<>"" Then
				arrDefaultValues=Split(strDefaultValue,":")
				'Adding Multiple Default Values
				For iCounter=0 To Ubound(arrDefaultValues)
					'Clicking "Add" Button to Invoke "AddValues" Dialog
					Call Fn_Button_Click("Fn_BMIDE_CreateGlobalConstant",ObjGlobalConstDialog,"Add")
					Call Fn_Edit_Box("Fn_BMIDE_CreateGlobalConstant",JavaWindow("Business Modeler").JavaWindow("AddGlobalConstValue"),"Value",arrDefaultValues(iCounter))
					Call Fn_Button_Click("Fn_BMIDE_CreateGlobalConstant",JavaWindow("Business Modeler").JavaWindow("AddGlobalConstValue"),"Finish")
				Next
			End If
			'Setting Status Of "Is Attachment Secured?" Option
			If bAttachmentSecured<>"" Then
				Call Fn_CheckBox_Set("Fn_BMIDE_CreateGlobalConstant",ObjGlobalConstDialog,"IsAttachmentSecured",bAttachmentSecured)
			End If
		Else
			'Setting Status Of "Is Multi Valued?" Option To OFF
			Call Fn_CheckBox_Set("Fn_BMIDE_CreateGlobalConstant",ObjGlobalConstDialog,"IsMultiValued","Off")
			'Setting Default Value
			If strDefaultValue<>"" Then
				Call Fn_Edit_Box("Fn_BMIDE_CreateGlobalConstant",ObjGlobalConstDialog,"DefaultValue",strDefaultValue)
			End If
		End If
	ElseIf LCase(strDataType)="boolean" Then
		'Setting Default Value
			If strDefaultValue<>"" Then
				Call Fn_List_Select("Fn_BMIDE_CreateGlobalConstant",ObjGlobalConstDialog,"DefaultValueList",strDefaultValue)
			End If
	ElseIf LCase(strDataType)="list" Then
			'Setting Multiple Values
			If strValues<>"" Then
				'Multiple Values Separeted by (~) Tilda
				arrValues=Split(strValues,"~")
				For iCounter=0 To Ubound(arrValues)
					arrValueSet=Split(arrValues(iCounter),":")
					'Clicking "Add" Button to Invoke "AddValues" Dialog
					Call Fn_Button_Click("Fn_BMIDE_CreateGlobalConstant",ObjGlobalConstDialog,"Add")
					Call Fn_Edit_Box("Fn_BMIDE_CreateGlobalConstant",JavaWindow("Business Modeler").JavaWindow("AddGlobalConstValue"),"Value",arrValueSet(0))
					If arrValueSet(1)<>"" Then
						Call Fn_CheckBox_Set("Fn_BMIDE_CreateGlobalConstant",JavaWindow("Business Modeler").JavaWindow("AddGlobalConstValue"),"Secured",arrValueSet(1))
					End If
					Call Fn_Button_Click("Fn_BMIDE_CreateGlobalConstant",JavaWindow("Business Modeler").JavaWindow("AddGlobalConstValue"),"Finish")
				Next
			End If
			'Setting Default Value
			If strDefaultValue<>"" Then
				'Call Fn_Edit_Box("Fn_BMIDE_CreateGlobalConstant",ObjGlobalConstDialog,"DefaultValue",strDefaultValue)
				Call Fn_Button_Click("Fn_BMIDE_CreateGlobalConstant", ObjGlobalConstDialog,"BrowseDefaultValue")
				Call Fn_Edit_Box("Fn_BMIDE_CreateGlobalConstant",JavaWindow("Business Modeler").JavaWindow("SelectValue"),"ValueEdit",strDefaultValue)
				Call Fn_Button_Click("Fn_BMIDE_CreateGlobalConstant",JavaWindow("Business Modeler").JavaWindow("SelectValue"),"OK")
			End If
	End If
	'Clicking "Finish" Button to Create New Global Constants
	Call Fn_Button_Click("Fn_BMIDE_CreateGlobalConstant",ObjGlobalConstDialog,"Finish")
	Fn_BMIDE_CreateGlobalConstant=True
	'Releasing Object Of "Global Constants Dialog"
	Set ObjGlobalConstDialog=Nothing
End Function

'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_CreateIDContext

'Description			 :	Function Used to Create :
										'1.New ID Context
										'2.New List Of Occurance Type
										'3.New Status
										'4.New Unit Of Measure
										'5.New View Type

'Parameters			   :	1.strName:Name
										'2.strDisplayName: Display Name
										'3.strDescription:Description

'Return Value		   : 	True Or False

'Pre-requisite			:	New Global COnstant Dialog Should be Appear on Screen

'Examples				: 	Call Fn_BMIDE_CreateIDContext("Test","Demo","Test Contect ID")

'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done								Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				07/12/2010			           1.0																								Sunny R
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'IMP NOTE:
'This Function Use to Create 
										'1.New ID Context
										'2.New List Of Occurance Type
										'3.New Status
										'4.New Unit Of Measure
										'5.New View Type
Public Function Fn_BMIDE_CreateIDContext(strName,strDisplayName,strDescription)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_CreateIDContext"
   'Variable Declaration
   Dim strPrefix
   Dim ObjIDContextDialog
	Fn_BMIDE_CreateIDContext=False
	'Checking Existance Of "NewIdContext" Window
	If Fn_UI_ObjectExist("Fn_BMIDE_CreateIDContext",JavaWindow("Business Modeler").JavaWindow("NewIdContext"))=True Then
		'Creating Object Of "NewIdContext" Window
		Set ObjIDContextDialog=Fn_UI_ObjectCreate("Fn_BMIDE_CreateIDContext",JavaWindow("Business Modeler").JavaWindow("NewIdContext"))
	Else
	    Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail: NewIdContext Dialog Is Not Exist")
		Exit Function
	End If
	'Retriving Prefix From "Name" Edit Box
	strPrefix=Fn_Edit_Box_GetValue("Fn_BMIDE_CreateIDContext",ObjIDContextDialog,"Name")
	'Attaching Prefix To Name
	strName=strPrefix+strName
	'Setting Name To ID Context
	Call Fn_Edit_Box("Fn_BMIDE_CreateIDContext",ObjIDContextDialog,"Name",strName)
	If strDisplayName<>"" Then
		'Setting Display Name To ID Context
		Call Fn_Edit_Box("Fn_BMIDE_CreateIDContext",ObjIDContextDialog,"DisplayName",strDisplayName)
	End If

	If strDescription<>"" Then
		'Setting Description
		Call Fn_Edit_Box("Fn_BMIDE_CreateIDContext",ObjIDContextDialog,"Description",strDescription)
	End If

	'Clicking "Finish" Button to Create New ID Context
	Call Fn_Button_Click("Fn_BMIDE_CreateIDContext",ObjIDContextDialog,"Finish")
	Fn_BMIDE_CreateIDContext=True
	'Releasing Object Of "ID Context Dialog"
	Set ObjIDContextDialog=Nothing
End Function
'-------------------------------------------------------------------Function Used to Create Bussiness Context-----------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_NewBussinessContext

'Description			 :	Function Used to Create Bussiness Context

'Parameters			   :   '1.strName:Name
										'2.strDesc:Description
										'3.strConnectionSettings:Connection Settings To Connect to Server
										'												"ProjectName:ProfileName:Password:Group:Role"
										'4.strAccessors: Accessors 

'Return Value		   : 	True Or False

'Pre-requisite			:	New Bussiness Context Dialog Should be appear

'Examples				: 	Call Fn_BMIDE_NewBussinessContext("Demo","Demo BussinessContext","Temp:Test:AutoTestDBA","Organization:Engineering:Designer~Organization:dba:DBA")

'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				07/12/2010			           1.0																				Sunny R
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'strConnectionSettings: - "ProjectName:ProfileName:Password:Group:Role"
'strAccessors:- "Accessor1~Accessor2~Accessor3"
Public Function Fn_BMIDE_NewBussinessContext(strName,strDesc,strConnectionSettings,strAccessors)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_NewBussinessContext"
	'Varaible Declaration
	Dim ObjContextDialog,ObjConnectionDialog,ObjAccessorDialog
	Dim arrConnSet,arrAccessor,iCounter,strPrefix
	Fn_BMIDE_NewBussinessContext=False
	'Creating Object Of Dialogs
	Set ObjContextDialog=JavaWindow("Business Modeler").JavaWindow("NewBusinessContext")
	Set ObjConnectionDialog= JavaWindow("Business Modeler").JavaWindow("TeamcenterRepositoryConnection")
	Set ObjAccessorDialog=JavaWindow("Business Modeler").JavaWindow("AccessorSelectionDialog")
			If strName<>"" Then
				'Setting Name To New Bussiness Context
				strPrefix=Fn_Edit_Box_GetValue("Fn_BMIDE_NewBussinessContext",ObjContextDialog,"Name")
				strName=strPrefix+strName
				Call Fn_Edit_Box("Fn_BMIDE_NewBussinessContext",ObjContextDialog,"Name",strName)
			End If
			If strDesc<>"" Then
				'Setting Description
				Call Fn_Edit_Box("Fn_BMIDE_NewBussinessContext",ObjContextDialog,"Description",strDesc)
			End If
			 If strAccessors<>"" Then
						'Selecting Organization
						Call Fn_Button_Click("Fn_BMIDE_NewBussinessContext",ObjContextDialog,"Add")
						'Checking Existance Of "TeamcenterRepositoryConnection" Dialog
						If Fn_UI_ObjectExist("Fn_BMIDE_NewBussinessContext", ObjConnectionDialog)=True Then	
									arrConnSet=Split(strConnectionSettings,":")
									If arrConnSet(0)<>"" Then
										'Selecting Project From Project List
										Call Fn_List_Select("Fn_BMIDE_NewBussinessContext", ObjConnectionDialog, "Project",arrConnSet(0))
									End If
									If arrConnSet(1)<>"" Then
										'Selecting Profile From Profile List
										If ObjConnectionDialog.JavaList("ServerProfile").GetROProperty("enabled") = "1"  Then
											Call Fn_List_Select("Fn_BMIDE_NewBussinessContext", ObjConnectionDialog, "ServerProfile",arrConnSet(1))
										End If
									End If
									wait(5)
									If  Fn_UI_ObjectExist("Fn_BMIDE_NewBussinessContext", ObjConnectionDialog.JavaEdit("Password"))=True Then
										If arrConnSet(2)<>"" Then
											'Setting password 
											Call Fn_Edit_Box("Fn_BMIDE_NewBussinessContext",ObjConnectionDialog,"Password",arrConnSet(2))
										End If
									End If
									If  Fn_UI_ObjectExist("Fn_BMIDE_NewBussinessContext", ObjConnectionDialog.JavaEdit("Group"))=True Then
										If arrConnSet(3)<>"" Then
											'Setting Group
											Call Fn_Edit_Box("Fn_BMIDE_NewBussinessContext",ObjConnectionDialog,"Group",arrConnSet(3))
										End If
									End If
									If  Fn_UI_ObjectExist("Fn_BMIDE_NewBussinessContext", ObjConnectionDialog.JavaEdit("Role"))=True Then
										If arrConnSet(4)<>"" Then
											'Setting Role
											Call Fn_Edit_Box("Fn_BMIDE_NewBussinessContext",ObjConnectionDialog,"Role",arrConnSet(4))
										End If
									End If
									'Clicking "Connect" Button To Connect To The Host
									If Fn_UI_ObjectExist("Fn_BMIDE_DeployProject",ObjConnectionDialog.JavaButton("Connect"))=True Then
										Call Fn_Button_Click("Fn_BMIDE_DeployProject",ObjConnectionDialog, "Connect")
									End If
									ObjConnectionDialog.JavaButton("Finish").WaitProperty "enabled","1",iTime
									'Clicking "Finish" Button
									Call Fn_Button_Click("Fn_BMIDE_DeployProject",ObjConnectionDialog, "Finish")
							End If
							arrAccessor=Split(strAccessors,"~")
							'Checking Existance of Search Organisation Dialog
							If  Fn_UI_ObjectExist("Fn_BMIDE_NewBussinessContext", ObjAccessorDialog)=True Then
								'Selecting Organisation From OrganizationTree 
								Call Fn_JavaTree_Select("Fn_BMIDE_NewBussinessContext", ObjAccessorDialog, "AccessorTree",arrAccessor(0))
								Call Fn_Button_Click("Fn_BMIDE_NewBussinessContext",ObjAccessorDialog, "Finish")
							End If				
							If UBound(arrAccessor)>=1 Then
								For iCounter=1 To UBound(arrAccessor)
										Call Fn_Button_Click("Fn_BMIDE_NewBussinessContext",ObjContextDialog,"Add")
										Call Fn_JavaTree_Select("Fn_BMIDE_NewBussinessContext", ObjAccessorDialog, "AccessorTree",arrAccessor(0))
										Call Fn_Button_Click("Fn_BMIDE_NewBussinessContext",ObjAccessorDialog, "Finish")
								Next
							End If
				End If
			'Clicking Finish Button to Create New Bussiness Context
			Call Fn_Button_Click("Fn_BMIDE_NewBussinessContext",ObjContextDialog,"Finish")
			Fn_BMIDE_NewBussinessContext=True
			'Releasing Object Of All Dialogs
			Set ObjContextDialog=Nothing
			Set ObjConnectionDialog=Nothing
			Set ObjAccessorDialog=Nothing
End Function

'-------------------------------------------------------------------Function Used to Verify Error Message which Appears while creating Project-------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_NewProjectErrorVerify

'Description			 :	Function Used to Verify Error Message which Appears while creating Project

'Parameters			   :	1.strAction: Action Name
										'2.strProjectName: Project Name
										'3.bDefaultOpt:Default Option
										'4.strLocation: Project Location
										'5.strPrefix:Project Prefix
										'6.strTempDirectory:Template Derectory
										'7.strDepdTemplate: Dependant Templates
										'8.strErrMsg: Expected Error Message

'Return Value		   : 	True Or False

'Pre-requisite			:	Should be Log In BMIDE

'Examples				: 	strMsg= "Invalid ""Prefix:"" field. The first character of the prefix must be an upper case letter."
'										Fn_BMIDE_NewProjectErrorVerify("IncorrectPrefix","Demo","","","00","","",strMsg)

'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done								Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep									   				07/12/2010			           1.0																								Sunny R
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_BMIDE_NewProjectErrorVerify(strAction,strProjectName,bDefaultOpt,strLocation,strPrefix,strTempDirectory,strDepdTemplate,strErrMsg)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_NewProjectErrorVerify"
   'Variable declaration
   Dim StrErrorMessage,strMenu
   Dim ObjNewProjectWindow
   'Setting Function equals to False
   Fn_BMIDE_NewProjectErrorVerify=False
   'Checking Existance of NewProject window
	If Fn_UI_ObjectExist("Fn_BMIDE_NewProjectErrorVerify",JavaWindow("Business Modeler").JavaWindow("NewProject"))=False Then
		'Taking Menu Name from Environmet File
		strMenu=Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\BMIDEConfig\BMIDE_Menu.xml", "NewProject")
		'Calling File:New:Project... Menu to open Project Dialog
        Call Fn_BMIDE_MenuOperation("Select", strMenu)
	End If
	'Creating Object Of NewProject window
	Set ObjNewProjectWindow=Fn_UI_ObjectCreate("Fn_BMIDE_NewProjectErrorVerify", JavaWindow("Business Modeler").JavaWindow("NewProject"))
	'Expanding Business Modeler IDE Node
    Call Fn_UI_JavaTree_Expand("Fn_BMIDE_NewProjectErrorVerify", ObjNewProjectWindow, "WizardsTree","Business Modeler IDE")
	'Selecting Project
	Call Fn_JavaTree_Select("Fn_BMIDE_NewProjectErrorVerify", ObjNewProjectWindow,"WizardsTree","Business Modeler IDE:New Business Modeler IDE Template Project")
	'Clicking On Next button to Go Next Wizard
	Call Fn_Button_Click("Fn_BMIDE_NewProjectErrorVerify", ObjNewProjectWindow, "Next")
	Select Case strAction
		Case "IncorrectPrefix"
		If strProjectName<>"" Then
			'Seeting project Name
			Call Fn_Edit_Box("Fn_BMIDE_NewProjectErrorVerify",ObjNewProjectWindow,"ProjectName",strProjectName)
		End If
		If bDefaultOpt=Cstr(True) Then
			If strLocation<>"" Then
				'Setting Project Location
				Call Fn_Edit_Box("Fn_BMIDE_NewProjectErrorVerify",ObjNewProjectWindow,"UseDefaultLocation",strLocation)
			End If
		End If
		'Clicking On Next button to Go Next Wizard
'		Call Fn_Button_Click("Fn_BMIDE_NewProjectErrorVerify", ObjNewProjectWindow, "Next")
		Call Fn_Edit_Box("Fn_BMIDE_NewProjectErrorVerify",ObjNewProjectWindow,"TemplateDesc","Test")
		'Setting Prefix To project
		Call Fn_Edit_Box("Fn_BMIDE_NewProjectErrorVerify",ObjNewProjectWindow,"Prefix",strPrefix)
		
        StrErrorMessage=Fn_Edit_Box_GetValue("Fn_BMIDE_NewProjectErrorVerify",ObjNewProjectWindow,"ErrorMsg")
		If Trim(StrErrorMessage)=Trim(strErrMsg) Then
			Fn_BMIDE_NewProjectErrorVerify=True
		End If
		Call Fn_Button_Click("Fn_BMIDE_NewProjectErrorVerify", ObjNewProjectWindow, "Cancel")
		Set ObjNewProjectWindow=Nothing
	End Select
End Function

'-------------------------------------------------------------------Function Used to Create New Business Object Constant----------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_CreateBusinessObjectConstant

'Description			 :	Function Used to Create New Business Object Constant

'Parameters			   :	1.strName: New BusinessObject Constant Name
										'2.strDesc: BusinessObject Constant Description
										'3.strDataType:Data Type
										'4.strScope: Scope Values  Seperated by  "~"
										'5.strDefaultValue:Default Values
										'6.strValues:Values ot  Enter when Datatype is  'List' seperated by "~"  and value of checkbox ON/OFF

'Return Value		   : 	True Or False

'Pre-requisite			:	New Business Object Constant Dialog Should be Appear on Screen

'Examples				: 						--	For Creating Business Object Constants  --
'											Fn_BMIDE_CreateBusinessObjectConstant("Demo","Demo","*","String","A","")
'										 Fn_BMIDE_CreateBusinessObjectConstant("Demo1","Demo","AbsOccData~*","Boolean","true","")
'										Fn_BMIDE_CreateBusinessObjectConstant("Demo2","Demo","AbsOccData~*","List","B","A:ON~B:OFF~C:ON")

'															--	For Creating property Constants  --
'										Fn_BMIDE_CreateBusinessObjectConstant("Demo","Demo","*","String","A","")
'										 Fn_BMIDE_CreateBusinessObjectConstant("Demo1","Demo","AbsOccData:lsd~*","Boolean","true","")
'										 Fn_BMIDE_CreateBusinessObjectConstant("Demo2","Demo","AbsOccData:lsd~*","List","B","A:ON~B:OFF~C:ON")
'													for property Constants 
'																strScope parameter should be as follows
'																AbsOccData:lsd -   here 		AbsOccData - Business object scope
'																																			isd	- property scope
'																for multiple values  -     seperate by  '~'			
'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done								Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Pranav Ingle								   				07/12/2010			           1.0																								Sunny R
'													pranav Ingle												08/12/2010			           1.0																								Sunny R
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_BMIDE_CreateBusinessObjectConstant(strName,strDesc,strScope,strDataType,strDefaultValue,strValues)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_CreateBusinessObjectConstant"
   'Variable Declaration
   Dim strPrefix,arrDefaultValues,iCounter,arrValues,arrValueSet,arrScope,arrListValue,arrPropScope
   Dim ObjGlobalConstDialog,ObjScopeDialog,ObjListValues,WshShell
	Fn_BMIDE_CreateBusinessObjectConstant=False
	'Checking Existance Of "NewBusinessObjectConstant" Window
	If Fn_UI_ObjectExist("Fn_BMIDE_CreateBusinessObjectConstant",JavaWindow("Business Modeler").JavaWindow("NewBusinessObjectConstant"))=True Then
		'Creating Object Of "NewBusinessObjectConstant" Window
		Set ObjGlobalConstDialog=Fn_UI_ObjectCreate("Fn_BMIDE_CreateBusinessObjectConstant",JavaWindow("Business Modeler").JavaWindow("NewBusinessObjectConstant"))
	Else
	    Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail: NewBusinessObjectConstant Dialog Is Not Exist")
		Exit Function
	End If
	'Retriving Prefix From "Name" Edit Box
	strPrefix=Fn_Edit_Box_GetValue("Fn_BMIDE_CreateBusinessObjectConstant",ObjGlobalConstDialog,"Name")
	'Attaching Prefix To Name
	strName=strPrefix+strName
	'Setting Name To Global Constant
	Call Fn_Edit_Box("Fn_BMIDE_CreateBusinessObjectConstant",ObjGlobalConstDialog,"Name",strName)
	If strDesc<>"" Then
		'Setting Description To New Global Constant
		Call Fn_Edit_Box("Fn_BMIDE_CreateBusinessObjectConstant",ObjGlobalConstDialog,"Description",strDesc)
	End If
'	Setting Scopes To New Global Constant
	If strScope<>"" Then
			arrScope= Split(strScope,"~")
			For iCounter = 0 to UBound(arrScope)
					arrPropScope = Split(arrScope(iCounter),":")
					' Click  on add button
					Call Fn_Button_Click("Fn_BMIDE_CreateBusinessObjectConstant",ObjGlobalConstDialog,"Add")
					If Fn_UI_ObjectExist("Fn_BMIDE_CreateBusinessObjectConstant",JavaWindow("Business Modeler").JavaWindow("NewBusinessObjectDefineScope"))=True Then
							'Creating Object Of "NewBusinessObjectDefineScope" Window
							Set ObjScopeDialog=Fn_UI_ObjectCreate("Fn_BMIDE_CreateBusinessObjectConstant",JavaWindow("Business Modeler").JavaWindow("NewBusinessObjectDefineScope"))
					Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail: NewBusinessObjectDefineScope Dialog Is Not Exist")
							Exit Function
					End If
					' Set  Business Object Scope value
					Call Fn_Edit_Box("Fn_BMIDE_CreateBusinessObjectConstant",ObjScopeDialog,"BusinessObjectScope",arrPropScope(0))
					If UBound(arrPropScope) = 1  Then
							' Set  Property Scope value
							Call Fn_Edit_Box("Fn_BMIDE_CreateBusinessObjectConstant",ObjScopeDialog,"PropertyScope",arrPropScope(1))
					End If
					' Click on finish Button
					Call Fn_Button_Click("Fn_BMIDE_CreateBusinessObjectConstant",ObjScopeDialog,"Finish")	
					Set ObjScopeDialog = Nothing
			Next
	End If
	If strDataType<>"" Then
		'Selecting Data Type
		Call Fn_List_Select("Fn_BMIDE_CreateBusinessObjectConstant",ObjGlobalConstDialog,"DataType",strDataType)
	End If
	If strDefaultValue<>"" Then
			If strDataType="String" Then
					Call Fn_Edit_Box("Fn_BMIDE_CreateBusinessObjectConstant",ObjGlobalConstDialog,"DefaultValue",strDefaultValue)
			End If
			If strDataType="Boolean" Then
					Call Fn_List_Select("Fn_BMIDE_CreateBusinessObjectConstant",ObjGlobalConstDialog,"DefaultValue",LCase(strDefaultValue))
			End If
			If strDataType="List" Then
					arrValues = Split(strValues,"~")
					For iCounter =0 to UBound(arrValues)
							arrListValue = Split(arrValues(iCounter),":")
							' Click  on add button
							Call Fn_Button_Click("Fn_BMIDE_CreateBusinessObjectConstant",ObjGlobalConstDialog,"Add2")
							If Fn_UI_ObjectExist("Fn_BMIDE_CreateBusinessObjectConstant",JavaWindow("Business Modeler").JavaWindow("NewBusinessObjectAddValue"))=True Then
									'Creating Object Of "NewBusinessObjectAddValue" Window
									Set ObjListValues=Fn_UI_ObjectCreate("Fn_BMIDE_CreateBusinessObjectConstant",JavaWindow("Business Modeler").JavaWindow("NewBusinessObjectAddValue"))
							Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail: NewBusinessObjectAddValue  Dialog Is Not Exist")
									Exit Function
							End If
							Call Fn_Edit_Box("Fn_BMIDE_CreateBusinessObjectConstant",ObjListValues,"Value",arrListValue(0))
							If  UCase(arrListValue(1)) = "ON" Then
								Call Fn_CheckBox_Select("Fn_BMIDE_CreateBusinessObjectConstant",ObjListValues,"Secured")
							End If
							' Click on finish Button
							Call Fn_Button_Click("Fn_BMIDE_CreateBusinessObjectConstant",ObjListValues,"Finish")	
							Set ObjListValues = Nothing	
					Next
					'Selecting Default value
					Call Fn_Edit_Box("Fn_BMIDE_CreateBusinessObjectConstant",ObjGlobalConstDialog,"DefaultValue",strDefaultValue)
					wait 1
					Set WshShell = CreateObject("WScript.Shell")
					WshShell.SendKeys "{ESC}"
					wait 1,500
					Set WshShell = nothing
    		End If
	End If
	'Clicking "Finish" Button to Create New Global Constants
	Call Fn_Button_Click("Fn_BMIDE_CreateBusinessObjectConstant",ObjGlobalConstDialog,"Finish")
	Fn_BMIDE_CreateBusinessObjectConstant=True
	'Releasing Object Of "Global Constants Dialog"
	Set ObjGlobalConstDialog=Nothing
End Function

'-------------------------------------------------------------------Function Used to Create New Condition---------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_CreateCondition

'Description			 :	Function Used to Create New Condition

'Parameters			   :	1.strName: New Condition Name
										'2.strDesc: Condition Description
										'3.bSecured: Secured Option (ON\OFF)
										'4.strInputParameter: Input Parameters Name
										'5.strSignature:Condition Signature
										'6.strExpression:Expression
										'7.bAttachmentSecured:Attachment Secured Option

'Return Value		   : 	True Or False

'Pre-requisite			:	New Condition Dialog Should be Appear on Screen

'Examples				: 	Call Fn_BMIDE_CreateCondition("DemoCondition","Demo Condition Description","ON","Custom","AbsOccFlags:Para1~AbsOccData:Para2","Test")
										'strInputParameter : - 1)Bussiness Object 2)Bussiness Object and User Session 3)Custom
'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done								Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				07/12/2010			           1.0																								Sunny R
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'strSignature : - If Input Parameters Is "Custom" Then Pass values ( : ) Colan Separated
							'Eg. :- "Parameter Type:Parameter Name ~ Parameter Type:Parameter Name"
							'Else Pass simple Value
							'Eg :- Parameter Type
Public Function Fn_BMIDE_CreateCondition(strName,strDesc,bSecured,strInputParameter,strSignature,strExpression)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_CreateCondition"
   'Variable Declaration
	Dim ObjConditionDialog
	Dim strPrefix,arrSignature,iCounter,arrSignatureValue
	Fn_BMIDE_CreateCondition=False
	'Checking Existance of "NewCondition" Window
   If Fn_UI_ObjectExist("Fn_BMIDE_CreateCondition",JavaWindow("Business Modeler").JavaWindow("NewCondition"))=True Then
	   'Creating Object Of "NewCondition" Window
		Set ObjConditionDialog=Fn_UI_ObjectCreate("Fn_BMIDE_CreateCondition",JavaWindow("Business Modeler").JavaWindow("NewCondition"))
	Else
		Exit Function
   End If
	If strName<>"" Then
		'Taking Project Prefix From 
		strPrefix=Fn_Edit_Box_GetValue("Fn_BMIDE_CreateCondition",ObjConditionDialog,"Name")
		strName=strPrefix+strName
		'Setting Condition Name
		Call Fn_Edit_Box("Fn_BMIDE_CreateCondition",ObjConditionDialog,"Name",strName)
	End If
	If strDesc<>"" Then
		'Setting Condition Description
		Call Fn_Edit_Box("Fn_BMIDE_CreateCondition",ObjConditionDialog,"Description",strDesc)
	End If
	If bSecured<>"" Then
		'Setting Status of Secured Option
		Call Fn_CheckBox_Set("Fn_BMIDE_CreateCondition",ObjConditionDialog,"Secured", bSecured)
	End If
	If strInputParameter<>"" Then
		'Setting Input Parameters Option
		Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_BMIDE_CreateCondition",ObjConditionDialog.JavaRadioButton("InputParameters"),"attached text",strInputParameter)
		Call Fn_UI_JavaRadioButton_SetON("Fn_BMIDE_CreateCondition",ObjConditionDialog,"InputParameters")
	End If
	If Trim(strInputParameter)="Custom" Then
		'Setting Signature
		If strSignature<>"" Then
			arrSignature=Split(strSignature,"~")
			Call Fn_Button_Click("Fn_BMIDE_CreateCondition",ObjConditionDialog,"Browse")
			For iCounter=0 To UBound(arrSignature)
				Call Fn_Button_Click("Fn_BMIDE_CreateCondition",JavaWindow("Business Modeler").JavaWindow("ConditionCustomParameters"),"Add")
				arrSignatureValue=Split(arrSignature(iCounter),":")
				Call Fn_Edit_Box("Fn_BMIDE_CreateCondition",JavaWindow("Business Modeler").JavaWindow("NewConditionParameters"),"ParameterType",arrSignatureValue(0))
				Call Fn_Edit_Box("Fn_BMIDE_CreateCondition",JavaWindow("Business Modeler").JavaWindow("NewConditionParameters"),"ParameterName",arrSignatureValue(1))
				Call Fn_Button_Click("Fn_BMIDE_CreateCondition",JavaWindow("Business Modeler").JavaWindow("NewConditionParameters"),"Finish")
			Next
			Call Fn_Button_Click("Fn_BMIDE_CreateCondition",JavaWindow("Business Modeler").JavaWindow("ConditionCustomParameters"),"Finish")
		End If
	Else
		If strSignature<>"" Then
			Call Fn_Button_Click("Fn_BMIDE_CreateCondition",ObjConditionDialog,"Browse")
			Call Fn_Edit_Box("Fn_BMIDE_CreateCondition",JavaWindow("Business Modeler").JavaWindow("Find Business Object"),"Project",strSignature)
			Call Fn_Button_Click("Fn_BMIDE_CreateCondition",JavaWindow("Business Modeler").JavaWindow("Find Business Object"),"OK")
		End If
	End If
	If strExpression<>"" Then
		'Setting Expression
		Call Fn_Edit_Box("Fn_BMIDE_CreateCondition",ObjConditionDialog,"Expression",strExpression)
	End If
	'Clicking On Finish Button To Create New Condition
	Call Fn_Button_Click("Fn_BMIDE_CreateCondition",ObjConditionDialog,"Finish")
	Fn_BMIDE_CreateCondition=True
	Set ObjConditionDialog=Nothing
End Function

'------------------------------------------------------------'Function Used to Create AliasId Rules-----------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_CreateAliasIdRule

'Description			 :	Function Used to Create Alternate ID Rules

'Parameters			   :   '1.strIdentifierContext :  Button Name and  Indentifier Context Value seperated by "~"
'										2.strIdentifierType		 :	Button Name and  Indentifier Type Value seperated by "~"
'										3. strDesciption  		  :   Alias id Description

'Return Value		   : 	True Or False

'Examples				:	 Fn_BMIDE_CreateAliasIdRule("Browse~D3Demo","Browse~Identifier","New Alias Id")
'										 Fn_BMIDE_CreateAliasIdRule("New~Demo1:DemoDisplayName:DemoDesc","Browse~Identifier","New Alias Id")
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Pranav  Ingle											08-Dec-2010								1.0																						Sunny
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------	
'strIdentifierContext : - If   'New'  Then Pass values ( : ) Colan Separated
										'Eg. :- "New~ IdentifierContextName: Identifier Context Display Name : Identifier Context Description"
'										If  'Browse' then Pass IdentifierContextName directly

Public Function Fn_BMIDE_CreateAliasIdRule(strIdentifierContext,strIdentifierType,strDescription)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_CreateAliasIdRule"
	Dim ObjAliasIdDialog, arrIdentifierContext,arrNewIdContext,arrIdentifierType
	Dim strPrefix
	Fn_BMIDE_CreateAliasIdRule=False

	If Fn_UI_ObjectExist("Fn_BMIDE_CreateAlternateIDRule",JavaWindow("Business Modeler").JavaWindow("NewAlternateIdRule"))=True Then
			'Creating Object Of "NewAliasIdRule" Window
			Set ObjAliasIdDialog=Fn_UI_ObjectCreate("Fn_BMIDE_CreateAliasIdRule",JavaWindow("Business Modeler").JavaWindow("NewAlternateIdRule"))
	Else 
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail: AliasIdDialog does not exist")
			Exit Function
	End If
	
	arrIdentifierContext = Split(strIdentifierContext,"~")
	' Set value Identifier  Context
	Select Case arrIdentifierContext(0)
			Case "Browse"
					Call Fn_Edit_Box("Fn_BMIDE_CreateAliasIdRule",ObjAliasIdDialog,"IdentifierContext",arrIdentifierContext(1))
			Case "New"
					Call Fn_Button_Click("Fn_BMIDE_CreateAliasIdRule", ObjAliasIdDialog, "New")
					arrNewIdContext = Split(arrIdentifierContext(1),":")
					'Taking Prefix from Name Edit Box
					strPrefix=Fn_Edit_Box_GetValue("Fn_BMIDE_CreateAliasIdRule",JavaWindow("Business Modeler").JavaWindow("NewIdentifierContext"),"Name")
					arrNewIdContext(0)=strPrefix+arrNewIdContext(0)
					'Setting  Name Of Identifier Context
					Call Fn_Edit_Box("Fn_BMIDE_CreateAliasIdRule",JavaWindow("Business Modeler").JavaWindow("NewIdentifierContext"),"Name",arrNewIdContext(0))
					'Setting Display Name Of Identifier Context
					Call Fn_Edit_Box("Fn_BMIDE_CreateAliasIdRule",JavaWindow("Business Modeler").JavaWindow("NewIdentifierContext"),"DisplayName",arrNewIdContext(1))
                	'Setting Description Name Of Identifier Context
					Call Fn_Edit_Box("Fn_BMIDE_CreateAliasIdRule",JavaWindow("Business Modeler").JavaWindow("NewIdentifierContext"),"Description",arrNewIdContext(2))
					'Clicking On Finish Button To Create New Identifier Context
					Call Fn_Button_Click("Fn_BMIDE_CreateAliasIdRule", JavaWindow("Business Modeler").JavaWindow("NewIdentifierContext"), "Finish")
					Fn_BMIDE_CreateAliasIdRule = True	
			 Case else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail: failed wrong Case entered for Identifier Context")
					 Exit Function
	End Select

	' Set value Identifier  Type
	arrIdentifierType = Split(strIdentifierType,"~")
	Select Case arrIdentifierType(0)
			Case "Browse"
					'Setting Identifier type
					Call Fn_Edit_Box("Fn_BMIDE_CreateAliasIdRule",ObjAliasIdDialog,"IdentifierType",arrIdentifierType(1))
			Case "New"
					'  yet to modify
			Case else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail: failed wrong Case entered for Identifier Type")
					 Exit Function
	End Select
	' Set value of Descrption
	Call Fn_Edit_Box("Fn_BMIDE_CreateAliasIdRule",ObjAliasIdDialog,"Description",strDescription)
	
	'Clicking on Finish Button To Create New AliasID Rule
	Call Fn_Button_Click("Fn_BMIDE_CreateAliasIdRule", ObjAliasIdDialog, "Finish")
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Pass: Successfully Created AliasId Rule")
			
	Fn_BMIDE_CreateAliasIdRule = True
	'Releasing Object Of AiasId Dialog
	Set ObjAliasIdDialog=Nothing
End Function

'-------------------------------------------------------------------Function Used to Perform Operations On Project Properties------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_ProjectPropertiesOperation

'Description			 :	Function Used to Perform Operations On Project Properties

'Parameters			   :	1.strAction: Action Name
										'2.strTempDisplayName:Template Display Name
										'3.strPrefix: Prefix
										'4.strTempDescription: Template Description
										'5.strDepdTempDirectory:Dependant Template Derectory
										'6.strDepdTemplate:Dependant Template

'Return Value		   : 	True Or False Or Status Of Check Box (ON Or OFF)

'Pre-requisite			:	Should be Log In BMIDE

'Examples				: 	Call Fn_BMIDE_ProjectPropertiesOperation("Modify","","X3","","","")
										'IMP NOTE : - User Have to Select Project In Navigator tab
										'In "Modify" Case Pass the Parameters which have to modify 
										'Case "TemplateOptionCurrentStatus" will Return Current State Of Template Option Check Boxes
										'strDepdTemplate :-  Pass Check Box Name
'										Fn_BMIDE_ProjectPropertiesOperation("TemplateOptionCurrentStatus","","","","","Enable Operational Data Updates?")
'										Case "SetTemplateOptionStatus" To set status of Template Option Chk
'										strDepdTemplate :-  Pass Check Box Name:Status
'										Fn_BMIDE_ProjectPropertiesOperation("SetTemplateOptionStatus","","","","","Enable Operational Data Updates?:Off")
'										Fn_BMIDE_ProjectPropertiesOperation("SetTemplateOptionStatus","","","","","Enable Operational Data Updates?:On")
'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done													Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				08/12/2010			           1.0																																		Sunny R
'													Sandeep N										   				11/01/2011			           1.1								Case "TemplateOptionCurrentStatus"		  Sunny R
'																																																				Case "SetTemplateOptionStatus"
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Fn_BMIDE_ProjectPropertiesOperation("Modify","","X3","","","")
Public Function Fn_BMIDE_ProjectPropertiesOperation(strAction,strTempDisplayName,strPrefix,strTempDescription,strDepdTempDirectory,strDepdTemplate)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_ProjectPropertiesOperation"
   'Variable Declaration
   Dim strMenu,strCurrStatus,arrTempOpt
   Dim ObjPropertiesDialog
   Fn_BMIDE_ProjectPropertiesOperation=False
   Set ObjPropertiesDialog=JavaWindow("Business Modeler").JavaWindow("ProjectProperties")
	If Not ObjPropertiesDialog.Exist(10) Then
		'Taking Menu Name from Environmet File
		strMenu=Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\BMIDEConfig\BMIDE_Menu.xml", "ProjectProperties")
		'Calling File:New:Project... Menu to open Project Dialog
        Call Fn_BMIDE_MenuOperation("Select", strMenu)	
	End If

    'Expanding Teamcenter Node
    Call Fn_UI_JavaTree_Expand("Fn_BMIDE_ProjectPropertiesOperation", ObjPropertiesDialog, "PropertiesTree","Teamcenter")
    'Selecting BMIDE Node
	Call Fn_JavaTree_Select("Fn_BMIDE_ProjectPropertiesOperation", ObjPropertiesDialog,"PropertiesTree","Teamcenter:BMIDE")
	Select Case strAction
			Case "Modify"
				If strTempDisplayName<>"" Then
					'Modifying Template Display Name
                    Call Fn_Edit_Box("Fn_BMIDE_ProjectPropertiesOperation",ObjPropertiesDialog,"TemplateDisplayName",strTempDisplayName)
				End If
				If strPrefix<>"" Then
					'Modifying Prefix
                    Call Fn_Edit_Box("Fn_BMIDE_ProjectPropertiesOperation",ObjPropertiesDialog,"Prefix",strPrefix)
				End If
				If strTempDescription<>"" Then
					'Modifying Template Description
                    Call Fn_Edit_Box("Fn_BMIDE_ProjectPropertiesOperation",ObjPropertiesDialog,"TemplateDescription",strTempDescription)
				End If
				If strDepdTempDirectory<>"" Then
					'Modifying Dependant Template Derectory
                    Call Fn_Edit_Box("Fn_BMIDE_ProjectPropertiesOperation",ObjPropertiesDialog,"DependentTempDirectory",strDepdTempDirectory)
				End If
				'Clicking OK Button To Modify Project Properties
                Call Fn_Button_Click("Fn_BMIDE_ProjectPropertiesOperation", ObjPropertiesDialog, "OK")
				Fn_BMIDE_ProjectPropertiesOperation=True

		Case "TemplateOptionCurrentStatus"
				If strDepdTemplate<>"" Then
					ObjPropertiesDialog.JavaCheckBox("TemplateOption").SetTOProperty "attached text",strDepdTemplate
					strCurrStatus=Fn_UI_Object_GetROProperty("Fn_BMIDE_ProjectPropertiesOperation",ObjPropertiesDialog.JavaCheckBox("TemplateOption"), "value")
					If CInt(strCurrStatus)=1 Then
						Fn_BMIDE_ProjectPropertiesOperation="ON"
					Else
						Fn_BMIDE_ProjectPropertiesOperation="OFF"
					End If
				End If
				Call Fn_Button_Click("Fn_BMIDE_ProjectPropertiesOperation", ObjPropertiesDialog, "OK")

			Case "SetTemplateOptionStatus"
				If strDepdTemplate<>"" Then
					arrTempOpt=Split(strDepdTemplate,":")
					ObjPropertiesDialog.JavaCheckBox("TemplateOption").SetTOProperty "attached text",arrTempOpt(0)
					Call Fn_CheckBox_Set("Fn_BMIDE_ProjectPropertiesOperation", ObjPropertiesDialog, "TemplateOption", arrTempOpt(1))
				End If
				Call Fn_Button_Click("Fn_BMIDE_ProjectPropertiesOperation", ObjPropertiesDialog, "OK")
				Fn_BMIDE_ProjectPropertiesOperation=True

	End Select
	'Releasing Properties Dialog Object
	Set ObjPropertiesDialog=Nothing
End Function

'-------------------------------------------------------------------Function Used to Connect to Teamcenter Repository---------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_TeamcenterRepositoryConnection

'Description			 :	Function Used to Connect to Teamcenter Repository

'Parameters			   :   '1.StrProjectName:Action to Perform
										'2.StrProfile:Organization Node Path
										'3.StrUserID: Project Name For Connection
										'4.StrPassword: Profile Name
										'5:StrGroup : Password For Connection
										'6:StrRole: Group name

'Return Value		   : 	True Or False

'Pre-requisite			:	Should be Log In BMIDE

'Examples				: 	Call Fn_BMIDE_TeamcenterRepositoryConnection("Temp","","AutoTestDBA","dba","DBA","isTrue","ON")

'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				08/12/2010			           1.0																				Sunny R
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_BMIDE_TeamcenterRepositoryConnection(StrProjectName,StrProfile,StrUserID,StrPassword,StrGroup,StrRole)
		GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_TeamcenterRepositoryConnection"
		Fn_BMIDE_TeamcenterRepositoryConnection=False
		'Creating Object Of "TeamcenterRepositoryConnection" Window
		Set ObjConnectionDialog=JavaWindow("Business Modeler").JavaWindow("TeamcenterRepositoryConnection")
		If StrProjectName<>"" Then
			'Selecting Project From Project List
			Call Fn_List_Select("Fn_BMIDE_TeamcenterRepositoryConnection", ObjConnectionDialog, "Project",StrProjectName)
		End If
		If StrProfile<>"" Then
			'Selecting Profile From Profile List
			Call Fn_List_Select("Fn_BMIDE_TeamcenterRepositoryConnection", ObjConnectionDialog, "ServerProfile",StrProfile)
		End If
		If  Fn_UI_ObjectExist("Fn_BMIDE_TeamcenterRepositoryConnection", ObjConnectionDialog.JavaEdit("Password"))=True Then
			If StrPassword<>"" Then
				'Setting password 
				Call Fn_Edit_Box("Fn_BMIDE_TeamcenterRepositoryConnection",ObjConnectionDialog,"Password",StrPassword)
			Else
				StrPassword = Fn_UI_Object_GetROProperty("", ObjConnectionDialog.JavaEdit("UserID") ,"text")
				Call Fn_Edit_Box("Fn_BMIDE_TeamcenterRepositoryConnection",ObjConnectionDialog,"Password",StrPassword)
			End If
		End If
		If  Fn_UI_ObjectExist("Fn_BMIDE_TeamcenterRepositoryConnection", ObjConnectionDialog.JavaEdit("Group"))=True Then
			If StrGroup<>"" Then
				'Setting Group
				Call Fn_Edit_Box("Fn_BMIDE_TeamcenterRepositoryConnection",ObjConnectionDialog,"Group",StrGroup)
			End If
		End If
		If  Fn_UI_ObjectExist("Fn_BMIDE_TeamcenterRepositoryConnection", ObjConnectionDialog.JavaEdit("Role"))=True Then
			If StrRole<>"" Then
				'Setting Role
				Call Fn_Edit_Box("Fn_BMIDE_TeamcenterRepositoryConnection",ObjConnectionDialog,"Role",StrRole)
			End If
		End If
		'Clicking "Connect" Button To Connect To The Host
		If Fn_UI_ObjectExist("Fn_BMIDE_DeployProject",ObjConnectionDialog.JavaButton("Connect"))=True Then
			Call Fn_Button_Click("Fn_BMIDE_DeployProject",ObjConnectionDialog, "Connect")
		End If
		ObjConnectionDialog.JavaButton("Finish").WaitProperty "enabled","1",iTime
		'Clicking "Finish" Button
		Call Fn_Button_Click("Fn_BMIDE_TeamcenterRepositoryConnection",ObjConnectionDialog, "Finish")
		Fn_BMIDE_TeamcenterRepositoryConnection=True
End Function

'-------------------------------------------------------------------Function Used to Create New Extension Defination-------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_CreateExtensionDefination

'Description			 :	Function Used to Create New Extension Defination

'Parameters			   :   '1.strName:Extension Defination Name
										'2.strLanguage:Language
										'3.strLibrary: Library
										'4.bInternal:Is Internal Extention Defination Option
										'5:strParameterList : Parameter List
										'6:strAvailability: Availability
										'7.strConnectionSett: Connection Settings 

'Return Value		   : 	True Or False

'Pre-requisite			:	New Extension Defination Dialog Should be Open ,Should be Log In BMIDE

'Examples				: 	Call Fn_BMIDE_CreateExtensionDefination("DemoExtension","ANSI_C","X3TestLib","","","","Temp:TestProfile::AutoTestDBA:dba:DBA")

'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				08/12/2010			           1.0																				Sunny R
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Fn_BMIDE_CreateExtensionDefination("DemoExtension","ANSI_C","X3TestLib","","","","Temp:TestProfile::AutoTestDBA:dba:DBA")

'strParameterList : - "Name:Type:Mandetory:Suggested Value:LOV Name:Query Name~ Name:Type:Mandetory:Suggested Value:LOV Name:Query Name"
'										"DemoPara:String:On:None::~DemoPara1:Integer:Off:TcLOV:BillCodes:"
'										"DemoPara2:Double:On:TcQuery::__Admin - Audit"
'strAvailability : -"Object Name:Property:Property Name:Operation Name:Extension Point~Object Name:Property:Property Name:Operation Name:Extension Point"
'								"AbsOccData:Property:Test Prop:Test Operation:PreCondition~AbsOccData:Type::Test Operation:PreCondition"
'strConnectionSett:-"Project Name : Profile Name:User Name:Password:Group:Role" 
'											"Temp:TestProfile::AutoTestDBA:dba:DBA"
'strConnectionSett : Mandetory Parameter
Public Function Fn_BMIDE_CreateExtensionDefination(strName,strLanguage,strLibrary,bInternal,strParameterList,strAvailability,strConnectionSett)
		GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_CreateExtensionDefination"
		Dim ObjExtensionDialog
		Dim strPrefix,arrParaList,iCounter,arrParaListValue,arrAvail,arrAvailValue
		Fn_BMIDE_CreateExtensionDefination=False
		Set ObjExtensionDialog=JavaWindow("Business Modeler").JavaWindow("NewExtensionDefinition")
		If Fn_UI_ObjectExist("Fn_BMIDE_CreateExtensionDefination",JavaWindow("Business Modeler").JavaWindow("NewExtensionDefinition"))=True Then
            Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: NewExtensionDefinition window Is Exist On Screen")
		Else
			Set ObjExtensionDialog=Nothing
			Exit Function
		End If
		If strName<>"" Then
			'Setting Name To Extension Defination
			strPrefix=Fn_Edit_Box_GetValue("Fn_BMIDE_CreateExtensionDefination",ObjExtensionDialog,"Name")
			strName=strPrefix+strName
			Call Fn_Edit_Box("Fn_BMIDE_CreateExtensionDefination",ObjExtensionDialog,"Name",strName)
		End If
		If strLanguage<>"" Then
			'Selecting Language
			Call Fn_List_Select("Fn_BMIDE_CreateExtensionDefination",ObjExtensionDialog,"Language",strLanguage)
		End If
		If strLibrary<>"" Then
			'Setting Library
			Call Fn_Edit_Box("Fn_BMIDE_CreateExtensionDefination",ObjExtensionDialog,"Library",strLibrary)
		End If
		If bInternal<>"" Then
			'Setting Internal Option
			Call Fn_CheckBox_Set("Fn_BMIDE_CreateExtensionDefination", ObjExtensionDialog,"Internal",bInternal)
		End If
		If strParameterList<>"" Then
			'Setting Multiple Parameter List
			arrParaList=Split(strParameterList,"~")
			For iCounter=0 To Ubound(arrParaList)
				Call Fn_Button_Click("Fn_BMIDE_CreateExtensionDefination",ObjExtensionDialog,"AddParameter")
				arrParaListValue=Split(arrParaList(iCounter),":")
				Call Fn_BMIDE_ExtensionParameterOperation("Add",arrParaListValue(0),arrParaListValue(1),arrParaListValue(2),arrParaListValue(3),arrParaListValue(4),arrParaListValue(5),strConnectionSett)
			Next
		End If
		If strAvailability<>"" Then
			'Setting Multiple Parameter List
			arrAvail=Split(strAvailability,"~")
			For iCounter=0 To Ubound(arrAvail)
				Call Fn_Button_Click("Fn_BMIDE_CreateExtensionDefination",ObjExtensionDialog,"AddAvailability")
				arrAvailValue=Split(arrAvail(iCounter),":")
				Call Fn_BMIDE_ExtensionAvailabilityOperations("Add",arrAvailValue(0),arrAvailValue(1),arrAvailValue(2),arrAvailValue(3),arrAvailValue(4))
			Next
		End If
		'Clicking On Finish Button To Create New Extention Defination
		Call Fn_Button_Click("Fn_BMIDE_CreateExtensionDefination",ObjExtensionDialog,"Finish")
		Fn_BMIDE_CreateExtensionDefination=True
		Set ObjExtensionDialog=Nothing
End Function


'-------------------------------------------------------------------Function Used to Perform Operation On Extension Availability-------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_ExtensionAvailabilityOperations

'Description			 :	Function Used to Perform Operation On Extension Availability

'Parameters			   :   '1.strAction:Action to Perform
										'2.strObjName:Bussiness Object Name
										'3.strPropetyOptn: Property option Name
										'4.strPropertyName: Property Name
										'5:strOperationName : Operation Name
										'6:strExtensionPoint: Extension Point

'Return Value		   : 	True Or False

'Pre-requisite			:	Should be Log In BMIDE

'Examples				: 	Call Fn_BMIDE_ExtensionAvailabilityOperations("Add","AbsOccData","Property","TestProp","TestOpration","PreCondition")
'										Call Fn_BMIDE_ExtensionAvailabilityOperations("Add","AbsOccData","Type","","TestOpration","PreCondition")

'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				08/12/2010			           1.0																				Sunny R
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_BMIDE_ExtensionAvailabilityOperations(strAction,strObjName,strPropetyOptn,strPropertyName,strOperationName,strExtensionPoint)
		GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_ExtensionAvailabilityOperations"
		Dim ObjExtAvl
		Fn_BMIDE_ExtensionAvailabilityOperations=False
		Set ObjExtAvl=JavaWindow("Business Modeler").JavaWindow("NewExtensionAvailability")
		Select Case strAction
				Case "Add"
					If Fn_UI_ObjectExist("Fn_BMIDE_ExtensionAvailabilityOperations",ObjExtAvl)=True Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: NewExtensionAvailability Dialog Is Exist")
					Else
				
					End If
					If strObjName<>"" Then
						'Setting Object Name
						Call Fn_Edit_Box("Fn_BMIDE_ExtensionAvailabilityOperations",ObjExtAvl,"BusinessObjectName",strObjName)
					End If
					If strPropetyOptn<>"" Then
						'Selecting Property Option
						Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_BMIDE_ExtensionAvailabilityOperations",ObjExtAvl.JavaRadioButton("Property"),"attached text",strPropetyOptn)
						Call Fn_UI_JavaRadioButton_SetON("Fn_BMIDE_ExtensionAvailabilityOperations",ObjExtAvl, "Property")
						If strPropetyOptn="Property" Then
							If strPropertyName<>"" Then
								'Selecting Property
								Call Fn_List_Select("Fn_BMIDE_ExtensionAvailabilityOperations", ObjExtAvl, "PropertyName",strPropertyName)
							End If
						End If
					End If
					If strOperationName<>"" Then
						'Selecting Operation
						Call Fn_List_Select("Fn_BMIDE_ExtensionAvailabilityOperations", ObjExtAvl, "OperationName",strOperationName)
					End If
					If strExtensionPoint<>"" Then
						'Selecting Extension Point
						Call Fn_List_Select("Fn_BMIDE_ExtensionAvailabilityOperations", ObjExtAvl, "ExtensionPoint",strExtensionPoint)
					End If
					Call Fn_Button_Click("Fn_BMIDE_ExtensionAvailabilityOperations", ObjExtAvl, "Finish")
					Fn_BMIDE_ExtensionAvailabilityOperations=True
		End Select
		Set ObjExtAvl=Nothing
End Function

'-------------------------------------------------------------------Function Used to Perform Operation On Extension Parameters-------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_ExtensionParameterOperation

'Description			 :	Function Used to Perform Operation On Extension Parameters

'Parameters			   :   '1.strAction:Action to Perform
										'2.strName:Extension Parameter Name
										'3.strType: Parametr Type
										'4.bMandetory: Mandetory Option
										'5:strSuggestedValue : Suggested Value
										'6:strLOVName: LOV Name
										'7.strQueryName: Query Name
										'8.strConnectionSetting: Connection Setting Values (Pass These parameter Always)

'Return Value		   : 	True Or False

'Pre-requisite			:	Should be Log In BMIDE

'Examples				: 	Call Fn_BMIDE_ExtensionParameterOperation("Add","Test","String","Off","TcQuery","","__Admin - Audit","Temp:TestProfile::AutoTestDBA:dba:DBA")

'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				08/12/2010			           1.0																				Sunny R
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Call Fn_BMIDE_ExtensionParameterOperation("Add","Test","String","Off","TcQuery","","__Admin - Audit","Temp:TestProfile::AutoTestDBA:dba:DBA")
'strConnectionSetting="Project Name:Profile Name:UserID:Password:Group:Role"
'UserID Pass Blank
Public Function Fn_BMIDE_ExtensionParameterOperation(strAction,strName,strType,bMandetory,strSuggestedValue,strLOVName,strQueryName,strConnectionSetting)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_ExtensionParameterOperation"
   Dim ObjExtParaDialog
   Dim arrConSet
   Fn_BMIDE_ExtensionParameterOperation=False
	Set ObjExtParaDialog=JavaWindow("Business Modeler").JavaWindow("NewExtensionParameter")
	Select Case strAction
		Case "Add"
				
				If Fn_UI_ObjectExist("Fn_BMIDE_ExtensionParameterOperation",ObjExtParaDialog)=True Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: NewExtensionParameter Dialog Is Exist On Screen")	
				Else
					'If User want to Open the dialog Explicitly then its provision
				End If
				If strName<>"" Then
					'Setting Name To Extension parameter
					Call Fn_Edit_Box("Fn_BMIDE_ExtensionParameterOperation",ObjExtParaDialog,"Name",strName)
				End If
				If strType<>"" Then
					'Setting Type
					Call Fn_List_Select("Fn_BMIDE_ExtensionParameterOperation",ObjExtParaDialog, "Type",strType)
				End If
				If bMandetory<>"" Then
					'Setting Status of Mandetory Option
					Call Fn_CheckBox_Set("Fn_BMIDE_ExtensionParameterOperation", ObjExtParaDialog, "Mandatory", bMandetory)
				End If
				'Selecting Suggested Value Type
				If strSuggestedValue<>"" Then
					Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_BMIDE_ExtensionParameterOperation",ObjExtParaDialog.JavaRadioButton("SuggestedValue"),"attached text",strSuggestedValue)
					Call Fn_UI_JavaRadioButton_SetON("Fn_BMIDE_ExtensionParameterOperation",ObjExtParaDialog, "SuggestedValue")
					If Trim(strSuggestedValue)="TcLOV" Then
						Call Fn_Edit_Box("Fn_BMIDE_ExtensionParameterOperation",ObjExtParaDialog,"LOVName",strLOVName)
					ElseIf Trim(strSuggestedValue)="TcQurery" Then
						Call Fn_Button_Click("Fn_BMIDE_ExtensionParameterOperation",ObjExtParaDialog,"BrowseQueryName")
						If Fn_UI_ObjectExist("Fn_BMIDE_ExtensionParameterOperation",JavaWindow("Business Modeler").JavaWindow("TeamcenterRepositoryConnection"))=True Then
							'Connecting To Teamcenter Repository
							arrConSet=Split(strConnectionSetting,":")
							Call Fn_BMIDE_TeamcenterRepositoryConnection(arrConSet(0),arrConSet(1),"",arrConSet(3),arrConSet(4),,arrConSet(5))
						End If
							Call Fn_Edit_Box("Fn_BMIDE_ExtensionParameterOperation",JavaWindow("Business Modeler").JavaWindow("Find Business Object"),"Criteria",strLOVName)
							Call Fn_Button_Click("Fn_BMIDE_ExtensionParameterOperation", JavaWindow("Business Modeler").JavaWindow("Find Business Object"),"OK")
					End If
				End If
				Call Fn_Button_Click("Fn_BMIDE_ExtensionParameterOperation",ObjExtParaDialog,"Finish")
				Fn_BMIDE_ExtensionParameterOperation=True

	End Select
	'Releasing Object ExtensionParameter Dialog
	Set ObjExtParaDialog=Nothing
End Function
'--------------------------------------------------Function Used Save Data Model---------------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_SaveDataModel

'Description			 :	Function Used Save Data Model
'
'Parameters:			:	1.strProjects: Project Name

'Return Value		   : 	True Or False

'Pre-requisite			:	Should be Log In BMIDE

'Examples				: 	'Call Fn_BMIDE_SaveDataModel("Trial")


'History					 :		
'													Developer Name												Date						Rev. No.						Changes Done				Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				08/12/2010			           1.0														Sunny
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_BMIDE_SaveDataModel(strProjects)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_SaveDataModel"
	Dim strMenu,ObjSaveDataModelWnd,iRowCount,arrProjects,iCounter,bFlag,iCount,strPrjName

		strMenu=Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\BMIDEConfig\BMIDE_Menu.xml", "SaveDataModel")
		Call Fn_BMIDE_MenuOperation("Select", strMenu)
	   'Checking Existance Of "SaveDataModel" Window
	   If Fn_UI_ObjectExist("Fn_BMIDE_SaveDataModel", JavaWindow("Business Modeler").JavaWindow("SaveDataModel"))=True Then
			'Creating Object of "SaveDataModel" Window
			Set ObjSaveDataModelWnd=Fn_UI_ObjectCreate("Fn_BMIDE_SaveDataModel", JavaWindow("Business Modeler").JavaWindow("SaveDataModel"))
			If strProjects<>"" Then
				'First Deselecting All Project 
				Call Fn_Button_Click("Fn_BMIDE_SaveDataModel", ObjSaveDataModelWnd, "DeselectAll")
				'Taking Row Count Of Project Table
				iRowCount=Fn_UI_Object_GetROProperty("Fn_BMIDE_SaveDataModel",ObjSaveDataModelWnd.JavaTable("Project"), "rows")
					arrProjects=Split(strProjects,":")
					For iCounter=0 To Ubound(arrProjects)
						bFlag=False
						For iCount=0 To iRowCount-1
							'Taking Row Data (Project name) From project Table
							strPrjName=ObjSaveDataModelWnd.JavaTable("Project").GetCellData(iCount,0)
							If Trim(arrProjects(iCounter))=Trim(strPrjName) Then
								'Selecting Row
								ObjSaveDataModelWnd.JavaTable("Project").SelectCell iCount,0
								ObjSaveDataModelWnd.JavaTable("Project").PressKey " "
								bFlag=True
								Exit For
							End If
						Next
					Next
				End If	
				'Releasing Object of "SaveDataModel" Window
				Set ObjSaveDataModelWnd=Nothing
		Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Pass : Data Model is already Saved [Save dialog not invoked]")
			Fn_BMIDE_SaveDataModel = True
			Exit Function
	    End If

		If bFlag=True Then
			'Clicking On OK Button to Exit BMIDE
			Call Fn_Button_Click("Fn_BMIDE_SaveDataModel", JavaWindow("Business Modeler").JavaWindow("SaveDataModel"), "OK")
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Pass :Successfully Saved Data Model")
			'Function Returns True
			Fn_BMIDE_SaveDataModel=True
		Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail : Invalid Project names ["+strProjects+"] pass by user thats why Save Data Model Canceled")
			Call Fn_Button_Click("Fn_BMIDE_SaveDataModel", JavaWindow("Business Modeler").JavaWindow("SaveDataModel"), "Cancel")
		End If

End Function

'-------------------------------------------------------------------Function Used to Create New Library-------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_CreateLibrary

'Description			 :	Function Used to Create New Library

'Parameters			   :   '1.bIsThirdParty: Is Third Party Option
										'2.strName:Library Name
										'3.strDesc: Library Description
										'4.strDependantOn: Dependant On Libraries List

'Return Value		   : 	True Or False

'Pre-requisite			:	New Lirary Dialog Should Be Open & Should be Log In BMIDE

'Examples				: 	Call Fn_BMIDE_CreateLibrary("Off","TestLibrary","This Is Test Library","Off","archive:cfm:tc")
'										Call Fn_BMIDE_CreateLibrary("","DemoLibrary","This Is Demo Library","ON","archive")

'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				09/12/2010			           1.0																				Sunny R
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'strDependantOn : - "Library:Library:Library"
'										"archive:cfm:tc"
Public Function Fn_BMIDE_CreateLibrary(bIsThirdParty,strName,strDesc,bSetActiveLibrary,strDependantOn)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_CreateLibrary"
   Dim strPrefix,arrDepdOn,iCounter
   Dim ObjLibraryDialog
	Fn_BMIDE_CreateLibrary=False
   If Fn_UI_ObjectExist("Fn_BMIDE_CreateLibrary", JavaWindow("Business Modeler").JavaWindow("NewLibrary"))=True Then
	   Set ObjLibraryDialog=JavaWindow("Business Modeler").JavaWindow("NewLibrary")
	Else
		Exit Function
   End If
   'Setting Status Of Is Third Party Option
	If bIsThirdParty<>"" Then
        Call Fn_CheckBox_Set("Fn_BMIDE_CreateLibrary",ObjLibraryDialog,"isThirdParty", bIsThirdParty)
	End If
	'Setting Name To Library
	If strName<>"" Then
		strPrefix=Fn_Edit_Box_GetValue("Fn_BMIDE_CreateLibrary",ObjLibraryDialog,"Name")
		strName=strPrefix+strName
		Call Fn_Edit_Box("Fn_BMIDE_CreateLibrary",ObjLibraryDialog,"Name",strName)
	End If
	'Setting Description To Library
	If strDesc<>"" Then
		Call Fn_Edit_Box("Fn_BMIDE_CreateLibrary",ObjLibraryDialog,"Description",strDesc)
	End If
	 'Setting Status Of Set As Active Library Option
	If bSetActiveLibrary<>"" Then
        Call Fn_CheckBox_Set("Fn_BMIDE_CreateLibrary",ObjLibraryDialog,"SetAsActiveLibrary", bSetActiveLibrary)
	End If
	If strDependantOn<>"" Then
		'Adding Dependant On Libraries
		arrDepdOn=Split(strDependantOn,":")
		For iCounter=0 To Ubound(arrDepdOn)
			Call Fn_Button_Click("Fn_BMIDE_CreateLibrary",ObjLibraryDialog,"Add")
			Call Fn_Edit_Box("Fn_BMIDE_CreateLibrary",JavaWindow("Business Modeler").JavaWindow("Find Business Object"),"Project",arrDepdOn(iCounter))
			Call Fn_Button_Click("Fn_BMIDE_CreateLibrary",JavaWindow("Business Modeler").JavaWindow("Find Business Object"),"OK")
		Next
	End If
	Call Fn_Button_Click("Fn_BMIDE_CreateLibrary",ObjLibraryDialog,"Finish")
   Fn_BMIDE_CreateLibrary=True
   Set ObjLibraryDialog=Nothing
End Function
'-------------------------------------------------------------------Function Used to Verify Message Which Appears while Deleting Object---------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_DeleteObjectErrorMsgVerify

'Description			 :	Function Used to Verify Message Which Appears while Deleting Object

'Parameters			   :   '1.strErrorMsg: Error Message

'Return Value		   : 	True Or False

'Pre-requisite			:	Delete Object Dialog Should be Appear On Screen

'Examples				: 	strErrMsg="This Business Object cannot be deleted because it is referenced by the following IDContextRules Rules: "" X3fhgdf@Identifier; """
'										Call Fn_BMIDE_DeleteObjectErrorMsgVerify(strErrMsg)

'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				09/12/2010			           1.0																				Sunny R
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'This Function Only Verify Error Messages Which Comes while Deleting Objects
'To Delete Object Use Function Fn_BMIDE_DeleteObject()
Public Function Fn_BMIDE_DeleteObjectErrorMsgVerify(strErrorMsg)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_DeleteObjectErrorMsgVerify"
   'Variable Declaration
   Dim bFlag,strMsg,iCounter,iRowCount
   Dim ObjDeleteDialog
   'Initially Function Returns False
   bFlag=False
   Fn_BMIDE_DeleteObjectErrorMsgVerify=False
   'Chaking Existance Of DeleteObject Dioalog
   If Fn_UI_ObjectExist("Fn_BMIDE_DeleteObjectErrorMsgVerify",JavaWindow("Business Modeler").JavaWindow("Deleteobject"))=True Then
	   Set ObjDeleteDialog=JavaWindow("Business Modeler").JavaWindow("Deleteobject")
   Else
	   'If Delete Object Not Exist Then Function Will Exit From Next Statement
		Exit Function
   End If
   If Fn_UI_ObjectExist("Fn_BMIDE_DeleteObjectErrorMsgVerify",ObjDeleteDialog.JavaTable("ErrorTable"))=True Then
		iRowCount=Fn_UI_Object_GetROProperty("Fn_BMIDE_DeleteObjectErrorMsgVerify",ObjDeleteDialog.JavaTable("ErrorTable"), "rows")
		For iCounter=0 To iRowCount-1
			strMsg=ObjDeleteDialog.JavaTable("ErrorTable").GetCellData(iCounter,"Messages")
			If InStr(1,Trim(LCase(strMsg)),Trim(LCase(strErrorMsg)))>=1 Then
				bFlag=True
				Exit For
			End If
		Next
	Else
		If strErrorMsg<>"" Then
				strMsg=Fn_UI_Object_GetROProperty("Fn_BMIDE_DeleteObjectErrorMsgVerify",ObjDeleteDialog.JavaEdit("ErrorEditBox"), "value")
				If InStr(1,Trim(LCase(strMsg)),Trim(LCase(strErrorMsg)))>=1 Then
					bFlag=True
				End If
		Else
			bFlag=True
		End If
   End If
   If bFlag=False Then
	    Call Fn_Button_Click("Fn_BMIDE_DeleteObjectErrorMsgVerify",ObjDeleteDialog,"Cancel")
		Set ObjDeleteDialog=Nothing
	   Exit Function
   End If
   'Clicking On Cancel Button to Delete Object
   Call Fn_Button_Click("Fn_BMIDE_DeleteObjectErrorMsgVerify",ObjDeleteDialog,"Cancel")
   Fn_BMIDE_DeleteObjectErrorMsgVerify=True
   Set ObjDeleteDialog=Nothing
End Function

'-------------------------------------------------------------------Function Used to Create New Note Type-------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_CreateNoteType

'Description			 :	Function Used to Create New Note Type

'Parameters			   :   '1.strName: Note Type Name
										'2.strDisplayName:Note Type Display Name
										'3.strDescription: Note Type Description
										'4.bAttachValueList: Attach Value List Option
										'5.strLOV:LOV Name
										'6:strDefaultValue:Default Value

'Return Value		   : 	True Or False

'Pre-requisite			:	 "New Note Type" Dialog Should be appear on screen

'Examples				: 	Call Fn_BMIDE_CreateNoteType("Test","DemoNoteType","New Demo Note Type","Off","","")

'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				16/12/2010			           1.0																				Sunny R
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'PreRequisite : - "New Note Type" Dialog Should be appear on screen
Public Function Fn_BMIDE_CreateNoteType(strName,strDisplayName,strDescription,bAttachValueList,strLOV,strDefaultValue)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_CreateNoteType"
   'Variable Declaration
   Dim ObjNoteDialog
   Dim strPrefix
   Fn_BMIDE_CreateNoteType=False
   'Creating Object Of "NewNoteType" window
   Set ObjNoteDialog=JavaWindow("Business Modeler").JavaWindow("NewNoteType")
   'Checking Existance of "NewNoteType" window
   If Fn_UI_ObjectExist("Fn_BMIDE_CreateNoteType", ObjNoteDialog)=True Then
       Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: NewNoteType Window is Exist")
	Else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail: NewNoteType Window is Not Exist")
		Set ObjNoteDialog=Nothing
		Exit Function
   End If
   'Setting Name To Note Type
   If strName<>"" Then
	   strPrefix=Fn_Edit_Box_GetValue("Fn_BMIDE_CreateNoteType",ObjNoteDialog,"Name")
	   strName=strPrefix+strName
	   Call Fn_Edit_Box("Fn_BMIDE_CreateNoteType",ObjNoteDialog,"Name",strName)
   End If
   'Setting Display Name To Note Type
   If strDisplayName<>"" Then
		Call Fn_Edit_Box("Fn_BMIDE_CreateNoteType",ObjNoteDialog,"DisplayName",strDisplayName)
   End If
   'Setting Description To Note Type
	If strDescription<>"" Then
		Call Fn_Edit_Box("Fn_BMIDE_CreateNoteType",ObjNoteDialog,"Description",strDescription)
	End If
	'Setting Status Of "Attach Value List" Option
	If bAttachValueList<>"" Then
		Call Fn_CheckBox_Set("Fn_BMIDE_CreateNoteType", ObjNoteDialog, "AttachValueList", bAttachValueList)
	End If
	'Setting LOV
	If strLOV<>"" Then
		Call Fn_Edit_Box("Fn_BMIDE_CreateNoteType",ObjNoteDialog,"LOV",strLOV)
	End If
	'Setting Default Va;ue To Note Type
	If  strDefaultValue<>"" Then
		Call Fn_Edit_Box("Fn_BMIDE_CreateNoteType",ObjNoteDialog,"DefaultValue",strDefaultValue)
	End If
	'Clicking On Finish button 
	Call Fn_Button_Click("Fn_BMIDE_CreateNoteType", ObjNoteDialog, "Finish")
	'Function Returns True
	Fn_BMIDE_CreateNoteType=True
	'Releasing Object Of "NewNoteType" window
	Set ObjNoteDialog=Nothing
End Function

'-------------------------------------------------------------------Function Used to Create New Class----------------------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_CreateNewClass

'Description			 :	Function Used to Create New Class

'Parameters			   :	1.strName: Class Name
										'2.strParent: Class parent
										'3.bExportable:Exportable Option
										'4.bUnInheritable: UnInheritable Option
										'5.bUnInstantiable:unInstantiable Option
										'6.strAppName:Application Name
										'7.strAttributes:Attributes

'Return Value		   : 	True Or False

'Pre-requisite			:	New Class Dialog Should Be Opened

'Examples				: Call Fn_BMIDE_CreateNewClass("TestClass","AbsOccData","ON","","","","")
'									  Call Fn_BMIDE_CreateNewClass("TestClass1","AbsOccData","OFF","ON","OFF","","")
'									  strAttr="TestAttr::TestAttribute:String:32::OFF:::::"
'									 Call Fn_BMIDE_CreateNewClass("TestClass2","AbsOccData","","ON","OFF","",strAttr)
'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done								Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				22/12/2010			           1.0																								Sunny R
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'strAttributes : - strName:strDisplayName:strDescription:strAttributeType:strStringLength:strReferenceClass:chkSetInitialValueToNull:strInitialValue:strLowerBound:strUpperBound:chkArray:chkKeys~strName:strDisplayName:strDescription:strAttributeType:strStringLength:strReferenceClass:chkSetInitialValueToNull:strInitialValue:strLowerBound:strUpperBound:chkArray:chkKeys
'Main Set Separated By ( ~ ) Tilda And Sub Set Seperated By ( : ) Colan
Public Function Fn_BMIDE_CreateNewClass(strName,strParent,bExportable,bUnInheritable,bUnInstantiable,strAppName,strAttributes)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_CreateNewClass"
   Dim strPrefix,arrAttribute,iCounter,arrAttrValues
   Dim ObjClassDialog
	Fn_BMIDE_CreateNewClass=False
	'Checking Existance Of "NewClass" Dialog
   If Fn_UI_ObjectExist("Fn_BMIDE_CreateNewClass",JavaWindow("Business Modeler").JavaWindow("NewClass"))=True Then
	   'Creating Object Of "NewClass" Dialog
	   Set ObjClassDialog=JavaWindow("Business Modeler").JavaWindow("NewClass")
	Else
		'If "NewClass" Dialog is Not Exist Then Function Will Return Flase And Exit From Here
		Exit Function
   End If
   'Setting Name To The New Class
   If strName<>"" Then
	   'Taking Prefix Of Project
		strPrefix=Fn_Edit_Box_GetValue("Fn_BMIDE_CreateNewClass",ObjClassDialog,"Name")
		strName=strPrefix+strName
		Call Fn_Edit_Box("Fn_BMIDE_CreateNewClass",ObjClassDialog,"Name",strName)
   End If
   'Setting Parent To Class
   If strParent<>"" Then
	   Call Fn_Edit_Box("Fn_BMIDE_CreateNewClass",ObjClassDialog,"Parent",strParent)
   End If
   'Setting Status Of "Exportable" option
   If bExportable<>"" Then
	   Call Fn_CheckBox_Set("Fn_BMIDE_CreateNewClass", ObjClassDialog, "Exportable", bExportable)
   End If
   'Setting Status Of "Uninheritable" option
   If bUnInheritable<>"" Then
		Call Fn_CheckBox_Set("Fn_BMIDE_CreateNewClass", ObjClassDialog, "Uninheritable", bUnInheritable)
   End If
   'Setting Status Of "Uninstantiable" option
   If bUnInstantiable<>"" Then
		Call Fn_CheckBox_Set("Fn_BMIDE_CreateNewClass", ObjClassDialog, "Uninstantiable", bUnInstantiable)
   End If
   'Selecting Application Name
   If strAppName<>"" Then
	   Call Fn_List_Select("Fn_BMIDE_CreateNewClass", ObjClassDialog, "ApplicationName",strAppName)
   End If
   'Setting Attributes
   If strAttributes<>"" Then
		JavaWindow("Business Modeler").JavaWindow("NewCustomProperty").SetTOProperty "title","New Attribute"
		arrAttribute=Split(strAttributes,"~")
		For iCounter=0 To UBound(arrAttribute)
			arrAttrValues=Split(arrAttribute(iCounter),":")
			Call Fn_Button_Click("Fn_BMIDE_CreateNewClass", ObjClassDialog, "Add")
			Call Fn_BMIDE_CreateNewCustomProperties(arrAttrValues(0),arrAttrValues(1),arrAttrValues(2),arrAttrValues(3),arrAttrValues(4),arrAttrValues(5),arrAttrValues(6),arrAttrValues(7),arrAttrValues(8),arrAttrValues(9),arrAttrValues(10),arrAttrValues(11))
		Next
   End If
   'Clicking On Finish Button To Create New Class
	Call Fn_Button_Click("Fn_BMIDE_CreateNewClass", ObjClassDialog, "Finish")
	'Function Rreturns True
	Fn_BMIDE_CreateNewClass=True
	'Releasing Object "NewClass" Dialog
	Set ObjClassDialog=Nothing
End Function

'-------------------------------------------------------------------Function Used to Create New Service Library Or Service-------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_CreateServiceLibrary

'Description			 :	Function Used to Create New Service Library Or Service

'Parameters			   :	1.strName: Serivce Library Or Service Name
										'2.strDescription: Serivce Library Or Service Description
										'3.strDependantLibrary:Service Library Dependant Libraries

'Return Value		   : 	True Or False

'Pre-requisite			:	New Service Library Or New Service Dialog Should Be Opened

'Examples				: To Create Service Library Use Belove Examples
'									  Call Fn_BMIDE_CreateServiceLibrary("Test","Test Service","bmf:cfm:cm")
'									  Call Fn_BMIDE_CreateServiceLibrary("Test1","Test Service","cm")
'									  Call Fn_BMIDE_CreateServiceLibrary("Test2","Test Service","")
'									 To Create Service Use Belove Examples
'									 Call Fn_BMIDE_CreateServiceLibrary("TestService","Test Service","")
'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done								Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				22/12/2010			           1.0																								Sunny R
'													Priyanka B										   				17/1/2013			           1.1						While adding Description for service new GUI added of Description edit 																		Sunny R
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'By Using This Function User Can Create "Service Library" And "Service"
Public Function Fn_BMIDE_CreateServiceLibrary(strName,strDescription,strDependantLibrary)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_CreateServiceLibrary"
   Dim strPrefix,arrDepdLib,iCounter
   Dim ObjLibraryDialog
   Fn_BMIDE_CreateServiceLibrary=False
   If Fn_UI_ObjectExist("Fn_BMIDE_CreateServiceLibrary", JavaWindow("Business Modeler").JavaWindow("NewServiceLibrary"))=True Then
	   Set ObjLibraryDialog=JavaWindow("Business Modeler").JavaWindow("NewServiceLibrary")
	Else
		Exit Function
   End If
   If strName<>"" Then
		strPrefix=Fn_Edit_Box_GetValue("Fn_BMIDE_CreateServiceLibrary",ObjLibraryDialog,"Name")
		strName=strPrefix+strName
		Call Fn_Edit_Box("Fn_BMIDE_CreateServiceLibrary",ObjLibraryDialog,"Name",strName)
   End If
   If strDescription<>"" Then
		If ObjLibraryDialog.JavaEdit("Description").Exist(3) then
			Call Fn_Edit_Box("Fn_BMIDE_CreateServiceLibrary",ObjLibraryDialog,"Description",strDescription)
		Else
			Call Fn_Button_Click("Fn_BMIDE_CreateServiceLibrary", ObjLibraryDialog, "ServiceDescription")
			Call Fn_Edit_Box("Fn_BMIDE_CreateServiceLibrary",JavaWindow("Business Modeler").JavaWindow("DescriptionEditor"),"InterfaceDescription",strDescription)
			Call Fn_Button_Click("Fn_BMIDE_CreateServiceLibrary", JavaWindow("Business Modeler").JavaWindow("DescriptionEditor"), "Finish")
		End if
   End If
	If strDependantLibrary<>"" Then
		arrDepdLib=Split(strDependantLibrary,":")
		For iCounter=0 To UBound(arrDepdLib)
			Call Fn_Button_Click("Fn_BMIDE_CreateServiceLibrary", ObjLibraryDialog, "Add")
			Call Fn_Edit_Box("Fn_BMIDE_CreateServiceLibrary",JavaWindow("Business Modeler").JavaWindow("Find Business Object"),"Criteria",arrDepdLib(iCounter))
			Call Fn_Button_Click("Fn_BMIDE_CreateServiceLibrary", JavaWindow("Business Modeler").JavaWindow("Find Business Object"), "OK")
		Next
	End If
	Call Fn_Button_Click("Fn_BMIDE_CreateServiceLibrary", ObjLibraryDialog, "Finish")
	Fn_BMIDE_CreateServiceLibrary=True
	Set ObjLibraryDialog=Nothing
End Function

'-------------------------------------------------------------------Function Used to Perform Operations On Operations Tree Which appears in Operations Inner Tab-------------------------------------------------------
'Function Name		:	Fn_BMIDE_OperationsTreeOperation

'Description			 :	Function Used to Perform Operations On Operations Tree Which appears in Operations Inner Tab

'Parameters			   :	1.sAction: Action Name
										'2.sNodeName: Node Name Separated By ( ~ ) Tilda
										'3.sPopupMenu:PopUp Menu 

'Return Value		   : 	True Or False

'Pre-requisite			:	Should Be Login In BMIDE

'Examples				: Call Fn_BMIDE_CreateServiceLibrary("Select","Operations:Legacy Operations","")
'									  Call Fn_BMIDE_CreateServiceLibrary("Expand","Operations:Legacy Operations","")
'							
'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done								Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				22/12/2010			           1.0																								Sunny R
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_BMIDE_OperationsTreeOperation(sAction,sNodeName,sPopupMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_OperationsTreeOperation"
   'Activating Operations tab
   Call Fn_BMIDE_InnerTabOperations("Activate","Operations")
   'Function returns false
   Fn_BMIDE_OperationsTreeOperation = False
   Select Case sAction
	 	Case "Select" 'Case To Select Node
			Call Fn_SISW_UI_JavaTable_Operations("Fn_BMIDE_OperationsTreeOperation", "ActivateRow", JavaWindow("Business Modeler") , "OperationsTable", "", "", sNodeName, "", "", "", "")
			Fn_BMIDE_OperationsTreeOperation = True
		Case "Expand" 'Case To Expand Node
			Call Fn_UI_JavaTree_Expand("Fn_BMIDE_OperationsTreeOperation",JavaWindow("Business Modeler"),"OperationsTree",sNodeName)
			Fn_BMIDE_OperationsTreeOperation = True
   End Select
End Function

'-------------------------------------------------------------------Function Used to Add Extension Rule------------------------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_AddExtensionRule

'Description			 :	Function Used to Add Extension Rule

'Parameters			   :	1.strExtension: Extension
										'2.strArguments: Rule Argument
										'3.strCondition:Rule Condition

'Return Value		   : 	True Or False

'Pre-requisite			:	Should Be Log In BMIDE

'Examples				: Call Fn_BMIDE_AddExtensionRule("createObjects","Bitmap:CM_list~Briefcase:DMI_markup","isTrue")

'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done								Reviewer				  Tc Release
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   		22/12/2010			                1.0																		Sunny R
'													Ankit N													17/07/2015					     	1.1              Selected of "Extension Rule" Via "Browse"            	ViVek A.         		Tc11.2_2015070100
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_BMIDE_AddExtensionRule(strExtension,strArguments,strCondition)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_AddExtensionRule"
   Dim arrArgument,iCounter,arrArgVal
   Dim ObjRuleDialog,objFindCondition
   Fn_BMIDE_AddExtensionRule=False
   Set ObjRuleDialog=JavaWindow("Business Modeler").JavaWindow("AddExtensionRule")
   Set objFindCondition = JavaWindow("Business Modeler").JavaWindow("FindOperationCondition")
	If strExtension<>"" Then
		'Call Fn_Edit_Box("Fn_BMIDE_AddExtensionRule",ObjRuleDialog,"Extension",strExtension)					''Tc112-2015070100-17_07_2015-Porting-NitishB-Modified Selection of "Extension Rule" Via "Browse"
		'Call Fn_KeyBoardOperation("SendKeys","{ENTER}")
		Call Fn_Button_Click("Fn_BMIDE_AddExtensionRule", ObjRuleDialog, "BrowseExtension")
		Call Fn_Edit_Box("Fn_BMIDE_AddExtensionRule",JavaWindow("Business Modeler").JavaWindow("Find Business Object"),"Project",strExtension)
		Call Fn_Button_Click("Fn_BMIDE_AddExtensionRule", JavaWindow("Business Modeler").JavaWindow("Find Business Object"), "OK")
	End If
	
	If strArguments<>"" Then
		arrArgument=Split(strArguments,"~")
		For iCounter=0 To Ubound(arrArgument)
			arrArgVal=Split(arrArgument(iCounter),":")
			
			'Call Fn_Button_Click("Fn_BMIDE_AddExtensionRule", ObjRuleDialog, "BrowseExtension")
		'Call Fn_Edit_Box("Fn_BMIDE_AddExtensionRule",JavaWindow("Business Modeler").JavaWindow("Find Business Object"),"Project",strExtension)
		'Call Fn_Button_Click("Fn_BMIDE_AddExtensionRule", JavaWindow("Business Modeler").JavaWindow("Find Business Object"), "OK")
			
			
			Call Fn_Button_Click("Fn_BMIDE_AddExtensionRule", ObjRuleDialog, "Add")
			If arrArgVal(0)<>"" Then
				Call Fn_Edit_Box("Fn_BMIDE_AddExtensionRule",JavaWindow("Business Modeler").JavaWindow("NewArgument"),"ObjectType",arrArgVal(0))
				Call Fn_KeyBoardOperation("SendKeys","{ENTER}")
			End If
			If arrArgVal(1)<>"" Then
				Call Fn_Edit_Box("Fn_BMIDE_AddExtensionRule",JavaWindow("Business Modeler").JavaWindow("NewArgument"),"RelationType",arrArgVal(1))
				Call Fn_KeyBoardOperation("SendKeys","{ENTER}")
			End If
			Call Fn_Button_Click("Fn_BMIDE_AddExtensionRule", JavaWindow("Business Modeler").JavaWindow("NewArgument"), "Finish")
		Next
	End If
	If strCondition<>"" Then
		'Call Fn_SISW_UI_JavaEdit_Operations("Fn_BMIDE_AddExtensionRule", "Set", ObjRuleDialog, "Condition", strCondition)
		Call Fn_Button_Click("", ObjRuleDialog,"BrowseCond")
		Call Fn_SISW_UI_JavaEdit_Operations("", "Set", objFindCondition, "ConditionField", strCondition)
		Call Fn_Button_Click("", objFindCondition,"OK")
	End If
	Call Fn_Button_Click("Fn_BMIDE_AddExtensionRule", ObjRuleDialog, "Finish")
	Fn_BMIDE_AddExtensionRule=True
	Set ObjRuleDialog=Nothing
	Set objFindCondition=Nothing
End Function
'-------------------------------------------------------------------Function Used to Perform Operation On Pre-Condition Table-----------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_PreConditionTableOperations

'Description			 :	Function Used to Perform Operation On Pre-Condition Table

'Parameters			   :  1.strAction : Action Name
'										2.strExtensn: Extension
										'3.Argument: Rule Argument
										'4.strCond:Rule Condition

'Return Value		   : 	True Or False

'Pre-requisite			:	Should Be Log In BMIDE

'Examples				: Call Fn_BMIDE_PreConditionTableOperations("Add","createObjects","Bitmap:CM_list~Briefcase:DMI_markup","isTrue")

'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done								Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				22/12/2010			           1.0																								Sunny R
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_BMIDE_PreConditionTableOperations(strAction,strExtensn,Argument,strCond)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_PreConditionTableOperations"
   Dim bReturn
   Fn_BMIDE_PreConditionTableOperations=False
   Call Fn_UI_JavaTab_Select("",JavaWindow("Business Modeler"),"MainInnerTab2", "Extension Attachments")
   Select Case strAction
	 	Case "Add" 'Case To Add "ExtensionRule"
       		Call Fn_JavaTree_Select("Fn_BMIDE_PreConditionTableOperations", JavaWindow("Business Modeler"), "OperationsTree","Pre-Condition")
			Call Fn_Button_Click("Fn_BMIDE_PreConditionTableOperations", JavaWindow("Business Modeler"), "AddCondition")
			bReturn=Fn_BMIDE_AddExtensionRule(strExtensn,Argument,strCond)
			If bReturn=True Then
				Fn_BMIDE_PreConditionTableOperations=True
			End If
			
   End Select
End Function

'-------------------------------------------------------------------Function Used to Perform Operation On Pre-Condition Table-----------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_PreConditionTableOperations

'Description			 :	Function Used to Perform Operation On Pre-Condition Table

'Parameters			   :  1.strAction : Action Name
'										2.strExtensn: Extension
										'3.Argument: Rule Argument
										'4.strCond:Rule Condition

'Return Value		   : 	True Or False

'Pre-requisite			:	Should Be Log In BMIDE

'Examples				: Call Fn_BMIDE_PostActionTableOperations("Add","createObjects","Bitmap:CM_list~Briefcase:DMI_markup","isTrue")
'									  Call Fn_BMIDE_PostActionTableOperations("Verify","createObjects","","")
'									Call Fn_BMIDE_PostActionTableOperations("VerifyActive","autoAssignToProject","true","")
'									Call Fn_BMIDE_PostActionTableOperations("VerifyTemplate","autoAssignToProject","regliveupdatebmidetemplate","")

'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done										Reviewer				Tc Release
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   		  22/12/2010			           1.0																				Sunny R
'													Pranav Ingle										   	  15/12/2011			           1.1						Added Case "VerifyTemplate"						    	Sandeep
'													Ankit N													  17/07/2015					   1.2				Modified Case "VerifyActive" and "VerifyTemplate"				ViVek A. 				Tc11.2_2015070100			
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_BMIDE_PostActionTableOperations(strAction,strExtensn,Argument,strCond)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_PostActionTableOperations"
   Dim bReturn,iRowCount,iCounter,strData,strActData
   Fn_BMIDE_PostActionTableOperations=False
   Call Fn_UI_JavaTab_Select("",JavaWindow("Business Modeler"),"MainInnerTab2", "Extension Attachments")
   Select Case strAction
	 	Case "Add" 'Case To Add "ExtensionRule"
       		Call Fn_JavaTree_Select("Fn_BMIDE_PreConditionTableOperations", JavaWindow("Business Modeler"), "OperationsTree","Post-Action")
			Call Fn_Button_Click("Fn_BMIDE_PreConditionTableOperations", JavaWindow("Business Modeler"), "AddCondition")
			bReturn=Fn_BMIDE_AddExtensionRule(strExtensn,Argument,strCond)
			If bReturn=True Then
				Fn_BMIDE_PostActionTableOperations=True
				Exit Function
			End If
		Case "Verify"
			iRowCount=Fn_UI_Object_GetROProperty("Fn_BMIDE_PostActionTableOperations",JavaWindow("Business Modeler").JavaTree("OperationsTree"), "items count")
			For iCounter=0 To iRowCount-1
					strData=JavaWindow("Business Modeler").JavaTree("OperationsTree").GetItem(iCounter)
					If Trim(strData)=Trim(strExtensn) Then
						Fn_BMIDE_PostActionTableOperations=True
						Exit Function
					End If
			Next
		Case "VerifyActive"											'Tc112-2015070100-17_07_2015-Porting-NitishB-Modified case as JavaTable("PostAction") changed to JavaTree("OperationsTree") as per design change
			sTreepath = Fn_UI_JavaTreeGetItemPathExt("", JavaWindow("Business Modeler").JavaTree("OperationsTree"), strExtensn, "~", "")
			sTreepath = Replace(sTreepath,"#","")
			arrpath = split(sTreepath,":")
			strActData=JavaWindow("Business Modeler").JavaTree("OperationsTree").object.GetItem(arrpath(0)).GetItem(arrpath(1)).getData().getExtAttach().getActive()
			If Trim(strActData)=Trim(Argument) Then
				Fn_BMIDE_PostActionTableOperations=True
				Exit Function
			End If
			
		Case "VerifyTemplate"										'Tc112-2015070100-17_07_2015-Porting-NitishB-Modified case as JavaTable("PostAction") changed to JavaTree("OperationsTree") as per design change
			iRowCount=Fn_UI_Object_GetROProperty("Fn_BMIDE_PostActionTableOperations",JavaWindow("Business Modeler").JavaTree("OperationsTree"), "items count")
			For iCounter=0 To iRowCount-1
					strData=JavaWindow("Business Modeler").JavaTree("OperationsTree").GetItem(iCounter)
					If Trim(strData)=Trim(strExtensn) Then
						strActData=JavaWindow("Business Modeler").JavaTree("OperationsTree").GetColumnValue(strData,"Template")
						If Trim(strActData)=Trim(Argument) Then
							Fn_BMIDE_PostActionTableOperations=True
							Exit For
						End If
					End If
			Next
   End Select
End Function 

'-------------------------------------------------------------------Function Used to Perform Operation On Business Object Constant Table---------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_BusinessObjectConstantTableOperation

'Description			 :	Function Used to Perform Operation On Business Object Constant Table

'Parameters			   :  1.strAction : Action Name
'										2.strName:Constant Name
										'3.strType: Type
										'4.strValue:Constant Value

'Return Value		   : 	True Or False

'Pre-requisite			:	Should Be Log In BMIDE

'Examples				: Call Fn_BMIDE_BusinessObjectConstantTableOperation("Edit","Fnd0MarkupControlObject","","ON")
'									  Call  Fn_BMIDE_BusinessObjectConstantTableOperation("Edit","RenderProviderName","","SQS")
'									 Call 'Fn_BMIDE_BusinessObjectConstantTableOperation("VerifyValue","BatchPrintProviderName","","SIEMENS")
'									Call Fn_BMIDE_BusinessObjectConstantTableOperation("Select","Fnd0MarkupControlObject","","")
'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done								Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				22/12/2010			           1.0																								Sunny R
'													Sandeep N										   				23/12/2010			           1.0							Case "VerifyValue"								Sunny R
'													Sandeep N										   				18/02/2011			           1.0							Case "Select"								Sunny R
'													Sandeep N										   				28/03/2011			           1.0							Modify Case  "Edit" For [ List	]							Sunny R
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_BMIDE_BusinessObjectConstantTableOperation(strAction,strName,strType,strValue)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_BusinessObjectConstantTableOperation"
	Dim iRowCount,iCounter,strConstName,bFlag,currType,currValue
	Dim ObjModelerDialog,ObjConstDialog
	Dim WshShell
	bFlag=False
	Fn_BMIDE_BusinessObjectConstantTableOperation=False
	'Activating Main Tab
		Call Fn_BMIDE_InnerTabOperations("Activate","Main")
		Set ObjModelerDialog=JavaWindow("Business Modeler")
		Select Case strAction
				Case "Edit"
					iRowCount=Fn_UI_Object_GetROProperty("Fn_BMIDE_BusinessObjectConstantTableOperation",ObjModelerDialog.JavaTable("BusinessObjectConstants"),"rows")
					For iCounter=0 To iRowCount-1
							strConstName=JavaWindow("Business Modeler").JavaTable("BusinessObjectConstants").GetCellData(iCounter,"Name")	
							If Trim(strConstName)=Trim(strName) Then
								JavaWindow("Business Modeler").JavaTable("BusinessObjectConstants").SelectCell iCounter,0
'								JavaWindow("Business Modeler").JavaTable("BusinessObjectConstants").ActivateRow iCounter
								If ObjModelerDialog.JavaButton("EditBusinessObjConst").GetROProperty("enabled") <> "1" Then
									JavaWindow("Business Modeler").JavaTable("BusinessObjectConstants").SelectCell 0,0
									For iCount=0 To iCounter-1
										Set WshShell = CreateObject("WScript.Shell")
										WshShell.SendKeys "{DOWN}"
										wait 0, 500
									Next
									Set WshShell =Nothing
								End If

								bFlag=True
								Exit For
							End If
					Next
					If bFlag=False Then
						Exit Function
					End If
					Call Fn_Button_Click("Fn_BMIDE_BusinessObjectConstantTableOperation", ObjModelerDialog, "EditBusinessObjConst")
					Set ObjConstDialog=JavaWindow("Business Modeler").JavaWindow("BusinessObjectConstant")
					If strValue<>"" Then
						currType=Fn_Edit_Box_GetValue("Fn_BMIDE_BusinessObjectConstantTableOperation",ObjConstDialog,"Type")
						If currType="String" Then
							'Call Fn_Edit_Box("Fn_BMIDE_BusinessObjectConstantTableOperation",ObjConstDialog,"Value",strValue)
							Call Fn_SISW_UI_JavaEdit_Operations("Fn_BMIDE_BusinessObjectConstantTableOperation", "SetExt", ObjConstDialog, "Value", strValue)
						End If
						If currType="Boolean" Then
							'Call Fn_CheckBox_Set("Fn_BMIDE_BusinessObjectConstantTableOperation", ObjConstDialog, "Value", strValue)
							If strValue<>"" Then
								Select Case lcase(strValue)
			     					Case "on"
			     						Call Fn_UI_Object_SetTOProperty_ExistCheck("",ObjConstDialog.JavaRadioButton("Value"),"attached text","True")
										Call Fn_UI_JavaRadioButton_SetON("Fn_BMIDE_GlobalConstantTableOperations", ObjConstDialog, "Value")
									Case "off"
										Call Fn_UI_Object_SetTOProperty_ExistCheck("",ObjConstDialog.JavaRadioButton("Value"),"attached text","False")
										Call Fn_UI_JavaRadioButton_SetON("Fn_BMIDE_GlobalConstantTableOperations", ObjConstDialog, "Value")
								End Select
						   End If
						End If
						If currType="List" Then
                            Call Fn_List_Select("Fn_BMIDE_BusinessObjectConstantTableOperation", ObjConstDialog, "ValueList",strValue)
						End If

					End If
					Call Fn_Button_Click("Fn_BMIDE_BusinessObjectConstantTableOperation", ObjConstDialog, "Finish")
					Fn_BMIDE_BusinessObjectConstantTableOperation=True
					Set ObjConstDialog=Nothing

				Case "VerifyValue"
					iRowCount=Fn_UI_Object_GetROProperty("Fn_BMIDE_BusinessObjectConstantTableOperation",ObjModelerDialog.JavaTable("BusinessObjectConstants"),"rows")
					For iCounter=0 To iRowCount-1
							strConstName=JavaWindow("Business Modeler").JavaTable("BusinessObjectConstants").GetCellData(iCounter,"Name")	
							If Trim(strConstName)=Trim(strName) Then
								currValue=JavaWindow("Business Modeler").JavaTable("BusinessObjectConstants").GetCellData(iCounter,"Value")	
								If  Trim(currValue)=Trim(strValue)Then
									bFlag=True
									Exit For
								End If
							End If
					Next
					If bFlag=True Then
						Fn_BMIDE_BusinessObjectConstantTableOperation=True
					End If

				Case "Select"
					iRowCount=Fn_UI_Object_GetROProperty("Fn_BMIDE_BusinessObjectConstantTableOperation",ObjModelerDialog.JavaTable("BusinessObjectConstants"),"rows")
					For iCounter=0 To iRowCount-1
							strConstName=JavaWindow("Business Modeler").JavaTable("BusinessObjectConstants").GetCellData(iCounter,"Name")	
							If Trim(strConstName)=Trim(strName) Then
                                JavaWindow("Business Modeler").JavaTable("BusinessObjectConstants").SelectCell iCounter,0
'								JavaWindow("Business Modeler").JavaTable("BusinessObjectConstants").ActivateRow iCounter
								bFlag=True
								Exit For
							End If
					Next
					If bFlag=False Then
						Exit Function
					End If
					Fn_BMIDE_BusinessObjectConstantTableOperation=True
		End Select
		Set ObjModelerDialog=Nothing
End Function 

'-------------------------------------------------------------------Function Used to Restart workspace -------------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_RestartWorkspace

'Description			 :	Function Used to restart workspace 

'Return Value		   : 	True Or False

'Pre-requisite			:	

'Examples				: 	'Call Fn_BMIDE_RestartWorkspace()

'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done																		Reviewer
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Pranav  Ingle									   				23/12/2010			           1.0																																															Sunny R
'													Sandeep N									   				25/08/2011			           1.1						Added code to handle [ TemplateProjectBackup ] window		 Sunny R
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_BMIDE_RestartWorkspace()
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_RestartWorkspace"
	Dim strMenu, iCount, iCounter
    strMenu=Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\BMIDEConfig\BMIDE_Menu.xml", "RestartWorkspace")
	Call Fn_BMIDE_MenuOperation("Select", strMenu)
	'- - - - - To Handle [ TemplateProjectBackup ] window : Added By Sandeep - - - - - - - - - - - - - - 
    If JavaWindow("Business Modeler").JavaWindow("TemplateProjectBackup").Exist(6) Then
		  Call Fn_Button_Click("Fn_BMIDE_RestartWorkspace", JavaWindow("Business Modeler").JavaWindow("TemplateProjectBackup"), "Finish")
		  JavaWindow("Business Modeler").JavaWindow("TemplateProjectBackup").SetTOProperty "index",1
		  wait 1
		 If JavaWindow("Business Modeler").JavaWindow("TemplateProjectBackup").Exist(6) Then
			 Call Fn_Button_Click("Fn_BMIDE_RestartWorkspace", JavaWindow("Business Modeler").JavaWindow("TemplateProjectBackup"), "OK")
		 End If
		 JavaWindow("Business Modeler").JavaWindow("TemplateProjectBackup").SetTOProperty "index",0
	End If
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	For iCount = 1 To 3
		If JavaWindow("Eclipse Launcher").Exist(iTime) Then
			JavaWindow("Eclipse Launcher").JavaButton("OK").Click			
		ElseIf Dialog("WorkspaceLauncher").Exist(iTime)=True Then
			Dialog("WorkspaceLauncher").WinButton("OK").Click
		End If
		For iCounter = 1 To 3
			If JavaWindow("Business Modeler").Exist(iTime) = True Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Pass:Business Modular Window Exist")
				Exit For 
			End If
		Next
		Exit For
	Next

	Fn_BMIDE_RestartWorkspace=True
End Function
'-------------------------------------------------------------------Function Used to Perform Operation On Attributes Table-----------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_AttributeTableOperatons

'Description			 :	Function Used to Perform Operation On Attributes Table

'Parameters			   :  1.strAction : Action Name
'										2.strAttributes:Attributes
										'3.strColName: Column Name
										'4.strValue:Value

'Return Value		   : 	True Or False

'Pre-requisite			:	Should Be Log In BMIDE

'Examples				: Call Fn_BMIDE_AttributeTableOperatons("Verify","","Attribute Name","ir_ref")
'									  Call  Fn_BMIDE_AttributeTableOperatons("Remove","bl_name","","")
'									'Use Belove Example For "Edit" Case
'									"strName:strDisplayName:strDescription:strAttributeType:strStringLength:strReferenceClass:chkSetInitialValueToNull:strInitialValue:strLowerBound:strUpperBound:chkArray:chkKeys"
'									sAttrbt="s3Test::NewDescription::64:::1::::"
'									Call Fn_BMIDE_AttributeTableOperatons("Edit",sAttrbt,"","")
'																													 'StrAction,"Attribute Name","Column Name","strVue"
'									Call Fn_BMIDE_AttributeTableOperatons("VerifyValues","s3Test","Storage Type","String[64]")

'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done								Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				23/12/2010			           1.0																								Sunny R
'													Sandeep N										   				23/12/2010			           1.0						Case "Remove"	        							Sunny R
'													Sandeep N										   				23/12/2010			           1.0						Case "Edit"	        										Sunny R
'													Sandeep N										   				23/12/2010			           1.0						Case "VerifyValues"	        					  Sunny R
'													Pallavi J										   				27/05/2013			           1.0						Case "Select"	        					  Sandeep N
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_BMIDE_AttributeTableOperatons(strAction,strAttributes,strColName,strValue)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_AttributeTableOperatons"
   Dim iRowCount,iCounter,currValue,bFlag,arrAttributes,strAttrName
   bFlag=False
   Fn_BMIDE_AttributeTableOperatons=False
   Select Case strAction
		 	Case "Verify"
				iRowCount=JavaWindow("Business Modeler").JavaTable("AttributesTable").GetROProperty("rows")
				For iCounter=0 To iRowCount-1
						'currValue=JavaWindow("Business Modeler").JavaTable("AttributesTable").GetCellData(iCounter,strColName)
						currValue=Fn_UI_JavaTable_GetCellData("Fn_BMIDE_AttributeTableOperatons", JavaWindow("Business Modeler"), "AttributesTable",iCounter,strColName)
						If Trim(currValue)=Trim(strValue) Then
							bFlag=True
							Exit For
						End If
				Next
				If bFlag=True Then
					Fn_BMIDE_AttributeTableOperatons=True
				End If
			Case "Remove"
				iRowCount=JavaWindow("Business Modeler").JavaTable("AttributesTable").GetROProperty("rows")
				For iCounter=0 To iRowCount-1
						currValue=JavaWindow("Business Modeler").JavaTable("AttributesTable").GetCellData(iCounter,"Attribute Name")
						If Trim(currValue)=Trim(strAttributes) Then
							JavaWindow("Business Modeler").JavaTable("AttributesTable").SelectCell iCounter,0
							JavaWindow("Business Modeler").JavaButton("RemoveAttribute").Click
							Call Fn_BMIDE_DeleteObject()
							bFlag=True
							Exit For
						End If
				Next
				If bFlag=True Then
					Fn_BMIDE_AttributeTableOperatons=True
				End If
			Case "Edit"
				arrAttributes=Split(strAttributes,":")
				iRowCount=JavaWindow("Business Modeler").JavaTable("AttributesTable").GetROProperty("rows")
				For iCounter=0 To iRowCount-1
						currValue=JavaWindow("Business Modeler").JavaTable("AttributesTable").GetCellData(iCounter,"Attribute Name")
						If Trim(currValue)=Trim(arrAttributes(0)) Then
							JavaWindow("Business Modeler").JavaTable("AttributesTable").SelectCell iCounter,0
							JavaWindow("Business Modeler").JavaButton("EditAttributes").Click
							bFlag=True
							Exit For
						End If
				Next
				If bFlag=False Then
					Exit Function
				End If
				 JavaWindow("Business Modeler").JavaWindow("NewCustomProperty").SetTOProperty "title","Modify Attribute"
				Call Fn_BMIDE_CreateNewCustomProperties("",arrAttributes(1),arrAttributes(2),arrAttributes(3),arrAttributes(4),arrAttributes(5),arrAttributes(6),arrAttributes(7),arrAttributes(8),arrAttributes(9),arrAttributes(10),arrAttributes(11))
				JavaWindow("Business Modeler").JavaWindow("NewCustomProperty").SetTOProperty "title","New Property"
				Fn_BMIDE_AttributeTableOperatons=True

			Case "VerifyValues"
				'iRowCount=JavaWindow("Business Modeler").JavaTable("AttributesTable").GetROProperty("rows")
				iRowCount = Fn_UI_Object_GetROProperty("Fn_BMIDE_AttributeTableOperatons", JavaWindow("Business Modeler").JavaTable("AttributesTable"),"rows")
				For iCounter=0 To iRowCount-1
						'strAttrName=JavaWindow("Business Modeler").JavaTable("AttributesTable").GetCellData(iCounter,"Attribute Name")
						strAttrName=Fn_UI_JavaTable_GetCellData("Fn_BMIDE_AttributeTableOperatons", JavaWindow("Business Modeler"), "AttributesTable",iCounter,"Attribute Name")
						If Trim(strAttrName)=Trim(strAttributes) Then
							currValue=JavaWindow("Business Modeler").JavaTable("AttributesTable").GetCellData(iCounter,strColName)
							If  Trim(currValue)=Trim(strValue) Then
								bFlag=True
								Exit For
							End If
						End If
				Next
				If bFlag=True Then
					Fn_BMIDE_AttributeTableOperatons=True
				End If
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			Case "Select"
				iRowCount=JavaWindow("Business Modeler").JavaTable("AttributesTable").GetROProperty("rows")
				For iCounter = 0 To iRowCount-1
						currValue=JavaWindow("Business Modeler").JavaTable("AttributesTable").GetCellData(iCounter,"Attribute Name")
						If Trim(currValue)=Trim(strAttributes) Then
							JavaWindow("Business Modeler").JavaTable("AttributesTable").SelectCell iCounter,0
                            bFlag=True
							Exit For
						End If
				Next
				If bFlag=True Then
					Fn_BMIDE_AttributeTableOperatons=True
				End If
   End Select
End Function

'-------------------------------------------------------------------Function Used to Create Package -------------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_CreatePackage

'Description			 :	Function Used to Create Package

'Return Value		   : 	True Or False

'Pre-requisite			:	

'Examples				: 	'Call Fn_BMIDE_CreatePackage("Trial")

'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sunny Ruparel								   				23/12/2010			           1.0																						Sunny R
'													Sandeep N								   				28/09/2011				           1.1																						Sunny R
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_BMIDE_CreatePackage(StrProjectName)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_CreatePackage"
	Dim strMenu,ObjPackageWnd

		Fn_BMIDE_CreatePackage = False
		strMenu=Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\BMIDEConfig\BMIDE_Menu.xml", "FileNewOther")
		Call Fn_BMIDE_MenuOperation("Select", strMenu)
		'Expanding ""Business Modeler IDE"" Node From "WizardsTree"
		Call Fn_UI_JavaTree_Expand("Fn_BMIDE_CreatePackage",JavaWindow("Business Modeler").JavaWindow("NewProject"),"WizardsTree","Business Modeler IDE")
		'Selecting "Business Modeler IDE:Package Template Extensions"" node 
		Call Fn_JavaTree_Select("Fn_BMIDE_CreatePackage",JavaWindow("Business Modeler").JavaWindow("NewProject"),"WizardsTree","Business Modeler IDE:Generate Software Package")
		'Clicking "Next" Button To open "Package Template Extensions" Dialog
		Call Fn_Button_Click("Fn_BMIDE_CreatePackage", JavaWindow("Business Modeler").JavaWindow("NewProject"), "Next")
		'Creating Object Of Package Template Extensions Window
		Set ObjPackageWnd=Fn_UI_ObjectCreate("Fn_BMIDE_CreatePackage", JavaWindow("Business Modeler").JavaWindow("PackageTemplateExtensions"))
		'Added Code to Handle [ Recomondation] dialog
		If ObjPackageWnd.JavaCheckBox("BackupDuringShutdownOfBMIDE").Exist(4) Then
			Call Fn_Button_Click("Fn_BMIDE_DeployProject", ObjPackageWnd, "Next")
		End If
		If StrProjectName<>"" Then
			'Selecting Project For Package Template Extensions
			Call Fn_List_Select("Fn_BMIDE_CreatePackage", ObjPackageWnd,"Project",StrProjectName)
		End If
		ObjPackageWnd.JavaButton("Finish").WaitProperty "enabled","1",iTime
		'Clicking "Finish" Button
		Call Fn_Button_Click("Fn_BMIDE_CreatePackage",ObjPackageWnd, "Finish")
		If Fn_UI_ObjectExist("Fn_BMIDE_CreatePackage",ObjPackageWnd) = True Then
			wait(20)
		Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Created Package from project [" + StrProjectName + "]")
			Fn_BMIDE_CreatePackage = True
		End If
		Set ObjPackageWnd=Nothing
End Function

'-------------------------------------------------------------------Function Used to Create Template using TEM.bat -------------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_CreateTemplate

'Description			 :	Function Used to Create Template using TEM.bat

'Return Value		   : 	True Or False

'Pre-requisite			:	

'Examples				: 	'Call Fn_BMIDE_CreateTemplate("Trial")

'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sunny Ruparel								   				23/12/2010			           1.0																						Sunny R
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_BMIDE_CreateTemplate(StrProjectName,strProjectPath)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_CreateTemplate"
   Dim sFolderPath,arrPath,pathLength,i,sRootPath,sPath,objShell,sWorkspacePath,sStringToType,objFSO,objStartFolder,objFolder,colFiles

	sFolderPath = Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\EnvVar_Ext.xml", "AppExecutable")

	arrPath = Split(sFolderPath,"\")
	pathLength = UBound(arrPath)-2
	For i=0 to pathLength
		sRootPath = sRootPath+arrPath(i)+"\"
	Next
	sPath = sRootPath +"install\tem.bat"
'	Set objShell = CreateObject("WScript.Shell")
'	'Invoking TEM.bat 
'	objShell.Run sPath
'	Set objShell = Nothing
	SystemUtil.Run sPath

	For i=0 to 5
		If Fn_UI_ObjectExist("Fn_BMIDE_CreateTemplate",JavaWindow("TEM")) = True Then
			Exit For
		Else
			wait(20)
		End If
	Next
	
	JavaWindow("TEM").JavaRadioButton("Maintenance").SetTOProperty "attached text","Configuration Manager"
	JavaWindow("TEM").JavaRadioButton("Maintenance").Set "ON"
	Call Fn_Button_Click("Fn_BMIDE_CreateTemplate",JavaWindow("TEM"),"Next")
	
	JavaWindow("TEM").JavaRadioButton("Maintenance").SetTOProperty "attached text", "Perform maintenance on an existing configuration"
	JavaWindow("TEM").JavaRadioButton("Maintenance").Set "ON"
	Call Fn_Button_Click("Fn_BMIDE_CreateTemplate",JavaWindow("TEM"),"Next")
	
	Call Fn_Button_Click("Fn_BMIDE_CreateTemplate",JavaWindow("TEM"),"Next")
	'Selecting option for Creating Template from Project
	JavaWindow("TEM").JavaRadioButton("Maintenance").SetTOProperty "attached text", ".*Add/Update Templates for working within the Business Modeler IDE Client"
	JavaWindow("TEM").JavaRadioButton("Maintenance").Set "ON"
	Call Fn_Button_Click("Fn_BMIDE_CreateTemplate",JavaWindow("TEM"),"Next")
	wait 2
	If JavaWindow("TEM").JavaButton("Next").GetROProperty("enabled")=0 Then
		JavaWindow("TEM").JavaTable("OriginalMediaLocation").ActivateRow "#0"
		JavaWindow("TEM").JavaTable("OriginalMediaLocation").SetCellData "#0","Update Location","Z:\install_kit"
	End If

	Call Fn_Button_Click("Fn_BMIDE_CreateTemplate",JavaWindow("TEM"),"Next")
	Call Fn_Button_Click("Fn_BMIDE_CreateTemplate",JavaWindow("TEM"),"Add")
	'Specify the packaged path
	
	Call Fn_Button_Click("Fn_BMIDE_CreateTemplate",JavaWindow("TEM").JavaDialog("SelectTemplates"),"Browse")
	sWorkspacePath = Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\BMIDEConfig\BMIDE_Config.xml", "BMIDEWorkspacePath")
	If Fn_UI_ObjectExist("Fn_BMIDE_CreateTemplate",JavaWindow("TEM").JavaDialog("Open")) = True Then
	
		'sStringToType = sWorkspacePath+"\"+StrProjectName+"\output\wntx64\packaging\full_update\feature_"+lcase(StrProjectName)+".xml"
	Call Fn_UI_EditBox_Type("Fn_BMIDE_CreateTemplate",JavaWindow("TEM").JavaDialog("Open"),"File name",strProjectPath)
	Call Fn_Button_Click("Fn_BMIDE_CreateTemplate",JavaWindow("TEM").JavaDialog("Open"),"Open")
	End If
	Call Fn_UI_EditBox_Type("Fn_BMIDE_CreateTemplate",JavaWindow("TEM").JavaDialog("SelectTemplates"),"Search",lcase(StrProjectName))
	Call Fn_SISW_UI_JavaList_Operations("Fn_BMIDE_CreateTemplate", "Select", JavaWindow("TEM").JavaDialog("SelectTemplates"), "Search","#0", "", "")
	Call Fn_Button_Click("Fn_BMIDE_CreateTemplate",JavaWindow("TEM").JavaDialog("SelectTemplates"),"OK")
	
	Call Fn_Button_Click("Fn_BMIDE_CreateTemplate",JavaWindow("TEM"),"Next")
	
	Call Fn_Button_Click("Fn_BMIDE_CreateTemplate",JavaWindow("TEM"),"Start")
	
	For i=0 to 5
		If Fn_UI_ObjectExist("Fn_BMIDE_CreateTemplate",JavaWindow("TEM").JavaButton("Restart")) = True Then
			Exit For
		Else
			wait(20)
		End If
	Next
	'Exiting TEM.bat
	Call Fn_Button_Click("Fn_BMIDE_CreateTemplate",JavaWindow("TEM"),"Close")
	Fn_BMIDE_CreateTemplate = True
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Created Template from project [" + StrProjectName + "] using TEM.bat")

End Function

'-------------------------------------------------------------------Function Used to Click ToolBar Button------------------- -----------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_ToolbarButtonClick

'Description			 :	Function Used to Click ToolBar Button

'Parameters			   :  1.iInstance : Instance Number
'										2.sButtonName:Button Name

'Return Value		   : 	True Or False

'Pre-requisite			:	Should Be Log In BMIDE

'Examples				: Call Fn_BMIDE_ToolbarButtonClick("","Find Class...")
'									  Call  Fn_BMIDE_ToolbarButtonClick(1,"Find Class...")

'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done								Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				23/12/2010			           1.0																								Sunny R
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_BMIDE_ToolbarButtonClick(iInstance,sButtonName)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_ToolbarButtonClick"
	Dim ObjDesc, ArrLists, iToolCnt, iCounter, sContents, iCnt

	If JavaWindow("Business Modeler").Exist(20) Then
		'Create Toolbar object
		Set ObjDesc = Description.Create() 
		ObjDesc("to_class").Value = "JavaToolbar" 
		ObjDesc("enabled").Value = 1
	
		JavaWindow("Business Modeler").Maximize
		
		'Get the total of Toolbar objects
		Set ArrLists =JavaWindow("Business Modeler").ChildObjects(ObjDesc)
		iToolCnt = JavaWindow("Business Modeler").ChildObjects(ObjDesc).count
		iCnt =1
		For iCounter = 0 to iToolCnt-1
			sContents = ArrLists(iCounter).GetContent()
			If instr(sContents, sButtonName) > 0 Then	
				If iInstance<>"" Then
					If iCnt = CInt(iInstance) Then
							ArrLists(iCounter).Press sButtonName
							Fn_BMIDE_ToolbarButtonClick = TRUE
							Exit For
					End If
					iCnt = iCnt +1
				Else
							ArrLists(iCounter).Press sButtonName
							Fn_BMIDE_ToolbarButtonClick = TRUE
							Exit For
				End If
			
			End If
		Next
	
		If iCounter = iToolCnt Then
			Fn_BMIDE_ToolbarButtonClick = FALSE
		End If

		Set ObjDesc = Nothing
		Set ArrLists = Nothing
	Else
		Fn_BMIDE_ToolbarButtonClick = FALSE
	End If
End Function
'-------------------------------------------------------------------Function Used to Search Objects---------------------------------- -----------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_FindObjects

'Description			 :	Function Used to Search Objects

'Parameters			   :  1.strProjectName : Project Name
'										2.strObjectName:Object Name

'Return Value		   : 	True Or False

'Pre-requisite			:	Find Object Dialog Should Be Appear On Screen

'Examples				: 	Fn_BMIDE_ToolbarButtonClick("","Find Class...")
'										Call Fn_BMIDE_FindObjects("DemoBatchProject","S3TestClass")

'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done								Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				23/12/2010			           1.0																								Sunny R
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_BMIDE_FindObjects(strProjectName,strObjectName)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_FindObjects"
   Dim ObjBusnssDialog
	Fn_BMIDE_FindObjects=False
	Set ObjBusnssDialog=JavaWindow("Business Modeler").JavaWindow("Find Business Object")
	If strProjectName<>"" Then
		Call Fn_List_Select("Fn_BMIDE_FindObjects",ObjBusnssDialog, "Project",strProjectName)
	End If
	Call Fn_Edit_Box("Fn_BMIDE_FindObjects",ObjBusnssDialog,"Criteria",strObjectName)
	ObjBusnssDialog.JavaTable("Table").SelectRow 0
    Call Fn_Button_Click("Fn_BMIDE_FindObjects", ObjBusnssDialog, "OK")
	Fn_BMIDE_FindObjects=True
	Set ObjBusnssDialog=Nothing
End Function 
'------------------------------------------------------------'Function Used to Perform Operation On Relation Properties-----------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_RelationPropertiesOperation

'Description			 :	Function Used to Perform Operation On Relation Properties

'Parameters			   :   '1.strAction: Action Name
'										1.strRelBusinessObj= Relation Bussiness Object
'										2.  strDescription= Relation Property Description
'										3. strDescription = String Descrption of Property

'Return Value		   : 	True Or False

'Pre-Requisit			: Should Be Log In BMIDE

'Examples				:	 Call Fn_BMIDE_RelationPropertiesOperation("Add","CM_state","Relation Property Description")
'										 Call Fn_BMIDE_RelationPropertiesOperation("SelectProperty","CM_state","")

'History					 :			
'							Developer Name					Date							Rev. No.						Changes Done								Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'							Sandeep N						29-Dec-2010						1.0																								Sunny
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------	
'							Pranav Ingle					13-mar-2013						 1.0						Modified Case SelectProperty				Sunny
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------	
Public Function Fn_BMIDE_RelationPropertiesOperation(strAction,strRelBusinessObj,strDescription)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_RelationPropertiesOperation"
	Dim ObjPropDialog
	Dim intRowCount,intCounter,strPropName
	Fn_BMIDE_RelationPropertiesOperation=False
	'Activating Properties Tab
	Call Fn_BMIDE_InnerTabOperations("Activate","Properties")
	Select Case strAction
		Case "Add"
				Call Fn_Button_Click("Fn_BMIDE_RelationPropertiesOperation", JavaWindow("Business Modeler"), "AddPropeties")
				wait(2)
				Set ObjPropDialog=JavaWindow("Business Modeler").JavaWindow("NewCustomProperty")
				'Checking Existance of NewCustomProperty Window
				If Not ObjPropDialog.Exist(5) Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail: NewCustomProperty Dialog Is Not Exist")
					Exit Function
				End If
				'Selecting Relation option
				Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_BMIDE_RelationPropertiesOperation",ObjPropDialog.JavaRadioButton("PropertyType"),"attached text","Relation")
				Call Fn_UI_JavaRadioButton_SetON("Fn_BMIDE_RelationPropertiesOperation",ObjPropDialog, "PropertyType")
				'Clicking On Next Button
				Call Fn_Button_Click("Fn_BMIDE_RelationPropertiesOperation", ObjPropDialog, "Next")
				If strRelBusinessObj<>"" Then
					'Setting Relation Bussiness Object
					Call Fn_Edit_Box("Fn_BMIDE_RelationPropertiesOperation",ObjPropDialog ,"RelationBusinessObj",strRelBusinessObj)
				End If
				If strDescription<>"" Then
					Call Fn_Edit_Box("Fn_BMIDE_RelationPropertiesOperation",ObjPropDialog ,"Description",strDescription)
				End If
				' Click on Finnish Button
				Call Fn_Button_Click("Fn_BMIDE_RelationPropertiesOperation",ObjPropDialog,"Finish")
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Pass: Function Completed Successfully")
				Fn_BMIDE_RelationPropertiesOperation = True
				Set ObjPropDialog=Nothing
		Case "SelectProperty"
				intRowCount=Fn_UI_Object_GetROProperty("Fn_BMIDE_RelationPropertiesOperation",JavaWindow("Business Modeler").JavaTable("PropertiesTable"), "rows")
				For intCounter=0 To intRowCount-1
					strPropName=JavaWindow("Business Modeler").JavaTable("PropertiesTable").GetCellData(intCounter,"Property Name")
					If Trim(strRelBusinessObj)=Trim(strPropName) Then
						JavaWindow("Business Modeler").JavaTable("PropertiesTable").SelectCell intCounter,0
						Fn_BMIDE_RelationPropertiesOperation=True
						Exit For
					End If
				Next
	End Select
End Function

'------------------------------------------------------------'Function Used to Perform Operation On GRM Rules-----------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_GRMRuleOperations

'Description			 :	Function Used to Perform Operation On GRM Rules

'Parameters			   :   '1.strAction: Action Name
'										 2.strPrimaryObj= Primary Object Name
'										 3.strSecondaryObj= Secondary Object Name
'										 4.strRelationObj = Relation Object Name
'										 5.strPrimaryCardinality = Primary Cardinality
'										 6.strSecondaryCardinality = Secondary Cardinality
'										 7.strChangeability = Changeability
'										 8.strAttachability = Attachability
'										 9.strDetachability = Detachability

'Return Value		   : 	True Or False

'Pre-Requisit			: Should Be Log In BMIDE

'Examples				:	 Call Fn_BMIDE_GRMRuleOperations("Add","AbsOccData","AbsOccGrmAnchor","CMHasWorkBreakdown","","","","","")
'										Call Fn_BMIDE_GRMRuleOperations("Verify","D2_bm_Airfrme400","Part Revision","TC_Is_Represented_By","","","","","")
'										Call Fn_BMIDE_GRMRuleOperations("Select","D2_bm_Airfrme400","Part Revision","TC_Is_Represented_By","","","","","")
'										Call Fn_BMIDE_GRMRuleOperations("Remove","D2_bm_Airfrme400","Part Revision","TC_Is_Represented_By","","","","","")
'										Call Fn_BMIDE_GRMRuleOperations("Modify","","","","2","2","","","")
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done							Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N												29-Dec-2010								1.0																					 Sunny
'													Sandeep N												29-Dec-2010								1.0									Case "Verify"							 Sunny
'													Pranav Ingle											 04-Jan-2012							  1.1							Cases "Select","Remove"				Sandeep
'													Pranav Ingle											 06-Jan-2012							  1.2							Cases "Modify"								Sandeep
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------	
Public Function Fn_BMIDE_GRMRuleOperations(strAction,strPrimaryObj,strSecondaryObj,strRelationObj,strPrimaryCardinality,strSecondaryCardinality,strChangeability,strAttachability,strDetachability)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_GRMRuleOperations"
   Dim ObjGRMRuleDialog,bReturn
   Dim iRowCount,iCounter,strRuleName,strScndObj,strRelObj
	Fn_BMIDE_GRMRuleOperations=False
	If JavaWindow("Business Modeler").JavaTab("InnerTab").Exist(5) then
		'Activating GRM Rules Tab
		Call Fn_BMIDE_InnerTabOperations("Activate","GRM Rules")
	End If

	Select Case strAction
		Case "Add"
			Set ObjGRMRuleDialog=JavaWindow("Business Modeler").JavaWindow("NewGRMRule")
			If Not  ObjGRMRuleDialog.Exist(5) Then
				Call Fn_Button_Click("Fn_BMIDE_GRMRuleOperations", JavaWindow("Business Modeler"), "AddGRMRule")
			End If
			If strPrimaryObj<>"" Then
				'Setting Primary Object
				If Trim(LCase(Fn_SISW_UI_JavaEdit_Operations("Fn_BMIDE_GRMRuleOperations", "GetText", ObjGRMRuleDialog, "PrimaryObject", ""))) <> LCase(strPrimaryObj) Then
					Call Fn_Edit_Box("Fn_BMIDE_GRMRuleOperations",ObjGRMRuleDialog ,"PrimaryObject",strPrimaryObj)	
				End If
			End If
			If strSecondaryObj<>"" Then
				'Setting Secondary Object
				Call Fn_Edit_Box("Fn_BMIDE_GRMRuleOperations",ObjGRMRuleDialog ,"SecondaryObject",strSecondaryObj)
			End If
			If strRelationObj<>"" Then
				'Setting Relation Object
				Call Fn_Edit_Box("Fn_BMIDE_GRMRuleOperations",ObjGRMRuleDialog ,"RelationObject",strRelationObj)
			End If
			
			'Setting Condition to [isTrue] (default condition), as fields has become mandatory since Tc113_0220 build
			Call Fn_Edit_Box("Fn_BMIDE_GRMRuleOperations",ObjGRMRuleDialog ,"Condition","isTrue")
			
			If strPrimaryCardinality<>"" Then
				'Setting Primary Cardinality
				Call Fn_Edit_Box("Fn_BMIDE_GRMRuleOperations",ObjGRMRuleDialog ,"PrimaryCardinality",strPrimaryCardinality)
			End If
			If strSecondaryCardinality<>"" Then
				'Setting Secondary Cardinality
				Call Fn_Edit_Box("Fn_BMIDE_GRMRuleOperations",ObjGRMRuleDialog ,"SecondaryCardinality",strSecondaryCardinality)
			End If
			If strChangeability<>"" Then
				'Selecting Changeability Of GRM Rule
				Call Fn_List_Select("Fn_BMIDE_GRMRuleOperations", ObjGRMRuleDialog, "Changeability",strChangeability)
			End If
			If strAttachability<>"" Then
				'Selecting Attachability Of GRM Rule
				Call Fn_List_Select("Fn_BMIDE_GRMRuleOperations", ObjGRMRuleDialog, "Attachability",strAttachability)
			End If
			If strDetachability<>"" Then
				'Selecting Detachability Of GRM Rule
				Call Fn_List_Select("Fn_BMIDE_GRMRuleOperations", ObjGRMRuleDialog, "Detachability",strDetachability)
			End If
			Call Fn_Button_Click("Fn_BMIDE_GRMRuleOperations", ObjGRMRuleDialog, "Finish")
			Fn_BMIDE_GRMRuleOperations=True
			Set ObjGRMRuleDialog=Nothing

		Case "Modify","Modify_DoubleClick"
			JavaWindow("Business Modeler").JavaWindow("NewGRMRule").SetTOProperty "title","Modify GRM Rule"
            If strAction="Modify_DoubleClick" Then
				Call Fn_BMIDE_GRMRuleOperations("DoubleClick",strPrimaryObj,strSecondaryObj,strRelationObj,"","","","","")
			Else
				Set ObjGRMRuleDialog=JavaWindow("Business Modeler").JavaWindow("NewGRMRule")
				If Not  ObjGRMRuleDialog.Exist(5) Then
					Call Fn_Button_Click("Fn_BMIDE_GRMRuleOperations", JavaWindow("Business Modeler"), "EditDeepCopyRule")
				End If
			End If
			Set ObjGRMRuleDialog=JavaWindow("Business Modeler").JavaWindow("NewGRMRule")
			
			If strPrimaryCardinality<>"" Then
				'Setting Primary Cardinality
				Call Fn_Edit_Box("Fn_BMIDE_GRMRuleOperations",ObjGRMRuleDialog ,"PrimaryCardinality",strPrimaryCardinality)
			End If
			If strSecondaryCardinality<>"" Then
				'Setting Secondary Cardinality
				Call Fn_Edit_Box("Fn_BMIDE_GRMRuleOperations",ObjGRMRuleDialog ,"SecondaryCardinality",strSecondaryCardinality)
			End If
			If strChangeability<>"" Then
				'Selecting Changeability Of GRM Rule
				Call Fn_List_Select("Fn_BMIDE_GRMRuleOperations", ObjGRMRuleDialog, "Changeability",strChangeability)
			End If
			If strAttachability<>"" Then
				'Selecting Attachability Of GRM Rule
				Call Fn_List_Select("Fn_BMIDE_GRMRuleOperations", ObjGRMRuleDialog, "Attachability",strAttachability)
			End If
			If strDetachability<>"" Then
				'Selecting Detachability Of GRM Rule
				Call Fn_List_Select("Fn_BMIDE_GRMRuleOperations", ObjGRMRuleDialog, "Detachability",strDetachability)
			End If
			Call Fn_Button_Click("Fn_BMIDE_GRMRuleOperations", ObjGRMRuleDialog, "Finish")

			JavaWindow("Business Modeler").JavaWindow("NewGRMRule").SetTOProperty "title","New GRM Rule"
			Fn_BMIDE_GRMRuleOperations=True
			Set ObjGRMRuleDialog=Nothing
		Case "Verify"
			iRowCount=JavaWindow("Business Modeler").JavaTable("GRMRuleTable").GetROProperty("rows")
			For iCounter=0 To iRowCount-1
				strRuleName=JavaWindow("Business Modeler").JavaTable("GRMRuleTable").GetCellData(iCounter,"Primary")
				If Trim(strRuleName)=Trim(strPrimaryObj) Then
					strScndObj=JavaWindow("Business Modeler").JavaTable("GRMRuleTable").GetCellData(iCounter,"Secondary")
					strRelObj=JavaWindow("Business Modeler").JavaTable("GRMRuleTable").GetCellData(iCounter,"Relation")
					If Trim(strScndObj)=Trim(strSecondaryObj) And Trim(strRelObj)=Trim(strRelationObj) Then
						If strChangeability<>"" Then
							If JavaWindow("Business Modeler").JavaTable("GRMRuleTable").GetCellData(iCounter,"Changeability")=strChangeability Then
							Else
								Exit for
							End If
						End If
						If strAttachability<>"" Then
							If JavaWindow("Business Modeler").JavaTable("GRMRuleTable").GetCellData(iCounter,"Attachability")=strAttachability Then
							Else
								Exit for
							End If
						End If
						If strDetachability<>"" Then
							If JavaWindow("Business Modeler").JavaTable("GRMRuleTable").GetCellData(iCounter,"Detachability")=strDetachability Then
							Else
								Exit for
							End If
						End If
						Fn_BMIDE_GRMRuleOperations=True
						Exit For
					End If
				End If
			Next
		Case "Select","DoubleClick"
			iRowCount=JavaWindow("Business Modeler").JavaTable("GRMRuleTable").GetROProperty("rows")
			For iCounter=0 To iRowCount-1
				
				strRuleName=JavaWindow("Business Modeler").JavaTable("GRMRuleTable").GetCellData(iCounter,"Primary")
				If Trim(strRuleName)=Trim(strPrimaryObj) Then
					strScndObj=JavaWindow("Business Modeler").JavaTable("GRMRuleTable").GetCellData(iCounter,"Secondary")
					strRelObj=JavaWindow("Business Modeler").JavaTable("GRMRuleTable").GetCellData(iCounter,"Relation")
					If Trim(strScndObj)=Trim(strSecondaryObj) And Trim(strRelObj)=Trim(strRelationObj) Then
						If strAction="Select" Then
							JavaWindow("Business Modeler").JavaTable("GRMRuleTable").SelectCell iCounter,0
							wait 1
						End If
						If strAction="DoubleClick" Then
							JavaWindow("Business Modeler").JavaTable("GRMRuleTable").ActivateRow iCounter
							wait 1
						End If
						Fn_BMIDE_GRMRuleOperations=True
						Exit For
					End If

				End If
			Next
		Case "Remove"
			bReturn=Fn_BMIDE_GRMRuleOperations("Select",strPrimaryObj,strSecondaryObj,strRelationObj,"","","","","")
			If  bReturn = False Then
				Exit Function
			Else
				Call Fn_Button_Click("Fn_BMIDE_GRMRuleOperations", JavaWindow("Business Modeler"), "RemoveDeepCopyRule")
				Fn_BMIDE_GRMRuleOperations=True
			End If

	End Select
End Function

'-------------------------------------------------------------------Function Used to Perform Operation On MasterAlternateIDRules Table----------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_MasterAlternateIDRulesTableOperations

'Description			 :	Function Used to Perform Operation On MasterAlternateIDRules Table

'Parameters			   :  1.strAction : Action Name
'										2.strName: Master Alternate ID Rule Name
										'3.strRule: Rule Argument
										'4.strDescription:Rule Description
										'5.strColumnName:Column Name
										'6.strExpValue:Expected Value

'Return Value		   : 	True Or False

'Pre-requisite			:	Should Be Log In BMIDE

'Examples				: Call Fn_BMIDE_MasterAlternateIDRulesTableOperations("Verify","S3_ACME36013@Identifier","","","Identifier Context","S3_ACME36013")
'									Call Fn_BMIDE_MasterAlternateIDRulesTableOperations("Verify","S3_Logan78648@Identifier","","","Rule","Test")
'									Call Fn_BMIDE_MasterAlternateIDRulesTableOperations("Remove","T3bb@Identifier","","","","")
'									Call Fn_BMIDE_MasterAlternateIDRulesTableOperations("DoubleClick","T3bb@Identifier","","","","")

'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done								Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				04/01/2011			           1.0																					Sunny R
'  													Swapna G										   				18/01/2011			           1.0							"Remove"		   								Sandeep N
'  													Pranav Ingle										   			2-Aug-2013			           1.1							"DoubleClick"		   							Sunny R
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_BMIDE_MasterAlternateIDRulesTableOperations(strAction,strName,strRule,strDescription,strColumnName,strExpValue)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_MasterAlternateIDRulesTableOperations"
    'Variable Declaration
   Dim bFlag,iRowCount,iCounter,strCurrName,strCurrVal
   Dim ObjBmideDialog,ObjTableDialog
   'Creating Object Of "Business Modeler" window and "MasterAlternateIdRulesTable" Table
	Set ObjBmideDialog=JavaWindow("Business Modeler")
	Set ObjTableDialog=JavaWindow("Business Modeler").JavaTable("MasterAlternateIdRulesTable")
	'Initially Function Returns False
	Fn_BMIDE_MasterAlternateIDRulesTableOperations=False
	'Setting bFlag=False
	bFlag=False
	'Activating Alternate Id Rules Tab
	Call Fn_BMIDE_InnerTabOperations("Activate","Alternate ID Rules")
	Select Case strAction
		Case "Verify" 'Case to Verify Table Values
			'Taking row count Of "MasterAlternateIdRulesTable" Table
			iRowCount=Fn_UI_Object_GetROProperty("Fn_BMIDE_MasterAlternateIDRulesTableOperations",ObjTableDialog, "rows")
            For iCounter=0 To iRowCount-1
				strCurrName=ObjTableDialog.GetCellData(iCounter,"Name")
				If Trim(strCurrName)=Trim(strName) Then
					strCurrVal=ObjTableDialog.GetCellData(iCounter,strColumnName)
					If Trim(strExpValue)=Trim(strCurrVal) Then
						bFlag=True
					End If
					Exit For
				End If
			Next
			If bFlag=True Then
				Fn_BMIDE_MasterAlternateIDRulesTableOperations=True
			End If

	Case "Remove","Select" , "DoubleClick"
			'Taking row count Of "MasterAlternateIdRulesTable" Table
			iRowCount=Fn_UI_Object_GetROProperty("Fn_BMIDE_MasterAlternateIDRulesTableOperations",ObjTableDialog, "rows")
            For iCounter=0 To iRowCount-1
				strCurrName=ObjTableDialog.GetCellData(iCounter,"Name")
				If Trim(strCurrName)=Trim(strName) Then
					JavaWindow("Business Modeler").JavaTable("MasterAlternateIdRulesTable").SelectCell iCounter, 0
					If strAction="DoubleClick" Then
						JavaWindow("Business Modeler").JavaTable("MasterAlternateIdRulesTable").ActivateRow iCounter
					End If
                    bFlag=True
                    Exit For
				End If
			Next
			If bFlag=True Then
				If strAction="Remove" Then
					Call Fn_Button_Click("Fn_BMIDE_MasterAlternateIDRulesTableOperations",JavaWindow("Business Modeler"), "RemoveAlternateIDRule")
				End If
                Fn_BMIDE_MasterAlternateIDRulesTableOperations=True
			End If

	End Select
	'Releasing Object of "Business Modeler" window
	Set ObjBmideDialog=Nothing
	'Releasing Object Of "MasterAlternateIdRulesTable" Table
	Set ObjTableDialog=Nothing
End Function


'-------------------------------------------------------------------Function Used to Perform Operation On SupplementalAlternateIDRules Table----------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_SupplementalAlternateIDRulesTableOperations

'Description			 :	Function Used to Perform Operation On SupplementalAlternateIDRules Table

'Parameters			   :  1.strAction : Action Name
'										2.strName: Master Alternate ID Rule Name
										'3.strRule: Rule Argument
										'4.strDescription:Rule Description
										'5.strColumnName:Column Name
										'6.strExpValue:Expected Value

'Return Value		   : 	True Or False

'Pre-requisite			:	Should Be Log In BMIDE

'Examples				: Call Fn_BMIDE_SupplementalAlternateIDRulesTableOperations("Verify","S3_ACME36013@IdentifierRev","","","Identifier Type","IdentifierRev")
'									  Call Fn_BMIDE_SupplementalAlternateIDRulesTableOperations("Verify","S3_Logan78648@IdentifierRev","","","Identifier Context","S3_Logan78648")

'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done								Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				04/01/2011			           1.0																								Sunny R
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_BMIDE_SupplementalAlternateIDRulesTableOperations(strAction,strName,strRule,strDescription,strColumnName,strExpValue)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_SupplementalAlternateIDRulesTableOperations"
   'Variable Declaration
   Dim bFlag,iRowCount,iCounter,strCurrName,strCurrVal
   Dim ObjBmideDialog,ObjTableDialog
   'Creating Object Of "Business Modeler" window and "SupplementalAlternateIdRulesTable" Table
	Set ObjBmideDialog=JavaWindow("Business Modeler")
	Set ObjTableDialog=JavaWindow("Business Modeler").JavaTable("SupplementalAlternateIdRulesTable")
	'Initially Function Returns False
	Fn_BMIDE_SupplementalAlternateIDRulesTableOperations=False
	'Setting bFlag=False
	bFlag=False
	'Activating Alternate Id Rules Tab
	Call Fn_BMIDE_InnerTabOperations("Activate","Alternate ID Rules")
	Select Case strAction
		Case "Verify" 'Case to Verify Table Values
			'Taking row count Of "SupplementalAlternateIdRulesTable" Table
			iRowCount=Fn_UI_Object_GetROProperty("Fn_BMIDE_SupplementalAlternateIDRulesTableOperations",ObjTableDialog, "rows")
            For iCounter=0 To iRowCount-1
				strCurrName=ObjTableDialog.GetCellData(iCounter,"Name")
				If Trim(strCurrName)=Trim(strName) Then
					strCurrVal=ObjTableDialog.GetCellData(iCounter,strColumnName)
					If Trim(strExpValue)=Trim(strCurrVal) Then
						bFlag=True
					End If
					Exit For
				End If
			Next
			If bFlag=True Then
				Fn_BMIDE_SupplementalAlternateIDRulesTableOperations=True
			End If
	End Select
	'Releasing Object of "Business Modeler" window
	Set ObjBmideDialog=Nothing
	'Releasing Object Of "SupplementalAlternateIdRulesTable" Table
	Set ObjTableDialog=Nothing
End Function


'------------------------------------------------------------'Function Used to Perform Operations On Compound Properties-------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_CompoundPropertiesOperation

'Description			 :	Function Used to Perform Operations On Compound Properties

'Parameters			   :   '1.strAction: Action Name
'										2.strName= Property name
'										3.  strDisplayName=Display name of property
'										4. strDescription = Descrption of Property
'										5.,bReadOnly = ReadOnly Option
'										6. strPropertyName=Compound Property Name
'										7. strObjectName = Business Object Name
'										8. strFinalSegProp=Final Segment Property Name

'Return Value		   : 	True Or False


'Examples				:	Fn_BMIDE_CompoundPropertiesOperation("Add","Test24","","Test","Off","bl_bomview","s3_Form61052","change")
'										Fn_BMIDE_CompoundPropertiesOperation("Add","Test24","","Test","Off","","s3_Form61052","change")

'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N												05-Jan-2011								1.0																						Sunny
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------	

Public Function Fn_BMIDE_CompoundPropertiesOperation(strAction,strName,strDisplayName,strDescripton,bReadOnly,strPropertyName,strObjectName,strFinalSegProp)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_CompoundPropertiesOperation"
   Dim strPrototype,strTopItem
   Dim ObjPropDialog,ObjPropSegDialog
   Fn_BMIDE_CompoundPropertiesOperation=False
	Set ObjPropDialog=JavaWindow("Business Modeler").JavaWindow("NewCustomProperty")
	Set ObjPropSegDialog=JavaWindow("Business Modeler").JavaWindow("AddCompoundPropertySeg")
	Select Case strAction
		Case "Add"
			If Not ObjPropDialog.Exist(10) Then
				'Clicking On Add Button To Add Compound Properties 
				Call Fn_Button_Click("Fn_BMIDE_CompoundPropertiesOperation", JavaWindow("Business Modeler"), "AddPropeties")			
			End If
			'Selecting Compound option
			Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_BMIDE_CompoundPropertiesOperation",ObjPropDialog.JavaRadioButton("PropertyType"),"attached text","Compound")
			Call Fn_UI_JavaRadioButton_SetON("Fn_BMIDE_CompoundPropertiesOperation",ObjPropDialog, "PropertyType")
			'Clicking On Next Button
			Call Fn_Button_Click("Fn_BMIDE_CompoundPropertiesOperation", ObjPropDialog, "Next")
			If strName<>"" Then
				strPrototype= Fn_Edit_Box_GetValue("Fn_BMIDE_CompoundPropertiesOperation",ObjPropDialog,"Name")
				strName=strPrototype+strName
				'Setting name to new compound property
				Call Fn_Edit_Box("Fn_BMIDE_CompoundPropertiesOperation",ObjPropDialog ,"Name",strName)
				If strDisplayName="" Then
					strDisplayName=strName
				End If
			End If
			If strDisplayName<>"" Then
				'Setting Display Name to new compound property
				Call Fn_Edit_Box("Fn_BMIDE_CompoundPropertiesOperation",ObjPropDialog ,"DisplayName",strDisplayName)
			End If
			If strDescripton<>"" Then
				'Setting Description to new compound property
				Call Fn_Edit_Box("Fn_BMIDE_CompoundPropertiesOperation",ObjPropDialog ,"Description",strDescripton)
			End If
			If bReadOnly<>"" Then
				'Setting Status Of "ReadOnly" Option
				Call Fn_CheckBox_Set("Fn_BMIDE_CompoundPropertiesOperation", ObjPropDialog, "ReadOnly", bReadOnly)
			End If
			

			If strPropertyName<>"" Then
				Call Fn_Button_Click("Fn_BMIDE_CompoundPropertiesOperation", ObjPropDialog, "AddSegment")
				wait(3)
				Call Fn_Edit_Box("Fn_BMIDE_CompoundPropertiesOperation",ObjPropSegDialog ,"PropertyName",strPropertyName)

			End If

			If strObjectName<>"" Then
				If Not ObjPropSegDialog.Exist(5) Then
					Call Fn_Button_Click("Fn_BMIDE_CompoundPropertiesOperation", ObjPropDialog, "AddSegment")
				End If
				wait(3)
				Call Fn_Button_Click("Fn_BMIDE_CompoundPropertiesOperation", ObjPropSegDialog, "Next")
				Call Fn_Edit_Box("Fn_BMIDE_CompoundPropertiesOperation",ObjPropSegDialog ,"PropertyName",strObjectName)
			End If

			Call Fn_Button_Click("Fn_BMIDE_CompoundPropertiesOperation", ObjPropSegDialog, "Finish")	
			If strFinalSegProp<>"" Then
				strTopItem=ObjPropDialog.JavaTree("PathTree").GetItem(1)
'				strTopItem=strTopItem+":"+strObjectName
				strTopItem=strTopItem
				ObjPropDialog.JavaTree("PathTree").Select strTopItem
				Call Fn_Button_Click("Fn_BMIDE_CompoundPropertiesOperation", ObjPropDialog, "AddFinalSegment")
				wait(3)
				Call Fn_Edit_Box("Fn_BMIDE_CompoundPropertiesOperation",ObjPropSegDialog ,"PropertyName",strFinalSegProp)
				Call Fn_Button_Click("Fn_BMIDE_CompoundPropertiesOperation", ObjPropSegDialog, "Finish")	
			End If
			Call Fn_Button_Click("Fn_BMIDE_CompoundPropertiesOperation", ObjPropDialog, "Finish")
			Fn_BMIDE_CompoundPropertiesOperation=True
	End Select
	Set ObjPropDialog=Nothing
	Set ObjPropSegDialog=Nothing
End Function
'-------------------------------------------------------------------Function Used to Create New LOV  of Date Type ---------------------------------------------------------------------------------------------------------------------------------------- 
'Function Name		:	Fn_BMIDE_NewLOVCreateExt

'Description			 :	Function Used to Create New List Of Value of Date Type 

'Parameters			   :	1.strProject: Projetct Name
										'2.strName:Name Of LOV
										'3.strDesc: LOV Description
										'4.strType:LOV type
										'5.bUsage:Usage Option
										'6.strReference:LOV Reference
										'7.strLower: Lower Boundry of LOV
										'8.strUpper:Upper Boundry of LOV
										'9.bCascdingView:Cascding View Option
										'10.arrProperties: Properties

'Return Value		   : 	True Or False

'Pre-requisite			:	Should be Log In BMIDE

'Examples				: abc="LA~LA~Los Engelis~isTrue:USA~United States~Unite States~"
'									  Call Fn_BMIDE_NewLOVCreateExt("","TestLOV4","First Test LOV Object7","ListOfValuesString","Exhaustive","","","","OFF",abc)

'										Fn_BMIDE_NewLOVCreateExt("","TestLOV4","First Test LOV Object7","ListOfValuesFilter","Exhaustive","Based On LOV~D3MyLOV1_07759","","","","","A~C")
'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Pranav Ingle										   			 071/2011			           1.0																					Sunny R
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_BMIDE_NewLOVCreateExt(strProject,strName,strDesc,strType,bUsage,strReference,strLower,strUpper,bCascdingView,arrProperties,strValues)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_NewLOVCreateExt"
   Dim ObjNewLOVWindow
   Dim bFlag,strPrototype,arrPropValues,arrPropPairs,iCounter,iCount
  'Function Return False
   Fn_BMIDE_NewLOVCreateExt=False
   'Checking Existance of NewLOVObject window
   If Fn_UI_ObjectExist("Fn_BMIDE_NewLOVCreateExt",JavaWindow("Business Modeler").JavaWindow("NewLOV"))=False Then
	   'If NewLOVObject window not exist then function will Exit
       Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail:NewLOVObject Window is not Exist ")
	   Exit Function
   End If
	'Creating Object of NewLOVObject window
	Set  ObjNewLOVWindow=Fn_UI_ObjectCreate("Fn_BMIDE_NewLOVCreateExt",JavaWindow("Business Modeler").JavaWindow("NewLOV"))

    Call Fn_UI_JavaRadioButton_SetON("Fn_BMIDE_NewLOVCreateExt",ObjNewLOVWindow, "ClassicLOV")
	Call Fn_Button_Click("Fn_BMIDE_NewLOVCreateExt", ObjNewLOVWindow, "Next")

	If strProject<>"" Then
		'Verifying Project is exist in Project List Or Not
		bFlag=Fn_UI_ListItemExist("Fn_BMIDE_NewLOVCreateExt", ObjNewLOVWindow, "Project",strProject)
		If bFlag=False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail:["+strProject+"] is not present in Project List")
			Set ObjNewLOVWindow=Nothing
			Exit Function
		End If
		'Selecting Project from Project List
		Call Fn_List_Select("Fn_BMIDE_NewLOVCreateExt", ObjNewLOVWindow, "Project",strProject)
	End If
	If strName<>"" Then
		strPrototype= Fn_Edit_Box_GetValue("Fn_BMIDE_NewLOVCreateExt",ObjNewLOVWindow,"Name")
		strName=strPrototype+strName
		'Setting Name to New LOV Object
        Call Fn_Edit_Box("Fn_BMIDE_NewLOVCreateExt",ObjNewLOVWindow,"Name",strName)
	End If
	If strDesc<>"" Then
		'Setting Description to New LOV Object
        Call Fn_Edit_Box("Fn_BMIDE_NewLOVCreateExt",ObjNewLOVWindow,"Description",strDesc)
	End If
	If strType<>"" Then
		'Setting Type to New LOV Object
        'Call Fn_Edit_Box("Fn_BMIDE_NewLOVCreateExt",ObjNewLOVWindow,"Type",strType)
        	Call Fn_Button_Click("Fn_BMIDE_NewLOVCreateExt",ObjNewLOVWindow, "BrowseLOVType")  
		Call Fn_Edit_Box("Fn_BMIDE_NewLOVCreateExt", JavaWindow("Business Modeler").JavaWindow("Find Business Object"),"Project",strType)  
		Call Fn_SISW_UI_JavaTable_Operations("Fn_BMIDE_NewLOVCreateExt", "SelectCell", JavaWindow("Business Modeler").JavaWindow("Find Business Object").JavaTable("Table") , "", "", "",strType , 0, "", "", "")
		Call Fn_Button_Click("Fn_BMIDE_NewLOVCreateExt", JavaWindow("Business Modeler").JavaWindow("Find Business Object"), "OK")
	End If
	If bUsage<>"" Then
		Call Fn_UI_Object_SetTOProperty("Fn_BMIDE_NewLOVCreateExt",ObjNewLOVWindow.JavaRadioButton("Usage"),"attached text",bUsage)
		'Setting Usage to New LOV Object
        Call Fn_UI_JavaRadioButton_SetON("Fn_BMIDE_NewLOVCreateExt",ObjNewLOVWindow, "Usage")
	End If
	If strReference<>"" Then
		arrPropValues = Split(strReference,"~")
		If UBound(arrPropValues) > 0 Then
				Call Fn_Edit_Box("Fn_BMIDE_NewLOVCreateExt",ObjNewLOVWindow,"BasedOnLOV",arrPropValues(1))
				If  strValues <> "" Then
					arrPropPairs = Split(strValues,"~")
					For iCounter = 0 to UBound(arrPropPairs)
							bReturn=Fn_JavaTree_NodeIndexExt("", ObjNewLOVWindow, "Reference", arrPropPairs(iCounter), ";", "")
							JavaWindow("Business Modeler").JavaTree("LOVTable").SetItemState "#"+CStr(bReturn),micChecked 

			'				JavaWindow("Business Modeler").JavaWindow("NewLOV").JavaTree("Reference").SetItemState "A",micChecked 
'							ObjNewLOVWindow.JavaTree("Reference").SetItemState arrPropPairs(iCounter),micChecked
					Next
					Call Fn_List_Select("Fn_BMIDE_NewLOVCreateExt",ObjNewLOVWindow,"Show","Selected")
				End If
				'JavaWindow("Business Modeler").JavaWindow("NewLOV").JavaList("Show").Select "Selected"
		Else
			'Setting Reference to New LOV Object
			Call Fn_Edit_Box("Fn_BMIDE_NewLOVCreateExt",ObjNewLOVWindow,"Reference",strReference)
		End If
	End If
	If strLower<>"" Then
		'Setting Lower Range Of  New LOV Object
        Call Fn_Edit_Box("Fn_BMIDE_NewLOVCreateExt",ObjNewLOVWindow,"Lower",strLower)
	End If
	If strUpper<>"" Then
		'Setting Upper Range Of New LOV Object
        Call Fn_Edit_Box("Fn_BMIDE_NewLOVCreateExt",ObjNewLOVWindow,"Upper",strUpper)
	End If
	If bCascdingView<>"" Then
		'Setting Cascading Option
        Call Fn_CheckBox_Set("Fn_BMIDE_NewLOVCreateExt", ObjNewLOVWindow, "ShowCascadingView", bCascdingView)
	End If
	If arrProperties<>"" Then
        arrPropValues=Split(arrProperties,"!")	
		For iCounter=0 To Ubound( arrPropValues)
				Call Fn_Button_Click("Fn_BMIDE_NewLOVCreateExt", ObjNewLOVWindow, "Add")
				 If Fn_UI_ObjectExist("Fn_BMIDE_NewLOVCreateExt",JavaWindow("Business Modeler").JavaWindow("AddLOVValue"))=True Then
				arrPropPairs=Split(arrPropValues(iCounter),"~")
				For iCount=0 To Ubound(arrPropPairs)
						If arrPropPairs(0)<>"" Then
							 Call Fn_Edit_Box("Fn_BMIDE_NewLOVCreateExt",JavaWindow("Business Modeler").JavaWindow("AddLOVValue"),"Value",arrPropPairs(0))				
						End If
						If arrPropPairs(1)<>"" Then
							Call Fn_Edit_Box("Fn_BMIDE_NewLOVCreateExt",JavaWindow("Business Modeler").JavaWindow("AddLOVValue"),"ValueDisplayName",arrPropPairs(1))
						End If
						If arrPropPairs(2)<>"" Then
							Call Fn_Edit_Box("Fn_BMIDE_NewLOVCreateExt",JavaWindow("Business Modeler").JavaWindow("AddLOVValue"),"Description",arrPropPairs(2))			
						End If
						If arrPropPairs(3)<>"" Then
							Call Fn_Edit_Box("Fn_BMIDE_NewLOVCreateExt",JavaWindow("Business Modeler").JavaWindow("AddLOVValue"),"Condition",arrPropPairs(3))				
						End If
						Call Fn_Button_Click("Fn_BMIDE_NewLOVCreateExt", JavaWindow("Business Modeler").JavaWindow("AddLOVValue"), "Finish")
						Exit For
				Next
				End If
		Next
	End If
    'Clicking On Finish Button create New LOV Object
	Call Fn_Button_Click("Fn_BMIDE_NewLOVCreateExt", ObjNewLOVWindow, "Finish")
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Pass:Successfully Create New LOV Object of Display Name ["+strName+"]")
	'Function returns True after creating New LOV Object
	Fn_BMIDE_NewLOVCreateExt=True
	'Releasing Object of NewLOVObject window
	Set ObjNewLOVWindow=Nothing
End Function

'-------------------------------------------------------------------Function Used to Perform Operations On  Server Connection Profile--------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_ServerConnectionProfileOperations

''Description			:1. Function Used to Create New Server Connection Profile

'Return Value		   : 	True or False

'Parameters     		:	1.strAction:Action Name
'										1. strProfileName : Server Connection Profile Name
										'2. strProtocol : Protocol Name for Connection
										'3. strHost: Host Name
										'4. strPort: Port Number
										'5. strAppName: Application Name
										'6. strUserID: User ID
										'7. strGroup: User Group Name
										'8. strRole: User Role
										'9.strErrorMsg : Error Message Or Static Text
'Pre-requisite			:	BMIDE Prespective should be Open.

'Examples				: Call Fn_BMIDE_ServerConnectionProfileOperations("VerifyStaticText","DemoProfile","","","","","","","","This profile is associated as a master profile")
'									  Call Fn_BMIDE_ServerConnectionProfileOperations("DeleteWithError","DemoProfile","","","","","","","","")
'									  Call Fn_BMIDE_ServerConnectionProfileOperations("ModifyCancel","DemoProfile","","","","","","","","")

'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				10/01/2011			           1.0																						Sunny R
'													Sandeep N										   				27/01/2011			           1.0																						Sunny R
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_BMIDE_ServerConnectionProfileOperations(strAction,strProfileName,strProtocol,strHost,strPort,strAppName,strUserID,strGroup,strRole,strErrorMsg)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_ServerConnectionProfileOperations"
   'Variable Declaration
   Dim strMenu,bFlag,strCurrText
   bFlag=False
	Fn_BMIDE_ServerConnectionProfileOperations=False
	'Checking Existance Of "Preferences" Dialog
   If Fn_UI_ObjectExist("Fn_BMIDE_ServerConnectionProfileOperations",JavaWindow("Business Modeler").JavaWindow("Preferences"))=False Then
	   'Calling Window:Preference Menu to open Preference Dialog
		strMenu=Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\BMIDEConfig\BMIDE_Menu.xml", "Preferences")
		Call Fn_BMIDE_MenuOperation("Select", strMenu)
   End If
   'Expanding "Teamcenter" Node From "PreferenceTree"
	Call Fn_UI_JavaTree_Expand("Fn_BMIDE_ServerConnectionProfileOperations",JavaWindow("Business Modeler").JavaWindow("Preferences"),"PreferenceTree","Teamcenter")
	'Selecting "Teamcenter:Server Connection Profiles" node 
	Call Fn_JavaTree_Select("Fn_BMIDE_ServerConnectionProfileOperations",JavaWindow("Business Modeler").JavaWindow("Preferences"),"PreferenceTree","Teamcenter:Server Connection Profiles")
	If strProfileName<>"" Then
		bFlag=Fn_UI_ListItemExist("Fn_BMIDE_ServerConnectionProfileOperations", JavaWindow("Business Modeler").JavaWindow("Preferences"),"ListOfProfiles",strProfileName)
		If bFlag=True Then
			Call Fn_List_Select("Fn_BMIDE_ServerConnectionProfileOperations", JavaWindow("Business Modeler").JavaWindow("Preferences"),"ListOfProfiles",strProfileName)
		End If
	End If
	Select Case strAction
		Case "VerifyStaticText"
			'Clicking "Modify" Button To open "TeamcenterRepositoryConnection" Dialog
			Call Fn_Button_Click("Fn_BMIDE_ServerConnectionProfileOperations", JavaWindow("Business Modeler").JavaWindow("Preferences"), "Modify")
			strCurrText=JavaWindow("Business Modeler").JavaWindow("TeamcenterRepositoryConnection").JavaStaticText("StaticTextMsg").GetROProperty("attached text")
			wait(2)
			If InStr(1,strCurrText,strErrorMsg)>=1 Then
				Fn_BMIDE_ServerConnectionProfileOperations=True
			End If
			Call Fn_Button_Click("Fn_BMIDE_ServerConnectionProfileOperations", JavaWindow("Business Modeler").JavaWindow("TeamcenterRepositoryConnection"), "Cancel")
		Case "DeleteWithError"
			Call Fn_Button_Click("Fn_BMIDE_ServerConnectionProfileOperations", JavaWindow("Business Modeler").JavaWindow("Preferences"), "Delete")
			If JavaWindow("Business Modeler").JavaWindow("DeleteProfile").Exist(5) Then
				Call Fn_Button_Click("Fn_BMIDE_ServerConnectionProfileOperations", JavaWindow("Business Modeler").JavaWindow("DeleteProfile"), "OK")
				Fn_BMIDE_ServerConnectionProfileOperations=True
			End If

		Case "Delete"
			Call Fn_Button_Click("Fn_BMIDE_ServerConnectionProfileOperations", JavaWindow("Business Modeler").JavaWindow("Preferences"), "Delete")
			Fn_BMIDE_ServerConnectionProfileOperations=True
	Case "ModifyCancel"
			'Clicking "Modify" Button To open "TeamcenterRepositoryConnection" Dialog
			Call Fn_Button_Click("Fn_BMIDE_ServerConnectionProfileOperations", JavaWindow("Business Modeler").JavaWindow("Preferences"), "Modify")
			If JavaWindow("Business Modeler").JavaWindow("TeamcenterRepositoryConnection").Exist(5) Then
				Fn_BMIDE_ServerConnectionProfileOperations=True
			End If
			Call Fn_Button_Click("Fn_BMIDE_ServerConnectionProfileOperations", JavaWindow("Business Modeler").JavaWindow("TeamcenterRepositoryConnection"), "Cancel")

	End Select
	
	Call Fn_Button_Click("Fn_BMIDE_ServerConnectionProfileOperations", JavaWindow("Business Modeler").JavaWindow("Preferences"), "OK")
End Function


'-------------------------------------------------------------------Function Used to Verify Error Message And Error Window---------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_ErrorWindowMsgVerify

'Description			 :	Function Used to Verify Error Message And Error Window

'Parameters			   :	1.strDialogName: Error Dialog Name
										'2.strErrorMsg:Expected Error Message
										'3.strButtton: Button Name

'Return Value		   : 	True Or False

'Pre-requisite			:	Error Message Should be Appear on Screen

'Examples				: 	Call Fn_BMIDE_ErrorWindowMsgVerify("Deployment Failed","","OK")
'										Call Fn_BMIDE_ErrorWindowMsgVerify("Deployment Failed","","View log Files")

'History					 :		
'													Developer Name												Date						Rev. No.						Changes Done														Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				10/01/2011			           1.0																																					Sunny R
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_BMIDE_ErrorWindowMsgVerify(strDialogName,strErrorMsg,strButtton)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_ErrorWindowMsgVerify"
   'Variablr Declaration
   Dim strMsg,bFlag
   GBL_EXPECTED_MESSAGE=strErrorMsg
   Fn_BMIDE_ErrorWindowMsgVerify=False
   bFlag=False
   'Setting Dialog Title
   Call Fn_UI_Object_SetTOProperty("Fn_BMIDE_ErrorWindowMsgVerify",JavaWindow("Business Modeler").JavaWindow("ErrorWindow"),"title",strDialogName)
   'Checking Existance Of Error Dialog
   If Fn_UI_ObjectExist("Fn_BMIDE_ErrorWindowMsgVerify", JavaWindow("Business Modeler").JavaWindow("ErrorWindow"))=True Then
	   bFlag=True
	   If strErrorMsg<>"" Then
		   bFlag=False
		   'Retriving Error Message which Apprers on Dialog
		   If JavaWindow("Business Modeler").JavaWindow("ErrorWindow").JavaStaticText("ErrMsg").exist Then
		  		strMsg=Fn_UI_Object_GetROProperty("Fn_BMIDE_ErrorWindowMsgVerify",JavaWindow("Business Modeler").JavaWindow("ErrorWindow").JavaStaticText("ErrMsg"),"text")		   	
		   Else
				strMsg=Fn_UI_Object_GetROProperty("Fn_BMIDE_ErrorWindowMsgVerify",JavaWindow("Business Modeler").JavaWindow("ErrorWindow").Static("ErrorMessage"),"text")		   
		   End If
		  'Verifying Error Message Match With Expected Error Message
			If InStr(1,strMsg,strErrorMsg)>=1 Then
			  'bFlag Set to True
			  bFlag=True
			Else
				GBL_ACTUAL_MESSAGE=strMsg
			End If
	   End If
	  'Clicking On strButtton Button
	  Call Fn_Button_Click("Fn_BMIDE_ErrorWindowMsgVerify", JavaWindow("Business Modeler").JavaWindow("ErrorWindow"), strButtton)
   End If
   If bFlag=True Then
	   Fn_BMIDE_ErrorWindowMsgVerify=True
   End If
End Function

'-------------------------------------------------------------------Function Used to Verify Logs From Main Tab---------------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_VerifyLogs

'Description			 :	Function Used to Verify Logs From Main Tab

'Parameters			   :	1.strLogText: Expected Log

'Return Value		   : 	True Or False

'Pre-requisite			:	Log Should be Appear on Screen

'Examples				: 	Call Fn_BMIDE_VerifyLogs("The update has been blocked because the preference BMIDE_ALLOW_OPS_DEPLOY is set to false.")

'History					 :		
'													Developer Name												Date						Rev. No.						Changes Done														Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				10/01/2011			           1.0																																					Sunny R
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_BMIDE_VerifyLogs(strLogText)
   Dim strCurrLog
	Fn_BMIDE_VerifyLogs=False
	strCurrLog=JavaWindow("Business Modeler").JavaEdit("LogEditBox").GetROProperty("value")
	If InStr(1,LCase(Trim(strCurrLog)),LCase(Trim(strLogText)))>=1 Then
		Fn_BMIDE_VerifyLogs=True
	End If
End Function

'-------------------------------------------------------------------Function Used to Create New Operational Data Project------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_CreateOperationalDataProject

'Description			 :	Function Used to Create New Operational Data Project

'Parameters			   :	1.strServerProfile: Server Profile Name
										'2.strPassword:Password For Connection
										'3.strGroup: Group Name
										'4.strRole:Role
										'5.strTemplate:Template Name

'Return Value		   : 	True Or False

'Pre-requisite			:	Should be Log In BMIDE

'Examples				: 	'1. Call Fn_BMIDE_CreateOperationalDataProject("DemoProfile","AutoTestDBA","","","demobatchproject (DemoBatchProject)")

'History					 :		
'													Developer Name												Date						Rev. No.						Changes Done														Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				11/01/2011			           1.0																																					Sunny R
'													Sandeep N										   				18/11/2011			           1.1				As Per Design Change Modified Node name												 Sunny R
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_BMIDE_CreateOperationalDataProject(strServerProfile,strPassword,strGroup,strRole,strTemplate)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_CreateOperationalDataProject"
   Dim iCounter,strLastItem,arrTemplate
   Dim ObjOpsDtProDialog
   Set ObjOpsDtProDialog=JavaWindow("Business Modeler").JavaWindow("NewOperationalDataProject")
   
	If Not ObjOpsDtProDialog.Exist(5) Then
		'Taking Menu Name from Environmet File
		strMenu=Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\BMIDEConfig\BMIDE_Menu.xml", "NewProject")
		'Calling File:New:Project... Menu to open Project Dialog
        Call Fn_BMIDE_MenuOperation("Select", strMenu)
	End If
	'Expanding Business Modeler IDE Node
    Call Fn_UI_JavaTree_Expand("Fn_BMIDE_CreateOperationalDataProject", JavaWindow("Business Modeler").JavaWindow("NewProject"), "WizardsTree","Business Modeler IDE")
	'Selecting Project
	Call Fn_JavaTree_Select("Fn_BMIDE_CreateOperationalDataProject", JavaWindow("Business Modeler").JavaWindow("NewProject"),"WizardsTree","Business Modeler IDE:New Live Update Project")
	'Clicking On Next button to Go Next Wizard
	Call Fn_Button_Click("Fn_BMIDE_CreateOperationalDataProject", JavaWindow("Business Modeler").JavaWindow("NewProject"), "Next")
	If strServerProfile<>"" Then
		'Selecting Server Profile
		Call Fn_List_Select("Fn_BMIDE_CreateOperationalDataProject", ObjOpsDtProDialog, "ServerProfile",strServerProfile)
	End If
	If strPassword<>"" Then
		If  ObjOpsDtProDialog.JavaEdit("Password").Exist(5) Then
            'Setting Password
			Call Fn_Edit_Box("Fn_BMIDE_CreateOperationalDataProject",ObjOpsDtProDialog,"Password",strPassword)
			wait 0,500
		End If
	End If
	If strGroup<>"" Then
		If  ObjOpsDtProDialog.JavaEdit("Group").Exist(5) Then
			'Setting Group
			Call Fn_Edit_Box("Fn_BMIDE_CreateOperationalDataProject",ObjOpsDtProDialog,"Group",strGroup)
		End IF
	End If
	If strRole<>"" Then
		If ObjOpsDtProDialog.JavaEdit("Role").Exist(5) Then
			'Setting Role
			Call Fn_Edit_Box("Fn_BMIDE_CreateOperationalDataProject",ObjOpsDtProDialog,"Role",strRole)
		End IF
	End If
	If ObjOpsDtProDialog.JavaButton("Connect").Exist(9) Then
		Call Fn_Button_Click("Fn_BMIDE_CreateOperationalDataProject", ObjOpsDtProDialog, "Connect")
	End If
    ObjOpsDtProDialog.JavaButton("Finish").WaitProperty "enabled",50000
	If Cint(ObjOpsDtProDialog.JavaButton("Next").GetROProperty("enabled"))=1 Then
		Call Fn_Button_Click("Fn_BMIDE_CreateOperationalDataProject", ObjOpsDtProDialog, "Next")
		If strTemplate<>"" Then
			'Selecting Template
			Call Fn_List_Select("Fn_BMIDE_CreateOperationalDataProject", ObjOpsDtProDialog, "Template",strTemplate)
		End If
	End If
	Call Fn_Button_Click("Fn_BMIDE_CreateOperationalDataProject", ObjOpsDtProDialog, "Finish")
	For iCounter=0 To 9
		If JavaWindow("Business Modeler").JavaWindow("NewProject").JavaWindow("ProgressInformation").Exist(10) Then
			wait(10)
		Else
			Exit For
		End If
	Next
	For iCounter=0 To 3
		If JavaWindow("Business Modeler").JavaWindow("OpenAssociatedPerspective").Exist(20) Then
			Call Fn_Button_Click("Fn_BMIDE_CreateOperationalDataProject", JavaWindow("Business Modeler").JavaWindow("OpenAssociatedPerspective"), "No")
			Exit For
		End If
	Next
	If JavaWindow("Business Modeler").JavaWindow("NewProject").JavaWindow("ViewConfiguration").Exist(20) Then
		Call Fn_Button_Click("Fn_BMIDE_CreateOperationalDataProject", JavaWindow("Business Modeler").JavaWindow("NewProject").JavaWindow("ViewConfiguration"), "OK")
		wait 2
	End If
'	strLastItem=JavaWindow("Business Modeler").JavaTree("BusinessObjectTree").GetROProperty("items count")
'    Fn_BMIDE_CreateOperationalDataProject=JavaWindow("Business Modeler").JavaTree("BusinessObjectTree").GetItem(strLastItem-1)
	 Call Fn_BMIDE_TabOperations("UpperLeft","Activate","Navigator")
	 Call Fn_BMIDE_TreeIndexIdentification()

	  arrTemplate=Split(strTemplate,"(")
	  For iCounter=0 To 10
			arrTemplate(0)=Replace(arrTemplate(0),"_","")
  			bFlag=False
			If iCounter=0 Then
				bFlag=Fn_BMIDE_BusinessObjectTreeOperations("Exist",LCase(Trim(arrTemplate(0)))+"_live_update",sPopupMenu)
			Else
				bFlag=Fn_BMIDE_BusinessObjectTreeOperations("Exist",LCase(Trim(arrTemplate(0)))+"_live_update"+CStr(iCounter),sPopupMenu)
			End If
			If bFlag=False Then
				If iCounter=0 Or iCounter=1 Then
					Fn_BMIDE_CreateOperationalDataProject=LCase(Trim(arrTemplate(0)))+"_live_update"
				Else
					Fn_BMIDE_CreateOperationalDataProject=LCase(Trim(arrTemplate(0)))+"_live_update"+CStr(iCounter-1)
				End If
				Exit for
			End If
	  Next

	    Call Fn_BMIDE_TabOperations("LowerLeft","Activate","Extensions")
	    Call Fn_BMIDE_TabOperations("UpperLeft","Activate","Business Objects")

		Call Fn_BMIDE_TreeIndexIdentification() 
End Function

'-------------------------------------------------------------------Function Used to Create New Classic Change--------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_CreateClassicChange

'Description			 :	Function Used to Create New Classic Change

'Parameters			   :	1.strName: Classic Change Name
										'2.strDispName:Classic Change Display Name
										'3.strDesc: Classic Change Description
										'4.strChangeID:Change ID
										'5.strRevID:Revision ID
										'6.strForms: Forms
										'7.strTemplate : Process Template
										'8strConnSettingCrdntial : Connection Settings

'Return Value		   : 	True Or False

'Pre-requisite			:	New Classic Change Dialog Should be Open

'Examples				: 	'1. Call Fn_BMIDE_CreateClassicChange("TestClssChange","","Test Classic Change","ON","1:2:Running~2:2:Static","3:147:Running","BOMChange Form","AutoDoDo","DemoBatchProject:DemoProfile::AutoTestDBA:dba:DBA")

'History					 :		
'													Developer Name												Date						Rev. No.						Changes Done														Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				12/01/2011			           1.0																																					Sunny R
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

' For 'strConnSettingCrdntial' Parameter take Ref of function 'Fn_BMIDE_TeamcenterRepositoryConnection'
'strChangeID ="Range:Value:Format~Range:Value:Format"  Eg:- "1:2:Running~2:2:Static"
'strRevID ="Range:Value:Format~Range:Value:Format"  Eg :- "3:147:Running"
Public Function Fn_BMIDE_CreateClassicChange(strName,strDispName,strDesc,bIsEffShared,strChangeID,strRevID,strForms,strTemplate,strConnSettingCrdntial)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_CreateClassicChange"
   Dim strPrefix,arrChangeID,iCounter,arrChangeIDVal,arrRevID,arrRevIDVal,arrForms,arrTemplate
   Dim ObjChageDailog,ObjChangeIDDialog
	Set ObjChageDailog=JavaWindow("Business Modeler").JavaWindow("NewClassicChange")
	Set ObjChangeIDDialog=JavaWindow("Business Modeler").JavaWindow("ChangeID")
	Fn_BMIDE_CreateClassicChange=False
	If strName<>"" Then
		strPrefix=Fn_Edit_Box_GetValue("Fn_BMIDE_CreateClassicChange",ObjChageDailog,"Name")
		strName=strPrefix+strName
		Call Fn_Edit_Box("Fn_BMIDE_CreateClassicChange",ObjChageDailog,"Name",strName)
	End If
	If strDispName<>"" Then
		Call Fn_Edit_Box("Fn_BMIDE_CreateClassicChange",ObjChageDailog,"DisplayName",strDispName)
	End If
	If strDesc<>"" Then
		Call Fn_Edit_Box("Fn_BMIDE_CreateClassicChange",ObjChageDailog,"Description",strDesc)
	End If
	If bIsEffShared<>"" Then
		Call Fn_CheckBox_Set("Fn_BMIDE_CreateClassicChange", ObjChageDailog, "IsEffectivityShared", bIsEffShared)
	End If
	Call Fn_Button_Click("Fn_BMIDE_CreateClassicChange", ObjChageDailog, "Next")

	If strChangeID<>"" Then
		arrChangeID=Split(strChangeID,"~")
		For iCounter=0 To UBound(arrChangeID)
			
			Call Fn_Button_Click("Fn_BMIDE_CreateClassicChange", ObjChageDailog, "AddChangeID")
			arrChangeIDVal=Split(arrChangeID(iCounter),":")
			If arrChangeIDVal(0)<>"" Then
				Call Fn_Edit_Box("Fn_BMIDE_CreateClassicChange",ObjChangeIDDialog,"Range",arrChangeIDVal(0))
			End If
			If arrChangeIDVal(1)<>"" Then
				Call Fn_Edit_Box("Fn_BMIDE_CreateClassicChange",ObjChangeIDDialog,"Value",arrChangeIDVal(1))
			End If
			If arrChangeIDVal(2)<>"" Then
                Call Fn_List_Select("Fn_BMIDE_CreateClassicChange", ObjChangeIDDialog, "Format",arrChangeIDVal(2))
			End If
			Call Fn_Button_Click("Fn_BMIDE_CreateClassicChange", ObjChangeIDDialog, "Finish")
		Next
	End If

	If strRevID<>"" Then
		arrRevID=Split(strRevID,"~")
		For iCounter=0 To UBound(arrRevID)
			Call Fn_Button_Click("Fn_BMIDE_CreateClassicChange", ObjChageDailog, "AddRevID")
			arrRevIDVal=Split(arrRevID(iCounter),":")
			If arrRevIDVal(0)<>"" Then
				Call Fn_Edit_Box("Fn_BMIDE_CreateClassicChange",ObjChangeIDDialog,"Range",arrRevIDVal(0))
			End If
			If arrRevIDVal(1)<>"" Then
				Call Fn_Edit_Box("Fn_BMIDE_CreateClassicChange",ObjChangeIDDialog,"Value",arrRevIDVal(1))
			End If
			If arrRevIDVal(2)<>"" Then
                Call Fn_List_Select("Fn_BMIDE_CreateClassicChange", ObjChangeIDDialog, "Format",arrRevIDVal(2))
			End If
			Call Fn_Button_Click("Fn_BMIDE_CreateClassicChange", ObjChangeIDDialog, "Finish")
		Next
	End If
	Call Fn_Button_Click("Fn_BMIDE_CreateClassicChange", ObjChageDailog, "Next")

	If strForms<>"" Then
		arrForms=Split(strForms,"~")
		For iCounter=0 To UBound(arrForms)
			Call Fn_Button_Click("Fn_BMIDE_CreateClassicChange", ObjChageDailog, "AddForms")
			Call Fn_Edit_Box("Fn_BMIDE_CreateClassicChange",JavaWindow("Business Modeler").JavaWindow("Find Business Object"),"Criteria",arrForms(iCounter))
			Call Fn_Button_Click("Fn_BMIDE_CreateClassicChange", JavaWindow("Business Modeler").JavaWindow("Find Business Object"), "OK")
		Next
	End If
	If strTemplate<>"" Then
		arrTemplate=Split(strTemplate,"~")
		For iCounter=0 To UBound(arrTemplate)
			
			Call Fn_Button_Click("Fn_BMIDE_CreateClassicChange", ObjChageDailog, "AddProcessTemp")
			wait(5)
			If JavaWindow("Business Modeler").JavaWindow("TeamcenterRepositoryConnection").Exist(10) Then
				arrConnCredential=Split(strConnSettingCrdntial,":")
				Call Fn_BMIDE_TeamcenterRepositoryConnection(arrConnCredential(0),arrConnCredential(1),arrConnCredential(2),arrConnCredential(3),arrConnCredential(4),arrConnCredential(5))
			End If
			Call Fn_Edit_Box("Fn_BMIDE_CreateClassicChange",JavaWindow("Business Modeler").JavaWindow("Find Business Object"),"Criteria",arrTemplate(iCounter))
			Call Fn_Button_Click("Fn_BMIDE_CreateClassicChange", JavaWindow("Business Modeler").JavaWindow("Find Business Object"), "OK")
		Next
	End If
	Call Fn_Button_Click("Fn_BMIDE_CreateClassicChange", ObjChageDailog, "Finish")
	Fn_BMIDE_CreateClassicChange=True
	Set ObjChageDailog=Nothing
	Set ObjChangeIDDialog=Nothing
End Function

'-------------------------------------------------------------------Function Used to Perform Operations On AccessorTable-------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_DeepCopyRuleOperations

'Description			 :	Function Used to Perform Operations On Deep Copy Rules Table

'Parameters			   :   '1.strAction:Action to Perform
										'2.strType: Accesor Type
										'3.strName: Accessor Name

'Return Value		   : 	True Or False

'Pre-requisite			:	Accesor Table Should Appear on Screen

'Examples				: 	Call Fn_BMIDE_AccessorTableOperation("Select","RoleInGroup","Designer.Engineering")
'										Call  Fn_BMIDE_AccessorTableOperation("Select","RoleInGroup","AutoRole1.AutoGrp1")
'										Call Fn_BMIDE_AccessorTableOperation("Add","","Organization:AutoGrp2:AutoRole3~Organization:AutoAdminGrp:AutoRoleDBA")
'										Call Fn_BMIDE_AccessorTableOperation("Remove","RoleInGroup","Designer.Engineering")

'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				13/01/2011			           1.0																				Sunny R
'													Sandeep N										   				16/01/2012			           1.1						Added Case "Remove"			Sunny R
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_BMIDE_AccessorTableOperation(strAction,strType,strName)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_AccessorTableOperation"
   Dim bFlag,iRowCount,iCounter,arrRowValue(1),strFullRowVal,strExpVal,arrName
   Dim ObjAccsrDialog,ObjAccsrTable
	Set ObjAccsrDialog=JavaWindow("Business Modeler").JavaWindow("AccessorDialog")
	Set ObjAccsrTable= JavaWindow("Business Modeler").JavaTable("AccessorTable")
	Fn_BMIDE_AccessorTableOperation=False
	bFlag=False
	Select Case strAction
		Case "Select"
				iRowCount=Fn_UI_Object_GetROProperty("Fn_BMIDE_AccessorTableOperation",ObjAccsrTable,"rows")
				For iCounter=0 To iRowCount-1
					arrRowValue(0)=ObjAccsrTable.GetCellData(iCounter,"Type")
					arrRowValue(1)=ObjAccsrTable.GetCellData(iCounter,"Name")

					strFullRowVal=arrRowValue(0)+arrRowValue(1)
					strExpVal=strType+strName
					If Trim(strFullRowVal)=Trim(strExpVal) Then
						ObjAccsrTable.ActivateRow iCounter
						bFlag=True
						Exit For
					End If
				Next
				If bFlag=True Then
					Fn_BMIDE_AccessorTableOperation=True
				End If
		Case "Verify"
				iRowCount=Fn_UI_Object_GetROProperty("Fn_BMIDE_AccessorTableOperation",ObjAccsrTable,"rows")
				For iCounter=0 To iRowCount-1
					arrRowValue(0)=ObjAccsrTable.GetCellData(iCounter,"Type")
					arrRowValue(1)=ObjAccsrTable.GetCellData(iCounter,"Name")

					strFullRowVal=arrRowValue(0)+arrRowValue(1)
					strExpVal=strType+strName
					If Trim(strFullRowVal)=Trim(strExpVal) Then
						bFlag=True
						Exit For
					End If
				Next
				If bFlag=True Then
					Fn_BMIDE_AccessorTableOperation=True
				End If
		Case "Add"
				If Not ObjAccsrDialog.Exist(5) Then
					arrName=Split(strName,"~")
					For iCounter=0 To UBound(arrName)
						Call Fn_Button_Click("Fn_BMIDE_AccessorTableOperation", JavaWindow("Business Modeler"), "AddAccessor")
						Call Fn_JavaTree_Select("Fn_BMIDE_AccessorTableOperation", ObjAccsrDialog, "AccessorTree",arrName(iCounter))
						Call Fn_Button_Click("Fn_BMIDE_AccessorTableOperation", ObjAccsrDialog, "Finish")
					Next
				Else	
					arrName=Split(strName,"~")
					For iCounter=0 To UBound(arrName)
						Call Fn_Button_Click("Fn_BMIDE_AccessorTableOperation", JavaWindow("Business Modeler"), "AddAccessor")
						Call Fn_JavaTree_Select("Fn_BMIDE_AccessorTableOperation", ObjAccsrDialog, "AccessorTree",arrName(iCounter))
						Call Fn_Button_Click("Fn_BMIDE_AccessorTableOperation", ObjAccsrDialog, "Finish")
					Next
				End If	
				Fn_BMIDE_AccessorTableOperation=True
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Remove"
				bFlag=Fn_BMIDE_AccessorTableOperation("Select",strType,strName)
				If bFlag=True Then
					Call Fn_Button_Click("Fn_BMIDE_AccessorTableOperation", JavaWindow("Business Modeler"), "RemoveAccessor")
					Fn_BMIDE_AccessorTableOperation=True
				End If
	End Select
	Set ObjAccsrDialog=Nothing
	Set ObjAccsrTable=Nothing
End Function


'-------------------------------------------------------------------Function Used to Select Project to Open Deployement Page-------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_OpsProjectSelect(sProjName)

'Description			 :	Function Used to Select Project to Open Deployement Page

'Parameters			   :   '1.sProjName:Project Name
										

'Return Value		   : 	True Or False

'Pre-requisite			:	Should Log In BMIDE

'Examples				: 	Call Fn_BMIDE_OpsProjectSelect(sProjName)
'										

'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Ashok K								   				17/01/2011			           1.0																				Sandeep N
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function Fn_BMIDE_OpsProjectSelect(sProjName)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_OpsProjectSelect"
	Fn_BMIDE_OpsProjectSelect=False
	If  Not JavaWindow("Business Modeler").JavaWindow("SelectProject").Exist(10) Then
		Call Fn_BMIDE_ToolbarButtonClick("","Deployment Page")
	End If
	If sProjName <>"" Then
		'Selecting Poject From the List
		Call Fn_List_Select("Fn_BMIDE_OpsProjectSelect", JavaWindow("Business Modeler").JavaWindow("SelectProject"),"ProjectList",sProjName)
	End If
	wait(3)
   'Clicking "Ok" Button
	Call Fn_Button_Click("Fn_BMIDE_OpsProjectSelect",JavaWindow("Business Modeler").JavaWindow("SelectProject"),"OK")
	Fn_BMIDE_OpsProjectSelect=True
End Function
'-------------------------------------------------------------------Function Used to Close Project From Navigator Tab-------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_SetView

'Description			 :	Function Used to Close Project From Navigator Tab

'Parameters			   :	1.strProjectName:Project Name

'Return Value		   : 	True Or False

'Pre-requisite			:	Should be Log In BMIDE

'Examples				:Call Fn_BMIDE_CloseProject("DemoBatch25448")
'									 Call Fn_BMIDE_CloseProject("DemoBatch25448:demobactch_opsdata1")

'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done																				Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				25/1/2011			           1.0																																																	Sunny R
'													Sandeep N										   				25/8/2011			           1.1			Modified Function To handle "TemplateProjectBackup" Dialog			 Sunny R
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_BMIDE_CloseProject(strProjectName)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_CloseProject"
   Dim arrProjectName,iCounter
   Fn_BMIDE_CloseProject=False
   	Call Fn_BMIDE_TabOperations("UpperLeft","Activate","Navigator")
	Call Fn_BMIDE_TreeIndexIdentification()
	arrProjectName=Split(strProjectName,":")
	For iCounter=0 To UBound(arrProjectName)
		Call Fn_BMIDE_BusinessObjectTreeOperations("Select",arrProjectName(iCounter),"")
		Call Fn_BMIDE_BusinessObjectTreeOperations("PopupMenuSelect",arrProjectName(iCounter),"Close Project")
		If JavaWindow("Business Modeler").JavaWindow("TemplateProjectBackup").Exist(15) Then
			  Call Fn_Button_Click("Fn_BMIDE_CloseProject", JavaWindow("Business Modeler").JavaWindow("TemplateProjectBackup"), "Finish")
			  JavaWindow("Business Modeler").JavaWindow("TemplateProjectBackup").SetTOProperty "index",1
			  wait 1
			 If JavaWindow("Business Modeler").JavaWindow("TemplateProjectBackup").Exist(6) Then
				 Call Fn_Button_Click("Fn_BMIDE_CloseProject", JavaWindow("Business Modeler").JavaWindow("TemplateProjectBackup"), "OK")
			 End If
			 JavaWindow("Business Modeler").JavaWindow("TemplateProjectBackup").SetTOProperty "index",0
		End If
	Next
	wait 5
	Call Fn_BMIDE_TabOperations("UpperLeft","Activate","Business Objects")
	Fn_BMIDE_CloseProject=True
End Function

'-------------------------------------------------------------------Function Used to Deploy Project with Detail verifications-------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_DeployProjectWithDetailVerifications

''Description			:1.Function Used to Deploy Project with Detail verifications

'Return Value		   : 	True or False

'Parameters     		:	1. StrProjectName : Project Name For Deploy
										'2. StrProfile : Profile
										'3. StrUserID: User ID ()
										'4. StrPassword: Password
										'5. StrGroup: User Group Name
										'6. StrRole: User Role

'Pre-requisite			:	BMIDE Prespective should be Open.

'Examples				: Call Fn_BMIDE_DeployProjectWithDetailVerifications("TestBMIDE305","DemoProfile","","AutoTestDBA","","")

'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				27/1/2010			           1.0																						Sunny R
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_BMIDE_DeployProjectWithDetailVerifications(StrProjectName,StrProfile,StrUserID,StrPassword,StrGroup,StrRole)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_DeployProjectWithDetailVerifications"
   Dim strMenu,i,iCounter
   Dim ObjDiployWnd
   Fn_BMIDE_DeployProjectWithDetailVerifications=False
	If Not  JavaWindow("Business Modeler").JavaWindow("Deploy").Exist(8) Then
			'Checking Existance Of "NewProject" Dialog
			 If Not JavaWindow("Business Modeler").JavaWindow("NewProject").Exist(8) Then
			   'Calling File:New:Other... Menu to open New Project  Dialog
				strMenu=Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\BMIDEConfig\BMIDE_Menu.xml", "FileNewOther")
				Call Fn_BMIDE_MenuOperation("Select", strMenu)
		   End If
			'Expanding ""Business Modeler IDE"" Node From "WizardsTree"
			Call Fn_UI_JavaTree_Expand("Fn_BMIDE_DeployProjectWithDetailVerifications",JavaWindow("Business Modeler").JavaWindow("NewProject"),"WizardsTree","Business Modeler IDE")
			'Selecting "Business Modeler IDE:Deployment"" node 
			Call Fn_JavaTree_Select("Fn_BMIDE_DeployProjectWithDetailVerifications",JavaWindow("Business Modeler").JavaWindow("NewProject"),"WizardsTree","Business Modeler IDE:Deployment")
			'Clicking "Next" Button To open "Deploy" Dialog
			Call Fn_Button_Click("Fn_BMIDE_DeployProjectWithDetailVerifications", JavaWindow("Business Modeler").JavaWindow("NewProject"), "Next")
	End If
	'Creating Object Of Deploy Window
	Set ObjDiployWnd=Fn_UI_ObjectCreate("Fn_BMIDE_DeployProjectWithDetailVerifications", JavaWindow("Business Modeler").JavaWindow("Deploy"))
	If StrProjectName<>"" Then
		If ObjDiployWnd.JavaList("Project").GetROProperty("enabled") = "1" Then
			'Selecting Project For Deploy
			Call Fn_List_Select("Fn_BMIDE_DeployProjectWithDetailVerifications", ObjDiployWnd,"Project",StrProjectName)
		End If
	End If
	If StrProfile<>"" Then
		If ObjDiployWnd.JavaList("ServerProfile").GetROProperty("enabled") = "1" Then
			'Selecting Server Profile For Deploy
			Call Fn_List_Select("Fn_BMIDE_DeployProjectWithDetailVerifications", ObjDiployWnd,"ServerProfile",StrProfile)
		End If
	End If
	If ObjDiployWnd.JavaEdit("UserID").Exist(2) Then
		If ObjDiployWnd.JavaEdit("UserID").GetROProperty("enabled") = "1" Then
			If StrUserID<>"" Then
				'Setting user ID
				Call Fn_Edit_Box("Fn_BMIDE_DeployProjectWithDetailVerifications",ObjDiployWnd,"UserID",StrUserID)
			Else
				arrUser = Split(Environment.Value("TcUserDBA"),":")
				Call Fn_Edit_Box("Fn_BMIDE_DeployProjectWithDetailVerifications",ObjDiployWnd,"UserID",arrUser(0))
			End If
		End If
	End If
	If  Fn_UI_ObjectExist("Fn_BMIDE_DeployProjectWithDetailVerifications", JavaWindow("Business Modeler").JavaWindow("Deploy").JavaEdit("Password"))=True Then
		If StrPassword<>"" Then
			'Setting Password
			 Call Fn_Edit_Box("Fn_BMIDE_DeployProjectWithDetailVerifications",ObjDiployWnd,"Password",StrPassword)
		End If
	Else
	Call Fn_Button_Click("Fn_BMIDE_DeployProjectWithDetailVerifications",ObjDiployWnd, "Finish")

	If Fn_UI_ObjectExist("Fn_BMIDE_DeployProjectWithDetailVerifications",JavaWindow("Business Modeler").JavaWindow("Deploying to Teamcenter")) = True Then
		
				If Fn_UI_ObjectExist("Fn_BMIDE_DeployProjectWithDetailVerifications",JavaWindow("Business Modeler").JavaWindow("Deploying to Teamcenter").JavaObject("MainProgressBar")) = True Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "MainProgressBar exist")
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "MainProgressBar does not exist")
					Exit Function
				End If
				If Fn_UI_ObjectExist("Fn_BMIDE_DeployProjectWithDetailVerifications",JavaWindow("Business Modeler").JavaWindow("Deploying to Teamcenter").JavaButton("Details")) = True Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Detais button exist")
					Call Fn_Button_Click("Fn_BMIDE_DeployProjectWithDetailVerifications",JavaWindow("Business Modeler").JavaWindow("Deploying to Teamcenter"), "Details")
					If Fn_UI_ObjectExist("Fn_BMIDE_DeployProjectWithDetailVerifications",JavaWindow("Business Modeler").JavaWindow("Deploying to Teamcenter").JavaObject("DetailProgressBar")) = True Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "DetailProgressBar exist")
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "DetailProgressBar does not exist")
						Exit Function
					End If
				Else
					If Fn_UI_ObjectExist("Fn_BMIDE_DeployProjectWithDetailVerifications",JavaWindow("Business Modeler").JavaWindow("Deploying to Teamcenter").JavaObject("DetailProgressBar")) = True Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "DetailProgressBar exist")
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "DetailProgressBar does not exist")
						Exit Function
					End If
				End If
		Else
			Exit Function
		End If

'		  For i=0 to 22
'			If Fn_UI_ObjectExist("Fn_BMIDE_DeployProjectWithDetailVerifications",JavaWindow("Business Modeler").JavaWindow("Deploying to Teamcenter")) = True Then
'                wait(20)
'			Else
'				Exit For
'			End If
'		 Next
		Do while JavaWindow("Business Modeler").JavaWindow("Deploying to Teamcenter").Exist(2)
			wait 5
		Loop
		If  Fn_UI_ObjectExist("Fn_BMIDE_DeployProjectWithDetailVerifications",JavaWindow("Business Modeler").JavaWindow("Deployment Complete")) = True Then
			Call Fn_Button_Click("Fn_BMIDE_DeployProjectWithDetailVerifications",JavaWindow("Business Modeler").JavaWindow("Deployment Complete"), "OK")
		End If
		Fn_BMIDE_DeployProjectWithDetailVerifications=True
		Set ObjDiployWnd=Nothing
		Exit Function
	End If
	If StrGroup<>"" Then
		'Setting Password
		 Call Fn_Edit_Box("Fn_BMIDE_DeployProjectWithDetailVerifications",ObjDiployWnd,"Group",StrGroup)
	End If
	If StrRole<>"" Then
		'Setting Password
		 Call Fn_Edit_Box("Fn_BMIDE_DeployProjectWithDetailVerifications",ObjDiployWnd,"Role",StrRole)
	End If
	If Fn_UI_ObjectExist("Fn_BMIDE_DeployProjectWithDetailVerifications", ObjDiployWnd.JavaButton("Connect"))=True Then
	'Clicking "Connect" Button To Connect To The Host
	Call Fn_Button_Click("Fn_BMIDE_DeployProjectWithDetailVerifications",ObjDiployWnd, "Connect")
	End If
	JavaWindow("Business Modeler").JavaWindow("Deploy").JavaButton("Finish").WaitProperty "enabled","1",iTime
	'Clicking "Finish" Button
	Call Fn_Button_Click("Fn_BMIDE_DeployProjectWithDetailVerifications",ObjDiployWnd, "Finish")
	'Checking Existance Of Save And Deploy dialog
	If Fn_UI_ObjectExist("Fn_BMIDE_DeployProjectWithDetailVerifications",JavaWindow("Business Modeler").JavaWindow("ConfirmSaveAndDeployment"))=True Then
		'Clicking "Connect" Button To Connect To The Host
		Call Fn_Button_Click("Fn_BMIDE_DeployProjectWithDetailVerifications",JavaWindow("Business Modeler").JavaWindow("ConfirmSaveAndDeployment"), "OK")
	End If
'    For i=0 to 23
'		If Fn_UI_ObjectExist("Fn_BMIDE_DeployProjectWithDetailVerifications",JavaWindow("Business Modeler").JavaWindow("Deploying to Teamcenter")) = True Then
'			wait(20)
'		Else
'			Exit For
'		End If
'     Next
	Do while JavaWindow("Business Modeler").JavaWindow("Deploying to Teamcenter").Exist(2)
			wait 5
	Loop
	If  Fn_UI_ObjectExist("Fn_BMIDE_DeployProjectWithDetailVerifications",JavaWindow("Business Modeler").JavaWindow("Deployment Complete")) = True Then
		Call Fn_Button_Click("Fn_BMIDE_DeployProjectWithDetailVerifications",JavaWindow("Business Modeler").JavaWindow("Deployment Complete"), "OK")
	ElseIf Fn_UI_ObjectExist("Fn_BMIDE_DeployProjectWithDetailVerifications",JavaWindow("DeploymentComplete")) = True Then
		Call Fn_Button_Click("Fn_BMIDE_DeployProjectWithDetailVerifications",JavaWindow("DeploymentComplete"), "OK")
	End If
	Fn_BMIDE_DeployProjectWithDetailVerifications=True
	'Releasing Object Of Deploy Eindow
	Set ObjDiployWnd=Nothing
End Function



'--------------------------------------------------Function Used to Start\Stop Remote Services---------------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_CommandOperations

'Description			 :	Function Used to fire Commands on Command prompt

'Return Value		   : 	Log File Path Or False
'
'Parameter				:	strCommand :- Command to Run
'Pre-requisite			:	

'Examples				:  	'Call Fn_CommandOperations("bmide_comparator -h")

'History					 :		
'													Developer Name												Date						Rev. No.						Changes Done																Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				01/02/2011			           1.0																																											  Sunny R
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Public Function Fn_CommandOperations(strCommand)
'   	Dim pCmd
'	Dim strTCRoot,strTDData,arrTRRoot,strAppLoc,strFilePath,strFullCommand,arrCommand,strFileName,strBatFilePath,arrTime
'	strTCRoot=Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\BMIDE_Config.xml", "TR_ROOT")
'	strTDData=Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\BMIDE_Config.xml", "TD_ROOT")
'	arrTRRoot=Split(strTCRoot,":")
'	strAppLoc=arrTRRoot(0)+":"
'	arrTime=Split(Time(),":")
'	arrCommand=Split(strCommand," ")
'	strFileName=arrCommand(0)+CStr(Day(Now()))+CStr(MonthName(Month(Now())))+CStr(arrTime(0))+CStr(arrTime(1))+".Log"
'	strFilePath="C:\Temp\"+strFileName
'	strBatFilePath=strTDData+"\tc_profilevars.bat"
'
'	strFullCommand="cmd /C "+strAppLoc & "& Set TC_ROOT="+strTCRoot & "& Set TC_DATA="+strTDData & " & "+strBatFilePath & "&"+strCommand+" >>"+strFilePath
'	Set pCmd=CreateObject("WSCript.Shell")
'	pCmd.Run strFullCommand
'	Set pCmd=Nothing
'	wait(20)
'	Fn_CommandOperations=strFilePath
'End Function
Function Fn_CommandOperations(sCommand)

const bytesToKb = 1024

Dim objShell
Dim objFSO, objFile
Dim sDriveName
Dim sTR_ROOT, sTR_DATA, sTempDir
Dim sCMDFileName, sLogFileName
Dim arrCommand
		
		arrCommand = Split(sCommand, "-", -1,1)
		sTR_ROOT = Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\BMIDEConfig\BMIDE_Config.xml", "TR_ROOT")
		sTR_DATA = Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\BMIDEConfig\BMIDE_Config.xml", "TD_ROOT")
		
		sTempDir = "C:\Temp"
		sCMDFileName = sTempDir & "\bmide_util.cmd"
		sLogFileName = sTempDir & "\" & Trim(arrCommand(0)) & ".log"
		
		Set objFSO = CreateObject("Scripting.FileSystemObject")
		
		' Delete log file
		if objFSO.FileExists(sLogFileName) then
			objFSO.DeleteFile(sLogFileName)				
		end if
		
		' Delete cmd file
		if objFSO.FileExists(sCMDFileName) then
			objFSO.DeleteFile(sCMDFileName)				
		end if
		
		' Create CMD file
		Set objFile = objFSO.CreateTextFile(sCMDFileName, True)
		sDriveName = objFSO.GetDriveName(sTR_ROOT)
		objFile.WriteLine sDriveName
		objFile.WriteLine "Set TC_ROOT=" & sTR_ROOT
		objFile.WriteLine "Set TC_DATA=" & sTR_DATA
		objFile.WriteLine "cd %TC_DATA%"
		objFile.WriteLine "Call tc_profilevars.bat"
		objFile.WriteLine "cd %TC_ROOT%"
		objFile.WriteLine sCommand & " > " & sLogFileName
		objFile.Close
		Set objFile = Nothing 
		
		' Creating Shell object to run cmd file'
		Set objShell = CreateObject("WScript.Shell")
		objShell.Run "%comspec% /c " & sCMDFileName, 2, True
		Set objShell = Nothing
		
		' Verifying the log file size, if its greater than 0, it return log file path, else it returns False'
		Set objFile =  objFSO.GetFile(sLogFileName)
		if cint(objFile.Size) > 0 then
			Fn_CommandOperations = sLogFileName	
		else
			Fn_CommandOperations = False	
		end if
		
		Set objFile = Nothing
		Set objFSO = Nothing

End Function


'-------------------------------------------------------------------Function Used to Set Colors------------------------------------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_ColorOperations

''Description			:1. Function Used to Set Colors

'Return Value		   : 	True or False

'Parameters     		:	1. strAction : Action Name
										'2. strRed :Red Color Band
										'3. strGreen: Green Color Band
										'4. strBlue: Blue Color Band
										'5. strHue: Hue
										'6. strSat: Sat
										'7. strLum:Lum

'Pre-requisite			:	BMIDE Prespective should be Open.

'Examples				: Call Fn_BMIDE_ColorOperations("Merged","254","242","237","12","221","231")
'									  Call Fn_BMIDE_ColorOperations("UserActionRequired","254","242","237","12","221","231")
'									  Call Fn_BMIDE_ColorOperations("NoUserActionRequired","254","242","237","12","221","231")
'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				7/2/2011			           1.0																						Sunny R
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_BMIDE_ColorOperations(strAction,strRed,strGreen,strBlue,strHue,strSat,strLum)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_ColorOperations"
   Fn_BMIDE_ColorOperations=False
   Dim ObjPrefDialog,ObjColorDialog

	Set ObjPrefDialog=JavaWindow("Business Modeler").JavaWindow("Preferences")
	Set ObjColorDialog=JavaWindow("Business Modeler").JavaWindow("Preferences").Dialog("Color")

    If Not ObjPrefDialog.Exist(7) Then
	   'Calling Window:Preference Menu to open Preference Dialog
		strMenu=Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\BMIDEConfig\BMIDE_Menu.xml", "Preferences")
		Call Fn_BMIDE_MenuOperation("Select", strMenu)
   End If
   'Expanding "Teamcenter" Node From "PreferenceTree"
	Call Fn_UI_JavaTree_Expand("Fn_BMIDE_ColorOperations",ObjPrefDialog,"PreferenceTree","Teamcenter")
	'Selecting "Teamcenter:Data Model Merge / Compare Tool" node 
	Call Fn_JavaTree_Select("Fn_BMIDE_ColorOperations",ObjPrefDialog,"PreferenceTree","Teamcenter:Data Model Merge / Compare Tool")
	Select Case strAction
		Case "UserActionRequired"
			Call Fn_Button_Click("Fn_BMIDE_ColorOperations", ObjPrefDialog, "UserActionRequired")
		Case "NoUserActionRequired"
			Call Fn_Button_Click("Fn_BMIDE_ColorOperations", ObjPrefDialog, "NoUserActionRequired")
		Case "Merged"
			Call Fn_Button_Click("Fn_BMIDE_ColorOperations", ObjPrefDialog, "Merged")
	End Select
	Call Fn_UI_WinButton_Click("Fn_BMIDE_ColorOperations", ObjColorDialog, "DefineCustomColors","","","")
	If strRed<>"" Then
		wait(1)
		ObjColorDialog.WinEdit("Red").Set strRed
	End If

	If strGreen<>"" Then
		wait(1)
		ObjColorDialog.WinEdit("Green").Set strGreen
	End If

	If strBlue<>"" Then
		wait(1)
		ObjColorDialog.WinEdit("Blue").Set strBlue
	End If

	If strHue<>"" Then
		wait(1)
		ObjColorDialog.WinEdit("Hue").Set strHue
	End If

	If strSat<>"" Then
		wait(1)
		ObjColorDialog.WinEdit("Sat").Set strSat
	End If

	If strLum<>"" Then
		wait(1)
		ObjColorDialog.WinEdit("Lum").Set strLum
	End If

	Call Fn_UI_WinButton_Click("Fn_BMIDE_ColorOperations", ObjColorDialog, "OK","","","")
	Call Fn_Button_Click("Fn_BMIDE_ColorOperations", ObjPrefDialog, "OK")
	Fn_BMIDE_ColorOperations=True
	Set ObjColorDialog=Nothing
	Set ObjPrefDialog=Nothing
End Function

'-------------------------------------------------------------------Function Used to ncorporateLatestOperationalDataChanges on Project------------------------------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_IncorporateLatestOperationalDataChanges

''Description			:1.Function Used to Incorporate Latest Operational Data Changes

'Return Value		   : 	True or False

'Parameters     		:	1. StrProjectName : Project Name For Incorporate Latest Operational Data Changes
										'2. StrProfile : Profile
										'3. strDatabaseSite: "TeamcenterServer / OperationalDataZip"
										'4. StrPassword: Password
										'5. StrGroup: User Group Name
										'6. StrRole: User Role

'Pre-requisite			:	BMIDE Prespective should be Open.

'Examples				:   Call Fn_BMIDE_IncorporateLatestOperationalDataChanges("REG_LiveUpdate_BMIDETemplate","DemoProfile","TeamcenterServer","AutoTestDBA","dba","DBA")
'									Call Fn_BMIDE_IncorporateLatestOperationalDataChanges("REG_LiveUpdate_BMIDETemplate","DemoProfile","OperationalDataZip","AutoTestDBA","dba","DBA")

'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done					   Reviewer
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Pranav Ingle								   				08/02/2011			           1.0																			     Sandeep N
'													Pranav Ingle								   				16/01/2012			           1.1							"Modified for TC9.1"					Sandeep N
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_BMIDE_IncorporateLatestOperationalDataChanges(StrProjectName,StrProfile,strDatabaseSite,StrPassword,StrGroup,StrRole)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_IncorporateLatestOperationalDataChanges"
   Dim strMenu,i
   Dim ObjOpDataChgWnd
   Fn_BMIDE_IncorporateLatestOperationalDataChanges=False

   'Checking Existance Of "Incorporate Latest Live update Changes" Dialog
	If Fn_UI_ObjectExist("Fn_BMIDE_IncorporateLatestOperationalDataChanges", JavaWindow("Business Modeler").JavaWindow("IncorporateLatestOperationalDataChanges"))=False Then
		strMenu=Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\BMIDEConfig\BMIDE_Menu.xml", "IncorporateLatestLiveUpdateChanges")
		Call Fn_BMIDE_MenuOperation("Select", strMenu)
	End If

	'Creating Object Of ncorporate Latest Live Update Changes Window
	Set ObjOpDataChgWnd=Fn_UI_ObjectCreate("Fn_BMIDE_IncorporateLatestOperationalDataChanges", JavaWindow("Business Modeler").JavaWindow("IncorporateLatestOperationalDataChanges"))

	If StrProjectName<>"" Then
		'Selecting Project For ncorporate Latest Operational Data Changes
		Call Fn_List_Select("Fn_BMIDE_IncorporateLatestOperationalDataChanges", ObjOpDataChgWnd,"Project",StrProjectName)
	End If
	If strDatabaseSite <>"" Then
		If  strDatabaseSite="TeamcenterServer" Then
			Call Fn_UI_JavaRadioButton_SetON("Fn_BMIDE_IncorporateLatestOperationalDataChanges",ObjOpDataChgWnd, "TeamcenterServer")
		End If
		If  strDatabaseSite="OperationalDataZip" Then
			Call Fn_UI_JavaRadioButton_SetON("Fn_BMIDE_IncorporateLatestOperationalDataChanges",ObjOpDataChgWnd, "OperationalDataZip")
		End If
	End If
	
	Wait 15
	Call Fn_Button_Click("Fn_BMIDE_IncorporateLatestOperationalDataChanges",ObjOpDataChgWnd, "Next")
	
	Wait 15
	If StrProfile<>"" Then
		'Selecting Server Profile For ncorporate Latest Operational Data Changes
		Call Fn_List_Select("Fn_BMIDE_IncorporateLatestOperationalDataChanges", ObjOpDataChgWnd,"ServerProfile",StrProfile)
	End If
    
	'Setting Password
	If StrPassword<>"" Then
		If  Fn_UI_ObjectExist("Fn_BMIDE_IncorporateLatestOperationalDataChanges", ObjOpDataChgWnd.JavaEdit("Password"))=True Then
			 Call Fn_Edit_Box("Fn_BMIDE_IncorporateLatestOperationalDataChanges",ObjOpDataChgWnd,"Password",StrPassword)
		End If
		'Setting Group
		If StrGroup<>"" Then
			If  Fn_UI_ObjectExist("Fn_BMIDE_IncorporateLatestOperationalDataChanges", ObjOpDataChgWnd.JavaEdit("Group"))=True Then
				If trim(lcase(StrGroup)) = "null" Then StrGroup = ""
				 Call Fn_Edit_Box("Fn_BMIDE_IncorporateLatestOperationalDataChanges",ObjOpDataChgWnd,"Group",StrGroup)
			End If
		End If
		'Setting Role
		If StrRole<>"" Then
			If  Fn_UI_ObjectExist("Fn_BMIDE_IncorporateLatestOperationalDataChanges", ObjOpDataChgWnd.JavaEdit("Role"))=True Then
				 If trim(lcase(StrRole)) = "null" Then StrRole = ""
				 Call Fn_Edit_Box("Fn_BMIDE_IncorporateLatestOperationalDataChanges",ObjOpDataChgWnd,"Role",StrRole)
			End If
		End If
		Wait 15
		
		If Fn_UI_ObjectExist("Fn_BMIDE_IncorporateLatestOperationalDataChanges", ObjOpDataChgWnd.JavaButton("Connect"))=True Then
		'Clicking "Connect" Button To Connect To The Host
		Call Fn_Button_Click("Fn_BMIDE_IncorporateLatestOperationalDataChanges",ObjOpDataChgWnd, "Connect")
		End If
	End If
	
	Wait 20
	
	If Fn_UI_ObjectExist("Fn_BMIDE_IncorporateLatestOperationalDataChanges", ObjOpDataChgWnd.JavaButton("Next"))=True Then
		Call Fn_Button_Click("Fn_BMIDE_IncorporateLatestOperationalDataChanges",ObjOpDataChgWnd, "Next")
		wait(1)
	End If
	If Fn_UI_ObjectExist("Fn_BMIDE_IncorporateLatestOperationalDataChanges", JavaWindow("Business Modeler").JavaWindow("NoSync"))=True Then
		Call Fn_Button_Click("Fn_BMIDE_IncorporateLatestOperationalDataChanges",JavaWindow("Business Modeler").JavaWindow("NoSync"), "Finish")
		wait(1)
	Else
		For i=0 to 23
			If  Fn_UI_ObjectExist("Fn_BMIDE_IncorporateLatestOperationalDataChanges", JavaWindow("Business Modeler").JavaWindow("MergeDataModel"))=True Then
				 Call Fn_Button_Click("Fn_BMIDE_IncorporateLatestOperationalDataChanges",JavaWindow("Business Modeler").JavaWindow("MergeDataModel"), "Finish")
				 wait(1)
				 If  Fn_UI_ObjectExist("Fn_BMIDE_IncorporateLatestOperationalDataChanges", JavaWindow("Business Modeler").JavaWindow("DataModelMergeDialog"))=True Then
					'Call Fn_Button_Click("Fn_BMIDE_IncorporateLatestOperationalDataChanges",JavaWindow("Business Modeler").JavaWindow("DataModelMergeResult"), "Abort")
					JavaWindow("Business Modeler").JavaWindow("DataModelMergeDialog").JavaButton("Abort").Click
					Fn_BMIDE_IncorporateLatestOperationalDataChanges=False
					Exit Function
				End If
				If  Fn_UI_ObjectExist("Fn_BMIDE_IncorporateLatestOperationalDataChanges", JavaWindow("Business Modeler").JavaWindow("DataModelMergeResult"))=True Then
					Call Fn_Button_Click("Fn_BMIDE_IncorporateLatestOperationalDataChanges",JavaWindow("Business Modeler").JavaWindow("DataModelMergeResult"), "OK")
				End If
				Exit For
			Else
				wait(10)
			End If
		 Next
	End If
	
	Fn_BMIDE_IncorporateLatestOperationalDataChanges=True
	'Releasing Object Of ncorporate Latest Operational Data Changes Eindow
	Set ObjOpDataChgWnd=Nothing
End Function 

'-------------------------------------------------------------------Function Used to Create ApplicationExtensionRule----------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_ApplicationExtensionRule

'Description			 :	Function Used to Create ApplicationExtensionRule

'Parameters			   :   1.strName:Name
'										2.strDesc:Description
'										3.strAddRuleDetails: Rule Detals seperated by  ~
'										4. strBusConSelection : Business Context Selection
'
'Return Value		   : 	True Or False

'Pre-requisite			:	New ApplicationExtensionRule Dialog Should be appear

'Examples				: 	Call Fn_BMIDE_ApplicationExtensionRule("NewExtRule","new Ext Rule","NewSiteId~OFF~OFF~OutputSite","P3NewBusContext")

'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Pranav Ingle										   			09/2/2011			           1.0																						SANDEEP 
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_BMIDE_ApplicationExtensionRule(strName,strDescription,strAddRuleDetails,strBusConSelection)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_ApplicationExtensionRule"
	'Varaible Declaration
	Dim ObjBussinessContextDialog,ObjDetails,arrAddRuletDetails,strPrefix

	Fn_BMIDE_ApplicationExtensionRule=False
	'Creating Object Of Dialogs
	Set ObjBussinessContextDialog=JavaWindow("Business Modeler").JavaWindow("NewApplicationExtensionPoint")
    If strName <> ""Then
			'Setting Name To New Bussiness Context
			strPrefix=Fn_Edit_Box_GetValue("Fn_BMIDE_ApplicationExtensionRule",ObjBussinessContextDialog,"Name")
			strName=strPrefix+strName
			Call Fn_Edit_Box("Fn_BMIDE_ApplicationExtensionRule",ObjBussinessContextDialog,"Name",strName)
	End If
    If  strDescription <> ""Then
			Call Fn_Edit_Box("Fn_BMIDE_ApplicationExtensionRule",ObjBussinessContextDialog,"Description",strDescription)
	End If

	Call Fn_Button_Click("Fn_BMIDE_ApplicationExtensionRule",ObjBussinessContextDialog,"Next")

'	Insert Rule Details
	If strAddRuleDetails<> "" Then
			Call Fn_Button_Click("Fn_BMIDE_ApplicationExtensionRule",ObjBussinessContextDialog,"Add")
			Set ObjDetails= JavaWindow("Business Modeler").JavaWindow("AddRuleDetails")
    		
			arrAddRuletDetails = Split(strAddRuleDetails,"~")
			If  arrAddRuletDetails(0) <> ""Then
					Call Fn_Edit_Box("Fn_BMIDE_ApplicationExtensionRule",ObjDetails,"siteId",arrAddRuletDetails(0))
			End If
			If  arrAddRuletDetails(1) <> "" Then
					Call Fn_CheckBox_Set("Fn_BMIDE_ApplicationExtensionRule",ObjDetails,"isPush", arrAddRuletDetails(1)) 
			End If
			If  arrAddRuletDetails(2) <> ""Then
					Call Fn_CheckBox_Set("Fn_BMIDE_ApplicationExtensionRule",ObjDetails,"isExport", arrAddRuletDetails(2)) 
			End If
			If  arrAddRuletDetails(3) <> ""Then
					Call Fn_Edit_Box("Fn_BMIDE_ApplicationExtensionRule",ObjDetails,"optionSetName",arrAddRuletDetails(3))
			End If
			Call Fn_Button_Click("Fn_BMIDE_ApplicationExtensionRule",ObjDetails,"Finish")
			Set ObjDetails= Nothing
	Else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail : Rule  Details can not be empty")
		Exit Function
	End If

'	Select strBusConSelection Details
	If strBusConSelection<> "" Then
			Call Fn_Button_Click("Fn_BMIDE_ApplicationExtensionRule",ObjBussinessContextDialog,"Add2")
			Call Fn_Edit_Box("Fn_BMIDE_ApplicationExtensionRule",JavaWindow("Business Modeler").JavaWindow("BusinessContextSelection"),"Project",strBusConSelection)
			Call Fn_Button_Click("Fn_BMIDE_ApplicationExtensionRule",JavaWindow("Business Modeler").JavaWindow("BusinessContextSelection"),"OK")
			Set ObjDetails= Nothing
    End If
'	Click fininsh to close the dialog
	If Fn_UI_ObjectExist("Fn_BMIDE_ApplicationExtensionRule", ObjBussinessContextDialog.JavaButton("Finish"))=True Then
		Call Fn_Button_Click("Fn_BMIDE_ApplicationExtensionRule",ObjBussinessContextDialog,"Finish")
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Pass :Function Fn_BMIDE_ApplicationExtensionRule successfully completed")
		Fn_BMIDE_ApplicationExtensionRule = true
	End If
'	Releasing Object Of ObjBussinessContextDialog Window
    Set ObjBussinessContextDialog = Nothing
End Function

'-------------------------------------------------------------------Function Used to Create ApplicationExtensionPoint----------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_ApplicationExtensionPoint

'Description			 :	Function Used to Create ApplicationExtensionPoint

'Parameters			   :   1.strId : Id  Should start with small case leter
'										2. strName:Name
'										3.strDesc:Description
'										4. strType : Type	
'										5.strInputDetails: Input Detals seperated by  ~
'										6. .strOutputDetails: Output Detals seperated by  ~
'
'Return Value		   : 	True Or False

'Pre-requisite			:	New ApplicationExtensionPoint Dialog Should be appear

'Examples				: 	Call Fn_BMIDE_ApplicationExtensionPoint("p31000","P3NewExtPoint","DecisionTable","New Ext Point","BusinessObject~Architecture","Primitive~String~Col2~Countries")
'										Call Fn_BMIDE_ApplicationExtensionPoint("p31000","P3NewExtPoint","DecisionTable","New Ext Point","Primitive~String~Col1~Countries","Primitive~String~Col2~Countries")
'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Pranav Ingle										   			09/2/2011			           1.0																						SANDEEP 
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_BMIDE_ApplicationExtensionPoint(strId,strName,strType,strDescription,strInputDetails,strOutputDetails)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_ApplicationExtensionPoint"
	'Varaible Declaration
	Dim ObjBussinessContextDialog,ObjDetails,arrInputDetails,arrOutputDetails,strPrefix

	Fn_BMIDE_ApplicationExtensionPoint=False
	'Creating Object Of Dialogs
	Set ObjBussinessContextDialog=JavaWindow("Business Modeler").JavaWindow("NewApplicationExtensionPoint")
	If strId <> "" Then
			Call Fn_Edit_Box("Fn_BMIDE_ApplicationExtensionPoint",ObjBussinessContextDialog,"Id",strId)
	End If
    If strName <> ""Then
			'Setting Name To New Bussiness Context
			strPrefix=Fn_Edit_Box_GetValue("Fn_BMIDE_ApplicationExtensionPoint",ObjBussinessContextDialog,"Name")
			strName=strPrefix+strName
			Call Fn_Edit_Box("Fn_BMIDE_ApplicationExtensionPoint",ObjBussinessContextDialog,"Name",strName)
	End If
	If  strType <> "" Then
			Call Fn_List_Select("Fn_BMIDE_ApplicationExtensionPoint",ObjBussinessContextDialog,"Type",strType)
	End If
    If  strDescription <> ""Then
			Call Fn_Edit_Box("Fn_BMIDE_ApplicationExtensionPoint",ObjBussinessContextDialog,"Description",strDescription)
	End If

	Call Fn_Button_Click("Fn_BMIDE_ApplicationExtensionPoint",ObjBussinessContextDialog,"Next")

'	Insert Input Details
	If strInputDetails<> "" Then
			Call Fn_Button_Click("Fn_BMIDE_ApplicationExtensionPoint",ObjBussinessContextDialog,"Add")
			Set ObjDetails= JavaWindow("Business Modeler").JavaWindow("InputOutputDetails")
			arrInputDetails = Split(strInputDetails,"~")
			If  arrInputDetails(0) <> ""Then
					Call Fn_List_Select("Fn_BMIDE_ApplicationExtensionPoint",ObjDetails,"Type",arrInputDetails(0))
			End If
			If  arrInputDetails(1) <> "" Then
					Call Fn_Edit_Box("Fn_BMIDE_ApplicationExtensionPoint",ObjDetails,"TypeName", arrInputDetails(1)) 
			End If

			If  arrInputDetails(0)= "Primitive" Then
					If  arrInputDetails(2) <> ""Then
					Call Fn_Edit_Box("Fn_BMIDE_ApplicationExtensionPoint",ObjDetails,"ColumnName", arrInputDetails(2)) 
					End If
					If  arrInputDetails(3) <> ""Then
							Call Fn_Edit_Box("Fn_BMIDE_ApplicationExtensionPoint",ObjDetails,"LOVName",arrInputDetails(3))
							ObjDetails.PressKey micEnter
					End If
			End If
			Call Fn_Button_Click("Fn_BMIDE_ApplicationExtensionPoint",ObjDetails,"Finish")
			Set ObjDetails= Nothing
	Else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail : Input Details can not be empty")
		Exit Function
	End If

'	Insert Input Details
	If strOutputDetails<> "" Then
			Call Fn_Button_Click("Fn_BMIDE_ApplicationExtensionPoint",ObjBussinessContextDialog,"Add2")
			Set ObjDetails= JavaWindow("Business Modeler").JavaWindow("InputOutputDetails")
			arrOutputDetails = Split(strOutputDetails,"~")
			If  arrOutputDetails(0) <> ""Then
					Call Fn_List_Select("Fn_BMIDE_ApplicationExtensionPoint",ObjDetails,"Type",arrOutputDetails(0))
			End If
			
			If  arrOutputDetails(1) <> "" Then
					Call Fn_Edit_Box("Fn_BMIDE_ApplicationExtensionPoint",ObjDetails,"TypeName", arrOutputDetails(1)) 
			End If
			
			If  arrOutputDetails(2) <> ""Then
				Call Fn_Edit_Box("Fn_BMIDE_ApplicationExtensionPoint",ObjDetails,"ColumnName", arrOutputDetails(2)) 
			End If
		
			If  arrOutputDetails(3) <> ""Then
				'Call Fn_Edit_Box("Fn_BMIDE_ApplicationExtensionPoint",ObjDetails,"LOVName",arrOutputDetails(3))     		'Code added by Nitish Bhardwaj
				ObjDetails.JavaStaticText("OutputName").SetTOProperty "label","LOV Name:"
				ObjDetails.JavaButton("Browse...").Click micLeftBtn
				Call Fn_Edit_Box("Fn_BMIDE_ApplicationExtensionPoint",JavaWindow("Business Modeler").JavaWindow("FindLOV"),"LOVName",arrOutputDetails(3))
				
				JavaWindow("Business Modeler").JavaWindow("FindLOV").JavaButton("OK").Click	
			End If
			
			Call Fn_Button_Click("Fn_BMIDE_ApplicationExtensionPoint",ObjDetails,"Finish")
			Set ObjDetails= Nothing
	Else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail : Output Details can not be empty")
		Exit Function
	End If

'	Click fininsh to close the dialog
	If Fn_UI_ObjectExist("Fn_BMIDE_ApplicationExtensionPoint", ObjBussinessContextDialog.JavaButton("Finish"))=True Then
		Call Fn_Button_Click("Fn_BMIDE_ApplicationExtensionPoint",ObjBussinessContextDialog,"Finish")
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Pass :Function Fn_BMIDE_ApplicationExtensionPoint successfully completed")
		Fn_BMIDE_ApplicationExtensionPoint = true
	End If
'	Releasing Object Of ObjBussinessContextDialog Window
    Set ObjBussinessContextDialog = Nothing
End Function

'-------------------------------------------------------------------Function Used to Create RevisionNamingRule----------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_CreateRevisionNamingRule

'Description			 :	Function Used to Create RevisionNamingRule

'Parameters			   :   1.strProjectName :Project name
'										2. strName:Name
'										3.strExclude : string value ON / OFF for checkbox
'										4. strInitialRevType : Initial rev Type value   (e.g  Alphabetic, Numeric) 
'										5. strInitialRevStart  : Intial Start Value (e.g A ,1 ) 
'										6. strInitialRevDesc : Initial Value Desc 
'										7 strSecRevType :  Sec rev Type value   (e.g  Alphabetic, Numeric) 
'										8. strSecRevStart :  Sec Start Value (e.g A ,1 ) 
'										9. strSecRevDesc  : Value Desc 
'										10.  strSuppRevType : Supplement Revision type 
'										11.  strSuppRevDesc	 : Supplement Revision Desc								
'Return Value		   : 	True Or False

'Pre-requisite			:	New RevisionNamingRule Dialog Should be appear

'Examples				: 	Fn_BMIDE_CreateRevisionNamingRule("","NewRevRule","ON","Alphabetic","A","InitRevDesc", "Numeric", "1", "SecRevDesc","", "")

'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Pranav Ingle										   			11/2/2011			           1.0																						SANDEEP 
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_BMIDE_CreateRevisionNamingRule(strProjectName,strName,strExclude,strInitialRevType,strInitialRevStart,strInitialRevDesc, strSecRevType, strSecRevStart, strSecRevDesc,strSuppRevType, strSuppRevDesc)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_CreateRevisionNamingRule"
	'Varaible Declaration
	Dim ObjBussinessContextDialog,strPrefix

	Fn_BMIDE_CreateRevisionNamingRule=False
	'Creating Object Of Dialogs
	Set ObjBussinessContextDialog=JavaWindow("Business Modeler").JavaWindow("NewRevisionNamingRules")
	If  strProjectName <> ""Then
			Call Fn_List_Select("Fn_BMIDE_CreateRevisionNamingRule",ObjBussinessContextDialog,"Project",strProjectName )
	End If
    If strName <> ""Then
			'Setting Name To New Bussiness Context
			strPrefix=Fn_Edit_Box_GetValue("Fn_BMIDE_CreateRevisionNamingRule",ObjBussinessContextDialog,"Name")
			strName=strPrefix+strName
			Call Fn_Edit_Box("Fn_BMIDE_CreateRevisionNamingRule",ObjBussinessContextDialog,"Name",strName)
	End If
    If  strExclude <> "" Then
			Call Fn_CheckBox_Set("Fn_BMIDE_CreateRevisionNamingRule",ObjBussinessContextDialog,"Exclude", strExclude) 
	End If

	If strInitialRevType<> ""Then
			Call Fn_List_Select("Fn_BMIDE_CreateRevisionNamingRule",ObjBussinessContextDialog,"InitialRevisionType",strInitialRevType)
	End If
	If strInitialRevStart<> ""Then
			Call Fn_Edit_Box("Fn_BMIDE_CreateRevisionNamingRule",ObjBussinessContextDialog,"InitialRevisionStart",strInitialRevStart)
	End If
	If strInitialRevDesc<> ""Then
			Call Fn_Edit_Box("Fn_BMIDE_CreateRevisionNamingRule",ObjBussinessContextDialog,"InitialRevisionDescription",strInitialRevDesc)
	End If	
	If strSecRevType<> ""Then
			Call Fn_List_Select("Fn_BMIDE_CreateRevisionNamingRule",ObjBussinessContextDialog,"SecondaryRevisionType",strSecRevType)
	End If
	If strSecRevStart<> ""Then
			Call Fn_Edit_Box("Fn_BMIDE_CreateRevisionNamingRule",ObjBussinessContextDialog,"SecondaryRevisionStart",strSecRevStart)
	End If
	If strSecRevDesc<> ""Then
			Call Fn_Edit_Box("Fn_BMIDE_CreateRevisionNamingRule",ObjBussinessContextDialog,"SecondaryRevisionDescription",strSecRevDesc)
	End If	
	If strSuppRevType<> ""Then
			Call Fn_List_Select("Fn_BMIDE_CreateRevisionNamingRule",ObjBussinessContextDialog,"SupplementalRevision",strSuppRevType)
	End If
	If strSuppRevDesc<> ""Then
			Call Fn_Edit_Box("Fn_BMIDE_CreateRevisionNamingRule",ObjBussinessContextDialog,"SupplementalRevision",strSuppRevDesc)
	End If	

    '	Click fininsh to close the dialog
	If Fn_UI_ObjectExist("Fn_BMIDE_CreateRevisionNamingRule", ObjBussinessContextDialog.JavaButton("Finish"))=True Then
		Call Fn_Button_Click("Fn_BMIDE_CreateRevisionNamingRule",ObjBussinessContextDialog,"Finish")
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Pass :Function Fn_BMIDE_CreateRevisionNamingRule successfully completed")
		Fn_BMIDE_CreateRevisionNamingRule = true
	End If
'	Releasing Object Of ObjBussinessContextDialog Window
    Set ObjBussinessContextDialog = Nothing
End Function

'-------------------------------------------------------------------Function Used to Create RevisionNamingRule----------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_RevisionNamingRuleAttach

'Description			 :	Function Used to Attach RevisionNamingRule

'Parameters			   :   1.strProjectName
'										2. strProperty:  Item name and Property name using " ."
'										3.strCase:case to select 
'										4.,strCondition: Condition
'										5. strOverride: Overide checkbox
'
'Return Value		   : 	True Or False

'Pre-requisite			:	New RevisionNamingRuleAttach Dialog Should be appear

'Examples				: 	Fn_BMIDE_RevisionNamingRuleAttach("","P3Item1Revision.item_revision_id","Mix","isTrue","OFF")

'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Pranav Ingle										   			11/2/2011			           1.0																						SANDEEP 
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_BMIDE_RevisionNamingRuleAttach(strProjectName,strProperty,strCase,strCondition,strOverride)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_RevisionNamingRuleAttach"
	'Varaible Declaration
	Dim ObjBussinessContextDialog,arrProperty

	Fn_BMIDE_RevisionNamingRuleAttach=False
	'Creating Object Of Dialogs
	Set ObjBussinessContextDialog=JavaWindow("Business Modeler").JavaWindow("RevisionNamingRuleAttach")
	If  strProjectName <> ""Then
			Call Fn_List_Select("Fn_BMIDE_RevisionNamingRuleAttach",ObjBussinessContextDialog,"Project",strProjectName )
	End If
	Call Fn_Button_Click("Fn_BMIDE_RevisionNamingRuleAttach",ObjBussinessContextDialog,"Browse") 

	If Fn_UI_ObjectExist("Fn_BMIDE_RevisionNamingRuleAttach", JavaWindow("Business Modeler").JavaWindow("PropertyAttachment"))=True Then
		 If strProperty <> ""Then
	   			 arrProperty = Split(strProperty,".")
				Call Fn_Edit_Box("Fn_BMIDE_RevisionNamingRuleAttach", JavaWindow("Business Modeler").JavaWindow("PropertyAttachment"),"Project",arrProperty(0))
				Call Fn_Edit_Box("Fn_BMIDE_RevisionNamingRuleAttach", JavaWindow("Business Modeler").JavaWindow("PropertyAttachment"),"Properties",arrProperty(1))
		End If
		Call Fn_Button_Click("Fn_BMIDE_RevisionNamingRuleAttach",JavaWindow("Business Modeler").JavaWindow("PropertyAttachment"),"OK")
	End If
	If  strCase <> ""Then
			Call Fn_List_Select("Fn_BMIDE_RevisionNamingRuleAttach",ObjBussinessContextDialog,"Case",strCase )
	End If
	If strCondition<> ""Then
			Call Fn_Edit_Box("Fn_BMIDE_RevisionNamingRuleAttach",ObjBussinessContextDialog,"Condition",strCondition)
	End If
	 If strExclude <> "" Then
			Call Fn_CheckBox_Set("Fn_BMIDE_RevisionNamingRuleAttach",ObjBussinessContextDialog,"Override", strExclude) 
	End If

    '	Click fininsh to close the dialog
	If Fn_UI_ObjectExist("Fn_BMIDE_RevisionNamingRuleAttach", ObjBussinessContextDialog.JavaButton("Finish"))=True Then
		Call Fn_Button_Click("Fn_BMIDE_RevisionNamingRuleAttach",ObjBussinessContextDialog,"Finish")
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Pass :Function Fn_BMIDE_RevisionNamingRuleAttach successfully completed")
		Fn_BMIDE_RevisionNamingRuleAttach = true
	End If
'	Releasing Object Of ObjBussinessContextDialog Window
    Set ObjBussinessContextDialog = Nothing
End Function


'-------------------------------------------------------------------Function Used to Create New Global Constant----------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_CreateGlobalConstantExt

'Description			 :	Function Used to Create New Global Constant

'Parameters			   :	1.strName: New Global Constant Name
										'2.strDesc: Global Constant Description
										'3.strDataType:Data Type
										'4.bMultiValued: Is Multi Valued Option
										'5.strDefaultValue:Default Values
										'6.strValues:Values
										'7.bAttachmentSecured:Attachment Secured Option

'Return Value		   : 	True Or False

'Pre-requisite			:	New Global COnstant Dialog Should be Appear on Screen

'Examples				: 	Call Fn_BMIDE_CreateGlobalConstantExt("Demo","Demo Global Constant","String","Off","Test","","","ON","ON")
'										Call Fn_BMIDE_CreateGlobalConstantExt("Demo1","Demo1 Global Constant","String","ON","DVal1:DVal2:DVal3","","ON","","")
'										Call Fn_BMIDE_CreateGlobalConstantExt("Demo2","Demo2 Global Constant","Boolean","","true","","")
'										Call Fn_BMIDE_CreateGlobalConstantExt("Demo3","Demo3 Global Constant","List","","Val1","Val1:On~Val2:Off~Val3:On","","OFF","ON")
'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done								Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				17/2/2011			           1.0																								Sunny R
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_BMIDE_CreateGlobalConstantExt(strName,strDesc,strDataType,bMultiValued,strDefaultValue,strValues,bAttachmentSecured,bAllowOpsData,bAllowOpsDataOvrr)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_CreateGlobalConstantExt"
   'Variable Declaration
   Dim strPrefix,arrDefaultValues,iCounter,arrValues,arrValueSet
   Dim ObjGlobalConstDialog
	Fn_BMIDE_CreateGlobalConstantExt=False
	'Checking Existance Of "NewGlobalConstants" Window
	If Fn_UI_ObjectExist("Fn_BMIDE_CreateGlobalConstantExt",JavaWindow("Business Modeler").JavaWindow("NewGlobalConstants"))=True Then
		'Creating Object Of "NewGlobalConstants" Window
		Set ObjGlobalConstDialog=Fn_UI_ObjectCreate("Fn_BMIDE_CreateGlobalConstantExt",JavaWindow("Business Modeler").JavaWindow("NewGlobalConstants"))
	Else
	    Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail: NewGlobalConstants Dialog Is Not Exist")
		Exit Function
	End If
	'Retriving Prefix From "Name" Edit Box
	strPrefix=Fn_Edit_Box_GetValue("Fn_BMIDE_CreateGlobalConstantExt",ObjGlobalConstDialog,"Name")
	'Attaching Prefix To Name
	strName=strPrefix+strName
	'Setting Name To Global Constant
	Call Fn_Edit_Box("Fn_BMIDE_CreateGlobalConstantExt",ObjGlobalConstDialog,"Name",strName)
	If strDesc<>"" Then
		'Setting Description To New Global Constant
		Call Fn_Edit_Box("Fn_BMIDE_CreateGlobalConstantExt",ObjGlobalConstDialog,"Description",strDesc)
	End If
	If strDataType<>"" Then
		'Selecting Data Type
		Call Fn_List_Select("Fn_BMIDE_CreateGlobalConstantExt",ObjGlobalConstDialog,"DataType",strDataType)
	End If
	If LCase(strDataType)="string" Then
		If Trim(UCase(bMultiValued))="ON" Then
			'Setting Status Of "Is Multi Valued?" Option To ON
			Call Fn_CheckBox_Set("Fn_BMIDE_CreateGlobalConstantExt",ObjGlobalConstDialog,"IsMultiValued","On")
			If strDefaultValue<>"" Then
				arrDefaultValues=Split(strDefaultValue,":")
				'Adding Multiple Default Values
				For iCounter=0 To Ubound(arrDefaultValues)
					'Clicking "Add" Button to Invoke "AddValues" Dialog
					Call Fn_Button_Click("Fn_BMIDE_CreateGlobalConstantExt",ObjGlobalConstDialog,"Add")
					Call Fn_Edit_Box("Fn_BMIDE_CreateGlobalConstantExt",JavaWindow("Business Modeler").JavaWindow("AddGlobalConstValue"),"Value",arrDefaultValues(iCounter))
					Call Fn_Button_Click("Fn_BMIDE_CreateGlobalConstantExt",JavaWindow("Business Modeler").JavaWindow("AddGlobalConstValue"),"Finish")
				Next
			End If
			'Setting Status Of "Is Attachment Secured?" Option
			If bAttachmentSecured<>"" Then
				Call Fn_CheckBox_Set("Fn_BMIDE_CreateGlobalConstantExt",ObjGlobalConstDialog,"IsAttachmentSecured",bAttachmentSecured)
			End If
		Else
			'Setting Status Of "Is Multi Valued?" Option To OFF
			Call Fn_CheckBox_Set("Fn_BMIDE_CreateGlobalConstantExt",ObjGlobalConstDialog,"IsMultiValued","Off")
			'Setting Default Value
			If strDefaultValue<>"" Then
				Call Fn_Edit_Box("Fn_BMIDE_CreateGlobalConstantExt",ObjGlobalConstDialog,"DefaultValue",strDefaultValue)
			End If
		End If
	ElseIf LCase(strDataType)="boolean" Then
		'Setting Default Value
			If strDefaultValue<>"" Then
				Call Fn_List_Select("Fn_BMIDE_CreateGlobalConstantExt",ObjGlobalConstDialog,"DefaultValueList",strDefaultValue)
			End If
	ElseIf LCase(strDataType)="list" Then
			'Setting Multiple Values
			If strValues<>"" Then
				'Multiple Values Separeted by (~) Tilda
				arrValues=Split(strValues,"~")
				For iCounter=0 To Ubound(arrValues)
					arrValueSet=Split(arrValues(iCounter),":")
					'Clicking "Add" Button to Invoke "AddValues" Dialog
					Call Fn_Button_Click("Fn_BMIDE_CreateGlobalConstantExt",ObjGlobalConstDialog,"Add")
					Call Fn_Edit_Box("Fn_BMIDE_CreateGlobalConstantExt",JavaWindow("Business Modeler").JavaWindow("AddGlobalConstValue"),"Value",arrValueSet(0))
					If arrValueSet(1)<>"" Then
						Call Fn_CheckBox_Set("Fn_BMIDE_CreateGlobalConstantExt",JavaWindow("Business Modeler").JavaWindow("AddGlobalConstValue"),"Secured",arrValueSet(1))
					End If
					Call Fn_Button_Click("Fn_BMIDE_CreateGlobalConstantExt",JavaWindow("Business Modeler").JavaWindow("AddGlobalConstValue"),"Finish")
				Next
			End If
			'Setting Default Value
			If strDefaultValue<>"" Then
				Call Fn_Edit_Box("Fn_BMIDE_CreateGlobalConstantExt",ObjGlobalConstDialog,"DefaultValue",strDefaultValue)
			End If
	End If

	If bAllowOpsData<>"" Then
		ObjGlobalConstDialog.JavaCheckBox("IsAttachmentSecured").SetTOProperty "attached text","Allow Operational Data Updates?"
		Call Fn_CheckBox_Set("Fn_BMIDE_CreateGlobalConstantExt",ObjGlobalConstDialog,"IsAttachmentSecured",bAllowOpsData)
	End If
	If bAllowOpsDataOvrr<>"" Then
		ObjGlobalConstDialog.JavaCheckBox("IsAttachmentSecured").SetTOProperty "attached text","Allow Operational Data Updates to the Constant Override?"
		Call Fn_CheckBox_Set("Fn_BMIDE_CreateGlobalConstantExt",ObjGlobalConstDialog,"IsAttachmentSecured",bAllowOpsData)
	End If

	'Clicking "Finish" Button to Create New Global Constants
	Call Fn_Button_Click("Fn_BMIDE_CreateGlobalConstantExt",ObjGlobalConstDialog,"Finish")
	Fn_BMIDE_CreateGlobalConstantExt=True
	'Releasing Object Of "Global Constants Dialog"
	Set ObjGlobalConstDialog=Nothing
End Function


'-------------------------------------------------------------------Function Used to Create New Tool------------------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_CreateNewTool

'Description			 :	Function Used to Create RevisionNamingRule

'Parameters			   :   1.strName :Tool name
'										2. strMIMETYPE:Mime Type
'										3.strShellSymbol : Shell Symbol
'										4. strVendorName : Vendor Name
'										5. strRevision  : Revision
'										6. strReleaseDate :Release Date
'										7 strDescription :  Description
'										8. strInput :  Input
'										9. strOutput  :Output
'										10.  strMacLunchComm :Mac Lunch Command
'										11.  strWinLunchComm	 : Win Lucnch Command			
'										12. strCheckBoxOptions: Check Boxes				
'Return Value		   : 	True Or False

'Pre-requisite			:	New Tool Dialog Should be appear

'Examples				: 	Fn_BMIDE_CreateNewTool("Test","TestMIME","DemoShellSym","SQS","A","","DemoTool","TestInp","TestOp","MacLnch","WinLnch","Markup Capable?:VVI Required?")
'										Fn_BMIDE_CreateNewTool("Test1","TestMIME1","DemoShellSym1","SQS","A","","DemoTool1","TestInp:TestInp1","TestOp:TestOp1","","","")

'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep										   			17/2/2011			           1.0																						Sunny
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_BMIDE_CreateNewTool(strName,strMIMETYPE,strShellSymbol,strVendorName,strRevision,strReleaseDate,strDescription,strInput,strOutput,strMacLunchComm,strWinLunchComm,strCheckBoxOptions)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_CreateNewTool"
   Dim ObjNewTool
   Dim strPrefix,arrInput,arrOutput,iCounter,iCounter1,iCount,arrCheckBoxOptions
	Fn_BMIDE_CreateNewTool=False
	Set ObjNewTool=JavaWindow("Business Modeler").JavaWindow("NewTool")
	'Checking Existance of NewTool Dialog
	If Not ObjNewTool.Exist(8) Then
		Exit Function
	End If
	
	strPrefix=Fn_Edit_Box_GetValue("Fn_BMIDE_CreateNewTool",ObjNewTool,"Name")
	strName=strPrefix+strName
	Call Fn_Edit_Box("Fn_BMIDE_CreateNewTool",ObjNewTool,"Name",strName)
	If strMIMETYPE<>"" Then
		Call Fn_Edit_Box("Fn_BMIDE_CreateNewTool",ObjNewTool,"MIMETYPE",strMIMETYPE)
	End If
	If strShellSymbol<>"" Then
		Call Fn_Edit_Box("Fn_BMIDE_CreateNewTool",ObjNewTool,"ShellSymbol",strShellSymbol)
	End If
	If strVendorName<>"" Then
		Call Fn_Edit_Box("Fn_BMIDE_CreateNewTool",ObjNewTool,"VendorName",strVendorName)
	End If
	If strRevision<>"" Then
		Call Fn_Edit_Box("Fn_BMIDE_CreateNewTool",ObjNewTool,"Revision",strRevision)
	End If
	If strDescription<>"" Then
		Call Fn_Edit_Box("Fn_BMIDE_CreateNewTool",ObjNewTool,"Description",strDescription)
	End If

	Call Fn_Button_Click("Fn_BMIDE_CreateNewTool", ObjNewTool, "Next")
	If strInput<>"" Then
		arrInput=Split(strInput,":")
		For iCounter=0 To UBound(arrInput)
			Call Fn_Button_Click("Fn_BMIDE_CreateNewTool", ObjNewTool, "AddInput")
			Call Fn_Edit_Box("Fn_BMIDE_CreateNewTool",JavaWindow("Business Modeler").JavaWindow("AddDefaultValue"),"DefaultValue",arrInput(iCounter))
			Call Fn_Button_Click("Fn_BMIDE_CreateNewTool", JavaWindow("Business Modeler").JavaWindow("AddDefaultValue"), "Finish")
		Next
	End If

	If strOutput<>"" Then
		arrOutput=Split(strOutput,":")
		For iCounter1=0 To UBound(arrOutput)
			Call Fn_Button_Click("Fn_BMIDE_CreateNewTool", ObjNewTool, "AddOutput")
			Call Fn_Edit_Box("Fn_BMIDE_CreateNewTool",JavaWindow("Business Modeler").JavaWindow("AddDefaultValue"),"DefaultValue",arrOutput(iCounter1))
			Call Fn_Button_Click("Fn_BMIDE_CreateNewTool", JavaWindow("Business Modeler").JavaWindow("AddDefaultValue"), "Finish")
		Next
	End If

	Call Fn_Button_Click("Fn_BMIDE_CreateNewTool", ObjNewTool, "Next")
	If  strMacLunchComm<>"" Then
		Call Fn_Edit_Box("Fn_BMIDE_CreateNewTool",ObjNewTool,"MacLaunchCommand",strMacLunchComm)
	End If
	If  strWinLunchComm<>"" Then
		Call Fn_Edit_Box("Fn_BMIDE_CreateNewTool",ObjNewTool,"WinLaunchCommand",strWinLunchComm)
	End If
	If strCheckBoxOptions<>"" Then
		arrCheckBoxOptions=Split(strCheckBoxOptions,":")
		For iCount=0 To UBound(arrCheckBoxOptions)
			ObjNewTool.JavaCheckBox("CheckBoxes").SetTOProperty "attached text",arrCheckBoxOptions(iCount)
			ObjNewTool.JavaCheckBox("CheckBoxes").Set "ON"
		Next
	End If
	Call Fn_Button_Click("Fn_BMIDE_CreateNewTool", ObjNewTool, "Finish")
	Fn_BMIDE_CreateNewTool=True
	Set ObjNewTool=Nothing
End Function

'-------------------------------------------------------------------Function Used to Create New Storage Media------------------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_CreateNewStorageMedia

'Description			 :	Function Used to Create New Storage Media

'Parameters			   :   1.strName :Storage Media name
'										2. strMediaType:Media Type
'										3.strLogicalDevice :Logical Device
'										4. strDescription :Description

'Return Value		   : 	True Or False

'Pre-requisite			:	New Storage Media Dialog Should be appear

'Examples				: Fn_BMIDE_CreateNewStorageMedia("MassStorage","Disk","360GBHardDisk","hardDiskWithSata")

'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep										   			17/2/2011			           1.0																						Sunny
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_BMIDE_CreateNewStorageMedia(strName,strMediaType,strLogicalDevice,strDescription)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_CreateNewStorageMedia"
   Dim ObjStorageMedia
   Dim strPrefix
   Fn_BMIDE_CreateNewStorageMedia=False
	Set ObjStorageMedia=JavaWindow("Business Modeler").JavaWindow("NewStorageMedia")
	If Not ObjStorageMedia.Exist(8) Then
		Set ObjStorageMedia=Nothing
		Exit Function
	End If
	strPrefix=Fn_Edit_Box_GetValue("Fn_BMIDE_CreateNewStorageMedia",ObjStorageMedia,"Name")
	strName=strPrefix+strName
	Call Fn_Edit_Box("Fn_BMIDE_CreateNewStorageMedia",ObjStorageMedia,"Name",strName)
	If strMediaType<>"" Then
		Call Fn_List_Select("Fn_BMIDE_CreateNewStorageMedia", ObjStorageMedia, "MediaType",strMediaType)
	End If
	If strLogicalDevice<>"" Then
		Call Fn_Edit_Box("Fn_BMIDE_CreateNewStorageMedia",ObjStorageMedia,"LogicalDevice",strLogicalDevice)
	End If
	If strDescription<>"" Then
		Call Fn_Edit_Box("Fn_BMIDE_CreateNewStorageMedia",ObjStorageMedia,"Description",strDescription)
	End If
	Call Fn_Button_Click("Fn_BMIDE_CreateNewStorageMedia", ObjStorageMedia, "Finish")
	Fn_BMIDE_CreateNewStorageMedia=True
	Set ObjStorageMedia=Nothing
End Function

'----------------------------------------------------------------------------Function use to Perform Operations On Global Constants Table---------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_GlobalConstantTableOperations

'Description			 :	Function use to Perform Operations On Global Constants Table

'Parameters			   :   '1.sAction = Action to Select
'										2. strName = Global Constant Name
'										3. strValue = New Value Of Global Constant
'										4.,strColumnName = Column Name
'								'		5.,strExpVal = Expected Value

'Return Value		   : 	True Or False

'Pre-requisite			:	Global Constants Table Should be appear on srceen

'Examples				:	 		Fn_BMIDE_GlobalConstantTableOperations("Edit","L3Test","NewVal","","")
'												Fn_BMIDE_GlobalConstantTableOperations("Edit","Fnd0SelectedLocales","NewVal","","")
'												Fn_BMIDE_GlobalConstantTableOperations("Select","ProjectTopLevelSmartFolders","","","")
'											Fn_BMIDE_GlobalConstantTableOperations("Verify","T3LU78","","Value","NewValue")
'											Fn_BMIDE_GlobalConstantTableOperations("VerifyGCExist","T3LU78","","","")

'History					 :			
'													Developer Name												Date						Rev. No.			Changes Done											Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N						   						  18-Feb-2011					1.0																							  Sunny
'													Sandeep N						   						  22-Mar-2011					1.0						Case "Select"												Sunny
'													Pranav Ingle						   					   09-Jan-2012					  1.1				 	  Case "Verify"												   Sandeep
'													Sandeep N						   						  08-Feb-2012					1.2						Case "VerifyGCExist"								Sunny
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------	
Public Function Fn_BMIDE_GlobalConstantTableOperations(strAction,strName,strValue,strColumnName,strExpVal)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_GlobalConstantTableOperations"
   Dim ObjBMIDE,objGlobalConstWin
   Dim iRowCount,iCounter,strCurrName,currValue,strType

   Fn_BMIDE_GlobalConstantTableOperations=False
	Set ObjBMIDE=JavaWindow("Business Modeler")
	If Not ObjBMIDE.JavaTable("GlobalConstant").Exist(7) Then
		Set ObjBMIDE=Nothing
		Exit Function
	End If
	Select Case strAction
		Case "Edit"
			iRowCount=ObjBMIDE.JavaTable("GlobalConstant").GetROProperty("rows")
			For iCounter=0 To iRowCount-1
				strCurrName=ObjBMIDE.JavaTable("GlobalConstant").GetCellData(iCounter,"Name")
				If Trim(strName)=Trim(strCurrName) Then
					ObjBMIDE.JavaTable("GlobalConstant").SelectCell iCounter,0
'					ObjBMIDE.JavaTable("GlobalConstant").ActivateRow(iCounter)
					
				End If
			Next
			If ObjBMIDE.JavaButton("EditGlobalConstant").CheckProperty("enabled",1)=False Then
				Set ObjBMIDE=Nothing
				Exit Function
			End If
			Call Fn_Button_Click("Fn_BMIDE_GlobalConstantTableOperations", ObjBMIDE, "EditGlobalConstant")
			wait(1)
			If strValue<>"" Then
				Set objGlobalConstWin = JavaWindow("Business Modeler").JavaWindow("GlobalConstant")
				If objGlobalConstWin.JavaEdit("Type").Exist(1) = True Then
					strType = Fn_Edit_Box_GetValue("Fn_BMIDE_GlobalConstantTableOperations",objGlobalConstWin,"Type")
				     Select Case Trim(strType)
				     		Case "Boolean" 
				     			bFlag1 = False
				     			Select Case lcase(strValue)
				     				Case "on"
				     						Call Fn_UI_Object_SetTOProperty_ExistCheck("",objGlobalConstWin.JavaRadioButton("Value"),"attached text","True")
											Call Fn_UI_JavaRadioButton_SetON("Fn_BMIDE_GlobalConstantTableOperations", objGlobalConstWin, "Value")
									Case "off"
											Call Fn_UI_Object_SetTOProperty_ExistCheck("",objGlobalConstWin.JavaRadioButton("Value"),"attached text","False")
											Call Fn_UI_JavaRadioButton_SetON("Fn_BMIDE_GlobalConstantTableOperations", objGlobalConstWin, "Value")
							End Select
						Case Else
							bFlag1 = True
					End Select
				Else
					bFlag1 = True
				End If
				If bFlag1 = True Then
					If JavaWindow("Business Modeler").JavaWindow("GlobalConstant").JavaEdit("Value").Exist(2) Then
						Call Fn_Edit_Box("Fn_BMIDE_GlobalConstantTableOperations",ObjBMIDE.JavaWindow("GlobalConstant"),"Value","")
						wait(1)
						Call Fn_UI_EditBox_Type("Fn_BMIDE_GlobalConstantTableOperations",ObjBMIDE.JavaWindow("GlobalConstant"),"Value",strValue)
						wait(2)
					ElseIf JavaWindow("Business Modeler").JavaWindow("GlobalConstant").JavaCheckBox("Value").Exist(2) Then
						Call Fn_CheckBox_Set("Fn_BMIDE_GlobalConstantTableOperations", ObjBMIDE.JavaWindow("GlobalConstant"), "Value", strValue)
					ElseIf JavaWindow("Business Modeler").JavaWindow("GlobalConstant").JavaList("ValueList").Exist(2) Then
						Call Fn_SISW_UI_JavaList_Operations("Fn_BMIDE_GlobalConstantTableOperations", "Select", JavaWindow("Business Modeler").JavaWindow("GlobalConstant"),"ValueList",strValue, "", "")
					End If
				End If
				Set objGlobalConstWin = Nothing
			End If
			Call Fn_Button_Click("Fn_BMIDE_GlobalConstantTableOperations", ObjBMIDE.JavaWindow("GlobalConstant"), "Finish")
			Fn_BMIDE_GlobalConstantTableOperations=True

		Case "Select"
				iRowCount=ObjBMIDE.JavaTable("GlobalConstant").GetROProperty("rows")
				For iCounter=0 To iRowCount-1
					strCurrName=ObjBMIDE.JavaTable("GlobalConstant").GetCellData(iCounter,"Name")
					If Trim(strName)=Trim(strCurrName) Then
						ObjBMIDE.JavaTable("GlobalConstant").SelectCell iCounter,0
'						ObjBMIDE.JavaTable("GlobalConstant").ActivateRow(iCounter)
						Fn_BMIDE_GlobalConstantTableOperations=True
					End If
				Next
		Case "Verify"
					iRowCount=Fn_UI_Object_GetROProperty("Fn_BMIDE_GlobalConstantTableOperations",ObjBMIDE.JavaTable("GlobalConstant"),"rows")
					For iCounter=0 To iRowCount-1
							strCurrName=JavaWindow("Business Modeler").JavaTable("GlobalConstant").GetCellData(iCounter,"Name")	
							If Trim(strCurrName)=Trim(strName) Then
								currValue=JavaWindow("Business Modeler").JavaTable("GlobalConstant").GetCellData(iCounter,strColumnName)	
								If  Trim(currValue)=Trim(strExpVal)Then
									bFlag=True
									Exit For
								End If
							End If
					Next
					If bFlag=True Then
						Fn_BMIDE_GlobalConstantTableOperations=True
					End If
			Case "VerifyGCExist"
					iRowCount=Fn_UI_Object_GetROProperty("Fn_BMIDE_GlobalConstantTableOperations",ObjBMIDE.JavaTable("GlobalConstant"),"rows")
					For iCounter=0 To iRowCount-1
							strCurrName=JavaWindow("Business Modeler").JavaTable("GlobalConstant").GetCellData(iCounter,"Name")	
							If Trim(strCurrName)=Trim(strName) Then
                                	bFlag=True
									Exit For
							End If
					Next
					If bFlag=True Then
						Fn_BMIDE_GlobalConstantTableOperations=True
					End If
	End Select
	Set ObjBMIDE=Nothing
End Function

'-------------------------------------------------------------------Function Used to Create New Business Object Constant----------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_CreateBusinessObjectConstantExt

'Description			 :	Function Used to Create New Business Object Constant

'Parameters			   :	1.strName: New BusinessObject Constant Name
										'2.strDesc: BusinessObject Constant Description
										'3.strDataType:Data Type
										'4.strScope: Scope Values  Seperated by  "~"
										'5.strDefaultValue:Default Values
										'6.strValues:Values ot  Enter when Datatype is  'List' seperated by "~"  and value of checkbox ON/OFF
										'7.bOpsDataUpdate:Allow Operation Data Update Option
										'8.bOpaDataOverride:Allow Operation Data Update Override Option

'Return Value		   : 	True Or False

'Pre-requisite			:	New Business Object Constant Dialog Should be Appear on Screen

'Examples				: 						--	For Creating Business Object Constants  --
'											Fn_BMIDE_CreateBusinessObjectConstantExt("Demo","Demo","*","String","A","","ON","ON")	
'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done								Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N								   				18/02/2011			           1.0																								Sunny R
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_BMIDE_CreateBusinessObjectConstantExt(strName,strDesc,strScope,strDataType,strDefaultValue,strValues,bOpsDataUpdate,bOpaDataOverride)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_CreateBusinessObjectConstantExt"
   'Variable Declaration
   Dim strPrefix,arrDefaultValues,iCounter,arrValues,arrValueSet,arrScope,arrListValue,arrPropScope
   Dim ObjGlobalConstDialog,ObjScopeDialog,ObjListValues
	Fn_BMIDE_CreateBusinessObjectConstantExt=False
	'Checking Existance Of "NewBusinessObjectConstant" Window
	If Fn_UI_ObjectExist("Fn_BMIDE_CreateBusinessObjectConstantExt",JavaWindow("Business Modeler").JavaWindow("NewBusinessObjectConstant"))=True Then
		'Creating Object Of "NewBusinessObjectConstant" Window
		Set ObjGlobalConstDialog=Fn_UI_ObjectCreate("Fn_BMIDE_CreateBusinessObjectConstantExt",JavaWindow("Business Modeler").JavaWindow("NewBusinessObjectConstant"))
	Else
	    Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail: NewBusinessObjectConstant Dialog Is Not Exist")
		Exit Function
	End If
	'Retriving Prefix From "Name" Edit Box
	strPrefix=Fn_Edit_Box_GetValue("Fn_BMIDE_CreateBusinessObjectConstantExt",ObjGlobalConstDialog,"Name")
	'Attaching Prefix To Name
	strName=strPrefix+strName
	'Setting Name To Global Constant
	Call Fn_Edit_Box("Fn_BMIDE_CreateBusinessObjectConstantExt",ObjGlobalConstDialog,"Name",strName)
	If strDesc<>"" Then
		'Setting Description To New Global Constant
		Call Fn_Edit_Box("Fn_BMIDE_CreateBusinessObjectConstantExt",ObjGlobalConstDialog,"Description",strDesc)
	End If
'	Setting Scopes To New Global Constant
	If strScope<>"" Then
			arrScope= Split(strScope,"~")
			For iCounter = 0 to UBound(arrScope)
					arrPropScope = Split(arrScope(iCounter),":")
					' Click  on add button
					Call Fn_Button_Click("Fn_BMIDE_CreateBusinessObjectConstantExt",ObjGlobalConstDialog,"Add")
					If Fn_UI_ObjectExist("Fn_BMIDE_CreateBusinessObjectConstantExt",JavaWindow("Business Modeler").JavaWindow("NewBusinessObjectDefineScope"))=True Then
							'Creating Object Of "NewBusinessObjectDefineScope" Window
							Set ObjScopeDialog=Fn_UI_ObjectCreate("Fn_BMIDE_CreateBusinessObjectConstantExt",JavaWindow("Business Modeler").JavaWindow("NewBusinessObjectDefineScope"))
					Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail: NewBusinessObjectDefineScope Dialog Is Not Exist")
							Exit Function
					End If
					' Set  Business Object Scope value
					Call Fn_Edit_Box("Fn_BMIDE_CreateBusinessObjectConstantExt",ObjScopeDialog,"BusinessObjectScope",arrPropScope(0))
					If UBound(arrPropScope) = 1  Then
							' Set  Property Scope value
							Call Fn_Edit_Box("Fn_BMIDE_CreateBusinessObjectConstantExt",ObjScopeDialog,"PropertyScope",arrPropScope(1))
					End If
					' Click on finish Button
					Call Fn_Button_Click("Fn_BMIDE_CreateBusinessObjectConstantExt",ObjScopeDialog,"Finish")	
					Set ObjScopeDialog = Nothing
			Next
	End If
	If strDataType<>"" Then
		'Selecting Data Type
		Call Fn_List_Select("Fn_BMIDE_CreateBusinessObjectConstantExt",ObjGlobalConstDialog,"DataType",strDataType)
	End If
	If strDefaultValue<>"" Then
			If strDataType="String" Then
					Call Fn_Edit_Box("Fn_BMIDE_CreateBusinessObjectConstantExt",ObjGlobalConstDialog,"DefaultValue",strDefaultValue)
			End If
			If strDataType="Boolean" Then
					Call Fn_List_Select("Fn_BMIDE_CreateBusinessObjectConstantExt",ObjGlobalConstDialog,"DefaultValue",LCase(strDefaultValue))
			End If
			If strDataType="List" Then
					arrValues = Split(strValues,"~")
					For iCounter =0 to UBound(arrValues)
							arrListValue = Split(arrValues(iCounter),":")
							' Click  on add button
							Call Fn_Button_Click("Fn_BMIDE_CreateBusinessObjectConstantExt",ObjGlobalConstDialog,"Add2")
							If Fn_UI_ObjectExist("Fn_BMIDE_CreateBusinessObjectConstantExt",JavaWindow("Business Modeler").JavaWindow("NewBusinessObjectAddValue"))=True Then
									'Creating Object Of "NewBusinessObjectAddValue" Window
									Set ObjListValues=Fn_UI_ObjectCreate("Fn_BMIDE_CreateBusinessObjectConstantExt",JavaWindow("Business Modeler").JavaWindow("NewBusinessObjectAddValue"))
							Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail: NewBusinessObjectAddValue  Dialog Is Not Exist")
									Exit Function
							End If
							Call Fn_Edit_Box("Fn_BMIDE_CreateBusinessObjectConstantExt",ObjListValues,"Value",arrListValue(0))
							If  UCase(arrListValue(1)) = "ON" Then
								Call Fn_CheckBox_Select("Fn_BMIDE_CreateBusinessObjectConstantExt",ObjListValues,"Secured")
							End If
							' Click on finish Button
							Call Fn_Button_Click("Fn_BMIDE_CreateBusinessObjectConstantExt",ObjListValues,"Finish")	
							Set ObjListValues = Nothing	
					Next
					'Selecting Default value
					Call Fn_Edit_Box("Fn_BMIDE_CreateBusinessObjectConstantExt",ObjGlobalConstDialog,"DefaultValue",strDefaultValue)
    		End If
	End If

	If bOpsDataUpdate<>"" Then
		Call Fn_CheckBox_Set("Fn_BMIDE_CreateBusinessObjectConstantExt", JavaWindow("Business Modeler").JavaWindow("NewBusinessObjectConstant"), "AllowOpsDataUpdate", bOpsDataUpdate)
	End If
	If bOpaDataOverride<>"" Then
		Call Fn_CheckBox_Set("Fn_BMIDE_CreateBusinessObjectConstantExt", JavaWindow("Business Modeler").JavaWindow("NewBusinessObjectConstant"), "AllowOpsDataOverride", bOpaDataOverride)
	End If

	'Clicking "Finish" Button to Create New Global Constants
	Call Fn_Button_Click("Fn_BMIDE_CreateBusinessObjectConstantExt",ObjGlobalConstDialog,"Finish")
	Fn_BMIDE_CreateBusinessObjectConstantExt=True
	'Releasing Object Of "Global Constants Dialog"
	Set ObjGlobalConstDialog=Nothing
End Function

'-------------------------------------------------------------------Function Used to Set Perspectives-----------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_SetPerspective

''Description			:	Function Used to Set Perspectives

'Return Value		   : 	True or False

'Parameters     		:	1. StrModule : Perspective Name

'Pre-requisite			:	BMIDE Client Should be Open.

'Examples				: Fn_BMIDE_SetPerspective("Business Modeler IDE")

'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				09/03/2011			           1.0																						Sunny R
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_BMIDE_SetPerspective(StrModule)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_SetPerspective"
	Dim strMenu,iRows,iCounter,bFlag
	Dim ObjPerspectiveDialog,ObjPerspectiveTable
	bFlag=False
	Fn_BMIDE_SetPerspective=False
	Set ObjPerspectiveDialog=JavaWindow("BMIDEDefaultWindow").JavaWindow("OpenPerspective")
	Call Fn_BMIDE_ToolbarButtonClick("","Open Perspective")
	If Not ObjPerspectiveDialog.Exist(5) Then
		strMenu=Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\BMIDEConfig\BMIDE_Menu.xml", "OpenPerspective")
		Call Fn_BMIDE_MenuOperation("SelectExt", strMenu)
	End If
	ObjPerspectiveDialog.Maximize
	Set ObjPerspectiveTable=ObjPerspectiveDialog.JavaTable("PerspectiveTable")
	iRows=Fn_UI_Object_GetROProperty("", ObjPerspectiveTable,"rows")
	'iRows=ObjPerspectiveTable.GetROProperty("rows")
	For iCounter=0 To iRows-1
		If Trim(StrModule) = Trim(ObjPerspectiveTable.GetCellData(iCounter,0)) Then
			ObjPerspectiveTable.SelectCell iCounter,0
			bFlag=True
			Exit For
		End If
	Next
	If bFlag=False Then
		ObjPerspectiveDialog.JavaButton("Cancel").Click micLeftBtn
		Set ObjPerspectiveDialog=Nothing
		Set ObjPerspectiveTable=Nothing
		Exit Function
	End If
	ObjPerspectiveDialog.JavaButton("OK").Click micLeftBtn
	Fn_BMIDE_SetPerspective=True
	Set ObjPerspectiveDialog=Nothing
	Set ObjPerspectiveTable=Nothing
End Function


'------------------------------------------------------------Function Used to Perform Operations On Naming Rule Pattern Table---------------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_PatternTableOperations

'Description			 :	Function Used to Perform Operations On Naming Rule Pattern Table

'Parameters			   :   '1.strAction: Action Name
'										 2.strPattern : Pattern Name
'										 3.strNewPattern : New Pattern Name
'										4.strDesc: Pattern Description
'										 4.bCounters: Generate counter Option
'										 5.strInitialValue: Initial value of Counter
'										6.strMaxValue: Maximum value of Counter

'Return Value		   : 	True Or False


'Examples				:	'Fn_BMIDE_PatternTableOperations("Edit","""Test-""NNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNN","""Test""NNN","","On","Test001","Test999")
'										Fn_BMIDE_PatternTableOperations("Verify","""Test""nnn","","","","Test001","Test999")

'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N													29-Mar-2011								1.0																				Sunny
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Pranav Ingle												20-Feb-2013								1.1							Modified Case "Edit"				Sunny
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------	
'													Avinash J													 16-July-2013								1.2							Added Case "Verify"				Pranav Ingle
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------	
Public Function Fn_BMIDE_PatternTableOperations(strAction,strPattern,strNewPattern,strDesc,bCounter,strInitialVal,strMaxVal)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_PatternTableOperations"
   Dim ObjPattTable,ObjPatternDialog
   Dim iRowCount,iCounter,strCurrPattern,bFlag, iCounter1

   Fn_BMIDE_PatternTableOperations=False
	Set ObjPattTable=JavaWindow("Business Modeler").JavaTable("NamingRulePatternTable")
	Set ObjPatternDialog=JavaWindow("Business Modeler").JavaWindow("NamingRulePattern")
	bFlag=False
	Select Case strAction
		Case "Edit", "Edit_DoubleClick"
			iRowCount=ObjPattTable.GetROProperty("rows")
			For iCounter=0 To iRowCount-1
				strCurrPattern=ObjPattTable.GetCellData(iCounter,"Pattern")
				If Trim(strCurrPattern)=Trim(strPattern) Then
'					ObjPattTable.ActivateRow iCounter
					ObjPattTable.SelectCell iCounter, "Pattern"
                    bFlag=True
					Exit For
				End If
			Next
			If bFlag=False Then
				Exit Function
			End If
			If  strAction = "Edit_DoubleClick" Then
				ObjPattTable.ActivateRow iCounter
			Else
				Call Fn_Button_Click("Fn_BMIDE_PatternTableOperations", JavaWindow("Business Modeler"), "EditNamingRulePattern")
			End If

			If strNewPattern<>"" Then
				Call Fn_Edit_Box("Fn_BMIDE_PatternTableOperations",ObjPatternDialog,"Pattern",strNewPattern)
			End If
			If bCounter<>"" Then
				Call Fn_CheckBox_Set("Fn_BMIDE_PatternTableOperations", ObjPatternDialog, "GenerateCounters", bCounter)
			End If

			If strDesc<>"" Then
				Call Fn_Edit_Box("Fn_BMIDE_PatternTableOperations",ObjPatternDialog,"Description",strDesc)
			End If

			If strInitialVal<>"" Then
				Call Fn_Edit_Box("Fn_BMIDE_PatternTableOperations",ObjPatternDialog,"InitialValue",strInitialVal)
			End If
			If strMaxVal<>"" Then
				Call Fn_Edit_Box("Fn_BMIDE_PatternTableOperations",ObjPatternDialog,"MaximumValue",strMaxVal)
			End If
			Call Fn_Button_Click("Fn_BMIDE_PatternTableOperations", ObjPatternDialog, "Finish")
			Fn_BMIDE_PatternTableOperations=True

		Case "Verify"
			bFlag=False
			iRowCount=ObjPattTable.GetROProperty("rows")
			For iCounter=0 To iRowCount-1
				strCurrPattern=ObjPattTable.GetCellData(iCounter,"Pattern")
				If Trim(strCurrPattern)=Trim(strPattern) Then
				bFlag=True
					If strInitialVal<> "" Then
							bFlag=False
							For iCounter1=0 To iRowCount-1
								strCurrPattern=ObjPattTable.GetCellData(iCounter1,"InitialValue")
								If Trim(strCurrPattern)=Trim(strInitialVal) Then
									 bFlag=True
									 Exit For
								End If
						   Next
					End If
					If strMaxVal<> "" Then
							bFlag=False
							For iCounter1=0 To iRowCount-1
								strCurrPattern=ObjPattTable.GetCellData(iCounter1,"MaximumValue")
								If Trim(strCurrPattern)=Trim(strMaxVal) Then
										 bFlag=True
										 Exit For
								End If
						   Next
					End If
					If strDesc<> "" Then
							bFlag=False
							For iCounter1=0 To iRowCount-1
								strCurrPattern=ObjPattTable.GetCellData(iCounter1,"Description")
								If Trim(strCurrPattern)=Trim(strDesc) Then
									 bFlag=True
									 Exit For
								 End If
							 Next
					 End If
					 Exit For
				End If
			Next
			If  bFlag=True Then
				Fn_BMIDE_PatternTableOperations=True
			End If

	Case "Remove"
			iRowCount=ObjPattTable.GetROProperty("rows")
			For iCounter=0 To iRowCount-1
				strCurrPattern=ObjPattTable.GetCellData(iCounter,"Pattern")
				If Trim(strCurrPattern)=Trim(strPattern) Then
'					ObjPattTable.ActivateRow iCounter
					ObjPattTable.SelectCell iCounter, "Pattern"
                    bFlag=True
					Exit For
				End If
			Next
			If bFlag=False Then
				Exit Function
			End If
        	Call Fn_Button_Click("Fn_BMIDE_PatternTableOperations", JavaWindow("Business Modeler"), "RemoveNamingRuleAttches")
			Fn_BMIDE_PatternTableOperations=True

	End Select
	Set ObjPatternDialog=Nothing
	Set ObjPattTable=Nothing
End Function
'------------------------------------------------------------Function Used to Perform Operations On Naming Rule Change ID Table--------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_NamingRuleChangeIDTableOperations

'Description			 :	Function Used to Perform Operations On Naming Rule Change ID Table

'Parameters			   :   '1.strAction: Action Name
'										 2.strRange : Range
'										 3.strValue : Value
'										4.strFormat: Format

'Return Value		   : 	True Or False


'Examples				:	'Fn_BMIDE_NamingRuleChangeIDTableOperations("Remove","1-3","","")
'										Fn_BMIDE_NamingRuleChangeIDTableOperations("Add","4-7","ASWW","Static")

'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N													30-Mar-2011								1.0																						Sunny
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------	
Public Function Fn_BMIDE_NamingRuleChangeIDTableOperations(strAction,strRange,strValue,strFormat)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_NamingRuleChangeIDTableOperations"
   Dim ObjBMIDEWindow,ObjChangeIDWindow,ObjChangeIDTable
   Dim iRowCount,iCounter,crrRange

   Fn_BMIDE_NamingRuleChangeIDTableOperations=False
	Set ObjBMIDEWindow=JavaWindow("Business Modeler")
	Set ObjChangeIDTable=JavaWindow("Business Modeler").JavaTable("NamingRuleChangeIDTable")
	Set ObjChangeIDWindow=JavaWindow("Business Modeler").JavaWindow("NamingRuleChangeID")
	Call Fn_BMIDE_InnerTabOperations("Activate","Naming Rules")
	Select Case strAction
			Case "Add"
				Call Fn_Button_Click("Fn_BMIDE_NamingRuleChangeIDTableOperations", ObjBMIDEWindow, "AddChangeID")
				If strRange<>"" Then
					Call Fn_Edit_Box("Fn_BMIDE_NamingRuleChangeIDTableOperations",ObjChangeIDWindow,"Range",strRange)
				End If
				If strValue<>"" Then
					Call Fn_Edit_Box("Fn_BMIDE_NamingRuleChangeIDTableOperations",ObjChangeIDWindow,"Value",strValue)
				End If
				If strFormat<>"" Then
					Call Fn_List_Select("Fn_BMIDE_NamingRuleChangeIDTableOperations", ObjChangeIDWindow, "Format",strFormat)
				End If
				Call Fn_Button_Click("Fn_BMIDE_NamingRuleChangeIDTableOperations", ObjChangeIDWindow, "Finish")
				Fn_BMIDE_NamingRuleChangeIDTableOperations=True
			Case "Remove"
				iRowCount=ObjChangeIDTable.GetROProperty("rows")
				For iCounter=0 To iRowCount-1
					crrRange=ObjChangeIDTable.GetCellData(iCounter,"Range")
					If Trim(crrRange)=Trim(strRange) Then
						ObjChangeIDTable.ActivateRow iCounter
						Call Fn_Button_Click("Fn_BMIDE_NamingRuleChangeIDTableOperations", ObjBMIDEWindow, "RemoveChangeID")
						Fn_BMIDE_NamingRuleChangeIDTableOperations=True
						Exit For
					End If
				Next
	End Select
	Set ObjBMIDEWindow=Nothing
	Set ObjChangeIDWindow=Nothing
	Set ObjChangeIDTable=Nothing
End Function
'------------------------------------------------------------Function Used to Perform Operations On Naming Rule Rev ID Table--------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_NamingRuleRevIDTableOperations

'Description			 :	Function Used to Perform Operations On Naming Rule Rev ID Table

'Parameters			   :   '1.strAction: Action Name
'										 2.strRange : Range
'										 3.strValue : Value
'										4.strFormat: Format

'Return Value		   : 	True Or False


'Examples				:	'Fn_BMIDE_NamingRuleRevIDTableOperations("Remove","1","","")
'										Fn_BMIDE_NamingRuleRevIDTableOperations("Add","1","A","Running")

'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N													30-Mar-2011								1.0																						Sunny
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------	
Public Function Fn_BMIDE_NamingRuleRevIDTableOperations(strAction,strRange,strValue,strFormat)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_NamingRuleRevIDTableOperations"
   Dim ObjBMIDEWindow,ObjChangeIDWindow,ObjRevIDTable
   Dim iRowCount,iCounter,crrRange

   Fn_BMIDE_NamingRuleRevIDTableOperations=False
	Set ObjBMIDEWindow=JavaWindow("Business Modeler")
	Set ObjChangeIDWindow=JavaWindow("Business Modeler").JavaWindow("NamingRuleChangeID")
	Set ObjRevIDTable=JavaWindow("Business Modeler").JavaTable("NamingRuleRevIDTable")
	Call Fn_BMIDE_InnerTabOperations("Activate","Naming Rules")
	Select Case strAction
			Case "Add"
				Call Fn_Button_Click("Fn_BMIDE_NamingRuleRevIDTableOperations", ObjBMIDEWindow, "AddRevID")
				If strRange<>"" Then
					Call Fn_Edit_Box("Fn_BMIDE_NamingRuleRevIDTableOperations",ObjChangeIDWindow,"Range",strRange)
				End If
				If strValue<>"" Then
					Call Fn_Edit_Box("Fn_BMIDE_NamingRuleRevIDTableOperations",ObjChangeIDWindow,"Value",strValue)
				End If
				If strFormat<>"" Then
					Call Fn_List_Select("Fn_BMIDE_NamingRuleRevIDTableOperations", ObjChangeIDWindow, "Format",strFormat)
				End If
				Call Fn_Button_Click("Fn_BMIDE_NamingRuleRevIDTableOperations", ObjChangeIDWindow, "Finish")
				Fn_BMIDE_NamingRuleRevIDTableOperations=True
			Case "Remove"
				iRowCount=ObjRevIDTable.GetROProperty("rows")
				For iCounter=0 To iRowCount-1
					crrRange=ObjRevIDTable.GetCellData(iCounter,"Range")
					If Trim(crrRange)=Trim(strRange) Then
						ObjRevIDTable.ActivateRow iCounter
						Call Fn_Button_Click("Fn_BMIDE_NamingRuleChangeIDTableOperations", ObjBMIDEWindow, "RemoveRevID")
						Fn_BMIDE_NamingRuleRevIDTableOperations=True
						Exit For
					End If
				Next
	End Select
	Set ObjBMIDEWindow=Nothing
	Set ObjChangeIDWindow=Nothing
	Set ObjRevIDTable=Nothing
End Function 

'------------------------------------------------------------Function Used to Select / Edit Naming Rule pattern---------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_NamingRulePattern

'Description			 :	Function Used to Select / Edit Naming Rule pattern

'Parameters			   :   '1.strAction: Select action to perform
'										 2.strPattern : Pattern
'										 3.strDescription : Pattern Description
'										 4.bCounters: Generate counter Option
'										 5.strInitialValue: Initial value of Counter
'										6.strMaxValue: Maximum value of Counter

'Return Value		   : 	True Or False

'Examples				:   Fn_BMIDE_NamingRulePattern("SelectPattern","""A""nn","","","","")
'										Fn_BMIDE_NamingRulePattern("EditPattern","{LOV:Activity Category}[RULE:T2test]","FirstDemoRulePattern","OFF","","")
'										Fn_BMIDE_NamingRulePattern("EditPattern","""B""nn","FirstDemoRulePattern","ON","B00","B99")
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Pranav Ingle												12-Apr-2011								1.0																						Sunny
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------	
Public Function Fn_BMIDE_NamingRulePattern(strAction,strPattern,strDescription,bCounters,strInitialValue,strMaxValue)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_NamingRulePattern"
   'Variable Declaration
   Dim strPrefix,arrPattern,iCounter,intRowCount,strPropName
   Dim ObjRulesDialog,ObjPatternDialog
   Fn_BMIDE_NamingRulePattern=False

	Select Case strAction
			Case "SelectPattern"
					intRowCount=JavaWindow("Business Modeler").JavaTable("NamingRulePatternTable").GetROProperty("rows")
					For iCounter=0 To intRowCount-1
							strPropName=JavaWindow("Business Modeler").JavaTable("NamingRulePatternTable").GetCellData(iCounter,"Pattern")
							If Trim(strPropName)=Trim(strPattern) Then
									JavaWindow("Business Modeler").JavaTable("NamingRulePatternTable").ActivateRow iCounter
									Exit For
							End If
					Next 
	
			Case "EditPattern"
					'Clicking On Next Button
					If Fn_UI_ObjectExist("Fn_BMIDE_NamingRulePattern", JavaWindow("Business Modeler").JavaButton("EditNamingRulePattern"))=True Then
							Call Fn_Button_Click("Fn_BMIDE_NamingRulePattern",  JavaWindow("Business Modeler"), "EditNamingRulePattern")
					 End If
					 wait(1)
				   'Creating Object Of "NamingRulePattern" Window
				   Set ObjPatternDialog=JavaWindow("Business Modeler").JavaWindow("NamingRulePattern")
						'Setting Naming Rule Pattern
						If strPattern<>"" Then
								Call Fn_Edit_Box("Fn_BMIDE_NamingRulePattern",ObjPatternDialog,"Pattern",strPattern)
						End If
						'Setting Naming Rule Pattern Description
						If strDescription<>"" Then
								Call Fn_Edit_Box("Fn_BMIDE_NamingRulePattern",ObjPatternDialog,"Description",strDescription)
						End If
						If bCounters<>"" Then
							If bCounters="ON" Then
									'Setting Stautus Of Generate Counter Option
									Call Fn_CheckBox_Set("Fn_BMIDE_NamingRulePattern",ObjPatternDialog,"GenerateCounters",bCounters)
									'Setting Initial Value
									If strInitialValue<>"" Then
											Call Fn_Edit_Box("Fn_BMIDE_NamingRulePattern",ObjPatternDialog,"InitialValue",strInitialValue)
									End If
									'Setting Maximum Value
									If strMaxValue<>"" Then
											Call Fn_Edit_Box("Fn_BMIDE_NamingRulePattern",ObjPatternDialog,"MaximumValue",strMaxValue)
									End If
							End If
						End If
						'Clicking On Finish Button
						Call Fn_Button_Click("Fn_BMIDE_NamingRulePattern", ObjPatternDialog, "Finish")
						'Releasing Object Of "NamingRulePattern" Window
						Set ObjPatternDialog=Nothing
		End Select
	Fn_BMIDE_NamingRulePattern=True
End Function

'------------------------------------------------------------Function Used to Restart Services---------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_RestartWarmServers

'Description			 :	Function Used to Restart Services

'Parameters			   :   '1.strAdminURL: URL of Admin for Server
'										 2.strUserName : Server Username
'										 3.strPassword : Server Password

'Return Value		   : 	True Or False

'Examples				:   Fn_BMIDE_RestartWarmServers("","","")
'										Fn_BMIDE_RestartWarmServers("http://pnv6s110/tc911109/admin","autoadmin","Password123")
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sunny Ruparel												29-Nov-2011								1.0																				Sandeep N
'													Sandeep N												02-Dec-2011								1.1																				Sunny R
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------	
Function Fn_BMIDE_RestartWarmServers(strAdminURL,strUserName,strPassword)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_RestartWarmServers"
	
	   Dim strServerPath,ObjBrowser,handlewin,strUserCredentials
	
  	    Fn_BMIDE_RestartWarmServers=False
		If strAdminURL = "" Then
			strServerPath = Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\BMIDEConfig\BMIDE_Config.xml", "AdminURL") 
		Else
			strServerPath = strAdminURL
		End If
		If strUserName = "" OR strPassword = "" Then
			strUserCredentials = Split(Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\BMIDEConfig\BMIDE_Config.xml", "ServerUser"),":")
		Else
			strUserCredentials = strUserName +":"+ strPassword
		End If
		Set ObjBrowser = CreateObject("InternetExplorer.Application")
		ObjBrowser.Visible = True
		'Launch Ie and enter the Admin URL
		ObjBrowser.Navigate strServerPath
		handlewin = ObjBrowser.HWND
		'Check for existence of the Login window
		If Dialog("CredentialWindow").Exist(60) Then

			Dialog("CredentialWindow").WinEdit("UserName").Set strUserCredentials(0)
			Dialog("CredentialWindow").WinEdit("Password").Set strUserCredentials(1)
			'Click OK
			If Dialog("CredentialWindow").WinButton("OK").Exist(10) Then
			   Dialog("CredentialWindow").WinButton("OK").Click
			End If

		ElseIf Browser("AdminBrowser").Dialog("CredentialWindow").Exist(60) Then
			'Enter Username and Password 
			Browser("AdminBrowser").Dialog("CredentialWindow").WinEdit("UserName").Set strUserCredentials(0)
			Browser("AdminBrowser").Dialog("CredentialWindow").WinEdit("Password").Set strUserCredentials(1)
			'Click OK
			If Browser("AdminBrowser").Dialog("CredentialWindow").WinButton("OK").Exist(10) Then
				Browser("AdminBrowser").Dialog("CredentialWindow").WinButton("OK").Click
			End If
		End If
		wait 2
		Window("hwnd:="+CStr(handlewin)).Maximize

		If Browser("AdminBrowser").WebButton("RestartWarmServers").Exist(30) Then
			Browser("AdminBrowser").WebButton("RestartWarmServers").Click
			wait(10)
			Fn_BMIDE_RestartWarmServers=True
	   End If
	'Close the window
	Browser("AdminBrowser").Close
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_BMIDE_RevisionNamingRuleAttachmentsOparation
'@@
'@@    Description				 :	Use to perform operations on Revision Naming Rule Attachments Table
'@@
'@@    Parameters			   :	1.StrAction :Action Name
'@@										 2.StrBOObject : Business Object name
'@@										 3.StrProperty : Property Name
'@@										 4.StrCondition : Condition
'@@										 5.StrOverride :
'@@										 6.StrTemplate : Template Name
'@@
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	RevisionNamingRuleAttachments table Should be Displayed
'@@
'@@    Examples					:	Call Fn_BMIDE_RevisionNamingRuleAttachmentsOparation("Verify","T3Liveupdate15Revision","item_id_revision","","","")
'@@    Examples					:	Call Fn_BMIDE_RevisionNamingRuleAttachmentsOparation("Detach","T3Liveupdate15Revision","item_id_revision","","","")
'@@
'@@	   History					 	:	
'@@													Developer Name												Date						Rev. No.						Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@												Sandeep Navghane									         30-Nov-2011						      1.0																Sunny Ruparel
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_BMIDE_RevisionNamingRuleAttachmentsOparation(StrAction,StrBOObject,StrProperty,StrCondition,StrOverride,StrTemplate)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_RevisionNamingRuleAttachmentsOparation"
	Dim ObjRevTable
	Dim iRowCount,iCounter,bFlag,crrBOObject

	Set ObjRevTable=JavaWindow("Business Modeler").JavaTable("RevisionNamingRuleAttachments")
	Fn_BMIDE_RevisionNamingRuleAttachmentsOparation=False
	Select Case StrAction
		Case "Verify"
			iRowCount=ObjRevTable.GetROProperty("rows")
			For iCounter=0 To iRowCount-1
				bFlag=False
				crrBOObject=ObjRevTable.GetCellData(iCounter,"Business Object.Property")
				If Trim(crrBOObject)=Trim(StrBOObject)+"."+Trim(StrProperty) Then
					bFlag=True
					Exit For
				End If
			Next
			If bFlag=True Then
				Fn_BMIDE_RevisionNamingRuleAttachmentsOparation=True
			End If

		Case "Detach"
			iRowCount=ObjRevTable.GetROProperty("rows")
			For iCounter=0 To iRowCount-1
				bFlag=False
				crrBOObject=ObjRevTable.GetCellData(iCounter,"Business Object.Property")
				If Trim(crrBOObject)=Trim(StrBOObject)+"."+Trim(StrProperty) Then
					ObjRevTable.ActivateRow iCounter
					bFlag=True
					Exit For
				End If
			Next
			If bFlag=True Then
				Call Fn_Button_Click("Fn_BMIDE_RevisionNamingRuleAttachmentsOparation",JavaWindow("Business Modeler"),"DetachRevisionNamingRule")
				Fn_BMIDE_RevisionNamingRuleAttachmentsOparation=True
			End If
   End Select
   Set ObjRevTable=Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_BMIDE_GetTemplatePrefix

'Description			 :	Function Used to Get Dependent Template Prefix from DependentTemplatePrefix.xls

'Parameters			   :   '1.strExcelPath: DependentTemplatePrefix.xls Path { Not mandetory }
'										 2.strTestCaseName : QTP Script Name { Not mandetory }
'										 3.intSheetNumber : Sheet Number { Not mandetory }

'Return Value		   : 	Prefix Or False

'Examples				:  Fn_BMIDE_GetTemplatePrefix("", "", "")
'									   Fn_BMIDE_GetTemplatePrefix("D:\mainline\TestData\BMIDEConfig\DependentTemplatePrefix.xls", "BMIDE_PropertyConstantsUpdate005", "")
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												22-Dec-2011								1.0																						Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_BMIDE_GetTemplatePrefix(strExcelPath,strTestCaseName,intSheetNumber)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_GetTemplatePrefix"

	Dim objFSO
	Dim objExcel
	Dim objWorkbook
	Dim objWorksheet
	Dim iCount,bFlag,crrTestCaseName,Prefix
	'Creating File System Object
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	If strExcelPath="" Then
		strExcelPath=Environment.Value("sPath") + "\TestData\BMIDEConfig\DependentTemplatePrefix.xls"
	End If
	If objFSO.FileExists(strExcelPath) Then
	
		Set objExcel = CreateObject("Excel.Application")
		objExcel.Visible=False
		Set objWorkbook = objExcel.Workbooks.Open(strExcelPath)
		
		If intSheetNumber = "" Then
			 intSheetNumber = 1
		End If 	
		
		Set objWorksheet = objWorkbook.Worksheets(intSheetNumber)
		objWorksheet.Activate
		For iCount=2 To objWorksheet.UsedRange.Rows.Count
			bFlag=False
			crrTestCaseName = objExcel.Cells(iCount, 1).Value
			If strTestCaseName="" Then
				strTestCaseName=Environment.Value("TestName")
			End If
			If crrTestCaseName=strTestCaseName Then
				Prefix = objExcel.Cells(iCount, 2).Value
				bFlag=True
			End If
			If bFlag=True Then
				Exit For
			End If
		Next
		If bFlag=True Then
			Fn_BMIDE_GetTemplatePrefix=Prefix
		Else
			Fn_BMIDE_GetTemplatePrefix=False
		End If
		
		objExcel.Quit
		Set objWorksheet = Nothing
		Set objWorkbook = Nothing
		Set objExcel = Nothing			

	Else
		Fn_BMIDE_GetTemplatePrefix = False	
	End If
	
	Set objFSO = Nothing
	
End Function


'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'Function Name		:	Fn_BMIDE_CreateDependentTemplateProject

'Description			 :	Function Used to Create New Dependent template Base Project 

'Parameters			   :	1.strProjectName: Name of Project To Be Created
										'2.bDefaultOpt:Template Location Option
										'3.strDescription: Project Description
										'4.strTempDirectory:Template Directory Path
										'5.strDepdTemplate:Dependant template Names
										'6.strLanguage:Language to Select

'Return Value		   : 	Project Prefix Or False

'Pre-requisite			:	Should be Log In BMIDE

'Examples				: 	'1. Call Fn_BMIDE_CreateDependentTemplateProject("Test123","","DemoDep Temp Project","","","")
										'Imp Note : For this Function All parameters come from environment(First Take Values From Env File And Pass to Function)
										'Pass Full Language Names by Collan Separeted (:)
										'strLanguage="cs_CZ - Czech (Czech Republic):zh_CN - Chinese (China)"
										'Call Fn_BMIDE_CreateDependentTemplateProject("Temp","True","","D:\Siemens\Teamcenter8\bmide\templates","",strLanguage)
										'strDepdTemplate :-Dependant Templates ( : ) sepearated
										'Call Fn_BMIDE_CreateDependentTemplateProject("Temp","True","TestProject","C:\Siemens\Teamcenter8\bmide\templates","Foundation:Change Management:Aerospace and Defense Change Management","")
'History					 :		
'													Developer Name												Date						Rev. No.						Changes Done														Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'													Sandeep N										   				26/12/2011			           1.0																																					Sunny R
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Public Function Fn_BMIDE_CreateDependentTemplateProject(strProjectName,bDefaultOpt,strDescription,strTempDirectory,strDepdTemplate,strLanguage)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_CreateDependentTemplateProject"
   'Variable declaration
   Dim strMenu,iCounter,arrLanguage,iItemCount,iCount,LangName,bFlag,arrDepTemp,iRowCount,strTempName,StrAlpha,arrAlpha
   Dim strPrefix,crrError,crrPrefix
   Dim ObjNewProjectWindow

   StrAlpha="A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z"
	
   'Setting Function equals to False
   Fn_BMIDE_CreateDependentTemplateProject=False
   'Checking Existance of NewProject window
	If Fn_UI_ObjectExist("Fn_BMIDE_CreateDependentTemplateProject",JavaWindow("Business Modeler").JavaWindow("NewProject"))=False Then
		'Taking Menu Name from Environmet File
		strMenu=Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\BMIDEConfig\BMIDE_Menu.xml", "NewProject")
		'Calling File:New:Project... Menu to open Project Dialog
        Call Fn_BMIDE_MenuOperation("Select", strMenu)
	End If
	'Creating Object Of NewProject window
	Set ObjNewProjectWindow=Fn_UI_ObjectCreate("Fn_BMIDE_CreateDependentTemplateProject", JavaWindow("Business Modeler").JavaWindow("NewProject"))
	'Expanding Business Modeler IDE Node
    Call Fn_UI_JavaTree_Expand("Fn_BMIDE_CreateDependentTemplateProject", ObjNewProjectWindow, "WizardsTree","Business Modeler IDE")
	'Selecting Project
	Call Fn_JavaTree_Select("Fn_BMIDE_CreateDependentTemplateProject", ObjNewProjectWindow,"WizardsTree","Business Modeler IDE:New Business Modeler IDE Template Project")
	'Clicking On Next button to Go Next Wizard
	Call Fn_Button_Click("Fn_BMIDE_CreateDependentTemplateProject", ObjNewProjectWindow, "Next")

	bFlag=False
	'Retriving Invalid Prifix Error
	StrError=Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\BMIDEConfig\BMIDE_ErrorMsg.xml", "DependentTemplatePrifixError")
	crrPrefix=Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\BMIDEConfig\BMIDE_Config.xml", "ProjectPrefix")

	'Generating Prefix
	arrAlpha=Split(StrAlpha,",")
	For iCounter=0 To UBound(arrAlpha)
		For iCount=2 To 9
			strPrefix=Trim(arrAlpha(iCounter))+CStr(iCount)
			If Trim(strPrefix)<>Trim(crrPrefix) Then
				'Setting Prefix To project
				Call Fn_Edit_Box("Fn_BMIDE_CreateDependentTemplateProject",ObjNewProjectWindow,"Prefix",strPrefix)
				crrError=Fn_Edit_Box_GetValue("Fn_BMIDE_CreateDependentTemplateProject",ObjNewProjectWindow,"ErrorMsg")
				If InStr(1,LCase(Trim(crrError)),LCase(Trim(StrError))) Then
	
				Else
					bFlag=True
					Exit For
				End If
			End If
		Next
		If bFlag=True Then
			Exit For
		End If
	Next
    	If bFlag=False Then
		Set ObjNewProjectWindow=Nothing
		Exit Function
	End If

	If strProjectName<>"" Then
		'Seeting project Name
		Call Fn_Edit_Box("Fn_BMIDE_CreateDependentTemplateProject",ObjNewProjectWindow,"ProjectName", lcase(strPrefix+strProjectName))
	End If
	If bDefaultOpt=Cstr(True) Then
'		If strLocation<>"" Then
'			'Setting Project Location
'			Call Fn_Edit_Box("Fn_BMIDE_CreateDependentTemplateProject",ObjNewProjectWindow,"UseDefaultLocation",strLocation)
'		End If
	End If

	
	If strDescription<>"" Then
		Call Fn_Edit_Box("Fn_BMIDE_CreateDependentTemplateProject",ObjNewProjectWindow,"TemplateDesc",strProjectName)
	End If

	If strTempDirectory<>"" Then
		'Setting Template Directory to project
		Call Fn_Edit_Box("Fn_BMIDE_CreateDependentTemplateProject",ObjNewProjectWindow,"TemplatesDirectoryPath",strTempDirectory)
	End If
	'Clicking On Next button to Go Next Wizard
	Call Fn_Button_Click("Fn_BMIDE_CreateDependentTemplateProject", ObjNewProjectWindow, "Next")

	'SelectingDependanat Template Derectory
	If strDepdTemplate<>"" Then
		arrDepTemp=Split(strDepdTemplate,":")
		iRowCount=Fn_UI_Object_GetROProperty("Fn_BMIDE_CreateDependentTemplateProject",ObjNewProjectWindow.JavaTable("DependentTemplates"),"rows")
		For iCount=0 To Ubound(arrDepTemp)
			bFlag=False
			For iCounter=0 To iRowCount-1
				strTempName=ObjNewProjectWindow.JavaTable("DependentTemplates").GetCellData(iCounter,"Template display name")
				If Trim(strTempName)=arrDepTemp(iCount) Then
					If arrDepTemp(iCount)<>"Foundation" Then
						ObjNewProjectWindow.JavaTable("DependentTemplates").SelectRow iCounter
						ObjNewProjectWindow.JavaTable("DependentTemplates").PressKey " "
					End If
					bFlag=True
					Exit For
				End If
			Next
			If bFlag=False Then
				Call Fn_Button_Click("Fn_BMIDE_CreateDependentTemplateProject", ObjNewProjectWindow, "Cancel")
				Set ObjNewProjectWindow=Nothing
				Exit Function
			End If
		Next
	End If
	
	'Clicking On Next button to Go Next Wizard
	Call Fn_Button_Click("Fn_BMIDE_CreateDependentTemplateProject", ObjNewProjectWindow, "Next")
	If strLanguage<>"" Then
		arrLanguage=Split(strLanguage,":")
		For iCounter=0 To Ubound(arrLanguage)
			iItemCount=Fn_UI_Object_GetROProperty("Fn_BMIDE_CreateDependentTemplateProject",ObjNewProjectWindow.JavaTable("LocaleTable"), "rows")
			For iCount=0 To iItemCount-1
				bFlag=False
				LangName=ObjNewProjectWindow.JavaTable("LocaleTable").GetCellData(iCount,"Locale")
				If Trim(LangName)=Trim(arrLanguage(iCounter)) Then
					ObjNewProjectWindow.JavaTable("LocaleTable").SelectCell iCount,"Locale"
					wait(1)
					ObjNewProjectWindow.JavaTable("LocaleTable").PressKey " "
					bFlag=True
					Exit For
				End If
			Next
			If bFlag=False Then
				Exit For
			End If
		Next
		If bFlag=False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"False : Wrong Language Name pass by User")
			Set ObjNewProjectWindow=Nothing
			Exit Function
		End If
	End If
	'Clicking On Finish  button to Go to Create Project
	Call Fn_Button_Click("Fn_BMIDE_CreateDependentTemplateProject", ObjNewProjectWindow, "Finish")
	For iCounter=0 to 20
		If Not JavaWindow("Business Modeler").JavaWindow("NewProject").JavaWindow("ProgressInformation").Exist(7) Then
			Fn_BMIDE_CreateDependentTemplateProject=strPrefix
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Pass : Successfully Created a Project")
			Exit For
		End If
		wait(7)
	Next
	Set ObjNewProjectWindow=Nothing
End Function
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDE_DeleteProject

'Description			 :	Function Used to Delete Project From Navigator Tab

'Parameters			   :	1.strProjectName:Project Names
'								 2.bDeleteProjectContents : Delete project contents on disk option

'Return Value		   : 	True Or False

'Pre-requisite			:	Should be Log In BMIDE

'Examples				:   Call Fn_BMIDE_DeleteProject("dec291140_live_update1","")
'							    Call Fn_BMIDE_DeleteProject("DemoBatch25448:demobactch_opsdata1","ON")

'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done																				Reviewer
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				28/12/2011			           1.0																																																	Sunny R
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_BMIDE_DeleteProject(strProjectName,bDeleteProjectContents)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_DeleteProject"
   Dim arrProjectName,iCounter,iCount

   Fn_BMIDE_DeleteProject=False
   Call Fn_BMIDE_TabOperations("UpperLeft","Activate","Navigator")
   Call Fn_BMIDE_TreeIndexIdentification()   	
   arrProjectName=Split(strProjectName,":")
   For iCounter=0 To UBound(arrProjectName)
		Call Fn_BMIDE_BusinessObjectTreeOperations("Select",arrProjectName(iCounter),"")
		'Calling Edit:Delete Menu
		strMenuName=Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\BMIDEConfig\BMIDE_Menu.xml", "EditDelete")
	    Call Fn_BMIDE_MenuOperation("Select", strMenuName)

		If JavaWindow("Business Modeler").JavaWindow("DeleteResources").Exist(20) Then
			  If bDeleteProjectContents<>"" Then
				  Call Fn_CheckBox_Set("Fn_BMIDE_DeleteProject", JavaWindow("Business Modeler").JavaWindow("DeleteResources"), "DeleteProjectContentsFromDisk", bDeleteProjectContents)
			  End If
			  Call Fn_Button_Click("Fn_BMIDE_DeleteProject", JavaWindow("Business Modeler").JavaWindow("DeleteResources"), "OK")
			   wait 2
			  If JavaWindow("Business Modeler").JavaWindow("DeleteResources").Exist(5) Then
					wait 5
			  End if
		End If
	Next	
	Call Fn_BMIDE_TabOperations("UpperLeft","Activate","Business Objects")
	Call Fn_BMIDE_TreeIndexIdentification()
	Fn_BMIDE_DeleteProject=True
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_CreateNewIRDC

'Description			 :	Function Used to create new IRDC

'Parameters			   :   '1.dicIRDC: IRDC Dictionary

'Return Value		   : 	True Or False

'Pre-requisite			:	New IRDC window Should be opened

'Examples				:   dicIRDC("Name")="IRDC1"
'										dicIRDC("Description")="Basic IRDC"
'										dicIRDC("AppliesToBusinessObject")="T3LiveUpdate61Revision"
'										dicIRDC("Condition")="isTrue"
'										Call Fn_CreateNewIRDC(dicIRDC)

'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												30-Dec-2011								1.0																						Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Fn_CreateNewIRDC : This function is not completely developed . Its develope only for basic fuctionality.
'Developer can implemet remaining part in same function as Input parameter is Dictionary Object
Public Function Fn_CreateNewIRDC(dicIRDC)
	GBL_FAILED_FUNCTION_NAME="Fn_CreateNewIRDC"
 	'Variable declaration
	Dim StrPrefix,objNewIRDC
	'Initially Function returns false
	Fn_CreateNewIRDC=False
	'Checking Existance of [ NewIRDC ] window
	If Not JavaWindow("Business Modeler").JavaWindow("NewIRDC").Exist(6) Then
		'If [ NewIRDC ] window is not exist then exit from function
		Exit Function
	End If
	'Creating Object of [ NewIRDC ] window
	Set objNewIRDC=JavaWindow("Business Modeler").JavaWindow("NewIRDC")
	'Appending Prefix with Name
	StrPrefix=Fn_Edit_Box_GetValue("Fn_CreateNewIRDC",objNewIRDC,"Name")
	dicIRDC("Name")=StrPrefix+dicIRDC("Name")
	'Setting name to IRDC :- Its compulsory parameter
	Call Fn_Edit_Box("Fn_CreateNewIRDC",objNewIRDC,"Name",dicIRDC("Name"))
	'Setting Description to IRDC :- Its compulsory parameter
	If dicIRDC("Description")<>"" Then
		Call Fn_Edit_Box("Fn_CreateNewIRDC",objNewIRDC,"Description",dicIRDC("Description"))
	End If
	'Setting Applies To Business Object to IRDC :- Its compulsory parameter
	If dicIRDC("AppliesToBusinessObject")<>"" Then
		Call Fn_Edit_Box("Fn_CreateNewIRDC",objNewIRDC,"AppliesToBusinessObject",dicIRDC("AppliesToBusinessObject"))
	End If
	'Setting Condition to IRDC
	If dicIRDC("Condition")<>"" Then
		Call Fn_Edit_Box("Fn_CreateNewIRDC",objNewIRDC,"Condition",dicIRDC("Condition"))
	End If
	'Clicking on Finish button
	Call Fn_Button_Click("Fn_CreateNewIRDC", objNewIRDC, "Finish")
	'Function returns True
	Fn_CreateNewIRDC=True
	'Releasing object of [ NewIRDC ] window
	Set objNewIRDC=Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_BMIDE_CreateNewFunctionality

'Description			 :	Function Used to create new Functionality

'Parameters			   :   '1.dicFunctionality: Functionality Dictionary

'Return Value		   : 	True Or False

'Pre-requisite			:	New IRDC window Should be opened

'Examples				:   dicFunctionality("Name")="Functionality1"
'										dicFunctionality("DisplayName")="Functionality1"
'										dicFunctionality("Description")="Test Functionality1"
'										dicFunctionality("EnableForVerificationRules")="ON"
'										dicFunctionality("BusinessObjectScope")="Item~Document"
'										dicFunctionality("SupportedConditionSignature")="Item:it~Document:dt"
'										dicFunctionality("SubGroupLOV")="BillCodes"
'										Call Fn_BMIDE_CreateNewFunctionality(dicFunctionality)

'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												30-Dec-2011								1.0																						Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Public Function Fn_BMIDE_CreateNewFunctionality(dicFunctionality)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_CreateNewFunctionality"
 	'Variable Declaration
	Dim objNewFunctionality,StrPrefix,arrBOScope,iCounter,arrSCSingature,arrParameters
	Fn_BMIDE_CreateNewFunctionality=False
	'Checking Existance of [ NewFunctionality ] window
	If Not JavaWindow("Business Modeler").JavaWindow("NewFunctionality").Exist(6) Then
		'If [ NewFunctionality ] window is not exist then Exit Function
		Exit Function
	End If
	Set objNewFunctionality=JavaWindow("Business Modeler").JavaWindow("NewFunctionality")
	'Appending Prefix with Name
	StrPrefix=Fn_Edit_Box_GetValue("Fn_BMIDE_CreateNewFunctionality",objNewFunctionality,"Name")
	dicFunctionality("Name")=StrPrefix+dicFunctionality("Name")
	'Setting name to Functionality :- Its compulsory parameter
	Call Fn_Edit_Box("Fn_BMIDE_CreateNewFunctionality",objNewFunctionality,"Name",dicFunctionality("Name"))
	'Setting Display Name to Functionality
	If dicFunctionality("DisplayName")<>"" Then
		Call Fn_Edit_Box("Fn_BMIDE_CreateNewFunctionality",objNewFunctionality,"DisplayName",dicFunctionality("DisplayName"))
	End If
	'Setting Description to Functionality
	If dicFunctionality("Description")<>"" Then
		Call Fn_Edit_Box("Fn_BMIDE_CreateNewFunctionality",objNewFunctionality,"Description",dicFunctionality("Description"))
	End If
	'Clicking On Next Button to add additional Functionality information
	Call Fn_Button_Click("Fn_BMIDE_CreateNewFunctionality", objNewFunctionality, "Next")
	'Setting Enable For Verification Rules option
	If dicFunctionality("EnableForVerificationRules")<>"" Then
		Call Fn_CheckBox_Set("Fn_BMIDE_CreateNewFunctionality", objNewFunctionality, "EnableForVerificationRules", dicFunctionality("EnableForVerificationRules"))
	End If
	'Selecting Business Object Scope
	If dicFunctionality("BusinessObjectScope")<>"" Then
		arrBOScope=Split(dicFunctionality("BusinessObjectScope"),"~")
		For iCounter=0 To UBound(arrBOScope)
			Call Fn_Button_Click("Fn_BMIDE_CreateNewFunctionality", objNewFunctionality, "AddBOScope")
			Call Fn_Edit_Box("Fn_BMIDE_CreateNewFunctionality",JavaWindow("Business Modeler").JavaWindow("NewBusinessObjectDefineScope"),"BusinessObjectScope",arrBOScope(iCounter))
			Call Fn_Button_Click("Fn_BMIDE_CreateNewFunctionality", JavaWindow("Business Modeler").JavaWindow("NewBusinessObjectDefineScope"), "Finish")
		Next
	End If
	'Selecting Supported Condition Signature
	If dicFunctionality("SupportedConditionSignature")<>"" Then
		arrSCSingature=Split(dicFunctionality("SupportedConditionSignature"),"~")
		Call Fn_Button_Click("Fn_BMIDE_CreateNewFunctionality", objNewFunctionality, "AddSCSignature")
		For iCounter=0 To UBound(arrSCSingature)
			arrParameters=Split(arrSCSingature(iCounter),":")
			Call Fn_Button_Click("Fn_BMIDE_CreateNewFunctionality", JavaWindow("Business Modeler").JavaWindow("ConditionCustomParameters"), "Add")
			Call Fn_Edit_Box("Fn_BMIDE_CreateNewFunctionality",JavaWindow("Business Modeler").JavaWindow("NewConditionParameters"),"ParameterType",arrParameters(0))
			Call Fn_Edit_Box("Fn_BMIDE_CreateNewFunctionality",JavaWindow("Business Modeler").JavaWindow("NewConditionParameters"),"ParameterName",arrParameters(1))
			Call Fn_Button_Click("Fn_BMIDE_CreateNewFunctionality", JavaWindow("Business Modeler").JavaWindow("NewConditionParameters"), "Finish")
		Next
		Call Fn_Button_Click("Fn_BMIDE_CreateNewFunctionality", JavaWindow("Business Modeler").JavaWindow("ConditionCustomParameters"), "Finish")
	End If

	'Setting Sub Group LOV
	If dicFunctionality("SubGroupLOV")<>"" Then				' Nitish B. 06 July 2015
		'Call Fn_Edit_Box("Fn_BMIDE_CreateNewFunctionality",objNewFunctionality,"SubGroupLOV",dicFunctionality("SubGroupLOV"))
		Call Fn_Button_Click("",objNewFunctionality,"Browse...")
		Call Fn_SISW_UI_JavaEdit_Operations("","Set",JavaWindow("Business Modeler").JavaWindow("FindLOV"),"LOVName",dicFunctionality("SubGroupLOV"))
		Call Fn_Button_Click("",JavaWindow("Business Modeler").JavaWindow("FindLOV"),"OK")
	End If
	Call Fn_Button_Click("Fn_BMIDE_CreateNewFunctionality", objNewFunctionality, "Finish")
	Set objNewFunctionality=Nothing
	Fn_BMIDE_CreateNewFunctionality=True
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_BMIDE_DeepCopyRuleOperationsExt

'Description			 :	Function Used to perform operations on Deep Copy Rules

'Parameters			   :   '1.StrAction: Action Name
'										 2.dicDeepCopyRuleInfo: Deep Copy Rule Information 
'
'Return Value		   : 	True Or False

'Pre-requisite			:	Editor Should be Open

'Examples				:   dicDeepCopyRuleInfo("OperationType")="SaveAs"
'										dicDeepCopyRuleInfo("PropertyType")="Relation"
'										dicDeepCopyRuleInfo("RelationType")="3DMarkup"
'										dicDeepCopyRuleInfo("ObjectType")="AbsOccData"
'										dicDeepCopyRuleInfo("Condition")="isTrue"
'										dicDeepCopyRuleInfo("ActionType")="Select"
'										dicDeepCopyRuleInfo("TargetPrimary")="ON"
'										dicDeepCopyRuleInfo("CopyPropertiesOnRelation")="ON"
'										dicDeepCopyRuleInfo("Required")="ON"
'										dicDeepCopyRuleInfo("Secured")="ON"

'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												18-Jan-2012								1.0																						Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Public Function Fn_BMIDE_DeepCopyRuleOperationsExt(StrAction,dicDeepCopyRuleInfo)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_DeepCopyRuleOperationsExt"
	'Variable Declaration
	Dim ObjDeepCopyRule,ObjAttachWindow
	'Function Returns False
	Fn_BMIDE_DeepCopyRuleOperationsExt=False
	'Activating "Deep Copy Rules Tab"
'	Call Fn_BMIDE_InnerTabOperations("Activate","Deep Copy Rules")
	Select Case strAction
		'"Add" Case to Add New Deep Copy Rule
		Case "Add"
			'Clicking On "AddDeepCopyRule" Button to Invoke "DeepCopyRule" Window
			Call Fn_Button_Click("Fn_BMIDE_DeepCopyRuleOperationsExt", JavaWindow("Business Modeler"), "AddDeepCopyRule")
			'Checking Existance of "DeepCopyRule" Window
			If Fn_UI_ObjectExist("Fn_BMIDE_DeepCopyRuleOperationsExt",JavaWindow("Business Modeler").JavaWindow("DeepCopyRule") )=True Then
				'Creating Object Of "DeepCopyRule" Window
				Set ObjDeepCopyRule=Fn_UI_ObjectCreate("Fn_BMIDE_DeepCopyRuleOperationsExt",JavaWindow("Business Modeler").JavaWindow("DeepCopyRule") )
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Successfully Invoked the DeepCopyRule Dialog")
			Else
				Exit Function
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Invok the DeepCopyRule Dialog")				
			End If
			If dicDeepCopyRuleInfo("OperationType")<>"" Then
				'Selecting Operation Type 
				Call Fn_List_Select("Fn_BMIDE_DeepCopyRuleOperationsExt", ObjDeepCopyRule, "OperationType",dicDeepCopyRuleInfo("OperationType"))
			End If
			'Setting Property Type
			If dicDeepCopyRuleInfo("PropertyType")<>"" Then
				ObjDeepCopyRule.JavaRadioButton("PropertyType").SetTOProperty "attached text",dicDeepCopyRuleInfo("PropertyType")
				Call Fn_UI_JavaRadioButton_SetON("Fn_BMIDE_DeepCopyRuleOperationsExt",ObjDeepCopyRule, "PropertyType")
			End If
			If dicDeepCopyRuleInfo("TargetPrimary")<>"" Then
				'Setting Status Of Target Primary Option
				Call Fn_CheckBox_Set("Fn_BMIDE_DeepCopyRuleOperationsExt", ObjDeepCopyRule, "TargetPrimary", dicDeepCopyRuleInfo("TargetPrimary"))
			End If
			If dicDeepCopyRuleInfo("PropertyType")= "Reference" Then
				If dicDeepCopyRuleInfo("ReferenceProperty")<>"" Then
					JavaWindow("Business Modeler").JavaWindow("DeepCopyRule").JavaStaticText("Relation").SetTOProperty "label","Reference Property.*"
					Call Fn_Button_Click("", JavaWindow("Business Modeler").JavaWindow("DeepCopyRule"), "RelationBrowse")
					Wait(3)
					Set ObjAttachWindow=Fn_UI_ObjectCreate("Fn_BMIDE_DeepCopyRuleOperationsExt",JavaWindow("Business Modeler").JavaWindow("PropertyAttachment"))
					Call Fn_Edit_Box("Fn_BMIDE_DeepCopyRuleOperationsExt",ObjAttachWindow,"Project",dicDeepCopyRuleInfo("ReferenceProperty"))
					Wait(2)
					Call Fn_Edit_Box("Fn_BMIDE_DeepCopyRuleOperationsExt",ObjAttachWindow,"TypeSelection",dicDeepCopyRuleInfo("ObjectType"))
					Call Fn_Button_Click("Fn_BMIDE_DeepCopyRuleOperationsExt", ObjAttachWindow, "OK")		
					JavaWindow("Business Modeler").JavaWindow("DeepCopyRule").JavaStaticText("Relation").SetTOProperty "label","Relation.*"
				End If
			Else
				If dicDeepCopyRuleInfo("RelationType")<>"" Then
					'Setting Relation Type
					Call Fn_Edit_Box("Fn_BMIDE_DeepCopyRuleOperationsExt",ObjDeepCopyRule,"RelationType",dicDeepCopyRuleInfo("RelationType"))
				End If
				If dicDeepCopyRuleInfo("ObjectType")<>"" Then
					'Setting Object Type
					Call Fn_Edit_Box("Fn_BMIDE_DeepCopyRuleOperationsExt",ObjDeepCopyRule,"ObjectType",dicDeepCopyRuleInfo("ObjectType"))
				End If
			End If
			If dicDeepCopyRuleInfo("Condition")<>"" Then
				'Setting Condition
				Call Fn_Edit_Box("Fn_BMIDE_DeepCopyRuleOperationsExt",ObjDeepCopyRule,"Condition",dicDeepCopyRuleInfo("Condition"))
			End If
			If dicDeepCopyRuleInfo("ActionType")<>"" Then
				'Selecting Action
				Call Fn_List_Select("Fn_BMIDE_DeepCopyRuleOperationsExt", ObjDeepCopyRule, "Action",dicDeepCopyRuleInfo("ActionType"))
			End If

			If dicDeepCopyRuleInfo("CopyPropertiesOnRelation")<>"" Then
				'Setting Status Of "Copy Properties On Relation" Option
				Call Fn_CheckBox_Set("Fn_BMIDE_DeepCopyRuleOperationsExt", ObjDeepCopyRule, "CopyPropertiesOnRelation", dicDeepCopyRuleInfo("CopyPropertiesOnRelation"))
			End If
			If dicDeepCopyRuleInfo("Required")<>"" Then
				'Setting Status Of "Required" Option
				Call Fn_CheckBox_Set("Fn_BMIDE_DeepCopyRuleOperationsExt", ObjDeepCopyRule, "Required", dicDeepCopyRuleInfo("Required"))
			End If
			If dicDeepCopyRuleInfo("Secured")<>"" Then
				'Setting Status Of "Secured" option
				Call Fn_CheckBox_Set("Fn_BMIDE_DeepCopyRuleOperationsExt", ObjDeepCopyRule, "Secured", dicDeepCopyRuleInfo("Secured"))
			End If
			'Clicking On Finish Button To Create "Deep Copy Rule"
			Call Fn_Button_Click("Fn_BMIDE_DeepCopyRuleOperationsExt", ObjDeepCopyRule, "Finish")
			'Function Returns True
			Fn_BMIDE_DeepCopyRuleOperationsExt=True
			'Releasing Object Of "DeepCopyRule" Dialog
			Set ObjDeepCopyRule=Nothing
        End Select
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_BMIDE_OperationInputPropertyTableOperations

'Description			 :	Function Used to perform operations on Operation Input Property Table

'Parameters			   :   '1.StrAction: Action Name
'										 2.dicOpsInputPropInfo: Operation Input Property Information 
'
'Return Value		   : 	True Or False

'Pre-requisite			:	Editor Should be Open

'Examples				:   dicOpsInputPropInfo("PropertyName")="creation_date"
'										dicOpsInputPropInfo("Required")="off"
'										dicOpsInputPropInfo("Visible")="on"
'										dicOpsInputPropInfo("Usage")="None"
'										dicOpsInputPropInfo("Description")="Creation date add"
'										Call Fn_BMIDE_OperationInputPropertyTableOperations("AddPropertyFromBO",dicOpsInputPropInfo)
'
'										dicOpsInputPropInfo("PropertyName")="t3Prop1"
'										dicOpsInputPropInfo("DisplayName")="Prop1"
'										dicOpsInputPropInfo("AttributeType")="String"
'										dicOpsInputPropInfo("StringLength")="64"
'										dicOpsInputPropInfo("Description")="New property added"
',										Call Fn_BMIDE_OperationInputPropertyTableOperations("AddRuntimeProperty",dicOpsInputPropInfo)
'
'										dicOpsInputPropInfo("PropertyName")="t3Prop1"
'										Call Fn_BMIDE_OperationInputPropertyTableOperations("Select",dicOpsInputPropInfo)
'
'										dicOpsInputPropInfo("PropertyName")="t3Prop1"
'										Call Fn_BMIDE_OperationInputPropertyTableOperations("Remove",dicOpsInputPropInfo)
'
'										dicOpsInputPropInfo("PropertyName")="current_id"
'										dicOpsInputPropInfo("Required")="on"
'										dicOpsInputPropInfo("Visible")="off"
'										dicOpsInputPropInfo("CopyFromOriginal")="off"
'										dicOpsInputPropInfo("Description")="Edited Description"
'										Call Fn_BMIDE_OperationInputPropertyTableOperations("EditPropertyFromBO",dicOpsInputPropInfo)
'
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done								Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												30-Jan-2012								1.0																								Sunny R
'													Sandeep N												10-Feb-2012								1.1				Added Case "EditPropertyFromBO"			Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Public Function Fn_BMIDE_OperationInputPropertyTableOperations(StrAction,dicOpsInputPropInfo)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_OperationInputPropertyTableOperations"
    'Variable Declaration
	Dim ObjDialog,iRowcount,iCounter,crrPropname,intRowCount,intCounter,strPropertyName,bFlag
	Fn_BMIDE_OperationInputPropertyTableOperations=False
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	Select Case StrAction
		 ' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		 'Case to Add Operation Input Property from Business Object
		 Case "AddPropertyFromBO"
				Set ObjDialog=JavaWindow("Business Modeler").JavaWindow("NewOperationInputProperty")
				If Not ObjDialog.Exist(5) Then
					'Activating [ Operation Descriptor ] tab
					Call Fn_BMIDE_InnerTabOperations("Activate","Operation Descriptor")
					'Clicking On "AddOperationInputProperty" to Open "New operationInput Property" Dialog
					Call Fn_Button_Click("Fn_BMIDE_OperationInputPropertyTableOperations", JavaWindow("Business Modeler"), "AddOperationInputProperty")
				End If
				
				'Selecting property Option [ Add a Property from Business Object ]
				Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_BMIDE_OperationInputPropertyTableOperations",ObjDialog.JavaRadioButton("AddPropertyFromBusinessObject"),"attached text","Add a Property from Business Object")
				'Selecting "Add a Property from Business Object" to Add properties from Business Objects
				Call Fn_UI_JavaRadioButton_SetON("Fn_BMIDE_OperationInputPropertyOperations",ObjDialog, "AddPropertyFromBusinessObject")
				'Clicking On Next button
				If Fn_UI_ObjectExist("Fn_BMIDE_OperationInputPropertyTableOperations",ObjDialog.JavaButton("Next"))=True Then
					'Clicking on Next button
					Call Fn_Button_Click("Fn_BMIDE_OperationInputPropertyTableOperations", ObjDialog,"Next")
				End If
				'Selecting [ Property ]
                If dicOpsInputPropInfo("PropertyName")<>"" Then
					Call Fn_Edit_Box("Fn_BMIDE_OperationInputPropertyTableOperations",ObjDialog,"PropertyName",dicOpsInputPropInfo("PropertyName"))
				End If
				'Setting Required option
				If dicOpsInputPropInfo("Required")<>"" Then
					Call Fn_CheckBox_Set("Fn_BMIDE_OperationInputPropertyTableOperations", ObjDialog, "Required", dicOpsInputPropInfo("Required"))
				End If
				'Setting Visible option
				If dicOpsInputPropInfo("Visible")<>"" Then
					Call Fn_CheckBox_Set("Fn_BMIDE_OperationInputPropertyTableOperations", ObjDialog, "Visible", dicOpsInputPropInfo("Visible"))
				End If
				'Setting Copy From Original option
				If dicOpsInputPropInfo("CopyFromOriginal")<>"" Then
					Call Fn_CheckBox_Set("Fn_BMIDE_OperationInputPropertyTableOperations", ObjDialog, "CopyFromOriginal", dicOpsInputPropInfo("CopyFromOriginal"))
				End If

				'Setting Usage Option
				If dicOpsInputPropInfo("Usage")<>"" Then
					Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_BMIDE_OperationInputPropertyTableOperations",ObjDialog.JavaRadioButton("Usage"),"attached text",dicOpsInputPropInfo("Usage"))
					Call Fn_UI_JavaRadioButton_SetON("Fn_BMIDE_OperationInputPropertyTableOperations",ObjDialog, "Usage")
				End If
				'Setting Compound Object Type
				If dicOpsInputPropInfo("CompoundObjectType")<>"" Then
					Call Fn_Edit_Box("Fn_BMIDE_OperationInputPropertyTableOperations",ObjDialog,"CompoundObjectType",dicOpsInputPropInfo("CompoundObjectType"))
				End If
				'Setting Compound Object Constant
				If dicOpsInputPropInfo("CompoundObjectConstant")<>"" Then
					Call Fn_Edit_Box("Fn_BMIDE_OperationInputPropertyTableOperations",ObjDialog,"CompoundObjectConstant",dicOpsInputPropInfo("CompoundObjectConstant"))
				End If
				'Setting Description
				If dicOpsInputPropInfo("Description")<>"" Then
					Call Fn_Edit_Box("Fn_BMIDE_OperationInputPropertyTableOperations",ObjDialog,"Description",dicOpsInputPropInfo("Description"))
				End If
				'Clicking on Finish button
				Call Fn_Button_Click("Fn_BMIDE_OperationInputPropertyTableOperations", ObjDialog, "Finish")
				Fn_BMIDE_OperationInputPropertyTableOperations=True
		   ' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		   'Case to Define and Add a new Runtime property from Business Object
		   Case "AddRuntimeProperty"
				Set ObjDialog=JavaWindow("Business Modeler").JavaWindow("NewOperationInputProperty")
				If Not ObjDialog.Exist(5) Then
					'Activating [ Operation Descriptor ] tab
					Call Fn_BMIDE_InnerTabOperations("Activate","Operation Descriptor")
					'Clicking On "AddOperationInputProperty" to Open "New operationInput Property" Dialog
					Call Fn_Button_Click("Fn_BMIDE_OperationInputPropertyTableOperations", JavaWindow("Business Modeler"), "AddOperationInputProperty")
				End If
				'Selecting property Option [ Add a Property from Business Object ]
				Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_BMIDE_OperationInputPropertyTableOperations",ObjDialog.JavaRadioButton("AddPropertyFromBusinessObject"),"attached text","Define and add a new Runtime Property from Business Object")
				'Selecting "Define and add a new Runtime Property from Business Object" to Add properties from Business Objects
				Call Fn_UI_JavaRadioButton_SetON("Fn_BMIDE_OperationInputPropertyOperations",ObjDialog, "AddPropertyFromBusinessObject")
				'Clicking On Next button
				If Fn_UI_ObjectExist("Fn_BMIDE_OperationInputPropertyTableOperations",ObjDialog.JavaButton("Next"))=True Then
					'Clicking on Next button
					Call Fn_Button_Click("Fn_BMIDE_OperationInputPropertyTableOperations", ObjDialog,"Next")
				End If
				'Setting New property Name
                If dicOpsInputPropInfo("PropertyName")<>"" Then
					Call Fn_Edit_Box("Fn_BMIDE_OperationInputPropertyTableOperations",ObjDialog,"Name",dicOpsInputPropInfo("PropertyName"))
				End If
				'Setting New property Display Name
                If dicOpsInputPropInfo("DisplayName")<>"" Then
					Call Fn_Edit_Box("Fn_BMIDE_OperationInputPropertyTableOperations",ObjDialog,"DisplayName",dicOpsInputPropInfo("DisplayName"))
				End If
				'Selecting Attribute Type
                If dicOpsInputPropInfo("AttributeType")<>"" Then
					Call Fn_List_Select("Fn_BMIDE_OperationInputPropertyTableOperations", ObjDialog, "AttributeType",dicOpsInputPropInfo("AttributeType"))
				End If
				'Setting String Length
                If dicOpsInputPropInfo("StringLength")<>"" Then
					Call Fn_Edit_Box("Fn_BMIDE_OperationInputPropertyTableOperations",ObjDialog,"StringLength",dicOpsInputPropInfo("StringLength"))
				End If
				'Setting Reference Business Object
                If dicOpsInputPropInfo("ReferenceBusinessObject")<>"" Then
					Call Fn_Edit_Box("Fn_BMIDE_OperationInputPropertyTableOperations",ObjDialog,"ReferenceBusinessObject",dicOpsInputPropInfo("ReferenceBusinessObject"))
				End If
				'Setting Description
				If dicOpsInputPropInfo("Description")<>"" Then
					Call Fn_Edit_Box("Fn_BMIDE_OperationInputPropertyTableOperations",ObjDialog,"PropDescription",dicOpsInputPropInfo("Description"))
				End If
				'Setting [ Array ] options
				If dicOpsInputPropInfo("Array")<>"" Then
					Call Fn_CheckBox_Set("Fn_BMIDE_OperationInputPropertyTableOperations", ObjDialog, "Array", dicOpsInputPropInfo("Array"))
					If LCase(dicOpsInputPropInfo("Array"))="on" Then
						'Setting Unlimited Array option
						If dicOpsInputPropInfo("Unlimited")<>"" Then
							Call Fn_CheckBox_Set("Fn_BMIDE_OperationInputPropertyTableOperations", ObjDialog, "Unlimited", dicOpsInputPropInfo("Unlimited"))
						End If
						'Setting Array Max Length
						If dicOpsInputPropInfo("MaxLength")<>"" Then
							Call Fn_Edit_Box("Fn_BMIDE_OperationInputPropertyTableOperations",ObjDialog,"MaxLength",dicOpsInputPropInfo("MaxLength"))
						End If
					End If
				End If
				'Clicking on Finish button
				Call Fn_Button_Click("Fn_BMIDE_OperationInputPropertyTableOperations", ObjDialog, "Finish")
				Fn_BMIDE_OperationInputPropertyTableOperations=True
		   ' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		   'Case to select property from Operation Input Property Table
		   Case "Select"
			   Call Fn_BMIDE_InnerTabOperations("Activate","Operation Descriptor")
			   iRowcount=Fn_UI_Object_GetROProperty("Fn_BMIDE_OperationInputPropertyTableOperations",JavaWindow("Business Modeler").JavaTable("OperationInputPropertyTable"), "rows")
              For iCounter=0 to iRowcount-1
				crrPropname=JavaWindow("Business Modeler").JavaTable("OperationInputPropertyTable").GetCellData(iCounter,"Name")
				If Trim(crrPropname)=Trim(dicOpsInputPropInfo("PropertyName")) Then
					 JavaWindow("Business Modeler").JavaTable("OperationInputPropertyTable").SelectCell iCounter,0
					 Fn_BMIDE_OperationInputPropertyTableOperations=True
					 Exit For
				End If
			  Next
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		   'Case to remove property from Operation Input Property Table
		   Case "Remove"
			  Call Fn_BMIDE_InnerTabOperations("Activate","Operation Descriptor")
			   iRowcount=Fn_UI_Object_GetROProperty("Fn_BMIDE_OperationInputPropertyTableOperations",JavaWindow("Business Modeler").JavaTable("OperationInputPropertyTable"), "rows")
              For iCounter=0 to iRowcount-1
				crrPropname=JavaWindow("Business Modeler").JavaTable("OperationInputPropertyTable").GetCellData(iCounter,"Name")
				If Trim(crrPropname)=Trim(dicOpsInputPropInfo("PropertyName")) Then
					 JavaWindow("Business Modeler").JavaTable("OperationInputPropertyTable").SelectCell iCounter,0
					 wait 1
					 'Clicking On "RemoveOperationInputProperty" to Remove property from Operation Input Property Table
					 Call Fn_Button_Click("Fn_BMIDE_OperationInputPropertyTableOperations", JavaWindow("Business Modeler"), "RemoveOperationInputProperty")
					 Fn_BMIDE_OperationInputPropertyTableOperations=Fn_BMIDE_DeleteObject()
					 Exit For
				End If
			  Next
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "EditPropertyFromBO"		'Case To Modify Business Object OperationInput Property
				Set ObjDialog=JavaWindow("Business Modeler").JavaWindow("NewOperationInputProperty")
				bFlag=False
                intRowCount=Fn_UI_Object_GetROProperty("Fn_BMIDE_OperationInputPropertyOperations",JavaWindow("Business Modeler").JavaTable("OperationInputPropertyTable"), "rows")
				For intCounter=0 To intRowCount-1
					strPropertyName=JavaWindow("Business Modeler").JavaTable("OperationInputPropertyTable").GetCellData(intCounter,"Name")
					If Trim(dicOpsInputPropInfo("PropertyName"))=Trim(strPropertyName) Then
						JavaWindow("Business Modeler").JavaTable("OperationInputPropertyTable").SelectCell intCounter,0
						bFlag=True
						Exit For
					End If
				Next
				If bFlag=False Then
					Exit Function
				End If
				Call Fn_Button_Click("Fn_BMIDE_OperationInputPropertyOperations", JavaWindow("Business Modeler"), "EditOperationInputProperty")
				Set ObjDilog=Fn_UI_ObjectCreate("Fn_BMIDE_OperationInputPropertyOperations", JavaWindow("Business Modeler").JavaWindow("NewOperationInputProperty"))

				'Editing Required option
				If dicOpsInputPropInfo("Required")<>"" Then
					Call Fn_CheckBox_Set("Fn_BMIDE_OperationInputPropertyTableOperations", ObjDialog, "Required", dicOpsInputPropInfo("Required"))
				End If
				'Editing Visible option
				If dicOpsInputPropInfo("Visible")<>"" Then
					Call Fn_CheckBox_Set("Fn_BMIDE_OperationInputPropertyTableOperations", ObjDialog, "Visible", dicOpsInputPropInfo("Visible"))
				End If
				'Editing Copy From Original option
				If dicOpsInputPropInfo("CopyFromOriginal")<>"" Then
					Call Fn_CheckBox_Set("Fn_BMIDE_OperationInputPropertyTableOperations", ObjDialog, "CopyFromOriginal", dicOpsInputPropInfo("CopyFromOriginal"))
				End If
				'Editing Description
				If dicOpsInputPropInfo("Description")<>"" Then
					Call Fn_Edit_Box("Fn_BMIDE_OperationInputPropertyTableOperations",ObjDialog,"Description",dicOpsInputPropInfo("Description"))
				End If
				'Clicking on Finish button
				Call Fn_Button_Click("Fn_BMIDE_OperationInputPropertyTableOperations", ObjDialog, "Finish")
				Fn_BMIDE_OperationInputPropertyTableOperations=True
	End Select
	Set ObjDialog=Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_SISW_BMIDE_ConvertBusinessObject

'Description			 :	Function Used to conver Business Object Primary to Secondary and vice versa

'Parameters			   :   1.StrAction: Action Name
'										2.StrMsg: Convrsion Message
'										3.StrButton: Button Name
'
'Return Value		   : 	true or false

'Pre-requisite			:	Confirm Business Object Conversion dialog should appear on screen

'Examples				:   bReturn=Fn_SISW_BMIDE_ConvertBusinessObject("Convert","","")
'
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												17-May-2012								1.0																						Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_SISW_BMIDE_ConvertBusinessObject(StrAction,StrMsg,StrButton)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_BMIDE_ConvertBusinessObject"
	Fn_SISW_BMIDE_ConvertBusinessObject=false
	'Checking existance of [ ConfirmBusinessObjectConversion ] dialog
	If JavaWindow("Business Modeler").JavaWindow("ConfirmBusinessObjectConversion").Exist(6) Then
		Select Case StrAction
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
			Case "Convert"
				'Clicking on [ OK ] button
				Call Fn_Button_Click("Fn_SISW_BMIDE_ConvertBusinessObject",JavaWindow("Business Modeler").JavaWindow("ConfirmBusinessObjectConversion"),"OK")
				'Checking existance of [ BusinessObjectConverted ]
				If JavaWindow("Business Modeler").JavaWindow("BusinessObjectConverted").Exist(6) Then
					'Clicking on [ OK ] button
					Call Fn_Button_Click("Fn_SISW_BMIDE_ConvertBusinessObject",JavaWindow("Business Modeler").JavaWindow("BusinessObjectConverted"),"OK")
					Fn_SISW_BMIDE_ConvertBusinessObject=true
				End If
		End Select
	End If
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_SISW_BMIDE_NamingRuleAttachmentsOperations

'Description			 :	Function Used to perform operations on Naming Rule Attachments

'Parameters			   :   1.StrAction: Action Name
'										2.StrBusinessObjectProperty: Business Object name . Property name
'										3.StrProperty: Property
'										4.StrCase: Case
'										5.StrCondition: Condition
'										6.bOverride: Override option
'										7.StrColName: column name
'										8.StrValue: expected value
'
'Return Value		   : 	true or false

'Pre-requisite			:	Naming Rule attachments table should be appear

'Examples				:   bReturn=Fn_SISW_BMIDE_NamingRuleAttachmentsOperations("Select","L3_TestItem.object_name","","","","","","")
'										bReturn=Fn_SISW_BMIDE_NamingRuleAttachmentsOperations("Detach","L3_TestItem.object_name","","","","","","")
'										bReturn=Fn_SISW_BMIDE_NamingRuleAttachmentsOperations("Attach","","L3_TestItem.object_name","Upper","isTrue","off","","")
'
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												30-Jul-2012								1.0																						Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												30-Jul-2012								1.0																						Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_SISW_BMIDE_NamingRuleAttachmentsOperations(StrAction,StrBusinessObjectProperty,StrProperty,StrCase,StrCondition,bOverride,StrColName,StrValue)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_BMIDE_NamingRuleAttachmentsOperations"
 	'Declaring variables
	Dim objAttachmentsDialog,objAttachmentsTable
	Dim bFlag,iRows,iCounter,crrBOObjectProperty
	'Creating object of [ NamingRuleAttachment ] dialog and [ NamingRuleAttachment ] table
	Set objAttachmentsDialog=JavaWindow("Business Modeler").JavaWindow("NamingRuleAttachment")
	Set objAttachmentsTable=JavaWindow("Business Modeler").JavaTable("NamingRuleAttachments")
	Select Case StrAction
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
			Case "Select"
				bFlag=false
				'Retriving number of rows exist in table
				iRows=Fn_UI_Object_GetROProperty("Fn_SISW_BMIDE_NamingRuleAttachmentsOperations",objAttachmentsTable, "rows")
				For iCounter=0 to iRows-1
					'retriving Business Object.Property from table
					crrBOObjectProperty=objAttachmentsTable.GetCellData(iCounter,"Business Object.Property")
					'Checking current Business Object.Property matches with expected Business Object.Property
					If trim(crrBOObjectProperty)=trim(StrBusinessObjectProperty) Then
						'Selecting expected Business Object.Property
						objAttachmentsTable.SelectCell iCounter, 0
						bFlag=true
						Exit for
					End If
				Next
				If bFlag=true Then
					Fn_SISW_BMIDE_NamingRuleAttachmentsOperations=true
				End If
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
			Case "Detach"
				bFlag=Fn_SISW_BMIDE_NamingRuleAttachmentsOperations("Select",StrBusinessObjectProperty,"","","","","","")
				If bFlag=true Then
					'detaching naming rule attachment
					Fn_SISW_BMIDE_NamingRuleAttachmentsOperations=Fn_Button_Click("Fn_SISW_BMIDE_NamingRuleAttachmentsOperations",JavaWindow("Business Modeler"), "DetachNamingRuleAttachments")
				End If
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
			Case "Attach"
				'Checking existance of [ Naming Rule Attachments ] dilaog
				If not objAttachmentsDialog.Exist(6) Then
					'Clikcing on [ Attach Naming Rule Attachments ] button to invoke [ Naming Rule Attachments ] dilaog
					Call Fn_Button_Click("Fn_SISW_BMIDE_NamingRuleAttachmentsOperations",JavaWindow("Business Modeler"), "AttachNamingRuleAttachments")
				End If
				'Setting Property name
				'Call Fn_Edit_Box("Fn_SISW_BMIDE_NamingRuleAttachmentsOperations",objAttachmentsDialog,"Property",StrProperty)
				arrProperty=split(StrProperty,".")
				Call Fn_Button_Click("Fn_SISW_BMIDE_NamingRuleAttachmentsOperations",objAttachmentsDialog, "Browse...")
				Call Fn_Edit_Box("Fn_SISW_BMIDE_NamingRuleAttachmentsOperations",JavaWindow("Business Modeler").JavaWindow("PropertyAttachment"),"Project",arrproperty(0))
				Call Fn_Edit_Box("Fn_SISW_BMIDE_NamingRuleAttachmentsOperations",JavaWindow("Business Modeler").JavaWindow("PropertyAttachment"),"Properties",arrproperty(1))
				Call Fn_Button_Click("Fn_SISW_BMIDE_NamingRuleAttachmentsOperations",JavaWindow("Business Modeler").JavaWindow("PropertyAttachment"), "OK")
				
				If StrCase<>"" Then
					'Selecting case
					Call Fn_List_Select("Fn_SISW_BMIDE_NamingRuleAttachmentsOperations",objAttachmentsDialog,"Case",StrCase)
				End If
				If StrCondition<>"" Then
					'Setting condition
					Call Fn_Edit_Box("Fn_SISW_BMIDE_NamingRuleAttachmentsOperations",objAttachmentsDialog,"Condition",StrCondition)
				End If
				If bOverride<>"" Then
					'setting override option
					Call Fn_CheckBox_Set("Fn_SISW_BMIDE_NamingRuleAttachmentsOperations",objAttachmentsDialog, "Override", bOverride)
				End If
				Fn_SISW_BMIDE_NamingRuleAttachmentsOperations=Fn_Button_Click("Fn_SISW_BMIDE_NamingRuleAttachmentsOperations",objAttachmentsDialog, "Finish")
	End Select
	Set objAttachmentsDialog=nothing
	Set objAttachmentsTable=nothing
End Function

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_SISW_BMIDE_NewEventType

'Description			 :	Function Used to create Event Type

'Parameters			   :   1.StrProject: Project name [ not mandetory field ]
'										2.StrID: Event Type ID
'										3.StrDisplayName: Event Type Display name
'										4.StrDescription: Event Type Description
'										5.StrButton: Button names
'
'Return Value		   : 	True or False

'Pre-requisite			:	New Event Type dialog should be exist

'Examples				:   bReturn=Fn_SISW_BMIDE_NewEventType("","00024","Event24","New event","")
'										bReturn=Fn_SISW_BMIDE_NewEventType("","00025","Event25","New event","Apply:Cancel")
'								
'
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												11-Sep-2012								1.0																						Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_SISW_BMIDE_NewEventType(StrProject,StrID,StrDisplayName,StrDescription,StrButton)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_BMIDE_NewEventType"
 	'Declaring variables
	Dim ObjEventType
	Dim StrPrefix,iCounter

	Fn_SISW_BMIDE_NewEventType=False
   'Checking existance of [ NewEventType ] dialog
   If not JavaWindow("Business Modeler").JavaWindow("NewEventType").Exist(6) Then
	   Exit function
   End If
   'Creating object of [ NewEventType ] dialog
   Set ObjEventType=JavaWindow("Business Modeler").JavaWindow("NewEventType")
   'Selecting project
   If StrProject<>"" Then
		Call Fn_List_Select("Fn_SISW_BMIDE_NewEventType", ObjEventType, "Project",StrProject)
   End If
   'Setting Event type ID
   If StrID<>"" Then
	   'Retriving current project Prefix
		StrPrefix= Fn_Edit_Box_GetValue("Fn_SISW_BMIDE_NewEventType",ObjEventType,"ID")
		StrID=StrPrefix+StrID
		Call Fn_Edit_Box("Fn_SISW_BMIDE_NewEventType",ObjEventType,"ID",StrID)
   End If
   'Setting Event type Display name
   If StrDisplayName<>"" Then
		Call Fn_Edit_Box("Fn_SISW_BMIDE_NewEventType",ObjEventType,"DisplayName",StrDisplayName)
   End If
	'Setting Event type Description
   If StrDescription<>"" Then
		Call Fn_Edit_Box("Fn_SISW_BMIDE_NewEventType",ObjEventType,"Description",StrDescription)
   End If
   'Clicking on button
   If StrButton<>"" Then
	   StrButton=Split(StrButton,":")
	   For iCounter=0 to ubound(StrButton)
			Call Fn_Button_Click("Fn_SISW_BMIDE_NewEventType", ObjEventType, StrButton(iCounter))
			wait 1
	   Next
	else
		Call Fn_Button_Click("Fn_SISW_BMIDE_NewEventType", ObjEventType, "Finish")
   End If
   If Err.Number < 0 Then
		Fn_SISW_BMIDE_NewEventType=False
	Else
		Fn_SISW_BMIDE_NewEventType=True
	End If
	'Releasing object of [ NewEventType ] dialog
	Set ObjEventType=nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_SISW_BMIDE_NewEventTypeMapping

'Description			 :	Function Used to create Event Type Mapping

'Parameters			   :   1.StrProject: Project name [ not mandetory field ]
'										2.StrPrimaryObject: Event Type Mapping Primary object
'										3.StrEventType: Event Type mapping Event Type
'										4.StrAuditType: Event Type Mapping Audit Type
'										5.StrSecondaryAuditType: Event Type Mapping Secondary Audit Type
'										6.bSubscribable: Event Type Mapping Subscribable option
'										7.bAuditable: Event Type Mapping Auditable option
'										8.StrDescription: Event Type Mapping Description
'										9.StrButton: Button Names to click
'
'Return Value		   : 	True or False

'Pre-requisite			:	New Event Type Mapping dialog should be exist

'Examples				:   bReturn=Fn_SISW_BMIDE_NewEventTypeMapping("","AbsOccData","L300024","Fnd0GeneralAudit","Fnd0SecondaryAudit","on","on","Event type mapping description","")
'								
'
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												11-Sep-2012								1.0																						Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_SISW_BMIDE_NewEventTypeMapping(StrProject,StrPrimaryObject,StrEventType,StrAuditType,StrSecondaryAuditType,bSubscribable,bAuditable,StrDescription,StrButton)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_BMIDE_NewEventTypeMapping"
 	'Declaring variables
	Dim ObjEventTypeMapping
	Dim iCounter

	Fn_SISW_BMIDE_NewEventTypeMapping=False
   'Checking existance of [ NewEventTypeMapping ] dialog
   If not JavaWindow("Business Modeler").JavaWindow("NewEventTypeMapping").Exist(6) Then
	   Exit function
   End If
   'Creating object of [ NewEventTypeMapping ] dialog
   Set ObjEventTypeMapping=JavaWindow("Business Modeler").JavaWindow("NewEventTypeMapping")
   'Selecting project
   If StrProject<>"" Then
		Call Fn_List_Select("Fn_SISW_BMIDE_NewEventTypeMapping", ObjEventTypeMapping, "Project",StrProject)
   End If
   'Setting Primary Object
   If StrPrimaryObject<>"" Then
		Call Fn_Edit_Box("Fn_SISW_BMIDE_NewEventTypeMapping",ObjEventTypeMapping,"PrimaryObject",StrPrimaryObject)
   End If
	'Setting Event Type
   If StrEventType<>"" Then
		Call Fn_Edit_Box("Fn_SISW_BMIDE_NewEventTypeMapping",ObjEventTypeMapping,"EventType",StrEventType)
   End If
	'Setting Audit Type
   If StrAuditType<>"" Then
		Call Fn_Edit_Box("Fn_SISW_BMIDE_NewEventTypeMapping",ObjEventTypeMapping,"AuditType",StrAuditType)
   End If
	'Setting Secondary Audit Type
   If StrSecondaryAuditType<>"" Then
		Call Fn_Edit_Box("Fn_SISW_BMIDE_NewEventTypeMapping",ObjEventTypeMapping,"SecondaryAuditType",StrSecondaryAuditType)
   End If
   'Selecting Subscribable option
   If bSubscribable<>"" Then
	   Call Fn_CheckBox_Set("Fn_SISW_BMIDE_NewEventTypeMapping", ObjEventTypeMapping, "Subscribable", bSubscribable)
   End If
	'Selecting Subscribable option
   If bAuditable<>"" Then
	   Call Fn_CheckBox_Set("Fn_SISW_BMIDE_NewEventTypeMapping", ObjEventTypeMapping, "Auditable", bAuditable)
   End If
	'Setting Event type Description
   If StrDescription<>"" Then
		Call Fn_Edit_Box("Fn_SISW_BMIDE_NewEventTypeMapping",ObjEventTypeMapping,"Description",StrDescription)
   End If
   'Clicking on button
   If StrButton<>"" Then
	   StrButton=Split(StrButton,":")
	   For iCounter=0 to ubound(StrButton)
			Call Fn_Button_Click("Fn_SISW_BMIDE_NewEventTypeMapping", ObjEventTypeMapping, StrButton(iCounter))
			wait 1
	   Next
	else
		Call Fn_Button_Click("Fn_SISW_BMIDE_NewEventTypeMapping", ObjEventTypeMapping, "Finish")
   End If
   If Err.Number < 0 Then
		Fn_SISW_BMIDE_NewEventTypeMapping=False
	Else
		Fn_SISW_BMIDE_NewEventTypeMapping=True
	End If
	'Releasing object of [ NewEventTypeMapping ] dialog
	Set ObjEventTypeMapping=nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_SISW_BMIDE_CreateAuditDefinationProperty

'Description			 :	Function Used to create Event Type Mapping

'Parameters			   :   1.StrProject: Project name [ not mandetory field ]
'										2.StrObjectType: Object Type
'										3.StrPropertyName: Property Name
'										4.StrTargetPropertyName: Target Property Name
'										5.StrReferenceDisplayPropertyName: Reference Display Property Name
'										6.StrTargetOldValuePropertyName: Target Old Value Property Name
'										7.StrReferenceOldValuePropertyName: Reference Old Value Property Name
'										8.StrEnableTracking: Enable Tracking
'										9.StrButton: Button Names to click
'
'Return Value		   : 	True or False

'Pre-requisite			:	Audit Defination Property dialog should be exist

'Examples				:   bReturn=Fn_SISW_BMIDE_CreateAuditDefinationProperty("","AbsOccData","archelemid","archelemid","","","","","")
'								
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												14-Sep-2012								1.0																						Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_SISW_BMIDE_CreateAuditDefinationProperty(StrProject,StrObjectType,StrPropertyName,StrTargetPropertyName,StrReferenceDisplayPropertyName,StrTargetOldValuePropertyName,StrReferenceOldValuePropertyName,StrEnableTracking,StrButton)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_BMIDE_CreateAuditDefinationProperty"
   'Declaring variables
   Dim ObjAuditDefinitionProperty
   Fn_SISW_BMIDE_CreateAuditDefinationProperty=false
   'Checking existance of [ AuditDefinitionProperty ] window
	If not JavaWindow("Business Modeler").JavaWindow("AuditDefinitionProperty").Exist(6) Then
		Exit function
	End If
	'creating object of [ AuditDefinitionProperty ] window
	Set ObjAuditDefinitionProperty=JavaWindow("Business Modeler").JavaWindow("AuditDefinitionProperty")
	'setting object type
	If StrObjectType<>"" Then
		Call Fn_Edit_Box("Fn_SISW_BMIDE_CreateAuditDefinationProperty",ObjAuditDefinitionProperty,"ObjectType",StrObjectType)
	End If
	'setting property name
	If StrPropertyName<>"" Then
		Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_SISW_BMIDE_CreateAuditDefinationProperty",ObjAuditDefinitionProperty.JavaStaticText("StaticText"),"label","Property Name:")
		Call Fn_Edit_Box("Fn_SISW_BMIDE_CreateAuditDefinationProperty",ObjAuditDefinitionProperty,"PropertyName",StrPropertyName)
	End If
	'setting target property name
	If StrTargetPropertyName<>"" Then
		Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_SISW_BMIDE_CreateAuditDefinationProperty",ObjAuditDefinitionProperty.JavaStaticText("StaticText"),"label","Target Property Name:")
		Call Fn_Edit_Box("Fn_SISW_BMIDE_CreateAuditDefinationProperty",ObjAuditDefinitionProperty,"TargetPropertyName",StrTargetPropertyName)
	End If
	'setting Reference Display Property Name
	If StrReferenceDisplayPropertyName<>"" Then
		Call Fn_Edit_Box("Fn_SISW_BMIDE_CreateAuditDefinationProperty",ObjAuditDefinitionProperty,"ReferenceDisplayPropertyName",StrReferenceDisplayPropertyName)
	End If
	'setting Target Old Value Property Name
	If StrTargetOldValuePropertyName<>"" Then
		Call Fn_Edit_Box("Fn_SISW_BMIDE_CreateAuditDefinationProperty",ObjAuditDefinitionProperty,"TargetOldValuePropertyName",StrTargetOldValuePropertyName)
	End If
	'setting Reference Old Value Property Name
	If StrReferenceOldValuePropertyName<>"" Then
		Call Fn_Edit_Box("Fn_SISW_BMIDE_CreateAuditDefinationProperty",ObjAuditDefinitionProperty,"ReferenceOldValuePropertyName",StrReferenceOldValuePropertyName)
	End If
	'Selectin enable tracking
	If StrEnableTracking<>"" Then
		Call Fn_List_Select("Fn_SISW_BMIDE_CreateAuditDefinationProperty", ObjAuditDefinitionProperty, "EnableTracking",StrEnableTracking)
	End If
	If StrButton<>"" Then
		Call Fn_Button_Click("Fn_SISW_BMIDE_CreateAuditDefinationProperty", ObjAuditDefinitionProperty,StrButton)
	else
		Call Fn_Button_Click("Fn_SISW_BMIDE_CreateAuditDefinationProperty", ObjAuditDefinitionProperty, "Finish")
	End If
	 If Err.Number < 0 Then
		Fn_SISW_BMIDE_CreateAuditDefinationProperty=False
	Else
		Fn_SISW_BMIDE_CreateAuditDefinationProperty=True
	End If
	Set ObjAuditDefinitionProperty=Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_SISW_BMIDE_NewAuditDefinition

'Description			 :	Function Used to create new Audit Definition

'Parameters			   :   1.dicAuditDefinitionInfo: New Audit Definition information
'
'Return Value		   : 	True or False

'Pre-requisite			:	New Audit Defination dialog should be exist

'Examples				:   Dim dicAuditDefinitionInfo
'										Set dicAuditDefinitionInfo=CreateObject("Scripting.Dictionary")
'										With dicAuditDefinitionInfo  
'													 .Add "PrimaryObject",""
'													 .Add "EventType",""
'													 .Add "AuditExtensions",""
'													 .Add "Description",""
'													 .Add "IsActive",""    ' On or Off
'													 .Add "TrackOldValues",""    ' On or Off
'													 .Add "AuditOnPropertyChange",""    ' On or Off
'													 .Add "PrimaryObjectAuditDefinitionProperty",""    ' Primary Object Audit Definition Properties : value separeted with ~ and property separated with $... eg: ~~absocc_attr_name~absocc_attr_name~~~~$~~childbv~childbv~childbvDisp~~~
'													 .Add "SecondaryObjectAuditDefinitionProperty",""    ' Secondary Object Audit Definition Properties : value separeted with ~ and property separated with $... eg: ~~absocc_attr_name~absocc_attr_name~~~~$~~childbv~childbv~childbvDisp~~~
'										End with
'										
'										dicAuditDefinitionInfo("PrimaryObject")="AbsOccData"
'										dicAuditDefinitionInfo("EventType")="L300024"
'										dicAuditDefinitionInfo("AuditExtensions")="Fnd0CICO_auditloghandler~Fnd0WriteSignoffDetails"
'										dicAuditDefinitionInfo("Description")="New audit definition"
'										dicAuditDefinitionInfo("IsActive")="on"
'										dicAuditDefinitionInfo("PrimaryObjectAuditDefinitionProperty")="~~absocc_attr_name~absocc_attr_name~~~~$~~childbv~childbv~childbvDisp~~~"
'										dicAuditDefinitionInfo("SecondaryObjectAuditDefinitionProperty")="~AbsOccData~absocc_attr_value~absocc_attr_value~~~~$~Bitmap~based_on~based_on~~~~"
'										
'										bReturn=Fn_SISW_BMIDE_NewAuditDefinition(dicAuditDefinitionInfo)
'								
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												14-Sep-2012								1.0																						Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_SISW_BMIDE_NewAuditDefinition(dicAuditDefinitionInfo)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_BMIDE_NewAuditDefinition"
 	'variable declaration
	Dim ObjAuditDefinition
	Dim aAuditExtensions,iCounter,iRows,iCount,bFlag,anPrimaryProperty,aPrimaryPropertyValue,anSecondaryProperty,aSecondaryPropertyValue
	Fn_SISW_BMIDE_NewAuditDefinition=false
 	'Checking existance of [ New Audit Definition... ] dialog
	If not JavaWindow("Business Modeler").JavaWindow("NewAuditDefinition").Exist(6) Then
		Exit function
	End If
	'creating object of [ New Audit Definition... ] dialog
	Set ObjAuditDefinition=JavaWindow("Business Modeler").JavaWindow("NewAuditDefinition")
	'Setting Primary Object
	If dicAuditDefinitionInfo("PrimaryObject")<>"" Then
		Call Fn_Edit_Box("Fn_SISW_BMIDE_NewAuditDefinition",ObjAuditDefinition,"PrimaryObject",dicAuditDefinitionInfo("PrimaryObject"))
	End If
	'Setting Event Type
	If dicAuditDefinitionInfo("EventType")<>"" Then
		Call Fn_Edit_Box("Fn_SISW_BMIDE_NewAuditDefinition",ObjAuditDefinition,"EventType",dicAuditDefinitionInfo("EventType"))
	End If
	'Selecting Audit Extensions
	If dicAuditDefinitionInfo("AuditExtensions")<>"" Then
		aAuditExtensions=Split(dicAuditDefinitionInfo("AuditExtensions"),"~")
		For iCounter=0 to ubound(aAuditExtensions)
			'clicking on [ Add.. ] button to add Audit Extensions
			Call Fn_Button_Click("Fn_SISW_BMIDE_NewAuditDefinition", ObjAuditDefinition, "AddAuditExtension")
			'Checking existance of [ FindAuditExtension ] dialog
			If not JavaWindow("Business Modeler").JavaWindow("FindAuditExtension").Exist(6) Then
				Set ObjAuditDefinition=Nothing
				Exit function
			End If
			iRows=JavaWindow("Business Modeler").JavaWindow("FindAuditExtension").JavaTable("AuditExtensionTypes").GetROProperty("rows")
			For iCount=0 to iRows-1
				bFlag=false
				If aAuditExtensions(iCounter)=JavaWindow("Business Modeler").JavaWindow("FindAuditExtension").JavaTable("AuditExtensionTypes").GetCellData(iCount,0) Then
					JavaWindow("Business Modeler").JavaWindow("FindAuditExtension").JavaTable("AuditExtensionTypes").ClickCell iCount,0
					wait 1
					Call Fn_Button_Click("Fn_SISW_BMIDE_NewAuditDefinition", JavaWindow("Business Modeler").JavaWindow("FindAuditExtension"), "OK")
					bFlag=true
					Exit for
				End If
			Next
			If bFlag=false Then
				Set ObjAuditDefinition=Nothing
				Exit function
			End If
		Next
	End If
	'Setting Description
	If dicAuditDefinitionInfo("Description")<>"" Then
		Call Fn_Edit_Box("Fn_SISW_BMIDE_NewAuditDefinition",ObjAuditDefinition,"Description",dicAuditDefinitionInfo("Description"))
	End If
	'setting Is Active option
	If dicAuditDefinitionInfo("IsActive")<>"" Then
		Call Fn_CheckBox_Set("Fn_SISW_BMIDE_NewAuditDefinition",ObjAuditDefinition, "IsActive", dicAuditDefinitionInfo("IsActive"))
	end if
	'setting Track Old Values option
	If dicAuditDefinitionInfo("TrackOldValues")<>"" Then
		Call Fn_CheckBox_Set("Fn_SISW_BMIDE_NewAuditDefinition",ObjAuditDefinition, "TrackOldValues", dicAuditDefinitionInfo("TrackOldValues"))
	end if
	'setting Audit On Property Change option
	If dicAuditDefinitionInfo("AuditOnPropertyChange")<>"" Then
		Call Fn_CheckBox_Set("Fn_SISW_BMIDE_NewAuditDefinition",ObjAuditDefinition, "AuditOnPropertyChange", dicAuditDefinitionInfo("AuditOnPropertyChange"))
	end if
	'Clicking on next button
	Call Fn_Button_Click("Fn_SISW_BMIDE_NewAuditDefinition", ObjAuditDefinition, "Next")	
	'Setting primary object defination properties
	If lcase(dicAuditDefinitionInfo("PrimaryObjectAuditDefinitionProperty"))<>"" then
		anPrimaryProperty=Split(dicAuditDefinitionInfo("PrimaryObjectAuditDefinitionProperty"),"$")
		For iCounter=0 to ubound(anPrimaryProperty)
			Call Fn_Button_Click("Fn_SISW_BMIDE_NewAuditDefinition", ObjAuditDefinition, "AddAuditDefinitionProperty")

			aPrimaryPropertyValue=split(anPrimaryProperty(iCounter),"~")
			bFlag=Fn_SISW_BMIDE_CreateAuditDefinationProperty(aPrimaryPropertyValue(0),aPrimaryPropertyValue(1),aPrimaryPropertyValue(2),aPrimaryPropertyValue(3),aPrimaryPropertyValue(4),aPrimaryPropertyValue(5),aPrimaryPropertyValue(6),aPrimaryPropertyValue(7),"")
			If bFlag=false Then
				Set ObjAuditDefinition=Nothing
				Exit function
			End If
		Next
	End if
	'Clicking on next button
	Call Fn_Button_Click("Fn_SISW_BMIDE_NewAuditDefinition", ObjAuditDefinition, "Next")	
	'Setting Secondary object defination properties
	If lcase(dicAuditDefinitionInfo("SecondaryObjectAuditDefinitionProperty"))<>"" then
		anSecondaryProperty=Split(dicAuditDefinitionInfo("SecondaryObjectAuditDefinitionProperty"),"$")
		For iCounter=0 to ubound(anSecondaryProperty)
			Call Fn_Button_Click("Fn_SISW_BMIDE_NewAuditDefinition", ObjAuditDefinition, "AddAuditDefinitionProperty")

			aSecondaryPropertyValue=split(anSecondaryProperty(iCounter),"~")
			bFlag=Fn_SISW_BMIDE_CreateAuditDefinationProperty(aSecondaryPropertyValue(0),aSecondaryPropertyValue(1),aSecondaryPropertyValue(2),aSecondaryPropertyValue(3),aSecondaryPropertyValue(4),aSecondaryPropertyValue(5),aSecondaryPropertyValue(6),aSecondaryPropertyValue(7),"")
			If bFlag=false Then
				Set ObjAuditDefinition=Nothing
				Exit function
			End If
		Next
	End if
	If dicAuditDefinitionInfo("Button")<>"" then
		If lcase(dicAuditDefinitionInfo("Button"))="donotclick" Then
		else
			Call Fn_Button_Click("Fn_SISW_BMIDE_NewAuditDefinition", ObjAuditDefinition, dicAuditDefinitionInfo("Button"))
		End If
	else
		'Clicking on next button
		Call Fn_Button_Click("Fn_SISW_BMIDE_NewAuditDefinition", ObjAuditDefinition, "Finish")
	end if
	
	If Err.Number < 0 Then
		Fn_SISW_BMIDE_NewAuditDefinition=False
	Else
		Fn_SISW_BMIDE_NewAuditDefinition=True
	End If
	'releasing object of [New Audit Definition ] dialog
	Set ObjAuditDefinition=Nothing
End Function 
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_SISW_BMIDE_JavaTreeGetItemPath

'Description			 :	Function Used to Return Index Of Tree Node

'Parameters			   :   1.ObjTree: Tree object
'										2.StrNode : Tree node
'
'Return Value		   : 	Tree node path or False
'
'Examples				:   bReturn=Fn_SISW_BMIDE_JavaTreeGetItemPath(JavaWindow("Business Modeler").JavaTree("Extension Tree"),"Test:Audit Manager:Audit Definitions:Discipline~__Create~isTrue")
'								
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												18-Sep-2012								1.0																						Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Public Function Fn_SISW_BMIDE_JavaTreeGetItemPath(ObjTree,StrNode)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_BMIDE_JavaTreeGetItemPath"
   'Variable Declaration
	Dim sItemPath,aStrNode,bFlag,i,iNodeItemsCount
	Dim oCurrentNode,eStrNode, iCount

	aStrNode = Split (StrNode, ":")
	bFlag=False
	Set oCurrentNode = ObjTree.Object
	'To Select first Occurance of Node
		For each eStrNode In aStrNode
			iNodeItemsCount = oCurrentNode.getItemCount()
			iCount=iCount+1
			bFlag=False
			For i = 0 to iNodeItemsCount - 1
				If inStr(1,eStrNode,"~") Then
					eStrNode=replace(eStrNode,"~",":")
				End If
				If Trim(oCurrentNode.getItem(i).getData().toString()) = Trim(eStrNode) Then
					Set oCurrentNode = oCurrentNode.getItem(i)
					If iCount=1 Then
						sItemPath="#" & i
					else
						sItemPath = sItemPath & ":#" & i
					End If
					bFlag=True
					Exit For
				End If
			Next
				If iCount=1 Then
					bFlag=True
				Else
					If bFlag=False Then
						Exit For
					End If
				End If
		Next 
		If bFlag=True Then
			'Function Returns Item Path
			Fn_SISW_BMIDE_JavaTreeGetItemPath = sItemPath
		Else
			Fn_SISW_BMIDE_JavaTreeGetItemPath = False
		End If
		Set oCurrentNode =Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_SISW_BMIDE_AuditExtensionsTableOperations

'Description			 :	Function Used to perform operations on [ Audit Extensions ] Table

'Parameters			   :   1.StrAction: Action name
'										2.StrTab: Tab Name
'										3.StrName: Audit Extension Name
'										4.StrValue: Expected value
'										5.StrColumnName: Column Name
'
'Return Value		   : 	True or False

'Pre-requisite			:	[ Audit Extensions ] Table should be exist

'Examples				:   bReturn=Fn_SISW_BMIDE_AuditExtensionsTableOperations("Add","Audit Definition","Fnd0CICO_auditloghandler~Fnd0USER_get_additional_log_info","","")
'										bReturn=Fn_SISW_BMIDE_AuditExtensionsTableOperations("Select","Audit Definition","Fnd0CICO_auditloghandler","","")
'										bReturn=Fn_SISW_BMIDE_AuditExtensionsTableOperations("Remove","Audit Definition","Fnd0CICO_auditloghandler~Fnd0USER_get_additional_log_info","","")
'								
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												18-Sep-2012								1.0																						Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_SISW_BMIDE_AuditExtensionsTableOperations(StrAction,StrTab,StrName,StrValue,StrColumnName)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_BMIDE_AuditExtensionsTableOperations"
   'Declaring variables
   Dim aName,iRows,iCounter,iCount,bFlag
   Fn_SISW_BMIDE_AuditExtensionsTableOperations=false
   If StrTab<>"" Then
	   Call Fn_BMIDE_InnerTabOperations("Activate",StrTab)
   End If
   Select Case StrAction
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to Remove Audit Extensions from table
	 	Case "Remove"
			'Spliting Audit Extensions name to remove multiple Audit Extensions at same time
			aName=Split(StrName,"~")
			For iCount=0 to ubound(aName)
				iRows=JavaWindow("Business Modeler").JavaTable("AuditExtensions").GetROProperty("rows")
				For iCounter=0 to iRows-1
					bFlag=False
					If aName(iCount)=JavaWindow("Business Modeler").JavaTable("AuditExtensions").GetCellData(iCounter,0) Then
						JavaWindow("Business Modeler").JavaTable("AuditExtensions").SelectCell iCounter,"Name"
						Call Fn_Button_Click("Fn_SISW_BMIDE_AuditExtensionsTableOperations", JavaWindow("Business Modeler"), "Remove")
						bFlag=True
						Exit for
					end if
				Next
				If bFlag=False Then
					Exit for
				End If
			Next
			If bFlag=True Then
				Fn_SISW_BMIDE_AuditExtensionsTableOperations=True
			End If
	 	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to Select Audit Extensions from table
	 	Case "Select"
			'iRows=JavaWindow("Business Modeler").JavaTable("AuditExtensions").GetROProperty("rows")
			iRows=Fn_UI_Object_GetROProperty("Fn_SISW_BMIDE_AuditExtensionsTableOperations",JavaWindow("Business Modeler").JavaTable("AuditExtensions"),"rows")
			For iCounter=0 to iRows-1
				bFlag=False
					If StrName=JavaWindow("Business Modeler").JavaTable("AuditExtensions").GetCellData(iCounter,0) Then
					JavaWindow("Business Modeler").JavaTable("AuditExtensions").SelectCell iCounter,"Name"
					bFlag=True
					Exit for
				end if
			Next
			If bFlag=True Then
				Fn_SISW_BMIDE_AuditExtensionsTableOperations=True
			End If
	 	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to add Audit Extensions
	 	Case "Add"
			'Spliting Audit Extensions name to add multiple Audit Extensions at same time
			aName=Split(StrName,"~")
			For iCounter=0 to ubound(aName)
				'Clicking on [ Add ] button
				If not JavaWindow("Business Modeler").JavaWindow("FindAuditExtension").Exist(1) Then
					Call Fn_Button_Click("Fn_SISW_BMIDE_AuditExtensionsTableOperations", JavaWindow("Business Modeler"), "Add")
				end if
				'Checking existance of [ FindAuditExtension ] dialog
				If not JavaWindow("Business Modeler").JavaWindow("FindAuditExtension").Exist(6) Then
					Exit function
				End If
				'Retriving number of rows exist in [ AuditExtension ] table
				iRows=JavaWindow("Business Modeler").JavaWindow("FindAuditExtension").JavaTable("AuditExtensionTypes").GetROProperty("rows")
				For iCount=0 to iRows-1
					bFlag=false
					'checking Audit Extension with expected Audit Extension
					If aName(iCounter)=JavaWindow("Business Modeler").JavaWindow("FindAuditExtension").JavaTable("AuditExtensionTypes").GetCellData(iCount,0) Then
						'Selecting specific [ AuditExtension ] from table 
						JavaWindow("Business Modeler").JavaWindow("FindAuditExtension").JavaTable("AuditExtensionTypes").ClickCell iCount,0
						wait 1
						'Clicking on [ OK ] button
						Call Fn_Button_Click("Fn_SISW_BMIDE_AuditExtensionsTableOperations", JavaWindow("Business Modeler").JavaWindow("FindAuditExtension"), "OK")
						bFlag=true
						Exit for
					End If
				Next
				If bFlag=false Then
					'Clicking on [ Cancel ] button
					Call Fn_Button_Click("Fn_SISW_BMIDE_AuditExtensionsTableOperations", JavaWindow("Business Modeler").JavaWindow("FindAuditExtension"), "Cancel")
					Exit function
				End If	
			Next
			'Function returns True
			Fn_SISW_BMIDE_AuditExtensionsTableOperations=true			
   End Select
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_SISW_BMIDE_JavaTreeGetItemPathExt

'Description			 :	Function Used to Build path Of Tree Node

'Parameters			   :   1.ObjTree: Tree object
'										2.StrNode : Tree node
'
'Return Value		   : 	Tree node path or False
'
'Examples				:   bReturn=Fn_SISW_BMIDE_JavaTreeGetItemPathExt(JavaWindow("Business Modeler").JavaTree("Extension Tree"),"Changes~Audit Manager~Audit Definitions~TCCalendar:__Modify:isTrue")
'								
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												10-Dec-2012								1.0																						Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Public Function Fn_SISW_BMIDE_JavaTreeGetItemPathExt(ObjTree,StrNode)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_BMIDE_JavaTreeGetItemPathExt"
   'Variable Declaration
	Dim aStrNode,bFlag,i,iCounter,iTemp

	aStrNode=split(StrNode,"~",-1,0)
	StrNode=aStrNode(0)
	For iCounter=0 to ubound(aStrNode)
		If iCounter=0 Then
			StrNode=aStrNode(0)
		Else
			StrNode=StrNode+":"+aStrNode(iCounter)
		End If
		bFlag=False
		For i=0 to ObjTree.GetROProperty("items count")-1
			If ObjTree.GetItem(i)=StrNode Then
				If inStr(1,aStrNode(iCounter),":") Then
					aStrNode(iCounter)="#"+CStr(i-iTemp-1)
				End If
				bFlag=True
				iTemp=i
				Exit for
			End If
		Next
		If bFlag=False Then
			Exit for
		End If
	Next
	If bFlag=False Then
		Fn_SISW_BMIDE_JavaTreeGetItemPathExt=False
	Else
		Fn_SISW_BMIDE_JavaTreeGetItemPathExt=Join(aStrNode,":")
	End If
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_SISW_BMIDE_AuditTypeMappingTableOperations

'Description			 :	Function Used to perform operations on [ Audit Type Mapping ] Table

'Parameters			   :   1.StrAction: Action name
'										2.StrObjectName: Object Name
'										3.StrEventType: Event Type Name
'										4.StrAuditLog: Audit Log
'										5.StrSecondaryAuditType: Secondary Audit Type Name
'										6.bSubscribable: Subscribable option
'										7.bAuditable: Auditable option
'										8.StrDescription: Description
'										9.StrValue: Expected value
'										10.StrColumnName: Expected column name
'
'Return Value		   : 	True or False

'Pre-requisite			:	[ Audit Type Mapping ] Table should be exist

'Examples				:   bReturn=Fn_SISW_BMIDE_AuditTypeMappingTableOperations("Add","L3Test","__Add","Fnd0LicenseChangeAudit","Fnd0SecondaryAudit","on","on","First event to add","","")
'										bReturn=Fn_SISW_BMIDE_AuditTypeMappingTableOperations("Verify","L3Test","__Add","","","","","","Fnd0LicenseChangeAudit","Audit Log")
'										bReturn=Fn_SISW_BMIDE_AuditTypeMappingTableOperations("Select","L3Test","__Add","","","","","","","")
'										bReturn=Fn_SISW_BMIDE_AuditTypeMappingTableOperations("Remove","L3Test","__Add","","","","","","","")
'								
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												20-Sep-2012								1.0																						Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_SISW_BMIDE_AuditTypeMappingTableOperations(StrAction,StrObjectName,StrEventType,StrAuditLog,StrSecondaryAuditType,bSubscribable,bAuditable,StrDescription,StrValue,StrColumnName)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_BMIDE_AuditTypeMappingTableOperations"
 	'Declaring variables
	Dim iRows,iCounter,bFlag,aObjectType,aEventType,iCount

	Fn_SISW_BMIDE_AuditTypeMappingTableOperations=False
	Select Case StrAction
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to add new [ New Event Type Mapping ]
		Case "Add"
			'Checking existance of [ New Event Type Mapping ] dialog
			If not JavaWindow("Business Modeler").JavaWindow("NewEventTypeMapping").Exist(6) Then
				'Clicking on [ Add ] button to open [ New Event Type Mapping ] dialog
				Call Fn_Button_Click("Fn_SISW_BMIDE_AuditTypeMappingTableOperations", JavaWindow("Business Modeler"), "Add")
			End If
			Fn_SISW_BMIDE_AuditTypeMappingTableOperations=Fn_SISW_BMIDE_NewEventTypeMapping("",StrObjectName,StrEventType,StrAuditLog,StrSecondaryAuditType,bSubscribable,bAuditable,StrDescription,"")
			If JavaWindow("Business Modeler").JavaWindow("NewEventTypeMapping").Exist(3) Then
				wait 5
			end if
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to select entry from [ AuditTypeMapping ] table
		Case "Select"
			'Retriving number of rows 
			iRows=JavaWindow("Business Modeler").JavaTable("AuditTypeMapping").GetROProperty("rows")
			For iCounter=0 to iRows-1
				bFlag=False
				'Checking [ Event Type ] with expected Event Type
				If trim(JavaWindow("Business Modeler").JavaTable("AuditTypeMapping").GetCellData(iCounter,"Event Type"))=trim(StrEventType) Then
					'Checking [ Object Type ] with expected Object Type
					If trim(JavaWindow("Business Modeler").JavaTable("AuditTypeMapping").GetCellData(iCounter,"Object Type"))=trim(StrObjectName) Then
						'Selecting specific entry
						JavaWindow("Business Modeler").JavaTable("AuditTypeMapping").SelectCell iCounter,"Object Type"
						bFlag=true
						Exit for
					End If
				End If
			Next
			If bFlag=true Then
				Fn_SISW_BMIDE_AuditTypeMappingTableOperations=true
			End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to Verify entry from [ AuditTypeMapping ] table
		Case "Verify"
			aObjectType=Split(StrObjectName,"~")
			aEventType=Split(StrEventType,"~")

			For iCount=0 to ubound(aObjectType)
				bFlag=False
				'Retriving number of rows 
				iRows=JavaWindow("Business Modeler").JavaTable("AuditTypeMapping").GetROProperty("rows")
				For iCounter=0 to iRows-1
					'Checking [ Event Type ] with expected Event Type
					If trim(JavaWindow("Business Modeler").JavaTable("AuditTypeMapping").GetCellData(iCounter,"Event Type"))=trim(StrEventType) Then
						'Checking [ Object Type ] with expected Object Type
						If trim(JavaWindow("Business Modeler").JavaTable("AuditTypeMapping").GetCellData(iCounter,"Object Type"))=trim(StrObjectName) Then
							if trim(JavaWindow("Business Modeler").JavaTable("AuditTypeMapping").GetCellData(iCounter,StrColumnName))=Trim(StrValue) then
								bFlag=true
								Exit for
							end if
						End If
					End If
				Next
				If bFlag=False Then
					Exit for
				End If
			Next
			If bFlag=true Then
				Fn_SISW_BMIDE_AuditTypeMappingTableOperations=true
			End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to remove Entry from [ AuditTypeMapping ] table
		Case "Remove"
			aObjectType=Split(StrObjectName,"~")
			aEventType=Split(StrEventType,"~")
			If ubound(aEventType)=ubound(aObjectType) Then
				For iCounter=0 to ubound(aEventType)
					bFlag=False
                    bFlag=Fn_SISW_BMIDE_AuditTypeMappingTableOperations("Select",aObjectType(iCounter),aEventType(iCounter),"","","","","","","")
					If bFlag=false Then
						Exit for
					End If
					'Clicking on [ Remove ] button to remove the entry from table
					Call Fn_Button_Click("Fn_SISW_BMIDE_AuditTypeMappingTableOperations", JavaWindow("Business Modeler"), "Remove")
				Next
			End If
			If bFlag=False Then
				Fn_SISW_BMIDE_AuditTypeMappingTableOperations=false
			else
				Fn_SISW_BMIDE_AuditTypeMappingTableOperations=Fn_BMIDE_DeleteObject()
			End If
	End Select
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_SISW_BMIDE_LOVAttachesOfPropertyTreetableOperations

'Description			 :	Function Used to perform operations on LOV Attaches of Property tree table

'Parameters			   :  1.StrAction: Action Name
'									2.StrNode: Node Path
'									3.StrColName: Column name
'									4.StrValue: Expected value
'									5.StrButton: Button name
'
'Return Value		   : 	True or False

'Pre-requisite			:	LOV Attaches Of Property Tree table page Should be appear

'Examples				:   bReturn= Fn_SISW_BMIDE_LOVAttachesOfPropertyTreetableOperations("Select","Part Source","","","")
'Examples				:   bReturn= Fn_SISW_BMIDE_LOVAttachesOfPropertyTreetableOperations("Attach","","","BillCodes","")
'
'History					 :			
'				Developer Name						Date					Rev. No.						Changes Done																				Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'				Sandeep N							9-May-2013				1.0																															Priyanka B
'				Pranav Ingle						27-Jun-2013				1.1																															Sandeep N
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_SISW_BMIDE_LOVAttachesOfPropertyTreetableOperations(StrAction,StrNode,StrColName,StrValue,StrButton)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_BMIDE_LOVAttachesOfPropertyTreetableOperations"
 	'Checking existance of [ Inner tab ]
	If  JavaWindow("Business Modeler").JavaTab("MainInnerTab2").Exist(5) Then
		'Click on [ LOV Attaches ] tab
		Call Fn_UI_JavaTab_Select("Fn_SISW_BMIDE_LOVAttachesOfPropertyTreetableOperations",JavaWindow("Business Modeler"),"MainInnerTab2", "LOV Attaches")
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Sucessfully Selected Inner Tab [ LOV Attaches ]")
	End If
	Fn_SISW_BMIDE_LOVAttachesOfPropertyTreetableOperations=False
	'Checking existance of [ LOVAttachesOfProperty ] tree table
	If not JavaWindow("Business Modeler").JavaTree("LOVAttachesOfProperty").Exist(5) Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ LOVAttachesOfProperty ] Tree table not found")
		Exit function
	End If
	Select Case StrAction
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to select specific node
		Case "Select"
			Fn_SISW_BMIDE_LOVAttachesOfPropertyTreetableOperations=Fn_JavaTree_Select("Fn_SISW_BMIDE_LOVAttachesOfPropertyTreetableOperations", JavaWindow("Business Modeler"), "LOVAttachesOfProperty",StrNode)
			If Fn_SISW_BMIDE_LOVAttachesOfPropertyTreetableOperations=True Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Sucessfully Selected LOV node [ "+StrNode+" ]")
			End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to Attach specific node
		Case "Attach"
			If Fn_UI_ObjectExist("Fn_SISW_BMIDE_LOVAttachesOfPropertyTreetableOperations",JavaWindow("Business Modeler").JavaWindow("SubLOVAttachment"))=False Then
				Call Fn_Button_Click("Fn_SISW_BMIDE_LOVAttachesOfPropertyTreetableOperations",JavaWindow("Business Modeler"),"AddNamingRuleAttaches")
			End If
			'Call Fn_Edit_Box("Fn_SISW_BMIDE_LOVAttachesOfPropertyTreetableOperations", JavaWindow("Business Modeler").JavaWindow("SubLOVAttachment"),"LOV",StrValue)
			Call Fn_Button_Click("Fn_SISW_BMIDE_LOVAttachesOfPropertyTreetableOperations",JavaWindow("Business Modeler").JavaWindow("SubLOVAttachment"),"BrowseLOV")
			Call Fn_UI_EditBox_Type("Fn_SISW_BMIDE_LOVAttachesOfPropertyTreetableOperations",Javawindow("Business Modeler").JavaWindow("FindLOV"),"LOVName",StrValue)
			Call Fn_Button_Click("Fn_SISW_BMIDE_LOVAttachesOfPropertyTreetableOperations",Javawindow("Business Modeler").JavaWindow("FindLOV"),"OK")
			'Here [ StrColName ] parameter used to set Condition
			If StrColName<> "" Then
				Call Fn_UI_Object_SetTOProperty("Fn_SISW_BMIDE_LOVAttachesOfPropertyTreetableOperations",JavaWindow("Business Modeler").JavaWindow("SubLOVAttachment").JavaButton("Browse"),"enabled","1")
				Call Fn_Button_Click("Fn_SISW_BMIDE_LOVAttachesOfPropertyTreetableOperations",JavaWindow("Business Modeler").JavaWindow("SubLOVAttachment"),"Browse")
				Call Fn_UI_EditBox_Type("Fn_SISW_BMIDE_LOVAttachesOfPropertyTreetableOperations",JavaWindow("Business Modeler").JavaWindow("FindAttachmentCondition"),"Condition",StrColName)
				Call Fn_Button_Click("Fn_SISW_BMIDE_LOVAttachesOfPropertyTreetableOperations",JavaWindow("Business Modeler").JavaWindow("FindAttachmentCondition"),"OK")
			End If
			Call Fn_Button_Click("Fn_SISW_BMIDE_LOVAttachesOfPropertyTreetableOperations",JavaWindow("Business Modeler").JavaWindow("SubLOVAttachment"),"Finish")

			Fn_SISW_BMIDE_LOVAttachesOfPropertyTreetableOperations=True
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Sucessfully Ataches LOV [ "+StrNode+" ]")

	End Select
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_SISW_BMIDE_HelpContentsTreeOperations
'
'Description			 :	Function Used to perform operations on Help Contents Tree
'
'Parameters			   :  1.StrAction: Action Name
'									2.StrNode: Node Path
'									3.StrPopupMenu: Popup menu 
'
'Return Value		   : 	True or False

'Pre-requisite			:	Should be login in BMIDE

'Examples				:   bReturn= Fn_SISW_BMIDE_HelpContentsTreeOperations("Expand","Business Modeler IDE Guide:Creating business rules:Introduction to business rules","")
'									bReturn= Fn_SISW_BMIDE_HelpContentsTreeOperations("Select","Business Modeler IDE Guide:Creating business rules:Alternate ID rules:Alternate ID rules characteristics","")
'									bReturn= Fn_SISW_BMIDE_HelpContentsTreeOperations("Exist","Business Modeler IDE Guide:Creating business rules:Alternate ID rules:Alternate ID rules characteristics","")
'									bReturn= Fn_SISW_BMIDE_HelpContentsTreeOperations("CloseHelp","Business Modeler IDE Guide:Creating business rules:Alternate ID rules:Alternate ID rules characteristics","")
'
'History					 :			
'				Developer Name						Date					Rev. No.						Changes Done																				Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'				Sandeep N							27-May-2013				1.0																																		Pallavi J
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_SISW_BMIDE_HelpContentsTreeOperations(StrAction,StrNode,StrPopupMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_BMIDE_HelpContentsTreeOperations"
	'Declaring variables
	Dim objHelpTree
	Dim intNodeCount,intCount,sTreeItem
 	'Checking Existance of [ Contents ] link
	If not JavaWindow("Business Modeler").JavaObject("HelpContentsHyperlink").Exist(5) Then
		'Press F1 to open Help contents
		JavaWindow("Business Modeler").PressKey micF1
		wait 2
		JavaWindow("Business Modeler").JavaObject("HelpContentsHyperlink").Click 1,1,"LEFT"
		wait 2
	End If
	If not JavaWindow("Business Modeler").JavaTree("HelpContentsTree").Exist(5) Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ HelpContentsTree ] not found in application")
		Exit function
	End If
	'Creating object of Help Contents tree
	Set objHelpTree=JavaWindow("Business Modeler").JavaTree("HelpContentsTree")
	Select Case StrAction
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Select"
			Fn_SISW_BMIDE_HelpContentsTreeOperations= Fn_JavaTree_Select("Fn_SISW_BMIDE_HelpContentsTreeOperations", JavaWindow("Business Modeler"), "HelpContentsTree",StrNode)
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Expand"
			Fn_SISW_BMIDE_HelpContentsTreeOperations=Fn_UI_JavaTree_Expand("Fn_SISW_BMIDE_HelpContentsTreeOperations",JavaWindow("Business Modeler"),"HelpContentsTree",StrNode)
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Exist"
			intNodeCount = objHelpTree.GetROProperty ("items count") 
			For intCount = 0 to intNodeCount - 1
				sTreeItem = objHelpTree.GetItem(intCount)
				If Trim(lcase(sTreeItem)) = Trim(Lcase(StrNode)) Then
					Fn_SISW_BMIDE_HelpContentsTreeOperations = TRUE	
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[" + StrNode + "] of JavaTree exist")
					Exit For
				End If
			Next
			If Cstr(intCount) = intNodeCount Then
				Fn_SISW_BMIDE_HelpContentsTreeOperations = FALSE						
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[" + StrNode + "] of JavaTree does not exist")
				Exit Function
			End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "CloseHelp"
			Fn_SISW_BMIDE_HelpContentsTreeOperations=Fn_BMIDE_TabOperations("Main","Close","Help")
	End Select
	'Releasing object of Help Contents tree
	Set objHelpTree=Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_SISW_BMIDE_ConfigureBusinessObjectsToDisplayOperations
'
'Description			 :	Function Used to perform operations on Configure Business Objects to Display
'
'Parameters			   :  1.StrAction: Action Name
'									   2.bCustomizeBOToShowOption: Customize Business Objects to Option
'									   3.StrBOName: Business Object names
'									   4.StrBOSearchValue: Business Object name to filter
'									   5.StrButton: Button Name
'
'Return Value		   : 	True or False

'Pre-requisite			:	Should be login in BMIDE Standar perspective

'Examples				:   bReturn= Fn_SISW_BMIDE_ConfigureBusinessObjectsToDisplayOperations("Add","","AbsOccData~AbsOccFlags","","")
'									bReturn= Fn_SISW_BMIDE_ConfigureBusinessObjectsToDisplayOperations("Remove","","AbsOccData~AbsOccFlags","","")
'									bReturn= Fn_SISW_BMIDE_ConfigureBusinessObjectsToDisplayOperations("AddAll","","","","")
'									bReturn= Fn_SISW_BMIDE_ConfigureBusinessObjectsToDisplayOperations("RemoveAll","","","","OK")
'
'History					 :			
'				Developer Name						Date					Rev. No.						Changes Done																				Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'				Sandeep N							25-June-2013				1.0																																			Avinash J	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_SISW_BMIDE_ConfigureBusinessObjectsToDisplayOperations(StrAction,bCustomizeBOToShowOption,StrBOName,StrBOSearchValue,StrButton)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_BMIDE_ConfigureBusinessObjectsToDisplayOperations"
 	'Declaring variables
	Dim ObjConfigureBODialog,ObjBOList
	Dim aBOName,iCount,iCounter,bFlag

	Fn_SISW_BMIDE_ConfigureBusinessObjectsToDisplayOperations=False
	'Checking existance of [ Configure Business Objects To Display ] dialog
	If JavaWindow("Business Modeler").JavaWindow("ConfigureBusinessObjectsToDisplay").Exist(10) Then
		'creating object of [ Configure Business Objects To Display ] dialog
		Set ObjConfigureBODialog=JavaWindow("Business Modeler").JavaWindow("ConfigureBusinessObjectsToDisplay")
		wait 2
	Else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail : [ Configure Business Objects To Display ] dialog not exist")
		Exit function
	End If
	'Set [ Customize Business Objects to Show ] option
	Call Fn_CheckBox_Set("Fn_SISW_BMIDE_ConfigureBusinessObjectsToDisplayOperations", ObjConfigureBODialog, "CustomizeBusinessObjectsToShow", "ON")
	Select Case StrAction		
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'case to Add Business Objects from Hidden Business Objects list to Displayed Business Objects list
		Case "Add"
			Set ObjBOList=ObjConfigureBODialog.JavaTable("HiddenBusinessObjects")
			aBOName=Split(StrBOName,"~")
			For iCount=0 to Ubound(aBOName)
				bFlag=False
				For iCounter=0 to Cint(ObjBOList.GetROProperty("rows"))-1
					If trim(aBOName(iCount))=trim(ObjBOList.GetCellData(iCounter,0)) Then
						ObjBOList.Type aBOName(iCount)
						wait 2
						Call Fn_Button_Click("Fn_SISW_BMIDE_ConfigureBusinessObjectsToDisplayOperations", ObjConfigureBODialog, "Add")
						wait 2
						bFlag=True
						Exit For
					End If
				Next
				If bFlag=False Then
					Exit for
				End If
			Next
			If bFlag=True Then
				Fn_SISW_BMIDE_ConfigureBusinessObjectsToDisplayOperations=True
			End If
			Set ObjBOList=Nothing
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'case to Add All Business Objects from Hidden Business Objects list to Displayed Business Objects list
		Case "AddAll"
			Fn_SISW_BMIDE_ConfigureBusinessObjectsToDisplayOperations=Fn_Button_Click("Fn_SISW_BMIDE_ConfigureBusinessObjectsToDisplayOperations", ObjConfigureBODialog,"AddAll")
			wait 2
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'case to Remove Business Objects from Displayed Business Objects list
		Case "Remove"
			Set ObjBOList=ObjConfigureBODialog.JavaTable("DisplayedBusinessObjects")
			aBOName=Split(StrBOName,"~")
			For iCount=0 to Ubound(aBOName)
				bFlag=False
				For iCounter=0 to Cint(ObjBOList.GetROProperty("rows"))-1
					If trim(aBOName(iCount))=trim(ObjBOList.GetCellData(iCounter,0)) Then
						ObjBOList.Type aBOName(iCount)
						wait 2
						Call Fn_Button_Click("Fn_SISW_BMIDE_ConfigureBusinessObjectsToDisplayOperations", ObjConfigureBODialog, "Remove")
						wait 2
						bFlag=True
						Exit For
					End If
				Next
				If bFlag=False Then
					Exit for
				End If
			Next
			If bFlag=True Then
				Fn_SISW_BMIDE_ConfigureBusinessObjectsToDisplayOperations=True
			End If
			Set ObjBOList=Nothing
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'case to Remove all Business Objects from Displayed Business Objects list
		Case "RemoveAll"
        	Fn_SISW_BMIDE_ConfigureBusinessObjectsToDisplayOperations=Fn_Button_Click("Fn_SISW_BMIDE_ConfigureBusinessObjectsToDisplayOperations", ObjConfigureBODialog,"RemoveAll")
			wait 2	
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'case to verify displayed business object in list
		Case "VerifyDisplayedObject"
			Set ObjBOList=ObjConfigureBODialog.JavaTable("DisplayedBusinessObjects")
			aBOName=Split(StrBOName,"~")
			For iCount=0 to Ubound(aBOName)
				bFlag=False
				For iCounter=0 to Cint(ObjBOList.GetROProperty("rows"))-1
					If trim(aBOName(iCount))=trim(ObjBOList.GetCellData(iCounter,0)) Then
						bFlag=True
						Exit For
					End If
				Next
				If bFlag=False Then
					Exit for
				End If
			Next
			If bFlag=True Then
				Fn_SISW_BMIDE_ConfigureBusinessObjectsToDisplayOperations=True
			End If
			Set ObjBOList=Nothing
			
	End Select
	If bCustomizeBOToShowOption<>"" Then
		Call Fn_CheckBox_Set("Fn_SISW_BMIDE_ConfigureBusinessObjectsToDisplayOperations", ObjConfigureBODialog, "CustomizeBusinessObjectsToShow", bCustomizeBOToShowOption)
	End If
	If StrButton<>"" Then
		Call Fn_Button_Click("Fn_SISW_BMIDE_ConfigureBusinessObjectsToDisplayOperations", ObjConfigureBODialog,StrButton)
	End If
	Set ObjConfigureBODialog=Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_SISW_BMIDE_FilterConfigurationOperations
'
'Description			 :	Function Used to perform operations on Elements tree of Filter Configuration
'
'Parameters			   :  1.StrAction: Action Name
'									   2.StrNode: Node name
'									   3.StrState: Item state
'									   4.StrButton: Button Name
'
'Return Value		   : 	True or False

'Pre-requisite			:	Filter Configuration dialog Should be appear

'Examples				:   bReturn=Fn_SISW_BMIDE_FilterConfigurationOperations("SetState","Extensions:Audit Manager","uncheck","")
'										bReturn=Fn_SISW_BMIDE_FilterConfigurationOperations("Expand","Extensions:Constants","","")
'										bReturn=Fn_SISW_BMIDE_FilterConfigurationOperations("Exist","Extensions:Constants:Global Constants1","","")
'										bReturn=Fn_SISW_BMIDE_FilterConfigurationOperations("Select","Extensions:Constants:Global Constants","","")
'										bReturn=Fn_SISW_BMIDE_FilterConfigurationOperations("SetState","Extensions:Audit Manager","check","OK")
'
'History					 :			
'				Developer Name						Date					Rev. No.						Changes Done																				Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'				Sandeep N							26-June-2013				1.0																																			Avinash J	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_SISW_BMIDE_FilterConfigurationOperations(StrAction,StrNode,StrState,StrButton)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_BMIDE_FilterConfigurationOperations"
 	'Declaring variables
	Dim ObjFilterConfigurationDialog
	Dim iCounter

	Fn_SISW_BMIDE_FilterConfigurationOperations=false
	'Checking existance of [ Filter Configuration ] dialog
	If JavaWindow("Business Modeler").JavaWindow("FilterConfiguration").Exist(10) Then
		'Creating object of [ Filter Configuration ] dialog
		Set ObjFilterConfigurationDialog=JavaWindow("Business Modeler").JavaWindow("FilterConfiguration")
	Else
	    Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail : [ Filter Configuration ] dialog not exist")
		Exit function
	End If
	Select Case StrAction
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'case to set state of item
		Case "SetState"
			If Lcase(StrState)="check" Then
				ObjFilterConfigurationDialog.JavaTree("ElementsTree").SetItemState StrNode,micChecked
				wait 1
			Elseif Lcase(StrState)="uncheck" then
				ObjFilterConfigurationDialog.JavaTree("ElementsTree").SetItemState StrNode,micUnchecked
				wait 1
			End If
			If Err.Number < 0 Then
				Fn_SISW_BMIDE_FilterConfigurationOperations=False
			Else
				Fn_SISW_BMIDE_FilterConfigurationOperations=True
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'case to expand node
		Case "Expand"
			ObjFilterConfigurationDialog.JavaTree("ElementsTree").Expand StrNode
			If Err.Number < 0 Then
				Fn_SISW_BMIDE_FilterConfigurationOperations=False
			Else
				Fn_SISW_BMIDE_FilterConfigurationOperations=True
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'case to select node
		Case "Select"
			ObjFilterConfigurationDialog.JavaTree("ElementsTree").Select StrNode
			If Err.Number < 0 Then
				Fn_SISW_BMIDE_FilterConfigurationOperations=False
			Else
				Fn_SISW_BMIDE_FilterConfigurationOperations=True
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'case to check existance of node
		Case "Exist"
			For iCounter=0 to Cint(ObjFilterConfigurationDialog.JavaTree("ElementsTree").GetROProperty("items count"))-1
				If trim(StrNode)=trim(ObjFilterConfigurationDialog.JavaTree("ElementsTree").GetItem(iCounter)) Then
					Fn_SISW_BMIDE_FilterConfigurationOperations=True
					Exit for
				End If
			Next
	End Select
	If StrButton<>"" Then
        Call Fn_Button_Click("Fn_SISW_BMIDE_FilterConfigurationOperations", ObjFilterConfigurationDialog,StrButton)
	End If
	'releasing object of [ Filter Configuration ] dialog
	Set ObjFilterConfigurationDialog=Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_SISW_BMIDE_CreateBatchLOV
'
'Description			 :	Function Used to create Batch LOV
'
'Parameters			   :  1.StrName: Batch LOV Name
'									   2.StrDescription: Batch LOV Description
'									   3.StrType: LOV Type
'									   4.StrUsage: Usage option
'
'Return Value		   : 	True or False

'Pre-requisite			:	New Batch LOV creation dialog Should be appear

'Examples				:   bReturn=Fn_SISW_BMIDE_CreateBatchLOV("BatchLOV1","Batch LOV desc","ListOfValuesInteger","Suggestive")
'
'History					 :			
'				Developer Name						Date					Rev. No.						Changes Done																				Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'				Sandeep N							26-June-2013				1.0																																			Avinash J	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_SISW_BMIDE_CreateBatchLOV(StrName,StrDescription,StrType,StrUsage)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_BMIDE_CreateBatchLOV"
 	'Declaring variables
	Dim ObjBatchLOVDialog,WshShell
	Dim StrPrifix
	Fn_SISW_BMIDE_CreateBatchLOV=False
	If not JavaWindow("Business Modeler").JavaWindow("NewBatchLOV").Exist(10) Then
        Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail : [ New Batch LOV ] creation dialog not exist")
		Exit function
	End If
	'Creating object of [ New Batch LOV ] dialog
	Set ObjBatchLOVDialog=JavaWindow("Business Modeler").JavaWindow("NewBatchLOV")
	'Retrive current prefix
    StrPrifix= Fn_Edit_Box_GetValue("Fn_SISW_BMIDE_CreateBatchLOV",ObjBatchLOVDialog,"Name")
	StrName=StrPrifix+StrName
	'setting Batch LOV name
    Call Fn_Edit_Box("Fn_SISW_BMIDE_CreateBatchLOV",ObjBatchLOVDialog,"Name",StrName)
	'Setting Batch LOV description
	If StrDescription<>"" Then
		Call Fn_Edit_Box("Fn_SISW_BMIDE_CreateBatchLOV",ObjBatchLOVDialog,"Description",StrDescription)
	End If
	'Setting Batch LOV Type
	Call Fn_Edit_Box("Fn_SISW_BMIDE_CreateBatchLOV",ObjBatchLOVDialog,"Type",StrType)
	wait 1
	Set WshShell = CreateObject("WScript.Shell")
	WshShell.SendKeys "{ESC}"
	wait 1
	Set WshShell = nothing
	'selecting usage option
	If StrUsage<>"" Then
		ObjBatchLOVDialog.JavaRadioButton("Usage").SetTOProperty "attached text",StrUsage
		Call Fn_UI_JavaRadioButton_SetON("Fn_SISW_BMIDE_CreateBatchLOV",ObjBatchLOVDialog, "Usage")
	End If
	'Clicking on [ Finish ] button
    Fn_SISW_BMIDE_CreateBatchLOV=Fn_Button_Click("Fn_SISW_BMIDE_CreateBatchLOV", ObjBatchLOVDialog, "Finish")
	wait 2
	'Releasing object of [ New Batch LOV ] dialog
	Set ObjBatchLOVDialog=Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_SISW_BMIDE_BuildQueryClauseOperations
'
'Description			 :	Function Used to perform operations on Build Query Clause
'
'Parameters			   :  1.StrAction: Action Name
'									   2.dicQueryInformation: Build Query Clause information
'									   3.StrButton: Button name
'
'Return Value		   : 	True or False

'Pre-requisite			:	Build Query Clause dialog Should be appear

'Examples				:   Dim dicQueryInformation
'										Set dicQueryInformation = CreateObject("Scripting.Dictionary")
'										dicQueryInformation("Filter")="Attributes and References"
'										dicQueryInformation("ReferenceType")="PSBOMView"
'										dicQueryInformation("PropertyName")="cd_tags"
'										bReturn=Fn_SISW_BMIDE_BuildQueryClauseOperations("BuildQuery",dicQueryInformation,"Next")
'
'										dicQueryInformation("ReferenceType")="POM_object"
'										dicQueryInformation("PropertyName")="timestamp"
'										bReturn=Fn_SISW_BMIDE_BuildQueryClauseOperations("BuildQuery",dicQueryInformation,"Finish")
'
'History					 :			
'				Developer Name						Date					Rev. No.						Changes Done																				Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'				Sandeep N							26-June-2013				1.0																																			Avinash J	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_SISW_BMIDE_BuildQueryClauseOperations(StrAction,dicQueryInformation,StrButton)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_BMIDE_BuildQueryClauseOperations"
 	'Declaring variables
	Dim ObjBuildQueryClauseDialog
	Dim bFlag,iCounter
	Fn_SISW_BMIDE_BuildQueryClauseOperations=False
	'Checking existance of [ Build Query Clause ] dialog
	If Not JavaWindow("Business Modeler").JavaWindow("BuildQueryClause").Exist(10) Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail : [ Build Query Clause ] dialog not exist")
		Exit function
	End If
	'creating object of [ Build Query Clause ] dialog
	Set ObjBuildQueryClauseDialog=JavaWindow("Business Modeler").JavaWindow("BuildQueryClause")
	
	Select Case StrAction
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to build query
		Case "BuildQuery"
			'setting filter
			If dicQueryInformation("Filter")<>"" Then
				ObjBuildQueryClauseDialog.JavaRadioButton("Filter").SetTOProperty "attached text",dicQueryInformation("Filter")
				Call Fn_UI_JavaRadioButton_SetON("Fn_SISW_BMIDE_BuildQueryClauseOperations",ObjBuildQueryClauseDialog, "Filter")
			End If
			'Setting Reference By Type
			If dicQueryInformation("ReferenceByType")<>"" Then
				Call Fn_Edit_Box("Fn_SISW_BMIDE_BuildQueryClauseOperations",ObjBuildQueryClauseDialog,"ReferencedByType",dicQueryInformation("ReferenceByType"))
				wait 2
			End if
			'Setting Reference Type
			If dicQueryInformation("ReferenceType")<>"" Then
				Call Fn_Edit_Box("Fn_SISW_BMIDE_BuildQueryClauseOperations",ObjBuildQueryClauseDialog,"ReferencedType",dicQueryInformation("ReferenceType"))
				wait 2
			End If
			'Selecting property
			If dicQueryInformation("PropertyName")<>"" Then
				bFlag=False
				For iCounter=0 to Cint(ObjBuildQueryClauseDialog.JavaTable("PropertyTable").GetROProperty("rows"))-1
					If trim(ObjBuildQueryClauseDialog.JavaTable("PropertyTable").GetCellData(iCounter,"Property Name"))=trim(dicQueryInformation("PropertyName")) Then
						ObjBuildQueryClauseDialog.JavaTable("PropertyTable").SelectCell iCounter,"Property Name"
						wait 2
						bFlag=True
						Exit for
					End if
				Next
				If bFlag=False Then
					Exit function
				End If
			End if
			'Clicking on specific button
			If StrButton<>"" Then
				Call Fn_Button_Click("Fn_SISW_BMIDE_BuildQueryClauseOperations", ObjBuildQueryClauseDialog, StrButton)
			End If
			Fn_SISW_BMIDE_BuildQueryClauseOperations=True
	End Select 
	'Releasing object of [ Build Query Clause ] dialog
	Set ObjBuildQueryClauseDialog=Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_SISW_BMIDE_NewDynamicLOVQueryCriteriaTableOperations
'
'Description			 :	Function Used to perform operations on Dynamic LOV Query criteria table
'
'Parameters			   :  1.StrAction: Action Name
'									   2.StrAttribute: Attribute name
'									   3.StrColumn: Column name
'									   4.StrValue: Value
'									   5.StrButton: Button Name
'
'Return Value		   : 	True or False

'Pre-requisite			:	Query Criteria table Should be appear

'Examples				:   bReturn=Fn_SISW_BMIDE_NewDynamicLOVQueryCriteriaTableOperations("SetData","AbsOccNote.archelemid","","AND","")
'										bReturn=Fn_SISW_BMIDE_NewDynamicLOVQueryCriteriaTableOperations("SetData","AbsOccNote.archelemid","Operator","!=","")
'										bReturn=Fn_SISW_BMIDE_NewDynamicLOVQueryCriteriaTableOperations("SetData","AbsOccNote.archelemid","Value","10","")
'
'History					 :			
'				Developer Name						Date					Rev. No.						Changes Done																				Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'				Sandeep N							26-June-2013				1.0																																			Avinash J	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_SISW_BMIDE_NewDynamicLOVQueryCriteriaTableOperations(StrAction,StrAttribute,StrColumn,StrValue,StrButton)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_BMIDE_NewDynamicLOVQueryCriteriaTableOperations"
 	'Declaring variables
	Dim ObjDynamicLOVDialog
	Dim bFlag,iCounter,arrDate

	If not JavaWindow("Business Modeler").JavaWindow("NewDynamicLOV").Exist(10) Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail : [ New Dynamic LOV ] dialog not exist")
		Exit function
	End If
	'Creating object of [ New Dynamic LOV ] dialog
   Set ObjDynamicLOVDialog=JavaWindow("Business Modeler").JavaWindow("NewDynamicLOV")
   Select Case StrAction
	 	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	 	'Case to set data in Query Criteria table
	 	Case "SetData"
				bFlag=False
				For iCounter=0 to Cint(ObjDynamicLOVDialog.JavaTable("QueryCriteria").GetROProperty("rows"))-1
					If trim(ObjDynamicLOVDialog.JavaTable("QueryCriteria").GetCellData(iCounter,"Attribute"))=trim(StrAttribute) Then
						If StrColumn="" Then
							ObjDynamicLOVDialog.JavaTable("QueryCriteria").SelectCell iCounter,0
						Else
							ObjDynamicLOVDialog.JavaTable("QueryCriteria").SelectCell iCounter,StrColumn
						End If
						wait 1
						bFlag=True
						Exit for
					End If
				Next
				If bFlag=True Then
					If StrColumn="" or StrColumn="Operator" Then
						If ObjDynamicLOVDialog.JavaList("QueryCriteriaList").Exist(5) Then
							ObjDynamicLOVDialog.JavaList("QueryCriteriaList").Select StrValue
							Fn_SISW_BMIDE_NewDynamicLOVQueryCriteriaTableOperations=True
						End If
					Elseif StrColumn="Value" then
						If ObjDynamicLOVDialog.JavaButton("SetQueryCriteriaDate").Exist(2) Then
							Call Fn_Button_Click("Fn_SISW_BMIDE_NewDynamicLOVQueryCriteriaTableOperations", ObjDynamicLOVDialog, "SetQueryCriteriaDate")					
							If ObjDynamicLOVDialog.JavaWindow("ChooseDate").Exist(5) Then
								
								'StrValue : 2013-June-4-3:12:12 , 2012-May-30-21:05:00
								arrDate=Split(StrValue,"-")
								'Setting year . Format := 2013,2012
								Call Fn_Edit_Box("Fn_SISW_BMIDE_NewDynamicLOVQueryCriteriaTableOperations",ObjDynamicLOVDialog.JavaWindow("ChooseDate"),"Year", arrDate(0))
								'Setting month . Format :=June,May
								Call Fn_List_Select("Fn_SISW_BMIDE_NewDynamicLOVQueryCriteriaTableOperations", ObjDynamicLOVDialog.JavaWindow("ChooseDate"),"Month" ,arrDate(1))
								'Setting Day . Format :=4,15,30
								Call Fn_List_Select("Fn_SISW_BMIDE_NewDynamicLOVQueryCriteriaTableOperations", ObjDynamicLOVDialog.JavaWindow("ChooseDate"),"Month" ,arrDate(1))
								ObjDynamicLOVDialog.JavaWindow("ChooseDate").JavaStaticText("Day").SetTOProperty "label",arrDate(2)
								ObjDynamicLOVDialog.JavaWindow("ChooseDate").JavaStaticText("Day").Click 1,1,"LEFT"
								'Setting time
								If ubound(arrDate)=3 Then
									ObjDynamicLOVDialog.JavaWindow("ChooseDate").JavaCalendar("Time").SetTime arrDate(3)
								End If
								wait 1
								Fn_SISW_BMIDE_NewDynamicLOVQueryCriteriaTableOperations=True
							End If
						ElseIf ObjDynamicLOVDialog.JavaEdit("QueryCriteriaEdit").Exist(3) Then
							ObjDynamicLOVDialog.JavaEdit("QueryCriteriaEdit").Set StrValue
							Fn_SISW_BMIDE_NewDynamicLOVQueryCriteriaTableOperations=True
						End If
					End If
					ObjDynamicLOVDialog.JavaTable("QueryCriteria").SelectCell 0,0
					wait 2
				End If
   End Select
   If StrButton<>"" Then
		Call Fn_Button_Click("Fn_SISW_BMIDE_NewDynamicLOVQueryCriteriaTableOperations", ObjDynamicLOVDialog, StrButton)
   End If
   Set ObjDynamicLOVDialog=Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_SISW_BMIDE_NewDynamicLOVOperations
'
'Description			 :	Function Used to perform operations on New Dynamic LOV creation
'
'Parameters			   :  1.StrAction: Action Name
'									   2.dicDynamicLOVInfo: Dynamic LOV Information
'									   3.StrButton: Button name
'
'Return Value		   : 	True or False

'Pre-requisite			:	New Dynamic LOV dialog Should be appear

'Examples				:   Dim dicDynamicLOVInfo
'										Set dicDynamicLOVInfo = CreateObject("Scripting.Dictionary")
'										dicDynamicLOVInfo("Name")="DLOV1"
'										dicDynamicLOVInfo("Description")="DLOV1 Desc"
'										dicDynamicLOVInfo("DataType")="String"
'										dicDynamicLOVInfo("Usage")="Exhaustive"
'										dicDynamicLOVInfo("Type")="Fnd0ListOfValuesDynamic"
'										dicDynamicLOVInfo("QueryType")="AbsOccData"
'										dicDynamicLOVInfo("LOVValueAttribute")="seqno"
'										dicDynamicLOVInfo("LOVDescriptionAttribute")="occname"
'										dicDynamicLOVInfo("FilterAttributes")="lsd~notetext"
'										bReturn=Fn_SISW_BMIDE_NewDynamicLOVOperations("Create",dicDynamicLOVInfo,"")
'										
'										bReturn=Fn_SISW_BMIDE_NewDynamicLOVOperations("AddQueryCriteria","","")
'
'History					 :			
'				Developer Name						Date					Rev. No.						Changes Done																				Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'				Sandeep N							27-June-2013				1.0																																			Avinash J	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_SISW_BMIDE_NewDynamicLOVOperations(StrAction,dicDynamicLOVInfo,StrButton)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_BMIDE_NewDynamicLOVOperations"
 	'Declaring variables
	Dim ObjDynamicLOVDialog
	Dim StrPrifix,aFilterAttributes,bFlag,iCounter,iCount

	Fn_SISW_BMIDE_NewDynamicLOVOperations=False
   	If not JavaWindow("Business Modeler").JavaWindow("NewDynamicLOV").Exist(10) Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail : [ New Dynamic LOV ] dialog not exist")
		Exit function
	End If
	'Creating object of [ New Dynamic LOV ] dialog
   Set ObjDynamicLOVDialog=JavaWindow("Business Modeler").JavaWindow("NewDynamicLOV")
   Select Case StrAction
	 	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	 	Case "Create"
				'Retrive current prefix
				StrPrifix= Fn_Edit_Box_GetValue("Fn_SISW_BMIDE_NewDynamicLOVOperations",ObjDynamicLOVDialog,"Name")
				dicDynamicLOVInfo("Name")=StrPrifix+dicDynamicLOVInfo("Name")
				'setting Batch LOV name
				Call Fn_Edit_Box("Fn_SISW_BMIDE_NewDynamicLOVOperations",ObjDynamicLOVDialog,"Name",dicDynamicLOVInfo("Name"))
				'Set Description
				If dicDynamicLOVInfo("Description")<>"" Then
					Call Fn_Edit_Box("Fn_SISW_BMIDE_NewDynamicLOVOperations",ObjDynamicLOVDialog,"Description",dicDynamicLOVInfo("Description"))
				End If
				'Select Data Type
				If dicDynamicLOVInfo("DataType")<>"" Then
					Call Fn_List_Select("Fn_SISW_BMIDE_NewDynamicLOVOperations", ObjDynamicLOVDialog, "DataType",dicDynamicLOVInfo("DataType"))
				End if
				'Set Usage option
				If dicDynamicLOVInfo("Usage")<>"" Then
					ObjDynamicLOVDialog.JavaRadioButton("Usage").SetTOProperty "attached text",dicDynamicLOVInfo("Usage")
					Call Fn_UI_JavaRadioButton_SetON("Fn_SISW_BMIDE_NewDynamicLOVOperations",ObjDynamicLOVDialog, "Usage")
				End If
				'Set Type
				If dicDynamicLOVInfo("Type")<>"" Then
					Call Fn_Edit_Box("Fn_SISW_BMIDE_NewDynamicLOVOperations",ObjDynamicLOVDialog,"Type",dicDynamicLOVInfo("Type"))
				End If
				'Set Query Type
				If dicDynamicLOVInfo("QueryType")<>"" Then
					Call Fn_Edit_Box("Fn_SISW_BMIDE_NewDynamicLOVOperations",ObjDynamicLOVDialog,"QueryType",dicDynamicLOVInfo("QueryType"))
				End If
				'Set LOV Value Attribute
				If dicDynamicLOVInfo("LOVValueAttribute")<>"" Then
					Call Fn_Edit_Box("Fn_SISW_BMIDE_NewDynamicLOVOperations",ObjDynamicLOVDialog,"LOVValueAttribute",dicDynamicLOVInfo("LOVValueAttribute"))
				End If
				'Set LOV Description Attribute
				If dicDynamicLOVInfo("LOVDescriptionAttribute")<>"" Then
					Call Fn_Edit_Box("Fn_SISW_BMIDE_NewDynamicLOVOperations",ObjDynamicLOVDialog,"LOVDescriptionAttribute",dicDynamicLOVInfo("LOVDescriptionAttribute"))
				End If
				'Set Filter Attributes
				If dicDynamicLOVInfo("FilterAttributes")<>"" Then
					aFilterAttributes=Split(dicDynamicLOVInfo("FilterAttributes"),"~")
					For iCounter=0 to Ubound(aFilterAttributes)
						bFlag=False
						Call Fn_Button_Click("Fn_SISW_BMIDE_NewDynamicLOVOperations", ObjDynamicLOVDialog, "AddFilterAttributes")
						If JavaWindow("Business Modeler").JavaWindow("PropertySelection").Exist(10) Then
							'For iCount=0 to cint(JavaWindow("Business Modeler").JavaWindow("PropertySelection").JavaTable("PropertyTable").GetROProperty("rows"))-1
								For iCount=0 to cint(Fn_UI_Object_GetROProperty("", JavaWindow("Business Modeler").JavaWindow("PropertySelection").JavaTable("PropertyTable"),"rows"))-1							
								If trim(aFilterAttributes(iCounter))=trim(JavaWindow("Business Modeler").JavaWindow("PropertySelection").JavaTable("PropertyTable").GetCellData(iCount,"Property Name")) Then
									JavaWindow("Business Modeler").JavaWindow("PropertySelection").JavaTable("PropertyTable").SelectCell iCount,"Property Name"
									wait 1
									Call Fn_Button_Click("Fn_SISW_BMIDE_NewDynamicLOVOperations", JavaWindow("Business Modeler").JavaWindow("PropertySelection"), "OK")
									wait 1
									bFlag=True
									Exit For
								End If
							Next
						Else
							Set ObjDynamicLOVDialog=Nothing
							Exit function
						End If
						If bFlag=False Then
							Set ObjDynamicLOVDialog=Nothing
							Exit function
						End If
					Next
				End If
				If StrButton<>"" Then
					Call Fn_Button_Click("Fn_SISW_BMIDE_NewDynamicLOVOperations", ObjDynamicLOVDialog,StrButton)
				End If
				Fn_SISW_BMIDE_NewDynamicLOVOperations=True
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "AddQueryCriteria"
			Fn_SISW_BMIDE_NewDynamicLOVOperations=Fn_Button_Click("Fn_SISW_BMIDE_NewDynamicLOVOperations", ObjDynamicLOVDialog,"AddQueryCriteria")
			wait 2
   End Select
   'Releasing object of [ Dynamic LOV Dialog ]
   Set ObjDynamicLOVDialog=Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_SISW_BMIDE_CloseAllDialogsOperations
'
'Description			 :	Function Used to reset BMIDE client to basic state
'
'Parameters			   :  1.StrAction: Action Name
'
'Return Value		   : 	True or False

'Pre-requisite			:	Should be Log in BMIDEclient

'Examples				:   bReturn=Fn_SISW_BMIDE_CloseAllDialogsOperations("Standard")
'
'History					 :			
'				Developer Name						Date					Rev. No.						Changes Done																				Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'				Sandeep N							27-June-2013				1.0																																			Avinash J	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_SISW_BMIDE_CloseAllDialogsOperations(StrAction)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_BMIDE_CloseAllDialogsOperations"
	'Declaring variables
	Dim ObjChild,ObjModelerChild,StrTabName
	Dim iCount,bFlag,StrButtonLabel,StrMenu
	Dim iX,iY,iHeight,iWidth

	Fn_SISW_BMIDE_CloseAllDialogsOperations=False
	bFlag=False
	'Closing all open dialogs
	If JavaWindow("Business Modeler").GetROProperty("enabled")=0 Then
		Set ObjChild=Description.Create()
		ObjChild("Class Name").value="JavaButton"
		Set ObjModelerChild=JavaWindow("Business Modeler").ChildObjects(ObjChild)
		For iCount=ObjModelerChild.Count-1 To 0 STEP -1
				StrButtonLabel=ObjModelerChild(iCount).GetROProperty("label")
				If LCase(StrButtonLabel)="cancel" OR  LCase(StrButtonLabel)="close" OR  LCase(StrButtonLabel)="ok" Then
					ObjModelerChild(iCount).Click
					bFlag=True
				End If
		Next
	Else
		bFlag=True
	End If 
	wait 1
	'Closing all open tabs
    If JavaWindow("Business Modeler").JavaTab("MainTab").Exist(2) Then
		If CInt(JavaWindow("Business Modeler").JavaTab("MainTab").GetROProperty("items count"))>0 Then
			If JavaWindow("Business Modeler").JavaTab("MainTab").GetROProperty("value")="Help" Then
				Call Fn_BMIDE_TabOperations("Main","Close","Help")
			ElseIf JavaWindow("Business Modeler").JavaTab("MainTab").GetROProperty("value")="Welcome" Then
				Call Fn_BMIDE_TabOperations("Main","Close","Welcome")
			ElseIf JavaWindow("Business Modeler").JavaTab("MainTab").GetROProperty("value")="BMIDE Assistant" Then
				Call Fn_BMIDE_TabOperations("Main","Close","BMIDE Assistant")
			End If
		End If
	End if
	'Closing all open tabs
	If JavaWindow("Business Modeler").JavaTab("CTabFolder").Exist(2) Then
		If CInt(JavaWindow("Business Modeler").JavaTab("CTabFolder").GetROProperty("items count"))>0 Then
			If JavaWindow("Business Modeler").JavaTab("CTabFolder").GetROProperty("value")="BMIDE Assistant" Then
				JavaWindow("Business Modeler").JavaTab("CTabFolder").CloseTab "BMIDE Assistant"
				wait 1
			End If
		End If
	End If
	If JavaWindow("Business Modeler").JavaTab("MainTab").Exist(2) Then
		If CInt(JavaWindow("Business Modeler").JavaTab("MainTab").GetROProperty("items count"))>0 Then
			StrTabName=JavaWindow("Business Modeler").JavaTab("MainTab").GetROProperty("value")
			JavaWindow("Business Modeler").JavaTab("MainTab").Click 5,5,"LEFT"
			Call Fn_BMIDE_TabOperations("Main","Activate",strTabName)
			StrMenu=Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\BMIDEConfig\BMIDE_Menu.xml", "CloseAll")
			Call Fn_BMIDE_MenuOperation("Select", StrMenu)
		End If
	End if
   Select Case StrAction
	 	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	 	Case "Standard"
			bFlag=Fn_BMIDE_ToolbarButtonClick("","Standard")
   End Select
   If bFlag=True Then
		Fn_SISW_BMIDE_CloseAllDialogsOperations=True
	End If
	'Releasing objects
	Set ObjModelerChild=Nothing
	Set ObjChild=Nothing
End Function
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_SISW_BMIDE_ToolbarButtonOperations
'
'Description			 :	Function Used to perform operations on ToolBar Buttons
'
'Parameters			   :  1.StrAction : Action Name
'									2.iInstance : Instance Number
'									3.sButtonName:Button Name
'									4.sMenu:Menu Name
'
'Return Value		   : 	True Or False
'
'Pre-requisite			:	Should Be Log In BMIDE
'
'Examples				: bReturn= Fn_SISW_BMIDE_ToolbarButtonOperations("Exist","","Find Class...","")
'								   bReturn= Fn_SISW_BMIDE_ToolbarButtonOperations("Select","","Find Class...","")
'								   bReturn= Fn_SISW_BMIDE_ToolbarButtonOperations("IsSelected","","Find Class...","")
'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done								Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Pranav Ingle										   		28-Jun-2013			           1.0																						Sandeep N
'													Pranav Ingle										   		26-July-2013			        1.0																						 Sandeep N
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_BMIDE_ToolbarButtonOperations(StrAction,iInstance,sButtonName,sMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_BMIDE_ToolbarButtonOperations"
	'Declaring variables
	Dim ObjDesc, ArrLists, iToolCnt, iCounter, sContents, iCnt
	Fn_SISW_BMIDE_ToolbarButtonOperations=False
	'checking existance of [ Business Modeler ] window
	If JavaWindow("Business Modeler").Exist(20) Then
		'Create Toolbar object
		Set ObjDesc = Description.Create() 
		ObjDesc("to_class").Value = "JavaToolbar" 
		ObjDesc("enabled").Value = 1
		'maximize [ Business Modeler ] window
		JavaWindow("Business Modeler").Maximize
		
		'Get the total of Toolbar objects
		Set ArrLists =JavaWindow("Business Modeler").ChildObjects(ObjDesc)
		iToolCnt = JavaWindow("Business Modeler").ChildObjects(ObjDesc).count
		iCnt =1
		Select Case StrAction
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
			'Case to select and check existance of toolbar buttons
			Case "Exist","Select"
				For iCounter = 0 to iToolCnt-1
					sContents = ArrLists(iCounter).GetContent()
					If instr(sContents, sButtonName) > 0 Then	
						If iInstance<>"" Then
							If iCnt = CInt(iInstance) Then
									If StrAction="Select" Then
										ArrLists(iCounter).Press sButtonName
									End If
									Fn_SISW_BMIDE_ToolbarButtonOperations = TRUE
									Exit For
							End If
							iCnt = iCnt +1
						Else
							If StrAction="Select" Then
								ArrLists(iCounter).Press sButtonName
							End If
							Fn_SISW_BMIDE_ToolbarButtonOperations = TRUE
							Exit For
						End If
					
					End If
				Next
			
				If iCounter = iToolCnt Then
					Fn_SISW_BMIDE_ToolbarButtonOperations = FALSE
				End If
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
			'Case to select and check existance of toolbar buttons
			Case "IsSelected"
				For iCounter = 0 to iToolCnt-1
					sContents = ArrLists(iCounter).GetContent()
					If instr(sContents, sButtonName) > 0 Then	
						If Trim(ArrLists(iCounter).GetSelection()) = Trim(sButtonName) Then
								Fn_SISW_BMIDE_ToolbarButtonOperations = TRUE
								Exit For
						End If
					End If
				Next
			
				If iCounter = iToolCnt Then
					Fn_SISW_BMIDE_ToolbarButtonOperations = FALSE
				End If
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
			Case else
				Fn_SISW_BMIDE_ToolbarButtonOperations = FALSE
		End Select
		'Releasing objects
		Set ObjDesc = Nothing
		Set ArrLists = Nothing
	Else
		Fn_SISW_BMIDE_ToolbarButtonOperations = FALSE
	End If
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_SISW_BMIDE_FilterAttributesTableOperations
'
'Description			 :	Function Used to perform operations on Filter Attribute table
'
'Parameters			   :  1.StrAction: Action Name
'									   2.StrHierarchy: Object Hierarchy
'									   3.StrFilterAttribute: Filter Attribute name
'									   4.StrColumn: Column name
'									   5.StrValue: Value
'
'Return Value		   : 	True or False

'Pre-requisite			:	Filter Attribute table Should be appear

'Example				'msgBox  Fn_SISW_BMIDE_FilterAttributesTableOperations("GetIndex","NewDynamicLOVDialogTable","p3Prop2","","")
									'msgBox  Fn_SISW_BMIDE_FilterAttributesTableOperations("Select","NewDynamicLOVDialogTable","p3Prop2","","")
									'msgBox  Fn_SISW_BMIDE_FilterAttributesTableOperations("MoveUp","NewDynamicLOVDialogTable","p3Prop2","","")
									'msgBox  Fn_SISW_BMIDE_FilterAttributesTableOperations("MoveDown","NewDynamicLOVDialogTable","p3Prop2","","")
									'msgBox  Fn_SISW_BMIDE_FilterAttributesTableOperations("Remove","NewDynamicLOVDialogTable","p3Prop2","","")
'History					 :			
'				Developer Name						Date					Rev. No.						Changes Done																				Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'				Sandeep N							4-July-2013				1.0																																			Avinash J	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_SISW_BMIDE_FilterAttributesTableOperations(StrAction,StrHierarchy,StrFilterAttribute,StrColumn,StrValue)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_BMIDE_FilterAttributesTableOperations"
   Dim objFilterAttributeTable
   Dim bFlag,iCounter
   Fn_SISW_BMIDE_FilterAttributesTableOperations=False
	Select Case StrHierarchy
		Case "NewDynamicLOVDialogTable"
				Set objFilterAttributeTable=JavaWindow("Business Modeler").JavaWindow("NewDynamicLOV")
		Case "MainTabTable"
			Set objFilterAttributeTable=JavaWindow("Business Modeler")
	End Select
   If not objFilterAttributeTable.Exist(10) Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail : [ Filter Attributes ] table not exist")
		Set objFilterAttributeTable=Nothing
		Exit function
   End If
   Select Case StrAction
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	 	Case "Select"
			bFlag=False
			For iCounter=0 to Cint(objFilterAttributeTable.JavaTable("FilterAttributes").GetROProperty("rows"))-1
				If trim(objFilterAttributeTable.JavaTable("FilterAttributes").GetCellData(iCounter,0))=trim(StrFilterAttribute) Then
					objFilterAttributeTable.JavaTable("FilterAttributes").ClickCell iCounter,0
					wait 1
                	objFilterAttributeTable.JavaTable("FilterAttributes").SelectCell iCounter,0
					wait 1
					bFlag=True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Pass : Successfully selected attribute "&StrFilterAttribute&" from [ Filter Attributes ] table")
					Exit for
				End If
			Next
			If bFlag=True Then
				Fn_SISW_BMIDE_FilterAttributesTableOperations=True
			End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "MoveUp"
				If Fn_SISW_BMIDE_FilterAttributesTableOperations("Select",StrHierarchy,StrFilterAttribute,"","")=True Then
					Wait 1
					Fn_SISW_BMIDE_FilterAttributesTableOperations= Fn_Button_Click("Fn_SISW_BMIDE_FilterAttributesTableOperations", objFilterAttributeTable, "MoveUpFilterAttributes")
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Pass : Successfully Move Up attribute "&StrFilterAttribute&" from [ Filter Attributes ] table")
				End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "MoveDown"
				If Fn_SISW_BMIDE_FilterAttributesTableOperations("Select",StrHierarchy,StrFilterAttribute,"","")=True Then
					Fn_SISW_BMIDE_FilterAttributesTableOperations= Fn_Button_Click("Fn_SISW_BMIDE_FilterAttributesTableOperations", objFilterAttributeTable, "MoveDownFilterAttributes")
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Pass : Successfully Move Down attribute "&StrFilterAttribute&" from [ Filter Attributes ] table")
				End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "Remove"
				If Fn_SISW_BMIDE_FilterAttributesTableOperations("Select",StrHierarchy,StrFilterAttribute,"","")=True Then
					Wait 1
					Fn_SISW_BMIDE_FilterAttributesTableOperations= Fn_Button_Click("Fn_SISW_BMIDE_FilterAttributesTableOperations", objFilterAttributeTable, "RemoveFilterAttributes")
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Pass : Successfully Remove attribute "&StrFilterAttribute&" from [ Filter Attributes ] table")
				End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "GetIndex"
			For iCounter=0 to Cint(objFilterAttributeTable.JavaTable("FilterAttributes").GetROProperty("rows"))-1
				If trim(objFilterAttributeTable.JavaTable("FilterAttributes").GetCellData(iCounter,0))=trim(StrFilterAttribute) Then
					wait 1
					Fn_SISW_BMIDE_FilterAttributesTableOperations=iCounter
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Pass : Index of attribute "&StrFilterAttribute&" is "&Cstr(iCounter)&" in [ Filter Attributes ] table")
					Exit for
				End If
			Next
   End Select
   'Releasing Object
   Set objFilterAttributeTable=Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_SISW_BMIDE_AddNewModelElementOperations
'
'Description			 :	Function Used to perform operations on Add New Model Element
'
'Parameters			   :  1.StrAction: Action Name
'									   2.StrProject: Project Name
'									   3.StrElement: Element name
'									   4.StrElementInfo: Element information
'									   5.StrFilterText: Filter text
'									   6.StrButton: Button Name
'
'Return Value		   : 	True or False
'
'Pre-requisite			:	Should be Login to BMIDE Standard perspective
'
'Example				  :	bReturn=Fn_SISW_BMIDE_AddNewModelElementOperations("Select","","Status","","","Next")
'
'History					 :			
'				Developer Name						Date					Rev. No.						Changes Done																				Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'				Sandeep N								5-July-2013				1.0																																					Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_SISW_BMIDE_AddNewModelElementOperations(StrAction,StrProject,StrElement,StrElementInfo,StrFilterText,StrButton)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_BMIDE_AddNewModelElementOperations"
 	'Declaring variables
	Dim ObjNewModelElementDialog
	Dim bFlag,iCounter

	Fn_SISW_BMIDE_AddNewModelElementOperations=False
	'checking existance of [ Add New Model Element ] dialog
	If not JavaWindow("Business Modeler").JavaWindow("AddNewModelElement").Exist(6) Then
		'Click on [ New Model Element ] toolbar to open [ Add New Model Element ] dialog
		If Fn_SISW_BMIDE_ToolbarButtonOperations("Select","","New Model Element (Ctrl+N)","")=False Then
			Exit function
		End If
		wait 2
	End If
	'Creating Object of [ Add New Model Element ] dialog
	Set ObjNewModelElementDialog=JavaWindow("Business Modeler").JavaWindow("AddNewModelElement")
	'Selecting project
	If StrProject<>"" Then
		Call Fn_List_Select("Fn_SISW_BMIDE_AddNewModelElementOperations", ObjNewModelElementDialog, "Project",StrProject)
	End If
	Select Case StrAction
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "Select"
			bFlag=False
			For iCounter=0 to Cint(Fn_UI_Object_GetROProperty("",ObjNewModelElementDialog.JavaTable("ElementsTable"),"rows"))-1
'			Cint(ObjNewModelElementDialog.JavaTable("ElementsTable").GetROProperty("rows"))-1
				If trim(StrElement)=trim(ObjNewModelElementDialog.JavaTable("ElementsTable").GetCellData(iCounter,0)) Then
					ObjNewModelElementDialog.JavaTable("ElementsTable").SelectCell iCounter,0
					wait 2
					bFlag=True
					Exit for
				End If
			Next
			If bFlag=True Then
				If StrButton<>"" Then
					Call Fn_Button_Click("Fn_SISW_BMIDE_AddNewModelElementOperations", ObjNewModelElementDialog, StrButton)
				End If
				Fn_SISW_BMIDE_AddNewModelElementOperations=True
			End If
	End Select
	'Releasing Object of [ Add New Model Element ] dialog
	Set ObjNewModelElementDialog=Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_SISW_BMIDE_DynamicLOVTestResultsTableOperations
'
'Description			 :	Function Used to perform operations on LOV Test Results Table
'
'Parameters			   :  1.StrAction: Action Name
'									   2.StrHierarchy: Hierarchy
'									   3.dicDynamicLOVTestInfo: Dynamic LOV Test Result Information
'									   4.StrButton: Button Name
'
'Return Value		   : 	True or False
'
'Pre-requisite			:	Should be Login to BMIDE
'
'Example				  :		  Set dicDynamicLOVTestInfo=CreateObject("Scripting.Dictionary")
'											dicDynamicLOVTestInfo("ServerProfile")="DemoProfile"
'											dicDynamicLOVTestInfo("UserID")="AutoTestDBA"
'											dicDynamicLOVTestInfo("Password")="AutoTestDBA"
'											dicDynamicLOVTestInfo("Role")="DBA"
'											dicDynamicLOVTestInfo("Group")="dba"
'											dicDynamicLOVTestInfo("PrimaryColumn")="item_id"
'											dicDynamicLOVTestInfo("PrimaryValue")="000028~000029~000030"
'											dicDynamicLOVTestInfo("Column")="object_name~object_name~object_name"
'											dicDynamicLOVTestInfo("Value")="TestItem~TestItem1~TestItem2"
'											bReturn=Fn_SISW_BMIDE_DynamicLOVTestResultsTableOperations("CellExist","NewDynamicLOVDialog",dicDynamicLOVTestInfo,"")
'
'History					 :			
'				Developer Name						Date			  Rev. No.						Changes Done																	Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'				Sandeep N						5-July-2013				1.0																											Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -- - -
'				Nitish Bharadwaj				19-Feb-2016				1.0					Added New Case "ColumnsExist"											[TC1122:2016012700:19Feb2016:AnkitN:NewDevelopment]		
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -- - -
Function Fn_SISW_BMIDE_DynamicLOVTestResultsTableOperations(StrAction,StrHierarchy,dicDynamicLOVTestInfo,StrButton)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_BMIDE_DynamicLOVTestResultsTableOperations"
 	'Declaring Variables
	Dim ObjTestDynamicLOVDialog,strcolumn
	Dim arrPrimaryValues,arrColumn,arrValue,iCount,iCounter,bFlag

	Fn_SISW_BMIDE_DynamicLOVTestResultsTableOperations=False
	If Not JavaWindow("Business Modeler").JavaWindow("TestDynamicLOV").Exist(6) Then
	   Select Case StrHierarchy
	 		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	 		'Case to perform operations from [ New Dynamic LOV ] dialog
			Case "NewDynamicLOVDialog"
				Call Fn_Button_Click("Fn_SISW_BMIDE_DynamicLOVTestResultsTableOperations",JavaWindow("Business Modeler").JavaWindow("NewDynamicLOV"),"TestDynamicLOV")
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
			Case "MainTabTable"
				Call Fn_Button_Click("Fn_SISW_BMIDE_DynamicLOVTestResultsTableOperations",JavaWindow("Business Modeler"),"TestDynamicLOV")
		End Select
		wait 2
	End If
	'Checking Existance of [ Test Dynamic LOV ] dialog
	If JavaWindow("Business Modeler").JavaWindow("TestDynamicLOV").Exist(6) Then
		'Creating Object of [ Test Dynamic LOV ] dialog
		Set ObjTestDynamicLOVDialog=JavaWindow("Business Modeler").JavaWindow("TestDynamicLOV")
	End If

	If ObjTestDynamicLOVDialog.JavaTable("DynamicLOVTestResultsTable").Exist(2) Then
		'Do nothing
	Else
		'Selecting project
		If dicDynamicLOVTestInfo("Project")<>"" Then
			Call Fn_List_Select("Fn_SISW_BMIDE_DynamicLOVTestResultsTableOperations", ObjTestDynamicLOVDialog,"Project",dicDynamicLOVTestInfo("Project"))
		End If
		'Selecting Server Profile
		If dicDynamicLOVTestInfo("ServerProfile")<>"" Then
			Call Fn_List_Select("Fn_SISW_BMIDE_DynamicLOVTestResultsTableOperations", ObjTestDynamicLOVDialog,"ServerProfile",dicDynamicLOVTestInfo("ServerProfile"))
		End If
		'Checking Existance of [ User ID ] field
		If ObjTestDynamicLOVDialog.JavaEdit("UserID").Exist(6) Then
			'Setting User ID
			If dicDynamicLOVTestInfo("UserID") <> "" Then
				Call Fn_Edit_Box("Fn_SISW_BMIDE_DynamicLOVTestResultsTableOperations",ObjTestDynamicLOVDialog,"UserID",dicDynamicLOVTestInfo("UserID"))
			End If	
			'Setting Password
			If dicDynamicLOVTestInfo("Password") <> "" Then
				Call Fn_Edit_Box("Fn_SISW_BMIDE_DynamicLOVTestResultsTableOperations",ObjTestDynamicLOVDialog,"Password",dicDynamicLOVTestInfo("Password"))
			End If
			'Setting Group
			If dicDynamicLOVTestInfo("Group") <> "" Then
				Call Fn_Edit_Box("Fn_SISW_BMIDE_DynamicLOVTestResultsTableOperations",ObjTestDynamicLOVDialog,"Group",dicDynamicLOVTestInfo("Group"))
			End If
			'Setting Role
			If dicDynamicLOVTestInfo("Role") <> "" Then
				Call Fn_Edit_Box("Fn_SISW_BMIDE_DynamicLOVTestResultsTableOperations",ObjTestDynamicLOVDialog,"Role",dicDynamicLOVTestInfo("Role"))
			End If
			'Clicking On [ Connect ] button
			Call Fn_Button_Click("Fn_SISW_BMIDE_DynamicLOVTestResultsTableOperations",ObjTestDynamicLOVDialog,"Connect")
		End If
		'Clicking On [ Next ] Button
		ObjTestDynamicLOVDialog.JavaButton("Next").WaitProperty "enabled",1,6000
		Call Fn_Button_Click("Fn_SISW_BMIDE_DynamicLOVTestResultsTableOperations",ObjTestDynamicLOVDialog,"Next")
	End If
	Select Case StrAction
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to check value exist in Cell
		Case "CellExist"
			If dicDynamicLOVTestInfo("PrimaryColumn")="" Then
				dicDynamicLOVTestInfo("PrimaryColumn")=0
			End If
			arrPrimaryValues=Split(dicDynamicLOVTestInfo("PrimaryValue"),"~")
			arrColumn=Split(dicDynamicLOVTestInfo("Column"),"~")
			arrValue=Split(dicDynamicLOVTestInfo("Value"),"~")
			For iCount=0 to ubound(arrPrimaryValues)
				bFlag=False
				For iCounter=0 to Cint(ObjTestDynamicLOVDialog.JavaTable("DynamicLOVTestResultsTable").GetROProperty("rows"))-1
					If isNumeric(arrPrimaryValues(iCount)) Then
						arrPrimaryValues(iCount)=Cdbl(arrPrimaryValues(iCount))
					End If
					If Trim(Cstr(ObjTestDynamicLOVDialog.JavaTable("DynamicLOVTestResultsTable").GetCellData(iCounter,dicDynamicLOVTestInfo("PrimaryColumn"))))=Trim(Cstr(arrPrimaryValues(iCount))) Then
						If Trim(ObjTestDynamicLOVDialog.JavaTable("DynamicLOVTestResultsTable").GetCellData(iCounter,arrColumn(iCount)))=arrValue(iCount) Then
							bFlag=True
							Exit for
						End If
					End If
				Next
				If bFlag=False Then
					Exit For
				End If
			Next
			If bFlag Then
				Fn_SISW_BMIDE_DynamicLOVTestResultsTableOperations=True
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to click button
		Case "ClickButton"
			Fn_SISW_BMIDE_DynamicLOVTestResultsTableOperations=Fn_Button_Click("Fn_SISW_BMIDE_DynamicLOVTestResultsTableOperations",ObjTestDynamicLOVDialog,dicDynamicLOVTestInfo("ButtonName"))
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to check column names
		Case "ColumnsExist"					'[TC1122:2016012700:19Feb2016:NitishB:NewDevelopment] - - Added New Case to verify column existance		
			arrPrimaryValues = split(dicDynamicLOVTestInfo("Column"),"~")
			strColumn = ObjTestDynamicLOVDialog.JavaTable("DynamicLOVTestResultsTable").GetROProperty("columns_names")
			arrColumn = split(strColumn,";")
			For iCounter = 0 To Ubound(arrPrimaryValues)
				bFlag = False
				For iCount = 0 To Ubound(arrColumn)
					If Trim(lcase(arrColumn(iCount))) = Trim(lcase(arrPrimaryValues(iCounter))) Then
						bFlag = true
						Exit For
					End If
				Next
				If bFlag = False Then
					Fn_SISW_BMIDE_DynamicLOVTestResultsTableOperations=False
					Exit For
				End If
			Next
			If bFlag Then
				Fn_SISW_BMIDE_DynamicLOVTestResultsTableOperations=True
			End If
			
	End Select
	If StrButton<>"" Then
		Call Fn_Button_Click("Fn_SISW_BMIDE_DynamicLOVTestResultsTableOperations",ObjTestDynamicLOVDialog,StrButton)
	End If
	'Releasing object
	Set ObjTestDynamicLOVDialog=Nothing
End Function

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_SISW_BMIDE_ToolbarCustmizationOperations
'
'Description			 :	Function Used to perform operations on Toolbar Custmization from Preference Window
'
'Parameters			   :  	  1.StrAction: Action Name
'									   2.sToolbarShownList: Toolbar shown List
'									   3.sToolbarHideList: Toolbar hidden List
'									   4.StrButton: Button Name
'
'Return Value		   : 	True or False

'Examples				:   bReturn= Fn_SISW_BMIDE_ToolbarCustmizationOperations("AddToShownList","","AliasID Rule~Tool","OK")
'									bReturn= Fn_SISW_BMIDE_ToolbarCustmizationOperations("AddToHideList","AliasID Rule~Tool","","OK")
'									bReturn= Fn_SISW_BMIDE_ToolbarCustmizationOperations("AddAllToHideList","","","OK")
'									bReturn= Fn_SISW_BMIDE_ToolbarCustmizationOperations("AddAllToShownList","","","","OK")
'									bReturn= Fn_SISW_BMIDE_ToolbarCustmizationOperations("RestoreDefaults","","","","OK")
'
'History					 :			
'				Developer Name						Date						Rev. No.						Changes Done									Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'				Pranav Ingle							23-July-2013				1.0																							Sandeep N
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Function Fn_SISW_BMIDE_ToolbarCustmizationOperations(StrAction,sToolbarShownList,sToolbarHideList,StrButton)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_BMIDE_ToolbarCustmizationOperations"
 	'Declaring variables
	Dim ObjPreference,ObjBOList, myDeviceReplay
	Dim aBOName,iCount,iCounter,bFlag

	Fn_SISW_BMIDE_ToolbarCustmizationOperations=False

	If Fn_UI_ObjectExist("Fn_SISW_BMIDE_ToolbarCustmizationOperations",JavaWindow("Business Modeler").JavaWindow("Preferences"))=False Then
	   'Calling Window:Preference Menu to open Preference Dialog
		strMenu=Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\BMIDEConfig\BMIDE_Menu.xml", "Preferences")
		Call Fn_BMIDE_MenuOperation("Select", strMenu)
   End If

	'Checking existance of [ Preferences ] dialog
	If JavaWindow("Business Modeler").JavaWindow("Preferences").Exist(5) Then
		'creating object of [ Preferences ] dialog
		Set ObjPreference=JavaWindow("Business Modeler").JavaWindow("Preferences")
		wait 2
	Else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail : [ Preferences ] dialog not exist")
		Exit function
	End If

	Call Fn_UI_JavaTree_Expand("Fn_SISW_BMIDE_ToolbarCustmizationOperations",ObjPreference,"PreferenceTree","Teamcenter")
	'Selecting "Teamcenter:Server Connection Profiles" node 
	Call Fn_JavaTree_Select("Fn_SISW_BMIDE_ToolbarCustmizationOperations",ObjPreference,"PreferenceTree","Teamcenter:Toolbar Customization")
	
	Select Case StrAction		
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'case to Add Business Objects from Hidden Business Objects list to Displayed Business Objects list
		Case "AddToShownList", "AddToHideList"
			If StrAction = "AddToShownList" Then
				Set ObjBOList=ObjPreference.JavaTable("ToolbarCustomization")
				aBOName=Split(sToolbarHideList,"~")
			Else
				Set ObjBOList=ObjPreference.JavaTable("ShownToolbar")
				aBOName=Split(sToolbarShownList,"~")
			End If
			
			For iCount=0 to Ubound(aBOName)
				bFlag=False
				For iCounter=0 to Cint(ObjBOList.GetROProperty("rows"))-1
					If trim(aBOName(iCount))=trim(ObjBOList.GetCellData(iCounter,0)) Then
						If iCount = 1 Then
							Set myDeviceReplay = CreateObject("Mercury.DeviceReplay")
							myDeviceReplay.KeyDown 29
						End If
						ObjBOList.ClickCell iCounter, 0,"LEFT"
						wait 1
						bFlag=True
						Exit For
					End If
				Next
				If bFlag=False Then
					Exit for
				End If
			Next
			If iCount > 0 Then
				myDeviceReplay.KeyUp 29
			End If
			If bFlag=True Then
				If StrAction = "AddToShownList" Then
					Fn_SISW_BMIDE_ToolbarCustmizationOperations = Fn_Button_Click("Fn_SISW_BMIDE_ToolbarCustmizationOperations", ObjPreference, "Right")
				Else
					Fn_SISW_BMIDE_ToolbarCustmizationOperations = Fn_Button_Click("Fn_SISW_BMIDE_ToolbarCustmizationOperations", ObjPreference,"Left")
				End If
			End If
			Set ObjBOList=Nothing
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'case to Add All Business Objects from Hidden Toolbar list to Shown Toolbar list
		Case "AddAllToShownList"
			Fn_SISW_BMIDE_ToolbarCustmizationOperations=Fn_Button_Click("Fn_SISW_BMIDE_ToolbarCustmizationOperations", ObjPreference, "RightAll")
			wait 2
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'case to Add All Business Objects from Shown Toolbar list To Hidden Toolbar list 
		Case "AddAllToHideList"
        	Fn_SISW_BMIDE_ToolbarCustmizationOperations=Fn_Button_Click("Fn_SISW_BMIDE_ToolbarCustmizationOperations", ObjPreference, "LeftAll")
			wait 2
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'case to Restore Defaults
		Case "RestoreDefaults"
        	Fn_SISW_BMIDE_ToolbarCustmizationOperations=Fn_Button_Click("Fn_SISW_BMIDE_ToolbarCustmizationOperations", ObjPreference, "Restore Defaults")
			wait 2
	End Select

	If StrButton<>"" Then
		Call Fn_Button_Click("Fn_SISW_BMIDE_ToolbarCustmizationOperations", ObjPreference,StrButton)
	End If
	Set ObjPreference=Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_SISW_BMIDE_HTMLReportOperations
'
'Description			 :	Function Used to perform operations Reports genearated by Report wizard
'
'Parameters			   :  1.StrAction: Action Name
'									   2.dicReportInfo: Report Information
'									   3.bOpen: Flag to open report in Browser
'									   4.bClose: Flag to close report Browser
'
'Return Value		   : 	True or False
'
'Pre-requisite			:	Reports tab should be Activated
'
'Example				  :		Set dicReportInfo=CreateObject("Scripting.Dictionary")
'										dicReportInfo("Attribute")="Name~Description"
'										dicReportInfo("Column")="Value"
'										dicReportInfo("Value")="Fnd0DMTemplateCondition~If application name of revision is RM then only create from template"
'										bReturn=Fn_SISW_BMIDE_HTMLReportOperations("VerifyFromConditionDetailsTable",dicReportInfo,"true","true")
'
'										dicReportInfo("FieldName")="Name"
'										dicReportInfo("Value")="AllocationMapRevMaster"
'										bReturn=Fn_SISW_BMIDE_HTMLReportOperations("VerifyFromFieldSummaryTable",dicReportInfo,"","")
'										
'										dicReportInfo("DataModelName")="Business Object"
'										bReturn=Fn_SISW_BMIDE_HTMLReportOperations("SelectTeamcenterDataModelType",dicReportInfo,"","")
'										
'										dicReportInfo("ObjectName")="BOMWindow"
'										bReturn=Fn_SISW_BMIDE_HTMLReportOperations("SelectObject",dicReportInfo,"","")
'										
'										dicReportInfo("Value")="This document is the Data Model specification for Teamcenter"
'										bReturn=Fn_SISW_BMIDE_HTMLReportOperations("VerifyOverviewContents",dicReportInfo,"","")
'										
'										dicReportInfo("TemplateName")="bmideprj25jun"
'										dicReportInfo("Column")="Display Name"
'										dicReportInfo("Value")="BMIDEPrj25Jun"
'										bReturn=Fn_SISW_BMIDE_HTMLReportOperations("VerifyFromTemplateDataModelTable",dicReportInfo,"","")
'										
'										dicReportInfo("Library")="None Found"
'										dicReportInfo("Column")="Library"
'										dicReportInfo("Value")="None Found"
'										bReturn=Fn_SISW_BMIDE_HTMLReportOperations("VerifyFromDeprecatedLibrariesTable",dicReportInfo,"","")
'										
'										dicReportInfo("BusinessObject")="ImanItemLine"
'										dicReportInfo("Column")="Operation"
'										dicReportInfo("Value")="fnd0getCustomConfiguredIrf"
'										bReturn=Fn_SISW_BMIDE_HTMLReportOperations("VerifyFromDeprecatedBusinessObjectOperationsTable",dicReportInfo,"","")
'
'History					 :			
'				Developer Name						Date					Rev. No.						Changes Done																				Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'				Sandeep N								30-July-2013				1.0																																	Preeti S
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Public Function Fn_SISW_BMIDE_HTMLReportOperations(StrAction,dicReportInfo,bOpen,bClose)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_BMIDE_HTMLReportOperations"
   'Declaring variables
   Dim StrURL,iColNumber,iCounter,iCount,bFlag,arrAttribute,arrValue
   Dim objBrowser,handlewin,objReportPage
   Dim iRowNumber,objLink,objChild,StrLinkName,objContent,StrContent,objTable
   Fn_SISW_BMIDE_HTMLReportOperations=False
   'Checking Open flag
   If LCase(bOpen)="true" Then
 		'Checking existance of [ Report URL ] list
		If dicReportInfo("ReportURL")<>"" Then
			StrURL=dicReportInfo("ReportURL")
		ElseIf JavaWindow("Business Modeler").JavaList("ReportURL").Exist(5) Then
			'Copy URL
			StrURL=JavaWindow("Business Modeler").JavaList("ReportURL").GetItem("#0")
		Else
			Exit function
		End If
		'Creating object of [ Internet Explorer ]
		Set objBrowser = CreateObject("InternetExplorer.Application")
		objBrowser.Visible = True
		'Set URL in browser
		objBrowser.Navigate StrURL
		handlewin = objBrowser.HWND
		Window("hwnd:="+CStr(handlewin)).Maximize
		'Releasing object of [ Internet Explorer ]
		Set objBrowser =Nothing
   End If
	'Checking Existance of [ Report ] page
	If Browser("Browser").Page("ReportPage").Exist(10) Then
		'Creating Object of [ Report ] page
		Set objReportPage=Browser("Browser").Page("ReportPage")
	Else
		Exit function
	End If
   Select Case StrAction
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "VerifyFromDeprecatedLibrariesTable"
			Set objChild= Description.Create()
			objChild("micclass").value="Link"
			objChild("html tag").value="A"
			objChild("text").value="Deprecated"
			Set objLink=objReportPage.Frame("Contents").WebTable("ContentType").ChildObjects(objChild)
			If Cint(objLink.count)=1 Then
				objLink(0).Click
				wait 2
			Else
				Set objChild=Nothing
				Set objLink=Nothing
				Set objReportPage=Nothing
				Exit function
			End if
			Set objChild=Nothing
			Set objLink=Nothing
        	Set objTable=objReportPage.Frame("Contents").WebTable("DeprecatedLibraries")
		
			'Checking Existance of [ Deprecated ] table
			If objTable.Exist(5) Then
				If dicReportInfo("Column")<>"" Then
					'get column number from [ Condition Details ] table of specific Column
					iColNumber=False
					For iCounter=0 to objTable.ColumnCount(1)
						'Matching column name with expected column
						If trim(objTable.GetCellData(1,iCounter))=trim(dicReportInfo("Column")) Then
							iColNumber=iCounter
							Exit for
						End If
					Next
					If iColNumber=False Then
						If lcase(bClose)="true" Then
							Browser("Browser").Close
					   End If
					   Exit function
					End If
					dicReportInfo("Column")=iColNumber
				Else
					dicReportInfo("Column")=0
				End If
				arrAttribute=Split(dicReportInfo("Library"),"~")
				arrValue=Split(dicReportInfo("Value"),"~")

				For iCount=0 to ubound(arrAttribute)
					bFlag=False
					'Loop to all row of table
					For iCounter=2 to objTable.RowCount
						'Matching Attribute
						If trim(objTable.GetCellData(iCounter,1))=trim(arrAttribute(iCount)) Then
							'Matching current value with expected value
							If trim(objTable.GetCellData(iCounter,dicReportInfo("Column")))=trim(arrValue(iCount)) Then
								bFlag=True
								Exit for
							End If	
						End If
					Next
					If bFlag=False Then
						Exit for
					End If
				Next
			End If
			If bFlag=True Then
				Fn_SISW_BMIDE_HTMLReportOperations=True
			End If
			Set objTable=Nothing
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	 	Case "VerifyFromDeprecatedBusinessObjectOperationsTable","VerifyFromDeprecatedPropertyOperationsTable"
			Set objChild= Description.Create()
			objChild("micclass").value="Link"
			objChild("html tag").value="A"
			objChild("text").value="Deprecated"
			Set objLink=objReportPage.Frame("Contents").WebTable("ContentType").ChildObjects(objChild)
			If Cint(objLink.count)=1 Then
				objLink(0).Click
				wait 2
			Else
				Set objChild=Nothing
				Set objLink=Nothing
				Set objReportPage=Nothing
				Exit function
			End if
			Set objChild=Nothing
			Set objLink=Nothing
			Select Case StrAction
				Case "VerifyFromDeprecatedBusinessObjectOperationsTable"
					Set objTable=objReportPage.Frame("Contents").WebTable("DeprecatedBusinessObjectOperations")
				Case "VerifyFromDeprecatedPropertyOperationsTable"
					Set objTable=objReportPage.Frame("Contents").WebTable("DeprecatedPropertyOperations")
			End Select
			'Checking Existance of [ Deprecated ] table
			If objTable.Exist(5) Then
				If dicReportInfo("Column")<>"" Then
					'get column number from [ Condition Details ] table of specific Column
					iColNumber=False
					For iCounter=0 to objTable.ColumnCount(1)
						'Matching column name with expected column
						If trim(objTable.GetCellData(1,iCounter))=trim(dicReportInfo("Column")) Then
							iColNumber=iCounter
							Exit for
						End If
					Next
					If iColNumber=False Then
						If lcase(bClose)="true" Then
							Browser("Browser").Close
					   End If
					   Exit function
					End If
					dicReportInfo("Column")=iColNumber
				Else
					dicReportInfo("Column")=0
				End If
				arrAttribute=Split(dicReportInfo("BusinessObject"),"~")
				arrValue=Split(dicReportInfo("Value"),"~")

				For iCount=0 to ubound(arrAttribute)
					bFlag=False
					'Loop to all row of table
					For iCounter=2 to objTable.RowCount
						'Matching Attribute
						If trim(objTable.GetCellData(iCounter,1))=trim(arrAttribute(iCount)) Then
							'Matching current value with expected value
							If trim(objTable.GetCellData(iCounter,dicReportInfo("Column")))=trim(arrValue(iCount)) Then
								bFlag=True
								Exit for
							End If	
						End If
					Next
					If bFlag=False Then
						Exit for
					End If
				Next
			End If
			If bFlag=True Then
				Fn_SISW_BMIDE_HTMLReportOperations=True
			End If
			Set objTable=Nothing
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	 	Case "VerifyFromFieldSummaryTable"
			'' create object for Field Summary Table
			set oDesc = Description.Create()
			oDesc("html tag").value = "TABLE"
			oDesc("text").value = "Field Summary.*"
			'oDesc("name").value = "RuntimeBusinessObject"
			set objFieldSummaryTable = objReportPage.Frame("Contents").ChildObjects(oDesc)

			'Checking Existance of [ FieldSummary ] table
			If objFieldSummaryTable(0).Exist(5) Then
				If dicReportInfo("Column")="" Then
					dicReportInfo("Column")=2
				End If
				arrAttribute=Split(dicReportInfo("FieldName"),"~")
				arrValue=Split(dicReportInfo("Value"),"~")

				For iCount=0 to ubound(arrAttribute)
					bFlag=False
					'Loop to all row of table
					For iCounter=2 to objFieldSummaryTable(0).RowCount
						'Matching Attribute
						If trim(objFieldSummaryTable(0).GetCellData(iCounter,1))=trim(arrAttribute(iCount)) Then
							'Matching current value with expected value
							If trim(objFieldSummaryTable(0).GetCellData(iCounter,dicReportInfo("Column")))=trim(arrValue(iCount)) Then
								bFlag=True
								Exit for
							End If	
						End If
					Next
					If bFlag=False Then
						Exit for
					End If
				Next
			End If
			If bFlag=True Then
				Fn_SISW_BMIDE_HTMLReportOperations=True
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	 	Case "VerifyFromTemplateDataModelTable"
			Set objChild= Description.Create()
			objChild("micclass").value="Link"
			objChild("html tag").value="A"
			objChild("text").value="Template Data Model"
			Set objLink=objReportPage.Frame("Contents").WebTable("ContentType").ChildObjects(objChild)
			If Cint(objLink.count)=1 Then
				objLink(0).Click
				wait 2
			Else
				Set objChild=Nothing
				Set objLink=Nothing
				Set objReportPage=Nothing
				Exit function
			End if
			Set objChild=Nothing
			Set objLink=Nothing
			'Checking Existance of [ Template Data Model ] table
			If objReportPage.Frame("Contents").WebTable("TemplateInformation").Exist(5) Then
				If dicReportInfo("Column")<>"" Then
					'get column number from [ Condition Details ] table of specific Column
					iColNumber=False
					For iCounter=0 to objReportPage.Frame("Contents").WebTable("TemplateInformation").ColumnCount(1)
						'Matching column name with expected column
						If trim(objReportPage.Frame("Contents").WebTable("TemplateInformation").GetCellData(1,iCounter))=trim(dicReportInfo("Column")) Then
							iColNumber=iCounter
							Exit for
						End If
					Next
					If iColNumber=False Then
						If lcase(bClose)="true" Then
							Browser("Browser").Close
					   End If
					   Exit function
					End If
					dicReportInfo("Column")=iColNumber
				Else
					dicReportInfo("Column")=0
				End If
				arrAttribute=Split(dicReportInfo("TemplateName"),"~")
				arrValue=Split(dicReportInfo("Value"),"~")

				For iCount=0 to ubound(arrAttribute)
					bFlag=False
					'Loop to all row of table
					For iCounter=2 to objReportPage.Frame("Contents").WebTable("TemplateInformation").RowCount
						'Matching Attribute
						If trim(objReportPage.Frame("Contents").WebTable("TemplateInformation").GetCellData(iCounter,1))=trim(arrAttribute(iCount)) Then
							'Matching current value with expected value
							If trim(objReportPage.Frame("Contents").WebTable("TemplateInformation").GetCellData(iCounter,dicReportInfo("Column")))=trim(arrValue(iCount)) Then
								bFlag=True
								Exit for
							End If	
						End If
					Next
					If bFlag=False Then
						Exit for
					End If
				Next
			End If
			If bFlag=True Then
				Fn_SISW_BMIDE_HTMLReportOperations=True
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to verify contents
	 	Case "VerifyOverviewContents","VerifyGlossaryContents","VerifyWhatsNewContents"
			Select Case StrAction
				Case "VerifyOverviewContents"
					StrLinkName="Overview"
				Case "VerifyGlossaryContents"
					StrLinkName="Glossary"
				Case "VerifyWhatsNewContents"
					StrLinkName="Whats new"
			End Select
			Set objChild= Description.Create()
			objChild("micclass").value="Link"
			objChild("html tag").value="A"
			objChild("text").value=StrLinkName
			Set objLink=objReportPage.Frame("Contents").WebTable("ContentType").ChildObjects(objChild)
			If Cint(objLink.count)=1 Then
				objLink(0).Click
				wait 2
				Set objChild=Nothing
				Set objChild= Description.Create()
				objChild("micclass").value="WebElement"
				Set objContent=objReportPage.Frame("Contents").ChildObjects(objChild)
				For iCounter=1 to objContent.count-1
					StrContent=objContent(iCounter).GetROProperty("innertext")
					If instr(1,Trim(StrContent),Trim(dicReportInfo("Value"))) Then
						Fn_SISW_BMIDE_HTMLReportOperations=True
						Exit for
					End If
				Next
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to Select Data Model from Overview List
	 	Case "SelectTeamcenterDataModelType"
			If objReportPage.Frame("TeamcenterDataModel").WebTable("DataModelList").Exist(5) Then
				iRowNumber=objReportPage.Frame("TeamcenterDataModel").WebTable("DataModelList").GetRowWithCellText(dicReportInfo("DataModelName"))
				If Cint(iRowNumber)<>-1 Then
					Set objLink=objReportPage.Frame("TeamcenterDataModel").WebTable("DataModelList").ChildItem(iRowNumber,1,"Link",0)
					If TypeName(objLink)<>"Nothing" Then
						objLink.Click
						wait 2
						Fn_SISW_BMIDE_HTMLReportOperations=True
					End If
					Set objLink=Nothing
				End If
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to Select Data Model from Overview List
	 	Case "SelectObject"
			If objReportPage.Frame("AllObject").WebTable("ObjectList").Exist(5) Then
					Set objChild= Description.Create()
					objChild("micclass").value="Link"
					objChild("html tag").value="A"
					objChild("text").value=dicReportInfo("ObjectName")
                    Set objLink=objReportPage.Frame("AllObject").WebTable("ObjectList").ChildObjects(objChild)
					If Cint(objLink.count)=1 Then
						objLink(0).Click
						wait 2
						Fn_SISW_BMIDE_HTMLReportOperations=True
					End If
					Set objLink=Nothing
					Set objChild=Nothing
			End If
	 	'Case to verify data from Condition Details table
	 	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	 	Case "VerifyFromConditionDetailsTable"
			'Checking Existance of [ Condition Details ] table
			If objReportPage.WebTable("ConditionDetails").Exist(5) Then
				If dicReportInfo("Column")<>"" Then
					'get column number from [ Condition Details ] table of specific Column
					iColNumber=False
					For iCounter=0 to objReportPage.WebTable("ConditionDetails").ColumnCount(1)
						'Matching column name with expected column
						If trim(objReportPage.WebTable("ConditionDetails").GetCellData(1,iCounter))=trim(dicReportInfo("Column")) Then
							iColNumber=iCounter
							Exit for
						End If
					Next
					If iColNumber=False Then
						If lcase(bClose)="true" Then
							Browser("Browser").Close
					   End If
					   Exit function
					End If
					dicReportInfo("Column")=iColNumber
				Else
					dicReportInfo("Column")=0
				End If
				arrAttribute=Split(dicReportInfo("Attribute"),"~")
				arrValue=Split(dicReportInfo("Value"),"~")

				For iCount=0 to ubound(arrAttribute)
					bFlag=False
					'Loop to all row of table
					For iCounter=2 to objReportPage.WebTable("ConditionDetails").RowCount
						'Matching Attribute
						If trim(objReportPage.WebTable("ConditionDetails").GetCellData(iCounter,1))=trim(arrAttribute(iCount)) Then
							'Matching current value with expected value
							If trim(objReportPage.WebTable("ConditionDetails").GetCellData(iCounter,dicReportInfo("Column")))=trim(arrValue(iCount)) Then
								bFlag=True
								Exit for
							End If	
						End If
					Next
					If bFlag=False Then
						Exit for
					End If
				Next
			End If
			If bFlag=True Then
				Fn_SISW_BMIDE_HTMLReportOperations=True
			End If
   End Select
   If lcase(bClose)="true" Then
		Browser("Browser").Close
   End If
   'Releasing Report page object
   Set objReportPage=Nothing
End Function

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_BMIDE_PushTemplatetoReferenceDir
'
'Description			 :	Function Used to perform operations for create template reference (Dependent Template)
'
'Parameters			   :  1.sProject: Template Project Name
'						2.sTarget: Target Folder Path
''
'Return Value		   : 	True or False
'
'Pre-requisite			:	Should be Login to BMIDE client
'
'Example				  :	bReturn=Fn_BMIDE_PushTemplatetoReferenceDir("t3testproject","C:\Program Files\Siemens\Teamcenter11\bmide\templates")
'
'History					 :			
'				Developer Name						Date					Rev. No.						Changes Done																				Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'				Paresh D								4-Aug-2014				1.0																										
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Public Function Fn_BMIDE_PushTemplatetoReferenceDir(sProject, sTarget)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDE_PushTemplatetoReferenceDir"
	'Declaring Variables
	Dim objTem , strMenu
	Fn_BMIDE_PushTemplatetoReferenceDir= False
	Set objTem = JavaWindow("Business Modeler").JavaWindow("PushTemplate")
	
	'Checking Existence
	If Fn_SISW_UI_Object_Operations("Fn_BMIDE_PushTemplatetoReferenceDir", "Exist", objTem, "") = False  Then
		strMenu=Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\BMIDEConfig\BMIDE_Menu.xml", "PushTemplatetoRefDir")
		Call Fn_BMIDE_MenuOperation("Select", strMenu)
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Sucessfully performed menu operation [ "+strMenu+" ]")	
	End If
	
	If Fn_SISW_UI_Object_Operations("Fn_BMIDE_PushTemplatetoReferenceDir", "Exist", objTem, "") = True Then
	'Selecting Project	
		If sProject <> "" Then
		  	Call Fn_SISW_UI_JavaList_Operations("Fn_BMIDE_PushTemplatetoReferenceDir", "Select", objTem, "Project", sProject, "", "")
			If Err.Number < 0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed To  Select  project [ "+sProject+" ]")
				Exit Function 
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Sucessfully Selected  project [ "+sProject+" ]")
			End If
		End If
	'Set target folder path
		If sTarget <> "" Then
			Call Fn_SISW_UI_JavaCheckBox_Operations("Fn_BMIDE_PushTemplatetoReferenceDir", "Set", objTem, "Usedefaultlocation", "OFF")
			Call Fn_SISW_UI_JavaEdit_Operations("Fn_BMIDE_PushTemplatetoReferenceDir", "Set", objTem, "Targetfolder", "")
			wait 1
			Call Fn_SISW_UI_JavaEdit_Operations("Fn_BMIDE_PushTemplatetoReferenceDir", "Set", objTem, "Targetfolder", sTarget)
			If Err.Number < 0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to enterTarget Folder path  [ "+sTarget+" ]")
				Exit Function 
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Sucessfully entered Target Folder path  [ "+sTarget+" ]")
			End If
		End If
	'Click on finish button
		Call Fn_SISW_UI_JavaButton_Operations("Fn_BMIDE_PushTemplatetoReferenceDir", "Click", objTem, "Finish")
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Sucessfully clicked on Finish button.")
	Else
		   Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to exist Push Template window object.")
		   Fn_BMIDE_PushTemplatetoReferenceDir = False
		   Exit Function
	End If

   Fn_BMIDE_PushTemplatetoReferenceDir = True
   Set objTem = nothing
End Function
