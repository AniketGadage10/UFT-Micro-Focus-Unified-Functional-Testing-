Option Explicit
'Function List
'************************************************************************************************************************************************************************************************************
'000. Fn_SISW_ReportBuilder_GetObject()
'001. Fn_RB_ReportBuilderTreeOperations()
'002. Fn_RB_ReportDataOperations()
'003. Fn_RB_CreateReportDefinitionTemplate()
'************************************************************************************************************************************************************************************************************
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'
''Function Name		 	:	Fn_SISW_ReportBuilder_GetObject
'
''Description		    :  	Function to get specified Object hierarchy.

''Parameters		    :	1. sObjectName : Object Handle name
								
''Return Value		    :  	Object \ Nothing
'
''Examples		     	:	Fn_SISW_ReportBuilder_GetObject("Remove")

'History:
'	Developer Name			Date							Rev. No.		Reviewer										Changes Done	
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Veena Gurjar			19-March-2013      			1					Sandeep Navghane
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------------------------------
'	Anurag Khera		 26-Sept-2013		2.0								Externalized object hierarchies
'-----------------------------------------------------------------------------------------------------------------------------------
Function Fn_SISW_ReportBuilder_GetObject(sObjectName)
	Dim sObjectXMLPath
	sObjectXMLPath = Fn_GetEnvValue("User", "AutomationDir") & "\TestData\AutomationXML\ObjectXML\ReportBuilder.xml"
	Set Fn_SISW_ReportBuilder_GetObject = Fn_SISW_Setup_GetObjectFromXML(sObjectXMLPath, sObjectName)
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name	:	Fn_RB_ReportBuilderTreeOperations
'@@
'@@    Description		:	Function Used to perform operations on Report Builder tree
'@@
'@@    Parameters		:	1. StrAction: Action to be performed
'@@    Parameters		:	2. StrNodeName: Action to be performed
'@@    Parameters		:	3. StrMenu: Action to be performed
'@@
'@@    Return Value		: 	True Or False
'@@
'@@    Pre-requisite	:	Report Builder perspective should be set							
'@@
'@@    Examples			:	Call Fn_RB_ReportBuilderTreeOperations(StrAction,StrNodeName,StrMenu)
'@@
'@@	   History			:	
'@@				Developer Name				Date			Rev. No.	Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe			20-Apr-2012			1.0			Created
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_RB_ReportBuilderTreeOperations(StrAction,StrNodeName,StrMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_RB_ReportBuilderTreeOperations"
	Dim objRBApplet
	
	If JavaWindow("ReportBuilderWindow").JavaApplet("ReportBuilderApplet").Exist Then
		Set objRBApplet = JavaWindow("ReportBuilderWindow").JavaApplet("ReportBuilderApplet")
	Else
		Set objRBApplet = JavaWindow("ReportBuilderWindow").JavaWindow("ReportBuilderWindow")
	End If
	
	Fn_RB_ReportBuilderTreeOperations = False
	Select Case StrAction
		'----------------------------------------------------------------------- For selecting single node -------------------------------------------------------------------------
		Case "Select"
			Fn_RB_ReportBuilderTreeOperations = Fn_JavaTree_Select("Fn_RB_ReportBuilderTreeOperations", objRBApplet, "NavTree",StrNodeName)
			If Fn_RB_ReportBuilderTreeOperations <> False Then
				Call Fn_ReadyStatusSync(1)
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Selected Node [" + StrNodeName + "] of NavTree")
			End If
		' - - - - - - - - - - Expand Node
		Case "Expand"
			Fn_RB_ReportBuilderTreeOperations = Fn_UI_JavaTree_Expand("Fn_RB_ReportBuilderTreeOperations", objRBApplet, "NavTree",StrNodeName)
			If Fn_RB_ReportBuilderTreeOperations <> False Then
				Call Fn_ReadyStatusSync(1)
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Expanded Node [" + StrNodeName + "] of NavTree")
			End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case Else
						Fn_RB_ReportBuilderTreeOperations = False
	End Select

	If Fn_RB_ReportBuilderTreeOperations <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), StrAction &" Sucessfully completed on Node [" + StrNodeName + "] of JavaTree of function Fn_RB_ReportBuilderTreeOperations")
	End If
	Set objRBApplet = nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name	:	Fn_RB_ReportDataOperations
'@@
'@@    Description		:	Function Used to perform operations on Report Builder tree
'@@
'@@    Parameters		:	1. StrAction: Action to be performed
'@@    Parameters		:	2. dicReportData - Define dictionary object
'@@
'@@    Return Value		: 	True Or False
'@@
'@@    Pre-requisite	:	Report Builder perspective should be set							
'@@
'@@    Examples			:	Dim dicReportData
'@@							Set dicReportData = CreateObject( "Scripting.Dictionary" )
'@@							With dicReportData  
'@@										 .Add "TreeNode","Reports Home:Teamcenter Reports:Admin - Items By Status"
'@@										 .Add "Name","kou"
'@@										 .Add "Description","kou"
'@@										 .Add "Source","Office Template"
'@@										 .Add "Query Source","Change Item Revision..."
'@@										 .Add "Report Format","PLMXML"
'@@										 .Add "Style-sheet Type","Teamcenter"
'@@										 .Add "Query Source","UserBasedProjects"
'@@										 .Add "Closure Rule","tcm_Exports"
'@@										 .Add "Property Set","ExportActivities"
'@@										 .Add "Class","ItemRevision"
'@@										 .Add "Transfer Mode","CRF_ECO_Details_Report"
'@@										 .Add "AddDefinedStyle-sheets","AllocatedTimeReport.xsl~AllocatedTimeReport.xsl"
'@@										 .Add "DeleteDefinedStyle-sheets","AllocatedTimeReport.xsl"
'@@							End with
'@@							msgbox Fn_RB_ReportDataOperations("ClearAndModify", dicReportData)
'@@							msgbox Fn_RB_ReportDataOperations("Modify", dicReportData)
'@@
'@@	   History			:	
'@@				Developer Name				Date			Rev. No.	Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe			23-Apr-2012			1.0			Created
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_RB_ReportDataOperations(sAction, dicReportData)
	GBL_FAILED_FUNCTION_NAME="Fn_RB_ReportDataOperations"
	Dim objRBApplet, iCnt, aReportData, aData
	Dim DictItems, DictKeys, iCount
	Dim objSelectType, objIntNoOfObjects, bFound
	If JavaWindow("ReportBuilderWindow").JavaApplet("ReportBuilderApplet").Exist Then
		Set objRBApplet = JavaWindow("ReportBuilderWindow").JavaApplet("ReportBuilderApplet")
	Else
		Set objRBApplet = JavaWindow("ReportBuilderWindow").JavaWindow("ReportBuilderWindow")
	End If
	Fn_RB_ReportDataOperations = False
	' select tree node from Report Builder tree
	If dicReportData("TreeNode") <> "" Then
		bFound = Fn_RB_ReportBuilderTreeOperations("Select",dicReportData("TreeNode"),"")
		If bFound = False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Fn_RB_ReportDataOperations : failed to select  [ " & dicReportData("TreeNode") & " ].")   	
			Exit function
		End If
	End If

	If trim(objRBApplet.JavaTab("RBTabbedPane").GetROProperty("value")) <> "Report Data" then
		objRBApplet.JavaTab("RBTabbedPane").Select "Report Data"
	End IF

	Select Case sAction
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Modify","ClearAndModify"
			If sAction = "ClearAndModify" Then
				Call Fn_Button_Click("Fn_RB_ReportDataOperations", objRBApplet, "Clear")
			End If

			' split report data with ~
                        DictItems = dicReportData.Items
			DictKeys = dicReportData.Keys
			For iCnt = 0 to dicReportData.Count - 1
				Select Case trim(DictKeys(iCnt))
					Case "TreeNode"
						' Do nothing
						'2 Add / remove / Move from list
					Case "AddDefinedStyle-sheets", "DeleteDefinedStyle-sheets"
						aData = split(DictItems(iCnt),"~")
						For iCount = 0 to UBound(aData)
							If iCount = 0 Then
								bFound = Fn_List_Select( "Fn_RB_ReportDataOperations",objRBApplet,"DefinedStylesheet", aData(iCount)) 
							Else
								bFound = Fn_UI_JavaList_ExtendSelect( "Fn_RB_ReportDataOperations",objRBApplet,"DefinedStylesheet", aData(iCount)) 
							End If
							If bFound = False Then
								' Can not set from list
								Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Fn_RB_ReportDataOperations : failed to set value for [ " & aData(iCount) & " ].")
								Exit function
							End If
						Next
						wait 3
						Select Case sAction
							Case "AddDefinedStyle-sheets"
								
								Call Fn_Button_Click("Fn_RB_ReportDataOperations",objRBApplet, "AddStylesheet")
							Case "DeleteDefinedStyle-sheets"
								Call Fn_Button_Click("Fn_RB_ReportDataOperations",objRBApplet, "DeleteStylesheet")
						End Select
					' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 	
					Case "RemoveSelectedStyle-sheets", "DeleteSelectedStyle-sheets"
						aData = split(DictItems(iCnt),"~")
						For iCount = 0 to UBound(aData)
							If iCount = 0 Then
								bFound = Fn_List_Select( "Fn_RB_ReportDataOperations",objRBApplet,"SelectedStylesheets", aData(iCount)) 
							Else
								bFound = Fn_UI_JavaList_ExtendSelect( "Fn_RB_ReportDataOperations",objRBApplet,"SelectedStylesheets", aData(iCount)) 
							End If
							If bFound = False Then
								' Can not set from list
								Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Fn_RB_ReportDataOperations : failed to set value for [ " & aData(iCount) & " ].")
								Exit function
							End If
						Next
						wait 3
						Select Case sAction
							Case "RemoveSelectedStyle-sheets"
								Call Fn_Button_Click("Fn_RB_ReportDataOperations",objRBApplet, "RemoveStylesheet")
							Case "DeleteSelectedStyle-sheets"
								Call Fn_Button_Click("Fn_RB_ReportDataOperations",objRBApplet, "DeleteStylesheet")
						End Select
					' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
					Case Else
						objRBApplet.JavaStaticText("DataLabel").SetTOProperty "label", trim(DictKeys(iCnt)) & ":"
						If objRBApplet.JavaStaticText("DataLabel").Exist(5) Then
							Select Case trim(DictKeys(iCnt))
								' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
								' Editbox
								Case "Name", "Description", "Process", "Output"
									If objRBApplet.JavaEdit("DataFieldEdit").Exist(5) Then
										objRBApplet.JavaEdit("DataFieldEdit").Set DictItems(iCnt)
									Else
										' edit box is not present
										Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Fn_RB_ReportDataOperations : failed to set [ " &  trim(DictKeys(iCnt)) & ": = "& DictKeys(iCnt) & " ].")
										Exit function
									End If
								' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
								' Combobox / List
								Case "Source", "Report Format", "Style-sheet Type"
									If Fn_List_Select("Fn_RB_ReportDataOperations",objRBApplet,"DataFieldList", DictItems(iCnt)) Then
										' do nothing
										wait 3
									Else
										' Can not set from list
										Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Fn_RB_ReportDataOperations : failed to set value for [ " & trim(DictKeys(iCnt)) & ": = "& DictKeys(iCnt) & " ].")
										Exit function
									End If
								' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
								' StaticText Combobox
								Case "Query Source", "Closure Rule", "Property Set", "Class", "Transfer Mode"
									bFound = False
									Call Fn_Button_Click("Fn_RB_ReportDataOperations", objRBApplet, "DataFieldDropDownButton")
									wait 5
									Set objSelectType=description.Create()
									objSelectType("Class Name").value = "JavaStaticText"					
									Set  objIntNoOfObjects = objRBApplet.ChildObjects(objSelectType)
									For  iCount = 0 to objIntNoOfObjects.count-1
										   If objIntNoOfObjects(iCount).getROProperty("label") = DictItems(iCnt) Then
												objIntNoOfObjects(iCount).Click 2,2
												bFound = TRUE
												Exit for
										   End If
									Next
									If  bFound = False Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Fn_RB_ReportDataOperations : failed to set [ " & trim(DictKeys(iCnt)) & ": = "& DictKeys(iCnt) & " ].")   	
										Exit function
									End If
								' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
							End Select
							Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Fn_RB_ReportDataOperations : Successfully set [ " & trim(DictKeys(iCnt)) & ": = "& DictKeys(iCnt) & " ].")   
						End If
				End Select
			Next
			' click on Modify button
			Call Fn_Button_Click("Fn_RB_ReportDataOperations", objRBApplet, "Modify")
			Fn_RB_ReportDataOperations = True
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Fn_RB_ReportDataOperations : Invalid action [ " & sAction & " ].")   
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	End Select
	If Fn_RB_ReportDataOperations <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Fn_RB_ReportDataOperations : executed successfully with action [ " & sAction & " ].")   	
	End If
	Set objRBApplet = Nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name	:	Fn_RB_CreateReportDefinitionTemplate
'@@
'@@    Description		:	Function used to create Report
'@@
'@@    Parameters		:	1. StrAction: Action to be performed
'@@    Parameters		:	2. dicReportData - Define dictionary object
'@@
'@@    Return Value		: 	Report ID Or False
'@@
'@@    Pre-requisite	:	Report Builder perspective should be set							
'@@
'@@    Examples			:	Dim dicReportData
'@@							Set dicReportData = CreateObject( "Scripting.Dictionary" )
'@@							With dicReportData  
'@@										 .Add "TreeNode","Reports Home:Teamcenter Reports:Admin - Items By Status"
'@@										 .Add "ReportType","Item report" / "Summary report"
'@@										 .Add "Report ID",""
'@@										 .Add "Name","kou"
'@@										 .Add "Description","kou"
'@@										 .Add "Source","Office Template"
'@@										 .Add "Query Source","Change Item Revision..."
'@@										 .Add "Report Format","PLMXML"
'@@										 .Add "Query Source","UserBasedProjects"
'@@										 .Add "Closure Rule","tcm_Exports"
'@@										 .Add "Property Set","ExportActivities"
'@@										 .Add "Class","ItemRevision"
'@@										 .Add "Transfer Mode","CRF_ECO_Details_Report"
'@@										 .Add "AddDefinedStyle-sheets","AllocatedTimeReport.xsl~AllocatedTimeReport.xsl"
'@@										 .Add "DeleteDefinedStyle-sheets","AllocatedTimeReport.xsl"
'@@							End with
'@@							msgbox Fn_RB_CreateReportDefinitionTemplate("Create", dicReportData)
'@@
'@@	   History			:	
'@@				Developer Name				Date			Rev. No.	Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe			24-Apr-2012			1.0			Created
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_RB_CreateReportDefinitionTemplate(sAction, dicReportData)
	GBL_FAILED_FUNCTION_NAME="Fn_RB_CreateReportDefinitionTemplate"
	Dim objCreateRDTDialog, iCnt, aReportData, aData,slable
	Dim DictItems, DictKeys, iCount
	Dim objSelectType, objIntNoOfObjects, bFound
'	Set objCreateRDTDialog = JavaWindow("ReportBuilderWindow").JavaWindow("RBWindow").JavaDialog("CreateReportDefinition")

' Modified Dialog hierarchy  : Modified by : Harshal Tanpure : Build : Teamcenter 10 (20120808.00) : 27-August-2012
	If JavaWindow("ReportBuilderWindow").JavaApplet("ReportBuilderApplet").Exist Then
		Set objCreateRDTDialog = JavaWindow("ReportBuilderWindow").JavaApplet("ReportBuilderApplet").JavaDialog("CreateReportDefinition")
	Else
		Set objCreateRDTDialog = JavaWindow("ReportBuilderWindow").JavaWindow("ReportBuilderWindow").JavaDialog("CreateReportDefinition")
	End If

	'Set objCreateRDTDialog = JavaWindow("ReportBuilderWindow").JavaApplet("ReportBuilderApplet").JavaDialog("CreateReportDefinition")

	Fn_RB_CreateReportDefinitionTemplate = False
	' select tree node from Report Builder tree
	If Fn_UI_ObjectExist("Fn_RB_CreateReportDefinitionTemplate",objCreateRDTDialog) = False Then
		If dicReportData("TreeNode") <> "" Then
			bFound = Fn_RB_ReportBuilderTreeOperations("Select",dicReportData("TreeNode"),"")
			If bFound = False Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Fn_RB_CreateReportDefinitionTemplate : failed to select  [ " & dicReportData("TreeNode") & " ].")   	
		 		Exit function
			End If
		End If

		Call Fn_MenuOperation("Select","File:Create Report")
		If Fn_UI_ObjectExist("Fn_RB_CreateReportDefinitionTemplate",objCreateRDTDialog) = False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Fn_RB_CreateReportDefinitionTemplate : failed to open  [ Create Report Definition Template ] window.")   	
			Exit function
		End IF
	End If

	Select Case sAction
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Create"

			' split report data with ~
			If dicReportData("ReportType") <> "" Then
				If Fn_UI_ListItemExist("Fn_RB_CreateReportDefinitionTemplate",objCreateRDTDialog,"SelectReportType",dicReportData("ReportType")) Then
					Call Fn_List_Select("Fn_RB_CreateReportDefinitionTemplate",objCreateRDTDialog,"SelectReportType",dicReportData("ReportType"))
				Else
					Exit function
				End If
			End If
			' clicking on next button
			Call Fn_Button_Click("Fn_RB_CreateReportDefinitionTemplate",objCreateRDTDialog, "Next")

			DictItems = dicReportData.Items
			DictKeys = dicReportData.Keys
			For iCnt = 0 to dicReportData.Count - 1
							slable = trim(DictKeys(iCnt))&":"
                            ' objCreateRDTDialog.JavaStaticText("DataLabel").SetTOProperty "lable", slable
                             JavaWindow("ReportBuilderWindow").JavaWindow("ReportBuilderWindow").JavaDialog("CreateReportDefinition").JavaEdit("DataFieldEdit").SetToProperty "attached text", slable

                             wait 2
				Select Case trim(DictKeys(iCnt))
					' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
					Case "Report ID"
						If objCreateRDTDialog.JavaEdit("DataFieldEdit").Exist(5) Then
							If DictItems(iCnt) <> "" Then
								objCreateRDTDialog.JavaEdit("DataFieldEdit").Set DictItems(iCnt)
							Else
								Call Fn_Button_Click("Fn_RB_CreateReportDefinitionTemplate", objCreateRDTDialog, "Assign")
								Fn_RB_CreateReportDefinitionTemplate = objCreateRDTDialog.JavaEdit("DataFieldEdit").getROProperty("value") 
							End If
						Else
							' edit box is not present
							Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Fn_RB_CreateReportDefinitionTemplate : failed to set [ " &  trim(DictKeys(iCnt)) & ": = "& DictKeys(iCnt) & " ].")
							Exit function
						End If
					' Editbox
					Case "Name", "Description"
						If JavaWindow("ReportBuilderWindow").JavaWindow("ReportBuilderWindow").JavaDialog("CreateReportDefinition").JavaEdit("DataFieldEdit").Exist(5) Then
							JavaWindow("ReportBuilderWindow").JavaWindow("ReportBuilderWindow").JavaDialog("CreateReportDefinition").JavaEdit("DataFieldEdit").Set DictItems(iCnt)
						Else
							' edit box is not present
							Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Fn_RB_CreateReportDefinitionTemplate : failed to set [ " &  trim(DictKeys(iCnt)) & ": = "& DictKeys(iCnt) & " ].")
							Fn_RB_CreateReportDefinitionTemplate = False
							Exit function
						End If
					' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
					' Combobox / List
					Case "Source", "Report Format", "Style-sheet Type"
					
						If Fn_List_Select("Fn_RB_CreateReportDefinitionTemplate",objCreateRDTDialog,"DataFieldList", DictItems(iCnt)) Then
							' do nothing
							wait 3
						Else
							' Can not set from list
							Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Fn_RB_CreateReportDefinitionTemplate : failed to set value for [ " & trim(DictKeys(iCnt)) & ": = "& DictKeys(iCnt) & " ].")
							Fn_RB_CreateReportDefinitionTemplate = False
							Exit function
						End If
					' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
					' StaticText Combobox
					Case "Query Source", "Closure Rule", "Property Set", "Class", "Transfer Mode"
							bFound = False
							objCreateRDTDialog.JavaStaticText("DataLabel").SetTOProperty "lable", slable
							If objCreateRDTDialog.JavaEdit("DataFieldEdit").Exist(5) Then
								Call Fn_Button_Click("Fn_RB_CreateReportDefinitionTemplate", JavaWindow("ReportBuilderWindow").JavaWindow("ReportBuilderWindow").JavaDialog("CreateReportDefinition"), "DataFieldDropDownButton")
							wait 10
							Set objSelectType=description.Create()
							objSelectType("Class Name").value = "JavaStaticText"					
							Set  objIntNoOfObjects = objCreateRDTDialog.ChildObjects(objSelectType)
							For  iCount = 0 to objIntNoOfObjects.count-1
								   If objIntNoOfObjects(iCount).getROProperty("label") = DictItems(iCnt) Then
										objIntNoOfObjects(iCount).Click 2,2
										bFound = TRUE
										Exit for
								   End If
							Next
						End If
						If  bFound = False Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Fn_RB_CreateReportDefinitionTemplate : failed to set [ " & trim(DictKeys(iCnt)) & ": = "& DictKeys(iCnt) & " ].")   	
							Fn_RB_CreateReportDefinitionTemplate = False
							Exit function
						End If
					' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				End Select
			Next

			If dicReportData("AddDefinedStyle-sheets") <> ""  OR dicReportData("RemoveSelectedStyle-sheets") <> "" Then
				' clicking on next button
				Call Fn_Button_Click("Fn_RB_CreateReportDefinitionTemplate",objCreateRDTDialog, "Next")
				If dicReportData("AddDefinedStyle-sheets") <> "" Then
				
					aData = split(dicReportData("AddDefinedStyle-sheets"),"~")
					For iCount = 0 to UBound(aData)
						If iCount = 0 Then
							bFound = Fn_List_Select( "Fn_RB_CreateReportDefinitionTemplate",objCreateRDTDialog,"DefinedStylesheet", aData(iCount)) 
						Else
							bFound = Fn_UI_JavaList_ExtendSelect( "Fn_RB_CreateReportDefinitionTemplate",objCreateRDTDialog,"DefinedStylesheet", aData(iCount)) 
						End If
						If bFound = False Then
							' Can not set from list
							Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Fn_RB_CreateReportDefinitionTemplate : failed to set value for [ " & aData(iCount) & " ].")
							Exit function
						End If
					Next
					wait 3
					Call Fn_Button_Click("Fn_RB_CreateReportDefinitionTemplate",objCreateRDTDialog, "AddStylesheet")
				End If
				If dicReportData("RemoveSelectedStyle-sheets") <> "" Then
					aData = split(DictItems(iCnt),"~")
					For iCount = 0 to UBound(aData)
						If iCount = 0 Then
							bFound = Fn_List_Select( "Fn_RB_CreateReportDefinitionTemplate",objCreateRDTDialog,"SelectedStylesheets", aData(iCount)) 
						Else
							bFound = Fn_UI_JavaList_ExtendSelect( "Fn_RB_CreateReportDefinitionTemplate",objCreateRDTDialog,"SelectedStylesheets", aData(iCount)) 
						End If
						If bFound = False Then
							' Can not set from list
							Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Fn_RB_CreateReportDefinitionTemplate : failed to set value for [ " & aData(iCount) & " ].")
							Exit function
						End If
					Next
					wait 3
					Call Fn_Button_Click("Fn_RB_CreateReportDefinitionTemplate",objCreateRDTDialog, "RemoveStylesheet")
				End If
			End If
			' click on Modify button
			Call Fn_Button_Click("Fn_RB_CreateReportDefinitionTemplate", objCreateRDTDialog, "Finish")
			If Fn_UI_ObjectExist("Fn_RB_CreateReportDefinitionTemplate",objCreateRDTDialog) = False Then
				Call Fn_Button_Click("Fn_RB_CreateReportDefinitionTemplate", objCreateRDTDialog, "Close")
			End IF
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Fn_RB_CreateReportDefinitionTemplate : Invalid action [ " & sAction & " ].")   
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	End Select
	If Fn_RB_CreateReportDefinitionTemplate <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Fn_RB_CreateReportDefinitionTemplate : executed successfully with action [ " & sAction & " ].")   	
	End If
	Set objCreateRDTDialog = Nothing
End Function
