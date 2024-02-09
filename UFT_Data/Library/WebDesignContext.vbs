Option Explicit
				'Function List
'************************************************************************************************************************************************************************************************************
'000. Fn_SISW_Web_DC_GetObject(sObjectName)
'001. Fn_Web_DC_ContextDefinitionOperations()
'002. Fn_Web_DC_ConfigureFiltersOperation()
'003. Fn_SISW_Web_DC_CreateNewZone()
'004. Fn_SISW_Web_DC_ResultsOperations()
'005. Fn_SISW_Web_DC_ConfigureWorkPartContextOperations()
'006. Fn_SISW_Web_DC_SaveStructureContextObject()
'007. Fn_SISW_Web_DC_FormAttributesTableOperation()
'************************************************************************************************************************************************************************************************************
'****************************************    Function to get Object hierarchy ***************************************
'
''Function Name		 	:	Fn_SISW_Web_DC_GetObject
'
''Description		    :  	Function to get Object hierarchy

''Parameters		    :	1. sObjectName : Object Handle name
								
''Return Value		    :  	Object \ Nothing
'
''Examples		     	:	Fn_SISW_Web_DC_GetObject("Teamcenter Web - Design")

'History:
'	Developer Name			Date			Rev. No.		Reviewer		Changes Done	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Sachin Joshi		 26-Sept-2012		1.0				
'-----------------------------------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------------------------------
'	Anurag Khera		 26-Sept-2013		2.0								Externalized object hierarchies
'-----------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_Web_DC_GetObject(sObjectName)
	Dim sObjectXMLPath
	sObjectXMLPath = Fn_GetEnvValue("User", "AutomationDir") & "\TestData\AutomationXML\ObjectXML\WebDesignContext.xml"
	Set Fn_SISW_Web_DC_GetObject = Fn_SISW_Setup_GetObjectFromXML(sObjectXMLPath, sObjectName)
End Function
'************************************************************************************************************************************************************************************************************
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name			:	Fn_Web_DC_ContextDefinitionOperations
'@@
'@@    Description				:	Function Used to perform operations on Context Definition in Design Context
'@@
'@@    Parameters			   	:	1. sAction : Action Name
'@@									2. dicWebDCContextDef : Dictionary object
'@@
'@@    Return Value		   	   	: 	True Or False
'@@
'@@    Pre-requisite			:	Should Be Log in Web Client And Design Context perspective should be open
'@@
'@@    Examples					: : 
'@@    								Dim dicWebDCContextDef
'@@    								Set dicWebDCContextDef = CreateObject("Scripting.Dictionary")
'@@    								With dicWebDCContextDef
'@@    									.Add "bClear", True
'@@    									.Add "sProductItem", "000030-dc[ Product ]"
'@@    									.Add "sProductContext", "000030/A;1-dc[ Product ]"
'@@    									.Add "sSearchType", "Processes"
'@@    									.Add "sSearchCriteriaText",  "000031~000032"
'@@    									.Add "sAddSearchCriteria", "A108-Automatic Transmission[ ]~A109-Manual Transmission[ ]$A112-CD Player[ ]~A113-DVD Player[ ]"
'@@    									.Add "sRemoveSearchCriteria", ""
'@@    									.Add "bUpdate", True
'@@    								End With
'@@    								Call  Fn_Web_DC_ContextDefinitionOperations("SetData", dicWebDCContextDef)
'@@	   History:				Developer Name				Date				Rev. No.		Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@							Koustubh Watwe				10-Jan-2012			1.0				Created.
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@							Koustubh Watwe				22-Feb-2012			2.0				Modified function.
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@    								With dicWebDCContextDef
'@@    									.Add "sProductItem", "000030-dc[ Product ]"
'@@    									.Add "sProductContext", "000030/A;1-dc[ Product ]"
'@@    									.Add "sSearchType", "Processes"
'@@    									.Add "sSearchCriteriaText", "A108-Automatic Transmission[ ]~A109-Manual Transmission[ ]~A112-CD Player[ ]~A113-DVD Player[ ]"
'@@    								End With
'@@    								Call Fn_Web_DC_ContextDefinitionOperations("Verify", dicWebDCContextDef)
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@	   	History:				
'@@    	Developer Name				Date				Rev. No.		Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@		Koustubh Watwe				10-Jan-2012			1.0				Created.
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@		Koustubh Watwe				22-Feb-2012			2.0				Modified function.
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@		Koustubh Watwe				28-Sept-2012		2.0				Added case Verify
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_Web_DC_ContextDefinitionOperations(sAction, dicWebDCContextDef)
	GBL_FAILED_FUNCTION_NAME="Fn_Web_DC_ContextDefinitionOperations"
	Dim objWebPage, iCnt, aSearchCri, objTable, aAddSrchCri, iSrchCri, aSearchRes, iCount, objCheckBox
	Fn_Web_DC_ContextDefinitionOperations = False
	Set objWebPage = Browser("Teamcenter Web - Design").Page("Teamcenter Web - Design")
	Select Case sAction
		Case "SetData"
				If dicWebDCContextDef("bClear") <> "" Then
					If cBool(dicWebDCContextDef("bClear")) Then
						' click on button
						Call Fn_Web_UI_Button_Click("Fn_Web_DC_ContextDefinitionOperations",objWebPage,  "Clear")
					End If
				End If

				If dicWebDCContextDef("sProductItem") <> "" Then
                    Call Fn_Web_UI_List_Select("Fn_Web_DC_ContextDefinitionOperations", objWebPage, "ProductItems",dicWebDCContextDef("sProductItem"))
				End If

				If dicWebDCContextDef("sProductContext") <> "" Then
                    Call Fn_Web_UI_List_Select("Fn_Web_DC_ContextDefinitionOperations", objWebPage, "ProductContext",dicWebDCContextDef("sProductContext"))
				End If

				If dicWebDCContextDef("sSearchType") <> "" Then
                    Call Fn_Web_UI_List_Select("Fn_Web_DC_ContextDefinitionOperations", objWebPage, "SearchType",dicWebDCContextDef("sSearchType"))
				End If

				If dicWebDCContextDef("sSearchCriteriaText") <> "" Then
						aSearchCri = split(dicWebDCContextDef("sSearchCriteriaText"),"~")
						For iCnt = 0 to UBound(aSearchCri)
'							Call Fn_Web_UI_WebEdit_Set("Fn_Web_DC_ContextDefinitionOperations",objWebPage, "SearchCriteriaEditbox", aSearchCri(iCnt))
							objWebPage.WebEdit("SearchCriteriaEditbox").Set aSearchCri(iCnt)
							objWebPage.WebEdit("SearchCriteriaEditbox").Click 0, 0,micLeftBtn
							Call Fn_KeyBoardOperation("SendKeys", " ")
							Call Fn_KeyBoardOperation("SendKeys", "{BACKSPACE}")
							Call Fn_Web_UI_Button_Click("Fn_Web_DC_ContextDefinitionOperations",objWebPage,  "AddSearchCriteria")
							Wait 2
							If inStr(aSearchCri(iCnt),"*") > 0 Then
								' select from search result table.
								If dicWebDCContextDef("sAddSearchCriteria") <> "" Then
									Set objTable = objWebPage.WebTable("ConextSearchResultTable")
									Select Case dicWebDCContextDef("sAddSearchCriteria")
										' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
										Case "SelectAll"
											Call Fn_Web_UI_Button_Click("Fn_Web_DC_ContextDefinitionOperations",objWebPage.WebTable("ButtonTable"),  "DeselectAll")
											Call Fn_Web_UI_Button_Click("Fn_Web_DC_ContextDefinitionOperations",objWebPage.WebTable("ButtonTable"),  "SelectAll")
										' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
										Case "DeselectAll"
											Call Fn_Web_UI_Button_Click("Fn_Web_DC_ContextDefinitionOperations",objWebPage.WebTable("ButtonTable"),  "DeselectAll")
										' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
										Case Else
											Call Fn_Web_UI_Button_Click("Fn_Web_DC_ContextDefinitionOperations",objWebPage.WebTable("ButtonTable"),  "DeselectAll")
											aAddSrchCri = split(dicWebDCContextDef("sAddSearchCriteria"), "$")
											aSearchRes = Split(aAddSrchCri(iCount),"~")
											For iSrchCri=0 to UBound(aSearchRes)
												For iCount = 0 to cInt(objTable.RowCount)
													If trim(cstr(objTable.GetCellData(iCount,2))) = trim(cstr(aSearchRes(iSrchCri))) Then
														Set objCheckBox = objTable.ChildItem(iCount,1,"WebCheckbox",0)
														If trim(typename(objCheckBox)) <> "" Then
															objCheckBox.set "ON"
														Else
															Exit Function
														End If
													End If
												Next' inner loop
											Next ' outer loop
										' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
									End Select
									' Clicking on Add button
									Call Fn_Web_UI_Button_Click("Fn_Web_DC_ContextDefinitionOperations",objWebPage.WebTable("ButtonTable"),  "Add")
								End IF
							End If
						Next
				End If

				If dicWebDCContextDef("sRemoveSearchCriteria") <> "" Then
						aSearchCri = split(dicWebDCContextDef("sRemoveSearchCriteria"),"~")
						For iCnt = 0 to UBound(aSearchCri)
							Call Fn_Web_UI_List_Select("Fn_Web_DC_ContextDefinitionOperations",objWebPage, "SearchCriteriaList", aSearchCri(iCnt))
							Call Fn_Web_UI_Button_Click("Fn_Web_DC_ContextDefinitionOperations",objWebPage,  "RemovePartResults")
						Next
				End If

				If dicWebDCContextDef("bUpdate") <> "" Then
					If cBool(dicWebDCContextDef("bUpadte")) Then
						' click on button
						Call Fn_Web_UI_Button_Click("Fn_Web_DC_ContextDefinitionOperations",objWebPage,  "Update")
					End If
				End If
				Fn_Web_DC_ContextDefinitionOperations = True
        ' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "Verify"
			If dicWebDCContextDef("sProductItem") <> "" Then
				If Fn_SISW_WebUI_WebListItemExist("Fn_Web_DC_ContextDefinitionOperations", objWebPage, "ProductItems",dicWebDCContextDef("sProductItem")) = False Then
					Exit Function
				End IF
			End If

			If dicWebDCContextDef("sProductContext") <> "" Then
				If Fn_SISW_WebUI_WebListItemExist("Fn_Web_DC_ContextDefinitionOperations", objWebPage, "ProductContext",dicWebDCContextDef("sProductContext")) = False Then
					Exit Function
				End IF
			End If

			If dicWebDCContextDef("sSearchType") <> "" Then
				If Fn_SISW_WebUI_WebListItemExist("Fn_Web_DC_ContextDefinitionOperations", objWebPage, "SearchType",dicWebDCContextDef("sSearchType")) = False Then
					Exit Function
				End IF
			End If
			If dicWebDCContextDef("sSearchCriteriaText") <> "" Then
				aSearchCri = split(dicWebDCContextDef("sSearchCriteriaText"),"~")
				For iCnt = 0 to UBound(aSearchCri)
					If Fn_SISW_WebUI_WebListItemExist("Fn_Web_DC_ContextDefinitionOperations", objWebPage, "SearchCriteriaList", aSearchCri(iCnt)) = False Then
						Exit Function
					End IF
				Next
			End If
			Fn_Web_DC_ContextDefinitionOperations = True
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "VerifyValues"
			If dicWebDCContextDef("sProductItem") <> ""  Then
				If objWebPage.WebList("ProductItems").GetROProperty("value") <> dicWebDCContextDef("sProductItem")  Then
					Exit Function
				End IF
			End If

			If dicWebDCContextDef("sProductContext") <> "" Then
				If objWebPage.WebList("ProductContext").GetROProperty("value") <> dicWebDCContextDef("sProductContext") Then
					Exit Function
				End IF
			End If

			If dicWebDCContextDef("sSearchType") <> "" Then
				If objWebPage.WebList("SearchType").GetROProperty("value") <> dicWebDCContextDef("sSearchType") Then
					Exit Function
				End IF
			End If

			If dicWebDCContextDef("SearchCriteriaList") <> "" Then
				If objWebPage.WebList("SearchCriteriaList").GetROProperty("value") <> dicWebDCContextDef("sAddSearchCriteria") Then
					Exit Function
				End IF
			End If

			If dicWebDCContextDef("sSearchCriteriaText") <> ""  Then
				If objWebPage.WebEdit("SearchCriteriaEditbox").GetROProperty("value") <> dicWebDCContextDef("sSearchCriteriaText")  Then
					Exit Function
				End IF
			End If

			Fn_Web_DC_ContextDefinitionOperations = True
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_Web_DC_ContextDefinitionOperations ] Invalid case [ " & sAction & " ].")
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
	End Select
	If Fn_Web_DC_ContextDefinitionOperations Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [ Fn_Web_DC_ContextDefinitionOperations ] executed successfuly with case [ " & sAction & " ].")
	End If
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name			:	Fn_Web_DC_ConfigureFiltersOperation
'@@
'@@    Description				:	Function Used to perform operations on Configure Filter in Design Context
'@@
'@@    Parameters			   	:	1. sAction : Action Name
'@@									2. dicWebDCConfigFilters : Dictionary Object
'@@
'@@    Return Value		   	   	: 	True Or False
'@@
'@@    Pre-requisite			:	Should Be Log in Web Client And Design Context perspective should be open
'@@
'@@    Examples					: 
'@@    								Dim dicWebDCConfigFilters
'@@    								Set dicWebDCConfigFilters = CreateObject("Scripting.Dictionary")
'@@    								With dicWebDCConfigFilters
'@@    										.Add "bClear", True
'@@    										.Add "ProximityFilter", "10"
'@@    										.Add "bTrueShapeFiltering", True
'@@    										.Add "bValidOverlaysOnly", False
'@@    										.Add "bAppendParts", True
'@@    										.Add "SavedQuery", "Item Revision..."
'@@    										.Add "SavedQuery_Criteria", "Name=workpart~Item ID=000060~Alternate Revision Type=Identifier"
'@@    										.Add "ZoneFilterID", "" 
'@@    										.Add "ZoneFilterOperator", ""
'@@    										.Add "RemoveZoneFilter", ""
'@@    										.Add "OccNotes_Field", "UG NAME~AIE_OCC_ID"
'@@    										.Add "OccNotes_Operator", "EQ~NE"
'@@    										.Add "OccNotes_Value", "ugnametext~asd"
'@@    										.Add "OccNotes_RemoveResults", ""
'@@    										.Add "RelationType", "Specifications" 
'@@    										.Add "ParentType", "Item"
'@@    										.Add "FormType", "EmailMaster"
'@@    										.Add "PropertyName", "Name"
'@@    										.Add "Operator", "EQ"
'@@    										.Add "SearchingValue" , "text"
'@@    										.Add "RemoveFormAttributes", ""
'@@    										.Add "bUpdate", True
'@@    								End With
'@@    								Call  Fn_Web_DC_ConfigureFiltersOperation("SetData", dicWebDCConfigFilters)
'@@
'@@									With dicWebDCConfigFilters
'@@								 			.Add "ProximityFilter","10.000000 "
'@@								 			.Add "TrueShapeFiltering","ON"
'@@											.Add "ValidOverlaysOnly","OFF"
'@@								 			.Add "Append Parts","OFF"
'@@								 			.Add "SavedQuery","Select Saved Query"
'@@									End With
'@@									bReturn=Fn_Web_DC_ConfigureFiltersOperation("VerifyData", dicWebDCConfigFilters)
'@@
'@@	   History:				Developer Name				Date				Rev. No.		Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@							Koustubh Watwe				10-May-2011			1.0				created
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@							Koustubh Watwe				27-Feb-2011			1.0				Added case "VerifyInResults"
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@							Sandeep N						25-Jul-2013			1.0				Added case "VerifyData"
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_Web_DC_ConfigureFiltersOperation(sAction, dicWebDCConfigFilters)
	GBL_FAILED_FUNCTION_NAME="Fn_Web_DC_ConfigureFiltersOperation"
	Dim objDCWebPage, aFields, aOperators, aValues, iCnt
	Dim iCount, iRowCnt, bFlag
	Dim aKey,aItem

	Set objDCWebPage = Browser("Teamcenter Web - Design").Page("Teamcenter Web - Design")
	Select Case sAction
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			Case "SetData"
						If dicWebDCConfigFilters("bClear") <> "" Then
							If cBool(dicWebDCConfigFilters("bClear") ) Then
								Call Fn_Web_UI_Button_Click("Fn_Web_DC_ConfigureFiltersOperation",objDCWebPage, "Clear")
							End If
							If Fn_SISW_UI_Object_Operations("Fn_Web_DC_ConfigureFiltersOperation", "Exist", Browser("TeamcenterWeb").Dialog("Dialog"), "") Then
								Call Fn_UI_WinButton_Click("Fn_Web_DC_ConfigureFiltersOperation", Browser("TeamcenterWeb").Dialog("Dialog"), "OK","","","")
							End If
						End If

						' setting Proximity Filter
						If dicWebDCConfigFilters("ProximityFilter") <> "" Then
							Call Fn_Web_UI_WebEdit_Set("Fn_Web_DC_ConfigureFiltersOperation", objDCWebPage, "ProximityFilter",dicWebDCConfigFilters("ProximityFilter"))
						End If

						' setting True Shape Filtering checkbox
						If dicWebDCConfigFilters("bTrueShapeFiltering") <> ""  Then
							If cBool(dicWebDCConfigFilters("bTrueShapeFiltering")) Then
								Call Fn_Web_UI_CheckBox_Set("Fn_Web_DC_ConfigureFiltersOperation",objDCWebPage, "TrueShapeFiltering","ON")
							Else
								Call Fn_Web_UI_CheckBox_Set("Fn_Web_DC_ConfigureFiltersOperation",objDCWebPage, "TrueShapeFiltering","OFF")
							End If
						End If

						'setting Valid Overlays Only check box
						If dicWebDCConfigFilters("bValidOverlaysOnly") <> ""  Then
							If cBool(dicWebDCConfigFilters("bValidOverlaysOnly")) Then
								Call Fn_Web_UI_CheckBox_Set("Fn_Web_DC_ConfigureFiltersOperation",objDCWebPage, "ValidOverlaysOnly","ON")
							Else
								Call Fn_Web_UI_CheckBox_Set("Fn_Web_DC_ConfigureFiltersOperation",objDCWebPage, "ValidOverlaysOnly","OFF")
							End If
						End If

						'setting Append Parts check box
						If dicWebDCConfigFilters("bAppendParts") <> ""  Then
							If cBool(dicWebDCConfigFilters("bAppendParts")) Then
								Call Fn_Web_UI_CheckBox_Set("Fn_Web_DC_ConfigureFiltersOperation",objDCWebPage, "Append Parts","ON")
							Else
								Call Fn_Web_UI_CheckBox_Set("Fn_Web_DC_ConfigureFiltersOperation",objDCWebPage, "Append Parts","OFF")
							End If
						End If

						' setting Saved Query
						If dicWebDCConfigFilters("SavedQuery") <> ""  Then
							Call Fn_Web_UI_List_Select("Fn_Web_DC_ConfigureFiltersOperation",objDCWebPage, "SavedQuery", dicWebDCConfigFilters("SavedQuery"))
							' clickinng on search image icon
							objDCWebPage.Image("search_savedQuery").Click 1,1,micLeftBtn
							Browser("Teamcenter Web - Design").Page("Teamcenter Web - Design").WebTable("SavedQueryTable").SetTOProperty "outerhtml", ".*" & replace(dicWebDCConfigFilters("SavedQuery"),".","") & ".*"

							If NOT(Fn_Web_UI_ObjectExist("", Browser("Teamcenter Web - Design").Page("Teamcenter Web - Design").WebTable("SavedQueryTable"))) Then
								' failed to find
								Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_Web_DC_ConfigureFiltersOperation ] Failed to Find [ " & dicWebDCConfigFilters("SavedQuery") & " ] dialog.") 
								Exit function
							End If

							aFields = split(dicWebDCConfigFilters("SavedQuery_Criteria"), "~")
							For iCnt = 0 to UBound(aFields)
								aValues = split(aFields(iCnt),"=")
								Browser("Teamcenter Web - Design").Page("Teamcenter Web - Design").WebTable("SavedQueryTable").WebElement("SavedQuery_FieldLabel").SetTOProperty "innertext", aValues(0) & ":"
								If Browser("Teamcenter Web - Design").Page("Teamcenter Web - Design").WebTable("SavedQueryTable").WebEdit("SavedQuery_ValueEditbox").Exist(5) Then
									Call Fn_Web_UI_WebEdit_Set("Fn_Web_DC_ConfigureFiltersOperation", Browser("Teamcenter Web - Design").Page("Teamcenter Web - Design").WebTable("SavedQueryTable"), "SavedQuery_ValueEditbox",aValues(1))
									Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Function [ Fn_Web_DC_ConfigureFiltersOperation ] Successfully set [ " & aValues(0) & " = " & aValues(1) & " ].") 										
								Else
									' no field is present
									Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_Web_DC_ConfigureFiltersOperation ] Failed to Find Field [ " & aValues(0) & " ].") 										
									Exit function
								End If
							Next

							' clicking on OK button
							Call Fn_Web_UI_Button_Click("Fn_Web_DC_ConfigureFiltersOperation", Browser("Teamcenter Web - Design").Page("Teamcenter Web - Design").WebTable("ButtonTable"),"OK")

						End If

						'setting Zone Filters
						If dicWebDCConfigFilters("ZoneFilterID") <> ""  Then
							If objDCWebPage.WebButton("ZoneFilter").Exist(5) Then
								Call Fn_Web_UI_Button_Click("Fn_Web_DC_ConfigureFiltersOperation",objDCWebPage, "ZoneFilter")
							End If
							'"ZoneResults"   "ZoneFilterOperatorID",    "ZoneFilterID"		
							aFields = split(dicWebDCConfigFilters("ZoneFilterID"),"~")
							aOperators  = split(dicWebDCConfigFilters("ZoneFilterOperator"),"~")

							For iCnt = 0 to UBound(aFields)
								' setting Occ Note ID
								Call Fn_Web_UI_List_Select("Fn_Web_DC_ConfigureFiltersOperation", objDCWebPage, "ZoneFilterID",aFields(iCnt))
								' setting Operator
								Call Fn_Web_UI_List_Select("Fn_Web_DC_ConfigureFiltersOperation", objDCWebPage, "ZoneFilterOperatorID",aOperators(iCnt))
								' setting value	
								Call Fn_Web_UI_Button_Click("Fn_Web_DC_ConfigureFiltersOperation",objDCWebPage, "AddZoneValueButton")
							Next
						End If
						
							If dicWebDCConfigFilters("RemoveZoneFilter") <> ""  Then
								If objDCWebPage.WebButton("ZoneFilter").Exist(5) Then
									Call Fn_Web_UI_Button_Click("Fn_Web_DC_ConfigureFiltersOperation",objDCWebPage, "ZoneResults")
								End If
								aValues = split(dicWebDCConfigFilters("RemoveZoneFilter"),"~")
								For iCnt = 0 to UBound(aValues)
									Call Fn_Web_UI_List_Select("Fn_Web_DC_ConfigureFiltersOperation", objDCWebPage, "ZoneResults",aValues(iCnt))
									Call Fn_Web_UI_Button_Click("Fn_Web_DC_ConfigureFiltersOperation",objDCWebPage, "RemoveZoneResults")
								Next	
							End If
						
						'setting Occurrence Notes
						If dicWebDCConfigFilters("OccNotes_Field") <> ""  Then
							If objDCWebPage.WebButton("OccurrenceNotes").Exist(5) Then
								Call Fn_Web_UI_Button_Click("Fn_Web_DC_ConfigureFiltersOperation",objDCWebPage, "OccurrenceNotes")
							End If
							aFields = split(dicWebDCConfigFilters("OccNotes_Field"),"~")
							aOperators  = split(dicWebDCConfigFilters("OccNotes_Operator"),"~")
							aValues = split(dicWebDCConfigFilters("OccNotes_Value"),"~")
							For iCnt = 0 to UBound(aFields)
								' setting Occ Note ID
								Call Fn_Web_UI_List_Select("Fn_Web_DC_ConfigureFiltersOperation", objDCWebPage, "OccurrenceNotesID",aFields(iCnt))
								' setting Operator
								Call Fn_Web_UI_List_Select("Fn_Web_DC_ConfigureFiltersOperation", objDCWebPage, "OccurrenceNotesOperatorID",aOperators(iCnt))
								' setting value
								Call Fn_Web_UI_WebEdit_Set("Fn_Web_DC_ConfigureFiltersOperation", objDCWebPage, "OccurrenceNoteCriteria",aValues(iCnt))
								' clicking on add button
								Call Fn_Web_UI_Button_Click("Fn_Web_DC_ConfigureFiltersOperation",objDCWebPage, "AddOccNoteValue")
							Next
						End If

						' removing Occ Note Criteria
						If dicWebDCConfigFilters("OccNotes_RemoveResults") <> ""  Then
							If objDCWebPage.WebButton("OccurrenceNotes").Exist(5) Then
								Call Fn_Web_UI_Button_Click("Fn_Web_DC_ConfigureFiltersOperation",objDCWebPage, "OccurrenceNotes")
							End If
							aValues = split(dicWebDCConfigFilters("OccNotes_RemoveResults"),"~")
							For iCnt = 0 to UBound(aValues)
								Call Fn_Web_UI_List_Select("Fn_Web_DC_ConfigureFiltersOperation", objDCWebPage, "OccurrenceNotesResults",aValues(iCnt))
								Call Fn_Web_UI_Button_Click("Fn_Web_DC_ConfigureFiltersOperation",objDCWebPage, "RemoveOccurrenceNotesResults")
							Next	
						End If
						'setting Form Attributes
						If dicWebDCConfigFilters("RelationType") <> ""  Then
							Call Fn_Web_UI_Button_Click("Fn_Web_DC_ConfigureFiltersOperation",objDCWebPage, "FormAttributes")
                            Call Fn_Web_UI_Button_Click("Fn_Web_DC_ConfigureFiltersOperation",objDCWebPage, "AddFormAttributes")

							If Fn_Web_UI_ObjectExist("Fn_Web_DC_ConfigureFiltersOperation", objDCWebPage.WebTable("AddFormAttributes")) Then

								If dicWebDCConfigFilters("RelationType") <> "" Then
									Call Fn_Web_UI_WebEdit_SetExt("Fn_Web_DC_ConfigureFiltersOperation","Set", objDCWebPage.WebTable("AddFormAttributes"), "RelationType",dicWebDCConfigFilters("RelationType"))
								End If

								If dicWebDCConfigFilters("ParentType") <> "" Then
									Call Fn_Web_UI_WebEdit_SetExt("Fn_Web_DC_ConfigureFiltersOperation", "Set",objDCWebPage.WebTable("AddFormAttributes"), "ParentType",dicWebDCConfigFilters("ParentType"))
								End If

								If dicWebDCConfigFilters("FormType") <> "" Then		' Modified code to set value in FormType List [TC11.2 Maintenance : Build(2015062400) : By Vivek Ahirrao]
									'[TC1121-2015111600-24_11_2015-VivekA-Maintenance] - Added by Jotiba T
									Call Fn_Web_UI_WebEdit_SetExt("Fn_Web_DC_ConfigureFiltersOperation","Set", objDCWebPage.WebTable("AddFormAttributes"), "FormType",dicWebDCConfigFilters("FormType"))
									'Call Fn_Web_UI_Button_Click("Fn_Web_DC_ConfigureFiltersOperation",objDCWebPage.WebTable("AddFormAttributes"), "WebButton")
									'objDCWebPage.WebTable("AddFormAttributes").WebElement("PropertyNameValue").SetTOProperty "innertext", dicWebDCConfigFilters("FormType")
								  	 
								  	'Call Fn_Web_UI_WebElement_Click("Fn_Web_DC_ConfigureFiltersOperation",objDCWebPage.WebTable("AddFormAttributes"), "PropertyNameValue", "","","")									
								End If

								If dicWebDCConfigFilters("PropertyName") <> "" Then
									Call Fn_Web_UI_WebEdit_SetExt("Fn_Web_DC_ConfigureFiltersOperation", "Set",objDCWebPage.WebTable("AddFormAttributes"), "PropertyName",dicWebDCConfigFilters("PropertyName"))
								End If

								If dicWebDCConfigFilters("Operator") <> "" Then
									Call Fn_Web_UI_WebEdit_SetExt("Fn_Web_DC_ConfigureFiltersOperation","Set", objDCWebPage.WebTable("AddFormAttributes"), "Operator",dicWebDCConfigFilters("Operator"))
								End If

								If dicWebDCConfigFilters("SearchingValue") <> "" Then
									Call Fn_Web_UI_WebEdit_Set("Fn_Web_DC_ConfigureFiltersOperation", objDCWebPage.WebTable("AddFormAttributes"), "SearchingValue",dicWebDCConfigFilters("SearchingValue"))
								End If

								' click on OK button
								Call Fn_Web_UI_Button_Click("Fn_Web_DC_ConfigureFiltersOperation", objDCWebPage.WebTable("ButtonTable"), "OK")
							Else 
								' "add form attribute" dialog is not present
								Exit function
							End If
						End If

						' remove selected form attributes from the list.
						If dicWebDCConfigFilters("RemoveFormAttributes") <> ""  Then
							' not yet implemented.
							Call Fn_Web_UI_Button_Click("Fn_Web_DC_ConfigureFiltersOperation", objDCWebPage, "RemoveFormAttributes")
						End If
						'Selecting update type : eg : Static , Dynamic ect
                        If dicWebDCConfigFilters("ConfigureFilterUpdateType") <> ""  Then
							Call Fn_Web_UI_List_Select("Fn_Web_DC_ConfigureFiltersOperation", objDCWebPage, "ConfigureFilterUpdateList",dicWebDCConfigFilters("ConfigureFilterUpdateType"))
						End if

						' clicking on Update button
						If dicWebDCConfigFilters("bUpdate") <> "" Then
							If cBool(dicWebDCConfigFilters("bUpdate") ) Then
								Call Fn_Web_UI_Button_Click("Fn_Web_DC_ConfigureFiltersOperation",objDCWebPage, "Update")
							End If
						End If
						Fn_Web_DC_ConfigureFiltersOperation = True
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			Case "VerifyInResults"
					If dicWebDCConfigFilters("PropertyName") <> "" Then
						aFields = split(dicWebDCConfigFilters("PropertyName"),"~")
						aValues = split(dicWebDCConfigFilters("SearchingValue"),"~")

						iRowCnt = objDCWebPage.WebTable("ConfiguredFiltersResults_Table").RowCount
						For iCnt = 0 to UBound(aFields)
							bFlag = False
							For iCount = 1 to iRowCnt -1
								If objDCWebPage.WebTable("ConfiguredFiltersResults_Table").GetCellData(iCount,1) = aFields(iCnt) Then
									If objDCWebPage.WebTable("ConfiguredFiltersResults_Table").GetCellData(iCount,2) = aValues(iCnt) Then
										bFlag = True
										Exit for
									End If
								ElseIf objDCWebPage.WebTable("ConfiguredFiltersResults_Table").GetCellData(iCount,3) = aFields(iCnt) Then
									If objDCWebPage.WebTable("ConfiguredFiltersResults_Table").GetCellData(iCount,4) = aValues(iCnt) Then
										bFlag = True
										Exit for
									End If
								End If
							Next
							If bFlag Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Function [ Fn_Web_DC_ConfigureFiltersOperation ] Successfully verified [ " & aFields(iCnt) & "=" & aValues(iCnt) & " ] in results.") 
							Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_Web_DC_ConfigureFiltersOperation ] Failed to verify [ " & aFields(iCnt) & "=" & aValues(iCnt) & " ] in results.") 
								Exit for
							End If
						Next
						Fn_Web_DC_ConfigureFiltersOperation = bFlag
					End If
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
            Case "VerifyData"
				aKey=dicWebDCConfigFilters.Keys
				aItem=dicWebDCConfigFilters.Items
				For iCount=0 to ubound(aKey)
					bFlag=False
					Select Case aKey(iCount)
	                    Case "ProximityFilterUnit"
							If Trim(objDCWebPage.WebElement(aKey(iCount)).GetROProperty("innerhtml"))=Trim(aItem(iCount)) Then
								bFlag=True
							End If
						Case "ProximityFilter"
							If Trim(objDCWebPage.WebEdit(aKey(iCount)).GetROProperty("value"))=Trim(aItem(iCount)) Then
								bFlag=True
							End If
						Case "TrueShapeFiltering","ValidOverlaysOnly","Append Parts"
							If LCase(dicWebDCConfigFilters(aKey(iCount)))="on" then
								aItem(iCount)=1
							Else
								aItem(iCount)=0
							End if
							If CInt(objDCWebPage.WebCheckBox(aKey(iCount)).GetROProperty("checked"))=CInt(aItem(iCount)) Then
								bFlag=True
							End If
						Case "SavedQuery"
							If Trim(objDCWebPage.WebList(aKey(iCount)).GetROProperty("value"))=Trim(aItem(iCount)) Then
								bFlag=True
							End If
					End Select
					If bFlag=False Then
						Exit for
					End If
				Next
				If bFlag=True Then
					Fn_Web_DC_ConfigureFiltersOperation=True
				Else
					Fn_Web_DC_ConfigureFiltersOperation=False
				End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			Case Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_Web_DC_ConfigureFiltersOperation ] Invalied case [ " & sAction& " ].") 
	End Select

	If Fn_Web_DC_ConfigureFiltersOperation = True Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Function [ Fn_Web_DC_ConfigureFiltersOperation ] Executed successfully with case [ " & sAction& " ].") 
	End If

	Set objDCWebPage = Nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name	:	Fn_SISW_Web_DC_CreateNewZone
'@@
'@@    Description		:	Function Used to create New Zone in Design Context
'@@
'@@    Parameters		:	1. dicNewZone 	: dictionary object to create Zone
'@@
'@@    Return Value		: 	True Or False
'@@
'@@    Pre-requisite	:	Should Be Log in Web Client And Design Context perspective should be open
'@@
'@@    Examples			: 	Dim dicNewZone
'@@							Set dicNewZone = CreateObject("Scripting.Dictionary")
'@@							With dicNewZone
'@@								.Add "Name", "name"
'@@								.Add "Description", "name"
'@@								.Add "Type", "RDVBoxZoneFormType"
'@@								.Add "ChangeID", "1"
'@@								.Add "Reason", "1"
'@@								.Add "ZoneDetails", "Edge Vector 2 - X Length=1~Edge Vector 1 - X Length=3"
'@@							End With
'@@							msgbox Fn_SISW_Web_DC_CreateNewZone(dicNewZone)
'@@
'@@	   History:				
'@@		Developer Name			Date			Rev. No.		Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@		Koustubh Watwe		25-Sept-2012		1.0				Created.
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_Web_DC_CreateNewZone(dicNewZone)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_Web_DC_CreateNewZone"
	Fn_SISW_Web_DC_CreateNewZone = False
	Dim objDialog, strWEBMenuPath, strMenu, objButtonPanel
	Dim arrDetails, iCount, arrData, iRowCount, objEdit,objDMR

	Set objDialog = Browser("Teamcenter Web - Design").Page("Teamcenter Web - Design").WebTable("NewZone")
	Set objButtonPanel = Browser("Teamcenter Web - Design").Page("Teamcenter Web - Design").WebTable("ButtonTable")
	If Fn_Web_UI_ObjectExist("Fn_SISW_Web_DC_CreateNewZone",objDialog ) = False Then
		'perform menu operations
		strWEBMenuPath = Fn_LogUtil_GetXMLPath("WEB_DC_Menu")
		strMenu = Fn_GetXMLNodeValue(strWEBMenuPath, "FileNewZone")
		Call Fn_Web_MenuOperation("Select",strMenu)

		If Fn_Web_UI_ObjectExist("Fn_SISW_Web_DC_CreateNewZone", objDialog) = False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_Web_DC_CreateNewZone ] Failed to find New Zone dialog.")
			Exit Function
		End If
	End If
	
	Call Fn_Web_UI_WebEdit_Set("Fn_SISW_Web_DC_CreateNewZone", objDialog,"Name", dicNewZone("Name"))

	If dicNewZone("Description") <> "" Then
		Call Fn_Web_UI_WebEdit_Set("Fn_SISW_Web_DC_CreateNewZone", objDialog,"Description", dicNewZone("Description"))
	End If

	If dicNewZone("Type") <> "" Then
        Call Fn_Web_UI_WebEdit_SetExt("Fn_SISW_Web_DC_CreateNewZone","Set", objDialog, "Type", dicNewZone("Type"))
	End If

	Call Fn_Web_UI_Button_Click("Fn_SISW_Web_DC_CreateNewZone", objButtonPanel, "OK")
	wait 3
	Call Fn_Web_UI_Button_Click("Fn_SISW_Web_DC_CreateNewZone", objButtonPanel, "CheckOutAndEdit")

	Call Fn_Web_CheckOutObject(dicNewZone("ChangeID"), dicNewZone("Reason"))

	Browser("Teamcenter Web - Design").Page("Teamcenter Web - Design").WebElement("GenericDialog_Title").SetTOProperty "innertext", dicNewZone("Name")
	Set objDialog = Browser("Teamcenter Web - Design").Page("Teamcenter Web - Design").WebTable("GenericDialog")
	'assign title to generic dialog
	If dicNewZone("ZoneDetails") <> "" Then
		
		arrDetails = split(dicNewZone("ZoneDetails"),"~")
		iRowCount = cInt(objDialog.RowCount())
		For iCount = 0 to UBound(arrDetails)
			Set objMDR=CreateObject("Mercury.DeviceReplay")
			arrData = split(arrDetails(iCount),"=")
			For iCnt = 1 to iRowCount
				If trim(objDialog.GetCellData(icnt,1)) = arrData(0) & ":" Then
					Set objEdit = objDialog.ChildItem(icnt, 2,"WebEdit",0)
					If TypeName(ObjEdit)<>"Nothing" Then
						ObjEdit.Set ""
						ObjEdit.Object.focus
                        objMDR.SendString arrData(1)
					Else
						Exit function
					End If
					Set ObjEdit=Nothing
					Exit for
				End If
			Next
		Next
		Set objMDR=Nothing
	End If

	Call Fn_Web_UI_Button_Click("Fn_SISW_Web_DC_CreateNewZone", objButtonPanel, "SaveAndCheckIn")

	Fn_SISW_Web_DC_CreateNewZone = True
	Set objButtonPanel = Nothing
	Set objDialog = Nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name	:	Fn_SISW_Web_DC_ResultsOperations
'@@
'@@    Description		:	Function Used to perform operations on Search Results in Design Context
'@@
'@@    Parameters		:	1. sAction	: Action to perform
'@@							2. sRow		: Row Text
'@@							3. sField	: Column / Field Name
'@@							4. sValue	: Value to be verified
'@@							5. bOpenInLifecycleVisualization	: Boolean value to click on Open In Lifecycle Visualization button
'@@
'@@    Return Value		: 	True Or False
'@@
'@@    Pre-requisite	:	Should Be Log in Web Client And Design Context perspective should be open
'@@
'@@    Examples			: 	msgbox Fn_SISW_Web_DC_ResultsOperations("GetColumnID", "", "Installation Assembly", "", "")
'@@							msgbox Fn_SISW_Web_DC_ResultsOperations("SelectInResultsTable", "0123flipfone_front_top/A;1", "", "", True)
'@@							msgbox Fn_SISW_Web_DC_ResultsOperations("VerifyInResultsTable", "0123flipfone_front_top/A;1", "", "", "")
'@@							msgbox Fn_SISW_Web_DC_ResultsOperations("VerifyInResultsTable", "0123flipfone_front_top/A;1", "Installation Assembly", "0123flipfone_assembly/A;1 (View)", "")
'@@
'@@	   History:				
'@@		Developer Name			Date			Rev. No.		Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@		Koustubh Watwe		25-Sept-2012		1.0				Created.
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_Web_DC_ResultsOperations(sAction, sRow, sField, sValue, bOpenInLifecycleVisualization)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_Web_DC_ResultsOperations"
	Dim objPage, iRowCount, iCnt, iColID, objCheckBox
	Fn_SISW_Web_DC_ResultsOperations = False
	iColID = -1
	Set objPage = Browser("Teamcenter Web - Design").Page("Teamcenter Web - Design")
	Select Case sAction
		Case "VerifyBackgroundPartAppearancesCount"
			iCnt=Split(objPage.WebElement("BackgroundPartAppearances_label").GetROProperty("innertext"),":")
			iCnt(1)=Trim(iCnt(1))
			If Cint(iCnt(1))=Cint(sValue) Then
				Fn_SISW_Web_DC_ResultsOperations=True
			End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "GetColumnID"
			Fn_SISW_Web_DC_ResultsOperations =  -1
			For iCnt = 2 to  objPage.WebTable("SearchResultTable").ColumnCount(1)
				If objPage.WebTable("SearchResultTable").GetCellData(1,iCnt) = sField Then
					Fn_SISW_Web_DC_ResultsOperations = iCnt
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [ Fn_SISW_Web_DC_ResultsOperations ] Successfuly found column [ " & sField & " ] at position [ " & iCnt & " ].")
					Exit function
				End if
			Next
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "SelectInResultsTable"
			iRowCount = cInt(objPage.WebTable("SearchResultTable").RowCount())
			For iCnt = 2 to iRowCount
				If objPage.WebTable("SearchResultTable").GetCellData(iCnt,2) = sRow Then
					Set objCheckBox = objPage.WebTable("SearchResultTable").ChildItem(iCnt,1,"WebCheckbox",0)
					If trim(typename(objCheckBox)) <> "" Then
						objCheckBox.set "ON"
						Fn_SISW_Web_DC_ResultsOperations = True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [ Fn_SISW_Web_DC_ResultsOperations ] Successfuly selected [ " & sRow & " ].")
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_Web_DC_ResultsOperations ] Failed to find checkbox for [ " & sRow & " ].")
						Exit Function
					End If
					Set objCheckBox=Nothing
					Exit for
				End If
			Next
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "VerifyInResultsTable"
			iRowCount = cInt(objPage.WebTable("SearchResultTable").RowCount())
			If sField <> "" Then
				iColID = Fn_SISW_Web_DC_ResultsOperations("GetColumnID", "", sField, "", "")
			End If
			For iCnt = 2 to iRowCount
				If objPage.WebTable("SearchResultTable").GetCellData(iCnt,2) = sRow Then
					Fn_SISW_Web_DC_ResultsOperations = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [ Fn_SISW_Web_DC_ResultsOperations ] Successfuly verified existence of [ " & sRow & " ].")
					If iColID <> -1 Then
						Fn_SISW_Web_DC_ResultsOperations = False
						If objPage.WebTable("SearchResultTable").GetCellData(iCnt, iColID) = sValue Then
							Fn_SISW_Web_DC_ResultsOperations = True
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [ Fn_SISW_Web_DC_ResultsOperations ] Successfuly verified [ " & sField & " = " & sValue & " ].")
						End If
					End If
					Exit for
				End If
			Next
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_Web_DC_ResultsOperations ] Invalid case [ " & sAction & " ].")
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
	End Select
	If bOpenInLifecycleVisualization <> "" Then
		If cBool(bOpenInLifecycleVisualization) Then
			Call Fn_Web_UI_Button_Click("Fn_SISW_Web_DC_ResultsOperations", objPage, "OpenInLifecycleVisualization")
		End If
	End If
	If Fn_SISW_Web_DC_ResultsOperations <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [ Fn_SISW_Web_DC_ResultsOperations ] executed successfuly with case [ " & sAction & " ].")
	End If
	Set objPage = Nothing
End Function

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name	:	Fn_SISW_Web_DC_ConfigureWorkPartContextOperations
'@@
'@@    Description		:	Function Used to perform operations on Configure Work Part Context block in Design Context
'@@
'@@    Parameters		:	1. sAction	: Action to perform
'@@							2. sRevisionRule ; Revision rule
'@@							3. sVariants 
'@@							4. sVariantValues
'@@							5. sVarName
'@@							6. sVarDescription
'@@							7. sSavedConfiguration
'@@							8. bAddConfigRule
'@@							9. sRow
'@@							10. sCol
'@@							11. sValue
'@@							12. bUpdate
'@@
'@@    Return Value		: 	True Or False
'@@
'@@    Pre-requisite	:	Should Be Log in Web Client And Design Context perspective should be open
'@@
'@@    Examples			: 	msgbox Fn_SISW_Web_DC_ConfigureWorkPartContextOperations("Set", "Latest Working", "001142:Levl1 (String)", "B", "", "", "C", True, "", "", "", True)
'@@							msgbox Fn_SISW_Web_DC_ConfigureWorkPartContextOperations("GetColumnID_ConfigureWorkPartContextTable", "", "", "", "", "", "", "", "", "Installation Assembly", "", "")
'@@							msgbox Fn_SISW_Web_DC_ConfigureWorkPartContextOperations("VerifyInConfigureWorkPartContextTable", "", "", "", "", "", "", "", "Coffee/A;1 (View)", "", "", "")
'@@							msgbox Fn_SISW_Web_DC_ConfigureWorkPartContextOperations("VerifyInConfigureWorkPartContextTable", "", "", "", "", "", "", "", "Coffee/A;1 (View)", "Installation Assembly", "HotDrink/A;1 (View)", "")
'@@							msgbox Fn_SISW_Web_DC_ConfigureWorkPartContextOperations("SelectInConfigureWorkPartContextTable", "", "", "", "", "", "", "", "Milk/A;1", "", "", "")
'@@							msgbox Fn_SISW_Web_DC_ConfigureWorkPartContextOperations("IsBOMLineSelectedInConfigureWorkPartContextTable", "", "", "", "", "", "", "", "Milk/A;1", "", "", "")
'@@
'@@	   History:				
'@@		Developer Name			Date			Rev. No.		Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@		Koustubh Watwe		27-Sept-2012		1.0				Created.
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@		Sandeep Navghane		29-July-2013		1.1				Added Case : VerifyCurrentRevisionRule		Veena G
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_Web_DC_ConfigureWorkPartContextOperations(sAction, sRevisionRule, sVariants, sVariantValues, sVarName, sVarDescription, sSavedConfiguration, bAddConfigRule, sRow, sCol, sValue, bUpdate)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_Web_DC_ConfigureWorkPartContextOperations"
	Dim objTable , objPage, iRowCount, iCnt, iColID,objChk
	Set objPage = Browser("Teamcenter Web - Design").Page("Teamcenter Web - Design")
	Set objTable = objPage.WebTable("ConfigureWorkPartContext_Table")
	Fn_SISW_Web_DC_ConfigureWorkPartContextOperations = False
	iColID = -1

	Select Case sAction
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Set"
			If sRevisionRule <> "" Then
				Call Fn_Web_UI_List_Select("Fn_SISW_Web_DC_ConfigureWorkPartContextOperations", objPage, "RevisionRule",sRevisionRule)
			End If

			If sVariants <> "" Then
				  Call  Fn_WebPSE_VariantConfigurationOperations("LoadVariantConfiguration",  sVariants, sVariantValues, sVarName, sVarDescription, sSavedConfiguration)
			End If

			If bAddConfigRule <> "" Then
				If cBool(bAddConfigRule) Then
					objPage.Image("simple_search").Click 1, 1,micLeftBtn
				End If
			End If
			Fn_SISW_Web_DC_ConfigureWorkPartContextOperations = True
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "GetColumnID_ConfigureWorkPartContextTable"
            Fn_SISW_Web_DC_ConfigureWorkPartContextOperations =  -1
			For iCnt = 2 to  objTable.ColumnCount(1)
				If objTable.GetCellData(1,iCnt) = sCol Then
					Fn_SISW_Web_DC_ConfigureWorkPartContextOperations = iCnt
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [ Fn_SISW_Web_DC_ConfigureWorkPartContextOperations ] Successfuly found column [ " & sCol & " ] at position [ " & iCnt & " ].")
					Exit function
				End if
			Next
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "VerifyInConfigureWorkPartContextTable"
			iRowCount = cInt(objTable.RowCount())
			If sCol <> "" Then
				iColID = Fn_SISW_Web_DC_ConfigureWorkPartContextOperations("GetColumnID_ConfigureWorkPartContextTable", "", "", "", "", "", "", "", "", sCol, "", "")
			End If
			For iCnt = 2 to iRowCount
                If objTable.GetCellData(iCnt, 2) = sRow Then
					Fn_SISW_Web_DC_ConfigureWorkPartContextOperations = true
					If iColID <> -1 Then
						Fn_SISW_Web_DC_ConfigureWorkPartContextOperations = False
						If objTable.GetCellData(iCnt, iColID) = sValue Then
							Fn_SISW_Web_DC_ConfigureWorkPartContextOperations = true
						End If
					End If
					Exit for
				End If
			Next
        '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "SelectInConfigureWorkPartContextTable"
			Fn_SISW_Web_DC_ConfigureWorkPartContextOperations = False
			iRowCount = cInt(objTable.RowCount())
			For iCnt = 2 to iRowCount
                If objTable.GetCellData(iCnt, 2) = sRow Then
						Set objChk=objTable.ChildItem(iCnt,0,"WebCheckBox",0)
						If Typename(objChk)<>"Nothing" Then
							objChk.Set "ON"
							Fn_SISW_Web_DC_ConfigureWorkPartContextOperations = true
						End If
						Set objChk=Nothing
						Exit for
				End If
			Next
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "IsBOMLineSelectedInConfigureWorkPartContextTable"
			Fn_SISW_Web_DC_ConfigureWorkPartContextOperations = False
			iRowCount = cInt(objTable.RowCount())
			For iCnt = 2 to iRowCount
                If objTable.GetCellData(iCnt, 2) = sRow Then
						Set objChk=objTable.ChildItem(iCnt,0,"WebCheckBox",0)
						If Typename(objChk)<>"Nothing" Then
							if Cint(objChk.GetROProperty("checked"))=1 then
								Fn_SISW_Web_DC_ConfigureWorkPartContextOperations = true
							End if
						End If
						Set objChk=Nothing
						Exit for
				End If
			Next
        '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "VerifyCurrentRevisionRule"
			If trim(objPage.WebList("RevisionRule").GetROProperty("value"))=trim(sRevisionRule) Then
				Fn_SISW_Web_DC_ConfigureWorkPartContextOperations=True
			End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_Web_DC_ConfigureWorkPartContextOperations ] Invalid case [ " & sAction & " ].")
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
	End Select
	If bUpdate <> "" Then
					If cBool(bUpdate) Then
						' click on button
						Call Fn_Web_UI_Button_Click("Fn_Web_DC_ContextDefinitionOperations",objPage,  "Update")
					End If
				End If
	If Fn_SISW_Web_DC_ConfigureWorkPartContextOperations <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [ Fn_SISW_Web_DC_ConfigureWorkPartContextOperations ] executed successfuly with case [ " & sAction & " ].")
	End If
	Set objPage = Nothing
End Function
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'
'    Function Name	:	Fn_SISW_Web_DC_SaveStructureContextObject
'
'    Description		:	Function Used to Save Structure Context Object
'
'    Parameters		:	1. StrAction	: Action to perform
'							2. StrName : Name
'							3. StrDescription  : Description 
'							4. StrMessage : Expected information message
'							5. StrButton
'
'    Return Value		: 	True Or False
'
'    Pre-requisite	:	Should Be Log in Web Client And Design Context perspective should be open
'
'    Examples			: 	bReturnFn_SISW_Web_DC_SaveStructureContextObject("SaveAndVerifyInformation","SaveObject2","Save Desc","The Structure Context object SaveObject2 is successfully saved","")
'										bReturnFn_SISW_Web_DC_SaveStructureContextObject("Save","SaveObject1","","","")
'
'	   History:				
'
'		Developer Name			Date			Rev. No.		Changes Done								Reviewer
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'		Sandeep N					24-July-2013		1.0				Created											Veena G
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function Fn_SISW_Web_DC_SaveStructureContextObject(StrAction,StrName,StrDescription,StrMessage,StrButton)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_Web_DC_SaveStructureContextObject"
   'Declaring variables
   Dim ObjSaveObject,ObjButtonPanel
   Dim StrCrrInfo
   Fn_SISW_Web_DC_SaveStructureContextObject=False
   'Checking Existance of [ Save Structure Context Object ] dialog
   If not Browser("Teamcenter Web - Design").Page("Teamcenter Web - Design").WebTable("SaveStructureContextObject").Exist(6) Then
 		'Calling menu [ File => Save Structure Context Object As... ] to invoke [ Save Structure Context Object ] dialog
		If Fn_Web_MenuOperation("Select","File:Save Structure Context Object As...")=False Then
			Exit function
		End If
		'Sync page
		Call Fn_Web_ReadyStatusSync(1)
   End If
   'Creating object of [ Save Structure Context Object ] dialog
   Set ObjSaveObject=Browser("Teamcenter Web - Design").Page("Teamcenter Web - Design").WebTable("SaveStructureContextObject")
   'Creating object of [ Button Panel ]
   Set ObjButtonPanel=Browser("Teamcenter Web - Design").Page("Teamcenter Web - Design").WebElement("ButtonPanel")
   Select Case StrAction 
	 	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	 	'Case Save := To Save  Structure Context Object
		'Case SaveAndVerifyInformation := To Save  Structure Context Object and verify Information
	 	Case "Save","SaveAndVerifyInformation"
			'Setting Name
			Call Fn_Web_UI_WebEdit_Set("Fn_SISW_Web_DC_SaveStructureContextObject", ObjSaveObject, "Name",StrName)
			'Setting Desciption
			If StrDescription<>"" Then
				Call Fn_Web_UI_WebEdit_Set("Fn_SISW_Web_DC_SaveStructureContextObject", ObjSaveObject, "Description",StrDescription)
			End If
			'Click on [ OK ] button
			Call Fn_Web_UI_Button_Click("Fn_SISW_Web_DC_SaveStructureContextObject", ObjButtonPanel, "OK")
			'Checking Existance of Information Dialog
			If Browser("TeamcenterWeb").Dialog("Dialog").Exist(5) Then
				If StrAction="SaveAndVerifyInformation" Then
					'Verify Information of Save Object
					StrCrrInfo=Browser("TeamcenterWeb").Dialog("Dialog").Static("ErrorMsg").GetROProperty("text")
					If instr(1,StrCrrInfo,StrMessage) Then
						'Clicking on [ OK ] button to finish operation
						Fn_SISW_Web_DC_SaveStructureContextObject=Fn_Web_UI_WinButton_Click("Fn_SISW_Web_DC_SaveStructureContextObject", Browser("TeamcenterWeb").Dialog("Dialog"), "OK","","","")
					Else
						'Clicking on [ OK ] button to finish operation
						Call Fn_Web_UI_WinButton_Click("Fn_SISW_Web_DC_SaveStructureContextObject", Browser("TeamcenterWeb").Dialog("Dialog"), "OK","","","")
					End If
				Else
					'Clicking on [ OK ] button to finish operation
					Fn_SISW_Web_DC_SaveStructureContextObject=Fn_Web_UI_WinButton_Click("Fn_SISW_Web_DC_SaveStructureContextObject", Browser("TeamcenterWeb").Dialog("Dialog"), "OK","","","")
				End If
				'Sync Page
				Call Fn_Web_ReadyStatusSync(1)
			End If
	End Select
	'Releasing Object
	Set ObjSaveObject=Nothing
	Set ObjButtonPanel=Nothing
End Function

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name	:	Fn_SISW_Web_DC_FormAttributesTableOperation
'@@
'@@    Description		:	Function Used to perform operations on  Form Attributes Table
'@@
'@@    Parameters		:	1. sAction	: Action to perform
'@@										 2. dicFormAttributes : Dictionary object of parameters     Note: Refer the Examples
'@@
'@@    Return Value		: 	True Or False
'@@
'@@    Pre-requisite	:	Should Be Log in Web Client And Design Context perspective should be open and form Aattributes are created
'@@
'@@    Examples			: 	  Set  dicFormAttributes = CreateObject("Scripting.Dictionary")
'@@ 									  dicFormAttributes("PrimaryColumn") = "Relation Type"
'@@ 									  dicFormAttributes("SelectValues") = "Specifications"
'@@ 									  dicFormAttributes("bUpdate") = True
'@@ 									  msgbox Fn_SISW_Web_DC_FormAttributesTableOperation("SelectInFormAttributeTable",dicFormAttributes)
'@@ 
'@@                                      Set  dicFormAttributes = CreateObject("Scripting.Dictionary")
'@@ 									  dicFormAttributes("PrimaryColumn") = "Relation Type"
'@@ 									  dicFormAttributes("SelectValues") = "Specifications@2"
'@@ 									  dicFormAttributes("bUpdate") = True
'@@ 									  msgbox Fn_SISW_Web_DC_FormAttributesTableOperation("SelectInFormAttributeTable",dicFormAttributes)							  
'@@
'@@                                      Set  dicFormAttributes = CreateObject("Scripting.Dictionary")
'@@ 									  dicFormAttributes("PrimaryColumn") = "Relation Type"
'@@ 									  dicFormAttributes("SelectValues") = "Specifications@2~References@3"
'@@ 									  dicFormAttributes("bUpdate") = False
'@@ 									  msgbox Fn_SISW_Web_DC_FormAttributesTableOperation("SelectInFormAttributeTable",dicFormAttributes)
'@@	   History:		
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------		
'@@		Developer Name	 \ 		Date			        \     Rev. No.	     \   Changes Done			    \ 				         	Reviewer
'@@---------------------------------------------------------- ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@		Pritam Shikare	      \ 	 07-Aug-20123    \ 	    1.0			        \	   Created.                         \                          Yogini Muluk
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_SISW_Web_DC_FormAttributesTableOperation(sAction, dicFormAttributes)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_Web_DC_FormAttributesTableOperation"

   Dim objAttrTable, objChk
   Dim iCount, jCount, iColmID, iRows, iCols, aValues, sValue,  aInstance, iInstance, iCurrInstance

	Fn_SISW_Web_DC_FormAttributesTableOperation = False

	Set objPage = Fn_SISW_Web_DC_GetObject("Teamcenter Web - Design")
	Set objAttrTable = objPage.WebTable("FormAttributeTable")

    Select Case sAction

	 	Case "SelectInFormAttributeTable"

				'get no. of Rows and Columns
				iRows = objAttrTable.GetROProperty("rows")
				iCols = objAttrTable.GetROProperty("cols")

				'get the Column No. i.e col ID of the Primary Column
				If dicFormAttributes("PrimaryColumn") <> "" Then
						 For iCount = 1 to iCols
							If objAttrTable.GetCellData(1,iCount) = dicFormAttributes("PrimaryColumn") Then
								iColmID = iCount
								Exit For
							End If
						 Next
				End If

				'Split Values
				If  dicFormAttributes("SelectValues") <> "" Then
					aValues = Split(dicFormAttributes("SelectValues"),"~",-1,1)
				End If

				'Select the Required Rows
				For iCount = 0 to uBound(aValues)

						'Get the Instance
						If Instr(1, aValues(iCount),"@") Then
							aInstance = Split(aValues(iCount),"@",-1,1)
							iInstance = Cint(aInstance(1))
							sValue = cstr(aInstance(0))
						Else
							iInstance = 1
							sValue = aValues(iCount)
						End If

						'Select the Row
						iCurrInstance = 0
						For jCount = 2 to iRows

								If objAttrTable.GetCellData(jCount,iColmID) = sValue Then
									iCurrInstance = iCurrInstance+1
									If Cint(iCurrInstance) = iInstance Then
										Set objChk=objAttrTable.ChildItem(jCount,0,"WebCheckBox",0)
										If Typename(objChk)<>"Nothing" Then
												objChk.Set "ON"
												wait 1
										End If
										Set objChk=Nothing
										Exit For
									End If
								End If
				
						Next '//Inner for lloop close

						If jCount = iRows+1 Then
							Fn_SISW_Web_DC_FormAttributesTableOperation = False
							Exit Function
						End If
				Next'//Outer For Loop

				'Click on Update Button
				If dicFormAttributes("bUpdate") = True Then
					objPage.WebButton("Update").Click 5,5,micLeftBtn
				End If

		Case else

				Fn_SISW_Web_DC_FormAttributesTableOperation = False
				Exit Function
    End Select

	Fn_SISW_Web_DC_FormAttributesTableOperation = True

End Function
