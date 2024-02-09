Option Explicit
'Function List
'************************************************************************************************************************************************************************************************************
'000. Fn_SISW_RDV_GetObject()
'001. Fn_RDV_ItemAttributes()
'002. Fn_RDV_ChangeSearch()
'003. Fn_RDV_FormAttributes()
'004. Fn_RDV_OccurrenceNotes()
'005. Fn_RDV_Classifications()
'006. Fn_RDV_SpatialCriteria()
'007. Fn_RDV_ItemIDSearchPanelOperations()
'008. Fn_RDV_SearchResultsOperations()
'009. Fn_RDV_FormAttributesSearchPanelOperations()
'010. Fn_RDV_OccurrenceNotesSearchPanelOperations()
'011. Fn_RDV_ClassificationSearchPanelOperations()
'012. Fn_RDV_MSM_ItemIDSearchPanelOperations()
'013. Fn_RDV_MSM_FormAttributesSearchPanelOperations()
'014. Fn_RDV_MSM_OccurrenceNotesSearchPanelOperations()
'015. Fn_RDV_MSM_ClassificationSearchPanelOperations()
'016. Fn_RDV_MSM_SearchResultsOperations()
'017. Fn_RDV_SpatialCriteriaSearchPanelOperations()
'018. Fn_RDV_MSM_SpatialCriteriaSearchPanelOperations()
'019. Fn_RDV_MSM_ConfirmationBoxOperations()
'020. Fn_RDV_SE_ItemIDSearchPanelOperations()
'021. Fn_RDV_SE_FormAttributesSearchPanelOperations()
'022. Fn_RDV_SE_OccurrenceNotesSearchPanelOperations()
'023. Fn_RDV_SE_ClassificationSearchPanelOperations()
'024. Fn_RDV_SE_SearchResultsOperations()
'025. Fn_RDV_SE_SpatialCriteriaSearchPanelOperations()
'026. Fn_RDV_SpatialFilterUseSelectionTableOperation()
'027. Fn_SISW_RDV_MSM_ScopesOperation()
'028. Fn_SISW_RDV_MSM_SearchCriteriaOperations()
'029. Fn_MSM_TabSet()
'030. Fn_SISW_RDV_SpatialFilterOperations()
'************************************************************************************************************************************************************************************************************
'****************************************    Function to get Object hierarchy ***************************************
'
''Function Name		 	:	Fn_SISW_RDV_GetObject
'
''Description		    :  	Function to get specified Object hierarchy.

''Parameters		    :	1. sObjectName : Object Handle name
								
''Return Value		    :  	Object \ Nothing
'
''Examples		     	:	Fn_SISW_RDV_GetObject("Remove")

'History:
'	Developer Name			Date			Rev. No.		Reviewer		Changes Done	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Koustubh Watwe		 7-June-2012		1.0	
'	Ashok kakade		 28-June-2012		1.0								Added Case "RDVJApplet"
'-----------------------------------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------------------------------
'	Anurag Khera		 26-Sept-2013		2.0								Externalized object hierarchies
'-----------------------------------------------------------------------------------------------------------------------------------
Function Fn_SISW_RDV_GetObject(sObjectName)
	Dim sObjectXMLPath
	sObjectXMLPath = Fn_GetEnvValue("User", "AutomationDir") & "\TestData\AutomationXML\ObjectXML\RDV.xml"
	Set Fn_SISW_RDV_GetObject = Fn_SISW_Setup_GetObjectFromXML(sObjectXMLPath, sObjectName)
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_RDV_ItemAttributes
'@@
'@@    Description				 :	Function Used to perform search operation on Item Attribute dialog
'@@
'@@    Parameters			   :	1. dicItemIDSearch: dictionary object
'@@
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	Item Attribute dialog should be displayed							
'@@
'@@    Examples					:	dicItemAttributes("bChangeSearch") = True
'@@    								dicItemAttributes("AdvancedDefaultSearchType") = "Item..."
'@@    								dicItemAttributes("bClearHistory") = "true"
'@@    								dicItemAttributes("RememberMyLastSearches") = "10"
'@@    								dicItemAttributes("SearchType") = "Item..."
'@@    								dicItemAttributes("SearchCriteria") = "Item ID=000038~Name=comp1"
'@@    								Call Fn_RDV_ItemAttributes(dicItemAttributes)
'@@
'@@	   History					 	:	
'@@				Developer Name				Date			Rev. No.	Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe			9-Jan-2012			1.0			Created
'@@				Ashok kakade			29-June-2012		1.0			Added New Hierarchy of Dialog ItemAttributes	
'@@				Koustubh Watwe			4-Sep-2012			1.0			Modified code to select date
'@@				Ganesh B				30-May-2014			1.0			Modified function to handle new Date Control
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_RDV_ItemAttributes(dicItemAttributes)
	GBL_FAILED_FUNCTION_NAME="Fn_RDV_ItemAttributes"
' verifying Item Attribute's window
	Dim objItemAttrib, arrSearchCriteria, arrFieldValue, iCnt
	Dim sTemplateType, intNoOfObjects, iCount
	Dim sDate1, arrDate
	Dim hieght,width
	Dim WshShell
	Fn_RDV_ItemAttributes = False
	If JavaWindow("RDV_StructureManager").JavaWindow("RDV_PSEWindow").JavaDialog("ItemAttributes").Exist(2) Then
			Set objItemAttrib = JavaWindow("RDV_StructureManager").JavaWindow("RDV_PSEWindow").JavaDialog("ItemAttributes")
	ElseIf JavaWindow("RDV_StructureManager").JavaWindow("RDV_JApplet").JavaDialog("ItemAttributes").Exist(2) Then
		Set objItemAttrib = JavaWindow("RDV_StructureManager").JavaWindow("RDV_JApplet").JavaDialog("ItemAttributes")
	Else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_RDV_ItemAttributes ] Failed to open [ Item Attribute ] dialog.") 
		Fn_RDV_ItemAttributes = False
		Exit function
	End If
'	
'	If Fn_UI_ObjectExist("Fn_RDV_ItemAttributes",objItemAttrib) = False Then
'		Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_RDV_ItemAttributes ] Failed to open [ Item Attribute ] dialog.") 
'		Exit function
'	End If
'	objItemAttrib.Resize 460, 900
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Resizing window
	hieght=JavaWindow("RDV_StructureManager").GetROProperty("height")
	width=JavaWindow("RDV_StructureManager").GetROProperty("width")
	objItemAttrib.Move 0,0
	wait 2
	objItemAttrib.Resize width-5,hieght-5
	wait 2

	If lcase(dicItemAttributes("bChangeSearch")) = "true" then
		Call Fn_Button_Click("Fn_RDV_ItemAttributes", objItemAttrib, "Change")
		IF Fn_RDV_ChangeSearch(dicItemAttributes) = False then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_RDV_ItemAttributes ] Failed to open [ change Search ] dialog.") 
			Exit function
		End IF
	End If
	Call Fn_ReadyStatusSync(2)
	' set values
	If dicItemAttributes("SearchCriteria") <> "" Then
		' clearing form
		Call Fn_Button_Click("Fn_RDV_ItemAttributes", objItemAttrib, "Clear")

		' clicking on more static text
		If Fn_UI_ObjectExist("Fn_RDV_ItemAttributes", objItemAttrib.JavaStaticText("More")) then
			objItemAttrib.JavaStaticText("More").Click 1,1,"LEFT"
		End If
		Call Fn_ReadyStatusSync(2)
		arrSearchCriteria = split(dicItemAttributes("SearchCriteria"),"~")
		For iCnt = 0 to UBound(arrSearchCriteria)
				arrFieldValue = split(arrSearchCriteria(iCnt),"=")
				objItemAttrib.JavaStaticText("Field").SetTOProperty  "label", trim(arrFieldValue(0)) & ":"
				wait 1
				Select Case arrFieldValue(0)
					Case "Created After", "Created Before", "Modified Before", "Modified After", "Released Before", "Released After"
						If lcase(trim(arrFieldValue(1) )) = "today" Then
							sDate1 = now
							arrDate = Split(sDate1, " ")	
						Else
							'' "14-May-2014 11:36"
							arrDate = Split(arrFieldValue(1), " ")	
						End if
						Call Fn_Edit_Box("Fn_RDV_ItemAttributes", objItemAttrib,"Field",trim(arrDate(0)))
						Set WshShell = CreateObject("WScript.Shell")
						wait(1)
						WshShell.SendKeys "{ESC}"
						wait(1)
						Call Fn_Edit_Box("Fn_RDV_ItemAttributes", objItemAttrib,"Time",trim(arrDate(1)))
						wait(1)
						WshShell.SendKeys "{TAB}"
						wait(1)
						Set WshShell =Nothing 
					Case Else
						If objItemAttrib.JavaButton("MultipleDropdownButton").Exist(2) Then
							objItemAttrib.JavaButton("MultipleDropdownButton").Click micLeftBtn
							wait(5)
							Set sTemplateType = Description.Create()
							sTemplateType("Class Name").value = "JavaStaticText"
							Set intNoOfObjects = objItemAttrib.ChildObjects(sTemplateType)
							For iCount = 0 to intNoOfObjects.count-1
								If  intNoOfObjects(iCount).getROProperty("label") = trim(arrFieldValue(1) ) Then
									intNoOfObjects(iCount).Click 1,1
									Exit for
								End If
							Next
							If iCount = intNoOfObjects.count Then
								Exit function
							End If
							Set sTemplateType = Nothing
						ElseIf objItemAttrib.JavaEdit("Field").exist(3) Then
							If  arrFieldValue(0) = "Type" OR arrFieldValue(0) = "Release Status" Then
								Call Fn_Edit_Box("Fn_RDV_ItemAttributes", objItemAttrib,"Field",trim(arrFieldValue(1) ) & vblf)
							Else
		                    	Call Fn_Edit_Box("Fn_RDV_ItemAttributes", objItemAttrib,"Field",trim(arrFieldValue(1) ))
							End If
						Else
							Exit function
						End If
				
				End Select
		Next
	End If
	' clicking on OK button
	Call Fn_Button_Click("Fn_RDV_ItemAttributes", objItemAttrib, "OK")
	Fn_RDV_ItemAttributes = True
	Set objItemAttrib = Nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_RDV_ChangeSearch
'@@
'@@    Description				 :	Function Used to perform search operation using Change Search Criteria
'@@
'@@    Parameters			   :	1. sAction: Action to be performed
'@@    								2. dicItemIDSearch: dictionary object
'@@
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	Change Search dialog should be displayed							
'@@
'@@    								dicChangeSearch("AdvancedDefaultSearchType") = "Item..."
'@@    								dicChangeSearch("bClearHistory") = "true"
'@@    								dicChangeSearch("RememberMyLastSearches") = "10"
'@@    								dicChangeSearch("SearchType") = "Item..."
'@@    								Call Fn_RDV_ChangeSearch("Search", dicChangeSearch)
'@@
'@@	   History					 	:	
'@@				Developer Name				Date			Rev. No.	Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe			9-Jan-2012			1.0			Created
'@@				Ashok kakade			29-June-2012		1.0			Added New Hierarchy of Dialog Change Search	
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_RDV_ChangeSearch(dicChangeSearch)
			GBL_FAILED_FUNCTION_NAME="Fn_RDV_ChangeSearch"
			Dim objChangeSearch
			Fn_RDV_ChangeSearch = False 
			If JavaWindow("RDV_StructureManager").JavaWindow("RDV_PSEWindow").JavaDialog("Change Search").Exist(2) = True Then
					Set objChangeSearch = JavaWindow("RDV_StructureManager").JavaWindow("RDV_PSEWindow").JavaDialog("Change Search")
			ElseIf JavaWindow("RDV_StructureManager").JavaWindow("RDV_JApplet").JavaDialog("Change Search").Exist(2) = True Then
					Set objChangeSearch = JavaWindow("RDV_StructureManager").JavaWindow("RDV_JApplet").JavaDialog("Change Search")
			Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_RDV_ChangeSearch ] Failed to find [ Change Search ] dialog.") 
					Exit function
			End If
		
			' checking existence of Change window
			If Fn_UI_ObjectExist("Fn_RDV_ChangeSearch", objChangeSearch) = False Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_RDV_ChangeSearch ] Failed to find [ Change Search ] dialog.") 
				Exit function
			End If
			' clicking on Advance and setting default searhc type
			' If dicChangeSearch("AdvancedDefaultSearchType") <> "" then
					' Call Fn_Button_Click("Fn_RDV_ChangeSearch", objChangeSearch, "Advanced")
					' Call Fn_ReadyStatusSync(2)

				' ' verifying default serach
				' If trim(cstr(objChangeSearch.JavaDialog("Advanced").JavaEdit("DefaultSearch").GetROProperty("value"))) <> trim(dicChangeSearch("AdvancedDefaultSearchType") ) then
					' Call Fn_Button_Click("Fn_RDV_ChangeSearch", objChangeSearch.JavaDialog("Advanced"),"Change")
					' Call Fn_ReadyStatusSync(2)
					' ' checking existence of Search type in Find Search list
					' If Fn_UI_ListItemExist("Fn_RDV_ChangeSearch", objChangeSearch.JavaDialog("Advanced").JavaDialog("Find Search"), "JList", dicChangeSearch("AdvancedDefaultSearchType")) = True then
						' ' selecting search type
						' Call Fn_List_Select("Fn_RDV_ChangeSearch", objChangeSearch.JavaDialog("Advanced").JavaDialog("Find Search"), "JList", dicChangeSearch("AdvancedDefaultSearchType"))
						' ' clicking on OK button of Advance default Search
						' Call Fn_Button_Click("Fn_RDV_ChangeSearch", objChangeSearch.JavaDialog("Advanced").JavaDialog("Find Search"),"OK")
					' Else
						' Exit function
					' End If
				' End If

				' ' clearing history
				' If dicChangeSearch("bClearHistory") <> "" then
					' If cBool(dicChangeSearch("bClearHistory")) Then
							' Call Fn_Button_Click("Fn_RDV_ChangeSearch", objChangeSearch.JavaDialog("Advanced"),"ClearSearchHistory")
					' End If
				' End If

				' ' setting remember serach count
				' If dicChangeSearch("RememberMyLastSearches") <> "" then
					' If trim(cstr(objChangeSearch.JavaDialog("Advanced").JavaEdit("RememberMyTest").GetROProperty("value"))) <> trim(dicChangeSearch("RememberMyLastSearches") ) then
						' Call Fn_Edit_Box("Fn_RDV_ChangeSearch", objChangeSearch.JavaDialog("Advanced"),"RememberMyTest", trim(dicChangeSearch("RememberMyLastSearches")))
					' End If
				' End If
				
				' 'clicking on Apply
				' If Fn_SISW_UI_Object_Operations("Fn_RDV_ChangeSearch", "Enabled", objChangeSearch.JavaDialog("Advanced").JavaButton("Apply"), 1) Then
					' Call Fn_Button_Click("Fn_RDV_ChangeSearch", objChangeSearch.JavaDialog("Advanced"),"Apply")
				' End If
				' 'clicking on OK
				' If Fn_SISW_UI_Object_Operations("Fn_RDV_ChangeSearch", "Enabled", objChangeSearch.JavaDialog("Advanced").JavaButton("OK"), 1) Then
					' Call Fn_Button_Click("Fn_RDV_ChangeSearch", objChangeSearch.JavaDialog("Advanced"),"OK")
				' End If
				' If Fn_SISW_UI_Object_Operations("Fn_RDV_ChangeSearch", "Exist", objChangeSearch.JavaDialog("Advanced"), 1)  Then
					' Call Fn_Button_Click("Fn_RDV_ChangeSearch", objChangeSearch.JavaDialog("Advanced"),"Cancel")
				' End If
			' End If

			' ' selecting history tab
			' objChangeSearch.JavaTab("JTabbedPane").Select "Search History"
			Call Fn_ReadyStatusSync(2)
			If Fn_UI_ListItemExist("Fn_RDV_ChangeSearch",objChangeSearch,"JList", dicChangeSearch("SearchType")) = False then
				' selecting tab
				objChangeSearch.JavaTab("JTabbedPane").Select "System Defined Searches"
				Call Fn_ReadyStatusSync(2)
				If Fn_UI_ListItemExist("Fn_RDV_ChangeSearch",objChangeSearch,"JList", dicChangeSearch("SearchType")) = False then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_RDV_ChangeSearch ] Failed to find [ " & dicChangeSearch("SearchType")  &" ] in Change Search Window.") 
					Exit function
				End If
			End If
			Call Fn_List_Select("Fn_RDV_ChangeSearch",objChangeSearch,"JList", dicChangeSearch("SearchType")) 
			Call Fn_ReadyStatusSync(2)
			' click on change button
			Call Fn_Button_Click("Fn_RDV_ChangeSearch", objChangeSearch, "ChangeSearch")
			Fn_RDV_ChangeSearch = True
	Set objChangeSearch = Nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_RDV_FormAttributes
'@@
'@@    Description				 :	Function Used to perform search operation using Form Attribute dialog
'@@
'@@    Parameters			   :	1.bClearFormAttributePanel : to clear Form Attribute table ( True / False / "" )
'@@   	 							2.sRelationType : ~ separated list of Relation Types ( String )
'@@   	 							3.sParentType :  ~ separated list of Parent Types ( String )
'@@   	 							4.sFormType :  ~ separated list of Form Types ( String )
'@@   	 							5.sPropertyName :  ~ separated list of Property Names ( String )
'@@   	 							6.sOperator :  ~ separated list of Operators ( String )
'@@   	 							7.sSearchingValue :  ~ separated list of Seawrching Values ( String )
'@@
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	Form Attribute dialog should be displayed						
'@@
'@@    Examples					:	
'@@    								bClearFormAttributePanel = "true"
'@@    								sRelationType = "3D Snapshot~"
'@@    								sParentType = "Item~"
'@@    								sFormType = "Arc Override Form~"
'@@    								sPropertyName = "Start Point RZ~"
'@@    								sOperator = "EQ~"
'@@    								sSearchingValue = "1~2"
'@@    								msgbox  Fn_RDV_FormAttributes(bClearFormAttributePanel, sRelationType, sParentType, sFormType, sPropertyName, sOperator, sSearchingValue)
'@@
'@@	   History					 	:	
'@@				Developer Name				Date			Rev. No.	Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe			10-Jan-2012			1.0			Created
'@@				Ashok kakade			29-June-2012		1.0			Added New Hierarchy of Dialog Form Attributes	
'@@            Vrushali  Wani			  24-July-2012			1.0         Added New  Hierarchy   for  'FormAttributes'  
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_RDV_FormAttributes(bClearFormAttributePanel, sRelationType, sParentType, sFormType, sPropertyName, sOperator, sSearchingValue)
		GBL_FAILED_FUNCTION_NAME="Fn_RDV_FormAttributes"
		Dim objFormAttrib, iCnt, iRowCnt 
		Dim arrRelationTypes, arrParentTypes, arrFormTypes, arrPropertyNames, arrOperators, arrSearchingValues

		If JavaWindow("RDV_StructureManager").JavaWindow("RDV_PSEWindow").JavaDialog("Form Attributes").Exist(2) = True Then
			Set objFormAttrib = JavaWindow("RDV_StructureManager").JavaWindow("RDV_PSEWindow").JavaDialog("Form Attributes")
		ElseIf JavaWindow("RDV_StructureManager").JavaWindow("RDV_JApplet").JavaDialog("FormAttributes").Exist(2)  = True Then
				Set objFormAttrib = JavaWindow("RDV_StructureManager").JavaWindow("RDV_JApplet").JavaDialog("FormAttributes")
		Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_RDV_ChangeSearch ] Failed to find [ Form Attributes ] dialog.") 
				Fn_RDV_FormAttributes = False
				Exit function
		End If

		' clearing Form Attribute details
		If bClearFormAttributePanel = "" then bClearFormAttributePanel = "False"
		If cBool(bClearFormAttributePanel) then
			Call Fn_Button_Click("Fn_RDV_FormAttributes", objFormAttrib, "Clear")
		End If

		arrRelationTypes = split(sRelationType,"~") 
		arrParentTypes = split(sParentType,"~") 
		arrFormTypes = split(sFormType,"~") 
		arrPropertyNames = split(sPropertyName,"~") 
		arrOperators = split(sOperator,"~") 
		arrSearchingValues = split(sSearchingValue,"~") 

		For iCnt = 0 to UBound(arrSearchingValues)
			' selecting relation type
			If trim(arrRelationTypes(iCnt)) <> "" then
				Call Fn_List_Select("Fn_RDV_FormAttributes", objFormAttrib, "RelationType", trim(arrRelationTypes(iCnt)))
			End If

			' selecting Parent type
			If trim(arrParentTypes(iCnt)) <> "" then
				Call Fn_List_Select("Fn_RDV_FormAttributes", objFormAttrib, "ParentType", trim(arrParentTypes(iCnt)))
			End If

			' selecting Form type
			If trim(arrFormTypes(iCnt)) <> "" then
				Call Fn_List_Select("Fn_RDV_FormAttributes", objFormAttrib, "FormType", trim(arrFormTypes(iCnt)))
			End If

			' clicking on Add button
			Call Fn_Button_Click("Fn_RDV_FormAttributes", objFormAttrib, "Add")

			iRowCnt = cInt(objFormAttrib.JavaTable("FormTypeTable").GetROProperty("rows")) - 1
			If trim(arrPropertyNames(iCnt)) <> "" then
				objFormAttrib.JavaTable("FormTypeTable").ClickCell iRowCnt,"Property Name","LEFT"
				Call Fn_List_Select("Fn_RDV_FormAttributes", objFormAttrib, "AttributeList", trim(arrPropertyNames(iCnt)))
			End If

			If trim(arrOperators(iCnt)) <> "" then
				objFormAttrib.JavaTable("FormTypeTable").ClickCell iRowCnt,"Operator","LEFT"
				Call Fn_List_Select("Fn_RDV_FormAttributes", objFormAttrib, "AttributeList", trim(arrOperators(iCnt)))
			End If

			If trim(arrSearchingValues(iCnt)) <> "" then
				objFormAttrib.JavaTable("FormTypeTable").SetCellData iRowCnt,"Searching Value", trim(arrSearchingValues(iCnt)) 
			End If
		Next
		' clicking on OK Button
		Call Fn_Button_Click("Fn_RDV_FormAttributes", objFormAttrib, "OK")
		Fn_RDV_FormAttributes = True
		Set objFormAttrib = Nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_RDV_OccurrenceNotes
'@@
'@@    Description				 :	Function Used to perform search operation using Occurrence Notes dialog
'@@
'@@    Parameters			   :	1.bClearFormAttributePanel : to clear Form Attribute table ( True / False / "" )
'@@   	 							2.sRelationType : ~ separated list of Relation Types ( String )
'@@   	 							3.sParentType :  ~ separated list of Parent Types ( String )
'@@   	 							4.sFormType :  ~ separated list of Form Types ( String )
'@@   	 							5.sPropertyName :  ~ separated list of Property Names ( String )
'@@   	 							6.sOperator :  ~ separated list of Operators ( String )
'@@   	 							7.sSearchingValue :  ~ separated list of Seawrching Values ( String )
'@@
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	Form Attribute dialog should be displayed						
'@@
'@@    Examples					:	
'@@    								bClear = "true"
'@@    								bClearOccurrenceNotes = "true"
'@@    								sOccurrenceNotes = "AIE_Exported~AIE_Exported"
'@@    								sOperators = "EQ~NE"
'@@    								sValues = "1~2"
'@@    								bClickOnSearchButton = ""
'@@    								msgbox Fn_RDV_OccurrenceNotes(bClearOccurrenceNotes, sOccurrenceNotes, sOperators, sValues)
'@@
'@@	   History					 	:	
'@@				Developer Name				Date			Rev. No.	Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe			10-Jan-2012			1.0			Created
'@@				Ashok kakade			29-June-2012		1.0			Added New Hierarchy of Dialog OccurrenceNotes	
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_RDV_OccurrenceNotes(bClearOccurrenceNotes, sOccurrenceNotes, sOperators, sValues)
	GBL_FAILED_FUNCTION_NAME="Fn_RDV_OccurrenceNotes"
	Dim objOccNotes
	Dim arrOccNotes, arrOperators, arrValues, iCnt, iRowCnt
	Fn_RDV_OccurrenceNotes = False

		If  JavaWindow("RDV_StructureManager").JavaWindow("RDV_PSEWindow").JavaDialog("OccurrenceNotes").Exist(2) Then
			Set objOccNotes = JavaWindow("RDV_StructureManager").JavaWindow("RDV_PSEWindow").JavaDialog("OccurrenceNotes")
		ElseIf  JavaWindow("RDV_StructureManager").JavaWindow("RDV_JApplet").JavaDialog("OccurrenceNotes").Exist(2) Then
			Set objOccNotes = JavaWindow("RDV_StructureManager").JavaWindow("RDV_JApplet").JavaDialog("OccurrenceNotes")
		Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_RDV_OccurrenceNotes ] Dialog [ Occurrence Notes ] does not exists for case [ " & sAction& " ].") 
			Set objOccNotes = Nothing
			Exit Function 
		End If
'	If Fn_UI_ObjectExist("Fn_RDV_OccurrenceNotes", objOccNotes ) = False Then
'		' form attribute window does not exist.
'		Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_RDV_OccurrenceNotes ] Dialog [ Occurrence Notes ] does not exists for case [ " & sAction& " ].") 
'		Set objOccNotes = Nothing
'		Exit function
'	End If	
	' clearing occurrence notes panel
	If bClearOccurrenceNotes <> "" Then
		If cBool(bClearOccurrenceNotes) Then
			Call Fn_Button_Click("Fn_RDV_OccurrenceNotes", objOccNotes, "Clear")
		End If
	End If
	
	'selecting occ note details
	arrOccNotes = split(sOccurrenceNotes, "~")
	arrOperators = split(sOperators, "~")
	arrValues = split(sValues, "~")
	iRowCnt = -1
	For iCnt = 0 to UBound(arrOccNotes)
			' clicking on + button
			Call Fn_Button_Click("Fn_RDV_OccurrenceNotes", objOccNotes, "Add")
			iRowCnt = iRowCnt + 1
			If trim(arrOccNotes(iCnt)) <> "" Then
'				objOccNotes.JavaTable("OccurrenceNotesTable").ClickCell iRowCnt,"Occurrence Notes","LEFT"
				objOccNotes.JavaTable("OccurrenceNotesTable").ClickCell iRowCnt,0,"LEFT"
			Call Fn_List_Select(	"Fn_RDV_OccurrenceNotes", objOccNotes, "OccNotesList", trim(arrOccNotes(iCnt)))
'				objOccNotes.JavaTable("OccurrenceNotesTable").ClickCell iRowCnt,"Occurrence Notes","LEFT"
				objOccNotes.JavaTable("OccurrenceNotesTable").ClickCell iRowCnt,0,"LEFT"
			End If
			' setting operator
			If trim(arrOperators(iCnt)) <> "" Then
				objOccNotes.JavaTable("OccurrenceNotesTable").ClickCell iRowCnt,"Operator","LEFT"
			'						wait 2
			'						objOccNotes.JavaTable("OccurrenceNotesTable").ClickCell iRowCnt,"Operator","LEFT"
			'						If objOccNotes.JavaList("OccNotesList").exist(5) = False then
			'							objOccNotes.JavaTable("OccurrenceNotesTable").ClickCell iRowCnt,"Operator","LEFT"
			'						End If
			'						Call Fn_List_Select("Fn_RDV_OccurrenceNotes", objOccNotes, "OccNotesList", trim(arrOperators(iCnt)))
				objOccNotes.JavaTable("OccurrenceNotesTable").SetCellData iRowCnt,"Operator", trim(arrOperators(iCnt))
			End IF
			' setting values
			If trim(arrValues(iCnt)) <> "" Then
					objOccNotes.JavaTable("OccurrenceNotesTable").ClickCell iRowCnt,"Value","LEFT"
					
					call Fn_SISW_UI_JavaEdit_Operations("Fn_RDV_OccurrenceNotes", "Set", objOccNotes, "OccurrenceNotesValue", trim(arrValues(iCnt)))
			'objOccNotes.JavaTable("OccurrenceNotesTable").SetCellData iRowCnt,"Value", trim(arrValues(iCnt))
			End If
	Next
	
	'clicking on OK
	Call Fn_Button_Click("Fn_RDV_OccurrenceNotes", objOccNotes, "OK")
	Fn_RDV_OccurrenceNotes = True
	Set objOccNotes = Nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_RDV_OccurrenceNotes
'@@
'@@    Description				 :	Function Used to perform search operation using Classifications dialog
'@@
'@@    Parameters			   :	1.bClearClassificationPanel : to clear Form Attribute table ( True / False / "" )
'@@   	 							2.sSearchClassificationClass : ~ separated list of Relation Types ( String )
'@@   	 							3.sSysOfMeasurement :  ~ separated list of Parent Types ( String )
'@@   	 							4.sPropertyNames :  ~ separated list of Property Names ( String )
'@@   	 							5.sOperators :  ~ separated list of Operators ( String )
'@@   	 							6.sValues :  ~ separated list of Values ( String )
'@@
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	Classifications dialog should be displayed						
'@@
'@@    Examples					:	
'@@    								bClear = "true"
'@@    								bClearClassificationPanel = "true"
'@@    								sSearchClassificationClass = "Classification Root:sc1 [1]"
'@@									sSysOfMeasurement = "non-metric"
'@@									sPropertyNames = "sc1.Measure~sc1.Measure"
'@@    								sOperators = "=~>"
'@@    								sValues = "1~2"
'@@    								bClickOnSearchButton = ""
'@@    								msgbox Fn_RDV_Classifications( bClearClassificationPanel, sSearchClassificationClass, sSysOfMeasurement, sPropertyNames, sOperators, sValues)
'@@
'@@	   History					 	:	
'@@				Developer Name				Date			Rev. No.	Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe			10-Jan-2012			1.0			Created
'@@				Ashok kakade			29-June-2012		1.0			Added New Hierarchy of Dialog Classification	
'@@				Avinash J			19-Mar-2013		1.1			Modified code to search node to select
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_RDV_Classifications( bClearClassificationPanel, sSearchClassificationClass, sSysOfMeasurement, sPropertyNames, sOperators, sValues)
	GBL_FAILED_FUNCTION_NAME="Fn_RDV_Classifications"
	Dim objClassification, objSelectType, intNoOfObjects
	Dim arrPropertyNames, arrOperators, arrValues, iCnt, iRowCnt,bSelect,bSelect2

	If  JavaWindow("RDV_StructureManager").JavaWindow("RDV_PSEWindow").JavaDialog("Classification").Exist(2) Then
			Set objClassification = JavaWindow("RDV_StructureManager").JavaWindow("RDV_PSEWindow").JavaDialog("Classification")
	ElseIf  JavaWindow("RDV_StructureManager").JavaWindow("RDV_JApplet").JavaDialog("Classification").Exist(2) Then
			Set objClassification = JavaWindow("RDV_StructureManager").JavaWindow("RDV_JApplet").JavaDialog("Classification")
	Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_RDV_Classifications ] Dialog [ Occurrence Notes ] does not exists for case [ " & sAction& " ].") 
			Set objClassification = Nothing
			Fn_RDV_Classifications = False
			Exit function
	End If
'	If Fn_UI_ObjectExist("Fn_RDV_Classifications", objClassification ) = False Then
'			' form attribute window does not exist.
'			Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_RDV_Classifications ] Dialog [ Occurrence Notes ] does not exists for case [ " & sAction& " ].") 
'			Set objClassification = Nothing
'			Exit function
'	End If
	
	' clearing classification panel
	If bClearClassificationPanel <> "" Then
		If cBool(bClearClassificationPanel) Then
			Call Fn_Button_Click("Fn_RDV_Classifications", objClassification, "Clear")
		End If
	End If
	
	'selecting type of class
	If sSearchClassificationClass <> "" Then
		Call Fn_CheckBox_Select("Fn_RDV_Classifications",objClassification,"ClassificationClassSelect")
        Set objSelectType = description.Create()             'bellow code is added to Expand the required noed first                  by Avinash j.
		objSelectType("Class Name").value = "JavaEdit"
		'objSelectType("attached text").value = "Class/Attribute Selection Popup"
		'objSelectType("tagname").value = "Class/Attribute Selekction Popup"
		objSelectType("toolkit class").value="com.teamcenter.rac.aif.common.AIFTree\$FindInDisplayPanel\$1"
		wait 1
		bSelect=split(sSearchClassificationClass, ":")                      
		bSelect2=split( bSelect(UBound(bSelect))," ")                                        
		Set  intNoOfObjects = JavaWindow("RDV_StructureManager").ChildObjects(objSelectType)
				'If intNoOfObjects(0).exist(3) then
				 intNoOfObjects(0).Set  bSelect2(0)
				Call Fn_KeyBoardOperation("SendKeys", "{ENTER}")
				'end If
		Set objSelectType = Nothing
		Set  intNoOfObjects = Nothing
		wait 1
		Err.Clear
		Set objSelectType = description.Create()
		objSelectType("Class Name").value = "JavaTree"
		'objSelectType("attached text").value = "Class/Attribute Selection Popup"
		'objSelectType("tagname").value = "Class/Attribute Selekction Popup"
		objSelectType("toolkit class").value="com.teamcenter.rac.common.pomclasstree.POMICSTree"
		Set  intNoOfObjects = JavaWindow("RDV_StructureManager").ChildObjects(objSelectType)
		wait 1
		intNoOfObjects(0).select sSearchClassificationClass
		intNoOfObjects(0).Activate sSearchClassificationClass
		wait 5
			If Err.Number < 0  Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_RDV_Classifications ] Failed to find Classification Class Tree.") 
				Set objSelectType = Nothing
				Set  intNoOfObjects = Nothing
				Set objClassification = Nothing
				Exit function
			Else
           		' objClassification.JavaCheckBox("ClassificationClassSelect").DblClick 1,1,"LEFT"
	     		Set objSelectType = Nothing
		   	    Set  intNoOfObjects = Nothing
			End If
	End If 
	'selecting type of class
	If sSysOfMeasurement <> "" Then
		Call Fn_CheckBox_Select("Fn_RDV_Classifications",objClassification,"SetSystemOfMeasurement")
		objClassification.JavaMenu("label:=" & sSysOfMeasurement ).select
	End If
	
	'selecting classification details
	arrPropertyNames = split(sPropertyNames, "~")
	arrOperators = split(sOperators, "~")
	arrValues = split(sValues, "~")
	iRowCnt = -1
	For iCnt = 0 to UBound(arrPropertyNames)
				  ' clicking on + button
		If iCnt <> 0 Then
			Call Fn_Button_Click("Fn_RDV_Classifications", objClassification, "Add")
		End If
		iRowCnt = iRowCnt + 1
		If trim(arrPropertyNames(iCnt)) <> "" Then
			objClassification.JavaTable("SearchClassificationTable").ClickCell iRowCnt,"Property Name","LEFT"
			Call Fn_List_Select(	"Fn_RDV_Classifications", objClassification, "SearchClassificationList", trim(arrPropertyNames(iCnt)))
		End If
		' setting operator
		If trim(arrOperators(iCnt)) <> "" Then
			objClassification.JavaTable("SearchClassificationTable").ClickCell iRowCnt,  2,"LEFT"
			Call Fn_List_Select(	"Fn_RDV_Classifications", objClassification, "SearchClassificationList", trim(arrOperators(iCnt)))
		End If
		' setting values
		If trim(arrValues(iCnt)) <> "" Then
			objClassification.JavaTable("SearchClassificationTable").SetCellData iRowCnt,"Searching Value", trim(arrValues(iCnt))
		End If
	Next
	
	'clicking on OK
	Call Fn_Button_Click("Fn_RDV_Classifications", objClassification, "OK")
	Fn_RDV_Classifications = True
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_RDV_SpatialCriteria
'@@
'@@    Description				 :	Function Used to perform search operation on Spatial Criteria dialog
'@@
'@@    Parameters			   :	s3DBoxCoordinates = 3D Box Coordinates  to be selected from list.
'@@    											sSearchType = Search Type
'@@    											XCoord =  X Coordinates
'@@    											XLen = X Length
'@@    											YCoord =  Y Coordinates 
'@@    											YLen = Y Length
'@@    											ZCoord =  Z Coordinates
'@@    											ZLen = Z Length
'@@    											bTrueShapeFiltering = boolean value to select checkbox True Shape Filtering
'@@    											sDistance = distance
'@@    											bCenterToSelected = True / False
'@@
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	Spatial Criteria dialog should be displayed							
'@@
'@@    Examples					:	s3DBoxCoordinates = 3D Box Coordinates  to be selected from list.
'@@    											sSearchType = Search Type
'@@    											XCoord =  X Coordinates
'@@    											XLen = X Length
'@@    											YCoord =  Y Coordinates 
'@@    											YLen = Y Length
'@@    											ZCoord =  Z Coordinates
'@@    											ZLen = Z Length
'@@    											bTrueShapeFiltering = boolean value to select checkbox True Shape Filtering
'@@    											sDistance = distance
'@@    											bCenterToSelected = True / False
'@@
'@@	   History					 	:	
'@@				Developer Name				Date			Rev. No.	Changes Done													Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe			28-Feb-2012			1.0			Created
'@@				Ashok kakade			29-June-2012		1.0			Added New Hierarchy of Dialog SpatialCriteria
'@@				Shrikant Narkhede		02-July-2012		1.0			modify function for case "3D Box" & "Proximity"
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Pallavi Jadhav			19-Sept-2013		2.0			redirected function to Fn_SISW_RDV_SpatialFilterOperations		Koustubh
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_RDV_SpatialCriteria(s3DBoxCoordinates, sSearchType, XCoord, XLen, YCoord, YLen, ZCoord, ZLen, bTrueShapeFiltering, sDistance, bCenterToSelected)
	GBL_FAILED_FUNCTION_NAME="Fn_RDV_SpatialCriteria"
	Dim dicSpatialFilter
	Set dicSpatialFilter = CreateObject( "Scripting.Dictionary" )
	dicSpatialFilter("SearchType") = sSearchType
	If sSearchType = "3D Box" Then
		dicSpatialFilter("Extent") = s3DBoxCoordinates
		dicSpatialFilter("XMin") = XCoord
		dicSpatialFilter("XMax") = XLen
		dicSpatialFilter("YMin") = YCoord
		dicSpatialFilter("YMax") = YLen
		dicSpatialFilter("ZMin") = ZCoord
		dicSpatialFilter("ZMax") = ZLen
	ElseIf sSearchType = "Proximity" Then
		dicSpatialFilter("Distance") = sDistance
		'bCenterToSelected  : Not Implemented
	End If
	If bTrueShapeFiltering <> "" Then
		If cBool(bTrueShapeFiltering) Then
			dicSpatialFilter("TrueShapeFiltering") = "ON"
		Else
			dicSpatialFilter("TrueShapeFiltering") = "OFF"
		End If
	End IF
	' clicking on OK
	dicSpatialFilter("ButtonName") = "OK"
	Fn_RDV_SpatialCriteria = Fn_SISW_RDV_SpatialFilterOperations("Set", dicSpatialFilter)
	Set dicSpatialFilter = Nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_RDV_ItemIDSearchPanelOperations
'@@
'@@    Description				 :	Function Used to perform search operation using Item ID Search Criteria
'@@
'@@    Parameters			   :	1. sAction: Action to be performed
'@@    								2. dicItemIDSearch: dictionary object
'@@
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	Structure Manager should be displayed							
'@@
'@@    Examples					:	dicItemIDSearch("bChangeSearch") = True
'@@    								dicItemIDSearch("AdvancedDefaultSearchType") = "Item..."
'@@    								dicItemIDSearch("bClearHistory") = "true"
'@@    								dicItemIDSearch("RememberMyLastSearches") = "10"
'@@    								dicItemIDSearch("SearchType") = "Item..."
'@@    								dicItemIDSearch("SearchCriteria") = "Item ID=000038~Name=comp1"
'@@    								dicItemIDSearch("bClickOnSearchButton") = ""
'@@    								Call Fn_RDV_ItemIDSearchPanelOperations("Search", dicItemIDSearch)
'@@
'@@	   History					 	:	
'@@				Developer Name				Date			Rev. No.	Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe			29-Nov-2011			1.0			Created
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public function Fn_RDV_ItemIDSearchPanelOperations(sAction, dicItemIDSearch)
	GBL_FAILED_FUNCTION_NAME="Fn_RDV_ItemIDSearchPanelOperations"
	Dim objApplet, bReturn
				
	Fn_RDV_ItemIDSearchPanelOperations = False
	Set objApplet = JavaWindow("RDV_StructureManager").JavaWindow("RDV_JApplet")
	' opening serach panel if it is not displayed.
	If Fn_UI_ObjectExist("Fn_RDV_ItemIDSearchPanelOperations", objApplet.JavaButton("ItemAttributesCashLessSearch")) = False Then
		bReturn = Fn_ToolBarOperation("Click", "Show/Hide Structure Manager Search Panel","")
		If bReturn <> True Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_RDV_ItemIDSearchPanelOperations ] Failed to click on [ Show/Hide Structure Manager Search Panel ].") 
				Fn_RDV_ItemIDSearchPanelOperations = False
				Exit function 
		Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Function [ Fn_RDV_ItemIDSearchPanelOperations ] Successfully clicked on [ Show/Hide Structure Manager Search Panel ].") 
		End If
	End If

	Select Case sAction
		Case "Search"
			If dicItemIDSearch("bClear") = "" then dicItemIDSearch("bClear") = "False"
			If cBool(dicItemIDSearch("bClear")) then
				Call Fn_Button_Click("Fn_RDV_ItemIDSearchPanelOperations", objApplet,"SearchPanelClear_16Button")
			End If
			' clicking on ... button of ItemID search criteria
			Call Fn_Button_Click("Fn_RDV_ItemIDSearchPanelOperations", objApplet, "ItemAttributesCashLessSearch")
			
			' Item Attribut function call
			IF Fn_RDV_ItemAttributes(dicItemIDSearch) = False then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_RDV_ItemAttributes ] Failed to open [ Item Attributes ] dialog.") 
				Exit function
			End IF
			
			' clicking on search button
			If dicItemIDSearch("bClickOnSearchButton") = "" then dicItemIDSearch("bClickOnSearchButton") = True
			If cBool(dicItemIDSearch("bClickOnSearchButton")) Then
				Call Fn_Button_Click("Fn_RDV_ItemIDSearchPanelOperations", objApplet,"SearchPanelSearch_16Button")
			End If
			Fn_RDV_ItemIDSearchPanelOperations = True
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_RDV_ItemIDSearchPanelOperations ] Invalid case [ " & sAction& " ].") 
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	End Select
	If Fn_RDV_ItemIDSearchPanelOperations = True Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Function [ Fn_RDV_ItemIDSearchPanelOperations ] Executed successfully with case [ " & sAction& " ].") 
	End If
	Set objApplet = Nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_RDV_SearchResultsOperations
'@@
'@@    Description			:	Function Used to Call Web Menu's
'@@
'@@    Parameters			:	1.strMenu: Menu Name
'@@
'@@    Return Value		   	: 	True Or False
'@@
'@@    Pre-requisite		:	Should be Log In Web Client							
'@@
'@@    Examples				:	Call Fn_RDV_SearchResultsOperations("VerifyInColumn", "Result 2", "", "Item Id", "000038", "")
'@@    Examples				:	Call Fn_RDV_SearchResultsOperations("SelectAllAndDisplay", "Result 2", "", "", "", True)
'@@    Examples				:	Call Fn_RDV_SearchResultsOperations("SelectAndDisplay", "Result 2", "Comp1", "Item Id", "", False)
'@@    Examples				:	Call Fn_RDV_SearchResultsOperations("MultiAndDisplay", "Result 2", "Comp1~comp2", "Item Id", "", False)
'@@    Examples				:	Call Fn_RDV_SearchResultsOperations("SelectAllAndDisplayInNewWindow", "Result 2", "", "", "", True)
'@@    Examples				:	Call Fn_RDV_SearchResultsOperations("SelectAndDisplayInNewWindow", "Result 2", "Comp1", "Item Id", "", False)
'@@    Examples				:	Call Fn_RDV_SearchResultsOperations("MultiAndDisplayInNewWindow", "Result 2", "Comp1~comp2", "Item Id", "", False)
'@@    Example              :    Call Fn_RDV_SearchResultsOperations("SelectAndCopy", "Result 2", "Comp1", "", "", False)
'@@    Example              :    Call Fn_RDV_SearchResultsOperations("VerifyInTabPopupMenuExist", "Result 2", "", "Rename", "", False)
'@@    Example              :    Call Fn_RDV_SearchResultsOperations("SelectInTabPopupMenu", "Result 2", "", "Rename", "", False)
'@@    Example              :    Call Fn_RDV_SearchResultsOperations("FindInResults", "Result 2", "Item Name", "!=", "sub2", False)
'@@    Example              :    Call Fn_RDV_SearchResultsOperations("GetSelected", "", "Item Name", "", "", False)
'@@    Example              :    Call  Fn_RDV_SearchResultsOperations("VerifyRow", "","Wheel_QS_2-000084@2", "","", "")
'@@    Example              :    Call  Fn_RDV_SearchResultsOperations("VerifyRow","","comp-000060","Parent~Preferred Ancestor","000059/A;1-top (View)~0000160-comp","False")
'@@    Example              :    Call  Fn_RDV_SearchResultsOperations("VerifyRow","","comp-123~000060","Parent~Preferred Ancestor","000059/A;1-top (View)~0000160-comp","False")
'@@    Examples				:	Call Fn_RDV_SearchResultsOperations("VerifyInColumn_InTargetsTable", "", "", "Item Name", "0632-039", "" )
'@@    Examples				:	Call Fn_RDV_SearchResultsOperations("InsertColumn", "", "", "Item Name", "", "")
'@@    Examples				:	Call Fn_RDV_SearchResultsOperations("InsertColumnAtPosition", "", "", "Item Name:0", "", "")
'@@	   Examples					:	Call= Fn_RDV_SearchResultsOperations("VerifyColumnName", "","", "BOM Line~Parent", "","")
'@@	
'@@	   History				:	
'@@				Developer Name				Date			Rev. No.	Changes Done													Reviewer							Tc Release		
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe			29-Nov-2011			1.0			Created
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe			29-Dec-2011			1.0			Added case MultiAndDisplayInNewWindow, 
'@@																		     	   SelectAllAndDisplayInNewWindow, 
'@@																				   SelectAndDisplayInNewWindow
'@@----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@            Vrushali Wani			30-Dec-2011			1.0			Added case SelectAndCopy
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe			05-Jan-2012			1.0			Added exit criteria for loop of SearchCompleted
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe			03-Feb-2012			1.0			Added new cases "VerifyInTabPopupMenuExist",
'@@																						"SelectInTabPopupMenu","FindInResults"
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Naveen Gupta			20-Mar-2012			1.0			Added new cases "GetSelected"									Koustubh Watwe
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe			29-Mar-2012			1.0			Added new cases "VerifyRow"
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe			2-Apr-2012			1.0			Modified case "VerifyRow"
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Amit Talegaonkar		11-Sept-2012	    1.0		    Added code to handele "Targets Table" 
'@@																		And case "VerifyInColumn_InTargetsTable"
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Ankit Nigam				27-Oct-2015			1.0			Modified cases "VerifyRow" , "SelectAndDisplay" ,	     		    Vivek A.					Tc1121-20150101200
'@@		     															"SelectAndCopy","MultiSelectAndDisplay",for design change 
'@@																		in Search Result Dialog										
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Vivek Ahirrao			18-Nov-2015			1.1			Modified cases "VerifyRow", "VerifyInColumn", "SelectAndDisplay", 	[TC1121-2015102600-18_11_2015-VivekA-Maintenance]
'@@																		"SelectAndDisplayInNewWindow", "SelectAndCopy", "FindInResults"
'@@																		- Search Results window design change, BOM Line and Parent columns are bydefault there in results tab
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Vivek Ahirrao			26-Nov-2015			1.1			Modified cases "VerifyInColumn", "GetSelected", 	[TC1121-2015102600-18_11_2015-VivekA-Maintenance]
'@@																		- As per design change
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe			13-Apr-2016			1.0			Added new cases "InsertColumn","InsertColumnAtPosition"				Koustubh Watwe
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Jotiba Takkekar			5-Dec-2017			1.0			Added new cases "VerifyColumnName"									Jotiba T		
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_RDV_SearchResultsOperations(sAction, sTab, sRow, sColumn, sValue, bCloseDialog)
	GBL_FAILED_FUNCTION_NAME="Fn_RDV_SearchResultsOperations"
	Dim objSrchRes, iCnt, iRowCount, aValue, objApplet, aOpr
	Dim aRows, iArrCnt, iWaitCnt, iInstance, iInstCnt , sNodeValue
	Dim arrValSet, aColumns, aRow, arrName, arrId , sName, iCounter, aValueVerify, bFlag1
	Dim sNodeID, sNodeName, sTempColumn, sReturn, aReturn, arrReturn
	Dim iTotalCols, sColumnName, sColName, iColcount
	
	Set objSrchRes = JavaWindow("RDV_StructureManager").JavaWindow("RDV_JApplet").JavaDialog("SearchResults")
	
	If Instr( sAction , "_InTargetsTable" ) > 0 Then
		objSrchRes.JavaTable("SearchResultTable").SetTOProperty "attached text" , "Targets Table:"
		Wait 1
		If objSrchRes.JavaTable("SearchResultTable").Exist(5) = False Then
			'Click on Target Button to open Target table
			Call Fn_Button_Click("Fn_RDV_SearchResultsOperations", objSrchRes,"ShowHideTargetsTable")
			Call Fn_ReadyStatusSync(2)
			
			If objSrchRes.JavaTable("SearchResultTable").Exist(5) = False Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_RDV_SearchResultsOperations ] Can not find [ Target Table ] table.") 
				Exit Function
			End If
		End If
	End If
	
	Fn_RDV_SearchResultsOperations = False
	If bCloseDialog = "" Then bCloseDialog = False
	If Fn_UI_ObjectExist("Fn_RDV_SearchResultsOperations",objSrchRes ) = False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_RDV_SearchResultsOperations ] Can not find winodw [ Search Results ].") 
		Set objSrchRes = Nothing
		Exit function
	End If

	' waiting to complete search process
	iWaitCnt = 1
	Do While NOT(Fn_UI_ObjectExist("Fn_RDV_SearchResultsOperations",objSrchRes.JavaStaticText("SearchCompleted") ))
		wait 1
		If iWaitCnt = 600 then 
			exit Do
		End If
		iWaitCnt = iWaitCnt + 1
	Loop

	' selecting Tab-
	If sTab <> "" Then
		Call Fn_UI_JavaTab_Select("Fn_RDV_SearchResultsOperations",objSrchRes,"ResultsTab", sTab)
	End If
	
	Select Case sAction
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "VerifyRow"
				iRowCount = cInt(objSrchRes.JavaTable("SearchResultTable").GetROProperty("rows"))
				iInstCnt = 1
				For iCnt = 0 to iRowCount -1
					If inStr(sRow,"~") > 0 Then
						arrValSet = split(sRow,"~")
					Else
						arrValSet = split(sRow,"-")
					End IF
					
					aRow = split(arrValSet(1) ,"@")
					if uBound(aRow) = 1 then
						iInstance = cInt(aRow(1))
						arrValSet(1) = trim(aRow(0))
					Else
						arrValSet(1) = trim(aRow(0))
						iInstance = 1
					End if
					
					arrId =  split(objSrchRes.JavaTable("SearchResultTable").Object.getValueAt(iCnt,0).toString() , "/")
					arrName = split (arrId(1) , "-" )
					
					If  cStr(arrValSet(0)) = cStr(arrValSet(1))   Then
						arrName(0) =  arrId(0)
					Else 						
						If instr( arrName(1) , " (View)") > 0 Then
							arrName(0) = replace(arrName(1) , " (View)", "")
						Else 
							arrName(0) =  arrName(1) 						
						End If					
					End If
										
						If arrId(0) = cStr(arrValSet(1)) and  arrName(0) = cStr(arrValSet(0)) Then	
							If iInstance = iInstCnt then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Function [ Fn_RDV_SearchResultsOperations ] Successfully verified [ " & cStr(arrValSet(0))  & "-" & cStr(arrValSet(1) ) & " ] is present.")
								Fn_RDV_SearchResultsOperations = True
										If sColumn <> ""  and sColumn <> "Preferred Ancestor" Then					'Tc1121-20150101200-27_10_2015-AnkitN-HC_Maintenance-Modified case for design change in Search Result Dialog
											Fn_RDV_SearchResultsOperations = False
											aColumns = split(sColumn,"~")
											aValue = split(sValue,"~")
											For iArrCnt = 0 to uBound(aColumns)
												If aColumns(iArrCnt) = "Preferred Ancestor" Then
													aValueVerify = Split(aValue(iArrCnt),"-")
													For iCounter = 0 To UBound(aValueVerify)
														bFlag1 = False
														iTotalCols = objSrchRes.JavaTable("SearchResultTable").GetROProperty("cols")
														For iColcount = 0 to iTotalCols-1
															sColumnName = objSrchRes.JavaTable("SearchResultTable").GetColumnName(iColcount)
															If sColumnName = "BOM Line" OR sColumnName = "BOM Line Name" Then
																sColName = sColumnName
																Exit for
															End If
														Next
														If Instr(trim(Cstr(trim(objSrchRes.JavaTable("SearchResultTable").GetCellData(iCnt, sColName)))) , trim(cStr(aValueVerify(iCounter))))>0 then
															bFlag1 = True
														Else 
															bFlag1 = False
															Exit For
														End If
													Next
													If bFlag1 = False Then
														Fn_RDV_SearchResultsOperations = False
														Exit For
													End If
													Fn_RDV_SearchResultsOperations = True
												Else
													If trim(Cstr(trim(objSrchRes.JavaTable("SearchResultTable").GetCellData(iCnt, aColumns(iArrCnt))))) <> trim(cStr(aValue(iArrCnt))) then
														Fn_RDV_SearchResultsOperations = False
														Exit for
													End If
													Fn_RDV_SearchResultsOperations = True
												End If
											Next
										End If
										If Fn_RDV_SearchResultsOperations = False Then
											Exit for
										End If
							End If
									iInstCnt = iInstCnt + 1
						End If		
				Next
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "InsertColumn", "InsertColumnAtPosition"
			Fn_RDV_SearchResultsOperations = False 
			dim objReportTable, objColumns
			Dim bFlag
			
			set objReportTable = JavaWindow("RDV_StructureManager").JavaWindow("RDV_JApplet").JavaDialog("SearchResults").JavaTable("SearchResultTable")
			'Set objColumns = JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Change Columns")
			
			Set objColumns = Window("TeamcenterWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Change Columns")

			aColumns = split(sColumn,":")
			iCol  = Fn_SISW_UI_JavaTable_Operations("Fn_RDV_SearchResultsOperations", "GetColumnIndex", objReportTable , "", "", aColumns(0), "", "", "", "", "")
			If iCol = -1 Then
				objReportTable.SelectColumnHeader 0, "RIGHT"
				wait SISW_MICRO_TIMEOUT
				set var = Description.Create()
				var("Class Name").value = "JavaMenu"
				set childObjects = JavaWindow("RDV_StructureManager").JavaWindow("RDV_JApplet").JavaDialog("SearchResults").ChildObjects(var)
				If childObjects.count <> 0 then
					JavaWindow("RDV_StructureManager").JavaWindow("RDV_JApplet").JavaDialog("SearchResults").JavaMenu("label:=Insert column.*").Select
				Else
					Exit function
				End If
				
				If Fn_SISW_UI_Object_Operations("Fn_RDV_SearchResultsOperations","Enabled", objColumns,"") Then
					Call Fn_SISW_UI_JavaList_Operations("Fn_RDV_SearchResultsOperations", "Select", objColumns,"ListAvailableCols",aColumns(0), "", "")
					wait 1
					Call Fn_SISW_UI_JavaButton_Operations("Fn_RDV_SearchResultsOperations", "Click", objColumns,"Add")
				else
					Exit function
				End If
			End If
			
			' manage location
			iCol  = Fn_SISW_UI_JavaList_Operations("Fn_RDV_SearchResultsOperations", "GetIndex", objColumns,"ListDisplayedCols", aColumns(0), "", "")
			
			If sAction = "InsertColumnAtPosition" Then
				If sAction = "InsertColumnAtPosition" AND iCol <> cInt(aColumns(1)) Then
					If Fn_SISW_UI_Object_Operations("Fn_RDV_SearchResultsOperations","Enabled", objColumns,"") = false Then
						objReportTable.SelectColumnHeader 0, "RIGHT"
						wait SISW_MICRO_TIMEOUT
						set var = Description.Create()
						var("Class Name").value = "JavaMenu"
						set childObjects = JavaWindow("RDV_StructureManager").JavaWindow("RDV_JApplet").JavaDialog("SearchResults").ChildObjects(var)
						If childObjects.count <> 0 then
							JavaWindow("RDV_StructureManager").JavaWindow("RDV_JApplet").JavaDialog("SearchResults").JavaMenu("label:=Insert column.*").Select
						Else
							Exit function
						End If
					End if
					
					bFlag = true
					iInstance = iCol - cInt(aColumns(1))
					If iInstance < 0 Then
						bFlag = false
						iInstance = iInstance * -1
					End If
					call Fn_SISW_UI_JavaList_Operations("Fn_RDV_SearchResultsOperations", "Select", objColumns,"ListDisplayedCols", aColumns(0), "", "")
					For iCnt = 0 To iInstance-1 Step 1
						If bFlag Then
							Call Fn_SISW_UI_JavaButton_Operations("Fn_RDV_SearchResultsOperations", "Click", objColumns,"Up")
						else
							Call Fn_SISW_UI_JavaButton_Operations("Fn_RDV_SearchResultsOperations", "Click", objColumns,"Down")
						End IF
					Next
				
				End If
			End If

			If Fn_SISW_UI_Object_Operations("Fn_RDV_SearchResultsOperations","Enabled", objColumns,"") Then
				Call Fn_SISW_UI_JavaButton_Operations("Fn_RDV_SearchResultsOperations", "Click", objColumns,"Apply")
				Call Fn_SISW_UI_JavaButton_Operations("Fn_RDV_SearchResultsOperations", "Click", objColumns,"Cancel")
			End If
			Fn_RDV_SearchResultsOperations = true
		
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "VerifyInColumn" , "VerifyInColumn_InTargetsTable"
				iRowCount = cInt(objSrchRes.JavaTable("SearchResultTable").GetROProperty("rows"))
				iInstCnt = 1
				aValue = split(sValue,"@")
				if uBound(aValue) = 1 then
					iInstance = cInt(aValue(1))
					sValue = trim(aValue(0))
				Else
					sValue = trim(aValue(0))
					iInstance = 1
				End if
				
				For iCnt = 0 to iRowCount -1
					sNodeValue = objSrchRes.JavaTable("SearchResultTable").GetCellData(iCnt, 0)	
					If sColumn = "Parent" Then
						sNodeValue = objSrchRes.JavaTable("SearchResultTable").GetCellData(iCnt,sColumn)	
						sName = sNodeValue
					Else
						If instr( sNodeValue , " (View)") > 0 Then
							sNodeValue = replace(sNodeValue , " (View)", "")
						End If
						arrId =  split(sNodeValue , "/")
									
						If sColumn = "Item Id" Then
							sName = arrId(0)
						ElseIf sColumn = "Item Name" Then
							If Instr(arrId(1),"-")>0 Then
								arrName = split(arrId(1),"-")
								sName = arrName(1) 
							Else
								sName = arrId(0)
							End If
						ElseIf sColumn = "Preferred Ancestor" Then
							sNodeID = arrId(0)
							If Instr(arrId(1),"-")>0 Then
								arrName = split(arrId(1),"-")
								sNodeName = arrName(1)
							Else
								sNodeName = arrId(0)
							End If
							sName = sNodeID+"-"+sNodeName
							If Cstr(trim(sName)) <> cStr(sValue) Then
								If Instr(Replace(LCase(sName)," ",""),LCase(sValue)) Then
									sName = sValue
								End If
							End If
						End If	
					End If
					If Cstr(trim(sName)) = cStr(sValue) then
						iF iInstCnt = iInstance Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Function [ Fn_RDV_SearchResultsOperations ] Successfully verified [ " & sValue &" ] is present in column [ " & sColumn & " ].") 
							Fn_RDV_SearchResultsOperations = True
							Exit for
						End If
						iInstCnt = iInstCnt + 1
					end if
				Next
				If Fn_RDV_SearchResultsOperations = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Function [ Fn_RDV_SearchResultsOperations ] Successfully verified [ " & sValue &" ] is not present in column [ " & sColumn & " ].") 
				End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 	
		Case "VerifyColumnName"   'TC11.4-2017111300-5_Dec_2017-JotibaT-Development
					For iCnt = 0 To UBound(Split(sColumn,"~"))
						bFlag=False
						iTotalCols=objSrchRes.JavaTable("SearchResultTable").Object.getColumnCount
							For iColcount = 0 To iTotalCols-1
								If LCase(Trim(objSrchRes.JavaTable("SearchResultTable").GetColumnName(iColcount)))=Lcase(Trim(Split(sColumn,"~") (iCnt))) Then 
									bFlag=True
									Fn_RDV_SearchResultsOperations=True
									Exit For
								End If 
							Next
							If bFlag=False Then
								Fn_RDV_SearchResultsOperations=False
								Exit For
							End If
					Next
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "SelectAndDisplay", "SelectAndDisplayInNewWindow","SelectAndCopy"
				iRowCount = cInt(objSrchRes.JavaTable("SearchResultTable").GetROProperty("rows"))
				For iCnt = 0 to iRowCount -1
					sNodeValue = objSrchRes.JavaTable("SearchResultTable").GetCellData(iCnt, 0)				'Tc1121-20150101200-27_10_2015-AnkitN-HC_Maintenance-Modified case for design change in Search Result Dialog
					If instr( sNodeValue , " (View)") > 0 Then
						sNodeValue = replace(sNodeValue , " (View)", "")
					End If
					arrId =  split(sNodeValue , "/")
					If sColumn = "Item Id" Then
						sName = arrId(0)
					Else
						If instr( arrId(1) , "-") > 0 Then
							arrName = split (arrId(1) , "-" )
						Else 
							arrName = arrId(0)						
						End If
						
						If instr( arrId(1) , "-") > 0 Then
							sName =  arrName(1)
						Else 						
							sName =  arrName
						End If					
					End If
					
					If trim(sName) = sRow then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Function [ Fn_RDV_SearchResultsOperations ] Successfully selected [ " & sRow &" ] in column [ " & "Item Name" & " ].") 
						objSrchRes.JavaTable("SearchResultTable").selectRow iCnt
						Fn_RDV_SearchResultsOperations = True
						Exit for
					End if
				Next   
				If sAction = "SelectAndDisplayInNewWindow" Then
					Call Fn_Button_Click("Fn_RDV_SearchResultsOperations", objSrchRes,"DisplayInNewWindow")
				Elseif sAction = "SelectAndCopy" Then
				      Call Fn_Button_Click("Fn_RDV_SearchResultsOperations", objSrchRes,"Copy")
				Else
					Call Fn_Button_Click("Fn_RDV_SearchResultsOperations", objSrchRes,"Display")
				End IF
				If JavaDialog("Confirmation").Exist(5) Then
					Call Fn_Button_Click("Fn_RDV_SearchResultsOperations", JavaDialog("Confirmation"),"Yes")
				End If

		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "MultiSelectAndDisplay", "MultiSelectAndDisplayInNewWindow"
				aRows = split(sRow,"~")
				iRowCount = cInt(objSrchRes.JavaTable("SearchResultTable").GetROProperty("rows"))
				For iArrCnt = 0 to uBound(aRows)
					For iCnt = 0 to iRowCount -1
						sNodeValue = objSrchRes.JavaTable("SearchResultTable").GetCellData(iCnt, 0)				'Tc1121-20150101200-27_10_2015-AnkitN-HC_Maintenance-Modified case for design change in Search Result Dialog
						If instr( sNodeValue , " (View)") > 0 Then
							sNodeValue = replace(sNodeValue , " (View)", "")
						End If
						arrId =  split(sNodeValue , "/")
						If instr( arrId(1) , "-") > 0 Then
							arrName = split (arrId(1) , "-" )
						Else 
							arrName = split (arrId(1) , "" )						
						End If
						
						If  instr( arrId(1) , "-") > 0 Then
							sName =  arrName(1)
						Else 						
							sName =  arrName(0)
						End If						
						If trim(sName) = aRows(iArrCnt) then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Function [ Fn_RDV_SearchResultsOperations ] Successfully selected [ " & sRow &" ] in column [ " & "Item Name" & " ].") 
							objSrchRes.JavaTable("SearchResultTable").ExtendRow iCnt
							Fn_RDV_SearchResultsOperations = True
'							Exit for
						end if
					Next
				Next
				If sAction = "MultiSelectAndDisplayInNewWindow" Then
					Call Fn_Button_Click("Fn_RDV_SearchResultsOperations", objSrchRes,"DisplayInNewWindow")
				Else
					Call Fn_Button_Click("Fn_RDV_SearchResultsOperations", objSrchRes,"Display")
				End IF
				If JavaDialog("Confirmation").Exist(5) Then
					Call Fn_Button_Click("Fn_RDV_SearchResultsOperations", JavaDialog("Confirmation"),"Yes")
				End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
        Case "SelectAllAndDisplay", "SelectAllAndDisplayInNewWindow"
'				Call Fn_Button_Click("Fn_RDV_SearchResultsOperations", objSrchRes, "SelectAll")
				Call Fn_CheckBox_Set("Fn_RDV_SearchResultsOperations",objSrchRes,"SelectAll", "ON")
				If sAction = "SelectAllAndDisplayInNewWindow" Then
					Call Fn_Button_Click("Fn_RDV_SearchResultsOperations", objSrchRes,"DisplayInNewWindow")
				Else
					Call Fn_Button_Click("Fn_RDV_SearchResultsOperations", objSrchRes,"Display")
				End IF
				If JavaDialog("Confirmation").Exist(5) Then
					Call Fn_Button_Click("Fn_RDV_SearchResultsOperations", JavaDialog("Confirmation"),"Yes")
				End If

				Fn_RDV_SearchResultsOperations = True
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "VerifyInTabPopupMenuExist"
				For iCnt = 0 to cInt(objSrchRes.JavaTab("ResultsTab").GetROProperty("items count"))
					If objSrchRes.JavaTab("ResultsTab").Object.getComponentAt(iCnt).isVisible() Then
						Exit for
					End If
				Next
				objSrchRes.JavaTab("ResultsTab").Click  iCnt * 70 + 35, 10, "RIGHT"
				wait 1
				Fn_RDV_SearchResultsOperations = objSrchRes.JavaMenu("label:=" & sColumn).exist(5)
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "SelectInTabPopupMenu"
				For iCnt = 0 to cInt(objSrchRes.JavaTab("ResultsTab").GetROProperty("items count"))
					If objSrchRes.JavaTab("ResultsTab").Object.getComponentAt(iCnt).isVisible() Then
						Exit for
					End If
				Next
				objSrchRes.JavaTab("ResultsTab").Click  iCnt * 70 + 35, 10, "RIGHT"
				wait 1
				If objSrchRes.JavaMenu("label:=" & sColumn ).exist(5) Then
					objSrchRes.JavaMenu("label:=" & sColumn ).Select
					Fn_RDV_SearchResultsOperations = True
				End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "FindInResults"
			Set objApplet = JavaWindow("RDV_StructureManager").JavaWindow("RDV_JApplet")
			aRows = Split(sRow,"~")
			aOpr = Split(sColumn,"~")
			aValue =  Split(sValue,"~")
			' activate Find In Results panel
			Call Fn_CheckBox_Set("Fn_RDV_SearchResultsOperations",objSrchRes,"FindInResults", "ON")
			If objApplet.JavaDialog("FindInDisplay").Exist(4) = False Then
				If objSrchRes.JavaStaticText("FindinDisplay").Exist(1) = True Then
					objSrchRes.JavaStaticText("FindinDisplay").DblClick 1,1,"LEFT"
					Wait 1
				Else
					Call Fn_CheckBox_Set("Fn_RDV_SearchResultsOperations",objSrchRes,"FindInResults", "ON")
					wait 1
					objSrchRes.JavaStaticText("FindinDisplay").DblClick 1,1,"LEFT"
				End If
			End If
			
			IF objApplet.JavaDialog("FindInDisplay").JavaButton("SearchResultFindInResult_Clear").Exist(10) = False Then
				Call Fn_CheckBox_Set("Fn_RDV_SearchResultsOperations",objSrchRes,"FindInResults", "ON")
			End If
			' click on clear
			wait 5
			IF objApplet.JavaDialog("FindInDisplay").JavaButton("SearchResultFindInResult_Clear").Exist(10) Then
					Call Fn_Button_Click("Fn_RDV_SearchResultsOperations", objApplet.JavaDialog("FindInDisplay"),"SearchResultFindInResult_Clear")
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_RDV_SearchResultsOperations ] Button [ SearchResultFindInResult_Clear ] is not present.") 
				Exit function
			End If
			' set data
			For iArrCnt = 0 to uBound(aRows)
				If aRows(iArrCnt) = "Item Name" OR aRows(iArrCnt) = "Item Id" Then
					aRows(iArrCnt) = "BOM Line Name"
				End If
				IF objApplet.JavaDialog("FindInDisplay").JavaButton("SearchResultFindInResult_Plus").Exist(10) Then
						Call Fn_Button_Click("Fn_RDV_SearchResultsOperations", objApplet.JavaDialog("FindInDisplay"),"SearchResultFindInResult_Plus")
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_RDV_SearchResultsOperations ] Button [ SearchResultFindInResult_Plus ] is not present.") 
					exit function 
				End If
				wait 1
				objApplet.JavaDialog("FindInDisplay").JavaTable("SearchResultFindInResult_Table").SetCellData iArrCnt, "Property Name", aRows(iArrCnt)
				wait 1
				objApplet.JavaDialog("FindInDisplay").JavaTable("SearchResultFindInResult_Table").ClickCell iArrCnt, 1
				objApplet.JavaDialog("FindInDisplay").JavaTable("SearchResultFindInResult_Table").SetCellData iArrCnt,2, aOpr(iArrCnt)
				wait 1
				objApplet.JavaDialog("FindInDisplay").JavaTable("SearchResultFindInResult_Table").SetCellData iArrCnt, "Searching Value", aValue(iArrCnt)
			Next
			' click on Find
			Call Fn_Button_Click("Fn_RDV_SearchResultsOperations", objApplet.JavaDialog("FindInDisplay"),"SearchResultFindInResult_Find")
			wait 1
'			Call Fn_CheckBox_Set("Fn_RDV_SearchResultsOperations",objSrchRes,"FindInResults", "OFF")
			objApplet.JavaDialog("FindInDisplay").Close
			Fn_RDV_SearchResultsOperations = True
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "GetSelected"
'			If  sColumn = "" Then
'				sColumn = "Item Name"
'			End If
			sTempColumn = sColumn
			If sColumn <> "Parent" Then
				sColumn = "BOM Line Name"
			End If
			Fn_RDV_SearchResultsOperations = False
			iRowCount = cInt(objSrchRes.JavaTable("SearchResultTable").GetROProperty("rows"))
			For iCnt =0  to iRowCount - 1 
				If objSrchRes.JavaTable("SearchResultTable").Object.isRowSelected(iCnt) Then
						sReturn = objSrchRes.JavaTable("SearchResultTable").GetCellData(iCnt, sColumn)
						If sTempColumn="Item Id" Then
							aReturn = Split(sReturn,"/")
							sReturn = aReturn(0)
						ElseIf sTempColumn="Item Name" Then
							aReturn = Split(sReturn,"/")
							If Instr(aReturn(1),"-")>0 Then
								arrReturn = Split(aReturn(1),"-")
								If UBound(arrReturn)>1 Then
									sReturn = arrReturn(1)+"-"+arrReturn(2)
								Else
									sReturn = arrReturn(1)
								End If
							Else
								sReturn = aReturn(0)
							End If
						Else
							sReturn = sReturn
						End If
						If Fn_RDV_SearchResultsOperations = False Then
								Fn_RDV_SearchResultsOperations = sReturn
						Else
								Fn_RDV_SearchResultsOperations = Fn_RDV_SearchResultsOperations & "~" & sReturn
						End If
				End if
			Next
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_RDV_SearchResultsOperations ] Invalid case [ " & sAction& " ].") 
	End Select
	If cBool(bCloseDialog) then objSrchRes.Close
	If Fn_RDV_SearchResultsOperations <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Function [ Fn_RDV_SearchResultsOperations ] Executed successfully with case [ " & sAction& " ].") 
	End If
	
	objSrchRes.JavaTable("SearchResultTable").SetTOProperty "attached text" , "Results Table:"
	
	Set objSrchRes = Nothing
	Set objApplet = Nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_RDV_FormAttributesSearchPanelOperations
'@@
'@@    Description				 :	Function Used to perform search operation using Form Attribute Search Criteria
'@@
'@@    Parameters			   :	1.sAction : Action to be performed
'@@   	 							2.bClear : to clear search criteria ( True / False / "" )
'@@   	 							3.bClearFormAttributePanel : to clear Form Attribute table ( True / False / "" )
'@@   	 							4.sRelationType : ~ separated list of Relation Types ( String )
'@@   	 							5.sParentType :  ~ separated list of Parent Types ( String )
'@@   	 							6.sFormType :  ~ separated list of Form Types ( String )
'@@   	 							7.sPropertyName :  ~ separated list of Property Names ( String )
'@@   	 							8.sOperator :  ~ separated list of Operators ( String )
'@@   	 							9.sSearchingValue :  ~ separated list of Seawrching Values ( String )
'@@   	 						   10.bClickOnSearchButton : to perform search operation ( True / False / "" )
'@@
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	Structure Manager should be displayed						
'@@
'@@    Examples					:	
'@@    								bClear = "true"
'@@    								bClearFormAttributePanel = "true"
'@@    								sRelationType = "3D Snapshot~"
'@@    								sParentType = "Item~"
'@@    								sFormType = "Arc Override Form~"
'@@    								sPropertyName = "Start Point RZ~"
'@@    								sOperator = "EQ~"
'@@    								sSearchingValue = "1~2"
'@@    								bClickOnSearchButton = ""
'@@    								sAction = "Search"
'@@    								msgbox  Fn_RDV_FormAttributesSearchPanelOperations("Search", bClear, bClearFormAttributePanel, sRelationType, sParentType, sFormType, sPropertyName, sOperator, sSearchingValue, bClickOnSearchButton)
'@@
'@@	   History					 	:	
'@@				Developer Name				Date			Rev. No.	Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe			1-Dec-2011			1.0			Created
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public function Fn_RDV_FormAttributesSearchPanelOperations(sAction, bClear, bClearFormAttributePanel, sRelationType, sParentType, sFormType, sPropertyName, sOperator, sSearchingValue, bClickOnSearchButton)
	GBL_FAILED_FUNCTION_NAME="Fn_RDV_FormAttributesSearchPanelOperations"
	Dim objApplet

	Fn_RDV_FormAttributesSearchPanelOperations = False
	Set objApplet = JavaWindow("RDV_StructureManager").JavaWindow("RDV_JApplet")
	' opening serach panel if it is not displayed.
	If Fn_UI_ObjectExist("Fn_RDV_FormAttributesSearchPanelOperations", objApplet.JavaButton("FormAttributesCashLessSearch")) = False Then
		bReturn = Fn_ToolBarOperation("Click", "Show/Hide Structure Manager Search Panel","")
		If bReturn <> True Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_RDV_FormAttributesSearchPanelOperations ] Failed to click on [ Show/Hide Structure Manager Search Panel ].") 
			Set objApplet = Nothing
			Exit function 
		Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Function [ Fn_RDV_FormAttributesSearchPanelOperations ] Successfully clicked on [ Show/Hide Structure Manager Search Panel ].") 
		End If
	End If
	
	If sParentType = "ItemRevision" Then sParentType = "Item Revision"
	
	Select Case sAction
		Case "Search"
				If bClear = "" then bClear = "False"
				If cBool(bClear) then
					Call Fn_Button_Click("Fn_RDV_FormAttributesSearchPanelOperations", objApplet,"SearchPanelClear_16Button")
				End If
	
				' clicking on ... button of Form Attribute search criteria
				Call Fn_Button_Click("Fn_RDV_FormAttributesSearchPanelOperations", objApplet, "FormAttributesCashLessSearch")
	
				If Fn_RDV_FormAttributes(bClearFormAttributePanel, sRelationType, sParentType, sFormType, sPropertyName, sOperator, sSearchingValue) = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_RDV_FormAttributesSearchPanelOperations ] Failed to execute function [ Fn_RDV_FormAttributes ].") 
					Exit Function
				End If
				' clicking on search button
				If bClickOnSearchButton = "" then bClickOnSearchButton = True
				If cBool(bClickOnSearchButton) Then
					Call Fn_Button_Click("Fn_RDV_FormAttributesSearchPanelOperations", objApplet,"SearchPanelSearch_16Button")
				End If
				Fn_RDV_FormAttributesSearchPanelOperations = True

		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_RDV_FormAttributesSearchPanelOperations ] Invalid case [ " & sAction& " ].") 
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	End Select

	If Fn_RDV_FormAttributesSearchPanelOperations = True Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Function [ Fn_RDV_FormAttributesSearchPanelOperations ] Executed successfully with case [ " & sAction& " ].") 
	End If
	Set objApplet = Nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_RDV_OccurrenceNotesSearchPanelOperations
'@@
'@@    Description				 :	Function Used to perform search operation using Occurrence Notes Search Criteria
'@@
'@@    Parameters			   :	1.sAction : Action to be performed
'@@   	 							2.bClear : to clear search criteria ( True / False / "" )
'@@   	 							3.bClearOccurrenceNotes : to clear Occurrence Notes table ( True / False / "" )
'@@   	 							4.sOccurrenceNotes : ~ separated list of Occurrence Notes ( String )
'@@   	 							5.sOperators :  ~ separated list of Operators ( String )
'@@   	 							6.sValues :  ~ separated list of Seawrching Values ( String )
'@@   	 						    7.bClickOnSearchButton : to perform search operation ( True / False / "" )
'@@
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	Structure Manager should be displayed						
'@@
'@@    Examples					:	
'@@    								bClear = "true"
'@@    								bClearOccurrenceNotes = "true"
'@@    								sOccurrenceNotes = "AIE_Exported~AIE_Exported"
'@@    								sOperators = "EQ~NE"
'@@    								sValues = "1~2"
'@@    								bClickOnSearchButton = ""
'@@    								sAction = "Search"
'@@    								msgbox  Fn_RDV_OccurrenceNotesSearchPanelOperations(sAction, bClear, bClearOccurrenceNotes, sOccurrenceNotes, sOperators, sValues, bClickOnSearchButton)
'@@
'@@	   History					 	:	
'@@				Developer Name				Date			Rev. No.	Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe			1-Dec-2011			1.0			Created
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public function Fn_RDV_OccurrenceNotesSearchPanelOperations(sAction, bClear, bClearOccurrenceNotes, sOccurrenceNotes, sOperators, sValues, bClickOnSearchButton)
	GBL_FAILED_FUNCTION_NAME="Fn_RDV_OccurrenceNotesSearchPanelOperations"
	Dim objApplet

	Fn_RDV_OccurrenceNotesSearchPanelOperations = False
	Set objApplet = JavaWindow("RDV_StructureManager").JavaWindow("RDV_JApplet")
	If Fn_UI_ObjectExist("Fn_RDV_OccurrenceNotesSearchPanelOperations", objApplet.JavaButton("OccurrenceNotesCashLessSearch")) = False Then
		bReturn = Fn_ToolBarOperation("Click", "Show/Hide Structure Manager Search Panel","")
		If bReturn <> True Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_RDV_OccurrenceNotesSearchPanelOperations ] Failed to click on [ Show/Hide Structure Manager Search Panel ].") 
				Set objApplet = Nothing
				Exit function 
		Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Function [ Fn_RDV_OccurrenceNotesSearchPanelOperations ] Successfully clicked on [ Show/Hide Structure Manager Search Panel ].") 
		End If
	End If
	
	Select Case sAction
		Case "Search"
				If bClear = "" then bClear = "False"
				If cBool(bClear) then
					Call Fn_Button_Click("Fn_RDV_OccurrenceNotesSearchPanelOperations", objApplet,"SearchPanelClear_16Button")
				End If
	
				' clicking on ... button of Form Attribute search criteria
				Call Fn_Button_Click("Fn_RDV_OccurrenceNotesSearchPanelOperations", objApplet, "OccurrenceNotesCashLessSearch")
	
				If Fn_RDV_OccurrenceNotes(bClearOccurrenceNotes, sOccurrenceNotes, sOperators, sValues) = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_RDV_OccurrenceNotesSearchPanelOperations ] Failed to execute function [ Fn_RDV_OccurrenceNotes ].") 
					Set objApplet = Nothing
					Exit function 
				End If
				' clicking on search button
				If bClickOnSearchButton = "" then bClickOnSearchButton = True
				If cBool(bClickOnSearchButton) Then
					Call Fn_Button_Click("Fn_RDV_OccurrenceNotesSearchPanelOperations", objApplet,"SearchPanelSearch_16Button")
				End If
				Fn_RDV_OccurrenceNotesSearchPanelOperations = True
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_RDV_OccurrenceNotesSearchPanelOperations ] Invalid case [ " & sAction& " ].") 
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	End Select

	If Fn_RDV_OccurrenceNotesSearchPanelOperations = True Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Function [ Fn_RDV_OccurrenceNotesSearchPanelOperations ] Executed successfully with case [ " & sAction& " ].") 
	End If
	Set objOccNotes = Nothing
	Set objApplet = Nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_RDV_ClassificationSearchPanelOperations
'@@
'@@    Description				 :	Function Used to perform search operation using Occurrence Notes Search Criteria
'@@
'@@    Parameters			   :	1.sAction : Action to be performed
'@@   	 							2.bClear : to clear search criteria ( True / False / "" )
'@@   	 							3.bClearClassificationPanel : to clear Occurrence Notes table ( True / False / "" )
'@@									4.sSearchClassificationClass : Tree Path of the classification class 
'@@									4.sSysOfMeasurement : System of Measurement ( metric / non-metric )
'@@   	 							4.sPropertyNames : ~ separated list of property names ( String )
'@@   	 							5.sOperators :  ~ separated list of Operators ( String )
'@@   	 							6.sValues :  ~ separated list of Seawrching Values ( String )
'@@   	 						    7.bClickOnSearchButton : to perform search operation ( True / False / "" )
'@@
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	Structure Manager should be displayed						
'@@
'@@    Examples					:	
'@@    								bClear = "true"
'@@    								bClearClassificationPanel = "true"
'@@    								sSearchClassificationClass = "Classification Root:sc1 [1]"
'@@									sSysOfMeasurement = "non-metric"
'@@									sPropertyNames = "sc1.Measure~sc1.Measure"
'@@    								sOperators = "=~>"
'@@    								sValues = "1~2"
'@@    								bClickOnSearchButton = ""
'@@    								sAction = "Search"
'@@    								msgbox  Fn_RDV_ClassificationSearchPanelOperations(sAction, bClear, bClearClassificationPanel, sSearchClassificationClass, sSysOfMeasurement, sPropertyNames, sOperators, sValues, bClickOnSearchButton)
'@@
'@@	   History					 	:	
'@@				Developer Name				Date			Rev. No.	Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe			2-Dec-2011			1.0			Created
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_RDV_ClassificationSearchPanelOperations(sAction, bClear, bClearClassificationPanel, sSearchClassificationClass, sSysOfMeasurement, sPropertyNames, sOperators, sValues, bClickOnSearchButton)
	GBL_FAILED_FUNCTION_NAME="Fn_RDV_ClassificationSearchPanelOperations"
	Dim objApplet

	Fn_RDV_ClassificationSearchPanelOperations = False
	Set objApplet = JavaWindow("RDV_StructureManager").JavaWindow("RDV_JApplet")
	If Fn_UI_ObjectExist("Fn_RDV_ClassificationSearchPanelOperations", objApplet.JavaButton("ClassificationCashLessSearch")) = False Then
		bReturn = Fn_ToolBarOperation("Click", "Show/Hide Structure Manager Search Panel","")
		If bReturn <> True Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_RDV_ClassificationSearchPanelOperations ] Failed to click on [ Show/Hide Structure Manager Search Panel ].") 
				Fn_RDV_ClassificationSearchPanelOperations = False
				Set objApplet = Nothing
				Exit function 
		Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Function [ Fn_RDV_ClassificationSearchPanelOperations ] Successfully clicked on [ Show/Hide Structure Manager Search Panel ].") 
		End If
	End If
	
	Select Case sAction
		Case "Search"
			If bClear = "" then bClear = "False"
			If cBool(bClear) then
				Call Fn_Button_Click("Fn_RDV_ClassificationSearchPanelOperations", objApplet,"SearchPanelClear_16Button")
			End If

			' clicking on ... button of Form Attribute search criteria
			Call Fn_Button_Click("Fn_RDV_ClassificationSearchPanelOperations", objApplet, "ClassificationCashLessSearch")

			If Fn_RDV_Classifications( bClearClassificationPanel, sSearchClassificationClass, sSysOfMeasurement, sPropertyNames, sOperators, sValues) = False Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_RDV_ClassificationSearchPanelOperations ] Failed to execute function [ Fn_RDV_Classifications ].") 
				Fn_RDV_ClassificationSearchPanelOperations = False
				Set objApplet = Nothing
				Exit function 
			End If
			' clicking on search button
			If bClickOnSearchButton = "" then bClickOnSearchButton = True
			If cBool(bClickOnSearchButton) Then
				Call Fn_Button_Click("Fn_RDV_ClassificationSearchPanelOperations", objApplet,"SearchPanelSearch_16Button")
			End If
			Fn_RDV_ClassificationSearchPanelOperations = True
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_RDV_ClassificationSearchPanelOperations ] Invalid case [ " & sAction& " ].") 
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	End Select

	If Fn_RDV_ClassificationSearchPanelOperations = True Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Function [ Fn_RDV_ClassificationSearchPanelOperations ] Executed successfully with case [ " & sAction& " ].") 
	End If
	Set objClassification = Nothing
	Set objApplet = Nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_RDV_MSM_ItemIDSearchPanelOperations
'@@
'@@    Description				 :	Function Used to perform search operation on Item Attribute dialog
'@@
'@@    Parameters			   :	1. dicItemIDSearch: dictionary object
'@@
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	Item Attribute dialog should be displayed							
'@@
'@@    Examples					:	dicItemIDSearch("bChangeSearch") = True
'@@    								dicItemIDSearch("AdvancedDefaultSearchType") = "Item..."
'@@    								dicItemIDSearch("bClearHistory") = "true"
'@@    								dicItemIDSearch("RememberMyLastSearches") = "10"
'@@    								dicItemIDSearch("SearchType") = "Item..."
'@@    								dicItemIDSearch("SearchCriteria") = "Item ID=000038~Name=comp1"
'@@    								dicItemIDSearch("bShowResultInNewTab") = True
'@@    								dicItemIDSearch("bClickOnSearchButton") = True
'@@    								dicItemIDSearch("BOMLine") = "000454/A;1-TopAssy (View)"
'@@    								Call Fn_RDV_MSM_ItemIDSearchPanelOperations("Search", dicItemIDSearch)
'@@
'@@	   History					 	:	
'@@				Developer Name				Date			Rev. No.	Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe			9-Jan-2012			1.0			Created
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Navin Gupta				21-Jan-2013			1.0			Modified code to open structure search in MSM
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_RDV_MSM_ItemIDSearchPanelOperations(sAction, dicItemIDSearch)
	GBL_FAILED_FUNCTION_NAME="Fn_RDV_MSM_ItemIDSearchPanelOperations"
	Dim objApplet, bReturn
				
	Fn_RDV_MSM_ItemIDSearchPanelOperations = False
	Set objApplet = JavaWindow("RDV_StructureManager")
	objApplet.JavaStaticText("MSM_SearchType").SetTOProperty "label", "Item attributes:"
	' opening serach panel if it is not displayed.
	If Fn_UI_ObjectExist("Fn_RDV_MSM_ItemIDSearchPanelOperations", objApplet.JavaButton("ItemAttributesCashLessSearch")) = False Then
'		Code Added by Naveen : Sturcture Search can not be open by menu [ Manufacturing:Structure Search ] in MSM
		bReturn = Fn_MSM_BOMTableNodeOpeations("PopupMenuSelect",dicItemIDSearch("BOMLine"),"","","Structure Search...")
'		bReturn = Fn_SetView("Manufacturing:Structure Search")
		If bReturn <> True Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_RDV_MSM_ItemIDSearchPanelOperations ] Failed to RMB on a BOM line ["+dicItemIDSearch("BOMLine")+ "] and select  [ Structure Search] .") 
				Fn_RDV_MSM_ItemIDSearchPanelOperations = False
				Exit function 
		Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Function [ Fn_RDV_MSM_ItemIDSearchPanelOperations ] Successfully RMB on a BOM line ["+dicItemIDSearch("BOMLine")+ "] and select  [ Structure Search] .") 
		End If
	End If

	Select Case sAction
		Case "Search"
			If dicItemIDSearch("bClear") = "" then dicItemIDSearch("bClear") = "False"
			If cBool(dicItemIDSearch("bClear")) then
				Call Fn_Button_Click("Fn_RDV_MSM_ItemIDSearchPanelOperations", objApplet,"MSM_SearchPanelClear_16Button")
				bReturn = Fn_RDV_MSM_ConfirmationBoxOperations("ClearAll", "", "OK")
			End If
			' clicking on ... button of ItemID search criteria
			Call Fn_Button_Click("Fn_RDV_MSM_ItemIDSearchPanelOperations", objApplet, "ItemAttributesCashLessSearch")
			
			' Item Attribut function call
			IF Fn_RDV_ItemAttributes(dicItemIDSearch) = False then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_RDV_ItemAttributes ] Failed to open [ Item Attributes ] dialog.") 
				Exit function
			End IF
			
			' open in new tab
			'Check existane of the Check box before clicking
			If JavaWindow("RDV_StructureManager").JavaCheckBox("MSM_ShowResultsInANew").Exist(2) Then
					if dicItemIDSearch("bShowResultInNewTab") <> "" then
						If cBool(dicItemIDSearch("bShowResultInNewTab")) Then
							Call Fn_CheckBox_Set("",JavaWindow("RDV_StructureManager"), "MSM_ShowResultsInANew","ON")
						Else
							Call Fn_CheckBox_Set("",JavaWindow("RDV_StructureManager"), "MSM_ShowResultsInANew","OFF")
						End If
					End if
			Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Function [ Fn_RDV_MSM_ItemIDSearchPanelOperations ] Check box  [Show Results In ANew TAB ]  not present") 
			End If
			
			' clicking on search button
			If dicItemIDSearch("bClickOnSearchButton") = "" then dicItemIDSearch("bClickOnSearchButton") = True
			If cBool(dicItemIDSearch("bClickOnSearchButton")) Then
				'	Call Fn_Button_Click("Fn_RDV_MSM_ItemIDSearchPanelOperations", objApplet,"MSM_SearchPanelSearch_16Button")'''               Commented This line added bellow one as Search button is not visible
                objApplet.JavaButton("MSM_SearchPanelSearch_16Button").Object.click                                                                                                     'Added by Avinash J.    19-march-2013
			End If
			Fn_RDV_MSM_ItemIDSearchPanelOperations = True
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_RDV_MSM_ItemIDSearchPanelOperations ] Invalid case [ " & sAction& " ].") 
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	End Select
	If Fn_RDV_MSM_ItemIDSearchPanelOperations = True Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Function [ Fn_RDV_MSM_ItemIDSearchPanelOperations ] Executed successfully with case [ " & sAction& " ].") 
	End If
	Set objApplet = Nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_RDV_MSM_FormAttributesSearchPanelOperations
'@@
'@@    Description				 :	Function Used to perform search operation using Form Attribute Search Criteria
'@@
'@@    Parameters			   :	1.sAction : Action to be performed
'@@   	 							2.bClear : to clear search criteria ( True / False / "" )
'@@   	 							3.bClearFormAttributePanel : to clear Form Attribute table ( True / False / "" )
'@@   	 							4.sRelationType : ~ separated list of Relation Types ( String )
'@@   	 							5.sParentType :  ~ separated list of Parent Types ( String )
'@@   	 							6.sFormType :  ~ separated list of Form Types ( String )
'@@   	 							7.sPropertyName :  ~ separated list of Property Names ( String )
'@@   	 							8.sOperator :  ~ separated list of Operators ( String )
'@@   	 							9.sSearchingValue :  ~ separated list of Seawrching Values ( String )
'@@   	 						   10.bClickOnSearchButton : to perform search operation ( True / False / "" )
'@@
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	Structure Manager should be displayed						
'@@
'@@    Examples					:	
'@@    								bClear = "true"
'@@    								bClearFormAttributePanel = "true"
'@@    								sRelationType = "3D Snapshot~"
'@@    								sParentType = "Item~"
'@@    								sFormType = "Arc Override Form~"
'@@    								sPropertyName = "Start Point RZ~"
'@@    								sOperator = "EQ~"
'@@    								sSearchingValue = "1~2"
'@@    								bClickOnSearchButton = ""
'@@    								sAction = "Search"
'@@    								Call Fn_RDV_MSM_FormAttributesSearchPanelOperations(sAction, bClear, bClearFormAttributePanel, sRelationType, sParentType, sFormType, sPropertyName, sOperator, sSearchingValue, bClickOnSearchButton)
'@@
'@@	   History					 	:	
'@@				Developer Name				Date			Rev. No.	Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe			9-Jan-2012			1.0			Created
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Navin Gupta				21-Jan-2013			1.0			Modified code to open structure search in MSM
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_RDV_MSM_FormAttributesSearchPanelOperations(sAction, bClear, bClearFormAttributePanel, sRelationType, sParentType, sFormType, sPropertyName, sOperator, sSearchingValue, bShowResultInNewTab, bClickOnSearchButton)
	GBL_FAILED_FUNCTION_NAME="Fn_RDV_MSM_FormAttributesSearchPanelOperations"
	Dim objApplet, bReturn
				
	Fn_RDV_MSM_FormAttributesSearchPanelOperations = False
	Set objApplet = JavaWindow("RDV_StructureManager")
	objApplet.JavaStaticText("MSM_SearchType").SetTOProperty "label", "Form attributes:"
	' opening serach panel if it is not displayed.
	If Fn_UI_ObjectExist("Fn_RDV_MSM_ItemIDSearchPanelOperations", objApplet.JavaButton("ItemAttributesCashLessSearch")) = False Then
'		Code Added by Naveen : Sturcture Search can not be open by menu [ Manufacturing:Structure Search ] in MSM
		bReturn = Fn_MSM_BOMTableNodeOpeations("PopupMenuSelect",dicItemIDSearch("BOMLine"),"","","Structure Search...")
'		bReturn = Fn_SetView("Manufacturing:Structure Search")
		If bReturn <> True Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_RDV_MSM_ItemIDSearchPanelOperations ] Failed to RMB on a BOM line ["+dicItemIDSearch("BOMLine")+ "] and select  [ Structure Search] .") 
				Fn_RDV_MSM_ItemIDSearchPanelOperations = False
				Exit function 
		Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Function [ Fn_RDV_MSM_ItemIDSearchPanelOperations ] Successfully RMB on a BOM line ["+dicItemIDSearch("BOMLine")+ "] and select  [ Structure Search] .") 
		End If
	End If

	Select Case sAction
		Case "Search"
			If bClear = "" then bClear = "False"
			If cBool(bClear) then
				Call Fn_Button_Click("Fn_RDV_MSM_FormAttributesSearchPanelOperations", objApplet,"MSM_SearchPanelClear_16Button")
				Call Fn_RDV_MSM_ConfirmationBoxOperations("ClearAll", "", "OK")
			End If
			' clicking on ... button of ItemID search criteria
			Call Fn_Button_Click("Fn_RDV_MSM_FormAttributesSearchPanelOperations", objApplet, "MSM_CashLessSearchDetails")
			
			' Item Attribut function call
			IF Fn_RDV_FormAttributes(bClearFormAttributePanel, sRelationType, sParentType, sFormType, sPropertyName, sOperator, sSearchingValue) = False then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_RDV_FormAttributes ] Failed to open [ Form Attributes ] dialog.") 
				Exit function
			End IF
			
			' open in new tab
			'Check existane of the Check box before clicking - Code added by Archana 25-11-13
			If JavaWindow("RDV_StructureManager").JavaCheckBox("MSM_ShowResultsInANew").Exist(2) Then
				if bShowResultInNewTab <> "" then
					If cBool(bShowResultInNewTab) Then
						Call Fn_CheckBox_Set("Fn_RDV_MSM_FormAttributesSearchPanelOperations",JavaWindow("RDV_StructureManager"), "MSM_ShowResultsInANew","ON")
					Else
						Call Fn_CheckBox_Set("Fn_RDV_MSM_FormAttributesSearchPanelOperations",JavaWindow("RDV_StructureManager"), "MSM_ShowResultsInANew","OFF")
					End If
				End if
           Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Function [ Fn_RDV_MSM_FormAttributesSearchPanelOperations ] Check box  [Show Results In ANew TAB ]  not present") 
			End If
			
			' clicking on search button
			If dicItemIDSearch("bClickOnSearchButton") = "" then dicItemIDSearch("bClickOnSearchButton") = True
			If cBool(dicItemIDSearch("bClickOnSearchButton")) Then
				Call Fn_Button_Click("Fn_RDV_MSM_FormAttributesSearchPanelOperations", objApplet,"MSM_SearchPanelSearch_16Button")
			End If
			Fn_RDV_MSM_FormAttributesSearchPanelOperations = True
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_RDV_MSM_FormAttributesSearchPanelOperations ] Invalid case [ " & sAction& " ].") 
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	End Select
	If Fn_RDV_MSM_FormAttributesSearchPanelOperations = True Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Function [ Fn_RDV_MSM_FormAttributesSearchPanelOperations ] Executed successfully with case [ " & sAction& " ].") 
	End If
	Set objApplet = Nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_RDV_MSM_OccurrenceNotesSearchPanelOperations
'@@
'@@    Description				 :	Function Used to perform search operation using Occurrence Notes Search Criteria
'@@
'@@    Parameters			   :	1.sAction : Action to be performed
'@@   	 							2.bClear : to clear search criteria ( True / False / "" )
'@@   	 							3.bClearOccurrenceNotes : to clear Occurrence Notes table ( True / False / "" )
'@@   	 							4.sOccurrenceNotes : ~ separated list of Occurrence Notes ( String )
'@@   	 							5.sOperators :  ~ separated list of Operators ( String )
'@@   	 							6.sValues :  ~ separated list of Seawrching Values ( String )
'@@   	 						    7.bClickOnSearchButton : to perform search operation ( True / False / "" )
'@@
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	Structure Manager should be displayed						
'@@
'@@    Examples					:	
'@@    								bClear = "true"
'@@    								bClearOccurrenceNotes = "true"
'@@    								sOccurrenceNotes = "AIE_Exported~AIE_Exported"
'@@    								sOperators = "EQ~NE"
'@@    								sValues = "1~2"
'@@    								bClickOnSearchButton = ""
'@@    								sAction = "Search"
'@@    								msgbox  Fn_RDV_MSM_OccurrenceNotesSearchPanelOperations(sAction, bClear, bClearOccurrenceNotes, sOccurrenceNotes, sOperators, sValues, bClickOnSearchButton)
'@@
'@@	   History					 	:	
'@@				Developer Name				Date			Rev. No.	Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe			1-Dec-2011			1.0			Created
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Navin Gupta				21-Jan-2013			1.0			Modified code to open structure search in MSM
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public function Fn_RDV_MSM_OccurrenceNotesSearchPanelOperations(sAction, bClear, bClearOccurrenceNotes, sOccurrenceNotes, sOperators, sValues, bShowResultInNewTab, bClickOnSearchButton)
	GBL_FAILED_FUNCTION_NAME="Fn_RDV_MSM_OccurrenceNotesSearchPanelOperations"
	Dim objApplet, bReturn
				
	Fn_RDV_MSM_OccurrenceNotesSearchPanelOperations = False

	Set objApplet = JavaWindow("RDV_StructureManager")
	objApplet.JavaStaticText("MSM_SearchType").SetTOProperty "label", "Occurrence notes:"
	' opening serach panel if it is not displayed.
	If Fn_UI_ObjectExist("Fn_RDV_MSM_ItemIDSearchPanelOperations", objApplet.JavaButton("ItemAttributesCashLessSearch")) = False Then
'		Code Added by Naveen : Sturcture Search can not be open by menu [ Manufacturing:Structure Search ] in MSM
		bReturn = Fn_MSM_BOMTableNodeOpeations("PopupMenuSelect",dicItemIDSearch("BOMLine"),"","","Structure Search...")
'		bReturn = Fn_SetView("Manufacturing:Structure Search")
		If bReturn <> True Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_RDV_MSM_ItemIDSearchPanelOperations ] Failed to RMB on a BOM line ["+dicItemIDSearch("BOMLine")+ "] and select  [ Structure Search] .") 
				Fn_RDV_MSM_ItemIDSearchPanelOperations = False
				Exit function 
		Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Function [ Fn_RDV_MSM_ItemIDSearchPanelOperations ] Successfully RMB on a BOM line ["+dicItemIDSearch("BOMLine")+ "] and select  [ Structure Search] .") 
		End If
	End If

	Select Case sAction
		Case "Search"
			If bClear = "" then bClear = "False"
			If cBool(bClear) then
				Call Fn_Button_Click("Fn_RDV_MSM_OccurrenceNotesSearchPanelOperations", objApplet,"MSM_SearchPanelClear_16Button")
				bReturn = Fn_RDV_MSM_ConfirmationBoxOperations("ClearAll", "", "OK")
			End If
			' clicking on ... button of ItemID search criteria
			Call Fn_Button_Click("Fn_RDV_MSM_OccurrenceNotesSearchPanelOperations", objApplet, "MSM_CashLessSearchDetails")
			
			' Item Attribut function call
			IF Fn_RDV_OccurrenceNotes(bClearOccurrenceNotes, sOccurrenceNotes, sOperators, sValues) = False then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_RDV_OccurrenceNotes ] Failed to open [ Occurrence notes ] dialog.") 
				Exit function
			End IF
			
			' open in new tab
'            Check existane of the Check box before clicking - Code added by Archana 25-11-13
			If JavaWindow("RDV_StructureManager").JavaCheckBox("MSM_ShowResultsInANew").Exist(2) Then
				if bShowResultInNewTab <> "" then
					If cBool(bShowResultInNewTab) Then
						Call Fn_CheckBox_Set("Fn_RDV_MSM_OccurrenceNotesSearchPanelOperations",JavaWindow("RDV_StructureManager"), "MSM_ShowResultsInANew","ON")
					Else
						Call Fn_CheckBox_Set("Fn_RDV_MSM_OccurrenceNotesSearchPanelOperations",JavaWindow("RDV_StructureManager"), "MSM_ShowResultsInANew","OFF")
					End If
				End if
            Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Function [ Fn_RDV_MSM_OccurrenceNotesSearchPanelOperations ] Check box  [Show Results In ANew TAB ]  not present") 
			End If
			
			' clicking on search button
			If bClickOnSearchButton = "" then bClickOnSearchButton = True
			If cBool(bClickOnSearchButton) Then
				Call Fn_Button_Click("Fn_RDV_MSM_OccurrenceNotesSearchPanelOperations", objApplet,"MSM_SearchPanelSearch_16Button")
			End If
			
			Fn_RDV_MSM_OccurrenceNotesSearchPanelOperations = True
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_RDV_MSM_OccurrenceNotesSearchPanelOperations ] Invalid case [ " & sAction& " ].") 
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	End Select
	If Fn_RDV_MSM_OccurrenceNotesSearchPanelOperations = True Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Function [ Fn_RDV_MSM_OccurrenceNotesSearchPanelOperations ] Executed successfully with case [ " & sAction& " ].") 
	End If
	Set objApplet = Nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_RDV_MSM_ClassificationSearchPanelOperations
'@@
'@@    Description				 :	Function Used to perform search operation using Occurrence Notes Search Criteria
'@@
'@@    Parameters			   :	1.sAction : Action to be performed
'@@   	 							2.bClear : to clear search criteria ( True / False / "" )
'@@   	 							3.bClearClassificationPanel : to clear Occurrence Notes table ( True / False / "" )
'@@									4.sSearchClassificationClass : Tree Path of the classification class 
'@@									4.sSysOfMeasurement : System of Measurement ( metric / non-metric )
'@@   	 							4.sPropertyNames : ~ separated list of property names ( String )
'@@   	 							5.sOperators :  ~ separated list of Operators ( String )
'@@   	 							6.sValues :  ~ separated list of Seawrching Values ( String )
'@@   	 						    7.bClickOnSearchButton : to perform search operation ( True / False / "" )
'@@
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	Structure Manager should be displayed						
'@@
'@@    Examples					:	
'@@    								bClear = "true"
'@@    								bClearClassificationPanel = "true"
'@@    								sSearchClassificationClass = "Classification Root:sc1 [1]"
'@@									sSysOfMeasurement = "non-metric"
'@@									sPropertyNames = "sc1.Measure~sc1.Measure"
'@@    								sOperators = "=~>"
'@@    								sValues = "1~2"
'@@    								bClickOnSearchButton = ""
'@@    								sAction = "Search"
'@@    								msgbox  Fn_RDV_MSM_ClassificationSearchPanelOperations(sAction, bClear, bClearClassificationPanel, sSearchClassificationClass, sSysOfMeasurement, sPropertyNames, sOperators, sValues, bShowResultInNewTab, bClickOnSearchButton)
'@@
'@@	   History					 	:	
'@@				Developer Name				Date			Rev. No.	Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe			2-Dec-2011			1.0			Created
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Navin Gupta				21-Jan-2013			1.0			Modified code to open structure search in MSM
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_RDV_MSM_ClassificationSearchPanelOperations(sAction, bClear, bClearClassificationPanel, sSearchClassificationClass, sSysOfMeasurement, sPropertyNames, sOperators, sValues, bClickOnSearchButton)
	GBL_FAILED_FUNCTION_NAME="Fn_RDV_MSM_ClassificationSearchPanelOperations"
	Dim objApplet, bReturn
				
	Fn_RDV_MSM_ClassificationSearchPanelOperations = False
	Set objApplet = JavaWindow("RDV_StructureManager")
	objApplet.JavaStaticText("MSM_SearchType").SetTOProperty "label", "Classification:"
	' opening serach panel if it is not displayed.
	If Fn_UI_ObjectExist("Fn_RDV_MSM_ItemIDSearchPanelOperations", objApplet.JavaButton("ItemAttributesCashLessSearch")) = False Then
'		Code Added by Naveen : Sturcture Search can not be open by menu [ Manufacturing:Structure Search ] in MSM
		bReturn = Fn_MSM_BOMTableNodeOpeations("PopupMenuSelect",dicItemIDSearch("BOMLine"),"","","Structure Search...")
'		bReturn = Fn_SetView("Manufacturing:Structure Search")
		If bReturn <> True Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_RDV_MSM_ItemIDSearchPanelOperations ] Failed to RMB on a BOM line ["+dicItemIDSearch("BOMLine")+ "] and select  [ Structure Search] .") 
				Fn_RDV_MSM_ItemIDSearchPanelOperations = False
				Exit function 
		Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Function [ Fn_RDV_MSM_ItemIDSearchPanelOperations ] Successfully RMB on a BOM line ["+dicItemIDSearch("BOMLine")+ "] and select  [ Structure Search] .") 
		End If
	End If

	Select Case sAction
		Case "Search"
			If bClear = "" then bClear = "False"
			If cBool(bClear) then
				Call Fn_Button_Click("Fn_RDV_MSM_ClassificationSearchPanelOperations", objApplet,"MSM_SearchPanelClear_16Button")
				bReturn = Fn_RDV_MSM_ConfirmationBoxOperations("ClearAll", "", "OK")
			End If
			' clicking on ... button of ItemID search criteria
			Call Fn_Button_Click("Fn_RDV_MSM_ClassificationSearchPanelOperations", objApplet, "MSM_CashLessSearchDetails")
			
			' Item Attribut function call
			If Fn_RDV_Classifications( bClearClassificationPanel, sSearchClassificationClass, sSysOfMeasurement, sPropertyNames, sOperators, sValues) = False Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_RDV_Classifications ] Failed to open [ Classification ] dialog.") 
				Exit function
			End IF
			
			' open in new tab
'Check existane of the Check box before clicking - Code added by Archana 25-11-13
			If JavaWindow("RDV_StructureManager").JavaCheckBox("MSM_ShowResultsInANew").Exist(2) Then
				if bShowResultInNewTab <> "" then
					If cBool(bShowResultInNewTab) Then
						Call Fn_CheckBox_Set("Fn_RDV_MSM_ClassificationSearchPanelOperations",JavaWindow("RDV_StructureManager"), "MSM_ShowResultsInANew","ON")
					Else
						Call Fn_CheckBox_Set("Fn_RDV_MSM_ClassificationSearchPanelOperations",JavaWindow("RDV_StructureManager"), "MSM_ShowResultsInANew","OFF")
					End If
				End if
			Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Function [ Fn_RDV_MSM_ClassificationSearchPanelOperations ] Check box  [Show Results In ANew TAB ]  not present") 
			End If	
			' clicking on search button
			If bClickOnSearchButton = "" then bClickOnSearchButton = True
			If cBool(bClickOnSearchButton) Then
				Call Fn_Button_Click("Fn_RDV_MSM_ClassificationSearchPanelOperations", objApplet,"MSM_SearchPanelSearch_16Button")
			End If
			Fn_RDV_MSM_ClassificationSearchPanelOperations = True
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_RDV_MSM_ClassificationSearchPanelOperations ] Invalid case [ " & sAction& " ].") 
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	End Select
	If Fn_RDV_MSM_ClassificationSearchPanelOperations = True Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Function [ Fn_RDV_MSM_ClassificationSearchPanelOperations ] Executed successfully with case [ " & sAction& " ].") 
	End If
	Set objApplet = Nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_RDV_MSM_SearchResultsOperations
'@@
'@@    Description				 :	Function Used to perform search operation using Occurrence Notes Search Criteria
'@@
'@@    Parameters			   :	1.sAction : Action to be performed
'@@   	 							2.bClear : to clear search criteria ( True / False / "" )
'@@   	 							3.bClearClassificationPanel : to clear Occurrence Notes table ( True / False / "" )
'@@									4.sSearchClassificationClass : Tree Path of the classification class 
'@@									4.sSysOfMeasurement : System of Measurement ( metric / non-metric )
'@@   	 							4.sPropertyNames : ~ separated list of property names ( String )
'@@   	 							5.sOperators :  ~ separated list of Operators ( String )
'@@   	 							6.sValues :  ~ separated list of Seawrching Values ( String )
'@@   	 						    7.bClickOnSearchButton : to perform search operation ( True / False / "" )
'@@
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	Structure Manager should be displayed						
'@@
'@@    Examples					:	
'@@    								bClear = "true"
'@@    								bClearClassificationPanel = "true"
'@@    								sSearchClassificationClass = "Classification Root:sc1 [1]"
'@@									sSysOfMeasurement = "non-metric"
'@@									sPropertyNames = "sc1.Measure~sc1.Measure"
'@@    								sOperators = "=~>"
'@@    								sValues = "1~2"
'@@    								bClickOnSearchButton = ""
'@@    								sAction = "Search"
'@@    								msgbox  Fn_RDV_MSM_SearchResultsOperations(sAction, bClear, bClearClassificationPanel, sSearchClassificationClass, sSysOfMeasurement, sPropertyNames, sOperators, sValues, bClickOnSearchButton)
'@@
'@@	   History					 	:	
'@@				Developer Name				Date			Rev. No.	Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe			2-Dec-2011			1.0			Created
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_RDV_MSM_SearchResultsOperations(sAction, sTab, sRow, sColumn, sValue, bCloseDialog)
	GBL_FAILED_FUNCTION_NAME="Fn_RDV_MSM_SearchResultsOperations"
   	Dim objSrchRes, iCnt, iRowCount, aValue, sText
	Dim aRows, iArrCnt, iWaitCnt, iInstance, iInstCnt
	Set objSrchRes = JavaWindow("RDV_StructureManager")
	Fn_RDV_MSM_SearchResultsOperations = False
	If bCloseDialog = "" Then bCloseDialog = False
	If Fn_UI_ObjectExist("Fn_RDV_MSM_SearchResultsOperations",objSrchRes.JavaTable("Searchcompleted.Found") ) = False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_RDV_MSM_SearchResultsOperations ] Can not find winodw [ Search Results ].") 
		Set objSrchRes = Nothing
		Exit function
	End If

	' waiting to complete search process
	iWaitCnt = 1
	Do While NOT(Fn_UI_ObjectExist("Fn_RDV_MSM_SearchResultsOperations",objSrchRes.JavaStaticText("SearchCompleted") ))
		wait 1
		If iWaitCnt = 30 then 
			exit Do
		End If
		iWaitCnt = iWaitCnt + 1
	Loop

	' selecting Tab
	If sTab <> "" Then
		If objSrchRes.JavaTab("MSM_ResultsTab").Exist(5) Then
			Call Fn_UI_JavaTab_Select("Fn_RDV_MSM_SearchResultsOperations",objSrchRes,"MSM_ResultsTab", sTab)
		End If
	End If

	Select Case sAction
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "VerifyInColumn"
				iRowCount = cInt(objSrchRes.JavaTable("Searchcompleted.Found").GetROProperty("rows"))
				iInstCnt = 1
				aValue = split(sValue,"@")
				if uBound(aValue) = 1 then
					iInstance = cInt(aValue(1))
					sValue = trim(aValue(0))
				Else
					sValue = trim(aValue(0))
					iInstance = 1
				End if
				For iCnt = 0 to iRowCount -1
					If sColumn <> "BOM Line" Then
						sText = Cstr(trim(objSrchRes.JavaTable("Searchcompleted.Found").GetCellData(iCnt, sColumn))) 
					Else
						sText = Cstr( trim(objSrchRes.JavaTable("Searchcompleted.Found").Object.getItem(iCnt).getdata().getProperty("bl_indented_title")))
					End If
					If sText = sValue then
						iF iInstCnt = iInstance Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Function [ Fn_RDV_MSM_SearchResultsOperations ] Successfully verified [ " & sValue &" ] is present in column [ " & sColumn & " ].") 
							Fn_RDV_MSM_SearchResultsOperations = True
							Exit for
						End If
						iInstCnt = iInstCnt + 1
					end if
				Next
				If Fn_RDV_MSM_SearchResultsOperations = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Function [ Fn_RDV_MSM_SearchResultsOperations ] Successfully verified [ " & sValue &" ] is not present in column [ " & sColumn & " ].") 
				End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Select"
				iRowCount = cInt(objSrchRes.JavaTable("Searchcompleted.Found").GetROProperty("rows"))
				For iCnt = 0 to iRowCount -1
					If trim(objSrchRes.JavaTable("MSM_SearchResultsTable").Object.getItem(iCnt).getData().toString()) = sRow then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Function [ Fn_RDV_MSM_SearchResultsOperations ] Successfully selected [ " & sRow &" ] in column [ " & "Item Name" & " ].") 
						objSrchRes.JavaTable("MSM_SearchResultsTable").selectRow iCnt
						Fn_RDV_MSM_SearchResultsOperations = True
						Exit for
					end if
				Next   
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "MultiSelect"
				aRows = split(sRow,"~")
				iRowCount = cInt(objSrchRes.JavaTable("Searchcompleted.Found").GetROProperty("rows"))
				For iArrCnt = 0 to uBound(aRows)
					For iCnt = 0 to iRowCount -1
						If trim(objSrchRes.JavaTable("Searchcompleted.Found").Object.getItem(iCnt).getData().toString()) = aRows(iArrCnt) then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Function [ Fn_RDV_MSM_SearchResultsOperations ] Successfully selected [ " & sRow &" ] in column [ " & "Item Name" & " ].") 
							objSrchRes.JavaTable("Searchcompleted.Found").ExtendRow iCnt
							Fn_RDV_MSM_SearchResultsOperations = True
							Exit for
						end if
					Next
				Next
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case Else
	                  Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_RDV_MSM_SearchResultsOperations ] Invalid case [ " & sAction& " ].") 
	End Select
	
	If Fn_RDV_MSM_SearchResultsOperations = True Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Function [ Fn_RDV_MSM_SearchResultsOperations ] Executed successfully with case [ " & sAction& " ].") 
	End If
	Set objSrchRes = Nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_RDV_SpatialCriteriaSearchPanelOperations
'@@
'@@    Description				 :	Function Used to perform search operation on Spatial Criteria in structure Manager
'@@
'@@    Parameters			   :	sAction = Action to be performed
'@@    											bClear = Clear Flag
'@@    											s3DBoxCoordinates = 3D Box Coordinates  to be selected from list.
'@@    											sSearchType = Search Type
'@@    											XCoord =  X Coordinates
'@@    											XLen = X Length
'@@    											YCoord =  Y Coordinates 
'@@    											YLen = Y Length
'@@    											ZCoord =  Z Coordinates
'@@    											ZLen = Z Length
'@@    											bTrueShapeFiltering = boolean value to select checkbox True Shape Filtering
'@@    											sDistance = distance
'@@    											bCenterToSelected = True / False
'@@    											bClickOnSearchButton = True / False value to click on Update button.
'@@
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	Structure Manager should be displayed							
'@@
'@@    Examples					:	sAction = Action to be performed
'@@    											bClear = Clear Flag
'@@    											s3DBoxCoordinates = 3D Box Coordinates  to be selected from list.
'@@    											sSearchType = Search Type
'@@    											XCoord =  X Coordinates
'@@    											XLen = X Length
'@@    											YCoord =  Y Coordinates 
'@@    											YLen = Y Length
'@@    											ZCoord =  Z Coordinates
'@@    											ZLen = Z Length
'@@    											bTrueShapeFiltering = boolean value to select checkbox True Shape Filtering
'@@    											sDistance = distance
'@@    											bCenterToSelected = True / False
'@@    											bClickOnSearchButton = True / False value to click on Update button.
'@@
'@@	   History					 	:	
'@@				Developer Name				Date			Rev. No.	Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe			28-Feb-2012			1.0			Created
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public function Fn_RDV_SpatialCriteriaSearchPanelOperations(sAction, bClear, s3DBoxCoordinates, sSearchType, XCoord, XLen, YCoord, YLen, ZCoord, ZLen, bTrueShapeFiltering, sDistance, bCenterToSelected, bClickOnSearchButton)
	GBL_FAILED_FUNCTION_NAME="Fn_RDV_SpatialCriteriaSearchPanelOperations"
	Dim objApplet, objSpatialCriteria
	Fn_RDV_SpatialCriteriaSearchPanelOperations = False
	If Window("RDV_StructureManagerWindow").JavaWindow("RDV_JApplet").Exist(1) Then
		Set objApplet = Window("RDV_StructureManagerWindow").JavaWindow("RDV_JApplet")
	ElseIf JavaWindow("RDV_StructureManager").JavaWindow("RDV_JApplet").Exist(1) then 
		Set objApplet = JavaWindow("RDV_StructureManager").JavaWindow("RDV_JApplet")
	End If
	Select Case sAction
		Case "Search"
				Set objSpatialCriteria = JavaWindow("RDV_StructureManager").JavaWindow("Spatial Filter")
				If Fn_SISW_UI_Object_Operations("Fn_RDV_SpatialCriteriaSearchPanelOperations","Enabled", objSpatialCriteria, SISW_MICRO_TIMEOUT) = False Then
					'opening search panel if it is not displayed.
					If Fn_SISW_UI_Object_Operations("Fn_RDV_SpatialCriteriaSearchPanelOperations","Enabled", objApplet.JavaButton("SpatialFilterCashLessSearch"), SISW_MICRO_TIMEOUT) = False Then
						bReturn = Fn_ToolBarOperation("Click", "Show/Hide Structure Manager Search Panel","")
						If bReturn <> True Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_RDV_SpatialCriteriaSearchPanelOperations ] Failed to click on [ Show/Hide Structure Manager Search Panel ].") 
							Set objApplet = Nothing
							Exit function 
						Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Function [ Fn_RDV_SpatialCriteriaSearchPanelOperations ] Successfully clicked on [ Show/Hide Structure Manager Search Panel ].") 
						End If
					End If
					If bClear = "" then bClear = "False"

					If cBool(bClear) then
						Call Fn_Button_Click("Fn_RDV_SpatialCriteriaSearchPanelOperations", objApplet,"SearchPanelClear_16Button")
					End If

					' clicking on ... button of Form Attribute search criteria
					Call Fn_Button_Click("Fn_RDV_SpatialCriteriaSearchPanelOperations", objApplet, "SpatialFilterCashLessSearch")
				End If
				If Fn_RDV_SpatialCriteria(s3DBoxCoordinates, sSearchType, XCoord, XLen, YCoord, YLen, ZCoord, ZLen, bTrueShapeFiltering, sDistance, bCenterToSelected) = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_RDV_SpatialCriteriaSearchPanelOperations ] Failed to execute function [ Fn_RDV_SpatialCriteria ].") 
					Exit Function
				End If
				' clicking on search button
				If bClickOnSearchButton = "" then bClickOnSearchButton = True
				If cBool(bClickOnSearchButton) Then
					Call Fn_Button_Click("Fn_RDV_SpatialCriteriaSearchPanelOperations", objApplet,"SearchPanelSearch_16Button")
				End If
				Fn_RDV_SpatialCriteriaSearchPanelOperations = True
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_RDV_SpatialCriteriaSearchPanelOperations ] Invalid case [ " & sAction& " ].") 
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	End Select
	If Fn_RDV_SpatialCriteriaSearchPanelOperations = True Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Function [ Fn_RDV_SpatialCriteriaSearchPanelOperations ] Executed successfully with case [ " & sAction& " ].") 
	End If
	Set objApplet = Nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_RDV_MSM_SpatialCriteriaSearchPanelOperations
'@@
'@@    Description				 :	Function Used to perform search operation on Spatial Criteria in Multi-structure Manager
'@@
'@@    Parameters			   :	sAction = Action to be performed
'@@    											bClear = Clear Flag
'@@    											s3DBoxCoordinates = 3D Box Coordinates  to be selected from list.
'@@    											sSearchType = Search Type
'@@    											XCoord =  X Coordinates
'@@    											XLen = X Length
'@@    											YCoord =  Y Coordinates 
'@@    											YLen = Y Length
'@@    											ZCoord =  Z Coordinates
'@@    											ZLen = Z Length
'@@    											bTrueShapeFiltering = boolean value to select checkbox True Shape Filtering
'@@    											sDistance = distance
'@@    											bCenterToSelected = True / False
'@@    											bShowResultInNewTab = True / False
'@@    											bClickOnSearchButton = True / False value to click on Update button.
'@@
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	Multi-Structure Manager should be displayed							
'@@
'@@    Examples					:	sAction = Action to be performed
'@@    											bClear = Clear Flag
'@@    											s3DBoxCoordinates = 3D Box Coordinates  to be selected from list.
'@@    											sSearchType = Search Type
'@@    											XCoord =  X Coordinates
'@@    											XLen = X Length
'@@    											YCoord =  Y Coordinates 
'@@    											YLen = Y Length
'@@    											ZCoord =  Z Coordinates
'@@    											ZLen = Z Length
'@@    											bTrueShapeFiltering = boolean value to select checkbox True Shape Filtering
'@@    											sDistance = distance
'@@    											bCenterToSelected = True / False
'@@    											bShowResultInNewTab = True / False
'@@    											bClickOnSearchButton = True / False value to click on Update button.
'@@
'@@	   History					 	:	
'@@				Developer Name				Date			Rev. No.	Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe			28-Feb-2012			1.0			Created
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_RDV_MSM_SpatialCriteriaSearchPanelOperations(sAction, bClear, sSearchCriteriaFor, s3DBoxCoordinates, sSearchType, XCoord, XLen, YCoord, YLen, ZCoord, ZLen, bTrueShapeFiltering, sDistance, bCenterToSelected, bShowResultInNewTab, bClickOnSearchButton)
	GBL_FAILED_FUNCTION_NAME="Fn_RDV_MSM_SpatialCriteriaSearchPanelOperations"
	Dim objApplet, bReturn
				
	Fn_RDV_MSM_SpatialCriteriaSearchPanelOperations = False
	Set objApplet = JavaWindow("RDV_StructureManager")
	objApplet.JavaStaticText("MSM_SearchType").SetTOProperty "label", "Spatial filter:"
	' opening serach panel if it is not displayed.
	If Fn_MSM_TabSet("Structure Search") = False Then
		If Fn_MSM_TabSet("*Structure Search") = False Then
			bReturn = Fn_SetView("Manufacturing:Structure Search")
			If bReturn <> True Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_RDV_MSM_SpatialCriteriaSearchPanelOperations ] Failed to click on [ Window > Show View >Others > Manufacturing > Structure Search ].") 
					Fn_RDV_MSM_SpatialCriteriaSearchPanelOperations = False
					Exit function 
			Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Function [ Fn_RDV_MSM_SpatialCriteriaSearchPanelOperations ] Successfully clicked on [ Window > Show View >Others > Manufacturing > Structure Search ].") 
			End If
		End If
	End If

	Select Case sAction
		Case "Search"
			If bClear = "" then bClear = "False"
			If cBool(bClear) then
				'Call Fn_Button_Click("Fn_RDV_MSM_SpatialCriteriaSearchPanelOperations", objApplet,"MSM_SearchPanelClear_16Button")
				' Code added by Archana to check state of Clear All Button.
				If JavaWindow("RDV_StructureManager").JavaButton("MSM_SearchPanelClear_16Button").GetROProperty("enabled") Then 

					JavaWindow("RDV_StructureManager").JavaButton("MSM_SearchPanelClear_16Button").Click micLeftBtn
					bReturn = Fn_RDV_MSM_ConfirmationBoxOperations("ClearAll", "", "OK")
				End if
			End If
			' clicking on ... button of Spatial Filter search criteria
			Call Fn_Button_Click("Fn_RDV_MSM_SpatialCriteriaSearchPanelOperations", objApplet, "MSM_CashLessSearchDetails")
			' Item Attribut function call
			If Fn_RDV_SpatialCriteria(s3DBoxCoordinates, sSearchType, XCoord, XLen, YCoord, YLen, ZCoord, ZLen, bTrueShapeFiltering, sDistance, bCenterToSelected) = False Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_RDV_Classifications ] Failed to open [ Fn_RDV_SpatialCriteria ] dialog.") 
				Exit function
			End IF
			
			' open in new tab
'Check existane of the Check box before clicking - Code added by Archana 25-11-13
			If JavaWindow("RDV_StructureManager").JavaCheckBox("MSM_ShowResultsInANew").Exist(2) Then
					If bShowResultInNewTab <> "" then
						If Fn_UI_ObjectExist(sFunctionName,JavaWindow("RDV_StructureManager").JavaCheckBox("MSM_ShowResultsInANew"))=True Then
									If cBool(bShowResultInNewTab) Then
										Call Fn_CheckBox_Set("Fn_RDV_MSM_SpatialCriteriaSearchPanelOperations",JavaWindow("RDV_StructureManager"), "MSM_ShowResultsInANew","ON")
									Else
										Call Fn_CheckBox_Set("Fn_RDV_MSM_SpatialCriteriaSearchPanelOperations",JavaWindow("RDV_StructureManager"), "MSM_ShowResultsInANew","OFF")
									End If
						 End If
					End if
            Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Function [ Fn_RDV_MSM_SpatialCriteriaSearchPanelOperations ] Check box  [Show Results In ANew TAB ]  not present") 
			End If
			
			' clicking on search button
			If bClickOnSearchButton = "" then bClickOnSearchButton = True
			If cBool(bClickOnSearchButton) Then
				Call Fn_Button_Click("Fn_RDV_MSM_SpatialCriteriaSearchPanelOperations", objApplet,"MSM_SearchPanelSearch_16Button")
			End If
			Fn_RDV_MSM_SpatialCriteriaSearchPanelOperations = True
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_RDV_MSM_SpatialCriteriaSearchPanelOperations ] Invalid case [ " & sAction& " ].") 
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	End Select
	If Fn_RDV_MSM_SpatialCriteriaSearchPanelOperations = True Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Function [ Fn_RDV_MSM_SpatialCriteriaSearchPanelOperations ] Executed successfully with case [ " & sAction& " ].") 
	End If
	Set objApplet = Nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_RDV_MSM_SpatialCriteriaSearchPanelOperations
'@@
'@@    Description				 :	Function Used to perform search operation on Spatial Criteria in Multi-structure Manager
'@@
'@@    Parameters			   :	sAction = Action to be performed
'@@    								sMessage = Message to verify  -  Not yet implemented
'@@    								sBtnName = Button Name : OK / Cancel
'@@
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	Multi-Structure Manager should be displayed							
'@@
'@@    Examples					:	Call Fn_RDV_MSM_ConfirmationBoxOperations("ClearAll", "", "Cancel")
'@@    Examples					:	Call Fn_RDV_MSM_ConfirmationBoxOperations("SetNewScope", "", "OK")
'@@    Examples					:	Call Fn_RDV_MSM_ConfirmationBoxOperations("StructureSearch", "", "OK")
'@@
'@@	   History					:	
'@@				Developer Name				Date			Rev. No.	Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe			18-Apr-2012			1.0			Created
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_RDV_MSM_ConfirmationBoxOperations(sAction, sMessage, sBtnName)
	GBL_FAILED_FUNCTION_NAME="Fn_RDV_MSM_ConfirmationBoxOperations"
   Dim objDialog
	Fn_RDV_MSM_ConfirmationBoxOperations = False
	Set objDialog = JavaWindow("RDV_StructureManager").JavaWindow("ConfirmationBox")
	Select Case sAction
		Case "ClearAll", "VerifyMessageClearAll"
				' set title to Clear All 
				objDialog.SetTOProperty "title", "Clear All"
				' check existence of Confirmation box
				If objDialog.Exist(5) = False Then
					' if not exists then click on clear all button
					Call Fn_Button_Click("Fn_RDV_MSM_ConfirmationBoxOperations", JavaWindow("RDV_StructureManager"),"MSM_SearchPanelClear_16Button")
				End If

		Case "StructureSearch", "VerifyMessageStructureSearch","SetNewScope", "VerifyMessageSetNewScope"
                ' set title to Structure search 
				objDialog.SetTOProperty "title", "Structure Search"
				If objDialog.Exist(5) = False Then
					' if not exists then click on clear all button
					Call Fn_Button_Click("Fn_RDV_MSM_ConfirmationBoxOperations", JavaWindow("RDV_StructureManager"),"MSM_SetNewScope")
				End If
	End Select

	' verifying existence of confirmation box
	If objDialog.Exist(10)  Then
		Fn_RDV_MSM_ConfirmationBoxOperations = True
	End IF

	' if message is specified
	If instr(sAction,"VerifyMessage") > 0 Then
		Fn_RDV_MSM_ConfirmationBoxOperations = False
		' check message
		' not yet implemented
	End If
    ' click on sBtnName  Button
	If sBtnName <> "" Then
		
' 		Click on OK button of Clear All dialog is not working.
'		Build 	0604 - Patch	20130730.00     - Koustubh, Sachin
		objDialog.Click 100, 50,  "LEFT"
		wait SISW_MIN_TIMEOUT
		Call Fn_KeyBoardOperation("SendKeys", "{ENTER}") 
'		Call Fn_SISW_UI_JavaButton_Operations("Fn_RDV_MSM_ConfirmationBoxOperations", "Click", objDialog, sBtnName)
	End If
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_RDV_SE_ItemIDSearchPanelOperations
'@@
'@@    Description				 :	Function Used to perform search operation on Item Attribute dialog
'@@
'@@    Parameters			   :	1. dicItemIDSearch: dictionary object
'@@
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	Item Attribute dialog should be displayed							
'@@
'@@    Examples					:	dicItemIDSearch("bChangeSearch") = True
'@@    								dicItemIDSearch("AdvancedDefaultSearchType") = "Item..."
'@@    								dicItemIDSearch("bClearHistory") = "true"
'@@    								dicItemIDSearch("RememberMyLastSearches") = "10"
'@@    								dicItemIDSearch("SearchType") = "Item..."
'@@    								dicItemIDSearch("SearchCriteria") = "Item ID=000038~Name=comp1"
'@@    								dicItemIDSearch("bShowResultInNewTab") = True
'@@    								dicItemIDSearch("bClickOnSearchButton") = True
'@@    								Call Fn_RDV_SE_ItemIDSearchPanelOperations("Search", dicItemIDSearch)
'@@
'@@	   History					 	:	
'@@				Developer Name				Date			Rev. No.	Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe			26-Apr-2012			1.0			Created
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_RDV_SE_ItemIDSearchPanelOperations(sAction, dicItemIDSearch)
	GBL_FAILED_FUNCTION_NAME="Fn_RDV_SE_ItemIDSearchPanelOperations"
	Dim objApplet, bReturn
				
	Fn_RDV_SE_ItemIDSearchPanelOperations = False
	Set objApplet = JavaWindow("RDV_StructureManager")
	objApplet.JavaStaticText("MSM_SearchType").SetTOProperty "label", "Item attributes:"
	' opening serach panel if it is not displayed.
	If Fn_UI_ObjectExist("Fn_RDV_SE_ItemIDSearchPanelOperations", objApplet.JavaButton("MSM_CashLessSearchDetails")) = False Then
		bReturn = Fn_SetView("Systems Engineering:Structure Search...")
		If bReturn <> True Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_RDV_SE_ItemIDSearchPanelOperations ] Failed to click on [ Window > Show View > Others > Systems Engineering > Structure Search... ].") 
			Fn_RDV_SE_ItemIDSearchPanelOperations = False
			Exit function 
		Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Function [ Fn_RDV_SE_ItemIDSearchPanelOperations ] Successfully clicked on [ Window > Show View > Others > Systems Engineering > Structure Search... ].") 
		End If
	End If

	Select Case sAction
		Case "Search"
			If dicItemIDSearch("bClear") = "" then dicItemIDSearch("bClear") = "False"
			If cBool(dicItemIDSearch("bClear")) then
				Call Fn_Button_Click("Fn_RDV_SE_ItemIDSearchPanelOperations", objApplet,"MSM_SearchPanelClear_16Button")
				bReturn = Fn_RDV_MSM_ConfirmationBoxOperations("ClearAll", "", "OK")
			End If
			' clicking on ... button of ItemID search criteria
			Call Fn_Button_Click("Fn_RDV_SE_ItemIDSearchPanelOperations", objApplet, "MSM_CashLessSearchDetails")
			
			' Item Attribut function call
			IF Fn_RDV_ItemAttributes(dicItemIDSearch) = False then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_RDV_ItemAttributes ] Failed to open [ Item Attributes ] dialog.") 
				Exit function
			End IF
			
			' open in new tab
           'Check existane of the Check box before clicking - Code added by Archana 25-11-13
			If JavaWindow("RDV_StructureManager").JavaCheckBox("MSM_ShowResultsInANew").Exist(2) Then
				if dicItemIDSearch("bShowResultInNewTab") <> "" then
					If cBool(dicItemIDSearch("bShowResultInNewTab")) Then
						Call Fn_CheckBox_Set("",JavaWindow("RDV_StructureManager"), "MSM_ShowResultsInANew","ON")
					Else
						Call Fn_CheckBox_Set("",JavaWindow("RDV_StructureManager"), "MSM_ShowResultsInANew","OFF")
					End If
				End if
            Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Function [ Fn_RDV_SE_ItemIDSearchPanelOperations ] Check box  [Show Results In ANew TAB ]  not present") 
			End If
				
			' clicking on search button
			If dicItemIDSearch("bClickOnSearchButton") = "" then dicItemIDSearch("bClickOnSearchButton") = True
			If cBool(dicItemIDSearch("bClickOnSearchButton")) Then
				Call Fn_Button_Click("Fn_RDV_SE_ItemIDSearchPanelOperations", objApplet,"MSM_SearchPanelSearch_16Button")
			End If
			Fn_RDV_SE_ItemIDSearchPanelOperations = True
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_RDV_SE_ItemIDSearchPanelOperations ] Invalid case [ " & sAction& " ].") 
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	End Select
	If Fn_RDV_SE_ItemIDSearchPanelOperations = True Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Function [ Fn_RDV_SE_ItemIDSearchPanelOperations ] Executed successfully with case [ " & sAction& " ].") 
	End If
	Set objApplet = Nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_RDV_SE_FormAttributesSearchPanelOperations
'@@
'@@    Description				 :	Function Used to perform search operation using Form Attribute Search Criteria
'@@
'@@    Parameters			   :	1.sAction : Action to be performed
'@@   	 							2.bClear : to clear search criteria ( True / False / "" )
'@@   	 							3.bClearFormAttributePanel : to clear Form Attribute table ( True / False / "" )
'@@   	 							4.sRelationType : ~ separated list of Relation Types ( String )
'@@   	 							5.sParentType :  ~ separated list of Parent Types ( String )
'@@   	 							6.sFormType :  ~ separated list of Form Types ( String )
'@@   	 							7.sPropertyName :  ~ separated list of Property Names ( String )
'@@   	 							8.sOperator :  ~ separated list of Operators ( String )
'@@   	 							9.sSearchingValue :  ~ separated list of Seawrching Values ( String )
'@@   	 						   10.bClickOnSearchButton : to perform search operation ( True / False / "" )
'@@
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	Structure Manager should be displayed						
'@@
'@@    Examples					:	
'@@    								bClear = "true"
'@@    								bClearFormAttributePanel = "true"
'@@    								sRelationType = "3D Snapshot~"
'@@    								sParentType = "Item~"
'@@    								sFormType = "Arc Override Form~"
'@@    								sPropertyName = "Start Point RZ~"
'@@    								sOperator = "EQ~"
'@@    								sSearchingValue = "1~2"
'@@    								bClickOnSearchButton = ""
'@@    								sAction = "Search"
'@@    								Call Fn_RDV_SE_FormAttributesSearchPanelOperations(sAction, bClear, bClearFormAttributePanel, sRelationType, sParentType, sFormType, sPropertyName, sOperator, sSearchingValue, bClickOnSearchButton)
'@@
'@@	   History					 	:	
'@@				Developer Name				Date			Rev. No.	Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe			26-Apr-2012			1.0			Created
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_RDV_SE_FormAttributesSearchPanelOperations(sAction, bClear, bClearFormAttributePanel, sRelationType, sParentType, sFormType, sPropertyName, sOperator, sSearchingValue, bShowResultInNewTab, bClickOnSearchButton)
	GBL_FAILED_FUNCTION_NAME="Fn_RDV_SE_FormAttributesSearchPanelOperations"
	Dim objApplet, bReturn
				
	Fn_RDV_SE_FormAttributesSearchPanelOperations = False
	Set objApplet = JavaWindow("RDV_StructureManager")
	objApplet.JavaStaticText("MSM_SearchType").SetTOProperty "label", "Form attributes:"

	' opening serach panel if it is not displayed.
	If Fn_UI_ObjectExist("Fn_RDV_SE_FormAttributesSearchPanelOperations", objApplet.JavaButton("MSM_CashLessSearchDetails")) = False Then
		bReturn = Fn_SetView("Systems Engineering:Structure Search...")
		If bReturn <> True Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_RDV_SE_ItemIDSearchPanelOperations ] Failed to click on [ Window > Show View > Others > Systems Engineering > Structure Search... ].") 
			Fn_RDV_SE_ItemIDSearchPanelOperations = False
			Exit function 
		Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Function [ Fn_RDV_SE_ItemIDSearchPanelOperations ] Successfully clicked on [ Window > Show View > Others > Systems Engineering > Structure Search... ].") 
		End If
	End If

	Select Case sAction
		Case "Search"
			If bClear = "" then bClear = "False"
			If cBool(bClear) then
				Call Fn_Button_Click("Fn_RDV_SE_FormAttributesSearchPanelOperations", objApplet,"MSM_SearchPanelClear_16Button")
				bReturn = Fn_RDV_MSM_ConfirmationBoxOperations("ClearAll", "", "OK")
			End If
			' clicking on ... button of ItemID search criteria
			Call Fn_Button_Click("Fn_RDV_SE_FormAttributesSearchPanelOperations", objApplet, "MSM_CashLessSearchDetails")
			
			' Item Attribut function call
			IF Fn_RDV_FormAttributes(bClearFormAttributePanel, sRelationType, sParentType, sFormType, sPropertyName, sOperator, sSearchingValue) = False then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_RDV_FormAttributes ] Failed to open [ Form Attributes ] dialog.") 
				Exit function
			End IF
			
			' open in new tab
            'Check existane of the Check box before clicking - Code added by Archana 25-11-13
			If JavaWindow("RDV_StructureManager").JavaCheckBox("MSM_ShowResultsInANew").Exist(2) Then
					if bShowResultInNewTab <> "" then
						If cBool(bShowResultInNewTab) Then
							Call Fn_CheckBox_Set("Fn_RDV_SE_FormAttributesSearchPanelOperations",JavaWindow("RDV_StructureManager"), "MSM_ShowResultsInANew","ON")
						Else
							Call Fn_CheckBox_Set("Fn_RDV_SE_FormAttributesSearchPanelOperations",JavaWindow("RDV_StructureManager"), "MSM_ShowResultsInANew","OFF")
						End If
					End if
			Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Function [ Fn_RDV_SE_FormAttributesSearchPanelOperations ] Check box  [Show Results In ANew TAB ]  not present") 
			End If
			' clicking on search button
			If bClickOnSearchButton = "" then bClickOnSearchButton = True
			If cBool(bClickOnSearchButton) Then
				Call Fn_Button_Click("Fn_RDV_SE_FormAttributesSearchPanelOperations", objApplet,"MSM_SearchPanelSearch_16Button")
			End If
			Fn_RDV_SE_FormAttributesSearchPanelOperations = True
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_RDV_SE_FormAttributesSearchPanelOperations ] Invalid case [ " & sAction& " ].") 
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	End Select
	If Fn_RDV_SE_FormAttributesSearchPanelOperations = True Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Function [ Fn_RDV_SE_FormAttributesSearchPanelOperations ] Executed successfully with case [ " & sAction& " ].") 
	End If
	Set objApplet = Nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_RDV_SE_OccurrenceNotesSearchPanelOperations
'@@
'@@    Description				 :	Function Used to perform search operation using Occurrence Notes Search Criteria
'@@
'@@    Parameters			   :	1.sAction : Action to be performed
'@@   	 							2.bClear : to clear search criteria ( True / False / "" )
'@@   	 							3.bClearOccurrenceNotes : to clear Occurrence Notes table ( True / False / "" )
'@@   	 							4.sOccurrenceNotes : ~ separated list of Occurrence Notes ( String )
'@@   	 							5.sOperators :  ~ separated list of Operators ( String )
'@@   	 							6.sValues :  ~ separated list of Seawrching Values ( String )
'@@   	 						    7.bClickOnSearchButton : to perform search operation ( True / False / "" )
'@@
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	Structure Manager should be displayed						
'@@
'@@    Examples					:	
'@@    								bClear = "true"
'@@    								bClearOccurrenceNotes = "true"
'@@    								sOccurrenceNotes = "AIE_Exported~AIE_Exported"
'@@    								sOperators = "EQ~NE"
'@@    								sValues = "1~2"
'@@    								bClickOnSearchButton = ""
'@@    								sAction = "Search"
'@@    								msgbox  Fn_RDV_SE_OccurrenceNotesSearchPanelOperations(sAction, bClear, bClearOccurrenceNotes, sOccurrenceNotes, sOperators, sValues, bClickOnSearchButton)
'@@
'@@	   History					 	:	
'@@				Developer Name				Date			Rev. No.	Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe			26-Apr-2012			1.0			Created
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public function Fn_RDV_SE_OccurrenceNotesSearchPanelOperations(sAction, bClear, bClearOccurrenceNotes, sOccurrenceNotes, sOperators, sValues, bShowResultInNewTab, bClickOnSearchButton)
	GBL_FAILED_FUNCTION_NAME="Fn_RDV_SE_OccurrenceNotesSearchPanelOperations"
	Dim objApplet, bReturn
				
	Fn_RDV_SE_OccurrenceNotesSearchPanelOperations = False
	Set objApplet = JavaWindow("RDV_StructureManager")
	objApplet.JavaStaticText("MSM_SearchType").SetTOProperty "label", "Occurrence notes:"

	' opening serach panel if it is not displayed.
	If Fn_UI_ObjectExist("Fn_RDV_SE_OccurrenceNotesSearchPanelOperations", objApplet.JavaButton("MSM_CashLessSearchDetails")) = False Then
		bReturn = Fn_SetView("Systems Engineering:Structure Search...")
		If bReturn <> True Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_RDV_SE_ItemIDSearchPanelOperations ] Failed to click on [ Window > Show View > Others > Systems Engineering > Structure Search... ].") 
			Fn_RDV_SE_ItemIDSearchPanelOperations = False
			Exit function 
		Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Function [ Fn_RDV_SE_ItemIDSearchPanelOperations ] Successfully clicked on [ Window > Show View > Others > Systems Engineering > Structure Search... ].") 
		End If
	End If

	Select Case sAction
		Case "Search"
			If bClear = "" then bClear = "False"
			If cBool(bClear) then
				Call Fn_Button_Click("Fn_RDV_SE_OccurrenceNotesSearchPanelOperations", objApplet,"MSM_SearchPanelClear_16Button")
				bReturn = Fn_RDV_MSM_ConfirmationBoxOperations("ClearAll", "", "OK")
			End If
			' clicking on ... button of ItemID search criteria
			Call Fn_Button_Click("Fn_RDV_SE_OccurrenceNotesSearchPanelOperations", objApplet, "MSM_CashLessSearchDetails")
			
			' Item Attribut function call
			IF Fn_RDV_OccurrenceNotes(bClearOccurrenceNotes, sOccurrenceNotes, sOperators, sValues) = False then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_RDV_OccurrenceNotes ] Failed to open [ Occurrence notes ] dialog.") 
				Exit function
			End IF
			
			' open in new tab
            'Check existane of the Check box before clicking - Code added by Archana 25-11-13
			If JavaWindow("RDV_StructureManager").JavaCheckBox("MSM_ShowResultsInANew").Exist(2) Then
					if bShowResultInNewTab <> "" then
						If cBool(bShowResultInNewTab) Then
							Call Fn_CheckBox_Set("Fn_RDV_SE_OccurrenceNotesSearchPanelOperations",JavaWindow("RDV_StructureManager"), "MSM_ShowResultsInANew","ON")
						Else
							Call Fn_CheckBox_Set("Fn_RDV_SE_OccurrenceNotesSearchPanelOperations",JavaWindow("RDV_StructureManager"), "MSM_ShowResultsInANew","OFF")
						End If
					End if
			Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Function [ Fn_RDV_SE_OccurrenceNotesSearchPanelOperations ] Check box  [Show Results In ANew TAB ]  not present") 
			End If
			
			' clicking on search button
			If bClickOnSearchButton = "" then bClickOnSearchButton = True
			If cBool(bClickOnSearchButton) Then
				Call Fn_Button_Click("Fn_RDV_SE_OccurrenceNotesSearchPanelOperations", objApplet,"MSM_SearchPanelSearch_16Button")
			End If
			
			Fn_RDV_SE_OccurrenceNotesSearchPanelOperations = True
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_RDV_SE_OccurrenceNotesSearchPanelOperations ] Invalid case [ " & sAction& " ].") 
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	End Select
	If Fn_RDV_SE_OccurrenceNotesSearchPanelOperations = True Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Function [ Fn_RDV_SE_OccurrenceNotesSearchPanelOperations ] Executed successfully with case [ " & sAction& " ].") 
	End If
	Set objApplet = Nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_RDV_SE_ClassificationSearchPanelOperations
'@@
'@@    Description				 :	Function Used to perform search operation using Occurrence Notes Search Criteria
'@@
'@@    Parameters			   :	1.sAction : Action to be performed
'@@   	 							2.bClear : to clear search criteria ( True / False / "" )
'@@   	 							3.bClearClassificationPanel : to clear Occurrence Notes table ( True / False / "" )
'@@									4.sSearchClassificationClass : Tree Path of the classification class 
'@@									4.sSysOfMeasurement : System of Measurement ( metric / non-metric )
'@@   	 							4.sPropertyNames : ~ separated list of property names ( String )
'@@   	 							5.sOperators :  ~ separated list of Operators ( String )
'@@   	 							6.sValues :  ~ separated list of Seawrching Values ( String )
'@@   	 						    7.bClickOnSearchButton : to perform search operation ( True / False / "" )
'@@
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	Structure Manager should be displayed						
'@@
'@@    Examples					:	
'@@    								bClear = "true"
'@@    								bClearClassificationPanel = "true"
'@@    								sSearchClassificationClass = "Classification Root:sc1 [1]"
'@@									sSysOfMeasurement = "non-metric"
'@@									sPropertyNames = "sc1.Measure~sc1.Measure"
'@@    								sOperators = "=~>"
'@@    								sValues = "1~2"
'@@    								bClickOnSearchButton = ""
'@@    								sAction = "Search"
'@@    								msgbox  Fn_RDV_SE_ClassificationSearchPanelOperations(sAction, bClear, bClearClassificationPanel, sSearchClassificationClass, sSysOfMeasurement, sPropertyNames, sOperators, sValues, bShowResultInNewTab, bClickOnSearchButton)
'@@
'@@	   History					 	:	
'@@				Developer Name				Date			Rev. No.	Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe			26-Apr-2012			1.0			Created
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_RDV_SE_ClassificationSearchPanelOperations(sAction, bClear, bClearClassificationPanel, sSearchClassificationClass, sSysOfMeasurement, sPropertyNames, sOperators, sValues, bClickOnSearchButton)
	GBL_FAILED_FUNCTION_NAME="Fn_RDV_SE_ClassificationSearchPanelOperations"
	Dim objApplet, bReturn
				
	Fn_RDV_SE_ClassificationSearchPanelOperations = False
	Set objApplet = JavaWindow("RDV_StructureManager")
	objApplet.JavaStaticText("MSM_SearchType").SetTOProperty "label", "Classification:"

	' opening serach panel if it is not displayed.
	If Fn_UI_ObjectExist("Fn_RDV_SE_ClassificationSearchPanelOperations", objApplet.JavaButton("MSM_CashLessSearchDetails")) = False Then
		bReturn = Fn_SetView("Systems Engineering:Structure Search...")
		If bReturn <> True Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_RDV_SE_ItemIDSearchPanelOperations ] Failed to click on [ Window > Show View > Others > Systems Engineering > Structure Search... ].") 
			Fn_RDV_SE_ItemIDSearchPanelOperations = False
			Exit function 
		Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Function [ Fn_RDV_SE_ItemIDSearchPanelOperations ] Successfully clicked on [ Window > Show View > Others > Systems Engineering > Structure Search... ].") 
		End If
	End If

	Select Case sAction
		Case "Search"
			If bClear = "" then bClear = "False"
			If cBool(bClear) then
				Call Fn_Button_Click("Fn_RDV_SE_ClassificationSearchPanelOperations", objApplet,"MSM_SearchPanelClear_16Button")
				bReturn = Fn_RDV_MSM_ConfirmationBoxOperations("ClearAll", "", "OK")
			End If
			' clicking on ... button of ItemID search criteria
			Call Fn_Button_Click("Fn_RDV_SE_ClassificationSearchPanelOperations", objApplet, "MSM_CashLessSearchDetails")
			
			' Item Attribut function call
			If Fn_RDV_Classifications( bClearClassificationPanel, sSearchClassificationClass, sSysOfMeasurement, sPropertyNames, sOperators, sValues) = False Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_RDV_Classifications ] Failed to open [ Classification ] dialog.") 
				Exit function
			End IF
			
			' open in new tab
			'Check existane of the Check box before clicking - Code added by Archana 25-11-13
			If JavaWindow("RDV_StructureManager").JavaCheckBox("MSM_ShowResultsInANew").Exist(2) Then
				if bShowResultInNewTab <> "" then
					If cBool(bShowResultInNewTab) Then
						Call Fn_CheckBox_Set("Fn_RDV_SE_ClassificationSearchPanelOperations",JavaWindow("RDV_StructureManager"), "MSM_ShowResultsInANew","ON")
					Else
						Call Fn_CheckBox_Set("Fn_RDV_SE_ClassificationSearchPanelOperations",JavaWindow("RDV_StructureManager"), "MSM_ShowResultsInANew","OFF")
					End If
				End if
			Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Function [ Fn_RDV_SE_ClassificationSearchPanelOperations ] Check box  [Show Results In ANew TAB ]  not present") 
			End If

			' clicking on search button
			If bClickOnSearchButton = "" then bClickOnSearchButton = True
			If cBool(bClickOnSearchButton) Then
				Call Fn_Button_Click("Fn_RDV_SE_ClassificationSearchPanelOperations", objApplet,"MSM_SearchPanelSearch_16Button")
			End If
			Fn_RDV_SE_ClassificationSearchPanelOperations = True
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_RDV_SE_ClassificationSearchPanelOperations ] Invalid case [ " & sAction& " ].") 
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	End Select
	If Fn_RDV_SE_ClassificationSearchPanelOperations = True Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Function [ Fn_RDV_SE_ClassificationSearchPanelOperations ] Executed successfully with case [ " & sAction& " ].") 
	End If
	Set objApplet = Nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_RDV_SE_SearchResultsOperations
'@@
'@@    Description				 :	Function Used to perform search operation using Occurrence Notes Search Criteria
'@@
'@@    Parameters			   :	1.sAction : Action to be performed
'@@   	 							2.bClear : to clear search criteria ( True / False / "" )
'@@   	 							3.bClearClassificationPanel : to clear Occurrence Notes table ( True / False / "" )
'@@									4.sSearchClassificationClass : Tree Path of the classification class 
'@@									4.sSysOfMeasurement : System of Measurement ( metric / non-metric )
'@@   	 							4.sPropertyNames : ~ separated list of property names ( String )
'@@   	 							5.sOperators :  ~ separated list of Operators ( String )
'@@   	 							6.sValues :  ~ separated list of Seawrching Values ( String )
'@@   	 						    7.bClickOnSearchButton : to perform search operation ( True / False / "" )
'@@
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	Structure Manager should be displayed						
'@@
'@@    Examples					:	
'@@    								bClear = "true"
'@@    								bClearClassificationPanel = "true"
'@@    								sSearchClassificationClass = "Classification Root:sc1 [1]"
'@@									sSysOfMeasurement = "non-metric"
'@@									sPropertyNames = "sc1.Measure~sc1.Measure"
'@@    								sOperators = "=~>"
'@@    								sValues = "1~2"
'@@    								bClickOnSearchButton = ""
'@@    								sAction = "Search"
'@@    								msgbox  Fn_RDV_SE_SearchResultsOperations(sAction, bClear, bClearClassificationPanel, sSearchClassificationClass, sSysOfMeasurement, sPropertyNames, sOperators, sValues, bClickOnSearchButton)
'@@
'@@	   History					 	:	
'@@				Developer Name				Date			Rev. No.	Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe			26-Apr-2012			1.0			Created
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_RDV_SE_SearchResultsOperations(sAction, sTab, sRow, sColumn, sValue, bCloseDialog)
	GBL_FAILED_FUNCTION_NAME="Fn_RDV_SE_SearchResultsOperations"
   	Dim objSrchRes, iCnt, iRowCount, aValue, sText
	Dim aRows, iArrCnt, iWaitCnt, iInstance, iInstCnt
	Set objSrchRes = JavaWindow("RDV_StructureManager")
	Fn_RDV_SE_SearchResultsOperations = False
	If bCloseDialog = "" Then bCloseDialog = False
	
	If Fn_UI_ObjectExist("Fn_RDV_SE_SearchResultsOperations",objSrchRes.JavaTable("MSM_SearchResultsTable") ) = False Then
'   If Fn_UI_ObjectExist("Fn_RDV_MSM_SearchResultsOperations",objSrchRes.JavaTable("Searchcompleted.Found") ) = False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_RDV_SE_SearchResultsOperations ] Can not find winodw [ Search Results ].") 
		Set objSrchRes = Nothing
		Exit function
	End If

	' waiting to complete search process
	iWaitCnt = 1
	Do While NOT(Fn_UI_ObjectExist("Fn_RDV_SE_SearchResultsOperations",objSrchRes.JavaStaticText("SearchCompleted") ))
		wait 1
		If iWaitCnt = 30 then 
			exit Do
		End If
		iWaitCnt = iWaitCnt + 1
	Loop

	' selecting Tab
	If sTab <> "" Then
		Call Fn_UI_JavaTab_Select("Fn_RDV_SE_SearchResultsOperations",objSrchRes,"MSM_ResultsTab", sTab)
	End If

	Select Case sAction
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "VerifyInColumn"
'				iRowCount = cInt(objSrchRes.JavaTable("MSM_SearchResultsTable").GetROProperty("rows"))
               iRowCount = cInt(objSrchRes.JavaTable("Searchcompleted.Found").GetROProperty("rows"))
				iInstCnt = 1
				aValue = split(sValue,"@")
				if uBound(aValue) = 1 then
					iInstance = cInt(aValue(1))
					sValue = trim(aValue(0))
				Else
					sValue = trim(aValue(0))
					iInstance = 1
				End if
				For iCnt = 0 to iRowCount -1
					If sColumn <> "BOM Line" Then
                         sText = Cstr(trim(objSrchRes.JavaTable("Searchcompleted.Found").GetCellData(iCnt, sColumn))) 
					Else
                          sText = Cstr( trim(JavaWindow("RDV_StructureManager").JavaTable("Searchcompleted.Found").Object.getItem(0).getData().toString()))
					End If
					If sText = sValue then
						iF iInstCnt = iInstance Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Function [ Fn_RDV_SE_SearchResultsOperations ] Successfully verified [ " & sValue &" ] is present in column [ " & sColumn & " ].") 
							Fn_RDV_SE_SearchResultsOperations = True
							Exit for
						End If
						iInstCnt = iInstCnt + 1
					end if
				Next
				If Fn_RDV_SE_SearchResultsOperations = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Function [ Fn_RDV_SE_SearchResultsOperations ] Successfully verified [ " & sValue &" ] is not present in column [ " & sColumn & " ].") 
				End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Select"
				iRowCount = cInt(objSrchRes.JavaTable("MSM_SearchResultsTable").GetROProperty("rows"))
				For iCnt = 0 to iRowCount -1
					If trim(objSrchRes.JavaTable("MSM_SearchResultsTable").Object.getItem(iCnt).getData().toString()) = sRow then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Function [ Fn_RDV_SE_SearchResultsOperations ] Successfully selected [ " & sRow &" ] in column [ " & "Item Name" & " ].") 
						objSrchRes.JavaTable("MSM_SearchResultsTable").selectRow iCnt
'						JavaWindow("RDV_StructureManager").JavaTable("MSM_SearchResultsTable").ActivateCell iCnt, sColumn
						Fn_RDV_SE_SearchResultsOperations = True
						Exit for
					end if
				Next   
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "MultiSelect"
				aRows = split(sRow,"~")
				iRowCount = cInt(objSrchRes.JavaTable("MSM_SearchResultsTable").GetROProperty("rows"))
				For iArrCnt = 0 to uBound(aRows)
					For iCnt = 0 to iRowCount -1
						If trim(objSrchRes.JavaTable("MSM_SearchResultsTable").Object.getItem(iCnt).getData().toString()) = aRows(iArrCnt) then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Function [ Fn_RDV_SE_SearchResultsOperations ] Successfully selected [ " & sRow &" ] in column [ " & "Item Name" & " ].") 
							objSrchRes.JavaTable("MSM_SearchResultsTable").ExtendRow iCnt
							Fn_RDV_SE_SearchResultsOperations = True
							Exit for
						end if
					Next
				Next
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "SelectCell"
				iRowCount = cInt(objSrchRes.JavaTable("MSM_SearchResultsTable").GetROProperty("rows"))
				For iCnt = 0 to iRowCount -1
					If trim(objSrchRes.JavaTable("MSM_SearchResultsTable").Object.getItem(iCnt).getData().toString()) = sRow then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Function [ Fn_RDV_SE_SearchResultsOperations ] Successfully selected [ " & sRow &" ] in column [ " & "Item Name" & " ].") 
						JavaWindow("RDV_StructureManager").JavaTable("MSM_SearchResultsTable").ActivateCell iCnt, sColumn
						Fn_RDV_SE_SearchResultsOperations = True
						Exit for
					end if
				Next   
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case Else
	                  Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_RDV_SE_SearchResultsOperations ] Invalid case [ " & sAction& " ].") 
	End Select
	
	If Fn_RDV_SE_SearchResultsOperations = True Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Function [ Fn_RDV_SE_SearchResultsOperations ] Executed successfully with case [ " & sAction& " ].") 
	End If
	Set objSrchRes = Nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_RDV_SE_SpatialCriteriaSearchPanelOperations
'@@
'@@    Description				 :	Function Used to perform search operation on Spatial Criteria in Multi-structure Manager
'@@
'@@    Parameters			   :	sAction = Action to be performed
'@@    											bClear = Clear Flag
'@@    											s3DBoxCoordinates = 3D Box Coordinates  to be selected from list.
'@@    											sSearchType = Search Type
'@@    											XCoord =  X Coordinates
'@@    											XLen = X Length
'@@    											YCoord =  Y Coordinates 
'@@    											YLen = Y Length
'@@    											ZCoord =  Z Coordinates
'@@    											ZLen = Z Length
'@@    											bTrueShapeFiltering = boolean value to select checkbox True Shape Filtering
'@@    											sDistance = distance
'@@    											bCenterToSelected = True / False
'@@    											bShowResultInNewTab = True / False
'@@    											bClickOnSearchButton = True / False value to click on Update button.
'@@
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	Multi-Structure Manager should be displayed							
'@@
'@@    Examples					:	sAction = Action to be performed
'@@    											bClear = Clear Flag
'@@    											s3DBoxCoordinates = 3D Box Coordinates  to be selected from list.
'@@    											sSearchType = Search Type
'@@    											XCoord =  X Coordinates
'@@    											XLen = X Length
'@@    											YCoord =  Y Coordinates 
'@@    											YLen = Y Length
'@@    											ZCoord =  Z Coordinates
'@@    											ZLen = Z Length
'@@    											bTrueShapeFiltering = boolean value to select checkbox True Shape Filtering
'@@    											sDistance = distance
'@@    											bCenterToSelected = True / False
'@@    											bShowResultInNewTab = True / False
'@@    											bClickOnSearchButton = True / False value to click on Update button.
'@@
'@@	   History					 	:	
'@@				Developer Name				Date			Rev. No.	Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe			26-Apr-2012			1.0			Created
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_RDV_SE_SpatialCriteriaSearchPanelOperations(sAction, bClear, sSearchCriteriaFor, s3DBoxCoordinates, sSearchType, XCoord, XLen, YCoord, YLen, ZCoord, ZLen, bTrueShapeFiltering, sDistance, bCenterToSelected, bShowResultInNewTab, bClickOnSearchButton)
	GBL_FAILED_FUNCTION_NAME="Fn_RDV_SE_SpatialCriteriaSearchPanelOperations"
	Dim objApplet, bReturn
				
	Fn_RDV_SE_SpatialCriteriaSearchPanelOperations = False
	Set objApplet = JavaWindow("RDV_StructureManager")
	objApplet.JavaStaticText("MSM_SearchType").SetTOProperty "label", "Spatial filter:"

	' opening serach panel if it is not displayed.
	If Fn_UI_ObjectExist("Fn_RDV_SE_SpatialCriteriaSearchPanelOperations", objApplet.JavaButton("MSM_CashLessSearchDetails")) = False Then
		bReturn = Fn_SetView("Manufacturing:Structure Search")
		If bReturn <> True Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_RDV_SE_SpatialCriteriaSearchPanelOperations ] Failed to click on [ Window > Show View >Others > Manufacturing > Structure Search ].") 
				Fn_RDV_SE_SpatialCriteriaSearchPanelOperations = False
				Exit function 
		Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Function [ Fn_RDV_SE_SpatialCriteriaSearchPanelOperations ] Successfully clicked on [ Window > Show View >Others > Manufacturing > Structure Search ].") 
		End If
	End If

	Select Case sAction
		Case "Search"
			If bClear = "" then bClear = "False"
			If cBool(bClear) then
				Call Fn_Button_Click("Fn_RDV_SE_SpatialCriteriaSearchPanelOperations", objApplet,"MSM_SearchPanelClear_16Button")
				bReturn = Fn_RDV_MSM_ConfirmationBoxOperations("ClearAll", "", "OK")
			End If
			' clicking on ... button of Spatial Filter search criteria
			Call Fn_Button_Click("Fn_RDV_SE_SpatialCriteriaSearchPanelOperations", objApplet, "MSM_CashLessSearchDetails")
			
			' Item Attribut function call
			If Fn_RDV_SpatialCriteria(s3DBoxCoordinates, sSearchType, XCoord, XLen, YCoord, YLen, ZCoord, ZLen, bTrueShapeFiltering, sDistance, bCenterToSelected) = False Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_RDV_Classifications ] Failed to open [ Fn_RDV_SpatialCriteria ] dialog.") 
				Exit function
			End IF
			
			' open in new tab
            'Check existane of the Check box before clicking - Code added by Archana 25-11-13
			If JavaWindow("RDV_StructureManager").JavaCheckBox("MSM_ShowResultsInANew").Exist(2) Then
				if bShowResultInNewTab <> "" then
					If cBool(bShowResultInNewTab) Then
						Call Fn_CheckBox_Set("Fn_RDV_SE_SpatialCriteriaSearchPanelOperations",JavaWindow("RDV_StructureManager"), "MSM_ShowResultsInANew","ON")
					Else
						Call Fn_CheckBox_Set("Fn_RDV_SE_SpatialCriteriaSearchPanelOperations",JavaWindow("RDV_StructureManager"), "MSM_ShowResultsInANew","OFF")
					End If
				End if
			Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Function [ Fn_RDV_SE_SpatialCriteriaSearchPanelOperations ] Check box  [Show Results In ANew TAB ]  not present") 
			End If
			' clicking on search button
			If bClickOnSearchButton = "" then bClickOnSearchButton = True
			If cBool(bClickOnSearchButton) Then
				Call Fn_Button_Click("Fn_RDV_SE_SpatialCriteriaSearchPanelOperations", objApplet,"MSM_SearchPanelSearch_16Button")
			End If
			Fn_RDV_SE_SpatialCriteriaSearchPanelOperations = True
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_RDV_SE_SpatialCriteriaSearchPanelOperations ] Invalid case [ " & sAction& " ].") 
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	End Select
	If Fn_RDV_SE_SpatialCriteriaSearchPanelOperations = True Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Function [ Fn_RDV_SE_SpatialCriteriaSearchPanelOperations ] Executed successfully with case [ " & sAction& " ].") 
	End If
	Set objApplet = Nothing
End Function

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_RDV_SpatialFilterUseSelectionTableOperation
'@@
'@@    Description				 :	Function Used to perform Operations on User Selection Table.
'@@
'@@    Parameters			   :	1.sAction : Action to be performed
'@@   	 							2.sColumn : Column Name.
'@@   	 							3.sValues : values in Column e.g. Item Name
'@@									4.bCloseDialog : Dialog to close or Not
'@@
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	Spatial Filter window should open.				
'@@
'@@    Examples					:	
'@@									msgbox Fn_RDV_SpatialFilterUseSelectionTableOperation("Verify","Item Name", "1602-049~8002-040",False)
'@@									msgbox Fn_RDV_SpatialFilterUseSelectionTableOperation("MultiSelect","Item Name", "1602-049~8002-040",False)
'@@									msgbox Fn_RDV_SpatialFilterUseSelectionTableOperation("Select","Item Name", "8002-040",False)
'@@									msgbox Fn_RDV_SpatialFilterUseSelectionTableOperation("removeall","", "",False)
'@@									msgbox Fn_RDV_SpatialFilterUseSelectionTableOperation("selectall","", "",False)
'@@
'@@	   History					 	:	
'@@				Developer Name				Date			Rev. No.	Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@
'@@				Sachin Joshi			10-Sept-2012			1.0			Created
'@@				Amit T			        11-Sept-2012			     1.0		   Modified Case "verify"
'@@				Amit T			        12-Sept-2012			    1.0		     Added Case "removeall" , "selectall"
'@@				Amit T			        12-Sept-2012			    1.0		     Added Case "verifywithinstance" , "removewithinstance" , "remove"
'@@				
'@@				Vivek Ahirrao			18-Nov-2015				1.1			Modified cases "removewithinstance", "remove", "select"		[TC1121-2015102600-18_11_2015-VivekA-Maintenance]
'@@																			- BOM Line and Parent columns are bydefault there in tab
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

Public Function Fn_RDV_SpatialFilterUseSelectionTableOperation(sAction,sColumn, sValues,bCloseDialog)
	GBL_FAILED_FUNCTION_NAME="Fn_RDV_SpatialFilterUseSelectionTableOperation"
	Dim objSpatialCriteria,sArrValues,sActValue, iCnt, iRowCount , iVals, iCounter
	Dim ArrInsNode , ArrNode , ArrIns
	Dim iTotalCols, sColumnName, sColName

	Fn_RDV_SpatialFilterUseSelectionTableOperation = False
	sArrValues = Split(sValues,"~")
	Set objSpatialCriteria = JavaWindow("RDV_StructureManager").JavaWindow("Spatial Filter")

	Select Case Lcase(sAction)
		Case "verify"
            If Cstr(objSpatialCriteria.JavaTable("UseSelectionTable").GetROProperty("rows")) <> "0" Then
				For iVals = 0 to Ubound(sArrValues)
					For iCnt = 0 to cInt(objSpatialCriteria.JavaTable("UseSelectionTable").GetROProperty("rows")) -1
						sActValue = JavaWindow("RDV_StructureManager").JavaWindow("Spatial Filter").JavaTable("UseSelectionTable").Object.getItem(iCnt).getData().toString()
						If Instr(sActValue,sArrValues(iVals)) > 0 Then
							Fn_RDV_SpatialFilterUseSelectionTableOperation = True
							Exit For
						End If
					Next					
					If iCnt > cInt(objSpatialCriteria.JavaTable("UseSelectionTable").GetROProperty("rows")) -1 Then
						Fn_RDV_SpatialFilterUseSelectionTableOperation = False
						Exit Function
					End If
				Next
			Else
				Fn_RDV_SpatialFilterUseSelectionTableOperation = False
                Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_RDV_SpatialFilterUseSelectionTableOperation ] Failed to No Items in Dialog.")
				Exit Function
			End If
			
		Case "verifywithinstance"
            If Cstr(objSpatialCriteria.JavaTable("UseSelectionTable").GetROProperty("rows")) <> "0" Then
				For iVals = 0 to Ubound(sArrValues)
				
					ArrInsNode = Split( sArrValues(iVals) , "@" )
					GlobIns = 1
					If Ubound(ArrInsNode) = 1 Then
						ArrNode = ArrInsNode(0)
						ArrIns = ArrInsNode(1)
					Else
						ArrNode = ArrInsNode(0)
						ArrIns = GlobIns
					End If	
					
					For iCnt = 0 to cInt(objSpatialCriteria.JavaTable("UseSelectionTable").GetROProperty("rows")) -1
						sActValue = JavaWindow("RDV_StructureManager").JavaWindow("Spatial Filter").JavaTable("UseSelectionTable").Object.getItem(iCnt).getData().toString()
						If Instr( sActValue , ArrNode ) > 0 Then
							If Cint(ArrIns) = Cint(GlobIns) Then
								Fn_RDV_SpatialFilterUseSelectionTableOperation = True
								Exit For
							Else
								GlobIns = Cint(GlobIns) + 1
							End If
						End If
					Next
					
					If iCnt > cInt(objSpatialCriteria.JavaTable("UseSelectionTable").GetROProperty("rows")) -1 Then
						Fn_RDV_SpatialFilterUseSelectionTableOperation = False
						Exit Function
					End If
				Next
			Else
				Fn_RDV_SpatialFilterUseSelectionTableOperation = False
                Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_RDV_SpatialFilterUseSelectionTableOperation ] Failed to No Items in Dialog.")
				Exit Function
			End If
			
		Case "removewithinstance" , "remove"

            If Cstr(objSpatialCriteria.JavaTable("UseSelectionTable").GetROProperty("rows")) <> "0" Then
				objSpatialCriteria.JavaTable("UseSelectionTable").SelectRow 0 
				objSpatialCriteria.JavaTable("UseSelectionTable").DeSelectRow 0 
				For iVals = 0 to Ubound(sArrValues)
					ArrInsNode = Split( sArrValues(iVals) , "@" )
					GlobIns = 1
					If Ubound(ArrInsNode) = 1 Then
						ArrNode = ArrInsNode(0)
						ArrIns = ArrInsNode(1)
					Else
						ArrNode = ArrInsNode(0)
						ArrIns = GlobIns
					End If	
					
					For iCnt = 0 to cInt(objSpatialCriteria.JavaTable("UseSelectionTable").GetROProperty("rows")) -1
						sActValue = objSpatialCriteria.JavaTable("UseSelectionTable").Object.getItem(iCnt).getData().toString()
						If Instr( sActValue , ArrNode ) > 0 Then
							If Cint(ArrIns) = Cint(GlobIns) Then
								objSpatialCriteria.JavaTable("UseSelectionTable").ActivateCell iCnt , "BOM Line Name"
								Wait 1
								'Click on Remove button
								Call Fn_Button_Click( "Fn_RDV_SpatialFilterUseSelectionTableOperation" , objSpatialCriteria , "Remove" )
								Call Fn_ReadyStatusSync(5)
								Fn_RDV_SpatialFilterUseSelectionTableOperation = True
								Exit For
							Else
								GlobIns = Cint(GlobIns) + 1
							End If
						End If
					Next
				Next
			Else
				Fn_RDV_SpatialFilterUseSelectionTableOperation = False
                Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_RDV_SpatialFilterUseSelectionTableOperation ] Failed to No Items in Dialog.")
				Exit Function
			End If

		Case "select"
			Call Fn_SISW_UI_JavaCheckBox_Operations("Fn_RDV_SpatialFilterUseSelectionTableOperation", "Set", objSpatialCriteria, "UseSelectionsFromTable", "ON")
			wait SISW_MIN_TIMEOUT
			objSpatialCriteria.JavaTable("UseSelectionTable").SelectRow 0 
			objSpatialCriteria.JavaTable("UseSelectionTable").DeSelectRow 0 

			iRowCount = cInt(objSpatialCriteria.JavaTable("UseSelectionTable").GetROProperty("rows"))
			For iCnt = 0 to iRowCount - 1
				sActValue = objSpatialCriteria.JavaTable("UseSelectionTable").Object.getItem(iCnt).getData().toString()
				If Instr(sActValue,sValues) Then
'					objSpatialCriteria.JavaTable("UseSelectionTable").ActivateCell iCnt , "Item Name"
						iTotalCols = objSpatialCriteria.JavaTable("UseSelectionTable").GetROProperty("cols")
						For iCounter = 0 To iTotalCols-1
							sColumnName = objSpatialCriteria.JavaTable("UseSelectionTable").GetColumnName(iCounter)
							If sColumnName = "BOM Line" OR sColumnName = "BOM Line Name" Then
								sColName = sColumnName
								Exit for
							End If
						Next
						objSpatialCriteria.JavaTable("UseSelectionTable").SelectCell iCnt, sColName
						If Err.Number < 0 Then
							Fn_RDV_SpatialFilterUseSelectionTableOperation = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_RDV_SpatialFilterUseSelectionTableOperation ] Failed to Select Node.")
							Exit Function
						End If
					Wait 1
					Fn_RDV_SpatialFilterUseSelectionTableOperation = True
					Exit Function
				End if
			Next

		Case "multiselect"
			Call Fn_SISW_UI_JavaCheckBox_Operations("Fn_RDV_SpatialFilterUseSelectionTableOperation", "Set", objSpatialCriteria, "UseSelectionsFromTable", "ON")
			wait SISW_MIN_TIMEOUT
			objSpatialCriteria.JavaTable("UseSelectionTable").SelectRow 0 
			objSpatialCriteria.JavaTable("UseSelectionTable").DeSelectRow 0 

			sArrValues = Split(sValues,"~")
			iRowCount = cInt(objSpatialCriteria.JavaTable("UseSelectionTable").GetROProperty("rows"))
			wait 1
			For iVals = 0 to Ubound(sArrValues)
				ArrInsNode = Split(sArrValues(iVals) , "@" )
				GlobIns = 1
				If Ubound(ArrInsNode) = 1 Then
					ArrIns = cInt(ArrInsNode(1))
				Else
					ArrIns = 1
				End If	
				ArrNode = trim(ArrInsNode(0))
				For iCnt = 0 to iRowCount - 1
					sActValue = objSpatialCriteria.JavaTable("UseSelectionTable").Object.getItem(iCnt).getData().toString()
					If  Instr( sActValue , ArrNode ) > 0 Then
						If ArrIns = 1 Then
							If iVals = 0 Then
								objSpatialCriteria.JavaTable("UseSelectionTable").SelectRow iCnt
							Else
								objSpatialCriteria.JavaTable("UseSelectionTable").ExtendRow iCnt
							End If
							Fn_RDV_SpatialFilterUseSelectionTableOperation = True
							exit for
						Else
							ArrIns = ArrIns - 1
						End If
					End if
				Next
			Next
			
		Case "removeall" , "selectall"
			Call Fn_SISW_UI_JavaCheckBox_Operations("Fn_RDV_SpatialFilterUseSelectionTableOperation", "Set", objSpatialCriteria, "UseSelectionsFromTable", "ON")
			wait SISW_MIN_TIMEOUT
			objSpatialCriteria.JavaTable("UseSelectionTable").SelectRow 0 
			objSpatialCriteria.JavaTable("UseSelectionTable").DeSelectRow 0 
			iRowCount = cInt(objSpatialCriteria.JavaTable("UseSelectionTable").GetROProperty("rows"))
			objSpatialCriteria.JavaTable("UseSelectionTable").DeselectRowsRange 0 , iRowCount - 1
			Wait 1
			For iCnt = 0 to iRowCount - 1
				If iCnt = 0 Then
					objSpatialCriteria.JavaTable("UseSelectionTable").SelectRow iCnt
				Else
					objSpatialCriteria.JavaTable("UseSelectionTable").ExtendRow iCnt
				End if
			Next
			
			If Lcase(sAction) = "removeall" Then
				'Click on Remove Button
				Call Fn_Button_Click( "Fn_RDV_SpatialFilterUseSelectionTableOperation" , objSpatialCriteria , "Remove" )		
			End If
			Call Fn_ReadyStatusSync(1)
			
			Fn_RDV_SpatialFilterUseSelectionTableOperation = True
			
		End Select
	'Close Search Dialog.
	If cBool(bCloseDialog) then objSpatialCriteria.Close
	
	If Fn_RDV_SpatialFilterUseSelectionTableOperation <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Function [ Fn_RDV_SpatialFilterUseSelectionTableOperation ] Executed successfully with case [ " & sAction& " ].") 
	Else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_RDV_SpatialFilterUseSelectionTableOperation ] is Not for case [ " & sAction& " ].") 
	End If
End Function

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_SISW_RDV_MSM_ScopesOperation
'@@
'@@    Description				 :	Function Used to perform Scope operation  in Multi-structure Manager
'@@
'@@    Parameters			   :	sAction = Action to be performed
'@@    											sBOMLine = BOMLine to be selected as a scope
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	Multi-Structure Manager should be displayed							
'@@
'@@    Examples					: Call Fn_SISW_RDV_MSM_ScopesOperation("Select","00.z120/A;1 (View)")
'@@
'@@	   History					 	:	
'@@				Developer Name				Date			Rev. No.	Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Vrushali Wani			07-Sept-2012     		1.0			Created
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Function Fn_SISW_RDV_MSM_ScopesOperation(sAction,sBOMLine)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_RDV_MSM_ScopesOperation"
	Dim objScopes, iRowCount,sText

    Fn_SISW_RDV_MSM_ScopesOperation = False

	Set  objScopes = Fn_SISW_RDV_GetObject("Scopes")

	' Verify the existance of Scopes object else open the window.
	If Fn_UI_ObjectExist("Fn_SISW_RDV_MSM_ScopesOperation", objScopes) = False Then
		bReturn = Fn_Button_Click("Fn_SISW_RDV_MSM_ScopesOperation",  JavaWindow("RDV_StructureManager"), "SearchScopeCashLessSearch")
		If bReturn <> True Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_SISW_RDV_MSM_ScopesOperation ] Failed to click on [ SearchScopeCashLessSearch ] button .") 
				Fn_SISW_RDV_MSM_ScopesOperation = False
				Exit function 
		Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Function [ Fn_SISW_RDV_MSM_ScopesOperation ] Successfully clicked on [ SearchScopeCashLessSearch ] button .") 
		End If
	End If

	Select Case sAction
	   Case "Select"
			iRowCount = objScopes.JavaTable("CurrentScopes").GetROProperty("rows")

			For iCount = 0 to iRowCount-1
				sText = objScopes.JavaTable("CurrentScopes").Object.getItem(iCount).getData().toString()
				If   cStr(sText)  =  sBOMLine Then
					objScopes.JavaTable("CurrentScopes").SelectCell iCount,"BOM Line Name"
					Fn_SISW_RDV_MSM_ScopesOperation = True
					Exit for 
				End If
			Next           
	   Case Else
		   Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_SISW_RDV_MSM_ScopesOperation ] Invalid case [ " & sAction& " ].") 
	End Select

	If Fn_SISW_RDV_MSM_ScopesOperation = True Then
			Call Fn_Button_Click("Fn_SISW_RDV_MSM_ScopesOperation",  objScopes, "OK")
			Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Function [ Fn_SISW_RDV_MSM_ScopesOperation ] Executed successfully with case [ " & sAction& " ].") 
	Else
	        Call Fn_Button_Click("Fn_SISW_RDV_MSM_ScopesOperation",  objScopes, "Cancel")
	End If
	Set objScopes = Nothing
End Function
'********************** Function to perform operations on Search Criteria panel in MSM ***************************************
'
''Function Name		 	:	Fn_SISW_RDV_MSM_SearchCriteriaOperations
'
''Description		    :	Function to perform operations on Search Criteria panel in MSM
'
''Parameters		    :	1. sAction 		: Action need to perform
'					  		2. sTab 		:	Tab Name
'					  		3. bCloseDialog	:	for future use.
'					  		4. dicSearchCri : Dictionary object to set Search Criteria data.
'								
'Return Value		    :  	True / False
'
'Pre-requisite		    :	MSM perspective should be opened.

''Examples  			:	Dim dicSearchCri
'					  		Set dicSearchCri = CreateObject("Scripting.Dictionary")

'							Dim dicSearchCri
'					  		Set dicSearchCri = CreateObject("Scripting.Dictionary")
'					  		dicSearchCri("Root") = "PACKED_000245_A/A;1 (View)"
'					  		dicSearchCri("ScopesTable_Row") = "PACKED_000245_A/A;1 (View)"
'					  		dicSearchCri("ScopesTable_Column") = "Item Type"
'					  		dicSearchCri("ScopesTable_Value") = "Item"
'					  		dicSearchCri("SpatialFilterValues") = ""
'					  		dicSearchCri("SpatialFilterTable_Row") = "PACKED_000245_A/A;1 x 5"
'					  		dicSearchCri("SpatialFilterTable_Column") = ""
'					  		dicSearchCri("SpatialFilterTable_Value") = ""
'					  		msgbox Fn_SISW_RDV_MSM_SearchCriteriaOperations("Verify","" ,"", dicSearchCri)

'History:
'	Developer Name			Date			Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Koustubh Watwe		12-Sept-2012			1.0					Created
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_RDV_MSM_SearchCriteriaOperations(sAction, sTab, bCloseDialog, dicSearchCri)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_RDV_MSM_SearchCriteriaOperations"
   	Dim objSrchRes, iCnt, iRowCount, sActValue, sText
	Dim aRows, iWaitCnt, iInstance, iInstCnt
	Dim bFlag
	Set objSrchRes = JavaWindow("RDV_StructureManager")
	Fn_SISW_RDV_MSM_SearchCriteriaOperations = False
	If bCloseDialog = "" Then bCloseDialog = False

	' selecting Tab
	If sTab <> "" Then
		If objSrchRes.JavaTab("MSM_ResultsTab").Exist(5) Then
			Call Fn_UI_JavaTab_Select("Fn_SISW_RDV_MSM_SearchCriteriaOperations",objSrchRes,"MSM_ResultsTab", sTab)
		End If
	End If

	If Fn_UI_ObjectExist("Fn_SISW_RDV_MSM_SearchCriteriaOperations",objSrchRes.JavaTable("MSM_ScopesSearchCriteriaTable") ) = False Then
        objSrchRes.JavaStaticText("MSM_SearchType").SetTOProperty "label", "Search Criteria"
		objSrchRes.JavaStaticText("MSM_SearchType").Click 1, 1,"LEFT"
		If Fn_UI_ObjectExist("Fn_SISW_RDV_MSM_SearchCriteriaOperations",objSrchRes.JavaTable("MSM_ScopesSearchCriteriaTable") ) = False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_SISW_RDV_MSM_SearchCriteriaOperations ] Can not find tab [ Search Criteria ].") 
			Set objSrchRes = Nothing
			Exit function
		End If
	End If

	Select Case sAction
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Verify"
				
				If dicSearchCri("Root") <> "" Then
					If dicSearchCri("Root") <> objSrchRes.JavaStaticText("MSM_RootSearchCriteria_Value").GetROProperty("label") Then
						Exit Function
					End If
				End If

				If dicSearchCri("ScopesTable_Row") <> "" Then
					bFlag = False
					iRowCount = cInt(objSrchRes.JavaTable("MSM_ScopesSearchCriteriaTable").GetROProperty("rows"))
					aRows = split(dicSearchCri("ScopesTable_Row"),"@")
					If UBound(aRows) > 0 Then
						iInstCnt = cInt(aRows(1))
					Else
						iInstCnt = 1
					End If
					 aRows(0) = trim( aRows(0))
					For iCnt = 0 to iRowCount - 1
						If aRows(0) = objSrchRes.JavaTable("MSM_ScopesSearchCriteriaTable").Object.getItem(iCnt).getdata().getProperty("bl_indented_title") then
							If iInstCnt = 1Then
								bFlag = True
								If dicSearchCri("ScopesTable_Value") <> "" Then
									bFlag = False
									sActValue = Fn_SISW_UI_JavaTableGetCellData("Fn_SISW_RDV_MSM_SearchCriteriaOperations", objSrchRes.JavaTable("MSM_ScopesSearchCriteriaTable"),iCnt, dicSearchCri("ScopesTable_Column"))
									If sActValue = dicSearchCri("ScopesTable_Value") Then
										bFlag = True
									End If
								End If
								Exit For
							Else
								iInstCnt = iInstCnt - 1
							End If
						End If
					Next
					If bFlag = False Then
						Exit Function
					End If
				End If

				If dicSearchCri("SpatialFilterValues") <> "" Then					
					If dicSearchCri("SpatialFilterValues") <> objSrchRes.JavaStaticText("MSM_SpatialFilterValues_Value").GetROProperty("label") Then
						Exit Function
					End If
				End If

				If dicSearchCri("SpatialFilterTable_Row") <> "" Then
					bFlag = False
					iRowCount = cInt(objSrchRes.JavaTable("MSM_SpatialFilterTargetSearchCriteriaTable").GetROProperty("rows"))
					aRows = split(dicSearchCri("SpatialFilterTable_Row"),"@")
					If UBound(aRows) > 0 Then
						iInstCnt = cInt(aRows(1))
					Else
						iInstCnt = 1
					End If
					aRows(0) = trim( aRows(0))
					For iCnt = 0 to iRowCount - 1
						If aRows(0) = objSrchRes.JavaTable("MSM_SpatialFilterTargetSearchCriteriaTable").Object.getItem(iCnt).getdata().getProperty("bl_indented_title") then
							If iInstCnt = 1 Then
								bFlag = True
								If dicSearchCri("SpatialFilterTable_Value") <> "" Then
									bFlag = False
									sActValue = Fn_SISW_UI_JavaTableGetCellData("Fn_SISW_RDV_MSM_SearchCriteriaOperations", objSrchRes.JavaTable("MSM_SpatialFilterTargetSearchCriteriaTable"),iCnt, dicSearchCri("SpatialFilterTable_Column"))
									If sActValue = dicSearchCri("ScopesTable_Value") Then
										bFlag = True
									End If
								End If
								Exit For
							Else
								iInstCnt = iInstCnt - 1
							End If
						End If
					Next
					If bFlag = False Then
						Exit Function
					End If
				End If
				Fn_SISW_RDV_MSM_SearchCriteriaOperations = True
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case Else
	            Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_SISW_RDV_MSM_SearchCriteriaOperations ] Invalid case [ " & sAction& " ].") 
	End Select
	
	If Fn_SISW_RDV_MSM_SearchCriteriaOperations = True Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Function [ Fn_SISW_RDV_MSM_SearchCriteriaOperations ] Executed successfully with case [ " & sAction& " ].") 
	End If
	Set objSrchRes = Nothing
End Function


'*********************************************************		Function to select  the Tab into Teamcenter		***********************************************************************
'Function Name		:				 Fn_MSM_TabSet

'Description			 :		 		 This function is used to select the required Tab.

'Parameters			   :	 			1.  StrTabName:Name of the Tab to be selected.
											
'Return Value		   : 				TRUE \ FALSE

'Pre-requisite			:		 		MSM application should be displayed.

'Examples				:				 Call   Fn_MSM_TabSet("Structure Search")
'													
'History:
'										Developer Name			Date			Rev. No.			Changes Done						Reviewer		Build
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Archana							01/04/2013								Added New function on basis of Fn_MyTcTabSet()
'---	----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function Fn_MSM_TabSet(StrTabName)
'	GBL_FAILED_FUNCTION_NAME="Fn_MSM_TabSet"
'	Dim oTabCtl,iCount
'	Set oTabCtl = Nothing
'	StrTabName = trim(StrTabName)
'	
'		Set oTabCtl =  JavaWindow("RDV_StructureManager").JavaObject("MSMTab")
'		 For iCount = 0 to Cint(oTabCtl.Object.getTabItemCount)-1
'			  If InStr(1, oTabCtl.Object.getItem(iCount).text, StrTabName, vbTextCompare) Then
'					 oTabCtl.Object.setSelectedTabAndNotifyListeners CInt(iCount), true				 
'					  Fn_MSM_TabSet = True
'					 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_MyTc_TabSet Execution Sucessful")
'					 Exit For
'				Else
'					 Fn_MSM_TabSet = False
'					 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_MyTc_TabSet Execution Falied Tab not Found")
'				End If
'		 Next
'	
'	 Set oTabCtl = Nothing

	GBL_FAILED_FUNCTION_NAME="Fn_MSM_TabSet"
	Dim objSelectType,objIntNoOfObjects,iCount
'	Dim oTabCtl,iCount
'	Set oTabCtl = Nothing
	StrTabName = trim(StrTabName)
'	
'	If JavaWindow("RDV_StructureManager").JavaObject("MSMTab").Exist = true Then
'		Set oTabCtl =  JavaWindow("RDV_StructureManager").JavaObject("MSMTab")
'		
'	ElseIf  JavaWindow("RDV_StructureManager").JavaTab("MSM_ResultsTab").Exist = true Then
'	    Set oTabCtl =  JavaWindow("RDV_StructureManager").JavaTab("MSM_ResultsTab")    
'		
'	End If
'	
'		'Set oTabCtl =  JavaWindow("RDV_StructureManager").JavaObject("MSMTab")
'		 For iCount = 0 to Cint(oTabCtl.Object.getTabItemCount)-1
'			
'			   If InStr(1, oTabCtl.Object.getItem(iCount).text, StrTabName, vbTextCompare) Then
'					 oTabCtl.Object.setSelectedTabAndNotifyListeners CInt(iCount), true				 
'					  Fn_MSM_TabSet = True
'					 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_MyTc_TabSet Execution Sucessful")
'					 Exit For
'				Else
'					 Fn_MSM_TabSet = False
'					 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_MyTc_TabSet Execution Falied Tab not Found")
'				End If
'		 Next
'	
'	 Set oTabCtl = Nothing

				Set objSelectType = description.Create()
				objSelectType("Class Name").value = "JavaTab"
				objSelectType("toolkit class").value = "org.eclipse.swt.custom.CTabFolder"						
				Set  objIntNoOfObjects = JavaWindow("DefaultWindow").ChildObjects(objSelectType)
	
				For icount = 0 To objIntNoOfObjects.Count-1 Step 1			
					iItemCount = cInt(objIntNoOfObjects(icount).Object.getItemCount())
					For iCounter = 0 To iItemCount- 1 Step 1
						If trim(StrTabName) = trim(objIntNoOfObjects(icount).Object.getItems().mic_arr_get(iCounter).getText()) Then
							objIntNoOfObjects(icount).Select StrTabName
							bFlag=True
							Exit For 
						End IF
					Next
					If bFlag=True Then Exit For 
				Next
				
				If bFlag=False Then
					Fn_MSM_TabSet = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Select ["+StrTabName+"] Tab.")
				Else
					Fn_MSM_TabSet = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully Selected ["+StrTabName+"] Tab.")
				End If
				
				Set objSelectType=Nothing
				Set  objIntNoOfObjects=Nothing


 End Function
'*********************************************************	Function for Spatila search criteria ***********************************************************************
'Function Name		 :			Fn_SISW_RDV_SpatialFilterOperations

'Description			 :			This function is used to set criteria for Spatial search - Proximity / 3DBox etc and to Search Results accordingly. 

'Parameters				:			1. sAction : Action to be performed
'											2. dicUserContext : Dictionary object to set data

'Return Value			:			TRUE \ FALSE

'Pre-requisite			 :			 Spatial filter window should be displayed.
	
'Note						 :			 This Function is used to set only criteria below Use Selection from table. (i.e. Proximity, 3DBox etc.)
'											 Function Fn_RDV_SpatialFilterUseSelectionTableOperation() is used for Use Selection from table.
'												
'Examples				 :			 Call Fn_SISW_RDV_SpatialFilterOperations("Set", dicSpatialFilter)
'
'											1. Search Type - 3D Box
'											Dim dicSpatialFilter
'											Set dicSpatialFilter = CreateObject( "Scripting.Dictionary" )
'											dicSpatialFilter.RemoveAll
'											dicSpatialFilter("SearchType") = "3DBox"
'											dicSpatialFilter("Enable3Dmanipulators") = "ON"
'											dicSpatialFilter("Extent") = "Minimum and Maximum"
'											dicSpatialFilter("SlideIncrement") = "0.01"
'											dicSpatialFilter("XMax") = "1.6026"
'											dicSpatialFilter("XMin") = "-4.4738"
'											dicSpatialFilter("YMax") = "-0.2213"
'											dicSpatialFilter("YMin") = "-3.8917"
'											dicSpatialFilter("ZMin") = "-2.3264"
'											dicSpatialFilter("ZMax") = "3.7500"
'											dicSpatialFilter("IncludePartsIntersect") = "OFF"
'											dicSpatialFilter("FindParts") = "Inside"
'											dicSpatialFilter("TrueShapeFiltering") = "ON"
'											dicSpatialFilter("ButtonName") = "OK"
'											dicSpatialFilter("ClickOnSearchButton") = ""
'											bReturn = Fn_SISW_RDV_SpatialFilterOperations("Set", dicSpatialFilter)
'
'											2. Search Type - Proximity
'											dicSpatialFilter("Reset") = True
'											dicSpatialFilter("UseSelectionsFromTable") = "ON"
'											dicSpatialFilter("ItemName") = "flipfone123_front_bottom"
'											dicSpatialFilter("SearchType") = "Proximity"
'											dicSpatialFilter("Distance") = "0.111"
'											dicSpatialFilter("ValidOverlaysOnly") = "ON"
'											dicSpatialFilter("TrueShapeFiltering") = "ON"
'											dicSpatialFilter("ButtonName") = "OK"
'											dicSpatialFilter("ClickOnSearchButton") = ""
'											bReturn = Fn_SISW_RDV_SpatialFilterOperations("Set", dicSpatialFilter)					
								
'History:
'										Developer Name			Date			Rev. No.		Reviewer					Build
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Pallavi Jadhav			17/09/2013		  1.0			Koustubh Watwe		20130902
'										Vivek Ahirrao			18/11/2015		  1.1			[TC1121-2015102600-18_11_2015-VivekA-Maintenance]
'																Modified case "Set" for "BOM Line" column change
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_RDV_SpatialFilterOperations(sAction, dicSpatialFilter)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_RDV_SpatialFilterOperations"
	Dim objSpatialCriteria , objApplet , iCount, sArrValues, bFlag, objDeviceReplay, sRowCount,iCnt, sNode

	Set objSpatialCriteria = Fn_SISW_RDV_GetObject("Spatial Filter")
	Set objApplet = Window("RDV_StructureManagerWindow").JavaWindow("RDV_JApplet")
	Fn_SISW_RDV_SpatialFilterOperations = False
	
	If Fn_SISW_UI_Object_Operations("Fn_SISW_RDV_SpatialFilterOperations", "Exist", objSpatialCriteria, SISW_MIN_TIMEOUT) = False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL | Function [ Fn_SISW_RDV_SpatialFilterOperations ] - Failed to find Spatial Criteria dialog.") 
		exit function
	End If
	
	Select Case sAction
		Case "Set"
			'Click on 'Reset'
			If lcase(dicSpatialFilter("Reset")) = "true" Then
				If Fn_SISW_UI_Object_Operations("Fn_SISW_RDV_SpatialFilterOperations","Enabled", objSpatialCriteria.JavaButton("Reset"), SISW_MIN_TIMEOUT) Then
					Call Fn_Button_Click("Fn_SISW_RDV_SpatialFilterOperations", objSpatialCriteria, "Reset")
				End If
			End If

			'Check 'Use Selections From Table' check box
			If dicSpatialFilter("UseSelectionsFromTable") <> "" Then
				'ON / OFF
				If Fn_SISW_UI_JavaCheckBox_Operations("Fn_SISW_RDV_SpatialFilterOperations", "Set", objSpatialCriteria, "UseSelectionsFromTable", dicSpatialFilter("UseSelectionsFromTable")) = False then exit function

				'Select BOM nodes if Check box is set to ON
				If dicSpatialFilter("UseSelectionsFromTable") = "ON" Then
					wait SISW_MIN_TIMEOUT
					objSpatialCriteria.JavaTable("UseSelectionTable").SelectRow 0 
					objSpatialCriteria.JavaTable("UseSelectionTable").DeSelectRow 0 
					wait SISW_MIN_TIMEOUT
					'Get Items to be selected
					sArrValues = Split(dicSpatialFilter("ItemName"), "~")
					Set objDeviceReplay = CreateObject("Mercury.DeviceReplay")
					For iCount = 0 to UBound(sArrValues)
						sRowCount = Fn_SISW_UI_JavaTable_Operations("Fn_SISW_RDV_SpatialFilterOperations", "GetRowCount", objSpatialCriteria , "UseSelectionTable", "", "", "", "", "", "", "")
						For iCnt = 0 To sRowCount-1
							sNode = Fn_SISW_UI_JavaTable_Operations("Fn_SISW_RDV_SpatialFilterOperations", "GetCellData", objSpatialCriteria , "UseSelectionTable", "Object.GetItem", "BOM Line Name", iCnt, "", "", "", "")	
							If Instr(sNode, dicSpatialFilter("ItemName"))>0 Then
								sArrValues(iCount) = sNode
							End If
						Next
					
'						bFlag = Fn_SISW_UI_JavaTable_Operations("Fn_SISW_RDV_SpatialFilterOperations", "ClickCell", objSpatialCriteria , "UseSelectionTable", "GetProperty", "Item Name", sArrValues(iCount), "Item Name", "", "", "")
						bFlag = Fn_SISW_UI_JavaTable_Operations("Fn_SISW_RDV_SpatialFilterOperations", "ClickCell", objSpatialCriteria , "UseSelectionTable", "GetProperty", "BOM Line Name", sArrValues(iCount), "BOM Line Name", "", "", "")
						If bFlag = False Then 
							exit for
						Else
							If iCount = 0 Then
								' KeyDown Ctrl key
								objDeviceReplay.KeyDown 29
							End If
						End If
					Next
					' release Ctrl key
					objDeviceReplay.KeyUp 29
					If bFlag = False Then Exit function
				Else
					'do nothing as after check box is set to OFF, selection table gets in disabled format			
				End If
			End If

			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
			'---- Search Type ---- 3DBox ---- 	
			If dicSpatialFilter("SearchType") = "3D Box" Then
			
				'Check '3D Box' check box
				Call Fn_SISW_UI_JavaRadioButton_Operations("Fn_SISW_RDV_SpatialFilterOperations", "Set", objSpatialCriteria, "3DBox", "ON")
				
				'Select 'Extent' type
				If dicSpatialFilter("Extent") <> "" Then
					Call Fn_List_Select("Fn_SISW_RDV_SpatialFilterOperations",objSpatialCriteria,"Extent", dicSpatialFilter("Extent"))
				End If
				
				'Set 'Enable 3D manipulators' check box to desired value
				If dicSpatialFilter("Enable3Dmanipulators") <> "" Then
					Call Fn_SISW_UI_JavaCheckBox_Operations("Fn_SISW_RDV_SpatialFilterOperations", "Set", objSpatialCriteria, "Enable3Dmanipulators", dicSpatialFilter("Enable3Dmanipulators"))
				End If
				
				'Set 'Slider Increment' value
				If dicSpatialFilter("SlideIncrement") <> "" Then
					Call Fn_SISW_UI_JavaEdit_Operations("Fn_SISW_RDV_SpatialFilterOperations", "Set",  objSpatialCriteria, "SlideIncrement", "" )
					Call Fn_SISW_UI_JavaEdit_Operations("Fn_SISW_RDV_SpatialFilterOperations", "Type",  objSpatialCriteria, "SlideIncrement", dicSpatialFilter("SlideIncrement") )
				End If

				'Set X Coordinates value - XMin 
				If dicSpatialFilter("XMin") <> "" Then
					Call Fn_SISW_UI_Spin_Edit("Fn_SISW_RDV_SpatialFilterOperations", objSpatialCriteria, "XMin", dicSpatialFilter("XMin"))
				End If

				'Set X Length value - XMax
				If dicSpatialFilter("XMax") <> "" Then
					Call Fn_SISW_UI_Spin_Edit("Fn_SISW_RDV_SpatialFilterOperations", objSpatialCriteria, "XMax", dicSpatialFilter("XMax"))
				End If

				'Set Y Coordinates value - YMin
				If dicSpatialFilter("YMin") <> "" Then
					Call Fn_SISW_UI_Spin_Edit("Fn_SISW_RDV_SpatialFilterOperations", objSpatialCriteria, "YMin", dicSpatialFilter("YMin"))
				End If

				'Set Y Length value - YMax
				If dicSpatialFilter("YMax") <> "" Then
					Call Fn_SISW_UI_Spin_Edit("Fn_SISW_RDV_SpatialFilterOperations", objSpatialCriteria, "YMax", dicSpatialFilter("YMax"))
				End If

				'Set Z Coordinates value - ZMin
				If dicSpatialFilter("ZMin") <> "" Then
					Call Fn_SISW_UI_Spin_Edit("Fn_SISW_RDV_SpatialFilterOperations", objSpatialCriteria, "ZMin", dicSpatialFilter("ZMin"))
				End If

				'Set Z Length value - ZMax
				If dicSpatialFilter("ZMax") <> "" Then
					Call Fn_SISW_UI_Spin_Edit("Fn_SISW_RDV_SpatialFilterOperations", objSpatialCriteria, "ZMax", dicSpatialFilter("ZMax"))
				End If

				'Select 'Find Parts' type
				If dicSpatialFilter("FindParts") <> "" Then
					Call Fn_List_Select("Fn_SISW_RDV_SpatialFilterOperations",objSpatialCriteria,"FindParts", dicSpatialFilter("FindParts"))
				End If

				'Set 'Include Parts Intersect' check box to desired value
				If dicSpatialFilter("IncludePartsIntersect") <> "" Then
					Call Fn_SISW_UI_JavaCheckBox_Operations("Fn_SISW_RDV_SpatialFilterOperations", "Set", objSpatialCriteria, "IncludePartsIntersect", dicSpatialFilter("IncludePartsIntersect"))
				End If

		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -	
			'---- Search Type ---- Proximity ---- 	
			ElseIf dicSpatialFilter("SearchType") = "Proximity" Then

				'Check 'Proximity' check box
				Call Fn_SISW_UI_JavaRadioButton_Operations("Fn_SISW_RDV_SpatialFilterOperations", "Set", objSpatialCriteria, "Proximity", "ON")
				'Set 'Distance' value
				If dicSpatialFilter("Distance") <> "" Then
					Call Fn_SISW_UI_JavaEdit_Operations("Fn_SISW_RDV_SpatialFilterOperations", "Set",  objSpatialCriteria, "Distance", "" )
					Call Fn_SISW_UI_JavaEdit_Operations("Fn_SISW_RDV_SpatialFilterOperations", "Type",  objSpatialCriteria, "Distance", dicSpatialFilter("Distance") )
				End If

				'Set 'Valid Overlays Only' check box to desired value
				If dicSpatialFilter("ValidOverlaysOnly") <> "" Then
					Call Fn_SISW_UI_JavaCheckBox_Operations("Fn_SISW_RDV_SpatialFilterOperations", "Set", objSpatialCriteria, "ValidOverlaysOnly", dicSpatialFilter("ValidOverlaysOnly"))
				End If

			Else
				Exit function
				Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL | Function [ Fn_SISW_RDV_SpatialFilterOperations ] - Failed to find Search type [ "+dicSpatialFilter("SearchType")+" ].") 
			End If

		'Set 'True Shape Filtering' check box to desired value
		If dicSpatialFilter("TrueShapeFiltering") <> "" Then
			Call Fn_SISW_UI_JavaCheckBox_Operations("Fn_SISW_RDV_SpatialFilterOperations", "Set", objSpatialCriteria, "TrueShapeFiltering", dicSpatialFilter("TrueShapeFiltering"))
		End If
	
		'Click on 'OK' Button
		If dicSpatialFilter("ButtonName") <> "" Then
			Call Fn_Button_Click("Fn_SISW_RDV_SpatialFilterOperations", objSpatialCriteria, dicSpatialFilter("ButtonName"))
		End If

		'Clicking on Search button
		If dicSpatialFilter("ClickOnSearchButton") <> "" Then 
			If lcase(dicSpatialFilter("ClickOnSearchButton")) = "true" Then
				Call Fn_Button_Click("Fn_SISW_RDV_SpatialFilterOperations", objApplet, "SearchPanelSearch_16Button")
			End If
		End If
		Fn_SISW_RDV_SpatialFilterOperations = True
	End Select

	If Fn_SISW_RDV_SpatialFilterOperations = True Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS | Function [ Fn_SISW_RDV_SpatialFilterOperations ] executed successfully with case [ "+sAction+" ].")
	Else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL | Failed to execute Function [ Fn_SISW_RDV_SpatialFilterOperations ] with case [ "+sAction+" ].") 
	End If

	Set objApplet = Nothing
	Set objSpatialCriteria = Nothing
	Set objDeviceReplay = Nothing
End Function
