Option Explicit

'*********************************************************	Function List***************************************************
'0.  Fn_SISW_Search_GetObject(sObjectName)
'1. Fn_QryBldr_CreateLocalQuery()
'2. Fn_QryBldr_DeleteQuery()
'3. Fn_QryBld_QueryTreeOperations()
'4. Fn_MySavedSearchesOperations()
'5. Fn_QryBldr_ShowHintsOperations()
'6. Fn_QryBldr_SearchCriteriaUpdate()
'7. Fn_QryBldr_SearchCriteriaValidate()
'8. Fn_QryBldr_ModifyLocalQuery()
'9. Fn_MyTcSrch_SpecifyQueryDetailsAndInvoke()
'10. Fn_QryBldr_ImportQuery()
'11. Fn_QryBldr_ExportQuery()
'12. Fn_Proj_CreateBasic()
'13. Fn_MyTc_SearchTreeValidate()   - Eliminated . Unused Function.
'14. Fn_MyTc_SearchHistoryOperation()
'15. Fn_QryBldr_CrtLocQryFrmExisting()
'16. Fn_QryBldr_CrtLocQryFrmExisting_Extn()
'17. Fn_MyTc_OrganizedSavedSrchOperations()
'18. Fn_MyTcSrch_SrchPreferenceOperation()
'19. Fn_MyTcSrch_ChangeSrchOperation()
'20.Fn_SrchExtMultiAppsTrgtListOperation()
'21. Fn_MyTc_ChangeSearchHistoryOperation
'22. Fn_MyTc_SrchPFFTableColSortOperation
'23. Fn_QryBldr_VerifyDetails()
'24. Fn_SISW_QryBldr_OrderByOperation()
'25. Fn_MyTcSrch_ErrorMessageVerify()
'26. Fn_QryBldr_ImportQueryExtn()
'27. Fn_SISW_Search_QuickSearch
'28. Fn_SISW_Search_QuickSearchOperations
'29. Fn_SISW_Search_SrchSavedSearchOperation
'30. Fn_SISW_Search_AdhocClassificationQuerySearchOperations
'31. Fn_SISW_Search_SimpleSearchBOTypeOprations
'32. Fn_SISW_Search_SimpleSearchPropertyTreeOprations
'33. Fn_SISW_Search_SimpleSearchEditClauseOprations
'34. Fn_SISW_Search_SimpleSearchCriteriaTableOprations
'35. Fn_SISW_Search_SearchSortOperations
'36. Fn_SISW_Search_ClassificationSearchDialogOperations
'37. Fn_SISW_Search_LoadQuery
'38. Fn_MyTcSrch_CompareReport_Operation(sAction,dicDetails,sButton)
'39. Fn_SISW_Search_QuickSearchAndVerifyError
'*********************************************************	Function List		***********************************************************************
'****************************************    Function to get Object hierarchy ***************************************
'
''Function Name		 	:	Fn_SISW_Search_GetObject
'
''Description		    :  	Function to get specified Object hierarchy.

''Parameters		    :	1. sObjectName : Object Handle name
								
''Return Value		    :  	Object \ Nothing
'
''Examples		     	:	Fn_SISW_Search_GetObject("QueryBuilderApplet")

'History:
'	Developer Name			Date			Rev. No.		Reviewer		Changes Done	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Sukhada Bakshi		 11-Dec-2012		1.0				
'-----------------------------------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------------------------------
'	Anurag Khera		 26-Sept-2013		2.0								Externalized object hierarchies
'-----------------------------------------------------------------------------------------------------------------------------------
Function Fn_SISW_Search_GetObject(sObjectName)
	Dim sObjectXMLPath
	sObjectXMLPath = Fn_GetEnvValue("User", "AutomationDir") & "\TestData\AutomationXML\ObjectXML\Search.xml"
	Set Fn_SISW_Search_GetObject = Fn_SISW_Setup_GetObjectFromXML(sObjectXMLPath, sObjectName)
		
End Function 

'#######################################################################################################################################################~
'########################################################  Create a Local Query with defined inputs as specified     ################################################~
'#
'# FUNCTION NAME:	 	  Fn_QryBldr_CreateLocalQuery
'#
'#  MyCommunity ID :	288
'#
'# MODULE: 						 Search Requirement, MyTC 
'#
'# PRE-REQUISITE:		RAC Session accessible and [Query Builder] Application loaded
'#
'# DESCRIPTION:			 Create a Local Query with defined inputs as specified
'#											 1. Input [Name] field details
'#											 2. Input [Description] field details
'#											 3. Click [Search Class] button
'#														 3a. Tree navigate to trace [Class Attribute], For Example: POM_object:WorkspaceObject:ItemRevision
'#														 3b. Select/Double Click on Tree Node of prefered [Class Attribute]
'#														 3b. Close [Class Attribute] window
'#											 4. Check the [Search Class] button label updated to requiste class
'#											 5. Set the [Display Setting] to required option [Class/All Attributes]
'#											 6. Set [Show Indented Results] option [On/Off]
'#											 7. Select attribute from [Attribute Selection] Tree
'#													>>	 Note: 	<<
'#													 7.1 Please note that the function arguement is a array of required attributes, seperated by ":"
'#													 7.2 Select the required attribute iteratively and click on [+] button to add the attribute
'#											 8. Invoke [Create] button
'#										
'# PARAMETERS   :      sQueryName: Name of the Local User Query
'#										   sQueryDescription: Description of the Local User Query
'#										   sSearchClass: Attribute Class POM Object of the Query
'#									       sDisplaySettings: Display Settings [Class/All Attributes]
'#										   bShowIndentedResults: [Show Indented Results] option [On/Off]
'#
'#									      aAttributes: Array of the class attributes to be added to the Search Query
'#													 >>   Note: ( 1.) Multiple Attributes to be seperated by "~" ( Tilde)  [ EXAMPLE - >> Attrib1~Attrib2~Attrib3]
'#  																   ( 2.) Inside Attribute  inner values to be seperated by "," (Comma)  [ EXAMPLE - >> First, Second, Third]
'#  																  ( 3.) Inside Values Reference Path  to be seperated by ":" (Colon)  [ EXAMPLE - >> Dataset:Revision
'#
'#				             	     								>>	InnerValues  = "First, Second, Third"
'#																						 First =  refpath for Attrib in main window - will activate /double click
'#				  				   								 				Second =  [For class Attrib Sel Dialog ] ->> FullRefPath Attrib to be selected.				 (Set the Class)   				
'#																									[For Class Selection Dialog} ->>EditBoxValue]
'#																					Third = Refpath for Attrib in main window - will activate /double click
'#
'#												->>	 (Multiple Attribute) Example>>  (First, Second, Third ~ First, Second, Third~ First, Second, Third)  << -
'#												->>	 (First, Second, Third ) Example>>  Refpath1,RefPath2,Refpath3<< -
'#												->>	 (Refpath1 ) Example>>  Home:Child << -
'#
'# RETURN VALUE : 		TRUE \ FALSE
'#
'# Examples	:				 Fn_QryBldr_CreateLocalQuery("LocQuery1", "Local Test Query", "Dataset", "AllAttributes", "OFF", "Dataset:Referenced By,ItemRevision:IMAN_specification,Dataset:Specifications [ ItemRevision ]:Revision~Dataset:pid" ) 
'#										Fn_QryBldr_CreateLocalQuery("LocQuery1", "Local Test Query", "Dataset", "AllAttributes:RealNames", "ON", "Dataset:pid" ) 
'#										Fn_QryBldr_CreateLocalQuery("Sample query","description test", "ItemRevision", "AllAttributes:RealNames", "ON", "ItemRevision:IMAN_specification, Dataset, ItemRevision:IMAN_specification [ Dataset ]:object_name")
'#
'#	History	:						Developer Name			Date			Version				Changes Done			Reviewer	
'#------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'#										Kavan Shah~					05-06-10		1.0							Sunil Rai
'#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'#										Koustubh W~					05-10-10		1.0	   Modified Selection of attribute from [Attribute Selection] Tree
'#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'#										Koustubh W~					11-3-11		1.0	   Modified Selection of attribute from [Attribute Selection] Tree.
'#																												added code to expand Attribute Tree node to handle special scenario
'#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'#										Koustubh W~					11-3-11		1.0	   Modified code to add selected attributes in search criteria
'#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'#										Koustubh W~					07-12-11		1.0	   Added code to select class
'#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'########################################################  Create a Local Query with defined inputs as specified     ################################################~

Public Function Fn_QryBldr_CreateLocalQuery(sQueryName, sQueryDescription, sSearchClass, sDisplaySettings, bShowIndentedResults, aAttributes)  
		GBL_FAILED_FUNCTION_NAME="Fn_QryBldr_CreateLocalQuery"
		Dim  ArrDispSet,  OuterArrAttrib, iOuterCounter, ArrAttrib,  ArrinnAttrib, iCounter
		Dim ObjQryApp, ObjQryAttribSel
		Dim jCnt, iCnt, aDummy, sPath
		Dim itemCnt, iTreeCnt, sRevisionRule, sFindText

		Set ObjQryApp =  Fn_UI_ObjectCreate( "Fn_QryBldr_CreateLocalQuery",JavaWindow("Search_QueryBuilder").JavaWindow("QueryBuilder"))

	'++++++++++<<    Chech weather the clear button enable or not   >>++++++++++
		If ObjQryApp.JavaButton("Clear").GetROProperty("enabled") = 1  Then
				Call Fn_Button_Click( "Fn_QryBldr_CreateLocalQuery", ObjQryApp, "Clear")
		End If

	'++++++++++<<    Set the Local Query   >>++++++++++
		Call Fn_List_Select("Fn_QryBldr_CreateLocalQuery", ObjQryApp , "ModifiableQueryTypes", "Local Query")

	'++++++++++<<    Input [Name] field details >>++++++++++
		Call Fn_Edit_Box("Fn_QryBldr_CreateLocalQuery",ObjQryApp,"Name",sQueryName)

	'++++++++++<<   Input [Description] field details>>++++++++++
		'[TC1122-20160504-26_05_2016-VivekA-NewDevelopment] - Added for Search new TCs
		If Instr(sQueryDescription,"$") Then
			aQueryDescription = Split(sQueryDescription,"$")
			Call Fn_Edit_Box("Fn_QryBldr_CreateLocalQuery",ObjQryApp,"Description",aQueryDescription(0))
			'Code to set Revision Rule value in List
			sRevisionRule = aQueryDescription(1)
		Else
			Call Fn_Edit_Box("Fn_QryBldr_CreateLocalQuery",ObjQryApp,"Description",sQueryDescription)		
		End If
		
'		If lcase(sSearchClass) = "itemrevision" Then
'			sSearchClass = "Item Revision"
'		ElseIf lcase(sSearchClass) = "documentrevision" Then
'			sSearchClass = "Document Revision"		
'		End If
		If lcase(sSearchClass) = "item revision" Then
			sSearchClass = "ItemRevision"
		ElseIf lcase(sSearchClass) = "document revision" Then
			sSearchClass = "DocumentRevision"		
		End If

	'++++++++++<<   Click [Search Class] button >>++++++++++
		Call Fn_CheckBox_Set("Fn_QryBldr_CreateLocalQuery", ObjQryApp, "SrchClass", "ON")
		Call Fn_Edit_Box("Fn_QryBldr_CreateLocalQuery",ObjQryApp,"Class/Attribute Selection",sSearchClass)
		Call Fn_Button_Click( "Fn_QryBldr_CreateLocalQuery", ObjQryApp, "Find")
		Call Fn_ReadyStatusSync(5)
		wait(1)
		ObjQryApp.JavaObject("Close").Click 1,1
		wait(1)

	'++++++++++<<   Set Revision Rule >>++++++++++
		If sRevisionRule<>"" Then
			'Code to set Revision Rule value in List
			Call Fn_List_Select("Fn_QryBldr_CreateLocalQuery",ObjQryApp,"RevisionRule",sRevisionRule)
			Wait 0,200
		End if
	 '++++++++++<<  Set the [Display Setting] to required option [Class/All Attributes] >>++++++++++
	 If sDisplaySettings<>"" Then
			ArrDispSet = split(sDisplaySettings, ":", -1,1)
			If Ubound(ArrDispSet) = 1 Then
					 Wait(1)
					Call Fn_CheckBox_Set("Fn_QryBldr_CreateLocalQuery", ObjQryApp, "DisplaySettings", "ON")
					Wait(1)
					Call  Fn_UI_JavaRadioButton_SetON("Fn_QryBldr_CreateLocalQuery",ObjQryApp, ArrDispSet(0))
					Wait(1)
					Call  Fn_UI_JavaRadioButton_SetON("Fn_QryBldr_CreateLocalQuery",ObjQryApp, ArrDispSet(1))
			Else
					Wait(1)
					Call Fn_CheckBox_Set("Fn_QryBldr_CreateLocalQuery", ObjQryApp, "DisplaySettings", "ON")
					Wait(1)
					Call  Fn_UI_JavaRadioButton_SetON("Fn_QryBldr_CreateLocalQuery",ObjQryApp, sDisplaySettings )
			End If
			ObjQryApp.JavaObject("Close").Click 1,1
	End If

	 '++++++++++<<  Set [Show Indented Results] option [On/Off] >>++++++++++
		Call Fn_CheckBox_Set("Fn_QryBldr_CreateLocalQuery", ObjQryApp, "ShowIndentedResults", bShowIndentedResults)
		wait 1
	'++++++++++<<   Select attribute from [Attribute Selection] Tree >>++++++++++
	
		OuterArrAttrib = split(aAttributes, "~", -1,1)
		For iOuterCounter = 0 To Ubound(OuterArrAttrib)
				ArrAttrib = split(OuterArrAttrib(iOuterCounter), ",", -1, 1)
				For iCounter = 0 to Ubound(ArrAttrib) 		
							ArrinnAttrib = split(ArrAttrib(iCounter), ":", -1, 1)
							Select Case iCounter
									Case "0"
											'ObjQryApp.JavaTree("AttributeSelectionList").Activate ArrAttrib(iCounter) 
											aDummy = split(ArrAttrib(iCounter),":")
											If uBound(aDummy) <> 0 Then
													For iCnt = 0 to ubound(aDummy) -1
																sPath = ""
																For jCnt = 0 to iCnt
																		If sPath = "" Then
																			sPath = aDummy(jCnt)
																			 If lcase(sPath) = "itemrevision" Then
																				sSearchClass = "Item Revision"
																			ElseIf lcase(sPath) = "documentrevision" Then
																				sSearchClass = "Document Revision"		
																			End If
																		Else
																			sPath = sPath & ":" & aDummy(jCnt)
																		End If
																Next
																If iCnt = 0 Then
																		itemCnt = cInt(ObjQryApp.JavaTree("AttributeSelectionList").getROProperty("items count"))
																		If itemCnt = 1 Then
																			ObjQryApp.JavaTree("AttributeSelectionList").Select sPath
																			ObjQryApp.JavaTree("AttributeSelectionList").Object.setExpandedState ObjQryApp.JavaTree("AttributeSelectionList").Object.getSelectionPath(), true
																		End If
																Else
																		ObjQryApp.JavaTree("AttributeSelectionList").Select sPath
																		ObjQryApp.JavaTree("AttributeSelectionList").Object.setExpandedState ObjQryApp.JavaTree("AttributeSelectionList").Object.getSelectionPath(), true
																		' special scenario tree node get expanded after dbl clicking on it.
																		itemCnt = cInt(ObjQryApp.JavaTree("AttributeSelectionList").GetROProperty("items count"))
																		For iTreeCnt = 0 to itemCnt - 1
																				If ObjQryApp.JavaTree("AttributeSelectionList").GetItem(iTreeCnt) = sPath then
																						If not (instr(ObjQryApp.JavaTree("AttributeSelectionList").GetItem(iTreeCnt + 1), sPath) > 0 ) Then
																								ObjQryApp.JavaTree("AttributeSelectionList").Activate sPath 
																								wait 2
																						End If
																						Exit for
																				End if 
																		Next
																End If
													Next
													wait 5
													if uBound(ArrAttrib) = 0 then
													If aDummy(uBound(aDummy))="Gov Classification" Then
														aDummy(uBound(aDummy))="Government Classification"
													End If
														ObjQryApp.JavaTree("AttributeSelectionList").Select sPath & ":" & aDummy(uBound(aDummy))
														Call Fn_Button_Click( "Fn_QryBldr_CreateLocalQuery", ObjQryApp, "Add")
													else
														ObjQryApp.JavaTree("AttributeSelectionList").Activate sPath & ":" & aDummy(uBound(aDummy))
													end If
													
											Else
													ObjQryApp.JavaTree("AttributeSelectionList").Activate ArrAttrib(iCounter)
											End If

							Case "1"
										If lcase(ArrinnAttrib(0)) = "item revision" Then
											sFindText = "ItemRevision"
										ElseIf lcase(ArrinnAttrib(0)) = "document revision" Then
											sFindText = "DocumentRevision"
										Else
											sFindText = ArrinnAttrib(0)							
										End If
                            
										'If  True = Fn_UI_ObjectExist("Fn_QryBldr_CreateLocalQuery", JavaWindow("Search_QueryBuilder").JavaWindow("QueryBuilder").JavaDialog("ClassAttributeSelection") )Then
										If  JavaWindow("Search_QueryBuilder").JavaWindow("QueryBuilder").JavaDialog("ClassAttributeSelection").exist(5) Then
												Set ObjQryAttribSel = JavaWindow("Search_QueryBuilder").JavaWindow("QueryBuilder").JavaDialog("ClassAttributeSelection") 
												Wait(2)
												Call Fn_CheckBox_Set("Fn_QryBldr_CreateLocalQuery", ObjQryAttribSel, "CAS_SrchClass", "ON")
												Wait(2)
												Call Fn_Edit_Box("Fn_QryBldr_CreateLocalQuery",ObjQryAttribSel,"CAS_Edit",sFindText )
												Wait(2)
												Call Fn_Button_Click( "Fn_QryBldr_CreateLocalQuery", ObjQryApp, "Find")
												Wait(2)
												ObjQryApp.JavaObject("Close").Click 1,1
												Wait (10)
												ObjQryAttribSel.JavaTree("CAS_SrchTree").Activate ArrAttrib(iCounter) 
										End If
										'If  True = Fn_UI_ObjectExist("Fn_QryBldr_CreateLocalQuery", JavaWindow("Search_QueryBuilder").JavaWindow("QueryBuilder").JavaDialog("ClassSelectionDialog")  )Then
										If JavaWindow("Search_QueryBuilder").JavaWindow("QueryBuilder").JavaDialog("ClassSelectionDialog").exist(5) OR JavaDialog("ClassSelectionDialog").Exist(5) Then
												If JavaWindow("Search_QueryBuilder").JavaWindow("QueryBuilder").JavaDialog("ClassSelectionDialog").exist(5) Then
													Set ObjQryAttribSel = JavaWindow("Search_QueryBuilder").JavaWindow("QueryBuilder").JavaDialog("ClassSelectionDialog")
												ElseIf JavaDialog("ClassSelectionDialog").exist(5) Then
													Set ObjQryAttribSel =JavaDialog("ClassSelectionDialog")
												End If
												Wait 5
												Call Fn_Edit_Box("Fn_QryBldr_CreateLocalQuery",ObjQryAttribSel,"SelectionField",ArrAttrib(iCounter) )
												wait 5
												Call Fn_Button_Click( "Fn_QryBldr_CreateLocalQuery", ObjQryAttribSel, "CSDFind")
												wait 5
												Call Fn_Button_Click( "Fn_QryBldr_CreateLocalQuery", ObjQryAttribSel, "CSDOK")
										End If										               
							Case "2"
										wait 5
										'ObjQryApp.JavaTree("AttributeSelectionList").Activate ArrAttrib(iCounter) 
										ObjQryApp.JavaTree("AttributeSelectionList").Select ArrAttrib(iCounter) 
										Call Fn_Button_Click( "Fn_QryBldr_CreateLocalQuery", ObjQryApp, "Add")
					End Select		
		Next
	Next

		 '++++++++++<<  Invoke [Create] button >>++++++++++
		 If  aAttributes <> "" Then
				Call Fn_Button_Click( "Fn_QryBldr_CreateLocalQuery", ObjQryApp, "Create")
				Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Pass: Local Query Successfully Created. ")   	
				Fn_QryBldr_CreateLocalQuery = True
		Else
				Call Fn_Button_Click( "Fn_QryBldr_CreateLocalQuery", ObjQryApp, "Clear")
				Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Pass: Local Query Successfully Tested. ")   	
				Fn_QryBldr_CreateLocalQuery = True
		 End If
	
	Set ObjQryApp = Nothing
	Set ObjQryAttribSel = Nothing
End Function
	
 
'####### 	E.O.F~. =  Fn_QryBldr_CreateLocalQuery		########################################################################################################~

'#######################################################################################################################################################~
'############################################################  Evaluate Query Tree Operations     ############################################################~
'#
'# FUNCTION NAME:	 	  Fn_QryBldr_DeleteQuery
'#
'#  MyCommunity ID :		298
'#
'# MODULE: 						 Search Requirement, MyTC 
'#
'# PRE-REQUISITE:		RAC Session accessible and [Query Builder] Application loaded
'#
'# DESCRIPTION:			 1. Load the Query by selecting it from Query Builder Tree
'#											2. Click [Delete] button
'#											3. Invoke [Yes] button on [Delete Confirmation] dialog
'#
'#										
'# PARAMETERS   :      sQueryPath: Tree Path of existing query to be selected
'#
'# RETURN VALUE : 	TRUE \ FALSE
'#
'#Examples	:					Fn_QryBldr_DeleteQuery("Saved Queries:__EINT_group_members")		
'#										
'#	History	:						Developer Name			Date			Version				Changes Done			Reviewer	
'#------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'#											Kavan Shah~				08-06-10		1.0										Sunil Rai						Sunil Rai
'#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'########################################################################################################################################################~
'############################################################  Evaluate Query Tree Operations     #############################################################~

Public Function Fn_QryBldr_DeleteQuery(sQueryPath)
GBL_FAILED_FUNCTION_NAME="Fn_QryBldr_DeleteQuery"
Dim iRowCnt
Dim ObjJavaApp  
Set ObjJavaApp = Fn_UI_ObjectCreate( "Fn_QryBldr_DeleteQuery",JavaWindow("Search_QueryBuilder").JavaWindow("QueryBuilder") )

	'++++++++++<<  Set Perspective >>++++++++++
	Call Fn_SetPerspective("Query Builder")

	 '++++++++++<<   Node Select  >>++++++++++
	If True =  Fn_UI_JavaTree_Activate_Select_Node("Fn_MyTc_ItemRevisionSearch", ObjJavaApp,"AbstractQueryBuilderApplicatio", sQueryPath ) Then

		Call Fn_Button_Click( "Fn_QryBldr_DeleteQuery", ObjJavaApp, "Delete")
		Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Pass: Requested Query Successfully Deleted. ")   	
		'++++++++++<<   Invoke [Yes] button on [Delete Confirmation] dialog >>++++++++++	
        Call Fn_Button_Click( "Fn_QryBldr_DeleteQuery", JavaDialog("DeleteConfirmation"), "Yes")
		Fn_QryBldr_DeleteQuery = True

	Else

		Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Fail: Requested Query Not Found. ")   		
		Fn_QryBldr_DeleteQuery = False

	End If
	
Set ObjJavaApp = Nothing
End Function
 
'####### 	E.O.F~. =  Fn_QryBldr_DeleteQuery		########################################################################################################~


'#######################################################################################################################################################~
'############################################################  Evaluate Query Tree Operations     ############################################################~
'#
'# FUNCTION NAME:	 	  Fn_QryBld_QueryTreeOperation
'#
'#  MyCommunity ID :	299
'#
'# MODULE: 						 Search Requirement, MyTC 
'#
'# PRE-REQUISITE:		Query Builder application loaded and the Query Builder Tree loaded on the LHS Pane
'#
'# DESCRIPTION:			 
'#											  Case: Query_Object_Exists
'#											 Case: Query_Object_Select
'#											Case: Query_Object_RMB_Menu_Action
'#										
'# PARAMETERS   :       sAction: This is action to be performed on the Teamcenter Query Builder Query Object Tree Node
'#										    sQueryTreePath: Name of the Teamcenter Builder Query Object Tree Node Element operated upon. 
'#										   sMenuAction: Context Menu Action
'#
'# RETURN VALUE : 		TRUE \ FALSE
'#
'#Examples	:				
'#										
'#	History	:						Developer Name			Date			Version				Changes Done			Reviewer	
'#------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'#										Kavan Shah~					08-06-10		1.0																	Sunil Rai
'#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'########################################################################################################################################################~
'############################################################  Evaluate Query Tree Operations     #############################################################~

Public Function Fn_QryBld_QueryTreeOperations(sAction, sQueryTreePath, sMenuAction)  
	GBL_FAILED_FUNCTION_NAME="Fn_QryBld_QueryTreeOperations"
Dim bExist, iNode,iCount, sNode ,arrMenuList,StrMenu
Dim ObjQryApp, ObjQryTree

Set ObjQryApp =  Fn_UI_ObjectCreate( "Fn_QryBld_QueryTreeOperations",JavaWindow("Search_QueryBuilder").JavaWindow("QueryBuilder"))
Set ObjQryTree = Fn_UI_ObjectCreate( "Fn_QryBld_QueryTreeOperations",ObjQryApp.JavaTree("AbstractQueryBuilderApplicatio"))
	Select Case sAction
		Case "Query_Object_Exists"

						If ObjQryTree.Exist Then
                			ObjQryTree.WaitProperty "enabled","1"									
							If  ObjQryTree.GetROProperty("enabled")="1" Then						
								iNode=ObjQryTree.GetROProperty("items count")				
								For  iCount=0 to iNode-1
									sNode = ObjQryTree.GetItem(iCount)
                        			If  sNode = sQueryTreePath Then
										bExist = true
										Exit For
									End If
								Next
								If True =  bExist Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Pass: Requested Query Object Exists. ")   	
									Fn_QryBld_QueryTreeOperations = True
								Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Fail: Requested Query Object Does Not Exist. ")   	
									Fn_QryBld_QueryTreeOperations = False
								End If 
							End If
						End If

		Case "Query_Object_Select"

			If True =  Fn_UI_JavaTree_Activate_Select_Node("Fn_QryBld_QueryTreeOperations",ObjQryApp,"AbstractQueryBuilderApplicatio",sQueryTreePath) Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Pass: Requested Query Object Selected. ")   	
				Fn_QryBld_QueryTreeOperations = True
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Fail: Requested Query Object Not Selected. ")   	
				Fn_QryBld_QueryTreeOperations = False
			End If 

		Case "Query_Object_RMB_Menu_Action"

			arrMenuList = split(sMenuAction, ":")
			iCount = Ubound(arrMenuList)

        	'Open context menu
			ObjQryTree.OpenContextMenu sQueryTreePath

			'Select Menu action - Level of menus
			Select Case iCount
					Case "0"
						StrMenu = JavaWindow("MyTeamcenter").WinMenu("ContextMenu").BuildMenuPath(arrMenuList(0))
					Case "1"
						StrMenu = JavaWindow("MyTeamcenter").WinMenu("ContextMenu").BuildMenuPath(arrMenuList(0),arrMenuList(1))
					Case "2"
						StrMenu = JavaWindow("MyTeamcenter").WinMenu("ContextMenu").BuildMenuPath(arrMenuList(0),arrMenuList(1),arrMenuList(2))
					Case Else
						Fn_QryBld_QueryTreeOperations = FALSE
						Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Fail: Requested RMB Menu Action Not Selected. ")   
						Exit Function
				End Select
				JavaWindow("MyTeamcenter").WinMenu("ContextMenu").Select StrMenu
				Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Fail: Requested RMB Menu Action Successfully Selected. ")   
				Fn_QryBld_QueryTreeOperations = TRUE

	End Select

Set ObjQryApp = Nothing
Set ObjQryTree = Nothing
End Function
 
'####### 	E.O.F~. =  Fn_QryBld_QueryTreeOperations		########################################################################################################~

'#######################################################################################################################################################~
'################################################    Function to do Keyword Search operation on Search Tab		###################################################~
'#
'# FUNCTION NAME:	Fn_MySavedSearchesOperations()
'#
'#  MyCommunity ID :	' ->>> NOTE: PLEASE USE "Fn_MyTc_SrchSavedSearchOperation" FUNCTION I NSTEAD OF THIS FUNCTION.<<--
'#
'# MODULE: 						 Search Requirement, MyTC 
'#
'# DESCRIPTION:			Function to do Show Hint Operation , ie verify the dialog, select the hints..
'#											Case: Add_Object_And_Verify ->> It will send the object to My Saved Searches Folder and verify its existance.
'#										
'#PARAMETERS   :  		sAction , (Two Optional Parameters for future enhancement)
'#											 
'# PRE REQUISITES:		  Object  on which action is to be performed is to be selected.
'#
'# RETURN VALUE : 		TRUE \ FALSE
'#
'#Examples	:				Fn_MySavedSearchesOperations("Add_Object_And_Verify","","")
'#										
'#	History	:						Developer Name			Date			Version				Changes Done			Reviewer	
'#------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'#										Kavan Shah~					Jenu@2010		1.0					Sunil Rai							Sunil Rai
'#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'########################################################################################################################################################~
'################################################    Function to do Keyword Search operation on Search Tab		###################################################~

Public Function   Fn_MySavedSearchesOperations(sAction, Optional1, Optional2)
	GBL_FAILED_FUNCTION_NAME="Fn_MySavedSearchesOperations"
' ->>> NOTE: PLEASE USE "Fn_MyTc_SrchSavedSearchOperation" FUNCTION I NSTEAD OF THIS FUNCTION.<<--

Dim sObjName
Dim objWindow, ObjCustSaveSrchDg, ObjAdSrchToMySvd

Set objWindow = Fn_UI_ObjectCreate( "Fn_MySavedSearchesOperations",JavaWindow("MyTeamcenter") )
Set ObjCustSaveSrchDg = JavaWindow("MyTeamcenter_Search").JavaWindow("CustomizeMySavedSrchs") ' UI Implementation is not possible on this Object
Set ObjAdSrchToMySvd =  JavaWindow("MyTeamcenter_Search").JavaWindow("AddSrchtoMySaved") ' UI Implementation is not possible on this Object

		Select Case sAction

			Case "Add_Object_And_Verify"

						'+++++++++++++++++++++
						If  True = Fn_UI_ObjectExist("Fn_MySavedSearchesOperations", objWindow.JavaTree("SearchResultTree") ) Then
								sObjName = Fn_UI_Object_GetROProperty("Fn_MySavedSearchesOperations",objWindow.JavaTree("SearchResultTree"), "value")

								If True = Fn_MyTc_SrchResltTreeOperation("PopupMenuSelect", sObjName, "Add to My Saved Searches") Then
		'								objWindow.JavaTree("SearchResultTree").OpenContextMenu sObjName
		'								JavaWindow("MyTeamcenter").WinMenu("ContextMenu").Select "Add to My Saved Searches"
										If  True = Fn_UI_ObjectExist("Fn_MySavedSearchesOperations", ObjAdSrchToMySvd )Then
												Call Fn_Button_Click("Fn_MySavedSearchesOperations",ObjAdSrchToMySvd, "OK")
												Call Fn_UI_JavaTable_SelectCell("Fn_MySavedSearchesOperations", objWindow, "Quick Links","4","0")

												If  True = Fn_UI_ObjectExist("Fn_MySavedSearchesOperations", ObjCustSaveSrchDg  ) Then
														If   True = Fn_JavaTree_Select("Fn_MySavedSearchesOperations", ObjCustSaveSrchDg, "MySavedSrchTree",  "My Saved Searches:"+sObjName)  Then 

																Call Fn_Button_Click("Fn_MySavedSearchesOperations", ObjCustSaveSrchDg, "Cancel")	
																Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS : Requested Object Successfully Verified in My Saved Searches Folder.")   
																Fn_MySavedSearchesOperations = True
														Else
																Call Fn_Button_Click("Fn_MySavedSearchesOperations", ObjCustSaveSrchDg, "Cancel")	
																Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Fail : Requested Object not found in My Saved Searches Folder.")   
																Fn_MySavedSearchesOperations = False
														End If

												Else
														Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Fail :Customize My Saved Search Dialog Does not Exist. ")   
														Fn_MySavedSearchesOperations = False	
												End If

										Else
											Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Fail :Add Search to my Saved Search Dialog Does not Exist ")   
											Fn_MySavedSearchesOperations = False	
										End If		
'																		
								Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Fail :Context Menu Operation Failed ")   
									Fn_MySavedSearchesOperations = False
								End If		
											
						Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Fail :Search Result Tree Does not Exist.  ")   
								Fn_MySavedSearchesOperations = False
						End If
						'+++++++++++++++++++++

			'Yet To Implement
'			Case "Delete"
'			Case "Rename"
'			Case "Execute"

		 Case Else 
						Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Fail :Search Result Tree Does not Exist.  ")   
						Fn_MySavedSearchesOperations = False
		End Select
	
Set objWindow = Nothing	
Set ObjCustSaveSrchDg = Nothing
Set ObjAdSrchToMySvd = Nothing
' ->>> NOTE: PLEASE USE "Fn_MyTc_SrchSavedSearchOperation" FUNCTION I NSTEAD OF THIS FUNCTION.<<--
End Function
 
'####### 	E.O.F~. =    Fn_MySavedSearchesOperations		#########################################################################################################~

'#######################################################################################################################################################~
'################################################    Function to do Keyword Search operation on Search Tab		###################################################~
'#
'# FUNCTION NAME:	Fn_QryBldr_ShowHintsOperations()
'#
'#  MyCommunity ID :	
'#
'# MODULE: 						 Search Requirement, MyTC 
'#
'# DESCRIPTION:			Function to do Show Hint Operation , ie verify the dialog, select the hints..
'#										
'#PARAMETERS   :     (sAction , sQueryPath)
'#											sAction
'#
'# RETURN VALUE : 		TRUE \ FALSE
'#
'#Examples	:				Fn_QryBldr_ShowHintsOperations()
'#										
'#	History	:						Developer Name			Date			Version				Changes Done			Reviewer	
'#------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'#										Kavan Shah~					Jenu@2010		1.0					Sunil Rai							Sunil Rai
'#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'########################################################################################################################################################~
'################################################    Function to do Keyword Search operation on Search Tab		###################################################~

Public Function   Fn_QryBldr_ShowHintsOperations(sAction, sQueryPath)
	GBL_FAILED_FUNCTION_NAME="Fn_QryBldr_ShowHintsOperations"
Dim objWindow
Set objWindow = Fn_UI_ObjectCreate( "Fn_QryBldr_ShowHintsOperations",JavaWindow("Search_QueryBuilder").JavaWindow("QueryBuilder") )

			Call Fn_CheckBox_Set("Fn_QryBldr_ShowHintsOperations", objWindow, "ShowHints", "ON")    
			Select Case sAction

					Case "VerifyDialog"

								If  True = Fn_UI_ObjectExist("Fn_QryBldr_ShowHintsOperations", objWindow.JavaObject("QueryHintPanel") ) Then
										Call Fn_CheckBox_Set("Fn_QryBldr_ShowHintsOperations", objWindow, "ChooseHint", "ON")  

										If  True = Fn_UI_ObjectExist("Fn_QryBldr_ShowHintsOperations", objWindow.JavaObject("ChooseHintDialog") ) Then
											Call Fn_Button_Click("Fn_QryBldr_ShowHintsOperations", objWindow, "OK")
											Call Fn_CheckBox_Set("Fn_QryBldr_ShowHintsOperations", objWindow, "ShowHints", "OFF")  
											Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Pass :Choose Hint Dialog is Verified.  ")   
                                            Fn_QryBldr_ShowHintsOperations = True
										Else
											Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Pass :Choose Hint Dialog Does not Exist.  ")   
											Fn_QryBldr_ShowHintsOperations = False
										End If
								Else
										Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Pass :Query HintPanel Does not Exist.  ")   
										Fn_QryBldr_ShowHintsOperations = False
								End If
		
					'Case "VerifyMsg" ' Not Implemented

					'Case "Select"' Not Implemented

					Case Else 

							Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Fail: Wrong Action Requested.  ")   	

			End Select
	
Set objWindow = Nothing	
End Function
 
'####### 	E.O.F~. =    Fn_QryBldr_ShowHintsOperations		#########################################################################################################~

'#######################################################################################################################################################~
'################################################    Create a Local Query with defined inputs as specified		######################################################~
'#
'# FUNCTION NAME:	Fn_QryBldr_SearchCriteriaUpdate
'#
'#  MyCommunity ID :	289
'#
'# MODULE: 						 Search Requirement 
'#
'# DESCRIPTION:			Create a Local Query with defined inputs as specified
'#
'#											1. Load the Query by selecting it from Query Builder Tree
'#											2. Update the Attribute definition
'#													>>  Note: Method to update the Attribute Search Criteria <<
'#													2.a Get the Row Index of the cell with requisite [Attribute Name] under Column [Attribute]
'#													2.b Update the Row with cell values [Relation, UserEntryL10Nkey, UserEntryName, Operator, DefaultValue]
'#										   3. Invoke [Modify] button
'#		
'# PRE-REQUISITE:		RAC Session accessible and [Query Builder] Application loaded with the Required Query Selected.
'#							
'#PARAMETERS   :      sAttributeName: Generic Internal Name of the attribute class
'#										   sRelation: First Column relation value [And/OR]
'#										   UserEntryL10Nkey:Name in Language
'#									   	  sOperator: Value of the relation operator [=, <, >, >=, <=, IS_NULL, IS_NOT_NULL]
'#										  sDefaultValue: Expected default expected value of search criteria
'#
'# RETURN VALUE : 		TRUE \ FALSE
'#
'#Examples	:				
'#										
'#	History	:						Developer Name			Date			Version				Changes Done			Reviewer	
'#------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'#										Kavan Shah~					05-06-10		1.0																	Sunil Rai
'#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'########################################################################################################################################################~
'################################################    Create a Local Query with defined inputs as specified		#######################################################~
Public Function Fn_QryBldr_SearchCriteriaUpdate(sAttributeNameIndex, sRelation, sAttributeName, UserEntryL10Nkey, sOperator, sDefaultValue) 
	GBL_FAILED_FUNCTION_NAME="Fn_QryBldr_SearchCriteriaUpdate"
Dim ObjSrchCrtr, iCounter, preVal, objSearchCriTable
Dim sValue, bSet, ArrItems
Dim bFlag,objLOVTreeTable,objChildObjects
Dim x, y, dr, objLOVEdit, objString

Set ObjSrchCrtr =  Fn_UI_ObjectCreate( "Fn_QryBldr_SearchCriteriaUpdate",JavaWindow("Search_QueryBuilder").JavaWindow("QueryBuilder") )
Set objSearchCriTable = ObjSrchCrtr.JavaTable("SrchCriteriaTable")
ArrItems = Array(sRelation, sAttributeName, UserEntryL10Nkey, "", sOperator, "") 
bSet = False		
If isNumeric(sAttributeNameIndex) Then 
	sAttributeNameIndex = cInt(sAttributeNameIndex)
End If
For iCounter = 0 To UBound(ArrItems)
		If  ArrItems(iCounter) <> "" Then
			' Added Code to handle Design change to set cell data as blank by Archana Dhadiwal[TC11.3_20170131_maintenance]
			If ArrItems(iCounter) = "<Blank>" Then
				 ArrItems(iCounter) = ""
			End If
			If  True = Fn_UI_JavaTable_SetCellData("Fn_QryBldr_SearchCriteriaUpdate",ObjSrchCrtr,"SrchCriteriaTable",sAttributeNameIndex, iCounter, ArrItems(iCounter))  Then
				bSet = True
			Else
				bSet = False
				Exit For
			End If
		End If
Next

If sDefaultValue<>"" Then
	Set objString = objSearchCriTable.Object.getValueAt( sAttributeNameIndex, 5)
	objSearchCriTable.Object.setValueAt objString.replaceAll(objString.toString(),""), sAttributeNameIndex, 5 
	wait 1
	objSearchCriTable.ClickCell sAttributeNameIndex,"Default Value"
	wait 2
	If ObjSrchCrtr.JavaButton("SearchCriteriaTable_dropdown_16").Exist(SISW_MIN_TIMEOUT) Then
		Call Fn_Button_Click("Fn_QryBldr_SearchCriteriaUpdate", ObjSrchCrtr,"SearchCriteriaTable_dropdown_16")
		wait 2
		bFlag = False
		Set objLOVEdit=Description.Create()
		objLOVEdit("Class Name").value="JavaEdit"
		objLOVEdit("path").value=".*LOVDisplayer.*"
		objLOVEdit("path").RegularExpression = true
		Set objChildObjects=JavaWindow("Search_QueryBuilder").JavaWindow("QueryBuilder").ChildObjects(objLOVEdit)
		objChildObjects(0).Click 1, 1
		wait 1
		'code modified to set multiple default values		
		If CBool(InStr(1,sDefaultValue,";")) = True Then
				objChildObjects(0).Set sDefaultValue
				wait 1
				If Err.Number < 0 Then
					bSet = False
				Else
					bSet = True				
				End If
		Else
				For iCounter = 1 to len(sDefaultValue)
					objChildObjects(0).Type mid(sDefaultValue, iCounter, 1)
					wait 1
				Next
				wait 5
				Set objLOVTreeTable=Description.Create()
				objLOVTreeTable("Class Name").value="JavaTable"
				objLOVTreeTable("tagname").value="LOVTreeTable"
				objLOVTreeTable("displayed").value=1
				Set objChildObjects=JavaWindow("Search_QueryBuilder").JavaWindow("QueryBuilder").ChildObjects(objLOVTreeTable)
				objChildObjects(1).SelectColumnHeader 0, "LEFT"
				wait 1
				x=objChildObjects(1).GetRoProperty("abs_x")
				y=objChildObjects(1).GetRoProperty("abs_y")
				Set dr = CreateObject("Mercury.DeviceReplay")
				dr.MouseMove Cint(x+10), Cint(y+10)
				wait 5
				bSet = Fn_SISW_UI_JavaTable_Operations("Fn_QryBldr_SearchCriteriaUpdate", "ClickCell", objChildObjects(1) , "", "GetValueAt.getDisplayableValue", 0, sDefaultValue, 0, "", "", "")
				'bFlag = bSet
				Set dr = Nothing
		End If
	Else
		ObjSrchCrtr.JavaTable("SrchCriteriaTable").SetCellData sAttributeNameIndex ,"Default Value", sDefaultValue
		'bFlag = True
		bSet = True
	End If
	bFlag = bSet
	If bFlag=False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail: Default vaue ["+sDefaultValue+"]not found in LOV Tree table")
	End If
End If
	
If  bSet = True Then	
		wait(2)
		Call Fn_Button_Click( "Fn_QryBldr_SearchCriteriaUpdate", ObjSrchCrtr, "Modify")
		Call Fn_ReadyStatusSync(2)
'		If Fn_UI_ObjectExist("Fn_QryBldr_SearchCriteriaUpdate",JavaWindow("Search_QueryBuilder").JavaWindow("User Entry Name Ignored"))Then
'			Call Fn_Button_Click( "Fn_QryBldr_SearchCriteriaUpdate", JavaWindow("Search_QueryBuilder").JavaWindow("User Entry Name Ignored"), "Ok")
'		End If
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Requested fields successfully Updated." )
    	Fn_QryBldr_SearchCriteriaUpdate = True
Else 
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed: Failed to Update Requested fields ." )
    	Fn_QryBldr_SearchCriteriaUpdate = False
End If

Set ObjSrchCrtr = Nothing
Set objSearchCriTable = Nothing
End Function
 
'####### 	E.O.F~. =    Fn_QryBldr_SearchCriteriaUpdate		##################################################################################################~

'#######################################################################################################################################################~
'################################################    Validate Query  Search Criteria as per specified details/inputs		##############################################~
'#
'# FUNCTION NAME:	Fn_QryBldr_SearchCriteriaValidate()
'#
'#  MyCommunity ID :	292
'#
'# MODULE: 						 Search Requirement, MyTC 
'#
'# DESCRIPTION:			Validate Query  Search Criteria as per specified details/inputs
'#											1. Validate the Search Criteria definition
'#												>>	Note: Method to Validate the Attribute Search Criteria <<
'#											2.a Get the Row Index of the cell with requisite [Attribute Name] under Column [Attribute]
'#											2.b Validate the Row with cell values [Relation, UserEntryL10Nkey, UserEntryName, Operator, DefaultValue]
'#
'# 
'# PRE-REQUISITE:		RAC Session accessible and [Query Builder] Application loaded. Search Criteria Loaded
'#										
'#PARAMETERS   :     sAttributeName: Generic Internal Name of the attribute class
'#										sRelation: First Column relation value [And/OR]
'#										UserEntryL10Nkey:Name in Language
'#										sOperator: Value of the relation operator [=, <, >, >=, <=, IS_NULL, IS_NOT_NULL]
'#										sDefaultValue: Expected default expected value of search criteria
'#
'#
'# RETURN VALUE : 		TRUE \ FALSE
'#
'#Examples	:				
'#										
'#	History	:						Developer Name			Date			Version				Changes Done			Reviewer	
'#------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'#										Kavan Shah~					05-06-10		1.0																	Sunil Rai
'#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'########################################################################################################################################################~
'################################################    Validate Query  Search Criteria as per specified details/inputs		##############################################~

Public Function Fn_QryBldr_SearchCriteriaValidate(sAttributeNameIndex, sRelation, sAttributeName, UserEntryL10Nkey, sUserEntryName, sOperator, sDefaultValue) 
GBL_FAILED_FUNCTION_NAME="Fn_QryBldr_SearchCriteriaValidate"
Dim ObjSrchCrtr, iCounter
Dim sValue, bMatched, ArrItems

If isNumeric(sAttributeNameIndex) Then 
	sAttributeNameIndex = cInt(sAttributeNameIndex)
End If

Set ObjSrchCrtr =  Fn_UI_ObjectCreate( "Fn_QryBldr_SearchCriteriaValidate",JavaWindow("Search_QueryBuilder").JavaWindow("QueryBuilder") )
ArrItems = Array(sRelation, sAttributeName, UserEntryL10Nkey, sUserEntryName, sOperator, sDefaultValue) 
bMatched = False		

For iCounter = 0 To UBound(ArrItems)
		If  ArrItems(iCounter) <> "" Then
			sValue =  Fn_UI_JavaTable_GetCellData("Fn_QryBldr_SearchCriteriaValidate",ObjSrchCrtr,"SrchCriteriaTable",sAttributeNameIndex, iCounter)
			If sValue = ArrItems(iCounter) Then
				bMatched = True
			Else
				bMatched = False
				Exit For
			End If
		End If
Next
	
If  bMatched = True Then	
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Validation for Requested fields successfully Verified." )
    	Fn_QryBldr_SearchCriteriaValidate = True
Else 
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed: Validation for Requested fields Failed ." )
    	Fn_QryBldr_SearchCriteriaValidate = False
End If

Set ObjSrchCrtr = Nothing	
End Function
 
'####### 	E.O.F~. =    Fn_QryBldr_SearchCriteriaValidate		#########################################################################################################~

'#######################################################################################################################################################~
'########################################################  Modify a Local Query with defined inputs as specified     ################################################~
'#
'# FUNCTION NAME:	 	  Fn_QryBldr_ModifyLocalQuery
'#
'#  MyCommunity ID :	
'#
'# MODULE: 						 Search Requirement, MyTC 
'#
'# PRE-REQUISITE:		RAC Session accessible and [Query Builder] Application loaded & Query to be modified should be loaded.
'#
'# DESCRIPTION:			 Modify a Local Query with defined inputs as specified
'#											 1. Input [Name] field details
'#											 2. Input [Description] field details
'#											 3. Click [Search Class] button
'#														 3a. Tree navigate to trace [Class Attribute], For Example: POM_object:WorkspaceObject:ItemRevision
'#														 3b. Select/Double Click on Tree Node of prefered [Class Attribute]
'#														 3b. Close [Class Attribute] window
'#											 4. Check the [Search Class] button label updated to requiste class
'#											 5. Set the [Display Setting] to required option [Class/All Attributes]
'#											 6. Set [Show Indented Results] option [On/Off]
'#											 7. Select attribute from [Attribute Selection] Tree
'#													>>	 Note: 	<<
'#													 7.1 Please note that the function arguement is a array of required attributes, seperated by ":"
'#													 7.2 Select the required attribute iteratively and click on [+] button to add the attribute
'#											 8. Invoke [Create] button
'#										
'# PARAMETERS   :      sQueryName: Name of the Local User Query
'#										   sQueryDescription: Description of the Local User Query
'#										   sSearchClass: Attribute Class POM Object of the Query
'#									       sDisplaySettings: Display Settings [Class/All Attributes]
'#										   bShowIndentedResults: [Show Indented Results] option [On/Off]
'#
'#									      aAttributes: Array of the class attributes to be added to the Search Query
'#													 >>   Note: ( 1.) Multiple Attributes to be seperated by "~" ( Tilde)  [ EXAMPLE - >> Attrib1~Attrib2~Attrib3]
'#  																   ( 2.) Inside Attribute  inner values to be seperated by "," (Comma)  [ EXAMPLE - >> First, Second, Third]
'#  																  ( 3.) Inside Values Reference Path  to be seperated by ":" (Colon)  [ EXAMPLE - >> Dataset:Revision
'#
'#				             	     								>>	InnerValues  = "First, Second, Third"
'#																						 First =  refpath for Attrib in main window - will activate /double click
'#				  				   								 				Second =  [For class Attrib Sel Dialog ] ->> FullRefPath Attrib to be selected.				 (Set the Class)   				
'#																									[For Class Selection Dialog} ->>EditBoxValue]
'#																					Third = Refpath for Attrib in main window - will activate /double click
'#
'#												->>	 (Multiple Attribute) Example>>  (First, Second, Third ~ First, Second, Third~ First, Second, Third)  << -
'#												->>	 (First, Second, Third ) Example>>  Refpath1,RefPath2,Refpath3<< -
'#												->>	 (Refpath1 ) Example>>  Home:Child << -
'#										sRemAttribNmIndex: Index Number for row to be deleted.
'#
'# RETURN VALUE : 		TRUE \ FALSE
'#
'#Examples	:				 Fn_QryBldr_ModifyLocalQuery("LocQuery1", "Local Test Query", "Dataset", "AllAttributes", "OFF", "Dataset:Referenced By,ItemRevision:IMAN_specification,Dataset:Specifications [ ItemRevision ]:Revision~Dataset:pid" , "1") 
'#										Fn_QryBldr_ModifyLocalQuery("LocQuery1", "Local Test Query", "Dataset", "AllAttributes:RealNames", "ON", "Dataset:pid" , "") 
'#
'#	History	:						Developer Name			Date			Version				Changes Done			Reviewer	
'#------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'#										Kavan Shah~					05-06-10		1.0																	Sunil Rai
'#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'########################################################################################################################################################~
'########################################################  Modify a Local Query with defined inputs as specified     ################################################~

Public Function  Fn_QryBldr_ModifyLocalQuery(sQueryName, sQueryDescription, sSearchClass, sDisplaySettings, bShowIndentedResults, aAttributes, sRemAttribNmIndex)  
GBL_FAILED_FUNCTION_NAME="Fn_QryBldr_ModifyLocalQuery"
  Dim  ArrDispSet,  OuterArrAttrib, iOuterCounter, ArrAttrib,  ArrinnAttrib, iCounter
Dim ObjQryApp, ObjQryAttribSel

Set ObjQryApp =  Fn_UI_ObjectCreate( "Fn_QryBldr_ModifyLocalQuery",JavaWindow("Search_QueryBuilder").JavaWindow("QueryBuilder"))

		
				'++++++++++<<    Input [Name] field details >>++++++++++
				If sQueryName <> "" Then
					Call Fn_Edit_Box("Fn_QryBldr_ModifyLocalQuery",ObjQryApp,"Name",sQueryName)
				End If
			
				'++++++++++<<   Input [Description] field details>>++++++++++
				If sQueryDescription <>"" Then
					Call Fn_Edit_Box("Fn_QryBldr_ModifyLocalQuery",ObjQryApp,"Description",sQueryDescription)
				End If
			
				'++++++++++<<   Click [Search Class] button >>++++++++++
				If  sSearchClass <> "" Then
					Call Fn_CheckBox_Set("Fn_QryBldr_ModifyLocalQuery", ObjQryApp, "SrchClass", "ON")
					Call Fn_Edit_Box("Fn_QryBldr_ModifyLocalQuery",ObjQryApp,"Class/Attribute Selection",sSearchClass)
					Call Fn_Button_Click( "Fn_QryBldr_ModifyLocalQuery", ObjQryApp, "Find")
					ObjQryApp.JavaObject("Close").Click 1,1
				End If
			
				 '++++++++++<<  Set the [Display Setting] to required option [Class/All Attributes] >>++++++++++
				 If sDisplaySettings <> ""  Then
					 ArrDispSet = split(sDisplaySettings, ":", -1,1)
					 If Ubound(ArrDispSet) = 1 Then
						Call Fn_CheckBox_Set("Fn_QryBldr_ModifyLocalQuery", ObjQryApp, "DisplaySettings", "ON")
						Call  Fn_UI_JavaRadioButton_SetON("Fn_QryBldr_ModifyLocalQuery",ObjQryApp, ArrDispSet(0))
						Call  Fn_UI_JavaRadioButton_SetON("Fn_QryBldr_ModifyLocalQuery",ObjQryApp, ArrDispSet(1))
					Else
						Call Fn_CheckBox_Set("Fn_QryBldr_ModifyLocalQuery", ObjQryApp, "DisplaySettings", "ON")
						Call  Fn_UI_JavaRadioButton_SetON("Fn_QryBldr_ModifyLocalQuery",ObjQryApp, sDisplaySettings )
					 End If
					ObjQryApp.JavaObject("Close").Click 1,1
				 End If
			
				 '++++++++++<<  Set [Show Indented Results] option [On/Off] >>++++++++++
				 If bShowIndentedResults <> ""  Then
					Call Fn_CheckBox_Set("Fn_QryBldr_ModifyLocalQuery", ObjQryApp, "ShowIndentedResults", bShowIndentedResults)
				 End If
			
				'++++++++++<<   Select attribute from [Attribute Selection] Tree >>++++++++++
				If  aAttributes <> "" Then
					OuterArrAttrib = split(aAttributes, "~", -1,1)
					For iOuterCounter = 0 To Ubound(OuterArrAttrib)
							ArrAttrib = split(OuterArrAttrib(iOuterCounter), ",", -1, 1)				
							For iCounter = 0 to Ubound(ArrAttrib) 		
										ArrinnAttrib = split(ArrAttrib(iCounter), ":", -1, 1)
										Select Case iCounter
										Case "0"
												ObjQryApp.JavaTree("AttributeSelectionList").Activate ArrAttrib(iCounter) 
										Case "1"
												If  True = Fn_UI_ObjectExist("Fn_QryBldr_ModifyLocalQuery", ObjQryApp.JavaDialog("ClassAttributeSelection") )Then
												
													Set ObjQryAttribSel = ObjQryApp.JavaDialog("ClassAttributeSelection") 
													Call Fn_CheckBox_Set("Fn_QryBldr_ModifyLocalQuery", ObjQryAttribSel, "CAS_SrchClass", "ON")
													Call Fn_Edit_Box("Fn_QryBldr_ModifyLocalQuery",ObjQryAttribSel,"CAS_Edit",ArrinnAttrib(0) )
													Call Fn_Button_Click( "Fn_QryBldr_ModifyLocalQuery", ObjQryApp, "Find")
													ObjQryApp.JavaObject("Close").Click 1,1
													ObjQryAttribSel.JavaTree("CAS_SrchTree").Activate ArrAttrib(iCounter) 
												End If
												If  True = Fn_UI_ObjectExist("Fn_QryBldr_ModifyLocalQuery",ObjQryApp.JavaDialog("ClassSelectionDialog")  )Then
													Set ObjQryAttribSel = ObjQryApp.JavaDialog("ClassSelectionDialog")
													Wait 2
													Call Fn_Edit_Box("Fn_QryBldr_ModifyLocalQuery",ObjQryAttribSel,"SelectionField",ArrAttrib(iCounter) )
													Call Fn_Button_Click( "Fn_QryBldr_ModifyLocalQuery", ObjQryAttribSel, "CSDFind")
													Call Fn_Button_Click( "Fn_QryBldr_ModifyLocalQuery", ObjQryAttribSel, "CSDOK")
												End If										               
										Case "2"
												wait 1
												ObjQryApp.JavaTree("AttributeSelectionList").Activate ArrAttrib(iCounter) 
										End Select		
							Next			
					Next
				End If
			
			   '++++++++++<<   Remove the Attribute Specified. >>++++++++++
			
				If sRemAttribNmIndex <> "" Then
					 If True = Fn_UI_JavaTable_SelectRow("Fn_QryBldr_ModifyLocalQuery", ObjQryApp, "SrchCriteriaTable",sRemAttribNmIndex) Then
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Fail : Local Query Search Criteria Not Deleted. ")   	
						Fn_QryBldr_ModifyLocalQuery = False
						Exit Function
					End If
				End If 

				'++++++++++<<  Invoke [Modify] button >>++++++++++
				 If True =  Fn_Button_Click( "Fn_QryBldr_ModifyLocalQuery", ObjQryApp, "Modify") Then
                        Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Pass: Local Query Successfully Modified. ")   	
						Fn_QryBldr_ModifyLocalQuery = True
				Else
						Call Fn_Button_Click( "Fn_QryBldr_ModifyLocalQuery", ObjQryApp, "Clear")
						Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Fail: Local Query Failed to be Modified. ")   	
						Fn_QryBldr_ModifyLocalQuery = False
				End If
		
	Set ObjQryApp = Nothing
	Set ObjQryAttribSel = Nothing
	End Function
	
 
'####### 	E.O.F~. =  Fn_QryBldr_ModifyLocalQuery		########################################################################################################~

'######################################################################################################################################################################################################################################################
'#	Function Name						:	 Fn_MyTcSrch_SpecifyQueryDetailsAndInvoke
'#
'#
'#	Function Teamcenter Action Association		:	Teamcenter Search
'#
'#		Teamcenter UI State Pre Condition		:	My Teamcenter Application - [Home] Tab loaded
'#
'#		Teamcenter UI State Post Condition		:	Preferred action performed under [Search Criteria] view to load requsite criteria and invoke search in My Teamcenter Context
'#
'#
'#	Function UI Control Types Exercised			:	
'#
'#		Java									:	Java Edit, Java Date, Java button, Java Combo
'#		Eclipse									:	-None-
'#		Web										:	-None-
'#		Windows									:	-None-
'#
'#
'#
'#	Function Logical Implementation Description	:	Function is designed to populate search criteria which are to be exercised
'#													under My Teamcenter Application Context and invoke search to populate search results
'#
'#
'#  	Function Parameter Details				:	Dictionary Key Value Pair
'#
'#
'#
'#	Function Return Value/Type					:	
'#	   General Case								:	Boolean
'#	   Specific (None)							:	
'#
'#
'#
'#	Function Dependancy Matrix					:	
'#		Parent Functions						:	
'#		Child Functions							:
'#
'#
'#
'#	Function Unit Test And Publication Tc Build	:	Teamcenter 9.0.20101103
'#
'#
'#	Function Usage Example						:
'#
'#
'#	Function Change History						:
'#-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'#	Created By					Date			Change Version		Function Change Unit Tests TcBuild (Build ID)			Change Review/Approval			Change Description
'#-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'#	Kavan					05/06/2010		1.0			Teamcenter 8.3 (2010????)					Mallikarjun				New Creation and Publication
'#	Mallikarjun				11/21/2010		2.0			Teamcenter 9.0 (20101103)					Mallikarjun				Updates Tc 9.0 GUI Element Change and Publication
'#	Sagar					10 dec-2010		3.0			Teamcenter 9.0(2010120100)											Updated case CreatedAfterDt-> added javabutton IR ORQuery handler code
'#	Deepak					27/06/2010		4.0			Teamcenter 9.0(2010162200)											Updated case EditBox-> Added DP for releaving Dependency From OR
'#	Deepak					27/06/2010		4.0			Teamcenter 9.0(2010162200)											Added case GenType:-> Added Case for Type Selection in General Search Pane
'#	Koustubh				07/12/2011		4.0			Teamcenter 9.1(2011113000)											Modifeid code to set values to search criteria.
'#	Manish Singh			28/06/2012		4.0			Teamcenter 9.1(2011113000)											Modifeid code to set User ID
'#	Sandeep Navghane			11/12/2012		5.0			Teamcenter 10.1 ( 1128 )											Modifeid case : "DsOwnUsrDrpDwn", "ItmRevOwningUsrDrpDwn", "OwningUsrDrpDwn", "GenOwningUsrDrpDwn",  "DsKwSrchOwnUsrDrpDwn"
'#																																																Modified case as Static Text replace with JavaTree in 10.1
'#Sandeep Navghane			31/12/2012			6.0			Teamcenter 10.1 ( 1212 )										Added Case : DropDownName
'#  Shwetambari Rathod      01/07/2014      7.0         Teamcenter 11.1(2014060400)                                         Added Case "Is This A Template:" to handle Javalist   
'######################################################################################################################################################################################################################################################
'Public Function Fn_MyTcSrch_SpecifyQueryDetailsAndInvoke(dicSearchCriteria)
'
'Dim DictItems, DictKeys, sObjElement, iCounter, bFound, innerCntr, ArrValue
'Dim ObjJavaWin,  objIntNoOfObjects, objSelectType,sAttachedText
'Dim objElementEdit,intNoOfObjects,objElement
'Set	ObjJavaWin = JavaWindow("MyTeamcenter_Search")
'
'		If ObjJavaWin.Exist Then
'
'			Call Fn_ToolbatButtonClick("Clear all search fields")
'			 
'			'Get the keys & items count from data dictionary.	
'			DictItems = dicSearchCriteria.Items
'			DictKeys = dicSearchCriteria.Keys
'			For iCounter = 0 to dicSearchCriteria.Count - 1
'					  IF IsNull(DictKeys(iCounter))  Then
'					 Else
'								IF  DictItems(iCounter) = "" Then										
'								Else
'										sObjElement = DictKeys(iCounter)
'										' Set the value as per the data dictioanry key.
'										Select case sObjElement
'
'												'++++++++++<< EditBox Set >>++++++++
'												Case  "CurrentTask", "DatasetID", "Description", "ItmItemID", "ItmName", "ItmOwningOrgID", "ItmRevAliasID", "ItmRevAliasIDCntxtNm", "ItmRevAltID", "ItmRevAltRev", "ItmRevAltRevIDCntxtNm", "ItmRevCurrentTask", "ItmRevDesc", "ItmRevItemID", "ItmRevName", "ItmRevRelStatus", "ItmRevRevision", "Keywords", "Name", "ReleaseStatus", "Revision", "PersonNm", "ObjectID", "ProjectID", "UserData1", "UserData2", "ExcludeStatus","LocationCode","Keyword","ItemName"
'
'														Select Case sObjElement
''															Case "CurrentTask"
''                                                            Case "DatasetID"
'															Case "Description"
'																sAttachedText ="Description:"
'															Case "ItmItemID"
'																sAttachedText = "Item ID:"
''															Case "ItmOwningOrgID"
''															Case "ItmRevAliasID"
''															Case "ItmRevAliasIDCntxtNm"
''															Case "ItmRevAltID"
''															Case "ItmRevAltRev"
''															Case "ItmRevAltRevIDCntxtNm"
''															Case "ItmRevCurrentTask"
''															Case "ItmRevDesc"
''															Case "ItmRevItemID"
''															Case "ItmRevRelStatus"
''															Case "ItmRevRevision"
'															Case "Keywords"
'																sAttachedText = "Keywords:"
'															Case "Keyword"
'																sAttachedText = "Keyword:"
'															Case "Name","ItmRevName","ItmName"
'																sAttachedText = "Name:"
'															Case"ItemName"
'																sAttachedText = "Item Name:"
'															Case "ReleaseStatus"
'																sAttachedText = "Release Status:"
'															Case "Revision"
'																sAttachedText = "Revision:"
''															Case "Revision"
''															Case "PersonNm"
''															Case "ObjectID"
''															Case "ProjectID"
''															Case "UserData1"
''															Case "UserData2"
''															Case "ExcludeStatus"
''															Case"LocationCode"
'															Case else
'																Call Fn_Edit_Box("Fn_MyTcSrch_SpecifyQueryDetailsAndInvoke",ObjJavaWin,sObjElement,DictItems(iCounter))
'																Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: The property "& sObjElement  & " changed successfully")
'																Call Fn_ToolbatButtonClick("Executes the search and displays the results in search result view")
'																Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Pass : MyTeamCenter Window does exist.") 
'																Fn_MyTcSrch_SpecifyQueryDetailsAndInvoke = True 
'																Exit function
'														End Select
'														
'														Set objElementEdit= Description.Create()
'                                                        Set  intNoOfObjects = JavaWindow("MyTeamcenter_Search").JavaObject("SearchTabObject").ChildObjects(objElementEdit)
'														For innerCntr = 0 to intNoOfObjects.count-1
'															If intNoOfObjects(innerCntr).getROProperty("label") = sAttachedText Then
'																Exit For
'															End If
'														Next
'														intNoOfObjects(innerCntr+1).Set DictItems(iCounter)
'														Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: The property "& sObjElement  & " changed successfully") 
'
'												'++++++++++<< EditBox Set >>++++++++
'												Case "Object ID:","Type:","Owning User:","Owning Group:","Owning Organization ID:","Name:","Item ID:","Alias ID:","Alias IdContext Name:","Alias Type:","Alternate ID:","Alternate Id Context Name:","Alternate Type:","Description:","Release Status:","Current Task:","Location Code:","User ID:","Dataset_Type:","Dataset Type:"
''- - - - -- - - - -- - - - -- - - - -- - - - -- - - - -- - - - -- - - - -- - - - -- - - - -- - - - -- - - - -- - - - -- - - - -- - - - -- - - - -- - - - -- - - - -- - - - -
'' Commented by Koustubh - 7 Dec 2011 Build :20111130
''														Set objElement = Description.Create()
''														
''														Set  intNoOfObjects = JavaWindow("MyTeamcenter_Search").JavaObject("SearchTabObject").ChildObjects(objElement)
''														For innerCntr = 0 to intNoOfObjects.count-1
''															If intNoOfObjects(innerCntr).getROProperty("label") = sObjElement Then
''																Exit For
''															End If
''														Next
''															intNoOfObjects(innerCntr+1).Set DictItems(iCounter)
'
'															' added code
'															JavaWindow("MyTeamcenter_Search").JavaEdit(replace(sObjElement,":","")).set DictItems(iCounter)
''- - - - -- - - - -- - - - -- - - - -- - - - -- - - - -- - - - -- - - - -- - - - -- - - - -- - - - -- - - - -- - - - -- - - - -- - - - -- - - - -- - - - -- - - - -- - - - -
'															wait (3)
'															Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: The property "& sObjElement  & " changed successfully")   	
'
'												'++++++++++<< Case For Type Selection in General Search >>++++++++
'												Case "GenType:"
'														Dim objType,objSubType
'														Set objType = Description.Create()
'														Set  intNoOfObjects = JavaWindow("MyTeamcenter_Search").JavaObject("SearchTabObject").ChildObjects(objType)
'														For innerCntr = 0 to intNoOfObjects.count-1
'															If intNoOfObjects(innerCntr).getROProperty("label") = "Type:" Then
'																Set objSubType = Description.Create()
'																objSubType("Class Name").Value= "JavaCheckBox"
'																If  JavaWindow("MyTeamcenter_Search").JavaObject("SearchTabObject").ChildObjects(objSubType).count > 0  Then
'																	JavaWindow("MyTeamcenter_Search").JavaWindow("Type").JavaCheckBox("Type").Set("ON")
'																	JavaWindow("MyTeamcenter_Search").JavaWindow("Type").JavaWindow("PopupGenralType").JavaEdit("Search").Set DictItems(iCounter)
'																	JavaWindow("MyTeamcenter_Search").JavaWindow("Type").JavaWindow("PopupGenralType").JavaButton("List_Find").Click
'																	JavaWindow("MyTeamcenter_Search").JavaWindow("Type").JavaWindow("PopupGenralType").JavaList("ListValue").DblClick 1,1
'																	Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: The property "& sObjElement  & " changed successfully")
'																	Exit For
'																Else
'																	intNoOfObjects(innerCntr+1).Type DictItems(iCounter)
'																	Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: The property "& sObjElement  & " changed successfully")   
'																End If		
'															End If
'														Next
'												Case "Synopsis:"
'														
'														Set objElement = Description.Create()
'														
'														Set  intNoOfObjects = JavaWindow("MyTeamcenter_Search").JavaObject("SearchTabObject").ChildObjects(objElement)
'														For innerCntr = 0 to intNoOfObjects.count-1
'															If intNoOfObjects(innerCntr).getROProperty("label") = sObjElement Then
'																Exit For
'															End If
'														Next
'															intNoOfObjects(innerCntr+1).Set DictItems(iCounter)
'															Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: The property "& sObjElement  & " changed successfully")
'												'++++++++++<< Date Set >>++++++++
'												Case "CreatedAfterDt", "CreatedBeforeDt", "ModifiedAfterDt", "ModifiedBeforeDt", "ReleasedAfterDt", "ReleasedBeforeDt"		
'														'ObjJavaWin.JavaCheckBox(sObjElement).Set "ON"
'														' new object added by sagar-> IRORQuery Button object
'														' code to click on button and invoke calender View
'														Call Fn_Button_Click( "Fn_MyTcSrch_SpecifyQueryDetailsAndInvoke", ObjJavaWin,sObjElement )
'														
'														' Code for setting Time
'														ArrValue = Split(DictItems(iCounter) ,"~",-1)
'														If Ubound(ArrValue) <> 0 Then
'															Call Fn_UI_SetDateAndTime("Fn_MyTcSrch_SpecifyQueryDetailsAndInvoke", ArrValue(0), ArrValue(1))
'														Else 
'															Call Fn_UI_SetDateAndTime("Fn_MyTcSrch_SpecifyQueryDetailsAndInvoke", DictItems(iCounter),"")
'														End If
'														
'														Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Date for Created After Successfully set.  ")         
'
'												'++++++++++<< DropDown Button (CheckBox) >>++++++++
'												Case "DsOwnGrpDrpDwn", "DsOwnUsrDrpDwn", "DsTypDrpDwn", "ItmRevAliasTypDrpDwn", "ItmRevAltRevTypDrpDwn", "ItmRevOwningGrpDrpDwn", "ItmRevOwningUsrDrpDwn", "ItmRevTypDrpDwn", "OwningGrpDrpDwn", "OwningUsrDrpDwn", "GenOwningGrpDrpDwn", "GenOwningUsrDrpDwn", "DsKwSrchDtsTypeDrpDwn", "DsKwSrchOwnGrpDrpDwn", "DsKwSrchOwnUsrDrpDwn", "UsrID"
'
'														bFound = False
'														wait 1
'														Call Fn_Button_Click( "Fn_MyTcSrch_SpecifyQueryDetailsAndInvoke", ObjJavaWin,sObjElement )
''														wait 1
''														bReturn =ObjJavaWin.JavaButton(sObjElement).GetROProperty("focused")
''														If bReturn=0 Then
''																ObjJavaWin.JavaButton(sObjElement).PressKey micTab
''														End If
'														wait 3
'														Set objSelectType=description.Create()
'														objSelectType("Class Name").value = "JavaStaticText"					
'														Set  objIntNoOfObjects = ObjJavaWin.ChildObjects(objSelectType)
'														For  innerCntr = 0 to objIntNoOfObjects.count-1
'															   If objIntNoOfObjects(innerCntr).getROProperty("label") = DictItems(iCounter) Then
'																	objIntNoOfObjects(innerCntr).Click 2,2
'																	bFound = TRUE
'																	Exit for
'															   End If
'														Next
'														If  bFound = True Then
'															Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Value for ["+DictKeys(iCounter)+"] Successfully set.  ")   	
'															Fn_MyTcSrch_SpecifyQueryDetailsAndInvoke = True
'														End If
'                
'										End select 
'								End If
'					End If                              
'			Next
'		    Call Fn_ToolbatButtonClick("Executes the search and displays the results in search result view")
'			Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Pass : MyTeamCenter Window does exist.") 
'			Fn_MyTcSrch_SpecifyQueryDetailsAndInvoke = True 
'		Else
'			Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Fail : MyTeamCenter Window does not exist.") 
'			Fn_MyTcSrch_SpecifyQueryDetailsAndInvoke = False
'		End If 
'Set objElementEdit = Nothing
'End Function


'######################################################################################################################################################################################################################################################
'#	Function Name						:	 Fn_MyTcSrch_SpecifyQueryDetailsAndInvoke
'#
'#
'#	Function Teamcenter Action Association		:	Teamcenter Search
'#
'#		Teamcenter UI State Pre Condition		:	My Teamcenter Application - [Home] Tab loaded
'#
'#		Teamcenter UI State Post Condition		:	Preferred action performed under [Search Criteria] view to load requsite criteria and invoke search in My Teamcenter Context
'#
'#
'#	Function UI Control Types Exercised			:	
'#
'#		Java									:	Java Edit, Java Date, Java button, Java Combo
'#		Eclipse									:	-None-
'#		Web										:	-None-
'#		Windows									:	-None-
'#
'#
'#
'#	Function Logical Implementation Description	:	Function is designed to populate search criteria which are to be exercised
'#													under My Teamcenter Application Context and invoke search to populate search results
'#
'#
'#  	Function Parameter Details				:	Dictionary Key Value Pair
'#
'#
'#
'#	Function Return Value/Type					:	
'#	   General Case								:	Boolean
'#	   Specific (None)							:	
'#
'#
'#
'#	Function Dependancy Matrix					:	
'#		Parent Functions						:	
'#		Child Functions							:
'#
'#
'#
'#	Function Unit Test And Publication Tc Build	:	Teamcenter 9.0.20101103
'#
'#
'#	Function Usage Example						:
'#
'#
'#	Function Change History						:
'#-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'#	Created By					Date			Change Version		Function Change Unit Tests TcBuild (Build ID)			Change Review/Approval			Change Description
'#-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'#	Kavan					05/06/2010		1.0			Teamcenter 8.3 (2010????)					Mallikarjun				New Creation and Publication
'#	Mallikarjun				11/21/2010		2.0			Teamcenter 9.0 (20101103)					Mallikarjun				Updates Tc 9.0 GUI Element Change and Publication
'#	Sagar					10 dec-2010		3.0			Teamcenter 9.0(2010120100)											Updated case CreatedAfterDt-> added javabutton IR ORQuery handler code
'#	Deepak					27/06/2010		4.0			Teamcenter 9.0(2010162200)											Updated case EditBox-> Added DP for releaving Dependency From OR
'#	Deepak					27/06/2010		4.0			Teamcenter 9.0(2010162200)											Added case GenType:-> Added Case for Type Selection in General Search Pane
'#	Koustubh				07/12/2011		4.0			Teamcenter 9.1(2011113000)											Modifeid code to set values to search criteria.
'#	Manish Singh			28/06/2012		4.0			Teamcenter 9.1(2011113000)											Modifeid code to set User ID
'#	Sandeep Navghane		11/12/2012		5.0			Teamcenter 10.1 ( 1128 )											Modifeid case : "DsOwnUsrDrpDwn", "ItmRevOwningUsrDrpDwn", "OwningUsrDrpDwn", "GenOwningUsrDrpDwn",  "DsKwSrchOwnUsrDrpDwn"
'#																																													Modified case as Static Text replace with JavaTree in 10.1
'#	Sandeep Navghane		31/12/2012		6.0			Teamcenter 10.1 ( 1212 )										Added Case : DropDownName
'#  Shwetambari Rathod      01/07/2014      7.0         Teamcenter 11.1(2014060400)                                         Added Case "Is This A Template:" to handle Javalist   
'#  Shwetambari Rathod      12/06/2015                  Teamcenter 11.2(2015052700)                                         Added case : "Activity Number:","Fault Code:",""Is Failure:","Discovered Before:","Discovered After:","Initiated Before:","Initiated After:","Due Before:","Due After:","Severity"	for MRO specific search operations	
'######################################################################################################################################################################################################################################################
 Public Function Fn_MyTcSrch_SpecifyQueryDetailsAndInvoke(dicSearchCriteria)
 		GBL_FAILED_FUNCTION_NAME="Fn_MyTcSrch_SpecifyQueryDetailsAndInvoke"
		Dim DictItems, DictKeys, sObjElement, iCounter, bFound, innerCntr, ArrValue
		Dim ObjJavaWin,  objIntNoOfObjects, objSelectType,sAttachedText
		Dim objElementEdit,intNoOfObjects,objElement
		Dim WshShell,bFlag

		Set	ObjJavaWin = JavaWindow("MyTeamcenter_Search")

		If ObjJavaWin.Exist Then
			Call Fn_ToolbatButtonClick("Clear all search fields")
			wait 2
			Call Fn_SyncTCObjects()
			wait 1
			'Get the keys & items count from data dictionary.	
			DictItems = dicSearchCriteria.Items
			DictKeys = dicSearchCriteria.Keys
			For iCounter = 0 to dicSearchCriteria.Count - 1
				bFlag = False
					  IF IsNull(DictKeys(iCounter))  Then
					 Else
								IF  DictItems(iCounter) = "" Then										
								Else
										sObjElement = DictKeys(iCounter)
										' Set the value as per the data dictioanry key.
										Select case sObjElement
												Case  "PartNumber", "ItmOwningOrgID", "ItmRevAliasID", "ItmRevAliasIDCntxtNm", "ItmRevAltID", "ItmRevAltRev", "ItmRevAltRevIDCntxtNm", "ItmRevCurrentTask", "ItmRevDesc", "ItmRevItemID", "ItmRevRelStatus", "ItmRevRevision", "PersonNm", "ObjectID", "ProjectID", "UserData1", "UserData2", "ExcludeStatus","LocationCode","ProjectUser","GroupName"
													Call Fn_Edit_Box("Fn_MyTcSrch_SpecifyQueryDetailsAndInvoke",ObjJavaWin,sObjElement,DictItems(iCounter))
													Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: The property "& sObjElement  & " changed successfully")
													Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Pass : MyTeamCenter Window does exist.") 
													Call Fn_KeyBoardOperation("SendKeys", "{TAB}")
													
												Case "p2Boolean1"		' Added new case for Search by Jotiba [TC1123-20160629-15_08_2016-JotibaT-Maintenance] 
													Select Case sObjElement
														Case "p2Boolean1"
															sObjElement = sObjElement & ":"
															DictItems(iCounter)= Ucase(DictItems(iCounter))
													End Select
													ObjJavaWin.JavaStaticText("srch_Type").SetTOProperty "label", sObjElement
													If Fn_SISW_UI_JavaList_Operations("Fn_MyTcSrch_SpecifyQueryDetailsAndInvoke", "Select", ObjJavaWin, "srch_ListBox", DictItems(iCounter), "", "")=False Then
														Fn_MyTcSrch_SpecifyQueryDetailsAndInvoke = False
														Exit function
													End If
												'++++++++++<< EditBox Set >>++++++++
												' Added new cases for Custom Attributes by Jotiba from TC1016
											Case  "CurrentTask", "DatasetID", "Description", "ItmItemID", "ItmName", "ItmRevName", "Keywords", "Name", "ReleaseStatus", "Revision", "Keyword","ItemName","Content Version","DocumentTitle","IssuingAuthority" , "ID","project_name","Owning User","Activity Number","Fault Code","p2Candid_String","p2Char1","p2Char2","p2Char3","p2Char4","p2Double1","p2Double2","p2Double3","p2Double4","p2Double5","p2Integer1","p2Integer2","p2Integer3","p2Integer4","p2String1","p2String2","p2Unique_Int1","p2Unique_Int2","OriginalFileName","Class","Object Identifier"
													Select Case sObjElement
															Case "project_name"
																sAttachedText = "project_name:"
															Case "CurrentTask"
																sAttachedText = "Current Task:"
															Case "DatasetID"
																sAttachedText = "Dataset ID:"
															Case "Description"
																sAttachedText ="Description:"
															Case "ItmItemID"
																sAttachedText = "Item ID:"
															Case "Keywords"
																sAttachedText = "Keywords:"
															Case "Keyword"
																sAttachedText = "Keyword:"
															Case "Name","ItmRevName","ItmName"
																sAttachedText = "Name:"
															Case"ItemName"
																sAttachedText = "Item Name:"
															Case "ReleaseStatus"
																sAttachedText = "Release Status:"
															Case "Revision"
																sAttachedText = "Revision:"
															Case "Content Version"
																sAttachedText = "Content Version:"
															Case "DocumentTitle"
																sAttachedText = "Document Title:"
															Case "IssuingAuthority"
																sAttachedText = "Issuing Authority:"
															Case "ID"                           'Added Case for [ Audit - General Logs ] Search by Harshal Tanpure : 13-September-2012 : Teamcenter 10 (20120905.00) 
																sAttachedText = "ID:"
															Case "Owning User"  'Added Case by Sanjeet on 27-Feb-2013
																sAttachedText = "Owning User:"
															Case "Activity Number","Fault Code","p2Candid_String","p2Char1","p2Char2","p2Char3","p2Char4","p2Double1","p2Double2","p2Double3","p2Double4","p2Double5","p2Integer1","p2Integer2","p2Integer3","p2Integer4","p2String1","p2String2","p2Unique_Int1","p2Unique_Int2"
																sAttachedText = sObjElement+":"
															Case "OriginalFileName"
																sAttachedText = "Original File Name:"
															Case "Class"
																sAttachedText ="Class:"
															Case "Object Identifier"
																sAttachedText ="Object Identifier:"
													End Select
														ObjJavaWin.JavaStaticText("srch_Type").SetTOProperty "label", sAttachedText
														If Fn_UI_ObjectExist("Fn_MyTcSrch_SpecifyQueryDetailsAndInvoke",ObjJavaWin.JavaEdit("srch_EditBox")) Then
															Call Fn_Edit_Box("Fn_MyTcSrch_SpecifyQueryDetailsAndInvoke", ObjJavaWin,"srch_EditBox",DictItems(iCounter))
															Call Fn_KeyBoardOperation("SendKey", "{TAB}")
														Else
																Fn_MyTcSrch_SpecifyQueryDetailsAndInvoke = False
																Exit function
														End If
										
												'++++++++++<< EditBox Set >>++++++++
												Case "Model Name:" , "Partition ID:" , "Partition Name:" , "Object ID:","Type:","Owning User:","Owning Group:","Owning Organization ID:","Name:","Item ID:","Alias ID:","Alias IdContext Name:","Alias Type:","Alternate ID:","Alternate Id Context Name:","Alternate IdContext Name:","Alternate Type:","Description:","Release Status:","Public ID:","Current Task:","Location Code:","User ID:","Dataset_Type:","Dataset Type:","Design Element Name:","Design Element ID:","Name:","License ID:" , "Dataset Name:","Activity Number:","Fault Code:","Message Subject:","Application Type:","Propagation Group:","License Level:","Status:"
																																						
														ObjJavaWin.JavaStaticText("srch_Type").SetTOProperty "label", sObjElement
														 JavaWindow("MyTeamcenter").JavaButton("Clear").SetTOProperty "label","More...>>>"   ''Modified by vidya 15/06/2012
'														  If JavaWindow("MyTeamcenter").JavaButton("Clear").Exist(1) Then
														If Fn_SISW_UI_Object_Operations("Fn_MyTcSrch_SpecifyQueryDetailsAndInvoke","Exist", JavaWindow("MyTeamcenter").JavaButton("Clear"),"") Then
																JavaWindow("MyTeamcenter").JavaButton("Clear").Click micLeftBtn 
														 End If 
'														If Fn_UI_ObjectExist("Fn_MyTcSrch_SpecifyQueryDetailsAndInvoke",ObjJavaWin.JavaEdit("srch_EditBox")) Then
														If Fn_SISW_UI_Object_Operations("Fn_MyTcSrch_SpecifyQueryDetailsAndInvoke","Exist", ObjJavaWin.JavaEdit("srch_EditBox"),"") Then
															Call Fn_Edit_Box("Fn_MyTcSrch_SpecifyQueryDetailsAndInvoke", ObjJavaWin,"srch_EditBox",DictItems(iCounter))
															'Need Key Press
															wait 2
'														   Call Fn_KeyBoardOperation("SendKeys", "{TAB}~{ENTER}")
															Call Fn_KeyBoardOperation("SendKeys", "{ENTER}")
															Fn_MyTcSrch_SpecifyQueryDetailsAndInvoke = True
														Else
															Fn_MyTcSrch_SpecifyQueryDetailsAndInvoke = False
															Exit function
														End If
												'++++++++++<< Case For Type Selection in General Search >>++++++++
												
												'------------------- Javalist Set added by shweta rathod--------------------------------------------------
												Case "Is This A Template:","Is Failure:"
														ObjJavaWin.JavaStaticText("srch_Type").SetTOProperty "label", sObjElement
														 JavaWindow("MyTeamcenter").JavaButton("Clear").SetTOProperty "label","More...>>>"   
														 
'														  If JavaWindow("MyTeamcenter").JavaButton("Clear").Exist(1) Then
														  If Fn_SISW_UI_Object_Operations("Fn_MyTcSrch_SpecifyQueryDetailsAndInvoke","Exist", JavaWindow("MyTeamcenter").JavaButton("Clear"),"") Then
																JavaWindow("MyTeamcenter").JavaButton("Clear").Click micLeftBtn 
															 End If 
'														If Fn_UI_ObjectExist("Fn_MyTcSrch_SpecifyQueryDetailsAndInvoke",ObjJavaWin.JavaList("DateList")) Then
														If Fn_SISW_UI_Object_Operations("Fn_MyTcSrch_SpecifyQueryDetailsAndInvoke","Exist", ObjJavaWin.JavaList("DateList"),"") Then
															ObjJavaWin.JavaList("DateList").Select DictItems(iCounter) 
															Fn_MyTcSrch_SpecifyQueryDetailsAndInvoke = True
														Else
															Fn_MyTcSrch_SpecifyQueryDetailsAndInvoke = False
															Exit function
														End If
												'--------------------------- Case For Javalist Selection in Content management Search >>++++++++
												Case "GenType:"
														Dim objType,objSubType
														Set objType = Description.Create()
														Set  intNoOfObjects = JavaWindow("MyTeamcenter_Search").JavaObject("SearchTabObject").ChildObjects(objType)
														For innerCntr = 0 to intNoOfObjects.count-1
															If intNoOfObjects(innerCntr).getROProperty("label") = "Type:" Then
																Set objSubType = Description.Create()
																objSubType("Class Name").Value= "JavaCheckBox"
																If  JavaWindow("MyTeamcenter_Search").JavaObject("SearchTabObject").ChildObjects(objSubType).count > 0  Then
																	JavaWindow("MyTeamcenter_Search").JavaWindow("Type").JavaCheckBox("Type").Set("ON")
																	JavaWindow("MyTeamcenter_Search").JavaWindow("Type").JavaWindow("PopupGenralType").JavaEdit("Search").Set DictItems(iCounter)
																	JavaWindow("MyTeamcenter_Search").JavaWindow("Type").JavaWindow("PopupGenralType").JavaButton("List_Find").Click
																	JavaWindow("MyTeamcenter_Search").JavaWindow("Type").JavaWindow("PopupGenralType").JavaList("ListValue").DblClick 1,1
																	Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: The property "& sObjElement  & " changed successfully")
																	Exit For
																Else
																	intNoOfObjects(innerCntr+1).Type DictItems(iCounter)
																	Wait 2
																	 Call Fn_KeyBoardOperation("SendKeys", "{TAB}") 
																	wait 1
																	Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: The property "& sObjElement  & " changed successfully")   
																End If		
															End If
														Next
												Case "Synopsis:"
														
'														Set objElement = Description.Create()
'														
'														Set  intNoOfObjects = JavaWindow("MyTeamcenter_Search").JavaObject("SearchTabObject").ChildObjects(objElement)
'														For innerCntr = 0 to intNoOfObjects.count-1
'															If intNoOfObjects(innerCntr).getROProperty("label") = sObjElement Then
'																Exit For
'															End If
'														Next
'														intNoOfObjects(innerCntr+1).Set DictItems(iCounter)
'														Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: The property "& sObjElement  & " changed successfully")
														
														ObjJavaWin.JavaStaticText("srch_Type").SetTOProperty "label", sObjElement
														If Fn_UI_ObjectExist("Fn_MyTcSrch_SpecifyQueryDetailsAndInvoke",ObjJavaWin.JavaEdit("srch_EditBox")) Then
															Call Fn_Edit_Box("Fn_MyTcSrch_SpecifyQueryDetailsAndInvoke", ObjJavaWin,"srch_EditBox",DictItems(iCounter))
														Else
																Fn_MyTcSrch_SpecifyQueryDetailsAndInvoke = False
																Exit function
														End If
												'++++++++++<< Date Set >>++++++++
												Case "CreatedAfterDt", "CreatedBeforeDt", "ModifiedAfterDt", "ModifiedBeforeDt", "ReleasedAfterDt", "ReleasedBeforeDt","DateCreatedDt","Discovered Before:","Discovered After:","Initiated Before:","Initiated After:","Due Before:","Due After:","Last Modified Date:"	
															Select Case sObjElement
																Case "CreatedAfterDt"
																	sAttachedText = "Created After:"
																Case "CreatedBeforeDt"
																	sAttachedText = "Created Before:"
																Case "ModifiedAfterDt"
																	sAttachedText = "Modified After:"
																Case "ModifiedBeforeDt"
																	sAttachedText = "Modified Before:"
																Case "ReleasedAfterDt"
																	sAttachedText = "Released After:"
																Case "ReleasedBeforeDt"
																	sAttachedText = "Modified Before:"
																Case "DateCreatedDt"
																	sAttachedText = "Date Created:"
																Case "Discovered Before:","Discovered After:","Initiated Before:","Initiated After:","Due Before:","Due After:"
																	sAttachedText = sObjElement
																	bFlag = True
																Case "Last Modified Date:"
																	sAttachedText = "Last Modified Date:"
																	bFlag = True
															End Select
															ObjJavaWin.JavaStaticText("srch_Type").SetTOProperty "label", sAttachedText
														'Call Fn_Button_Click( "Fn_MyTcSrch_SpecifyQueryDetailsAndInvoke", ObjJavaWin,"srch_DateButton")
'															ArrValue = Split(DictItems(iCounter) ,"~",-1)
'															If Ubound(ArrValue) <> 0 Then
'																Call Fn_UI_SetDateAndTime("Fn_MyTcSrch_SpecifyQueryDetailsAndInvoke", ArrValue(0), ArrValue(1))
'															Else 
'																Call Fn_UI_SetDateAndTime("Fn_MyTcSrch_SpecifyQueryDetailsAndInvoke", DictItems(iCounter),"")
'															End If
 													   
															ArrValue = Split(DictItems(iCounter) ,"~",-1)
															If DictItems(iCounter) <> "" Then
																If bFlag <> True then
																	ObjJavaWin.JavaEdit(Replace(Replace(sAttachedText," ",""),":","")).set ArrValue(0)
																else
																	ObjJavaWin.JavaEdit("srch_EditBox").Set ArrValue(0)
																End if
																wait 2
																Set WshShell = CreateObject("WScript.Shell")
																WshShell.SendKeys "{TAB}"
																Set WshShell = Nothing
																wait 2
																If  ubound (ArrValue) = 1 Then
																	ObjJavaWin.JavaList("DateList").Select ArrValue(1)
																End If
															End If
                  			
															
															Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Date for Created After Successfully set.  ")  
                                                 
                                              			 ' Added new cases for Custom Attributes by Jotiba [TC1123-20160629-15_08_2016-JotibaT-Maintenance] ------------------------
												Case "p2Date1","p2Date2","p2Date3"
														sAttachedText = sObjElement+":"
														ArrValue = Split(DictItems(iCounter) ,"~",-1)
														ObjJavaWin.JavaStaticText("srch_Type").SetTOProperty "label", sAttachedText
													      If ArrValue(0)<>"" Then
													      		If Fn_SISW_UI_Object_Operations("Fn_MyTcSrch_SpecifyQueryDetailsAndInvoke","Exist", ObjJavaWin.JavaEdit("srch_EditBox"),"") Then
'															      If Fn_UI_ObjectExist("Fn_MyTcSrch_SpecifyQueryDetailsAndInvoke",ObjJavaWin.JavaEdit("srch_EditBox")) Then
																	Call Fn_Edit_Box("Fn_MyTcSrch_SpecifyQueryDetailsAndInvoke", ObjJavaWin,"srch_EditBox",ArrValue(0))
																Else
																		Fn_MyTcSrch_SpecifyQueryDetailsAndInvoke = False
																		Exit function
																End If
													      End If
													       If ArrValue(1)<>"" Then
														       		If Fn_SISW_UI_JavaList_Operations("Fn_MyTcSrch_SpecifyQueryDetailsAndInvoke", "Select", ObjJavaWin, "srch_ListBox", ArrValue(1), "", "")=False Then
																		Fn_MyTcSrch_SpecifyQueryDetailsAndInvoke = False
																		Exit function
																	End If
													       End If 
													
				                                               Case "p2LOV3_Ind1","p2LOV3_Ind1_Sub","p2LOV4_Ind1","p2LOV4_Ind1_Sub","p2string_LOV","p2sub_LOV1","p2LOV5_Ind1","p2LOV5_Ind1_Sub","Checked-Out by User","Group","Role"
				                                                        sAttachedText = sObjElement+":"
																		ObjJavaWin.JavaStaticText("srch_Type").SetTOProperty "label", sAttachedText
																		If Fn_SISW_UI_Object_Operations("Fn_MyTcSrch_SpecifyQueryDetailsAndInvoke","Exist", ObjJavaWin.JavaEdit("srch_EditBox"),"") Then
'																		If Fn_UI_ObjectExist("Fn_MyTcSrch_SpecifyQueryDetailsAndInvoke",ObjJavaWin.JavaEdit("srch_EditBox")) Then
																			Call Fn_Edit_Box("Fn_MyTcSrch_SpecifyQueryDetailsAndInvoke", ObjJavaWin,"srch_EditBox",DictItems(iCounter))
																			Call Fn_KeyBoardOperation("SendKeys", "{TAB}")
																		Else
																				Fn_MyTcSrch_SpecifyQueryDetailsAndInvoke = False
																				Exit function
																		End If     
                                                '----------------------------------------------------------------------------------------------------------------------------------
												'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
												'Modified below case as Static Text replace with JavaTree in 10.1 By Sandeep N
												Case "DsOwnUsrDrpDwn", "ItmRevOwningUsrDrpDwn", "OwningUsrDrpDwn", "GenOwningUsrDrpDwn",  "DsKwSrchOwnUsrDrpDwn", "DsOwnGrpDrpDwn", "ItmRevOwningGrpDrpDwn", "OwningGrpDrpDwn", "GenOwningGrpDrpDwn",  "DsKwSrchOwnGrpDrpDwn", "DsTypDrpDwn",  "DsKwSrchDtsTypeDrpDwn","User ID:","DropDownName","Apply Class Name:","StyleSheetType","StyleSheetResourceContentType","Severity"
													Select Case sObjElement
														'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
														Case "DsOwnUsrDrpDwn", "ItmRevOwningUsrDrpDwn", "OwningUsrDrpDwn", "GenOwningUsrDrpDwn",  "DsKwSrchOwnUsrDrpDwn"
															sAttachedText = "Owning User:"
														Case "DsOwnGrpDrpDwn", "ItmRevOwningGrpDrpDwn", "OwningGrpDrpDwn", "GenOwningGrpDrpDwn",  "DsKwSrchOwnGrpDrpDwn"
															sAttachedText = "Owning Group:"
														Case "DsTypDrpDwn",  "DsKwSrchDtsTypeDrpDwn"
															sAttachedText = "Dataset Type:"
														Case "User ID:"
															sAttachedText = "User ID:"
														Case "DropDownName"
															sAttachedText = "Name:"
														Case "Apply Class Name:"
															sAttachedText = "Apply Class Name:"
														' For Content Mgmt
														Case "StyleSheetType"
															sAttachedText = "Style Sheet Type:"
														Case "StyleSheetResourceContentType"
															sAttachedText = "Style Sheet Resource Content Type:"
														Case "Severity"
															sAttachedText = "Severity:"
													End Select
													JavaWindow("MyTeamcenter").JavaButton("Clear").SetTOProperty "label","More...>>>"   
													If Fn_SISW_UI_Object_Operations("Fn_MyTcSrch_SpecifyQueryDetailsAndInvoke","Exist", JavaWindow("MyTeamcenter").JavaButton("Clear"),"") Then
'													 If JavaWindow("MyTeamcenter").JavaButton("Clear").Exist(1) Then
													 	JavaWindow("MyTeamcenter").JavaButton("Clear").Click micLeftBtn 
													 End If
													 
													JavaWindow("MyTeamcenter_Search").JavaStaticText("srch_Type").SetTOProperty "label",sAttachedText
													JavaWindow("MyTeamcenter_Search").JavaButton("srch_MultipleDropDown").Click
													wait 2
													
													Call Fn_KeyBoardOperation("SendKeys", "{TAB}")
													wait 1
													Call Fn_KeyBoardOperation("SendKeys", "{DOWN}")
													wait 1
													Dim iLastItem,i
													For i=0 to 2
														iLastItem=JavaWindow("MyTeamcenter_Search").JavaWindow("SearchCriteriaTreeShell").JavaTree("Tree").GetROProperty("items count")
														iLastItem=iLastItem-1
														JavaWindow("MyTeamcenter_Search").JavaWindow("SearchCriteriaTreeShell").JavaTree("Tree").Select "#"&Cstr(iLastItem)
														wait 1
													Next
													JavaWindow("MyTeamcenter_Search").JavaWindow("SearchCriteriaTreeShell").JavaTree("Tree").Activate  DictItems(iCounter)
													wait 1
'													If not JavaWindow("MyTeamcenter_Search").JavaWindow("SearchCriteriaTreeShell").JavaTree("Tree").Exist(3) Then
													If Fn_SISW_UI_Object_Operations("Fn_MyTcSrch_SpecifyQueryDetailsAndInvoke","Exist", JavaWindow("MyTeamcenter_Search").JavaWindow("SearchCriteriaTreeShell").JavaTree("Tree"),"")=False Then
														Fn_MyTcSrch_SpecifyQueryDetailsAndInvoke = True
													End If
												'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
												'++++++++++<< DropDown Button (CheckBox) >>++++++++
												Case "DsOwnGrpDrpDwn", "DsOwnUsrDrpDwn", "OwnUsrNameDrpDwn" , "DsTypDrpDwn", "ItmRevAliasTypDrpDwn", "ItmRevAltRevTypDrpDwn", "ItmRevOwningGrpDrpDwn", "ItmRevOwningUsrDrpDwn", "ItmRevTypDrpDwn", "OwningGrpDrpDwn", "OwningUsrDrpDwn", "GenOwningGrpDrpDwn", "GenOwningUsrDrpDwn", "DsKwSrchDtsTypeDrpDwn", "DsKwSrchOwnGrpDrpDwn", "DsKwSrchOwnUsrDrpDwn", "UsrID", "Checked-Out by User"
													
                                                                     ObjJavaWin.JavaStaticText("srch_Type").SetTOProperty "label", sObjElement
												 JavaWindow("MyTeamcenter").JavaButton("Clear").SetTOProperty "label","More...>>>"   ''Modified by vidya 15/06/2012

'																		  If JavaWindow("MyTeamcenter").JavaButton("Clear").Exist Then
																		  If Fn_SISW_UI_Object_Operations("Fn_MyTcSrch_SpecifyQueryDetailsAndInvoke","Exist", JavaWindow("MyTeamcenter").JavaButton("Clear"),"") Then
																					JavaWindow("MyTeamcenter").JavaButton("Clear").Click micLeftBtn 
																		 End If 
																	Select Case sObjElement
																		Case "DsOwnGrpDrpDwn", "ItmRevOwningGrpDrpDwn", "OwningGrpDrpDwn", "GenOwningGrpDrpDwn",  "DsKwSrchOwnGrpDrpDwn"
																			sAttachedText = "Owning Group:"
																		Case "DsOwnUsrDrpDwn", "ItmRevOwningUsrDrpDwn", "OwningUsrDrpDwn", "GenOwningUsrDrpDwn",  "DsKwSrchOwnUsrDrpDwn"
																			sAttachedText = "Owning User:"
																		Case "DsTypDrpDwn",  "DsKwSrchDtsTypeDrpDwn"
																			sAttachedText = "Dataset Type:"
																		Case "ItmRevAliasTypDrpDwn"
																			sAttachedText = "Alias Type:"
																		Case "ItmRevAltRevTypDrpDwn"
																			sAttachedText = "Alternate Revision Type:"
																		Case "ItmRevTypDrpDwn"
																			sAttachedText = "Type:"
																		Case "UsrID"
																			sAttachedText = "User ID:"
																		Case "Checked-Out by User"
																			sAttachedText = "Checked-Out by User:"
																		Case "OwnUsrNameDrpDwn"
																			sAttachedText = "Owning User Name:"
																																					
                     												End Select
																bFound = False
																wait 1
																ObjJavaWin.JavaStaticText("srch_Type").SetTOProperty "label", sAttachedText
'																Call Fn_Button_Click( "Fn_MyTcSrch_SpecifyQueryDetailsAndInvoke", ObjJavaWin,"srch_MultipleDropDown" )
																Call Fn_SISW_UI_JavaButton_Operations("Fn_MyTcSrch_SpecifyQueryDetailsAndInvoke", "Click", ObjJavaWin, "srch_MultipleDropDown")
'																wait 10
																wait 3
																Set objSelectType=description.Create()
																objSelectType("Class Name").value = "JavaStaticText"					
																Set  objIntNoOfObjects = ObjJavaWin.ChildObjects(objSelectType)
																For  innerCntr = 0 to objIntNoOfObjects.count-1
																	   If objIntNoOfObjects(innerCntr).getROProperty("label") = DictItems(iCounter) Then
																			objIntNoOfObjects(innerCntr).Click 2,2
																			bFound = TRUE
																			Exit for
																	   End If
																Next
																If  bFound = True Then
																	Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Value for ["+DictKeys(iCounter)+"] Successfully set.  ")   	
																	Fn_MyTcSrch_SpecifyQueryDetailsAndInvoke = True
																End If
										End select 
								End If
					End If                              
			Next
		    Call Fn_ToolbatButtonClick("Executes the search and displays the results in search result view")
			Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Pass : MyTeamCenter Window does exist.") 
			Fn_MyTcSrch_SpecifyQueryDetailsAndInvoke = True 
		Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Fail : MyTeamCenter Window does not exist.") 
			Fn_MyTcSrch_SpecifyQueryDetailsAndInvoke = False
		End If 
Set objElementEdit = Nothing
End Function 
'####### E.O.F~. =  Fn_MyTcSrch_SpecifyQueryDetailsAndInvoke	###################################################################################################~
'##################################################################################################################################################################~

'###################################################  Import a existing Query in Query Builder Application   ######################################################~
'#
'# FUNCTION NAME:	 	  Fn_QryBldr_ImportQuery
'#
'#  MyCommunity ID :	297
'#
'# MODULE: 						 Search Requirement, MyTC 
'#
'# PRE-REQUISITE:		RAC Session accessible and [Query Builder] Application loaded
'#
'# DESCRIPTION:			  Import a existing Query in Query Builder Application
'#										
'# PARAMETERS   :        sFilePath: Source path of Import file format
'#
'# RETURN VALUE : 		TRUE \ FALSE
'#
'#Examples	:		Fn_QryBldr_ImportQuery("C:\mainline\Reports\SrcSrchPrefPFFCreate.png")  		
'#										
'#	History	:						Developer Name			Date			Version				Changes Done			Reviewer	
'#------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'#										Kavan Shah~					10-06-10		1.0																	Sunil Rai
'#
'#										Sagar Shivade				6-Jan11			9.0 					Comment out error code
'#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'####################################################################################################################################################################~
'###################################################  Import a existing Query in Query Builder Application   ########################################################~

Public Function Fn_QryBldr_ImportQuery(sFilePath)  
GBL_FAILED_FUNCTION_NAME="Fn_QryBldr_ImportQuery"
Dim ObjQryApp
Set ObjQryApp =  Fn_UI_ObjectCreate( "Fn_QryBldr_ImportQuery",JavaWindow("Search_QueryBuilder").JavaWindow("QueryBuilder"))

		'++++++++++<<    Click [Import] button  >>++++++++++
		Call Fn_CheckBox_Select("Fn_QryBldr_ImportQuery", ObjQryApp, "Import")

		'++++++++++<<     Invoke [...] button on [Import] dialog  >>++++++++++
		Call Fn_Button_Click( "Fn_QryBldr_ImportQuery",ObjQryApp, "BrowseFilePath")

		'++++++++++<<      Specify \Path\FileName under [Read Query Definition] dialog  >>++++++++++
		Call Fn_Edit_Box("Fn_QryBldr_ImportQuery",JavaDialog("ImportPath"),"Filename",sFilePath)

		'++++++++++<<     Invoke [Import] button on [Read Query Definition] dialog  >>++++++++++
		Call Fn_Button_Click( "Fn_QryBldr_ImportQuery",JavaDialog("ImportPath") , "Import")

		'++++++++++<<     Checks if any error has occured.  >>++++++++++
'		If Err.Number > 0 Then
'
'			Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Fail: Failed to Import  the Requested Query. ")  
'			Fn_QryBldr_ImportQuery =False
'            Err.clear

'		Else
            
			'++++++++++<<     Invoke [Verify] button on [Import] dialog  >>++++++++++
			Call Fn_Button_Click( "Fn_QryBldr_ImportQuery",ObjQryApp , "Verify")
	
			'++++++++++<<    Invoke [OK] button on [Import] dialog   >>++++++++++
			Call Fn_Button_Click( "Fn_QryBldr_ImportQuery",ObjQryApp , "OK")
			Call Fn_ReadyStatusSync(2)
			Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Pass: Successfully Imported the Requested Query. ")   		
			Fn_QryBldr_ImportQuery = True

'		End If

		Call Fn_Button_Click( "Fn_QryBldr_ImportQuery", ObjQryApp, "Create")

Set ObjQryApp = Nothing
End Function
 
'####### 	E.O.F~. =  Fn_QryBldr_ImportQuery		################################################################################################################~
'###################################################################################################################################################################~

'###################################################################################################################################################################~
'###################################################  Export a existing Query in Query Builder Application    ######################################################~
'#
'# FUNCTION NAME:	 	  Fn_QryBldr_ExportQuery
'#
'#  MyCommunity ID :	296
'#
'# MODULE: 						 Search Requirement, MyTC 
'#
'# PRE-REQUISITE:		RAC Session accessible and [Query Builder] Application loaded
'#
'# DESCRIPTION:			 Export a existing Query in Query Builder Application
'#										
'# PARAMETERS   :        sQueryPath: Tree Path of existing query to be selected
'#										     sFilePath: Destination path of Export file format
'#
'# RETURN VALUE : 		TRUE \ FALSE
'#
'#Examples	:				
'#										
'#	History	:						Developer Name			Date			Version				Changes Done			Reviewer	
'#------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'#										Kavan Shah~					10-06-10		1.0																	Sunil Rai
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'#										Sandeep N					05-Sep-2012		1.1																	Sukhada B
'#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'###################################################################################################################################################################~
'###################################################  Export a existing Query in Query Builder Application    ######################################################~

Public Function Fn_QryBldr_ExportQuery(sQueryPath, sFilePath)  
GBL_FAILED_FUNCTION_NAME="Fn_QryBldr_ExportQuery"
Dim ObjQryApp
Set ObjQryApp =  Fn_UI_ObjectCreate( "Fn_QryBldr_ExportQuery",JavaWindow("Search_QueryBuilder").JavaWindow("QueryBuilder"))

	'++++++++++<<     Load the Query by selecting it from Query Builder Tree  >>++++++++++
	If True =  Fn_UI_JavaTree_Activate_Select_Node("Fn_QryBldr_ExportQuery", ObjQryApp,"AbstractQueryBuilderApplicatio", sQueryPath ) Then

		'++++++++++<<    Click [Export] button  >>++++++++++
		Call Fn_CheckBox_Select("Fn_QryBldr_ExportQuery", ObjQryApp, "Export")

		'++++++++++<<     Invoke [Save] button on [Print] dialog  >>++++++++++
		Call Fn_Button_Click( "Fn_QryBldr_ExportQuery",ObjQryApp.JavaDialog("ExportOption"), "Save")
		'Added code to handle Hierarchy : Dialog("SaveExport")
		If Dialog("SaveExport").Exist(4) Then
			wait 2
			Dialog("SaveExport").WinEdit("FileName").Set sFilePath
			wait 2
			Dialog("SaveExport").WinButton("Save").Click
		elseif ObjQryApp.JavaDialog("SaveExport").Exist(4) then
			'++++++++++<<     Specify \Path\FileName under [Save] dialog  >>++++++++++
			Call Fn_Edit_Box("Fn_QryBldr_ExportQuery",ObjQryApp.JavaDialog("SaveExport"),"FileName",sFilePath)
			'++++++++++<<     Invoke [Save] button on [Save] dialog  >>++++++++++
			Call Fn_Button_Click( "Fn_QryBldr_ExportQuery",ObjQryApp.JavaDialog("SaveExport"), "Save")
		else
			Fn_QryBldr_ExportQuery = False
			Set ObjQryApp = Nothing
			Exit function
		End If

		Call Fn_Button_Click( "Fn_QryBldr_ExportQuery",ObjQryApp.JavaDialog("ExportOption"), "Close")

		Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Pass: Successfully Exported the Requested Query. ")   		
		Fn_QryBldr_ExportQuery = True

	Else

		Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Fail: Requested Query Not Found. ")   		
		Fn_QryBldr_ExportQuery = False

	End If	

Set ObjQryApp = Nothing
End Function
 
'####### 	E.O.F~. =  Fn_QryBldr_ExportQuery		################################################################################################################~
'#########################################################################################################################################################~

'#######################################################################################################################################################~
'########################################################  Create a Local Query with defined inputs as specified     ################################################~
'#
'# FUNCTION NAME:	 	  Fn_Proj_CreateBasic 
'#
'# MODULE: 					Search/Project 
'#
'# PRE-REQUISITE:		RAC Session accessible and [Project] Application is to be loaded
'#
'# DESCRIPTION:			 Create Basic Project with defined inputs as specified 
'#											 0. Input [ID] field details
'#											 1. Input [Name] field details
'#											 2. Input [Description] field details
'#											 3. Set [Status] 
'#											 4. Set[Program Security] 
'#											 5. Select requested members [Array to be passed]
'#											 6. Select requested Administrator
'#										
'# PARAMETERS   :      iID:  Valid ID [Num]
'#										   sName: Valid Project Name [Alpha-Num]
'#										   sDesc: Valid Description 
'#									       sStatus: Valid Status Name []
'#										   sProgSecurity: Valid Program Security ID
'#											aMembers: All the Nodes from organization group and to put it in selected Members.  
'#											sProjAdmin: Admin name to select from "Select a Team Administrator " Dialog.
'#
'# RETURN VALUE : 		TRUE \ FALSE
'#
'#Examples	:				[ Create array  for Members and then pass Array as parameter. ]
'#										Arr = Array("Organization:Engineering:Designer:AutoTest2 (autotest2)","Organization:Engineering:Designer:AutoTest7 (autotest7)","Organization:Engineering:Designer:AutoTest6 (autotest6)", "Organization:Project Administration:Project Administrator:AutoTestDBA (autotestdba)" )
'#										MsgBox Fn_Proj_CreateBasic("123", "first", "Testing", "Inactive", "ON", Arr, "AutoTestDBA (autotestdba)")
'#										
'#	History	:						Developer Name			Date			Version				Changes Done			Reviewer	
'#------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'#										Kavan Shah~					July@2010		1.0															 	Sunil Rai

'#									sAGAR sHIVADE				Jan 32011		porting 9.0			changes object path for create project object Line 1174
'#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'########################################################################################################################################################~
'########################################################  Create Basic Project with defined inputs as specified     ################################################~

Public Function  Fn_Proj_CreateBasic(iID, sName, sDesc, sStatus, sProgSecurity, aMembers, sProjAdmin)
GBL_FAILED_FUNCTION_NAME="Fn_Proj_CreateBasic"
Dim iCntr, iabound, ArrPrevTmMem, iInnBound, aRefExpnd, sNodeEle,IExpndCntr
Dim ObjProjWndDef

'Set ObjProjWndDef =  Fn_UI_ObjectCreate( "Fn_Proj_CreateBasic ",JavaWindow("Project - Teamcenter 8").JavaApplet("ProjectApplet") )
 Set ObjProjWndDef =  Fn_UI_ObjectCreate( "Fn_Proj_CreateBasic ",Window("ProjectSearch").JavaWindow("Project_Srch") )
	Call Fn_Button_Click( "Fn_Proj_CreateBasic ", ObjProjWndDef, "Clear")

	'++++++++++<<    Set the ID   >>++++++++++
	If iID <> "" Then
		Call Fn_Edit_Box("Fn_Proj_CreateBasic ",ObjProjWndDef,"ID",iID)
	End If
    '++++++++++<<    Set the Name >>++++++++++
	If sName <> "" Then
		Call Fn_Edit_Box("Fn_Proj_CreateBasic ",ObjProjWndDef,"Name",sName)
	End If
	'++++++++++<<    Set the Description >>++++++++++
	If sDesc <> "" Then
		Call Fn_Edit_Box("Fn_Proj_CreateBasic ",ObjProjWndDef,"Desc",sDesc)
	End If
	'++++++++++<<    Set the Status  >>++++++++++
	If sStatus <> "" Then
		Call Fn_UI_Object_SetTOProperty("Fn_Proj_CreateBasic",ObjProjWndDef.JavaRadioButton("StatusActive") ,"attached text",sStatus)
		Call  Fn_UI_JavaRadioButton_SetON("Fn_Proj_CreateBasic ",ObjProjWndDef, "StatusActive")
	End If
	'++++++++++<<    Set the Program Security  >>++++++++++
	If sProgSecurity <> "" Then
		Call Fn_CheckBox_Set("Fn_Proj_CreateBasic ", ObjProjWndDef, "UsePrgmSec", sProgSecurity)
	End If
	'++++++++++<<    Selecting the members  >>++++++++++
	
	If UBound(aMembers) <> 0 Then
		iabound = UBound(aMembers)
		For iCntr = 0 to iabound
			aRefExpnd = split(aMembers(iCntr), ":", -1, 1) 
			sNodeEle = aRefExpnd(0)
			If  UBound(aRefExpnd) > 0 Then
				For IExpndCntr = 1 to UBound(aRefExpnd)
					sNodeEle = sNodeEle+":"+aRefExpnd(IExpndCntr)
					Wait(2)
					Call Fn_UI_JavaTree_Expand("Fn_Proj_CreateBasic", ObjProjWndDef, "OrgGrpTree",sNodeEle )
					Wait(2)
				Next
			End If
            Call Fn_JavaTree_Node_Activate("Fn_Proj_CreateBasic",ObjProjWndDef,"OrgGrpTree",aMembers(iCntr) )
		Next

		Call Fn_Button_Click( "Fn_Proj_CreateBasic ", ObjProjWndDef, "PrjctTeam")
		If True = Fn_UI_ObjectExist("Fn_Proj_CreateBasic", ObjProjWndDef.JavaDialog("SelectPrvlgdTeam"))  Then
			Call Fn_Button_Click( "Fn_Proj_CreateBasic ",ObjProjWndDef.JavaDialog("SelectPrvlgdTeam") , "MultipleOut")
			Call Fn_Button_Click( "Fn_Proj_CreateBasic ",ObjProjWndDef.JavaDialog("SelectPrvlgdTeam") , "OK")
		End If
	End If

	'++++++++++<<    Selecting the project admin  >>++++++++++
	If sProjAdmin <> "" Then

		Call Fn_Button_Click( "Fn_Proj_CreateBasic ", ObjProjWndDef, "PrivilegedMem")

		If True = Fn_UI_ObjectExist("Fn_Proj_CreateBasic", ObjProjWndDef.JavaDialog("TeamAdministrator"))  Then
			ArrPrevTmMem = split(sProjAdmin, ":", -1,1)
			iInnBound = UBound(ArrPrevTmMem)
			Call Fn_UI_JavaList_ExtendSelect("Fn_Proj_CreateBasic", ObjProjWndDef.JavaDialog("TeamAdministrator"), "TeamAdminList", ArrPrevTmMem(iInnBound) )	
			Call Fn_Button_Click( "Fn_Proj_CreateBasic ",ObjProjWndDef.JavaDialog("TeamAdministrator") , "OK")
		End If
    End If

    '++++++++++<<  Invoke [Create] button >>++++++++++
	If  Fn_UI_Object_GetROProperty("Fn_Proj_CreateBasic",ObjProjWndDef, "enabled") = "1" Then
		Call Fn_Button_Click( "Fn_Proj_CreateBasic ", ObjProjWndDef, "Create")
		Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Pass: Project Successfully Created. ")   	
		Fn_Proj_CreateBasic  = True
	Else
		Call Fn_Button_Click( "Fn_Proj_CreateBasic ", ObjProjWndDef, "Clear")
		Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Fail: Project Not Created due to wrong parameters. ")   	
		Fn_Proj_CreateBasic  = False
	End If
	
	Set ObjProjWndDef = Nothing

End Function
 
'####### 	E.O.F~. =  Fn_Proj_CreateBasic 		##################################################################################################################~

'##########################################################################################################################################################~
'# FUNCTION NAME:	 	Fn_MyTc_SearchHistoryOperation
'#
'#
'# MODULE: 			Search Requirement
'#				
'# DESCRIPTION:			This Operate on Search History
'#							
'# PARAMETERS   :     	 	
'#														  
'# RETURN VALUE : 	   	TRUE \ FALSE
'#
'# Examples	:		Fn_MyTc_SearchHistoryOperation()	
'#										
'# History	:		Developer Name			Date					Rev. No.			Changes Done			Reviewer	
'#------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'#							Mallikarjun						05-Aug-2010		001
'#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'###########################################################################################################################################################~


Public Function Fn_MyTc_SearchHistoryOperation(strAction, sSearchHistoryItems)
GBL_FAILED_FUNCTION_NAME="Fn_MyTc_SearchHistoryOperation"
Dim ObjPrefDialog, ObjJavaWin, ObjJavaWinPrevSer, objDetailsTable
Dim sText, bReturn, oCounter, iCounter, aObjList, intItemCount, bFlag

				'1. Operate on the Main Menu: Window;Show View;Other...
					Call Fn_MenuOperation("Select","Window:Show View:Other...")

				'2. Select Tree Node [Others;Search Results] under [Show View] dialog and click [OK] button
					Set ObjJavaWin = Fn_UI_ObjectCreate("Fn_MyTc_SearchHistoryOperation",JavaWindow("DefaultWindow").JavaWindow("Show View") )
					ObjJavaWin.JavaTree("ViewTree").Activate "Other:Search Results"
'					Call  Fn_Button_Click("Fn_MyTc_SearchHistoryOperation", ObjJavaWin, "OK") 

				'3. After the search results Panel is loaded, [Double Click] on Toolbar Icon [Show Previous Searches]. This should invoke [Previous Search Results] Dialog
					Call Fn_ToolbatButtonClick("Show Previous Searches")



	Select Case StrAction


		'----------------------------------------------------------------------- For Clearing Stale Searches -------------------------------------------------------------------------
		Case "Clear Stale Search History"
 
				
				'4. Select all the listed Search results under the dialog, if they exists, and click on [Delete] button. Click [Cancel] to close the dialog
					Set ObjJavaWinPrevSer = Fn_UI_ObjectCreate("Fn_MyTc_SearchHistoryOperation",JavaWindow("MyTeamcenter").JavaWindow("PreviousSrchRslts") )
					If 0 = Fn_UI_Object_GetROProperty("Fn_MyTc_SearchHistoryOperation",ObjJavaWinPrevSer.JavaTable("Select the search result"),"rows") Then
						Call  Fn_Button_Click("Fn_MyTc_SearchHistoryOperation", ObjJavaWinPrevSer, "Cancel")  
					Else
						ObjJavaWinPrevSer.JavaTable("Select the search result").PressKey "a",micCtrl
						Call  Fn_Button_Click("Fn_MyTc_SearchHistoryOperation", ObjJavaWinPrevSer, "Delete")  
						Call  Fn_Button_Click("Fn_MyTc_SearchHistoryOperation", ObjJavaWinPrevSer, "Cancel")  
					End If
				'5. Operate on the Main Menu: Window;Preferences to invoke [Preferences] dialog
					Call Fn_MenuOperation("Select","Window:Preferences")

				'6. Select the Tree Node [Search; Results] to load the [Results] panel
					Set ObjPrefDialog = Fn_UI_ObjectCreate("Fn_MyTc_SearchHistoryOperation",JavaWindow("DefaultWindow").JavaWindow("Preferences") )
					Call Fn_UI_JavaTree_Expand("Fn_MyTc_SearchHistoryOperation", ObjPrefDialog, "Tree","Search")
					ObjPrefDialog.JavaTree("Tree").Select "Search:Results"
			
				'7. Clcik on [Clear Search History] button and then click [Apply] button
					Call Fn_Button_Click("Fn_MyTc_SearchHistoryOperation", ObjPrefDialog, "ClearSearchHistory") 
					Call  Fn_Button_Click("Fn_MyTc_SearchHistoryOperation", ObjPrefDialog, "Apply") 
		
				'8. Click on [OK] button to close the [Preferences] dialog
					Call  Fn_Button_Click("Fn_MyTc_SearchHistoryOperation", ObjPrefDialog, "OK") 

				' Writing Log
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Performed all the Steps for function Fn_MyTc_SearchHistoryOperation.")
					Fn_MyTc_SearchHistoryOperation = True


		'----------------------------------------------------------------------- For Validating Entries in Search History -------------------------------------------------------------------------
		Case "Validate Search History"


				'4. Validate all the listed Search results under the dialog, if NULL exists, Click [Cancel] to close the dialog
					Set ObjJavaWinPrevSer = Fn_UI_ObjectCreate("Fn_MyTc_SearchHistoryOperation",JavaWindow("MyTeamcenter").JavaWindow("PreviousSrchRslts") )
					Set objDetailsTable = Fn_UI_ObjectCreate("Fn_MyTc_SearchHistoryOperation",JavaWindow("MyTeamcenter").JavaWindow("PreviousSrchRslts").JavaTable("Select the search result") )
					
					aObjList = Split(sSearchHistoryItems,":")
					intItemCount =ubound(aObjList)
					bFlag = false
					
					If 0 = Fn_UI_Object_GetROProperty("Fn_MyTc_SearchHistoryOperation",ObjJavaWinPrevSer.JavaTable("Select the search result"),"rows") Then
						'Call  Fn_Button_Click("Fn_MyTc_SearchHistoryOperation", ObjJavaWinPrevSer, "Cancel") 
						If aObjList(0) = "NULL" Then
							Fn_MyTc_SearchHistoryOperation = True
							bFlag = true
						End If
						'Call  Fn_Button_Click("Fn_MyTc_SearchHistoryOperation", ObjJavaWinPrevSer, "Cancel")

					Else
 						bReturn = objDetailsTable.GetROProperty("rows")
						For oCounter=0 to intItemCount
							For iCounter=0 to bReturn-1
								sText = objDetailsTable.GetCellData(iCounter,0)						
								If IsNumeric(aObjList(oCounter)) Then
							 		If cstr(sText) = cstr(cint(aObjList(oCounter)))  Then
								 		bFlag = true
								 		Exit for
									End If
								elseIf cstr(sText) = cstr(aObjList(oCounter))  Then
								 	bFlag = true
								 	Exit for
								End If
							Next
							
							'If Item Not Found:
							If bFlag = false Then
								Fn_MyTc_SearchHistoryOperation = False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Find Item [" + aObjList(oCounter) + "] under Search History Dialog")
								Exit for
							End If									
						Next

					End If
					
					If bFlag = true Then
						Fn_MyTc_SearchHistoryOperation = True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Validated Item [" + sSearchHistoryItems + "] under Search History Dialog")
					End If						
					
					'Dismiss the Search History Dialog:
					Call  Fn_Button_Click("Fn_MyTc_SearchHistoryOperation", ObjJavaWinPrevSer, "Cancel") 

 
		Case Else
						Fn_MyTc_SearchHistoryOperation = False
						'Report Invalid Action Arguement Error
						'Call Fn_WriteLogFile("Fn_MyTc_SearchHistoryOperation", 1, Err.Number,"FAIL: Wrong Action Parameter [" + StrAction + "] for [Fn_MyTc_SearchHistoryOperation] Function Call")
	End Select


Set ObjJavaWinPrevSer = Nothing
Set ObjJavaWin = Nothing
Set ObjPrefDialog = Nothing
Set objDetailsTable = Nothing
End Function

'########################################################  Modify a Local Query with defined inputs as specified     ################################################~
'#
'# FUNCTION NAME:	 	  Fn_QryBldr_CrtLocQryFrmExisting
'#
'#  MyCommunity ID :	
'#
'# MODULE: 						 Search Requirement, MyTC 
'#
'# PRE-REQUISITE:		RAC Session accessible and [Query Builder] Application loaded & Query to be modified should be loaded.
'#
'# DESCRIPTION:			 Modify a Local Query with defined inputs as specified
'#											 1. Input [Name] field details
'#											 2. Input [Description] field details
'#											 3. Click [Search Class] button
'#														 3a. Tree navigate to trace [Class Attribute], For Example: POM_object:WorkspaceObject:ItemRevision
'#														 3b. Select/Double Click on Tree Node of prefered [Class Attribute]
'#														 3b. Close [Class Attribute] window
'#											 4. Check the [Search Class] button label updated to requiste class
'#											 5. Set the [Display Setting] to required option [Class/All Attributes]
'#											 6. Set [Show Indented Results] option [On/Off]
'#											 7. Select attribute from [Attribute Selection] Tree
'#													>>	 Note: 	<<
'#													 7.1 Please note that the function arguement is a array of required attributes, seperated by ":"
'#													 7.2 Select the required attribute iteratively and click on [+] button to add the attribute
'#											 8. Invoke [Create] button
'#										
'# PARAMETERS   :      sQueryName: Name of the Local User Query
'#										   sQueryDescription: Description of the Local User Query
'#										   sSearchClass: Attribute Class POM Object of the Query
'#									       sDisplaySettings: Display Settings [Class/All Attributes]
'#										   bShowIndentedResults: [Show Indented Results] option [On/Off]
'#
'#									      aAttributes: Array of the class attributes to be added to the Search Query
'#													 >>   Note: ( 1.) Multiple Attributes to be seperated by "~" ( Tilde)  [ EXAMPLE - >> Attrib1~Attrib2~Attrib3]
'#  																   ( 2.) Inside Attribute  inner values to be seperated by "," (Comma)  [ EXAMPLE - >> First, Second, Third]
'#  																  ( 3.) Inside Values Reference Path  to be seperated by ":" (Colon)  [ EXAMPLE - >> Dataset:Revision
'#
'#				             	     								>>	InnerValues  = "First, Second, Third"
'#																						 First =  refpath for Attrib in main window - will activate /double click
'#				  				   								 				Second =  [For class Attrib Sel Dialog ] ->> FullRefPath Attrib to be selected.				 (Set the Class)   				
'#																									[For Class Selection Dialog} ->>EditBoxValue]
'#																					Third = Refpath for Attrib in main window - will activate /double click
'#
'#												->>	 (Multiple Attribute) Example>>  (First, Second, Third ~ First, Second, Third~ First, Second, Third)  << -
'#												->>	 (First, Second, Third ) Example>>  Refpath1,RefPath2,Refpath3<< -
'#												->>	 (Refpath1 ) Example>>  Home:Child << -
'#										sRemAttribNmIndex: Index Number for row to be deleted.
'#
'# RETURN VALUE : 		TRUE \ FALSE
'#
'#Examples	:				 Fn_QryBldr_ModifyLocalQuery("LocQuery1", "Local Test Query", "Dataset", "AllAttributes", "OFF", "Dataset:Referenced By,ItemRevision:IMAN_specification,Dataset:Specifications [ ItemRevision ]:Revision~Dataset:pid" , "1") 
'#										Fn_QryBldr_ModifyLocalQuery("LocQuery1", "Local Test Query", "Dataset", "AllAttributes:RealNames", "ON", "Dataset:pid" , "") 
'#
'#	History	:						Developer Name			Date			Version				Changes Done			Reviewer	
'#------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'#										Deepak		03-08-10		1.0																	Sunil Rai
'#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'########################################################################################################################################################~
'########################################################  Modify a Local Query with defined inputs as specified     ################################################~


Public Function  Fn_QryBldr_CrtLocQryFrmExisting(sQueryName, sQueryDescription, sSearchClass, sDisplaySettings, bShowIndentedResults, aAttributes, sRemAttribNmIndex)  
  GBL_FAILED_FUNCTION_NAME="Fn_QryBldr_CrtLocQryFrmExisting"
Dim  ArrDispSet,  OuterArrAttrib, iOuterCounter, ArrAttrib,  ArrinnAttrib, iCounter
Dim ObjQryApp, ObjQryAttribSel,bReturn,RmvArrAttrib,iCnt,iRows,iCot,iDelCounter

Set ObjQryApp =  Fn_UI_ObjectCreate( "Fn_QryBldr_CrtLocQryFrmExisting",JavaWindow("Search_QueryBuilder").JavaWindow("QueryBuilder"))

		
				'++++++++++<<    Input [Name] field details >>++++++++++
				If sQueryName <> "" Then
					Call Fn_Edit_Box("Fn_QryBldr_CrtLocQryFrmExisting",ObjQryApp,"Name",sQueryName)
				End If
			
				'++++++++++<<   Input [Description] field details>>++++++++++
				If sQueryDescription <>"" Then
					Call Fn_Edit_Box("Fn_QryBldr_CrtLocQryFrmExisting",ObjQryApp,"Description",sQueryDescription)
				End If
			
				'++++++++++<<   Click [Search Class] button >>++++++++++
				If  sSearchClass <> "" Then
					Call Fn_CheckBox_Set("Fn_QryBldr_CrtLocQryFrmExisting", ObjQryApp, "SrchClass", "ON")
					Call Fn_Edit_Box("Fn_QryBldr_CrtLocQryFrmExisting",ObjQryApp,"Class/Attribute Selection",sSearchClass)
					Call Fn_Button_Click( "Fn_QryBldr_CrtLocQryFrmExisting", ObjQryApp, "Find")
					ObjQryApp.JavaObject("Close").Click 1,1
				End If
			
				 '++++++++++<<  Set the [Display Setting] to required option [Class/All Attributes] >>++++++++++
				 If sDisplaySettings <> ""  Then
					 ArrDispSet = split(sDisplaySettings, ":", -1,1)
					 If Ubound(ArrDispSet) = 1 Then
						Call Fn_CheckBox_Set("Fn_QryBldr_CrtLocQryFrmExisting", ObjQryApp, "DisplaySettings", "ON")
						Call  Fn_UI_JavaRadioButton_SetON("Fn_QryBldr_CrtLocQryFrmExisting",ObjQryApp, ArrDispSet(0))
						Call  Fn_UI_JavaRadioButton_SetON("Fn_QryBldr_CrtLocQryFrmExisting",ObjQryApp, ArrDispSet(1))
					Else
						Call Fn_CheckBox_Set("Fn_QryBldr_CrtLocQryFrmExisting", ObjQryApp, "DisplaySettings", "ON")
						Call  Fn_UI_JavaRadioButton_SetON("Fn_QryBldr_CrtLocQryFrmExisting",ObjQryApp, sDisplaySettings )
					 End If
					ObjQryApp.JavaObject("Close").Click 1,1
				 End If
			
				 '++++++++++<<  Set [Show Indented Results] option [On/Off] >>++++++++++
				 If bShowIndentedResults <> ""  Then
					Call Fn_CheckBox_Set("Fn_QryBldr_CrtLocQryFrmExisting", ObjQryApp, "ShowIndentedResults", bShowIndentedResults)
				 End If
			
				'++++++++++<<   Select attribute from [Attribute Selection] Tree >>++++++++++
				If  aAttributes <> "" Then
					OuterArrAttrib = split(aAttributes, "~", -1,1)
					For iOuterCounter = 0 To Ubound(OuterArrAttrib)
							ArrAttrib = split(OuterArrAttrib(iOuterCounter), ",", -1, 1)				
							For iCounter = 0 to Ubound(ArrAttrib) 		
										ArrinnAttrib = split(ArrAttrib(iCounter), ":", -1, 1)
										Select Case iCounter
										Case "0"
												ObjQryApp.JavaTree("AttributeSelectionList").Activate ArrAttrib(iCounter) 
										Case "1"
												If  True = Fn_UI_ObjectExist("Fn_QryBldr_CrtLocQryFrmExisting", ObjQryApp.JavaDialog("ClassAttributeSelection") )Then
												
													Set ObjQryAttribSel = ObjQryApp.JavaDialog("ClassAttributeSelection") 
													Call Fn_CheckBox_Set("Fn_QryBldr_CrtLocQryFrmExisting", ObjQryAttribSel, "CAS_SrchClass", "ON")
													Call Fn_Edit_Box("Fn_QryBldr_CrtLocQryFrmExisting",ObjQryAttribSel,"CAS_Edit",ArrinnAttrib(0) )
													Call Fn_Button_Click( "Fn_QryBldr_CrtLocQryFrmExisting", ObjQryApp, "Find")
													ObjQryApp.JavaObject("Close").Click 1,1
													ObjQryAttribSel.JavaTree("CAS_SrchTree").Activate ArrAttrib(iCounter) 
												End If
												If  True = Fn_UI_ObjectExist("Fn_QryBldr_CrtLocQryFrmExisting",ObjQryApp.JavaDialog("ClassSelectionDialog")  )Then
													Set ObjQryAttribSel = ObjQryApp.JavaDialog("ClassSelectionDialog")  
													Call Fn_Edit_Box("Fn_QryBldr_CrtLocQryFrmExisting",ObjQryAttribSel,"SelectionField",ArrAttrib(iCounter) )
													Call Fn_Button_Click( "Fn_QryBldr_CrtLocQryFrmExisting", ObjQryAttribSel, "ClassAttSelFind")
													Call Fn_Button_Click( "Fn_QryBldr_CrtLocQryFrmExisting", ObjQryAttribSel, "CSDOK")
												End If										               
										Case "2"
												ObjQryApp.JavaTree("AttributeSelectionList").Activate ArrAttrib(iCounter) 
										End Select		
							Next			
					Next
				End If
			
	   '++++++++++<<   Remove the Attribute Specified. >>++++++++++
	'		Dim bReturn,RmvArrAttrib,iCnt,iRows,iCot,iDelCounter
				If sRemAttribNmIndex <> "" Then
					RmvArrAttrib = split(sRemAttribNmIndex, ":", -1,1)
					iRows = ObjQryApp.JavaTable("SrchCriteriaTable").GetROProperty("rows")
					iCot = 0
					bReturn = False
					iDelCounter = 0
					For iCnt = 0 to iRows - 1
						If  iCnt = cInt ( RmvArrAttrib(iCot)) Then
							ObjQryApp.JavaTable("SrchCriteriaTable").SelectRow(RmvArrAttrib(iCot)-iDelCounter)
							Call Fn_Button_Click( "Fn_QryBldr_CrtLocQryFrmExisting", ObjQryApp, "Remove")	
							iDelCounter = iDelCounter +1
							iCot = iCot +1
							Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Fail : Local Query Search Criteria Row no [" + Cstr(iCnt) + "] Not Deleted") 
							bReturn = True
								If  iCot > Ubound (RmvArrAttrib) Then
									Exit For
								End If
					  End If
					Next
					
					
					 If bReturn = True	Then
						 Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Fail : Local Query Search Criteria Row Not Deleted")   	
						 Fn_QryBldr_CrtLocQryFrmExisting = False
					Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Fail : Local Query Search Criteria  Deleted")   	
							Fn_QryBldr_CrtLocQryFrmExisting = False
							Exit Function
					End If
				End If 
				'++++++++++<<  Invoke [Create] button >>++++++++++
				 If True =  Fn_Button_Click( "Fn_QryBldr_CrtLocQryFrmExisting", ObjQryApp, "Create") Then
                        Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Pass: Local Query Successfully Modified with new name. ")   	
						Fn_QryBldr_CrtLocQryFrmExisting = True
				Else
						Call Fn_Button_Click( "Fn_QryBldr_CrtLocQryFrmExisting", ObjQryApp, "Clear")
						Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Fail: Local Query Failed to be Modified with new name. ")   	
						Fn_QryBldr_CrtLocQryFrmExisting = False
				End If
		
	Set ObjQryApp = Nothing
	Set ObjQryAttribSel = Nothing
	End Function
	
'####### 	E.O.F~. =  Fn_QryBldr_CrtLocQryFrmExisting		########################################################################################################~

'##########################################################################################################################################################~
'# FUNCTION NAME:	 	Fn_MyTc_Search_SortTableOperations
'#
'#
'# MODULE: 			Search Requirement
'#				
'# DESCRIPTION:			This Operate on Search Sort
'#							
'# PARAMETERS   :     	 	
'#														  
'# RETURN VALUE : 	   	TRUE \ FALSE
'#
'# Examples	:		Fn_MyTc_Search_SortTableOperations()	
'#										
'# History	:		Developer Name			Date			Rev. No.			Changes Done			Reviewer	
'#------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'#				Mallikarjun			09-Aug-2010		001		1.0
'#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'#				Koustubh			07-Jun-2012		001		1.1					modified case SortItemCellUpdate, selected expectedvalue descriptively
'#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'#				Manish				08-Jun-2012		001		1.1					modified case SortItemValidate
'#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_MyTc_Search_SortTableOperations(sAction, sObjectName, sPropertyName, sExpectedValue)
	GBL_FAILED_FUNCTION_NAME="Fn_MyTc_Search_SortTableOperations"
	Dim ObjJavaSortWindow, objSortTable,bReturn,iCounter,aObjList,intItemCount,oCounter, rowIndex, bFlag, aMenuList, intCount, sMenu, sText, aMenuList1()
	Dim objSelectType, intNoOfObjects, i

	'Set False Flag to the Function: 
	Fn_MyTc_Search_SortTableOperations = False


	'1. Operate on the Main Menu: Window;Show View;Other...
	Call Fn_MenuOperation("Select","Window:Show View:Search")

	'2. After the search results Panel is loaded, [Double Click] on Toolbar Icon [Show Previous Searches]. This should invoke [Previous Search Results] Dialog
	Call Fn_ToolbatButtonClick("Sort")
	wait 4

	Set ObjJavaSortWindow = Fn_UI_ObjectCreate("Fn_MyTc_SearchHistoryOperation",JavaWindow("MyTeamcenter").JavaWindow("Search_Sort") )
	If NOT ObjJavaSortWindow.GetROProperty("Maximized") Then
		ObjJavaSortWindow.maximize
	End If
	
	Set objSortTable = JavaWindow("MyTeamcenter").JavaWindow("Search_Sort").JavaTable("Table")

	Select Case sAction

		 Case "SortItemExist"
   				bFlag = false
				'Count number of rows of Table
				bReturn = objSortTable.GetROProperty("rows")	
				'Extract the index of row at which the object exist.
				For iCounter=0 to bReturn - 1
					sText = objSortTable.GetCellData(iCounter,"Sort By")						
						If IsNumeric(sObjectName) Then
							 If cstr(sText) = cstr(cint(sObjectName))  Then
								 bFlag = true
								 Exit for
							End If
						elseIf cstr(sText) = cstr(sObjectName)  Then
								 bFlag = true
								 Exit for
						End If									
				Next
				If bFlag = false Then
						Fn_MyTc_Search_SortTableOperations = FALSE				 
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Fn_MyTc_Search_SortTableOperations : Row with Object "+ sObjectName +" does not exist")	
						Exit function
				Else 
						Fn_MyTc_Search_SortTableOperations = TRUE				 
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Fn_MyTc_Search_SortTableOperations: Row with Object "+ sObjectName +" exist")	
				End If
				
		 Case "SortItemGetIndex"
   				bFlag = -1
				'Count number of rows of Table
				bReturn = objSortTable.GetROProperty("rows")	
				'Extract the index of row at which the object exist.
				For iCounter=0 to bReturn - 1
					sText = objSortTable.GetCellData(iCounter,"Sort By")						
						If IsNumeric(sObjectName) Then
							 If cstr(sText) = cstr(cint(sObjectName))  Then
								 bFlag = iCounter
								 Exit for
							End If
						elseIf cstr(sText) = cstr(sObjectName)  Then
								 bFlag = iCounter
								 Exit for
						End If									
				Next
				If bFlag = -1 Then
						Fn_MyTc_Search_SortTableOperations = -1
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Fn_MyTc_Search_SortTableOperations : Failed to Get Index Of Object ["+ sObjectName +"]")	
						Call  Fn_Button_Click("Fn_MyTc_Search_SortTableOperations", ObjJavaSortWindow, "Cancel")
						Set ObjJavaSortWindow = Nothing
						Set objSortTable = Nothing						
						Exit function
				Else 
						Fn_MyTc_Search_SortTableOperations = iCounter + 1				 
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Fn_MyTc_Search_SortTableOperations: Successfully Got Index of Object ["+ sObjectName +"] :: ["+cstr( iCounter) +"]")	
						Call  Fn_Button_Click("Fn_MyTc_Search_SortTableOperations", ObjJavaSortWindow, "Cancel") 
						Set ObjJavaSortWindow = Nothing
						Set objSortTable = Nothing
						Exit function
				End If				

		'Case "SortItemSelect"
				'Count number of rows of Table
		'		bReturn = objSortTable.GetROProperty("rows")	
				'Extract the index of row at which the object exist.
		'		For iCounter=0 to bReturn - 1
		'		sText = objSortTable.GetCellData(iCounter,"Sort By")						
		'			If IsNumeric(sObjectName) Then
		'				 If cstr(sText) = cstr(cint(sObjectName))  Then
		'					 objSortTable.ClickCell iCounter,"Sort By","LEFT"
		'					 Exit for
		'				End If
		'			elseIf cstr(sText) = cstr(sObjectName)  Then
		'					 objSortTable.ClickCell iCounter,"Sort By","LEFT"
		'					 Exit for
		'			End If									
		'		Next

		 Case "SortItemMultiSelect"
				'Split the string where " : " exist
				aObjList = Split(sObjectName,":")
				intItemCount =ubound(aObjList)
				'Count number of rows of Table
				bReturn = objSortTable.GetROProperty("rows")	
				'Extract the index of row at which the object exist.
				For oCounter=0 to intItemCount
						For iCounter=0 to bReturn-1
						sText = objSortTable.GetCellData(iCounter,"Sort By")						
						If IsNumeric(aObjList(oCounter)) Then
							 If cstr(sText) = cstr(cint(aObjList(oCounter)))  Then
								 objSortTable.ClickCell iCounter, "Sort By","LEFT","CONTROL"
								 Exit for
							End If
						elseIf cstr(sText) = cstr(aObjList(oCounter))  Then
								 objSortTable.ClickCell iCounter, "Sort By","LEFT","CONTROL"
								 Exit for
						End If									
						Next
				Next
				
		Case "SortItemValidate"
				'Count number of rows of Table
				Fn_MyTc_Search_SortTableOperations = False
				'Count number of rows of Table
				bReturn = objSortTable.GetROProperty("rows")	
				'Extract the index of row of which relation is to be changed
				For iCounter=0 to bReturn - 1
					    sText = objSortTable.GetCellData(iCounter,"Sort By")						
						If IsNumeric(sObjectName) Then
							 If cstr(sText) = cstr(cint(sObjectName))  Then
								 Exit for
							End If
						elseIf cstr(sText) = cstr(sObjectName)  Then
								 Exit for
						End If								
				Next

				If iCounter < bReturn Then
					objSortTable.ActivateCell iCounter,"Order By"
					wait 1
					Set objSelectType = description.Create()
					objSelectType("Class Name").value = "JavaList"
					objSelectType("path").value = "List;Shell;Shell;Shell;"
					Set  intNoOfObjects = ObjJavaSortWindow.ChildObjects(objSelectType)
					  For i = 0 to intNoOfObjects.count-1
						  If intNoOfObjects(i).Exist(3) Then
							   intNoOfObjects(i).select sExpectedValue
							   Fn_MyTc_Search_SortTableOperations = True
								Exit for
						  End If
					  Next
						Set intNoOfObjects = Nothing
						Set objSelectType = Nothing
				End If

   		 Case "SortItemCellUpdate"
				Fn_MyTc_Search_SortTableOperations = False
				'Count number of rows of Table
				bReturn = objSortTable.GetROProperty("rows")	
				'Extract the index of row of which relation is to be changed
				For iCounter=0 to bReturn - 1
					    sText = objSortTable.GetCellData(iCounter,"Sort By")						
						If IsNumeric(sObjectName) Then
							 If cstr(sText) = cstr(cint(sObjectName))  Then
								 Exit for
							End If
						elseIf cstr(sText) = cstr(sObjectName)  Then
								 Exit for
						End If								
				Next

				If iCounter < bReturn Then
					objSortTable.ActivateCell iCounter,"Order By"
					wait 1
					Set objSelectType = description.Create()
					objSelectType("Class Name").value = "JavaList"
					objSelectType("path").value = "List;Shell;Shell;Shell;"
					Set  intNoOfObjects = ObjJavaSortWindow.ChildObjects(objSelectType)
					  For i = 0 to intNoOfObjects.count-1
						  If intNoOfObjects(i).Exist(3) Then
							   intNoOfObjects(i).select sExpectedValue
							   Fn_MyTc_Search_SortTableOperations = True
								Exit for
						  End If
					  Next
						Set intNoOfObjects = Nothing
						Set objSelectType = Nothing
				End If

																				
		Case Else
						Fn_MyTc_Search_SortTableOperations = False
						'Report Invalid Action Arguement Error
						'Call Fn_WriteLogFile("Fn_MyTc_Search_SortTableOperations", 1, Err.Number,"FAIL: Wrong Action Parameter [" + StrAction + "] for [Fn_MyTc_Search_SortTableOperations] Function Call")
	End Select
	
	'Close the Sort Dialog:
	Call  Fn_Button_Click("Fn_MyTc_Search_SortTableOperations", ObjJavaSortWindow, "OK") 	
	
	
						
	Fn_MyTc_Search_SortTableOperations = TRUE				 
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_MyTc_Search_SortTableOperations passed with case ["+ sAction +"] on Object [" + sObjectName +"]")	

	'Release Objects:
	Set ObjJavaSortWindow = Nothing
	Set objSortTable = Nothing 

End Function


'##########################################################################################################################################################~
'# FUNCTION NAME:	 	Fn_MyTc_ChangeSearchTreeOperation
'#
'#
'# MODULE: 			Search Requirement
'# 
'# TEST REQUIREMENT:		http://cipgweb/qacgi-bin/tt_view.cgi?release=TC_2008&feature=REGQ&cobid=84190&tcobid=587374
'#							QTP Testcase: SrchMySvdSrchsRedsgnHigh07
'#				
'# DESCRIPTION:			This Operate on Search Tree under Change Search dialog
'#
'#
'# PRE-REQUISITE:		RAC Session accessible and My Teacenter Application Search Pane loaded
'#							
'# PARAMETERS   :     	 	
'#														  
'# RETURN VALUE : 	   	TRUE \ FALSE
'#
'# Examples	:		Fn_MyTc_ChangeSearchTreeOperation()	
'#										
'# History	:		Developer Name			Date			Rev. No.			Changes Done			Reviewer	
'#------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'#				Mallikarjun			28-Aug-2010		001
'#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'###########################################################################################################################################################~

Public Function Fn_MyTc_ChangeSearchTreeOperation(sAction, sTreePath, iBase)

	GBL_FAILED_FUNCTION_NAME="Fn_MyTc_ChangeSearchTreeOperation"
	' Variable Declaration:
	Dim ObjChngSrch
	Dim iNode, iCount, sNode, bExist, aArrNodes, iOuterCounter, iTreeNodeRelIndex


	' Error Handler Initialization:


	' Function Core Logic:

	If True =  Fn_ToolbatButtonClick("Select a Search") Then

			'Create UI Object:
			Set ObjChngSrch =  Fn_UI_ObjectCreate( "Fn_MyTc_ChangeSearchTreeOperation",JavaWindow("MyTeamcenter").JavaWindow("Change Search") )

			Wait(2)	
			'Expand Majot Parent Nodes:
			Call Fn_UI_JavaTree_Expand("Fn_MyTc_ChangeSearchTreeOperation", ObjChngSrch, "SearchOptions","My Saved Searches")
			Call Fn_UI_JavaTree_Expand("Fn_MyTc_ChangeSearchTreeOperation", ObjChngSrch, "SearchOptions","System Defined Searches")	
	Else

		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed: Window [Change Search] Does Not Exist." )
		Fn_MyTc_ChangeSearchTreeOperation = False

	End IF

	If iBase <> "" Then
		iCount = iBase
	Else
		iBase = 0	
	End If


	Select Case sAction

		 Case "Change Search Tree Validate"


			'Get the Number of TreePath Elements to be Assessed:
			aArrNodes = Split(sTreePath,",",-1,1)
			
			'Get Total Tree Item Count:
			iNode=Fn_UI_Object_GetROProperty("Fn_MyTc_ChangeSearchTreeOperation",ObjChngSrch.JavaTree("SearchOptions"), "items count")

			For iOuterCounter = 0 To Ubound(aArrNodes)

				'Set Found Flag to [False]
				bExist = False

				'Loop Counter  for TreePaths:
				For  iCount=0 to iNode-1
						sNode = ObjChngSrch.JavaTree("SearchOptions").GetItem (iCount)
						if  sNode = aArrNodes(iOuterCounter) Then
							bExist = True
							Exit For
						End If
				Next	

				'Exit Loop, If Tree Item Not Found After Assessing All Existant Items:	
				If  bExist = False Then
					Exit For
				End If

			Next

			' Log Results Of Validation:
			If bExist = True Then
				Call Fn_Button_Click("Fn_MyTc_ChangeSearchTreeOperation", ObjChngSrch, "Cancel")
				Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Pass: Requested Query Object Exists in Change Search Dialog. ")   	
				Fn_MyTc_ChangeSearchTreeOperation = True
				Set ObjChngSrch = Nothing
				Exit Function
			Else
				Call Fn_Button_Click("Fn_MyTc_ChangeSearchTreeOperation", ObjChngSrch, "Cancel")
				Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Fail: Requested Query Object :["+ aArrNodes(iOuterCounter) +"] Does Not Exist Change Search Dialog. ")   	
				Fn_MyTc_ChangeSearchTreeOperation = False
				Set ObjChngSrch = Nothing
				Exit Function
			End If 	


		 Case "Change Search Tree Node Get Relative Index"

			'Get Total Tree Item Count:
			iNode=Fn_UI_Object_GetROProperty("Fn_MyTc_ChangeSearchTreeOperation",ObjChngSrch.JavaTree("SearchOptions"), "items count")

			'Set Found Flag to [False]
			bExist = False

			'Loop Counter  for TreePaths:
			For  iCount = iBase to iNode-1
				sNode = ObjChngSrch.JavaTree("SearchOptions").GetItem (iCount)
				if  sNode = sTreePath Then
					bExist = True
					Exit For
				End If
			Next	

			' Log Results Of Validation:
			If bExist = True Then
				Call Fn_Button_Click("Fn_MyTc_ChangeSearchTreeOperation", ObjChngSrch, "Cancel")
				iTreeNodeRelIndex = CInt(iBase) + Cint(iCount) 
				Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Pass: Requested Query Object Exists at Tree Index :["+ Cstr(iTreeNodeRelIndex) +"] in Change Search Dialog. ")   	
				Fn_MyTc_ChangeSearchTreeOperation = iTreeNodeRelIndex
				Set ObjChngSrch = Nothing
				Exit Function
			Else
				Call Fn_Button_Click("Fn_MyTc_ChangeSearchTreeOperation", ObjChngSrch, "Cancel")
				Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Fail: Requested Query Object :["+ Cstr(sTreePath) +"] Does Not Exist Change Search Dialog. ")   	
				Fn_MyTc_ChangeSearchTreeOperation = 0
				Set ObjChngSrch = Nothing
				Exit Function
			End If 


																				
		Case Else
						Fn_MyTc_ChangeSearchTreeOperation = False
						'Report Invalid Action Arguement Error
						'Call Fn_WriteLogFile("Fn_MyTc_ChangeSearchTreeOperation", 1, Err.Number,"FAIL: Wrong Action Parameter [" + StrAction + "] for [Fn_MyTc_ChangeSearchTreeOperation] Function Call")
	End Select
	

						
	Fn_MyTc_ChangeSearchTreeOperation = TRUE				 
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_MyTc_ChangeSearchTreeOperation passed with case [" + sAction +"] on Change Search Tree")	

	'Release UI Object:
	Set ObjChngSrch = Nothing

End Function

'##########################################################################################################################################################~
'# FUNCTION NAME:	 	Fn_MyTc_SearchPFFTableOperation
'#
'#
'# MODULE: 			Search Requirement
'#				
'# DESCRIPTION:			This Operate on PFF Table under Search Panel
'#							
'# PARAMETERS   :     	 	
'#														  
'# RETURN VALUE : 	   	TRUE \ FALSE \ Numeric (Index)
'#
'# Examples	:		Case: 
'# 					Fn_MyTc_SearchPFFTableOperation("PFF Row Item GetIndex", "Item1", "Object Name", "")	==> Returns Row Index of "Item1" under Column "Object Name"
'#
'# 					Case: 
'# 					Fn_MyTc_SearchPFFTableOperation("PFF Object Property Validate", "Item1", "Release Status", "TCM Released")	==> Validates Properties of "Item1": Property [Release Status] for Expected Value [TCM Released], Return [True], if matched
'#
'# 					Case: 
'# 					Fn_MyTc_SearchPFFTableOperation("PFF Columns Exist", "", "Object Name:Object Type:Release Status", "")	==> Validates Existance of Property Columns [Object Name:Object Type:Release Status] under PFF Table, Return [True], if all of them exist
'#
'# 					Case: 
'# 					Fn_MyTc_SearchPFFTableOperation("PFF Columns Sort", "", "Object Name", "")	==> Applies Sort on Property Column [Object Name] under PFF Table, Return [True], if successful

'#										
'# History	:		Developer Name			Date			Rev. No.			Changes Done			Reviewer	
'#------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'#				Mallikarjun			01-Sep-2010		001
'#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'###########################################################################################################################################################~

Public Function Fn_MyTc_SearchPFFTableOperation(sAction, sObjectName, sPropertyName, sExpectedValue)
	GBL_FAILED_FUNCTION_NAME="Fn_MyTc_SearchPFFTableOperation"
	Dim aArrCols, objPFFTable,bReturn,iCounter, icolIndex, bFlag, sText, bExist

	'Set False Flag to the Function: 
	Fn_MyTc_SearchPFFTableOperation = False

	'1. Operate on the Refresh Toolbar Button under Search Results Pane
	Call Fn_ToolbatButtonClick("Refresh property formetter search") 

	Set objPFFTable = Fn_UI_ObjectCreate("Fn_MyTc_SearchPFFTableOperation",JavaWindow("MyTeamcenter_Search").JavaTable("PffTable") )

	If sPropertyName <> "" Then
	
		aArrCols = split(sPropertyName, ":",-1,1)
		If UBound(aArrCols) = 0 Then

			'Get Count Property Column Count of PFF Table:
			bReturn = objPFFTable.GetROProperty("cols")
	
			'Extract the index of Requisite Property Column:
			bExist = False
			For iCounter=0 to bReturn-1
				sText = objPFFTable.GetColumnName(iCounter)
				If Cstr(sText) =  Cstr(sPropertyName) Then
					bExist = True
					Exit For
				End If
			Next

			If bExist = False Then
				Fn_MyTc_SearchPFFTableOperation = False
				Exit Function
			Else
				icolIndex = iCounter
			End If
		End If
	
	End If

	Select Case sAction
				
		 Case "PFF Row Item GetIndex"
   				bFlag = -1
				'Count number of rows of Table
				bReturn = objPFFTable.GetROProperty("rows")	
				'Extract the index of row at which the object exist.
				For iCounter=0 to bReturn - 1
					sText = objPFFTable.GetCellData(iCounter,icolIndex)						
						If IsNumeric(sObjectName) Then
							 If cstr(sText) = cstr(cint(sObjectName))  Then
								 bFlag = iCounter
								 Exit for
							End If
						elseIf cstr(sText) = cstr(sObjectName)  Then
								 bFlag = iCounter
								 Exit for
						End If									
				Next
				If bFlag = -1 Then
						Fn_MyTc_SearchPFFTableOperation = -1
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Fn_MyTc_SearchPFFTableOperation : Failed to Get Index Of Object ["+ Cstr(sObjectName) +"] under Column ["+ Cstr(sPropertyName) +"]")	
						Set objPFFTable = Nothing						
						Exit function
				Else 
						sExpectedValue = Cstr(iCounter + 1)
						Fn_MyTc_SearchPFFTableOperation = Cstr(iCounter + 1)
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Fn_MyTc_SearchPFFTableOperation: Successfully Got Index of Object ["+ Cstr(sObjectName) +"] :: ["+ Cstr(iCounter + 1) +"]")	
						Set objPFFTable = Nothing
						Exit function
				End If				

				
		Case "PFF Object Property Validate"
				'Count number of rows of Table
				bReturn = objPFFTable.GetROProperty("rows")

				'Extract the index of row of which relation is to be changed
				bFlag = -1
				For iCounter=0 to bReturn - 1
					sText = objPFFTable.GetCellData(iCounter,"Object Name")
					If IsNumeric(sObjectName) Then
						If cstr(sText) = cstr(cint(sObjectName))  Then
							bFlag = iCounter
							Exit for
						End If
					elseIf cstr(sText) = cstr(sObjectName)  Then
						 bFlag = iCounter
						 Exit for
					End If
				Next

				If bFlag = -1 Then
					Fn_MyTc_SearchPFFTableOperation = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Fn_MyTc_SearchPFFTableOperation : Failed to Get Index Of Object ["+ Cstr(sObjectName) +"] under Column [Object Name]")	
					Set objPFFTable = Nothing						
					Exit function
				End If

				sText = objPFFTable.GetCellData(iCounter, sPropertyName)
				If cstr(sText) = cstr(sExpectedValue)  Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Pass: Property of Object Validated under PFF Table. Object:: ["+ Cstr(sObjectName) +"] Property: ["+ sPropertyName +"] Value:["+ sExpectedValue +"]")
					Fn_MyTc_SearchPFFTableOperation = True
					Set objPFFTable = Nothing
					Exit function				
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: " + sPropertyName + " Validation Failed")
					Fn_MyTc_SearchPFFTableOperation = False
					Set objPFFTable = Nothing
					Exit function
				End If


   		 Case "PFF Columns Exist"
				'Count number of rows of Table
				bReturn = objPFFTable.GetROProperty("cols")	
				'Extract the index of row of which relation is to be changed

				'aArrCols = split(sObjectName, ":",-1,1)

				For oCounter=0 to UBound(aArrCols)
					bExist = False
					For iCounter=0 to bReturn-1
					    sText = objPFFTable.GetColumnName(iCounter)					
						If sText =  aArrCols(oCounter) Then
							bExist = True
							Exit For
						End If
					Next
					If bExist = False Then
						Exit For
						Fn_MyTc_SearchPFFTableOperation = False
					End If
				Next
			
				If True =  bExist Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Pass: Requested Column is present in PFF Table. ")
					Fn_MyTc_SearchPFFTableOperation = True
					Set objPFFTable = Nothing
					Exit function					
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Fail: Requested Column :["+ aArrCols(oCounter) + "] Does Not Exist in PFF Table. ")
					Fn_MyTc_SearchPFFTableOperation = False
					Set objPFFTable = Nothing
					Exit function					
				End If



   		 Case "PFF Column Sort"

																				
		Case Else
				Fn_MyTc_SearchPFFTableOperation = False
				'Report Invalid Action Arguement Error
				'Call Fn_WriteLogFile("Fn_MyTc_SearchPFFTableOperation", 1, Err.Number,"FAIL: Wrong Action Parameter [" + StrAction + "] for [Fn_MyTc_SearchPFFTableOperation] Function Call")
	End Select
	
	'Release Objects:
	Set objPFFTable = Nothing 

End Function

'####################################################################################################################################		
'########################################################      Verify  Search Tab title			###############################################################~
'#
'# FUNCTION NAME:	Fn_MyTcSrch_HdrMsgVerify(sAction, sHeaderTitle)
'#
'#  MyCommunity ID :	Not available
'#
'# MODULE: 						 Search Requirement 
'#
'# DESCRIPTION:			 1. To validate the Title of  Search tab 
'#							
'#		
'# PRE-REQUISITE:		1. Search Tab should b open
'#											
'#							
'# PARAMETERS   :       sAction: Name of the case to exercise pertaining to saved search
'#										 	 sHeaderTitle: Title of the Search header
'#											
'#
'# RETURN VALUE : 		TRUE \ FALSE
'#
'#Examples	:				Fn_MyTcSrch_HdrMsgVerify("HeaderVerify", "Item")
'#										
'#	History	:						Developer Name			Date			Version				Changes Done			Reviewer	
'#------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'#									Deepak kumar			31	August 2010	1.0																	Sunil Rai
'#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'########################################################################################################################################################~
'########################################################    Verify  Search Tab title		###############################################################~

 Public Function   Fn_MyTcSrch_HdrMsgVerify(sAction, sHeaderTitle)
 GBL_FAILED_FUNCTION_NAME="Fn_MyTcSrch_HdrMsgVerify"
 Dim sValue
			Select Case sAction

					Case "HeaderVerify"
							sValue = JavaWindow("MyTeamcenter_Search").JavaStaticText("SrchHdrMsg").GetROProperty("attached text")
							If sValue = sHeaderTitle Then
								Fn_MyTcSrch_HdrMsgVerify = True
								Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Pass: Verified Title Send as parameter ")   
							Else
								Fn_MyTcSrch_HdrMsgVerify = False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Fail: Wrong Title Send as parameter ")   
							End If
					Case Else 
				End Select

End Function

'#######################################################################################################################################################~
'########################################################    Set Requisite Search Preference		###############################################################~
'#
'# FUNCTION NAME:	Fn_MyTc_OrganizedSavedSrchOperations
'#
'#  MyCommunity ID :	Not available
'#
'# MODULE: 						 Search Requirement 
'#
'# DESCRIPTION:			 1. Check/UnCheck the[Is Shared] box
'#											2. Click on [Create In] button on [Add Search to My Saved Searches]
'#											3. Expand and select prefered Tree Path under [My Saved Search]
'#											
'#											Case: Add_To_My_Saved_Search  		' 
'#													a. Specify [Name] of the saved search									
'#													b. Specify Folder Name on [Folder Information] dialog and click [OK]
'#													c. Select Path
'#											Case: Saved Search Delete
'#													a. Click on [Delete] button
'#													b. Click [OK] on Warning dialog		
'#											Case: Saved Search Rename
'#												a. Click on [Rename] button
'#												b. Specify New Name through send Key operation
'#												c. Send [Enter] Key						
'#											Case: Saved Search Validate		' Need to pass Full Reference path in "sSearchName" for existance check.							
'#												a. Validate existance of saved search
'#													4. Click [OK]
'#		
'# PRE-REQUISITE:		1. RAC Session accessible and My Teacenter Application Search Pane loaded
'#											 2. Search Criteria applied with various input values 
'#									>>   3. Click on Toolbar button [Add Search To My Saved Searches] under Search Pane
'#							
'# PARAMETERS   :       sAction: Name of the case to exercise pertaining to saved search
'#                                           bIsShared : ON/OFF
'#										 	 sSourceFolderPath: Existing Search Folder Path
'#											 sNewName: New Name for folder OR Search
'#
'# RETURN VALUE : 		TRUE \ FALSE
'#
'#Examples	:		Fn_MyTc_OrganizedSavedSrchOperations("Add_To_My_Saved_Search",  "ON", "A:B:C", "NewName") 
'#										
'#	History	:						Developer Name			Date			   Version					Changes Done			                                  Reviewer	
'#------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'#										Deepak kumar		   28-august 2010		1.0																	                               Sunil Rai
'#                                      Nilesh Gadekar        06-Jan-2011                         Added Case "Saved_Search_Rename"         Deepak Kumar     
'#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'########################################################################################################################################################~
'########################################################    Set Requisite Search Preference		###############################################################~

Public Function  Fn_MyTc_OrganizedSavedSrchOperations(sAction,  bIsShared, sSourceFolderPath, sNewName)
GBL_FAILED_FUNCTION_NAME="Fn_MyTc_OrganizedSavedSrchOperations"
Dim  iCount, bExist, aRefPath, afullRefPath,  bFolder
Dim  ObjAddSrch, iCont, sFullPath,ObjOrgSavedSrchTree
	
		If True = Fn_UI_ObjectExist("Fn_MyTc_OrganizedSavedSrchOperations",JavaWindow("MyTeamcenter_Search").JavaWindow("Organize My Saved Searches")) Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "My Saved Searches window allready exist")		
		Else
			Fn_ToolbatButtonClick("Organize My Saved Searches")
            Call Fn_ReadyStatusSync(2)  'Added Sync by Nilesh on 11-Oct-2013
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Organize My Saved Searches Button Successfully Clicked")
		End If
		Set ObjAddSrch = Fn_UI_ObjectCreate( "Fn_MyTc_OrganizedSavedSrchOperations",JavaWindow("MyTeamcenter_Search").JavaWindow("Organize My Saved Searches"))
		Select Case sAction

				Case "Add_To_My_Saved_Search" 
			
						'	[Example ( last parameter is for name and its compulsary )]			Fn_MyTc_SrchSavedSearchOperation("Add_To_My_Saved_Search",  "ON", "A:B:C", "NewName") 
						'<<<<<<<<<<<<<<<<<<<<<<<  Add to my saved search  >>>>>>>>>>>>>>>>>>>>>>>>> 
						If True = Fn_UI_ObjectExist("Fn_MyTc_OrganizedSavedSrchOperations",JavaWindow("MyTeamcenter_Search").JavaWindow("Organize My Saved Searches")) Then

								If sSourceFolderPath <> "" Then

											aRefPath = Split(sSourceFolderPath,":")
											For iCount = 0 To UBound(aRefPath) 
												If iCount = 0 Then
													afullRefPath = aRefPath(iCount)
												Else
													afullRefPath = afullRefPath + ":"+CStr(aRefPath(iCount))
												End If
												wait 5												
												Call Fn_UI_JavaTree_Expand("Fn_MyTc_OrganizedSavedSrchOperations",ObjAddSrch,"Existing Saved Searches",afullRefPath)
											Next
											wait 5
											If True =   Fn_JavaTree_Select("Fn_MyTc_OrganizedSavedSrchOperations", ObjAddSrch, "Existing Saved Searches", afullRefPath) then
												bExist = True
											Else
												bExist = False
											End If										
														If True =  bExist Then
																	'CreateFolder
																	Call Fn_Button_Click("Fn_MyTc_OrganizedSavedSrchOperations", ObjAddSrch, "New Folder...")
																	If True = Fn_UI_ObjectExist("Fn_MyTc_OrganizedSavedSrchOperations",JavaWindow("MyTeamcenter_Search").JavaWindow("Folder Information") ) Then
																		Call Fn_UI_EditBox_Type("Fn_MyTc_OrganizedSavedSrchOperations",JavaWindow("MyTeamcenter_Search").JavaWindow("Folder Information"),"Name", sNewName )
																		Call Fn_Button_Click("Fn_MyTc_OrganizedSavedSrchOperations", JavaWindow("MyTeamcenter_Search").JavaWindow("Folder Information"), "OK")		
																		Call Fn_ReadyStatusSync(5)
																		Call Fn_Button_Click("Fn_MyTc_OrganizedSavedSrchOperations", ObjAddSrch, "Cancel")
																	Else
																		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed: Window Folder Create Does Not Exist ")
																		Fn_MyTc_OrganizedSavedSrchOperations = False
																	End If
														End If 	
											
								End If
						
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Passed:Search Successfully Added to My Saved Searches" )
								Fn_MyTc_OrganizedSavedSrchOperations = True
						Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed: Window Does Not Exist." )
								Fn_MyTc_OrganizedSavedSrchOperations = False
						End If
						'<<<<<<<<<<<<<<<<<<<<<<<  Add to my saved search  >>>>>>>>>>>>>>>>>>>>>>>>>

			
				Case "Saved_Search_Delete" ' Fn_MyTc_OrganizedSavedSrchOperations("Saved_Search_Delete",  "", sSourceFolderPath, "")

                    		Call Fn_UI_JavaTree_Activate_Select_Node("Fn_MyTc_OrganizedSavedSrchOperations",ObjAddSrch,"Existing Saved Searches",sSourceFolderPath)
							'Added by Akshay
							Call JavaWindow("MyTeamcenter_Search").JavaWindow("Organize My Saved Searches").JavaButton("Delete").Click

							If True = Fn_UI_ObjectExist("Fn_MyTc_OrganizedSavedSrchOperations", ObjAddSrch.JavaWindow("Warning")) Then
								Call Fn_Button_Click("Fn_MyTc_SrchSavedSearchOperation", ObjAddSrch.JavaWindow("Warning"), "OK")
                                Call Fn_Button_Click("Fn_MyTc_SrchSavedSearchOperation", ObjAddSrch, "Cancel")
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Sucessfully Deleted Requested Query:" &sSourceFolderPath)
								Fn_MyTc_OrganizedSavedSrchOperations = True
							Else
								Call Fn_Button_Click("Fn_MyTc_SrchSavedSearchOperation", ObjAddSrch, "Cancel")
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed To  Delete Requested Query:" &sSourceFolderPath)
								Fn_MyTc_OrganizedSavedSrchOperations = False
							End If
						
        			
				Case "Saved_Search_Rename"   ' Fn_MyTc_SrchSavedSearchOperation("Saved_Search_Rename",  "", sSourceFolderPath, sNewName)
					
							If True = Fn_UI_ObjectExist("Fn_MyTc_OrganizedSavedSrchOperations",JavaWindow("MyTeamcenter_Search").JavaWindow("Organize My Saved Searches")) Then

								If sSourceFolderPath <> "" Then

											aRefPath = Split(sSourceFolderPath,":")
											For iCount = 0 To UBound(aRefPath) 
												If iCount = 0 Then
													afullRefPath = aRefPath(iCount)
												Else
													afullRefPath = afullRefPath + ":"+CStr(aRefPath(iCount))
												End If
												Call Fn_UI_JavaTree_Expand("Fn_MyTc_OrganizedSavedSrchOperations",ObjAddSrch,"Existing Saved Searches",afullRefPath)
											Next
											wait 5
											If True =   Fn_JavaTree_Select("Fn_MyTc_OrganizedSavedSrchOperations", ObjAddSrch, "Existing Saved Searches", afullRefPath) then
												bExist = True
											Else
												bExist = False
											End If										
														If True =  bExist Then
																	'Click on Rename Button
																	Call Fn_Button_Click("Fn_MyTc_OrganizedSavedSrchOperations", ObjAddSrch, "Rename")
																	wait(5)
																	ObjAddSrch.JavaTree("Existing Saved Searches").Type sNewName
																	wait(5)
                                                                    Call Fn_ReadyStatusSync(2)
																	Call Fn_Button_Click("Fn_MyTc_OrganizedSavedSrchOperations", ObjAddSrch, "Cancel")
																	wait(5)
                                                 		End If 	
											
								End If
						
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Passed:Successfully Renamed Saved Searches" )
								Fn_MyTc_OrganizedSavedSrchOperations = True
							Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed: Window Does Not Exist." )
								Fn_MyTc_OrganizedSavedSrchOperations = False
							End If

				Case "Saved_Search_Validate" ' Fn_MyTc_SrchSavedSearchOperation("Saved_Search_Validate",  "", sSourceFolderPath, "")

							bExist = False
                        	bFolder = Split(sSourceFolderPath,":") 
							For iCont = 0 To UBound(bFolder) - 1
								If iCont = 0 Then
									sFullPath = bFolder(iCont)
								Else
									sFullPath = sFullPath + ":"+CStr(bFolder(iCont))
								End If
                            Call Fn_UI_JavaTree_Expand("Fn_MyTc_OrganizedSavedSrchOperations",ObjAddSrch,"Existing Saved Searches",sFullPath)
							Next
							wait 5
							If True =  Fn_UI_JavaTree_Activate_Select_Node("Fn_MyTc_OrganizedSavedSrchOperations",ObjAddSrch,"Existing Saved Searches",sSourceFolderPath) then
								bExist = True
							Else
								bExist = False
							End If
							If bExist = True Then
									Call Fn_Button_Click("Fn_MyTc_OrganizedSavedSrchOperations", ObjAddSrch, "Cancel")
									Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Pass: Requested Query Object Exists. ")   	
									Fn_MyTc_OrganizedSavedSrchOperations = True
							Else
									Call Fn_Button_Click("Fn_MyTc_OrganizedSavedSrchOperations", ObjAddSrch, "Cancel")
									Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Fail: Requested Query Object Does Not Exist. ")   	
									Fn_MyTc_OrganizedSavedSrchOperations = False
							End If

                Case "Saved_Search_Button_Enabled" ' Fn_MyTc_OrganizedSavedSrchOperations("Saved_Search_Button_Enabled",  "", "Rename", "")	

						If ObjAddSrch.JavaButton(sNewName).GetROProperty("enabled") =1 then
								Call Fn_Button_Click("Fn_MyTc_OrganizedSavedSrchOperations", ObjAddSrch, "Cancel")
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Passed:Button Enabeled" )
								Fn_MyTc_OrganizedSavedSrchOperations = True
						Else
								Call Fn_Button_Click("Fn_MyTc_OrganizedSavedSrchOperations", ObjAddSrch, "Cancel")
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed:  Button  not Enabeled." )
								Fn_MyTc_OrganizedSavedSrchOperations = False
						End If

				End Select

Set ObjAddSrch = Nothing

End Function


'########################################################  Modify a Local Query with defined inputs as specified     ################################################~
'#
'# FUNCTION NAME:	 	  Fn_QryBldr_CrtLocQryFrmExisting_Extn
'#
'#  MyCommunity ID :	
'#
'# MODULE: 						 Search Requirement, MyTC 
'#
'# PRE-REQUISITE:		RAC Session accessible and [Query Builder] Application loaded & Query to be modified should be loaded.
'#
'# DESCRIPTION:			 Modify a Local Query with defined inputs as specified
'#											 1. Input [Name] field details
'#											 2. Input [Description] field details
'#											 3. Click [Search Class] button
'#														 3a. Tree navigate to trace [Class Attribute], For Example: POM_object:WorkspaceObject:ItemRevision
'#														 3b. Select/Double Click on Tree Node of prefered [Class Attribute]
'#														 3b. Close [Class Attribute] window
'#											 4. Check the [Search Class] button label updated to requiste class
'#											 5. Set the [Display Setting] to required option [Class/All Attributes]
'#											 6. Set [Show Indented Results] option [On/Off]
'#											 7. Select attribute from [Attribute Selection] Tree
'#													>>	 Note: 	<<
'#													 7.1 Please note that the function arguement is a array of required attributes, seperated by ":"
'#													 7.2 Select the required attribute iteratively and click on [+] button to add the attribute
'#											 8. Invoke [Create] button
'#										
'# PARAMETERS   :      sQueryName: Name of the Local User Query
'#										   sQueryDescription: Description of the Local User Query
'#										   sSearchClass: Attribute Class POM Object of the Query
'#									       sDisplaySettings: Display Settings [Class/All Attributes]
'#										   bShowIndentedResults: [Show Indented Results] option [On/Off]
'#
'#									      aAttributes: Array of the class attributes to be added to the Search Query
'#													 >>   Note: ( 1.) Multiple Attributes to be seperated by "~" ( Tilde)  [ EXAMPLE - >> Attrib1~Attrib2~Attrib3]
'#  																   ( 2.) Inside Attribute  inner values to be seperated by "," (Comma)  [ EXAMPLE - >> First, Second, Third]
'#  																  ( 3.) Inside Values Reference Path  to be seperated by ":" (Colon)  [ EXAMPLE - >> Dataset:Revision
'#
'#				             	     								>>	InnerValues  = "First, Second, Third"
'#																						 First =  refpath for Attrib in main window - will activate /double click
'#				  				   								 				Second =  [For class Attrib Sel Dialog ] ->> FullRefPath Attrib to be selected.				 (Set the Class)   				
'#																									[For Class Selection Dialog} ->>EditBoxValue]
'#																					Third = Refpath for Attrib in main window - will activate /double click
'#
'#												->>	 (Multiple Attribute) Example>>  (First, Second, Third ~ First, Second, Third~ First, Second, Third)  << -
'#												->>	 (First, Second, Third ) Example>>  Refpath1,RefPath2,Refpath3<< -
'#												->>	 (Refpath1 ) Example>>  Home:Child << -
'#										sRemAttribNmIndex: Cell Data
'#
'# RETURN VALUE : 		TRUE \ FALSE
'#
'#Examples	:				 Fn_QryBldr_CrtLocQryFrmExisting_Extn("LocQuery1", "Local Test Query", "Dataset", "AllAttributes", "OFF", "Dataset:Referenced By,ItemRevision:IMAN_specification,Dataset:Specifications [ ItemRevision ]:Revision~Dataset:pid" , "BOM ID:CMDisposition:Synopsis") 
'#							 Fn_QryBldr_CrtLocQryFrmExisting_Extn("LocQuery1", "Local Test Query", "Dataset", "AllAttributes:RealNames", "ON", "Dataset:pid" , "") 
'#
'#	History	:						Developer Name			Date			Version				Changes Done			Reviewer	
'#------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'#										Sandeep		24-09-10		1.0																	Tushar B
'#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function  Fn_QryBldr_CrtLocQryFrmExisting_Extn(sQueryName, sQueryDescription, sSearchClass, sDisplaySettings, bShowIndentedResults, aAttributes, sRemAttribNmIndex)  
  GBL_FAILED_FUNCTION_NAME="Fn_QryBldr_CrtLocQryFrmExisting_Extn"
Dim  ArrDispSet,  OuterArrAttrib, iOuterCounter, ArrAttrib,  ArrinnAttrib, iCounter
Dim ObjQryApp, ObjQryAttribSel,bReturn,RmvArrAttrib,iCnt,iRows,iCot,iDelCounter,strData,bFlag

Set ObjQryApp =  Fn_UI_ObjectCreate( "Fn_QryBldr_CrtLocQryFrmExisting_Extn",JavaWindow("Search_QueryBuilder").JavaWindow("QueryBuilder"))

		
				'++++++++++<<    Input [Name] field details >>++++++++++
				If sQueryName <> "" Then
					Call Fn_Edit_Box("Fn_QryBldr_CrtLocQryFrmExisting_Extn",ObjQryApp,"Name",sQueryName)
				End If
			
				'++++++++++<<   Input [Description] field details>>++++++++++
				If sQueryDescription <>"" Then
					Call Fn_Edit_Box("Fn_QryBldr_CrtLocQryFrmExisting_Extn",ObjQryApp,"Description",sQueryDescription)
				End If
			
				'++++++++++<<   Click [Search Class] button >>++++++++++
				If  sSearchClass <> "" Then
					Call Fn_CheckBox_Set("Fn_QryBldr_CrtLocQryFrmExisting_Extn", ObjQryApp, "SrchClass", "ON")
					Call Fn_Edit_Box("Fn_QryBldr_CrtLocQryFrmExisting_Extn",ObjQryApp,"Class/Attribute Selection",sSearchClass)
					Call Fn_Button_Click( "Fn_QryBldr_CrtLocQryFrmExisting_Extn", ObjQryApp, "Find")
					ObjQryApp.JavaObject("Close").Click 1,1
				End If
			
				 '++++++++++<<  Set the [Display Setting] to required option [Class/All Attributes] >>++++++++++
				 If sDisplaySettings <> ""  Then
					 ArrDispSet = split(sDisplaySettings, ":", -1,1)
					 If Ubound(ArrDispSet) = 1 Then
						Call Fn_CheckBox_Set("Fn_QryBldr_CrtLocQryFrmExisting_Extn", ObjQryApp, "DisplaySettings", "ON")
						Call  Fn_UI_JavaRadioButton_SetON("Fn_QryBldr_CrtLocQryFrmExisting_Extn",ObjQryApp, ArrDispSet(0))
						Call  Fn_UI_JavaRadioButton_SetON("Fn_QryBldr_CrtLocQryFrmExisting_Extn",ObjQryApp, ArrDispSet(1))
					Else
						Call Fn_CheckBox_Set("Fn_QryBldr_CrtLocQryFrmExisting_Extn", ObjQryApp, "DisplaySettings", "ON")
						Call  Fn_UI_JavaRadioButton_SetON("Fn_QryBldr_CrtLocQryFrmExisting_Extn",ObjQryApp, sDisplaySettings )
					 End If
					ObjQryApp.JavaObject("Close").Click 1,1
				 End If
			
				 '++++++++++<<  Set [Show Indented Results] option [On/Off] >>++++++++++
				 If bShowIndentedResults <> ""  Then
					Call Fn_CheckBox_Set("Fn_QryBldr_CrtLocQryFrmExisting_Extn", ObjQryApp, "ShowIndentedResults", bShowIndentedResults)
				 End If
			
				'++++++++++<<   Select attribute from [Attribute Selection] Tree >>++++++++++
				If  aAttributes <> "" Then
					OuterArrAttrib = split(aAttributes, "~", -1,1)
					For iOuterCounter = 0 To Ubound(OuterArrAttrib)
							ArrAttrib = split(OuterArrAttrib(iOuterCounter), ",", -1, 1)				
							For iCounter = 0 to Ubound(ArrAttrib) 		
										ArrinnAttrib = split(ArrAttrib(iCounter), ":", -1, 1)
										Select Case iCounter
										Case "0"
												ObjQryApp.JavaTree("AttributeSelectionList").Activate ArrAttrib(iCounter) 
										Case "1"
												If  True = Fn_UI_ObjectExist("Fn_QryBldr_CrtLocQryFrmExisting_Extn", ObjQryApp.JavaDialog("ClassAttributeSelection") )Then
												
													Set ObjQryAttribSel = ObjQryApp.JavaDialog("ClassAttributeSelection") 
													Call Fn_CheckBox_Set("Fn_QryBldr_CrtLocQryFrmExisting_Extn", ObjQryAttribSel, "CAS_SrchClass", "ON")
													Call Fn_Edit_Box("Fn_QryBldr_CrtLocQryFrmExisting_Extn",ObjQryAttribSel,"CAS_Edit",ArrinnAttrib(0) )
													Call Fn_Button_Click( "Fn_QryBldr_CrtLocQryFrmExisting_Extn", ObjQryApp, "Find")
													ObjQryApp.JavaObject("Close").Click 1,1
													ObjQryAttribSel.JavaTree("CAS_SrchTree").Activate ArrAttrib(iCounter) 
												End If
												If  True = Fn_UI_ObjectExist("Fn_QryBldr_CrtLocQryFrmExisting_Extn",ObjQryApp.JavaDialog("ClassSelectionDialog")  )Then
													Set ObjQryAttribSel = ObjQryApp.JavaDialog("ClassSelectionDialog")  
													Call Fn_Edit_Box("Fn_QryBldr_CrtLocQryFrmExisting_Extn",ObjQryAttribSel,"SelectionField",ArrAttrib(iCounter) )
													Call Fn_Button_Click( "Fn_QryBldr_CrtLocQryFrmExisting_Extn", ObjQryAttribSel, "ClassAttSelFind")
													Call Fn_Button_Click( "Fn_QryBldr_CrtLocQryFrmExisting_Extn", ObjQryAttribSel, "CSDOK")
												End If										               
										Case "2"
												ObjQryApp.JavaTree("AttributeSelectionList").Activate ArrAttrib(iCounter) 
										End Select		
							Next			
					Next
				End If

	If sRemAttribNmIndex <> "" Then
					RmvArrAttrib = split(sRemAttribNmIndex, ":", -1,1)
					iRows = JavaWindow("Search_QueryBuilder").JavaWindow("QueryBuilder").JavaTable("SrchCriteriaTable").GetROProperty("rows")
					For iCnt = 0 to iRows - 1	
	                        iRows = JavaWindow("Search_QueryBuilder").JavaWindow("QueryBuilder").JavaTable("SrchCriteriaTable").GetROProperty("rows")
							If iCnt>iRows - 1 Then
								Exit For
							End If
							strData= JavaWindow("Search_QueryBuilder").JavaWindow("QueryBuilder").JavaTable("SrchCriteriaTable").GetCellData(iCnt,"User Entry L10N Key")
                            For iCounter=0 To Ubound(RmvArrAttrib)
								bFlag=False
								If strData=RmvArrAttrib(iCounter) Then
									bFlag=True
									Exit For
								End If
							Next
							If bFlag=False Then
								JavaWindow("Search_QueryBuilder").JavaWindow("QueryBuilder").JavaTable("SrchCriteriaTable").SelectRow iCnt
								Call Fn_Button_Click( "Fn_QryBldr_CrtLocQryFrmExisting_Extn", ObjQryApp, "Remove")	
								iCnt=0
							End If
					Next
	End If 
	iRows = JavaWindow("Search_QueryBuilder").JavaWindow("QueryBuilder").JavaTable("SrchCriteriaTable").GetROProperty("rows")
	If Ubound(RmvArrAttrib)=iRows-1 Then
		Fn_QryBldr_CrtLocQryFrmExisting_Extn = True
	Else
		Fn_QryBldr_CrtLocQryFrmExisting_Extn = False
	End If
				'++++++++++<<  Invoke [Create] button >>++++++++++
				 If True =  Fn_Button_Click( "Fn_QryBldr_CrtLocQryFrmExisting_Extn", ObjQryApp, "Create") Then
                        Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Pass: Local Query Successfully Modified with new name. ")   	
						Fn_QryBldr_CrtLocQryFrmExisting_Extn = True
				Else
						Call Fn_Button_Click( "Fn_QryBldr_CrtLocQryFrmExisting_Extn", ObjQryApp, "Clear")
						Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Fail: Local Query Failed to be Modified with new name. ")   	
						Fn_QryBldr_CrtLocQryFrmExisting_Extn = False
				End If
		
	Set ObjQryApp = Nothing
	Set ObjQryAttribSel = Nothing
	End Function

'####################################################################################################################################		
'########################################################      Verify  Search Tab title			###############################################################~
'#
'# FUNCTION NAME:	Fn_MyTcSrch_SearchPreferenceSetting( sSrchItemPath)
'#
'#  MyCommunity ID :	Not available
'#
'# MODULE: 						 Search Requirement 
'#
'# DESCRIPTION:			 1. To Change  the settings  of  Search tab prefrence option
'#							
'#		
'# PRE-REQUISITE:		1. Search Tab should be open & Preference dialog should also be open
'#											
'#							
'# PARAMETERS   :       sSrchItemPath: Full path for selecting object  from System Defined Searches
'#										 	 
'#											
'#
'# RETURN VALUE : 		TRUE \ FALSE
'#
'#Examples	:				Fn_MyTcSrch_SearchPreferenceSetting( sSrchItemPath)
'#										
'#	History	:						Developer Name			Date			Version				Changes Done			Reviewer	
'#------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'#									Deepak kumar			31	August 2010	1.0																	Sunil Rai
'#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'########################################################################################################################################################~
'########################################################    Verify  Search Tab title		###############################################################~
Public Function   Fn_MyTcSrch_SearchPreferenceSetting( sSrchItemPath)
	GBL_FAILED_FUNCTION_NAME="Fn_MyTcSrch_SearchPreferenceSetting"
	Dim ObjPrefDialog,aSrchItemPath,ObjJavaTree
	aSrchItemPath = Split(sSrchItemPath, ":", -1, 1)

	If True = Fn_UI_ObjectExist("Fn_MyTcSrch_SearchPreferenceSetting",JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("IndexOptions")) Then

			Set ObjPrefDialog = Fn_UI_ObjectCreate("Fn_MyTcSrch_SearchPreferenceSetting",JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("IndexOptions"))
			Call Fn_JavaTree_Select("Fn_MyTc_ClearStaleSearch", ObjPrefDialog, "OptionsTree","Options:Search:General")
			Call  Fn_Button_Click("Fn_MyTcSrch_SearchPreferenceSetting", ObjPrefDialog, "ChangeDefaultSearch")
		
			Set ObjJavaTree = Fn_UI_ObjectCreate("Fn_MyTcSrch_SearchPreferenceSetting",  JavaWindow("MyTeamcenter").JavaWindow("Change Search") )
			Wait 3
			Call Fn_UI_JavaTree_Expand("Fn_MyTcSrch_SearchPreferenceSetting", ObjJavaTree, "SearchOptions",aSrchItemPath(0))
			If True = Fn_JavaTree_Select("Fn_MyTcSrch_SearchPreferenceSetting", ObjJavaTree, "SearchOptions",sSrchItemPath) Then
					Call  Fn_Button_Click("Fn_MyTc_ClearStaleSearch", ObjJavaTree, "OK") 
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Selected Requested Node: ["+sSrchItemPath+"] ")					
			Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Select Requested Node: ["+sSrchItemPath+"] ")			
					Fn_MyTcSrch_SearchPreferenceSetting = False
			End IF

			Call  Fn_Button_Click("Fn_MyTc_ClearStaleSearch", ObjPrefDialog, "OK") 
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Performed all the Steps for Selecting Search item from the list.")
			Fn_MyTcSrch_SearchPreferenceSetting = True
	 Else
			Fn_MyTcSrch_SearchPreferenceSetting = False
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Preference window does not exist & Preference menu not selected  from the menu list.")
	 End IF

End Function

'########################################################  Modify the Preference values    ################################################~
'#
'# FUNCTION NAME:	 	  Fn_MyTcSrch_SrchPreferenceOperation
'#
'#  MyCommunity ID :	
'#
'# MODULE: 						 Search Requirement, MyTC 
'#
'# PRE-REQUISITE:		RAC Session accessible and Search Preference Window should be Open
'#
'# DESCRIPTION:			 Modification of Preferences In the Search Preference Window
'#                                           										
'# PARAMETERS   :      sAction: Action or Case to be selected
'#										   sQueryName: Query Name to be worked upon
'#										   sOption: ON/OFF to be selected for the Query
'#									       iHistSize: History size to be worked upon
'#										   iPageSize: Page size of the Search Result pane
'#										  iRsltLimit: Result limit to be set for the Search result window
'#									      bDisplay: Display of the Saved searches, ON/OFF
'#										  bFlag: True/False
'#
'# RETURN VALUE : 		TRUE \ FALSE
'#
'#Examples	:				 Fn_MyTcSrch_SrchPreferenceOperation("Favorites","Remote...:All Sequence", "OFF:ON", "", "", "", "", "")
'											Fn_MyTcSrch_SrchPreferenceOperation("VerifyFavorites","Remote...:All Sequence", "", "", "", "", "", "")
'#
'#	History	:						Developer Name			Date			Version				Changes Done			Reviewer	
'#------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'#										Deepak Kumar		06-10-10		1.0																	Sunil R
'#------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'#											Sagar Shivade	22-Dec			9.0 Porting				Chnaged Hierarchy  Teamcenter:Search
'#------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'#											Nilesh Gadekar	11-Jul-2012			10.0 Porting				Modified code for Search load limit functionality
'#------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'#											Sandeep Navghane		25-Dec-2012								Modified case : Favorites & Added Case : VerifyFavorites
'#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_MyTcSrch_SrchPreferenceOperation(sAction,sQueryName, sOption, iHistSize, iPageSize, iRsltLimit, bDisplay, bFlag)
GBL_FAILED_FUNCTION_NAME="Fn_MyTcSrch_SrchPreferenceOperation"
Dim iCntr, iRowCount, iCount, sData,bReturn, sQryName, sOpt,objPrefDialog,iLoadLimit,aAction,objPrefRowCount

			If Not Fn_UI_ObjectExist("Fn_MyTcSrch_SrchPreferenceOperation",JavaWindow("MyTeamcenter_Search").JavaWindow("SrchDefaultWndw").JavaDialog("Options")) Then
				Call Fn_MenuOperation("Select","Edit:Options...")
			End If

			Set objPrefDialog =Fn_UI_ObjectCreate("Fn_MyTcSrch_SrchPreferenceOperation",JavaWindow("MyTeamcenter_Search").JavaWindow("SrchDefaultWndw").JavaDialog("Options"))
			'Code Added by Nilesh to take care of Load limit parameter on 11-Jul-12
			If Instr(sAction,"~") >0 Then
				aAction=Split(sAction,"~",-1,1)
				If Ubound(aAction)<>0 Then
					sAction=aAction(0)
					iLoadLimit=aAction(1)
				End If
			End If
			'End
                Select Case sAction

                Case "VerifyFavorites"
							Call Fn_UI_JavaTree_Expand("Fn_MyTcSrch_SrchPreferenceOperation", objPrefDialog, "OptionsTree","Options:Search")
							Call Fn_JavaTree_Select("Fn_MyTcSrch_SrchPreferenceOperation", objPrefDialog, "OptionsTree","Options:Search:Favorites")
							sQryName=Split(sQueryName, ":", -1, 1)
							For iCntr = 0 To Ubound(sQryName)
								bReturn=False
								iRowCount=objPrefDialog.JavaTable("FavoritesTable").GetROProperty("rows")
								For iCount=0 To iRowCount - 1
									sData=objPrefDialog.JavaTable("FavoritesTable").GetCellData(iCount,"Query Name")
									If Trim(sQryName(iCntr)) = Trim(sData) Then
										If objPrefDialog.JavaTable("FavoritesTable").GetCellData(iCount,"Favorite")=1 Then
											bReturn=True
										End If
										Exit for
									End if
								Next
								If bReturn=False Then
									Exit for
								End If
							Next
							If bReturn=True Then
								Fn_MyTcSrch_SrchPreferenceOperation=true
							Else
								Fn_MyTcSrch_SrchPreferenceOperation=false
								Call Fn_Button_Click("Fn_MyTcSrch_SrchPreferenceOperation",objPrefDialog,"Close")   
								Set objPrefDialog =Nothing
								Exit function
							End If
                Case "Favorites"
								Call Fn_UI_JavaTree_Expand("Fn_MyTcSrch_SrchPreferenceOperation", objPrefDialog, "OptionsTree","Options:Search")
								Call Fn_JavaTree_Select("Fn_MyTcSrch_SrchPreferenceOperation", objPrefDialog, "OptionsTree","Options:Search:Favorites")
                                
                                sQryName=Split(sQueryName, ":", -1, 1)
                                sOpt=Split(sOption, ":", -1, 1)
                                Set objPrefRowCount=objPrefDialog.JavaTable("FavoritesTable")
                                For iCntr = 0 To Ubound(sQryName)
                                iRowCount=Fn_UI_Object_GetROProperty("Fn_MyTcSrch_SrchPreferenceOperation",objPrefRowCount,"rows")
                                          '      iRowCount=objPrefDialog.JavaTable("FavoritesTable").GetROProperty("rows")
                                                For iCount=0 To iRowCount - 1
                                                                sData=objPrefDialog.JavaTable("FavoritesTable").GetCellData(iCount,"Query Name")
                                                                If Trim(sQryName(iCntr)) = Trim(sData) Then
                                                                                If sOpt(iCntr) = Ucase("ON") Then
																					objPrefDialog.JavaTable("FavoritesTable").SetCellData iCount,"Favorite",1
                                                                                Else
																					objPrefDialog.JavaTable("FavoritesTable").SetCellData iCount,"Favorite",0
                                                                                End If
                                                                                Exit For
                                                                End If
                                                Next
                                Next
								

                'Case "Remote"

                Case "Results"
								'Commented by Sanjeet on 20-Feb-13
'                               Call Fn_UI_JavaTree_Expand("Fn_MyTcSrch_SrchPreferenceOperation", objPrefDialog, "FavoritesTree","Teamcenter:Search")
'								Call Fn_JavaTree_Select("Fn_MyTcSrch_SrchPreferenceOperation", objPrefDialog, "FavoritesTree","Teamcenter:Search:Results")
								'Added by Sanjeet on 20-Feb-13
								Call Fn_UI_JavaTree_Expand("Fn_MyTcSrch_SrchPreferenceOperation", objPrefDialog, "OptionsTree","Options:Search")
								Call Fn_JavaTree_Select("Fn_MyTcSrch_SrchPreferenceOperation", objPrefDialog, "OptionsTree","Options:Search:Results")

                                If iHistSize <> "" Then
									Call Fn_Edit_Box("Fn_MyTcSrch_SrchPreferenceOperation",objPrefDialog,"ConfigureSrchHistory",iHistSize)                                                
                                End If

								If iPageSize <> "" Then
									Call Fn_Edit_Box("Fn_MyTcSrch_SrchPreferenceOperation",objPrefDialog,"SetLoadingPageSize",iPageSize)           
                                End If

								If iRsltLimit <> "" Then
									Call Fn_Edit_Box("Fn_MyTcSrch_SrchPreferenceOperation",objPrefDialog,"SetSrchRsltLimit",iRsltLimit)           
                                End If

                                If iLoadLimit <> "" Then
									Call Fn_Edit_Box("Fn_MyTcSrch_SrchPreferenceOperation",objPrefDialog,"LoadLimit",iLoadLimit)           
                                End If

								If bFlag = "ON" Then
									Call Fn_Button_Click("Fn_MyTcSrch_SrchPreferenceOperation",objPrefDialog,"ClearSearchHistory")   
                                End If

                Case "Saved Searches"
								 Call Fn_UI_JavaTree_Expand("Fn_MyTcSrch_SrchPreferenceOperation", objPrefDialog, "FavoritesTree","Search")
								Call Fn_JavaTree_Select("Fn_MyTcSrch_SrchPreferenceOperation", objPrefDialog, "FavoritesTree","Search:Saved Searches")
                               ' JavaWindow("MyTeamcenter_Search").JavaWindow("Preferences").JavaTree("FavoritesTree").Select "Search:Saved Searches

								If bDisplay = "ON" Then
                                    JavaWindow("MyTeamcenter_Search").JavaWindow("Preferences").JavaCheckBox("DisplayMySvdSrchs").Set "ON"
                                Else
                                    JavaWindow("MyTeamcenter_Search").JavaWindow("Preferences").JavaCheckBox("DisplayMySvdSrchs").Set "OFF"
                                End If

                Case Else

		End Select

				Call Fn_Button_Click("Fn_MyTcSrch_SrchPreferenceOperation",objPrefDialog,"Apply")   
				' JavaWindow("MyTeamcenter_Search").JavaWindow("Preferences").JavaButton("Apply").Click micLeftBtn
				bReturn =  Fn_Button_Click("Fn_MyTcSrch_SrchPreferenceOperation",objPrefDialog,"OK")   
				If bReturn = True  Then
					Fn_MyTcSrch_SrchPreferenceOperation = True
				Else
					Fn_MyTcSrch_SrchPreferenceOperation = False
				End If
                ' JavaWindow("MyTeamcenter_Search").JavaWindow("Preferences").JavaButton("OK").Click micLeftBtn

Set objPrefDialog = Nothing
End Function

'########################################################  Modify the Change Search Values   ################################################~
'#
'# FUNCTION NAME:	 	  Fn_MyTcSrch_ChangeSrchOperation
'#
'#  MyCommunity ID :	
'#
'# MODULE: 						 Search Requirement, MyTC 
'#
'# PRE-REQUISITE:		RAC Session accessible and Change Search Dialog Accessible
'#
'# DESCRIPTION:			 Validation of System Defined Searches in the Change Search Dialog
'#                                           										
'# PARAMETERS   :      sAction: Action or Case to be selected
'#										   StrNodeName: Query Name to be worked upon
'#
'# RETURN VALUE : 		TRUE \ FALSE
'#
'#Examples	:				 Fn_MyTcSrch_ChangeSrchOperation("GetIndex","Item...)
'#
'#	History	:						Developer Name			Date			Version				Changes Done			Reviewer	
'#------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'#										Deepak Kumar		06-10-10		1.0																	Sunil R
'#		
'#										Sagar Shivade		22-Dec-10		9.0 Porting	(2.0)		Added  do while case for OK Botton
'#										Sagar Shivade		23-Dec-10		9.0 Porting	(2.1)		Added  if loop to check dialog existance
'#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function Fn_MyTcSrch_ChangeSrchOperation(sAction,StrNodeName)
GBL_FAILED_FUNCTION_NAME="Fn_MyTcSrch_ChangeSrchOperation"
Dim iCounter,sCmpItm

	If Not JavaWindow("MyTeamcenter").JavaWindow("Change Search").Exist Then
		Call Fn_ToolbatButtonClick("Select a Search") 
		Call Fn_ReadyStatusSync(2)
	End If	
		Select Case sAction
		
		Case "GetIndex"
				'Index of Item1
				For iCounter=0 to JavaWindow("MyTeamcenter").JavaWindow("Change Search").JavaTree("SearchOptions").GetROProperty ("items count")-1
					sCmpItm=JavaWindow("MyTeamcenter").JavaWindow("Change Search").JavaTree("SearchOptions").GetItem (iCounter)
					If   sCmpItm = StrNodeName Then
						Fn_MyTcSrch_ChangeSrchOperation = iCounter
							Do While JavaWindow("MyTeamcenter").JavaWindow("Change Search").Exist(10)
									Call  Fn_Button_Click("Fn_MyTcSrch_ChangeSrchOperation",JavaWindow("MyTeamcenter").JavaWindow("Change Search"),"OK")   
							Loop
						
						Exit For
					End If
				Next
	
		Case Else
	
		End Select

End Function


'###########################################	Warning Message Validation for Saved Search under a given folder    #####################################################~
'#
'# FUNCTION NAME:	 	  Fn_MyTcSrch_SrchWarningValidate
'#
'# MODULE: 						 My Teamcenter Search
'#
'# PRE-REQUISITE:	   1. RAC Session accessible and My Teacenter Application Search Pane loaded
'#											2. Search Criteria applied with various input values and Executed
'#											3. Expected Warning Dialog accessible
'#
'# DESCRIPTION:			 Warning Message Validation
'#
'#											1. Validate the Warning Message
'#											2. Click [OK]
'#						
'# PARAMETERS   :      sWindowCaption: Name of the Warning window/dialog						NOTE    - >>	' Default Title  = "Search ..."  << If Default Window title change then pass title in sWindowCaption.
'#											sWarnMessage: Details of Warning message to be validated			NOTE   - >>    ' Need to pass Exactly same Warning Message..
'#
'# RETURN VALUE : 	TRUE \ FALSE
'#
'#Examples	:				  Fn_MyTcSrch_SrchWarningValidate("Warning", "Deleting the folder will not delete the contents of the folder. Do you want to continue?")
'#										
'#	History	:					  Developer Name			Date			Version				Changes Done			Reviewer	
'#------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'#										Sunil Rai					 18/10/2010			1.0																	Sunil Rai
'#	
'#										Sagar Shivade		29/12/10`			9.0 pORING			Replace if loop and compair massage 		Deepak
'#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'########################################################################################################################################################~

Public Function Fn_MyTcSrch_SrchWarningValidate(sWindowCaption, sWarnMessage)
GBL_FAILED_FUNCTION_NAME="Fn_MyTcSrch_SrchWarningValidate"
Dim sDispMsg,objErrorDialog
Set	objErrorDialog = JavaWindow("MyTeamcenter_Search").JavaWindow("Organize My Saved Searches").JavaWindow("Warning")

	'++++++++++<<   Set the Caption of the window >>++++++++++
	If sWindowCaption <> ""  Then
		objErrorDialog.SetTOProperty "title", sWindowCaption
	End If

	If  True = Fn_UI_ObjectExist("Fn_MyTc_SrchWarningValidate",objErrorDialog) Then

			'objErrorDialog.JavaStaticText("WarnMsg").SetTOProperty "label", sWarnMessage
			'If objErrorDialog.JavaStaticText("WarnMsg").Exist Then
			If objErrorDialog.JavaStaticText("WarnMsg").GetROProperty ("label")=sWarnMessage Then	
        			Call Fn_Button_Click("Fn_MyTc_SrchWarningValidate", objErrorDialog,"OK")
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Warning Message Dialog is Verified")
					Fn_MyTcSrch_SrchWarningValidate = TRUE
			Else
					Call Fn_Button_Click("Fn_MyTc_SrchWarningValidate", objErrorDialog,"OK")
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Warning Message is not matched")
					Fn_MyTcSrch_SrchWarningValidate = FALSE
			End If 
	Else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Warning Message Window Not Found")
		Fn_MyTcSrch_SrchWarningValidate = FALSE
	End If

Set objErrorDialog = Nothing
End Function

'######################################################################################################################################################################################
'#
'# FUNCTION NAME	:	Fn_SrchExtMultiAppsTrgtListOperation
'#
'# MODULE			: 	My Teamcenter Search
'#
'# PRE-REQUISITE	:	1. RAC Session accessible and My Teacenter Application Search Pane loaded
'#						2. Search window should open 
'#						3. Expected Warning Dialog accessible
'#
'# DESCRIPTION		:	Extebded Multi Application Search window Handeled and verified information in Target Tab.
'#						1. Set TARGET
'#						2.	Validate information
'#						3. Click [OK]
'#						
'# PARAMETERS   	:	bClipBrd = ClipBoard CheckBox in Target Window
'#						sRef = References Target   in Target Window
'#						sPrefSrch = Prior Search  Target   in Target Window
'#						ActWorkflow = Active Workflow  Target   in Target Window
'#						sApplication = Application  Target   in Target Window
'#						sStructure = Dropdown for Structure 
'#						sVerifyInfo = Information to verify relative target window
'#											
'# RETURN VALUE 	: 	TRUE \ FALSE
'#	
'# Examples			:	CALL Fn_SrchExtMultiAppsTrgtListOperation("","","","","Launch Pad","","0 ...On List") 
'#										
'# History			:	Developer Name		Date		Version		Changes Done											Reviewer	
'#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'#						Sagar Shivade	  11-1-11	  	 1.0 		CREATED        			 								Deepak
'#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'# Modified By  	:	Deepak 
'#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'# Modified By  	:	Vivek Ahirrao	  18-05-16		 1.1		Added code for Multiple selection and verification		[TC1122-20160504-18_05_2016-VivekA-NewDevelopment]
'#			Example : Fn_SrchExtMultiAppsTrgtListOperation1("","","Item Revision... (15)~Item Revision... (14)~Item Revision... (13)~Item Revision... (11)","","","","Prior Search:1 ...On List")
'######################################################################################################################################################################################
Public Function Fn_SrchExtMultiAppsTrgtListOperation(bClipBrd,sRef,sPrefSrch,ActWorkflow,sApplication,sStructure,sVerifyInfo)
	GBL_FAILED_FUNCTION_NAME="Fn_SrchExtMultiAppsTrgtListOperation"
	Dim objAdvanceSearch, ObjChild, oDeviceReplay, objSelectType
	Dim aRef, aPrefSrch, aActWorkflow, aApplication, aVerifyInfo
	Dim sAppInfo, sValue, iCount, x_Axis, y_Axis
	Const VK_CONTROL = 29
	'Code tested for Application only
	If Not Fn_UI_ObjectExist("Fn_SrchExtMultiAppsTrgtListOperation",JavaWindow("MyTeamcenter_Search").JavaWindow("SrchDefaultWndw").JavaDialog("Advanced_MultiSite")) Then
		Call Fn_ToolbarButtonClick_Ext(1,"View Menu")
		JavaWindow("DefaultWindow").WinMenu("ContextMenu").Select "Extended Multi-Application Search..."
		Wait 1
	End If
	
	Set objAdvanceSearch = JavaWindow("MyTeamcenter_Search").JavaWindow("SrchDefaultWndw").JavaDialog("Advanced_MultiSite")
	' selecting tab
	objAdvanceSearch.JavaTab("TabSwitch").Select "Target List"
	Wait 1
	' clicking on cler
	objAdvanceSearch.JavaButton("Clear").Click micLeftBtn	
	Wait 1
	'Set value of Clipboard check box
	If bClipBrd<>"" Then
		objAdvanceSearch.JavaCheckBox("Clipboard").Set bClipBrd
		Wait 0,100
	End If
	
	'To set value in Referencers and verify for same
	If sRef<>"" Then
		objAdvanceSearch.JavaButton("Referencers").Click
		Wait 0,500
		sRef = Replace(sRef,"(","\(")
		sRef = Replace(sRef,")","\)")
		'For multiple value selection
		If Instr(sRef,"~")>0 Then
			aRef = Split(sRef,"~")
			'Create Device Replay Object
			Set oDeviceReplay = CreateObject("Mercury.DeviceReplay")
			For iCount = 0 To UBound(aRef)
				Set objSelectType = Description.Create()
				objSelectType("Class Name").value = "JavaStaticText"
				objSelectType("label").value = aRef(iCount)
				Set ObjChild = objAdvanceSearch.ChildObjects(objSelectType)
				If ObjChild.Count > 0 Then
					'Hold Down control key
					oDeviceReplay.KeyDown VK_CONTROL
					'Click on Static Text in list displayed
					x_Axis = ObjChild(0).GetROProperty("abs_x")
					y_Axis = ObjChild(0).GetROProperty("abs_y")
					oDeviceReplay.MouseClick x_Axis+5,y_Axis+5,LEFT_MOUSE_BUTTON
					'Release control key
					oDeviceReplay.KeyUp VK_CONTROL
					Set objSelectType = Nothing
					Set ObjChild = Nothing
				Else
					Set objSelectType = Nothing
					Set ObjChild = Nothing
					Set oDeviceReplay = Nothing
					Set objAdvanceSearch = Nothing
					Fn_SrchExtMultiAppsTrgtListOperation = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to set Referencers value as ["+aRef(iCount)+"]")
					Exit Function
				End If
			Next
			'To set focus on window
			objAdvanceSearch.Click 1,1,"LEFT"
			Set oDeviceReplay = Nothing
		Else
			'For single value selection
			Set objSelectType = Description.Create()
			objSelectType("Class Name").value = "JavaStaticText"
			objSelectType("label").value = sRef
			Set ObjChild = objAdvanceSearch.ChildObjects(objSelectType)
			If ObjChild.Count > 0 Then
				ObjChild(0).Click 1,1,"LEFT"
			Else
				Set objSelectType = Nothing
				Set ObjChild = Nothing
				Set objAdvanceSearch = Nothing
				Fn_SrchExtMultiAppsTrgtListOperation = False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to set Referencers value as ["+sRef+"]")
				Exit Function
			End If
		End If
		Wait 0,500
		'Verify value in field such as "0 ...On List"
		If Instr(sVerifyInfo,"Referencers")>0 Then
			aVerifyInfo = Split(sVerifyInfo,"~")
			For iCount = 0 To UBound(aVerifyInfo)
				If Instr(aVerifyInfo(iCount),"Referencers")>0 Then
					sValue = Replace(aVerifyInfo(iCount),"Referencers:","")
					'set index for info text field and retrive information
					objAdvanceSearch.JavaEdit("InfoTextField").SetTOProperty "Index","1"
					sAppInfo = objAdvanceSearch.JavaEdit("InfoTextField").GetROProperty("value")
					If Trim(Lcase(sAppInfo)) = Trim(Lcase(sValue)) Then
						Fn_SrchExtMultiAppsTrgtListOperation = True
					Else
						Set objAdvanceSearch = Nothing
						Fn_SrchExtMultiAppsTrgtListOperation = False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Verify value of Referencers as ["+sValue+"]")
						Exit Function
					End If
					Exit For
				End If
			Next
		End If
	End If
	
	'To set value in Prior Search and verify for same
	If sPrefSrch<>"" Then
		objAdvanceSearch.JavaButton("PriorSrch").Click
		Wait 0,500
		sPrefSrch = Replace(sPrefSrch,"(","\(")
		sPrefSrch = Replace(sPrefSrch,")","\)")
		'For multiple value selection
		If Instr(sPrefSrch,"~")>0 Then
			aPrefSrch = Split(sPrefSrch,"~")
			'Create Device Replay Object
			Set oDeviceReplay = CreateObject("Mercury.DeviceReplay")
			For iCount = 0 To UBound(aPrefSrch)
				Set objSelectType = Description.Create()
				objSelectType("Class Name").value = "JavaStaticText"
				objSelectType("label").value = aPrefSrch(iCount)
				Set ObjChild = objAdvanceSearch.ChildObjects(objSelectType)
				If ObjChild.Count > 0 Then
					'Hold Down control key
					oDeviceReplay.KeyDown VK_CONTROL
					'Click on Static Text in list displayed
					x_Axis = ObjChild(0).GetROProperty("abs_x")
					y_Axis = ObjChild(0).GetROProperty("abs_y")
					oDeviceReplay.MouseClick x_Axis+5,y_Axis+5,LEFT_MOUSE_BUTTON
					'Release control key
					oDeviceReplay.KeyUp VK_CONTROL
					Set objSelectType = Nothing
					Set ObjChild = Nothing
				Else
					Set objSelectType = Nothing
					Set ObjChild = Nothing
					Set oDeviceReplay = Nothing
					Set objAdvanceSearch = Nothing
					Fn_SrchExtMultiAppsTrgtListOperation = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to set Prior Search value as ["+aPrefSrch(iCount)+"]")
					Exit Function
				End If
			Next
			'To set focus on window
			objAdvanceSearch.Click 1,1,"LEFT"
			Set oDeviceReplay = Nothing
		Else
			'For single value selection
			Set objSelectType = Description.Create()
			objSelectType("Class Name").value = "JavaStaticText"
			objSelectType("label").value = sPrefSrch
			Set ObjChild = objAdvanceSearch.ChildObjects(objSelectType)
			If ObjChild.Count > 0 Then
				ObjChild(0).Click 1,1,"LEFT"
			Else
				Set objSelectType = Nothing
				Set ObjChild = Nothing
				Set objAdvanceSearch = Nothing
				Fn_SrchExtMultiAppsTrgtListOperation = False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to set Prior Search value as ["+sPrefSrch+"]")
				Exit Function
			End If
		End If
		Wait 0,500
		'Verify value in field such as "0 ...On List"
		If Instr(sVerifyInfo,"Prior Search")>0 Then
			aVerifyInfo = Split(sVerifyInfo,"~")
			For iCount = 0 To UBound(aVerifyInfo)
				If Instr(aVerifyInfo(iCount),"Prior Search")>0 Then
					sValue = Replace(aVerifyInfo(iCount),"Prior Search:","")
					'set index for info text field and retrive information
					objAdvanceSearch.JavaEdit("InfoTextField").SetTOProperty "Index","2"
					sAppInfo = objAdvanceSearch.JavaEdit("InfoTextField").GetROProperty("value")
					If Trim(Lcase(sAppInfo)) = Trim(Lcase(sValue)) Then
						Fn_SrchExtMultiAppsTrgtListOperation = True
					Else
						Fn_SrchExtMultiAppsTrgtListOperation = False
						Set objAdvanceSearch = Nothing
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Verify value of Prior Search as ["+sValue+"]")
						Exit Function
					End If
					Exit For
				End If
			Next
		End If
	End If
	
	'To set value in Active Workflows and verify for same
	If ActWorkflow<>"" Then
		objAdvanceSearch.JavaButton("ActiveWorkflows").Click
		Wait 0,500
		ActWorkflow = Replace(ActWorkflow,"(","\(")
		ActWorkflow = Replace(ActWorkflow,")","\)")
		'For multiple value selection
		If Instr(ActWorkflow,"~")>0 Then
			aActWorkflow = Split(ActWorkflow,"~")
			'Create Device Replay Object
			Set oDeviceReplay = CreateObject("Mercury.DeviceReplay")
			For iCount = 0 To UBound(aActWorkflow)
				Set objSelectType = Description.Create()
				objSelectType("Class Name").value = "JavaStaticText"
				objSelectType("label").value = aActWorkflow(iCount)
				Set ObjChild = objAdvanceSearch.ChildObjects(objSelectType)
				If ObjChild.Count > 0 Then
					'Hold Down control key
					oDeviceReplay.KeyDown VK_CONTROL
					'Click on Static Text in list displayed
					x_Axis = ObjChild(0).GetROProperty("abs_x")
					y_Axis = ObjChild(0).GetROProperty("abs_y")
					oDeviceReplay.MouseClick x_Axis+5,y_Axis+5,LEFT_MOUSE_BUTTON
					'Release control key
					oDeviceReplay.KeyUp VK_CONTROL
					Set objSelectType = Nothing
					Set ObjChild = Nothing
				Else
					Set objSelectType = Nothing
					Set ObjChild = Nothing
					Set oDeviceReplay = Nothing
					Set objAdvanceSearch = Nothing
					Fn_SrchExtMultiAppsTrgtListOperation = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to set Active Workflows value as ["+aActWorkflow(iCount)+"]")
					Exit Function
				End If
			Next
			'To set focus on window
			objAdvanceSearch.Click 1,1,"LEFT"
			Set oDeviceReplay = Nothing
		Else
			'For single value selection
			Set objSelectType = Description.Create()
			objSelectType("Class Name").value = "JavaStaticText"
			objSelectType("label").value = ActWorkflow
			Set ObjChild = objAdvanceSearch.ChildObjects(objSelectType)
			If ObjChild.Count > 0 Then
				ObjChild(0).Click 1,1,"LEFT"
			Else
				Set objSelectType = Nothing
				Set ObjChild = Nothing
				Set objAdvanceSearch = Nothing
				Fn_SrchExtMultiAppsTrgtListOperation = False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to set Active Workflows value as ["+ActWorkflow+"]")
				Exit Function
			End If
		End If
		Wait 0,500
		'Verify value in field such as "0 ...On List"
		If Instr(sVerifyInfo,"Active Workflows")>0 Then
			aVerifyInfo = Split(sVerifyInfo,"~")
			For iCount = 0 To UBound(aVerifyInfo)
				If Instr(aVerifyInfo(iCount),"Active Workflows")>0 Then
					sValue = Replace(aVerifyInfo(iCount),"Active Workflows:","")
					Exit For
				End If
			Next
		Else
			sValue = sVerifyInfo
		End If
		'set index for info text field and retrive information
		objAdvanceSearch.JavaEdit("InfoTextField").SetTOProperty "Index","3"
		sAppInfo = objAdvanceSearch.JavaEdit("InfoTextField").GetROProperty("value")
		If Trim(Lcase(sAppInfo)) = Trim(Lcase(sValue)) Then
			Fn_SrchExtMultiAppsTrgtListOperation = True
		Else
			Fn_SrchExtMultiAppsTrgtListOperation = False
			Set objAdvanceSearch = Nothing
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Verify value of Active Workflows as ["+sValue+"]")
			Exit Function
		End If
	End If
	
	'To set value in Applications and verify for same
	If sApplication<>"" Then
		objAdvanceSearch.JavaButton("ActiveWorkflows").Click
		Wait 0,500
		sApplication = Replace(sApplication,"(","\(")
		sApplication = Replace(sApplication,")","\)")
		'For multiple value selection
		If Instr(sApplication,"~")>0 Then
			aApplication = Split(sApplication,"~")
			'Create Device Replay Object
			Set oDeviceReplay = CreateObject("Mercury.DeviceReplay")
			For iCount = 0 To UBound(aApplication)
				Set objSelectType = Description.Create()
				objSelectType("Class Name").value = "JavaStaticText"
				objSelectType("label").value = aApplication(iCount)
				Set ObjChild = objAdvanceSearch.ChildObjects(objSelectType)
				If ObjChild.Count > 0 Then
					'Hold Down control key
					oDeviceReplay.KeyDown VK_CONTROL
					'Click on Static Text in list displayed
					x_Axis = ObjChild(0).GetROProperty("abs_x")
					y_Axis = ObjChild(0).GetROProperty("abs_y")
					oDeviceReplay.MouseClick x_Axis+5,y_Axis+5,LEFT_MOUSE_BUTTON
					'Release control key
					oDeviceReplay.KeyUp VK_CONTROL
					Set objSelectType = Nothing
					Set ObjChild = Nothing
				Else
					Set objSelectType = Nothing
					Set ObjChild = Nothing
					Set oDeviceReplay = Nothing
					Set objAdvanceSearch = Nothing
					Fn_SrchExtMultiAppsTrgtListOperation = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to set Applications value as ["+aApplication(iCount)+"]")
					Exit Function
				End If
			Next
			'To set focus on window
			objAdvanceSearch.Click 1,1,"LEFT"
			Set oDeviceReplay = Nothing
		Else
			'For single value selection
			Set objSelectType = Description.Create()
			objSelectType("Class Name").value = "JavaStaticText"
			objSelectType("label").value = sApplication
			Set ObjChild = objAdvanceSearch.ChildObjects(objSelectType)
			If ObjChild.Count > 0 Then
				ObjChild(0).Click 1,1,"LEFT"
			Else
				Set objSelectType = Nothing
				Set ObjChild = Nothing
				Set objAdvanceSearch = Nothing
				Fn_SrchExtMultiAppsTrgtListOperation = False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to set Applications value as ["+sApplication+"]")
				Exit Function
			End If
		End If
		Wait 0,500
		'Verify value in field such as "0 ...On List"
		If Instr(sVerifyInfo,"Active Workflows")>0 Then
			aVerifyInfo = Split(sVerifyInfo,"~")
			For iCount = 0 To UBound(aVerifyInfo)
				If Instr(aVerifyInfo(iCount),"Active Workflows")>0 Then
					sValue = Replace(aVerifyInfo(iCount),"Active Workflows:","")
					Exit For
				End If
			Next
		Else
			sValue = sVerifyInfo
		End If
		'set index for info text field and retrive information
		objAdvanceSearch.JavaEdit("InfoTextField").SetTOProperty "Index","4"
		sAppInfo = objAdvanceSearch.JavaEdit("InfoTextField").GetROProperty("value")
		If Trim(Lcase(sAppInfo)) = Trim(Lcase(sValue)) Then
			Fn_SrchExtMultiAppsTrgtListOperation = True
		Else
			Fn_SrchExtMultiAppsTrgtListOperation = False
			Set objAdvanceSearch = Nothing
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Verify value of Applications as ["+sValue+"]")
			Exit Function
		End If
	End If
	
	'Set Structure Manager check box value
	If sStructure<>"" Then
		objAdvanceSearch.JavaCheckBox("Structure Manager").Set(sStructure)
		wait(1)
		Call Fn_Button_Click("Fn_SrchExtMultiAppsTrgtListOperation",objAdvanceSearch, "SMListButton")
		If JavaWindow("MyTeamcenter_Search").JavaWindow("SrchDefaultWndw").JavaDialog("Collect BomElements for").Exist Then
			Call Fn_Button_Click("Fn_SrchExtMultiAppsTrgtListOperation", JavaWindow("MyTeamcenter_Search").JavaWindow("SrchDefaultWndw").JavaDialog("Collect BomElements for"), "GO")
			wait(3)
		Else 
			Fn_SrchExtMultiAppsTrgtListOperation = False
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Collect BomElements not Exist" )
		End If
		'set index for info text field and retrive information
		objAdvanceSearch.JavaEdit("InfoTextField").SetTOProperty "Index","5"
		wait(3)
		sAppInfo = objAdvanceSearch.JavaEdit("InfoTextField").GetROProperty("value")
		wait(3)
		If Trim (Lcase (sAppInfo)) = Trim (Lcase (sVerifyInfo)) Then
			Fn_SrchExtMultiAppsTrgtListOperation = True
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Successfully set Target and verified information from Structure Manager")
		Else
			Fn_SrchExtMultiAppsTrgtListOperation = False
			Set objAdvanceSearch = Nothing
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to  set Target and verified information from Structure Manager")	
			Exit Function
		End If
	End If

	'Click on OK Button
	Call Fn_Button_Click("Fn_SrchExtMultiAppsTrgtListOperation", objAdvanceSearch, "OK")
	Fn_SrchExtMultiAppsTrgtListOperation = True
	Set ObjChild = Nothing
	Set objAdvanceSearch = Nothing
End Function

'##########################################################################################################################################################~
'# FUNCTION NAME:         Fn_MyTc_ChangeSearchHistoryOperation
'#
'#
'# MODULE:             Search Requirement
'#               
'# DESCRIPTION:            This Operate on Search History
'#                           
'# PARAMETERS   :             
'#                                                         
'# RETURN VALUE :            TRUE \ FALSE
'#
'# Examples    :        Fn_MyTc_ChangeSearchHistoryOperation("Validate","Item Revision... (1):Item... (1)" ,"Change search from search history")   
'#                                       
'# History    :        Developer Name                            Date                    Rev. No.            Changes Done            Reviewer   
'#------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'#                            Pranav Ingle                            22-Aug-2011                001
'#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'###########################################################################################################################################################~


Public Function Fn_MyTc_ChangeSearchHistoryOperation(strAction, sSearchHistoryItems,strButton)
GBL_FAILED_FUNCTION_NAME="Fn_MyTc_ChangeSearchHistoryOperation"
Dim iCounter, bFlag,objSelectType,intNoOfObjects,iCount,arrSearchHistoryItems

                arrSearchHistoryItems = Split(sSearchHistoryItems,":")
                '1. Operate on the Main Menu: Window;Show View;Other...
                    Call Fn_MenuOperation("Select","Window:Show View:Search")

                '2. Change Search History toolbat button click
                    Call Fn_ToolbatButtonClick(strButton)

    Select Case StrAction

        '----------------------------------------------------------------------- For Select Searches from Searches History-------------------------------------------------------------------------
        Case "Select"   'Implemented by Pritam S.

				bFlag = False
				Set objSelectType=Description.Create()
				objSelectType("Class Name").value = "JavaMenu"
				Set  intNoOfObjects =JavaWindow("MyTeamcenter_Search").ChildObjects(objSelectType)
				For iCounter = 0 to intNoOfObjects.count-1
					If intNoOfObjects(iCounter).getROProperty("label") = arrSearchHistoryItems(0)  Then
						intNoOfObjects(iCounter).select
						bFlag = True
						Exit For
					End If
				Next
				 If bFlag = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile")," Failed To  select search ["+arrSearchHistoryItems(iCount)+"] " )
					Fn_MyTc_ChangeSearchHistoryOperation = False 
					Set intNoOfObjects=Nothing
                    Set objSelectType=Nothing  
					Exit Function
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile")," Successfully select search ["+arrSearchHistoryItems(iCount)+"] " )
					Fn_MyTc_ChangeSearchHistoryOperation = True   
				End If   
				Set intNoOfObjects=Nothing
                Set objSelectType=Nothing
        '----------------------------------------------------------------------- For Validating Entries in Search History -------------------------------------------------------------------------
        Case "Validate"   

                    For iCount = 0 to Ubound(arrSearchHistoryItems)
                        bFlag = False
                        Set objSelectType=Description.Create()
                        objSelectType("Class Name").value = "JavaMenu"
                        Set  intNoOfObjects =JavaWindow("MyTeamcenter_Search").ChildObjects(objSelectType)
                   
                        For  iCounter = 0 to intNoOfObjects.count-1
                            If intNoOfObjects(iCounter).getROProperty("label") = arrSearchHistoryItems(iCount) Then
                                    bFlag = True                                    
                                    Exit For
                            End If
                        Next
                        If bFlag = False Then
                                Call Fn_WriteLogFile(Environment.Value("TestLogFile")," Failed To  Verify saved searches ["+arrSearchHistoryItems(iCount)+"] " )
                                Fn_MyTc_ChangeSearchHistoryOperation = False   
								Exit Function
                        End If   
                    Next
                    Fn_MyTc_ChangeSearchHistoryOperation = True
                    Set intNoOfObjects=Nothing
                    Set objSelectType=Nothing

        Case Else
                        Fn_MyTc_ChangeSearchHistoryOperation = False
    End Select
End Function

'#######################################################################################################################################################~
'#######################		Sort And Validate the column are Sorted under PFF Search Option Table in My Teacenter Application Search Results Pane			######################~
'#
'# FUNCTION NAME:	Fn_MyTc_SrchPFFTableColSortOperation
'#
'# MODULE: 						 My Teamcenter
'#
'# PRE-REQUISITE:	    RAC Session accessible and My Teacenter Application Search Results Pane loaded. 
'#											Set the requisite PFF Search Option in My Teacenter Application Search Results Pane
'# 				
'# DESCRIPTION:			Validate the existance of columns under PFF Search Option Table in My Teacenter Application Search Results Pane
'#
'#											1. Sort the column under PFF Options Table
'#											2. Validate column under PFF Options Table are Sorted
'#										
'#PARAMETERS   :     	strAction : Action To Perform
'#  										aPFFOptionColumns: array of ":" seperated PFF column names							
'#
'# RETURN VALUE : 		TRUE \ FALSE
'#
'#Examples	:				 	Fn_MyTc_SrchPFFTableColSortOperation("Sort","Object Name")
'#											Fn_MyTc_SrchPFFTableColSortOperation("Validate Sort Ascending","Object Name")
'#											Fn_MyTc_SrchPFFTableColSortOperation("Validate Sort Descending","Object Name")
'#										
'#	History	:						Developer Name			Date						Version						Changes Done			Reviewer	
'#------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'#										Pranav Ingle				26-Aug -2011				1.0																		
'#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'########################################################################################################################################################~
'#######################		Validate the existance of columns under PFF Search Option Table in My Teacenter Application Search Results Pane			######################~

Public Function Fn_MyTc_SrchPFFTableColSortOperation(strAction,aPFFOptionColumns)
	GBL_FAILED_FUNCTION_NAME="Fn_MyTc_SrchPFFTableColSortOperation"
Dim iCols,iRows, aArrCols, iOuterCount, iCount, iCounter, sColName, bFlag
Dim ObjJavaTbl

			
			Set ObjJavaTbl = JavaWindow("MyTeamcenter_Search").JavaTable("PffTable")

          Select Case strAction
	  			Case "Sort"
							
								Call Fn_ToolbatButtonClick("Refresh property formetter search") 
								wait(1)

								Call Fn_ReadyStatusSync("2")
								iCols = Fn_UI_Object_GetROProperty("Fn_MyTc_SrchPFFTableColSortOperation",ObjJavaTbl, "cols")
                            					
								bFlag = False
								For iCount = 0 To iCols-1
											sColName =  ObjJavaTbl.GetColumnName(iCount)
											If sColName = aPFFOptionColumns Then
													ObjJavaTbl.SelectColumnHeader aPFFOptionColumns
													bFlag = true
													Exit For
											End If
									Next
								If  bFlag = True  Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Pass: Requested Column is present in PFF Table. ")   	
									Fn_MyTc_SrchPFFTableColSortOperation = True
								Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Fail: Requested Column :["+aPFFOptionColumns+"] Does Not Exist in PFF Table. ")   	
									Fn_MyTc_SrchPFFTableColSortOperation = False
								End If 
								
					Case "Validate Sort Ascending"
								iCols = Fn_UI_Object_GetROProperty("Fn_MyTc_SrchPFFTableColSortOperation",ObjJavaTbl, "cols")
								iRows = Fn_UI_Object_GetROProperty("Fn_MyTc_SrchPFFTableColSortOperation",ObjJavaTbl, "rows")
                    					
								bFlag = False
								For iCount = 0 To iCols-1
											sColName =  ObjJavaTbl.GetColumnName(iCount)
											If sColName = aPFFOptionColumns Then
                                                	bFlag = true
													Exit For
											End If
								Next
								If  bFlag = True  Then
										For iCount = 0 To iRows-2
												iCounter = iCount +1
												If  ObjJavaTbl.GetCellData(iCount,aPFFOptionColumns) > ObjJavaTbl.GetCellData(iCounter,aPFFOptionColumns)  Then
														Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Fail: Requested Column in PFF Table is Not sorted")   	
                                                        Fn_MyTc_SrchPFFTableColSortOperation = False
														Exit Function
												End If
										Next
								Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Fail: Requested Column :["+aPFFOptionColumns+"] Does Not Exist in PFF Table. ")   	
									Fn_MyTc_SrchPFFTableColSortOperation = False
									Exit Function
								End If 

								Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Pass: Requested Column in PFF Table is sorted")   	
								Fn_MyTc_SrchPFFTableColSortOperation = True

					Case "Validate Sort Descending"
								iCols = Fn_UI_Object_GetROProperty("Fn_MyTc_SrchPFFTableColSortOperation",ObjJavaTbl, "cols")
								iRows = Fn_UI_Object_GetROProperty("Fn_MyTc_SrchPFFTableColSortOperation",ObjJavaTbl, "rows")
                    					
								bFlag = False
								For iCount = 0 To iCols-1
											sColName =  ObjJavaTbl.GetColumnName(iCount)
											If sColName = aPFFOptionColumns Then
                                                	bFlag = true
													Exit For
											End If
								Next
								If  bFlag = True  Then
										For iCount = 0 To iRows-2
												iCounter = iCount +1
												If  ObjJavaTbl.GetCellData(iCount,aPFFOptionColumns) < ObjJavaTbl.GetCellData(iCounter,aPFFOptionColumns)  Then
														Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Fail: Requested Column in PFF Table is Not sorted")   	
                                                        Fn_MyTc_SrchPFFTableColSortOperation = False
														Exit Function
												End If
										Next
                                Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Fail: Requested Column :["+aPFFOptionColumns+"] Does Not Exist in PFF Table. ")   	
									Fn_MyTc_SrchPFFTableColSortOperation = False
									Exit Function
								End If 

								Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Pass: Requested Column in PFF Table is sorted")   	
								Fn_MyTc_SrchPFFTableColSortOperation = True

		  End Select
            
Set ObjJavaTbl  =  Nothing
End Function

'########################################################################################################
'#
'# FUNCTION NAME:	 	  Fn_QryBldr_VerifyDetails
'#
'# RETURN VALUE : 		TRUE \ FALSE
'#
'#	Examples	:		Call Fn_QryBldr_VerifyDetails("Admin - Employee Information_49465", "New Description for New Query", "", "", "", "", "", "user_id:UserId:User ID:=:", "")	 
'#
'#	History	:						Developer Name			Date			Version				Changes Done				
'#-------------------------------------------------------------------------------------------------------------------------------------
'#										Pranav S.		22-03-12			  1.0									
'#-------------------------------------------------------------------------------------------------------------------------------------
'########################################################################################################
Public Function Fn_QryBldr_VerifyDetails(sQueryName, sQueryDescription, sQueryTypes, sSearchClass, sDisplaySettings, bShowIndentedResults, aAttributes, sSearchCriteria, sbuttons)  
  	GBL_FAILED_FUNCTION_NAME="Fn_QryBldr_VerifyDetails"
	Dim iCounter, ObjQryApp, RmvArrAttrib, iCnt, iRows, strData, iCols, iCount, aButtons, dicDetailQuery
	
	Set ObjQryApp = Fn_UI_ObjectCreate("Fn_QryBldr_VerifyDetails",JavaWindow("Search_QueryBuilder").JavaWindow("QueryBuilder"))
	
	Fn_QryBldr_VerifyDetails=True
	'++++++++++<<    Input [Name] field details >>++++++++++
	If sQueryName <> "" Then
		If Fn_Edit_Box_GetValue("Fn_QryBldr_VerifyDetails",ObjQryApp,"Name") = sQueryName Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "sQueryName value matches with actual value")
			Fn_QryBldr_VerifyDetails=True
		Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "sQueryName value dose'nt matches with actual value")
			Fn_QryBldr_VerifyDetails=False
			Exit Function
		End If
	End If

	'++++++++++<<   Input [Description] field details>>++++++++++
	If sQueryDescription <>"" Then
								
		If Fn_Edit_Box_GetValue("Fn_QryBldr_VerifyDetails",ObjQryApp,"Description") = sQueryDescription Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Description value matches with actual value")
			Fn_QryBldr_VerifyDetails=True
		Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Description value dose'nt matches with actual value")
			Fn_QryBldr_VerifyDetails = False
			Exit Function
		End If
	End If

	'++++++++++<<   Query Typres  field details>>++++++++++
	If sQueryTypes <>"" Then
		If Fn_Edit_Box_GetValue("Fn_QryBldr_VerifyDetails",ObjQryApp,"QueryTypes") = sQueryTypes Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "QueryTypes value matches with actual value")
			Fn_QryBldr_VerifyDetails=True
		Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "QueryTypes value dose'nt matches with actual value")
			Fn_QryBldr_VerifyDetails = False
			Exit Function
		End If
	End If

	'++++++++++<<   Verify [Search Class] values >>++++++++++'------Add as per need
	
	'[TC1122-20160504-25_05_2016-VivekA-NewDevelopment] - Added for Search new TCs
	'Check whether sDisplaySettings is Dictionary object or not
	If varType(sDisplaySettings) = "9" Then
		Set dicDetailQuery = sDisplaySettings
		
		'Verify Revision Rule
		If dicDetailQuery("Revision Rule")<>"" Then
			If Fn_SISW_UI_JavaList_Operations("Fn_QryBldr_VerifyDetails","GetText",ObjQryApp,"RevisionRule","","","") = dicDetailQuery("Revision Rule") Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Revision Rule value matches with actual value")
				Fn_QryBldr_VerifyDetails = True
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Revision Rule value does'nt matches with actual value")
				Fn_QryBldr_VerifyDetails = False
				Exit Function
			End If
		End If
		
		'Verify Display Settings
			' Future as per need
	End If
	'----------------------------------------------------

	'++++++++++<<  Verify [Display Settings] values >>++++++++++------Add as per need
				
	'++++++++++<<  Set [Show Indented Results] option [On/Off] >>++++++++++------Add as per need
	 
	'++++++++++<<   Select attribute from [Attribute Selection] Tree >>++++++++++---Add as per need

	'++++ Verify the Search Criteria++++
	If sSearchCriteria <> "" then
		RmvArrAttrib = split(sSearchCriteria, ":", -1,1)
		iRows = JavaWindow("Search_QueryBuilder").JavaWindow("QueryBuilder").JavaTable("SrchCriteriaTable").GetROProperty("rows")	
		iCols = JavaWindow("Search_QueryBuilder").JavaWindow("QueryBuilder").JavaTable("SrchCriteriaTable").GetROProperty("cols")
		For iCounter=0 to Ubound(RmvArrAttrib)
			For iCnt=0 to iRows - 1
				For iCount=0 to iCols - 1
                    	If RmvArrAttrib(iCounter) = Trim(JavaWindow("Search_QueryBuilder").JavaWindow("QueryBuilder").JavaTable("SrchCriteriaTable").GetCellData(iCnt,iCount)) Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Search value matches with actual value")
						Fn_QryBldr_VerifyDetails=True
						Exit For
					End If
				Next
			Next
		Next

	End If

	'++++++++++<<  Verify buttons >>++++++++++
	If sButtons<>"" Then
		aButtons = split(sButtons, ":",-1,1)
		iCounter = Ubound(aButtons)
		For iCount=0 to iCounter
			If ObjQryApp.JavaButton(aButtons(iCount)).Exist Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Button Exist")
				Fn_QryBldr_VerifyDetails=True
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Button does'nt Exist")
				Fn_QryBldr_VerifyDetails=False
				Exit Function
			End If
		Next
	End If


Set ObjQryApp = Nothing
End Function
'########################################################################################################
'#
'# FUNCTION NAME:	 	 Fn_SISW_QryBldr_OrderByOperation
'#
'# RETURN VALUE : 		    TRUE \ FALSE
'#
'#	Examples	:		bReturn = Fn_SISW_QryBldr_OrderByOperation("Modify","","","","","OFF","ItemRevision:Sequence Limit","","","")
'#
'#                                  aValue = Array("","sequence_limit","Sequence Limit","")
'#                                  bReturn = Fn_SISW_QryBldr_OrderByOperation("Verify","","","","","","","","1",aValue) 
'#
'#									bReturn = Fn_SISW_QryBldr_OrderByOperation("SetTableDataWithoutModify","","","","","","","",1,aValues)

'#                                  bReturn = Fn_SISW_QryBldr_OrderByOperation("SetTableDataWithModify","","","","","","","",1,aValues)

'#                                  bReturn = Fn_SISW_QryBldr_OrderByOperation("SelectRow","","","","","","","",2,"")
'#
'#
'#
'#	History	:						Developer Name			Date			Version				Changes Done				
'#-------------------------------------------------------------------------------------------------------------------------------------
'#										   Pritam Shikare    		12-12-12			  1.0									
'#-------------------------------------------------------------------------------------------------------------------------------------
'########################################################################################################
Public Function  Fn_SISW_QryBldr_OrderByOperation(sAction, sQueryName, sQueryDescription, sSearchClass, sDisplaySettings, bShowIndentedResults, aAttributes, sRemAttribNmIndex, sRowIndex,aValues)  
  	GBL_FAILED_FUNCTION_NAME="Fn_SISW_QryBldr_OrderByOperation"
	Dim  ArrDispSet,  OuterArrAttrib, iOuterCounter, ArrAttrib,  ArrinnAttrib, iCounter
	Dim ObjQryApp, ObjQryAttribSel,objOrderByTable
	Dim sValue, bMatched,objApplet
	
	If isNumeric(sRowIndex) Then 
	sRowIndex = cInt(sRowIndex)
	End If

	If isNumeric(sRemAttribNmIndex) Then 
	sRemAttribNmIndex = cInt(sRemAttribNmIndex)
	End If
	
	Set ObjQryApp =  Fn_UI_ObjectCreate( "Fn_SISW_QryBldr_OrderByOperation",JavaWindow("Search_QueryBuilder").JavaWindow("QueryBuilder"))
	Set objOrderByTable = Fn_SISW_Search_GetObject("SearchCriteriaTab")
    Set objApplet = Fn_SISW_Search_GetObject("Project_Srch")'Added by Priyanka on 20-Dec-2012
   Select Case sAction

				Case "Modify" ,"ModifyWithoutButton" 'Added Case "ModifyWithoutButton" by Priyanka on 20-Dec-2012
							'++++++++++<<    Input [Name] field details >>++++++++++
							If sQueryName <> "" Then
								Call Fn_Edit_Box("Fn_SISW_QryBldr_OrderByOperation",ObjQryApp,"Name",sQueryName)
							End If
						
							'++++++++++<<   Input [Description] field details>>++++++++++
							If sQueryDescription <>"" Then
								Call Fn_Edit_Box("Fn_SISW_QryBldr_OrderByOperation",ObjQryApp,"Description",sQueryDescription)
							End If
						
							'++++++++++<<   Click [Search Class] button >>++++++++++
							If  sSearchClass <> "" Then
								Call Fn_CheckBox_Set("Fn_SISW_QryBldr_OrderByOperation", ObjQryApp, "SrchClass", "ON")
								Call Fn_Edit_Box("Fn_SISW_QryBldr_OrderByOperation",ObjQryApp,"Class/Attribute Selection",sSearchClass)
								Call Fn_Button_Click( "Fn_SISW_QryBldr_OrderByOperation", ObjQryApp, "Find")
								ObjQryApp.JavaObject("Close").Click 1,1
							End If
						
							 '++++++++++<<  Set the [Display Setting] to required option [Class/All Attributes] >>++++++++++
							 If sDisplaySettings <> ""  Then
								 ArrDispSet = split(sDisplaySettings, ":", -1,1)
								 If Ubound(ArrDispSet) = 1 Then
									Call Fn_CheckBox_Set("Fn_SISW_QryBldr_OrderByOperation", ObjQryApp, "DisplaySettings", "ON")
									Call  Fn_UI_JavaRadioButton_SetON("Fn_SISW_QryBldr_OrderByOperation",ObjQryApp, ArrDispSet(0))
									Call  Fn_UI_JavaRadioButton_SetON("Fn_SISW_QryBldr_OrderByOperation",ObjQryApp, ArrDispSet(1))
								Else
									Call Fn_CheckBox_Set("Fn_SISW_QryBldr_OrderByOperation", ObjQryApp, "DisplaySettings", "ON")
									Call  Fn_UI_JavaRadioButton_SetON("Fn_SISW_QryBldr_OrderByOperation",ObjQryApp, sDisplaySettings )
								 End If
								ObjQryApp.JavaObject("Close").Click 1,1
							 End If
						
							 '++++++++++<<  Set [Show Indented Results] option [On/Off] >>++++++++++
							 If bShowIndentedResults <> ""  Then
								Call Fn_CheckBox_Set("Fn_SISW_QryBldr_OrderByOperation", ObjQryApp, "ShowIndentedResults", bShowIndentedResults)
							 End If
			
						   objOrderByTable.Select "Order By"
			
							
							'++++++++++<<   Select attribute from [Attribute Selection] Tree >>++++++++++
							If  aAttributes <> "" Then
								OuterArrAttrib = split(aAttributes, "~", -1,1)
								For iOuterCounter = 0 To Ubound(OuterArrAttrib)
										ArrAttrib = split(OuterArrAttrib(iOuterCounter), ",", -1, 1)				
										For iCounter = 0 to Ubound(ArrAttrib) 		
													ArrinnAttrib = split(ArrAttrib(iCounter), ":", -1, 1)
													Select Case iCounter
													Case "0"
															Wait 2
															ObjQryApp.JavaTree("AttributeSelectionList").Activate ArrAttrib(iCounter) 
													Case "1"
															If  True = Fn_UI_ObjectExist("Fn_SISW_QryBldr_OrderByOperation", ObjQryApp.JavaDialog("ClassAttributeSelection") )Then
															
																Set ObjQryAttribSel = ObjQryApp.JavaDialog("ClassAttributeSelection") 
																Call Fn_CheckBox_Set("Fn_SISW_QryBldr_OrderByOperation", ObjQryAttribSel, "CAS_SrchClass", "ON")
																Call Fn_Edit_Box("Fn_SISW_QryBldr_OrderByOperation",ObjQryAttribSel,"CAS_Edit",ArrinnAttrib(0) )
																Call Fn_Button_Click( "Fn_SISW_QryBldr_OrderByOperation", ObjQryApp, "Find")
																ObjQryApp.JavaObject("Close").Click 1,1
																Wait 2
																ObjQryAttribSel.JavaTree("CAS_SrchTree").Activate ArrAttrib(iCounter) 
															End If
															If  True = Fn_UI_ObjectExist("Fn_SISW_QryBldr_OrderByOperation",ObjQryApp.JavaDialog("ClassSelectionDialog")  )Then
																Set ObjQryAttribSel = ObjQryApp.JavaDialog("ClassSelectionDialog")
																Wait 2
																Call Fn_Edit_Box("Fn_SISW_QryBldr_OrderByOperation",ObjQryAttribSel,"SelectionField",ArrAttrib(iCounter) )
																Call Fn_Button_Click( "Fn_SISW_QryBldr_OrderByOperation", ObjQryAttribSel, "CSDFind")
																Call Fn_Button_Click( "Fn_SISW_QryBldr_OrderByOperation", ObjQryAttribSel, "CSDOK")
															End If										               
													Case "2"
															Wait 2
															ObjQryApp.JavaTree("AttributeSelectionList").Activate ArrAttrib(iCounter) 
													End Select		
										Next			
								Next
							End If
						
						   '++++++++++<<   Remove the Attribute Specified. >>++++++++++
							If sRemAttribNmIndex <> "" Then
								 If True = Fn_UI_JavaTable_SelectRow("Fn_SISW_QryBldr_OrderByOperation", ObjQryApp, "OrderByTable",sRemAttribNmIndex) Then
									 If True = Fn_Button_Click( "Fn_SISW_QryBldr_OrderByOperation", objApplet, "Remove") Then  'Added by Priyanka on 20-Dec-2012
										 Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Pass : Local Query Search Criteria Deleted. ") 
										 Fn_SISW_QryBldr_OrderByOperation = True  	
									End If
								Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Fail : Local Query Search Criteria Not Deleted. ")   	
									Fn_SISW_QryBldr_OrderByOperation = False
									Exit Function
								End If
							End If 
							Fn_SISW_QryBldr_OrderByOperation = True  	     			
							'++++++++++<<  Invoke [Modify] button >>++++++++++
							If  sAction = "Modify" Then
								 If True =  Fn_Button_Click( "Fn_SISW_QryBldr_OrderByOperation", ObjQryApp, "Modify") Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Pass: Local Query Successfully Modified. ")   	
										Fn_SISW_QryBldr_OrderByOperation = True
								Else
										Call Fn_Button_Click( "Fn_SISW_QryBldr_OrderByOperation", ObjQryApp, "Clear")
										Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Fail: Local Query Failed to be Modified. ")   	
										Fn_SISW_QryBldr_OrderByOperation = False

								End If
							End If
				'##################CASE : Verify    to Verify the Values in the Table  ##########################################################################################
					Case "Verify"
							'Set ObjSrchCrtr =  Fn_UI_ObjectCreate( "Fn_QryBldr_SearchCriteriaValidate",JavaWindow("Search_QueryBuilder").JavaWindow("QueryBuilder") )
							bMatched = False		
							objOrderByTable.Select "Order By"
							For iCounter = 0 To UBound(aValues)
									If  aValues(iCounter) <> "" Then
										sValue =  Fn_UI_JavaTable_GetCellData("Fn_SISW_QryBldr_OrderByOperation",ObjQryApp,"OrderByTable",sRowIndex, iCounter)
										If sValue = aValues(iCounter) Then
											bMatched = True
										Else
											bMatched = False
											Exit For
										End If
									End If
							Next
								
							If  bMatched = True Then	
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Validation for Requested fields successfully Verified." )
									Fn_SISW_QryBldr_OrderByOperation = True
							Else 
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed: Validation for Requested fields Failed ." )
									Fn_SISW_QryBldr_OrderByOperation = False
							End If

					Case "SetTableDataWithoutModify", "SetTableDataWithModify"
							'Set ObjSrchCrtr =  Fn_UI_ObjectCreate( "Fn_QryBldr_SearchCriteriaValidate",JavaWindow("Search_QueryBuilder").JavaWindow("QueryBuilder") )
							bMatched = False		
							objOrderByTable.Select "Order By"
							For iCounter = 0 To UBound(aValues)
									If  aValues(iCounter) <> "NA" Then
										bReturn = Fn_UI_JavaTable_SetCellData("Fn_SISW_QryBldr_OrderByOperation",ObjQryApp,"OrderByTable",sRowIndex, iCounter,aValues(iCounter))
										If bReturn = False Then
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed: Validation for Requested fields Failed ." )
											Fn_SISW_QryBldr_OrderByOperation = False
											Exit For
									   Else
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Validation for Requested fields successfully Verified." )
											Fn_SISW_QryBldr_OrderByOperation = True
										End If
									End If
							Next

							If sAction = "SetTableDataWithModify" Then
								If True =  Fn_Button_Click( "Fn_SISW_QryBldr_OrderByOperation", ObjQryApp, "Modify") Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Pass: Table Data Successfully Modified. ")   	
									Fn_SISW_QryBldr_OrderByOperation = True
								Else
									Call Fn_Button_Click( "Fn_SISW_QryBldr_OrderByOperation", ObjQryApp, "Clear")
									Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Fail: Table Data Failed to be Modified. ")   	
									Fn_SISW_QryBldr_OrderByOperation = False
								End If
							End If

						Case "SelectRow"

									bReturn = Fn_UI_JavaTable_SelectRow("Fn_SISW_QryBldr_OrderByOperation", ObjQryApp, "SrchCriteriaTable",sRowIndex)
									If bReturn = False Then
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed: Failed to select a Row ." )
											Fn_SISW_QryBldr_OrderByOperation = False
									   Else
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully selected a row." )
											Fn_SISW_QryBldr_OrderByOperation = True
									End If

		End Select
						
		Set ObjQryApp = Nothing
		Set ObjQryAttribSel = Nothing
		Set objOrderByTable = Nothing
		Set objApplet =Nothing
End Function

'###########################################	Function Used to Verify Error Message    #####################################################~
'#
'# FUNCTION NAME:	 	  Fn_MyTcSrch_ErrorMessageVerify
'#
'# MODULE: 						 My Teamcenter Search
'#
'# PRE-REQUISITE:	   1. RAC Session accessible and Query Builder Perspective loaded
'#										2. Couple of fields which are not part of search criteria are selected in Order By tab
'#										3. Expected Error dialog is opened
'#
'# DESCRIPTION:			 Function Used to Verify Error Message
'#
'#										
'#						
'# PARAMETERS   :     1. strDialogName: Name of the Error dialog	
'#										2. strErrorMsg : Expected Error Message
'#										3. sCheckBox : CheckBox Name
'#										4. sAction : Action Name
'#										5. strButtton: Button Name
'#
'#
'#
'# RETURN VALUE : 	TRUE \ FALSE
'#
'#Examples	:				  bReturn = Fn_MyTcSrch_ErrorMessageVerify("Sort Attribute Count Exceeded","No more than 3 sort attributes can be defined.","","ErrorMessageVerify","")
'#										 bReturn = Fn_MyTcSrch_ErrorMessageVerify("Sort Attribute Count Exceeded","No more than 3 attributes can be added to the Order By table of a Saved Query.","More...","DetailMessageVerify","OK")
'#
'#
'#										
'#	History	:					  Developer Name			Date			Version				Changes Done			Reviewer	
'#------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'#										Priyanka Bhave			20/12/2012		1.0																Nilesh Gadekar
'#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'########################################################################################################################################################~
Public Function Fn_MyTcSrch_ErrorMessageVerify(strDialogName,strErrorMsg,sCheckBox,sAction,strButtton)
	GBL_FAILED_FUNCTION_NAME="Fn_MyTcSrch_ErrorMessageVerify"
	GBL_EXPECTED_MESSAGE=strErrorMsg
   'Variable Declaration
   Dim strMsg,objErrDialog
   Set objErrDialog= Window("ProjectSearch").JavaWindow("WEmbeddedFrame").JavaDialog("Error")
   Fn_MyTcSrch_ErrorMessageVerify=False
   'Setting Dialog Title
   Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_MyTcSrch_ErrorMessageVerify",objErrDialog,"title",strDialogName)

   Select Case sAction
 	Case "ErrorMessageVerify"
		'Checking Existance Of Error Dialog
		 If Fn_UI_ObjectExist("Fn_MyTcSrch_ErrorMessageVerify", objErrDialog)=True Then
			'Retriving Error Message which Apprers on Dialog
			 strMsg=Fn_UI_Object_GetROProperty("Fn_MyTcSrch_ErrorMessageVerify",objErrDialog.JavaEdit("Msg"),"value")
			 'Verifying Error Message Match With Expected Error Message
			 If InStr(strMsg,strErrorMsg)>0 Then
			  'Function Returns true
			  	Fn_MyTcSrch_ErrorMessageVerify=True
			  Else
			  	GBL_ACTUAL_MESSAGE=strMsg
			End If
		 End If   

	Case "DetailMessageVerify"
			'Checking Existance Of Error Dialog
		 If Fn_UI_ObjectExist("Fn_MyTcSrch_ErrorMessageVerify", objErrDialog)=True Then
		 If sCheckBox <> "" Then
			 Call Fn_CheckBox_Select("Fn_MyTcSrch_ErrorMessageVerify", objErrDialog, sCheckBox)
		 End If	
			'Retriving Error Message which Apprers on Dialog
			 strMsg=Fn_UI_Object_GetROProperty("Fn_MyTcSrch_ErrorMessageVerify",objErrDialog.JavaEdit("DetailMsg"),"value")
    		 'Verifying Error Message Match With Expected Error Message
			 If InStr(strMsg,strErrorMsg)>0 Then
			  'Function Returns true
			  	Fn_MyTcSrch_ErrorMessageVerify=True
			  Else 
			  	GBL_ACTUAL_MESSAGE=strMsg
			End If
		 End If   
   End Select
   If  strButtton <> "" Then
		'Clicking On strButtton Button
		Call Fn_Button_Click("Fn_MyTcSrch_ErrorMessageVerify", objErrDialog, strButtton)
	End If
	Set objErrDialog=Nothing
End Function 
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'#
'# FUNCTION NAME:	 	  Fn_QryBldr_ImportQuery
'#
'#  MyCommunity ID :	297
'#
'# MODULE: 						 Search Requirement, MyTC 
'#
'# PRE-REQUISITE:		RAC Session accessible and [Query Builder] Application loaded
'#
'# DESCRIPTION:			  Import a existing Query in Query Builder Application
'#										
'# PARAMETERS   :        sFilePath: Source path of Import file format
'#											sNewQueryName: New query name
'#
'# RETURN VALUE : 		TRUE \ FALSE
'#
'#Examples	:		Fn_QryBldr_ImportQueryExtn("C:\mainline\Reports\SrcSrchPrefPFFCreate.png","Query1")  		
'#										
'#	History	:						Developer Name			Date			Version				Changes Done			Reviewer	
'#------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'#										Sukhada B					31-12-12		1.0																	Sandeep N
'#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_QryBldr_ImportQueryExtn(sFilePath,sNewQueryName)  
	GBL_FAILED_FUNCTION_NAME="Fn_QryBldr_ImportQueryExtn"
Dim ObjQryApp
Set ObjQryApp =  Fn_UI_ObjectCreate( "Fn_QryBldr_ImportQueryExtn",JavaWindow("Search_QueryBuilder").JavaWindow("QueryBuilder"))

		'++++++++++<<    Click [Import] button  >>++++++++++
		Call Fn_CheckBox_Select("Fn_QryBldr_ImportQueryExtn", ObjQryApp, "Import")

		'++++++++++<<     Invoke [...] button on [Import] dialog  >>++++++++++
		Call Fn_Button_Click( "Fn_QryBldr_ImportQueryExtn",ObjQryApp, "BrowseFilePath")

		'++++++++++<<      Specify \Path\FileName under [Read Query Definition] dialog  >>++++++++++
		Call Fn_Edit_Box("Fn_QryBldr_ImportQueryExtn",JavaDialog("ImportPath"),"Filename",sFilePath)

		'++++++++++<<     Invoke [Import] button on [Read Query Definition] dialog  >>++++++++++
		Call Fn_Button_Click( "Fn_QryBldr_ImportQueryExtn",JavaDialog("ImportPath") , "Import")

			'++++++++++<<     Invoke [Verify] button on [Import] dialog  >>++++++++++
			Call Fn_Button_Click( "Fn_QryBldr_ImportQueryExtn",ObjQryApp , "Verify")
	
			'++++++++++<<    Invoke [OK] button on [Import] dialog   >>++++++++++
			Call Fn_Button_Click( "Fn_QryBldr_ImportQueryExtn",ObjQryApp , "OK")
	If sNewQueryName <> "" Then

				Call Fn_Edit_Box("Fn_QryBldr_ImportQueryExtn",ObjQryApp,"Name",sNewQueryName)
	End If
			Call Fn_ReadyStatusSync(2)
			Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Pass: Successfully Imported the Requested Query. ")   		
			Fn_QryBldr_ImportQueryExtn = True

'		End If

		Call Fn_Button_Click( "Fn_QryBldr_ImportQueryExtn", ObjQryApp, "Create")

Set ObjQryApp = Nothing
End Function

''*********************************************************  Fn_SISW_Search_QuickSearch.  ***********************************************************************
'Function Name  :    Fn_SISW_Search_QuickSearch
'Description		:	For Quick Search Operation
'Parameters      :   1. sSrchType 2. sSrchText
'Return Value   :   True/False
'Pre-requisite  :    Quick Search Pane should be visible
'Examples		:  	Fn_SISW_Search_QuickSearch("Item ID", 000211)
'History      		: 	Name            		Date      						Rev. No.      	Changes Done      								Release(Build)
'Develop By	:		
'Modified By    Harshal Agrawal     30 Dec 2010      		2.0           Dialog open Quick Search Result 		Teamcenter9(20101208)
'Modified By    Sonal P     				22 Feb 2011      		3.0           Added Case "StringID"
'Modified By    Koustubh     				11 Apr 2012      		3.0           Removed case "StringID", commented unnecessary code
''Modified By    Ashok k.     				17 May 2012      		3.0           Modified Hierarchy  of dialog QuickOpenResults
''Modified By    Sandeep N     				05 Feb 2013      		4.0           Modified Case "Item ID" replace condition : If Cint(sSrchText) = Cint(aCellData(0)) Then bFlag = True with If Cdbl(sSrchText) = Cdbl(aCellData(0)) Then bFlag = True
'*********************************************************************************************************************************************************
Public Function Fn_SISW_Search_QuickSearch(sSrchType, sSrchText)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_Search_QuickSearch"
	Dim sMenu,iCount, bFlag,sSrchWild,sSrch
	Dim objQuickOpen,iRowCnt,sCellData ,aCellData

    JavaWindow("DefaultWindow").JavaToolbar("QuickSearchToolbar").ShowDropdown "Perform Search"
	wait 1
	'Added by VivekA -----------------------------------------------
	sSrchWild = ""
	If instr( 1 , sSrchText , "~" , 1 ) > 0 Then
		sSrch = split(sSrchText, "~")
		sSrchText =  sSrch(0)	'Item id 
		sSrchWild = sSrch(1)	'wild character	
	End If	
	'JavaWindow("DefaultWindow").WinMenu("ContextMenu").Select sSrchType
	' -- added by Koustubh
	If  sSrchType = "StringID" Then sSrchType = "Item ID"
	
	If Environment.Value("ProductName") = sUFTProductName Then 
		JavaWindow("DefaultWindow").WinMenu("ContextMenu").Select sSrchType
	Else
		JavaWindow("DefaultWindow").JavaMenu("Label:=" & sSrchType).Select
	End If
	wait 2
	'-------------------------------------------------------------------
	'Added by VivekA -----------------------------------------------
	If sSrchWild <> "" Then
		JavaWindow("DefaultWindow").JavaEdit("QuickSearch").Set sSrchWild
	Else 
		JavaWindow("DefaultWindow").JavaEdit("QuickSearch").Set sSrchText	
	End If
	'-----------------------------------------------------------------------
	wait 3
	'JavaWindow("DefaultWindow").JavaEdit("QuickSearch").Set sSrchText
	JavaWindow("DefaultWindow").JavaToolbar("QuickSearchToolbar").Press "Perform Search"
	wait 3
	Fn_SISW_Search_QuickSearch = True
	'[Harshal 06 Dec 2011]: for mulitple Hierarchy of object
	If JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("NoAccessibleObjects").Exist(SISW_MIN_TIMEOUT) Then
		JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("NoAccessibleObjects").JavaButton("OK").Click
		Fn_SISW_Search_QuickSearch = False
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Quick Search Object [" + sSrchText + "]")
	ElseIf JavaWindow("DefaultWindow").JavaWindow("NoAccessibleObjects").Exist(SISW_MICRO_TIMEOUT) Then
		JavaWindow("DefaultWindow").JavaWindow("NoAccessibleObjects").JavaButton("OK").Click
		Fn_SISW_Search_QuickSearch = False
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Quick Search Object [" + sSrchText + "]")
	End If
	If JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("QuickOpenResults").Exist(SISW_MIN_TIMEOUT)  Then
		Set objQuickOpen = Fn_UI_ObjectCreate("Fn_SISW_Search_QuickSearch", JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("QuickOpenResults") )
	ElseIf JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("QuickOpenResults").Exist(SISW_MIN_TIMEOUT)  Then
		Set objQuickOpen = Fn_UI_ObjectCreate("Fn_SISW_Search_QuickSearch",JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("QuickOpenResults") )
	Else
		Exit Function
	End If
		 				  
	If objQuickOpen.JavaButton("LoadAll").exist(SISW_DEFAULT_TIMEOUT) then
		If cInt(Fn_UI_Object_GetROProperty("Fn_SISW_Search_QuickSearch",objQuickOpen.JavaButton("LoadAll"), "enabled")) = 1 Then
			Call Fn_Button_Click("Fn_SISW_Search_QuickSearch", objQuickOpen, "LoadAll")
			Call Fn_ReadyStatusSync(5)
		End If
	End If
	iRowCnt = Fn_UI_Object_GetROProperty("Fn_SISW_Search_QuickSearch",objQuickOpen.JavaTable("QckSrchRsltsTable"),"rows")
	For iCount = 0 to iRowCnt - 1
		sCellData =  Fn_UI_JavaTable_GetCellData("Fn_SISW_Search_QuickSearch", objQuickOpen, "QckSrchRsltsTable", iCount,"0")
		Select Case sSrchType
			Case "Item ID"
				bFlag = False
							aCellData = Split(sCellData ,"-",-1)
							If isNumeric(aCellData(0)) Then
'								If Cint(sSrchText) = Cint(aCellData(0)) Then bFlag = True
								If Cdbl(sSrchText) = Cdbl(aCellData(0)) Then bFlag = True
							Else
								If (sSrchText) = (aCellData(0)) Then bFlag = True
							End If
							If bFlag Then
								 Call Fn_UI_JavaTable_SelectRow("Fn_SISW_Search_QuickSearch", objQuickOpen,"QckSrchRsltsTable",iCount)
								 Call Fn_Button_Click("Fn_SISW_Search_QuickSearch", objQuickOpen, "Open")
								 Fn_SISW_Search_QuickSearch = True
								 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Data Found in JavaTable on Row " & iCount)
								 Exit For
							Else 
								 Fn_SISW_Search_QuickSearch = False
								 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Data Not Found in JavaTable on Row " & iCount)
							End If

			Case "Dataset Name"
				If InStr(1,CStr(sCellData),CStr(sSrchText)) Then    'Modified by Manish Singh 14June 12
					Call Fn_UI_JavaTable_SelectRow("Fn_SISW_Search_QuickSearch", objQuickOpen,"QckSrchRsltsTable",iCount)
					Call Fn_Button_Click("Fn_SISW_Search_QuickSearch", objQuickOpen, "Open")
					Fn_SISW_Search_QuickSearch = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Data Found in JavaTable on Row " & iCount)
					Exit For
				Else 
					Fn_SISW_Search_QuickSearch = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Data Not Found in JavaTable on Row " & iCount)
				End If
				'- - -- - - - - - - - - - - - - - - -- - - - - - - - - - - - - - - -- - - - - - - - - - - - - - - -- - - - - - - - - - - - - - - -- - - - - - - - - - - - - - - -- - - - - - - - - - - - - 
			End Select
	Next

'	Wait (5)
'	For iCount=0 to 0 ' [Ketan 30 Dec 2010]: for mulitple Hierarchy of object
'		If JavaWindow("DefaultWindow").JavaWindow("NoAccessibleObjects").Exist(5) Then
'			JavaWindow("DefaultWindow").JavaWindow("NoAccessibleObjects").JavaButton("OK").Click
'			Fn_SISW_Search_QuickSearch = False
'			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Quick Search Object [" + sSrchText + "]")
'			Exit For
'			Exit Function
'		End If
'		JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("ErrorDialog").SetTOProperty "title","No accessible objects found"
'		If JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("ErrorDialog").Exist(5) Then
'			JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("ErrorDialog").JavaButton("OK").Click
'			Fn_SISW_Search_QuickSearch = False
'			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Quick Search Object [" + sSrchText + "]")
'			Exit For
'			Exit Function
'		End If
'	Next
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Object [" + sSrchText + "] Found through Quick Search")
	Set objQuickOpen = Nothing
End Function

'#####################################################  This function allows "Select" , "Verify" & "Get" operations under Quick Search.	##############################################~
'#	FUNCTION NAME		:	 	 Fn_SISW_Search_QuickSearchOperations
'#
'#	FUNCTION ID			: 		 273
'#
'#	MODULE				: 		 Search Requirement
'#			
'#	DESCRIPTION			:		 This function allows "Select" , "Verify" and Get Selected Text operations under Quick Search.	 
'#							
'#	PARAMETERS   		:        1. sAction		: Switch case (Select/Verify/Get)											
'#								 2. sSrchType	: Type of search
'#								 3. sSrchText	: Text to be entered into Edit box
'#	
'#	RETURN VALUE 		: 		 TRUE \ FALSE \ Return Value of Selected Text in Quick Search Edit
'#
'#	Examples				:	 CASE "MenuSelect" 		: Call Fn_SISW_Search_QuickSearchOperations("MenuSelect", "Bitmap", "ABC") 
'#								 CASE "GetSelectedText" : Call Fn_SISW_Search_QuickSearchOperations("GetSelectedText","","whySo")
'#										
'#	History				:		 
'#
'#	Developer Name			Date			Rev. No.				Changes Done											Reviewer	
'#---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'#	  Kavan Shah~		  02-06-10		   																					Sunil Rai
'#---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'#	  Ankit Nigam		02-Feb-2016		   					Added New Case "GetSelectedText"		[TC11.2.2:201611300:02FEB2016:REG-MyTC:VivekA:NewDevelopment]
'#---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'################################################     This function allows "Select" , "Verify" & "Get" operations under Quick Search. ##################################################~

Public Function Fn_SISW_Search_QuickSearchOperations(sAction, sSrchType, sSrchText)  
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_Search_QuickSearchOperations"
	Dim objJavaEdit, objCheck

	Select Case sAction
		Case "MenuSelect"
			Call Fn_QuickSearch(sSrchType, sSrchText)
		Case "VerifyActive"
	        Call Fn_ToolBar_ShowDropdown("Fn_SISW_Search_QuickSearchOperations",JavaWindow("DefaultWindow"),"QuickSearchToolbar","#1")
			objCheck = JavaWindow("DefaultWindow").WinMenu("ContextMenu").CheckItemProperty(sSrchType, "Checked", True , 10)
			Set  objJavaEdit = Fn_UI_ObjectCreate("Fn_QuickOpenResults", JavaWindow("DefaultWindow").JavaEdit("QuickSearch"))
			objJavaEdit.SetFocus	
		Case "GetSelectedText"	
			Set  objJavaEdit = Fn_UI_ObjectCreate("Fn_QuickOpenResults", JavaWindow("DefaultWindow").JavaEdit("QuickSearch"))
			objJavaEdit.Set sSrchText
			objJavaEdit.DblClick 2,2,"LEFT"
	End Select
	If Err.Number < 0 Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Result Not Found  For  " & sSrchText)
		Fn_SISW_Search_QuickSearchOperations = False
	Else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Result  Found  For  " & sSrchText)
		If sAction = "GetSelectedText" Then
			Fn_SISW_Search_QuickSearchOperations = objJavaEdit.Object.getSelectionText()						
		Else 
			Fn_SISW_Search_QuickSearchOperations = True
		End If
	End If
Set objCheck = Nothing
Set objJavaEdit = Nothing
End Function 

'-------------------------------------------------------------------- Set Requisite Search Preference-----------------------------------------------------------------------------------------------------------
'#
	'# FUNCTION NAME:	Fn_SISW_Search_SrchSavedSearchOperation
'#
'#
'# DESCRIPTION:			 1. Check/UnCheck the[Is Shared] box
'#											2. Click on [Create In] button on [Add Search to My Saved Searches]
'#											3. Expand and select prefered Tree Path under [My Saved Search]
'#											
'#											Case: Add_To_My_Saved_Search  		' 
'#													a. Specify [Name] of the saved search									
'#													b. Specify Folder Name on [Folder Information] dialog and click [OK]
'#													c. Select Path
'#											Case: Saved Search Delete
'#													a. Click on [Delete] button
'#													b. Click [OK] on Warning dialog		
'#											Case: Saved Search Rename
'#												a. Click on [Rename] button
'#												b. Specify New Name through send Key operation
'#												c. Send [Enter] Key						
'#											Case: Saved Search Validate		' Need to pass Full Reference path in "sSearchName" for existance check.							
'#												a. Validate existance of saved search
'#													4. Click [OK]
'#		
'# PRE-REQUISITE:		1. RAC Session accessible and My Teacenter Application Search Pane loaded
'#											 2. Search Criteria applied with various input values 
'#									>>   3. Click on Toolbar button [Add Search To My Saved Searches] under Search Pane
'#							
'# PARAMETERS   :       sAction: Name of the case to exercise pertaining to saved search
'#                                           bIsShared : ON/OFF
'#										 	 sSourceFolderPath: Existing Search Folder Path
'#											 sNewName: New Name for folder OR Search
'#
'# RETURN VALUE : 		TRUE \ FALSE
'#
'#Examples	:				Fn_SISW_Search_SrchSavedSearchOperation("Add_To_My_Saved_Search",  "ON", "A:B:C", "NewName") 
'#										Fn_SISW_Search_SrchSavedSearchOperation("Add_To_My_Saved_Search", "On", "","Sandeep") 
'#
'#	History	:						Developer Name			Date			Version				Changes Done			Reviewer	
'#------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'#										Sandeep N						07-10-10			1.0																Sunny R
'#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Imp Note :--- To devolope other Cases take Help of function Fn_MyTc_SrchSavedSearchOperation
'							And for OR help of Search.tsr
Public Function  Fn_SISW_Search_SrchSavedSearchOperation(sAction,  bIsShared, sSourceFolderPath, sNewName)
GBL_FAILED_FUNCTION_NAME="Fn_SISW_Search_SrchSavedSearchOperation"
Dim iNode, sNode, iCount, bExist, aRefPath, iCountOuter, afullRefPath, ArrNode, iEle, bFolder
Dim ObjJavaWndw, ObjQkLnkWndwTr, ObjQkLnkWndw, ObjAddSrch, iCont, sFullPath

		Select Case sAction
				Case "Add_To_My_Saved_Search" 			
						'	[Example ( last parameter is for name and its compulsary )]			Fn_SISW_Search_SrchSavedSearchOperation("Add_To_My_Saved_Search",  "ON", "A:B:C", "NewName") 
						' If Dont want to save/create new folder then pass blank "" in parameter "sSourceFolderPath"  ->> this will save query under parent node.
						'<<<<<<<<<<<<<<<<<<<<<<<  Add to my saved search  >>>>>>>>>>>>>>>>>>>>>>>>> 
						If True = Fn_UI_ObjectExist("Fn_SISW_Search_SrchSavedSearchOperation",JavaWindow("DefaultWindow").JavaWindow("AddSrchtoMySaved")) Then

'								Set ObjAddSrch = Fn_UI_ObjectCreate( "Fn_SISW_Search_SrchSavedSearchOperation",JavaWindow("DefaultWindow").JavaWindow("AddSrchtoMySaved"))		
								Set ObjAddSrch =  Fn_SISW_UI_Object_Operations("Fn_SISW_Search_SrchSavedSearchOperation","Create", JavaWindow("DefaultWindow").JavaWindow("AddSrchtoMySaved"), SISW_MIN_TIMEOUT)

								If sSourceFolderPath <> "" Then
												
											Call Fn_Button_Click("Fn_SISW_Search_SrchSavedSearchOperation", ObjAddSrch, "CreateIn")
				
											aRefPath = Split(sSourceFolderPath,":")
											afullRefPath = "My Saved Searches"
										
											For iCountOuter = 0 To  UBound(aRefPath)
														bExist = False
														For  iCount=1 to ObjAddSrch.JavaTree("ExistingSavedSrchs").GetROProperty("items count")-1
																sNode = ObjAddSrch.JavaTree("ExistingSavedSrchs").GetItem(iCount)
																If  sNode = afullRefPath+":"+aRefPath(iCountOuter) Then
																	bExist = true
																	Exit For
																End If
														Next
														If True =  bExist Then
																	afullRefPath =afullRefPath+":"+aRefPath(iCountOuter) 
																	Call Fn_JavaTree_Select("Fn_SISW_Search_SrchSavedSearchOperation", ObjAddSrch, "ExistingSavedSrchs", afullRefPath)
														Else
																'CreateFolder
																Call Fn_Button_Click("Fn_SISW_Search_SrchSavedSearchOperation", ObjAddSrch, "NewFolder")
'																If True = Fn_UI_ObjectExist("Fn_SISW_Search_SrchSavedSearchOperation",JavaWindow("DefaultWindow").JavaWindow("Folder Information") ) Then
																If True = Fn_SISW_UI_Object_Operations("Fn_SISW_Search_SrchSavedSearchOperation", "Exist", JavaWindow("DefaultWindow").JavaWindow("Folder Information"), SISW_MIN_TIMEOUT)  Then
																		Call Fn_Edit_Box("Fn_SISW_Search_SrchSavedSearchOperation",JavaWindow("DefaultWindow").JavaWindow("Folder Information"),"Name","")
																		Call Fn_UI_EditBox_Type("Fn_SISW_Search_SrchSavedSearchOperation",JavaWindow("DefaultWindow").JavaWindow("Folder Information"),"Name", aRefPath(iCountOuter) )
																		Call Fn_Button_Click("Fn_SISW_Search_SrchSavedSearchOperation",JavaWindow("DefaultWindow").JavaWindow("Folder Information"), "OK")		
																		Call Fn_ReadyStatusSync(5)
																		afullRefPath =afullRefPath+":"+aRefPath(iCountOuter) 
																		Call Fn_JavaTree_Select("Fn_SISW_Search_SrchSavedSearchOperation", ObjAddSrch, "ExistingSavedSrchs", afullRefPath)
																Else
																		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed: Window Folder Create Does Not Exist ")
																		Fn_SISW_Search_SrchSavedSearchOperation = False
																End If
														End If 	
											Next
								End If
								Call Fn_CheckBox_Set("Fn_SISW_Search_SrchSavedSearchOperation", ObjAddSrch, "IsShared", bIsShared )
								Call Fn_Edit_Box("Fn_SISW_Search_SrchSavedSearchOperation",ObjAddSrch,"Name","")
								Call Fn_UI_EditBox_Type("Fn_SISW_Search_SrchSavedSearchOperation",ObjAddSrch,"Name", sNewName )
                                Call Fn_Button_Click("Fn_SISW_Search_SrchSavedSearchOperation", ObjAddSrch, "OK")

								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Passed:Search Successfully Added to My Saved Searches" )
								Fn_SISW_Search_SrchSavedSearchOperation = True
						Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed: Window Does Not Exist." )
								Fn_SISW_Search_SrchSavedSearchOperation = False
						End If                				
				Case Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Fail: Requested Action Not Valid. ")   	
							Fn_SISW_Search_SrchSavedSearchOperation = False
				End Select

Set ObjAddSrch = Nothing
End Function			


'*********************************************************	Fn_SISW_Search_AdhocClassificationQuerySearchOperations  ***********************************************************************
'Function Name   :				Fn_SISW_Search_AdhocClassificationQuerySearchOperations()

'Description	 :		 		 The function is used to perform operations on Adhoc Classification Query Search.

'Parameters	:	 			1.  sAction
'							2.  sSearchClassificationClass - list of classnames, conditions, unit of measurements in 
'										' Condition:ClassName:unit~Condition1:ClassName1:unit1' format
'							3.  sRemoveList - for future use 
'							4.  sRow - list of row numbers separated by ~
'							5.  sCol - list of column names separated by ~
'									Column Names are - Condition, Property Name, Operator, Searching Value
'							6.  sValue - for future use
'							7.  sBtnName : Buttun Name to be clicked
'										 OK / Cancel
													
'Return Value	: 			True / False

'Pre-requisite	:		 		Search Tab should be closed..

'Examples		:				Call Fn_SISW_Search_AdhocClassificationQuerySearchOperations("SearchAdhocClassificationQuery", ":Storage_55290:~OR:ICOClass_23109:", "","", "","","OK")
'								Note :: Column Names are - Condition, Property Name, Operator, Searching Value
'								Call  Fn_SISW_Search_AdhocClassificationQuerySearchOperations("VerifyInPropertyTable", "", "", "0~0", "Condition~Operator", "~>", "")
'								bReturn=  Fn_SISW_Search_AdhocClassificationQuerySearchOperations("SelectSearchedAdhocClassificationQuery", "Storage_14912:Storage_14912.AutomatedLov", "", "", "", "", "OK")
									  
'History:
'					Developer Name			Date			Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'					Koustubh W				4-1-2011
'					SHREYAS					3-02-2011		Prasanna
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public function Fn_SISW_Search_AdhocClassificationQuerySearchOperations(sAction, sSearchClassificationClass, sRemoveList, sRow, sCol, sValue, sBtnName)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_Search_AdhocClassificationQuerySearchOperations"
	Dim objAdvanceSearch, aClassificationClass, aClassDetails,iRowCount,sCellData
	Dim objSelectType, intNoOfObjects, iCount, iCounter,sNode
	Fn_SISW_Search_AdhocClassificationQuerySearchOperations = False
	Set objAdvanceSearch = JavaWindow("MyTeamcenter_Search").JavaWindow("SrchDefaultWndw").JavaDialog("Advanced_MultiSite")
	If objAdvanceSearch.Exist(SISW_MIN_TIMEOUT) = False Then
		If True = Fn_ToolbatButtonClick("Open Search View") Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Clicked : 'Open Search View' ToolBar Button")					
		Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : [ Fn_SISW_Search_AdhocClassificationQuerySearchOperations ] Failed to Click on Open 'Open Search View' ToolBar Button")
			exit function
		End IF

		If True = Fn_ToolbatButtonClick("View Menu") Then
			JavaWindow("DefaultWindow").WinMenu("ContextMenu").Select "Extended Multi-Application Search..."
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Clicked : View Menu ToolBar Button")					
		Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : [ Fn_SISW_Search_AdhocClassificationQuerySearchOperations ] Failed to Click on Open View Menu ToolBar Button")
			exit function
		End IF
		If objAdvanceSearch.Exist(SISW_MIN_TIMEOUT) = False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : [ Fn_SISW_Search_AdhocClassificationQuerySearchOperations ] Failed to Click on Open Advanced Search window.")
			exit function
		End If
	End If
	Select Case sAction
		Case "SearchAdhocClassificationQuery"
			' selecting tab
			objAdvanceSearch.JavaTab("TabSwitch").Select "Adhoc Classification Query"
			' clicking on cler
			objAdvanceSearch.JavaButton("Clear").Click micLeftBtn
			' setting classsification class
			If sSearchClassificationClass <> "" Then
				aClassificationClass = split(sSearchClassificationClass,"~")
				For iCounter = 0 to UBound(aClassificationClass)
					aClassDetails = split(aClassificationClass(iCounter),":")
					objAdvanceSearch.JavaCheckBox("Search Classification Class").Set "ON"
                                                  Fn_ReadyStatusSync(2)
					Set objSelectType = description.Create()
					objSelectType("Class Name").value = "JavaButton"
					Set  intNoOfObjects = JavaWindow("MyTeamcenter_Search").ChildObjects(objSelectType)
					 'Clicking on Clear button
					  For iCount = 0 to intNoOfObjects.count-1
						   If  intNoOfObjects(iCount).getROProperty("attached text") = "clear_16" Then
							intNoOfObjects(iCount).click  micLeftBtn
							Exit for
						   End If
					  Next
					' setting class name
					JavaWindow("MyTeamcenter_Search").JavaEdit("Name").SetTOProperty "attached text", "Class/Attribute Selection Popup"
					JavaWindow("MyTeamcenter_Search").JavaEdit("Name").Set trim(aClassDetails(1) )
	
					 'Clicking on Find button
					  For iCount = 0 to intNoOfObjects.count-1
						   If  intNoOfObjects(iCount).getROProperty("attached text") = "find_16" Then
							intNoOfObjects(iCount).click  micLeftBtn
							Exit for
						   End If
					  Next
					 'double clicking on selected class
					objSelectType("Class Name").value = "JavaTree"
					Set  intNoOfObjects = JavaWindow("MyTeamcenter_Search").ChildObjects(objSelectType)
					  For iCount = 0 to intNoOfObjects.count-1
						   If  intNoOfObjects(iCount).getROProperty("attached text") = "Class/Attribute Selection Popup" Then
							'sNode=intNoOfObjects(iCount).GetROProperty ("value")
							intNoOfObjects(iCount).Activate intNoOfObjects(iCount).GetROProperty ("value")
							Exit for
						   End If
					  Next
					  'setting measurement unit
					  If trim(aClassDetails(2)) <> ""  Then
						objAdvanceSearch.JavaCheckBox("Set system of measurement").Set "ON"
						objAdvanceSearch.JavaMenu("Label:=" & trim(aClassDetails(2))).Select
					  End If
					  ' setting condition
					  If trim(aClassDetails(0)) <> ""  Then
						objAdvanceSearch.JavaTable("Search Classification").ClickCell objAdvanceSearch.JavaTable("Search Classification").Object.getSelectedRow, 0,"LEFT"  
						If objAdvanceSearch.JavaList("SearchClassification").Exist(SISW_MIN_TIMEOUT) Then
							objAdvanceSearch.JavaList("SearchClassification").Select trim(aClassDetails(0))
						Else
							objAdvanceSearch.JavaTable("Search Classification").ClickCell objAdvanceSearch.JavaTable("Search Classification").Object.getSelectedRow,"Property Name","LEFT"  
							objAdvanceSearch.JavaList("SearchClassification").Select trim(aClassDetails(0))
						End If
					  End If
				Next
			End If
			Fn_SISW_Search_AdhocClassificationQuerySearchOperations = True
' - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - -  - - - - - -
		Case "SelectSearchedAdhocClassificationQuery"
			Dim aCols
			' selecting tab
			objAdvanceSearch.JavaTab("TabSwitch").Select "Adhoc Classification Query"
			' clicking on cler
			objAdvanceSearch.JavaButton("ClearTable").Click micLeftBtn
			' setting classsification class
			If sSearchClassificationClass <> "" Then
				aClassificationClass = split(sSearchClassificationClass,"~")
				If Trim(sCol) = "" Then
					sCol = " "
				End If
				aCols = split(sCol,"~")
				For iCounter = 0 to UBound(aClassificationClass)
					aClassDetails = split(aClassificationClass(iCounter),":")
					objAdvanceSearch.JavaCheckBox("Search Classification Class").Set "ON"
                                                  Fn_ReadyStatusSync(2)
					Set objSelectType = description.Create()
					objSelectType("Class Name").value = "JavaButton"
					Set  intNoOfObjects = JavaWindow("MyTeamcenter_Search").ChildObjects(objSelectType)
					 'Clicking on Clear button
					  For iCount = 0 to intNoOfObjects.count-1
						   If  intNoOfObjects(iCount).getROProperty("attached text") = "clear_16" Then
							intNoOfObjects(iCount).click  micLeftBtn
							Exit for
						   End If
					  Next
					  
					' setting class name
					'JavaWindow("MyTeamcenter_Search").JavaEdit("Name").SetTOProperty "attached text", "Class/Attribute Selection Popup"
					'JavaWindow("MyTeamcenter_Search").JavaEdit("Name").Set trim(aClassDetails(0) )
					'[TC1122-2016010600-21_01_2016-VivekA-Maintenance] - Added by Shantan S
					If JavaWindow("MyTeamcenter_Search").JavaWindow("SrchDefaultWndw").JavaDialog("Advanced_MultiSite").JavaWindow("PopupWindow").JavaEdit("ClassAttributeSelection").Exist(1) Then
						JavaWindow("MyTeamcenter_Search").JavaWindow("SrchDefaultWndw").JavaDialog("Advanced_MultiSite").JavaWindow("PopupWindow").JavaEdit("ClassAttributeSelection").Set Trim(aClassDetails(0))
					Else
						JavaWindow("MyTeamcenter_Search").JavaEdit("Name").SetTOProperty "attached text", "Class/Attribute Selection Popup"
						JavaWindow("MyTeamcenter_Search").JavaEdit("Name").Set Trim(aClassDetails(0))	
					End If
					
					 'Clicking on Find button
					  For iCount = 0 to intNoOfObjects.count-1
						   If  intNoOfObjects(iCount).getROProperty("attached text") = "find_16" Then
							intNoOfObjects(iCount).click  micLeftBtn
							Exit for
						   End If
					  Next
					 'double clicking on selected class
					objSelectType("Class Name").value = "JavaTree"
					Set  intNoOfObjects = JavaWindow("MyTeamcenter_Search").ChildObjects(objSelectType)
					  For iCount = 0 to intNoOfObjects.count-1
						   If  intNoOfObjects(iCount).getROProperty("attached text") = "Class/Attribute Selection Popup" Then
							intNoOfObjects(iCount).Activate intNoOfObjects(iCount).GetROProperty ("value")
							Exit for
						   End If
					  Next

					'Select the value from the class Dropdown list
					iRowCount=objAdvanceSearch.JavaTable("Search Classification").GetROProperty("rows")
					For iCount=0 to iRowCount-1
						sCellData=objAdvanceSearch.JavaTable("Search Classification").GetCellData(iCount,1)
						If (LCase( Trim(sCellData) ) = LCase( Trim( aClassDetails(0)))) Then
							objAdvanceSearch.JavaTable("Search Classification").ClickCell iCount,1,"LEFT"
							wait(3)
							objAdvanceSearch.JavaList("SearchClassification").Select trim(aClassDetails(1))
							objAdvanceSearch.JavaTable("Search Classification").ClickCell iCount,3,"LEFT"
							wait(1)
							If objAdvanceSearch.JavaList("SearchClassification").Exist then 
								objAdvanceSearch.JavaList("SearchClassification").Select trim(aCols(iCounter))
							Elseif objAdvanceSearch.JavaEdit("SearchClassification").Exist then 
									objAdvanceSearch.JavaEdit("SearchClassification").Set trim(aCols(iCounter))
							Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Neither a List Nor a EditBox is present")
									Fn_SISW_Search_AdhocClassificationQuerySearchOperations = False
									Exit function
							 End if
							Exit For
						End If
					Next
				Next
			End If
		'       check if list exists..
'			If objAdvanceSearch.JavaList("SearchClassification").Exist then 
'									objAdvanceSearch.JavaList("SearchClassification").Select sCol
'			Elseif objAdvanceSearch.JavaEdit("SearchClassification").Exist then 
'									objAdvanceSearch.JavaEdit("SearchClassification").Set sCol
'			Else
'									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Neither a List Nor a EditBox is present")
'									Fn_SISW_Search_AdhocClassificationQuerySearchOperations = False
'									Exit function
'			 End if

			'it was observed that OK button was not getting clicked so handled that part over here..
			'so whenever using this case, the "sBtnName" parameter should be passed blank

			objAdvanceSearch.JavaButton("OK").Click micLeftBtn

			Fn_SISW_Search_AdhocClassificationQuerySearchOperations = True
' - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - -  - - - - - -
		Case "VerifyInPropertyTable"
			Dim sVal, aRows, aColumns, aValues  
			aRows = split(sRow,"~") 
			aColumns = split(sCol,"~") 
			aValues = split(sValue,"~")
			For iCounter = 0 to UBound(aRows)
				Fn_SISW_Search_AdhocClassificationQuerySearchOperations = False
				Select Case trim(aColumns(iCounter))
					Case "Condition"
						sVal = objAdvanceSearch.JavaTable("Search Classification").GetCellData(cInt(trim(aRows(iCounter))), 0)
					Case "Property Name"
						sVal = objAdvanceSearch.JavaTable("Search Classification").GetCellData(cInt(trim(aRows(iCounter))), 1)
					Case "Operator"
						sVal = objAdvanceSearch.JavaTable("Search Classification").GetCellData(cInt(trim(aRows(iCounter))), 2)
					Case "Searching Value"
						sVal = objAdvanceSearch.JavaTable("Search Classification").GetCellData(cInt(trim(aRows(iCounter))), 3)
					Case Else
						Exit for
				End Select
				If trim(sVal) <> trim(aValues(iCounter)) Then
					Exit for
				End If
				Fn_SISW_Search_AdhocClassificationQuerySearchOperations = True
			Next
' - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - -  - - - - - -
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : [ Fn_SISW_Search_AdhocClassificationQuerySearchOperations ] Invalid Case [ " & sAction & " ]")
			exit function
' - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - -  - - - - - -
	End Select
	'clicking on button
	If sBtnName <> "" Then
			If lcase(sBtnName) = "ok" Then
				objAdvanceSearch.JavaButton(sBtnName).SetToProperty "Index",0				
			End If
				objAdvanceSearch.JavaButton(sBtnName).Click micLeftBtn
	End If
	Set objAdvanceSearch = nothing
	Set  intNoOfObjects = nothing
	Set objSelectType = nothing
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : [ Fn_SISW_Search_AdhocClassificationQuerySearchOperations ] Successfully executed with case [ " & sAction &" ].")	
End Function

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_SISW_Search_SimpleSearchBOTypeOprations
'@@
'@@    Description				 :	Function Used to perform operations on Simple Search Business Object Types
'@@
'@@    Parameters			   :	1.StrAction: Action Name
'@@											  2.StrBOType: Business Object Type
'@@
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	Simple Search Tab Should be open							
'@@
'@@    Examples					:	Call Fn_SISW_Search_SimpleSearchBOTypeOprations("Select","Item")
'@@											 Call Fn_SISW_Search_SimpleSearchBOTypeOprations("Exist","ItemRevision")
'@@											 Call Fn_SISW_Search_SimpleSearchBOTypeOprations("SelectFromToolBar","ANDList")
'@@											 Call Fn_SISW_Search_SimpleSearchBOTypeOprations("GetCurrentBOType","")
'@@											 Call Fn_SISW_Search_SimpleSearchBOTypeOprations("GetBOTypeCurrentPosition","Item")
'@@											 Call Fn_SISW_Search_SimpleSearchBOTypeOprations("BOTypeExistInDropDown","Item")
'@@											 Call Fn_SISW_Search_SimpleSearchBOTypeOprations("GetBOTypeToolTipText","")
'@@											 Call Fn_SISW_Search_SimpleSearchBOTypeOprations("BOTypeExistInDialog","ANDList~Company~Dataset")
'@@
'@@	   History					 	:	
'@@													Developer Name												Date						Rev. No.						Changes Done								Reviewer
'@@------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@												Sandeep Navghane									08-Aug-2011						1.0																								Sunny Ruparel
'@@												Sandeep Navghane									09-Aug-2011						1.1									Added Cases										Sunny Ruparel
'@@																																															GetBOTypeCurrentPosition,BOTypeExistInDropDown
'@@												Sandeep Navghane									12-Aug-2011						1.2						 GetBOTypeToolTipText
'@@												Sandeep Navghane									1-Sep-2011						 1.3					 BOTypeExistInDialog
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'Pre-Requisites : - Simple Search Tab Should be open
'BOType=Business Object Type
Function Fn_SISW_Search_SimpleSearchBOTypeOprations(StrAction,StrBOType)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_Search_SimpleSearchBOTypeOprations"
   Dim bFlag, WshShell,arrBOTypes,iCounter,iCount,iRowCount
	Fn_SISW_Search_SimpleSearchBOTypeOprations=False
	bFlag=False
	Select Case StrAction
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Select"
			bFlag=Fn_SISW_Search_SimpleSearchBOTypeOprations("Exist",StrBOType)
			If bFlag=True Then
				Call Fn_List_Select("Fn_SISW_Search_SimpleSearchBOTypeOprations", JavaWindow("DefaultWindow"), "BusinessObjectType",StrBOType)
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),  "Sucessfully Selected Business Object Type :" & StrBOType &" from Business Object Type list")
				Fn_SISW_Search_SimpleSearchBOTypeOprations=True
			End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
         Case "Exist"
			 bFlag=Fn_UI_ListItemExist("Fn_SISW_Search_SimpleSearchBOTypeOprations", JavaWindow("DefaultWindow"), "BusinessObjectType",StrBOType)
			 Call Fn_WriteLogFile(Environment.Value("TestLogFile"),  "Sucessfully verified Business Object Type :" & StrBOType &" is exist in [ Business Object Type ] list")
			 If bFlag=True Then
				Fn_SISW_Search_SimpleSearchBOTypeOprations=True
			 End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "SelectFromToolBar"
			If Not JavaWindow("DefaultWindow").JavaWindow("SelectBusinessObjectType").Exist(SISW_DEFAULT_TIMEOUT) Then
				Call Fn_ToolbatButtonClick("Select a Business Object Type To Search.")
			End If
			wait 5
			'Added by Sanjeet Kumar on 28-Feb-2013
            JavaWindow("DefaultWindow").JavaWindow("SelectBusinessObjectType").JavaEdit("ChooseBOType").Set StrBOType
			Wait 2
			iRowCount=Fn_UI_Object_GetROProperty("Fn_SISW_Search_SimpleSearchBOTypeOprations",JavaWindow("DefaultWindow").JavaWindow("SelectBusinessObjectType").JavaTable("BusinessObjectType"), "rows")
			For iCounter=0 To iRowCount-1
				strCurrBOType=JavaWindow("DefaultWindow").JavaWindow("SelectBusinessObjectType").JavaTable("BusinessObjectType").GetCellData(iCounter,0)
				If Trim(strCurrBOType)=Trim(StrBOType) Then
					wait 1
                    JavaWindow("DefaultWindow").JavaWindow("SelectBusinessObjectType").JavaTable("BusinessObjectType").SelectRow iCounter
					JavaWindow("DefaultWindow").JavaWindow("SelectBusinessObjectType").JavaTable("BusinessObjectType").SelectCell iCounter,"0"
					Call Fn_Button_Click("Fn_SISW_Search_SimpleSearchBOTypeOprations", JavaWindow("DefaultWindow").JavaWindow("SelectBusinessObjectType"), "OK")
					bFlag=True
					Exit For
				End If
			Next
			If bFlag=True Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),  "Sucessfully Selected Business Object Type :" & StrBOType &" from Business Object Type Table")
				Fn_SISW_Search_SimpleSearchBOTypeOprations=True
			Else
				Call Fn_Button_Click("Fn_SISW_Search_SimpleSearchBOTypeOprations", JavaWindow("DefaultWindow").JavaWindow("SelectBusinessObjectType"), "Cancel")
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "GetCurrentBOType"
				Fn_SISW_Search_SimpleSearchBOTypeOprations=Fn_UI_Object_GetROProperty("Fn_SISW_Search_SimpleSearchBOTypeOprations",JavaWindow("DefaultWindow").JavaList("BusinessObjectType"), "value")
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "GetBOTypeCurrentPosition"
			  Fn_SISW_Search_SimpleSearchBOTypeOprations=JavaWindow("DefaultWindow").JavaList("BusinessObjectType").GetItemIndex(StrBOType)
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "BOTypeExistInDropDown"
			Call Fn_ToolBarOperation("ShowDropdown", "Select a Business Object Type To Search.", "" )
			Fn_SISW_Search_SimpleSearchBOTypeOprations=Fn_MenuOperation("Exist",StrBOType)
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Type"
				'Type value in Business Object Type List.
				JavaWindow("DefaultWindow").JavaList("BusinessObjectType").Type StrBOType
				'SendKey for Esc and Enter
				Set WshShell = CreateObject("WScript.Shell")
					WshShell.SendKeys "{ESC}"
					WshShell.SendKeys "{ENTER}"
				Set WshShell = nothing
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),  "Sucessfully entered " & StrBOType &" from Business Object Type list")
				Fn_SISW_Search_SimpleSearchBOTypeOprations=True
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	Case "GetBOTypeToolTipText"
			If JavaWindow("DefaultWindow").JavaList("BusinessObjectType").Exist(5) Then
				Fn_SISW_Search_SimpleSearchBOTypeOprations=JavaWindow("DefaultWindow").JavaList("BusinessObjectType").Object.getToolTipText()
			End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "BOTypeExistInDialog"
			If Not JavaWindow("DefaultWindow").JavaWindow("SelectBusinessObjectType").Exist(8) Then
				Call Fn_ToolbatButtonClick("Select a Business Object Type To Search.")
			End If
			arrBOTypes=Split(StrBOType,"~")
			wait 5
			iRowCount=Fn_UI_Object_GetROProperty("Fn_SISW_Search_SimpleSearchBOTypeOprations",JavaWindow("DefaultWindow").JavaWindow("SelectBusinessObjectType").JavaTable("BusinessObjectType"), "rows")
			For iCount=0 To UBound(arrBOTypes)
				bFlag=False
				For iCounter=0 To iRowCount-1
					strCurrBOType=JavaWindow("DefaultWindow").JavaWindow("SelectBusinessObjectType").JavaTable("BusinessObjectType").GetCellData(iCounter,0)
					If Trim(strCurrBOType)=Trim(arrBOTypes(iCount)) Then
						wait 1
						bFlag=True
						Exit For
					End If
				Next
				If bFlag=False Then
					Exit For
				End If
			Next
			If bFlag=True Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),  "Sucessfully Selected Business Object Type :" & StrBOType &" from Business Object Type Table")
				Fn_SISW_Search_SimpleSearchBOTypeOprations=True
			End If
			Call Fn_Button_Click("Fn_SISW_Search_SimpleSearchBOTypeOprations", JavaWindow("DefaultWindow").JavaWindow("SelectBusinessObjectType"), "Cancel")
	End Select
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_SISW_Search_SimpleSearchPropertyTreeOprations
'@@
'@@    Description				 :	Function Used to perform operations on Property Tree
'@@
'@@    Parameters			   :	1.StrAction: Action Name
'@@											  2.StrProperty: Property Name
'@@
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	Property Tree Should be appear on screen							
'@@
'@@    Examples					:	Call Fn_SISW_Search_SimpleSearchPropertyTreeOprations("Select","Name")
'@@											 Call Fn_SISW_Search_SimpleSearchPropertyTreeOprations("Select","Owning Project(Project ID)")
'@@											 Call Fn_SISW_Search_SimpleSearchPropertyTreeOprations("Verify","Owning Project(Project ID)")
'@@											 Call Fn_SISW_Search_SimpleSearchPropertyTreeOprations("GetAllNodes","")
'@@
'@@	   History					 	:	
'@@													Developer Name												Date						Rev. No.						Changes Done								Reviewer
'@@------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@												Sandeep Navghane									08-Aug-2011						1.0																								Sunny Ruparel
'@@												Sandeep Navghane									12-Aug-2011						1.1							Added Case "Verify"									Sunny Ruparel
'@@												Sandeep Navghane									19-Aug-2011						1.2							Added Case "GetAllNodes"									Sunny Ruparel
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Function Fn_SISW_Search_SimpleSearchPropertyTreeOprations(StrAction,StrProperty)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_Search_SimpleSearchPropertyTreeOprations"
   Dim ObjPropertyTree
   Dim iNodeCount,iCounter,strCurrProperty,StrUpdatedProp,bFlag
   Fn_SISW_Search_SimpleSearchPropertyTreeOprations=False
   Set ObjPropertyTree=JavaWindow("DefaultWindow").JavaTree("SimpleSearchProperties")
   Select Case StrAction
	 	Case "Select"
	 		'[TC1123-20160518-27_05_2016-VivekA-Maintenance] - Added by Vivek as per Design change and as discussed with Akshay J
	 		If StrProperty = "Date Created" Then
	 			bFlag = Fn_SISW_Search_SimpleSearchPropertyTreeOprations("Verify",StrProperty)
	 			If bFlag = False Then
	 				StrProperty = "Creation Date"
	 			End If
	 		End If
	 		'-------------------------------------------------
			ObjPropertyTree.Select StrProperty
			If  Err.Number < 0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Select Property [" + StrProperty + "] from Property Tree")
				Fn_SISW_Search_SimpleSearchPropertyTreeOprations=False
			Else
                 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Selected Property [" + StrProperty + "] from Property Tree")
				 Fn_SISW_Search_SimpleSearchPropertyTreeOprations=True
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Verify"
			iNodeCount=ObjPropertyTree.GetROProperty("items count")
			For iCounter=0 To iNodeCount-1
				strCurrProperty=ObjPropertyTree.GetItem(iCounter)
				If Trim(strCurrProperty)=Trim(StrProperty) Then
					Fn_SISW_Search_SimpleSearchPropertyTreeOprations=True
					Exit For
				End If
			Next

	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "GetAllNodes"
			iNodeCount=ObjPropertyTree.GetROProperty("items count")
			For iCounter=0 To iNodeCount-1
				strCurrProperty=ObjPropertyTree.GetItem(iCounter)
				If iCounter<>iNodeCount-1 Then
					StrUpdatedProp=StrUpdatedProp+strCurrProperty+":"
				Else
					StrUpdatedProp=StrUpdatedProp+strCurrProperty
				End If
			Next
			Fn_SISW_Search_SimpleSearchPropertyTreeOprations=StrUpdatedProp
   End Select
   Set ObjPropertyTree=Nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_SISW_Search_SimpleSearchEditClauseOprations
'@@
'@@    Description				 :	Function Used to perform operations on Edit clauses of Simple Search
'@@
'@@    Parameters			   :	1.StrAction: Action Name
'@@											  2.StrOperator: Operator
'@@											  3.StrValue : Value
'@@													
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	Simple Search Tab Should be open							
'@@
'@@    Examples					:	Call Fn_SISW_Search_SimpleSearchEditClauseOprations("EnterCriteria","=","EditBox~Item")
'@@											 Call Fn_SISW_Search_SimpleSearchEditClauseOprations(Action Name,Operator,Control Type~Attached text (Name))
'@@											 Call Fn_SISW_Search_SimpleSearchEditClauseOprations("EnterCriteria","!=","RadioButton~True")
'@@											 Call Fn_SISW_Search_SimpleSearchEditClauseOprations("GetAllOperators","","")
'@@											 Call Fn_SISW_Search_SimpleSearchEditClauseOprations("GetValueToolTip","","")
'@@											 Call Fn_SISW_Search_SimpleSearchEditClauseOprations("EnterCriteria",">","Date~11-Aug-2011 11:00:21")
'@@											 Call Fn_SISW_Search_SimpleSearchEditClauseOprations("EnterCriteria","=","Date~01-Jan-2011 01:00:21")
'@@											 Call Fn_SISW_Search_SimpleSearchEditClauseOprations("EnterCriteria","=","DropDown~AutoTest1 (autotest1)")
'@@											bReturn= Fn_SISW_Search_SimpleSearchEditClauseOprations("VerifyInvalidDateExists&checkButton","","OK")
'@@											bReturn= Fn_SISW_Search_SimpleSearchEditClauseOprations("Type&SetDate","","12 03 12")
'@@											bReturn= Fn_SISW_Search_SimpleSearchEditClauseOprations("VerifyDay","","3")
'@@											bReturn= Fn_SISW_Search_SimpleSearchEditClauseOprations("VerifyDay","abc","11")
'@@											bReturn= Fn_SISW_Search_SimpleSearchEditClauseOprations("VerifyInvalidDateBlankEditBox","abc","")
'@@	   History					 	:	
'@@													Developer Name												Date						Rev. No.						Changes Done								Reviewer
'@@------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@												Sandeep Navghane									08-Aug-2011						1.0																								Sunny Ruparel
'@@												Sandeep Navghane									12-Aug-2011						1.1						Added Case "GetValueToolTip"				 Sunny Ruparel
'@@												Sandeep Navghane									12-Aug-2011						1.2						Modified Case "EnterCriteria"				       Sunny Ruparel
'@@												Sandeep Navghane									17-Aug-2011						1.3						Added Case "DropDown"				 			  Sunny Ruparel
'@@												Shreyas Waichal											  09-Sept-2012					 `.4					Added 4 new Cases								
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Function Fn_SISW_Search_SimpleSearchEditClauseOprations(StrAction,StrOperator,StrValue)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_Search_SimpleSearchEditClauseOprations"
   'Variable declaration
   Dim bFlag,ObjDefaultWnd,sDate,WshShell,iCnt,sValue,iCnter
   Dim arrValues,iOprCount,iCounter,StrCrrOperator,StrUpdatedOpr
   Dim iCount,iItemcnt,objDateControl,objShell
   Set ObjDefaultWnd=JavaWindow("DefaultWindow")
   Fn_SISW_Search_SimpleSearchEditClauseOprations=False
   bFlag=False
   Select Case StrAction
   ' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "EnterCriteria"
			'Selecting Operator 
			If StrOperator<>"" Then
				bFlag=Fn_UI_ListItemExist("Fn_SISW_Search_SimpleSearchEditClauseOprations", ObjDefaultWnd, "EditClauseOperator",StrOperator)
				If bFlag=True Then
'					Call Fn_List_Select("Fn_SISW_Search_SimpleSearchEditClauseOprations", ObjDefaultWnd, "EditClauseOperator",StrOperator)
					' *Added by Nilesh on 5-March-2013
					 iItemcnt=ObjDefaultWnd.JavaList("EditClauseOperator").GetRoProperty("items count")
					For iCount=0 to iItemcnt-1
						If  ObjDefaultWnd.JavaList("EditClauseOperator").GetItem(iCount)=StrOperator Then
								ObjDefaultWnd.JavaList("EditClauseOperator").Object.Select 	iCount
								bFlag=True
								Exit For
						Else
								bFlag=False
						End If
					Next
					If bFlag=False Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Operator " & StrOperator &"  is not set in Operators List")
						Set ObjDefaultWnd=Nothing
						Exit Function
					End If
					'*End
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Operator " & StrOperator &"  is not exist in Operators List")
					Set ObjDefaultWnd=Nothing
					Exit Function
				End If
			End If
			If StrValue<>"" Then
				arrValues=Split(StrValue,"~")
				Select Case arrValues(0)
					Case "EditBox"
							wait 1
							Call Fn_Edit_Box("Fn_SISW_Search_SimpleSearchEditClauseOprations",ObjDefaultWnd,"EditClauseValue",arrValues(1))
							wait 1
					Case "DropDown"
							Call Fn_Edit_Box("Fn_SISW_Search_SimpleSearchEditClauseOprations",ObjDefaultWnd,"EditClauseValue",arrValues(1))
							wait 1
							Set WshShell = CreateObject("WScript.Shell")
							wait 1
							WshShell.SendKeys "{ENTER}"
							Set WshShell =Nothing
					Case "RadioButton"
							ObjDefaultWnd.JavaRadioButton("EditClauseValue").SetTOProperty "attached text",arrValues(1)
							wait 2
							Call Fn_UI_JavaRadioButton_SetON("Fn_SISW_Search_SimpleSearchEditClauseOprations",ObjDefaultWnd, "EditClauseValue")
					Case "Date"
                          	sDate=Split(arrValues(1)," ")
                           JavaWindow("DefaultWindow").JavaEdit("EditClauseValue").Set sDate(0)
						   	wait 2
                           Set WshShell = CreateObject("WScript.Shell")
							WshShell.SendKeys "{TAB}"
							Set WshShell = Nothing
							wait 2
                           JavaWindow("DefaultWindow").JavaList("SimpleSearchCondition").Select sDate(1)
      

							'Call Fn_Button_Click("Fn_SISW_Search_SimpleSearchEditClauseOprations", JavaWindow("DefaultWindow"), "SimpleSearchDateSet")
						  'Call Fn_UI_SetDateAndTime("Fn_SISW_Search_SimpleSearchEditClauseOprations",sDate(0),sDate(1))
				End Select
			End If
			Fn_SISW_Search_SimpleSearchEditClauseOprations=True
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "GetAllOperators"
			iOprCount=Fn_UI_Object_GetROProperty("Fn_SISW_Search_SimpleSearchEditClauseOprations",ObjDefaultWnd.JavaList("EditClauseOperator"),"items count")
			For iCounter=0 To iOprCount-1
				StrCrrOperator=JavaWindow("DefaultWindow").JavaList("EditClauseOperator").GetItem(iCounter)
				If iCounter<>iOprCount-1 Then
					StrUpdatedOpr=StrUpdatedOpr+StrCrrOperator+":"
				Else
					StrUpdatedOpr=StrUpdatedOpr+StrCrrOperator
				End If
			Next
			Fn_SISW_Search_SimpleSearchEditClauseOprations=StrUpdatedOpr
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "GetValueToolTip"
			If JavaWindow("DefaultWindow").JavaEdit("EditClauseValue").Exist(SISW_MIN_TIMEOUT) Then
				Fn_SISW_Search_SimpleSearchEditClauseOprations=JavaWindow("DefaultWindow").JavaEdit("EditClauseValue").Object.getToolTipText()
			End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 

	Case "Type&SetDate"

'This is a pseudo code to determine to invoke Date Control
'Pass any Random String in "StrOperator" Parameter

				If StrOperator="" Then
						Call Fn_Button_Click("Fn_SISW_Search_SimpleSearchEditClauseOprations", JavaWindow("DefaultWindow"), "SimpleSearchDateSet")
						wait 2
				End If
				For iCnt = 1 to 30
						JavaWindow("MyTcShell").SetTOProperty "index", iCnt
						If JavaWindow("MyTcShell").JavaWindow("Date Control").Exist(SISW_MICRO_TIMEOUT) Then
								Set objDateControl=JavaWindow("MyTcShell").JavaWindow("Date Control")
								Exit For
						End If
				Next
				objDateControl.JavaEdit("Text").Set StrValue

				'Press Ok Button
				If StrOperator="" Then
				objDateControl.JavaButton("OK").Click micLeftBtn
				End if
				If err.number<0 Then
					Fn_SISW_Search_SimpleSearchEditClauseOprations=False
				Else
					Fn_SISW_Search_SimpleSearchEditClauseOprations=True
				End If

	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 

	Case "VerifyInvalidDateExists&checkButton"

'This is a pseudo code to determine to invoke Date Control
'Pass any Random String in "StrOperator" Parameter

				If StrOperator="" Then
						Call Fn_Button_Click("Fn_SISW_Search_SimpleSearchEditClauseOprations", JavaWindow("DefaultWindow"), "SimpleSearchDateSet")
						wait 2
				End If				
				For iCnt = 1 to 30
				JavaWindow("MyTcShell").SetTOProperty "index", iCnt
						If JavaWindow("MyTcShell").JavaWindow("Date Control").Exist(SISW_MIN_TIMEOUT) Then
								Set objDateControl=JavaWindow("MyTcShell").JavaWindow("Date Control")
								Exit For
						End If
				Next

				If objDateControl.JavaEdit("Text").Exist(SISW_MIN_TIMEOUT) Then
					objDateControl.JavaEdit("Text").Set ""
				Elseif objDateControl.JavaEdit("InvalidDate").Exist(SISW_MICRO_TIMEOUT) then
					objDateControl.JavaEdit("InvalidDate").Set ""
				End If

				Set objShell=CreateObject("WScript.Shell")
				objShell.SendKeys "{TAB}"
				
				For iCnt = 1 to 30
				JavaWindow("MyTcShell").SetTOProperty "index", iCnt
						If JavaWindow("MyTcShell").JavaWindow("Date Control").Exist(SISW_MICRO_TIMEOUT) Then
								Set objDateControl=JavaWindow("MyTcShell").JavaWindow("Date Control")
								Exit For
						End If
				Next

				If objDateControl.JavaStaticText("Invalid").Exist Then
					Fn_SISW_Search_SimpleSearchEditClauseOprations=true
				Else
					Fn_SISW_Search_SimpleSearchEditClauseOprations=false
				End If

				If  StrValue<>"" Then
					sValue=objDateControl.JavaButton("OK").GetROProperty("enabled")
					If cint(sValue)=0 Then
						Fn_SISW_Search_SimpleSearchEditClauseOprations=true
					Else
						Fn_SISW_Search_SimpleSearchEditClauseOprations=false
					End If
			End If

				If StrOperator="" Then  'Code added by Anjali on 13th Sep 2012
					objDateControl.Close
				End if

'		End if
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 

	Case "VerifyDay"

'This is a pseudo code to determine to invoke Date Control
'Pass any Random String in "StrOperator" Parameter

				If StrOperator="" Then
							Call Fn_Button_Click("Fn_SISW_Search_SimpleSearchEditClauseOprations", JavaWindow("DefaultWindow"), "SimpleSearchDateSet")
							wait 2
				End If
				For iCnt = 1 to 30
				JavaWindow("MyTcShell").SetTOProperty "index", iCnt
						If JavaWindow("MyTcShell").JavaWindow("Date Control").Exist(SISW_MICRO_TIMEOUT) Then
								Set objDateControl=JavaWindow("MyTcShell").JavaWindow("Date Control")
								Exit For
						End If
				Next
				If objDateControl.JavaCalendar("Date").Exist(SISW_MIN_TIMEOUT)  Then
					sValue=objDateControl.JavaCalendar("Date").GetROProperty("day")
				Elseif objDateControl.JavaCalendar("InvalidDate").Exist(SISW_MICRO_TIMEOUT) then
					sValue=objDateControl.JavaCalendar("InvalidDate").GetROProperty("day")
				End If
				If cstr(sValue)=cstr(StrValue) Then
						Fn_SISW_Search_SimpleSearchEditClauseOprations=true
						If StrOperator="" Then
						objDateControl.Close
						End if
				Else
						Fn_SISW_Search_SimpleSearchEditClauseOprations=false
				End If

'	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'
	Case "VerifyDateInEditBox"

'This is a pseudo code to determine to invoke Date Control
'Pass any Random String in "StrOperator" Parameter

				If StrOperator="" Then
							Call Fn_Button_Click("Fn_SISW_Search_SimpleSearchEditClauseOprations", JavaWindow("DefaultWindow"), "SimpleSearchDateSet")
							wait 2
				End If	

				For iCnt = 1 to 30
				JavaWindow("MyTcShell").SetTOProperty "index", iCnt
						If JavaWindow("MyTcShell").JavaWindow("Date Control").Exist(SISW_MICRO_TIMEOUT) Then
								Set objDateControl=JavaWindow("MyTcShell").JavaWindow("Date Control")
								Exit For
						End If
				Next
				If objDateControl.JavaEdit("Text").Exist(SISW_MIN_TIMEOUT) Then
					sValue=objDateControl.JavaEdit("Text").GetROProperty("value")
				Elseif objDateControl.JavaEdit("InvalidDate").Exist(SISW_MICRO_TIMEOUT) then
					sValue=objDateControl.JavaEdit("InvalidDate").GetROProperty("value")
				End If
				If cstr(sValue)=cstr(StrValue) Then
						Fn_SISW_Search_SimpleSearchEditClauseOprations=true
						If StrOperator="" Then
						objDateControl.Close
						End if
				Else
						Fn_SISW_Search_SimpleSearchEditClauseOprations=false
				End If

	Case "VerifyInvalidDateBlankEditBox"

			'This is a pseudo code to determine to invoke Date Control
			'Pass any Random String in "StrOperator" Parameter

				If StrOperator="" Then
							Call Fn_Button_Click("Fn_SISW_Search_SimpleSearchEditClauseOprations", JavaWindow("DefaultWindow"), "SimpleSearchDateSet")
							wait 2
				End If
				For iCnt = 1 to 30
				JavaWindow("MyTcShell").SetTOProperty "index", iCnt
						If JavaWindow("MyTcShell").JavaWindow("Date Control").Exist(SISW_MICRO_TIMEOUT) Then
								Set objDateControl=JavaWindow("MyTcShell").JavaWindow("Date Control")
								Exit For
						End If
				Next

				objDateControl.JavaEdit("Text").SetTOProperty "attached text","Invalid date"
				If objDateControl.JavaEdit("Text").Exist Then
					sValue=objDateControl.JavaEdit("Text").GetROProperty("caret_position")
					If cint(sValue)=0 Then
							If StrOperator="" Then
								objDateControl.Close
							End if
							Fn_SISW_Search_SimpleSearchEditClauseOprations=True
					Else
						Fn_SISW_Search_SimpleSearchEditClauseOprations=false
					End If
				End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	Case "VerifyInvalidDateExists&checkButtonAfterClick"

			'This is a pseudo code to determine to invoke Date Control
			'Pass any Random String in "StrOperator" Parameter
				If StrOperator="" Then
							Call Fn_Button_Click("Fn_SISW_Search_SimpleSearchEditClauseOprations", JavaWindow("DefaultWindow"), "SimpleSearchDateSet")
							wait 1
				End If

				For iCnt = 1 to 30
				JavaWindow("MyTcShell").SetTOProperty "index", iCnt
						If JavaWindow("MyTcShell").JavaWindow("Date Control").Exist(1) Then
								Set objDateControl=JavaWindow("MyTcShell").JavaWindow("Date Control")
								Exit For
						End If
				Next
				Set objShell=CreateObject("WScript.Shell")
				sValue=objDateControl.JavaEdit("Text").GetROProperty("value")
				objDateControl.JavaEdit("Text").Click 0,0,"LEFT"
				Wait 1
				For iCnter=0 to len(sValue)-1
						objShell.SendKeys "{DEL}"
				Next
					
				objDateControl.JavaButton("OK").Click micLeftBtn

				For iCnt = 1 to 30
				JavaWindow("MyTcShell").SetTOProperty "index", iCnt
						If JavaWindow("MyTcShell").JavaWindow("Date Control").Exist(SISW_MICRO_TIMEOUT) Then
								Set objDateControl=JavaWindow("MyTcShell").JavaWindow("Date Control")
								Exit For
						End If
				Next

				If objDateControl.JavaStaticText("Invalid").Exist Then
					Fn_SISW_Search_SimpleSearchEditClauseOprations=true
				Else
					Fn_SISW_Search_SimpleSearchEditClauseOprations=false
				End If

				If  StrValue<>"" Then
					sValue=objDateControl.JavaButton("OK").GetROProperty("enabled")
					If cint(sValue)=0 Then
						Fn_SISW_Search_SimpleSearchEditClauseOprations=true
					Else
						Fn_SISW_Search_SimpleSearchEditClauseOprations=false
					End If
					objDateControl.Close
			End If
'		End if

   End Select
	Set ObjDefaultWnd=Nothing
	Set objDateControl=Nothing
	Set objShell=Nothing
End Function

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_SISW_Search_SimpleSearchCriteriaTableOprations
'@@
'@@    Description				 :	Function Used to perform operations on Selected Search Criteria Table
'@@
'@@    Parameters			   :	1.StrAction: Action Name
'@@											  2.StrProperty: Proprty Name
'@@											  3.StrOperator: Operator
'@@											  2.StrValue: Value
'@@
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	Simple Search Tab Should be open							
'@@
'@@    Examples					:	Call Fn_SISW_Search_SimpleSearchCriteriaTableOprations("Select","Name","=","Item")
'@@											 Call Fn_SISW_Search_SimpleSearchCriteriaTableOprations("Verify","Configuration Item?","!=","TRUE")
'@@											 Call Fn_SISW_Search_SimpleSearchCriteriaTableOprations("Remove","Name","=","Item")
'@@											 Call Fn_SISW_Search_SimpleSearchCriteriaTableOprations("Search","","","")
'@@											 Call Fn_SISW_Search_SimpleSearchCriteriaTableOprations("Remove","","","")
'@@											 Call Fn_SISW_Search_SimpleSearchCriteriaTableOprations("RemoveAll","Name","=","Item")
'@@											 Call Fn_SISW_Search_SimpleSearchCriteriaTableOprations("ChangeCriteria","AND~Description","=","TestItem")
'@@											 Call Fn_SISW_Search_SimpleSearchCriteriaTableOprations("GetLineCriteria","Description","Contain","item")
'@@											 Call Fn_SISW_Search_SimpleSearchCriteriaTableOprations("SetCriteria","aaa~Description","=","TestItem")
'@@											 Call Fn_SISW_Search_SimpleSearchCriteriaTableOprations("GetRowNumber","Description","=","TestItem")
'@@											 Call Fn_SISW_Search_SimpleSearchCriteriaTableOprations("Up","3~Description","=","TestItem")
'@@											 Call Fn_SISW_Search_SimpleSearchCriteriaTableOprations("Down","1~Description","=","TestItem")
'@@
'@@	   History					 	:	
'@@													Developer Name												Date						Rev. No.						Changes Done								Reviewer
'@@------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@												Sandeep Navghane									08-Aug-2011						1.0																								Sunny Ruparel
'@@												Sandeep Navghane									16-Aug-2011						1.1						Added Case "ChangeCriteria"						Sunny Ruparel
'@@												Sandeep Navghane									16-Aug-2011						1.2						Added Case "GetLineCriteria"					 Sunny Ruparel
'@@												Sandeep Navghane									17-Aug-2011						1.3						Added Case "SetCriteria"						     Sunny Ruparel
'@@												Sandeep Navghane									17-Aug-2011						1.4						Added Case "GetRowNumber"					  Sunny Ruparel
'@@																																																				  "Up"
'@@																																																				   "Down"
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Function Fn_SISW_Search_SimpleSearchCriteriaTableOprations(StrAction,StrProperty,StrOperator,StrValue)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_Search_SimpleSearchCriteriaTableOprations"
	'Variable Declaration
    Dim ObjSearchCriteriaTbl
    Dim iRowCount,iCounter,strCurrProperty,strCurrOpr,strCurrVal,bFlag,arrProperty
	'Creating Object "SimpleSearchSelectedSearchCriteria" Table
   Set ObjSearchCriteriaTbl=JavaWindow("DefaultWindow").JavaTable("SimpleSearchSelectedSearchCriteria")
   Fn_SISW_Search_SimpleSearchCriteriaTableOprations=False
   bFlag=False
	Select Case StrAction
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Select"
			iRowCount=Fn_UI_Object_GetROProperty("Fn_SISW_Search_SimpleSearchCriteriaTableOprations",ObjSearchCriteriaTbl, "rows")
			For iCounter=0 To iRowCount-1
				strCurrProperty=ObjSearchCriteriaTbl.GetCellData(iCounter,"Property")
				If strCurrProperty=StrProperty Then
					strCurrOpr=ObjSearchCriteriaTbl.GetCellData(iCounter,"Operator")
					strCurrVal=ObjSearchCriteriaTbl.GetCellData(iCounter,"Value")
					If Trim(strCurrOpr+strCurrVal)=Trim(StrOperator+StrValue) Then
						ObjSearchCriteriaTbl.SelectRow iCounter
						wait 1
						Fn_SISW_Search_SimpleSearchCriteriaTableOprations=True
						Exit For
					End If
				End If
			Next
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Verify"
			iRowCount=Fn_UI_Object_GetROProperty("Fn_SISW_Search_SimpleSearchCriteriaTableOprations",ObjSearchCriteriaTbl, "rows")
			For iCounter=0 To iRowCount-1
				strCurrProperty=ObjSearchCriteriaTbl.GetCellData(iCounter,"Property")
				If strCurrProperty=StrProperty Then
					strCurrOpr=ObjSearchCriteriaTbl.GetCellData(iCounter,"Operator")
					strCurrVal=ObjSearchCriteriaTbl.GetCellData(iCounter,"Value")
					If Trim(strCurrOpr+CStr(strCurrVal))=Trim(StrOperator+StrValue) Then
						Fn_SISW_Search_SimpleSearchCriteriaTableOprations=True
						Exit For
					End If
				End If
			Next
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Search"
			bFlag=Fn_Button_Click("Fn_SISW_Search_SimpleSearchCriteriaTableOprations", JavaWindow("DefaultWindow"), "SimpleSearch")
			If bFlag=True Then
				Fn_SISW_Search_SimpleSearchCriteriaTableOprations=True
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Remove"
			If StrProperty<>"" And StrOperator<>"" And StrValue<>"" Then
				bFlag=Fn_SISW_Search_SimpleSearchCriteriaTableOprations("Select",StrProperty,StrOperator,StrValue)
				If bFlag=False Then
					Set ObjSearchCriteriaTbl=Nothing
					Exit Function
				End If
			End If
			Call Fn_Button_Click("Fn_SISW_Search_SimpleSearchCriteriaTableOprations", JavaWindow("DefaultWindow"), "SimpleSearchRemove")
			Fn_SISW_Search_SimpleSearchCriteriaTableOprations=True
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "RemoveAll"
			bFlag=Fn_Button_Click("Fn_SISW_Search_SimpleSearchCriteriaTableOprations", JavaWindow("DefaultWindow"), "SimpleSearchRemoveAll")
			If bFlag=True Then
				Fn_SISW_Search_SimpleSearchCriteriaTableOprations=True
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "ChangeCriteria"
			arrProperty=Split(StrProperty,"~")
			iRowCount=Fn_UI_Object_GetROProperty("Fn_SISW_Search_SimpleSearchCriteriaTableOprations",ObjSearchCriteriaTbl, "rows")
			For iCounter=0 To iRowCount-1
				strCurrProperty=ObjSearchCriteriaTbl.GetCellData(iCounter,"Property")
				If strCurrProperty=arrProperty(1) Then
					strCurrOpr=ObjSearchCriteriaTbl.GetCellData(iCounter,"Operator")
					strCurrVal=ObjSearchCriteriaTbl.GetCellData(iCounter,"Value")
					If Trim(strCurrOpr+strCurrVal)=Trim(StrOperator+StrValue) Then
						ObjSearchCriteriaTbl.SelectCell iCounter,"0"
						wait 1
						JavaWindow("DefaultWindow").JavaList("SimpleSearchCondition").Select arrProperty(0)
						wait 1
						ObjSearchCriteriaTbl.SelectCell iCounter,"1"
						Fn_SISW_Search_SimpleSearchCriteriaTableOprations=True
						Exit For
					End If
				End If
			Next
        ' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "GetLineCriteria"
			iRowCount=Fn_UI_Object_GetROProperty("Fn_SISW_Search_SimpleSearchCriteriaTableOprations",ObjSearchCriteriaTbl, "rows")
				For iCounter=0 To iRowCount-1
					strCurrProperty=ObjSearchCriteriaTbl.GetCellData(iCounter,"Property")
					If strCurrProperty=StrProperty Then
						strCurrOpr=ObjSearchCriteriaTbl.GetCellData(iCounter,"Operator")
						strCurrVal=ObjSearchCriteriaTbl.GetCellData(iCounter,"Value")
						If Trim(strCurrOpr+strCurrVal)=Trim(StrOperator+StrValue) Then
							Fn_SISW_Search_SimpleSearchCriteriaTableOprations=ObjSearchCriteriaTbl.GetCellData(iCounter,"")
							Exit For
						End If
					End If
				Next
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "SetCriteria"
			arrProperty=Split(StrProperty,"~")
			iRowCount=Fn_UI_Object_GetROProperty("Fn_SISW_Search_SimpleSearchCriteriaTableOprations",ObjSearchCriteriaTbl, "rows")
			For iCounter=0 To iRowCount-1
				strCurrProperty=ObjSearchCriteriaTbl.GetCellData(iCounter,"Property")
				If strCurrProperty=arrProperty(1) Then
					strCurrOpr=ObjSearchCriteriaTbl.GetCellData(iCounter,"Operator")
					strCurrVal=ObjSearchCriteriaTbl.GetCellData(iCounter,"Value")
					If Trim(strCurrOpr+strCurrVal)=Trim(StrOperator+StrValue) Then
						ObjSearchCriteriaTbl.SetCellData iCounter,"0",arrProperty(0)
						wait 1
						ObjSearchCriteriaTbl.SelectCell iCounter,"1"
						Fn_SISW_Search_SimpleSearchCriteriaTableOprations=True
						Exit For
					End If
				End If
			Next
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "GetRowNumber"
			iRowCount=Fn_UI_Object_GetROProperty("Fn_SISW_Search_SimpleSearchCriteriaTableOprations",ObjSearchCriteriaTbl, "rows")
				For iCounter=0 To iRowCount-1
					strCurrProperty=ObjSearchCriteriaTbl.GetCellData(iCounter,"Property")
					If strCurrProperty=StrProperty Then
						strCurrOpr=ObjSearchCriteriaTbl.GetCellData(iCounter,"Operator")
						strCurrVal=ObjSearchCriteriaTbl.GetCellData(iCounter,"Value")
						If Trim(strCurrOpr+strCurrVal)=Trim(StrOperator+StrValue) Then
							Fn_SISW_Search_SimpleSearchCriteriaTableOprations=iCounter
							Exit For
						End If
					End If
				Next	
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Up"
			arrProperty=Split(StrProperty,"~")
			bFlag=Fn_SISW_Search_SimpleSearchCriteriaTableOprations("Select",arrProperty(1),StrOperator,StrValue)
			If bFlag=True Then
				For iCounter=1 To CInt(arrProperty(0))
					Call Fn_Button_Click("Fn_SISW_Search_SimpleSearchCriteriaTableOprations", JavaWindow("DefaultWindow"), "SimpleSearchUp")
					wait 1
					Fn_SISW_Search_SimpleSearchCriteriaTableOprations=True
				Next
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Down"
			arrProperty=Split(StrProperty,"~")
			bFlag=Fn_SISW_Search_SimpleSearchCriteriaTableOprations("Select",arrProperty(1),StrOperator,StrValue)
			If bFlag=True Then
				For iCounter=1 To CInt(arrProperty(0))
					Call Fn_Button_Click("Fn_SISW_Search_SimpleSearchCriteriaTableOprations", JavaWindow("DefaultWindow"), "SimpleSearchDown")
					wait 1
					Fn_SISW_Search_SimpleSearchCriteriaTableOprations=True
				Next
			End If
	End Select
	Set ObjSearchCriteriaTbl=Nothing
End Function

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_SISW_Search_SearchSortOperations

'Description			 :	Function Used to perform operations Sort Option in Search criteria pane 

'Parameters			   :   1.StrAction: Action Name
'										2.StrSortBy: Sort By option
'										3.StrOrderBy: Order By Option
'										4.StrColumn: Column Name
'										5.StrValue: Expected value
'										6.StrButtonName: Button Name
'
'Return Value		   : 	True or False

'Pre-requisite			:	should be exist on search pane

'Examples				:   bReturn=Fn_SISW_Search_SearchSortOperations("GetSortOptionIndex","ID~Is VI?~Name~Type","","Cancel")
'										bReturn=Fn_SISW_Search_SearchSortOperations("Select","ID","","","","")
'										bReturn=Fn_SISW_Search_SearchSortOperations("Verify","ID","","Sort By","ID","")
'										bReturn=Fn_SISW_Search_SearchSortOperations("Verify","ID","","Order By","None","")
'
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												13-Dec-2012								1.0																						Pranav S
'													Sandeep N												20-Dec-2012								1.1							Added Case : Select,Verify		 Sukhada B
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Public Function Fn_SISW_Search_SearchSortOperations(StrAction,StrSortBy,StrOrderBy,StrColumn,StrValue,StrButtonName)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_Search_SearchSortOperations"
	'Declaring variables
	Dim objSortDialog
	Dim arrSortBy,iRows,iCounter,iCount,bFlag,iIndex

	Fn_SISW_Search_SearchSortOperations=False
	'Creating object of [ Sort ] dialog
   Set	objSortDialog=JavaWindow("DefaultWindow").JavaWindow("Sort")
	'checking existance of [ Sort ] dialog
	If not objSortDialog.Exist(SISW_MIN_TIMEOUT) Then
		'clicking on Sort toolbar button to invoke [ Sort ] dialog
		Call Fn_ToolbatButtonClick("Sort")
		wait 2
	End if

   Select Case StrAction
	    '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	 	Case "Select"
			bFlag=False
			iRows=objSortDialog.JavaTable("Table").GetROProperty("rows")
			For iCounter=0 to iRows-1
				If trim(objSortDialog.JavaTable("Table").GetCellData(iCounter,"Sort By"))=trim(StrSortBy) Then
					objSortDialog.JavaTable("Table").SelectRow iCounter
					wait 1
					bFlag=True		
					Exit for
				End if
			Next
			If bFlag=True Then
				Fn_SISW_Search_SearchSortOperations=True
			End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Verify"
			bFlag=False
			iRows=objSortDialog.JavaTable("Table").GetROProperty("rows")
			For iCounter=0 to iRows-1
				If trim(objSortDialog.JavaTable("Table").GetCellData(iCounter,"Sort By"))=trim(StrSortBy) Then
					If trim(objSortDialog.JavaTable("Table").GetCellData(iCounter,StrColumn))=trim(StrValue) Then
                    	bFlag=True		
					End if
					Exit for
				End if
			Next
			If bFlag=True Then
				Fn_SISW_Search_SearchSortOperations=True
			End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
        Case "GetSortOptionIndex"
			'splitin Sort By option
			arrSortBy=Split(StrSortBy,"~")
			'Retriving number of rows exist ing table
			iRows=objSortDialog.JavaTable("Table").GetROProperty("rows")
			For iCounter=0 to ubound(arrSortBy)
				bFlag=False
				For iCount=0 to iRows-1
					'matching current Sort By option with expected option
					If trim(objSortDialog.JavaTable("Table").GetCellData(iCount,"Sort By"))=trim(arrSortBy(iCounter)) Then
						If iCounter=0 Then
							iIndex=CStr(iCount)
						Else
							iIndex=CStr(iIndex)+":"+CStr(iCount)
						End If
						bFlag=True
						Exit for
					End If
				Next
				If bFlag=False Then
					Exit for
				End If
			Next
	If bFlag=True Then
		Fn_SISW_Search_SearchSortOperations=iIndex
	End If
    		
   End Select
	'Click on button
	If StrButtonName<>"" Then
		Call Fn_Button_Click( "Fn_SISW_Search_SearchSortOperations", objSortDialog,StrButtonName)
	End If
End Function


'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'  FUNCTION NAME   :     Fn_SISW_Search_ClassificationSearchDialogOperations
'
'  DESCRIPTION     :   		This function is used to perform operations on  'Classification Search Dialog' Operations
'
'  PARAMETERS      :  		strAction -  valid action name
'                    					 		dicClassificationSearchDialog-  valid dictionary object
'                            					StrButton-  valid button name

'
'  EXAMPLE            :    					dicClassificationSearchDialog("ObjectID")="000064"
'																With dicClassificationSearchDialog
'																	.Add "ObjectID",""
'																	.Add "Search",""
'																End With
'													1. Fn_SISW_Search_ClassificationSearchDialogOperations("Search",dicClassificationSearchDialog,"OK")
		
				'										 dicClassificationSearchDialog("ObjectID")="000064"
'																With dicClassificationSearchDialog
'																	.Add "ObjectID",""
'																	.Add "Search",""
'																'End With
'													2.  Fn_SISW_Search_ClassificationSearchDialogOperations("TableTabOperation",dicClassificationSearchDialog,"OK")
'													3.  Fn_SISW_Search_ClassificationSearchDialogOperations("AddComponentID",dicClassificationSearchDialog,"")
'													4.  Fn_SISW_Search_ClassificationSearchDialogOperations("ClearAllValue",dicClassificationSearchDialog,"")
'													5.  Fn_SISW_Search_ClassificationSearchDialogOperations("AddRevisionRule",dicClassificationSearchDialog,"")
'		
' History : 
'								
'          Developer Name         Date      	Rev. No.      Changes Done      																Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'          Shailendra  Sahu     13/05/2013       1.0    		           																		Pranav
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'		   Ankit Nigam			14/01/2016		 1.1		Added Cases "DialogExist","SelectClassandAddClassifiedObject"		[TC1122-2016010600-14_01_2016-AnkitN-Maintenance]
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'		   Ankit Nigam			14/01/2016		 1.1		Added Cases "SelectClassandAddICOwithoutItem"						[TC1122-2016010600-14_01_2016-AnkitN-Maintenance]
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Function Fn_SISW_Search_ClassificationSearchDialogOperations(strAction,dicClassificationSearchDialog,StrButton)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_Search_ClassificationSearchDialogOperations"
	Dim objClassificationSearchDialog,TempValue,intCount
	Dim aClassPath, sClassPath,iCnt
	Dim aValue, iCount, bFlag, aGetValues, sTextVal

	'Added by VivekA --------------------------------------
	If Instr(lCase(JavaWindow("DefaultWindow").GetROProperty("title")), "manufacturing process planner")>0 then
		If lcase (dicClassificationSearchDialog("InvokeOption")) = "toolbar" then 
			Call Fn_ToolBarOperation("Click", "Assign resource from library/classification","")
		End if
	End if
	'------------------------------------------------------
	'Added by VivekA --------------------------------------[TC1015-2015081100-08_09_2015-VivekA-NewDevelopment]
	If Instr(JavaWindow("DefaultWindow").GetROProperty("title"), "Resource Manager")>0 Then
		If Fn_SISW_Search_GetObject("ClassificationSearchDialog3").Exist Then
			Set objClassificationSearchDialog = Fn_SISW_Search_GetObject("ClassificationSearchDialog3")
		End If
	ElseIf Fn_SISW_Search_GetObject("ClassificationSearchDialog1").Exist Then
		Set objClassificationSearchDialog = Fn_SISW_Search_GetObject("ClassificationSearchDialog1")
	ElseIf Fn_SISW_Search_GetObject("ClassificationSearchDialog2").Exist Then
		Set objClassificationSearchDialog = Fn_SISW_Search_GetObject("ClassificationSearchDialog2")
	End If

	Set objClassificationSearchDialog = Fn_UI_ObjectCreate("Fn_GeneralItem_OptionsSettings", objClassificationSearchDialog)
	'------------------------------------------------------
	
	If Not objClassificationSearchDialog.Exist(SISW_MIN_TIMEOUT) Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Classification Search Dialog box' is not present")
		Fn_SISW_Search_ClassificationSearchDialogOperations=false
		Exit Function
	End If
		
	Select Case strAction
			Case "DialogExist" 'Added by VivekA ---------------------------------
				If Not objClassificationSearchDialog.Exist(SISW_MIN_TIMEOUT) Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Classification Search Dialog box' is not present")
					Fn_SISW_Search_ClassificationSearchDialogOperations=false
					Exit Function
				Else
					If  Fn_SISW_UI_JavaButton_Operations("Fn_SISW_Search_ClassificationSearchDialogOperations", "Click", objClassificationSearchDialog, "Cancel") = false then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed To click on Cancel Button")
						Exit Function
					End If
					Fn_SISW_Search_ClassificationSearchDialogOperations=True
				End If
				' ---------------------------------------------------------------
			Case "Search"
				Call Fn_SISW_UI_JavaTab_Operations("Fn_SISW_Search_ClassificationSearchDialogOperations", "Select", objClassificationSearchDialog, "AllTabs", "Search")	
				' Writing Object ID
				objClassificationSearchDialog.JavaEdit("ObjectID").Set ""
				Call Fn_Edit_Box("Fn_SISW_Search_ClassificationSearchDialogOperations",objClassificationSearchDialog,"ObjectID",dicClassificationSearchDialog("ObjectID"))
				Call Fn_ReadyStatusSync(1)
				'Clicking on Find Button
				Call Fn_Button_Click("Fn_SISW_Search_ClassificationSearchDialogOperations",objClassificationSearchDialog,"find_16")
				Call Fn_ReadyStatusSync(1)
				Fn_SISW_Search_ClassificationSearchDialogOperations=True
'=================================================================================================================
			Case "TableTabOperation"
                    If  objClassificationSearchDialog.JavaTab("AllTabs").GetROProperty("enabled")="1" Then
						'Clicking on Table tab
                        Call Fn_ReadyStatusSync(1)
						Call Fn_UI_JavaTab_Select("",objClassificationSearchDialog,"AllTabs","Table")
                        Call Fn_ReadyStatusSync(1)
						'Checking for list of material
						'If objClassificationSearchDialog.JavaTable("materialTable").GetROProperty("rows")>0 Then
						If Fn_UI_Object_GetROProperty("Fn_SISW_Search_ClassificationSearchDialogOperations",objClassificationSearchDialog.JavaTable("materialTable"), "rows") >0 Then
							For intCount=0 to objClassificationSearchDialog.JavaTable("materialTable").GetROProperty("rows")
								TempValue=Fn_UI_JavaTable_GetCellData("",objClassificationSearchDialog,"materialTable",intCount,2)
								If isNumeric(TempValue) Then
									If Cint(dicClassificationSearchDialog("ObjectID"))=CInt(TempValue) Then
'										 Call Fn_UI_JavaTable_SelectRow("Fn_SISW_Search_ClassificationSearchDialogOperations",objClassificationSearchDialog,"materialTable",intCount)
										wait 1
										objClassificationSearchDialog.JavaTable("materialTable").SelectRowsRange intCount,intCount
										wait 1
										 Exit For
									End If
								Else
									If dicClassificationSearchDialog("ObjectID")=TempValue Then
'										 Call Fn_UI_JavaTable_SelectRow("Fn_SISW_Search_ClassificationSearchDialogOperations",objClassificationSearchDialog,"materialTable",intCount)
										wait 1
										objClassificationSearchDialog.JavaTable("materialTable").SelectRowsRange intCount,intCount
										wait 1
										 Exit For
									End If
								End If
							Next
							If intCount<>objClassificationSearchDialog.JavaTable("materialTable").GetROProperty("rows") Then
								'Clicking on Ok button
								 Call Fn_Button_Click("Fn_SISW_Search_ClassificationSearchDialogOperations",objClassificationSearchDialog,StrButton)
								 Call Fn_ReadyStatusSync(1)
							Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Object ID not found in the List")
								Exit Function
							End If
						End If
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Table tab is not enabled")
						Exit Function
					End If
					Fn_SISW_Search_ClassificationSearchDialogOperations=True
'=================================================================================================================
			Case"AddComponentID"
				If objClassificationSearchDialog.JavaButton("add_16").GetROProperty("enabled")="1" Then
						Call Fn_Button_Click("Fn_SISW_Search_ClassificationSearchDialogOperations",objClassificationSearchDialog,"add_16")
						Call Fn_ReadyStatusSync(1)
				End If
				Fn_SISW_Search_ClassificationSearchDialogOperations=True
'=================================================================================================================
			Case "ClearAllValue"
				If objClassificationSearchDialog.JavaButton("Clear").GetROProperty("enabled")="1" Then
						Call Fn_Button_Click("Fn_SISW_Search_ClassificationSearchDialogOperations",objClassificationSearchDialog,"Clear")
						Call Fn_ReadyStatusSync(1)
				End If
					Fn_SISW_Search_ClassificationSearchDialogOperations=True
'=================================================================================================================
			Case "AddRevisionRule"
				Call Fn_UI_JavaStaticText_Click("Fn_SISW_Search_ClassificationSearchDialogOperations",objClassificationSearchDialog,"ClicktoAddaRevision","0","0","LEFT")
				   	Fn_SISW_Search_ClassificationSearchDialogOperations=True
'=================================================================================================================
			Case "MapClassificationObject"

''=================================================================================================================
			Case "SelectClassandAddClassifiedObject", "SelectClassandAddICOwithoutItem"	'Added by VivekA ---------------------------------------
				Fn_SISW_Search_ClassificationSearchDialogOperations=false
				If dicClassificationSearchDialog("ClassPath") <> "" Then
					aClassPath = Split(dicClassificationSearchDialog("ClassPath") , ":")
					For iCnt = 0 To Ubound(aClassPath)- 1 Step 1
						If iCnt = 0 Then
							sClassPath = aClassPath(iCnt)
						Else
							sClassPath = sClassPath &":" & aClassPath(iCnt)
						End If
						If Fn_UI_JavaTree_Expand("Fn_SISW_Search_ClassificationSearchDialogOperations", objClassificationSearchDialog, "Hierarchy",sClassPath)= False Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed To EXPAND Class in Hierarchy Tree")
							Exit Function	
						End If 
					Next
					If Fn_JavaTree_Node_Activate("Fn_SISW_Search_ClassificationSearchDialogOperations",objClassificationSearchDialog,"Hierarchy",dicClassificationSearchDialog("ClassPath")) = False then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Class is is not found in Hierarchy Tree")
						Exit Function
					End if
				End if
				wait 2
				Call Fn_SISW_UI_JavaTab_Operations("Fn_SISW_Search_ClassificationSearchDialogOperations", "Select", objClassificationSearchDialog, "AllTabs", "Search")
				If  Fn_SISW_UI_JavaButton_Operations("Fn_SISW_Search_ClassificationSearchDialogOperations", "Click", objClassificationSearchDialog, "Search") = false then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed To click on Search Button")
					Exit Function
				End If
				wait 2
				If dicClassificationSearchDialog("ObjectID")<> "" then
					Call Fn_SISW_UI_JavaTab_Operations("Fn_SISW_Search_ClassificationSearchDialogOperations", "Select", objClassificationSearchDialog, "AllTabs", "Table")
					wait 2
					
'					For intCount=0 to objClassificationSearchDialog.JavaTable("materialTable").GetROProperty("rows")-1
					For intCount=0 to Fn_UI_Object_GetROProperty("Fn_SISW_Search_ClassificationSearchDialogOperations",objClassificationSearchDialog.JavaTable("materialTable"), "rows")-1
						TempValue=Fn_UI_JavaTable_GetCellData("",objClassificationSearchDialog,"materialTable",intCount,2)
						If isNumeric(TempValue) Then
							If cLng(dicClassificationSearchDialog("ObjectID"))=cLng(TempValue) Then
								wait 2
								objClassificationSearchDialog.JavaTable("materialTable").SelectRowsRange intCount,intCount
								wait 1
'									objClassificationSearchDialog.JavaTable("materialTable").SelectRow  
								 Exit For
							End If
						Else
							If dicClassificationSearchDialog("ObjectID")=TempValue Then
								wait 2
								objClassificationSearchDialog.JavaTable("materialTable").SelectRowsRange intCount,intCount
								wait 1
'									objClassificationSearchDialog.JavaTable("materialTable").SelectRow 
								Exit For
							End If
						End If
					Next
				End if
				If dicClassificationSearchDialog("OccurrenceType")<> "" then
					If Fn_SISW_UI_JavaList_Operations("Fn_SISW_Search_ClassificationSearchDialogOperations", "Select", objClassificationSearchDialog, "OccurrenceType", dicClassificationSearchDialog("OccurrenceType"), "", "") = False then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Select Occurrence Type")
						Exit Function
					End if
				End if
				
				If StrButton <> "" Then 
					If Fn_SISW_UI_JavaButton_Operations("Fn_SISW_Search_ClassificationSearchDialogOperations", "Click", objClassificationSearchDialog, StrButton) = false then
					    Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed To click on OK Button")
					   Exit Function
				      End iF
				End If
				
				If strAction ="SelectClassandAddICOwithoutItem" Then
					If  Fn_UI_ObjectExist("Fn_SISW_Search_ClassificationSearchDialogOperations", JavaDialog("ICOwithoutItem")) = True Then
						If dicClassificationSearchDialog("IcoItemType") <> "" Then
							If Fn_SISW_UI_JavaList_Operations("Fn_SISW_Search_ClassificationSearchDialogOperations", "Select", JavaDialog("ICOwithoutItem"), "ItemType", dicClassificationSearchDialog("IcoItemType"), "", "") = False then
								call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Select Occurrence Type")
								Exit Function
							End iF
						End If
	
						If Fn_SISW_UI_JavaButton_Operations("Fn_SISW_Search_ClassificationSearchDialogOperations", "Click", JavaDialog("ICOwithoutItem"), "OK") = false then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed To click on OK Button")
							Exit Function
						End iF
						JavaDialog("Error").SetTOProperty "title","Connect item with classification object"
						If  Fn_UI_ObjectExist("Fn_SISW_Search_ClassificationSearchDialogOperations", JavaDialog("Error")) = True Then
							If Fn_SISW_UI_JavaButton_Operations("Fn_SISW_Search_ClassificationSearchDialogOperations", "Click", JavaDialog("Error"), "Yes") = false then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed To click on Yes Button")
								Exit Function
							End iF
						End if
					Else
						Fn_SISW_Search_ClassificationSearchDialogOperations=False
						Exit Function
					End If
				End If
				Fn_SISW_Search_ClassificationSearchDialogOperations=True	'-----------------------------------
''=================================================================================================================
			Case "CreateGraphics","UpdateGrapics"
'			
''=================================================================================================================
			Case "PropertyOfSelectedData"
				  If  objClassificationSearchDialog.JavaTab("AllTabs").GetROProperty("enabled")="1" Then
						'Clicking on Table tab
						objClassificationSearchDialog.JavaTab("AllTabs").Select("Table")
						'Checking for list of material
						If objClassificationSearchDialog.JavaTable("materialTable").GetROProperty("rows")>0 Then
							For intCount=0 to objClassificationSearchDialog.JavaTable("materialTable").GetROProperty("rows")
								TempValue=Fn_UI_JavaTable_GetCellData("Fn_SISW_Search_ClassificationSearchDialogOperations",objClassificationSearchDialog,"materialTable",intCount,2)
								If Cint(dicClassificationSearchDialog("ObjectID"))=CInt(TempValue) Then
									 Call Fn_UI_JavaTable_SelectRow("",objClassificationSearchDialog,"materialTable",intCount)
									 Exit For
								End If
							Next
							'Clicking on Ok button
							Call Fn_Button_Click("Fn_SISW_Search_ClassificationSearchDialogOperations",objClassificationSearchDialog,"Properties")
                             Call Fn_ReadyStatusSync(1)
						End If
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Table tab is not enabled")
						Exit Function
					End If
					Fn_SISW_Search_ClassificationSearchDialogOperations=True
			''=================================================================================================================
			Case "CopyToOSClipboard"
			''=================================================================================================================
			Case "VerifyPropertiesOfSelectedData"		'[TC1015-08_09_2015-2015082000-VivekA-NewDevelopment] - Added for Properties tab verification
				'Select Properties tab 
				If dicClassificationSearchDialog("ObjectRevID")<> "" Then
					Call Fn_SISW_UI_JavaTab_Operations("Fn_SISW_Search_ClassificationSearchDialogOperations", "Select", objClassificationSearchDialog, "AllTabs", "Properties")
					'get value from object ID static text 
					objClassificationSearchDialog.JavaStaticText("ObjectIdText").SetTOProperty "label","<html><body>Object ID:&nbsp;&nbsp;<b>"+dicClassificationSearchDialog("ObjectRevID")+"</b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<i>"+dicClassificationSearchDialog("Class")+"</i></html><body>"
					TempValue = objClassificationSearchDialog.JavaStaticText("ObjectIdText").GetROProperty("attached text")
					If instr(1,cstr(TempValue),cstr(dicClassificationSearchDialog("ObjectRevID"))) > 0 Then
						Fn_SISW_Search_ClassificationSearchDialogOperations=True
					Else
						Fn_SISW_Search_ClassificationSearchDialogOperations=False					
					End If
				End If 
				
				If StrButton <> "" Then 
					If Fn_SISW_UI_JavaButton_Operations("Fn_SISW_Search_ClassificationSearchDialogOperations", "Click", objClassificationSearchDialog, StrButton) = false then
					    Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed To click on OK Button")
					   Exit Function
				    End iF
				End If
			''=================================================================================================================
			Case "VerifyAttributeValuesFromProperties"		'[TC1015-10_08_2015-2015092200-VivekA-NewDevelopment] - Added for Properties tab verification
				'Select Properties tab 
				Call Fn_SISW_UI_JavaTab_Operations("Fn_SISW_Search_ClassificationSearchDialogOperations", "Select", objClassificationSearchDialog, "AllTabs", "Properties")
				
				If dicClassificationSearchDialog("sAttributeType")<> "" Then
					Select Case dicClassificationSearchDialog("sAttributeType")
						Case "List"
							If Instr(1,dicClassificationSearchDialog("sAttributeValue"),",") Then
								aValue = Split(dicClassificationSearchDialog("sAttributeValue"),"~",-1,1)
							Else
								aValue = Array(dicClassificationSearchDialog("sAttributeValue"))
							End If
							For iCount = 0 to Ubound(aValue)
								bFlag = False
								aGetValues = split(aValue(iCount),":",-1,1)
								objClassificationSearchDialog.JavaStaticText("AttributeName_Label").SetTOProperty "label", aGetValues(0)
								Wait 1
								sTextVal = objClassificationSearchDialog.JavaList("AttributeValueList").GetROProperty("value")
								Wait 1
								If Trim(Cstr(aGetValues(1))) = Trim(Cstr(sTextVal)) Then
									bFlag = True
								Else															
									Fn_SISW_Search_ClassificationSearchDialogOperations = False
									Set objClassificationSearchDialog = Nothing
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Verify value [" + aGetValues(1) + "] for Attribute [" + aGetValues(0) + "] " )
									Exit Function
								End If
							Next
							If bFlag = True Then
								Fn_SISW_Search_ClassificationSearchDialogOperations = True
							End If
					End Select					
				End If 
				
				If StrButton <> "" Then 
					If Fn_SISW_UI_JavaButton_Operations("Fn_SISW_Search_ClassificationSearchDialogOperations", "Click", objClassificationSearchDialog, StrButton) = false then
					    Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed To click on OK Button")
					   Exit Function
					   Set objClassificationSearchDialog = Nothing
				    End IF
				End If
				'-------------------------------------------
		End Select
	Set objClassificationSearchDialog = nothing
End Function

'-------------------------------------------------------------------------Load a given Search Query-------------------- ------------------------------------------------------------------------------------------
' Function Name			:	  Fn_SISW_Search_LoadQuery()

'Pre-Requisite				:	  RAC Session accessible and  Application loaded
' 				
'Description				  :		Load a given Search Query
'										
'Parameters			 		:    strQueryPath: Tree Path of the Named Query       											
'
'Return Value		        : 	True \ False
'
' Examples				      :	 Fn_SISW_Search_LoadQuery("System Defined Searches:Change Request Revision...")		
'										
'History				           :	Developer Name			Date			Version				Changes Done			Reviewer	
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N				07-10-10		1.0																	Sunny R
'													Sandeep N				20-01-11		1.0																	Sunny R
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_Search_LoadQuery(strQueryPath)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_Search_LoadQuery"
	Dim ObjJavaTree, aQryPath

			aQryPath = Split(strQueryPath, ":", -1, 1)
			'Click on [Open Search View] Icon on the Toolbar
			If True = Fn_MenuOperation("Select","Window:Show View:Other...") Then
					Call Fn_SetView ("Teamcenter:Search")
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Clicked : Open Search View ToolBar Button")					
			Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Click on Open Search View ToolBar Button")
					Fn_SISW_Search_LoadQuery = False
			End IF
            Call Fn_ReadyStatusSync(1)
			'Click on [Select a Search] toolbar button under Search Criteria Panel
			If True = Fn_ToolbatButtonClick("Select a Search") Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Clicked : Select a Search Button")					
			Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Click on [Select a Search] Button")
					Fn_SISW_Search_LoadQuery = False
			End IF
			Call Fn_ReadyStatusSync(1)
			'#	Select search criteria listed [Change Search] dialog & Selecting Requested Node
        	Set ObjJavaTree = Fn_UI_ObjectCreate("Fn_SISW_Search_LoadQuery",JavaWindow("DefaultWindow").JavaWindow("Change Search"))
			Wait 3
			Call Fn_UI_JavaTree_Expand("Fn_SISW_Search_LoadQuery", ObjJavaTree, "SearchOptions",aQryPath(0))
             If True = Fn_JavaTree_Select("Fn_SISW_Search_LoadQuery", ObjJavaTree, "SearchOptions",strQueryPath) Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Selected Requested Node: ["+strQueryPath+"] ")					
			Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Select Requested Node: ["+strQueryPath+"] ")			
					Fn_SISW_Search_LoadQuery = False
			End IF
			'#	Invoke [OK] button
        	If True = Fn_Button_Click("Fn_SISW_Search_LoadQuery", ObjJavaTree, "OK")  Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Clicked : [Open Search View] Button")					
			Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Click on [Open Search View] Button")
					Fn_SISW_Search_LoadQuery = False
			End IF
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Loaded the Query")
			Fn_SISW_Search_LoadQuery = True
	Set ObjJavaTree  = Nothing
End Function

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name :	Fn_MyTcSrch_CompareReport_Operation
'@@
'@@    Description	 :	Function Used to Perform operations on Compare report dialog
'@@
'@@    Parameters	 :	1. sAction		: Action to be performed
'@@					 :	2. dicDetails	: Dictionary object
'@@					 :	3. sButton 		: Button name
'@@
'@@    Return Value	 : 	True Or False
'@@
'@@    Pre-requisite :	Search Result Tree should be opened.
'@@
'@@    Examples		 : 	Set dicDetails = CreateObject("Scripting.Dictionary")
'@@    									dicDetails("NodeName") = "Item ID (3)"
'@@    									dicDetails("CompareToMenu") = "Compare To:Item Revision... (1)"
'@@    									dicDetails("LeftTreeDifference") = "0 Difference"
'@@    									dicDetails("RightTreeDifference") = "0 Difference"
'@@    									dicDetails("BottomText") = "0"
'@@    									bReturn = Fn_MyTcSrch_CompareReport_Operation("Compare",dicDetails,"Close")		
'@@	   History		 :	
'@@		Developer Name		Date	  		Rev. No.	Changes Done										Reviewer
'@@-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@		Vivek Ahirrao	 27-May-2016		1.0		  	Created for Search TC's								[TC1122-20160504-27_05_2016-VivekA-NewDevelopment]
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_MyTcSrch_CompareReport_Operation(sAction,dicDetails,sButton)
	GBL_FAILED_FUNCTION_NAME="Fn_MyTcSrch_CompareReport_Operation"
	Dim objCompareReportDlg
	Dim bFlag
	
	Fn_MyTcSrch_CompareReport_Operation = False
	
	If JavaWindow("MyTeamcenter_Search").JavaWindow("MyTc").JavaDialog("CompareReport").Exist = False Then
		If dicDetails("NodeName")<>"" AND dicDetails("CompareToMenu")<>"" Then
			bFlag = Fn_MyTc_SrchResltTreeOperation("PopupMenuSelect",dicDetails("NodeName"),dicDetails("CompareToMenu"))
			If bFlag = False Then
				Exit Function
			End If
		End If
	End If
	
	Set objCompareReportDlg = JavaWindow("MyTeamcenter_Search").JavaWindow("MyTc").JavaDialog("CompareReport")
	
	If objCompareReportDlg.Exist(2) Then
		'Check if Warning dialog is present or not "The views being compared are completely different."
		objCompareReportDlg.SetTOProperty "Index","2"
		If objCompareReportDlg.Exist = False Then
			'Warning dialog is present, so Click on OK button
			objCompareReportDlg.SetTOProperty "Index","1"
			Call Fn_Button_Click("Fn_MyTcSrch_CompareReport_Operation",objCompareReportDlg,"OK")
		End If
		'Check if "The views being compared are identical. Do you want to continue?" dialog is present
		objCompareReportDlg.SetTOProperty "Index","0"
		If objCompareReportDlg.JavaButton("Yes").Exist(2) Then
			Call Fn_Button_Click("Fn_MyTcSrch_CompareReport_Operation",objCompareReportDlg,"Yes")
		End If
		Wait 2
	End If
	
	'Now operation on Main dialog of Compare Report
	Select Case sAction
		Case "Compare"
				'Verify Left Tree Static Text
				If dicDetails("LeftTreeDifference")<>"" Then
					objCompareReportDlg.JavaStaticText("DifferenceNumberTxt").SetTOProperty "Index","0"
					objCompareReportDlg.JavaStaticText("DifferenceNumberTxt").SetTOProperty "label",dicDetails("LeftTreeDifference")
					If objCompareReportDlg.JavaStaticText("DifferenceNumberTxt").Exist = False Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Verify [Left Tree Text Difference] is ["+dicDetails("LeftTreeDifference")+"].")
						Set objCompareReportDlg = Nothing
						Fn_MyTcSrch_CompareReport_Operation = False
						Exit Function
					End If
				End If
				
				'Verify Left Tree Static Text
				If dicDetails("RightTreeDifference")<>"" Then
					objCompareReportDlg.JavaStaticText("DifferenceNumberTxt").SetTOProperty "Index","1"
					objCompareReportDlg.JavaStaticText("DifferenceNumberTxt").SetTOProperty "label",dicDetails("RightTreeDifference")
					If objCompareReportDlg.JavaStaticText("DifferenceNumberTxt").Exist = False Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Verify [Right Tree Text Difference] is ["+dicDetails("RightTreeDifference")+"].")
						Set objCompareReportDlg = Nothing
						Fn_MyTcSrch_CompareReport_Operation = False
						Exit Function
					End If
				End If
				
				'Verify Bottom Static Text of total Differnet objects
				If dicDetails("BottomText")<>"" Then
					objCompareReportDlg.JavaStaticText("TotalDifferencesTxt").SetTOProperty "label",dicDetails("BottomText") & " total different object(s) found."
					If objCompareReportDlg.JavaStaticText("TotalDifferencesTxt").Exist = False Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Verify [Bottom Text Total Difference] is ["+dicDetails("BottomText")+"].")
						Set objCompareReportDlg = Nothing
						Fn_MyTcSrch_CompareReport_Operation = False
						Exit Function
					End If
				End If
				Fn_MyTcSrch_CompareReport_Operation = True
				
			Case Else
					Fn_MyTcSrch_CompareReport_Operation = False
					Set objCompareReportDlg = Nothing
					Exit Function
	End Select
	Set objCompareReportDlg = Nothing
End Function

'#######################################################################################################################################################~
'########################################################  Extension function to perform Local Query operation with defined inputs as specified     ################################################~
'#
'# FUNCTION NAME:	 	  Fn_QryBldr_LocalQuery_Operations
'#
'# MODULE: 						 Search Requirement, MyTC 
'#
'# PRE-REQUISITE:		RAC Session accessible and [Query Builder] Application loaded
'#
'# DESCRIPTION:			 Create a Local Query with defined inputs as specified
'#											 1. Input [Name] field details
'#											 2. Input [Description] field details
'#											 3. Click [Search Class] button
'#														 3a. Tree navigate to trace [Class Attribute], For Example: POM_object:WorkspaceObject:ItemRevision
'#														 3b. Select/Double Click on Tree Node of prefered [Class Attribute]
'#														 3b. Close [Class Attribute] window
'#											 4. Check the [Search Class] button label updated to requiste class
'#											 5. Set the [Display Setting] to required option [Class/All Attributes]
'#											 6. Set [Show Indented Results] option [On/Off]
'#											 7. Select attribute from [Attribute Selection] Tree
'#													>>	 Note: 	<<
'#													 7.1 Please note that the function arguement is a array of required attributes, seperated by ":"
'#													 7.2 Select the required attribute iteratively and click on [+] button to add the attribute
'#											 8. Invoke [Create] button
'#										
'# PARAMETERS   :      					   sAction:Name of the action to be performed
'#										   sQueryName: Name of the Local User Query
'#										   sQueryDescription: Description of the Local User Query
'#										   sSearchClass: Attribute Class POM Object of the Query
'#									       sDisplaySettings: Display Settings [Class/All Attributes]
'#										   bShowIndentedResults: [Show Indented Results] option [On/Off]
'#
'#									       aAttributes: Array of the class attributes to be added to the Search Query
'#													 >>   Note: ( 1.) Multiple Attributes to be seperated by "~" ( Tilde)  [ EXAMPLE - >> Attrib1~Attrib2~Attrib3]
'#  																   ( 2.) Inside Attribute  inner values to be seperated by "," (Comma)  [ EXAMPLE - >> First, Second, Third]
'#  																  ( 3.) Inside Values Reference Path  to be seperated by ":" (Colon)  [ EXAMPLE - >> Dataset:Revision
'#
'#				             	     								>>	InnerValues  = "First, Second, Third"
'#																						 First =  refpath for Attrib in main window - will activate /double click
'#				  				   								 				Second =  [For class Attrib Sel Dialog ] ->> FullRefPath Attrib to be selected.				 (Set the Class)   				
'#																									[For Class Selection Dialog} ->>EditBoxValue]
'#																					Third = Refpath for Attrib in main window - will activate /double click
'#
'#												->>	 (Multiple Attribute) Example>>  (First, Second, Third ~ First, Second, Third~ First, Second, Third)  << -
'#												->>	 (First, Second, Third ) Example>>  Refpath1,RefPath2,Refpath3<< -
'#												->>	 (Refpath1 ) Example>>  Home:Child << -
'#											sButton: Button Name
'#
'# RETURN VALUE : 		TRUE \ FALSE
'#
'# Examples	:				aAttributes = "Select|Item:Type~SelectAndActivate|Item:Referenced By,SelectClassAttribute|ImanRelation:Primary Reference,SelectAndActivate|Item:Primary Reference [ ImanRelation ]:Relation Type [ ImanType ],Select|Item:Primary Reference [ ImanRelation ]:Relation Type [ ImanType ]:Name~SelectAndActivate|Item:Primary Reference [ ImanRelation ]:Secondary Reference,SelectClassAttribute|Item,Select|Item:Primary Reference [ ImanRelation ]:Secondary Reference [ Item ]:ID"
'#							Fn_QryBldr_LocalQuery_Operations("Create","Query_Create_564", "", "Item", "AllAttributes", "", aAttributes,"")  
'#
'#	History	:						Developer Name			Date			Version				Changes Done			Reviewer	
'#------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'#									Poonam Chopade			29-08-2017		1.0					Created					Swapna Ghatge
'#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'########################################################  Extension function to perform Local Query operation with defined inputs as specified     ################################################~

Public Function Fn_QryBldr_LocalQuery_Operations(sAction, sQueryName, sQueryDescription, sSearchClass, sDisplaySettings, bShowIndentedResults, aAttributes, sButton)  
		GBL_FAILED_FUNCTION_NAME="Fn_QryBldr_LocalQuery_Operations"
		Dim  ArrDispSet,  OuterArrAttrib, iOuterCounter, ArrAttrib,  ArrinnAttrib, iCounter
		Dim ObjQryApp, ObjQryAttribSel
		Dim jCnt, iCnt, aDummy, sPath
		Dim itemCnt, iTreeCnt, sRevisionRule

		Set ObjQryApp =  Fn_UI_ObjectCreate( "Fn_QryBldr_LocalQuery_Operations",JavaWindow("Search_QueryBuilder").JavaWindow("QueryBuilder"))

		Select Case sAction
			Case "Create"
				'++++++++++<<    Chech weather the clear button enable or not   >>++++++++++
				If ObjQryApp.JavaButton("Clear").GetROProperty("enabled") = 1  Then
						Call Fn_Button_Click( "Fn_QryBldr_LocalQuery_Operations", ObjQryApp, "Clear")
				End If
		
			'++++++++++<<    Set the Local Query   >>++++++++++
				Call Fn_List_Select("Fn_QryBldr_LocalQuery_Operations", ObjQryApp , "ModifiableQueryTypes", "Local Query")
		
			'++++++++++<<    Input [Name] field details >>++++++++++
				Call Fn_Edit_Box("Fn_QryBldr_LocalQuery_Operations",ObjQryApp,"Name",sQueryName)
		
			'++++++++++<<   Input [Description] field details>>++++++++++
				'[TC1122-20160504-26_05_2016-VivekA-NewDevelopment] - Added for Search new TCs
				If Instr(sQueryDescription,"$") Then
					aQueryDescription = Split(sQueryDescription,"$")
					Call Fn_Edit_Box("Fn_QryBldr_LocalQuery_Operations",ObjQryApp,"Description",aQueryDescription(0))
					'Code to set Revision Rule value in List
					sRevisionRule = aQueryDescription(1)
				Else
					Call Fn_Edit_Box("Fn_QryBldr_LocalQuery_Operations",ObjQryApp,"Description",sQueryDescription)		
				End If
		
			'++++++++++<<   Click [Search Class] button >>++++++++++
				Call Fn_CheckBox_Set("Fn_QryBldr_LocalQuery_Operations", ObjQryApp, "SrchClass", "ON")
				Call Fn_Edit_Box("Fn_QryBldr_LocalQuery_Operations",ObjQryApp,"Class/Attribute Selection",sSearchClass)
				Call Fn_Button_Click( "Fn_QryBldr_LocalQuery_Operations", ObjQryApp, "Find")
				Call Fn_ReadyStatusSync(5)
				ObjQryApp.JavaObject("Close").Click 1,1
				wait(1)
		
			'++++++++++<<   Set Revision Rule >>++++++++++
				If sRevisionRule<>"" Then
					'Code to set Revision Rule value in List
					Call Fn_List_Select("Fn_QryBldr_LocalQuery_Operations",ObjQryApp,"RevisionRule",sRevisionRule)
					Wait 0,200
				End if
			 '++++++++++<<  Set the [Display Setting] to required option [Class/All Attributes] >>++++++++++
			 If sDisplaySettings<>"" Then
					ArrDispSet = split(sDisplaySettings, ":", -1,1)
					If Ubound(ArrDispSet) = 1 Then
						 Wait(1)
						Call Fn_CheckBox_Set("Fn_QryBldr_LocalQuery_Operations", ObjQryApp, "DisplaySettings", "ON")
						Wait(1)
						Call  Fn_UI_JavaRadioButton_SetON("Fn_QryBldr_LocalQuery_Operations",ObjQryApp, ArrDispSet(0))
						Wait(1)
						Call  Fn_UI_JavaRadioButton_SetON("Fn_QryBldr_LocalQuery_Operations",ObjQryApp, ArrDispSet(1))
					Else
						Wait(1)
						Call Fn_CheckBox_Set("Fn_QryBldr_LocalQuery_Operations", ObjQryApp, "DisplaySettings", "ON")
						Wait(1)
						Call  Fn_UI_JavaRadioButton_SetON("Fn_QryBldr_LocalQuery_Operations",ObjQryApp, sDisplaySettings )
					End If
					ObjQryApp.JavaObject("Close").Click 1,1
			End If
		
			 '++++++++++<<  Set [Show Indented Results] option [On/Off] >>++++++++++
				Call Fn_CheckBox_Set("Fn_QryBldr_LocalQuery_Operations", ObjQryApp, "ShowIndentedResults", bShowIndentedResults)
				wait 1
			'++++++++++<<   Select attribute from [Attribute Selection] Tree >>++++++++++
				OuterArrAttrib = split(aAttributes, "~", -1,1)
				For iOuterCounter = 0 To Ubound(OuterArrAttrib)
						ArrAttrib = split(OuterArrAttrib(iOuterCounter), ",", -1, 1)
						For iCounter = 0 to Ubound(ArrAttrib) 
									ArrAction = split(ArrAttrib(iCounter), "|", -1, 1)			
									ArrinnAttrib = split(ArrAction(1), ":", -1, 1)
		'							For iCounter = 0 to Ubound(ArrinnAttrib)
										Select Case ArrAction(0)
												Case "Select"
													aDummy = split(ArrAction(1),":")
													If uBound(aDummy) <> 0 Then
														For iCnt = 0 to ubound(aDummy) -1
															sPath = ""
															For jCnt = 0 to iCnt
																If sPath = "" Then
																	sPath = aDummy(jCnt)
																Else
																	sPath = sPath & ":" & aDummy(jCnt)
																End If
															Next
															If iCnt = 0 Then
																itemCnt = cInt(ObjQryApp.JavaTree("AttributeSelectionList").getROProperty("items count"))
																If itemCnt = 1 Then
																	ObjQryApp.JavaTree("AttributeSelectionList").Select sPath
																	ObjQryApp.JavaTree("AttributeSelectionList").Object.setExpandedState ObjQryApp.JavaTree("AttributeSelectionList").Object.getSelectionPath(), true
																End If
															Else
																ObjQryApp.JavaTree("AttributeSelectionList").Select sPath
																ObjQryApp.JavaTree("AttributeSelectionList").Object.setExpandedState ObjQryApp.JavaTree("AttributeSelectionList").Object.getSelectionPath(), true
																' special scenario tree node get expanded after dbl clicking on it.
																itemCnt = cInt(ObjQryApp.JavaTree("AttributeSelectionList").GetROProperty("items count"))
																For iTreeCnt = 0 to itemCnt - 1
																	If ObjQryApp.JavaTree("AttributeSelectionList").GetItem(iTreeCnt) = sPath then
																		If not (instr(ObjQryApp.JavaTree("AttributeSelectionList").GetItem(iTreeCnt + 1), sPath) > 0 ) Then
																			ObjQryApp.JavaTree("AttributeSelectionList").Activate sPath 
																			wait 2
																		End If
																		Exit for
																	End if 
																Next
															End If
														Next
														wait 5
														if uBound(ArrAttrib) = 0 then
															If aDummy(uBound(aDummy))="Gov Classification" Then
																aDummy(uBound(aDummy))="Government Classification"
															End If
															ObjQryApp.JavaTree("AttributeSelectionList").Select sPath & ":" & aDummy(uBound(aDummy))
														else
															ObjQryApp.JavaTree("AttributeSelectionList").Select sPath & ":" & aDummy(uBound(aDummy))
														end If	
													Else
														ObjQryApp.JavaTree("AttributeSelectionList").Select ArrAttrib(iCounter)
													End If
													
												Case "SelectAndActivate"
													aDummy = split(ArrAction(1),":")
													If uBound(aDummy) <> 0 Then
														For iCnt = 0 to ubound(aDummy) -1
															sPath = ""
															For jCnt = 0 to iCnt
																If sPath = "" Then
																	sPath = aDummy(jCnt)
																Else
																	sPath = sPath & ":" & aDummy(jCnt)
																End If
															Next
															If iCnt = 0 Then
																itemCnt = cInt(ObjQryApp.JavaTree("AttributeSelectionList").getROProperty("items count"))
																If itemCnt = 1 Then
																	ObjQryApp.JavaTree("AttributeSelectionList").Select sPath
																	ObjQryApp.JavaTree("AttributeSelectionList").Object.setExpandedState ObjQryApp.JavaTree("AttributeSelectionList").Object.getSelectionPath(), true
																End If
															Else
																ObjQryApp.JavaTree("AttributeSelectionList").Select sPath
																ObjQryApp.JavaTree("AttributeSelectionList").Object.setExpandedState ObjQryApp.JavaTree("AttributeSelectionList").Object.getSelectionPath(), true
																' special scenario tree node get expanded after dbl clicking on it.
																itemCnt = cInt(ObjQryApp.JavaTree("AttributeSelectionList").GetROProperty("items count"))
																For iTreeCnt = 0 to itemCnt - 1
																	If ObjQryApp.JavaTree("AttributeSelectionList").GetItem(iTreeCnt) = sPath then
																		If not (instr(ObjQryApp.JavaTree("AttributeSelectionList").GetItem(iTreeCnt + 1), sPath) > 0 ) Then
																			ObjQryApp.JavaTree("AttributeSelectionList").Activate sPath 
																			wait 2
																		End If
																		Exit for
																	End if 
																Next
															End If
														Next
														wait 5
														if uBound(ArrAttrib) = 0 then
															If aDummy(uBound(aDummy))="Gov Classification" Then
																aDummy(uBound(aDummy))="Government Classification"
															End If
															ObjQryApp.JavaTree("AttributeSelectionList").Select sPath & ":" & aDummy(uBound(aDummy))
															ObjQryApp.JavaTree("AttributeSelectionList").Activate sPath & ":" & aDummy(uBound(aDummy))
														else
															ObjQryApp.JavaTree("AttributeSelectionList").Select sPath & ":" & aDummy(uBound(aDummy))
															ObjQryApp.JavaTree("AttributeSelectionList").Activate sPath & ":" & aDummy(uBound(aDummy))
														end If	
													Else
														ObjQryApp.JavaTree("AttributeSelectionList").Select ArrAttrib(iCounter)
														ObjQryApp.JavaTree("AttributeSelectionList").Activate ArrAttrib(iCounter)
													End If
			
												Case "SelectClassAttribute"
													If  JavaWindow("Search_QueryBuilder").JavaWindow("QueryBuilder").JavaDialog("ClassAttributeSelection").exist(5) Then
														Set ObjQryAttribSel = JavaWindow("Search_QueryBuilder").JavaWindow("QueryBuilder").JavaDialog("ClassAttributeSelection") 
														Wait(2)
														Call Fn_CheckBox_Set("Fn_QryBldr_LocalQuery_Operations", ObjQryAttribSel, "CAS_SrchClass", "ON")
														Wait(2)
														Call Fn_Edit_Box("Fn_QryBldr_LocalQuery_Operations",ObjQryApp,"Class/Attribute Selection",ArrinnAttrib(0) )
														Wait(2)
														Call Fn_Button_Click( "Fn_QryBldr_LocalQuery_Operations", ObjQryApp, "Find")
														Wait(2)
														ObjQryApp.JavaObject("Close").Click 1,1
														Wait (10)
														ObjQryAttribSel.JavaTree("CAS_SearchClassTree").Activate ArrAction(1)
													End If
													If JavaWindow("Search_QueryBuilder").JavaWindow("QueryBuilder").JavaDialog("ClassSelectionDialog").exist(5) OR JavaDialog("ClassSelectionDialog").Exist(5) Then
														If JavaWindow("Search_QueryBuilder").JavaWindow("QueryBuilder").JavaDialog("ClassSelectionDialog").exist(5) Then
															Set ObjQryAttribSel = JavaWindow("Search_QueryBuilder").JavaWindow("QueryBuilder").JavaDialog("ClassSelectionDialog")
														ElseIf JavaDialog("ClassSelectionDialog").exist(5) Then
															Set ObjQryAttribSel =JavaDialog("ClassSelectionDialog")
														End If
														Wait 5
														Call Fn_Edit_Box("Fn_QryBldr_LocalQuery_Operations",ObjQryAttribSel,"SelectionField",ArrAction(1) )
														wait 5
														Call Fn_Button_Click( "Fn_QryBldr_LocalQuery_Operations", ObjQryAttribSel, "CSDFind")
														wait 5
														Call Fn_Button_Click( "Fn_QryBldr_LocalQuery_Operations", ObjQryAttribSel, "CSDOK")
													End If
									End Select							
							Next
							Call Fn_Button_Click( "Fn_QryBldr_LocalQuery_Operations", ObjQryApp, "Add")
						Next
		
				 '++++++++++<<  Invoke [Create] button >>++++++++++
				 If  aAttributes <> "" Then
					Call Fn_Button_Click( "Fn_QryBldr_LocalQuery_Operations", ObjQryApp, "Create")
					Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Pass: Local Query Successfully Created. ")   	
					Fn_QryBldr_LocalQuery_Operations = True
				Else
					Call Fn_Button_Click( "Fn_QryBldr_LocalQuery_Operations", ObjQryApp, "Clear")
					Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Pass: Local Query Successfully Tested. ")   	
					Fn_QryBldr_LocalQuery_Operations = True
				End If
				Set ObjQryApp = Nothing
				Set ObjQryAttribSel = Nothing
		 End Select
End Function
'#************************************************************************************************************************************************************************
'#	Function Name		:				Fn_SISW_Search_QuickSearchAndVerifyError
'#
'#	Description	        :		 	    Verify Serach error message when object not found
'#
'#	Parameters			:           	1.sSrchType: Name of the query to select
'#									    2.sSrchText : Search value
'#									    3.sErrMsg : Error message to verify	
'#
'#	Return Value		: 		   		TRUE \ FALSE
'#
'#	Examples		   :		  	   Call Fn_SISW_Search_QuickSearchAndVerifyError("Item Name", "TestItemName","You have returned zero objects count")
'#
'#	History				:
'#									Developer Name			Date			Rev. No.		Changes Done			Reviewer		Reviewed Date
'#************************************************************************************************************************************************************************
'#									Poonam Chopade		26-Sept-2017		1.0				Created				Tc11.4(2017091200)_NewDevelopment_PoonamC_26Sept2017
'#************************************************************************************************************************************************************************
Public Function Fn_SISW_Search_QuickSearchAndVerifyError(sSrchType, sSrchText,sErrMsg)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_Search_QuickSearchAndVerifyError"
	Dim  sSrchWild,sSrch,sAppMsg

    JavaWindow("DefaultWindow").JavaToolbar("QuickSearchToolbar").ShowDropdown "Perform Search"
	wait 1
	'Added by VivekA -----------------------------------------------
	sSrchWild = ""
	If instr( 1 , sSrchText , "~" , 1 ) > 0 Then
		sSrch = split(sSrchText, "~")
		sSrchText =  sSrch(0)	'Item id 
		sSrchWild = sSrch(1)	'wild character	
	End If	
	'JavaWindow("DefaultWindow").WinMenu("ContextMenu").Select sSrchType
	' -- added by Koustubh
	If sSrchType = "StringID" Then sSrchType = "Item ID"
	JavaWindow("DefaultWindow").JavaMenu("Label:=" & sSrchType).Select
	wait 2
	'-------------------------------------------------------------------
	'Added by VivekA -----------------------------------------------
	If sSrchWild <> "" Then
		JavaWindow("DefaultWindow").JavaEdit("QuickSearch").Set sSrchWild
	Else 
		JavaWindow("DefaultWindow").JavaEdit("QuickSearch").Set sSrchText	
	End If
	'-----------------------------------------------------------------------
	wait 3
	JavaWindow("DefaultWindow").JavaToolbar("QuickSearchToolbar").Press "Perform Search"
	wait 3
	If JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("NoAccessibleObjects").Exist(SISW_MIN_TIMEOUT) Then
		sAppMsg = JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("NoAccessibleObjects").JavaEdit("JTextArea").GetROProperty("value")
		If instr(trim(sAppMsg),trim(sErrMsg)) > 0 Then
			Fn_SISW_Search_QuickSearchAndVerifyError = True
		Else
			Fn_SISW_Search_QuickSearchAndVerifyError = False
		End If
		JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("NoAccessibleObjects").JavaButton("OK").Click
		Wait 2
	Else
		Fn_SISW_Search_QuickSearchAndVerifyError = False	
	End If
End Function

