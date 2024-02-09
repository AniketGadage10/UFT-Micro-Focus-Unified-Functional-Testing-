Option Explicit
'Function List
'************************************************************************************************************************************************************************************************************
'000. Fn_SISW_DC_GetObject()
'001. Fn_DC_ContextDefinitionOperations()
'002. Fn_DC_ProductItemsOperations()
'003. Fn_DC_WorkPartsOperations()
'004. Fn_DC_EngChangeRevisionOperations()
'005. Fn_DC_ProcessesOperations()
'006. Fn_DC_OccurrenceNotesSearchPanelOperations()
'007. Fn_DC_SearchResultOperations()
'008. Fn_DC_CloseDesignContext()
'009. Fn_DC_ItemIDSearchPanelOperations()
'010. Fn_DC_ZoneOperations()
'011. Fn_DC_SpatialSearchPanelOperations()
'012. Fn_SISW_DC_SaveStructureContextObject()
'013. Fn_SISW_DC_FormAttributeSearchPanelOperations()
'014. Fn_SISW_DC_ExcuteSCOSearch()
'*********************************************************	Function List		***********************************************************************
'****************************************    Function to get Object hierarchy ***************************************
'
''Function Name		 	:	Fn_SISW_DC_GetObject
'
''Description		    :  	Function to get specified Object hierarchy.

''Parameters		    :	1. sObjectName : Object Handle name
								
''Return Value		    :  	Object \ Nothing
'
''Examples		     	:	Fn_SISW_DC_GetObject("DesignContextApplet")

'History:
'	Developer Name			Date			Rev. No.		Reviewer		Changes Done	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Koustubh Watwe		 26-Jul-2012		1.0	
'-----------------------------------------------------------------------------------------------------------------------------------
'	Ashwini Kumar		 25-Sept-2013		2.0								Externalized object hierarchies
'-----------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_DC_GetObject(sObjectName)
	Dim sObjectXMLPath
	sObjectXMLPath = Fn_GetEnvValue("User", "AutomationDir") & "\TestData\AutomationXML\ObjectXML\DesignContext.xml"
	Set Fn_SISW_DC_GetObject = Fn_SISW_Setup_GetObjectFromXML(sObjectXMLPath, sObjectName)
End Function
'************************************************************************************************************************************************************************************************************
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_DC_ContextDefinitionOperations
'@@
'@@    Description			:	Function Used to perform operations on Context Definitions
'@@
'@@    Parameters			:	1. sAction		: Action to be performed
'@@    							2. sObjectRow	: Item node name from Object Column
'@@    							3. sColumn		: Column Name
'@@    							4. sValue		: Value to be verified. 
'@@
'@@    Return Value		   	: 	True Or False
'@@
'@@    Pre-requisite		:	Design Context perspective should be displayed							
'@@
'@@    Examples				:	Call Fn_DC_ContextDefinitionOperations("Select", "000023/A;1-top", "", "")
'@@    						:	Call Fn_DC_ContextDefinitionOperations("Remove", "000024/A;1-t2", "", "")
'@@    						:	Call Fn_DC_ContextDefinitionOperations("Verify", "000024/A;1-t2", "Description", "Product")
'@@
'@@	   History					 	:	
'@@				Developer Name				Date			Rev. No.	Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe			15-Dec-2011			1.0			Created
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_DC_ContextDefinitionOperations(sAction, sObjectRow, sColumn, sValue)
	GBL_FAILED_FUNCTION_NAME="Fn_DC_ContextDefinitionOperations"
	Dim objApplet, iColCnt, bFlag, objTable, iCnt 
	Set objApplet = JavaWindow("DesignContextWindow").JavaWindow("DesignContextApplet")
	Set objTable = objApplet.JavaTable("SelectedProductContexts")
	Fn_DC_ContextDefinitionOperations = False
	bFlag = False
	' checking existence of JavaTable
	If Fn_UI_ObjectExist("Fn_DC_ContextDefinitionOperations",objTable) = False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Function - [ Fn_DC_ContextDefinitionOperations ] Failed to find Java Table [ Selected Product Contexts  ]")
		Set objApplet = Nothing
		Set objTable = Nothing
		Exit function
	End If
	' checking existence of specified column
	If sColumn <> "" Then
			For iColCnt = 0 to cInt(objTable.GetROProperty("cols"))
				If trim(sColumn) = trim(objTable.Object.getColumnName(iColCnt)) Then
					bFlag = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Function - [ Fn_DC_ContextDefinitionOperations ] Successfully identified column [ " & sColumn & " ]")
					Exit for
				End If
			Next
			If NOT(bFlag) Then
				' clumn not found
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Function - [ Fn_DC_ContextDefinitionOperations ] Failed to find column [ " & sColumn & " ]")
				Set objApplet = Nothing
				Set objTable = Nothing
				Exit function
			End If
	End If

	Select Case sAction
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Select"
			bFlag = False
				For  iCnt = 0 to cInt(objTable.GetROProperty("rows")) -1
					If sObjectRow = objTable.GetCellData(iCnt, 0) Then
						objTable.SelectRow iCnt
						bFlag = True
						Exit for
					End If
				Next
				Fn_DC_ContextDefinitionOperations = bFlag 
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Remove"
				bFlag = False
				For  iCnt = 0 to cInt(objTable.GetROProperty("rows")) -1
					If sObjectRow = objTable.GetCellData(iCnt, 0) Then
						objTable.SelectRow iCnt
						bFlag = True
						Exit for
					End If
				Next
				If bFlag Then
					Call Fn_Button_Click("Fn_DC_ContextDefinitionOperations",objApplet,"ContextDefinitionDelete")
					Fn_DC_ContextDefinitionOperations = True
				End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Verify"
				bFlag = False
				For  iCnt = 0 to cInt(objTable.GetROProperty("rows")) -1
					If sObjectRow = objTable.GetCellData(iCnt, 0) Then
						If sValue  = objTable.GetCellData(iCnt, iColCnt ) Then
							bFlag = True
							Exit for
						End If
					End If
				Next
				Fn_DC_ContextDefinitionOperations = bFlag
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Function - [ Fn_DC_ContextDefinitionOperations ] Invalid case [ " & sAction & " ]")
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	End Select

	If Fn_DC_ContextDefinitionOperations Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS : Function - [ Fn_DC_ContextDefinitionOperations ] executed successfully with case [ " & sAction & " ]")
	End If
	Set objApplet = Nothing
	Set objTable = Nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_DC_ProductItemsOperations
'@@
'@@    Description			:	Function Used to perform operations on Product Items
'@@
'@@    Parameters			:	1. sAction		: Action to be performed
'@@    							2. sObjectRow	: Item node name from Object Column
'@@    							3. sColumn		: Column Name
'@@    							4. sValue		: Value to be verified. 
'@@
'@@    Return Value		   	: 	True Or False
'@@
'@@    Pre-requisite		:	Design Context perspective should be displayed							
'@@
'@@    Examples				:	Call Fn_DC_ProductItemsOperations("Select", "000023/A;1-top", "", "")
'@@    						:	Call Fn_DC_ProductItemsOperations("AddToSelectedProductContexts", "000024/A;1-t2", "", "")
'@@    						:	Call Fn_DC_ProductItemsOperations("Verify", "000024/A;1-t2", "Description", "Product")
'@@
'@@	   History					 	:	
'@@				Developer Name				Date			Rev. No.	Changes Done								Reviewer
'@@----------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe			15-Dec-2011			1.0			Created
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_DC_ProductItemsOperations(sAction, sObjectRow, sColumn, sValue)
	GBL_FAILED_FUNCTION_NAME="Fn_DC_ProductItemsOperations"
	Dim objApplet, iColCnt, bFlag, objTable, iCnt  
	Set objApplet = JavaWindow("DesignContextWindow").JavaWindow("DesignContextApplet")
	Set objTable = objApplet.JavaTable("TCCompositePropertyTable")
	Fn_DC_ProductItemsOperations = False
	bFlag = False
	' selecting tab
	objApplet.JavaTab("ProductItemsTabbedPane").Select "Product Items"
'	' checking existence of JavaTable
	If Fn_UI_ObjectExist("Fn_DC_ProductItemsOperations",objTable) = False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Function - [ Fn_DC_ProductItemsOperations ] Failed to find Java Table [ Product Contexts  ]")
		Set objApplet = Nothing
		Set objTable = Nothing
		Exit function
	End If

' checking existence of specified column
	If sColumn <> "" Then
			For iColCnt = 0 to cInt(objTable.GetROProperty("cols"))
				If trim(sColumn) = trim(objTable.Object.getColumnName(iColCnt)) Then
					bFlag = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Function - [ Fn_DC_ContextDefinitionOperations ] Successfully identified column [ " & sColumn & " ]")
					Exit for
				End If
			Next
			If NOT(bFlag) Then
				' clumn not found
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Function - [ Fn_DC_ContextDefinitionOperations ] Failed to find column [ " & sColumn & " ]")
				Set objApplet = Nothing
				Set objTable = Nothing
				Exit function
			End If
	End If

	Select Case sAction
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Select", "AddToSelectedProductContexts"
			bFlag = False
				For  iCnt = 0 to cInt(objTable.GetROProperty("rows")) -1
					If sObjectRow = objTable.GetCellData(iCnt, 0) Then
						objTable.SelectRow iCnt
						If sAction = "AddToSelectedProductContexts" Then
							Call Fn_Button_Click("Fn_DC_ProductItemsOperations", objApplet, "AddContextDefinition")
						End If
						bFlag = True
						Exit for
					End If
				Next
				Fn_DC_ProductItemsOperations = bFlag 
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Verify"
                  bFlag = False
				For  iCnt = 0 to cInt(objTable.GetROProperty("rows")) -1
					If sObjectRow = objTable.GetCellData(iCnt,  0) Then
						If sValue  = objTable.GetCellData(iCnt, iColCnt ) Then
							bFlag = True
							Exit for
						End If
					End If
				Next
				Fn_DC_ProductItemsOperations = bFlag
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Function - [ Fn_DC_ProductItemsOperations ] Invalid case [ " & sAction & " ]")
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	End Select

	If Fn_DC_ProductItemsOperations Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS : Function - [ Fn_DC_ProductItemsOperations ] executed successfully with case [ " & sAction & " ]")
	End If
	Set objApplet = Nothing
	Set objTable = Nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_DC_WorkPartsOperations
'@@
'@@    Description			:	Function Used to perform operations on WorkParts
'@@
'@@    Parameters			:	1. sAction		: Action to be performed
'@@    							2. sWorkItem	: WorkParts Item Name ( ~ separated list of workparts in case of Verify )
'@@    							3. sNewWorkItem	: New WorkParts Item Name 
'@@
'@@    Return Value		   	: 	True Or False
'@@
'@@    Pre-requisite		:	Design Context perspective should be displayed							
'@@
'@@    Examples				:	Call Fn_DC_WorkPartsOperations("Add", "000023", "")
'@@    						:	Call Fn_DC_WorkPartsOperations("Replace", "000157-wp", "000158")
'@@    						:	Call Fn_DC_WorkPartsOperations("Remove", "000157-wp", "")
'@@    						:	Call Fn_DC_WorkPartsOperations("Verify", "000157-wp~000158-wp1", "")
'@@
'@@	   History					 	:	
'@@				Developer Name				Date			Rev. No.	Changes Done								Reviewer
'@@----------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe			16-Dec-2011			1.0			Created
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_DC_WorkPartsOperations(sAction, sWorkItem, sNewWorkItem)
	GBL_FAILED_FUNCTION_NAME="Fn_DC_WorkPartsOperations"
	Dim objApplet, iCnt, arrWorkParts, bFlag 
	Set objApplet = JavaWindow("DesignContextWindow").JavaWindow("DesignContextApplet")
	Fn_DC_WorkPartsOperations = False
	bFlag = False

	Select Case sAction
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Add"
			Call Fn_Edit_Box("Fn_DC_WorkPartsOperations",objApplet, "WorkParts", sWorkItem)
			Call Fn_Button_Click("Fn_DC_WorkPartsOperations",objApplet, "WorkPartsAdd")
			Fn_DC_WorkPartsOperations = True
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Replace"
			Call Fn_List_Select("Fn_DC_WorkPartsOperations",objApplet, "WorkParts", sWorkItem)
			Call Fn_Edit_Box("Fn_DC_WorkPartsOperations",objApplet, "WorkParts", sNewWorkItem)
			Call Fn_Button_Click("Fn_DC_WorkPartsOperations",objApplet, "WorkPartsReplace")
			Fn_DC_WorkPartsOperations = True
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Remove"
			Call Fn_List_Select("Fn_DC_WorkPartsOperations",objApplet, "WorkParts", sWorkItem)
			Call Fn_Button_Click("Fn_DC_WorkPartsOperations",objApplet, "WorkPartsRemove")
			Fn_DC_WorkPartsOperations = True
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Verify"
			arrWorkParts = split(sWorkItem, "~")
			For iCnt = 0 to UBound(arrWorkParts)
					bFlag = Fn_UI_ListItemExist("Fn_DC_WorkPartsOperations",objApplet, "WorkParts", arrWorkParts(iCnt))
					Fn_DC_WorkPartsOperations = bFlag
					If bFlag Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS : Function - [ Fn_DC_WorkPartsOperations ] Successfully verified existence of [ " & arrWorkParts(iCnt) & " ] in [ WorkParts ] list")
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Function - [ Fn_DC_WorkPartsOperations ] Failed to verify existence of [ " & arrWorkParts(iCnt) & " ] in [ WorkParts ] list")
						Set objApplet = Nothing
						Exit function
					End If
			Next
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Function - [ Fn_DC_WorkPartsOperations ] Invalid case [ " & sAction & " ]")
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	End Select

	If Fn_DC_WorkPartsOperations Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS : Function - [ Fn_DC_WorkPartsOperations ] executed successfully with case [ " & sAction & " ]")
	End If
	Set objApplet = Nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_DC_EngChangeRevisionOperations
'@@
'@@    Description			:	Function Used to perform operations on WorkParts
'@@
'@@    Parameters			:	1. sAction			: Action to be performed
'@@    							2. sEngchangeRev	: EngChange Revision Item Name ( ~ separated list of EngChange Revision in case of Verify )
'@@    							3. sText			: Text Field
'@@    							4. sNewEngchangeRev	: New EngChange Revision Item Name
'@@    							5. sNewText			: New Text Field
'@@
'@@    Return Value		   	: 	True Or False
'@@
'@@    Pre-requisite		:	Design Context perspective should be displayed							
'@@
'@@    Examples				:	Call Fn_DC_EngChangeRevisionOperations("Add", "000023", "*","","")
'@@    						:	Call Fn_DC_EngChangeRevisionOperations("Replace", "000157-wp", "","000157-wp","*")
'@@    						:	Call Fn_DC_EngChangeRevisionOperations("Remove", "000157-wp", "", "", "")
'@@    						:	Call Fn_DC_EngChangeRevisionOperations("Verify", "000157-wp~000158-wp1", "","","")
'@@
'@@	   History					 	:	
'@@				Developer Name				Date			Rev. No.	Changes Done								Reviewer
'@@----------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe			16-Dec-2011			1.0			Created
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_DC_EngChangeRevisionOperations(sAction, sEngchangeRev, sText, sNewEngchangeRev, sNewText)
	GBL_FAILED_FUNCTION_NAME="Fn_DC_EngChangeRevisionOperations"
	Dim objApplet, arrEnvChangeRev, iCnt 
	Set objApplet = JavaWindow("DesignContextWindow").JavaWindow("DesignContextApplet")
	Fn_DC_EngChangeRevisionOperations = False
	bFlag  = False
	Select Case sAction
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Add"
			Call Fn_Edit_Box("Fn_DC_EngChangeRevisionOperations",objApplet, "EngChangeRevision", sEngchangeRev)
			If sText <> "" Then
				Call Fn_Edit_Box("Fn_DC_EngChangeRevisionOperations",objApplet, "iTextField", sText)
			End If
			Call Fn_Button_Click("Fn_DC_EngChangeRevisionOperations",objApplet, "EngChangeRevisionAdd")
			Fn_DC_EngChangeRevisionOperations = True
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Replace"
			Call Fn_List_Select("Fn_DC_EngChangeRevisionOperations",objApplet, "EngChangeRevision", sEngchangeRev)
			Call Fn_Edit_Box("Fn_DC_EngChangeRevisionOperations",objApplet, "EngChangeRevision", sNewEngchangeRev)
			If sNewText <> "" Then
				Call Fn_Edit_Box("Fn_DC_EngChangeRevisionOperations",objApplet, "iTextField", sNewText)
			End If
			Call Fn_Button_Click("Fn_DC_EngChangeRevisionOperations",objApplet, "EngChangeRevisionReplace")
			Fn_DC_EngChangeRevisionOperations = True
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Remove"
			Call Fn_List_Select("Fn_DC_EngChangeRevisionOperations",objApplet, "EngChangeRevision", sEngchangeRev)
			Call Fn_Button_Click("Fn_DC_EngChangeRevisionOperations",objApplet, "EngChangeRevisionRemove")
			Fn_DC_EngChangeRevisionOperations = True
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Verify"
			arrEnvChangeRev = split(sEngchangeRev, "~")
			For iCnt = 0 to UBound(arrEnvChangeRev)
					bFlag = Fn_UI_ListItemExist("Fn_DC_EngChangeRevisionOperations",objApplet, "EngChangeRevision", arrEnvChangeRev(iCnt))
					Fn_DC_EngChangeRevisionOperations = bFlag
					If bFlag Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS : Function - [ Fn_DC_EngChangeRevisionOperations ] Successfully verified existence of [ " & arrEnvChangeRev(iCnt) & " ] in [ WorkParts ] list")
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Function - [ Fn_DC_EngChangeRevisionOperations ] Failed to verify existence of [ " & arrEnvChangeRev(iCnt) & " ] in [ WorkParts ] list")
						Set objApplet = Nothing
						Exit function
					End If
			Next
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Function - [ Fn_DC_EngChangeRevisionOperations ] Invalid case [ " & sAction & " ]")
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	End Select

	If Fn_DC_EngChangeRevisionOperations Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS : Function - [ Fn_DC_EngChangeRevisionOperations ] executed successfully with case [ " & sAction & " ]")
	End If
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_DC_ProcessesOperations
'@@
'@@    Description			:	Function Used to perform operations on WorkParts
'@@
'@@    Parameters			:	1. sAction		: Action to be performed
'@@    							2. sProcess		: Process Item Name ( ~ separated list of Process in case of Verify )
'@@    							3. sNewProcess	: New Process Item Name 
'@@
'@@    Return Value		   	: 	True Or False
'@@
'@@    Pre-requisite		:	Design Context perspective should be displayed							
'@@
'@@    Examples				:	Call Fn_DC_ProcessesOperations("Add", "000023", "")
'@@    						:	Call Fn_DC_ProcessesOperations("Replace", "000157-wp", "000158")
'@@    						:	Call Fn_DC_ProcessesOperations("Remove", "000157-wp", "")
'@@    						:	Call Fn_DC_ProcessesOperations("Verify", "000157-wp~000158-wp1", "")
'@@
'@@	   History					 	:	
'@@				Developer Name				Date			Rev. No.	Changes Done								Reviewer
'@@----------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe			16-Dec-2011			1.0			Created
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_DC_ProcessesOperations(sAction, sProcess, sNewProcess)
	GBL_FAILED_FUNCTION_NAME="Fn_DC_ProcessesOperations"
	Dim objApplet, iCnt, arrProcesses, bFlag 
	Set objApplet = JavaWindow("DesignContextWindow").JavaWindow("DesignContextApplet")
	Fn_DC_ProcessesOperations = False
	bFlag = False

	Select Case sAction
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Add"
			Call Fn_Edit_Box("Fn_DC_ProcessesOperations",objApplet, "Processes", sProcess)
			Call Fn_Button_Click("Fn_DC_ProcessesOperations",objApplet, "ProcessesAdd")
			Fn_DC_ProcessesOperations = True
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Replace"
			Call Fn_Edit_Box("Fn_DC_ProcessesOperations",objApplet, "Processes", sNewProcess)
			Call Fn_List_Select("Fn_DC_ProcessesOperations",objApplet, "Processes", sProcess)
			Call Fn_Button_Click("Fn_DC_ProcessesOperations",objApplet, "ProcessesReplace")
			Fn_DC_ProcessesOperations = True
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Remove"
			Call Fn_List_Select("Fn_DC_ProcessesOperations",objApplet, "Processes", sProcess)
			Call Fn_Button_Click("Fn_DC_ProcessesOperations",objApplet, "ProcessesRemove")
			Fn_DC_ProcessesOperations = True
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Verify"
			arrProcesses = split(sProcess, "~")
			For iCnt = 0 to UBound(arrProcesses)
					bFlag = Fn_UI_ListItemExist("Fn_DC_ProcessesOperations",objApplet, "Processes", arrProcesses(iCnt))
					Fn_DC_ProcessesOperations = bFlag
					If bFlag Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS : Function - [ Fn_DC_ProcessesOperations ] Successfully verified existence of [ " & arrProcesses(iCnt) & " ] in [ Processes ] list")
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Function - [ Fn_DC_ProcessesOperations ] Failed to verify existence of [ " & arrProcesses(iCnt) & " ] in [ Processes ] list")
						Set objApplet = Nothing
						Exit function
					End If
			Next
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Function - [ Fn_DC_ProcessesOperations ] Invalid case [ " & sAction & " ]")
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	End Select

	If Fn_DC_ProcessesOperations Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS : Function - [ Fn_DC_ProcessesOperations ] executed successfully with case [ " & sAction & " ]")
	End If
	Set objApplet = Nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_DC_OccurrenceNotesSearchPanelOperations
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
'@@    								msgbox  Fn_DC_OccurrenceNotesSearchPanelOperations(sAction, bClear, bClearOccurrenceNotes, sOccurrenceNotes, sOperators, sValues, bClickOnSearchButton)
'@@
'@@	   History					 	:	
'@@				Developer Name				Date			Rev. No.	Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe			1-Dec-2011			1.0			Created
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public function Fn_DC_OccurrenceNotesSearchPanelOperations(sAction, bClear, sOccurrenceNotes, sOperators, sValues, bClickOnSearchButton)
	GBL_FAILED_FUNCTION_NAME="Fn_DC_OccurrenceNotesSearchPanelOperations"
	Dim objApplet
	Dim arrOccNotes, arrOperators, arrValues, iCnt, iRowCnt

	Fn_DC_OccurrenceNotesSearchPanelOperations = False
	Set objApplet = JavaWindow("DesignContextWindow").JavaWindow("DesignContextApplet")
	
	'setting index to 0 to open ItemID serach criteria window. 
	objApplet.JavaStaticText("SavedQuery").SetTOProperty "label", "Occurrence Notes"
	objApplet.JavaStaticText("SavedQuery").Click 1,1,"LEFT"

	' opening serach panel if it is not displayed.
	If Fn_UI_ObjectExist("Fn_DC_OccurrenceNotesSearchPanelOperations", objApplet.JavaTable("OccurrenceNotesTable")) = False Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_DC_OccurrenceNotesSearchPanelOperations ] Failed to click on [ Show/Hide Structure Manager Search Panel ].") 
				Fn_DC_OccurrenceNotesSearchPanelOperations = False
				Set objApplet = Nothing
				Exit function 
	End If
	

	Select Case sAction
		Case "Search"
				If bClear = "" then bClear = "False"
				If cBool(bClear) then
					Call Fn_Button_Click("Fn_DC_OccurrenceNotesSearchPanelOperations", objApplet,"Clear")
				End If
	
				'selecting occ note details
				arrOccNotes = split(sOccurrenceNotes, "~")
				arrOperators = split(sOperators, "~")
				arrValues = split(sValues, "~")
				iRowCnt = -1
				For iCnt = 0 to UBound(arrOccNotes)
                              ' clicking on + button
					Call Fn_Button_Click("Fn_DC_OccurrenceNotesSearchPanelOperations", objApplet, "ConfigFilter_Add")
					iRowCnt = iRowCnt + 1
					If trim(arrOccNotes(iCnt)) <> "" Then
						objApplet.JavaTable("OccurrenceNotesTable").ClickCell iRowCnt,"Occurrence Notes","LEFT"
						Call Fn_List_Select(	"Fn_DC_OccurrenceNotesSearchPanelOperations", objApplet, "ConfigFilterList", trim(arrOccNotes(iCnt)))
						objApplet.JavaTable("OccurrenceNotesTable").ClickCell iRowCnt,"Occurrence Notes","LEFT"
					End If
					' setting operator
					If trim(arrOperators(iCnt)) <> "" Then
						objApplet.JavaTable("OccurrenceNotesTable").ClickCell iRowCnt,"Operator","LEFT"
'						wait 2
'						objApplet.JavaTable("OccurrenceNotesTable").ClickCell iRowCnt,"Operator","LEFT"
'						If objApplet.JavaList("OccNotesList").exist(5) = False then
'							objApplet.JavaTable("OccurrenceNotesTable").ClickCell iRowCnt,"Operator","LEFT"
'						End If
'						Call Fn_List_Select("Fn_DC_OccurrenceNotesSearchPanelOperations", objApplet, "OccNotesList", trim(arrOperators(iCnt)))
						objApplet.JavaTable("OccurrenceNotesTable").SetCellData iRowCnt,"Operator", trim(arrOperators(iCnt))
					End If
					' setting values
					If trim(arrValues(iCnt)) <> "" Then
'						objApplet.JavaTable("OccurrenceNotesTable").SetCellData iRowCnt,"Value", trim(arrValues(iCnt))
						objApplet.JavaTable("OccurrenceNotesTable").ClickCell iRowCnt,"Value"
						call Fn_SISW_UI_JavaEdit_Operations("Fn_DC_OccurrenceNotesSearchPanelOperations", "Set", objApplet, "OccurrenceNotesValue", trim(arrValues(iCnt)))

						
					End If
				Next

				' clicking on search button
				If bClickOnSearchButton = "" then bClickOnSearchButton = True
				If cBool(bClickOnSearchButton) Then
					Call Fn_Button_Click("Fn_DC_OccurrenceNotesSearchPanelOperations", objApplet,"Update")
				End If
				Fn_DC_OccurrenceNotesSearchPanelOperations = True
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_DC_OccurrenceNotesSearchPanelOperations ] Invalid case [ " & sAction& " ].") 
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	End Select

	If Fn_DC_OccurrenceNotesSearchPanelOperations = True Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Function [ Fn_DC_OccurrenceNotesSearchPanelOperations ] Executed successfully with case [ " & sAction& " ].") 
	End If
	Set objApplet = Nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_DC_SearchResultOperations
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
'@@    Examples					:	msgbox  Fn_DC_SearchResultOperations("VerifyInBackgroundPartAppearances", "Search Result...1", "", "Item Name", "Comp2")
'@@			    					msgbox  Fn_DC_SearchResultOperations("VerifyInBackgroundPartInstallationAssemblies", "", "", "", "Comp2")
'@@    Examples					:	msgbox  Fn_DC_SearchResultOperations("SelectInBackgroundPartAppearances", "Search Result...1", "Comp2", "", "")
'@@    								msgbox  Fn_DC_SearchResultOperations("SelectInBackgroundPartInstallationAssemblies", "", "Comp2", "", "")
'@@    Examples					:	msgbox  Fn_DC_SearchResultOperations("SelectInTargetPartAppearances", "Search Result...1", "Comp2", "", "")
'@@    								msgbox  Fn_DC_SearchResultOperations("VerifyRowInBackgroundPartAppearances", "","comp-000060","Parent~Item Description" ,"000059/A;1-top (View)~ ")
'@@    								msgbox  Fn_DC_SearchResultOperations("VerifyRowInBackgroundPartInstallationAssemblies", "","comp-000060","Parent~Item Description" ,"000059/A;1-top (View)~ ")
'@@									msgbox Fn_DC_SearchResultOperations("GetColumIndexInBackgroundPartAppearances", "", "",  "APN UID", "")
'@@									msgbox Fn_DC_SearchResultOperations("AddColumInBackgroundPartAppearances", "", "",  "APN UID", "")
'@@									msgbox Fn_DC_SearchResultOperations("RemoveColumInBackgroundPartAppearances", "", "",  "APN UID", "")
'@@									bReturn=Fn_DC_SearchResultOperations("MultiSelectInBackgroundPartAppearances", "", "DVD Player~CD Player", "", "")
'@@									bReturn=Fn_DC_SearchResultOperations("CloseTab", "Search Result...3", "", "", "")
'@@
'@@	   History					 	:	
'@@				Developer Name				Date			Rev. No.	Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe			1-Dec-2011			1.0			Created
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe		   13-Jan-2012			1.0			Added cases SelectInBackgroundPartAppearances, SelectInBackgroundPartInstallationAssemblies
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Sandeep Navghane		  05-Jul-2013			1.1			Added cases MultiSelectInBackgroundPartAppearances
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_DC_SearchResultOperations(sAction, sTab, sRow, sColumn, sValue)
	GBL_FAILED_FUNCTION_NAME="Fn_DC_SearchResultOperations"
	Dim objApplet, objTable, iRowCnt, iCnt , aNode, iInstanceCnt, arrValSet, aValue, aColumns, aRow, objChangeColumnDialog 
	Dim iCounter,aNodeInstance,bFlag
	Set objApplet = JavaWindow("DesignContextWindow").JavaWindow("DesignContextApplet")
	Fn_DC_SearchResultOperations = False
	'select search panel
	If sTab <> "" Then
		objApplet.JavaStaticText("SearchPanel").SetTOProperty "label", sTab
		objApplet.JavaStaticText("SearchPanel").Click 1,1,"LEFT"
	End If
	Select Case sAction
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "AddColumInBackgroundPartAppearances", "AddColumInBackgroundPartInstallationAssemblies"
			If sAction = "AddColumInBackgroundPartInstallationAssemblies" Then
				Set objTable = objApplet.JavaTable("BackgroundPartInstallationAssemblies")
				iInstanceCnt =  Fn_DC_SearchResultOperations("GetColumIndexInBackgroundPartInstallationAssemblies", sTab, "", sColumn, "")
			Else
				Set objTable = objApplet.JavaTable("BackgroundPartAppearances")
				iInstanceCnt =  Fn_DC_SearchResultOperations("GetColumIndexInBackgroundPartAppearances", sTab, "", sColumn, "")
			End If
			If iInstanceCnt = -1 Then
				' add column
				objTable.SelectColumnHeader 0 ,"RIGHT"
				wait 1
				Call Fn_UI_JavaMenu_Select("",JavaWindow("DesignContextWindow"),"Insert column\(s\) ...")
				If JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("ChangeColumns").Exist(2) Then
					Set objChangeColumnDialog = JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("ChangeColumns")
				Else
					Set objChangeColumnDialog = JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("ChangeColumns")
				End If
                
				If NOT objChangeColumnDialog.Exist  Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: RMB action Insert Column(s).... Executed successfully in the Application.")			
					Exit function
				End If

				objChangeColumnDialog.JavaList("AvailableCol").ExtendSelect sColumn
                Call Fn_Button_Click("Fn_DC_SearchResultOperations", objChangeColumnDialog, "Add")
				' Hit  Apply Button after selection
                Call Fn_Button_Click("Fn_DC_SearchResultOperations", objChangeColumnDialog, "Apply")
                Call Fn_Button_Click("Fn_DC_SearchResultOperations", objChangeColumnDialog, "Cancel")
                Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Pass:Successfully Added  Column  ["& sColumn &"] in BOMTable")									
				Fn_DC_SearchResultOperations = TRUE
			Else
				Fn_DC_SearchResultOperations = True
			End If
			
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "RemoveColumInBackgroundPartAppearances", "RemoveColumInBackgroundPartInstallationAssemblies"
			If sAction = "RemoveColumInBackgroundPartInstallationAssemblies" Then
				Set objTable = objApplet.JavaTable("BackgroundPartInstallationAssemblies")
				iInstanceCnt =  Fn_DC_SearchResultOperations("GetColumIndexInBackgroundPartInstallationAssemblies", sTab, "", sColumn, "")
			Else
				Set objTable = objApplet.JavaTable("BackgroundPartAppearances")
				iInstanceCnt =  Fn_DC_SearchResultOperations("GetColumIndexInBackgroundPartAppearances", sTab, "", sColumn, "")
			End If
			If iInstanceCnt <> -1 Then
				' remove column
				objTable.SelectColumnHeader iInstanceCnt ,"RIGHT"
				wait 1
				Call Fn_UI_JavaMenu_Select("",JavaWindow("DesignContextWindow"),"Remove this column")
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Pass: Successfully removed Column  ["& sColumn &"] from BOMTable.")
				'Fn_DC_SearchResultOperations = Fn_Button_Click("Fn_DC_SearchResultOperations", JavaWindow("DefaultWindow").JavaWindow("RemoveColumn"), "Yes")
				If JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Remove Column").Exist(2) Then
					Fn_DC_SearchResultOperations = Fn_Button_Click("Fn_DC_SearchResultOperations", JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Remove Column"), "Yes")
				ElseIf JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Remove Column").Exist(2) Then
					Fn_DC_SearchResultOperations = Fn_Button_Click("Fn_DC_SearchResultOperations", JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Remove Column"), "Yes")				
				End If
				'Fn_DC_SearchResultOperations = Fn_Button_Click("Fn_DC_SearchResultOperations", JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Remove Column"), "Yes")
			else
				Fn_DC_SearchResultOperations = True
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "GetColumIndexInBackgroundPartAppearances", "GetColumIndexInBackgroundPartInstallationAssemblies"
			If sAction = "GetColumIndexInBackgroundPartInstallationAssemblies" Then
				Set objTable = objApplet.JavaTable("BackgroundPartInstallationAssemblies")
			Else
				Set objTable = objApplet.JavaTable("BackgroundPartAppearances")
			End If
			Fn_DC_SearchResultOperations = -1 
			iInstanceCnt = cInt(objTable.GetROProperty("cols"))
			For iCnt = 0 to iInstanceCnt - 1
				If objTable.Object.getColumnName(iCnt) = sColumn Then
					Fn_DC_SearchResultOperations = iCnt
					Exit for
				End If
			Next
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "VerifyInBackgroundPartAppearances", "VerifyInBackgroundPartInstallationAssemblies"
			If sAction = "VerifyInBackgroundPartInstallationAssemblies" Then
				Set objTable = objApplet.JavaTable("BackgroundPartInstallationAssemblies")
			Else
				Set objTable = objApplet.JavaTable("BackgroundPartAppearances")
			End If

			If sColumn = "" Then sColumn = "Item Name"
			aNode = split(sValue, "@")
			iInstanceCnt = 1
			If uBound( aNode) = 1  Then
				iInstanceCnt = cInt(aNode(1))
			End If
			wait 5
			For iCnt = 0 to cInt(objTable.GetROProperty("rows")) - 1
					If cStr(objTable.GetCellData(iCnt, sColumn)) = cstr(aNode(0)) then
						If iInstanceCnt = 1 Then
								Fn_DC_SearchResultOperations = True
								Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Value [ " & sValue & " ] found in column [ " & sColumn & " ].")
								Exit for
						Else
							iInstanceCnt = iInstanceCnt - 1
						End If
					End If
			Next
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "SelectInBackgroundPartAppearances", "SelectInBackgroundPartInstallationAssemblies"
			If sAction = "SelectInBackgroundPartInstallationAssemblies" Then
				Set objTable = objApplet.JavaTable("BackgroundPartInstallationAssemblies")
			Else
				Set objTable = objApplet.JavaTable("BackgroundPartAppearances")
			End If

			sColumn = "Item Name"
			aNode = split(sRow, "@")
			iInstanceCnt = 1
			If uBound( aNode) = 1  Then
				iInstanceCnt = cInt(aNode(1))
			End If
			wait 5
			For iCnt = 0 to cInt(objTable.GetROProperty("rows")) - 1
					If cStr(objTable.GetCellData(iCnt, sColumn)) = cstr(aNode(0)) then
						If iInstanceCnt = 1 Then
							objTable.SelectRow iCnt
							Fn_DC_SearchResultOperations = True
							Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Value [ " & sRow & " ] found in column [ " & sColumn & " ].")
							Exit for
						Else
							iInstanceCnt = iInstanceCnt - 1
						End If
					End If
			Next
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "SelectInTargetPartAppearances", "SelectInTargetPartInstallationAssemblies"
			If sAction = "SelectInTargetPartInstallationAssemblies" Then
				'Set objTable = objApplet.JavaTable("BackgroundPartInstallationAssemblies")
				' do nothing
			Else
				Set objTable = objApplet.JavaTable("TargetPartAppearances")
			End If

			sColumn = "Item Name"
			aNode = split(sRow, "@")
			iInstanceCnt = 1
			If uBound( aNode) = 1  Then
				iInstanceCnt = cInt(aNode(1))
			End If
			wait 5
			For iCnt = 0 to cInt(objTable.GetROProperty("rows")) - 1
					If cStr(objTable.GetCellData(iCnt, sColumn)) = cstr(aNode(0)) then
						If iInstanceCnt = 1 Then
							objTable.SelectRow iCnt
							Fn_DC_SearchResultOperations = True
							Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Value [ " & sRow & " ] found in column [ " & sColumn & " ].")
							Exit for
						Else
							iInstanceCnt = iInstanceCnt - 1
						End If
					End If
			Next

		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "VerifyRowInBackgroundPartAppearances", "VerifyRowInBackgroundPartInstallationAssemblies"
			If sAction = "VerifyRowInBackgroundPartInstallationAssemblies" Then
				Set objTable = objApplet.JavaTable("BackgroundPartInstallationAssemblies")
			Else
				Set objTable = objApplet.JavaTable("BackgroundPartAppearances")
			End If

			iRowCount = cInt(objApplet.JavaTable("BackgroundPartAppearances").GetROProperty("rows"))
			iInstCnt = 1
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

			For iCnt = 0 to iRowCount -1
				If Cstr(trim(objTable.GetCellData(iCnt, "Item Name"))) = cStr(arrValSet(0)) then
					If IsNumeric(arrValSet(1)) Then
						arrValSet(1) = cStr(cInt(arrValSet(1)))
					End IF
					If Cstr(trim(objTable.GetCellData(iCnt, "Item Id"))) = cStr(arrValSet(1)) then
							iF iInstCnt = iInstance Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Function [ Fn_DC_SearchResultOperations ] Successfully verified [ " & cStr(arrValSet(0)) & "-" & arrValSet(1) & " ] is present.") 
								Fn_DC_SearchResultOperations = True
								If sColumn <> "" Then
									Fn_DC_SearchResultOperations = False
									aColumns = split(sColumn,"~")
									aValue = split(sValue,"~")
									For iArrCnt = 0 to uBound(aColumns)
										If trim(Cstr(trim(objTable.GetCellData(iCnt, aColumns(iArrCnt))))) <> trim(cStr(aValue(iArrCnt))) then
											Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_DC_SearchResultOperations ] Successfully verified [ " & aValue(iArrCnt) &" ] is present in column [ " & aColumns(iArrCnt) & " ].") 
											Fn_DC_SearchResultOperations = False
											Exit for
										End If
										Fn_DC_SearchResultOperations = True
									Next
								End If
								IF Fn_DC_SearchResultOperations = False Then
									Exit for
								End If
							End If
							iInstCnt = iInstCnt + 1
					End If
				end if
			Next
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
        Case "MultiSelectInBackgroundPartAppearances"
			Set objTable = objApplet.JavaTable("BackgroundPartAppearances")
			sColumn = "Item Name"
			aNode = split(sRow, "~")
		
			For iCounter=0 to ubound(aNode)
				bFlag = False
				iInstanceCnt = 1
				aNodeInstance=Split(aNode(iCounter),"@")
				If uBound( aNodeInstance) = 1  Then
					iInstanceCnt = cInt(aNodeInstance(1))
				End If
				wait 5
				For iCnt = 0 to cInt(objTable.GetROProperty("rows")) - 1
						If cStr(objTable.GetCellData(iCnt, sColumn)) = cstr(aNodeInstance(0)) then
							If iInstanceCnt = 1 Then
								If iCounter=0 Then
									objTable.SelectRow iCnt
								Else
									objTable.ExtendRow iCnt
								End If
								bFlag = True
								Exit for
							Else
								iInstanceCnt = iInstanceCnt - 1
							End If
						End If
				Next
				If bFlag = False Then
					Exit for
				End If
			Next
			If bFlag = True Then
				Fn_DC_SearchResultOperations=True
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "CloseTab","CloseAllTab"
					objApplet.JavaStaticText("SearchPanel").Click 1,1,"RIGHT"
					wait 1
					If sAction="CloseTab" Then
						objApplet.JavaMenu("label:=Close","index:=0").Select
					Else
						objApplet.JavaMenu("label:=Close All","index:=0").Select
					End If
					wait 1
					If Err.Number < 0 Then
						Fn_DC_SearchResultOperations=False
					Else
						Fn_DC_SearchResultOperations=True
					End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_DC_SearchResultOperations ] Invalid case [ " & sAction & " ].") 
				Exit function
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	End Select
	Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Function [ Fn_DC_SearchResultOperations ] executed successfully with case [ " & sAction & " ].") 
	Set objTable = Nothing
	Set objApplet = Nothing
End Function

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_DC_CloseDesignContext
'@@
'@@    Description				 :	Function Used to close Desaign Context
'@@
'@@    Parameters			   :	1.sAction : Action to be performed
'@@   	 							2.sBtnName : to clear search criteria ( True / False / "" )
'@@   	 							3.sValue : For future use
'@@
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	Desaign Context should be displayed						
'@@
'@@    Examples					:	msgbox Fn_DC_CloseDesignContext("Menu", "Yes", "")
'@@
'@@	   History					 	:	
'@@				Developer Name				Date			Rev. No.	Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe			1-Dec-2011			1.0			Created
'@@				Sachin Joshi			14-Sept-2012			1.0			Modified
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_DC_CloseDesignContext(sAction, sBtnName, sValue)
	GBL_FAILED_FUNCTION_NAME="Fn_DC_CloseDesignContext"
	Dim objExitDC,objSMClose,objClose
	Set objExitDC = Fn_SISW_DC_GetObject("DesignContextExit")
	Set objSMClose = Fn_SISW_DC_GetObject("There are unsaved Structure")
	Set objClose = Fn_SISW_GetObject("ConfirmationDialog")
	Fn_DC_CloseDesignContext = False
	Select Case sAction
		Case "Menu", ""
				bReturn =  Fn_MenuOperation("Select","File:Close")
				If bReturn = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Failed to close Design Context.") 
					Exit function
				End If
				if sBtnName = ""  then sBtnName = "Yes"
				If objExitDC.Exist(10) Then
					objExitDC.WinButton(sBtnName).Click 1,1,micLeftBtn
					Fn_DC_CloseDesignContext = True
				End If
				If objSMClose.Exist(10) Then
					objSMClose.WinButton(sBtnName).Click 1,1,micLeftBtn
					Fn_DC_CloseDesignContext = True
				End If
				objClose.SetTOProperty "title","There are unsaved Structure Manager changes."
				wait 1
				If objClose.Exist(10) Then
					Call Fn_Button_Click("Fn_DC_CloseDesignContext",objClose, sBtnName)
					Fn_DC_CloseDesignContext = True
				Else
					'Handle to avoid unnecessary failure when no dialog appears
					JavaWindow("DefaultWindow").JavaStaticText("MyTeamcenter").SetTOProperty "label","DesignContext"
					If Not JavaWindow("DefaultWindow").JavaStaticText("MyTeamcenter").Exist(5) Then
						Fn_DC_CloseDesignContext = True
					End If
				End If
				
	End Select
	If Fn_DC_CloseDesignContext Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Function [ Fn_DC_CloseDesignContext ] executed successfully with case [ " & sAction & " ].") 
	End If
	Set objExitDC = Nothing
	Set objSMClose = Nothing
End Function

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name			:	Fn_DC_CloseDesignContext
'@@
'@@    Description				:	Function Used to perform search operation on Item Attribute dialog
'@@
'@@    Parameters			    :	1. dicItemIDSearch: dictionary object
'@@
'@@    Return Value		   	    : 	True Or False
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
'@@    								Call Fn_DC_ItemIDSearchPanelOperations("Search", dicItemIDSearch)
'@@
'@@	   History					:
'@@				Developer Name				Date			Rev. No.	Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe			24-Jan-2012			1.0			Created
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_DC_ItemIDSearchPanelOperations(sAction, dicItemIDSearch)
	GBL_FAILED_FUNCTION_NAME="Fn_DC_ItemIDSearchPanelOperations"
	Dim objApplet, bReturn, sTemplateType, iCnt, arrFieldValue, intNoOfObjects, iCount, arrSearchCriteria
	Set objApplet = JavaWindow("DesignContextWindow").JavaWindow("DesignContextApplet")

    'setting index to 0 to open ItemID serach criteria window. 
	objApplet.JavaStaticText("SavedQuery").SetTOProperty "label", "Saved Query"
	objApplet.JavaStaticText("SavedQuery").Click 1,1,"LEFT"

	' opening serach panel if it is not displayed.
	If Fn_UI_ObjectExist("Fn_DC_OccurrenceNotesSearchPanelOperations", objApplet.JavaButton("SavedQueryChange")) = False Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_DC_ItemIDSearchPanelOperations ] Failed to select [ Saved Query ].") 
				Fn_DC_ItemIDSearchPanelOperations = False
				Set objApplet = Nothing
				Exit function 
	End If

	Select Case sAction
		Case "Search"
				If dicItemIDSearch("bChangeSearch") <> "" Then
					If cBool(dicItemIDSearch("bChangeSearch")) Then
						' clicking on ... button of ItemID search criteria
						Call Fn_Button_Click("Fn_RDV_ItemIDSearchPanelOperations", objApplet, "SavedQueryChange")
		
						IF Fn_RDV_ChangeSearch(dicItemIDSearch) = False then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_DC_ItemIDSearchPanelOperations ] Failed to open [ change Search ] dialog.") 
							Exit function
						End IF
					End If
				End If

				If dicItemIDSearch("SearchCriteria") <> "" Then
					If dicItemIDSearch("bClear") = "" then dicItemIDSearch("bClear") = "False"
					If cBool(dicItemIDSearch("bClear")) then
						' clearing form
						Call Fn_Button_Click("Fn_DC_OccurrenceNotesSearchPanelOperations", objApplet, "SavedQueryClear" )
					End IF
					
					
					Call Fn_ReadyStatusSync(2)
					arrSearchCriteria = split(dicItemIDSearch("SearchCriteria"),"~")
					For iCnt = 0 to UBound(arrSearchCriteria)
						arrFieldValue = split(arrSearchCriteria(iCnt),"=")
						objApplet.JavaStaticText("Field").SetTOProperty  "label", trim(arrFieldValue(0)) & ":"
						wait 1
						If objApplet.JavaButton("MultipleDropdownButton").Exist(2) Then
							objApplet.JavaButton("MultipleDropdownButton").Click micLeftBtn
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
						ElseIf objApplet.JavaEdit("Field").exist(3) Then
							Call Fn_Edit_Box("Fn_RDV_ItemAttributes", objApplet,"Field", trim(arrFieldValue(1) ))
						ElseIf objApplet.JavaCheckBox("DateField").exist(3) Then
							objApplet.JavaCheckBox("DateField").Object.setDate trim(arrFieldValue(1) )
						Else
							Exit function
						End If
					Next
				End If

			' clicking on search button
			If dicItemIDSearch("bClickOnSearchButton") = "" then 
					dicItemIDSearch("bClickOnSearchButton") = True
			End If

			If cBool(dicItemIDSearch("bClickOnSearchButton")) Then
				Call Fn_Button_Click("Fn_DC_ItemIDSearchPanelOperations", objApplet, "Update")
			End If
			Fn_DC_ItemIDSearchPanelOperations = True
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_DC_ItemIDSearchPanelOperations ] Invalid case [ " & sAction& " ].") 
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		End Select

	If Fn_DC_ItemIDSearchPanelOperations = True Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Function [ Fn_DC_ItemIDSearchPanelOperations ] Executed successfully with case [ " & sAction& " ].") 
	End If

	Set objApplet = Nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name			:	Fn_DC_ZoneOperations
'@@
'@@    Description				:	Function Used to perform operations related to Zone.
'@@
'@@    Parameters			    :	1. dicZone: dictionary object
'@@
'@@    Return Value		   	    : 	True Or False
'@@
'@@    Pre-requisite			:	Design Context perspective should be displayed
'@@
'@@    								dicZone("Description") = "desc"
'@@    Examples					:	
'@@    								Dim dicZone
'@@    								Set dicZone = CreateObject("Scripting.Dictionary")
'@@    								 With dicZone
'@@    									.Add "Name","name" 
'@@    									.Add "Type","HRN_Core"
'@@    									.Add "Description","desc" 
'@@    									.Add "bOpenOnCreate", False 
'@@    									.Add "Field","Description~Name" 
'@@    									.Add "Value","desc~comp1"
'@@    								End With
'@@    								Call Fn_DC_ZoneOperations("Create", dicZone)
'@@    								Call Fn_DC_ZoneOperations("CreateAndEditDetails", dicZone)
'@@    								Call Fn_DC_ZoneOperations("CreateAndVerifyDetails", dicZone)
'@@
'@@	   History					:
'@@				Developer Name				Date			Rev. No.		Reviewer		Changes Done							
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe			13-Feb-2012			1.0								Created
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Pooja Bondarde			21-Jun-2012			1.0				Koustubh		modified object hierarchy for Tc10.0 Build 2012061300
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_DC_ZoneOperations(sAction, dicZone)
	GBL_FAILED_FUNCTION_NAME="Fn_DC_ZoneOperations"
	Dim objZone, objZoneDetails, iCnt, aFields, aFieldValues, bFlag
	Dim objSelectType, intNoOfObjects
	Fn_DC_ZoneOperations = False
	bFlag = False
	Select Case sAction
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Create", "CreateAndVerifyDetails", "CreateAndEditDetails"
					Set objZone = JavaWindow("DesignContextWindow").JavaWindow("DesignContextApplet").JavaDialog("NewZone")
					If Fn_UI_ObjectExist("Fn_DC_ZoneOperations", objZone) = False Then
						' perform menu operation
						Call Fn_MenuOperation("Select", "File:New:Zone...")
						If Fn_UI_ObjectExist("Fn_DC_ZoneOperations", objZone) = False Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_DC_ZoneOperations ] Failed to find New Zone dialog.") 
								Set objZone = Nothing
								Exit function
						End If
					End If
					' setting name
					Call Fn_Edit_Box("Fn_DC_ZoneOperations", objZone,"Name",dicZone("Name"))
	
					' setting description
					If dicZone("Description") <> "" Then
						Call Fn_Edit_Box("Fn_DC_ZoneOperations", objZone,"Description",dicZone("Description"))
					End If
	
					' setting type
					If dicZone("Type") <>"" Then
							Call Fn_CheckBox_Set("Fn_DC_ZoneOperations",objZone, "More",  "ON" )
							wait(3)
							Call Fn_ReadyStatusSync(2)
							'Set Form Type
							Set objSelectType=description.Create()
							objSelectType("Class Name").value = "JavaStaticText"
							objSelectType("label").value = dicZone("Type")
							Set  intNoOfObjects = objZone.ChildObjects(objSelectType)
					        If Environment.Value("ProductName")=sUFTProductName Then
					        	For  iCounter = 0 to intNoOfObjects.count-1
	                             	If  intNoOfObjects(iCounter).getROProperty("label") = dicZone("Type") Then
			                        	intNoOfObjects(iCounter).Click 1,1
					                    bFlag=True
			                    	    Exit For
		                            End If
                                Next
                            Else
                               If intNoOfObjects(0).exist(5) Then
								   intNoOfObjects(0).Click 1,1
							   Else 
								   Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_DC_ZoneOperations ] Failed to find Zone type [ " & dicZone("Type") &" ].") 
								   Set objZone = Nothing
								   Exit function
							   End If
					        End If
				   End If
					' setting open on create
					If dicZone("bOpenOnCreate") <> "" Then
						If cBool(dicZone("bOpenOnCreate")) Then
							Call Fn_CheckBox_Set("Fn_DC_ZoneOperations", objZone,"OpenOnCreate","ON")
							bFlag = True
						Else
							Call Fn_CheckBox_Set("Fn_DC_ZoneOperations", objZone,"OpenOnCreate","OFF")
						End If
					End If
					' clickin on OK button
					Call Fn_Button_Click("Fn_DC_ZoneOperations", objZone,"OK")
					Fn_DC_ZoneOperations = True
					Call Fn_ReadyStatusSync(2)
					If bFlag = true Then
							Fn_DC_ZoneOperations = False
							' verifying for Zone details window..
							Set objZoneDetails = JavaWindow("DesignContextWindow").JavaWindow("DesignContextApplet").JavaDialog("ZoneDetails")
							If Fn_UI_ObjectExist("Fn_DC_ZoneOperations", objZoneDetails) = False Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_DC_ZoneOperations ] Failed to find Zone Details window.") 
									Set objZone = Nothing
									Set objZoneDetails = Nothing
									Exit function
							End If
							 
							 If dicZone("Field") <> "" Then
								 aFields = split(dicZone("Field"),"~")
								 aFieldValues = split(dicZone("Value"),"~")
							 End If
							Select Case sAction
							' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
								Case  "CreateAndVerifyDetails"
										For iCnt = 0 to UBound(aFields)
											Fn_DC_ZoneOperations = true
											objZoneDetails.JavaStaticText("FieldLabel").SetTOProperty "label" , aFields(iCnt) &":"
											If  Fn_UI_ObjectExist("Fn_DC_ZoneOperations", objZoneDetails.JavaEdit("FieldEditbox")) Then
												If objZoneDetails.JavaEdit("FieldEditbox").GetROProperty("value") <> aFieldValues(iCnt) Then
													Fn_DC_ZoneOperations = False
													Exit for
												End IF
											ElseIf  Fn_UI_ObjectExist("Fn_DC_ZoneOperations", objZoneDetails.JavaList("FieldListBox")) Then
												If NOT(Fn_UI_ListItemExist("Fn_DC_ZoneOperations", objZoneDetails, "FieldListBox", aFieldValues(iCnt))) Then
													Fn_DC_ZoneOperations = False
													Exit for
												End IF
											Else
												Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_DC_ZoneOperations ] Can not find field for [ " & aFields(iCnt) & " ].") 
												Set objZone = Nothing
												Set objZoneDetails = Nothing
												Exit function
											End If
										Next
							' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
							Case  "CreateAndEditDetails"
										Call Fn_Button_Click("Fn_DC_ZoneOperations", objZoneDetails,"CheckOutAndEdit")

										If  Fn_UI_ObjectExist("Fn_DC_ZoneOperations",JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Check-Out")) Then
											Call Fn_Button_Click("Fn_DC_ZoneOperations",JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Check-Out"),"Yes")
											Call Fn_ReadyStatusSync(2)
										ElseIf Fn_UI_ObjectExist("Fn_DC_ZoneOperations",JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Check-Out")) Then
											Call Fn_Button_Click("Fn_DC_ZoneOperations",JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Check-Out"),"Yes")
											Call Fn_ReadyStatusSync(2)
										End If

										For iCnt = 0 to UBound(aFields)
											objZoneDetails.JavaStaticText("FieldLabel").SetTOProperty "label" , aFields(iCnt) &":"
											If  Fn_UI_ObjectExist("Fn_DC_ZoneOperations", objZoneDetails.JavaEdit("FieldEditbox")) Then
												Call Fn_Edit_Box("Fn_DC_ZoneOperations", objZoneDetails,"FieldEditbox",aFieldValues(iCnt))
											ElseIf  Fn_UI_ObjectExist("Fn_DC_ZoneOperations", objZoneDetails.JavaList("FieldListBox")) Then
											Else
												Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_DC_ZoneOperations ] Can not find edit field for [ " & aFields(iCnt) & " ].") 
												Set objZone = Nothing
												Set objZoneDetails = Nothing
												Exit function
											End If
										Next

										Call Fn_Button_Click("Fn_DC_ZoneOperations", objZoneDetails,"SaveAndCheckIn")

										If  Fn_UI_ObjectExist("Fn_DC_ZoneOperations",JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Check-In")) Then
											Call Fn_Button_Click("Fn_DC_ZoneOperations",JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Check-In"),"Yes")
											Call Fn_ReadyStatusSync(2)
									    ElseIf Fn_UI_ObjectExist("Fn_DC_ZoneOperations",JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Check-In")) Then 
									    	Call Fn_Button_Click("Fn_DC_ZoneOperations",JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Check-In"),"Yes")
											Call Fn_ReadyStatusSync(2)
										End If

										'Call Fn_Button_Click("Fn_DC_ZoneOperations", objZoneDetails,"Close")
										Fn_DC_ZoneOperations = True
								' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
							End Select
					End If
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			Case Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_DC_ZoneOperations ] Invalied case [ " & sAction& " ].") 
	End Select

	If Fn_DC_ZoneOperations = True Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Function [ Fn_DC_ZoneOperations ] Executed successfully with case [ " & sAction& " ].") 
	End If

	Set objZone = Nothing
	Set objZoneDetails = Nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name			:	Fn_DC_SpatialSearchPanelOperations
'@@
'@@    Description				:	Function Used to perform operations related to Spatial Search Panel.
'@@
'@@    Parameters			    :	1. dicSpatial: dictionary object
'@@
'@@    Return Value		   	    : 	True Or False
'@@
'@@    Pre-requisite			:	Design Context perspective should be displayed
'@@
'@@    								
'@@    Examples					:	
'@@    								Dim dicSpatial
'@@    								Set dicSpatial = CreateObject("Scripting.Dictionary")
'@@    								 With dicSpatial
'@@    									.Add "Proximity","name" 
'@@    									.Add "bTrueShapeFiltering",True
'@@    									.Add "bValidOverlaysOnly",True 
'@@    									.Add "bAppendParts", False 
'@@    									.Add "bClickOnSearchButton", True
'@@    								End With
'@@    								Call Fn_DC_SpatialSearchPanelOperations("Create", dicSpatial)
'@@
'@@								Dim dicSpatial
'@@								Set dicSpatial = CreateObject("Scripting.Dictionary")
'@@								 With dicSpatial
'@@    									.Add "Proximity","name" 
'@@    									.Add "bTrueShapeFiltering",True
'@@    									.Add "bValidOverlaysOnly",True 
'@@    									.Add "bAppendParts", False 
'@@									   .Add "bClickOnSearchButton", True
'@@									  .Add "Name", "BoxZone~Zone_Pred"
'@@									  .Add "Operator", "Outside~Within"
'@@								End With
'@@								msgbox Fn_DC_SpatialSearchPanelOperations("Create", dicSpatial)
'@@								Dim dicSpatial
'@@								Set dicSpatial = CreateObject("Scripting.Dictionary")
'@@								With dicSpatial
'@@									.Add "Name", "BoxZone~Zone_Pred"
'@@									.Add "Operator", "Outside~Within"
'@@								End With
'@@								bReturn = Fn_DC_SpatialSearchPanelOperations("Verify", dicSpatial)
'@@
'@@									Set dicSpatial = CreateObject("Scripting.Dictionary")
'@@									With dicSpatial
'@@										.Add "Name", "Box890"
'@@										.Add "Value", "Within"
'@@										.Add "CloumnName", "Operator"
'@@									End With
'@@									bReturn=Fn_DC_SpatialSearchPanelOperations("Modify", dicSpatial)
'@@
'@@	   History					:
'@@				Developer Name				Date			Rev. No.	Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe			21-Feb-2012			1.0			Created
'@@				Sachin Joshi				12-SEPT-2012	  1.1			Modified Case "Create"
'@@				Sachin Joshi				12-SEPT-2012	  1.1			Added Case "Verify"
'@@				Sandeep Navghane				04-JUL-2013	  1.2			Added Case "Modify"
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_DC_SpatialSearchPanelOperations(sAction, dicSpatial)
	GBL_FAILED_FUNCTION_NAME="Fn_DC_SpatialSearchPanelOperations"
   Dim objApplet,ObjMDR, iRowCnt, iCnt, arrNames, arrOperators, iLstCount,iCounter,arrValues,arrColumn
   Dim iX,iY,iTempRowHight,iTempWidth,iWidth,iHight

	Fn_DC_SpatialSearchPanelOperations = False
	bFlag = False
	Set objApplet  = JavaWindow("DesignContextWindow").JavaWindow("DesignContextApplet")

	Select Case sAction
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Create"
			' setting proximity
			If dicSpatial("Proximity") <> "" Then
				call Fn_SISW_UI_JavaEdit_Operations("Fn_DC_SpatialSearchPanelOperations", "Set", objApplet, "Proximity", dicSpatial("Proximity"))
'				Call Fn_List_Select("Fn_DC_SpatialSearchPanelOperations", objApplet,"ProximityList",dicSpatial("Proximity"))
			End IF
			' setting True Shape Filtering
			If dicSpatial("bTrueShapeFiltering") <> "" Then
				If cBool( dicSpatial("bTrueShapeFiltering") ) Then
					Call Fn_CheckBox_Set("Fn_DC_SpatialSearchPanelOperations", objApplet,"TrueShapeFiltering","ON")
				Else
					Call Fn_CheckBox_Set("Fn_DC_SpatialSearchPanelOperations", objApplet,"TrueShapeFiltering","OFF")
				End If
			End IF
			' setting Valid overlays Only
			If dicSpatial("bValidOverlaysOnly") <> "" Then
				If cBool( dicSpatial("bValidOverlaysOnly") ) Then
					Call Fn_CheckBox_Set("Fn_DC_SpatialSearchPanelOperations", objApplet,"ValidOverlaysOnly","ON")
				Else
					Call Fn_CheckBox_Set("Fn_DC_SpatialSearchPanelOperations", objApplet,"ValidOverlaysOnly","OFF")
				End If
			End IF
			' Append Parts
			If dicSpatial("bAppendParts") <> "" Then
				If cBool( dicSpatial("bAppendParts") ) Then
					Call Fn_CheckBox_Set("Fn_DC_SpatialSearchPanelOperations", objApplet,"AppendParts","ON")
				Else
					Call Fn_CheckBox_Set("Fn_DC_SpatialSearchPanelOperations", objApplet,"AppendParts","OFF")
				End If
			End IF

			'Selecting Zone Filter Table Details
				arrNames = split(dicSpatial("Name"), "~")
				arrOperators = split(dicSpatial("Operator"), "~")
				iRowCnt = -1
				For iCnt = 0 to UBound(arrNames)
                   ' Click on + button
					wait 1
					Call Fn_Button_Click("Fn_DC_SpatialSearchPanelOperations", objApplet, "Add(+)")
					iRowCnt = iRowCnt + 1
					wait 2
					' Setting Name
					If trim(arrNames(iCnt)) <> "" Then
						objApplet.JavaTable("ZoneFilerTable").ClickCell iRowCnt,"Name","LEFT"
						Call Fn_List_Select("Fn_DC_SpatialSearchPanelOperations", objApplet, "ZoneFilterName", trim(arrNames(iCnt)))
					End If

					wait 2
					' Setting Operator
					If dicSpatial("Operator") <> "" Then
						objApplet.JavaTable("ZoneFilerTable").DoubleClickCell iRowCnt,"Operator","LEFT"
						For iLstCount = 0 to 2
							objApplet.JavaTable("ZoneFilerTable").ClickCell iRowCnt,"Name","LEFT"
							wait 1
							objApplet.JavaTable("ZoneFilerTable").ClickCell iRowCnt,"Operator","LEFT"
							wait 1
							objApplet.JavaTable("ZoneFilerTable").DoubleClickCell iRowCnt,"Operator","LEFT"
							wait 1
							If objApplet.JavaList("ZoneFilterOperator").Exist(2) Then
								Call Fn_List_Select("Fn_DC_SpatialSearchPanelOperations", objApplet, "ZoneFilterOperator", trim(arrOperators(iCnt)))
								Exit For
							End If
						Next
						If iLstCount = 3 Then
							objApplet.JavaTable("ZoneFilerTable").setCellData iRowCnt,"Operator", trim(arrOperators(iCnt))
						End If
					End If
					wait 2
				Next

            ' clicking on search button
			If dicSpatial("bClickOnSearchButton") = "" then 
				dicSpatial("bClickOnSearchButton") = True
			End If

			If cBool(dicSpatial("bClickOnSearchButton")) Then
				Call Fn_Button_Click("Fn_DC_SpatialSearchPanelOperations", objApplet, "Update")
			End If
			Fn_DC_SpatialSearchPanelOperations = True
			'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		Case "Modify1"		
				
				' setting proximity
			If dicSpatial("Proximity") <> "" Then
				call Fn_SISW_UI_JavaEdit_Operations("Fn_DC_SpatialSearchPanelOperations", "Set", objApplet, "Proximity", dicSpatial("Proximity"))
'				Call Fn_List_Select("Fn_DC_SpatialSearchPanelOperations", objApplet,"ProximityList",dicSpatial("Proximity"))
			End IF
			' setting True Shape Filtering
			If dicSpatial("bTrueShapeFiltering") <> "" Then
				If cBool( dicSpatial("bTrueShapeFiltering") ) Then
					Call Fn_CheckBox_Set("Fn_DC_SpatialSearchPanelOperations", objApplet,"TrueShapeFiltering","ON")
				Else
					Call Fn_CheckBox_Set("Fn_DC_SpatialSearchPanelOperations", objApplet,"TrueShapeFiltering","OFF")
				End If
			End IF
			' setting Valid overlays Only
			If dicSpatial("bValidOverlaysOnly") <> "" Then
				If cBool( dicSpatial("bValidOverlaysOnly") ) Then
					Call Fn_CheckBox_Set("Fn_DC_SpatialSearchPanelOperations", objApplet,"ValidOverlaysOnly","ON")
				Else
					Call Fn_CheckBox_Set("Fn_DC_SpatialSearchPanelOperations", objApplet,"ValidOverlaysOnly","OFF")
				End If
			End IF
			' Append Parts
			If dicSpatial("bAppendParts") <> "" Then
				If cBool( dicSpatial("bAppendParts") ) Then
					Call Fn_CheckBox_Set("Fn_DC_SpatialSearchPanelOperations", objApplet,"AppendParts","ON")
				Else
					Call Fn_CheckBox_Set("Fn_DC_SpatialSearchPanelOperations", objApplet,"AppendParts","OFF")
				End If
			End IF

			'Selecting Zone Filter Table Details
				arrNames = split(dicSpatial("Name"), "~")
				arrOperators = split(dicSpatial("Operator"), "~")
				iRowCnt = -1
				For iCnt = 0 to UBound(arrNames)
                   ' Click on + button
					wait 1
					'Call Fn_Button_Click("Fn_DC_SpatialSearchPanelOperations", objApplet, "Add(+)")
					iRowCnt = iRowCnt + 1
					wait 2
					' Setting Name
					If trim(arrNames(iCnt)) <> "" Then
						objApplet.JavaTable("ZoneFilerTable").ClickCell iRowCnt,"Name","LEFT"
						Call Fn_List_Select("Fn_DC_SpatialSearchPanelOperations", objApplet, "ZoneFilterName", trim(arrNames(iCnt)))
					End If

					wait 2
					' Setting Operator
					If dicSpatial("Operator") <> "" Then
						objApplet.JavaTable("ZoneFilerTable").DoubleClickCell iRowCnt,"Operator","LEFT"
						For iLstCount = 0 to 2
							objApplet.JavaTable("ZoneFilerTable").ClickCell iRowCnt,"Name","LEFT"
							wait 1
							objApplet.JavaTable("ZoneFilerTable").ClickCell iRowCnt,"Operator","LEFT"
							wait 1
							objApplet.JavaTable("ZoneFilerTable").DoubleClickCell iRowCnt,"Operator","LEFT"
							wait 1
							If objApplet.JavaList("ZoneFilterOperator").Exist(2) Then
								Call Fn_List_Select("Fn_DC_SpatialSearchPanelOperations", objApplet, "ZoneFilterOperator", trim(arrOperators(iCnt)))
								Exit For
							End If
						Next
						If iLstCount = 3 Then
							objApplet.JavaTable("ZoneFilerTable").setCellData iRowCnt,"Operator", trim(arrOperators(iCnt))
						End If
					End If
					wait 2
				Next

            ' clicking on search button
			If dicSpatial("bClickOnSearchButton") = "" then 
				dicSpatial("bClickOnSearchButton") = True
			End If

			If cBool(dicSpatial("bClickOnSearchButton")) Then
				Call Fn_Button_Click("Fn_DC_SpatialSearchPanelOperations", objApplet, "Update")
			End If
			Fn_DC_SpatialSearchPanelOperations = True
			
	    ' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Modify"
			arrNames = split(dicSpatial("Name"), "~")
			arrValues= split(dicSpatial("Value"), "~")	
			arrColumn=split(dicSpatial("ColumnName"),"~")
            
			iX=objApplet.JavaTable("ZoneFilerTable").GetROProperty("abs_x")
			iY=objApplet.JavaTable("ZoneFilerTable").GetROProperty("abs_y")
			iTempRowHight=objApplet.JavaTable("ZoneFilerTable").Object.getRowHeight()
            
			For iCnt = 0 to UBound(arrNames)
				For iCount=0 to cint(objApplet.JavaTable("ZoneFilerTable").GetROProperty("cols"))-1
					iTempWidth=objApplet.JavaTable("ZoneFilerTable").Object.getColumnModel().getColumn(iCount).getWidth()
					If arrColumn(iCnt)=objApplet.JavaTable("ZoneFilerTable").GetColumnName(iCount) Then
						iTempWidth=iWidth/2
					End If
					iWidth=iWidth+iTempWidth
				Next
				bFlag=False
				For iCounter=0 to cint(objApplet.JavaTable("ZoneFilerTable").GetROProperty("rows"))-1

					If Trim(objApplet.JavaTable("ZoneFilerTable").GetCellData(iCounter,"Name"))=Trim(arrNames(iCnt)) Then
						iHight=iTempRowHight*(iCounter+1)-iTempRowHight/2
						
						For iLstCount = 0 to 2
							Set ObjMDR=CreateObject("Mercury.DeviceReplay")
							ObjMDR.MouseDblClick iX+iWidth,iY+iHight,0
							wait 2
							Set ObjMDR=Nothing

							If objApplet.JavaList("ZoneFilterOperator").Exist(2) Then
								Call Fn_List_Select("Fn_DC_SpatialSearchPanelOperations", objApplet, "ZoneFilterOperator", trim(arrValues(iCnt)))
								bFlag=True
								Exit For
							End If
						Next
					End If
				Next
				If bFlag=False Then
					Exit for
				End If
			Next
			If bFlag=True Then
				Fn_DC_SpatialSearchPanelOperations = True
				Call Fn_Button_Click("Fn_DC_SpatialSearchPanelOperations", objApplet, "Update")
			End If
        ' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
        Case "Verify"
				'Selecting Zone Filter Table Details
				arrNames = split(dicSpatial("Name"), "~")
				arrOperators = split(dicSpatial("Operator"), "~")
				iRowCnt = -1
				wait 1
				Call Fn_Button_Click("Fn_DC_SpatialSearchPanelOperations", objApplet, "Add(+)")
				iRowCnt = iRowCnt + 1
				For iCnt = 0 to UBound(arrNames)
                   ' Click on + button
					
					wait 2
					' Setting Name
					If trim(arrNames(iCnt)) <> "" Then
						objApplet.JavaTable("ZoneFilerTable").ClickCell iRowCnt,"Name","LEFT"
'						Call Fn_List_Select("Fn_DC_SpatialSearchPanelOperations", objApplet, "ZoneFilterName", trim(arrNames(iCnt)))
                        Fn_DC_SpatialSearchPanelOperations = Fn_UI_ListItemExist("Fn_DC_SpatialSearchPanelOperations", objApplet, "ZoneFilterName",trim(arrNames(iCnt)))		
					End If

					wait 2
					' Setting Operator
					If dicSpatial("Operator") <> "" Then
						objApplet.JavaTable("ZoneFilerTable").DoubleClickCell iRowCnt,"Operator","LEFT"
						For iLstCount = 0 to 2
							objApplet.JavaTable("ZoneFilerTable").ClickCell iRowCnt,"Name","LEFT"
							wait 1
							objApplet.JavaTable("ZoneFilerTable").DoubleClickCell iRowCnt,"Operator","LEFT"
							wait 1
							If objApplet.JavaList("ZoneFilterOperator").Exist(2) Then
								Fn_DC_SpatialSearchPanelOperations = Fn_UI_ListItemExist("Fn_DC_SpatialSearchPanelOperations", objApplet, "ZoneFilterOperator",trim(arrOperators(iCnt)))
								Exit For
							End If
						Next
					End If
					wait 2
				Next
			wait 1
			Call Fn_Button_Click("Fn_DC_SpatialSearchPanelOperations", objApplet, "Remove(-)")
			wait 2
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_DC_SpatialSearchPanelOperations ] Invalied case [ " & sAction& " ].") 
	End Select

	If Fn_DC_SpatialSearchPanelOperations = True Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Function [ Fn_DC_SpatialSearchPanelOperations ] Executed successfully with case [ " & sAction& " ].") 
	End If

	Set objApplet = Nothing
End Function

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name			:	Fn_SISW_DC_SaveStructureContextObject
'@@
'@@    Description				:	Function Used to Save Structure Context Object.
'@@
'@@    Parameters			    :	1. sName			: Name string
'@@									2. sDescription		: Description string
'@@									3. sType			: Type
'@@									4. bAddToClipboard	: Boolean value to select Add To Clipboard checkbox
'@@									5. bAddToNewstuff	: Boolean value to select Add To Newstuff checkbox
'@@
'@@    Return Value		   	    : 	True Or False
'@@
'@@    Pre-requisite			:	Design Context perspective should be displayed
'@@
'@@    								
'@@    Examples					:	
'@@								bReturn = Fn_SISW_DC_SaveStructureContextObject("name", "desc", "VisStructureContext", True, False)
'@@
'@@	   History					:
'@@				Developer Name				Date			Rev. No.	Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe			17-Sept-2012			1.0			Created
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_SISW_DC_SaveStructureContextObject(sName, sDescription, sType, bAddToClipboard, bAddToNewstuff)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_DC_SaveStructureContextObject"
	Dim objDialog
	Fn_SISW_DC_SaveStructureContextObject = False
	Set objDialog = JavaWindow("DesignContextWindow").JavaWindow("DesignContextApplet").JavaDialog("SaveStructureContext")
	 
	If objDialog.Exist(5) = False Then
		'Old menu
'		Call Fn_MenuOperation("Select", "File:Save Structure Context Object")
		'New menu
		Call Fn_MenuOperation("Select", "File:Save")
		If objDialog.Exist(15) = False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Function - [ Fn_SISW_DC_SaveStructureContextObject ] Failed to find [Save Structure Context Object] window.")
			Exit function
		End If
	End If

	objDialog.JavaEdit("Name").Type sName
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS : Function - [ Fn_SISW_DC_SaveStructureContextObject ] Successfully Set Name = " & sName)

	objDialog.JavaEdit("Description").Type sDescription
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS : Function - [ Fn_SISW_DC_SaveStructureContextObject ] Successfully Set Description = " & sDescription)

	Call Fn_List_Select("Fn_DC_FormAttributeSearchPanelOperations", objDialog, "Type", sType)
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS : Function - [ Fn_SISW_DC_SaveStructureContextObject ] Successfully Set Type = " & sType)

	If bAddToClipboard <> "" Then
		If cBool(bAddToClipboard) Then
			objDialog.JavaCheckBox("AddToClipboard").Set "ON"
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS : Function - [ Fn_SISW_DC_SaveStructureContextObject ] Successfully Set Add To Clipboard = ON ")
		Else
			objDialog.JavaCheckBox("AddToClipboard").Set "OFF"
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS : Function - [ Fn_SISW_DC_SaveStructureContextObject ] Successfully Set Add To Clipboard = OFF ")
		End If
	End If
	If bAddToNewstuff <> "" Then
		If cBool(bAddToNewstuff) Then
			objDialog.JavaCheckBox("AddToNewstuff").Set "ON"
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS : Function - [ Fn_SISW_DC_SaveStructureContextObject ] Successfully Set Add To Newstuff = ON ")
		Else
			objDialog.JavaCheckBox("AddToNewstuff").Set "OFF"
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS : Function - [ Fn_SISW_DC_SaveStructureContextObject ] Successfully Set Add To Newstuff = OFF ")
		End If
	End If
	Fn_SISW_DC_SaveStructureContextObject = Fn_Button_Click("Fn_DC_SpatialSearchPanelOperations", objDialog, "OK")
	Call Fn_ReadyStatusSync(2)
	If JavaWindow("DesignContextWindow").JavaWindow("DesignContextApplet").JavaDialog("DesignContextApplicationSave").Exist(5) Then
		Call Fn_Button_Click("Fn_DC_SpatialSearchPanelOperations",  JavaWindow("DesignContextWindow").JavaWindow("DesignContextApplet").JavaDialog("DesignContextApplicationSave"), "OK")
	End If
    Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS : Function - [ Fn_SISW_DC_SaveStructureContextObject ] executed successfully.")
	Set objDialog = Nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name			:	Fn_SISW_DC_FormAttributeSearchPanelOperations
'@@
'@@    Description				:	Function Used to set values for Form Attribute in Design Context.
'@@
'@@    Parameters			    :	1. sAction : Action to perform
'@@									2. bClearFormAttributePanel : True / False value to click on clear button
'@@									3. sLogicalOperator	: Logical operator seprated by ~ ( OR / AND )
'@@									4. sRelationType : List of Relation Types separated by ~
'@@									5. sParentType : List of Parent Types separated by ~
'@@									6. sFormType : List of Form Types separated by ~
'@@									7. sPropertyName : List of Property Names separated by ~
'@@									8. sOperator : List of Operators separated by ~
'@@									9. sSearchingValue : List of Searching Value separated by ~
'@@									10. bUpdateBtnClick : True / False value to click on Update button
'@@
'@@    Return Value		   	    : 	True Or False
'@@
'@@    Pre-requisite			:	Design Context perspective should be displayed
'@@
'@@    								
'@@    Examples					:	
'@@								bReturn = Fn_SISW_DC_FormAttributeSearchPanelOperations("", "", "~AND", "Specifications~", "ItemRevision~", "ItemRevision Master~", "User Data 3~User Data 3", "GT~GE", "123~100", "")
'@@
'@@	   History					:
'@@				Developer Name				Date			Rev. No.	Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe			24-Sept-2012			1.0			Created
'@@				Sachin							25-Sept-2012			1.1			Added Case "VerifyFormAttribute"
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_SISW_DC_FormAttributeSearchPanelOperations(sAction, bClearFormAttributePanel, sLogicalOperator, sRelationType, sParentType, sFormType, sPropertyName, sOperator, sSearchingValue, bUpdateBtnClick)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_DC_FormAttributeSearchPanelOperations"
	Dim objApplet, iCounter, iRowCnt
	Dim aLogicalOperator, aRelationType, aParentType, aFormType, aPropertyName, aOperator, aSearchValue,bFlag
	Fn_SISW_DC_FormAttributeSearchPanelOperations = False
	Set objApplet = JavaWindow("DesignContextWindow").JavaWindow("DesignContextApplet")
	objApplet.JavaStaticText("Field").SetTOProperty "label", "Form Attributes"
	If objApplet.JavaList("RelationTypeList").Exist(3) = False Then
		objApplet.JavaStaticText("Field").Click 1, 1, "LEFT"
	End If
	' clearing Form Attribute details
	If bClearFormAttributePanel = "" then bClearFormAttributePanel = "False"
	If cBool(bClearFormAttributePanel) then
		Call Fn_Button_Click("Fn_SISW_DC_FormAttributeSearchPanelOperations", objFormAttrib, "Clear")
	End If

	Select Case sAction
		Case "SetFormAttribute"
			arrLogicalOperator = split(sLogicalOperator,"~") 
			arrRelationTypes = split(sRelationType,"~") 
			arrParentTypes = split(sParentType,"~") 
			arrFormTypes = split(sFormType,"~") 
			arrPropertyNames = split(sPropertyName,"~") 
			arrOperators = split(sOperator,"~") 
			arrSearchingValues = split(sSearchingValue,"~")

			For iCnt = 0 to UBound(arrSearchingValues)

				' selecting relation type
				If trim(arrRelationTypes(iCnt)) <> "" then
					Call Fn_List_Select("Fn_SISW_DC_FormAttributeSearchPanelOperations", objApplet, "RelationTypeList", trim(arrRelationTypes(iCnt)))
				End If
	
				' selecting Parent type
				If trim(arrParentTypes(iCnt)) <> "" then
					Call Fn_List_Select("Fn_SISW_DC_FormAttributeSearchPanelOperations", objApplet, "ParentTypeList", trim(arrParentTypes(iCnt)))
				End If
	
				' selecting Form type
				If trim(arrFormTypes(iCnt)) <> "" then
					Call Fn_List_Select("Fn_SISW_DC_FormAttributeSearchPanelOperations", objApplet, "FormTypeList", trim(arrFormTypes(iCnt)))
				End If
	
				' clicking on Add button
				Call Fn_Button_Click("Fn_SISW_DC_FormAttributeSearchPanelOperations", objApplet, "FormTypeAdd")
	
				If iCnt <> 0 Then
					If trim(arrLogicalOperator(iCnt)) <> "" then
						objApplet.JavaTable("FormTypeTable").ClickCell iRowCnt,"Logical Operator","LEFT"
						Call Fn_List_Select("Fn_SISW_DC_FormAttributeSearchPanelOperations", objApplet, "FormTypeTableList", trim(arrLogicalOperator(iCnt)))
					End If
				End If
				Wait 2
				iRowCnt = cInt(objApplet.JavaTable("FormTypeTable").GetROProperty("rows")) - 1
				If trim(arrPropertyNames(iCnt)) <> "" then
					objApplet.JavaTable("FormTypeTable").ClickCell iRowCnt,"Property Name","LEFT"
					Call Fn_List_Select("Fn_SISW_DC_FormAttributeSearchPanelOperations", objApplet, "FormTypeTableList", trim(arrPropertyNames(iCnt)))
				End If
	
				If trim(arrOperators(iCnt)) <> "" then
					objApplet.JavaTable("FormTypeTable").ClickCell iRowCnt,"Operator","LEFT"
					Call Fn_List_Select("Fn_SISW_DC_FormAttributeSearchPanelOperations", objApplet, "FormTypeTableList", trim(arrOperators(iCnt)))
				End If
	
				If trim(arrSearchingValues(iCnt)) <> "" then
					objApplet.JavaTable("FormTypeTable").SetCellData iRowCnt,"Searching Value", trim(arrSearchingValues(iCnt)) 
				End If
			Next

			Fn_SISW_DC_FormAttributeSearchPanelOperations = True
			If bUpdateBtnClick <> "" Then
				If cBool(bUpdateBtnClick) Then
					Call Fn_Button_Click("Fn_SISW_DC_FormAttributeSearchPanelOperations", objApplet, "Update")
				End If
			End If
		Case "VerifyFormAttribute"
			arrLogicalOperator = split(sLogicalOperator,"~")
			arrRelationTypes = split(sRelationType,"~") 
			arrParentTypes = split(sParentType,"~") 
			arrFormTypes = split(sFormType,"~") 
			arrPropertyNames = split(sPropertyName,"~") 
			arrOperators = split(sOperator,"~") 
			arrSearchingValues = split(sSearchingValue,"~")

			bFlag = False

			For iCnt = 0 to UBound(arrSearchingValues)
				If sLogicalOperator <> "" Then
					If trim(arrLogicalOperator(iCnt)) <> "" Then
                    	If Cstr(trim(arrLogicalOperator(iCnt))) = Cstr(objApplet.JavaTable("FormTypeTable").GetCellData(iCnt,"Logical Operator")) Then 
							bFlag = True
						Else
							bFlag = False
							Exit For
						End If
					End If
				End If

				If sRelationType <> "" Then
					If trim(arrRelationTypes(iCnt)) <> "" Then
                    	If Cstr(trim(arrRelationTypes(iCnt))) = Cstr(objApplet.JavaTable("FormTypeTable").GetCellData(iCnt,"Relation Type")) Then 
							bFlag = True
						Else
							bFlag = False
							Exit For
						End If
						End If
				End If

				If sParentType <> "" Then
					If trim(arrParentTypes(iCnt)) <> "" Then
                    	If Cstr(trim(arrParentTypes(iCnt))) = Cstr(objApplet.JavaTable("FormTypeTable").GetCellData(iCnt,"Parent Type")) Then 
							bFlag = True
						Else
							bFlag = False
							Exit For
						End If
						End If
				End If

				If sFormType <> "" Then
					If trim(arrFormTypes(iCnt)) <> "" Then
                    	If Cstr(trim(arrFormTypes(iCnt))) = Cstr(objApplet.JavaTable("FormTypeTable").GetCellData(iCnt,"Form Type")) Then 
							bFlag = True
						Else
							bFlag = False
							Exit For
						End If
						End If
				End If

				If sPropertyName <> "" Then	
					If trim(arrPropertyNames(iCnt)) <> "" Then
						If Cstr(trim(arrPropertyNames(iCnt))) = Cstr(objApplet.JavaTable("FormTypeTable").GetCellData(iCnt,"Property Name")) Then 
							bFlag = True
						Else
							bFlag = False
							Exit For
						End If
					End If
					End If

				If sOperator <> "" Then
					If trim(arrOperators(iCnt)) <> "" Then
						If Cstr(trim(arrOperators(iCnt))) = Cstr(objApplet.JavaTable("FormTypeTable").GetCellData(iCnt,"Operator")) Then 
							bFlag = True
						Else
							bFlag = False
							Exit For
						End If
						End If
				End If

				If sSearchingValue <> "" Then
					If trim(arrSearchingValues(iCnt)) <> "" Then
						If Cstr(trim(arrSearchingValues(iCnt))) = Cstr(objApplet.JavaTable("FormTypeTable").GetCellData(iCnt,"Searching Value")) Then
							bFlag = True
						Else
							bFlag = False
							Exit For
						End If
						End If
				End If
			Next
			If bFlag Then
				Fn_SISW_DC_FormAttributeSearchPanelOperations = True
			Else
				Fn_SISW_DC_FormAttributeSearchPanelOperations = false
			End If

		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_SISW_DC_FormAttributeSearchPanelOperations ] Invalied case [ " & sAction& " ].") 
	End Select

	If Fn_SISW_DC_FormAttributeSearchPanelOperations = True Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Function [ Fn_SISW_DC_FormAttributeSearchPanelOperations ] Executed successfully with case [ " & sAction& " ].") 
	End If

	Set objApplet = Nothing
End Function

'*******************************************************************************
'
''Function Name		 	:	Fn_SISW_DC_ExcuteSCOSearch
'
''Description		    :  	Function to perform SCO search operation in Design Context.

''Parameters		    :	1. sSearchType : Search Type
								
''Return Value		    :  	True \ False
'
''Examples		     	:	Fn_SISW_DC_ExcuteSCOSearch("SCO Evaluation Dynamic")

'History:
'	Developer Name			Date			Rev. No.		Reviewer		Changes Done	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Koustubh Watwe		 24-Sept-2012		1.0				
'-----------------------------------------------------------------------------------------------------------------------------------
Function Fn_SISW_DC_ExcuteSCOSearch(sSearchType)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_DC_ExcuteSCOSearch"
	Dim objApplet
	Fn_SISW_DC_ExcuteSCOSearch = False
	Set objApplet = Fn_SISW_DC_GetObject("DesignContextApplet")
	
	If Fn_UI_ListItemExist("Fn_SISW_DC_ExcuteSCOSearch", objApplet, "BottomRightComboBox", sSearchType ) Then
		Call Fn_List_Select("Fn_SISW_DC_ExcuteSCOSearch", objApplet, "BottomRightComboBox", sSearchType )
		Fn_SISW_DC_ExcuteSCOSearch = Fn_Button_Click("Fn_SISW_DC_ExcuteSCOSearch", objApplet,"executesearch_16")
	End If
	Set objApplet = Nothing
End Function
