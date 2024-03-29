Option Explicit
iTimeOut = 40

'---------------------------------------------------------	Function List		------------------------------------------------------------------------------------------------------------------------------------------------
'000 Fn_SISW_ABM_GetObject()
'001 Fn_ABM_RootStructureTableColumnOperations()
'002 Fn_ABM_RootStructureTableRowIndex()
'003 Fn_ABM_RootStructureTableOperations()
'004 Fn_ABM_LotOperations()
'005 Fn_ASB_NavTree_NodeOperation()
'006 Fn_ABM_SerialNoGenerator()
'007 Fn_ABM_GenerateAsBuiltStructure()
'008 Fn_ABM_AssignLot()
'009 Fn_SISW_ABM_SearchDialogOperations()
'010 Fn_SISW_ABM_InstallPhysicalPartOperations()
'011 Fn_SISW_ABM_ReplacePhysicalPartOperations()
'012 Fn_SISW_ABM_UnInstallPhysicalPartOperations()
'013 Fn_SISW_ABM_RebuildAsBuiltStructure()
'014 Fn_SISW_ABM_SetupDeviationOperations()
'015 Fn_SISW_ABM_RebasePhysicalPartOperations()
'016 Fn_SISW_ABM_RenamePhysicalPartOperations()
'017 Fn_SISW_ABM_AsBuildCompareOperations()
'018 Fn_SISW_ABM_DuplicateAsBuiltStructure
'****************************************    Function to perform Find and Select Physical Part ***************************************
'
''Function Name		 	:	Fn_SISW_ABM_GetObject
'
''Description		    :  	Function to get objects of As Build Manager / MRO

''Parameters		    :	1. sObjectName : Object Handle name
								
''Return Value		    :  	Object \ Nothing
'
''Examples		     	:	Fn_SISW_ABM_GetObject("InstallPhysicalPart")

'History:
'	Developer Name			Date			Rev. No.		Reviewer		Changes Done	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Koustubh Watwe		 31-May-2012		1.0				

'-----------------------------------------------------------------------------------------------------------------------------------
'	Ashwini Kumar		 25-Sept-2013		2.0								Externalized object hierarchies
'-----------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_ABM_GetObject(sObjectName)
	Dim sObjectXMLPath
	sObjectXMLPath = Fn_GetEnvValue("User", "AutomationDir") & "\TestData\AutomationXML\ObjectXML\AsBuildManager.xml"
	Set Fn_SISW_ABM_GetObject = Fn_SISW_Setup_GetObjectFromXML(sObjectXMLPath, sObjectName)
End Function
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
''****************************************    Function to perform operatiosn fon columnsin MRO ***************************************
'
''Function Name		 	:			  Fn_ABM_RootStructureTableColumnOperations
'
''Description		    :  	      Function to perform operatiosn fon columns in MRO 
'
''Parameters		    :	 	1. sAction : Action need to perform
'					   			2. sColumnNames : Column's name
'								
''Return Value		    :  		True \ False | column nhumber \ -1
'
''Pre-requisite		    :		MRO perspective should be selected

''Examples		     	:	Call  Fn_ABM_RootStructureTableColumnOperations("GetIndex", "Lot")
''Examples		     	:	Call  Fn_ABM_RootStructureTableColumnOperations("Remove", "Lot")
''Examples		     	:	Call  Fn_ABM_RootStructureTableColumnOperations("Add", "Lot")

'History:
'						Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'						   Koustubh Watwe		5-Sept-2011			1.0					 Created
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public function Fn_ABM_RootStructureTableColumnOperations(sAction, sColumnNames)
	GBL_FAILED_FUNCTION_NAME="Fn_ABM_RootStructureTableColumnOperations"
	Dim iColIndex, objTable, objParentObject, strMenu, iCols, ArrCol
	Dim sColToAdd, iIndex, objChangeColumnDialog, objList,intCol
	Fn_ABM_RootStructureTableColumnOperations = False
	Set objParentObject = JavaWindow("AsBuiltManager").JavaApplet("JApplet")
	Set objTable = JavaWindow("AsBuiltManager").JavaApplet("JApplet").JavaTable("RootStructuresTable")
	objParentObject.JavaObject("RootStructureTablePanel").Click 0,0, "LEFT" 

	Select Case sAction
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			Case "GetIndex"
				iCols = Cint(objTable.GetROProperty("cols"))
				Fn_ABM_RootStructureTableColumnOperations = -1
				For iColIndex =0 to iCols - 1
					If objTable.GetColumnName(iColIndex) = sColumnNames Then
						Fn_ABM_RootStructureTableColumnOperations = iColIndex
						Exit for
					End If
				Next
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			Case "Add"
                        ArrCol = Split(sColumnNames,":",-1,1)
				sColToAdd = ""
				 For iIndex = 0 To Ubound(ArrCol)
						'Check that Column is present in the BOMTable.
						iColIndex =  Fn_ABM_RootStructureTableColumnOperations("GetIndex", ArrCol(iIndex))		
						If iColIndex = -1 Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Warning: Column does not  exist in the Application.Need to Add Column ["& ArrCol(iIndex) &"]." )
								sColToAdd = sColToAdd +":"+ArrCol(iIndex)
						Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Column ["& ArrCol(iIndex) &"] exists in the Application" )
								Fn_ABM_RootStructureTableColumnOperations =TRUE
						End if
				Next
				If sColToAdd <>""  Then
						sColToAdd = Mid(sColToAdd, 2,Len(sColToAdd))
						ArrCol = Split(sColToAdd,":",-1,1)
						'Invoke Choose Column Window if it is not present on the screen
						Set objChangeColumnDialog = JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Change Columns")
						If NOT objChangeColumnDialog.Exist( 1)  Then
								objTable.SelectColumnHeader "#1","RIGHT"       	
								objParentObject.JavaMenu("label:=Insert column\(s\) ...").Select 										       
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: RMB action Insert Column(s).... Executed successfully in the Application.")			
								Set objList = objChangeColumnDialog.JavaList("ListAvailableCols").Object
						End If
						'Synchronization for Ready state
						Call Fn_ReadyStatusSync(2)
						For iIndex = 0 To Ubound(ArrCol)							
								'Select Col to be added from the lsit
								intCol = objChangeColumnDialog.JavaList("ListAvailableCols").GetItemIndex(ArrCol(iIndex))
								objList.ensureIndexIsVisible intCol
								objChangeColumnDialog.JavaList("ListAvailableCols").ExtendSelect ArrCol(iIndex)
						Next
						' Hit  Add Column  Button after every Column selection
						Call Fn_Button_Click("Fn_ABM_RootStructureTableColumnOperations",objChangeColumnDialog, "Add")
						' Hit  Apply Button after selection
						Call Fn_Button_Click("Fn_ABM_RootStructureTableColumnOperations",objChangeColumnDialog, "Apply")
						Call Fn_Button_Click("Fn_ABM_RootStructureTableColumnOperations",objChangeColumnDialog, "Cancel")
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Pass:Successfully Added  Column  ["& sColToAdd &"] in BOMTable")									
						Fn_ABM_RootStructureTableColumnOperations = TRUE					
				End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			Case "Remove"
					ArrCol = Split(sColumnNames,":",-1,1)
					For iIndex = 0 To Ubound(ArrCol)										
							'Check that Column is present in the BOMTable
							iColIndex =  Fn_ABM_RootStructureTableColumnOperations("GetIndex", ArrCol(iIndex))						
							If iColIndex = -1 Then							
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"WARNING:Column dose not  exist in the Application.No Need to Remove Column ["& ArrCol(iIndex) &"]")
									Fn_ABM_RootStructureTableColumnOperations  = FALSE
							Else
								'Remove the given Colum.													
								objTable.SelectColumnHeader iColIndex,"RIGHT"
								objParentObject.JavaMenu("label:=Remove this column").Select		
								Call Fn_Button_Click("Fn_ABM_RootStructureTableColumnOperations",JavaWindow("DefaultWindow").JavaWindow("RemoveColumn"),"Yes")
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Pass: Successfully removed Column  ["& ArrCol(iIndex) &"] from BOMTable.")          																
								Fn_ABM_RootStructureTableColumnOperations  =TRUE										 						
							End if
					Next
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	End Select
	Set objParentObject = Nothing
	Set objTable = Nothing
	Set objChangeColumnDialog = Nothing
	Set objList = Nothing
End Function

''****************************************    Function to get Row INdex in MRO ***************************************
'
''Function Name		 	:			  Fn_ABM_RootStructureTableColumnOperations
'
''Description		    :  	      Function to to get Row INdex in MRO 
'
''Parameters		    :	 	1. objTable : Action need to perform
'					   			2. sNodeName : Root Structure Node Path
'								
''Return Value		    :  		row nhumber \ -1
'
''Pre-requisite		    :		MRO perspective should be selected

''Examples		     	:	Call  Fn_ABM_RootStructureTableRowIndex(objTable, "000554/A;1-TopPart (View):000555/A;1-Ch1 (View)")

'History:
'						Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'						   Koustubh Watwe		5-Sept-2011			1.0					 Created
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public function Fn_ABM_RootStructureTableRowIndex(objTable, sNodeName)
	GBL_FAILED_FUNCTION_NAME="Fn_ABM_RootStructureTableRowIndex"
	Dim nodeArr, aRowNode, iColIndex, aPath
	Dim iRowCounter, sNode, iInstance, iNodeCounter, iPathCounter, bFound 
	Dim iRows, sNodePath, sPath, StrNodePath
	sPath = ""

	If Fn_UI_ObjectExist("Fn_ABM_RootStructureTableRowIndex", objTable) = False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_ABM_RootStructureTableRowIndex ] Table does not exist.")	
		Fn_ABM_RootStructureTableRowIndex = -1
		Exit function
	End If
	iColIndex = 0
	bFound = False
	If sNodeName <> "" Then
		' identifying RowId
		iRows = cInt(objTable.GetROProperty ("rows"))
		nodeArr = split(sNodeName , ":")
		iRowCounter = 0
		For iNodeCounter=0 to UBound(nodeArr)
				aRowNode = split(trim((nodeArr(iNodeCounter))),"@")
				If sPath = "" Then
							sPath =  trim(aRowNode(0))
				Else
							sPath = sPath &":"& trim(aRowNode(0))
				End If
		Next
		For iNodeCounter=0 to UBound(nodeArr)
			If iRowCounter = iRows  Then
				Exit for
			End If
			aRowNode = split(trim((nodeArr(iNodeCounter))),"@")
			iInstance = 0
			bFound = False
			do While iRowCounter < iRows
				If uBound(aRowNode) > 0 Then
					' instance number exist in name
					' initialize instance num
					' ith row matches with aRowNode(0) then
					sNodePath = objTable.object.getValueAt(iRowCounter, iColIndex).toString()
					If trim(sNodePath) = trim(aRowNode(0)) then
							StrNodePath =objTable.Object.getPathForRow(iRowCounter).toString()
							StrNodePath = Right(StrNodePath, (Len(StrNodePath)-Instr(1, StrNodePath, ",", 1)))					
							StrNodePath = Left(StrNodePath, Len(StrNodePath)-1)
							If instr(StrNodePath, "@BOM::") > 0 Then
								StrNodePath = trim(replace(StrNodePath,"""",""))
								aPath = split(StrNodePath,",")
								StrNodePath = ""
								For icnt = 0 to uBound(aPath)
									aPath(iCnt) = Left(aPath(iCnt), instr(aPath(iCnt),"@")-1)
									If StrNodePath = "" Then
										StrNodePath = trim(aPath(iCnt))
									else
										StrNodePath = StrNodePath & ", " & trim(aPath(iCnt))
									End If
								Next
							End If

							StrNodePath = trim(replace(StrNodePath,", ",":"))
							If instr(sPath, StrNodePath ) > 0 Then
								iInstance = iInstance +1
								If iInstance = cInt(aRowNode(1)) Then 
										If UBound(nodeArr) = iNodeCounter Then
												bFound = True
										End If
										Exit do
								End If
								'exit loop
							End If
					End if
				Else
					'ith row matches with aRowNode(0) then
					sNodePath = objTable.object.getValueAt(iRowCounter, iColIndex).toString()
					If trim(sNodePath) = trim(aRowNode(0)) then
						StrNodePath =objTable.Object.getPathForRow(iRowCounter).toString()
						StrNodePath = Right(StrNodePath, (Len(StrNodePath)-Instr(1, StrNodePath, ",", 1)))					
						StrNodePath = Left(StrNodePath, Len(StrNodePath)-1)
						If instr(StrNodePath, "@BOM::") > 0 Then
							StrNodePath = trim(replace(StrNodePath,"""",""))
							aPath = split(StrNodePath,",")
							StrNodePath = ""
							For icnt = 0 to uBound(aPath)
								aPath(iCnt) = Left(aPath(iCnt), instr(aPath(iCnt),"@")-1)
								If StrNodePath = "" Then
									StrNodePath = trim(aPath(iCnt))
								else
									StrNodePath = StrNodePath & ", " & trim(aPath(iCnt))
								End If
							Next
						End If
						StrNodePath = trim(replace(StrNodePath,", ",":"))
						If instr(sPath, StrNodePath ) > 0 Then
								If UBound(nodeArr) = iNodeCounter Then
									bFound = True
								End If
								Exit do
								'exit loop
						End if
					End if
				End If
				iRowCounter = iRowCounter + 1
				' increment counter
			loop
		Next
	End If
	If bFound Then
		Fn_ABM_RootStructureTableRowIndex = iRowCounter
	Else
		Fn_ABM_RootStructureTableRowIndex = -1
	End If
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_ABM_RootStructureTableRowIndex ] executed successfully.")
End Function
''****************************************    Function to perform operations on Root Structure Table in MRO ***************************************
'
''Function Name		 	:			  Fn_ABM_RootStructureTableColumnOperations
'
''Description		    :  	      Function to perform operations on Root Structure Table in MRO 
'
''Parameters		    :	 	1. objTable : Action need to perform
'					   			2. sRootStructureHeader : Root Structure header tab label
'					   			3. sRootStructure : Root Structure item
'					   			4. sNodeName : Root Structure Node Path
'					   			5. sColName : column name
'					   			6. sValue : value 
'					   			7. sPopupMenu : Popup menu to select
'								
''Return Value		    :  		row nhumber \ -1
'
''Pre-requisite		    :		MRO perspective should be selected

'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Examples		     	:	Call  Fn_ABM_RootStructureTableOperations("TabPopupMenuSelect", "\(000059\/MRO\-06\-A\)", "", "", "", "", "Split Panel")
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
''Examples		     	:	Call  Fn_ABM_RootStructureTableOperations(sAction, "", "", "000554/A;1-TopPart (View):000555/A;1-Ch1 (View)", "", "", "")
''Examples		     	:	Call  Fn_ABM_RootStructureTableOperations("Select", "", "", "000554/A;1-TopPart (View):000555/A;1-Ch1 (View)", "", "", "")
''Examples		     	:	Call  Fn_ABM_RootStructureTableOperations("Exist", "", "", "000554/A;1-TopPart (View):000555/A;1-Ch1 (View)", "", "", "")
''Examples		     	:	Call  Fn_ABM_RootStructureTableOperations("Expand", "", "", "000554/A;1-TopPart (View):000555/A;1-Ch1 (View)", "", "", "")
''Examples		     	:	Call  Fn_ABM_RootStructureTableOperations("ExpandBelow", "", "", "000554/A;1-TopPart (View)", sColName, sValue, "")
''Examples		     	:	Call  Fn_ABM_RootStructureTableOperations("PopupSelect", "", "", "000554/A;1-TopPart (View)", sColName, sValue, "")
''Examples		     	:	Call  Fn_ABM_RootStructureTableOperations("CellEdit", "", "", "000554/A;1-TopPart (View):000555/A;1-Ch1 (View)", "", "", "")
''Examples		     	:	Call  Fn_ABM_RootStructureTableOperations("Select_OnBelowTable", "", "", "000554/A;1-TopPart (View):000555/A;1-Ch1 (View)", "", "", "")
''Examples		     	:	Call  Fn_ABM_RootStructureTableOperations("Exist_OnBelowTable", "", "", "000554/A;1-TopPart (View):000555/A;1-Ch1 (View)", "", "", "")
''Examples		     	:	Call  Fn_ABM_RootStructureTableOperations("Expand_OnBelowTable", "", "", "000554/A;1-TopPart (View):000555/A;1-Ch1 (View)", "", "", "")
''Examples		     	:	Call  Fn_ABM_RootStructureTableOperations("ExpandBelow_OnBelowTable", "", "", "000554/A;1-TopPart (View)", sColName, sValue, "")
''Examples		     	:	Call  Fn_ABM_RootStructureTableOperations("PopupSelect_OnBelowTable", "", "", "000554/A;1-TopPart (View)", sColName, sValue, "")
''Examples		     	:	Call  Fn_ABM_RootStructureTableOperations("CellEdit_OnBelowTable", "", "", "000554/A;1-TopPart (View):000555/A;1-Ch1 (View)", "", "", "")
''Examples		     	:	Call  Fn_ABM_RootStructureTableOperations("VerifyForegroundColour", "", "", "000554/A;1-TopPart (View)", "", "GREEN", "")
''Examples		     	:	Call  Fn_ABM_RootStructureTableOperations("VerifyBackgroundColour", "", "", "000554/A;1-TopPart (View):000555/A;1-Ch1 (View)", "", "YELLOW", "")
''Examples		     	:	Call  Fn_ABM_RootStructureTableOperations("VerifyForegroundColour_OnBelowTable","", "", "000554/A;1-TopPart (View)", "", "GREEN", "")
''Examples		     	:	Call  Fn_ABM_RootStructureTableOperations("VerifyBackgroundColour_OnBelowTable", "", "", "000554/A;1-TopPart (View):000555/A;1-Ch1 (View)", "", "YELLOW", "")

'History:
'	Developer Name			Date				Rev. No.	Reviewer			Changes Done
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Koustubh Watwe		5-Sept-2011			1.0								Created
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Vrushali Wani		13-June-2012		1.0								Modified 'ExpandBelow' dialog hierarchy
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Vrushali Wani		21-June-2012		1.0			Koustubh			Modified case 'CellVerify' 
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Vrushali Wani		26-June-2012		1.0			Koustubh			added  case 'TabPopupMenuSelect' 
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Koustubh Watwe	     26-June-2012	     1.1		Koustubh			added  cases to perform operations on splitted tables.
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Koustubh Watwe	     20-Aug- 2012	     1.1		Koustubh			added  cases VerifyForegroundColour, VerifyBackgroundColour.
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Vrushali Wani	     22-Aug- 2012	     1.1		Koustubh			added  cases VerifyBackgroundColour_OnBelowTable, VerifyForegroundColour_OnBelowTable.
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Sandeep Navghane	    27-Mar- 2013	     1.2		Anumol			Modified case : "TabPopupMenuSelect" Added new Hierarchy of Popup menu parent object.
'																															Old: JavaWindow("AsBuiltManager").JavaMenu(...)
'																															New: JavaWindow("AsBuiltManager").JavaApplet("JApplet").JavaMenu(...)
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public function Fn_ABM_RootStructureTableOperations(sAction, sRootStructureHeader, sRootStructure, sNodeName, sColName, sValue, sPopupMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_ABM_RootStructureTableOperations"
	Dim iRowIndex, aMenu, objTable, objParentObject, strMenu, iColIndex,objExpandbelow,bReturn
	Dim sValuearr

	Fn_ABM_RootStructureTableOperations = False
	Set objParentObject = Fn_SISW_ABM_GetObject("JApplet")
	Set objTable = objParentObject.JavaTable("RootStructuresTable")

	objParentObject.JavaObject("RootStructureTablePanel").Click 0,0, "LEFT" 

	' selcting Root Structure Header.
	If sRootStructureHeader <> "" then
	    If inStr(1,sRootStructureHeader,"\(") Then
            'Do nothing
		Else
	        If inStr(1,sRootStructureHeader,"(") Then
                sRootStructureHeader=replace(sRootStructureHeader,"(","\(")
				If inStr(1,sRootStructureHeader,"\)") Then
					'Do nothing
				Else
	                If inStr(1,sRootStructureHeader,")") Then
							sRootStructureHeader=replace(sRootStructureHeader,")","\)")
					End If
				End If
			End if
		End If

		objParentObject.JavaStaticText("RootStructureHeader").SetTOProperty "label", sRootStructureHeader
		wait 1
        If objParentObject.JavaStaticText("RootStructureHeader").Exist(2) Then
			objParentObject.JavaStaticText("RootStructureHeader").Click 1,1,"LEFT"
			wait 1
		End if
	End IF

	' selecting Root Structure from List.
	If sRootStructure <> "" Then
		Call Fn_List_Select("Fn_ABM_RootStructureTableOperations",objParentObject,"RootStructures",sRootStructure)
	End If

	If Instr(sAction,"_OnBelowTable") > 0 Then
		Set objTable = objParentObject.JavaTable("RootStructuresBottomTable")
	End If

	Select Case sAction
	    '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "TabSelect"
				If objParentObject.JavaStaticText("RootStructureHeader").Exist(2) Then
					Fn_ABM_RootStructureTableOperations = True
				End if
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "TabPopupMenuSelect"
				objParentObject.JavaStaticText("RootStructureHeader").SetTOProperty "label", sRootStructureHeader
				objParentObject.JavaStaticText("RootStructureHeader").Click 1,1,"RIGHT"
				wait 2
				Fn_ABM_RootStructureTableOperations = Fn_UI_JavaMenu_Select("",JavaWindow("AsBuiltManager"),sPopupMenu)
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Exist", "Exists", "Exist_OnBelowTable","Exists_OnBelowTable"
				  iRowIndex = Fn_ABM_RootStructureTableRowIndex(objTable, sNodeName)
				  If iRowIndex <> -1 Then
					  Fn_ABM_RootStructureTableOperations = True
					  Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_ABM_RootStructureTableRowIndex ] Successfully verified existence of Node [ " & sNodeName & " ].")
				  Else
    					  Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_ABM_RootStructureTableRowIndex ] Node [ " & sNodeName & " ] is not exists.")
				  End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Select","Select_OnBelowTable"
				  iRowIndex = Fn_ABM_RootStructureTableRowIndex(objTable, sNodeName)
				  If iRowIndex <> -1 Then
					  wait 5
					  Call Fn_UI_JavaTable_SelectRow("Fn_ABM_RootStructureTableRowIndex",objParentObject ,"RootStructuresTable",iRowIndex)
					  Fn_ABM_RootStructureTableOperations = True
				  End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "MultiSelect","MultiSelect_OnBelowTable"
			' for future use
			' Not implemented yet
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Expand","Expand_OnBelowTable"
				  iRowIndex = Fn_ABM_RootStructureTableRowIndex(objTable, sNodeName)
				  If iRowIndex <> -1 Then
					  Call Fn_UI_JavaTable_SelectRow("Fn_ABM_RootStructureTableRowIndex",objParentObject ,"RootStructuresTable",iRowIndex)
					  wait 2
					  Fn_ABM_RootStructureTableOperations = Fn_MenuOperation("WinMenuSelect", "View:Expand")
				  Else 
				      Fn_ABM_RootStructureTableOperations = False
				  End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "ExpandBelow","ExpandBelow_OnBelowTable"
				  iRowIndex = Fn_ABM_RootStructureTableRowIndex(objTable, sNodeName)
				  If iRowIndex <> -1 Then
					  Call Fn_UI_JavaTable_SelectRow("Fn_ABM_RootStructureTableRowIndex",objParentObject ,"RootStructuresTable",iRowIndex)
					  wait 2
					  Fn_ABM_RootStructureTableOperations = Fn_MenuOperation("WinMenuSelect", "View:Expand Below")   
					  Set objExpandbelow =  JavaWindow("AsBuiltManager").JavaApplet("JApplet").JavaDialog("ExpandBelow")
					  If Fn_UI_ObjectExist("Fn_ABM_RootStructureTableRowIndex",objExpandbelow)  then
						  'Click Yes Button 
						  Call Fn_Button_Click("Fn_ABM_RootStructureTableRowIndex", objExpandbelow, "Yes")
					  End If
					  Set objExpandbelow = nothing
                  Else 
				     Fn_ABM_RootStructureTableOperations = False
				  End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "PopupSelect","PopupSelect_OnBelowTable"
				   iRowIndex = Fn_ABM_RootStructureTableRowIndex(objTable, sNodeName)
				   
				   If sColName = "" Then
					   sColName = "Neutral Structure"
					   iColIndex = 0
				   Else
						iColIndex = Fn_ABM_RootStructureTableColumnOperations("GetIndex", sColName)						
				   End If
				   
				  If iRowIndex <> -1 AND iColIndex <> -1 Then
				         wait 3
						'Call Fn_UI_JavaTable_SelectRow("Fn_ABM_RootStructureTableRowIndex",objParentObject ,"RootStructuresTable",iRowIndex)
						aMenu = split(sPopupMenu,":",-1,1)
						objTable.ClickCell iRowIndex, iColIndex ,"RIGHT"
						wait 3
						Select Case Ubound(aMenu)
							Case 0
								strMenu = JavaWindow("AsBuiltManager").WinMenu("ContextMenu").BuildMenuPath(aMenu(0))
								JavaWindow("AsBuiltManager").WinMenu("ContextMenu").Select strMenu
							Case 1
								strMenu = JavaWindow("AsBuiltManager").WinMenu("ContextMenu").BuildMenuPath(aMenu(0),aMenu(1))
								JavaWindow("AsBuiltManager").WinMenu("ContextMenu").Select strMenu
						End Select
						Fn_ABM_RootStructureTableOperations = True
				  End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "CellEdit","CellEdit_OnBelowTable"
				   iRowIndex = Fn_ABM_RootStructureTableRowIndex(objTable, sNodeName)
				   
				   If sColName = "" Then
					   sColName = "Neutral Structure"
					   iColIndex = 0
				   Else
						iColIndex = Fn_ABM_RootStructureTableColumnOperations("GetIndex", sColName)						
				   End If
				   
				  If iRowIndex <> -1 AND iColIndex <> -1 Then
					Call Fn_UI_JavaTable_SelectRow("Fn_ABM_RootStructureTableRowIndex",objParentObject ,"RootStructuresTable",iRowIndex)
						objTable.ClickCell iRowIndex, iColIndex,"LEFT"
						If JavaWindow("AsBuiltManager").JavaApplet("JApplet").JavaEdit("RootStructTblCellEdit").Exist(2) Then
							'Workaround fpr PR#6813387
'							JavaWindow("AsBuiltManager").JavaApplet("JApplet").JavaEdit("RootStructTblCellEdit").Set ""
							
 							call Fn_Edit_Box("Fn_ABM_RootStructureTableOperations",JavaWindow("AsBuiltManager").JavaApplet("JApplet"),"RootStructTblCellEdit",sValue)

'							JavaWindow("AsBuiltManager").JavaApplet("JApplet").JavaEdit("RootStructTblCellEdit").Set sValue                          
							JavaWindow("AsBuiltManager").JavaApplet("JApplet").JavaEdit("RootStructTblCellEdit").Activate
						Else
							Call Fn_UI_JavaTable_SetCellData("Fn_SrvMgr_RootStructureTableOperations",objParentObject ,"RootStructuresTable",iRowIndex,iColIndex,sValue)
						End If
					Call Fn_UI_JavaTable_SetCellData("Fn_ABM_RootStructureTableRowIndex",objParentObject ,"RootStructuresTable",iRowIndex,iColIndex,sValue)
					Fn_ABM_RootStructureTableOperations = True
				  End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "CellVerify","CellVerify_OnBelowTable"
				'Workaround fpr PR#6813387
				If Trim(LCase(sValue)) = "y" Then
					sValue = "True"
				ElseIf Trim(sValue) = "" Then
					sValue = "False"
				End If

				iRowIndex = Fn_ABM_RootStructureTableRowIndex(objTable, sNodeName)
				
				If sColName = "" Then
					sColName = "Neutral Structure"
					iColIndex = 0
				Else
					iColIndex = Fn_ABM_RootStructureTableColumnOperations("GetIndex", sColName)						
				End If
				   
				If iRowIndex <> -1 AND iColIndex <> -1 Then
					Call Fn_UI_JavaTable_SelectRow("Fn_ABM_RootStructureTableRowIndex",objParentObject ,"RootStructuresTable",iRowIndex)
					sActValue =  Fn_UI_JavaTable_GetCellData("Fn_ABM_RootStructureTableRowIndex", objParentObject, "RootStructuresTable",iRowIndex,iColIndex)

					If 	sColName = "Installation Time" OR	sColName = "Manufacturing Date" Then
						sValuearr = split(sValue," ")
						' checking whether date is present in sActValue
							If instr(sActValue,sValuearr(0))  Then
								 Fn_ABM_RootStructureTableOperations = True
							Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_ABM_RootStructureTableOperations ] Date [ " & sValuearr(0) & " ] is not present in actual value [ " & sActValue & " ] of column [ " & sColName & " ].")
							End If
					Else
							If Trim(sActValue) = Trim(sValue) Then
								Fn_ABM_RootStructureTableOperations = True
							Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_ABM_RootStructureTableOperations ] Value [ " & sValue & " ] is not present in column [ " & sColName & " ].")
							End If
					End If
				End If
          	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			Case "VerifyForegroundColour", "VerifyBackgroundColour", "VerifyForegroundColour_OnBelowTable", "VerifyBackgroundColour_OnBelowTable"
				iRowIndex = Fn_ABM_RootStructureTableRowIndex(objTable, sNodeName)
				If cint(iRowIndex) = -1 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [Fn_ABM_RootStructureTableOperations] Couldnt find  Root Structure Table Node [" + sNodeName + "]")
					Exit function
				End If
				Set  objNodeForRow =  objTable.Object.getNodeForRow(cint(iRowIndex))
				' if background colour
				If sAction = "VerifyBackgroundColour" OR sAction = "VerifyBackgroundColour_OnBelowTable" Then
					sColour = objTable.Object.getBackground(objNodeForRow,False).toString()
				Else
				' if foreground colour
					sColour = objTable.Object.getForeground(objNodeForRow,False).toString()
				End If
		
				sColour =  mid(sColour ,instr(sColour ,"[")  ,instr(sColour ,"]") )
				' comparing colour codes RGB
				Select Case lcase(sColour)
					Case "[r=0,g=0,b=0]"
						sColour = "BLACK"
					Case "[r=159,g=255,b=159]"
						sColour = "GREEN"
					Case "[r=255,g=255,b=128]", "[r=255,g=200,b=0]", "[r=255,g=255,b=0]"
						sColour = "YELLOW"
					Case "[r=0,g=255,b=255]"
						sColour = "CYAN"
					Case "[r=0,g=0,b=255]"
						sColour = "BLUE"
					Case "[r=255,g=121,b=121]", "[r=255,g=0,b=0]"
						sColour = "RED"
					Case Else
						sColour = ""
				End Select
				If sValue = sColour Then
					Fn_ABM_RootStructureTableOperations = True
				Else
					Fn_ABM_RootStructureTableOperations = False
				ENd If
			'This case is used to get the Cell Value----------------------------------------------
			'[TC1122-20151116d-29_12_2015-VivekA-Maintenance] - Added from TC1015
			Case "GetCellData"

				iRowIndex = Fn_ABM_RootStructureTableRowIndex(objTable, sNodeName)
				
				If sColName = "" Then
					sColName = "Neutral Structure"
					iColIndex = 0
				Else
					iColIndex = Fn_ABM_RootStructureTableColumnOperations("GetIndex", sColName)						
				End If
				   
				If iRowIndex <> -1 AND iColIndex <> -1 Then
					Call Fn_UI_JavaTable_SelectRow("Fn_ABM_RootStructureTableOperations",objParentObject ,"RootStructuresTable",iRowIndex)
					Fn_ABM_RootStructureTableOperations =  Fn_UI_JavaTable_GetCellData("Fn_ABM_RootStructureTableOperations", objParentObject, "RootStructuresTable",iRowIndex,iColIndex)
				End IF
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_ABM_RootStructureTableOperations ] Invalid case [ " & sAction & " ].")
			Fn_ABM_RootStructureTableOperations = False

	End Select

	If Fn_ABM_RootStructureTableOperations <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_ABM_RootStructureTableOperations ] executed successfully with case [ " & sAction & " ].")
	End If
	Set objParentObject = Nothing
	Set objTable = Nothing
End Function
''****************************************    Function to perform operations on Lots in MRO ***************************************
'
''Function Name		 	:		Fn_ABM_LotOperations
'
''Description		    :  	    Function to perform operations on Lots in MRO 
'
''Parameters		    :	 	1. sAction : Action need to perform
'					   		    2. sOpenDialogBy : to open Lot dialog ( NewLot_RMB / NewLot_Menu )
'					   		    3. sRootStructureNode : Root Structure Node Path
'								4. dicLotOperations : Dictionary object for Lot operations
'								
''Return Value		    :  		True \ False
'
''Pre-requisite		    :		MRO perspective should be selected

' 							   sRootStructureNode = "000087/A;1-Top (View)"
'							   dicLotOperations("LotNumber") = "0001"
'							   dicLotOperations("ManufacturersID") = "1001"
'							   dicLotOperations("LotSize") = "1"

''Examples		     	:	Call Fn_ABM_LotOperations("NewLot", "NewLot_RMB", "000087/A;1-Top (View)", dicLotOperations)
''Examples		     	:	Call Fn_ABM_LotOperations("NewLot", "NewLot_Menu", "", dicLotOperations)
''Examples		     	:	Call Fn_ABM_LotOperations("NewLot", "", "", dicLotOperations)

'History:
'						Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'						   Koustubh Watwe		6-Sept-2011			1.0					 Created
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_ABM_LotOperations(sAction, sOpenDialogBy, sRootStructureNode, dicLotOperations)
	GBL_FAILED_FUNCTION_NAME="Fn_ABM_LotOperations"
	Dim objLotDialog, bReturn
	Fn_ABM_LotOperations = False
	Set objLotDialog = JavaWindow("AsBuiltManager").JavaWindow("Lot")
	If Fn_UI_ObjectExist("Fn_ABM_LotOperations",objLotDialog) = False then
		Select Case sOpenDialogBy
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
			Case "NewLot_RMB"
					If sRootStructureNode <> "" then
						bReturn = Fn_ABM_RootStructureTableOperations("PopupSelect", "", "", sRootStructureNode, "", "", "New:Lot...")
						If bReturn = False Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_ABM_LotOperations ] Failed to perform [ RMB : New:Lot... ] on Root Strcture Node [ " & dicLotOperations("sRootStructureNode") & " ].")
							Set objLotDialog = Nothing
							Exit function
						End If
					End If
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
			Case "NewLot_Menu", ""
					If dicLotOperations("sRootStructureNode") <> "" then
						bReturn = Fn_ABM_RootStructureTableOperations("Select", "", "", sRootStructureNode, "", "", "")
						If bReturn = False Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_ABM_LotOperations ] Failed to select Root Strcture Node [ " & dicLotOperations("sRootStructureNode") & " ].")
							Set objLotDialog = Nothing
							Exit function
						End If
					End If
					Call Fn_MenuOperation("Select","File:New:Lot...")
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
			End Select
			If Fn_UI_ObjectExist("Fn_ABM_LotOperations",objLotDialog) = False then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_ABM_LotOperations ] Failed to open New Lot Window.")
				Set objLotDialog = Nothing
				Exit function
			End IF
	End IF
	Select Case sAction
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "NewLot"
			' set Lot number
			Call Fn_Edit_Box("Fn_ABM_LotOperations", objLotDialog, "LotNumber",dicLotOperations("LotNumber"))

			' set Manufacturer's ID
			'JavaWindow("AsBuiltManager").JavaWindow("Lot").JavaEdit("ManufacturersID").Type dicLotOperations("ManufacturersID")
			JavaWindow("AsBuiltManager").JavaWindow("Lot").JavaEdit("ManufacturersID").Set dicLotOperations("ManufacturersID")
			'Work-around for enabling OK button [Only 'Type' method on edit fails to print '0' and 'Set' method doesn't invoke event]
			JavaWindow("AsBuiltManager").JavaWindow("Lot").JavaEdit("ManufacturersID").Type "a"
			Call Fn_KeyBoardOperation("SendKeys", "{BKSP}")
			' set Lot Size
			Call Fn_Edit_Box("Fn_ABM_LotOperations", objLotDialog, "LotSize",dicLotOperations("LotSize"))

			'clicking on OK button
			Call Fn_Button_Click("Fn_ABM_LotOperations", objLotDialog, "OK")
			Fn_ABM_LotOperations = True
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_ABM_LotOperations ] Invalid case [ " & sAction & " ].")
			Fn_ABM_LotOperations = False
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	End Select
	If Fn_ABM_LotOperations = True Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_ABM_LotOperations ] executed successfully with case [ " & sAction & " ].")
	End If
	Set objLotDialog = Nothing
End Function

'*********************************************************		Function to action perform on NavTree of  As Built Manager ***********************************************************************
'Function Name		:				Fn_ASB_NavTree_NodeOperation

'Description			 :		 		 Actions performed in this function are:
'																	1. Node Select
'																	2. Node multi-select
'																	3. Node Expand
'																	4. Node Collapse
'																	5. Node Popup menu select
'																	6. Node double-click
'																	7. Node Deselect
'																	8. Node Exist
'																	9. Node SelectRange

'Parameters			   :	 			1. StrAction: Action to be performed
'													2. StrNodeName: Fully qulified tree Path (delimiter as ':') [multiple node are separated by "," ] 
'												   3. StrMenu: Context menu to be selected

'Return Value		   : 				TRUE \ FALSE

'Pre-requisite			:		 		As Built Manager  module window should be displayed

'Examples				:				  Fn_ASB_NavTree_NodeOperation("PopupMenuSelect","Home:Newstuff","Copy Ctrl+C")
'										  EXAMPLE for Case "Select" : Call Fn_ASB_NavTree_NodeOperation( "Select" ,  "Home:Newstuff:000032-CarModel_VI_LS1:000032 @2" , "" ) 
'										  Call Fn_ASB_NavTree_NodeOperation( "Select" ,  "Home:Newstuff:000032-CarModel_VI_LS1:000032" , "" ) 
'History					 :		
'	Developer Name				Date						Rev. No.			Changes Done						Reviewer
'	------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Rupali Palhade				16/09/2011			          1.0					Created
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Koustubh Watwe				06/06/2012			          1.1				 Modifeid code to generate node path
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Vrushali Wani				   07/06/2012			          1.1				Modifeid case PopupSelect
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_ASB_NavTree_NodeOperation(StrAction,StrNodeName,StrMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_ASB_NavTree_NodeOperation"
	Dim NodeLists, intNodeCount, intCount, StrExist, aMenuList, sTreeItem,sCmpItm
	Dim objJavaTreeNav,ArrNodeName
	Dim ArrStrcomp, sArrStr1,sArrStr2, iCounter
	Dim iRows, colonCnt
	Dim iItemCount, aNodePath,  iInstance, instCount, aNodes
	Dim sPath, sEle ,iCnt, bFound
	Dim iLen,iIndex,iTotal,iCount,sReturn,iReturn,arr
	Fn_ASB_NavTree_NodeOperation = False
	Set objJavaWindowMyTc = JavaWindow("AsBuiltManager")
	Set objJavaTreeNav = JavaWindow("AsBuiltManager").JavaTree("NavTree")

	Select Case StrAction
		Case "Multiselect", "SelectRange"
			' do nothing
		Case Else
			If StrNodeName <> "" Then
				sPath =  Fn_UI_JavaTreeGetItemPathExt("Fn_ASB_NavTree_NodeOperation", objJavaTreeNav, StrNodeName, "", "")
				If sPath = False Then
					Exit function
				End If
			End If
	End Select

	Select Case StrAction

		'----------------------------------------------------------------------- For selecting single node -------------------------------------------------------------------------
		Case "Select"
			objJavaTreeNav.select sPath
			 Fn_ASB_NavTree_NodeOperation = TRUE
		'----------------------------------------------------------------------- For selecting single node -------------------------------------------------------------------------

		Case "Deselect"
			objJavaTreeNav.Deselect sPath
			 Fn_ASB_NavTree_NodeOperation = TRUE
		'----------------------------------------------------------------------- For expanding a particular  node-------------------------------------------------------------------------
		Case "Expand"
		     objJavaTreeNav.Expand sPath
			 Fn_ASB_NavTree_NodeOperation = TRUE
		'----------------------------------------------------------------------- For selecting multiple node at a time -------------------------------------------------------------------------
		Case "Multiselect"
			' needs modifications
			Set objJavaTreeNav = JavaWindow("AsBuiltManager").JavaTree("NavTree")
			 Call Fn_UI_JavaTree_ExtendSelect("Fn_ASB_NavTree_NodeOperation",objJavaWindowMyTc,"NavTree", StrNodeName)
			 Fn_ASB_NavTree_NodeOperation = TRUE

		'----------------------------------------------------------------------- For collapssing a particular  node-------------------------------------------------------------------------
		Case "Collapse"
			objJavaTreeNav.Collapse sPath
		    Fn_ASB_NavTree_NodeOperation = True

		'----------------------------------------------------------------------- For selecting popup menu of  a particular  node-------------------------------------------------------------------------
		Case "PopupMenuSelect"
			Select Case StrMenu
				Case "Show in As-Built Manager"
						'Select node
						Call Fn_JavaTree_Select("Fn_ASB_NavTree_NodeOperation",objJavaWindowMyTc,"NavTree",sPath)
			
						'Open context menu
						Call Fn_UI_JavaTree_OpenContextMenu("Fn_ASB_NavTree_NodeOperation",objJavaWindowMyTc,"NavTree",sPath)

						JavaWindow("AsBuiltManager").JavaMenu("Menu").SetTOProperty "label", "Show in As-Built Manager"
						JavaWindow("AsBuiltManager").JavaMenu("Menu").Select

						Fn_ASB_NavTree_NodeOperation = True
				Case Else
						'Build the Popup menu to be selected
						aMenuList = split(StrMenu, ":",-1,1)
						intCount = Ubound(aMenuList)
			
						'Select node
						Call Fn_JavaTree_Select("Fn_ASB_NavTree_NodeOperation",objJavaWindowMyTc,"NavTree",sPath)
			
						'Open context menu
						Call Fn_UI_JavaTree_OpenContextMenu("Fn_ASB_NavTree_NodeOperation",objJavaWindowMyTc,"NavTree",sPath)
						
						'Select Menu action
						Select Case intCount
							Case 0
								 StrMenu = objJavaWindowMyTc.WinMenu("ContextMenu").BuildMenuPath(aMenuList(0))
			
							Case 1
								StrMenu = objJavaWindowMyTc.WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1))
			
							Case 2
								StrMenu = objJavaWindowMyTc.WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1),aMenuList(2))
			
							Case Else
								Fn_ASB_NavTree_NodeOperation = FALSE
								Exit Function
						End Select
			
						objJavaWindowMyTc.WinMenu("ContextMenu").Select StrMenu
						Fn_ASB_NavTree_NodeOperation = True
			End Select
		'----------------------------------------------------------------------- For doble clicking on a particular  node-------------------------------------------------------------------------
		Case "DoubleClick"
				objJavaTreeNav.Activate sPath
				Fn_ASB_NavTree_NodeOperation = True
		'----------------------------------------------------------------------- For doble clicking on a particular  node-------------------------------------------------------------------------
		Case "Exist"
				Fn_ASB_NavTree_NodeOperation = True
	     '----------------------------------------------------------------------- For  Select Range of  Nav tree -------------------------------------------------------------------------
		Case "SelectRange"
			' needs modifications
			ReDim ArrNodeName(2)
			ArrNodeName = Split(StrNodeName,"|")
			JavaWindow("AsBuiltManager").JavaTree("NavTree").SelectRange ArrNodeName(0),ArrNodeName(1)
			If err.number < 0 Then
				Fn_ASB_NavTree_NodeOperation = false
			else
				Fn_ASB_NavTree_NodeOperation = True
			End If
							
		'****************************************************************************************	
		Case Else
				Fn_ASB_NavTree_NodeOperation = FALSE
	End Select

	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), StrAction &" Sucessfully completed on Node [" + StrNodeName + "] of JavaTree of function Fn_ASB_NavTree_NodeOperation")
	Set objJavaWindowMyTc = nothing
	Set objJavaTreeNav = nothing

End Function
''****************************************    Function to perform Serial number genrator ***************************************
'
''Function Name		 	:		Fn_ABM_SerialNoGenerator
'
''Description		    :  	    Function to perform  serial number generator

''Parameters		    :	 	1. sAction : Action need to perform
'					   		    2. sOpenDialogBy : to open Lot dialog ( SerialNumberGenerator_RMB / SerialNumber_Menu )
'					   		   3. sRootStructureNode : Root Structure Node Path
'							  4. sSeriesPrefix : Series prefix value 
'                             5. sSeriesSuffix : Series Suffix value 
'                             6. sSeriesCurrent : Series Current  value 
'                             7. sIncrement : Series increment value 
'                             8. sCharacter : Series Character value 
'                             9. sMaximum : Series Maximum value 
'								
''Return Value		    :  		True \ False
'
''Pre-requisite		    :		As Build Manager perspective should be selected

''Examples		     	:	Fn_ABM_SerialNoGenerator("NewSerial", "NewLot_Menu", "000087/A;1-Top (View)", "MRO-" ,"" , "04" , "10","1" , "2")

'History:
'						Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'						Rupali Palhade		19-Sept-2011			1.0					 Created
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_ABM_SerialNoGenerator(sAction, sOpenDialogBy, sRootStructureNode, sSeriesPrefix ,sSeriesSuffix , sSeriesCurrent ,sMaximum , sIncrement , sCharacter)
	GBL_FAILED_FUNCTION_NAME="Fn_ABM_SerialNoGenerator"
	Dim objSeriesDialog, bReturn
	Fn_ABM_SerialNoGenerator = False
	Set objSeriesDialog = JavaWindow("AsBuiltManager").JavaWindow("NewSerialNoGenerator")
	If Fn_UI_ObjectExist("Fn_ABM_SerialNoGenerator",objSeriesDialog) = False then

		Select Case sOpenDialogBy
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
			Case "NewLot_RMB"
					If sRootStructureNode <> "" then
						wait 2
						bReturn = Fn_ABM_RootStructureTableOperations("PopupSelect", "", "", sRootStructureNode, "", "", "New:Serial Number Generator...")
						If bReturn = False Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_ABM_SerialNoGenerator ] Failed to perform [ RMB : New:Serial Number Generator... ] on Root Strcture Node [ " & sRootStructureNode & " ].")
							Set objSeriesDialog = Nothing
							Exit function
						End If
					End If
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
			Case "NewLot_Menu", ""
					If sRootStructureNode <> "" then
						bReturn = Fn_ABM_RootStructureTableOperations("Select", "", "", sRootStructureNode, "", "", "")
						If bReturn = False Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_ABM_SerialNoGenerator ] Failed to select Root Strcture Node [ " & sRootStructureNode & " ].")
							Set objSeriesDialog = Nothing
							Exit function
						End If
					End If
					Call Fn_MenuOperation("Select","File:New:Serial Number Generator...")
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
			End Select

			If Fn_UI_ObjectExist("Fn_ABM_SerialNoGenerator",objSeriesDialog) = False then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_ABM_SerialNoGenerator ] Failed to open New Serial Generator Window.")
				Set objSeriesDialog = Nothing
				Exit function
			End If

	End IF

	Select Case sAction
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "NewSerial"
			' set  Series Prefix
			Call Fn_Edit_Box("Fn_ABM_SerialNoGenerator", objSeriesDialog, "SeriesPrefix",sSeriesPrefix )

			' set  Series Suffix
			Call Fn_Edit_Box("Fn_ABM_SerialNoGenerator", objSeriesDialog, "SeriesSuffix",sSeriesSuffix )

			' set  Series Current
			Call Fn_UI_EditBox_Type("Fn_ABM_SerialNoGenerator", objSeriesDialog, "SeriesCurrent",sSeriesCurrent )

			' set  Series Increment
			Call Fn_UI_EditBox_Type("Fn_ABM_SerialNoGenerator", objSeriesDialog, "Increment",sIncrement )

			' set  Series Maximum
			Call Fn_UI_EditBox_Type("Fn_ABM_SerialNoGenerator", objSeriesDialog, "SeriesMaximum",sMaximum )

			' set  Series Character
			Call Fn_UI_EditBox_Type("Fn_ABM_SerialNoGenerator", objSeriesDialog, "SeriesCharacters",sCharacter)

			'clicking on OK button
			Call Fn_Button_Click("Fn_ABM_SerialNoGenerator", objSeriesDialog, "OK")
			Fn_ABM_SerialNoGenerator = True
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_ABM_SerialNoGenerator ] Invalid case [ " & sAction & " ].")
			Fn_ABM_SerialNoGenerator = False
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	End Select

	If Fn_ABM_SerialNoGenerator = True Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_ABM_SerialNoGenerator ] executed successfully with case [ " & sAction & " ].")
	End If

	Set objSeriesDialog = Nothing
End Function

''****************************************    Function to perform Generate As Built Structure ***************************************

''Function Name		 	:		Fn_ABM_GenerateAsBuiltStructure
'
''Description		    :  	    Function to perform  Generate As built  structure

''Parameters		    :	 	1. sAction : Action need to perform
'					   		   2. sRootStructureNode : Root Structure Node Path
'					   		   3. sSerialNo : Serial Number
'                             4. sLot : Lot value need to set 
'                             5. SManufactueID :  Manufacure' ID 
'                             6. sStuctContextName : Structure Context Name
'                             7. sManufDate : Manufacture Date
'                             8. sInstallationTime : Installation time 
'                            9. sNoofLevel : Number of level/
'                           10.sRootStructureHeader : Header of the table need to click
'								
''Return Value		    :  		True \ False
'
''Pre-requisite		    :		As Build Manager perspective should be selected

''Examples		     	:	Fn_ABM_GenerateAsBuiltStructure("GenerateAsBuiltStructure", "000087/A;1-Top (View)"," " ,"L1" , "000020" , "000020","22-Sep-2011" , "00:00:00","3")

''History:
'						Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'						Rupali Palhade		19-Sept-2011			1.0					 Created
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'						Nikhil D			04-Jun-2014				1.0					 modified to set label property of Property_Name object
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function Fn_ABM_GenerateAsBuiltStructure(sAction,sRootStructureNode,sSerialNo,sLot,sManufactueID,sStuctContextName,sManufDate,sInstallationTime,sNoofLevel)
	GBL_FAILED_FUNCTION_NAME="Fn_ABM_GenerateAsBuiltStructure"
	Dim objGenerateABSDialog, bReturn , arrDateTime
	Fn_ABM_GenerateAsBuiltStructure = False
	Set objGenerateABSDialog = JavaWindow("AsBuiltManager").JavaWindow("GenerateAsBuiltStructure")
	If Fn_UI_ObjectExist("Fn_ABM_GenerateAsBuiltStructure",objGenerateABSDialog) = False then

			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
					If sRootStructureNode <> "" then
						'wait 2
						bReturn = Fn_ABM_RootStructureTableOperations("PopupSelect","", "", sRootStructureNode, "", "", "Generate As-Built Structure")
						If bReturn = False Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_ABM_GenerateAsBuiltStructure ] Failed to perform [ RMB : New:Generate As-Built Structure ] on Root Strcture Node [ " & sRootStructureNode & " ].")
							Set objGenerateABSDialog = Nothing
							Exit function
						End If
					End If
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
			
			'Added to handle 'Show Unconfigured is set' dialog popping up since Tc1122_1119 build
			If JavaWindow("DefaultWindow").JavaWindow("ShowUnconfigured_is_set").Exist(3) Then
				JavaWindow("DefaultWindow").JavaWindow("ShowUnconfigured_is_set").JavaButton("OK").Click micLeftBtn
			End If

			If Fn_UI_ObjectExist("Fn_ABM_GenerateAsBuiltStructure",objGenerateABSDialog) = False then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_ABM_GenerateAsBuiltStructure ] Failed to open Generate As Built Structure Window.")
				Set objGenerateABSDialog = Nothing
				Exit function
			End If

	End IF

	Select Case sAction
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "GenerateAsBuiltStructure"
			' set  Serial Number
			If sSerialNo <> "" Then
				objGenerateABSDialog.JavaStaticText("Property_Name").SetTOProperty "label", "Serial Number :"
				wait 1
				Call Fn_UI_EditBox_Type("Fn_ABM_GenerateAsBuiltStructure", objGenerateABSDialog, "SerialNumber",sSerialNo )
                wait 1
			End If
			
			' select Checkbox  use serial number generator
			Call Fn_CheckBox_Select("Fn_ABM_GenerateAsBuiltStructure", objGenerateABSDialog, "UseSerialNoGenerators")


			' set  Lot Value
			If sLot <> "" Then
				objGenerateABSDialog.JavaStaticText("Property_Name").SetTOProperty "label", "Lot :"
				wait 1
				Call Fn_List_Select("Fn_ABM_GenerateAsBuiltStructure", objGenerateABSDialog,"Lot",sLot)
			End If

			' set  Manufacture's ID
			If sManufactueID <> "" Then
				objGenerateABSDialog.JavaStaticText("Property_Name").SetTOProperty "label", "Manufacturer's ID :"
                wait 1
			    Call Fn_UI_EditBox_Type("Fn_ABM_GenerateAsBuiltStructure", objGenerateABSDialog, "ManufacturersID",sManufactueID )
			End If

		    ' set  Structure Context Name 
			If sStuctContextName <> "" Then
			    Call Fn_UI_EditBox_Type("Fn_ABM_GenerateAsBuiltStructure", objGenerateABSDialog, "StructContext Name",sStuctContextName )
			End If

			' set Manufacturing Date 
			If sManufDate <> "" Then
			    arrDateTime = Split(sManufDate," ")
            	objGenerateABSDialog.JavaEdit("ManufacturingDate").Set arrDateTime(0)
				wait 1
				Set WshShell = CreateObject("WScript.Shell")
				WshShell.SendKeys "{TAB}"
				Set WshShell = Nothing
				wait 1
				If ubound(arrDateTime) =1 Then
					objGenerateABSDialog.JavaList("ManufacturingDate").Select arrDateTime(1)
				End If
			End If
			
            ' set  Installation Time 
			If sInstallationTime <> "" Then
				objGenerateABSDialog.JavaStaticText("Property_Name").SetTOProperty "label", "Installation Time :"
				wait 1
			    arrDateTime = Split(sInstallationTime," ")
				objGenerateABSDialog.JavaEdit("InstallationTime").Set arrDateTime(0)
				wait 1
				Set WshShell = CreateObject("WScript.Shell")
				WshShell.SendKeys "{TAB}"
				Set WshShell = Nothing
				wait 1
				If ubound(arrDateTime) =1 Then
					objGenerateABSDialog.JavaList("InstallationTime").Select arrDateTime(1)
				End If
			End If

			' set  Number of Level 
			If sNoofLevel <> "" Then
			objGenerateABSDialog.JavaStaticText("Property_Name").SetTOProperty "label", "Number of Levels:"
			wait 1
			    Call Fn_UI_EditBox_Type("Fn_ABM_GenerateAsBuiltStructure", objGenerateABSDialog, "NumberofLevels",sNoofLevel )
			End If
			wait 1
			'clicking on OK button
			Call Fn_Button_Click("Fn_ABM_GenerateAsBuiltStructure", objGenerateABSDialog, "OK")
			Fn_ABM_GenerateAsBuiltStructure = True
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_ABM_GenerateAsBuiltStructure ] Invalid case [ " & sAction & " ].")
			Fn_ABM_GenerateAsBuiltStructure = False
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	End Select

	If Fn_ABM_GenerateAsBuiltStructure = True Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_ABM_GenerateAsBuiltStructure ] executed successfully with case [ " & sAction & " ].")
	End If

	Set objGenerateABSDialog = Nothing
End Function


'****************************************    Function to perform Assign lot operation ***************************************
'
''Function Name		 	:		Fn_ABM_AssignLot
'
''Description		    :  	    Function to perform  assign lot operation

''Parameters		    :	 	1. sAction : Action need to perform
'					   		    2. sOpenDialogBy : to open Lot dialog ( AssignLot_RMB / AssignLot_Menu )
'					   		   3. sRootStructureNode : Root Structure Node Path
'							  4. sLotNumber : Lot Number 
'                             5. sManufactureID : Manufacture' ID 
								
''Return Value		    :  		True \ False
'
''Pre-requisite		    :		As Build Manager perspective should be selected

''Examples		     	:	Fn_ABM_AssignLot("AssignLot", "AssignLot_Menu", "000087/A;1-Top (View)", "L-1" ,"")

'History:
'						Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'						Rupali Palhade		 21-Sept-2011			1.0					 Created
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'						Shweta Rathod		 10-July-2015			1.0					 Modified case "AssignLot"  'work around implemeted for PR#7436486  by shweta rathod on 06-july-2014
'						[TC11.2 Maintenence : Build(2015062400) By Vivek Ahirrao] 						
'-------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_ABM_AssignLot(sAction, sOpenDialogBy, sRootStructureNode, sLotNumber ,sManufactureID)
	GBL_FAILED_FUNCTION_NAME="Fn_ABM_AssignLot"
	Dim objAssignLot, bReturn , sLotIndex
	Dim sMru, objABMWindowApplet
	Fn_ABM_AssignLot = False
	Set objAssignLot = JavaWindow("AsBuiltManager").JavaWindow("AssignLot")
	sMru = split(sRootStructureNode,"~")
	sRootStructureNode = sMru(0)
	If Fn_UI_ObjectExist("Fn_ABM_AssignLot",objAssignLot) = False then
	
	arrNode = split(sRootStructureNode, ":", -1, 1)	
	bReturn = Fn_ABM_RootStructureTableOperations("ExpandBelow", "", "", arrNode(0), "", "", "")
	wait(1)

		Select Case sOpenDialogBy
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
			Case "AssignLot_RMB"
					If sRootStructureNode <> "" then
						bReturn = Fn_ABM_RootStructureTableOperations("PopupSelect", "", "", sRootStructureNode, "", "", "Assign Lot...")
						If bReturn = False Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_ABM_AssignLot ] Failed to perform [ Tools:Assign Lot...] on Root Strcture Node [ " & sRootStructureNode & " ].")
							Set objAssignLot = Nothing
							Exit function
						End If
					End If
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
			Case "AssignLot_Menu", ""
					If sRootStructureNode <> "" then
						bReturn = Fn_ABM_RootStructureTableOperations("Select", "", "", sRootStructureNode, "", "", "")
						If bReturn = False Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_ABM_AssignLot ] Failed to select Root Strcture Node [ " & sRootStructureNode & " ].")
							Set objAssignLot = Nothing
							Exit function
						End If
					End If
					wait(1)
					Call Fn_MenuOperation("Select","Tools:Assign Lot...")
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
			End Select

			If Fn_UI_ObjectExist("Fn_ABM_AssignLot",objAssignLot) = False then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_ABM_AssignLot ] Failed to open Assign Lot  Window.")
				Set objAssignLot = Nothing
				Exit function
			End If

	End IF

	Select Case sAction
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "AssignLot"
			' set  Lot Number
			If sLotNumber <> "" Then
				
				sLotIndex = objAssignLot.JavaList("LotNumber").GetItemIndex(sLotNumber)
				objAssignLot.JavaList("LotNumber").Select "#"&sLotIndex
			End If

			' set  Manufacturer's ID 
			If sManufactureID <> "" Then
			Call Fn_UI_EditBox_Type("Fn_ABM_AssignLot", objAssignLot, "ManufacturerID",sManufactureID )
			End If

			'clicking on OK button
			Call Fn_Button_Click("Fn_ABM_AssignLot", objAssignLot, "OK")
			
			'work around implemeted for PR#7436486  by shweta rathod on 06-july-2014	[TC11.2 Maintenence : Build(2015062400) By Vivek Ahirrao]	
			If uBound(sMru) > 0 then
				Call  Fn_ABM_RootStructureTableOperations("TabPopupMenuSelect", "("+sMru(1)+")", "", "", "", "", "Close Panel")
				wait 1
				Set objABMWindowApplet=Fn_UI_ObjectCreate("Fn_ABM_AssignLot",JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame"))
				Call Fn_CheckBox_Set("Fn_ABM_AssignLot",objABMWindowApplet, "MRUButton",  "ON" )
				wait 2
				objABMWindowApplet.JavaButton("MRUListButton").SetTOProperty "label",sMru(1)
				Call Fn_Button_Click("Fn_ABM_AssignLot",objABMWindowApplet,"MRUListButton")
			End if
			'end of work around implemeted for PR#7436486  by shweta rathod on 06-july-2014
			Fn_ABM_AssignLot = True
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_ABM_AssignLot ] Invalid case [ " & sAction & " ].")
			Fn_ABM_AssignLot = False
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	End Select

	If Fn_ABM_AssignLot = True Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_ABM_AssignLot ] executed successfully with case [ " & sAction & " ].")
	End If

	Set objAssignLot = Nothing
End Function
'****************************************    Function to perform Find and Select Physical Part ***************************************
'
''Function Name		 	:	Fn_SISW_ABM_SearchDialogOperations
'
''Description		    :  	Function to perform  assign lot operation

''Parameters		    :	1. sAction : Action need to perform
'					   		2. sSearchType : to open Lot dialog ( AssignLot_RMB / AssignLot_Menu )
'					   		3. dicABMSearchDialog
								
''Return Value		    :  	True \ False
'
''Pre-requisite		    :	Search Window should be present.

''Examples		     	:	
'							Dim dicABMSearchDialog
'							Set dicABMSearchDialog = CreateObject( "Scripting.Dictionary" )
							'dicABMSearchDialog("bClear") = True
							'dicABMSearchDialog("PartNumber")
							'dicABMSearchDialog("bSerialize") = True
							'dicABMSearchDialog("SerialNumber") 
							'dicABMSearchDialog("SerialNumberAfter")
							'dicABMSearchDialog("SerialNumberBefore")
							'dicABMSearchDialog("bLot") = False
							'dicABMSearchDialog("LotNumber")
							'dicABMSearchDialog("ManufacturerID")
							'dicABMSearchDialog("ManufactureredAfterDate")
							'dicABMSearchDialog("ManufactureredBeforeDate")
							'dicABMSearchDialog("ManufactureredBeforeDate")
							'dicABMSearchDialog("ManufactureredBeforeTime")

'Examples		     	:	Case "InstallablePhysicalParts"
'								dicABMSearchDialog("PhysicalPart")
'								Fn_SISW_ABM_SearchDialogOperations("FindAndSelect", "InstallablePhysicalParts", dicABMSearchDialog)

'Examples		     	:	Case "ReplacePhysicalParts"
'								dicABMSearchDialog("SearchCriteria") = "Name=A*"
'								dicABMSearchDialog("AlternateParts")
'								dicABMSearchDialog("SubstituteParts")
'								dicABMSearchDialog("DeviationItem") =
'								Call Fn_SISW_ABM_SearchDialogOperations("FindAndSelect", "ReplacePhysicalParts", dicABMSearchDialog)

'Examples		     	:	Case "SetDeviation_FindAndSelect"
'								dicABMSearchDialog("SearchCriteria") = "Name=A*"
'								dicABMSearchDialog("DocumentItem") = "000001/A;-Asd"
'								Call Fn_SISW_ABM_SearchDialogOperations("SetDeviation_FindAndSelect", "SetDeviation", dicABMSearchDialog)
'-----------------------------------------------------------------------------------------------------------------------------------
'								Call Fn_SISW_ABM_SearchDialogOperations("CloseDialog", "", "")
'-----------------------------------------------------------------------------------------------------------------------------------
'								dicABMSearchDialog("sButton") = "Find"
'								Call Fn_SISW_ABM_SearchDialogOperations("IsDialogExists", "", dicABMSearchDialog)

'History:
'	Developer Name			Date			Rev. No.		Reviewer		Changes Done	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Koustubh Watwe		 31-May-2012		1.0				
'-----------------------------------------------------------------------------------------------------------------------------------
'	Koustubh Watwe		 07-Jun-2012		1.0				Koustubh		Added cases IsDialogExists, FindAndSelect
'-----------------------------------------------------------------------------------------------------------------------------------
'	Koustubh Watwe		 21-Jun-2012		1.0				Koustubh		Added cases SetDeviation_FindAndSelect
'-----------------------------------------------------------------------------------------------------------------------------------
'	Koustubh Watwe		 04-Oct-2012		1.0				Koustubh		Added cases SetDeviation
'-----------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_ABM_SearchDialogOperations(sAction, sSearchType, dicABMSearchDialog)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_ABM_SearchDialogOperations"
	Dim objSearch, sTitle, iCnt, bFlag, iRowCount, aSearchCri , aFieldSet
	Set objShell = JavaWindow("Shell")
	Set objSearch = objShell.JavaWindow("Search")
	bFlag = False
	Fn_SISW_ABM_SearchDialogOperations = False
	Select Case sSearchType
		Case "InstallablePhysicalParts", "ReplacePhysicalParts"
				sTitle = "Installable Physical Parts"
				objSearch.SetTOProperty "title", sTitle
				For iCnt = 0 to 50
					objShell.SetTOProperty "Index", iCnt
					If objSearch.Exist(2) Then
						bFlag = True
						Exit for
					End If
				Next

		Case "SetDeviation"
			Set objSearch = JavaWindow("AsBuiltManager").JavaWindow("SetupDeviation").JavaWindow("Search")
			If objSearch.Exist(5) Then
				bFlag = true
			End If
		Case Else
				sTitle = "Search"
				objSearch.SetTOProperty "title", sTitle
				For iCnt = 0 to 50
					objShell.SetTOProperty "Index", iCnt
					If objSearch.Exist(2) Then
						bFlag = True
						Exit for
					End If
				Next
	End Select

	If Not(bFlag) Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_ABM_SearchDialogOperations ] Failed to find [ " & sTitle & " ] window.")
		Exit function
	End If

	Select Case sAction
		Case "CloseDialog"
				Call Fn_Button_Click("Fn_SISW_ABM_SearchDialogOperations", objSearch, "Cancel")
				Fn_SISW_ABM_SearchDialogOperations = True
		Case "IsDialogExists"
				If dicABMSearchDialog("sButtonName")  <> "" Then
					Call Fn_Button_Click("Fn_SISW_ABM_SearchDialogOperations", objSearch, dicABMSearchDialog("sButtonName"))
				End If
				Fn_SISW_ABM_SearchDialogOperations = True
		Case "FindAndSelect"
			' clearing fields
			If dicABMSearchDialog("bClear") <> "" Then
				If cBool(dicABMSearchDialog("bClear")) Then
					Call Fn_Button_Click("Fn_SISW_ABM_SearchDialogOperations", objSearch, "Clear")
				End If
			End If
			'setting part number
			If dicABMSearchDialog("PartNumber") <> "" Then
				objSearch.JavaEdit("PartNumber").Activate
				Call Fn_Edit_Box("Fn_SISW_ABM_SearchDialogOperations", objSearch,"PartNumber",dicABMSearchDialog("PartNumber"))
			End If

			' setting serialixed radio button
			If dicABMSearchDialog("bSerialize") <> "" Then
				If cBool(dicABMSearchDialog("bSerialize")) Then
					objSearch.JavaRadioButton("SerializedRadioButton").SetTOProperty "attached text","true" 
				Else
					objSearch.JavaRadioButton("SerializedRadioButton").SetTOProperty "attached text","false" 
				End If
				Call Fn_UI_JavaRadioButton_SetON("Fn_SISW_ABM_SearchDialogOperations", objSearch,"SerializedRadioButton")
			End If

			' setting serial number
			If dicABMSearchDialog("SerialNumber") <> "" Then
				Call Fn_Edit_Box("Fn_SISW_ABM_SearchDialogOperations", objSearch,"SerialNumber",dicABMSearchDialog("SerialNumber"))
			End If

			' setting serial number After
			If dicABMSearchDialog("SerialNumberAfter") <> "" Then
				Call Fn_Edit_Box("Fn_SISW_ABM_SearchDialogOperations", objSearch,"SerialNumberAfter",dicABMSearchDialog("SerialNumberAfter"))
			End If

			' setting serial number Before
			If dicABMSearchDialog("SerialNumberBefore") <> "" Then
				Call Fn_Edit_Box("Fn_SISW_ABM_SearchDialogOperations", objSearch,"SerialNumberBefore",dicABMSearchDialog("SerialNumberBefore"))
			End If

			' setting lot radio button
			If dicABMSearchDialog("bLot") <> "" Then
				If cBool(dicABMSearchDialog("bLot")) Then
					objSearch.JavaRadioButton("LotRadioButton").SetTOProperty "attached text","true" 
				Else
					objSearch.JavaRadioButton("LotRadioButton").SetTOProperty "attached text","false" 
				End If
				Call Fn_UI_JavaRadioButton_SetON("Fn_SISW_ABM_SearchDialogOperations", objSearch,"LotRadioButton")
			End If

			' setting Lot number
			If dicABMSearchDialog("LotNumber") <> "" Then
				Call Fn_Edit_Box("Fn_SISW_ABM_SearchDialogOperations", objSearch,"LotNumber",dicABMSearchDialog("LotNumber"))
			End If

			' setting Lot number
			If dicABMSearchDialog("ManufacturerID") <> "" Then
				Call Fn_Edit_Box("Fn_SISW_ABM_SearchDialogOperations", objSearch,"ManufacturerID",dicABMSearchDialog("ManufacturerID"))
			End If

			If dicABMSearchDialog("ManufactureredAfterDate") <> "" Then
				Call Fn_Button_Click("Fn_SISW_ABM_SearchDialogOperations", objSearch, "ManufacturedAfterDateButton")
				wait 1
				Call Fn_UI_SetDateAndTime("Fn_SISW_ABM_SearchDialogOperations",dicABMSearchDialog("ManufactureredAfterDate"),dicABMSearchDialog("ManufactureredAfterTime"))				
			End If

			If dicABMSearchDialog("ManufactureredBeforeDate") <> "" Then
				Call Fn_Button_Click("Fn_SISW_ABM_SearchDialogOperations", objSearch, "ManufacturedBeforeDateButton")
				wait 1
				Call Fn_UI_SetDateAndTime("Fn_SISW_ABM_SearchDialogOperations",dicABMSearchDialog("ManufactureredBeforeDate"),dicABMSearchDialog("ManufactureredBeforeTime"))				
			End If

			' clicking on FInd button
			Call Fn_Button_Click("Fn_SISW_ABM_SearchDialogOperations", objSearch, "Find")
            objSearch.Maximize

   			Select Case sSearchType
				'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				Case "InstallablePhysicalParts"
					If dicABMSearchDialog("PhysicalPart") <> ""  Then
						iRowCount = cInt(objSearch.JavaTable("SearchResultList").GetROProperty("rows"))
						For iCnt = 0 to iRowCount -1
							If objSearch.JavaTable("SearchResultList").Object.getItem(iCnt).getData().toString() = dicABMSearchDialog("PhysicalPart") then
								wait 2
								objSearch.JavaTable("SearchResultList").ActivateRow iCnt
								Exit for
							End If
						Next
					End If
				'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				Case "ReplacePhysicalParts"
					If dicABMSearchDialog("PreferredParts") <> ""  Then
						wait 2
						iRowCount = cInt(objSearch.JavaTable("PreferredParts").GetROProperty("rows"))
						wait 2
						For iCnt = 0 to iRowCount -1
							If objSearch.JavaTable("PreferredParts").Object.getItem(iCnt).getData().toString() = dicABMSearchDialog("PreferredParts") then
								wait 2
								objSearch.JavaTable("PreferredParts").ActivateRow iCnt
								Exit for
							End If
						Next
					End If
					If dicABMSearchDialog("AlternateParts") <> ""  Then
						wait 2
						iRowCount = cInt(objSearch.JavaTable("AlternateParts").GetROProperty("rows"))
						wait 2
						For iCnt = 0 to iRowCount 
							wait 5
							'If objSearch.JavaTable("AlternateParts").GetCellData(iCnt,0) = dicABMSearchDialog("AlternateParts") then
							If objSearch.JavaTable("AlternateParts").Object.getItem(iCnt).getData().tostring() = dicABMSearchDialog("AlternateParts") then
								objSearch.JavaTable("AlternateParts").ActivateRow iCnt
								Exit for
							End If
						Next
					End If
					If dicABMSearchDialog("SubstituteParts") <> ""  Then
						wait 2
						iRowCount = cInt(objSearch.JavaTable("SubstituteParts").GetROProperty("rows"))
						wait 2
						For iCnt = 0 to iRowCount -1
							wait 2
						   'If objSearch.JavaTable("SubstituteParts").GetCellData(iCnt,0) = dicABMSearchDialog("SubstituteParts") then
							If objSearch.JavaTable("SubstituteParts").Object.getItem(iCnt).getData().tostring() = dicABMSearchDialog("SubstituteParts") then
							Wait 3
								objSearch.JavaTable("SubstituteParts").ActivateRow iCnt
								Exit for
							End If
						Next
					End If
					If dicABMSearchDialog("DeviatedParts") <> "" Then
						wait 2
						iRowCount = cInt(objSearch.JavaTable("DeviatedParts").GetROProperty("rows"))
						wait 2
						For iCnt = 0 to iRowCount -1
							If objSearch.JavaTable("DeviatedParts").Object.getItem(iCnt).getData().toString()= dicABMSearchDialog("DeviatedParts") Then
								wait 2
								objSearch.JavaTable("DeviatedParts").ActivateRow iCnt
								Exit for
							End If
						Next
					End If
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			Case Else
					If dicABMSearchDialog("SearchResultsPart") <> ""  Then
						wait 2
						iRowCount = cInt(objSearch.JavaTable("SearchResultList").GetROProperty("rows"))
						wait 2
						For iCnt = 0 to iRowCount -1
							If objSearch.JavaTable("SearchResultList").Object.getItem(iCnt).getData().toString() = dicABMSearchDialog("SearchResultsPart") then
								wait 2
								objSearch.JavaTable("SearchResultList").ActivateRow iCnt
								Exit for
							End If
						Next
					End If
			End Select
			
			If objSearch.Exist(10) Then
				Call Fn_Button_Click("Fn_SISW_ABM_SearchDialogOperations", objSearch, "OK")
			End If
			Fn_SISW_ABM_SearchDialogOperations = True
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "SetDeviation_FindAndSelect"
			aSearchCri = split(dicABMSearchDialog("SearchCriteria"),"~")
			Dim aDate
			For iCnt = 0 to uBound(aSearchCri)
				aFieldSet = split(aSearchCri(iCnt),"=")
				objSearch.JavaStaticText("Field_label").SetTOProperty "label", aFieldSet(0) & ":"
				If objSearch.JavaButton("DateButton").Exist(5) Then
					' set Date
					aDate = split(aFieldSet(1)," ")
					Call Fn_Button_Click("Fn_SISW_ABM_SearchDialogOperations", objSearch, "DateButton")
					wait 1
					Call Fn_UI_SetDateAndTime("Fn_SISW_ABM_SearchDialogOperations",aDate(0),aDate(1))				
				ElseIf objSearch.JavaEdit("FieldEditbox").Exist(5) Then
					Call Fn_Edit_Box("Fn_SISW_ABM_SearchDialogOperations", objSearch,"FieldEditbox",aFieldSet(1))
				Else
					' No field found...
					Exit function
				End If
			Next

			' clicking on FInd button
			Call Fn_Button_Click("Fn_SISW_ABM_SearchDialogOperations", objSearch, "Find")
			wait 2
			If dicABMSearchDialog("DocumentItem") <> ""  Then
				iRowCount = cInt(objSearch.JavaTable("SearchResultList").GetROProperty("rows"))
				For iCnt = 0 to iRowCount -1
					If objSearch.JavaTable("SearchResultList").GetCellData(iCnt,0) = dicABMSearchDialog("DocumentItem") then
						objSearch.JavaTable("SearchResultList").ActivateCell  iCnt,0
					Exit for
				End If
				Next
			End If

			If objSearch.Exist(15) Then
				Call Fn_Button_Click("Fn_SISW_ABM_SearchDialogOperations", objSearch, "OK")
			End If
			Fn_SISW_ABM_SearchDialogOperations = True
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_ABM_SearchDialogOperations ] invalied case [ " & sAction & " ] on  [ " & sTitle & " ] window.")
	End Select

	If  Fn_SISW_ABM_SearchDialogOperations <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [ Fn_SISW_ABM_SearchDialogOperations ] executed successfuly with case [ " & sAction & " ] on  [ " & sTitle & " ] window.")
	End If

	Set objShell = Nothing
	Set objSearch = Nothing
End Function
'****************************************    Function to perform Install Physical Part ***************************************
'
''Function Name		 	:	Fn_SISW_ABM_InstallPhysicalPartOperations
'
''Description		    :  	Function to perform  Install Physical Part

''Parameters		    :	1. sAction : Action need to perform
'					   		2. sOpenDialogBy : to open Lot dialog ( AssignLot_RMB / AssignLot_Menu )
'					   		3. sRootStructureNode
'					   		4. dicInstallPhysicalPart
								
''Return Value		    :  	True \ False
'
''Pre-requisite		    :	As Build Manager perspective should be present.

''Examples		     	:	
'							Dim dicInstallPhysicalPart
'							Set dicInstallPhysicalPart = CreateObject( "Scripting.Dictionary" )
							'dicInstallPhysicalPart("Usage") = ""

							'dicInstallPhysicalPart("ToBeInstalledPhysicalPart") = True
							'dicInstallPhysicalPart("bClear") = True
							'dicInstallPhysicalPart("PartNumber")
							'dicInstallPhysicalPart("bSerialize") = True
							'dicInstallPhysicalPart("SerialNumber") 
							'dicInstallPhysicalPart("SerialNumberAfter")
							'dicInstallPhysicalPart("SerialNumberBefore")
							'dicInstallPhysicalPart("bLot") = False
							'dicInstallPhysicalPart("LotNumber")
							'dicInstallPhysicalPart("ManufacturerID")
							'dicInstallPhysicalPart("ManufactureredAfterDate")
							'dicInstallPhysicalPart("ManufactureredBeforeDate")
							'dicInstallPhysicalPart("ManufactureredBeforeDate")
							'dicInstallPhysicalPart("ManufactureredBeforeTime")
							'dicInstallPhysicalPart("PhysicalPart")

							'dicInstallPhysicalPart("InstallationDate") = ""
							'dicInstallPhysicalPart("InstallationTime") = ""
							'dicInstallPhysicalPart("ExtraToDesign") = ""

''Examples		     	:	Fn_SISW_ABM_InstallPhysicalPartOperations("InstallPhysicalPart", "RMB", "000087/A;1-Top (View)", dicInstallPhysicalPart)
''Examples		     	:	Fn_SISW_ABM_InstallPhysicalPartOperations("InstallPhysicalPart", "SrvMgr_Menu", "", dicInstallPhysicalPart)

'History:
'	Developer Name			Date			Rev. No.		Reviewer		Changes Done	
'-----------------------------------------------------------------------------------------------------------------------------------
'	Koustubh Watwe		 31-May-2012		1.0				
'-----------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_ABM_InstallPhysicalPartOperations(sAction, sOpenDialogBy, sRootStructureNode, dicInstallPhysicalPart)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_ABM_InstallPhysicalPartOperations"
	Dim objInstallPhysicalParts, bFlag, iCnt, iItemCount
	Dim objSelectType, objUsageTable, iWidth
	Dim sMenu, arrDate

	Set objInstallPhysicalParts = Fn_SISW_ABM_GetObject("InstallPhysicalPart")
'	Set objInstallPhysicalParts = JavaWindow("AsBuiltManager").JavaWindow("InstallPhysicalPart")
	bFlag = False
	Fn_SISW_ABM_InstallPhysicalPartOperations = False

	'If Fn_UI_ObjectExist("Fn_SISW_ABM_InstallPhysicalPartOperations", objInstallPhysicalParts) = False Then
	If Fn_SISW_UI_Object_Operations("Fn_SISW_ABM_InstallPhysicalPartOperations","Exist", objInstallPhysicalParts,SISW_MINLESS_TIMEOUT) = False Then
		Select Case sOpenDialogBy
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
				Case "RMB", "SrvMgr_RMB"
					sMenu = "Install Physical Part"
					If sOpenDialogBy = "SrvMgr_RMB" Then sMenu = sMenu & "..."
					 
					If sRootStructureNode <> "" then
						bReturn = Fn_ABM_RootStructureTableOperations("PopupSelect", "", "", sRootStructureNode, "", "", sMenu)
						If bReturn = False Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_ABM_InstallPhysicalPartOperations ] Failed to perform [ RMB : "& sMenu &" ] on Root Strcture Node [ " & sRootStructureNode & " ].")
							Set objInstallPhysicalParts = Nothing
							Exit function
						End If
					End If
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
			Case "Menu", "", "SrvMgr_Menu"
					If sRootStructureNode <> "" then
						bReturn = Fn_ABM_RootStructureTableOperations("Select", "", "", sRootStructureNode, "", "", "")
						If bReturn = False Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_ABM_InstallPhysicalPartOperations ] Failed to select Root Strcture Node [ " & sRootStructureNode & " ].")
							Set objInstallPhysicalParts = Nothing
							Exit function
						End If
					End If
					sMenu = "Tools:Install Physical Part"
					If sOpenDialogBy = "SrvMgr_Menu" Then sMenu = sMenu & "..."
					
					Call Fn_MenuOperation("Select",sMenu)
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		End Select
		Call Fn_ReadyStatusSync(2)
		If Fn_UI_ObjectExist("Fn_SISW_ABM_InstallPhysicalPartOperations", objInstallPhysicalParts) = False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_ABM_InstallPhysicalPartOperations ] Failed to find [ Install Physical Part ] window.")
			Set objInstallPhysicalParts = Nothing
			Exit function
		End If
	End If

	Select Case sAction
		Case "InstallPhysicalPart"
			' select usage
			If dicInstallPhysicalPart("Usage") <> "" Then
				iItemCount = cInt(objInstallPhysicalParts.JavaObject("UsageCombo").Object.getItemCount())
				For iCnt = 0 to iItemCount -1
					If dicInstallPhysicalPart("Usage") = objInstallPhysicalParts.JavaObject("UsageCombo").Object.getItem(iCnt).getText() Then
'						objInstallPhysicalParts.JavaObject("UsageCombo").Object.select(iCnt)
						iWidth = cInt(objInstallPhysicalParts.JavaObject("UsageCombo").GetROProperty("width"))
						objInstallPhysicalParts.JavaObject("UsageCombo").Click iWidth / 2 ,10,"LEFT"
						wait 3
						Dim WshShell
						Set WshShell = CreateObject("WScript.Shell")
						For iItemCount=0 to iCnt
							WshShell.SendKeys "{DOWN}"
							wait 1
						Next
						WshShell.SendKeys "{ENTER}"
						Set WshShell =Nothing

'						Set objSelectType=description.Create()
'						objSelectType("Class Name").value = "JavaTable"
'						objSelectType("path").value = "Table;Shell;Shell;Shell;"
'						objSelectType("toolkit class").value = "org.eclipse.swt.widgets.Table"
'						Set  objUsageTable = objInstallPhysicalParts.ChildObjects(objSelectType)
'						objUsageTable(0).ClickCell iCnt,0
						Exit for
					End If
				Next
			End If

			' setting To Be Installed Physical Part
			If dicInstallPhysicalPart("ToBeInstalledPhysicalPart") Then
				If cBool(dicInstallPhysicalPart("ToBeInstalledPhysicalPart")) Then
					objInstallPhysicalParts.JavaStaticText("ToBeInstalledPhysicalDropDown").Click 1,1,"LEFT"
					wait 1
					objInstallPhysicalParts.JavaMenu("label:=Add").Select
					If dicInstallPhysicalPart("SearchResultType") <> "" Then
						bFlag = Fn_SISW_ABM_SearchDialogOperations("FindAndSelect", dicInstallPhysicalPart("SearchResultType"), dicInstallPhysicalPart)
						If bFlag = False Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_ABM_InstallPhysicalPartOperations ] Failed to execute function Fn_SISW_ABM_SearchDialogOperations to find and select Physical Part")
							Exit Function
						End If
					Else
						bFlag = Fn_SISW_ABM_SearchDialogOperations("FindAndSelect", "InstallablePhysicalParts", dicInstallPhysicalPart)
						If bFlag = False Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_ABM_InstallPhysicalPartOperations ] Failed to execute function Fn_SISW_ABM_SearchDialogOperations to find and select Physical Part")
							Exit Function
						End If
					End If
				End If
			End If

			' setting Installation date
			If dicInstallPhysicalPart("InstallationDate") <> "" Then    ''modified to handle new date control object on  26-Jun-2014
				arrDate = split(dicInstallPhysicalPart("InstallationDate"), "~")
				objInstallPhysicalParts.JavaEdit("InstallationTimeEditBox").Set arrDate(0)
				Call Fn_KeyBoardOperation("SendKeys", "{TAB}")
				wait 1
				If UBound(arrDate)>0 then
					objInstallPhysicalParts.JavaList("Time").Type arrDate(1)
					Call Fn_KeyBoardOperation("SendKeys", "{TAB}")
				End if
			End If
			If dicInstallPhysicalPart("InstallationTime") <> "" Then
				objInstallPhysicalParts.JavaList("Time").Type dicInstallPhysicalPart("InstallationTime")
				Call Fn_KeyBoardOperation("SendKeys", "{TAB}")
			End If

			' extra to design
			If dicInstallPhysicalPart("ExtraToDesign") <> "" Then
				Call Fn_Edit_Box("Fn_SISW_ABM_InstallPhysicalPartOperations", objInstallPhysicalParts, "ExtraToDesignEditBox", dicInstallPhysicalPart("ExtraToDesign")) 
			End If
			wait 65  'Added by Vrushali on 2-Apr-2013
			' clicking on OK button
			If dicInstallPhysicalPart("sButtonName") = "NA" Then
				'Do  Nothing 
			ElseIf dicInstallPhysicalPart("sButtonName") <> "" Then
				Call Fn_Button_Click("Fn_SISW_ABM_SearchDialogOperations", objInstallPhysicalParts, dicInstallPhysicalPart("sButtonName"))
			 Else
				Call Fn_Button_Click("Fn_SISW_ABM_SearchDialogOperations", objInstallPhysicalParts, "OK")
			End If
            Fn_SISW_ABM_InstallPhysicalPartOperations = True			
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_ABM_InstallPhysicalPartOperations ] invalied case [ " & sAction & " ].")
	End Select

	If  Fn_SISW_ABM_InstallPhysicalPartOperations <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [ Fn_SISW_ABM_InstallPhysicalPartOperations ] executed successfuly with case [ " & sAction & " ].")
	End If

	Set objInstallPhysicalParts = Nothing
End Function
'****************************************    Function to perform Replace Physical Part ***************************************
'
''Function Name		 	:	Fn_SISW_ABM_ReplacePhysicalPartOperations
'
''Description		    :  	Function to perform  Replace Physical Part

''Parameters		    :	1. sAction : Action need to perform
'					   		2. sOpenDialogBy : to open Lot dialog ( AssignLot_RMB / AssignLot_Menu )
'					   		3. sRootStructureNode
'					   		4. dicReplacePhysicalPart
								
''Return Value		    :  	True \ False
'
''Pre-requisite		    :	As Build Manager perspective should be present.

''Examples		     	:	dicReplacePhysicalPart("ToBeInstalledPhysicalPart") = True
'										dicReplacePhysicalPart("bClear")
'										dicReplacePhysicalPart("PartNumber")
'										dicReplacePhysicalPart("bSerialize") = True
'										dicReplacePhysicalPart("SerialNumber") 
'										dicReplacePhysicalPart("SerialNumberAfter")
'										dicReplacePhysicalPart("SerialNumberBefore")
'										dicReplacePhysicalPart("bLot") = False
'										dicReplacePhysicalPart("LotNumber")
'										dicReplacePhysicalPart("ManufacturerID")
'										dicReplacePhysicalPart("ManufactureredAfterDate")
'										dicReplacePhysicalPart("ManufactureredBeforeDate")
'										dicReplacePhysicalPart("ManufactureredBeforeDate")
'										dicReplacePhysicalPart("ManufactureredBeforeTime")
'										dicReplacePhysicalPart("PreferredParts")
'										dicReplacePhysicalPart("AlternateParts")
'										dicReplacePhysicalPart("SubstituteParts")
'										dicReplacePhysicalPart("DeviatedParts")

'										dicReplacePhysicalPart("SelectedUsage")
'										dicReplacePhysicalPart("ReplaceDate")
'										dicReplacePhysicalPart("ReplaceTime")

'										Fn_SISW_ABM_ReplacePhysicalPartOperations("ReplacePhysicalPart", "RMB", "000087/A;1-Top (View)", dicReplacePhysicalPart)
''Examples		     	:	Fn_SISW_ABM_ReplacePhysicalPartOperations("ReplacePhysicalPart", "", "", dicReplacePhysicalPart)

'History:
'	Developer Name			Date			Rev. No.		Reviewer		Changes Done	
'-----------------------------------------------------------------------------------------------------------------------------------
'	Koustubh Watwe		 31-May-2012		1.0				
'-----------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_ABM_ReplacePhysicalPartOperations(sAction, sOpenDialogBy, sRootStructureNode, dicReplacePhysicalPart)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_ABM_ReplacePhysicalPartOperations"
	Dim objReplacePhysicalParts, bFlag, iCnt, iItemCount,WshShell
	Set objReplacePhysicalParts = Fn_SISW_ABM_GetObject("ReplacePhysicalPart")
'	Set objReplacePhysicalParts = JavaWindow("AsBuiltManager").JavaWindow("ReplacePhysicalPart")
	bFlag = False
	Fn_SISW_ABM_ReplacePhysicalPartOperations = False

	If Fn_UI_ObjectExist("Fn_SISW_ABM_ReplacePhysicalPartOperations", objReplacePhysicalParts) = False Then
		Select Case sOpenDialogBy
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
			Case "RMB"
					If sRootStructureNode <> "" then
						bReturn = Fn_ABM_RootStructureTableOperations("PopupSelect", "", "", sRootStructureNode, "", "", "Replace Physical Part")
						If bReturn = False Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_ABM_ReplacePhysicalPartOperations ] Failed to perform [ RMB : Replace Physical Part ] on Root Strcture Node [ " & sRootStructureNode & " ].")
							Set objReplacePhysicalParts = Nothing
							Exit function
						End If
					End If
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
			Case "Menu", ""
					If sRootStructureNode <> "" then
						bReturn = Fn_ABM_RootStructureTableOperations("Select", "", "", sRootStructureNode, "", "", "")
						If bReturn = False Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_ABM_ReplacePhysicalPartOperations ] Failed to select Root Strcture Node [ " & sRootStructureNode & " ].")
							Set objReplacePhysicalParts = Nothing
							Exit function
						End If
					End If
					Call Fn_MenuOperation("Select","Tools:Replace Physical Part")
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		End Select
	
		If Fn_UI_ObjectExist("Fn_SISW_ABM_InstallPhysicalPartOperations", objReplacePhysicalParts) = False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_ABM_ReplacePhysicalPartOperations ] Failed to find [ Replace Physical Part ] window.")
			Set objReplacePhysicalParts = Nothing
			Exit function
		End If
	End If

	Select Case sAction
		Case "ReplacePhysicalPart"
			' setting To Be Installed Physical Part
			If dicReplacePhysicalPart("ToBeInstalledPhysicalPart") Then
				If cBool(dicReplacePhysicalPart("ToBeInstalledPhysicalPart")) Then
					objReplacePhysicalParts.JavaStaticText("ToBeInstalledPhysicalPartDropDown").Click 1,1,"LEFT"
					wait 1
					objReplacePhysicalParts.JavaMenu("label:=Add").Select
					bFlag = Fn_SISW_ABM_SearchDialogOperations("FindAndSelect", "ReplacePhysicalParts", dicReplacePhysicalPart)
					If bFlag = False Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_ABM_ReplacePhysicalPartOperations ] Failed to execute function Fn_SISW_ABM_SearchDialogOperations to find and select Physical Part")
						Exit Function
					End If
				End If
			End If

			If dicReplacePhysicalPart("SelectedUsage") <> "" Then
				Call Fn_Edit_Box("Fn_SISW_ABM_ReplacePhysicalPartOperations",objReplacePhysicalParts,"SelectedUsageEditBox",dicReplacePhysicalPart("SelectedUsage"))
			End If

			If dicReplacePhysicalPart("ReplaceDate") <> "" Then
				objReplacePhysicalParts.JavaEdit("ReplaceDateEditBox").Set dicReplacePhysicalPart("ReplaceDate")
				Set WshShell = CreateObject("WScript.Shell")
				WshShell.SendKeys "{TAB}"
				Set WshShell = Nothing
				wait 2
            End If

			If dicReplacePhysicalPart("ReplaceTime") <> "" Then
				objReplacePhysicalParts.JavaList("ReplaceDateListBox").Select dicReplacePhysicalPart("ReplaceTime")
            End If

			If dicReplacePhysicalPart("LocationName") <> "" Then
				Call Fn_List_Select("Fn_SISW_ABM_ReplacePhysicalPartOperations", objReplacePhysicalParts,"LocationName", dicReplacePhysicalPart("LocationName")) 
			End If

			If dicReplacePhysicalPart("DispositionValue") <> "" Then
				Call Fn_List_Select("Fn_SISW_ABM_ReplacePhysicalPartOperations", objReplacePhysicalParts,"DispositionValue", dicReplacePhysicalPart("DispositionValue")) 
			End If
			' clicking on OK button
			Call Fn_Button_Click("Fn_SISW_ABM_ReplacePhysicalPartOperations", objReplacePhysicalParts, "OK")
			Fn_SISW_ABM_ReplacePhysicalPartOperations = True
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_ABM_ReplacePhysicalPartOperations ] Invalid case [ " & sAction & " ].")
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
	End Select
	If  Fn_SISW_ABM_ReplacePhysicalPartOperations <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [ Fn_SISW_ABM_ReplacePhysicalPartOperations ] executed successfuly with case [ " & sAction & " ].")
	End If
	Set objReplacePhysicalParts = Nothing
End Function
'****************************************    Function to perform  Uninstall Physical Part ***************************************
'
''Function Name		 	:	Fn_SISW_ABM_UnInstallPhysicalPartOperations
'
''Description		    :  	Function to perform  Uninstall Physical Part

''Parameters		    :	 1. sOpenDialogBy : to open Uninstall dialog ( RMB / Menu )	   		
'							 2. sRootStructureNode : Root Structure node path
'							 3. sMessage : Message to verify
'							 4. sButton : Button Name
								
''Return Value		    :  	True \ False
'
''Pre-requisite		    :	

''Examples		     	:	1. Fn_SISW_ABM_UnInstallPhysicalPartOperations("RMB","000188/4c-A (View):000190/--A","Do you want to uninstall the Physical Part from As-Built structure?","")

'History:
'	Developer Name			Date			Rev. No.		Reviewer		Changes Done	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Vrushali Wani		     06-June-2012		1.0			Koustubh W
'-----------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_ABM_UnInstallPhysicalPartOperations(sOpenDialogBy,sRootStructureNode,sMessage,sButton)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_ABM_UnInstallPhysicalPartOperations"
	Dim objUnInstallPhysicalPart
	Set objUnInstallPhysicalPart = Fn_SISW_ABM_GetObject("UninstallPhysicalPart")
	bFlag = False
	Fn_SISW_ABM_UnInstallPhysicalPartOperations = False

	If Fn_UI_ObjectExist("Fn_SISW_ABM_InstallPhysicalPartOperations", objUnInstallPhysicalPart) = False Then
		Select Case sOpenDialogBy
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
			Case "RMB"
					If sRootStructureNode <> "" then
						bReturn = Fn_ABM_RootStructureTableOperations("PopupSelect", "", "", sRootStructureNode, "", "", "Uninstall Physical Part")
						If bReturn = False Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_ABM_UnInstallPhysicalPartOperations ] Failed to perform [ RMB : Uninstall Physical Part ] on Root Strcture Node [ " & sRootStructureNode & " ].")
							Set objUnInstallPhysicalPart = Nothing
							Exit function
						End If
					End If
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
			Case "Menu", ""
					If sRootStructureNode <> "" then
						bReturn = Fn_ABM_RootStructureTableOperations("Select", "", "", sRootStructureNode, "", "", "")
						If bReturn = False Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_ABM_UnInstallPhysicalPartOperations ] Failed to select Root Strcture Node [ " & sRootStructureNode & " ].")
							Set objUnInstallPhysicalPart = Nothing
							Exit function
						End If
					End If
					Call Fn_MenuOperation("Select","Tools:Uninstall Physical Part")
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		End Select
	
		If Fn_UI_ObjectExist("Fn_SISW_ABM_UnInstallPhysicalPartOperations", objUnInstallPhysicalPart) = False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_ABM_UnInstallPhysicalPartOperations ] Failed to find [ Uninstall Physical Part ] window.")
			Set objUnInstallPhysicalPart = Nothing
			Exit function
		End If
	End If

	If 	sMessage <> "" Then   
		sUninstallMessage = objUnInstallPhysicalPart.JavaStaticText("UninstallMessage").GetROProperty("value")
		If  sUninstallMessage = sMessage  Then
			Fn_SISW_ABM_UnInstallPhysicalPartOperations = True
		End If
	Else
		Fn_SISW_ABM_UnInstallPhysicalPartOperations = True
	End If

	' clicking on OK/Cancel button
	If sButton <> "" Then
		Call Fn_Button_Click("Fn_SISW_ABM_UnInstallPhysicalPartOperations", objUnInstallPhysicalPart, sButton)
	End If
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	If  Fn_SISW_ABM_UnInstallPhysicalPartOperations <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [ Fn_SISW_ABM_UnInstallPhysicalPartOperations ] executed successfuly .")
	End If
	Set objInstallPhysicalParts = Nothing
End Function
'****************************************    Function to perform Rebuild As-Built Strucuture ***************************************
''Function Name		 	:	Fn_SISW_ABM_RebuildAsBuiltStructure
'
''Description		    :  	Function to perform Rebuild As-Built Strucuture

''Parameters		    :	 1. sAction : Action to berform   		
'							         2. sRootStructureNode : Root Structure node path
'							         3. sPhysicalPart :Physical Part - for future use
'							         4. sRebuildDate : Rebuild date
'							         5. dicProperties : Dictionary object to handle Property dialog - for future use
								
''Return Value		    :  	Revision ID \ False
'
''Examples		     	:	Call Fn_SISW_ABM_RebuildAsBuiltStructure("RebuildAsBuiltStructure","000188/4c-A (View):000190/--A","","22-Sep-2011","")

'History:
'	Developer Name			Date			Rev. No.		Reviewer		Changes Done	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Vrushali Wani		     15-June-2012		1.0			Koustubh W
'-----------------------------------------------------------------------------------------------------------------------------------
'	Vrushali Wani		     19-June-2012		1.0			Koustubh W			Modfifiec function to return next Revision ID
'-----------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_ABM_RebuildAsBuiltStructure(sAction,sRootStructureNode,sPhysicalPart, sRebuildDate, dicProperties)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_ABM_RebuildAsBuiltStructure"
	Dim objRebuildABSDialog, bReturn , arrDateTime,sRevId,sPart
	Fn_SISW_ABM_RebuildAsBuiltStructure = False
	Set objRebuildABSDialog = Fn_SISW_ABM_GetObject("RebuildAsBuiltStructure")

	If Fn_UI_ObjectExist("Fn_SISW_ABM_RebuildAsBuiltStructure",objRebuildABSDialog) = False then
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
			If sRootStructureNode <> "" then
				bReturn = Fn_ABM_RootStructureTableOperations("Select","", "", sRootStructureNode, "", "", "")
				If bReturn = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_ABM_RebuildAsBuiltStructure ] Failed to select Root Strcture Node [ " & sRootStructureNode & " ].")
					Set objRebuildABSDialog = Nothing
					Exit function
				End If

				If instr(sAction,"Toolbar") > 0 Then
					' toolbar call
				Else
					Call Fn_MenuOperation("Select", "Tools:Rebuild As-Built Structure")
				End If
			End If
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
			If Fn_UI_ObjectExist("Fn_SISW_ABM_RebuildAsBuiltStructure",objRebuildABSDialog) = False then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_ABM_RebuildAsBuiltStructure ] Failed to open Rebuild As Built Structure Window.")
				Set objRebuildABSDialog = Nothing
				Exit function
			End If
	End IF


	Select Case sAction
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "RebuildAsBuiltStructure","Toolbar_RebuildAsBuiltStructure"

             ' code to handle Properties dialog
			' not implemented yet.

           If Fn_UI_ObjectExist("Fn_SISW_ABM_RebuildAsBuiltStructure",objRebuildABSDialog.JavaButton("Yes")) = True  then
				Call Fn_Button_Click("Fn_SISW_ABM_RebuildAsBuiltStructure", objRebuildABSDialog, "Yes")
			End If

			sPart= objRebuildABSDialog.JavaObject("SelectedPhysicalPart").Object.getText()
			sRevId = Chr( Asc(right(sPart,1)) + 1)
			
			' set Rebuild Date 

			If sRebuildDate <> "" Then
				If lcase(sRebuildDate) = "today" then
					arrDateTime = Split(Now," ")
				Else
					   arrDateTime = Split(sRebuildDate," ")
				End If
				objRebuildABSDialog.JavaEdit("RebuildDate").Set arrDateTime(0)
				wait 1
				call Fn_KeyBoardOperation("SendKeys", "{TAB}")
				objRebuildABSDialog.JavaList("Time").Type arrDateTime(1)
			End If

			
			'clicking on OK button
			Call Fn_Button_Click("Fn_SISW_ABM_RebuildAsBuiltStructure", objRebuildABSDialog, "OK")
			Fn_SISW_ABM_RebuildAsBuiltStructure = sRevId
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_ABM_RebuildAsBuiltStructure ] Invalid case [ " & sAction & " ].")
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	End Select

    If Fn_SISW_ABM_RebuildAsBuiltStructure <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SISW_ABM_RebuildAsBuiltStructure ] executed successfully with case [ " & sAction & " ].")
	End If
	Set objRebuildABSDialog = Nothing
End Function
'****************************************    Function to perform Rebuild As-Built Strucuture ***************************************
''Function Name		 	:	Fn_SISW_ABM_RebuildAsBuiltStructure
'
''Description		    :  	Function to perform Rebuild As-Built Strucuture

''Parameters		    :	 1. sAction : Action to berform   		
'							         2. sOpenDialogBy
'							         3. sRootStructureNode : Root Structure node path
'							         4. sSelectedPhysicalPart = Selected Physical Part
'							         5. sCopiedPhysicalPart : Copied Physical Part - for future use
'							         6. sDocumentID : Document ID
'							         7. dicProperties : Dictionary object to handle Property dialog - for future use
'							         8. dicSearch : Dictionary object to handle Search dialog
								
''Return Value		    :  	True \ False
'
''Examples		     	:	Call Fn_SISW_ABM_SetupDeviationOperations("SetupDeviation", "RMB", "000188/4c-A (View):000190/--A","","","","",dicSearch)

'History:
'	Developer Name			Date			Rev. No.		Reviewer		Changes Done	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Koustubh Watwe		21-June-2012		1.0			Koustubh W
'-----------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_ABM_SetupDeviationOperations(sAction, sOpenDialogBy, sRootStructureNode, sSelectedPhysicalPart, sCopiedPhysicalPart, sDocumentID, dicProperties, dicSearch)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_ABM_SetupDeviationOperations"
	Dim objShell, objSearch
    Dim objSetDev
'	Set objSetDev = Fn_SISW_ABM_GetObject("SetupDeviation")
	Set objSetDev = JavaWindow("AsBuiltManager").JavaWindow("SetupDeviation")
	bFlag = False
	Fn_SISW_ABM_SetupDeviationOperations = False

	If Fn_UI_ObjectExist("Fn_SISW_ABM_SetupDeviationOperations", objSetDev) = False Then
		Select Case sOpenDialogBy
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
			Case "RMB"
					If sRootStructureNode <> "" then
						bReturn = Fn_ABM_RootStructureTableOperations("PopupSelect", "", "", sRootStructureNode, "", "", "Setup Deviation...")
						If bReturn = False Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_ABM_SetupDeviationOperations ] Failed to perform [ RMB : Setup Deviation... ] on Root Strcture Node [ " & sRootStructureNode & " ].")
							Set objSetDev = Nothing
							Exit function
						End If
					End If
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
			Case "Menu", ""
					If sRootStructureNode <> "" then
						bReturn = Fn_ABM_RootStructureTableOperations("Select", "", "", sRootStructureNode, "", "", "")
						If bReturn = False Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_ABM_SetupDeviationOperations ] Failed to select Root Strcture Node [ " & sRootStructureNode & " ].")
							Set objSetDev = Nothing
							Exit function
						End If
					End If
					Call Fn_MenuOperation("Select","Tools:Setup Deviation...")
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		End Select
	
		If Fn_UI_ObjectExist("Fn_SISW_ABM_SetupDeviationOperations", objSetDev) = False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_ABM_SetupDeviationOperations ] Failed to find [ Setup Deviation ] window.")
			Set objSetDev = Nothing
			Exit function
		End If
	End If

	Select Case sAction
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "SetupDeviation"
			' setting To Be Installed Physical Part
			If dicSearch("AddDocumentID") <> "" Then
				If cBool(dicSearch("AddDocumentID")) Then
					objSetDev.JavaStaticText("DocumentId_DrpDwn").Click 1,1,"LEFT"
					wait 1
					objSetDev.JavaMenu("label:=Add").Select
					bFlag = Fn_SISW_ABM_SearchDialogOperations("SetDeviation_FindAndSelect", "SetDeviation", dicSearch)
					If bFlag = False Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_ABM_SetupDeviationOperations ] Failed to execute function Fn_SISW_ABM_SearchDialogOperations to find and select Physical Part")
						Exit Function
					End If
				End If
			End If
	        'clicking on OK button
			Call Fn_Button_Click("Fn_SISW_ABM_SetupDeviationOperations", objSetDev, "OK")
			Fn_SISW_ABM_SetupDeviationOperations = True
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "VerifySetupDeviation"
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_ABM_SetupDeviationOperations ] invalied case [ " & sAction & " ].")
	End Select

	If Fn_SISW_ABM_SetupDeviationOperations <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [ Fn_SISW_ABM_SetupDeviationOperations ] executed successfuly with case [ " & sAction & " ].")
	End If

	Set objShell = Nothing
	Set objSearch = Nothing
End Function
'****************************************    Function to perform  rebase physical part in As-Built Strucuture ***************************************
''Function Name		 	:	Fn_SISW_ABM_RebasePhysicalPartOperations
'
''Description		    :  	Function to perform rebase physical part in As-Built Strucuture

''Parameters		    :	 1. sAction : Action to berform   		
'							 2. sPhysicalPart : Physical Part
'							 3. sRealizedPart = Realized Part
'							 4. sRebaseTo : Rebase To text
'							 5. sRebaseDate : Rebase Date
'							 6. sStructureContextName : Structure Context Name
'							 7. dicProperties : Dictionary object to handle Properties dialog = for future use
								
''Return Value		    :  	True \ False
'
''Examples		     	:	Call Fn_SISW_ABM_RebasePhysicalPartOperations("Verify", "000063/--A", "000063/A;1-top", "", "", "000063-B", "")

'History:
'	Developer Name			Date			Rev. No.		Reviewer		Changes Done	
'------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Koustubh Watwe		26-June-2012		1.0			Koustubh W		 	Created
'------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_ABM_RebasePhysicalPartOperations(sAction, sPhysicalPart, sRealizedPart, sRebaseTo, sRebaseDate, sStructureContextName, dicProperties)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_ABM_RebasePhysicalPartOperations"
	Dim objRebase
	Set objRebase = Fn_SISW_ABM_GetObject("RebasePhysicalPart")
	Fn_SISW_ABM_RebasePhysicalPartOperations = False

	If Fn_UI_ObjectExist("Fn_SISW_ABM_RebasePhysicalPartOperations", objRebase) = False Then
		Call Fn_MenuOperation("Select","Tools:Rebase Physical Part...")
		If Fn_UI_ObjectExist("Fn_SISW_ABM_RebasePhysicalPartOperations", objRebase) = False Then
			Exit function
		End If
	End If

	Select Case sAction
		Case "Verify"
			'Physical Part
			If sPhysicalPart <> "" Then
				If objRebase.JavaObject("PhysicalPartHyperlink").Object.getText() <> sPhysicalPart Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SISW_ABM_RebasePhysicalPartOperations ] Successfully verified [ Physical Part != " & sPhysicalPart & " ]")
					Call Fn_Button_Click("Fn_SISW_ABM_RebasePhysicalPartOperations", objRebase,"Cancel")
					Exit function
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SISW_ABM_RebasePhysicalPartOperations ] Successfully verified [ Physical Part = " & sPhysicalPart & " ]")
				End If
			End If
			'Realized Part
			If sRealizedPart <> "" Then
				If objRebase.JavaObject("RealizedFromHyperlink").Object.getText() <> sRealizedPart Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SISW_ABM_RebasePhysicalPartOperations ] Successfully verified [ Realized Part != " & sRealizedPart & " ]")
					Call Fn_Button_Click("Fn_SISW_ABM_RebasePhysicalPartOperations", objRebase,"Cancel")
					Exit function
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SISW_ABM_RebasePhysicalPartOperations ] Successfully verified [ Realized Part = " & sRealizedPart & " ]")
				End If
			End If
			'Rebase To 
			If sRebaseTo <> "" Then
				If objRebase.JavaObject("RebaseToHyperlink").Object.getText() <> sRebaseTo Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SISW_ABM_RebasePhysicalPartOperations ] Successfully verified [ Rebase To != " & sRebaseTo & " ]")
					Call Fn_Button_Click("Fn_SISW_ABM_RebasePhysicalPartOperations", objRebase,"Cancel")
					Exit function
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SISW_ABM_RebasePhysicalPartOperations ] Successfully verified [ Rebase To = " & sRebaseTo & " ]")
				End If
			End If
			'Rebase Date 
			If sRebaseDate <> "" Then
				If objRebase.JavaEdit("RebaseDate").GetROProperty("value") <> sRebaseDate Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SISW_ABM_RebasePhysicalPartOperations ] Successfully verified [ Rebase Date != " & sRebaseDate & " ]")
					Call Fn_Button_Click("Fn_SISW_ABM_RebasePhysicalPartOperations", objRebase,"Cancel")
					Exit function
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SISW_ABM_RebasePhysicalPartOperations ] Successfully verified [ Rebase Date = " & sRebaseDate & " ]")
				End If
			End If
			'Structure Context Name 
			If sStructureContextName <> "" Then
				If objRebase.JavaEdit("StructureContextName").GetROProperty("value")   <> sStructureContextName Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SISW_ABM_RebasePhysicalPartOperations ] Successfully verified [ Structure Context Name != " & sStructureContextName & " ]")
					Call Fn_Button_Click("Fn_SISW_ABM_RebasePhysicalPartOperations", objRebase,"Cancel")
					Exit function
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SISW_ABM_RebasePhysicalPartOperations ] Successfully verified [ Structure Context Name = " & sStructureContextName & " ]")
				End If
			End If
			Call Fn_Button_Click("Fn_SISW_ABM_RebasePhysicalPartOperations", objRebase,"Cancel")
			Fn_SISW_ABM_RebasePhysicalPartOperations = True
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_ABM_RebasePhysicalPartOperations ] invalied case [ " & sAction & " ].")
	End Select
	If Fn_SISW_ABM_RebasePhysicalPartOperations <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SISW_ABM_RebasePhysicalPartOperations ] Executed successfully with case [ " & sAction & " ]")
	End If
	Set objRebase = Nothing
End Function
'****************************************    Function to perform Rename Physical Part ***************************************
'
''Function Name		 	:	Fn_SISW_ABM_RenamePhysicalPartOperations
'
''Description		    :  	Function to perform  Rename Physical Part

''Parameters		    :	 1. sAction : Action to berform   		
'							 2. sAsMaintainedBOMLine : Physical Part
'							 3. PartNumber = New Part Number
'							 4. NewSerialNumber : New Serial Number
'							 5. ManufacturesID : ManufacturesID
'							 6. dicProperties : Dictionary object to handle Properties dialog = for future use
								
''Return Value		    :  	True \ False
'
''Examples		     	:	Call  Fn_SISW_ABM_RenamePhysicalPartOperations("RenamePhysicalPart", "","NewPart001", "001", "", "")

'History:
'	Developer Name			Date			                   Rev. No.		Reviewer		              Changes Done	
'------------------------------------------------------------------------------------------------------------------------------------------------------------------
'    Vrushali   Wani               13-August-2012             001          Koustubh Watwe              'Created
'------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_ABM_RenamePhysicalPartOperations(sAction, sAsMaintainedBOMLine,sPartNumber, sNewSerialNumber, sManufacturesID, dicProperties)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_ABM_RenamePhysicalPartOperations"
	Dim objRenamePhysicalParts
	Set objRenamePhysicalParts = Fn_SISW_ABM_GetObject("RenamePhysicalPart")
	Fn_SISW_ABM_RenamePhysicalPartOperations = False

	If Fn_UI_ObjectExist("Fn_SISW_ABM_RenamePhysicalPartOperations", objRenamePhysicalParts) = False Then
		Call Fn_MenuOperation("Select","Tools:Rename Physical Part...")
		If Fn_UI_ObjectExist("Fn_SISW_ABM_RenamePhysicalPartOperations", objRenamePhysicalParts) = False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_ABM_RenamePhysicalPartOperations ] Failed to find [ Rename Physical Part ].")
			Exit function
		End If
	End If

	Select Case sAction
		Case "RenamePhysicalPart" 
			'New Part Number
			If sPartNumber <> "" Then
				objRenamePhysicalParts.JavaEdit("NewPartNumber").Type sPartNumber
			End If

			'New Searial Number
			If sNewSerialNumber <> "" Then
				objRenamePhysicalParts.JavaEdit("NewSerialNumber").Type sNewSerialNumber
			End If
			
			' clicking on OK button
			Call Fn_Button_Click("Fn_SISW_ABM_RenamePhysicalPartOperations", objRenamePhysicalParts, "OK")
			Fn_SISW_ABM_RenamePhysicalPartOperations = True
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_ABM_RenamePhysicalPartOperations ] invalid case [ " & sAction & " ].")
	End Select

	If  Fn_SISW_ABM_RenamePhysicalPartOperations <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [ Fn_SISW_ABM_RenamePhysicalPartOperations ] executed successfuly with case [ " & sAction & " ].")
	End If
	Set objRenamePhysicalParts = Nothing
End Function
'****************************************    Function to perform  As-Build Compare in As-Built Strucuture ***************************************
''Function Name		 : Fn_SISW_ABM_AsBuildCompareOperations
'
''Description		     : Function to perform rebase physical part in As-Built Strucuture

''Parameters		    : 	1. sAction : Action to berform   		
'							2. dicCompare : Dictionary object to handle As-Build Compare dialog
								
''Return Value		    : True \ False
'
''Examples		     :	Set dicCompare = CreateObject( "Scripting.Dictionary" )

'						dicCompare("SelectedTargetObjects") = "000119/A;1-Spec1 (View)"
'						dicCompare("SelectedSourceObjects") = "000112/A;1-Proc1 (View)"
'						dicCompare("FlipSourceAndTarget") = True / False
'						dicCompare("DisplayOptions") = True / False
'						dicCompare("FullMatch") = True
'						dicCompare("PartialMatch") = False
'						dicCompare("MultipleMatch") = False
'						dicCompare("MultiplePartialMatch") = True
'						dicCompare("MissingTarget") = True
'						dicCompare("MissingSource") = False
'						dicCompare("CompareOptions") = True / False	
'						dicCompare("TreatSameIDInContext") = True
'						dicCompare("CompareAdditionalProperties") = False
'						dicCompare("AddAvailableProperties") = "All Notes~BOM Line~Attachments"
'						dicCompare("RemoveSelectedProperties") = "Usage"
'                       dicCompare("Mode")= "Single Level" 
'					Msgbox Fn_SISW_ABM_AsBuildCompareOperations("AsBuildCompare", dicCompare)	
'History:
'	Developer Name			Date			Rev. No.		Reviewer		Changes Done	
'------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Koustubh Watwe		22-Aug-2012		1.0			Koustubh W		 	Created
'------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public function Fn_SISW_ABM_AsBuildCompareOperations(sAction, dicCompare)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_ABM_AsBuildCompareOperations"
	Dim objCompareDiag, arrProperties, iCnt
	Dim objList, iCount, sListValue
	Fn_SISW_ABM_AsBuildCompareOperations = False
	Set objCompareDiag = JavaWindow("AsBuiltManager").JavaApplet("JApplet").JavaDialog("AsBuiltCompare")

	If Fn_UI_ObjectExist("Fn_SISW_ABM_AsBuildCompareOperations",objCompareDiag) = False Then
		' perform menu operations
		Call Fn_MenuOperation("Select","Tools:Compare...")
		If Fn_UI_ObjectExist("Fn_SISW_ABM_AsBuildCompareOperations",objCompareDiag) = False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_ABM_AsBuildCompareOperations ] Failed to find [ As-Build Compare ] window.")
			Exit Function
		End If
	End If
	Select Case sAction
		Case "AsBuildCompare"
			'setting values in "Check" tab
			objCompareDiag.JavaTab("Tab").Select  "Check"

			'select from selected target objects
			If dicCompare("SelectedTargetObjects") <> "" Then
				If Fn_UI_ListItemExist("Fn_SISW_ABM_AsBuildCompareOperations",objCompareDiag, "SelectedTargetObjects", dicCompare("SelectedTargetObjects"))Then
					Call Fn_List_Select("Fn_SISW_ABM_AsBuildCompareOperations",objCompareDiag, "SelectedTargetObjects", dicCompare("SelectedTargetObjects"))
				else
					Exit function
				End If
			End If
			
			'select from selected source objects
			If dicCompare("SelectedSourceObjects") <> "" Then
				If Fn_UI_ListItemExist("Fn_SISW_ABM_AsBuildCompareOperations",objCompareDiag, "SelectedSourceObjects" , dicCompare("SelectedSourceObjects"))Then
					Call Fn_List_Select("Fn_SISW_ABM_AsBuildCompareOperations", objCompareDiag, "SelectedSourceObjects", dicCompare("SelectedSourceObjects"))
				else
					Exit function
				End If
			End If
			'click on flip button
			If dicCompare("FlipSourceAndTarget") <> "" Then
				If cBool(dicCompare("FlipSourceAndTarget")) Then
					Call Fn_Button_Click("Fn_SISW_ABM_AsBuildCompareOperations", objCompareDiag, "FlipSourceAndTarget")
				End If
			End If
			'select compare with context
			If dicCompare("CompareWithContext") <> "" Then
				If cBool(dicCompare("CompareWithContext")) Then
					Call Fn_UI_JavaRadioButton_SetON("Fn_SISW_ABM_AsBuildCompareOperations", objCompareDiag, "CompareWithContext")
				Else
					Call Fn_UI_JavaRadioButtont_setOff("Fn_SISW_ABM_AsBuildCompareOperations", objCompareDiag, "CompareWithContext")
				End If
			End If
			'select Compare without context
			If dicCompare("CompareWithoutContext") <> "" Then
				If cBool(dicCompare("CompareWithoutContext")) Then
					Call Fn_UI_JavaRadioButton_SetON("Fn_SISW_ABM_AsBuildCompareOperations", objCompareDiag, "CompareWithoutContext")
				Else
					Call Fn_UI_JavaRadioButtont_setOff("Fn_SISW_ABM_AsBuildCompareOperations", objCompareDiag, "CompareWithoutContext")
				End If
			End If

			' setting values in "Display Options" tab
			If dicCompare("DisplayOptions") <> "" Then
				If cBool(dicCompare("DisplayOptions")) Then
					objCompareDiag.JavaTab("Tab").Select  "Display Options"
					' Full Match
					If dicCompare("FullMatch") <> "" Then
						If cBool(dicCompare("FullMatch")) Then
							Call Fn_CheckBox_Set("Fn_SISW_ABM_AsBuildCompareOperations", objCompareDiag, "FullMatch", "ON")
						Else
							Call Fn_CheckBox_Set("Fn_SISW_ABM_AsBuildCompareOperations", objCompareDiag, "FullMatch", "OFF")
						End If
					End If

					' Partial Match
					If dicCompare("PartialMatch") <> "" Then
						If cBool(dicCompare("PartialMatch")) Then
							Call Fn_CheckBox_Set("Fn_SISW_ABM_AsBuildCompareOperations", objCompareDiag, "PartialMatch", "ON")
						Else
							Call Fn_CheckBox_Set("Fn_SISW_ABM_AsBuildCompareOperations", objCompareDiag, "PartialMatch", "OFF")
						End If
					End If

					' Multiple Match
					If dicCompare("MultipleMatch") <> "" Then
						If cBool(dicCompare("MultipleMatch")) Then
							Call Fn_CheckBox_Set("Fn_SISW_ABM_AsBuildCompareOperations", objCompareDiag, "MultipleMatch", "ON")
						Else
							Call Fn_CheckBox_Set("Fn_SISW_ABM_AsBuildCompareOperations", objCompareDiag, "MultipleMatch", "OFF")
						End If
					End If

					' Multiple Partial Match
					If dicCompare("MultiplePartialMatch") <> "" Then
						If cBool(dicCompare("MultiplePartialMatch")) Then
							Call Fn_CheckBox_Set("Fn_SISW_ABM_AsBuildCompareOperations", objCompareDiag, "MultiplePartialMatch", "ON")
						Else
							Call Fn_CheckBox_Set("Fn_SISW_ABM_AsBuildCompareOperations", objCompareDiag, "MultiplePartialMatch", "OFF")
						End If
					End If

					' Missing Target
					If dicCompare("MissingTarget") <> "" Then
						If cBool(dicCompare("MissingTarget")) Then
							Call Fn_CheckBox_Set("Fn_SISW_ABM_AsBuildCompareOperations", objCompareDiag, "MissingTarget", "ON")
						Else
							Call Fn_CheckBox_Set("Fn_SISW_ABM_AsBuildCompareOperations", objCompareDiag, "MissingTarget", "OFF")
						End If
					End If

					' Missing Source
					If dicCompare("MissingSource") <> "" Then
						If cBool(dicCompare("MissingSource")) Then
							Call Fn_CheckBox_Set("Fn_SISW_ABM_AsBuildCompareOperations", objCompareDiag, "MissingSource", "ON")
						Else
							Call Fn_CheckBox_Set("Fn_SISW_ABM_AsBuildCompareOperations", objCompareDiag, "MissingSource", "OFF")
						End If
					End If
				End If
			End If ' end of If dicCompare("DisplayOptions") <> "" Then

			' setting values in "Compare Options" tab
			If dicCompare("CommonOptions") <> "" Then
				If cBool(dicCompare("CommonOptions")) Then
					objCompareDiag.JavaTab("Tab").Select "Compare Options"

					' Treat Same ID In Context
					If dicCompare("TreatSameIDInContext") <> "" Then
						If cBool(dicCompare("TreatSameIDInContext")) Then
							Call Fn_CheckBox_Set("Fn_SISW_ABM_AsBuildCompareOperations", objCompareDiag, "TreatSameIDInContext", "ON")
						Else
							Call Fn_CheckBox_Set("Fn_SISW_ABM_AsBuildCompareOperations", objCompareDiag, "TreatSameIDInContext", "OFF")
						End If
					End If

					' Compare Additional Properties
					If dicCompare("CompareAdditionalProperties") <> "" Then
						If cBool(dicCompare("CompareAdditionalProperties")) Then
							Call Fn_CheckBox_Set("Fn_SISW_ABM_AsBuildCompareOperations", objCompareDiag, "CompareAdditionalProperties", "ON")
						Else
							Call Fn_CheckBox_Set("Fn_SISW_ABM_AsBuildCompareOperations", objCompareDiag, "CompareAdditionalProperties", "OFF")
						End If
					End If
					'AddAvailableProperties
					If dicCompare("AddAvailableProperties") <> "" Then
						arrProperties = split(dicCompare("AddAvailableProperties"), "~")
						For iCnt = 0 to UBound(arrProperties)
							If Fn_UI_ListItemExist("Fn_SISW_ABM_AsBuildCompareOperations",objCompareDiag, "AvailableProperties", arrProperties(iCnt))Then
								Call Fn_List_Select("Fn_SISW_ABM_AsBuildCompareOperations",objCompareDiag, "AvailableProperties", arrProperties(iCnt))
								wait 1
								Call Fn_Button_Click("Fn_SISW_ABM_AsBuildCompareOperations", objCompareDiag, "Add")
							End If
						Next
					End If

					'RemoveSelectedProperties
					If dicCompare("RemoveSelectedProperties") <> "" Then
						arrProperties = split(dicCompare("RemoveSelectedProperties"), "~")
						For iCnt = 0 to UBound(arrProperties)
							If Fn_UI_ListItemExist("Fn_SISW_ABM_AsBuildCompareOperations",objCompareDiag, "SelectedProperties", arrProperties(iCnt))Then
								Call Fn_List_Select("Fn_SISW_ABM_AsBuildCompareOperations",objCompareDiag, "SelectedProperties", arrProperties(iCnt))
								wait 1
								Call Fn_Button_Click("Fn_SISW_ABM_AsBuildCompareOperations", objCompareDiag, "Remove")
							End If
						Next
					End If
				End If

			 '  Compare Option Move when ''Compare (without context)  radio button is checked
				If  dicCompare("Mode") <> "" Then
						If   Fn_UI_ListItemExist("Fn_SISW_ABM_AsBuildCompareOperations", objCompareDiag, "Mode",dicCompare("Mode") )  Then 
								Call Fn_List_Select("Fn_SISW_ABM_AsBuildCompareOperations", objCompareDiag, "Mode",dicCompare("Mode") )
						End If
				End If
			End If ' end of If dicCompare("CommonOptions") <> "" Then
			Fn_SISW_ABM_AsBuildCompareOperations = Fn_Button_Click("Fn_SISW_ABM_AsBuildCompareOperations", objCompareDiag, "Check")

		Case Else
			' Invalid case
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_ABM_AsBuildCompareOperations ] Invalid case [ " & sAction & " ].")
	End Select

	If Fn_SISW_ABM_AsBuildCompareOperations <> "" Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SISW_ABM_AsBuildCompareOperations ] successfully executed with case [ " & sAction & " ].")
	End If
	Set objCompareDiag = Nothing
End Function
''****************************************    Function to perform Duplicate As Built Structure ***************************************

''Function Name		 	:		Fn_SISW_ABM_DuplicateAsBuiltStructure
'
''Description		    :  	    Function to perform  Generate As built  structure

''Parameters		    :	 	1. sAction : Action need to perform
'					   		   2. sRootStructureNode : Root Structure Node Path
'					   		   3. sSerialNo : Serial Number
'                             4. sLot : Lot value need to set 
'                             5. SManufactueID :  Manufacure' ID 
'                             6. sManufDate : Manufacture Date
'                             7. sInstallationTime : Installation time 
'                            8. sNoofLevel : Number of level/
'                            9.sRootStructureHeader : Header of the table need to click
'								
''Return Value		    :  		True \ False
'
''Pre-requisite		    :		As Build Manager perspective should be selected

''Examples		     	:	Fn_SISW_ABM_DuplicateAsBuiltStructure("GenerateAsBuiltStructure", "000087/A;1-Top (View)"," " ,"L1" , "000020","22-Sep-2011" , "00:00:00","3")

''History:
'						Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'						Ashwini Kumar			10-Jan-2014        1.0
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_ABM_DuplicateAsBuiltStructure(sAction,sRootStructureNode,sSerialNo,sLot,sManufactueID,sManufDate,sInstallationTime,sNoofLevel)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_ABM_DuplicateAsBuiltStructure"
	Dim objDuplicateABSDialog, bReturn , arrDateTime
	Fn_SISW_ABM_DuplicateAsBuiltStructure = False
	Set objDuplicateABSDialog = JavaWindow("AsBuiltManager").JavaWindow("DuplicateAsBuiltStructure")
	If Fn_UI_ObjectExist("Fn_SISW_ABM_DuplicateAsBuiltStructure",objDuplicateABSDialog) = False then

			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
					If sRootStructureNode <> "" then
						wait 2
						bReturn = Fn_ABM_RootStructureTableOperations("PopupSelect","", "", sRootStructureNode, "", "", "Duplicate As-Built Structure")
						If bReturn = False Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_ABM_DuplicateAsBuiltStructure ] Failed to perform [ RMB : New:Generate As-Built Structure ] on Root Strcture Node [ " & sRootStructureNode & " ].")
							Set objDuplicateABSDialog = Nothing
							Exit function
						End If
					End If
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

			If Fn_UI_ObjectExist("Fn_SISW_ABM_DuplicateAsBuiltStructure",objDuplicateABSDialog) = False then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_ABM_DuplicateAsBuiltStructure ] Failed to open Generate As Built Structure Window.")
				Set objDuplicateABSDialog = Nothing
				Exit function
			End If

	End IF

	Select Case sAction
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "DuplicateAsBuiltStructure"
			' set  Manufacture's ID
			If sManufactueID <> "" Then
                wait 3
				Call Fn_SISW_UI_JavaEdit_Operations("Fn_SISW_ABM_DuplicateAsBuiltStructure", "Type",  objDuplicateABSDialog, "ManufacturersID", sManufactueID)
			End If

			' set  Serial Number
			If sSerialNo <> "" Then
				wait 3
				Call Fn_SISW_UI_JavaEdit_Operations("Fn_SISW_ABM_DuplicateAsBuiltStructure", "Type",  objDuplicateABSDialog, "SerialNumber", sSerialNo)
                wait 3
			End If
			
			' select Checkbox  use serial number generator
			Call Fn_CheckBox_Select("Fn_SISW_ABM_DuplicateAsBuiltStructure", objDuplicateABSDialog, "UseSerialNoGenerators")


			' set  Lot Value
			If sLot <> "" Then
				Call Fn_List_Select("Fn_SISW_ABM_DuplicateAsBuiltStructure", objDuplicateABSDialog,"Lot",sLot)
			End If

			' set Manufacturing Date 
			If sManufDate <> "" Then
				If lcase(sManufDate) = "today" then
					arrDateTime = Split(Now," ")
				Else
					   arrDateTime = Split(sManufDate," ")
				End If
				objDuplicateABSDialog.JavaEdit("ManufacturingDate").Set arrDateTime(0)
				wait 1
				call Fn_KeyBoardOperation("SendKeys", "{TAB}")
				objDuplicateABSDialog.JavaStaticText("PropertyName").SetTOProperty "label", "Manufacturing Date :"
				objDuplicateABSDialog.JavaList("Time").Type arrDateTime(1)
			End If
			
            ' set  Installation Time 
			If sInstallationTime <> "" Then
				If lcase(sInstallationTime) = "today" then
					arrDateTime = Split(Now," ")
				Else
			    		arrDateTime = Split(sInstallationTime," ")
				End If
				objDuplicateABSDialog.JavaEdit("InstallationTime").Set arrDateTime(0)
				wait 1
				call Fn_KeyBoardOperation("SendKeys", "{TAB}")
				objDuplicateABSDialog.JavaStaticText("PropertyName").SetTOProperty "label", "Installation Time :"
				objDuplicateABSDialog.JavaList("Time").Type arrDateTime(1)
			End If

			' set  Number of Level 
			If sNoofLevel <> "" Then
				'to clear the field and then using type to enable  OK button
				Call Fn_SISW_UI_JavaEdit_Operations("Fn_SISW_ABM_DuplicateAsBuiltStructure", "Set",  objDuplicateABSDialog, "NumberofLevels", "")
				Call Fn_SISW_UI_JavaEdit_Operations("Fn_SISW_ABM_DuplicateAsBuiltStructure", "Type",  objDuplicateABSDialog, "NumberofLevels", sNoofLevel)
			End If
			wait 5
			'clicking on OK button
			Call Fn_Button_Click("Fn_SISW_ABM_DuplicateAsBuiltStructure", objDuplicateABSDialog, "OK")
			Fn_SISW_ABM_DuplicateAsBuiltStructure = True
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_ABM_DuplicateAsBuiltStructure ] Invalid case [ " & sAction & " ].")
			Fn_SISW_ABM_DuplicateAsBuiltStructure = False
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	End Select

	If Fn_SISW_ABM_DuplicateAsBuiltStructure = True Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SISW_ABM_DuplicateAsBuiltStructure ] executed successfully with case [ " & sAction & " ].")
	End If

	Set objDuplicateABSDialog = Nothing
End Function
