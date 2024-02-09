Option Explicit
Dim sNXUIFail
Dim SISW_NX_MICRO_TIMEOUT
SISW_NX_MICRO_TIMEOUT = 1 'time in seconds
'=====================================================================================================================
' Function List
'-------------------------------------------------------------------------------------------------------------------------------------------------------
'1. Fn_SISW_NX_UI_EditBoxOperation
'2. Fn_SISW_NX_UI_ButtonOperation
'3. Fn_SISW_NX_UI_ComboBoxOperation
'4. Fn_SISW_NX_UI_GetNodePath
'5. Fn_SISW_NX_UI_CheckBoxOperation
'6. Fn_SISW_NX_UI_WinTabOperation
'=====================================================================================================================

'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Function Name		:				Fn_SISW_NX_UI_EditBoxOperation

'Description			 :		 		 Use to Set \VerIfy\Get  Values of EditBox

''Parameters			   :	1.sFunctionName : The name of the caller function
'											2. sAction : Action to Perform (create\verIfy)
'											3. objHierarchy : Hierarchy of th WinEdit Object
'											4. sEditBoxName : Name of the WinEdit Object
'											5.	sValue: The Value which is to be Set/get/verIfy

'Return Value		   : 				True \ False

'Pre-requisite			:		 		NX Dialog should be launched

'Examples				:				 Call  Fn_SISW_NX_UI_EditBoxOperation("Fn_LaunchNX()","Set",Window("NX 8").Dialog("New"),"Name","Model123")

'History					 :		

'	Developer Name											Date						Rev. No.			Changes							Reviewer
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Pranav Ingle											07-Dec-2013					1.0																	Self
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function Fn_SISW_NX_UI_EditBoxOperation(sFunctionName,sAction,objHierarchy,sEditBoxName,sValue)

	Dim objWinEdit, EditBox_Data
	sNXUIFail = sFunctionName + "> Fn_SISW_NX_UI_EditBoxOperation : [ " & objHierarchy.toString & " ] : Action = " & sAction & " : " & sEditBoxName

 	'Set an Edit Object on variable
   Set objWinEdit= objHierarchy.WinEdit(sEditBoxName)
   'Checking the editbox  exists or not
	If objWinEdit.Exist = False Then
			Fn_SISW_NX_UI_EditBoxOperation= False
			Call Fn_UpdateLogFiles("FAIL : EditBox "&sEditBoxName&" Does Not  Exist", "FAIL: EditBox "&sEditBoxName&" Does Not  Exist")
			Exit Function
	End If
		Select Case  sAction
			Case "Set"
					If  objWinEdit.GetROProperty("enabled") = "True"  Then
						'Setting the editbox Value
						objWinEdit.Set sValue
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Text " &sValue  &"Successfully  Set in EditBox" & sEditBoxName  & " of Function " & sFunctionName)
						Fn_SISW_NX_UI_EditBoxOperation= True
					Else
						Fn_SISW_NX_UI_EditBoxOperation= False
						Call Fn_UpdateLogFiles("FAIL : "+sNXUIFail, "FAIL:"+ sNXUIFail)
				   End If
			Case "Type"
					If  objWinEdit.GetROProperty("enabled") = "True"  Then
						'Setting the editbox Value
						objWinEdit.Type sValue
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Text " &sValue  &"Successfully Typed in EditBox" & sEditBoxName  & " of Function " &sFunctionName)
						Fn_SISW_NX_UI_EditBoxOperation= True
					Else
						Fn_SISW_NX_UI_EditBoxOperation= False
						Call Fn_UpdateLogFiles("FAIL : "+sNXUIFail, "FAIL:"+ sNXUIFail)
					End If
			Case "GetValue"
					If  objWinEdit.GetROProperty("visible") = "True"  Then
						EditBox_Data=objWinEdit.GetROProperty("text")
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The retrived value of  EditBox is ["+EditBox_Data+"] ")
						Fn_SISW_NX_UI_EditBoxOperation= EditBox_Data
					Else
						Fn_SISW_NX_UI_EditBoxOperation=False
						Call Fn_UpdateLogFiles("FAIL : "+sNXUIFail, "FAIL:"+ sNXUIFail)
					End If
		  Case Else
				 Fn_SISW_NX_UI_EditBoxOperation= False 		
		End Select
	Set objWinEdit = Nothing 
End Function

'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Function Name		:				Fn_SISW_NX_UI_ButtonOperation


'Description			 :		 		 Use to Click/VerIfy  the Win Button

''Parameters			   :	 			1.sFunctionName : The Name of the caller function
'													 2. sAction : Action to perform click/verIfy
'													 3. objHierarchy : Hierarchy of th Button Object
'													4. sButtonName : Name of the Button Object
'Return Value		   : 				True \ False

'Examples				:				 Fn_SISW_NX_UI_ButtonOperation( sFunctionName, "Click", Window("NX 8").Dialog("New"), "OK")

'History					 :		

'	Developer Name											Date						Rev. No.			Changes							Reviewer
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Pranav Ingle											07-Dec-2013					1.0																	Self
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function Fn_SISW_NX_UI_ButtonOperation(sFunctionName, sAction, objHierarchy, sButtonName)
		Dim objWinButton
		sNXUIFail = sFunctionName + "> Fn_SISW_NX_UI_ButtonOperation : [ " & objHierarchy.toString & " ] : Action = " & sAction & " : " & sButtonName

		'Object Creation
		Set objWinButton = objHierarchy.WinButton(sButtonName)
		'VerIfy WinButton object exists
		If objWinButton.Exist = False Then
				Fn_SISW_NX_UI_ButtonOperation = False
				Call Fn_UpdateLogFiles("FAIL : Button "&sButtonName&" Does Not  Exist", "FAIL: Button "&sButtonName&" Does Not  Exist")
				Exit Function
		End If

		Select Case sAction
			Case "Click"
				If objWinButton.GetROProperty("enabled") = "True"  Then
					objWinButton.Click	 
					Fn_SISW_NX_UI_ButtonOperation = True
                    Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Button " &sButtonName & " Clicked Successfully " )
				Else
					Fn_SISW_NX_UI_ButtonOperation = False
					Call Fn_UpdateLogFiles("FAIL : "+sNXUIFail, "FAIL:"+ sNXUIFail)
				End If
			Case "Verify"
				If objWinButton.GetROProperty("enabled") = "True"  Then
					Fn_SISW_NX_UI_ButtonOperation = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Button " &sButtonName & " Clicked Successfully " )
				Else
					Fn_SISW_NX_UI_ButtonOperation = False
					Call Fn_UpdateLogFiles("FAIL : "+sNXUIFail, "FAIL:"+ sNXUIFail)
				End If
			Case Else
				 Fn_SISW_NX_UI_ButtonOperation= False 	
		End Select
		Set objWinButton = Nothing 
End Function

'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Function Name     	    :  Fn_SISW_NX_UI_ComboBoxOperation()
' Function Description  :  	Function used to check existence and get value of combobox
'Parameter				:   	sAction :- Action name
'										objWinDialog:- Dialog Hierachy
'										ObjCombo:- Combo box name
'										sComboName:-name of combo box
'										sValue:- Value to select from combo box
'										Reserve:- Reserved variable for future use
'78
' Return Value		    	:   True/False/value
'
' Examples		    		:   Fn_SISW_NX_UI_ComboBoxOperation("Exist",Window("NXWindow").Dialog("Flange"),"Combo","Match Face","","")

'History					 :		

'	Developer Name											Date						Rev. No.			Changes							Reviewer
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Pranav Ingle											07-Dec-2013					1.0																	Self
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_NX_UI_ComboBoxOperation(sAction,objWinDialog,ObjCombo,sComboName,sValue,Reserve)
On Error Resume Next
	 Dim bResult,objComboBox
	 Dim iEelecount,arrSelectList,iCounter,bFlag,iCnt
	 
	Set objComboBox = objWinDialog.WinComboBox(ObjCombo)

	 sNXUIFail = sFunctionName + "> Fn_SISW_NX_UI_ComboBoxOperation : [ " & objWinDialog.toString & " ] : Action = " & sAction & " : " & sComboName
	If sComboName<> "" Then
		objComboBox.SetToProperty "attached text",sComboName
	End If

	 If objComboBox.Exist = False Then
			Fn_SISW_NX_UI_ComboBoxOperation = False
			Call Fn_UpdateLogFiles("FAIL : ComboBox "& sComboName &" Does Not  Exist", "FAIL: ComboBox "& sComboName &" Does Not  Exist")
			Exit Function
	End If

	Select Case sAction
		Case "Value"
                Fn_SISW_NX_UI_ComboBoxOperation=objWinDialog.WinComboBox(ObjCombo).GetROProperty("text")
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully got value "+Fn_SISW_NX_UI_ComboBoxOperation+" of combobox "+sComboName+" in "+objWinDialog.Tostring()+" Dialog ")
		Case "Set"
				objWinDialog.WinComboBox(ObjCombo).Select sValue
				If Err.Number < 0 Then
					Fn_SISW_NX_UI_ComboBoxOperation=False
					Call Fn_UpdateLogFiles("FAIL : "+sNXUIFail+" Failed To set value ="+sValue, "FAIL:"+ sNXUIFail+" Failed To set value ="+sValue)
				Else
					Fn_SISW_NX_UI_ComboBoxOperation=True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully selected value "+sValue+" in combobox "+sComboName+" of "+objWinDialog.Tostring()+" Dialog ")
				End If	
				
		Case "Exist"
				' get total items from list
				iEelecount =objWinDialog.WinComboBox(ObjCombo).GetROProperty("items count")
				arrSelectList = split(sValue, "~")
				For iCounter = 0 To UBound(arrSelectList)
						bFlag = False
						For iCnt = 0 To iEelecount-1
							If Trim(cstr(objWinDialog.WinComboBox(ObjCombo).GetItem(iCnt))) = Trim(arrSelectList(iCounter)) Then
								bFlag=True
								Fn_SISW_NX_UI_ComboBoxOperation=True
								Exit For
							End If
						Next	
						If bFlag = False Then
							Fn_SISW_NX_UI_ComboBoxOperation=False
							Exit For
						End If				
				Next
	End Select
	Set objComboBox=Nothing
End Function

'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Function Name     	    :  Fn_SISW_NX_UI_GetNodePath()
'
' Function Description  :  	Function used to Get path (Required for Macro ) from NX Teamcenter Navigator Tree 
'
' Parameter				:   	sNodePath :- Node Path w
'										objTable:- Table Hierarchy 
'										sColName:- Column Name
'										intCnt:- 	Counter of  node in Table to Start with
'										sPath:- 		Path of Object to start with in Macro
'										sSpaceBefore:-    To Calculate length of Node to check is the correct object or not  /  Right now reserved
'										sMacroName:- 	Name of Macro to run to open Browser

' Return Value		    	:   Nav Tree Path / False
'
' Examples		    		:   Call  Fn_SISW_NX_UI_GetNodePath("Home:AutomatedTests","AssemblyTable", "Object", sSpaceBefore,3,"0 0", "NX_TCTree_Export_Browser.macro")

'History					 :		

'	Developer Name											Date						Rev. No.			Changes							Reviewer
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Pranav Ingle											07-Dec-2013					1.0																	Self
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_NX_UI_GetNodePath(sNodePath, sColName, intCnt, sPath, sSpaceBefore,sMacroName)
	Dim iStart,iColCount,sCol,iCount,aNode,bResult,sFinalPath
	Dim iRowCount,sRowContent,iSpaceLen,jCount,iLength, objTable

	'Export all the things From Teamcenter Navtree in Browser
	'---------------------------------------------------------------------------------------------------------------------------------------------
	Call Fn_SISW_MakeIEDefaultBrowser()
   'Kill Existing Excel files
   Call Fn_WindowsApplications("TerminateAll","iexplore.EXE")
   Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Close all browser instance " )
	' Run Macro to Open nav tree in Browser
	bResult=Fn_SISW_NX_Setup_LoadRunMacro("Set",Environment.Value("sPath")&"\TestData\NX\Macro\"&sMacroName)
	If bResult=False Then
		Fn_SISW_NX_UI_GetNodePath=False
		Call Fn_UpdateLogFiles("FAIL : Failed to run Macro file to export TC tree Content in IE", "FAIL : Failed to run Macro file to export TC tree Content in IE")
		Exit Function
	End If
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Succesfully run Macro file to export TC tree Content in IE" )
	wait 5
	'---------------------------------------------------------------------------------------------------------------------------------------------

    ' Check if the browser is open
	If Browser("AssemblyExportBrowser").Exist(5) = False and Browser("TeamcenterExportBrowser").Exist(5) = False  Then
		Call Fn_UpdateLogFiles("FAIL : Failed to Get Exported Assembly Structure in IE", "FAIL : Failed to Get Exported Assembly Structure in IE")
		Fn_SISW_NX_UI_GetNodePath=False
		Exit Function
	End If
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Succesfully check existance of internet explorer" )

	' Create the object of Assembly Table
	If Browser("AssemblyExportBrowser").Exist(5) = True Then
		Set objTable = Browser("AssemblyExportBrowser").Page("Page").WebTable("AssemblyTable")
	ElseIf Browser("TeamcenterExportBrowser").Exist(5) = True Then	
		Set objTable = Browser("TeamcenterExportBrowser").Page("TeamcenterPage").WebTable("TeamcenterTable")
	End If
	
	aNode=Split(sNodePath,":",-1,1)

	'Extract all the Nodes From browser
	iColCount=objTable.GetROProperty("cols")
	bResult=False
	For iCount=1 to iColCount
		sCol=objTable.GetCellData(1,iCount)
		If Trim(sCol)=sColName Then
			iCol=iCount
			bResult=True
			Exit For
		End If
	Next

	iRowCount=objTable.GetROProperty("rows")

	'Get Node path
	For iCount=2 to Ubound(aNode)
			bResult = False
			iStart=0
'					sSpaceBefore=sSpaceBefore & " "    ' To Calculate length of Node to check is the correct object or not
'					iSpaceLen=Len(sSpaceBefore)
			'iSpaceLen=iCount+1
			For jCount = intCnt To  iRowCount
					sRowContent=RTrim(objTable.GetCellData(jCount,iCol))
					If Trim(sRowContent)= Trim(aNode(iCount)) Then
						sPath = sPath& " " & iStart
						bResult = True
						jCount = jCount +1
						Exit For
'					ElseIf  Len(sRowContent) - Len(LTrim(sRowContent)) = iSpaceLen Then
					Else
						iStart = iStart + 1
					End If
			Next
			intCnt = jCount
			If bResult = False Then
				Call Fn_UpdateLogFiles("FAIL : Failed to Find "&aNode(iCount)&" in Navtree", "FAIL : Failed to Find "&aNode(iCount)&" in Navtree")
				Fn_SISW_NX_UI_GetNodePath=False
				Exit Function
			End If
	Next
	iLength = Ubound(aNode)+1
	sFinalPath=" * " & iLength & " * (" &Trim(sPath) &") ! " & aNode(Ubound(aNode))
	Fn_SISW_NX_UI_GetNodePath=sFinalPath
	
	If Browser("AssemblyExportBrowser").Exist(3) = True Then
		Browser("AssemblyExportBrowser").Close()
	ElseIf Browser("TeamcenterExportBrowser").Exist(3) = True Then	
		Browser("TeamcenterExportBrowser").Close()
	End If
	Set objTable=Nothing
End Function

'-----------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_SISW_NX_UI_CheckBoxOperation

'Description		:	This function is used to perform operations on WinCheckBox component.

'Parameters		    :	1. sFunctionName	: Caller function's name
'						2. sAction			: Action to be performed
'						3. objDialog		: Parent UI Component or WinCheckBox object
'						4. sCheckBoxName 	: WinCheckBox Control name
'						5. sValue			: value

'Return Value		: 	TRUE \ FALSE

'Pre-requisite		:	Dialog of Window which contains WinCheckBox should be opened.

'Examples			:	bReturn = Fn_SISW_NX_UI_CheckBoxOperation("Function Name", "Set", objDialog, "PropertyCheckBox", "ON")

'History			:
'-----------+--------------------+-------------+--------------------+---------------------+-----------------------------------
'	Developer Name		|	  Date		|	Rev. No.   |		Reviewer		|	Changes Done	
'-----------+--------------------+-------------+-------------------+----------------------+-----------------------------------
'	Vivek Ahirrao		|  07-Oct-2016	|	 1.0   	   |	Vivek Ahirrao	 	| 	Created
'-----------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_NX_UI_CheckBoxOperation(sFunctionName, sAction, objDialog, sCheckBoxName, sValue)
	Dim sFuncLog, objCheckBox
	Fn_SISW_NX_UI_CheckBoxOperation = False
	'Object Creation
	If sCheckBoxName <> "" Then
		Set objCheckBox = objDialog.WinCheckBox(sCheckBoxName)
		sFuncLog = sFunctionName + " > Fn_SISW_NX_UI_CheckBoxOperation  : [ " &  objDialog.toString & " ] : [ " +  objCheckBox.toString + " ] : Action = " & sAction & " : "
	Else
		Set objCheckBox = objDialog
		sFuncLog = sFunctionName + " > Fn_SISW_NX_UI_CheckBoxOperation  : [ " +  objCheckBox.toString + " ] : Action = " & sAction & " : "
	ENd IF
	
	'Verify JavaCheckBox object exists
	If objCheckBox.Exist = False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & " : FAIL : Does not exist.")
		Fn_SISW_NX_UI_CheckBoxOperation = False
		Set objCheckBox = Nothing 
		Exit Function
	End If
	
	Select case sAction
		Case "Set"
			If UCase(Trim(CStr(sValue))) = "TRUE" OR UCase(Trim(CStr(sValue))) = "ON" Then
				objCheckBox.Set "ON" 
			Else
				objCheckBox.Set "OFF" 			
			End If
			Fn_SISW_NX_UI_CheckBoxOperation = True
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & "PASS : Sucessfully set WinCheckBox [ " & sValue & " ].")
		Case Else
			'Do Nothing
	End Select
	'Clear memory of JavaCheckBox object.
	Set objCheckBox = Nothing 
End Function

'-----------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_SISW_NX_UI_WinTabOperation

'Description		:	This function to perform operations on WinTab component.

'Parameters		    :	1. sFunctionName	: Caller function's name
'						2. sAction			: Action to be performed
'						3. objDialog 		: Prent UI Component or WinTab object
'						4. sTabObjectName 	: WinTab Control name
'						5. sItem			: Tab Text

'Return Value		: 	TRUE \ FALSE

'Pre-requisite		:	WinTab must be displayed in Object repository

'Examples			:	bReturn = Fn_SISW_NX_UI_WinTabOperation("Function Name", "Exist", objDialog, "InternalTab", "Assembly")
'						bReturn = Fn_SISW_NX_UI_WinTabOperation("Function Name", "Select", objDialog, "InternalTab", "Assembly")

'History			:
'-----------+--------------------+-------------+--------------------+---------------------+-----------------------------------
'	Developer Name		|	  Date		|	Rev. No.   |		Reviewer		|	Changes Done	
'-----------+--------------------+-------------+-------------------+----------------------+-----------------------------------
'	Vivek Ahirrao		|  07-Oct-2016	|	 1.0   	   |	Vivek Ahirrao	 	| 	Created
'-----------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_NX_UI_WinTabOperation(sFunctionName, sAction, objDialog, sTabObjectName, sItem)
	Dim sFuncLog, objTab, sitems, iItemCount, iCounter
	'Object Creation
	Fn_SISW_NX_UI_WinTabOperation = False
	If sTabObjectName <> "" Then
		Set objTab = objDialog.WinTab(sTabObjectName)
		sFuncLog = sFunctionName + " > Fn_SISW_NX_UI_WinTabOperation  : [ " &  objDialog.toString & " ] : [ " +  sTabObjectName + " ] : Action = " & sAction & " : "
	Else
		Set objTab = objDialog
		sFuncLog = sFunctionName + " > Fn_SISW_NX_UI_WinTabOperation  : [ " +  objTab.toString + " ] : Action = " & sAction & " : "
	ENd IF
	
	'Verify WinTab object exists
	If objTab.Exist = False Then	
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & " : FAIL : Does not exist.")
		Set objTab = Nothing 
		Exit Function
	End If
	
	Select case sAction
		Case "Select"
			If Fn_SISW_NX_UI_WinTabOperation(sFunctionName, "Exist", objDialog, sTabObjectName, sItem) = True Then
				On Error Resume Next
				objTab.Select sItem
				If Err.Number <> 0 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & "FAIL : Failed to Find WinTab [ " & sItem + " ].")
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & "FAIL : Error Description - " & Err.Description )
                    On Error GoTo 0
				Else
					Fn_SISW_NX_UI_WinTabOperation = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & "PASS : Sucessfully Set WinTab [ " & sItem & " ].")
				End If				
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & "FAIL : Failed to Find WinTab [ " & sItem & " ].")
			End If
		Case "Exist"
			iItemCount = CInt(objTab.GetItemsCount())
			For iCounter = 0 to iItemCount - 1
				If sItem = objTab.getItem(iCounter) Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & "PASS : Sucessfully verified WinTab [ " & sItem & " ].")
					Fn_SISW_NX_UI_WinTabOperation = True
					Exit For
				End If
			Next
	End Select
	'Clear memory of WinTab object.
	Set objTab = Nothing 
End Function
