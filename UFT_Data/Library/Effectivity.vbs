Option Explicit
'*********************************************************	Function List		***********************************************************************
'1. Fn_SISW_Eff_MappingOperations
'2. Fn_SISW_Eff_DataOperations
'3. Fn_SISW_Eff_EffectivitySetDate
'4. Fn_SISW_Eff_OccurrenceEffOperations
'5. Fn_SISW_Eff_RevisionEffOperations
'*********************************************************	Function List		***********************************************************************


'****************************************    Function to  perform Effectivity Mapping Operations  ***************************************

'Function Name		      :			  Fn_SISW_Eff_MappingOperations

'Description			     :  	      Function to  perform Effectivity Mapping Operations

'Parameters			   		:	   	 	1. sModName : Perspective name 
'												  2. sAction : Action need to perform
'												  3. dicEffectivityMapping : Dictionary object of dicEffectivityMapping
'											
											
'Return Value		       : 			True\False

'Pre-requisite			    :		 	Item revision should be selected.

'Examples				    :			  	 
																	''Declaration for Effectivity Mapping Dictioanry objects 
																	''														
																	'dicEffectivityMapping("sColName") = "End Item~Sub Effectivity~Unit/Date Range"   - case Verify
																	'dicEffectivityMapping("sValue") = "000287-test1~000287-test1~1-56"   - case Verify
																	'dicEffectivityMapping("aRowNum") = "0~2~1"   - case Verify
																	'dicEffectivityMapping("bPackEffectivities") = True ' -  True / False
																	'dicEffectivityMapping("bUsedSharedEffectivity") = True ' -  True / False
																	'dicEffectivityMapping("bCreateNew") = True ' -  True / False
																	'dicEffectivityMapping("sEffectivityId") = ""
																	'dicEffectivityMapping("sEndItem") = "000287" / "Clear"
																	'dicEffectivityMapping("sEndItemSelectType") = "OpenByName" - "OpenByName", "MRUList", "PasteFromClipboard"
																	'dicEffectivityMapping("sEndItemMRU") = ""
																	'dicEffectivityMapping("sEndItemName") = "test1"
																	'dicEffectivityMapping("sEndItemRev") = ""
																	'dicEffectivityMapping("bUnit") = ""  ' -  True / False
																	'dicEffectivityMapping("sUnit") = ""
																	'dicEffectivityMapping("bDate") = True  ' -  True / False
																	'dicEffectivityMapping("sStartDates") = "24-Sep-2010~19-20-20:27-Sep-2010~19-20-20"'  - should Follow Format  DD-MMM-YYYY~HH-MM-SS     HH[00-23] seperated by :
																	'dicEffectivityMapping("sEndDates") = "26-Sep-2010~10-20-20:30-Sep-2010~10-20-20"'- should Follow Format  DD-MMM-YYYY~HH-MM-SS     HH[00-23] seperated by :
																	'dicEffectivityMapping("sSubEffectivityEndItem") = "000287"
																	'dicEffectivityMapping("sSubEffectivityEndItemSelectType") = "OpenByName" - "OpenByName", "MRUList", "PasteFromClipboard"
																	'dicEffectivityMapping("sSubEffectivityEndItemMRU") = ""
																	'dicEffectivityMapping("sSubEffectivityEndItemName") = "test1"
																	'dicEffectivityMapping("sSubEffectivityUnit") = "20"
																	'dicEffectivityMapping("sSubEffectivityDate") = ""
																	'dicEffectivityMapping("bUseLastReleaseDate") = True ' -  True / False


'													  sModName =  "My Teamcenter", "Structure  Manager", "" - if effectivity mapping dialog is already opened.
'													  sAction = "Create", "Close", "Verify", "Edit", "Delete"

'													  Call Fn_SISW_Eff_MappingOperations("My Teamcenter", "Create", dicEffectivityMapping)
'													  Call Fn_SISW_Eff_MappingOperations("Structure Manager", "Edit", dicEffectivityMapping)
'													  Call Fn_SISW_Eff_MappingOperations("Structure Manager", "Copy", dicEffectivityMapping)
'													  Call Fn_SISW_Eff_MappingOperations("", "Close", dicEffectivityMapping) 

'History:
'	Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Koustubh W	 		24-Sept-2010		 1.0														Rupali
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Koustubh W	 		23-Nov-2010		     1.0				Added Case Copy
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Koustubh W	 		24-Nov-2010		     1.0				Modified case Delete
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Koustubh W	 		13-Sept-2011		 1.0				Modified case verify
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Koustubh W	 		25-May -2012		 1.1				Modified object hierarchy
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_Eff_MappingOperations(sModName, sAction, dicEffectivityMapping)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_Eff_MappingOperations"
	Dim objEffectivityMapping, iCount, aRows
	Dim aCols, aVals, bReturn, iRows, jCount
	Dim objEff1, objEff2
				
	Fn_SISW_Eff_MappingOperations = True
	Set objEff1 = JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Effectivity Mapping_1")
	Set objEff2 = JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Effectivity Mapping_1")

        If Fn_UI_ObjectExist("Fn_SISW_Eff_MappingOperations",objEff1) = False AND Fn_UI_ObjectExist("Fn_SISW_Eff_MappingOperations",objEff2) = False Then
		Select Case sModName
			Case "My Teamcenter"
				Call Fn_MenuOperation("Select", "View:Effectivity:Effectivity Mapping...")
			
			Case "Structure Manager"
'------------------Changes are made to identify the proper QTP version and act accordingly---------------------------				
				If Environment.Value("ProductName") = sUFTProductName Then
					Call Fn_MenuOperation("WinMenuSelect", "Tools:Effectivity:Effectivity Mapping...")
				Else
    				Call Fn_MenuOperation("Select", "Tools:Effectivity:Effectivity Mapping...")
				End If
				'Call Fn_MenuOperation("Select", "Tools:Effectivity:Effectivity Mapping...")
		End Select
		wait(2)
		If Fn_UI_ObjectExist("Fn_SISW_Eff_MappingOperations",objEff1) Then
			Set objEffectivityMapping = objEff1
		ElseIf Fn_UI_ObjectExist("Fn_SISW_Eff_MappingOperations",objEff2) then
			Set objEffectivityMapping = objEff2
		Else
			Fn_SISW_Eff_MappingOperations = False
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Function [ Fn_SISW_Eff_MappingOperations ] Effectivity Dialog does not exist.")
			set objEffectivityMapping = nothing
			Exit function
		End if
	End If

	If dicEffectivityMapping("bPackEffectivities") <> "" Then
		' setting Pack Effectivity checkbox
		If dicEffectivityMapping("bPackEffectivities") = "True" OR dicEffectivityMapping("bPackEffectivities") = True Then
			Call Fn_CheckBox_Select("Fn_SISW_Eff_MappingOperations", objEffectivityMapping, "Pack effectivities")
		ElseIf dicEffectivityMapping("bPackEffectivities") = "False" OR dicEffectivityMapping("bPackEffectivities") = False Then
			Call Fn_CheckBox_Set("Fn_SISW_Eff_MappingOperations", objEffectivityMapping, "Pack effectivities","OFF")
		End If
	End If
	
	Select Case sAction
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "Create"
			Call Fn_Button_Click("Fn_SISW_Eff_MappingOperations", objEffectivityMapping,"Create...")
			Fn_SISW_Eff_MappingOperations =  Fn_SISW_Eff_DataOperations( "Create", dicEffectivityMapping)
			' closing effectivity dialog window
			Call Fn_Button_Click("Fn_SISW_Eff_MappingOperations", objEffectivityMapping,"Close")
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -						
		Case "Close" 
			' closing effectivity dialog window
			Call Fn_Button_Click("Fn_SISW_Eff_MappingOperations", objEffectivityMapping,"Close")
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -						
		Case "Verify"
			Dim aValFormat, iCnt, sDt, sModVal
			iRows = objEffectivityMapping.JavaTable("EffectivityTable").GetROProperty("rows")
			aCols = split(dicEffectivityMapping("sColName"),"~")
			aVals = split(dicEffectivityMapping("sValue"),"~")
			aRows = split(dicEffectivityMapping("aRowNum"),"~")

			'Code added by Archana to verify the value details with dt "DD-MON-YYYY"
			For  iCount =0 to uBound(aVals)
				aValFormat = Split(aVals(iCount),"to")
				For iCnt = 0 to uBound(aValFormat)
					' added if statement by Koustubh
					If instr(lcase(aValFormat(iCnt)),"am") > 0 OR instr(lcase(aValFormat(iCnt)),"pm") > 0 Then
						If isDate(aValFormat(iCnt)) Then
							aValFormat(iCnt) = Trim(aValFormat(iCnt))
							sDt = Mid(Trim(aValFormat(iCnt)),1,instr(aValFormat(iCnt),"-")-1)
							sDt = Trim(sDt)
							if Len(sDt)  = 1 then 
								sDt = "0" + sDt
								aValFormat(iCnt) = sDt + Mid(aValFormat(iCnt),instr(aValFormat(iCnt),"-"), Len(aValFormat(iCnt)))
							End if
							If Instr(aVals(iCount),"to") > 0 and iCnt <> uBound(aValFormat) then
								aValFormat(iCnt) = aValFormat(iCnt) + " to "
							End If
						End If									
					End If
				Next
				sModVal = ""
				For iCnt = 0 to UBound(aValFormat) 
					sModVal = sModVal + aValFormat(iCnt)
				Next
				aVals(iCount) = sModVal
			Next			
								
			For iCount = 0 to uBound(aRows)
				If cstr(objEffectivityMapping.JavaTable("EffectivityTable").GetCellData (cint(aRows(iCount)), aCols(iCount))) <> aVals(iCount) Then
					Fn_SISW_Eff_MappingOperations = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Function [ Fn_SISW_Eff_MappingOperations ] failed with case [ Verify ] .")
					exit for
				End If	
			Next
			' closing effectivity dialog window
			Call Fn_Button_Click("Fn_SISW_Eff_MappingOperations", objEffectivityMapping,"Close")
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -						
		Case "Edit"
			' selecting row
			If dicEffectivityMapping("aRowNum") <> "" Then
				objEffectivityMapping.JavaTable("EffectivityTable").SelectRow cInt(dicEffectivityMapping("aRowNum"))
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Function [ Fn_SISW_Eff_MappingOperations ] failed with case [ Edit ] .")
				Fn_SISW_Eff_MappingOperations = False
				Exit function
			End If
			' clicking on edit
			Call Fn_Button_Click("Fn_SISW_Eff_MappingOperations", objEffectivityMapping,"Edit...")
			Fn_SISW_Eff_MappingOperations =  Fn_SISW_Eff_DataOperations( "Edit", dicEffectivityMapping)
			' closing effectivity dialog window
			Call Fn_Button_Click("Fn_SISW_Eff_MappingOperations", objEffectivityMapping,"Close")
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -						
		Case "Delete"
			' selecting row
			If dicEffectivityMapping("aRowNum") <> "" Then
				iRows = objEffectivityMapping.JavaTable("EffectivityTable").GetROProperty("rows")
				aRows = split(dicEffectivityMapping("aRowNum"),"~")
				objEffectivityMapping.JavaTable("EffectivityTable").Object.clearSelection
				For iCount = 0 to uBound(aRows)
					Call Fn_UI_JavaTable_ExtendRow("Fn_SISW_Eff_MappingOperations", objEffectivityMapping, "EffectivityTable",cInt(trim(aRows(iCount))))		
				Next
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Function [ Fn_SISW_Eff_MappingOperations ] failed with case [ Delete ] .")
				Fn_SISW_Eff_MappingOperations = False
			End IF
			' clicking on delete
			Call Fn_Button_Click("Fn_SISW_Eff_MappingOperations", objEffectivityMapping,"Delete")
			wait(3)
			' closing effectivity dialog window
			Call Fn_Button_Click("Fn_SISW_Eff_MappingOperations", objEffectivityMapping,"Close")
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -						
		Case "Copy"
			' selecting row
			If dicEffectivityMapping("aRowNum") <> "" Then
				objEffectivityMapping.JavaTable("EffectivityTable").SelectRow cInt(dicEffectivityMapping("aRowNum"))
			Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Function [ Fn_SISW_Eff_MappingOperations ] failed with case [ Copy ] .")
					Fn_SISW_Eff_MappingOperations = False
					Exit function
			End If
			
			' clicking on edit
			Call Fn_Button_Click("Fn_SISW_Eff_MappingOperations", objEffectivityMapping,"Copy...")
			Fn_SISW_Eff_MappingOperations =  Fn_SISW_Eff_DataOperations( "Copy", dicEffectivityMapping)
			' closing effectivity dialog window
			Call Fn_Button_Click("Fn_SISW_Eff_MappingOperations", objEffectivityMapping,"Close")
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -						
		Case Else
			Fn_SISW_Eff_MappingOperations = False
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Function [ Fn_SISW_Eff_MappingOperations ] Invalid Case [ " & sAction & " ].")
	End Select
	If Fn_SISW_Eff_MappingOperations <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS : Function [ Fn_SISW_Eff_MappingOperations ] executed successfully with case [ " & sAction & " ].")
	End If
	Set objEffectivityMapping = nothing
End Function

'****************************************    Function to perform Effectivity Data Operations  ***************************************

'Function Name		      :			  Fn_SISW_Eff_DataOperations

'Description			     :  	      Function to perform Effectivity Data Operations

'Parameters			   		:	   	 	1. sAction : Action need to perform
'												  2. dicEffectivityMapping : Dictionary object of dicEffectivityMapping
'											
'Return Value		       : 			True \ False
'Pre-requisite			    :		 	 Revision of an Item should be selected.
'Examples				    :			  	 
																	''Declaration for Effectivity Mapping Dictioanry objects 
																	''														
																	'dicEffectivityMapping("sColName") = "End Item~Sub Effectivity~Unit/Date Range"
																	'dicEffectivityMapping("sValue") = "000287-test1~000287-test1~1-56"
																	'dicEffectivityMapping("aRowNum") = "0~2~1"
																	'dicEffectivityMapping("bPackEffectivities") = True ' -  True / False
																	'dicEffectivityMapping("bUsedSharedEffectivity") = True ' -  True / False
																	'dicEffectivityMapping("bCreateNew") = True ' -  True / False
																	'dicEffectivityMapping("sEffectivityId") = ""
																	'dicEffectivityMapping("sEndItem") = "000287" / "Clear"
																	'dicEffectivityMapping("sEndItemSelectType") = "OpenByName" - "OpenByName", "MRUList", "PasteFromClipboard"
																	'dicEffectivityMapping("sEndItemMRU") = ""
																	'dicEffectivityMapping("sEndItemName") = "test1"
																	'dicEffectivityMapping("sEndItemRev") = ""
																	'dicEffectivityMapping("bUnit") = ""  ' -  True / False
																	'dicEffectivityMapping("sUnit") = ""
																	'dicEffectivityMapping("bDate") = True  ' -  True / False
																	'dicEffectivityMapping("sStartDates") = "24-Sep-2010~19-20-20:27-Sep-2010~19-20-20"'  - should Follow Format  DD-MMM-YYYY~HH-MM-SS     HH[00-23] seperated by :
																	'dicEffectivityMapping("sEndDates") = "26-Sep-2010~10-20-20:30-Sep-2010~10-20-20"'- should Follow Format  DD-MMM-YYYY~HH-MM-SS     HH[00-23] seperated by :
																	'dicEffectivityMapping("sSubEffectivityEndItem") = "000287" / "Clear"
																	'dicEffectivityMapping("sSubEffectivityEndItemSelectType") = "OpenByName" - "OpenByName", "MRUList", "PasteFromClipboard"
																	'dicEffectivityMapping("sSubEffectivityEndItemMRU") = ""
																	'dicEffectivityMapping("sSubEffectivityEndItemName") = "test1"
																	'dicEffectivityMapping("sSubEffectivityUnit") = "20"
																	'dicEffectivityMapping("sSubEffectivityDate") = "" / "Nov 24, 2010 (7:10 PM)" / "Clear"
																	'dicEffectivityMapping("bUseLastReleaseDate") = True ' -  True / False

'													  sAction = "Create"  "Edit" 

'												  Call Fn_SISW_Eff_DataOperations("Create", dicEffectivityMapping)
'												  Call Fn_SISW_Eff_DataOperations("Edit", dicEffectivityMapping)

'History:
'						Developer Name			Date				Rev. No.			Changes Done								Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'					 	   Koustubh W	 		25-Sept-2010		     1.0														Rupali
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'					 	   Koustubh W	 		22-Oct-2010		         1.0			Added code for buttons SO and UP			Rupali
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'					 	   Koustubh W	 		25-May-2012		         1.1			Modified object hierarchy
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_Eff_DataOperations(sAction, dicEffectivityMapping)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_Eff_DataOperations"
	Dim objEffectivityMapping, iCount, objTable,arrDate,WshShell
	Dim aStartDate, aEndDate, bReturn, iRows, jCount
	
	Fn_SISW_Eff_DataOperations = True
	If JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Effectivity Mapping_2").Exist(5) Then
		Set objEffectivityMapping = JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Effectivity Mapping_2")
	ElseIf JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Effectivity Mapping_2").Exist(5) Then
		Set objEffectivityMapping = JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Effectivity Mapping_2")
	Else
		Fn_SISW_Eff_DataOperations = False
		Exit function
	End If

	Select Case sAction
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  
		Case "Create","Edit","Copy"
			' setting used share effectivity checkbox
			If dicEffectivityMapping("bUsedSharedEffectivity") <> ""  Then
				If  dicEffectivityMapping("bUsedSharedEffectivity") = True Then
					Call Fn_CheckBox_Set("Fn_SISW_Eff_DataOperations",objEffectivityMapping,"Use shared effectivity", "ON")
				End If
			End If

			' setting create new / edit existing checkbox
			If dicEffectivityMapping("bCreateNew") <> "" Then
				If  cBool(dicEffectivityMapping("bCreateNew")) = True Then
					If sAction ="Edit" Then
						objEffectivityMapping.JavaCheckBox("Create new").SetTOProperty "attached text" ,"Edit existing"
					Else
						objEffectivityMapping.JavaCheckBox("Create new").SetTOProperty "attached text" ,"Create new"
					End If
					Call Fn_CheckBox_Set("Fn_SISW_Eff_DataOperations",objEffectivityMapping,"Create new", "ON")
				End If
			End If
					
			' setting effectivity id
			If dicEffectivityMapping("sEffectivityId") <> "" Then
				Call Fn_Edit_Box("Fn_SISW_Eff_DataOperations",objEffectivityMapping, "Effectivity ID", dicEffectivityMapping("sEffectivityId"))
				objEffectivityMapping.JavaEdit("Effectivity ID").Activate
			End If

			'setting end item id
			If  dicEffectivityMapping("sEndItemSelectType") = "" Then
				If dicEffectivityMapping("sEndItem") <> "" Then
					If uCase(dicEffectivityMapping("sEndItem")) <> "CLEAR" Then
						Call Fn_Edit_Box("Fn_SISW_Eff_DataOperations",objEffectivityMapping, "End Item", dicEffectivityMapping("sEndItem"))
						objEffectivityMapping.JavaEdit("End Item").Activate

						'setting end item rev.
						If dicEffectivityMapping("sEndItemRev") <> "" Then
							Call Fn_List_Select("Fn_SISW_Eff_DataOperations",objEffectivityMapping, "EndItemRevId", dicEffectivityMapping("sEndItemRev"))
						End If
					Else
						' clicking on clear end item
						Call Fn_Button_Click("Fn_SISW_Eff_DataOperations", objEffectivityMapping, "ClearEndItem")
					End IF
				End If
			Else
				' future use
				Select Case dicEffectivityMapping("sEndItemSelectType")
					Case "MRUList"
						' not yet implemented
					Case "OpenByName"
						Call Fn_CheckBox_Select("Fn_SISW_Eff_DataOperations", objEffectivityMapping, "EndItemOpenByName" )
						Call Fn_OpenByNameOperations("CellDoubleClick", dicEffectivityMapping("sEndItemName"), dicEffectivityMapping("sEndItem"),"","","")
					Case "PasteFromClipboard"
						' not yet implemented
				End Select
			End If

			' setting unit radio button
			If  dicEffectivityMapping("bUnit") = True Then
				'--- insert code for removal of dates
				If cInt( Fn_UI_Object_GetROProperty("Fn_SISW_Eff_DataOperations", objEffectivityMapping.JavaRadioButton("Dates"), "enabled")) = 1Then
					If cint(Fn_UI_Object_GetROProperty("Fn_SISW_Eff_DataOperations", objEffectivityMapping.JavaRadioButton("Dates"), "value"))  = 1Then
						Set objTable = objEffectivityMapping.JavaTable("DateRangeTable")
						iRows = objTable.GetROProperty("rows")
						For iCount =0 to iRows -1
							If trim(objTable.GetCellData (iCount,"From Date")) = "" AND trim(objTable.GetCellData (iCount,"To Date")) = "" then 
								exit for
							End If
							objTable.SelectCell iCount,"From Date"
							Call Fn_Button_Click("Fn_SISW_Eff_DataOperations", objEffectivityMapping, "Clear Date")
							wait 1
							objTable.SelectCell iCount,"To Date"
							Call Fn_Button_Click("Fn_SISW_Eff_DataOperations", objEffectivityMapping, "Clear Date")
						Next
					End If
				End If

				Call Fn_UI_JavaRadioButton_SetON("Fn_SISW_Eff_DataOperations", objEffectivityMapping,"Units")
				If trim(dicEffectivityMapping("sUnit")) <> "SO" AND trim(dicEffectivityMapping("sUnit")) <> "UP" Then
					Call Fn_Edit_Box("Fn_SISW_Eff_DataOperations",objEffectivityMapping, "Units", dicEffectivityMapping("sUnit"))
				Else
					Call Fn_Button_Click("Fn_SISW_Eff_DataOperations", objEffectivityMapping, trim(dicEffectivityMapping("sUnit")))
				End If
			End If

			' setting date radio
			If  dicEffectivityMapping("bDate") = True Then
				' insert code for removal of units
				If cInt( Fn_UI_Object_GetROProperty("Fn_SISW_Eff_DataOperations", objEffectivityMapping.JavaRadioButton("Units"), "enabled")) = 1Then
					If cint(Fn_UI_Object_GetROProperty("Fn_SISW_Eff_DataOperations", objEffectivityMapping.JavaRadioButton("Units"), "value"))  = 1Then
						objEffectivityMapping.JavaEdit("Units").Set "" 
					End If
				End If
				Call Fn_UI_JavaRadioButton_SetON("Fn_SISW_Eff_DataOperations", objEffectivityMapping,"Dates")
				if sAction = "Edit" then
					Set objTable = objEffectivityMapping.JavaTable("DateRangeTable")
					iRows = objTable.GetROProperty("rows")
					For iCount =0 to iRows -1
						If trim(objTable.GetCellData (iCount,"From Date")) = "" AND trim(objTable.GetCellData (iCount,"To Date")) = "" then 
							exit for
						End If
						objTable.SelectCell iCount,"From Date"
						Call Fn_Button_Click("Fn_SISW_Eff_DataOperations", objEffectivityMapping, "Clear Date")
						wait 1
						objTable.SelectCell iCount,"To Date"
						Call Fn_Button_Click("Fn_SISW_Eff_DataOperations", objEffectivityMapping, "Clear Date")
					Next
					objTable.SelectCell 0,"From Date"
				End If

				' selecting date
				aStartDate  = split(dicEffectivityMapping("sStartDates"),":")
				aEndDate =  split(dicEffectivityMapping("sEndDates"),":")
				For iCount = 0 to uBound(aStartDate)
					Call Fn_SISW_Eff_EffectivitySetDate("Mapping", aStartDate(iCount))
					If trim(aEndDate(iCount)) <> "SO" AND trim(aEndDate(iCount)) <> "UP" Then
						Call Fn_SISW_Eff_EffectivitySetDate("Mapping", aEndDate(iCount))
					Else
						Call Fn_Button_Click("Fn_SISW_Eff_DataOperations", objEffectivityMapping, trim(aEndDate(iCount)))
					End If
				Next
			End If
					
			'setting sub effectivity end item
			If  dicEffectivityMapping("sSubEffectivityEndItemSelectType") = "" Then
				If dicEffectivityMapping("sSubEffectivityEndItem") <> "" Then
					If ucase(dicEffectivityMapping("sSubEffectivityEndItem")) <> "CLEAR" Then
						Call Fn_Edit_Box("Fn_SISW_Eff_DataOperations",objEffectivityMapping, "Sub End Item", dicEffectivityMapping("sSubEffectivityEndItem"))
						objEffectivityMapping.JavaEdit("Sub End Item").Activate
					Else
						' clicking on clear
						Call Fn_Button_Click("Fn_SISW_Eff_DataOperations", objEffectivityMapping, "ClearSubEffectivity")
					End If
				End If
			Else
				' future use
				Select Case dicEffectivityMapping("sSubEffectivityEndItemSelectType")
					Case "MRUList"
						' not yet impletemented
					Case "OpenByName"
						Call Fn_CheckBox_Select("Fn_SISW_Eff_DataOperations", objEffectivityMapping, "SubItemOpenByName" )
						Call Fn_OpenByNameOperations("CellDoubleClick",dicEffectivityMapping("sSubEffectivityEndItemName"),dicEffectivityMapping("sSubEffectivityEndItem"),"","","")
					Case "PasteFromClipboard"
						' not yet impletemented
				End Select
			End If

			If dicEffectivityMapping("sSubEffectivityUnit") <> ""  Then
				'setting sub effectivity end item unit
				Call Fn_Edit_Box("Fn_SISW_Eff_DataOperations",objEffectivityMapping, "Unit", dicEffectivityMapping("sSubEffectivityUnit"))
			End If
					
			'setting sub effectivity - use last release date
			If dicEffectivityMapping("bUseLastReleaseDate") <> ""  then
				If cBool(dicEffectivityMapping("bUseLastReleaseDate")) <> False Then
					Call Fn_CheckBox_Set("Fn_SISW_Eff_DataOperations",objEffectivityMapping,"Use last release date", "ON")
				Else
					Call Fn_CheckBox_Set("Fn_SISW_Eff_DataOperations",objEffectivityMapping,"Use last release date", "OFF")
					If dicEffectivityMapping("sSubEffectivityDate") ="Clear" Then         '' function modified to handle new date control object
						'Clear the SubEffectivity Date
						objEffectivityMapping.JavaEdit("EffectiveDate").dblclick 1,1
						wait 5
						objEffectivityMapping.JavaEdit("EffectiveDate").Set ""
						wait 1
						objEffectivityMapping.JavaEdit("EffectiveDate").Activate
						'objEffectivityMapping.JavaEdit("EffectiveDate").Type ""
						Set WshShell = CreateObject("WScript.Shell")
						WshShell.SendKeys "{TAB}"
						Set WshShell = Nothing
					Else	
						if instr(1,dicEffectivityMapping("sSubEffectivityDate"),"~") > 0 Then
							arrDate = Split(dicEffectivityMapping("sSubEffectivityDate"),"~")
						Else
							arrDate = Split(dicEffectivityMapping("sSubEffectivityDate"),"(")	
						End If	
						objEffectivityMapping.JavaEdit("EffectiveDate").dblclick 1,1
						wait 5
						Call Fn_SISW_UI_JavaEdit_Operations("Fn_SISW_Eff_DataOperations", "Set", objEffectivityMapping,"EffectiveDate", arrDate(0))
						wait 1
						objEffectivityMapping.JavaEdit("EffectiveDate").Activate
						Set WshShell = CreateObject("WScript.Shell")
						WshShell.SendKeys "{TAB}"
						Set WshShell = Nothing
						wait 2
						If Len(arrDate(1))=8 Then
							objEffectivityMapping.JavaList("EffectiveDate").Select Left(arrDate(1),8)		
						Else
							objEffectivityMapping.JavaList("EffectiveDate").Select Left(arrDate(1),9)		
						End If	
				
					End If
				End If
			End If
					
			' clicking on OK.
			Call Fn_Button_Click("Fn_SISW_Eff_DataOperations", objEffectivityMapping, "OK")
			Fn_SISW_Eff_DataOperations = True	
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  
	Case Else
		Fn_SISW_Eff_DataOperations = False
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Function [ Fn_SISW_Eff_DataOperations ] Invalid Case [ " & sAction & " ].")
	End Select
	If Fn_SISW_Eff_DataOperations <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS : Function [ Fn_SISW_Eff_DataOperations ] executed successfully with case [ " & sAction & " ].")
	End If
End Function
'*********************************************************		Fn_SISW_Eff_EffectivitySetDate		***********************************************************************
'Function Name		:				Fn_SISW_Eff_EffectivitySetDate

'Description			 :		 		 Select the Date and set to From Date And To Date Table

'Parameters			   :	 			sDateTime should Follow Format  DD-MMM-YYYY~HH-MM-SS     HH[00-23]
											
'Return Value		   : 				TRUE \ FALSE

'Pre-requisite			:		 			1.  Effectivity window should be Open 
'													 2. Radio Button Date should be Set ON
'													 3. Cell in FromDate and ToDate Table should be selected

'Examples				:				Call Fn_SISW_Eff_EffectivitySetDate("Mapping","15-Dec-2009~10-05-00") 
'												 Call Fn_SISW_Eff_EffectivitySetDate("Occurrence","15-Dec-2009~10-05-00") 
'History:
'	Developer Name			Date		  Rev. No.			Changes Done													Reviewer		Reviewed Date
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Harshal				31-Mar-2010			1.0																				Sameer 			31-Mar-2010
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Koustubh			25-Sept-2010		1.0		added from pilot function to general Functions with extra parameter, UI calls and new log calls
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Vrushali W			4-Feb-2011			1.0		added >= conidion in comparing month number.
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Ashok kakade		15-May-2012			1.0		Modified case  "ReleaseStatusEffectivity" 
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Koustubh W			15-May-2012			1.0		Modified case  "Occurrence", "LegacyOccurrenceEffectivity"
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function Fn_SISW_Eff_EffectivitySetDate(sDialogName, sDateTime)					'Note:sDateTime should Follow Format  DD-MMM-YYYY~HH-MM-SS
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_Eff_EffectivitySetDate"
	Dim arrDateTime,arrDate,arrTime,iCurMonth,iMonth,iCntMonth,iCounter, objEffectivityMapping,objTemp1,objTemp2
   'Spliting Date and Time in Seperate Arrays//////////////////////////////
	arrDateTime = Split(sDateTime,"~",-1)
	arrDate = Split(arrDateTime(0),"-",-1)
	arrTime = Split(arrDateTime(1),"-",-1)

	Select Case sDialogName
		Case "Mapping", "mapping"
			If JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Effectivity Mapping_2").Exist(5) Then
				Set objEffectivityMapping = JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Effectivity Mapping_2")
			ElseIf JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Effectivity Mapping_2").Exist(5) Then
				Set objEffectivityMapping = JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Effectivity Mapping_2")
			Else
				Fn_SISW_Eff_EffectivitySetDate = False
				Exit function
			End If
		Case "Occurrence", "occurrence"
			Set objTemp1 = JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Occurrence Effectivity_2")
			Set objTemp2 = JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Occurrence Effectivity_2")
			IF Fn_UI_ObjectExist("Fn_SISW_Eff_EffectivitySetDate", objTemp1) Then
				Set objEffectivityMapping = objTemp1
			ElseIF Fn_UI_ObjectExist("Fn_SISW_Eff_EffectivitySetDate", objTemp2) Then 
				Set objEffectivityMapping = objTemp2
			Else
				Fn_SISW_Eff_EffectivitySetDate = False
				Exit function
			End If
		Case "ReleaseStatusEffectivity"
			If JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Release Status Effectivity").Exist(5) Then
				Set objEffectivityMapping = JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Release Status Effectivity")
			ElseIf JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Release Status Effectivity").Exist(5)  Then
			  Set objEffectivityMapping = JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Release Status Effectivity")
			Else
				Fn_SISW_Eff_EffectivitySetDate = False
				Exit function
			End If

		Case "LegacyOccurrenceEffectivity"
			If JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Legacy Occurrence Effectivity").Exist(5) Then
				Set objEffectivityMapping = JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Legacy Occurrence Effectivity")
			ElseIf JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Legacy Occurrence Effectivity").Exist(5) Then
				Set objEffectivityMapping = JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Legacy Occurrence Effectivity")
			Else
				Fn_SISW_Eff_EffectivitySetDate = False
				Exit function
			End If
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Function [ Fn_SISW_Eff_EffectivitySetDate ] Invalid Dialog box.")
			Fn_SISW_Eff_EffectivitySetDate = false
			Exit function 
	End Select
	

   '////////////////////Condition Validation for date and Time//////////////////////////////////
	If  (arrDate(0) >0 And arrDate(0)<32)And(arrTime(0)>=00 And arrTime(0)<24)And(arrTime(1)>=00 And arrTime(1)<60)And(arrTime(2)>=00 And arrTime(2)<60) Then

		'Setting the current  date time
		Call Fn_Button_Click("Fn_SISW_Eff_EffectivitySetDate", objEffectivityMapping, "Clear Date")
		iCurMonth =  Month(Now)

		Select Case arrDate(1)							   'Select Case For Months
			Case "Apr", "Jun", "Sep", "Nov"
				If arrDate(0)>30 Then											
						Fn_SISW_Eff_EffectivitySetDate = False																										'Validation for Number of days in a Month
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Function [ Fn_SISW_Eff_EffectivitySetDate ] Invalid Date.")
						Exit Function
				Else
						Select Case arrDate(1)
								Case "Apr"
										iMonth = 4
								Case "Jun"
										iMonth = 6
								Case "Sep"
										iMonth = 9
								Case "Nov"
										iMonth = 11
						End Select
				End If
        Case "Jan"
			iMonth = 1
		Case "Feb"
			iMonth = 2
			If (arrDate(2) Mod 4) = 0 And arrDate(0)>29 Then																							'Validation for Number of days in a Month	
				Fn_SISW_Eff_EffectivitySetDate = False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Function [ Fn_SISW_Eff_EffectivitySetDate ] Invalid Date."& arrDate(2))
				Exit Function
			elseif (arrDate(2) Mod 4) <> 0 And arrDate(0)>28 Then
				Fn_SISW_Eff_EffectivitySetDate = False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Function [ Fn_SISW_Eff_EffectivitySetDate ] Invalid Date."& arrDate(2))
				Exit Function
			End If
		Case "Mar"
			iMonth = 3
        Case "May"
			iMonth = 5
		Case "Jul"
			iMonth = 7
		Case "Aug"
			iMonth = 8
		Case "Oct"
			iMonth = 10
		Case "Dec"
			iMonth = 12
		Case Else
			Fn_SISW_Eff_EffectivitySetDate = False
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Function [ Fn_SISW_Eff_EffectivitySetDate ] Invalid Date."& arrDate(1))														'Validation for Month
			Exit Function
	End Select
	'/////////////////////////////////////Setting The Month///////////////////////////////////////////////
			iCntMonth = iCurMonth - iMonth
			If iCntMonth>=1 Then
				For iCounter = 1 to iCntMonth 
					'objEffectivityMapping.JavaButton("PreviousMonth").Click 
					Call Fn_Button_Click("Fn_SISW_Eff_EffectivitySetDate", objEffectivityMapping, "PreviousMonth")
				Next
			elseif iCntMonth<0 then
				For iCounter = iCntMonth to -1 
					'objEffectivityMapping.JavaButton("NextMonth").Click
					Call Fn_Button_Click("Fn_SISW_Eff_EffectivitySetDate", objEffectivityMapping, "NextMonth")
				Next
			End If
	'///////////////////////Setting the values of Year, Date, Hour, Minuite and Seconds ////////////////////////////////
	'------------------Changes are made to identify the proper QTP version and act accordingly---------------------------			
		call Fn_SISW_UI_JavaEdit_Operations("Fn_SISW_Eff_EffectivitySetDate", "Set",  objEffectivityMapping.JavaEdit("Year"), "", arrDate(2))
		call Fn_SISW_UI_JavaEdit_Operations("Fn_SISW_Eff_EffectivitySetDate", "Set",  objEffectivityMapping.JavaEdit("Hour"), "", arrTime(0))
        call Fn_SISW_UI_JavaEdit_Operations("Fn_SISW_Eff_EffectivitySetDate", "Set",  objEffectivityMapping.JavaEdit("Minute"), "", arrTime(1))
        call Fn_SISW_UI_JavaEdit_Operations("Fn_SISW_Eff_EffectivitySetDate", "Set",  objEffectivityMapping.JavaEdit("Second"), "", arrTime(2))
        '----------------------End------------------------------------------------------------------------------------------
		objEffectivityMapping.JavaCheckBox("DateDigit").SetTOProperty "attached text", cStr(cInt(arrDate(0)))
		Call Fn_CheckBox_Select("Fn_SISW_Eff_EffectivitySetDate",objEffectivityMapping , "DateDigit")
		Wait(3)
        'objEffectivityMapping.JavaButton("Set Date").Click
		Call Fn_Button_Click("Fn_SISW_Eff_EffectivitySetDate", objEffectivityMapping, "Set Date")
		
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS : Function [ Fn_SISW_Eff_EffectivitySetDate ] Sucessfully Executed ")
		Fn_SISW_Eff_EffectivitySetDate =True
	Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Function [ Fn_SISW_Eff_EffectivitySetDate ] Invalid Date Or Time Value.")
			Fn_SISW_Eff_EffectivitySetDate = False
	End If
End Function


'****************************************    Function to perform Occurrence Effectivity Operations  ***************************************

'Function Name		      :			  Fn_SISW_Eff_OccurrenceEffOperations

'Description			     :  	      Function to perform Occurrence Effectivity Operations

'Parameters			   		:	   	 	1. sAction : Action need to perform
'										  2. dicEffectivityMapping : Dictionary object of dicEffectivityMapping
'											
'Return Value		       : 			True \ False
'Pre-requisite			    :		 	 Revision of an Item should be selected.
																	'Need to set following preference:  CFMOccEffMode = maintenance
'Examples				    :			  Otherwise It will open in Legacy mode ( CFMOccEffMode = legacy )
																	''Declaration for Effectivity Mapping Dictioanry objects 
																					'dicEffectivityMapping("sColName") = ""
																					'dicEffectivityMapping("sValue") = ""
																					'dicEffectivityMapping("aRowNum") = "1~2"
																					'dicEffectivityMapping("bPackEffectivities") = True ' -  True / False
																					'dicEffectivityMapping("bUsedSharedEffectivity") = True ' -  True / False
																					'dicEffectivityMapping("bCreateNew") = True ' -  True / False
													'								dicEffectivityMapping("sVerifyEffectivityId") = "re2~res3"
													'								dicEffectivityMapping("sSelectEffectivityId") = "re2"
													'								dicEffectivityMapping("sEffectivityId") = "d*"																					
																					'dicEffectivityMapping("bEffectivityProtection") = ""
																					'dicEffectivityMapping("sEndItem") = "000124" / "Clear"
																					'dicEffectivityMapping("sEndItemSelectType") = "OpenByName" '- "OpenByName", "MRUList", "PasteFromClipboard"
																					'dicEffectivityMapping("sEndItemMRU") = ""
																					'dicEffectivityMapping("sEndItemName") = "Top"
																					'dicEffectivityMapping("sEndItemRev") = ""
																					'dicEffectivityMapping("bUnit") = ""  ' -  True / False
																					'dicEffectivityMapping("sUnit") = ""
																					'dicEffectivityMapping("bDate") = True  ' -  True / False
																					'dicEffectivityMapping("sStartDates") = "24-Oct-2010~19-20-20:27-Oct-2010~19-20-20"'  - should Follow Format  DD-MMM-YYYY~HH-MM-SS     HH[00-23] seperated by :
'  																   								  "Oct 26, 2010  (12:00 AM)"'  - For Case "VerifyRevisionEffectivityDetails" - should Follow Format  MMM DD, YYYY  (HH:MM AM/PM)  seperated by ~
																					'dicEffectivityMapping("sEndDates") = "26-Oct-2010~10-20-20:30-Oct-2010~10-20-20"'- should Follow Format  DD-MMM-YYYY~HH-MM-SS     HH[00-23] seperated by :
'																  								 "Oct 29, 2010  (12:00 AM)"'  - For Case "VerifyRevisionEffectivityDetails" - should Follow Format  MMM DD, YYYY  (HH:MM AM/PM)  seperated by ~

'													  sAction = "Create"  "Edit" , "CreateOnMultipleBOMLines", "Copy", "Delete", "VerifyOccEffectivityData", "VerifyOccEffectivityCreate", "VerifyOccEffectivityCreateOnMultipleBOMLines"

'												  Call Fn_SISW_Eff_OccurrenceEffOperations("Create", dicEffectivityMapping)
'												  Call Fn_SISW_Eff_OccurrenceEffOperations("CreateOnMultipleBOMLines", dicEffectivityMapping)
'												  Call Fn_SISW_Eff_OccurrenceEffOperations("Edit", dicEffectivityMapping)
'												  Call Fn_SISW_Eff_OccurrenceEffOperations("Copy", dicEffectivityMapping)
'												  Call Fn_SISW_Eff_OccurrenceEffOperations("Delete", dicEffectivityMapping)
'												  Call Fn_SISW_Eff_OccurrenceEffOperations("VerifyOccEffectivityData", dicEffectivityMapping)
'												  Call Fn_SISW_Eff_OccurrenceEffOperations("VerifyOccEffectivityCreate", dicEffectivityMapping)
'												  Call Fn_SISW_Eff_OccurrenceEffOperations("VerifyOccEffectivityCreateOnMultipleBOMLines", dicEffectivityMapping)
'												  Call Fn_SISW_Eff_OccurrenceEffOperations("LegacyOccurrenceEffectivity", dicEffectivityMapping)
'												  Call Fn_SISW_Eff_OccurrenceEffOperations("RemoveLegacyOccurrenceEffectivity", dicEffectivityMapping)
'												  Call Fn_SISW_Eff_OccurrenceEffOperations("VerifyLegacyOccurrenceEffectivity", dicEffectivityMapping)



'History:
'						Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'					 	   Koustubh W	 		6-Oct-2010		    1.0										    Rupali
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'					 	   Koustubh W	 		11-Oct-2010		    1.0			Added cases VerifyOccEffectivityCreate, VerifyOccEffectivityCreateOnMultipleBOMLines
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'					 	   Koustubh W	 		22-Oct-2010		    1.0			Added code for buttons SO and UP			Rupali
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'					 	   Koustubh W	 		22-Oct-2010			1.0			Added cases "LegacyOccurrenceEffectivity", 
'																									"RemoveLegacyOccurrenceEffectivity"	
'																									"VerifyLegacyOccurrenceEffectivity"
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'					 	   Koustubh W	 		25-May-2012			2.0			Modifeid function according to TC10.0 changes in OR.
'																				Copied OccurrenceEffectivityDialog, VariantItemRevise from StructureManager.tsr to General.tsr
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'					 	   Naveen Gupta	 		11-Feb-2013		    2.0			Modified code to set End Item,
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_Eff_OccurrenceEffOperations(sAction, dicEffectivityMapping)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_Eff_OccurrenceEffOperations"
	Dim bReturn, objEditOccEffDiag, objCreateOccEffDiag, aRRsEndItem
	Dim aStartDate, aEndDate, iRows, jCount, iCount, aRows
	Dim aCols, aVals
	Dim objTable
	Dim objOccEffID, aEffID, bFlag, iCnt
	Dim objTempOcc1, objTempOcc2, objTempOccDetails1, objTempOccDetails2

	Set objTempOcc1 = JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Occurrence Effectivity_1")
	Set objTempOcc2 = JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Occurrence Effectivity_1")
	'Set objEditOccEffDiag = JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Occurrence Effectivity_1")
	
	Set objTempOccDetails1 = JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Occurrence Effectivity_2")
	Set objTempOccDetails2 = JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Occurrence Effectivity_2")
	'Set objCreateOccEffDiag = JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Occurrence Effectivity_2")
	Fn_SISW_Eff_OccurrenceEffOperations = True

	Select Case sAction
		Case "Create", "Edit", "Copy", "VerifyOccEffectivityCreate", "VerifyOccEffectivityEdit", "VerifyOccEffectivityData", "Delete"
			' Occ1
			IF Fn_UI_ObjectExist("Fn_SISW_Eff_OccurrenceEffOperations", objTempOcc1) Then
				Set objEditOccEffDiag = objTempOcc1
			ElseIF Fn_UI_ObjectExist("Fn_SISW_Eff_OccurrenceEffOperations", objTempOcc2) Then 
				Set objEditOccEffDiag = objTempOcc2
			ElseIF Fn_UI_ObjectExist("Fn_SISW_Eff_OccurrenceEffOperations", objTempOccDetails1) Then 
				Set objEditOccEffDiag = objTempOccDetails1
			ElseIF Fn_UI_ObjectExist("Fn_SISW_Eff_OccurrenceEffOperations", objTempOccDetails2) Then 
				Set objEditOccEffDiag = objTempOccDetails2				
			Else
				'------------------Changes are made to identify the proper QTP version and act accordingly---------------------------				
				If Environment.Value("ProductName") = sUFTProductName Then
					Call Fn_MenuOperation("WinMenuSelect","Tools:Effectivity:Occurrence Effectivity...:View, Edit and Create...")
			   Else
			    	Call Fn_MenuOperation("Select","Tools:Effectivity:Occurrence Effectivity...:View, Edit and Create...")
			    End If
				'Call Fn_MenuOperation("Select","Tools:Effectivity:Occurrence Effectivity...:View, Edit and Create...")
				Call Fn_ReadyStatusSync(2)
				IF Fn_UI_ObjectExist("Fn_SISW_Eff_OccurrenceEffOperations", objTempOcc1) Then
					Set objEditOccEffDiag = objTempOcc1
				ElseIF Fn_UI_ObjectExist("Fn_SISW_Eff_OccurrenceEffOperations", objTempOcc2) Then 
					Set objEditOccEffDiag = objTempOcc2
				ElseIF Fn_UI_ObjectExist("Fn_SISW_Eff_OccurrenceEffOperations", objTempOccDetails1) Then 
					Set objEditOccEffDiag = objTempOccDetails1
				ElseIF Fn_UI_ObjectExist("Fn_SISW_Eff_OccurrenceEffOperations", objTempOccDetails2) Then 
					Set objEditOccEffDiag = objTempOccDetails2					
				Else
					Fn_SISW_Eff_OccurrenceEffOperations = False
					Exit Function
				End If
			End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "CreateOnMultipleBOMLines",  "VerifyOccEffectivityCreateOnMultipleBOMLines"
			' Occ2
			IF Fn_UI_ObjectExist("Fn_SISW_Eff_OccurrenceEffOperations", objTempOccDetails1) Then
				Set objCreateOccEffDiag = objTempOccDetails1
			ElseIF Fn_UI_ObjectExist("Fn_SISW_Eff_OccurrenceEffOperations", objTempOccDetails2) Then 
				Set objCreateOccEffDiag = objTempOccDetails2
			Else
				'------------------Changes are made to identify the proper QTP version and act accordingly---------------------------				
				If Environment.Value("ProductName") = sUFTProductName Then
					Call Fn_MenuOperation("WinMenuSelect","Tools:Effectivity:Occurrence Effectivity...:Create on Multiple BOM Lines...")
			    Else
			    	Call Fn_MenuOperation("Select","Tools:Effectivity:Occurrence Effectivity...:Create on Multiple BOM Lines...")
			    End If
'				Call Fn_MenuOperation("Select","Tools:Effectivity:Occurrence Effectivity...:Create on Multiple BOM Lines...")
				Call Fn_ReadyStatusSync(2)
				IF Fn_UI_ObjectExist("Fn_SISW_Eff_OccurrenceEffOperations", objTempOccDetails1) Then
					Set objCreateOccEffDiag = objTempOccDetails1
				ElseIF Fn_UI_ObjectExist("Fn_SISW_Eff_OccurrenceEffOperations", objTempOccDetails2) Then 
					Set objCreateOccEffDiag = objTempOccDetails2
				Else
					Fn_SISW_Eff_OccurrenceEffOperations = False
					exit function
				End IF
			End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "LegacyOccurrenceEffectivity", "RemoveLegacyOccurrenceEffectivity","VerifyLegacyOccurrenceEffectivity"
			' legOccEff
			If JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Legacy Occurrence Effectivity").Exist(5) Then
				Set objCreateOccEffDiag = JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Legacy Occurrence Effectivity")
			ElseIf JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Legacy Occurrence Effectivity").Exist(5) Then
				Set objCreateOccEffDiag = JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Legacy Occurrence Effectivity")
			Else
			'------------------Changes are made to identify the proper QTP version and act accordingly---------------------------								
				If Environment.Value("ProductName") = sUFTProductName Then
					Call Fn_MenuOperation("WinMenuSelect","Tools:Effectivity:Occurrence Effectivity...:View, Edit and Create...")
			    Else
			    	Call Fn_MenuOperation("Select","Tools:Effectivity:Occurrence Effectivity...:View, Edit and Create...")
			    End If
				Call Fn_ReadyStatusSync(2)
				If JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Legacy Occurrence Effectivity").Exist(5) Then
					Set objCreateOccEffDiag = JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Legacy Occurrence Effectivity")
				ElseIf JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Legacy Occurrence Effectivity").Exist(5) Then
					Set objCreateOccEffDiag = JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Legacy Occurrence Effectivity")
				Else
					Fn_SISW_Eff_OccurrenceEffOperations = False
					Exit Function
				End If
			End If
	End Select
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	Select Case sAction
		Case "Create", "Edit", "Copy", "CreateOnMultipleBOMLines", "LegacyOccurrenceEffectivity"
			Select Case sAction
				Case "Create", "Edit","Copy"
					' setting Pack Effectivity checkbox
					If dicEffectivityMapping("bPackEffectivities") <> "" then
						If cBool(dicEffectivityMapping("bPackEffectivities")) = True Then
							Call Fn_CheckBox_Select("Fn_SISW_Eff_OccurrenceEffOperations", objEditOccEffDiag, "Pack effectivities")
						ElseIf cBool(dicEffectivityMapping("bPackEffectivities")) = False Then
							Call Fn_CheckBox_Set("Fn_SISW_Eff_OccurrenceEffOperations", objEditOccEffDiag, "Pack effectivities","OFF")
						End If
					End If
					' selecting row
					If dicEffectivityMapping("aRowNum") <> "" Then
						Call Fn_UI_JavaTable_SelectRow("Fn_SISW_Eff_OccurrenceEffOperations", objEditOccEffDiag, "EffectivityTable", cInt(dicEffectivityMapping("aRowNum")))
					End If
					' clicking on create button
					Call Fn_Button_Click("Fn_SISW_Eff_OccurrenceEffOperations", objEditOccEffDiag, sAction & "...")

					IF Fn_UI_ObjectExist("Fn_SISW_Eff_OccurrenceEffOperations", objTempOccDetails1) Then
						Set objCreateOccEffDiag = objTempOccDetails1
					ElseIF Fn_UI_ObjectExist("Fn_SISW_Eff_OccurrenceEffOperations", objTempOccDetails2) Then 
						Set objCreateOccEffDiag = objTempOccDetails2
					End If

				Case Else
					' do nothing
			End Select

			' setting used share effectivity checkbox
			If dicEffectivityMapping("bUsedSharedEffectivity") <> ""  Then
				If Cbool(dicEffectivityMapping("bUsedSharedEffectivity")) Then
					Call Fn_CheckBox_Set("Fn_SISW_Eff_OccurrenceEffOperations",objCreateOccEffDiag,"Use shared effectivity", "ON")
				End If
			End If

			' setting create new / edit existing checkbox
			If dicEffectivityMapping("bCreateNew") <> ""   Then
				If  cBool(dicEffectivityMapping("bCreateNew")) Then
					If sAction ="Edit" Then
						objCreateOccEffDiag.JavaCheckBox("Create new").SetTOProperty "attached text" ,"Edit existing"
					Else
						objCreateOccEffDiag.JavaCheckBox("Create new").SetTOProperty "attached text" ,"Create new"
					End If
					Call Fn_CheckBox_Set("Fn_SISW_Eff_OccurrenceEffOperations",objCreateOccEffDiag,"Create new", "ON")
				End If
			End If
			
			' setting effectivity id
			If dicEffectivityMapping("sEffectivityId") <> ""  Then
				objCreateOccEffDiag.JavaEdit("EffectivityID").Set ""
				objCreateOccEffDiag.JavaEdit("EffectivityID").Type dicEffectivityMapping("sEffectivityId")
			'----------------------End------------------------------------------------------------------------------------------
'				Call Fn_Edit_Box("Fn_SISW_Eff_OccurrenceEffOperations",objCreateOccEffDiag, "EffectivityID", dicEffectivityMapping("sEffectivityId"))
				if dicEffectivityMapping("bCreateNew") = "" Then
					objCreateOccEffDiag.JavaEdit("EffectivityID").Activate
				ElseIf  cBool(dicEffectivityMapping("bCreateNew") ) = False Then
					objCreateOccEffDiag.JavaEdit("EffectivityID").Activate
				End If
				If instr(dicEffectivityMapping("sEffectivityId"), "*") > 0 Then
					If JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("OccurrenceEffectivityDialog").Exist(15) Then
						Set objOccEffID = JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("OccurrenceEffectivityDialog")
					ElseIf JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("OccurrenceEffectivityDialog").Exist(15) Then
						Set objOccEffID = JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("OccurrenceEffectivityDialog")
					End If
					if objOccEffID.exist(5) then
						if dicEffectivityMapping("sSelectEffectivityId") <> "" then
							'select specified ID from list
							iRows = cInt(objOccEffID.JavaTable("EffectivityIDTable").getROProperty("rows"))
							bFlag = False
							for iCnt = 0 to iRows - 1
								if cstr(objOccEffID.JavaTable("EffectivityIDTable").getCellData(iCnt, "ID")) = dicEffectivityMapping("sSelectEffectivityId") Then
									objOccEffID.JavaTable("EffectivityIDTable").SelectRow iCnt
									bFlag = True
									exit for
								End If
							Next
							if  bFlag <> True then
								Call Fn_Button_Click("Fn_SISW_Eff_OccurrenceEffOperations", objOccEffID,"Close")
								Fn_SISW_Eff_OccurrenceEffOperations = False
								Set objOccEffID = nothing
								Exit function
							End If
							Call Fn_Button_Click("Fn_SISW_Eff_OccurrenceEffOperations", objOccEffID,"OK")
							Else
							Call Fn_Button_Click("Fn_SISW_Eff_OccurrenceEffOperations", objOccEffID,"Close")
						End IF
					End If ' end of if objOccEffID.exist(20) then
				End IF
			End If ' end of If dicEffectivityMapping("sEffectivityId") <> ""  Then

			'Effectivity Protection
			If dicEffectivityMapping("bEffectivityProtection") <> ""  Then
				If   dicEffectivityMapping("bEffectivityProtection") = True Then
					Call Fn_CheckBox_Set("Fn_SISW_Eff_OccurrenceEffOperations",objCreateOccEffDiag,"Apply Access Manager effectivi", "ON")
				Else
					Call Fn_CheckBox_Set("Fn_SISW_Eff_OccurrenceEffOperations",objCreateOccEffDiag,"Apply Access Manager effectivi", "OFF")
				End If
			End If
			'setting end item id
			If  dicEffectivityMapping("sEndItemSelectType") = "" Then
				If dicEffectivityMapping("sEndItem") <> "" Then
					If ucase(dicEffectivityMapping("sEndItem")) <> "CLEAR" Then
						'Call Fn_CheckBox_Select("Fn_SISW_Eff_OccurrenceEffOperations", objCreateOccEffDiag, "EndItemOpenByName" )
						'aRRsEndItem = Split(dicEffectivityMapping("sEndItem"), "-")
						'Call Fn_OpenByNameOperations("CellDoubleClick", aRRsEndItem(1), aRRsEndItem(0),"","","")
						'Call Fn_Edit_Box("Fn_SISW_Eff_OccurrenceEffOperations",objCreateOccEffDiag, "EndItem", dicEffectivityMapping("sEndItem"))
						objCreateOccEffDiag.JavaEdit("EndItem").Activate
						wait 2
						objCreateOccEffDiag.JavaEdit("EndItem").Set ""
						objCreateOccEffDiag.JavaEdit("EndItem").Type dicEffectivityMapping("sEndItem")
						objCreateOccEffDiag.JavaEdit("EndItem").Activate
						wait 2
						'setting end item rev.
						If dicEffectivityMapping("sEndItemRev") <> "" Then
							Call Fn_List_Select("Fn_SISW_Eff_OccurrenceEffOperations",objCreateOccEffDiag, "EndItemRevId", dicEffectivityMapping("sEndItemRev"))
						End If
					Else
						' click on clear
						Call Fn_Button_Click("Fn_SISW_Eff_OccurrenceEffOperations", objCreateOccEffDiag,"ClearEndItem")
					End If
				End If
			Else
				' future use
				Select Case dicEffectivityMapping("sEndItemSelectType")
					Case "MRUList"
						' not yet implemented
					Case "OpenByName"
						Call Fn_CheckBox_Select("Fn_SISW_Eff_OccurrenceEffOperations", objCreateOccEffDiag, "EndItemOpenByName" )
						Call Fn_OpenByNameOperations("CellDoubleClick", dicEffectivityMapping("sEndItemName"), dicEffectivityMapping("sEndItem"),"","","")
					Case "PasteFromClipboard"
						' not yet implemented
				End Select
			End If

			' setting unit radio button
			If  dicEffectivityMapping("bUnit") <> ""  Then
				If Cbool(dicEffectivityMapping("bUnit")) = True Then
					'--- insert code for removal of dates
					If cInt( Fn_UI_Object_GetROProperty("Fn_SISW_Eff_OccurrenceEffOperations", objCreateOccEffDiag.JavaRadioButton("Dates"), "enabled")) = 1Then
						If cint(Fn_UI_Object_GetROProperty("Fn_SISW_Eff_OccurrenceEffOperations", objCreateOccEffDiag.JavaRadioButton("Dates"), "value"))  = 1Then
							Set objTable = objCreateOccEffDiag.JavaTable("DateRangeTable")
							iRows = objTable.GetROProperty("rows")
							For iCount =0 to iRows -1
								If trim(objTable.GetCellData (iCount,"From Date")) = "" AND trim(objTable.GetCellData (iCount,"To Date")) = "" then exit for
								objTable.SelectCell iCount,"From Date"
								Call Fn_Button_Click("Fn_SISW_Eff_OccurrenceEffOperations", objCreateOccEffDiag, "Clear Date")
								wait 2
								objTable.SelectCell iCount,"To Date"
								Call Fn_Button_Click("Fn_SISW_Eff_OccurrenceEffOperations", objCreateOccEffDiag, "Clear Date")
							Next
						End If
					End If

					'setting units
					Call Fn_UI_JavaRadioButton_SetON("Fn_SISW_Eff_OccurrenceEffOperations", objCreateOccEffDiag,"Units")
					If trim(dicEffectivityMapping("sUnit")) <> "SO" AND trim(dicEffectivityMapping("sUnit")) <> "UP" Then
						Call Fn_Edit_Box("Fn_SISW_Eff_OccurrenceEffOperations",objCreateOccEffDiag, "Units", dicEffectivityMapping("sUnit"))
						wait 2
					Else
						Call Fn_Button_Click("Fn_SISW_Eff_OccurrenceEffOperations", objCreateOccEffDiag, trim(dicEffectivityMapping("sUnit")))
						wait 2
					End If
				End If
			End If
			' setting date radio
			If  dicEffectivityMapping("bDate") <>""  Then
				If  cBool(dicEffectivityMapping("bDate") ) = True Then
					' insert code for removal of units
					If cInt( Fn_UI_Object_GetROProperty("Fn_SISW_Eff_OccurrenceEffOperations", objCreateOccEffDiag.JavaRadioButton("Units"), "enabled")) = 1Then
						If cint(Fn_UI_Object_GetROProperty("Fn_SISW_Eff_OccurrenceEffOperations", objCreateOccEffDiag.JavaRadioButton("Units"), "value"))  = 1Then
							objCreateOccEffDiag.JavaEdit("Units").Set "" 
						End If
					End If
					Call Fn_UI_JavaRadioButton_SetON("Fn_SISW_Eff_OccurrenceEffOperations", objCreateOccEffDiag,"Dates")
					if sAction = "Edit" then
						Set objTable = objCreateOccEffDiag.JavaTable("DateRangeTable")
						iRows = objTable.GetROProperty("rows")
						For iCount =0 to iRows -1
							If trim(objTable.GetCellData (iCount,"From Date")) = "" AND trim(objTable.GetCellData (iCount,"To Date")) = "" then exit for
							objTable.SelectCell iCount,"From Date"
							Call Fn_Button_Click("Fn_SISW_Eff_OccurrenceEffOperations", objCreateOccEffDiag, "Clear Date")
							wait 1
							objTable.SelectCell iCount,"To Date"
							Call Fn_Button_Click("Fn_SISW_Eff_OccurrenceEffOperations", objCreateOccEffDiag, "Clear Date")
						Next
						objTable.SelectCell 0,"From Date"
					End If
					' selecting date
					aStartDate  = split(dicEffectivityMapping("sStartDates"),":")
					aEndDate =  split(dicEffectivityMapping("sEndDates"),":")

					For iCount = 0 to uBound(aStartDate)
						If sAction = "LegacyOccurrenceEffectivity" Then
							Call Fn_EffectivitySetDate("LegacyOccurrenceEffectivity", aStartDate(iCount))
						Else
							Call Fn_EffectivitySetDate("Occurrence", aStartDate(iCount))
						End If
						
						
						If trim(aEndDate(iCount)) <> "SO" AND trim(aEndDate(iCount)) <> "UP" Then
							If sAction = "LegacyOccurrenceEffectivity" Then
								Call Fn_EffectivitySetDate("LegacyOccurrenceEffectivity", aStartDate(iCount))
							Else
								Call Fn_EffectivitySetDate("Occurrence", aEndDate(iCount))
							End If
						Else
							Call Fn_Button_Click("Fn_SISW_Eff_OccurrenceEffOperations", objCreateOccEffDiag, trim(aEndDate(iCount)))
						End If
					Next
				End If
			Else
				If sAction =  "LegacyOccurrenceEffectivity" Then
					If dicEffectivityMapping("sStartDates") <> "" AND  dicEffectivityMapping("sEndDates") <> "" Then
						' selecting date
						aStartDate  = split(dicEffectivityMapping("sStartDates"),":")
						aEndDate =  split(dicEffectivityMapping("sEndDates"),":")

						For iCount = 0 to uBound(aStartDate)
							Call Fn_EffectivitySetDate("LegacyOccurrenceEffectivity", aStartDate(iCount))
							Call Fn_EffectivitySetDate("LegacyOccurrenceEffectivity", aEndDate(iCount))
						Next
					End If
				End If
			End If
		
			'clciking on OK
			Call Fn_Button_Click("Fn_SISW_Eff_OccurrenceEffOperations",objCreateOccEffDiag,"OK")
			' addding code for upgrade
			JavaDialog("VariantItemRevise").setTOProperty "title","Upgrade Legacy Occurrence Effectivity?"
			If JavaDialog("VariantItemRevise").exist(10) Then
				Call Fn_Button_Click("Fn_SISW_Eff_OccurrenceEffOperations",JavaDialog("VariantItemRevise"),"Yes")
			End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
                Case "Delete"
                        ' selecting row
			If dicEffectivityMapping("aRowNum") <> "" Then
				aRows = split(dicEffectivityMapping("aRowNum"), "~")
				For iCount = 0 to uBound(aRows )
					Call Fn_UI_JavaTable_ExtendRow("Fn_SISW_Eff_OccurrenceEffOperations", objEditOccEffDiag, "EffectivityTable", cInt(dicEffectivityMapping("aRowNum")))
					'objEditOccEffDiag.JavaTable("EffectivityTable").ExtendRow  cInt(aRows(iCount))
				Next
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Function [ Fn_SISW_Eff_OccurrenceEffOperations ] failed with case [ Delete ].")
				Fn_SISW_Eff_OccurrenceEffOperations = False
				Set objEditOccEffDiag = nothing
				Set objCreateOccEffDiag = nothing
				Exit function
			End If
			' clicking on Delete
			Call Fn_Button_Click("Fn_SISW_Eff_OccurrenceEffOperations", objEditOccEffDiag,"Delete")
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "RemoveLegacyOccurrenceEffectivity"
			' clickin on remove
			Call Fn_Button_Click("Fn_SISW_Eff_OccurrenceEffOperations", objEditOccEffDiag,"Remove")
			' clickin on Ok
			Call Fn_Button_Click("Fn_SISW_Eff_OccurrenceEffOperations", objEditOccEffDiag,"OK")
			Fn_SISW_Eff_OccurrenceEffOperations = True
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "VerifyOccEffectivityData"
			iRows = objEditOccEffDiag.JavaTable("EffectivityTable").GetROProperty("rows")
			aCols = split(dicEffectivityMapping("sColName"),"~")
			aVals = split(dicEffectivityMapping("sValue"),"~")
			aRows = split(dicEffectivityMapping("aRowNum"),"~")

			For iCount = 0 to uBound(aRows)
				If cstr(objEditOccEffDiag.JavaTable("EffectivityTable").GetCellData (cint(aRows(iCount)), aCols(iCount))) <> aVals(iCount) Then
					Fn_SISW_Eff_OccurrenceEffOperations = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Function [ Fn_SISW_Eff_OccurrenceEffOperations ] failed with case [ Verify ] .")
					exit for
				End If	
			Next
			' closing effectivity dialog window
			Call Fn_Button_Click("Fn_SISW_Eff_OccurrenceEffOperations", objEditOccEffDiag,"Close")
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "VerifyOccEffectivityCreate","VerifyOccEffectivityEdit",  "VerifyOccEffectivityCreateOnMultipleBOMLines","VerifyLegacyOccurrenceEffectivity"
			Select Case sAction
				'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				Case "VerifyOccEffectivityCreate"
					' setting Pack Effectivity checkbox
					If dicEffectivityMapping("bPackEffectivities") = "True" OR dicEffectivityMapping("bPackEffectivities") = True Then
						Call Fn_CheckBox_Select("Fn_SISW_Eff_OccurrenceEffOperations", objEditOccEffDiag, "Pack effectivities")
					ElseIf dicEffectivityMapping("bPackEffectivities") = "False" OR dicEffectivityMapping("bPackEffectivities") = False Then
						Call Fn_CheckBox_Set("Fn_SISW_Eff_OccurrenceEffOperations", objEditOccEffDiag, "Pack effectivities","OFF")
					End If
					
					' clicking on create button
					Call Fn_Button_Click("objCreateOccEffDiag", objEditOccEffDiag,"Create...")
				'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				Case "VerifyOccEffectivityEdit"
					' setting Pack Effectivity checkbox
					If dicEffectivityMapping("bPackEffectivities") = "True" OR dicEffectivityMapping("bPackEffectivities") = True Then
						Call Fn_CheckBox_Select("Fn_SISW_Eff_OccurrenceEffOperations", objEditOccEffDiag, "Pack effectivities")
					ElseIf dicEffectivityMapping("bPackEffectivities") = "False" OR dicEffectivityMapping("bPackEffectivities") = False Then
						Call Fn_CheckBox_Set("Fn_SISW_Eff_OccurrenceEffOperations", objEditOccEffDiag, "Pack effectivities","OFF")
					End If

					' selecting row
					If dicEffectivityMapping("aRowNum") <> "" Then
						Call Fn_UI_JavaTable_SelectRow("Fn_SISW_Eff_OccurrenceEffOperations", objEditOccEffDiag, "EffectivityTable", cInt(dicEffectivityMapping("aRowNum")))
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Function [ Fn_SISW_Eff_OccurrenceEffOperations ] failed with case [ Edit ] .")
						Fn_SISW_Eff_OccurrenceEffOperations = False
						Set objEditOccEffDiag = nothing
						Set objCreateOccEffDiag = nothing
						Exit function
					End If
					Call Fn_Button_Click("Fn_SISW_Eff_OccurrenceEffOperations", objEditOccEffDiag,"Edit...")
			End Select

			IF Fn_UI_ObjectExist("Fn_SISW_Eff_OccurrenceEffOperations", objTempOccDetails1) Then
				Set objCreateOccEffDiag = objTempOccDetails1
			ElseIF Fn_UI_ObjectExist("Fn_SISW_Eff_OccurrenceEffOperations", objTempOccDetails2) Then 
				Set objCreateOccEffDiag = objTempOccDetails2
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_Eff_OccurrenceEffOperations ] Failed to finde Occurrence Effectivity window.")
				Fn_SISW_Eff_OccurrenceEffOperations = False
				Exit Function
			End if	 

			' setting used share effectivity checkbox
			If dicEffectivityMapping("bUsedSharedEffectivity") <> ""  Then
				If  Cbool(dicEffectivityMapping("bUsedSharedEffectivity")) = True Then
					Call Fn_CheckBox_Set("Fn_SISW_Eff_OccurrenceEffOperations",objCreateOccEffDiag,"Use shared effectivity", "ON")
				End If
			End If

			' setting create new / edit existing checkbox
			If dicEffectivityMapping("bCreateNew") <> ""   Then
				If  cBool(dicEffectivityMapping("bCreateNew")) = True Then
					If sAction ="Edit" Then
						objCreateOccEffDiag.JavaCheckBox("Create new").SetTOProperty "attached text" ,"Edit existing"
					Else
						objCreateOccEffDiag.JavaCheckBox("Create new").SetTOProperty "attached text" ,"Create new"
					End If
					Call Fn_CheckBox_Set("Fn_SISW_Eff_OccurrenceEffOperations",objCreateOccEffDiag,"Create new", "ON")
				End If
			End If

			' setting effectivity id
			If dicEffectivityMapping("sEffectivityId") <> ""  Then
				Call Fn_Edit_Box("Fn_SISW_Eff_OccurrenceEffOperations",objCreateOccEffDiag, "EffectivityID", dicEffectivityMapping("sEffectivityId"))
				'if dicEffectivityMapping("bCreateNew") = "" Then
				'	objCreateOccEffDiag.JavaEdit("EffectivityID").Activate
				'ElseIf  cBool(dicEffectivityMapping("bCreateNew")) = False Then
				'	objCreateOccEffDiag.JavaEdit("EffectivityID").Activate
				'End If
				objCreateOccEffDiag.JavaEdit("EffectivityID").Activate

				If instr(dicEffectivityMapping("sEffectivityId"), "*") > 0 Then
					If JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("OccurrenceEffectivityDialog").Exist(15) Then
						Set objOccEffID = JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("OccurrenceEffectivityDialog")
					ElseIf JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("OccurrenceEffectivityDialog").Exist(15) Then
						Set objOccEffID = JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("OccurrenceEffectivityDialog")
					End If
					if objOccEffID.exist(5) then
						if dicEffectivityMapping("sSelectEffectivityId") <> "" then
							'select specified ID from list
							iRows = cInt(objOccEffID.JavaTable("EffectivityIDTable").getROProperty("rows"))
							bFlag = False
							For iCnt = 0 to iRows - 1
								if cstr(objOccEffID.JavaTable("EffectivityIDTable").getCellData(iCnt, "ID")) = dicEffectivityMapping("sSelectEffectivityId") Then
									objOccEffID.JavaTable("EffectivityIDTable").SelectRow iCnt
									bFlag = True
									exit for
								End If
							Next
							if  bFlag <> True then
								Call Fn_Button_Click("Fn_SISW_Eff_OccurrenceEffOperations", objOccEffID,"Close")
								Fn_SISW_Eff_OccurrenceEffOperations = False
								Set objOccEffID = nothing
								Exit function
							End If
							Call Fn_Button_Click("Fn_SISW_Eff_OccurrenceEffOperations", objOccEffID,"OK")
						Else
							Call Fn_Button_Click("Fn_SISW_Eff_OccurrenceEffOperations", objOccEffID,"Close")
						End IF
					End If ' end of if objOccEffID.exist(20) then
				End IF
			End If ' end of If dicEffectivityMapping("sEffectivityId") <> ""  Then

			'Effectivity Protection
			If dicEffectivityMapping("bEffectivityProtection") <> ""  Then
				If   dicEffectivityMapping("bEffectivityProtection") = True Then
					Call Fn_CheckBox_Set("Fn_SISW_Eff_OccurrenceEffOperations",objCreateOccEffDiag,"Apply Access Manager effectivi", "ON")
				Else
					Call Fn_CheckBox_Set("Fn_SISW_Eff_OccurrenceEffOperations",objCreateOccEffDiag,"Apply Access Manager effectivi", "OFF")
				End If
			End If

			'setting end item id
			If  dicEffectivityMapping("sEndItemSelectType") = "" Then
				If dicEffectivityMapping("sEndItem") <> "" Then
					If ucase(dicEffectivityMapping("sEndItem")) <> "CLEAR" Then
						Call Fn_Edit_Box("Fn_SISW_Eff_OccurrenceEffOperations",objCreateOccEffDiag, "EndItem", dicEffectivityMapping("sEndItem"))
						objCreateOccEffDiag.JavaEdit("EndItem").Activate
						'setting end item rev.
						If dicEffectivityMapping("sEndItemRev") <> "" Then
							Call Fn_List_Select("Fn_SISW_Eff_OccurrenceEffOperations",objCreateOccEffDiag, "EndItemRevId", dicEffectivityMapping("sEndItemRev"))
						End If
					Else
						' click on clear
						Call Fn_Button_Click("Fn_SISW_Eff_OccurrenceEffOperations", objCreateOccEffDiag,"ClearEndItem")
					End If
				End If
			Else
				' future use
				Select Case dicEffectivityMapping("sEndItemSelectType")
					Case "MRUList"
						' not yet implemented
					Case "OpenByName"
						Call Fn_CheckBox_Select("Fn_SISW_Eff_OccurrenceEffOperations", objCreateOccEffDiag, "EndItemOpenByName" )
						Call Fn_OpenByNameOperations("CellDoubleClick", dicEffectivityMapping("sEndItemName"), dicEffectivityMapping("sEndItem"),"","","")
					Case "PasteFromClipboard"
						' not yet implemented
				End Select
			End If
			' setting unit radio button
			If  dicEffectivityMapping("bUnit") <> ""  Then
				If Cbool(dicEffectivityMapping("bUnit")) = True Then
					'--- insert code for removal of dates
					If cInt( Fn_UI_Object_GetROProperty("Fn_SISW_Eff_OccurrenceEffOperations", objCreateOccEffDiag.JavaRadioButton("Dates"), "enabled")) = 1Then
						If cint(Fn_UI_Object_GetROProperty("Fn_SISW_Eff_OccurrenceEffOperations", objCreateOccEffDiag.JavaRadioButton("Dates"), "value"))  = 1Then
							Set objTable = objCreateOccEffDiag.JavaTable("DateRangeTable")
							iRows = objTable.GetROProperty("rows")
							For iCount =0 to iRows -1
								If trim(objTable.GetCellData (iCount,"From Date")) = "" AND trim(objTable.GetCellData (iCount,"To Date")) = "" then 
									exit for
								End If
								objTable.SelectCell iCount,"From Date"
								Call Fn_Button_Click("Fn_SISW_Eff_OccurrenceEffOperations", objCreateOccEffDiag, "Clear Date")
								wait 1
								objTable.SelectCell iCount,"To Date"
								Call Fn_Button_Click("Fn_SISW_Eff_OccurrenceEffOperations", objCreateOccEffDiag, "Clear Date")
							Next
						End If
					End If

					'setting units
					Call Fn_UI_JavaRadioButton_SetON("Fn_SISW_Eff_OccurrenceEffOperations", objCreateOccEffDiag,"Units")
					If trim(dicEffectivityMapping("sUnit")) <> "SO" AND trim(dicEffectivityMapping("sUnit")) <> "UP" Then
						Call Fn_Edit_Box("Fn_SISW_Eff_OccurrenceEffOperations",objCreateOccEffDiag, "Units", dicEffectivityMapping("sUnit"))
					Else
						Call Fn_Button_Click("Fn_SISW_Eff_OccurrenceEffOperations", objCreateOccEffDiag, trim(dicEffectivityMapping("sUnit")))
					End If
				End If
			End If
			' setting date radio
			If  dicEffectivityMapping("bDate") <>""  Then
				If  cBool(dicEffectivityMapping("bDate") ) = True Then
					' insert code for removal of units
					If cInt( Fn_UI_Object_GetROProperty("Fn_SISW_Eff_OccurrenceEffOperations", objCreateOccEffDiag.JavaRadioButton("Units"), "enabled")) = 1Then
						If cint(Fn_UI_Object_GetROProperty("Fn_SISW_Eff_OccurrenceEffOperations", objCreateOccEffDiag.JavaRadioButton("Units"), "value"))  = 1Then
							objCreateOccEffDiag.JavaEdit("Units").Set "" 
						End If
					End If
					Call Fn_UI_JavaRadioButton_SetON("Fn_SISW_Eff_OccurrenceEffOperations", objCreateOccEffDiag,"Dates")
					
					' selecting date
					aStartDate  = split(dicEffectivityMapping("sStartDates"),":")
					aEndDate =  split(dicEffectivityMapping("sEndDates"),":")

					For iCount = 0 to uBound(aStartDate)
						If instr(sAction, "LegacyOccurrenceEffectivity") > 0 Then
							Call Fn_EffectivitySetDate("LegacyOccurrenceEffectivity", aStartDate(iCount))
						Else
							Call Fn_EffectivitySetDate("Occurrence", aStartDate(iCount))
						End If
						If trim(aEndDate(iCount)) <> "SO" AND trim(aEndDate(iCount)) <> "UP" Then
							If instr(sAction, "LegacyOccurrenceEffectivity") > 0 Then
								Call Fn_EffectivitySetDate("LegacyOccurrenceEffectivity", aStartDate(iCount))
							Else
								Call Fn_EffectivitySetDate("Occurrence", aEndDate(iCount))
							End If
						Else
							Call Fn_Button_Click("Fn_SISW_Eff_OccurrenceEffOperations", objCreateOccEffDiag, trim(aEndDate(iCount)))
						End If
					Next
				End If
			Else
				If sAction =  "LegacyOccurrenceEffectivity" Then
					If dicEffectivityMapping("sStartDates") <> "" AND  dicEffectivityMapping("sEndDates") <> "" Then
						' selecting date
						aStartDate  = split(dicEffectivityMapping("sStartDates"),":")
						aEndDate =  split(dicEffectivityMapping("sEndDates"),":")

						For iCount = 0 to uBound(aStartDate)
							Call Fn_EffectivitySetDate("LegacyOccurrenceEffectivity", aStartDate(iCount))
							Call Fn_EffectivitySetDate("LegacyOccurrenceEffectivity", aEndDate(iCount))
						Next
					End If
				End If
			End If
			'clciking on OK
			Call Fn_Button_Click("Fn_SISW_Eff_OccurrenceEffOperations",objCreateOccEffDiag,"OK")
			' addding code for upgrade
			JavaDialog("VariantItemRevise").setTOProperty "title","Upgrade Legacy Occurrence Effectivity?"
			If JavaDialog("VariantItemRevise").exist(10) Then
				Call Fn_Button_Click("Fn_SISW_Eff_OccurrenceEffOperations",JavaDialog("VariantItemRevise"),"Yes")
			End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case Else
			Fn_SISW_Eff_OccurrenceEffOperations = False
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_Eff_OccurrenceEffOperations ] Execution Failed. [ " & sAction & " ] is invalid.")
	End Select
	Wait 2
	If lcase(typename(objEditOccEffDiag)) <> "empty" Then
		If Fn_UI_ObjectExist("Fn_SISW_Eff_OccurrenceEffOperations", objEditOccEffDiag) <> False Then
			Call Fn_Button_Click("Fn_SISW_Eff_OccurrenceEffOperations", objEditOccEffDiag,"Close")
		End If
	End If

	Set objEditOccEffDiag = nothing
	Set objCreateOccEffDiag = nothing
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [ Fn_SISW_Eff_OccurrenceEffOperations ] Successfully executed.")
End Function
'****************************************    Function to perform Revision Effectivity Operations  ***************************************

'Function Name		      :			  Fn_SISW_Eff_RevisionEffOperations

'Description			  :  	      Function to perform Revision Effectivity Operations

'Parameters			 :	   	 	1. sModuleName : Module Name
'								  2. sAction : Action need to perform
'								  3. dicEffectivityMapping : Dictionary object of dicEffectivityMapping
'											
'Return Value		       	: 			True \ False
'Pre-requisite			 :		 	 Revision of an Item should be selected.
											'Need to set following preference:  CFMOccEffMode = maintenance
'Examples			 :			  Otherwise It will open in Legacy mode ( CFMOccEffMode = legacy )
											''Declaration for Effectivity Mapping Dictioanry objects 
'																						
'								dicEffectivityMapping("sColName") = ""
'								dicEffectivityMapping("sValue") = ""
'								dicEffectivityMapping("aRowNum") = "1"
'								dicEffectivityMapping("bPackEffectivities") = True ' -  True / False
'								dicEffectivityMapping("bUsedSharedEffectivity") = True ' -  True / False
'								dicEffectivityMapping("bCreateNew") = True ' -  True / False
'								dicEffectivityMapping("sVerifyEffectivityId") = "re2~res3"
'								dicEffectivityMapping("sSelectEffectivityId") = "re2"
'								dicEffectivityMapping("sEffectivityId") = "d*"
'								dicEffectivityMapping("bEffectivityProtection") = ""
'								dicEffectivityMapping("sEndItem") = "000065" / "Clear"
'								dicEffectivityMapping("sEndItemSelectType") = "OpenByName" '- "OpenByName", "MRUList", "PasteFromClipboard"
'								dicEffectivityMapping("sEndItemMRU") = ""
'								dicEffectivityMapping("sEndItemName") = "Top"
'								dicEffectivityMapping("sEndItemRev") = "A"
'								dicEffectivityMapping("bUnit") = ""  ' -  True / False
'								dicEffectivityMapping("sUnit") = ""
'								dicEffectivityMapping("bDate") = True  ' -  True / False
'								dicEffectivityMapping("sStartDates") = "24-Oct-2010~19-20-20"'  - should Follow Format  DD-MMM-YYYY~HH-MM-SS     HH[00-23] seperated by :
'  																   "Oct 26, 2010  (12:00 AM)"'  - For Case "VerifyRevisionEffectivityDetails" - should Follow Format  MMM DD, YYYY  (HH:MM AM/PM)  seperated by ~
'								dicEffectivityMapping("sEndDates") = "27-Oct-2010~10-20-20"'- should Follow Format  DD-MMM-YYYY~HH-MM-SS     HH[00-23] seperated by :
'																  "Oct 29, 2010  (12:00 AM)"'  - For Case "VerifyRevisionEffectivityDetails" - should Follow Format  MMM DD, YYYY  (HH:MM AM/PM)  seperated by ~

'								 sAction = "Create"  "Edit" , "Copy", "Delete","VerifyRevisionEffectivity"

'								Call Fn_SISW_Eff_RevisionEffOperations("My Teamcenter","Create", dicEffectivityMapping)
'								Call Fn_SISW_Eff_RevisionEffOperations("Structure Manager","Edit", dicEffectivityMapping)
'								Call Fn_SISW_Eff_RevisionEffOperations("Structure Manager","Copy", dicEffectivityMapping)
'								Call Fn_SISW_Eff_RevisionEffOperations("Structure Manager","Delete", dicEffectivityMapping)
'								Call Fn_SISW_Eff_RevisionEffOperations("Structure Manager","VerifyRevisionEffectivity", dicEffectivityMapping)
'								Call Fn_SISW_Eff_RevisionEffOperations("Structure Manager","VerifyRevisionEffectivityDetails", dicEffectivityMapping)
'								Call Fn_SISW_Eff_RevisionEffOperations("CM Viewer","Edit", dicEffectivityMapping)

'History:
'						Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'					 	   Koustubh W	 		6-Oct-2010		     1.0										    Rupali
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'					 	   Koustubh W	 		11-Oct-2010		     1.0					Added Case VerifyRevisionEffectivityDetails
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'					 	   Koustubh W	 		22-Oct-2010		     1.0			Added code for buttons SO and UP			Rupali
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'					 	   Koustubh W	 		27-Oct-2010		     1.0			Added case for Structure Manage Edit IC
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'					 	   Koustubh W	 		 8-Nov-2010		     1.0			Added case for CM Viewer perspective
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'					 	   Koustubh W	 		 19-Nov-2010		     1.0		Added code to set Pack Effectivity checkbox
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'					 	   Koustubh W	 		 19-Nov-2010		     1.0		Added code to set verify multiple effectivity IDs
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'					 	   Koustubh W	 		 10-Mar-2011		     1.0		Updated function deficnition from Tc8.3
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'					 	   Koustubh W	 		 14-Mar-2011		     1.0		Updated function modified case Edit of Structure Manager_EditIC
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'					 	   Koustubh W	 		 05-Apr-2011		     1.0		Updated function modified case Edit
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'					 	  Ashok kakade	 		 15-May-2012		     1.0		Updated  object Hierarchy of  objRevEffectivity & objRevEffectDetails
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_Eff_RevisionEffOperations(sModuleName, sAction, dicEffectivityMapping )
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_Eff_RevisionEffOperations"
	Dim objRevEffectivity, objRevEffectDetails, aCols, aVals, aRows, iCount, iRows
	Dim aStartDate, aEndDate
	Dim objTable,objTcDefaultWindow,objJavaApplet,objReleaseStatEff,objReleaseStatEff2
	Dim objOccEffID, aEffID, bFlag, iCnt
	Dim sMenuFile, sMenu
	
	Fn_SISW_Eff_RevisionEffOperations = True
         Set objTcDefaultWindow = Fn_SISW_GetObject("Effectivity")
         Set objJavaApplet = Fn_SISW_GetObject("Effectivity@1")
         Set objReleaseStatEff= Fn_SISW_GetObject("Release Status Effectivity")
 	Set objReleaseStatEff2= Fn_SISW_GetObject("Release Status Effectivity @2")
 	
	Select Case sModuleName
		Case "My Teamcenter", "Structure Manager", "CM Viewer"

			If Fn_SISW_UI_Object_Operations("Fn_SISW_Eff_RevisionEffOperations", "Exist", objTcDefaultWindow, SISW_MICRO_TIMEOUT) Then
				Set objRevEffectivity = objTcDefaultWindow
			ElseIF Fn_SISW_UI_Object_Operations("Fn_SISW_Eff_RevisionEffOperations", "Exist", objJavaApplet, SISW_MICRO_TIMEOUT) Then
				Set objRevEffectivity = objJavaApplet
			Else
				Select Case sModuleName
					Case "My Teamcenter", "CM Viewer"
								sMenuFile = Fn_LogUtil_GetXMLPath("MyTc_Menu")
								sMenu = Fn_GetXMLNodeValue(sMenuFile, "ViewEffectivityRevisionEffectivity")
								Call Fn_MenuOperation("Select",sMenu)
								Call Fn_ReadyStatusSync(1)
								
								If Fn_SISW_UI_Object_Operations("Fn_SISW_Eff_RevisionEffOperations", "Exist", objTcDefaultWindow, SISW_MICROLESS_TIMEOUT) Then
									Set objRevEffectivity = objTcDefaultWindow
								ElseIF Fn_SISW_UI_Object_Operations("Fn_SISW_Eff_RevisionEffOperations", "Exist", objJavaApplet, SISW_MICROLESS_TIMEOUT) Then
									Set objRevEffectivity = objJavaApplet
								Else	
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Function [ Fn_SISW_Eff_RevisionEffOperations ] Failed to open Revision Effectivity Dialogbox.")
									Fn_SISW_Eff_RevisionEffOperations = False
									Exit function
								End If
											
					 Case "Structure Manager"
								sMenuFile = Fn_LogUtil_GetXMLPath("PSE_Menu")
								sMenu = Fn_GetXMLNodeValue(sMenuFile, "ToolsEffectivityRevisionEffectivity")
'------------------Changes are made to identify the proper QTP version and act accordingly---------------------------
								If Environment.Value("ProductName") = sUFTProductName Then
									Call Fn_MenuOperation("WinMenuSelect", sMenu)
							    Else
							    	Call Fn_MenuOperation("Select", sMenu)
							    End If
								'Call Fn_MenuOperation("Select",sMenu)
								Call Fn_ReadyStatusSync(1)
								
								If Fn_SISW_UI_Object_Operations("Fn_SISW_Eff_RevisionEffOperations", "Exist", objTcDefaultWindow, SISW_MICROLESS_TIMEOUT) Then
									Set objRevEffectivity = objTcDefaultWindow
								ElseIF Fn_SISW_UI_Object_Operations("Fn_SISW_Eff_RevisionEffOperations", "Exist", objJavaApplet, SISW_MICROLESS_TIMEOUT) Then
									Set objRevEffectivity = objJavaApplet
								Else	
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Function [ Fn_SISW_Eff_RevisionEffOperations ] Failed to open Revision Effectivity Dialogbox.")
									Fn_SISW_Eff_RevisionEffOperations = False
									Exit function
								End If	
					End Select		 
				End If
'				If JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Effectivity").Exist(5) Then
'					Set objRevEffectivity = JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Effectivity")
'				ElseIF JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Effectivity").Exist(5) Then
'					Set objRevEffectivity = JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Effectivity")
'				If Fn_SISW_UI_Object_Operations("Fn_SISW_Eff_RevisionEffOperations", "Exist", objTcDefaultWindow, SISW_MICROLESS_TIMEOUT) Then
'					Set objRevEffectivity = objTcDefaultWindow
'				ElseIF Fn_SISW_UI_Object_Operations("Fn_SISW_Eff_RevisionEffOperations", "Exist", objJavaApplet, SISW_MICROLESS_TIMEOUT) Then
'					Set objRevEffectivity = objJavaApplet	
'				Else 
'					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Function [ Fn_SISW_Eff_RevisionEffOperations ] Failed to open Revision Effectivity Dialogbox.")
'					Fn_SISW_Eff_RevisionEffOperations = False
'					Exit function
'			   End If
'			End If
				If Fn_SISW_UI_Object_Operations("Fn_SISW_Eff_RevisionEffOperations", "Exist", objTcDefaultWindow, SISW_MICROLESS_TIMEOUT) Then
					Set objRevEffectivity = objTcDefaultWindow
				ElseIF Fn_SISW_UI_Object_Operations("Fn_SISW_Eff_RevisionEffOperations", "Exist", objJavaApplet, SISW_MICROLESS_TIMEOUT) Then
					Set objRevEffectivity = objJavaApplet
				Else	
					Call Fn_MenuOperation("Select",sMenu)
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Function [ Fn_SISW_Eff_RevisionEffOperations ] Failed to open Revision Effectivity Dialogbox.")
							Fn_SISW_Eff_RevisionEffOperations = False
							Exit function
				End If
		Case "Structure Manager_EditIC"
			' Do nothing
			'Set objRevEffectivity = JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Effectivity")
			Set objRevEffectivity =objJavaApplet
	End Select

	Select Case sAction
	     	Case "Create", "Edit", "Copy"
			Select Case sAction
				Case  "Create"
						' setting Pack Effectivity checkbox
					  	If dicEffectivityMapping("bPackEffectivities") <> "" then
							If dicEffectivityMapping("bPackEffectivities") = "True" OR dicEffectivityMapping("bPackEffectivities") = True Then
								Call Fn_CheckBox_Select("Fn_EffectivityMappingOperations", objRevEffectivity, "Pack effectivities")
							Else
								Call Fn_CheckBox_Set("Fn_EffectivityMappingOperations", objRevEffectivity, "Pack effectivities","OFF")
							End If
   					   End If
'						If JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Release Status Effectivity").Exist(5) = False Then
'							If JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Release Status Effectivity").Exist(5) = False Then
'								Call Fn_Button_Click("Fn_SISW_Eff_RevisionEffOperations", objRevEffectivity,"Create...")
'							End If
'						End If
						
						If Fn_SISW_UI_Object_Operations("Fn_SISW_Eff_RevisionEffOperations", "Exist", objReleaseStatEff, SISW_MICROLESS_TIMEOUT)  = False Then
							If Fn_SISW_UI_Object_Operations("Fn_SISW_Eff_RevisionEffOperations", "Exist", objReleaseStatEff2, SISW_MICROLESS_TIMEOUT) = False Then
								Call Fn_Button_Click("Fn_SISW_Eff_RevisionEffOperations", objRevEffectivity,"Create...")
							End If
						End If

				Case  "Edit", "Copy"
					Select Case sModuleName
						Case "My Teamcenter", "Structure Manager", "CM Viewer"
								If objRevEffectivity.Exist(2) Then
									'Do Nothing
								Else
									'Set objRevEffectivity = JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Effectivity")
									Set objRevEffectivity = objJavaApplet
								End If
								' setting Pack Effectivity checkbox
								if dicEffectivityMapping("bPackEffectivities") <> "" then
									If dicEffectivityMapping("bPackEffectivities") = "True" OR dicEffectivityMapping("bPackEffectivities") = True Then
										Call Fn_CheckBox_Select("Fn_EffectivityMappingOperations", objRevEffectivity, "Pack effectivities")
									Else
										Call Fn_CheckBox_Set("Fn_EffectivityMappingOperations", objRevEffectivity, "Pack effectivities","OFF")
									End If
								End If
								
								' selecting row
								If dicEffectivityMapping ("aRowNum") <> "" Then
									Call Fn_UI_JavaTable_SelectRow("Fn_SISW_Eff_RevisionEffOperations", objRevEffectivity, "EffectivityTable", cInt(dicEffectivityMapping("aRowNum")))
								Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Function [ Fn_SISW_Eff_RevisionEffOperations ] failed with case [ Edit ] .")
									Fn_SISW_Eff_RevisionEffOperations = False
									Set objRevEffectivity = nothing
									Set objRevEffectDetails = nothing
									Exit function
								End If
								
								If sAction = "Edit" Then
									' clicking on edit
									Call Fn_Button_Click("Fn_SISW_Eff_RevisionEffOperations", objRevEffectivity,"Edit...")
								Else
									' clicking on edit
									Call Fn_Button_Click("Fn_SISW_Eff_RevisionEffOperations", objRevEffectivity,"Copy...")
								End If
						Case "Structure Manager_EditIC"
								' Do nothing
								Set objRevEffectivity = JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Effectivity")
					End Select
			End Select

'			If JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Release Status Effectivity").Exist(5) Then
'				Set objRevEffectDetails = JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Release Status Effectivity")
'			ElseIF JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Release Status Effectivity").Exist(5) Then
'				Set objRevEffectDetails = JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Release Status Effectivity")
			If Fn_SISW_UI_Object_Operations("Fn_SISW_Eff_RevisionEffOperations", "Exist", objReleaseStatEff, SISW_MINLESS_TIMEOUT) Then
				Set objRevEffectDetails = objReleaseStatEff
			ElseIF  Fn_SISW_UI_Object_Operations("Fn_SISW_Eff_RevisionEffOperations", "Exist", objReleaseStatEff2, SISW_MINLESS_TIMEOUT)  Then
				Set objRevEffectDetails = objReleaseStatEff2
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Function [ Fn_SISW_Eff_RevisionEffOperations ] failed to find [ Release Status Effectivity ] window.")
				Fn_SISW_Eff_RevisionEffOperations = False
				Exit Function 
			End If

			' setting used share effectivity checkbox
			If dicEffectivityMapping("bUsedSharedEffectivity") <> ""  Then
				If  Cbool(dicEffectivityMapping("bUsedSharedEffectivity")) = True Then
					Call Fn_CheckBox_Set("Fn_SISW_Eff_RevisionEffOperations",objRevEffectDetails,"Use shared effectivity", "ON")
				End If
			End If

			' setting create new / edit existing checkbox
			If dicEffectivityMapping("bCreateNew") <> ""   Then
				If   cBool(dicEffectivityMapping("bCreateNew")) = True Then
					If sAction ="Edit" Then
						objRevEffectDetails.JavaCheckBox("Create new").SetTOProperty "attached text" ,"Edit existing"
						If cBool(dicEffectivityMapping("bEditCreateNew")) = True Then
							objRevEffectDetails.JavaCheckBox("Create new").SetTOProperty "attached text" ,"Create new"
						End If
					Else
						objRevEffectDetails.JavaCheckBox("Create new").SetTOProperty "attached text" ,"Create new"
					End If
					
					Call Fn_CheckBox_Set("Fn_SISW_Eff_RevisionEffOperations",objRevEffectDetails,"Create new", "ON")
				End If
			End If
			
			' setting effectivity id
			If dicEffectivityMapping("sEffectivityId") <> ""  Then
				Call Fn_Edit_Box("Fn_SISW_Eff_RevisionEffOperations",objRevEffectDetails, "EffectivityID", dicEffectivityMapping("sEffectivityId"))
				if dicEffectivityMapping("bCreateNew") = "" Then
						objRevEffectDetails.JavaEdit("EffectivityID").Activate
				ElseIf  cBool(dicEffectivityMapping("bCreateNew") ) = False Then
						objRevEffectDetails.JavaEdit("EffectivityID").Activate
				End If
				If instr(dicEffectivityMapping("sEffectivityId") , "*") > 0 Then
					Set objOccEffID = JavaWindow("StructureManager").JavaWindow("PSEWindow").JavaDialog("OccurrenceEffectivityDialog")
					if objOccEffID.exist(20) then
						if dicEffectivityMapping("sSelectEffectivityId") <> "" then
							'select specified ID from list
							iRows = cInt(objOccEffID.JavaTable("EffectivityIDTable").getROProperty("rows"))
							bFlag = False
							for iCnt = 0 to iRows - 1
								if cstr(objOccEffID.JavaTable("EffectivityIDTable").getCellData(iCnt, "ID")) = dicEffectivityMapping("sSelectEffectivityId") Then
									objOccEffID.JavaTable("EffectivityIDTable").SelectRow iCnt
									Call Fn_SISW_UI_JavaTable_Operations("Fn_SISW_Eff_RevisionEffOperations", "SelectRow", objOccEffID.JavaTable("EffectivityIDTable") , "", "", "", iCnt, "", "", "", "")
									bFlag = True
									exit for
								End If
							Next
							if  bFlag <> True then
								Call Fn_Button_Click("Fn_SISW_Eff_RevisionEffOperations", objOccEffID,"Close")
								Fn_SISW_Eff_RevisionEffOperations = False
								Set objOccEffID = nothing
								Exit function
							End If
						End IF
						Call Fn_Button_Click("Fn_SISW_Eff_RevisionEffOperations", objOccEffID,"OK")
					End If ' end of if objOccEffID.exist(20) then
				End IF
			End If ' end of If dicEffectivityMapping("sEffectivityId") <> ""  Then

			'Effectivity Protection
			If dicEffectivityMapping("bEffectivityProtection") <> ""  Then
				If   dicEffectivityMapping("bEffectivityProtection") = True Then
					Call Fn_CheckBox_Set("Fn_SISW_Eff_RevisionEffOperations",objRevEffectDetails,"ApplyAccessManagerEffectivityProtection", "ON")
				Else
					Call Fn_CheckBox_Set("Fn_SISW_Eff_RevisionEffOperations",objRevEffectDetails,"ApplyAccessManagerEffectivityProtection", "OFF")
				End If
			End If
			'setting end item id
			If  dicEffectivityMapping("sEndItemSelectType") = "" Then
				If dicEffectivityMapping("sEndItem") <> "" Then
					If uCase(dicEffectivityMapping("sEndItem")) <> "CLEAR" Then
						Call Fn_Edit_Box("Fn_SISW_Eff_RevisionEffOperations",objRevEffectDetails, "EndItem", dicEffectivityMapping("sEndItem"))
						objRevEffectDetails.JavaEdit("EndItem").Activate
						'setting end item rev.
						If dicEffectivityMapping("sEndItemRev") <> "" Then
								Call Fn_List_Select("Fn_SISW_Eff_RevisionEffOperations",objRevEffectDetails, "EndItemRevId", dicEffectivityMapping("sEndItemRev"))
						End If
					Else
						' clicking on clear end item
						Call Fn_Button_Click("Fn_SISW_Eff_RevisionEffOperations", objRevEffectDetails, "ClearEndItem")
					End IF
				End If
			Else
				' future use
				Select Case dicEffectivityMapping("sEndItemSelectType")
					Case "MRUList"
					Case "OpenByName"
							Call Fn_CheckBox_Select("Fn_SISW_Eff_RevisionEffOperations", objRevEffectDetails, "EndItemOpenByName" )
							Call Fn_OpenByNameOperations("CellDoubleClick", dicEffectivityMapping("sEndItemName"), dicEffectivityMapping("sEndItem"),"","","")
					Case "PasteFromClipboard"
				End Select
			End If

			' setting unit radio button
			If  dicEffectivityMapping("bUnit") <> ""  Then
				If Cbool(dicEffectivityMapping("bUnit")) = True Then
						' clearing dates
						If cInt( Fn_UI_Object_GetROProperty("Fn_SISW_Eff_RevisionEffOperations", objRevEffectDetails.JavaRadioButton("Dates"), "enabled")) = 1Then
							If cint(Fn_UI_Object_GetROProperty("Fn_SISW_Eff_RevisionEffOperations", objRevEffectDetails.JavaRadioButton("Dates"), "value"))  = 1Then
								Set objTable = objRevEffectDetails.JavaTable("DateRangeTable")
								'iRows = objTable.GetROProperty("rows")
								iRows = cint(Fn_UI_Object_GetROProperty("Fn_SISW_Eff_RevisionEffOperations", objTable, "rows"))
								For iCount =0 to iRows -1
									If trim(objTable.GetCellData (iCount,"From Date")) = "" AND trim(objTable.GetCellData (iCount,"To Date")) = "" then exit for
									'objTable.SelectCell iCount,"From Date"
									Call Fn_SISW_UI_JavaTable_Operations("Fn_SISW_Eff_RevisionEffOperations", "SelectCell", objTable , "", "", "", iCount, "From Date", "", "", "")
									Call Fn_Button_Click("Fn_SISW_Eff_RevisionEffOperations", objRevEffectDetails, "Clear Date")
									wait 1
									'objTable.SelectCell iCount,"To Date"
									Call Fn_SISW_UI_JavaTable_Operations("Fn_SISW_Eff_RevisionEffOperations", "SelectCell", objTable , "", "", "", iCount, "To Date", "", "", "")
									Call Fn_Button_Click("Fn_SISW_Eff_RevisionEffOperations", objRevEffectDetails, "Clear Date")
								Next
							End If
						End If
						' setting units
						Call Fn_UI_JavaRadioButton_SetON("Fn_SISW_Eff_RevisionEffOperations", objRevEffectDetails,"Units")
						If trim(dicEffectivityMapping("sUnit")) <> "SO" AND trim(dicEffectivityMapping("sUnit")) <> "UP" Then
							Call Fn_Edit_Box("Fn_SISW_Eff_RevisionEffOperations",objRevEffectDetails, "Units", dicEffectivityMapping("sUnit"))
						Else
							Call Fn_Button_Click("Fn_SISW_Eff_RevisionEffOperations", objRevEffectDetails, trim(dicEffectivityMapping("sUnit")))
						End If
				End If
			End If
			' setting date radio
			If  dicEffectivityMapping("bDate") <>""  Then
				If  cBool(dicEffectivityMapping("bDate") ) = True Then
						' clearing Units
						If cInt( Fn_UI_Object_GetROProperty("Fn_SISW_Eff_RevisionEffOperations", objRevEffectDetails.JavaRadioButton("Units"), "enabled")) = 1Then
							If cint(Fn_UI_Object_GetROProperty("Fn_SISW_Eff_RevisionEffOperations", objRevEffectDetails.JavaRadioButton("Units"), "value"))  = 1Then
								objRevEffectDetails.JavaEdit("Units").Set "" 
							End If
						End If
						Call Fn_UI_JavaRadioButton_SetON("Fn_SISW_Eff_RevisionEffOperations", objRevEffectDetails,"Dates")
						If sAction = "Edit" OR sAction = "Create" Then
							Set objTable = objRevEffectDetails.JavaTable("DateRangeTable")
							if sAction = "Edit" then
								Set objTable = objRevEffectDetails.JavaTable("DateRangeTable")
								'iRows = objTable.GetROProperty("rows")
								iRows = cint(Fn_UI_Object_GetROProperty("Fn_SISW_Eff_RevisionEffOperations", objTable, "rows"))
								For iCount =0 to iRows -1
									If trim(objTable.GetCellData (iCount,"From Date")) = "" AND trim(objTable.GetCellData (iCount,"To Date")) = "" then exit for
									'objTable.SelectCell iCount,"From Date"
									Call Fn_SISW_UI_JavaTable_Operations("Fn_SISW_Eff_RevisionEffOperations", "SelectCell", objTable , "", "", "", iCount, "From Date", "", "", "")
									Call Fn_Button_Click("Fn_SISW_Eff_RevisionEffOperations", objRevEffectDetails, "Clear Date")
									wait 1
									'objTable.SelectCell iCount,"To Date"
									Call Fn_SISW_UI_JavaTable_Operations("Fn_SISW_Eff_RevisionEffOperations", "SelectCell", objTable , "", "", "", iCount, "To Date", "", "", "")
									Call Fn_Button_Click("Fn_SISW_Eff_RevisionEffOperations", objRevEffectDetails, "Clear Date")
								Next
							End If
							objTable.SelectCell 0,"From Date"
							Set objTable = nothing
						End If
						' selecting date
						aStartDate  = split(dicEffectivityMapping("sStartDates"),":")
						aEndDate =  split(dicEffectivityMapping("sEndDates"),":")
						For iCount = 0 to uBound(aStartDate)
								Call Fn_EffectivitySetDate("ReleaseStatusEffectivity", aStartDate(iCount))
								If trim(aEndDate(iCount)) <> "SO" AND trim(aEndDate(iCount)) <> "UP" Then
									Call Fn_EffectivitySetDate("ReleaseStatusEffectivity", aEndDate(iCount))
								Else
									Call Fn_Button_Click("Fn_SISW_Eff_RevisionEffOperations", objRevEffectDetails, trim(aEndDate(iCount)))
								End If
						Next
					End If
				End If
		
			'clciking on OK
			Call Fn_Button_Click("Fn_SISW_Eff_RevisionEffOperations",objRevEffectDetails,"OK")
			wait(1)
			'If objRevEffectivity.exist(1) Then
			If Fn_SISW_UI_Object_Operations("Fn_SISW_Eff_RevisionEffOperations", "Exist", objRevEffectivity, SISW_MINLESS_TIMEOUT) Then
				' closing effectivity dialog window
				Call Fn_Button_Click("Fn_SISW_Eff_RevisionEffOperations", objRevEffectivity,"Close")
			End If
		
		'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		Case "Delete"
						' setting Pack Effectivity checkbox
						If dicEffectivityMapping("bPackEffectivities") = "True" OR dicEffectivityMapping("bPackEffectivities") = True Then
							Call Fn_CheckBox_Select("Fn_EffectivityMappingOperations", objRevEffectivity, "Pack effectivities")
						Else
							Call Fn_CheckBox_Set("Fn_EffectivityMappingOperations", objRevEffectivity, "Pack effectivities","OFF")
						End If
						If dicEffectivityMapping ("aRowNum") <> "" Then
								Call Fn_UI_JavaTable_SelectRow("Fn_SISW_Eff_RevisionEffOperations", objRevEffectivity, "EffectivityTable", cInt(dicEffectivityMapping("aRowNum")))
						Else
								Fn_SISW_Eff_RevisionEffOperations = False
						End If
			
						' Clicking on Delete
						Call Fn_Button_Click("Fn_SISW_Eff_RevisionEffOperations", objRevEffectivity,"Delete")
			
						' closing effectivity dialog window
						Call Fn_Button_Click("Fn_SISW_Eff_RevisionEffOperations", objRevEffectivity,"Close")

	'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		Case "VerifyRevisionEffectivity"
									' setting Pack Effectivity checkbox
						If dicEffectivityMapping("bPackEffectivities") = "True" OR dicEffectivityMapping("bPackEffectivities") = True Then
							Call Fn_CheckBox_Select("Fn_EffectivityMappingOperations", objRevEffectivity, "Pack effectivities")
						Else
							Call Fn_CheckBox_Set("Fn_EffectivityMappingOperations", objRevEffectivity, "Pack effectivities","OFF")
						End If
						iRows = objRevEffectivity.JavaTable("EffectivityTable").GetROProperty("rows")
						aCols = split(dicEffectivityMapping("sColName"),"~")
						aVals = split(dicEffectivityMapping("sValue"),"~")
						aRows = split(dicEffectivityMapping("aRowNum"),"~")
			
						For iCount = 0 to uBound(aRows)
							If cstr(objRevEffectivity.JavaTable("EffectivityTable").GetCellData (cint(aRows(iCount)), aCols(iCount))) <> aVals(iCount) Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Function [ Fn_SISW_Eff_RevisionEffOperations ] failed with case [ "& sAction & " ] .")
								Fn_SISW_Eff_RevisionEffOperations = False
								Set objRevEffectivity = nothing
								Set objRevEffectDetails = nothing
								Exit function
							End If	
						Next
						' closing effectivity dialog window
						Call Fn_Button_Click("Fn_SISW_Eff_RevisionEffOperations", objRevEffectivity,"Close")
			'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		Case "VerifyRevisionEffectivityDetails"
				' setting Pack Effectivity checkbox
				If dicEffectivityMapping("bPackEffectivities") = "True" OR dicEffectivityMapping("bPackEffectivities") = True Then
					Call Fn_CheckBox_Select("Fn_EffectivityMappingOperations", objRevEffectivity, "Pack effectivities")
				Else
					Call Fn_CheckBox_Set("Fn_EffectivityMappingOperations", objRevEffectivity, "Pack effectivities","OFF")
				End If
				Call Fn_Button_Click("Fn_SISW_Eff_RevisionEffOperations", objRevEffectivity,"Create...")

'				If JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Release Status Effectivity").Exist Then
'					Set objRevEffectDetails = JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Release Status Effectivity")
'				ElseIF JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Release Status Effectivity").Exist Then
'					Set objRevEffectDetails = JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Release Status Effectivity")
				If Fn_SISW_UI_Object_Operations("Fn_SISW_Eff_RevisionEffOperations", "Exist", objReleaseStatEff, SISW_MICROLESS_TIMEOUT) Then
					Set objRevEffectDetails =objReleaseStatEff
				ElseIF  Fn_SISW_UI_Object_Operations("Fn_SISW_Eff_RevisionEffOperations", "Exist", objReleaseStatEff2, SISW_MICROLESS_TIMEOUT) Then
					Set objRevEffectDetails = objReleaseStatEff2
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Function [ Fn_SISW_Eff_RevisionEffOperations ] failed to find [ Release Status Effectivity ] window.")
					Fn_SISW_Eff_RevisionEffOperations = False
					Exit Function 
				End If
				'If objRevEffectDetails.Exist(1) = False Then
				If Fn_SISW_UI_Object_Operations("Fn_SISW_Eff_RevisionEffOperations", "Exist", objRevEffectDetails, SISW_MICRO_TIMEOUT) = False Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Function [ Fn_SISW_Eff_RevisionEffOperations ] failed with case [ Verify ] .")
						Fn_SISW_Eff_RevisionEffOperations = False
						Set objRevEffectivity = nothing
						Set objRevEffectDetails = nothing
						Exit function
				End If
	
				' setting used share effectivity checkbox
				If dicEffectivityMapping("bUsedSharedEffectivity") <> ""  Then
					If  Cbool(dicEffectivityMapping("bUsedSharedEffectivity")) = True Then
						Call Fn_CheckBox_Set("Fn_SISW_Eff_RevisionEffOperations",objRevEffectDetails,"Use shared effectivity", "ON")
					End If
				End If
	
				' setting create new / edit existing checkbox
				If dicEffectivityMapping("bCreateNew") <> ""   Then
					If  cBool(dicEffectivityMapping("bCreateNew")) = True Then
						If sAction ="Edit" Then
							'objRevEffectDetails.JavaCheckBox("Create new").SetTOProperty "attached text" ,"Edit existing"
							Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_SISW_Eff_RevisionEffOperations",objRevEffectDetails.JavaCheckBox("Create new"),"attached text","Edit existing")
						Else
							'objRevEffectDetails.JavaCheckBox("Create new").SetTOProperty "attached text" ,"Create new"
							Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_SISW_Eff_RevisionEffOperations",objRevEffectDetails.JavaCheckBox("Create new"),"attached text","Create new")
						End If
						Call Fn_CheckBox_Set("Fn_SISW_Eff_RevisionEffOperations",objRevEffectDetails,"Create new", "ON")
					End If
				End If
				
				' setting effectivity id
				If dicEffectivityMapping("sEffectivityId") <> ""  Then
					Call Fn_Edit_Box("Fn_SISW_Eff_RevisionEffOperations",objRevEffectDetails, "EffectivityID", dicEffectivityMapping("sEffectivityId"))
					if dicEffectivityMapping("bCreateNew") = "" Then
							objRevEffectDetails.JavaEdit("EffectivityID").Activate
					ElseIf  cBool(dicEffectivityMapping("bCreateNew") ) = False Then
							objRevEffectDetails.JavaEdit("EffectivityID").Activate
					End If
					Wait(1)
					If instr(dicEffectivityMapping("sEffectivityId"),"*") > 0 Then
					    If JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("OccurrenceEffectivityDialog").Exist(5) Then
							Set objOccEffID = JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("OccurrenceEffectivityDialog")
						ElseIF JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("OccurrenceEffectivityDialog").Exist(1) Then
							Set objOccEffID = JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("OccurrenceEffectivityDialog")
						End If
						if objOccEffID.exist(20) then
							if dicEffectivityMapping("sVerifyEffectivityId") <> "" Then
								'verify speciifed IDs in list
								aEffID = split(dicEffectivityMapping("sVerifyEffectivityId"),"~")
								iRows = cInt(objOccEffID.JavaTable("EffectivityIDTable").getROProperty("rows"))
								for iCount = 0 to uBound(aEffID)
									bFlag = False
									for iCnt = 0 to iRows - 1
										if cstr(objOccEffID.JavaTable("EffectivityIDTable").getCellData(iCnt, "ID")) = aEffID(iCount) Then
											bFlag = True
											exit for
										End If
									Next
									if bFlag = False then
										exit for
									End IF
								Next
								if bFlag = False then
									'click on cancel
									Call Fn_Button_Click("Fn_SISW_Eff_RevisionEffOperations", objOccEffID,"Close")
									Fn_SISW_Eff_RevisionEffOperations = False
									Set objOccEffID = nothing
									Exit function
								End If
							End IF ' end of if dicEffectivityMapping("sVerifyEffectivityId") <> "" Then
							if dicEffectivityMapping("sSelectEffectivityId") <> "" then
								'select specified ID from list
								iRows = cInt(objOccEffID.JavaTable("EffectivityIDTable").getROProperty("rows"))
								bFlag = False
								for iCnt = 0 to iRows - 1
									if cstr(objOccEffID.JavaTable("EffectivityIDTable").getCellData(iCnt, "ID")) = dicEffectivityMapping("sSelectEffectivityId") Then
										'objOccEffID.JavaTable("EffectivityIDTable").SelectRow iCnt
										Call Fn_SISW_UI_JavaTable_Operations("Fn_SISW_Eff_RevisionEffOperations", "SelectRow", objOccEffID.JavaTable("EffectivityIDTable") , "", "", "", iCnt, "", "", "", "")
										bFlag = True
										exit for
									End If
								Next
								if  bFlag <> True then
									Call Fn_Button_Click("Fn_SISW_Eff_RevisionEffOperations", objOccEffID,"Close")
									Fn_SISW_Eff_RevisionEffOperations = False
									Set objOccEffID = nothing
									Exit function
								End If
							End IF
							Call Fn_Button_Click("Fn_SISW_Eff_RevisionEffOperations", objOccEffID,"OK")
						End If ' end of if objOccEffID.exist(20) then
					End IF
				End If ' end of If dicEffectivityMapping("sEffectivityId") <> ""  Then
			'Effectivity Protection
			If dicEffectivityMapping("bEffectivityProtection") <> ""  Then
				If cInt(objRevEffectDetails.JavaCheckBox("ApplyAccessManagerEffectivityProtection").GetROProperty("value")) = 1 AND cBool(dicEffectivityMapping("bEffectivityProtection") ) = True Then
				ElseIf cInt(objRevEffectDetails.JavaCheckBox("ApplyAccessManagerEffectivityProtection").GetROProperty("value")) = 0 AND cBool(dicEffectivityMapping("bEffectivityProtection") ) = False Then
				Else
					Fn_SISW_Eff_RevisionEffOperations = False
					Exit function
				End If
			End If
			'verifying end item id
			If  dicEffectivityMapping("sEndItemSelectType") = "" Then
				If dicEffectivityMapping("sEndItem") <> "" Then
					If objRevEffectDetails.JavaEdit("EndItem").GetROProperty("value") <> dicEffectivityMapping("sEndItem") then
						Fn_SISW_Eff_RevisionEffOperations = False
						Exit function
					End If
					
					'verifying end item rev.
					If dicEffectivityMapping("sEndItemRev") <> "" Then
						If objRevEffectDetails.JavaList("EndItemRevId").GetROProperty("value") <> dicEffectivityMapping("sEndItemRev")  then
							Fn_SISW_Eff_RevisionEffOperations = False
							Exit function
						End If
					End If
				End If
			Else

			' setting unit radio button
			If dicEffectivityMapping("sUnit") <> ""  Then
				If dicEffectivityMapping("sUnit") <> objRevEffectDetails.JavaEdit("Units").GetROProperty("value") Then
					Fn_SISW_Eff_RevisionEffOperations = False
					Exit function
				End If
			End If
			End If
			' verifying date radio
			If  dicEffectivityMapping("sStartDates") <>"" AND dicEffectivityMapping("sEndDates") <> ""  Then
				' verifying date
				aStartDate  = split(dicEffectivityMapping("sStartDates"),"~")
				aEndDate =  split(dicEffectivityMapping("sEndDates"),"~")

				Set objTable = objRevEffectDetails.JavaTable("DateRangeTable")
				For iCount = 0 to uBound(aStartDate)
					If  objTable.GetCellData(iCount,"From Date") <> trim(aStartDate(iCount)) AND  objTable.GetCellData(iCount,"To Date") <> trim(aEndDate(iCount)) Then
						Fn_SISW_Eff_RevisionEffOperations = False
						Set objTable = nothing
						Exit function
					End If
				Next
			End If
		
			'clciking on Cancel
			Call Fn_Button_Click("Fn_SISW_Eff_RevisionEffOperations",objRevEffectDetails,"Cancel")
			If objRevEffectivity.exist(1) Then
				' closing effectivity dialog window
				Call Fn_Button_Click("Fn_SISW_Eff_RevisionEffOperations", objRevEffectivity,"Close")
			End If
			Fn_SISW_Eff_RevisionEffOperations = True
	'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		Case Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Function [ Fn_SISW_Eff_RevisionEffOperations ] Invalid case [ "& sAction & " ] .")
				Fn_SISW_Eff_RevisionEffOperations = False
	End Select
	Set objRevEffectivity = nothing
	Set objRevEffectDetails = nothing
	Set objTcDefaultWindow = nothing
	Set objJavaApplet  = nothing
	Set objReleaseStatEff = nothing
	Set objReleaseStatEff2 = nothing
			
End Function
