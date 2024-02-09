Option Explicit

'=======================================================================================================================================================
' Function List
'-------------------------------------------------------------------------------------------------------------------------------------------------------
'						Function Name																		|					Created By
'-------------------------------------------------------------------------------------------------------------------------------------------------------
'1. Fn_SISW_LifeView_GetObject()											 |	Vallari S (vallari.shimpukade@siemens.com)
'2. Fn_SISW_LifeView_SaveSession()											|	Vallari S (vallari.shimpukade@siemens.com)
'3. Fn_SISW_LifeView_MenuOperation()									 |	Vallari S (vallari.shimpukade@siemens.com)
'4. Fn_SISW_LifeView_RHSCanvasOperations()							  |	Vallari S (vallari.shimpukade@siemens.com)
'5. Fn_SISW_LifeView_SavePLMXML()										  |	Vallari S (vallari.shimpukade@siemens.com)
'6. Fn_SISW_LifeView_ImgPopupSelect()									  |	Vallari S (vallari.shimpukade@siemens.com)
'7. Fn_SISW_LifeView_3DTextMarkupCreate()							   |	Vallari S (vallari.shimpukade@siemens.com)
'8. Fn_SISW_LifeView_TcVizExit()											    |	Vallari S (vallari.shimpukade@siemens.com)
'9. Fn_SISW_LifeView_LaunchTcViz()										     |	Pritam  S (Pritam.Shikare@ugs.com)
'10. Fn_SISW_LifeView_FileOpenInsertOperation()                       |	  Pritam  S (Pritam.Shikare@ugs.com)
'11. Fn_SISW_LifeView_DialogHandleVerifyMessage()                 |	  Pritam  S (Pritam.Shikare@ugs.com)
'12. Fn_SISW_LifeView_ProductStructureConfigure()                    |	 Pritam  S (Pritam.Shikare@ugs.com)
'13. Fn_SISW_LifeView_TCIntegrationPrefOperation()                   |	 Pritam  S (Pritam.Shikare@ugs.com)
'14. Fn_SISW_LifeView_NetworkLogin()                                       |   Pritam  S (Pritam.Shikare@ugs.com)
'15. Fn_SISW_LifeView_FileUsageConfirmationOperation()            |	 Pritam  S (Pritam.Shikare@ugs.com)
'16. Fn_SISW_LifeView_2DMarkupCreate()                              	 |   Pritam  S (Pritam.Shikare@ugs.com)
'17. Fn_SISW_LifeView_3DMarkupCreate()                             		 |   Pritam  S (Pritam.Shikare@ugs.com)
'18. Fn_SISW_LifeView_PartTrasformation							           |   Pritam  S (Pritam.Shikare@ugs.com)
'19. Fn_SISW_LifeView_PSMandMSM_ProductViewOperations				|   Pritam  S (Pritam.Shikare@ugs.com)
'20. Fn_SISW_LifeView_ViewerTableRowIndex							|   Pritam  S (Pritam.Shikare@ugs.com)
'21. Fn_SISW_LifeView_ViewerBOMTreeTableNodeOperation				|   Pritam  S (Pritam.Shikare@ugs.com)
'22. Fn_SISW_LifeView_ViewPreferencesOperation						| Reema W
'23. Fn_SISW_LifeView_ColorOperations								| Ankit T
'24. Fn_SISW_LifeView_2DMarkupDelete								| Reema W
'25. Fn_SISW_LifeView_MergeSessions									| Reema W
'26. Fn_SISW_LifeView_MarkupLayerTreeNodeOperation									| Reema W
'27. Fn_SISW_LifeView_NavigatorOperations							|Ankit T
'28. Fn_SISW_LifeView_PreferencesOperation							|Ankit T
'29. Fn_SISW_LifeView_GDTMarkupCreate()								|Ankit T
'30. Fn_SISW_LifeView_SaveAsTeamcenterProductView
'31. Fn_SISW_LifeView_AssemblyTreeOperation
'32. Fn_SISW_LifeView_ViewProductStructure
'33. Fn_SISW_LifeView_ExportImage									| Ankit T
'34. Fn_SISW_LifeView_LoadOptionPrefOperation()						| Paresh
'35. Fn_SISW_LifeView_ExportDialogOperation							| Ankit T
'36. Fn_SISW_LifeView_ConceptAppearancePartColorOperation				|Rinki A
'37. Fn_SISW_LifeView_FilePreferencesPLMXMLOperation				|Rinki A
'38. Fn_SISW_LifeView_2DSnapshotFormDialogOperation					| Ankit T
'39. Fn_SISW_LifeView_2DLoaderPreferencesOperations					| Ankit T
'40. Fn_SISW_LifeView_2DMarkupPreferencesOperations					| Ankit T
'41. Fn_SISW_LifeView_AutoFileSearchPreferences()					| Reema W
'42. Fn_SISW_LifeView_Section3DPreferencesOperations()				| Reema W
'43. Fn_SISW_LifeView_Section3D_CreateRepositionSection()			| Reema W
'44. Fn_SISW_LifeView_FilterManager_Operations()						| Shweta Rathod
'45 Fn_SISW_LifeView_3DLoaderPrefOperation()						| Shweta Rathod
'46 Fn_SISW_LifeView_Inspector_Operations()							
'47 Fn_SISW_LifeView_LayerFilter_Operations							| Priyanka Kakade
'48 Fn_SISW_LifeView_DeformationDisplay_Operations					| Priyanka Kakade
'49 Fn_SISW_LifeView_ColorDisplay_Operations						| Priyanka Kakade
'50 Fn_SISW_LifeView_DisplayOptions_Operations						| Priyanka Kakade
'51 Fn_SISW_LifeView_IdentifyDialog_Operations()					| Poonam Chopade
'52 Fn_SISW_LifeView_ExportImage_Save_Ops()							| Poonam Chopade	
'53 Fn_SISW_LifeView_VIZ_ExportDialogOperation()					| Pravin Bhoyar
'54 Fn_SISW_LifeView_AppearanceImagePalette_Ops()					| Poonam Chopade
'55 Fn_SISW_LifeView_ComparisonPreferences_Ops()					| Poonam Chopade
'=======================================================================================================================================================
'****************************************    Function to get Object hierarchy **************************************
'
''Function Name		 	:	Fn_SISW_LifeView_GetObject()
'
''Description		    :  	Function to get specified Object hierarchy.

''Parameters		    :	1. sObjectName : Object Handle name
								
''Return Value		    :  	Object \ Nothing
'
''Examples		     	:	Fn_SISW_LifeView_GetObject("TC_SaveSession")

'History:
'	Developer Name			Date			Rev. No.		Reviewer		Changes Done	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Vallari S			 14-Feb-2013			1.0				
'-----------------------------------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------------------------------
'	Anurag Khera		 26-Sept-2013		2.0								Externalized object hierarchies
'-----------------------------------------------------------------------------------------------------------------------------------
Function Fn_SISW_LifeView_GetObject(sObjectName)
	Dim sObjectXMLPath
	sObjectXMLPath = Fn_GetEnvValue("User", "AutomationDir") & "\TestData\AutomationXML\ObjectXML\TcViz_LifecycleView.xml"
	Set Fn_SISW_LifeView_GetObject = Fn_SISW_Setup_GetObjectFromXML(sObjectXMLPath, sObjectName)
		
End Function 

'****************************************    Function to Save Session ***************************************
'
''Function Name		 	:	Fn_SISW_LifeView_SaveSession()
'
''Description		    :  	Function to get specified Object hierarchy.

''Parameters		    :	1. sCalledFrom : Called from Teamcenter or Standalone Viz  (TC or VIZ)
''								  2. sAction		: Save/SaveAs action
''							      3. sStorageLocPref : 
''							      4. SStorageLoc : location   (eg. Home:Newstuff), eg. C:\mainline\Scripts\Test
''							      5. sFileName : Name of the file
''							      6. sCapture :  '
''							      7. bUICheck : UI verification (TRUE or FALSE)
''							      8. sButtons :   buttons name of Save Session or Save Session As dialog
								
''Return Value		    :  	True \ False
'
''Examples		     	:	Call Fn_SISW_LifeView_SaveSession("TC", "SaveAs", "BaseDocument", "", "Test123","", True, "OK")
'								 Call Fn_SISW_LifeView_SaveSession("VIZ", "Save", "Location", "Home:Newstuff", "abc", "", False, "OK")
'								 Call Fn_SISW_LifeView_SaveSession("VIZ", "Save", "", "C:\mainline\Scripts\Test", "abc", "", False, "OK")

'History:
'	Developer Name			Date			Rev. No.		Reviewer		Changes Done	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Vallari S			 14-Feb-2013			1.0		
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Pritam S			14-Feb-2013			   1.1               Vallari            Handled the case for MyComputer and Servers location
'																								   Handled the  cases for AttachTo and SaveAs dialog	
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Shweta Rathod		09-Dec-2014            1.2               Vallari			Added code to Uncheck - Save extended 3D content into PLMXML Preferences
'																											Save inserted models into PLMXML Preferences From PLMXML Prefrence dialog			
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Ankit T			12-Jan-2015			   1.3             Paresh             Change the sequence of code to set Session name in Save Session dialog
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_LifeView_SaveSession(sCalledFrom, sAction, sStorageLocPref, SStorageLoc, sFileName, sOptions, bUICheck, sButtons)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_LifeView_SaveSession"
   Dim objWin, objParent, sMenu, sMenuPath, bReturn, aOptions, aSubOptions, iCount, sDefName
   Dim sTcWebServer, sTcWebBuild, aWebURLPath
   Dim objPLMXMLWin,objDialog,oNWPwd

   Fn_SISW_LifeView_SaveSession = False
   'Added code to Unckeck PlmXml Preferences
   	
	sMenuPath=Fn_LogUtil_GetXMLPath("Viz_Menu")		
	'Get menu
	sMenu = Fn_GetXMLNodeValue(sMenuPath, "FilePrefrencesPLMXML")	
	'bReturn = Fn_MenuOperation("WinMenuSelect", sMenu)	 	
	
	'Invoke Menu from the Application
	Select Case sCalledFrom
		Case "TC"
			'----------Invoke From TC LCV--------------------------------
			bReturn = Fn_MenuOperation("WinMenuSelect", sMenu)
			wait(1)
			Set objPLMXMLWin = Fn_SISW_LifeView_GetObject("TC_PLMXML")
		Case "VIZ"
			'----------Invoke From Standalone TC VIZ----------------
			wait 5
			Window("VizMainWin").Activate micLeftBtn
			bReturn = Fn_SISW_LifeView_MenuOperation("WinMenuSelect", sMenu)
			wait(1)
			Set objPLMXMLWin = Fn_SISW_LifeView_GetObject("VIZ_PLMXML")
	End Select
	If bReturn = False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Operate Menu [" + sMenu + "]")
		Set objWin = Nothing
		Exit Function
	End If
			
	If objPLMXMLWin.Exist(10) Then	
		
		'set PLMXML Tab
		err.clear
		If Instr(objPLMXMLWin.WinTab("MainTab").GetContent, "Save") > 0 Then
			objPLMXMLWin.WinTab("MainTab").Select "Save", micLeftBtn
			If err.number < 0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Set [Save] Tab")
				Set objPLMXMLWin = Nothing
				Exit Function		
			End If
			Wait 1		
			'OFF SaveExtended3D,SaveInserted  options
		
			if objPLMXMLWin.WinCheckBox("SaveExtended3D").GetROProperty("Checked") = "ON" and objPLMXMLWin.WinCheckBox("SaveInserted").GetROProperty("Checked") = "ON" then
				err.clear
				objPLMXMLWin.WinCheckBox("SaveExtended3D").Set "OFF"
				wait 1
				If err.number < 0 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Set [Save extended 3D content into PLMXML Preferences] Option")
					Set objPLMXMLWin = Nothing
					Exit Function		
				End If
				
				err.clear
				objPLMXMLWin.WinCheckBox("SaveInserted").Set "OFF"	
				wait 1
				If err.number < 0 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Set [Save inserted models into PLMXML Preferences] Option")
					Set objPLMXMLWin = Nothing
					Exit Function		
				End If
				
				err.clear
				objPLMXMLWin.WinButton("Apply").Click 5, 5,micLeftBtn
				wait 1
				If err.number < 0 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Click on [Apply] Button of [PLMXML Prefrences] dialog")
					Set objPLMXMLWin = Nothing
					Exit Function		
				End If
			End if
		End If
		
		err.clear
		objPLMXMLWin.WinButton("OK").Click 5, 5,micLeftBtn
		wait 1
		If err.number < 0 Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Click on [OK] Button of [PLMXML Prefrences] dialog")
			Set objPLMXMLWin = Nothing
			Exit Function		
		End If
   else
   		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [PLMXML Prefrences] dialog does not exist")
		Set objPLMXMLWin = Nothing
		Exit Function
   End if
   'End of added code to unchk PLMXML

   'Select the application, TC or Standalone TcViz
   Select Case sCalledFrom
   	Case "TC"
		'------ Called From Teamcenter LCV ---------------------------------
   		Set objParent = Fn_SISW_LifeView_GetObject("LifeViewWin")
   		Set objWin = Fn_SISW_LifeView_GetObject("TC_SaveSession")
		
   	Case "VIZ"
		'------ Called From Standalone TcVIZ ---------------------------------
   		Set objParent = Fn_SISW_LifeView_GetObject("VizMainWin")
   		Set objWin = Fn_SISW_LifeView_GetObject("VIZ_SaveSession")
   End Select       '///End of Select Statement

	'Set the Object, either Session Save As or Session Save As 
	If Not objWin.Exist(5) Then  'Block 1

		'If Session Save As doesnt Exist then Check for the Session Save
		objWin.SetTOProperty "text","Session Save"

		'If Session Save  doesnt Exist then invoke the menu
		If Not objWin.Exist(5) Then   'Block 2
			'Find File Path for Lifecycle Viewer Menu XML
			sMenuPath=Fn_LogUtil_GetXMLPath("Viz_Menu")

			'Select the Action . Wether Save Session or Save Session As
			Select Case sAction
				Case "Save", "SaveDefNameVerify"
					sMenu = Fn_GetXMLNodeValue(sMenuPath, "SaveSession")
				Case "SaveAs", "SaveAsDefNameVerify"
					sMenu = Fn_GetXMLNodeValue(sMenuPath, "SaveSessionAs")
			End Select
	
			wait(2)
			'Invoke Menu from the Application
			Select Case sCalledFrom
				Case "TC"
					'----------Invoke From TC LCV--------------------------------
					bReturn = Fn_MenuOperation("WinMenuSelect", sMenu)
					wait(1)
				Case "VIZ"
					'----------Invoke From Standalone TC VIZ----------------
					bReturn = Fn_SISW_LifeView_MenuOperation("WinMenuSelect", sMenu)
					wait(1)
			End Select
			If bReturn = False Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Operate Menu [" + sMenu + "]")
				Set objWin = Nothing
				Exit Function
			End If

		End If '//End Of Block 2
	End If   '//End Of Block 1

	'Dismiss Warning Dialog, if popped up
	If objParent.Dialog("Warning").Exist(2) Then
		objParent.Dialog("Warning").WinButton("OK").Click 5, 5,micLeftBtn
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Warning: [Warning] Dialog Handled Successfully")
		If sCalledFrom = "TC" Then
			bReturn = Fn_MenuOperation("WinMenuSelect", sMenu)
		ElseIf sCalledFrom = "VIZ" Then
			bReturn = Fn_SISW_LifeView_MenuOperation("WinMenuSelect", sMenu)
		End If
	End If

	'Check for the Save Session of Save Session As dialog
	If Not objWin.Exist(10) Then
		objWin.SetTOProperty "text","Session Save As"
	End If

	'if Doesnt Exist then End the Function
	If Not objWin.Exist(10) Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Save Session dialog does not exist")
		Set objWin = Nothing
		Exit Function
	End If

	'UI Verification
	If CBool(bUICheck) Then
		bReturn = objWin.WinRadioButton("AlternateLocation").GetROProperty("checked")
		If Trim(cstr(bReturn)) <> "ON" Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Check Default Storage Location Preference as [Alternate Location]")
			objWin.WinButton("Cancel").Click 5, 5,micLeftBtn
			Set objWin = Nothing
			Exit Function
		End If
		bReturn = objWin.WinEdit("StorageLoc").GetROProperty("text")
		If Trim(cstr(bReturn)) <> "Newstuff" Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Check Default Storage Location as [Newstuff]")
			objWin.WinButton("Cancel").Click 5, 5,micLeftBtn
			Set objWin = Nothing
			Exit Function
		End If
	End If

	'Select the Storage Loc Preference, Radio Buttons
	If sStorageLocPref <> "" Then 'Block 3
		Select Case sStorageLocPref
			Case "BaseDocument"
				objWin.WinRadioButton("AttachToBaseDocument").Set
			Case "Bomline"
				objWin.WinRadioButton("AttachToSelectedBomline").Set
			Case "Location"
				objWin.WinRadioButton("AlternateLocation").Set
		End Select

		If Err.Number < 0 Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Set Storage Location as [" + sStorageLocPref + "]")
			objWin.WinButton("Cancel").Click 5, 5,micLeftBtn
			Set objWin = Nothing
			Exit Function
		End If

	End If '//End of Block 3

	Select Case sAction

		Case "SaveDefNameVerify", "SaveAsDefNameVerify"
			'==============Case Save And SaveAs=============================
			If sFileName <> "" Then
				objWin.WinObject("SessionTree").DblClick 50, 30, micLeftBtn
				Wait 3
				If objWin.Dialog("ItemName").exist(5) Then
					sDefName = objWin.Dialog("ItemName").WinEdit("SessionName").GetRoProperty("text")
					objWin.Dialog("ItemName").WinButton("OK").Click 5, 5,micLeftBtn
					Wait 2
				End If
	
				If sFileName <> sDefName Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Set File Name as [" + sFileName + "] in the ItemName dialog")
					objWin.WinButton("Cancel").Click 5, 5,micLeftBtn
					Set objWin = Nothing
					Exit Function
				End If
			End If
			'========end of case VerifyDefaultName================================

		Case "Save", "SaveAs"
			'==============Case Save And SaveAs=============================
			'Storage Location, File save to Folder Path
			If SStorageLoc <> "" Then        '--------------------------------  Block 4
		
				'Extract The File Location into an array
				If Instr(1,SStorageLoc,"Home:") > 0 Then '// Block 5
					'If the Home Folder exist in location then Save on to the Servers
					aStorageLoc = Split(SStorageLoc,":",-1,1)
					sSaveOnTo = "Servers"
				Else
					aStorageLoc = Split(SStorageLoc,"/",-1,1)
					'Save on to the My Computer
					sSaveOnTo = "My Computer"
				End If     '// Block 5 ends
		
				'Click the Browse Button
				If instr(SStorageLoc, objWin.WinEdit("StorageLoc").GetROProperty("text")) > 0 Then
					'Do Nothing
				Else
					objWin.WinButton("Browse").Click 5, 5,micLeftBtn
					If Err.Number < 0 Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to click [ Browse ] button")
						Set objWin = Nothing
						Exit Function
					End If
			
					'If the Attach To Dialog Exists the Preform Operation On the Attach to Dialog
					If objWin.Dialog("Attach to").Exist(10)  Then '// Block 6
						'Select the Save onto Case
						Select Case sSaveOnTo
							Case "Servers"
								'===========Save Onto the Server Location==============================
							   If  instr(objWin.Dialog("Attach to").WinComboBox("LookIn").GetROProperty("Selection"), "Home") = 0 Then  '// Block 7
									'If the Home foldr is not by default selected then select the folder, click on the Server on the LHS pane
									objWin.Dialog("Attach to").WinButton("Servers").Click 5,5,micLeftBtn
									If Err.Number < 0 Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to click [ server ] button")
										Set objWin = Nothing
										Exit Function
									End If
									
									'Get Serevr Path
									aWebURLPath = split(Environment("TcWebServer"), "/")
	                                sTcWebServer = aWebURLPath(2)
									sTcWebBuild = aWebURLPath(3)
			
									'Get the No of count  of items in the List
									iItemsCnt =  objWin.Dialog("Attach to").WinListView("Folders").GetItemsCount()
									For iCount = 0 to cint(iItemsCnt)-1
										sItem = objWin.Dialog("Attach to").WinListView("Folders").GetItem(iCount) 
										'Check if the Servers TcWeb path is present
										If Instr(1,sItem,"/"+ sTcWebServer +":") > 0 and Instr(1, sItem,"/"+ sTcWebBuild +"/")Then
											'if found then activate the node
											objWin.Dialog("Attach to").WinListView("Folders").Activate sItem
											If Err.Number < 0 Then
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Activate the Server Node [ "+sItem+" ]")
												Set objWin = Nothing
												Exit Function
											End If
											Exit For
										End If
									Next
									Wait 3
			
									'if not found then exit from the function
									If iCount = iItemsCnt Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL:  Server Node [ "+sItem+" ] not found in the list")
										Set objWin = Nothing
										Exit Function
									End If
								End If'// Block 7 Ends
			
								'Activate the Full path
								For iCount = 1 to Ubound(aStorageLoc)-1
									objWin.Dialog("Attach to").WinListView("Folders").Activate aStorageLoc(iCount)
									If Err.Number < 0 Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Activate the  Node [ "+aStorageLoc(iCount)+" ]")
										Set objWin = Nothing
										Exit Function
									End If
									Wait 2
								Next
			
								'Select the Destination folder
								objWin.Dialog("Attach to").WinListView("Folders").Select aStorageLoc(iCount)
								If Err.Number < 0 Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Select the  Node [ "+aStorageLoc(iCount)+" ]")
									Set objWin = Nothing
									Exit Function
								End If
								Wait 1
			
								'Click on the Select Button
								objWin.Dialog("Attach to").WinButton("Select").Click 5, 5,micLeftBtn
								If Err.Number < 0 Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Click the [ Select ] button")
									Set objWin = Nothing
									Exit Function
								End If
								Wait 2
							'=========end of Case "Servers"==========
			
							Case "My Computer"
								'===========Case My Computer=======================
								'Click the My Computer on LHS pane
								objWin.Dialog("Attach to").WinButton("MyComputer").Click 5,5,micLeftBtn
								If Err.Number < 0 Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to click [ Browse ] button")
									Set objWin = Nothing
									Exit Function
								End If
			
								'Form the Path
								sPath = ""
								For iCount = 0 to uBound(aStorageLoc)-1
									sPath  = sPath+aStorageLoc(iCount)+"\"
								Next
			
								'Open the Path, by entering the Path in the Edit Box
								objWin.Dialog("Attach to").WinEdit("FolderName").Type sPath
								If Err.Number < 0 Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to click [ Browse ] button")
									Set objWin = Nothing
									Exit Function
								End If
			
								'Click Select Button
								objWin.Dialog("Attach to").WinButton("Select").Click 5, 5,micLeftBtn
								If Err.Number < 0 Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Click the [ Select ] button")
									Set objWin = Nothing
									Exit Function
								End If
								Wait 1
			
								'Select the Destination folder
								objWin.Dialog("Attach to").WinListView("Folders").Select aStorageLoc(iCount)
								If Err.Number < 0 Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Select the  Node [ "+aStorageLoc(iCount)+" ]")
									Set objWin = Nothing
									Exit Function
								End If
								Wait 1
			
								'Click on the Select button
								objWin.Dialog("Attach to").WinButton("Select").Click 5, 5,micLeftBtn
								If Err.Number < 0 Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Click the [ Select ] button")
									Set objWin = Nothing
									Exit Function
								End If
								Wait 2
							'=========end of Case "My Computer"==========
						'End Select
						End Select ' //End the Select Case
					
				ElseIf objWin.Dialog("SaveAs").Exist(5) Then  '// ElseIf of Block 6   (If SaveAs dialog Exists)
				'Select the Case for the SaveOnTo 
					Select Case sSaveOnTo
						Case "Servers"
							'=============Case : Servers  -> Save Onto th Server Location
							If  objWin.Dialog("SaveAs").WinComboBox("LookIn").GetROProperty("Selection") <> "Home" Then ' // Block 8
								'If the Home foldr is not by default selected then select the folder, click on the Server on the LHS pane
								objWin.Dialog("SaveAs").WinButton("Servers").Click 5,5,micLeftBtn
								If Err.Number < 0 Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to click [ Servers ] button")
									Set objWin = Nothing
									Exit Function
								End If
		
								'Activate the TcWeb path of the Server
								aWebURLPath = split(Environment("TcWebServer"), "/")
                                sTcWebServer = aWebURLPath(2)
								sTcWebBuild = aWebURLPath(3)
								iItemsCnt =  objWin.Dialog("SaveAs").WinListView("Folders").GetItemsCount()
								For iCount = 0 to cint(iItemsCnt)-1
									sItem = objWin.Dialog("SaveAs").WinListView("Folders").GetItem(iCount) 
									If Instr(1,sItem,"/"+ sTcWebServer) > 0 and Instr(1, sItem,"/"+ sTcWebBuild +"/")Then
										objWin.Dialog("SaveAs").WinListView("Folders").Activate sItem
										If Err.Number < 0 Then
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Activate the Server Node [ "+sItem+" ]")
											Set objWin = Nothing
											Exit Function
										End If
										Exit For
									End If
								Next
								
								'If System Ask for network password then provide
									Set oNWPwd = Fn_SISW_LifeView_GetObject("VIZ_NetworkPassword")
									wait(5)
									If oNWPwd.Exist(5) Then
										'Set Username
										bReturn = Fn_SISW_LifeView_NetworkLogin(sOptions,"")
										If bReturn = False Then
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to invoke Menu [ "+sMenu+" ]")
											Set oOpenInsertFile = Nothing
											Set oNWPwd = Nothing
											Exit Function	
										End If
									End If
									Set oNWPwd = Nothing
						
								'If not found then Exit from the Function
								If iCount = iItemsCnt Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL:  Server Node [ "+sItem+" ] not found in the list")
									Set objWin = Nothing
									Exit Function
								End If 
							End If    ' // Block 8 ends
		
							'Activate the Full path
							For iCount = 1 to Ubound(aStorageLoc)
								objWin.Dialog("SaveAs").WinListView("Folders").Activate aStorageLoc(iCount)
								If Err.Number < 0 Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Activate the  Node [ "+aStorageLoc(iCount)+" ]")
									Set objWin = Nothing
									Exit Function
								End If
								Wait 2
							Next
		
							'Specify the File Name
							objWin.Dialog("SaveAs").WinEdit("FileName").Set sFileName
							If Err.Number < 0 Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Specify the  FileName in the SaveAs Dilaog")
								Set objWin = Nothing
								Exit Function
							End If
		
							'Click Save Button
							objWin.Dialog("SaveAs").WinButton("Save").Click 5, 5,micLeftBtn
							If Err.Number < 0 Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Click the [ Save ] button")
								Set objWin = Nothing
								Exit Function
							End If
							Wait 2
							'====Case Server ends==========
		
						Case "My Computer"
							'====Case My Computer     =>  Save onto te Local location ==========
							'Click the My Computer on LHS pane
							objWin.Dialog("SaveAs").WinButton("MyComputer").Click 5,5,micLeftBtn
							If Err.Number < 0 Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to click [ MyComputer ] button")
								Set objWin = Nothing
								Exit Function
							End If
		
							'Specify the File path with the Filename
							objWin.Dialog("SaveAs").WinEdit("FileName").Set ""
							Wait 1
							objWin.Dialog("SaveAs").WinEdit("FileName").Type SStorageLoc+"\"+sFileName
							If Err.Number < 0 Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Type the Filename Location in the Filename field")
								Set objWin = Nothing
								Exit Function
							End If
		
							'Click Save button
							objWin.Dialog("SaveAs").WinButton("Save").Click 5, 5,micLeftBtn
							If Err.Number < 0 Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Click the [ Save ] button")
								Set objWin = Nothing
								Exit Function
							End If
							Wait 2
					End Select '//End Select Case
		
				Else   '//  Else Part Block 6 Ends
					'End Function
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL:SaveAs Dialog or Atttach to dialog do not exists")
					Set objWin = Nothing
					Exit Function
					
				End If  '//  Block 6 Ends
			End If
			
			'Activate the Tree noe to Edit the Name
			If sFileName <> "" Then
				Set objDialog = Window("VizMainWin").Dialog("SessionSaveAs").Dialog("Error/Warning")
				objDialog.SetToProperty "text","Warning"
				If objDialog.Exist(1) <> TRUE Then				'Added to perform operation only if Warning Dialog does not comes - By Ankit T[11.2 Porting 29 Apr 15]
					objWin.WinObject("SessionTree").DblClick 50, 30, micLeftBtn
					Wait 3
					If objWin.Dialog("ItemName").exist(5) Then
						objWin.Dialog("ItemName").WinEdit("SessionName").Set sFileName
						objWin.Dialog("ItemName").WinButton("OK").Click 5, 5,micLeftBtn
						Wait 2
					End If
		
					If Err.Number < 0 Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Set File Name as [" + sFileName + "] in the ItemName dialog")
						objWin.WinButton("Cancel").Click 5, 5,micLeftBtn
						Set objWin = Nothing
						Exit Function
					End If
				End If
			End If
			
		End If   '// Block 4 Ends

	End Select '///End of Action Select Case

	
	'Select teh Capture Option or Save Session As Package
	If sOptions <> "" Then
		aOptions  = Split(sOptions,"$",-1,1)
		For iCount = 0 to Ubound(aOptions)
			aSubOptions = Split(aOptions(iCount),"~",-1,1)
			Select Case aSubOptions(0)
				Case "SaveAsPackage"
						objWin.WinCheckBox("SaveAsSessionPackage").WaitProperty "enabled", 1, 20000
						objWin.WinCheckBox("SaveAsSessionPackage").Set aSubOptions(1)
						If Err.Number < 0 Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Set the SaveAsSessionPackage checkbox [ "+aSubOptions(1)+" ] ")
							Set objWin = Nothing
							Exit Function
					End If
				Case "CaptureStatic"
						If objWin.WinCheckBox("CaptureStatic").CheckProperty("enabled", True) Then
							objWin.WinCheckBox("CaptureStatic").Set aSubOptions(1)
							If Err.Number < 0 Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Set the CaptureStatic checkbox [ "+aSubOptions(1)+" ] ")
								Set objWin = Nothing
								Exit Function
							End If
						End If
			End Select
		Next
	End If

	'Click on the Save Button
	If sButtons<>"" Then
		aButtons = Split(sButtons,":",-1,1)
		For iCount = 0 to Ubound(aButtons)
			objWin.WinButton(aButtons(iCount)).Click 5, 5,micLeftBtn
			wait(2)
			If Err.Number < 0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to click [ "+aButtons(iCount)+" ] button")
				Set objWin = Nothing
				Exit Function
			End If
		Next
	End If

	Fn_SISW_LifeView_SaveSession = True
	Set objParent = Nothing
	Set objWin = Nothing
	Set objPLMXMLWin = Nothing

End Function

'****************************************    Function to Select Standalone Viz Menu ***************************************
'
''Function Name		 	:	Fn_SISW_LifeView_MenuOperation()
'
''Description		    :  	Function to get specified Object hierarchy.

''Parameters		    :	1. StrAction : Menu Selection Action
''						:	2. StrMenuLabel	: Menu to be Selected
								
''Return Value		    :  	True \ False
'
''Examples		     	:	Fn_SISW_LifeView_MenuOperation("WinMenuSelect", "File:Save")
'							sValue = Fn_SISW_LifeView_MenuOperation("GetItemProperty", "Tools:Markup~checked")
'                           Fn_SISW_LifeView_MenuOperation("CheckItemProperty", "Tools:Markup~enabled~True")
'                           Fn_SISW_LifeView_MenuOperation("WinMenuExist", "File:Save")
'							Fn_SISW_LifeView_MenuOperation("SelectMenuwithColon", "Window$1 shuttle.hpg:1") Note: For this case use '$' as Menu seperator not ':' or ';'

'History:
'	Developer Name			Date			Rev. No.		Reviewer		Changes Done	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Vallari S			 14-Feb-2013			1.0				
'	Pritam S			 10-Sep-2013			1.1	                        Added Cases GetItemProperty, CheckItemProperty, WinMenuExist
'	Pritam S			 17-Sep-2013			1.2							Handle the case if Menu Item name contains ';' semicolon
'	Ankit Tewari		25-Aug-2014				1.3							Added case SelectMenuwithColon to select menu having colon as 'Window$1 shuttle.hpg:1'
'-----------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_LifeView_MenuOperation(strAction, strMenuLabel)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_LifeView_MenuOperation"
	Dim dMenu, strMenu, objWin, arrMenuLabel
	Dim sMenuPath, arrMenu, iCount, jCount, iSubMenuCnt,sTempMenuPath
	
	Fn_SISW_LifeView_MenuOperation = False

	
	If Instr(1,strMenuLabel,"~") Then
		arrMenuLabel = Split(strMenuLabel,"~",-1,1) 
		strMenu = arrMenuLabel(0)
	Else
		strMenu = strMenuLabel
	End If
	
	Set objWin = Fn_SISW_LifeView_GetObject("VizMainWin")
	Set dMenu=description.create()
	dMenu("menuobjtype").value=2
	
	'Added by Vallari - 7-Nov-2014
	If Instr(1,strMenu,"$") and Instr(1,strMenu,";") Then
		strMenu=Replace(strMenu,"$",":")
	End If
	
	'If strAction = "SelectMenuwithColon" Then
		If Instr(1,strMenu,";") Then
			arrMenu = Split(strMenu,":",-1,1)
			sMenuPath = arrMenu(0)
			For iCount = 1 to Ubound(arrMenu)
				bMenuExist = False
				If Instr(1,arrMenu(iCount),";") Then
					iSubMenuCnt= objWin.WinMenu(dMenu).GetItemProperty(sMenuPath,"SubMenuCount")
					For jCount =1 to iSubMenuCnt
						sTempMenuPath = sMenuPath + ";<Item "+cstr(jCount)+">"
						If objWin.WinMenu(dMenu).GetItemProperty(sTempMenuPath,"Label") = arrMenu(iCount) Then
							sMenuPath=sTempMenuPath
							sTempMenuPath=""
							bMenuExist = True
							Exit For
						ElseIf instr(1, Right(arrMenu(iCount), Len(arrMenu(iCount)) -2), Right(objWin.WinMenu(dMenu).GetItemProperty(sTempMenuPath,"Label"), Len(objWin.WinMenu(dMenu).GetItemProperty(sTempMenuPath,"Label")) -2)) > 0 Then
							sMenuPath=sTempMenuPath
							sTempMenuPath=""
							bMenuExist = True
							Exit For
						End If
					Next 
				End If
			Next
			If bMenuExist Then
				strMenu = sMenuPath
			End If
			
		'Else
			strMenu=Replace(strMenu,":",";")
			'bMenuExist = True
		'End If
		ElseIf Instr(1,strMenu,"$") Then
			strMenu=Replace(strMenu,"$",";")
			bMenuExist = True
		Else
			strMenu=Replace(strMenu,":",";")
			bMenuExist = True
		End If

	If bMenuExist = False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL:  Menu Item  not Found")
		Set objWin = Nothing
		Set dMenu=Nothing
		Fn_SISW_LifeView_MenuOperation = False
		Exit Function
	End If

	objWin.winmenu(dMenu).WaitItemProperty strMenu,"exists",true,10

	Select Case strAction
	    '===============Select WinMenu=======================
		Case "WinMenuSelect","SelectMenuwithColon"
			objWin.winmenu(dMenu).Select(strMenu)
			If Err.Number < 0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Select Menu Item  [" + strMenu + "]")
				Set objWin = Nothing
				Exit Function
			Else
				Fn_SISW_LifeView_MenuOperation = True
			End If

		 '===============Get WinMenu Item Property value=======================
		Case "GetItemProperty"
			Fn_SISW_LifeView_MenuOperation = objWin.winmenu(dMenu).GetItemProperty(strMenu,arrMenuLabel(1))

		 '===============Check WinMenu Item Property value=======================
		Case "CheckItemProperty"
			Fn_SISW_LifeView_MenuOperation = objWin.winmenu(dMenu).CheckItemProperty(strMenu,arrMenuLabel(1),arrMenuLabel(2),5)

		 '===============Check Existence of WinMenu Item =======================
		Case "WinMenuExist"
			Fn_SISW_LifeView_MenuOperation = objWin.winmenu(dMenu).CheckItemProperty(strMenu,"exists",True,10)
			
		'===============Select Menu by Pressing Alt<First letter> =======================
		'Added by Vallari - 7-Nov-2014
		Case "KeyPress"
			arrMenu = Split(strMenu,":",-1,1)
			objWin.click 10, 10
			objWin.Type micAltDwn
			For iCount = 1 to Ubound(arrMenu)
				objWin.Type Left(arrMenu(iCount), 1)
			Next
			Fn_SISW_LifeView_MenuOperation = True
			
	End Select
	
	Set objWin = Nothing
End Function


'****************************************    Function to Select Standalone Viz Menu ***************************************
'
''Function Name		 	:	Fn_SISW_LifeView_Open2D3DDocument()
'
''Description		    :  	Function to get specified Object hierarchy.

''Parameters		    :	1. sObjectName : Object Handle name
								
''Return Value		    :  	True \ False
'
''Examples		     	:	Fn_SISW_LifeView_Open2D3DDocument(sCalledFrom, sOpenOpt, bMarkup)

'History:
'	Developer Name			Date			Rev. No.		Reviewer		Changes Done	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Vallari S			 14-Feb-2013			1.0				
'-----------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_LifeView_Open2D3DDocument(sCalledFrom, sOpenOpt, bMarkup)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_LifeView_Open2D3DDocument"
	Dim bReturn
	Dim objWin
	Fn_SISW_LifeView_Open2D3DDocument = False
	
	Select Case sCalledFrom
		Case "TC"
	   		Set objWin = Fn_SISW_LifeView_GetObject("TC_LoadDoc")
	   	Case "VIZ"
	   		Set objWin = Fn_SISW_LifeView_GetObject("VIZ_LoadDoc")
	End Select
	
	Select Case sOpenOpt
		Case "Open"
			objWin.WinRadioButton("DocOption").SetTOProperty "text", "Open document"
			wait(1)
			objWin.WinRadioButton("DocOption").Set
		Case "Insert"
			objWin.WinRadioButton("DocOption").SetTOProperty "text", "Insert document into active window"
			wait(1)
			objWin.WinRadioButton("DocOption").Set			
	End Select
	
	If Err.Number < 0 Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Set Open Options as [" + sOpenOpt + "]")
		objWin.WinButton("Cancel").Click 5, 5,micLeftBtn
		Set objWin = Nothing
		Exit Function
	End If
	
	If trim(bMarkup) <> "" Then
		If objWin.WinCheckBox("OpenWithMarkups").GetROProperty("enabled") = 1 Then
			If cbool(bMarkup) Then
				objWin.WinCheckBox("OpenWithMarkups").Set "ON"
			Else
				objWin.WinCheckBox("OpenWithMarkups").Set "OFF"
			End If
			
			If Err.Number < 0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Set Markup Options as [" + Cstr(bMarkup) + "]")
				objWin.WinButton("Cancel").Click 5, 5,micLeftBtn
				Set objWin = Nothing
				Exit Function
			End If
				
		Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Markup Option is Disabled")
			objWin.WinButton("Cancel").Click 5, 5,micLeftBtn
			Set objWin = Nothing
			Exit Function			
		End If
	End If
	
	objWin.WinButton("OK").Click 5, 5,micLeftBtn
	If Err.Number < 0 Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Click [OK] Button")
		'objWin.WinButton("Cancel").Click 5, 5,micLeftBtn
		Set objWin = Nothing
		Exit Function
	End If
	
	Fn_SISW_LifeView_Open2D3DDocument = True
End Function

'****************************************    Function to Select Standalone Viz Menu ***************************************
'
''Function Name		 	:	Fn_SISW_LifeView_RHSCanvasOperations()
'
''Description		    :  	Function to do operations on Canvas.

''Parameters		    :	1. sAction : Action to be performed
'							2. arrParam : Parameter array
								
''Return Value		    :  	True \ False
'
''Examples		     	:	Fn_SISW_LifeView_RHSCanvasOperations("PopupMenuSelect", arrParam("Fit All"))

'History:
'	Developer Name			Date			Rev. No.		Reviewer		Changes Done	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Vallari S			 14-Feb-2013			1.0				
'-----------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_LifeView_RHS3DViewerOperations(sAction, arrParam)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_LifeView_RHS3DViewerOperations"
	Dim bReturn
	Dim objWin3dViewer
	
	Fn_SISW_LifeView_RHS3DViewerOperations = False
	
	Set objWin3dViewer = Fn_SISW_LifeView_GetObject("J3DViewer")
	
	Select Case sAction
		Case "AllOn"
			Err.clear
			objWin3dViewer.Object.loadAll()
			If Err.Number < 0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Load 3DViewer Panel")
			Else
				Fn_SISW_LifeView_RHS3DViewerOperations = true
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Successfully Loaded 3DViewer Panel")
			End If
	End Select	
	
	Set objWin3dViewer = Nothing
End Function

'****************************************    Function to Save PLMXML ***************************************
'
''Function Name		 	:	Fn_SISW_LifeView_SavePLMXML()
'
''Description		    :  	Function to Save PLMXML

''Parameters		    :	1. sCalledFrom : TC/VIZ
'							2. sPLMXML		: Save PLMXML Checkbox
								
''Return Value		    :  	True \ False
'
''Examples		     	:	Fn_SISW_LifeView_SavePLMXML("TC", "ON", "", "", "", "", "", "Home:Newstuff", "")

'History:
'	Developer Name			Date			Rev. No.		Reviewer		Changes Done	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Vallari S			 14-Feb-2013			1.0		
'   Pritam S.		    15-Jul-2013              1.1              Vallari S.          Handle the cases for the Servers and My Computers Location  					
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_LifeView_SavePLMXML(sCalledFrom, sPLMXML, sSaveInserted, sCopy, sRetainRef, sIncludeLate, sAlwaysAsk, sLocation, sName)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_LifeView_SavePLMXML"
	Dim objPLMXMLWin, objSaveWin
	Dim sMenuPath, sMenu, sSaveOnTo
	Dim bReturn
	Dim arrPath,stemp,TCSite
	Dim iCnt
	
	Fn_SISW_LifeView_SavePLMXML = False
	If InStr(sLocation,"Home:") OR sLocation = "Home" Then
		arrPath = split(sLocation, ":", -1, 1)
		sSaveOnTo = "Servers"
	Else
		arrPath = split(sLocation, "/", -1, 1)
		sSaveOnTo = "My Computer"
	End If
	
	'Get all the required window references
	Select Case sCalledFrom
		Case "TC"
			Set objPLMXMLWin = Fn_SISW_LifeView_GetObject("TC_PLMXML")
			Set objSaveWin = Fn_SISW_LifeView_GetObject("TC_SaveFile")
		Case "VIZ"
			Set objPLMXMLWin = Fn_SISW_LifeView_GetObject("VIZ_PLMXML")
			Set objSaveWin = Fn_SISW_LifeView_GetObject("VIZ_SaveFile")
	End Select
	
	If Not objSaveWin.Exist(5) Then		
		'Get Viz menu file path
		sMenuPath=Fn_LogUtil_GetXMLPath("Viz_Menu")
		
		'Get menu
	    sMenu = Fn_GetXMLNodeValue(sMenuPath, "SavePLMXML")
	
		wait(2)
		Select Case sCalledFrom
			Case "TC"
				bReturn = Fn_MenuOperation("WinMenuSelect", sMenu)
				wait(1)
			Case "VIZ"
				bReturn = Fn_SISW_LifeView_MenuOperation("WinMenuSelect", sMenu)
				wait(3)
		End Select
		
		If bReturn = False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Operate Menu [" + sMenu + "]")
			Set objPLMXMLWin = Nothing
			Set objSaveWin = Nothing
			Exit Function
		End If
	End If
	
'	If Not objPLMXMLWin.Exist(5) Then
'		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Open Dialog [Save PLMXML]")
'		Set objPLMXMLWin = Nothing
'		Set objSaveWin = Nothing
'		Exit Function
'	End If

	If objPLMXMLWin.Exist(5) Then	
		'set PLMXML Tab
		err.clear
		objPLMXMLWin.WinTab("MainTab").Select "Save", micLeftBtn
		If err.number < 0 Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Set [Save] Tab")
			Set objPLMXMLWin = Nothing
			Set objSaveWin = Nothing
			Exit Function		
		End If
		Wait 1
		
		'Check ON/OFF selected options
		If sPLMXML <> "" Then
			err.clear
			If sPLMXML = "ON" Then
				If objPLMXMLWin.WinCheckBox("SaveExtended3D").GetROProperty("checked") = "OFF" Then
					objPLMXMLWin.WinCheckBox("SaveExtended3D").Click 2,2,micLeftBtn
				End If
			Else
				If objPLMXMLWin.WinCheckBox("SaveExtended3D").GetROProperty("checked")= "ON" Then
					objPLMXMLWin.WinCheckBox("SaveExtended3D").Click 2,2,micLeftBtn
				End If
			End If
			If err.number < 0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Set [Save extended 3D content into PLMXML] Option")
				Set objPLMXMLWin = Nothing
				Set objSaveWin = Nothing
				Exit Function		
			End If
		End If
		Wait 1
	
		If sSaveInserted <> "" Then
			err.clear
			objPLMXMLWin.WinCheckBox("SaveInserted").Set sSaveInserted
			If err.number < 0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Set [Save inserted models] Option")
				Set objPLMXMLWin = Nothing
				Set objSaveWin = Nothing
				Exit Function		
			End If
		End If
		Wait 1
	
		If sCopy <> "" Then
			err.clear
			objPLMXMLWin.WinCheckBox("CopyParts").Set sCopy
			If err.number < 0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Set [Copy parts locally] Option")
				Set objPLMXMLWin = Nothing
				Set objSaveWin = Nothing
				Exit Function		
			End If
		End If
		Wait 1
	
		If sRetainRef <> "" Then
			err.clear
			objPLMXMLWin.WinCheckBox("RetainReferences").Set sRetainRef
			If err.number < 0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Set [Retain references to original product structure] Option")
				Set objPLMXMLWin = Nothing
				Set objSaveWin = Nothing
				Exit Function		
			End If
		End If
		Wait 1
		
		If sIncludeLate <> "" Then
			err.clear
			objPLMXMLWin.WinCheckBox("IncludeLateLoaded").Set sIncludeLate
			If err.number < 0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Set [Include late loaded attributes] Option")
				Set objPLMXMLWin = Nothing
				Set objSaveWin = Nothing
				Exit Function		
			End If
		End If
		Wait 1
		
		If sAlwaysAsk <> "" Then
			err.clear
			objPLMXMLWin.WinCheckBox("AlwaysAsk").Set sAlwaysAsk
			If err.number < 0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Set [ALways ask at save time] Option")
				Set objPLMXMLWin = Nothing
				Set objSaveWin = Nothing
				Exit Function		
			End If
		End If
		Wait 1
		
		err.clear
		objPLMXMLWin.WinButton("OK").Click
		If err.number < 0 Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Click [OK] Button")
			Set objPLMXMLWin = Nothing
			Set objSaveWin = Nothing
			Exit Function		
		End If
		Wait 2
	End If
	
	If objSaveWin.exist(20) Then
		Select Case sSaveOnTo
			'-------------------Save on to servers database-----------------------------
			Case  "Servers"
				If sCalledFrom <> "TC" Then
					'Implement the code, Click the servers button on the LHS  side
					'if Login dialog apperas login to TC
				End If

                If Instr(objSaveWin.WinComboBox("LookIn").GetROProperty("Selection"), "Home") <= 0 Then  '// Block 7
						'If the Home foldr is not by default selected then select the folder, click on the Server on the LHS pane
						objSaveWin.WinButton("Servers").Click 5,5,micLeftBtn
						If Err.Number < 0 Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to click [ Browse ] button")
							Set objSaveWin = Nothing
							Exit Function
						End If
						wait 5
						'Get the No of count  of items in the List
						iItemsCnt =  objSaveWin.WinListView("WinListView").GetItemsCount()
						For iCount = 0 to cint(iItemsCnt)-1
							sItem = objSaveWin.WinListView("WinListView").GetItem(iCount) 
							'Check if the Servers TcWeb path is present
							stemp = split(Environment("TcWebServer"),"/")
							TCSite = Replace("http://server:80/webuild/","server",Environment("TcServer"))
							TCSite = Replace(TCSite,"webuild",stemp(3))
							If (Instr(1,Lcase(sItem),Lcase("/"+Environment("TcServer")+":")) > 0 and Instr(1, Lcase(sItem),LCase("/"+Environment("TcWebBuild")+"/"))) OR Instr(1,Lcase(Environment("TcWebBuild")),Lcase(sItem)) > 0 Then
								'if found then activate the node
								objSaveWin.WinListView("WinListView").Activate sItem
								If Err.Number < 0 Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Activate the Server Node [ "+sItem+" ]")
									Set objSaveWin = Nothing
									Exit Function
								End If
								Exit For
							ElseIf(Instr(Lcase(sItem),TCSite) >0) Then
								objSaveWin.WinListView("WinListView").Activate sItem
								If Err.Number < 0 Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Activate the Server Node [ "+sItem+" ]")
									Set objSaveWin = Nothing
									Exit Function
								End If
								Exit For
							End If
						Next
						Wait 3

						'if not found then exit from the function
						If iCount = iItemsCnt Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL:  Server Node [ "+sItem+" ] not found in the list")
							Set objSaveWin = Nothing
							Exit Function
						End If
				End If'// Block 7 Ends

				For iCnt = 1 To Ubound(arrPath) Step 1
					err.clear
					'Activate the Path of the Location
					objSaveWin.WinListView("WinListView").Activate arrPath(iCnt), micLeftBtn
					If err.number < 0 Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Select [" + arrPath(iCnt) + "] from ListView")
						Set objPLMXMLWin = Nothing
						Set objSaveWin = Nothing
						Exit Function			
					End If
					Wait 1
				Next

				'Set the File name
				If sName <> "" Then
					err.clear
					objSaveWin.WinEdit("FileName").Set sName
					If err.number < 0 Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Set Name as [" + sName + "]")
						Set objPLMXMLWin = Nothing
						Set objSaveWin = Nothing
						Exit Function			
					End If	
					Wait 1		
				End If

				'Click Save
				err.clear
				objSaveWin.WinButton("Save").Click 5,5,micLeftBtn
				If err.number < 0 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Click [Save] Button")
					Set objPLMXMLWin = Nothing
					Set objSaveWin = Nothing
					Exit Function			
				End If	
				Wait 1
			'--------------Save on the Local Hard Drive------------------------
			Case "My Computer"
				'Set the File path
				objSaveWin.WinEdit("FileName").Set sLocation+"/"+sName
				If err.number < 0 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Set Name as [" + sName + "]")
					Set objPLMXMLWin = Nothing
					Set objSaveWin = Nothing
					Exit Function			
				End If	
				Wait 1

				'Click the Save button
				objSaveWin.WinButton("Save").Click 5,5,micLeftBtn
				If err.number < 0 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Click [Save] Button")
					Set objPLMXMLWin = Nothing
					Set objSaveWin = Nothing
					Exit Function			
				End If	
				Wait 1

		End Select
				
	Else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Find [Save File] Dialog")
		Set objPLMXMLWin = Nothing
		Set objSaveWin = Nothing
		Exit Function		
	End If
	
	Set objPLMXMLWin = Nothing
	Set objSaveWin = Nothing
	Fn_SISW_LifeView_SavePLMXML = True	

End Function
'****************************************    Function to Select Image Area Popup Menu ***************************************
'
''Function Name		 	:	Fn_SISW_LifeView_ImgPopupSelect()
'
''Description		    :  	Function to Select Image Area Popup Menu

''Parameters		    :	1. sCalledFrom : TC/VIZ
'							2. sMenu		: Popup Menu to be Selected
								
''Return Value		    :  	True \ False
'
''Examples		     	:	Fn_SISW_LifeView_ImgPopupSelect("VIZ", "AllOn")

'History:
'	Developer Name			Date			Rev. No.		Reviewer		Changes Done	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Vallari S			 11-Jun-2013			1.0				
'-----------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_LifeView_ImgPopupSelect(sCalledFrom, sMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_LifeView_ImgPopupSelect"
	Dim breturn
	Dim objWin
	
	Dim sLeft, strTop, sRight, sBottom, xParentCord, yParentCord
	
	Fn_SISW_LifeView_ImgPopupSelect = False
	
	Select Case sCalledFrom
		Case "TC"
			Set objWin = Fn_SISW_LifeView_GetObject("LifeViewWin")
		Case "Viz"
			Set objWin = Fn_SISW_LifeView_GetObject("VIZ_ScratchPad")
	End Select
	
	Select Case sMenu
		Case "AllOn"
			sMenu = "All On (Ctrl-Break to cancel)"
	End Select
	
	Err.clear
	
	If objWin.Exist(5) Then
		objWin.WinObject("DMUtils").Click 50, 50, micRightBtn
		objWin.WinMenu("ContextMenu").Select sMenu
			If Err.Number < 0 Then
			    bReturn = objWin.GetTextLocation(sMenu, sLeft, strTop, sRight, sBottom, False) 
				If bReturn = True Then
					xParentCord = (sLeft+sRight) / 2 
				    yParentCord =(strTop+sBottom) / 2 
				
					wait 3
					objWin.Click xParentCord,yParentCord, micLeftBtn
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Select Popup Menu ["&sMenu&"]")
					Set objWin = Nothing
					Exit Function	
				End If
			End If
		' Pranav:- 25-Aug-2014 ->  Added Temporarily, as winmenu object issue
		'-----------------------------------------------------------------------------------------------
'		bReturn = objWin.GetTextLocation(sMenu, sLeft, strTop, sRight, sBottom, False) 
'		If bReturn = True Then
'			xParentCord = (sLeft+sRight) / 2 
'		    yParentCord =(strTop+sBottom) / 2 
'		
'		wait 3
'		objWin.Click xParentCord,yParentCord, micLeftBtn
		'-----------------------------------------------------------------------------------------------
		
	End If
	
	Fn_SISW_LifeView_ImgPopupSelect = True
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Successfully Selected Popup Menu")
	
End Function

Public Function Fn_SISW_LifeView_DialogOperation(sCalledFrom, sMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_LifeView_ImgPopupSelect"
	Dim breturn
	Dim objWin
	
	Dim sLeft, strTop, sRight, sBottom, xParentCord, yParentCord
	
	Fn_SISW_LifeView_DialogOperation = False
	
	Select Case sCalledFrom
		Case "TC"
			Set objWin=Window("LifeViewWin")
		Case "Viz"
			Set objWin = Fn_SISW_LifeView_GetObject("VIZ_ScratchPad")
	End Select
	
	Select Case sMenu
		Case "Snap"
			sMenu = "Save as Teamecenter Snapshot..."
	End Select
	
	Err.clear
	
	If objWin.Exist(5) Then
		JavaWindow("LCV_localWin").InsightObject("InsightObject_snap2").Click 50, 50, micRightBtn
		
'		objWin.WinMenu("ContextMenu").Select sMenu
	
		' Pranav:- 25-Aug-2014 ->  Added Temporarily, as winmenu object issue
		'-----------------------------------------------------------------------------------------------
		bReturn = objWin.GetTextLocation(sMenu, sLeft, strTop, sRight, sBottom, False) 
		If bReturn = False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Select Popup Menu ["&sMenu&"]")
			Set objWin = Nothing
			Exit Function	
		End If
		xParentCord = (sLeft+sRight) / 2 
		yParentCord =(strTop+sBottom) / 2 
		
		wait 3
		objWin.Click xParentCord,yParentCord, micLeftBtn
		'-----------------------------------------------------------------------------------------------
		
	End If
	
	Fn_SISW_LifeView_ImgPopupSelect = True
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Successfully Selected Popup Menu")
	
End Function
'=======================================================================================================================================================
'****************************************    Function to create 3D Text Markup ***************************************
'
''Function Name		 	:	Fn_SISW_LifeView_3DTextMarkupCreate()
'
''Description		    :  	Function to create 3D Text Markup

''Parameters		    :	1. sCalledFrom : TC/VIZ
'							2. sAction		: Base/Advance
'							3. arrParam		: Array of parameters
								
''Return Value		    :  	True \ False
'
''Examples		     	:	Fn_SISW_LifeView_3DTextMarkupCreate("VIZ", "Base", Array("TestMarkup"))

'History:
'	Developer Name			Date			Rev. No.		Reviewer		Changes Done	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Vallari S			 13-Jun-2013			1.0				
'-----------------------------------------------------------------------------------------------------------------------------------
'	Reema W			 	 29-Aug-2014			1.0							added Case "TC_PSE"
'-----------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_LifeView_3DTextMarkupCreate(sCalledFrom, sAction, arrParam)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_LifeView_3DTextMarkupCreate"
	Dim bReturn
	Dim objWin, objDialog, objDMUtil
	Dim dMenu
	Dim sMenuPath
	Dim sMarkupMenu, sTextMarkupMenu
	
	Fn_SISW_LifeView_3DTextMarkupCreate = False
	
	Select Case sCalledFrom
		Case "TC"
			Set objWin = Fn_SISW_LifeView_GetObject("LifeViewWin")
			Set objDMUtil = objWin.WinObject("DMUtils")
			Set objDialog = Fn_SISW_LifeView_GetObject("TC_Markup")
		Case "VIZ"
			Set objWin = Fn_SISW_LifeView_GetObject("VizMainWin")
			Set objDMUtil = Fn_SISW_LifeView_GetObject("VIZ_ScratchPad").WinObject("DMUtils")
			Set objDialog = Fn_SISW_LifeView_GetObject("VIZ_Markup")
		Case "TC_PSE"
			Set objDMUtil = Window("TcVizStructureManager").WinObject("3DImageViewer")
			Set objDialog = Fn_SISW_LifeView_GetObject("PSE_Markup")
	End Select


	If sCalledFrom =  "TC" OR sCalledFrom =  "VIZ" Then
		'Find File Path for Lifecycle Viewer Menu XML
		 sMenuPath=Fn_LogUtil_GetXMLPath("Viz_Menu")
		 
		 'Extract Menu Paths from XML
		 sMarkupMenu = Fn_GetXMLNodeValue(sMenuPath, "EnableMarkup")
		 sTextMarkupMenu = Fn_GetXMLNodeValue(sMenuPath, "TextMarkup")
		 
		 sMarkupMenu = Replace(sMarkupMenu, ":", ";")
		 sTextMarkupMenu = Replace(sTextMarkupMenu, ":", ";")
		
		'Invoke Marup Dialog by Menu Action
		set dMenu=description.create()
		dMenu("menuobjtype").value=2
		bReturn = objWin.winmenu(dMenu).CheckItemProperty(sMarkupMenu, "checked", true)
		If bReturn = False Then
			objWin.winmenu(dMenu).Select sMarkupMenu
			wait(1)
		End If
		objWin.winmenu(dMenu).Select sTextMarkupMenu
	ElseIf sCalledFrom =  "TC_PSE" Then
			If Fn_SISW_TcViz_ToolbarOperation("IsButtonChecked", "3D Markup", "3D Markup") =False Then
					call Fn_SISW_TcViz_ToolbarOperation("ButtonClick", "3D Markup", "3D Markup")
			End If
			If Fn_SISW_TcViz_ToolbarOperation("IsButtonChecked", "3D Markup", "Text") =False Then
				call Fn_SISW_TcViz_ToolbarOperation("ButtonClick", "3D Markup", "Text")
			End If
	End If
	
	objDMUtil.Click 25, 25, micLeftBtn
	
	'Check Existance of Dialog
	If Not objDialog.Exist(10) Then
		Set objWin = Nothing
		Set objDMUtil = Nothing
		Set objDialog = Nothing
		set dMenu=nothing
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to Find [MarkupText] Dialog")
		Exit Function		
	End If
	
	Err.clear
	
	'Base OR Advance creation of Text Markup
	Select Case sAction
		Case "Base"
			objDialog.WinEditor("MarkupText").Type arrParam(0)
			If Err.Number < 0 Then
				Set objWin = Nothing
				Set objDMUtil = Nothing
				Set objDialog = Nothing
				set dMenu=nothing
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to Set Markup Text on [MarupText] Dialog")
				Exit Function
			End If
			objDialog.WinButton("OK").Click 5,5,micLeftBtn
			If Err.Number < 0 Then
				Set objWin = Nothing
				Set objDMUtil = Nothing
				Set objDialog = Nothing
				set dMenu=nothing
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to Click [OK] Button on [MarupText] Dialog")
				Exit Function
			End If
			
		Case "Advance"
		
		Case Else
			Set objWin = Nothing
			Set objDMUtil = Nothing
			Set objDialog = Nothing
			set dMenu=nothing
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to Find Required Case")
			Exit Function			
	End Select
	
	'Disable Markup Menu
	If sCalledFrom =  "TC" OR sCalledFrom =  "VIZ" Then
		objWin.winmenu(dMenu).Select sMarkupMenu	
	ElseIf sCalledFrom =  "TC_PSE" Then
		call Fn_SISW_TcViz_ToolbarOperation("ButtonClick", "3D Markup", "3D Markup")
	End If
	
	Set objWin = Nothing
	Set objDMUtil = Nothing
	Set objDialog = Nothing
	set dMenu=nothing
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Successfully Created 3D Text Markup")
	Fn_SISW_LifeView_3DTextMarkupCreate = True
	
End Function
'=======================================================================================================================================================
'****************************************    Function to create 3D Text Markup ***************************************
'
''Function Name		 	:	Fn_SISW_LifeView_TcVizExit()
'
''Description		    :  	Function to Exit Standalone Viz

''Parameters		    :	1. sButton : Yes/No/Cancel

''Return Value		    :  	True \ False
'
''Examples		     	:	Fn_SISW_LifeView_TcVizExit(sButton)

'History:
'	Developer Name			Date			Rev. No.		Reviewer		Changes Done	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Vallari S			 13-Jun-2013			1.0				
'-----------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_LifeView_TcVizExit(sButton)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_LifeView_TcVizExit"
	Dim objWin
	Dim sMenuPath
	Dim sExitMenu
	Dim dMenu
	
	Fn_SISW_LifeView_TcVizExit = False
	
	Set objWin = Fn_SISW_LifeView_GetObject("VizMainWin")
	
	
	'Find File Path for Lifecycle Viewer Menu XML
     sMenuPath=Fn_LogUtil_GetXMLPath("Viz_Menu")
     
     'Extract Menu Paths from XML
     sExitMenu = Fn_GetXMLNodeValue(sMenuPath, "Exit")
     sExitMenu = Replace(sExitMenu, ":", ";")
     
     Err.Clear
     
       'If Warning Dialog Exists then click the Button
     If objWin.Dialog("Warning").Exist(10) Then
     	objWin.Dialog("Warning").WinButton(sButton).Click 5, 5,micLeftBtn
     	If Err.Number < 0 Then
			Set objWin = Nothing
			set dMenu = Nothing
			Fn_SISW_LifeView_TcVizExit = False
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to Click Button on [Warning] Dialog")
			Exit Function
		Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Successfully Clicked [ OK ] Button on [Warning] Dialog")
			Fn_SISW_LifeView_TcVizExit = True
		End If
'	 Else
'	 	Set objWin = Nothing
'		set dMenu = Nothing
'		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to Find [Warning] Dialog")
'		Exit Function
     End If

     'If Viz is open the Perform File:Exit
     If objWin.Exist(10) Then
     	set dMenu=description.create()
		dMenu("menuobjtype").value=2
		objWin.winmenu(dMenu).WaitItemProperty sExitMenu,"enabled","true",20
		objWin.winmenu(dMenu).Select sExitMenu
		If Err.Number < 0 Then
			Set objWin = Nothing
			set dMenu = Nothing
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to Operate [Exit] Menu")
			Exit Function
		Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Successfully performed [File:Exit ] Menu operation")
			Fn_SISW_LifeView_TcVizExit = True
		End If
	 Else
	 	Set objWin = Nothing
		set dMenu = Nothing
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to Find [Visualization] Window")
		Exit Function	 	
     End If

  	If objWin.Dialog("Warning").Exist(10) Then
     	objWin.Dialog("Warning").WinButton(sButton).Click 5, 5,micLeftBtn
     	If Err.Number < 0 Then
			Set objWin = Nothing
			set dMenu = Nothing
			Fn_SISW_LifeView_TcVizExit = False
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to Click Button on [Warning] Dialog")
			Exit Function
		Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Successfully Clicked [ OK ] Button on [Warning] Dialog")
			Fn_SISW_LifeView_TcVizExit = True
		End If
	'	 Else
	'	 	Set objWin = Nothing
	'		set dMenu = Nothing
	'		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to Find [Warning] Dialog")
	'		Exit Function
     End If
	
	Set objWin = Nothing
	set dMenu = Nothing
	Fn_SISW_LifeView_TcVizExit = True
End Function

'=========================================================================================================================================
'****************************************    Function to Launch the  TcViz application***************************************
'
''Function Name		 	:	Fn_SISW_LifeView_LaunchTcViz()
'
''Description		    :  	Function to invoke Viz

''Parameters		    :	1. bLaunchFromTcRAC : TRUE/FALSE (TRUE  : If TcViz is to be launched from the RAC toolbar button, 
'															                            FALSE :  If TcRAC to be invoked through Systemutil)

''Return Value		    :  	True \ False
'
''Examples		     	:		'If Viz to be launched standalone by SystemUtil
'									Fn_SISW_LifeView_LaunchTcViz(False)

'									If Viz to be launched through teamcenter
'									Fn_SISW_LifeView_LaunchTcViz(True)

'History:
'	Developer Name			Date			Rev. No.		Reviewer		Changes Done	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Pritam S			    10-Jul-2013		   1.0				  Vallari S.
'-----------------------------------------------------------------------------------------------------------------------------------

Public Function Fn_SISW_LifeView_LaunchTcViz(bLaunchFromTcRAC)																																		 
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_LifeView_LaunchTcViz"
   On Error resume next
	Dim sPath, sXmlFilePath, bReturn, sToolbarBtn1, sToolbarBtn2
	Dim oVizScratchPad, oVizWin
	Fn_SISW_LifeView_LaunchTcViz = True

	If bLaunchFromTcRAC  Then
		'##### Launch TcViz from the RAC##############
		'Get the  path of the RAC_Toolbar.xml file
		sXmlFilePath = Fn_LogUtil_GetXMLPath("RAC_Toolbar")
		'Get the Toolbar button name of the 'StartOpenInLifecycleVisualization'
		sToolbarBtn1 = Fn_GetXMLNodeValue(sXmlFilePath, "StartOpenInLifecycleVisualization")

		sXmlFilePath = Fn_LogUtil_GetXMLPath("PSM_Toolbar")
		'Get the Toolbar button name of the 'StartOpenInLifecycleVisualization'
		sToolbarBtn2 = Fn_GetXMLNodeValue(sXmlFilePath, "StartOpenInLifecycleVisualization")
		'Click the Toolbar button to launch, or open the Object in the TcViz
		If  Fn_ToolBarOperation( "ButtonExist", sToolbarBtn1, "") Then
			bReturn = Fn_ToolbarButtonClick_Ext(1,sToolbarBtn1)
		ElseIf Fn_ToolBarOperation( "ButtonExist", sToolbarBtn2, "") Then
			bReturn = Fn_ToolbarButtonClick_Ext(1,sToolbarBtn2)
		Else
			Fn_SISW_LifeView_LaunchTcViz = FALSE
			 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Toolbar button [ "+sToolbarBtn1+"  ] does not exits")
			 Exit Function
		End If

		If bReturn = False Then
			 Fn_SISW_LifeView_LaunchTcViz = FALSE
			 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to click the Toolbar button [ "+sToolbarBtn1+"  ] from Teamcenter")
			 Exit Function
		End If
	
	Else
		'####### Launch TcViz through the SystemUtil #####
		sPath = Environment.Value("VizInstallDir")+"/Products/"+Environment.Value("VizLicenseLevel")+"/VisView.exe"
		SystemUtil.Run sPath
		If Err.Number < 0 Then
			Fn_SISW_LifeView_LaunchTcViz = FALSE
			 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Invoke Tc Viz Application from [" + sPath + "]")
			 Exit Function
		End If
	End If

	'Verify if the TcViz main Window is displayed	
	Set oVizScratchPad = Fn_SISW_LifeView_GetObject("VIZ_ScratchPad")
	Set oVizWin = Fn_SISW_LifeView_GetObject("VizMainWin")
	If bLaunchFromTcRAC  Then		
		If oVizScratchPad.Exist(iTimeOut) Then						  							
				Fn_SISW_LifeView_LaunchTcViz = TRUE  
				oVizWin.Maximize 																								
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Invoked Tc Viz from the Teamcenter")
		Else
				 Fn_SISW_LifeView_LaunchTcViz = FALSE
				 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Invoke Tc Viz from the Teamcenter")
				 Set oVizScratchPad = Nothing
				 Exit Function
		End If
	Else
		If oVizWin.Exist(iTimeOut) Then						  							
				Fn_SISW_LifeView_LaunchTcViz = TRUE 
				oVizWin.Maximize 																						
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Invoked Tc Viz from from [" + sPath + "]")
		Else
				 Fn_SISW_LifeView_LaunchTcViz = FALSE
				 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Invoke Tc Viz Application from [" + sPath + "]")
				 Set oVizWin = Nothing
				 Exit Function
		End If
	End If
	Wait 2

	Set oVizWin = Nothing
	Set oVizScratchPad = Nothing
End Function


'=========================================================================================================================================
'****************************************    Function to Launch the  TcViz application***************************************
'
''Function Name		 	:	Fn_SISW_LifeView_FileOpenInsertOperation()
'
''Description		    :  	Function to invoke Viz

''Parameters		    :	sAction : (Open or Insert)
'                                 dicOpenFile : Dictionary , refer the example given below

''Return Value		    :  	True \ False
'
''Examples		     	:	 Set dicOpenFile = CreateObject("Scripting.Dictionary")
'								  dicOpenFile("OpenFrom") = "Menu"
'								  dicOpenFile("Storage") = "MyComputer"
'								  dicOpenFile("NetworkCredentials") = "AutoTest1:AutoTest1:Engineering:Designer::autotest1"
'								  dicOpenFile("FileFolderPath") = "D:\mainline\Scripts\Reg-Visualization\Testcase"
'								  dicOpenFile("FileName") = "objPart.plmxml"
'  								  Fn_SISW_LifeView_FileOpenInsertOperation("Open",dicFileOpen)

'								  Set dicOpenFile = CreateObject("Scripting.Dictionary")
'								  dicOpenFile("OpenFrom") = "Menu"
'								  dicOpenFile("Storage") = "MyComputer"
'								  dicOpenFile("NetworkCredentials") = "AutoTest1:AutoTest1:Engineering:Designer::autotest1"
'								  dicOpenFile("FileFolderPath") = "D:\mainline\Scripts\Reg-Visualization\Testcase"
'								  dicOpenFile("FileName") = "objPart.plmxml"
'  								  Fn_SISW_LifeView_FileOpenInsertOperation("Insert",dicFileOpen)

'History:
'	Developer Name			Date			  Rev. No.		 Reviewer		Changes Done	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Pritam S			       11-Jul-2013			 1.0		     Vallari
'	Pritam S		           24-Jul-2013           1.1             Vallari          Handle the cases Open/Insert
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function Fn_SISW_LifeView_FileOpenInsertOperation(sAction, dicOpenFile)																																		 
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_LifeView_FileOpenInsertOperation"
   On Error resume next

   Dim sPath, sXmlFilePath, bReturn
   Dim sToolbarBtn, sMenu
   Dim oVizWin, oNWPwd, oOpenFile, oProdStrucDlg ,oAttachDlg
   Dim iCount, iItemsCnt

   Fn_SISW_LifeView_FileOpenInsertOperation = False

	Set oOpenInsertFile = Fn_SISW_LifeView_GetObject("VIZ_OpenFile")


	'Select Case for Action Open/Insert
	Select Case sAction
		Case "Open"
			'Open File
			Set oOpenInsertFile = Fn_SISW_LifeView_GetObject("VIZ_OpenFile")
			sXMLFilePath = Fn_LogUtil_GetXMLPath("Viz_Menu")
			sMenu = Fn_GetXMLNodeValue(sXmlFilePath, "FileOpen")
				'---------------------------------------------------------------------------------------
		Case "Insert"
			'Insert File
			Set oOpenInsertFile = Fn_SISW_LifeView_GetObject("VIZ_OpenFile")
			oOpenInsertFile.SetToProperty "text", "Insert File"
			sXMLFilePath = Fn_LogUtil_GetXMLPath("Viz_Menu")
			sMenu = Fn_GetXMLNodeValue(sXmlFilePath, "FileInsert")
				'---------------------------------------------------------------------------------------
	End Select

	'Select Case for Menu/Toolbar
	Select Case dicOpenFile("OpenFrom")
		case "Toolbar"
			sXMLFilePath = Fn_LogUtil_GetXMLPath("Viz_Toolbar")
			sToolbarBtn = Fn_GetXMLNodeValue(sXmlFilePath, "Open")
			bReturn = Fn_ToolbarButtonClick_Ext(1,sToolbarBtn)
		'---------------------------------------------------------------------------------------
		Case "Menu"
			bReturn = Fn_SISW_LifeView_MenuOperation("WinMenuSelect",sMenu)
			If bReturn = False Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to invoke Menu [ "+sMenu+" ]")
				Set oOpenInsertFile = Nothing
				Exit Function		
			End If
		'---------------------------------------------------------------------------------------	
	End Select

	'If System Ask for network password then provide
	Set oNWPwd = Fn_SISW_LifeView_GetObject("VIZ_NetworkPassword")
	wait(5)
	If oNWPwd.Exist(5) Then
		'Set Username
		bReturn = Fn_SISW_LifeView_NetworkLogin(dicOpenFile("NetworkCredentials"),"")
		If bReturn = False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to invoke Menu [ "+sMenu+" ]")
			Set oOpenInsertFile = Nothing
			Set oNWPwd = Nothing
			Exit Function	
		End If
	End If
	Set oNWPwd = Nothing

	'Check if the Open Dialog box exist
	If oOpenInsertFile.Exist(20) Then
		Select Case dicOpenFile("Storage")

			'------------------Open the File from the Local Hard Drive-----------------------------------
			Case "MyComputer"
				'Open the Folder
				If dicOpenFile("FileFolderPath") <> "" Then
					'Set the File Folder name 
					Err.Clear
					oOpenInsertFile.WinEdit("FileName").Type dicOpenFile("FileFolderPath")
					If Err.Number < 0 Then
						Fn_SISW_LifeView_FileOpenInsertOperation = FALSE
						 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Type the File Folder Name [ "+dicOpenFile("FileFolderPath")+" ] in the FileName filed")
						 Set oOpenInsertFile = Nothing
						 Exit Function
					End If
					Wait 3

					'Click Open button
					oOpenInsertFile.WinButton("Open").Click 5,5,micLeftBtn
					If Err.Number < 0 Then
						 Fn_SISW_LifeView_FileOpenInsertOperation = FALSE
						 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to click [ Open ] button")
						 Set oOpenInsertFile = Nothing
						 Exit Function
					End If
				End If
				Wait 2

				'Open the File
				If dicOpenFile("FileName") <> "" Then
					'Check wether the specific file is present, it present then select and Open
					iItemsCnt = oOpenInsertFile.WinListView("FoldesFilesList").GetItemsCount()
					For iCount =0 to iItemsCnt-1
						If oOpenInsertFile.WinListView("FoldesFilesList").GetItem(iCount) = dicOpenFile("FileName") Then
							'Select the File
							oOpenInsertFile.WinListView("FoldesFilesList").Select iCount, micLeftBtn
							If Err.Number < 0 Then
							     Fn_SISW_LifeView_FileOpenInsertOperation = FALSE
								 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Select the File [ "+dicOpenFile("FileName")+" ]from the List")
								 Set oOpenInsertFile = Nothing
								 Exit Function
							End If
							Wait 1

							'Click the Open Button
							oOpenInsertFile.WinButton("Open").Click 5,5,micLeftBtn
							If Err.Number < 0 Then
							     Fn_SISW_LifeView_FileOpenOperation = FALSE
								 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to click [ Open ] button")
								 Set oOpenInsertFile = Nothing
								 Exit Function
							End If
							Wait 1
							Exit For
						End If
					Next

					'If the File is not found, exit from the function
					If iCount = iItemsCnt Then
						Fn_SISW_LifeView_FileOpenOperation = FALSE
						oOpenInsertFile.WinButton("Cancel").Click 5,5,micLeftBtn
						 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Specified File [ "+dicOpenFile("FileName")+" ] is not found in the Directory ["+dicOpenFile("FileFolderPath")+" ] ")
						 Set oOpenInsertFile = Nothing
						 Exit Function
					End If
				End If

			'------------------Open the File from the Application, on the Specific server-----------------------------------
			Case "Server"
				'Will be developed as per the requirement

		End Select
	Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Open dialog box does not exists")
			Fn_SISW_LifeView_FileOpenOperation = FALSE
			Set oOpenInsertFile = Nothing
			Exit Function	
	End If
	Set oNWPwd = Nothing
	Set oOpenInsertFile = Nothing

	'If System Ask for network password then provide
	Set oNWPwd = Fn_SISW_LifeView_GetObject("VIZ_NetworkPassword")
	If oNWPwd.Exist(5) Then
		'Set Username
		bReturn = Fn_SISW_LifeView_NetworkLogin(dicOpenFile("NetworkCredentials"),"")
		If bReturn = False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to invoke Menu [ "+sMenu+" ]")
			Set oOpenInsertFile = Nothing
			Set oNWPwd = Nothing
			Exit Function	
		End If
	End If
	Set oNWPwd = Nothing

	'Handle Product Structure Config Dialog
	Set oProdStrucDlg  = Fn_SISW_LifeView_GetObject("VIZ_ProductStructure")
	If oProdStrucDlg.Exist(15) Then
		bReturn =  Fn_SISW_LifeView_ProductStructureConfigure(dicOpenFile("ProdStrConfigOption"),"OK")
		If bReturn = False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Set the Option [ "+dicOpenFile("ProdStrConfigOption")+" ] in Product Structure Configuration")
			Fn_SISW_LifeView_FileOpenOperation = FALSE
			Set oProdStrucDlg = Nothing
			Exit Function	
		End If
	End If
	
	'Handle Attachments Dialog
	Set oAttachDlg  = Fn_SISW_LifeView_GetObject("VIZ_attachment")
	If oAttachDlg.Exist(15) Then
		oAttachDlg.close
	End If
	
	'Verify wether the Object is loaded
	If dicOpenFile("FileUsageConf") = "" OR dicOpenFile("FileUsageConf") = False  Then
		Set oVizWin = Fn_SISW_LifeView_GetObject("VIZ_ScratchPad")
		If  Not(oVizWin.Exist(iTimeOut)) Then
			Fn_SISW_LifeView_FileOpenOperation = FALSE
			 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Faliled to open file ")
			 Set oVizWin = Nothing
			 Exit Function
		End If
		Set oVizWin = Nothing
	End If

	Fn_SISW_LifeView_FileOpenInsertOperation = True
End Function

'=========================================================================================================================================
'****************************************    Function to Handle Dialog and Verify Messages***************************************
'
''Function Name		 	:	Fn_SISW_LifeView_DialogHandleVerifyMessage()
'
''Description		    :  	   Function to Verify Messages

''Parameters		    :	  sAction : (Menu or Toolbar)
'                                   sTitle  : Title of the Dialog
'                                   bDetailMessage   : Deatil message to be verified or not (TRUE/FALSE)
'                                   sMessage    : Error / Warning Message
'                                   sButtons   : Buttons name separated by ':'
 
''Return Value		    :  	True \ False
'
''Examples		     	:	 ' To Verify the Basic Message and not detail Message,   SaveAsWarningErrorVerify
'  									Call  Fn_SISW_LifeView_DialogHandleVerifyMessage("SaveAsWarningErrorVerify","Warning",False,"unable to save to local reference","OK")

'									To only handle the dialog of Save as Warning
'  									Call  Fn_SISW_LifeView_DialogHandleVerifyMessage("SaveAsWarningErrorVerify","Warning",False,"","OK")

'History:
'	Developer Name			Date			  Rev. No.		Reviewer		Changes Done	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Pritam S			      25-Jul-2013			 1.0			Vallari	 S.
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_LifeView_DialogHandleVerifyMessage(sAction, sTitle, bDetailMessage, sMessage, sButtons)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_LifeView_DialogHandleVerifyMessage"
	GBL_EXPECTED_MESSAGE=sMessage
   Dim objDialog
   Dim sMsg, iCount, aButtons

   Fn_SISW_LifeView_DialogHandleVerifyMessage = False
	
   Select Case sAction
		'----Case : Error Or Warning After performing Save/ SaveAs operation
		Case "SaveAsWarningErrorVerify"
				Set objDialog = Fn_SISW_LifeView_GetObject("VIZ_SaveAsWarningError")
				objDialog.SetToProperty "text",sTitle
				'If Dialog Exist
				If objDialog.Exist(10) Then
					If Not bDetailMessage Then
							If objDialog.Static("Message").Exist(5) Then
									sMsg = objDialog.Static("Message").GetRoProperty("text")
							Else
									Fn_SISW_LifeView_DialogHandleVerifyMessage = FALSE
									 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Error Message object not Present")
									 Set objDialog = Nothing
									Exit Function
							End If
					Else
						'Implement Later
					End If
				Else
					Fn_SISW_LifeView_DialogHandleVerifyMessage = FALSE
					 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Dialog [ "+sTitle+" ] not Present ")
					 Set objDialog = Nothing
					 Exit Function
				End If
			'[Tc1123(20161205c00)_PoonamC_NewDevelopment_10Mar2017 Added New Case : Warning After entering zero value in Arrow Size/Scale edit box By Priyanka Kakade ]
			Case "InvalidValueWarning"
				Set objDialog = Dialog("Warning")
				'If Dialog Exist
				If objDialog.Exist(3) Then
					If Not bDetailMessage Then
						If objDialog.Static("WarningMessage").Exist(3) Then
							sMsg = objDialog.Static("WarningMessage").GetRoProperty("text")
						Else
							Fn_SISW_LifeView_DialogHandleVerifyMessage = FALSE
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Error Message object not Present")
							Set objDialog = Nothing
							Exit Function
						End If
					Else
						'Implement Later
					End If
				Else
					Fn_SISW_LifeView_DialogHandleVerifyMessage = FALSE
					 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Dialog [ Warning ] not Present ")
					 Set objDialog = Nothing
					 Exit Function
				End If
			'---------------Case Else-----------------------------------------
			Case Else
				Fn_SISW_LifeView_DialogHandleVerifyMessage = FALSE
				 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Case Mismatch ")
				 Exit Function
   End Select

	'Click Buttons
   aButtons = Split(sButtons,":",-1,1)
   For iCount = 0 to UBound(aButtons)
		objDialog.WinButton(aButtons(iCount)).Click 5,5,micLeftBtn
   Next

	'Verify Meassage
	If sMessage <> "" Then
		If Instr(1,sMsg, sMessage) Then
			Fn_SISW_LifeView_DialogHandleVerifyMessage = TRUE
		Else
			GBL_ACTUAL_MESSAGE=sMsg
			Fn_SISW_LifeView_DialogHandleVerifyMessage = FALSE
			 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Error Message Not Verified ")
			 Set objDialog = Nothing
			 Exit Function
		End If
	End If

	'On Completion Return True
	Fn_SISW_LifeView_DialogHandleVerifyMessage = TRUE

End Function


'=========================================================================================================================================
'****************************************    Function to Handle Dialog and Verify Messages***************************************
'
''Function Name		 	:	Fn_SISW_LifeView_ProductStructureConfigure()
'
''Description		         :  	   Function to Verify Messages

''Parameters		       :	  sOption : Radiobutton
'		                                       sButton  : button name
''		                             
 
''Return Value		      :  	True \ False
'
''Examples		         	:

'								         	To only handle the dialog of Save as Warning
'  									         Call  Fn_SISW_LifeView_ProductStructureConfigure("","OK")

'History:
'	Developer Name			Date			  Rev. No.		Reviewer		Changes Done	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Pritam S			      25-Jul-2013			 1.0			Vallari	 S.
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_LifeView_ProductStructureConfigure(sOption,sButton)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_LifeView_ProductStructureConfigure"

	Dim objDialog

	Fn_SISW_LifeView_ProductStructureConfigure = FALSE

   Set objDialog = Fn_SISW_LifeView_GetObject("VIZ_ProductStructure")

   If  objDialog.Exist(10) Then
		objDialog.WinRadioButton("Options").SetTOProperty "text", sOption
		objDialog.WinRadioButton("Options").Click
		If Err.Number < 0 Then
			 Fn_SISW_LifeView_ProductStructureConfigure = FALSE
			 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Set the Radio button [ "+sOption+" ]")
			 Set objDialog = Nothing
			 Exit Function
		End If

	   If sButton <> "" Then
		    objDialog.WaitProperty "visible",True,10000
			objDialog.WinButton(sButton).Click 5,5,micLeftBtn
			If Err.Number < 0 Then
				 Fn_SISW_LifeView_ProductStructureConfigure = FALSE
				 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Click [ "+sButton+" ] button")
				 Set objDialog = Nothing
				 Exit Function
			End If
	   End If
    End If
	Set objDialog = Nothing
	Fn_SISW_LifeView_ProductStructureConfigure = TRUE

End Function


'=========================================================================================================================================
'****************************************    Function to perform the Opertion on Teamcenter Integration Preferences ***************************************
'
''Function Name		 	:	Fn_SISW_LifeView_TCIntegrationPrefOperation()
'
''Description		         :  	   Function to perform the Opertion on Teamcenter Integration Preferences

''Parameters		   		:	   sCalledFrom				    : (Menu or Toolbar)
'                                   			sAction  							: Title of the Dialog
'                                   			dicTCIntegrationPref   : Deatil message to be verified or not (TRUE/FALSE)
'                                  		  	    sButtons   						   : Buttons name separated by ':'
 
''Return Value		   	    :  	True \ False
'
''Examples		     	      :	 ' To set the 'Save configured and static representations of the current structure' in the 3D Save tab of Teancenter Integration Preferences Dialog
'
'											   Set dicTCIntegrationPref = CreateObject("Scripting.Dictionary")
'											   dicTCIntegrationPref("ProdStrucOption") = "Save configured and static representations of the current structure"
'  											   bReturn =   Fn_SISW_LifeView_TCIntegrationPrefOperation("VIZ","3D Save",dicTCIntegrationPref,"OK")

'												dicTCIntegrationPref("Show attributes form on save") = "ON"
'												dicTCIntegrationPref("Capture 2D geometry asset") = "ON"
'											   bReturn =   Fn_SISW_LifeView_TCIntegrationPrefOperation("TC","Snapshot",dicTCIntegrationPref,"Apply:OK")

'History:
'	Developer Name			Date			  Rev. No.		Reviewer		Changes Done	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Pritam S			      25-Jul-2013			 1.0			Vallari	 S.
'	Ankit Tewari			  30-Sep-2014											Added Case "Snapshot"
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_LifeView_TCIntegrationPrefOperation(sCalledFrom, sAction, dicTCIntegrationPref, sButtons)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_LifeView_TCIntegrationPrefOperation"

	Dim objDialog, aButtons, iCount,dicKeys,dicItems

	Fn_SISW_LifeView_TCIntegrationPrefOperation = False

	'Select case, to set the Object of Viz or TC
   Select Case sCalledFrom
		 	Case "TC"
				Set objDialog = Fn_SISW_LifeView_GetObject("TC_TeamcenterIntegrationPref")
				If NOT objDialog.Exist(5) Then
					sMenuXML  = Fn_LogUtil_GetXMLPath("Viz_Menu")
					sMenu = Fn_GetXMLNodeValue(sMenuXML, "FilePrefrencesTCIntegration")
					 bReturn =  Fn_MenuOperation("WinMenuSelect", sMenu)
				End If
			Case "VIZ"
				Set objDialog = Fn_SISW_LifeView_GetObject("VIZ_TeamcenterIntegrationPref")
				If NOT objDialog.Exist(5) Then
					sMenuXML  = Fn_LogUtil_GetXMLPath("Viz_Menu")
                    sMenu = Fn_GetXMLNodeValue(sMenuXML, "FilePrefrencesTCIntegration")
					bReturn = Fn_SISW_LifeView_MenuOperation("WinMenuSelect",sMenu)
				End If
   End Select
	If bReturn = False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to invoke Menu [ "+sMenu+" ]")
		Set oOpenInsertFile = Nothing
		Exit Function		
	End If


	'Select case, i.e select the Particular TAB
	If objDialog.Exist(10) Then

	   Select Case sAction
	
				'---------------------------------------------------------------------------------------------------------------------
				'Case to Perfor the 3D Save options selections
				Case "3D Save"
	
						objDialog.WinTab("TcIntegrationPrefTab").Select "3D Save"
						Wait 1

'						objDialog.WinRadioButton("ProdStrucOption").SetTOProperty "text", dicTCIntegrationPref("ProdStrucOption")
'						objDialog.WinRadioButton("ProdStrucOption").Click 2,2,micLeftBtn


						Select Case LCase(dicTCIntegrationPref("ProdStrucOption"))
							Case "save only configured structure", "saveonlyconfigured", "save configured"
								objDialog.WinRadioButton("SaveOnlyConfigured").Click 2,2,micLeftBtn
								wait 1
							Case "Save configured and static representations of the current structure", "saveconfiguredandstatic", "save configured and static"
								objDialog.WinRadioButton("SaveConfiguredAndStatic").Click 2,2,micLeftBtn
								wait 1
						End Select

						If Err.Number < 0 Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Set ON the [ "+dicTCIntegrationPref("ProdStrucOption")+" ] radio button")
							Set objDialog = Nothing
							Exit Function
						End If
				'---------------------------------------------------------------------------------------------------------------------
				Case "3D Loader"
					'Develop as per requirements
				'---------------------------------------------------------------------------------------------------------------------					
				Case "Snapshot"
					objDialog.WinTab("TcIntegrationPrefTab").Select sAction
					Wait 1
					
					dicKeys = dicTCIntegrationPref.Keys
					dicitems = dicTCIntegrationPref.Items
					For iCount = 0 to Ubound(dicKeys)
						objDialog.WinCheckBox(dicKeys(iCount)).Set "ON"
					Next
					
					If Err.Number < 0 Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Set ON the Checkboxes.")
						Set objDialog = Nothing
						Exit Function
					End If
	
				'---------------------------------------------------------------------------------------------------------------------
				Case "Session"
					'Develop as per requirements
	
				'---------------------------------------------------------------------------------------------------------------------
				Case Else
					Fn_SISW_LifeView_TCIntegrationPrefOperation = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Invalid Action Case")
					Set objDialog = Nothing
					Exit Function
	
	   End Select
	
		'Click the Buttons as passed
	   If  sButtons <> "" Then
		   aButtons = Split(sButtons, ":", -1,1)
		   For iCount = 0 to Ubound(aButtons)
				objDialog.WinButton(aButtons(iCount)).Click 5,5,micLeftBtn
				If Err.Number < 0 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Click [ "+aButtons(iCount)+" ] button")
					Set objDialog = Nothing
					Exit Function
				End If
		   Next
	   End If

	End If'///End the Outer most If block, Existence of the Preference dialog

	'Return True
	Set objDialog = Nothing
	Fn_SISW_LifeView_TCIntegrationPrefOperation = True
End Function

'=========================================================================================================================================
'****************************************    Function to login to Network for TcViz***************************************
'
''Function Name		 	:	Fn_SISW_LifeView_NetworkLogin()
'
''Description		         :   Function to perform the Opertion  on File Usage Confirmation dialog which appears after performing File:Open

''Parameters		   		:	 sUser  		 : User Name
'                                   		   sReserve   :   Reserve for Future use                                      			     
 
''Return Value		   	    :  	True \ False
'
''Examples		     	      :	    bReturn =   Fn_SISW_LifeView_NetworkLogin("shimpuka:shimpuka:Engineering:Designer","")


'History:
'	Developer Name			Date			  Rev. No.		Reviewer		Changes Done	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Pritam S			         2-Sep-2013			 1.0		       	Veena	 P.
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Reema W			         11-Oct-2014			 1.0		       	Added case to handle network password in PSE
'-------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_LifeView_NetworkLogin(sUser,sReserve)
		GBL_FAILED_FUNCTION_NAME="Fn_SISW_LifeView_NetworkLogin"
		Dim aUser, oNWPwd
		Fn_SISW_LifeView_NetworkLogin = False	
		If sReserve = "PSE" Then
			Set oNWPwd = Fn_SISW_TcViz_GetObject("TC_NetworkPassword")	
		Else
			Set oNWPwd = Fn_SISW_LifeView_GetObject("VIZ_NetworkPassword")
		End If
		If oNWPwd.Exist(8) Then
			aUser = Split(sUser,":",-1,1)
			'Set Username
			oNWPwd.WinEdit("UserName").Set aUser(0)
			If err.number < 0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Set UserName [ "+aUser(0)+" ]")
				Set oNWPwd = Nothing
				Exit Function		
			End If
			Wait 1
			'Set Password
			oNWPwd.WinEdit("Password").Set aUser(1)
			If err.number < 0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Set Password [ "+aUser(1)+" ]")
				Set oNWPwd = Nothing
				Exit Function		
			End If
			Wait 1
			'Set Group
			oNWPwd.WinEdit("Group").Set aUser(2)
			If err.number < 0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Set Group [ "+aUser(2)+" ]")
				Set oNWPwd = Nothing
				Exit Function		
			End If
			Wait 1
			'Set Role
			oNWPwd.WinEdit("Role").Set aUser(3)
			If err.number < 0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Set Role [ "+aUser(3)+" ]")
				Set oNWPwd = Nothing
				Exit Function		
			End If
			Wait 1
			'Click OK
			oNWPwd.WinButton("OK").Click 5,5,micLeftBtn
			If err.number < 0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Click [ OK ] button")
				Set oNWPwd = Nothing
				Exit Function		
			End If
			Wait 1
			Set oNWPwd = Nothing
		End If
		Fn_SISW_LifeView_NetworkLogin = True
End Function

'=========================================================================================================================================
'****************************************    Function to perform the Opertion  on File Usage Confirmation dialog***************************************
'
''Function Name		 	:	Fn_SISW_LifeView_FileUsageConfirmationOperation()
'
''Description		         :   Function to perform the Opertion  on File Usage Confirmation dialog which appears after performing File:Open

''Parameters		   		:	 sAction  							: Title of the Dialog
'                                   		   dicLocateFile   : Deatil message to be verified or not (TRUE/FALSE)                                           			     
 
''Return Value		   	    :  	True \ False
'
''Examples		     	      :	 ' To open the file from the Location C:\mainline\TestData, and if File Usage Confirmation  dialog is again to be appeared then : 
'
'											   Set dicLocateFile                                =    CreateObject("Scripting.Dictionary")
'											   dicLocateFile("Storage") = "MyComputer"
'											   dicLocateFile("FileFolderPath") =  Environment.Value("sPath") & "\TestData\TcViz"
'										 	   dicLocateFile("FileName") ="File.jt"
'											   dicLocateFile("LocateFile") = True                                         'Set this parameter to true if the File Confirmation Dialog will be appearing again, else set Flase or leave empty
'  											   bReturn =   Fn_SISW_LifeView_FileUsageConfirmationOperation("BrowseAndOpen",dicLocateFile)


'History:
'	Developer Name			Date			  Rev. No.		Reviewer		Changes Done	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Pritam S			         2-Sep-2013			 1.0		       	Veena	 P.
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function Fn_SISW_LifeView_FileUsageConfirmationOperation(sAction, dicLocateFile)																																		 
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_LifeView_FileUsageConfirmationOperation"
   On Error resume next

   Dim  bReturn
   Dim oVizWin, oFileConfWin, oLocateFileDia
   Dim iCount, iItemsCnt

   Fn_SISW_LifeView_FileUsageConfirmationOperation = False


	Set oFileConfWin = Window("VizMainWin").Dialog("FileUsageConfirmation")
	Set oLocateFileDia = Window("VizMainWin").Dialog("LocateFile")


	'Select Case for Action Open/Insert
	Select Case sAction

		Case "BrowseAndOpen"
			'if file Usage confirmation dialog exists then click browse button and search for the required file
			     If oFileConfWin.Exist(5)  Then
'						oFileConfWin.WinButton("Browse").Click 5,5,micLeftBtn
						oFileConfWin.WinButton("Browse").WaitProperty "enabled",True,100000
						oFileConfWin.WinButton("Browse").Click
						If Err.Number < 0 Then
								 Fn_SISW_LifeView_FileUsageConfirmationOperation = FALSE
								 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to click [ Browse ] button of [ File Usage Confirmation Dialog]")
								 Set oLocateFileDia = Nothing
								 Exit Function
						End If
				 End If

				 'If Locate file dialog exist then perform the File Locate operation
				 If oLocateFileDia.Exist(5) Then
						Select Case dicLocateFile("Storage")

							'------------------Open the File from the Local Hard Drive-----------------------------------
							Case "MyComputer"
								'Open the Folder
								If dicLocateFile("FileFolderPath") <> "" Then

									oLocateFileDia.WinComboBox("FilesOfType").Select "All Files (*.*)"
									If Err.Number < 0 Then
										Fn_SISW_LifeView_FileUsageConfirmationOperation = FALSE
										 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Select the [ All Files (*.*) ] from the files of type combo box")
										 Set oLocateFileDia = Nothing
										 Exit Function
									End If
									Wait 1

									'Set the File Folder name 
									oLocateFileDia.WinEdit("FileName").Set ""
									oLocateFileDia.WinEdit("FileName").Type dicLocateFile("FileFolderPath")
									If Err.Number < 0 Then
										Fn_SISW_LifeView_FileUsageConfirmationOperation = FALSE
										 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Type the File Folder Name [ "+dicLocateFile("FileFolderPath")+" ] in the FileName filed")
										 Set oLocateFileDia = Nothing
										 Exit Function
									End If
									Wait 1
				
									'Click Open button
									oLocateFileDia.WinButton("Open").Click 5,5,micLeftBtn
									If Err.Number < 0 Then
										 Fn_SISW_LifeView_FileUsageConfirmationOperation = FALSE
										 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to click [ Open ] button")
										 Set oLocateFileDia = Nothing
										 Exit Function
									End If
								End If
								Wait 2
				
								'Open the File
								If dicLocateFile("FileName") <> "" Then
									'Check wether the specific file is present, it present then select and Open
									iItemsCnt = oLocateFileDia.WinListView("FoldesFilesList").GetItemsCount()
									For iCount =0 to iItemsCnt-1
										If oLocateFileDia.WinListView("FoldesFilesList").GetItem(iCount) = dicLocateFile("FileName") Then
											'Select the File
											oLocateFileDia.WinListView("FoldesFilesList").Select iCount, micLeftBtn
											If Err.Number < 0 Then
												 Fn_SISW_LifeView_FileUsageConfirmationOperation = FALSE
												 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Select the File [ "+dicLocateFile("FileName")+" ]from the List")
												 Set oLocateFileDia = Nothing
												 Exit Function
											End If
											Wait 1
				
											'Click the Open Button
											oLocateFileDia.WinButton("Open").Click 5,5,micLeftBtn
											If Err.Number < 0 Then
												 Fn_SISW_LifeView_FileOpenOperation = FALSE
												 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to click [ Open ] button")
												 Set oLocateFileDia = Nothing
												 Exit Function
											End If
											Wait 1
											Exit For
										End If
									Next
				
									'If the File is not found, exit from the function
									If iCount = iItemsCnt Then
										Fn_SISW_LifeView_FileOpenOperation = FALSE
										oLocateFileDia.WinButton("Cancel").Click 5,5,micLeftBtn
										wait 1
										 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Specified File [ "+dicLocateFile("FileName")+" ] is not found in the Directory ["+dicLocateFile("FileFolderPath")+" ] ")
										 Set oLocateFileDia = Nothing
										 Exit Function
									End If
								End If
				
							'------------------Open the File from the Application, on the Specific server-----------------------------------
							Case "Server"
								'Will be developed as per the requirement
				
						End Select
				 End If

		Case "Details"
			'Implement As Required
				'---------------------------------------------------------------------------------------
	End Select

	'Verify wether the Object is loaded
	If dicLocateFile("LocateFile") = "" OR dicLocateFile("LocateFile") = False  Then
		Set oVizWin = Fn_SISW_LifeView_GetObject("VIZ_ScratchPad")
		If  Not(oVizWin.Exist(iTimeOut)) Then
			Fn_SISW_LifeView_FileOpenOperation = FALSE
			 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Faliled to open file ")
			 Set oVizWin = Nothing
			 Exit Function
		End If
		Set oVizWin = Nothing
	End If

	Fn_SISW_LifeView_FileUsageConfirmationOperation = True
End Function

'=======================================================================================================================================================
'****************************************    Function to create 2D Markup ***************************************
'
''Function Name		 	:	Fn_SISW_LifeView_2DMarkupCreate()
'
''Description		    :    	Function to create 2D  Markup

''Parameters		    :	1. sCalledFrom : TcViz
'										  2. sType		: Type of Markup (Text, Ellipse,etc)
'									      3. dicMarkupCreate		:  Dictionary of parameters
								
''Return Value		    :  	True \ False
'
''Examples		     	:	 set dicMarkupCreate = CreateObject("Scripting.Dictionary")
'										 dicMarkupCreate("TextType")	= "Unrestricted"
'										 dicMarkupCreate("MarkupText")	= "ABCDE"
'										  Fn_SISW_LifeView_2DMarkupCreate("VIZ", "Text", dicMarkupCreate)

'						dicMarkupCreate("ImageStorage")="MyComputer"
'						dicMarkupCreate("ImageFolderPath") = "C:\Program Files\Siemens\Teamcenter10.1\Visualization\Examples\2D\shuttle.cgm"
'						dicMarkupCreate("X1")=400
'						dicMarkupCreate("X2")=500
'						dicMarkupCreate("Y1")=500
'						dicMarkupCreate("Y2")=600
'						 bReturn = Fn_SISW_LifeView_2DMarkupCreate("TC", "InsetImage", dicMarkupCreate)

'History:
'	Developer Name			Date			     Rev. No.		Reviewer		Changes Done	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Pritam  S			          13-Sep-2013			1.0				
'-----------------------------------------------------------------------------------------------------------------------------------
'	Reema  W			          19-Aug-2014			1.0						Added code to select Menu in "Tc LCV" application
'-----------------------------------------------------------------------------------------------------------------------------------
'	Ankit T			          18-Nov-2014			1.0						Added code to insert 'inset image' in "LCV" Prespective 
'-----------------------------------------------------------------------------------------------------------------------------------
'   Shweta Rathod                 02-Dec-2014            1.0                    Added case  "Ellipse","Line","Rectangle","FreeHandLine"
'-----------------------------------------------------------------------------------------------------------------------------------
'   Shweta Rathod                 11-Dec-2014            1.0                    Modified case "Text" - 'added code to position a mouse curser
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_LifeView_2DMarkupCreate(sCalledFrom, sType, dicMarkupCreate)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_LifeView_2DMarkupCreate"
	Dim bReturn, iItemsCnt, iCount
	Dim objWin, objDMUtil, objDialog
	Dim sMenuXMLPath
	Dim sToolsMarkupMenu, sMarkupMenu
	
	Fn_SISW_LifeView_2DMarkupCreate = False
	
		'Find File Path for Lifecycle Viewer Menu XML
     sMenuXMLPath=Fn_LogUtil_GetXMLPath("Viz_Menu")
     
     'Extract Menu Paths from XML
     sToolsMarkupMenu = Fn_GetXMLNodeValue(sMenuXMLPath, "EnableMarkup")
     
	'Select the Application , Viz or TC LCV
	Select Case sCalledFrom
		Case "TC"
			Set objWin = Fn_SISW_LifeView_GetObject("LifeViewWin")
			Set objDMUtil = Fn_SISW_LifeView_GetObject("RHSImageCanvas")
		
			'Check if the Tools:Markup menu is checked or Not
			bReturn = Fn_MenuOperation("WinMenuCheck", sToolsMarkupMenu )
			If bReturn = False Then
				'Select Tools:Markup menu 
				bReturn =  Fn_MenuOperation("WinMenuSelect", sToolsMarkupMenu )
				If bReturn = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to Select Menu [ "+sToolsMarkupMenu+" ]")
					Set objWin = Nothing
					Set objDMUtil = Nothing
					Exit Function
				End If
				wait(1)
			End If
			
		Case "VIZ"
			Set objWin = Fn_SISW_LifeView_GetObject("VizMainWin")
			Set objDMUtil = Fn_SISW_LifeView_GetObject("VIZ_ScratchPad").WinObject("DMUtils")
			
			'Check if the Tools:Markup menu is checked or Not
			bReturn = Fn_SISW_LifeView_MenuOperation("CheckItemProperty",sToolsMarkupMenu+"~checked~True")
			If bReturn = False Then
				'Select Tools:Markup menu 
				bReturn =  Fn_SISW_LifeView_MenuOperation("WinMenuSelect",sToolsMarkupMenu)
				If bReturn = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to Select Menu [ "+sToolsMarkupMenu+" ]")
					Set objWin = Nothing
					Set objDMUtil = Nothing
					Exit Function
				End If
				wait(1)
			End If
		End Select
		
	'Select Type of Markup
	Select Case sType
		'------------------------------------------Insert The Text Markup----------------------------------------
		Case "Text"

			Set objDialog = objWin.Dialog("TextEditor")

			'Set the Text according to Restricted or Unrestricted
			If dicMarkupCreate("TextType") = "Restricted" Then
					'Implement as required
			Else
				'Select the Menu for the Unrestricted Text
				sMarkupMenu = Fn_GetXMLNodeValue(sMenuXMLPath, "2DMarkupTextUnRestricted")
				If sCalledFrom = "TC"  Then
					bReturn =  Fn_MenuOperation("WinMenuSelect", sMarkupMenu )	
				ElseIf sCalledFrom = "VIZ" Then
					bReturn =  Fn_SISW_LifeView_MenuOperation("WinMenuSelect",sMarkupMenu)
				End If
				If bReturn = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to Select Menu [ "+sMarkupMenu+" ]")
					Set objWin = Nothing
					Set objDMUtil = Nothing
					Set objDialog = Nothing
					Exit Function
				End If
				objDMUtil.WaitProperty "enabled",True,10
				if dicMarkupCreate("XCord") <> "" or dicMarkupCreate("YCord") <> "" then 'added by shweta to position a curser
					If dicMarkupCreate("XCord") = "" then
						dicMarkupCreate("XCord") = 25
					End If
					If dicMarkupCreate("YCord") = "" then
						dicMarkupCreate("YCord") = 25
					End If
					objDMUtil.Click cint(dicMarkupCreate("XCord")),cint(dicMarkupCreate("YCord")), micLeftBtn
				else
					objDMUtil.Click 25, 25, micLeftBtn
				End if
			End If
				wait 5
			'If Text Editor Dialog Exits then
			If objDialog.Exist(30) Then
					'Set the Markup Text
					objDialog.WinEditor("MarkupTextEdit").Type dicMarkupCreate("MarkupText")
					If Err.Number < 0 Then
						Set objWin = Nothing
						Set objDMUtil = Nothing
						Set objDialog = Nothing
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to Set Markup Text on [MarupText] Dialog")
						Exit Function
					End If

					'Click OK
					objDialog.WinButton("OK").Click 5,5,micLeftBtn
					If Err.Number < 0 Then
						Set objWin = Nothing
						Set objDMUtil = Nothing
						Set objDialog = Nothing
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to Click [OK] Button on [MarupText] Dialog")
						Exit Function
					End If
			Else
					Set objWin = Nothing
					Set objDMUtil = Nothing
					Set objDialog = Nothing
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to Find Text Editor Dialog")
					Exit Function
			End If
			Set objDialog = Nothing

		'------------------------------------------Insert The Inset Image----------------------------------------
		Case "InsetImage"

			Set objDialog = objWin.Dialog("InsetImage")

			'Select the Menu for inset image
			sMarkupMenu = Fn_GetXMLNodeValue(sMenuXMLPath, "2DMarkupInsetImage")			
			If sCalledFrom = "TC"  Then
				bReturn =  Fn_MenuOperation("WinMenuSelect", sMarkupMenu )	
			ElseIf sCalledFrom = "VIZ" Then
				bReturn =  Fn_SISW_LifeView_MenuOperation("WinMenuSelect",sMarkupMenu)
			End If
			If bReturn = False Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to Select Menu [ "+sMarkupMenu+" ]")
				Set objWin = Nothing
				Set objDMUtil = Nothing
				Set objDialog = Nothing
				Exit Function
			End If
			
			If objDialog.Exist(20) Then
					Select Case dicMarkupCreate("ImageStorage")
			
						'------------------Open the File from the Local Hard Drive-----------------------------------
						Case "MyComputer"
							'Open the Folder
							If dicMarkupCreate("ImageFolderPath") <> "" Then
								'Set the File Folder name 
								objDialog.WinEdit("FileName").Type dicMarkupCreate("ImageFolderPath")
								If Err.Number < 0 Then
									 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Type the File Folder Name [ "+dicOpenFile("ImageFolderPath")+" ] in the FileName filed")
									 Set objWin = Nothing
									 Set objDMUtil = Nothing
									 Set objDialog = Nothing
									 Exit Function
								End If
								Wait 3
			
								'Click Open button
								objDialog.WinButton("Open").Click 5,5,micLeftBtn
								If Err.Number < 0 Then
									 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to click [ Open ] button")
									 Set objWin = Nothing
									 Set objDMUtil = Nothing
									 Set objDialog = Nothing
									 Exit Function
								End If
							End If
							Wait 2
			
							'Open the File
							If dicMarkupCreate("ImageFileName") <> "" Then
								'Check wether the specific file is present, it present then select and Open
								iItemsCnt = objDialog.WinListView("FoldesFilesList").GetItemsCount()
								For iCount =0 to iItemsCnt-1
									If objDialog.WinListView("FoldesFilesList").GetItem(iCount) = dicMarkupCreate("ImageFileName") Then
										'Select the File
										objDialog.WinListView("FoldesFilesList").Select iCount, micLeftBtn
										If Err.Number < 0 Then
											 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Select the File [ "+dicMarkupCreate("ImageFileName")+" ]from the List")
											 Set objWin = Nothing
											 Set objDMUtil = Nothing
											 Set objDialog = Nothing
											 Exit Function
										End If
										Wait 1
			
										'Click the Open Button
										objDialog.WinButton("Open").Click 5,5,micLeftBtn
										If Err.Number < 0 Then
											 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to click [ Open ] button")
											 Set objWin = Nothing
											 Set objDMUtil = Nothing
											 Set objDialog = Nothing
											 Exit Function
										End If
										Wait 1
										Exit For
									End If
								Next
			
								'If the File is not found, exit from the function
								If iCount = iItemsCnt Then
									objDialog.WinButton("Cancel").Click 5,5,micLeftBtn
									 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Specified File [ "+dicOpenFile("ImageFileName")+" ] is not found in the Directory ["+dicOpenFile("ImageFolderPath")+" ] ")
									 Set objWin = Nothing
									 Set objDMUtil = Nothing
									 Set objDialog = Nothing
									 Exit Function
								End If
							End If
			
						'------------------Open the File from the Application, on the Specific server-----------------------------------
						Case "Server"
							'Will be developed as per the requirement
			
					End Select
			Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Inset Image dialog box does not exists")
						Fn_SISW_LifeView_2DMarkupCreate = FALSE
						Set objWin = Nothing
						 Set objDMUtil = Nothing
						 Set objDialog = Nothing
						Exit Function	
			End If
			If sCalledFrom = "TC"  Then
				objDMUtil.MouseDrag dicMarkupCreate("X1"),dicMarkupCreate("X2"),dicMarkupCreate("Y1"),dicMarkupCreate("Y2"),"LEFT"
			ElseIf sCalledFrom = "VIZ" Then
				objDMUtil.WaitProperty "enabled",True,10
				objDMUtil.Drag 25,25,micLeftBtn
				objDMUtil.Drop 100,100,micLeftBtn
			End If
			If Err.Number < 0 Then
				 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Drag and Drop the Image")
				 Set objWin = Nothing
				 Set objDMUtil = Nothing
				 Set objDialog = Nothing
				 Exit Function
			End If

		Case "Ellipse","Line","Rectangle","FreeHandLine"
			Select Case sCalledFrom
				Case "TC" 						
					sMarkupMenu = Fn_GetXMLNodeValue(sMenuXMLPath, "2DMarkup"+sType)
					If sType = "FreeHandLine" Then
						sMarkupMenu = Fn_GetXMLNodeValue(sMenuXMLPath, "2DMarkupFreehandLine")
					End If
					bReturn = Fn_MenuOperation("WinMenuCheck",sMarkupMenu)
					wait 2
					If bReturn = False Then
						bReturn = Fn_MenuOperation("WinMenuSelect", sMarkupMenu )
						If bReturn = False Then
							Set objWin = Nothing
							Set objDMUtil = Nothing
							Exit Function
						End If
					End If
				
				Case "VIZ"	
					
				Case "TC_PSE"
						
				End Select
		
				Set objDeviceReplay = CreateObject("Mercury.DeviceReplay")
				objDeviceReplay.DragAndDrop(objDMUtil.GetROProperty("abs_x") +objDMUtil.GetROProperty("width")/2 - dicMarkupCreate("abs_X") ),(objDMUtil.GetROProperty("abs_y") + objDMUtil.GetROProperty("height")/2-dicMarkupCreate("abs_Y") ), (objDMUtil.GetROProperty("abs_x") +objDMUtil.GetROProperty("width")/2+dicMarkupCreate("abs_X")),(objDMUtil.GetROProperty("abs_y") + objDMUtil.GetROProperty("height")/2+dicMarkupCreate("abs_Y")), 0
				If Err.Number < 0  Then
					Fn_SISW_LifeView_3DMarkupCreate = False
					Set objDeviceReplay = nothing
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to draw ["+sType+"].")
				Else
					Fn_SISW_LifeView_2DMarkupCreate = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully drawn ["+sType+"].")
				End If

		Case "Circle"

			'Implement as required
		
		Case Else
			Set objWin = Nothing
			Set objDMUtil = Nothing
			Set objDialog = Nothing
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to Find Required Case")
			Exit Function			
	End Select
	
	'Disable Markup Menu
	If sCalledFrom = "TC"  Then
		bReturn =  Fn_MenuOperation("WinMenuSelect", sToolsMarkupMenu )	
	ElseIf sCalledFrom = "VIZ" Then
		bReturn =  Fn_SISW_LifeView_MenuOperation("WinMenuSelect",sToolsMarkupMenu)
	End If
	
	Set objWin = Nothing
	Set objDMUtil = Nothing
	Set objDialog = Nothing
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Successfully Created 3D Text Markup")
	Fn_SISW_LifeView_2DMarkupCreate = True
	
End Function

'=======================================================================================================================================================
'****************************************    Function to create 3D Markup ***************************************
'
''Function Name		 	:	Fn_SISW_LifeView_3DMarkupCreate()
'
''Description		    :    	Function to create 3D  Markup

''Parameters		    :	1. sCalledFrom : TcViz
'									 2. sType		: Type of Markup (Text, Ellipse,etc)
'									 3. dicMarkupCreate		:  Dictionary of parameters
								
''Return Value		    :  	True \ False
'
''Examples		     	:	 Set dicMarkupCreate = CreateObject("Scripting.Dictionary")
'									dicMarkupCreate("ImageStorage")	= "MyComputer"
'									dicMarkupCreate("ImageFolderPath")	= "C:\Temp\Test1"
'									dicMarkupCreate("ImageFileName") = DataTable("FileName") 
'									dicMarkupCreate("NetworkCredentials") = "AutoTest1:AutoTest1:Engineering:Designer::autotest1"				' added new parameter 'networkcredentials '
'									bReturn = Fn_SISW_LifeView_3DMarkupCreate("VIZ", "InsetImage", dicMarkupCreate)

'
'									dicMarkupCreate("AnchotText")	= "Anchor Text Markup"
'									dicMarkupCreate("X_Drag") = 29
'									dicMarkupCreate("Y_Drag") = 30
'									dicMarkupCreate("X_Drop") = 354
'									dicMarkupCreate("Y_Drop") = 54
'									bReturn = Fn_SISW_LifeView_3DMarkupCreate("TC_PSE", "AnchorTextMarkup", dicMarkupCreate)
'
'									dicMarkupCreate("abs_X") = 50
'									dicMarkupCreate("abs_Y) = 50
'									bReturn = Fn_SISW_LifeView_3DMarkupCreate("TC_PSE", "Ellipse", dicMarkupCreate)
'									bReturn = Fn_SISW_LifeView_3DMarkupCreate("TC", "Ellipse", dicMarkupCreate)    for LCV
'History:
'	Developer Name			Date			     Rev. No.		Reviewer		Changes Done	
'-----------------------------------------------------------------------l--------------------------------------------------------------------------------------------------------------------------------------------------------
'	Pritam  S			          13-Sep-2013			1.0				
'-----------------------------------------------------------------------------------------------------------------------------------
'	Reema  W			          29-Aug-2014			1.0					added Case "TC_PSE"
'-----------------------------------------------------------------------------------------------------------------------------------
'	Reema  W			          01-Sep-2014			1.0					added Case "AnchorTextMarkup"
'-----------------------------------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------------------------------
'	Reema  W			          01-Sep-2014			1.0					added Case "AdvancedAnchorTextMarkup"
'-----------------------------------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------------------------------
'	Reema  W			          10-Oct-2014			1.0					Modified case "InsetImage" to handle networkcredentials dialog
'-----------------------------------------------------------------------------------------------------------------------------------
'	Shweta Rathod		          11-Nov-2014			1.0					Added Case "AnchorLine","AnchorFreeHandLine","AnchorRectangle","AnchorEllipse"  
'-------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_LifeView_3DMarkupCreate(sCalledFrom, sType, dicMarkupCreate)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_LifeView_3DMarkupCreate"
	Dim bReturn, iItemsCnt, iCount
	Dim objWin, objDMUtil, objDialog
	Dim sMenuXMLPath
	Dim sToolsMarkupMenu, sMarkupMenu
	Dim Xcoordinate, Ycoordinate, objDeviceReplay
	Dim oNWPwd
	
	Fn_SISW_LifeView_3DMarkupCreate = False

	'Select the Application , Viz or TC LCV
	Select Case sCalledFrom
		Case "TC"
			Set objWin = Fn_SISW_LifeView_GetObject("LifeViewWin")
			Set objDMUtil = objWin.WinObject("DMUtils")
		Case "TC_PSE"
			Set objWin = Window("TcVizStructureManager")
			Set objDMUtil = objWin.WinObject("3DImageViewer")
		Case "VIZ"
			Set objWin = Fn_SISW_LifeView_GetObject("VizMainWin")
			Set objDMUtil = Fn_SISW_LifeView_GetObject("VIZ_ScratchPad").WinObject("DMUtils")
	End Select

	Select Case sCalledFrom
			Case "TC"
				'Find File Path for Lifecycle Viewer Menu XML
				 sMenuXMLPath=Fn_LogUtil_GetXMLPath("Viz_Menu")
				 
				 'Extract Menu Paths from XML
				 sToolsMarkupMenu = Fn_GetXMLNodeValue(sMenuXMLPath, "EnableMarkup")
			
				'Check if the Tools:Markup menu is checked or Not
				bReturn =Fn_MenuOperation("WinMenuCheck", sToolsMarkupMenu )
				If bReturn = False Then
					'Select Tools:Markup menu 
					bReturn =  Fn_MenuOperation("WinMenuSelect", sToolsMarkupMenu )
					If bReturn = False Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to Select Menu [ "+sToolsMarkupMenu+" ]")
						Set objWin = Nothing
						Set objDMUtil = Nothing
						Exit Function
					End If
					wait(1)
				End If
			Case "VIZ"
					'Find File Path for Lifecycle Viewer Menu XML
					 sMenuXMLPath=Fn_LogUtil_GetXMLPath("Viz_Menu")
					 
					 'Extract Menu Paths from XML
					 sToolsMarkupMenu = Fn_GetXMLNodeValue(sMenuXMLPath, "EnableMarkup")
				
					'Check if the Tools:Markup menu is checked or Not
					bReturn = Fn_SISW_LifeView_MenuOperation("CheckItemProperty",sToolsMarkupMenu+"~checked~True")
					If bReturn = False Then
						'Select Tools:Markup menu 
						bReturn =  Fn_SISW_LifeView_MenuOperation("WinMenuSelect",sToolsMarkupMenu)
						If bReturn = False Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to Select Menu [ "+sToolsMarkupMenu+" ]")
							Set objWin = Nothing
							Set objDMUtil = Nothing
							Exit Function
						End If
						wait(1)
					End If
			Case "TC_PSE"
					If Fn_SISW_TcViz_ToolbarOperation("IsButtonChecked", "3D Markup", "3D Markup") =False Then
							call Fn_SISW_TcViz_ToolbarOperation("ButtonClick", "3D Markup", "3D Markup")
					End If

	End Select
	'Select Type of Markup
	Select Case sType
		'------------------------------------------Insert The Text Markup----------------------------------------
		Case "Text"

'			Set objDialog = objWin.Dialog("TextEditor")
'
'			'Set the Text according to Restricted or Unrestricted
'			If dicMarkupCreate("TextType") = "Restricted" Then
'					'Implement as required
'			Else
'				'Select the Menu for the Unrestricted Text
'				sMarkupMenu = Fn_GetXMLNodeValue(sMenuXMLPath, "2DMarkupTextUnRestricted")
'				bReturn =  Fn_SISW_LifeView_MenuOperation("WinMenuSelect",sMarkupMenu)
'				If bReturn = False Then
'					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to Select Menu [ "+sMarkupMenu+" ]")
'					Set objWin = Nothing
'					Set objDMUtil = Nothing
'					Set objDialog = Nothing
'					Exit Function
'				End If
'				objDMUtil.WaitProperty "enabled",True,10
'				objDMUtil.Click 25, 25, micLeftBtn
'			End If
'
'			'If Text Editor Dialog Exits then
'			If objDialog.Exist(10) Then
'					'Set the Markup Text
'					objDialog.WinEditor("MarkupTextEdit").Type dicMarkupCreate("MarkupText")
'					If Err.Number < 0 Then
'						Set objWin = Nothing
'						Set objDMUtil = Nothing
'						Set objDialog = Nothing
'						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to Set Markup Text on [MarupText] Dialog")
'						Exit Function
'					End If
'
'					'Click OK
'					objDialog.WinButton("OK").Click 5,5,micLeftBtn
'					If Err.Number < 0 Then
'						Set objWin = Nothing
'						Set objDMUtil = Nothing
'						Set objDialog = Nothing
'						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to Click [OK] Button on [MarupText] Dialog")
'						Exit Function
'					End If
'			Else
'					Set objWin = Nothing
'					Set objDMUtil = Nothing
'					Set objDialog = Nothing
'					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to Find Text Editor Dialog")
'					Exit Function
'			End If
'			Set objDialog = Nothing
	'------------------------------------------Insert Advanced AnchorText Markup----------------------------------------
		Case "AdvancedAnchorTextMarkup"
					Set objDialog = Fn_SISW_LifeView_GetObject("TC_Markup")
					Select Case sCalledFrom

							Case "TC" , "VIZ"
									set dMenu=description.create()
									dMenu("menuobjtype").value=2
									sTextMarkupMenu = Fn_GetXMLNodeValue(sMenuXMLPath, "TextMarkup")
									sTextMarkupMenu = Replace(sTextMarkupMenu, ":", ";")
									objWin.winmenu(dMenu).Select sTextMarkupMenu    
									' Added Call to check if Menu is already selected
									bReturn = objWin.winmenu(dMenu).CheckItemProperty(sTextMarkupMenu, "checked", true)
									If bReturn = False Then
										objWin.winmenu(dMenu).Select sTextMarkupMenu
									End If
									
									sTextMarkupMenu = Fn_GetXMLNodeValue(sMenuXMLPath, "EnableAnchor")
									sTextMarkupMenu = Replace(sTextMarkupMenu, ":", ";")
									
'									objWin.winmenu(dMenu).Select sTextMarkupMenu    
									' Added Call to check if Menu is already selected
									bReturn = objWin.winmenu(dMenu).CheckItemProperty(sTextMarkupMenu, "checked", true)
									If bReturn = False Then
										objWin.winmenu(dMenu).Select sTextMarkupMenu
									End If
									
							Case "TC_PSE"

					End Select
		
					Xcoordinate =  cInt((objDMUtil.GetROProperty("abs_x") +objDMUtil.GetROProperty("width")/2)) + cint(dicMarkupCreate("X_Drag"))
					Ycoordinate =  cInt((objDMUtil.GetROProperty("abs_y") + objDMUtil.GetROProperty("height")/2)) + cint(dicMarkupCreate("Y_Drag"))				
					
					'objDMUtil.Click Xcoordinate, Ycoordinate, micLeftbtn
					
				Set objDeviceReplay = CreateObject("Mercury.DeviceReplay")
				objDeviceReplay.MouseClick Xcoordinate, Ycoordinate, micLeftbtn
					
					'Check Existance of Dialog
					If Not objDialog.Exist(10) Then
						Set objWin = Nothing
						Set objDMUtil = Nothing
						Set objDialog = Nothing
						Set objDeviceReplay = Nothing
						set dMenu=nothing
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to Find [MarkupText] Dialog")
						Exit Function		
					End If
					
					Err.clear
					
					' click the advanced button
					objDialog.WinButton("Advanced").Click 5,5,micLeftBtn
					If Err.Number < 0 Then
						Set objWin = Nothing
						Set objDMUtil = Nothing
						Set objDialog = Nothing
						Set objDeviceReplay = Nothing
						set dMenu=nothing
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to click Advanced button on [MarupText] Dialog")
						Exit Function
					End If
					'expand part information node 
					objDialog.WinTreeView("EditMarkupTree").ExpandAll "Part Information"
					If Err.Number < 0 Then
						Set objWin = Nothing
						Set objDMUtil = Nothing
						Set objDialog = Nothing
						Set objDeviceReplay = Nothing
						set dMenu=nothing
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to expand 'part information' node of [EditMarkupTree] on [MarupText] Dialog")
						Exit Function
					End If
					'select part information; <part Name>
					objDialog.WinTreeView("EditMarkupTree").Select "Part Information;<Part Name>"
					If Err.Number < 0 Then
						Set objWin = Nothing
						Set objDMUtil = Nothing
						Set objDialog = Nothing
						Set objDeviceReplay = Nothing
						set dMenu=nothing
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to select 'Part Information;<Part Name>' node of [EditMarkupTree] on [MarupText] Dialog")
						Exit Function
					End If
					'Click Add Key button
					objDialog.WinButton("AddKey").Click 5,5,micLeftBtn
					If Err.Number < 0 Then
						Set objWin = Nothing
						Set objDMUtil = Nothing
						Set objDialog = Nothing
						Set objDeviceReplay = Nothing
						set dMenu=nothing
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to click [Add <Key> ->] Button on [MarupText] Dialog")
						Exit Function
					End If
					'Click OK button
					objDialog.WinButton("OK").Click 5,5,micLeftBtn
					If Err.Number < 0 Then
						Set objWin = Nothing
						Set objDMUtil = Nothing
						Set objDialog = Nothing
						Set objDeviceReplay = Nothing
						set dMenu=nothing
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to click [OK] Button on [MarupText] Dialog")
						Exit Function
					End If
	
		'------------------------------------------Insert The AnchorText Markup----------------------------------------
		Case "AnchorTextMarkup"
					Select Case sCalledFrom

							Case "TC_PSE"
									Set objDialog = Fn_SISW_LifeView_GetObject("PSE_Markup")
									If Fn_SISW_TcViz_ToolbarOperation("IsButtonChecked", "3D Markup", "Text") =False Then
										call Fn_SISW_TcViz_ToolbarOperation("ButtonClick", "3D Markup", "Text")
									End If
									If Fn_SISW_TcViz_ToolbarOperation("IsButtonChecked", "3D Markup", "Anchor Mode") =False Then
											call Fn_SISW_TcViz_ToolbarOperation("ButtonClick", "3D Markup", "Anchor Mode")
									End If

							Case "TC"
									Set objDialog = Fn_SISW_LifeView_GetObject("TC_Markup")
									sMarkupMenu = Fn_GetXMLNodeValue(sMenuXMLPath, "TextMarkup")
									bReturn = Fn_MenuOperation("WinMenuCheck",sMarkupMenu)
									If bReturn = False Then
										bReturn = Fn_MenuOperation("WinMenuSelect", sMarkupMenu )
										If bReturn = False Then
											Set objWin = Nothing
											Set objDMUtil = Nothing
											Exit Function
										End If
									End If
									sMarkupMenu = Fn_GetXMLNodeValue(sMenuXMLPath, "EnableAnchor")
									bReturn = Fn_MenuOperation("WinMenuCheck",sMarkupMenu)
									If bReturn = False Then
										bReturn = Fn_MenuOperation("WinMenuSelect", sMarkupMenu )
										If bReturn = False Then
											Set objWin = Nothing
											Set objDMUtil = Nothing
											Exit Function
										End If
									End If

							Case "VIZ"

					End Select
		
					Xcoordinate =  cInt((objDMUtil.GetROProperty("abs_x") +objDMUtil.GetROProperty("width")/2)) + cint(dicMarkupCreate("X_Drag"))
					Ycoordinate =  cInt((objDMUtil.GetROProperty("abs_y") + objDMUtil.GetROProperty("height")/2)) + cint(dicMarkupCreate("Y_Drag"))					
					
					Set objDeviceReplay = CreateObject("Mercury.DeviceReplay")
					objDeviceReplay.MouseClick Xcoordinate, Ycoordinate, micLeftbtn

					'Check Existance of Dialog
					If Not objDialog.Exist(10) Then
						Set objWin = Nothing
						Set objDMUtil = Nothing
						Set objDialog = Nothing
						Set objDeviceReplay = Nothing
						set dMenu=nothing
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to Find [MarkupText] Dialog")
						Exit Function		
					End If
					
					Err.clear
					
					' creation of Anchot Text Markup
		
					objDialog.WinEditor("MarkupText").Type dicMarkupCreate("AnchotText")	
					If Err.Number < 0 Then
						Set objWin = Nothing
						Set objDMUtil = Nothing
						Set objDialog = Nothing
						Set objDeviceReplay = Nothing
						set dMenu=nothing
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to Set Markup Text on [MarupText] Dialog")
						Exit Function
					End If
					objDialog.WinButton("OK").Click 5,5,micLeftBtn
					If Err.Number < 0 Then
						Set objWin = Nothing
						Set objDMUtil = Nothing
						Set objDialog = Nothing
						Set objDeviceReplay = Nothing
						set dMenu=nothing
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to Click [OK] Button on [MarupText] Dialog")
						Exit Function
					End If

					objDeviceReplay.DragAndDrop Xcoordinate, Ycoordinate, Xcoordinate+cint(dicMarkupCreate("X_Drop")), Ycoordinate+cint(dicMarkupCreate("Y_Drop")), 0
					If Err.Number < 0 Then
						 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Drag and Drop the Image")
						 Set objWin = Nothing
						 Set objDMUtil = Nothing
 						Set objDeviceReplay = Nothing
						 Set objDialog = Nothing
						 Exit Function
					End If
		
		'------------------------------------------Insert The Inset Image----------------------------------------
		Case "InsetImage"
			Select Case sCalledFrom
				Case "TC", "VIZ"
							Set objDialog = objWin.Dialog("InsetImage")
							'Select the Menu for inset image
							sMarkupMenu = Fn_GetXMLNodeValue(sMenuXMLPath, "3DMarkupInsetImage")
							bReturn =  Fn_SISW_LifeView_MenuOperation("WinMenuSelect",sMarkupMenu)
							If bReturn = False Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to Select Menu [ "+sMarkupMenu+" ]")
								Set objWin = Nothing
								Set objDMUtil = Nothing
								Set objDialog = Nothing
								Exit Function
							End If
				Case "TC_PSE"
						Set objDialog = Fn_SISW_LifeView_GetObject("PSE_InsetImage")
						bReturn =  Fn_SISW_TcViz_ToolbarOperation("ButtonClick", "3D Markup", "Inset Image")
						If bReturn = False Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to Select toolbar button [ Inset Image ]")
							Set objDMUtil = Nothing
							Set objDialog = Nothing
							Exit Function
						End If
			End Select
	
	
			
			If objDialog.Exist(20) Then
					Select Case dicMarkupCreate("ImageStorage")
			
						'------------------Open the File from the Local Hard Drive-----------------------------------
						Case "MyComputer"
							'Open the Folder
							If dicMarkupCreate("ImageFolderPath") <> "" Then
								'Set the File Folder name 
								objDialog.WinEdit("FileName").Type dicMarkupCreate("ImageFolderPath")
								If Err.Number < 0 Then
									 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Type the File Folder Name [ "+dicOpenFile("ImageFolderPath")+" ] in the FileName filed")
									 Set objWin = Nothing
									 Set objDMUtil = Nothing
									 Set objDialog = Nothing
									 Exit Function
								End If
								Wait 3
			
								'Click Open button
								objDialog.WinButton("Open").Click 5,5,micLeftBtn
								If Err.Number < 0 Then
									 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to click [ Open ] button")
									 Set objWin = Nothing
									 Set objDMUtil = Nothing
									 Set objDialog = Nothing
									 Exit Function
								End If
									'If System Ask for network password then provide
								Set oNWPwd = Fn_SISW_TcViz_GetObject("TC_NetworkPassword")
								If oNWPwd.Exist(5) Then
									'Set Username
									bReturn = Fn_SISW_LifeView_NetworkLogin(dicMarkupCreate("NetworkCredentials"),"")
									If bReturn = False Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to login")
										Set objWin = Nothing
									 	Set objDMUtil = Nothing
									 	Set objDialog = Nothing
										Set oNWPwd = Nothing
										Exit Function	
									End If
								End If
								Set oNWPwd = Nothing
							
							End If
							Wait 2
			
							'Open the File
							If dicMarkupCreate("ImageFileName") <> "" Then
								'Check wether the specific file is present, it present then select and Open
								iItemsCnt = objDialog.WinListView("FoldesFilesList").GetItemsCount()
								For iCount =0 to iItemsCnt-1
									If objDialog.WinListView("FoldesFilesList").GetItem(iCount) = dicMarkupCreate("ImageFileName") Then
										'Select the File
										objDialog.WinListView("FoldesFilesList").Select iCount, micLeftBtn
										If Err.Number < 0 Then
											 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Select the File [ "+dicMarkupCreate("ImageFileName")+" ]from the List")
											 Set objWin = Nothing
											 Set objDMUtil = Nothing
											 Set objDialog = Nothing
											 Exit Function
										End If
										Wait 1
			
										'Click the Open Button
										objDialog.WinButton("Open").Click 5,5,micLeftBtn
										If Err.Number < 0 Then
											 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to click [ Open ] button")
											 Set objWin = Nothing
											 Set objDMUtil = Nothing
											 Set objDialog = Nothing
											 Exit Function
										End If
										Wait 1
										Exit For
									End If
								Next
			
								'If the File is not found, exit from the function
								If iCount = iItemsCnt Then
									objDialog.WinButton("Cancel").Click 5,5,micLeftBtn
									 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Specified File [ "+dicOpenFile("ImageFileName")+" ] is not found in the Directory ["+dicOpenFile("ImageFolderPath")+" ] ")
									 Set objWin = Nothing
									 Set objDMUtil = Nothing
									 Set objDialog = Nothing
									 Exit Function
								End If
							End If
			
						'------------------Open the File from the Application, on the Specific server-----------------------------------
						Case "Server"
							'Will be developed as per the requirement
			
					End Select
			Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Inset Image dialog box does not exists")
						Fn_SISW_LifeView_3DMarkupCreate = FALSE
						Set objWin = Nothing
						 Set objDMUtil = Nothing
						 Set objDialog = Nothing
						Exit Function	
			End If
			objDMUtil.WaitProperty "enabled",True,10
			objDMUtil.Drag 100,100,micLeftBtn
			objDMUtil.Drop 200,200,micLeftBtn
			If Err.Number < 0 Then
				 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Drag and Drop the Image")
				 Set objWin = Nothing
				 Set objDMUtil = Nothing
				 Set objDialog = Nothing
				 Exit Function
			End If

		Case "Ellipse","Line","Rectangle","FreeHand Line"
			Select Case sCalledFrom
				Case "TC" 
						
						sMarkupMenu = Fn_GetXMLNodeValue(sMenuXMLPath, "3DMarkup"+sType)
						If sType = "FreeHand Line" Then
							sMarkupMenu = Fn_GetXMLNodeValue(sMenuXMLPath, "3DMarkupFreehandLine")
						End If
						bReturn = Fn_MenuOperation("WinMenuCheck",sMarkupMenu)
						wait 2
						If bReturn = False Then
							bReturn = Fn_MenuOperation("WinMenuSelect", sMarkupMenu )
							If bReturn = False Then
								Set objWin = Nothing
								Set objDMUtil = Nothing
								Exit Function
							End If
						End If
					
				Case "VIZ"	
						sMarkupMenu = Fn_GetXMLNodeValue(sMenuXMLPath, "3DMarkup"+sType)
						If sType = "FreeHand Line" Then
							sMarkupMenu = Fn_GetXMLNodeValue(sMenuXMLPath, "3DMarkupFreehandLine")
						End If
						bReturn = Fn_SISW_LifeView_MenuOperation("CheckItemProperty",sMarkupMenu+"~checked~True")
						wait 2
						If bReturn = False Then
							bReturn =  Fn_SISW_LifeView_MenuOperation("WinMenuSelect",sMarkupMenu)
							If bReturn = False Then
								Set objWin = Nothing
								Set objDMUtil = Nothing
								Exit Function
							End If
						End If
				Case "TC_PSE"
							If Fn_SISW_TcViz_ToolbarOperation("IsButtonChecked", "3D Markup", sType) =False Then
								call Fn_SISW_TcViz_ToolbarOperation("ButtonClick", "3D Markup", sType)
							End if
						
			End Select
		
			Set objDeviceReplay = CreateObject("Mercury.DeviceReplay")
			objDeviceReplay.DragAndDrop(objDMUtil.GetROProperty("abs_x") +objDMUtil.GetROProperty("width")/2 - dicMarkupCreate("abs_X") ),(objDMUtil.GetROProperty("abs_y") + objDMUtil.GetROProperty("height")/2-dicMarkupCreate("abs_Y") ), (objDMUtil.GetROProperty("abs_x") +objDMUtil.GetROProperty("width")/2+dicMarkupCreate("abs_X")),(objDMUtil.GetROProperty("abs_y") + objDMUtil.GetROProperty("height")/2+dicMarkupCreate("abs_Y")), 0
			If Err.Number < 0  Then
					Fn_SISW_LifeView_3DMarkupCreate = False
					Set objDeviceReplay = nothing
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to draw ["+sType+"].")
			Else
					Fn_SISW_LifeView_3DMarkupCreate = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully drawn ["+sType+"].")
			End If
			
			
			Select Case sCalledFrom
				Case "TC" 
					
					bReturn = Fn_MenuOperation("WinMenuSelect", sMarkupMenu )
					If bReturn = False Then
						Set objDeviceReplay = nothing
						Set objWin = Nothing
						Set objDMUtil = Nothing
						Exit Function
					End If
						
					
				Case "VIZ"

'						bReturn =  Fn_SISW_LifeView_MenuOperation("WinMenuSelect",sToolsMarkupMenu)
'						If bReturn = False Then
'							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to Select Menu [ "+sToolsMarkupMenu+" ]")
'							Set objWin = Nothing
'							Set objDMUtil = Nothing
'							Exit Function
'						End If
				Case "TC_PSE"		
					call Fn_SISW_TcViz_ToolbarOperation("ButtonClick", "3D Markup", sType)	
			End Select
			
		Case "Circle"

			'Implement as required
		Case "AnchorLine","AnchorFreeHandLine","AnchorRectangle","AnchorEllipse" 
			Select Case sCalledFrom
				Case "TC" 	
					sMarkupMenu = Fn_GetXMLNodeValue(sMenuXMLPath, "EnableAnchor")
					bReturn = Fn_MenuOperation("WinMenuCheck",sMarkupMenu)
					If bReturn = False Then
						bReturn = Fn_MenuOperation("WinMenuSelect", sMarkupMenu )
						If bReturn = False Then
							Set objWin = Nothing
							Set objDMUtil = Nothing
							Exit Function
						End If
					End If	
					if sType = "AnchorLine" then
						sType= "Line"
					ElseIf sType = "AnchorFreeHandLine" then
						sType = "FreeHand Line"
					ElseIf sType = "AnchorRectangle" then
						sType = "Rectangle"
					ElseIf sType = "AnchorEllipse" then
						sType = "Ellipse"
					End if	
					
					sMarkupMenu = Fn_GetXMLNodeValue(sMenuXMLPath, "3DMarkup"+sType)
					If sType = "FreeHand Line" Then
						sMarkupMenu = Fn_GetXMLNodeValue(sMenuXMLPath, "3DMarkupFreehandLine")
					End If
					bReturn = Fn_MenuOperation("WinMenuCheck",sMarkupMenu)
					wait 2
					If bReturn = False Then
						bReturn = Fn_MenuOperation("WinMenuSelect", sMarkupMenu )
						If bReturn = False Then
							Set objWin = Nothing
							Set objDMUtil = Nothing
							Exit Function
						End If
					End If
				Case "VIZ"	
				Case "TC_PSE"
					bReturn = Fn_SISW_TcViz_ToolbarOperation("IsButtonChecked", "3D Markup", "Anchor Mode")
					If bReturn = False Then
						call Fn_SISW_TcViz_ToolbarOperation("ButtonClick", "3D Markup", "Anchor Mode")
					End If
					If sType = "AnchorLine" then
						sType= "Line"
					End If
					If Fn_SISW_TcViz_ToolbarOperation("IsButtonChecked", "3D Markup", sType) =False Then
						call Fn_SISW_TcViz_ToolbarOperation("ButtonClick", "3D Markup", sType)
					End if
			End Select
			
			Xcoordinate =  cInt((objDMUtil.GetROProperty("abs_x") +objDMUtil.GetROProperty("width")/2)) + cint(dicMarkupCreate("X_Drag"))
			Ycoordinate =  cInt((objDMUtil.GetROProperty("abs_y") + objDMUtil.GetROProperty("height")/2)) + cint(dicMarkupCreate("Y_Drag"))					
			
			Set objDeviceReplay = CreateObject("Mercury.DeviceReplay")
			objDeviceReplay.MouseClick Xcoordinate, Ycoordinate, micLeftbtn		
			
			Set objDeviceReplay = CreateObject("Mercury.DeviceReplay")
			objDeviceReplay.DragAndDrop(objDMUtil.GetROProperty("abs_x") +objDMUtil.GetROProperty("width")/2 - dicMarkupCreate("abs_X") ),(objDMUtil.GetROProperty("abs_y") + objDMUtil.GetROProperty("height")/2-dicMarkupCreate("abs_Y") ), (objDMUtil.GetROProperty("abs_x") +objDMUtil.GetROProperty("width")/2+dicMarkupCreate("abs_X")),(objDMUtil.GetROProperty("abs_y") + objDMUtil.GetROProperty("height")/2+dicMarkupCreate("abs_Y")), 0
			If Err.Number < 0  Then
				Fn_SISW_LifeView_3DMarkupCreate = False
				Set objDeviceReplay = nothing
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to draw ["+sType+"].")
			Else
				Fn_SISW_LifeView_3DMarkupCreate = True
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully drawn ["+sType+"].")
			End If
			
			Select Case sCalledFrom
				Case "TC" 			
					bReturn = Fn_MenuOperation("WinMenuSelect", sMarkupMenu )
					If bReturn = False Then
						Set objDeviceReplay = nothing
						Set objWin = Nothing
						Set objDMUtil = Nothing
						Exit Function
					End If			
				Case "VIZ"	
				Case "TC_PSE"		
					call Fn_SISW_TcViz_ToolbarOperation("ButtonClick", "3D Markup", sType)	
			End Select	
			
		Case Else
			Set objWin = Nothing
			Set objDMUtil = Nothing
			Set objDialog = Nothing
			Set objDeviceReplay = nothing
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to Find Required Case")
			Exit Function			
	End Select
	
	'Disable Markup Menu
	Select Case sCalledFrom
		Case "VIZ"
			sMarkupMenu = Fn_GetXMLNodeValue(sMenuXMLPath, "EnableAnchor")
			bReturn = Fn_SISW_LifeView_MenuOperation("CheckItemProperty",sMarkupMenu+"~checked~True")
			If bReturn Then
				'Select Tools:Markup menu 
				bReturn =  Fn_SISW_LifeView_MenuOperation("WinMenuSelect",sMarkupMenu)
			End If
			
			bReturn =  Fn_SISW_LifeView_MenuOperation("WinMenuSelect",sToolsMarkupMenu)
			If bReturn = False Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to Select Menu [ "+sToolsMarkupMenu+" ]")
				Set objWin = Nothing
				Set objDMUtil = Nothing
				Exit Function
			End If
			wait(1)
		
		Case "TC"
			
			'Disable Anchor Mode
			sMarkupMenu = Fn_GetXMLNodeValue(sMenuXMLPath, "EnableAnchor")
			bReturn = Fn_MenuOperation("WinMenuCheck",sMarkupMenu)
			If bReturn Then
				Call Fn_MenuOperation("WinMenuSelect", sMarkupMenu )
			End If
			wait(1)
			
			'Disable 3D Markup
			bReturn = Fn_MenuOperation("WinMenuSelect", sToolsMarkupMenu )
			If bReturn = False Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to Select Menu [ "+sToolsMarkupMenu+" ]")
				Set objWin = Nothing
				Set objDMUtil = Nothing
				Exit Function
			End If
			
			wait(1)
					
		Case "TC_PSE"
			If Fn_SISW_TcViz_ToolbarOperation("IsButtonChecked", "3D Markup", "Anchor Mode") =True Then
				call Fn_SISW_TcViz_ToolbarOperation("ButtonClick", "3D Markup", "Anchor Mode")
			End If
			If Fn_SISW_TcViz_ToolbarOperation("IsButtonChecked", "3D Markup", "3D Markup") =True Then
				call Fn_SISW_TcViz_ToolbarOperation("ButtonClick", "3D Markup", "3D Markup")
			End If
	End Select
	Set objWin = Nothing
	Set objDMUtil = Nothing
	Set objDialog = Nothing
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Successfully Created 3D Text Markup")
	Fn_SISW_LifeView_3DMarkupCreate = True
End Function


'=======================================================================================================================================================
'****************************************    Function to Perform the Part Transformation on the Graphics of PSM and MSM and LCV ***************************************
'
''Function Name		 	:	Fn_SISW_LifeView_PartTrasformation
'
''Description		    :    	Function to Perform the Part Transformation on the Graphics of PSM and MSM and LCV

''Parameters		    :	   1. sCalledFrom : LCV, PSM, MSM
'									 2. dicTrasform		:  Dictionary of parameters
								
''Return Value		    :  	True \ False
'
''Examples		     	:	   Set dicTrasform = CreateObject("Scripting.Dictionary")
'									dicTrasform("CoordinateSystem")	= "Global"
'									dicTrasform("Transform")	= "Rotate"
'									dicTrasform("AxisDelta") = "X:12.0000~Y:20.0000~Z:10.0000"
'									dicTrasform("Buttons") = "OK" 
'									bReturn = Fn_SISW_LifeView_PartTrasformation("LCV", dicTrasform)
'									bReturn = Fn_SISW_LifeView_PartTrasformation("PSM", dicTrasform)
'									bReturn = Fn_SISW_LifeView_PartTrasformation("MSM", dicTrasform)

'History:
'	Developer Name			Date			     Rev. No.		Reviewer		Changes Done	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Pritam  S			          7-Oct-2013			1.0				
'-----------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_LifeView_PartTrasformation(sCalledFrom, dicTrasform)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_LifeView_PartTrasformation"
	Dim sMenuXMLPath, sMenu, objDialog, iCount, arrButtons, arrDelta, arrValues
	
	Fn_SISW_LifeView_PartTrasformation =False
	
	'Find File Path for Lifecycle Viewer Menu XML
	 sMenuXMLPath=Fn_LogUtil_GetXMLPath("Viz_Menu")

	 Select Case sCalledFrom
			Case "LCV"
				'Set the Object of TemporaryTrasformation Dialog
				 Set objDialog = Fn_SISW_LifeView_GetObject("LCV_PartTransformation")
				 'Extract Menu Paths from XML
				 sMenu = Fn_GetXMLNodeValue(sMenuXMLPath, "LCVActionPartTrasformation")
			Case "PSM", "MSM"
				'Set the Object of TemporaryTrasformation Dialog
				 Set objDialog = Fn_SISW_LifeView_GetObject("PSM_TemporaryTransformation")
				 'Extract Menu Paths from XML
				 sMenu = Fn_GetXMLNodeValue(sMenuXMLPath, "PSMGraphicsTemporaryTrasformation")
	 End Select

	'Select Tools:Markup menu 
	If NOT(objDialog.Exist(10)) Then ''verify dialog non existence
		bReturn =  Fn_MenuOperation("WinMenuSelect",sMenu)
		If bReturn = False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to Select Menu [ "+sMenu+" ]")
			Set objDialog = Nothing
			Exit Function
		End If
	End if

	If objDialog.Exist(10) Then
		'Select Coordinate System
		If dicTrasform("CoordinateSystem") <> "" Then
			objDialog.WinComboBox("CoordinateSystem").Select dicTrasform("CoordinateSystem")
		End If

		'Select the Type of Transform
		If dicTrasform("Transform") <> "" Then
		   Select Case dicTrasform("Transform")
				Case "Translate"
						objDialog.WinRadioButton("Transform").SetTOProperty "text", "Translate"
				Case "Rotate"
						objDialog.WinRadioButton("Transform").SetTOProperty "text", "Rotate"
				Case "Scale"
						objDialog.WinRadioButton("Transform").SetTOProperty "text", "Scale"
		   End Select
			objDialog.WinRadioButton("Transform").Set
			If Err.Number < 0 Then
				 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Set on the Radio button for the [ "+dicTrasform("Transform")+" ]")
				 Set objDialog = Nothing
				 Exit Function
			End If
		End If

		'Set the Value of Delta transform
		If dicTrasform("AxisDelta") <> "" Then
			arrDelta = Split(dicTrasform("AxisDelta"),"~",-1,1)
			For iCount = 0 to Ubound(arrDelta)
				arrValues = split(arrDelta(iCount),":",-1,1)

				'Select the Axis of Transform
				objDialog.WinRadioButton("Axis").SetTOProperty "text",arrValues(0)
				objDialog.WinRadioButton("Axis").Set
				If Err.Number < 0 Then
					 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Set on the Radio bUtton for the [ "+arrValues(0)+" ]")
					 Set objDialog = Nothing
					 Exit Function
				End If
				Wait 1

				'set the value of Delta
				objDialog.WinEdit("Delta").Set arrValues(1)
				If Err.Number < 0 Then
					 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Set on the Delta value [ "+arrValues(1)+" ] for [ "+arrValues(0)+" ]")
					 Set objDialog = Nothing
					 Exit Function
				End If
				Wait 1
			Next
		End If

		'Click the buttons
		If dicTrasform("Buttons") <> "" Then
		arrButtons = Split(dicTrasform("Buttons"),"~",-1,1)
			For iCount = 0 to Ubound(arrButtons)
				objDialog.WinButton(arrButtons(iCount)).Click 5,5,micLeftBtn
				If Err.Number < 0 Then
					 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to click button [ "+arrButtons(iCount)+" ]")
					 Set objDialog = Nothing
					 Exit Function
				End If
				Wait 1
			Next
		End If
	End If
	Fn_SISW_LifeView_PartTrasformation =True
End Function

'=======================================================================================================================================================
'****************************************    Function to create, Verify Product Views  ***************************************
'
''Function Name		 	:	Fn_SISW_LifeView_PSMandMSM_ProductViewOperations()
'
''Description		    :    	Function to create, Verify Product Views 

''Parameters		    :	   1. strAction                : Action to perform eg. Create , ViewExists etc 
'									 3. dicProductView		:  Dictionary of parameters
								
''Return Value		    :  	True \ False
'
''Examples		     	:	   Set dicProductView = CreateObject("Scripting.Dictionary")
'									dicProductView("ProductViewName")	= "TestView"
'									dicProductView("Description")	= "test Description"
'									dicProductView("PVGalleryButtons") = "OK"
'									bReturn = Fn_SISW_LifeView_PSMandMSM_ProductViewOperations("Create", dicProductView)

'									Set dicProductView = CreateObject("Scripting.Dictionary")
'									dicProductView("ProductViewName")	= "TestView"
'									dicProductView("PVGalleryButtons") = "OK"
'									bReturn = Fn_SISW_LifeView_PSMandMSM_ProductViewOperations("ViewExists", dicProductView)

'History:
'	Developer Name			Date			     Rev. No.		Reviewer		Changes Done	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Pritam  S			          7-Oct-2013			1.0				
'-----------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_LifeView_PSMandMSM_ProductViewOperations(strAction, dicProductView)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_LifeView_PSMandMSM_ProductViewOperations"
	Dim iCount, sViewName, aButtons
	Dim objProductViewGallery, objNewProductView

	Fn_SISW_LifeView_PSMandMSM_ProductViewOperations = False

	'Set the Objects
	Set objProductViewGallery = Fn_SISW_LifeView_GetObject("ProductViewGallery")

	If objProductViewGallery.Exist(3) = False Then
		bReturn = Fn_SISW_TcViz_ToolbarOperation("Enable", "Create Markup", "")
		If bReturn = False Then
			 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to unable the toolbar [ Create Markup ]")
			 Set objProductViewGallery = Nothing
			 Exit Function
		End If
		Call Fn_ReadyStatusSync(2)

		bReturn = Fn_SISW_TcViz_ToolbarOperation("ButtonClick", "Create Markup", "Create 3D Product Views")
		If bReturn = False Then
			 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to click Create Markup toolbar button [ Create 3D Product Views ]")
			 Set objProductViewGallery = Nothing
			 Exit Function
		End If
		Call Fn_ReadyStatusSync(2)
	End If

	'If Product View Gallery dialog does not exists then Exit Function
	If objProductViewGallery.Exist(3) = False Then
		 Set objProductViewGallery = Nothing
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Dialog [ProductViewGallery] does not Exist")
		Exit Function
	End If

	'Select Case
	Select Case strAction
		'--------------------------Case : Create Product View--------------------------------------
		Case "Create"
			'Click on Create a new product view button
			objProductViewGallery.JavaButton("CreateProductView").Click micLeftBtn
			If Err.Number < 0 Then
					 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to click button [ Create a new product view ]")
					 Set objProductViewGallery = Nothing
					 Exit Function
			End If

			'If New Product view dialog does not exists then exit function
			Set objNewProductView = Fn_SISW_TcViz_GetObject("NewProductView")
			If objNewProductView.Exist(20) = False Then
					 Set objProductViewGallery = Nothing
					 Set objNewProductView = Nothing
					 Exit Function
			End If

			'Set the Name of the Product view
			If Trim(dicProductView("ProductViewName")) <> "" Then
				objNewProductView.JavaEdit("ProductViewName").Set dicProductView("ProductViewName")
				If Err.Number < 0 Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Set Product View Name on Dialog [NewProductView]")
						Set objProductViewGallery = Nothing
						Set objNewProductView = Nothing
						Exit Function
				End If
				objProductViewGallery.JavaRadioButton("View").SetTOProperty "attached text", dicProductView("ProductViewName")
			Else
				'Else take the default name
				sViewName = objNewProductView.JavaEdit("ProductViewName").GetROProperty("value")
				objProductViewGallery.JavaRadioButton("View").SetTOProperty "attached text", sViewName
				Fn_SISW_LifeView_PSMandMSM_ProductViewOperations = sViewName
			End If

			'Set the Description
			If trim(dicProductView("Description")) <> "" Then
				objNewProductView.JavaEdit("Description").Set dicProductView("Description")
				If Err.Number < 0 Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Set Description on Dialog [NewProductView]")
						Set objProductViewGallery = Nothing
						Set objNewProductView = Nothing
						Exit Function
				End If
			End If

			'Click the OK button of the New Product view dialog
			objNewProductView.JavaButton("OK").Click micLeftBtn
			If Err.Number < 0 Then
					 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to click button [ OK ] of New Product view dialog")
					 Set objProductViewGallery = Nothing
					 Set objNewProductView = Nothing
					 Exit Function
			End If

			
			If objProductViewGallery.JavaRadioButton("View").Exist(5) Then
				objProductViewGallery.JavaRadioButton("View").Click 5,5,"LEFT"
			End If
			Call Fn_ReadyStatusSync(5)
			Set objNewProductView = Nothing

		'--------------------------Case : ViewExists : to verify the View existss or not--------------------------------------
		Case "ViewExists"

			'Set the Object the Product View
			objProductViewGallery.JavaRadioButton("View").SetTOProperty "attached text", dicProductView("ProductViewName")
			If Not objProductViewGallery.JavaRadioButton("View").Exist(5) Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Product View "+dicProductView("ProductViewName")+" does not exists")
				objProductViewGallery.JavaButton("Cancel").Click micLeftBtn
				Set objProductViewGallery = Nothing
				Exit Function
			End If

			'--------------------------Case : ViewExists : to verify the View existss or not--------------------------------------
		Case "DeleteView"

			'Set the Object the Product View
			objProductViewGallery.JavaRadioButton("View").SetTOProperty "attached text", dicProductView("ProductViewName")
			If objProductViewGallery.JavaRadioButton("View").Exist(5) Then
				objProductViewGallery.JavaButton("DeleteProductView").Click micLeftBtn
				If Err.Number < 0 Then
					 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to click button [ Delete ] of Product view  Gallery dialog")
					 Set objProductViewGallery = Nothing
					 Exit Function
				End If

				If Window("TcVizStructureManager").JavaWindow("JApplet").JavaDialog("Delete").Exist(5) Then
					Window("TcVizStructureManager").JavaWindow("JApplet").JavaDialog("Delete").JavaButton("Yes").Click micLeftBtn
					If Err.Number < 0 Then
						 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to click button [ Yes ] of Delete dialog")
						 Set objProductViewGallery = Nothing
						 Exit Function
					End If
				End If
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Product View "+dicProductView("ProductViewName")+" does not exists")
				objProductViewGallery.JavaButton("Cancel").Click micLeftBtn
				Set objProductViewGallery = Nothing
				Exit Function
			End If

		'--------------------------Case : DoubleClickView : to verify the View existss or not--------------------------------------
		Case "DoubleClickView"
			'Set the Object the Product View
			objProductViewGallery.JavaRadioButton("View").SetTOProperty "attached text", dicProductView("ProductViewName")
			If Not objProductViewGallery.JavaRadioButton("View").Exist(5) Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Product View "+dicProductView("ProductViewName")+" does not exists")
				objProductViewGallery.JavaButton("Cancel").Click micLeftBtn
				Set objProductViewGallery = Nothing
				Exit Function
			Else
				objProductViewGallery.JavaRadioButton("View").DblClick 5,5,"LEFT"
				If Err.Number < 0 Then
					 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to double click Product View "+dicProductView("ProductViewName"))
					 Set objProductViewGallery = Nothing
					 Set objNewProductView = Nothing
					 Exit Function
				End If
			End If
			
	End Select

	'Click the Buttons of the Product View Galley
	If dicProductView("PVGalleryButtons") <> "" Then
		aButtons = Split(dicProductView("PVGalleryButtons"),":",-1,1)
		For iCount = 0 to UBound(aButtons)
			objProductViewGallery.JavaButton(aButtons(iCount)).Click micLeftBtn
			If Err.Number < 0 Then
					 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to click button [ "+aButtons(iCount)+" ] of Product view  Gallery dialog")
					 Set objProductViewGallery = Nothing
					 Exit Function
			End If
		Next
	End If
	Set objProductViewGallery = Nothing

	'Return the Result
	If strAction = "Create" AND dicProductView("ProductViewName") = "" Then
		Fn_SISW_LifeView_PSMandMSM_ProductViewOperations = sViewName
	Else
		Fn_SISW_LifeView_PSMandMSM_ProductViewOperations = True
	End If
End Function 


'=======================================================================================================================================================
'****************************************    Function to find the Row Index of the Node in the BOM viewer table  ***************************************
'
''Function Name		 	:	Fn_SISW_LifeView_ViewerTableRowIndex()
'
''Description		    :    	Function to create, Verify Product Views 

''Parameters		    :	   1. objTable                : Tree Table object
'									 2. sNodeName	       :  Node Name
'									3,4. sRes1, sRes2        : Reserved for future use
								
''Return Value		    :  	True \ False
'
''Examples		     	:	   Set objTable = JavaWindow("DefTcVizWindow").JavaApplet("TcViewerJApplet").JavaTable("ViewerBOMTreeTable")
'									iRowIndex = Fn_SISW_LifeView_ViewerTableRowIndex(objTable,"33000047/A;1-fishing_reel (View):33000055/A;1-spool_assembly (View)","","") 

'History:
'	Developer Name			Date			     Rev. No.		Reviewer		Changes Done	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Pritam  S			          16-Oct-2013			1.0				
'-----------------------------------------------------------------------------------------------------------------------------------

Public Function Fn_SISW_LifeView_ViewerTableRowIndex(objTable,sNodeName,sRes1, sRes2)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_LifeView_ViewerTableRowIndex"
	Dim arrNode, arrTemp, arrNodePath, sNodePath, sPath, sNode
	Dim iCount, iRows, iRowCounter, iInstance, iNodeInstance

	arrNode = split(sNodeName,":",-1,1)
	sNodePath = ""
	For iCount = 0 to ubound(arrNode)
		arrTemp = split(arrNode(iCount),"@",-1,1)
		If iCount=0 Then
			sNodePath = arrTemp(0)
		Else
			sNodePath = sNodePath+":"+arrTemp(0)
		End If
	Next
	arrNodePath = split(sNodePath,":",-1,1)

	iRows = cInt(objTable.GetROProperty("rows"))
	iRowCounter = -1
	sPath=""
	For iCount = 0 to uBound(arrNode)
		If Instr(1,arrNode(iCount),"@") Then
			arrTemp = split(arrNode(iCount),"@",-1,1)
			iNodeInstance = arrTemp(1)
		Else
			iNodeInstance = 1
		End If

		iInstance = 0
		Do 
			iRowCounter = iRowCounter+1
			sNode = objTable.object.getValueAt(iRowCounter, 0).toString()
			If sNode = arrNodePath(iCount) Then
				iInstance = iInstance+1
				If iInstance = iNodeInstance Then
					bFound = True
					If sPath="" Then
						sPath = sNode
					Else
						sPath = sPath+":"+sNode
					End If
					Exit Do
				End If
			End If
		Loop Until iRowCounter = iRows

		If Not bFound Then
			Exit For
		End If

	Next

	If sNodePath = sPath Then
		Fn_SISW_LifeView_ViewerTableRowIndex = iRowCounter
	else
		Fn_SISW_LifeView_ViewerTableRowIndex = -1
	End If
	
End Function


'=======================================================================================================================================================
'****************************************    Function to perform operations on the Nodes of the BOM tree Table  ***************************************
'
''Function Name		 	:	Fn_SISW_LifeView_ViewerBOMTreeTableNodeOperation()
'
''Description		    :    	Function to create, Verify Product Views 

''Parameters		    :	   1. sAction                : Tree Table object
'									 2. sNodeName	     :  Node Name
'									 3. sColName           : Column Name
'									 4. sValue                 : Value to set or verify
'									 5. sPopupMenu       : Popup Menu
								
''Return Value		    :  	True \ False
'
''Examples		     	:	 bReturn = Fn_SISW_LifeView_ViewerBOMTreeTableNodeOperation("Select',"33000047/A;1-fishing_reel (View):33000055/A;1-spool_assembly (View)","","","")
'								  bReturn = Fn_SISW_LifeView_ViewerBOMTreeTableNodeOperation("Deselect',"33000047/A;1-fishing_reel (View):33000055/A;1-spool_assembly (View)","","","") 
'								  bReturn = Fn_SISW_LifeView_ViewerBOMTreeTableNodeOperation("LoadInViewer',"33000047/A;1-fishing_reel (View):33000055/A;1-spool_assembly (View)","","","") 
'								  bReturn = Fn_SISW_LifeView_ViewerBOMTreeTableNodeOperation("UnloadFromViewer',"33000047/A;1-fishing_reel (View):33000055/A;1-spool_assembly (View)","","","") 

'History:
'	Developer Name			Date			     Rev. No.		Reviewer		Changes Done	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Pritam  S			   16-Oct-2013			1.0				
'-----------------------------------------------------------------------------------------------------------------------------------
'	Reema  W			    19-Jun-2014			1.0								added case "VerifyLoadInViewerState"
'-----------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_LifeView_ViewerBOMTreeTableNodeOperation(sAction, sNodeName, sColName, sValue, sPopupMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_LifeView_ViewerBOMTreeTableNodeOperation"
	Dim iRowIndex, objTable

	Fn_SISW_LifeView_ViewerBOMTreeTableNodeOperation = False		
				
	Set objTable = Fn_SISW_LifeView_GetObject("ViewerBOMTreeTable")
	If sNodeName <> "" Then
		iRowIndex = Fn_SISW_LifeView_ViewerTableRowIndex(objTable,sNodeName,"","") 
	End If

	Select Case sAction
		'------------------------Case : Select => to select the node from the Tree Table--------------------------------------------
		Case "Select"
			If iRowIndex <> -1 Then
				objTable.Object.clearSelection  
				objTable.SelectRow iRowIndex 
				If Err.Number < 0 Then
					 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Select Node [ "+sNodeName+" ]")
					 Set objTable = Nothing
					 Exit Function
				Else
					Fn_SISW_LifeView_ViewerBOMTreeTableNodeOperation = True
				End If
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to find the node [ "+sNodeName+" ] in the table ")
				 Set objTable = Nothing
				 Exit Function
			End If

		'------------------------Case : Deselect => to Deselect the node from the Tree Table--------------------------------------------
		Case "Deselect"
			If iRowIndex <> -1 Then
				objTable.DeselectRow iRowIndex 
				If Err.Number < 0 Then
					 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to DeSelect Node [ "+sNodeName+" ]")
					 Set objTable = Nothing
					 Exit Function
				Else
					Fn_SISW_LifeView_ViewerBOMTreeTableNodeOperation = True
				End If
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to find the node [ "+sNodeName+" ] in the table ")
				 Set objTable = Nothing
				 Exit Function
			End If

		'------------------------Case : LoadInViewer => to put the Check mark against the node--------------------------------------------
		Case "LoadInViewer"
			If iRowIndex <> -1 Then
				Set oNode = objTable.object.getNodeForRow(iRowIndex)
				If oNode.getLoaded() = False OR lCase(oNode.getLoaded()) = "false"  Then
					oNode.stateIconClicked()
					If Err.Number < 0 Then
						 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to load [ "+sNodeName+" ] in Viewer")
						 Set objTable = Nothing
						 Exit Function
					Else
						Fn_SISW_LifeView_ViewerBOMTreeTableNodeOperation = True
					End If
				End If
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to find the node [ "+sNodeName+" ] in the table ")
				 Set objTable = Nothing
				 Exit Function
			End If

		'------------------------Case : VerifyLoadInViewerState => verify node is check or unchecked--------------------------------------------
		Case "VerifyLoadInViewerState"
			If iRowIndex <> -1 Then
				Set oNode = objTable.object.getNodeForRow(iRowIndex)
				If  sValue = False Then
					If oNode.getLoaded() = False OR lCase(oNode.getLoaded()) = "false"  Then
							 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "successfully  verified [ "+sNodeName+" ] node is not loaded in Viewer")
							Fn_SISW_LifeView_ViewerBOMTreeTableNodeOperation = True
					ElseIf oNode.getLoaded() = True OR lCase(oNode.getLoaded()) = "true" Then
							 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  verify  [ "+sNodeName+" ] node is not loaded in Viewer")
							 Set objTable = Nothing
							 Exit Function
					End If
				ElseIf  sValue = True Then
					If oNode.getLoaded() = True OR lCase(oNode.getLoaded()) = "true"  Then
							 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "successfully  verified [ "+sNodeName+" ] node is loaded in Viewer")
							Fn_SISW_LifeView_ViewerBOMTreeTableNodeOperation = True
					ElseIf oNode.getLoaded() = False OR lCase(oNode.getLoaded()) = "false" Then
							 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  verify  [ "+sNodeName+" ] node is loaded in Viewer")
							 Set objTable = Nothing
							 Exit Function
					End If
				End If
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to find the node [ "+sNodeName+" ] in the table ")
				 Set objTable = Nothing
				 Exit Function
			End If
		'------------------------Case : LoadInViewer => to uncheck the node--------------------------------------------
		Case "UnloadFromViewer"
			If iRowIndex <> -1 Then
				Set oNode = objTable.object.getNodeForRow(iRowIndex)
				If oNode.getLoaded() = True OR lCase(oNode.getLoaded()) = "true" Then
					oNode.stateIconClicked()
					If Err.Number < 0 Then
						 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to unload [ "+sNodeName+" ] from Viewer")
						 Set objTable = Nothing
						 Exit Function
					Else
						Fn_SISW_LifeView_ViewerBOMTreeTableNodeOperation = True
					End If
				End If
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to find the node [ "+sNodeName+" ] in the table ")
				 Set objTable = Nothing
				 Exit Function
			End If
	End Select

	Set objTable = Nothing
End Function

'****************************************    Function to perform various operation on RHS Image Canvas object( for 2d Images) ***************************************
'
''Function Name		 	:	Fn_SISW_LifeView_RHSCanvasOperations()
'
''Description		    :    	Function to perform various operation on RHS Image Canvas object( for 2d Images)for lcv and standalone

''Parameters		    :	1. sCalledFrom : Select the Application , Viz or TC LCV
'							2. sAction : Action to be performed
'							3. dicImageCanvasInfo		:  for Future use
								
'
''Examples		     	:	 
'									Set dicImageCanvasInfo = CreateObject( "Scripting.Dictionary" )
'
'									msgbox Fn_SISW_LifeView_RHSCanvasOperations("TC", "getSeekByValueAlongX", "")
'									msgbox Fn_SISW_LifeView_RHSCanvasOperations("TC", "getSeekByValueAlongY", "")
'									msgbox Fn_SISW_LifeView_RHSCanvasOperations("TC", "getPanByValueAlongY", "")
'									msgbox Fn_SISW_LifeView_RHSCanvasOperations("TC", "getPanByValueAlongX", "")
'									msgbox Fn_SISW_LifeView_RHSCanvasOperations("TC", "getZoomOutByValue", "")
'									msgbox Fn_SISW_LifeView_RHSCanvasOperations("TC", "getZoomInByValue", "")
'									msgbox Fn_SISW_LifeView_RHSCanvasOperations("TC", "getRotationbyValue", "")
'									
'									dicImageCanvasInfo("Color") = "LIGHTGRAY"
'									msgbox Fn_SISW_LifeView_RHSCanvasOperations("TC", "VerifyBackGroundColor", dicImageCanvasInfo)
'									msgbox Fn_SISW_LifeView_RHSCanvasOperations("TC", "isMonoColor", "")
'
'									dicImageCanvasInfo("Zoom_Y") = 125 			NOte:- y Coordinate for Zoom in and Zoom out.  dicImageCanvasInfo("Zoom_Y")  should be less than objImageCanvas.GetROProperty("height")/2
'									dicImageCanvasInfo("Zoom_X") = 125			NOte:- X Coordinate for Zoom in and Zoom out.  dicImageCanvasInfo("Zoom_X") should be less than objImageCanvas.GetROProperty("width")/2
'									msgbox Fn_SISW_LifeView_RHSCanvasOperations("TC", "DragDrop_ZoomIn~Toolbar", dicImageCanvasInfo)
'									msgbox Fn_SISW_LifeView_RHSCanvasOperations("TC", "DragDrop_ZoomOut~Toolbar", dicImageCanvasInfo)
'									msgbox Fn_SISW_LifeView_RHSCanvasOperations("TC", "DragDrop_ZoomIn~PopupMenu", dicImageCanvasInfo)
'									msgbox Fn_SISW_LifeView_RHSCanvasOperations("TC", "DragDrop_ZoomOut~PopupMenu", dicImageCanvasInfo)
'
'									msgbox Fn_SISW_LifeView_RHSCanvasOperations("TC", "RotateClockwise", "")
'
'									dicImageCanvasInfo("Seek") = 125			NOte:- X or Y Coordinate for Seek.  dicImageCanvasInfo("Seek") should be less than objImageCanvas.GetROProperty("width")/2 or objImageCanvas.GetROProperty("height")/2
'									msgbox Fn_SISW_LifeView_RHSCanvasOperations("TC", "SeekUp~Toolbar", dicImageCanvasInfo)
'
'									dicImageCanvasInfo("Pan") = 125			NOte:- X or Y Coordinate for Seek.  dicImageCanvasInfo("Pan") should be less than objImageCanvas.GetROProperty("width")/2 or objImageCanvas.GetROProperty("height")/2
'									msgbox Fn_SISW_LifeView_RHSCanvasOperations("TC", "PanDown~Toolbar", dicImageCanvasInfo)
'History:
'	Developer Name					Date			     Rev. No.		Reviewer		Changes Done	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Nitish S 			          08-Aug-2014				1.0				
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Ankit T 			          11-Aug-2014				1.0												added cases  "isMonoColor", VerifyBackGroundColor
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Ankit T 			          12-Aug-2014				1.0												added cases  "DragDrop_ZoomOut", DragDrop_ZoomIn
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Ankit T 			          19-Aug-2014				1.0												added cases  "VerifyBackGroundColor" for "Viz"
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Ankit T 			          20-Aug-2014				1.0												added Case "PopupMenu" under  Case "DragDrop_ZoomIn"	, "DragDrop_ZoomOut" 
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Ankit T 			          20-Aug-2014				1.0												added 	Case "SeekUp", "SeekDown", "SeekRight", "SeekLeft" and Case "SeekUp", "SeekDown", "SeekRight", "SeekLeft"
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Ankit T 			          21-Aug-2014				1.0												added 	Case "RotateClockwise"
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Ankit T 			          25-Aug-2014				1.0												added 	Case Case "DragDrop_ZoomIn"	Case "DragDrop_ZoomOut"Case "PanUp", "PanDown", "PanRight", "PanLeft" for Viz 
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Shweta Rathod		          10-Nov-2014				1.0												added 	Case "DragDrop_ZoomIn"	Case getCalibrationUOM, "setCalibrationUOM", "Raster Linear", "Persist Measurements"
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Reema W		          		25-Nov-2014				1.0													added 	Case "getFlyToState"	
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Ankit Tewari		          	02-Dec-2014				1.0													added 	Case "2DTextMarkupRMB"	
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Vivek Ahirrao		          29-jun-2015				1.0													added 	Case "SelectAndDelete"	
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function Fn_SISW_LifeView_RHSCanvasOperations(sCalledFrom, sAction, dicImageCanvasInfo)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_LifeView_RHSCanvasOperations"
	Dim sValue, objImageCanvas, sColour, objDeviceReplay, aAction, bFlag
	Dim hDCSource, hWndSource, backColor
	Dim iX, iY
	Dim aMarkupPointstoLoc, aDrag, aDrop, iDrag_X, iDrag_Y, iDrop_X, iDrop_Y
	Dim Xcoordinate, Ycoordinate,WShell
	Dim objCalibrateDistance, sToolbars2DMeasurementMenu, sMenuPath, dMenu
	
	Fn_SISW_LifeView_RHSCanvasOperations = False
	aAction = Split(sAction, "~")
	sAction =aAction(0)
		'Select the Application , Viz or TC 
	Select Case sCalledFrom
		Case "TC"
			Set objImageCanvas = Fn_SISW_LifeView_GetObject("RHSImageCanvas")
		Case "VIZ"
			Set objImageCanvas = Window("VizMainWin").Window("ScratchPad").WinObject("DMUtils")
		Case "TC_PSE"
			Set objImageCanvas = Window("TcVizStructureManager").WinObject("3DImageViewer")
	End Select
	If objImageCanvas.Exist(1)   AND sCalledFrom = "TC" Then
			Select Case sAction
				Case "getZoomOutByValue"
						sValue =  objImageCanvas.Object.getViewer().getViewportScale()
				Case "getZoomInByValue"
						sValue =  objImageCanvas.Object.getViewer().getViewportScale()
				Case "getRotationbyValue"
						sValue =  objImageCanvas.Object.getViewer().getViewPortRotation()
				Case "getSeekByValueAlongX"
						sValue =  objImageCanvas.Object.getViewer().getViewPortCenterX()
				Case "getSeekByValueAlongY"
						sValue =  objImageCanvas.Object.getViewer().getViewPortCenterY()
				Case "getPanByValueAlongX"
						sValue =  objImageCanvas.Object.getViewer().getViewPortCenterX()
				Case "getPanByValueAlongY"
						sValue =  objImageCanvas.Object.getViewer().getViewPortCenterY()
				Case "isMonoColor"
						sValue =  objImageCanvas.Object.getViewer().isMonoColor()
						If lCase(sValue) = "true" Then
							sValue = True
						ElseIf lCase(sValue) = "false" Then
							sValue = False
						End If	
				Case "getFlyToState"	
						sValue =  objImageCanvas.Object.getViewer().getFlyToState()
						If Err.Number < 0  Then
							Fn_SISW_LifeView_RHSCanvasOperations = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to getFlyToState value.")
							Exit Function
						End If
						sValue = CBool(sValue)						
						
				Case "DragDrop_ZoomOut"
					Select Case aAction(1)
						Case "Toolbar" 
							bFlag = Fn_ToolBarOperation("Click", "Zoom", "")
						Case "PopupMenu"
							objImageCanvas.Click 50,50,micRightBtn
							bFlag = Fn_SISW_Window_ContextMenu_Operation("Select", "Zoom", "")
						Case "Menu"
							bFlag = Fn_MenuOperation("WinMenuSelect", "Navigation:Zoom Area")
					End Select
					If bFlag = True Then
						Set objDeviceReplay = CreateObject("Mercury.DeviceReplay")
						objDeviceReplay.DragAndDrop (objImageCanvas.GetROProperty("abs_x") +(objImageCanvas.GetROProperty("width")/2)+dicImageCanvasInfo("Zoom_X")),(objImageCanvas.GetROProperty("abs_y") + (objImageCanvas.GetROProperty("height")/2)+dicImageCanvasInfo("Zoom_Y")), (objImageCanvas.GetROProperty("abs_x") +(objImageCanvas.GetROProperty("width")/2)-dicImageCanvasInfo("Zoom_X")),(objImageCanvas.GetROProperty("abs_y") + (objImageCanvas.GetROProperty("height")/2)-dicImageCanvasInfo("Zoom_Y")), 0
						Set objDeviceReplay = Nothing
						Set objImageCanvas = Nothing
						If Err.Number < 0 Then
							 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to perform [ " & sAction & " ] Action.")
							 Exit Function
						Else
							Fn_SISW_LifeView_RHSCanvasOperations = True
							 Exit Function
						End If
					End If
				Case "PanUp", "PanDown", "PanRight", "PanLeft"
					Select Case aAction(1)
						Case "Toolbar" 
							bFlag = Fn_ToolBarOperation("Click", "Pan", "")
						Case "PopupMenu"
							objImageCanvas.Click 50,50,micRightBtn
							bFlag = Fn_SISW_Window_ContextMenu_Operation("Select", "Pan", "")
					End Select
					If bFlag = True Then
						Set objDeviceReplay = CreateObject("Mercury.DeviceReplay")
						If  sAction = "PanUp" Then
							objDeviceReplay.DragAndDrop (objImageCanvas.GetROProperty("abs_x") + objImageCanvas.GetROProperty("width")/2),(objImageCanvas.GetROProperty("abs_y") + (objImageCanvas.GetROProperty("height")/2)), (objImageCanvas.GetROProperty("abs_x") +(objImageCanvas.GetROProperty("width")/2)),(objImageCanvas.GetROProperty("abs_y") + (objImageCanvas.GetROProperty("height")/2)-dicImageCanvasInfo("Pan")), 0
						ElseIf sAction = "PanDown" Then
							objDeviceReplay.DragAndDrop (objImageCanvas.GetROProperty("abs_x") + objImageCanvas.GetROProperty("width")/2),(objImageCanvas.GetROProperty("abs_y") + (objImageCanvas.GetROProperty("height")/2)), (objImageCanvas.GetROProperty("abs_x") +(objImageCanvas.GetROProperty("width")/2)),(objImageCanvas.GetROProperty("abs_y") + (objImageCanvas.GetROProperty("height")/2)+dicImageCanvasInfo("Pan")), 0
						ElseIf sAction = "PanRight" Then
							objDeviceReplay.DragAndDrop (objImageCanvas.GetROProperty("abs_x") +(objImageCanvas.GetROProperty("width")/2)),(objImageCanvas.GetROProperty("abs_y") + (objImageCanvas.GetROProperty("height")/2)), (objImageCanvas.GetROProperty("abs_x") +(objImageCanvas.GetROProperty("width")/2)+dicImageCanvasInfo("Pan")),(objImageCanvas.GetROProperty("abs_y") + (objImageCanvas.GetROProperty("height")/2)), 0
						ElseIf sAction = "PanLeft" Then
							objDeviceReplay.DragAndDrop (objImageCanvas.GetROProperty("abs_x") +(objImageCanvas.GetROProperty("width")/2)),(objImageCanvas.GetROProperty("abs_y") + (objImageCanvas.GetROProperty("height")/2)), (objImageCanvas.GetROProperty("abs_x") +(objImageCanvas.GetROProperty("width")/2)-dicImageCanvasInfo("Pan")),(objImageCanvas.GetROProperty("abs_y") + (objImageCanvas.GetROProperty("height")/2)), 0
						Else
							Fn_SISW_LifeView_RHSCanvasOperations = False
							Set objImageCanvas = Nothing
							Exit function
						End If
						Set objDeviceReplay = Nothing
						Set objImageCanvas = Nothing
						If Err.Number < 0 Then
							 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to perform [ " & sAction & " ] Action.")
							 Exit Function
						Else
							Fn_SISW_LifeView_RHSCanvasOperations = True
							 Exit Function
						End If
					End If

				Case "SeekUp", "SeekDown", "SeekRight", "SeekLeft"
					Select Case aAction(1)
						Case "Toolbar" 
							bFlag = Fn_ToolBarOperation("Click", "Seek", "")
						Case "PopupMenu"
							objImageCanvas.Click 50,50,micRightBtn
							bFlag = Fn_SISW_Window_ContextMenu_Operation("Select", "Seek", "")
					End Select
					If bFlag = True Then
						Set objDeviceReplay = CreateObject("Mercury.DeviceReplay")
						If  sAction = "SeekUp" Then
							objDeviceReplay.DragAndDrop (objImageCanvas.GetROProperty("abs_x") + objImageCanvas.GetROProperty("width")/2),(objImageCanvas.GetROProperty("abs_y") + (objImageCanvas.GetROProperty("height")/2)), (objImageCanvas.GetROProperty("abs_x") +(objImageCanvas.GetROProperty("width")/2)),(objImageCanvas.GetROProperty("abs_y") + (objImageCanvas.GetROProperty("height")/2)-dicImageCanvasInfo("Seek")), 0
						ElseIf sAction = "SeekDown" Then
							objDeviceReplay.DragAndDrop (objImageCanvas.GetROProperty("abs_x") + objImageCanvas.GetROProperty("width")/2),(objImageCanvas.GetROProperty("abs_y") + (objImageCanvas.GetROProperty("height")/2)), (objImageCanvas.GetROProperty("abs_x") +(objImageCanvas.GetROProperty("width")/2)),(objImageCanvas.GetROProperty("abs_y") + (objImageCanvas.GetROProperty("height")/2)+dicImageCanvasInfo("Seek")), 0
						ElseIf sAction = "SeekRight" Then
							objDeviceReplay.DragAndDrop (objImageCanvas.GetROProperty("abs_x") +(objImageCanvas.GetROProperty("width")/2)),(objImageCanvas.GetROProperty("abs_y") + (objImageCanvas.GetROProperty("height")/2)), (objImageCanvas.GetROProperty("abs_x") +(objImageCanvas.GetROProperty("width")/2)+dicImageCanvasInfo("Seek")),(objImageCanvas.GetROProperty("abs_y") + (objImageCanvas.GetROProperty("height")/2)), 0
						ElseIf sAction = "SeekLeft" Then
							objDeviceReplay.DragAndDrop (objImageCanvas.GetROProperty("abs_x") +(objImageCanvas.GetROProperty("width")/2)),(objImageCanvas.GetROProperty("abs_y") + (objImageCanvas.GetROProperty("height")/2)), (objImageCanvas.GetROProperty("abs_x") +(objImageCanvas.GetROProperty("width")/2)-dicImageCanvasInfo("Seek")),(objImageCanvas.GetROProperty("abs_y") + (objImageCanvas.GetROProperty("height")/2)), 0
						Else
							Fn_SISW_LifeView_RHSCanvasOperations = False
							Set objImageCanvas = Nothing
							Exit function
						End If
						Set objDeviceReplay = Nothing
						Set objImageCanvas = Nothing
						If Err.Number < 0 Then
							 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to perform [ " & sAction & " ] Action.")
							 Exit Function
						Else
							Fn_SISW_LifeView_RHSCanvasOperations = True
							 Exit Function
						End If
					End If

				Case "DragDrop_ZoomIn"	
					Select Case aAction(1)
						Case "Toolbar" 
							bFlag = Fn_ToolBarOperation("Click", "Zoom", "")
						Case "PopupMenu"
							objImageCanvas.Click 50,50,micRightBtn
							bFlag = Fn_SISW_Window_ContextMenu_Operation("Select", "Zoom", "")
					End Select
					If bFlag = True Then
						Set objDeviceReplay = CreateObject("Mercury.DeviceReplay")
						objDeviceReplay.DragAndDrop (objImageCanvas.GetROProperty("abs_x") +(objImageCanvas.GetROProperty("width")/2)-dicImageCanvasInfo("Zoom_X")),(objImageCanvas.GetROProperty("abs_y") + (objImageCanvas.GetROProperty("height")/2)-dicImageCanvasInfo("Zoom_Y")), (objImageCanvas.GetROProperty("abs_x") +(objImageCanvas.GetROProperty("width")/2)+dicImageCanvasInfo("Zoom_X")),(objImageCanvas.GetROProperty("abs_y") + (objImageCanvas.GetROProperty("height")/2)+dicImageCanvasInfo("Zoom_Y")), 0
						Set objDeviceReplay = Nothing
						Set objImageCanvas = Nothing
						If Err.Number < 0 Then
							 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to perform [ " & sAction & " ] Action.")
							 Exit Function
						Else
							Fn_SISW_LifeView_RHSCanvasOperations = True
							 Exit Function
						End If
					End If
				Case "MoveAnchorTextMarkup"	'' added by Swapna to move markup.
					aMarkupPointstoLoc = Split(dicImageCanvasInfo("MarkupPointstoLoc"), "~")
					aDrag = Split(dicImageCanvasInfo("Drag"), "~")
					aDrop = Split(dicImageCanvasInfo("Drop"), "~")
					iDrag_X =  cInt((objImageCanvas.GetROProperty("abs_x") + objImageCanvas.GetROProperty("width")/2)) + cint(aMarkupPointstoLoc(0))+cint(aDrag(0))
					iDrag_Y =  cInt((objImageCanvas.GetROProperty("abs_y") + objImageCanvas.GetROProperty("height")/2)) + cint(aMarkupPointstoLoc(1))+ cint(aDrag(1))
					Set objDeviceReplay = CreateObject("Mercury.DeviceReplay")
					objDeviceReplay.MouseClick iDrag_X, iDrag_Y, micLeftbtn
					wait 0,500
					iDrop_X =  cInt((objImageCanvas.GetROProperty("abs_x") + objImageCanvas.GetROProperty("width")/2)) + cint(aDrop(1))
					iDrop_Y =  cInt((objImageCanvas.GetROProperty("abs_y") + objImageCanvas.GetROProperty("height")/2)) + cint(aDrop(1))
					objDeviceReplay.DragAndDrop iDrag_X, iDrag_Y,iDrop_X, iDrop_Y, 0
					wait 0,200
					objDeviceReplay.MouseClick cInt(objImageCanvas.GetROProperty("abs_x"))+3, cInt(objImageCanvas.GetROProperty("abs_y"))+3, micLeftbtn
					Set objDeviceReplay = Nothing
					Set objImageCanvas = Nothing
						If Err.Number < 0 Then
							 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to perform [ " & sAction & " ] Action.")
							 Exit Function
						Else
							Fn_SISW_LifeView_RHSCanvasOperations = True
							 Exit Function
						End If
						
				Case "2DTextMarkupRMB"
					 Set WShell = CreateObject("WScript.Shell")
					If Fn_ToolBarOperation("IsSelected", "Select", "") = False Then
						call Fn_ToolBarOperation("Click", "Select", "")
					End If
					 objImageCanvas.Click 30, 30,"LEFT"
					wait 0,200
					objImageCanvas.Click 30, 30,"RIGHT"
					WShell.SendKeys "{DOWN}"
					bFlag = Fn_SISW_Window_ContextMenu_Operation("Select", dicImageCanvasInfo("PopupMenu"), "")
					If bFlag = False Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to perform [ " & sAction & " ] Action.")
						 Exit Function
					Else
						Set objWin = Nothing
						Set WShell = Nothing
						Set objImageCanvas = Nothing	
						Fn_SISW_LifeView_RHSCanvasOperations = True
						Exit Function
					End If
					
				Case "RotateClockwise"
					Select Case aAction(1)
						Case "Toolbar" 
							bFlag = Fn_ToolBarOperation("Click", "Rotate Clockwise", "")
						Case "Menu"
							bFlag = Fn_MenuOperation("WinMenuSelect","View:Rotate:Clockwise") 
					End Select
					If bFlag = True Then
						If Err.Number < 0 Then
							 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to perform [ " & sAction & " ] Action.")
							 Exit Function
						Else
							Fn_SISW_LifeView_RHSCanvasOperations = True
							 Exit Function
						End If
					End If	
				Case "getCalibrationUOM"
					sUOM = objImageCanvas.object.getViewer.getMeasure2D.getCalibrationUnits()
					If Err.Number < 0 Then
						 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to perform [ " & sAction & " ] Action.")
						 Exit Function
					Else
						Select Case sUOM
							Case "0"
								Fn_SISW_LifeView_RHSCanvasOperations = "Millimeters"
							Case "1"
								Fn_SISW_LifeView_RHSCanvasOperations = "Centimeters"
							Case "2"
								Fn_SISW_LifeView_RHSCanvasOperations = "Meters"
							Case "3"
								Fn_SISW_LifeView_RHSCanvasOperations = "Inches"
							Case "4"
								Fn_SISW_LifeView_RHSCanvasOperations = "Feet"
							Case "5"
								Fn_SISW_LifeView_RHSCanvasOperations = "Yards"
							Case Else
								Fn_SISW_LifeView_RHSCanvasOperations = false
						End Select
 	          			Exit Function
					End If	
				Case "setCalibrationUOM", "Raster Linear", "Persist Measurements"
					Select Case aAction(1)
						Case "Toolbar" 
							'Find File Path for Lifecycle Viewer Menu XML
							 sMenuPath=Fn_LogUtil_GetXMLPath("LifecycleViewer_Menu")
							 
							 'Extract Menu Paths from XML
							 sToolbars2DMeasurementMenu = Fn_GetXMLNodeValue(sMenuPath, "Toolbars2DMeasurement")
					 
							 sToolbars2DMeasurementMenu = Replace(sToolbars2DMeasurementMenu, ":", ";")
							
							'Invoke Marup Dialog by Menu Action
							Set objWin = Fn_SISW_LifeView_GetObject("LifeViewWin")
							set dMenu=description.create()
							dMenu("menuobjtype").value=2
							bReturn = objWin.winmenu(dMenu).CheckItemProperty(sToolbars2DMeasurementMenu, "checked", true)
							If bReturn = False Then
								objWin.winmenu(dMenu).Select sToolbars2DMeasurementMenu
								wait(1)
							End If
							If Fn_ToolBarOperation("IsSelected", "Enable Measurement", "") = False Then
								bFlag =  Fn_ToolBarOperation("Click", "Enable Measurement", "")
							End If
							If sAction = "setCalibrationUOM" Then
								If Fn_ToolBarOperation("IsSelected", "Calibrate Raster", "") = False Then
									call Fn_ToolBarOperation("Click", "Calibrate Raster", "")
								End If
							Else
								If Fn_ToolBarOperation("IsSelected", sAction, "") = False Then
									call Fn_ToolBarOperation("Click", sAction, "")
								End If
							End If
						Case "Menu"
							'' future use
					End Select
					Set objDeviceReplay = CreateObject("Mercury.DeviceReplay")
					Select Case sAction
						Case "setCalibrationUOM"
								set objCalibrateDistance = Fn_SISW_LifeView_GetObject("CalibrateDistance") 

								Xcoordinate =  cInt((objImageCanvas.GetROProperty("abs_x") + objImageCanvas.GetROProperty("width")/2)) - cint(dicImageCanvasInfo("FirstPoint_X"))
								Ycoordinate =  cInt((objImageCanvas.GetROProperty("abs_y") + objImageCanvas.GetROProperty("height")/2)) - cint(dicImageCanvasInfo("FirstPoint_Y"))	
								objDeviceReplay.MouseClick Xcoordinate, Ycoordinate, micLeftbtn
								Wait 1
								Xcoordinate =  cInt((objImageCanvas.GetROProperty("abs_x") + objImageCanvas.GetROProperty("width")/2)) + cint(dicImageCanvasInfo("SecondPoint_X"))
								Ycoordinate =  cInt((objImageCanvas.GetROProperty("abs_y") + objImageCanvas.GetROProperty("height")/2)) + cint(dicImageCanvasInfo("SecondPoint_Y"))	
								objDeviceReplay.MouseClick Xcoordinate, Ycoordinate, micLeftbtn
								
								If objCalibrateDistance.Exist(3) Then
									If dicImageCanvasInfo("CalibrateDistanceValue") <> "" then 
										objCalibrateDistance.WinEdit("Value").Set dicImageCanvasInfo("CalibrateDistanceValue") 
									End if
									If dicImageCanvasInfo("CalibrateDistancUnits") <> "" then 
										objCalibrateDistance.WinComboBox("Units").Select dicImageCanvasInfo("CalibrateDistancUnits") 
									End if
									wait(1)
									call Fn_UI_WinButton_Click("Fn_SISW_LifeView_RHSCanvasOperations",objCalibrateDistance,"OK",5,5,micLeftBtn)
									Set objCalibrateDistance = Nothing
								Else
									Set objCalibrateDistance = Nothing
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Open Calibrate Distance Dialog.")
							 		Exit Function
								End If
								
						Case "Raster Linear", "Persist Measurements"
							wait 2
							Xcoordinate =  cInt((objImageCanvas.GetROProperty("abs_x") + objImageCanvas.GetROProperty("width")/2)) - cint(dicImageCanvasInfo("FirstPoint_X"))
							Ycoordinate =  cInt((objImageCanvas.GetROProperty("abs_y") + objImageCanvas.GetROProperty("height")/2)) - cint(dicImageCanvasInfo("FirstPoint_Y"))	
							objDeviceReplay.MouseClick Xcoordinate, Ycoordinate, micLeftbtn
							Wait 2
							Xcoordinate =  cInt((objImageCanvas.GetROProperty("abs_x") + objImageCanvas.GetROProperty("width")/2)) + cint(dicImageCanvasInfo("SecondPoint_X"))
							Ycoordinate =  cInt((objImageCanvas.GetROProperty("abs_y") + objImageCanvas.GetROProperty("height")/2)) + cint(dicImageCanvasInfo("SecondPoint_Y"))	
							objDeviceReplay.MouseClick Xcoordinate, Ycoordinate, micLeftbtn
							Wait 2
							Xcoordinate =  cInt((objImageCanvas.GetROProperty("abs_x") + objImageCanvas.GetROProperty("width")/2)) + (cint(dicImageCanvasInfo("FirstPoint_X"))+cint(dicImageCanvasInfo("SecondPoint_X")))/2+(cint(dicImageCanvasInfo("FirstPoint_X"))-cint(dicImageCanvasInfo("SecondPoint_X")))/2
							Ycoordinate =  cInt((objImageCanvas.GetROProperty("abs_y") + objImageCanvas.GetROProperty("height")/2)) + (cint(dicImageCanvasInfo("FirstPoint_Y"))+cint(dicImageCanvasInfo("SecondPoint_Y")))/2+(cint(dicImageCanvasInfo("FirstPoint_Y"))-cint(dicImageCanvasInfo("SecondPoint_Y")))/2		
							objDeviceReplay.MouseClick Xcoordinate, Ycoordinate, micLeftbtn
					End Select
					Set objWin = Nothing
					Set objDeviceReplay = Nothing
					Set objImageCanvas = Nothing	
					If Err.Number < 0 Then
						 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to perform [ " & sAction & " ] Action.")
						 Exit Function
					Else
						Fn_SISW_LifeView_RHSCanvasOperations = True
						 Exit Function
					End If						
				Case "VerifyBackGroundColor"
						sColour =  objImageCanvas.Object.getViewer().getViewerBackground().toString()
						sColour =  mid(sColour ,instr(sColour ,"[")  ,instr(sColour ,"]") )
							' comparing colour codes RGB
							Select Case UCase(dicImageCanvasInfo("Color") )
								Case "LIGHTGRAY"
										If sColour = "[r=192,g=192,b=192]" Then
											sValue = True
										Else
											Fn_SISW_LifeView_RHSCanvasOperations = False
											Set objImageCanvas = Nothing
											Exit function
										End If
								Case "DARKGREEN"
										If sColour = "[r=0,g=128,b=128]" Then
											sValue = True
										Else
											Fn_SISW_LifeView_RHSCanvasOperations = False
											Set objImageCanvas = Nothing
											Exit function
										End If
                            	Case Else
										Fn_SISW_LifeView_RHSCanvasOperations = False
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Invalid Action Case")
										Set objImageCanvas = Nothing
										Exit function
							End Select		
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [Fn_SISW_LifeView_RHSCanvasOperations] Successfully verified colour [ " &  dicImageCanvasInfo("Color")  & " ] for case [" & sAction & "]")
				Case Else
						Fn_SISW_LifeView_RHSCanvasOperations = False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Invalid Action Case")
						Set objImageCanvas = Nothing
						Exit function
			End Select
			If Err.Number < 0 Then
				 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to perform [ " & sAction & " ] Action.")
				 Set objImageCanvas = Nothing
				 Exit Function
			Else
					 Set objImageCanvas = Nothing
					Fn_SISW_LifeView_RHSCanvasOperations = sValue
			End If
	ElseIf objImageCanvas.Exist(1)   AND sCalledFrom = "VIZ" Then
			Select Case sAction
					Case "isMonoColor"
							sValue = False
							Extern.Declare micLong,"GetPixel","gdi32","GetPixel",micLong,micLong,micLong
							Extern.Declare micLong,"GetWindowDC","user32","GetWindowDC",micLong
							Extern.Declare micLong,"ReleaseDC","user32","ReleaseDC",micLong,micLong
							Extern.Declare micLong,"GetDC","user32","GetDC",micLong
							Extern.Declare micLong,"SetForegroundWindow","user32","SetForegroundWindow",micLong
							hWndSource = objImageCanvas.GetROProperty("hwnd")
							extern.SetForegroundWindow hWndSource
							hDCSource = Clng(Extern.GetDC(hWndSource))
							
							iX = objImageCanvas.GetROProperty("width")/2
							iY = objImageCanvas.GetROProperty("height")/2
							
							backColor=Clng(Extern.GetPixel(hDCSource, iX,iY))
							Extern.ReleaseDC hWndSource, hDCSource 
'							sColour = Fn_SISW_getRGBColorValue(backColor)
							If backColor = 12632256 OR backColor = 0  Then
								sValue = True
							End If
						
					Case "VerifyBackGroundColor"
							sValue =  False
							Extern.Declare micLong,"GetPixel","gdi32","GetPixel",micLong,micLong,micLong
							Extern.Declare micLong,"GetWindowDC","user32","GetWindowDC",micLong
							Extern.Declare micLong,"ReleaseDC","user32","ReleaseDC",micLong,micLong
							Extern.Declare micLong,"GetDC","user32","GetDC",micLong
							Extern.Declare micLong,"SetForegroundWindow","user32","SetForegroundWindow",micLong
							hWndSource = objImageCanvas.GetROProperty("hwnd")
							extern.SetForegroundWindow hWndSource
							hDCSource = Clng(Extern.GetDC(hWndSource))
							backColor=Clng(Extern.GetPixel(hDCSource, Clng(1),Clng(1)))
							Extern.ReleaseDC hWndSource, hDCSource 
							sColour = Fn_SISW_getRGBColorValue(backColor)
							
							Select Case UCase(dicImageCanvasInfo("Color") )
								Case "LIGHTGRAY"
										If sColour = "[r=192,g=192,b=192]" Then
											sValue = True
										Else
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [Fn_SISW_LifeView_RHSCanvasOperations] FAILED To  verify colour [ " &  dicImageCanvasInfo("Color")  & " ] for case [" & sAction & "]")
											Fn_SISW_LifeView_RHSCanvasOperations = False
											Set objImageCanvas = Nothing
											Exit function
										End If
								Case "DARKGREEN"
										If sColour = "[r=0,g=128,b=128]" Then
											sValue = True
										Else
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [Fn_SISW_LifeView_RHSCanvasOperations] FAILED To  verify colour [ " &  dicImageCanvasInfo("Color")  & " ] for case [" & sAction & "]")
											Fn_SISW_LifeView_RHSCanvasOperations = False
											Set objImageCanvas = Nothing
											Exit function
										End If
								Case "WHITE"
										If sColour = "[r=255,g=255,b=255]" Then
											sValue = True
										Else
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [Fn_SISW_LifeView_RHSCanvasOperations] FAILED To  verify colour [ " &  dicImageCanvasInfo("Color")  & " ] for case [" & sAction & "]")
											Fn_SISW_LifeView_RHSCanvasOperations = False
											Set objImageCanvas = Nothing
											Exit function
										End If
								Case Else
											Fn_SISW_LifeView_RHSCanvasOperations = False
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [Fn_SISW_LifeView_RHSCanvasOperations] FAILED To  verify colour [ " &  dicImageCanvasInfo("Color")  & " ] for case [" & sAction & "]")
											Set objImageCanvas = Nothing
											Exit function
							End Select
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [Fn_SISW_LifeView_RHSCanvasOperations] Successfully verified colour [ " &  dicImageCanvasInfo("Color")  & " ] for case [" & sAction & "]")
					
					Case "DragDrop_ZoomIn"	
						Select Case aAction(1)
							Case "Toolbar" 
								bFlag = Fn_ToolBarOperation("Click", "Zoom", "")
							Case "PopupMenu"
								objImageCanvas.Click 50,50,micRightBtn
								bFlag = Fn_SISW_Window_ContextMenu_Operation("Select", "Zoom", "")
						End Select
						If bFlag = True Then
							Set objDeviceReplay = CreateObject("Mercury.DeviceReplay")
							objDeviceReplay.DragAndDrop (objImageCanvas.GetROProperty("abs_x") +(objImageCanvas.GetROProperty("width")/2)-dicImageCanvasInfo("Zoom_X")),(objImageCanvas.GetROProperty("abs_y") + (objImageCanvas.GetROProperty("height")/2)-dicImageCanvasInfo("Zoom_Y")), (objImageCanvas.GetROProperty("abs_x") +(objImageCanvas.GetROProperty("width")/2)+dicImageCanvasInfo("Zoom_X")),(objImageCanvas.GetROProperty("abs_y") + (objImageCanvas.GetROProperty("height")/2)+dicImageCanvasInfo("Zoom_Y")), 0
							Set objDeviceReplay = Nothing
							Set objImageCanvas = Nothing
							If Err.Number < 0 Then
								 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to perform [ " & sAction & " ] Action.")
								 Exit Function
							Else
								Fn_SISW_LifeView_RHSCanvasOperations = True
								 Exit Function
							End If
						End If
						
					Case "DragDrop_ZoomOut"
						Select Case aAction(1)
							Case "Toolbar" 
								bFlag = Fn_ToolBarOperation("Click", "Zoom", "")
							Case "PopupMenu"
								objImageCanvas.Click 50,50,micRightBtn
								bFlag = Fn_SISW_Window_ContextMenu_Operation("Select", "Zoom", "")
						End Select
						If bFlag = True Then
							Set objDeviceReplay = CreateObject("Mercury.DeviceReplay")
							objDeviceReplay.DragAndDrop (objImageCanvas.GetROProperty("abs_x") +(objImageCanvas.GetROProperty("width")/2)+dicImageCanvasInfo("Zoom_X")),(objImageCanvas.GetROProperty("abs_y") + (objImageCanvas.GetROProperty("height")/2)+dicImageCanvasInfo("Zoom_Y")), (objImageCanvas.GetROProperty("abs_x") +(objImageCanvas.GetROProperty("width")/2)-dicImageCanvasInfo("Zoom_X")),(objImageCanvas.GetROProperty("abs_y") + (objImageCanvas.GetROProperty("height")/2)-dicImageCanvasInfo("Zoom_Y")), 0
							Set objDeviceReplay = Nothing
							Set objImageCanvas = Nothing
							If Err.Number < 0 Then
								 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to perform [ " & sAction & " ] Action.")
								 Exit Function
							Else
								Fn_SISW_LifeView_RHSCanvasOperations = True
								 Exit Function
							End If
						End If
					
					Case "PanUp", "PanDown", "PanRight", "PanLeft"
						Select Case aAction(1)
							Case "Toolbar" 
								bFlag = Fn_ToolBarOperation("Click", "Pan", "")
							Case "PopupMenu"
								objImageCanvas.Click 50,50,micRightBtn
								bFlag = Fn_SISW_Window_ContextMenu_Operation("Select", "Pan", "")
						End Select
						If bFlag = True Then
							Set objDeviceReplay = CreateObject("Mercury.DeviceReplay")
							If  sAction = "PanUp" Then
								objDeviceReplay.DragAndDrop (objImageCanvas.GetROProperty("abs_x") + objImageCanvas.GetROProperty("width")/2),(objImageCanvas.GetROProperty("abs_y") + (objImageCanvas.GetROProperty("height")/2)), (objImageCanvas.GetROProperty("abs_x") +(objImageCanvas.GetROProperty("width")/2)),(objImageCanvas.GetROProperty("abs_y") + (objImageCanvas.GetROProperty("height")/2)-dicImageCanvasInfo("Pan")), 0
							ElseIf sAction = "PanDown" Then
								objDeviceReplay.DragAndDrop (objImageCanvas.GetROProperty("abs_x") + objImageCanvas.GetROProperty("width")/2),(objImageCanvas.GetROProperty("abs_y") + (objImageCanvas.GetROProperty("height")/2)), (objImageCanvas.GetROProperty("abs_x") +(objImageCanvas.GetROProperty("width")/2)),(objImageCanvas.GetROProperty("abs_y") + (objImageCanvas.GetROProperty("height")/2)+dicImageCanvasInfo("Pan")), 0
							ElseIf sAction = "PanRight" Then
								objDeviceReplay.DragAndDrop (objImageCanvas.GetROProperty("abs_x") +(objImageCanvas.GetROProperty("width")/2)),(objImageCanvas.GetROProperty("abs_y") + (objImageCanvas.GetROProperty("height")/2)), (objImageCanvas.GetROProperty("abs_x") +(objImageCanvas.GetROProperty("width")/2)+dicImageCanvasInfo("Pan")),(objImageCanvas.GetROProperty("abs_y") + (objImageCanvas.GetROProperty("height")/2)), 0
							ElseIf sAction = "PanLeft" Then
								objDeviceReplay.DragAndDrop (objImageCanvas.GetROProperty("abs_x") +(objImageCanvas.GetROProperty("width")/2)),(objImageCanvas.GetROProperty("abs_y") + (objImageCanvas.GetROProperty("height")/2)), (objImageCanvas.GetROProperty("abs_x") +(objImageCanvas.GetROProperty("width")/2)-dicImageCanvasInfo("Pan")),(objImageCanvas.GetROProperty("abs_y") + (objImageCanvas.GetROProperty("height")/2)), 0
							Else
								Fn_SISW_LifeView_RHSCanvasOperations = False
								Set objImageCanvas = Nothing
								Exit function
							End If
							Set objDeviceReplay = Nothing
							Set objImageCanvas = Nothing
							If Err.Number < 0 Then
								 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to perform [ " & sAction & " ] Action.")
								 Exit Function
							Else
								Fn_SISW_LifeView_RHSCanvasOperations = True
								 Exit Function
							End If
						End If
						Set objDeviceReplay = Nothing
						Set objImageCanvas = Nothing
					
					Case Else
						Fn_SISW_LifeView_RHSCanvasOperations = False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Invalid Action Case")
						Set objImageCanvas = Nothing
						Exit function
			End Select
			If Err.Number < 0 Then
				 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to perform [ " & sAction & " ] Action.")
				 Set objImageCanvas = Nothing
				 Exit Function
			Else
					 Set objImageCanvas = Nothing
					Fn_SISW_LifeView_RHSCanvasOperations = sValue
			End If
	ElseIf objImageCanvas.Exist(1)   AND sCalledFrom = "TC_PSE" Then
		Select Case sAction
			Case "3DViewerPopupMenuSelect"
				Set WShell = CreateObject("WScript.Shell")
				objImageCanvas.Click dicImageCanvasInfo("XCord"),dicImageCanvasInfo("YCord"),micLeftBtn
				Wait 0,200
				objImageCanvas.Click dicImageCanvasInfo("XCord"),dicImageCanvasInfo("YCord"),micRightBtn
				WShell.SendKeys "{DOWN}"
				bFlag = Fn_SISW_Window_ContextMenu_Operation("Select",dicImageCanvasInfo("PopupMenu"), "")
				If bFlag = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to perform [ " & sAction & " ] Action.")
					 Exit Function
				Else
					Set objWin = Nothing
					Set WShell = Nothing
					Set objImageCanvas = Nothing	
					Fn_SISW_LifeView_RHSCanvasOperations = True
					Exit Function
				End If
			Case "SelectAndDelete"
				Set WShell = CreateObject("WScript.Shell")
				Set objDeviceReplay = CreateObject("Mercury.DeviceReplay")
				objDeviceReplay.MouseClick dicImageCanvasInfo("XCord"),dicImageCanvasInfo("YCord"),0
				Wait 0,200
				objDeviceReplay.MouseClick dicImageCanvasInfo("XCord"),dicImageCanvasInfo("YCord"),2
				WShell.SendKeys "{DOWN}"
				bFlag = Fn_SISW_Window_ContextMenu_Operation("Select",dicImageCanvasInfo("PopupMenu"), "")
				If bFlag = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to perform [ " & sAction & " ] Action.")
					Exit Function
				Else
					Set objWin = Nothing
					Set WShell = Nothing
					Set objImageCanvas = Nothing	
					Fn_SISW_LifeView_RHSCanvasOperations = True
					Exit Function
				End If
			Case Else
				Fn_SISW_LifeView_RHSCanvasOperations = False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Invalid Action Case")
				Set objImageCanvas = Nothing
				Exit function
		End Select
	Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Image View does not exists")
			 Set objImageCanvas = Nothing
	End If
End function 


'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'=======================================================================================================================================================
'****************************************    Function to set and verify values in View Preferences Dialog ***************************************
'
''Function Name		 	:	Fn_SISW_LifeView_ViewPreferencesOperation()
'
''Description		    :    	Function to set, Verify View Preferences Dialog 

''Parameters		    :	   				 1. sAction                : Set Or Verify
'									 2. sCalledFrom	     :  Called from Teamcenter or Standalone Viz  (TC or VIZ)
'									 3. sInvokeFrm           : Menu Or RMB
'									 4. dicViewPreferences               :
'									 
'										Set dicViewPreferences = CreateObject( "Scripting.Dictionary" )
'										dicViewPreferences("flipmousedirectionforzoom") = True				'
'										dicViewPreferences("showborder") = True
'										dicViewPreferences("viewblackandwhiteimages") = True
'										dicViewPreferences("viewbaselayerinmonocolor") = True
'										'dicViewPreferences("No Rotation") = True
'										dicViewPreferences("90 Degrees CW") = True
'										dicViewPreferences("Zoom") = "ON"
'										dicViewPreferences("MeasuredWidth") = "157"	
'										dicViewPreferences("BackgroundColor")= "LightGrey"		
'Return Value		    :  	True \ False
'
''Examples		     	:	 bReturn = Fn_SISW_LifeView_ViewPreferencesOperation("Verify","TC","Menu",dicViewPreferences)
'								bReturn = Fn_SISW_LifeView_ViewPreferencesOperation("GetSelectedInitialValueMode","TC","Menu","")
'
'History:
'	Developer Name			Date			     Rev. No.		Reviewer		Changes Done	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Reema  W			    08-Aug-2014			1.0			
'   Ankit Tewari			11-Aug-2014											Added Case 'BackgroundColor' to set and verify Background color.
'   Ankit Tewari			13-Aug-2014											Added Case 'GetSelectedInitialValueMode' to get selected Initial Value Mode.
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function Fn_SISW_LifeView_ViewPreferencesOperation(sAction,sCalledFrom, sInvokeFrm ,dicViewPreferences)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_LifeView_ViewPreferencesOperation"
   Dim objViewPrefences,objCheckbox, aKeyFields, aValueFields,iCount
   Dim bFlag, rgbValue(2), sValue, arrInitialValueMode
    bFlag = True
    Fn_SISW_LifeView_ViewPreferencesOperation = False
   	Select Case sCalledFrom
		Case "TC"
			Set objViewPrefences = Window("VizWindow").Dialog("ViewPreferences")
		   	If Fn_UI_ObjectExist("Fn_SISW_LifeView_ViewPreferencesOperation",objViewPrefences )=False Then
				'Invoking View Prefrences Winidow
				If sInvokeFrm = "Menu" Then
					Call Fn_MenuOperation("Select", "View:Preferences...")
				Else
					Fn_SISW_LifeView_ViewPreferencesOperation = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to execute Function [ Fn_SISW_LifeView_ViewPreferencesOperation ] Invalid case [ " & sAction & " ] ")
					Exit function
				End If
		   End If
		Case "VIZ"					'Added to handle Object in Visualization Standalone
			Set objViewPrefences = Window("VizMainWin").Dialog("ViewPreferences")
		   	If Fn_UI_ObjectExist("Fn_SISW_LifeView_ViewPreferencesOperation",objViewPrefences )=False Then
				'Invoking View Prefrences Winidow
				If sInvokeFrm = "Menu" Then
					 call Fn_SISW_LifeView_MenuOperation("WinMenuSelect","View:Preferences...")
				Else
					Fn_SISW_LifeView_ViewPreferencesOperation = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to execute Function [ Fn_SISW_LifeView_ViewPreferencesOperation ] Invalid case [ " & sAction & " ] ")
					Exit function
				End If
		   End If
	End Select

	
   If Fn_UI_ObjectExist("Fn_SISW_LifeView_ViewPreferencesOperation",objViewPrefences )=True Then
	Select Case sAction
		Case "Set"
			aKeyFields = dicViewPreferences.Keys
			aValueFields = dicViewPreferences.Items
			For iCount = 0 to Ubound(aKeyFields)
			
				Select Case aKeyFields(iCount)
					Case "flipmousedirectionforzoom","showborder","viewblackandwhiteimages","viewbaselayerinmonocolor"
						Set objCheckbox = objViewPrefences.WinCheckBox(aKeyFields(iCount))
						If aValueFields(iCount) =True Then
							objCheckbox.Set "ON"
						ElseIf aValueFields(iCount) = False Then
		                      objCheckbox.Set "OFF"
						End If
						
					Case "90 Degrees CW" ,"90 Degrees CCW","No Rotation","180 Degrees"
						objViewPrefences.WinRadioButton("InitialViewRotation").SetTOProperty "text" ,aKeyFields(iCount)
						If aValueFields(iCount) =True Then
							objViewPrefences.WinRadioButton("InitialViewRotation").Set
						Else
							Fn_SISW_LifeView_ViewPreferencesOperation = False
							Set objViewPrefences = nothing
							Exit function
						End If
						
					Case "Browse","Zoom","Seek","Zoom Area","Pan"
						objViewPrefences.WinRadioButton("InitialViewMode").SetTOProperty "text" ,aKeyFields(iCount)
						If aValueFields(iCount) =True Then
							objViewPrefences.WinRadioButton("InitialViewMode").Set 
						Else
							Fn_SISW_LifeView_ViewPreferencesOperation = False
							Set objViewPrefences = nothing
							Exit function
						End If
						
					Case "MeasuredWidth"
						If aValueFields(iCount) <> "" Then
							objViewPrefences.WinEdit("MeasuredWidth").Set aValueFields(iCount)
						End If
						
					Case "BackgroundColor"							
						objViewPrefences.WinObject("BackgroundColor").Click
						Window("ColorWindow").WinButton("Other").Click
						bFlag=Fn_SISW_LifeView_ColorOperations(sAction,aValueFields(iCount),sCalledFrom)
						Fn_SISW_LifeView_ViewPreferencesOperation=bFlag
						
					Case Else
						Fn_SISW_LifeView_ViewPreferencesOperation = False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to execute Function [ Fn_SISW_LifeView_ViewPreferencesOperation ] Invalid case [ " & sAction & " ] ")
						Set objViewPrefences = nothing
						Exit function
				End Select
				
				If Err.Number < 0 or bFlag=False Then
					Fn_SISW_LifeView_ViewPreferencesOperation = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to set value.")
					Set objViewPrefences = nothing
					Exit function
				Else
					Fn_SISW_LifeView_ViewPreferencesOperation = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully Set Value.")
				End If	
			Next
			
		Case "Verify"
			aKeyFields = dicViewPreferences.Keys
			aValueFields = dicViewPreferences.Items
			For iCount = 0 to Ubound(aKeyFields)
			
				Select Case aKeyFields(iCount)
				
					Case "flipmousedirectionforzoom","showborder","viewblackandwhiteimages","viewbaselayerinmonocolor"
						Set objCheckbox = objViewPrefences.WinCheckBox(aKeyFields(iCount))
						If aValueFields(iCount) =True Then
							If  Fn_UI_Object_GetROProperty("Fn_SetPerspective",objCheckbox,"Checked") = "ON" Then
								Fn_SISW_LifeView_ViewPreferencesOperation = True
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [ Fn_SISW_LifeView_ViewPreferencesOperation ] Successfully verified [ " & trim(aKeyFields(iCount)) & " = " & trim(aValueFields(iCount)) & " ].")
							Else
								Fn_SISW_LifeView_ViewPreferencesOperation = False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail: [ Fn_SISW_LifeView_ViewPreferencesOperation ] Failed to verify [ " & aKeyFields(iCount) & " ].")
								Set objViewPrefences = nothing
								Exit function
							End If
						ElseIf aValueFields(iCount) = False Then
		                    If  Fn_UI_Object_GetROProperty("Fn_SetPerspective",objCheckbox,"Checked") = "OFF" Then
								Fn_SISW_LifeView_ViewPreferencesOperation = True
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [ Fn_SISW_LifeView_ViewPreferencesOperation ] Successfully verified [ " & trim(aKeyFields(iCount)) & " = " & trim(aValueFields(iCount)) & " ].")
							Else
								Fn_SISW_LifeView_ViewPreferencesOperation = False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail: [ Fn_SISW_LifeView_ViewPreferencesOperation ] Failed to verify [ " & aKeyFields(iCount) & " ].")
								Set objViewPrefences = nothing
								Exit function
							End If
						End If
				
					Case "90 Degrees CW" ,"90 Degrees CCW","No Rotation","180 Degrees"
						objViewPrefences.WinRadioButton("InitialViewRotation").SetTOProperty "text" ,aKeyFields(iCount)
						If  Fn_UI_Object_GetROProperty("Fn_SetPerspective",objViewPrefences.WinRadioButton("InitialViewRotation"),"Checked") = aValueFields(iCount) Then
							Fn_SISW_LifeView_ViewPreferencesOperation = True
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [ Fn_SISW_LifeView_ViewPreferencesOperation ] Successfully verified [ " & trim(aKeyFields(iCount)) & " = " & trim(aValueFields(iCount)) & " ].")
						Else
							Fn_SISW_LifeView_ViewPreferencesOperation = False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail: [ Fn_SISW_LifeView_ViewPreferencesOperation ] Failed to verify [ " & aKeyFields(iCount) & " ].")
							Set objViewPrefences = nothing
							Exit function
						End If
				
					Case "Browse","Zoom","Seek","Zoom Area","Pan"
						objViewPrefences.WinRadioButton("InitialViewMode").SetTOProperty "text" ,aKeyFields(iCount)
						If  Fn_UI_Object_GetROProperty("Fn_SetPerspective", objViewPrefences.WinRadioButton("InitialViewMode"),"Checked") = aValueFields(iCount) Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [ Fn_SISW_LifeView_ViewPreferencesOperation ] Successfully verified [ " & trim(aKeyFields(iCount)) & " = " & trim(aValueFields(iCount)) & " ].")
							Fn_SISW_LifeView_ViewPreferencesOperation = True
						Else
							Fn_SISW_LifeView_ViewPreferencesOperation = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail: [ Fn_SISW_LifeView_ViewPreferencesOperation ] Failed to verify [ " & aKeyFields(iCount) & " ].")
							Set objViewPrefences = nothing
							Exit function
						End If
				
					Case "MeasuredWidth"
						If  Fn_UI_Object_GetROProperty("Fn_SetPerspective",objViewPrefences.WinEdit("MeasuredWidth"),"text") = aValueFields(iCount) Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [ Fn_SISW_LifeView_ViewPreferencesOperation ] Successfully verified [ " & trim(aKeyFields(iCount)) & " = " & trim(aValueFields(iCount)) & " ].")
							Fn_SISW_LifeView_ViewPreferencesOperation = True
						Else
							Fn_SISW_LifeView_ViewPreferencesOperation = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail: [ Fn_SISW_LifeView_ViewPreferencesOperation ] Failed to verify [ " & aKeyFields(iCount) & " ].")
							Set objViewPrefences = nothing
							Exit function
						End If
						
					Case "BackgroundColor"	
						bFlag = false
						objViewPrefences.WinObject("BackgroundColor").Click
						Window("ColorWindow").WinButton("Other").Click
						bFlag=Fn_SISW_LifeView_ColorOperations(sAction,aValueFields(iCount),sCalledFrom)
						If bFlag = True Then
							Fn_SISW_LifeView_ViewPreferencesOperation = True
						Else
							Fn_SISW_LifeView_ViewPreferencesOperation = False
						End If
					Case Else
						Fn_SISW_LifeView_ViewPreferencesOperation = False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to execute Function [ Fn_SISW_LifeView_ViewPreferencesOperation ] Invalid case [ " & sAction & " ] ")
						Set objViewPrefences = nothing
						Exit function
				End Select
			Next
			
		Case "GetSelectedInitialValueMode"
			arrInitialValueMode=Array("Browse","Zoom Area","Seek","Pan","Zoom")
			For iCount = 0 to uBound(arrInitialValueMode)
				objViewPrefences.WinRadioButton("InitialViewMode").SetTOProperty "text" ,arrInitialValueMode(iCount)
				If  Fn_UI_Object_GetROProperty("Fn_SISW_LifeView_ViewPreferencesOperation", objViewPrefences.WinRadioButton("InitialViewMode"),"Checked") = "ON" then
					sValue=arrInitialValueMode(iCount)
					Exit For
				Else
					sValue=""
				End if
			Next
		
		Case Else
			Fn_SISW_LifeView_ViewPreferencesOperation = False
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to execute Function [ Fn_SISW_LifeView_ViewPreferencesOperation ] Invalid case [ " & sAction & " ] ")
			Set objViewPrefences = nothing
			Exit function	
	End Select
			
	If sAction="GetSelectedInitialValueMode" Then
		objViewPrefences.WinButton("Cancel").Click
		If sValue<>"" Then
			Fn_SISW_LifeView_ViewPreferencesOperation = sValue
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully retrieved value of Initial View Mode radio button.")
		Else
			Fn_SISW_LifeView_ViewPreferencesOperation = False
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to retrieve value of Initial View Mode radio button.")
		End If
	Else	
		objViewPrefences.WinButton("OK").Click
		If Err.Number < 0 Then
			Fn_SISW_LifeView_ViewPreferencesOperation = False
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Click On [ OK ] Button.")
			Set objViewPrefences = nothing
			Exit function
		Else
			Fn_SISW_LifeView_ViewPreferencesOperation = True
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully Clicked On [ OK ] Button.")	
		End If
	End if
  Else
	Fn_SISW_LifeView_ViewPreferencesOperation = False
	Set objViewPrefences = nothing
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Open [ View Preference ] Dialog. ")
	Exit function
  End IF
	Set objViewPrefences = nothing
End Function

'****************************************    Function to perform various operation on Color dialog ***************************************
'
''Function Name		 	:	Fn_SISW_LifeView_ColorOperations()
'
''Description		    :    	Function to perform various operation on Color dialog

''Parameters		    :	1. sAction : Action to be performed
'							2.ColorInfo : Name of color  
'							3. sCalledFrom	:  Tc or Viz application
								
'
''Examples		     	:	 
'									Call Fn_SISW_LifeView_ColorOperations("Set", "LightGrey", "")
'									Call Fn_SISW_LifeView_ColorOperations("Verify", "DarkGreen", "")

'History:
'	Developer Name					Date			     Rev. No.		Reviewer		Changes Done	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Ankit Tewari 			        11-Aug-2014				1.0				
'-----------------------------------------------------------------------------------------------------------------------------------
Function Fn_SISW_LifeView_ColorOperations(sAction,ColorInfo,sCalledFrom)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_LifeView_ColorOperations"
	Dim sValue, objColor, iRed, iGreen, iBlue, arrRGBValue
	Fn_SISW_LifeView_ColorOperations = False
	Select Case sCalledFrom
		Case "TC"
			Set objColor = Window("VizWindow").Dialog("Color")
		Case "VIZ"					
			Set objColor = Window("VizMainWin").Dialog("Color")
	End Select
	
	If objColor.Exist(1)  Then
		Select Case sAction
			Case "Set"
				If ColorInfo="LightGrey" Then
					iRed="192"
					iGreen="192"
					iBlue="192"
				ElseIf ColorInfo="DarkGreen" Then
					iRed="0"
					iGreen="128"
					iBlue="128"	
				ElseIf ColorInfo="Red" Then
					iRed="255"
					iGreen="0"
					iBlue="0"	
				End If	
				If iRed<>"" and iGreen<> "" and iBlue <> "" Then
					objColor.WinEdit("Red").set iRed
					objColor.WinEdit("Green").set iGreen
					objColor.WinEdit("Blue").set iBlue
					sValue=True
				Else
					sValue=False
				End If
			Case "Verify"
				sValue=""
				sValue=objColor.WinEdit("Red").GetROProperty("text")
				sValue=sValue+":"+objColor.WinEdit("Green").GetROProperty("text")
				sValue=sValue+":"+objColor.WinEdit("Blue").GetROProperty("text")
				arrRGBValue=split(sValue,":")
				Select Case ColorInfo
					Case "LightGrey"
						If arrRGBValue(0)="192" and arrRGBValue(1)="192" and arrRGBValue(2)="192" Then
							sValue = True
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [Fn_SISW_LifeView_ColorOperations] Successfully verified colour [ " & ColorInfo & " ].")
						Else
							Fn_SISW_LifeView_ColorOperations = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail: [ Fn_SISW_LifeView_ColorOperations ] Failed to verify colour [ " & ColorInfo & " ].")
						End If
					Case "DarkGreen"
						If arrRGBValue(0)="0" and arrRGBValue(1)="128" and arrRGBValue(2)="128" Then
							sValue = True
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [Fn_SISW_LifeView_ColorOperations] Successfully verified colour [ " & ColorInfo & " ].")
						Else
							Fn_SISW_LifeView_ColorOperations = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail: [ Fn_SISW_LifeView_ColorOperations ] Failed to verify colour [ " & ColorInfo & " ].")
						End If
				End Select
		End Select	
		
		Err.Clear		
		If sAction="Set" Then
			objColor.WinButton("OK").Click
		Else
			objColor.WinButton("Cancel").Click
		End If
		If Err.Number < 0 Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to perform [ " & sAction & " ] Action.")				 
			Exit Function
		Else
			Fn_SISW_LifeView_ColorOperations = sValue
		End If
	Else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Color Dialog does not exists")
	End If

	Set objColor = Nothing
End function 

'=======================================================================================================================================================
'****************************************    Function to create 2D Markup ***************************************
'
''Function Name		 	:	Fn_SISW_LifeView_2DMarkupDelete()
'
''Description		    :    	Function to create 2D  Markup

''Parameters		    :	1. sCalledFrom : TcViz
'							2. dicMarkupDelete		:  Dictionary of parameters for Future Use
								
''Return Value		    :  	True \ False
'
''Examples		     	:	 set dicMarkupDelete = CreateObject("Scripting.Dictionary")
'							call Fn_SISW_LifeView_2DMarkupDelete("TC",  "")
'							call Fn_SISW_LifeView_2DMarkupDelete("VIZ",  "")

'History:
'	Developer Name			Date			     Rev. No.		Reviewer		Changes Done	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Reema  W			          19-Aug-2014			1.0				
'-----------------------------------------------------------------------------------------------------------------------------------
'	Reema  W			          22-Aug-2014			1.0				added Case "VIZ"
'-----------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_LifeView_2DMarkupDelete(sCalledFrom,  dicMarkupDelete)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_LifeView_2DMarkupDelete"
	Dim bReturn
	Dim objDMUtil
	Dim sMenuXMLPath
	Dim sToolsMarkupMenu
	
	Fn_SISW_LifeView_2DMarkupDelete = False
			
	'Find File Path for Lifecycle Viewer Menu XML
     sMenuXMLPath=Fn_LogUtil_GetXMLPath("Viz_Menu")

	'Select the Application , Viz or TC LCV
	Select Case sCalledFrom
		Case "TC"
			
			Set objDMUtil = Fn_SISW_LifeView_GetObject("RHSImageCanvas")

			'Extract Menu Paths from XML
			sToolsMarkupMenu = Fn_GetXMLNodeValue(sMenuXMLPath, "Toolbars2DMarkup")
		 
			'Check if the Tools:Markup menu is checked or Not
			bReturn = Fn_MenuOperation("WinMenuCheck", sToolsMarkupMenu )
			If bReturn = False Then
				'Select Tools:Markup menu 
				bReturn =  Fn_MenuOperation("WinMenuSelect", sToolsMarkupMenu )
				If bReturn = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to Select Menu [ "+sToolsMarkupMenu+" ]")
					Set objDMUtil = Nothing
					Exit Function
				End If
			End If
			Call Fn_ToolBarOperation("Click","Select","")
			objDMUtil.Click 25, 25, micLeftBtn
			bReturn = Fn_ToolBarOperation("Click","Delete the selected object permanently (Delete)","")
			If bReturn = False Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to Select Menu [ Delete the selected object permanently (Delete) ]")
				Set objDMUtil = Nothing
				Exit Function
			End If
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		Case "VIZ"
			Set objDMUtil = Fn_SISW_LifeView_GetObject("VIZ_ScratchPad").WinObject("DMUtils")
						 
			 'Extract Menu Paths from XML
			 sToolsMarkupMenu = Fn_GetXMLNodeValue(sMenuXMLPath, "EnableMarkup")
			 
			'Check if the Tools:Markup menu is checked or Not
			bReturn = Fn_SISW_LifeView_MenuOperation("CheckItemProperty",sToolsMarkupMenu+"~checked~True")
			If bReturn = False Then
				'Select Tools:Markup menu 
				bReturn =  Fn_SISW_LifeView_MenuOperation("WinMenuSelect",sToolsMarkupMenu)
				If bReturn = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to Select Menu [ "+sToolsMarkupMenu+" ]")
					Set objDMUtil = Nothing
					Exit Function
				End If
				wait(1)
			End If
			
			bReturn =  Fn_SISW_LifeView_MenuOperation("WinMenuSelect","Markup:Select")
			objDMUtil.Click 30, 30, micLeftBtn
			bReturn =  Fn_SISW_LifeView_MenuOperation("WinMenuSelect","Edit:Delete	Del")
			If bReturn = False Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to Select Menu [ Edit:Delete ]")
				Set objDMUtil = Nothing
				Exit Function
			End If
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		Case Else
			Set objDMUtil = Nothing
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to Find Required Case")
			Exit Function		
	End Select		
	'Disable Markup Menu
	If sCalledFrom = "TC"  Then
		bReturn =  Fn_MenuOperation("WinMenuSelect", sToolsMarkupMenu )	
	ElseIf sCalledFrom = "VIZ" Then
		bReturn =  Fn_SISW_LifeView_MenuOperation("WinMenuSelect",sToolsMarkupMenu)
	End If
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Successfully deleted 2D Markup")
	Set objDMUtil = Nothing
	Fn_SISW_LifeView_2DMarkupDelete = True
End Function

'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'=======================================================================================================================================================
'****************************************    Function to Merge Sessions and Load Sessions***************************************
'
''Function Name		 	:	Fn_SISW_LifeView_MergeSessions()
'
''Description		    :    	Function to Merge and Load Sessions

''Parameters		    :	   				 
'									 1. sCalledFrom	     :  Called from Teamcenter or Standalone Viz  (TC or VIZ)
'									 2. bSessionLoadReuseExisting           : True \ False to set 'Resuse existing windows if possible' checkbox on [Session Load ] dialog
'									3. bSesionLoadDontShow               :	True \ False to set 'Dont show this dialog again' checkbox on [Session Load ] dialog
'									4. sOpenOpt						: "Open" \ "Insert" to set Radiobutton on [2D load Options ] dialog
'									5. bMarkup						: "True \ False to set 'Open with markups' checkbox on [2D load Options ] dialog
'	
'Return Value		    :  	True \ False
'
''Examples		     	:	 bReturn = Fn_SISW_LifeView_MergeSessions("TC", "", "", "Open", "")
'
'History:
'	Developer Name			Date			     Rev. No.		Reviewer		Changes Done	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Reema  W			    19-Aug-2014			1.0			
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function Fn_SISW_LifeView_MergeSessions(sCalledFrom, bSessionLoadReuseExisting, bSesionLoadDontShow, sOpenOpt, bMarkup)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_LifeView_MergeSessions"
   Dim objSessionLoad,objMergeSessions
   Fn_SISW_LifeView_MergeSessions = False
   
   	Select Case sCalledFrom
		Case "TC"
			Set objMergeSessions = Window("LifeViewWin").Dialog("MergeSessions")
			Set objSessionLoad = Window("LifeViewWin").Dialog("SessionLoad")
		Case "VIZ"
			'' do nothing
	End Select


	If Fn_UI_ObjectExist("Fn_SISW_LifeView_MergeSessions",objMergeSessions )=True  Then
		objMergeSessions.WinButton("Merge").Click
		If Err.Number < 0  Then
			Fn_SISW_LifeView_MergeSessions = False
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Click On [ Merge ] Button.")
			Set objMergeSessions = nothing
			Set objSessionLoad = nothing
			Exit function
		End If	
    Else
      	Fn_SISW_LifeView_MergeSessions = False
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to find [Merge Sessions] Dialog.")
		Set objMergeSessions = nothing
		Set objSessionLoad = nothing
		Exit function
	End If
	
	If Fn_UI_ObjectExist("Fn_SISW_LifeView_MergeSessions",objSessionLoad )=True Then	
        If trim(bSessionLoadReuseExisting) <> "" Then	
			If cbool(bSessionLoadReuseExisting) = True Then
				 objSessionLoad.WinCheckBox("Reuseexistingwindows").Set "ON"	
			ElseIf cbool(bSessionLoadReuseExisting) = False Then
				 objSessionLoad.WinCheckBox("Reuseexistingwindows").Set "OFF"
			End If
			If Err.Number < 0  Then
				Fn_SISW_LifeView_MergeSessions = False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Set 'Reuse existing windows if possible' Options to [" + Cstr(bSessionLoadReuseExisting) + "]")
				Set objMergeSessions = nothing
				Set objSessionLoad = nothing
				Exit function
			End If	
		End If
		
		
		If trim(bSesionLoadDontShow) <> "" Then
			If cbool(bSesionLoadDontShow) = True Then
				 objSessionLoad.WinCheckBox("Donotshowthisdialog").Set "ON"	
			ElseIf cbool(bSesionLoadDontShow) = False Then
				 objSessionLoad.WinCheckBox("Donotshowthisdialog").Set "OFF"
			End If
			If Err.Number < 0  Then
				Fn_SISW_LifeView_MergeSessions = False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Set 'Do not Show this Dialog Again' Options to [" + Cstr(bSesionLoadDontShow) + "]")
				Set objMergeSessions = nothing
				Set objSessionLoad = nothing
				Exit function
			End If	
		End If
		
		objSessionLoad.WinButton("OK").Click
		If Err.Number < 0  Then
			Fn_SISW_LifeView_MergeSessions = False
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Click On [ OK ] Button.")
			Set objMergeSessions = nothing
			Set objSessionLoad = nothing
			Exit function
		End If	
	Else
		Fn_SISW_LifeView_MergeSessions = False
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to find [Session Load ] Dialog.")
		Set objMergeSessions = nothing
		Set objSessionLoad = nothing
		Exit function
	End If	
	
'	bReturn = fn_SISW_LifeView_Open2D3DDocument(sCalledFrom, sOpenOpt , bMarkup)
'	If bReturn = False Then
'		Fn_SISW_LifeView_MergeSessions = False
'		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Load Document.")
'		Set objMergeSessions = nothing
'		Set objSessionLoad = nothing
'		Exit function
'	End If	
	
	Set objMergeSessions = nothing
	Set objSessionLoad = nothing
	Fn_SISW_LifeView_MergeSessions = True
End Function
'*********************************************************		Function to action perform on MarkupLayer tree		**********************************************************************
'Function Name		:				Fn_SISW_LifeView_MarkupLayerTreeNodeOperation
'Description			 :		 		Function to action perform on MarkupLayer tree	

'Parameters			   :	 			1) sWindowName: Valid Window Name
											
'Return Value		   : 				TRUE \ FALSE

'Pre-requisite			:		 		MarkupLayer Tree Should be opened

'Examples				:				msgbox Fn_SISW_LifeView_MarkupLayerTreeNodeOperation("Exist","20140825-162643","")
'													msgbox Fn_SISW_LifeView_MarkupLayerTreeNodeOperation("Exist","20140825-162643:Layer1","")

'History:
'										Developer Name			Date			Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Reema W			25-Aug-2014				1.0								No			
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Priyanka K		28-Jul-2015				1.0		Added new case "GetXYCoOrdinates"	Vivek A
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Function Fn_SISW_LifeView_MarkupLayerTreeNodeOperation(StrAction,StrNodeName,StrMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_LifeView_MarkupLayerTreeNodeOperation"
	Dim objJavaWin, objJApplet, iPath, sWinTitle, objMarkupLayerTreeView
	Dim xCord, yCord, iNodeDiff, WShell, MyClipboard, iHeight, iCounter, bReturn, sClipBoardText
	
	Set objJavaWin =  Fn_UI_ObjectCreate("Fn_SISW_LifeView_MarkupLayerTreeNodeOperation",JavaWindow("DefaultWindow"))
	Set objJApplet =  Fn_UI_ObjectCreate("Fn_SISW_LifeView_MarkupLayerTreeNodeOperation",Window("TcVizStructureManager").JavaWindow("JApplet"))
	Set objMarkupLayerTreeView = JavaWindow("TcVizMainWin").JavaObject("MarkupLayerTreeView")
	
	'Calling UI GetRO property function for geting the window title value asString
	sWinTitle=  Fn_UI_Object_GetROProperty("Fn_SISW_LifeView_MarkupLayerTreeNodeOperation",objJavaWin , "title")

	Fn_SISW_LifeView_MarkupLayerTreeNodeOperation = False
	Select Case StrAction
		' TC112-2015071500-28_07_2015-VivekA-Porting-Added new case "GetXYCoOrdinates"
		'-----------Case : GetXYCoOrdinates => to get the XY Cord of node from the Tree Table
		Case "GetXYCoOrdinates"
				xCord = 140
				iNodeDiff = 15
				Set WShell = CreateObject("WScript.Shell")
				Set MyClipboard = CreateObject("Mercury.Clipboard")
				iHeight = objMarkupLayerTreeView.GetROProperty("height")
				For iCounter = 110 To iHeight Step iNodeDiff
					MyClipboard.Clear
					yCord = iCounter
					objMarkupLayerTreeView.Click xCord, yCord, "LEFT"
					wait 1
					objMarkupLayerTreeView.Click xCord, yCord, "LEFT"
					wait 1
					MyClipboard.Clear
					bReturn = Fn_KeyBoardOperation("SendKey", "^(c)")
				 	wait 1
					If bReturn = True Then
						wait 0,200
						WShell.SendKeys "{ESC}"
						wait 0,200
						sClipBoardText = MyClipboard.GetText
						sClipBoardText = replace(Trim(sClipBoardText), vbtab, "")
						If Trim(sClipBoardText) = Trim(StrNodeName) Then
							Fn_SISW_LifeView_MarkupLayerTreeNodeOperation = xCord &"~"& yCord
							Exit For
						ElseIf sClipBoardText = "" AND iCounter>150 Then
							Fn_SISW_LifeView_MarkupLayerTreeNodeOperation = False
							Exit For						
						End If
					Else
						Exit For
					End If
				Next
				Set objMarkupLayerTreeView = Nothing
				Exit Function
		' - - - - - - - - - - Existance of Node
		Case "Exist"  'TC112-2015071500-28_07_2015-VivekA-Porting-As per Design change as JavaTree is changed to WinObject.
				If Instr(sWinTitle, "Structure Manager")>0 Then
					StrNodeName = "#0:"  & StrNodeName
					iPath = Fn_JavaTree_NodeIndex("Fn_SISW_LifeView_MarkupLayerTreeNodeOperation",objJApplet,"MarkupLayerTree",StrNodeName)
					If iPath>=0 Then
						Fn_SISW_LifeView_MarkupLayerTreeNodeOperation = True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Node [" + StrNodeName + "] Exist in MarkupLayerTree")
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Node [" + StrNodeName + "] Not exist in MarkupLayerTree")
						Fn_SISW_LifeView_MarkupLayerTreeNodeOperation = False
					End If
				ElseIf Instr(sWinTitle, "My Teamcenter")>0 Then
					bReturn = Fn_SISW_LifeView_MarkupLayerTreeNodeOperation("GetXYCoOrdinates",StrNodeName,"")
					If bReturn <> False Then
						Fn_SISW_LifeView_MarkupLayerTreeNodeOperation = True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Node [" + StrNodeName + "] Exist in MarkupLayerTree")
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Node [" + StrNodeName + "] Not exist in MarkupLayerTree")
						Fn_SISW_LifeView_MarkupLayerTreeNodeOperation = False
					End If				
				End If
				'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case Else
				Fn_SISW_LifeView_MarkupLayerTreeNodeOperation = False
	End Select

	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), StrAction &" Sucessfully completed on Node [" + StrNodeName + "] of JavaTree of function Fn_SISW_LifeView_MarkupLayerTreeNodeOperation")
	Set objJApplet = nothing
	Set objJavaWin = nothing
End Function
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'****************************************    Function to perform various operation on Navigator object( for 2d Images) ***************************************
'
''Function Name		 	:	Fn_SISW_LifeView_NavigatorOperations()
'
''Description		    :    	Function to perform various operation on Navigator object( for 2d Images)

''Parameters		    :	1. sCalledFrom : Select the Application , Viz or TC LCV
'							2. sAction : Action to be performed
'							3. dicNavInfo		:  for Future use
								
'
''Examples		     	:	 
'									Set dicNavInfo = CreateObject( "Scripting.Dictionary" )
'									msgbox Fn_SISW_LifeView_NavigatorOperations("TC", "DragDropRubberBand", dicNavInfo)
'									msgbox Fn_SISW_LifeView_NavigatorOperations("TC", "PopupMenuSelect","","Fit All")
'									msgbox Fn_SISW_LifeView_NavigatorOperations("TC", "Exist", "","")
'
'History:
'	Developer Name					Date			     Rev. No.		Reviewer		Changes Done	
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Ankit T 			          27-Aug-2014				1.0												
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function Fn_SISW_LifeView_NavigatorOperations(sCalledFrom, sAction, dicNavInfo,sPopupMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_LifeView_NavigatorOperations"
	Dim objNav,objDeviceReplay, aAction, bFlag
	Dim xParentCord, yParentCord, l,t,r,b, bTextLocation
	Fn_SISW_LifeView_NavigatorOperations = False

	Select Case sCalledFrom
		Case "TC"
			Set objNav = Fn_SISW_LifeView_GetObject("Navigator")
		Case "VIZ"
				'do nothing
	End Select
	If objNav.Exist(1) AND sCalledFrom = "TC" Then
			Select Case sAction
					
				Case "DragDropRubberBand"
					Set objDeviceReplay = CreateObject("Mercury.DeviceReplay")
					objDeviceReplay.DragAndDrop (cInt(objNav.GetROProperty("abs_x")) +(cInt(objNav.GetROProperty("width"))/2)-dicNavInfo("Drag_X")),(cInt(objNav.GetROProperty("abs_y")) + (cInt(objNav.GetROProperty("height"))/2)-dicNavInfo("Drag_Y")), (cInt(objNav.GetROProperty("abs_x")) +(cInt(objNav.GetROProperty("width"))/2)+dicNavInfo("Drag_X")),(cInt(objNav.GetROProperty("abs_y")) + (cint(objNav.GetROProperty("height"))/2)+dicNavInfo("Drag_Y")), 0
					Set objDeviceReplay = Nothing
					
					If Err.Number < 0 Then
						 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to perform [ " & sAction & " ] Action.")
						 Set objNav = Nothing
						 Exit Function
					Else
						Fn_SISW_LifeView_NavigatorOperations = True
						 Exit Function
					End If
					
				Case "PopupMenuSelect"
					'call Fn_TabFolder_Operation("Select", "Navigator", "")
					Call Fn_TabFolder_Operation("DoubleClickTab", "Navigator", "")
					wait 1
					Window("TcVizStructureManager").WinObject("ImageNavigator").Click 50,50,micRightBtn
					If sPopupMenu = "Zoom In" Then
						sPopupMenu = "Zoom"
					End If
					bTextLocation = Window("TcVizStructureManager").GetTextLocation(sPopupMenu, l,t,r,b, false)
					If bTextLocation = False Then
						Set objNav = Nothing
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to perform [ " & sAction & " ] Action.")
						Exit Function
					End If
					
					Select Case sPopupMenu
						Case "Zoom"
							xParentCord = Cint(l+r)/2
						Case Else
							xParentCord = Cint(l)-5
					End Select
					
					yParentCord =(t+b) / 2 
					
					Window("TcVizStructureManager").Click xParentCord, yParentCord, micLeftBtn
					wait 1
					If sPopupMenu <> "Preferences" Then
						call Fn_TabFolder_Operation("DoubleClickTab", "Navigator", "")	
					End If			
					If err.Number = 0 Then
						Fn_SISW_LifeView_NavigatorOperations=True
					Else
						Set objNav = Nothing
						Fn_SISW_LifeView_NavigatorOperations=False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to perform [ " & sAction & " ] Action.")
						Exit Function
					End If
					
				Case "Exist"
					bFlag = Fn_TabFolder_Operation("Exist", "Navigator", "")
					If objNav.Exist(1) and bFlag = True Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [Fn_SISW_LifeView_NavigatorOperations] Successfully verified [Navigator] Tab Exists.")
						Fn_SISW_LifeView_NavigatorOperations = True
					Else
						Set objNav = Nothing
						Fn_SISW_LifeView_NavigatorOperations = False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to perform [ " & sAction & " ] Action.")
						Exit function
					End If
					
				Case Else
						Fn_SISW_LifeView_NavigatorOperations = False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Invalid Action Case")
						Set objImageCanvas = Nothing
						Exit function
			End Select
			
	End If
	 Set objNav = Nothing
End function 

'*********************************************************		Function to action perform on Preferences Dialog in navigator Tab		**********************************************************************
'Function Name		:				Fn_SISW_LifeView_PreferencesOperation
'Description			 :		 		Function to action perform on Preferences Dialog in navigator Tab

'Parameters			   :	 			1) strAction: Valid Action name
'										2) sButton: Name of button to click
'										3) sResrv:For future use	
'Return Value		   : 				TRUE \ FALSE

'Pre-requisite			:		 		Preferences Dialog must be open

'Examples				:				msgbox Fn_SISW_LifeView_PreferencesOperation("Exist","OK","")

'History:
'	Developer Name			Date			Rev. No.						Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Ankit Tewari		28-Aug-2014				1.0								No			
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Function Fn_SISW_LifeView_PreferencesOperation(strAction,sButton,sResrv)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_LifeView_PreferencesOperation"
	Dim objDialog

	Set objDialog = Window("LifeViewWin").Dialog("Preferences")
	  Fn_SISW_LifeView_PreferencesOperation = False
	Select Case StrAction
		' - - - - - - - - - - Existance of Dialog
		Case "Exist"
				If objDialog.Exist(2) Then
					Fn_SISW_LifeView_PreferencesOperation = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verified [Preferences] dialog Exist.")
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully performed [ " & sAction & " ] Action.")
				Else
					 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to verify [Preferences] dialog Exist.")
					 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to perform [ " & sAction & " ] Action.")
					  Fn_SISW_LifeView_PreferencesOperation = False
				End If
				'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case Else
				Fn_SISW_LifeView_PreferencesOperation = False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Invalid Action Case")
				Set objDialog = Nothing
				Exit function
	End Select
	
	If objDialog.WinButton(sButton).Exist(2) Then
		objDialog.WinButton(sButton).Click	
	End If
	
	Set objDialog = nothing
End Function

'=======================================================================================================================================================
'****************************************    Function to perform operations on the Nodes of the TC Vis view WinObject  ***************************************
'
''Function Name		 	:	Fn_SISW_TCVisView_TreeTableNodeOperation()
'
''Description		    :    	Function to create, Verify Product Views 

''Parameters		    :	   1. sAction                : Tree Table object
'									 2. sNodeName	     :  Node Name (Do not pass complete path)
'									 3. sReserve           : Reserve
'									 5. sPopupMenu       : Popup Menu
								
''Return Value		    :  	True \ False
'
''Examples		     	:	 bReturn = Fn_SISW_TCVisView_TreeTableNodeOperation("Select',"33000047/A;1-fishing_reel (View):33000055/A;1-spool_assembly (View)","","")
'								  bReturn = Fn_SISW_TCVisView_TreeTableNodeOperation("Exist',"33000047/A;1-fishing_reel (View):33000055/A;1-spool_assembly (View)","","") 

'History:
'	Developer Name				Date			     Rev. No.		Reviewer		Changes Done	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Pranav Ingle			16-Oct-2013			1.0				
'-----------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_TCVisView_TreeTableNodeOperation(sAction, sNodeName, sReserve, sPopupMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_TCVisView_TreeTableNodeOperation"
	Dim iRowIndex, objHPTreeView, bTextLocation
	Dim oDeviceReplay,aNodes,iCnt
	
	Dim xParentCord,iNodeDiff,WShell,MyClipboard
	Dim iCounter,bReturn,sClipBoardText,arrCord
	
	Fn_SISW_TCVisView_TreeTableNodeOperation = False		
				
	Set objHPTreeView = Window("VizMainWin").WinObject("HPTreeView")
	Set WShell = CreateObject("WScript.Shell")
	Select Case sAction
	'------------------------Case : GetXYCoOrdinates => to get the XY Cord of node from the Tree Table--------------------------------------------
		Case "GetXYCoOrdinates"
			xParentCord = 100
			iNodeDiff = 10
			
			Set MyClipboard = CreateObject("Mercury.Clipboard")
			'iHeight = Window("VizMainWin").WinObject("HPTreeView").GetROProperty("height")
			iHeight = Fn_UI_Object_GetROProperty("Fn_SISW_TCVisView_TreeTableNodeOperation",Window("VizMainWin").WinObject("HPTreeView"), "height")
			For iCounter = 50 To iHeight Step iNodeDiff
				MyClipboard.Clear
				yParentCord = iCounter
				Window("VizMainWin").WinObject("HPTreeView").Click xParentCord, yParentCord,micLeftBtn
				wait 1
				Window("VizMainWin").WinObject("HPTreeView").Click xParentCord, yParentCord,micLeftBtn
				wait 1
				WShell.SendKeys "+^(a)"
				wait 1
				Window("VizMainWin").WinObject("HPTreeView").Click xParentCord, yParentCord, micRightBtn
				wait 1
				MyClipboard.Clear
			 	WShell.SendKeys "{DOWN}"
			 	wait 1
			 	bReturn = Fn_SISW_Window_ContextMenu_Operation("Select", "Copy", "")
			 	wait(1)
				If bReturn = True Then
					wait 0,200
					WShell.SendKeys "{ESC}"
					wait 0,200
					sClipBoardText = MyClipboard.GetText
					sClipBoardText = replace(Trim(sClipBoardText), vbtab, "")
					If Trim(sClipBoardText) = Trim(sNodeName) Then
						Fn_SISW_TCVisView_TreeTableNodeOperation = xParentCord &"~"& yParentCord
						Exit For						
					End If
				End If
			Next
			Exit Function
		'------------------------Case : Select => to select the node from the Tree Table--------------------------------------------
		Case "Exist", "Select", "PopupMenuSelect"
		
			'bTextLocation = objHPTreeView.GetTextLocation(sNodeName, l,t,r,b, false)
			bReturn = Fn_SISW_TCVisView_TreeTableNodeOperation("GetXYCoOrdinates", sNodeName, "", "")
			'If bTextLocation = False Then
			'	Set objHPTreeView = Nothing
			'	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to verify Item ["&sNodeName&"] exict in Tree View")
			'	Exit Function
			'End If
			If bReturn = False Then
				Set objHPTreeView = Nothing
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to verify Item ["&sNodeName&"] exict in Tree View")
				Exit Function
			End If
			
			xParentCord = Cint(l+r)/2
			yParentCord = Cint(t+b)/2
			
			If sAction = "Select" Then
				Window("VizMainWin").WinObject("HPTreeView").Click xParentCord, yParentCord, micLeftBtn
				wait 1	
			ElseIf sAction = "PopupMenuSelect" Then
				Window("VizMainWin").WinObject("HPTreeView").Click xParentCord, yParentCord, micLeftBtn
				wait 1
				Window("VizMainWin").WinObject("HPTreeView").Click xParentCord, yParentCord, micRightBtn
				wait 1
				
				bTextLocation = Window("VizMainWin").GetTextLocation(sPopupMenu, l,t,r,b, false)
				If bTextLocation = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to verify Menu ["& sPopupMenu &"] exict in RMB Menu")
					Exit Function
				End If
				
				xParentCord = Cint(l+r)/2
				yParentCord = Cint(t+b)/2
				
				Window("VizMainWin").Click xParentCord, yParentCord, micLeftBtn
			End If

			If err.Number <> 0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to perform [ " & sAction & " ] Action.")
				Exit Function
			End If
		'------------------------Case : Find in Visview--------------------------------------------		
		 Case "Find"      
                    Set ObjFind=Window("VizMainWin").Dialog("Find")
						If ObjFind.exist(2) = True Then
						  ObjFind.WinEdit("Find What:").Set ""
						   sNodeName=Left(sNodeName,8)
						   ObjFind.WinEdit("Find What:").Set sNodeName
						   wait 1
						    If  ObjFind.WinRadioButton("Text in Assembly Workspace").GetROProperty("checked") = "OFF" Then

						   	  ObjFind.WinRadioButton("Text in Assembly Workspace").Set "ON"
						   End If
						   wait 1
						   If  ObjFind.WinRadioButton("All").GetROProperty("checked") = "OFF" Then

						   	 ObjFind.WinRadioButton("All").Set 
						   End If
						   wait 1
                          
						   ObjFind.WinButton("Find Next").Click
                          
						  if ObjFind.Dialog("Find").Exist(2) = True Then
						    if  instr(ObjFind.Dialog("Find").Static("ErrorMsg").getroproperty("text"),"The item you searched for was not found.")>0 then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to searched  for Item")
								ObjFind.Dialog("Find").WinButton("OK").Click
								bReturn=false
								Exit Function
						    End if 
						  End If
						 ObjFind.WinButton("Cancel").Click
						
						If Err.Number < 0 Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to searched  for Item")
								bReturn=false
						End If 
						
				End if
			
		'------------------------Case : Select => to select the layer from the Tree Table--------------------------------------------
		Case "SelectLayer","PopupMenuSelectLayer"
		
			bReturn = Fn_SISW_TCVisView_TreeTableNodeOperation("GetXYCoOrdinates", sNodeName, "", "")
			wait 0,700
			If bReturn = False Then
				Set objHPTreeView = Nothing
				Set WShell = Nothing
				Exit Function
			End If
			arrCord = Split(bReturn, "~" )
			xParentCord = CInt(arrCord(0))
			yParentCord = Cint(arrCord(1))
			If sAction = "SelectLayer"  Then
				objHPTreeView.Click xParentCord, yParentCord, micLeftBtn
				WShell.SendKeys "{ESC}"
			ElseIf sAction = "PopupMenuSelectLayer" Then
				objHPTreeView.Click xParentCord, yParentCord, micRightBtn
				wait 1
				WShell.SendKeys "{DOWN}"
			 	wait 1
				bReturn = Fn_SISW_Window_ContextMenu_Operation("Select", sPopupMenu, "")
				If bReturn = false Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to perform [ " & sAction & " ] Action.")
					Exit Function
				End If
			End If
			
			If err.Number <> 0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to perform [ " & sAction & " ] Action.")
				Exit Function
			End If
	'--------------------------------------------------------------------
		Case "RenameLayer"
		
			bReturn = Fn_SISW_TCVisView_TreeTableNodeOperation("GetXYCoOrdinates", sNodeName, "", "")
			wait 0,700
			If bReturn = False Then
				Set objHPTreeView = Nothing
				Set WShell = Nothing
				Exit Function
			End If
			arrCord = Split(bReturn, "~" )
			xParentCord = CInt(arrCord(0))
			yParentCord = Cint(arrCord(1))
			
			objHPTreeView.DblClick xParentCord, yParentCord, micLeftBtn
			wait 1
			WShell.SendKeys sReserve
			wait 0,500
			WShell.SendKeys "{ENTER}"
			
			If err.Number <> 0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to perform [ " & sAction & " ] Action.")
				Exit Function
			End If			
		'------------------------Case : Deselect => to Deselect the node from the Tree Table--------------------------------------------
		Case "Deselect" ' For Future Use
		'-------------------TC11.5(20180122.00)_NewDevelopment_PoonamC_15Feb2018 :( Added Case to multi select nodes) -----------------
		Case "Multiselect"
			Set oDeviceReplay = CreateObject("Mercury.DeviceReplay")
			Const VK_CONTROL = 29
			'Split the nodes
			aNodes = Split(sNodeName,"~")
			For iCnt = 0 To UBound(aNodes)
				If iCnt = 0 Then
						bReturn = Fn_SISW_TCVisView_TreeTableNodeOperation("Select",aNodes(iCnt), "", "")
						If bReturn = False Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select [ " & aNodes(iCnt) & " ].")
							Exit Function
						End If
						oDeviceReplay.KeyDown VK_CONTROL
				Else
						bReturn = Fn_SISW_TCVisView_TreeTableNodeOperation("Select",aNodes(iCnt), "", "")
						If bReturn = False Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select [ " & aNodes(iCnt) & " ].")
							oDeviceReplay.KeyUp VK_CONTROL
							Exit Function
						End If
				End If
			Next
			oDeviceReplay.KeyUp VK_CONTROL
		'--------------------------------------------------------------------
	End Select	
	Set objHPTreeView = Nothing
	Set WShell = Nothing
	Fn_SISW_TCVisView_TreeTableNodeOperation = True
End Function


'=======================================================================================================================================================
'****************************************    Function to create GDT Markup ***************************************
'
''Function Name		 	:	Fn_SISW_LifeView_GDTMarkupCreate()
'
''Description		    :    	Function to create GDT Markup

''Parameters		    :	1. sCalledFrom : TcViz
'									 2. sType		: Type of Markup (GDTAnnotationEditor,etc)
'									 3. dicMarkupCreate		:  Dictionary of parameters
								
''Return Value		    :  	True \ False
'
''Examples		     	:	 Set dicMarkupCreate = CreateObject("Scripting.Dictionary")

'									dicMarkupCreate("X_Drag") = 29
'									dicMarkupCreate("Y_Drag") = 30
'									dicMarkupCreate("X_Drop") = 100
'									dicMarkupCreate("Y_Drop") = 54
'									bReturn = Fn_SISW_LifeView_GDTMarkupCreate("TC_PSE", "AnchoredGDTAnnotationEditor", dicMarkupCreate)
'History:
'	Developer Name			Date			     Rev. No.		Reviewer		Changes Done	
'-----------------------------------------------------------------------l--------------------------------------------------------------------------------------------------------------------------------------------------------
'	Ankit Tewari		19-Sep-2014			1.0				
'-----------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_LifeView_GDTMarkupCreate(sCalledFrom, sType, dicMarkupCreate)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_LifeView_GDTMarkupCreate"
	Dim bReturn
	Dim objWin, objDMUtil
	Dim Xcoordinate, Ycoordinate, objDeviceReplay
	
	Fn_SISW_LifeView_GDTMarkupCreate = False

	'Select the Application , Viz or TC LCV
	Select Case sCalledFrom
		Case "TC"
			'Do nothing
		Case "TC_PSE"
			Set objWin = Window("TcVizStructureManager")
			Set objDMUtil = objWin.WinObject("3DImageViewer")
		Case "VIZ"
			'Do nothing
	End Select

	Select Case sCalledFrom
			Case "TC", "VIZ"
				'Do Nothing
			Case "TC_PSE"
					If Fn_SISW_TcViz_ToolbarOperation("IsButtonChecked", "3D GDT Markup", "GDT Markup") =False Then
							call Fn_SISW_TcViz_ToolbarOperation("ButtonClick", "3D GDT Markup", "GDT Markup")
					End If

	End Select
	'Select Type of Markup
	Select Case sType
		'------------------------------------------Create Anchored GDT Annotation Editor----------------------------------------
		Case "AnchoredGDTAnnotationEditor"
		
			If Fn_SISW_TcViz_ToolbarOperation("IsButtonChecked", "3D GDT Markup", "Anchor Mode") =False Then
				call Fn_SISW_TcViz_ToolbarOperation("ButtonClick", "3D GDT Markup", "Anchor Mode")
			End If
			
			If Fn_SISW_TcViz_ToolbarOperation("IsButtonChecked", "3D GDT Markup", "GDT Annotation Editor") =False Then
				call Fn_SISW_TcViz_ToolbarOperation("ButtonClick", "3D GDT Markup", "GDT Annotation Editor")
			End If
			
			bReturn = Fn_SISW_PSE_GDTAnnotationEditor("Set","Feature Control Frame","", "OK", "","")
			If bReturn = False Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to perform operation on dialog")
				Set objWin = Nothing
				Set objDMUtil = Nothing
				Exit Function
			End If
			
			Xcoordinate =  cInt((objDMUtil.GetROProperty("abs_x") +objDMUtil.GetROProperty("width")/2)) + cint(dicMarkupCreate("X_Drag"))
			Ycoordinate =  cInt((objDMUtil.GetROProperty("abs_y") + objDMUtil.GetROProperty("height")/2)) + cint(dicMarkupCreate("Y_Drag"))	
			
			Set objDeviceReplay = CreateObject("Mercury.DeviceReplay")
			objDeviceReplay.MouseClick Xcoordinate, Ycoordinate, micLeftbtn
					
			Err.clear

			objDeviceReplay.DragAndDrop Xcoordinate+5, Ycoordinate-5, Xcoordinate+cint(dicMarkupCreate("X_Drop")), Ycoordinate+cint(dicMarkupCreate("Y_Drop")), 0
			If Err.Number < 0 Then
				 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Drag and Drop in the Image")
				 Set objWin = Nothing
				 Set objDMUtil = Nothing
				 Exit Function
			End If
			
		Case "StackMode"
			'Implement as required

		Case "CopyGDTAnnotation"
			'Implement as required

		Case "PasteGDTAnnotation"
			'Implement as required

		Case "NewLayer"
			'Implement as required
			
		Case "Preferences"
			'Implement as required
		
		Case Else
			Set objWin = Nothing
			Set objDMUtil = Nothing
			Set objDialog = Nothing
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to Find Required Case")
			Exit Function			
	End Select
	
	'Disable Markup Menu
	Select Case sCalledFrom
		Case "TC", "VIZ"
		'Implement as required
					
		Case "TC_PSE"
		If Fn_SISW_TcViz_ToolbarOperation("IsButtonChecked", "3D GDT Markup", "GDT Markup") =True Then
			call Fn_SISW_TcViz_ToolbarOperation("ButtonClick", "3D GDT Markup", "GDT Markup")
		End If
	End Select
	Set objWin = Nothing
	Set objDMUtil = Nothing
	Set objDialog = Nothing
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Successfully Created GDT Markup")
	Fn_SISW_LifeView_GDTMarkupCreate = True
End Function

'=======================================================================================================================================================
'****************************************    Function to perform operations on the Nodes of the TC Vis view WinObject  ***************************************
'
''Function Name		 	:	Fn_SISW_LifeView_AssemblyTreeOperation()
'
''Description		    :    	Function to create, Verify Product Views 

''Parameters		    :	   1. sAction                : Tree Table object
'									 2. sNodeName	     :  Node Name (Do not pass complete path)
'									 3. sReserve           : Reserve
'									 5. sPopupMenu       : Popup Menu
								
''Return Value		    :  	True \ False
'
''Examples		     	:	 1. "GetXYCoOrdinates"
'						Returns XCord  & YCord  Cord of any node.  e.g.   "150~275" 
'
'						2. Select: 
'						Note:  to operate on markup nodes  please get xCord and yCord of node  by using case "GetXYCoOrdinates"

'						 bReturn = Fn_SISW_LifeView_AssemblyTreeOperation("", "Select","Markup", "Ellipse", "", XCord, YCord,"", "")
'						 bReturn = Fn_SISW_LifeView_AssemblyTreeOperation("", "Select","Part", "000032/A;1-Test", "", "", "", "", "")
'						 bReturn = Fn_SISW_LifeView_AssemblyTreeOperation("", "Select","Layer", "Layer1", "", "", "", "", "")

'						3. Exist
'						 bReturn = Fn_SISW_LifeView_AssemblyTreeOperation("", "Exist","Markup", "Ellipse", "", "", "", "", "")
'						4. PopupmMenuSelect
'						 bReturn = Fn_SISW_LifeView_AssemblyTreeOperation("", "PopupMenuSelect","Layer", "Layer1", "", "", 150, 275, "New Layer") 
'						5. Rename
'						 bReturn = Fn_SISW_LifeView_AssemblyTreeOperation("", "Rename","Layer", "Layer2", "Layer1", "", 150, 275, "") 
'						6. RMB_Snapshot
'						 bReturn = Fn_SISW_LifeView_AssemblyTreeOperation("", "RMB_Snapshot","", "", "", "", "", "", "Save as Teamcenter Product View") 
'						7. IsSelected
'						bReturn = Fn_SISW_LifeView_AssemblyTreeOperation("", "IsSelected","Layer", "Layer1", "", "", "", "", "")
'						8. UnloadAll
'						bReturn = Fn_SISW_LifeView_AssemblyTreeOperation("", "UnloadAll","", "", "", "", "", "", "")
'						8. LoadAll
'						bReturn = Fn_SISW_LifeView_AssemblyTreeOperation("", "LoadAll","", "", "", "", "", "", "")
'						9. SaveAllLayers
'						bReturn = Fn_SISW_LifeView_AssemblyTreeOperation("", "SaveAllLayers","", "", "TestSaveAllLayers", "", "", "", "")
'						10. RemoveLayer
'						bReturn = Fn_SISW_LifeView_AssemblyTreeOperation("", "RemoveLayer","", "Layer2", "", "", "", "", "")
'						11. SaveLayer
'						bReturn = Fn_SISW_LifeView_AssemblyTreeOperation("", "SaveLayer","", "Layer2", "", "", "", "", "")
'						12. SaveSelectedLayersAs
'						bReturn = Fn_SISW_LifeView_AssemblyTreeOperation("", "SaveSelectedLayersAs","", "Layer2", "TestSaveSelectedLayers, "", "", "", "")
'						13. DeselectAllObjects
'						bReturn = Fn_SISW_LifeView_AssemblyTreeOperation("", "DeselectAllObjects","", "", "", "", "", "", "")
'
'						NOTE : sReserve used to represents instance of snapshot to be clicked under case "RMB_2DImageSnapshot" .i.e if sReserve is "2",than it will click on second snapshot
'History:
'	Developer Name				Date			     Rev. No.		Reviewer		Changes Done	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Pranav Ingle			29-Sept-2014			1.0				
'	Reema W					25-Nov-14												Added case : DeselectAllObjects
'	Ankit Tewari				1-Dec-14										Modified Case "RMB_2DImageSnapshot" to RMB click on desired instance of snapshot.			
'-----------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_LifeView_AssemblyTreeOperation(sCalledFrom, sAction, sNodeType, sNodeName, sNewNodeName, xCord,  yCord, sReserve, sPopupMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_LifeView_AssemblyTreeOperation"

	Dim bReturn, iRowIndex, objAssemblyTreeView, arrCord
	Dim MyClipboard, iHeight, sClipBoardText, WShell
	Dim l, t, r, b, xParentCord, yParentCord,iCounter,XCounter
	Fn_SISW_LifeView_AssemblyTreeOperation = False
	Err.Clear
	Set WShell = CreateObject("WScript.Shell")
	
	Set objAssemblyTreeView = JavaWindow("LifeViewJavaWin").JavaObject("AssemblyTreeView")
		
		Select Case sAction
			'------------------------Case : GetXYCoOrdinates => to get the XY Cord of node from the Tree Table--------------------------------------------
			Case "GetXYCoOrdinates"
					xCord = 140
					iNodeDiff = 15
					
					Set MyClipboard = CreateObject("Mercury.Clipboard")
					iHeight = objAssemblyTreeView.GetROProperty("height")
					For iCounter = 50 To iHeight Step iNodeDiff
						MyClipboard.Clear
						yCord = iCounter
						objAssemblyTreeView.Click xCord, yCord, "LEFT"
						wait 1
						objAssemblyTreeView.DblClick xCord, yCord, "LEFT"
						wait 1
						MyClipboard.Clear
						objAssemblyTreeView.Click xCord, yCord, "RIGHT"
						wait 1
					 	WShell.SendKeys "{DOWN}"
					 	wait 1
						bReturn =  Fn_SISW_Window_ContextMenu_Operation("Select", "Copy", "")
						If bReturn = True Then
							wait 0,700
							WShell.SendKeys "{ESC}"
							wait 0,700
							sClipBoardText = MyClipboard.GetText
							sClipBoardText = replace(Trim(sClipBoardText), vbtab, "")
							If Trim(sClipBoardText) = Trim(sNodeName) Then
								Fn_SISW_LifeView_AssemblyTreeOperation = xCord &"~"& yCord
								Exit For
							End If
						Else
							Exit For
						End If
					Next
					Exit Function
			'------------------------Case : Select => to select the node from the Tree Table--------------------------------------------
			Case "Exist"
					If lCase(sNodeType) = "part" Then
						bReturn = objAssemblyTreeView.Object.getViewer().findPart(sNodeName)
						If bReturn = 0 Then
							Set objAssemblyTreeView = Nothing
							Set WShell = Nothing
							Exit Function
						End If
					ElseIf lCase(sNodeType) = "layer" Then
						bReturn = objAssemblyTreeView.Object.getViewer().getLayerIndex(sNodeName)
						If bReturn = "-1" Then
							Set objAssemblyTreeView = Nothing
							Set WShell = Nothing
							Exit Function
						End If
					ElseIf lCase(sNodeType) = "markup" Or lCase(sNodeType) = "models"Then
						bReturn = Fn_SISW_LifeView_AssemblyTreeOperation("", "GetXYCoOrdinates", "Markup", sNodeName, "", "", "", "", "")
						wait 0,700
						If bReturn = False Then
							Set objAssemblyTreeView = Nothing
							Set WShell = Nothing
							Exit Function
						End If
					Else
						Set objAssemblyTreeView = Nothing
						Set WShell = Nothing
						Exit Function
					End If
			Case "Select"
					If lCase(sNodeType) = "part" Then
						objAssemblyTreeView.Object.getViewer().select(sNodeName)
						
					ElseIf lCase(sNodeType) = "layer" Then
						bReturn = objAssemblyTreeView.Object.getViewer().getLayerIndex(sNodeName)
						If bReturn = -1 Then
							Set objAssemblyTreeView = Nothing
							Set WShell = Nothing
							Exit Function
						End If
						objAssemblyTreeView.Object.getViewer().setActiveLayer(Cint(bReturn))
					ElseIf lCase(sNodeType) = "markup" Or lCase(sNodeType) = "models"Then
						If xCord = "" OR yCord = "" Then
							bReturn = Fn_SISW_LifeView_AssemblyTreeOperation("", "GetXYCoOrdinates", "Markup", sNodeName, "", "", "", "", "")
							wait 0,700
							If bReturn = False Then
								Set objAssemblyTreeView = Nothing
								Set WShell = Nothing
								Exit Function
							End If
							arrCord = Split(bReturn, "~" )
							xCord = CInt(arrCord(0))
							yCord = Cint(arrCord(1))
						End If
						objAssemblyTreeView.Click xCord, yCord, "LEFT"
					Else
						Set objAssemblyTreeView = Nothing
						Set WShell = Nothing
						Exit Function
					End If
			Case "LoadImage", "UnloadImage"  '' added case to check or uncheck check box in Assembly tree
				bReturn = Fn_SISW_LifeView_AssemblyTreeOperation("", "GetXYCoOrdinates", "", sNodeName, "", "", "", "", "")
				If bReturn = False Then
					Set objAssemblyTreeView = Nothing
					Set WShell = Nothing
					Exit Function
				End If
				arrCord = Split(bReturn, "~" )
				xCord = CInt(arrCord(0)) -80
				yCord = Cint(arrCord(1))
				objAssemblyTreeView.Click xCord, yCord, "LEFT"

			Case "PopupMenuSelect"
				If xCord = "" OR yCord = "" Then
					bReturn = Fn_SISW_LifeView_AssemblyTreeOperation("", "GetXYCoOrdinates", sNodeType, sNodeName, "", "", "", "", "")
					wait 0,700
					If bReturn = False Then
						Set objAssemblyTreeView = Nothing
						Set WShell = Nothing
						Exit Function
					End If
					arrCord = Split(bReturn, "~" )
					xCord = CInt(arrCord(0))
					yCord = Cint(arrCord(1))
				End If
				
				'objAssemblyTreeView.Click xCord, yCord, "LEFT"
				wait 1
				objAssemblyTreeView.Click xCord, yCord, "RIGHT"
				wait 1
				WShell.SendKeys "{DOWN}"
				wait 0,800
				bReturn = Fn_SISW_Window_ContextMenu_Operation("Select", sPopupMenu, "")
				wait 0,700
				If bReturn = False Then
					Set objAssemblyTreeView = Nothing
					Set WShell = Nothing
					Exit Function
				End If
				
			Case "Rename"
					If xCord = "" OR yCord = "" Then
						bReturn = Fn_SISW_LifeView_AssemblyTreeOperation("", "GetXYCoOrdinates", sNodeType, sNodeName, "", "", "", "", "")
						If bReturn = False Then
							Set objAssemblyTreeView = Nothing
							Set WShell = Nothing
							Exit Function
						End If
						arrCord = Split(bReturn, "~" )
						xCord = CInt(arrCord(0))
						yCord = Cint(arrCord(1))
					End If
					
					objAssemblyTreeView.DblClick xCord, yCord, "LEFT"
					wait 1
					WShell.SendKeys sNewNodeName
					wait 0,500
					WShell.SendKeys "{ENTER}"
					
			Case "IsSelected"
					If lCase(sNodeType) = "part" Then
						bReturn = objAssemblyTreeView.Object.getViewer().IsSelected(sNodeName)
						If bReturn = false Then
							Set objAssemblyTreeView = Nothing
							Set WShell = Nothing
							Exit Function
						End If
					ElseIf lCase(sNodeType) = "layer" Then
						bReturn = objAssemblyTreeView.Object.getViewer().getActiveLayer()
						If bReturn <> sNodeName Then
							Set objAssemblyTreeView = Nothing
							Set WShell = Nothing
							Exit Function
						End If
						objAssemblyTreeView.Object.getViewer().setActiveLayer(Cint(bReturn))
					ElseIf lCase(sNodeType) = "markup" Then
						' Not Implemented yet
						
'						If xCord = "" OR yCord = "" Then
'							bReturn = Fn_SISW_LifeView_AssemblyTreeOperation("", "GetXYCoOrdinates", "Markup", sNodeName, "", "", "", "", "")
'							If bReturn = False Then
'								Exit Function
'							End If
'							arrCord = Split(bReturn, "~" )
'							xCord = arrCord(0)
'							yCord = arrCord(1)
'						End If
'						objAssemblyTreeView.Click xCord, yCord, "LEFT"
					Else
						Set objAssemblyTreeView = Nothing
						Set WShell = Nothing
						Exit Function
					End If
					
			'------------------------Case : IsChecked => to Check checkbox of assemblytree is set or not--------------------------------------------			
			Case "IsChecked" 
                    bReturn =  objAssemblyTreeView.Object.getViewer().isVisible(sNodeName)
                    If Err.Number < 0  Then
                        Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to get Checkbox State.")
                        Exit Function
                    End If            
                    If lCase(bReturn) = "false" Then
                        Set WShell=nothing
                        Set objAssemblyTreeView=nothing
						Exit Function 
                    End If
             '------------------------Case : Find--------------------------------------------	       
               Case "Find"      
                    Set ObjFind= Window("LifeCycleViewWinTeamcenter").Dialog("Find")
						If ObjFind.exist(2) = True Then
						  ObjFind.WinEdit("Find What:").Set ""
						   sNodeName=Left(sNodeName,8)
						   ObjFind.WinEdit("Find What:").Set sNodeName
						   wait 1
						   If  ObjFind.WinCheckBox("Match whole word only").GetROProperty("checked") = "OFF" Then
						   	  ObjFind.WinCheckBox("Match whole word only").Set "ON"
						   End If
						    If  ObjFind.WinRadioButton("Teamcenter").GetROProperty("checked") = "OFF" Then
						   	  ObjFind.WinRadioButton("Teamcenter").Set "ON"
						   End If
						   wait 1
                           If  ObjFind.WinCheckBox("WinCheckBox").GetROProperty("checked") = "OFF" Then
						   	  ObjFind.WinCheckBox("WinCheckBox").Set "ON"
						   End If
						   wait 1
						   ObjFind.WinButton("Find Next").Click
                          
						  if ObjFind.Dialog("Find").Exist(2) = True Then
						    if  instr(ObjFind.Dialog("Find").Static("ErrorMsg").getroproperty("text"),"The item you searched for was not found.")>0 then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to searched  for Item")
								ObjFind.Dialog("Find").WinButton("OK").Click
								bReturn=false
								Exit Function
						    End if 
						  End If
						 ObjFind.WinButton("Cancel").Click
						
						If Err.Number < 0 Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to searched  for Item")
								bReturn=false
						End If 
						
				End if
			'------------------------Case : RMB_Snapshot => to perform RMB operation on Snapshot Tab--------------------------------------------
			Case "RMB_Snapshot","RMB_2DImageSnapshot"
				If  sReserve <> "" Then 'sReserve represents instance of snapshot to be clicked.i.e if sReserve is "2",than it will click on second snapshot
					XCounter=0
					For iCounter=1 to sReserve
						If  iCounter = 1 Then
							XCounter=0
						Else
							XCounter=XCounter+65
						End If
					Next
					objAssemblyTreeView.Click 50+XCounter, 50, "LEFT"
					wait 1
					objAssemblyTreeView.Click 50+XCounter, 50, "RIGHT"
					wait 1
				Else
					objAssemblyTreeView.Click 50, 50, "LEFT"
					wait 1
					objAssemblyTreeView.Click 50, 50, "RIGHT"
					wait 1
				End If
				
				If sAction ="RMB_Snapshot"  Then
					WShell.SendKeys "{DOWN}"
					wait 1
					bReturn = Fn_SISW_Window_ContextMenu_Operation("Select", sPopupMenu, "")
					If bReturn = False Then
						Set objAssemblyTreeView = Nothing
						Set WShell = Nothing
						Exit Function
					End If
				ElseIf sAction ="RMB_2DImageSnapshot"  Then
					Select Case sPopupMenu
						Case "Add"
							wait(3)
							JavaWindow("LifeViewJavaWin").JavaObject("AssemblyTreeView").Click 1, 1, micRightBtn
							wait(1)
							If JavaWindow("LifeViewJavaWin").InsightObject("SnapshotsAdd").Exist = False Then
								JavaWindow("LifeViewJavaWin").InsightObject("SnapshotsAdd_1").Click 10, 10, micLeftBtn
							Else
								JavaWindow("LifeViewJavaWin").InsightObject("SnapshotsAdd").Click 10, 10, micLeftBtn								
							End If
						Case "Save as Teamcenter Snapshot..."
							wait(1)
							If JavaWindow("LifeViewJavaWin").InsightObject("SaveAsTcSnapshot").Exist = False Then
								JavaWindow("LifeViewJavaWin").InsightObject("SaveAsTcSnapshot_1").Click 10, 10, micLeftBtn
							Else
								JavaWindow("LifeViewJavaWin").InsightObject("SaveAsTcSnapshot").Click 10, 10, micLeftBtn								
							End If
						Case else
							bTextLocation =Window("LifeViewWin").GetTextLocation(sPopupMenu, l,t,r,b, false)
							xParentCord = (l+r)/2
							yParentCord= (t+b)/2
							Window("LifeViewWin").Click xParentCord, yParentCord, micLeftBtn
							If err.number < 0 Then
								Set objAssemblyTreeView = Nothing
								Set WShell = Nothing
								Exit Function
							End If
					End Select
				Else
					Set objAssemblyTreeView = Nothing
					Set WShell = Nothing
					Exit Function
				End If
			'------------------------Case : UnloadAll => To Unload Assembly --------------------------------------------
			Case "UnloadAll"
				objAssemblyTreeView.Object.getviewer.UnloadAll
				
			'------------------------Case : LoadAll => To Load Assembly --------------------------------------------
			Case "LoadAll"
				objAssemblyTreeView.Object.getviewer.LoadAll
				
			'------------------------Case : DeselectAllObjects => To Deselect all object --------------------------------------------
			Case "DeselectAllObjects"
				objAssemblyTreeView.Object.getviewer.deselectAllObjects
			
			'------------------------Case : CreateNewLayer => To Create New Layer -------------------------------------------
			Case "CreateNewLayer"	' sNewNodeName - is used to pass "new Layer name" to function 
				objAssemblyTreeView.Object.getviewer.CreateNewLayer(sNewNodeName)
				
			'------------------------Case : SaveAllLayers => To save all Layers--------------------------------------------
			Case "SaveAllLayers"	 ' sNewNodeName - is used to pass "new save dataset name" to function 
				objAssemblyTreeView.Object.getviewer.SaveAllLayers
				wait 3
				If Window("LifeViewWin").Dialog("NewMarkupDataset").Exist(10) Then
					Window("LifeViewWin").Dialog("NewMarkupDataset").WinEdit("MarkupName").Set sNewNodeName
					wait 1
					Window("LifeViewWin").Dialog("NewMarkupDataset").WinButton("OK").Click 1, 1, micLeftBtn
				End If
				wait 2	
			'------------------------Case : SaveLayer => to save 1 layer--------------------------------------------
			Case "SaveLayer"	 ' sNewNodeName - is used to pass "new save dataset name" to function 
'				bReturn = objAssemblyTreeView.Object.getviewer.getLayerIndex(sNodeName)
'				objAssemblyTreeView.Object.getviewer.savelayer bReturn,sNodeName

'					-----------Commented code and used option 'SaveSelectedLayers' instead of 'SaveLayer' as per discussion with vallari  TC112(2015071500)
				If xCord = "" OR yCord = "" Then
					bReturn = Fn_SISW_LifeView_AssemblyTreeOperation("", "GetXYCoOrdinates", sNodeType, sNodeName, "", "", "", "", "")
					If bReturn = False Then
						Set objAssemblyTreeView = Nothing
						Set WShell = Nothing
						Exit Function
					End If
					arrCord = Split(bReturn, "~" )
					xCord = CInt(arrCord(0))
					yCord = Cint(arrCord(1))
				End If
				'objAssemblyTreeView.Click xCord, yCord, "LEFT"
				wait 1
				objAssemblyTreeView.Click xCord, yCord, "RIGHT"
				wait 1
				WShell.SendKeys "{DOWN}"
				wait 0,200
				bReturn = Fn_SISW_Window_ContextMenu_Operation("Select", "Save Selected Layers", "")
				If bReturn = False Then
					Set objAssemblyTreeView = Nothing
					Set WShell = Nothing
					Exit Function
				End If
			'------------------------Case : SaveSelectedLayerAs => to save 1 layerAs--------------------------------------------
			Case "SaveSelectedLayersAs"
				If xCord = "" OR yCord = "" Then
					bReturn = Fn_SISW_LifeView_AssemblyTreeOperation("", "GetXYCoOrdinates", sNodeType, sNodeName, "", "", "", "", "")
					If bReturn = False Then
						Set objAssemblyTreeView = Nothing
						Set WShell = Nothing
						Exit Function
					End If
					arrCord = Split(bReturn, "~" )
					xCord = CInt(arrCord(0))
					yCord = Cint(arrCord(1))
				End If
				
				'objAssemblyTreeView.Click xCord, yCord, "LEFT"
				wait 1
				objAssemblyTreeView.Click xCord, yCord, "RIGHT"
				wait 1
				WShell.SendKeys "{DOWN}"
				wait 0,200
				bReturn = Fn_SISW_Window_ContextMenu_Operation("Select", "Save Selected Layers As...", "")
				If bReturn = False Then
					Set objAssemblyTreeView = Nothing
					Set WShell = Nothing
					Exit Function
				End If
				
				If Window("LifeViewWin").Dialog("NewMarkupDataset").Exist(5) = False Then
					JavaWindow("LifeViewJavaWin").InsightObject("SaveSelectedLayerAs").Click 10, 10, micLeftBtn
				End If
				
				If Window("LifeViewWin").Dialog("NewMarkupDataset").Exist(10) Then
					Window("LifeViewWin").Dialog("NewMarkupDataset").WinEdit("MarkupName").Set sNewNodeName
					wait 1
					Window("LifeViewWin").Dialog("NewMarkupDataset").WinButton("OK").Click 1, 1, micLeftBtn
				Else
					Set objAssemblyTreeView = Nothing
					Set WShell = Nothing
					Exit Function
				End If
				wait 2	
			'------------------------Case : RemoveLayer => To Remove a  Layer--------------------------------------------
			Case "RemoveLayer"  ' sNewNodeName - is used to pass warning message to function 
				'objAssemblyTreeView.Object.getviewer.RemoveLayer()
				bReturn = objAssemblyTreeView.Object.getviewer.getLayerIndex(sNodeName)
				If bReturn = False Then
					Set objAssemblyTreeView = Nothing
					Set WShell = Nothing
					Exit Function
				End If
				objAssemblyTreeView.Object.getviewer.RemoveLayer(bReturn)
				wait 3
				If Window("LifeViewWin").Dialog("Warning").Exist(10) Then
					If sNewNodeName <> "" Then
						Window("LifeViewWin").Dialog("Warning").Static("Message").SetTOProperty "text", sNewNodeName
						If Not(Window("LifeViewWin").Dialog("Warning").Static("Message").Exist(1)) Then
							Window("LifeViewWin").Dialog("Warning").WinButton("OK").Click 1, 1, micLeftBtn
							Set objAssemblyTreeView = Nothing
							Set WShell = Nothing
							Exit Function
						End If
					End If
					wait 1
					Window("LifeViewWin").Dialog("Warning").WinButton("OK").Click 1, 1, micLeftBtn
				End If
				wait 2
		End Select
		
		If Err.Number <> 0 Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to perform [ " & sAction & " ] Action.")
			Set objAssemblyTreeView = Nothing
			Set WShell = Nothing
			Exit Function
		End If
		
	Set objAssemblyTreeView = Nothing
	Set WShell = Nothing
	Fn_SISW_LifeView_AssemblyTreeOperation = True
End Function

'****************************************    Function to Save As Product Teamcenter View ***************************************
'
''Function Name		 	:	Fn_SISW_LifeView_SaveAsTeamcenterProductView()
'
''Description		    :  	Function to Save As Product Teamcenter View

''Parameters		    :	1. sCalledFrom : TC/VIZ
'					2. sAction : Action To Call
'					3. sMsgInvalidAssmState : To Verify Message from Dialog Invalid Assmembly State 
'					4. sNewProductViewName : To enter new Product View Dataset
'					5. sInformation : Check Message from Information Dialog
'					6. sReserve : Reserve for Future Use

								
''Return Value		    :  	True \ False
'
''Examples		     	:	Fn_SISW_LifeView_SaveAsTeamcenterProductView("TC", "SaveAsTCProductView", "", "TestProduct", "Successfully created Teamcenter Product View.", "")

'History:
'	Developer Name			Date			Rev. No.		Reviewer		Changes Done	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Pranav Ingle			 29-Sept-2014			1.0		
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_LifeView_SaveAsTeamcenterProductView(sCalledFrom, sAction, sMsgInvalidAssmState, sNewProductViewName, sInformation, sReserve)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_LifeView_SaveAsTeamcenterProductView"
	Dim bReturn, objInvalidAssemblyState , objNewTeamcenterProduct, objInformation
	Fn_SISW_LifeView_SaveAsTeamcenterProductView = False

	'Get all the required window references
	Select Case sCalledFrom
		Case "TC"
			Set objInvalidAssemblyState =  Window("LifeViewWin").Dialog("InvalidAssemblyState")
			Set objNewTeamcenterProduct = Window("LifeViewWin").Dialog("NewTeamcenterProductViewDataset")
			Set objInformation = Window("LifeViewWin").Dialog("Information")
		Case "VIZ"
			
	End Select
		
	err.clear
	Select Case sAction
		Case "SaveAsTCProductView"
			If objInvalidAssemblyState.Exist(5) Then		
				'If sMsgInvalidAssmState <> "" Then      '------------------For Future Use
					
					'Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Operate Menu [" + sMenu + "]")
					'Set objInvalidAssemblyState =  Nothing
					'Set objNewTeamcenterProduct = Nothing
					'Set objInformation = Nothing
					'Exit Function
				'End If
			
			
				objInvalidAssemblyState.WinButton("Proceed").Click 1,1,micLeftBtn
				If err.number < 0 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Cllick  Proceed Button of [ Invalid Assembly State ] Dialog")
					Set objInvalidAssemblyState =  Nothing
					Set objNewTeamcenterProduct = Nothing
					Set objInformation = Nothing
					Exit Function		
				End If
				Wait 2
			End If
			
			If objNewTeamcenterProduct.Exist(5) Then		
				
				If sNewProductViewName <> "" Then
					objNewTeamcenterProduct.WinEdit("NewProductName").Set sNewProductViewName
					If err.number < 0 Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Enter NewProductViewName [ NewTeamcenterProductView ] Dialog")
						Set objInvalidAssemblyState =  Nothing
						Set objNewTeamcenterProduct = Nothing
						Set objInformation = Nothing
						Exit Function		
					End If
					Wait 1
				End If
				
					
				objNewTeamcenterProduct.WinButton("OK").Click 1,1,micLeftBtn
				If err.number < 0 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Cllick  OK Button of [ NewTeamcenterProductView ] Dialog")
					Set objInvalidAssemblyState =  Nothing
					Set objNewTeamcenterProduct = Nothing
					Set objInformation = Nothing
					Exit Function		
				End If
				Wait 5
			End If
			
			If objInformation.Exist(5) Then		
				
				If sInformation <> "" Then
					objInformation.Static("sMsg").SetTOProperty "text", sInformation
					If Not (objInformation.Static("sMsg").Exist(1)) Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Verify message "&sInformation&" from [ Information ] Dialog")
						Set objInvalidAssemblyState =  Nothing
						Set objNewTeamcenterProduct = Nothing
						Set objInformation = Nothing
						Exit Function		
					End If
					Wait 1
				End If
					
				objInformation.WinButton("OK").Click 1,1,micLeftBtn
				If err.number < 0 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Cllick  OK Button of [ Information ] Dialog")
					Set objInvalidAssemblyState =  Nothing
					Set objNewTeamcenterProduct = Nothing
					Set objInformation = Nothing
					Exit Function		
				End If
				Wait 1
			End If
	End Select
	
	
Set objInvalidAssemblyState =  Nothing
Set objNewTeamcenterProduct = Nothing
Set objInformation = Nothing

Fn_SISW_LifeView_SaveAsTeamcenterProductView = True	

End Function

'****************************************    Function to View Product Structure in Lyfe Cycle Viewer ***************************************
'
''Function Name		 	:	Fn_SISW_LifeView_ViewProductStructure()
'
''Description		    :  	Function to Save As Product Teamcenter View

''Parameters		    :	1. sCalledFrom : TC/VIZ
'					2. sAction : Action To Call
'					3. sReserve : Reserve for Future Use
								
''Return Value		    :  	True \ False
'
''Examples		     	:	Fn_SISW_LifeView_ViewProductStructure("TC", "ConfigureAnUpdatedStructure", "")
'						Fn_SISW_LifeView_ViewProductStructure("TC", "LoadTheStaticAsSaved", "")

'History:
'	Developer Name			Date			Rev. No.		Reviewer		Changes Done	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Pranav Ingle			 30-Sept-2014			1.0		
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_LifeView_ViewProductStructure(sCalledFrom, sProdStructureLoadView, sReserve)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_LifeView_ViewProductStructure"
	Dim bReturn, objProductStructure
	Fn_SISW_LifeView_ViewProductStructure = False

	'Get all the required window references
	Select Case sCalledFrom
		Case "TC"
			Set objProductStructure =  Dialog("ProductStructure")
		Case "VIZ"
			
	End Select
	
	Err.Clear
	
	If objProductStructure.Exist(5) Then		
		
		If sProdStructureLoadView <> "" Then
			objProductStructure.WinRadioButton(sProdStructureLoadView).Set
			If err.number < 0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Select Radio Button  "&sProdStructureLoadView&" on [ Product Structure ] Dialog")
				Set objProductStructure =  Nothing
				Exit Function		
			End If
			Wait 0, 200
		End If
			
		objProductStructure.WinButton("OK").Click 1,1,micLeftBtn
		If err.number < 0 Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Cllick  OK Button of [ Product Structure  ] Dialog")
			Set objProductStructure =  Nothing
			Exit Function		
		End If
		Wait 1
	End If
	
Set objProductStructure =  Nothing

Fn_SISW_LifeView_ViewProductStructure = True	

End Function


''*********************************************************		Function to Perform operation on Export Image Dialog in LifeCycle Viewer	***********************************************************************
'
''Function Name		:		Fn_SISW_LifeView_ExportImage
'
''Description			 :		 Function to Perform operation on Export Image Dialog in LifeCycle Viewer
'
''Parameters		:		1.	sAction = Action To Perform
'							   			2.  sTabName = pass the value of tab you want to select .  eg "Export Image"
'										3. sButton = Name of button to click. 		eg. "OK" or "cancel" 		
'										4. dicInfo = Dictionary Object
'										5. sReserve = For Future Use  		NOTE : For Enter Name Dialog that appears after Export Image Dialog use 'sReserve'		
'										6. sCheck = "ON/OFF" To check and uncheck checkboxes.

'			  										
''Return Value		   : 				True or False 
'
''Pre-requisite			:				Export Image Dialog should be displayed in LCV.
'
'
''Examples				:		 			Case "Export Image" 'Case handled according to Tab of the Dialog
'											Dim dicInfo
'											Set dicInfo=CreateObject("Scripting.Dictionary")
											'dicInfo("Export") = "Image"
											'dicInfo("Width") = "3.9400"
											'dicInfo("Height") = "5.3200"
											'dicInfo("Units") ="inches"
											'dicInfo("Resolution") = "300"
											'dicInfo("ColorMode") = "Color"
											'dicInfo("Type") = "bmp"
											'dicInfo("RatioOriginaIImage") = "100"

'											
												'Fn_SISW_LifeView_ExportImage("Set","Export Image",dicInfo, "OK/Cancel", "ExportedImage1234","ON/OFF")

'											dicInfo("Width") = "False"
'											dicInfo("Height") = "False"
											'Fn_SISW_LifeView_ExportImage("VerifyEnabled","Export Image",dicInfo, "OK/Cancel", "","")
'
'History:
'										Developer Name				Date						Rev. No.							Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Ankit Tewari					29-Sep-2014					1.0									
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_LifeView_ExportImage(sAction, sTabName , dicInfo , sButton, sReserve, sCheck)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_LifeView_ExportImage"
	Dim DictItems,DictKeys,iCount
	Dim iCounter,sValue,objDialog,objDialogEnterName

	Fn_SISW_LifeView_ExportImage=False
	iCount = 0
	Set ObjDialog = Window("LifeViewWin").Dialog("ExportImage")
	Set ObjDialogEnterName = Window("LifeViewWin").Dialog("EnterName")
	If ObjDialog.Exist(2) <> True Then
		Set ObjDialog=nothing
		Exit Function
	Else
		ObjDialog.WinTab("Tab").Select sTabName
	End If

   	Select Case sTabName
		Case "Export Image"
			Select Case sAction
				Case "Set"
					DictKeys = dicInfo.Keys
					DictItems = dicInfo.Items
					For iCounter = 0 to Ubound(DictKeys)							
						Select Case DictKeys(iCounter)
							Case "Width","Height","RatioOriginaIImage"
								If DictItems(iCounter) <> "" Then
									'ObjDialog.WinEdit(DictKeys(iCounter)).Set DictItems(iCounter)	
									If ObjDialog.WinEdit(DictKeys(iCounter)).Exist(1) Then
										ObjDialog.WinEdit(DictKeys(iCounter)).Set DictItems(iCounter)	
									End If
								End If
								
							Case "Export","Units","Resolution","ColorMode","Type"
								If DictItems(iCounter) <> "" Then
									'ObjDialog.WinComboBox(DictKeys(iCounter)).Select DictItems(iCounter)
									If ObjDialog.WinComboBox(DictKeys(iCounter)).Exist(1) Then
										ObjDialog.WinComboBox(DictKeys(iCounter)).Select DictItems(iCounter)
									End If
								End If
								
							Case Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_LifeView_ExportImage ] Invalid case [ " & sAction & " ].")
								Exit function
						End Select
					Next
					
					If sCheck = "ON" Then
						ObjDialog.WinCheckBox("Retaindpi").Set "ON"
					End If
					
					If Err.Number < 0 Then
						Fn_SISW_LifeView_ExportImage=False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_LifeView_ExportImage ] Failed to performed [ " & sAction & " ] Action on [ "+sTabName+" ].")
						ObjDialog=nothing
						Exit Function				
					Else
						Fn_SISW_LifeView_ExportImage=True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_LifeView_ExportImage ] Successfully performed [ " & sAction & " ] Action on [ "+sTabName+" ].")
					End If
	'------------------------------------------------------------------------------------------------------------------------------------------------------------------				
				Case "VerifyEnabled"
					DictKeys = dicInfo.Keys
					DictItems = dicInfo.Items
					For iCounter = 0 to Ubound(DictKeys)							
						Select Case DictKeys(iCounter)
							Case "Width","Height","RatioOriginaIImage"
								sValue = ObjDialog.WinEdit(DictKeys(iCounter)).GetROProperty("enabled")
								If DictItems(iCounter) <> "" Then
									If sValue = CBool(DictItems(iCounter)) Then										
										iCount=iCount+1
									End If									
								End If
								
								If iCount = Ubound(DictKeys)+1 Then
									Fn_SISW_LifeView_ExportImage=True
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_LifeView_ExportImage ] Successfully performed [ " & sAction & " ] Action on [ "+sTabName+" ].")	
								End If
								
							Case Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_LifeView_ExportImage ] Invalid case [ " & sAction & " ].")
								Exit function
						End Select
					Next
			
	'------------------------------------------------------------------------------------------------------------------------------------------------------------------			
				Case Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_LifeView_ExportImage ] Invalid case [ " & sAction & " ].")
					Exit function
			End Select
	'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		Case "Watermark"
			'Implement as required
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		Case else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_LifeView_ExportImage ] Invalid case [ " & sAction & " ].")
			Exit function
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	End Select
	
	If sButton <> "" Then
		If ObjDialog.WinButton(sButton).Exist(2) = True Then
			ObjDialog.WinButton(sButton).Click	
		else
			Set ObjDialog = nothing
			Exit function
		End if
	End If

	Err.Clear
	If ObjDialogEnterName.Exist(1) = True Then
		If sReserve <> "" Then
			'ObjDialogEnterName.WinEdit("Name").Set sReserve
			ObjDialogEnterName.WinEdit("Name").Set ""
			ObjDialogEnterName.WinEdit("Name").Type sReserve
			wait 1
			ObjDialogEnterName.WinButton("OK").Click	
		End If	
	Else
		Set ObjDialogEnterName = nothing
	End If
	
	If Err.Number = 0  Then
		Fn_SISW_LifeView_ExportImage = True
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [ Fn_SISW_LifeView_ExportImage ] executed successfuly with case [ " & sAction & " ].")
	else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_LifeView_ExportImage ] Failed to executed with case [ " & sAction & " ].")
	End If

	Set ObjDialog = nothing
	Set ObjDialogEnterName = nothing
End Function 

'****************************************    Function to handle Load Option Preferences dialog ***************************************
'
''Function Name		 	:	Fn_SISW_LifeView_LoadOptionPrefOperation()
'
''Description		    :  	Function to do operations on Load Option Preferences dialog.

''Parameters		    :	1. sAction : Action to be performed
'							2. sTab : Tab name 2D/3D/ECAD
'							3. Dim dicLoadOptPreference
'								Set dicLoadOptPreference = CreateObject("Scripting.Dictionary")
'								dicLoadOptPreference("DocumentRadioPref")="OpenDocument"
'								dicLoadOptPreference("DocumentCheckboxPref")="Askatloadtime:ON"
'								dicLoadOptPreference("MarkupCheckboxPref")="Askatloadtime:ON~Openwithmarkups:ON"
'							4. sReserve : this parameter used as "sCalledFrom" :  Called from Teamcenter or Standalone Viz  (TC or VIZ)
								
''Return Value		    :  	True \ False
'
''Examples		     	:	Fn_SISW_LifeView_LoadOptionPrefOperation("Set", "3D", dicLoadOptPreference, "OK", "")
'							Fn_SISW_LifeView_LoadOptionPrefOperation("Set", "2D", dicLoadOptPreference, "OK", "")
'History:
'	Developer Name			Date			Rev. No.		Reviewer		Changes Done	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Paresh			 	16-Oct-2013			1.0				
'-----------------------------------------------------------------------------------------------------------------------------------
'	shweta rathod		05-Dec-2016         1.1         shweta				added code to work with standalone viz viewer. 
'-----------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_LifeView_LoadOptionPrefOperation(sAction, sTab, dicLoadOptPreference, sButton, sReserve)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_LifeView_LoadOptionPrefOperation"
	'Declaring Variables
	Dim objLoadOpt,aPrefvalue,icnt,aPreferences,sMenu
	Fn_SISW_LifeView_LoadOptionPrefOperation = False
	
	If sReserve = "VIZ" then
		Set objLoadOpt = Fn_SISW_LifeView_GetObject("VIZ_LoadOptPrefrence") 'Fn_SISW_LifeView_GetObject("VIZ_PLMXML")				
	else
		Set objLoadOpt = Fn_SISW_LifeView_GetObject("LoadOptPrefrence") 		
	End if
	
	wait(1)
	If Fn_SISW_UI_Object_Operations("Fn_SISW_LifeView_LoadOptionPrefOperation","Exist", objLoadOpt,"") = False Then
		sMenu =  Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("Viz_Menu"), "LoadOptions")    'Get menu
		If sReserve = "VIZ" then			
			Call Fn_SISW_LifeView_MenuOperation("WinMenuSelect", sMenu)
		else
			Call Fn_MenuOperation("WinMenuSelect",sMenu)
		End if
		Call Fn_ReadyStatusSync(2)	
	End If
	
	If objLoadOpt.Exist(4) = True Then
		If sTab <> "" Then
			objLoadOpt.WinTab("LoadOptionstab").Select sTab
			If err.number < 0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Set ["+ sTab +"] Tab")
				Set objLoadOpt = Nothing
				Exit Function
			End If	
		End If
		
		Select Case sAction
			Case "Set"
	'Set product structure or document radio preferences
				If dicLoadOptPreference("DocumentRadioPref")<>"" Then
					Select Case dicLoadOptPreference("DocumentRadioPref")
						Case "OpenDocument"
							objLoadOpt.WinRadioButton("OpenDocument").SetTOProperty "index","0"							
							objLoadOpt.WinRadioButton("OpenDocument").Set 
						Case "InsertDocument"
							objLoadOpt.WinRadioButton("InsertDocument").SetTOProperty "index","0"							
							objLoadOpt.WinRadioButton("InsertDocument").Set 						
						Case "MergeOpenDocument"
							objLoadOpt.WinRadioButton("Mergedocument").Set 						
							wait 1
							objLoadOpt.WinRadioButton("OpenDocument").SetTOProperty "index","1"							
							objLoadOpt.WinRadioButton("OpenDocument").Set 
						Case "MergeInsertDocument"
							objLoadOpt.WinRadioButton("Mergedocument").Set
							Wait 1
							objLoadOpt.WinRadioButton("InsertDocument").SetTOProperty "index","1"							
							objLoadOpt.WinRadioButton("InsertDocument").Set 						
					End Select
				If err.number < 0 Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Set Radiobutton ["+dicLoadOptPreference("DocumentRadioPref")+"] Preferences.")
						Set objLoadOpt = Nothing
						Exit Function
					End If
				End If
			
		'set Product structure or Document checkbox preferences
				If dicLoadOptPreference("DocumentCheckboxPref")<>"" Then
					aPreferences = split(dicLoadOptPreference("DocumentCheckboxPref"),":")
						If aPreferences(0) = "Askatloadtime" Then
							objLoadOpt.WinCheckBox("AskatLoadTime").SetTOProperty "index","0"
							If lcase(aPreferences(1))="on" Then
								objLoadOpt.WinCheckBox("AskatLoadTime").Set "ON"
							Else
								objLoadOpt.WinCheckBox("AskatLoadTime").Set "OFF"					
							End If
						End If
						If err.number < 0 Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Set document checkbox ["+dicLoadOptPreference("DocumentCheckboxPref")+"] preferences.")
							Set objLoadOpt = Nothing
							Exit Function
						End If
				End If
				
				If sReserve<>"" Then
					'further use
				End If
			
		'set Markups checkbox preferences
				If dicLoadOptPreference("MarkupCheckboxPref")<>"" Then
					aPreferences = split(dicLoadOptPreference("MarkupCheckboxPref"),"~")
					For icnt = 0 To ubound(aPreferences)
						aPrefvalue = Split(aPreferences(icnt),":")
						Select Case aPrefvalue(0)
							Case "Askatloadtime"
								objLoadOpt.WinCheckBox("AskatLoadTime").SetTOProperty "index","1"
								If lcase(aPrefvalue(1))="on" Then
									objLoadOpt.WinCheckBox("AskatLoadTime").Set "ON"
								Else
									objLoadOpt.WinCheckBox("AskatLoadTime").Set "OFF"					
								End If
							Case "Openwithmarkups"
								If lcase(aPrefvalue(1))="on" Then
									objLoadOpt.WinCheckBox("Openwithmarkups").Set "ON"
								Else
									objLoadOpt.WinCheckBox("Openwithmarkups").Set "OFF"					
								End If
						End Select
						If err.number < 0 Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Set ["+aPreferences(icnt)+"] Preferences.")
							Set objLoadOpt = Nothing
							Exit Function
						End If
					Next
				End If	
			Case Else
				Set objLoadOpt = nothing
				Exit Function		
		End Select
	Else
		Set objLoadOpt = nothing
		Exit Function
	End If
		
	'click on given button
	If sButton<>"" Then
		Call Fn_UI_WinButton_Click("Fn_SISW_LifeView_LoadOptionPrefOperation",objLoadOpt,sButton,5,5,micLeftBtn)
	End If		
	
	Set objLoadOpt = nothing
	Fn_SISW_LifeView_LoadOptionPrefOperation = True
End Function


''*********************************************************		Function to Perform operation on Export Dialog in LifeCycle Viewer	***********************************************************************
'
''Function Name		:		Fn_SISW_LifeView_ExportDialogOperation
'
''Description			 :		 Function to Perform operation on Export Dialog in LifeCycle Viewer
'
''Parameters		:					1.	sAction = Action To Perform
'										2. dicInfo = Dictionary Object
'										3. sButton = Name of button to click. 		eg. "OK" or "cancel" 
'										4. sReserve = For Future Use  		NOTE : For Enter Name Dialog that appears after Export Image Dialog use 'sReserve'		

'			  										
''Return Value		   : 				True or False 
'
''Pre-requisite			:				Export Dialog should be displayed in LCV.
'
'
''Examples				:		 			
'											Dim dicInfo
'											Set dicInfo=CreateObject("Scripting.Dictionary")
'											dicInfo("File") = "C:\Block_8567+".plmxml"
'											dicInfo("FileFormat") = "Product Structure (*.plmxml)"
'											dicInfo("Hierarchy") = "(AH) Alt Hier"										
'											bReturn= Fn_SISW_LifeView_ExportDialogOperation("ExportImageandVerifyMessage", dicInfo , "OK", "")
'
'										Set dicInfo=CreateObject("Scripting.Dictionary")
'											dicInfo("File") = "C:\Block_8567+".plmxml"
'											dicInfo("FileFormat") = "Product Structure (*.plmxml)"
'											dicInfo("Hierarchy") = "(AH) Alt Hier"										
'											
'										Set dicOptionInfo=CreateObject("Scripting.Dictionary")
'											dicOptionInfo("Tab1") = "General"
'											dicOptionInfo("WinCheckBox1") = "Product structure data:ON"
'											dicOptionInfo("WinCheckBox2") = "Product view data:ON"
'											dicOptionInfo("WinCheckBox3") = "Current state of document:ON"
'											dicOptionInfo("Tab2") = "Product Structure"
'											dicOptionInfo("WinRadioButton1") = "Complete product structure:ON"
'											dicOptionInfo("WinRadioButton2") = "JT geometry creation:ON"
'											dicOptionInfo("WinButton") = "OK"
'									bReturn = Fn_SISW_LifeView_ExportDialogOperation("ExportImageandVerifyMessage",dicInfo,"OK",dicOptionInfo)
'
'History:
'										Developer Name				Date						Rev. No.							Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Ankit Tewari					16-Oct-2014					1.0									
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
''										Poonam Chopade					01-Mar-2018					1.1				Added Code to set Options in Export		TC_20180212.00_NewDevelopment_PoonamC_01Mar2018							
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_LifeView_ExportDialogOperation(sAction, dicInfo , sButton, sReserve)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_LifeView_ExportDialogOperation"
	Dim DictItems,DictKeys,iCount
	Dim iCounter,sValue,objDialog,objDialogEnterName
	Dim ObjExporterDialog
	Dim sSubAction,sProperty,ObjplmXmlExp
	
	Fn_SISW_LifeView_ExportDialogOperation=False
	iCount = 0

	Set ObjDialog = Window("LifeViewWin").Dialog("Export")
	Set ObjExporterDialog = Window("LifeViewWin").Dialog("Exporter")
	If NOT ObjDialog.exist(4) Then
		call Fn_MenuOperation("Select", "File:Export...")
	End if
	wait 3
	If NOT ObjDialog.exist(4) Then
		Fn_SISW_LifeView_ExportDialogOperation=False
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_LifeView_ExportDialogOperation ] Export Dialog is not opened.")
		ObjDialog=nothing
		Exit Function	
	End if
	Select Case sAction
		Case "ExportImageandVerifyMessage"
			DictKeys = dicInfo.Keys
			DictItems = dicInfo.Items
			For iCounter = 0 to Ubound(DictKeys)							
				Select Case DictKeys(iCounter)
					Case "File"
						If DictItems(iCounter) <> "" Then
							If ObjDialog.WinEdit(DictKeys(iCounter)).Exist(1) Then
								ObjDialog.WinEdit(DictKeys(iCounter)).Set DictItems(iCounter)	
							End If
						End If
						
					Case "FileFormat"
						If DictItems(iCounter) <> "" Then
							If ObjDialog.WinComboBox(DictKeys(iCounter)).Exist(1) Then
								ObjDialog.WinComboBox(DictKeys(iCounter)).Select DictItems(iCounter)
							End If
						End If
				
					Case "Hierarchy"
						If DictItems(iCounter) <> "" Then
							If ObjDialog.WinList(DictKeys(iCounter)).Exist(1) Then
								ObjDialog.WinList(DictKeys(iCounter)).Select DictItems(iCounter)
							End If
						End If
					Case Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_LifeView_ExportDialogOperation ] Invalid case [ " & sAction & " ].")
						Exit function
				End Select	
			Next
			'-------------------------------------------------------------------------------------------------------------------------
			'[TC11.5(20180212.00)_NewDevelopment_PoonamC_28Feb2018 : Added below code set PLMXML export options ]
			If vartype(sReserve) = 9 Then	
				'Click on Options button
				Call Fn_UI_WinButton_Click("Fn_SISW_LifeView_ExportDialogOperation", ObjDialog, "Options","","","")
				Wait 1
				Set ObjplmXmlExp = Dialog("PLMXMLexport")
				'Check Dialog existence
				If NOT ObjplmXmlExp.exist(4) Then
					Fn_SISW_LifeView_VIZ_ExportDialogOperation = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_LifeView_ExportDialogOperation ] PLMXML Export Dialog is not opened.")
					Set ObjDialog = nothing
					Set ObjplmXmlExp = Nothing
					Exit Function	
				End if
			
				DictKeys = sReserve.Keys
				DictItems = sReserve.Items
				For iCounter = 0 to Ubound(DictKeys)	
					If Instr(DictKeys(iCounter),"Tab")>0 Then
						sSubAction = "Tab"
					ElseIf Instr(DictKeys(iCounter),"WinCheckBox")>0 Then
						sSubAction = "WinCheckBox"
					ElseIf Instr(DictKeys(iCounter),"WinRadioButton")>0 Then
						sSubAction = "WinRadioButton"
					ElseIf Instr(DictKeys(iCounter),"WinButton")>0 Then
						sSubAction = "WinButton"						
					Else
						sSubAction = DictKeys(iCounter)
					End If
					sProperty = DictItems(iCounter)
					
					Select Case sSubAction
						Case "Tab"
							ObjplmXmlExp.WinTab("TabName").Select sProperty
							If Err.Number < 0 Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to select ["+sProperty+"] tab on [PLMXML export] dialog.")
								Set ObjplmXmlExp = Nothing
								Set ObjDialog = nothing
								Exit Function
							End If
							Wait 1
						Case "WinCheckBox"
							sProperty = Split(sProperty,":")
							If sProperty(0) = "Product structure data" then sProperty(0) = "Product structure"&vblf&"data"
							If sProperty(0) = "Current state of document" then sProperty(0) = "Current state of"&vblf&"document"
							If sProperty(0) = "Maintain JT file structure" then sProperty(0) = "Maintain JT file"&vblf&"structure"
							
							ObjplmXmlExp.WinCheckBox("CheckBoxName").SetTOProperty "text",sProperty(0)
							ObjplmXmlExp.WinCheckBox("CheckBoxName").Set sProperty(1)
							If Err.Number < 0 Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to set ["+sProperty(0)+"] as ["&sProperty(1)&"] on [PLMXML export] dialog.")
								Set ObjplmXmlExp = Nothing
								Set ObjDialog = nothing
								Exit Function
							End If
							Wait 1
						Case "WinRadioButton"
							sProperty = Split(sProperty,":")
							ObjplmXmlExp.WinRadioButton("Radiobutton").SetTOProperty "text",sProperty(0)
							ObjplmXmlExp.WinRadioButton("Radiobutton").Set 
							If Err.Number < 0 Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to set ["+sProperty(0)+"] as ["&sProperty(1)&"] on [PLMXML export] dialog.")
								Set ObjplmXmlExp = Nothing
								Set ObjDialog = nothing
								Exit Function
							End If
							Wait 1
						Case "WinButton"
							ObjplmXmlExp.WinButton(sProperty).Click 5,5,micLeftBtn
							If Err.Number < 0 Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to click ["+sProperty+"] button on [PLMXML export] dialog.")
								Set ObjplmXmlExp = Nothing
								Set ObjDialog = nothing
								Exit Function
							End If
							Wait 1
					End Select
				Next
				Set ObjplmXmlExp = Nothing
			End If
			'------------------------------------------------------------------------------------------------------------------------
			If sButton <> "" Then
				If ObjDialog.WinButton(sButton).Exist(2) Then
					ObjDialog.WinButton(sButton).Click	
					Fn_SISW_LifeView_ExportDialogOperation=true
				else
					Set ObjDialog = nothing
					Exit function
				End if
			End If
		
			If ObjExporterDialog.Exist(4) = True Then
				If ObjExporterDialog.Static("Message").Exist(1) Then
					ObjExporterDialog.WinButton("OK").Click	
					Fn_SISW_LifeView_ExportDialogOperation=True
				Else
					ObjExporterDialog.WinButton("OK").Click	
					Fn_SISW_LifeView_ExportDialogOperation=False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_LifeView_ExportDialogOperation ] Failed to verify [ Export Succeeeded ] Message.")
					ObjDialog=nothing
					Exit Function	
				End If
			Else
				Set ObjExporterDialog = nothing
			End If
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_LifeView_ExportDialogOperation ] Invalid case [ " & sAction & " ].")
			Exit function
	End Select
	
	If ObjDialog.exist(4) Then
		ObjDialog.WinButton("Cancel").Click	
	End if
	
	If Fn_SISW_LifeView_ExportDialogOperation <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [ Fn_SISW_LifeView_ExportDialogOperation ] executed successfuly with case [ " & sAction & " ].")
	else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_LifeView_ExportDialogOperation ] Failed to executed with case [ " & sAction & " ].")
	End If

	Set ObjDialog = nothing
	Set ObjExporterDialog = nothing
End Function


'****************************************    Function to Save PLMXML ***************************************
'
''Function Name		 	:	Fn_SISW_LifeView_FilePreferencesPLMXMLOperation()
'
''Description		    :  	Function to Save PLMXML

''Parameters		    :	
'							1. sAction : Action to be performed
'							2. sCalledFrom : TC/VIZ
'							3. sInvokeFrm : Menu/Toolbar
'							4. dicPartColorInfo		: dictinary object							
''Return Value		    :  	True \ False
'
''Examples		     	:	Set dicPartColorInfo = CreateObject( "Scripting.Dictionary" )			
'							dicPartColorInfo("partcolor") = "Red"
'							dicPartColorInfo("parttransparency") = "0%"
'							bReturn = Fn_SISW_LifeView_ConceptAppearancePartColorOperation("Set","TC", "menu" ,dicPartColorInfo)

'History:
'	Developer Name			Date			Rev. No.		Reviewer		Changes Done	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Rinki A			 13-Nov-2014			1.0		
'  
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function Fn_SISW_LifeView_ConceptAppearancePartColorOperation(sAction,sCalledFrom, sInvokeFrm ,dicPartColorInfo)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_LifeView_ConceptAppearancePartColorOperation"
   Dim objPartColor, objImageCanvas
   Dim aKeyFields, aValueFields,iCount
   Dim sXMLFilePath, sMenu
   Dim bFlag

    Fn_SISW_LifeView_ConceptAppearancePartColorOperation = False
   	Select Case sCalledFrom
		Case "TC"
			Set objPartColor = Window("LifeViewWin").Dialog("PartColor")
			Set objImageCanvas = Fn_SISW_LifeView_GetObject("RHSImageCanvas")
		   	If Fn_UI_ObjectExist("Fn_SISW_LifeView_ConceptAppearancePartColorOperation",objPartColor )=False Then
				'Invoking View Prefrences Winidow
				If lcase(sInvokeFrm) = "menu" Then
					sXMLFilePath = Fn_LogUtil_GetXMLPath("Viz_Menu")
					sMenu = Fn_GetXMLNodeValue(sXMLFilePath, "ConceptAppearancePartColor")
					call Fn_MenuOperation("WinMenuSelect", sMenu)
				Else
					Fn_SISW_LifeView_ConceptAppearancePartColorOperation = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to execute Function [ Fn_SISW_LifeView_ConceptAppearancePartColorOperation ] Invalid case [ " & sAction & " ] ")
					Set objImageCanvas = nothing
					Set objPartColor = nothing
					Exit function
				End If
		   End If
		Case "VIZ"					
			' for Futur use
	End Select

   If Fn_UI_ObjectExist("Fn_SISW_LifeView_ConceptAppearancePartColorOperation",objPartColor )=True Then
	Select Case sAction
		Case "Set"
			aKeyFields = dicPartColorInfo.Keys
			aValueFields = dicPartColorInfo.Items
			For iCount = 0 to Ubound(aKeyFields)
				Select Case lCase(aKeyFields(iCount))
					Case "partcolor"							
						objPartColor.WinButton("PartColor").Click
						Window("ColorWindow").WinButton("Other").Click
						bFlag=Fn_SISW_LifeView_ColorOperations(sAction,aValueFields(iCount),sCalledFrom)
						If bFlag = False Then
							Fn_SISW_LifeView_ConceptAppearancePartColorOperation = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to set [ " & aValueFields(iCount) &" ] color.")
							Set objImageCanvas = nothing
							Set objPartColor = nothing
							Exit function
						End If
					Case "parttransparency"							
						objPartColor.WinComboBox("PartTransparency").Select aValueFields(iCount)
					Case Else
						Fn_SISW_LifeView_ConceptAppearancePartColorOperation = False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to execute Function [ Fn_SISW_LifeView_ConceptAppearancePartColorOperation ] Invalid case [ " & aKeyFields(iCount) & " ] ")
						Set objImageCanvas = nothing
						Set objPartColor = nothing
						Exit function
				End Select
			Next
			objPartColor.Close
			wait 1
			If 	objImageCanvas.exist(2) Then
				objImageCanvas.Click 3,3
			End If
			If Err.Number < 0 Then
				Fn_SISW_LifeView_ConceptAppearancePartColorOperation = False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to set specified color to selected part.")
				Set objImageCanvas = nothing
				Set objPartColor = nothing
				Exit function
			Else
				Fn_SISW_LifeView_ConceptAppearancePartColorOperation = True
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully Set specified color to selected part.")
			End If	
		Case Else
			Fn_SISW_LifeView_ConceptAppearancePartColorOperation = False
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to execute Function [ Fn_SISW_LifeView_ConceptAppearancePartColorOperation ] Invalid case [ " & sAction & " ] ")
			Set objImageCanvas = nothing
			Set objPartColor = nothing
			Exit function	
	End Select
			
  Else
	Fn_SISW_LifeView_ConceptAppearancePartColorOperation = False
	Set objPartColor = nothing
	Set objImageCanvas = nothing
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Open [Part Color ] Dialog. ")
	Exit function
  End IF
	Set objPartColor = nothing
	Set objImageCanvas = nothing
End Function


'****************************************    Function to Save PLMXML ***************************************
'
''Function Name		 	:	Fn_SISW_LifeView_FilePreferencesPLMXMLOperation()
'
''Description		    :  	Function to Save PLMXML

''Parameters		    :	
'							1. sCalledFrom : TC/VIZ
'							2. sPLMXMLUnits		: PLM XML Units	to be selected
'							3. sPLMXML		: Save PLMXML Checkbox
'							4. sSaveInserted : Save inserted models]
'							5. sCopy		: Copy parts locally
'							6. sRetainRef		: Retain references to original product structure
'							7. sIncludeLate		: Include late loaded attributes
'							8. sAlwaysAsk		: ALways ask at save time
''Return Value		    :  	True \ False
'
''Examples		     	:	bReturn = Fn_SISW_LifeView_FilePreferencesPLMXMLOperation("TC", "", "ON", "", "", "", "", "ON")

'History:
'	Developer Name			Date			Rev. No.		Reviewer		Changes Done	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Rinki A			 13-Nov-2014			1.0		
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_LifeView_FilePreferencesPLMXMLOperation(sCalledFrom, sPLMXMLUnits, sPLMXML, sSaveInserted, sCopy, sRetainRef, sIncludeLate, sAlwaysAsk)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_LifeView_FilePreferencesPLMXMLOperation"
	Dim objPLMXMLWin
	Dim sMenuPath, sMenu
	Dim bFlag
	
	Fn_SISW_LifeView_FilePreferencesPLMXMLOperation = False
	
	'Get all the required window references
	Select Case sCalledFrom
		Case "TC"
			Set objPLMXMLWin = Fn_SISW_LifeView_GetObject("TC_PLMXML")
		Case "VIZ"
			Set objPLMXMLWin = Fn_SISW_LifeView_GetObject("VIZ_PLMXML")
	End Select
	
	If Not objPLMXMLWin.Exist(5) Then		
		'Get Viz menu file path
		sMenuPath=Fn_LogUtil_GetXMLPath("Viz_Menu")
		
		'Get menu
	    sMenu = Fn_GetXMLNodeValue(sMenuPath, "FilePrefrencesPLMXML")
		Select Case sCalledFrom
			Case "TC"
				bFlag = Fn_MenuOperation("WinMenuSelect", sMenu)
				wait(1)
			Case "VIZ"
				bFlag = Fn_SISW_LifeView_MenuOperation("WinMenuSelect", sMenu)
				wait(2)
		End Select
		
		If bFlag = False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Operate Menu [" + sMenu + "]")
			Set objPLMXMLWin = Nothing
			Exit Function
		End If
	End If
	
	If objPLMXMLWin.Exist(5) Then
		'set load Tab
		err.clear
		objPLMXMLWin.WinTab("MainTab").Select "Load", micLeftBtn
		If err.number < 0 Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Set [Load] Tab")
			Set objPLMXMLWin = Nothing

			Exit Function		
		End If
		err.clear
		'select sPLMXMLUnits
		If sPLMXMLUnits <> "" Then
			objPLMXMLWin.WinComboBox("PLMXMLUnits").Select sPLMXMLUnits
			If err.number < 0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to select [ " & sPLMXMLUnits & " ] from PLM XML Units")
				Set objPLMXMLWin = Nothing
				Exit Function		
			End If
		End If
		
		'set PLMXML Tab
		err.clear
		objPLMXMLWin.WinTab("MainTab").Select "Save", micLeftBtn
		If err.number < 0 Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Set [Save] Tab")
			Set objPLMXMLWin = Nothing
			Exit Function		
		End If
		Wait 1
		
		'Check ON/OFF selected options
		If sPLMXML <> "" Then
			err.clear
			If sPLMXML = "ON" Then
				If objPLMXMLWin.WinCheckBox("SaveExtended3D").GetROProperty("checked") = "OFF" Then
					objPLMXMLWin.WinCheckBox("SaveExtended3D").Click 2,2,micLeftBtn
				End If
			Else
				If objPLMXMLWin.WinCheckBox("SaveExtended3D").GetROProperty("checked")= "ON" Then
					objPLMXMLWin.WinCheckBox("SaveExtended3D").Click 2,2,micLeftBtn
				End If
			End If
			If err.number < 0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Set [Save extended 3D content into PLMXML] Option")
				Set objPLMXMLWin = Nothing
				Exit Function		
			End If
		End If
		Wait 1
	
		If sSaveInserted <> "" Then
			err.clear
			objPLMXMLWin.WinCheckBox("SaveInserted").Set sSaveInserted
			If err.number < 0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Set [Save inserted models] Option")
				Set objPLMXMLWin = Nothing
				Exit Function		
			End If
		End If
		Wait 1
	
		If sCopy <> "" Then
			err.clear
			objPLMXMLWin.WinCheckBox("CopyParts").Set sCopy
			If err.number < 0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Set [Copy parts locally] Option")
				Set objPLMXMLWin = Nothing
				Exit Function		
			End If
		End If
		Wait 1
	
		If sRetainRef <> "" Then
			err.clear
			objPLMXMLWin.WinCheckBox("RetainReferences").Set sRetainRef
			If err.number < 0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Set [Retain references to original product structure] Option")
				Set objPLMXMLWin = Nothing
				Exit Function		
			End If
		End If
		Wait 1
		
		If sIncludeLate <> "" Then
			err.clear
			objPLMXMLWin.WinCheckBox("IncludeLateLoaded").Set sIncludeLate
			If err.number < 0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Set [Include late loaded attributes] Option")
				Set objPLMXMLWin = Nothing
				Exit Function		
			End If
		End If
		Wait 1
		
		If sAlwaysAsk <> "" Then
			err.clear
			objPLMXMLWin.WinCheckBox("AlwaysAsk").Set sAlwaysAsk
			If err.number < 0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Set [ALways ask at save time] Option")
				Set objPLMXMLWin = Nothing
				Exit Function		
			End If
		End If
		Wait 1
		
		err.clear
		objPLMXMLWin.WinButton("OK").Click
		If err.number < 0 Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Click [OK] Button")
			Set objPLMXMLWin = Nothing
			Exit Function		
		End If
		Wait 2
	End If
	Set objPLMXMLWin = Nothing
	Fn_SISW_LifeView_FilePreferencesPLMXMLOperation = True	

End Function


''*********************************************************		Function to Perform operation on 2D Snapshot Form Dialog in LifeCycle Viewer	***********************************************************************
'
''Function Name		:		Fn_SISW_LifeView_2DSnapshotFormDialogOperation
'
''Description			 :		 Function to Perform operation on 2D Snapshot Form Dialog in LifeCycle Viewer
'
''Parameters		:						1. sCalledFrom=LCV or VIZ
'										2.sAction = Action To Perform
'										3. dicInfo = Dictionary Object
'										4. sButton = Name of button to click. 		eg. "OK" or "cancel" 
'										5. sReserve = For Future Use  		
'			  										
''Return Value		   : 				True or False 
'
''Pre-requisite			:				2D Snapshot Form Dialog should be displayed in LCV.
'
'
''Examples				:		 			
'											Dim dicInfo
'											Set dicInfo=CreateObject("Scripting.Dictionary")
'											dicInfo("Revision ID") = "B"
'											dicInfo("Page Number") = "1"										
'											bReturn= Fn_SISW_LifeView_2DSnapshotFormDialogOperation("LCV"."Set", dicInfo , "OK", "")
'
'History:
'										Developer Name				Date						Rev. No.							Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Ankit Tewari					10-Nov-2014					1.0									
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_LifeView_2DSnapshotFormDialogOperation(sCalledFrom,sAction, dicInfo , sButton, sReserve)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_LifeView_2DSnapshotFormDialogOperation"
	Dim DictItems,DictKeys
	Dim iCounter,objDialog
	Dim ObjFormDialog
	Err.clear
	Fn_SISW_LifeView_2DSnapshotFormDialogOperation=False

		   'Select the application, TC or Standalone TcViz
   	Select Case sCalledFrom
   		Case "TC"
			'------ Called From Teamcenter LCV ---------------------------------
			Set ObjDialog = Window("LifeViewWin").Dialog("2DSnapshotFormDialog")
			Set ObjFormDialog = Dialog("FormEdit")
		
   		Case "VIZ"
			'------ Called From Standalone TcVIZ ---------------------------------
   	End Select  	

	If NOT ObjDialog.exist(4) Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_LifeView_2DSnapshotFormDialogOperation ] 2D Snapshot Form Dialog is not opened.")
		ObjDialog=nothing
		Exit Function	
	End if
	
	Select Case sAction
		Case "Set"
			DictKeys = dicInfo.Keys
			DictItems = dicInfo.Items
			For iCounter = 0 to Ubound(DictKeys)							
				Select Case DictKeys(iCounter)
				
					Case "Revision ID","Page Number"
						If DictItems(iCounter) <> "" Then
							ObjDialog.WinListView("ListView").Activate(DictKeys(iCounter))
							If ObjFormDialog.exist(4) Then
								ObjFormDialog.WinEdit("Value").Set DictItems(iCounter) 
								If ObjFormDialog.WinButton("OK").Exist(2) Then
									ObjFormDialog.WinButton("OK").Click
									If Err.Number < 0 Then
			 							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Click on [OK] Button.")
			 							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : [ Fn_SISW_LifeView_2DSnapshotFormDialogOperation ] Failed to executed with case [ " & sAction & " ].")
			 							Exit Function
									End If
								End If
							End If
						End If
						
					Case Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_LifeView_2DSnapshotFormDialogOperation ] Invalid case [ " & sAction & " ].")
						Exit function			
				End Select	
			Next
		
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_LifeView_2DSnapshotFormDialogOperation ] Invalid case [ " & sAction & " ].")
			Exit function	
	End Select	
	
	If sButton <> "" Then
		If ObjDialog.WinButton(sButton).Exist(2) Then
			ObjDialog.WinButton(sButton).Click	
			If Err.Number < 0 Then
			 	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Click on ["+sButton+"] Button.")
				Exit Function
			End If
		End if
	End If
	
	Fn_SISW_LifeView_2DSnapshotFormDialogOperation=True
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [ Fn_SISW_LifeView_2DSnapshotFormDialogOperation ] executed successfuly with case [ " & sAction & " ].")

	Set ObjDialog = nothing
	Set ObjFormDialog = nothing
End Function


''*********************************************************		Function to Perform operation on 2D Loader Preference Dialog in LifeCycle Viewer	***********************************************************************
'
''Function Name		:		Fn_SISW_LifeView_2DLoaderPreferencesOperations
'
''Description			 :		 Function to Perform operation on 2D Loader Preference Dialog in LifeCycle Viewer
'
''Parameters		:						1. SCalledFrom=LCV 0r VIZ
'										2. sAction = Action To Perform
'							   			3. sTabName = pass the value of tab you want to select .  eg "Raster"
'										4. sButton = Name of button to click. 		eg. "OK" or "cancel" 		
'										5. dicInfo = Dictionary Object
'										6. sReserve = For Future Use  	
'										7.sBackValue = value of the background

'			  										
''Return Value		   : 				True or False 
'
''Pre-requisite			:				2D Loader Preference Dialog should be displayed in LCV.
'
'
''Examples				:		 			Case "Raster" 'Case handled according to Tab of the Dialog
'											'Fn_SISW_LifeView_2DLoaderPreferencesOperations("LCV","SetWhiteBackground","Raster","", "OK/Cancel", "","85")
'
'History:
'										Developer Name				Date						Rev. No.							Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Ankit Tewari					11-Nov-2014					1.0									
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_LifeView_2DLoaderPreferencesOperations(sCalledFrom,sAction, sTabName , dicInfo , sButton, sReserve,sBackValue)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_LifeView_2DLoaderPreferencesOperations"
	Dim DictItems,DictKeys
	Dim iCounter,sValue,objDialog
	Err.Clear
	Fn_SISW_LifeView_2DLoaderPreferencesOperations=False
	
	   'Select the application, TC or Standalone TcViz
   Select Case sCalledFrom
   	Case "TC"
		'------ Called From Teamcenter LCV ---------------------------------
		Set ObjDialog = Window("LifeViewWin").Dialog("2DLoaderPreferences")
		If ObjDialog.Exist(2) <> True Then
			Call Fn_MenuOperation("WinMenuSelect","File:Preferences:2D Loader...") 
		End If
   	Case "VIZ"
		'------ Called From Standalone TcVIZ ---------------------------------
   End Select  
	
	If ObjDialog.Exist(2) <> True Then
		Set ObjDialog=nothing
		Exit Function
	Else
		ObjDialog.WinTab("Tab").Select sTabName
	End If

   	Select Case sTabName
		Case "Raster"
			Select Case sAction
				Case "SetWhiteBackground" 		'Case to set background threshold value white i.e greater than sBackValue (Near to 95)
					ObjDialog.WinButton("Reset").Click
					sMax=ObjDialog.WinObject("Background").GetROProperty("x")
					sValue=ObjDialog.WinEdit("Value").GetROProperty("text")
					If sValue <= "95" Then
						Do While sValue < sBackValue
 							ObjDialog.WinObject("Background").Click sMax
 							If Err.Number < 0 Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_LifeView_2DLoaderPreferencesOperations ] Failed to performed [ " & sAction & " ] Action on [ "+sTabName+" ].")
								ObjDialog=nothing
								Exit Function
							End If	
 							sValue=ObjDialog.WinEdit("Value").GetROProperty("text")
						Loop	
					End If		
	'------------------------------------------------------------------------------------------------------------------------------------------------------------------			
				Case Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_LifeView_2DLoaderPreferencesOperations ] Invalid case [ " & sAction & " ].")
					Exit function
			End Select
	'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		Case "Text"
			'Implement as required
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		Case else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_LifeView_2DLoaderPreferencesOperations ] Invalid case [ " & sAction & " ].")
			Exit function
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	End Select
	
	If sButton <> "" Then
		If ObjDialog.WinButton(sButton).Exist(2) = True Then
			ObjDialog.WinButton(sButton).Click	
			If Err.Number < 0 Then
				 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to click on ["+sButton+"] Button.")
				 Set ObjDialog = nothing
				 Exit Function
			End If
		End if
	End If
	
	Fn_SISW_LifeView_2DLoaderPreferencesOperations=True
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [ Fn_SISW_LifeView_2DLoaderPreferencesOperations ] executed successfuly with case [ " & sAction & " ].")
	Set ObjDialog = nothing
End Function 


''*********************************************************		Function to Perform operation on 2D Markup Preferences Dialog in LifeCycle Viewer	***********************************************************************
'
''Function Name		:		Fn_SISW_LifeView_2DMarkupPreferencesOperations()
'
''Description			 :		 Function to Perform operation on 2D Markup Preferences Dialog in LifeCycle Viewer
'
''Parameters		:						1. sCalledFrom=LCV or VIZ
'										1. sAction = Action To Perform
'							   			2. sTabName = pass the value of tab you want to select .  eg "Text"
'										3. sButton = Name of button to click. 		eg. "OK" or "cancel" 		
'										4. dicInfo = Dictionary Object
'										5. sReserve = For Future Use  		
'			  										
''Return Value		   : 				True or False 
'
''Pre-requisite			:				2D Markup Preferences Dialog should be displayed in LCV.
'
'
''Examples				:		 			Case "Text"		 'Case handled according to Tab of the Dialog
'											Dim dicInfo
'											Set dicInfo=CreateObject("Scripting.Dictionary")
											'dicInfo("FontName") = "MS Gothic"									
											'Fn_SISW_LifeView_2DMarkupPreferencesOperations("LCV","Exist","Text",dicInfo, "OK/Cancel", "")
'History:
'										Developer Name				Date						Rev. No.							Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Ankit Tewari					12-Nov-2014					1.0									
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_LifeView_2DMarkupPreferencesOperations(sCalledFrom,sAction, sTabName , dicInfo , sButton, sReserve)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_LifeView_2DMarkupPreferencesOperations"
	Dim DictItems,DictKeys,l,t,r,b
	Dim iCounter,sValue,objDialog,iCount,WshShell
	Err.clear
	Fn_SISW_LifeView_2DMarkupPreferencesOperations=False

		   'Select the application, TC or Standalone TcViz
  	 Select Case sCalledFrom
   		Case "TC"
			'------ Called From Teamcenter LCV ---------------------------------
			Set ObjDialog =Window("LifeViewWin").Dialog("2DMarkupPreferences")
			If ObjDialog.Exist(2) <> True Then
				Call Fn_ToolBarOperation("Click","Markup Preferences","")
			End If
   		Case "VIZ"
			'------ Called From Standalone TcVIZ ---------------------------------
   	End Select  
   	
	If ObjDialog.Exist(2) <> True Then
		Set ObjDialog=nothing
		Exit Function
	Else
		ObjDialog.WinTab("Tab").Select sTabName
	End If

   	Select Case sTabName
		Case "Text"
			Select Case sAction			
				Case "Exist"
					DictKeys = dicInfo.Keys
					DictItems = dicInfo.Items
					For iCounter = 0 to Ubound(DictKeys)							
						Select Case DictKeys(iCounter)
							Case "FontName"
								If DictItems(iCounter) <> "" Then
									ObjDialog.WinEdit("Fontname").Set DictItems(iCounter)
									sValue = ObjDialog.WinObject("FontnameList").GetTextLocation(DictItems(iCounter),l,t,r,b,False)
									wait 2
									sValue = ObjDialog.GetTextLocation(DictItems(iCounter),l,t,r,b,False)
									If sValue = True Then										
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [ Fn_SISW_LifeView_2DMarkupPreferencesOperations ] Successfully performed [ " & sAction & " ] Action on [ "+sTabName+" ].")									
									End If									
								End If
							Case "FontName_Ext"
									If DictItems(iCounter) <> "" Then
									ObjDialog.WinEdit("Fontname").Set ""
									wait 1
									ObjDialog.WinEdit("Fontname").Type Left(DictItems(iCounter),3)
									For iCount = 1 To 10
									   wait 2
										Set WshShell = CreateObject("WScript.Shell")
										WshShell.SendKeys "{DOWN}"
										sValue = ObjDialog.WinEdit("Fontname").GetROProperty("text")
										If sValue = DictItems(iCounter) Then										
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [ Fn_SISW_LifeView_2DMarkupPreferencesOperations ] Successfully performed [ " & sAction & " ] Action on [ "+sTabName+" ].")									
											Fn_SISW_LifeView_2DMarkupPreferencesOperations=True
											Exit for
										End If	
									Next
									End If									
							Case Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_LifeView_2DMarkupPreferencesOperations ] Invalid case [ " & sAction & " ].")
								Exit function
						End Select
					Next
			
	'------------------------------------------------------------------------------------------------------------------------------------------------------------------			
				Case Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_LifeView_2DMarkupPreferencesOperations ] Invalid case [ " & sAction & " ].")
					Exit function
			End Select
	'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		Case "Line"
			'Implement as required
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		Case else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_LifeView_2DMarkupPreferencesOperations ] Invalid case [ " & sAction & " ].")
			Exit function
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	End Select
	
	If sButton <> "" Then
		If ObjDialog.WinButton(sButton).Exist(2) = True Then
			ObjDialog.WinButton(sButton).Click	
			If Err.Number < 0 Then
				 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to click on ["+sButton+"] Button.")
				 Set ObjDialog = nothing
				 Exit Function
			End If
		End if
	End If
	
	If sValue = True Then
		Fn_SISW_LifeView_2DMarkupPreferencesOperations = True
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [ Fn_SISW_LifeView_2DMarkupPreferencesOperations ] executed successfuly with case [ " & sAction & " ].")
	End If

	Set ObjDialog = nothing
End Function 

''*********************************************************		Function to Perform operation on Export Dialog in LifeCycle Viewer	***********************************************************************
'
''Function Name		:		Fn_SISW_LifeView_AutoFileSearchPreferences
'
''Description		:		 Function to perform operation on AutoFile search Preferences
'
''Parameters		:					1.	sAction = Action To Perform
'										2. sTab = Tab opened
'										3. dicInfo = Dictionary Object
'										4. sButton = Name of button to click. 		eg. "OK" or "cancel" 
'										5. sReserve = For Future Use  				

'			  										
''Return Value		: 				True or False 
'
''Pre-requisite		:				LCV Prespective.
'
'
''Examples			:		 			
'											Dim dicDocumentSearchOrderInfo
'											Set dicDocumentSearchOrderInfo=CreateObject("Scripting.Dictionary")
'											dicDocumentSearchOrderInfo("MoveValue") = "Original File Directory Set"
'											dicDocumentSearchOrderInfo("Column") = "Directory Set"
'											dicDocumentSearchOrderInfo("FromValue") = "Relative File Directory Set"										
'											bReturn= Fn_SISW_LifeView_AutoFileSearchPreferences("MoveUp","Document Search Order", dicDocumentSearchOrderInfo , "OK", "")
'
'History:
'										Developer Name				Date						Rev. No.							Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Reema Wadhwa					12-Nov-2014					1.0														Paresh 
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_LifeView_AutoFileSearchPreferences(sAction, sTab, dicDocumentSearchOrderInfo , sButton, sReserve)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_LifeView_AutoFileSearchPreferences"
	Dim  iIndexValue , iIndexFromValue, bReturn
	Dim ObjAutoFileSearchPref
	Fn_SISW_LifeView_AutoFileSearchPreferences=False

	Set ObjAutoFileSearchPref = Window("LifeViewWin").Dialog("AutoFileSearchPreferences")
	If NOT ObjAutoFileSearchPref.exist(4) Then
		call Fn_MenuOperation("WinMenuSelect", "File:Preferences:File Locate...")
	End if
	
	If NOT ObjAutoFileSearchPref.exist(4) Then
		Fn_SISW_LifeView_AutoFileSearchPreferences=False
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_LifeView_AutoFileSearchPreferences ]  AutoFileSearchPreferences Dialog is not opened.")
		ObjAutoFileSearchPref = nothing
		Exit Function	
	End if
	
	If sTab <> "" Then
		call Window("LifeViewWin").Dialog("AutoFileSearchPreferences").WinTab("AutoFilePreferencesTab").Select(sTab)
	End If
	wait 3
	
	
	Select Case sAction
		Case "MoveUp"
			If dicDocumentSearchOrderInfo("MoveValue") <> "" and  dicDocumentSearchOrderInfo("Column") <> "" and dicDocumentSearchOrderInfo("FromValue") <> "" Then
				bReturn = Fn_SISW_UI_WinListView_Operations("Fn_DocumentSearchOrder", "Select", ObjAutoFileSearchPref ,"DocSearchOrderList",dicDocumentSearchOrderInfo("MoveValue"), dicDocumentSearchOrderInfo("Column"))
				If bReturn = True Then
					iIndexValue = Fn_SISW_UI_WinListView_Operations("Fn_DocumentSearchOrder", "GetRowIndex", ObjAutoFileSearchPref ,"DocSearchOrderList",dicDocumentSearchOrderInfo("MoveValue"), dicDocumentSearchOrderInfo("Column"))
					iIndexFromValue = Fn_SISW_UI_WinListView_Operations("Fn_DocumentSearchOrder", "GetRowIndex", ObjAutoFileSearchPref ,"DocSearchOrderList",dicDocumentSearchOrderInfo("FromValue"), dicDocumentSearchOrderInfo("Column"))
					If iIndexValue <> False and iIndexFromValue <> False Then
						Do Until iIndexValue < iIndexFromValue
							ObjAutoFileSearchPref.WinButton("MoveUp").Click
							iIndexValue = Fn_SISW_UI_WinListView_Operations("Fn_DocumentSearchOrder", "GetRowIndex", ObjAutoFileSearchPref ,"DocSearchOrderList",dicDocumentSearchOrderInfo("MoveValue"), dicDocumentSearchOrderInfo("Column"))
							iIndexFromValue = Fn_SISW_UI_WinListView_Operations("Fn_DocumentSearchOrder", "GetRowIndex", ObjAutoFileSearchPref ,"DocSearchOrderList",dicDocumentSearchOrderInfo("FromValue"), dicDocumentSearchOrderInfo("Column"))							
						Loop
						If iIndexValue < iIndexFromValue Then
							Fn_SISW_LifeView_AutoFileSearchPreferences = True	
						End If
					End If
				End If
			End If
		Case "VerifyMoveUp"
			If dicDocumentSearchOrderInfo("MoveValue") <> "" and  dicDocumentSearchOrderInfo("Column") <> "" and dicDocumentSearchOrderInfo("FromValue") <> "" Then
				
					iIndexValue = Fn_SISW_UI_WinListView_Operations("Fn_DocumentSearchOrder", "GetRowIndex", ObjAutoFileSearchPref ,"DocSearchOrderList",dicDocumentSearchOrderInfo("MoveValue"), dicDocumentSearchOrderInfo("Column"))
					iIndexFromValue = Fn_SISW_UI_WinListView_Operations("Fn_DocumentSearchOrder", "GetRowIndex", ObjAutoFileSearchPref ,"DocSearchOrderList",dicDocumentSearchOrderInfo("FromValue"), dicDocumentSearchOrderInfo("Column"))
					If iIndexValue < iIndexFromValue and iIndexValue <> False and iIndexFromValue <> False Then
						Fn_SISW_LifeView_AutoFileSearchPreferences = True	
					End If
			End If
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_LifeView_AutoFileSearchPreferences ] Invalid case [ " & sAction & " ].")
			Exit function
	End Select	
	
	If sButton <> "" Then
		If ObjAutoFileSearchPref.WinButton(sButton).Exist(2) Then
			ObjAutoFileSearchPref.WinButton(sButton).Click	
		else
			Set ObjAutoFileSearchPref = nothing
			Exit function
		End if
	End If
	

	If Fn_SISW_LifeView_AutoFileSearchPreferences <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [ Fn_SISW_LifeView_AutoFileSearchPreferences ] executed successfuly with case [ " & sAction & " ].")
	else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_LifeView_AutoFileSearchPreferences ] Failed to executed with case [ " & sAction & " ].")
	End If

	Set ObjAutoFileSearchPref = nothing
End Function


''*********************************************************		Function to Perform operation on Section 3d  Preferences Dialog ***********************************************************************
'
''Function Name		:		Fn_SISW_LifeView_Section3DPreferencesOperations()
'
''Description		:		 Function to Perform operation on 3D Section Preferences Dialog
'
''Parameters		:					1. sCalledFrom=LCV or VIZ
'										2. sAction = Action To Perform
'							   			3. sTabName = pass the value of tab you want to select .  eg "Text"
'										4. sButton = Name of button to click. 		eg. "OK" or "cancel" 		
'										5. dicInfo = Dictionary Object
'										6. sReserve = For Future Use  		
'			  										
''Return Value		: 				True or False 
'
''Pre-requisite		:				Section3D Preferences Dialog should be displayed in LCV.
'
'
''Examples			:		 			Case "Viewer"		 'Case handled according to Tab of the Dialog
'											Dim dicInfo
'											Set dicInfo=CreateObject("Scripting.Dictionary")
											'dicInfo("DisplayActiveSection") = "ON"							
											'Fn_SISW_LifeView_Section3DPreferencesOperations("TC","Set","Viewer",dicInfo, "OK/Cancel", "")
'										Case "Grid"		 'Case handled according to Tab of the Dialog
'											Dim dicInfo
'											Set dicInfo=CreateObject("Scripting.Dictionary")
											'dicInfo("ShowGrid") = "ON"		
'											'dicInfo("ShowGridLabels") = "ON"												
											'Fn_SISW_LifeView_Section3DPreferencesOperations("TC","Set","Viewer",dicInfo, "OK/Cancel", "")
'History:
'										Developer Name				Date						Rev. No.							Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Reema W					17-Nov-2014					1.0									
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_LifeView_Section3DPreferencesOperations(sCalledFrom,sAction, sTabName , dicInfo , sButton, sReserve)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_LifeView_Section3DPreferencesOperations"
	Dim DictItems,DictKeys, bReturn
	Dim iCounter,sValue,objDialog
	Dim sMenuXMLPath, sToolsMarkupMenu
	Err.clear
	Fn_SISW_LifeView_Section3DPreferencesOperations=False

   	Select Case sCalledFrom
		Case "TC"  'for lcv
			Set ObjDialog =Window("LifeViewWin").Dialog("CrossSectionPreferences")
			
		Case "VIZ"
		
	End Select

	If ObjDialog.Exist(2) <> True Then
		Select Case sCalledFrom
				Case "TC"		'for lcv
					'Find File Path for Lifecycle Viewer Menu XML
					 sMenuXMLPath=Fn_LogUtil_GetXMLPath("LifecycleViewer_Menu")
					 
					 'Extract Menu Paths from XML
					 sToolsMarkupMenu = Fn_GetXMLNodeValue(sMenuXMLPath, "ToolsSection3D")
				
					'Check if the Tools:Section3D menu is checked or Not
					bReturn =Fn_MenuOperation("WinMenuCheck", sToolsMarkupMenu )
					If bReturn = False Then
						'Select Tools:Markup menu 
						bReturn =  Fn_MenuOperation("WinMenuSelect", sToolsMarkupMenu )
						If bReturn = False Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_LifeView_Section3DPreferencesOperations ] to Perform [ "+sToolsMarkupMenu+" ]")
							Set ObjDialog = Nothing
							Exit Function
						End If
					End If
					Wait 2
					sToolsMarkupMenu = Fn_GetXMLNodeValue(sMenuXMLPath, "Section3DPreferences")
					bReturn =  Fn_MenuOperation("WinMenuSelect", sToolsMarkupMenu )
					If bReturn = False Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_LifeView_Section3DPreferencesOperations ] to Perform [ "+sToolsMarkupMenu+" ]")
						Set ObjDialog = Nothing
						Exit Function
					End If
				Case "VIZ"
		End Select
	End If
	
	If ObjDialog.Exist(2) <> True Then
		Set ObjDialog=nothing
		Exit Function
	Else
		ObjDialog.WinTab("CrossSecTab").Select sTabName
	End If

   	Select Case sTabName
		Case "Viewer"
			Select Case sAction			
			Case "Set"
				DictKeys = dicInfo.Keys
				DictItems = dicInfo.Items
				For iCounter = 0 to Ubound(DictKeys)							
					Select Case DictKeys(iCounter)
						Case "DisplayActiveSection"
							If DictItems(iCounter) <> "" Then
								ObjDialog.WinCheckBox(DictKeys(iCounter)).Set DictItems(iCounter)
								If Err.Number < 0 Then
									 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to set on ["+DictKeys(iCounter)+"] checkbox.")
									 Set ObjDialog = nothing
									 Exit Function
								End If
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [ Fn_SISW_LifeView_Section3DPreferencesOperations ] set [ " + DictKeys(iCounter) +" ] = [ " & DictItems(iCounter) & " ].")
							End If
							
						Case Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_LifeView_Section3DPreferencesOperations ] Invalid Dicinfo [ " & DictKeys(iCounter) & " ].")
							Exit function
					End Select
				Next
		
'------------------------------------------------------------------------------------------------------------------------------------------------------------------			
				Case Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_LifeView_Section3DPreferencesOperations ] Invalid sAction [ " & sAction & " ].")
					Exit function
			End Select
	'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		Case "Grid"
			Select Case sAction			
			Case "Set"
				DictKeys = dicInfo.Keys
				DictItems = dicInfo.Items
				For iCounter = 0 to Ubound(DictKeys)							
					Select Case DictKeys(iCounter)
						Case "ShowGrid" , "ShowGridLabels"
							If DictItems(iCounter) <> "" Then
								ObjDialog.WinCheckBox(DictKeys(iCounter)).Set DictItems(iCounter)
								If Err.Number < 0 Then
									 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to set on ["+DictKeys(iCounter)+"] checkbox.")
									 Set ObjDialog = nothing
									 Exit Function
								End If
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [ Fn_SISW_LifeView_Section3DPreferencesOperations ]  set [ " + DictKeys(iCounter) +" ] = [ " & DictItems(iCounter) & " ].")
							End If
							
						Case Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_LifeView_Section3DPreferencesOperations ] Invalid Dicinfo [ " & DictKeys(iCounter) & " ].")
							Exit function
					End Select
				Next
		
	'------------------------------------------------------------------------------------------------------------------------------------------------------------------			
				Case Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_LifeView_Section3DPreferencesOperations ] Invalid sAction [ " & sAction & " ].")
					Exit function
			End Select
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		Case "General"
			Select Case sAction			
			Case "Set"
				DictKeys = dicInfo.Keys
				DictItems = dicInfo.Items
				For iCounter = 0 to Ubound(DictKeys)							
					Select Case DictKeys(iCounter)
						Case "SectionLinesWidth"
							
							If DictItems(iCounter) <> "" Then
								ObjDialog.WinEdit(DictKeys(iCounter)).Set DictItems(iCounter)
								If Err.Number < 0 Then
									 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to set ["+DictKeys(iCounter)+"] EditBox.")
									Fn_SISW_LifeView_Section3DPreferencesOperations = False
									Set ObjDialog = nothing
									 Exit Function
								End If
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [ Fn_SISW_LifeView_Section3DPreferencesOperations ]  set [ " + DictKeys(iCounter) +" ] = [ " & DictItems(iCounter) & " ].")
							End If
							
						Case Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_LifeView_Section3DPreferencesOperations ] Invalid Dicinfo [ " & DictKeys(iCounter) & " ].")
							Exit function
					End Select
				Next
					
			Case Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_LifeView_Section3DPreferencesOperations ] Invalid sAction [ " & sAction & " ].")
					Exit function
			End Select
					
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		Case "Clipping"
			Select Case sAction			
			Case "Set"
				DictKeys = dicInfo.Keys
				DictItems = dicInfo.Items
				For iCounter = 0 to Ubound(DictKeys)							
					Select Case DictKeys(iCounter)
						Case "ShowLinesinPartColor"
							
							If DictItems(iCounter) <> "" Then
								ObjDialog.WinCheckBox(DictKeys(iCounter)).Set DictItems(iCounter)
								If Err.Number < 0 Then
									 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to set on ["+DictKeys(iCounter)+"] checkbox.")
									Fn_SISW_LifeView_Section3DPreferencesOperations = False
									Set ObjDialog = nothing
									 Exit Function
								End If
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [ Fn_SISW_LifeView_Section3DPreferencesOperations ]  set [ " + DictKeys(iCounter) +" ] = [ " & DictItems(iCounter) & " ].")
							End If
							
						Case Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_LifeView_Section3DPreferencesOperations ] Invalid Dicinfo [ " & DictKeys(iCounter) & " ].")
							Exit function
					End Select
				Next
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
			Case "Verify"
				DictKeys = dicInfo.Keys
				DictItems = dicInfo.Items
				For iCounter = 0 to Ubound(DictKeys)							
					Select Case DictKeys(iCounter)
						Case "ShowLinesinPartColor"
							
							If DictItems(iCounter) <> "" Then
								sValue = ObjDialog.WinCheckBox(DictKeys(iCounter)).GetROProperty("checked")
								If sValue <> DictItems(iCounter) Then
									 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Verify value on ["+DictKeys(iCounter)+"] checkbox.")
									Fn_SISW_LifeView_Section3DPreferencesOperations = False
									Set ObjDialog = nothing
									 Exit Function
								End If
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [ Fn_SISW_LifeView_Section3DPreferencesOperations ] verified [ " + DictKeys(iCounter) +" ] = [ " & DictItems(iCounter) & " ].")
							End If
							
						Case Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_LifeView_Section3DPreferencesOperations ] Invalid Dicinfo [ " & DictKeys(iCounter) & " ].")
							Exit function
					End Select
				Next
				
			Case Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_LifeView_Section3DPreferencesOperations ] Invalid sAction [ " & sAction & " ].")
				Exit function
			End Select
		Case else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_LifeView_Section3DPreferencesOperations ] Invalid sTabName [ " & sAction & " ].")
			Exit function
		End Select
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

	
	If sButton <> "" Then
		If ObjDialog.WinButton(sButton).Exist(2) = True Then
			ObjDialog.WinButton(sButton).Click	
			If Err.Number < 0 Then
				 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to click on ["+sButton+"] Button.")
				 Set ObjDialog = nothing
				 Exit Function
			End If
		End if
	End If
	Fn_SISW_LifeView_Section3DPreferencesOperations=True	

	Set ObjDialog = nothing
End Function 

''*********************************************************		Function to Reposition Section Pane	***********************************************************************
'
''Function Name		:		Fn_SISW_LifeView_Section3D_CreateRepositionSection()
'
''Description		:		 Function to Reposition Pane, Verify Reposition Pane
'
''Parameters		:					1. sCalledFrom=TC or VIZ
'										2. sAction = Action To Perform
'							   			3. sAxis = X/Y/Z axis for Position Pane
'										4. sValue = value for position pane	
'										5. sReserve = For Future Use  		
'			  										
''Return Value		: 				True or False 
'
'
''Examples			:		 					
											'Fn_SISW_LifeView_Section3D_CreateRepositionSection("TC","Create","Z","","true", "")	
											'Fn_SISW_LifeView_Section3D_CreateRepositionSection("TC","Reposition","Z","0.25", "", "")												
'History:
'										Developer Name				Date						Rev. No.							Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Reema W					17-Nov-2014					1.0									
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_LifeView_Section3D_CreateRepositionSection(sCalledFrom,sAction, sAxis, sValue, sClose , sReserve)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_LifeView_Section3D_CreateRepositionSection"
	Dim bReturn
	Dim ObjDialog,iCounter
	Dim sMenuXMLPath, sToolsMarkupMenu
	
	Err.clear
	Fn_SISW_LifeView_Section3D_CreateRepositionSection=False

   	Select Case sCalledFrom
		Case "TC"  'for lcv
			Set ObjDialog =Window("LifeViewWin").Dialog("PositionPlane")	
		Case "VIZ"
		
	End Select

	If ObjDialog.Exist(2) <> True Then
		Select Case sCalledFrom
			Case "TC"		'for lcv
			'Find File Path for Lifecycle Viewer Menu XML
			 sMenuXMLPath=Fn_LogUtil_GetXMLPath("LifecycleViewer_Menu")
			 'Extract Menu Paths from XML
			 sToolsMarkupMenu = Fn_GetXMLNodeValue(sMenuXMLPath, "ToolsSection3D")
		
			'Check if the Tools:Section3D menu is checked or Not
			bReturn =Fn_MenuOperation("WinMenuCheck", sToolsMarkupMenu )
			If bReturn = False Then
				'Select Tools:Section3D menu 
				bReturn =  Fn_MenuOperation("WinMenuSelect", sToolsMarkupMenu )
				If bReturn = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_LifeView_Section3D_CreateRepositionSection ] to Perform [ "+sToolsMarkupMenu+" ]")
					Set ObjDialog = Nothing
					Exit Function
				End If
			End If
			Wait 2
			
			Select Case sAction
				Case "Create"
					sToolsMarkupMenu = Fn_GetXMLNodeValue(sMenuXMLPath, "Section3DCreateSection" + sAxis)
					'Select Section3D:create section:x/y/z menu 
					bReturn =  Fn_MenuOperation("WinMenuSelect", sToolsMarkupMenu )
					If bReturn = False Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_LifeView_Section3D_CreateRepositionSection ] to Perform [ "+sToolsMarkupMenu+" ]")
						Set ObjDialog = Nothing
						Exit Function
					End If
	
				Case "Reposition"
					sToolsMarkupMenu = Fn_GetXMLNodeValue(sMenuXMLPath, "Section3DPositionPlane")
					'Select Section3D:Position Pane menu 
					bReturn =  Fn_MenuOperation("WinMenuSelect", sToolsMarkupMenu )
					If bReturn = False Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_LifeView_Section3D_CreateRepositionSection ] to Perform [ "+sToolsMarkupMenu+" ]")
						Set ObjDialog = Nothing
						Exit Function
					End If
					
				Case Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_LifeView_Section3D_CreateRepositionSection ] Invalid sAction [ " & sAction & " ].")
					Set ObjDialog = Nothing
					Exit function
			End Select	
		
		Case "VIZ"
		'-------------
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_LifeView_Section3D_CreateRepositionSection ] Invalid Value [ " & sCalledFrom & " ].")
			Set ObjDialog = Nothing
			Exit function
		End Select
	End If
	If ObjDialog.Exist(2) <> True Then
		Set ObjDialog=nothing
		Exit Function
	End If
	
	Select Case sAction
	
		Case "Reposition", "Create"
			If sAxis<> "" Then
				ObjDialog.WinComboBox("Axis").select sAxis
			End If
			wait 1
			If sValue<> "" Then
				ObjDialog.WinEdit("Value").Set sValue
			End If
			wait 1
			ObjDialog.WinButton("Apply").Click
			If Err.Number < 0 Then
				 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to click on [ Apply ] Button.")
				 ObjDialog.Close
				 Set ObjDialog = nothing
				 Exit Function
			End If
					
			If Window("LifeViewWin").Dialog("Warning").Exist(2) = True Then
				Window("LifeViewWin").Dialog("Warning").WinButton("OK").Click
				ObjDialog.Close
				Set ObjDialog = nothing
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_LifeView_Section3D_CreateRepositionSection ] Successfully Repositioned Pane At Axis[" + aAxis + "] and Value [" + sValue + "]")
				Exit function
			End If
		
	End Select
	
	If lCase(sClose) = "true"  Then
		ObjDialog.Close
	End If
	
	Fn_SISW_LifeView_Section3D_CreateRepositionSection = True


	Set ObjDialog = nothing
End Function 


'****************************************    Function to handle Load Option Preferences dialog ***************************************
'
''Function Name		 	:	Fn_SISW_LifeView_3DLoaderPrefOperation()
'
''Description		    :  	Function to do operations on 3D Loader Option Preferences dialog.

''Parameters		    :	1. sAction : Action to be performed
'							2. sTab : Tab name - Others
'							3. Dim dicLoadOptPreference
'							4. For further use	
'							dicLoadOptPreference("Load geometry lazily ( Recommended )") = "OFF"
'							dicLoadOptPreference("SetDefaultLayer") = "ON"
'							
								
''Return Value		    :  	True \ False
'
''Examples		     	:	Fn_SISW_LifeView_3DLoaderPrefOperation("TC","Set","Other",dicLoadOptPreference,"OK")
'
'History:
'	Developer Name			Date			Rev. No.		Reviewer		Changes Done	
'----------------------------------------------------------------------------------------------------------------------------------
'	Shweta Rathod		20-Nov-2014		1.0				
'-----------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_LifeView_3DLoaderPrefOperation(sCalledFrom, sAction, sTab, dicLoadOptPreference, sButton, sReserve)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_LifeView_3DLoaderPrefOperation"
	'Declaring Variables
	Dim objLoadOpt,aPrefvalue,icnt,aPreferences,DictKeys,DictItems
	Fn_SISW_LifeView_3DLoaderPrefOperation = False
	
	'Set load option dialog object
	
	Select Case sCalledFrom
		Case "TC"
			Set objLoadOpt = Window("LifeViewWin").Dialog("3DLoaderPreferences")
		Case "VIZ"	
	End Select
	
	'if not exists then perform menu call
	If Fn_SISW_UI_Object_Operations("Fn_SISW_LifeView_LoadOptionPrefOperation","Exist", objLoadOpt,"") = False Then
		Call Fn_MenuOperation("WinMenuSelect","File:Preferences:3D Loader...")
		Call Fn_ReadyStatusSync(2)	
	End If
	
	If objLoadOpt.Exist(10) = True Then
		If sTab <> "" Then
			objLoadOpt.WinTab("OtherTab").Select sTab
			If err.number < 0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Set ["+ sTab +"] Tab")
				Set objLoadOpt = Nothing
				Exit Function
			End If	
		End If
		Select Case sTab
			Case "Other"
				Select Case sAction
					Case "Set"
						DictKeys = dicLoadOptPreference.Keys
						DictItems = dicLoadOptPreference.Items
						For iCounter = 0 to Ubound(DictKeys)							
							If DictItems(iCounter) <> "" Then
								objLoadOpt.WinCheckBox("GeneralChkBox").SetTOProperty "text",DictKeys(iCounter)
								objLoadOpt.WinCheckBox("GeneralChkBox").Set DictItems(iCounter)
								If Err.Number < 0 Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to set on ["+DictKeys(iCounter)+"] checkbox.")
									Fn_SISW_LifeView_Section3DPreferencesOperations = False
									Set objLoadOpt = nothing
									Exit Function
								End If
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [ Fn_SISW_LifeView_3DLoaderPrefOperation ]  set [ " + DictKeys(iCounter) +" ] = [ " & DictItems(iCounter) & " ].")
							End If
						Next
					Case Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_LifeView_3DLoaderPrefOperation ] Invalid Dicinfo [ " & DictKeys(iCounter) & " ].")
						Exit function
				End Select
		End Select
		'click on given button
		If sButton<>"" Then
			Call Fn_UI_WinButton_Click("Fn_SISW_LifeView_3DLoaderPrefOperation",objLoadOpt,sButton,5,5,micLeftBtn)
		End If	
	else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_LifeView_3DLoaderPrefOperation ] failed to open dialog .")
		Exit function	
	End if
	If Err.Number < 0 Then
		 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to click on [ Add ] Button.")
		 Set ObjDialog = nothing
		 Exit Function
	End If
	Set objLoadOpt = nothing
	Fn_SISW_LifeView_3DLoaderPrefOperation = True
End Function

'*********************************************************		Function to Perform operation on Export Dialog in LifeCycle Viewer	***********************************************************************
'
''Function Name		:		Fn_SISW_LifeView_FilterManager_Operations
'
''Description		:		Function to Perform operation on Export Dialog in LifeCycle Viewer
'
''Parameters		:					1.	sAction = Action To Perform
'										2. dicInfo = Dictionary Object
'										3. sButton = Name of button to click. 		eg. "OK" or "cancel" 
'										4. sReserve = For Future Use  		NOTE : For Enter Name Dialog that appears after Export Image Dialog use 'sReserve'		

'			  										
''Return Value		: 				True or False 
'
''Pre-requisite		:				Export Dialog should be displayed in LCV.
'
'
''Examples			:		 			
'											Dim dicInfo
'											Set dicInfo=CreateObject("Scripting.Dictionary")
'											dicInfo("File") = "C:\Block_8567+".plmxml"
'											dicInfo("FileFormat") = "Product Structure (*.plmxml)"
'															
'											bReturn = Fn_SISW_LifeView_FilterManager_Operations("TC","Export", dicInfo , "Close", "")
'
'History:
'										Developer Name				Date						Rev. No.							Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Shweta Rathod				12-Nov-2014					1.0									
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_LifeView_FilterManager_Operations(sCalledFrom,sAction, dicInfo , sButton, sReserve)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_LifeView_FilterManager_Operations"
	Dim DictItems,DictKeys,iCount
	Dim iCounter,sValue,objDialogEnterName
	Dim ObjExporterDialog,ObjImportDialog,objDialog
	Fn_SISW_LifeView_FilterManager_Operations=False
	iCount = 0	
	'Select the application, TC or Standalone TcViz
   	Select Case sCalledFrom
   		Case "TC"
			Set objDialog = Fn_SISW_LifeView_GetObject("FilterManager")
			Set ObjExporterDialog = Fn_SISW_LifeView_GetObject("FilterExport")
	
   		Case "VIZ"
			'------ Called From Standalone TcVIZ ---------------------------------

  	 End Select       '///End of Select Statement   
   
	If NOT ObjDialog.exist(4) Then
		call Fn_MenuOperation("WinMenuSelect", "Action:Filters...")
	End if
	wait 3
	If NOT ObjDialog.exist(4) Then
		Fn_SISW_LifeView_FilterManager_Operations=False
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_LifeView_ExportDialogOperation ] Export Dialog is not opened.")
		ObjDialog=nothing
		Exit Function	
	End if
	If dicInfo("AvailableFilter") <> "" Then
		objDialog.WinComboBox("AvailableFilters").SetTOProperty "attached text","Available Filters"
		If objDialog.WinComboBox("AvailableFilters").Exist(1) Then
			objDialog.WinComboBox("AvailableFilters").Select dicInfo("AvailableFilter")
			objDialog.WinButton("CommonButton").SetTOProperty "text","Add"
			objDialog.WinButton("CommonButton").Click
			If Err.Number < 0 Then
				 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to click on [ Add ] Button.")
				 Set ObjDialog = nothing
				 Exit Function
			End If
		Else
			Set ObjDialog = nothing
			Exit Function
		End If
	End If
	
	If dicInfo("FilterInputSel") <> "" Then
		objDialog.WinComboBox("AvailableFilters").SetTOProperty "attached text","Filter Input Select:"
		If objDialog.WinComboBox("AvailableFilters").Exist(1) Then
			objDialog.WinComboBox("AvailableFilters").Select dicInfo("FilterInputSel")
		Else
			Set ObjDialog = nothing
			Exit Function
		End If
	End If
	
	If dicInfo("FilterOutputSel") <> "" Then
		objDialog.WinComboBox("AvailableFilters").SetTOProperty "attached text","Filter Output Action:"
		If objDialog.WinComboBox("AvailableFilters").Exist(1) Then
			objDialog.WinComboBox("AvailableFilters").Select dicInfo("FilterOutputSel")
		Else
			Set ObjDialog = nothing
			Exit Function
		End If
	End If
	
	If dicInfo("SaveFilter") <> "" Then
		objDialog.WinButton("CommonButton").SetTOProperty "text","Save"
		objDialog.WinButton("CommonButton").Click
		If Window("LifeViewWin").Dialog("RenameFilter").Exist(10) Then
			Window("LifeViewWin").Dialog("RenameFilter").WinEdit("Name").Set dicInfo("SaveFilter")
			Window("LifeViewWin").Dialog("RenameFilter").WinButton("OK").Click	
			If Err.Number < 0 Then
				 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to click on [ OK ] Button.")
				 Set ObjDialog = nothing
				 Exit Function
			End If

			Window("LifeViewWin").Dialog("Warning").SetTOProperty "Text","FilterManager"
			If Window("LifeViewWin").Dialog("Warning").Exist(2) then
				Window("LifeViewWin").Dialog("Warning").WinButton("OK").Click		
			End If
		Else
			Exit Function
		End If
		
	End If
	
	DictKeys = dicInfo.Keys
	DictItems = dicInfo.Items
	Select Case sAction
		Case "Export"
			Window("LifeViewWin").Dialog("ExportFilters").SetTOProperty "text","Export Filters"
			objDialog.WinButton("CommonButton").SetTOProperty "text","Export..."
			objDialog.WinButton("CommonButton").Click
			For iCounter = 0 to Ubound(DictKeys)							
				Select Case DictKeys(iCounter)
					Case "File"
						If DictItems(iCounter) <> "" Then
							If ObjExporterDialog.WinEdit(DictKeys(iCounter)).Exist(1) Then
								ObjExporterDialog.WinEdit(DictKeys(iCounter)).Set DictItems(iCounter)	
							Else
								Exit Function
							End If
						End If
						
					Case "FileFormat"
						If DictItems(iCounter) <> "" Then
							If ObjExporterDialog.WinComboBox(DictKeys(iCounter)).Exist(1) Then
								ObjExporterDialog.WinComboBox(DictKeys(iCounter)).Select DictItems(iCounter)
								wait 2
							Else
								Exit Function
							End If
						End If
				End Select	
			Next
			
				ObjExporterDialog.WinButton("Cancel").SetTOProperty "text","&Save"
				ObjExporterDialog.WinButton("Cancel").Click
			
		'------------------------------------------import
		Case "Import"
			set ObjExporterDialog = Window("LifeViewWin").Dialog("ExportFilters")
			ObjExporterDialog.SetTOProperty "text","Import Filters"
			objDialog.WinButton("CommonButton").SetTOProperty "text","Import..."
			objDialog.WinButton("CommonButton").Click
			For iCounter = 0 to Ubound(DictKeys)							
				Select Case DictKeys(iCounter)
					Case "File"
						If DictItems(iCounter) <> "" Then
							If ObjExporterDialog.WinEdit(DictKeys(iCounter)).Exist(1) Then
								ObjExporterDialog.WinEdit(DictKeys(iCounter)).Set DictItems(iCounter)	
							Else
								Exit Function
							End If
						End If
						
					Case "FileFormat"
						If DictItems(iCounter) <> "" Then
							ObjExporterDialog.WinComboBox("FileFormat").SetTOProperty "attached text","Files of &type:"
							If ObjExporterDialog.WinComboBox(DictKeys(iCounter)).Exist(1) Then
								ObjExporterDialog.WinComboBox(DictKeys(iCounter)).Select DictItems(iCounter)
								wait 2
							Else
								Exit Function
							End If
						End If
				End Select	
			Next
			ObjExporterDialog.WinButton("Cancel").SetTOProperty "text","&Open"
			ObjExporterDialog.WinButton("Cancel").Click
		
			
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_LifeView_ExportDialogOperation ] Invalid case [ " & sAction & " ].")
			Exit function
	End Select
	
	Window("LifeViewWin").Dialog("Warning").SetTOProperty "Text","Warning"	
	If Window("LifeViewWin").Dialog("Warning").Exist(2) then
		Window("LifeViewWin").Dialog("Warning").WinButton("OK").Click		
	End If	
	
	If sButton <> "" Then
		objDialog.WinButton("CommonButton").SetTOProperty "text",sButton
		If ObjDialog.WinButton("CommonButton").Exist(2) Then
			ObjDialog.WinButton("CommonButton").Click	
		else
			Set ObjDialog = nothing
			Exit function
		End if
	End If
	If ObjDialog.exist(4) Then
		ObjDialog.WinButton("Cancel").Click	
	End if
	If Err.Number < 0 Then
		 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to click on [ OK ] Button.")
		 Set ObjDialog = nothing
		 Exit Function
	End If

	Fn_SISW_LifeView_FilterManager_Operations=True

	Set ObjDialog = nothing
	Set ObjExporterDialog = nothing
End Function
''*********************************************************		Function to Perform operation on Inspector Dialog in LifeCycle Viewer	***********************************************************************
'
''Function Name		:		Fn_SISW_LifeView_Inspector_Operations
'
''Description		:		 Function to Perform operation on Inspector Dialog in LifeCycle Viewer
'
''Parameters		:		1. sAction = Action To Perform
'							2. dicInfo = Dictionary Object
'							3. sButton = Name of button to click. 		eg. "OK" or "cancel" 	
'		  										
''Return Value		 : 	True or False 
'
''Pre-requisite		 :	Inspector Dialog should be displayed in LCV.
'
''Examples			 :		 			
'						 Set dicInfo=CreateObject("Scripting.Dictionary")
'						 dicInfo("EditBox1") = "Tolerance"
'						 dicInfo("Button") = "Browse..."
'						 bReturn= Fn_SISW_LifeView_Inspector_Operations("VerifyFieldExist", dicInfo , "OK")
'
'						 dicInfo("InsptreeNode") = "JT Fidelity;Cloud Of Point"
'						 bReturn = Fn_SISW_LifeView_Inspector_Operations("Select",dicInfo,"OK")
'
'History			:
'						Developer Name					Date		Rev. No.		Changes Done		Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'						Nishigandha J				02-Feb-2017		 1.0			Created				Poonam C										
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_LifeView_Inspector_Operations(sAction, dicInfo , sButton)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_LifeView_Inspector_Operations"
	Dim dictItems,dictKeys,iCount,dicCount,iCnt,sItems
	Dim iCounter,objDialog,aMetrics,sToolsMarkupMenu,sSubAction,sField

	Fn_SISW_LifeView_Inspector_Operations=False
	iCount = 0

	Set ObjDialog = Fn_SISW_LifeView_GetObject("VIZ_Inspector")
	
	If Fn_SISW_UI_Object_Operations("","Exist",ObjDialog,"") <> True Then
		 'Extract Menu Paths from XML
		 sToolsMarkupMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("Viz_Menu"), "ActionsInspector")

		'Select Action:Inspector... menu 
		bReturn =  Fn_SISW_LifeView_MenuOperation("WinMenuSelect",sToolsMarkupMenu)
		If bReturn = False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to Select Menu [ "+sToolsMarkupMenu+" ]")
			Set objDialog = Nothing
			Exit Function
		End If
		If Fn_SISW_UI_Object_Operations("","Exist",ObjDialog,"") <> True Then	
			Fn_SISW_LifeView_InspectorDialogOperations=False
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_LifeView_Inspector_Operations ] Inspector Dialog is not opened.")
			ObjDialog=nothing
			Exit Function	
		End If
	End If	
	
	Select Case sAction
	' - - - - -  Select Option from Inspector Tree -----------------
			Case "Select"
				If dicInfo("InsptreeNode") <> "" Then
						ObjDialog.WinTreeView("ConfigTree").Select dicInfo("InsptreeNode")
						If Err.Number < 0 Then
							Set objDialog = Nothing
							Fn_SISW_LifeView_Inspector_Operations = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to select option '"+dicInfo("InsptreeNode")+"'")
							Exit Function
						Else
							Fn_SISW_LifeView_Inspector_Operations = True
						End If
				Else
						Set objDialog = Nothing
						Fn_SISW_LifeView_Inspector_Operations = False
				End If
	' - - - - - - - - - - Verify Field from Inspector dialog -------------
			Case "VerifyFieldExist"
					dicCount = dicInfo.Count
					dicItems = dicInfo.Items
					dicKeys = dicInfo.Keys
					
					For iCounter = 0 To dicCount - 1
						If Instr(dicKeys(iCounter),"Button") > 0  Then
							sSubAction = "Button"
						ElseIf Instr(dicKeys(iCounter),"EditBox") > 0  Then 
							sSubAction = "EditBox"
						ElseIf Instr(dicKeys(iCounter),"ComboList") > 0  Then 
							sSubAction = "ComboList"	
						End if
						sField = dicItems(iCounter)
					
						Select Case sSubAction
							Case "Button"
									ObjDialog.WinButton("Button").SetTOProperty "text",sField
									bReturn = Fn_SISW_UI_Object_Operations("Fn_SISW_LifeView_Inspector_Operations","Exist",ObjDialog.WinButton("Button"),"")
									If bReturn=False Then
										Fn_SISW_LifeView_Inspector_Operations = False
										Set ObjDialog = Nothing
										Exit Function
									Else
										Fn_SISW_LifeView_Inspector_Operations = True
									End If
									
							Case "EditBox"
								If sField = "Cloud Of Points Data File" Then
									sField = ".*"
								End If
								ObjDialog.WinEdit("EditboxField").SetTOProperty "attached text",sField+":"	
								bReturn = Fn_SISW_UI_Object_Operations("Fn_SISW_LifeView_Inspector_Operations","Exist", ObjDialog.WinEdit("EditboxField"),"")
								If bReturn=False Then
									bFlag = False
								Else
									bFlag = True
								End If
								If bFlag = False Then
									Fn_SISW_LifeView_Inspector_Operations = False
									Set ObjDialog = Nothing
									Exit Function
								Else
									Fn_SISW_LifeView_Inspector_Operations = True
								End If
								Fn_SISW_LifeView_Inspector_Operations = True
								
							Case "ComboList"
								aMetrics = Split(sField,"~")
								sItems = ObjDialog.WinComboBox("Metric").GetROProperty("all items")
								
								For iCnt = 0 To UBound(aMetrics)
									If instr(1,sItems,aMetrics(iCnt)) Then
									 	Fn_SISW_LifeView_Inspector_Operations = True
									Else
										Fn_SISW_LifeView_Inspector_Operations = False
										Set ObjDialog = Nothing
										Exit Function
									End If
								Next

						     End Select
						Next

		Case Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_LifeView_ExportDialogOperation ] Invalid case [ " & sAction & " ].")
				Exit function
	End Select
	
	If sButton <> "" Then
		ObjDialog.WinButton(sButton).Click
		Call Fn_ReadyStatusSync(1)
	End If
	
	Set ObjDialog = nothing
	
End Function
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''''/$$$$
''''/$$$$  FUNCTION NAME   	:  Fn_SISW_LifeView_DisplayOptions_Operations
''''/$$$$
''''/$$$$  DESCRIPTION     	:  Function is used to perform operations on Display Options dialog
''''/$$$$ 
''''/$$$$
''''/$$$$  PARAMETERS   	:  sAction 		: 	Action to be performed
''''/$$$$ 					   dicDisplayOptions 	: 	Dictionary object
''''/$$$$ 					   sButton 		: 	Button to be clicked
''''/$$$$ 					   sReserve 		: 	For future use
''''/$$$$	
''''/$$$$	Return Value 	:  True or False
''''/$$$$										
''''/$$$$	How To Use 		:  Set dicDisplayOptions = CreateObject("Scripting.Dictionary")
''''/$$$$						dicDisplayOptions("WinCheckBox1") 		= "Color Display:OFF"
''''/$$$$						dicDisplayOptions("WinCheckBox2") 		= "Deformation Display:ON"
''''/$$$$						dicDisplayOptions("WinButton") 		= "ColorDisplayConfigure"
''''/$$$$						bReturn = Fn_SISW_LifeView_DisplayOptions_Operations("Set",dicDisplayOptions,"","")
''''/$$$$	
''''/$$$$					Developer Name			Date		Version		Changes				Reviewer
''''/$$$$	Created by  :	Priyanka Kakade	 	08/03/2017	  	1.0		   Created				Poonam C
''''/$$$$  	
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Public Function Fn_SISW_LifeView_DisplayOptions_Operations(sAction,dicDisplayOptions,sButton,sReserve)
	
	Dim dicCount, dicItems, dicKeys
	Dim bFlag, bFlag1, iCounter, sSubAction, sProperty,sMenu
	Dim aProperty
	Dim objDisplayOptions
	
	Fn_SISW_LifeView_DisplayOptions_Operations = False
	On Error Resume Next
	
	Set objDisplayOptions = Window("VizMainWin").Dialog("Display Options")
	
	If objDisplayOptions.Exist(5) = False Then
		sMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("Viz_Menu"), "CAEViewingDisplayOptions")
		'Select CAE Viewing:Display Options... menu 
		bReturn =  Fn_SISW_LifeView_MenuOperation("WinMenuSelect",sMenu)
		If bReturn = False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to Select Menu [ "+sMenu+" ]")
			Set objDisplayOptions = Nothing
			Exit Function
		End If
		If objDisplayOptions.Exist(5) = False Then
			Set objDisplayOptions = Nothing
			Exit Function
		End If
	End If
	
	Select Case sAction
		Case "Set"
			dicCount = dicDisplayOptions.Count
			dicItems = dicDisplayOptions.Items
			dicKeys = dicDisplayOptions.Keys
			
			For iCounter = 0 To dicCount - 1
				If Instr(dicKeys(iCounter),"WinCheckBox")>0 Then
					sSubAction = "WinCheckBox"
				ElseIf Instr(dicKeys(iCounter),"WinButton")>0 Then
					sSubAction = "WinButton"
				Else
					sSubAction = dicKeys(iCounter)
				End If
				sProperty = dicItems(iCounter)
				bFlag = False
				
				Select Case sSubAction
					Case "WinCheckBox"
					If sProperty<>"" Then
						aProperty = Split(sProperty,":")
						If objDisplayOptions.WinCheckBox(aProperty(0)).Exist Then
							objDisplayOptions.WinCheckBox(aProperty(0)).Set aProperty(1)
							If Err.Number < 0 Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Set checkbox ["+aProperty(0)+"] ["+aProperty(1)+"].")
								Set objDisplayOptions = Nothing
								Exit Function
							End If
							Wait 1
							bFlag = True
						End If
					End If
					Case "WinButton"
						If sProperty<>"" Then
							'Click on button provided
							If objDisplayOptions.WinButton(sProperty).Exist Then
								objDisplayOptions.WinButton(sProperty).Click 5,5,micLeftBtn
								If Err.Number < 0 Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to click ["+sProperty+"] button on [Display Options] dialog.")
									Set objDisplayOptions = Nothing
									Exit Function
								End If
								Wait 1
								bFlag = True
							End If
						End If
				End Select
				
				If bFlag = False Then
					Fn_SISW_LifeView_DisplayOptions_Operations = False
					Set objDisplayOptions = Nothing
					Exit Function
				Else
					Fn_SISW_LifeView_DisplayOptions_Operations = True
				End If		
			Next
	End Select
	
	'Click on button provided
	If sButton<>"" Then
		bFlag1 = Fn_SISW_UI_WinButton_Operations("Fn_SISW_LifeView_DisplayOptions_Operations", "Click", objDisplayOptions,sButton,"","","")
		If bFlag1 = False Then
			Fn_SISW_LifeView_DisplayOptions_Operations = False
			Set objDisplayOptions = Nothing
			Exit Function
		End If
	End If
	
	Set objDisplayOptions = Nothing
End Function
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''''/$$$$
''''/$$$$  FUNCTION NAME   	:  Fn_SISW_LifeView_ColorDisplay_Operations
''''/$$$$
''''/$$$$  DESCRIPTION     	:  Function is used to perform operations on Color Display dialog
''''/$$$$ 
''''/$$$$
''''/$$$$  PARAMETERS   	:  sAction 		: 	Action to be performed
''''/$$$$ 					   dicDisplayOptions 	: 	Dictionary object
''''/$$$$ 					   sButton 		: 	Button to be clicked
''''/$$$$ 					   sReserve 		: 	For future use
''''/$$$$	
''''/$$$$	Return Value 	:  True or False
''''/$$$$										
''''/$$$$	How To Use 		:  Set dicDisplayOptions = CreateObject("Scripting.Dictionary")
''''/$$$$						dicDisplayOptions("WinComboBox") 		= "Color Mode:Arrows"
''''/$$$$						dicDisplayOptions("WinEdit") 		= "Arrow Size:5.00"
''''/$$$$						dicDisplayOptions("WinButton") 		= "OK"
''''/$$$$						bReturn = Fn_SISW_LifeView_ColorDisplay_Operations("Set",dicDisplayOptions,"")
''''/$$$$	
''''/$$$$					Developer Name			Date		Version			Changes				Reviewer
''''/$$$$	Created by  :	Priyanka Kakade	 	08/03/2017	  	1.0		   		Created				Poonam C
''''/$$$$	
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Public Function Fn_SISW_LifeView_ColorDisplay_Operations(sAction,dicDisplayOptions,sReserve)
	
	Dim dicCount, dicItems, dicKeys
	Dim bFlag, bFlag1, iCounter, sSubAction, sProperty
	Dim aProperty
	Dim objDisplayOptions,objColorDisplay
	
	Fn_SISW_LifeView_ColorDisplay_Operations = False
	On Error Resume Next
	
	Set objDisplayOptions = Window("VizMainWin").Dialog("Display Options")
	Set objColorDisplay = Window("VizMainWin").Dialog("Color Display")
	
	If objColorDisplay.Exist(5) = False Then
		objDisplayOptions.WinButton("ColorDisplayConfigure").Click micLeftBtn
		If Err.Number < 0 Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to click [Configure] button on [Display Options] dialog.")
			Set objDisplayOptions = Nothing
			Set objColorDisplay = Nothing
			Exit Function
		End If
		If objColorDisplay.Exist(5) = False Then
			Set objDisplayOptions = Nothing
			Set objColorDisplay = Nothing
			Exit Function
		End If
	End If
	
	Select Case sAction
		Case "Set"
			dicCount = dicDisplayOptions.Count
			dicItems = dicDisplayOptions.Items
			dicKeys = dicDisplayOptions.Keys
			
			For iCounter = 0 To dicCount - 1
				If Instr(dicKeys(iCounter),"WinComboBox")>0 Then
					sSubAction = "WinComboBox"
				ElseIf Instr(dicKeys(iCounter),"WinEdit")>0 Then
					sSubAction = "WinEdit"
				ElseIf Instr(dicKeys(iCounter),"WinButton")>0 Then
					sSubAction = "WinButton"
				Else
					sSubAction = dicKeys(iCounter)
				End If
				sProperty = dicItems(iCounter)
				bFlag = False
				
				Select Case sSubAction
					Case "WinEdit"
						If sProperty<>"" Then
							aProperty = Split(sProperty,":")
							objColorDisplay.WinEdit("Arrow Size").SetTOProperty "attached text",aProperty(0)
							If objColorDisplay.WinEdit("Arrow Size").Exist Then
								objColorDisplay.WinEdit("Arrow Size").Set aProperty(1)
								If Err.Number < 0 Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Set ["+aProperty(1)+"] in ["+aProperty(0)+"] WinEdit Box.")
									Set objDisplayOptions = Nothing
									Set objColorDisplay = Nothing
									Exit Function
								End If
								Wait 1
								bFlag = True
							End If
						End If
					Case "WinComboBox"
						If sProperty<>"" Then
							aProperty = Split(sProperty,":")
							'Set text property 
							objColorDisplay.WinComboBox("Color Mode").SetTOProperty "attached text",aProperty(0)
							Wait 1
							If objColorDisplay.WinComboBox("Color Mode").Exist Then
								objColorDisplay.WinComboBox("Color Mode").Select aProperty(1)
								If Err.Number < 0 Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to select ["+aProperty(1)+"] from ["+aProperty(0)+"] WinCombo Box.")
									Set objDisplayOptions = Nothing
									Set objColorDisplay = Nothing
									Exit Function
								End If
								Wait 1
								bFlag = True
							End If
						End If
					Case "WinButton"
						If sProperty<>"" Then
							'Set text property 
							If objColorDisplay.WinButton(sProperty).Exist Then
								objColorDisplay.WinButton(sProperty).Click 5,5,micLeftBtn
								If Err.Number < 0 Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to click ["+sProperty+"] button on [Color Display] dialog.")
									Set objDisplayOptions = Nothing
									Set objColorDisplay = Nothing
									Exit Function
								End If
								Wait 1
								bFlag = True
							End If
						End If
				End Select
				
				If bFlag = False Then
					Fn_SISW_LifeView_ColorDisplay_Operations = False
					Set objDisplayOptions = Nothing
					Set objColorDisplay = Nothing
					Exit Function
				Else
					Fn_SISW_LifeView_ColorDisplay_Operations = True
				End If
			Next	
	End Select
	
	Set objDisplayOptions = Nothing
	Set objColorDisplay = Nothing
End Function
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''''/$$$$
''''/$$$$  FUNCTION NAME   	:  Fn_SISW_LifeView_DeformationDisplay_Operations
''''/$$$$
''''/$$$$  DESCRIPTION     	:  Function is used to perform operations on Deformation Display dialog
''''/$$$$ 
''''/$$$$
''''/$$$$  PARAMETERS   	:  sAction 		: 	Action to be performed
''''/$$$$ 					   dicDisplayOptions 	: 	Dictionary object
''''/$$$$ 					   sReserve 		: 	For future use
''''/$$$$	
''''/$$$$	Return Value 	:  True or False
''''/$$$$										
''''/$$$$	How To Use 		:  Set dicDisplayOptions = CreateObject("Scripting.Dictionary")
''''/$$$$						dicDisplayOptions("WinEdit") 		= "Scale:5.00"
''''/$$$$						dicDisplayOptions("WinButton") 		= "OK"
''''/$$$$						bReturn = Fn_SISW_LifeView_DeformationDisplay_Operations("Set",dicDisplayOptions,"")
''''/$$$$	
''''/$$$$					Developer Name			Date		Version			Changes				Reviewer
''''/$$$$	Created by  :	Priyanka Kakade	 	09/03/2017	  	1.0		   		Created				Poonam C
''''/$$$$	
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Public Function Fn_SISW_LifeView_DeformationDisplay_Operations(sAction,dicDisplayOptions,sReserve)
	
	Dim dicCount, dicItems, dicKeys
	Dim bFlag, bFlag1, iCounter, sSubAction, sProperty
	Dim aProperty
	Dim objDeformation
	
	Fn_SISW_LifeView_DeformationDisplay_Operations = False
	On Error Resume Next
	
	Set objDeformation = Window("VizMainWin").Dialog("Deformation Display")
	
	If objDeformation.Exist(5) = False Then
		Set objDeformation = Nothing
		Exit Function
	End If
	
	Select Case sAction
		Case "Set"
			dicCount = dicDisplayOptions.Count
			dicItems = dicDisplayOptions.Items
			dicKeys = dicDisplayOptions.Keys
			
			For iCounter = 0 To dicCount - 1
				If Instr(dicKeys(iCounter),"WinEdit")>0 Then
					sSubAction = "WinEdit"
				ElseIf Instr(dicKeys(iCounter),"WinButton")>0 Then
					sSubAction = "WinButton"
				ElseIf Instr(dicKeys(iCounter),"WinComboBox")>0 Then
					sSubAction = "WinComboBox"
				Else
					sSubAction = dicKeys(iCounter)
				End If
				sProperty = dicItems(iCounter)
				bFlag = False
				
				Select Case sSubAction
					Case "WinEdit"
						If sProperty<>"" Then
							aProperty = Split(sProperty,":")
							If objDeformation.WinEdit("Scale").Exist Then
								objDeformation.WinEdit("Scale").Set aProperty(1)
								If Err.Number < 0 Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Set ["+aProperty(1)+"] in ["+aProperty(0)+"] WinEdit Box.")
									Set objDeformation = Nothing
									Exit Function
								End If
								Wait 1
								bFlag = True
							End If
						End If
					Case "WinButton"
						If sProperty<>"" Then
							'Set text property 
							If objDeformation.WinButton(sProperty).Exist Then
								objDeformation.WinButton(sProperty).Click 5,5,micLeftBtn
								If Err.Number < 0 Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to click ["+sProperty+"] button on [Deformation Display] dialog.")
									Set objDeformation = Nothing
									Exit Function
								End If
								Wait 1
								bFlag = True
							End If
						End If
					Case "WinComboBox" '[TC11.5(20180122.00)_NewDevelopment_PoonamC_15Feb2018 : Added Case to select values for combo box]
						If sProperty<>"" Then
							aProperty = Split(sProperty,":")
							objDeformation.WinComboBox("ResultSelComboBox").SetTOProperty "attached text",aProperty(0)
							If objDeformation.WinComboBox("ResultSelComboBox").Exist Then
								objDeformation.WinComboBox("ResultSelComboBox").Select aProperty(1)
								If Err.Number < 0 Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to select ["+aProperty(1)+"] in ["+aProperty(0)+"] WinCombo Box.")
									Set objDeformation = Nothing
									Exit Function
								End If
								Wait 1
								bFlag = True
							End If
						End If	
				End Select
				
				If bFlag = False Then
					Fn_SISW_LifeView_DeformationDisplay_Operations = False
					Set objDeformation = Nothing
					Exit Function
				Else
					Fn_SISW_LifeView_DeformationDisplay_Operations = True
				End If
			Next	
	End Select
	
	Set objDeformation = Nothing
End Function
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''''/$$$$
''''/$$$$  FUNCTION NAME   	:  Fn_SISW_LifeView_LayerFilter_Operations
''''/$$$$
''''/$$$$  DESCRIPTION     	:  Function is used to perform operations on Layer Filter dialog
''''/$$$$ 
''''/$$$$
''''/$$$$  PARAMETERS   	:  sAction 		: 	Action to be performed
''''/$$$$ 					   dicDisplayOptions 	: 	Dictionary object
''''/$$$$ 					   sButton 		: 	Button to be clicked
''''/$$$$ 					   sReserve 		: 	For future use
''''/$$$$	
''''/$$$$	Return Value 	:  True or False
''''/$$$$										
''''/$$$$	How To Use 		:  Set dicDisplayOptions = CreateObject("Scripting.Dictionary")
''''/$$$$						dicDisplayOptions("WinButton") 	= "Find Layers"
''''/$$$$						dicDisplayOptions("WinComboBox") = "All"
''''/$$$$						dicDisplayOptions("WinList") 	= "[Default Filter]"
''''/$$$$						dicDisplayOptions("WinCheckBox") = "Use default filter when selected filter is not available:ON"
''''/$$$$						bReturn = Fn_SISW_LifeView_LayerFilter_Operations("Set",dicDisplayOptions,"Apply","")
''''/$$$$	
''''/$$$$					Developer Name			Date		Version		Changes				Reviewer
''''/$$$$	Created by  :	Priyanka Kakade	 	10/03/2017	  	1.0		   Created				Poonam C
''''/$$$$  	
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Public Function Fn_SISW_LifeView_LayerFilter_Operations(sAction,dicLayerFilterOptions,sButton,sReserve)
	
	Dim dicCount, dicItems, dicKeys
	Dim bFlag, bFlag1, iCounter, sSubAction, sProperty,sMenu
	Dim aProperty
	Dim objLayerFilter
	
	Fn_SISW_LifeView_LayerFilter_Operations = False
	On Error Resume Next
	
	Set objLayerFilter = Window("VizMainWin").Dialog("LayerFilter")
	
	If objLayerFilter.Exist(5) = False Then
		sMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("Viz_Menu"), "ActionsLayerFilter")
		
		bReturn =  Fn_SISW_LifeView_MenuOperation("WinMenuSelect",sMenu)
		If bReturn = False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to Select Menu [ "+sMenu+" ]")
			Set objLayerFilter = Nothing
			Exit Function
		End If
		If objLayerFilter.Exist(5) = False Then
			Set objLayerFilter = Nothing
			Exit Function
		End If
	End If
	
	Select Case sAction
		Case "Set"
			dicCount = dicLayerFilterOptions.Count
			dicItems = dicLayerFilterOptions.Items
			dicKeys = dicLayerFilterOptions.Keys
			
			For iCounter = 0 To dicCount - 1
				If Instr(dicKeys(iCounter),"WinCheckBox")>0 Then
					sSubAction = "WinCheckBox"
				ElseIf Instr(dicKeys(iCounter),"WinButton")>0 Then
					sSubAction = "WinButton"
				ElseIf Instr(dicKeys(iCounter),"WinComboBox")>0 Then
					sSubAction = "WinComboBox"
				ElseIf Instr(dicKeys(iCounter),"WinList")>0 Then
					sSubAction = "WinList"
				Else
					sSubAction = dicKeys(iCounter)
				End If
				sProperty = dicItems(iCounter)
				bFlag = False
				
				Select Case sSubAction
					Case "WinCheckBox"
						If sProperty<>"" Then
							aProperty = Split(sProperty,":")
							objLayerFilter.WinCheckBox("SelectAllComponents").SetTOProperty "text",aProperty(0)
							If objLayerFilter.WinCheckBox("SelectAllComponents").Exist Then
								objLayerFilter.WinCheckBox("SelectAllComponents").Set aProperty(1)
								If Err.Number < 0 Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Set checkbox ["+aProperty(0)+"] ["+aProperty(1)+"].")
									Set objLayerFilter = Nothing
									Exit Function
								End If
								Wait 1
								bFlag = True
							End If
						End If
					Case "WinButton"
						If sProperty<>"" Then
							'Click on button provided
							If objLayerFilter.WinButton(sProperty).Exist Then
								objLayerFilter.WinButton(sProperty).Click 5,5,micLeftBtn
								If Err.Number < 0 Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to click ["+sProperty+"] button on [Layer Filter] dialog.")
									Set objLayerFilter = Nothing
									Exit Function
								End If
								Wait 1
								bFlag = True
							End If
						End If
					Case "WinComboBox"
						If sProperty<>"" Then
							If objLayerFilter.WinComboBox("SearchScope").Exist Then
								objLayerFilter.WinComboBox("SearchScope").Select sProperty
								If Err.Number < 0 Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to select ["+sProperty+"] from ComboBox [SearchScope] on [ Layer Filter ]dialog.")
									Set objLayerFilter = Nothing
									Exit Function
								End If
								Wait 1
								bFlag = True
							End If
						End If
					Case "WinList"
						If sProperty<>"" Then
							If objLayerFilter.WinList("ListBox").Exist Then
								objLayerFilter.WinList("ListBox").Select sProperty
								If Err.Number < 0 Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to select ["+sProperty+"] from List on [ Layer Filter ]dialog.")
									Set objLayerFilter = Nothing
									Exit Function
								End If
								Wait 1
								bFlag = True
							End If
						End If
				End Select
				
				If bFlag = False Then
					Fn_SISW_LifeView_LayerFilter_Operations = False
					Set objLayerFilter = Nothing
					Exit Function
				Else
					Fn_SISW_LifeView_LayerFilter_Operations = True
				End If		
			Next
			
	End Select
	
	'Click on button provided
	If sButton<>"" Then
		bFlag1 = Fn_SISW_UI_WinButton_Operations("Fn_SISW_LifeView_LayerFilter_Operations", "Click", objLayerFilter,sButton,"","","")
		If bFlag1 = False Then
			Fn_SISW_LifeView_LayerFilter_Operations = False
			Set objLayerFilter = Nothing
			Exit Function
		End If
	End If
	
	Set objLayerFilter = Nothing
End Function
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''''/$$$$
''''/$$$$  FUNCTION NAME   	:  Fn_SISW_LifeView_PartEdit_Operations
''''/$$$$
''''/$$$$  DESCRIPTION     	:  Function is used to perform operations on Part Edit dialog
''''/$$$$ 
''''/$$$$
''''/$$$$  PARAMETERS   	:  sAction 		: 	Action to be performed
''''/$$$$ 					   dicDisplayOptions 	: 	Dictionary object
''''/$$$$ 					   sButton 		: 	Button to be clicked
''''/$$$$ 					   sReserve 		: 	For future use
''''/$$$$	
''''/$$$$	Return Value 	:  True or False
''''/$$$$										
''''/$$$$	How To Use 		:  Set dicPartEdit = CreateObject("Scripting.Dictionary")
''''/$$$$						dicPartEdit.RemoveAll
''''/$$$$						dicPartEdit("VerifyTabActivate") = "B-Rep"
''''/$$$$						dicPartEdit("WinObject1") = "Input"
''''/$$$$						dicPartEdit("WinObject2") 	= "FaceEditing"
''''/$$$$						bReturn = Fn_SISW_LifeView_PartEdit_Operations("VerifyFields",dicPartEdit,"","")
''''/$$$$	
''''/$$$$					Developer Name			Date		Version		Changes				Reviewer
''''/$$$$	Created by  :	Priyanka Kakade	 	15/03/2017	  	1.0		   Created				Poonam C
''''/$$$$  	
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Public Function Fn_SISW_LifeView_PartEdit_Operations(sAction,dicPartEdit,sButton,sReserve)
	
	Dim dicCount, dicItems, dicKeys
	Dim bFlag, bFlag1, iCounter,iIndex
	Dim sContent,sNode, sSubAction, sProperty,sMenu,sAppValue
	Dim objPartEdit
	
	Fn_SISW_LifeView_PartEdit_Operations = False
	On Error Resume Next
	
	Set objPartEdit = Window("VizMainWin").Dialog("PartEdit")
	
	If objPartEdit.Exist(5) = False Then
		Set objPartEdit = Nothing
		Exit Function
	End If
	
	Select Case sAction
		Case "Set"
			dicCount = dicPartEdit.Count
			dicItems = dicPartEdit.Items
			dicKeys = dicPartEdit.Keys
			
			For iCounter = 0 To dicCount - 1
				If Instr(dicKeys(iCounter),"WinTab")>0 Then
					sSubAction = "WinTab"
				ElseIf Instr(dicKeys(iCounter),"WinButton")>0 Then
					sSubAction = "WinButton"
				ElseIf Instr(dicKeys(iCounter),"WinComboBox")>0 Then
					sSubAction = "WinComboBox"
				ElseIf Instr(dicKeys(iCounter),"WinList")>0 Then
					sSubAction = "WinList"
				Else
					sSubAction = dicKeys(iCounter)
				End If
				sProperty = dicItems(iCounter)
				bFlag = False
				
				Select Case sSubAction
					Case "WinTab"
						If sProperty<>"" Then
							objPartEdit.WinTab("SysTabControl32").Select sProperty
							If Err.Number < 0 Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Select tab ["+sProperty+"] on [Part Edit] dialog")
								Set objPartEdit = Nothing
								Exit Function
							End If
							Wait 1
							bFlag = True
						End If
					Case "WinButton"
						If sProperty<>"" Then
							'Click on button provided
							If objPartEdit.WinButton(sProperty).Exist Then
								objPartEdit.WinButton(sProperty).Click 5,5,micLeftBtn
								If Err.Number < 0 Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to click ["+sProperty+"] button on [Part Edit] dialog.")
									Set objPartEdit = Nothing
									Exit Function
								End If
								Wait 1
								bFlag = True
							End If
						End If
					Case "WinComboBox"
						If sProperty<>"" Then
							If objPartEdit.WinComboBox("Parts").Exist Then
								objPartEdit.WinComboBox("Parts").Select sProperty
								If Err.Number < 0 Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to select ["+sProperty+"] from ComboBox [Parts] on [ Part Edit ]dialog.")
									Set objPartEdit = Nothing
									Exit Function
								End If
								Wait 1
								bFlag = True
							End If
						End If
				End Select
				
				If bFlag = False Then
					Fn_SISW_LifeView_PartEdit_Operations = False
					Set objPartEdit = Nothing
					Exit Function
				Else
					Fn_SISW_LifeView_PartEdit_Operations = True
				End If		
			Next
		
		Case "VerifyFields"	
			dicCount = dicPartEdit.Count
			dicItems = dicPartEdit.Items
			dicKeys = dicPartEdit.Keys
			
			For iCounter = 0 To dicCount - 1
				If Instr(dicKeys(iCounter),"VerifyTabActivate")>0 Then
					sSubAction = "VerifyTabActivate"
				ElseIf Instr(dicKeys(iCounter),"WinButton")>0 Then
					sSubAction = "WinButton"
				ElseIf Instr(dicKeys(iCounter),"WinComboBox")>0 Then
					sSubAction = "WinComboBox"
				ElseIf Instr(dicKeys(iCounter),"WinObject")>0 Then
					sSubAction = "WinObject"
				Else
					sSubAction = dicKeys(iCounter)
				End If
				sProperty = dicItems(iCounter)
				bFlag = False
				
				Select Case sSubAction
					Case "VerifyTabActivate"
						If sProperty<>"" Then
							sAppValue = objPartEdit.WinTab("SysTabControl32").GetSelection()
							If sAppValue <> sProperty Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to verify tab ["+sProperty+"] is selected on [Part Edit] dialog")
								Set objPartEdit = Nothing
								Exit Function
							End If
							Wait 1
							bFlag = True
						End If
					Case "WinButton"
						If sProperty<>"" Then
							If NOT objPartEdit.WinButton(sProperty).Exist Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to verify ["+sProperty+"] button exist on [Part Edit] dialog.")
								Set objPartEdit = Nothing
								Exit Function
							End If
							Wait 1
							bFlag = True
						End If
					Case "WinObject"
						If sProperty<>"" Then
							If NOT objPartEdit.WinObject(sProperty).Exist Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to verify ["+sProperty+"] exist on [Part Edit] dialog.")
								Set objPartEdit = Nothing
								Exit Function
							End If
							Wait 1
							bFlag = True
						End If
					Case "WinComboBox"
						If sProperty<>"" Then
							If objPartEdit.WinComboBox("Parts").Exist Then
								sAppValue = objPartEdit.WinComboBox("Parts").GetROProperty("text")
								If sAppValue <> sProperty Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to verify value ["+sProperty+"] is set in ComboBox [Parts] on [ Part Edit ]dialog.")
									Set objPartEdit = Nothing
									Exit Function
								End If
								Wait 1
								bFlag = True
							End If
						End If
				End Select
				
				If bFlag = False Then
					Fn_SISW_LifeView_PartEdit_Operations = False
					Set objPartEdit = Nothing
					Exit Function
				Else
					Fn_SISW_LifeView_PartEdit_Operations = True
				End If		
			Next
	
	Case "ControlPanelOperations"	
				dicCount = dicPartEdit.Count
				dicItems = dicPartEdit.Items
				dicKeys = dicPartEdit.Keys
				
				For iCounter = 0 To dicCount - 1
					If Instr(dicKeys(iCounter),"ExpandNode")>0 Then
						sSubAction = "ExpandNode"
					ElseIf Instr(dicKeys(iCounter),"VerifyNode")>0 Then
						sSubAction = "VerifyNode"
					ElseIf Instr(dicKeys(iCounter),"VerifyCheckMark")>0 Then
						sSubAction = "VerifyCheckMark"
					ElseIf Instr(dicKeys(iCounter),"SelectCheckMark")>0 Then
						sSubAction = "SelectCheckMark"
					Else
						sSubAction = dicKeys(iCounter)
					End If
					sProperty = dicItems(iCounter)
					bFlag = False
					
					Select Case sSubAction
						Case "ExpandNode"
							If sProperty<>"" Then
								objPartEdit.WinTreeView("PartsTree").Expand(sProperty)
								If Err.Number < 0 Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to expand ["+sProperty+"] on [Control] panel of [ Part Edit ]dialog.")
									Set objPartEdit = Nothing
									Exit Function
								End If
								Wait 1
								bFlag = True
							End If
							
						Case "VerifyNode"
							If sProperty<>"" Then
								sContent = objPartEdit.WinTreeView("PartsTree").GetContent()
								If Instr(1,sContent,sProperty) = 0 Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to verify existence of node ["+sProperty+"] on [Control] panel of [ Part Edit ]dialog.")
									Set objPartEdit = Nothing
									Exit Function
								End If
								Wait 1
								bFlag = True
							End If
							
						Case "VerifyCheckMark"
							If sProperty<>"" Then
								iIndex = Window("VizMainWin").Dialog("PartEdit").WinTreeView("PartsTree").GetROProperty("checked")
								sNode = Window("VizMainWin").Dialog("PartEdit").WinTreeView("PartsTree").GetItem(cint(iIndex))
								If sProperty<>sNode Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to verify node ["+sProperty+"] has checkmark on [Control] panel of [ Part Edit ]dialog.")
									Set objPartEdit = Nothing
									Exit Function
								End If
								Wait 1
								bFlag = True
							End If
						
						Case "SelectCheckMark"
							If sProperty<>"" Then
								objPartEdit.WinTreeView("PartsTree").SetItemState sProperty, micDblClick
								If Err.Number < 0 Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to select checkmark for node ["+sProperty+"] on [Control] panel of [ Part Edit ]dialog.")
									Set objPartEdit = Nothing
									Exit Function
								End If
								Wait 1
								bFlag = True
							End If
					End Select
					
					If bFlag = False Then
						Fn_SISW_LifeView_PartEdit_Operations = False
						Set objPartEdit = Nothing
						Exit Function
					Else
						Fn_SISW_LifeView_PartEdit_Operations = True
					End If		
				Next
			
	End Select
	
	'Click on button provided
	If sButton<>"" Then
		bFlag1 = Fn_SISW_UI_WinButton_Operations("Fn_SISW_LifeView_PartEdit_Operations", "Click", objPartEdit,sButton,"","","")
		If bFlag1 = False Then
			Fn_SISW_LifeView_PartEdit_Operations = False
			Set objPartEdit = Nothing
			Exit Function
		End If
	End If
	
	Set objPartEdit = Nothing
End Function
'=============================================================================================================================
'  FUNCTION NAME   	:  Fn_SISW_LifeView_IdentifyDialog_Operations
'
'  DESCRIPTION     	:  Function is used to perform operations on Identify dialog
'
'
'  PARAMETERS   	:  sAction 		: 	Action to be performed
'				  	   dicIdentifyDtls 	: 	Dictionary object
'					   sButton 		: 	Button to be clicked
'					   sReserve 		: 	For future use
'
'	Return Value 	:  True or False
'										
'	How To Use 		:  Set dicIdentifyDtls = CreateObject("Scripting.Dictionary")
'							dicIdentifyDtls("WinButton") 	= "Clear"
'							dicIdentifyDtls("Mark Selection") = "Mark Result Values"
'						bReturn = Fn_SISW_LifeView_IdentifyDialog_Operations("Set",dicIdentifyDtls,"Cancel","")
'						bReturn = Fn_SISW_LifeView_IdentifyDialog_Operations("GetResults","","","")
'
'	History			:	Developer Name			Date			Version		Changes				Reviewer
'=============================================================================================================================
'	Created by  	:	Poonam Chopade	 	  15-Feb-2018	  	1.0		   	Created				TC11.5(20180122.00)_NewDevelopment_PoonamC
'=============================================================================================================================
Public Function Fn_SISW_LifeView_IdentifyDialog_Operations(sAction,dicIdentifyDtls,sButton,sReserve)
	
	Dim dicCount, dicItems, dicKeys
	Dim bFlag, iCounter, sSubAction, sProperty,sMenu
	Dim objIdentify
	
	Fn_SISW_LifeView_IdentifyDialog_Operations = False
	On Error Resume Next
	
	Set objIdentify = Fn_SISW_LifeView_GetObject("Identify")
	
	'check existence of dialog
	If Fn_SISW_UI_Object_Operations("Fn_SISW_LifeView_IdentifyDialog_Operations","Exist", objIdentify,"") = False Then
		sMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("Viz_Menu"), "CAEViewingIdentify")
		bReturn =  Fn_SISW_LifeView_MenuOperation("WinMenuSelect",sMenu)
		If bReturn = False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to Select Menu [ "+sMenu+" ]")
			Set objIdentify = Nothing
			Exit Function
		End If
		If Fn_SISW_UI_Object_Operations("Fn_SISW_LifeView_IdentifyDialog_Operations","Exist", objIdentify,"") = False Then
			Set objIdentify = Nothing
			Exit Function
		End If
	End If
	
	Select Case sAction
		Case "Set"
			dicCount = dicIdentifyDtls.Count
			dicItems = dicIdentifyDtls.Items
			dicKeys = dicIdentifyDtls.Keys
			
			For iCounter = 0 To dicCount - 1
				If Instr(dicKeys(iCounter),"WinButton")>0 Then
					sSubAction = "WinButton"
				Else
					sSubAction = dicKeys(iCounter)
				End If
				
				sProperty = dicItems(iCounter)
				bFlag = False
				
				Select Case sSubAction
					Case "Mark Selection"
						If sProperty<>"" Then
							objIdentify.WinComboBox("Mark Selection").Select sProperty
							If Err.Number < 0 Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to select option ["+sProperty+"] as [ Mark Selection ].")
								Set objIdentify = Nothing
								Exit Function
							End If
							Wait 1
							bFlag = True 
						End If
					Case "WinButton"
						If sProperty<>"" Then
							'Click on button provided
							objIdentify.WinButton(sProperty).Click 5,5,micLeftBtn
							If Err.Number < 0 Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to click ["+sProperty+"] button on [Identity] dialog.")
								Set objIdentify = Nothing
								Exit Function
							End If
							Wait 1
							bFlag = True
						End If
				End Select
				If bFlag = False Then
					Fn_SISW_LifeView_IdentifyDialog_Operations = False
					Set objIdentify = Nothing
					Exit Function
				Else
					Fn_SISW_LifeView_IdentifyDialog_Operations = True
				End If		
			Next
		Case "GetResults"
			Fn_SISW_LifeView_IdentifyDialog_Operations = Fn_UI_Object_GetROProperty("Fn_SISW_LifeView_IdentifyDialog_Operations",objIdentify.WinEditor("Results"),"text")
	End Select
	
	'Click on button provided
	If sButton<>"" Then
		bFlag = Fn_SISW_UI_WinButton_Operations("Fn_SISW_LifeView_IdentifyDialog_Operations", "Click", objIdentify,sButton,"","","")
		If bFlag = False Then
			Fn_SISW_LifeView_IdentifyDialog_Operations = False
			Set objIdentify = Nothing
			Exit Function
		End If
	End If
	
	Set objIdentify = Nothing
End Function
'=============================================================================================================================
'  FUNCTION NAME   	:  Fn_SISW_LifeView_ExportImage_Save_Ops
'
'  DESCRIPTION     	:  Function is used to perform operations on Export Image to Save dialog
'
'  PARAMETERS   	:  sAction 		: 	Action to be performed
'				  	   dicInfo 		: 	Dictionary object
'					   sButton 		: 	Button to be clicked
'
'	Return Value 	:  True or False
'										
'	How To Use 		:  Set dicInfo = CreateObject("Scripting.Dictionary")
'						   dicInfo("Filename") = "Comparison - Tricolor Mapping 1.png"						
'						bReturn = Fn_SISW_LifeView_ExportImage_Save_Ops("VerifyFileName",dicInfo,"")
'
'						Set dicInfo = CreateObject("Scripting.Dictionary")
'						    dicInfo("Filename") = "Test123.png"
'						bReturn = Fn_SISW_LifeView_ExportImage_Save_Ops("SaveFile",dicInfo,"Save")
'
'						bReturn = Fn_SISW_LifeView_ExportImage_Save_Ops("ButtonClick","","Save")
'
'	History			:	Developer Name			Date			Version		Changes				Reviewer
'=============================================================================================================================
'	Created by  	:	Poonam Chopade	 	  20-Feb-2018	  	1.0		   	Created				TC11.5(20180122.00)_NewDevelopment_PoonamC
'=============================================================================================================================
Public Function Fn_SISW_LifeView_ExportImage_Save_Ops(sAction,dicInfo,sButton)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_LifeView_ExportImage_Save_Ops"
	Dim ObjDialog	
	Fn_SISW_LifeView_ExportImage_Save_Ops=False
	Set ObjDialog = Window("LifeViewWin").Dialog("ExportImage")
	
	'Check existence
	If ObjDialog.Exist(2) <> True Then
		Set ObjDialog=nothing
		Exit Function
	End If
	
   	Select Case sAction
		Case "VerifyFileName"
			If trim(ObjDialog.WinEdit("Filename").GetROProperty("text")) = trim(dicInfo("FileName")) Then
				Fn_SISW_LifeView_ExportImage_Save_Ops=True
			Else
				Fn_SISW_LifeView_ExportImage_Save_Ops=False
			End If
	'-------------------------------------------------------------------------------------------------------------------------
		Case "SaveFile"
			ObjDialog.WinEdit("Filename").set dicInfo("FileName")
			wait 3
			If Err.Number < 0 Then
				Fn_SISW_LifeView_ExportImage_Save_Ops=False
			Else
				Fn_SISW_LifeView_ExportImage_Save_Ops=True
			End If	
	'-------------------------------------------------------------------------------------------------------------------------			
		Case "ButtonClick"
			Fn_SISW_LifeView_ExportImage_Save_Ops = Fn_SISW_UI_WinButton_Operations("Fn_SISW_LifeView_ExportImage_Save_Ops","click",ObjDialog, sButton,5,5,micLeftBtn)
	End Select
	
	If sButton <> "" Then
		If ObjDialog.WinButton(sButton).Exist(2) = True Then
			ObjDialog.WinButton(sButton).Click	
		else
			Set ObjDialog = nothing
			Exit function
		End if
	End If
	
	Set ObjDialog = nothing
	
End Function

''*********************************************************		Function to Perform operation on Export Dialog in Viz   ***********************************************************************
'
''Function Name		:		   Fn_SISW_LifeView_VIZ_ExportDialogOperation
'
''Description		:		   Function to Perform operation on Export Dialog in Vis Mockup that i.e standalone
'
'Parameters		    :			1. sAction = Action To Perform
'								2. dicInfo = Dictionary Object
'								3. sButton = Name of button to click. 	 Ex. "OK" or "cancel" 
'								4. sReserve = For Future Use  		NOTE : For Enter Name Dialog that appears after Export Image Dialog use 'sReserve'		
'
'			  										
''Return Value	    : 			True or False 
'
''Pre-requisite		:			Export Dialog should be displayed in LCV.
'
'
''Examples			:		 	Dim dicInfo
'								Set dicInfo=CreateObject("Scripting.Dictionary")
'								dicInfo("File") = "C:\Block_8567+".plmxml"
'								dicInfo("FileFormat") = "Product Structure (*.plmxml)"
'								dicInfo("Hierarchy") = "(AH) Alt Hier"										
'								bReturn= Fn_SISW_LifeView_VIZ_ExportDialogOperation("ExportImageandVerifyMessage", dicInfo , "OK", "")
'
'History:			Developer Name				Date						Rev. No.		Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'					Pravin Bhoyar			 26-Feb-2018					1.0				Created					Poonam Chopade					
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_LifeView_VIZ_ExportDialogOperation(sAction, dicInfo , sButton, sReserve)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_LifeView_VIZ_ExportDialogOperation"
	Dim DictItems,DictKeys,iCount,sMenu
	Dim iCounter,ObjDialog
	Dim ObjExporterDialog
	Fn_SISW_LifeView_VIZ_ExportDialogOperation=False
	iCount = 0
	
	Set ObjDialog = Fn_SISW_LifeView_GetObject("Export") 
	' Check existence of Export dialog
	If Fn_UI_ObjectExist("Fn_SISW_LifeView_VIZ_ExportDialogOperation",ObjDialog)=False Then
		sMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("LifecycleViewer_Menu"), "FileExport")
		Call Fn_SISW_LifeView_MenuOperation("WinMenuSelect",sMenu)
		Call Fn_ReadyStatusSync(3)
		
		wait 3
		If NOT ObjDialog.exist(4) Then
			Fn_SISW_LifeView_VIZ_ExportDialogOperation = False
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_LifeView_VIZ_ExportDialogOperation ] Export Dialog is not opened.")
			ObjDialog = nothing
			Exit Function	
		End if
	End If
	
	Select Case sAction
		Case "ExportImageandVerifyMessage"
			DictKeys = dicInfo.Keys
			DictItems = dicInfo.Items
			For iCounter = 0 to Ubound(DictKeys)							
				Select Case DictKeys(iCounter)
					Case "File"
						If DictItems(iCounter) <> "" Then
							If ObjDialog.WinEdit(DictKeys(iCounter)).Exist(1) Then
								ObjDialog.WinEdit(DictKeys(iCounter)).Set DictItems(iCounter)	
							End If
						End If
						
					Case "FileFormat"
						If DictItems(iCounter) <> "" Then
							If ObjDialog.WinComboBox(DictKeys(iCounter)).Exist(1) Then
								ObjDialog.WinComboBox(DictKeys(iCounter)).Select DictItems(iCounter)
							End If
						End If
				
					Case "Hierarchy"
						If DictItems(iCounter) <> "" Then
							If ObjDialog.WinList(DictKeys(iCounter)).Exist(1) Then
								ObjDialog.WinList(DictKeys(iCounter)).Select DictItems(iCounter)
							End If
						End If
					Case Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_LifeView_VIZ_ExportDialogOperation ] Invalid case [ " & sAction & " ].")
						Exit function
				End Select	
			Next
			If sButton <> "" Then
				If ObjDialog.WinButton(sButton).Exist(2) Then
					ObjDialog.WinButton(sButton).Click	
					Fn_SISW_LifeView_VIZ_ExportDialogOperation=true
				else
					Set ObjDialog = nothing
					Exit function
				End if
			End If
		
			Set ObjExporterDialog = Fn_SISW_LifeView_GetObject("Exporter")
			If ObjExporterDialog.Exist(4) = True Then
				If ObjExporterDialog.Static("Message").Exist(1) Then
					ObjExporterDialog.WinButton("OK").Click	
					Fn_SISW_LifeView_VIZ_ExportDialogOperation=True
				Else
					ObjExporterDialog.WinButton("OK").Click	
					Fn_SISW_LifeView_VIZ_ExportDialogOperation=False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_LifeView_VIZ_ExportDialogOperation ] Failed to verify [ Export Succeeeded ] Message.")
					ObjDialog=nothing
					Set ObjExporterDialog = nothing
					Exit Function	
				End If
			Else
				Set ObjExporterDialog = nothing
			End If
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_LifeView_VIZ_ExportDialogOperation ] Invalid case [ " & sAction & " ].")
			Exit function
	End Select
	
	If ObjDialog.exist(4) Then
		ObjDialog.WinButton("Cancel").Click	
	End if
	
	If Fn_SISW_LifeView_VIZ_ExportDialogOperation <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [ Fn_SISW_LifeView_VIZ_ExportDialogOperation ] executed successfuly with case [ " & sAction & " ].")
	else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_LifeView_VIZ_ExportDialogOperation ] Failed to executed with case [ " & sAction & " ].")
	End If

	Set ObjDialog = nothing
	Set ObjExporterDialog = nothing
End Function
'=============================================================================================================================
'  FUNCTION NAME   	:  Fn_SISW_LifeView_AppearanceImagePalette_Ops()
'
'  DESCRIPTION     	:  Function is used to perform operations on Appearance Image Palette dialog
'
'  PARAMETERS   	:  sAction 		: 	Action to be performed
'				  	   dicInfo 		: 	Dictionary object
'					   sClose 		: 	Yes / No
'
'	Return Value 	:  True or False
'										
'	How To Use 		:  Set dicInfo = CreateObject("Scripting.Dictionary")
'						   dicInfo("TabName") = "Images"
'						   dicInfo("Storage") = "MyComputer"	
'						   dicInfo("FileFolderPath") = "C:\Temp\image1.jpg"
'						bReturn = Fn_SISW_LifeView_AppearanceImagePalette_Ops("AddNewImage",dicInfo,"Yes")
'
'	History			:	Developer Name			Date			Version		Changes				Reviewer
'=============================================================================================================================
'	Created by  	:	Poonam Chopade	 	  02-Mar-2018	  	1.0		   	Created				TC11.5(20180212.00)_NewDevelopment_PoonamC
'=============================================================================================================================
Public Function Fn_SISW_LifeView_AppearanceImagePalette_Ops(sAction,dicInfo,sClose)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_LifeView_AppearanceImagePalette_Ops"
	Dim ObjAppWindow,ObjAppPalette,sMenu,bFlag
	
	Fn_SISW_LifeView_AppearanceImagePalette_Ops=False
	Set ObjAppWindow = Window("VizMainWin").Window("AppearancePalette")
	Set ObjAppPalette = Window("VizMainWin").WinObject("AppearancePalette")
	
	'Check existence
	If ObjAppWindow.Exist(2) <> True and ObjAppPalette.Exist(2) <> True Then
		sMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("Viz_Menu"),"ConceptAppearanceImagePalette")
		Call Fn_SISW_LifeView_MenuOperation("WinMenuSelect",sMenu)
		Wait 2
		'check existence
		If ObjAppWindow.Exist(2) <> True Then
			If ObjAppPalette.Exist(2) = True Then
				ObjAppPalette.DblClick 5,5,micLeftBtn
				Wait 2			
			Else
				Set ObjAppWindow = Nothing
				Set ObjAppPalette = Nothing
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_LifeView_AppearanceImagePalette_Ops ] Appearance Image Palette is not Exists.")
				Exit Function
			End If
			
			If ObjAppWindow.Exist(2) <> True Then
				Set ObjAppWindow = Nothing
				Set ObjAppPalette = Nothing
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_LifeView_AppearanceImagePalette_Ops ] Appearance Image Palette is not Exists.")
				Exit Function
			End if
		End If
	End If
	
   	Select Case sAction
		Case "AddNewImage"
			If dicInfo("TabName")<>"" Then 'Select Tab 
				ObjAppWindow.WinTab("TabName").Select dicInfo("TabName")
				Wait 1
			End If
			'Click on New button
			ObjAppWindow.WinButton("New").Click 5,5,micLeftBtn  
			Wait 1
			'Select file name
			If dicInfo("Storage") <> "" and dicInfo("FileFolderPath") <> "" Then
				bFlag = Fn_SISW_LifeView_FileOpenInsertOperation("Open",dicInfo)
				If bFlag = False Then
					Set ObjAppWindow = Nothing
					Set ObjAppPalette = Nothing
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_LifeView_AppearanceImagePalette_Ops ] failed to open file [ "+dicInfo("FileFolderPath")+" ].")
					Exit Function
				Else
					Fn_SISW_LifeView_AppearanceImagePalette_Ops = bFlag
				End If
			End If			
	End Select
	
	If lcase(sClose) = "yes" Then
		 ObjAppWindow.Close()
	End If
	
	Set ObjAppWindow = nothing
	Set ObjAppPalette = Nothing
	
End Function
'=============================================================================================================================
'  FUNCTION NAME   	:  Fn_SISW_LifeView_ComparisonPreferences_Ops()
'
'  DESCRIPTION     	:  Function is used to perform operations on Comparison Preferences dialog
'
'  PARAMETERS   	:  sAction 		: 	Action to be performed
'				  	   dicInfo 		: 	Dictionary object
'					   sButton 		: 	OK / Cancel
'
'	Return Value 	:  True or False
'										
'	How To Use 		:  Set dicInfo = CreateObject("Scripting.Dictionary")
'						   dicInfo("TabName") = "Tricolor"
'						   dicInfo("IncreaseTolerance") = 20
'					   bReturn = Fn_SISW_LifeView_ComparisonPreferences_Ops("IncreaseTolerance",dicInfo,"OK")
'
'					    Set dicInfo = CreateObject("Scripting.Dictionary")
'							 dicInfo("TabName") = "Tricolor"
'						    dicInfo("DecreaseTolerance") = 10		
'					   bReturn = Fn_SISW_LifeView_ComparisonPreferences_Ops("DecreaseTolerance",dicInfo,"OK")						   
'						
'
'	History			:	Developer Name			Date			Version		Changes				Reviewer
'=============================================================================================================================
'	Created by  	:	Poonam Chopade	 	  20-Mar-2018	  	1.0		   	Created				TC11.5(20180305a.00)_NewDevelopment_PoonamC
'=============================================================================================================================
Public Function Fn_SISW_LifeView_ComparisonPreferences_Ops(sAction,dicInfo,sButton)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_LifeView_ComparisonPreferences_Ops"
	Dim ObjComparisonPref,iCount
	
	Fn_SISW_LifeView_ComparisonPreferences_Ops=False
	Set ObjComparisonPref = Window("LifeViewWin").Dialog("ComparisonPreferences")
	
	'Check existence
	If ObjComparisonPref.Exist(2) <> True Then
		Set ObjComparisonPref = Nothing
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_LifeView_ComparisonPreferences_Ops ] Comparison Preferences dialog is not Exists.")
		Exit Function
	End If
	
	'Select Tab
	If dicInfo("TabName")<>"" Then
		ObjComparisonPref.WinTab("TabName").Select dicInfo("TabName")
		Wait 1
	End If
	
   	Select Case sAction
		Case "IncreaseTolerance"
			If dicInfo("IncreaseTolerance")<>"" Then 'Click on button
				For iCount = 1 to cint(dicInfo("IncreaseTolerance"))
					ObjComparisonPref.WinButton("MaxTolerance").Click 5,5,micLeftBtn  
					Wait 1
				Next
			End If
			If Err.Number < 0 Then
				Fn_SISW_LifeView_ComparisonPreferences_Ops = False
				Set ObjComparisonPref = Nothing
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_LifeView_ComparisonPreferences_Ops ] failed to Increase Tolerance.")
				Exit Function
			End If
		Case "DecreaseTolerance"
			If dicInfo("DecreaseTolerance")<>"" Then 'Click on button
				For iCount = 1 to cint(dicInfo("DecreaseTolerance"))
					ObjComparisonPref.WinButton("MinTolerance").Click 5,5,micLeftBtn  
					Wait 1
				Next
			End If
			If Err.Number < 0 Then
				Fn_SISW_LifeView_ComparisonPreferences_Ops = False
				Set ObjComparisonPref = Nothing
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_LifeView_ComparisonPreferences_Ops ] failed to Decrease Tolerance.")
				Exit Function
			End If					
	End Select
	
	If sButton <> "" Then
		sButton = Split(sButton,"~")
		For iCount = 0 to UBound(sButton)
			ObjComparisonPref.WinButton(sButton(iCount)).Click 5,5,micLeftBtn  
			Wait 1
		Next	
	End If
	
	Fn_SISW_LifeView_ComparisonPreferences_Ops = True
	
	Set ObjComparisonPref = nothing
	
End Function
