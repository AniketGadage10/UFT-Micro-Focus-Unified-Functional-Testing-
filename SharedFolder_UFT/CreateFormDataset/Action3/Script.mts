'Option Explicit
'		
''**********************************************************************************
''Variable Declaration
''**********************************************************************************
'Dim bReturn
'
''**********************************************************************************
''Load Required Library Files
''**********************************************************************************
'ExecuteFile (Environment.Value("sPath") + "\Library\LogUtil.vbs")
'ExecuteFile (Environment.Value("sPath") + "\Library\Setup.vbs")
'ExecuteFile (Environment.Value("sPath") + "\Library\UI_Library.vbs")
'ExecuteFile (Environment.Value("sPath") + "\Library\GeneralFunctions.vbs")
'ExecuteFile (Environment.Value("sPath") + "\Library\MyTeamcenter.vbs")
'ExecuteFile (Environment.Value("sPath") + "\Library\StructureMananger.vbs")
'
''**********************************************************************************
''Action 3 Execution
''**********************************************************************************
'Call Fn_UpdateLogFiles("-------------------------------------", "")
'Call Fn_UpdateLogFiles("******************* Action 3 Execution *********************", "")
'Call Fn_UpdateLogFiles("-------------------------------------", "")
'Call Fn_UpdateLogFiles(vblf, "")
'
''**********************************************************************************
''Print Associated Libraries Details
'''**********************************************************************************
'Call Fn_UpdateLogFiles("List Library Operation: Start", "")
'Call Fn_UpdateLogFiles("Library File [LogUtil.vbs] Associated", "")
'Call Fn_UpdateLogFiles("Library File [Setup.vbs] Associated", "")
'Call Fn_UpdateLogFiles("Library File [GeneralFunctions.vbs] Associated", "")
'Call Fn_UpdateLogFiles("Library File [StructureMananger.vbs] Associated", "")
'Call Fn_UpdateLogFiles("List Library Operation: End", "")
'Call Fn_UpdateLogFiles(vblf, "")
'
''**********************************************************************************
''Associate Object Repositories
'''**********************************************************************************
'Call Fn_UpdateLogFiles("OR Association Operation: Start", "")
'bReturn = Fn_LoadORFile("Action3", Environment.Value("sPath") + "\ObjectRepository\General.tsr")
'If bReturn = False Then
'	Call Fn_UpdateLogFiles("Failed to Associate Object Repository [General.tsr]", "FAIL:Failed to Associate [General.tsr] OR")
'	ExitTest
'Else
'	Call Fn_UpdateLogFiles("Successfully Associate Object Repository [General.tsr]", "")
'End If
'bReturn = Fn_LoadORFile("Action3", Environment.Value("sPath") + "\ObjectRepository\StructureManager.tsr")
'If bReturn = False Then
'	Call Fn_UpdateLogFiles("Failed to Associate Object Repository [StructureManager.tsr]", "FAIL:Failed to Associate [StructureManager.tsr] OR")
'	ExitTest
'Else
'	Call Fn_UpdateLogFiles("Successfully Associate Object Repository [StructureManager.tsr]", "")
'End If
'Call Fn_UpdateLogFiles("OR Association Operation: End", "")
'Call Fn_UpdateLogFiles(vblf, "")
'
'''**********************************************************************************
'''Set the "Tc_Allow_Longer_ID_Name" to False
'''**********************************************************************************
''Call Fn_UpdateLogFiles("[" + Cstr(now) + "] Preference Set Operation: Start", "")
''bReturn = Fn_Preference_Search_Operation("Modify","TC_Allow_Longer_ID_Name","True")
''If bReturn = False Then
''	Call Fn_UpdateLogFiles("Failed to Modify the Set Preference", "FAIL:Failed to Modify the Set Preference")
''Else
''	Call Fn_UpdateLogFiles("Successfully Selected  and Replaced Preference Options", "")
''End If
''Call Fn_UpdateLogFiles("[" + Cstr(now) + "] Preference Set Operation: End", "")
''Call Fn_UpdateLogFiles(vblf, "")
'
''**********************************************************************************
''Kill current session and invoke new session as TcUser2
'''**********************************************************************************
'Call Fn_UpdateLogFiles("[" + Cstr(now) + "] Teamcenter Login Operation: Start", "")
'bReturn = Fn_ReUserTcSession(False, False, Environment.Value("TcUser2"))
'If bReturn = False Then
' Call Fn_UpdateLogFiles("Failed to Find Tc Session for User [" + Environment.Value("TcUser2") + "]", "FAIL:Tc Session Not Found")
' ExitTest
'Else
' Call Fn_UpdateLogFiles("Successfully Found Tc Session for User [" + Environment.Value("TcUser2") + "]", "")
'End If
'Call Fn_UpdateLogFiles("[" + Cstr(now) + "] Teamcenter Login Operation: End", "")
'Call Fn_UpdateLogFiles(vblf, "")
'Call Fn_ReadyStatusSync(1)
'
''**********************************************************************************
''Set the MyTeamcenter Perspective
'''**********************************************************************************
'Call Fn_UpdateLogFiles("[" + Cstr(now) + "] Set Perspective Operation: Start", "")
'bReturn = Fn_SetPerspective("My Teamcenter")
'If bReturn = False Then
'	Call Fn_UpdateLogFiles("Failed to Set My Teamcenter Perspective", "FAIL:Failed to Go to MyTc Module")
'	ExitTest
'Else
'	Call Fn_UpdateLogFiles("Successfully set My Teamcenter Perspective", "")
'End If
'Call Fn_UpdateLogFiles("[" + Cstr(now) + "] Set Perspective Operation: End", "")
'Call Fn_UpdateLogFiles(vblf, "")
'
''**********************************************************************************
''Synchronization for Ready state
''**********************************************************************************
'Call Fn_ReadyStatusSync(1)
'
''**********************************************************************************
''Reset MyTc Perspective to Display Default state
''**********************************************************************************
'Call Fn_ResetPerspective()
'Call Fn_ReadyStatusSync(1)
'
''**********************************************************************************
''Select the newly created Dataset
'''**********************************************************************************
'Call Fn_MyTc_NavTree_NodeOperation("Select","Home:" + Datatable("DatasetName", dtGlobalSheet),"")
'
''**********************************************************************************
''Create New Form (More than 32 chara length) under the newly created Dataset 
'''**********************************************************************************
'Call Fn_UpdateLogFiles("[" + Cstr(now) + "] Test Form Creation Operation: Start", "")
'bReturn = Fn_FormCreate(Datatable("LongFormName", dtGlobalSheet), DataTable("FormDesc", dtGlobalSheet), DataTable("FormType", dtGlobalSheet),DataTable("OpenOnCreate", dtGlobalSheet))
'If bReturn = False Then
'	Call Fn_UpdateLogFiles("Failed to Create Test Form [" + DataTable("LongFormName", dtGlobalSheet) + "]", "FAIL: Test Form Creation Failed")
'	ExitTest
'Else
'	Call Fn_ReadyStatusSync(1)
'	Call Fn_UpdateLogFiles("Successfully Created Test Form [" + DataTable("LongFormName", dtGlobalSheet) + "]", "")
'End If
'Call Fn_UpdateLogFiles("[" + Cstr(now) + "] Test Form Creation Operation: End", "")
'Call Fn_UpdateLogFiles(vblf, "")
'
''**********************************************************************************
''Verify the character length of the last created form. 
'''**********************************************************************************
'If Fn_MyTc_NavTree_NodeOperation("Exist","Home:" + Datatable("DatasetName", dtGlobalSheet) + ":" + Datatable("LongFormName", dtGlobalSheet),"") = True Then
'	Call Fn_WriteLogFile("","Fail: Form created with more than 32 chara length even though Tc_Allow_Longer_ID_Name is set to False")
'Else
'	Call Fn_WriteLogFile("", "Pass: Form NOT created with more than 32 chara length when Tc_Allow_Longer_ID_Name is set to False")
'End if
'
''**********************************************************************************
''Log Test Result
'''**********************************************************************************
'Call Fn_UpdateLogFiles("Execution End Time :: " + cStr(now), "")
'Call Fn_UpdateLogFiles("-------------------------------------", "")
'Call Fn_UpdateLogFiles("Test Execution Result: PASS", "PASS: All VP Pass")
'Call Fn_UpdateLogFiles("-------------------------------------", "")
'
''**********************************************************************************
''Call for Code Coverage
''**********************************************************************************
'Call Fn_CodeCover_Exit()
'
'