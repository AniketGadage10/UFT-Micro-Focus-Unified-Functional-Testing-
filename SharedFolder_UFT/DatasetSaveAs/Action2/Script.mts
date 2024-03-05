Option Explicit
		
'**********************************************************************************
'Variable Declaration
'**********************************************************************************
Dim bReturn, iRanNo
Dim sName,sMenu,sNewDataset

'**********************************************************************************
'Action 2 Execution Start
'**********************************************************************************
Call Fn_UpdateLogFiles("["+"["+Cstr(time) + "] ********************************* QTP Action2 - Start ************************************", "")
Environment.Value("UniqueNo") = Fn_RandNoGenerate()

''**********************************************************************************************
'Create TestCase Folder under AutomatedTests folder
'**********************************************************************************
bReturn = Fn_MyTc_TestcaseFolderCreate("","","")
If bReturn = False Then
	Call Fn_UpdateLogFiles("[" + Cstr(now) + "] - Action - FAIL | TestCase Folder [ "&bReturn&" ] is not Created under [ "&GBL_AUTOMATEDTEST_FOLDER_PATH&" ] ", "[" + Cstr(now) + "] - Action - FAIL | TestCase Folder Not Created")
	Call Fn_KillProcess(Environment.Value("KillProcesses"))
	ExitTest
Else
	Call Fn_UpdateLogFiles("[" + Cstr(now) + "] - Action - PASS | Successfully Created TestCase Folder [ "&bReturn&" ] under [ "&GBL_AUTOMATEDTEST_FOLDER_PATH&" ] ", "")
End If 
Environment.Value("TestFolderName") = GBL_AUTOMATEDTEST_FOLDER_PATH  &  ":"  &  bReturn
		
'**********************************************************************************
'Expand and Select Test Case Folder
''**********************************************************************************
bReturn =  Fn_MyTc_NavTree_NodeOperation("Expand",Environment.Value("TestFolderName"),"")
If bReturn = False Then
	Call Fn_UpdateLogFiles("[" + Cstr(now) + "] - Action - FAIL | Failed to Expand TestCase Folder [ "&Environment.Value("TestFolderName")&" ] under [ "&GBL_AUTOMATEDTEST_FOLDER_PATH&" ] ", "[" + Cstr(now) + "] - Action - FAIL | TestCase Folder Not Created")
	Call Fn_KillProcess(Environment.Value("KillProcesses"))
	ExitTest
Else
	Call Fn_UpdateLogFiles("[" + Cstr(now) + "] - Action - PASS | Successfully Expanded TestCase Folder [ "&Environment.Value("TestFolderName")&" ] under [ "&GBL_AUTOMATEDTEST_FOLDER_PATH&" ] ", "")
End If
Call Fn_ReadyStatusSync(1)



'**********************************************************************************
'Create Dataset  under Test case Folder
''**********************************************************************************
Environment.Value("DatasetType") = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("Dataset_Type"), "Text")
sName = Environment.Value("UniqueNo") + "_1"
bReturn = Fn_DatasetCreate(sName,DataTable("DatasetDesc", dtGlobalSheet),DataTable("ToolsUsed", dtGlobalSheet),Environment.Value("sPath") + "\" +DataTable("Filename", dtGlobalSheet),DataTable("OpenonCreate", dtGlobalSheet),Environment.Value("DatasetType"))
If bReturn = False Then
    	Call Fn_UpdateLogFiles("["+Cstr(time) + "]  - Action - FAIL | Failed to Create Dataset [" + sName + "]", "FAIL : Dataset Creation Failed")
	Call Fn_KillProcess(Environment.Value("KillProcesses"))
    	ExitTest
End If

Call Fn_ReadyStatusSync(1)
Call Fn_UpdateLogFiles("["+Cstr(time) + "]  - Action - PASS | Successfully Created Dataset[" + sName + "]", "")
Environment.Value("DatasetName") = Environment.Value("TestFolderName") +":"+ sName 

'**********************************************************************************
'Select  the Created Dataset and  do copy after right click on the dataset
'**********************************************************************************
bReturn= Fn_MyTc_NavTree_NodeOperation("Select",Environment.Value("DatasetName")   ,"")
If bReturn=True then
	Call Fn_UpdateLogFiles(Time() & " - " & "Action - Pass | Successfully Selected  [ "&Datatable("DataSetName",dtGlobalsheet)&"] ", "")
	Call Fn_ReadyStatusSync(1)
Else
	Call Fn_UpdateLogFiles(Time() & " - " & "Action - Fail |  Failed to Select Dataset [ "&Datatable("DataSetName",dtGlobalsheet)&"] ", "FAIL :  Failed to select  Dataset ")
	Call Fn_KillProcess("")
	ExitTest
End If

sMenu=Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("RAC_Menu"),"EditCopy")
If Fn_MenuOperation("Select",sMenu) = True Then
	Call Fn_UpdateLogFiles(Time() & " - " & "Action - Pass | Successfully Performed Menu operation [ "&sMenu&" ] ", "")
	Call Fn_ReadyStatusSync(1)
Else
	Call Fn_UpdateLogFiles(Time() & " - " & "Action - Fail |  Failed to Perform Menu operation [ "&sMenu&" ]", "FAIL : Failed to Perform Menu operation [ "&sMenu&" ]")
	Call Fn_KillProcess("")
	ExitTest
End If


bReturn = Fn_MyTc_NavTree_NodeOperation("Select",Environment.Value("TestFolderName"),"")
If bReturn=True then
	Call Fn_UpdateLogFiles(Time() & " - " & "Action - Pass | Successfully Selected  Folder [ "&Environment.Value("TestFolderName")&"] ", "")
'	Call Fn_ReadyStatusSync(1)
Else
	Call Fn_UpdateLogFiles(Time() & " - " & "Action - Fail |  Failed to Select Folder [ "&Environment.Value("TestFolderName")&"] ", "FAIL :  Failed to select  Folder [ "&Environment.Value("TestFolderName")&"]")
	Call Fn_KillProcess("")
	ExitTest
End If

'**********************************************************************************
'Create  Folder under Test Case Folder
''**********************************************************************************
bReturn = Fn_MyTc_FolderCreate("Folder",Environment.Value("UniqueNo") ,DataTable("FolderDesc", dtGlobalSheet),DataTable("FolderOnCreate", dtGlobalSheet))
If bReturn = False Then
    	Call Fn_UpdateLogFiles("["+Cstr(time) + "]  - Action - FAIL | Failed to Create  Folder as  [" + Environment.Value("UniqueNo") + "]", "FAIL : Folder Creation Failed")
    	Call Fn_KillProcess(Environment.Value("KillProcesses"))
	ExitTest
End If
Call Fn_ReadyStatusSync(1)
Call Fn_UpdateLogFiles("["+Cstr(time) + "]  - Action - PASS | Successfully Created Folder[" +  Environment.Value("UniqueNo")  + "]", "")
Environment.Value("TestCaseFolderName") = Environment.Value("TestFolderName") + ":" + Environment.Value("UniqueNo") 

'**********************************************************************************
'Expand and Select Tests Case folder under Automated Test Folder
'**********************************************************************************
bReturn = Fn_MyTc_NavTree_NodeOperation("Select",Environment.Value("TestCaseFolderName") ,"")
If bReturn=True then
	Call Fn_UpdateLogFiles(Time() & " - " & "Action - Pass | Successfully Selected  Folder [ "&Environment.Value("TestCaseFolderName")&"] ", "")
'	Call Fn_ReadyStatusSync(1)
Else
	Call Fn_UpdateLogFiles(Time() & " - " & "Action - Fail |  Failed to Select Folder [ "&Environment.Value("TestCaseFolderName")&"] ", "FAIL :  Failed to select  Folder [ "&Environment.Value("TestCaseFolderName")&"]")
	Call Fn_KillProcess("")
	ExitTest
End If
'**********************************************************************************
'Paste selected Dataset in to Created folder 
''**********************************************************************************
sMenu=Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("RAC_Menu"),"EditPaste")
bReturn = Fn_MenuOperation("Select",sMenu)
If bReturn = FALSE Then
	Call Fn_UpdateLogFiles("["+Cstr(time) + "] - Verify - VP1 - FAIL | Failed to Verify That The Dataset ["+ sName +"] That was Copied is Pasted Under The Folder[" + Environment.Value("UniqueNo") + "]", " FAIL : Failed to Paste Dataset in to Created Folder")
   	Call Fn_KillProcess(Environment.Value("KillProcesses"))
	ExitTest
End If
Call Fn_UpdateLogFiles("["+Cstr(time) + "] - Verify - VP1 - PASS | Successfully  Verified That The Dataset ["+ sName +"] That was Copied is Pasted Under The Folder[" + Environment.Value("UniqueNo") + "]", "")
Environment.Value("CopiedDatset") =  Environment.Value("TestCaseFolderName") + ":" + sName
Call Fn_ReadyStatusSync(1)

''**********************************************************************************
'Select the copied dataset
'**********************************************************************************
bReturn= Fn_MyTc_NavTree_NodeOperation("Select",Environment.Value("CopiedDatset"),"")
If bReturn=True then
	Call Fn_UpdateLogFiles(Time() & " - " & "Action - Pass | Successfully Dataset [ "&Environment.Value("CopiedDatset")&"] ", "")
'	Call Fn_ReadyStatusSync(1)
Else
	Call Fn_UpdateLogFiles(Time() & " - " & "Action - Fail |  Failed to Select Dataset [ "&Environment.Value("CopiedDatset")&"] ", "FAIL :  Failed to select  Dataset [ "&Environment.Value("CopiedDatset")&"]")
	Call Fn_KillProcess("")
	ExitTest
End If

'**********************************************************************************
'Dataset  save as operation under newstuff  folder
'''********************************************************************************
bReturn = Fn_MyTc_DatasetSaveAs(DataTable("DatasetSaveAs", dtGlobalSheet) +  Environment.Value("UniqueNo") ,DataTable("DataSetSaveAsDesc", dtGlobalSheet), "1", DataTable("OpenonCreate", dtGlobalSheet)) 
If bReturn = FALSE Then
	Call Fn_UpdateLogFiles("["+Cstr(time) + "]   - Action - FAIL | Failed to Save as Dataset ["+DataTable("DatasetSaveAs", dtGlobalSheet) +  Environment.Value("UniqueNo")+"] Under NewStuff Folder", "FAIL : Dataset Save as Failed")
    	Call Fn_KillProcess(Environment.Value("KillProcesses"))
	ExitTest
End If
Call Fn_UpdateLogFiles("["+Cstr(time) + "]  - Action - PASS | Successfully Save as Dataset ["+DataTable("DatasetSaveAs", dtGlobalSheet) +  Environment.Value("UniqueNo") +"] Under NewStuff Folder", "")
Environment.Value("SaveAsDataSetPath") = GBL_NEWSTUFF_FOLDER_PATH +":" + DataTable("DatasetSaveAs", dtGlobalSheet) +  Environment.Value("UniqueNo")


''**********************************************************************************
' Select  NewStuff folder and save as DataSet
'**********************************************************************************
bReturn= Fn_MyTc_NavTree_NodeOperation("Expand",GBL_NEWSTUFF_FOLDER_PATH,"")
If bReturn=True then
	Call Fn_UpdateLogFiles(Time() & " - " & "Action - Pass | Successfully Expanded Folder [ "&GBL_NEWSTUFF_FOLDER_PATH&"] ", "")

Else
	Call Fn_UpdateLogFiles(Time() & " - " & "Action - Fail |  Failed to Expand Folder [ "&GBL_NEWSTUFF_FOLDER_PATH&"] ", "FAIL :  Failed to Expand Folder [ "&GBL_NEWSTUFF_FOLDER_PATH&"]")
	Call Fn_KillProcess("")
	ExitTest
End If
Call Fn_ReadyStatusSync(1)

bReturn = Fn_MyTc_NavTree_NodeOperation("Exist",Environment.Value("SaveAsDataSetPath"),"")
If bReturn = FALSE Then
    	Call Fn_UpdateLogFiles("["+Cstr(time) + "] - Verify - VP2 - FAIL | Failed to Verify That when File-->Save As... is Performed The Dataset gets Saved in The NewStuff Folder with The Specified Name as ["+DataTable("DatasetSaveAs", dtGlobalSheet) +  Environment.Value("UniqueNo") +"]", " FAIL : Dataset Not Placed In NewStuff Folder")
   	Call Fn_KillProcess(Environment.Value("KillProcesses"))
	ExitTest
End If
Call Fn_UpdateLogFiles("["+Cstr(time) + "] - Verify - VP2 - PASS | Successfully Verified That when File-->Save As... is Performed The Dataset gets Saved in The NewStuff Folder with The Specified Name as ["+DataTable("DatasetSaveAs", dtGlobalSheet) +  Environment.Value("UniqueNo") +"]", "")


'**********************************************************************************
'Verify the text file contains the same text as the original
'**********************************************************************************
bReturn= Fn_MyTc_NavTree_NodeOperation("DoubleClick",Environment.Value("SaveAsDataSetPath") ,"")
If bReturn=True then
	Call Fn_UpdateLogFiles(Time() & " - " & "Action - Pass | Successfully Clicked on Dataset [ "&Environment.Value("SaveAsDataSetPath")&"] ", "")
	Call Fn_ReadyStatusSync(1)
Else
	Call Fn_UpdateLogFiles(Time() & " - " & "Action - Fail |  Failed to Click on Dataset [ "&Environment.Value("SaveAsDataSetPath")&"] ", "FAIL :  Failed to Click on Dataset [ "&Environment.Value("SaveAsDataSetPath")&"]")
	Call Fn_KillProcess("")
	ExitTest
End If

bReturn = Fn_Dataset_Operations("TextVerify",DataTable("FileContent", dtGlobalSheet))
If bReturn = FALSE Then
    	Call Fn_UpdateLogFiles("["+Cstr(time) + "] - Verify - VP3 - FAIL | Dataset  ["+DataTable("DatasetSaveAs", dtGlobalSheet) +  Environment.Value("UniqueNo") +"] Contains The Same Text as The Original", "FAIL : Dataset contains the same text as the original")
    	Call Fn_KillProcess(Environment.Value("KillProcesses"))
	ExitTest
End If
Call Fn_UpdateLogFiles("["+Cstr(time) + "] - Verify - VP3 - PASS | Successfully Verified Dataset  ["+DataTable("DatasetSaveAs", dtGlobalSheet) +  Environment.Value("UniqueNo") +"] Contains The Same Text as The Original ", "")


'**********************************************************************************
'Look for Existing Tc Session for TcUserDBA from EnvVar_Ext.xml file
''**********************************************************************************
Call Fn_UpdateLogFiles("[" + Cstr(time) + "] -----------------------------------------PR#1741114 Implementation---------------------------------------------.", "")		
bPrefReset = false
bReturn = Fn_ReUserTcSession(True, True, Environment.Value("TcUserDBA") )
bPrefReset = true
If bReturn = False Then
	Call Fn_UpdateLogFiles("["+Cstr(time) + "] - Action - Fail  | Failed to Find Tc Session for User [" + Environment.Value("TcUserDBA") + "]", "FAIL:Tc Session Not Found")
    	Call Fn_KillProcess(Environment.Value("KillProcesses"))
   	ExitTest
Else
	Call Fn_UpdateLogFiles("[" + Cstr(time) + "] - Action - Pass | Successfully Found Tc Session for User [" + Environment.Value("TcUserDBA") + "]", "")
End If
Call Fn_ReadyStatusSync(1)


'*********************************************************************************
''Set TcUserDBA Sessions
'*********************************************************************************
Call Fn_SetTCSession("AutoTestDBA")

'*********************************************************************************
''Verify the preference 
'*********************************************************************************
bReturn = Fn_PreferenceOperations("VerifyPreference","TC_Allow_Longer_ID_Name","","","","","","","","","","")
If bReturn = False Then
	Call Fn_UpdateLogFiles("[" + Cstr(time) + "] - ACTION - FAIL | Failed to Verify the preference TC_Allow_Longer_ID_Name", "Fail : Fail to Verify Preference.")
Else
	Call Fn_UpdateLogFiles("[" + Cstr(time) + "] - ACTION - PASS | Successfully Verified the Preference TC_Allow_Longer_ID_Name.", "")
End If


'*********************************************************************************
''Set the preference
'*********************************************************************************
bReturn= Fn_Preference_Search_Operation("Modify","TC_Allow_Longer_ID_Name","false")
If bReturn = False Then
	Call Fn_UpdateLogFiles("[" + Cstr(time) + "] - ACTION - FAIL | Failed to set the TC_Allow_Longer_ID_Name preference", "FAIL : Failed to set the TC_Allow_Longer_ID_name preference")
    	Call Fn_KillProcess(Environment.Value("KillProcesses"))
    	ExitTest
Else
	Call Fn_UpdateLogFiles("[" + Cstr(time) + "] - ACTION - PASS | Successfully Set the TC_Allow_Longer_ID_Name Preference to False.", "")
End If


'*********************************************************************************
''Set Regular User Sessions
'*********************************************************************************
Call Fn_SetTCSession("AutoTest1")

'**********************************************************************************
'Set the MyTeamcenter Perspective
''*********************************************************************************
bReturn = Fn_SetPerspective("My Teamcenter")
If bReturn = False Then
	Call Fn_UpdateLogFiles("["+Cstr(time) + "] - Action - Fail | Failed to Set My Teamcenter Perspective", "FAIL:Failed to Go to MyTc Module")
	Call Fn_KillProcess(Environment.Value("KillProcesses"))
    	ExitTest
Else
	Call Fn_UpdateLogFiles("["+Cstr(time) + "] - Action - Pass | Successfully set My Teamcenter Perspective", "")
End If

'**********************************************************************************************
'Check the Existance of AutomatedTest folder under Home & Create if not exist
''**********************************************************************************************
bReturn = Fn_MyTc_NavTree_NodeOperation("Exist",GBL_AUTOMATEDTEST_FOLDER_PATH,"")
If bReturn = False Then
	Call Fn_MyTc_NavTree_NodeOperation("Select",Split(GBL_AUTOMATEDTEST_FOLDER_PATH,":")(0),"")
	bReturn = Fn_MyTc_FolderCreate("Folder",Split(GBL_AUTOMATEDTEST_FOLDER_PATH,":")(1),"Automation Artifact","OFF")
		If bReturn = False Then
			Call Fn_UpdateLogFiles("["+Cstr(time) + "]   - Action - FAIL | Failed to Create [ "&Split(GBL_AUTOMATEDTEST_FOLDER_PATH,":")(1)&" ] Folder under [ "&Split(GBL_AUTOMATEDTEST_FOLDER_PATH,":")(0)&" ] in NavTree", "FAIL : AutomatedTests Folder Creation Failed")
			Call Fn_KillProcess(Environment.Value("KillProcesses"))
            		ExitTest
		End If
	Call Fn_ReadyStatusSync(1)
	Call Fn_UpdateLogFiles("["+Cstr(time) + "]   - Action - PASS | Successfully Created [ "&Split(GBL_AUTOMATEDTEST_FOLDER_PATH,":")(1)&" ] Folder under [ "&Split(GBL_AUTOMATEDTEST_FOLDER_PATH,":")(0)&" ]", "")
End If


'**********************************************************************************
'Expand and Select AutomatedTests Folder
''**********************************************************************************
Call Fn_MyTc_NavTree_NodeOperation("ExpandAndSelect",GBL_AUTOMATEDTEST_FOLDER_PATH,"")

'**********************************************************************************
'Create Dataset  under Test case Folder
''**********************************************************************************
DataTable("DatasetName", dtGlobalSheet) = Fn_RandomString("DataSet",32)
bReturn = Fn_DatasetCreate(DataTable("DatasetName", dtGlobalSheet),DataTable("DatasetDesc", dtGlobalSheet),DataTable("ToolsUsed", dtGlobalSheet), "" ,DataTable("OpenonCreate", dtGlobalSheet),Environment.Value("DatasetType"))
If bReturn = False Then
    	Call Fn_UpdateLogFiles("["+Cstr(time) + "]  - Action - FAIL | Failed to Create Dataset [" + DataTable("DatasetName", dtGlobalSheet) + "]", "FAIL : Dataset Creation Failed")
	Call Fn_KillProcess(Environment.Value("KillProcesses"))
   	ExitTest
End If

Call Fn_UpdateLogFiles("["+Cstr(time) + "]  - Action - PASS | Successfully Created Dataset[" + DataTable("DatasetName", dtGlobalSheet) + "]", "")

'**********************************************************************************
'Select  the Created Dataset & Save As
'**********************************************************************************
Call Fn_MyTc_NavTree_NodeOperation("Select",GBL_AUTOMATEDTEST_FOLDER_PATH +":" + DataTable("DatasetName", dtGlobalSheet)  ,"")
bReturn = Fn_MyTc_DatasetSaveAs(DataTable("DatasetName", dtGlobalSheet) , "" , "1", "")
If bReturn = False Then
    	Call Fn_UpdateLogFiles("["+Cstr(time) + "]  - Action - FAIL | Failed to Save As Dataset [" + DataTable("DatasetName", dtGlobalSheet) + "]", "FAIL : Dataset Creation Failed")
	Call Fn_KillProcess(Environment.Value("KillProcesses"))
    ExitTest
End If

Call Fn_UpdateLogFiles("["+Cstr(time) + "]  - Action - PASS | Successfully Save As Dataset [" + DataTable("DatasetName", dtGlobalSheet) + "]", "")

'**********************************************************************************
'Expand Newstuff Folder
''**********************************************************************************
Call Fn_MyTc_NavTree_NodeOperation("Expand",GBL_NEWSTUFF_FOLDER_PATH,"")
Call Fn_ReadyStatusSync(1)

'**********************************************************************************
'verify Dataset allows not more than 32 characters
'**********************************************************************************
bReturn= Fn_MyTc_NavTree_NodeOperation("Exist",GBL_NEWSTUFF_FOLDER_PATH +":" + DataTable("DatasetName", dtGlobalSheet) , "")
If bReturn = False Then
    	Call Fn_UpdateLogFiles("["+Cstr(time) + "]  - Verify - VP4 - FAIL | Failed to Verify that the name field should not allowed more than 32 character", "FAIL : Failed to Verify the name field.")
	Call Fn_KillProcess(Environment.Value("KillProcesses"))
   	ExitTest
End If
Call Fn_UpdateLogFiles("["+Cstr(time) + "]  - Verify - VP4 - PASS | Successfully Verified  that the name field should not allowed more than 32 character.", "")

Call Fn_CodeCover_Exit()
Call Fn_SetTCSession("AutoTestDBA")

'*********************************************************************************
''Verify the preference 
'*********************************************************************************
bReturn = Fn_PreferenceOperations("VerifyPreference","TC_Allow_Longer_ID_Name","","","","","","","","","","")
If bReturn = False Then
	Call Fn_UpdateLogFiles("[" + Cstr(time) + "] - ACTION - FAIL | Failed to Verify the preference TC_Allow_Longer_ID_Name", "Fail : Fail to Verify Preference.")
Else
	Call Fn_UpdateLogFiles("[" + Cstr(time) + "] - ACTION - PASS | Successfully Verified the Preference TC_Allow_Longer_ID_Name.", "")
End If


'*********************************************************************************
''Set the preference
'*********************************************************************************
bReturn= Fn_Preference_Search_Operation("Modify","TC_Allow_Longer_ID_Name","True")
If bReturn = False Then
	Call Fn_UpdateLogFiles("[" + Cstr(time) + "] - ACTION - FAIL | Failed to set the TC_Allow_Longer_ID_Name preference", "FAIL : Failed to set the TC_Allow_Longer_ID_name preference")
    	Call Fn_KillProcess(Environment.Value("KillProcesses"))
    	ExitTest
Else
	Call Fn_UpdateLogFiles("[" + Cstr(time) + "] - ACTION - PASS | Successfully Set the TC_Allow_Longer_ID_Name Preference to True.", "")
End If

	
Call Fn_SetTCSession("AutoTest1")

'**********************************************************************************
'Select  the Created Dataset & Save As
'**********************************************************************************
Call Fn_MyTc_NavTree_NodeOperation("Select",GBL_AUTOMATEDTEST_FOLDER_PATH +":" + DataTable("DatasetName", dtGlobalSheet)  ,"")
sNewDataset= Fn_RandomString("DataSet",128)
bReturn = Fn_MyTc_DatasetSaveAs(sNewDataset , "" , "1", "")
If bReturn = False Then
    	Call Fn_UpdateLogFiles("["+Cstr(time) + "]  - Action - FAIL | Failed to Save As Dataset [" + sNewDataset + "]", "FAIL : Dataset Creation Failed")
	Call Fn_UpdateLogFiles("["+Cstr(time) + "]  - Verify - VP5 - FAIL | Failed to Verify that the name field should not allowed morethan 128 character.", "FAIL : VP5-Failed to Verify the name field.")
	Call Fn_KillProcess(Environment.Value("KillProcesses"))
    	ExitTest
End If

Call Fn_UpdateLogFiles("["+Cstr(time) + "]  - Action - PASS | Successfully Save As Dataset [" + sNewDataset + "]", "")
Call Fn_UpdateLogFiles("["+Cstr(time) + "]  - Verify - VP5 - PASS | Successfully Verified that the name field should not allowed more than 128 character..", "")



'**********************************************************************************
Call Fn_Setup_TestCaseExit(True)
