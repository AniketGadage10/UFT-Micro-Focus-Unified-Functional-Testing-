Option Explicit
		
'**********************************************************************************
'Variable Declaration
'**********************************************************************************
Dim bReturn, iRanNo, sTestFolderName ,sFileName,sErrorMsg,dicErrorInfo,objNewForm,sPathName
Dim aUserDetails,sPerfName,sPerfValue,sPerfScope

'**********************************************************************************
'Action 2 Execution
'**********************************************************************************
Call Fn_UpdateLogFiles("-----------------------------------------------------------------------------------------------", "")
Call Fn_UpdateLogFiles("[" + Cstr(time) + "] - QTP Action 2 - Start", "")

'**********************************************************************************************
'Check the Existance of AutomatedTest folder under Home & Create if not exist
''**********************************************************************************************
bReturn = Fn_MyTc_NavTree_NodeOperation("Exist","Home:AutomatedTests","")
If bReturn = false Then
	Call Fn_MyTc_NavTree_NodeOperation("Select","Home","")
	bReturn = Fn_MyTc_FolderCreate("Folder","AutomatedTests","Automation Artifact","OFF")
		If bReturn = False Then
			Call Fn_UpdateLogFiles("[" + Cstr(time) + "] - Action - Fail | Failed to Create AutomatedTests Folder under Home in NavTree", "FAIL:AutomatedTests Folder not Created")
			Call Fn_KillProcess(Environment.Value("KillProcesses"))
			ExitTest
		End If
Call Fn_ReadyStatusSync(3)
Call Fn_UpdateLogFiles("[" + Cstr(time) + "] - Action - Pass | Successfully Created AutomatedTests Folder under Home", "")
End If
		
'**********************************************************************************
'Expand and Select AutomatedTests Folder
''**********************************************************************************
bReturn= Fn_MyTc_NavTree_NodeOperation("Expand","Home:AutomatedTests","")
If bReturn = False Then
	Call Fn_UpdateLogFiles("[" + Cstr(time) + "] - Action - Fail | Failed to Expand [Home:AutomatedTests] folder", "FAIL:Failed to Expand [Home:AutomatedTests] folder")
	Call Fn_KillProcess(Environment.Value("KillProcesses"))
	ExitTest
Else
	Call Fn_UpdateLogFiles("[" + Cstr(time) + "] - Action - Pass | Successfully Expanded  [Home:AutomatedTests] folder", "")
End If
Call Fn_ReadyStatusSync(3)

bReturn= Fn_MyTc_NavTree_NodeOperation("Select","Home:AutomatedTests","")
If bReturn = False Then
	Call Fn_UpdateLogFiles("[" + Cstr(time) + "] - Action - Fail | Failed to Select [Home:AutomatedTests] folder", "FAIL:Failed to Select [Home:AutomatedTests] folder")
	Call Fn_KillProcess(Environment.Value("KillProcesses"))
	ExitTest
Else
	Call Fn_UpdateLogFiles("[" + Cstr(time) + "] - Action - Pass | Successfully Selected  [Home:AutomatedTests] folder", "")
End If
Call Fn_ReadyStatusSync(3)
'**********************************************************************************
'Create TestCase Folder under AutomatedTests folder
''**********************************************************************************
iRanNo = Fn_RandNoGenerate()
sTestFolderName = Environment.Value("TestName") + "_" + Cstr(iRanNo)
bReturn = Fn_MyTc_FolderCreate("Folder",sTestFolderName,"Automation Artifact - Test Case Folder","OFF")
If bReturn = False Then
	Call Fn_UpdateLogFiles("[" + Cstr(time) + "] - Action - Fail | Failed to Create Test case Folder [" + sTestFolderName + "]", "FAIL:Test Case Folder not Created")
	Call Fn_KillProcess(Environment.Value("KillProcesses"))
	ExitTest
End If
Call Fn_ReadyStatusSync(3)
Call Fn_UpdateLogFiles("[" + Cstr(time) + "] - Action - Pass | Successfully Created Test case Folder [" + sTestFolderName + "]", "")

'If Test Case Folder Path is Required to be Refered Across the Actions, Store it
Environment.Value("TestFolderName") = "Home:AutomatedTests:" + sTestFolderName
		
'**********************************************************************************
'Expand and Select Test Case Folder
''**********************************************************************************
bReturn =Fn_MyTc_NavTree_NodeOperation("Expand",Environment.Value("TestFolderName"),"")
If bReturn = False Then
	Call Fn_UpdateLogFiles("[" + Cstr(time) + "] - Action - Fail | Failed to Expand ["+Environment.Value("TestFolderName")+"]", "FAIL:Failed to Expand ["+Environment.Value("TestFolderName")+"]")
	Call Fn_KillProcess(Environment.Value("KillProcesses"))
	ExitTest
Else
	Call Fn_UpdateLogFiles("[" + Cstr(time) + "] - Action - Pass | Successfully Expanded  ["+Environment.Value("TestFolderName")+"]", "")
End If
Call Fn_ReadyStatusSync(3)

bReturn= Fn_MyTc_NavTree_NodeOperation("Select",Environment.Value("TestFolderName"),"")
If bReturn = False Then
	Call Fn_UpdateLogFiles("[" + Cstr(time) + "] - Action - Fail | Failed to Select ["+Environment.Value("TestFolderName")+"]", "FAIL:Failed to Select ["+Environment.Value("TestFolderName")+"]")
	Call Fn_KillProcess(Environment.Value("KillProcesses"))
	ExitTest
Else
	Call Fn_UpdateLogFiles("[" + Cstr(time) + "] - Action - Pass | Successfully Selected  ["+Environment.Value("TestFolderName")+"]", "")
End If
Call Fn_ReadyStatusSync(3)
'**************************************************************************************************
'Create a New Dataset on the selected Folder
'**************************************************************************************************
bReturn = Fn_DatasetCreate(DataTable("DatasetName", dtGlobalSheet), DataTable("DatasetDesc", dtGlobalSheet), DataTable("DatasetTool", dtGlobalSheet), DataTable("FilePath", dtGlobalSheet), DataTable("OpenOnCreate", dtGlobalSheet), DataTable("DatasetType", dtGlobalSheet))
If bReturn = False Then
	Call Fn_UpdateLogFiles("[" + Cstr(time) + "] - Action - Fail | Failed to Create New Dataset", "FAIL:Failed to Select the Created Folder")
	Call Fn_KillProcess(Environment.Value("KillProcesses"))
	ExitTest
Else
	Call Fn_UpdateLogFiles("[" + Cstr(time) + "] - Action - Pass | Successfully Created New Datset [ "+DataTable("DatasetName", dtGlobalSheet)+" ]", "")
End If
Call Fn_ReadyStatusSync(1)

'**********************************************************************************
'Select the newly created Dataset
''**********************************************************************************
bReturn = Fn_MyTc_NavTree_NodeOperation("Select",Environment.Value("TestFolderName") + ":" + Datatable("DatasetName", dtGlobalSheet),"")
If bReturn = False Then
	Call Fn_UpdateLogFiles("[" + Cstr(time) + "] - Action - Fail | Failed to Select  Dataset", "FAIL:Failed to Select  Dataset")
	Call Fn_KillProcess(Environment.Value("KillProcesses"))
	ExitTest
End If
Call Fn_ReadyStatusSync(3)
Call Fn_UpdateLogFiles("[" + Cstr(time) + "] - Action - Pass | Successfully Select New Datset [ "+DataTable("DatasetName", dtGlobalSheet)+" ]", "")
'**********************************************************************************
'Create New Form under the newly created Dataset
''**********************************************************************************
bReturn = Fn_FormCreate(Datatable("FormName", dtGlobalSheet), DataTable("FormDesc", dtGlobalSheet), DataTable("FormType", dtGlobalSheet),DataTable("OpenOnCreate", dtGlobalSheet))
If bReturn = False Then
	Call Fn_UpdateLogFiles("[" + Cstr(time) + "] - Action - Fail | Failed to Create Form [" + DataTable("FormName", dtGlobalSheet) + "]", "FAIL: Form Creation Failed")
	Call Fn_KillProcess(Environment.Value("KillProcesses"))
	ExitTest
Else
	Call Fn_UpdateLogFiles("[" + Cstr(time) + "] - Action - Pass | Successfully Created Form [" + DataTable("FormName", dtGlobalSheet) + "]", "")
End If
Call Fn_ReadyStatusSync(3)

''**********************************************************************************
''Expand the newly created Dataset
'''**********************************************************************************
bReturn = Fn_MyTc_NavTree_NodeOperation("Expand",Environment.Value("TestFolderName") + ":" + Datatable("DatasetName", dtGlobalSheet) ,"")
If bReturn = False Then
	Call Fn_UpdateLogFiles("[" + Cstr(time) + "] - Action - Fail | Failed to Expand  Dataset", "FAIL:Failed to Expand  Dataset")
	Call Fn_KillProcess(Environment.Value("KillProcesses"))
	ExitTest
End If
Call Fn_ReadyStatusSync(3)
Call Fn_UpdateLogFiles("[" + Cstr(time) + "] - Action - Pass | Successfully Expand  Dataset [" + DataTable("DatasetName", dtGlobalSheet) + "]", "")
'**********************************************************************************
'Verify if the newly created form is successfully created
''**********************************************************************************
If Fn_MyTc_NavTree_NodeOperation("Exist",Environment.Value("TestFolderName") + ":" + Datatable("DatasetName", dtGlobalSheet) + ":" + Datatable("FormName", dtGlobalSheet),"") = True Then
	Call Fn_UpdateLogFiles("[" + Cstr(time) + "] - Verify - Pass | VP1 - Successfully verified the form [ "+Datatable("FormName", dtGlobalSheet)+" ] created by expanding the dataset [ "+Datatable("DatasetName", dtGlobalSheet)+" ] ", "")
Else
		Call Fn_UpdateLogFiles("[" + Cstr(time) + "] - Verify - Fail | VP1 - Fail to Verify the form created by expanding the dataset [ "+Datatable("FormName", dtGlobalSheet)+" ]", "")
		Call Fn_KillProcess(Environment.Value("KillProcesses"))
		ExitTest
End If

'*********************************************************************************
'Select the newly created FOrm
''**********************************************************************************
bReturn = Fn_MyTc_NavTree_NodeOperation("Select",Environment.Value("TestFolderName") + ":" + Datatable("DatasetName", dtGlobalSheet),"")
 If bReturn = False Then
		Call Fn_UpdateLogFiles("[" + Cstr(time) + "] - Action - Fail | Failed to Select  Form", "FAIL:Failed to Select  Form")
		Call Fn_KillProcess(Environment.Value("KillProcesses"))
		ExitTest
End If
Call Fn_ReadyStatusSync(3)
Call Fn_UpdateLogFiles("[" + Cstr(time) + "] - Action - Pass | Successfully Select  Form [" + DataTable("FormName", dtGlobalSheet) + "]", "")
'**********************************************************************************
'Create New Form (More than 32 chara length) under the newly created Dataset 
''**********************************************************************************
bReturn = Fn_FormCreate(Datatable("LongFormName", dtGlobalSheet), DataTable("FormDesc", dtGlobalSheet), DataTable("FormType", dtGlobalSheet),DataTable("OpenOnCreate", dtGlobalSheet))
If bReturn = False Then
	Call Fn_UpdateLogFiles("[" + Cstr(time) + "] - Action - Fail | Failed to Create Form [" + DataTable("FormName", dtGlobalSheet) + "]", "FAIL: Form Creation Failed")
	Call Fn_KillProcess(Environment.Value("KillProcesses"))
	ExitTest
Else
	Call Fn_UpdateLogFiles("[" + Cstr(time) + "] - Action - Pass | Successfully Peform Menu operation [ File:New:Form...]", "")
	Call Fn_UpdateLogFiles("[" + Cstr(time) + "] - Action - Pass | Successfully Set name [ "+DataTable("LongFormName", dtGlobalSheet)+" ] which exceeds the allowed length of 32.", "")
	Call Fn_UpdateLogFiles("[" + Cstr(time) + "] - Action - Pass | Successfully Set Description [ "+DataTable("FormDesc", dtGlobalSheet)+" ] .", "")
	Call Fn_UpdateLogFiles("[" + Cstr(time) + "] - Action - Pass | Successfully Clicked On Finish.", "")
	Call Fn_UpdateLogFiles("[" + Cstr(time) + "] - Action - Pass | Successfully Created Form [" + DataTable("FormName", dtGlobalSheet) + "]", "")
End If
Call Fn_ReadyStatusSync(3)

'**********************************************************************************
'Verify the character length of the last created form. 
''**********************************************************************************
Datatable("LongFormName_32", dtGlobalSheet) = Left (Datatable("LongFormName", dtGlobalSheet),32)
bReturn = Fn_MyTc_NavTree_NodeOperation("Exist",Environment.Value("TestFolderName") + ":" + Datatable("DatasetName", dtGlobalSheet) + ":" + Datatable("LongFormName_32", dtGlobalSheet),"")
If bReturn = False Then
	Call Fn_UpdateLogFiles("[" + Cstr(time) + "] - Verify - Fail | VP2 - Fail to verify Form created with the first 32 character when Tc_Allow_Longer_ID_Name is set to False", "Fail : VP2 - Fail to verify Form created with the first 32 Characters when Tc_Allow_Longer_ID_Name is set to False")
	Call Fn_KillProcess(Environment.Value("KillProcesses"))
	ExitTest
Else
	Call Fn_UpdateLogFiles("[" + Cstr(time) + "] - Verify - Pass | VP2.1 - Successfully Verified that When the preference [ Tc_Allow_Longer_ID_Name ] is set to [ false ] then the user is allowed to Create an Item [ "+Datatable("LongFormName", dtGlobalSheet)+ "] of Type [ "+ Datatable("DatasetName", dtGlobalSheet)+" ] having more than 32 Characters in its Name", "")
	Call Fn_UpdateLogFiles("[" + Cstr(time) + "] - Verify - Pass | VP2.2 - Successfully Verified that When the preference [ Tc_Allow_Longer_ID_Name ] is set to [ false ] then the user is allowed to enter more than 32 Characters  but will dispay only first 32 Characters in Items Name,", "")
End If

'''*********************************************************************************************
''Close the  Perspective
'''*********************************************************************************************
sFileName  = Fn_LogUtil_GetXMLPath("RAC_Menu")
sPathName = Fn_GetXMLNodeValue(sFileName, "FileClose")
bReturn= Fn_MenuOperation("Select",sPathName)
If bReturn=false Then
		Call Fn_UpdateLogFiles("[" + Cstr(time) + "] - Action - Fail | Perspective was not closed","FAIL:Perspective was not closed")
		Call Fn_KillProcess(Environment.Value("KillProcesses"))
         ExitTest
else
		Call Fn_UpdateLogFiles("[" + Cstr(time) + "] - Action - Pass | My Teamcenter Perspective successfully closed","")
End If

'Exit from teamcenter
Call Fn_TeamcenterExit()

'**********************************************************************************
' Check if Preference Exists... if not,, then Create it and set the value
'**********************************************************************************
aUserDetails = split(Environment.Value("TcUserDBA"), ":")
sPerfName = "TC_Allow_Longer_ID_Name"
sPerfValue = "True"
sPerfScope = "site"

'**********************************************************************************
' set Preference with its default value
'**********************************************************************************
bReturn = Fn_SOA_SetPreference("AutoTestDBA", sPerfName, sPerfValue,sPerfScope)
If bReturn = False Then
	Call Fn_UpdateLogFiles("[" + Cstr(now) + "] - ACTION - FAIL | Failed to Set the site-preference ["+sPerfName+"] to ["+sPerfValue+"]", "FAIL:Failed to Set the site-preference 'maintenance'")
	ExitTest
Else
	Call Fn_UpdateLogFiles("[" + Cstr(now) + "] - ACTION - PASS | Successfully Set site Preference ["+sPerfName+"] to default value ["+sPerfValue+"]", "")
End If

'**********************************************************************************
'Action 2 Execution End
'**********************************************************************************
Call Fn_UpdateLogFiles("---------------------------------------------------------------------------------------", "")
Call Fn_UpdateLogFiles("******************* Action 2 Execution : END*********************", "")
Call Fn_UpdateLogFiles("----------------------------------------------------------------------------------------", "")

'**********************************************************************************
'Log Test Result
''**********************************************************************************
Call Fn_UpdateLogFiles("Execution End Time :: " + cStr(now), "")
Call Fn_UpdateLogFiles("-------------------------------------", "")
Call Fn_UpdateLogFiles("Test Execution Result: PASS", "PASS: All VP Pass")
Call Fn_UpdateLogFiles("-------------------------------------", "")

'**********************************************************************************
'Call for Code Coverage
'**********************************************************************************
Call Fn_Setup_TestcaseExit(True)

