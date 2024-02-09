Option Explicit
'**********************************************************************************
'Variables Declaration
'**********************************************************************************	
Dim bReturn
Dim bExitTest
Dim strXMLParentFolder , strXMLFileName, arrTemp

'Call Fn_Setup_TestcaseInit()
' IMP Variable : Environment.Value("bExitTest")  = true / false  : to terminate test execution
' IMP Variable : Environment.Value("XMLFile") = AM Rules XML File path

bExitTest = Environment.Value("bExitTest")
arrTemp = split(Environment.Value("XMLFile"), "\")
strXMLFileName = arrTemp(UBound(arrTemp))
strXMLParentFolder = Environment.Value("XMLFile")
strXMLParentFolder = replace(strXMLParentFolder , strXMLFileName,"")
strXMLParentFolder = left(strXMLParentFolder,len(strXMLParentFolder)-1)  

Call Fn_UpdateLogFiles("[" + Cstr(now) + "] - Recovery Scenario [ Importing default AM Rules XML ] - Start", "")
Call Fn_UpdateLogFiles("", "")
'**********************************************************************************
'Login with infodba User
'**********************************************************************************
bReturn = Fn_ReUserTcSession(true, true, Environment.Value("TcUserDBA"))
If bReturn = False Then
	Call Fn_UpdateLogFiles("[" + Cstr(now) + "] - Recovery Scenario [ Importing default AM Rules XML ] - FAIL: Failed to Find Tc Session for User [" + Environment.Value("TcUserDBA") + "]", "")
	ExitTest
Else
	Call Fn_UpdateLogFiles("[" + Cstr(now) + "] - Recovery Scenario [ Importing default AM Rules XML ] - PASS: Successfully Found Tc Session for User [" + Environment.Value("TcUserDBA") + "]", "")
End If
Call Fn_ReadyStatusSync(1)
bReturn = Fn_SetPerspective("Access Manager")
If bReturn = False Then
	Call Fn_UpdateLogFiles("[" + Cstr(now) + "] - Recovery Scenario [ Importing default AM Rules XML ] - FAIL: Failed to Set Access Manager Perspective", "")
    Call Fn_KillProcess(Environment.Value("KillProcesses"))
    ExitTest
Else
	Call Fn_UpdateLogFiles("[" + Cstr(now) + "] - Recovery Scenario [ Importing default AM Rules XML ] - PASS: Successfully Set Access Manager Perspective.", "")
End If
Call Fn_ReadyStatusSync(1)

bReturn=Fn_AccMgr_ImportExportAMRules("Import_WithOutDelete", strXMLParentFolder , strXMLFileName)
If bReturn = False Then
	Call Fn_UpdateLogFiles("[" + Cstr(now) + "] - Recovery Scenario [ Importing default AM Rules XML ] - FAIL: Failed to imported AM Rules XML", "")
	Call Fn_KillProcess(Environment.Value("KillProcesses"))
	ExitTest
Else
	Call Fn_UpdateLogFiles("[" + Cstr(now) + "] - Recovery Scenario [ Importing default AM Rules XML ] - PASS: Successfully imported AM Rules XML", "")
End If
Call Fn_ReadyStatusSync(1)

Call Fn_TeamcenterExit()
'Call Fn_KillProcess(Environment.Value("KillProcesses"))
Call Fn_UpdateLogFiles("", "")
Call Fn_UpdateLogFiles("[" + Cstr(now) + "] - Recovery Scenario [ Importing default AM Rules XML ] - END", "")
Call Fn_UpdateLogFiles("", "")

If bExitTest = true Then
	ExitTest	
End If

