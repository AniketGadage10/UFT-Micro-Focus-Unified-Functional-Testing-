'############################################## InitialLoginSetup_RAC ##############################################
'//###    HISTORY        		 :   		AUTHOR              		DATE        				VERSION
'//###
'//###    CREATED BY      :   			Sunny R						11-04-2011	   			1.0
'//###
'//###    REVIWED BY      :	
'//### 
'//###    PARAMETERS   :			CacheClear,Relaunch,LoginCredentials,PerspectiveName,[OPTIONAL]SOAReturnvalue
'//###	  EXAMPLE				:					
'//###
'//###											Case 1: When SOA is not implemented
'//###													RunAction "InitialLoginSetup_RAC [GeneralSetup]", oneIteration, "False", "False", "TcUser1", "My Teamcenter"
'//###											Case 2: When SOA is Implemented
'//###													Step1: 	Call LoadEnvXML()
'//###													Step2: 	bReturn = Fn_SOA_SetPreference(sUserDetail, sPrefName, sPerfValue,sPerfScope)
'//###													Step3: 	RunAction "InitialLoginSetup_RAC [GeneralSetup]", oneIteration, "False", "False", "TcUser1", "My Teamcenter",bReturn
'//###													Step4: 	If bReturn = False Then
'//###																			Call Fn_UpdateLogFiles(Time() & " - " & "Action - Fail", "FAIL:Failed")
'//###																			ExitTest
'//###																	Else
'//###																			Call Fn_UpdateLogFiles(Time() & " - " & "Action - Pass", "")
'//###																	End If
'###################################################################################################

Option Explicit
'**********************************************************************************
'Variable Declaration
'**********************************************************************************
Dim bReturn, sLogFile
Dim Result
ReDim TestData(8)

''-----------------------------------------------------------------------------------------------------------
'' Import Environment Variable File
''-----------------------------------------------------------------------------------------------------------
Call LoadEnvXML()

'**********************************************************************************
'Run Action Only for First Row of DataTable
'**********************************************************************************
Call Fn_SetActionIterationMode("oneIteration")

'**********************************************************************************
'Call for Code Coverage
'**********************************************************************************
Call Fn_CodeCover_Init()

'**********************************************************************************
'Create Test Log File
'**********************************************************************************
sLogFile = Fn_CreateLogFile(Environment.Value("TestName") + ".log")
If sLogFile <> "" Then
	Environment.Value("TestLogFile") = sLogFile
Else
	Reporter.ReportEvent micFail, "LogFileCreate", "Failed to Create Test Log File"
	Call Fn_KillProcess("")
	ExitTest
End If

'**********************************************************************************
'Write QART Specific Info to log
'**********************************************************************************
Call Fn_PrintQARTLog()

'Write Test Details to Batch Log File
TestData(1) = Environment.Value("QARTRoot")
TestData(2) = Environment.Value("QARTRelease")
TestData(3) =DataTable("Feature", dtGlobalSheet)
TestData(4) = DataTable("Category", dtGlobalSheet)
TestData(5) =  Environment.Value("TestName")
TestData(6) = Date & " - " & Time
Result = Fn_Update_TestDetail( "", TestData, 1)

'**********************************************************************************
'Action 1 Execution Start
''**********************************************************************************
Call Fn_UpdateLogFiles("******************* Reusable Action - Start*********************", "")

'[Harshal]: in case SOA is implemented.
If Parameter("SOAReturnValue") = False Then
	ExitAction
End If

''''''**********************************************************************************
''''' 'Login to TC 
'''''**********************************************************************************
bReturn = Fn_ReUserTcSession(Parameter("CacheClear"),Parameter("Relaunch"), Environment.Value(Parameter("UserCredential")))
If bReturn = False Then
				Call Fn_UpdateLogFiles(Time() & " - " & "Action - Fail | Failed to Find Tc Session for User [" + Environment.Value(Parameter("UserCredential")) + "]", "FAIL:Failed to Find Tc Session for User [" + Environment.Value(Parameter("UserCredential")) + "]")
				 ExitTest
Else
				Call Fn_UpdateLogFiles(Time() & " - " & "Action - Pass | Successfully Found Tc Session for User [" + Environment.Value(Parameter("UserCredential")) + "]", "")
End If
				Call Fn_ReadyStatusSync(1)

''*************************************************************************************
''Maximize the window
''************************************************************************************
Call Fn_Window_Maximize(Environment.Value("TestName"),JavaWindow("DefaultWindow"))
Call Fn_UpdateLogFiles(Time() & " - " & "Action - Pass | Successfully maximized the window", "")

''*************************************************************************************
''Launch the Required Perspective
''************************************************************************************
bReturn = Fn_SetPerspective(Parameter("PerspectiveName"))
If bReturn = False Then
				Call Fn_UpdateLogFiles(Time() & " - " & "Action - Fail | Failed to Set "+Parameter("PerspectiveName")+" Perspective", "FAIL:Failed to Set "+Parameter("PerspectiveName")+" Perspective")
			    Call Fn_KillProcess(Environment.Value("KillProcesses"))
			    ExitTest
Else
				Call Fn_UpdateLogFiles(Time() & " - " & "Action - Pass | Successfully set "+Parameter("PerspectiveName")+" Perspective.", "")
End If
				Call Fn_ReadyStatusSync(1)

''*************************************************************************************
''Reset the Perspective
''************************************************************************************
Call Fn_ResetPerspective()

'**********************************************************************************
'Action 1 Execution End
''**********************************************************************************
Call Fn_UpdateLogFiles("******************* Reusable Action - End*********************", "")

