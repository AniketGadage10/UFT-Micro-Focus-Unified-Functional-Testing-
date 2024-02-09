'############################################## InitialLoginSetup_Web ##############################################
'//###    HISTORY        		 :   		AUTHOR              		DATE        				VERSION
'//###
'//###    CREATED BY      :   			Sunny R						13-04-2011	   			1.0
'//###
'//###    REVIWED BY      :	
'//### 
'//###    PARAMETERS   :			Username,Password
'//###	  EXAMPLE				:					
'//###														sUserName=Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\Web_Users.xml", "webUser02")
'//###														aUserDetails=Split(sUserName,":")
'//###													    RunAction "InitialLoginSetup_Web [GeneralSetup]", oneIteration, aUserDetails(0),aUserDetails(1)
'//###
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


''''''**********************************************************************************
''''' 'Login to Web Client
'''''**********************************************************************************
bReturn =Fn_Web_Login(Parameter("UserName"),Parameter("Password"))
If bReturn = True Then  
                Call Fn_UpdateLogFiles(Time() & " - " & "Action - PASS | Successfully Loged into "+Environment.Value("WebBrowserName")+" with User [ "+Parameter("UserName")+" ]", "")
Else
                Call Fn_UpdateLogFiles(Time() & " - " & "Action - FAIL | Failed To Log into "+Environment.Value("WebBrowserName")+" with User [ "+Parameter("UserName")+" ]", "FAIL: Failed To Log into "+Environment.Value("WebBrowserName")+" with User [ "+Parameter("UserName")+" ]")
                Call Fn_Web_KillProcess("")
                ExitTest
End If

'**********************************************************************************
'Action 1 Execution End
''**********************************************************************************
Call Fn_UpdateLogFiles("******************* Reusable Action - End*********************", "")



