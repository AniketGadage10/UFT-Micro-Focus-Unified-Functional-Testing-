'#######################################################################################
'//###    TESTCASE NAME   :   DataSetSaveAs
'//###
'//###    DESCRIPTION     :   	Copy/Add based on an existing Dataset object
'//###
'//###		QART Link	8.3	    :		http://cipgweb/qacgi-bin/tt_view.cgi?release=Tc83_Aut&feature=REGM&cobid=132231&tcobid=938276
'//###
'//###		QART Link	9.0		:		http://cipgweb/qacgi-bin/tt_view.cgi?release=Tc_9.0&feature=REGMN&cobid=146025&tcobid=1046679
'//###
'//###    HISTORY        		 :   		AUTHOR              		DATE        				VERSION
'//###
'//###    CREATED BY      :   			Rajeev			    		12-May-2010		   			1.0
'//###
'//###    REVIWED BY      :				Sameer						15-05-2010	
'//###
'//###    MODIFIED BY     :				Rajeev					13-July_2010
'//###
'//###    REVIWED BY     :				Mohit					13-July_2010
'//###
'//###    Run on Tc Build  :
'//###
'//###    PORTED  BY	:			   Ketan 				20-12-2010			(Server::PNV6S166, 	Build::2010120100)
'//######################################################################################
'//######################################################################################
'Porting Details[Teamcenter 9. 1]
'QART Link 9.1 :	http://cipgweb/qacgi-bin/tt_view.cgi?release=TC_9.1&feature=REGMN&cobid=168845&tcobid=1241606
  'QART  Version :0001
  'BUILD  :2011071300(pnv6s166)
  'Porting Done By :Saurabh Thakur	    Date : 28-07-2011
  'Porting Reviewed By : Ketan Raje
  'Porting Comments: 1.Modified the Script to Select Dataset under New Stuff folder
'//######################################################################################
'Porting Details[Teamcenter 11.2]
'QART Link 11.2:	https://tidev.industrysoftware.automation.siemens.com/qart/qacgi-bin/tt_view.cgi?release=TC_11.1&feature=REGMN&category=Datasets&testcase=DatasetSaveAs
  'BUILD  :2015021800(pnv6s222)
  'Porting Done By :Jotiba Takkekar	    Date : 16-March-2015
   'Porting Comments: 1.Modified the Script to handle AutoTestDBA session to set preferences. 
 '                                        2. Removed Unused code. 
'##################################################################################

Option Explicit
'----------------------------------------------------------------------------------
'Variable Declaration
'----------------------------------------------------------------------------------
Dim bReturn
Dim sDefaultObj

sPrefName_Reset = "TC_Allow_Longer_ID_Name"
sPreVal_Reset = "true"
sScope_Reset = "site"
bPrefReset = true

Call fn_setup_TestCaseinit()

''**********************************************************************************
'Call Fn_UpdateLogFiles(Cstr(time) + " ***************************** QTP Action1 - Start **********************************", "")

'**********************************************************************************
'Look for Existing Tc Session for TcUserDBA from EnvVar_Ext.xml file
''**********************************************************************************
bPrefReset = false
bReturn = Fn_ReUserTcSession(true, true, Environment.Value("TcUser3") )
bPrefReset = true
If bReturn = False Then
	Call Fn_UpdateLogFiles("["+Cstr(time) + "] - Action - Fail  | Failed to Find Tc Session for User [" + Environment.Value("TcUser3") + "]", "FAIL:Tc Session Not Found")
    	Call Fn_KillProcess(Environment.Value("KillProcesses"))
    	ExitTest
Else

C:\Users\agadage\Desktop\TeamCenterApplication\TestData\TeamCenterLogin.xlsx

	Call Fn_UpdateLogFiles("[" + Cstr(time) + "] - Action - Pass | Successfully Found Tc Session for User [" + Environment.Value("TcUser3") + "]", "")
End If
Call Fn_ReadyStatusSync(1)

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
Call Fn_ReadyStatusSync(1)

'''**********************************************************************************
'Reset MyTc Perspective to Display Default state
'**********************************************************************************
Call Fn_ResetPerspective()
Call Fn_ReadyStatusSync(1)


'**********************************************************************************
'Action 1 Execution End
'**********************************************************************************
 Call Fn_UpdateLogFiles(Cstr(time) + " ***************************** QTP Action1 - End **********************************", "")
