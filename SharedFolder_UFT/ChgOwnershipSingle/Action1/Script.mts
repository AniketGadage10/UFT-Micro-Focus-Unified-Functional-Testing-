'#######################################################################################
'//###    TESTCASE NAME   :   ChgOwnershipSingle
'//###
'//###    DESCRIPTION         :  Verify you can Change Ownerships for Single Objects
'//###
'//###		QART Link			   :	http://cipgweb/qacgi-bin/tt_view.cgi?release=Tc_9.0&feature=REGMN&cobid=146005&tcobid=1046548
'//###
'//###    HISTORY        		  :   		AUTHOR              		DATE        				VERSION
'//###
'//###    CREATED BY       :   		  Saurabh		    		    9-Jan-2011		   			   1.0
'//###
'//###    REVIWED BY       :		  Harshal
'//###
'//###    Maintenance  BY     :		Rima Patil						23-July-2012	 
'//###
'//###    Run on Tc Build  :		 Run on Server 110 TC Build 9.0-	2011011900
'//######################################################################################
'//###   															Porting Details
'#######################################################################################
'//###    TESTCASE NAME   :   ChgOwnershipSingle
'//###
'//###    DESCRIPTION         :  Verify you can Change Ownerships for Single Objects
'//###
'//###		QART Link			   :	http://cipgweb/qacgi-bin/tt_view.cgi?release=Tc_9.0&feature=REGMN&cobid=146005&tcobid=1046548
'//###
'//###    Ported  BY     :		        Pooja B					26-July-2012	 
'//###
'//###    Run on Tc Build  :		 Run on Server 166  Build  2012071800
'//######################################################################################
'#######################################################################################
'//###    TESTCASE NAME   :   ChgOwnershipSingle
'//###
'//###    DESCRIPTION         :  Verify you can Change Ownerships for Single Objects
'//###
'//###		QART Link			   :	https://tidev.industrysoftware.automation.siemens.com/qart/qacgi-bin/tt_view.cgi?release=TC_11.1&feature=REGMN&category=ObjectProtection&testcase=ChgOwnershipSingle
'//###
'//###    Ported  BY     :		        Ankit Nigam 					18-Mar-2015	 
'//###
'//###    Run on Tc Build  :		 Run on Server 222  Build  2015021800
'//######################################################################################


Option Explicit
	'**********************************************************************************
	'Variable Declaration
	'**********************************************************************************
	Dim bReturn
	'**********************************************************************************
	'Initialize testcase
	'**********************************************************************************
	Call Fn_Setup_TestcaseInit()

	'**********************************************************************************
	'Look for Existing Tc Session for TcUser1 from EnvVar_Ext.xml file
	''**********************************************************************************
	bReturn = Fn_ReUserTcSession(False, True, Environment.Value("TcUser1"))
	If bReturn = False Then
		Call Fn_UpdateLogFiles("[" + Cstr(now) + "] - Action - Fail  | Failed to Find Tc Session for User [" + Environment.Value("TcUser1") + "]", "FAIL:Tc Session Not Found")
		Call Fn_KillProcess("")
		ExitTest
	Else
		Call Fn_UpdateLogFiles("[" + Cstr(now) + "] - Action - PASS | Successfully Found Tc Session for User [" + Environment.Value("TcUser1") + "]", "")
	End If
	
	'**********************************************************************************
	'Set the MyTeamcenter Perspective
	''*********************************************************************************
	bReturn = Fn_SetPerspective("My Teamcenter")
	If bReturn = False Then
		Call Fn_UpdateLogFiles("[" + Cstr(now) + "] - Action - Fail | Failed to Set My Teamcenter Perspective", "FAIL:Failed to Go to My Teamcenter Module")
		Call Fn_KillProcess("")
		ExitTest
	Else
		Call Fn_ReadyStatusSync(1)
		Call Fn_UpdateLogFiles("[" + Cstr(now) + "] - Action - PASS | Successfully set My Teamcenter Perspective", "")
	End If
	 
	'**********************************************************************************
	'Reset MyTc Perspective to Display Default state
	'**********************************************************************************
	bReturn = Fn_ResetPerspective()
	If bReturn = False Then
		Call Fn_UpdateLogFiles("[" + Cstr(now) + "] - Action - Fail | Failed to Reset My Teamcenter Perspective", "FAIL:Failed to Reset My Teamcenter Perspective")
		Call Fn_KillProcess("")
		ExitTest
	Else
		Call Fn_ReadyStatusSync(1)
		Call Fn_UpdateLogFiles("[" + Cstr(now) + "] - Action - PASS | Successfully Reset My Teamcenter Perspective", "")
	End If
	
	Call Fn_UpdateLogFiles("[" + Cstr(now) + "] - QTP Action 1 - End", "")
	Call Fn_UpdateLogFiles("-----------------------------------------------------------------------------------------------", "")
