
'#######################################################################################
'//###    TESTCASE NAME   :   CreateFormDataset 
'//###
'//###    DESCRIPTION     :   	Create Form on a Dataset. This test also verifies that if "TC_Allow_Longer_ID_Name" perference is set to false and an attempt is made to create a form with > 32 characters that appropriate feedback is received and the form is not created.
'//###
'//###		QART Link			:		http://cipgweb/qacgi-bin/tt_view.cgi?release=Tc83_Aut&feature=REGM&cobid=132237&tcobid=938328  
'//###
'//###    HISTORY        		 :   		AUTHOR              		DATE        				VERSION
'//###
'//###    CREATED BY      :   			Rizwan  		    		30-Apr-2010		   			1.0
'//###
'//######################################################################################

'#######################################################################################
'//###    TESTCASE NAME   :   CreateFormDataset 
'//###
'//###    MAINTAIN BY      :	Swapna Ghatge 			14th-Dec-2011						
'//###
'//###    Run on Tc Build  : Pnv6s169 on Build 1207
'//###
'//######################################################################################
'#######################################################################################
'//###
'//###    QART Link      :   http://cipgweb/qacgi-bin/tt_view.cgi?release=TC_10.1&feature=REGMN&cobid=214321&tcobid=1659208
'//###
'//###    HISTORY        :     AUTHOR                DATE            VERSION      
'//###
'//###    MODIFIED BY     :      Sonal P    	 14-Jan-2013            0001
'//###
'//###    Run on Tc Build  :   pnv6s203\2012121200
'//#####################################################################################

'#######################################################################################
'//###
'//###    QART Link           :   http://cipgweb/qacgi-bin/tt_view.cgi?release=TC_10.1&feature=REGMN&cobid=214321&tcobid=1659208
'//###
'//###    HISTORY            :     AUTHOR                         DATE            
'//###
'//###    MODIFIED BY    :    Preeti Shendre    	 23-Sep-2013       
'//###
'//###    Run on Tc Build   :   pnv6s222\2013060400A(Patch 902)
'//### 
'//###    Changes               :   1) Removed unused variables.
'//###									     2) Used XML for Menu and Toolbar operation.
'//###		                                 3) Changed Log. 
'//###    
'//######################################################################################
Option Explicit

'**********************************************************************************
'Variable Declaration
'**********************************************************************************
Dim bReturn
Dim sPrefName_Reset,sPreVal_Reset,sScope_Reset,aSOAPerfInput
Dim bPrefReset,aUserDetails,sPerfName,sPerfValue,sPerfScope

Call Fn_Setup_TestcaseInit()

'**********************************************************************************
' Re-Setting Preference Values
'**********************************************************************************
sPrefName_Reset = "TC_Allow_Longer_ID_Name"
sPreVal_Reset = "True"
sScope_Reset = "site"
bPrefReset  = True

'**********************************************************************************
' Check if Preference Exists... if not,, then Create it and set the value
'**********************************************************************************
aUserDetails = split(Environment.Value("TcUserDBA"), ":")
sPerfName = "TC_Allow_Longer_ID_Name"
sPerfValue = "False"
sPerfScope = "site"

aSOAPerfInput = Array(aUserDetails(0) , "createpreference", sPerfName, sPerfScope,sPerfValue)
bReturn = Fn_SOA_PrefOperation(aSOAPerfInput)
If bReturn = False Then
		Call Fn_UpdateLogFiles("[" + Cstr(time)  + "] - ACTION - PASS | Preference - ["+sPerfName+"] already Exists in TeamCenter", "")

		'  NowModify the pref value
		bReturn = Fn_SOA_SetPreference(aUserDetails(0), sPerfName, sPerfValue,sPerfScope)
		If bReturn = False Then
			Call Fn_UpdateLogFiles("[" + Cstr(time)  + "] - ACTION - FAIL | Failed to Set the site-preference ["+sPerfName+"] to ["+sPerfValue+"]", "FAIL:Failed to Set the site-preference 'maintenance'")
			ExitTest
		Else
			Call Fn_UpdateLogFiles("[" + Cstr(time)  + "] - ACTION - PASS | Successfully Set site Preference ["+sPerfName+"] to ["+sPerfValue+"]", "")
		End If
Else
		Call Fn_UpdateLogFiles("[" +Cstr(time)  + "] - ACTION - PASS | Successfully Created Preference ["+sPerfName+"] with Value ["+sPerfValue+"]", "")
End If
Wait 5
'**********************************************************************************
'Look for Existing Tc Session for TcUser2 from EnvVar_Ext.xml file
''**********************************************************************************
bPrefReset  = False
bReturn = Fn_ReUserTcSession(True, True, Environment.Value("TcUser2"))
bPrefReset  = True
If bReturn = False Then
	Call Fn_UpdateLogFiles("[" + Cstr(time) + "] - Action - Fail | Failed to Find Tc Session for User [" + Environment.Value("TcUser2") + "]", "FAIL:Tc Session Not Found")
    Call Fn_KillProcess(Environment.Value("KillProcesses"))
	ExitTest
Else
	Call Fn_UpdateLogFiles("[" + Cstr(time) + "] - Action - Pass | Successfully Found Tc Session for User [" + Environment.Value("TcUser2") + "]", "")
End If
Call Fn_ReadyStatusSync(1)
'**********************************************************************************
'Set the MyTeamcenter Perspective
''**********************************************************************************
bReturn = Fn_SetPerspective("My Teamcenter")
If bReturn = False Then
	Call Fn_UpdateLogFiles("[" + Cstr(time) + "] - Action - Fail | Failed to Set My Teamcenter Perspective", "FAIL:Failed to Go to MyTc Module")
    Call Fn_KillProcess(Environment.Value("KillProcesses"))
	ExitTest
Else
	Call Fn_UpdateLogFiles("[" + Cstr(time) + "] - Action - Pass | Successfully set My Teamcenter Perspective", "")
End If
Call Fn_ReadyStatusSync(1)
Call Fn_UpdateLogFiles("[" + Cstr(time) + "] - QTP Action 1 - End", "")
