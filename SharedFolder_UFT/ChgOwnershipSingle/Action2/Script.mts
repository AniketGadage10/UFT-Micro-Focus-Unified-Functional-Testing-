Option Explicit
		
'**********************************************************************************
'Variable Declaration
'**********************************************************************************
Dim  bReturn,iRanNo,sTestFolderName,sFolderpath,sItemPath,sItemselPath,aItmInfo,ObjChng
Dim sGetText,sUser2,sUser1,sRootNode,sItemtosearch,bReturn2,sUserData,aUserInfo,sUserData1,aUserInfo1


Call Fn_UpdateLogFiles("-----------------------------------------------------------------------------------------------", "")
Call Fn_UpdateLogFiles("[" + Cstr(now) + "] - QTP Action 2 - Start", "")

sUserData = Environment.Value("TcUser2")
aUserInfo=split(sUserData,":",-1,1)
'**********************************************************************************
'Check the Existance of AutomatedTest folder under Home
''**********************************************************************************
bReturn = Fn_MyTc_NavTree_NodeOperation("Exist","Home:AutomatedTests","")
If bReturn = false Then
	bReturn = Fn_MyTc_FolderCreate("Folder","Automated_Tests","Automation Artifact","OFF")
	If bReturn = False Then
		Call Fn_UpdateLogFiles("[" + Cstr(now) + "]  -Action- FAIL |Failed to Create AutomatedTests Folder under Home in NavTree", "FAIL:AutomatedTests Folder not Created")
		Call Fn_KillProcess("")
		ExitTest
	End If
	Call Fn_UpdateLogFiles("[" + Cstr(now) + "]  -Action- PASS | Successfully Created AutomatedTests Folder under Home", "")
	Call Fn_ReadyStatusSync(1)
End If

'**********************************************************************************
'Expand and Select AutomatedTests Folder
''**********************************************************************************
Call Fn_MyTc_NavTree_NodeOperation("Expand","Home:AutomatedTests","")
Call Fn_ReadyStatusSync(1)
Call Fn_MyTc_NavTree_NodeOperation("Select","Home:AutomatedTests","")

'**********************************************************************************
'Create TestCase Folder under AutomatedTests folder
''**********************************************************************************
iRanNo = Fn_RandNoGenerate()
sTestFolderName = Environment.Value("TestName") + "_" + Cstr(iRanNo)
bReturn = Fn_MyTc_FolderCreate("Folder",sTestFolderName,"Automation Artifact - Test Case Folder","OFF")
If bReturn = False Then
	Call Fn_UpdateLogFiles("[" + Cstr(now) + "] Failed to Create Test case Folder [" + sTestFolderName + "]", "FAIL:Test Case Folder not Created")
	Call Fn_KillProcess("")
	ExitTest
End If
Call Fn_ReadyStatusSync(2)
Call Fn_UpdateLogFiles("[" + Cstr(now) + "] - Action - PASS |  Successfully Created Test case Folder [" + sTestFolderName + "]", "")
sFolderpath="Home:AutomatedTests:" + sTestFolderName

'**********************************************************************************
'Expand and Select Test Case Folder
''**********************************************************************************
Call Fn_MyTc_NavTree_NodeOperation("Select",sFolderpath,"")
Call Fn_ReadyStatusSync(1)

''**********************************************************************************
'Creating Item under Test Case folder
''***********************************************************************************
bReturn= Fn_ItemBasicCreate(DataTable("ItemType", dtGlobalSheet),"OFF",DataTable("ItemID", dtGlobalSheet),DataTable("ItemRevision", dtGlobalSheet),DataTable("ItemName", dtGlobalSheet),DataTable("ItemDescription", dtGlobalSheet),"")

aItmInfo = split(bReturn, "-", -1, 1)
DataTable("ItemID", dtGlobalSheet) = aItmInfo(0)
DataTable("ItemRevision", dtGlobalSheet) = aItmInfo(1)
sItemPath=aItmInfo(0)+"-"+DataTable("ItemName", dtGlobalSheet)
sItemselPath=sFolderpath+":"+sItemPath

If instr(bReturn, "-") = 0 Then
	Call Fn_UpdateLogFiles("[" + Cstr(now) + "] -Action - Fail  | Failed to create Item  ["+sItemPath+"] under Testcase folder ", "Action - Fail  | Failed to create Item  ["+sItemPath+"]  under Testcase folder")
	Call Fn_KillProcess("")
	ExitTest
Else
	Call Fn_UpdateLogFiles("[" + Cstr(now) + "] - Action - PASS |  Successfully  created Item ["+sItemPath+"]  under Testcase folder", "")
    Call Fn_ReadyStatusSync(1)
End If

'**********************************************************************************
'Expand and Select Test Case Folder
''**********************************************************************************
Call Fn_MyTc_NavTree_NodeOperation("Expand",sFolderpath,"")
Call Fn_ReadyStatusSync(1)

'***********************************************************************************
' Selecting the Item 
'***********************************************************************************
bReturn= Fn_MyTc_NavTree_NodeOperation("Select",sItemselPath,"")
If bReturn = False Then
	Call Fn_UpdateLogFiles( "[" + Cstr(now) + "]-Action - Fail |  Fail to  Select Item ["+sItemPath+"]  under Testcase folder ", "Action - Fail |  Fail to Select  Item ["+sItemPath+"]  under Testcase folder")
	Call Fn_KillProcess("")
	ExitTest
Else
	Call Fn_UpdateLogFiles("[" + Cstr(now) + "] - Action - PASS |  Successfully Selected  Item ["+sItemPath+"]  under Testcase folder", "")
	Call Fn_ReadyStatusSync(1)
End If


'********************************************************************************************************************************************
' Verify the Change Ownership dialog box is displayed with the selected object and user name of the owner of the object
'*********************************************************************************************************************************************
Set ObjChng=Fn_SISW_MyTc_GetObject("ChangeOwnership")
sUserData1 = Environment.Value("TcUser1")
aUserInfo1=split(sUserData1,":",-1,1)
Call Fn_MenuOperation("Select","Edit:Change Ownership...")
Call Fn_ReadyStatusSync(1)
sUser1=aUserInfo1(2)+"/"+aUserInfo1(0)+" ("+aUserInfo1(5)+")"
If ObjChng.Exist Then
	sGetText=ObjChng.JavaButton("OwingUser").GetROProperty("attached text")
	If sUser1=sGetText Then
		Call Fn_UpdateLogFiles("[" + Cstr(now) + "] - VP1    - PASS |  Successfully Verified that  the Change Ownership dialog box is displayed with Selected object  ["+sItemPath+"] and  name ["+sUser1+"]  of the Owner of the object", "")
		Call Fn_ReadyStatusSync(1)
	Else
		Call Fn_UpdateLogFiles("[" + Cstr(now) + "] - VP1     - Fail |  Fail to  Verify that  the Change Ownership dialog box is displayed with Selected object  ["+sItemPath+"] and  name ["+sUser1+"]  of the Owner of the object", " VP1     - Fail |  Fail to  Verify that  the Change Ownership dialog box is displayed with Selected object  ["+sItemPath+"] and name ["+sUser1+"]  of the Owner of the object")
		Call Fn_KillProcess("")
		ExitTest
	End If
Else
	Call Fn_UpdateLogFiles("[" + Cstr(now) + "] - VP1     - Fail |  Fail to  Verify that  the Change Ownership dialog box is displayed with Selected object  ["+sItemPath+"] and user name ["+sUser1+"]  of the Owner of the object", " VP1     - Fail |  Fail to  Verify that  the Change Ownership dialog box is displayed with Selected object  ["+sItemPath+"] and user name ["+sUser1+"]  of the Owner of the object")
	Call Fn_KillProcess("")
	ExitTest
End If
Call Fn_Button_Click("ChgOwnershipSingle",JavaWindow("MyTeamcenter").JavaWindow("MyTcApplet").JavaDialog("ChangeOwnership"),"No")
Call Fn_ReadyStatusSync(1)

'***********************************************************************************
' Changing Ownership
'***********************************************************************************
sUser2=aUserInfo(2)+"/"+aUserInfo(3)+"/"+aUserInfo(0)+" ("+aUserInfo(5)+")"
bReturn= Fn_MyTc_ChangeOwnership("VerifyChangeOwner",sItemselPath,"Organization:"+aUserInfo(2)+":"+aUserInfo(3)+":"+aUserInfo(0)+" ("+aUserInfo(5)+")")
If bReturn = False Then
	Call Fn_UpdateLogFiles("[" + Cstr(now) + "] - VP2     - Fail |  Successfully Verified that  the Organization Selection dialog displays the Site's organization tree", "")
	Call Fn_UpdateLogFiles("[" + Cstr(now) + "] - VP3     - Fail |  Successfully Verified that  the Organization Selection dialog box is dismissed and name of User Selected["+sUser2+"] is Displaying on the New Owing User Button ", "")
	Call Fn_KillProcess("")
	ExitTest
Else
	Call Fn_UpdateLogFiles("[" + Cstr(now) + "] - VP2    - PASS |  Successfully Verified that  the Organization Selection dialog displays the Site's organization tree", "")
	Call Fn_UpdateLogFiles("[" + Cstr(now) + "] - VP3    - PASS |  Successfully Verified that  the Organization Selection dialog box is dismissed and name of User Selected["+sUser2+"] is Displaying on the New Owing User Button ", "")
	Call Fn_ReadyStatusSync(1)
End If

'***********************************************************************************
' Refreshing Window
'***********************************************************************************
Call Fn_MyTc_NavTree_NodeOperation("Select",sFolderpath,"")
Call Fn_ReadyStatusSync(1)
bReturn= Fn_MyTc_NavTree_NodeOperation("Select",sItemselPath,"")
If bReturn = False Then
	Call Fn_UpdateLogFiles( "[" + Cstr(now) + "]-Action - Fail |  Fail to  Refesh the Window  ", "Action - Fail |  Fail to  Refesh the Window")
	Call Fn_KillProcess("")
	ExitTest
Else
	Call Fn_UpdateLogFiles("[" + Cstr(now) + "] - Action - PASS |  Successfully Refeshed the Window", "")
	Call Fn_ReadyStatusSync(1)
End If
Call Fn_MyTc_NavTree_NodeOperation("Select",sFolderpath,"")
Call Fn_ReadyStatusSync(1)
Call Fn_MyTc_NavTree_NodeOperation("Select",sItemselPath,"")
Call Fn_ReadyStatusSync(1)

'***********************************************************************************
' Verifying that Ownership is Changed
'***********************************************************************************
bReturn=Fn_ObjectPropertyPanelVerify("Properties:Owner",aUserInfo(0)+" ("+aUserInfo(5)+")")
If bReturn=False Then
	Call Fn_UpdateLogFiles("[" + Cstr(now) + "] - VP4     - Fail |  Fail to Verify that  the User ["+aUserInfo(0)+" ("+aUserInfo(5)+")"+"] is now the Owner of the Object  ["+sItemPath+"]", "VP2     - Fail |  Fail to Verify that  the User ["+aUserInfo(0)+" ("+aUserInfo(5)+")"+"] is now the Owner of the Object  ["+sItemPath+"]")
Else
	Call Fn_UpdateLogFiles("[" + Cstr(now) + "] - VP4    - PASS |  Successfully Verified that  the User ["+aUserInfo(0)+" ("+aUserInfo(5)+")"+"] is now the Owner of the Object  ["+sItemPath+"]", "")
	Call Fn_ReadyStatusSync(1)
End If

'***********************************************************************************
' Changing the Ownership
'***********************************************************************************
sUserData1 = Environment.Value("TcUser4")
aUserInfo1=split(sUserData1,":",-1,1)
Call Fn_MyTc_ChangeOwnership("ChangeOwner",sItemselPath,"Organization:"+aUserInfo1(2)+":"+aUserInfo1(3)+":"+aUserInfo1(0)+" ("+aUserInfo1(5)+")")
If ObjChng.Exist Then
     bReturn=Fn_DetailRedButtonErrorMessageVerify("The access is denied.",ObjChng)
	 If   bReturn=False  Then
		 Call Fn_UpdateLogFiles("[" + Cstr(now) + "] - VP5     - Fail |  Failed to  Verify that the Error Message is Appearing [The access is denied.] ,when we are trying to change the Ownership Again", " VP5     - Fail |  Failed to  Verify that the Error Message is Appearing [The access is denied.] ,when we are trying to change the Ownership Again")
		 Call Fn_KillProcess("")
		 ExitTest
	Else
		Call Fn_UpdateLogFiles("[" + Cstr(now) + "] - VP5    - PASS |  Successfully Verified that the Error Message is Appearing [The access is denied.] ,when we are trying to change the Ownership Again", "")
		Call Fn_ReadyStatusSync(1)
	 End If
Else
	Call Fn_UpdateLogFiles("[" + Cstr(now) + "] - VP5     - Fail |  Failed to  Verify that the Error Message is Appearing [The access is denied.] ,when we are trying to change the Ownership Again", " VP5     - Fail |  Failed to  Verify that the Error Message is Appearing [The access is denied.] ,when we are trying to change the Ownership Again")
	Call Fn_KillProcess("")
	ExitTest	
End if 
'
'***************************************************************************************
'Logging in Team Center with User 2
''***************************************************************************************
Call  Fn_InvokeTeamCenter()
bReturn = Fn_TeamcenterLogin(aUserInfo(0),aUserInfo(0),aUserInfo(2) ,aUserInfo(3),"")
If bReturn = False Then
	Call Fn_UpdateLogFiles( "[" + Cstr(now) + "]-Action - Fail |  Fail to  Found Tc Session for User [" + Environment.Value("TcUser2") + "]","Action - Fail |  Fail to  Found Tc Session for User [" + Environment.Value("TcUser2") + "]")
	Call Fn_KillProcess("")
	ExitTest
Else
	Call Fn_UpdateLogFiles("[" + Cstr(now) + "] - Action - PASS |  Successfully Found Tc Session for User [" + Environment.Value("TcUser2") + "]" , "")
    Call Fn_ReadyStatusSync(1)
End If
Call Fn_SetTCSession(aUserInfo(0))
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
	Call Fn_UpdateLogFiles("[" + Cstr(now) + "] - Action - PASS |  Successfully set My Teamcenter Perspective", "")
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
	Call Fn_UpdateLogFiles("[" + Cstr(now) + "] - Action - PASS |  Successfully Reset My Teamcenter Perspective", "")
End If

'**********************************************************************************
'Searching for the Item
'**********************************************************************************
Call Fn_MenuOperation("Select","Window:Show View:Search")
Call Fn_ReadyStatusSync(1)
bReturn = Fn_MyTc_ItemSearch("", "", "","" ,DataTable("ItemID", dtGlobalSheet) ,"", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
If bReturn = False Then
	Call Fn_UpdateLogFiles( "[" + Cstr(now) + "]-Action - Fail |  Fail to Search the  Item ["+sItemPath+"] by the New Owner[AutoTest2 (autotest2)]","Action - Fail |  Fail to Search the  Item ["+sItemPath+"] by the New Owner[AutoTest2 (autotest2)]")
	Call Fn_KillProcess("")
	ExitTest
Else
	Call Fn_UpdateLogFiles("[" + Cstr(now) + "] - Action - PASS |  Successfully  Searched the  Item ["+sItemPath+"] by the New Owner [AutoTest2 (autotest2)]", "")
    Call Fn_ReadyStatusSync(1)
End If
'
'**********************************************************************************
'Verifying that Object is there for User2
'**********************************************************************************
sRootNode = JavaWindow("MyTeamcenter").JavaTree("SearchResultTree").GetItem(0)
sItemtosearch=sRootNode+":"+sItemPath
bReturn = Fn_MyTc_SrchResltTreeOperation("Select",sItemtosearch, "")
bReturn2=Fn_ObjectPropertyPanelVerify("Properties:Owner",aUserInfo(0)+" ("+aUserInfo(5)+")")
If bReturn =True and  bReturn2=True Then
	Call Fn_UpdateLogFiles("[" + Cstr(now) + "] - Action - PASS |  Successfully  Selected   Item ["+sItemPath+"] by the New Owner["+aUserInfo(0)+" ("+aUserInfo(5)+")"+"]", "")
	Call Fn_UpdateLogFiles("[" + Cstr(now) + "] - VP6    - PASS |  Successfully Verified that the Object ["+sItemPath+"] Exists and Owner is [AutoTest2 (autotest2)] when logging in with the new Owner [AutoTest2 (autotest2)]", "")
	Call Fn_ReadyStatusSync(1)
Else
	Call Fn_UpdateLogFiles("[" + Cstr(now) + "] - Action - Fail |  Failed to   Select   Item ["+sItemPath+"] by the New Owner["+aUserInfo(0)+" ("+aUserInfo(5)+")"+"]", "")
	Call Fn_UpdateLogFiles("[" + Cstr(now) + "] - VP6    - Fail |  Failed to Verify that the Object ["+sItemPath+"] Exists and Owner is ["+aUserInfo(0)+" ("+aUserInfo(5)+")"+"] when logging in with the new Owner ["+aUserInfo(0)+" ("+aUserInfo(5)+")"+"]", "")
	Call Fn_KillProcess("")
	ExitTest
End If

'**********************************************************************************
	'Log Test Result and Exit Testcase
'**********************************************************************************
	Call Fn_Setup_TestcaseExit(True)




'call Fn_Setup_TestcaseInit()
'
'msgbox Fn_MyTc_ChangeOwnership("VerifyChangeOwner","Home:AutomatedTests:ChgOwnershipSingle_18177:000024-Item1","Organization:Change Analysts:Change Analyst:cmuser01 (cmuser01)")
