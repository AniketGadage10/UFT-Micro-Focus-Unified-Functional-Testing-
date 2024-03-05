'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

'/$$$$   TESTCASE NAME   :   Checkout_Disable_After_Checkout
'/$$$$ 
'/$$$$   DESCRIPTION     :  Check out is grayed out when item is revised
'/$$$$  										
'/$$$$
'/$$$$	QART Link	10.1	:	http://cipgweb/qacgi-bin/tt_view.cgi?release=TC_10.1&feature=REGMN&cobid=222088&tcobid=1725260
'/$$$$
'/$$$$	HISTORY			:	      		AUTHOR			              DATE		                   VERSION				Build					
'/$$$$
'/$$$$	CREATED BY      :		Pooja Shilwant				25-June-2013				1.0					  2013060400			
'/$$$$
'/$$$$	REVIWED BY      :		
'/$$$$
'/$$$$  SERVER               :     pnv6s224
'/$$$$  
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'//$$$$    Porting ( Test case Name)          : Checkout_Disable_After_Checkout
'//$$$$
'//$$$$    QART Link           :https://tidev.industrysoftware.automation.siemens.com/qart/qacgi-bin/tt_view.cgi?release=TC_11.1&feature=REGMN&category=CheckIN-OUT&testcase=Checkout_Disable_After_Checkout
'//$$$$ 
'//$$$$    PORTED BY           :Jotiba Takkekar                     Date: 26-02-2015
'//$$$$ 
'//$$$$
'//$$$$    Run on Tc Build     : TC 11.2 (2015020400)  Server pnv6s106 
'//$$$$   
'//$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$


Option Explicit
'**********************************************************************************
'Variable Declaration
'**********************************************************************************
Dim bReturn
Dim iRanNo,iCnt
Dim sFilePath,sTopAssyPath,strMenuPath,strMenu,sPopMenuFile,sPopMenu
Dim aUser,aSOAInput
Dim objSummaryEdit,objWindow

'-************************************************************************************
'' Assign AutomationDir to sPath env variable –
' Optional (Only if script needs to have path for further use)
'*************************************************************************************
Environment.Value("sPath") = Fn_GetEnvValue("User", "AutomationDir")

'**********************************************************************************
'Set up test case Init
'**********************************************************************************
Call Fn_Setup_TestcaseInit()
Call Fn_UpdateLogFiles("-----------------------------------------------------------------------------------------------", "")

'*****************************************************************************************************
'Create Root Item Under Testcase Folder.
''*****************************************************************************************************
DataTable.SetCurrentRow(1)
iRanNo = Fn_RandNoGenerate()
Environment.Value("sChildFolder") =  Mid(Environment.Value("TestName"),1, 25) &"_"& iRanNo

ReDim aSOAInput(6)
aUser = split (Environment.Value("TcUser4"),":",-1,1)

aSOAInput(0) = aUser(0)
aSOAInput(1) =  "AutomatedTests"
aSOAInput(2) = Environment.Value("sChildFolder")
aSOAInput(3) = "ImportStructure"
aSOAInput(4) = Environment.Value("TestDir")& "\import_structure.xml"
aSOAInput(5) = 0

bReturn = Fn_SOA_CreateTCObject(aSOAInput)
If bReturn = False Then
		Call Fn_UpdateLogFiles(Cstr(now) + " - ACTION - FAIL | Failed to create structure.", "FAIL: Failed to create structure.")
		Call Fn_KillProcess("")
		ExitTest
End If

'**********************************************************************************
'	Filling Item details in Data table
'1. Create an Item structure in PSE
''**********************************************************************************
Environment.Value("TestFolderName") = "Home:AutomatedTests:" + Environment.Value("sChildFolder")

bReturn = Fn_SaveItemDetailsInDataTable("Append","","","")
If bReturn = False Then
		Call Fn_UpdateLogFiles(Cstr(now) + " - ACTION - FAIL | Failed to save structure.", "FAIL: Failed to save structure.")
		Call Fn_KillProcess("")
		ExitTest
End If
'- - - - - - - - - - - -- - - - - - - - - - - -- - - - - - - - - - - -- - - - - - - - - - - -- - - - - - - - - - - -- - - - - - - - - - - -- - - - - - - - - - - -
DataTable.SetCurrentRow(1)
DataTable("ItemPath", dtGlobalSheet) = Environment.Value("TestFolderName") + ":" + DataTable("ItemID", dtGlobalSheet)+"-"+DataTable("ItemName",dtGlobalSheet)
DataTable("ItemRevision", dtGlobalSheet) = DataTable("ItemID", dtGlobalSheet) +"/"+DataTable("ItemRevID", dtGlobalSheet)+";1"+ "-" + DataTable("ItemName", dtGlobalSheet) 
DataTable("BOMPath", dtGlobalSheet) = DataTable("ItemRevision", dtGlobalSheet)+" (View)"
'- - - - - - - - - - - -- - - - - - - - - - - -- - - - - - - - - - - -- - - - - - - - - - - -- - - - - - - - - - - -- - - - - - - - - - - -- - - - - - - - - - - -
sTopAssyPath = DataTable("BOMPath", dtGlobalSheet)
Call Fn_UpdateLogFiles("[" + Cstr(now) + "] - ACTION - PASS | Successfully created Node - ["+DataTable("ItemRevision", dtGlobalSheet)+"] for user - ["& aUser(0) &"] under TestCase folder - [" & Environment.Value("TestFolderName") &"]  using SOA", "")
'- - - - - - - - - - - -- - - - - - - - - - - -- - - - - - - - - - - -- - - - - - - - - - - -- - - - - - - - - - - -- - - - - - - - - - - -- - - - - - - - - - - -
DataTable.SetCurrentRow(2)
DataTable("ItemRevision", dtGlobalSheet) = DataTable("ItemID", dtGlobalSheet) +"/"+DataTable("ItemRevID", dtGlobalSheet)+";1"+ "-" + DataTable("ItemName", dtGlobalSheet)
DataTable("BOMPath", dtGlobalSheet) = sTopAssyPath+":" + DataTable("ItemRevision", dtGlobalSheet)
Call Fn_UpdateLogFiles("[" + Cstr(now) + "] - ACTION - PASS | Successfully Created Node - ["+DataTable("ItemRevision", dtGlobalSheet)+"] for user - ["& aUser(0) &"] under TestCase folder - [" & Environment.Value("TestFolderName") &"]  using SOA", "")

'**********************************************************************************
' Login to RAC.
'**********************************************************************************
DataTable.SetCurrentRow (1)
bReturn = Fn_ReUserTcSession(True, True , Environment.Value("TcUser4"))
If bReturn = False Then
		Call Fn_UpdateLogFiles("[" + Cstr(now) + "] - ACTION - FAIL  | Failed to Log into Teamcenter with user [" + Environment.Value("TcUser4") + "] ", "FAIL:Failed to Log into Teamcenter with user [" + Environment.Value("TcUser4") + "] ")
		Call Fn_KillProcess("")
		ExitTest
Else
		Call Fn_UpdateLogFiles("[" + Cstr(now) + "] - ACTION - PASS | Successfully Logged into Teamcenter with user [" + Environment.Value("TcUser4") + "] ", "")
		Call Fn_ReadyStatusSync(1)
End If

''**********************************************************************************
''Set the MyTeamcenter Perspective
'''*********************************************************************************
bReturn = Fn_SetPerspective("My Teamcenter")
If bReturn = False Then
		Call Fn_UpdateLogFiles("[" + Cstr(now) + "] - ACTION - FAIL | Failed to Set MyTeamcenter Perspective", "FAIL:Failed to Go to MyTeamcenter Module")
		Call Fn_KillProcess("")
		ExitTest
Else
		Call Fn_ReadyStatusSync(3)
		Call Fn_UpdateLogFiles("[" + Cstr(now) + "] - ACTION - PASS | Successfully set [ MyTeamcenter ] Perspective", "")
		Call Fn_RefreshWindow()
End If

'**********************************************************************************
'Reset  Perspective to Display Default state
'**********************************************************************************
bReturn = Fn_ResetPerspective()
If bReturn = False Then
		Call Fn_UpdateLogFiles("[" + Cstr(now) + "] - ACTION - FAIL | Failed to Reset Perspective", "FAIL:Failed to Reset Perspective")
		Call Fn_KillProcess(Environment.Value("KillProcesses"))
		ExitTest
Else
		Call Fn_ReadyStatusSync(1)
		Call Fn_UpdateLogFiles("[" + Cstr(now) + "] - ACTION - PASS | Successfully Reset [ MyTeamcenter ] Perspective", "")
		Call Fn_ReadyStatusSync(1)
End If

'**********************************************************************************
'Expand and Select AutomatedTests Folder
'**********************************************************************************
DataTable.SetCurrentRow(1)

bReturn =  Fn_MyTc_NavTree_NodeOperation("Expand","Home:AutomatedTests","")
If bReturn = False Then
		Call Fn_UpdateLogFiles("[" + Cstr(now) + "] - ACTION - FAIL | Failed to Expand Folder [Home:AutomatedTests]", "Fail: Failed to expand folder")
		Call Fn_KillProcess("")
		ExitTest
Else
		Call Fn_UpdateLogFiles("[" + Cstr(now) + "] - ACTION - PASS | Successfully Expanded folder [Home:AutomatedTests]", "")
		Call Fn_ReadyStatusSync(1)	
End If
'- - - - - - - - - - - -- - - - - - - - - - - -- - - - - - - - - - - -- - - - - - - - - - - -- - - - - - - - - - - -- - - - - - - - - - - -- - - - - - - - - - - -
bReturn =  Fn_MyTc_NavTree_NodeOperation("Expand",Environment.Value("TestFolderName") ,"")
If bReturn = False Then
		Call Fn_UpdateLogFiles("[" + Cstr(now) + "] - ACTION - FAIL | Failed to Expand Folder ["& Environment.Value("TestFolderName")  &"]", "Fail: Failed to expand folder")
		Call Fn_KillProcess("")
		ExitTest
Else
		Call Fn_UpdateLogFiles("[" + Cstr(now) + "] - ACTION - PASS | Successfully Expanded folder ["& Environment.Value("TestFolderName")  &"]", "")
		Call Fn_ReadyStatusSync(1)	
End If

'***********************************************************************************
'Send Root Item to PSE
''**********************************************************************************
sPopMenuFile=Fn_LogUtil_GetXMLPath("RAC_PopupMenu")
sPopMenu=Fn_GetXMLNodeValue(sPopMenuFile,"SendToStructureManager")
bReturn = Fn_MyTc_NavTree_NodeOperation("PopupMenuSelect", DataTable("ItemPath", dtGlobalSheet), sPopMenu)
If bReturn = FALSE Then
		Call Fn_UpdateLogFiles("[" + Cstr(now) + "] - ACTION - FAIL | Failed to Send Item To PSE ", "FAIL:Send To Send Item To PSE ")
		Call Fn_KillProcess("")
		ExitTest
End If
Call Fn_UpdateLogFiles("[" + Cstr(now) + "] - ACTION - PASS | Sucessfully Sent Root Item ["+ DataTable("ItemPath", dtGlobalSheet)+"] to PSE ", "")
Call Fn_ReadyStatusSync(5)

'*******************************************************************************************
'2. Select a BOM line and show both the Summary tab and Properties view 
'*******************************************************************************************
bReturn= Fn_PSE_BOMTable_NodeOperationExt("Select",sTopAssyPath, "", "", "")
If bReturn = False Then
		Call Fn_UpdateLogFiles("[" + Cstr(now) + "] - ACTION - FAIL | Failed to Select the Top BOM line - ["&sTopAssyPath& "] in Structure Manager ", "FAIL : Failed toSelect the Top BOM line - ["&sTopAssyPath& "] in Structure Manager ")
		Call Fn_KillProcess("")
		ExitTest
Else
		Call Fn_UpdateLogFiles("[" + Cstr(now) + "] - ACTION - PASS | Successfully Selected the Top BOM line - ["&sTopAssyPath& "] in Structure Manager ", "")
		Call Fn_ReadyStatusSync(1)
End If

'*******************************************************************************************
'show both the Summary tab and Properties view 
'*******************************************************************************************
bReturn =Fn_SetView("General:Summary")
If bReturn = False Then	
		Call Fn_UpdateLogFiles("[" + Cstr(now) + "] - ACTION - FAIL | Failed To Open [ Summary ] tab ", "FAIL: Failed To Open [ Summary ] tab")
		Call Fn_KillProcess("")
		ExitTest
Else
		Call Fn_UpdateLogFiles("[" + Cstr(now) + "] - ACTION - PASS | Successfully Opened [ Summary ] tab ", "")
		Call Fn_ReadyStatusSync(5)
End If	
'- - - - - - - - - - - -- - - - - - - - - - - -- - - - - - - - - - - -- - - - - - - - - - - -- - - - - - - - - - - -- - - - - - - - - - - -- - - - - - - - - - - -
strMenuPath=Fn_LogUtil_GetXMLPath("RAC_Menu")
strMenu=Fn_GetXMLNodeValue(strMenuPath, "WindowShowViewProperties")

bReturn = Fn_MenuOperation("Select",strMenu)
If bReturn = False Then	
		Call Fn_UpdateLogFiles("[" + Cstr(now) + "] - ACTION - FAIL | Failed To Perform [ "+strMenu+" ] Menu operation ", "FAIL: Failed To Perform [ "+strMenu+" ] Menu operation ")
		Call Fn_KillProcess("")
		ExitTest
Else
		Call Fn_UpdateLogFiles("[" + Cstr(now) + "] - ACTION - PASS | Successfully Performed [ "+strMenu+" ] Menu operation to open [Properties] panel ", "")
		Call Fn_ReadyStatusSync(5)
End If	

'**********************************************************************************
'3. Revise the Item (File->revise)
'**********************************************************************************
bReturn= Fn_PSE_BOMTable_NodeOperationExt("Select",sTopAssyPath, "", "", "")
If bReturn = False Then
		Call Fn_UpdateLogFiles("[" + Cstr(now) + "] - ACTION - FAIL | Failed to Select the Top BOM line - ["&sTopAssyPath& "] in Structure Manager ", "FAIL : Failed toSelect the Top BOM line - ["&sTopAssyPath& "] in Structure Manager ")
		Call Fn_KillProcess("")
		ExitTest
End If
strMenuPath=Fn_LogUtil_GetXMLPath("RAC_Menu")
strMenu=Fn_GetXMLNodeValue(strMenuPath, "FileRevise")
bReturn =Fn_MenuOperation("Select",strMenu)
If bReturn = False Then	
		Call Fn_UpdateLogFiles("[" + Cstr(now) + "] - ACTION - FAIL | Failed To Perform [ File:Revise... ] Menu operation ", "FAIL: Failed To Perform [ File:Revise... ] Menu operation ")
		Call Fn_KillProcess("")
		ExitTest
Else
		Call Fn_UpdateLogFiles("[" + Cstr(now) + "] - ACTION - PASS | Successfully Performed [ File:Revise... ] Menu operation to Revise the Item ", "")
		Call Fn_ReadyStatusSync(5)
End If	
'- - - - - - - - - - - -- - - - - - - - - - - -- - - - - - - - - - - -- - - - - - - - - - - -- - - - - - - - - - - -- - - - - - - - - - - -- - - - - - - - - - - -
DataTable.SetCurrentRow(1)
DataTable("ItemRevision", dtGlobalSheet) = DataTable("ItemID", dtGlobalSheet) +"/B;1"+ "-" + DataTable("ItemName", dtGlobalSheet) 

bReturn =Fn_ObjectRevRevise("", "", "" )
If bReturn = False Then	
		Call Fn_UpdateLogFiles("[" + Cstr(now) + "] - ACTION - FAIL | Failed To Revise the Item to [ "+DataTable("ItemRevision", dtGlobalSheet) +"] ", "FAIL: Failed To Revise the Item to [ "+DataTable("ItemRevision", dtGlobalSheet) +"] ")
		Call Fn_KillProcess("")
		ExitTest
Else
		Call Fn_UpdateLogFiles("[" + Cstr(now) + "] - ACTION - PASS | Successfully Revised the Item to [ "+DataTable("ItemRevision", dtGlobalSheet) +"] ", "")
		Call Fn_ReadyStatusSync(5)
End If	

'**********************************************************************************
'(V)Verify:
'1.  The BOM line 
'2.  and properties view and 
'3.  summary Tab are getting updated.
'**********************************************************************************
sTopAssyPath=DataTable("ItemRevision", dtGlobalSheet)+" (View)"

bReturn= Fn_PSE_BOMTable_NodeOperationExt("Exists",sTopAssyPath, "", "", "")
If bReturn = False Then
		Call Fn_UpdateLogFiles("[" + Cstr(now) + "] - ACTION - FAIL | Failed to Verify the BOM line is updated to - ["&sTopAssyPath& "] in Structure Manager ", "FAIL : Failed to Verify the BOM line is updated to - ["&sTopAssyPath& "] in Structure Manager ")
		Call Fn_KillProcess("")
		ExitTest
Else
		Call Fn_UpdateLogFiles("[" + Cstr(now) + "] - VERIFY - PASS | VP1: Successfully Verified the BOM line is updated to - ["&sTopAssyPath& "] in Structure Manager ", "")
		Call Fn_ReadyStatusSync(1)
End If
'- - - - - - - - - - - -- - - - - - - - - - - -- - - - - - - - - - - -- - - - - - - - - - - -- - - - - - - - - - - -- - - - - - - - - - - -- - - - - - - - - - - -
Datatable.SetCurrentRow(1)

bReturn=Fn_ObjectPropertyPanelVerify("Properties:Item Name;Properties:Revision","TopA;B")
If bReturn =False Then
		Call Fn_UpdateLogFiles("[" + Cstr(now) + "] - VERIFY - FAIL | VP2: Failed to Verify Item Revision is updated to [ "+sTopAssyPath+" ] from the Properties Panel ","FAIL | Failed to Verify Item Revision is updated to [ "+sTopAssyPath+" ] from the Properties Panel ")
		Call Fn_KillProcess(Environment.Value("KillProcesses"))
		ExitTest
Else
		Call Fn_UpdateLogFiles("[" + Cstr(now) + "] - VERIFY - PASS | VP2.1:Successfully Verified Property [Revision] has value[B] from the Properties Panel ","")
		Call Fn_UpdateLogFiles("[" + Cstr(now) + "] - VERIFY - PASS | VP2.2:Successfully Verified Property [Item Name] has value[TopA] from the Properties Panel ","")
		Call Fn_UpdateLogFiles("[" + Cstr(now) + "] - VERIFY - PASS | VP2.3:Hence Verified Item Revision is updated to [ "+sTopAssyPath+" ] from the Properties Panel ","")
		Call Fn_ReadyStatusSync(1)
End If 
'- - - - - - - - - - - -- - - - - - - - - - - -- - - - - - - - - - - -- - - - - - - - - - - -- - - - - - - - - - - -- - - - - - - - - - - -- - - - - - - - - - - -
sTopAssyPath=DataTable("ItemRevision", dtGlobalSheet)
Set objWindow=Fn_SISW_GetObject("DefaultWindow")
Set objSummaryEdit=objWindow.JavaEdit("SummaryHeader")

bReturn=Fn_UI_Object_GetROProperty("",objSummaryEdit, "value")
Call Fn_ReadyStatusSync(1)
If bReturn=sTopAssyPath Then
		Call Fn_UpdateLogFiles("[" + Cstr(now) + "] - VERIFY - PASS | VP3:Successfully Verified Item Revision is updated to [ "+sTopAssyPath+" ] from Summary tab ","")
		Call Fn_ReadyStatusSync(1)
Else
		Call Fn_UpdateLogFiles("[" + Cstr(now) + "] - VERIFY - FAIL | VP3: Failed to Verify Item Revision is updated to [ "+sTopAssyPath+" ] from Summary tab ","FAIL | Failed to Verify Item Revision is updated to[ "+sTopAssyPath+" ] from Summary tab ")
		Call Fn_KillProcess(Environment.Value("KillProcesses"))
		ExitTest
End If

'**********************************************************************************
'(V)Verify:
'and also Check-out and Edit icon should be enabled after revision.
'**********************************************************************************
bReturn=Fn_ToolBarOperation("IsEnabled", "Check Out...","" )
If bReturn=True Then
		Call Fn_UpdateLogFiles("[" + Cstr(now) + "] - VERIFY - PASS | VP4:Successfully Verified [Check-out and Edit] icon is enabled after revision ","")
		Call Fn_ReadyStatusSync(1)
Else
		Call Fn_UpdateLogFiles("[" + Cstr(now) + "] - VERIFY - FAIL | VP4: Failed to Verify [Check-out and Edit] icon is enabled after revision ","FAIL | Failed to Verify [Check-out and Edit] icon is enabled after revision ")
		Call Fn_KillProcess(Environment.Value("KillProcesses"))
		ExitTest
End If

''*********************************************************************************
'Call for Code Coverage
'**********************************************************************************
Set objSummaryEdit=Nothing
Set objWindow=Nothing

'- - - - - - - - - - - -- - - - - - - - - - - -- - - - - - - - - - - -- - - - - - - - - - - -- - - - - - - - - - - -
Call Fn_UpdateLogFiles("-----------------------------------------------------------------------------------------------", "")
Call Fn_Setup_TestcaseExit(True)

