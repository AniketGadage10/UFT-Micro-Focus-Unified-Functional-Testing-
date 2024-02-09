Option Explicit
'===================================================================================================================
' Function List
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'001.	Fn_SISW_NX_Setup_Restore
'002.	Fn_SISW_NX_Setup_RoutingOperation
'003.	Fn_SISW_NX_Setup_DefaultState
'004.	Fn_SISW_NX_Setup_ReadyStatusSync
'005.	Fn_SISW_NX_Setup_Get_Cursor
'006.	Fn_SISW_NX_Setup_SetCursor
'007.	Fn_SISW_NX_Setup_InvokeCommandFinder
'008.	Fn_SISW_NX_Setup_CmdFinderOperation
'009.	Fn_SISW_NX_Setup_LoadRunMacro
'010.	Fn_SISW_NX_Setup_NXExit
'011.	Fn_SISW_NX_Setup_DisplayResourceBar
'012.	Fn_SISW_NX_Setup_UserInterFacePrefrences
'013.	Fn_SISW_NX_Setup_TabOperation
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Function Name     			: Fn_SISW_NX_Setup_Restore
' Function Description 			 : Function used to Restore NX window
' Parameters			: 			NA						  					
' Return Value		    : 		Nothing
' 
' Examples		    	: 			Call Fn_SISW_NX_Setup_Restore()
' History               :  
'		Developer Name		 		Date	  				Rev. No. 				Changes					Reviewer
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'		Pranav Ingle 				03-Dec-2013	 				1.0							
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_NX_Setup_Restore()
	On Error Resume Next
   If  Window("NXWindow").Exist(1)  Then
	   If Window("NXWindow").GetROProperty("visible")=False Then
			Window("NXWindow").Restore
	   End If
   End If
End Function
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Function Name     			: Fn_SISW_NX_Setup_RoutingOperation
' Function Description 			 : To handle routing window till NXlaunch
' Parameters			: 			NA
' Return Value           :        True/False
'
' Examples		    	: 		Call 	 Fn_SISW_NX_Setup_RoutingOperation()
' History               :  
'		Developer Name			 		Date	  					Rev. No. 				Changes					Reviewer
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'		Nilesh Gadekar  	    	07-Sept-2013					1.0							
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_NX_Setup_RoutingOperation()
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_NX_Setup_RoutingOperation"
	Dim objRouting,iTimeOut,objNXWin
	Set objRouting=Window("Routing")
	Set objNXWin=Window("NXWindow")
	iTimeOut=240 
	If objRouting.Exist(10) Then
		Fn_SISW_NX_Setup_RoutingOperation=True
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: NX Application launched successfully")
	Else
		Fn_SISW_NX_Setup_RoutingOperation=False
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: NX Application launched successfully")
	End If

	If objNXWin.Exist( iTimeOut) Then
		Fn_SISW_NX_Setup_RoutingOperation=True
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: NX Application launched successfully")
	Else
		Fn_SISW_NX_Setup_RoutingOperation=False
        Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: NX Application launching taken more than 4 minutes ")
	End If
End Function
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Function Name     			: Fn_SISW_NX_Setup_DefaultState
' Function Description 			 : To Rest NXinto its default state
' Parameters			: 			NA
' Return Value           :        True/False
' Pre-requisite		    : 			Nothing
' Function Call          :  			Fn_SISW_NX_Setup_RoutingOperation
' Examples		    	: 		Call 	 Fn_SISW_NX_Setup_DefaultState()
' History               :  
'			Developer Name				 Date	  				Rev. No. 				Changes					Reviewer
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'			Nilesh Gadekar  	    07-Sept-2013					1.0							
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_NX_Setup_DefaultState()
 GBL_FAILED_FUNCTION_NAME="Fn_SISW_NX_Setup_DefaultState"
   Dim objNXWin,bResult,objChild,objChilds,iCount
   Set objNXWin=Window("NXWindow")
   Fn_SISW_NX_Setup_DefaultState=True
   bResult=Fn_SISW_NX_Setup_RoutingOperation()
   If bResult=False Then
	   Call Fn_SISW_NX_Setup_RoutingOperation()
   End If
	'Close all the unexpected dialogs and windos under NX window
   If objNXWin.Exist(5)  Then
	   Set objChild=Description.Create()
		objChild("Class Name").Value="Window|Dialog"
		objChild("Class Name").RegularExpression=True
		If objNXWin.Dialog("Welcome Page").Exist(5) Then
			Call Fn_SISW_NX_UI_ButtonOperation("Fn_SISW_NX_Setup_DefaultState", "Click", objNXWin.Dialog("Welcome Page"), "Don't")
		End If

	   Set objChilds=objNXWin.ChildObjects(objChild)
	   For iCount=0 to objChilds.Count-1
			objChilds(iCount).Close()
			Wait 1
	   Next
 		If objNXWin.GetROProperty("visible")=False Then
			objNXWin.Restore()
			Wait 1
		End If
		'Maximize the NXWindow
		objNXWin.Maximize()
   End If

   Set objNXWin=Nothing
   Set objChilds=Nothing
   Set objChild=Nothing
End Function

'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Function Name		:				Fn_SISW_NX_Setup_ReadyStatusSync(iIterations)

'Description			 :		 		 This function waits till Application`s System Progress Monitor Icon  Exist/Visible
'Parameters			   :	 			1. iIterations: No. of times to be checked for Existance of System Progress Monitor Icon  text
											
'Return Value		   : 				None

'Pre-requisite			:		 		NX Should Be Launched

'Examples				:				 Call Fn_SISW_NX_Setup_ReadyStatusSync(2)

'History:
'				Developer Name					Date								Rev. No.			Changes Done			Reviewer	
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'				Pranav 	Ingle					07-Sept-2013							1.0
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_NX_Setup_ReadyStatusSync(iIterations)
	Dim iCounter,objHierarchy,i
	Dim iTimeout,bReturn

    Set	objHierarchy=Window("WorkInProgress").WinObject("SysProgressIcon")
	For iCounter = 1 to iIterations
		'-------To Check The Existance Of The System Progress ICON--------------------------------
		If objHierarchy.Exist(1) Then
			'-------To Wait  either for 10 sec or  untill  System Progress ICON get dissapear  -------------------------------
			objHierarchy.WaitProperty "exist",true,10000
		Else
			Exit For
		End If
	Next

	Call Fn_SISW_NX_Setup_SetCursor()

	iTimeout=240
	For iCounter =1 To iIterations
			For i=1 To iTimeout
				bReturn=Fn_SISW_NX_Setup_Get_Cursor()
'				If  bReturn="65539" OR bReturn="65541" Then   '///  Normal pointers
				If  bReturn="65543" OR bReturn="65561" Then  '//  Poiters for wait 
						Wait 1
				Else
					Exit For
				End If
			Next
	Next

	If Dialog("ServerBusy").Exist(1) Then
		wait 5
		For iCounter = 1 To iIterations
			If Dialog("ServerBusy").Exist(1) Then
				wait 2
			Else
				Exit For
			End If
		Next
	End If
End Function

'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Function Name     			: Fn_SISW_NX_Setup_Get_Cursor
' Function Description 			 : Get the Curser status value in Integer format
' Parameters			: 			Nothing
' Return Value           :        Cursor status value
' 	
' Examples		    	: 			Call  Fn_SISW_NX_Setup_Get_Cursor()

' History               :  Developer Name			 Date	  				Rev. No. 				Changes					Reviewer			Reviewed Date	
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'								Pranav Ingle		 	  07-Sept-2013				1.0							
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_NX_Setup_Get_Cursor()

   Dim hwnd,pid,thread_id

	extern.Declare micLong,"GetForegroundWindow","user32.dll","GetForegroundWindow"
	extern.Declare micLong,"AttachThreadInput","user32.dll","AttachThreadInput", micLong, micLong,micLong
	extern.Declare micLong,"GetWindowThreadProcessId","user32.dll","GetWindowThreadProcessId", micLong, micLong
	extern.Declare micLong,"GetCurrentThreadId","kernel32.dll","GetCurrentThreadId"
	extern.Declare micLong,"GetCursor","user32.dll","GetCursor"

    hwnd = extern.GetForegroundWindow()

    pid = extern.GetWindowThreadProcessId(hWnd, NULL)
    thread_id=extern.GetCurrentThreadId()
    extern.AttachThreadInput pid,thread_id,True

    Fn_SISW_NX_Setup_Get_Cursor=extern.GetCursor()

    extern.AttachThreadInput pid,thread_id,False

End function

'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Function Name     			: Fn_SISW_NX_Setup_SetCursor
' Function Description 			 : To set  Curser on top of NX window
' Parameters			: 			NA
' Return Value           :        True/False
' Pre-requisite		    : 			Nothing
' Function Call          :  			Fn_SISW_NX_Setup_Restore
' Examples		    	: 			 Fn_SISW_NX_Setup_SetCursor()

' History               :  Developer Name			 Date	  				Rev. No. 				Changes					Reviewer			Reviewed Date	
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'								Pranav Ingle 	    	07-Dec-2013					1.0							
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_NX_Setup_SetCursor()
	On Error Resume Next
   Dim width, height
   Call Fn_SISW_NX_Setup_Restore()
	width=Window("NXWindow").getroproperty("width")
	height=Window("NXWindow").getroproperty("height")
   Window("NXWindow").Click Cint(width/2),(height-15)
End Function
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Function Name     			: Fn_SISW_NX_Setup_InvokeCommandFinder
' Function Description 			 : To Invoke Command Finder Dialog
' Parameters			: 			NA
' Return Value           :        True/False
' Pre-requisite		    : 			Nothing
' Function Call          :  			
' Examples		    	: 			 Fn_NX_SetCursor()
' History               :  Developer Name			 Date	  				Rev. No. 				Changes					Reviewer			Reviewed Date	
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'								Nilesh Gadekar  	    28-Aug-2013				1.0							
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_NX_Setup_InvokeCommandFinder()
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_NX_Setup_InvokeCommandFinder"
	Dim objCmd
	
	Call Fn_SISW_NX_Setup_Get_Cursor()
	Set objCmd=Window("NXWindow").Dialog("Command Finder")
	If  objCmd.Exist(5) Then
		Fn_SISW_NX_Setup_InvokeCommandFinder=True
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Command finder dialog already opened")
	Else
		Call Fn_SISW_NX_Setup_LoadRunMacro("Set",Environment.Value("sPath")+"\TestData\NX\Macro\InvokeCMD.macro")
		If  objCmd.Exist(5) Then
			Fn_SISW_NX_Setup_InvokeCommandFinder=True
		Else
			Fn_SISW_NX_Setup_InvokeCommandFinder=False
		End If
	End If
	Set objCmd=Nothing
End Function
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Function Name     			: Fn_SISW_NX_Setup_CmdFinderOperation
' Function Description 			 : To  perform search and enable activity in command finder dialog
' Parameters			: 			sCommand: Command to be find out and performoperation on it
'											sAction:  Action to perform on command
'											sReserve: For future use
'
' Return Value           :        True/False
' 
' Examples		    	: 			 Fn_SISW_NX_Setup_CmdFinderOperation("See-Thru All","Show on Menu","")
' History               :  Developer Name			 Date	  				Rev. No. 				Changes					Reviewer			Reviewed Date	
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'								Nilesh Gadekar  	28-Aug-2013	 				1.0							
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_NX_Setup_CmdFinderOperation(sCommand,sAction,sReserve)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_NX_Setup_CmdFinderOperation"
	Dim ObjCmdDialog,bResult,bFlag
	Dim left1,top,right1,bottom,x,y
	Set ObjCmdDialog=Window("NXWindow").Dialog("Command Finder")

	If  ObjCmdDialog.Exist(5)=False Then
		Call  Fn_SISW_NX_Setup_InvokeCommandFinder()
		Wait 3
	End If

	If sCommand<>"" Then
		bResult=Fn_SISW_NX_UI_EditBoxOperation("Fn_SISW_NX_Setup_CmdFinderOperation","Set",ObjCmdDialog,"Search", sCommand)
		If bResult=False Then
			Fn_SISW_NX_Setup_CmdFinderOperation=False
			Set ObjCmdDialog=Nothing
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to set value "+sCommand+ " in Search Editbox of Command finder dialog")
			Exit Function
		Else
			Fn_SISW_NX_Setup_CmdFinderOperation=True
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Pass : Successfully set value "+sCommand+ " in Search Editbox of Command finder dialog")
		End If
	End If
    
	bResult=Fn_SISW_NX_UI_ButtonOperation("Fn_SISW_NX_Setup_CmdFinderOperation","Click",ObjCmdDialog,"Find Command")
	If bResult=False Then
		Fn_SISW_NX_Setup_CmdFinderOperation=False
		Set ObjCmdDialog=Nothing
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Click on button Find Command of Command finder dialog")
		Exit Function
	Else
		Fn_SISW_NX_Setup_CmdFinderOperation=True
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Clicked on button Find Command of Command finder dialog")
	End If
	Call Fn_SISW_NX_Setup_ReadyStatusSync(1)

	Select Case sAction
		Case "Start"
			ObjCmdDialog.WinList("ListBox").Select 0
		Case "Start2"
			ObjCmdDialog.WinList("ListBox").Select 1
	End Select
	Set ObjCmdDialog=Nothing
	wait 1
End Function

'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Function Name		:				Fn_SISW_NX_Setup_LoadRunMacro (sAction,sMacroPath)

'Description			 :		 		 Loads the Executable Via Tool--> Macro--->Replay...

''Parameters			   :	 			1. sAction : Action to Perform (create\verify)
'									
'														2. sMacroPath : Location of  Macro File
'														
'Return Value		   : 				TRUE \ FALSE

'Pre-requisite			:		 		 Nx should be launched  and Executable must be present at desired location
'Examples				:				 

'History					 :		
'		Developer Name				Date						Rev. No.											Reviewer
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'		Nilesh Gadekar			28-Aug-2013					1
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public function Fn_SISW_NX_Setup_LoadRunMacro(sAction,sMacroPath)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_NX_Setup_LoadRunMacro"
Dim ObjHierarchy,sFileName,bReturn
Set ObjHierarchy=Dialog("MacroPlayback")
Call Fn_SISW_NX_Setup_Restore()
 sFileName=sMacroPath
 Fn_SISW_NX_Setup_LoadRunMacro=False
 
If  Window("NXWindow").Exist Then

	'===============To select  Tools--> Macro--->PlayBack...============
	If ObjHierarchy.Exist(5)=False Then
		bReturn=Fn_SISW_NX_General_MenuOperation("Select","macrorun","")
	End If
	If  ObjHierarchy.Exist(5)=False Then
		Set ObjHierarchy=Dialog("MacroPlayback")
	End If
	If  bReturn=true and ObjHierarchy.Exist(10) Then
		'------------------To Set  the Macro  file name-------------------------
		If  Fn_SISW_NX_UI_EditBoxOperation("Fn_SISW_NX_Setup_LoadRunMacro","Set",ObjHierarchy,"MacroFileName",sFileName) Then
		Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL :Fail to Set the Macro File Name "+sMacroName+" into File Name Edit Box ")
			Exit Function
		End If
		wait 2
		 '=======================To  Click On OK Button ======================
		If Fn_SISW_NX_UI_ButtonOperation("Fn_SISW_NX_Setup_LoadRunMacro",  "Click", ObjHierarchy, "OK") Then
			 Fn_SISW_NX_Setup_LoadRunMacro=True
		else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL :Fail to Click on  [ OK ] Button")
			Exit Function
		End If
		
		'If Error occurs during Macro run then, to return false, handled it
		If Window("NXWindow").Dialog("MacroOutOfSync").Exist(10) Then
			Window("NXWindow").Dialog("MacroOutOfSync").WinButton("OK").Click 5,5,micLeftBtn
			Wait 1
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : To execute Macro : [ "+sFileName+" ]")
			Fn_SISW_NX_Setup_LoadRunMacro=False
		End If
	else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Dialog"+ObjHierarchy.ToString()+" Not Found")
	end if
else 
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Window "+ Window("NXWindow").ToString()+" Does Not Exist")
End If
Set ObjHierarchy=nothing
End Function

'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Function Name		:				Fn_SISW_NX_Setup_NXExit (sAction,sMacroPath)

'Description			 :		 		 Loads the Executable Via Tool--> Macro--->Replay...

''Parameters			   :	 			1. sSaveOption : Click on Button of Save Dialog
'														
'Return Value		   : 				TRUE \ FALSE

'Examples				:				 Fn_SISW_NX_Setup_NXExit("")

'History					 :		
'		Developer Name					Date						Rev. No.											Reviewer
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'		Pranav Ingle					28-Novs-2013					1
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'		Modified by
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'		Ashwini Patil					14-Jan-2014						1.1													Pranav Ingale
'__________________________________________________________________________________________________________________	
Public Function Fn_SISW_NX_Setup_NXExit(sSaveOption,sMsg)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_NX_Setup_NXExit"
   Dim objExit,bResult,sErrMsg
   Set objExit=Window("NXWindow").Dialog("Exit")
   Set objSave = Window("NXWindow").Dialog("Save")

   If  objExit.Exist(5)=False Then
	   Call Fn_SISW_NX_General_MenuOperation("Select","Exit","")
	   Call Fn_SISW_NX_Setup_ReadyStatusSync(2)
   End If
   If objExit.Exist(5)=True Then
	   If sMsg <> ""  Then
			sErrMsg=Dialog("Exit").Static("ErrMsg").GetROProperty("Text")
			If Trim(sMsg) <> Trim(sErrMsg) Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to verify the message of Exit dialog box ")
				Set ObjDialog=Nothing
				Exit Function
			End If
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verified the message of Exit dialog box ")
		End If
		Fn_SISW_NX_Setup_NXExit = Fn_SISW_NX_UI_ButtonOperation("Fn_SISW_NX_Setup_NXExit", "Click", objExit, sSaveOption)
		'Added Condition for No case by PoonamC_DIPRO NX_NewDevelopment_TC11.5_20180329.00
		If sSaveOption <> "No" Then
			If objSave.Exist = False Then
				Call Fn_UpdateLogFiles("FAIL : Save dialog box does Not  Exist", "FAIL: Save dialog box does Not  Exist")
				Exit Function
			End If
			Fn_SISW_NX_Setup_NXExit=Fn_SISW_NX_UI_ButtonOperation("Fn_SISW_NX_Setup_NXExit", "Click", objSave, sSaveOption)
		End If
	Else
		Fn_SISW_NX_Setup_NXExit=True
   End If
End Function

'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Function Name     			: Fn_SISW_NX_Setup_DisplayResourceBar
' Function Description 			 : To set Display resource bar preference
' Parameters			: 			sType:   Type to set resource bar
' Return Value           :        True/False
' 
' Examples		    	: 			 Fn_SISW_NX_Setup_DisplayResourceBar("On Left")
' History               :  Developer Name			 Date	  				Rev. No. 				Changes					Reviewer
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'								Nilesh Gadekar  	8-Dec-2012 				1.0							
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function Fn_SISW_NX_Setup_DisplayResourceBar(sType)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_NX_Setup_DisplayResourceBar"
   Dim bReturn, dicPrefrenceInfo
 
	' Set Values To Pass values to UserInterface Function
	Set dicPrefrenceInfo=CreateObject("Scripting.Dictionary")
	dicPrefrenceInfo("sAction")="Select"
	dicPrefrenceInfo("sTabCaseName")="UserInterfacePrefrencesTab"
	dicPrefrenceInfo("sTabName")="Layout"
	dicPrefrenceInfo( "sPrefOperation")="displayresourcebar"
	dicPrefrenceInfo( "sWinComboBoxValue")=sType
	dicPrefrenceInfo( "sButton")="OK"

	Select Case sType
		Case "On Left"
			If Window("NXWindow").WinObject("NavigatorTabTitle").Exist(3)=False Then
				bReturn=Fn_SISW_NX_Setup_UserInterFacePrefrences(dicPrefrenceInfo)
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Pass :Successfully Set  Resource Bar  As ["+sType+"]")
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Pass : Resource Bar Already Set As ["+sType+"]")
				bReturn=True
			End If
		Case "As Toolbar"
			If  Window("NXWindow").WinObject("NavigatorTabTitle").Exist(3)=True Then
				bReturn=Fn_SISW_NX_Setup_UserInterFacePrefrences(dicPrefrenceInfo)
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Pass :Successfully Set  Resource Bar  As ["+sType+"]")
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Pass : Resource Bar Already Set As ["+sType+"]")
				bReturn=True
			End If
	
	End Select
	If  bReturn=False	Then
		Fn_SISW_NX_Setup_DisplayResourceBar=False
	Else
		Fn_SISW_NX_Setup_DisplayResourceBar=True
	End If
End Function

'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Function Name		:				Fn_SISW_NX_Setup_UserInterFacePrefrences( dicPrefrenceInfo)

'Description		:		 		 To   Select The Tab From User InterFace Prefrences Dialog and set  the prefrence values

''Parameters		:	 			1. dicPrefrenceInfo :: The Information of the Prefrence

'Return Value		: 				TRUE \ FALSE

'Pre-requisite		:		 		 Nx  and User InterFace Prefrences Dialog should be launched

'Examples			:				 call Fn_SISW_NX_Setup_UserInterFacePrefrences( dicPrefrenceInfo)
'												

'History			:		
'			Developer Name							Date						Rev. No.											Reviewer
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'			Nilesh Gadekar						16-Jan-2013						1.0
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_NX_Setup_UserInterFacePrefrences( dicPrefrenceInfo)

GBL_FAILED_FUNCTION_NAME="Fn_SISW_NX_Setup_UserInterFacePrefrences"
Dim objHierarchy ,ItemsCount,Flag,bResult
set objHierarchy=Window("NXWindow").Dialog("UserInterfacePreferences")
Fn_SISW_NX_Setup_UserInterFacePrefrences=False
Flag=0
Call Fn_SISW_NX_Setup_Restore()
'To Launch The Prefrence Dialog 
If objHierarchy.Exist(5)=False Then
	 If  Fn_SISW_NX_General_MenuOperation("Select","User Interface Preferences", "") Then
		 Flag=1
		 else
		  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Invoke  [ User Interface Preferences ] Dialog")
		  Exit Function
	 End If
End If

' To Set Prefrence Values 
	Select Case  trim(lcase(dicPrefrenceInfo("sAction")))
	Case "select"
		'To Select the TAB 
		bResult=Fn_SISW_NX_Setup_TabOperation( dicPrefrenceInfo("sTabName"),dicPrefrenceInfo("sAction"),objHierarchy, dicPrefrenceInfo("sTabCaseName"))
		If bResult=True Then
		Else
		  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select the tab  [ " +dicPrefrenceInfo("sTabName") + " ] ")
		  Exit Function
		End If

		Select Case trim(lcase(dicPrefrenceInfo("sPrefOperation")))
			Case  "displayresourcebar"
				'To Select  option from Display Resource Bar sWinComboBoxValue
				bResult=Fn_SISW_NX_UI_ComboBoxOperation( "Set",objHierarchy, "DisplayResourceBar","", dicPrefrenceInfo("sWinComboBoxValue"),"")
				If  bResult=True Then
						Fn_SISW_NX_Setup_UserInterFacePrefrences=True
				Else
					  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select the Option  [ " +dicPrefrenceInfo("sWinComboBoxValue") + " ] ComboBox ")
					  Exit Function
				End If
				'To Click On OK Button 
				If Fn_SISW_NX_UI_ButtonOperation( "Fn_SISW_NX_Setup_UserInterFacePrefrences", "Click", objHierarchy, dicPrefrenceInfo("sButton")) Then
				End If

		'Window("NXWindow").MouseMove 5,5	
		'Window("NXWindow").WinTab("NavigatorTab").MouseMove x,y	
			Case Else
		End Select

	Case "verify"

	Case else
			 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : The Select Case ["+dicPrefrenceInfo("sAction") +"]  Not Found ") 
			 Exit Function
    End Select

set objHierarchy=nothing
End Function
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Function Name     			: Fn_SISW_NX_Setup_TabOperation
' Function Description 	    : Function used to Select Tab from Dialog
' Parameters	   				: 	1.sTab: Tab to be select
'											 2.sAction: Action i.e Select/Verify
'											3.objHierarchy: Dialog Hierarchy
'											4.Tab name in OR
'
' Return Value		    : 		True/False
' 	
' Examples		    	: 			bReturn=Fn_SISW_NX_Setup_TabOperation("Layout","select",Window("NXWindow").Dialog("UserInterfacePreferences"),"UserInterfacePrefrencesTab")
' History               :  Developer Name		 Date	  				Rev. No. 				Changes					Reviewer
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'								Nilesh Gadekar  	2-Sep-2013				1.0							
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_NX_Setup_TabOperation( sTab,sAction,objHierarchy, sTabName)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_NX_Setup_TabOperation"
	Dim objTab,icount,counter,Flag,NavigatorTitle
	Call Fn_SISW_NX_Setup_Restore()
	Select Case LCase(sAction)
		Case "select"
			Set objTab=objHierarchy.WinTab(sTabName)
			sActTab=sTab
			If objTab.Exist(5) Then
				sTabList=objTab.GetroProperty("all items")
				aTab=Split(sTabList,vblf,-1,1)
				If Ubound(aTab)>0 Then
					For iCount=0 to Ubound(aTab)-1
						If sActTab=aTab(iCount) Then
							objTab.Select iCount
							Fn_SISW_NX_Setup_TabOperation=True
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS :Successfully selected the  Tab ["+sTabName+"] ")
							Exit For
						End If
					Next
				End IF
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL :Fail to select the  Tab ["+sTabName+"] ")
				Fn_SISW_NX_Setup_TabOperation=False
			End IF
		Case "verify"
			Set objTab=objHierarchy.WinTab(sTabName)
			sActTab=sTab
			If objTab.Exist(5) Then
				sTabList=objTab.GetroProperty("all items")
				aTab=Split(sTabList,vblf,-1,1)
				If Ubound(aTab)>0 Then
					For iCount=0 to Ubound(aTab)-1
						If sActTab=aTab(iCount) Then
							Fn_SISW_NX_Setup_TabOperation=True
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS :Successfully verified the  Tab ["+sTabName+"] ")
							Exit For
						End If
					Next
				End IF
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL :Fail to select the  Tab ["+sTabName+"] ")
				Fn_SISW_NX_Setup_TabOperation=False
			End IF
	End Select
	Set objTab=Nothing
End Function
