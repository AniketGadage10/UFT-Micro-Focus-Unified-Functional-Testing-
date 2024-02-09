Option Explicit

Dim sModule
Dim sPrefName_Reset
Dim sPreVal_Reset
Dim sScope_Reset
Dim bPrefReset
Dim gCacheClear
Dim sSOAUser                    				'Added by Nilesh to reset user level preferences on 27-Nov-12
Dim bCaptureImgVP
Dim sQTPProductName
Dim sUFTProductName
Dim sFeatureName
Dim bGblFuncRetVal
Dim bSiteReset
Dim bConsoleLog			   
sFeatureName=""
sQTPProductName = "HP Unified Functional Testing"
sUFTProductName="Micro Focus Unified Functional Testing"
bPrefReset = True
'--------------------------------------------------------------------------------------
'POC for HC - Global variables for folder path 
'--------------------------------------------------------------------------------------
Public GBL_AUTOMATEDTEST_FOLDER_PATH,GBL_NEWSTUFF_FOLDER_PATH
GBL_AUTOMATEDTEST_FOLDER_PATH="Home:AutomatedTests"
GBL_NEWSTUFF_FOLDER_PATH="Home:Newstuff"

'[TC12-13_7_2017-JotibaT]  Added variable for 4GD as per design change to close extra tab
Dim GBL_4GD_EXTRA_TAB_CLOSER:GBL_4GD_EXTRA_TAB_CLOSER = False

'--------------------------------------------------------------------------------------
'[TC1017-20161101-16_11_2016-VivekA-Maintenance] - As discussed with Dhananjay, Declared Global variable, to prevent deleting (removing) preference in Login function
' - As Login function calls "Fn_KillProcess()" which contains Pref deletion operation for RM Testcases
' - So if we set Pref or create Pref before login using SOA, then it removes it in Login function, to handle this
Dim bGblPrefDeleteAtLogin
bGblPrefDeleteAtLogin = True
'----------------------------------------------------------------------------------------------------------

'POC Jotiba Takkekar***************************************************************************************
Dim GBL_RESET_PERSPECTIVE:GBL_RESET_PERSPECTIVE = True
Dim GBL_FAILED_FUNCTION_NAME:GBL_FAILED_FUNCTION_NAME="" 'Declare global variable to displaying business function name for failure test in batch run result excel
Dim GBL_EXPECTED_MESSAGE:GBL_EXPECTED_MESSAGE=""  'Declare global variable to displaying Expected error message in batch run result excel
Dim GBL_ACTUAL_MESSAGE:GBL_ACTUAL_MESSAGE=""     'Declare global variable to displaying Actual error message in batch run result excel

'POC Vivek Ahirrao***************************************************************************************

'*********************************************************	Function List		***********************************************************************
'0.  Fn_SISW_GetObject(sObjectName)
'1.  Fn_LoadORFile(sActionName, sFilePath)
'2.  Fn_AddRecoveryScenario(sFileName, sScenName, iIndex, sMode)
'3.  Fn_SetActionIterationMode(sMode)
'4.  Fn_CreateLogFile(sLogFileName)
'5.  Fn_ClearCache()
'6.  Fn_InvokeTeamCenter()
'7.  Fn_TeamcenterLogin(StrUserName,StrPassWord, StrGroup,StrRole,StrServer )
'8.  Fn_ReUserTcSession(bCacheClear, bRelaunch, sLoginDetails)
'9.  Fn_SetPerspective(StrModule)
'10. Fn_ResetPerspective()
'11. Fn_MenuOperation(StrAction, StrMenuLabel)
'12. Fn_TeamcenterExit()
'13. Fn_CodeCover_Init()
'14. Fn_CodeCover_Exit()
'15. Fn_SetTCSession()
'16. Fn_KillProcess()
'17. Fn_SetPerspectiveExtn()
'18. Fn_SetMyTcSession()
'19. Fn_Setup_TestcaseInit()
'20. Fn_Setup_TestcaseExit()
'21. Fn_TeamcenterExit_Extn()
'22. Fn_Setup_RandNoGenerate()
'23. Fn_Setup_NetworkDriveOperations()
'24. Fn_GetTestCaseDetailsFromExcel()
'25. Fn_Setup_TestcaseExitWithConsoleLog()
'26. Fn_Setup_BatFileOperations()
'27  Fn_EnableTcExcelAddin()
'28  Fn_Set_ExcelAddinRegistryVal()
'29  Fn_ExcelErrorClose()
'30  Fn_SISW_GetHierarchy()
'31 Fn_SISW_Reload_Addin(sAddinName)
'32 Fn_SISW_Setup_CaptureDesktopImg()
'33 Fn_SISW_Setup_VerifyWinProc()
'34 Fn_SISW_Setup_CompareBitmap()
'35 Fn_SISW_LoadLibrary()
'36 Fn_SISW_Setup_ArrayStringContains()
'37 Fn_SISW_Setup_GetObjectFromXML()
'38 Fn_SISW_MakeIEDefaultBrowser()
'39 Fn_InvokeTeamCenterExt()
'40 Fn_GetFeatureNameOfTestCase()
'41 Fn_Setup_GetActivePerspectiveName
'42 Fn_SecondsToMinutes
'43 Fn_SISW_ConvertToItemID 
'44 Fn_SISW_ImgComp_Operations(ActualBmp,ExpectedBmp)
'*********************************************************	Function List		***********************************************************************
'****************************************    Function to get FeatureName of TestCase ***************************************
'
'Function Name		 	:	Fn_GetFeatureNameOfTestCase
'
'Description		    :  	Function to get FeatureName of TestCase.
'
'Parameters		    :	null
'								
'Return Value		    :  	Feature Name/Nothing (Global variable)
'
'Examples		     	:	Fn_GetFeatureNameOfTestCase()
'
'History:
'		Developer Name			  Date			Rev. No.		Reviewer		Changes Done	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'		Chandrakant Tyagi	   25-5-2015		 1.0		  Vivek Ahirrao		
'	
'-----------------------------------------------------------------------------------------------------------------------------------
Function Fn_GetFeatureNameOfTestCase()
	Dim colCount, iterator, colName
	Datatable.SetCurrentRow(1)
	colCount=DataTable.GetSheet("Global").GetParameterCount
	For iterator = 1 To colCount Step 1
		colName=Datatable.GetSheet("Global").GetParameter(iterator).Name
		If Lcase(colName)="feature" Then
		   	If Instr(1,Lcase(DataTable.Value(colName)),Lcase("RequirementsManagement")) > 0 Then
				sFeatureName="REG - RequirementsManagement"
			   	Exit for
			Else
				sFeatureName=DataTable.Value(colName)
				Exit for	
		   	End If			
		End If
	Next
End Function

'****************************************    Function to get Object hierarchy ***************************************
'
''Function Name		 	:	Fn_SISW_GetObject
'
''Description		    :  	Function to get specified Object hierarchy.

''Parameters		    :	1. sObjectName : Object Handle name
								
''Return Value		    :  	Object \ Nothing
'
''Examples		     	:	Fn_SISW_GetObject("Remove")

'History:
'	Developer Name			Date			Rev. No.		Reviewer		Changes Done	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Koustubh Watwe		 1-June-2012		1.0				
'	Snehal Salunkhe	   22-June-2012				Pranav S.
'-----------------------------------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------------------------------
'	Anurag Khera		 26-Sept-2013		2.0								Externalized object hierarchies
'-----------------------------------------------------------------------------------------------------------------------------------
Function Fn_SISW_GetObject(sObjectName)
	Dim sObjectXMLPath
	sObjectXMLPath = Fn_GetEnvValue("User", "AutomationDir") & "\TestData\AutomationXML\ObjectXML\Setup.xml"
	Set Fn_SISW_GetObject = Fn_SISW_Setup_GetObjectFromXML(sObjectXMLPath, sObjectName)
End Function
'*********************************************************		Function to Creates Log File and folder.		***********************************************************************
'Function Name		:				Fn_LoadORFile(sActionName, sFilePath)

'Description			 :		 		 Loads shared Object Repository for specified Action

'Parameters			   :	 			1. sActionName : Name of the Action for which Object Repository to be assigned

'													2.	sFilePath	-  Path of the Object Repository file

'Return Value		   : 			True/False

'Pre-requisite			:		 	Nothing

'Examples				:			Call Fn_LoadORFile("Action1", "C:\Automation\General.tsr")

'History					 :		
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'												------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Vallari		 													   22/04/2010			              1.0										Created
'												------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_LoadORFile(sActionName, sFilePath)
		Dim App
		Dim qtRepositories

		Fn_LoadORFile = TRUE
'		sActionName = Environment("ActionName")
		Set App = CreateObject("QuickTest.Application")
		Set qtRepositories = App.Test.Actions(Environment("ActionName")).ObjectRepositories
		If (qtRepositories.Find(sFilePath) = -1) Then ' If the repository cannot be found in the collection 
			If ((qtRepositories.Add(sFilePath, 1)) = -1) Then ' Add the repository to the collection 
				Services.LogMessage "Failed to load GUI file ["  &sFilePath &"]", ErrorMsg 
				Fn_LoadORFile = FALSE			
			End If
		End If
		Set qtRepositories = Nothing ' Release the action's shared repositories collection 
		Set App = Nothing ' Release the App
End Function


'*********************************************************		Function to Creates Log File and folder.		***********************************************************************
'Function Name		:				Fn_AddRecoveryScenario(sFileName, sScenName, iIndex, sMode)

'Description			 :		 		 Adds pre-recorded recovery scenario to the script

'Parameters			   :	 			1. sFileName : Name of the File where Recovery Scenario is recorded

'													2.	sScenName	-  Scenario Name

'													3.	iIndex - Position of the Scenario for the script

'													4.	sMode - Mode to invoke Recovery

'Return Value		   : 			True/False

'Pre-requisite			:		 	Nothing

'Examples				:			Call Fn_AddRecoveryScenario("C:\Automation\Recovery.qrs", "FCC_Cache", 1, "onError")
'
'History					 :		
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'												------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Vallari		 													   22/04/2010			              1.0										Created
'												------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_AddRecoveryScenario(sFileName, sScenName, iIndex, sMode)
	Dim qtAppRec, qtTestRecovery

	On Error Resume Next

	Set qtAppRec = CreateObject("QuickTest.Application")
	Set qtTestRecovery = qtAppRec.Test.Settings.Recovery
	qtTestRecovery.SetActivationMode sMode
	
	qtTestRecovery.Add sFileName, sScenName, iIndex
	
	Set qtTestRecovery = Nothing ' Release the Recovery object
	Set qtAppRec = Nothing ' Release the Application object
End Function


'*********************************************************		Function to Creates Log File and folder.		***********************************************************************
'Function Name		:				Fn_SetActionIterationMode(sMode)

'Description			 :		 		Sets Action iteration as per the Data in the datatable

'Parameters			   :	 			1. sMode : Mode for execution

'Return Value		   : 			True/False

'Pre-requisite			:		 	Nothing

'Examples				:			Call Fn_SetActionIterationMode("oneIteration")
'
'History					 :		
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'												------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Vallari		 													   22/04/2010			              1.0										Created
'												------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public function Fn_SetActionIterationMode(sMode)
	Dim qtAppItrMod

	On Error Resume Next

	Set qtAppItrMod = CreateObject("QuickTest.Application")

	qtAppItrMod.Test.Settings.Run.IterationMode = sMode
    qtAppItrMod.Options.Run.RunMode = "Normal"

	Set qtAppItrMod = Nothing
	
	'Added by Nilesh for RM excel Test cases on 1st June 2012
	If DataTable("Feature", dtGlobalSheet)="" Then 'Added by shrikant
		 ' Do nothing
	Else
		 If Instr(Lcase(DataTable("Feature", dtGlobalSheet)),"requirement")>0 Then
			Call Fn_EnableTcExcelAddin()'Enable Excel TC addin
		End If
	End If
	'Handle Restart application dialog in windows 7
	Call Fn_ExcelErrorClose()

End Function


''*********************************************************		Function to Kill all the Processes		***********************************************************************
'Function Name		:				Fn_KillProcess

'Description			 :		 		 The function is used to kill all the process which is passed by the user in a comma seprated string.

'Parameters			   :	 			1. sProcessToKill.
											
'Return Value		   : 				Nothing.

'Pre-requisite			:		 		Nothing.

'Examples				:				 Fn_KillProcess("iexplore.exe,notepad.exe") 

'History:
'	Developer Name			Date					Rev. No.			Changes Done			Reviewer	Reviewed  Date
'--------------------------------------------------------------------------------------------------------------------------------
'	Vallari		 			23-Apr-2010       		1.0					
'--------------------------------------------------------------------------------------------------------------------------------
'	Koustubh W		 		10-Apr-2011       		1.0					Modified code to reset preference using SOA
'--------------------------------------------------------------------------------------------------------------------------------
'	Sandeep N		 		15-Jun-2011       		1.1					Modified code to Handle BMIDE window
'--------------------------------------------------------------------------------------------------------------------------------
'	Ashwini Patil		 	20-Mar-2014       		1.2					Added Code to handle modified QTP Release
'--------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_KillProcess(sProcessToKill)
			Dim sArrData,  iCount, strComputer,sImgPath
			Dim objWMIService, objProcess, colProcess
'			Dim sSOAUser
			Dim aPrefName
			Dim aPrefVal
			Dim aPrefScope
			Dim objTcWin_1, iWinCnt_1, i_1, sWinTitle,objTCWin
			Dim bReturn, aSOAPerfInput
			Dim App,userVarPath,bFlag,objOptionDialog
			Dim Count, Iterator, sRepository
			Dim sCreatedPrefrenceName, sAutoDir
			'- - - - - - - - - - - - - - - - - - - - - - - - Added Code to handle modified QTP Release - - - - - - - - - - - - - - - - - - - -
			Set App = CreateObject("QuickTest.Application")
			If CDbl(App.Version) > 11 Then
				Count=App.Test.Actions(Environment("ActionName")).ObjectRepositories.Count() 
				bFlag = False
        		For Iterator = 1 To Count 
        			sRepository=App.Test.Actions(Environment("ActionName")).ObjectRepositories.Item(Iterator)
					If instr(1,sRepository,"General.tsr") Then
						bFlag = True
						Exit For	
					Else
						bFlag = False				
					End If	        	
	        	Next
	      	
			Else
				userVarPath	= Fn_GetEnvValue("User", "AutomationDir")
				bFlag=App.Test.Actions(Environment("ActionName")).ObjectRepositories.Find(userVarPath&"\ObjectRepository\General.tsr")
				'- - - - - - - - - - - - - - - - - - - - - - - - Added Code to handle BMIDE Scenarion - - - - - - - - - - - - - - - - - - - -
        	End IF
        	Set App = Nothing
			If Cint(bFlag)<>-1 or bFlag = True Then	
				'Added By Nilesh on 1st June 2012
				Call Fn_ExcelErrorClose()
				'End
				If  sProcessToKill = "" Then
					sProcessToKill = Environment.Value("KillProcesses")                        
				End If
				
				'Restore Image of the Application before killing Application.
				sImgPath = Environment.Value("BatchFldName") +"\"   + Environment.Value("TestName") + ".png"
				Desktop.CaptureBitmap sImgPath,True
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Application Image is stored at [" + sImgPath + "]")				

				'Added by Vallari - Stop Code Coverage
				Call Fn_CodeCover_Exit()
				
				'Call Tc_Exit before killing processes
'				Call Fn_TeamcenterExit_Extn()
'				wait(5)

				'Exit Teamcenter Sessions & Not BMIDE
                Set objTcWin_1 = description.Create()

				objTcWin_1("Class Name").Value = "JavaWindow"
				objTcWin_1("title").Value = ".*Teamcenter.*"
				objTcWin_1("title").RegularExpression = True
				'wait(2)
'				iWinCnt_1 = objTcWin_1.Count
				iWinCnt_1 = desktop.ChildObjects(objTcWin_1).Count
				Set objTCWin=desktop.ChildObjects(objTcWin_1)
				For i_1 = 0 to iWinCnt_1 - 1
'					sWinTitle = JavaWindow("DefaultWindow").GetROProperty("title")
					sWinTitle = objTCWin(i_1).GetROProperty("title")
					If instr(sWinTitle, "Business Modeler IDE") > 0 OR trim(sWinTitle) = "" Then
						'Don't do anything on this window
					Else
						JavaWindow("DefaultWindow").SetTOProperty "title", sWinTitle
						'*Migrated from TC10_1 to Mainline by Nilesh Gadekar on 8-Jan-2013
						'Code added by Nilesh on 28-Sep-2012 to handle Option Dialog (To avoid "Instance in Use" Error while performing various Preference operation)
						Set objOptionDialog=Fn_SISW_GetObject("IndexOptions")
						If objOptionDialog.Exist(1)=True Then
							objOptionDialog.Restore()
							If  objOptionDialog.JavaButton("Cancel").Exist(1)Then
								If objOptionDialog.JavaButton("Cancel").GetRoProperty("enabled")=1 Then
									objOptionDialog.JavaButton("Cancel").Click
									Wait 1
								End If
								objOptionDialog.Close()
							End If
							If  objOptionDialog.JavaButton("Save").Exist(1)Then
								objOptionDialog.Restore()
								If objOptionDialog.JavaButton("Save").GetRoProperty("enabled")=1 Then
									objOptionDialog.JavaButton("Save").Click
									Wait 1
								End If
									objOptionDialog.Close()
							End If
						End If
						Set objOptionDialog=Nothing
						'End
						'JavaWindow("DefaultWindow").Maximize - by Vallari [27/1/2012] This call doesn't get executed when there is any child dialog on Tc main win
						Call Fn_TeamcenterExit_Extn()
						'wait(2)        
					End If
					JavaWindow("DefaultWindow").SetTOProperty "title", ".* - Teamcenter .*"
				Next
				Set objTcWin_1 = Nothing
				Set objTCWin=Nothing
				'Kill Teamcenter.exe for killing Tc Sessions for which any modal dialog is up
				sArrData = split(sProcessToKill, ":",-1,1)

				strComputer = "." 
				Set objWMIService = GetObject("winmgmts:"& "{impersonationLevel=impersonate}!\\"& strComputer & "\root\cimv2") 

				For iCount = 0 to ubound(sArrData)
					Set colProcess = objWMIService.ExecQuery("Select * from Win32_Process Where Name ='"+sArrData(iCount)+"'")  
					'For Each objProcess in colProcess 
					For Each objProcess in colProcess 
						objProcess.Terminate() 
					Next 
				Next

				'- - - - - - - - - - - - - - - - - - - - - - - - Added Code to handle BMIDE Scenarion - - - - - - - - - - - - - - - - - - - -
				JavaWindow("DefaultWindow").SetTOProperty "index",0
				If JavaWindow("DefaultWindow").Exist(1) Then
					sWinTitle=JavaWindow("DefaultWindow").GetROProperty("title")
					If instr(sWinTitle, "Business Modeler IDE") Then
						For iCount=0 to 1
							JavaWindow("DefaultWindow").SetTOProperty "index",1
							sWinTitle=JavaWindow("DefaultWindow").GetROProperty("title")
							If instr(sWinTitle, "Business Modeler IDE") > 0 or sWinTitle="" Then
								'Don't do anything on this window
							Else
								'wait(2)
								systemutil.CloseProcessByWndTitle sWinTitle,False
								'Closes RCAF Console
								systemutil.CloseProcessByWndTitle "RCAF Console",True
							End If
							'Closes Teamcenter popup, which is invokes occasionaly after closing Tc application			
						Next	
					Else
						'wait(2)
						systemutil.CloseProcessByWndTitle ".*Teamcenter.*",True			
						'Closes RCAF Console
						systemutil.CloseProcessByWndTitle "RCAF Console",True
					End If
				Else
					'wait(2)
					systemutil.CloseProcessByWndTitle ".*Teamcenter.*",True			
					'Closes RCAF Console
					systemutil.CloseProcessByWndTitle "RCAF Console",True
					systemutil.CloseProcessByWndTitle "TAO ImR",False
				End If
				If JavaWindow("Teamcenter Login").Exist(1) Then
					'wait(1)
					systemutil.CloseProcessByWndTitle "Teamcenter Login",False
				End If
				'- - - - - - - - - - - - - - - - - - - - - - - - Added Code to handle BMIDE Scenarion - - - - - - - - - - - - - - - - - - - -
				If Window("Class Name:=Window","text:=TAO ImR").Exist(5) Then
					Window("Class Name:=Window","text:=TAO ImR").Close	
				End If

				Set objWMIService = Nothing
				Set colProcess = Nothing

				If Err.Number <> 0 Then
					Fn_KillProcess =False
					'Call Fn_WriteLogFile("Fn_ClearCache()", 1, Err.Number ,"FAIL : Clear Cache Operation Failed")
				Else
					Fn_KillProcess =True
					'Call Fn_WriteLogFile("Fn_ClearCache()", 3, Err.Number ,"PASS : Clear Cache Operation Passed")
				End If

'---------------------------------------------------------------------------------------------------------
'Added by Chandrakant Tyagi to deal with the prefrence of Requirement Manager 25-5-2015
				If sFeatureName="REG - RequirementsManagement" Then
						If gPrefName <> "" Then							
							sCreatedPrefrenceName = split(gPrefName,":",-1,1)
							sAutoDir = Fn_GetEnvValue("User", "AutomationDir")
							If UBound(sCreatedPrefrenceName) > -1 Then
								For Iterator = 0 To UBound(sCreatedPrefrenceName)
							  		sPrefrenceName = Fn_GetXMLNodeValue(sAutoDir + "\TestData\AutomationXML\CreatedPreferenceXML\RM_Created_Prefrence.xml", Trim(sCreatedPrefrenceName(Iterator)))
									If sPrefrenceName <> False Then
										sPrefName_Reset=Trim(gPrefName)
										sPreVal_Reset=gPrefValue
										sScope_Reset=gPrefScope												
										Exit for
									End If				  		
								Next
							End If
						End If	
			
						If sPrefName_Reset = "PTN_Default_Partition_RevRule" AND sPreVal_Reset = "Any Status; No Working" AND lcase(Scope_Reset) = "site" Then 
						'' Do nothing
						Elseif sPrefName_Reset <> "" And CBool(bPrefReset) Then
								If  sSOAUser="" Then   
									sSOAUser = "TcUserDBA"
								End If
								aPrefName = Split(sPrefName_Reset, ":", -1, 1)
								aPrefVal = Split(sPreVal_Reset, ":", -1,1)
								aPrefScope = Split(sScope_Reset, ":", -1,1)
								if uBound(aPrefName) <> uBound(aPrefVal) OR inStr(sPreVal_Reset,"~") > 0 then
									aPrefName = Split(sPrefName_Reset, "~", -1, 1)
									aPrefVal = Split(sPreVal_Reset, "~", -1,1)
									aPrefScope = Split(sScope_Reset, "~", -1,1)
								End If
								sPrefName_Reset = ""
								sPreVal_Reset = ""
								sScope_Reset = ""
								For iCount = 0 to Ubound(aPrefName)										
									If gPrefName <> "" Then
										'To handle RM testcases scenerio
										If bGblPrefDeleteAtLogin = True Then
											aSOAPerfInput = Array(sSOAUser, "REMOVEEPREFERENCE", aPrefName(iCount),  aPrefScope(iCount), aPrefVal(iCount))
											bReturn = Fn_SOA_PrefOperation(aSOAPerfInput)
										End If
									Else
										aSOAPerfInput = Array(sSOAUser, "SetMultiValuePreference", aPrefName(iCount),  aPrefScope(iCount), aPrefVal(iCount))
										bReturn = Fn_SOA_PrefOperation(aSOAPerfInput)
									End If		
								Next
							End If
'----------------------------------------------------------------------------------------------------------------------------------					
				Else					
						'Reset Preference Using SOA
						If sPrefName_Reset = "PTN_Default_Partition_RevRule" AND sPreVal_Reset = "Any Status; No Working" AND lcase(Scope_Reset) = "site" Then 
						'' Do nothing
						Elseif sPrefName_Reset <> "" And CBool(bPrefReset) Then
		'					sSOAUser = "TcUserDBA"
							If  sSOAUser="" Then   'Added by Nilesh to reset user level preference on 27-Nov-12
								sSOAUser = "TcUserDBA"
							End If
							aPrefName = Split(sPrefName_Reset, ":", -1, 1)
							aPrefVal = Split(sPreVal_Reset, ":", -1,1)
							aPrefScope = Split(sScope_Reset, ":", -1,1)
							if uBound(aPrefName) <> uBound(aPrefVal) OR inStr(sPreVal_Reset,"~") > 0 then
								aPrefName = Split(sPrefName_Reset, "~", -1, 1)
								aPrefVal = Split(sPreVal_Reset, "~", -1,1)
								aPrefScope = Split(sScope_Reset, "~", -1,1)
							End If
							sPrefName_Reset = ""
							sPreVal_Reset = ""
							sScope_Reset = ""
							For iCount = 0 to Ubound(aPrefName)
								aSOAPerfInput = Array(sSOAUser, "SetMultiValuePreference", aPrefName(iCount),  aPrefScope(iCount), aPrefVal(iCount))
								bReturn = Fn_SOA_PrefOperation(aSOAPerfInput)
								'bReturn = Fn_SOA_SetPreference(sSOAUser, aPrefName(iCount), aPrefVal(iCount), aPrefScope(iCount))
							Next
						End If
				End If
				Call Fn_ClearCache()
            Else
				'Restore Image of the Application before killing Application.
				sImgPath = Environment.Value("BatchFldName") +"\"   + Environment.Value("TestName") + ".png"
				Desktop.CaptureBitmap sImgPath,True
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Application Image is stored at [" + sImgPath + "]")		
			End If
			
			'=============================================================================================
			' Code to Reset Site1 details as default site  : Added code for CMS TCs : PoonamC_02May2018
			If bSiteReset = True Then
				userVarPath = Fn_GetXMLNodeValue(Environment.Value("sPath") & "\TestData\Sites.xml","Site1")
				Call Fn_SISW_CMS_SiteOperations("Set",userVarPath,"")
				Wait 1
				bSiteReset = False
			End If
			'=============================================================================================
End Function


'*********************************************************		Function Cancels Check Out		***********************************************************************
'Function Name		:				Fn_ClearCache

'Description			 :		 		 Function to clear chache

'Parameters			   :	 			NA
											
'Return Value		   : 				TRUE \ FALSE

'Pre-requisite			:		 		NA

'Examples				:				Fn_ClearCache()

'History:
'										Developer Name			Date			Rev. No.			Changes Done			Reviewer		Reviewed Date	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Sameer						25-Mar-2010		1.0																	Santosh				25-Mar-10 											
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function Fn_ClearCache()

		On error resume next
		Dim objNetwork, objFSO, objFolder, objSubFolder, objFiles
		Dim sDrive, sUserName, sPath 
		Dim WshShell
		'Creates function for Network and File System
		Set objNetwork =CreateObject("WScript.Network")
		Set objFSO = CreateObject("Scripting.FileSystemObject")
		Set WshShell =CreateObject("WScript.Shell")
		sPath=WshShell.ExpandEnvironmentStrings("%USERPROFILE%")
		'Object of User folder
		Set objFolder = objFSO.GetFolder(sPath)
		'Objects of subfolders and files within User folder
		Set objSubFolder = objFolder.SubFolders
		'Deletes folders and files 
		If objFSO.FolderExists(sPath & "\Teamcenter") Then
						objSubFolder.Item("Teamcenter").Delete
		End If
		'Clears out objects
		Set objNetwork = nothing
		Set objFSO = nothing
		Set objFolder = nothing
		Set objSubFolder = nothing
		If Err.Number <> 0 Then
						Fn_ClearCache =False
						'Call Fn_WriteLogFile("Fn_ClearCache()", 1, Err.Number ,"FAIL : Clear Cache Operation Failed")
		Else
						Fn_ClearCache =True
						'Call Fn_WriteLogFile("Fn_ClearCache()", 3, Err.Number ,"PASS : Clear Cache Operation Passed")
		End If
End Function


'*********************************************************		Function to invoke the application	************************************************************************

'Function Name		:		Fn_InvokeTeamCenter()

'Description			:		 This function invokes the team center application

'Parameters			  :	 			

'Return Value		   : 		The String which represents the result : "PASS" or "FAIL" with the reason

'Pre-requisite			:		  The Team Centre Application 2007 should be installed

'Examples				:		  FnInvokeTeamCenter() will invoke team center application

'History					 :		
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'												------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													
'												------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function Fn_InvokeTeamCenter()																																		 
	GBL_FAILED_FUNCTION_NAME="Fn_InvokeTeamCenter"
   On Error resume next
	Dim sPath, sModuleName, sAutoDir,sLanguage
	
	sLanguage = ""
	sModuleName = ""
	sPath = Environment.Value("AppExecutable")
	
	'========= Added Code for DIPRO to luanch portal.bat with language specified  : Added by PoonamC_05June2018 ==========
	Select Case Environment.Value("Locale")
		Case "English"
			sLanguage = " -nl en_US"
		Case else
		 	sLanguage = ""
	End Select	
	'======================================================================================================================
	
	If sModule <> "" Then
		sAutoDir = Fn_GetEnvValue("User", "AutomationDir")
		sModuleName = Fn_GetXMLNodeValue(sAutoDir + "\TestData\TcModules.xml", sModule)
	Else
		sModuleName = "com.teamcenter.rac.gettingstarted.GettingStartedPerspective"
	End If

	SystemUtil.Run sPath, "-perspective " & sModuleName & sLanguage
	'SystemUtil.Run sPath

	If JavaWindow("Teamcenter Login").Exist(iTimeOut) Then						  							
			Fn_InvokeTeamCenter = TRUE  																						
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Invoked Teamcenter Application from [" + sPath + "]")
	Else
			 Fn_InvokeTeamCenter = FALSE
			 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Invoke Teamcenter Application from [" + sPath + "]")
			 Exit Function
	End If

End Function


''*********************************************************		Function to login into Teamcenter		***********************************************************************
'Function Name		:				Fn_TeamcenterLogin

'Description			 :		 		 This function logins into Teamcenter

'Parameters			   :	 			1. StrUserName: This is the username to be used for login
'													 2. StrPassWord: This is the password corresponding to the username
'													3. StrGroup: Group of the user logging in
'													4. StrRole: Assigned role of the user from the specified Group
'													5. Server: Database to be connected

'Return Value		   : 				TRUE \ FALSE

'Pre-requisite			:		 		All passed in parameters should be correct	  

'Examples				:				 Fn_TeamcenterLogin("infodba","infodba", "dba","DBA")

'History					 :		
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'												------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Jhuma Debnath										08/03/2010			              1.0										Created
'													Vallari S.														10/03/2011						2.0											Modified	Updated function for PR#6497348
'													Sandeep N.													15/03/2013						3.0										 Added condition to check existance of  : JavaWindow("Teamcenter Login").JavaWindow("Login") dialog
'													Ashwini P											30/04/2014						   4.0								When Role and group is spaces to enable Login button for Teamcenter login window.		
'												------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function Fn_TeamcenterLogin(StrUserName,StrPassWord, StrGroup,StrRole,StrServer )
	GBL_FAILED_FUNCTION_NAME="Fn_TeamcenterLogin"
	Dim objWindowTeamCenterLogin,objWindowLogin
	Dim bReturn, bSessionGrpErr
	Dim bExist, iIterate

	bSessionGrpErr = False
	On Error Resume Next

	'Check the existence of Login window
	Set objWindowTeamCenterLogin = JavaWindow("Teamcenter Login")
	If Not objWindowTeamCenterLogin.Exist(20) Then
        Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Find [Teamcenter Login] Window of Function Fn_TeamcenterLogin" )
		Set objWindowTeamCenterLogin = Nothing
		Fn_TeamcenterLogin = False
		Exit Function
	End If
	
	objWindowTeamCenterLogin.JavaButton("Clear").Click micLeftBtn					'Click on "Clear" button
		
		'Username and password cannot be null
		If  StrUserName <> "" And StrPassWord <> "" Then
'			If StrGroup <> "" and StrRole <> "" then
'				objWindowTeamCenterLogin.JavaEdit("User ID:").Set Trim(StrUserName)	 					'Enter "Username"
'				objWindowTeamCenterLogin.JavaEdit("Password:").Set Trim(StrPassWord)					'Enter "Password"	
'			Else
		'If Role and Group is null to Enable 'Login' button	
				objWindowTeamCenterLogin.JavaEdit("User ID:").Set ""
				objWindowTeamCenterLogin.JavaEdit("User ID:").Type Trim(StrUserName)					'Enter "Username"
				objWindowTeamCenterLogin.JavaEdit("Password:").Set ""
				objWindowTeamCenterLogin.JavaEdit("Password:").Type Trim(StrPassWord)					'Enter "Password"
'			End If	
			
			If StrGroup <> "" Then
				objWindowTeamCenterLogin.JavaEdit("Group:").Set Trim(StrGroup)  						'Enter "Group"
            End If
			If StrRole <> "" Then
				objWindowTeamCenterLogin.JavaEdit("Role:").Set Trim(StrRole)							'Enter Role" 	
			End If
			If StrServer <> "" Then
				objWindowTeamCenterLogin.JavaEdit("Server:").Set Trim(StrServer)				  'Select "Server"
			End If
			If Err.Number < 0 Then
				Fn_TeamcenterLogin = FALSE
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Set Login Credentials")		
			End If
		Else
			Fn_TeamcenterLogin = FALSE
            Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Username or  password cannot be empty")		
		End If

		If objWindowTeamCenterLogin.JavaButton("Login").GetROProperty("enabled") = 1 Then
			objWindowTeamCenterLogin.JavaButton("Login").Click micLeftBtn		'Click on "Login" button
		Else
			Fn_TeamcenterLogin =FALSE
            Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL:Login Button Disabled")
		End If

		'If  FCC error occur click on "OK" button
        'If Fn_UI_ObjectExist("Fn_TeamcenterLogin",objWindowTeamCenterLogin.JavaWindow("Login"))=True Then
		''Adde by Vallari - 30-Mar-2012
		
		If gCacheClear Then
			Dim iSynWin, objSynchWin
			iSynWin = 0
	'		Do While objWindowTeamCenterLogin.JavaWindow("SynchronizingInstalledFiles").Exist
			If JavaWindow("SynchronizingInstalledFiles").Exist(5) Then
				Set objSynchWin = JavaWindow("SynchronizingInstalledFiles")
			ElseIf JavaWindow("Teamcenter Login").JavaWindow("SynchronizingInstalledFiles").Exist(5) Then
				Set objSynchWin = JavaWindow("Teamcenter Login").JavaWindow("SynchronizingInstalledFiles")
			Else
				Set objSynchWin = JavaWindow("Teamcenter Login")
			End If
			For i=1 to 30
				If objSynchWin.Exist(1) Then
					iSynWin = iSynWin + 1
					wait 1
				Else
					Exit For
				End If
			Next
			Set objSynchWin = Nothing
			gCacheClear = False
		End If
'		Loop
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[Synchronizing installed files] popup was displayed for [" + iSynWin + "] Seconds")
        
		Set objWindowDefaultWindow = JavaWindow("DefaultWindow")
		bExist = False
		
		For iIterate = 1 to 50
			If objWindowDefaultWindow.Exist(1) Then
				If instr(JavaWindow("DefaultWindow").GetROProperty("title"), "Business Modeler IDE") = 0  Then
					bExist = True
					Exit For
				End If
			ElseIf objWindowTeamCenterLogin.JavaWindow("Login").Exist(1) Then
				'++++++++<<<<<<<<<<Code to handle the invalid login error msgs.>>>>>>>>>>+++++++++
				'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
				Dim ObjDialog, Desc, ObjDesc
				Dim iCounter,  bMsgFound
				'Set ObjDialog =  Fn_UI_ObjectCreate("Fn_TeamcenterLogin", JavaWindow("Teamcenter Login").JavaWindow("Login") )
				Set ObjDialog =  JavaWindow("Teamcenter Login").JavaWindow("Login")
				Set Desc = Description.Create()
				Desc("to_class").Value = "JavaStaticText"
				Set ObjDesc = ObjDialog.ChildObjects(Desc)
				bMsgFound = False
				For iCounter=0 to(ObjDesc.Count) -1
					 If instr(ObjDesc(iCounter).Object.text , "The given group is unknown") > 0 _
					 OR instr(ObjDesc(iCounter).Object.text , "either the user ID or the password is invalid") > 0 _
					 OR instr(ObjDesc(iCounter).Object.text , "The role entered is invalid") > 0 _
					 OR instr(ObjDesc(iCounter).Object.text , "The specified user is inactive") > 0 Then
							 bMsgFound = True
							 Exit for 
					 Elseif instr(ObjDesc(iCounter).Object.text , "Login service group does not match") > 0 then
							bSessionGrpErr = True
							Exit for
					 End If 
				Next
				If bMsgFound = False Then
					If objWindowTeamCenterLogin.JavaWindow("Login").Exist(1) Then
						Call Fn_Button_Click("Fn_TeamcenterLogin",objWindowTeamCenterLogin.JavaWindow("Login"),"OK")
					End If
				ElseIf  instr(ObjDesc(iCounter).Object.text , "operation was successful but had some warning") > 0  Then
					'Call Fn_Button_Click("Fn_TeamcenterLogin",objWindowTeamCenterLogin.JavaWindow("Login"),"OK")
					bSessionGrpErr = True
				End If
				Set Desc = Nothing
				Set ObjDesc = Nothing
				Set ObjDialog = Nothing

				'Work-around for PR#6497348
				If bSessionGrpErr Then
					'Dismiss Login error dialog
					If objWindowTeamCenterLogin.JavaWindow("Login").Exist(1) Then
						Call Fn_Button_Click("Fn_TeamcenterLogin",objWindowTeamCenterLogin.JavaWindow("Login"),"OK")
					End if
					'Try to login only with UserName & Password
					If objWindowTeamCenterLogin.Exist(5) Then
						objWindowTeamCenterLogin.JavaEdit("User ID:").Set Trim(StrUserName)	 					'Enter "Username"
						objWindowTeamCenterLogin.JavaEdit("Password:").Set Trim(StrPassWord)				'Enter "Password"
						objWindowTeamCenterLogin.JavaEdit("Group:").Set ""															  'Enter "Group" as blank
						objWindowTeamCenterLogin.JavaEdit("Role:").Set ""																'Enter Role" as blank	
						objWindowTeamCenterLogin.JavaButton("Login").Click micLeftBtn										'Click on "Login" button
					End If
				End If
				'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
				If bSessionGrpErr or bMsgFound Then
					Exit For
				End If
			End If
		Next

		If bExist = False Then
			For iIterate = 1 to 50
				If objWindowDefaultWindow.Exist(1) And instr(JavaWindow("DefaultWindow").GetROProperty("title"), "Business Modeler IDE") = 0 Then
					bExist = True
					Exit For
				'Else
					wait 1
				End If
			Next
		End If
		
		'If objWindowDefaultWindow.Exist(600) And instr(JavaWindow("DefaultWindow").GetROProperty("title"), "Business Modeler IDE") = 0 Then	'Exclude BMIDE Window
		If bExist Then
			objWindowDefaultWindow.Maximize
			'Added by Vallari - For Tracking Tc Code Coverage
			Call Fn_CodeCover_Init()
			'After logging in change session settings as required in case of session setting error
			If bSessionGrpErr Then
				bReturn = Fn_UserSessionSettings("Set", StrGroup, StrRole, "", "", "", "", "", "", "")
				If bReturn = False Then
						Fn_TeamcenterLogin = False			
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Log in with User [" +StrUserName + "] in Group/Role [" + StrGroup + " / " + StrRole + "]")
				Else
					Fn_TeamcenterLogin = TRUE			
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully Logged in with User [" +StrUserName + "] in Group/Role [" + StrGroup + " / " + StrRole + "]")
				End If
			Else
				Fn_TeamcenterLogin = TRUE			
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully Logged in with User [" +StrUserName + "]")
			End If
		Else
			Fn_TeamcenterLogin = False			
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Log in with User [" +StrUserName + "]")
		End If

		Set objWindowTeamCenterLogin=Nothing
		Set objWindowLogin=Nothing
		Set objWindowDefaultWindow = Nothing

End Function

''*********************************************************		Function to Check Exisitng Teamcenter	Session	***********************************************************************
'Function Name		:				Fn_ReUserTcSession(bCacheClear, bRelaunch, sLoginDetails)

'Description			 :		 		 This function logins into Teamcenter

'Parameters			   :	 			1. bCacheClear: Boolean to be set to True if Test Script specifically requires clearing of Cache
'													 2. bRelaunch: Boolean to be set to True if Test Script specifically requires New Tc Session
'													3. sLoginDetails: ":" separated login data

'Return Value		   : 				TRUE \ FALSE

'Pre-requisite			:		 		All passed in parameters should be correct	  

'Examples				:				 Fn_ReUserTcSession(False, False, "AutoTest1:AutoTest1:Engineering:Designer:")

'History					 :		
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'												------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Vallari																27/04/2010			              1.0										Created
'												------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function Fn_ReUserTcSession(bCacheClear, bRelaunch, sLoginDetails)
			GBL_FAILED_FUNCTION_NAME="Fn_ReUserTcSession"
			Dim bReuse, bReturn, sSession, aLogin, bRefreshValue
	
			On Error Resume Next
			bConsoleLog = 0
			'To handle RM testcases scenerio
			bGblPrefDeleteAtLogin = False

			aLogin  =Split(sLoginDetails,":",-1,1)
			bReuse = True

			If bCacheClear = True Then
					bRelaunch = True
					'Set a Global variable
					gCacheClear = True
			End If

			'************* If bCacheClear = True, Kill Tc session and Relaunch after Clearing Cache ****************
			If CBool(bRelaunch) =True Then
					'Set the bReuse flag to False, as bRelaunch is True
					bReuse = False
					'Kill Teamcenter Application, if it exists
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[" + Cstr(now) + "] Teamcenter Kill Process Operation: Start")
					bReturn = Fn_KillProcess("")
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[" + Cstr(now) + "] Teamcenter Kill Process Operation: Start")

					If CBool(bCacheClear) =True Then
							'Call a Function to Clear Cache
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[" + Cstr(now) + "] Teamcenter Cache Clear Operation: Start")
							Call Fn_ClearCache()
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[" + Cstr(now) + "] Teamcenter Cache Clear Operation: End")
					End If

					'Call Function to launch Tc Application from the Path mentioned in EnvVar_Ext.xml
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[" + Cstr(now) + "] Teamcenter Launch Operation: Start")
					bReturn = Fn_InvokeTeamCenter()
					If bReturn = False Then
							Fn_ReUserTcSession = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to launch Application")
							sModule = ""
							'To handle RM testcases scenerio
							bGblPrefDeleteAtLogin = True
							Exit Function
					End If
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[" + Cstr(now) + "] Teamcenter Launch Operation: End")

					'Call Function Login To Tc Application with Supplied Data
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[" + Cstr(now) + "] Teamcenter Login Operation: Start")
					bReturn = Fn_TeamcenterLogin(aLogin(0),aLogin(1), aLogin(2),aLogin(3),aLogin(4))
					If bReturn = False Then
							Fn_ReUserTcSession = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Login to Application")
							sModule = ""
							'To handle RM testcases scenerio
							bGblPrefDeleteAtLogin = True
							Exit Function
					Else
							Fn_ReUserTcSession = True
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Tc Session for User [" + aLogin(0) + "] is Available")
					End If
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[" + Cstr(now) + "] Teamcenter Login Operation: End")
			End If

			'************* If bReuse = True, Check for the Existing Session Details ****************
			If bReuse = True Then
					sSession = ".*" + Lcase(aLogin(0)) + ".*" + aLogin(2) + " / " + aLogin(3) + ".*"
					
					If JavaWindow("DefaultWindow").Exist(20) Then
						JavaWindow("DefaultWindow").Maximize
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Checking for Exisitng Teamcenter Session: Start")
						JavaWindow("DefaultWindow").JavaStaticText("SessionInfo").SetTOProperty "label", sSession
						wait(1)
						If JavaWindow("DefaultWindow").JavaStaticText("SessionInfo").Exist(5) Then
							Fn_ReUserTcSession = True
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Tamcenter Session Exists with User [" + aLogin(0) + "]")
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Checking for Exisitng Teamcenter Session: End")
							'Added by Vallari - For Tracking Tc Code Coverage
							Call Fn_CodeCover_Init()
						Else
							bReuse = False
						End If
					Else
						'Teamcenter Window NOT Found
						bReuse = False
					End if

					If bReuse = False Then
						'Kill Teamcenter Application, if it exists
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[" + Cstr(now) + "] Teamcenter Kill Process Operation: Start")
						bReturn = Fn_KillProcess("")
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[" + Cstr(now) + "] Teamcenter Kill Process Operation: Start")
	
						'Call Function to launch Tc Application from the Path mentioned in EnvVar_Ext.xml
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[" + Cstr(now) + "] Teamcenter Launch Operation: Start")
						bReturn = Fn_InvokeTeamCenter()
						If bReturn = False Then
								Fn_ReUserTcSession = False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to launch Application")
								sModule = ""
								'To handle RM testcases scenerio
								bGblPrefDeleteAtLogin = True
								Exit Function
						End If
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[" + Cstr(now) + "] Teamcenter Launch Operation: End")
	
						'Call Function Login To Tc Application with Supplied Data
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[" + Cstr(now) + "] Teamcenter Login Operation: Start")
						bReturn = Fn_TeamcenterLogin(aLogin(0),aLogin(1), aLogin(2),aLogin(3),aLogin(4))
						If bReturn = False Then
								Fn_ReUserTcSession = False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Login to Application")
								sModule = ""
								'To handle RM testcases scenerio
								bGblPrefDeleteAtLogin = True
								Exit Function
						Else
								Fn_ReUserTcSession = True
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Tc Session for User [" + aLogin(0) + "] is Available")
						End If
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[" + Cstr(now) + "] Teamcenter Login Operation: End")
					End If
			End If

	sModule = ""
	'To handle RM testcases scenerio
	bGblPrefDeleteAtLogin = True
	
	'POC Jotiba T***************************************************************************************
	' To Refresh Preference 
	bRefreshValue=CInt(Fn_GetXMLNodeValue(Fn_GetEnvValue("User", "AutomationDir") & "\TestData\EnvVar_Ext.xml", "PreferenceRefresh"))
	If bRefreshValue = 1 Then
		Call Fn_PreferenceOperations("Refresh","","","","","","","","","","","")
	End If
	'To handle Reset Perspective call
	If bCacheClear = True Then
		GBL_RESET_PERSPECTIVE = False
	End If
End Function

'*********************************************************		Function to set perspective into Teamcenter		***********************************************************************
'Function Name		:				Fn_SetPerspective

'Description			 :		 		 This function open the specified module.

'Parameters			   :	 			1.StrModule : Name of the module.

'Return Value		   : 				TRUE \ FALSE

'Pre-requisite			:		 		should be logged in

'Examples				:				 Fn_SetPerspective("My Teamcenter")

'History					 :		
'	Developer Name												Date						Rev. No.						Changes Done						Reviewer
'	------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Vallari																2304/2010			              1.0										Created
'	------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SetPerspective(StrModule)
	GBL_FAILED_FUNCTION_NAME="Fn_SetPerspective"
	Dim objWindowOpenPerspective, objDefaultWindow, objDefaultWindow1
	Dim bFlag,sMenu
	Dim iRootindex, StrTitle, sDefaultWindowTitle
	
	Fn_SetPerspective = False
	
	Set objDefaultWindow = JavaWindow("DefaultWindow")
    Set objWindowOpenPerspective = objDefaultWindow.JavaWindow("Open Perspective")    	
    Select Case StrModule
		Case "Collaborative Product Development"
			StrModule = "4G Designer"
		Case "Getting Started"
			StrModule = "Getting Started (default)"
		Case "PLMXML Export Import Administration", "PLM XML Export Import Administration"
			StrModule = "PLM XML/TC XML Export Import Administration"
	End Select
	
'	Select menu ["Window - > Open Perspective -> Other...]
   	If Fn_SISW_UI_Object_Operations("Fn_SetPerspective", "Exist", objWindowOpenPerspective,SISW_MICRO_TIMEOUT) = False Then
   		sMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("RAC_Menu"), "WindowOpenPerspectiveOther")
		Call Fn_MenuOperation("Select",sMenu)
		Call Fn_ReadyStatusSync(1)
		If Fn_SISW_UI_Object_Operations("Fn_SetPerspective", "Exist", objWindowOpenPerspective,SISW_MICRO_TIMEOUT) = False Then
			sMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("RAC_Menu"), "WindowOpenPerspectiveOther")
			Call Fn_MenuOperation("Select",sMenu)
			Call Fn_ReadyStatusSync(1)
			If Fn_SISW_UI_Object_Operations("Fn_SetPerspective", "Exist", objWindowOpenPerspective,"") = False Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Operate Menu [ "&sMenu&" ] of Function Fn_SetPerspective")
				Fn_SetPerspective = False
				Exit Function
			End If
		End If
	End If

	'Swapnil : Added for 4GD 
   	'Maximize the open perspective window
	objWindowOpenPerspective.Maximize
	wait SISW_MIN_TIMEOUT
	'Selecting Module
	Fn_SetPerspective = Fn_SISW_UI_JavaTable_Operations("Fn_SetPerspective", "ClickCell", objWindowOpenPerspective , "Table", "", 0, StrModule , 0, "", "", "")
	
	'Swapnil:20-Nov-2013:Clickcell Method not able to select project prespective, hence adding below code.

	If Fn_SISW_UI_Object_Operations("Fn_SetPerspective", "Enabled", objWindowOpenPerspective.JavaButton("OK"),SISW_MICRO_TIMEOUT)  <> True AND StrModule = "Project" Then
		Call Fn_JavaTable_Type("Fn_SetPerspective",objWindowOpenPerspective, "Table",Left(StrModule, 4))
	End If 

	If Fn_SISW_UI_Object_Operations("Fn_SetPerspective", "Exist", objWindowOpenPerspective.JavaButton("OK"),SISW_MICRO_TIMEOUT) Then
		If Fn_SISW_UI_JavaButton_Operations("Fn_SetPerspective", "Click", objWindowOpenPerspective, "OK") = False Then
			Fn_SetPerspective = False
			Exit function
		End If
	End IF	
	
	sDefaultWindowTitle = objDefaultWindow.GetROProperty("title")
	
	If Instr(sDefaultWindowTitle, "Classification") > 0 Then
		If objDefaultWindow.JavaWindow("ErrorJavaWindow").Exist(5) Then
			objDefaultWindow.JavaWindow("ErrorJavaWindow").JavaButton("OK").Click micLeftBtn
		End If
	End If
	
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	'Swapnil :Added for 4GD
	iRootindex = "0"
	StrTitle = "4G Designer"
	
	If Instr(sDefaultWindowTitle,cstr(StrTitle)) > 0 Then
		Set objDefaultWindow1 = JavaWindow("Collaborative Product")
		' if Nav. tree is not empty.
		If cInt(objDefaultWindow1.JavaTree("NavTree").Object.getItemCount()) > 0 Then
			Wait 2
			objDefaultWindow1.JavaTree("NavTree").Deselect cint(iRootindex)
		End If
		Set objDefaultWindow1 = Nothing
	End  If		
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	' - - - - - - - Added code by Koustubh 13-Feb-2013 to reassign X Y values for function Fn_SyncTCObjects after setting perspective.
	Dim objQSearchEdit 
	Set objQSearchEdit = objDefaultWindow.JavaEdit("QuickSearch")
	Fn_SyncTCObjects_Xaxis = cInt(objQSearchEdit.getROProperty("abs_x")) + cInt( objQSearchEdit.getROProperty("width")) + 47
	Fn_SyncTCObjects_Yaxis = cInt(objQSearchEdit.getROProperty("abs_y")) + (cInt( objQSearchEdit.getROProperty("height")) / 2 )+ 5
	Set objQSearchEdit = Nothing
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 

	Set objWindowOpenPerspective = Nothing
	Set objDefaultWindow = Nothing
	Fn_SetPerspective = True
End Function

'*********************************************************		Function to reset perspective Teamcenter		***********************************************************************
'Function Name		:				Fn_ResetPerspective

'Description			 :		 		 This reset the specified module.

'Parameters			   :	 			Nill

'Return Value		   : 				TRUE \ FALSE

'Pre-requisite			:		 		should be logged in

'Examples				:				 Fn_ResetPerspective()

'History					 :		
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'												------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Jhuma Debnath										10/03/2010			              1.0										Created
'												------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_ResetPerspective()
	GBL_FAILED_FUNCTION_NAME="Fn_ResetPerspective"
	Dim bFlag, objDefaultWindow, objDefaultWindow1,i
	Dim iRootindex, StrTitle, sDefaultWindowTitle, sMenu
	
	Set objDefaultWindow = JavaWindow("DefaultWindow")
	'POC Vivek Ahirrao***************************************************************************************

	
		If GBL_RESET_PERSPECTIVE = True Then
			
			
			'Select menu ["Window - > Reset Perspective...]
			sMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("RAC_Menu"), "WindowResetPerspective")
			bFlag = Fn_MenuOperation("WinMenuSelect",sMenu)
			If bFlag = False Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Operate Menu [ "&sMenu&" ] of Function Fn_ResetPerspective")
				Fn_ResetPerspective = False
				Exit Function
			End If
'		
'			'Set objWindowResetPerspective = JavaWindow("DefaultWindow").JavaWindow("Reset Perspective")
'			' Click on [OK] button
'			'Call Fn_Button_Click("Fn_ResetPerspective", objWindowResetPerspective, "OK")
			objDefaultWindow.JavaWindow("Reset Perspective").JavaButton("OK").WaitProperty "enabled",True,10
			objDefaultWindow.JavaWindow("Reset Perspective").JavaButton("OK").Click micLeftBtn
			If Err.Number < 0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Click Button [OK] on Reset Perspective Dialog of Function Fn_ResetPerspective")
				Fn_ResetPerspective = False
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Successfully Clicked Button [OK] on Reset Perspective Dialog of Function Fn_ResetPerspective")
				Fn_ResetPerspective = True
				Call Fn_SyncTCObjects()		
			End If
		End If

'		
		sDefaultWindowTitle = objDefaultWindow.GetROProperty("title")
'		
		StrTitle = "Schedule Manager"
		If Instr(sDefaultWindowTitle,cstr(StrTitle)) = 0 Then
			If Instr(sDefaultWindowTitle,"Product Configurator") = 0 Then
				Call Fn_RefreshWindow()
			End If
		End If
		
'		'Swapnil : By default in CPD perspective root object gets selected that is causing problem , so added code to deselect root object.
		iRootindex = "0"
		StrTitle = "4G Designer"
		If Instr(sDefaultWindowTitle,cstr(StrTitle)) > 0 Then
			Set objDefaultWindow1 = JavaWindow("Collaborative Product")
			' if Nav. tree is not empty.
			If cInt(objDefaultWindow1.JavaTree("NavTree").Object.getItemCount()) > 0 Then
				Wait 1
				objDefaultWindow1.JavaTree("NavTree").Deselect cint(iRootindex)
			End If
			Set objDefaultWindow1 = Nothing
		End If
'		
		StrTitle = "Product Configurator"
		If Instr(sDefaultWindowTitle,cstr(StrTitle)) > 0 Then
			Set objDefaultWindow1 = JavaWindow("ProductConfigurator")
			' if Nav. tree is not empty.
			If objDefaultWindow1.JavaTree("NavTreeTable").Exist(1) Then
				If cInt(objDefaultWindow1.JavaTree("NavTreeTable").Object.getItemCount()) > 0 Then
					Wait 1
					objDefaultWindow1.JavaTree("NavTreeTable").Deselect cint(iRootindex)
				End If
			End If
			Set objDefaultWindow1 = Nothing
		End If

	Call Fn_RefreshWindow()
	Set objDefaultWindow = Nothing
	
	GBL_RESET_PERSPECTIVE = True
	Fn_ResetPerspective = True
End Function
'******************************************************************Function to perform menu operations************************************************************************************************************

'Function Name					:Fn_MenuOperation

'Description						:1. This function is used to select the menu option. The menu which is to be selected is decided from the argument value passed by the script.
'												2. This function is used to check existance the menu option.
'												3. This function is used to check the state(Enabled/Disabled) of the menu option.

'Parameters						:1. StrAction: Action to be performed (Eg : Select/Exist;/State )
'											   2. StrMenuLabel: Exact name of the menu option (delimiter as ':')
'											  

'Return Value		   : 	    TRUE \ FALSE

'Pre-requisite					:Base state should be present 

'Examples						:Fn_MenuOperation("Select","File:New:Item...)
'											Call Fn_MenuOperation("WinMenuSelect", "View:Show Unconfigured Variants")
'											Call Fn_MenuOperation("WinMenuState", "View:Show Unconfigured Changes")
'											Call Fn_MenuOperation("WinMenuUnCheck", "View:Show Unconfigured Changes")
'											Call Fn_MenuOperation("WinMenuCheck", "View:Show Unconfigured Variants")
'											Call Fn_MenuOperation("SelectMenuObjects", "Tools:Reports:Report Designer Reports")											

'History:
'										Developer Name			Date			Rev. No.			Changes Done										Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Jhuma Debnath			09/03/09			1.0  									
'										Sachin Joshi				25/09/10																							   Rupali
'										Prasanna B.                 22/10/10                               Added SelectMenuObjects Case
'										Koustubh W.                 04/11/11                     Added code for build 910928
'										Nilesh G.                 14/02/13                     Added code to handle special characters {  ( and ) } in menu . eg : File:SaveAs...:Item(Revision)...
'										Sandeep N.                 13/03/13                     Updated code to handle special characters {  ( and ) } in menu . eg : File:SaveAs...:Item(Revision)...
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Public Function Fn_MenuOperation(StrAction, StrMenuLabel)
'	On Error Resume Next
'	
'    Dim ArrMenu, NumObjects, iCounter,	dMenu,StrMenu,iMenuCount,objMenu,StrMenuLabel1
'	Dim bReturn
'	
'	GBL_LAST_INVOKE_MENU_NAME = StrMenuLabel
'	gLastMenuCall = StrMenuLabel
'	
'	If Instr(StrMenuLabel,"\(")>0 and Instr(StrMenuLabel,"\)")>0 Then
'		'Do nothing
'	Else
'		If Instr(StrMenuLabel,"(")>0 Then
'			StrMenuLabel=Replace(StrMenuLabel,"(","\(")
'		End If
'		If Instr(StrMenuLabel,")")>0 Then
'			StrMenuLabel=Replace(StrMenuLabel,")","\)")
'		End If
'	End If
'	
'	JavaWindow("DefaultWindow").Object.setActive
'	
'	'Vallari [29Jun11] - In few PSE Category tests, NavTree becomes irresponsive if MMP perspective is closed
'	If trim(StrMenuLabel) = "File:Close" Then
'			'If instr(JavaWindow("DefaultWindow").GetROProperty("title"), "Manufacturing Process Planner") > 0 OR instr(JavaWindow("DefaultWindow").GetROProperty("title"), "My Teamcenter")  > 0 OR instr(JavaWindow("DefaultWindow").GetROProperty("title"), "Systems Engineering")  > 0 OR instr(JavaWindow("DefaultWindow").GetROProperty("title"), "Organization")  > 0 OR instr(JavaWindow("DefaultWindow").GetROProperty("title"), "Structure Manager")  > 0 Then
'	
'			 'Amit  - [ 09-Nov-2011 ] - Removing PSE and MPP perspectives from above condition since now [ File -> Close ] does NOT cause Irresponsiveness.
'			If Instr(JavaWindow("DefaultWindow").GetROProperty("title"), "My Teamcenter")  > 0 OR Instr(JavaWindow("DefaultWindow").GetROProperty("title"), "Systems Engineering")  > 0 OR Instr(JavaWindow("DefaultWindow").GetROProperty("title"), "Organization")  > 0 Then
'					Call Fn_ToolBarOperation("Click", "Back", "") 
'					Fn_MenuOperation = True
'					Exit Function
'			End If
'	End If
'
''- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
''	Added code by Koustubh [4-Nov-2011] 
''	In Manufacturing Process Planner perspective Pack, unpack menus are WinMenu object on build 910928.( MPP is unstable on build 9101026  )
''
'	If trim(StrMenuLabel) = "View:Pack Unpack:Unpack All" OR trim(StrMenuLabel) = "View:Pack Unpack:Pack All" OR trim(StrMenuLabel) = "View:Pack Unpack:Unpack	Ctrl+Shift+N" OR trim(StrMenuLabel) = "View:Pack Unpack:Pack	Ctrl+Shift+M" Then
'		If instr(JavaWindow("DefaultWindow").GetROProperty("title"), "Manufacturing Process Planner") > 0 then
'			If StrAction = "Select" Then StrAction = "WinMenuSelect"
'		End If
'	End If
''- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'	
''	Added code by Chandrakant Tyagi [8-Jan-2015]
''	added code to handle Chnged Menus  in Systems Engineering perspective  for TC 11.2
'	Dim sPath, sMenu
'	sMenu = StrMenuLabel
'	If instr(JavaWindow("DefaultWindow").GetROProperty("title"), "Systems Engineering") > 0  Then
'    	StrMenuLabel = Replace(StrMenuLabel, ":", "")
'		StrMenuLabel = Replace(StrMenuLabel, " ", "")
'		If instr(StrMenuLabel,"End Trace Link...")=0  Then
'			StrMenuLabel = Replace(StrMenuLabel, ".", "")
'		End If
'		Select Case StrMenuLabel
'			Case "FileNewInterface", "FileNewConnection", "FileOpenText", "FileNewCreateDiagram", "FileNewCustomNote", "FileNewBlocksOther", "FileNewRequirementsSpec", "FileNewBudgetDefinition", "ToolsSetClosureRule","ViewShowCustomNotes", "ToolsSignalManagerAssociateSignalToTarget"
'				sPath= Fn_LogUtil_GetXMLPath("RM_Menu")
'				StrMenuLabel =Fn_GetXMLNodeValue(sPath, StrMenuLabel)
'			Case "ToolsTraceLinkTraceabilityReportComplying", "ToolsTraceLinkTraceabilityReportDefining"
'				bReturn =  Fn_SetView("Systems Engineering:Traceability")
'				Fn_MenuOperation = bReturn
'				Exit Function	
'			Case "ViewExpandOptionsExpand", "ViewExpandOptionsExpandBelow", "ViewExpandOptionsExpandBelow...", "ViewExpandOptionsExpandToType", "ViewRefresh"
'					StrMenuLabel = sMenu
'					If StrAction = "Select" Then StrAction = "WinMenuSelect"
'			Case "ViewAccess"
'				sPath= Fn_LogUtil_GetXMLPath("RM_Menu")
'				StrMenuLabel =Fn_GetXMLNodeValue(sPath, StrMenuLabel)		
'			Case Else
'				StrMenuLabel = sMenu
'		End Select
'	End If
'
''- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'
'	If StrMenuLabel = "Tools:Revision Rule:Save Configuration..." then 
'		StrMenuLabel = "Tools:Revision Rule:Save as New Configuration Context"
'	End if
'
'	If StrMenuLabel = "Edit:Paste Special" then 
'		StrMenuLabel = "Edit:Paste Special..."
'	End if
'
'	If StrMenuLabel = "File:CC:Save as New Structure Context..." then
'			StrMenuLabel = "File:Collaboration Context:Save as New Structure Context..."
'  
'	 ElseIf StrMenuLabel = "File:CC:Save Configuration" then
'			 StrMenuLabel = "File:Collaboration Context:Save Configuration"
'	  
'	 ElseIf StrMenuLabel = "File:CC:Save as New Configuration Context" then
'			  StrMenuLabel = "File:Collaboration Context:Save as New Configuration Context"
'	  
'	 ElseIf StrMenuLabel = "File:CC:Apply Configuration Context..." then
'			 StrMenuLabel = "File:Collaboration Context:Apply Configuration Context..."
'	
'	ElseIf StrMenuLabel = "File:CC:Save as New Collaboration Context..." then
'			 StrMenuLabel = "File:Collaboration Context:Save as New Collaboration Context..."
'	End If
'
'	If trim(StrMenuLabel) = "File:Save" Then
'		If instr(JavaWindow("DefaultWindow").GetROProperty("title"), "Platform Designer") > 0 then
'				   bolState = JavaWindow("DefaultWindow").JavaMenu("label:=File","index:=0").JavaMenu("label:=Save","index:=0").GetROProperty("enabled")
'				   If Cbool(bolState) = False Then
'								Fn_MenuOperation = True
'								Exit Function
'					End if								
'		End If
'	End If
'
'	Select Case StrAction
'
'		'.---------------------------------------This case is used to select the menu option.----------------------------------------------
'		Case "Select"
'				If Trim(StrMenuLabel) =  "Tools:Variants:Update Variant Item..."Then
'						set dMenu=description.create()
'						dMenu("menuobjtype").value=2
'		
'						JavaWindow("DefaultWindow").winmenu(dMenu).select "Tools;Variants;Update Variant Item..."
'						Fn_MenuOperation = True 
'
'				Else
'						Fn_MenuOperation = Fn_UI_JavaMenu_Select("Fn_MenuOperation",JavaWindow("DefaultWindow"),StrMenuLabel)
'						If instr(1,StrMenuLabel,"File:Close") Then
'							If JavaWindow("DefaultWindow").JavaWindow("LastPerspective").Exist(3) Then
'								JavaWindow("DefaultWindow").JavaWindow("LastPerspective").JavaButton("OK").Click micLeftBtn
'							End If
'						End If
'				End If
'		'.---------------------------------------This case is used to select the menu option.----------------------------------------------
'		Case "KeyPress"
'			'Split Menu String
'			ArrMenu=Split(StrMenuLabel,":") 
'			NumObjects = ubound(ArrMenu)
'
'			'This is a Special case to operate menu by KeyPress method - Few Tc menus are not getting selected by traditional way
'			Select Case NumObjects
'				Case "1"
'						JavaWindow("DefaultWindow").PressKey Left(ArrMenu(0), 1), micAlt
'						JavaWindow("DefaultWindow").PressKey Left(ArrMenu(1), 1)
'				Case "2"
'						JavaWindow("DefaultWindow").PressKey Left(ArrMenu(0), 1), micAlt
'						JavaWindow("DefaultWindow").PressKey Left(ArrMenu(1), 1)
'						JavaWindow("DefaultWindow").PressKey Left(ArrMenu(2), 1)
'			End Select
'			Fn_MenuOperation = TRUE							
'			'Call Fn_WriteLogFile("Fn_MenuOperation", 3, Err.Number,"PASS: Menu selected succefully")
'			If instr(1,StrMenuLabel,"File:Close") Then
'				If JavaWindow("DefaultWindow").JavaWindow("LastPerspective").Exist(3) Then
'					JavaWindow("DefaultWindow").JavaWindow("LastPerspective").JavaButton("OK").Click micLeftBtn
'				End If
'			End If
'		'.---------------------------------------This case is used to check existance the menu option.---------------------------------------
'		Case "Exist" 
'			Dim bolExist
'			StrMenuLabel1=Replace(StrMenuLabel,";",":")
'			ArrMenu=Split(StrMenuLabel1,":")
'			NumObjects = ubound(ArrMenu) 
'
'			Select Case NumObjects
'					Case "0"								
'						bolExist =JavaWindow("DefaultWindow").JavaMenu("label:="&ArrMenu(0)&"","index:=0", "path:=MenuItem;Shell;").Exist(10)
'                        Fn_MenuOperation =bolExist
'
'					Case "1"								
'						bolExist =JavaWindow("DefaultWindow").JavaMenu("label:="&ArrMenu(0)&"","index:=0").JavaMenu("label:="&ArrMenu(1)&"","index:=0").Exist(10)
'                        Fn_MenuOperation =bolExist
'                        						
'					Case "2"
'						bolExist = JavaWindow("DefaultWindow").JavaMenu("label:="&ArrMenu(0)&"","index:=0").JavaMenu("label:="&ArrMenu(1)&"","index:=0").JavaMenu("label:="&ArrMenu(2)&"","index:=0").Exist(10)
'                        Fn_MenuOperation = bolExist
'            
'					 Case "3"
'							bolExist = JavaWindow("DefaultWindow").JavaMenu("label:="&ArrMenu(0)&"","index:=0").JavaMenu("label:="&ArrMenu(1)&"","index:=0").JavaMenu("label:="&ArrMenu(2)&"","index:=0").JavaMenu("label:="&ArrMenu(3)&"","index:=0").Exist(10)
'                            Fn_MenuOperation =bolExist
'							
'					Case "4"
'							bolExist = JavaWindow("DefaultWindow").JavaMenu("label:="&ArrMenu(0)&"","index:=0").JavaMenu("label:="&ArrMenu(1)&"","index:=0").JavaMenu("label:="&ArrMenu(2)&"","index:=0").JavaMenu("label:="&ArrMenu(3)&"","index:=0").JavaMenu("label:="&ArrMenu(4)&"","index:=0").Exist(10)
'                        	Fn_MenuOperation =bolExist
'							
'					Case Else
'							bolExist =JavaWindow("DefaultWindow").JavaMenu("label:="&ArrMenu(0)&"","index:=0", "path:=MenuItem;Shell;").Exist(10)
'							Fn_MenuOperation =bolExist
'
'			End Select
'
'		'.---------------------------------------This case is used to check the state(Enabled/Disabled) of the menu option..---------------------------------------
'		Case "State" 
'			Dim bolState
'			ArrMenu=Split(StrMenuLabel,":")
'			NumObjects = ubound(ArrMenu) 
'
'			Select Case NumObjects
'					Case "1"	
'						   bolState = JavaWindow("DefaultWindow").JavaMenu("label:="&ArrMenu(0)&"","index:=0").JavaMenu("label:="&ArrMenu(1)&"","index:=0").GetROProperty("enabled")						
'							If  Cbool(bolState) Then
'								Fn_MenuOperation = True
'					     	Else 
'								Fn_MenuOperation = False
'						   End If	
'
'					Case "2"	
'						   bolState = JavaWindow("DefaultWindow").JavaMenu("label:="&ArrMenu(0)&"","index:=0").JavaMenu("label:="&ArrMenu(1)&"","index:=0").JavaMenu("label:="&ArrMenu(2)&"","index:=0").GetROProperty("enabled")						
'							If  Cbool(bolState) Then
'								Fn_MenuOperation = True
'					     	Else 
'								Fn_MenuOperation = False
'						   End If	
'
'					Case "3"	
'						   bolState = JavaWindow("DefaultWindow").JavaMenu("label:="&ArrMenu(0)&"","index:=0").JavaMenu("label:="&ArrMenu(1)&"","index:=0").JavaMenu("label:="&ArrMenu(2)&"","index:=0").JavaMenu("label:="&ArrMenu(3)&"","index:=0").GetROProperty("enabled")						
'							If  Cbool(bolState) Then
'								Fn_MenuOperation = True
'					     	Else 
'								Fn_MenuOperation = False
'						   End If	
'
'					Case "4"	
'						   bolState = JavaWindow("DefaultWindow").JavaMenu("label:="&ArrMenu(0)&"","index:=0").JavaMenu("label:="&ArrMenu(1)&"","index:=0").JavaMenu("label:="&ArrMenu(2)&"","index:=0").JavaMenu("label:="&ArrMenu(3)&"","index:=0").JavaMenu("label:="&ArrMenu(4)&"","index:=0").GetROProperty("enabled")						
'							If  Cbool(bolState) Then
'								Fn_MenuOperation = True
'					     	Else 
'								Fn_MenuOperation = False
'						   End If	
'					Case Else
'							Fn_MenuOperation = "PASS:Invalid option"
'			End Select
'		'.---------------------------------------This case is used to Select Window Menu Item..---------------------------------------
'		' all Below Cases Works only for Structure manager Menu's (View,Tools-->Variants)
'		Case "WinMenuSelect"
'				StrMenu=Replace(StrMenuLabel,":",";")
'				set dMenu=description.create()
'				dMenu("menuobjtype").value=2
'				wait(1)
'				JavaWindow("DefaultWindow").winmenu(dMenu).Select(StrMenu)
'				wait 1
'				Fn_MenuOperation = True
'				'Returns true when the Menu Item is enabled
'			'.---------------------------------------This case is used to Check Window Menu Item State (Enabled)..---------------------------------------
'		Case "WinMenuState"
'				StrMenu=Replace(StrMenuLabel,":",";")
'				set dMenu=description.create()
'				dMenu("menuobjtype").value=2
'                If (JavaWindow("DefaultWindow").winmenu(dMenu).GetItemProperty (StrMenu,"enabled")) Then
'					Fn_MenuOperation=True
'				Else
'					Fn_MenuOperation=False
'				End If
'       '.---------------------------------------This case is used to Check Window Menu Item State (Checked)..---------------------------------------
'					'Returns true when the Item state is Checked
'		Case "WinMenuCheck"
'				StrMenu=Replace(StrMenuLabel,":",";")
'                set dMenu=description.create()
'				dMenu("menuobjtype").value=2
'				wait 1
'				bReturn = JavaWindow("DefaultWindow").winmenu(dMenu).CheckItemProperty(StrMenu, "checked","1",1)
'				If bReturn Then
'					Fn_MenuOperation=True
'				Else
'					Fn_MenuOperation=False
'				End If
'			'.---------------------------------------This case is used to Check Window Menu Item State (Checked)..---------------------------------------
'			'Returns true when the Item state is UnChecked
'		Case "WinMenuUnCheck"
'				StrMenu=Replace(StrMenuLabel,":",";")
'                set dMenu=description.create()
'				dMenu("menuobjtype").value=2
'				If (JavaWindow("DefaultWindow").winmenu(dMenu).CheckItemProperty(StrMenu, "checked","1",1)) Then
'					Fn_MenuOperation=False
'				Else
'					Fn_MenuOperation=True
'				End If
'		  '.---------------------------------------This case is used  for those menu operations that do not work with Select case---------------------------------------
'		 Case "SelectMenuObjects"
'				ArrMenu = split(StrMenuLabel,":",-1,1)
'				 iMenuCount = Ubound(ArrMenu) 
'				set dMenu=description.create()
'				dMenu("menuobjtype").value=2
'				dMenu("label").value=ArrMenu(iMenuCount)		
'				set objMenu = JavaWindow("DefaultWindow").ChildObjects(dMenu)
'				objMenu(0).Select
'				If Err.Number < 0 Then
'					Fn_MenuOperation = FALSE
'				Else
'					Fn_MenuOperation = True
'				End If
'
'		Case Else
'						Fn_MenuOperation = FALSE
'						Reporter.ReportEvent micPass,"Menu operation","Invalid option"
'		End Select
'
'End Function

Public Function Fn_MenuOperation(StrAction, StrMenuLabel)
	GBL_FAILED_FUNCTION_NAME="Fn_MenuOperation"
	On Error Resume Next

	'Declaring variables
    Dim aMenu
	Dim bReturn
	Dim objMenu,objChildObjects
	Dim sMenu,sMenuLabel1,sPath
	Dim iNumObjects,iCounter,iMenuCount
	
	gLastMenuCall = StrMenuLabel
	
	If Instr(StrMenuLabel,"\(")>0 and Instr(StrMenuLabel,"\)")>0 Then
		'Do nothing
	Else
		If Instr(StrMenuLabel,"(")>0 Then
			StrMenuLabel=Replace(StrMenuLabel,"(","\(")
		End If
		If Instr(StrMenuLabel,")")>0 Then
			StrMenuLabel=Replace(StrMenuLabel,")","\)")
		End If
	End If
	'Activating teamcenter window
	JavaWindow("DefaultWindow").Object.SetActive
	
	If trim(StrMenuLabel) = "File:Close" Then
		'Amit  - [ 09-Nov-2011 ] - Removing PSE and MPP perspectives from above condition since now [ File -> Close ] does NOT cause Irresponsiveness.
		If Instr(JavaWindow("DefaultWindow").GetROProperty("title"), "My Teamcenter")  > 0 OR Instr(JavaWindow("DefaultWindow").GetROProperty("title"), "Systems Engineering")  > 0 OR Instr(JavaWindow("DefaultWindow").GetROProperty("title"), "Organization")  > 0 Then
			Call Fn_ToolBarOperation("Click", "Back", "") 
			Fn_MenuOperation = True
			Exit Function
		End If
	End If

	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	'	Added code by Koustubh [4-Nov-2011] 
	'	In Manufacturing Process Planner perspective Pack, unpack menus are WinMenu object on build 910928.( MPP is unstable on build 9101026  )
	'
	If trim(StrMenuLabel) = "View:Pack Unpack:Unpack All" OR trim(StrMenuLabel) = "View:Pack Unpack:Pack All" OR trim(StrMenuLabel) = "View:Pack Unpack:Unpack	Ctrl+Shift+N" OR trim(StrMenuLabel) = "View:Pack Unpack:Pack	Ctrl+Shift+M" Then
		If instr(JavaWindow("DefaultWindow").GetROProperty("title"), "Manufacturing Process Planner") > 0 then
			If StrAction = "Select" Then StrAction = "WinMenuSelect"
		End If
	End If
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	'	Added code by Jotiba Takkekar [6-August-2019] 
	If Environment.Value("ProductName") = sUFTProductName Then
		sMenuName=split(StrMenuLabel,":")
		If Ubound(sMenuName) >= 3 Then
			If StrAction = "Select" Then StrAction = "WinMenuSelect"
		End If
	End If
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 	
	'	Added code by Chandrakant Tyagi [8-Jan-2015]
	'	added code to handle Chnged Menus  in Systems Engineering perspective  for TC 11.2
	sMenu = StrMenuLabel
	If instr(JavaWindow("DefaultWindow").GetROProperty("title"), "Systems Engineering") > 0  Then
    	StrMenuLabel = Replace(StrMenuLabel, ":", "")
		StrMenuLabel = Replace(StrMenuLabel, " ", "")
		If instr(StrMenuLabel,"End Trace Link...")=0  Then
			StrMenuLabel = Replace(StrMenuLabel, ".", "")
		End If
		
		Select Case StrMenuLabel
			Case "FileNewInterface", "FileNewConnection", "FileOpenText", "FileNewCreateDiagram", "FileNewCustomNote", "FileNewBlocksOther", "FileNewRequirementsSpec", "FileNewBudgetDefinition", "ToolsSetClosureRule","ViewShowCustomNotes", "ToolsSignalManagerAssociateSignalToTarget"
				sPath= Fn_LogUtil_GetXMLPath("RM_Menu")
				StrMenuLabel =Fn_GetXMLNodeValue(sPath, StrMenuLabel)
			Case "ToolsTraceLinkTraceabilityReportComplying", "ToolsTraceLinkTraceabilityReportDefining"
				bReturn =  Fn_SetView("Systems Engineering:Traceability")
				Fn_MenuOperation = bReturn
				Exit Function	
			Case "ViewExpandOptionsExpand", "ViewExpandOptionsExpandBelow", "ViewExpandOptionsExpandBelow...", "ViewExpandOptionsExpandToType", "ViewRefresh"
				StrMenuLabel = sMenu
				If StrAction = "Select" Then StrAction = "WinMenuSelect"
			Case "ViewAccess"
				sPath= Fn_LogUtil_GetXMLPath("RM_Menu")
				StrMenuLabel =Fn_GetXMLNodeValue(sPath, StrMenuLabel)		
			Case Else
				StrMenuLabel = sMenu
		End Select
	End If
	
	Select Case StrMenuLabel
		Case "Tools:Revision Rule:Save Configuration..."
			 StrMenuLabel = "Tools:Revision Rule:Save as New Configuration Context"
		Case "Edit:Paste Special"
			 StrMenuLabel = "Edit:Paste Special..."	
		Case "File:CC:Save as New Structure Context..."
			 StrMenuLabel = "File:Collaboration Context:Save as New Structure Context..."
		Case "File:CC:Save Configuration"
			 StrMenuLabel = "File:Collaboration Context:Save Configuration"
		Case "File:CC:Save as New Configuration Context" 
			 StrMenuLabel = "File:Collaboration Context:Save as New Configuration Context"
		Case "File:CC:Apply Configuration Context..." 
			StrMenuLabel = "File:Collaboration Context:Apply Configuration Context..."
		Case "File:CC:Save as New Collaboration Context..." 
			 StrMenuLabel = "File:Collaboration Context:Save as New Collaboration Context..."
	End Select

	If trim(StrMenuLabel) = "File:Save" Then
		If instr(JavaWindow("DefaultWindow").GetROProperty("title"), "Platform Designer") > 0 then
			bReturn = JavaWindow("DefaultWindow").JavaMenu("label:=File","index:=0").JavaMenu("label:=Save","index:=0").GetROProperty("enabled")
			If Cbool(bReturn) = False Then
				Fn_MenuOperation = True
				Exit Function
			End if
		End If
	End If

	Select Case StrAction
		'---------------------------------------This case is used to select the menu option.----------------------------------------------
		Case "Select"
			If Trim(StrMenuLabel) =  "Tools:Variants:Update Variant Item..." Then
				Set objMenu=description.create()
				objMenu("menuobjtype").value=2
				JavaWindow("DefaultWindow").winmenu(objMenu).select "Tools;Variants;Update Variant Item..."
				Fn_MenuOperation = True 
				Set objMenu=Nothing
			Else
				Fn_MenuOperation = Fn_UI_JavaMenu_Select("Fn_MenuOperation",JavaWindow("DefaultWindow"),StrMenuLabel)
				If instr(1,StrMenuLabel,"File:Close") Then
					If JavaWindow("DefaultWindow").JavaWindow("LastPerspective").Exist(2) Then
						JavaWindow("DefaultWindow").JavaWindow("LastPerspective").JavaButton("OK").Click micLeftBtn
					End If
				End If
			End If
		'---------------------------------------This case is used to select the menu option.----------------------------------------------
		Case "KeyPress"
			'Split Menu String
			aMenu=Split(StrMenuLabel,":") 
			iNumObjects = ubound(aMenu)

			'This is a Special case to operate menu by KeyPress method - Few Tc menus are not getting selected by traditional way
			Select Case iNumObjects
				Case "1"
					JavaWindow("DefaultWindow").PressKey Left(aMenu(0), 1), micAlt
					JavaWindow("DefaultWindow").PressKey Left(aMenu(1), 1)
				Case "2"
					JavaWindow("DefaultWindow").PressKey Left(aMenu(0), 1), micAlt
					JavaWindow("DefaultWindow").PressKey Left(aMenu(1), 1)
					JavaWindow("DefaultWindow").PressKey Left(aMenu(2), 1)
			End Select
			Fn_MenuOperation = TRUE							
			'Call Fn_WriteLogFile("Fn_MenuOperation", 3, Err.Number,"PASS: Menu selected succefully")
			If instr(1,StrMenuLabel,"File:Close") Then
				If JavaWindow("DefaultWindow").JavaWindow("LastPerspective").Exist(2) Then
					JavaWindow("DefaultWindow").JavaWindow("LastPerspective").JavaButton("OK").Click micLeftBtn
				End If
			End If
		'---------------------------------------This case is used to check existance the menu option.---------------------------------------
		Case "Exist"   'Modified by Chaitali R. 
			sMenuLabel1=Replace(StrMenuLabel,";",":")
			aMenu=Split(sMenuLabel1,":")
			iNumObjects = ubound(aMenu) 

			Select Case iNumObjects
				Case "0"								
					bReturn =JavaWindow("DefaultWindow").JavaMenu("label:="&aMenu(0)&"","index:=0", "path:=MenuItem;Shell;").Exist(5)
					Fn_MenuOperation =bReturn
				Case "1"	
					bReturn =JavaWindow("DefaultWindow").JavaMenu("label:="&aMenu(0)&"","index:=0", "path:=MenuItem;Shell;").Select				
					Call Fn_ReadyStatusSync(1)
					bReturn =JavaWindow("DefaultWindow").JavaMenu("label:="&aMenu(0)&"","index:=0").JavaMenu("label:="&aMenu(1)&"","index:=0").Exist(5)
					Fn_MenuOperation =bReturn                        						
				Case "2"
					bReturn =JavaWindow("DefaultWindow").JavaMenu("label:="&aMenu(0)&"","index:=0").JavaMenu("label:="&aMenu(1)&"","index:=0").Select
					Call Fn_ReadyStatusSync(1)
					bReturn = JavaWindow("DefaultWindow").JavaMenu("label:="&aMenu(0)&"","index:=0").JavaMenu("label:="&aMenu(1)&"","index:=0").JavaMenu("label:="&aMenu(2)&"","index:=0").Exist(5)
					Fn_MenuOperation = bReturn            
				 Case "3"
				 	bReturn = JavaWindow("DefaultWindow").JavaMenu("label:="&aMenu(0)&"","index:=0").JavaMenu("label:="&aMenu(1)&"","index:=0").JavaMenu("label:="&aMenu(2)&"","index:=0").Select
					Call Fn_ReadyStatusSync(1)
					bReturn = JavaWindow("DefaultWindow").JavaMenu("label:="&aMenu(0)&"","index:=0").JavaMenu("label:="&aMenu(1)&"","index:=0").JavaMenu("label:="&aMenu(2)&"","index:=0").JavaMenu("label:="&aMenu(3)&"","index:=0").Exist(5)
					Fn_MenuOperation =bReturn							
				Case "4"
					bReturn = JavaWindow("DefaultWindow").JavaMenu("label:="&aMenu(0)&"","index:=0").JavaMenu("label:="&aMenu(1)&"","index:=0").JavaMenu("label:="&aMenu(2)&"","index:=0").JavaMenu("label:="&aMenu(3)&"","index:=0").Select
					Call Fn_ReadyStatusSync(1)
					bReturn = JavaWindow("DefaultWindow").JavaMenu("label:="&aMenu(0)&"","index:=0").JavaMenu("label:="&aMenu(1)&"","index:=0").JavaMenu("label:="&aMenu(2)&"","index:=0").JavaMenu("label:="&aMenu(3)&"","index:=0").JavaMenu("label:="&aMenu(4)&"","index:=0").Exist(5)
					Fn_MenuOperation =bReturn						
				Case Else
					bReturn =JavaWindow("DefaultWindow").JavaMenu("label:="&aMenu(0)&"","index:=0", "path:=MenuItem;Shell;").Exist(10)
					Fn_MenuOperation =bReturn
			End Select
			Call Fn_ReadyStatusSync(1)
			bReturn =JavaWindow("DefaultWindow").JavaMenu("label:="&aMenu(0)&"","index:=0", "path:=MenuItem;Shell;").Select
		'---------------------------------------This case is used to check the state(Enabled/Disabled) of the menu option..---------------------------------------
		Case "State"   'Modified by Chaitali R.
			aMenu=Split(StrMenuLabel,":")
			iNumObjects = ubound(aMenu) 

			Select Case iNumObjects
				Case "1"	
					bReturn = JavaWindow("DefaultWindow").JavaMenu("label:="&aMenu(0)&"","index:=0").Select						
					Call Fn_ReadyStatusSync(1)
					bReturn = JavaWindow("DefaultWindow").JavaMenu("label:="&aMenu(0)&"","index:=0").JavaMenu("label:="&aMenu(1)&"","index:=0").GetROProperty("enabled")						
					If  Cbool(bReturn) Then
						Fn_MenuOperation = True
					Else 
						Fn_MenuOperation = False
					End If	
				Case "2"	
					bReturn = JavaWindow("DefaultWindow").JavaMenu("label:="&aMenu(0)&"","index:=0").JavaMenu("label:="&aMenu(1)&"","index:=0").Select
					Call Fn_ReadyStatusSync(1)
					bReturn = JavaWindow("DefaultWindow").JavaMenu("label:="&aMenu(0)&"","index:=0").JavaMenu("label:="&aMenu(1)&"","index:=0").JavaMenu("label:="&aMenu(2)&"","index:=0").GetROProperty("enabled")						
					If  Cbool(bReturn) Then
						Fn_MenuOperation = True
					Else 
						Fn_MenuOperation = False
					End If	
				Case "3"	
					bReturn = JavaWindow("DefaultWindow").JavaMenu("label:="&aMenu(0)&"","index:=0").JavaMenu("label:="&aMenu(1)&"","index:=0").JavaMenu("label:="&aMenu(2)&"","index:=0").Select					
					Call Fn_ReadyStatusSync(1)
					bReturn = JavaWindow("DefaultWindow").JavaMenu("label:="&aMenu(0)&"","index:=0").JavaMenu("label:="&aMenu(1)&"","index:=0").JavaMenu("label:="&aMenu(2)&"","index:=0").JavaMenu("label:="&aMenu(3)&"","index:=0").GetROProperty("enabled")						
					If  Cbool(bReturn) Then
						Fn_MenuOperation = True
					Else 
						Fn_MenuOperation = False
					End If	
				Case "4"	
					bReturn = JavaWindow("DefaultWindow").JavaMenu("label:="&aMenu(0)&"","index:=0").JavaMenu("label:="&aMenu(1)&"","index:=0").JavaMenu("label:="&aMenu(2)&"","index:=0").JavaMenu("label:="&aMenu(3)&"","index:=0").Select
					Call Fn_ReadyStatusSync(1)
					bReturn = JavaWindow("DefaultWindow").JavaMenu("label:="&aMenu(0)&"","index:=0").JavaMenu("label:="&aMenu(1)&"","index:=0").JavaMenu("label:="&aMenu(2)&"","index:=0").JavaMenu("label:="&aMenu(3)&"","index:=0").JavaMenu("label:="&aMenu(4)&"","index:=0").GetROProperty("enabled")						
					If  Cbool(bReturn) Then
						Fn_MenuOperation = True
					Else 
						Fn_MenuOperation = False
					End If	
				Case Else
					Fn_MenuOperation = "PASS:Invalid option"
			End Select
			Call Fn_ReadyStatusSync(1)
			bReturn =JavaWindow("DefaultWindow").JavaMenu("label:="&aMenu(0)&"","index:=0", "path:=MenuItem;Shell;").Select
		'---------------------------------------This case is used to Select Window Menu Item..---------------------------------------
		' all Below Cases Works only for Structure manager Menu's (View,Tools-->Variants)
		Case "WinMenuSelect"
			sMenu=Replace(StrMenuLabel,":",";")
			Set objMenu=description.create()
			objMenu("menuobjtype").value=2
			JavaWindow("DefaultWindow").winmenu(objMenu).Select(sMenu)
			Wait 1
			Set objMenu=Nothing
			Fn_MenuOperation = True
			'Returns true when the Menu Item is enabled
		'---------------------------------------This case is used to Check Window Menu Item State (Enabled)..---------------------------------------
		Case "WinMenuState"
			sMenu=Replace(StrMenuLabel,":",";")
			Set objMenu=description.create()
			objMenu("menuobjtype").value=2
			If (JavaWindow("DefaultWindow").winmenu(objMenu).GetItemProperty (sMenu,"enabled")) Then
				Fn_MenuOperation=True
			Else
				Fn_MenuOperation=False
			End If
			Set objMenu=Nothing
       '---------------------------------------This case is used to Check Window Menu Item State (Checked)..---------------------------------------
		'Returns true when the Item state is Checked
		Case "WinMenuCheck"
			sMenu=Replace(StrMenuLabel,":",";")
			Set objMenu=description.create()
			objMenu("menuobjtype").value=2
			bReturn = JavaWindow("DefaultWindow").winmenu(objMenu).CheckItemProperty(sMenu,"checked","1",1)
			If bReturn Then
				Fn_MenuOperation=True
			Else
				Fn_MenuOperation=False
			End If
			Set objMenu=Nothing
		'---------------------------------------This case is used to Check Window Menu Item State (Checked)..---------------------------------------
		'Returns true when the Item state is UnChecked
		Case "WinMenuUnCheck"
			sMenu=Replace(StrMenuLabel,":",";")
			Set objMenu=description.create()
			objMenu("menuobjtype").value=2
			If (JavaWindow("DefaultWindow").winmenu(objMenu).CheckItemProperty(sMenu, "checked","1",1)) Then
				Fn_MenuOperation=False
			Else
				Fn_MenuOperation=True
			End If
			Set objMenu=Nothing
		'---------------------------------------This case is used  for those menu operations that do not work with Select case---------------------------------------
		Case "SelectMenuObjects"
			aMenu = split(StrMenuLabel,":",-1,1)
			iMenuCount = Ubound(aMenu) 
			Set objMenu=description.create()
			objMenu("menuobjtype").value=2
			objMenu("label").value=aMenu(iMenuCount)		
			Set objChildObjects = JavaWindow("DefaultWindow").ChildObjects(objMenu)
			objChildObjects(0).Select
			If Err.Number < 0 Then
				Fn_MenuOperation = FALSE
			Else
				Fn_MenuOperation = True
			End If
			Set objChildObjects = Nothing
			Set objMenu=Nothing
		Case Else
			Fn_MenuOperation = FALSE
			Reporter.ReportEvent micPass,"Menu operation","Invalid option"
		End Select

End Function


'*********************************************************		Function to exit from Teamcenter		***********************************************************************
'Function Name		:				Fn_TeamcenterExit

'Description			 :		 		 Close down Tc session / Logout

'Parameters			   :	 			Nill

'Return Value		   : 				TRUE \ FALSE

'Pre-requisite			:		 		should be logged in

'Examples				:				 Fn_TeamcenterExit()

'History					 :		
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'												------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Jhuma Debnath										   10/03/2010			              1.0										Created
'													Sandeep Navghane								  25/11/2011			              1.1							Added Code to Handle BMIDE Window
'												------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function Fn_TeamcenterExit()
GBL_FAILED_FUNCTION_NAME="Fn_TeamcenterExit"
 Dim objJavaWindowDefault, objJavaWindowExit, oDesc, intNoOfObjects
 Dim sWinTitle
 Set objJavaWindowDefault = Fn_UI_ObjectCreate("Fn_TeamcenterExit",JavaWindow("DefaultWindow"))
	'Added by Vallari - Stop Code Coverage
	Call Fn_CodeCover_Exit()

 ' If JavaWindow("DefaultWindow").Exist(iTimeOut) Then 
   'Select menu [File -> Exit]
   Call Fn_MenuOperation("WinMenuSelect","File:Exit")
   Set objJavaWindowExit = Fn_UI_ObjectCreate("Fn_TeamcenterExit", JavaWindow("DefaultWindow").JavaWindow("Exit"))
     'If JavaWindow("DefaultWindow").JavaWindow("Exit").Exist(10) Then
   'Click on [Yes] button
	If Fn_UI_ObjectExist("Fn_TeamcenterExit",JavaWindow("DefaultWindow").JavaWindow("Exit"))=True Then
		Set oDesc = Description.Create()
		oDesc("Class Name").value = "JavaStaticText"
		Set intNoOfObjects = JavaWindow("DefaultWindow").JavaWindow("Exit").ChildObjects(oDesc)
		If intNoOfObjects(0).GetRoProperty("label") = "Is it ok to exit?" Then
			Call Fn_WriteLogFile("Fn_TeamcenterExit","PASS:Message 'Is it ok to exit?' Verified")
			If JavaWindow("DefaultWindow").JavaWindow("Exit").JavaCheckBox("Always exit without prompt").Exist Then
				Call Fn_WriteLogFile("Fn_TeamcenterExit","PASS: CheckBox 'Always exit without prompt' Exist")
			Else
				Call Fn_WriteLogFile("Fn_TeamcenterExit","Fail: CheckBox 'Always exit without prompt' does not Exist")
			End If
		Else
			Call Fn_WriteLogFile("Fn_TeamcenterExit","Fail:Message 'Is it ok to exit?' not Verified")
		End If
		Call Fn_Button_Click("Fn_TeamcenterExit", objJavaWindowExit, "Yes") 
    End If 
    'If Teamceneter window exists after log out
	'If JavaWindow("DefaultWindow").Exist(10) Then
	If Fn_UI_ObjectExist("Fn_TeamcenterExit",JavaWindow("DefaultWindow")) = True Then
		sWinTitle = JavaWindow("DefaultWindow").GetROProperty("title")
		If instr(sWinTitle, "Business Modeler IDE") > 0 Then
			Fn_TeamcenterExit = True
			Set objJavaWindowDefault =Nothing
			Set objJavaWindowExit =Nothing
			Exit Function
		Else
			Fn_TeamcenterExit = FALSE  
		End If
	Else
		Fn_TeamcenterExit = TRUE  
	End If
End Function

'*********************************************************		Function to Fn_CodeCover_Init()		***********************************************************************
'Function Name		:			Fn_CodeCover_Init()
'*********************************************************		Function to Fn_CodeCover_Init()		***********************************************************************									   
Public Function Fn_CodeCover_Init()
    Dim bReturn
	Dim App,userVarPath,bFlag
	sTestStack = "pv.rep:module.seq:" + Environment.Value("TestName") + ":"
	sRootDir = Fn_GetEnvValue("User", "AutomationDir")
	sBaseline = "tc10.1.0.2013060400"

	Fn_CodeCover_Init = False

	Err.Clear
	'- - - - - - - - - - - - - - - - - - - - - - - - Added Code to handle Scenarion where General.tsr is not associated with action- - - - - - - - - - - 
	Set App = CreateObject("QuickTest.Application")
	userVarPath	= Fn_GetEnvValue("User", "AutomationDir")
	bFlag=App.Test.Actions(Environment("ActionName")).ObjectRepositories.Find(userVarPath&"\ObjectRepository\General.tsr")
	'- - - - - - - - - - - - - - - - - - - - - - - - Added Code to handle Scenarion where General.tsr is not associated with action- - - - - - - - - - - 
	Set App =Nothing		
	If Cint(bFlag)<>-1 Then
		If JavaWindow("DefaultWindow").Exist(5) Then
			bReturn = Fn_MenuOperation("Exist", "Code Coverage")
			If bReturn Then
				Call Fn_MenuOperation("WinMenuSelect", "Code Coverage:Start Code Coverage")
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "This is not a Code Coverage Build to Start Code Coverage")
				Exit Function
			End If
			
			If  JavaWindow("DefaultWindow").JavaWindow("StartCodeCoverage").Exist(10) Then
				 JavaWindow("DefaultWindow").JavaWindow("StartCodeCoverage").JavaEdit("CCOV_TestStack").Set sTestStack
				 If Err.Number < 0 Then
					 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Set [TestStackName] on [Start Code Coverage] Window")
					 Exit Function
				 End If
				 JavaWindow("DefaultWindow").JavaWindow("StartCodeCoverage").JavaEdit("RootDirectory").Set sRootDir
				  If Err.Number < 0 Then
					 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Set [RootDirectory] on [Start Code Coverage] Window")
					Exit Function
				 End If
				 JavaWindow("DefaultWindow").JavaWindow("StartCodeCoverage").JavaEdit("CCOV_Baseline").Set sBaseline
				  If Err.Number < 0 Then
					 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Set [Baseline] on [Start Code Coverage] Window")
					 Exit Function
				 End If
				 javawindow ("DefaultWindow").JavaWindow("StartCodeCoverage").JavaButton("Start").Click micLeftBtn
				 If Err.Number < 0 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Click [Start] Butoon on [Start Code Coverage] Window")
					Exit Function
				 End If
			ElseIf JavaWindow("DefaultWindow").JavaWindow("Warn").Exist(5) Then
				JavaWindow("DefaultWindow").JavaWindow("Warn").JavaButton("OK").Click micLeftBtn
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[Start Code Coverage] Window Not Found")
				Exit Function
			End If
		Else
			'Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Can not Start Code Coverage, as Teamcenter Main Window not Present")
		End If
	End If

	Fn_CodeCover_Init = True
	
End Function

'*********************************************************		Function to CodeCover_Exit()	***********************************************************************
'Function Name		:			CodeCover_Exit()
'*********************************************************'*********************************************************'*********************************************************																																												   					  
Public Function Fn_CodeCover_Exit()
	On Error Resume Next

	Fn_CodeCover_Exit = False
	'Added by Nilesh to Handle Excel restart dialog
	Call Fn_ExcelErrorClose()

	Err.Clear

	If JavaWindow("DefaultWindow").Exist(1)  Then		
			if lcase(Environment("DetailLog")) = "console" and bConsoleLog = 0 then
				Call Fn_Setup_TestcaseExitWithConsoleLog("False")	
			End If 
		'bReturn = Fn_MenuOperation("Exist", "Code Coverage")			
		'If bReturn Then
		If JavaWindow("DefaultWindow").JavaMenu("label:=Code Coverage","index:=0").Exist(1) Then			
			Call Fn_MenuOperation("WinMenuSelect", "Code Coverage:Stop Code Coverage")
		Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "This is not a Code Coverage Build to Start/Stop Code Coverage")
			Exit Function
		End If

		If JavaWindow("DefaultWindow").JavaWindow("StopCodeCoverage").Exist(1) Then
			JavaWindow("DefaultWindow").JavaWindow("StopCodeCoverage").JavaButton("OK").Click micLeftBtn
			 If Err.Number < 0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Click [OK] Butoon on [Stop Code Coverage] Window")
				Exit Function
			 End If
		ElseIf JavaWindow("DefaultWindow").JavaWindow("Warn").Exist(1) Then
			JavaWindow("DefaultWindow").JavaWindow("Warn").JavaButton("OK").Click micLeftBtn
		Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[Stop Code Coverage] Window Not Found")
			Exit Function
		End If
	Else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Can not Start/Stop Code Coverage, as Teamcenter Main Window not Present")
	End If

	'**********************************************************************************
	'Deleting Snapshot Image file, snapshot is not needed if ALL VP PASS
	'**********************************************************************************
	Dim filePath, objFSO
	filePath = Environment.Value("BatchFldName") + "\" + Environment.Value("TestName") + ".png"
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	if objFSO.FileExists(filePath) then
		objFSO.DeleteFile filePath, True
	End if
	Set objFSO = Nothing
		
	Fn_CodeCover_Exit = True

End Function


'*********************************************  Function handle Error Dialog ..**************************************************************

'Function Name		:					Fn_SetTCSession

'Description			 :		 		  The function handles multiple Tc sessions

'Parameters			   :	 			1.  sUser : User whoes session needs to be activated

'Return Value		   : 				True/False

'Pre-requisite			:		 		Multiple Tc sessions with different users are exist

'Examples				:				 Fn_SetTCSession("AutoTest3")

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'											Vallari						10-Jun-2010	   		1.0
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function  Fn_SetTCSession(sUser)
		GBL_FAILED_FUNCTION_NAME="Fn_SetTCSession"
		Dim i, iWinCnt, sSession, objTcWin

		On Error Resume Next
		
		wait(5)

		sSession = ".*" + lcase(sUser) + ".*"
		
		For i= 1 to 120
			Set objTcWin = description.Create()
			objTcWin("Class Name").Value = "JavaWindow"
			objTcWin("title").Value = ".*- Teamcenter.*"
			objTcWin("title").RegularExpression = True
			iWinCnt = desktop.ChildObjects(objTcWin).Count
			If Cint(iWinCnt)=2 Then
				wait 2
				Exit For
			Else
				wait 1
			End If
			Set objTcWin =Nothing			
		Next
		
		Set objTcWin = description.Create()
		objTcWin("Class Name").Value = "JavaWindow"
		objTcWin("title").Value = ".*Teamcenter.*"
		objTcWin("title").RegularExpression = True
		iWinCnt = desktop.ChildObjects(objTcWin).Count
		
		For i= 0 to iWinCnt - 1
			JavaWindow("DefaultWindow").SetTOProperty "index", i
			JavaWindow("DefaultWindow").Maximize
			JavaWindow("DefaultWindow").JavaStaticText("SessionInfo").SetTOProperty "label", sSession
			
			If JavaWindow("DefaultWindow").JavaStaticText("SessionInfo").Exist(5) Then
				JavaWindow("DefaultWindow").highlight
				Fn_SetTCSession = True
				JavaWindow("MyTeamcenter").SetTOProperty "index", i
				JavaWindow("MyTeamcenter_Search").SetTOProperty "index", i
				JavaWindow("StructureManager").SetTOProperty "index", i
				JavaWindow("MyWorkListWindow").SetTOProperty "index", i
				wait 1
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Successfully Found Tc Session for User [" + sUser + "]")
				Exit for
			Else
			 	JavaWindow("DefaultWindow").Minimize          
			End If
		Next
		
		If i = iWinCnt Then
			Fn_SetTCSession = False
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to Find Tc Session for User [" + sUser + "]")
		End If
		Set objTcWin = Nothing		
End Function

'*********************************************************		Function to set perspective into Teamcenter		***********************************************************************
'Function Name		:				Fn_SetPerspectiveExtn

'Description			 :		 		 This function open the specified module.

'Parameters			   :	 			1.StrModule : Name of the module.

'Return Value		   : 				TRUE \ FALSE

'Pre-requisite			:		 		should be logged in

'Examples				:				 Fn_SetPerspectiveExtn("My Teamcenter")

'History					 :		
'													Developer Name				Date						Rev. No.			Changes Done						Reviewer
'												------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Harshal Agrawal					24June2010			2.0						None										Harshal Agrawal
'												------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function Fn_SetPerspectiveExtn(StrModule)

		Fn_SetPerspectiveExtn = Fn_SetPerspective(StrModule)
'            'Declearing Variables
'			Dim strModuleName, intRowCount, intCount,sInitial
'			Dim objTable, intSelIndex,objWindowOpenPerspective, objOpenPerspectiveTable
'			Dim iRootindex,StrTitle
'			
'			'Swapnil: Added for 4GD
'			
'			If StrModule = "Collaborative Product Development" Then
'				StrModule = "4G Designer"
'			End If
'		
'		   'Select menu ["Window - > Open Perspective -> Other...]
'			Call Fn_MenuOperation("Select","Window:Open Perspective:Other...")
'
'			'Setting the Object  for  Prespective Window
'			Set objWindowOpenPerspective=Fn_UI_ObjectCreate("Fn_SetPerspective",JavaWindow("DefaultWindow").JavaWindow("Open Perspective"))
'		
'		   'Maximize the open perspective window
'			Call Fn_Window_Maximize("Fn_SetPerspective",objWindowOpenPerspective)
'
'			'Setting the Object  Table inside Prespective Window
'        	Set objOpenPerspectiveTable=Fn_UI_ObjectCreate("Fn_SetPerspective",objWindowOpenPerspective.JavaTable("Table"))
'
'			'Select the First Row 
'			Call Fn_UI_JavaTable_SelectCell("Fn_SetPerspective", objWindowOpenPerspective, "Table","#0", "#0")
'
'			'Hit The initial  Letter of the Prespectrive to be set
'			Call Fn_JavaTable_Type("Fn_SetPerspective",objWindowOpenPerspective, "Table",Left(StrModule, 1))
'
'			'Setting the object for  Table methods
'			set objTable = objWindowOpenPerspective.JavaTable("Table").Object
'
'			'Getting the index of theFocused item
'			intSelIndex = objTable.getFocusIndex()
'
'			'Storing the initial  value of  Focused Index
'            sInitial = Fn_UI_JavaTable_GetCellData("Fn_SetPerspective", objWindowOpenPerspective, "Table",cstr(intSelIndex),"0")
'
'			'Get thing the value of  Focused Index
'	 		strModuleName = Fn_UI_JavaTable_GetCellData("Fn_SetPerspective", objWindowOpenPerspective, "Table",cstr(intSelIndex),"0")
'
'			'Do loop Starts
'			Do 
'				' If  the condition not matches then
'			If Trim(Lcase(strModuleName)) <> Trim(Lcase(StrModule)) Then
'					'Hit The initial  Letter of the Prespectrive to be set
'					Call Fn_JavaTable_Type("Fn_SetPerspective",objWindowOpenPerspective, "Table",Left(StrModule, 1))
'					'Getting the index of thecurrent Focused item
'					intSelIndex = objTable.getFocusIndex()
'			'Else
'			Else
'					'if matches Click on [OK] button
'					Call Fn_Button_Click("Fn_SetPerspective", JavaWindow("DefaultWindow").JavaWindow("Open Perspective"), "OK")
'					'Set the Return Value to TRUE
'					Fn_SetPerspectiveExtn = True
'					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [" + StrModule + "] Module selected successfully of Function Fn_SetPerspective" )
'					'Exit from Do Loop
'					Exit Do
'			'End if
'			End If
'
'				'Get thing the value of  Focused Index
'				strModuleName = Fn_UI_JavaTable_GetCellData("Fn_SetPerspective", objWindowOpenPerspective, "Table",cstr(intSelIndex),"0")
'
'		'End of Do Loop if Condition Matches
'		Loop Until Trim(Lcase(strModuleName))= Trim(Lcase(sInitial))
'
'		'if Condition Matches
'		If Trim(Lcase(strModuleName))= Trim(Lcase(sInitial)) Then
'
'				'Set the Return Value to TRUE
'				Fn_SetPerspectiveExtn =  True
'				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [" + StrModule + "] Module selected successfully of Function Fn_SetPerspective" )
'		'End if
'		End If
'
'		'Swapnil : By default in CPD perspective root object gets selected that is causing problem , so added code to deselect root object.	
'	
'		iRootindex = "0"
'		StrTitle = "4G Designer"
'		
'		If Instr(JavaWindow("DefaultWindow").GetROProperty("title"),cstr(StrTitle)) > 0 Then
'			' if Nav. tree is not empty.
'			If cInt(JavaWindow("Collaborative Product").JavaTree("NavTree").GetROProperty("objects count")) > 0 Then
'				JavaWindow("Collaborative Product").JavaTree("NavTree").Deselect cint(iRootindex)
'			End If
'		End  If	
'		
'		'release all Objects
'		Set objTable = Nothing
'		Set objWindowOpenPerspective=Nothing
'		Set objOpenPerspectiveTable=Nothing
End Function
'*********************************************  Function handle Error Dialog ..**************************************************************

'Function Name		:					Fn_SetMyTcSession

'Description			 :		 		  The function handles multiple Tc sessions

'Parameters			   :	 			1.  sUser : User whoes session needs to be activated

'Return Value		   : 				True/False

'Pre-requisite			:		 		Multiple Tc sessions with different users are exist

'Examples				:				 Fn_SetMyTcSession("AutoTest3")

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'											Vallari						09-Sept-2010	   		1.0
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function  Fn_SetMyTcSession(sUser)
		GBL_FAILED_FUNCTION_NAME="Fn_SetMyTcSession"
		Dim i, iWinCnt, sSession, objTcWin

		On Error Resume Next

		sSession = ".*" + lcase(sUser) + ".*"
		
		Set objTcWin = description.Create()

		objTcWin("Class Name").Value = "JavaWindow"
		objTcWin("title").Value = ".*Teamcenter.*"
		objTcWin("title").RegularExpression = True
		wait(2)
'		iWinCnt = objTcWin.Count
		iWinCnt = desktop.ChildObjects(objTcWin).Count
		
		For i= 0 to iWinCnt - 1
					JavaWindow("MyTeamcenter").SetTOProperty "index", i
					JavaWindow("MyTeamcenter").Maximize
					JavaWindow("MyTeamcenter").JavaStaticText("SessionInfo").SetTOProperty "label", sSession
					If JavaWindow("MyTeamcenter").JavaStaticText("SessionInfo").Exist(5) Then
							'Set the same index for MyWorklist window
							JavaWindow("MyWorkListWindow").SetTOProperty "index", i
							Fn_SetMyTcSession = True
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Successfully Found Tc Session for User [" + sUser + "]")
							Exit for
					Else
							 JavaWindow("MyTeamcenter").Minimize          
					End If
		Next
		If i = iWinCnt Then
				Fn_SetMyTcSession = False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to Find Tc Session for User [" + sUser + "]")
		End If

		Set objTcWin = Nothing
End Function


'*********************************************  Function Componentise Actiob1 in the Script ..**************************************************************

'Function Name		:					Fn_Setup_TestcaseInit

'Description			 :		 		  The function handles Test script initialization part

'Parameters			   :	 			

'Return Value		   : 				True/False

'Pre-requisite			:		 		None

'Examples				:
'

'History:
'										Developer Name			Date			Rev. No.			Changes Done												Reviewer	
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Vallari				04-Oct-2010	   		1.0
'										Shweta Rathod		29-Aug-2016			1.0					creating report folder and subreport folders for BB			Koustubh Watwe
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_Setup_TestcaseInit()

	'**********************************************************************************
	'Variables Declaration
	'**********************************************************************************	
	Dim iCnt
	Dim sLogFile
	Dim Result
	Dim bReturn
	ReDim TestData(8)

  	On Error Resume Next

	'**********************************************************************************
	'Set the SVN mainline path value to Script Variable
	'**********************************************************************************
    Environment.Value("sPath") = Fn_GetEnvValue("User", "AutomationDir")

	'**********************************************************************************
	'Import Environment Variable File
	'**********************************************************************************
	bReturn = LoadEnvXML()
	
	'**********************************************************************************
	'Run Action Only for First Row of DataTable
	'**********************************************************************************
	Call Fn_SetActionIterationMode("oneIteration")
	
	'**********************************************************************************
	'Call for Code Coverage
	'**********************************************************************************
	'Call Fn_CodeCover_Init()

    '**********************************************************************************
	'Create Test Log File & Write QART Specifi Info
	'**********************************************************************************
	'Create Log File
	sLogFile = Fn_CreateLogFile(Environment.Value("TestName") + ".log")
	If sLogFile <> "" Then
			Environment.Value("TestLogFile") = sLogFile
		Else
			Reporter.ReportEvent micFail, "LogFileCreate", "Failed to Create Test Log File"
			Fn_Setup_TestcaseInit = False
			Exit Function
	End If
	
	'**********************************************************************************
	'creating report folder and subreport folders for BB
	'**********************************************************************************
	If lcase(Environment.Value("BBFlag")) = "true" then
		sBBReportFolderPath = Fn_SISW_BB_Setup_CreateReportFolder(Environment.Value("BatchFldName") +"\"+Environment.Value("TestName"))
		If sBBReportFolderPath <> "" Then
			Environment.Value("BBReportFolderPath") = sBBReportFolderPath	
			'load BBXML file
			Call Fn_SISW_BB_Setup_LoadBBXML()
		Else
			Reporter.ReportEvent micFail, "LogFileCreate", "Failed to Create Test Log File"
			Fn_Setup_TestcaseInit = False
			Exit Function
		End If	
	End if
	
	'Print QART Derails in Test Log File
	Call Fn_PrintQARTLog()
	
	' Call added by Chandrakant Tyagi to get Feature Name of every testcase in global variable "sFeatureName" --------------- TC1015-2015071400-31_07_2015
	Call Fn_GetFeatureNameOfTestCase()
	'------------------------------------------------------------------------------------------------------------------------
	'Write Test Details to Batch Log File
	TestData(1) = Environment.Value("QARTRoot")
	TestData(2) = Environment.Value("QARTRelease")
	TestData(3) =DataTable("Feature", dtGlobalSheet)
	TestData(4) = DataTable("Category", dtGlobalSheet)
	TestData(5) =  Environment.Value("TestName")
	TestData(6) = Date & " - " & Time
	
	Result = Fn_Update_TestDetail( "", TestData, 1)
	
	'============'Flag set for CMS TCs=======
	If Instr(TestData(3),"Multi-Site") > 0 Then
		bSiteReset = True
	End If
	'========================================
	'Print Action Execution Start
	Call Fn_UpdateLogFiles("[" + Cstr(time) + "]  -  QTP [" + Environment.Value("ActionName") + "] - Start", "")

	Fn_Setup_TestcaseInit = True

End Function

'*********************************************  Function Componentise Actiob1 in the Script ..**************************************************************

'Function Name		:					Fn_Setup_TestcaseExit

'Description			 :		 		  The function handles test script end part

'Parameters			   :	 			

'Return Value		   : 				True/False

'Pre-requisite			:		 		None

'Examples				:
'

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'											Vallari						05-Oct-2010	   		1.0
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function Fn_Setup_TestcaseExit(bTcKill)

		On Error Resume Next
		Dim timeTaken

		''**********************************************************************************
		''Kill Teamcenter if Flag is True
		''**********************************************************************************
		If Cbool(bTcKill) Then
			'Added by Vallari - Closing logged in Perspective, to get Tc default state
			Call Fn_MenuOperation("Select","File:Close")
			'Call Fn_TeamcenterExit()
			Call Fn_KillProcess("")
		End If
		
		'**********************************************************************************
		'Log Test Result
		''**********************************************************************************
		Call Fn_UpdateLogFiles("[" + Cstr(time) + "] - QTP [" + Environment.Value("ActionName") + "] - End", "")
		Call Fn_UpdateLogFiles("-----------------------------------------------------------------------------------------------", "")
		If bCaptureImgVP = True Then
			Call Fn_UpdateLogFiles("[" + Cstr(time) + "] - Final - Pass | Test Execution Result: PASS without comparing images", "PASS: All VP Pass without comparing images")
		Else
			Call Fn_UpdateLogFiles("[" + Cstr(time) + "] - Final - Pass | Test Execution Result: PASS", "PASS: All VP Pass")
		End If
		Call Fn_UpdateLogFiles("-----------------------------------------------------------------------------------------------", "")
		'------------- Set TC Execution End Time [PoonamC_26Dec2016]---------------
		Environment.Value("TCEndTime") = now()
		timeTaken = Fn_SecondsToMinutes(DateDiff("s", Environment.Value("TCStartTime"), Environment.Value("TCEndTime")))
		Call Fn_UpdateLogFiles("Total Test Execution Time : " & timeTaken,"")
		Call Fn_UpdateLogFiles("-----------------------------------------------------------------------------------------------", "")
		'**********************************************************************************
		'Deleting Snapshot Image file, snapshot is not needed if ALL VP PASS
		'**********************************************************************************
		Dim filePath, objFSO
		filePath = Environment.Value("BatchFldName") + "\" + Environment.Value("TestName") + ".png"
		Set objFSO = CreateObject("Scripting.FileSystemObject")
		if objFSO.FileExists(filePath) then
			objFSO.DeleteFile filePath, True
		End if
		Set objFSO = Nothing
		'**********************************************************************************
		'Call for Code Coverage
		'**********************************************************************************
		Call Fn_CodeCover_Exit()		
End Function

'*********************************************************		Function to exit from Current Session of Teamcenter		***********************************************************************
'Function Name		:				Fn_TeamcenterExit_Extn

'Description			 :		 		 Close down the Active or Current Tc session / Logout without  Closing Concurrent or Previous Session of Teamcenter

'Parameters			   :	 			Nill

'Return Value		   : 				TRUE \ FALSE

'Pre-requisite			:		 		Should be logged in

'Examples				:				 Fn_TeamcenterExit_Extn()

'History					 :		
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'												Omkar Kulkarni												   24/11/2010			           	  1.0									Created
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_TeamcenterExit_Extn()
 Dim objJavaWindowDefault, objJavaWindowExit, oDesc, intNoOfObjects,bReturn
 Set objJavaWindowDefault = JavaWindow("DefaultWindow")

	   If JavaWindow("DefaultWindow").Exist(10) Then
		   'Added by Vallari - Stop Code Coverage
			Call Fn_CodeCover_Exit()
		 'Select Menu [File -> Exit]	
			bReturn= Fn_MenuOperation("Select","File:Exit")
			If bReturn = False Then								
						Call Fn_WriteLogFile("Fn_TeamcenterExit_Extn", "FAIL : Failed to Perform Menu Operation File ->Exit")
						 Set objJavaWindowDefault = Nothing
						Exit Function
			Else
						Call Fn_ReadyStatusSync(3)
						Call Fn_WriteLogFile("Fn_TeamcenterExit_Extn", "PASS : Successfully Performed Menu Operation File ->Exit ")
			End If	


		    Set objJavaWindowExit = JavaWindow("DefaultWindow").JavaWindow("Exit")
			 'If JavaWindow("DefaultWindow").JavaWindow("Exit").Exist(10) Then
		   'Click on [Yes] button
			If JavaWindow("DefaultWindow").JavaWindow("Exit").Exist(10) Then


				'Set oDesc = Description.Create()
				'oDesc("Class Name").value = "JavaStaticText"
				'Set intNoOfObjects = JavaWindow("DefaultWindow").JavaWindow("Exit").ChildObjects(oDesc)
				'If intNoOfObjects(0).GetRoProperty("label") = "Is it ok to Exit?" Then
								'Call Fn_WriteLogFile("Fn_TeamcenterExit_Extn","PASS:Message 'Is it ok to exit?' Verified")	 
								 'If 	JavaWindow("DefaultWindow").JavaWindow("Exit").JavaCheckBox("Always exit without prompt").Exist Then
												'Call Fn_WriteLogFile("Fn_TeamcenterExit_Extn","PASS: CheckBox 'Always exit without prompt' Exist")

									'Else
												'Call Fn_WriteLogFile("Fn_TeamcenterExit_Extn","FAIL: CheckBox 'Always exit without prompt' does not Exist")


								 'End If
				'Else		
											'Call Fn_WriteLogFile("Fn_TeamcenterExit_Extn","FAIL:Message 'Is it ok to exit?' not Verified")

				'End If
				 'Click on Yes button
				 JavaWindow("DefaultWindow").JavaWindow("Exit").JavaButton("Yes").Click micLeftBtn 												
				If Err.Number < 0 Then
							Call Fn_WriteLogFile("Fn_TeamcenterExit_Extn", "Failed to Click on Yes Button")

							Fn_TeamcenterExit_Extn = False                                                              																
							Exit Function

				Else
							Call Fn_WriteLogFile("Fn_TeamcenterExit_Extn", "Successfully Clicked on Yes Button")	

				End If          														   
			Else
				 Fn_TeamcenterExit_Extn = FALSE
				Call Fn_WriteLogFile("Fn_TeamcenterExit_Extn","FAIL:'Exit' Window does not exist")
				Exit Function
			End If 


			Fn_TeamcenterExit_Extn = TRUE 
			Call Fn_WriteLogFile("Fn_TeamcenterExit_Extn", "PASS : Successfully Closed TeamCenter Session")
		Else
			 Fn_TeamcenterExit_Extn = FALSE
			Call Fn_WriteLogFile("Fn_TeamcenterExit_Extn","FAIL:'Default Window' does not exist")
			Exit Function										 
		End If 

		 Set objJavaWindowDefault =Nothing
		 Set objJavaWindowExit =Nothing
		 Set oDesc=Nothing
		 Set intNoOfObjects=Nothing			
End Function


'#########################################################################################################
'###
'###    FUNCTION NAME   :   Fn_UserSessionSettings(sAction, sGroup, sRole, sProject, sWorkContext, sVolume, bAppLogging, bJournalling, bSecLogging, bBypass)
'###
'###    DESCRIPTION        :   This function allows the user to switch the session privileges from Engineering(Normal User) to DBA.
'###
'###    PARAMETERS      :   sAction: Select option for setting the session setting or Verify.
'###             								sGroup: Group to be selected
'###  									         sRole: Role to be selected
'###  									         sProject: Project to be selected
'###  									         sWorkContext: WorkContext to be selected only if a project is assigned [ Not Implemented in Function Yet]
'###        								   sVolume: Volume to be selected
'### 								          bAppLogging: Flag to set be set Yes/No
'### 								          bJournalling: Flag to be set Yes/No
'### 								          bSecLogging: Flag to be set Yes/No
'###  								         bBypass: Flag to be set Yes/No
'###
'###    RETURNS      		 :   True / False
'###
'###  HISTORY             		 :   AUTHOR                       DATE        VERSION
'###
'###    CREATED BY      :    Mahendra Bhandarkar           26/05/2010         1.0
'###    REVIWED BY      :    Mohit Khare					26/05/2010   1.0
'###
'###	Modified by		:	Harshal Agrawal					11/06/2010			2.0
'###
'###	Modified by		:	Jeevan Mutha					08/06/2012			3.0			Modified object hierarchy
'###    REVIWED BY      :   Koustubh Watwe					08/06/2012			3.0
'###	Modified by      : Dipali							19/02/2013						Modified case IsEnabled 	
'###
'###    EXAMPLE           	: Call  Fn_UserSessionSettings("Set",  "dba", "DBA", "", "", "volume1", "ON", "ON", "ON","ON")
'###						Call Fn_UserSessionSettings("Verify","dba","DBA","","","volume1","0","1","","")
'###						Call Fn_UserSessionSettings("Set","dba","DBA","none","","volume1","0","1","","") ||None is used to set the Project value to blank||
'###						Call  Fn_UserSessionSettings("IsEnable",1, 1, 1 ,"","","","","","")|| 1 for enable and 0 for disable||
'#############################################################################################################
Public Function Fn_UserSessionSettings(sAction, sGroup, sRole, sProject, sWorkContext, sVolume, bAppLogging, bJournalling, bSecLogging, bBypass)
	GBL_FAILED_FUNCTION_NAME="Fn_UserSessionSettings"
	Dim objDialog, bFlag, ObjShell,sRACMenuFilePath,sNewItemMenu
	bFlag = True
	
	Set ObjShell = CreateObject("WScript.Shell")
	Set objDialog = JavaWindow("DefaultWindow").JavaWindow("UserSettings")
	
	sRACMenuFilePath = Fn_LogUtil_GetXMLPath("RAC_Menu")
	sNewItemMenu = Fn_GetXMLNodeValue(sRACMenuFilePath,"UserSettings")

	If JavaWindow("DefaultWindow").JavaWindow("UserSettings").Exist(SISW_MINLESS_TIMEOUT)  = False Then
		 Call Fn_MenuOperation("Select",sNewItemMenu)
		Call Fn_ReadyStatusSync(1)
	End If
	
	If JavaWindow("DefaultWindow").JavaWindow("UserSettings").Exist(SISW_MINLESS_TIMEOUT)  Then
		Set objDialog = JavaWindow("DefaultWindow").JavaWindow("UserSettings")
	Else 
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Fn_UserSessionSettings : Failed to find [ User Settings ] window")
		Fn_UserSessionSettings = False
		Exit Function
	End If

	objDialog.Activate
	objDialog.JavaTab("MainTab").Select "Session"	

	Select Case sAction
	 Case "Set"
		 If sGroup<>"" Then
			   Call Fn_List_Select("Fn_UserSessionSettings", objDialog, "Group", sGroup)
		 End If
		 If sRole<>"" Then
			   Call Fn_List_Select("Fn_UserSessionSettings", objDialog, "Role", sRole)
		 End If
		 If sVolume<>"" Then
			   Call Fn_List_Select("Fn_UserSessionSettings", objDialog, "Volume", sVolume)
		 End If
		 If bAppLogging<>"" Then
				objDialog.JavaTab("MainTab").Select "Administrative"
			   Call Fn_CheckBox_Set("Fn_UserSessionSettings", objDialog, "ApplicationLogging", bAppLogging)
		 End If
		If bJournalling<>"" Then
				objDialog.JavaTab("MainTab").Select "Administrative"
			   Call Fn_CheckBox_Set("Fn_UserSessionSettings", objDialog, "Journalling", bJournalling)
		End If
		   Wait 2
		   If Trim(sProject) <> "" Then
			   objDialog.JavaTab("MainTab").Select "Session"
				If Trim(sProject) = "none" Then
'					objDialog.JavaList("Project").Select ""		-----------------------------------------------------Code doesn't work for setting the value as blank & repopulates the previous value
					objDialog.JavaList("Project").Select 0
					ObjShell.SendKeys "{DELETE}"
					'Call Fn_List_Select("Fn_UserSessionSettings",objDialog, "Project", "")||UI is not working||
				Else
					Call Fn_List_Select("Fn_UserSessionSettings",objDialog, "Project", sProject)
				End If
		   End If
		   'Call Fn_Button_Click("Fn_UserSessionSettings", objDialog, "OK")
		   If  Trim(bSecLogging) <> "" OR Trim(bBypass) <> "" Then
				'Call Fn_List_Select("Fn_UserSessionSettings",objDialog, "Steps", "Administrative")
				objDialog.JavaTab("MainTab").Select "Administrative"
				Call Fn_CheckBox_Set("Fn_UserSessionSettings", objDialog, "Bypass", bBypass)
				Call Fn_CheckBox_Set("Fn_UserSessionSettings", objDialog, "SecurityLogging", bSecLogging)
				'Call Fn_Button_Click("Fn_UserSessionSettings", objDialog, "Apply")
		   End If
		   'Sandeep [ 06-Feb-2012 ]
		   'Call Fn_Button_Click("Fn_UserSessionSettings", objDialog, "OK")
	
	Case "Verify"   'Modified By Anjali 2-Jan-13
	   If sGroup<>"" Then
'				 If objDialog.JavaEdit("Group").GetROProperty("text") <> sGroup Then
				 If Fn_UI_Object_GetROProperty("",objDialog.JavaList("Group"),"value") <> sGroup Then
				'If objDialog.JavaList("Group").GetROProperty("value") <> sGroup Then
					bFlag = False
			   End If
			End If
		  If sRole<>"" Then
'			   If objDialog.JavaEdit("Role").GetROProperty("text") <> sRole Then
				 If Fn_UI_Object_GetROProperty("",objDialog.JavaList("Role"),"value") <> sRole Then
			  'If objDialog.JavaList("Role").GetROProperty("value") <> sRole Then
					bFlag = False
			   End If
		  End If
		  If sVolume<>"" Then
'				If objDialog.JavaEdit("Volume").GetROProperty("text") <> sRole Then
				If objDialog.JavaList("Volume").GetROProperty("value") <> sVolume Then
					bFlag = False
				End If
		 End If
		 If bAppLogging<>"" Then
'			   objDialog.JavaTab("MainTab").Select "Administrative" 'Modified by Nilesh on 16-Jul-2013
               objDialog.JavaTab("MainTab").Select "Administrative"
			   If objDialog.JavaCheckBox("ApplicationLogging").GetROProperty("value") <> bAppLogging Then
				bFlag = False
			  End If
		 End If
		 If bJournalling<>""  Then
'			   objDialog.JavaTab("MainTab").Select "Administrative" 'Modified by Nilesh on 16-Jul-2013
			   objDialog.JavaTab("MainTab").Select "Administrative"
			   If objDialog.JavaCheckBox("Journalling").GetROProperty("value") <> bJournalling Then
					bFlag = False
			  End If
		 End If
	     If Trim(sProject) <> "" Then
			   objDialog.JavaTab("MainTab").Select "Session"
				If Trim(sProject) = "none" Then
					objDialog.JavaList("Project").Select ""
					'Call Fn_List_Select("Fn_UserSessionSettings",objDialog, "Project", "")||UI is not working||
				Else
					Call Fn_List_Select("Fn_UserSessionSettings",objDialog, "Project", sProject)
				End If
	     End If

		 If  Trim(bSecLogging) <> "" OR Trim(bBypass) <> "" Then
'			Call Fn_List_Select("Fn_UserSessionSettings",objDialog, "Steps", "Administrative")
			objDialog.JavaTab("MainTab").Select "Administrative"
			If objDialog.JavaCheckBox("Security Administrative").GetROProperty("value") <> bSecLogging Then
				bFlag = False
			End If
			If objDialog.JavaCheckBox("bBypass").GetROProperty("value") <> bBypass Then
				bFlag = False
			End If
		 End If
	
	Case "IsEnable"
		If sGroup<>"" Then
			If cint(objDialog.JavaList("Group").GetROProperty("enabled")) <> cint(sGroup) Then
				bFlag = False
			 End If
		End If
	
		 If sRole<>"" Then
			If cint(objDialog.JavaList("Role").GetROProperty("enabled")) <> cint(sRole) Then
				bFlag = False
			End If
		End If
	
		If sVolume<>"" Then
			If cint(objDialog.JavaList("Volume").GetROProperty("enabled")) <> cint(sVolume) Then
				bFlag = False
			End If
		End If
	
		If bAppLogging<>"" Then
			If cint(objDialog.JavaCheckBox("ApplicationLogging").GetROProperty("enabled")) <> cint(bAppLogging) Then
				bFlag = False
			End If
		End If
	
		If bJournalling<>""  Then
			If cint(objDialog.JavaCheckBox("Journalling").GetROProperty("enabled")) <> cint(bJournalling) Then
				bFlag = False
			End If
		End If
	
		If sProject<>"" Then
			If cint(objDialog.JavaList("Project").GetROProperty("enabled")) <> cint(sProject) Then 'NOT CHANGED  modify if needed
				bFlag = False
			End If
		End If
	End Select

	If objDialog.JavaButton("OK").GetROProperty("enabled") = 1 Then
			Call Fn_Button_Click("Fn_UserSessionSettings", objDialog, "OK")  
	End If

	If JavaWindow("DefaultWindow").JavaWindow("Shell").JavaWindow("No Assign Privilege").Exist(3) Then
		Call Fn_Button_Click("Fn_UserSessionSettings", JavaWindow("DefaultWindow").JavaWindow("Shell").JavaWindow("No Assign Privilege"), "OK")
	End If
	If JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("No Assign Privilege").Exist(3) Then
		Call Fn_Button_Click("Fn_UserSessionSettings", JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("No Assign Privilege"), "OK")
	End If

	If objDialog.Exist(3) Then
		objDialog.Close
	End If
	If bFlag = False Then
			Fn_UserSessionSettings = False
	Else
			Fn_UserSessionSettings = True
	End If
	Set objDialog = Nothing
	Set bFlag = Nothing
End Function


'*********************************************************		Function to Creates Log File and folder.		***********************************************************************
'Function Name		:				Fn_LHSModuleLinkoperation(sAction, sTabName)

'Description			 :		 		 LHS Tab operation

'Parameters			   :	 			1. sAction : Name of the Action to perform

'										2.	sTabName	-  Tab name

'Return Value		   : 			True/False

'Pre-requisite			:		 	Nothing

'Examples				:			Call Fn_LHSModuleLinkoperation("LHSModuleSelect", "My Teamcenter")

'History					 :		
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'												------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Vallari		 											24/03/2011			              1.0										Created
'												------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public function Fn_LHSModuleLinkoperation(sAction, sTabName)
	Dim objAppTab
	Dim bReturn
	Dim sTabText
	Dim iCnt, i

	Select Case sAction
		Case "LHSModuleSelect"
			Set objAppTab = description.Create()
			objAppTab("Class Name").value = "JavaObject"
			objAppTab("toolkit class").value = "com.teamcenter.rac.aif.console.PrimaryButton"
			iCnt = JavaWindow("DefaultWindow").ChildObjects(objAppTab).count
		
			For i = 0 to iCnt - 1
				JavaWindow("DefaultWindow").JavaObject("LHNMyteamcenter").SetTOProperty "index", cstr(i)
				sTabText = JavaWindow("DefaultWindow").JavaObject("LHNMyteamcenter").Object.getText
				If trim(lcase(sTabText)) = trim(lcase(sTabName)) Then
					JavaWindow("DefaultWindow").JavaObject("LHNMyteamcenter").Click 1,1,"LEFT"
					Fn_LHSModuleLinkoperation = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Clicked on LHS Tab [" + sTabName + "]")
					Exit For
				End If
			Next
		
			If i = iCnt Then
				Fn_LHSModuleLinkoperation = False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Clicked on LHS Tab [" + sTabName + "]")
			End If
		
			Set objAppTab = Nothing
	End Select

End Function
'-------------------------------------------------------------------Function Used to Create Random Number-------------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_Setup_RandNoGenerate

'Description			 :	Function Used to Create Random Number

'Return Value		   : 	Random Number

'Parameters     		:	1. iLength : Length Of Random Number

'Examples				: 	'Call Fn_Setup_RandNoGenerate(2)
										'Call Fn_Setup_RandNoGenerate(3)
										'Call Fn_Setup_RandNoGenerate(4)
										'Call Fn_Setup_RandNoGenerate(5)
										'Call Fn_Setup_RandNoGenerate(6)
										'Call Fn_Setup_RandNoGenerate(7)

'History					 :			
'		Developer Name												Date						Rev. No.						Changes Done						Reviewer
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'		Sandeep Navghane										19-Apr-2011			           1.0												-							      Sunny Ruparel
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_Setup_RandNoGenerate(iLength)
	Dim iNumber,iStartNumber, sInitZeros, iCount
    Randomize
	sInitZeros = "0"
	iStartNumber="9"
	For iCount = 1 To iLength-1
		iStartNumber=Cstr(iStartNumber)+"0"
	Next
	' Generate random value between 1 and iStartNumber. 
	iNumber = Int((iStartNumber * Rnd) + 1)
 	If Len(Cstr(iNumber)) < iLength Then
			For iCount = Len(Cstr(iNumber))+1 To iLength-1
				sInitZeros = sInitZeros & "0"
			Next
			Fn_Setup_RandNoGenerate = sInitZeros & Cstr(iNumber)
	 Else
			Fn_Setup_RandNoGenerate = Cstr(iNumber)
	 End If
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Generated Random Number is [" & Fn_Setup_RandNoGenerate & "]")		
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_Setup_NetworkDriveOperations

'Description			 :	Function Used to Share & Unshare network drives

'Parameters			   :   '1.StrAction: Action Name
'										 2.StrSharePath : Network path which is to share to local drive
'										 3.StrDriveLetter : Local share drive letter
'										 4.StrUsername : Username for login to server
'										 5.StrPassword : Password for login to server

'Return Value		   : 	True Or False

'Examples				:   Fn_Setup_NetworkDriveOperations("Share","\\pnv6s108\Siemens" ,"Z:","autoadmin","Password123")
'										Fn_Setup_NetworkDriveOperations("Unshare","\\pnv6s108\Siemens" ,"Z:","","")
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												19-Dec-2011								1.0																						Sunny R
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Swapnil G												15-May-2013								1.0							used net use command																						Sunny R
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Public Function Fn_Setup_NetworkDriveOperations(StrAction,StrSharePath,StrDriveLetter,StrUsername,StrPassword)
	Dim objNetwork,objFSO
	Fn_Setup_NetworkDriveOperations=False
	'Creating Object Of  [ Network ] & [ FileSystemObject ]
    Set objNetwork = CreateObject("WScript.Network") 
	'Set objFSO = CreateObject("Scripting.FileSystemObject")
	Select Case StrAction
		'- - - - - - - - - - - - - - Case to share Drive from Network - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "Share"
			'If objFSO.FolderExists(StrDriveLetter) Then
			'	objNetwork.RemoveNetworkDrive StrDriveLetter,True,True
			'End If
			'objNetwork.MapNetworkDrive StrDriveLetter, StrSharePath,True,StrUsername,
			
			Set objNetwork = CreateObject("WScript.shell")
            objNetwork.Run "cmd.exe /c net use "&StrDriveLetter&" "&" /delete /y & timeout /t 5", 1, true
			wait 4
            objNetwork.Run "cmd.exe /c net use "&StrDriveLetter&" "&StrSharePath&" "&StrPassword&" /USER:"&StrUsername &" & timeout /t 5", 1, true
			
			If Err.Number < 0 Then
				Fn_Setup_NetworkDriveOperations=False
			Else
				Fn_Setup_NetworkDriveOperations=True
			End If
			
		'- - - - - - - - - - - - - - Case to unshare Drive from Network - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "Unshare"
		
			Set objNetwork = CreateObject("WScript.shell")
            objNetwork.Run "cmd.exe /c net use "&StrDriveLetter&" "&" /delete /y & timeout /t 5", 1, true
			wait 1
            If Err.Number < 0 Then
				Fn_Setup_NetworkDriveOperations=False
			Else
				Fn_Setup_NetworkDriveOperations=True
			End If
	End Select
	'Releasing Object Of [ Network ] & [ FileSystemObject ]
	Set objNetwork =Nothing
	'Set objFSO = Nothing
End Function

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_GetTestCaseDetailsFromExcel

'Description			:	Function Used to set Feature and Category from specified excel file

'Parameters			    :   '1.sFilePath : Excel file name

'Return Value		    : 	True Or False

'Examples				:   Call Fn_GetTestCaseDetailsFromExcel("d:\test1.xlsx")
'History				:			
'					Developer Name					Date					Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'					Koustubh Watwe					20-Dec-2011				1.0																						Sunny R
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Public Function Fn_GetTestCaseDetailsFromExcel(sFilePath)
		'-----------------------------------------
		' Variable Declaration
		'-----------------------------------------
		Dim loopCount, objExcel, workbook
		DataTable.SetCurrentRow 1
		Fn_GetTestCaseDetailsFromExcel = False
		If sFilePath = "" Then
			Environment.Value("sPath") = Fn_GetEnvValue("User", "AutomationDir")
			sFilePath = Environment.Value("sPath")+"\TestData\PSE_Feature.xls"
		End If
		
		Set objExcel = CreateObject("Excel.Application")
		objExcel.Visible = False 
		objExcel.DisplayAlerts = 0
			
		Set workbook = objExcel.Workbooks.Open(sFilePath)
		loopCount = 1
		Do while not isempty(objExcel.Cells(loopCount, 1).Value)
			If  objExcel.Cells(loopCount, 3).Value = Environment.Value("TestName") Then
				DataTable("Category", dtGlobalSheet) = objExcel.Cells(loopCount, 2).Value
				DataTable("Feature", dtGlobalSheet) = objExcel.Cells(loopCount, 1).Value
				Fn_GetTestCaseDetailsFromExcel = True
				Exit do
			End If
			loopCount  = loopCount +1
		Loop
		
		objExcel.Workbooks.Close
		objExcel.quit
		objExcel = Empty
		workbook = Empty
End Function
'*********************************************  Function Componentise Actiob1 in the Script ..**************************************************************

'Function Name		:					Fn_Setup_TestcaseExitWithConsoleLog

'Description			 :		 		  The function handles test script end part with console log.

'Return Value		   : 				True / False

'Pre-requisite			:		 		None

'Examples				:
'

'History:
'										Developer Name				Date				Rev. No.			Changes Done			Reviewer	
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Koustubh						23-Jan-2011	   		1.0							Created
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function Fn_Setup_TestcaseExitWithConsoleLog(bTcKill)	
Dim sConsoleMessage, sPassStr, objFile, objFSO
Dim sLogFile,sBaseline,sLogFolder,sTimeStamp

On Error Resume Next		
	Fn_Setup_TestcaseExitWithConsoleLog = False
	Set objFSO = CreateObject("Scripting.FileSystemObject")				
	
	Call Fn_SyncTCObjects()
	''Checking existence of Teamcenter window, will exist on non-existence of teamcenter window	
	If cBool(JavaWindow("BMIDEWindow").Exist(2)) = true Then
        call Fn_UI_Object_SetTOProperty("",JavaWindow("DefaultWindow"),"index",1)
		'JavaWindow("TeamcenterWindow").SetTOProperty("index",1)
		If JavaWindow("DefaultWindow").exist(3) = False Then			
			Fn_Setup_TestcaseExitWithConsoleLog = False
			Exit function
		End If 
	End if	
	
	' getting data from Console, console errors
	sConsoleMessage = trim(Fn_ConsoleOperations("OpenGetContentsAndClose","",""))
	
	' syncing objects after closing Console tab
	'Call Fn_SyncTCObjects()
	If sConsoleMessage = False or (len(sConsoleMessage) < 1)  = true Then	
		If bTcKill = True Then
			Call Fn_UpdateLogFiles("[" + Cstr(time) + "] - Final - Pass | Test Execution Result: PASS without console log" , "PASS: All VP Pass")
			Call Fn_MenuOperation("Select","File:Close")			
			Call Fn_TeamcenterExit()
		End If
		Fn_Setup_TestcaseExitWithConsoleLog = True
		Exit function
'	Else
'		Call Fn_ConsoleOperations("CloseTab", "", "")
	End If
	
	If ( instr(lcase(sConsoleMessage),"error") > 0 ) OR ( instr(lcase(sConsoleMessage),"exception") > 0 ) Then
		sPassStr = sPassStr & " with Console errors."
	End If
	
	sBaseline = Environment("TcRelease") & "." & Environment("TcBuild")
	If cBool(bTcKill)= False Then
		sLogFolder = "\\pnv6s1324\Automation_Share\TcConsoleLogs\" + sBaseline
	Elseif cBool(bTcKill) = True and lcase(Environment("DetailLog")) = "console" then 
		sLogFolder = Environment.Value("BatchFldName")
	End If 
	
	If objFSO.FolderExists(sLogFolder) = False Then
		objFSO.CreateFolder(sLogFolder)
	End If
	
	sTimeStamp = Day(Date) & MonthName(Month(Date),True) & Year(Date) & "_" & Hour(Time) & Minute(Time) & Second(Time)
	sLogFile = sLogFolder + "\" + Environment.Value("TestName") &"_Console" + sTimeStamp + ".log"	
	Call objFSO.CreateTextFile(sLogFile)
	Set objFile = objFSO.OpenTextFile(sLogFile,8)
																														  
	objFile.WriteLine sConsoleMessage
	objFile.Close
	Set objFile = Nothing
	''**********************************************************************************
	''Kill Teamcenter if Flag is True
	''**********************************************************************************
	If Cbool(bTcKill) Then
		'Added by Vallari - Closing logged in Perspective, to get Tc default state
		Call Fn_MenuOperation("Select","File:Close")			
		Call Fn_TeamcenterExit()			
		'**********************************************************************************
		'Log Test Result
		''**********************************************************************************
		Call Fn_UpdateLogFiles("[" + Cstr(time) + "] - QTP [" + Environment.Value("ActionName") + "] - End", "")
		Call Fn_UpdateLogFiles("-----------------------------------------------------------------------------------------------", "")
		Call Fn_UpdateLogFiles("[" + Cstr(time) + "] - Final - Pass | Test Execution Result: PASS" & sPassStr, "PASS: All VP Pass" & sPassStr)
		Call Fn_UpdateLogFiles("-----------------------------------------------------------------------------------------------", "")
		'**********************************************************************************
		'Deleting Snapshot Image file, snapshot is not needed if ALL VP PASS
		'**********************************************************************************
		Dim filePath
		filePath = Environment.Value("BatchFldName") + "\" + Environment.Value("TestName") + ".png"
		if objFSO.FileExists(filePath) then
			objFSO.DeleteFile filePath, True
		End if						  
	End If		
	
	Set objFSO = Nothing
	'**********************************************************************************
	'Call for Code Coverage
	'**********************************************************************************
	bConsoleLog = 1
	Fn_Setup_TestcaseExitWithConsoleLog = True
	
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_Setup_BatFileOperations

'Description			 :	Function Used to perform operations in Bat files

'Parameters			   :   '1.StrAction: Action Name
'										 2.dicBatFileDetails : Bat file details

'Return Value		   : 	True Or False

'Examples				: 'dicBatFileDetails("BatFilePath")="C:\mainline\TestData\Test.bat"
									 'dicBatFileDetails("TC_ROOT")="C:\apps\siemens\TR"
									 'dicBatFileDetails("TC_DATA")="C:\apps\siemens\TD"
									 'dicBatFileDetails("cdTC_DATA")="True"
									 'dicBatFileDetails("Calltc_profilevarsBat")="True"
									 'dicBatFileDetails("Command")="create_project -u=infodba -p=infodba -g=dba -input=C:\util_log\projlist.lst > C:\util_share\create_project.log"
									 'Msgbox Fn_Setup_BatFileOperations("Create",dicBatFileDetails)

									'dicBatFileDetails("BatFilePath")="C:\mainline\TestData\Test.bat"
									'dicBatFileDetails("BatFilePastePath")="C:\mainline\Test12.bat"
									'Msgbox Fn_Setup_BatFileOperations("CopyPaste",dicBatFileDetails)

									'dicBatFileDetails("BatFilePath")="C:\mainline\TestData\MyTest.bat"
									'Msgbox Fn_Setup_BatFileOperations("Run",dicBatFileDetails)

									'dicBatFileDetails("BatFilePath")="C:\mainline\TestData\MyTest.bat"
									'dicBatFileDetails("PsExecPath")="C:\mainline\Utilities\BMIDE"
									'dicBatFileDetails("Command")="PsExec.exe -h \\pnv6s106 -u pnv6s106\autoadmin -p Password123 C:\util_share\Test.bat"
									'
									'Msgbox Fn_Setup_BatFileOperations("Create",dicBatFileDetails)
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												14-Feb-2011								1.0																						Sunny R
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sachin Joshi											21-Feb-2012								1.1																						
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Public Function Fn_Setup_BatFileOperations(StrAction,dicBatFileDetails)
		'Variable declaration
		const bytesToKb = 1024
		Dim objShell
		Dim objFSO, objFile
		Dim sDriveName
		Fn_Setup_BatFileOperations=False
		'Creating File System object
		Set objFSO = CreateObject("Scripting.FileSystemObject")		
		Select Case StrAction
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			'Case to create bat file
			Case "Create"
				'Delete bat file
				if objFSO.FileExists(dicBatFileDetails("BatFilePath")) then
					objFSO.DeleteFile(dicBatFileDetails("BatFilePath"))		
				end if
				'Create bat file
				Set objFile = objFSO.CreateTextFile(dicBatFileDetails("BatFilePath"), True)

				'Added by Vallari - Required for Win 7  & 64-bit XP m/c
				If dicBatFileDetails("PsExecPath")<>"" Then
					objFile.WriteLine "reg add ""hklm\system\currentcontrolset\control"" /f /v SCMApiConnectionParam /t REG_DWORD /d 1"
				End If

				If dicBatFileDetails("TC_ROOT")<>"" Then
					objFile.WriteLine "Set TC_ROOT=" & dicBatFileDetails("TC_ROOT")
				End If
				If dicBatFileDetails("TC_DATA")<>"" Then
					objFile.WriteLine "Set TC_DATA=" & dicBatFileDetails("TC_DATA")
				End If
				If dicBatFileDetails("cdTC_DATA")="True" Then
					objFile.WriteLine "cd %TC_DATA%"
				End If
				If dicBatFileDetails("Calltc_profilevarsBat")="True" Then
					objFile.WriteLine "Call %TC_DATA%\tc_profilevars.bat"
				End If
				If dicBatFileDetails("cdTC_ROOT")="True" Then
					objFile.WriteLine "cd %TC_ROOT%"
				End If
				
				If dicBatFileDetails("CPDCommand")<>""  Then
					objFile.WriteLine "CD\"
					objFile.WriteLine "Set TC_Command=" & dicBatFileDetails("CPDCommand")
				End If
				If dicBatFileDetails("cdCPDCommand")="True" Then
					objFile.WriteLine "cd %TC_Command%"
				End If
				
				If dicBatFileDetails("PsExecPath")<>"" Then
					sDriveName = objFSO.GetDriveName(dicBatFileDetails("PsExecPath"))
					objFile.WriteLine sDriveName
					objFile.WriteLine "Set Drive_Path=" & dicBatFileDetails("PsExecPath")
					objFile.WriteLine "cd %Drive_Path%"
				End If

				If instr(1,lcase(dicBatFileDetails("Command")),"cpd_populate_cd") Then
                    dicBatFileDetails("Command")=replace(dicBatFileDetails("Command"),"cpd_populate_cd","4gd_populate_cd")
					dicBatFileDetails("Command")=replace(dicBatFileDetails("Command"),"-partition_names=assy","")
					dicBatFileDetails("Command")=replace(dicBatFileDetails("Command"), " -dtition_names=assy","")
				    dicBatFileDetails("Command")=replace(dicBatFileDetails("Command"), " -design_element_names=assy","")

					'Removing Log file location
					Dim sDummyStr
					ReDim sDummyStr(2)
					sDummyStr = Split (dicBatFileDetails("Command"), ">")
                    sDummyStr(0) = sDummyStr(0)+"-use_original_names = true"
                    dicBatFileDetails("Command") = sDummyStr(0)
				End If

				If dicBatFileDetails("Command")<>"" Then
					If instr(dicBatFileDetails("Command"), "PsExec.exe") > 0 Then
						replace dicBatFileDetails("Command"), "PsExec.exe", "PsExec.exe/accepteula"
					End If
					objFile.WriteLine dicBatFileDetails("Command")
				End If

				objFile.Close
				Set objFile = Nothing
				Set objFile =  objFSO.GetFile(dicBatFileDetails("BatFilePath"))
				if cint(objFile.Size) > 0 then
					Fn_Setup_BatFileOperations = True	
				else
					Fn_Setup_BatFileOperations = False	
				end if
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			'Case to copy bat file from source and paste to destination
			Case "CopyPaste"
				'Delete bat file
				if objFSO.FileExists(dicBatFileDetails("BatFilePastePath")) then
					objFSO.DeleteFile(dicBatFileDetails("BatFilePastePath"))		
				end if
				objFSO.CopyFile dicBatFileDetails("BatFilePath"),dicBatFileDetails("BatFilePastePath"),True
				if objFSO.FileExists(dicBatFileDetails("BatFilePastePath")) then
					Fn_Setup_BatFileOperations = True
				Else
					Fn_Setup_BatFileOperations = False
				end if

			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			'Case to run bat file
			Case "Run"
				
				If objFSO.FileExists(dicBatFileDetails("BatFilePath")) then
					Set objShell = CreateObject("WScript.Shell")
					objShell.Run "%comspec% /c " & dicBatFileDetails("BatFilePath"), 2, True
					Set objShell = Nothing
					Fn_Setup_BatFileOperations = True
				Else
					Fn_Setup_BatFileOperations = False
				End If
				
		End Select
		'Releasing File System object
		Set objFSO =Nothing
End Function



'********************************************* **************************************************************

'Function Name		:					Fn_EnableTcExcelAddin

'Description			 :		 		  The function is used to Enable Teamcenter Addin for Excel

'Return Value		   : 				True / False

'Pre-requisite			:		 		None

'Examples				:		Call Fn_EnableTcExcelAddin()
'

'History:
'										Developer Name				Date				Rev. No.			Changes Done			Reviewer	
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Nilesh					01-June-2012		   		1.0							Created
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_EnableTcExcelAddin()
   On Error Resume Next
   Dim objExcel,iAddinCount,iCount,bFlag,Obj1,bResult,l,t,r,b
   Dim bAddinCheck
   Call Fn_Set_ExcelAddinRegistryVal() 'Set Registry value =3 of LoadBehavior reg key of TC Excel Addin

	bAddinCheck=False
	Set objExcel=CreateObject("Excel.Application")
	iAddinCount=objExcel.COMAddIns.Count
	For iCount=1 To iAddinCount
		If Instr(Lcase(objExcel.COMAddIns(iCount).ProgId),"tcexceladdin")>0 Then
            bAddinCheck=True
			bFlag=objExcel.COMAddIns(iCount).Connect
			objExcel.COMAddIns(iCount).Connect=True
			If Err.Number>0 Then
				Fn_EnableTcExcelAddin=True
				Exit Function
			End If
		End If
	Next

	If  bFlag=False And bAddinCheck=True Then
		objExcel.Visible=True
		wait 2
		'Open Blan Excel Sheet
		Call Fn_KeyBoardOperation("SendKeys","%~F~T~{ENTER}")		
		Wait 2
		'Open Excerl Option Window
		Call Fn_KeyBoardOperation("SendKeys","%~F~T")
		'Click on Add-ins 
    	Set Obj1=Window("MicrosoftExcel").Window("Excel Options").WinObject("NetUIHWND")
		bResult=Obj1.GetTextLocation("Add-Ins",l,t,r,b,True)
'		Print bResult
		If bResult=True Then
			Obj1.Click Cint((l+r)/2),Cint((t+b)/2)
			Wait 2
		Else
			Window("MicrosoftExcel").Window("Excel Options").Activate
			Call Fn_KeyBoardOperation("SendKeys","A~A")
			Wait 2
		End If
		'Select Disabled Items from Manage Dropdown
		Call Fn_KeyBoardOperation("SendKeys","{Tab}~{Tab}~{Tab}~{DOWN}~{UP}")
		wait 2
		Call Fn_KeyBoardOperation("SendKeys","{ENTER}")
		Wait 2
		'Click on Go... button
		Call Fn_KeyBoardOperation("SendKeys","{Tab}")
		wait 2
        Call Fn_KeyBoardOperation("SendKeys","{ENTER}")
		wait 1
		'Select Disabled Addin from Disabled Items window
		Call Fn_KeyBoardOperation("SendKeys"," ")
		wait 2
		'Click on Enable button
        Call Fn_KeyBoardOperation("SendKeys","%E")
		Wait 2
		'Close Disable Items Window
		Window("MicrosoftExcel").Window("Excel Options").Window("Disabled Items").Close
		Wait 1
		'Close Excel Option Window
		Window("MicrosoftExcel").Window("Excel Options").Close
		'Close Excel Application
		Window("MicrosoftExcel").Close
		Fn_EnableTcExcelAddin=True
	Else
		Fn_EnableTcExcelAddin=True
	End If

	iAddinCount=objExcel.COMAddIns.Count
	For iCount=1 To iAddinCount
		If Instr(Lcase(objExcel.COMAddIns(iCount).ProgId),"tcexceladdin")>0 Then
			bFlag=objExcel.COMAddIns(iCount).Connect
		End If
	Next
	If  bFlag=False And bAddinCheck=True Then
		Call Fn_EnableTcExcelAddin()
	End If
	Set Obj1=Nothing
	Set objExcel=Nothing
End Function

'********************************************* **************************************************************

'Function Name		:					Fn_Set_ExcelAddinRegistryVal

'Description			 :		 		  The function is used to  set Registry value =3 to LoadBehavior key of TC Excel Addin

'Return Value		   : 				Nothing

'Pre-requisite			:		 		None

'Examples				:		Call Fn_Set_ExcelAddinRegistryVal()
'

'History:
'										Developer Name				Date				Rev. No.			Changes Done			Reviewer	
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Nilesh					04-June-2012		   		1.0							Created
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function Fn_Set_ExcelAddinRegistryVal()
   On Error Resume Next
	Dim strComputer,oReg,strKeyPath,strValueName,dwValue

	Const HKEY_LOCAL_MACHINE = &H80000002
	strComputer = "."
	Set oReg=GetObject( "winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
	strKeyPath="SOFTWARE\Microsoft\Office\Excel\Addins\TcExcelAddin"
	strValueName="LoadBehavior"
	oReg.GetDWORDValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,dwValue
	If dwValue<>3 Then
		oReg.SetDWORDValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,3
	End If
	Set oReg=Nothing

End Function

'********************************************* **************************************************************

'Function Name		:					Fn_ExcelErrorClose

'Description			 :		 		  The function is used to  Close Error Dialog of Excel

'Return Value		   : 				Nothing

'Pre-requisite			:		 		None

'Examples				:		Call Fn_ExcelErrorClose()
'

'History:
'										Developer Name				Date				Rev. No.			Changes Done			Reviewer	
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Nilesh					08-June-2012		   		1.0							Created
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function Fn_ExcelErrorClose()
   On Error Resume Next
   If Dialog("ExceClose").Exist(1) Then
		Dialog("ExceClose").Type "N"
		If Dialog("ExceClose").Exist(3)Then
			Dialog("ExceClose").Close()
		End If
	End If
End Function

'********************************************* **************************************************************
'Function Name		:					Fn_SISW_GetHierarchy

'Description			 :		 		  The function is used to  to get hirearchy of frequently changing menus, tree items etc

'Return Value		   : 				Hierarchy for input item

'Pre-requisite			:		 		None

'Examples				:				dicGetHierarchy("Project ID") = ""	
'													Call Fn_SISW_GetHierarchy(dicGetHierarchy)

'History:
'										Developer Name				Date				Rev. No.			Changes Done			Reviewer	
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Prasanna					 12-July-2012		   		1.0							Created
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function Fn_SISW_GetHierarchy(dicGetHierarchy)

			Dim dictItems,dictKeys,sObjElement
	
			dictItems = dicGetHierarchy.Items
			dictKeys = dicGetHierarchy.Keys
			
			For i = 0 to dicGetHierarchy.Count - 1
						sObjElement = dictKeys(i)
						 Select Case sObjElement 			   				
										 Case "Project ID"	
																Fn_SISW_GetHierarchy = "Unassigned:Project IDs"
										 Case Else
																Fn_SISW_GetHierarchy = false
																Exit function
										 End Select
			Next

End Function
'********************************************* **************************************************************
'Function Name		:					Fn_SISW_Reload_Addin

'Description			 :		 		  The function is used to  to get reload specific addin from opened QTP test

'Return Value		   : 				Always True

'Pre-requisite			:		 		None

'Examples				:				Call Fn_SISW_Reload_Addin("Java")

'History:
'										Developer Name				Date				Rev. No.			Changes Done			Reviewer	
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Vallari							 12-Dec-2012		   		1.0							Created
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public function Fn_SISW_Reload_Addin(sAddinName)
   Dim objQTPApp, arrAddIns, arrModAddIns
   Dim i, j

   Set objQTPApp = CreateObject("QuickTest.Application")

   arrAddIns = objQTPApp.Test.GetAssociatedAddins

   ReDim arrModAddIns(Ubound(arrAddIns) - 1)

   j = 0

   For i = 0 to Ubound(arrAddIns)
	   If trim(lcase(arrAddIns(i))) <> trim(lcase(sAddinName)) Then
		   arrModAddIns(j) = arrAddIns(i)
		   j = j + 1
	   End If
   Next

   objQTPApp.Test.SetAssociatedAddins arrModAddIns
   wait(3)
	objQTPApp.Test.SetAssociatedAddins arrAddIns
	wait(2)

   Set objQTPApp = Nothing

   Fn_SISW_Reload_Addin = True

End Function

'*********************************************************		Function Captures Desktop Image		***********************************************************************
'Function Name		:				Fn_SISW_Setup_CaptureDesktopImg

'Description		:		 		 Function to capture desktop image

'Parameters			   :	 		sFolderName - FOlder will be created (as per Feature) with this name under Report folder
'									sObjRef - Object frame to be captured, else send "" as param value for capturing desktop
'									sFileNmae - name given for the saved image file
											
'Return Value		   : 				TRUE \ FALSE

'Pre-requisite			:		 		NA

'Examples				:				Fn_SISW_Setup_CaptureDesktopImg("TcViz_DesktopImg", "TestName_Assmy")

'History:
'										Developer Name			Date			Rev. No.			Changes Done			Reviewer		Reviewed Date	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Vallari					18-Apr-2013		1.0																	 											
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_Setup_CaptureDesktopImg(sFolderName, sObjRef, sFilenmae)
	Dim sImgPath
	Dim objFSO
	
	Fn_SISW_Setup_CaptureDesktopIm = False
	
	
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	if NOt objFSO.FolderExists(Environment.Value("BatchFldName") + "\" + sFolderName) Then
		objFSO.CreateFolder Environment.Value("BatchFldName") + "\" + sFolderName
	End If
	
	Set objFSO = Nothing
	
	sImgPath = Environment.Value("BatchFldName") + "\" + sFolderName + "\" + sFilenmae + ".bmp"
	Err.clear
	
	If sObjRef.toString() <> "" Then
		sObjRef.WaitProperty "enabled",True,10
		sObjRef.CaptureBitmap sImgPath, True		
	else
		Desktop.CaptureBitmap sImgPath,True				
	End If	
	
	If Err.Number < 0 Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Capture [" + sFilenmae + "] Iamge")
	Else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Successfully Captured [" + sFilenmae + "] Iamge")
		
		Fn_SISW_Setup_CaptureDesktopImg = sImgPath
	End If
		
End Function

'*********************************************************		Function to Verify Windows Process		***********************************************************************
'Function Name		:				Fn_SISW_Setup_VerifyWinProc

'Description		:		 		 Function to Verify Windows Process

'Parameters			   :	 		sProName - Name of the process as seen in Task Manager
											
'Return Value		   : 				TRUE \ FALSE

'Pre-requisite			:		 		NA

'Examples				:				Fn_SISW_Setup_VerifyWinProc("javaw.exe")

'History:
'										Developer Name			Date			Rev. No.			Changes Done			Reviewer		Reviewed Date	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Vallari					8-May-2013		1.0																														
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_Setup_VerifyWinProc(sProcName)
	Dim strComputer
	Dim objWMIService
	Dim colProcess
	
	strComputer = "." 
	Set objWMIService = GetObject("winmgmts:"& "{impersonationLevel=impersonate}!\\"& strComputer & "\root\cimv2") 

	Set colProcess = objWMIService.ExecQuery("Select * from Win32_Process Where Name ='" + sProcName + "'")  
	'For Each objProcess in colProcess 
	If colProcess.Count > 0 Then
		Fn_SISW_Setup_VerifyWinProc = True
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Windows Process [" + sProcName + "] Exists")
	Else
		Fn_SISW_Setup_VerifyWinProc = False
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Windows Process [" + sProcName + "] Doesn't Exist")		
	End If
	
End Function

'*********************************************************		Function to Verify Image		***********************************************************************
'Function Name		:				Fn_SISW_ImgComp_Operations

'Description		:		 		 Function to Compare Images

'Parameters			   :	 		ActualBmp & ExpectedBmp - file paths
										
'Return Value		   : 				TRUE \ FALSE

'Pre-requisite			:		 		Installation Of Node JS 

'Examples				:				Fn_SISW_ImgComp_Operations("C:\test.bmp", "C:\standard.bmp")

'History:
'										Developer Name			Date			 Rev. No.			Changes Done			Reviewer		Reviewed Date	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Akshay Bhagwat		3-Apr-2020	         1.0																														
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function Fn_SISW_ImgComp_Operations(ActualBmp,ExpectedBmp)
	Dim objXMLDoc,objNodeVal,objfso
	Set objfso = createobject("Scripting.FileSystemObject")
	If objfso.FileExists(ActualBmp) AND objfso.FileExists(ExpectedBmp) Then
		
			Set objXMLDoc = CreateObject("Microsoft.XMLDOM")
			objXMLDoc.async=false
			objXMLDoc.load( Environment.Value("sPath") +"\Utilities\imageCompare\image.xml")
	
			If (objXMLDoc.ParseError.ErrorCode = 0) Then
				Set objNodeVal = objXMLDoc.SelectSingleNode("/Environment/firstimagepath")			
				objNodeVal.Text=ActualBmp
				
				
				Set objNodeVal = objXMLDoc.SelectSingleNode("/Environment/secondimagepath")			
			    objNodeVal.Text=ExpectedBmp
			    
			    Set objNodeVal = objXMLDoc.SelectSingleNode("/Environment/result")			
			    objNodeVal.Text= "False"
			    objXMLDoc.Save(Environment.Value("sPath") +"\Utilities\imageCompare\image.xml")
			   
	    		Set objNodeVal = nothing 
				Set objXMLDoc = nothing
			    Set objShellApp = CreateObject("WScript.Shell")
					objShellApp.CurrentDirectory=Environment.Value("sPath") +"\Utilities\imageCompare"
					objShellApp.Run "run.bat"
				Set objShellApp = nothing
				
				wait(2)
				Set objXMLDoc = CreateObject("Microsoft.XMLDOM")
				objXMLDoc.async=false
				objXMLDoc.load(Environment.Value("sPath") +"\Utilities\imageCompare\image.xml")
				 
				Set objNodeVal = objXMLDoc.SelectSingleNode("/Environment/result")			
				
			    Fn_SISW_ImgComp_Operations = objNodeVal.Text
			  End  If
	End If
	Set objfso = Nothing
End  Function
'*********************************************************		Function to Verify picture files		***********************************************************************
'Function Name		:				Fn_SISW_Setup_CompareBitmap

'Description		:		 		 Function to Verify Windows Process

'Parameters			   :	 		ActualBmp & ExpectedBmp - file paths
										
'Return Value		   : 				TRUE \ FALSE

'Pre-requisite			:		 		NA

'Examples				:				Fn_SISW_Setup_CompareBitmap("C:\test.bmp", "C:\standard.bmp")

'History:
'										Developer Name			Date			Rev. No.			Changes Done			Reviewer		Reviewed Date	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Vallari					9-May-2013		1.0																														
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function Fn_SISW_Setup_CompareBitmap(ActualBmp, ExpectedBmp) 

	Dim fCompare 
	Fn_SISW_Setup_CompareBitmap = False
  
	Set fCompare = CreateObject("Mercury.FileCompare")   
	  
	If fCompare.IsEqualBin(ExpectedBmp, ActualBmp, 0, 1) Then  
	   Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Bitmap COmparison for files [" + ActualBmp + "] and [" + ExpectedBmp + "] is Successful")		
	   Fn_SISW_Setup_CompareBitmap=True  
	Else  
	   Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Bitmap COmparison for files [" + ActualBmp + "] and [" + ExpectedBmp + "] Failed")		
	   Fn_SISW_Setup_CompareBitmap=False     
	End If  
  
End Function 

'*********************************************************		Function to Load Library at runtime if not loaded		***********************************************************************
'Function Name		:				Fn_SISW_LoadLibrary(sFilePath)

'Description			 :		 		 Loads shared Object Repository for specified Action

'Parameters			   :	 			1. sFilePath	-  Path of the Library 

'Return Value		   : 			True/False

'Pre-requisite			:		 	Nothing

'Examples				:			Call Fn_SISW_LoadLibrary("C:\mainline\Library\Preference.vbs")

'History					 :		
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'	--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sushma		 													   9-Jul-13			              1.0										Created
'	---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_LoadLibrary(sFilePath)

		Dim objQTPApp, objLib
		Set objQTPApp = CreateObject("QuickTest.Application")
        Set objLib = objQTPApp.Test.Settings.Resources.Libraries

        If objLib.Find(sFilePath) = -1 Then ' If library is not already added
				LoadFunctionLibrary(sFilePath )
				If Err.Number < 0  Then
					Services.LogMessage "Failed to load GUI file ["  &sFilePath &"]", Err.Description 
					Fn_SISW_LoadLibrary = FALSE
				Else
					Fn_SISW_LoadLibrary = TRUE
				End If				
		Else
			Fn_SISW_LoadLibrary = TRUE
		End If
		Set objLib = Nothing  ' Release Lib Object
		Set objQTPApp = Nothing ' Release QTP Object

End Function
'-----------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_SISW_Setup_ArrayStringContains

'Description		:	This function is used to check whether specified value is present in given Array String.

'Parameters			:	1. sString - Valid Action
'						2. sValue - Panel Type
'						3. sSeparator - Dictionary object
'
'RETURNVALUE     	:   True/False
'
'PRE-REQUISITES  	:  None
'
'Examples			:	Call Fn_SISW_Setup_ArrayStringContains("A~B~C", "B", "~")
'
'-------------------------------+---------------+---------------+-------------------------------+-----------------------------------
'	Developer Name				|	Date		|	Rev. No.	|	Reviewer					|	Changes Done	
'-------------------------------+---------------+---------------+-------------------------------+-----------------------------------
'	Koustubh Watwe			 	|26-Dec-2012	|	1.0			|	Koustubh Watwe		 		| 	Created
'-------------------------------+---------------+---------------+-------------------------------+-----------------------------------
Public Function Fn_SISW_Setup_ArrayStringContains(sString, sValue, sSeparator)
	Dim iCnt, arrValues
	Fn_SISW_Setup_ArrayStringContains = False
	arrValues = split(sString, sSeparator)
	For iCnt = 0 to UBound(arrValues)
		If arrValues(iCnt) = sValue Then
			Fn_SISW_Setup_ArrayStringContains = True
			Exit for
		End If
	Next
End Function
'-----------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_SISW_Setup_GetObjectFromXML

'Description		:	Function to get specified Object hierarchy.

'Parameters			:	1. sXMLPath : XMl File Name
'						2. sObjectName : Object Name

'Return Value		: 	Object \ Nothing

'Pre-requisite		:	Teamcenter application should be displayed

'Examples			:	Set obj = Fn_SISW_Setup_GetObjectFromXML("c:\object.xml", "DefaultWindow")

'History			:
'-------------------------------+---------------+---------------+-------------------------------+-----------------------------------
'	Developer Name				|	Date		|	Rev. No.	|	Reviewer					|	Changes Done	
'-------------------------------+---------------+---------------+-------------------------------+-----------------------------------
'	Koustubh Watwe		 		|21-Sept-2013	|	1.0			|	Koustubh Watwe		 		| Copied Fn_SISW_Setup_GetObject from TcPMM Mainline and modified
'-------------------------------+---------------+---------------+-------------------------------+-----------------------------------
Function Fn_SISW_Setup_GetObjectFromXML(sXMLPath, sObjectName)
	Dim sFuncLog, bResult, aFilePath
	Dim objXMLDoc	
	Dim intNodeLength
	Dim intNodeCount	
	Dim objNodeName
	Dim objNodeVal
	bResult = ""
	aFilePath = split(sXMLPath, "\")
	sFuncLog = "Fn_SISW_Setup_GetObjectFromXML > " & aFilePath(uBound(aFilePath)) & " > "
	Set Fn_SISW_Setup_GetObjectFromXML = Nothing
	set objXMLDoc=CreateObject("Microsoft.XMLDOM")												' Create XMLDOM object
	objXMLDoc.async="false"
	objXMLDoc.load(sXMLPath)																	' Loading QTP Environment XML

	If (objXMLDoc.parseError.errorCode <> 0) Then
		Exit function
	Else
		intNodeLength = objXMLDoc.getElementsByTagName("Variable").length
		For intNodeCount = 0 to (intNodeLength - 1)
			Set objNodeName = objXMLDoc.SelectSingleNode("/Environment/Variable[" & intNodeCount &"]/Name")			
			Set objNodeVal = objXMLDoc.SelectSingleNode("/Environment/Variable[" & intNodeCount &"]/Value")
			If LCase(objNodeName.Text) = LCase(sObjectName) Then				
				bResult = objNodeVal.Text
				Exit For
			End IF
			Set objNodeVal = Nothing
			Set objNodeName = Nothing
		Next
		Set objNodeVal = nothing 
		Set objNodeName = nothing
		Set objXMLDoc = nothing
	End if	

	If bResult <> "" AND bResult <> False Then 
		Set Fn_SISW_Setup_GetObjectFromXML = eval(bResult)
		'Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & "PASS : Successfully returned object hierarchy of [" & Fn_SISW_Setup_GetObjectFromXML.toString() & "]")
	Else
		' Failed to Find Toolbar button
		'Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & "FAIL : Object hierarchy of [" & sObjectName & "] does not exist.")
	End IF
End Function


'*********************************************************		Function to Load Library at runtime if not loaded		***********************************************************************
'Function Name		:			Fn_SISW_MakeIEDefaultBrowser(sFilePath)

'Description			:		 	To make IEas Default Browser

'Parameters			:	 		NA

'Return Value		   	: 			True/False

'Pre-requisite			:		 	Nothing

'Examples			:			Call Fn_SISW_MakeIEDefaultBrowser

'History				:			Developer Name				Date						Rev. No.				Changes Done						Reviewer
'	--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Created By 			:			Snehal Salunkhe		 			25-Jul-2014			1.0					Added
'	---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_MakeIEDefaultBrowser()
   On Error Resume Next
	Dim objShell,temp
	Set objShell = CreateObject("WScript.Shell")
	
	objShell.RegDelete("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.htm\UserChoice\")
	objShell.RegDelete("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.html\UserChoice\")
	
	temp=objShell.RegWrite("HKEY_CURRENT_USER\Software\Microsoft\Windows\Shell\Associations\UrlAssociations\https\UserChoice\Progid","IE.HTTPS","REG_SZ")
	temp=objShell.RegWrite("HKEY_CURRENT_USER\Software\Microsoft\Windows\Shell\Associations\UrlAssociations\http\UserChoice\Progid","IE.HTTP","REG_SZ")
	temp=objShell.RegWrite("HKEY_CURRENT_USER\Software\Microsoft\Windows\Shell\Associations\UrlAssociations\ftp\UserChoice\Progid","IE.FTP","REG_SZ")
	
	
	temp=objShell.RegWrite("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.htm\UserChoice\Progid","IE.AssocFile.HTM","REG_SZ")
	temp=objShell.RegWrite("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.html\UserChoice\Progid","IE.AssocFile.HTM","REG_SZ")
	
	temp=objShell.RegWrite("HKEY_CURRENT_USER\Software\Clients\StartmenuInternet\","iexplore.exe","REG_SZ")
	
	If Err.Number<0 Then
		Fn_SISW_MakeIEDefaultBrowser=False
	Else
		Fn_SISW_MakeIEDefaultBrowser=True
	End If
End Function

'*********************************************************		Function to invoke the application	************************************************************************

'Function Name		:		Fn_InvokeTeamCenterExt()

'Description			:		 This function invokes the team center application with prespective

'Parameters			  :	 			

'Return Value		   : 		The String which represents the result : "PASS" or "FAIL" with the reason

'Pre-requisite			:		  The Team Centre Application 2007 should be installed

'Examples				:		  Fn_InvokeTeamCenterExt() will invoke team center application

'History					 :		
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'												------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Reema W														22-Oct-2014
'												------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function Fn_InvokeTeamCenterExt(sModule)
   On Error resume next
	Dim sPath, sModuleName, sAutoDir

	sPath = Environment.Value("AppExecutable")

	If sModule <> "" Then
		sAutoDir = Fn_GetEnvValue("User", "AutomationDir")
		sModuleName = Fn_GetXMLNodeValue(sAutoDir + "\TestData\TcModules.xml", sModule)
		SystemUtil.Run sPath, "-perspective " + sModuleName
	Else
		SystemUtil.Run sPath
	End If

	If JavaWindow("Teamcenter Login").Exist(iTimeOut) Then						  							
			Fn_InvokeTeamCenterExt = TRUE  																						
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Invoked Teamcenter Application from [" + sPath + "]")
	Else
			 Fn_InvokeTeamCenterExt = FALSE
			 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Invoke Teamcenter Application from [" + sPath + "]")
			 Exit Function
	End If

End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - @Function Header Start 
'Function Name			 :	Fn_Setup_GetActivePerspectiveName
'
'Function Description	 :	Function used to get teamcenter active perspective name
'
'Function Parameters	 :  1.sAction : Action to perform						
'
'Function Return Value	 : 	Perspective name
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	Teamcenter application should be displayed
'
'Function Usage		     :	Call Fn_Setup_GetActivePerspectiveName("")
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	      Reviewer		|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Vrushali Sahare			    |  26-Feb-2016	    |	 1.0		|	  Ganesh Bhosale	| 		Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Public Function Fn_Setup_GetActivePerspectiveName(sAction)
	GBL_FAILED_FUNCTION_NAME="Fn_Setup_GetActivePerspectiveName"
	'Declaring Variables
	Dim objDefaultWindow
	Dim sPerspectiveName
	
	Fn_Setup_GetActivePerspectiveName=""
	
	'Creating object of Teamcenter Default Window
	Set objDefaultWindow = JavaWindow("title:=.* - Teamcenter .*","tagname:=.* - Teamcenter .*","resizable:=1","index:=0")
	Set objDefaultWindow = JavaWindow("DefaultWindow")
	Select Case Lcase(sAction)
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "getnameext"
			sPerspectiveName=Split(objDefaultWindow.GetROProperty("title"),"-")
			sPerspectiveName(0)=Trim(sPerspectiveName(0))
			Fn_Setup_GetActivePerspectiveName=Replace(sPerspectiveName(0)," ","")
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -	
		Case "getname",""
			sPerspectiveName=Split(objDefaultWindow.GetROProperty("title"),"-")
			Fn_Setup_GetActivePerspectiveName=Trim(sPerspectiveName(0))			
	End Select	
End Function
'*********************************************************		Function to Calculate time ************************************************************************

'Function Name		:	Fn_SecondsToMinutes()

'Description		:	This function Calculates Hrs , Minutes & seconds from passed seconds parameter 

'Parameters			:	Pass intSeconds as total seconds		

'Return Value		: 	The String which represents the result : Minutes & second or Hrs & minutes & seconds

'Examples			:	Fn_SecondsToMinutes(504)

'History			:		
'						Developer Name							Date				Rev. No.	    Changes Done       Reviewer
'						------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'						Poonam C								26-Dec-2016			  1.0
'						------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Function Fn_SecondsToMinutes(ByVal intSeconds)
	Dim hours, minutes, seconds
	
	' calculates whole hours (like a div operator)
	hours = intSeconds \ 3600
	
	' calculates the remaining number of seconds
	intSeconds = intSeconds Mod 3600
	
	' calculates the whole number of minutes in the remaining number of seconds
	minutes = intSeconds \ 60
	
	' calculates the remaining number of seconds after taking the number of minutes
	seconds = intSeconds Mod 60
	
	If hours <> 0 Then
		' returns as a string
		 Fn_SecondsToMinutes = hours & " hrs, " & minutes & " mins, " & seconds & " seconds"
	Else
		' returns as a string
		 Fn_SecondsToMinutes = minutes & " mins, " & seconds & " seconds"
	End If
	
End Function

'*********************************************************		Function to Calculate time ************************************************************************

'Function Name		:	Fn_SISW_ConvertToItemID()

'Description		:	This function to convert integer item id to string 

'Parameters			:	Pass intItemNumber item id in integer format

'Return Value		: 	6 digit item id in String format

'Examples			:	Fn_SISW_ConvertToItemID(504)

'History			:		
'Developer Name							Date				Rev. No.	    Changes Done       Reviewer
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Koustubh Watwe						31-Aug-2017			  1.0
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_ConvertToItemID(intItemNumber )
	Dim iStartNumber, sInitZeros, iLength, iCount
	iLength = 6
	sInitZeros = "0"
	
	If Len(Cstr(intItemNumber)) < iLength Then
		For iCount = Len(Cstr(intItemNumber)) + 1 To iLength-1
			sInitZeros = sInitZeros & "0"
		Next
		Fn_SISW_ConvertToItemID = sInitZeros & Cstr(intItemNumber)
	 Else
		Fn_SISW_ConvertToItemID = Cstr(intItemNumber)
	 End If
End Function
'********************************************* **************************************************************
'Function Name		:					Fn_map_drive

'Description			 :		 		This function is used to disconnect and connect mapdrive 

'Return Value		   : 				True or False

'Pre-requisite			:		 		None

'Examples				:				
'													
'History:
'										Developer Name				Date				Rev. No.			Changes Done			Reviewer	
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Rohit Sangwai				 08-April-2020		   		1.0							Created
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Function Fn_map_drive(strDrive,StrName,user,password)
Dim sharepath,retVal
err.clear
retVal=False
sharepath="\\"&StrName&"\util_share"

Set WshNetwork = CreateObject("WScript.Network")

WshNetwork.RemoveNetworkDrive strDrive
on error resume next

WshNetwork.MapNetworkDrive strDrive, sharepath, true,user,password

If err.number<0 Then
	retVal=False
else
	retVal=True
End If
Fn_map_drive=retVal
End Function

'********************************************* **************************************************************
'Function Name		:					Fn_create_utilrun_file

'Description			 :		 		This function is used to create a file which trigger batchfile 

'Return Value		   : 				True or False

'Pre-requisite			:		 		None

'Examples				:					
'													
'History:
'										Developer Name				Date				Rev. No.			Changes Done			Reviewer	
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Rohit Sangwai				 08-April-2020		   		1.0							Created
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function Fn_create_utilrun_file(filename,dir,user,pass,cmdfilename,servername) 
Set objFSO = CreateObject("Scripting.FileSystemObject")
        
        batchName = "\"&filename&".ps1"	
        currDir=dir  
        		
		Set objFile = objFSO.CreateTextFile(currDir & batchName , True) 
		objFile.WriteLine "$userPassword = ConvertTo-SecureString -String "&pass&" -AsPlainText -Force"
		objFile.WriteLine "$userCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList plm\"&user&", $userPassword"
		objFile.WriteLine "$sb={Invoke-Expression -Command:"&CHR(34)&"cmd.exe /c 'C:\util_share\"&cmdfilename&".bat'"&CHR(34)& "}"
		'objFile.WriteLine "$server_name="&servername
		objFile.WriteLine "Invoke-command -computer "&servername& " -Credential $userCredential -ScriptBlock $sb"
		
		objFile.WriteLine cmdParam
		objFile.Close
		
		
		If objFSO.FileExists(currDir+batchName) then
    Fn_create_utilrun_file=True
    
  Else

    Fn_create_utilrun_file=False

End If

Set objFile = Nothing
Set objFSO = Nothing
		
        
End Function


