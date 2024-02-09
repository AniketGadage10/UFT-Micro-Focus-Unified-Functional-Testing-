Option Explicit
'1. Fn_SISW_BB_GetObject(sObjectName)
'2. Fn_SISW_BB_Setup_InvokeBriefcaseBrowser()
'3. Fn_SISW_BB_Setup_ReadyStatusSync(iIterations)
'4. Fn_SISW_BB_Setup_Exit(sAction,sPath)
'5. Fn_SISW_BB_Setup_LoadBBXML()
'6. Fn_SISW_BB_Setup_TestcaseExit(bTcKill)
'7. Fn_SISW_BB_Setup_CreateReportFolder(sBBReportFolderPath).
'8. Fn_SISW_BB_Setup_ClearCache(sPath)
'9. Fn_SISW_BB_Setup_KillProcess(sProcessToKill)

'****************************************    Function to get Object hierarchy ***************************************
'Function Name		 	:	Fn_SISW_BB_GetObject
'
'Description		    :  	Function to get Object hierarchy
'
'Parameters		    	:	1. sObjectName : Object Handle name
'
'Return Value		    :  	Object \ Nothing
'
'Examples		     	:	Fn_SISW_BB_GetObject("BBWindow")
'
'History:
'	Developer Name			Date			Rev. No.		Reviewer		Changes Done	
'----------------------------------------------------------------------------------------------------------------------------------
'	Shweta Rathod		 19-Jul-2016		1.0				Koustubh Watwe
'----------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_BB_GetObject(sObjectName)
	Dim sObjectXMLPath
	sObjectXMLPath = Fn_GetEnvValue("User", "AutomationDir") & "\TestData\AutomationXML\ObjectXML\BriefcaseBrowser.xml"
	Set Fn_SISW_BB_GetObject = Fn_SISW_Setup_GetObjectFromXML(sObjectXMLPath, sObjectName)
End Function


'****************************************    Function to get Object hierarchy ***************************************
'Function Name		 	:	Fn_SISW_BB_Setup_InvokeBriefcaseBrowser
'
'Description		    :  	Function to Invoke BriefcaseBrowser application
'
'Return Value		    :  	TRUE \ FALSE
'
'Examples		     	:	Fn_SISW_BB_GetObject("BBWindow")
'
'History:
'	Developer Name			Date			Rev. No.		Reviewer		Changes Done	
'----------------------------------------------------------------------------------------------------------------------------------
'	Shweta Rathod		 19-Jul-2016		1.0				Koustubh Watwe
'----------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_BB_Setup_InvokeBriefcaseBrowser()																																		 
	SystemUtil.Run Environment.Value("BriefcaseBrowserPath"),"",Replace(Environment.Value("BriefcaseBrowserPath"),"BriefcaseBrowser.exe","") ,""
	If Fn_SISW_UI_Object_Operations("Fn_SISW_BB_Setup_InvokeBriefcaseBrowser","Exist",JavaWindow("BriefcaseBrowser"),SISW_MAX_TIMEOUT) Then						  							
		Fn_SISW_BB_Setup_InvokeBriefcaseBrowser = TRUE  																						
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Invoked Briefcase Browser Application from [" + Environment.Value("BriefcaseBrowserPath") + "]")
	Else
		 Fn_SISW_BB_Setup_InvokeBriefcaseBrowser = FALSE
		 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Invoke Briefcase Browser Application from [" + Environment.Value("BriefcaseBrowserPath") + "]")
		 Exit Function
	End If
End Function

'*********************************************************		Function to Synchronize on Application Response	***********************************************************************
'Function Name		:				Fn_SISW_BB_Setup_ReadyStatusSync(iIterations)

'Description			 :		 		 This function waits till Application comes to Ready state

'Parameters			   :	 			1. iIterations: No. of times to be checked for Ready text
											
'Return Value		   : 				TRUE \ FALSE

'Pre-requisite			:		 		Briefcase browser application should be displayed

'Examples				:				 Fn_SISW_BB_Setup_ReadyStatusSync(2)

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Shweta Rathod			25-Jul-2016			1.0										Koustubh Watwe
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_BB_Setup_ReadyStatusSync(iIterations)
	Dim iCounter, bFound, iCnt, objQSearchEdit
	Fn_SISW_BB_Setup_ReadyStatusSync =  false
	JavaWindow("BriefcaseBrowser").JavaStaticText("ReadyStatus").SetTOProperty "label","Open \(\Finished at.*"
	For iCounter = 1 to iIterations
		If JavaWindow("BriefcaseBrowser").Exist(SISW_DEFAULT_TIMEOUT) Then
			JavaWindow("BriefcaseBrowser").JavaStaticText("ReadyStatus").WaitProperty "label", "ReadyStatus", 20000
			If JavaWindow("BriefcaseBrowser").JavaStaticText("ReadyStatus").Exist(1) Then
				exit for
			End If
		Else
			Fn_SISW_BB_Setup_ReadyStatusSync = FALSE
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Briefcase Browser window does not exist.")	
			exit function
		End If
	Next
	If JavaWindow("BriefcaseBrowser").JavaStaticText("ReadyStatus").Exist(1) Then
		For iCounter = 1 to iIterations
			For iCnt = 1 to 20
				wait 1
				If JavaWindow("BriefcaseBrowser").Exist(SISW_DEFAULT_TIMEOUT) Then
					' exit from inner loop if progressbar disappears
					If JavaWindow("BriefcaseBrowser").JavaObject("ProgressBar").exist(1) = FALSE Then
						Fn_SISW_BB_Setup_ReadyStatusSync = TRUE
						Exit for
					End If
				Else
					Fn_SISW_BB_Setup_ReadyStatusSync = FALSE
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: BriefcaseBrowser window does not exist.")	
					exit function
				End If
			Next
			' exit from main loop if progressbar disappears
			If Fn_SISW_BB_Setup_ReadyStatusSync Then Exit for
		Next
	end if
	If JavaWindow("BriefcaseBrowser").JavaStaticText("ReadyStatus").Exist(1) = FALSE OR Fn_SISW_BB_Setup_ReadyStatusSync = FALSE Then
		Fn_SISW_BB_Setup_ReadyStatusSync = FALSE
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: BriefcaseBrowser Not Ready after [" + CStr(iIterations) + "] sync iterations")		
	Else
		Fn_SISW_BB_Setup_ReadyStatusSync = TRUE
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: BriefcaseBrowser is Ready in [" + CStr(iIterations) + "] sync iterations")		
	End If
End Function


'*********************************************************		Function to Synchronize on Application Response	***********************************************************************
'Function Name		:				Fn_SISW_BB_Setup_Exit(sAction,sPath)

'Description			 :		 		exit form briefcase browser application

'Parameters			   :	 			1. sAction: There are two ways to exit from application accordingly need to set this value 
'										2. sPath : path to save the briefcase browser
'
'Return Value		   : 				Path of tree \ FALSE

'Pre-requisite			:		 		Briefcase browser should be open and all other dialoues inside application should close if open

'Examples				:				 Fn_SISW_BB_Setup_Exit("withsave","C:\Temp\BBAssembly\TestCaseName_IrandNum","")
'										Fn_SISW_BB_Setup_Exit("withoutsave","","")

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Shweta Rathod			25-Jul-2016			1.0											Koustubh Watwe
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function Fn_SISW_BB_Setup_Exit(sAction,sPath)
	Dim objJavaWindowExit
	Fn_SISW_BB_Setup_Exit = false
	
	Set objJavaWindowExit =Fn_SISW_BB_GetObject("SaveResource")
	bGblFuncRetVal = Fn_SISW_UI_Object_Operations("Fn_SISW_BB_Setup_Exit","Exist",objJavaWindowExit,SISW_MICRO_TIMEOUT)
	if bGblFuncRetVal = false then
		call Fn_UI_JavaMenu_Select("Fn_SISW_BB_Setup_Exit",JavaWindow("BriefcaseBrowser"), Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("BriefcaseBrowser_Menu"), "BBExit"))
	End if 
	
	Select Case sAction
	 	Case "withoutsave"
	 			bGblFuncRetVal = Fn_SISW_UI_Object_Operations("Fn_SISW_BB_Setup_Exit","Exist",objJavaWindowExit,SISW_MICRO_TIMEOUT)
				If bGblFuncRetVal = true Then call Fn_SISW_UI_JavaButton_Operations("Fn_SISW_BB_Setup_Exit", "Click", objJavaWindowExit,"No")	
	 	Case "withsave"
	 End Select
	
	Set objJavaWindowExit = nothing
	 Fn_SISW_BB_Setup_Exit = true
End Function



'*************************************************  Function to perform various opeartions on the prefrence operation window ***********************************************************************
'Function Name		:					Fn_SISW_BB_Setup_LoadBBXML()

'Description			 :		 		this will load all the enviornment level variable which are given in configuration file BBEnvVar.xml

'
'Return Value		   : 				nothing

'Pre-requisite			:		 		File "BBEnvVar.xml" should be present with all the valid values

'Examples				:				Fn_SISW_BB_Setup_LoadBBXML()

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Shweta Rathod			01-Aug-2016			1.0											Koustubh watwe
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_BB_Setup_LoadBBXML()	
	Environment.Value("BBCustomDatasetMappings") = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("BriefcaseBrowser_Envvar"), "BBCustomDatasetMappings")
	Environment.Value("BriefcaseBrowserPath") = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("BriefcaseBrowser_Envvar"), "BriefcaseBrowserPath")
	Environment.Value("BBVersion") = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("BriefcaseBrowser_Envvar"), "BBVersion")
	Environment.Value("NXVersion") = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("BriefcaseBrowser_Envvar"), "NXVersion")
	
	'creating folder to save NX assembly
	call Fn_UpdateEnvXMLNode(Fn_LogUtil_GetXMLPath("BriefcaseBrowser_Envvar"),"BBNX_AssemblyPath",Environment.Value("BBReportFolderPath")+"\NX")
	Environment.Value("BBNX_AssemblyPath") = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("BriefcaseBrowser_Envvar"), "BBNX_AssemblyPath")
	bGblFuncRetVal = fn_SISW_util_folder_operation("exist",Environment.Value("BBNX_AssemblyPath"))
	If bGblFuncRetVal = false then
		bGblFuncRetVal = fn_SISW_util_folder_operation("createfolder",Environment.Value("BBNX_AssemblyPath"))
		If bGblFuncRetVal = false then 
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: falied to create the folder [" + Environment.Value("BBNX_AssemblyPath") + "] ")		
			Fn_SISW_BB_Setup_LoadBBXML = false 
			Exit function
		End if
	End if
	
	'updating XML path to create bczfile 
	call Fn_UpdateEnvXMLNode(Fn_LogUtil_GetXMLPath("BriefcaseBrowser_Envvar"),"BBAssemblyPath",Environment.Value("BBReportFolderPath")+"\bczfile")
	Environment.Value("BBAssemblyPath") = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("BriefcaseBrowser_Envvar"), "BBAssemblyPath")
	bGblFuncRetVal = fn_SISW_util_folder_operation("exist",Environment.Value("BBAssemblyPath"))
	If bGblFuncRetVal = false then
		bGblFuncRetVal = fn_SISW_util_folder_operation("createfolder",Environment.Value("BBAssemblyPath"))
		If bGblFuncRetVal = false then 
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: falied to create the folder [" + Environment.Value("BBAssemblyPath") + "] ")		
			Fn_SISW_BB_Setup_LoadBBXML = false 
			Exit function
		End if
	End if
	Environment.Value("Site1") = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("BriefcaseBrowser_Envvar"), "Site1")
	Environment.Value("BBConfigXMLPath") = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("BriefcaseBrowser_Envvar"), "BBConfigXMLPath")
End Function

'*********************************************  Function to exit the testcase **************************************************************

'Function Name		:					Fn_SISW_BB_Setup_TestcaseExit

'Description			 :		 		  The function handles test script end part

'Parameters			   :	 			

'Return Value		   : 				True/False

'Pre-requisite			:		 		None

'Examples				:
'

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										shweta Rathod			05-Oct-2016	   		1.0											Koustubh Watwe
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function Fn_SISW_BB_Setup_TestcaseExit(bTcKill)
	Dim bReturn,filePath,bCaptureImgVP
	if bTcKill <> false then
		'**********************************************************************************
		'exit without save
		''**********************************************************************************
		bGblFuncRetVal = Fn_SISW_BB_Setup_Exit("withoutsave","")
		If bGblFuncRetVal = false then 
			Call Fn_KillProcess("BriefcaseBrowser.exe")
			exit function		
		End if
		'**********************************************************************************
		'kill Briefcase browser process
		''**********************************************************************************
		Call Fn_KillProcess("BriefcaseBrowser.exe")
	end if
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
	
	'**********************************************************************************
	'Deleting Snapshot Image file, snapshot is not needed if ALL VP PASS
	'**********************************************************************************
	filePath = Environment.Value("BatchFldName") + "\" + Environment.Value("TestName") + ".png"
	bGblFuncRetVal = fn_splm_util_file_operation("exist",filePath)
	if bGblFuncRetVal = true then
		call fn_splm_util_file_operation("delete",filePath)
	End if
	
	'**********************************************************************************
	'Deleting NX,bcz file, extracted files are not needed if ALL VP PASS
	'**********************************************************************************
	call Fn_SISW_BB_Setup_ClearCache(Environment.Value("BBReportFolderPath"))
	
	bGblFuncRetVal = fn_SISW_util_folder_operation("exist",Environment.Value("BBReportFolderPath"))
	if bGblFuncRetVal = true then
		call fn_splm_util_file_operation("delete",Environment.Value("BBReportFolderPath"))
	End if
End Function


'*********************************************************		Function to Creates ReportFolder ***********************************************************************
'Function Name		:				Fn_SISW_BB_Setup_CreateReportFolder(sReportFolderName)

'Description			 :		 		 Creates report folder under batch location

'Parameters			   :	 			1. sLogFileName : Name with which batch file needs to be creatde

'Return Value		   : 			sReportFolderPath : Returns report folder path

'Pre-requisite			:		 	Nothing

'Examples				:			Call Fn_SISW_BB_Setup_CreateReportFolder(Environment.Value("TestName"))

'History					 :		
'	Developer Name												Date						Rev. No.						Changes Done						Reviewer
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Shweta Rathod		 										26/08/2016			            1.0								Created						Koustubh Watwe
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_BB_Setup_CreateReportFolder(sBBReportFolderPath)
	Dim bret
	Fn_SISW_BB_Setup_CreateReportFolder = ""
	' create test case specific folder
'	sPath = sBBReportFolderPath
		
	'checking existence of report folder
	bret = fn_SISW_util_folder_operation("exist",sBBReportFolderPath)
	If bret = true then
		'deleting report folder if exist
		'bret = fn_SISW_util_folder_operation("deletefolder",sBBReportFolderPath)
		Call Fn_SISW_BB_Setup_ClearCache(sBBReportFolderPath)
		wait 1
		If bret = false Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to delete briefcase browser report folder [ "+sBBReportFolderPath+"]")
			Exit Function
		else
			'creating report folder 
			bret = fn_SISW_util_folder_operation("createfolder",sBBReportFolderPath)
			If bret = false Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to create briefcase browser report folder [ "+sBBReportFolderPath+"]")
				Exit Function
			End if
		End if
	else
		'creating report folder if not - exist
		bret = fn_SISW_util_folder_operation("createfolder",sBBReportFolderPath)
		If bret = false Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to create briefcase browser report folder [ "+sBBReportFolderPath+"]")
			Exit Function
		End if
	End if
	Fn_SISW_BB_Setup_CreateReportFolder = sBBReportFolderPath
End Function


'*********************************************************		Function to Creates ReportFolder ***********************************************************************
'Function Name		:				Fn_SISW_BB_Setup_ClearCache(sReportFolderName)
'
'Description		:		 	deleted report folder,NX,bczfile and extractedfiles from the report folder
'
'Return Value		 : 			 Returns nothing
'
'Pre-requisite		 :		 	Nothing
'
'Examples			 :			Call Fn_SISW_BB_Setup_ClearCache(sPath)
'
'History					 :		
'	Developer Name												Date						Rev. No.						Changes Done								Reviewer
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Shweta Rathod		 										13/09/2016			            1.0								Created						
'   Shweta Rathod												13/10/2016														modified code to make it generalise		shweta	
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function Fn_SISW_BB_Setup_ClearCache(sPath)
	Dim objFSO,fp1,objSubFolder,sPath2,fp2

	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set fp1 = objFSO.GetFolder(sPath+"\")

	'geting subfolder name of report folder 
	Set objSubFolder = fp1.SubFolders
	For Each Subfolder in fp1.SubFolders
		sPath2 = objSubFolder.Item(Subfolder.Name)
		
		Set fp2 = objFSO.GetFolder(sPath2+"\")
	
		On Error Resume Next
	    For Each objFile In fp2.Files
	        objFile.Delete True 'setting force to true deletes read-only file                        
	    Next
    
	    'deleting subfolders of report folder 
		For Each Subfolder2 in fp2.SubFolders
			objSubFolder2.Item(Subfolder2.Name).Delete
			wait 1
		next   
	
		'delete f2
		If objFSO.FolderExists(sPath2) Then
			objFSO.DeleteFolder(sPath2)
			wait 1
		End if
	next

	On Error Resume Next
	For Each objFile In fp1.Files
	    objFile.Delete True 'setting force to true deletes read-only file                        
	Next

	'delete fp1
	If objFSO.FolderExists(sPath) Then
		objFSO.DeleteFolder(sPath)
		wait 1
	End if
	
	Set objFSO = Nothing
	Set fp1 = Nothing
	Set objSubFolder = Nothing
	Set fp2 = Nothing

End Function

'*********************************************************		Function to ProcessToKill ***********************************************************************
'Function Name		:				Fn_SISW_BB_Setup_KillProcess(sProcessToKill)
'
'Description		:		 	kill the process  "EXCEL.EXE","WINWORD.EXE","ugraf.exe","Teamcenter.exe"
'
'Return Value		 : 			 Returns nothing
'
'Pre-requisite		 :		 	Nothing
'
'Examples			 :			Fn_SISW_BB_Setup_KillProcess("EXCEL.EXE:WINWORD.EXE")
'
'History					 :		
'	Developer Name												Date						Rev. No.						Changes Done								Reviewer
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Shweta Rathod		 										19/10/2016			            1.0								Created						
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function Fn_SISW_BB_Setup_KillProcess(sProcessToKill)
	Dim sArrData,strComputer,objWMIService,iCount
	sArrData = split(sProcessToKill, ":",-1,1)
	strComputer = "."
	For iCount = 0 to ubound(sArrData)
		Select Case sArrData(iCount)
			Case "EXCEL.EXE","WINWORD.EXE","ugraf.exe","BriefcaseBrowser.exe"
				Set objWMIService = GetObject("winmgmts:"& "{impersonationLevel=impersonate}!\\"& strComputer & "\root\cimv2") 
				Set colProcess = objWMIService.ExecQuery("Select * from Win32_Process Where Name ='"+sArrData(iCount)+"'")  
				'For Each objProcess in colProcess 
				For Each objProcess in colProcess 
						objProcess.Terminate() 
				Next 
			Case "Teamcenter.exe"
				Call Fn_KillProcess("")
		End Select		
	Next
	Set objWMIService = nothing
End function
