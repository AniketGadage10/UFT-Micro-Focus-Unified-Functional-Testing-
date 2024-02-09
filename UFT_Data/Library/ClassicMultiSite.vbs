Option Explicit
'*********************************************************	Function List *******************************************
'1. Fn_SISW_CMS_GetObject
'2. Fn_SISW_CMS_SiteOperations
'3. Fn_CommandOperations
'4. Fn_SISW_UpdateEnvXMLNode
'5. Fn_DeleteCacheFiles
'6. Fn_CMS_RemoteExport_Operation
'7. Fn_CMS_RemoteImport_Operation
'8. Fn_CMS_RemoteSiteSelection_Ops
'9. Fn_CMS_ImportExportSettings_Operation
'10. Fn_CMS_OptionsSettings_Ops
'11. Fn_CMS_ExportObjects_Operation
'12. Fn_CMS_ImportObjects_Operation
'13. Fn_CMS_Folder_Operation
'14. Fn_CMS_RemoteImportProgress_Operations
'15. Fn_CMS_ImportRemoteDialog_Operations
'16. Fn_CMS_MultiSiteSyncSettings_Operation
'17. Fn_CMS_MultiSiteSynchronisation_Operation
'18. Fn_CMS_Delete_UserLevel_Preference
'*********************************************************	Function List		**************************************

'=====================================================================================================================
' Function Name	:	Fn_SISW_CMS_GetObject
'
' Description	:  	Function to get specified Object hierarchy.
'
''Parameters	:	1. sObjectName : Object Handle name
								
' Return Value 	:  	Object \ Nothing
'
' Examples	  	:	 Fn_SISW_MyTc_GetObject("View/Edit Multi Unit Configuration")
'
' History	 	: 	Developer Name			Date			Rev. No.		Reviewer								Changes Done	
'-----------------------------------------------------------------------------------------------------------------------------------
'					Poonam Chopade		08-May-2018		  Created		Tc11.5_20180402.00_NewDevelopment		
'-----------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_CMS_GetObject(sObjectName)
	Dim sObjectXMLPath
	sObjectXMLPath = Fn_GetEnvValue("User", "AutomationDir") & "\TestData\AutomationXML\ObjectXML\ClassicMultiSite.xml"
	Set Fn_SISW_CMS_GetObject = Fn_SISW_Setup_GetObjectFromXML(sObjectXMLPath, sObjectName)
End Function
'=====================================================================================================================
'#	Function Name		:	Fn_SISW_CMS_SiteOperations
'#
'#	Description			:	Function is used to set Site details 
'#
'#	Parameters			:	sAction : Action Name 
'#						:	sSiteDetails : Server Details
'#						:   sReserve : 	Reserve variable
'#
'#	Return Value		: 	True/False 
'#
'#	Examples			:	Call Fn_SISW_CMS_SiteOperations("Set","pnv6s1458|C:\apps\siemens\multisite\pnv6s1458\portal\portal.bat|http://pnv6s1458/tc1120715/webclient|20180402.00","")
'#
'#	History				:	Developer Name	   		Date			 Rev. No.	    Changes Done       Reviewer
'#	
'#=====================================================================================================================
'#						:  Poonam Chopade		 02-May-2018			1.0			Created			TC11.5_2018040200
'#=====================================================================================================================
Public Function Fn_SISW_CMS_SiteOperations(sAction,sSiteDetails,sReserve)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_CMS_SiteOperations"
	Dim arrSiteDetails,bFlag,sQTPEnvXML,stccsPath
	
	Fn_SISW_CMS_SiteOperations = False
	Select Case sAction
		Case "Set"
			If sSiteDetails <> "" Then
				arrSiteDetails = Split(sSiteDetails,"|")
	
				'Update Env file
				sQTPEnvXML = Environment.Value("sPath") & "\TestData\EnvVar_Ext.xml"
				
				'Set TcServer
				bFlag = Fn_SISW_UpdateEnvXMLNode(sQTPEnvXML, "TcServer", arrSiteDetails(0))
				If bFlag = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to set Environment variable [ TcServer as " & arrSiteDetails(0) & "]")
					Fn_SISW_CMS_SiteOperations = False
					Exit Function
				End If
				
				'Set AppExecutable
				bFlag = Fn_SISW_UpdateEnvXMLNode(sQTPEnvXML, "AppExecutable", arrSiteDetails(1))
				If bFlag = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to set Environment variable [ AppExecutable as " & arrSiteDetails(1) & "]")
					Fn_SISW_CMS_SiteOperations = False
					Exit Function
				End If
				
				'Set TcWebServer
				bFlag = Fn_SISW_UpdateEnvXMLNode(sQTPEnvXML, "TcWebServer", arrSiteDetails(2))
				If bFlag = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to set Environment variable [ TcWebServer as " & arrSiteDetails(2) & "]")
					Fn_SISW_CMS_SiteOperations = False
					Exit Function
				End If
				
				'Set TcBuild
				bFlag = Fn_SISW_UpdateEnvXMLNode(sQTPEnvXML, "TcBuild", arrSiteDetails(3))
				If bFlag = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to set Environment variable [ TcBuild as " & arrSiteDetails(3) & "]")
					Fn_SISW_CMS_SiteOperations = False
					Exit Function
				End If
				
				'Clear Cache
				Call Fn_WindowsApplications("TerminateAll","java.exe")
                Call Fn_WindowsApplications("TerminateAll","javaw.exe") 
				Call Fn_DeleteCacheFiles
				
				'Generate tccs path
				stccsPath = Replace(Replace(Replace(arrSiteDetails(1),"\rac",""),"\portal.bat",""),"\portal","")
				stccsPath = stccsPath & "\tccs" 
				
				'Set FMS Home
				Call Fn_CMS_CommandOperations(stccsPath)
				Wait 1
				
				'Load Env Variables in QTP
				bFlag = LoadEnvXML()
				If bFlag = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to load Environment Variable in QTP")
					Fn_SISW_CMS_SiteOperations = False
					Exit Function
				End If	
				
				'Set SOA web client 
				sSrvURL = Fn_GetXMLNodeValue(sQTPEnvXML, "TcWebServer")
				If Instr(sSrvURL, "webclient") Then
					sSrvURL = Left(sSrvURL, (Len(sSrvURL) - Len("/webclient")))
				End If	
				
			End If
	End Select
	
	Fn_SISW_CMS_SiteOperations = True
End Function
'--------------------------------------------------------------------------------------------------------------------                                                                             
' Function Name     	: Fn_CommandOperations
' Function Usage    	: Create & Run command to Set FMS_Home
'--------------------------------------------------------------------------------------------------------------------
Function Fn_CMS_CommandOperations(sCommand)
	const bytesToKb = 1024
	Dim objShell
	Dim objFSO, objFile
	Dim sDriveName
	Dim sTR_ROOT, sTR_DATA, sTempDir
	Dim sCMDFileName, sLogFileName
	Dim arrCommand,iRanNo
    
	Set objFSO = CreateObject("Scripting.FileSystemObject") 
	
	arrCommand = Split(sCommand, "-", -1,1)
	sTR_ROOT = sCommand
	iRanNo = Fn_Setup_RandNoGenerate(4)
	
	Const WindowsFolder = 0
	Const SystemFolder = 1
	Const TemporaryFolder = 2
		
	sTempDir = objFSO.GetSpecialFolder(TemporaryFolder)
	sCMDFileName = sTempDir & "\FMS_Set_util"&iRanNo&".cmd"
	sLogFileName = sTempDir & "\" & Trim(arrCommand(0)) & ".log"
	
	'Delete log file
	if objFSO.FileExists(sLogFileName) then
	    objFSO.DeleteFile(sLogFileName)                
	end if
	
	'Delete cmd file
	if objFSO.FileExists(sCMDFileName) then
	    objFSO.DeleteFile(sCMDFileName)                
	end if
	
	'Create CMD file
	Set objFile = objFSO.CreateTextFile(sCMDFileName, True)
	sDriveName = objFSO.GetDriveName(sTR_ROOT)
	objFile.WriteLine sDriveName
	objFile.WriteLine "Set FMS_HOME=" & sTR_ROOT
	objFile.WriteLine "%FMS_HOME%\bin\fccstat.exe -restart"
	objFile.WriteLine sCommand & " > " & sLogFileName
	objFile.Close
	Set objFile = Nothing
	
	' Creating Shell object to run cmd file'
	Set objShell = CreateObject("WScript.Shell")
	objShell.Run "%comspec% /c " & sCMDFileName, 2, True
	Set objShell = Nothing
	
	Set objFile = Nothing
	Set objFSO = Nothing
End Function
'--------------------------------------------------------------------------------------------------------------------                                                                             
' Function Name     	: Fn_SISW_UpdateEnvXMLNode
' Function Description  : Update QTP Environment XML with the BatchResult folder path
' Function Usage    	: Result = Fn_UpdateEnvXMLNode(XMLDataFile, sNodeName, sNodeValue)
'							XMLDataFile	- Location of QTP Environment XML on test machine
'							sNodeName	- Node Name in XML (e.g. BatchFldName in QTP Environment XML)
'							sNodeValue	- Node Value for the sNodeName (e.g. BatchResult folder path)
'                     		return True on success, False on failure
'--------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_UpdateEnvXMLNode(XMLDataFile, sNodeName, sNodeValue)
	Dim objXMLDoc
	Dim objChildNodes
	Dim objSelectNode
	Dim intNodeLength
	Dim intNodeCount
	Dim intChildNodeCount
	Dim strNodeSting
	
	set objXMLDoc=CreateObject("Microsoft.XMLDOM")												' Create XMLDOM object
	objXMLDoc.async="false"
	objXMLDoc.load(XMLDataFile)																	' Loading QTP Environment XML
	
	If (objXMLDoc.parseError.errorCode <> 0) Then
		Fn_UpdateEnvXMLNode = False
	Else
		intNodeLength = objXMLDoc.getElementsByTagName("Variable").length
		For intNodeCount = 0 to (intNodeLength - 1)
			Set objChildNodes = objXMLDoc.documentElement.childNodes.item(intNodeCount).childNodes
				strNodeSting = ""
				For intChildNodeCount = 0 to (objChildNodes.length - 1)
						strNodeSting = strNodeSting & objChildNodes(intChildNodeCount).text 
				Next
				If Instr(strNodeSting, sNodeName) Then
					If sNodeName = "TcWebServer" Then
						Set objSelectNode = objXMLDoc.SelectSingleNode("/Environment/Variable[" & intNodeCount-1 &"]/Value")
					Else
						Set objSelectNode = objXMLDoc.SelectSingleNode("/Environment/Variable[" & intNodeCount &"]/Value")
					End If
					objSelectNode.Text = sNodeValue
					Exit For
				End If
		Next
		objXMLDoc.Save(XMLDataFile)
		Set objSelectNode = nothing 
		Set objChildNodes = nothing
		Set objXMLDoc = nothing
		Fn_SISW_UpdateEnvXMLNode = True
	End if	
End Function
'--------------------------------------------------------------------------------------------------------------------       
'Function Name	:	Fn_ClearCache
'Description	:	Function to clear cache
'--------------------------------------------------------------------------------------------------------------------  
Function Fn_DeleteCacheFiles()
		On error resume next
		Dim objNetwork, objFSO, objFolder, objSubFolder, objFiles
		Dim sDrive, sUserName, sPath,sFoldName,sArrData,fileIdx,iCnt,files 
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
			objSubFolder.Item("Teamcenter").Delete(True)
		End If
		
		''Deleting FCCCache folders
		sFoldName = "FCCCache" 
		For Each iCnt in objSubFolder
			sArrData = split(iCnt.Name, "_",-1,1)
			if sFoldName = sArrData(0) then
				If objFSO.FolderExists(sPath & "\"&iCnt.Name) Then
					objSubFolder.Item(iCnt.Name).Delete
					Exit For
				End If		
			End if 
		Next
		
		' Deleting Siemens folders
		If objFSO.FolderExists(sPath & "\Siemens") Then
			objSubFolder.Item("Siemens").Delete 
		End If
		
		' Deleting .Administrator folders
		sFoldName = ".Administrator"
		For Each iCnt in objSubFolder
			sArrData = split(iCnt.Name, "_",-1,1)
			if sFoldName = sArrData(0) then
				If objFSO.FolderExists(sPath & "\"&iCnt.Name) Then
					objSubFolder.Item(iCnt.Name).Delete
					Exit For
				End If	
			End if 
		Next
		
		objFSO.DeleteFile (objFolder+"\"+ "fcc.*"),DeleteReadOnly 'Deleting Fcc files		
		
		Const WindowsFolder = 0
		Const SystemFolder = 1
		Const TemporaryFolder = 2
		
		Set objFolder = objFSO.GetSpecialFolder(TemporaryFolder)  'Clear Temp folder
		Set objSubFolder = objFolder.SubFolders
		For Each iCnt in objSubFolder
			Err.Clear
			objFSO.DeleteFolder objFolder +"\"+ iCnt.Name,true
		Next
		
		Set files = objFolder.Files
		For each fileIdx In Files    
			objFSO.DeleteFile fileIdx,true
		Next

		'Clears out objects
		Set objNetwork = nothing
		Set objFSO = nothing
		Set objFolder = nothing
		Set objSubFolder = nothing
		Set files = nothing
		
		If Err.Number <> 0 Then
			Fn_ClearCache =False
			Call Fn_WriteLogFile("Fn_ClearCache()", 1, Err.Number ,"FAIL : Clear Cache Operation Failed")
		Else
			Fn_ClearCache =True
			Call Fn_WriteLogFile("Fn_ClearCache()", 3, Err.Number ,"PASS : Clear Cache Operation Passed")
		End If
End Function
'=====================================================================================================================================================================================
'@@
'@@    Function Name			:	Fn_CMS_RemoteExport_Operation
'@@
'@@    Description				:	Function used to perform Remote Export on Object
'@@
'@@    Parameters			   	:	1. sAction		: ActionName
'@@								:	2. dicDetails	: Information for to remove
'@@								:   3. sButton 		: Button Name										 	         
'@@
'@@    Return Value		   	   	: 	True Or False
'@@
'@@   Example 					:	Set dicREOptions = CreateObject("Scripting.Dictionary")
'@@										dicREOptions("SelectTab1") = "General"
'@@										dicREOptions("ItemOptions") = "Include all revisions:ON"
'@@										dicREOptions("DatasetOptions") = "Include all versions:ON~Include all files:ON"
'@@										dicREOptions("StructureManagerOptions") = "Include entire BOM:ON"
'@@										dicREOptions("SelectTab2") = "Advanced"
'@@										
'@@									Set dicREOptionSettings = CreateObject("Scripting.Dictionary")
'@@										dicREOptionSettings("Transfer Options") = "None"
'@@										dicREOptionSettings("Item Options") = "Include all revisions"
'@@										dicREOptionSettings("General Options") = "None"
'@@										dicREOptionSettings("Structure Manager Options") = "Include entire BOM"
'@@									
'@@									Set dicREDetails = CreateObject("Scripting.Dictionary")
'@@										dicREDetails("Reason") = "Test1"
'@@										dicREDetails("TargetSites") = "SiteB--1886227743"
'@@									bReturn = Fn_CMS_RemoteExport_Operation("RemoteExport",dicREDetails,dicREOptions,"Yes",dicREOptionSettings)
'@@
'@@    History					:	
'@@ ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@	  Developer Name			Date			 Rev. No.	   				Changes Done								 			Reviewer
'@@ ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@	  Poonam Chopade	 	30-Apr-2018	 		  1.0			Created - Added for MultiSite new TC's development			TC11.5(20180402.00)_CMS_NewDevelopment_PoonamC_30Apr2018
'=====================================================================================================================================================================================
Public Function Fn_CMS_RemoteExport_Operation(sAction,dicREDetails,dicREOptions,sContinueButton,dicREOptionSettings)
	GBL_FAILED_FUNCTION_NAME="Fn_CMS_RemoteExport_Operation"
	Dim ObjRemoteExport,sMenu,bReturn

	Fn_CMS_RemoteExport_Operation = False
	Set ObjRemoteExport = Fn_SISW_CMS_GetObject("RemoteImportExport")
	
	'Check Dialog existence
	If Fn_UI_ObjectExist("Fn_CMS_RemoteExport_Operation",ObjRemoteExport)  = False  Then
		sMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("MyTc_Menu"),"ToolsExportRemote")
		Call Fn_MenuOperation("Select",sMenu)
		Call Fn_ReadyStatusSync(2)
		
		'Check Dialog existence
		If Fn_UI_ObjectExist("Fn_CMS_RemoteExport_Operation",ObjRemoteExport)  = False  Then
			Set ObjRemoteExport = Nothing
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_CMS_RemoteExport_Operation ] Remote Export dialog dose not exists ].")
			Exit Function
		End If
		
	End If
	
	Select Case sAction
		Case "RemoteExport"
			'Enter Reason
			If dicREDetails("Reason") <> "" Then 
				 bReturn = Fn_SISW_UI_JavaEdit_Operations("Fn_CMS_RemoteExport_Operation", "Type", ObjRemoteExport, "Reason", dicREDetails("Reason"))	
				 If bReturn = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_CMS_RemoteExport_Operation ]  Failed to enter Reason as [ "+ dicREDetails("Reason")  +" ]") 
					Set ObjRemoteExport = Nothing
					Fn_CMS_RemoteExport_Operation = False
					Exit Function
				 End If
			 End If	
			 'Select Target Sites
			 If dicREDetails("TargetSites") <> "" Then 
			 	 'Click on Select target remote site icon
			 	 Call Fn_Button_Click("Fn_CMS_RemoteExport_Operation",ObjRemoteExport,"site_16")
			 	 Wait 1
			 	 bReturn = Fn_CMS_RemoteSiteSelection_Ops("Add",dicREDetails("TargetSites"),"OK")
				 If bReturn = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_CMS_RemoteExport_Operation ]  Failed to select Target Sites as [ "+ dicREDetails("TargetSites")+" ]") 
					Set ObjRemoteExport = Nothing
					Fn_CMS_RemoteExport_Operation = False
					Exit Function
				 End If
			 End If	
			 'Display or Set Export options
			 If vartype(dicREOptions) = 9 Then 
			 	'Click on button Display/set remote export option
			 	Call Fn_Button_Click("Fn_CMS_RemoteExport_Operation",ObjRemoteExport,"exportsettings_16")
			 	Wait 1
				bReturn = Fn_CMS_ImportExportSettings_Operation("SetImportExportOptions",dicREOptions,"OK")
				If bReturn = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_CMS_RemoteExport_Operation ]  Failed to set Remote Export options") 
					Set ObjRemoteExport = Nothing
					Fn_CMS_RemoteExport_Operation = False
					Exit Function
				 End If
			 End If
			 'Click Yes/No
			 If sContinueButton <> "" Then 
			 	bReturn = Fn_Button_Click("Fn_CMS_RemoteExport_Operation",ObjRemoteExport,sContinueButton)
			 	If bReturn = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_CMS_RemoteExport_Operation ]  Failed to click on [ "&sContinueButton&" ] button") 
					Set ObjRemoteExport = Nothing
					Fn_CMS_RemoteExport_Operation = False
					Exit Function
				 End If
			 End If
			 'Check Display or Set Export options in Option Settings  
			 If vartype(dicREOptionSettings) = 9 Then 
			 	bReturn = Fn_CMS_OptionsSettings_Ops("VerifyOptionsSettings",dicREOptionSettings,"Yes")
			 	If bReturn = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_CMS_RemoteExport_Operation ]  Failed to verify set Remote Export options") 
					Set ObjRemoteExport = Nothing
					Fn_CMS_RemoteExport_Operation = False
					Exit Function
				 End If
			 End If
	Case "VerifyExportOptionsChecked" 
			If vartype(dicREOptions) = 9 Then 
			 	'Click on button Display/set remote Import option
			 	Call Fn_Button_Click("Fn_CMS_RemoteExport_Operation",ObjRemoteExport,"exportsettings_16")
			 	Wait 1
				bReturn = Fn_CMS_ImportExportSettings_Operation("IsChecked",dicREOptions,"Cancel")
				If bReturn = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_CMS_RemoteExport_Operation ]  Failed to verify Remote Export options") 
					Set ObjRemoteExport = Nothing
					Fn_CMS_RemoteExport_Operation = False
					Exit Function
				 End If
			 End If
	End Select
	
	Fn_CMS_RemoteExport_Operation = True
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_CMS_RemoteExport_Operation ] executed successfully.") 
	Set ObjRemoteExport = nothing
End Function
'=====================================================================================================================================================================================
'@@
'@@    Function Name			:	Fn_CMS_RemoteImport_Operation
'@@
'@@    Description				:	Function used to perform Remote Export on Object
'@@
'@@    Parameters			   	:	1. sAction		: ActionName
'@@								:	2. dicDetails	: Information for to remove
'@@								:   3. sButton 		: Button Name										 	         
'@@
'@@    Return Value		   	   	: 	True Or False
'@@
'@@    Example					:	Set dicIEOptions = CreateObject("Scripting.Dictionary")
'@@										dicIEOptions("SelectTab1") = "General"
'@@										dicIEOptions("ItemOptions") = "Include all revisions:ON"
'@@										dicIEOptions("DatasetOptions") = "Include all versions:ON~Include all files:ON"
'@@										dicIEOptions("StructureManagerOptions") = "Include entire BOM:ON"
'@@										dicIEOptions("SelectTab2") = "Advanced"
'@@										
'@@									Set dicIEOptionSettings = CreateObject("Scripting.Dictionary")
'@@										dicIEOptionSettings("Transfer Options") = "None"
'@@										dicIEOptionSettings("Item Options") = "Include all revisions"
'@@										dicIEOptionSettings("General Options") = "None"
'@@										dicIEOptionSettings("Structure Manager Options") = "Include entire BOM"
'@@									
'@@									Set dicIEDetails = CreateObject("Scripting.Dictionary")
'@@										dicIEDetails("Reason") = "Test1"
'@@									bReturn = Fn_CMS_RemoteImport_Operation("RemoteImport",dicIEDetails,dicIEOptions,"Yes",dicIEOptionSettings)
'@@
'@@    History					:	
'@@ ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@	  Developer Name			Date			 Rev. No.	   				Changes Done								 			Reviewer
'@@ ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@	  Poonam Chopade	 	30-Apr-2018	 		  1.0			Created - Added for MultiSite new TC's development			TC11.5(20180402.00)_CMS_NewDevelopment_PoonamC_30Apr2018
'=====================================================================================================================================================================================
Public Function Fn_CMS_RemoteImport_Operation(sAction,dicIEDetails,dicIEOptions,sContinueButton,dicIEOptionSettings)
	GBL_FAILED_FUNCTION_NAME="Fn_CMS_RemoteImport_Operation"
	Dim ObjRemoteImport,sMenu,ObjRemoteImport1
	Dim bReturn
	Fn_CMS_RemoteImport_Operation = False
	Set ObjRemoteImport = Fn_SISW_CMS_GetObject("RemoteImportExport")
	Set ObjRemoteImport1 = Fn_SISW_CMS_GetObject("RemoteImportExport@1")
	
	'Check Dialog existence
	If Fn_UI_ObjectExist("Fn_CMS_RemoteImport_Operation",ObjRemoteImport)  = False and Fn_UI_ObjectExist("Fn_CMS_RemoteImport_Operation",ObjRemoteImport1) = False Then
		sMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("RAC_Menu"),"ToolsImportRemote")
		Call Fn_MenuOperation("Select",sMenu)
		Call Fn_ReadyStatusSync(2)
		
		If Fn_UI_ObjectExist("Fn_CMS_RemoteImport_Operation",ObjRemoteImport)  = False  Then
			Set ObjRemoteImport = Fn_SISW_CMS_GetObject("RemoteImportExport@1")
		End If
		
		If Fn_UI_ObjectExist("Fn_CMS_RemoteImport_Operation",ObjRemoteImport)  = False  Then
			Set ObjRemoteImport = Nothing
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_CMS_RemoteImport_Operation ] Remote Import dialog dose not exists ].")
			Exit Function
		End If
	Else
		'Set path
		If Fn_UI_ObjectExist("Fn_CMS_RemoteImport_Operation",ObjRemoteImport) = False  Then
			Set ObjRemoteImport = Fn_SISW_CMS_GetObject("RemoteImportExport@1")
		End If
	End If
	
	Set ObjRemoteImport1 = Nothing
	
	Select Case sAction
		Case "RemoteImport"
			'Enter Reason
			  If vartype(dicIEDetails) = 9 Then 
				If dicIEDetails("Reason") <> "" Then 
					 bReturn = Fn_SISW_UI_JavaEdit_Operations("Fn_CMS_RemoteImport_Operation", "Type", ObjRemoteImport, "Reason", dicIEDetails("Reason"))	
					 Call Fn_ReadyStatusSync(1)	
					 If bReturn = False Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_CMS_RemoteImport_Operation ]  Failed to enter Reason as [ "+ dicIEDetails("Reason")  +" ]") 
						Set ObjRemoteImport = Nothing
						Fn_CMS_RemoteImport_Operation = False
						Exit Function
					 End If
				 End If	
			 End If 
			 'Display or Set Import options
			 If vartype(dicIEOptions) = 9 Then 
			 	'Click on button Display/set remote Import option
			 	Call Fn_Button_Click("Fn_CMS_RemoteImport_Operation",ObjRemoteImport,"importremote_16")
			 	Wait 1
				bReturn = Fn_CMS_ImportExportSettings_Operation("SetImportExportOptions",dicIEOptions,"OK")
				If bReturn = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_CMS_RemoteImport_Operation ]  Failed to set Remote Export options") 
					Set ObjRemoteImport = Nothing
					Fn_CMS_RemoteImport_Operation = False
					Exit Function
				 End If
			 End If
			 'Click Yes/No
			 If sContinueButton <> "" Then 
			 	bReturn = Fn_Button_Click("Fn_CMS_RemoteImport_Operation",ObjRemoteImport,sContinueButton)
			 	If bReturn = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_CMS_RemoteImport_Operation ]  Failed to click on [ "&sContinueButton&" ] button") 
					Set ObjRemoteImport = Nothing
					Fn_CMS_RemoteImport_Operation = False
					Exit Function
				 End If
				 wait 2
			 End If
			 'Check Display or Set Export options in Option Settings  
			 If vartype(dicIEOptionsettings) = 9 Then 
			 	bReturn = Fn_CMS_OptionsSettings_Ops("VerifyOptionsSettings",dicIEOptionsettings,"Yes")
			 	If bReturn = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_CMS_RemoteImport_Operation ]  Failed to verify set Remote Export options") 
					Set ObjRemoteImport = Nothing
					Fn_CMS_RemoteImport_Operation = False
					Exit Function
				 End If
			 End If 
		Case "VerifyImportOptionsChecked","VerifyValueExists" 
			If vartype(dicIEOptions) = 9 Then 
			 	'Click on button Display/set remote Import option
			 	Call Fn_Button_Click("Fn_CMS_RemoteImport_Operation",ObjRemoteImport,"importremote_16")
			 	Wait 1
				bReturn = Fn_CMS_ImportExportSettings_Operation("IsChecked",dicIEOptions,"Cancel")
				If bReturn = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_CMS_RemoteImport_Operation ]  Failed to verify Remote Import options") 
					Set ObjRemoteImport = Nothing
					Fn_CMS_RemoteImport_Operation = False
					Exit Function
				 End If
			 End If
	End Select
	
	Fn_CMS_RemoteImport_Operation = True
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_CMS_RemoteImport_Operation ] executed successfully.") 
	Set ObjRemoteImport = nothing
End Function
'=====================================================================================================================================================================================
'@@    Function Name			:	Fn_CMS_RemoteSiteSelection_Ops
'@@
'@@    Description				:	Function used to perform operations on Remote Site Selection dialog
'@@
'@@    Parameters			   	:	1. sAction		: ActionName
'@@								:	2. sSites		: Sites names (Seperated with ~)
'@@								:   3. sButtons 	: Button Names										 	         
'@@
'@@    Return Value		   	   	: 	True Or False
'@@
'@@    Examples					:	Call Fn_CMS_RemoteSiteSelection_Ops("Add","Site2","OK")
'@@
'@@    History					:	
'@@ ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@	  Developer Name			Date			 Rev. No.	   				Changes Done								 			Reviewer
'@@ ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@	  Poonam Chopade	 	30-Apr-2018	 		  1.0			Created - Added for MultiSite new TC's development			TC11.5(20180402.00)_CMS_NewDevelopment_PoonamC_30Apr2018
'=====================================================================================================================================================================================
Public Function Fn_CMS_RemoteSiteSelection_Ops(sAction,sSites,sButtons)
	GBL_FAILED_FUNCTION_NAME="Fn_CMS_RemoteSiteSelection_Ops"
	Dim objRSSelectionDialog,iCounter, bReturn, aSites, aButtons  
	
	Set objRSSelectionDialog = Fn_SISW_CMS_GetObject("RemoteSiteSelection")
	Set objRSSelectionDialog = Fn_UI_ObjectCreate("Fn_CMS_RemoteSiteSelection_Ops",objRSSelectionDialog)
	Fn_CMS_RemoteSiteSelection_Ops = False
	
	Select Case sAction 
			Case "Add"	'Add Target Site					
				If sSites<>"" Then
						aSites = split(sSites, "~",-1,1)
						For iCounter = 0 to UBound(aSites)
							bReturn = Fn_SISW_UI_JavaList_Operations("Fn_CMS_RemoteSiteSelection_Ops", "Select", objRSSelectionDialog, "AvailableSites", aSites(iCounter), "", "")
							If bReturn = True Then
								Call Fn_Button_Click("Fn_CMS_RemoteSiteSelection_Ops", objRSSelectionDialog, "Add")
							End If
						Next
				End If
			Case "Remove" 'Remove Target Site	
				If sSites<>"" Then
						aSites = split(sSites, "~",-1,1)
						For iCounter = 0 to UBound(aSites)
							bReturn = Fn_SISW_UI_JavaList_Operations("Fn_CMS_RemoteSiteSelection_Ops", "Select", objRSSelectionDialog, "SelectedSites", aSites(iCounter), "", "")
							If bReturn = True Then
								Call Fn_Button_Click("Fn_CMS_RemoteSiteSelection_Ops", objRSSelectionDialog, "Remove")
							End If
						Next
				End If				
	End Select
	
	'Click on Buttons
	If sButtons<>"" Then
			aButtons = split(sButtons, ":",-1,1)
			For iCounter=0 to Ubound(aButtons)
				Call Fn_Button_Click("Fn_CMS_RemoteSiteSelection_Ops", objRSSelectionDialog, aButtons(iCounter))
                Call Fn_ReadyStatusSync(2)
			Next
	End If
		
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" Successfully completed of function Fn_CMS_RemoteSiteSelection_Ops")
	Fn_CMS_RemoteSiteSelection_Ops = TRUE
    Set objRSSelectionDialog = nothing 	
	
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name	:	Fn_CMS_ImportExportSettings_Operation
'@@
'@@    Description		:	Function Used to perform operations on "Remote Import/Export Options" dialog
'@@
'@@    Parameters		:	1. sAction		: Action to be performed
'@@						:	2. dicDetails	: Dictionary object
'@@						:	3. sButton		: OK / Cancel button
'@@
'@@    Return Value		: 	True Or False
'@@
'@@    Examples			:	Set dicIEOptions = CreateObject("Scripting.Dictionary")
'@@								dicIEOptions("SelectTab1") = "General"
'@@								dicIEOptions("ItemOptions") = "Include all revisions:ON"
'@@								dicIEOptions("DatasetOptions") = "Include all versions:ON~Include all files:ON"
'@@								dicIEOptions("StructureManagerOptions") = "Include entire BOM:ON"
'@@								dicIEOptions("SelectTab2") = "Advanced"
'@@							Call Fn_CMS_ImportExportSettings_Operation("SetImportExportOptions",dicIEOptions,"OK")
'@@
'@@-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@	   History			:	Developer Name			Date	   		Rev. No.		Changes Done										Reviewer
'@@-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@	  						Poonam Chopade	 	30-Apr-2018			1.0			Created - Added for PSM new TC's development		[TC115-2018040200-30Apr2018-PoonamC-NewDevelopment]
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_CMS_ImportExportSettings_Operation(sAction,dicDetails,sButton)
	GBL_FAILED_FUNCTION_NAME="Fn_CMS_ImportExportSettings_Operation"
	Dim objRemoteIEOptions,dicCount, dicItems, dicKeys,aOptions,aOptionsVals
	Dim iCounter, iCount, bFlag,sSubAction, sProperty

	Fn_CMS_ImportExportSettings_Operation = False
	Set objRemoteIEOptions = Fn_SISW_CMS_GetObject("RemoteImportExportOptions")
	
	If Fn_UI_ObjectExist("Fn_CMS_ImportExportSettings_Operation",objRemoteIEOptions) = False Then
		Set objRemoteIEOptions = Fn_SISW_CMS_GetObject("RemoteImportExportOptions@1")
	End If
	
	'Check Dialog existence
	If Fn_UI_ObjectExist("Fn_CMS_ImportExportSettings_Operation",objRemoteIEOptions)  = False  Then
		Set objRemoteIEOptions = Nothing
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_CMS_ImportExportSettings_Operation ] Remote Import/Export Options dialog dose not exists ].")
		Exit Function
	End If
	
	Select Case sAction
		Case "SetImportExportOptions"
				dicCount = dicDetails.Count
				dicItems = dicDetails.Items
				dicKeys = dicDetails.Keys
				
				For iCounter = 0 To dicCount - 1
					If Instr(dicKeys(iCounter),"SelectTab")>0 Then
						sSubAction = "SelectTab"
					Else
						sSubAction = dicKeys(iCounter)
					End If
					sProperty = dicItems(iCounter)
					bFlag = False
					
					Select Case sSubAction
						Case "SelectTab"   					'Select Tab Name
							If sProperty<>"" Then
								objRemoteIEOptions.JavaTab("Tab").Select sProperty
								wait 1
								If Err.Number >= 0 Then
									bFlag = True
								End If
							End If
						Case "TransferOptions"   					'Select Transfer options
								aOptions = Split(sProperty,"~")
								For iCount = 0 To UBound(aOptions)
									aOptionsVals = Split(aOptions(iCount),":")
									objRemoteIEOptions.JavaCheckBox("TransferOptions").SetTOProperty "attached text",aOptionsVals(0)
									bFlag = Fn_SISW_UI_JavaCheckBox_Operations("Fn_CMS_ImportExportSettings_Operation", "Set", objRemoteIEOptions, "TransferOptions", aOptionsVals(1))
									If bFlag = False Then
										Exit For
									End If  
									wait 1
								Next
						Case "ItemOptions","CustomItemOptions"   					'Select Item Options
								aOptions = Split(sProperty,"~")
								For iCount = 0 To UBound(aOptions)
									aOptionsVals = Split(aOptions(iCount),":")
									objRemoteIEOptions.JavaCheckBox("ItemOptions").SetTOProperty "attached text",aOptionsVals(0)
									bFlag = Fn_SISW_UI_JavaCheckBox_Operations("Fn_CMS_ImportExportSettings_Operation", "Set", objRemoteIEOptions, "ItemOptions", aOptionsVals(1))
									If bFlag = False Then
										Exit For
									End If  
									wait 1
								Next
						Case "ItemOptionsList"   					'Select value from dropdown list in Item Options
								bFlag = Fn_SISW_UI_JavaList_Operations("Fn_CMS_ImportExportSettings_Operation","Select", objRemoteIEOptions, "ItemOptionsList", sProperty, "", "")
						Case "GeneralOptions","CustomGeneralOptions"   					'Select General Options
								aOptions = Split(sProperty,"~")
								For iCount = 0 To UBound(aOptions)
									aOptionsVals = Split(aOptions(iCount),":")
									objRemoteIEOptions.JavaCheckBox("GeneralOptions").SetTOProperty "attached text",aOptionsVals(0)
									bFlag = Fn_SISW_UI_JavaCheckBox_Operations("Fn_CMS_ImportExportSettings_Operation", "Set", objRemoteIEOptions, "GeneralOptions", aOptionsVals(1))
									If bFlag = False Then
										Exit For
									End If
									Wait 1									
								Next
						Case "TCXMLsessionOptions"   					'Select TC XML session Options
								bFlag = Fn_SISW_UI_JavaList_Operations("Fn_CMS_ImportExportSettings_Operation","Select", objRemoteIEOptions, "TCXMLsessionoptions", sProperty, "", "")
						Case "SaveOptions"   					'Select Save Options
								aOptionsVals = Split(sProperty,":")
								objRemoteIEOptions.JavaCheckBox("SaveOptions").SetTOProperty "attached text",aOptionsVals(0)
								bFlag = Fn_SISW_UI_JavaCheckBox_Operations("Fn_CMS_ImportExportSettings_Operation", "Set", objRemoteIEOptions, "SaveOptions", aOptionsVals(1))	
					   	Case "DatasetOptions"   					'Select dataset Options
								aOptions = Split(sProperty,"~")
								For iCount = 0 To UBound(aOptions)
									aOptionsVals = Split(aOptions(iCount),":")
									objRemoteIEOptions.JavaCheckBox("DatasetOptions").SetTOProperty "attached text",aOptionsVals(0)
									bFlag = Fn_SISW_UI_JavaCheckBox_Operations("Fn_CMS_ImportExportSettings_Operation", "Set", objRemoteIEOptions, "DatasetOptions", aOptionsVals(1))
									If bFlag = False Then
										Exit For
									End If
									Wait 1									
								Next
						Case "StructureManagerOptions"   					'Select Structure Manager Options
								aOptions = Split(sProperty,"~")
								For iCount = 0 To UBound(aOptions)
									aOptionsVals = Split(aOptions(iCount),":")
									objRemoteIEOptions.JavaCheckBox("StructureMgrOptions").SetTOProperty "attached text",aOptionsVals(0)
									bFlag = Fn_SISW_UI_JavaCheckBox_Operations("Fn_CMS_ImportExportSettings_Operation", "Set", objRemoteIEOptions, "StructureMgrOptions", aOptionsVals(1))
									If bFlag = False Then
										Exit For
									End If  
								Next
						Case "SessionOptions"   					'Select Session Options
								aOptions = Split(sProperty,"~")
								For iCount = 0 To UBound(aOptions)
									aOptionsVals = Split(aOptions(iCount),":")
									objRemoteIEOptions.JavaCheckBox("SessionOptions").SetTOProperty "attached text",aOptionsVals(0)
									bFlag = Fn_SISW_UI_JavaCheckBox_Operations("Fn_CMS_ImportExportSettings_Operation", "Set", objRemoteIEOptions, "SessionOptions", aOptionsVals(1))
									If bFlag = False Then
										Exit For
									End If  
								Next						
						Case "AddIncludeRelation"   					'Add relations from Exclude Reference list to Include Reference list
								aOptions = Split(sProperty,"~")
								For iCount = 0 To UBound(aOptions)
									bFlag = Fn_SISW_UI_JavaList_Operations("Fn_CMS_ImportExportSettings_Operation","Exist", objRemoteIEOptions, "Include Reference", aOptions(iCount),"", "")
									If bFlag = False Then
										bFlag = Fn_SISW_UI_JavaList_Operations("Fn_CMS_ImportExportSettings_Operation","Select", objRemoteIEOptions, "Exclude Reference", aOptions(iCount),"", "")
										If bFlag = False Then
											Exit For
										End If
										'Click button
										Call Fn_Button_Click("Fn_CMS_ImportExportSettings_Operation",objRemoteIEOptions,"RelationAdd")
									End If
								Next										
						Case "AddExcludeRelation"   					'Add relations from Include Reference list to Exclude Reference list
								aOptions = Split(sProperty,"~")
								For iCount = 0 To UBound(aOptions)
									bFlag = Fn_SISW_UI_JavaList_Operations("Fn_CMS_ImportExportSettings_Operation","Exist", objRemoteIEOptions, "Exclude Reference", aOptions(iCount),"", "")
									If bFlag = False Then
										bFlag = Fn_SISW_UI_JavaList_Operations("Fn_CMS_ImportExportSettings_Operation","Select", objRemoteIEOptions, "Include Reference", aOptions(iCount),"", "")
										If bFlag = False Then
											Exit For
										End If
										'Click button
										Call Fn_Button_Click("Fn_CMS_ImportExportSettings_Operation",objRemoteIEOptions,"RelationRemove")
									End If
								Next
						Case "SyncNotificationOptions"   					'Select Synchronization/Notification Options
								aOptions = Split(sProperty,"~")
								For iCount = 0 To UBound(aOptions)
									aOptionsVals = Split(aOptions(iCount),":")
									objRemoteIEOptions.JavaCheckBox("SyncNotifnOptions").SetTOProperty "attached text",aOptionsVals(0)
									bFlag = Fn_SISW_UI_JavaCheckBox_Operations("Fn_CMS_ImportExportSettings_Operation", "Set", objRemoteIEOptions, "SyncNotifnOptions", aOptionsVals(1))
									If bFlag = False Then
										Exit For
									End If  
								Next
						Case "SendOption_SelectNewUser"   					'Select New Owning User from Send Options	
								'For new User selection code here
						Case "SendOption_UsedefaultUserGrp"	   					'Select send Option as Use default user/group ownership rules	
							aOptionsVals = Split(sProperty,":")							 
							objRemoteIEOptions.JavaCheckBox("Usedefaultusergroup").SetTOProperty "attached text",aOptionsVals(0)
							bFlag = Fn_SISW_UI_JavaCheckBox_Operations("Fn_CMS_ImportExportSettings_Operation", "Set", objRemoteIEOptions, "Usedefaultusergroup", aOptionsVals(1)) 								
						
						Case "AssembliesOption"	   					'Select Assemblies With Part Family Member Components									
							aOptionsVals = Split(sProperty,":")
							objRemoteIEOptions.JavaRadioButton("AssembliesOption").SetTOProperty "attached text",aOptionsVals(0)
							bFlag = Fn_SISW_UI_JavaRadioButton_Operations("Fn_CMS_ImportExportSettings_Operation", "Set", objRemoteIEOptions, "AssembliesOption", aOptionsVals(1))					
						
						Case "PartFamiltyTemplates"	   					'Select Part Family Templates 		
							aOptionsVals = Split(sProperty,":")
							objRemoteIEOptions.JavaCheckBox("PartFamiltyTempOption").SetTOProperty "attached text",aOptionsVals(0)
							bFlag = Fn_SISW_UI_JavaCheckBox_Operations("Fn_CMS_ImportExportSettings_Operation", "Set", objRemoteIEOptions, "PartFamiltyTempOption", aOptionsVals(1))
							
						Case "PartFamiltyMembers"	   					'Select Part Family Members option 			
							aOptionsVals = Split(sProperty,":")
							objRemoteIEOptions.JavaCheckBox("PartFamiltyMemberOption").SetTOProperty "attached text",aOptionsVals(0)
							bFlag = Fn_SISW_UI_JavaCheckBox_Operations("Fn_CMS_ImportExportSettings_Operation", "Set", objRemoteIEOptions, "PartFamiltyMemberOption", aOptionsVals(1))
							
						Case "4GDOptions"	   					'Select 4GD Options				
							aOptions = Split(sProperty,"~")
							For iCount = 0 To UBound(aOptions)
								aOptionsVals = Split(aOptions(iCount),":")
								objRemoteIEOptions.JavaCheckBox("4GDOptions").SetTOProperty "attached text",aOptionsVals(0)
								bFlag = Fn_SISW_UI_JavaCheckBox_Operations("Fn_CMS_ImportExportSettings_Operation", "Set", objRemoteIEOptions, "4GDOptions", aOptionsVals(1))
								If bFlag = False Then
									Exit For
								End If 
								Wait 1								
							Next
						Case "APSContainerOptions"	   					'Select APS Container Options 	
							aOptionsVals = Split(sProperty,":")
							objRemoteIEOptions.JavaCheckBox("APSOptions").SetTOProperty "attached text",aOptionsVals(0)
							bFlag = Fn_SISW_UI_JavaCheckBox_Operations("Fn_CMS_ImportExportSettings_Operation", "Set", objRemoteIEOptions, "APSOptions", aOptionsVals(1))
						
					End Select
					
					If bFlag = False Then
						Fn_CMS_ImportExportSettings_Operation = False
						Set objRemoteIEOptions = Nothing
						Call Fn_WriteLogFile("","FAIL : Function [Fn_CMS_ImportExportSettings_Operation] Failed to Perform Case ["&sAction&"] SubCase ["+sSubAction+"].")
						Exit Function
					End If
				Next				
		Case "IsChecked"
				dicCount = dicDetails.Count
				dicItems = dicDetails.Items
				dicKeys = dicDetails.Keys
				For iCounter = 0 To dicCount - 1
					If Instr(dicKeys(iCounter),"SelectTab")>0 Then
						sSubAction = "SelectTab"
					ElseIf Instr(dicKeys(iCounter),"CheckBox")>0 Then
						sSubAction = "CheckBox"	
					ElseIf Instr(dicKeys(iCounter),"JavaList")>0 Then
						sSubAction = "JavaList"		
					Else
						sSubAction = ""
					End If
					sProperty = dicItems(iCounter)
					bFlag = False
					
					Select Case sSubAction
						Case "SelectTab"   					'Select Tab Name
							If sProperty<>"" Then
								objRemoteIEOptions.JavaTab("Tab").Select sProperty
								wait 1
								If Err.Number >= 0 Then
									bFlag = True
								End If
							End If
						Case "CheckBox"   					'Select Transfer options
								aOptions = Split(sProperty,"~")
								For iCount = 0 To UBound(aOptions)
									aOptionsVals = Split(aOptions(iCount),":")
									objRemoteIEOptions.JavaCheckBox("GeneralOptions").SetTOProperty "attached text",aOptionsVals(0) 
									If cbool(Fn_UI_Object_GetROProperty("Fn_CMS_ImportExportSettings_Operation",objRemoteIEOptions.JavaCheckBox("GeneralOptions"),"value")) = cbool(aOptionsVals(1)) Then
										bFlag = True
										Exit For
									End If  
								Next
						Case "JavaList"	
							aOptions = Split(sProperty,":")
							aOptionsVals = Split(aOptions(1),"~")
							For iCount = 0 To UBound(aOptionsVals)
								bFlag = Fn_SISW_UI_JavaList_Operations("Fn_CMS_ImportExportSettings_Operation","Exist", objRemoteIEOptions, aOptions(0), aOptionsVals(iCount),"", "")
								If bFlag = False Then
									Exit For
								End If
							Next	
					End Select
					
					If bFlag = False Then
						Fn_CMS_ImportExportSettings_Operation = False
						Set objRemoteIEOptions = Nothing
						Call Fn_WriteLogFile("","FAIL : Function [Fn_CMS_ImportExportSettings_Operation] Failed to Perform Case ["&sAction&"] SubCase ["+sSubAction+"].")
						Exit Function
					End If
				Next
				
			Case "IsEnabled"
				'For future use			
		End Select
		
		If sButton<>"" Then
			Call Fn_Button_Click("Fn_CMS_ImportExportSettings_Operation",objRemoteIEOptions,sButton)
			Wait 1
		End If
		
		Fn_CMS_ImportExportSettings_Operation = True
		Set objRemoteIEOptions = Nothing
End Function
'=====================================================================================================================================================================================
'@@    Function Name			:	Fn_CMS_OptionsSettings_Ops
'@@
'@@    Description				:	Function used to perform operations on Remote Site Selection dialog
'@@
'@@    Parameters			   	:	1. sAction		: ActionName
'@@								:	2. sSites		: Sites names (Seperated with ~)
'@@								:   3. sButtons 	: Button Names										 	         
'@@
'@@    Return Value		   	   	: 	True Or False
'@@
'@@    Examples					:	Set dicOptionSettings = CreateObject("Scripting.Dictionary")
'@@										dicOptionSettings("Transfer Options") = "None"
'@@										dicOptionSettings("Item Options") = "Include all revisions"
'@@										dicOptionSettings("General Options") = "None"
'@@										dicOptionSettings("Structure Manager Options") = "Include entire BOM"
'@@									Call Fn_CMS_OptionsSettings_Ops("VerifyOptionsSettings",dicOptionSettings,"Yes")
'@@    History					:	
'@@ ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@	  Developer Name			Date			 Rev. No.	   				Changes Done								 			Reviewer
'@@ ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@	  Poonam Chopade	 	02-May-2018	 		  1.0			Created - Added for MultiSite new TC's development			TC11.5(20180402.00)_CMS_NewDevelopment_PoonamC_02Apr2018
'=====================================================================================================================================================================================
Public Function Fn_CMS_OptionsSettings_Ops(sAction,dicDetails,sButton)
	GBL_FAILED_FUNCTION_NAME="Fn_CMS_OptionsSettings_Ops"
	Dim ObjOptionsSettings,iCounter, sOptionName, sProperty,aOptionsVals
	Dim dicCount,dicItems,dicKeys,iCount,sValues
	
	Set ObjOptionsSettings = Fn_SISW_CMS_GetObject("OptionsSettings")
	If Fn_UI_ObjectExist("Fn_CMS_OptionsSettings_Ops",ObjOptionsSettings) = False Then
		Set ObjOptionsSettings = Fn_SISW_CMS_GetObject("OptionsSettings@1")
		Set ObjOptionsSettings = Fn_UI_ObjectCreate("Fn_CMS_OptionsSettings_Ops",ObjOptionsSettings)
	End If
	
	Fn_CMS_OptionsSettings_Ops = False
	
	Select Case sAction
		Case "VerifyOptionsSettings"
			dicCount = dicDetails.Count
			dicItems = dicDetails.Items
			dicKeys = dicDetails.Keys
			
			For iCounter = 0 To dicCount - 1
				sOptionName = dicKeys(iCounter)
				sProperty = dicItems(iCounter)
				
				ObjOptionsSettings.JavaObject("MLabel").SetTOProperty "attached text",sOptionName
				sValues = ObjOptionsSettings.JavaObject("MLabel").GetROProperty("text")
				
				'Split values if multi line
				aOptionsVals = Split(sProperty,"~")
				For iCount = 0 To UBound(aOptionsVals)
					If instr(sValues,aOptionsVals(iCount)) = 0  Then
						Fn_CMS_OptionsSettings_Ops = False
						Set ObjOptionsSettings = Nothing
						Call Fn_WriteLogFile("","FAIL : Function [Fn_CMS_OptionsSettings_Ops] Failed to verify ["&sOptionName&"] as ["+sValues+"].")
						Exit Function
					End If
				Next
		   Next
	End Select 
	
	'Click on button  
	If sButton<>"" Then
		Call Fn_Button_Click("Fn_CMS_OptionsSettings_Ops",ObjOptionsSettings,sButton)
		wait 1
		Call Fn_ReadyStatusSync(5)
	End If
		
	Fn_CMS_OptionsSettings_Ops = TRUE
    Set ObjOptionsSettings = nothing 	
	
End Function
'=====================================================================================================================================================================================
'@@
'@@    Function Name			:	Fn_CMS_ExportObjects_Operation
'@@
'@@    Description				:	Function used to perform Export Object 
'@@
'@@    Parameters			   	:	1. sAction		: ActionName
'@@								:	2. dicEODetails	: Information for Export Objects
'@@								:   3. dicEOOptions : Options Settings info for Export										 	         
'@@								:	4. sButtons : Buttons name
'@@
'@@    Return Value		   	   	: 	True Or False
'@@
'@@   Example 					:	Set dicEOOptions = CreateObject("Scripting.Dictionary")
'@@										dicEOOptions("SelectTab1") = "General"
'@@										dicEOOptions("ItemOptions") = "Include all revisions:ON"
'@@										dicEOOptions("DatasetOptions") = "Include all versions:ON~Include all files:ON"
'@@										dicEOOptions("StructureManagerOptions") = "Include entire BOM:ON"
'@@										dicEOOptions("SelectTab2") = "Advanced"
'@@									
'@@									Set dicEODetails = CreateObject("Scripting.Dictionary")
'@@										dicEODetails("OptionName") = "Teamcenter"
'@@										dicEODetails("ParentDirectory") = "C:\Temp"
'@@										dicEODetails("ExportDirectory") = "000021-Item1"										
'@@										dicEODetails("Reason") = "Test1"
'@@										dicEODetails("TargetSites") = "SiteB--1886227743"
'@@									bReturn = Fn_CMS_ExportObjects_Operation("ExportObjects",dicEODetails,dicEOOptions,"OK")
'@@
'@@    History					:	
'@@ ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@	  Developer Name			Date			 Rev. No.	   				Changes Done								 			Reviewer
'@@ ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@	  Poonam Chopade	 	03-May-2018	 		  1.0			Created - Added for MultiSite new TC's development			TC11.5(20180402.00)_CMS_NewDevelopment_PoonamC_03May2018
'=====================================================================================================================================================================================
Public Function Fn_CMS_ExportObjects_Operation(sAction,dicEODetails,dicEOOptions,sButtons)
	GBL_FAILED_FUNCTION_NAME="Fn_CMS_ExportObjects_Operation"
	Dim ObjExport,sMenu,bReturn,ObjSelectDir

	Fn_CMS_ExportObjects_Operation = False
	Set ObjExport = Fn_SISW_CMS_GetObject("ExportObject")
	
	'Check Dialog existence
	If Fn_UI_ObjectExist("Fn_CMS_ExportObjects_Operation",ObjExport)  = False  Then
		sMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("MyTc_Menu"),"ToolsExportObjects")
		Call Fn_MenuOperation("Select",sMenu)
		Call Fn_ReadyStatusSync(2)
		
		'Check Dialog existence
		If Fn_UI_ObjectExist("Fn_CMS_ExportObjects_Operation",ObjExport)  = False  Then
			Set ObjExport = Nothing
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_CMS_ExportObjects_Operation ] Remote Export dialog dose not exists ].")
			Exit Function
		End If
	End If
	
	Set ObjSelectDir = 	Fn_SISW_CMS_GetObject("SelectDirectory")
	
	Select Case sAction
		Case "ExportObjects"
			 'Select check box
			  If dicEODetails("OptionName") <> "" Then 
					ObjExport.JavaCheckBox("OptionName").SetTOProperty "attached text",dicEODetails("OptionName")
					bReturn = Fn_SISW_UI_JavaCheckBox_Operations("Fn_CMS_ExportObjects_Operation", "Set", ObjExport, "OptionName", "ON")
					If bReturn = False Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_CMS_ExportObjects_Operation ]  Failed to click on [ "+ dicEODetails("OptionName")+" ]") 
						Set ObjExport = Nothing
						Set ObjSelectDir = Nothing
						Fn_CMS_ExportObjects_Operation = False
						Exit Function
					 End If
			  End If		
			 'Enter Parent Directory
			 If dicEODetails("ParentDirectory") <> "" Then 
			 	 Call Fn_Button_Click("Fn_CMS_ExportObjects_Operation",ObjExport,"browse") 'click on browse button
			 	 Wait 2
'				 bReturn = Fn_SISW_UI_JavaEdit_Operations("Fn_CMS_ExportObjects_Operation", "Set", ObjSelectDir, "Foldername", dicEODetails("ParentDirectory"))	
				 ObjSelectDir.JavaEdit("Foldername").Set ""
				 Wait 2
				  ObjSelectDir.JavaEdit("Foldername").Type dicEODetails("ParentDirectory")
				 wait 2
				 bReturn = Fn_Button_Click("Fn_CMS_ExportObjects_Operation",ObjSelectDir,"Select") 
				 If bReturn = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_CMS_ExportObjects_Operation ]  Failed to enter Parent Directory [ "+ dicEODetails("ParentDirectory")  +" ]") 
					Set ObjExport = Nothing
					Set ObjSelectDir = Nothing
					Fn_CMS_ExportObjects_Operation = False
					Exit Function
				 End If
				 wait 1
			 End If	
			'Enter Export Directory
			If dicEODetails("ExportDirectory") <> "" Then 
				 bReturn = Fn_SISW_UI_JavaEdit_Operations("Fn_CMS_ExportObjects_Operation", "Set", ObjExport, "ExportDirectory", dicEODetails("ExportDirectory"))	
				 Wait 1
				 If bReturn = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_CMS_ExportObjects_Operation ]  Failed to enter Export Directory [ "+ dicEODetails("ExportDirectory")  +" ]") 
					Set ObjExport = Nothing
					Set ObjSelectDir = Nothing
					Fn_CMS_ExportObjects_Operation = False
					Exit Function
				 End If
			 End If
			 'Enter Reason
			If dicEODetails("Reason") <> "" Then 
				 bReturn = Fn_SISW_UI_JavaEdit_Operations("Fn_CMS_ExportObjects_Operation", "Set", ObjExport, "Reason", dicEODetails("Reason"))	
				 Wait 1
				 If bReturn = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_CMS_ExportObjects_Operation ]  Failed to enter Reason as [ "+ dicEODetails("Reason")  +" ]") 
					Set ObjExport = Nothing
					Set ObjSelectDir = Nothing
					Fn_CMS_ExportObjects_Operation = False
					Exit Function
				 End If
			 End If	
			 'Select Target Sites
			 If dicEODetails("TargetSites") <> "" Then 
			 	 'Click on Select target remote site icon
			 	 Call Fn_Button_Click("Fn_CMS_ExportObjects_Operation",ObjExport,"site_16")
			 	 Wait 1
			 	 bReturn = Fn_CMS_RemoteSiteSelection_Ops("Add",dicEODetails("TargetSites"),"OK")
				 If bReturn = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_CMS_ExportObjects_Operation ]  Failed to select Target Sites as [ "+ dicEODetails("TargetSites")+" ]") 
					Set ObjExport = Nothing
					Set ObjSelectDir = Nothing
					Fn_CMS_ExportObjects_Operation = False
					Exit Function
				 End If
			 End If	
			 'Display or Set Export options
			 If vartype(dicEOOptions) = 9 Then 
			 	'Click on button Display/set remote export option
			 	Call Fn_Button_Click("Fn_CMS_ExportObjects_Operation",ObjExport,"exportsettings_16")
			 	Wait 1
				bReturn = Fn_CMS_ImportExportSettings_Operation("SetImportExportOptions",dicEOOptions,"OK")
				If bReturn = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_CMS_ExportObjects_Operation ]  Failed to set Export options") 
					Set ObjExport = Nothing
					Set ObjSelectDir = Nothing
					Fn_CMS_ExportObjects_Operation = False
					Exit Function
				 End If
			 End If
	End Select
	
	'Click on Buttons
	If sButtons<>"" Then
		sButtons = split(sButtons, ":",-1,1)
		For iCounter=0 to Ubound(sButtons)
			Call Fn_Button_Click("Fn_CMS_ExportObjects_Operation", ObjExport, sButtons(iCounter))
			Call Fn_ReadyStatusSync(2)
		Next
	End If
	
	Fn_CMS_ExportObjects_Operation = True
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_CMS_ExportObjects_Operation ] executed successfully.") 
	Set ObjExport = nothing
	Set ObjSelectDir = Nothing
	
End Function
'=====================================================================================================================================================================================
'@@
'@@    Function Name			:	Fn_CMS_ImportObjects_Operation
'@@
'@@    Description				:	Function used to perform Import Object 
'@@
'@@    Parameters			   	:	1. sAction		: ActionName
'@@								:	2. dicIODetails	: Information for Import Objects
'@@								:   3. sButtons 	: Buttons name									 	         
'@@								 	
'@@    Return Value		   	   	: 	True Or False
'@@
'@@   Example 					:  Set dicIODetails = CreateObject("Scripting.Dictionary")
'@@										dicIODetails("OptionName") = "Teamcenter"
'@@										dicIODetails("ImportingObject") = "C:\Temp\000021-item1"
'@@										dicIODetails("ObjectButton") = "Select All"										
'@@										dicIODetails("ReportOptions") = "Generate Import report:ON~Preview Import Report:ON"
'@@									bReturn = Fn_CMS_ImportObjects_Operation("ImportObjects",dicIODetails,"OK")
'@@
'@@    History					:	
'@@ ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@	  Developer Name			Date			 Rev. No.	   				Changes Done								 			Reviewer
'@@ ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@	  Poonam Chopade	 	03-May-2018	 		  1.0			Created - Added for MultiSite new TC's development			TC11.5(20180402.00)_CMS_NewDevelopment_PoonamC_03May2018
'=====================================================================================================================================================================================
Public Function Fn_CMS_ImportObjects_Operation(sAction,dicIODetails,sButtons)
	GBL_FAILED_FUNCTION_NAME="Fn_CMS_ImportObjects_Operation"
	Dim ObjImport,sMenu,bReturn

	Fn_CMS_ImportObjects_Operation = False
	Set ObjImport = Fn_SISW_CMS_GetObject("ImportObjects")
	
	'Check Dialog existence
	If Fn_UI_ObjectExist("Fn_CMS_ImportObjects_Operation",ObjImport)  = False  Then
		sMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("MyTc_Menu"),"ToolsImportObjects")
		Call Fn_MenuOperation("Select",sMenu)
		Call Fn_ReadyStatusSync(2)
		
		'Check Dialog existence
		If Fn_UI_ObjectExist("Fn_CMS_ImportObjects_Operation",ObjImport)  = False  Then
			Set ObjImport = Nothing
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_CMS_ImportObjects_Operation ] Remote Export dialog dose not exists ].")
			Exit Function
		End If
		
	End If
	
	Select Case sAction
		Case "ImportObjects"
			 'Select check box
			  If dicIODetails("OptionName") <> "" Then 
					ObjImport.JavaCheckBox("OptionName").SetTOProperty "attached text",dicIODetails("OptionName")
					bReturn = Fn_SISW_UI_JavaCheckBox_Operations("Fn_CMS_ImportObjects_Operation", "Set", ObjImport, "OptionName", "ON")
					If bReturn = False Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_CMS_ImportObjects_Operation ]  Failed to click on [ "+ dicIODetails("OptionName")+" ]") 
						Set ObjImport = Nothing
						Fn_CMS_ImportObjects_Operation = False
						Exit Function
					 End If
			  End If		
			 'select Importing Object
			 If dicIODetails("ImportingObject") <> "" Then 
				 Call Fn_Button_Click("Fn_CMS_ImportObjects_Operation",ObjImport,"browse_16") 'click on browse button
			 	 Wait 1
				 Call Fn_SISW_UI_JavaEdit_Operations("Fn_CMS_ImportObjects_Operation", "Type", ObjImport.JavaDialog("SelectObject"), "Filename", dicIODetails("ImportingObject"))
				 Wait 1
				 bReturn = Fn_Button_Click("Fn_CMS_ImportObjects_Operation",ObjImport.JavaDialog("SelectObject"),"Select") 
				 wait 2
				 If bReturn = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_CMS_ImportObjects_Operation ]  Failed to select Importing Object [ "+ dicIODetails("ImportingObject")  +" ]") 
					Set ObjImport = Nothing
					Fn_CMS_ImportObjects_Operation = False
					Exit Function
				 End If
			 End If	
			'Click Select All/Select None button
			If dicIODetails("ObjectButton") <> "" Then 
				 ObjImport.JavaButton("ObjectButton").SetTOProperty "label",dicIODetails("ObjectButton")
				 wait 1
				 bReturn = Fn_Button_Click("Fn_CMS_ImportObjects_Operation",ObjImport,"ObjectButton") 
				 If bReturn = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_CMS_ImportObjects_Operation ]  Failed to click on [ "+ dicIODetails("ObjectButton")+" ] button") 
					Set ObjImport = Nothing
					Fn_CMS_ImportObjects_Operation = False
					Exit Function
				 End If
			 End If
			 'Select Report Options
			If dicIODetails("ReportOptions") <> "" Then 
				aOptions = Split(dicIODetails("ReportOptions"),"~")
				For iCount = 0 To UBound(aOptions)
					aOptionsVals = Split(aOptions(iCount),":")
					ObjImport.JavaCheckBox("ReportOptions").SetTOProperty "attached text",aOptionsVals(0)
					bReturn = Fn_SISW_UI_JavaCheckBox_Operations("Fn_CMS_ImportObjects_Operation", "Set", ObjImport, "ReportOptions", aOptionsVals(1))
					If bReturn = False Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_CMS_ImportObjects_Operation ]  Failed to set [ "+aOptionsVals(0)+" as "+aOptionsVals(1)+"]") 
						Set ObjImport = Nothing
						Fn_CMS_ImportObjects_Operation = False
						Exit Function
					 End If 
				Next
			 End If	
	End Select
	
	'Click on Buttons
	If sButtons<>"" Then
		sButtons = split(sButtons, ":",-1,1)
		For iCounter=0 to Ubound(sButtons)
			Call Fn_Button_Click("Fn_CMS_ImportObjects_Operation", ObjImport, sButtons(iCounter))
			Call Fn_ReadyStatusSync(2)
		Next
	End If
	
	Fn_CMS_ImportObjects_Operation = True
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_CMS_ImportObjects_Operation ] executed successfully.") 
	Set ObjImport = nothing
	
End Function
'=====================================================================================================================================================================================
'@@
'@@    Function Name			:	Fn_CMS_Folder_Operation
'@@
'@@    Description				:	Function used to perform operation with folder
'@@
'@@    Parameters			   	:	1. sAction		: ActionName
'@@								:	2. sFolderPath	: folder path
'@@								:   3. sReserve 	: for future use									 	         
'@@								 	
'@@    Return Value		   	   	: 	True Or False
'@@
'@@   Example 					:  Call Fn_CMS_Folder_Operation("Create","C:\logs","")
'@@								   Call Fn_CMS_Folder_Operation("Exists","C:\logs","")						
'@@								   Call Fn_CMS_Folder_Operation("Delete","C:\logs","")	
'@@    History					:	
'@@ ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@	  Developer Name			Date			 Rev. No.	   				Changes Done								 			Reviewer
'@@ ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@	  Poonam Chopade	 	04-May-2018	 		  1.0			Created - Added for MultiSite new TC's development			TC11.5(20180402.00)_CMS_NewDevelopment_PoonamC_04May2018
'=====================================================================================================================================================================================
Public Function Fn_CMS_Folder_Operation(sAction,sFolderPath,sReserve)
	GBL_FAILED_FUNCTION_NAME="Fn_CMS_Folder_Operation"
	Dim objFSO,objFolder
	
	Fn_CMS_Folder_Operation = False
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	
	Select Case sAction
		Case "Create"
			Set objFolder = objFSO.CreateFolder(sFolderPath)	
			If objFSO.FolderExists(sFolderPath) = True	Then	
				Fn_CMS_Folder_Operation = True
			End if
		Case "Exists"	
			If objFSO.FolderExists(sFolderPath) = True	Then	
				Fn_CMS_Folder_Operation = True
			End if
		Case "Delete"
			Call objFSO.DeleteFolder(sFolderPath,True)		
			If objFSO.FolderExists(sFolderPath) = False	Then	
				Fn_CMS_Folder_Operation = True
			End if
	End Select
	
	Set objFSO = Nothing
	Set objFolder = Nothing
	
End Function
'=====================================================================================================================================================================================
'@@    Function Name			:	Fn_CMS_RemoteImportProgress_Operations
'@@
'@@    Description				:	Function used to perform operations on Remote Import Progress dialog
'@@
'@@    Parameters			   	:	1. sAction		: ActionName
'@@								:	2. sObjectName	: Object Name
'@@								:   3. sColumns 	: Column names (Separated with ~)	
'@@								:	4. sValues 		: Values (Separated with ~)									 	         
'@@								:   5. sButton		: Button Name
'@@
'@@    Return Value		   	   	: 	True Or False
'@@
'@@    Examples					:	Call Fn_CMS_RemoteImportProgress_Operations("VerifyColValues","000107-Item1","Operation~Progress (Objects/Bytes/Files)","Remote Import Complete~Succeeded",Close)
'@@
'@@    History					:	
'@@ ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@	  Developer Name			Date			 Rev. No.	   				Changes Done								 			Reviewer
'@@ ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@	  Poonam Chopade	 	16-May-2018	 		  1.0			Created - Added for MultiSite new TC's development			TC11.5(20180402.00)_CMS_NewDevelopment_PoonamC_16May2018
'=====================================================================================================================================================================================
Public Function Fn_CMS_RemoteImportProgress_Operations(sAction,sObjectName,sColumns,sValues,sButton)
	GBL_FAILED_FUNCTION_NAME="Fn_CMS_RemoteImportProgress_Operations"
	Dim objRIProgressDialog,iCounter, iRows, sAppvalue,bFlag ,iRowIndex,aCols,aVals
	
	Set objRIProgressDialog = Fn_SISW_CMS_GetObject("RemoteImportProgress")
	Set objRIProgressDialog = Fn_UI_ObjectCreate("Fn_CMS_RemoteImportProgress_Operations",objRIProgressDialog)
	Fn_CMS_RemoteImportProgress_Operations = False
	
	'Activate Dialog
	objRIProgressDialog.Activate()
	Wait 1
	
	Select Case sAction 
		   Case "VerifyColValues"
			 If sObjectName <> "" Then
					iRows = objRIProgressDialog.JavaTable("JTable").GetROProperty("rows")
					For iCounter = 0 to iRows-1
						bFlag = False
						sAppvalue = Fn_UI_JavaTable_GetCellData("Fn_CMS_RemoteImportProgress_Operations",objRIProgressDialog,"JTable",iCounter,"Object Name")
						If Instr(sAppvalue,sObjectName) > 0 Then
							bFlag = True
							iRowIndex = iCounter
							Exit For
						End If
					Next
					'Verify Column values for object
					aCols = Split(sColumns,"~")
					aVals = Split(sValues,"~")
					For iCounter = 0 to UBound(aCols)	
						sAppvalue = Fn_UI_JavaTable_GetCellData("Fn_CMS_RemoteImportProgress_Operations",objRIProgressDialog,"JTable",iRowIndex,aCols(iCounter))	
						If Instr(sAppvalue,aVals(iCounter)) = 0 Then
							Set objRIProgressDialog = Nothing
							Fn_CMS_RemoteImportProgress_Operations = False
							Exit Function
						End If
					Next	
			 End If			
	End Select
	
	'Click on Buttons
	If sButton <> "" Then
		Call Fn_Button_Click("Fn_CMS_RemoteImportProgress_Operations", objRIProgressDialog,sButton)
        Call Fn_ReadyStatusSync(2)
	End If
		
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" Successfully completed of function Fn_CMS_RemoteImportProgress_Operations")
	Fn_CMS_RemoteImportProgress_Operations = TRUE
    Set objRIProgressDialog = nothing 	
	
End Function
'=====================================================================================================================================================================================
'@@    Function Name			:	Fn_CMS_ImportRemoteDialog_Operations
'@@
'@@    Description				:	Function used to perform operations on Import Remote Dialog
'@@
'@@    Parameters			   	:	1. sAction		: ActionName
'@@								:	2. sNodeNames	: Item Names
'@@								:   3. sOption		: Option ( All or None ) 
'@@								:   3. sButton 		: Button Name ( OK or Cancel )										 	         
'@@								:   4. sReserve		: for future use
'@@
'@@    Return Value		   	   	: 	True Or False
'@@
'@@    Examples					:	Call Fn_CMS_ImportRemoteDialog_Operations("SelectOption","","All","OK","")
'@@ 								Call Fn_CMS_ImportRemoteDialog_Operations("VerifySelectList","000552-SubItem1~000554-SubItem2~000556-SubItem3","","Cancel","")
'@@
'@@    History					:	
'@@ ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@	  Developer Name			Date			 Rev. No.	   				Changes Done								 			Reviewer
'@@ ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@	  Poonam Chopade	 	29-May-2018	 		  1.0			Created - Added for MultiSite new TC's development			TC11.5(20180402.00)_CMS_NewDevelopment_PoonamC_29May2018
'=====================================================================================================================================================================================
Public Function Fn_CMS_ImportRemoteDialog_Operations(sAction,sNodeNames,sOption,sButton,sReserve)
	GBL_FAILED_FUNCTION_NAME="Fn_CMS_ImportRemoteDialog_Operations"
	Dim objImpRemoteDialog,iCounter,bReturn,aNodes,sMenu  

	Set objImpRemoteDialog = Fn_SISW_CMS_GetObject("ImportRemote")
	Fn_CMS_ImportRemoteDialog_Operations = False
	
	'Check Dialog existence
	If Fn_UI_ObjectExist("Fn_CMS_ImportRemoteDialog_Operations",objImpRemoteDialog)  = False  Then
		sMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("RAC_Menu"),"ToolsImportRemote")
		Call Fn_MenuOperation("Select",sMenu)
		Call Fn_ReadyStatusSync(2)
		
		'Check Dialog existence
		If Fn_UI_ObjectExist("Fn_CMS_ImportRemoteDialog_Operations",objImpRemoteDialog)  = False  Then
			Set objImpRemoteDialog = Nothing
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_CMS_ImportRemoteDialog_Operations ] Import Remote dialog dose not exists ].")
			Exit Function
		End If
	End If
	
	Select Case sAction 
			Case "SelectOption"				
				If sOption <> "" Then
						bReturn = Fn_Button_Click("Fn_CMS_ImportRemoteDialog_Operations", objImpRemoteDialog,sOption)
						Call Fn_ReadyStatusSync(2)
						If bReturn = False Then
							Set objImpRemoteDialog = Nothing
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fail to click on [ "&sOption&" ] button in Import Remote dialog.")
							Exit Function
						End If
				End If
			Case "VerifySelectList"	
				If sNodeNames <> "" Then
						aNodes = split(sNodeNames, "~",-1,1)
						For iCounter = 0 to UBound(aNodes)
							bReturn = Fn_SISW_UI_JavaList_Operations("Fn_CMS_ImportRemoteDialog_Operations", "Exist", objImpRemoteDialog, "Select", aNodes(iCounter), "", "")
							If bReturn = False Then
								Set objImpRemoteDialog = Nothing
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fail to verify List Item [ "&aNodes(iCounter)&" ] in Import Remote dialog.")
								Exit Function
							End If
						Next
				End If				
	End Select
	
	'Click button OK / Cancel
	If sButton <> "" Then
		Call Fn_Button_Click("Fn_CMS_ImportRemoteDialog_Operations", objImpRemoteDialog,sButton)
		Call Fn_ReadyStatusSync(2)
	End If
	
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" Successfully completed of function Fn_CMS_ImportRemoteDialog_Operations")
	Fn_CMS_ImportRemoteDialog_Operations = TRUE
    Set objImpRemoteDialog = nothing 	
	
End Function

'=====================================================================================================================================================================================
'@@
'@@    Function Name			:	Fn_CMS_MultiSiteSynchronisation_Operation
'@@
'@@    Description				:	Function used to perform Multi Site Synchronisation on Component or Assembly
'@@
'@@    Parameters			   	:	1. sAction		: ActionName
'@@								:	2. dicMSOptions	: Information for to remove
'@@								:   3. sButton 		: Button Name										 	         
'@@
'@@    Return Value		   	   	: 	True Or False
'@@
'@@   Example 					:	Set dicMSOptions = CreateObject("Scripting.Dictionary")
'@@ 									dicMSOptions("SelectTab1") = "General"
'@@ 									dicMSOptions("TCXMLsessionOptions") = "MultiSiteExpOptSet"
'@@ 									dicMSOptions("SyncOptions") = "Report Only:ON"
'@@ 									dicMSOptions("RevisionRuleOptions") = "Specific Revision Rule:ON"
'@@ 									dicMSOptions("GeneralOptions") = "Exclude folder contents:ON"
'@@										
'@@									Set dicMSOptionSettings = CreateObject("Scripting.Dictionary")
'@@										dicMSOptionSettings("Sync Options") = "Report Only"
'@@										dicMSOptionSettings("Revision Rule Options") = "Specific Revision Rule"
'@@									
'@@									sMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("MyTc_Menu"),"ToolsMultiSiteSynchronisationComponent")
'@@									bReturn = Fn_CMS_MultiSiteSynchronisation_Operation("MultiSiteSyncComponent",dicMSOptions,"Yes",dicMSOptionSettings,sMenu)
'@@
'@@    History					:	
'@@ ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@	  Developer Name			Date			 Rev. No.	   				Changes Done								 			Reviewer
'@@ ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@	  Pravin Bhoyar	 	    28-May-2018	 		  1.0			Created - Added for MultiSite new TC's development			TC11.5(20180402.00)_CMS_NewDevelopment_BhoyarP_28May2018
'=====================================================================================================================================================================================
Public Function Fn_CMS_MultiSiteSynchronisation_Operation(sAction,dicMSOptions,sContinueButton,dicMSOptionsettings,sMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_CMS_MultiSiteSynchronisation_Operation"
	Dim ObjMultiSiteSync,bReturn

	Fn_CMS_MultiSiteSynchronisation_Operation = False
	Set ObjMultiSiteSync = Fn_SISW_CMS_GetObject("MultiSiteSynchronisation")
	
	'Check Dialog existence
	If Fn_UI_ObjectExist("Fn_CMS_MultiSiteSynchronisation_Operation",ObjMultiSiteSync)  = False  Then
		Call Fn_MenuOperation("Select",sMenu)
		Call Fn_ReadyStatusSync(2)
		
		'Check Dialog existence
		If Fn_UI_ObjectExist("Fn_CMS_MultiSiteSynchronisation_Operation",ObjMultiSiteSync)  = False  Then
			Set ObjMultiSiteSync = Nothing
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_CMS_MultiSiteSynchronisation_Operation ] Remote Export dialog dose not exists ].")
			Exit Function
		End If
	End If
	
	Select Case sAction
		Case "MultiSiteSynchronisation"
		
			 If vartype(dicMSOptions) = 9 Then 
			 	'Click to change syncronisation preferences
			 	Call Fn_Button_Click("Fn_CMS_MultiSiteSynchronisation_Operation",ObjMultiSiteSync,"sync_object_16")
			 	Wait 1
				bReturn = Fn_CMS_MultiSiteSyncSettings_Operation("SetMultiSiteSyncOptions",dicMSOptions,"OK")
				If bReturn = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_CMS_MultiSiteSynchronisation_Operation ]  Failed to set Remote Export options") 
					Set ObjMultiSiteSync = Nothing
					Fn_CMS_MultiSiteSynchronisation_Operation = False
					Exit Function
				 End If
			 End If
			 'Click Yes/No
			 If sContinueButton <> "" Then 
			 	bReturn = Fn_Button_Click("Fn_CMS_MultiSiteSynchronisation_Operation",ObjMultiSiteSync,sContinueButton)
			 	If bReturn = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_CMS_MultiSiteSynchronisation_Operation ]  Failed to click on [ "&sContinueButton&" ] button") 
					Set ObjMultiSiteSync = Nothing
					Fn_CMS_MultiSiteSynchronisation_Operation = False
					Exit Function
				 End If
			 End If
			 'Check Display or Set Export options in Option Settings  
			 If vartype(dicMSOptionsettings) = 9 Then 
			 	bReturn = Fn_CMS_OptionsSettings_Ops("VerifyOptionsSettings",dicMSOptionsettings,"Yes")
			 	If bReturn = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_CMS_MultiSiteSynchronisation_Operation ]  Failed to verify set Synchronisation options") 
					Set ObjMultiSiteSync = Nothing
					Fn_CMS_MultiSiteSynchronisation_Operation = False
					Exit Function
				 End If
			 End If 
	End Select
	
	Fn_CMS_MultiSiteSynchronisation_Operation = True
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_CMS_MultiSiteSynchronisation_Operation ] executed successfully.") 
	Set ObjMultiSiteSync = nothing
End Function


'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name	:	Fn_CMS_MultiSiteSyncSettings_Operation
'@@
'@@    Description		:	Function Used to perform operations on "Remote Import/Export Options" dialog
'@@
'@@    Parameters		:	1. sAction		: Action to be performed
'@@						:	2. dicMSOptions	: Dictionary object
'@@						:	3. sButton		: OK / Cancel button
'@@
'@@    Return Value		: 	True Or False
'@@
'@@    Examples			:	Set dicMSOptions = CreateObject("Scripting.Dictionary")
'@@								dicMSOptions("SelectTab1") = "General"
'@@								dicMSOptions("TCXMLsessionOptions") = "MultiSiteExpOptSet"
'@@								dicMSOptions("SyncOptions") = "Report Only:ON"
'@@								dicMSOptions("RevisionRuleOptions") = "Specific Revision Rule:ON"
'@@								dicMSOptions("GeneralOptions") = "Exclude folder contents:ON"
'@@	
'@@							Call Fn_CMS_MultiSiteSyncSettings_Operation("SetMultiSiteSyncOptions",dicMSOptions,"OK")
'@@
'@@    History					:	
'@@ ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@	  Developer Name			Date			 Rev. No.	   				Changes Done								 			Reviewer
'@@ ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@	  Pravin Bhoyar	 	    28-May-2018	 		  1.0			Created - Added for MultiSite new TC's development			TC11.5(20180402.00)_CMS_NewDevelopment_BhoyarP_28May2018
'=====================================================================================================================================================================================
Public Function Fn_CMS_MultiSiteSyncSettings_Operation(sAction,dicDetails,sButton)
	GBL_FAILED_FUNCTION_NAME="Fn_CMS_MultiSiteSyncSettings_Operation"
	Dim objRemoteMSSOptions,dicCount, dicItems, dicKeys,aOptions,aOptionsVals
	Dim iCounter, iCount, bFlag,sSubAction, sProperty

	Fn_CMS_MultiSiteSyncSettings_Operation = False
	Set objRemoteMSSOptions = Fn_SISW_CMS_GetObject("MultiSiteSynchronisationOptions")
	
	'Check Dialog existence
	If Fn_UI_ObjectExist("Fn_CMS_MultiSiteSyncSettings_Operation",objRemoteMSSOptions)  = False  Then
		Set objRemoteMSSOptions = Nothing
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_CMS_MultiSiteSyncSettings_Operation ] Remote Import/Export Options dialog dose not exists ].")
		Exit Function
	End If
	
	Select Case sAction
		Case "SetMultiSiteSyncOptions"
				dicCount = dicDetails.Count
				dicItems = dicDetails.Items
				dicKeys = dicDetails.Keys
				
				For iCounter = 0 To dicCount - 1
					If Instr(dicKeys(iCounter),"SelectTab")>0 Then
						sSubAction = "SelectTab"
					Else
						sSubAction = dicKeys(iCounter)
					End If
					sProperty = dicItems(iCounter)
					bFlag = False
					
					Select Case sSubAction
						Case "SelectTab"   					'Select Tab Name
							If sProperty<>"" Then
								objRemoteMSSOptions.JavaTab("Tab").Select sProperty
								wait 1
								If Err.Number >= 0 Then
									bFlag = True
								End If
							End If
						Case "SyncOptions"   					'Select Sync options
								aOptions = Split(sProperty,"~")
								For iCount = 0 To UBound(aOptions)
									aOptionsVals = Split(aOptions(iCount),":")
									objRemoteMSSOptions.JavaCheckBox("SyncOptions").SetTOProperty "attached text",aOptionsVals(0)
									bFlag = Fn_SISW_UI_JavaCheckBox_Operations("Fn_CMS_MultiSiteSyncSettings_Operation", "Set", objRemoteMSSOptions, "SyncOptions", aOptionsVals(1))
									If bFlag = False Then
										Exit For
									End If  
									wait 1
								Next
						Case "RevisionRuleOptions"					'Select Revision Rule Options
								aOptions = Split(sProperty,"~")
								For iCount = 0 To UBound(aOptions)
									aOptionsVals = Split(aOptions(iCount),":")
									objRemoteMSSOptions.JavaCheckBox("RevisionRuleOptions").SetTOProperty "attached text",aOptionsVals(0)
									bFlag = Fn_SISW_UI_JavaCheckBox_Operations("Fn_CMS_MultiSiteSyncSettings_Operation", "Set", objRemoteMSSOptions, "RevisionRuleOptions", aOptionsVals(1))
									If bFlag = False Then
										Exit For
									End If  
									wait 1
								Next
						Case "RevRuleOptionsList"   					'Select value from dropdown list in rev Rule Options list
								bFlag = Fn_SISW_UI_JavaList_Operations("Fn_CMS_MultiSiteSyncSettings_Operation","Select", objRemoteMSSOptions, "RevRuleOptionsList", sProperty, "", "")
						Case "GeneralOptions" 					'Select General Options
								aOptions = Split(sProperty,"~")
								For iCount = 0 To UBound(aOptions)
									aOptionsVals = Split(aOptions(iCount),":")
									objRemoteMSSOptions.JavaCheckBox("GeneralOptions").SetTOProperty "attached text",aOptionsVals(0)
									bFlag = Fn_SISW_UI_JavaCheckBox_Operations("Fn_CMS_MultiSiteSyncSettings_Operation", "Set", objRemoteMSSOptions, "GeneralOptions", aOptionsVals(1))
									If bFlag = False Then
										Exit For
									End If
									Wait 1									
								Next
						Case "TCXMLsessionOptions"   					'Select TC XML session Options
								bFlag = Fn_SISW_UI_JavaList_Operations("Fn_CMS_MultiSiteSyncSettings_Operation","Select", objRemoteMSSOptions, "TCXMLsessionoptions", sProperty, "", "")
						Case "SaveOptions"   					'Select Save Options
								aOptionsVals = Split(sProperty,":")
								objRemoteMSSOptions.JavaCheckBox("SaveOptions").SetTOProperty "attached text",aOptionsVals(0)
								bFlag = Fn_SISW_UI_JavaCheckBox_Operations("Fn_CMS_MultiSiteSyncSettings_Operation", "Set", objRemoteMSSOptions, "SaveOptions", aOptionsVals(1))	
					   	Case "AssemblyOptions"   					'Select Assembly Options
								aOptions = Split(sProperty,"~")
								For iCount = 0 To UBound(aOptions)
									aOptionsVals = Split(aOptions(iCount),":")
									objRemoteMSSOptions.JavaCheckBox("AssemblyOptions").SetTOProperty "attached text",aOptionsVals(0)
									bFlag = Fn_SISW_UI_JavaCheckBox_Operations("Fn_CMS_MultiSiteSyncSettings_Operation", "Set", objRemoteMSSOptions, "AssemblyOptions", aOptionsVals(1))
									If bFlag = False Then
										Exit For
									End If
									Wait 1									
								Next
						Case "AssemblyOptionsList"   					'Select value from dropdown list in Assembly Options
								bFlag = Fn_SISW_UI_JavaList_Operations("Fn_CMS_MultiSiteSyncSettings_Operation","Select", objRemoteMSSOptions, "AssemblyOptionsList", sProperty, "", "")		
						Case "SessionOptions"   					'Select Session Options
								aOptions = Split(sProperty,"~")
								For iCount = 0 To UBound(aOptions)
									aOptionsVals = Split(aOptions(iCount),":")
									objRemoteMSSOptions.JavaCheckBox("SessionOptions").SetTOProperty "attached text",aOptionsVals(0)
									bFlag = Fn_SISW_UI_JavaCheckBox_Operations("Fn_CMS_MultiSiteSyncSettings_Operation", "Set", objRemoteMSSOptions, "SessionOptions", aOptionsVals(1))
									If bFlag = False Then
										Exit For
									End If  
								Next						
						Case "AddIncludeRelation"   					'Add relations from Exclude Reference list to Include Reference list
								aOptions = Split(sProperty,"~")
								For iCount = 0 To UBound(aOptions)
									bFlag = Fn_SISW_UI_JavaList_Operations("Fn_CMS_MultiSiteSyncSettings_Operation","Select", objRemoteMSSOptions, "Exclude Reference", aOptions(iCount),"", "")
									If bFlag = False Then
										Exit For
									End If
									'Click button
									Call Fn_Button_Click("Fn_CMS_MultiSiteSyncSettings_Operation",objRemoteMSSOptions,"RelationAdd")
								Next										
						Case "AddExcludeRelation"   					'Add relations from Include Reference list to Exclude Reference list
								aOptions = Split(sProperty,"~")
								For iCount = 0 To UBound(aOptions)
									bFlag = Fn_SISW_UI_JavaList_Operations("Fn_CMS_MultiSiteSyncSettings_Operation","Select", objRemoteMSSOptions, "Include Reference", aOptions(iCount),"", "")
									If bFlag = False Then
										Exit For
									End If
									'Click button
									Call Fn_Button_Click("Fn_CMS_MultiSiteSyncSettings_Operation",objRemoteMSSOptions,"RelationRemove")
								Next						
					End Select
					
					If bFlag = False Then
						Fn_CMS_MultiSiteSyncSettings_Operation = False
						Set objRemoteMSSOptions = Nothing
						Call Fn_WriteLogFile("","FAIL : Function [Fn_CMS_MultiSiteSyncSettings_Operation] Failed to Perform Case ["&sAction&"] SubCase ["+sSubAction+"].")
						Exit Function
					End If
				Next				
		End Select
		
		If sButton<>"" Then
			Call Fn_Button_Click("Fn_CMS_MultiSiteSyncSettings_Operation",objRemoteMSSOptions,sButton)
			Wait 1
		End If
		
		Fn_CMS_MultiSiteSyncSettings_Operation = True
		Set objRemoteMSSOptions = Nothing
End Function
'=====================================================================================================================================================================================
'@@    Function Name			:	Fn_CMS_Delete_UserLevel_Preference
'@@
'@@    Description				:	Function used to delete User level Preference
'@@
'@@    Parameters			   	:	1. sUserName	: User Name 
'@@								:	2. sSiteName	: Site Name
'@@								:    3. sPreferencename	: Preference Name
'@@								:   3. sScope 		:  Scope Name 						 	         
'@@								:   4. sReserve		: for future use
'@@
'@@    Examples					:	Call Fn_CMS_Delete_UserLevel_Preference(sSite1,Environment.Value("TcUser6"),"Site2",Environment.Value("TcUser6"),"TC_use_tcxml_multisite","User","")
'@@
'@@    History					:	
'@@ ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@	  Developer Name			Date			 Rev. No.	   				Changes Done								 			Reviewer
'@@ ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@	  Jotiba Takkekar	 	21-June-2018	 		 1.0			Created - Added for MultiSite new TC's development			TC11.5(20180402.00)_CMS_NewDevelopment_JotibaT_21June2018
'=====================================================================================================================================================================================
Public Function Fn_CMS_Delete_UserLevel_Preference(sSiteName1,sUserName1,sSiteName2,sUserName2,sPreferencename,sReserve )
	GBL_FAILED_FUNCTION_NAME="Fn_CMS_Delete_UserLeve_Preference"
	Dim bReturn,sSOAUser
	
	'To kill existing process if running 
	Call Fn_KillProcess("")	
	
	sPrefName_Reset = sPreferencename
	sPreVal_Reset =  "true"
	sScope_Reset = "Site"
	sSOAUser = Split(Environment.Value("TcUserDBA"),":")(0)
	
	' Set site1
	If sSiteName1<>"" Then
		bReturn=Fn_SISW_CMS_SiteOperations("Set",sSiteName1,"")
		If bReturn=False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_CMS_Delete_UserLevel_Preference ]  Failed to set site - [ "&sSiteName1&" ]") 
		End If
		'==================== Added Code te reset Preference as per design chnage ==========================================
		 Call Fn_SOA_SetPreference(sSOAUser,sPrefName_Reset,sPreVal_Reset,sScope_Reset)
		'====================================================================================================================
		
		' Login to Teamcenter for Site1
'		If sUserName1<>"" Then
'			bSiteReset = False
'			bReturn = Fn_ReUserTcSession(True,True,sUserName1)
'			bSiteReset = True
'			If bReturn=False Then
'				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_CMS_Delete_UserLevel_Preference ] Failed to log into teamcenter session with user [ "+sUserName1+ " ]") 
'			End If
'			Call Fn_ReadyStatusSync(4)
'		End If
'		
'		' Delete User level Preference from site1
'		If sPreferencename<>""  Then
'			bReturn = Fn_SISW_Pref_PreferenceOperations("VerifyPreferenceWithScope",sPreferencename,"","User","","","","","","","","")
'			If bReturn = True Then
'				bReturn= Fn_SISW_Pref_PreferenceOperations("DeletePreferenceWithScope",sPreferencename,"","User","","","","","","","","")
'				If bReturn = False Then
'					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_CMS_Delete_UserLevel_Preference ] Failed to Delete Preference [ "+sPreferencename+ " ]") 
'				End If 
'				Call Fn_ReadyStatusSync(2)
'			End IF 
'		End If
'		'Call kill process
'		Call Fn_KillProcess("")
	End If
		
'	Set Site 2
	If sSiteName2<>"" Then
		bReturn=Fn_SISW_CMS_SiteOperations("Set",sSiteName2,"")
		If bReturn=False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_CMS_Delete_UserLevel_Preference ]  Failed to set site - [ "&sSiteName2&" ]") 
		End If
		'==================== Added Code te reset Preference as per design chnage ==========================================
		 Call Fn_SOA_SetPreference(sSOAUser,sPrefName_Reset,sPreVal_Reset,sScope_Reset)
		 If sSiteName1<>"" Then
		 	Call Fn_SISW_CMS_SiteOperations("Set",sSiteName1,"")	
		 End If
		'====================================================================================================================
		
		' Login to Teamcenter for Site2
'		If sUserName2<>"" Then
'			bSiteReset = False
'			bReturn = Fn_ReUserTcSession(True,True,sUserName2)
'			bSiteReset = True
'			If bReturn=False Then
'				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_CMS_Delete_UserLevel_Preference ] Failed to log into teamcenter session with user [ "+sUserName2+ " ]") 
'			End If
'			Call Fn_ReadyStatusSync(4)
'		End If
'		
'		' Delete User level Preference from site2
'		If sPreferencename<>""  Then
'			bReturn = Fn_SISW_Pref_PreferenceOperations("VerifyPreferenceWithScope",sPreferencename,"","User","","","","","","","","")
'			If bReturn = True Then
'				bReturn= Fn_SISW_Pref_PreferenceOperations("DeletePreferenceWithScope",sPreferencename,"","User","","","","","","","","")
'				If bReturn = False Then
'					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_CMS_Delete_UserLevel_Preference ] Failed to Delete Preference [ "+sPreferencename+ " ]") 
'				End If 
'				Call Fn_ReadyStatusSync(2)
'			End IF 
'			'Call kill process
'			Call Fn_KillProcess("")
'		End If
	End If
End Function
