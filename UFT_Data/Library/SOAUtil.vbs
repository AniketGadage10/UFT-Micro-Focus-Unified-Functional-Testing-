Option Explicit
Dim gPrefName, gPrefValue, gPrefScope
gPrefName= ""
gPrefValue = ""
gPrefScope=""

''=======================================================================================================================================================
'' Function List
''-------------------------------------------------------------------------------------------------------------------------------------------------------
''						Function Name											|					Created By
''-------------------------------------------------------------------------------------------------------------------------------------------------------
'' 1. Fn_CheckFileExists														| 	Samir Thosar (samir.thosar@siemens.com)
'' 2. Fn_SOA_GetClassPath														|	Samir Thosar (samir.thosar@siemens.com)
'' 3. Fn_SOA_VerifyFMSHome														|   Samir Thosar (samir.thosar@siemens.com)
'' 4. Fn_SOA_Execute															|	Samir Thosar (samir.thosar@siemens.com)
'' 5. Fn_SOA_CreateTCObject														|	Samir Thosar (samir.thosar@siemens.com)
'' 6. Fn_SOA_SetPreference														|   Samir Thosar (samir.thosar@siemens.com)
'' 7. Fn_SOA_PrefOperation														|   Samir Thosar (samir.thosar@siemens.com)
'' 8. Fn_SOA_ImportOperation													|   Samir Thosar (samir.thosar@siemens.com)
''=======================================================================================================================================================

Dim sAutoFolder, sSOAFolder
Dim sJavaHome, sFMSHome
Dim sQTPEnvXML, sSOAInputXML, sSOAOutputXML
Dim sSrvURL

''-------------------------------------------------------------------------------------
'' Initialize library level variables
''-------------------------------------------------------------------------------------
'' Autoamtion Directory Path
sAutoFolder = Fn_GetEnvValue("USER", "AutomationDir")

'' SOA Folder
sSOAFolder = sAutoFolder & "\SOA"

'' JAVA_HOME Path
sJavaHome = Fn_GetEnvValue("SYSTEM", "JAVA_HOME")

'' FMS_HOME Path
sFMSHome = Fn_GetEnvValue("USER", "FMS_HOME")

'' QTP Envirnoment XML
sQTPEnvXML = sAutoFolder & "\TestData\EnvVar_Ext.xml"

'' SOA Input XML
sSOAInputXML = sAutoFolder & "\SOA\soainput.xml"

'' SOA Output XML
sSOAOutputXML = sAutoFolder & "\SOA\soaoutput.xml"

'' Server URL
sSrvURL = Fn_GetXMLNodeValue(sQTPEnvXML, "TcWebServer")
If Instr(sSrvURL, "webclient") Then
	sSrvURL = Left(sSrvURL, (Len(sSrvURL) - Len("/webclient")))
End If

''----------------------------------------------------------------------------------------------------------------------

''----------------------------------------------------------------------------------------------------------------------
'' Function Number   	: 1                                                                              
'' Function Name     	: Fn_CheckFileExists(sFilePath)
'' Function Description : Function used to verify file existance
'' Function Usage    	: Result = Fn_CheckFileExists(sFilePath)
''							sFilePath	- File path
'' Function History
''----------------------------------------------------------------------------------------------------------------------
''	Developer Name		|	  Date		|Rev. No.|		    Changes Done			|	Reviewer	|	Reviewed Date
''----------------------------------------------------------------------------------------------------------------------
'' 	Samir Thosar		|  25-Nov-2010	| 	1.0	 |									|				|  
''----------------------------------------------------------------------------------------------------------------------
Public Function Fn_CheckFileExists(sFilePath)

	Dim objFSO

	Set objFSO = CreateObject("Scripting.FileSystemObject")
	If objFSO.FileExists(sFilePath) Then
		Fn_CheckFileExists = True
	Else
		Fn_CheckFileExists = False
	End If

	Set objFSO = Nothing

End Function
''----------------------------------------------------------------------------------------------------------------------
'' Function Number   	: 2
'' Function Name     	: Fn_SOA_GetClassPath(sFolderPath)
'' Function Description : Function used to generate list of jars present under SOA\lib folder to generate CLASSPATH
'' Function Usage    	: Result = Fn_SOA_GetClassPath(sFolderPath)
''							sFolderPath	- Folder path
'' Function History
''----------------------------------------------------------------------------------------------------------------------
''	Developer Name		|	  Date		|Rev. No.|		    Changes Done			|	Reviewer	|	Reviewed Date
''----------------------------------------------------------------------------------------------------------------------
'' 	Samir Thosar		|  25-Nov-2010	| 	1.0	 |									|				|  
''----------------------------------------------------------------------------------------------------------------------
Public Function Fn_SOA_GetClassPath(sFolderPath)

	Dim objFSO, objFolder, objFile
	Dim sSOALibFld
	Dim sClassPath
	Dim iFileCnt

	sSOALibFld = sSOAFolder & "\lib"
	sClassPath = ""

	Set objFSO = CreateObject("Scripting.FileSystemObject")

	If objFSO.FolderExists(sSOALibFld) Then
		Set objFolder = objFSO.GetFolder(sSOALibFld)
		Set objFile = objFolder.Files
		sClassPath = "set CLASSPATH=." & VbCrLf
        sClassPath = sClassPath & "set CLASSPATH=%CLASSPATH%;" & sSOALibFld & "\*"
		Fn_SOA_GetClassPath = sClassPath
	Else
		Fn_SOA_GetClassPath = False
	End If

	Set objFile = Nothing
	Set objFolder = Nothing
	Set objFSO = Nothing

End Function

''----------------------------------------------------------------------------------------------------------------------
'' Function Number   	: 3
'' Function Name     	: Fn_SOA_VerifyFMSHome(sFilePath)
'' Function Description : Verify the FMS_HOMEset in Env Varibale and FMS_HOME defined in soa.bat
'' Function Usage    	: Result = Fn_SOA_VerifyFMSHome(sFilePath)
''							sFilePath	- SOA batch file path
'' Function History
''----------------------------------------------------------------------------------------------------------------------
''	Developer Name		|	  Date		|Rev. No.|		    Changes Done			|	Reviewer	|	Reviewed Date
''----------------------------------------------------------------------------------------------------------------------
'' 	Samir Thosar		|  25-Nov-2010	| 	1.0	 |									|				|  
''----------------------------------------------------------------------------------------------------------------------
Public Function Fn_SOA_VerifyFMSHome(sFilePath)

	Dim objFSO, objFile
	Dim strData, arrLines, strLine
	
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	If objFSO.FileExists(sFilePath) Then
		Set objFile = objFSO.OpenTextFile(sFilePath,1, True)
		strData = objFile.ReadAll
		arrLines = Split(strData,vbCrLf, -1, 1)
		For strLine = 0 to Ubound(arrLines) - 1
			If Instr(arrLines(strLine), "FMS_HOME") Then
				If Fn_strUtil_SubField(arrLines(strLine),"=",1) <> sFMSHome then
					Fn_SOA_VerifyFMSHome = True
				Else
					Fn_SOA_VerifyFMSHome = False
				End If
				Exit For
			End If			
		Next
	Else
		Fn_SOA_VerifyFMSHome = False	
	End If
End Function

''----------------------------------------------------------------------------------------------------------------------
'' Function Number   	: 4
'' Function Name     	: Fn_SOA_Execute(sFolderPath)
'' Function Description : Function used to run the SOA
'' Function Usage    	: Result = Fn_SOA_Execute(sFolderPath)
''							sFolderPath	- Folder path
'' Function History
''----------------------------------------------------------------------------------------------------------------------
''	Developer Name		|	  Date		|Rev. No.|		    Changes Done					|	Reviewer	|	Reviewed Date
''----------------------------------------------------------------------------------------------------------------------
'' 	Samir Thosar		|  25-Nov-2010	| 	1.0	 |											|				|  
'' 	Samir Thosar		|  07-June-2011	|   1.1  |	Added %FMS_HOME%\jar\* in CLASSPATH				
''----------------------------------------------------------------------------------------------------------------------
Public Function Fn_SOA_Execute(sFolderPath)

	Dim objShell, objFSO, objFile
	Dim sNavFolder
    Dim sDriveLetter
	Dim bReturn
	Dim sBatFileName
	Dim sBatchLog
	Dim sClassPath
	Dim objWMIService, sTAOMgrPath, strWMIQuery, colProcess 


	
	If Instr(sSrvURL, "iiop") Then
		
		sTAOMgrPath = Fn_GetXMLNodeValue(sQTPEnvXML, "TAOManagerPath")
		SystemUtil.Run sTAOMgrPath		
		strWMIQuery = "Select * from Win32_Process where name like 'tao_imr_activator.exe'"
		Set objWMIService = GetObject("winmgmts:"& "{impersonationLevel=impersonate}!\\"& "." & "\root\cimv2") 		
		Do Until objWMIService.ExecQuery(strWMIQuery).Count > 0
			wait 2
		Loop
	
	End If

	sDriveLetter = Split(sFolderPath, ":", -1, 1)
	sNavFolder = "cd " & sFolderPath
	sBatFileName = "soa.bat"
	sBatchLog = "soa_out.log"

	If Fn_CheckFileExists(sFolderPath & "\" & sBatchLog) = True Then
		Set objFSO = CreateObject("Scripting.FileSystemObject")
		objFSO.DeleteFile(sFolderPath & "\" & sBatchLog)
		Set objFSO = Nothing
	End If
	' Added by Koustubh
	' Setting SOA result to False
	' Update SOA Action
	bReturn = Fn_UpdateEnvXMLNode(sSOAOutputXML, "SOAResult", "False")
	' Update SOA Data to empty String
	bReturn = Fn_UpdateEnvXMLNode(sSOAOutputXML, "SOAData", "SOA Utility - ApplicationManager is not executed.")

	If Fn_CheckFileExists(sFolderPath & "\soa.bat") = False OR Fn_SOA_VerifyFMSHome(sFolderPath & "\soa.bat") = True Then

		sClassPath = Fn_SOA_GetClassPath(sFolderPath)

		Set objFSO = CreateObject("Scripting.FileSystemObject")	
		Set objFile = objFSO.CreateTextFile(sFolderPath & "\" & sBatFileName, True)
		objFile.WriteLine ("Set JAVA_HOME=" & sJavaHome)
		objFile.WriteLine ("Set FMS_HOME=" & sFMSHome)
		objFile.WriteLine ("Set PATH=%JAVA_HOME%\bin;%FMS_HOME%\jar;%FMS_HOME%\lib;%PATH%")
		objFile.WriteLine (sClassPath) & ";" & sJavaHome & "\jre\lib\*;" & sFMSHome & "\jar\*" 
		objFile.WriteLine ("javac *.java")
		objFile.WriteLine ("java AutomationManager")
		objFile.Close	

		Set objShell = CreateObject("WScript.Shell")
		objShell.Run "%comspec% /c " & sDriveLetter(0) & ":" & "&" & sNavFolder & "&" & sBatFileName & ">> " & sBatchLog, 2, True
		Set objShell = Nothing

		Fn_SOA_Execute = True

		Set objFile = Nothing
		Set objFSO = Nothing

	Else

		Set objShell = CreateObject("WScript.Shell")
		objShell.Run "%comspec% /c " & sDriveLetter(0) & ":" & "&" & sNavFolder & "&" & sBatFileName & ">> " & sBatchLog, 2, True
		Set objShell = Nothing
		Fn_SOA_Execute = True

	End If

	If Instr(sSrvURL, "iiop") Then
	
		strWMIQuery = "Select * from Win32_Process"
		Set colProcess =  objWMIService.ExecQuery(strWMIQuery)
		For Each objProcess in colProcess
			If objProcess.Name = "tao_imr_activator.exe" OR objProcess.Name = "tao_imr_locator.exe" Then
				objProcess.Terminate()
			End If
		Next

		Set colProcess = Nothing
		Set objWMIService = Nothing

	End if


	If Fn_SOA_Execute = "" Then
		Fn_SOA_Execute = False
	End If
End Function

''----------------------------------------------------------------------------------------------------------------------
'' Function Number   	: 5
'' Function Name     	: Fn_SOA_CreateTCObject(aSOAInputData)
'' Function Description : Function used to create Teamcenter objects using SOA
'' Function Usage    	: Result = Fn_SOA_CreateTCObject(aSOAInputData)
''							aSOAInputData - Array of data required to use and create teamcenter objects
'' Function History
''----------------------------------------------------------------------------------------------------------------------
''	Developer Name		|	  Date		|Rev. No.|		    Changes Done			|	Reviewer	|	Reviewed Date
''----------------------------------------------------------------------------------------------------------------------
'' 	Samir Thosar		|  25-Nov-2010	| 	1.0	 |									|				|  
''----------------------------------------------------------------------------------------------------------------------
'' 	Koustubh Watwe		|  27-Sept-2011	| 	1.0	 |			Added case to create	|				|   
''															empty child folder
''----------------------------------------------------------------------------------------------------------------------
Public Function Fn_SOA_CreateTCObject(aSOAInputData())

	Dim bReturn

	' Update the server url
	bReturn = Fn_UpdateEnvXMLNode(sSOAInputXML, "AppURL", sSrvURL)

	' Update SOA Action	
    bReturn = Fn_UpdateEnvXMLNode(sSOAInputXML, "SOAAction", "Create")

	' Update username
	bReturn = Fn_UpdateEnvXMLNode(sSOAInputXML, "SOAUser", Fn_GetXMLNodeValue(sQTPEnvXML, aSOAInputData(0)))

	' Update Parent Folder Name
    bReturn = Fn_UpdateEnvXMLNode(sSOAInputXML, "ParentFolder", aSOAInputData(1))

	' Update Child Folder Name
	bReturn = Fn_UpdateEnvXMLNode(sSOAInputXML, "ChildFolder", aSOAInputData(2))

	' Update Object Type
    bReturn = Fn_UpdateEnvXMLNode(sSOAInputXML, "ObjectType", aSOAInputData(3))
	
	if aSOAInputData(3) <> "Folder" then
		' Update Object Sub Type
		bReturn = Fn_UpdateEnvXMLNode(sSOAInputXML, "ObjectData", aSOAInputData(4))

		' Update Number of items to be created
		bReturn = Fn_UpdateEnvXMLNode(sSOAInputXML, "ObjectCount", aSOAInputData(5))
	End IF
	' Create object using SOA
	bReturn = Fn_SOA_Execute(sSOAFolder)
	If bReturn = True Then
		bReturn = Fn_GetXMLNodeValue(sSOAOutputXML, "SOAResult")
		If LCase(bReturn) = LCase(true) Then
			Fn_SOA_CreateTCObject = True
		Else
			Fn_SOA_CreateTCObject = False
		End If		
	Else
		Fn_SOA_CreateTCObject = False
	End If

End Function

''----------------------------------------------------------------------------------------------------------------------
'' Function Number   	: 6
'' Function Name     	: Fn_SOA_SetPreference(sUserDetail, sPrefName, sPerfValue, sPerfScope)
'' Function Description : Function used to run the SOA
'' Function Usage    	: Result = Fn_SOA_SetPreference(sUserDetail, sPrefName, sPerfValue, sPerfScope)
''							sUserDetail	- Teamcenter User information 
''							sPrefName	- Prefernece Name
''							sPerfValue	- Desired Preference Value
''							sPerfScope	- Preference Scope (Site / Group / Role / User)
'' Function History
''----------------------------------------------------------------------------------------------------------------------
''	Developer Name		|	  Date		|Rev. No.|		    Changes Done			|	Reviewer	|	Reviewed Date
''----------------------------------------------------------------------------------------------------------------------
'' 	Samir Thosar		|  25-Nov-2010	| 	1.0	 |									|				|  
''----------------------------------------------------------------------------------------------------------------------

Public Function Fn_SOA_SetPreference(sUserDetail, sPrefName, sPerfValue, sPerfScope)

	Dim bReturn
   
   'Added calls for incorrectly setting PTN Default Revision rule preference for 4GD test cases
    If  sPrefName = "PTN_Default_Partition_RevRule" AND sPerfValue = "Any Status; No Working" Then
		Fn_SOA_SetPreference = True
		Exit Function
	End If
   
	' Update the server url
	bReturn = Fn_UpdateEnvXMLNode(sSOAInputXML, "AppURL", sSrvURL)

	' Update username
	bReturn = Fn_UpdateEnvXMLNode(sSOAInputXML, "SOAUser", Fn_GetXMLNodeValue(sQTPEnvXML, sUserDetail))

	' Update SOA Action
    bReturn = Fn_UpdateEnvXMLNode(sSOAInputXML, "SOAAction", "PREFERENCE")

	' Update Preference Action
    bReturn = Fn_UpdateEnvXMLNode(sSOAInputXML, "PREFACTION", "SETPREFERENCE")

	' Update Preference Name
    bReturn = Fn_UpdateEnvXMLNode(sSOAInputXML, "PERFNAME", sPrefName)

	' Update Preference Value
    bReturn = Fn_UpdateEnvXMLNode(sSOAInputXML, "PERFVALUE", sPerfValue)

	' Update Preference Scope
    bReturn = Fn_UpdateEnvXMLNode(sSOAInputXML, "PERFSCOPE", sPerfScope)

	' Set Preference using SOA
	bReturn = Fn_SOA_Execute(sSOAFolder)
	If bReturn = True Then
		bReturn = Fn_GetXMLNodeValue(sSOAOutputXML, "SOAResult")
		If LCase(bReturn) = LCase(true) Then
			Fn_SOA_SetPreference = True
		Else
			Fn_SOA_SetPreference = False
		End If		
	Else
		Fn_SOA_SetPreference = False
	End If

End Function

''----------------------------------------------------------------------------------------------------------------------
'' Function Number   	: 7
'' Function Name     	: Fn_SOA_PrefOperation(sSOAPerfData)
'' Function Description : Function used get / update and create preference
'' Function Usage    	: Result = Fn_SOA_SetPreference(sSOAPerfData)
''							sSOAPerfData - Array of data required to get / update and create preference
'' Function History
''----------------------------------------------------------------------------------------------------------------------
''	Developer Name		|	  Date		|Rev. No.|		    Changes Done					|	Reviewer	|	Reviewed Date
''----------------------------------------------------------------------------------------------------------------------
'' 	Samir Thosar		|  10-Dec-2010	| 	1.0	 |											|				|  
'' 	Koustubh Watwe		|  10-Jan-2011	| 	1.1	 |	Added case SetMultiValuePreference		|				|  
'' 	Koustubh Watwe		|  26-Jul-2011	| 	1.1	 |	Added case RemoveMultiValuePreference	|				|  
''  Sagar Shivade			22/11/2011					Modified case 'Createpreference' :- $ Seperator replaced by Pipe (|) 	
''----------------------------------------------------------------------------------------------------------------------


Public Function Fn_SOA_PrefOperation(sSOAPerfData())

   
	Dim bReturn
	Dim sAutoDir,sPrefrenceName,temp,Iterator

    ' Update the server url
	bReturn = Fn_UpdateEnvXMLNode(sSOAInputXML, "AppURL", sSrvURL)

	' Update username
	bReturn = Fn_UpdateEnvXMLNode(sSOAInputXML, "SOAUser", Fn_GetXMLNodeValue(sQTPEnvXML, sSOAPerfData(0)))

	' Update SOA Action
    bReturn = Fn_UpdateEnvXMLNode(sSOAInputXML, "SOAAction", "PREFERENCE")

	' Update Preference Action
	if lcase(sSOAPerfData(1)) = "create" then
		sSOAPerfData(1) = "CreatePreference"
	end if
    bReturn = Fn_UpdateEnvXMLNode(sSOAInputXML, "PREFACTION", sSOAPerfData(1))

	' Update Preference Name
	If sSOAPerfData(2) = "QS_TRUESHAPE_GENERATION_ENABLED" Then sSOAPerfData(2) = "QS_TRUSHAPE_GENERATION_ENABLED"
    bReturn = Fn_UpdateEnvXMLNode(sSOAInputXML, "PERFNAME", sSOAPerfData(2))
    
	If Lcase(sSOAPerfData(2)) = "psevariantsmode" OR Lcase(sSOAPerfData(2)) = "psevariantmode" Then
		sSOAPerfData(2) = "PSEVariantsMode"
		sSOAPerfData(3) = "User"
		Select Case Lcase(sSOAPerfData(4))
		Case "hybrid"
					sSOAPerfData(4) = "hybrid"
		Case "legacy"
					sSOAPerfData(4) = "legacy"
		Case "modular"
					sSOAPerfData(4) = "modular"
		End Select
	End If
	
	' Update Preference Scope
    bReturn = Fn_UpdateEnvXMLNode(sSOAInputXML, "PERFSCOPE", sSOAPerfData(3))

	' Update Preference Value
	'If (LCase(sSOAPerfData(1)) = LCase("SETPREFERENCE")) OR  (LCase(sSOAPerfData(1)) =LCase("SETMULTIVALUEPREFERENCE")) OR (LCase(sSOAPerfData(1)) = LCase("CREATEPREFERENCE")) Then

	'	 bReturn = Fn_UpdateEnvXMLNode(sSOAInputXML, "PERFVALUE", sSOAPerfData(4))

	'End If
	Select Case  LCase(sSOAPerfData(1))
		Case "setpreference", "setmultivaluepreference", "createpreference", "createmultivaluepreference", "addmultivaluepreference", "removemultivaluepreference","removepreference"
				
'---------------------------------------------------------------------------------------------------------
'Added by Chandrakant Tyagi to deal with the prefrence of Requirement Manager related prefrence 25-5-2015
			If sFeatureName="REG - RequirementsManagement" Then
				If LCase(sSOAPerfData(1))="createpreference" or LCase(sSOAPerfData(1))="createmultivaluepreference" or  LCase(sSOAPerfData(1))="setpreference" or  LCase(sSOAPerfData(1))="setmultivaluepreference" Then
					sAutoDir = Fn_GetEnvValue("User", "AutomationDir")
					sPrefrenceName = Fn_GetXMLNodeValue(sAutoDir + "\TestData\AutomationXML\CreatedPreferenceXML\RM_Created_Prefrence.xml", sSOAPerfData(2))
					If sPrefrenceName <> False Then
						If gPrefName = "" Then
							gPrefName =sPrefrenceName
							gPrefValue = sSOAPerfData(4)
							gPrefScope=	sSOAPerfData(3)
						Else
							temp=split(gPrefName,":")
							For Iterator = 0 To ubound(temp) Step 1
								If sPrefrenceName <> temp(Iterator) And Iterator = ubound(temp) Then
									gPrefName =gPrefName+":"+sPrefrenceName
									gPrefValue =gPrefValue+":"+sSOAPerfData(4)	
									gPrefScope =gPrefScope+":"+sSOAPerfData(3)	
								Else
									If sPrefrenceName = temp(Iterator) Then
										Exit for
									End If											
								End If
							Next
						End If
					Else
						'Do nothing					
					End If
				End If
			End If	
'-----------------------------------------------------------------------------------------------------------
				
				If sSOAPerfData(1) <> "removepreference" Then
					If InStr(1,sSOAPerfData(4),"$")<>0 Then
						sSOAPerfData(4)=Replace (sSOAPerfData(4),"$","|") 					' $ Seperator replaced by Pipe (|) 		By Sagar S. For 10.0 Porting 	22/11/12
					End If
					bReturn = Fn_UpdateEnvXMLNode(sSOAInputXML, "PERFVALUE", sSOAPerfData(4))
				End If		

	End Select
   
   'Added calls for incorrectly setting PTN Default Revision rule preference for 4GD test cases   

			If  sSOAPerfData(1) <> "removepreference" Then
			    If sSOAPerfData(2) = "PTN_Default_Partition_RevRule" and sSOAPerfData(4) = "Any Status; No Working" then
					Fn_SOA_PrefOperation = true
					Exit function
				End If
			End If
	' Set Preference using SOA
	bReturn = Fn_SOA_Execute(sSOAFolder)
	If bReturn = True Then
		bReturn = Fn_GetXMLNodeValue(sSOAOutputXML, "SOAResult")
		If LCase(bReturn) = LCase(true) Then
			Fn_SOA_PrefOperation = True
		Else
			Fn_SOA_PrefOperation = False
		End If		
	Else
		Fn_SOA_PrefOperation = False
	End If

End Function

''----------------------------------------------------------------------------------------------------------------------
'' Function Number   	: 8
'' Function Name     	: Fn_SOA_ImportOperation(sSOAPImpData)
'' Function Description : Function used to import workflow template
'' Function Usage    	: Result = Fn_SOA_SetPreference(sSOAPerfData)
''							sSOAPImpData - Array of data required to import workflow template
'' Function History
''----------------------------------------------------------------------------------------------------------------------
''	Developer Name		|	  Date		|Rev. No.|		    Changes Done			|	Reviewer	|	Reviewed Date
''----------------------------------------------------------------------------------------------------------------------
'' 	Samir Thosar		|  10-Dec-2010	| 	1.0	 |									|				|  
''----------------------------------------------------------------------------------------------------------------------

Public Function Fn_SOA_ImportOperation(sSOAInputData())

	Dim bReturn

    ' Update the server url
	bReturn = Fn_UpdateEnvXMLNode(sSOAInputXML, "AppURL", sSrvURL)

	' Update username
	bReturn = Fn_UpdateEnvXMLNode(sSOAInputXML, "SOAUser", Fn_GetXMLNodeValue(sQTPEnvXML, sSOAInputData(0)))

	' Update SOA Action
    bReturn = Fn_UpdateEnvXMLNode(sSOAInputXML, "SOAAction", "IMPORT")

	' Update Object Type
    bReturn = Fn_UpdateEnvXMLNode(sSOAInputXML, "ObjectType", sSOAInputData(1))

	' Update Object Sub Type
    bReturn = Fn_UpdateEnvXMLNode(sSOAInputXML, "ObjectData", sSOAInputData(2))

	' Create object using SOA
	bReturn = Fn_SOA_Execute(sSOAFolder)
	If bReturn = True Then
		bReturn = Fn_GetXMLNodeValue(sSOAOutputXML, "SOAResult")
		If LCase(bReturn) = LCase(true) Then
			Fn_SOA_ImportOperation = True
		Else
			Fn_SOA_ImportOperation = False
		End If		
	Else
		Fn_SOA_ImportOperation = False
	End If


End Function


''=======================================================================================================================================================
