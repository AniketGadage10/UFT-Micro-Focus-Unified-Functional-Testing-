Option Explicit

'=======================================================================================================================================================
' Function List
'-------------------------------------------------------------------------------------------------------------------------------------------------------
'						Function Name																		|					Created By
'-------------------------------------------------------------------------------------------------------------------------------------------------------
'1. Fn_LifecycleViewer_SaveSession()												|	Vallari S (vallari.shimpukade@siemens.com)
'=======================================================================================================================================================

'--------------------------------------------------------------------------------------------------------------------
' Function Number   	: 1                                                                                
' Function Name     	: Fn_LifecycleViewer_SaveSession
' Function Description  : Saves Session
' Function Usage    	: Result = Fn_LifecycleViewer_SaveSession(sAction, sStorageLocPref, SStorageLoc, sName, sCapture)
'							
'                     		return True/False
'--------------------------------------------------------------------------------------------------------------------
Public Function Fn_LifecycleViewer_SaveSession(sAction, sStorageLocPref, SStorageLoc, sName, sCapture, bUICheck)
   Dim objWin, sMenu, sMenuPath, bReturn

   Fn_LifecycleViewer_SaveSession = False
   Set objWin = JavaWindow("LifecycleViewerMainWin").Dialog("SessionSaveAs")

	If Not objWin.Exist(5) Then
		'Find File Path for Lifecycle Viewer Menu XML
        sMenuPath=Fn_LogUtil_GetXMLPath("LifecycleViewer_Menu")

		Select Case sAction
			Case "Save"
                sMenu = Fn_GetXMLNodeValue(sMenuPath, "SaveSession")
			Case "SaveAs"
				sMenu = Fn_GetXMLNodeValue(sMenuPath, "SaveSessionAs")
		End Select

		wait(2)
		bReturn = Fn_MenuOperation("WinMenuSelect", sMenu)
		wait(1)
        If bReturn = False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Operate Menu [" + sMenu + "]")
			Set objWin = Nothing
			Exit Function
		End If
	End If

	If Not objWin.Exist(5) Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Open Dialog [Save Session As]")
		Set objWin = Nothing
		Exit Function
	End If

	'UI Verification
	If CBool(bUICheck) = True Then
		bReturn = objWin.WinRadioButton("AlternateLocation").GetROProperty("checked")
		If Trim(cstr(bReturn)) <> "ON" Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Check Default Storage Location Preference as [Alternate Location]")
			objWin.WinButton("Cancel").Click 5, 5,micLeftBtn
			Set objWin = Nothing
			Exit Function
		End If
		bReturn = objWin.WinEdit("StorageLoc").GetROProperty("text")
		If Trim(cstr(bReturn)) <> "Newstuff" Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Check Default Storage Location as [Newstuff]")
			objWin.WinButton("Cancel").Click 5, 5,micLeftBtn
			Set objWin = Nothing
			Exit Function
		End If
	End If

	If sStorageLocPref <> "" Then
		Select Case sStorageLocPref
			Case "BaseDocument"
				objWin.WinRadioButton("AttachToBaseDocument").Set
			Case "Bomline"
				objWin.WinRadioButton("AttachToSelectedBomline").Set
			Case "Location"
				objWin.WinRadioButton("AlternateLocation").Set
				If SStorageLoc <> "" Then
					objWin.WinButton("Browse").Click 5, 5,micLeftBtn
					objWin.Dialog("Attach to").WinEdit("FolderName").Set SStorageLoc
					objWin.Dialog("Attach to").WinButton("Select").Click 5, 5,micLeftBtn
				End If
		End Select

		If Err.Number < 0 Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Set Storage Location as [" + sStorageLocPref + "]")
			objWin.WinButton("Cancel").Click 5, 5,micLeftBtn
			Set objWin = Nothing
			Exit Function
		End If
	End If

	If sName <> "" Then
		objWin.WinObject("SessionTree").DblClick 50, 30, micLeftBtn
		JavaWindow("LifecycleViewerMainWin").Dialog("ItemName").WinEdit("SessionName").Set sName
		JavaWindow("LifecycleViewerMainWin").Dialog("ItemName").WinButton("OK").Click 5, 5,micLeftBtn

		If Err.Number < 0 Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Set Storage Location Name as [" + sName + "]")
			objWin.WinButton("Cancel").Click 5, 5,micLeftBtn
			Set objWin = Nothing
			Exit Function
		End If
	End If

	If sCapture <> "" Then

	End If

	objWin.WinButton("Save").Click 5, 5,micLeftBtn
	wait(2)

	If Err.Number < 0 Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Save Session")
		objWin.WinButton("Cancel").Click 5, 5,micLeftBtn
		Set objWin = Nothing
		Exit Function
	End If

	Fn_LifecycleViewer_SaveSession = True
	Set objWin = Nothing

End Function
