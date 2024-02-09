Option Explicit
'------------------------'Global variables for Teamcenter Perspective Names--------------------------------------------------------------
Public GBL_PERSPECTIVE_ADA_LICENSE
GBL_PERSPECTIVE_ADA_LICENSE = "ADA License"
'------------------------'Global variables for Teamcenter Perspective Names--------------------------------------------------------------
'*********************************************************	Function List		***********************************************************************
'0. Fn_SISW_ADA_GetObject()
'1. Fn_ADALicense_CheckButtonsdisabled()
'2. Fn_ADALicense_TabSet()
'3. Fn_GetITARUser()
'4. Fn_ADALicense_Options()
'5. Fn_ADALicense_LicenseTreeOperations()
'6. Fn_ADALicense_GroupsTreeOperations()
'7. Fn_ADALicense_UsersTreeOperations()
'8. Fn_ADALicense_CheckControl_IsEnable()
'9. Fn_ADALicense_UsersGroups_Operations()
'10. Fn_ADALicense_AuditLogOperations()
'11. Fn_ADALicense_LicenseAccosiation()
'12. Fn_AuditLogTableRMB()
'13. Fn_ADALicense_PrintOptionVerify()
'*********************************************************	Function List		***********************************************************************

'- - - - - - - - - - - - - - - - - - - - - - - - - 			Function to get Object hierarchy  		- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
''Function Name		:	Fn_SISW_ADA_GetObject
'
''Description		  	 :  	Function to get Object hierarchy

''Parameters		   :	1. sObjectName : Object Handle name
								
''Return Value		   :  	Object \ Nothing
'
''Examples		     	:	Fn_SISW_ADA_GetObject("ADATree")

'History:                
'	Developer Name							Date				Rev. No.		Reviewer		Changes Done	
'- - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'	Amol Lanke		 				14-Sept-2012				
'-----------------------------------------------------------------------------------------------------------------------------------
'	Ashwini Kumar		 25-Sept-2013		2.0								Externalized object hierarchies
'-----------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_ADA_GetObject(sObjectName)
	Dim sObjectXMLPath
	sObjectXMLPath = Fn_GetEnvValue("User", "AutomationDir") & "\TestData\AutomationXML\ObjectXML\ADALicense.xml"
	Set Fn_SISW_ADA_GetObject = Fn_SISW_Setup_GetObjectFromXML(sObjectXMLPath, sObjectName)
End Function

'*********************************************************		Function to Check the buttons are enabled / disabled		**********************************************************************
'Function Name		:				Fn_ADALicense_CheckButtonsdisabled(sReferencePath,sButtons)

'Description			 :		 		 To Check the button is enabled or disabled.

'Parameters			   :	 			1) sReferencePath
'													 2) sButtons
											
'Return Value		   : 				TRUE \ FALSE

'Pre-requisite			:		 		ADALicense prespective should be displayed.

'Examples				:				Fn_ADALicense_CheckButtonsdisabled(JavaWindow("ADA License - Teamcenter"),"Create:Delete:Modify:Clear:ViewAuditLog")

'History:
'										Developer Name			Date			Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Ketan Raje					05-07-10			1.0	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function Fn_ADALicense_CheckButtonsdisabled(sReferencePath,sButtons)
	GBL_FAILED_FUNCTION_NAME="Fn_ADALicense_CheckButtonsdisabled"
   Dim aButtons,intCount,iCounter
		aButtons = split(sButtons, ":",-1,1)
		intCount = Ubound(aButtons)
		For iCounter=0 to intCount
				If Fn_UI_Object_GetROProperty("Fn_ADALicense_CheckButtonsdisabled",sReferencePath.JavaButton(aButtons(iCounter)), "enabled")="0" Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The "&aButtons(iCounter)&" button is disabled")
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The "&aButtons(iCounter)&" button is enabled")
					Fn_ADALicense_CheckButtonsdisabled = False
					Exit Function
				End If
		Next
		Fn_ADALicense_CheckButtonsdisabled = True
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The function Fn_ADALicense_CheckButtonsdisabled completed succesfully")
End Function

'*********************************************************		Function to select  the Tab into ADALicense***********************************************************************
'Function Name		:				Fn_ADALicense_TabSet(StrTabName)

'Description			 :		 		 This function is used to select the required Tab.

'Parameters			   :	 			1.  StrTabName:Name of the Tab to be selected.
											
'Return Value		   : 				TRUE \ FALSE

'Pre-requisite			:		 		ADALicense prespective should be displayed.

'Examples				:				 Fn_ADALicense_TabSet("Groups")

'History:
'										Developer Name			Date			Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Ketan Raje					05/07/2010																			
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function Fn_ADALicense_TabSet(StrTabName)
	GBL_FAILED_FUNCTION_NAME="Fn_ADALicense_TabSet"
	Dim objJavaWindow
	Set objJavaWindow = Fn_UI_ObjectCreate( "Fn_ADALicense_TabSet", JavaWindow("ADA License - Teamcenter"))
	   Select Case StrTabName
				'For selecting Groups Tab
			   Case "Groups" 				
						Call Fn_UI_JavaTab_Select("Fn_ADALicense_TabSet", objJavaWindow, "ACATab", "Groups")
						Fn_ADALicense_TabSet = TRUE				
			    'For selecting UsersTab
				Case "Users" 				
						Call Fn_UI_JavaTab_Select("Fn_ADALicense_TabSet", objJavaWindow, "ACATab", "Users")
						Fn_ADALicense_TabSet = TRUE								
				'Error message If the above Tab is not selected
				Case Else 
						 Fn_ADALicense_TabSet = FALSE
	   End Select
	   Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Select Tab " & StrTabName & " succeeded")
		Set objJavaWindow = Nothing 
End Function
'*********************************************************		Function to select  the Tab into ADALicense***********************************************************************
'Function Name		:				Fn_GetITARUser(FileLocation, UserType, iSheetNumber)

'Description			 :		 		 This function is used to fetch ITAR user details from the excel 
'Parameters			   :	 			1. FileLocation: Location of ITAR user excel file 
'                           2. UserType: User type for user details
'                               ADA Site Admin
'                               ITAR Admin 01
'                               ITAR Admin 02
'                               IP Admin 01
'                               IP Admin 02
'                               IP Admin 03
'                               IP Admin 04
'                               Test User 01
'                               Test User 02
'                               Test User 03
'                               Test User 04
'                               Test User 05
'                               Test User 06    
'                           3. iSheetNumber: Excel sheet number 
											
'Return Value		   : 				TRUE \ FALSE

'Pre-requisite			:		 		

'Examples				:				 Fn_GetITARUser(FileLocation, "ITAR Admin 01", 1)

'History:
'										Developer Name			Date			Rev. No.			Changes Done			Reviewer	
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Samir Thosar					05/07/2010																			
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_GetITARUser(FileLocation, UserType, iSheetNumber)
	GBL_FAILED_FUNCTION_NAME="Fn_GetITARUser"

	Const xlCellTypeLastCell = 11
	
	Dim objFSO, objFile ,iRowId
	Dim objExcel, objWorkbook, objWorksheet

	Set objFSO = CreateObject("Scripting.FileSystemObject")
	If objFSO.FileExists (FileLocation) Then
		Set objExcel = CreateObject("Excel.Application")
		Set objWorkbook = objExcel.Workbooks.Open(FileLocation)
		objExcel.Visible = False
		objExcel.DisplayAlerts = False
		If iSheetNumber = "" Then
			iSheetNumber = 1
		End If
		Set objWorksheet = objWorkbook.Worksheets(iSheetNumber)
		objWorksheet.Activate
    objWorksheet.UsedRange.SpecialCells(xlCellTypeLastCell).Activate
		For iRowId = 1 To objExcel.ActiveCell.Row
				If LCase(objExcel.Cells(iRowId, 1).Value) = LCase(UserType) Then
					Fn_GetITARUser = objExcel.Cells(iRowId, 2).Value
					Exit For
				End If
		Next

		objExcel.Quit

		Set objWorksheet = Nothing
		Set objWorkbook = Nothing
		Set objExcel = Nothing

		If Fn_GetITARUser = "" Then
			Fn_GetITARUser = False
		End If

	Else
		Fn_GetITARUser = False
	End If

End Function
'#########################################################################################################
'###
'###    FUNCTION NAME   :   Fn_ADALicense_Options(sAction,sLicenseID,sLicenseType,sLockDate,sLicenseExp,sReason,sAccordance)
'###
'###    DESCRIPTION        :   Create / Modify / Delete / Clear / ViewAuditLog all these functionality should be done for ADA Licenses.
'###
'###    PARAMETERS      :   1. sAction: Create / Modify / Delete / Clear / ViewAuditLog
'###											 2.	sLicenseID
'###											 3.	sLicenseType
'###											 4.	sLockDate
'###											 5.	sLicenseExp
'###											 6.	sReason
'###											 7.	sAccordance
'###                                         
'###    Function Calls       :   Fn_WriteLogFile(), Fn_Edit_Box(), Fn_List_Select(), Fn_UI_ObjectPressKey(), Fn_UI_ObjectCreate(), Fn_UI_SetDateAndTime(), Fn_Button_Click()
'###
'###	 HISTORY             :   AUTHOR                 DATE        VERSION
'###
'###    CREATED BY     :   Ketan Raje           05/07/2010         1.0
'###
'###    REVIWED BY     :   
'###
'###    MODIFIED BY   :  Pranav Shirode
'###   
'###    MODIFIED BY   :  Sanjeet K 			05-Feb-2013			1.1				Modified case : CreateWithCategory Added function call Call Fn_KeyBoardOperation("SendKeys", "{ENTER}")  after setting value in Category filed
'###
'###    EXAMPLE          : 		Case "Create"  : Call Fn_ADALicense_Options("Create","KR1","ITAR_License","06-Jul-2010 01:00:00","07-Jul-2011 01:00:00","Nothing","")
'###										 Case "Modify" : Call Fn_ADALicense_Options("Modify","","","","07-Jul-2011 01:00:00","No","")
'###										 Case "Delete" : Call Fn_ADALicense_Options("Delete","","","","","","")
'###										 Case "Clear" : Call Fn_ADALicense_Options("Clear","","","","","","")
'###										 Case "Verify" : Call Fn_ADALicense_Options("Verify","LID17428","ITAR_License","06-Jul-2010 17:04:30","07-Jul-2011 17:04:32","","")
'###										Case "CreateWithCategory"  : Call Fn_ADALicense_Options("CreateWithCategory","KR1:Category A","ITAR_License","06-Jul-2010 01:00:00","07-Jul-2011 01:00:00","Nothing","")
'###										Case "VerifyWithCategory"  : Call Fn_ADALicense_Options("VerifyWithCategory","KR1:Category A","ITAR_License","06-Jul-2010 01:00:00","07-Jul-2011 01:00:00","Nothing","")
'#############################################################################################################
Public Function Fn_ADALicense_Options(sAction,sLicenseID,sLicenseType,sLockDate,sLicenseExp,sReason,sAccordance)
	GBL_FAILED_FUNCTION_NAME="Fn_ADALicense_Options"
	Dim objLicense,objJavaEditLockDate,objJavaEditExpDate,objJavaWin,objJavaEditCreatedAfterDt,WshShell,iCnt,objAccordance
	Dim DateTime,objDateControl, iCount, bFlag
	bFlag=False
	'replaced _ with space, change is specific to TC12 -Added changes from TC11.4 by Jotiba T
	if instr(sLicenseType,"_") > 0 then
		sLicenseType = replace(sLicenseType,"_"," ")
	end if
	
	Set objLicense = Fn_UI_ObjectCreate("Fn_ADALicense_Options", JavaWindow("ADA License - Teamcenter"))
	Set objAccordance = Fn_SISW_ADA_GetObject("InAccordanceWith")
' Commented as mentioned in PR 7420643 - Koustubh [06-Jul-2015]
'	If sAction="Create" OR sAction="CreateWithCategory" Then
'		Call Fn_ADALicense_LicenseTreeOperations("Select","ADA Licenses","")
'	End If
		Select Case sAction
			Case "Create","Modify"
						If sAction = "Create" Then
							bGblFuncRetVal = Fn_ADALicense_CheckControl_IsEnable(objLicense.JavaButton("Clear"))
							If bGblFuncRetVal = true Then
								' click on clear
								bGblFuncRetVal = Fn_SISW_UI_JavaButton_Operations("Fn_ADALicense_Options", "Click", objLicense, "Clear")
								If bGblFuncRetVal = false Then
									Fn_ADALicense_Options = false
									Exit function
								End If
							End If
						End If
						'Set License ID
						If sLicenseID<>"" Then
							Call Fn_Edit_Box("Fn_ADALicense_Options",objLicense,"LicenseID",sLicenseID)
							wait 3
						End If
						'Set License Type
						If sLicenseType<>"" Then
							objLicense.JavaList("LicenseType").Click 0,0
							 iCount=objLicense.JavaList("LicenseType").GetROProperty("items count")
                             For iCnt = 0 To iCount-1
	                            If objLicense.JavaList("LicenseType").Object.GetItem(iCnt)=Trim(sLicenseType) Then
	                            	bFlag=True
	                            	Exit For 
	                            End If 	
                             Next
                             If bFlag=True Then
                             	'objLicense.JavaList("LicenseType").Object.select(iCnt)
								objLicense.JavaList("LicenseType").Select iCnt 
								wait 1
                             End If

'							Call Fn_List_Select("Fn_ADALicense_Options", objLicense, "LicenseType",sLicenseType)
'							wait 1
						End If
						'Set Lock Date
						DateTime = Split(sLockDate, " ", -1,1)
						If Trim(sLockDate) <> "" And Trim(sLockDate)<> "none" Then
							objLicense.JavaEdit("LockDate").Set DateTime(0)
							Wait 1
							Set WshShell = CreateObject("WScript.Shell")
							WshShell.SendKeys "{ESC}"
							Set WshShell = Nothing
                         	'[TC1121-20151012-23_10_2015-VivekA-Maintenance] - Added sync to enter time after date
							Call Fn_SyncTCObjects()
							If Ubound(DateTime) = 1 Then
								objLicense.JavaList("LockDate").Select ""
								Wait 1
								objLicense.JavaList("LockDate").Select DateTime(1)
							End If
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Date selected as " + sLockDate)
						ElseIf Trim(sLockDate) = "none" Then
							objLicense.JavaEdit("LockDate").Set " "
							Set WshShell = CreateObject("WScript.Shell")
							WshShell.SendKeys "{ESC}"
							Set WshShell = Nothing
							Wait 1
							If Ubound(DateTime) = 1 Then
								objLicense.JavaList("LockDate").Select " "
							End If
							Fn_ADALicense_Options = True
							Set objJavaWin = Nothing
							Set objJavaEditCreatedAfterDt = Nothing
							Set objDateControl = Nothing
              			End If 
						'Set License Expiry Date.
						DateTime = Split(sLicenseExp, " ", -1,1)
						If Trim(sLicenseExp) <> "" And Trim(sLicenseExp) <> "none" Then
							objLicense.JavaEdit("LicenseExpiry").Set DateTime(0)
							Wait 1
							Set WshShell = CreateObject("WScript.Shell")
							WshShell.SendKeys "{ESC}"
							Set WshShell = Nothing
							'[TC1121-20151012-23_10_2015-VivekA-Maintenance] - Added sync to enter time after date
							Call Fn_SyncTCObjects()
							If Ubound(DateTime) = 1 Then
								objLicense.JavaList("LicenseExpiry").Select ""
								Wait 1
								objLicense.JavaList("LicenseExpiry").Select DateTime(1)
							End If
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Date selected as " + sLicenseExp)
						ElseIf Trim(sLicenseExp) = "none" Then
							objLicense.JavaEdit("LicenseExpiry").Set " "
							Set WshShell = CreateObject("WScript.Shell")
							WshShell.SendKeys "{ESC}"
							Set WshShell = Nothing
							Wait 1
							If Ubound(DateTime) = 1 Then
								objLicense.JavaList("LicenseExpiry").Select " "
							End If
							Fn_ADALicense_Options = True
							Set objJavaWin = Nothing
							Set objJavaEditCreatedAfterDt = Nothing
							Set objDateControl = Nothing
     					End If
						'Set Reason
						If sReason<>"" Then
							Call Fn_Edit_Box("Fn_ADALicense_Options",objLicense,"Reason",sReason)
						End If
						'Set In Accordance With.
						If sAccordance<>"" Then
							' added wait statements for stability - Koustubh
							wait 5
							Call Fn_Edit_Box("Fn_ADALicense_Options", objLicense, "InAccordanceWith",sAccordance)
							wait 5
						End If
						'Click on Create / Modify button.
						Call Fn_Button_Click("Fn_ADALicense_Options", objLicense, sAction)
						Call Fn_ReadyStatusSync(1)
						Set objJavaEditLockDate = nothing 	
						Set objJavaEditExpDate = nothing

			Case "CreateWithCategory"
						'Set License ID
						If sLicenseID<>"" Then
							sLicenseID=split(sLicenseID,":")
							Call Fn_Edit_Box("Fn_ADALicense_Options",objLicense,"LicenseID",sLicenseID(0))
						End If
						'Set License Type
                        If sLicenseType<>"" Then
'							Call Fn_List_Select("Fn_ADALicense_Options", objLicense, "LicenseType",sLicenseType)
							bFlag=False
							iEelecount = objLicense.JavaList("LicenseType").GetROProperty("items count")
							For iCnt=0 to iEelecount-1
								If trim(objLicense.JavaList("LicenseType").Object.getItem(iCnt))=Trim(sLicenseType) Then
									objLicense.JavaList("LicenseType").Select iCnt
									bFlag=True
									Exit for
								End If
							Next
							If bFlag=false Then
								Set objLicense = nothing
								Fn_ADALicense_Options=false
								Exit function
							End If
						End If
						'Set Category
								Call Fn_Edit_Box("Fn_ADALicense_Options",objLicense,"Category",sLicenseID(1))
								Call Fn_KeyBoardOperation("SendKeys", "{ENTER}")
						'Set Lock Date
						DateTime = Split(sLockDate, " ", -1,1)
						If 	trim(sLockDate) <> "" and 	trim(sLockDate)<> "none" Then
							objLicense.JavaEdit("LockDate").Set DateTime(0)
							Set WshShell = CreateObject("WScript.Shell")
							WshShell.SendKeys "{ESC}"
							Set WshShell = Nothing
							wait 2
							If ubound(DateTime) = 1Then
								objLicense.JavaList("LockDate").Select DateTime(1)
							End If
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Date selected as " + sLockDate)
						Elseif trim(sLockDate) = "none" then
							objLicense.JavaEdit("LockDate").Set " "
							Set WshShell = CreateObject("WScript.Shell")
							WshShell.SendKeys "{ESC}"
							Set WshShell = Nothing
							wait 2
							If ubound(DateTime) = 1Then
								objLicense.JavaList("LockDate").Select " "
							End If
								Fn_ADALicense_Options = True
								Set objJavaWin = Nothing
								Set objJavaEditCreatedAfterDt = Nothing
								Set objDateControl = Nothing
                         End If 
						'Set License Expiry Date.
						DateTime = Split(sLicenseExp, " ", -1,1)
						If 	trim(sLicenseExp) <> "" And 	trim(sLicenseExp) <> "none"Then
							objLicense.JavaEdit("LicenseExpiry").Set DateTime(0)
							Set WshShell = CreateObject("WScript.Shell")
							WshShell.SendKeys "{ESC}"
							Set WshShell = Nothing
							wait 2
							If ubound(DateTime) = 1Then
								objLicense.JavaList("LicenseExpiry").Select DateTime(1)
							End If
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Date selected as " + sLicenseExp)
							Elseif trim(sLicenseExp) = "none" then
							objLicense.JavaEdit("LicenseExpiry").Set " "
							Set WshShell = CreateObject("WScript.Shell")
							WshShell.SendKeys "{ESC}"
							Set WshShell = Nothing
							wait 2
							If ubound(DateTime) = 1Then
								objLicense.JavaList("LicenseExpiry").Select " "
							End If
								Fn_ADALicense_Options = True
								Set objJavaWin = Nothing
								Set objJavaEditCreatedAfterDt = Nothing
								Set objDateControl = Nothing
     				End If
						'Set Reason
						If sReason<>"" Then
							Call Fn_Edit_Box("Fn_ADALicense_Options",objLicense,"Reason",sReason)
						End If
						'Set In Accordance With.
						If sAccordance<>"" Then							
							Call Fn_Edit_Box("Fn_ADALicense_Options", objLicense, "InAccordanceWith",sAccordance)
						End If
						'Click on Create / Modify button.
						Call Fn_Button_Click("Fn_ADALicense_Options", objLicense, "Create")
						Set objJavaEditLockDate = nothing 	
						Set objJavaEditExpDate = nothing

			Case "ModifyWithCategory"   'Added Case : By : Harshal Tanpure : 9-October-2012 : Build : Teamcenter 10 (20120919.00)
						'Set License ID
						If sLicenseID<>"" Then
							sLicenseID=split(sLicenseID,":")
							Call Fn_Edit_Box("Fn_ADALicense_Options",objLicense,"LicenseID",sLicenseID(0))
						End If
						'Set License Type
						If sLicenseType<>"" Then							
							Call Fn_List_Select("Fn_ADALicense_Options", objLicense, "LicenseType",sLicenseType)
						End If
						'Set Category
								Call Fn_Edit_Box("Fn_ADALicense_Options",objLicense,"Category",sLicenseID(1))
						'Set Lock Date
						DateTime = Split(sLockDate, " ", -1,1)
						If 	trim(sLockDate) <> "" and 	trim(sLockDate)<> "none" Then
							objLicense.JavaEdit("LockDate").Set DateTime(0)
							Set WshShell = CreateObject("WScript.Shell")
							WshShell.SendKeys "{ESC}"
							Set WshShell = Nothing
							wait 2
							If ubound(DateTime) = 1Then
								objLicense.JavaList("LockDate").Select DateTime(1)
							End If
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Date selected as " + sLockDate)
						Elseif trim(sLockDate) = "none" then
							objLicense.JavaEdit("LockDate").Set " "
							Set WshShell = CreateObject("WScript.Shell")
							WshShell.SendKeys "{ESC}"
							Set WshShell = Nothing
							wait 2
							If ubound(DateTime) = 1Then
								objLicense.JavaList("LockDate").Select " "
							End If
								Fn_ADALicense_Options = True
								Set objJavaWin = Nothing
								Set objJavaEditCreatedAfterDt = Nothing
								Set objDateControl = Nothing
                         End If 
						'Set License Expiry Date.
						DateTime = Split(sLicenseExp, " ", -1,1)
						If 	trim(sLicenseExp) <> "" And 	trim(sLicenseExp) <> "none"Then
							objLicense.JavaEdit("LicenseExpiry").Set DateTime(0)
							Set WshShell = CreateObject("WScript.Shell")
							WshShell.SendKeys "{ESC}"
							Set WshShell = Nothing
							wait 2
							If ubound(DateTime) = 1Then
								objLicense.JavaList("LicenseExpiry").Select DateTime(1)
							End If
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Date selected as " + sLicenseExp)
							Elseif trim(sLicenseExp) = "none" then
							objLicense.JavaEdit("LicenseExpiry").Set " "
							Set WshShell = CreateObject("WScript.Shell")
							WshShell.SendKeys "{ESC}"
							Set WshShell = Nothing
							wait 2
							If ubound(DateTime) = 1Then
								objLicense.JavaList("LicenseExpiry").Select " "
							End If
								Fn_ADALicense_Options = True
								Set objJavaWin = Nothing
								Set objJavaEditCreatedAfterDt = Nothing
								Set objDateControl = Nothing
     				End If
						'Set Reason
						If sReason<>"" Then
							Call Fn_Edit_Box("Fn_ADALicense_Options",objLicense,"Reason",sReason)
						End If
						'Set In Accordance With.
						If sAccordance<>"" Then							
							Call Fn_Edit_Box("Fn_ADALicense_Options", objLicense, "InAccordanceWith",sAccordance)
						End If
						'Click on Create / Modify button.
						Call Fn_Button_Click("Fn_ADALicense_Options", objLicense, "Modify")
						Set objJavaEditLockDate = nothing 	
						Set objJavaEditExpDate = nothing

			Case "Delete","Clear"
						'Click on Delete / Clear button
						Call Fn_Button_Click("Fn_ADALicense_Options", objLicense, sAction)
						If sAction="Delete" Then
							'Click on OK button 
							objLicense.Dialog("ErrorDialog").SetTOProperty "text","Delete License"
							'Click on OK button
							objLicense.Dialog("ErrorDialog").WinButton("OK").Click 1,1,micLeftBtn
						End If
			Case "Verify"
						'Check License ID
						If sLicenseID<>"" Then
							If Trim(Lcase(sLicenseID)) = Trim(Lcase(Fn_Edit_Box_GetValue("Fn_ADALicense_Options",objLicense,"LicenseID"))) Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "License ID value matches")
							Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "License ID value does not matches")
								Fn_ADALicense_Options = False
								Exit Function
							End If
						End If
						'Check License Type
						If sLicenseType<>"" Then							
							If Trim(Lcase(sLicenseType)) = Trim(Lcase(objLicense.JavaList("LicenseType").GetROProperty("value"))) Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "License Type value matches")
							Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "License Type value does not matches")
								Fn_ADALicense_Options = False
								Exit Function
							End If							
						End If
						'Check Lock Date
						If 	trim(sLockDate) <> "" Then
							sDate=Trim(Lcase(Fn_Edit_Box_GetValue("Fn_ADALicense_Options",objLicense,"LockDate"))) 

							sTime=Trim(Lcase(Fn_SISW_UI_JavaList_Operations("Fn_ADALicense_Options","GetText", objLicense, "LockDate", "", "", "")))

							sDateTime=sDate+Space(1)+sTime
							If Trim(Lcase(sLockDate)) = sDateTime Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Lock Date value matches")
							Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Lock Date value does not matches")
								Fn_ADALicense_Options = False
								Exit Function
							End If
						End If 
						'Check License Expiry Date.
						If 	trim(sLicenseExp) <> "" Then
							sDate=Trim(Lcase(Fn_Edit_Box_GetValue("Fn_ADALicense_Options",objLicense,"LicenseExpiry"))) 

							sTime=Trim(Lcase(Fn_SISW_UI_JavaList_Operations("Fn_ADALicense_Options","GetText", objLicense, "LicenseExpiry", "", "", "")))

							sDateTime=sDate+Space(1)+sTime
							If Trim(Lcase(sLicenseExp)) =sDateTime  Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "License Expiry Date value matches")
							Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "License Expiry Date value does not matches")
								Fn_ADALicense_Options = False
								Exit Function
							End If
						End If
						'Check Reason
						If sReason<>"" Then
							If Trim(Lcase(sReason)) = Trim(Lcase(Fn_Edit_Box_GetValue("Fn_ADALicense_Options",objLicense,"Reason"))) Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Reason value matches")
							Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Reason value does not matches")
								Fn_ADALicense_Options = False
								Exit Function
							End If
						End If
						'Check In Accordance With.
						If sAccordance<>"" Then							
							If Trim(Lcase(sAccordance)) = Trim(Lcase(objAccordance.GetROProperty("value"))) Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "In Accordance with value matches")
							Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "In Accordance with value does not matches")
								Fn_ADALicense_Options = False
								Exit Function
							End If							
						End If
						'Log to specify all values match correctly
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "All specified values match actual values")

			Case "VerifyWithCategory"
						sLicenseID=split(sLicenseID,":")
						'Check License ID
						If sLicenseID(0) <>"" Then
							If Trim(Lcase(sLicenseID(0))) = Trim(Lcase(Fn_Edit_Box_GetValue("Fn_ADALicense_Options",objLicense,"LicenseID"))) Then
								Call Fn_WriteLogFile (Environment.Value("TestLogFile"), "License ID value matches")
							Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "License ID value does not matches")
								Fn_ADALicense_Options = False
								Exit Function
							End If
						End If
						'Check License Type
						If sLicenseType<>"" Then							
							If Trim(Lcase(sLicenseType)) = Trim(Lcase(objLicense.JavaList("LicenseType").GetROProperty("value"))) Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "License Type value matches")
							Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "License Type value does not matches")
								Fn_ADALicense_Options = False
								Exit Function
							End If							
						End If
						'Check for Category
						If sLicenseID(1) <>"" Then
							If Trim(Lcase(sLicenseID(1))) = Trim(Lcase(Fn_Edit_Box_GetValue("Fn_ADALicense_Options",objLicense,"Category"))) Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Category value matches")
							Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Category value does not matches")
								Fn_ADALicense_Options = False
								Exit Function
							End If
						End If
						'Check Lock Date
						If 	trim(sLockDate) <> "" Then
							If Trim(Lcase(sLockDate)) = Trim(Lcase(Fn_Edit_Box_GetValue("Fn_ADALicense_Options",objLicense,"LockDate")) &" " & (objLicense.JavaList("LockDate").GetROProperty("value"))) Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Lock Date value matches")
							Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Lock Date value does not matches")
								Fn_ADALicense_Options = False
								Exit Function
							End If
						End If 
						'Check License Expiry Date.
						If 	trim(sLicenseExp) <> "" Then
							If Trim(Lcase(sLicenseExp)) = Trim(Lcase(Fn_Edit_Box_GetValue("Fn_ADALicense_Options",objLicense,"LicenseExpiry")) &" " & (objLicense.JavaList("LicenseExpiry").GetROProperty("value"))) Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "License Expiry Date value matches")
							Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "License Expiry Date value does not matches")
								Fn_ADALicense_Options = False
								Exit Function
							End If
						End If
						'Check Reason
						If sReason<>"" Then
							If Trim(Lcase(sReason)) = Trim(Lcase(Fn_Edit_Box_GetValue("Fn_ADALicense_Options",objLicense,"Reason"))) Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Reason value matches")
							Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Reason value does not matches")
								Fn_ADALicense_Options = False
								Exit Function
							End If
						End If
						'Check In Accordance With.
						If sAccordance<>"" Then							
							If Trim(Lcase(sAccordance)) = Trim(Lcase(objAccordance.GetROProperty("value"))) Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "In Accordance with value matches")
							Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "In Accordance with value does not matches")
								Fn_ADALicense_Options = False
								Exit Function
							End If							
						End If
						'Log to specify all values match correctly
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "All specified values match actual values")
						
						
			Case Else						
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_ADALicense_Options function failed")
						Fn_ADALicense_Options = FALSE
						Exit Function						
		End Select
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" Sucessfully completed of function Fn_ADALicense_Options")
    Set objLicense = nothing 	
	Fn_ADALicense_Options = TRUE
End Function
''*********************************************************		Function to action perform on ADALicense Tree	***********************************************************************
'Function Name		:				Fn_ADALicense_LicenseTreeOperations()

'Description			 :		 		 Actions performed in this function are:
'																	1. Node Select
'                                                                   2. Node Expand
'																	3. Node Collapse
'																	4. Node Exist

'Parameters			   :	 			1. sAction: Action to be performed
'													2. sNodeName: Fully qulified tree Path (delimiter as ':') 
'												   3. sLicenseType

'Return Value		   : 				TRUE \ FALSE

'Pre-requisite			:		 		ADALicense Prespective is Open.

'Examples				:				Case "Select" : Call Fn_ADALicense_LicenseTreeOperations("Select","ADA Licenses:Ketan1","ALL")
'													Case "Expand" : Call Fn_ADALicense_LicenseTreeOperations("Expand","ADA Licenses","")
'													Case "Collapse" : Call Fn_ADALicense_LicenseTreeOperations("Collapse","ADA Licenses","")
'													Case "Exist" : Call Fn_ADALicense_LicenseTreeOperations("Exist","ADA Licenses :Ketan1","")
'													Case "GetIndex" : Call Fn_ADALicense_LicenseTreeOperations("Exist","ADA Licenses :Ketan1","")
  												
'History					 :		
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'												------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Ketan Raje										   			06/07/2010			              1.0										Created	
'												------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function Fn_ADALicense_LicenseTreeOperations(sAction,sNodeName,sLicenseType)
	GBL_FAILED_FUNCTION_NAME="Fn_ADALicense_LicenseTreeOperations"
	Dim objJavaWindowADA, objJavaTreeADA, intNodeCount, intCount, sTreeItem
	Set objJavaWindowADA = Fn_UI_ObjectCreate( "Fn_ADALicense_LicenseTreeOperations",JavaWindow("ADA License - Teamcenter"))
	
	'objJavaWindowADA.JavaTree("ADALicenseTree").Activate "#0"
	'Select Item in the License Type List.
	
	If sLicenseType<>"" Then
	    sLicenseType = replace(sLicenseType,"_"," ")
		Call Fn_List_Select("Fn_ADALicense_LicenseTreeOperations", objJavaWindowADA, "LicenseList",sLicenseType)
	End If
	
	' Modified as per design change, as ADA License node is not present. 7-July-2015 - By Poonam
   	If Instr(1,Trim(sNodeName),"ADA Licenses:")>0  Then
		sNodeName = Replace(Trim(sNodeName),"ADA Licenses:","")
	ElseIf Instr(1,Trim(sNodeName),"ADA Licenses :")>0  Then
		sNodeName = Replace(Trim(sNodeName),"ADA Licenses :","")
	ElseIf StrComp(Trim(sNodeName),"ADA Licenses")=0 Then
 		 Fn_ADALicense_LicenseTreeOperations = TRUE
 		 Exit Function
 	End If
	
'	objJavaWindowADA.JavaTree("ADALicenseTree").Activate "#0"
	Select Case sAction
		'----------------------------------------------------------------------- For selecting single node -------------------------------------------------------------------------
		Case "Select"
					Call Fn_JavaTree_Select("Fn_ADALicense_LicenseTreeOperations", objJavaWindowADA, "ADALicenseTree",sNodeName)
                    Fn_ADALicense_LicenseTreeOperations = TRUE
		'----------------------------------------------------------------------- For expanding a particular  node-------------------------------------------------------------------------
		Case "Expand"
                   Call Fn_UI_JavaTree_Expand("Fn_ADALicense_LicenseTreeOperations",objJavaWindowADA,"ADALicenseTree",sNodeName)
				   Fn_ADALicense_LicenseTreeOperations = TRUE
		'----------------------------------------------------------------------- For collapssing a particular  node-------------------------------------------------------------------------
		Case "Collapse"
                    Call Fn_UI_JavaTree_Collapse("Fn_ADALicense_LicenseTreeOperations", objJavaWindowADA,"ADALicenseTree",sNodeName)
					Fn_ADALicense_LicenseTreeOperations = TRUE
		'----------------------------------------------------------------------- For Checking existance of a particular  node-------------------------------------------------------------------------
		Case "Exist"
				Set objJavaTreeADA = Fn_UI_ObjectCreate( "Fn_ADALicense_LicenseTreeOperations", objJavaWindowADA.JavaTree("ADALicenseTree"))
					intNodeCount = objJavaTreeADA.GetROProperty ("items count") 
					For intCount = 0 to intNodeCount - 1
						sTreeItem = objJavaTreeADA.GetItem(intCount)
						If Trim(lcase(sTreeItem)) = Trim(Lcase(sNodeName)) Then
							Fn_ADALicense_LicenseTreeOperations = TRUE	
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[" + sNodeName + "] of JavaTree exist")
							Exit For
						End If
					Next
					If Cstr(intCount) = intNodeCount Then
						Fn_ADALicense_LicenseTreeOperations = FALSE						
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[" + sNodeName + "] of JavaTree does not exist")
						Exit Function
					End If
		'----------------------------------------------------------------------- Get Index value of a particular node-------------------------------------------------------------------------
		Case "GetIndex"
				bFlag = False
				For intCount=0 to objJavaWindowADA.JavaTree("ADALicenseTree").GetROProperty ("items count")-1
					sTreeItem = objJavaWindowADA.JavaTree("ADALicenseTree").GetItem (intCount)
					If Trim(lcase(sTreeItem)) = Trim(Lcase(sNodeName)) Then
						Fn_ADALicense_LicenseTreeOperations = intCount
						bFlag = True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The index of the given node is "&intCount)
						Exit For
					End If
				Next
				If bFlag = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The given node does not exist")
					Fn_ADALicense_LicenseTreeOperations = FALSE
				End If
		'------------------------------------------		Case Else	-----------------------------------------
		Case Else
						Fn_ADALicense_LicenseTreeOperations = FALSE
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_ADALicense_LicenseTreeOperations function failed")
						Exit Function
	End Select
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" Sucessfully completed on Node [" + sNodeName + "] of JavaTree of function Fn_ADALicense_LicenseTreeOperations")
	Set objJavaWindowADA = nothing
	Set objJavaTreeADA = nothing
End Function
''*********************************************************		Function to action perform on ADALicense GroupsTree	***********************************************************************
'Function Name		:				Fn_ADALicense_GroupsTreeOperations()

'Description			 :		 		 Actions performed in this function are:
'																	1. Node Select
'                                                                   2. Node Expand
'																	3. Node Collapse
'																	4. Node Exist
'																	5. Node GetIndex

'Parameters			   :	 			1. sAction: Action to be performed
'													2. sNodeName: Fully qulified tree Path (delimiter as ':') 

'Return Value		   : 				TRUE \ FALSE

'Pre-requisite			:		 		ADALicense Prespective is Open.

'Examples				:				Case "Select" : Call Fn_ADALicense_GroupsTreeOperations("Select","Engineering")
'													Case "Multiselect" : Call Fn_ADALicense_GroupsTreeOperations("Multiselect","Engineering,ITAR_Admin")
'													Case "Expand" : Call Fn_ADALicense_GroupsTreeOperations("Expand","Engineering")
'													Case "Collapse" : Call Fn_ADALicense_GroupsTreeOperations("Collapse","Engineering")
'													Case "Exist" : Call Fn_ADALicense_GroupsTreeOperations("Exist","Engineering:Designer")
'													Case "GetIndex" : Call Fn_ADALicense_GroupsTreeOperations("GetIndex","Engineering:Designer")
  												
'History					 :		
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'												------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Ketan Raje										   			06/07/2010			              1.0										Created	
'												------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function Fn_ADALicense_GroupsTreeOperations(sAction,sNodeName)
	GBL_FAILED_FUNCTION_NAME="Fn_ADALicense_GroupsTreeOperations"
	Dim objJavaWindowADA, objJavaTreeADA, intNodeCount, intCount, sTreeItem
	Set objJavaWindowADA = Fn_UI_ObjectCreate( "Fn_ADALicense_GroupsTreeOperations",JavaWindow("ADA License - Teamcenter"))
	Select Case sAction
		'----------------------------------------------------------------------- For selecting single node -------------------------------------------------------------------------
		Case "Select"
                    Call Fn_JavaTree_Select("Fn_ADALicense_GroupsTreeOperations", objJavaWindowADA, "AvailableGroupsTree",sNodeName)
					Fn_ADALicense_GroupsTreeOperations = TRUE
		'----------------------------------------------------------------------- For selecting multiple node at a time -------------------------------------------------------------------------
		Case "Multiselect"
                    Call Fn_UI_JavaTree_ExtendSelect("Fn_ADALicense_GroupsTreeOperations",objJavaWindowADA,"AvailableGroupsTree", sNodeName)
					Fn_ADALicense_GroupsTreeOperations = TRUE
		'----------------------------------------------------------------------- For expanding a particular  node-------------------------------------------------------------------------
		Case "Expand"
                    Call Fn_UI_JavaTree_Expand("Fn_ADALicense_GroupsTreeOperations",objJavaWindowADA,"AvailableGroupsTree",sNodeName)
					Fn_ADALicense_GroupsTreeOperations = TRUE
		'----------------------------------------------------------------------- For expanding a particular  node by double clicking on it-------------------------------------------------------------------------
		Case "Activate"
                    objJavaWindowADA.JavaTree("AvailableGroupsTree").Activate sNodeName
					Fn_ADALicense_GroupsTreeOperations = TRUE
		'----------------------------------------------------------------------- For collapssing a particular  node-------------------------------------------------------------------------
		Case "Collapse"
                    Call Fn_UI_JavaTree_Collapse("Fn_ADALicense_GroupsTreeOperations", objJavaWindowADA,"AvailableGroupsTree",sNodeName)
					Fn_ADALicense_GroupsTreeOperations = TRUE
		'----------------------------------------------------------------------- For Checking existance of a particular  node-------------------------------------------------------------------------
		Case "Exist"
				Set objJavaTreeADA = Fn_UI_ObjectCreate( "Fn_ADALicense_GroupsTreeOperations", objJavaWindowADA.JavaTree("AvailableGroupsTree"))
					intNodeCount = objJavaTreeADA.GetROProperty ("items count") 
					For intCount = 0 to intNodeCount - 1
						sTreeItem = objJavaTreeADA.GetItem(intCount)
						If Trim(lcase(sTreeItem)) = Trim(Lcase(sNodeName)) Then
							Fn_ADALicense_GroupsTreeOperations = TRUE	
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[" + sNodeName + "] of JavaTree exist")
							Exit For
						End If
					Next
					If Cstr(intCount) = intNodeCount Then
						Fn_ADALicense_GroupsTreeOperations = FALSE						
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[" + sNodeName + "] of JavaTree does not exist")
						Exit Function
					End If
		'----------------------------------------------------------------------- Get Index value of a particular node-------------------------------------------------------------------------
		Case "GetIndex"
				bFlag = False
				For intCount=0 to objJavaWindowADA.JavaTree("AvailableGroupsTree").GetROProperty ("items count")-1
					sTreeItem = objJavaWindowADA.JavaTree("AvailableGroupsTree").GetItem (intCount)
					If Trim(lcase(sTreeItem)) = Trim(Lcase(sNodeName)) Then
						Fn_ADALicense_GroupsTreeOperations = intCount
						bFlag = True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The index of the given node is "&intCount)
						Exit For
					End If
				Next
				If bFlag = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The given node does not exist")
					Fn_ADALicense_GroupsTreeOperations = FALSE
				End If
		'----------------------------------------------------------------------- Case Else-------------------------------------------------------------------------
		Case Else
						Fn_ADALicense_GroupsTreeOperations = FALSE
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_ADALicense_GroupsTreeOperations function failed")
						Exit Function
	End Select
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" Sucessfully completed on Node [" + sNodeName + "] of JavaTree of function Fn_ADALicense_GroupsTreeOperations")
	Set objJavaWindowADA = nothing
	Set objJavaTreeADA = nothing
End Function
''*********************************************************		Function to action perform on ADALicense UsersTree	***********************************************************************
'Function Name		:				Fn_ADALicense_UsersTreeOperations()

'Description			 :		 		 Actions performed in this function are:
'																	1. Node Select
'                                                                   2. Node Expand
'																	3. Node Collapse
'																	4. Node Exist
'																	5. Node GetIndex

'Parameters			   :	 			1. sAction: Action to be performed
'													2. sNodeName: Fully qulified tree Path (delimiter as ':') 

'Return Value		   : 				TRUE \ FALSE

'Pre-requisite			:		 		ADALicense Prespective is Open.

'Examples				:				Case "Select" : Call Fn_ADALicense_UsersTreeOperations("Select","Engineering:Designer")
'													Case "Multiselect" : Call Fn_ADALicense_UsersTreeOperations("Multiselect","Engineering:Designer:AutoTest2 (autotest2),Engineering:Designer:AutoTest3 (autotest3)")
'													Case "Expand" : Call Fn_ADALicense_UsersTreeOperations("Expand","Engineering:Designer")
'													Case "Collapse" : Call Fn_ADALicense_UsersTreeOperations("Collapse","Engineering:Designer")
'													Case "Exist" : Call Fn_ADALicense_UsersTreeOperations("Exist","Engineering:Designer")
'													Case "GetIndex" : Call Fn_ADALicense_UsersTreeOperations("GetIndex","Engineering:Designer")
  												
'History					 :		
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'												------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Ketan Raje										   			06/07/2010			              1.0										Created	
'												------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function Fn_ADALicense_UsersTreeOperations(sAction,sNodeName)
	GBL_FAILED_FUNCTION_NAME="Fn_ADALicense_UsersTreeOperations"
	Dim objJavaWindowADA, objJavaTreeADA, intNodeCount, intCount, sTreeItem
	Set objJavaWindowADA = Fn_UI_ObjectCreate( "Fn_ADALicense_UsersTreeOperations",JavaWindow("ADA License - Teamcenter"))
	Select Case sAction
		'----------------------------------------------------------------------- For selecting single node -------------------------------------------------------------------------
		Case "Select"
                    Call Fn_JavaTree_Select("Fn_ADALicense_UsersTreeOperations", objJavaWindowADA, "AvailableUsersTree",sNodeName)
					Fn_ADALicense_UsersTreeOperations = TRUE
		'----------------------------------------------------------------------- For selecting multiple node at a time -------------------------------------------------------------------------
		Case "Multiselect"
                    Call Fn_UI_JavaTree_ExtendSelect("Fn_ADALicense_UsersTreeOperations",objJavaWindowADA,"AvailableUsersTree", sNodeName)
					Fn_ADALicense_UsersTreeOperations = TRUE
		'----------------------------------------------------------------------- For expanding a particular  node-------------------------------------------------------------------------
		Case "Expand"
                    Call Fn_UI_JavaTree_Expand("Fn_ADALicense_UsersTreeOperations",objJavaWindowADA,"AvailableUsersTree",sNodeName)
					Fn_ADALicense_UsersTreeOperations = TRUE
		'----------------------------------------------------------------------- For collapssing a particular  node-------------------------------------------------------------------------
		Case "Collapse"
                    Call Fn_UI_JavaTree_Collapse("Fn_ADALicense_UsersTreeOperations", objJavaWindowADA,"AvailableUsersTree",sNodeName)
					Fn_ADALicense_UsersTreeOperations = TRUE
		'----------------------------------------------------------------------- For Checking existance of a particular  node-------------------------------------------------------------------------
		Case "Exist"
				Set objJavaTreeADA = Fn_UI_ObjectCreate( "Fn_ADALicense_UsersTreeOperations", objJavaWindowADA.JavaTree("AvailableUsersTree"))
					intNodeCount = objJavaTreeADA.GetROProperty ("items count") 
					For intCount = 0 to intNodeCount - 1
						sTreeItem = objJavaTreeADA.GetItem(intCount)
						If Trim(lcase(sTreeItem)) = Trim(Lcase(sNodeName)) Then
							Fn_ADALicense_UsersTreeOperations = TRUE	
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[" + sNodeName + "] of JavaTree exist")
							Exit For
						End If
					Next
					If Cstr(intCount) = intNodeCount Then
						Fn_ADALicense_UsersTreeOperations = FALSE						
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[" + sNodeName + "] of JavaTree does not exist")
						Exit Function
					End If
		'----------------------------------------------------------------------- Get Index value of a particular node-------------------------------------------------------------------------
		Case "GetIndex"
				bFlag = False
				For intCount=0 to objJavaWindowADA.JavaTree("AvailableUsersTree").GetROProperty ("items count")-1
					sTreeItem = objJavaWindowADA.JavaTree("AvailableUsersTree").GetItem (intCount)
					If Trim(lcase(sTreeItem)) = Trim(Lcase(sNodeName)) Then
						Fn_ADALicense_UsersTreeOperations = intCount
						bFlag = True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The index of the given node is "&intCount)
						Exit For
					End If
				Next
				If bFlag = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The given node does not exist")
					Fn_ADALicense_LicenseTreeOperations = FALSE
				End If
		'----------------------------------------------------------------------- Case Else-------------------------------------------------------------------------
		Case Else
						Fn_ADALicense_UsersTreeOperations = FALSE
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_ADALicense_UsersTreeOperations function failed")
						Exit Function
	End Select
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" Sucessfully completed on Node [" + sNodeName + "] of JavaTree of function Fn_ADALicense_UsersTreeOperations")
	Set objJavaWindowADA = nothing
	Set objJavaTreeADA = nothing
End Function
'*********************************************************		Function to Check the Status of control, wheather its enabled / disabled		**********************************************************************
'Function Name		:				Fn_ADALicense_CheckControl_IsEnable(sReferencePath)

'Description			 :		 		 To Check the Control is enabled or disabled.

'Parameters			   :	 			1) sReferencePath
										
'Return Value		   : 				TRUE \ FALSE

'Pre-requisite			:		 		ADALicense prespective should be displayed.

'Examples				:				Call Fn_ADALicense_CheckControl_IsEnable(JavaWindow("ADA License - Teamcenter").JavaList("InAccordanceWith"))

'										Developer Name			Date			Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Ketan Raje					07-07-10			1.0	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function Fn_ADALicense_CheckControl_IsEnable(sReferencePath)
			GBL_FAILED_FUNCTION_NAME="Fn_ADALicense_CheckControl_IsEnable"
			If Fn_UI_Object_GetROProperty("Fn_ADALicense_CheckControl_IsEnable",sReferencePath, "enabled")="1" Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The control is enabled")
				Fn_ADALicense_CheckControl_IsEnable = True
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The control is disabled")
				Fn_ADALicense_CheckControl_IsEnable = False
			End If		
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The function Fn_ADALicense_CheckControl_IsEnable completed succesfully")
End Function
'#########################################################################################################
'###
'###    FUNCTION NAME   :   Fn_ADALicense_UsersGroups_Operations(sAction,sUserGroup)
'###
'###    DESCRIPTION        :   Add / Remove Users or Groups
'###
'###    PARAMETERS      :   1. sAction: Add / Remove
'###											 2.	sUsreGroup
'###                                         
'###    Function Calls       :   Fn_WriteLogFile(), Fn_UI_ObjectCreate(), Fn_UI_JavaTable_GetCellData(), Fn_UI_JavaTable_SelectRow(), Fn_Button_Click()
'###
'###	 HISTORY             :   AUTHOR                 DATE        VERSION
'###
'###    CREATED BY     :   Ketan Raje           08/07/2010         1.0
'###
'###    REVIWED BY     :   
'###
'###    MODIFIED BY   :  
'###
'###    EXAMPLE          : 		Case "AddUsers" : Call Fn_ADALicense_UsersGroups_Operations("AddUsers","Engineering:Designer:AutoTest2 (autotest2),Engineering:Designer:AutoTest3 (autotest3)")
'###										 Case "RemoveUsers" : Call Fn_ADALicense_UsersGroups_Operations("RemoveUsers","autotest2:autotest3")
'###										 Case "AddGroups" : Call Fn_ADALicense_UsersGroups_Operations("AddGroups","Engineering,ITAR_Admin")
'###										 Case "RemoveGroups" : Call Fn_ADALicense_UsersGroups_Operations("RemoveGroups","Engineering:ITAR_Admin")
'###										 Case "VerifyUsers" : Call Fn_ADALicense_UsersGroups_Operations("VerifyUsers","autotest3")
'###										 Case "VerifyGroups" : Call Fn_ADALicense_UsersGroups_Operations("VerifyGroups","ITAR_Admin")
'#############################################################################################################
Public Function Fn_ADALicense_UsersGroups_Operations(sAction,sUserGroup)
	GBL_FAILED_FUNCTION_NAME="Fn_ADALicense_UsersGroups_Operations"
	Dim ObjADA, iCounter, bReturn, aColname, iCount, iRowData
	Set ObjADA = Fn_UI_ObjectCreate("Fn_ADALicense_UsersGroups_Operations", JavaWindow("ADA License - Teamcenter"))
		Select Case sAction
				Case "AddUsers"
						If Instr(1,sUserGroup,",")<>0 Then
							'Call the milti-select  case of Users Tree
							Call Fn_ADALicense_UsersTreeOperations("Multiselect",sUserGroup)
						Else
							'Call the Select case of Users Tree
							Call Fn_ADALicense_UsersTreeOperations("Select",sUserGroup)
						End If
						'Click on Add button
						Call Fn_Button_Click("Fn_ADALicense_UsersGroups_Operations", ObjADA, "Add")						
				Case "RemoveUsers"
						JavaWindow("ADA License - Teamcenter").JavaTable("SelectedUsers").ActivateRow 0
						bReturn = JavaWindow("ADA License - Teamcenter").JavaTable("SelectedUsers").GetROProperty("rows")
						aColname = split(sUserGroup, ":",-1,1)
						iCount = Ubound(aColname)
						For iRowData=0 to iCount
							For iCounter=0 to Cint(bReturn)-1
								If Trim(lcase(Fn_UI_JavaTable_GetCellData("Fn_ADALicense_UsersGroups_Operations", ObjADA, "SelectedUsers",iCounter,0))) = Trim(lcase(aColname(iRowData))) then
									'Select Row of Users Table
									'Call Fn_UI_JavaTable_SelectRow("Fn_ADALicense_UsersGroups_Operations", ObjADA, "SelectedUsers",iCounter)
									JavaWindow("ADA License - Teamcenter").JavaTable("SelectedUsers").ActivateRow iCounter
									'Click on remove button
									Call Fn_Button_Click("Fn_ADALicense_UsersGroups_Operations", ObjADA, "Remove")	
								End If
									Exit For 								
							Next
						Next
				Case "AddGroups"
						If Instr(1,sUserGroup,",")<>0 Then
							'Call the milti-select  case of Groups Tree
							Call Fn_ADALicense_GroupsTreeOperations("Multiselect",sUserGroup)
						Else
							'Call the Select case of Users Tree
							Call Fn_ADALicense_GroupsTreeOperations("Select",sUserGroup)
						End If
						'Click on Add button
						Call Fn_Button_Click("Fn_ADALicense_UsersGroups_Operations", ObjADA, "Add")						
				Case "RemoveGroups"
						JavaWindow("ADA License - Teamcenter").JavaTable("SelectedGroups").SelectCell "0","0"
						bReturn = JavaWindow("ADA License - Teamcenter").JavaTable("SelectedGroups").GetROProperty("rows")
						aColname = split(sUserGroup, ":",-1,1)
						iCount = Ubound(aColname)
						For iRowData=0 to iCount
							For iCounter=0 to Cint(bReturn)-1
								If Trim(lcase(Fn_UI_JavaTable_GetCellData("Fn_ADALicense_UsersGroups_Operations", ObjADA, "SelectedGroups",iCounter,0))) = Trim(lcase(aColname(iRowData))) then
									'Select Row of Users Table
									'Call Fn_UI_JavaTable_SelectRow("Fn_ADALicense_UsersGroups_Operations", ObjADA, "SelectedGroups",iCounter)
									JavaWindow("ADA License - Teamcenter").JavaTable("SelectedGroups").SelectCell iCounter, "0"
									'Click on remove button
									Call Fn_Button_Click("Fn_ADALicense_UsersGroups_Operations", ObjADA, "Remove")	
								End If
									Exit For 								
							Next
						Next
			Case "VerifyUsers"
				bReturn = JavaWindow("ADA License - Teamcenter").JavaTable("SelectedUsers").GetROProperty("rows")
						If Cint(bReturn) <> 0 Then						
								'Extract the index of row at which the object exist.
								aColname = split(sUserGroup, ":",-1,1)
								iCount = Ubound(aColname)
								For iRowData=0 to iCount
									For iCounter=0 to Cint(bReturn)-1
										If Trim(lcase(Fn_UI_JavaTable_GetCellData("Fn_ADALicense_UsersGroups_Operations", ObjADA, "SelectedUsers",iCounter,0))) = Trim(lcase(aColname(iRowData))) then
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), aColname(iRowData) &" Sucessfully found in Selected Users Table")									
											Exit For 
										Elseif iCounter = Cint(bReturn)-1 Then
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), aColname(iRowData) &" not found in Selected Users Table")
											Fn_ADALicense_UsersGroups_Operations = FALSE
											Exit Function											
										End If
									Next
								Next
						Else
								Fn_ADALicense_UsersGroups_Operations = FALSE
								Exit Function											
						End If
			Case "VerifyGroups"
						bReturn = JavaWindow("ADA License - Teamcenter").JavaTable("SelectedGroups").GetROProperty("rows")
						If Cint(bReturn) <> 0 Then						
								'Extract the index of row at which the object exist.
								aColname = split(sUserGroup, ":",-1,1)
								iCount = Ubound(aColname)
								For iRowData=0 to iCount
									For iCounter=0 to Cint(bReturn)-1
										If Trim(lcase(Fn_UI_JavaTable_GetCellData("Fn_ADALicense_UsersGroups_Operations", ObjADA, "SelectedGroups",iCounter,0))) = Trim(lcase(aColname(iRowData))) then
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), aColname(iRowData) &" Sucessfully found in Selected Groups Table")									
											Exit For 
										Elseif iCounter = Cint(bReturn)-1 Then
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), aColname(iRowData) &" not found in Selected Groups Table")
											Fn_ADALicense_UsersGroups_Operations = FALSE
											Exit Function											
										End If
									Next
								Next
						Else
								Fn_ADALicense_UsersGroups_Operations = FALSE
								Exit Function											
						End If
			Case Else						
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_ADALicense_UsersGroups_Operations function failed")
						Fn_ADALicense_UsersGroups_Operations = FALSE
						Exit Function						
		End Select
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" Sucessfully completed of function Fn_ADALicense_UsersGroups_Operations")
	Fn_ADALicense_UsersGroups_Operations = True
    Set ObjADA = nothing 	
End Function
'*********************************************************		Function to Find / Verify the Audit Log Of License		**********************************************************************
'Function Name		:				Fn_ADALicense_AuditLogOperations(sAction,sObjID,sObjName,sObjRev,sObjTypeName,sObjSeqNum,sObjSecID,sObjSecName,sObjSecRev,sObjSecType,sObjSecSeqNum,sErrCode,sGroupName,sEventTypeName,sUserID,sDateCreBefore,sDateCreAfter,sClose)

'Description			 :		 		 To View the Audit Log of the Licenses

'Parameters			   :	 			sAction,sObjID,sObjName,sObjRev,sObjTypeName,sObjSeqNum,sObjSecID,sObjSecName,sObjSecRev,sObjSecType,sObjSecSeqNum,sErrCode,sGroupName,sEventTypeName,sUserID,sDateCreBefore,sDateCreAfter,sClose
											
'Return Value		   : 				TRUE \ FALSE

'Pre-requisite			:		 		ADALicense prespective should be displayed.

'Examples				:				Case "Find" : Call Fn_ADALicense_AuditLogOperations("Find","","samir123","","","","","","","","","","","Add Users/Groups","","","","no")
'													Case "Verify" : Call Fn_ADALicense_AuditLogOperations("Verify","UserData","thosars","","","","","","","","","","","","","","","no")
'													Case "GetList" : aRecord = Fn_ADALicense_AuditLogOperations("GetList","0","","","","","","","","","","","","","","","","no")
'													Case "Export" : Call Fn_ADALicense_AuditLogOperations("Export","all","excel","","","","","","","","","","","","","","","no")
'													Case "ExcelVerify" : Call Fn_ADALicense_AuditLogOperations("ExcelVerify","2",aRecord,"","","","","","","","","","","","","","","no")
'													Case "LoadAllVerify" : Call Fn_ADALicense_AuditLogOperations("LoadAllVerify","","","","","","","","","","","","","","","","","yes")
'													Case "VerifyObject" : Call Fn_ADALicense_AuditLogOperations("VerifyObject","PR-000001","ChgName","A","","","","","","","","","","","","","","")
'													Case "AdvancedFind" : Call Fn_ADALicense_AuditLogOperations("AdvancedFind","","Test","","EPMTask","","","","","Type","","","","Update Process","","","","")
'													Case "VerifyUser" : Call Fn_ADALicense_AuditLogOperations("VerifyUser","","","","","","","","","","","","","","AutoTest1 (autotest1)~cmuser01 (cmuser01)","","","yes")
'History:
'										Developer Name			Date			Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Ketan Raje					08-07-10			1.0	
'										Mahendra Bhandarkar	08-07-10			1.0	
'										Mahendra Bhandarkar	13-10-10
'										Amit T				29-08-11								Modified case Verify
'										Sandeep N				23-09-11								Modified case VerifyUser
'										Sanjeet K				15-Feb-2013							Modified case VerifyUser
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_ADALicense_AuditLogOperations(sAction,sObjID,sObjName,sObjRev,sObjTypeName,sObjSeqNum,sObjSecID,sObjSecName,sObjSecRev,sObjSecType,sObjSecSeqNum,sErrCode,sGroupName,sEventTypeName,sUserID,sDateCreBefore,sDateCreAfter,sClose)
	GBL_FAILED_FUNCTION_NAME="Fn_ADALicense_AuditLogOperations"
	Dim objLicense, DateTime, iRowCount, iRow, sReturn, iColCount, sCol, iCol, ObjExport, aTableRecord, objExcel, iCount, sExceldata, iReturn, iCounter, bFlag
	Dim objStaticTxt,arrUserID, sColumns, aColumns, sAllColumns
	ReDim aTableRecord (18)
    		
    		
    		If JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Audit Log").exist(3) = False Or JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Audit Log").exist(3) = False Then
    			JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").SetTOProperty "Index",0
    			JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").SetTOProperty "Index",0
    		End If

    	Select Case sAction
			Case "Find"
						'Click on View Audit Log button
'						If JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Audit Log").Exist = False Then
'							'Click on View Audit Log button
'							Call Fn_Button_Click("Fn_ADALicense_AuditLogOperations", JavaWindow("ADA License - Teamcenter"), "ViewAuditLog")	
'						End If
'						Set objLicense = Fn_UI_ObjectCreate("Fn_ADALicense_AuditLogOperations", JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Audit Log"))
						'Click on Clear button 
						'Call Fn_Button_Click("Fn_ADALicense_AuditLogOperations", objLicense, "Clear")

'						JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Audit Log").JavaButton("Clear").Click micLeftBtn
						'*Added by Nilesh on 27-Mar-2013
						'If (JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Audit Log").Exist OR JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Audit Log").Exist(5)) =False Then
						If Fn_SISW_UI_Object_Operations("Fn_ADALicense_AuditLogOperations","Exist", JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Audit Log"),SISW_DEFAULT_TIMEOUT) = False OR Fn_SISW_UI_Object_Operations("Fn_ADALicense_AuditLogOperations","Exist", JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Audit Log"),SISW_MIN_TIMEOUT) =False Then
							'Click on View Audit Log button
							Call Fn_Button_Click("Fn_ADALicense_AuditLogOperations", JavaWindow("ADA License - Teamcenter"), "ViewAuditLog")	
						End If
						Call Fn_ReadyStatusSync(1)
						'If  JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Audit Log").Exist = True Then
						If Fn_SISW_UI_Object_Operations("Fn_ADALicense_AuditLogOperations","Exist", JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Audit Log"),SISW_MICROLESS_TIMEOUT) = True Then
							Set objLicense = Fn_UI_ObjectCreate("Fn_ADALicense_AuditLogOperations", JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Audit Log"))
						Else
							Set objLicense = Fn_UI_ObjectCreate("Fn_ADALicense_AuditLogOperations",JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Audit Log"))
						End If
						objLicense.JavaButton("Clear").Click micLeftBtn
						'*End
						'Set Object ID
						If sObjID<>"" Then
							Call Fn_Edit_Box("Fn_ADALicense_AuditLogOperations",objLicense,"ObjectID",sObjID)
						End If
						'Set Object Name
						If sObjName<>"" Then
							Call Fn_Edit_Box("Fn_ADALicense_AuditLogOperations",objLicense,"ObjectName",sObjName)
						End If
						'Set Object Revision
						If sObjRev<>"" Then
							Call Fn_Edit_Box("Fn_ADALicense_AuditLogOperations",objLicense,"ObjectRevision",sObjRev)
						End If
						'Set Object Type Name
						If sObjTypeName<>"" Then
							Call Fn_Edit_Box("Fn_ADALicense_AuditLogOperations",objLicense,"ObjectTypName",sObjTypeName)
						End If
						'Set Object Sequence Number
						If sObjSeqNum<>"" Then
							Call Fn_Edit_Box("Fn_ADALicense_AuditLogOperations",objLicense,"ObjectSeqNo",sObjSeqNum)
						End If
						'Set Secondary Object ID
						If sObjSecID<>"" Then
							Call Fn_Edit_Box("Fn_ADALicense_AuditLogOperations",objLicense,"SecondaryObjID",sObjSecID)
						End If
						'Set Secondary Object Name
						If sObjSecName<>"" Then
							Call Fn_Edit_Box("Fn_ADALicense_AuditLogOperations",objLicense,"SecondaryObjName",sObjSecName)
						End If
						'Set Secondary Object Revision
						If sObjSecRev<>"" Then
							Call Fn_Edit_Box("Fn_ADALicense_AuditLogOperations",objLicense,"SecondaryObjRev",sObjSecRev)
						End If
						'Set Secondary Object Type
						If sObjSecType<>"" Then
							Call Fn_Edit_Box("Fn_ADALicense_AuditLogOperations",objLicense,"SecondaryObjType",sObjSecType)
						End If
						'Set Secondary Object Sequence Number
						If sObjSecSeqNum<>"" Then
							Call Fn_Edit_Box("Fn_ADALicense_AuditLogOperations",objLicense,"SecondaryObjSeqNo",sObjSecSeqNum)
						End If
						'Set Error Code
						If sErrCode<>"" Then
							Call Fn_Edit_Box("Fn_ADALicense_AuditLogOperations",objLicense,"ErrorCode",sErrCode)
						End If
						'Set Group Name
						If sGroupName<>"" Then
							Call Fn_Edit_Box("Fn_ADALicense_AuditLogOperations",objLicense,"GroupName",sGroupName)
						End If
						'Select Event Type Name
						If sEventTypeName<>"" Then
							Call Fn_List_Select("Fn_ADALicense_Options", objLicense, "EventTypeName",sEventTypeName)
						End If
						'Set User ID
						If sUserID<>"" Then
							Call Fn_Edit_Box("Fn_ADALicense_AuditLogOperations",objLicense,"UserID",sUserID)
							Wait 1
                            Call Fn_KeyBoardOperation("SendKeys", "{ENTER}")
						End If
						'Set Date Created Before
						If sDateCreBefore<>"" Then
							objLicense.JavaCheckBox("DateCreatedBefore").Object.setDate(sDateCreBefore)
						End If
						'Set Date Created After
						If sDateCreAfter<>"" Then
							objLicense.JavaCheckBox("DateCreatedAfter").Object.setDate(sDateCreAfter)
						End If
						'Click on Find button
						Call Fn_Button_Click("Fn_ADALicense_AuditLogOperations", objLicense, "Find")
						If Trim(Lcase(sClose)) = "yes" Then
							objLicense.Close 
						End If
			Case "Verify"
						'*Added by Nilesh on 27-Mar-2013
		                'If  JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Audit Log").Exist = True Then
		                If Fn_SISW_UI_Object_Operations("Fn_ADALicense_AuditLogOperations","Exist", JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Audit Log"),SISW_MINLESS_TIMEOUT) Then
							Set objLicense = Fn_UI_ObjectCreate("Fn_ADALicense_AuditLogOperations", JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Audit Log"))
						Else
							Set objLicense = Fn_UI_ObjectCreate("Fn_ADALicense_AuditLogOperations",JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Audit Log"))
						End If
						'*End

'						Set objLicense = Fn_UI_ObjectCreate("Fn_ADALicense_AuditLogOperations", JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Audit Log"))
						
						'[TC1123-20161031-10_11_2016-VivekA-Maintenance] - Added by Archana D
						If objLicense.javaButton("LoadAll").Exist(1) Then
							If objLicense.javaButton("LoadAll").GetROProperty("enabled") Then
								objLicense.JavaButton("LoadAll").Click
								Call Fn_ReadyStatusSync(2)
							End If
						End If
						'-------------------------------------------------
						
						'iRowCount = objLicense.JavaTable("LogTable").GetROProperty("rows")
						iRowCount = Fn_UI_Object_GetROProperty("Fn_ADALicense_AuditLogOperations",objLicense.JavaTable("LogTable"), "rows")
						If iRowCount < 1 Then
							Fn_ADALicense_AuditLogOperations = False
							Set objLicense = Nothing
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Audit Log Table does not exist")
							Exit Function
						End If
						sColumns = objLicense.JavaTable("LogTable").GetROProperty("column names")
						aColumns = Split(sColumns," ")
						For iCol = 0 to Ubound(aColumns)
							If iCol = 0 Then
								sAllColumns = aColumns(0)
							Else
								sAllColumns = sAllColumns & aColumns(iCol)
							End If
						Next
						aColumns = Split(sAllColumns,";")
						sObjID = Replace(sObjID,"Sec_ObjectId","SecondaryObjectID")
						For iCol = 0 to Ubound(aColumns)
							If Trim(Lcase(aColumns(iCol))) = Trim(Lcase(sObjID)) Then
								Exit For
							End If
						Next
						iRowCount = Fn_Table_GetRowCount("Fn_ADALicense_AuditLogOperations",objLicense, "LogTable")
						iRowCount=CInt(iRowCount)
						For iRow=0 to iRowCount-1
							sReturn = objLicense.JavaTable("LogTable").Object.GetValueAt(iRow,iCol).ToString()
							If IsNumeric(sReturn) Then
								If instr(1,Cstr(Cint(sReturn)),Cstr(Cint(sObjName)))<>0 Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sObjName &" Sucessfully found in row "&iRow)
									Exit For
								ElseIf sObjID <> "UserData" Then
									IF instr(1,Cstr(sReturn),Cstr(sObjName))<>0 Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sObjName &" Sucessfully found in row "&iRow)							
										Exit For
									End If
								End If
							Else
								If sObjID = "ObjectType" Then
									if Trim(Lcase(Cstr(sObjName))) = "ada_license" OR _
									   Trim(Lcase(Cstr(sObjName))) = "ip_license" OR _
									   Trim(Lcase(Cstr(sObjName))) = "exclude_license" OR _ 
									   Trim(Lcase(Cstr(sObjName))) = "itar_license" Then
									   
									   sObjName = replace(sObjName, "_", " ")
									End If
								End If
								If instr(1,Trim(Lcase(Cstr(sReturn))),Trim(Lcase(Cstr(sObjName)))) <> 0 Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sObjName &" Sucessfully found in row "&iRow)							
										Exit For
								ElseIf sObjID <> "UserData" Then
										IF instr(1,Cstr(sReturn),Cstr(sObjName))<>0 Then
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sObjName &" Sucessfully found in row "&iRow)
											Exit For
										End If
								End If
							End If
						Next
						'Return value of function.
						If iRow = iRowCount Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sObjName &" not found in the Log Table")								
								JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").SetTOProperty "Index",1
    							JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").SetTOProperty "Index",1	
								Fn_ADALicense_AuditLogOperations = False
								If Trim(Lcase(sClose)) = "yes" Then
									objLicense.Close 
								End If
								Set objLicense = nothing 		
								Exit Function
						End If
						'Close the Audit Log dialog.
						If Trim(Lcase(sClose)) = "yes" Then
							objLicense.Close 
						End If
						Set objLicense = nothing 		
			Case "Export"
						'Call Fn_KillProcess("EXCEL.EXE")
						Call Fn_WindowsApplications("TerminateAll", "EXCEL.EXE")
						

						JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Audit Log").JavaTable("LogTable").Object.selectAll
						wait 1

'						JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Audit Log").JavaButton("Export Audit Log").Object.doClick(1)
'						Call Fn_Button_Click( "Fn_ADALicense_AuditLogOperations" , JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Audit Log") , "Export Audit Log" )
						'[TC1122-20160420-03_05_2016-VivekA-Maintenance] - Added from TC1015 ---------------------------------------
						Call Fn_SISW_UI_JavaButton_Operations("Fn_ADALicense_AuditLogOperations", "Click", JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Audit Log").JavaButton("Export Audit Log"), "")
						'-----------------------------------------------------------------------------------------------------------
						Wait 1

						JavaWindow("DefaultWindow").JavaWindow("ExportAuditLog").Activate

						'Set ObjExport = Fn_UI_ObjectCreate("Fn_ADALicense_AuditLogOperations", JavaWindow("DefaultWindow").JavaWindow("ExportAuditLog"))
						Set ObjExport = JavaWindow("DefaultWindow").JavaWindow("ExportAuditLog")

						'Select Object Selection radio button
						If Trim(Lcase(sObjID))="all" Then
							Call Fn_UI_JavaRadioButton_SetON("Fn_ADALicense_AuditLogOperations",ObjExport, "ExportAllObjectsIn")
						Else
							Call Fn_UI_JavaRadioButton_SetON("Fn_ADALicense_AuditLogOperations",ObjExport, "ExportSelectedObjects")
						End If

						If Trim(Lcase(sObjName))="excel" Then
							Call Fn_UI_JavaRadioButton_SetON("Fn_ADALicense_AuditLogOperations",ObjExport, "UseExcel")
						ElseIf Trim(Lcase(sObjName))="csv" Then
							Call Fn_UI_JavaRadioButton_SetON("Fn_ADALicense_AuditLogOperations",ObjExport, "UseCSV")
						End If

						'Click on OK button	
'						Call Fn_Button_Click("Fn_ADALicense_AuditLogOperations", ObjExport, "OK")
						'[TC1122-20160420-03_05_2016-VivekA-Maintenance] - Added from TC1015 ---------------------------------------
						Call Fn_SISW_UI_JavaButton_Operations("Fn_ADALicense_AuditLogOperations", "DeviceReplay.Click",ObjExport, "OK")
						'-----------------------------------------------------------------------------------------------------------
						If ObjExport.exist(2) Then	''' through automation Export Audit log dialog is not closed. added code to close dialog if it exist
							Call Fn_Button_Click("Fn_ADALicense_AuditLogOperations", ObjExport, "Cancel")
						End If
						If Trim(Lcase(sClose)) = "yes" Then
							objLicense.Close 
						End If

			Case "GetList"
						Set objLicense = Fn_UI_ObjectCreate("Fn_ADALicense_AuditLogOperations", JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Audit Log"))
						iColCount = objLicense.JavaTable("LogTable").GetROProperty("cols")
						iRowCount = objLicense.JavaTable("LogTable").GetROProperty("rows")
						If iRowCount < sObjID Then
							Fn_ADALicense_AuditLogOperations = False
							Set objLicense = Nothing
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sObjID &" does not exist in Audit Log Table.")
							Exit Function
						End If
						For iCol=0 to iColCount-1
							aTableRecord(iCol) = objLicense.JavaTable("LogTable").Object.GetValueAt(sObjID,iCol).ToString()
						Next
						Fn_ADALicense_AuditLogOperations = aTableRecord
						Set objLicense = Nothing
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" Sucessfully completed of function Fn_ADALicense_AuditLogOperations")
						If Trim(Lcase(sClose)) = "yes" Then
							objLicense.Close 
						End If
						JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").SetTOProperty "Index",1
    					JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").SetTOProperty "Index",1	
						Exit Function
			Case "ExcelVerify"
						iCount = 0
						Wait(30)
						Set objLicense = Fn_UI_ObjectCreate("Fn_ADALicense_AuditLogOperations", JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Audit Log"))
						Set objExcel = Nothing
		                Set objExcel = GetObject(,"Excel.Application")	'This line gives syntax warning, but works successfully, please don't remove the comma.
						iColCount = objExcel.ActiveSheet.UsedRange.Columns.Count
							For iCol = 2 to iColCount
								For iCounter=iCount to objLicense.JavaTable("LogTable").GetROProperty("cols")-1
									If LCase(objExcel.Cells(sObjID, iCol).Value)="" Then
										sExceldata = " "
									Else
										sExceldata = LCase(objExcel.Cells(sObjID, iCol).Value)
									End If
									If sExceldata = LCase(sObjName(iCounter)) Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sObjName(iCounter) &" match successfully")
										iCount = iCount+1
										Exit For
									Else
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sObjName(iCounter) &" does not match")
										JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").SetTOProperty "Index",1
    									JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").SetTOProperty "Index",1	
										Fn_ADALicense_AuditLogOperations = False
										Set objLicense = Nothing
										Set objExcel = Nothing
										Exit Function
									End If
								Next
							Next
							'Call Fn_KillProcess("EXCEL.EXE")
							Call Fn_WindowsApplications("TerminateAll", "EXCEL.EXE")
							If Trim(Lcase(sClose)) = "yes" Then
								objLicense.Close 
							End If
			Case "LoadAllVerify"
						Set objLicense = JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Audit Log")
						'Check wheather LoadAll button is enabled
						If Fn_UI_Object_GetROProperty("Fn_ADALicense_AuditLogOperations",objLicense.JavaButton("LoadAll"), "enabled")=0 Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Load All button is disabled")
							Fn_ADALicense_AuditLogOperations = False
							If Trim(Lcase(sClose)) = "yes" Then
								objLicense.Close 
							End If
							Set objLicense = Nothing
							Exit Function
						End If
						'Click on LoadAll button.
						Call Fn_Button_Click("Fn_ADALicense_AuditLogOperations", objLicense, "LoadAll")
						Call Fn_ReadyStatusSync(2)
						'Get row count of Audit Log Table.
						iRowCount = Fn_Table_GetRowCount("Fn_ADALicense_AuditLogOperations",objLicense, "LogTable")						
						sReturn = objLicense.JavaStaticText("SearchObjectsFound").getROProperty("label")
						iReturn = split(sReturn, " ",-1,1)
						If Cint(iRowCount)=Cint(iReturn(0)) Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), iReturn(0) &" Objects Sucessfully found in Audit Log Table")							
						Else 
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), iReturn(0) &" Objects not found in Audit Log Table")
							Fn_ADALicense_AuditLogOperations = False
							If Trim(Lcase(sClose)) = "yes" Then
								objLicense.Close 
							End If
							Set objLicense = Nothing
							JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").SetTOProperty "Index",1
    						JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").SetTOProperty "Index",1	
							Exit Function
						End If
						If Trim(Lcase(sClose)) = "yes" Then
							objLicense.Close 
						End If
			Case "VerifyObject"
						iCount = 0
						iCounter = 0
						'Click on View Audit Log button
						If Fn_UI_ObjectExist("Fn_ADALicense_AuditLogOperations", JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Audit Log")) = False Then
									Call Fn_MenuOperation("Select","View:Audit:View Audit Logs")
									Call Fn_ReadyStatusSync(1)
						End If
						Set objLicense = Fn_UI_ObjectCreate("Fn_ADALicense_AuditLogOperations", JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Audit Log"))
						'Set Object ID
						If Trim(sObjID) <> "" Then
							iCount = iCount + 1
							If Trim(Lcase(Fn_UI_Object_GetROProperty("Fn_ADALicense_AuditLogOperations",objLicense.JavaEdit("ObjectID"),"text"))) = Trim(Lcase(sObjID)) Then
								iCounter = iCounter + 1
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Selected Object ID Matches with value ["+Trim(sObjID)+"] Verified Successfully.")
							End If
						End If
						'Set Object Name
						If sObjName<>"" Then
							iCount = iCount + 1
							If Fn_UI_ObjectExist("Fn_ADALicense_AuditLogOperations", objLicense.JavaEdit("ObjectName")) = True Then
								If Trim(Lcase(Fn_UI_Object_GetROProperty("Fn_ADALicense_AuditLogOperations",objLicense.JavaEdit("ObjectName"),"text"))) = Trim(Lcase(sObjName)) Then
									iCounter = iCounter + 1
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Selected Object Name Matches with value ["+Trim(sObjName)+"] Verified Successfully.")
								End If
							Else
								If Trim(Lcase(Fn_UI_Object_GetROProperty("Fn_ADALicense_AuditLogOperations",objLicense.JavaEdit("JobName"),"text"))) = Trim(Lcase(sObjName)) Then
									iCounter = iCounter + 1
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Selected Job Name Matches with value ["+Trim(sObjName)+"] Verified Successfully.")
								End If
							End If
						End If
						'Set Object Revision
						If sObjRev<>"" Then
							iCount = iCount + 1
							If Trim(Lcase(Fn_UI_Object_GetROProperty("Fn_ADALicense_AuditLogOperations",objLicense.JavaEdit("ObjectRevision"),"text"))) = Trim(Lcase(sObjRev)) Then
								iCounter = iCounter + 1
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Selected Object Revision Matches with value ["+Trim(sObjRev)+"] Verified Successfully.")
							End If
						End If
						'Set Object Type Name
						If sObjTypeName<>"" Then
							iCount = iCount + 1
							If Trim(Lcase(Fn_UI_Object_GetROProperty("Fn_ADALicense_AuditLogOperations",objLicense.JavaEdit("ObjectTypName"),"text"))) = Trim(Lcase(sObjTypeName)) Then
								iCounter = iCounter + 1
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Selected Object Type Name Matches with value ["+Trim(sObjTypeName)+"] Verified Successfully.")
							End If
						End If
						'Set Object Sequence Number
						If sObjSeqNum<>"" Then
							iCount = iCount + 1
							If Trim(Lcase(Fn_UI_Object_GetROProperty("Fn_ADALicense_AuditLogOperations",objLicense.JavaEdit("ObjectSeqNo"),"text"))) = Trim(Lcase(sObjSeqNum)) Then
								iCounter = iCounter + 1
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Selected Object Sequence Number Matches with value ["+Trim(sObjSeqNum)+"] Verified Successfully.")
							End If
						End If
						'Set Secondary Object ID
						If sObjSecID<>"" Then
							iCount = iCount + 1
							If Trim(Lcase(Fn_UI_Object_GetROProperty("Fn_ADALicense_AuditLogOperations",objLicense.JavaEdit("SecondaryObjID"),"text"))) = Trim(Lcase(sObjSecID)) Then
								iCounter = iCounter + 1
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Selected Secondary Object ID Matches with value ["+Trim(sObjSecID)+"] Verified Successfully.")								
							End If
						End If
						'Set Secondary Object Name
						If sObjSecName<>"" Then
							iCount = iCount + 1
							If Trim(Lcase(Fn_UI_Object_GetROProperty("Fn_ADALicense_AuditLogOperations",objLicense.JavaEdit("SecondaryObjName"),"text"))) = Trim(Lcase(sObjSecName)) Then
								iCounter = iCounter + 1
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Selected Secondary Object Name Matches with value ["+Trim(sObjSecName)+"] Verified Successfully.")
							End If
						End If
						'Set Secondary Object Revision
						If sObjSecRev<>"" Then
							iCount = iCount + 1
							If Trim(Lcase(Fn_UI_Object_GetROProperty("Fn_ADALicense_AuditLogOperations",objLicense.JavaEdit("SecondaryObjRev"),"text"))) = Trim(Lcase(sObjSecRev)) Then
								iCounter = iCounter + 1
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Selected Secondary Object Revision Matches with value ["+Trim(sObjSecRev)+"] Verified Successfully.")
							End If
						End If
						'Set Secondary Object Type
						If sObjSecType<>"" Then
							iCount = iCount + 1
							If Trim(Lcase(Fn_UI_Object_GetROProperty("Fn_ADALicense_AuditLogOperations",objLicense.JavaEdit("SecondaryObjType"),"text"))) = Trim(Lcase(sObjSecType)) Then
								iCounter = iCounter + 1
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Selected Secondary Object Type Matches with value ["+Trim(sObjSecType)+"] Verified Successfully.")
							End If
						End If
						'Set Secondary Object Sequence Number
						If sObjSecSeqNum<>"" Then
							iCount = iCount + 1
							If Trim(Lcase(Fn_UI_Object_GetROProperty("Fn_ADALicense_AuditLogOperations",objLicense.JavaEdit("SecondaryObjSeqNo"),"text"))) = Trim(Lcase(sObjSecSeqNum)) Then
								iCounter = iCounter + 1
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Selected Secondary Object Sequence Number Matches with value ["+Trim(sObjSecSeqNum)+"] Verified Successfully.")
							End If
						End If
						'Set Error Code
						If sErrCode<>"" Then
							iCount = iCount + 1
							If Trim(Lcase(Fn_UI_Object_GetROProperty("Fn_ADALicense_AuditLogOperations",objLicense.JavaEdit("ErrorCode"),"text"))) = Trim(Lcase(sErrCode)) Then
								iCounter = iCounter + 1
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Selected Error Code Matches with value ["+Trim(sErrCode)+"] Verified Successfully.")
							End If
						End If
						'Set Group Name
						If sGroupName<>"" Then
							iCount = iCount + 1
							If Trim(Lcase(Fn_UI_Object_GetROProperty("Fn_ADALicense_AuditLogOperations",objLicense.JavaEdit("GroupName"),"text"))) = Trim(Lcase(sGroupName)) Then
								iCounter = iCounter + 1
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Selected Group Name Matches with value ["+Trim(sGroupName)+"] Verified Successfully.")
							End If
						End If
						'Select Event Type Name
						If sEventTypeName<>"" Then
							iCount = iCount + 1
							If Trim(Lcase(Fn_UI_Object_GetROProperty("Fn_ADALicense_AuditLogOperations",objLicense.JavaEdit("EventTypeName"),"text"))) = Trim(Lcase(sEventTypeName)) Then
								iCounter = iCounter + 1
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Selected Event Type Name Matches with value ["+Trim(sEventTypeName)+"] Verified Successfully.")
							End If
						End If
						'Set User ID
						If sUserID<>"" Then
							iCount = iCount + 1
							If Trim(Lcase(Fn_UI_Object_GetROProperty("Fn_ADALicense_AuditLogOperations",objLicense.JavaEdit("UserID"),"text"))) = Trim(Lcase(sUserID)) Then
								iCounter = iCounter + 1
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Selected User ID Matches with value ["+Trim(sUserID)+"] Verified Successfully.")
							End If
						End If
						'Set Date Created Before
						If sDateCreBefore<>"" Then
							iCount = iCount + 1
							If InStr(1, Trim(objLicense.JavaCheckBox("DateCreatedBefore").GetROProperty("label")), Trim(sDateCreBefore), 1) > 0 Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verified Date Created Before ["+CStr(sDateCreBefore)+"]")
								iCounter = iCounter + 1
								Call Fn_ReadyStatusSync(2)
							End If
						End If
						'Set Date Created After
						If sDateCreAfter<>"" Then
							iCount = iCount + 1
							If InStr(1, objLicense.JavaCheckBox("DateCreatedAfter").GetROProperty("label"), Trim(sDateCreAfter), 1) > 0 Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verified Date Created Before ["+CStr(sDateCreAfter)+"]")
								iCounter = iCounter + 1
								Call Fn_ReadyStatusSync(2)
							End If
						End If
						'Click on Find button
						Call Fn_Button_Click("Fn_ADALicense_AuditLogOperations", objLicense, "Find")
						If Trim(Lcase(sClose)) = "yes" Then
							objLicense.Close 
						End If
						If iCount <> iCounter Then
							JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").SetTOProperty "Index",1
    						JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").SetTOProperty "Index",1	
							Set objLicense = Nothing
							Fn_ADALicense_AuditLogOperations = False
							Exit Function
						End If
			 Case "LegacyDataViewAudit"
						'Click on View Audit Log button
						bFlag = Fn_Button_Click("Fn_ADALicense_AuditLogOperations", JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Audit Log"), "LegacyDataViewAudit")
						If bFlag = False Then
								Fn_ADALicense_AuditLogOperations = FALSE
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Click on Legacy Data View Audit Button")
						Else
								Fn_ADALicense_AuditLogOperations = TRUE
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Clicked on Legacy Data View Audit Button")
						End If
			Case "AdvancedFind"
						'Case for Advanced Tab Operation
						Set objLicense = Fn_UI_ObjectCreate("Fn_ADALicense_AuditLogOperations", JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Audit Log"))
						objLicense.JavaTab("Tab").Select "Advanced"
						objLicense.RefreshObject
						'sObjID,sObjName,sObjRev,sObjTypeName,sObjSeqNum,sObjSecID,sObjSecName,sObjSecRev,sObjSecType,sObjSecSeqNum,sErrCode,sGroupName,sEventTypeName,sUserID,sDateCreBefore,sDateCreAfter,sClose
						If  Trim(sObjTypeName) <> "" Then
							Call Fn_List_Select("Fn_ADALicense_AuditLogOperations", objLicense, "AdvancedObjType",sObjTypeName)							
						End If
						If Trim(sEventTypeName) <> "" Then
							Call Fn_List_Select("Fn_ADALicense_AuditLogOperations", objLicense, "AdvancedEventType",sEventTypeName)
						End If
						objLicense.RefreshObject
						If Trim(sObjName) <> "" Then
							Call Fn_Edit_Box("Fn_ADALicense_AuditLogOperations", objLicense, "Name",sObjName)
						End If
						If Trim(sObjSecType) <> "" Then
							Call Fn_Edit_Box("Fn_ADALicense_AuditLogOperations", objLicense, "TaskType",sObjSecType)
						End If
						'Click on Find button
						Call Fn_Button_Click("Fn_ADALicense_AuditLogOperations", objLicense, "Find")
						' for Closing Dialog
						If Trim(Lcase(sClose)) = "yes" Then
							objLicense.Close 
						End If
			Case "Cancel"
						'Click on View Audit Log button
						bFlag = Fn_Button_Click("Fn_ADALicense_AuditLogOperations", JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Audit Log"), "Cancel")
						If bFlag = False Then
								Fn_ADALicense_AuditLogOperations = FALSE
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Click on Cancel Button")
						Else
								Fn_ADALicense_AuditLogOperations = TRUE
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Clicked on Cancel Button")
						End If
			Case "VerifyAvailableProperties"			'Added By Ketan On 09-May-2011.
						iCount = 0
						iCounter = 0
						'Case for Advanced Tab Operation
						Set objLicense = Fn_UI_ObjectCreate("Fn_ADALicense_AuditLogOperations", JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Audit Log"))
						aTableRecord = Split(sObjID,":",-1,1)
						'Verify all Object Properties.
						iRowCount = objLicense.JavaList("AvailableProperties").GetROProperty("items count")
						For iReturn = 0 to Ubound(aTableRecord)
							iCount = iCount + 1
							For iRow = 0 to iRowCount-1
								If Trim(Lcase(objLicense.JavaList("AvailableProperties").GetItem(iRow)))=Trim(Lcase(aTableRecord(iReturn))) Then
									iCounter = iCounter + 1
									Exit For
								End If
							Next
						Next
						If iCount = iCounter Then
								Fn_ADALicense_AuditLogOperations = True
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Values in the Available Properties match.")
						Else
								Fn_ADALicense_AuditLogOperations = False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Values in the Available Properties do not match.")
								If Trim(Lcase(sClose)) = "yes" Then
									objLicense.Close 
								End If
								JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").SetTOProperty "Index",1
    							JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").SetTOProperty "Index",1	
								Set objLicense = Nothing
								Exit Function
						End If
						' for Closing Dialog
						If Trim(Lcase(sClose)) = "yes" Then
							objLicense.Close 
						End If
			Case "VerifyUser"
'					Set objLicense = Fn_UI_ObjectCreate("Fn_ADALicense_AuditLogOperations", JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Audit Log"))
'					bFlag=False
'					arrUserID=Split(sUserID,"~")
'					JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Audit Log").JavaButton("UserID").Click micLeftBtn
'					wait 1
'					For iCounter=0 To UBound(arrUserID)
'						Set objStaticTxt=Description.Create()
'						objStaticTxt("Class Name").value = "JavaStaticText"
'						Set ObjChld=JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Audit Log").ChildObjects(objStaticTxt)
'						For iCount=0 To ObjChld.Count-1
'							If ObjChld(iCount).GetROProperty("label")=arrUserID(iCounter) Then
'								bFlag=True
'								Exit For
'							Else
'								bFlag=False
'							End If
'						Next
'						If bFlag=False Then
'							Exit For
'						End If
'					Next
					'*Code Added by Sanjeet Kumar 15-Feb-2013
					Dim objLOVTable,i,WshShell
					Set objLicense = Fn_UI_ObjectCreate("Fn_ADALicense_AuditLogOperations", JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Audit Log"))
					bFlag=False
					arrUserID=Split(sUserID,"~")
					JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Audit Log").JavaButton("UserID").Click micLeftBtn
					wait 2
					Set WshShell = CreateObject("WScript.Shell")
					wait 1
					For i=0 to 2
						wait 1			
						WshShell.SendKeys "^"		
						WshShell.SendKeys "^{END}"
					Next
					Set WshShell =nothing
					wait 1

					Set ObjChld=JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Audit Log").JavaTable("UserTable")

					For iCounter=0 To UBound(arrUserID)
							bFlag=False
							For iCount=0 To Cint(ObjChld.GetROProperty("rows"))-1
									If ObjChld.Object.getValueAt(iCount,0).getDisplayableValue()=arrUserID(iCounter) Then
										bFlag=True
										Exit For
									End If
							Next
							If bFlag=False Then
								Exit For
							End If
					Next

					Set objLOVTable=Nothing
					Set ObjChld=Nothing

					If Trim(Lcase(sClose)) = "yes" Then
						objLicense.Close 
					End If
					If bFlag=False Then
						Fn_ADALicense_AuditLogOperations = FALSE
						JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").SetTOProperty "Index",1
    					JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").SetTOProperty "Index",1	
						Exit Function
					End If

			Case Else			
						JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").SetTOProperty "Index",1
    					JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").SetTOProperty "Index",1
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_ADALicense_AuditLogOperations function failed")
						Fn_ADALicense_AuditLogOperations = FALSE
						Exit Function						
		End Select
	
	JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").SetTOProperty "Index",1
    JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").SetTOProperty "Index",1	
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" Sucessfully completed of function Fn_ADALicense_AuditLogOperations")
    Set objLicense = nothing 
	Fn_ADALicense_AuditLogOperations = TRUE		
End Function
'#########################################################################################################
'###
'###    FUNCTION NAME   :   Fn_ADALicense_LicenseAccosiation(sAction,sLicense,sApply,sOK)
'###
'###    DESCRIPTION        :   AddAttachLicenses / RemoveAttachLicenses / AddDetachLicenses / RemoveDetachLicenses / VerifyAttachLicenses / VerifyDetachLicenses
'###
'###    PARAMETERS      :   1. sAction
'###											 2.	sLicenses
'###											 3.	sApply
'###											 4.	sOK
'###                                         
'###    Function Calls       :   Fn_WriteLogFile(), Fn_UI_ObjectCreate(), Fn_Button_Click()
'###
'###	 HISTORY             :   AUTHOR                 DATE        VERSION
'###
'###    CREATED BY     :   Ketan Raje           12/07/2010         1.0
'###
'###    REVIWED BY     :   Harshal
'###
'###    MODIFIED BY   :  
'###
'###    EXAMPLE          : 		Case "AddAttachLicenses" : Call Fn_ADALicense_LicenseAccosiation("AddAttachLicenses","ITAR_Lic107201015114:ITAR_License1272010154944","yes","yes")
'###										 Case "RemoveAttachLicenses" : Call Fn_ADALicense_LicenseAccosiation("RemoveAttachLicenses","ITAR_Lic107201015114:ITAR_License1272010154944","no","no")
'###										Case "AddDetachLicenses" : Call Fn_ADALicense_LicenseAccosiation("AddDetachLicenses","ITAR_Lic21272010152555:ITAR_License1272010154944","no","no")
'###										Case "RemoveDetachLicenses" : Call Fn_ADALicense_LicenseAccosiation("RemoveDetachLicenses","ITAR_Lic21272010152555:ITAR_License1272010154944","yes","yes")
'###										Case "EditAuthTable" : Call Fn_ADALicense_LicenseAccosiation("EditAuthTable","license 123:ABC","no","no")
'###										Case "VerifyAttachAvailablelist" : Call Fn_ADALicense_LicenseAccosiation("VerifyAttachAvailablelist","license 123:TestSamir1","no","no")
'###										Case "VerifyDetachAvailablelist" : Call Fn_ADALicense_LicenseAccosiation("VerifyDetachAvailablelist","ITAR_Lic11572010113940aix:ITAR_Lic11572010143233g4","no","no")
'###										Case "SelectLicenseType" : Call Fn_ADALicense_LicenseAccosiation("SelectLicenseType","IP_License","","")
'###										Case "VerifyAttachButtonsEnable" : Call Fn_ADALicense_LicenseAccosiation("VerifyAttachButtonsEnable","AddAllLicenses:AddLicenses:RemoveAllLicenses:RemoveLicenses:Apply:Cancel:OK","","")
'###										Case "AttachListLastVisibleIndex" : Call Fn_ADALicense_LicenseAccosiation("AttachListLastVisibleIndex","","","")
'#############################################################################################################
Public Function Fn_ADALicense_LicenseAccosiation(sAction,sLicense,sApply,sOK)
	GBL_FAILED_FUNCTION_NAME="Fn_ADALicense_LicenseAccosiation"
	Dim objLicense, iCounter, bReturn, aColname, iCount, iRowData, aLicense, sRowCount, sLicensesName, ObjList, sAuthData
	'wait(5)
		Select Case sAction
				Case "AddAttachLicenses"
						Set objLicense = Fn_UI_ObjectCreate("Fn_ADALicense_LicenseAccosiation", JavaWindow("MyTeamcenter").JavaWindow("MyTcApplet").JavaDialog("AttachLicenses"))	
						If sLicense<>"" Then
								JavaWindow("MyTeamcenter").JavaWindow("MyTcApplet").JavaDialog("AttachLicenses").JavaList("AvailableLicenses").WaitProperty "displayed",1
                                bReturn = objLicense.JavaList("AvailableLicenses").GetROProperty("items count")
								'Extract the index of row at which the object exist.
								aColname = split(sLicense, ":",-1,1)
								iCount = Ubound(aColname)
								For iRowData=0 to iCount
									For iCounter=0 to bReturn-1
										If Trim(lcase(objLicense.JavaList("AvailableLicenses").GetItem(iCounter))) = Trim(lcase(aColname(iRowData))) then
											objLicense.JavaList("AvailableLicenses").Select aColname(iRowData)
											'Click on Add Button
											Call Fn_Button_Click("Fn_ADALicense_LicenseAccosiation", objLicense, "AddLicenses")
											Exit For 
										End If
									Next
								Next
						End If
										
				Case "RemoveAttachLicenses"
						Set objLicense = Fn_UI_ObjectCreate("Fn_ADALicense_LicenseAccosiation", JavaWindow("MyTeamcenter").JavaWindow("MyTcApplet").JavaDialog("AttachLicenses"))	
						If sLicense<>"" Then
								bReturn = objLicense.JavaList("SelectedLicenses").GetROProperty("items count")
								'Extract the index of row at which the object exist.
								aColname = split(sLicense, ":",-1,1)
								iCount = Ubound(aColname)
								For iRowData=0 to iCount
									For iCounter=0 to bReturn-1
										If Trim(lcase(objLicense.JavaList("SelectedLicenses").GetItem(iCounter))) = Trim(lcase(aColname(iRowData))) then
											objLicense.JavaList("SelectedLicenses").Select aColname(iRowData)
											'Click on Remove Button
											Call Fn_Button_Click("Fn_ADALicense_LicenseAccosiation", objLicense, "RemoveLicenses")
											Exit For 
										End If
									Next
								Next
						End If

				Case "AddDetachLicenses"
						Set objLicense = Fn_UI_ObjectCreate("Fn_ADALicense_LicenseAccosiation", JavaWindow("MyTeamcenter").JavaWindow("MyTcApplet").JavaDialog("DetachLicenses"))	
						If sLicense<>"" Then
								bReturn = objLicense.JavaList("AvailableLicenses").GetROProperty("items count")
								'Extract the index of row at which the object exist.
								aColname = split(sLicense, ":",-1,1)
								iCount = Ubound(aColname)
								For iRowData=0 to iCount
									For iCounter=0 to bReturn-1
										If Trim(lcase(objLicense.JavaList("AvailableLicenses").GetItem(iCounter))) = Trim(lcase(aColname(iRowData))) then
											objLicense.JavaList("AvailableLicenses").Select aColname(iRowData)
											'Click on Add Button
											Call Fn_Button_Click("Fn_ADALicense_LicenseAccosiation", objLicense, "AddLicenses")										
											Exit For 
										End If
									Next
								Next
						End If				

				Case "RemoveDetachLicenses"
						Set objLicense = Fn_UI_ObjectCreate("Fn_ADALicense_LicenseAccosiation", JavaWindow("MyTeamcenter").JavaWindow("MyTcApplet").JavaDialog("DetachLicenses"))	
						If sLicense<>"" Then
								bReturn = objLicense.JavaList("SelectedLicenses").GetROProperty("items count")
								'Extract the index of row at which the object exist.
								aColname = split(sLicense, ":",-1,1)
								iCount = Ubound(aColname)
								For iRowData=0 to iCount
									For iCounter=0 to bReturn-1
										If Trim(lcase(objLicense.JavaList("SelectedLicenses").GetItem(iCounter))) = Trim(lcase(aColname(iRowData))) then
											objLicense.JavaList("SelectedLicenses").Select aColname(iRowData)
											'Click on Remove Button
											Call Fn_Button_Click("Fn_ADALicense_LicenseAccosiation", objLicense, "RemoveLicenses")
											Exit For 
										End If
									Next
								Next
						End If

				Case "EditAuthTable","VerifyAuthTable"
                        aLicense = Split(sLicense,":")
                        Set objLicense = Fn_UI_ObjectCreate("Fn_ADALicense_LicenseAccosiation", JavaWindow("MyTeamcenter").JavaWindow("MyTcApplet").JavaDialog("AttachLicenses"))
                        'Check enablity of Authorizing Table
                        If objLicense.JavaTable("AuthorizingTable").GetROProperty("enabled")=0 Then
                            Fn_ADALicense_LicenseAccosiation = False
                            Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "AuthorizingTable is disabled")
                            Exit Function                                
                        End If
                        'Get rows of Authorizing Table
                        sRowCount = objLicense.JavaTable("AuthorizingTable").GetROProperty("rows")
                        For iCounter = 0 to cint(sRowCount) -1
                        sLicensesName = objLicense.JavaTable("AuthorizingTable").GetCellData(iCounter,"License Name")
                        If trim(lcase(aLicense(0))) = trim(lcase(sLicensesName)) Then
                            If Trim(Lcase(sAction)) = "editauthtable" Then
                                objLicense.JavaTable("AuthorizingTable").ClickCell iCounter,1
                                objLicense.JavaTable("AuthorizingTable").SetCellData iCounter,"Authorizing Paragraph",aLicense(1)
                                objLicense.JavaTable("AuthorizingTable").ClickCell iCounter,1
                                Fn_ADALicense_LicenseAccosiation = True
                                Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Edited value "+aLicense(1)+" of License "+aLicense(0)+" In AuthorizingTable")
                            ElseIf Trim(Lcase(sAction)) = "verifyauthtable" Then
                                sAuthData = objLicense.JavaTable("AuthorizingTable").GetCellData(iCounter,"Authorizing Paragraph")
                                If Trim(Lcase(sAuthData)) = Trim(Lcase(aLicense(1))) Then
                                    Fn_ADALicense_LicenseAccosiation = True
                                    Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verified License Data of "+aLicense(0)+" In AuthorizingTable")
                                Else
                                    Fn_ADALicense_LicenseAccosiation = False
                                    Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to verify License Data of "+aLicense(0)+" In AuthorizingTable")
                                    Exit Function                                                                
                                End If
                            End If
                            Exit For
                        End If
                        Next
                        If Cint(iCounter) = Cint(sRowCount) Then
                            Fn_ADALicense_LicenseAccosiation = False
                            Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Find License "+aLicense(0)+" In AuthorizingTable")
                            Exit Function                            
                        End If

				Case "VerifyAttachAvailablelist","VerifyAttachSelectedlist"
						Set objLicense = Fn_UI_ObjectCreate("Fn_ADALicense_LicenseAccosiation", JavaWindow("MyTeamcenter").JavaWindow("MyTcApplet").JavaDialog("AttachLicenses"))	
						If sAction = "VerifyAttachAvailablelist" Then
							Set ObjList = Fn_UI_ObjectCreate("Fn_ADALicense_LicenseAccosiation", JavaWindow("MyTeamcenter").JavaWindow("MyTcApplet").JavaDialog("AttachLicenses").JavaList("AvailableLicenses"))								
						Else						
							Set ObjList = Fn_UI_ObjectCreate("Fn_ADALicense_LicenseAccosiation", JavaWindow("MyTeamcenter").JavaWindow("MyTcApplet").JavaDialog("AttachLicenses").JavaList("SelectedLicenses"))	
						End If
							bReturn = ObjList.GetROProperty("items count")							
							If sLicense<>"" AND Cint(bReturn)<> 0 Then 
									'Extract the index of row at which the object exist.
									aColname = split(sLicense, ":",-1,1)
									iCount = Ubound(aColname)
									For iRowData=0 to iCount
										For iCounter=0 to Cint(bReturn)-1
											If Trim(lcase(ObjList.GetItem(iCounter))) = Trim(lcase(aColname(iRowData))) then
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), aColname(iRowData) &" Sucessfully found in License List")
												Fn_ADALicense_LicenseAccosiation = TRUE
												Exit For 
											Elseif iCounter = Cint(bReturn)-1 Then
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), aColname(iRowData) &" not found in License List")
												Fn_ADALicense_LicenseAccosiation = FALSE
												Exit Function											
											End If
										Next
									Next
							Else
									Fn_ADALicense_LicenseAccosiation = FALSE
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "No License present in License List")
									Exit Function 
							End If

				Case "VerifyDetachAvailablelist","VerifyDetachSelectedlist"
						Set objLicense = Fn_UI_ObjectCreate("Fn_ADALicense_LicenseAccosiation", JavaWindow("MyTeamcenter").JavaWindow("MyTcApplet").JavaDialog("DetachLicenses"))
						If sAction = "VerifyDetachAvailablelist" Then
								Set ObjList = Fn_UI_ObjectCreate("Fn_ADALicense_LicenseAccosiation", JavaWindow("MyTeamcenter").JavaWindow("MyTcApplet").JavaDialog("DetachLicenses").JavaList("AvailableLicenses"))
						Else												
								Set ObjList = Fn_UI_ObjectCreate("Fn_ADALicense_LicenseAccosiation", JavaWindow("MyTeamcenter").JavaWindow("MyTcApplet").JavaDialog("DetachLicenses").JavaList("SelectedLicenses"))
						End If
								bReturn = ObjList.GetROProperty("items count")	
							If sLicense<>"" AND Cint(bReturn)<> 0 Then
									'Extract the index of row at which the object exist.
									aColname = split(sLicense, ":",-1,1)
									iCount = Ubound(aColname)
									For iRowData=0 to iCount
										For iCounter=0 to Cint(bReturn)-1
											If Trim(lcase(ObjList.GetItem(iCounter))) = Trim(lcase(aColname(iRowData))) then
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), aColname(iRowData) &" Sucessfully found in Available License List")
												Fn_ADALicense_LicenseAccosiation = TRUE
												Exit For 
											Elseif iCounter = Cint(bReturn)-1 Then
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), aColname(iRowData) &" not found in Available License List")
												Fn_ADALicense_LicenseAccosiation = FALSE
												Exit Function											
											End If
										Next
									Next
							Else
									Fn_ADALicense_LicenseAccosiation = FALSE
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "No License present in Available License List")
									Exit Function 
							End If

				Case "SelectLicenseType"
							sLicense= replace(sLicense,"_"," ")
							Set objLicense = Fn_UI_ObjectCreate("Fn_ADALicense_LicenseAccosiation", JavaWindow("MyTeamcenter").JavaWindow("MyTcApplet").JavaDialog("AttachLicenses").JavaList("LicenseType"))
							objLicense.Select sLicense	

				Case "VerifyAttachButtonsEnable"
						Set objLicense = Fn_UI_ObjectCreate("Fn_ADALicense_LicenseAccosiation", JavaWindow("MyTeamcenter").JavaWindow("MyTcApplet").JavaDialog("AttachLicenses"))
						aLicense = Split(sLicense,":")
						For iCounter=0 to ubound(aLicense)
							If Fn_UI_Object_GetROProperty("Fn_ADALicense_LicenseAccosiation",objLicense.JavaButton(aLicense(iCounter)), "enabled")=1 Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), aLicense(iCounter) &" button is enabled")
							Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), aLicense(iCounter) &" button is disabled")
								Set objLicense = Nothing
								Fn_ADALicense_LicenseAccosiation = False
								Exit Function
							End If
						Next

				Case "AttachListLastVisibleIndex"
							Set objLicense = Fn_UI_ObjectCreate("Fn_ADALicense_LicenseAccosiation", JavaWindow("MyTeamcenter").JavaWindow("MyTcApplet").JavaDialog("AttachLicenses").JavaList("AvailableLicenses"))
							Fn_ADALicense_LicenseAccosiation = objLicense.Object.getLastVisibleIndex
							Set objLicense = nothing 	
							Exit Function

				'-------------------------------------------------------------------------------------------------------------------
				Case "AddAttachLicenses_SM" '[TC11.5_20180616b.00_NewDevelopment_PoonamC_10Oct2018]
						Set objLicense = Fn_SISW_PSE_GetObject("PSEApplet")
						Set objLicense = Fn_UI_ObjectCreate("Fn_ADALicense_LicenseAccosiation", objLicense.JavaDialog("AttachLicenses"))	
						If sLicense<>"" Then
								objLicense.JavaList("AvailableLicenses").WaitProperty "displayed",1
                                bReturn = objLicense.JavaList("AvailableLicenses").GetROProperty("items count")
								'Extract the index of row at which the object exist.
								aColname = split(sLicense, ":",-1,1)
								iCount = Ubound(aColname)
								For iRowData=0 to iCount
									For iCounter=0 to bReturn-1
										If Trim(lcase(objLicense.JavaList("AvailableLicenses").GetItem(iCounter))) = Trim(lcase(aColname(iRowData))) then
											objLicense.JavaList("AvailableLicenses").Select aColname(iRowData)
											'Click on Add Button
											Call Fn_Button_Click("Fn_ADALicense_LicenseAccosiation", objLicense, "AddLicenses")
											Exit For 
										End If
									Next
								Next
						End If
				Case "SelectLicenseType_SM" '[TC11.5_20180616b.00_NewDevelopment_PoonamC_10Oct2018]
							Set objLicense = Fn_SISW_PSE_GetObject("PSEApplet")
							Set objLicense = Fn_UI_ObjectCreate("Fn_ADALicense_LicenseAccosiation", objLicense.JavaDialog("AttachLicenses"))
							sLicense = replace(sLicense, "_", " ") 
							Fn_ADALicense_LicenseAccosiation = Fn_SISW_UI_JavaList_Operations("Fn_ADALicense_LicenseAccosiation", "Select", objLicense, "LicenseType", sLicense, "", "")			
				'----------------------------------------------------------------------------------------------------------------------------------		
				Case Else						
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_ADALicense_LicenseAccosiation function failed")
							Fn_ADALicense_LicenseAccosiation = FALSE
							Set objLicense = nothing 
							Exit Function						
		End Select
		If Trim(Lcase(sApply))="yes" Then
			'Click on Apply button
			Call Fn_Button_Click("Fn_ADALicense_LicenseAccosiation", objLicense, "Apply")
		End If
		If Trim(Lcase(sOK))="yes" Then
			'Click on OK button
			Call Fn_Button_Click("Fn_ADALicense_LicenseAccosiation", objLicense, "OK")
		End If
		Fn_ADALicense_LicenseAccosiation = True
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" Sucessfully completed of function Fn_ADALicense_LicenseAccosiation")
    Set objLicense = nothing 
	Set ObjList = nothing
End Function
'#########################################################################################################
'###
'###    FUNCTION NAME   :   Fn_AuditLogTableRMB(sMenu,aColumnNames)
'###
'###    DESCRIPTION        :   RMB operation on Audit Log Table
'###
'###    PARAMETERS      :   1. sMenu
'###											 2.	aColumnNames
'###                                         
'###    Function Calls       :   Fn_WriteLogFile()
'###
'###	 HISTORY             :   AUTHOR                 DATE        VERSION
'###
'###    CREATED BY     :   Ketan Raje           16/07/2010         1.0
'###
'###    REVIWED BY     :   Harshal
'###
'###    MODIFIED BY   :  Ketan Raje				14/09/2011			1.0
'###
'###    EXAMPLE          : 		Call Fn_AuditLogTableRMB("Print Table:Graphics","UserId")
'###    Case "Exist"      : 	  Msgbox Fn_AuditLogTableRMB("Exist|Print Table:Graphics","Event Type Name")
'#########################################################################################################
Public Function Fn_AuditLogTableRMB(sMenu,aColumnNames)
		GBL_FAILED_FUNCTION_NAME="Fn_AuditLogTableRMB"
		Dim aMenuList, intCount, bFlag, bReturnm, bColFoundFlag, iCounter, aColname, aColHeadName, var, childObjects
		Dim sAction, aMenu
		sAction = ""
			If Instr(1,sMenu,"|") <> 0 Then
				aMenu = Split(sMenu, "|")
				sAction = aMenu(0)
				sMenu = aMenu(1)
			End If
			bFlag = False
			
			If JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Audit Log").exist(5) = False Then
				JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").SetTOProperty "Index",0
			End If
			'get  Count of Columns 
			bReturn = JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Audit Log").JavaTable("LogTable").GetROProperty("cols")
			bColFoundFlag = False
			' RMB on Column Header
			For iCounter = 0 to bReturn - 1 
						' get  text property of column 
						aColname= split(JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Audit Log").JavaTable("LogTable").GetColumnName (iCounter),"text=")
						aColHeadName = split(aColname(0),",")
						'Right click on table column header  if  column  is available  in details table  to open java menu
						If aColHeadName(0) = aColumnNames Then
									bColFoundFlag = true
									Exit For
						End If
			Next
			If bColFoundFlag = False Then
					iCounter = 0 
			End If
			If sAction <> "" Then
				Select Case sAction
							Case "Exist"
									JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Audit Log").JavaTable("LogTable").SelectColumnHeader iCounter,"RIGHT"
									'  Selecting item from RMB menu
									set var = Description.Create()
									var("Class Name").value = "JavaMenu"
									set childObjects = JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Audit Log").ChildObjects(var)
									If childObjects.count <> 0 then
													aMenuList = split(sMenu, ":",-1,1)
													intCount = Ubound(aMenuList)
													Select Case intCount
														Case "0"
															If Not JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Audit Log").JavaMenu("label:="&aMenuList(0),"index:=0").Exist Then
																Fn_AuditLogTableRMB = False
																Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sMenu&" does not exist.")
																JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").SetTOProperty "Index",1
																Exit Function
															End If
														Case "1"
															If Not JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Audit Log").JavaMenu("label:="&aMenuList(0),"index:=0").JavaMenu("label:="&aMenuList(1),"index:=1").Exist Then
																Fn_AuditLogTableRMB = False
																Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sMenu&" does not exist.")
																JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").SetTOProperty "Index",1
																Exit Function
															End If
														Case "2"
															If Not JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Audit Log").JavaMenu("label:="&aMenuList(0),"index:=0").JavaMenu("label:="&aMenuList(1),"index:=1").JavaMenu("label:="&aMenuList(2),"index:=2").Exist Then
																Fn_AuditLogTableRMB = False
																Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sMenu&" does not exist.")
																JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").SetTOProperty "Index",1
																Exit Function
															End If
														Case Else
															Fn_AuditLogTableRMB = FALSE
															Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_AuditLogTableRMB Failed to select Menu "&sMenu)
															JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").SetTOProperty "Index",1
															Exit Function
													End Select								
									Else
										   Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_AuditLogTableRMB Failed to select Menu "&sMenu)
										   Fn_AuditLogTableRMB = False									   
									End If
				End Select
			Else
					JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Audit Log").JavaTable("LogTable").SelectColumnHeader iCounter,"RIGHT"
					'  Selecting item from RMB menu
					set var = Description.Create()
					var("Class Name").value = "JavaMenu"
					set childObjects = JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Audit Log").ChildObjects(var)
					If childObjects.count <> 0 then
									aMenuList = split(sMenu, ":",-1,1)
									intCount = Ubound(aMenuList)
									Select Case intCount
										Case "0"
											JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Audit Log").JavaMenu("label:="&aMenuList(0),"index:=0").Select
										Case "1"
											JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Audit Log").JavaMenu("label:="&aMenuList(0),"index:=0").JavaMenu("label:="&aMenuList(1),"index:=1").Select
										Case "2"
											JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Audit Log").JavaMenu("label:="&aMenuList(0),"index:=0").JavaMenu("label:="&aMenuList(1),"index:=1").JavaMenu("label:="&aMenuList(2),"index:=2").Select
										Case Else
											Fn_AuditLogTableRMB = FALSE
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_AuditLogTableRMB Failed to select Menu "&sMenu)
											JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").SetTOProperty "Index",1
											Exit Function
									End Select								
					Else
						   Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_AuditLogTableRMB Failed to select Menu "&sMenu)
							JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").SetTOProperty "Index",1						  
							Fn_AuditLogTableRMB = False									   
					End If
			End If
		Fn_AuditLogTableRMB = TRUE
		JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").SetTOProperty "Index",1
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_AuditLogTableRMB Sucessfully Selected Menu "&sMenu)
End Function
'*********************************************************Function to select  the Tab into ADALicense***********************************************************************
'Function Name		:				Fn_ADALicense_PrintOptionVerify(sAction)

'Description			 :		 		 Check the Printer Option

'Parameters			   :	 			sAction
											
'Return Value		   : 				TRUE \ FALSE

'Pre-requisite			:		 		Print Dialog should be present

'Examples				:				 Call Fn_ADALicense_PrintOptionVerify("Page Setup")
													 'Call Fn_ADALicense_PrintOptionVerify("HTML")

'History:
'	Developer Name			Date			Rev. No.			Changes Done		Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Harshal						19/07/2010						
'	Veena						03/01/2013						Added Case Print	Kowstubh Watwe
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function Fn_ADALicense_PrintOptionVerify(sAction)
	GBL_FAILED_FUNCTION_NAME="Fn_ADALicense_PrintOptionVerify"
	Dim ObjPrint,ObjPrinter
	Select Case sAction
	Case "HTML"
    	Set ObjPrint = Fn_UI_ObjectCreate("Fn_ADALicense_PrintOptionVerify",JavaWindow("ADA License - Teamcenter").JavaWindow("ADA License").JavaDialog("Print"))
		If  Fn_UI_ObjectExist("Fn_ADALicense_PrintOptionVerify",objPrint) Then
			Call Fn_Button_Click("Fn_ADALicense_PrintOptionVerify",ObjPrint,"Print")
			Set ObjPrinter = Fn_UI_ObjectCreate("Fn_ADALicense_PrintOptionVerify",Dialog("Print"))
				If Fn_UI_ObjectExist("Fn_ADALicense_PrintOptionVerify",objPrinter)  Then
					If cint(ObjPrinter.WinButton("Print").GetROProperty("enabled")) =True then
						 Fn_ADALicense_PrintOptionVerify = True
						 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Verify Print Button Is Enable")
						Call Fn_UI_WinButton_Click("Fn_ADALicense_PrintOptionVerify",ObjPrinter,"Cancel","","","")
						Call Fn_Button_Click("Fn_ADALicense_PrintOptionVerify",ObjPrint,"Close")
					 Else
						Fn_ADALicense_PrintOptionVerify =False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fali ToVerify Print Button Is Enable")
				End If 
			Else
					Fn_ADALicense_PrintOptionVerify =False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail To Verify Print Dialog Existance")
			End If
		Else
			Fn_ADALicense_PrintOptionVerify =False
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail To Verify TCPrint Dialog Existance")
		End If
	Case "Page Setup"
		'Set ObjPrinter  =  Fn_UI_ObjectCreate("Fn_ADALicense_PrintOptionVerify",JavaDialog("ADAPrint"))
		Set ObjPrinter  =  Fn_UI_ObjectCreate("Fn_ADALicense_PrintOptionVerify",JavaWindow("ADA License - Teamcenter").JavaWindow("ADA License").JavaDialog("ADAPrint"))
			If  Fn_UI_ObjectExist("Fn_ADALicense_PrintOptionVerify",objPrinter) Then
				ObjPrinter.JavaTab("Accepting jobs").Select "Page Setup"
				If ObjPrinter.JavaButton("Print").GetROProperty("enabled") =1 Then
					Fn_ADALicense_PrintOptionVerify = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Verify Print Button Is Enable")
					wait(2)
'					Call Fn_UI_WinButton_Click("Fn_ADALicense_PrintOptionVerify",ObjPrinter,"Cancel","","","")
					Call Fn_Button_Click("Fn_ADALicense_PrintOptionVerify", ObjPrinter, "Cancel")
                 Else
					Fn_ADALicense_PrintOptionVerify = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail to Verify Print Button Is Enable")
				End If
			Else
					Fn_ADALicense_PrintOptionVerify = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail to Find Page Setup Dialog")
			End If
	Case "Print"  'As per telephonic disscussion with Dhananjay Niwal Added Case for Print dialog by Veena Gurjar [3-Jan-2013]
		Set ObjPrinter  =  Fn_UI_ObjectCreate("Fn_ADALicense_PrintOptionVerify",JavaDialog("Print"))
			If  Fn_UI_ObjectExist("Fn_ADALicense_PrintOptionVerify",objPrinter) Then
				If JavaDialog("Print").JavaButton("Print").GetROProperty("enabled") =1 Then
					Fn_ADALicense_PrintOptionVerify = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Verify Print Button Is Enable")
						Call Fn_Button_Click("Fn_ADALicense_PrintOptionVerify",ObjPrinter,"Cancel")
                 Else
					Fn_ADALicense_PrintOptionVerify = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail to Verify Print Button Is Enable")
				End If
			Else
					Fn_ADALicense_PrintOptionVerify = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail to Find Print Dialog")
			End If
	End Select
	Set ObjPrint = Nothing
	Set ObjPrinter = Nothing
End Function
