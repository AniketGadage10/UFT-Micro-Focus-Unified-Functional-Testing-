Option Explicit

'= = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
' Function name should be : Fn_SISW_SPM_FunctionName
'= = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

' Function List
'= = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'1.  Fn_SISW_SPM_DeleteOverrideRecordOperations 
'= = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_SISW_SPM_DeleteOverrideRecordOperations

'Description			 :	Function Used to perform operations on  Delete Override Dialog

'Parameters			   :  1.StrAction : Action name
'									2.StrMessage: Message To Verify From Delete Override Dialog
'								 	3.StrDeleteOverrideOption: Delete override options
'								    4.StrButton: Button name
'
'Return Value		   : 	True or False

'Pre-requisite			:	Delete Override Dialog should be open					

'Examples				:  

'                       
'History					 :			
'										Developer Name							Date						Rev. No.				Changes Done											Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'										Sandeep N								15-Nov-2012					1.0																								Sonal P
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Public Function Fn_SISW_SPM_DeleteOverrideRecordOperations(StrAction,StrMessage,StrDeleteOverrideOption,StrButton)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SPM_DeleteOverrideRecordOperations"
   'declaring variables
   Dim objDeleteOverrideRecord, sDispMsg
	'creating object of [ DeleteOverrideRecord ] dialog	
	Set objDeleteOverrideRecord = JavaWindow("SoftwareParameterManager").JavaWindow("DeleteOverrideRecord")
	Fn_SISW_SPM_DeleteOverrideRecordOperations = False
    'checking existance of [ DeleteOverrideRecord ] dialog	
	If Fn_UI_ObjectExist("Fn_SISW_SPM_DeleteOverrideRecordOperations",objDeleteOverrideRecord ) = False  Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_SPM_DeleteOverrideRecordOperations ] Nat Table object of [ Traceability Matrix ] is not visible.")
		Exit function
	End If

	Select Case StrAction
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		'case to verify available delete override options
		Case "VerifyDeleteOverrideOptions"
				arrOptions=Split(StrDeleteOverrideOption,"~")
				For iCounter=0 to ubound(arrOptions)
					bFlag=False
					objDeleteOverrideRecord.JavaRadioButton("DeleteOverrideRecordOption").SetTOProperty "attached text",arrOptions(iCounter)
					If objDeleteOverrideRecord.JavaRadioButton("DeleteOverrideRecordOption").Exist(2) Then
						bFlag=True
					End If
					If bFlag=False Then
						Exit for
					End If
				Next
				If bFlag=True Then
					Fn_SISW_SPM_DeleteOverrideRecordOperations=True
				End If
		 '- - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		'case to delete override
		Case "Delete"
			If StrDeleteOverrideOption<>"" Then
				objDeleteOverrideRecord.JavaRadioButton("DeleteOverrideRecordOption").SetTOProperty "attached text",StrDeleteOverrideOption
				Call Fn_UI_JavaRadioButton_SetON("Fn_SISW_SPM_DeleteOverrideRecordOperations",objDeleteOverrideRecord, "DeleteOverrideRecordOption")
			End If
			If StrButton="" Then
				StrButton="Yes"
			End If
			Fn_SISW_SPM_DeleteOverrideRecordOperations = Fn_Button_Click("Fn_SISW_SPM_DeleteOverrideRecordOperations", objDeleteOverrideRecord,StrButton)
		 '- - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "VerifyMessage"
        	sDispMsg = objDeleteOverrideRecord.JavaStaticText("sMsg").GetROProperty("value")
			If  Instr(1,sDispMsg,StrMessage) > 0 Then
				Fn_SISW_SPM_DeleteOverrideRecordOperations=True
			Else
				Fn_SISW_SPM_DeleteOverrideRecordOperations=False
			End If

    		Call Fn_Button_Click("Fn_SISW_SPM_DeleteOverrideRecordOperations", objDeleteOverrideRecord, StrButton)

	' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_SPM_DeleteOverrideRecordOperations ] Invalid case [ " & StrAction & " ].")
	' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
	End Select
	If  Fn_SISW_SPM_DeleteOverrideRecordOperations <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [ Fn_SISW_SPM_DeleteOverrideRecordOperations ] executed successfuly with case [ " & StrAction & " ].")
	End If
	'releasing object of [ DeleteOverrideRecord ] dialog	
	Set objDeleteOverrideRecord = Nothing
End Function
