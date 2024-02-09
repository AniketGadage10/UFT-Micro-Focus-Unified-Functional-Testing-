
' *********************************************************	UI_AWSLibrary  Function List		***********************************************************************
'1.  Fn_DSM_UI_ObjectExist()												This function is Use to check Existance of given Object
'2.  Fn_DSM_UIButton_Operations()											     This function is Use to perform button operation
' *********************************************************	End Of List		****************************************************************************************




'#########################################################################################################
'###
'###    FUNCTION NAME   :  Fn_DSM_UI_ObjectExist(sFunctionName, sReferencePath)
'###
'###    DESCRIPTION     :   This function is Use to check Existance of given Object
'###
'###    PARAMETERS      :   1.sFunctionName: Valid Function Name
'###			    							2.sReferencePath: Valid reference path
'###                                         
'###	 HISTORY         :   AUTHOR                 DATE        
'###
'###    CREATED BY      :   Amruta Patil       		  16/07/2021     

'###    EXAMPLE         : Fn_DSM_UI_ObjectExist("Fn_DSM_UI_ObjectExist", "UIAWindow("Teamcenter RAC (Eclipse)").UIATab("Summary").UIAObject("LeftTab").UIAButton("TabUIButton") ")
'#############################################################################################################

Function Fn_DSM_UI_ObjectExist(sFunctionName, sReferencePath)
		Dim objDialog
		Set objDialog = sReferencePath

		If objDialog.Exist Then			
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"UI_AWS : New Object is Exist for " & sReferencePath.toString & " in Function " & sFunctionName)
				Fn_DSM_UI_ObjectExist=True
		Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"UI_AWS : New Object is Not Exist for " & sReferencePath.toString & " in Function " & sFunctionName)
				Fn_DSM_UI_ObjectExist= False
		End If

		Set objDialog = Nothing 
End Function

'#########################################################################################################
'###
'###    FUNCTION NAME   :  Fn_DSM_UIButton_Operations(sFunctionName, sAction, objDialog, sUIButton)
'###
'###    DESCRIPTION     :   This function is Use to check Existance of given Object
'###
'###    PARAMETERS      :   1.sFunctionName: Valid Function Name
'###			    							2.sAction: Valid Action
'###                                         	3.objDialog : Object dialog
'###											4.sUIButton : Button name
'###	 HISTORY         :   AUTHOR                 DATE        
'###
'###    CREATED BY      :   Amruta Patil       		  16/07/2021     
'###    EXAMPLE         : Fn_DSM_UIButton_Operations("Fn_DSM_UI_ObjectExist", "Click","UIAWindow("Teamcenter RAC (Eclipse)").UIATab("Summary").UIAObject("LeftTab")","TabUIButton")
'#############################################################################################################

Public Function Fn_DSM_UIButton_Operations(sFunctionName, sAction, objDialog, sUIButton)
	Dim objUIButton, sFuncLog, objDeviceReplay
	bGblFailedFunctionName = sFunctionName
	Fn_DSM_UIButton_Operations = False
	'Object Creation
		Set objUIButton = objDialog.UIAButton(sUIButton)
		sFuncLog = sFunctionName + " > Fn_DSM_UIButton_Operations  : [ " &  objDialog.toString & " ] : [ " +  objUIButton.toString + " ] : Action = " & sAction & " : "
	'Verify JavaButton object exists
	If Fn_DSM_UI_ObjectExist("Fn_DSM_UIButton_Operations",objUIButton) = False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & " : FAIL : Does not exist.")
		Set objUIButton = Nothing 
		Exit Function
	End If
	
	Select Case sAction
	
		Case "Click"
			objUIButton.Click	
			wait 5
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & "PASS : Sucessfully clicked on JavaButton.")
			Fn_DSM_UIButton_Operations = True
	End Select
	'Clear memory of JavaButton object.
	Set objDeviceReplay = Nothing
	Set objUIButton = Nothing 
End Function
