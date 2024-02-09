Option Explicit
' *********************************************************	DataShareManager Library  Function List		***********************************************************************
'1. Fn_DSM_Window_Exist()										 This function is Use to setTOProperty of DataShareManager window and check existance of DSM 
'2. Fn_Maximize_DSM_Window()									 This function is Use to maximize the Data Share Manager window
'3. Fn_DSM_Table_Operation()									 This function is use to perform table operations on dsm window
'	
' *********************************************************	End Of List		****************************************************************************************

'#########################################################################################################
'###
'###    FUNCTION NAME   :  Fn_DataShareManager_Exist()
'###
'###    DESCRIPTION     :   This function is Use to setTOProperty of DataShareManager window and check existance of DSM 
'###
'###    PARAMETERS      :   
'###	 HISTORY         :   AUTHOR               		  DATE            Reviwer  
'### 
'###    CREATED BY      :   Vaishali Deshmukh      		  21/01/2022       Amruta P    
'###
'###    EXAMPLE         : call Fn_DataShareManager_Exist()
'#############################################################################################################
Function Fn_DataShareManager_Exist()
	Dim objDesc,objDSM,tabobj,iCount,WShell
	GBL_FAILED_FUNCTION_NAME = "Fn_DataShareManager_Exist"
	Fn_DataShareManager_Exist=False
	Set objDSM = UIAWindow("Teamcenter Data Share")
	Set objDesc = Description.Create
	objDesc("Class Name").value = "UIAWindow"
	Set tabobj = Desktop.ChildObjects(objDesc)
	
	For iCount = 1 To tabobj.Count-1
		If (tabobj(iCount).GetROProperty("name") = "Teamcenter Data Share Manager") Then
			objDSM.SetTOProperty "hwnd", tabobj(iCount).GetROProperty("hwnd")
			objDSM.SetTOProperty "path", tabobj(iCount).GetROProperty("path")
			objDSM.SetTOProperty "name", tabobj(iCount).GetROProperty("name")
			objDSM.SetTOProperty "controltype", tabobj(iCount).GetROProperty("controltype")
			Exit For
		End If
	Next
	If objDSM.Exist(5) then
			set WShell = CreateObject("Wscript.shell")
			WShell.AppActivate "Teamcenter Data Share Manager"
 			objDSM.Highlight
 			WShell.AppActivate "Teamcenter Data Share Manager"
 			wait 2
 			WShell.AppActivate "Teamcenter Data Share Manager"
'			UIAWindow("Teamcenter Data Share").Activate
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Data Share Manager window exist")
			Fn_DataShareManager_Exist=True
			Exit Function
	Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Data Share Manager window not exist")
			Fn_DataShareManager_Exist= False
			Exit Function
	End If
	Set objDSM = Nothing
	Set WShell = Nothing	
End Function

'#########################################################################################################
'###
'###    FUNCTION NAME   :  Fn_Maximize_DSM_Window()
'###
'###    DESCRIPTION     :   This function is Use to maximize the Data Share Manager window
'###
'###    PARAMETERS      :   sAction
'###                                         
'###	 HISTORY         :   AUTHOR                        DATE            Reviwer  
'### 
'###    CREATED BY      :   Vaishali Deshmukh     		  21/01/2022       Amruta P
'###
'###    EXAMPLE         : call Fn_Maximize_DSM_Window()
'#############################################################################################################
Function Fn_Maximize_DSM_Window(sAction)
	Dim objDesc,objDSM,tabobj,iCount,objButton
	GBL_FAILED_FUNCTION_NAME = "Fn_Maximize_DSM_Window"
	Fn_Maximize_DSM_Window=False
	Set objDSM = UIAWindow("Teamcenter Data Share")
	Set objButton = UIAWindow("Teamcenter Data Share").UIAButton("Maximize")
	UIAWindow("Teamcenter Data Share").Activate
	Set objDesc = Description.Create
	objDesc("controltype").Value="Button"
	Set tabobj = objDSM.ChildObjects(objDesc)
	
	For iCount = 0 To tabobj.Count-1
		If (tabobj(iCount).GetROProperty("helptext") = "Click for Main Window") Then
			objButton.SetTOProperty "hwnd",tabobj(iCount).GetROProperty("hwnd")
			objButton.SetTOProperty "path",tabobj(iCount).GetROProperty("path")
			objButton.SetTOProperty "name",tabobj(iCount).GetROProperty("name")
			Exit For
		End If
	Next
			
	Select Case sAction		
			Case "ButtonClick"
					bReturn = Fn_DSM_UIButton_Operations("Fn_Maximize_DSM_Window", "Click",UIAWindow("Teamcenter Data Share"),"Maximize")
					If  bReturn=True Then
					  	  Fn_Maximize_DSM_Window = True
					  	  
					  	  Exit Function
					Else			
						  Fn_Maximize_DSM_Window = False
					  	  Exit Function
					End If
			End Select
	Set objButton = Nothing 
	Set objDSM = Nothing
	Set objDesc = Nothing	
End Function

'#########################################################################################################
'###
'###    FUNCTION NAME   :  Fn_DSM_Table_Operation()
'###
'###    DESCRIPTION     :   This function is Use to perform the table operations on Data Share Manager window
'###
'###    PARAMETERS      :   sAction
'###                                         
'###	 HISTORY         :   AUTHOR                			 DATE            Reviwer  
'### 
'###    CREATED BY      :   Vaishali Deshmukh      		  21/01/2022       Amruta P    
'###
'###    EXAMPLE         : call Fn_DSM_Table_Operation()
'#############################################################################################################
Public Function Fn_DSM_Table_Operation(strAction,sFileName,dicDetails)
	Dim objDSM,objTbale,ColCount,RowCount,iCount,objDesc,objButton,RowName
	GBL_FAILED_FUNCTION_NAME="Fn_DSM_Table_Operation"
   	Fn_DSM_Table_Operation = FALSE
	Set objDSM = UIAWindow("Teamcenter Data Share")
	Set objTbale = UIAWindow("Teamcenter Data Share").UIATable("DSMTable")
	Call Fn_DataShareManager_Exist()
	UIAWindow("Teamcenter Data Share").Activate
	Set objDesc = Description.Create
	objDesc("controltype").Value="Table"
	Set tabobj = objDSM.ChildObjects(objDesc)
	If tabobj.count <> 0 Then
		objTbale.SetTOProperty "hwnd",tabobj(0).GetROProperty("hwnd")
		objTbale.SetTOProperty "path",tabobj(0).GetROProperty("path")
		objTbale.SetTOProperty "name",tabobj(0).GetROProperty("name")
	Else
		Fn_DSM_Table_Operation = FALSE
		Exit Function
	End If
	
	If dicDetails("FileStatus") <> "" Then
				Set objButton = UIAWindow("Teamcenter Data Share").UIAButton("StatusButton")
				
				Set objDesc = Description.Create
				objDesc("controltype").Value="Button"
				Set tabobj = objDSM.ChildObjects(objDesc)
				For iCount = 0 To tabobj.Count-1
					If instr(tabobj(iCount).GetROProperty("helptext"), dicDetails("FileStatus")) Then  
					
						objButton.SetTOProperty "hwnd",tabobj(iCount).GetROProperty("hwnd")
						objButton.SetTOProperty "path",tabobj(iCount).GetROProperty("path")
						objButton.SetTOProperty "name",tabobj(iCount).GetROProperty("name")
						Exit For
					End If
				
				Next
				
				call Fn_DSM_UIButton_Operations("Fn_DSM_Table_Operation", "Click",UIAWindow("Teamcenter Data Share"),"StatusButton")
				wait 2
			
	End If
	Select Case strAction
		Case "DatasetFileExist"
		
					ColCount = UIAWindow("Teamcenter Data Share").UIATable("DSMTable").GetROProperty("columncount")
					RowCount = UIAWindow("Teamcenter Data Share").UIATable("DSMTable").GetROProperty("rowcount")
					 
					 For i = 0 To RowCount-1
					 	For j = 1 To ColCount-1
							RowName = UIAWindow("Teamcenter Data Share").UIATable("DSMTable").GetCellName(i,j)
							If RowName <> "" Then
								If RowName = sFileName Then
								Fn_DSM_Table_Operation = True
								Exit Function
								Else
									If i = RowCount - 1 Then
										Fn_DSM_Table_Operation = False
										Exit Function
									End If
								End If
							End If
						Next
					 Next
			
				 
		Case "Actions"
					If dicDetails("Actions") <> "" Then
						Set objButton = UIAWindow("Teamcenter Data Share").UIATable("DSMTable").UIAButton("Actions")
						Set objDesc = Description.Create
						objDesc("controltype").Value="Button"
						Set tabobj = objDSM.ChildObjects(objDesc)
						For iCount = 0 To tabobj.Count-1
							If (tabobj(iCount).GetROProperty("helptext") = dicDetails("Actions")) Then  
								objButton.SetTOProperty "hwnd",tabobj(iCount).GetROProperty("hwnd")
								objButton.SetTOProperty "path",tabobj(iCount).GetROProperty("path")
								objButton.SetTOProperty "name",tabobj(iCount).GetROProperty("name")
								Exit For
							End If
						Next
						
						ColCount = UIAWindow("Teamcenter Data Share").UIATable("DSMTable").GetROProperty("columncount")
						RowCount = UIAWindow("Teamcenter Data Share").UIATable("DSMTable").GetROProperty("rowcount")
				 
						 For i = 0 To RowCount-1
						 	For j = 1 To ColCount-1
						 	 	RowName = UIAWindow("Teamcenter Data Share").UIATable("DSMTable").GetCellName(i,j)
								If RowName <> "" Then
									If RowName = sFileName Then
										'Rownum = i
										If dicDetails("Actions")="Click to close this completed transaction" Then
											UIAWindow("Teamcenter Data Share").UIATable("DSMTable").ClickCell i,6,60,10,"LEFT"
											wait 2
											Fn_DSM_Table_Operation = True
											Exit Function
										ElseIf dicDetails("Actions")="Click to open this file" Then
											UIAWindow("Teamcenter Data Share").UIATable("DSMTable").ClickCell i,6,40,10,"LEFT"
											Fn_DSM_Table_Operation = True
											Exit Function
										ElseIf dicDetails("Actions")="Click to open containing folder for this file" Then
											UIAWindow("Teamcenter Data Share").UIATable("DSMTable").ClickCell i,6,20,10,"LEFT"
											Fn_DSM_Table_Operation = True
											Exit Function
										End If
									Else
										If i = RowCount - 1 Then
											Fn_DSM_Table_Operation = False
											Exit Function
										End If
									End If
								End If
							Next
						 Next
				End If
	End Select
	Set objDSM = Nothing
	Set objButton = Nothing
	Set objTbale = Nothing
	Set objDesc = Nothing
End  Function

