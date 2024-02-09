' global variables for recursive function Fn_UI_getTreeIndex - Koustubh
Public WEB_DEFAULT_TIMEOUT, WEB_MICRO_TIMEOUT, WEB_MIN_TIMEOUT, WEB_MAX_TIMEOUT, WEB_MINLESS_TIMEOUT, WEB_MICROLESS_TIMEOUT

WEB_MAX_TIMEOUT = 240 'time in seconds
WEB_DEFAULT_TIMEOUT = 10 'time in seconds
WEB_MIN_TIMEOUT = 5 'time in seconds
WEB_MINLESS_TIMEOUT = 3
WEB_MICROLESS_TIMEOUT = 2
WEB_MICRO_TIMEOUT = 1 'time in second

' *********************************************************	UI_WebLibrary  Function List		***********************************************************************
'1.  Fn_Web_UI_ObjectExist()												This function is Use to check Existance of given Object
'2   Fn_Web_UI_ObjectEnable()											This function is Use to check  given Object Enable or Not i.e. State of Object
'3.  ExitFromWeb_UI(sProcessToKill)									 This function  is used to  exit from Process  after doing kill  active process mentioned in parameter .
' 4 .  Fn_Web_UI_Button_Click()												This function is used to click the Button.
'5.  Fn_Web_UI_WebEdit_Set()                               ---			This function is Use to Set value of WebEdit Box
'6		Fn_Web_UI_ObjectVisible												This function is Use to check  given Object Visible or Not i.e. Visibility State of Object
'7		Fn_Web_UI_WinButton_Click										Function to verify WinButton is enabled and to Click the mouse button at X,Y Co-ordinates. 
'8. Fn_Web_UI_List_Select()														This function is used to select the element from the List
'9. Fn_Web_UI_CheckBox_Set()											This function is used to Select or DeSelect the CheckBox.
'10. Fn_Web_UI_WebElement_Click          						This function is used to click the Web element Like "Menu"
'11. Fn_Web_UI_Link_Click              										This function is used to click the Link.
'12. Fn_Web_UI_Image_Click             									This function is used to click the Image
'13. Fn_WEB_UI_Object_SetTOProperty							This function  is used to Set TO Property For given Object
'14. Fn_WebUI_ImageOperations										This function  is used to Perform Operations On Image
'15. Fn_WebUI_TableRowIndex
'16. Fn_WebUI_TableColumnIndex
'17. Fn_SISW_WebUI_WebListItemExist
'18. Fn_Web_UI_WebWinEdit_GetValue()						This function will return value from a parametrised WinEdit on a browser
'19. Fn_Web_UI_WebEdit_SetExt()
'20. Fn_Web_UI_WebEdit_GetValue()
'21. Fn_Web_UI_Button_ClickExt()
' *********************************************************	End Of List		****************************************************************************************

'#########################################################################################################
'###
'###    FUNCTION NAME   :  Fn_Web_UI_ObjectExist(sFunctionName, sReferencePath)
'###
'###    DESCRIPTION     :   This function is Use to check Existance of given Object
'###
'###    PARAMETERS      :   1.sFunctionName: Valid Function Name
'###			    							2.sReferencePath: Valid reference path
'###                                         
'###    Function Calls  :  Fn_WriteLogFile()
'###
'###	 HISTORY         :   AUTHOR                 DATE        VERSION
'###
'###    CREATED BY      :   Sagar       		  06/07/2010      1.0
'###
'###    REVIWED BY      :   Sameer		 		  06/07/2010	  1.0
'###
'###    EXAMPLE         : Fn_Web_UI_ObjectExist("Fn_Web_UI_ObjectExist", "Browser("Teamcenter Login").Page("Teamcenter Login") ")
'#############################################################################################################

Function Fn_Web_UI_ObjectExist(sFunctionName, sReferencePath)
		Dim objDialog
		Set objDialog = sReferencePath

		If objDialog.Exist Then			
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"UI_Web : New Object is Exist for " & sReferencePath.toString & " in Function " & sFunctionName)
				Fn_Web_UI_ObjectExist=True
		Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"UI_Web : New Object is Not Exist for " & sReferencePath.toString & " in Function " & sFunctionName)
				Fn_Web_UI_ObjectExist= False
		End If

		Set objDialog = Nothing 
End Function

'#########################################################################################################
'###
'###    FUNCTION NAME   :  Fn_Web_UI_ObjectEnable(sFunctionName, sReferencePath)
'###
'###    DESCRIPTION     :   This function is Use to check  given Object Enable or Not i.e. State of Object
'###
'###    PARAMETERS      :   1.sFunctionName: Valid Function Name
'###			    							2.sReferencePath: Valid reference path
'###                                         
'###    Function Calls  :  Fn_WriteLogFile()
'###
'###	 HISTORY         :   AUTHOR                 DATE        VERSION
'###
'###    CREATED BY      :   Sagar       		  06/07/2010      1.0
'###
'###    REVIWED BY      :   Sameer		 		  06/07/2010	  1.0
'###
'###    EXAMPLE         : 		Fn_Web_UI_ObjectEnable("Fn_Web_UI_ObjectEnable", "Browser("Teamcenter Login").Page("Teamcenter Login") ")
'#############################################################################################################

Function Fn_Web_UI_ObjectEnable(sFunctionName, sReferencePath)
  Dim objDialog
  Set objDialog = sReferencePath
  If objDialog.CheckProperty("Enabled","1")= True OR objDialog.CheckProperty("Enabled",True)=True  OR objDialog.CheckProperty("disabled","0") = True Then 
      Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"UI_Web : New Object is Enabled for " & sReferencePath.toString & " in Function " & sFunctionName)
      Fn_Web_UI_ObjectEnable=True
    Else
      Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"UI_Web : New Object is Not Enabled for " & sReferencePath.toString & " in Function " & sFunctionName)
      Fn_Web_UI_ObjectEnable= False
  End If
  Set objDialog = Nothing 
End Function

'#################################################################################################################
'###    FUNCTION NAME   :    ExitFromWeb_UI(sProcessToKill)
'###
'###    DESCRIPTION     :   This function  is used to  exit from Process  after doing kill  active process mentioned in parameter .
'###
'###    PARAMETERS      :   sProcessToKill - list of processes seperated by colon (:) 
'###             
'###    Function Calls  :   Fn_KillProcess(sProcessToKill)
'###
'###    HISTORY         :   
'###
'###    CREATED BY      :   Sagar Shivade
'###
'###    REVIWED BY      :      Sameer Chitnis
'###
'###    MODIFIED BY     :
'###    EXAMPLE         :    ExitFromWeb_UI("Teamcenter.exe:java.exe:jucheck.exe:jusched.exe:javaw.exe")
'###												ExitFromWeb_UI("")
'################################################################################################################

Function ExitFromWeb_UI(sProcessToKill)
   
	' call function killProcess .
'   Call Fn_Web_KillProcess(sProcessToKill)
'   'exit from current test.
'   'Code added by Archana
'   Call Fn_UpdateLogFiles("UI_Web : FAIL|Exiting From Web_UI", "FAIL:Failed From Web_UI")
	'ExitTestIteration 

End Function



'#######################################################################################################
'###    FUNCTION NAME   :   Fn_Web_UI_Button_Click()
'###
'###    DESCRIPTION     :   This function is used to click the Button.
'###
'###    PARAMETERS      :   sFunctionName : Valid Function name, 
'###                       						 objPage : Valid Dialog box object name,
'###                        					sWebButton   : Valid Button name
'###
'###    Function Calls  :   Fn_WriteLogFile()
'###
'###    HISTORY         :   AUTHOR         DATE          VERSION
'###
'###    CREATED BY      :   Sagar          7/07/2010    1.0
'###
'###    REVIWED BY      :	Sameer			7/07/2010 1.0
'###
'###    MODIFIED BY     :   
'###    EXAMPLE         :   
'#######################################################################################################

Function Fn_Web_UI_Button_Click(sFunctionName, objPage, sWebButton)

	Dim objWebButton, bReturn,objButton,iCounter,objPageChld,bHgt
	'Object Creation
Fn_Web_UI_Button_Click=False
If Instr(objPage.toString,"ButtunPanel") > 0  or Instr(objPage.toString,"WebButtons") > 0 Then

	Set objButton=Description.Create
	objButton("micClass").value="WebButton"
	objButton("Name").value=sWebButton
	Set objPageChld= Browser("TeamcenterWeb").Page("MyTeamCenter").ChildObjects(objButton)
	
	For iCounter= 0 To objPageChld.count-1
		If Instr(objPage.toString,"WebButtons") > 0 and sWebButton = "OK" Then
			iCounter = iCounter + 1
			bHgt=objPageChld(iCounter).GetROProperty("height")
			If bHgt>0 Then
				objPageChld(iCounter).click
				Fn_Web_UI_Button_Click=True
			End If
		ElseIf Instr(objPage.toString,"ButtunPanel") > 0 and sWebButton = "Save and Check-In" Then
			iCounter = iCounter + 1
			bHgt=objPageChld(iCounter).GetROProperty("height")
			If bHgt>0 Then
				objPageChld(iCounter).click
				Fn_Web_UI_Button_Click=True
			End If
		Else
			bHgt=objPageChld(iCounter).GetROProperty("height")
			If bHgt>0 Then
				objPageChld(iCounter).click
				Fn_Web_UI_Button_Click=True
			End If
		End If
	Next

Else
		Set objWebButton = objPage.WebButton(sWebButton)
		bReturn= Fn_Web_UI_ObjectExist("Fn_Web_UI_ObjectCreate", objWebButton)
		If  bReturn=True Then
							 If objWebButton.CheckProperty("disabled","0") AND   objWebButton.CheckProperty("visible",True)Then 
								  objWebButton.Click	
								  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "UI_Web : Sucessfully clicked on WebButton " & sWebButton & " of Function " & sFunctionName)
								   Fn_Web_UI_Button_Click = True
							Else			
							'Report error/message when Web Button object is disable.
							  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "UI_Web : WebButton " &sWebButton &"  is not enabled of Function " & sFunctionName)
							  Fn_Web_UI_Button_Click = False
							  Call ExitFromWeb_UI("")
						End If
			
			 
			'Report error/message when WebButton object does not exists.
			Else
						   Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "UI_Web : WebButton " &sWebButton & "  does not exist of Function " &sFunctionName )
						   Fn_Web_UI_Button_Click = False
						   Call ExitFromWeb_UI("")
		End If
	
		'Clear memory of WebButton object.
		Set objWebButton = Nothing 
		
End If

End Function

'#######################################################################################################
'###    FUNCTION NAME   :   Fn_Web_UI_WebEdit_Set()
'###
'###    DESCRIPTION     :   This function is used to click the Button.
'###
'###    PARAMETERS      :   sFunctionName : Valid Function name, 
'###                       						 objPage : Valid Dialog box object name,
'###                        					sWinEdit   : Valid Edit box
'###                        					sValue   : Valid value to set
'###
'###    Function Calls  :   Fn_WriteLogFile()
'###
'###    HISTORY         :   AUTHOR         DATE          VERSION
'###
'###    CREATED BY      :   Sagar          14/07/2010    1.0
'###
'###    REVIWED BY      :	Sameer			14/07/2010 1.0
'###
'###    MODIFIED BY     :   
'###    EXAMPLE         :   
'#######################################################################################################

Public Function Fn_Web_UI_WebEdit_Set(sFunctionName, objPage, sWinEdit, sValue)
		Dim objWebEdit, bReturn,objMDR
		Set objWebEdit = objPage.WebEdit(sWinEdit)
		'creating object of Mercury device replay
		Set objMDR = CreateObject("Mercury.DeviceReplay")
		bReturn= Fn_Web_UI_ObjectExist("Fn_Web_UI_ObjectCreate", objWebEdit)

		If  bReturn=True Then
				If objWebEdit.CheckProperty("disabled","0") AND   objWebEdit.CheckProperty("visible",True)Then 
'						objWebEdit.Object.focus
					     objWebEdit.Set ""
						 objWebEdit.Object.focus
                        objMDR.SendString sValue
                        Wait 0, 300
						Fn_Web_UI_WebEdit_Set = True
						Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"UI_Web :PASS: Successfully WinEdit "+CStr(objWebEdit.toString)+" Box  Set with Value "+CStr(sValue)+" on Function "+CStr(sFunctionName)+"")
				Else
						'Report error/message when WebEditBox object is disable.
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "UI_Web : Fail : WebEdit " +CStr(objWebEdit.toString)+"  is not enabled of Function " & sFunctionName)
						Fn_Web_UI_WebEdit_Set = False
						Call ExitFromWeb_UI("")
				End If
		Else
				Fn_Web_UI_WebEdit_Set = False
				Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"UI_Web :FAIL: Failed to Set with Value "+CStr(sValue)+" in WinEdit "+CStr(objWebEdit.toString)+" Box on Function "+CStr(sFunctionName)+" ")
				Call ExitFromWeb_UI("")
		End If

		'Clear memory of WebButton object.
		Set objWebEdit = Nothing
		Set objMDR =Nothing
End Function


'#########################################################################################################
'###
'###    FUNCTION NAME   :  Fn_Web_UI_ObjectVisible(sFunctionName, sReferencePath)
'###
'###    DESCRIPTION     :   This function is Use to check  given Object Visible or Not i.e. VisibleState of Object
'###
'###    PARAMETERS      :   1.sFunctionName: Valid Function Name
'###			    							2.sReferencePath: Valid reference path
'###                                         
'###    Function Calls  :  Fn_WriteLogFile()
'###
'###	 HISTORY         :   AUTHOR                 DATE        VERSION
'###
'###    CREATED BY      :   Sagar       		  08/07/2010      1.0
'###    CREATED BY      :   Sagar       		  08/07/2010      1.0
'###    REVIWED BY      :   Sameer		 		  08/07/2010	  1.0
'###
'###    EXAMPLE         : 		Fn_Web_UI_ObjectVisible("Fn_Web_UI_ObjectEnable", "Browser("Teamcenter Login").Page("Teamcenter Login") ")
'#############################################################################################################

Function Fn_Web_UI_ObjectVisible(sFunctionName, sReferencePath)
		Dim objDialog
		Set objDialog = sReferencePath
		
		If objDialog.CheckProperty("Visible","1")= True OR objDialog.CheckProperty("Visible",True) = True Then	
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"UI_Web : New Object is Visible  for " & objDialog.toString & " in Function " & sFunctionName)
						Fn_Web_UI_ObjectVisible=True
				Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"UI_Web : New Object is Not Visible for " & objDialog.toString & " in Function " & sFunctionName)
						Fn_Web_UI_ObjectVisible= False
		End If

		Set objDialog = Nothing 
End Function

'##################################################################################################################################################
'###    FUNCTION NAME   :   Fn_Web_UI_WinButton_Click(sFunctionName, objJavaDialog, sWinButton,iXValue,iYValue,coMicButton) 
'###
'###    DESCRIPTION     :   Function to verify WinButton is enabled and to Click the mouse button at X,Y Co-ordinates. 
'###
'###    PARAMETERS      :   sFunctionName : Valid Function Name, 
'###			    		objJavaDialog : Valid Dialog Path,
'###					    sWinButton    : Valid Button Name,
'###					    iXValue	  : Valid X Co-ordiate value,
'###			    		iYValue	  : Valid Y Co-ordinate Value,
'###			    		coMicButtonToClick : Valid Mouse button to be Clicked 
'###                        
'###    Function Calls  :   Fn_WriteLogFile()
'###
'###    HISTORY         :   AUTHOR      						DATE        VERSION
'###
'###    CREATED BY      :   Sagar Shivade				15/07/10		1.0
'###
'###    REVIWED BY      :	Sameer 								15/07/10		1.0
'###	
'###    EXAMPLE         :   Fn_Web_UI_WinButton_Click(Fn_Web_FileDownloadOperations,"File Download","Open","","",micLeftBtn)
'##################################################################################################################################################


Function Fn_Web_UI_WinButton_Click(sFunctionName, objDialog, sWinButton,iXValue,iYValue,coMicButton)
	Dim objWinButton,bReturn
'Object Creation
	Set objWinButton = objDialog.WinButton(sWinButton)

'Verify  WinButton object exists
	bReturn= Fn_Web_UI_ObjectExist("Fn_Web_UI_WinButton_Click", objWinButton)
	If  bReturn=True Then
		'Verify Object is enabled 
				If objDialog.CheckProperty("Enabled","1")= True OR objDialog.CheckProperty("Enabled",True) = True Then	
			
							If  iXValue <> "" AND iYValue <> "" Then
										'Click the mouse button at X,Y Co-ordinates
										objWinButton.Click iXValue,iYValue,coMicButton
										'log on success
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),  "Sucessfully clicked on WinButton" & sWinButton &" at Co-ordinates " &  iXValue &"," &iYValue & " of Function " & sFunctionName)
							  Else
										objWinButton.Click
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Clicked on " & sWinButton & "WinButton of Function " & sFunctionName)
							 End If
			
					  Fn_Web_UI_WinButton_Click = True			
			'Report Error when WinButton object is disable.
					Else
						   Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sWinButton & "WinButton  is disabled of Function " &sFunctionName)
						   Fn_Web_UI_WinButton_Click = False
						 Call ExitFromWeb_UI("")
					End If
			'Report Error when WinButton object does not exists.
	Else
						  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sWinButton & "WinButton does not exist of Function " &sFunctionName)
						  Fn_Web_UI_WinButton_Click = False
						  Call ExitFromWeb_UI("")
	End If

Set objWinButton = Nothing
End Function

'########################################################################################################################################

'##############################################################################################################################################
'###    FUNCTION NAME   :   Fn_Web_UI_List_Select()
'###
'###    DESCRIPTION     :   This function is used to select the element from the List
'###
'###    PARAMETERS      :   sFunctionName	: Valid Function Name,
'###			    							objJavaDialog	: Valid Dialog/Window Path,
'###			    							sWebList		: Valid Weblist Name,
'###			    							sElementToSelect 	: Valid Element to be selected 
'###
'###    Function Calls  :   Fn_WriteLogFile (To report errors )
'###
'###    HISTORY         :   AUTHOR            DATE        VERSION
'###
'###    CREATED BY      :   Sagar 			16/07/10		1.0
'###
'###    REVIWED BY      :   Sameer 			16/07/10		1.0
'###
'###    MODIFIED BY     :   NA
'###    EXAMPLE         :   Call Fn_Web_UI_List_Select("Fn_Web_PasteAs", Browser("Teamcenter").Page("MyTeamcenter").WebTable("NewFolder"), "Type","Folder")
'###############################################################################################################################################

Function Fn_Web_UI_List_Select(sFunctionName, objDialog, sWeblist,sElementToSelect)

		' Variable declaration
		Dim bReturn																															
		Dim objWeblist

		'Object Creation
		Set objWeblist = objDialog.Weblist(sWeblist)														
		bReturn= Fn_Web_UI_ObjectExist("Fn_Web_UI_ObjectCreate", objWeblist)

		If  bReturn=True Then
				If objWeblist.CheckProperty("disabled","0") AND   objWeblist.CheckProperty("visible",True)Then 
						objWeblist.Select sElementToSelect
						'  Report message of Selected the Element from Weblist succesfully.				
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),  "Sucessfully Selected element " & sElementToSelect &" of Web List " & sWeblist & " of Function " &sFunctionName)
						Fn_Web_UI_List_Select = True
				
				Else
						' Report error when object is disable.
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Weblist " & sWeblist &" is disable of Function " &sFunctionName)
						Fn_Web_UI_List_Select = False
						Call ExitFromUI("")
				End If
		Else
			' Report error when Web List object does not exists.
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Weblist " &sWeblist & " does not exist of Function " &sFunctionName)
			Fn_Web_UI_List_Select = False
			Call ExitFromUI("")
		End If

		Set objWeblist=Nothing

End Function																																	

'#######################################################################################


'##########################################################################################################################################################################
'###    FUNCTION NAME   :   Fn_Web_UI_CheckBox_Set(sFunctionName, objDialog, sWebCheckBox, sStatus)
'###
'###    DESCRIPTION     :   This function is used to Select or DeSelect the CheckBox.
'###
'###    PARAMETERS      :   sFunctionName - Valid function name
'###                        					objDialog - Valid Java Dialog Name
'###                       						 sWebCheckBox- Vaild CheckBox Name
'###                        					sStatus - Valid Status( ON or OFF)
'###                       
'###    Function Calls  :   Fn_WriteLogFile()
'###
'###    HISTORY         :   AUTHOR      	DATE        	VERSION
'###
'###    CREATED BY      :   Sagar Shivade    19/07/2010 		1.0
'###
'###    REVIWED BY      :   Sameer         		23/07/2010		1.0
'###
'###
'###    EXAMPLE         :   Call Fn_Web_UI_CheckBox_Set("Fn_Web_UI_CheckBox_Set" ,Browser("Teamcenter").Page("MyTeamcenter").WebTable("NewItem").WebTable("ItemObjects"),"CreateAlternateID", "OFF") 
'###########################################################################################################################################################################


Function Fn_Web_UI_CheckBox_Set(sFunctionName, objDialog, sWebCheckBox, sStatus)

Dim objWebCheckBox,bReturn

Set objWebCheckBox = objDialog.WebCheckBox(sWebCheckBox)
    
		bReturn= Fn_Web_UI_ObjectExist("Fn_Web_UI_CheckBox_Set", objWebCheckBox)

		If  bReturn=True Then
				If objWebCheckBox.CheckProperty("disabled","0") AND   objWebCheckBox.CheckProperty("visible",True)Then 
			
							'Check weather checkbox is unchecked or not.
							If  UCase(sStatus) = "ON"  Then	
									objWebCheckBox.Set "ON"
									'log the Success Result
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "UI_Web :Sucessfully Selected WebCheckBox " & sWebCheckBox & " of Function " & sFunctionName)
									Fn_Web_UI_CheckBox_Set = True
							ElseIf UCase(sStatus) = "OFF" Then
									objWebCheckBox.Set "OFF"
									'log the Success Result
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "UI_Web : WebCheckBox " & sWebCheckBox & "is DeSelected of Function " & sFunctionName)
									Fn_Web_UI_CheckBox_Set = True
							End If
				Else 
					'log the Failure Result
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),  "UI_Web :WebCheckBox " &sWebCheckBox & " is Disable  of Function " & sFunctionName)
					Fn_Web_UI_CheckBox_Set = False 
					Call ExitFromWeb_UI("")
				End If
		Else
	
				'log the Failure Result
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),  "UI_Web :WebCheckBox " & sWebCheckBox & " does not Exist of Function " & sFunctionName)	
				Fn_Web_UI_CheckBox_Set = False
				Call ExitFromWeb_UI("")
	End If

Set objWebCheckBox = Nothing 
End Function

'#######################################################################################################
'###    FUNCTION NAME   :   Fn_Web_UI_WebElement_Click(sFunctionName, objDialog, sWebElement,  iXValue,iYValue,coMicButton)
'###
'###    DESCRIPTION     :   This function is used to click the element
'###
'###    PARAMETERS      :   sFunctionName : Valid Function name, 
'###                       						 objDialog : Valid Dialog box object ,
'###                        					sWebElement   : Valid element name
'###                        					iXValue   : Valid X axis Cordinate Number
'###                        					iYValue   : Valid Y axis Cordinate Number
'###                        					coMicButton   : Valid Mouse button Constant 
'###
'###    Function Calls  :   Fn_WriteLogFile()
'###
'###    HISTORY         :   AUTHOR         DATE          VERSION
'###
'###    CREATED BY      :   Sagar          28/07/2010    1.0
'###
'###    REVIWED BY      :	Mahendra		28/07/2010 		1.0
'###
'###    EXAMPLE         :   Call Fn_Web_UI_WebElement_Click("Fn_Web_QuickLinkOperations", Browser("Teamcenter").Page("MyTeamcenter").WebTable("MenuTable"), "MenuItem", "", "", "")
'#######################################################################################################

Public Function Fn_Web_UI_WebElement_Click(sFunctionName, objDialog, sWebElement, iXValue,iYValue,coMicButton)

	Dim objWebElement, bReturn

	'Set an list object on variable        
	Set objWebElement = objDialog.WebElement(sWebElement)
    bReturn = Fn_Web_UI_ObjectExist("Fn_Web_UI_WebElement_Click", objWebElement)
        'checking List object exist or not
		If bReturn = True Then  

		   'Adding Syncronization point and Check Property disabled and visible
            If Fn_Web_UI_ObjectVisible("Fn_Web_UI_WebElement_Click", objWebElement) = True Then

							If  iXValue <> "" AND iYValue <> "" Then
										'Click the mouse button at X,Y Co-ordinates
										objWebElement.Click iXValue,iYValue,coMicButton
										'log on success
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),  "UI_Web : Sucessfully clicked on WebElement [" & sWebElement &"] with Co-ordinates " &  iXValue &"," &iYValue & " of Function " & sFunctionName)
							  Else
										objWebElement.Click
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "UI_Web : Clicked on " & sWebElement & " WebElement of Function " & sFunctionName)
							 End If

						' Return True from Function
						Fn_Web_UI_WebElement_Click = True
				 Else
						 'log the failure when list not enabled
						 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "UI_Web : "&sWebElement & " WebElement is not visible of Function " & sFunctionName)
						 'Return False from function
						  Fn_Web_UI_WebElement_Click = False
						  Call ExitFromWeb_UI("")
				End If
		Else
		        'log the failure when list does not exist
                Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"UI_Web : "&sWebElement & " WebElement does not exist of Function " & sFunctionName)
				'Return False from function
				Fn_Web_UI_WebElement_Click = False
                Call ExitFromWeb_UI("")
	End If

	 Set objWebElement = Nothing
End Function

'#######################################################################################################
'###    FUNCTION NAME   :   Fn_Web_UI_Link_Click(sFunctionName, objDialog, sLink,  iXValue,iYValue,coMicButton)
'###
'###    DESCRIPTION     :   This function is used to click on link
'###
'###    PARAMETERS      :   sFunctionName : Valid Function name, 
'###                       						 objDialog : Valid Dialog box object ,
'###                        					sLink   : Valid Link name
'###                        					iXValue   : Valid X axis Cordinate Number
'###                        					iYValue   : Valid Y axis Cordinate Number
'###                        					coMicButton   : Valid mouse button Constant 
'###
'###    Function Calls  :   Fn_WriteLogFile()
'###
'###    HISTORY         :   AUTHOR         DATE          VERSION
'###
'###    CREATED BY      :   Sagar          28/07/2010    1.0
'###
'###    REVIWED BY      :	Mahendra		28/07/2010 		1.0
'###
'###    MODIFIED BY     :   
'###    EXAMPLE         :   Call Fn_Web_UI_Link_Click("Fn_Web_QuickLinkOperations", Browser("Teamcenter").Page("MyTeamcenter").WebTable("MenuTable"), "MenuItem", "", "", "")
'#######################################################################################################

Public Function Fn_Web_UI_Link_Click(sFunctionName, objDialog, sLink,  iXValue,iYValue,coMicButton)

	Dim objLink, bReturn

	'Set an list object on variable        
	Set objLink = objDialog.Link(sLink)
    bReturn = Fn_Web_UI_ObjectExist("Fn_Web_UI_Link_Click", objLink)
        'checking List object exist or not
		If bReturn = True Then  
		   'Adding Syncronization point and Check Property disabled and visible
            If  Fn_Web_UI_ObjectVisible("Fn_Web_UI_Link_Click", objLink) = True Then
							If  iXValue <> "" AND iYValue <> "" Then
										'Click the mouse button at X,Y Co-ordinates
										objLink.Click iXValue,iYValue,coMicButton
										'log on success
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),  "UI_Web : Sucessfully clicked on Link [" & sLink &"] at Co-ordinates " &  iXValue &"," &iYValue & " of Function " & sFunctionName)
							  Else
										objLink.Click
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "UI_Web : Clicked on " & sLink & " Link of Function " & sFunctionName)
							 End If
						' Return True from Function
						Fn_Web_UI_Link_Click = True
				 Else
					 'log the failure when list not enabled
					 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "UI_Web : "&sLink & " Link is disabled in Function " & sFunctionName)
					 'Return False from function
					  Fn_Web_UI_Link_Click = False
					  Call ExitFromWeb_UI("")
				End If
		Else
		        'log the failure when list does not exist
                Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"UI_Web : "&sLink & " Link does not exist of Function " & sFunctionName)
				'Return False from function
				Fn_Web_UI_Link_Click = False
                Call ExitFromWeb_UI("")
	End If

	 Set objLink = Nothing
End Function
'############################################################################################################


'######################################################################################################
'###    FUNCTION NAME   :   Fn_Web_UI_Image_Click()
'###
'###    DESCRIPTION     :   This function is used to click on image.
'###
'###    PARAMETERS      :   sFunctionName : Valid Function name, 
'###                       						 objPage : Valid Dialog box object 
'###                        					sWebImage   : Valid Image name
'###
'###    Function Calls  :   Fn_WriteLogFile()
'###
'###    HISTORY         :   AUTHOR         DATE          VERSION
'###
'###    CREATED BY      :   Sagar          23/07/2010    		1.0
'###
'###    REVIWED BY      :	Mahendra			23/07/2010 		1.0
'###
'###    EXAMPLE         :  Call   Fn_Web_UI_Image_Click("Fn_Web_UI_Image_Click", Browser("Teamcenter").Page("MyTeamcenter").WebTable("QuickSearch"), "Go")
'######################################################################################################

Function Fn_Web_UI_Image_Click(sFunctionName, objPage, sWebImage)
	' Variable Initialization
	Dim objWebImage, bReturn
	'Object Creation
	Set objWebImage = objPage.Image(sWebImage)
	bReturn= Fn_Web_UI_ObjectExist("Fn_Web_UI_Image_Click", objWebImage)
	If  bReturn=True Then
						 If   Fn_Web_UI_ObjectVisible("Fn_Web_UI_Image_Click", objWebImage) = True Then 
								  objWebImage.Click	
								  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "UI_Web : Sucessfully clicked on Image " & sWebImage & " of Function " & sFunctionName)
								   Fn_Web_UI_Image_Click = True
						Else			
						'Report error/message when WebImage object is disable.
								  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "UI_Web : Image " & sWebImage &"  is not visible of Function " & sFunctionName)
								  Fn_Web_UI_Image_Click = False
								  Call ExitFromWeb_UI("")
					End If
		Else
						'Report error/message when WebImage object does not exists.
					   Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "UI_Web : Image " & sWebImage & "  does not exist of Function " &sFunctionName )
					   Fn_Web_UI_Image_Click = False
					   Call ExitFromWeb_UI("")
	End If

	'Clear memory of WebImage object.
	Set objWebImage = Nothing 

End Function

'#################################################################################################################
'###    FUNCTION NAME   :    Fn_WEB_UI_Object_SetTOProperty(sFunctionName,objDialog,sProperty,sPropValue)
'###
'###    DESCRIPTION     :   This function  is used to Set TO Property For given Object
'###
'###    PARAMETERS      :   sFunctionName - Valid function name
'###                        					objDialog - Valid  Dialog Object
'###                              				sProperty-Valid Property Name
'###                        					sPropValue-New Property value
'###                        
'###    Function Calls  :   Fn_WriteLogFile()
'###
'###    HISTORY         :   AUTHOR                   DATE        VERSION
'###
'###    CREATED BY      :   Sagar Shivade      			29/7/10 	1.0  
'###
'###    REVIWED BY      :   Mahendra						29/7/10		1.0
'###
'###    EXAMPLE         :Call Fn_WEB_UI_Object_SetTOProperty("Fn_UI_Object_SetTOProperty",Browser("Teamcenter").Link("Logout"),"text","Login")
'###										
'################################################################################################################

Function Fn_WEB_UI_Object_SetTOProperty(sFunctionName,objDialog,sProperty,sPropValue)
		Dim objSetTOProperty
		Set objSetTOProperty=objDialog

		'Checking existance of object

		objSetTOProperty.SetTOProperty sProperty,sPropValue
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"UI_Web :The Property "&sProperty &" value is set as " & cStr(sPropValue) &" for " & objSetTOProperty.toString & "  of Function " &sFunctionName)
		Fn_WEB_UI_Object_SetTOProperty=True
	
		Set objSetTOProperty = Nothing 
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_WebUI_ImageOperations
'@@
'@@    Description				 :	Function Used to Perform Operations on Image
'@@
'@@    Parameters			   :	1.StrFunctionName:Function Name
'@@												 2.StrAction:Action Name
'@@												 3.ObjPath:Dialog Path
'@@												 4.StrImgName:Image Name
'@@
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	Object Should be Selected 
'@@
'@@    Examples					:	Call Fn_WebUI_ImageOperations("","Verify",Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("Object").WebTable("ObjectInfo"),"screw")
'@@
'@@	   History					 	:	
'@@													Developer Name												Date						Rev. No.						Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@												Sandeep Navghane									11-May-2011						1.0																								Sunny Ruparel
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Function Fn_WebUI_ImageOperations(StrFunctionName,StrAction,ObjPath,StrImgName)
	Dim objDialog,Img,objChld,iCount,StrFileName,StrImg
	Set objDialog=ObjPath
	Fn_WebUI_ImageOperations=False
	Set Img=Description.Create
	Img("micclass").value="Image"
	Img("html tag").value="IMG"
	Set objChld=objDialog.ChildObjects(Img)
	Select Case StrAction
		Case "Verify"
			For iCount=0 To objChld.count-1
					StrFileName=objChld(iCount).GetROProperty("file name")
					StrImg=Split(StrFileName,".")
					If LCase(StrImg(0))=LCase(StrImgName) Then
						Fn_WebUI_ImageOperations=True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"UI_Web :The Image "&StrImgName &" Exist on " & ObjPath.toString & "  of Function " &StrFunctionName)
						Exit For
					End If
			Next
	End Select
	Set objDialog=Nothing
	Set Img=Nothing
	Set objChld=Nothing
End Function
'*********************************************************		Function to Get Table Node Index into***********************************************************************

'Function Name		:				Fn_WebUI_TableRowIndex

'Description			 :		 		  This function is used to get the BOM Table Node Index.

'Parameters			   :	 			1.  objTable : Table Object
'									    			 2.  sNodeName:Name of the Node to retrieve Index for.
'									    			 3.  sColName : Column Name ( optional )
											
'Return Value		   : 				 Node index

'Pre-requisite			:		 		Table should be displayed in web browser .

'Examples				:				 Fn_WebUI_TableRowIndex(Browser("TeamcenterStructureWeb").Page("TeamcenterWebStructure").WebTable("BOMTable"), "000053/A;1-top (View):000093/A;1-sub (View)", "")
'Examples				:				 Fn_WebUI_TableRowIndex(Browser("TeamcenterStructureWeb").Page("TeamcenterWebStructure").WebTable("BOMTable"), "000015/A;1-top (View):000016/A;1-sub (View) @2:000305/A;1-asm @3", "")

'History:
'										Developer Name				Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Kaustubh\Sandeep			12-May-2011			1.0
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Kaustubh					12-May-2011			1.0					Changed logic to handle multiple occurrences.
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_WebUI_TableRowIndex(objTable, sNodeName, sColName)
		Dim iRowCnt, iCounter, aSubElement, jCounter, bFlag
		Dim iInstanceCnt, aNode, sNode
		Fn_WebUI_TableRowIndex = -1
		bFlag = False
		If sColName = "" Then
			sColName =  Fn_WebUI_TableColumnIndex(objTable,"Name")
		else
			sColName = Fn_WebUI_TableColumnIndex(objTable, sColName)
		End If
		If Fn_Web_UI_ObjectExist("Fn_WebUI_TableRowIndex", objTable) = False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Fn_WebUI_TableRowIndex: Specified table does not exist in web browser.")
			Exit function
		End If
		iRowCnt = objTable.RowCount
		aSubElement = Split(sNodeName, ":", -1, 1)
		jCounter = 0
		sNode = ""
		For iCounter = 1 To iRowCnt 
            ' For the Node Hierarchy of an Element
			If sNode = "" then
				 iInstanceCnt = 1
				aNode = split(Trim(aSubElement(jCounter)),"@")
				If UBound(aNode) = 1 Then
					iInstanceCnt = cInt(aNode(1))
				End If
				sNode = trim( aNode(0) )
			end if
			If Trim(objTable.GetCellData(iCounter, sColName)) = sNode Then
				iInstanceCnt = iInstanceCnt - 1
				If iInstanceCnt = 0 Then
					' For Last Node of an Element
					If jCounter = UBound(aSubElement) Then
						bFlag = True
						Exit for
					end if
					jCounter = jCounter + 1
					sNode = ""
				End If
			End If
		Next
		If bFlag  Then
			Fn_WebUI_TableRowIndex = iCounter
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS : Fn_WebUI_TableRowIndex: Node [ " & sNodeName & " ] is present at index [ " & Fn_WebUI_TableRowIndex & " ] in specified able.")
		else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Fn_WebUI_TableRowIndex: Node [ " & sNodeName & " ] is not present in specified able.")
		End If
End function

'*********************************************************		Function to Get 	 Table Column Index *************************************************************************

'Function Name		:		Fn_WebUI_TableColumnIndex

'Description			 :		  This function is used to get Table column index.

'Parameters			   :	 			1.  objTable : Table Object
'									   				 2.  sColName : Column Name
											
'Return Value		   : 				 Column index

'Pre-requisite			:		 		Table be displayed in web browser .

'Examples				:				 Fn_WebUI_TableColumnIndex(Browser("TeamcenterStructureWeb").Page("TeamcenterWebStructure").WebTable("BOMTable"), "BOM Line")

'History:
'										Developer Name				Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Kaustubh\Sandeep						 12-May-2011		1.0
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
public function Fn_WebUI_TableColumnIndex(objTable,sColName)
		Dim iColCount, iCounter
		Fn_WebUI_TableColumnIndex = -1
		If Not objTable.Exist(5) Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Fn_WebUI_TableColumnIndex: Specified table does not exist in web browser.")
			Exit function
		End If
		iColCount = objTable.ColumnCount(1)
		For iCounter = 1 to iColCount
				If  sColName = trim(objTable.GetCellData(1,iCounter)) Then
							Fn_WebUI_TableColumnIndex = iCounter
							Exit for
				End If
		Next
		If Fn_WebUI_TableColumnIndex = -1 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Fn_WebUI_TableColumnIndex: Column [ " & sColName & " ] is not present in specified able.")
		else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS : Fn_WebUI_TableColumnIndex: Column [ " & sColName & " ] is present at index [ " & Fn_WebUI_TableColumnIndex & " ] in specified able.")
		End If
End function
'************************* Fn_SISW_WebUI_WebListItemExist ********************************************************

'Function Name		:	Fn_SISW_WebUI_WebListItemExist

'Description		:	This function is use to check for item present in list or not.
'
'PARAMETERS      	:   1. sFunctionName		: Valid Function Name
'			    		2. objWebPage			: Valid Web Page	
'			    		3. sWebList				: Valid Web List Name,
'			            4. sElementToSelect 	: Valid element to the selected   ( Single Element)		
'			   							
'Return Value		: 	True / False

'Pre-requisite		:	Web List must be displayed in web browser .

'Examples			:	Fn_SISW_WebUI_WebListItemExist("Fn_Web_DC_ContextDefinitionOperations", Browser("Teamcenter Web - Design").Page("Teamcenter Web - Design"), "ProductContext", "ad")

'History:
'		Developer Name				Date		Rev. No.	Reviewer			Changes Done
'-----------------------------------------------------------------------------------------------------------------
'		Kaustubh W				 12-May-2011	1.0			
'-----------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_WebUI_WebListItemExist(sFunctionName, objWebPage, sWebList, sElementToSelect)
	Dim objWebList, sUIFail, iElecount, iCounter
	
	sUIFail = sFunctionName + ">> Fn_SISW_WebUI_WebListItemExist >> " +  objWebPage.toString +">> " +  sWebList
	Fn_SISW_WebUI_WebListItemExist = False
	Set objWebList = objWebPage.WebList(sWebList)	
												
	' Synchronization Point for an Java List Object 
	If Fn_Web_UI_ObjectExist("Fn_SISW_WebUI_WebListItemExist", objWebList) = False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Fn_SISW_WebUI_WebListItemExist : WebList " &sWebList & " does not exist of Function " &sFunctionName)
	End If

	' get total items from list
	iEelecount = objWebList.GetROProperty("items count")
	For iCounter = 1 To iEelecount
		If objWebList.GetItem(iCounter) <> "" Then
			If trim(cstr(objWebList.GetItem(iCounter))) = trim(cstr(sElementToSelect)) Then
				Fn_SISW_WebUI_WebListItemExist = True
				Exit For
			End If
		End If
	Next
	If Fn_SISW_WebUI_WebListItemExist = true  Then
		' log result 
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),  "PASS : Fn_SISW_WebUI_WebListItemExist : Item " & sElementToSelect &" is present in  WebList " & sWebList & " in function " &sFunctionName)
	Else
		'Report error when item not present in the list
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Fn_SISW_WebUI_WebListItemExist : Item " & sElementToSelect &" is not present in WebList " & sWebList & " in function " &sFunctionName)
	End If
	' Clear Object
	Set objWebList =Nothing 
End Function



'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'/$$$$
'/$$$$   FUNCTION NAME   :   Fn_Web_UI_WebWinEdit_GetValue(sFunctionName, objPage, sWinEdit)
'/$$$$
'/$$$$   DESCRIPTION        :  This function will return value from a WinEdit on a browser
'/$$$$
'/$$$$    PARAMETERS      :   1.) sFunctionName : 'Valid Function /Test Nme
'/$$$$                                      2.) objPage : Hierarchy of the Page / Browser  that contains the Win Edit
'/$$$$
'/$$$$    Function Calls       :   Fn_WriteLogFile()
'/$$$$									  
'/$$$$
'/$$$$	 HISTORY           :   AUTHOR                 DATE        VERSION
'/$$$$
'/$$$$    CREATED BY     :   SHREYAS           19/10/2012         1.0
'/$$$$
'/$$$$    REVIWED BY     :   shreyas
'/$$$$
'/$$$$    MODIFIED BY   :  Shreyas : Added cases for Verify, VerifyBlank, ExistInTree, ClickFromTree(22nd Dec 10)
'/$$$$
'/$$$$   EXAMPLE          : 	sValue=Fn_Web_UI_WebWinEdit_GetValue(Environment.Value("TestName"), Browser("Browser"), "AddressBar")
'/$$$$ 										 msgbox sValue ,vbinformation,"Return Value"
'/$$$$ 
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

Public Function Fn_Web_UI_WebWinEdit_GetValue(sFunctionName, objPage, sWinEdit)
		Dim objWebEdit, bReturn
		Set objWebEdit = objPage.WinEdit(sWinEdit)

		bReturn= Fn_Web_UI_ObjectExist("Fn_Web_UI_ObjectCreate", objWebEdit)

		If  bReturn=True Then
				If  objWebEdit.CheckProperty("visible",True)Then 
						objWebEdit.GetROProperty("text")
						Fn_Web_UI_WebWinEdit_GetValue = objWebEdit.GetROProperty("text")
						Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"UI_Web :PASS: Successfully WinEdit "+CStr(objWebEdit.toString)+" Box fetched Value "+CStr(objWebEdit.GetROProperty("text"))+" on Function "+CStr(sFunctionName)+"")
				Else
						'Report error/message when WebEditBox object is disable.
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "UI_Web : Fail : WebEdit " +CStr(objWebEdit.toString)+"  is not enabled of Function " & sFunctionName)
						Fn_Web_UI_WebWinEdit_GetValue = False
						Call ExitFromWeb_UI("")
				End If
		Else
				Fn_Web_UI_WebWinEdit_GetValue = False
				Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"UI_Web :FAIL: Failed to fetch Value "+CStr(objWebEdit.GetROProperty("text"))+" in WinEdit "+CStr(objWebEdit.toString)+" Box on Function "+CStr(sFunctionName)+" ")
				Call ExitFromWeb_UI("")
		End If

		'Clear memory of WebButton object.
		Set objWebEdit = Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'    FUNCTION NAME   :   Fn_Web_UI_WebEdit_SetExt()
'
'    DESCRIPTION     :   This function is used to Edit.
'
'    PARAMETERS      :   sFunctionName : Valid Function name, 
'                                               sAction :  Valid Action
'                       						 objPage : Valid Dialog box object name,
'                        					sWinEdit   : Valid Edit box
'                        					sValue   : Valid value to set
'
'    Function Calls  :   Fn_WriteLogFile()
'
'    HISTORY         :   	AUTHOR         			DATE          VERSION
'
'    CREATED BY      :   Sandeep N          25/03/2013    1.0
'
'    EXAMPLE         :   Fn_Web_UI_WebEdit_SetExt("", "Set",objTeamcenterWebStructure.WebTable("CreateSnapshot"), "SnapshotName", "SnapShot1")
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Public Function Fn_Web_UI_WebEdit_SetExt(sFunctionName, sAction,objPage, sWinEdit, sValue)
		Dim objWebEdit, bReturn, objobjMDR
		If isNull(objPage) = False Then
			'Creating Object of Edit Box
			Set objWebEdit = objPage.WebEdit(sWinEdit)
		Else
			Set objWebEdit = sWinEdit
		End If
		'Checking Existance of Edit Box
		bReturn= Fn_Web_UI_ObjectExist("Fn_Web_UI_WebEdit_SetExt", objWebEdit)
		If  bReturn=True Then
				If objWebEdit.CheckProperty("disabled","0") AND   objWebEdit.CheckProperty("visible",True)Then 
						Select Case sAction
							Case "Set"
								objWebEdit.Set sValue	

							Case "SendString"
								Set objobjMDR = CreateObject("Mercury.DeviceReplay")
								objWebEdit.Set ""
								objWebEdit.Object.focus
								objobjMDR.SendString sValue
								Set objobjMDR =Nothing
						End Select
						Fn_Web_UI_WebEdit_SetExt = True
						Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"UI_Web :PASS: Successfully WinEdit "+CStr(objWebEdit.toString)+" Box  Set with Value "+CStr(sValue)+" on Function "+CStr(sFunctionName)+"")
				Else
						'Report error/message when WebEditBox object is disable.
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "UI_Web : Fail : WebEdit " +CStr(objWebEdit.toString)+"  is not enabled of Function " & sFunctionName)
						Fn_Web_UI_WebEdit_SetExt = False
						Call ExitFromWeb_UI("")
				End If
		Else
				Fn_Web_UI_WebEdit_SetExt = False
				Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"UI_Web :FAIL: Failed to Set with Value "+CStr(sValue)+" in WinEdit "+CStr(objWebEdit.toString)+" Box on Function "+CStr(sFunctionName)+" ")
				Call ExitFromWeb_UI("")
		End If
		'Clear memory of WebButton object.
		Set objWebEdit = Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'    FUNCTION NAME   :   Fn_Web_UI_WebEdit_GetValue()
'
'    DESCRIPTION     :   This function is used to get value from an WebEdit.
'
'    PARAMETERS      :   sFunctionName : Valid Function name, 
'                       						 objPage : Valid Dialog box object name,
'                        					sWebEdit   : Valid Edit box
'
'    Function Calls  :   Fn_WriteLogFile()
'
'    HISTORY         :   	AUTHOR         			DATE          VERSION
'
'    CREATED BY      :  Sandeep N			27-May-2013
'
'    EXAMPLE         :   bReturn= Fn_Web_UI_WebEdit_GetValue("", Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("UserSettings"), "ProgramName")
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Public Function Fn_Web_UI_WebEdit_GetValue(sFunctionName, objPage, sWebEdit)
		Dim objWebEdit, bReturn
		Set objWebEdit = objPage.WebEdit(sWebEdit)

		bReturn= Fn_Web_UI_ObjectExist("Fn_Web_UI_ObjectCreate", objWebEdit)

		If  bReturn=True Then
				If  objWebEdit.CheckProperty("visible",True)Then 
						Fn_Web_UI_WebEdit_GetValue = objWebEdit.GetROProperty("value")
						Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"UI_Web :PASS: Successfully fetched Value "+CStr(objWebEdit.GetROProperty("value"))+" from WebEdit "+CStr(objWebEdit.toString))
				Else
						'Report error/message when WebEditBox object is disable.
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "UI_Web : Fail : WebEdit " +CStr(objWebEdit.toString))
						Fn_Web_UI_WebEdit_GetValue = False
						Call ExitFromWeb_UI("")
				End If
		Else
				Fn_Web_UI_WebEdit_GetValue = False
				Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"UI_Web :FAIL: Failed to fetch Value "+CStr(objWebEdit.GetROProperty("text"))+" from WebEdit "+CStr(objWebEdit.toString))
				Call ExitFromWeb_UI("")
		End If

		'Clear memory of WebEdit object.
		Set objWebEdit = Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'    FUNCTION NAME   :   Fn_Web_UI_Button_ClickExt()
'
'    DESCRIPTION     :   This function is used to perform operations on Web Button
'
'    PARAMETERS      :   sFunctionName : Valid Function name, 
'                                               sAction :  Valid Action
'                       						 objPage : Valid Dialog box object name,
'                        					sWebButton   : Valid Button
'
'    Function Calls  :   Fn_WriteLogFile()
'
'    HISTORY         :   	AUTHOR         			DATE          VERSION
'
'    CREATED BY      :   Sandeep N          04/06/2013    1.0
'
'    EXAMPLE         :   Fn_Web_UI_Button_ClickExt("", "Click",objTeamcenterWebStructure.WebElement("ButtonPanel"), "SaveAndCheckIn")
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Function Fn_Web_UI_Button_ClickExt(sFunctionName, sAction,objPage, sWebButton)
	Dim objWebButton, bReturn
	'Object Creation
	Fn_Web_UI_Button_ClickExt=False
	'Creating object of button
	Set objWebButton = objPage.WebButton(sWebButton)
	bReturn= Fn_Web_UI_ObjectExist("Fn_Web_UI_ObjectCreate", objWebButton)
	If  bReturn=True Then
		 If objWebButton.CheckProperty("disabled","0") AND   objWebButton.CheckProperty("visible",True) Then 
			  Select Case sAction
		  			Case "Click"
						  objWebButton.Click	
						  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "UI_Web : Sucessfully clicked on WebButton " & sWebButton & " of Function " & sFunctionName)
						   Fn_Web_UI_Button_ClickExt = True
			  End Select
		Else			
		'Report error/message when Web Button object is disable.
		  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "UI_Web : WebButton " &sWebButton &"  is not enabled of Function " & sFunctionName)
		  Fn_Web_UI_Button_ClickExt = False
		  Call ExitFromWeb_UI("")
	End If
	'Report error/message when WebButton object does not exists.
	Else
	   Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "UI_Web : WebButton " &sWebButton & "  does not exist of Function " &sFunctionName )
	   Fn_Web_UI_Button_ClickExt = False
	   Call ExitFromWeb_UI("")
	End If	
	'Clear memory of WebButton object.
	Set objWebButton = Nothing 
End Function

'#################################################################################################################
'###    FUNCTION NAME   :    Fn_WEB_UI_Object_GetROProperty(sFunctionName,objDialog,sProperty,sPropValue)
'###
'###    DESCRIPTION     :   This function  is used to Get RO Property For given Object
'###
'###    PARAMETERS      :   sFunctionName - Valid function name
'###                        					objDialog - Valid  Dialog Object
'###                              				sProperty-Valid Property Name
'###                        
'###    Function Calls  :   Fn_WriteLogFile()
'###
'###    HISTORY         :   AUTHOR                   DATE        	VERSION
'###
'###    CREATED BY      :   Vrushali S      		04-Jan-2017 	  1.0  
'###
'###    REVIWED BY      :   Prasad					04-Jan-2017 	  1.0 
'###
'###    EXAMPLE         :Call Fn_WEB_UI_Object_GetROProperty("Fn_WEB_UI_Object_GetROProperty",Browser("Teamcenter").Link("Logout"),"text")
'###										
'################################################################################################################

Function Fn_WEB_UI_Object_GetROProperty(sFunctionName,objDialog,sProperty)
		Dim objGetROProperty
		Fn_WEB_UI_Object_GetROProperty = False
		Set objGetROProperty=objDialog
		If objGetROProperty.Exist Then
			'Checking existance of object
			Fn_WEB_UI_Object_GetROProperty=objGetROProperty.GetROProperty(sProperty)
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"UI_Web :The Property "&sProperty &" value is " & cStr(Fn_WEB_UI_Object_GetROProperty) &" for " & objGetROProperty.toString & "  of Function " &sFunctionName)	
		End If
		Set objGetROProperty = Nothing 
End Function