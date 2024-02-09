Option Explicit
'*********************************************************	Function List		***********************************************************************
'1. Fn_SISW_ChkInOut_ObjectCheckOut
'2. Fn_SISW_ChkInOut_ObjectCancelCheckOut
'3. Fn_SISW_ChkInOut_CheckOutMessageVerify
'4. Fn_SISW_ChkInOut_ObjectCheckIn
'5. Fn_SISW_ChkInOut_CheckOutMsgVerify
'6. Fn_SISW_ChkInOut_OptionsSettings
'7. Fn_SISW_ChkInOut_ChkOutHistoryVerification
'8. Fn_SISW_ChkInOut_GetChkInChkOutObject
'*********************************************************	Function List		***********************************************************************


'*********************************************************		Function Checks Out the Teamcenter Obejct		***********************************************************************
'Function Name		:				Fn_SISW_ChkInOut_ObjectCheckOut

'Description			 :		 		 General utility function which performs Check-Out operation from various application interface initiation points

'Parameters			   : 	               1)sAction: Action Label to guide code segment to navigate to prefered case statements
'														2)sChangeID: Change ID value for Check-Out operation
'														3)sComment: Comment value for Check-Out operation
'														4)bExport: True/False flag to export dataset on Check-out operation
'														5)bOverwrite: True/Fasle flag to overwrite already exported dataset
'														6)aExploreOption: array with content details of handling [Explore] dialog details
'														7)sItem : Item which error need to verify.
'														8)sErrorMsg : Error message need to verify.

'Return Value		   : 				TRUE \ FALSE

'Pre-requisite			:			 		1. Teamcenter RAC Application accessible
'														2. Requisite Business Object Selected under Application context

'Examples				:				 Call Fn_SISW_ChkInOut_ObjectCheckOut("Summary pane More Properties Check- out Edit",  "",  "","" ,"", "", "", "" )

'History:
'										Developer Name			Date			Rev. No.			Changes Done											Reviewer	Reviewed date
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Harshal	Tanpure					18-May-2010		1.0																	Sameer	18-May-10
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Koustubh Watwe					3-Dec-2010		1.0	         Modified case Menu Check out, replaced menu operation call according to menu.
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Ashok Kakade					3-Dec-2010		1.0	         Modified case Viewer Tab Check-Out and Edit 			Koustubh Watwe
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Swapnil Gore					18-May-2012		1.0			 Modified the Menu Check-Out
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										 Pooja B					      10-Dec-2012		1.0			 Modified the Summary Toolbar Check-Out and edit
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										 Sandeep N					    02-Aug-2013		1.1			 Added New Hierarchy : JavaWindow("DefaultWindow").JavaWindow("Check-Out") For perspective : [ My Teamcenter ] And [ Change Manager ]
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										 Pranav Ingle					  17-Oct-2013		1.1			Modified Case "Check out with error verify" for new checkout dialog: JavaWindow("DefaultWindow").JavaWindow("Check-Out")
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function Fn_SISW_ChkInOut_ObjectCheckOut(sAction, sChangeID,  sComment, bExport, bOverwrite, aExploreOption, sItem, sErrorMsg )	
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_ChkInOut_ObjectCheckOut"
	Dim objCheckOut, objProperties, objExplore,  objErrorDialog, objDesBtnCloseRed, objDesFetched_static_text, objaFetchText, bErrorDialog
	Dim StrTitle
	Dim bPropertiesDialog, bCheckOutDialog, bExploreDialog, aFetchText, iCount, iCounter, sErrorText
	
	'Getting current title
	StrTitle=JavaWindow("DefaultWindow").GetROProperty("title")
	Select Case sAction
			Case "Menu CheckOut"
				
						Set objCheckOut=Fn_SISW_ChkInOut_GetChkInChkOutObject("CheckOut")
						If typename(objCheckOut) = "Nothing" Then
						  JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").SetTOProperty "Index" ,0 
						  JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").SetTOProperty "Index" ,0 
						  Set objCheckOut=Fn_SISW_ChkInOut_GetChkInChkOutObject("CheckOut")
						   If typename(objCheckOut) = "Nothing"	   Then
						   		Call Fn_MenuOperation("Select", "Tools:Check-In/Out:Check Out...")
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Menu Item Tools > Check-Out selected successfully")
						   End If		
						End If
						Set objCheckOut= Nothing

						Set objCheckOut=Fn_SISW_ChkInOut_GetChkInChkOutObject("CheckOut")
						If typename(objCheckOut) = "Nothing" Then
						  JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").SetTOProperty "Index" ,1 
						  JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").SetTOProperty "Index" ,1 
						  Set objCheckOut=Fn_SISW_ChkInOut_GetChkInChkOutObject("CheckOut")
						   If typename(objCheckOut) = "Nothing"	   Then
						   		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Check-out Dialog Does not exixts...")
								Exit Function
						   End If	
						End If

						'Entering values in Check Out Dialog
						Call Fn_Edit_Box("Fn_SISW_ChkInOut_ObjectCheckOut",objCheckOut,"ChangeID",sChangeID)
						Call Fn_Edit_Box("Fn_SISW_ChkInOut_ObjectCheckOut",objCheckOut,"Comments",sComment)
						If bExport = TRUE Then
                                Call Fn_CheckBox_Set("Fn_SISW_ChkInOut_ObjectCheckOut", objCheckOut, "ExportDtsetOnCheck-Out", "ON")
                        ElseIf bExport = FALSE Then
								Call Fn_CheckBox_Set("Fn_SISW_ChkInOut_ObjectCheckOut", objCheckOut, "ExportDtsetOnCheck-Out", "OFF")	
                        End If
						If bOverwrite=TRUE Then
								Call Fn_CheckBox_Set("Fn_SISW_ChkInOut_ObjectCheckOut", objCheckOut, "OverwriteExistingFiles", "ON")
                        Elseif bOverwrite=FALSE Then
									Call Fn_CheckBox_Set(" Fn_SISW_ChkInOut_ObjectCheckOut", objCheckOut, "OverwriteExistingFiles", "OFF")
                        End If	
						
						'Swapnil: Activate the dialog as it's unable to click on "Yes" button sometimes
						
						objCheckOut.Activate
						
						Call Fn_Button_Click("Fn_SISW_ChkInOut_ObjectCheckOut", objCheckOut, "Yes")
                        Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Values  enetered in the Check Out Dialog Successfully")
						  JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").SetTOProperty "Index" ,1 
						  JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").SetTOProperty "Index" ,1 
						Set objCheckOut= Nothing
			Case  "Summary Toolbar Check-Out and edit"
							'Selecting Check Out  and edit buttoin from Summary Toolbar
							'*Modiffied by Pooja B. according to TC10.1 Changes on 10-Dec-2012
					         Call Fn_ToolbatButtonClick("Check Out...")
								'Entering values in Check Out Dialog
							If Instr(1,StrTitle,"My Teamcenter") or Instr(1,StrTitle,"Change Manager") Then
								Set objCheckOut=Fn_SISW_GetObject("Check-Out@2")
							Else
								Set  objCheckOut = Fn_UI_ObjectCreate("Fn_SISW_ChkInOut_ObjectCheckOut", JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Check-Out"))
							End if
							Call Fn_Edit_Box("Fn_SISW_ChkInOut_ObjectCheckOut",objCheckOut,"ChangeID",sChangeID)
							Call Fn_Edit_Box("Fn_SISW_ChkInOut_ObjectCheckOut",objCheckOut,"Comments",sComment)
							If bExport = TRUE Then
									Call Fn_CheckBox_Set(" Fn_SISW_ChkInOut_ObjectCheckOut", objCheckOut, "ExportDtsetOnCheck-Out", "ON")
                            ElseIf bExport = FALSE Then
									Call Fn_CheckBox_Set(" Fn_SISW_ChkInOut_ObjectCheckOut", objCheckOut, "ExportDtsetOnCheck-Out", "OFF")
                            End If
							If bOverwrite=TRUE Then
									Call Fn_CheckBox_Set(" Fn_SISW_ChkInOut_ObjectCheckOut", objCheckOut, "OverwriteExistingFiles", "ON")
                            Elseif bOverwrite=FALSE Then
										Call Fn_CheckBox_Set(" Fn_SISW_ChkInOut_ObjectCheckOut", objCheckOut, "OverwriteExistingFiles", "OFF")
                            End If	
							Call Fn_Button_Click("Fn_SISW_ChkInOut_ObjectCheckOut", objCheckOut, "Yes")
                            Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Values  enetered in the Check Out Dialog Successfully")
			Case  "Viewer Tab Check-Out and Edit"
							If Instr(1,StrTitle,"My Teamcenter") or Instr(1,StrTitle,"Change Manager") Then
								Set objCheckOut=Fn_SISW_GetObject("Check-Out@2")
							Else
								Set objCheckOut=JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Check-Out")
							End If
						   If Fn_UI_ObjectExist("Fn_SISW_ChkInOut_ObjectCheckOut", objCheckOut) = False Then
								'Selecting Check Out  and edit buttoin from Summary Toolbar
								If JavaWindow("MyTeamcenter").JavaWindow("MyTcJApplet").JavaButton("Check-Out and Edit").Exist(5) Then
									Call Fn_Button_Click("Fn_SISW_ChkInOut_ObjectCheckOut", JavaWindow("MyTeamcenter").JavaWindow("MyTcJApplet"), "Check-Out and Edit")
								ElseIf  JavaWindow("MyTeamcenter").JavaToolbar("CheckInCheckOutToolbar").Exist(5) Then
									' buttons are replaced with JavaToolbar after deploying ADS template - Ashok Kakade, Koustubh Watwe
									Call Fn_UI_JavaToolbar_Press("Fn_SISW_ChkInOut_ObjectCheckOut", JavaWindow("MyTeamcenter"), "CheckInCheckOutToolbar","Check Out...")
								ElseIf 	Fn_ToolbatButtonClick("Check Out...") =True then
										objCheckOut=True
								Else							
									 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_SISW_ChkInOut_ObjectCheckOut : failed to find Check out... button / toolbar.")
									exit function
								End If
							End If
							If JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Check-Out").exist(2) =  False Then
							     JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").SetTOProperty "Index",0
							End If
							If Instr(1,StrTitle,"My Teamcenter") or Instr(1,StrTitle,"Change Manager") Then
								Set objCheckOut=Fn_SISW_GetObject("Check-Out@2")
							Else
								Set  objCheckOut = Fn_UI_ObjectCreate("Fn_SISW_ChkInOut_ObjectCheckOut", JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Check-Out"))
							End If
						    
						    Call Fn_Edit_Box("Fn_SISW_ChkInOut_ObjectCheckOut",objCheckOut,"ChangeID",sChangeID)
						    Call Fn_Edit_Box("Fn_SISW_ChkInOut_ObjectCheckOut",objCheckOut,"Comments",sComment)
						    If bExport = TRUE Then
								 Call Fn_CheckBox_Set(" Fn_SISW_ChkInOut_ObjectCheckOut", objCheckOut, "ExportDtsetOnCheck-Out", "ON")
							ElseIf bExport = FALSE Then
								 Call Fn_CheckBox_Set(" Fn_SISW_ChkInOut_ObjectCheckOut", objCheckOut, "ExportDtsetOnCheck-Out", "OFF")
							End If
						    If bOverwrite=TRUE Then
								Call Fn_CheckBox_Set(" Fn_SISW_ChkInOut_ObjectCheckOut", objCheckOut, "OverwriteExistingFiles", "ON")
							Elseif bOverwrite=FALSE Then
								Call Fn_CheckBox_Set(" Fn_SISW_ChkInOut_ObjectCheckOut", objCheckOut, "OverwriteExistingFiles", "OFF")
							End If 
							Call Fn_Button_Click("Fn_SISW_ChkInOut_ObjectCheckOut", objCheckOut, "Yes")
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Values  enetered in the Check Out Dialog Successfully")
							JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").SetTOProperty "Index",1
			Case "Summary Toolbar Save and Keep Check-Out"
							'Selecting Save and Keep Check-Out  buttoin from Summary Toolbar
							Call Fn_ToolbatButtonClick("Save and keep Checked-out")
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Save and Keep Checked-out button clicked Successfully")
			Case "View Menu Check-Out and edit"
			Case "View Menu Save and Keep Check-Out"
			Case "Summary pane More Properties Check- out Edit"
					'click the More Properties from Summary Pane
					JavaWindow("MyTeamcenter").JavaObject("SummaryMoreProperties").Click 1,1
					'Verify the Property dialog is exist
					Set  objProperties = Fn_UI_ObjectCreate("Fn_SISW_ChkInOut_ObjectCheckOut",JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Properties"))
					'If JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Properties").Exist then
					Call Fn_Button_Click("Fn_SISW_ChkInOut_ObjectCheckOut", objProperties, "Check-Out and Edit")
					gLastMenuCall = "CheckOutAndEdit"
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Property Dialog Exist")	
					Set objCheckOut = Fn_SISW_GetChkInChkOutObject("CheckOut")
					'Verify the Check-out dialog is exist
					If Typename(objCheckOut) <> "Nothing" Then
						Call Fn_Button_Click("Fn_SISW_ChkInOut_ObjectCheckOut", objCheckOut, "Yes")
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Check-Out Dialog does Not Exist")	
						Fn_SISW_ChkInOut_ObjectCheckOut = False
						Exit Function 
					End If
					'If   JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Check-Out").Exist Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Check-Out Dialog Exist")	
			Case "Check out with Explore button to select all the component"
						Call Fn_MenuOperation("Select", "Tools:Check-In/Out:Check Out...")
						If Instr(1,StrTitle,"My Teamcenter") or Instr(1,StrTitle,"Change Manager") Then
								Set objCheckOut=Fn_SISW_GetObject("Check-Out@2")
						Else
							Set  objCheckOut = Fn_UI_ObjectCreate("Fn_SISW_ChkInOut_ObjectCheckOut", JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Check-Out"))
						End If
						Call Fn_Button_Click("Fn_SISW_ChkInOut_ObjectCheckOut", objCheckOut, "ExploerComponents")
						Set objExplore = JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Explore")
						If Fn_SISW_UI_Object_Operations("Fn_SISW_ChkInOut_ObjectCheckOut", "Exist", objExplore, "") = False Then
							Set  objExplore = JavaDialog("Explore")
							If Fn_SISW_UI_Object_Operations("Fn_SISW_ChkInOut_ObjectCheckOut", "Exist", objExplore, SISW_MICRO_TIMEOUT) = False Then
								Fn_SISW_ChkInOut_ObjectCheckOut = False
								Exit function
							End If 
						End If
										Call Fn_Button_Click("Fn_SISW_ChkInOut_ObjectCheckOut", objExplore, "SelectAll")
                                        Call Fn_Button_Click("Fn_SISW_ChkInOut_ObjectCheckOut", objExplore, "OK")
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
										If aExploreOption <> "" Then										
										'Dim desFetched_static_text
										Set objDesFetched_static_text=description.Create()
										objDesFetched_static_text("Class Name").value="JavaStaticText"
										Set objaFetchText=objCheckOut.ChildObjects(objDesFetched_static_text)
										iFetchTextCount=objaFetchText.Count-4 
                                    	iArrayCount=Ubound(aExploreOption)
										iCount=0
                                    	If iFetchTextCount=iArrayCount then
											For iCounter=3 to objaFetchText.Count-1
													If aExploreOption(iCount)=objaFetchText(iCounter).GetRoProperty("attached text") Then
															iCount=iCount+1
													Else
															Fn_SISW_ChkInOut_ObjectCheckOut =FALSE
															Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "array details and selected items are not equal")	
															Exit Function
													End If
											Next
										Else
												Fn_SISW_ChkInOut_ObjectCheckOut =FALSE
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "array details and selected items are not equal")	
												Exit Function
										End If
								  End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                  						Call Fn_Button_Click("Fn_SISW_ChkInOut_ObjectCheckOut", objCheckOut, "Yes")
                                		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Explore Dialog Exist")
										'Verify the Check-Out dialog is exist	
									'	If JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Check-Out").Exist  Then
										'If Fn_UI_ObjectExist("Fn_SISW_ChkInOut_ObjectCheckOut", objCheckOut) = False then
												'	Fn_SISW_ChkInOut_ObjectCheckOut=FALSE
													'Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Check-Out Dialog Exist")	
													'Exit Function
										'End If

			Case "Check out with error verify"
						Set objCheckOut=Fn_SISW_ChkInOut_GetChkInChkOutObject("CheckOut")
						If typename(objCheckOut) = "Nothing" Then
							Call Fn_MenuOperation("Select", "Tools:Check-In/Out:Check Out...")
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Menu Item Tools > Check-Out selected successfully")
						End If
						Set objCheckOut= Nothing

						Set objCheckOut=Fn_SISW_ChkInOut_GetChkInChkOutObject("CheckOut")
						If typename(objCheckOut) = "Nothing" Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Check-out Dialog Does not exixts...")
							Exit Function
						End If

						'Enter details in Check Out dialog
						Call Fn_Edit_Box("Fn_SISW_ChkInOut_ObjectCheckOut",objCheckOut,"ChangeID",sChangeID)
						Call Fn_Edit_Box("Fn_SISW_ChkInOut_ObjectCheckOut",objCheckOut,"Comments",sComment)
						If bExport = TRUE Then
								Call Fn_CheckBox_Set(" Fn_SISW_ChkInOut_ObjectCheckOut", objCheckOut, "ExportDtsetOnCheck-Out", "ON")
						ElseIf bExport = FALSE Then
								Call Fn_CheckBox_Set(" Fn_SISW_ChkInOut_ObjectCheckOut", objCheckOut, "ExportDtsetOnCheck-Out", "OFF")
						End If
						If bOverwrite=TRUE Then
									Call Fn_CheckBox_Set(" Fn_SISW_ChkInOut_ObjectCheckOut", objCheckOut, "OverwriteExistingFiles", "ON")
						Elseif bOverwrite=FALSE Then
									Call Fn_CheckBox_Set(" Fn_SISW_ChkInOut_ObjectCheckOut", objCheckOut, "OverwriteExistingFiles", "OFF")
						End If	
						Call Fn_Button_Click("Fn_SISW_ChkInOut_ObjectCheckOut", objCheckOut, "Yes")
						'Dim objDesBtnCloseRed
						Call Fn_ReadyStatusSync(5)
						wait(5)
						Set objDesBtnCloseRed=description.Create()
						objDesBtnCloseRed("label").value="defaulterror_16"
						wait(3)						
						If  objCheckOut.JavaButton(objDesBtnCloseRed).Exist(2) Then
							objCheckOut.JavaButton(objDesBtnCloseRed).Click micLeftBtn
							Set objErrorDialog=Fn_UI_ObjectCreate("Fn_SISW_ChkInOut_ObjectCheckOut", Dialog("ErrorMsgDialog"))
							wait(5)
							bErrorDialog=Fn_UI_ObjectExist("Fn_SISW_ChkInOut_ObjectCheckOut", objErrorDialog)
							If bErrorDialog then
								'Added By Sandeep N
								If Instr(1,StrTitle,"My Teamcenter") or Instr(1,StrTitle,"Change Manager") Then
									sErrorText=JavaWindow("MyTeamcenter").JavaWindow("CheckOut").JavaDialog("Error").JavaEdit("ErrText").GetROProperty("value")
								Else
									sErrorText=JavaWindow("MyTeamcenter").JavaWindow("MyTcApplet").JavaDialog("CheckOut").JavaDialog("Error").JavaEdit("ErrText").GetROProperty("value")
								End if
								wait(5)
								If Instr(1,sErrorText,sErrorMsg)=0 then
									Fn_SISW_ChkInOut_ObjectCheckOut =FALSE
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Error msg is not correct")
									Exit Function
								end if
							Else
								Fn_SISW_ChkInOut_ObjectCheckOut =FALSE
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL:Error Dialog Exist")
								Exit Function
							End If
							wait(5)
	'						Dialog("ErrorMsgDialog").WinButton("OK").Click
							If Instr(1,StrTitle,"My Teamcenter") or Instr(1,StrTitle,"Change Manager") Then
								JavaWindow("MyTeamcenter").JavaWindow("CheckOut").JavaDialog("Error").JavaButton("OK").Click
							Else
								JavaWindow("MyTeamcenter").JavaWindow("MyTcApplet").JavaDialog("CheckOut").JavaDialog("Error").JavaButton("OK").Click
							End if
						wait(5)
						ElseIf JavaWindow("DefaultWindow").InsightObject("RedCrossImage").Exist(2) Then
							JavaWindow("DefaultWindow").InsightObject("RedCrossImage").Click 5,5,micLeftBtn
							wait(3)
							Set objErrorDialog=Fn_UI_ObjectCreate("Fn_SISW_ChkInOut_ObjectCheckOut", objCheckOut.JavaWindow("Warning"))
							
							bErrorDialog=Fn_UI_ObjectExist("Fn_SISW_ChkInOut_ObjectCheckOut", objErrorDialog)
							If bErrorDialog then
								sErrorText = objErrorDialog.JavaEdit("DetailMsg").GetROProperty("value")
								wait(1)
								If Instr(1,sErrorText,sErrorMsg)=0 then
									Fn_SISW_ChkInOut_ObjectCheckOut =FALSE
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Error msg is not correct")
									Exit Function
								end if
							Else
								Fn_SISW_ChkInOut_ObjectCheckOut =FALSE
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL:Error Dialog Exist")
								Exit Function
							End If
							wait(1)
							objErrorDialog.JavaButton("OK").Click							

						Else 		' 17-Oct-2013 : Pranav    -  Added to handle new Checkout Dialog 
							sErrorText=objCheckOut.JavaTable("CheckOutList").GetCellData(0,1)
							If Instr(1,sErrorMsg,sErrorText)=0 then
								Fn_SISW_ChkInOut_ObjectCheckOut =FALSE
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Error msg is not correct")
								Exit Function
							End If
						End If
						
						wait(5)
						Call Fn_Button_Click("Fn_SISW_ChkInOut_ObjectCheckOut", objCheckOut,"OK")	
		Case  "Check-OutFromCustomizeToolBar"
							'Selecting Check Out  and edit buttoin from Summary Toolbar
							Call Fn_ToolbatButtonClick("Check Out...")
							If Instr(1,StrTitle,"My Teamcenter") or Instr(1,StrTitle,"Change Manager") Then
								Set objCheckOut=Fn_SISW_GetObject("Check-Out@2")
							Else
								Set  objCheckOut = Fn_UI_ObjectCreate("Fn_SISW_ChkInOut_ObjectCheckOut", JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Check-Out"))
							End If
							Call Fn_Button_Click("Fn_SISW_ChkInOut_ObjectCheckOut", objCheckOut, "Yes")
                            Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Values  enetered in the Check Out Dialog Successfully")				
			Case Else
						Fn_SISW_ChkInOut_ObjectCheckOut =FALSE
						Exit Function
		End Select
	Fn_SISW_ChkInOut_ObjectCheckOut =TRUE
	Set objCheckOut=Nothing
	Set objProperties=Nothing
	Set objExplore=Nothing
    Set objErrorDialog=Nothing
	Set objDesBtnCloseRed=Nothing
	Set objDesFetched_static_text=Nothing
	Set objaFetchText=Nothing
End Function

'*********************************************************		Function Checks Out the Teamcenter Obejct		***********************************************************************
'Function Name		:				Fn_SISW_ChkInOut_ObjectCancelCheckOut(sAction, sErrorButtonLabel, sErrorMessage)  

'Description			 :		 		 General utility function which performs [Cancel Check-Out] operation from various application interface initiation points

'Parameters			   : 	            1) sAction: Action Label to guide code segment to navigate to prefered case statements
'
'													2) sErrorButtonLabel: Error object label [Cancel Check-Out] dialog
'
'													3) sErrorMessage: Error message to validate
'

'Return Value		   : 				TRUE \ FALSE

'Pre-requisite			:			 		1. Teamcenter RAC Application accessible
'														2. Requisite Business Object Selected under Application context in Checked-Out state


'Examples				:				 Call Fn_SISW_ChkInOut_ObjectCancelCheckOut("Summary Toolbar Cancel Check-Out and revert back to original", "", "")  

'History:
'										Developer Name			Date			Rev. No.			Changes Done			Reviewer	Reviewed date
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Harshal	 Tanpure				27-May-2010		1.0													Manisha		27-May-10
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Koustubh Watwe			17-Nov-2011		1.0					Modified case else
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function Fn_SISW_ChkInOut_ObjectCancelCheckOut(sAction, sErrorButtonLabel, sErrorMessage)  
		GBL_FAILED_FUNCTION_NAME="Fn_SISW_ChkInOut_ObjectCancelCheckOut"
		Dim objCancelCheckOut
		Dim StrTitle

		Fn_SISW_ChkInOut_ObjectCancelCheckOut = False

		Set objCancelCheckOut=Fn_SISW_ChkInOut_GetChkInChkOutObject("CancelCheckOut")
		
		Select Case sAction
				Case "Summary Toolbar Cancel Check-Out and revert back to original"
						'Invoke Summary Toolbar button Tools > Cancel Check-Out and revert back to original
						If typename(objCancelCheckOut) = "Nothing" Then
							Call Fn_ToolbatButtonClick("Cancel Checkout...")
						End If
						Set objCancelCheckOut= Nothing
						
						'Call Fn_ToolbatButtonClick("Cancel Check-Out and revert back to original") ''commented by Avinash J. - 18Jan2013
						Set objCancelCheckOut=Fn_SISW_ChkInOut_GetChkInChkOutObject("CancelCheckOut")
						If typename(objCancelCheckOut) <> "Nothing" Then
							Call Fn_Button_Click("Fn_SISW_ChkInOut_ObjectCancelCheckOut", objCancelCheckOut, "Yes")
						Else
							Set objCancelCheckOut= Nothing
							Exit Function
						End If
				Case "Menu Cancel Check-Out"
						'Invoke  Tools > Cancel Check-Out and revert back to original
						If typename(objCancelCheckOut) = "Nothing" Then
							Call Fn_MenuOperation("Select","Tools:Check-In/Out:Cancel Checkout...")
						End If
						Set objCancelCheckOut= Nothing

                        Set objCancelCheckOut=Fn_SISW_ChkInOut_GetChkInChkOutObject("CancelCheckOut")
						If typename(objCancelCheckOut) <> "Nothing" Then
							Call Fn_Button_Click("Fn_SISW_ChkInOut_ObjectCancelCheckOut", objCancelCheckOut, "Yes")     
						Else
							Set objCancelCheckOut= Nothing
							Exit Function
						End If
						                  
				Case Else
						If objCancelCheckOut.exist(10) Then
							Call Fn_Button_Click("Fn_SISW_ChkInOut_ObjectCancelCheckOut", objCancelCheckOut, "Yes")                       
						Else
							Fn_SISW_ChkInOut_ObjectCancelCheckOut =FALSE
							Exit function
						End If
		End Select

		Fn_SISW_ChkInOut_ObjectCancelCheckOut=TRUE
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Cancel Check-Out done successfully")
		Set objCancelCheckOut=Nothing

End Function

'#############################################################################################################################################
'###    FUNCTION NAME   :  Fn_SISW_ChkInOut_CheckOutMessageVerify(sObject,sMsg,sButton)
'###
'###    DESCRIPTION     :   Function is used  for  varify the message in Read only window
'###
'###    PARAMETERS   :       1.sObject:           :   Valid passed Window text 
 '###                                              2. sMessage	:  Vallid Message to be passed 
 '###                                             3.sButton        : Button to be pressed
'###
'###
'###    HISTORY         :   AUTHOR          DATE           VERSION
'###	
'###    CREATED BY      :   Vidya       	12/04/2010            1.0
'###
'###    REVIWED BY      :	
'###
'###    EXAMPLE         :   Fn_SISW_ChkInOut_CheckOutMessageVerify("text","text is checked out by Vidya Kulkarni (x_kulkvi). Any changes will be lost."+vblf+"Do you want to continue?","No")
'#############################################################################################################################################
Function Fn_SISW_ChkInOut_CheckOutMessageVerify(sObject,sMsg,sButton)
		GBL_FAILED_FUNCTION_NAME="Fn_SISW_ChkInOut_CheckOutMessageVerify"
         Dim sFinalText,sVar
					   sVar="Checked Out -- "
                       sVar=sVar&sObject
                      JavaDialog("ReadOnly").SetTOProperty "title",sVar
					If  JavaDialog("ReadOnly").Exist Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Checked Out Dialog box exists")
							  sFinalText = JavaDialog("ReadOnly").JavaObject("Message").Object.GetText
							  If lcase(sFinalText) = lcase(sMsg)  Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Pass: Checkout Message Verification passed" )
								Call Fn_Button_Click("Fn_SISW_ChkInOut_CheckOutMessageVerify",JavaDialog("ReadOnly"),sButton)
								Fn_SISW_ChkInOut_CheckOutMessageVerify = True
							Else 
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail: Checkout Message Verification failed") 
								Fn_SISW_ChkInOut_CheckOutMessageVerify = False
							End If

					Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Checked Out Dialog box does not exists")
							Fn_SISW_ChkInOut_CheckOutMessageVerify = False
					End If

     Set sFinalText = Nothing

End Function 

'**********************************************************************************************************************************
' Replaced Prasanna's half complete function as per his request.
'**********************************************************************************************************************************
'*********************************************************		Function Checks Out the Teamcenter Obejct		***********************************************************************
'Function Name		:				Fn_SISW_ChkInOut_ObjectCheckIn

'Description			 :		 		 General utility function which performs Check-In operation from various application interface initiation points

'Parameters			   : 	               1) sAction: Action Label to guide code segment to navigate to prefered case statements

' 														2)aExploreOption: Array containing Dependent Object Names to be added through [Explore] dialog
'
' 														3) sErrorMessage: Error message to validate
'

'Return Value		   : 				TRUE \ FALSE

'Pre-requisite			:			 		1. Teamcenter RAC Application accessible
'														2. Requisite Business Object Selected under Application context in Checked-Out state


'Examples				:				 Call Fn_SISW_ChkInOut_ObjectCheckIn("Summary Toolbar Save properties and Check-In", "" , "")

'History:
'		Developer Name					Date			Rev. No.			Changes Done										Reviewer	Reviewed date
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'		Harshal	 Tanpure			25-May-2010			1.0																		Sameer		27-May-10
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'		Ashok Kakade				17-Feb-2012			1.0					Modified case "Viewer Tab Save and check-In"		Koustubh	17-Feb-2012
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'		Sachin Joshi				29-March-2012		1.0					Modified case "Menu Check-In","Check-In with Error"	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'		Ashok Kakade				02-Apr-2012			1.0					Modified code to check existence of the dialogbox	- Koustubh	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'		Snehal							04-Jul-2012			1.1					Modified code to check existence of new object Hierarchy of Check in Dialog in "Menu Check-In" Case - Sandeep Navghane
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'		Jeevan							02-Jul-2013			1.2					Added New Hierarchy : JavaWindow("DefaultWindow").JavaWindow("Check-In"). As design change in My Teamcenter & Change Manager perspective
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'		Ganesh B						21-May-2014			1.2					Added New case "VerifyChekInErrorMessage" to verify Error Message
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Function Fn_SISW_ChkInOut_ObjectCheckIn(sAction, aExploreOption, sErrorMessage)
		GBL_FAILED_FUNCTION_NAME="Fn_SISW_ChkInOut_ObjectCheckIn"
		
		Dim objCheckIn, objExplore, objDesFetched_static_text, objaFetchText, objDesBtnCloseRed, objErrorDialog
		Dim bSaveButtonExist, iFetchTextCount, iArrayCount, iCount, iCounter,  sErrorText,StrTitle
		Dim  sErrorMsg
		StrTitle=JavaWindow("DefaultWindow").GetROProperty("title")
		Select Case sAction
				Case "Menu Check-In"
					Set objCheckIn=Fn_SISW_ChkInOut_GetChkInChkOutObject("CheckIn")
					If typename(objCheckIn) = "Nothing" Then
						JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").SetTOProperty "Index" ,0 
						JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").SetTOProperty "Index" ,0
						Set objCheckIn=Fn_SISW_ChkInOut_GetChkInChkOutObject("CheckIn")
						If typename(objCheckIn) = "Nothing" Then
							Call Fn_MenuOperation("Select", "Tools:Check-In/Out:Check In...")
							Call Fn_ReadyStatusSync(1)
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Menu Item Tools:Check-In/Out:Check In... selected successfully")
						End If
					End If
					Set objCheckIn= Nothing
					'wait 2
					For iCounter=0 to 9
						Set objCheckIn=Fn_SISW_ChkInOut_GetChkInChkOutObject("CheckIn")
						If typename(objCheckIn) <> "Nothing" Then
							Exit For
						Else
						   JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").SetTOProperty "Index" ,1
						   JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").SetTOProperty "Index" ,1
						   Set objCheckIn=Fn_SISW_ChkInOut_GetChkInChkOutObject("CheckIn")
						   If  typename(objCheckIn) <> "Nothing" Then
						   	   Exit For
						   End If
						   JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").SetTOProperty "Index" ,0
						   JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").SetTOProperty "Index" ,0
						End If
						Set objCheckIn=Nothing
					Next
					
					If typename(objCheckIn) = "Nothing" Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed Function [Fn_SISW_ChkInOut_ObjectCheckIn]  to find [Check-In] dialog.")
						Set objCheckIn= Nothing
						Exit Function
					End If
					
					Call Fn_Button_Click("Fn_SISW_ChkInOut_ObjectCheckIn", objCheckIn, "Yes")
					Call Fn_ReadyStatusSync(1)					
					
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Check-In done successfully")
					JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").SetTOProperty "Index" ,1
					JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").SetTOProperty "Index" ,1
				Case "Summary Toolbar Save properties and Check-In"
						'Selecting  Save properties and Check-In button from Summary Toolbar
						'Wait 2
						Call Fn_ToolbatButtonClick("Check In...")''By Avinash J.as 10.1 change -18-Jan2013
						'Wait 2
						Call Fn_ReadyStatusSync(2)
						If Instr(1,StrTitle,"My Teamcenter") or Instr(1,StrTitle,"Change Manager") or Instr(1,StrTitle,"4G Designer") Then
							Set objCheckIn =Fn_SISW_GetObject("Check-In@2")
						Else
							Set objCheckIn = Fn_UI_ObjectCreate("Fn_SISW_ChkInOut_ObjectCheckIn", JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Check-In"))
						End if
						bSaveButtonExist=False
						For iCounter=0 to 9
							If objCheckIn.Exist(1) Then
								Call Fn_Button_Click("Fn_SISW_ChkInOut_ObjectCheckIn", objCheckIn, "Yes")
								bSaveButtonExist=True
								Call Fn_ReadyStatusSync(1)
								Exit For
							End If
						Next
						If bSaveButtonExist=False Then
							Set objCheckIn= Nothing
							Exit Function
						End If
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Check-In done successfully")						
				'Case "View Menu Save and Check-In"
				Case "Viewer Tab Save and check-In"
						'Select Viewer Tab
						Call Fn_MyTc_TabOperation("Activate", "Viewer")
						Call Fn_ReadyStatusSync(1)
						If Fn_SISW_UI_Object_Operations("Fn_SISW_ChkInOut_ObjectCheckIn", "Exist", Fn_SISW_GetObject("Check-In"), SISW_DEFAULT_TIMEOUT) Then
							' do Nothing
						ElseIf Fn_SISW_UI_Object_Operations("Fn_SISW_ChkInOut_ObjectCheckIn", "Exist", Fn_SISW_GetObject("Check-In@1"), SISW_MICRO_TIMEOUT) Then
							' do Nothing
						ElseIf Fn_SISW_UI_Object_Operations("Fn_SISW_ChkInOut_ObjectCheckIn", "Exist", Fn_SISW_GetObject("Check-In@2"), SISW_MICRO_TIMEOUT) Then
							' do Nothing
						Else
							If Fn_SISW_UI_Object_Operations("Fn_SISW_ChkInOut_ObjectCheckIn", "Exist", JavaWindow("MyTeamcenter").JavaWindow("MyTcJApplet").JavaButton("Save and Check-In"), SISW_MICRO_TIMEOUT) Then
								Call Fn_Button_Click("Fn_SISW_ChkInOut_ObjectCheckIn", JavaWindow("MyTeamcenter").JavaWindow("MyTcJApplet"), "Save and Check-In")
								Call Fn_ReadyStatusSync(1)
							Else
								'buttons are replaced with JavaToolbar after deploying ADS template - Ashok Kakade, Koustubh Watwe
								If Fn_ToolbatButtonClick("Check In...")=False Then
									Call Fn_SISW_UI_JavaToolbar_Operations("Fn_SISW_ChkInOut_ObjectCheckIn", "Click", JavaWindow("MyTeamcenter"), "", "Check In...", "", "", "")
								End If
								Call Fn_ReadyStatusSync(1)
							End If
						End If
						
						If Fn_SISW_UI_Object_Operations("Fn_SISW_ChkInOut_ObjectCheckIn", "Exist", Fn_SISW_GetObject("Check-In"), SISW_DEFAULT_TIMEOUT) Then
							Set objCheckIn =Fn_SISW_GetObject("Check-In")
						ElseIf Fn_SISW_UI_Object_Operations("Fn_SISW_ChkInOut_ObjectCheckIn", "Exist", Fn_SISW_GetObject("Check-In@1"), SISW_MICRO_TIMEOUT) Then
							Set objCheckIn =Fn_SISW_GetObject("Check-In@1")
						ElseIf Fn_SISW_UI_Object_Operations("Fn_SISW_ChkInOut_ObjectCheckIn", "Exist", Fn_SISW_GetObject("Check-In@2"), SISW_MICRO_TIMEOUT) Then
							Set objCheckIn =Fn_SISW_GetObject("Check-In@2")
						Else
							JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").SetTOProperty "Index" ,0
							JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").SetTOProperty "Index" ,0
							If Fn_SISW_UI_Object_Operations("Fn_SISW_ChkInOut_ObjectCheckIn", "Exist", Fn_SISW_GetObject("Check-In"), SISW_DEFAULT_TIMEOUT) Then
							   Set objCheckIn =Fn_SISW_GetObject("Check-In")
							ElseIf Fn_SISW_UI_Object_Operations("Fn_SISW_ChkInOut_ObjectCheckIn", "Exist", Fn_SISW_GetObject("Check-In@1"), SISW_MICRO_TIMEOUT) Then
							   Set objCheckIn =Fn_SISW_GetObject("Check-In@1")
							Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_SISW_ChkInOut_ObjectCheckIn : failed to find Check In dialog.")
							    exit function
							End If
						End If

						Call Fn_Button_Click("Fn_SISW_ChkInOut_ObjectCheckIn", objCheckIn, "Yes")
						Call Fn_ReadyStatusSync(1)
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Check In done successfully")
						JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").SetTOProperty "Index" ,1
						JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").SetTOProperty "Index" ,1
			Case "Check-In with select all component","Check-In with select all component_Ext"
						'invoke Check_In dialog from Main menu
						If sAction = "Check-In with select all component" Then
							Call Fn_MenuOperation("Select", "Tools:Check-In/Out:Check In...")
							Call Fn_ReadyStatusSync(1)
						End If
                    	
                        If Instr(1,StrTitle,"My Teamcenter") or Instr(1,StrTitle,"Change Manager") Then
							Set  objCheckIn =Fn_SISW_GetObject("Check-In@2")
						Else
							Set  objCheckIn = Fn_UI_ObjectCreate("Fn_SISW_ChkInOut_ObjectCheckIn", JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Check-In"))
						End if
						
						bSaveButtonExist=False
						For iCounter=0 to 9
							If objCheckIn.Exist(1) Then
								bSaveButtonExist=True
								Exit For
							End If
						Next
						If bSaveButtonExist=False Then
							Set objCheckIn= Nothing
							Exit Function
						End If
						Call Fn_Button_Click("Fn_SISW_ChkInOut_ObjectCheckIn", objCheckIn, "Explore")
						Call Fn_ReadyStatusSync(1)
						
						'verify Explore dialog exist
						If JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Explore").Exist(2) Then
							Set  objExplore = JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Explore")
						Else 
							Set  objExplore = JavaDialog("Explore")
						End If
						Call Fn_Button_Click("Fn_SISW_ChkInOut_ObjectCheckIn", objExplore, "SelectAll")
                        Call Fn_Button_Click("Fn_SISW_ChkInOut_ObjectCheckIn", objExplore, "OK")

						If aExploreOption <> ""  Then
							'Validate all the objects are added to [Check-In] dialog
							Set objDesFetched_static_text=description.Create()
							objDesFetched_static_text("Class Name").value="JavaStaticText"
							Set objaFetchText=objCheckIn.ChildObjects(objDesFetched_static_text)
							iFetchTextCount=objaFetchText.Count-2 
							iArrayCount=Ubound(aExploreOption)
							iCount=0
							If iFetchTextCount=iArrayCount then
								For iCounter=1 to objaFetchText.Count-1
									If aExploreOption(iCount)=objaFetchText(iCounter).GetRoProperty("attached text") Then
										iCount=iCount+1
									Else
										Fn_SISW_ChkInOut_ObjectCheckIn =FALSE
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Array details and selected items are not equal")	
										Exit Function
									End If
								Next
							Else
								Fn_SISW_ChkInOut_ObjectCheckIn =FALSE
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Array details and selected items are not equal")	
								Exit Function
							End If	
						End If
						Call Fn_Button_Click("Fn_SISW_ChkInOut_ObjectCheckIn", objCheckIn, "Yes")
						Call Fn_ReadyStatusSync(1)
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Check-In done successfully")
						
				Case "Check-In with Error"
					Set objCheckIn=Fn_SISW_ChkInOut_GetChkInChkOutObject("CheckIn")
					If typename(objCheckIn) = "Nothing" Then
						Call Fn_MenuOperation("Select", "Tools:Check-In/Out:Check In...")
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Menu Item Tools:Check-In/Out:Check In... selected successfully")
					End If
					Set objCheckIn= Nothing

						Set objCheckIn=Fn_SISW_ChkInOut_GetChkInChkOutObject("CheckIn")
						If typename(objCheckIn) = "Nothing" Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed Function [Fn_SISW_ChkInOut_ObjectCheckIn]  to find [Check-In] dialog.")
							Set objCheckIn= Nothing
							Exit Function
						End If
											
						Call Fn_Button_Click("Fn_SISW_ChkInOut_ObjectCheckIn", objCheckIn, "Yes")

						If aExploreOption <> "" Then
							'Validate all the objects are added to [Check-In] dialog
							Set objDesFetched_static_text=description.Create()
							objDesFetched_static_text("Class Name").value="JavaStaticText"
							Set objaFetchText=objCheckIn.ChildObjects(objDesFetched_static_text)
							iFetchTextCount=objaFetchText.Count-2 
							iArrayCount=Ubound(aExploreOption)
							iCount=0
							If iFetchTextCount=iArrayCount then
									For iCounter=1 to objaFetchText.Count-1
											If aExploreOption(iCount)=objaFetchText(iCounter).GetRoProperty("attached text") Then
													iCount=iCount+1
											Else
													Fn_SISW_ChkInOut_ObjectCheckIn =FALSE
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Array details and selected items are not equal")	
													Exit Function
											End If
									Next
							Else
									Fn_SISW_ChkInOut_ObjectCheckIn =FALSE
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Array details and selected items are not equal")	
									Exit Function
							End If	
						End If

						'Click on red cross button and verify the error message
						Set objDesBtnCloseRed=description.Create()
						objDesBtnCloseRed("label").value="defaulterror_16"							
						If  objCheckIn.JavaButton(objDesBtnCloseRed).Exist(3) Then
							objCheckIn.JavaButton(objDesBtnCloseRed).Click
						ElseIf objCheckIn.JavaTree("CheckInErrorMsgTree").Exist(1) Then
							objCheckIn.JavaTree("CheckInErrorMsgTree").Activate "#0"
						Else
							objCheckIn.JavaTable("CheckInList").ClickCell 0,1
						End If
						For iCount = 0 to 0
							If JavaWindow("DefaultWindow").JavaWindow("ErrorWindow").Exist Then
								Set objErrorDialog = Fn_UI_ObjectCreate("Fn_SISW_ChkInOut_ObjectCheckIn",JavaWindow("DefaultWindow").JavaWindow("ErrorWindow"))
									sErrorText=objErrorDialog.JavaEdit("DetailsMsg").GetROProperty("value")
									If sErrorMessage<>sErrorText then
											Fn_SISW_ChkInOut_ObjectCheckIn =FALSE
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Error message is not correct")
											Exit Function
									End if					
									Call Fn_Button_Click("Fn_SISW_ChkInOut_ObjectCheckIn", objErrorDialog, "OK")
								Exit For
							End If
							If JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("ErrorDialog").Exist Then
								Set objErrorDialog = Fn_UI_ObjectCreate("Fn_SISW_ChkInOut_ObjectCheckIn",JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("ErrorDialog"))
								If sErrorMessage <> "" Then
									sErrorText=objErrorDialog.JavaEdit("DetailMsg").GetROProperty("value")
									If sErrorMessage <> sErrorText Then
											Fn_SISW_ChkInOut_ObjectCheckIn =FALSE
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Error message is not correct")
											Exit Function
									End If	
								End If										
									Call Fn_Button_Click("Fn_SISW_ChkInOut_ObjectCheckIn", objErrorDialog, "OK")
									Exit For
							End If
						Next
						If objCheckIn.JavaButton("OK").Exist Then
							 Call Fn_Button_Click("Fn_SISW_ChkInOut_ObjectCheckIn", objCheckIn,"OK")
						Else 
							 Call Fn_Button_Click("Fn_SISW_ChkInOut_ObjectCheckIn", objCheckIn,"Yes")
						End If
						 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Check-In done successfully")
				Case "Check-InFromCustomizeToolBar"
						'Selecting  Save properties and Check-In button from Summary Toolbar
						Call Fn_ToolbatButtonClick("Check In...")
						Call Fn_ReadyStatusSync(1)
						
                        If Instr(1,StrTitle,"My Teamcenter") or Instr(1,StrTitle,"Change Manager") Then
							Set  objCheckIn =Fn_SISW_GetObject("Check-In@2")
						Else
							Set  objCheckIn = Fn_UI_ObjectCreate("Fn_SISW_ChkInOut_ObjectCheckIn", JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Check-In"))
						End if
						'Call Fn_Button_Click("Fn_SISW_ChkInOut_ObjectCheckIn", objCheckIn, "Yes")
						
						bSaveButtonExist=False
						For iCounter=0 to 9
							If objCheckIn.Exist(1) Then
								Call Fn_Button_Click("Fn_SISW_ChkInOut_ObjectCheckIn", objCheckIn, "Yes")
								bSaveButtonExist=True
								Call Fn_ReadyStatusSync(1)
								Exit For
							End If
						Next
						If bSaveButtonExist=False Then
							Set objCheckIn= Nothing
							Exit Function
						End If						
						
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Check-In done successfully")
				Case "VerifyChekInErrorMessage"
					'Selecting  Save properties and Check-In button from Summary Toolbar
					If Instr(1,StrTitle,"My Teamcenter") or Instr(1,StrTitle,"Change Manager") Then
						Set  objCheckIn =Fn_SISW_GetObject("Check-In@2")
					Else
						Set  objCheckIn =Fn_SISW_GetObject("Check-In")
					End if
					If NOT objCheckIn.Exist(1) Then
						Call Fn_MenuOperation("Select", "Tools:Check-In/Out:Check In...")
						Call Fn_ReadyStatusSync(1)
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Menu Item Tools:Check-In/Out:Check In... selected successfully")
					End If
					
					If  objCheckIn.Exist(1) Then
						Call Fn_Button_Click("Fn_SISW_ChkInOut_ObjectCheckIn", objCheckIn, "Yes")
						sErrorMsg = objCheckIn.JavaTree("CheckInErrorMsgTree").GetColumnValue (0,1)
						If strComp(trim(sErrorMsg), trim(sErrorMessage)) = 0  Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: suucessfully Verified Error Mesage.")
							Fn_SISW_ChkInOut_ObjectCheckIn = True
						Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Verify Error Mesage.")
						End If
						Call Fn_Button_Click("Fn_SISW_ChkInOut_ObjectCheckIn", objCheckIn, "OK")
					End If
				Case Else
					Fn_SISW_ChkInOut_ObjectCheckIn =FALSE
					Exit Function								
		End Select

		Fn_SISW_ChkInOut_ObjectCheckIn =TRUE
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Check-In done successfully")

		Set objCheckIn=Nothing
  		Set  objExplore=Nothing
		Set  objDesFetched_static_text=Nothing
		Set  objaFetchText=Nothing
		Set   objDesBtnCloseRed=Nothing
		Set	 objErrorDialog=Nothing
End Function

'#######################################################################################
'###     FUNCTION NAME   :   Fn_SISW_ChkInOut_CheckOutMsgVerify(sAction, aObject, aMessage)
'###
'###    DESCRIPTION     :   Verify the error messages popping up after Check-Out operation
'###
'###    PARAMETERS      :   sAction
'###									aObject
'###									aMessage
'###
'###    Return Value  	:   	True/False 
'###
'###    HISTORY         :   	AUTHOR              	DATE        		VERSION
'###
'###    CREATED BY      :   Ketan Raje				15-Sept-2010   			1.0
'###
'###    REVIWED BY      :	Harshal					15-Sept-2010   			
'###
'###    MODIFIED BY     :
'###    EXAMPLE         :   Msgbox Fn_SISW_ChkInOut_CheckOutMsgVerify("ButtonToolTip", "StartTestWC", "Object  is of a type not supported by Check-Out facility.")
'#############################################################################################
Public Function Fn_SISW_ChkInOut_CheckOutMsgVerify(sAction, aObject, aMessage)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_ChkInOut_CheckOutMsgVerify"
Dim objDialog, sObj, txtMsg, sElement, iCount, bFlag, sMsg, iCountChecked,StrTitle,iCounter
iCountChecked = 0
StrTitle=JavaWindow("DefaultWindow").GetROProperty("title")
bFlag=False
If Instr(1,StrTitle,"My Teamcenter") or Instr(1,StrTitle,"Change Manager") Then
	Set objDialog =JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Check-Out")
	If objDialog.Exist(6)=False Then
		Call Fn_MenuOperation("Select","Edit:Properties")
	End if
    bFlag=True
Else
	If JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Check-Out").Exist(6) Then
		Set objDialog = Fn_UI_ObjectCreate("Fn_SISW_ChkInOut_CheckOutMsgVerify", JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Check-Out"))		
	ElseIf JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Check-Out").Exist(2)  Then
		Set objDialog = Fn_UI_ObjectCreate("Fn_SISW_ChkInOut_CheckOutMsgVerify", JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Check-Out"))		
	Else
		Call Fn_MenuOperation("Select","Edit:Properties")
		'@Added by Swapna on 31-July-2013
		If JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Check-Out").Exist(6) Then
			Set objDialog = Fn_UI_ObjectCreate("Fn_SISW_ChkInOut_CheckOutMsgVerify", JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Check-Out"))	
		ElseIf JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Check-Out").Exist(2)  Then
			Set objDialog = Fn_UI_ObjectCreate("Fn_SISW_ChkInOut_CheckOutMsgVerify", JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Check-Out"))		
		End If
	End if
End If

'Click on yes button 
If Fn_UI_ObjectExist("Fn_SISW_ChkInOut_CheckOutMsgVerify", objDialog.JavaButton("ExploerComponents"))=True Then
	Call Fn_Button_Click("Fn_SISW_ChkInOut_CheckOutMsgVerify", objDialog, "Yes")
End If 
'@End

sObj = Split(aObject, ":", -1, 1)
txtMsg = Split(aMessage, ":", -1, 1)
Select Case sAction
    Case "ButtonToolTip"
				If bFlag=True Then
					objDialog.JavaLink("More").SetTOProperty "attached text","More..."
					If objDialog.JavaLink("More").Exist(1) Then
						objDialog.JavaLink("More").Click 1,1,"LEFT"
					End If
					If aObject<>"" Then
						For iCount=0 to ubound(sObj)
							bFlag = False
							For iCounter=0 to Cint(objDialog.JavaTable("CheckOutList").GetROProperty("rows"))-1
								If Trim(objDialog.JavaTable("CheckOutList").Object.getItem(iCounter).getData().getAIFContext().toString())=Trim(sObj(iCount)) Then
									If instr(1,objDialog.JavaTable("CheckOutList").GetCellData(iCounter,1),txtMsg(iCount))>0 Then
										bFlag = True
										Exit for
									End If
								End If
							Next
							If bFlag = False Then
								Exit for
							End If
						Next
					Else
						If instr(1,objDialog.JavaTable("CheckOutList").GetCellData(0,1),txtMsg(iCount))>0 Then
							bFlag = True
						Else
							bFlag = False
						End If
					End If
				Else
					objDialog.JavaButton("CheckOutErrorBtn").SetTOProperty "Index",iCount
					objDialog.Click 1,1,"LEFT"
					sMsg = objDialog.JavaButton("CheckOutErrorBtn").GetROProperty("tool_tip_text")
							If InStr(1, sMsg, txtMsg(iCount), 1) > 0 Then
								bFlag = True
							Else
								bFlag = False
							End If			
				End If		
    Case "Label"
				If InStr(1, objDialog.JavaObject("CheckOutErrorMessage").GetROProperty("attached text"), aMessage, 1) > 0 Then
						bFlag = True
				End If
End Select
	Call Fn_Button_Click("Fn_SISW_ChkInOut_CheckOutMsgVerify", objDialog, "OK")
		If bFlag = True Then
				Fn_SISW_ChkInOut_CheckOutMsgVerify = True
				Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Pass: CheckOut Error Message Verified Successfully. ")
		Else
				Fn_SISW_ChkInOut_CheckOutMsgVerify = False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: CheckOut Error Message Not Verified. ")				
		End If
Set objDialog = Nothing
Set sElement = Nothing
End Function

'################################################################################################################################
'# 	Function Name		:				Fn_SISW_ChkInOut_OptionsSettings()
'#
'#	Description			 :		 		     Perform the functionality of Check-In and Check-Out from General Tab of Options settings.
'#											
'#	Return Value		   : 				TRUE \ FALSE 
'#
'#	Pre-requisite			:		 		 Team Center perspective is open .
'#
'#	Examples				:			     
'#										Developer Name			Date			Rev. No.		Changes Done		Reviewer	
'######################################################################################################################################
'#										Ketan Raje				04-Jan-2011			1											Harshal
'######################################################################################################################################
Public Function Fn_SISW_ChkInOut_OptionsSettings(sAction, bGeneral, sRemoveFiles, sCheckOutDir, sExportFiles, bSysAdmin, sAutoCheckOut, sButtons)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_ChkInOut_OptionsSettings"
	Dim ObjDialog, aButtons, iCnt
    'Select menu [Edit  -> Options...]
	If Not Fn_UI_ObjectExist("Fn_SISW_ChkInOut_OptionsSettings",JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Options")) Then
			Call Fn_MenuOperation("Select","Edit:Options...")
	End If    
	Call Fn_ReadyStatusSync(1)
	Set ObjDialog = Fn_UI_ObjectCreate("Fn_SISW_ChkInOut_OptionsSettings", JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Options"))
	If Not Fn_UI_ObjectExist("Fn_SISW_ChkInOut_OptionsSettings",JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Options").JavaTab("ItemTab")) Then
			'Select General : Check-In/Check-Out under OptionsTree.
			Call Fn_JavaTree_Select("Fn_SISW_ChkInOut_OptionsSettings", JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Options"), "OptionsTree","Options:General:Check-In/Check-Out")
	End If
	Call Fn_ReadyStatusSync(1)
		Select Case sAction
				Case "Set"
							If Trim(Lcase(bGeneral)) = "true" Then
								'Select General Tab.
								Call Fn_UI_JavaTab_Select("Fn_SISW_ChkInOut_OptionsSettings", ObjDialog, "ItemTab", "General")
								'Set the check box for Remove file on checkIn.
								If sRemoveFiles <> "" Then
									Call Fn_CheckBox_Set("Fn_SISW_ChkInOut_OptionsSettings",ObjDialog,"RemoveFilesOnCheckIn",sRemoveFiles)					
								End If
								'Set the Edit box for Check Out Dir.
								If sCheckOutDir <> "" Then
									Call Fn_Edit_Box("Fn_SISW_ChkInOut_OptionsSettings",ObjDialog,"CheckOutDirectory",sCheckOutDir)
								End If													
								'Set the check box for Export file on checkOut.
								If sExportFiles <> "" Then
									Call Fn_CheckBox_Set("Fn_SISW_ChkInOut_OptionsSettings",ObjDialog,"ExportFilesOnCheckOut",sExportFiles)					
								End If
							End If
							If Trim(Lcase(bSysAdmin)) = "true" Then
								'Select SysAdmin Tab.
								Call Fn_UI_JavaTab_Select("Fn_SISW_ChkInOut_OptionsSettings", ObjDialog, "ItemTab", "Sys Admin")
								'Set the check box for Auto Check Out.
								If sAutoCheckOut <> "" Then
									Call Fn_CheckBox_Set("Fn_SISW_ChkInOut_OptionsSettings",ObjDialog,"AutoCheckOut",sAutoCheckOut)					
								End If
							End If
							Fn_SISW_ChkInOut_OptionsSettings = True
							Call Fn_WriteLogFile(Environment.Value("TestLogFile")," CheckIn and CheckOut options set successfully" )		
				Case Else
							Fn_SISW_ChkInOut_OptionsSettings = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_SISW_ChkInOut_OptionsSettings function failed")
		End Select
		'Click on Buttons
		If sButtons<>"" Then
				aButtons = split(sButtons, ":",-1,1)
				For iCnt=0 to Ubound(aButtons)
					'Click on Buttons
					Call Fn_Button_Click("Fn_SISW_ChkInOut_OptionsSettings", ObjDialog, aButtons(iCnt))
				Next
		End If
		Set ObjDialog = Nothing		
End Function

'*********************************************************		Function to perform operations on CheckOutHistory Table***********************************************
'Function Name		:			Fn_SISW_ChkInOut_ChkOutHistoryVerification

'Description			 :		 	  Perform operations on CheckOutHistory Table

'Return Value		   : 			-1 if False And Zero or positive number if True.

'Pre-requisite			:		 	Item Should be Selected on which have to perform the operations

'Examples				:			Case "Verify"  : Fn_SISW_ChkInOut_ChkOutHistoryVerification("Verify", "", "", "AutoTestDBA (autotestdba)", "Check-Out", "3", "000293/A;2-item", "")

'History					 :			
'										Developer Name				Date						Rev. No.							Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Ketan Raje					25-Apr-2011			           1.0									Harshal A
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public function Fn_SISW_ChkInOut_ChkOutHistoryVerification(sAction, sItem, sDateTime, sUser, sActivity, sChangeID, sComments, aGlobalDic)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_ChkInOut_ChkOutHistoryVerification"
	'Function Returning False
	Fn_SISW_ChkInOut_ChkOutHistoryVerification=-1
	'Declaring Variables
	Dim iFlag, ChkOutHistoryTable, iRowCnt, iCount, iCnt, iRow, arrCols, iColIndex
	'Checking Existance Of ChkOutHistory
	If Fn_UI_ObjectExist("Fn_SISW_ChkInOut_ChkOutHistoryVerification", JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Check-Out History").JavaTable("ChkOutHistory"))=False Then
		'Calling Menu Operation to open ChkOutHistory Table Dialog
		Call Fn_MenuOperation("Select","Tools:Check-In/Out:Checkout History...")
	End If
	'Creating Object of Check-Out History Dialog
	Set ChkOutHistoryTable=Fn_UI_ObjectCreate("Fn_SISW_ChkInOut_ChkOutHistoryVerification", JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Check-Out History"))
	'Selecting Action
	Select Case sAction
		'Case For verifying activity done by sUserName
		Case "GetColumnIndex"
			arrCols = split(ChkOutHistoryTable.JavaTable("ChkOutHistory").GetROProperty("column names"), ";")
			Fn_SISW_ChkInOut_ChkOutHistoryVerification = -1
			For iCount = 0 To UBound(arrCols) Step 1
				If sItem = arrCols(iCount) Then
					Fn_SISW_ChkInOut_ChkOutHistoryVerification = iCount
					Exit function
				End If
			Next 
			
		Case "Verify"
					'Taking number of rows from ChkOutHistory Table
					iRowCnt = Fn_Table_GetRowCount("Fn_SISW_ChkInOut_ChkOutHistoryVerification",ChkOutHistoryTable,"ChkOutHistory")
					For iRow = 0 to iRowCnt-1
						iCount = 0
						iCnt = 0
							'Verify the value of Date/Time
							If sDateTime<>"" Then
								iCount = iCount + 1
								iColIndex = Fn_SISW_ChkInOut_ChkOutHistoryVerification("GetColumnIndex", "Date/Time", "", "", "", "", "", "")
								If Trim(Lcase(sDateTime)) = Trim(Lcase(JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Check-Out History").JavaTable("ChkOutHistory").GetCellData(iRow,iColIndex))) Then
									iCnt = iCnt + 1
								End If
							End If
							'Verify the value of User
							If sUser<>"" Then
								iCount = iCount + 1
								iColIndex = Fn_SISW_ChkInOut_ChkOutHistoryVerification("GetColumnIndex", "User", "", "", "", "", "", "")
								If Trim(Lcase(sUser)) = Trim(Lcase(JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Check-Out History").JavaTable("ChkOutHistory").GetCellData(iRow, iColIndex))) Then
									iCnt = iCnt + 1
								End If
							End If
							'Verify the value of Activity
							If sActivity<>"" Then
								iCount = iCount + 1
								iColIndex = Fn_SISW_ChkInOut_ChkOutHistoryVerification("GetColumnIndex", "Activity", "", "", "", "", "", "")
								If Trim(Lcase(sActivity)) = Trim(Lcase(JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Check-Out History").JavaTable("ChkOutHistory").GetCellData(iRow,iColIndex))) Then
									iCnt = iCnt + 1
								End If
							End If
							'Verify the value of ChangeID
							If sChangeID<>"" Then
								iCount = iCount + 1
								iColIndex = Fn_SISW_ChkInOut_ChkOutHistoryVerification("GetColumnIndex", "Change ID", "", "", "", "", "", "")
								If Trim(Lcase(sChangeID)) = Trim(Lcase(JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Check-Out History").JavaTable("ChkOutHistory").GetCellData(iRow,iColIndex))) Then
									iCnt = iCnt + 1
								End If
							End If
							'Verify the value of Comments
							If sComments<>"" Then
								iCount = iCount + 1
								iColIndex = Fn_SISW_ChkInOut_ChkOutHistoryVerification("GetColumnIndex", "Comments", "", "", "", "", "", "")
								If Trim(Lcase(sComments)) = Trim(Lcase(JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Check-Out History").JavaTable("ChkOutHistory").GetCellData(iRow, iColIndex))) Then
									iCnt = iCnt + 1
								End If
							End If
						If iCount=iCnt Then
							Exit For
						End If
					Next
					If iCount=iCnt Then
						Fn_SISW_ChkInOut_ChkOutHistoryVerification = iRow
					Else
						Fn_SISW_ChkInOut_ChkOutHistoryVerification = -1
					End If
	End Select
	'Closing Table
    Call Fn_Button_Click("Fn_SISW_ChkInOut_ChkOutHistoryVerification",ChkOutHistoryTable,"Close")
	'Setting objects to nothing
	Set ChkOutHistoryTable=Nothing
End Function

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_SISW_ChkInOut_GetChkInChkOutObject()
'
'Description			 :	

'Parameters			   :  

'Return Value		   : 	True or False

'Pre-requisite			:	 Login to RAC Client

'Examples				:  	Call Fn_SISW_ChkInOut_GetChkInChkOutObject("CheckOut")
'									
'History				   :			
'				Developer Name						Date					Rev. No.						Changes Done																																Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'				Pritam Shikare						23-08-2013 				1.0																																																												Sandeep	
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'				Pranav Ingle						 5-Feb-2014 			 1.1							Modified Code to Take  only new Dialog  in "MyTC" &  "CM"
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Public Function Fn_SISW_ChkInOut_GetChkInChkOutObject(sAction)
	
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_ChkInOut_GetChkInChkOutObject"
	Dim StrTitle, sDialogName
	StrTitle=JavaWindow("DefaultWindow").GetROProperty("title")

   	Select Case sAction
		Case "CheckOut"
			sDialogName =  "Check-Out"
		Case "CheckIn"
			sDialogName = "Check-In"
		Case "CancelCheckOut"
			sDialogName = "Cancel Check-Out"
	End Select

	'  Commented  code for other Object Hierarchies for  perspective  so it will take only  new Dialog in  "MyTC" &  "CM"
	If Instr(1,StrTitle,"My Teamcenter") or Instr(1,StrTitle,"Change Manager") Then
'		Select Case gLastMenuCall
'			Case "Edit:Properties", "View:Properties", "View Properties	Alt+Enter", "Edit Properties","Edit","SaveAndCheckIn","CheckOutAndEdit"
'	
'				If JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog(sDialogName).Exist(SISW_MIN_TIMEOUT) Then
'						Set Fn_SISW_ChkInOut_GetChkInChkOutObject = JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog(sDialogName)
'				ElseIf JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog(sDialogName).Exist(SISW_MICRO_TIMEOUT) Then
'						Set Fn_SISW_ChkInOut_GetChkInChkOutObject = JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog(sDialogName)
'				Else
'						Set Fn_SISW_ChkInOut_GetChkInChkOutObject = Nothing
'				End If
'
'			Case Else
                If JavaWindow("DefaultWindow").JavaWindow(sDialogName).Exist(2)  Then
					Set Fn_SISW_ChkInOut_GetChkInChkOutObject = JavaWindow("DefaultWindow").JavaWindow(sDialogName)
				ElseIf JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("CheckOut").Exist(1)  Then
					Set Fn_SISW_ChkInOut_GetChkInChkOutObject =JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("CheckOut")						
				ElseIf  JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Edit Properties").JavaDialog("Check-In").Exist(1) Then
					Set Fn_SISW_ChkInOut_GetChkInChkOutObject=JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Edit Properties").JavaDialog("Check-In")

'				ElseIf sDialogName = "Check-In" Then
'					If JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Edit Properties").Exist(1) Then  ' Pranav - Added Case where above menus are not called
'						Set Fn_SISW_ChkInOut_GetChkInChkOutObject = JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog(sDialogName)
'					Else
'						Set Fn_SISW_ChkInOut_GetChkInChkOutObject = Nothing
'					End If
				Else
					Set Fn_SISW_ChkInOut_GetChkInChkOutObject = Nothing
				End If		
'		End Select
	Else
		If JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog(sDialogName).Exist(SISW_MIN_TIMEOUT) Then
			Set Fn_SISW_ChkInOut_GetChkInChkOutObject = JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog(sDialogName)
		ElseIf JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog(sDialogName).Exist(SISW_MICRO_TIMEOUT) Then
			Set Fn_SISW_ChkInOut_GetChkInChkOutObject = JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog(sDialogName)
		ElseIf JavaWindow("DefaultWindow").JavaWindow(sDialogName).Exist(2)  Then
			Set Fn_SISW_ChkInOut_GetChkInChkOutObject = JavaWindow("DefaultWindow").JavaWindow(sDialogName)
		Else
			Set Fn_SISW_ChkInOut_GetChkInChkOutObject = Nothing
		End If
	End If
	
End Function
