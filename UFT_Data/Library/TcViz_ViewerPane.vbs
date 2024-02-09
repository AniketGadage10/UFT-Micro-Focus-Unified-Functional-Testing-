Option Explicit

'=======================================================================================================================================================
' Function List
'-------------------------------------------------------------------------------------------------------------------------------------------------------
'						Function Name																		|					Created By
'-------------------------------------------------------------------------------------------------------------------------------------------------------
'1. Fn_SISW_TcViz_GetObject()												|	Vallari S (vallari.shimpukade@siemens.com)
'2. Fn_SISW_TcViz_ToolbarOperation()										|	Vallari S (vallari.shimpukade@siemens.com)
'3. Fn_SISW_TcViz_ProductView_Operations()									|	Vallari S (vallari.shimpukade@siemens.com)
'4. Fn_SISW_TcViz_Options_Setting()											|	Vallari S (vallari.shimpukade@siemens.com)
'5. Fn_SISW_GraphicsTab_3DViewOperation
'6. Fn_SISW_TcViz_CheckOutProductView										|	Reema W
'7. Fn_SISW_TcViz_CheckInProductView										|	Reema W
'8. Fn_SISW_TcViz_ImageInViewerOperations										|   Rinki A
'9. Fn_SISW_TcViz_MyTcAssemblyTreeOperations									|   Reema W
'=======================================================================================================================================================
'****************************************    Function to get Object hierarchy ***************************************
'
''Function Name		 	:	Fn_SISW_TcViz_GetObject
'
''Description		    :  	Function to get specified Object hierarchy.

''Parameters		    :	1. sObjectName : Object Handle name
								
''Return Value		    :  	Object \ Nothing
'
''Examples		     	:	Fn_SISW_TcViz_GetObject("NewProductView")

'History:
'	Developer Name			Date			Rev. No.		Reviewer		Changes Done	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Vallari S			 4-Feb-2013			1.0				
'-----------------------------------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------------------------------
'	Anurag Khera		 26-Sept-2013		2.0								Externalized object hierarchies
'-----------------------------------------------------------------------------------------------------------------------------------
Function Fn_SISW_TcViz_GetObject(sObjectName)
	Dim sObjectXMLPath
	sObjectXMLPath = Fn_GetEnvValue("User", "AutomationDir") & "\TestData\AutomationXML\ObjectXML\TcViz_ViewerPane.xml"
	Set Fn_SISW_TcViz_GetObject = Fn_SISW_Setup_GetObjectFromXML(sObjectXMLPath, sObjectName)
End Function 
'--------------------------------------------------------------------------------------------------------------------
' Function Number   	: 2                                                                                
' Function Name     	: Fn_SISW_TcViz_myTc_ToolbarOperation
' Function Description  : Saves Session
' Function Usage    	: Result = Fn_SISW_TcViz_myTc_ToolbarOperation(sAction, sToolbarName, sButtonName)
'							
'                     		return True/False
'--------------------------------------------------------------------------------------------------------------------
'	Developer Name			Date			Rev. No.		Reviewer		Changes Done	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Reema W			 		20-Aug-2014			1.0							modified Case "Enable", "Disable" to use Fn_SISW_Window_ContextMenu_Operation
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Vivek A				25-May-2015				1.0							Modified Case 
'-----------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_TcViz_ToolbarOperation(sAction, sToolbarName, sButtonName)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_TcViz_ToolbarOperation"
   	Dim strMenu, WShell
   	Dim arrItems, iCnt,arrToolbarName, AbstractInsightObj
   	Dim objImageCanvas, sToolbar, bFlag, objJavaApplet
   	Dim sToolbarButton, objToolbar, sHeight, sWidth, x, y, iCounter, iDevider, iButtonID
   	Dim iCount, iLoop, midy
	
	IF JavaWindow("TcVizMainWin").JavaWindow("JApplet").JavaObject("ImageCanvas").Exist(3) Then
		Set objImageCanvas = JavaWindow("TcVizMainWin").JavaWindow("JApplet").JavaObject("ImageCanvas")
	End If

	Fn_SISW_TcViz_ToolbarOperation = False
  	Set AbstractInsightObj=Window("VizWindow").InsightObject("InsightObject_2")
  	
	Select Case sAction
			Case "ButtonClick"
					Select Case sToolbarName
							Case "2D Viewing"
	

							Case "2D Markup"
								ReDim arrItems(19)
								arrItems(1) = "Enable Markup"
								arrItems(2) = "Select"
								arrItems(3) = "Freehand Marker"
								arrItems(4) = "Intersection Marker"
								arrItems(5) = "Midpoint Marker"
								arrItems(6) = "Centerpoint Marker"
								arrItems(7) = "Freehand Line"
								arrItems(8) = "Leader Line"
								arrItems(9) = "Line"
								arrItems(10) = "Polyline"
								arrItems(11) = "Ellipse"
								arrItems(12) = "Polygon"
								arrItems(13) = "Rectangle"
								arrItems(14) = "Restricted Text"
								arrItems(15) = "Unrestricted Text"
								arrItems(16) = "Inset Image"
								arrItems(17) = "Rubber Stamp"
								arrItems(18) = "New Layer"
								arrItems(19) = "Markup Preferences"		
								

							Case "MyTc:Create Markup"
								If objImageCanvas.Exist(3) Then
									objImageCanvas.Object.getViewerBean.setToolbarVisibility "Create Markup", True
									Wait(1)
								End If
								
								ReDim arrItems(5)
								arrItems(1) = "Save 3D Layers"
								arrItems(2) = "Create 2D Markup"
								arrItems(3) = "Image Capture"
								arrItems(4) = "Show Hide PMI Tree"
								arrItems(5) = "Show Hide Assembly Tree"
							
							Case "Create Markup"
								If objImageCanvas.Exist(3) Then
									objImageCanvas.Object.getViewerBean.setToolbarVisibility "Create Markup", True
									Wait(1)
								End If
								
								ReDim arrItems(3)
								arrItems(1) = "Image Capture"
								arrItems(2) = "Create 3D Product Views"
								arrItems(3) = "Show Hide PMI Tree"
								
							Case "3D Markup"
								If objImageCanvas.Exist(3) Then
									objImageCanvas.Object.getViewerBean.setToolbarVisibility "Markup3D", True
									Wait(1)
								End If
								
								ReDim arrItems(19)
								arrItems(1) = "3D Markup"
								arrItems(2) = "New Layer"
								arrItems(3) = "Select"
								arrItems(4) = "Freehand Line"
								arrItems(5) = "Line"
								arrItems(6) = "Polyline"
								arrItems(7) = "Ellipse"
								arrItems(8) = "Polygon"
								arrItems(9) = "Rectangle"
								arrItems(10) = "Inset Image"
								arrItems(11) = "Text"
								arrItems(12) = "Anchor Mode"
								arrItems(13) = "Fill Mode"
								arrItems(14) = "Auto Create"
								arrItems(15) = "Use Pre-defined Text Mode"
								arrItems(16) = "Align"
								arrItems(17) = "Distribute"
								arrItems(18) = "Position"
								arrItems(19) = "Resequence Callout"
								
							Case "3D GDT Markup"
								If objImageCanvas.Exist(3) Then
									objImageCanvas.Object.getViewerBean.setToolbarVisibility "3D GDTMarkupUI", True
									Wait(1)
								End If
								
								ReDim arrItems(8)
								arrItems(1) = "GDT Markup"
								arrItems(2) = "GDT Annotation Editor"
								arrItems(3) = "Anchor Mode"
								arrItems(4) = "Stack Mode"
								arrItems(5) = "Copy GDT Annotation"
								arrItems(6) = "Paste GDT Annotation"
								arrItems(7) = "New Layer"
								arrItems(8) = "Preferences"
					End Select
					
					If inStr(1,sToolbarName,":")  > 0 Then
						arrToolbarName = Split(sToolbarName,":")
						If arrToolbarName(1) <> "" Then
							sToolbarName = arrToolbarName(1)	
						End If
					End If
					
					For iCnt = 1 To Ubound(arrItems) Step 1
						If lcase(trim(sButtonName)) = lcase(trim(arrItems(iCnt))) Then
							If Window("VizWindow").WinToolbar(sToolbarName).Exist(2) Then
								Window("VizWindow").WinToolbar(sToolbarName).Press iCnt, micLeftBtn
							ElseIf Window("VizWindow").WinObject(sToolbarName).Exist(2) Then
								Select Case sToolbarName
									Case "3D GDT Markup", "3D Markup", "Create Markup"
											If sToolbarName = "3D Markup" AND arrItems(iCnt)<>"Preferences" Then
												Set objJavaApplet = JavaWindow("TcVizMainWin").JavaWindow("JApplet")
												If Instr(JavaWindow("DefaultWindow").GetROProperty("title"), "Multi-Structure Manager" )>0 Then
													objJavaApplet.JavaButton("BasicSplitPaneDivider").SetToProperty "index", "2"
													Wait 1
													objJavaApplet.JavaButton("BasicSplitPaneDivider").Click
													Wait 1
												ElseIf Instr(JavaWindow("DefaultWindow").GetROProperty("title"), "Structure Manager" )>0 Then
													objJavaApplet.JavaButton("BasicSplitPaneDivider").SetToProperty "index", "0"
													Wait 1
													objJavaApplet.JavaButton("BasicSplitPaneDivider").Click
													Wait 1
												ElseIf Instr(JavaWindow("DefaultWindow").GetROProperty("title"), "Manufacturing Process Planner" )>0 Then
													Call Fn_MPP_TabOperations("DoubleClick", "Graphics")
													Wait 1
												ElseIf Instr(JavaWindow("DefaultWindow").GetROProperty("title"), "My Teamcenter" )>0 Then
													Call Fn_TabFolder_Operation("DoubleClickTab", "Viewer", "")
													Wait 1
												End If
											End If
																																

											Set objToolbar = Window("VizWindow").WinObject(sToolbarName)
											sToolbarButton = arrItems(iCnt)
											
											sHeight = objToolbar.GetROProperty("height")
											
											If Cint(sHeight)>35 Then
												midy = Cint(sHeight /2)
												y = Cint(sHeight /4)
												iLoop = 2
											Else 
												y = Cint(sHeight /2)
												iLoop = 1
											End If
											'y = Cint(sHeight /2)
											
											sWidth = objToolbar.GetROProperty("width")
											iDevider = Cint(sWidth/10)
											xFact = Cint(sWidth/iDevider)
											
											For iCount = 1 To iLoop
												If iCnt > 1 Then
													y = y+midy
												End If
												For iCounter = 1 To iDevider
													x = iCounter*xFact
													
													objToolbar.MouseMove x,y
													'Wait 1
													If Window("nativeclass:=tooltips_class32").Exist(1) Then
														bResult = Window("nativeclass:=tooltips_class32").GetROProperty("text")
														If bResult = sToolbarButton Then
															objToolbar.Click x,y
															Exit For
														End If
													End If
												Next
											Next
											
											If sToolbarName = "3D Markup" AND arrItems(iCnt)<>"Preferences" Then
												If Instr(JavaWindow("DefaultWindow").GetROProperty("title"), "Multi-Structure Manager" )>0 Then
													objJavaApplet.JavaButton("BasicSplitPaneDivider").SetToProperty "index", "3"
													Wait 1
													objJavaApplet.JavaButton("BasicSplitPaneDivider").Click
													Wait 1
												ElseIf Instr(JavaWindow("DefaultWindow").GetROProperty("title"), "Structure Manager" )>0 Then
													objJavaApplet.JavaButton("BasicSplitPaneDivider").SetToProperty "index", "1"
													Wait 1
													objJavaApplet.JavaButton("BasicSplitPaneDivider").Click
													Wait 1
												ElseIf Instr(JavaWindow("DefaultWindow").GetROProperty("title"), "Manufacturing Process Planner" )>0 Then
													Call Fn_MPP_TabOperations("DoubleClick", "Graphics")
													Wait 1
												ElseIf Instr(JavaWindow("DefaultWindow").GetROProperty("title"), "My Teamcenter" )>0 Then
													Call Fn_TabFolder_Operation("DoubleClickTab", "Viewer", "")
													Wait 1
												End If
												Set objJavaApplet = Nothing
											End If
											Set objToolbar = Nothing		
									Case Else 
											AbstractInsightObj.SetTOProperty "ImgSrc", Environment.Value("sPath")+"\TestData\VizInsightImages\"+arrItems(iCnt)+".jpg"
											Wait 5
											AbstractInsightObj.Click
								End Select					
							End If
							Exit For
						End If
					Next

					If Err.Number < 0  OR iCnt > Ubound(arrItems) Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Click [" + sButtonName + "] of [" + sToolbarName + "] Toolbar")
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Successfully Clicked [" + sButtonName + "] of [" + sToolbarName + "] Toolbar")
						Fn_SISW_TcViz_ToolbarOperation = True
					End If
					
			Case "IsButtonChecked"

					Select Case sToolbarName
							Case "3D Markup"
									If objImageCanvas.Exist(3) Then
										objImageCanvas.Object.getViewerBean.setToolbarVisibility "Markup3D", True
										Wait(1)
									End If
									
									ReDim arrItems(19)
									arrItems(1) = "3D Markup"
									arrItems(2) = "New Layer"
									arrItems(3) = "Select"
									arrItems(4) = "Freehand Line"
									arrItems(5) = "Line"
									arrItems(6) = "Polyline"
									arrItems(7) = "Ellipse"
									arrItems(8) = "Polygon"
									arrItems(9) = "Rectangle"
									arrItems(10) = "Inset Image"
									arrItems(11) = "Text"
									arrItems(12) = "Anchor Mode"
									arrItems(13) = "Fill Mode"
									arrItems(14) = "Auto Create"
									arrItems(15) = "Use Pre-defined Text Mode"
									arrItems(16) = "Align"
									arrItems(17) = "Distribute"
									arrItems(18) = "Position"
									arrItems(19) = "Resequence Callout"
									
							Case "3D GDT Markup"
									If objImageCanvas.Exist(3) Then
										objImageCanvas.Object.getViewerBean.setToolbarVisibility "3D GDTMarkupUI", True
										Wait(1)
									End If
									
									ReDim arrItems(8)
									arrItems(1) = "GDT Markup"
									arrItems(2) = "GDT Annotation Editor"
									arrItems(3) = "Anchor Mode"
									arrItems(4) = "Stack Mode"
									arrItems(5) = "Copy GDT Annotation"
									arrItems(6) = "Paste GDT Annotation"
									arrItems(7) = "New Layer"
									arrItems(8) = "Preferences"
							
					End Select
					
					For iCnt = 1 To Ubound(arrItems) Step 1
						If lcase(trim(sButtonName)) = lcase(trim(arrItems(iCnt))) Then
							If Window("VizWindow").WinToolbar(sToolbarName).Exist(2) Then
								bFlag = Window("VizWindow").WinToolbar(sToolbarName).GetItemProperty(iCnt, "Checked")
							Else
								If sToolbarName = "3D Markup" Then
									Select Case arrItems(iCnt)
										Case "Text"
												iButtonID = JavaWindow("TcVizMainWin").JavaWindow("JApplet").JavaObject("JTViewer").object.getViewer3DBean.getMarkup3D().ID_TEXT
										Case "Anchor Mode"
												iButtonID = JavaWindow("TcVizMainWin").JavaWindow("JApplet").JavaObject("JTViewer").object.getViewer3DBean.getMarkup3D().ID_ANCHOR_MODE
										Case "Line"
												iButtonID = JavaWindow("TcVizMainWin").JavaWindow("JApplet").JavaObject("JTViewer").object.getViewer3DBean.getMarkup3D().ID_LINE
										Case "Freehand Line"
												iButtonID = JavaWindow("TcVizMainWin").JavaWindow("JApplet").JavaObject("JTViewer").object.getViewer3DBean.getMarkup3D().ID_FREEHAND_LINE
										Case "Ellipse"
												iButtonID = JavaWindow("TcVizMainWin").JavaWindow("JApplet").JavaObject("JTViewer").object.getViewer3DBean.getMarkup3D().ID_ELLIPSE
										Case "Rectangle"
												iButtonID = JavaWindow("TcVizMainWin").JavaWindow("JApplet").JavaObject("JTViewer").object.getViewer3DBean.getMarkup3D().ID_RECTANGLE												
									End Select
								ElseIf sToolbarName = "3D GDT Markup" Then
									Select Case arrItems(iCnt)
										Case "Anchor Mode"
												iButtonID = JavaWindow("TcVizMainWin").JavaWindow("JApplet").JavaObject("JTViewer").object.getViewer3DBean.getGDTMarkup3D().GDTMARKUP3D_ID_ANCHOR 
										Case "GDT Annotation Editor"
												iButtonID = JavaWindow("TcVizMainWin").JavaWindow("JApplet").JavaObject("JTViewer").object.getViewer3DBean.getGDTMarkup3D().GDTMARKUP3D_ID_EDITOR
									End Select
								End If
								
								If sToolbarName = "3D Markup" Then
									If arrItems(iCnt)="3D Markup" Then ' add images for checked buttons to execute this code in \TestData\VizInsightImages\ folder
										AbstractInsightObj.SetTOProperty "ImgSrc", Environment.Value("sPath")+"\TestData\VizInsightImages\"+arrItems(iCnt)+"_Checked.jpg"
										Wait 2
										If AbstractInsightObj.exist(5) Then
											bFlag = True
										Else
											bFlag = False
										End If
									ElseIf arrItems(iCnt)="Text" OR arrItems(iCnt)="Anchor Mode" OR arrItems(iCnt)="Line" OR arrItems(iCnt)="Freehand Line" OR arrItems(iCnt)="Ellipse" OR arrItems(iCnt)="Rectangle" Then
										bFlag = JavaWindow("TcVizMainWin").JavaWindow("JApplet").JavaObject("JTViewer").object.getViewer3DBean.getMarkup3D().isCmdIDPressed(Cint(iButtonID))
									End If
								ElseIf  sToolbarName = "3D GDT Markup" Then
									If arrItems(iCnt)="GDT Markup" Then
										AbstractInsightObj.SetTOProperty "ImgSrc", Environment.Value("sPath")+"\TestData\VizInsightImages\"+arrItems(iCnt)+"_Checked.jpg"
										Wait 2
										If AbstractInsightObj.exist(5) Then
											bFlag = True
										Else
											bFlag = False
										End If
									ElseIf arrItems(iCnt)="Anchor Mode" OR arrItems(iCnt)="GDT Annotation Editor"  Then
										bFlag = JavaWindow("TcVizMainWin").JavaWindow("JApplet").JavaObject("JTViewer").object.getViewer3DBean.getGDTMarkup3D().isCmdIDPressed(Cint(iButtonID))
									End If
								End If
								
							End If
							Exit For
						End If
					Next
									
					If Trim(LCase(bFlag)) = Trim(LCase(True)) Then
						Fn_SISW_TcViz_ToolbarOperation = True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Successfully Verified [" + sButtonName + "] of [" + sToolbarName + "] Toolbar is Checked")
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to verify  [" + sButtonName + "] of [" + sToolbarName + "] Toolbar is Checked")
						Fn_SISW_TcViz_ToolbarOperation = False
					End If

			Case "Enable"
										
					If Not Window("VizWindow").WinToolbar(sToolbarName).Exist(3) OR Not Window("VizWindow").WinObject(sToolbarName).Exist(3) Then
						If Window("VizWindow").WinToolbar("2D Viewing").Exist(2) Then
							Window("TcVizStructureManager").WinObject("2DImageViewer").Click 10,10,micLeftBtn
		                    			Window("VizWindow").WinToolbar("2D Viewing").Click 3,4,micRightBtn
		                		Elseif Window("VizWindow").WinToolbar("3D Navigation").Exist(1) Then
		                    			Window("TcVizStructureManager").WinObject("3DImageViewer").Click 10,10,micLeftBtn
							Window("VizWindow").WinToolbar("3D Navigation").Click 3,4,micRightBtn
		                		Elseif Window("VizWindow").WinToolbar("Create Markup").Exist(1) Then
		                    			Window("TcVizStructureManager").WinObject("3DImageViewer").Click 10,10,micLeftBtn
							Window("VizWindow").WinToolbar("Create Markup").Click 3,4,micRightBtn
						ElseIf sToolbarName="3D Markup" OR sToolbarName="3D GDT Markup" OR sToolbarName="3d Alignment" OR sToolbarName="3D CAE Viewing" OR sToolbarName="3D Clearance" OR sToolbarName="3D Comparison" OR sToolbarName="3D Coordinate System" OR sToolbarName="3D Measurement" Then
							Select Case sToolbarName
									Case "3D Markup"
										sToolbar = "Markup3D"
									Case "3D GDT Markup"
										sToolbar = "3D GDTMarkupUI"
									Case "3d Alignment"
										sToolbar = "VPAlignment"
									Case "3D CAE Viewing"
										sToolbar = "CAE Viewing"
									Case "3D Clearance"
										sToolbar = "ClearanceFramework"
									Case "3D Comparison"
										sToolbar = "VPCompare"
									Case "3D Coordinate System"
										sToolbar = "VPCoordinateSystem"
									Case "3D Measurement"
										sToolbar = "VPMeasurement"
							End Select
							objImageCanvas.Object.getViewerBean.setToolbarVisibility sToolbar, True
							Wait(5)
							bFlag = True
						ElseIf sToolbarName="3D Section" OR sToolbarName="3D Appearance" OR sToolbarName="3D Constraints" OR sToolbarName="3D Display Modes" OR sToolbarName="3D Movie Capture" OR sToolbarName="3D Navigation" OR sToolbarName="3D PMI" OR sToolbarName="3D Selection" OR sToolbarName="3D Standard Views" OR sToolbarName="3D Thrustline Editor" OR sToolbarName="3d Visibility" OR sToolbarName="3D Visual Report" OR sToolbarName="Create Markup" Then
							objImageCanvas.Object.getViewerBean.setToolbarVisibility sToolbarName, True
							Wait(5)
							bFlag = True
						Else
							Window("VizWindow").InsightObject("InsightObject").SetTOProperty "Index", 0
							Window("VizWindow").InsightObject("InsightObject").Click 5,5,micRightBtn
						End If
						
						If bFlag = False Then
							Set WShell = CreateObject("WScript.Shell")		
							WShell.SendKeys "{DOWN}"
							Call  Fn_SISW_Window_ContextMenu_Operation("Select", sToolbarName,"")	'' added code to selct context menu
							Set WShell = Nothing
						End If
					End If

					If Window("VizWindow").WinToolbar(sToolbarName).Exist(10) Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Successfully Enabled Toolbar [" + sToolbarName + "]")
						Fn_SISW_TcViz_ToolbarOperation = True
					ElseIf Window("VizWindow").WinObject(sToolbarName).Exist(5) Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Successfully Enabled Toolbar [" + sToolbarName + "]")
						Fn_SISW_TcViz_ToolbarOperation = True
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Enable Toolbar [" + sToolbarName + "]")
					End If
					Set WShell = nothing
			Case "Disable"
					bFlag = False
					If Window("VizWindow").WinToolbar(sToolbarName).Exist(10) OR Window("VizWindow").WinObject(sToolbarName).Exist(10) Then
						If Window("VizWindow").WinToolbar("2D Viewing").Exist(3) Then
							Window("TcVizStructureManager").WinObject("2DImageViewer").Click 10,10,micLeftBtn
                        				Window("VizWindow").WinToolbar("2D Viewing").Click 3,4,micRightBtn
                            			Elseif Window("VizWindow").WinToolbar("3D Navigation").Exist(3) Then
							Window("TcVizStructureManager").WinObject("3DImageViewer").Click 10,10,micLeftBtn
							Window("VizWindow").WinToolbar("3D Navigation").Click 3,4,micRightBtn
                            			Elseif Window("VizWindow").WinToolbar("Create Markup").Exist(1) Then
							Window("TcVizStructureManager").WinObject("3DImageViewer").Click 10,10,micLeftBtn
							Window("VizWindow").WinToolbar("Create Markup").Click 3,4,micRightBtn
						ElseIf sToolbarName="3D Markup" OR sToolbarName="3D GDT Markup" OR sToolbarName="3d Alignment" OR sToolbarName="3D CAE Viewing" OR sToolbarName="3D Clearance" OR sToolbarName="3D Comparison" OR sToolbarName="3D Coordinate System" OR sToolbarName="3D Measurement" Then
							Select Case sToolbarName
									Case "3D Markup"
										sToolbar = "Markup3D"
									Case "3D GDT Markup"
										sToolbar = "3D GDTMarkupUI"
									Case "3d Alignment"
										sToolbar = "VPAlignment"
									Case "3D CAE Viewing"
										sToolbar = "CAE Viewing"
									Case "3D Clearance"
										sToolbar = "ClearanceFramework"
									Case "3D Comparison"
										sToolbar = "VPCompare"
									Case "3D Coordinate System"
										sToolbar = "VPCoordinateSystem"
									Case "3D Measurement"
										sToolbar = "VPMeasurement"
							End Select
							objImageCanvas.Object.getViewerBean.setToolbarVisibility sToolbar, False
							Wait(5)
							bFlag = True
						ElseIf sToolbarName="3D Section" OR sToolbarName="3D Appearance" OR sToolbarName="3D Constraints" OR sToolbarName="3D Display Modes" OR sToolbarName="3D Movie Capture" OR sToolbarName="3D Navigation" OR sToolbarName="3D PMI" OR sToolbarName="3D Selection" OR sToolbarName="3D Standard Views" OR sToolbarName="3D Thrustline Editor" OR sToolbarName="3d Visibility" OR sToolbarName="3D Visual Report" OR sToolbarName="Create Markup" Then
							objImageCanvas.Object.getViewerBean.setToolbarVisibility sToolbarName, False
							Wait(5)
							bFlag = True							
						Else
							Window("VizWindow").InsightObject("InsightObject").Click 5,5,micRightBtn
						End If
						
						If bFlag = False Then
							Set WShell = CreateObject("WScript.Shell")
							WShell.SendKeys "{DOWN}"
							Call  Fn_SISW_Window_ContextMenu_Operation("Select", sToolbarName,"")
							Set WShell = Nothing
'							strMenu = JavaWindow("TcVizMainWin").WinMenu("ContextMenu").BuildMenuPath(sToolbarName)
'							JavaWindow("TcVizMainWin").WinMenu("ContextMenu").Select strMenu
						End If
					End If

					If Window("VizWindow").WinToolbar(sToolbarName).Exist(3) OR Window("VizWindow").WinObject(sToolbarName).Exist(3) Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Disable Toolbar [" + sToolbarName + "]")
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Successfully Disabled Toolbar [" + sToolbarName + "]")
						Fn_SISW_TcViz_ToolbarOperation = True
					End If

			Case "Exists"
					If Window("VizWindow").WinToolbar(sToolbarName).Exist(10) Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Successfully Found Existance of Toolbar [" + sToolbarName + "]")
						Fn_SISW_TcViz_ToolbarOperation = True
					ElseIf Window("VizWindow").WinObject(sToolbarName).Exist(10) Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Successfully Found Existance of Toolbar [" + sToolbarName + "]")
						Fn_SISW_TcViz_ToolbarOperation = True
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Find Toolbar [" + sToolbarName + "]")
					End If
					
	End Select
	Set objImageCanvas = Nothing
	Set AbstractInsightObj = Nothing
End Function

'--------------------------------------------------------------------------------------------------------------------
' Function Number   	: 3                                                                                
' Function Name     	: Fn_SISW_TcViz_ProductView_Operations
' Function Description  : Saves Session
' Function Usage    	: Result = Fn_SISW_TcViz_ProductView_Operations(ProductViewPopupSelect, Array("View1_4356", "Check-In/Out...:Check-Out"))
'									bReturn = Fn_SISW_TcViz_ProductView_Operations("Delete", Array("abc","", "","","Cancel"))
' NOTE					: Second argument is an Array, Format is {"ViewName", "Description", "PopupMenu", "NewPVButton", "GalleryButon", ...}
'
'                     		return True/False
'History:
'	Developer Name			Date			Rev. No.		Reviewer		Changes Done	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Vallari S			 4-Feb-2013			1.0				
'-----------------------------------------------------------------------------------------------------------------------------------
'	Reema W			 	20-Jun-2014			1.0								added Case "Delete"
'-----------------------------------------------------------------------------------------------------------------------------------
'	Reema W			 	20-Jun-2014			1.0								modified Case "ProductViewPopupSelect", Create
'-----------------------------------------------------------------------------------------------------------------------------------
'	Reema W			 	27-Aug-2014			1.0								added case "select" and added code to open  Product View Gallery dialog if it not exist
'-----------------------------------------------------------------------------------------------------------------------------------
'	Shweta Rathod		 	29-Sep-2014			1.0								Added Case "ProductViewPopupExist"
'-----------------------------------------------------------------------------------------------------------------------------------
'	Shweta Rathod		 	13-Oct-2014			1.0								Added Case "ProductViewTabPopupSelect"
'-----------------------------------------------------------------------------------------------------------------------------------
'	Shweta Rathod		 	13-Oct-2014			1.0								modified Case "ProductViewPopupSelect" to work with MPP perspective
'-----------------------------------------------------------------------------------------------------------------------------------
'	Shweta Rathod		 	15-Oct-2014			1.0								modified Case "Create"  added a code to handle InvaliAssemblyState to work with MPP perspective
'-----------------------------------------------------------------------------------------------------------------------------------
'	Priyanka kakade		 	13-Feb-2017			1.0								Added New Cases : "CreateWithoutWarning","CreateWithoutErrorDialog","CreateWithoutName","VerifyInvalidAssemblyStateError"
'-----------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_TcViz_ProductView_Operations(strAction, arrArg())
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_TcViz_ProductView_Operations"
	Dim bReturn
	Dim StrMenu, sViewName, aMenuList, intCount, sDesc
	Dim objProductViewGallery, objNewProductView
	Dim sPVButton, sGalleryButton
	Dim objDelete,sToggleName

	Set objProductViewGallery = Fn_SISW_TcViz_GetObject("ProductViewGallery")
	Set objNewProductView = Fn_SISW_TcViz_GetObject("NewProductView")
	
	Fn_SISW_TcViz_ProductView_Operations = False
	
	sViewName = arrArg(0)
	sDesc = arrArg(1)
	StrMenu = arrArg(2)
	sPVButton = arrArg(3)
	sGalleryButton = arrArg(4)
	
	If objProductViewGallery.Exist(3) = false Then
		If JavaWindow("TcVizMainWin").JavaWindow("JavaWinFrame").JavaDialog("ProductViewGallery").Exist(3) Then
			Set objProductViewGallery = JavaWindow("TcVizMainWin").JavaWindow("JavaWinFrame").JavaDialog("ProductViewGallery")
		Else
			bReturn = Fn_SISW_TcViz_ToolbarOperation("ButtonClick", "Create Markup", "Create 3D Product Views")
			If bReturn = False Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Dialog [ProductViewGallery] does not Exist")
				Exit Function
			End If
		End If
	End If
	
	Select Case strAction
		Case "Create","CreateWithoutWarning","CreateWithoutErrorDialog"
			objProductViewGallery.JavaButton("CreateProductView").Click micLeftBtn
			
			If strAction = "CreateWithoutWarning" Then
				If Window("VizWindow").Dialog("InvaliAssemblyState").Static("ActiveViewTogglesWarning").Exist(3) then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Dialog [InvaliAssemblyState] Exist")
					Exit Function
				End If
			ElseIf strAction = "CreateWithoutErrorDialog" Then
				If Window("VizWindow").Dialog("InvaliAssemblyState").WinEditor("ActiveViewTogglesError").Exist(3) then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Dialog [InvaliAssemblyState] Exist")
					Exit Function
				End If
			Else
				If Window("VizWindow").Dialog("InvaliAssemblyState").Exist(10) then
					 call Fn_UI_WinButton_Click("Fn_SISW_TcViz_ProductView_Operations",Window("VizWindow").Dialog("InvaliAssemblyState"),"Proceed",5,5,micLeftBtn)
				End If
			End If 	
			
			If objNewProductView.Exist(10) = false Then
				If JavaWindow("TcVizMainWin").JavaWindow("JavaWinFrame").JavaDialog("NewProductView").Exist(3) Then
					Set objNewProductView = JavaWindow("TcVizMainWin").JavaWindow("JavaWinFrame").JavaDialog("NewProductView")
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Dialog [NewProductView] does not Exist")
					Exit Function
				End If	 
			End If
			
			Err.Clear		
			If trim(sViewName) <> "" Then
				objNewProductView.JavaEdit("ProductViewName").Set sViewName
				Fn_SISW_TcViz_ProductView_Operations = True
			Else
				sViewName = objNewProductView.JavaEdit("ProductViewName").GetROProperty("value")
				Fn_SISW_TcViz_ProductView_Operations = sViewName
			End If
			'msgbox Err.Number
			If Err.Number < 0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Set Product View Name on Dialog [NewProductView]")
				Fn_SISW_TcViz_ProductView_Operations = False
				Exit Function
			End If
			
			If trim(sDesc) <> "" Then
				objNewProductView.JavaEdit("Description").Set sDesc
			End If
			
			If trim(sPVButton) = "" Then
				sPVButton = "OK"
			End If
			
			objNewProductView.JavaButton(sPVButton).Click micLeftBtn
			
			If Err.Number < 0 Then
'				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Click [OK] Button on Dialog [NewProductView]")
'				Fn_SISW_TcViz_ProductView_Operations = False
'				Exit Function
			Else
				'Synchronisation
				Call Fn_ReadyStatusSync(5)
			End If
			
			If trim(sGalleryButton) <> "" Then
				objProductViewGallery.JavaButton(sGalleryButton).Click micLeftBtn
				Call Fn_ReadyStatusSync(2)				
			End If
			
			If Err.Number < 0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Click [Cancel] Button on Dialog [ProductViewGallery]")
				Fn_SISW_TcViz_ProductView_Operations = False
				Exit Function
			End If
			
		Case "ViewExists"
		
			objProductViewGallery.JavaRadioButton("View").SetTOProperty "attached text", sViewName
			
			If objProductViewGallery.JavaRadioButton("View").Exist(5) Then
				Fn_SISW_TcViz_ProductView_Operations = True
			End If
			
			If trim(sGalleryButton) <> "" Then
				objProductViewGallery.JavaButton(sGalleryButton).Click micLeftBtn
				Call Fn_ReadyStatusSync(2)				
			End If
			
			If Err.Number < 0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Click [" + sGalleryButton + "] Button on Dialog [ProductViewGallery]")
				Fn_SISW_TcViz_ProductView_Operations = False
				Exit Function
			End If
			
		
		Case "ProductViewPopupSelect", "ProductViewTabPopupSelect"
	
			'Build the Popup menu to be selected
			aMenuList = split(StrMenu, ":",-1,1)
			intCount = Ubound(aMenuList)
			
			If strAction = "ProductViewTabPopupSelect" Then
			    objProductViewGallery.Activate	
				objProductViewGallery.JavaTab("ProductViewTab").Click 5, 5, "RIGHT"
			Else
				If trim(sViewName) <> "" Then
				objProductViewGallery.JavaRadioButton("View").SetTOProperty "attached text", sViewName			
				End If	
				objProductViewGallery.Activate	
				objProductViewGallery.JavaRadioButton("View").Click 20, 20, "RIGHT"
			End If
			
			'Select Menu action
			Select Case intCount
					Case "0"								
						objProductViewGallery.JavaMenu("label:="&aMenuList(0)&"","index:=0").Select
							bReturn = True
					Case "1"								
							objProductViewGallery.JavaMenu("label:="&aMenuList(0)&"","index:=0").JavaMenu("label:="&aMenuList(1)&"","index:=0").Select
							bReturn = True
			End Select
			If Err.Number < 0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Invoke [" + StrMenu + "] Menu on Dialog [ProductViewGallery]")
				Fn_SISW_TcViz_ProductView_Operations = False
				Exit Function
			End If
			If trim(sGalleryButton) <> "" Then
				objProductViewGallery.JavaButton(sGalleryButton).Click micLeftBtn
				wait 2		
			End If
			
			If Err.Number < 0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Click [" + sGalleryButton + "] Button on Dialog [ProductViewGallery]")
				Fn_SISW_TcViz_ProductView_Operations = False
				Exit Function
			End If
			Fn_SISW_TcViz_ProductView_Operations = bReturn
		
		Case "Delete"
			Set objDelete = JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Delete")
			If trim(sViewName) <> "" Then
				objProductViewGallery.JavaRadioButton("View").SetTOProperty "attached text", sViewName			
			End If
			
			objProductViewGallery.JavaRadioButton("View").Click 20, 20, "LEFT"

			objProductViewGallery.JavaButton("DeleteProductView").Click micLeftBtn

			If 	objDelete.Exist(10) = false Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Delete JavaDialog does not Exist")
				Fn_SISW_TcViz_ProductView_Operations = False
				Exit Function
			End If

			objDelete.JavaButton("Yes").Click micLeftBtn
			If Err.Number < 0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Click Yes on Delete JavaDialog ")
				Fn_SISW_TcViz_ProductView_Operations = False
				Set objDelete = Nothing
				Exit Function
			End If

			If trim(sGalleryButton) <> "" Then
				objProductViewGallery.JavaButton(sGalleryButton).Click micLeftBtn
				wait 2		
			End If
			
			If Err.Number < 0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Click [" + sGalleryButton + "] Button on Dialog [ProductViewGallery]")
				Fn_SISW_TcViz_ProductView_Operations = False
				Set objDelete = Nothing
				Exit Function
			Else
				Fn_SISW_TcViz_ProductView_Operations = True
			End If
		
		Case "ProductViewPopupExist"					
					aMenuList = split(StrMenu, ":",-1,1)
					intCount = Ubound(aMenuList)
					
					If trim(sViewName) <> "" Then
						objProductViewGallery.JavaRadioButton("View").SetTOProperty "attached text", sViewName			
					End If
					objProductViewGallery.Activate
					objProductViewGallery.JavaRadioButton("View").Click 20, 20, "RIGHT"				
					wait 3
					Select Case intCount
						Case "0"
							 bReturn = objProductViewGallery.JavaMenu("label:="&aMenuList(0)&"","index:=0").exist(5)
						Case "1"								
							bReturn = objProductViewGallery.JavaMenu("label:="&aMenuList(0)&"","index:=0").JavaMenu("label:="&aMenuList(1)&"","index:=0").exist(5)
                                 Case Else
							Fn_SISW_TcViz_ProductView_Operations = False
                       			 Exit Function
					End Select
					Call Fn_KeyBoardOperation("SendKeys", "{ESC}")
					Fn_SISW_TcViz_ProductView_Operations = bReturn	
					
		Case "Select"
			If trim(sViewName) <> "" Then
				objProductViewGallery.JavaRadioButton("View").SetTOProperty "attached text", sViewName			
			End If
			
			objProductViewGallery.JavaRadioButton("View").Click 20, 20, "LEFT"
			If Err.Number < 0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select Product View [" + sViewName + "] ")
				Fn_SISW_TcViz_ProductView_Operations = False
				Set objDelete = Nothing
				Exit Function
		     Else
				Fn_SISW_TcViz_ProductView_Operations = True
			End If	
			
		Case "CreateWithoutName"
			objProductViewGallery.JavaButton("CreateProductView").Click micLeftBtn
			
			If Window("VizWindow").Dialog("InvaliAssemblyState").Exist(10) then
				 call Fn_UI_WinButton_Click("Fn_SISW_TcViz_ProductView_Operations",Window("VizWindow").Dialog("InvaliAssemblyState"),"Proceed",5,5,micLeftBtn)
			End If
			
			If objNewProductView.Exist(5) OR JavaWindow("TcVizMainWin").JavaWindow("JavaWinFrame").JavaDialog("NewProductView").Exist(1) Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Dialog [NewProductView] Exist")
				Exit Function
			Else
				sViewName = objProductViewGallery.JavaRadioButton("View").GetROProperty("attached text")
				Fn_SISW_TcViz_ProductView_Operations = sViewName
			End If
			
			If trim(sGalleryButton) <> "" Then
				objProductViewGallery.JavaButton(sGalleryButton).Click micLeftBtn
				Call Fn_ReadyStatusSync(2)				
			End If
			
			If Err.Number < 0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Click [Cancel] Button on Dialog [ProductViewGallery]")
				Fn_SISW_TcViz_ProductView_Operations = False
				Exit Function
			End If
			
		Case "VerifyInvalidAssemblyStateError"
			objProductViewGallery.JavaButton("CreateProductView").Click micLeftBtn
		
			If Window("VizWindow").Dialog("InvaliAssemblyState").WinEditor("ActiveViewTogglesError").Exist(5) = False then
				 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Dialog [InvaliAssemblyState] does not Exist")
				Exit Function
			End If
			
			Window("VizWindow").Dialog("InvaliAssemblyState").Activate
			sToggleName = Window("VizWindow").Dialog("InvaliAssemblyState").WinEditor("ActiveViewTogglesError").GetROProperty("text")
			If instr(sToggleName,StrMenu)>0 Then
				Fn_SISW_TcViz_ProductView_Operations = True
			End If 
			
			Window("VizWindow").Dialog("InvaliAssemblyState").Close
			wait(1)
			If Fn_SISW_UI_Object_Operations("Fn_SISW_TcViz_ProductView_Operations","Exist", Window("VizWindow").Dialog("InvaliAssemblyState"),SISW_MIN_TIMEOUT) Then
				Window("VizWindow").Dialog("InvaliAssemblyState").Close
			End If
			If Err.Number < 0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Click [OK] Button on Dialog [InvaliAssemblyState]")
				Fn_SISW_TcViz_ProductView_Operations = False
				Exit Function
			End If
			
			If trim(sGalleryButton) <> "" Then
				objProductViewGallery.JavaButton(sGalleryButton).Click micLeftBtn
				Call Fn_ReadyStatusSync(1)				
			End If
			
			If Err.Number < 0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Click [Cancel] Button on Dialog [ProductViewGallery]")
				Fn_SISW_TcViz_ProductView_Operations = False
				Exit Function
			End If
	End Select
		Set objProductViewGallery = Nothing
		Set objNewProductView = Nothing
		Set objDelete = Nothing
End Function

'--------------------------------------------------------------------------------------------------------------------
' Function Number   	: 4                                                                                
' Function Name     	: Fn_SISW_TcViz_Options_Setting
' Function Description  : Set Visualization Options
' Function Usage    	: Result = Fn_SISW_TcViz_Options_Setting("OFF", "ON")
'                     		return True/False
'History:
'	Developer Name			Date			Rev. No.		Reviewer		Changes Done	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Vallari S			 11-Jun-2013			1.0				
'-----------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_TcViz_Options_Setting(chkViewNX, chkOpenInViz)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_TcViz_Options_Setting"
	Dim bReturn
	Dim objWin
	
	Fn_SISW_TcViz_Options_Setting = False
	
	Set objWin = Fn_SISW_TcViz_GetObject("Options")
	
	If Not objWin.Exist(3) Then
		bReturn = Fn_MenuOperation("WinMenuSelect", "Edit:Options...")
		If bReturn = False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Menu Operation [Edit:Options...] Failed")
			Set objWin = Nothing
			Exit Function
		End If
	End If
	
	Err.clear
	
	If objWin.Exist(10) Then
		objWin.JavaStaticText("BottomLink").SetTOProperty "label", "Options"
		objWin.JavaStaticText("BottomLink").Click 5, 5, "LEFT"
		If Err.Number < 0 Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Select [Options] Bottom Link")
			Set objWin = Nothing
			Exit Function			
		End If
		
		objWin.JavaTree("OptionsTree").Select "Options:Visualization:Lifecycle Visualization"
		If Err.Number < 0 Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Select [Visualization:Lifecycle Visualization] Tree Node")
			Set objWin = Nothing
			Exit Function			
		End If
		
		objWin.JavaCheckBox("ViewNX").Set chkViewNX
		If Err.Number < 0 Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Set [ViewNX] Checkbox")
			Set objWin = Nothing
			Exit Function			
		End If
		
		objWin.JavaCheckBox("OpenInViz").Set chkOpenInViz
		If Err.Number < 0 Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Set [OprnInViz] Checkbox")
			Set objWin = Nothing
			Exit Function			
		End If
		
		objWin.JavaButton("OK").Click micLeftBtn
		If Err.Number < 0 Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Click [OK] Button")
			Set objWin = Nothing
			Exit Function			
		End If
	Else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Dialog [Options] does not Exist")
		Set objWin = Nothing
		Exit Function
	End If
	
	Fn_SISW_TcViz_Options_Setting = True
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Set [Viz-Options]")
	
End Function



'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_SISW_GraphicsTab_3DViewOperation
'
'Description			 :	Function Used to perform operations on Change History table of Summary Tab

'Parameters			   : 1.  sAction: valid Action Name 
'										2.   dicGraphicsViewInfo
'
'Return Value		   : 	True or False

'Pre-requisite			:	Graphics tab should be activated

'Examples				:  		Set dicGraphicsViewInfo = CreateObject( "Scripting.Dictionary" )
''
''													dicGraphicsViewInfo("RotateBy") = 90				''""NOTE:- 	value in Degrees by which image is rotatted
''											msgbox Fn_SISW_GraphicsTab_3DViewOperation("Rotate", dicGraphicsViewInfo)
''										
''												dicGraphicsViewInfo("ZoomBy") = 1.2					''""NOTE:- 	vvalue to zoom in and Zoom out -for Zoom in give input as1.x,	for Zoom OUT give input as 0.x WHERE X IS FROM 0 TO 9
''										msgbox Fn_SISW_GraphicsTab_3DViewOperation("Zoom", dicGraphicsViewInfo)
'History					 :			
'				Developer Name						Date					Rev. No.				Changes Done								 Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'				Reema W						18-Jun-2014				1.0						created new function					
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_SISW_GraphicsTab_3DViewOperation(sAction, dicGraphicsViewInfo)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_GraphicsTab_3DViewOperation"
   Dim objGraphics, obj3DBean, objDeviceReplay, sValue
   Set objJTViewer = JavaWindow("TcVizMainWin").JavaWindow("JApplet").JavaObject("JTViewer")
	Fn_SISW_GraphicsTab_3DViewOperation = False
	Select Case sAction
		Case "Rotate"
			Set obj3DBean = objJTViewer.Object.getViewer3DBean
			obj3DBean.Rotate dicGraphicsViewInfo("RotateBy")
			If Err.Number < 0  Then
					Fn_SISW_GraphicsTab_3DViewOperation = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Rotate Image.")
			Else
					Fn_SISW_GraphicsTab_3DViewOperation = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully Rotated Image.")
			End If
			
		Case "getFlyToState"
			sValue =  objJTViewer.Object.getFlyToState()
			If Err.Number < 0  Then
					Fn_SISW_GraphicsTab_3DViewOperation = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to getFlyToState value.")
					Exit Function
			End If			
			
			If lCase(sValue) = "true" Then
				Fn_SISW_GraphicsTab_3DViewOperation = True
			ElseIf lCase(sValue) = "false" Then
				Fn_SISW_GraphicsTab_3DViewOperation = False	
			End If
			
		Case "Zoom"
			Set obj3DBean = objJTViewer.Object.getViewer3DBean
			obj3DBean.Zoom dicGraphicsViewInfo("ZoomBy")
				If Err.Number < 0  Then
						Fn_SISW_GraphicsTab_3DViewOperation = False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Zoom Image.")
				Else
						Fn_SISW_GraphicsTab_3DViewOperation = True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully Zoomed Image.")
				End If

		Case "IsWireFrameState"
			Set objImageCanvas = JavaWindow("TcVizMainWin").JavaWindow("JApplet").JavaObject("ImageCanvas")
			bFlag = objImageCanvas.Object.getViewerBean().getDrawWireframeState()
			If Cbool(bFlag) = True Then
				Fn_SISW_GraphicsTab_3DViewOperation = True
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully verified [" & sAction & " ] is [True].")
			Else
				Fn_SISW_GraphicsTab_3DViewOperation = False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Verify [" &sAction & " ] is [True].")					
			End If
				
		Case "PopupMenuSelect"
			Window("TcVizStructureManager").WinObject("3DImageViewer").Click 169,138,micRightBtn
			wait 1
			Set WShell = CreateObject("WScript.Shell")		
			WShell.SendKeys "{DOWN}"
			bFlag = Fn_SISW_Window_ContextMenu_Operation("Select",dicGraphicsViewInfo("PopupMenu"),"")	
			If bFlag = True Then
				Fn_SISW_GraphicsTab_3DViewOperation = True
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully performed Action [" & sAction & " ].")
			Else
				Fn_SISW_GraphicsTab_3DViewOperation = False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to perform Action [" &sAction & " ].")
			End If			
			Set WShell = Nothing
			
		Case Else
					Fn_SISW_GraphicsTab_3DViewOperation = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to execute Function [ Fn_SISW_GraphicsTab_3DViewOperation ] Invalid case [ " & sAction & " ] ")
					Exit function
	End Select
End Function


'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_SISW_TcViz_CheckOutProductView
'
'Description			 :	Function Used to CheckOut a Product View From ProductView Gallery Dialog Using RMB.

'Parameters			   : 	1.  sProductView: 	ProductView name to be check-out
'							2.  sOption:		CheckOut RadioButton to be select on Check-Out Warning Dialog
'							3.  sAction:		Case"CheckOut"for internal use of function.
'							4.	sChangeID:		JavaEdit'ChangeID' field on Check-out Dialog
'							5.	sComment:		JavaEdit 'Comments' field on Check-out Dialog
'							6.	sFutureUse:		For Future Use
'Return Value		   : 	True or False

'Pre-requisite			:	ProductView Gallery Dialog Should Be Opened

'Examples				:  		 Fn_SISW_TcViz_CheckOutProductView("PV_17536","CheckOutAndOverwrite", "CheckOut" , "123", "comment", "")
'History					 :			
'				Developer Name						Date					Rev. No.				Changes Done								 Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'				Reema W						26-Aug-2014				1.0						created new function					
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'				Reema W						27-Aug-2014				1.0						 modified condition to check existence of warning dialog			
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Public Function Fn_SISW_TcViz_CheckOutProductView(sProductView, sOption, sAction , sChangeID,  sComment, sFutureUse )
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_TcViz_CheckOutProductView"
	Dim bReturn
	Dim objCheckOutWarning,objCheckOut
	
	Fn_SISW_TcViz_CheckOutProductView = False
	
	Set objCheckOutWarning = Fn_SISW_TcViz_GetObject("Check-OutWarning")
	Set objCheckOut= Fn_SISW_GetObject("Check-Out")
	
	''  modified condition to check existence of warning dialog
	If Not objCheckOutWarning.Exist(3) or not objCheckOut.Exist(3) Then
		bReturn = Fn_SISW_TcViz_ProductView_Operations("ProductViewPopupSelect", Array(sProductView,"", "Check-In/Out...:Check-Out","",""))
		If bReturn = False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed To RMB 'Check-In/Out...:Check-Out' on ["+sProductView+"]"  )
			Set objCheckOutWarning = Nothing
			Set objCheckOut= Nothing
			Exit Function
		End If
	End If
	
	Err.clear
	 ''	modified condition to check existence of warning dialog
	If sOption <>"" and objCheckOutWarning.Exist Then
		Select Case sOption
			
			Case "CreateNewSnapshot"
				objCheckOutWarning.JavaRadioButton("SnapshotOption").SetTOProperty "attached text", "Create a new Snapshot to save your current work and preserve the last saved Snapshot."
			Case "DiscardCurrentWorkAndCheckOut"
				objCheckOutWarning.JavaRadioButton("SnapshotOption").SetTOProperty "attached text", "Discard your current work and checkout and apply the last saved Snapshot."
			Case "CheckOutAndOverwrite"
				objCheckOutWarning.JavaRadioButton("SnapshotOption").SetTOProperty "attached text", "Continue with check-out and overwrite the last saved Snapshot."
			Case Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "case for option ["+sOption+"] does not exists")
				Set objCheckOutWarning = Nothing
				Set objCheckOut= Nothing
				Exit Function
		End Select  
		
		bReturn = Fn_SISW_UI_JavaRadioButton_Operations("objCheckOutWarning", "Set", objCheckOutWarning, "SnapshotOption", "ON")	
		If bReturn = False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed To set JavaRadioButton(SnapshotOption) to ON "  )
			Set objCheckOutWarning = Nothing
			Set objCheckOut= Nothing
			Exit Function
		End If
		
		objCheckOutWarning.JavaButton("OK").Click micLeftBtn
		If Err.Number < 0 Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Click [OK] Button")
			Set objCheckOutWarning = Nothing
			Set objCheckOut= Nothing
			Exit Function	
		End If
		
	End If
	
	If sAction <>"" Then
		Select Case sAction
			
			Case "CheckOut"
			
				If sChangeID <> "" Then
					Call Fn_SISW_UI_JavaEdit_Operations("Fn_SISW_ChkInOut_ObjectCheckOut", "Set",  objCheckOut, "ChangeID", sChangeID)
				End If
				If sComment <> "" Then
					Call Fn_SISW_UI_JavaEdit_Operations("Fn_SISW_ChkInOut_ObjectCheckOut", "Set",  objCheckOut, "Comments", sComment)
				End If
				Call Fn_SISW_UI_JavaButton_Operations("Fn_SISW_ChkInOut_ObjectCheckOut", "Click", objCheckOut,"Yes")
				If Err.Number < 0 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Click [OK] Button")
					Set objCheckOutWarning = Nothing
					Set objCheckOut= Nothing
					Exit Function
				End If
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Values entered in the Check Out Dialog Successfully")
			Case Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "case for option ["+sAction+"] does not exists")
				Set objCheckOutWarning = Nothing
				Set objCheckOut= Nothing
				Exit Function
		End Select
	End If
	
	Fn_SISW_TcViz_CheckOutProductView = True
	Set objCheckOutWarning = Nothing
	Set objCheckOut= Nothing
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Set Checkout ProductView")
	
End Function


'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_SISW_TcViz_CheckInProductView
'
'Description			 :	Function Used to Check-In a Product View From ProductView Gallery Dialog Using RMB.

'Parameters			   : 	1.  sAction: 	Case 'Check-In' for Internal Use
'							2.  sProductView:		ProductView Name to be check-In
'							3.  aExploreOption:		For Future Use of Explore options
'							4.	sErrorMessage:		For Future Use to verify Error
'Return Value		   : 	True or False

'Pre-requisite			:	ProductView Gallery Dialog Should Be Opened

'Examples				:  		Fn_SISW_TcViz_CheckInProductView("Check-In", "PV_17536" , "", "")
'History					 :			
'				Developer Name						Date					Rev. No.				Changes Done								 Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'				Reema W						26-Aug-2014				1.0						created new function					
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_SISW_TcViz_CheckInProductView(sAction, sProductView, aExploreOption, sErrorMessage)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_TcViz_CheckInProductView"
	Dim bReturn
	Dim objCheckIn
	Fn_SISW_TcViz_CheckInProductView = False

		Select Case sAction
			Case "Check-In"

				Set objCheckIn=Fn_SISW_GetObject("Check-In")
				If Not objCheckIn.Exist(3) Then
					bReturn = Fn_SISW_TcViz_ProductView_Operations("ProductViewPopupSelect", Array(sProductView,"", "Check-In/Out...:Check-In","",""))
					If bReturn = False Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed To RMB 'Check-In/Out...:Check-In' on ["+sProductView+"]"  )
						Set objCheckIn = Nothing
						Exit Function
					End If
				End If
				bReturn = Fn_Button_Click("Fn_SISW_ChkInOut_ObjectCheckIn", objCheckIn, "Yes")
				If bReturn = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Click [Yes] Button")
					Set objCheckIn = Nothing
					Exit Function
				End If
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Check-In done successfully")
			Case Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "case for option ["+sAction+"] does not exists")
				Set objCheckIn = Nothing
				Exit Function
        End Select
            
		Fn_SISW_TcViz_CheckInProductView =TRUE
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Check-In done successfully")

		Set objCheckIn=Nothing
End Function


''*********************************************************		Function to Perform operation on Preference Dialog in Structure Manager	***********************************************************************
'
''Function Name		:		Fn_SISW_TcViz_PreferenceOperation
'
''Description			 :		 Function to Perform operation on Preference Dialog in Structure Manager
'
''Parameters		:		1.	sAction = Action To Perform
'							   			2.  sTabName = pass the value of tab you want to select .  eg "Display"
'										3. sButton = Name of button to click. 		eg. "OK" or "cancel" 		
'										4. dicPrefInfo = Dictionary Object
'										5. sReserve = For Future Use  		NOTE : For Defining Objects checkbox and Complying Objects checkbox use 'sReserve'		
'										6. sCheck = "ON/OFF" To check and uncheck checkboxes.

'			  										
''Return Value		   : 				True or False 
'
''Pre-requisite			:				Preference Dialog should be displayed in SM.
'
'
''Examples				:		 Dim dicPrefInfo
'											Set dicPrefInfo=CreateObject("Scripting.Dictionary")

'											Case "Display" 'Case handled according to Tab of the Dialog

												'dicPrefInfo("Model")="Tessellation" 
												'Fn_SISW_TcViz_PreferenceOperation("Set","Display",dicPrefInfo, "OK/Cancel", "","ON")	 
'
'								Set dicPrefInfo = CreateObject("Scripting.Dictionary")
'									dicPrefInfo("WinCheckBox1") = "View interpolation:ON"
'									dicPrefInfo("ModelSetasdefault") = "ON"
'									dicPrefInfo("WinButton1") = "Apply"
'								bReturn = Fn_SISW_TcViz_PreferenceOperation("Set","General",dicPrefInfo,"", "","")
'
'
'								Set dicPrefInfo = CreateObject("Scripting.Dictionary")
'									dicPrefInfo("WinCheckBox1") = "Show WCS reference grid in Orthographic, axis-aligned views:ON"
'									dicPrefInfo("GridSetasdefault") = "ON"
'									dicPrefInfo("WinButton1") = "OK"
'								bReturn = Fn_SISW_TcViz_PreferenceOperation("Set","Grid",dicViewPreferences,"OK","","")	
'
'History:
'										Developer Name				Date						Rev. No.							Changes Done						Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Ankit Tewari					11-Sep-2014					1.0		
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Poonam Chopade					20-Feb-2018					1.1						Added Cases "General" , "Grid"			TC11.5_2018012200_NewDevelopment_PoonamC							
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_TcViz_PreferenceOperation(sAction, sTabName , dicPrefInfo , sButton, sReserve,sCheck)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_TcViz_PreferenceOperation"
	Dim DictItems,DictKeys
	Dim iCounter,sValue,ObjDialog
	Dim sSubAction,sProperty,bFlag
	
	Fn_SISW_TcViz_PreferenceOperation=False
	Set ObjDialog = JavaWindow("TcVizMainWin").Dialog("Preferences")
	
	If ObjDialog.Exist(2) <> True Then
		ObjDialog=nothing
		Exit Function
	Else
		ObjDialog.WinTab("PrefTab").Select sTabName
	End If

   	Select Case sTabName
		Case "General" 'TC11.5(20180122.00)_NewDevelopment_PoonamC_19Feb2018
			Select Case sAction
				Case "Set"
					DictItems = dicPrefInfo.Items
					DictKeys = dicPrefInfo.Keys	
					For iCounter = 0 to Ubound(Dictkeys)
						If Instr(DictKeys(iCounter),"WinCheckBox")>0 Then
							sSubAction = "WinCheckBox"
						ElseIf Instr(DictKeys(iCounter),"WinButton")>0 Then
							sSubAction = "WinButton"
						Else
							sSubAction = DictKeys(iCounter)
						End If
						
						sProperty = DictItems(iCounter)
						bFlag = False
						
						Select Case sSubAction
							Case "WinButton"
								If sProperty<>"" Then
									'Click on button provided
									ObjDialog.WinButton(sProperty).Click 5,5,micLeftBtn
									If Err.Number < 0 Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to click ["+sProperty+"] button on [ Preferences ] window.")
										Set ObjDialog = Nothing
										Exit Function
									End If
									Wait 1
									bFlag = True
								End If
							Case "WinCheckBox"	
								If sProperty<>"" Then 
									sProperty = Split(sProperty,":")
									ObjDialog.WinCheckBox("CheckBox").SetTOProperty "text",sProperty(0)
									ObjDialog.WinCheckBox("CheckBox").Set sProperty(1)
									If Err.Number < 0 Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to set checkbox ["+sProperty(0)+"] as [ "+sProperty(1)+" ]in [ Preferences ] window.")
										Set ObjDialog = Nothing
										Exit Function
									End If
									Wait 1
									bFlag = True 
								End If
							Case "ModelSetasdefault"
								If sProperty<>"" Then 
									ObjDialog.WinCheckBox("ModelSetasdefault").Set sProperty
									If Err.Number < 0 Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to set checkbox [ModelSetasdefault] as [ "+sProperty+" ]in [ Preferences ] window.")
										Set ObjDialog = Nothing
										Exit Function
									End If
									Wait 1
									bFlag = True 
								End If
						End Select
						If bFlag = False Then
							Fn_SISW_TcViz_PreferenceOperation = False
							Set ObjDialog = Nothing
							Exit Function
						Else
							Fn_SISW_TcViz_PreferenceOperation = True
						End If	
					Next
			End Select	
	'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		Case "Display"
			If sAction="Set" Then
				DictItems = dicPrefInfo.Items
				DictKeys = dicPrefInfo.Keys	
				For iCounter = 0 to Ubound(Dictkeys)
					If DictKeys(iCounter) = "Model" Then
						If ObjDialog.WinButton(DictItems(iCounter)).exist(1) Then
							ObjDialog.WinButton(DictItems(iCounter)).Click
						Else
							ObjDialog.WinObject(DictItems(iCounter)).Click
						End If
						bFlag=True
					End If
				Next
				
				If sCheck = "ON" Then
					ObjDialog.WinCheckBox("Setasdefault").Set "ON"
				End If
				
				ObjDialog.WinButton("Apply").Click			
			End If
			
			If bFlag = True Then
				Fn_SISW_TcViz_PreferenceOperation=True
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_TcViz_PreferenceOperation ] Successfully performed [ " & sAction & " ] Action on [ "+sTabName+" ].")
			Else
				Fn_SISW_TcViz_PreferenceOperation=False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_TcViz_PreferenceOperation ] Failed to performed [ " & sAction & " ] Action on [ "+sTabName+" ].")
				ObjDialog=nothing
				Exit Function
			End If
	'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		Case "Grid" 'TC11.5(20180122.00)_NewDevelopment_PoonamC_19Feb2018
			Select Case sAction
				Case "Set"
					DictItems = dicPrefInfo.Items
					DictKeys = dicPrefInfo.Keys	
					For iCounter = 0 to Ubound(Dictkeys)
						If Instr(DictKeys(iCounter),"WinCheckBox")>0 Then
							sSubAction = "WinCheckBox"
						ElseIf Instr(DictKeys(iCounter),"WinButton")>0 Then
							sSubAction = "WinButton"
						Else
							sSubAction = DictKeys(iCounter)
						End If
						
						sProperty = DictItems(iCounter)
						bFlag = False
						
						Select Case sSubAction
							Case "WinButton"
								If sProperty<>"" Then
									'Click on button provided
									ObjDialog.WinButton(sProperty).Click 5,5,micLeftBtn
									If Err.Number < 0 Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to click ["+sProperty+"] button on [ Preferences ] window.")
										Set ObjDialog = Nothing
										Exit Function
									End If
									Wait 1
									bFlag = True
								End If
							Case "WinCheckBox"	
								If sProperty<>"" Then 
									sProperty = Split(sProperty,":")
									ObjDialog.WinCheckBox("CheckBox").SetTOProperty "text",sProperty(0)
									ObjDialog.WinCheckBox("CheckBox").Set sProperty(1)
									If Err.Number < 0 Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to set checkbox ["+sProperty(0)+"] as [ "+sProperty(1)+" ]in [ Preferences ] window.")
										Set ObjDialog = Nothing
										Exit Function
									End If
									Wait 1
									bFlag = True 
								End If
							Case "GridSetasdefault"	
								If sProperty<>"" Then 
									ObjDialog.WinCheckBox("Setasdefault").Set sProperty
									If Err.Number < 0 Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to set checkbox [GridSetasdefault] as [ "+sProperty+" ]in [ Preferences ] window.")
										Set ObjDialog = Nothing
										Exit Function
									End If
									Wait 1
									bFlag = True 
							   End If
						End Select
						If bFlag = False Then
							Fn_SISW_TcViz_PreferenceOperation = False
							Set ObjDialog = Nothing
							Exit Function
						Else
							Fn_SISW_TcViz_PreferenceOperation = True
						End If	
					Next
				End Select	
	'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		Case "Selection"
		'Do Nothing
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		Case else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_TcViz_PreferenceOperation ] Invalid case [ " & sAction & " ].")
			Exit function
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	End Select
	
	If sButton <> "" Then
		If ObjDialog.WinButton(sButton).Exist(2)=True Then
			ObjDialog.WinButton(sButton).Click 5,5,micLeftBtn	
		else
			Set ObjDialog = nothing
			Exit function
		End if
	End If
	
	If Fn_SISW_TcViz_PreferenceOperation<> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [ Fn_SISW_TcViz_PreferenceOperation ] executed successfuly with case [ " & sAction & " ].")
	else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_TcViz_PreferenceOperation ] Failed to executed with case [ " & sAction & " ].")
	End If

	Set ObjDialog = nothing
End Function 


''*********************************************************		Function to Perform operation on Image in Viewer Tab	***********************************************************************
'
''Function Name		:		Fn_SISW_TcViz_ImageInViewerOperations
'
''Description			 :		 Function to Perform operation on on Image in Viewer Tab
'
''Parameters		:		1.	sAction = Action To Perform
'							2.  sImg = Name of the Image Tab.
'							3. sReserve = Reserved for future use	

'			  										
''Return Value		   : 				True or False 
'
''Pre-requisite			:				Viewer Tab should be Opened in MyTC.
'
''Examples				:		 
										'Fn_SISW_TcViz_PreferenceOperation("Verify","ds_123","")	 
'										'Fn_SISW_TcViz_PreferenceOperation("Verify","ds_123~ds_456","")	 
'History:
'										Developer Name				Date						Rev. No.							Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Rinki A					1-Oct-2014					1.0									
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function Fn_SISW_TcViz_ImageInViewerOperations(sAction,sImg,sReserve)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_TcViz_ImageInViewerOperations"
Dim objDesc,obj,arrImg
Select Case sAction
		Case "Verify"
						arrImg=Split(sImg,"~",-1,1)
						Set objDesc=Description.Create()
						objDesc("toolkit class").value="org.eclipse.swt.widgets.Canvas"
						objDesc("tagname").value="Canvas"
						set obj=JavaWindow("TcVizMainWin").ChildObjects(objDesc)
						For iCnt = 1 To obj.Count - 1 
							If trim(arrImg(iCnt-1)) = trim(obj(iCnt).GetROProperty("text")) Then
								Fn_SISW_TcViz_ImageInViewerOperations=True
							Else
								Fn_SISW_TcViz_ImageInViewerOperations = False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to execute Function [ Fn_SISW_TcViz_CheckShuttleInView ] with case [ " & sAction & " ] ")
								Exit function
							End If 
						Next
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully execute Function [ Fn_SISW_TcViz_CheckShuttleInView ]with case [ " & sAction & " ] ")
								
		Case Else
						Fn_SISW_TcViz_ImageInViewerOperations = False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to execute Function [ Fn_SISW_GraphicsTab_3DViewOperation ] Invalid case [ " & sAction & " ] ")
						Exit function
		End Select
		
Set objDesc = nothing
Set obj = nothing
End Function

''*********************************************************		Function to Perform operation on Product View in Option Dialog 	***********************************************************************
''Function Name		:		Fn_SISW_TcViz_Options_ProductView
'
''Description			 :		 Function to Perform operation on Product View in Option Dialog
'
''Parameters		:		      1.	sAction = Action To Perform
'							2.  sNode = Name of node in Option tree
'							3. dicInfo = dictionary object
'							4. sButton = Name of the button
'							5. sReserve = Reserved for future use	
		  										
''Return Value		   : 		True or False 
'
''Pre-requisite			:	Options Dialog must be opened.
'
''Examples				:	Fn_SISW_TcViz_Options_ProductView("Verify","Options:Visualization:Product View","dicInfo","OK","")	 
						
'										Set dicInfo=CreateObject("Scripting.Dictionary")
'										dicInfo("Show Unconfigured Variants") ="0"
'										dicInfo("Show Unconfigured Changes") ="0"
'										dicInfo("Show Suppressed Occurrences") ="0"

'							'Fn_SISW_TcViz_Options_ProductView("IsEnabled","Options:Visualization:Product View",dicInfo, "OK","")
'History:
'										Developer Name				Date						Rev. No.							Changes Done				Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Ankit Tewari			11-Sept-2014					1.0									
'										Priyanka kakade  		23-Feb-2017						1.1 					Added New Case : "CheckBoxSelect"    [TC1123(20161205c00)_PoonamC_NewDevelopment]																															
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_TcViz_Options_ProductView(sAction,sNode,dicInfo,sButton,sReserve)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_TcViz_Options_ProductView"
	Dim bReturn
	Dim objWin,dicItems,dicKeys
	Fn_SISW_TcViz_Options_ProductView = False
	iCount = 0
	Set objWin = Fn_SISW_TcViz_GetObject("Options")
	
	If Not objWin.Exist(2) Then
		bReturn = Fn_MenuOperation("WinMenuSelect", "Edit:Options...")
		If bReturn = False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Menu Operation [Edit:Options...] Failed")
			Set objWin = Nothing
			Exit Function
		End If
	End If
	
	Err.clear
	
	If objWin.Exist(2) Then
		objWin.JavaStaticText("BottomLink").SetTOProperty "label", "Options"
		objWin.JavaStaticText("BottomLink").Click 5, 5, "LEFT"
		If Err.Number < 0 Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Select [Options] Bottom Link")
			Set objWin = Nothing
			Exit Function			
		End If
		
		objWin.JavaTree("OptionsTree").Select sNode
		If Err.Number < 0 Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Select ["+sNode+"] Tree Node")
			Set objWin = Nothing
			Exit Function			
		End If
	Else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Dialog [Options] does not Exist")
		Set objWin = Nothing
		Exit Function
	End If	
	Select Case sAction	
		Case "Verify"
			DictKeys = dicInfo.Keys
			DictItems = dicInfo.Items			
				For iCounter = 0 to Ubound(DictKeys)							
					Select Case DictKeys(iCounter)
						Case "ViewToggleWarningLevel","GeometryAsset"
							If DictItems(iCounter) <> "" Then					
								If objWin.JavaList(DictKeys(iCounter)).Exist(5) Then
									sValue = objWin.JavaList(DictKeys(iCounter)).GetROProperty("value")
									If sValue =  DictItems(iCounter) Then
										iCount=iCount+1
									End If
								End If
							End If
																					
							If iCount = Ubound(DictKeys)+1 Then
								Fn_SISW_TcViz_Options_ProductView=True
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_TcViz_Options_ProductView ] Successfully performed [ " & sAction & " ] Action.")	
							End If
							
						Case Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_TcViz_Options_ProductView ] Invalid case [ " & sAction & " ].")
							Exit function
					End Select
				Next
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------							
				Case "IsChecked"
					DictKeys = dicInfo.Keys
					DictItems = dicInfo.Items
					For iCounter = 0 to Ubound(DictKeys)							
						Select Case DictKeys(iCounter)
							Case "Show Unconfigured Variants","Show Unconfigured Changes","Show Suppressed Occurrences","Show Unconfigured Assigned Occurrences","Show Unconfigured By Occurrence Effectivity","Show GCS Connection Points"
								If DictItems(iCounter) <> "" Then	
									objWin.JavaCheckBox("ViewTogglestoConsider").SetTOProperty "attached text",DictKeys(iCounter) 
									If objWin.JavaCheckBox("ViewTogglestoConsider").Exist(1) Then								
										sValue = objWin.JavaCheckBox("ViewTogglestoConsider").GetROProperty("value")
										If sValue =  DictItems(iCounter) Then
											iCount=iCount+1
										End If
									End If
								End If								
								
								If iCount = Ubound(DictKeys)+1 Then
									Fn_SISW_TcViz_Options_ProductView=True
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_TcViz_Options_ProductView ] Successfully performed [ " & sAction & " ] Action.")	
								End If
								
							Case Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_TcViz_Options_ProductView ] Invalid case [ " & sAction & " ].")
								Exit function
						End Select	
					Next	
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------							
				Case "CheckBoxSelect"
					DictKeys = dicInfo.Keys
					DictItems = dicInfo.Items
					For iCounter = 0 to Ubound(DictKeys)							
						Select Case DictKeys(iCounter)
							Case "Show Suppressed Occurrences"
								If DictItems(iCounter) <> "" Then	
									objWin.JavaCheckBox("ViewTogglestoConsider").SetTOProperty "attached text",DictKeys(iCounter) 
									If objWin.JavaCheckBox("ViewTogglestoConsider").Exist(1) Then								
										If Fn_SISW_UI_JavaCheckBox_Operations("Fn_SISW_TcViz_Options_ProductView", "Set", objWin.JavaCheckBox("ViewTogglestoConsider"), "" ,DictItems(iCounter)) Then
										 	iCount=iCount+1
										End If
									End If
								End If								
								
								If iCount = Ubound(DictKeys)+1 Then
									Fn_SISW_TcViz_Options_ProductView=True
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_TcViz_Options_ProductView ] Successfully performed [ " & sAction & " ] Action.")	
								End If
								
							Case Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_TcViz_Options_ProductView ] Invalid case [ " & sAction & " ].")
								Exit function
						End Select	
					Next					
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------							
				Case "IsEnabled"
					DictKeys = dicInfo.Keys
					DictItems = dicInfo.Items
					For iCounter = 0 to Ubound(DictKeys)							
						Select Case DictKeys(iCounter)
							Case "Show Unconfigured Variants","Show Unconfigured Changes","Show Suppressed Occurrences","Show Unconfigured Assigned Occurrences","Show Unconfigured By Occurrence Effectivity","Show GCS Connection Points"
								If DictItems(iCounter) <> "" Then	
									objWin.JavaCheckBox("ViewTogglestoConsider").SetTOProperty "attached text",DictKeys(iCounter) 
									If objWin.JavaCheckBox("ViewTogglestoConsider").Exist(1) Then								
										sValue = objWin.JavaCheckBox("ViewTogglestoConsider").GetROProperty("enabled")
										If sValue =  DictItems(iCounter) Then
											iCount=iCount+1
										End If
									End If
								End If								
								
								If iCount = Ubound(DictKeys)+1 Then
									Fn_SISW_TcViz_Options_ProductView=True
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_TcViz_Options_ProductView ] Successfully performed [ " & sAction & " ] Action.")	
								End If
								
							Case Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_TcViz_Options_ProductView ] Invalid case [ " & sAction & " ].")
								Exit function
						End Select
					Next	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------					
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_TcViz_Options_ProductView ] Invalid case [ " & sAction & " ].")
			Exit function
	End Select
	
	If sButton <> "" Then
		If objWin.JavaButton(sButton).Exist(2) = True Then
			objWin.JavaButton(sButton).Click	
		else
			Set objWin = nothing
			Exit function
		End if
	End If
	
	If Fn_SISW_TcViz_Options_ProductView <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [ Fn_SISW_TcViz_Options_ProductView ] executed successfuly with case [ " & sAction & " ].")
	else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_TcViz_Options_ProductView ] Failed to executed with case [ " & sAction & " ].")
	End If

	Set objWin = nothing
End Function

'=======================================================================================================================================================
'****************************************    Function to perform operations on the Nodes of the Mytc Viewer WinObject  ***************************************
'
''Function Name		 	:	Fn_SISW_TcViz_MyTcAssemblyTreeOperations()
'
''Description		    :    	Function to Perform operations on Assembly Tree in MyTeamcenter Persppective

''Parameters		    :	   1. sCalledFrom                : Called from which perspective
'									 2. sAction                : Action to perform
'									 3. sNodeName	     :  Node Name (Do not pass complete path)
'									 4. sReserve           : Reserve
'									 5. sPopupMenu       : Popup Menu
								
''Return Value		    :  	True \ False
'
''Examples				1. Select: 
'						bReturn = Fn_SISW_TcViz_MyTcAssemblyTreeOperations("", "Select", "Ellipse", "", "")
'						2. Exist
'						 bReturn = Fn_SISW_TcViz_MyTcAssemblyTreeOperations("", "Exist", "Ellipse", "", "")
'						3. IsChecked
'						bReturn = Fn_SISW_TcViz_MyTcAssemblyTreeOperations("", "IsChecked", "Ellipse", "", "")
'						4. IsSelected
'						 bReturn = Fn_SISW_TcViz_MyTcAssemblyTreeOperations("", "IsSelected", "Ellipse", "", "")
'						5. DeselectAllObjects
'						bReturn = Fn_SISW_TcViz_MyTcAssemblyTreeOperations("", "DeselectAllObjects","", "", "")

'History:
'	Developer Name				Date			     Rev. No.		Reviewer		Changes Done	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Reema W					1-Dec-14												
'-----------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_TcViz_MyTcAssemblyTreeOperations(sCalledFrom, sAction, sNodeName, sReserve, sPopupMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_TcViz_MyTcAssemblyTreeOperations"
	Dim bReturn, objAssemblyTreeView
	Fn_SISW_TcViz_MyTcAssemblyTreeOperations = False
	Err.Clear
	
	Set objAssemblyTreeView = JavaWindow("TcVizMainWin").JavaWindow("JApplet").JavaObject("JTViewer")
		
		Select Case sAction
				'------------------------Case : Exist => to find existance of the node from the Tree Table--------------------------------------------
		
			Case "Exist"
						bReturn = objAssemblyTreeView.Object.getViewer3DBean.findPart(sNodeName)
						If bReturn = 0 Then
							Set objAssemblyTreeView = Nothing
							Exit Function
						End If
				'------------------------Case : Select => to select the node from the Tree Table--------------------------------------------	
			Case "Select"
						objAssemblyTreeView.Object.getViewer3DBean.select(sNodeName)
				'------------------------Case : IsSelected => to select the node from the Tree Table--------------------------------------------	
			Case "IsSelected"
						bReturn = objAssemblyTreeView.Object.getViewer3DBean.IsSelected(sNodeName)
						If bReturn = false Then
							Set objAssemblyTreeView = Nothing
							Exit Function
						End If
					
			'------------------------Case : IsChecked => to Check checkbox of assemblytree is set or not--------------------------------------------			
			Case "IsChecked" 
                    bReturn =  objAssemblyTreeView.Object.getViewer3DBean.isVisible(sNodeName)
                    If Err.Number < 0  Then
                        Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to get Checkbox State.")
                        Exit Function
                    End If            
                    If lCase(bReturn) = "false" Then
                        Set objAssemblyTreeView=nothing
						Exit Function 
                    End If
				
			'------------------------Case : DeselectAllObjects => To Deselect all object --------------------------------------------
			Case "DeselectAllObjects"
				objAssemblyTreeView.Object.getViewer3DBean.deselectAllObjects	
				
		End Select
		
		If Err.Number <> 0 Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to perform [ " & sAction & " ] Action.")
			Set objAssemblyTreeView = Nothing
			Exit Function
		End If
		
	Set objAssemblyTreeView = Nothing
	Fn_SISW_TcViz_MyTcAssemblyTreeOperations = True
End Function
