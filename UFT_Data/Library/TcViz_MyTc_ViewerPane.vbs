Option Explicit

'=======================================================================================================================================================
' Function List
'-------------------------------------------------------------------------------------------------------------------------------------------------------
'						Function Name																		|					Created By
'-------------------------------------------------------------------------------------------------------------------------------------------------------
'1. Fn_TcViz_myTc_ToolbarOperation()												|	Vallari S (vallari.shimpukade@siemens.com)
'=======================================================================================================================================================

'--------------------------------------------------------------------------------------------------------------------
' Function Number   	: 1                                                                                
' Function Name     	: Fn_TcViz_myTc_ToolbarOperation
' Function Description  : Saves Session
' Function Usage    	: Result = Fn_TcViz_myTc_ToolbarOperation(sAction, sToolbarName, sButtonName)
'							
'                     		return True/False
'--------------------------------------------------------------------------------------------------------------------
Public Function Fn_TcViz_MyTc_ToolbarOperation(sAction, sToolbarName, sButtonName)
	GBL_FAILED_FUNCTION_NAME="Fn_TcViz_MyTc_ToolbarOperation"
   Dim strMenu
   Dim arrItems, iCnt

   Fn_TcViz_myTc_ToolbarOperation = False

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
								For iCnt = 1 To Ubound(arrItems) Step 1
									If lcase(trim(sButtonName)) = lcase(trim(arrItems(iCnt))) Then
										JavaWindow("TcVizMyTcMainWin").WinToolbar("2D Markup").Press iCnt, micLeftBtn
										Exit For
									End If
								Next
								
'									Select Case sButtonName
'											Case "Enable Markup"
'													JavaWindow("TcVizMyTcMainWin").WinToolbar("2D Markup").Press 1,micLeftBtn
'											Case "Select"
'													JavaWindow("TcVizMyTcMainWin").WinToolbar("2D Markup").Press 2,micLeftBtn
'											Case "Restricted Text"
'													JavaWindow("TcVizMyTcMainWin").WinToolbar("2D Markup").Press 14,micLeftBtn
'											Case "Unrestricted Text"
'													JavaWindow("TcVizMyTcMainWin").WinToolbar("2D Markup").Press 15,micLeftBtn
'									End Select

							Case "3D Markup"

					End Select

					If Err.Number < 0 Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Click [" + sButtonName + "] of [" + sToolbarName + "] Toolbar")
					Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Successfully Clicked [" + sButtonName + "] of [" + sToolbarName + "] Toolbar")
							Fn_TcViz_myTc_ToolbarOperation = True
					End If

			Case "Enable"
					If Not JavaWindow("TcVizMyTcMainWin").WinToolbar(sToolbarName).Exist(3) Then
							If JavaWindow("TcVizMyTcMainWin").WinToolbar("2D Viewing").Exist(10) Then
                        			JavaWindow("TcVizMyTcMainWin").WinToolbar("2D Viewing").Click 3,4,micRightBtn
                            Elseif JavaWindow("TcVizMyTcMainWin").WinToolbar("3D Navigation").Exist(10) Then
									JavaWindow("TcVizMyTcMainWin").WinToolbar("3D Navigation").Click 3,4,micRightBtn
							End If
							strMenu = JavaWindow("TcVizMyTcMainWin").WinMenu("ContextMenu").BuildMenuPath(sToolbarName)
							JavaWindow("TcVizMyTcMainWin").WinMenu("ContextMenu").Select strMenu
					End If

					If JavaWindow("TcVizMyTcMainWin").WinToolbar(sToolbarName).Exist(10) Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Successfully Enabled Toolbar [" + sToolbarName + "]")
							Fn_TcViz_myTc_ToolbarOperation = True
					Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Enable Toolbar [" + sToolbarName + "]")
					End If

			Case "Disable"
					If JavaWindow("TcVizMyTcMainWin").WinToolbar(sToolbarName).Exist(10) Then
							If JavaWindow("TcVizMyTcMainWin").WinToolbar("2D Viewing").Exist(10) Then
                        			JavaWindow("TcVizMyTcMainWin").WinToolbar("2D Viewing").Click 3,4,micRightBtn
                            Elseif JavaWindow("TcVizMyTcMainWin").WinToolbar("3D Navigation").Exist(10) Then
									JavaWindow("TcVizMyTcMainWin").WinToolbar("3D Navigation").Click 3,4,micRightBtn
							End If
							strMenu = JavaWindow("TcVizMyTcMainWin").WinMenu("ContextMenu").BuildMenuPath(sToolbarName)
							JavaWindow("TcVizMyTcMainWin").WinMenu("ContextMenu").Select strMenu
					End If

					If JavaWindow("TcVizMyTcMainWin").WinToolbar(sToolbarName).Exist(3) Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Disable Toolbar [" + sToolbarName + "]")
					Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Successfully Disabled Toolbar [" + sToolbarName + "]")
							Fn_TcViz_myTc_ToolbarOperation = True
					End If

			Case "Exists"
					If JavaWindow("TcVizMyTcMainWin").WinToolbar(sToolbarName).Exist(10) Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Successfully Found Existance of Toolbar [" + sToolbarName + "]")
							Fn_TcViz_myTc_ToolbarOperation = True
					Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Find Toolbar [" + sToolbarName + "]")
					End If

	End Select

End Function