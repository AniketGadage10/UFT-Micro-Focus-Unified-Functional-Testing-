Option explicit
'001. Fn_SISW_RAC_NatTable_Init()
'002. Fn_SISW_RAC_NatTable_MoveScrollBar()
'003. Fn_SISW_RAC_NatTable_ResetHorizontalScrollBar()
'004. Fn_SISW_RAC_NatTable_ResetVerticalScrollBar()
'005. Fn_SISW_RAC_NatTable_SetColumnVisible()
'006. Fn_SISW_RAC_NatTable_GetColumnIndex()
'007. Fn_SISW_RAC_NatTable_GetRowIndex()
'008. Fn_SISW_RAC_NatTable_TreeTable_GetRowIndex()
'009. Fn_SISW_RAC_NatTable_SetColumnVisibleExt()
'010. Fn_SISW_RAC_NatTable_GetColumnIndexExt()
' - - - - - - - - -  - - - - - - - - -  - - - - - - - - -  - - - - - - - - -  - - - - - - - - -  - - - - - - - - -  - - - - - - - - -  - - - - - - - - - 
Dim NT_HScrollBarX, NT_HScrollBarY, NT_HScrollBarH, NT_HScrollBarW
Dim NT_HScrollBarMax, NT_HScrollBarThumb

Dim NT_VScrollBarX, NT_VScrollBarY, NT_VScrollBarH, NT_VScrollBarW
Dim NT_VScrollBarMax, NT_VScrollBarThumb

Dim NT_iColumnIndex, NT_objHScrollBar, NT_objVScrollBar

' - - - - - - - - -  - - - - - - - - -  - - - - - - - - -  - - - - - - - - -  - - - - - - - - -  - - - - - - - - -  - - - - - - - - -  - - - - - - - - - 
NT_HScrollBarX = 0
NT_HScrollBarY = 0
NT_HScrollBarH = 0 
NT_HScrollBarW = 0
NT_HScrollBarMax = 0
NT_HScrollBarThumb = 0

NT_VScrollBarX = 0
NT_VScrollBarY = 0
NT_VScrollBarW = 0
NT_VScrollBarH = 0
NT_VScrollBarMax = 0
NT_VScrollBarThumb = 0
' - - - - - - - - -  - - - - - - - - -  - - - - - - - - -  - - - - - - - - -  - - - - - - - - -  - - - - - - - - -  - - - - - - - - -  - - - - - - - - - 
'*******************************************************************************************************************
'
''Function Name		 	:	Fn_SISW_RAC_NatTable_Init
'
''Description		    :  	Function to initialize NatTable values.

''Parameters		    :	1. objNatTable : Object Handle name
								
''Return Value		    :  	True \ false
'
''Examples		     	:	Fn_SISW_RAC_NatTable_Init(JavaWindow("Product Master Manager").JavaObject("NatTable"))

'History:
'	Developer Name				Date			Rev. No.	Reviewer		Changes Done	
'*******************************************************************************************************************
'	Koustubh Watwe		        17-Oct-2012	    1.0			Self	
'*******************************************************************************************************************
Public Function Fn_SISW_RAC_NatTable_Init(objNatTable)
	Dim sBounds, aBounds, sFuncLog
	sFuncLog = "Fn_SISW_RAC_NatTable_Init : on [ " & objNatTable.toString() & " ] : "
	Fn_SISW_RAC_NatTable_Init = False
'	If Fn_SISW_RAC_UI_Object_Operations("Fn_SISW_RAC_NatTable_Init", "Exist", objNatTable,"") = False Then
'		Call Fn_SISW_LogUtil_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & " FAIL : Failed to find NatTable object.")		
'		Call Fn_SISW_RAC_UI_ExitFromUI(sFuncLog)
'		Exit Function
'	End If
	NT_HScrollBarX = 0
	NT_HScrollBarY = 0
	NT_HScrollBarH = 0 
	NT_HScrollBarW = 0
	NT_HScrollBarMax = 0
	NT_HScrollBarThumb = 0

	NT_VScrollBarX = 0
	NT_VScrollBarY = 0
	NT_VScrollBarW = 0
	NT_VScrollBarH = 0
	NT_VScrollBarMax = 0
	NT_VScrollBarThumb = 0

	' Initializing horizontal scroll bar
	Set NT_objHScrollBar = objNatTable.Object.getHorizontalBar()
	If NT_objHScrollBar.isVisible() Then
		sBounds = NT_objHScrollBar.getBounds().toString()
		sBounds = mid(sBounds,instr(sBounds,"{")+1, len(sBounds) -instr(sBounds,"{")-1)
		aBounds = split(sBounds,",")
		NT_HScrollBarX = cDbl(trim(aBounds(0)))
		NT_HScrollBarY = cDbl(trim(aBounds(1)))
		NT_HScrollBarW = cDbl(trim(aBounds(2)))
		NT_HScrollBarH = cDbl(trim(aBounds(3)))
		NT_HScrollBarMax = cDbl(NT_objHScrollBar.getMaximum())
		NT_HScrollBarThumb = cDbl(NT_objHScrollBar.getThumb())
	End If

	' Initializing vertical scroll bar
	Set NT_objVScrollBar = objNatTable.Object.getVerticalBar()
	If NT_objVScrollBar.isVisible()  Then
		sBounds = NT_objVScrollBar.getBounds().toString()
		sBounds = mid(sBounds,instr(sBounds,"{")+1, len(sBounds) -instr(sBounds,"{")-1)
		aBounds = split(sBounds,",")
		NT_VScrollBarX = cDbl(trim(aBounds(0)))
		NT_VScrollBarY = cDbl(trim(aBounds(1)))
		NT_VScrollBarW = cDbl(trim(aBounds(2)))
		NT_VScrollBarH = cDbl(trim(aBounds(3)))
		NT_VScrollBarMax = cDbl(NT_objVScrollBar.getMaximum())
		NT_VScrollBarThumb = cDbl(NT_objVScrollBar.getThumb())
	End If
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & " PASS : Successfully initialized NatTable Object..")		
	Fn_SISW_RAC_NatTable_Init = True
End Function
'*******************************************************************************************************************
'
''Function Name		 	:	Fn_SISW_RAC_NatTable_MoveScrollBar
'
''Description		    :  	Function to scroll scrollbars of NatTable.

''Parameters		    :	1. objNatTable : Object Handle name
'						:	2. sDirection  : Direction Up / Down / Left / Right
								
''Return Value		    :  	True \ False
'
''Examples		     	:	Fn_SISW_RAC_NatTable_MoveScrollBar(JavaWindow("Product Master Manager").JavaObject("NatTable"), "Up")
''Examples		     	:	Fn_SISW_RAC_NatTable_MoveScrollBar(JavaWindow("Product Master Manager").JavaObject("NatTable"), "Down")
''Examples		     	:	Fn_SISW_RAC_NatTable_MoveScrollBar(JavaWindow("Product Master Manager").JavaObject("NatTable"), "Left")
''Examples		     	:	Fn_SISW_RAC_NatTable_MoveScrollBar(JavaWindow("Product Master Manager").JavaObject("NatTable"), "Right")

'History:
'	Developer Name				Date			Rev. No.	Reviewer		Changes Done	
'*******************************************************************************************************************
'	Koustubh Watwe		        17-Oct-2012	    1.0			Self	
'*******************************************************************************************************************
Public Function Fn_SISW_RAC_NatTable_MoveScrollBar(objNatTable, sDirection)
	Dim iCurrentX, iCurrentY, sFuncLog
	sFuncLog = "Fn_SISW_RAC_NatTable_MoveScrollBar : on [ " & objNatTable.toString() & " ] : "
	Call Fn_SISW_RAC_NatTable_Init(objNatTable)
	Select Case sDirection
		Case "Right"
			If NT_objHScrollBar.isVisible() Then
				iCurrentX = NT_objHScrollBar.getSelection()
				objNatTable.Click NT_HScrollBarX + NT_HScrollBarW - 5, NT_HScrollBarY + 5, "LEFT"
				Wait SISW_MICRO_TIMEOUT
				If iCurrentX < NT_objHScrollBar.getSelection() Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & " PASS : Successfully scrolled in direction [ " & sDirection & " ]")	
					Fn_SISW_RAC_NatTable_MoveScrollBar = True
				Else
					Fn_SISW_RAC_NatTable_MoveScrollBar = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & " FAIL : Failed to scroll in direction [ " & sDirection & " ]")				
				End If
			Else
				Fn_SISW_RAC_NatTable_MoveScrollBar = True
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & " PASS : Scrollbar is not available to scroll in direction [ " & sDirection & " ]")	
			End If
		Case "RightGrayArea"
			If NT_objHScrollBar.isVisible() Then
				iCurrentX = NT_objHScrollBar.getSelection()
				objNatTable.Click NT_HScrollBarX + NT_HScrollBarW - 20, NT_HScrollBarY + 5, "LEFT"
				Wait SISW_MICRO_TIMEOUT
				If iCurrentX < NT_objHScrollBar.getSelection() Then
					Fn_SISW_RAC_NatTable_MoveScrollBar = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & " PASS : Successfully scrolled in direction [ " & sDirection & " ]")	
				Else
					Fn_SISW_RAC_NatTable_MoveScrollBar = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & " FAIL : Failed to scroll in direction [ " & sDirection & " ]")				
				End If
			Else
				Fn_SISW_RAC_NatTable_MoveScrollBar = True
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & " PASS : Scrollbar is not available to scroll in direction [ " & sDirection & " ]")	
			End If
		Case "Left"
			If NT_objHScrollBar.isVisible() Then
				iCurrentX = NT_objHScrollBar.getSelection()
				objNatTable.Click NT_HScrollBarX + 5, NT_HScrollBarY + 5, "LEFT"
				Wait SISW_MICRO_TIMEOUT
				If iCurrentX > NT_objHScrollBar.getSelection() OR ( iCurrentX = 0 AND NT_objHScrollBar.getSelection() = 0 ) Then
					Fn_SISW_RAC_NatTable_MoveScrollBar = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & " PASS : Successfully scrolled in direction [ " & sDirection & " ]")	
				Else
					Fn_SISW_RAC_NatTable_MoveScrollBar = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & " FAIL : Failed to scroll in direction [ " & sDirection & " ]")				
				End If
			Else
				Fn_SISW_RAC_NatTable_MoveScrollBar = True
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & " PASS : Scrollbar is not available to scroll in direction [ " & sDirection & " ]")	
			End If
		Case "LeftGrayArea"
			If NT_objHScrollBar.isVisible() Then
				iCurrentX = NT_objHScrollBar.getSelection()
				objNatTable.Click NT_HScrollBarX + 20, NT_HScrollBarY + 5, "LEFT"
				Wait SISW_MICRO_TIMEOUT
				If iCurrentX > NT_objHScrollBar.getSelection() OR ( iCurrentX = 0 AND NT_objHScrollBar.getSelection() = 0 ) Then
					Fn_SISW_RAC_NatTable_MoveScrollBar = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & " PASS : Successfully scrolled in direction [ " & sDirection & " ]")	
				Else
					Fn_SISW_RAC_NatTable_MoveScrollBar = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & " FAIL : Failed to scroll in direction [ " & sDirection & " ]")				
				End If
			Else
				Fn_SISW_RAC_NatTable_MoveScrollBar = True
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & " PASS : Scrollbar is not available to scroll in direction [ " & sDirection & " ]")	
			End If
		Case "Up"
			If NT_objVScrollBar.isVisible() Then
				iCurrentY = NT_objVScrollBar.getSelection()
				objNatTable.Click NT_VScrollBarX + 5, NT_VScrollBarY + 5, "LEFT"
				Wait SISW_MICRO_TIMEOUT
				If iCurrentY > NT_objVScrollBar.getSelection() OR ( iCurrentY = 0 AND NT_objVScrollBar.getSelection() = 0 ) Then
					Fn_SISW_RAC_NatTable_MoveScrollBar = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & " PASS : Successfully scrolled in direction [ " & sDirection & " ]")	
				Else
					Fn_SISW_RAC_NatTable_MoveScrollBar = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & " FAIL : Failed to scroll in direction [ " & sDirection & " ]")				
				End If
			Else
				Fn_SISW_RAC_NatTable_MoveScrollBar = True
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & " PASS : Scrollbar is not available to scroll in direction [ " & sDirection & " ]")	
			End If
		Case "UpGrayArea"
			If NT_objVScrollBar.isVisible() Then
				iCurrentY = NT_objVScrollBar.getSelection()
				objNatTable.Click NT_VScrollBarX + 5, NT_VScrollBarY + 20, "LEFT"
				Wait SISW_MICRO_TIMEOUT
				If iCurrentY > NT_objVScrollBar.getSelection() OR ( iCurrentY = 0 AND NT_objVScrollBar.getSelection() = 0 ) Then
					Fn_SISW_RAC_NatTable_MoveScrollBar = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & " PASS : Successfully scrolled in direction [ " & sDirection & " ]")	
				Else
					Fn_SISW_RAC_NatTable_MoveScrollBar = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & " FAIL : Failed to scroll in direction [ " & sDirection & " ]")				
				End If
			Else
				Fn_SISW_RAC_NatTable_MoveScrollBar = True
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & " PASS : Scrollbar is not available to scroll in direction [ " & sDirection & " ]")	
			End If
		Case "Down"
			If NT_objVScrollBar.isVisible() Then
				iCurrentY = NT_objVScrollBar.getSelection()
				objNatTable.Click NT_VScrollBarX + 5, NT_VScrollBarY + NT_VScrollBarH - 5, "LEFT"
				Wait SISW_MICRO_TIMEOUT
				If iCurrentY < NT_objVScrollBar.getSelection() Then
					Fn_SISW_RAC_NatTable_MoveScrollBar = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & " PASS : Successfully scrolled in direction [ " & sDirection & " ]")	
				Else
					Fn_SISW_RAC_NatTable_MoveScrollBar = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & " FAIL : Failed to scroll in direction [ " & sDirection & " ]")				
				End If
			Else
				Fn_SISW_RAC_NatTable_MoveScrollBar = True
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & " PASS : Scrollbar is not available to scroll in direction [ " & sDirection & " ]")	
			End If
		Case "DownGrayArea"
			If NT_objVScrollBar.isVisible() Then
				iCurrentY = NT_objVScrollBar.getSelection()
				objNatTable.Click NT_VScrollBarX + 5, NT_VScrollBarY + NT_VScrollBarH - 20, "LEFT"
				Wait SISW_MICRO_TIMEOUT
				If iCurrentY < NT_objVScrollBar.getSelection() Then
					Fn_SISW_RAC_NatTable_MoveScrollBar = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & " PASS : Successfully scrolled in direction [ " & sDirection & " ]")	
				Else
					Fn_SISW_RAC_NatTable_MoveScrollBar = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & " FAIL : Failed to scroll in direction [ " & sDirection & " ]")				
				End If
			Else
				Fn_SISW_RAC_NatTable_MoveScrollBar = True
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & " PASS : Scrollbar is not available to scroll in direction [ " & sDirection & " ]")	
			End If
	End Select	
End Function
'*******************************************************************************************************************
'
''Function Name		 	:	Fn_SISW_RAC_NatTable_ResetHorizontalScrollBar
'
''Description		    :  	Function to reset position of horizontal scroll scrollbar of NatTable.

''Parameters		    :	1. objNatTable : Object Handle name
								
''Return Value		    :  	True \ false
'
''Examples		     	:	Fn_SISW_RAC_NatTable_ResetHorizontalScrollBar(JavaWindow("Product Master Manager").JavaObject("NatTable"))

'History:
'	Developer Name				Date			Rev. No.	Reviewer		Changes Done	
'*******************************************************************************************************************
'	Koustubh Watwe		        17-Oct-2012	    1.0			Self	
'*******************************************************************************************************************
Public Function Fn_SISW_RAC_NatTable_ResetHorizontalScrollBar(objNatTable)
	Dim sFuncLog
	sFuncLog = "Fn_SISW_RAC_NatTable_ResetHorizontalScrollBar : on [ " & objNatTable.toString() & " ] : "
	Fn_SISW_RAC_NatTable_ResetHorizontalScrollBar = False
	bResult = Fn_SISW_RAC_NatTable_Init(objNatTable)
	If bResult = False Then
		Call Fn_SISW_RAC_UI_ExitFromUI(sFuncLog)
		Exit Function
	End IF
	If NT_objHScrollBar.isVisible() Then
		objNatTable.Object.getHorizontalBar().setSelection 0
		Do until cDbl(objNatTable.Object.getHorizontalBar().getSelection()) = 0
			bResult = Fn_SISW_RAC_NatTable_MoveScrollBar(objNatTable, "LeftGrayArea")
			If bResult = False Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & " FAIL : Failed to reset horizontal scroll bar.")	
				Exit Function
			End IF			
		Loop
	End If
	Fn_SISW_RAC_NatTable_ResetHorizontalScrollBar = True
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & " PASS : Successfully reset horizontal scroll bar.")	
End Function
'*******************************************************************************************************************
'
''Function Name		 	:	Fn_SISW_RAC_NatTable_ResetVerticalScrollBar
'
''Description		    :  	Function to reset position of vertical scroll scrollbar of NatTable.

''Parameters		    :	1. objNatTable : Object Handle name
								
''Return Value		    :  	True \ false
'
''Examples		     	:	Fn_SISW_RAC_NatTable_ResetVerticalScrollBar(JavaWindow("Product Master Manager").JavaObject("NatTable"))

'History:
'	Developer Name				Date			Rev. No.	Reviewer		Changes Done	
'*******************************************************************************************************************
'	Koustubh Watwe		        17-Oct-2012	    1.0			Self	
'*******************************************************************************************************************
Public Function Fn_SISW_RAC_NatTable_ResetVerticalScrollBar(objNatTable)
	Dim sFuncLog
	sFuncLog = "Fn_SISW_RAC_NatTable_ResetVerticalScrollBar : on [ " & objNatTable.toString() & " ] : "
	Fn_SISW_RAC_NatTable_ResetVerticalScrollBar = False
	bResult = Fn_SISW_RAC_NatTable_Init(objNatTable)
	If bResult = False Then
		Call Fn_SISW_RAC_UI_ExitFromUI(sFuncLog)
		Exit Function
	End IF
	If NT_objVScrollBar.isVisible() Then
		objNatTable.Object.getVerticalBar().setSelection 0
		Do until cDbl(objNatTable.Object.getVerticalBar().getSelection()) = 0
			bResult = Fn_SISW_RAC_NatTable_MoveScrollBar(objNatTable, "UpGrayArea")
			If bResult = False Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & " FAIL : Failed to reset vertical scroll bar.")	
				Exit Function
			End IF
		Loop
	End If
	Fn_SISW_RAC_NatTable_ResetVerticalScrollBar = True
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & " PASS : Successfully reset vertical scroll bar.")	
End Function
'*******************************************************************************************************************
'
''Function Name		 	:	Fn_SISW_RAC_NatTable_SetColumnVisible
'
''Description		    :  	Function to make specified column visible in NatTable.

''Parameters		    :	1. objNatTable : Object Handle name
''						:	2. sColumnName : Column Name
								
''Return Value		    :  	True \ false
'
''Examples		     	:	Fn_SISW_RAC_NatTable_SetColumnVisible(JavaWindow("Product Master Manager").JavaObject("NatTable"),"Part Number")

'History:
'	Developer Name				Date			Rev. No.	Reviewer		Changes Done	
'*******************************************************************************************************************
'	Koustubh Watwe		        17-Oct-2012	    1.0			Self	
'*******************************************************************************************************************
Public Function Fn_SISW_RAC_NatTable_SetColumnVisible(objNatTable, sColumnName )
	Dim NatTableColumns, iCnt, sFuncLog
	sFuncLog = "Fn_SISW_RAC_NatTable_SetColumnVisible : on [ " & objNatTable.toString() & " ] : "
	Fn_SISW_RAC_NatTable_SetColumnVisible = False
	bResult = Fn_SISW_RAC_NatTable_Init(objNatTable)
	If bResult = False Then
		Call Fn_SISW_RAC_UI_ExitFromUI(sFuncLog)
		Exit Function
	End IF
	For iCnt = 1 to objNatTable.Object.getColumnCount() -1
		If sColumnName = objNatTable.Object.getCellByPosition(iCnt, 1).getDataValue().toString() Then
			Fn_SISW_RAC_NatTable_SetColumnVisible = True
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & " PASS : Successfully set column [ " & sColumnName & " ] visible.")
			Exit Function
		End If
	Next
	Call Fn_SISW_RAC_NatTable_MoveScrollBar(objNatTable, "Right")
	NatTableColumns = ""
	If NT_HScrollBarMax <> 0 Then
		Call Fn_SISW_RAC_NatTable_ResetHorizontalScrollBar(objNatTable) 
		Do until (cDbl(NT_objHScrollBar.getSelection()) + NT_HScrollBarThumb) = NT_HScrollBarMax
			For iCnt = 1 to objNatTable.Object.getColumnCount() -1
				If NatTableColumns <>  "" Then
					If instr(NatTableColumns, objNatTable.Object.getCellByPosition(iCnt, 1).getDataValue().toString()) > 0 Then
					Else
						NatTableColumns = NatTableColumns & "~" & objNatTable.Object.getCellByPosition(iCnt, 1).getDataValue().toString()	
					End If
				Else
					NatTableColumns = objNatTable.Object.getCellByPosition(iCnt, 1).getDataValue().toString()	
				End If
			Next
			If instr(NatTableColumns, sColumnName) > 0 Then
				Fn_SISW_RAC_NatTable_SetColumnVisible = True
				Exit Function
			End If
			Call Fn_SISW_RAC_NatTable_MoveScrollBar(objNatTable, "Right")
		loop	
	End If

	' addditional code to make Cell visible
	If Fn_SISW_RAC_NatTable_SetColumnVisible <> False Then
		Call Fn_SISW_RAC_NatTable_MoveScrollBar(objNatTable, "Right")
		For iCnt = 1 to objNatTable.Object.getColumnCount() -1
			If sColumnName = objNatTable.Object.getCellByPosition(iCnt, 1).getDataValue().toString() Then
				Fn_SISW_RAC_NatTable_SetColumnVisible = True
				Exit Function
			End If
		Next
	End If
	If Fn_SISW_RAC_NatTable_SetColumnVisible <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & " PASS : Successfully set column [ " & sColumnName & " ] visible.")
	Else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & " FAIL : Failed to set column [ " & sColumnName & " ] visible.")
	End If
End Function
'*******************************************************************************************************************
'
''Function Name		 	:	Fn_SISW_RAC_NatTable_GetColumnIndex
'
''Description		    :  	Function to return column number of specified column from NatTable.

''Parameters		    :	1. objNatTable : Object Handle name
''						:	2. sColumnName : Column Name
								
''Return Value		    :  	Column Number \ -1
'
''Examples		     	:	Fn_SISW_RAC_NatTable_GetColumnIndex(JavaWindow("Product Master Manager").JavaObject("NatTable"),"Part Number")

'History:
'	Developer Name				Date			Rev. No.	Reviewer		Changes Done	
'*******************************************************************************************************************
'	Koustubh Watwe		        17-Oct-2012	    1.0			Self	
'*******************************************************************************************************************
Public Function Fn_SISW_RAC_NatTable_GetColumnIndex(objNatTable, sColumnName )
	Dim iCnt, sFuncLog
	sFuncLog = "Fn_SISW_RAC_NatTable_GetColumnIndex : on [ " & objNatTable.toString() & " ] : "
	Fn_SISW_RAC_NatTable_GetColumnIndex = -1
	Call Fn_SISW_RAC_NatTable_SetColumnVisible(objNatTable, sColumnName)
	For iCnt = 1 to objNatTable.Object.getColumnCount() -1
		If sColumnName = objNatTable.Object.getCellByPosition(iCnt, 1).getDataValue().toString() then
			Fn_SISW_RAC_NatTable_GetColumnIndex = iCnt
			Exit Function
		End If
	Next
	If Fn_SISW_RAC_NatTable_GetColumnIndex <> -1 Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & " PASS : Successfully found column [ " & sColumnName & " ].")
	Else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & " FAIL : Failed to find column [ " & sColumnName & " ].")
	End If
End Function
'*******************************************************************************************************************
'
''Function Name		 	:	Fn_SISW_RAC_NatTable_GetRowIndex
'
''Description		    :  	Function to return column number of specified column from NatTable.

''Parameters		    :	1. objNatTable : Object Handle name
''						:	2. sColumnName : Column Name
''						:	3. sValue : Row Name
''						:	4. sInstanceHandler : instance handler
								
''Return Value		    :  	Row Number \ -1
'
''Examples		     	:	Fn_SISW_RAC_NatTable_GetRowIndex(JavaWindow("Product Master Manager").JavaObject("NatTable"),"Part Number", "Prt001","")
''Examples		     	:	Fn_SISW_RAC_NatTable_GetRowIndex(JavaWindow("Product Master Manager").JavaObject("NatTable"),"Part Number", "Prt001 @3","@")

'History:
'	Developer Name				Date			Rev. No.	Reviewer		Changes Done	
'*******************************************************************************************************************
'	Koustubh Watwe		        17-Oct-2012	    1.0			Self	
'*******************************************************************************************************************
Public Function Fn_SISW_RAC_NatTable_GetRowIndex(objNatTable, sColumnName, sValue, sInstanceHandler)
	Dim sFuncLog
	sFuncLog = "Fn_SISW_RAC_NatTable_GetRowIndex : on [ " & objNatTable.toString() & " ] : "
	Dim bFound, iColIndex, iRowIndex, iInstanceCounter, iInstanceNumber, aValue
	Fn_SISW_RAC_NatTable_GetRowIndex = -1
	bFound = False
	If sInstanceHandler = "" Then
		sInstanceHandler = "@"
	End If

	aValue = split(sValue, sInstanceHandler)
	iInstanceCounter = 1
	aValue(0) = trim(aValue(0))
	If UBound(aValue) = 1 Then
		iInstanceCounter = cDbl(aValue(1))
	End If
	iInstanceNumber = iInstanceCounter

	iColIndex = Fn_SISW_RAC_NatTable_GetColumnIndex(objNatTable, sColumnName )
	If iColIndex = -1 Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & " FAIL : Failed to find column [ " & sColumnName & " ].")
		Exit function
	End If
	Call Fn_SISW_RAC_NatTable_ResetVerticalScrollBar(objNatTable) 
	sFuncLog = "Fn_SISW_RAC_NatTable_GetRowIndex : on [ " & objNatTable.toString() & " ] : "
	Do until False
		For iRowIndex = 3 to objNatTable.Object.getRowCount() -1
			If aValue(0) = trim(objNatTable.Object.getCellByPosition(iColIndex, iRowIndex).getDataValue().toString()) Then
				If iInstanceCounter = 1 Then
					bFound = True
					Fn_SISW_RAC_NatTable_GetRowIndex = iRowIndex
					Exit for
				End If
				iInstanceCounter = iInstanceCounter - 1
			End If
		Next
		If bFound Then
			Exit do
		End If

		If NT_objVScrollBar.isVisible() Then
			If (cDbl(NT_objVScrollBar.getSelection()) + NT_VScrollBarThumb) = NT_VScrollBarMax Then
				Exit do
			Else
				Call Fn_SISW_RAC_NatTable_MoveScrollBar(objNatTable, "Down")
				iInstanceCounter = iInstanceNumber 
			End If
		Else
			Exit do
		End If
	loop	
	If iRowIndex <> -1 Then
		' addditional code to make Cell visible
		Call Fn_SISW_RAC_NatTable_MoveScrollBar(objNatTable, "Down")
		iInstanceCounter = iInstanceNumber 
		For iRowIndex = 3 to objNatTable.Object.getRowCount() -1
			If aValue(0) = trim(objNatTable.Object.getCellByPosition(iColIndex, iRowIndex).getDataValue().toString()) Then
				If iInstanceCounter = 1 Then
					bFound = True
					Fn_SISW_RAC_NatTable_GetRowIndex = iRowIndex
					Exit for
				End If
				iInstanceCounter = iInstanceCounter - 1
			End If
		Next
	End If
	If Fn_SISW_RAC_NatTable_GetRowIndex <> -1 Then 
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & " PASS : Successfully found [ " & sValue & " in " & sColumnName & " ].")
	Else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & " FAIL : Failed to find [ " & sValue & " in " & sColumnName & " ].")
	End If
End Function
'*******************************************************************************************************************
''Function Name		 	:	Fn_SISW_RAC_NatTable_TreeTable_GetRowIndex
'
''Description		    :  	Function to return row number of specified value in given column from NatTable.

''Parameters		    :	1. objNatTable : Object name
''						:	2. StrNode : Node Path
''						:	3. sDelimiter : Delimiter
''						:	4. sInstanceHandler : Instance handler 
''						:	5. sColumnIndex : Column Index
'							6: iCoumnNumber : column number
''						:	6. sRowStartingIndex : starting Row Index
''						:   7. iRowIndexIncrementor : Row index incrementor
								
''Return Value		    :  	Row Number \ -1
'
''Examples		     	:	Fn_SISW_RAC_NatTable_TreeTable_GetRowIndex(JavaWindow("ProductMasterManager").JavaObject("NatTable"),  "CM48830:CM38830:AS58830", "","", "","","",2)

'History:
'	Developer Name				Date			Rev. No.	Reviewer		Changes Done	
'*******************************************************************************************************************
'	Sandeep			6-May-2013		1.0			Self	
'*******************************************************************************************************************
Public Function Fn_SISW_RAC_NatTable_TreeTable_GetRowIndex(objNatTable, sColumnIndex,iCoumnNumber, StrNode, sRowStartingIndex, sDelimiter, sInstanceHandler,iRowIndexIncrementor)
   'Variable Declaration
	Dim sItemPath,aStrNode,bFlag,i,iNodeItemsCount, sFuncLog
	Dim oCurrentNode,eStrNode, iCount, iNodecnt
	Dim iInstanceCnt, aNode,iOccCnt
	Dim sTreeNodeStr
	Dim objRootObjects, objDataProvider, iRowIndex, iColIndex
	
	sFuncLog = "Fn_SISW_RAC_NatTable_TreeTable_GetRowIndex : on [ " & objNatTable.toString() & " ] : "

	If sColumnIndex = "" Then
		iColIndex = 1
    Elseif sColumnIndex="" and iColumnNumber<>"" then
		iColIndex = cDbl(iColumnNumber)
	Else
		iColIndex =Fn_SISW_RAC_NatTable_GetColumnIndexExt(objNatTable, sColumnIndex,"","","" )
		If iColIndex = -1 Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & " FAIL : Failed to find column [ " & sColumnIndex & " ].")
			Exit function
		End If
	End If
	If sRowStartingIndex = "" Then
		sRowStartingIndex = 2
	End If
	If Environment.Value("ProductName") = sQTPProductName OR Environment.Value("ProductName") = sUFTProductName Then '' added code by Ankit T(02-Jul-2014) to handle UFT related issue as some method not supported in UFT 
		Set objDataProvider = objNatTable.Object.getCellByPosition(iColIndex, sRowStartingIndex).getlayer.getRowHeaderLayer.getBaselayer.getCellByPosition(iColIndex,sRowStartingIndex).getSourcelayer().getDataProvider()
	Else
		Set objDataProvider = objNatTable.Object.getCellByPosition(iColIndex, sRowStartingIndex).getSourceLayer().getDataProvider()
	End If
	Set objRootObjects = objDataProvider.getTreeList().getRoots()

	If sDelimiter = "" Then sDelimiter = ":"
	If sInstanceHandler = "" Then sInstanceHandler = "@"
	Fn_SISW_RAC_NatTable_TreeTable_GetRowIndex = False
	sTreeNodeStr = ""
	'Initial Item Path
	sItemPath= -1
	aStrNode = Split (StrNode, sDelimiter)
	bFlag=False
	
	'To handle the situation where operation needs to be performed on Root Node
	iOccCnt = 1
	For iCount = 0 to cDbl(objRootObjects.size()) - 1
'	For iCount = 0 to cDbl(objDataProvider.getRowCount) - 1
		If Instr(aStrNode(0), sInstanceHandler) > 0 Then
			aNode = split(aStrNode(0),sInstanceHandler)
			eStrNode = trim(aNode(0))
			iInstanceCnt = cDbl(aNode(1) )
		Else
			eStrNode = trim(aStrNode(0))
			iInstanceCnt = 1
		End If
		If iRowIndexIncrementor="" Then
			iRowIndexIncrementor=1
		End If
		' Index of the Row then get Row text
 		iRowIndex = objDataProvider.indexOfRowObject(objRootObjects.get(iCount).getElement())
		iRowIndex = cDbl(iRowIndex) 
		iRowIndex = cDbl(iRowIndex) + iRowIndexIncrementor
		
		sTreeNodeStr = objNatTable.Object.getCellByPosition(0, iRowIndex).getDataValue().getData().toString()
		If sTreeNodeStr="" Then
			sTreeNodeStr = objNatTable.Object.getCellByPosition(0, iRowIndex).getDataValue().getData().getName()
		End If
		If sTreeNodeStr = eStrNode Then
			If  iOccCnt = iInstanceCnt Then
				Set oCurrentNode = objRootObjects.get(iCount)
				sItemPath = iRowIndex 
				bFlag = True
				Exit For
			else
				iOccCnt = iOccCnt + 1
			End If
		End If
	Next
	If UBound(aStrNode) = 0 Then
		Fn_SISW_RAC_NatTable_TreeTable_GetRowIndex = sItemPath
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & " PASS : executed successfully for item [ " & StrNode & " ]"  )
		Exit Function
	End If
	If bFlag Then
		bFlag = False
	Else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & "Failed to find item [ " & StrNode & " ]"  )
		Exit function
	End If
		'To Select first Occurance of Node
'		For each eStrNode In 
		For iNodecnt = 1 to UBound(aStrNode)
			eStrNode = aStrNode(iNodecnt)
			iNodeItemsCount = cDbl(oCurrentNode.getChildren().size())
'			iNodeItemsCount = cDbl(oCurrentNode.size())
			bFlag=False
			iOccCnt = 1
			If Instr(eStrNode, sInstanceHandler) > 0 Then
				aNode = split(eStrNode,sInstanceHandler)
				eStrNode = trim(aNode(0))
				iInstanceCnt = cDbl(aNode(1) )
			Else
				iInstanceCnt = 1
			End If
			For i = 0 to iNodeItemsCount - 1
				' get text from table
				iRowIndex = objDataProvider.indexOfRowObject(oCurrentNode.getChildren().get(i).getElement())
				iRowIndex = cDbl(iRowIndex) + 1
				sTreeNodeStr = objNatTable.Object.getCellByPosition(0, iRowIndex).getDataValue().getData().toString()
				If Trim(sTreeNodeStr) = Trim(eStrNode) Then
					If  iOccCnt = iInstanceCnt Then
						sItemPath = objDataProvider.indexOfRowObject(oCurrentNode.getChildren().get(i).getElement())
						Set oCurrentNode = oCurrentNode.getChildren().get(i)
						bFlag=True
						Exit For
					else
						iOccCnt = iOccCnt + 1
					End If
				End If
			Next
			If bFlag=False Then
				Exit For
			End If
		Next 
	If bFlag=True Then
		'Function Returns Item Path
		Fn_SISW_RAC_NatTable_TreeTable_GetRowIndex = (sItemPath + 1)
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & " PASS : executed successfully for item [ " & StrNode & " ]"  )
	Else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & "Failed to find item [ " & StrNode & " ]"  )
		Fn_SISW_RAC_NatTable_TreeTable_GetRowIndex = False
	End If
	Set oCurrentNode =Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_SISW_RAC_NatTable_SetColumnVisibleExt

'Description			 :	Function to make specified column visible in NatTable

'Parameters			   :  1.objNatTable : Object Handle name
'									2.sColumnName : Column Name
'									3.iColumnNumber: Column number
'									3.iStartColIndex: Column start index { option }
'									4.iColumnRow: Column Row number { option }
'
'Return Value		   : 	True or False

'Examples				:   bReturn=Fn_SISW_RAC_NatTable_SetColumnVisibleExt(JavaWindow("ProductConfigurator").JavaObject("VariantNatTable"),"Part Number","","")
'									    bReturn=Fn_SISW_RAC_NatTable_SetColumnVisibleExt(JavaWindow("ProductConfigurator").JavaObject("VariantNatTable"),"Part Number",1,0)
'
'History					 :			
'				Developer Name						Date					Rev. No.						Changes Done																				Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'				Sandeep N							8-May-2013				1.0																															Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Public Function Fn_SISW_RAC_NatTable_SetColumnVisibleExt(objNatTable,sColumnName,iColumnNumber,iStartColIndex,iColumnRow)
	Dim NatTableColumns, iCnt, sFuncLog
	sFuncLog = "Fn_SISW_RAC_NatTable_SetColumnVisibleExt : on [ " & objNatTable.toString() & " ] : "
	Fn_SISW_RAC_NatTable_SetColumnVisibleExt = False
	bResult = Fn_SISW_RAC_NatTable_Init(objNatTable)
	If bResult = False Then
		Call Fn_SISW_RAC_UI_ExitFromUI(sFuncLog)
		Exit Function
	End IF
	If iColumnRow="" Then
		iColumnRow=0
	End If
	If iStartColIndex="" Then
		If Environment.Value("ProductName") = sQTPProductName OR Environment.Value("ProductName") = sUFTProductName Then
			iStartColIndex=0
		Else
			iStartColIndex=1
		End If
	End If
    If sColumnName="" and iColumnNumber<>"" Then
		If cDbl(objNatTable.Object.getColumnCount() )-1>=cDbl(iColumnNumber) Then
			Fn_SISW_RAC_NatTable_SetColumnVisibleExt = True
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & " PASS : Successfully set column [ " & iColumnNumber & " ] visible.")
			Exit Function
		End If
	Else
		For iCnt = iStartColIndex to objNatTable.Object.getColumnCount() -1
			If Environment.Value("ProductName") = sQTPProductName OR Environment.Value("ProductName") = sUFTProductName Then  '' '' added code by Ankit T(02-Jul-2014) to handle UFT related issue as some method not supported in UFT 
				   If Not IsEmpty(objNatTable.Object.getCellByPosition(iCnt, iColumnRow).getlayer.getColumnHeaderlayer.getDatavalueByPosition(iCnt,iColumnRow)) Then
					If CStr(sColumnName) = CStr(objNatTable.Object.getCellByPosition(iCnt, iColumnRow).getlayer.getColumnHeaderlayer.getDatavalueByPosition(iCnt,iColumnRow).tostring) Then
						Fn_SISW_RAC_NatTable_SetColumnVisibleExt = True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & " PASS : Successfully set column [ " & sColumnName & " ] visible.")
						Exit Function
					End if
				End If
			Else
				If Not objNatTable.Object.getCellByPosition(iCnt, iColumnRow).getDataValue() is Nothing Then
					If CStr(sColumnName) = CStr(objNatTable.Object.getCellByPosition(iCnt, iColumnRow).getDataValue().toString()) Then
			 			Fn_SISW_RAC_NatTable_SetColumnVisibleExt = True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & " PASS : Successfully set column [ " & sColumnName & " ] visible.")
						Exit Function
					End if
				End If
			End If
		Next
	End if
	Call Fn_SISW_RAC_NatTable_MoveScrollBar(objNatTable, "Right")
	NatTableColumns = ""
	If NT_HScrollBarMax <> 0 Then
		Call Fn_SISW_RAC_NatTable_ResetHorizontalScrollBar(objNatTable) 
		Do until (cDbl(NT_objHScrollBar.getSelection()) + NT_HScrollBarThumb) = NT_HScrollBarMax
			For iCnt = iStartColIndex to objNatTable.Object.getColumnCount() -1
				If NatTableColumns <>  "" Then
					If instr(NatTableColumns, objNatTable.Object.getCellByPosition(iCnt, iColumnRow).getDataValue().toString()) > 0 Then
					Else
						NatTableColumns = NatTableColumns & "~" & objNatTable.Object.getCellByPosition(iCnt, iColumnRow).getDataValue().toString()	
					End If
				Else
					NatTableColumns = objNatTable.Object.getCellByPosition(iCnt, iColumnRow).getDataValue().toString()	
				End If
			Next
			If instr(NatTableColumns, CStr(sColumnName)) > 0 Then
				Fn_SISW_RAC_NatTable_SetColumnVisibleExt = True
				Exit Function
			End If
			Call Fn_SISW_RAC_NatTable_MoveScrollBar(objNatTable, "Right")
		loop	
	End If

	' addditional code to make Cell visible
	If Fn_SISW_RAC_NatTable_SetColumnVisibleExt <> False Then
		Call Fn_SISW_RAC_NatTable_MoveScrollBar(objNatTable, "Right")
		For iCnt = iStartColIndex to objNatTable.Object.getColumnCount() -1
			If Cstr(sColumnName) = CStr(objNatTable.Object.getCellByPosition(iCnt, iColumnRow).getDataValue().toString()) Then
				Fn_SISW_RAC_NatTable_SetColumnVisibleExt = True
				Exit Function
			End If
		Next
	End If
	If Fn_SISW_RAC_NatTable_SetColumnVisibleExt <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & " PASS : Successfully set column [ " & sColumnName & " ] visible.")
	Else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & " FAIL : Failed to set column [ " & sColumnName & " ] visible.")
	End If
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_SISW_RAC_NatTable_GetColumnIndexExt

'Description			 :	Function to return column number of specified column from NatTable

'Parameters			   :  1.objNatTable : Object Handle name
'									2.sColumnName : Column Name
'									3.iCoumnNumber : Column number
'									3.iStartColIndex: Column start index { option }
'									4.iColumnRow: Column Row number { option }
'
'Return Value		   : 	True or False
'
'Examples				:   bReturn=Fn_SISW_RAC_NatTable_GetColumnIndexExt(JavaWindow("ProductConfigurator").JavaObject("VariantNatTable"),"Part Number","","")
'									    bReturn=Fn_SISW_RAC_NatTable_GetColumnIndexExt(JavaWindow("ProductConfigurator").JavaObject("VariantNatTable"),"Part Number",1,0)
'
'History					 :			
'				Developer Name						Date					Rev. No.						Changes Done																				Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'				Sandeep N							8-May-2013				1.0																															Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Public Function Fn_SISW_RAC_NatTable_GetColumnIndexExt(objNatTable, sColumnName,iCoumnNumber,iStartColIndex,iColumnRow)
	Dim iCnt, sFuncLog
	sFuncLog = "Fn_SISW_RAC_NatTable_GetColumnIndexExt : on [ " & objNatTable.toString() & " ] : "
	Fn_SISW_RAC_NatTable_GetColumnIndexExt = -1
    If sColumnName="" and iCoumnNumber<>"" Then
		If cDbl(objNatTable.Object.getColumnCount()) -1>= cDbl(iCoumnNumber) Then
			Fn_SISW_RAC_NatTable_GetColumnIndexExt = cDbl(iCoumnNumber)
			Exit Function
		End If
	End if
	Call Fn_SISW_RAC_NatTable_SetColumnVisibleExt(objNatTable, sColumnName,"",iStartColIndex,iColumnRow)
	If iColumnRow="" Then
		iColumnRow=0
	End If
	If iStartColIndex="" Then
		If Environment.Value("ProductName") = sQTPProductName OR Environment.Value("ProductName") = sUFTProductName Then
			iStartColIndex=0
		Else
			iStartColIndex=1
		End If
	End If

	For iCnt = iStartColIndex to objNatTable.Object.getColumnCount() -1
		If Environment.Value("ProductName") = sQTPProductName OR Environment.Value("ProductName") = sUFTProductName Then  '' added code by Ankit T(02-Jul-2014) to handle UFT related issue as some method not supported in UFT 
           If Not IsEmpty(objNatTable.Object.getCellByPosition(iCnt, iColumnRow).getlayer.getColumnHeaderlayer.getDatavalueByPosition(iCnt,iColumnRow))Then
				If CStr(sColumnName) = CStr(objNatTable.Object.getCellByPosition(iCnt, iColumnRow).getlayer.getColumnHeaderlayer.getDatavalueByPosition(iCnt,iColumnRow).tostring) Then
					Fn_SISW_RAC_NatTable_GetColumnIndexExt = iCnt
					Exit Function
				End If
			End If
		Else
	    	If Not objNatTable.Object.getCellByPosition(iCnt, iColumnRow).getDataValue() is Nothing Then
				If Cstr(sColumnName) = CStr(objNatTable.Object.getCellByPosition(iCnt, iColumnRow).getDataValue().toString()) Then
 					Fn_SISW_RAC_NatTable_GetColumnIndexExt = iCnt
					Exit Function
				End If
			End If
		End If
	Next
	If Fn_SISW_RAC_NatTable_GetColumnIndexExt <> -1 Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & " PASS : Successfully found column [ " & sColumnName & " ].")
	Else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & " FAIL : Failed to find column [ " & sColumnName & " ].")
	End If
End Function
