Option Explicit
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'Declaring dectionary for Power Point Operations : Fn_MSO_PPTOperations
Dim dicPPTInfo
Set dicPPTInfo = CreateObject( "Scripting.Dictionary" )
With dicPPTInfo  
			 .Add "PresentationName",""
			 .Add "SlideNumber",""
			 .Add "PresentationCloseFlag",""
			 .Add "PowerPointQuitFlag",""
			 .Add "Title",""	
			 .Add "BodyText",""	
End with
extern.Declare micLong,"EmptyClipboard","user32.dll","EmptyClipboard"
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 

'*********************************************************	Function List		***********************************************************************
'00. Fn_SISW_MSO_GetObject()
'01. Fn_MSO_WordShapesOperation()
'02. Fn_MSO_WordEditOperation()
'03. Fn_MSO_ExcelEditOperations()
'04. Fn_MSO_WordHeaderFooter()
'05. Fn_MSO_PictureManagerOperation()
'06. Fn_MSO_FileOperations()
'07. Fn_MSO_ExcelColHeader()
'08. Fn_MSO_ExtractExcelRowNumber()
'09. Fn_MSO_ExtractExcelColumnName()
'10. Fn_MSO_ExcelDialogHandler()
'11. Fn_MSO_TeamcenterLogin()
'12. Fn_MSO_ImportToTeamcenter()
'13. Fn_MSO_DataFileAssignIDs()
'14. Fn_MSO_DocumentMapOperation()
'15. Fn_MSO_CreateAndFollow_HyperLink()
'16. Fn_MSO_Find_And_ReplaceText()
'17. Fn_MSO_ItemBasicCreate()
'18. Fn_MSO_CreateFolder()
'19. Fn_MSO_SpecificationOperations()
'20. Fn_MSO_Click_ImportToTeamCenter_Button()
'21. Fn_MSO_RenameSheetColumnHeader()
'22. Fn_MSO_OutlineOperations()
'23. Fn_MSO_PPTOperations()
'24. Fn_MSO_MarkupOperations()
'25. Fn_MSO_Word_RequirementTreeOperations()
'26. Fn_SISW_MSO_FolderViewTreeOperations()
'27. Fn_MSO_ZipFileOperations()
'28. Fn_SISW_MSO_PropertyOperations()
'29. Fn_MSO_VerifyErrorAfterEdit()
'30. Fn_SISW_MSO_RibbonbuttonClick()
'31. Fn_RemoveHyperLink()   ' Note : Function was used in script but was not present in any library.
'32. Fn_MSO_WordErrorDialogOperations()
'33. Fn_MSO_HiddenExcelOperation()
'34. Fn_MSO_CancelChckInChckOut()
'35. Fn_MSO_WorkFlowProcess
'36. Fn_MSO_CheckIn_CheckoutOperations
'37. Fn_MSO_WpfButton_Click
'38. Fn_MSO_Navigate
'39. Fn_MSO_FolderViewTreeOperations
'40. Fn_MSO_RibbonButton_Operations
'41. Fn_MSO_TeamcenterLogout
'42. Fn_MSO_SetFocusOnApplicationWindow
'43. Fn_MSO_BasicTeamcenterPreferences_Ops
'44. Fn_MSO_Session_Operations
'45. Fn_MSO_CalendarDialogOps
'46. Fn_MSO_AdvancedSearchOps
'47. Fn_MSO_PerformWorkflowTaskOps
'48. Fn_MSO_VerifyPopupMessage
'49. Fn_MSO_MyWorkList_Operations
'50. Fn_MSO_DeleteOperations
'51. Fn_MSO_NavTreeInWindow_Operations
'52. Fn_MSO_TeamcenterSaveAs_Operations
'53. Fn_MSO_SelectSignoffTeam_Ops
'54. Fn_MSO_SignoffTaskAndReviewersDecisions_Ops
'55. Fn_MSO_Delegate_Ops
'56. Fn_MSO_Revise_Ops
'57. Fn_MSO_MainTabOperations
'58. Fn_MSO_TeamcenterOpen_Operations
'59. Fn_MSO_ErrorDialog_Ops
'60. Fn_MSO_BrowseTreeOperations
'*********************************************************	Function List		***********************************************************************
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 			Function to get Object hierarchy  		- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''Function Name		:	Fn_SISW_MSO_GetObject
'
''Description		  	 :  	Function to get Object hierarchy

''Parameters		   :	1. sObjectName : Object Handle name
								
''Return Value		   :  	Object \ Nothing
'
''Examples		     	:	Fn_SISW_MSO_GetObject("MicrosoftExcel")

'History:                
'								Developer Name							Date				Rev. No.		Reviewer		Changes Done	
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'									Sonal Padmawar		 				14-Feb-2013				1.0					Sandeep
'-----------------------------------------------------------------------------------------------------------------------------------
'	Ashwini Kumar		 25-Sept-2013		2.0								Externalized object hierarchies
'-----------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_MSO_GetObject(sObjectName)
	Dim sObjectXMLPath
	sObjectXMLPath = Fn_GetEnvValue("User", "AutomationDir") & "\TestData\AutomationXML\ObjectXML\MSOffice.xml"
	Set Fn_SISW_MSO_GetObject = Fn_SISW_Setup_GetObjectFromXML(sObjectXMLPath, sObjectName)
End Function
'------------------------------------------------------------------------Function for Drawing Shapes in MS Word.----------------------------------------------------------------------------------------------------------------------------
'Function Name		  :	  Fn_MSO_WordShapesOperation

'Description			 :	 Function for Drawing Shapes in MS Word and filling color in the shapes.

'Parameters			   :	sAction, sFileStatus, sFilePath, sShape, iRed, iGreen, iBlue, Xco, Yco, iHeight, iWidth, sExtra

'Return Value		   : 	True Or False

'Examples				:	Msgbox Fn_MSO_WordShapesOperation("DrawAndColor", "Closed", "D:\Mainline\Test.docx", "Rectangle", "100", "54", "200", "100", "100", "100", "100", "")	
'							   Msgbox Fn_MSO_WordShapesOperation("DrawAndColor", "Opned", "D:\Mainline\Test1.docx", "Rectangle", "0", "255", "0", "100", "100", "100", "100", "")	
'							   Msgbox Fn_MSO_WordShapesOperation("DrawAndColor", "GetShapeColour", "D:\Mainline\Test1.docx", "Rectangle", "", "", "", "100", "100", "100", "100", "")	
'							   'Case "GetShapeColour" :-   it returns value equivalent to RGB(iRed, iGreen, iBlue)

'History			:	Developer Name				Date				Rev. No.			Reviewed by			Changes Done
'---------------------------------------------------------------------------------------------------------------------------------------------
'						Ketan Raje				10/02/2011			     1.0		
'						Madhura P				31_07_2015				 1.0				Vivek Ahirrao		Added new case "AddPicture"
'---------------------------------------------------------------------------------------------------------------------------------------------
Function Fn_MSO_WordShapesOperation(sAction, sFileStatus, sFilePath, sShape, iRed, iGreen, iBlue, Xco, Yco, iHeight, iWidth, sExtra)	
	GBL_FAILED_FUNCTION_NAME="Fn_MSO_WordShapesOperation"
   Dim lobjWord, objShape
	   Select Case sAction
	   Case "DrawAndColor"
				Select Case sFileStatus
				Case "Closed"
						Set lobjWord = CreateObject("Word.Application")
						lobjWord.Visible = True
						lobjWord.Documents.Open sFilePath,False,False
						If Trim(Lcase(sShape)) = "rectangle" Then
							Set objShape = lobjWord.ActiveDocument.Shapes.AddShape(1, Xco, Yco, iHeight, iWidth)
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),sShape &" shape drawn successfully.")
						End If
						objShape.Select
						objShape.Fill.Visible = msoTrue
						objShape.Fill.ForeColor.RGB = RGB(iRed, iGreen, iBlue)
						objShape.Fill.Solid
						lobjWord.ActiveDocument.Activate
						lobjWord.ActiveDocument.SaveAs sFilePath
						lobjWord.Quit
						Fn_MSO_WordShapesOperation = True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Pass: Fn_MSO_WordShapesOperation successfully executed "+sAction+" action.")
				Case "Opned"
						Set lobjWord = GetObject(,"Word.Application")'This Case May Give Warning But Will not create any issues in Single or Batch Run.
						lobjWord.Visible = True						
						If Trim(Lcase(sShape)) = "rectangle" Then
							Set objShape = lobjWord.ActiveDocument.Shapes.AddShape(1, Xco, Yco, iHeight, iWidth)
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),sShape &" shape drawn successfully.")
						End If
						objShape.Select
						objShape.Fill.Visible = msoTrue
						objShape.Fill.ForeColor.RGB = RGB(iRed, iGreen, iBlue)
						objShape.Fill.Solid
						lobjWord.ActiveDocument.Activate
						lobjWord.ActiveDocument.SaveAs sFilePath
						lobjWord.Quit
						Fn_MSO_WordShapesOperation = True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Pass: Fn_MSO_WordShapesOperation successfully executed "+sAction+" action.")
				Case "GetShapeColour"
						If sFilePath <> "" then
							Set lobjWord = CreateObject("Word.Application")
							lobjWord.Documents.Open sFilePath,False,False
						Else
							Set lobjWord = GetObject(,"Word.Application")'This Case May Give Warning But Will not create any issues in Single or Batch Run.
						End If
						lobjWord.Visible = True	
						If Trim(Lcase(sShape)) = "rectangle" Then
							Set objShape = lobjWord.ActiveDocument.Shapes(1)
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),sShape &" shape drawn successfully.")
						End If
						objShape.Select
						objShape.Fill.Visible = msoTrue
					'	lobjWord.Quit
						Fn_MSO_WordShapesOperation = "" & objShape.Fill.ForeColor.RGB
						lobjWord.ActiveDocument.Activate
						lobjWord.ActiveDocument.SaveAs sFilePath
						lobjWord.Quit
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Pass: Fn_MSO_WordShapesOperation successfully executed "+sAction+" action.")
				Case Else 
						Fn_MSO_WordShapesOperation = False			 				
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail: Fn_MSO_WordShapesOperation failed due to Invalid arguments")
				End Select
		Case "AddPicture"   'TC112-2015071500-31_07_2015-VivekA-Porting-Added case from TC1013 to mainline by Madhura 
			'sFilePath is used for Path of Inserting Image
			Set lobjWord = GetObject(,"Word.Application")
			lobjWord.Visible = True
			Call Fn_KeyBoardOperation("SendKeys","%~N~P")
			wait 2
			If Window("MicrosoftWord").Dialog("Insert Picture").Exist(2) = True Then
				If sFilePath<>"" Then
					Window("MicrosoftWord").Dialog("Insert Picture").WinEdit("FileName").Set sFilePath
					wait 1
				End If
				Window("MicrosoftWord").Dialog("Insert Picture").WinObject("Insert").Click
				wait 1
				
				Dialog("ConfirmationBox").SetTOProperty "index","0"
				If Dialog("ConfirmationBox").Exist(1) Then
					Call Fn_UI_WinButton_Click("",Dialog("ConfirmationBox"),"OK","","","")	
					wait 1
					Fn_MSO_WordShapesOperation=False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail: Fn_MSO_WordShapesOperation failed due to Path entered for Inserting picture is invalid")
				End If
				Fn_MSO_WordShapesOperation = True	
			Else
				Fn_MSO_WordShapesOperation=False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail: Fn_MSO_WordShapesOperation failed due to Insert Picture dialog is not available")				
			End If
		Case Else 
				Fn_MSO_WordShapesOperation = False			 				
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail: Fn_MSO_WordShapesOperation failed due to Invalid arguments")
	   End Select	
	   Set lobjWord = Nothing
	   Set objShape = Nothing
End Function
'*********************************************************		Function to handle operations with MS Word.  ********************************************************************
'Function Name		:				Fn_MSO_WordEditOperation  

'Description			 :		 		This function will perform all edit edit operations related to MSWord
'Description			 :		 		This function will perform all edit edit operations related to MSWord

'Parameters			   :	 			1. sAction
' 											   2. sFilePath
'											  3. sString

'Return Value		   : 				True/False  

'Examples				:		 Fn_MSO_WordEditOperation("WordEraseAll","","")	
'										Case "TCWordTabVerify" : Msgbox Fn_MSO_WordEditOperation("TCWordTabVerify" ,"", "Req_TL16_38384")		Added By Ketan on 11-Mar-2011.
'										Case "TCWordTabModify" : Msgbox Fn_MSO_WordEditOperation("TCWordTabModify" ,"", " Modified")		Added By Ketan on 11-Mar-2011.
'										Case "WordGetContent" : Fn_MSO_WordEditOperation("WordGetContent","","")	'Added By Ketan on 6-May-2011.
'										 Case "WordCopyContent"  Fn_MSO_WordEditOperation("WordCopyContent","","")
'										Msgbox Fn_MSO_WordEditOperation("WordGetFontColor" ,"", "")
'										Msgbox Fn_MSO_WordEditOperation("WordFontColorModify" ,"", "blue")
'										Msgbox Fn_MSO_WordEditOperation("WordFontSizeModify" ,"", 22)
'										Msgbox Fn_MSO_WordEditOperation("WordGetFontSize" ,"", "")
'										Msgbox Fn_MSO_WordEditOperation("WordFontModify" ,"", "Arial Black")
'										Msgbox Fn_MSO_WordEditOperation("WordGetFont" ,"", "")
'								bReturn = Fn_MSO_WordEditOperation("ExistImageInWord","","Height:351~Width:468")
'								bReturn = Fn_MSO_WordEditOperation("ExistImageInWord","","")
'								bReturn = Fn_MSO_WordEditOperation("ExistHyperlinkInWord","","000097-Vivek_2")
'								bReturn = Fn_MSO_WordEditOperation("FollowHyperlink","","000097-Vivek_1")
'
'History					 :	
'										Developer Name				Date				Rev. No.			Changes Done																											Tc Release
'									----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Harshal						25/01/2011		       1.0				Created
'									-------------------------------------------------------------------------------------------------
'										Koustubh W					19/05/2011		       1.0				Added case TCWordTabModifyWithoutSave
'									-------------------------------------------------------------------------------------------------
'										Sandeep N					14/10/2011		       1.1				Added case  "WordCopyContent"
'									-------------------------------------------------------------------------------------------------
'										Sandeep N					17/10/2011		       1.2				Added case  "WordModifyWithoutSaveAndClose"      &  "WordSaveAndClose"
'									-------------------------------------------------------------------------------------------------
'										Ketan Raje					18/10/2011		       1.2				Added case  "TCWordFontSizeModify", "TCWordFontSizeModifyWithoutSave","TCWordGetFontSize"
'																																	And 	"WordFontSizeModify", "WordGetFontSize", "WordFontColorModify", "GetWordFontColor"
'																																	And		"TCWordFontColorModify", "TCWordFontColorModifyWithoutSave","TCWordGetFontColor"
'									-------------------------------------------------------------------------------------------------
'										Ketan Raje					18/10/2011		       1.2				Added case  "TCWordFontModify", "TCWordFontModifyWithoutSave","TCWordGetFont"
'																																	And 	"WordFontModify", "WordGetFont"
'                                  --------------------------------------------------------------------------------------------------
 '                                       Vrushali  Wani           31/10/2011                              Added case  "AddTable"
'                                  --------------------------------------------------------------------------------------------------
'                                       Sachin Joshi          22/05/2012                              Modified Case "TCWordTabModify" added wait and Set Word object Visible to True
'								   ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------			
'										Ankit Nigam					22/07/2015			   1.2				Added case  "OpenWord"																										Tc11.2_2015071800											
'	Vivek Ahirrao		23/06/2016			Added new cases "ExistImageInWord", "ExistHyperlinkInWord", "FollowHyperlink"
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------			
'	Vivek Ahirrao		28/06/2016		Added new cases "ExistTableInWord","VerifyTableCellInWord" 
'								Example for same : 
'								Set dicWordDetails = CreateObject("Scripting.Dictionary")
'									dicWordDetails("InstanceOfTable") = 1
'									dicWordDetails("ColumnName") = "Owner"
'									dicWordDetails("ColumnValue") = "AutoTest6 (autotest6)"
'								bReturn = Fn_MSO_WordEditOperation("VerifyTableCellInWord","",dicWordDetails)
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------			
'	Shweta Rathod		29-Aug-2016		Modified Case "WordGetContent", "WordGetContentWithoutClose", "WordGetSomeContents" - modified case, added condition to work with briefcase browser
'**********************************************************************************************************************************************************************************************************************************************************
Public Function Fn_MSO_WordEditOperation(sAction ,sFilePath, sString)'Due to Spelling Mistake while checkin
		Fn_MSO_WordEditOperation = Fn_MSO_WorsEditOperation(sAction ,sFilePath, sString)
End function
Public Function Fn_MSO_WorsEditOperation(sAction ,sFilePath, sString)
	GBL_FAILED_FUNCTION_NAME="Fn_MSO_WorsEditOperation"
	Dim objWord,sWString,WaitCounter, iLen
	Const wdStory = 6
	Const wdMove = 0
	WaitCounter =1
	Set	objWord=nothing
	'Added by Nilesh on 27-Feb-2013 to remove focus from Word document
	If Fn_SISW_UI_Object_Operations("Fn_MSO_WordEditOperation", "Exist", JavaWindow("DefaultWindow"), SISW_MICRO_TIMEOUT)Then
		If Fn_UI_Object_GetROProperty("Fn_MSO_WordEditOperation",JavaWindow("DefaultWindow"), "enabled") Then 
			JavaWindow("DefaultWindow").Minimize
		End If
	Else
		If Window("MicrosoftWord").Exist(5)=True Then
			Window("MicrosoftWord").Minimize
			wait 1
			Window("MicrosoftWord").Maximize
		End If
	End If
	Select Case sAction
		'Case to Verify if a Word document contains Excel Sheet as embedded.
		Case "ExistTableInWord"
			Fn_MSO_WorsEditOperation = False
			Set objWord = Nothing
			If JavaWindow("DefaultWindow").Exist(1) Then
				JavaWindow("DefaultWindow").Maximize	'To remove the Focus From Word File
			End If
			
			Set objWord = GetObject(,"Word.Application")
			objWord.Visible = True
			
			If objWord.ActiveDocument.Tables.Count < 1 Then
				Set objWord = Nothing
				Exit Function
			End If
			Fn_MSO_WorsEditOperation = True
		'Case to Verify Cell in Table contained in Word
		Case "VerifyTableCellInWord"
			Fn_MSO_WorsEditOperation = False
			If varType(sString)<>"9" Then
				Exit Function
			End If
			
			If JavaWindow("DefaultWindow").Exist(1) Then
				JavaWindow("DefaultWindow").Maximize	'To remove the Focus From Word File
			End If
			
			Set objWord = GetObject(,"Word.Application")
			objWord.Visible = True
			
			If objWord.ActiveDocument.Tables.Count < 1 Then
				Set objWord = Nothing
				Exit Function
			End If
			'If there are multiple instances of tables in Word Document
			If sString("InstanceOfTable")<>"" Then
				iInstance = sString("InstanceOfTable")
			Else
				iInstance = 1
			End If
			
			'Take count of total Rows & Columns in Table(iInstance)
			iTotalRow = objWord.ActiveDocument.tables(iInstance).Rows.Count
			iTotalCol = objWord.ActiveDocument.tables(iInstance).Columns.Count
			
			If sString("ColumnName")="" AND sString("ColumnValue")="" Then
				Set objWord = Nothing
				Exit Function
			End If
			For iRowCount = 2 To iTotalRow
				For iColCount = 1 To iTotalCol
					'Get Column name
					sAppColText = Replace(Replace(Trim(objWord.ActiveDocument.Tables(iInstance).Cell(1,iColCount).Range.Text),CHR(13),"*"),"*","")
					'Compare Col value for iColCount column
					If sAppColText=sString("ColumnName") OR sString("ColumnName")=Left(sAppColText ,Len(sAppColText )-1) Then	'As "BEL" char gets added when retrieved from Application
						'Get Row Value for iColCount column
						sAppRowText = Replace(Replace(Trim(objWord.ActiveDocument.Tables(iInstance).Cell(iRowCount,iColCount).Range.Text),CHR(13),"*"),"*","")
						'Compare Col value for iColCount column
						If sString("ColumnValue") = sAppRowText OR sString("ColumnValue") = Left(sAppRowText,Len(sAppRowText)-1) Then	'As "BEL" char gets added when retrieved from Application
							bFlag = True
							Exit For
						End If
						Exit For
					End If
				Next
			Next
			If bFlag = False Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Column Value ["+sString("ColumnValue")+"] for Column ["+sString("ColumnName")+"] does not Exist.")
				Set objWord = Nothing
				Exit Function
			Else
				Set objWord = Nothing
				Fn_MSO_WorsEditOperation = True
			End If
			
		'Case to Verify if a Word document contains Excel Sheet as embedded.
		Case "ExistExcelInWord"
			Fn_MSO_WorsEditOperation = False
			Set objWord = Nothing
			If JavaWindow("DefaultWindow").Exist(1) Then
				JavaWindow("DefaultWindow").Maximize	'To remove the Focus From Word File
			End If
			
			Set objWord = GetObject(,"Word.Application")
			objWord.Visible = True
			
			For iOLE = 1 To objWord.ActiveDocument.InlineShapes.Count 'These are the embedded objects
				If Not objWord.ActiveDocument.InlineShapes(iOLE).OLEFormat Is Nothing Then
					If Instr(objWord.ActiveDocument.InlineShapes(iOLE).OLEFormat.ProgID,"Excel")>0 Then
						Fn_MSO_WorsEditOperation = True
						Exit For
					End If
				End If
			Next
		'Case to click on Hyperlink
		Case "FollowHyperlink"
			Fn_MSO_WorsEditOperation = False
			Set objWord = Nothing
			If JavaWindow("DefaultWindow").Exist(1) Then
				JavaWindow("DefaultWindow").Maximize	'To remove the Focus From Word File
			End If
			
			Set objWord = GetObject(,"Word.Application")
			objWord.Visible = True
			
			For Each wHlink In objWord.ActiveDocument.Hyperlinks
				If sString<>"" Then
					If wHlink.TextToDisplay = sString Then
						wHlink.Follow
						Fn_MSO_WorsEditOperation = True
						Set objWord = Nothing
						Exit For
					End If
				End If
			Next
		'Case to check existence of Hyperlink in Word document
		'sString = "000097-Vivek_1"
		Case "ExistHyperlinkInWord"
			Fn_MSO_WorsEditOperation = False
			Set objWord = Nothing
			If JavaWindow("DefaultWindow").Exist(1) Then
				JavaWindow("DefaultWindow").Maximize	'To remove the Focus From Word File
			End If
			
			Set objWord = GetObject(,"Word.Application")
			objWord.Visible = True
			
			For Each wHlink In objWord.ActiveDocument.Hyperlinks
				If sString<>"" Then
					If wHlink.TextToDisplay = sString Then
						Fn_MSO_WorsEditOperation = True
						Set objWord = Nothing
						Exit For
					End If
				End If
			Next
		'Case to check existence of Image in Word document
		'sString = Height:468~Width:351
		Case "ExistImageInWord"
			Fn_MSO_WorsEditOperation = False
			Set objWord = Nothing
			If JavaWindow("DefaultWindow").Exist(1) Then
				JavaWindow("DefaultWindow").Maximize	'To remove the Focus From Word File
			End If
			
			Set objWord = GetObject(,"Word.Application")
			objWord.Visible = True
			
			iImageCount = objWord.ActiveDocument.InlineShapes.count
			
			If iImageCount<1 Then
				Fn_MSO_WorsEditOperation = False
				Set objWord = Nothing
			Else
				If sString<>"" Then
					For iCount = 1 To iImageCount
						iAppHeight = objWord.ActiveDocument.InlineShapes(iCount).Height
						iAppWidth = objWord.ActiveDocument.InlineShapes(iCount).Width
						
						If sString<>"" Then
							aString = Split(sString,"~")
							aHeight = Split(aString(0),":")
							aWidth = Split(aString(1),":")
							If iAppHeight <> CInt(aHeight(1)) AND iAppWidth <> CInt(aWidth(1)) Then
								Fn_MSO_WorsEditOperation = False
								Set objWord = Nothing
							Else
								Fn_MSO_WorsEditOperation = True
								Set objWord = Nothing
								Exit For
							End If
						End If
					Next
				Else
					Fn_MSO_WorsEditOperation = True
					Set objWord = Nothing
				End If
			End If
			
		Case "WordEraseAll" 'This Case May Give Warning But Will not create any issues in Single or Batch Run.
						Set objWord = Nothing
						JavaWindow("DefaultWindow").Maximize'To remove the Focus From Word File
						Set objWord = GetObject(,"Word.Application")
						objWord.Visible = True
						objWord.Selection.WholeStory
						objWord.Selection.Delete
						objWord.ActiveDocument.Save
						objWord.Quit
						Fn_MSO_WorsEditOperation  =True
						Set objWord = Nothing
       Case "AddTable" 
						Window("MicrosoftWordWin").Maximize
						Call Fn_KeyBoardOperation("SendKeys","%~N~T~I~{ENTER}")
						Set objWord=GetObject(,"Word.application")
								objWord.Visible=true
								objWord.Selection.TypeText(sString)
								objWord.Quit	
								If Window("MicrosoftWordWin").Dialog("Microsoft Office Word").Exist(5) Then
										Window("MicrosoftWordWin").Dialog("Microsoft Office Word").WinButton("Yes").Click
								End If
						Fn_MSO_WorsEditOperation  =True
						Set objWord = Nothing

		Case "WordVerifyForActiveSession"'This Case May Give Warning But Will not create any issues in Single or Batch Run.
					JavaWindow("DefaultWindow").Maximize'To remove the Focus From Word File
					Set objWord = GetObject(,"Word.Application")
					objWord.Visible = True
					objWord.Selection.WholeStory
					Set oSelection = objWord.Selection 
					sWString=  CStr(oSelection.Text)
					'Verify the string
					  If  InStr(1, sWString, sString, 1)>0 Then
							Fn_MSO_WorsEditOperation=True
							'''' TC112-2015071500-29_07_2015-AnkitN-Porting-Handled ConfirmationBox dialog after performing quit Word
							objWord.Quit 0
							If Dialog("ConfirmationBox").Exist(2) then
								If Dialog("ConfirmationBox").WinButton("No").Exist(5) Then
								    Dialog("ConfirmationBox").WinButton("No").Click 10,10	
								End If
							End If
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Pass:Dataset Word type Verified Successfully - Matched String - "&sString)
					Else
							Fn_MSO_WorsEditOperation=False
							objWord.Quit
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"False:Dataset Word Type Verification failed -String "&sString&" does not match")
					End if
					Set objWord = Nothing
		Case "WordGetContent", "WordGetContentWithoutClose", "WordGetSomeContents" 'This Case May Give Warning But Will not create any issues in Single or Batch Run.
					if sString <> "BBVerification" then 'added condition to work with briefcase browser
						JavaWindow("DefaultWindow").Maximize'To remove the Focus From Word File
					else
							'do nothig
					End if
					Set objWord = GetObject(,"Word.Application")
					objWord.Visible = True
					objWord.Selection.WholeStory
					Set oSelection = objWord.Selection 
					sWString=  CStr(oSelection.Text)
					If sAction = "WordGetSomeContents" Then 
						If Instr(1,sWString,chr(13))>0 Then
							sWString = Trim(Replace(Replace(sWString,chr(13),"*"),"*"," "))
						End If
						Fn_MSO_WorsEditOperation=sWString
					Else
						Fn_MSO_WorsEditOperation=sWString
					End If

					If sAction <> "WordGetContentWithoutClose" Then
						objWord.Quit
					End If
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Pass:Content of the Word file extracted successfully.")
					Set objWord = Nothing
		Case "WordIsEmpty"'This Case May Give Warning But Will not create any issues in Single or Batch Run.
					JavaWindow("DefaultWindow").Maximize'To remove the Focus From Word File
					Set objWord = GetObject(,"Word.Application")
					objWord.Visible = True
					iLen = objWord.ActiveDocument.Content.StoryLength
					If Cint(iLen) = 1 Then
						Fn_MSO_WorsEditOperation = True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Pass:WordFile is Empty")
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Pass:WordFile is not Empty")
						Fn_MSO_WorsEditOperation = False
					End If					
					objWord.Quit
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Pass:Content of the Word file extracted successfully.")
					Set objWord = Nothing
		' TC112-2015071500-29_07_2015-VivekA-Porting-Added new case as per design change, as word file contains data with extra charactors and comparing some contents only
		Case "TCWordTabVerify", "TCWordTabVerifySomeContents"	'This Case May Give Warning But Will not create any issues in Single or Batch Run.
					Set objWord = GetObject(,"Word.Application")
					objWord.Selection.WholeStory
					Set oSelection = objWord.Selection 
					sWString = CStr(oSelection.Text)
					
					If sAction = "TCWordTabVerifySomeContents" Then ' TC112-2015071500-29_07_2015-VivekA-Porting-Added new case as per design change, as word file contains data with extra charactors and comparing some contents only
						If Instr(1,sWString,chr(13))>0 Then
							sWString = Trim(Replace(Replace(sWString,chr(13),"*"),"*"," "))
						End If
						If InStr(1,sString, sWString)>0 Then
							Fn_MSO_WorsEditOperation=True
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Pass:Dataset Word type Verified Successfully - Matched String - "&sString)
						Else
							Fn_MSO_WorsEditOperation=False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"False:Dataset Word Type Verification failed -String "&sString&" does not match")
						End If
					Else
						'Verify the string
						'' Chage instr function parameter  by Dipali 
						If InStr(1,sWString, sString)>0 Then
							Fn_MSO_WorsEditOperation=True
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Pass:Dataset Word type Verified Successfully - Matched String - "&sString)
						Else
							Fn_MSO_WorsEditOperation=False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"False:Dataset Word Type Verification failed -String "&sString&" does not match")
						End If
					End If
		Case "WordModifyForActiveSession","WordModifyForActiveSessionExt"
				'To remove the Focus From Word File
				JavaWindow("DefaultWindow").Maximize
				Wait(2)
				Set objWord=GetObject(,"Word.application")
				objWord.Visible=true
				Wait(2)
				'TC11.4_2017081500_NewDevelopment_PoonamC_29Aug2017 : Added Case to select content from Word and modify with new content.
				If sAction = "WordModifyForActiveSessionExt" Then
					objWord.Selection.WholeStory
					Wait(2)
				End If
				objWord.Selection.TypeText(sString)
				Wait(2)
				objWord.ActiveDocument.Save
				Wait(2)
				objWord.Quit
				Fn_MSO_WorsEditOperation=True
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Pass:Sucessfully Modfied MsWord")
		Case "TCWordTabModify", "TCWordTabModifyWithoutSave"
				Set objWord = Nothing
				If JavaWindow("DefaultWindow").Exist(2) Then
					JavaWindow("DefaultWindow").Maximize'To remove the Focus From Word File
				End If
				Set objWord = GetObject(,"Word.Application")
				Set oSelection = objWord.Selection 
				oSelection.EndKey wdStory, wdMove
				oSelection.TypeText(sString)
				Wait(3)
				If sAction <> "TCWordTabModifyWithoutSave" then
					objWord.ActiveDocument.Save
				end If
				Fn_MSO_WorsEditOperation=True
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Pass:Sucessfully Modfied MsWord")

			Case "WordCopyContent"
				Set objWord = GetObject(,"Word.Application")
				objWord.Visible = True
				objWord.Selection.WholeStory
				Set oSelection = objWord.Selection 
				oSelection.Copy
				Fn_MSO_WorsEditOperation=True

			Case "WordModifyWithoutSaveAndClose"
				JavaWindow("DefaultWindow").Maximize
				Wait(2)
				Set objWord=GetObject(,"Word.application")
				objWord.Visible=true
				objWord.Selection.TypeText(sString)
				Fn_MSO_WorsEditOperation=True
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Pass:Sucessfully Modfied MsWord but Not Saved & Closed")

			Case "WordSaveAndClose"
				JavaWindow("DefaultWindow").Maximize
				Wait(2)
				Set objWord=GetObject(,"Word.application")
				objWord.Visible=true
				objWord.ActiveDocument.Save
				objWord.Quit
				Fn_MSO_WorsEditOperation=True
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Pass:Sucessfully Saved and Closed MsWord")

		Case "TCWordFontSizeModify", "TCWordFontSizeModifyWithoutSave","TCWordFontModify", "TCWordFontModifyWithoutSave"
				Set objWord=GetObject(,"Word.application")
				objWord.Selection.WholeStory
				If sAction = "TCWordFontSizeModify" OR sAction = "TCWordFontSizeModifyWithoutSave" Then
					objWord.Selection.Font.Size = Cint(sString)
				ElseIf sAction = "TCWordFontModify" OR sAction = "TCWordFontModifyWithoutSave" Then
					objWord.Selection.Font.Name = sString
				End If			    
				Wait(2)
				If sAction <> "TCWordFontSizeModifyWithoutSave" OR sAction <> "TCWordFontModifyWithoutSave" then
					objWord.ActiveDocument.Save
				end If
				Fn_MSO_WorsEditOperation = True
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Pass:Sucessfully Modfied Font in MSWord")

		Case "TCWordGetFontSize", "TCWordGetFont"
				Set objWord=GetObject(,"Word.application")
				objWord.Selection.WholeStory
				If sAction = "TCWordGetFontSize" Then
					Fn_MSO_WorsEditOperation = objWord.Selection.Font.Size
				ElseIf sAction = "TCWordGetFont" Then
					Fn_MSO_WorsEditOperation = objWord.Selection.Font.Name
				End If
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Pass:Font of text in MS-Word is extracted successfully.")

		Case "WordFontSizeModify", "WordFontModify"
				Set objWord = GetObject(,"Word.Application")
				objWord.Visible = True
				objWord.Selection.WholeStory
				If sAction = "WordFontSizeModify" Then
					objWord.Selection.Font.Size = Cint(sString)
				ElseIf sAction = "WordFontModify" Then
					objWord.Selection.Font.Name = sString
				End If
				objWord.ActiveDocument.Save
				objWord.Quit
				Fn_MSO_WorsEditOperation = True
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Pass:Sucessfully Modfied Font Size in MSWord")
				
		Case "WordGetFontSize", "WordGetFont"
				Set objWord = GetObject(,"Word.Application")
				objWord.Visible = True
				objWord.Selection.WholeStory
				If sAction = "WordGetFontSize" Then
					Fn_MSO_WorsEditOperation = objWord.Selection.Font.Size
				ElseIf sAction = "WordGetFont" Then
					Fn_MSO_WorsEditOperation = objWord.Selection.Font.Name
				End If
				objWord.ActiveDocument.Save
				objWord.Quit				
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Pass:FontSize of text in MS-Word is extracted successfully.")

		Case "WordFontColorModify"
				Select Case Trim(Lcase(sString))
								Case "white"
									sString = 8
								Case "yellow"
									sString = 7
								Case "red"
									sString = 6
								Case "pink"
									sString = 5
								Case "green"
									sString = 4
								Case "blue"
									sString = 2
								Case "aqua"
									sString = 3
								Case "black"
									sString = 1
				End Select
				Set objWord = GetObject(,"Word.Application")
				objWord.Visible = True
				objWord.Selection.WholeStory
				objWord.Selection.Font.ColorIndex = Cint(sString)
				objWord.ActiveDocument.Save
				objWord.Quit	
				Fn_MSO_WorsEditOperation = True			
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Pass:Color of text in MS-Word is modified successfully.")

		Case "WordGetFontColor"
				Set objWord = GetObject(,"Word.Application")
				objWord.Visible = True
				objWord.Selection.WholeStory
				sString = objWord.Selection.Font.ColorIndex
				Select Case sString
								Case "8"
									sString = "white"
								Case "7"
									sString = "yellow"
								Case "6"
									sString = "red"
								Case "5"
									sString = "pink"
								Case "4"
									sString = "green"
								Case "2"
									sString = "blue"
								Case "3"
									sString = "aqua"
								Case "1"
									sString = "black"
								Case Else
									sString = "Unknown"
				End Select
				objWord.ActiveDocument.Save
				objWord.Quit		
				Fn_MSO_WorsEditOperation = sString					
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Pass:Color of text in MS-Word is modified successfully.")

		Case "TCWordFontColorModify", "TCWordFontColorModifyWithoutSave"
				Select Case Trim(Lcase(sString))
								Case "white"
									sString = 8
								Case "yellow"
									sString = 7
								Case "red"
									sString = 6
								Case "pink"
									sString = 5
								Case "green"
									sString = 4
								Case "blue"
									sString = 2
								Case "aqua"
									sString = 3
								Case "black"
									sString = 1
				End Select
				Set objWord=GetObject(,"Word.application")
				objWord.Selection.WholeStory
			    objWord.Selection.Font.ColorIndex = Cint(sString)
				Wait(2)
				If sAction <> "TCWordColorModifyWithoutSave" then
					objWord.ActiveDocument.Save
				end If
				Fn_MSO_WorsEditOperation = True
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Pass:Sucessfully Modfied Color in MSWord")

		Case "TCWordGetFontColor"
				Set objWord=GetObject(,"Word.application")
				objWord.Selection.WholeStory			    
				sString = objWord.Selection.Font.ColorIndex
				Select Case sString
								Case "8"
									sString = "white"
								Case "7"
									sString = "yellow"
								Case "6"
									sString = "red"
								Case "5"
									sString = "pink"
								Case "4"
									sString = "green"
								Case "2"
									sString = "blue"
								Case "3"
									sString = "aqua"
								Case "1"
									sString = "black"
								Case Else
									sString = "Unknown"
				End Select
				Fn_MSO_WorsEditOperation = sString					
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Pass:Color of text in MS-Word is extracted successfully.")
		Case "WordVerifyForIEActiveSession"'Added by Nilesh on 26-March-2013
					If JavaWindow("DefaultWindow").Exist(5)=True Then
						JavaWindow("DefaultWindow").Maximize  'To remove the Focus From Word File
					End If
					Set objWord = GetObject( ,"Word.Application")
					objWord.Selection.WholeStory
					Set oSelection = objWord.Selection 
					sWString=  CStr(oSelection.Text)
					'Verify the string
					  If  InStr(1, sWString, sString, 1)>0 Then
							Fn_MSO_WorsEditOperation=True
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Pass:Dataset Word type Verified Successfully - Matched String - "&sString)
					Else
							Fn_MSO_WorsEditOperation=False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"False:Dataset Word Type Verification failed -String "&sString&" does not match")
					End if
					Set objWord = Nothing
		Case "OpenWord"			' TC112-2015070800-22_07_2015-Porting-AnkitN-Added case to Open Instance of MsWord .
					If Fn_UI_ObjectExist("Fn_MSO_WorsEditOperation", JavaWindow("DefaultWindow")) Then
						JavaWindow("DefaultWindow").Maximize'To remove the Focus From Excel File
					End If
					If sFilePath = "" Then 
						Fn_MSO_WorsEditOperation = False
						Exit Function
					Else
						Set wrdApp = CreateObject("Word.Application")
						wrdApp.DisplayAlerts=0
						wrdApp.Documents.Open sFilePath,False,True
					End If 			
					'SystemUtil.Run sFilePath
					wait 10
					Fn_MSO_WorsEditOperation = True
			'Example : bReturn=Fn_MSO_WordEditOperation("OpenWordDefault","","")
			Case "OpenWordDefault"
					On error resume next
					'Get existing instance of Word if it exists.
				   	Set wrdApp = GetObject(, "Word.Application")
					
				   	If Err.Number <> 0 then
				    	'If GetObject fails, then use CreateObject instead.
				      	Set wrdApp = CreateObject("Word.Application")
				      	wait 10
				   	End If
				   	wrdApp.Visible = True
					wrdApp.DisplayAlerts=0
				   	'Add a new document.
				   	wrdApp.Documents.Add
				   	
				   	Set wrdApp = Nothing
					Fn_MSO_WorsEditOperation = True
    End Select
   If JavaWindow("DefaultWindow").Exist(5)=True  then
	  if Fn_UI_Object_GetROProperty("Fn_MSO_WordEditOperation",JavaWindow("DefaultWindow"), "enabled") Then
         JavaWindow("DefaultWindow").Maximize
      End if 
   End If
	Set	objWord=nothing
End function
'*********************************************************		Function to handle operations with MS Excel.  ********************************************************************
'Function Name		:				Fn_MSO_ExcelEditOperations  

'Description			 :		 		This function will perform all edit edit operations related to MSWord

'Parameters			   :	 			sAction,sFilePath,iWorksheets,sCellPosition,sString,bCloseExcel

'Return Value		   : 				True/False  

'Examples				:		 Msgbox Fn_MSO_ExcelEditOperations("ChangeDateFormat","",1,"A1","mm-dd-yy","")
								'Msgbox Fn_MSO_ExcelEditOperations("ModifyCellValue","D:\Mainline\Test.xls",1,"C3","Ketan","Yes")
								'Msgbox Fn_MSO_ExcelEditOperations("GetCellPosition","","Sheet1","","Group","")
								'Msgbox Fn_MSO_ExcelEditOperations("GetCellPosition","","Sheet1","","Group @2","")
								'Msgbox Fn_MSO_ExcelEditOperations("GetColumnData","","Sheet1","2-5","Group","")
								'Msgbox Fn_MSO_ExcelEditOperations("GetColumnData","","Sheet1","5","Group","")
								'Msgbox Fn_MSO_ExcelEditOperations("GetRowData","","Sheet1","2-5","Group","")
								'Msgbox Fn_MSO_ExcelEditOperations("GetRowData","","Sheet1","5","Group","")
								'Msgbox Fn_MSO_ExcelEditOperations("VerifySaveAsDialog","","","","","")
								'Msgbox Fn_MSO_ExcelEditOperations("VerifyCellBackground","","","A1","BLACK","")
								'Msgbox Fn_MSO_ExcelEditOperations("OpenExcel","D:\Mainline\Test.xls",1,"","","")
								'Msgbox Fn_MSO_ExcelEditOperations("OpenLiveExcel","D:\Mainline\Test.xls","","","","")
								'Msgbox Fn_MSO_ExcelEditOperations("OpenLiveExcel","D:\Mainline\Test.xls","Sheet1","A1","","")
								'Msgbox Fn_MSO_ExcelEditOperations("SelectCell","",1,"C1:C10,G1,G3:G10","","")
								'Msgbox Fn_MSO_ExcelEditOperations("SelectCell","",1,"C1:C10","","")
								'Msgbox Fn_MSO_ExcelEditOperations("SelectCell","",1,"C10","","")
								'Msgbox Fn_MSO_ExcelEditOperations("VerifyCellFontSize","",1,"G1",11,"")
								'Msgbox Fn_MSO_ExcelEditOperations("VerifyCellFontName","",1,"G1","Calibri","")
								'Msgbox Fn_MSO_ExcelEditOperations("VerifyCellFontBold","",1,"G1","","")
								'Msgbox Fn_MSO_ExcelEditOperations("VerifyCellFontItalic","",1,"G1","","")
								'Msgbox Fn_MSO_ExcelEditOperations("VerifyCellFontRegular","",1,"G1","","")
								'Msgbox Fn_MSO_ExcelEditOperations("GetCellFontColour","",1,"G1","","")
								'Msgbox Fn_MSO_ExcelEditOperations("InsertRow","",1,"G1","","")
								'Msgbox Fn_MSO_ExcelEditOperations("InsertColumn","",1,"G1","","")
								'Msgbox Fn_MSO_ExcelEditOperations("ShiftCellsRight","",1,"G1","","")
								'Msgbox Fn_MSO_ExcelEditOperations("ShiftCellsDown","",1,"G1","","")
								'Msgbox Fn_MSO_ExcelEditOperations("DeleteEntireColumn","",1,"G1","","")
								'Msgbox Fn_MSO_ExcelEditOperations("DeleteEntireRow","",1,"G1","","")
								'Msgbox Fn_MSO_ExcelEditOperations("ShiftCellsLeft","",1,"G1","","")
								'Msgbox Fn_MSO_ExcelEditOperations("ShiftCellsUp","",1,"G1","","")
								'Msgbox Fn_MSO_ExcelEditOperations("GroupCells","",1,"G1:G10","Columns","")
								'Msgbox Fn_MSO_ExcelEditOperations("GetCellComment","",1,"G1","","")
								'Msgbox Fn_MSO_ExcelEditOperations("GetCellValue","",1,"G1","","")
								'Msgbox Fn_MSO_ExcelEditOperations("ActivateExcel","",1,"","","")
								'Msgbox Fn_MSO_ExcelEditOperations("FollowHyperlink","",1,"G1","","")
								'Msgbox Fn_MSO_ExcelEditOperations("HideRow","",1,"F3","","")
								'Msgbox Fn_MSO_ExcelEditOperations("UnhideRow","",1,"F3","","")
								'Msgbox Fn_MSO_ExcelEditOperations("HideColumn","",1,"F3","","")
								'Msgbox Fn_MSO_ExcelEditOperations("UnhideColumn","",1,"F3","","")
								'Msgbox Fn_MSO_ExcelEditOperations("GetSelectedCellNumber","",1,"","","")
								'Msgbox Fn_MSO_ExcelEditOperations("GetRowContents","","Sheet1",1,"","")
								'bReturn=Fn_MSO_ExcelEditOperations("GetSheetName","","",1,"","")
								'bReturn=Fn_MSO_ExcelEditOperations("ModifySheetName","","",1,"Shreyas
								'bReturn=Fn_MSO_ExcelEditOperations("GetLastRowIndex","","",1,"","")
								'bReturn=Fn_MSO_ExcelEditOperations("CutPasteCol","","",1,"Spec:PSOccurrence_struct","") 
								'bReturn= Fn_MSO_ExcelEditOperations("UnGroupCells","",1,"A1:A2","Rows","")
								'bReturn= Fn_MSO_ExcelEditOperations("GetMultiCellValues","",1,"A1:A2:A3","","")
								'bReturn= Fn_MSO_ExcelEditOperations("VerifyMultiCellValues","","Complying","A2:A3:A4:A5","req2~req3~req4~req5","")
								'bReturn = Fn_MSO_ExcelEditOperations("TCExcelTabExists","","","","","")
								'bReturn = Fn_MSO_ExcelEditOperations("TCExcelTabCellVerify","","","B2","Test","")
								'bReturn = Fn_MSO_ExcelEditOperations("ExistsHyperlinkInCell","","","$E$5","000071-Vivek","")
								'bReturn = Fn_MSO_ExcelEditOperations("ExistsImageInCell","","","$C$4","","")
'		Note:	to work on hidden / embedded excel pass "~HiddenExcel" with sAction

								'Msgbox Fn_MSO_ExcelEditOperations("VerifyCellFontSize~HiddenExcel","",1,"G1",11,"")
								'Msgbox Fn_MSO_ExcelEditOperations("VerifyCellFontName~HiddenExcel","",1,"G1","Calibri","")
								'Msgbox Fn_MSO_ExcelEditOperations("VerifyCellFontBold~HiddenExcel","",1,"G1","","")
								'Msgbox Fn_MSO_ExcelEditOperations("VerifyCellFontItalic~HiddenExcel","",1,"G1","","")
								'Msgbox Fn_MSO_ExcelEditOperations("VerifyCellFontRegular~HiddenExcel","",1,"G1","","")

'History					 :		
'		Developer Name				Date				Rev. No.			Changes Done	
'----------------------------------------------------------------------------------------------------------------
'		Harshal						08/03/2011				1.0				Created
'----------------------------------------------------------------------------------------------------------------
'		Koustubh					19/03/2011				1.0				Added cases GetCellPosition, GetColumnData
'----------------------------------------------------------------------------------------------------------------
'		Koustubh					19/03/2011				1.0				Added case VerifySaveAsDialog
'----------------------------------------------------------------------------------------------------------------
'		Koustubh					19/03/2011				1.0				Added case VerifyCellBackground
'----------------------------------------------------------------------------------------------------------------
'		Koustubh					29/05/2011				1.0				Added case OpenExcel
'----------------------------------------------------------------------------------------------------------------
'		Koustubh					03/06/2011				1.0				Added case OpenLiveExcel
'																			SelectCell, VerifyCellFontSize, VerifyCellFontName
'																			VerifyCellFontBold, VerifyCellFontItalic, VerifyCellFontRegular
'----------------------------------------------------------------------------------------------------------------
'		Koustubh					29/05/2011				1.0				Added case GroupCells, GetCellComment
'----------------------------------------------------------------------------------------------------------------
'		Koustubh					13/06/2011				1.0				Added case InsertRow, InsertColumn, ShiftCellsRight, ShiftCellsDown
'----------------------------------------------------------------------------------------------------------------
'		Amit T						14/06/2011				1.0				Added case GetCellValue
'----------------------------------------------------------------------------------------------------------------
'		Koustubh					15/06/2011				1.0				Added cases "DeleteEntireRow", "DeleteEntireColumn", "ShiftCellsLeft", "ShiftCellsUp"
'----------------------------------------------------------------------------------------------------------------
'		Koustubh					16/06/2011				1.0				Added cases "ActivateExcel"
'----------------------------------------------------------------------------------------------------------------
'		Koustubh					16/06/2011				1.0				Modified case "GetCellPosition"
'----------------------------------------------------------------------------------------------------------------
'		Koustubh					23/06/2011				1.0				Added case "FollowHyperlink"
'----------------------------------------------------------------------------------------------------------------
'		Koustubh					23/06/2011				1.0				Added cases  "HideRow", "HideColumn","UnhideRow","UnhideColumn"
'----------------------------------------------------------------------------------------------------------------
'		Harshal A					27/06/2011				1.0				Added cases  "GetCellFontName","GetCellFontSize"
'----------------------------------------------------------------------------------------------------------------
'		Ketan Raje					18/10/2011				1.0				Added cases  "GetSelectedCellNumber","GetRowContents"
'----------------------------------------------------------------------------------------------------------------
'Shreyas                            27-02-2012				 1.1			Added Cases "GetSheetName"  &  "ModifySheetName"
'----------------------------------------------------------------------------------------------------------------
'Shreyas                            05-03-2012				 1.1			Added Case "GetLastRowIndex"
'----------------------------------------------------------------------------------------------------------------
'Shreyas                            12-03-2012				 1.1			Added Case "CutPasteCol"
'----------------------------------------------------------------------------------------------------------------
'Shreyas                            14-03-2012				 1.1			Added Case "UngroupCells"
'----------------------------------------------------------------------------------------------------------------
'Pranav Ingle                   31-07-2012				 1.1			Added Case "GetMultiCellValues"
'----------------------------------------------------------------------------------------------------------------
'Madhura Puranik                02-09-2015				 1.1			Added Case "VerifyMultiCellValues"			Ankit Nigam		Tc10.1.5_2015081100
'----------------------------------------------------------------------------------------------------------------
'Vivek Ahirrao					23-06-2016					1.2		Added new cases "ExistsHyperlinkInCell","ExistsImageInCell"
'----------------------------------------------------------------------------------------------------------------
'Shweta Rathod					29-Aug-2016			1.1     			Added new "VerifyCellValue_BB"	
'----------------------------------------------------------------------------------------------------------------
Function Fn_MSO_ExcelEditOperations(sAction,sFilePath,iWorksheets,sCellPosition,sString,bCloseExcel)
	GBL_FAILED_FUNCTION_NAME="Fn_MSO_ExcelEditOperations"
	Dim objExcel,sAppValue, objShell, arrAction
	Dim iExcel,sRowCount, bVisibleFlag, arrValues, arrString
	Dim iCellCnt, iColCnt, bFlag, xlsCellPosition, sReturnData, aRange, iCnt
	Dim sColour, WshShell, sKey
	Dim iInstance, aCol,sValue
	Dim X,Y,left,top,right,bottom
	Dim oSheet,sCol,sColname,shell,bReturn,iTimeout,iCount
	Dim iHeight, iWidth
	bVisibleFlag = True
	On Error Resume Next
	arrAction = split(sAction,"~")
	If UBound(arrAction) > 0  Then
		If arrAction(1) = "HiddenExcel" Then
			bVisibleFlag = False
		End If
	End If
	sAction = arrAction(0)
	sReturnData = ""
	Set objExcel = Nothing
	Fn_MSO_ExcelEditOperations = False
    ' Added by Nilesh on 1 June 2012 to handle Restart application dialog
    If Dialog("ExceClose").Exist(1) Then
		Dialog("ExceClose").Type "N"
	End If
	'End 

	Select Case sAction
		Case "TCExcelTabSelectCell"
		 	'To set focus on Excel tab in Teamcenter RAC
			If JavaWindow("DefaultWindow").JavaObject("OleClientSite").Exist Then
				iHeight = JavaWindow("DefaultWindow").JavaObject("OleClientSite").GetROProperty("height")
				iWidth = JavaWindow("DefaultWindow").JavaObject("OleClientSite").GetROProperty("width")
				JavaWindow("DefaultWindow").JavaObject("OleClientSite").Click iWidth/2, iHeight/2,"LEFT"
				Wait 0,500
			End If
			Set objExcel = GetObject(,"Excel.Application")
			If iWorksheets <> "" Then
				objExcel.Worksheets(iWorksheets).Activate
				Wait 0,500
			End If	
			objExcel.Range(sCellPosition).Select
			Wait 0,500
			Fn_MSO_ExcelEditOperations = True
		Case "TCExcelTabModifyCellValue"
			'To set focus on Excel tab in Teamcenter RAC
			If JavaWindow("DefaultWindow").JavaObject("OleClientSite").Exist Then
				iHeight = JavaWindow("DefaultWindow").JavaObject("OleClientSite").GetROProperty("height")
				iWidth = JavaWindow("DefaultWindow").JavaObject("OleClientSite").GetROProperty("width")
				JavaWindow("DefaultWindow").JavaObject("OleClientSite").Click iWidth/2, iHeight/2,"LEFT"
				Wait 0,500
			End If
			Set objExcel = GetObject(,"Excel.Application")
			If iWorksheets <> "" Then
				objExcel.Worksheets(iWorksheets).Activate
				objExcel.Worksheets(iWorksheets).Range(sCellPosition).value = sString
				Wait 0,500
				Fn_MSO_ExcelEditOperations = Fn_KeyBoardOperation("SendKeys", "^S")
				Wait 0,500
			End If
		'Case to verify Cell value in Excel Tab in Teamcenter RAC
		Case "TCExcelTabCellVerify"
			'To set focus on Excel tab in Teamcenter RAC
			If JavaWindow("DefaultWindow").JavaObject("OleClientSite").Exist Then
				iHeight = JavaWindow("DefaultWindow").JavaObject("OleClientSite").GetROProperty("height")
				iWidth = JavaWindow("DefaultWindow").JavaObject("OleClientSite").GetROProperty("width")
				JavaWindow("DefaultWindow").JavaObject("OleClientSite").Click iWidth/2, iHeight/2,"LEFT"
			End If
			Set objExcel = GetObject(,"Excel.Application")
			sAppValue = objExcel.Range(sCellPosition).value
			If sAppValue = sString Then
				Fn_MSO_ExcelEditOperations = True
			End If
		'Case to check existance of Excel in Teamcenter RAC Tab
		Case "TCExcelTabExists"
			Wait 1
			Fn_MSO_ExcelEditOperations = JavaWindow("DefaultWindow").JavaObject("OleClientSite").Exist
		Case"ChangeDateFormat"
				If Fn_UI_ObjectExist("Fn_MSO_ExcelEditOperations", JavaWindow("DefaultWindow")) Then
					JavaWindow("DefaultWindow").Maximize'To remove the Focus From Excel File
				End If
				Set objExcel = GetObject(,"Excel.Application")
				objExcel.Visible = bVisibleFlag
				objExcel.Worksheets(iWorksheets).Activate
				objExcel.Range(sCellPosition).Select
				objExcel.Selection.NumberFormat = "[$-409]"+sString+";@"
				objExcel.ActiveWorkbook.Save
				Fn_MSO_ExcelEditOperations = True
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case"OpenExcel"
				If Fn_UI_ObjectExist("Fn_MSO_ExcelEditOperations", JavaWindow("DefaultWindow")) Then
					JavaWindow("DefaultWindow").Maximize'To remove the Focus From Excel File
				End If
				
'				Set objShell = CreateObject("WScript.Shell")
'				objShell.Run sFilePath
				If sFilePath = "" Then sFilePath = "EXCEL.exe"
				SystemUtil.Run sFilePath
				wait 10
				Set objExcel = GetObject(,"Excel.Application")
				objExcel.Visible = bVisibleFlag
				If iWorksheets <> "" Then
					objExcel.Worksheets(iWorksheets).Activate
				End If
				Fn_MSO_ExcelEditOperations = True
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case"ActivateExcel"
				Set objExcel = GetObject(,"Excel.Application")
				objExcel.Visible = bVisibleFlag
				If iWorksheets <> "" Then
					objExcel.Worksheets(iWorksheets).Activate
				End If
				Window("MicrosoftExcel").Maximize
				Fn_MSO_ExcelEditOperations = True
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "OpenLiveExcel"
				If Fn_UI_ObjectExist("Fn_MSO_ExcelEditOperations",Window("MicrosoftExcel").Window("FileOpen")) = False then
						Call Fn_MenuOperation("Select", "Tools:Open Live Excel")
				End If
				If Fn_UI_ObjectExist("Fn_MSO_ExcelEditOperations",Window("MicrosoftExcel").Window("FileOpen"))  then
					'Window("MicrosoftExcel").Window("FileOpen").WinObject("FileName").Type sFilePath
					Window("MicrosoftExcel").Window("FileOpen").WinEdit("FileName").Type sFilePath
					Set WshShell = CreateObject("WScript.Shell")
					WshShell.SendKeys "{ENTER}"
					Set WshShell = nothing
				End If
				Fn_MSO_ExcelEditOperations = True
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "VerifyCellValue"
				If Fn_UI_ObjectExist("Fn_MSO_ExcelEditOperations", JavaWindow("DefaultWindow")) Then
					JavaWindow("DefaultWindow").Minimize
					JavaWindow("DefaultWindow").Maximize	'To remove the Focus From Excel File
					'JavaWindow("DefaultWindow").Click 1, 1,"LEFT"
				End If
				Set objExcel = GetObject(,"Excel.Application")
				objExcel.Visible = bVisibleFlag
				If iWorksheets <> "" Then
					objExcel.Worksheets(iWorksheets).Activate
				End If
				sAppValue = objExcel.Range(sCellPosition).text
				If sAppValue = sString Then
					Fn_MSO_ExcelEditOperations = True
				End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -		
		'[Tc11.2.2-2016033000-02_09_2015-MadhuraP-Added case to verify multiple values from excel]
		Case "VerifyMultiCellValues"
				bFlag = False
		        Set objExcel = GetObject(,"Excel.Application")
				objExcel.Visible = bVisibleFlag
				If iWorksheets = "" Then
					iWorksheets = 1
				End If
				objExcel.Worksheets(iWorksheets).Activate
				sValue=""
				aRange = Split(sCellPosition,":")
				' Cell Value here
                For iCount = 0 To Ubound(aRange)
					If iCount = 0 Then
						sValue=objExcel.Range(aRange(iCount)).text
					Else
						sValue = sValue +"~"+ objExcel.Range(aRange(iCount)).text
					End If
				Next
				arrValues = Split(sValue,"~")
				arrString = Split(sString,"~")
				For iCount = 0 to UBound(arrString)
					For iCnt = 0 To UBound(arrValues)
						If trim(arrString(iCount)) = trim(arrValues(iCnt)) Then
							bFlag = True
							Exit For
						End If
					Next
					If bFlag <> True Then
						Exit For
					End If
				Next
				If bFlag = True Then
					Fn_MSO_ExcelEditOperations = True
				End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "ModifyCellValue","ModifyCellValueAndSave"
				If Window("MicrosoftExcel").Exist(5)=True Then
					Window("MicrosoftExcel").Minimize
					wait 1
					Window("MicrosoftExcel").Maximize
					wait 3
				End If
				Set objExcel = GetObject(,"Excel.Application")
				 objExcel.Visible = True
				
				If iWorksheets <> "" Then
					objExcel.Worksheets(iWorksheets).Activate
					Wait 2
					objExcel.Worksheets(iWorksheets).Range(sCellPosition).value = sString
					If sAction = "ModifyCellValueAndSave" Then
						objExcel.ActiveWorkbook.Save
					Else	
						objExcel.Worksheets(iWorksheets).SaveAs sFilePath
					End If
					Fn_MSO_ExcelEditOperations = True
				End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		 Case "ModifyCellValue_WithOutSave"
				Set objExcel = GetObject(,"Excel.Application")
				'objExcel.Visible = True
				objExcel.Worksheets(iWorksheets).Activate
				objExcel.Worksheets(iWorksheets).Range(sCellPosition).value = sString
				wait(2)
				Fn_MSO_ExcelEditOperations = True
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		 Case "SelectCell"
				if JavaWindow("DefaultWindow").Exist(1) Then
	            	JavaWindow("DefaultWindow").Minimize
	           		JavaWindow("DefaultWindow").Maximize
	            End If
	            Set objExcel = GetObject(,"Excel.Application")
				objExcel.Visible = bVisibleFlag


				If iWorksheets <> "" Then
					objExcel.Worksheets(iWorksheets).Activate
				End If	
				if Window("MicrosoftExcel").Exist(1) Then
					Window("MicrosoftExcel").Minimize
					Window("MicrosoftExcel").Maximize
				End if
				objExcel.Range(sCellPosition).Select
				Fn_MSO_ExcelEditOperations = True
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		 Case "VerifyCellFontSize", "VerifyCellFontName", "VerifyCellFontRegular", "VerifyCellFontBold","VerifyCellFontItalic","GetCellFontColour", "GetCellFontColor","GetCellFontName","GetCellFontSize"
				Set objExcel = GetObject(,"Excel.Application")
				objExcel.Visible = bVisibleFlag
				If iWorksheets <> "" Then
					objExcel.Worksheets(iWorksheets).Activate
				End If
				Select Case sAction
					Case "GetCellFontSize"
						Fn_MSO_ExcelEditOperations = Cstr(objExcel.Range(sCellPosition).Font.Size)
					Case "VerifyCellFontSize"
						If cInt(objExcel.Range(sCellPosition).Font.Size) = cInt(sString) then
							Fn_MSO_ExcelEditOperations = True
						End If
					Case "GetCellFontName"
						Fn_MSO_ExcelEditOperations = uCase(trim(objExcel.Range(sCellPosition).Font.Name))
					Case "VerifyCellFontName"
						If uCase(trim(objExcel.Range(sCellPosition).Font.Name)) = uCase(sString) then
							Fn_MSO_ExcelEditOperations = True
						End IF
					Case "VerifyCellFontBold"
						Fn_MSO_ExcelEditOperations = objExcel.Range(sCellPosition).Font.BOLD
					Case "VerifyCellFontItalic"
						Fn_MSO_ExcelEditOperations = objExcel.Range(sCellPosition).Font.ITALIC
					Case "VerifyCellFontRegular"
						If objExcel.Range(sCellPosition).Font.ITALIC = False AND objExcel.Range(sCellPosition).Font.BOLD = False Then
							Fn_MSO_ExcelEditOperations = True
						End If
					Case "GetCellFontColour", "GetCellFontColor"
						Select Case objExcel.Range(sCellPosition).Font.ColorIndex
							Case 1 'BLACK
								sColour = "BLACK"
							case 2 'WHITE
								sColour = "WHITE"
							Case 3 'RED
								sColour = "RED"
							Case 4 'GREEN
								sColour = "GREEN"
							Case 5 'BLUE
								sColour = "BLUE"
							Case 6, 27 ' YELLOW
								sColour = "YELLOW"
							Case 7, 26 'MAGENTA
								sColour = "MAGENTA"
							Case 8 'CYAN
								sColour = "CYAN"
							Case 46
								sColour = "ORANGE"
							Case 47
								sColour = "PURPLE"
							Case Else
								sColour = objExcel.Range(sCellPosition).Font.ColorIndex
						end Select
						Fn_MSO_ExcelEditOperations = sColour
				End Select
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "AddNewSheet"
			Set objExcel = GetObject(,"Excel.Application")
			If iWorksheets <> "" Then
				objExcel.Worksheets(iWorksheets).Activate
			End If
			objExcel.Sheets.Add
			'objExcel.Visible = True
			Fn_MSO_ExcelEditOperations = True
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "GetCellPosition"
				Set objExcel = GetObject(,"Excel.Application")
				'objExcel.Visible = True
				If iWorksheets = "" Then
					iWorksheets = 1
				End If
				objExcel.Worksheets(iWorksheets).Activate
				bFlag = False
				If sString <> "" Then
					aCol = split(sString,"@")
					aCol(0) = trim(aCol(0))
					If UBound(aCol) = 1 then
						iInstance = CInt(aCol(1))
					Else
						iInstance = 1
					End If
					
					For iCellCnt = 1 to objExcel.Worksheets(iWorksheets).UsedRange.Rows.Count
						For iColCnt = 1 to objExcel.Worksheets(iWorksheets).UsedRange.Columns.Count
							If objExcel.Worksheets(iWorksheets).cells(iCellCnt, iColCnt).value = aCol(0) then
									iInstance = iInstance -1
									If iInstance = 0 Then
										' if value matches exit from inner loop
										bFlag = true
										Exit For
									End If
							end if
						Next
						If bFlag Then
							' if value matches exit from outer loop
							Exit For
						End If
					Next
				End If
				If bFlag Then
					Fn_MSO_ExcelEditOperations = (Fn_MSO_ExcelColHeader("GetColumnHeaderName", iColCnt) & "" &iCellCnt)
				else
					Fn_MSO_ExcelEditOperations = False
				End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "GetColumnData","GetColumnDataExt"
			'Added by Nilesh to handle Synchronization after Excel export
				If Fn_UI_ObjectExist("Fn_MSO_ExcelEditOperations", JavaWindow("DefaultWindow")) Then
				     JavaWindow("DefaultWindow").Minimize
					JavaWindow("DefaultWindow").Maximize'To remove the Focus From Excel File Added by Nilesh on 5-Nov-2012
				End If
	            iTimeout=120
			   bFlag=False
				Set objExcel = GetObject(,"Excel.Application")
				For iCount= 0 To iTimeout
						If objExcel.Visible= True Then
							bFlag=True
							Exit For
						End If
				Next
				'End 

				If iWorksheets = "" Then
					iWorksheets = 1
				End If
				Window("MicrosoftExcel").Maximize
				objExcel.Worksheets(iWorksheets).Activate
				'objExcel.Visible = True
				xlsCellPosition =  Fn_MSO_ExcelEditOperations("GetCellPosition","",iWorksheets,"", sString ,"false")
				If xlsCellPosition = False Then
						' column not found.
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_ExcelEditOperations:Specified column [ " & sString & " ] is not present in given excel sheet.")
					Exit function
				End If
				'For iCnt = 1 to len(xlsCellPosition)
				'	sChar = mid(xlsCellPosition,iCnt,1)
				'	If  Asc(sChar) >= Asc("A") and   Asc(sChar) <=  Asc("Z") then
				'		' do nothing :)
				'	else
				'		' exit from loop
				'		Exit For
				'	end if
				'Next
				' row number	
				'iCellCnt = cInt( mid(xlsCellPosition,iCnt, len(xlsCellPosition)))
				' column number
				'iColCnt =  Fn_MSO_ExcelColHeader("GetColumnHeaderNumber", mid(xlsCellPosition,1, iCnt -1)) 
				
				' row number	
				iCellCnt = Fn_MSO_ExtractExcelRowNumber(xlsCellPosition)
				' column number
				iColCnt = Fn_MSO_ExcelColHeader("GetColumnHeaderNumber", Fn_MSO_ExtractExcelColumnName(xlsCellPosition))
				
				aRange = split(sCellPosition,"-")
				Select Case UBound(aRange)
					Case 0
						sReturnData = objExcel.Worksheets(iWorksheets).cells((iCellCnt + cInt(aRange(0))), iColCnt).value
					Case 1
						For iCnt = (iCellCnt + cInt(aRange(0))) to (iCellCnt + cInt(aRange(1)))
								If sReturnData = "" Then
									sReturnData = objExcel.Worksheets(iWorksheets).cells(iCnt, iColCnt).value
								else
									sReturnData = sReturnData & "~" & objExcel.Worksheets(iWorksheets).cells(iCnt, iColCnt).value
								End If
						Next
				End Select
				
				If sAction = "GetColumnDataExt" Then ' Case to return value if Cell is Empty - Tc11.5_20180616b_NewDevelopment_PoonamC_04Jul2018
					Fn_MSO_ExcelEditOperations = sReturnData
				Else
					If sReturnData <> "" Then
						Fn_MSO_ExcelEditOperations = sReturnData
					End If
				End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "VerifyCellValue_BB"    'added case to work with briefcase browser [Briefcase Browser 11.2.3 (P11000.2.0.30_20160608.00) 64-bit]
				If Window("MicrosoftExcel").Exist(5)=True Then
					Window("MicrosoftExcel").Minimize
					wait 1
					Window("MicrosoftExcel").Maximize
				End If
				Set objExcel = GetObject(,"Excel.Application")
				objExcel.Visible = bVisibleFlag
				If iWorksheets <> "" Then
					objExcel.Worksheets(iWorksheets).Activate
				End If
				sAppValue = objExcel.Range(sCellPosition).text
				If sAppValue = sString Then
					Fn_MSO_ExcelEditOperations = True
				End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "GetRowData"
				Set objExcel = GetObject(,"Excel.Application")
				If iWorksheets = "" Then
					iWorksheets = 1
				End If
				objExcel.Worksheets(iWorksheets).Activate
				'objExcel.Visible = True
				xlsCellPosition =  Fn_MSO_ExcelEditOperations("GetCellPosition","",iWorksheets,"", sString ,"false")
				If xlsCellPosition = False Then
						' column not found.
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_ExcelEditOperations:Specified column [ " & sString & " ] is not present in given excel sheet.")
					Exit function
				End If
				'For iCnt = 1 to len(xlsCellPosition)
				'	sChar = mid(xlsCellPosition,iCnt,1)
				'	If  Asc(sChar) >= Asc("A") and   Asc(sChar) <=  Asc("Z") then
				'		' do nothing :)
				'	else
				'		' exit from loop
				'		Exit For
				'	end if
				'Next
				' row number	
				'iCellCnt = cInt( mid(xlsCellPosition,iCnt, len(xlsCellPosition)))
				' column number
				'iColCnt =  Fn_MSO_ExcelColHeader("GetColumnHeaderNumber", mid(xlsCellPosition,1, iCnt -1)) 
				
				' row number	
				iCellCnt = Fn_MSO_ExtractExcelRowNumber(xlsCellPosition)
				' column number
				iColCnt =  Fn_MSO_ExcelColHeader("GetColumnHeaderNumber", Fn_MSO_ExtractExcelColumnName(xlsCellPosition))
				
				aRange = split(sCellPosition,"-")
				Select Case UBound(aRange)
					Case 0
						sReturnData = objExcel.Worksheets(iWorksheets).cells(iCellCnt, iColCnt + cInt(aRange(0)) ).value
					Case 1
						For iCnt = (iColCnt + cInt(aRange(0))) to (iColCnt + cInt(aRange(1)))
								If sReturnData = "" Then
									sReturnData = objExcel.Worksheets(iWorksheets).cells(iCellCnt, iCnt).value
								else
									sReturnData = sReturnData & "~" & objExcel.Worksheets(iWorksheets).cells(iCellCnt,iCnt).value
								End If
						Next
				End Select
				If sReturnData <> "" Then
					Fn_MSO_ExcelEditOperations = sReturnData
				End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -				
		Case "GetRowContents"
					sReturnData = ""
					Set objExcel = GetObject(,"Excel.Application")
					If iWorksheets = "" Then
						iWorksheets = 1
					End If
					objExcel.Worksheets(iWorksheets).Activate
					iColCnt = objExcel.Worksheets(iWorksheets).UsedRange.Columns.Count
					For iCnt = 1 to iColCnt
							If sReturnData = "" Then
										If objExcel.Worksheets(iWorksheets).cells(sCellPosition,iCnt).Value = "" Then
											sReturnData = "Blank"
										Else
											sReturnData = objExcel.Worksheets(iWorksheets).cells(sCellPosition,iCnt).Value
										End If
							Else
										If objExcel.Worksheets(iWorksheets).cells(sCellPosition,iCnt).Value = "" Then
											sReturnData = sReturnData &"~"&"Blank"
										Else
											sReturnData = sReturnData &"~"& objExcel.Worksheets(iWorksheets).cells(sCellPosition,iCnt).Value
										End If
							End If
					Next
				If sReturnData <> "" Then
					Fn_MSO_ExcelEditOperations = sReturnData
				End If				
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "VerifySaveAsDialog"
				'If Fn_UI_ObjectExist("Fn_MSO_ExcelEditOperations", Window("MicrosoftExcelWin")) Then
				if bCloseExcel = "" then bCloseExcel = "Save"
				if bCloseExcel = "Yes" then bCloseExcel = "Save"
				if bCloseExcel = "No" then bCloseExcel = "Don't Save"
				If Window("MicrosoftExcelWin").exist(10) then
						Window("MicrosoftExcelWin").Close       
						'Click on 'YES' button
						If Fn_UI_ObjectExist("Fn_MSO_ExcelEditOperations",Window("MicrosoftExcelWin").Window("Microsoft Excel").WinObject("Microsoft Excel")) Then
							Window("MicrosoftExcelWin").Window("Microsoft Excel").Activate
							Window("MicrosoftExcelWin").Window("Microsoft Excel").Highlight
							Call Fn_UI_WinButton_Click("Fn_DataSet_Operations",Window("MicrosoftExcelWin").Window("Microsoft Excel").WinObject("Microsoft Excel"),bCloseExcel,"","","")	
							wait 1
							if bCloseExcel = "Save" then bCloseExcel = ""
						ElseIf Fn_UI_ObjectExist("Fn_MSO_ExcelEditOperations", Window("MicrosoftExcelWin").Dialog("Microsoft Excel")) Then
						'Call Fn_UI_WinButton_Click("Fn_DataSet_Operations",Window("MicrosoftExcelWin").Dialog("Microsoft Excel"),bCloseExcel,"","","")	
							bReturn= Window("MicrosoftExcelWin").Dialog("Microsoft Excel").GetTextLocation(bCloseExcel,left,top,right,bottom,False)
							X=left+right
							Y=top+bottom
							If bReturn=True Then
							 Window("MicrosoftExcelWin").Dialog("Microsoft Excel").Click X/2,Y/2
							End If
							wait 1
							If bCloseExcel = "Save" Then
								bCloseExcel = ""
							End If
						End If
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_ExcelEditOperations: [ Save As ] Dialog verified successfully.")
						Fn_MSO_ExcelEditOperations =True
				else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL:Excel Window does not exist")
						Fn_MSO_ExcelEditOperations=False
						'ExitTest
				End If
		Case "VerifyCellBackground"
				Set objExcel = GetObject(,"Excel.Application")
				objExcel.Visible = bVisibleFlag
				If iWorksheets <> "" Then
					objExcel.Worksheets(iWorksheets).Activate
				End If
				' row number	
				iCellCnt = Fn_MSO_ExtractExcelRowNumber(sCellPosition)
				' column number
				iColCnt = Fn_MSO_ExcelColHeader("GetColumnHeaderNumber", Fn_MSO_ExtractExcelColumnName(sCellPosition))
				Select Case objExcel.Cells(iCellCnt, iColCnt).Interior.ColorIndex
					Case 1 'BLACK
						sColour = "BLACK"
					case 2 'WHITE
						sColour = "WHITE"
					Case 3, 22 'RED
						sColour = "RED"
					Case 4 'GREEN
						sColour = "GREEN"
					Case 5 'BLUE
						sColour = "BLUE"
					Case 6, 27 ' YELLOW
						sColour = "YELLOW"
					Case 7, 26 'MAGENTA
						sColour = "MAGENTA"
					Case 8 'CYAN
						sColour = "CYAN"
					Case 46 'ORANGE  R:255, G:102, B:0
						sColour = "ORANGE"
				end Select
				if sColour = UCase(sString) then Fn_MSO_ExcelEditOperations = True
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "GroupCells"
				bFlag = Fn_MSO_ExcelEditOperations("SelectCell",sFilePath,iWorksheets,sCellPosition,sString, bCloseExcel)
                If bFlag = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_ExcelEditOperations: Failed to select cell [ " & sCellPosition & " ].")
					Exit function
				End If
				wait 2
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_ExcelEditOperations: Successfully selected cells [ " & sCellPosition & " ].")
				Call Fn_KeyBoardOperation("SendKeys", "%~A~G~G")
				Select Case sString
					Case "Rows"
						Call Fn_KeyBoardOperation("SendKeys", "R")
					Case "Columns"
						Call Fn_KeyBoardOperation("SendKeys", "C")
				End Select
				Call Fn_KeyBoardOperation("SendKeys", "{ENTER}")
				Fn_MSO_ExcelEditOperations = True
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "GetCellComment"
				Set objExcel = GetObject(,"Excel.Application")
				objExcel.Visible = bVisibleFlag
				If iWorksheets <> "" Then
					objExcel.Worksheets(iWorksheets).Activate
				End If
				If NOT objExcel.Range(sCellPosition).comment Is Nothing Then
					Fn_MSO_ExcelEditOperations = objExcel.Range(sCellPosition).comment.Text
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_ExcelEditOperations: Successfully returned cell comment of [ " & sCellPosition & " ].")
				Else
					Fn_MSO_ExcelEditOperations = False 
				End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "InsertRow", "InsertColumn", "ShiftCellsRight", "ShiftCellsDown"
				bFlag = Fn_MSO_ExcelEditOperations("SelectCell",sFilePath,iWorksheets,sCellPosition,sString, bCloseExcel)
				If bFlag = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_ExcelEditOperations: Failed to select cell [ " & sCellPosition & " ].")
					Exit function
				End If
				wait 2
				Select Case sAction
					Case "InsertRow"
						sKey = "R"
					Case "InsertColumn"
						sKey = "C"
					Case "ShiftCellsRight"
						sKey = "I"
					Case "ShiftCellsDown"
						sKey = "D"
				End Select
				Fn_MSO_ExcelEditOperations = Fn_KeyBoardOperation("SendKeys", "^+=~" & sKey & "~{ENTER}") 
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "HideRow", "HideColumn","UnhideRow","UnhideColumn"
				If sCellPosition <> "" Then
					bFlag = Fn_MSO_ExcelEditOperations("SelectCell",sFilePath,iWorksheets,sCellPosition,sString, bCloseExcel)
					If bFlag = False Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_ExcelEditOperations: Failed to select cell [ " & sCellPosition & " ].")
						Exit function
					End If
				End If
				wait 2
				Select Case sAction
					Case "HideRow"
							'To hide selected rows – Ctrl + 9
							sKey = "^(9)"
					Case "HideColumn"
							'To hide selected columns – Ctrl + 0
							sKey = "^(0)"
					Case "UnhideRow"
							'To unhide hidden rows within the selected range – Ctrl + Shift + (
							sKey = "^+9"
					Case "UnhideColumn"
							'To unhide hidden columns within the selected range – Ctrl + Shift + )
							sKey = "^+0"
				End Select
				Fn_MSO_ExcelEditOperations = Fn_KeyBoardOperation("SendKeys",  sKey )
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "DeleteEntireRow", "DeleteEntireColumn", "ShiftCellsLeft", "ShiftCellsUp"
				bFlag = Fn_MSO_ExcelEditOperations("SelectCell",sFilePath,iWorksheets,sCellPosition,sString, bCloseExcel)
				If bFlag = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_ExcelEditOperations: Failed to select cell [ " & sCellPosition & " ].")
					Exit function
				End If
				wait 2
				Select Case sAction
					Case "DeleteEntireRow"
						sKey = "R"
					Case "DeleteEntireColumn"
						sKey = "C"
					Case "ShiftCellsLeft"
						sKey = "L"
					Case "ShiftCellsUp"
						sKey = "U"
				End Select
				Fn_MSO_ExcelEditOperations = Fn_KeyBoardOperation("SendKeys", "^-~" & sKey & "~{ENTER}") 
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "GetCellValue"
				If Fn_UI_ObjectExist("Fn_MSO_ExcelEditOperations", JavaWindow("DefaultWindow")) Then
					JavaWindow("DefaultWindow").Maximize	''To remove the Focus From Excel File
				End If
                Set objExcel = GetObject(,"Excel.Application")
				objExcel.Visible = bVisibleFlag
				If iWorksheets = "" Then
					iWorksheets = 1
				End If
				objExcel.Worksheets(iWorksheets).Activate
				' Cell Value here
				Fn_MSO_ExcelEditOperations = objExcel.Range(sCellPosition).text
		'Case to verify if a Cell contains an Image or not
		Case "ExistsImageInCell"
				If Fn_UI_ObjectExist("Fn_MSO_ExcelEditOperations", JavaWindow("DefaultWindow")) Then
					JavaWindow("DefaultWindow").Maximize	''To remove the Focus From Excel File
				End If
                		Set objExcel = GetObject(,"Excel.Application")
				objExcel.Visible = bVisibleFlag
				If iWorksheets = "" Then
					iWorksheets = 1
				End If
				objExcel.Worksheets(iWorksheets).Activate
				
				For Each wShape In objExcel.Worksheets(iWorksheets).Shapes
					If wShape.TopLeftCell.Address = sCellPosition Then
				       		Fn_MSO_ExcelEditOperations = True
				       		Exit For
					Else
				       		Fn_MSO_ExcelEditOperations = False
					End If
				Next
		'Case to verify if a Cell contains a Hyperlink or not	
		Case "ExistsHyperlinkInCell"
				If Fn_UI_ObjectExist("Fn_MSO_ExcelEditOperations", JavaWindow("DefaultWindow")) Then
					JavaWindow("DefaultWindow").Maximize	''To remove the Focus From Excel File
				End If
                		Set objExcel = GetObject(,"Excel.Application")
				objExcel.Visible = bVisibleFlag
				If iWorksheets = "" Then
					iWorksheets = 1
				End If
				objExcel.Worksheets(iWorksheets).Activate
				bFlag = False
				For Each wHlink In objExcel.Worksheets(iWorksheets).Hyperlinks
					If wHlink.Range.Address = sCellPosition Then
						If wHlink.Name = sString Then
							bFlag = True
							Exit For
						End If
					End If
				Next
				If bFlag = True Then
					Fn_MSO_ExcelEditOperations = True
				Else
					Fn_MSO_ExcelEditOperations = False
				End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		' Added For More than 1 Cell Value sCellPosition = Seperated by  ' : '
		Case "GetMultiCellValues"
                Set objExcel = GetObject(,"Excel.Application")
				objExcel.Visible = bVisibleFlag
				If iWorksheets = "" Then
					iWorksheets = 1
				End If
				objExcel.Worksheets(iWorksheets).Activate
				sValue=""
				aRange = Split(sCellPosition,":")
				' Cell Value here
                For iCount = 0 To Ubound(aRange)
					If iCount = 0 Then
						sValue=objExcel.Range(aRange(iCount)).text
					Else
						sValue = sValue +"~"+ objExcel.Range(aRange(iCount)).text
					End If
				Next
				Fn_MSO_ExcelEditOperations = Trim(sValue)
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "FollowHyperlink"
				bFlag = Fn_MSO_ExcelEditOperations("SelectCell",sFilePath,iWorksheets,sCellPosition,sString, bCloseExcel)
				If bFlag = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_ExcelEditOperations: Failed to select cell [ " & sCellPosition & " ].")
					Exit function
				End If
				wait 2
				Set objExcel = GetObject(,"Excel.Application")
				objExcel.Visible = bVisibleFlag
				If iWorksheets <> "" Then
					objExcel.Worksheets(iWorksheets).Activate
				End If
				objExcel.Range(sCellPosition).Hyperlinks(1).Follow
				Fn_MSO_ExcelEditOperations = True
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "GetSelectedCellNumber"	
				Set objExcel = GetObject(,"Excel.Application")
				objExcel.Visible = bVisibleFlag
				If iWorksheets <> "" Then
					objExcel.Worksheets(iWorksheets).Activate
				End If
				iRow = objExcel.ActiveCell.Row
				iCol = objExcel.ActiveCell.Column			
				iCol = Fn_MSO_ExcelColHeader("GetColumnHeaderName", iCol)
				Fn_MSO_ExcelEditOperations = iCol&""&iRow
		' - - - - - - - - - - - - - - - - - - - - -[SHREYAS 27-02-2012] - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "GetSheetName"
				If Fn_UI_ObjectExist("Fn_MSO_ExcelEditOperations", JavaWindow("DefaultWindow")) Then
					JavaWindow("DefaultWindow").Minimize'To remove the Focus From Excel File Added by Jotiba T on 2-Nov-2017
	           		JavaWindow("DefaultWindow").Maximize
				End If
	            iTimeout=120
			   	bFlag=False
				Set objExcel = GetObject(,"Excel.Application")
				For iCount= 0 To iTimeout
						If objExcel.Visible= True Then
							bFlag=True
							Exit For
						End If
				Next
				'End 
			Fn_MSO_ExcelEditOperations = objExcel.Worksheets(cint(sCellPosition)).Name
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_ExcelEditOperations: Executed successfully with case [ " & sAction & " ].")
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

		Case "ModifySheetName"

			Set objExcel = GetObject(,"Excel.Application")
			objExcel.Visible = bVisibleFlag
			wait(2)
			objExcel.Worksheets(sCellPosition).Activate
			wait(2)
			objExcel.Worksheets(cint(sCellPosition)).Name = sString
			Fn_MSO_ExcelEditOperations = True
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

		Case "GetLastRowIndex"

			Set objExcel=GetObject(,"Excel.Application")
			Set iExcel=objExcel.Worksheets(sCellPosition)
			sRowCount=iExcel.usedrange.rows.count
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Total Number of Used Rows is ["+cstr(sRowCount)+"]" )
			wait 1
			Fn_MSO_ExcelEditOperations = sRowCount
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "CutPasteCol"

				aCol=split(sString,":",-1,1)
				Set objExcel = GetObject(,"Excel.Application")
				objExcel.Visible = bVisibleFlag
				objExcel.Worksheets(sCellPosition).Activate
				wait 3
				Set oSheet = objExcel.Activesheet
				
						bReturn= Fn_MSO_ExcelEditOperations("GetCellPosition","",sCellPosition,"",aCol(0),"")
						sValue = mid(bReturn,1,1)
						sCol=Fn_MSO_ExcelColHeader("GetColumnHeaderNumber", sValue)
				
				sColname=Fn_MSO_ExcelColHeader("GetColumnHeaderName", sCol)
				oSheet.Columns(sColname).Select
				
				oSheet.Columns(sColname).Cut
				wait 3
				
						bReturn= Fn_MSO_ExcelEditOperations("GetCellPosition","",sCellPosition,"",aCol(1),"")
						sValue = mid(bReturn,1,1)
						sCol=Fn_MSO_ExcelColHeader("GetColumnHeaderNumber", sValue)
				
				sColname=Fn_MSO_ExcelColHeader("GetColumnHeaderName", sCol)
				objExcel.Columns(sColname).Select
				wait 3
				'Paste the Row
				Set shell=CreateObject("Wscript.Shell")
				shell.SendKeys("^{+}")

				Fn_MSO_ExcelEditOperations=true
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "UnGroupCells"
				bFlag = Fn_MSO_ExcelEditOperations("SelectCell",sFilePath,iWorksheets,sCellPosition,sString, bCloseExcel)
                If bFlag = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_ExcelEditOperations: Failed to select cell [ " & sCellPosition & " ].")
					Exit function
				End If
				wait 2
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_ExcelEditOperations: Successfully selected cells [ " & sCellPosition & " ].")
				Call Fn_KeyBoardOperation("SendKeys", "%~A~U~U")
				Select Case sString
					Case "Rows"
						Call Fn_KeyBoardOperation("SendKeys", "R")
					Case "Columns"
						Call Fn_KeyBoardOperation("SendKeys", "C")
				End Select
				Call Fn_KeyBoardOperation("SendKeys", "{ENTER}")
				Fn_MSO_ExcelEditOperations = True
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'[TC1123_20161205c00_PoonamC-NewDevelopment : Added Cases to modify excel file & close when multiple instances opened]
		Case "ModifyCellForMultipleInstance"
		    Set objExcel = Window("MSExcelMultiInstance")
			Wait(5)
		    objExcel.SetTOProperty "regexpwndtitle",sFilePath
		    ' Chnages made for MSExcel2013 
		    If Not Fn_SISW_UI_Object_Operations("Fn_MSO_ExcelEditOperations","Exist", objExcel,SISW_MIN_TIMEOUT) Then
		    	objExcel.SetTOProperty "regexpwndtitle","Excel"
		    	objExcel("Text").RegularExpression=True
		    	objExcel("Text").Value=Replace(sFilePath,".xlsx",".*")
		    End If 
		    If Not Fn_SISW_UI_Object_Operations("Fn_MSO_ExcelEditOperations","Exist", objExcel,SISW_MIN_TIMEOUT) Then
		    	objExcel.SetTOProperty "regexpwndtitle","Saved"
		    	objExcel("Text").RegularExpression=True
		    	objExcel("Text").Value=Replace(sFilePath,".xlsx",".*")
		    End If 
		    If Fn_SISW_UI_Object_Operations("Fn_MSO_ExcelEditOperations","Exist", objExcel,SISW_MIN_TIMEOUT) Then
		     	objExcel.highlight 
				Call Fn_KeyBoardOperation("SendKeys","^{HOME}")
				objExcel.Type sString
			    Wait(1)
			    objExcel.WinObject("Ribbon").WinToolbar("QuickAccessToolbar").Press "Save"
			    Wait(1)
				Fn_MSO_ExcelEditOperations = True
			Else
				Fn_MSO_ExcelEditOperations = False
			End If
			     		
		Case "CloseMultipleInstance"
			 Set objExcel = Window("MSExcelMultiInstance")
			 objExcel.SetTOProperty "regexpwndtitle",sFilePath
			   ' Chnages made for MSExcel2013 
			    If Not Fn_SISW_UI_Object_Operations("Fn_MSO_ExcelEditOperations","Exist", objExcel,SISW_MIN_TIMEOUT) Then
			    	objExcel.SetTOProperty "regexpwndtitle","Excel"
			    	objExcel("Text").RegularExpression=True
			    	objExcel("Text").Value=Replace(sFilePath,".xlsx",".*")
			    End If 
				If Not Fn_SISW_UI_Object_Operations("Fn_MSO_ExcelEditOperations","Exist", objExcel,SISW_MIN_TIMEOUT) Then
			    	objExcel.SetTOProperty "regexpwndtitle","Saved"
			    	objExcel("Text").RegularExpression=True
			    	objExcel("Text").Value=Replace(sFilePath,".xlsx",".*")
			    End If 
			 If Fn_SISW_UI_Object_Operations("Fn_MSO_ExcelEditOperations","Exist", objExcel,SISW_MIN_TIMEOUT) Then
				 objExcel.highlight 
			     objExcel.Close()
			     Wait(1)
				Fn_MSO_ExcelEditOperations = True	
			Else
				Fn_MSO_ExcelEditOperations = False
			End If	
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -	
		'TC11.4(20171113a00)_DIPRO_NewDevelopment_PoonamC_22Nov2017 : Added Case to SaveAs file to location speified
		Case "SaveAs"
				If Window("MicrosoftExcel").Exist(5)=True Then
					Window("MicrosoftExcel").Minimize
					wait 1
					Window("MicrosoftExcel").Maximize
					wait 3
				End If
				Set objExcel = GetObject(,"Excel.Application")
				 	objExcel.Visible = True 
				If iWorksheets <> "" Then
					objExcel.Worksheets(iWorksheets).SaveAs sFilePath
					Fn_MSO_ExcelEditOperations = True
				End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -	
		'TC11.6(20181001b00)_DIPRO_NewDevelopment_PoonamC_26Oct2018 
		Case "VerifyValueInVisibleCells"
				If Fn_UI_ObjectExist("Fn_MSO_ExcelEditOperations", JavaWindow("DefaultWindow")) Then
					JavaWindow("DefaultWindow").Minimize
					JavaWindow("DefaultWindow").Maximize
				End If	
				iTimeout=120
			  	bFlag=False
				Set objExcel = GetObject(,"Excel.Application")
				For iCount=0 To iTimeout
					If objExcel.Visible= True Then
						bFlag=True
						Exit For
					End If
				Next	
				If iWorksheets = "" Then
					iWorksheets = 1
				End If
				Window("MicrosoftExcel").Maximize
				objExcel.Worksheets(iWorksheets).Activate
				
				Const xlCellTypeVisible = 12
				Set objExcel = GetObject(,"Excel.Application")
					objExcel.Visible = True 
				Set bVisibleFlag = objExcel.Worksheets(iWorksheets).Range(sCellPosition).SpecialCells(xlCellTypeVisible).Find(sString)	
				If bVisibleFlag Is Nothing Then
				    Fn_MSO_ExcelEditOperations = False
				Else
				    Fn_MSO_ExcelEditOperations = True
				End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -		
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_ExcelEditOperations: Invalid case [ " & sAction & " ].")
			Fn_MSO_ExcelEditOperations = False
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	End Select

	Select Case bCloseExcel
		Case "True", "Yes", True
			objExcel.Quit
    End Select
    Wait 2
	Window("MicrosoftExcel").Window("Excel Options").SetTOProperty "regexpwndtitle","Microsoft Excel"
	If Window("MicrosoftExcel").Window("Excel Options").Exist(5) Then
		If Window("MicrosoftExcel").Window("Excel Options").WinObject("MicrosoftExcel").WinButton("Save").Exist(1) Then
			Call Fn_UI_WinButton_Click("Fn_MSO_ExcelEditOperations",Window("MicrosoftExcel").Window("Excel Options").WinObject("MicrosoftExcel"), "Save",5,5,micLeftBtn)
		End If
	End If
	if Fn_MSO_ExcelEditOperations = True then Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_ExcelEditOperations: Executed successfully with case [ " & sAction & " ].")
	Set objExcel = Nothing
End Function


'*********************************************************	Function to Insert Headers and Footers in MS Word application.  **************************************
'Function Name		:				Fn_MSO_WordHeaderFooter  

'Description			 :		 		This function will Insert or Verify Headers and Footers in MS Word application.

'Return Value		   : 				True/False  

'Examples				:				'Msgbox Fn_MSO_WordHeaderFooter("InsertHeaderFooter" ,"D:\Mainline\Test.docx", "Programmer : Ketan", "Contact : 9821087858", "")
											  'Msgbox Fn_MSO_WordHeaderFooter("InsertHeaderFooter" ,"", "Programmer : Ketan", "Contact : 9821087858", "")

'History					 :		
'										Developer Name				Date				Rev. No.			Changes Done	
'									-------------------------------------------------------------------------------------------------
'										Ketan Raje					14/03/2011		       1.0						Created
'									-------------------------------------------------------------------------------------------------
'**********************************************************************************************************************************************************************************
Public Function Fn_MSO_WordHeaderFooter(sAction ,sFilePath, sHeader, sFooter, aGlobalDictionary)
	GBL_FAILED_FUNCTION_NAME="Fn_MSO_WordHeaderFooter"
	Dim lobjWord, ObjSection, iCount, iCounter
	Fn_MSO_WordHeaderFooter = False
	Set	lobjWord = Nothing
	If sFilePath <> "" Then
		Set lobjWord = CreateObject("Word.Application")
		lobjWord.Visible = True
		lobjWord.Documents.Open sFilePath,False,False
	Else
		Set lobjWord = GetObject(,"Word.Application")		
	End If	
	Select Case sAction
				Case"InsertHeaderFooterInsideTC","InsertHeaderFooterOutsideTC" 'This Case May Give Warning But Will not create any issues in Single or Batch Run.
						Set ObjSection = lobjWord.ActiveDocument.Sections(1)
							If sHeader <> "" Then
								ObjSection.Headers(1).Range.Text = sHeader
							End If
							If sFooter <> "" Then
								ObjSection.Footers(1).Range.Text = sFooter
							End If
							Fn_MSO_WordHeaderFooter = True
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Pass:Sucessfully Inserted Header/Footer.")
				Case"VerifyHeaderFooterInsideTC","VerifyHeaderFooterOutsideTC" 'This Case May Give Warning But Will not create any issues in Single or Batch Run.
						iCount = 0
						iCounter = 0
						Set ObjSection = lobjWord.ActiveDocument.Sections(1)
							If sHeader <> "" Then
								iCount = iCount + 1
								If Instr(1,Trim(Lcase(ObjSection.Headers(1).Range.Text)),Trim(Lcase(sHeader))) <> 0 Then
									iCounter = iCounter + 1
								End If
							End If
							If sFooter <> "" Then
								iCount = iCount + 1
								If Instr(1,Trim(Lcase(ObjSection.Footers(1).Range.Text)),Trim(Lcase(sFooter))) <> 0 Then
									iCounter = iCounter + 1
								End If
							End If
							If iCount = iCounter Then
								Fn_MSO_WordHeaderFooter = True
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Pass:Sucessfully Inserted Header/Footer.")
							End If
    End Select
			If sFilePath <> "" Then
				lobjWord.ActiveDocument.SaveAs sFilePath
				lobjWord.Quit
			Else
				lobjWord.ActiveDocument.Save
				If Instr(1,sAction,"OutsideTC") <> 0 Then
					lobjWord.Quit
				End If
			End If
			Set	lobjWord=nothing
End function
'------------------------------------------------------------------------Function to Perform Operations On MS Picture Manager.---------
'Function Name		  :	  Fn_MSO_PictureManagerOperation

'Description			 :	 Function to Perform Operations On MS Picture Manager

'Parameters			   :	sAction, StrText

'Return Value		   : 	True Or False

'Examples				:	Call Fn_MSO_PictureManagerOperation("Exist","")

'History					 :					Developer Name												Date						Rev. No.	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep Navghane										05/05/2011			           1.0		
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function Fn_MSO_PictureManagerOperation(StrAction,StrText)
	GBL_FAILED_FUNCTION_NAME="Fn_MSO_PictureManagerOperation"
	Dim objPicMgr
	Set objPicMgr=Window("text:=Microsoft Office Picture Manager")
	Fn_MSO_PictureManagerOperation=False
	Select Case StrAction
		Case "Exist"
			If objPicMgr.Exist(6) Then
				If objPicMgr.GetVisibleText()<>"" Then
					Fn_MSO_PictureManagerOperation=True	
				End If
				objPicMgr.Close
			End If
	End Select
	Set objPicMgr=Nothing
End Function
'*********************************************************		Function to handle operations with MSO files.  ********************************************************************
' Function Name	:	Fn_MSO_FileOperations  

' Return Value	:	True/False  and also FileName.

' Examples		:	Msgbox Fn_MSO_FileOperations("GetFileName" ,"Word", "", "")
'					Msgbox Fn_MSO_FileOperations("GetFileName" ,"Excel", "", "")
'					Msgbox Fn_MSO_FileOperations("FileSaveAs" ,"Word", "D:\Mainline\ABC.docx", "")
'					Msgbox Fn_MSO_FileOperations("FileSaveAs" ,"Excel", "D:\Mainline\ABC.csv", "")
'					Msgbox Fn_MSO_FileOperations("FileClose" ,"Word", "", "")
'					Msgbox Fn_MSO_FileOperations("FileClose" ,"Excel", "", "")
'					Msgbox Fn_MSO_FileOperations("IsReadOnly" ,"Excel", "", "")
'
'				Example to get HWND property of word window	
'					Set dicFileDetails = CreateObject("Scripting.Dictionary")
'					dicFileDetails.RemoveAll
'					dicFileDetails("SubAction") = "GetHWND"
'					bReturn = Fn_MSO_FileOperations("MSWordMultipleInstance","Word","",dicFileDetails)
'
'				Example to verify Word file is Read only or not using HWND property
'					dicFileDetails.RemoveAll
'					dicFileDetails("SubAction") = "IsReadOnly"
'					dicFileDetails("HWNDProperty") = sHWNDProperty2
'					bReturn = Fn_MSO_FileOperations("MSWordMultipleInstance","Word","",dicFileDetails)
'
'				Example to Modify Word file using HWND property
'					dicFileDetails.RemoveAll
'					dicFileDetails("SubAction") = "ModifyWordDoc"
'					dicFileDetails("HWNDProperty") = sHWNDProperty4
'					dicFileDetails("TextToModify") = "Test"
'					bReturn = Fn_MSO_FileOperations("MSWordMultipleInstance","Word","",dicFileDetails)
'					
'				Example to Close Word Instance file using HWND property		'First get HWND propety at runtime of opened word using "GetHWND" case
'					dicFileDetails("SubAction") = "CloseInstance"
'					dicFileDetails("HWNDProperty") = "12465783"
'					bReturn = Fn_MSO_FileOperations("MSWordMultipleInstance","Word","",dicFileDetails)
'
'				Example to Modify Word Instance with same name using HWND (to modify 2nd instance) 	'First get HWND propety at runtime of opened word using "GetHWND" case
'					dicFileDetails("SubAction") = "ModifyWordInstanceOfSameName"
'					dicFileDetails("HWNDProperty") = "50726278"
'					dicFileDetails("TextToModify") = "Tweest"
'					bReturn = Fn_MSO_FileOperations("MSWordMultipleInstance","Word","",dicFileDetails)
'
'				Example to Verify Word Instance with same name using HWND (to Verify 2nd instance) 	'First get HWND propety at runtime of opened word using "GetHWND" case
'					dicFileDetails("SubAction") = "VerifyWordInstanceOfSameName"
'					dicFileDetails("HWNDProperty") = "50726278"
'					dicFileDetails("TextToVerify") = "TestTweest"
'					bReturn = Fn_MSO_FileOperations("MSWordMultipleInstance","Word","",dicFileDetails)
'					
' History	:		
'	Developer Name		Date	Rev. No.	Changes Done										Reviewer
'-----------------------------------------------------------------------------------------------------------------------------------------
'	Ketan Raje		18/05/2011	 1.0		Created
'-----------------------------------------------------------------------------------------------------------------------------------------
'	Koustubh W		09/06/2011	 1.1		Added case FileClose
'-----------------------------------------------------------------------------------------------------------------------------------------
'	Sandeep N		17/10/2011	 1.2		Added case FileSave
'-----------------------------------------------------------------------------------------------------------------------------------------
'	Ketan Raje		19/10/2011	 1.2		Added case IsReadOnly
'-----------------------------------------------------------------------------------------------------------------------------------------
'	Vivek Ahirrao	07/04/2016	 1.3		Added case "MSWordMultipleInstance"					[TC1122-20160323-07_04_2016-VivekA-NewDevelopment]
'**********************************************************************************************************************************************************************************
Public Function Fn_MSO_FileOperations(sAction ,sFileType, sFilePath, aGlobalDic)
	GBL_FAILED_FUNCTION_NAME="Fn_MSO_FileOperations"
	Dim objFile,sWString
	Fn_MSO_FileOperations = False
	Set	objFile=nothing
	Select Case sAction
		Case "MSWordMultipleInstance"
			Set objWordWindow = Window("MSWordMultiInstance")
				Select Case sFileType
						'Case for Multiple instances of Word file
						Case "Word"
								Select Case aGlobalDic("SubAction")
										'Case to get Handle Window property of latest opened Doc file
										Case "GetHWND"
											If objWordWindow.Exist Then
												sHWNDProperty = objWordWindow.GetROProperty("hwnd")
												Fn_MSO_FileOperations = sHWNDProperty
												Set objWordWindow = Nothing
												Exit Function
											End If
											Set objWordWindow = Nothing
										Case "IsReadOnly"
											If aGlobalDic("HWNDProperty")<>"" Then
												Set objWordWindow = Window("regexpwndtitle:=.*Word.*","hwnd:="&aGlobalDic("HWNDProperty"))
												'objWordWindow.SetTOProperty "hwnd",aGlobalDic("HWNDProperty")
												objWordWindow.Highlight
												If objWordWindow.Exist Then
													sTitle = objWordWindow.GetROProperty("text")
													If Instr(sTitle,"Read-Only")>0 Then
														Fn_MSO_FileOperations = True
													Else 
														Fn_MSO_FileOperations = False
													End If
												End If
											End If
											Set objWordWindow = Nothing
										'[TC1123(20161205c00)_PoonamC_NewDevelopment_07Feb2017 : Added SubCase "VerifyTitle" - Case to verify title for opened Doc file]
										Case "VerifyTitle"
											If aGlobalDic("HWNDProperty")<>"" Then
												Set objWordWindow = Window("regexpwndtitle:=.*Word.*","hwnd:="&aGlobalDic("HWNDProperty"))
												'objWordWindow.SetTOProperty "hwnd",aGlobalDic("HWNDProperty")
												If objWordWindow.Exist Then
													sTitle = objWordWindow.GetROProperty("text")
													If Instr(sTitle,aGlobalDic("title"))>0 Then
														Fn_MSO_FileOperations = True
													Else 
														Fn_MSO_FileOperations = False
													End If
												End If
											End If
											Set objWordWindow = Nothing	
										Case "ModifyWordDoc","ModifyWordDocWtoutSave"
											If aGlobalDic("HWNDProperty")<>"" Then
												Set objWordWindow = Window("regexpwndtitle:=.*Word.*","hwnd:="&aGlobalDic("HWNDProperty"))
												'objWordWindow.SetTOProperty "hwnd",aGlobalDic("HWNDProperty")
												If objWordWindow.Exist Then
													bFlag = Fn_KeyBoardOperation("SendKeys", "+{HOME}")
													If bFlag = True Then
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Successfully entered END key to get Cursor location at END.")
														Fn_MSO_FileOperations=True
													Else
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to enter END key to get Cursor location at END.")
													   	Fn_MSO_FileOperations=False
													   	Set objWordWindow = Nothing
													   	Exit function	
													End If
													If aGlobalDic("TextToModify")<>"" Then
														Set objWord = GetObject(,"Word.application")
														objWord.Visible=true
														Wait(2)
														objWord.Selection.TypeText(aGlobalDic("TextToModify"))
														Wait(2)
														If aGlobalDic("SubAction")<>"ModifyWordDocWtoutSave" Then
															objWord.ActiveDocument.Save
															Wait(2)
															objWord.Quit
														End If
														Fn_MSO_FileOperations=True
														Set objWord = Nothing
														Set objWordWindow = Nothing
														Exit Function
													Else 
														Fn_MSO_FileOperations=False
														Set objWordWindow = Nothing
														Exit Function
													End If
												End If
												'Window("MSWordMultiInstance").SetTOProperty "hwnd",""
											End If
											Set objWordWindow = Nothing
										Case "WordGetContent"
											If aGlobalDic("HWNDProperty")<>"" Then
												Set objWordWindow = Window("regexpwndtitle:=.*Word.*","hwnd:="&aGlobalDic("HWNDProperty"))
												If objWordWindow.Exist Then
													Set objWord = GetObject(,"Word.application")
													objWord.Visible=true
													Wait(2)
													objWord.Selection.WholeStory 
													Set oSelection = objWord.Selection 
													sWString=  CStr(oSelection.Text)
													If Instr(1,sWString,chr(13))>0 Then
														sWString = Trim(Replace(Replace(sWString,chr(13),"*"),"*"," "))
													End If
													If sWString<>"" Then
														Fn_MSO_FileOperations = sWString
														Set objWord = Nothing
														Set objWordWindow = Nothing
														Exit Function
													Else 
														Fn_MSO_FileOperations = False 
														Set objWord = Nothing
														Set objWordWindow = Nothing
														Exit Function
													End If
												End If
											End If
										Case "CloseInstance"
											If aGlobalDic("HWNDProperty")<>"" Then
												Set objWordWindow = Window("regexpwndtitle:=.*Word.*","hwnd:="&aGlobalDic("HWNDProperty"))
												If objWordWindow.Exist Then
													objWordWindow.Close
													Fn_MSO_FileOperations=True
													Set objWordWindow = Nothing
													Exit Function
												End If
											End If
										Case "ActivateInstance"
											If aGlobalDic("HWNDProperty")<>"" Then
												Set objWordWindow = Window("regexpwndtitle:=.*Word.*","hwnd:="&aGlobalDic("HWNDProperty"))
												If objWordWindow.Exist Then
													objWordWindow.Activate
													Fn_MSO_FileOperations=True
													Set objWordWindow = Nothing
													Exit Function
												End If
											End If
										Case "ModifyWordInstanceOfSameName"
											If aGlobalDic("HWNDProperty")<>"" Then
												Set objWordWindow = Window("regexpwndtitle:=.*Word.*","hwnd:="&aGlobalDic("HWNDProperty"))
												Set objWord = Window("MSWordMultiInstance")
												Set objDoc = Window("MSWordMultiInstance").WinObject("MicrosoftWordDocument")
												If objWordWindow.Exist Then
													objWordWindow.Activate
													If objDoc.Exist Then
														objDoc.Click 5,5,micLeftBtn
														Wait 0,500
														Set WshShell = CreateObject("WScript.Shell")
														WshShell.SendKeys "{END}"
														Wait 0,500
														WshShell.SendKeys aGlobalDic("TextToModify")
														Set WshShell = nothing
														Wait 0,500
														objWord.WinObject("Ribbon").WinButton("Save").Click 5,5,micLeftBtn
														Fn_MSO_FileOperations=True
													End If
												End If
												Set objWordWindow = Nothing
												Set objWord = Nothing
												Set objDoc = Nothing
											End If
										Case "VerifyWordInstanceOfSameName"
											If aGlobalDic("HWNDProperty")<>"" Then
												Set objWordWindow = Window("regexpwndtitle:=.*Word.*","hwnd:="&aGlobalDic("HWNDProperty"))
												Set objWord = Window("MSWordMultiInstance")
												Set objDoc = Window("MSWordMultiInstance").WinObject("MicrosoftWordDocument")
												objWordWindow.Maximize                              'To remove the Focus From Word File
												If objWordWindow.Exist Then
													objWordWindow.Activate
													If objDoc.Exist Then
														objDoc.Click 5,5,micLeftBtn
														Wait 1
														'sAppText = Window("MSWordMultiInstance").WinObject("MicrosoftWordDocument").GetVisibleText
														'sAppText=replace(sAppText,vbcrlf,"")
														sAppText = objDoc.GetVisibleText()           'changed object MicrosoftWordDocument to get visible text from MS Word
														sAppText= Split(sAppText,vbcrlf,-1,1)					
														If sAppText(0) = aGlobalDic("TextToVerify") Then
															Fn_MSO_FileOperations = True
														Else
															Fn_MSO_FileOperations = False
															Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to verify the contents ["+aGlobalDic("TextToVerify")+"].")
														End If
														Wait 1
													End If
												End If
												Set objWordWindow = Nothing
												Set objWord = Nothing
												Set objDoc = Nothing
											End If
								End Select
						'Case for Multiple instances of Excel file
						Case "Excel"
								'Future Use
				End Select
		Case "GetFileName","IsReadOnly" 'This Case May Give Warning But Will not create any issues in Single or Batch Run.
					Select Case sFileType
								Case "Word"
										If Fn_UI_ObjectExist("Fn_MSO_FileOperations", JavaWindow("DefaultWindow")) Then
											JavaWindow("DefaultWindow").Minimize
											wait 1
											JavaWindow("DefaultWindow").Maximize'To remove the Focus From Word File
										End If
										Wait(2)
										Set objFile = GetObject(,"Word.Application")
										objFile.Visible = True
										If sAction = "GetFileName" Then
											Fn_MSO_FileOperations = objFile.ActiveDocument.Name
										ElseIf sAction = "IsReadOnly" Then
											Fn_MSO_FileOperations = objFile.ActiveDocument.ReadOnly
										End If										
								Case "Excel"
										If Fn_UI_ObjectExist("Fn_MSO_FileOperations", JavaWindow("DefaultWindow")) Then
											JavaWindow("DefaultWindow").Maximize'To remove the Focus From Word File
										End If
										Wait(2)
										Set objFile = GetObject(,"Excel.Application")
										objFile.Visible = True																				
										If sAction = "GetFileName" Then
											Fn_MSO_FileOperations = objFile.ActiveWorkbook.Name
										ElseIf sAction = "IsReadOnly" Then
											Fn_MSO_FileOperations = objFile.ActiveWorkbook.ReadOnly
										End If	
								Case "PowerPoint"
										If Fn_UI_ObjectExist("Fn_MSO_FileOperations", JavaWindow("DefaultWindow")) Then
											JavaWindow("DefaultWindow").Maximize'To remove the Focus From PowerPoint File
										End If
										Wait(2)
										Set objFile = GetObject(,"PowerPoint.Application")
										objFile.Visible = True
										If sAction = "GetFileName" Then
											Fn_MSO_FileOperations = objFile.ActivePresentation.Name
										ElseIf sAction = "IsReadOnly" Then
											Fn_MSO_FileOperations = objFile.ActivePresentation.ReadOnly
										End If										
					End Select		
		Case "FileSaveAs" 'This Case May Give Warning But Will not create any issues in Single or Batch Run.				
					Select Case sFileType
								Case "Word","WordWithoutExit"
										If Fn_UI_ObjectExist("Fn_MSO_ExcelEditOperations", JavaWindow("DefaultWindow")) Then
											JavaWindow("DefaultWindow").Maximize'To remove the Focus From Word File
										End If
										Wait(2)
										Set objFile = GetObject(,"Word.Application")
										objFile.Visible = True										
										objFile.ActiveDocument.SaveAs sFilePath
										If sFileType<>"WordWithoutExit" Then
											objFile.Quit
										End If
										Fn_MSO_FileOperations = True
								Case "Excel"
										If Fn_UI_ObjectExist("Fn_MSO_ExcelEditOperations", JavaWindow("DefaultWindow")) Then
											JavaWindow("DefaultWindow").Maximize'To remove the Focus From Excel File
										End If
										Wait(2)
										Set objFile = GetObject(,"Excel.Application")
										objFile.Visible = True										
										objFile.ActiveWorkbook.SaveCopyAs sFilePath
                            			objFile.Quit
										Fn_MSO_FileOperations = True
					End Select
		Case "FileSave"
					Select Case sFileType
                                Case "Excel"
										If Fn_UI_ObjectExist("Fn_MSO_ExcelEditOperations", JavaWindow("DefaultWindow")) Then
											JavaWindow("DefaultWindow").Maximize'To remove the Focus From Excel File
										End If
										Wait(2)
										Set objFile = GetObject(,"Excel.Application")
										objFile.Visible = True										
										objFile.ActiveWorkbook.Save
                            			objFile.Quit
										Fn_MSO_FileOperations = True
					End Select
		Case "FileClose"
					Select Case sFileType
								Case "Word"
'							Set objFile = GetObject(,"Word.Application")
'										objFile.Visible = True										
'										objFile.Quit
'										Fn_MSO_FileOperations = True
'									  If Window("MicrosoftWordWin").Dialog("Microsoft Office Word").Exist(5) Then         '''Added by Avinash J        27-Dec-2012
'														Window("MicrosoftWordWin").Dialog("Microsoft Office Word").WinButton("No").Click
'														  If WpfWindow("Save").Exist(5) Then
'                                                    			 WpfWindow("Save").WpfButton("OK").Click 5,5,micLeftBtn
'														  End If
'									 End If
							Call Fn_WindowsApplications("TerminateAll", "WINWORD.EXE")
							' TC112-2015070800-22_07_2015-Porting-AnkitN-Handled Close Dialog As per design change .
							If WpfWindow("Close").Exist(5) Then
								WpfWindow("Close").WpfButton("OK").Click 10, 10
								If Err.number < 0  Then
									Fn_MSO_FileOperations = False
									Exit Function
								End If 
							End If

							Fn_MSO_FileOperations = True
								Case "Excel"
										'Code Added by Nilesh to handle Dialogs and Windows after Excel close 
										Dim iOpenWorkbook,objDialog,iCount,objOpenWindow,objOpenDialog,i
										Set objFile = GetObject(,"Excel.Application")
										objFile.Visible = True										
                            			objFile.Quit
										Fn_MSO_FileOperations = True
										Set objOpenWindow =  Fn_SISW_UI_Object_GetChildObjects("Fn_MSO_FileOperations", Window("MicrosoftExcel") , "Class Name", "Window")
										If TypeName(objOpenWindow) <> "Nothing"  Then
											'To close Open Window
											For i=0 to objOpenWindow.Count-1
												objOpenWindow(i).Type "N"
'												wait SISW_MICRO_TIMEOUT
'												If objOpenWindow(i).Exist(SISW_MICRO_TIMEOUT) Then
'													objOpenWindow(i).Type "N"
'													objOpenWindow(i).Close()
'												End If
											Next
										End If
										Set objDialog =  Fn_SISW_UI_Object_GetChildObjects("Fn_MSO_FileOperations", Window("MicrosoftExcel") , "Class Name", "Dialog")
										If TypeName(objDialog) <> "Nothing"  Then
											'To close Open Window
											For i=0 to objDialog.Count-1
												objDialog(i).Type "N"
'												wait SISW_MICRO_TIMEOUT
'												If objDialog(i).Exist(SISW_MICRO_TIMEOUT) Then
'													objDialog(i).Close()
'												End If
											Next
										End If
'										objFile.close
										If Dialog("ExceClose").Exist(SISW_MICRO_TIMEOUT) Then
											Dialog("ExceClose").Type "N"
											If Dialog("ExceClose").Exist(SISW_MICRO_TIMEOUT)Then
												Dialog("ExceClose").Close()
											End If
										End If
					End Select
    End Select
	Set	objFile=nothing
End function
'***********************************	Function to get Excel Column Header Name / Number.  ********************************************************************
'Function Name			:				Fn_MSO_ExcelColHeader  

'Return Value		    : 			Excel Column Header Name / Number  / False

'Examples				:		 Msgbox Fn_MSO_ExcelColHeader("GetColumnHeaderName", 700)
'						 		 Msgbox Fn_MSO_ExcelColHeader("GetColumnHeaderNumber", "ZX")
'History				:		
'						Developer Name				Date				Rev. No.			Changes Done	
'						-------------------------------------------------------------------------------------------------
'						Koustubh					19/05/2011		       1.0				Created
'						-------------------------------------------------------------------------------------------------
'**********************************************************************************************************************************************************************************
Public Function Fn_MSO_ExcelColHeader(sAction, sColumn)
	GBL_FAILED_FUNCTION_NAME="Fn_MSO_ExcelColHeader"
	Dim temp, iCnt, iColCnt 
	Select Case sAction
		Case "GetColumnHeaderName"
				If sColumn > 26 Then
						temp = (sColumn Mod 26)
						sColumn = (sColumn - temp) / 26
						Fn_MSO_ExcelColHeader = Chr(64 + sColumn) & Chr(64 + temp)
				Else
						Fn_MSO_ExcelColHeader = Chr(64 + sColumn)
				End If
				
		Case "GetColumnHeaderNumber"
				Fn_MSO_ExcelColHeader = 0
				If len(sColumn) > 1 Then
						iColCnt = 1
						For iCnt = len(sColumn)-1 to 0  step -1
							sChar = mid(sColumn,iColCnt,1)
							Fn_MSO_ExcelColHeader = Fn_MSO_ExcelColHeader + (Asc(sChar) - 64 )  * (26^iCnt)
							iColCnt = iColCnt+1
						Next
				else
						Fn_MSO_ExcelColHeader = Asc(sColumn) - 64
				End If
	End Select
End Function
'***********************************	Function to extract Row Number from cellname.  ********************************************************************
'Function Name			:				Fn_MSO_ExtractExcelRowNumber  

'Return Value		    : 			Excel Column Row Number  / False

'Examples				:		 Msgbox Fn_MSO_ExtractExcelRowNumber("AA4")
'History				:		

'Developer Name				Date				Rev. No.			Changes Done	
'-------------------------------------------------------------------------------------------------
'Koustubh					26/05/2011		       1.0				Created
'-------------------------------------------------------------------------------------------------
Public function Fn_MSO_ExtractExcelRowNumber(sCell)
	GBL_FAILED_FUNCTION_NAME="Fn_MSO_ExtractExcelRowNumber"
	dim iCnt,sChar
	Fn_MSO_ExtractExcelRowNumber = False
	For iCnt = 1 to len(sCell)
		sChar = mid(sCell,iCnt,1)
		If  Asc(sChar) >= Asc("A") and Asc(sChar) <= Asc("Z") then
			' do nothing :)
		else
			' exit from loop
			Exit For
		end if
	Next
	if iCnt <> len(sCell)+1 then
		' row number	
		Fn_MSO_ExtractExcelRowNumber = cInt(mid(sCell,iCnt, len(sCell)))
	End If
end function
'***********************************	Function to extract column header name from cellname.  ********************************************************************
'Function Name			:				Fn_MSO_ExtractExcelColumnName  

'Return Value		    : 			Excel Column column header name  / False

'Examples				:		 Msgbox Fn_MSO_ExtractExcelRowNumber("AA4")
'History				:		

'Developer Name				Date				Rev. No.			Changes Done	
'-------------------------------------------------------------------------------------------------
'Koustubh					26/05/2011		       1.0				Created
'-------------------------------------------------------------------------------------------------
Public function Fn_MSO_ExtractExcelColumnName(sCell)
	GBL_FAILED_FUNCTION_NAME="Fn_MSO_ExtractExcelColumnName"
	dim iCnt, sChar
	Fn_MSO_ExtractExcelColumnName = False
	For iCnt = 1 to len(sCell)
		sChar = mid(sCell,iCnt,1)
		If  Asc(sChar) >= Asc("A") and Asc(sChar) <= Asc("Z") then
			' do nothing :)
		else
			' exit from loop
			Exit For
		end if
	Next
	if iCnt <> len(sCell)+1 then
		' column Name
		Fn_MSO_ExtractExcelColumnName = mid(sCell,1, iCnt -1)
		if Fn_MSO_ExtractExcelColumnName = "" then Fn_MSO_ExtractExcelColumnName = False
	end if

end function

'***********************************	Function to handle excel dialog.  ********************************************************************
'Function Name			:	Fn_MSO_ExcelDialogHandler  

'Return Value		    : 	True  / False

'Examples				:	Msgbox  Fn_MSO_ExcelDialogHandler("Teamcenter Extensions for Microsoft Office", "", "Yes")
'History				:		

'Developer Name				Date				Rev. No.			Changes Done	
'-------------------------------------------------------------------------------------------------
'Koustubh					31/05/2011		       1.0				Created
'-------------------------------------------------------------------------------------------------
Public function Fn_MSO_ExcelDialogHandler(sTitle, sTextMessage, sBtnName)
	GBL_FAILED_FUNCTION_NAME="Fn_MSO_ExcelDialogHandler"
	Dim objDialog
	Fn_MSO_ExcelDialogHandler=False
	Set objDialog = Window("MicrosoftExcel").Dialog("ExeclImportErrorDialog")
	If Window("MicrosoftExcel").Exist Then
			if sTitle <> "" then
				objDialog.SetTOProperty "text",sTitle
			end if
			'Click on button
			If Fn_UI_ObjectExist("Fn_MSO_ExcelDialogHandler", objDialog) Then
				Fn_MSO_ExcelDialogHandler = True
				IF sTextMessage <> "" THEN
					Fn_MSO_ExcelDialogHandler = False
					if sTextMessage = objDialog.Static("ErrorMsg").GetROProperty("text") Then
						Fn_MSO_ExcelDialogHandler = True
					End IF
				END IF
				Call Fn_UI_WinButton_Click("Fn_MSO_ExcelDialogHandler", objDialog, sBtnName,"","","")	
			End If
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_ExcelDialogHandler: Dialog verified successfully.")
	else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL:Fn_MSO_ExcelDialogHandler: Window does not exist")
			'ExitTest
	End If
	Set objDialog = nothing
End Function

'------------------------------------- Function to login to teamcenter using MS Office ---------
'Function Name		  :	  Fn_MSO_TeamcenterLogin

'Description			 :	 Function to login to teamcenter using MS Office

'Parameters			   :		  sAction : Action to perform
'								sUserName : User Name
'								sPassword : Password
'								sGroup : Group name
'								sRole  : Role

'Return Value		   : 		  True / False

'Examples				:	Call Fn_MSO_TeamcenterLogin("ExcelLogin", "x_watwe", "x_watwe", "", "" )
'Examples				:	Call Fn_MSO_TeamcenterLogin("WpfExcelLogin", "x_watwe", "x_watwe", "", "" )

'History					 :					
'		Developer Name			Date		Rev. No.	Changes Done										Reviewed By
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------
'		Koustubh Watwe		31/05/2011		1.0		
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------
'		Koustubh Watwe		1/06/2011		1.0			Added case WpfExcelLogin
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------
'		Vivek Ahirrao		10/06/2016		1.1			Added cases "WpfWordLogin", "WpfPowerPointLogin"	[TC1123-20160504-10_06_2016-VivekA-NewDevelopment]
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public function Fn_MSO_TeamcenterLogin(sAction, sUserName, sPassword, sGroup, sRole )
	GBL_FAILED_FUNCTION_NAME="Fn_MSO_TeamcenterLogin"
	Dim wshshell
	Fn_MSO_TeamcenterLogin = False
	'Added by Nilesh on 1 June 2012
'	If Window("MicrosoftExcel").Exist(5) Then
'		If Window("MicrosoftExcel").GetRoProperty("enabled")=True Then
'			Call Fn_MSO_Click_ImportToTeamCenter_Button()
'		End If
'	End If
	Select Case sAction
		Case "ExcelLogin"
				' handling 
				Call Fn_MSO_ExcelDialogHandler("Teamcenter Extensions for Microsoft Office", "", "Yes")
				If Fn_UI_ObjectExist("Fn_MSO_TeamcenterLogin", SwfWindow("MSO_Login")) Then
					'sUserName
					If sUserName <> "" Then
						SwfWindow("MSO_Login").SwfEdit("txtUserName").Set sUserName
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_TeamcenterLogin: User name [ " & sUserName & " ] set successfully.")
					End If
					'sPassword
					If sPassword <> "" Then
						SwfWindow("MSO_Login").SwfEdit("txtPassword").Set sPassword
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_TeamcenterLogin: Password [ " & sPassword & " ] set successfully.")
					End If
					'sGroup
					If sGroup <> "" Then
						SwfWindow("MSO_Login").SwfEdit("txtGroup").Set sGroup
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_TeamcenterLogin: Group [ " & sGroup & " ] set successfully.")
					End If
					'sRole
					If sRole <> "" Then
						SwfWindow("MSO_Login").SwfEdit("txtRole").Set sRole
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_TeamcenterLogin: Role [ " & sRole & " ] set successfully.")
					End If
					'clicking on OK
					Call Fn_UI_SwfButtonClick("Fn_MSO_TeamcenterLogin", SwfWindow("MSO_Login"), "OK")
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_TeamcenterLogin: Clicked on [ OK ] button successfully.")
					Fn_MSO_TeamcenterLogin = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_TeamcenterLogin: Dialog verified successfully.")
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_TeamcenterLogin: Dialog object is not visible.")
				End If
	
		Case "WpfExcelLogin"
				If Fn_UI_ObjectExist("Fn_MSO_TeamcenterLogin", WpfWindow("Teamcenter Login"))=False Then
					Call Fn_MSO_SetFocusOnApplicationWindow("MSExcel","")
					'Activate Teamcenter tab
					Call Fn_MSO_MainTabOperations("Activate","MSExcel","Teamcenter","")
					wait 1
					Call Fn_MSO_RibbonButton_Operations("MSExcel","Click","Current Settings:Login","")
					Wait 1
					If Fn_UI_ObjectExist("Fn_MSO_TeamcenterLogin", WpfWindow("Teamcenter Login"))=False Then
						Call Fn_MSO_RibbonButton_Operations("MSExcel","Click","Current Settings:Login","")
						Wait 1
					End If	
				End If
				If Fn_UI_ObjectExist("Fn_MSO_TeamcenterLogin", WpfWindow("Teamcenter Login")) Then
					Call Fn_UI_WpfButtonClick("Fn_MSO_TeamcenterLogin", WpfWindow("Teamcenter Login"), "Clear")
					'sUserName
					'Added calll of Fn_UI_WpfObjectClick in all WPF object operation by Nilesh on 1st June 2012
					If sUserName <> "" Then
						Call Fn_UI_WpfObjectClick("Fn_MSO_TeamcenterLogin",WpfWindow("Teamcenter Login").WpfEdit("UserID"))
						WpfWindow("Teamcenter Login").WpfEdit("UserID").Type sUserName
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_TeamcenterLogin: User name [ " & sUserName & " ] set successfully.")
					End If
					'sPassword
					If sPassword <> "" Then
						Call Fn_UI_WpfObjectClick("Fn_MSO_TeamcenterLogin",WpfWindow("Teamcenter Login").WpfEdit("Password"))
						WpfWindow("Teamcenter Login").WpfEdit("Password").Type sPassword
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_TeamcenterLogin: Password [ " & sPassword & " ] set successfully.")
					End If
					'sGroup
					If sGroup <> "" Then
                       			Call Fn_UI_WpfObjectClick("Fn_MSO_TeamcenterLogin",WpfWindow("Teamcenter Login").WpfEdit("Group"))
						WpfWindow("Teamcenter Login").WpfEdit("Group").Type sGroup
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_TeamcenterLogin: Group [ " & sGroup & " ] set successfully.")
					End If
					'sRole
					If sRole <> "" Then
						Call Fn_UI_WpfObjectClick("Fn_MSO_TeamcenterLogin",WpfWindow("Teamcenter Login").WpfEdit("Role"))
						WpfWindow("Teamcenter Login").WpfEdit("Role").Type sRole
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_TeamcenterLogin: Role [ " & sRole & " ] set successfully.")
					End If
					'clicking on OK
					Call Fn_UI_WpfButtonClick("Fn_MSO_TeamcenterLogin", WpfWindow("Teamcenter Login"), "Login")
					Wait 1
					Do While Fn_MSO_GetCursorType() = "65545"
						Wait 2
					Loop
'					Do While Fn_UI_ObjectExist("Fn_MSO_TeamcenterLogin", WpfWindow("Teamcenter Login")) = True
'						Wait 3
'					Loop
					Wait 5
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_TeamcenterLogin: Clicked on [ OK ] button successfully.")
					Fn_MSO_TeamcenterLogin = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_TeamcenterLogin: successfully entered Login data to Teamcenter Login window.")
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_TeamcenterLogin: Dialog object is not visible.")
				End If
		Case "WpfWordLogin"
				If Fn_UI_ObjectExist("Fn_MSO_TeamcenterLogin", WpfWindow("Teamcenter Login"))=False Then
					Call Fn_MSO_SetFocusOnApplicationWindow("MSWord","")
					'Activate Teamcenter tab
					Call Fn_MSO_MainTabOperations("Activate","MSWord","Teamcenter","")
					wait 1
					Call Fn_MSO_RibbonButton_Operations("MSWord","Click","Current Settings:Login","")
					Wait 1
					If Fn_UI_ObjectExist("Fn_MSO_TeamcenterLogin", WpfWindow("Teamcenter Login"))=False Then
						Call Fn_MSO_RibbonButton_Operations("MSWord","Click","Current Settings:Login","")
						Wait 1
					End If				
				End If
				If Fn_UI_ObjectExist("Fn_MSO_TeamcenterLogin", WpfWindow("Teamcenter Login")) Then
					Call Fn_UI_WpfButtonClick("Fn_MSO_TeamcenterLogin", WpfWindow("Teamcenter Login"), "Clear")
					'sUserName
					'Added calll of Fn_UI_WpfObjectClick in all WPF object operation by Nilesh on 1st June 2012
					If sUserName <> "" Then
						Call Fn_UI_WpfObjectClick("Fn_MSO_TeamcenterLogin",WpfWindow("Teamcenter Login").WpfEdit("UserID"))
						WpfWindow("Teamcenter Login").WpfEdit("UserID").Type sUserName
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_TeamcenterLogin: User name [ " & sUserName & " ] set successfully.")
					End If
					'sPassword
					If sPassword <> "" Then
						Call Fn_UI_WpfObjectClick("Fn_MSO_TeamcenterLogin",WpfWindow("Teamcenter Login").WpfEdit("Password"))
						WpfWindow("Teamcenter Login").WpfEdit("Password").Type sPassword
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_TeamcenterLogin: Password [ " & sPassword & " ] set successfully.")
					End If
					'sGroup
					If sGroup <> "" Then
                       			Call Fn_UI_WpfObjectClick("Fn_MSO_TeamcenterLogin",WpfWindow("Teamcenter Login").WpfEdit("Group"))
						WpfWindow("Teamcenter Login").WpfEdit("Group").Type sGroup
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_TeamcenterLogin: Group [ " & sGroup & " ] set successfully.")
					End If
					'sRole
					If sRole <> "" Then
						Call Fn_UI_WpfObjectClick("Fn_MSO_TeamcenterLogin",WpfWindow("Teamcenter Login").WpfEdit("Role"))
						WpfWindow("Teamcenter Login").WpfEdit("Role").Type sRole
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_TeamcenterLogin: Role [ " & sRole & " ] set successfully.")
					End If
					'clicking on OK
					Call Fn_UI_WpfButtonClick("Fn_MSO_TeamcenterLogin", WpfWindow("Teamcenter Login"), "Login")
					Wait 1
					Do While Fn_MSO_GetCursorType() = "65545"
						Wait 2
					Loop
'					Do While Fn_UI_ObjectExist("Fn_MSO_TeamcenterLogin", WpfWindow("Teamcenter Login")) = True
'						Wait 3
'					Loop
					Wait 5
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_TeamcenterLogin: Clicked on [ OK ] button successfully.")
					Fn_MSO_TeamcenterLogin = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_TeamcenterLogin: successfully entered Login data to Teamcenter Login window.")
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_TeamcenterLogin: Dialog object is not visible.")
				End If
				
		Case "WpfPowerPointLogin"
				If Fn_UI_ObjectExist("Fn_MSO_TeamcenterLogin", WpfWindow("Teamcenter Login"))=False Then
					Call Fn_MSO_SetFocusOnApplicationWindow("MSPowerPoint","")
					'Activate Teamcenter tab
					Call Fn_MSO_MainTabOperations("Activate","MSPowerPoint","Teamcenter","")
					wait 3
					Call Fn_MSO_RibbonButton_Operations("MSPowerPoint","Click","Current Settings:Login","")
					Wait 1
					If Fn_UI_ObjectExist("Fn_MSO_TeamcenterLogin", WpfWindow("Teamcenter Login"))=False Then
						Call Fn_MSO_RibbonButton_Operations("MSPowerPoint","Click","Current Settings:Login","")
						Wait 1
					End If				
				End If
				If Fn_UI_ObjectExist("Fn_MSO_TeamcenterLogin", WpfWindow("Teamcenter Login")) Then
					Call Fn_UI_WpfButtonClick("Fn_MSO_TeamcenterLogin", WpfWindow("Teamcenter Login"), "Clear")
					'sUserName
					'Added calll of Fn_UI_WpfObjectClick in all WPF object operation by Nilesh on 1st June 2012
					If sUserName <> "" Then
						Call Fn_UI_WpfObjectClick("Fn_MSO_TeamcenterLogin",WpfWindow("Teamcenter Login").WpfEdit("UserID"))
						WpfWindow("Teamcenter Login").WpfEdit("UserID").Type sUserName
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_TeamcenterLogin: User name [ " & sUserName & " ] set successfully.")
					End If
					'sPassword
					If sPassword <> "" Then
						Call Fn_UI_WpfObjectClick("Fn_MSO_TeamcenterLogin",WpfWindow("Teamcenter Login").WpfEdit("Password"))
						WpfWindow("Teamcenter Login").WpfEdit("Password").Type sPassword
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_TeamcenterLogin: Password [ " & sPassword & " ] set successfully.")
					End If
					'sGroup
					If sGroup <> "" Then
                       			Call Fn_UI_WpfObjectClick("Fn_MSO_TeamcenterLogin",WpfWindow("Teamcenter Login").WpfEdit("Group"))
						WpfWindow("Teamcenter Login").WpfEdit("Group").Type sGroup
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_TeamcenterLogin: Group [ " & sGroup & " ] set successfully.")
					End If
					'sRole
					If sRole <> "" Then
						Call Fn_UI_WpfObjectClick("Fn_MSO_TeamcenterLogin",WpfWindow("Teamcenter Login").WpfEdit("Role"))
						WpfWindow("Teamcenter Login").WpfEdit("Role").Type sRole
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_TeamcenterLogin: Role [ " & sRole & " ] set successfully.")
					End If
					'clicking on OK
					Call Fn_UI_WpfButtonClick("Fn_MSO_TeamcenterLogin", WpfWindow("Teamcenter Login"), "Login")
					Wait 1
					Do While Fn_MSO_GetCursorType() = "65545"
						Wait 2
					Loop
'					Do While Fn_UI_ObjectExist("Fn_MSO_TeamcenterLogin", WpfWindow("Teamcenter Login")) = True
'						Wait 3
'					Loop
					Wait 5
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_TeamcenterLogin: Clicked on [ OK ] button successfully.")
					Fn_MSO_TeamcenterLogin = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_TeamcenterLogin: successfully entered Login data to Teamcenter Login window.")
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_TeamcenterLogin: Dialog object is not visible.")
				End If
				
		Case else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL:Fn_MSO_TeamcenterLogin: Invalid case [ " & sAction & " ]")
	End Select
	If Fn_MSO_TeamcenterLogin = True Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Fn_MSO_TeamcenterLogin: Executed successfully with case [ " & sAction & " ].")
	End If
End Function
'------------------------------------- Function to import to teamcenter using MS Office ---------
'Function Name		  :	  Fn_MSO_ImportToTeamcenter

'Description			 :	 Function to import to teamcenter using MS Office

'Parameters			   :		  sAction : Action to perform
'								dicImportTeamcenter : Import Teamcenter dictionary object

'Return Value		   : 		  True / False
'								dicImportTeamcenter("UserID") = "AutoTestDBA"
'								dicImportTeamcenter("Password")= "AutoTestDBA"
'								dicImportTeamcenter("Group")=""
'								dicImportTeamcenter("Role")=""
'								dicImportTeamcenter("ControlFilePath") = "d:\testctrlfile.xlsx"
'								dicImportTeamcenter("ControlFileSheet") = ""
'								dicImportTeamcenter("bCloseExcel") = True
'								dicImportTeamcenter("ValidationErrorAllowed") = "50"
'								dicImportTeamcenter("sErrorMessages") = "Validation Error:[H7] - Property 'revision_list' on type 'Item' is read-only."""
'								dicImportTeamcenter("sErrorMessagesCount") = "5"
'								dicImportTeamcenter("sDialogError") = ""
'								dicImportTeamcenter("ImportLogFileName") = "c:\importlog.log"
'								dicImportTeamcenter("ImportLogMessage") = "Comparing structures and Importing file..." & VbLf & "Done!" 

'Examples				:	Call Fn_MSO_ImportToTeamcenter("ImportToTeamcenter", dicImportTeamcenter )
'Examples				:	Call Fn_MSO_ImportToTeamcenter("VerifyAndSaveReports", dicImportTeamcenter )
'Examples				:	Call Fn_MSO_ImportToTeamcenter("VerifyInvalidLogin", dicImportTeamcenter )
'Examples				:	Call Fn_MSO_ImportToTeamcenter("VerifyValidationError", dicImportTeamcenter )
'Examples				:	Call Fn_MSO_ImportToTeamcenter("AssignID", dicImportTeamcenter )
'Examples				:	Call Fn_MSO_ImportToTeamcenter("VerifySelectSheet", dicImportTeamcenter )
'Examples				:	Call Fn_MSO_ImportToTeamcenter("VerifyErrorMessage", dicImportTeamcenter )
'Examples				:	bReturn=Fn_MSO_ImportToTeamcenter("CancelOperation", dicImportTeamcenter )

'History					 :					
'		Developer Name				Date						Rev. No.				Changes Done				Reviewed By
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------
'		Koustubh Watwe			07/06/2011			           1.0		
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------
'		Koustubh Watwe			08/06/2011			           1.0						modified cases ImportToTeamcenter, VerifyInvalidLogin
'																						replced standard window objects to SwfWindow
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------
'		Koustubh Watwe			09/06/2011			           1.0		                Added code to close opened excel
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------
'		Koustubh Watwe			09/06/2011			           1.0		                Added case VerifyValidationError, 
'																					    Replaced Login code with function call
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------
'		Koustubh Watwe			15/06/2011			           1.0		                Added case VerifyAndSaveReports, AssignID
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------
'		Koustubh Watwe			15/06/2011			           1.0		                Added case VerifySelectSheet
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------
'		Koustubh Watwe			21/06/2011			           1.0		                Added case VerifyErrorMessage
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------
'		Koustubh Watwe			23/06/2011			           1.0		                modified cases ImportToTeamcenter, VerifyValidationError
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------
'		Sandeep Navghane		13/10/2011			           1.0		                Added case VerifyTabValues
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------
'		Koustubh Watwe			19/10/2011			           1.0		                Added case VerifyColumnMapping
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	  Shreyas Waichal				29-02-2012						1.1						Modified Code to click on  "Import to Teamcenter" button
'																															Added Case "CancelOperation"
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	 Nilesh Gadekar			01-05-2012						1.2						Added code to handle OR change on Windows 7 in some cases
'																															
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public function Fn_MSO_ImportToTeamcenter(sAction, dicImportTeamcenter)
		GBL_FAILED_FUNCTION_NAME="Fn_MSO_ImportToTeamcenter"
		Dim objWMIService, colItems, iXpos, iYpos, objItem, sResolution
        Dim sLeft,Top,sRight,sBottom,x,y,xAxis,yAxis,dr
		Dim objValidationError, aErrors, bFlag, iCnt, iCount, iErrorCnt, currVal
		Dim aItems, iItemCount, iNodeCnt
		Dim objDialog, objCMTree, aElements
		Fn_MSO_ImportToTeamcenter = False
		Set objExcel = GetObject(,"Excel.Application")
		objExcel.Visible = True
		Select Case sAction
			Case "AssignID"
				'Call Fn_KeyBoardOperation("SendKeys","%~Y~C~3~{ENTER}")
				'Call Fn_KeyBoardOperation("SendKeys","%~Y1~YF~{ENTER}")
				Call Fn_MSO_SetFocusOnApplicationWindow("MSExcel","")   'to set Focus on MSExcel
	                     wait 2
				Call Fn_KeyBoardOperation("SendKeys","%~Y2~YG~{ENTER}")
				wait 2
			Case Else
				'Wait For Synchronisation 
				Wait 3
				'Click on Import To Teamcenter Button
				 Call Fn_MSO_Click_ImportToTeamCenter_Button()
		End Select
		Select Case sAction
				' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
				Case "ImportToTeamcenter", "VerifyAndSaveReports", "VerifyErrorMessage", "VerifyColumnMapping","ErrorMessageVerify"
						' Import to Teamcenter login window
						If Fn_SISW_UI_Object_Operations("Fn_MSO_ImportToTeamcenter", "Exist", WpfWindow("Teamcenter Login"), "") = True Then
							If Fn_MSO_TeamcenterLogin("WpfExcelLogin", dicImportTeamcenter("UserID"), dicImportTeamcenter("Password"), dicImportTeamcenter("Group") , dicImportTeamcenter("Role") ) = True Then
								' select control file
								If Window("MicrosoftExcel").Window("FileOpen").exist(120) Then
									If dicImportTeamcenter("ControlFilePath") <> "" Then
	'									Window("MicrosoftExcel").Window("FileOpen").WinObject("FileName").Type dicImportTeamcenter("ControlFilePath") 
	'									Added by Nilesh  on 1st june 2012 to handle  Object change on windows 7
									   If Window("MicrosoftExcel").Window("FileOpen").WinObject("FileName").Exist(2) Then
											Window("MicrosoftExcel").Window("FileOpen").WinObject("FileName").Type dicImportTeamcenter("ControlFilePath") 
									   End If
									   If Window("MicrosoftExcel").Window("FileOpen").WinEdit("FileName").Exist(2) Then
											Window("MicrosoftExcel").Window("FileOpen").WinEdit("FileName").Set dicImportTeamcenter("ControlFilePath") 
									   End If
									'Code end
										Window("MicrosoftExcel").Window("FileOpen").Activate
										Set WshShell = CreateObject("WScript.Shell")
										WshShell.SendKeys "{ENTER}"
										Set WshShell = nothing
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_ImportToTeamcenter: Successfully opened Control file [ " & dicImportTeamcenter("ControlFilePath") & " ].")
									End If
								End If ' end of open control file file								
							End If
						'ElseIf SwfWindow("AssociateControlFile").Exist(10) Then
						ElseIf Fn_UI_ObjectExist("Fn_MSO_ImportToTeamcenter", SwfWindow("AssociateControlFile")) Then
							If dicImportTeamcenter("ControlFilePath") <> "" Then
								'SwfWindow("AssociateControlFile").SwfEdit("FilePath").Type dicImportTeamcenter("ControlFilePath") 
								Call Fn_UI_SwfButtonClick("Fn_MSO_ImportToTeamcenter", SwfWindow("AssociateControlFile"), "Browse")
								If Fn_UI_ObjectExist("Fn_MSO_ImportToTeamcenter", SwfWindow("AssociateControlFile").Window("SelectControlFile")) Then
									If dicImportTeamcenter("ControlFilePath") <> "" Then
										SwfWindow("AssociateControlFile").Window("SelectControlFile").WinObject("FileName").Type dicImportTeamcenter("ControlFilePath") 
										SwfWindow("AssociateControlFile").Window("SelectControlFile").Activate
										Set WshShell = CreateObject("WScript.Shell")
										WshShell.SendKeys "{ENTER}"
										Set WshShell = nothing
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_ImportToTeamcenter: Successfully opened Control file [ " & dicImportTeamcenter("ControlFilePath") & " ].")
									End If
								End If ' end of open control file file
								Call Fn_UI_SwfButtonClick("Fn_MSO_ImportToTeamcenter", SwfWindow("AssociateControlFile"), "Continue")
							End If
						Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_ImportToTeamcenter: Failed to perform Import To Teamcer operation.")
						End IF
						' selecting control file sheet.
						'If SwfWindow("Select Sheet").Exist(30)  then
						If Fn_UI_ObjectExist("Fn_MSO_ImportToTeamcenter", SwfWindow("Select Sheet")) Then
							If dicImportTeamcenter("ControlFileSheet") <> "" Then
								SwfWindow("Select Sheet").SwfComboBox("SheetNames").Select dicImportTeamcenter("ControlFileSheet")
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_ImportToTeamcenter: Successfully selected Control sheet [ " & dicImportTeamcenter("ControlFileSheet") & " ].")
							Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_ImportToTeamcenter: Successfully opened Default Control sheet.")
							End If
							Call Fn_UI_SwfButtonClick("Fn_MSO_ImportToTeamcenter", SwfWindow("Select Sheet"), "OK")
						Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_ImportToTeamcenter: Failed to open Control sheet.")
						End IF

						' Control File parsing is complete!
						'If Window("MicrosoftExcel").Dialog("ExeclImportErrorDialog").Exist(20) then
						If Fn_UI_ObjectExist("Fn_MSO_ImportToTeamcenter", Window("MicrosoftExcel").Dialog("ExeclImportErrorDialog")) Then
								'Window("MicrosoftExcel").Dialog("ExeclImportErrorDialog").WinButton("OK").Click 1,1,micLeftBtn
								Call Fn_UI_WinButton_Click("Fn_MSO_ImportToTeamcenter", Window("MicrosoftExcel").Dialog("ExeclImportErrorDialog"), "OK","","","")
						ElseIf Fn_UI_ObjectExist("Fn_MSO_ImportToTeamcenter", Dialog("ConfirmationBox")) Then
								Call Fn_UI_WinButton_Click("Fn_MSO_ImportToTeamcenter", Dialog("ConfirmationBox"), "OK","","","")
						ElseIf Fn_UI_ObjectExist("Fn_MSO_ImportToTeamcenter", Dialog("ExcelImport")) Then
								Call Fn_UI_WinButton_Click("Fn_MSO_ImportToTeamcenter", Dialog("ExcelImport"), "OK","","","")
						End If
						' cliclicking OK of Column Mapping window.
						If Fn_UI_ObjectExist("Fn_MSO_ImportToTeamcenter",SwfWindow("Column Mappings")) Then
								If sAction = "VerifyColumnMapping" then
									If dicImportTeamcenter("ColumnMapping") <> "" Then
										Set objCMTree = SwfWindow("Column Mappings").SwfTreeView("ColumnMappingTree")
										aElements = split(dicImportTeamcenter("ColumnMapping"),"~")
										For iCnt = 0 to uBound(aElements)
											aItems = split(aElements(iCnt),":")
											iCount = 0
											iNodeCnt = 0
											Fn_MSO_ImportToTeamcenter = False
											iItemCount = objCMTree.getROProperty("items count")
											Do until iNodeCnt >= iItemCount 
												If iCount = uBound(aItems) Then
													If trim(objCMTree.getItem(iNodeCnt)) = trim(aItems(iCount)) then
														Fn_MSO_ImportToTeamcenter = True
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_ImportToTeamcenter: [ " & aElements(iCnt) & " ] is present in Column Mapping tree.")
														Exit do
													End If
												ElseIf trim(objCMTree.getItem(iNodeCnt)) = trim(aItems(iCount)) then
													objCMTree.expand iNodeCnt
													wait 1
													iItemCount = objCMTree.getROProperty("items count")
													iCount = iCount + 1
												End If
												iNodeCnt = iNodeCnt + 1
											Loop
											If Fn_MSO_ImportToTeamcenter = False Then
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_ImportToTeamcenter: [ " & aElements(iCnt) & " ] is not present in Column Mapping tree.")
												Exit For
											End If
										Next
										Set objCMTree = nothing
									End If
								End If
								Call Fn_UI_SwfButtonClick("Fn_MSO_ImportToTeamcenter", SwfWindow("Column Mappings"), "OK")
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_ImportToTeamcenter: Successfully clicked on [ OK ] of Column Mapping window.")
						Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_ImportToTeamcenter: Failed to open Column Mapping window.")
						End If

						'Excel Import Wizard
						If Fn_UI_ObjectExist("Fn_MSO_ImportToTeamcenter",SwfWindow("Excel Import Wizard")) Then
								' Step 1
								'setting Create new revision checkbox
								If dicImportTeamcenter("CreateNewRevision") <> "" Then
									Select Case dicImportTeamcenter("CreateNewRevision")
										Case "True", True, "ON"
											SwfWindow("Excel Import Wizard").SwfCheckBox("Create new revision").Set "ON"
										Case "False", false, "OFF"
											SwfWindow("Excel Import Wizard").SwfCheckBox("Create new revision").Set "OFF"
									End Select
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_ImportToTeamcenter: Successfully set Create new revision to [ " & dicImportTeamcenter("CreateNewRevision") & " ].")
								End If
								'setting number of validation error allowed
								If dicImportTeamcenter("ValidationErrorAllowed") <> "" Then
									SwfWindow("Excel Import Wizard").SwfSpin("NoOfErrorsAllowed").Set trim(dicImportTeamcenter("ValidationErrorAllowed")) 
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_ImportToTeamcenter: Successfully set Number of Validation Errors allowed as [ " & dicImportTeamcenter("ValidationErrorAllowed") & " ].")
								End If

								Call Fn_UI_SwfButtonClick("Fn_MSO_ImportToTeamcenter", SwfWindow("Excel Import Wizard"), "Next")
								' checking Warning message
								If Fn_UI_ObjectExist("Fn_MSO_ImportToTeamcenter",SwfWindow("Excel Import Wizard").Dialog("ConfirmationBox")) Then
										If dicImportTeamcenter("sDialogError") <> "" Then
											If trim(SwfWindow("Excel Import Wizard").Dialog("ConfirmationBox").Static("TextMessage").GetROProperty("text")) <> dicImportTeamcenter("sDialogError") Then
													Fn_MSO_ImportToTeamcenter = False
													Exit function
											End If
										End If
										Call Fn_UI_WinButton_Click("Fn_MSO_ImportToTeamcenter", SwfWindow("Excel Import Wizard").Dialog("ConfirmationBox"),"Yes", "","","")
								End If
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_ImportToTeamcenter: Successfully clicked on [ Next ] of Excel Import Wizard [ Step 1 ].")
								SwfWindow("Excel Import Wizard").SwfLabel("StatusLabel").WaitProperty "visible", true, 15000
								' Step 2
								Call Fn_UI_SwfButtonClick("Fn_MSO_ImportToTeamcenter", SwfWindow("Excel Import Wizard"), "Next")
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_ImportToTeamcenter: Successfully clicked on [ Next ] of Excel Import Wizard [ Step 2 ].")
								SwfWindow("Excel Import Wizard").SwfLabel("StatusLabel").WaitProperty "visible", true, 15000
								' Step 3
								Call Fn_UI_SwfButtonClick("Fn_MSO_ImportToTeamcenter", SwfWindow("Excel Import Wizard"), "Next")
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_ImportToTeamcenter: Successfully clicked on [ Next ] of Excel Import Wizard [ Step 3 ].")
								
								'Step 4
								Select Case sAction

									Case "ErrorMessageVerify"

										If SwfWindow("Excel Import Wizard").Dialog("ConfirmationBox").Exist then
												If dicImportTeamcenter("sButton") <> "" Then
													 SwfWindow("Excel Import Wizard").Dialog("ConfirmationBox").WinButton("Yes").SetTOProperty  "text","&"+dicImportTeamcenter("sButton")
'													Call Fn_UI_SwfButtonClick("Fn_MSO_ImportToTeamcenter", SwfWindow("Excel Import Wizard").Dialog("ConfirmationBox"),dicImportTeamcenter("sButton"))
										SwfWindow("Excel Import Wizard").Dialog("ConfirmationBox").WinButton("Yes").Click 5,5

											Fn_MSO_ImportToTeamcenter = True
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_ImportToTeamcenter: Successfully Clicked on the button [ " & dicImportTeamcenter("sButton") & " ].")
												End If
										End If
									'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
									' Normal Import
									Case "ImportToTeamcenter"
											SwfWindow("Excel Import Wizard").SwfLabel("FeedbackMessage").SetTOProperty "text", "Excel import process completed successfully."
                                            Wait 5
											SwfWindow("Excel Import Wizard").SwfLabel("FeedbackMessage").WaitProperty "visible", true, 36000 ''Added by Avinash J on 130 build -18-feb-13
											SwfWindow("Excel Import Wizard").SwfLabel("FeedbackMessage").WaitProperty "visible", true, 24000
											
											For iCount = 1 To 10 Step 1
												If Fn_UI_ObjectExist("Fn_MSO_ImportToTeamcenter", SwfWindow("Excel Import Wizard").SwfLabel("FeedbackMessage"))=False Then
														SwfWindow("Excel Import Wizard").SwfLabel("FeedbackMessage").WaitProperty "visible", true, 3000
												Else
													Exit For 
												End If
											Next
											
											If Fn_UI_ObjectExist("Fn_MSO_ImportToTeamcenter", SwfWindow("Excel Import Wizard").SwfLabel("FeedbackMessage")) Then
													Fn_MSO_ImportToTeamcenter = True
											End If
									'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
									' Verify and Save Reports
									Case "VerifyAndSaveReports"
											' Verifying Log
											If dicImportTeamcenter("ImportLogMessage") <> "" Then
												wait 5
'												If trim(dicImportTeamcenter("ImportLogMessage")) = trim(SwfWindow("Excel Import Wizard").SwfEditor("FeedbackMessageDetails").GetROProperty("text")) Then
												'Modified beacuse log contains actual mesaage + detailed log by Nilesh on 1st June 2012
												If Instr(trim(dicImportTeamcenter("ImportLogMessage")) , trim(SwfWindow("Excel Import Wizard").SwfEditor("FeedbackMessageDetails").GetROProperty("text")) )>0Then
													Fn_MSO_ImportToTeamcenter = True
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_ImportToTeamcenter: Successfully verified Feedback report with [ " & dicImportTeamcenter("ImportLogMessage") & " ].")
												End IF
											End If
											' Save Report Log
											If dicImportTeamcenter("ImportLogFileName") <> "" Then
												' clicking on save report button
												Call Fn_UI_SwfButtonClick("Fn_MSO_ImportToTeamcenter", SwfWindow("Excel Import Wizard"), "SaveReport")
												' check existance of Save Report Log dialog
												If SwfWindow("Excel Import Wizard").Dialog("SaveAs").Exist(15) Then
													' typing import log file path
'													SwfWindow("Excel Import Wizard").Dialog("SaveAs").WinEdit("FileName").Type dicImportTeamcenter("ImportLogFileName")
'														Added to work on Windows 7 by Nilesh
													If SwfWindow("Excel Import Wizard").Dialog("SaveAs").WinEdit("FileName").Exist(5) Then
														SwfWindow("Excel Import Wizard").Dialog("SaveAs").WinEdit("FileName").Type dicImportTeamcenter("ImportLogFileName")
													End If
													
													If SwfWindow("Excel Import Wizard").Dialog("SaveAs").WinEdit("FileName_1").Exist(5) Then
														SwfWindow("Excel Import Wizard").Dialog("SaveAs").WinEdit("FileName_1").Type dicImportTeamcenter("ImportLogFileName")
													End If
	'												End

													' clicking on save button
													'SwfWindow("Excel Import Wizard").Dialog("SaveAs").WinButton("Save").Click 1,1,micLeftBtn
													Call Fn_UI_WinButton_Click("Fn_MSO_ImportToTeamcenter", SwfWindow("Excel Import Wizard").Dialog("SaveAs"), "Save","","","")
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_ImportToTeamcenter: Successfully saved Feedback report in file  [ " & dicImportTeamcenter("ImportLogFileName") & " ].")
												End IF
											End If
									'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
									Case "VerifyErrorMessage"
											If Fn_UI_ObjectExist("Fn_MSO_ImportToTeamcenter",SwfWindow("Excel Import Wizard").Dialog("ConfirmationBox")) Then
												If dicImportTeamcenter("sDialogError") <> "" Then
													If trim(SwfWindow("Excel Import Wizard").Dialog("ConfirmationBox").Static("TextMessage").GetROProperty("text")) = dicImportTeamcenter("sDialogError") Then
															Fn_MSO_ImportToTeamcenter = True
															Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_ImportToTeamcenter: Successfully verified error message  [ " & dicImportTeamcenter("sDialogError") & " ].")
													End If
												End If
													If SwfWindow("Excel Import Wizard").Dialog("ConfirmationBox").WinButton("OK").Exist Then
														Call Fn_UI_WinButton_Click("Fn_MSO_ImportToTeamcenter", SwfWindow("Excel Import Wizard").Dialog("ConfirmationBox"),"OK","","","")
													Elseif SwfWindow("Excel Import Wizard").Dialog("ConfirmationBox").WinButton("Yes").Exist then
														Call Fn_UI_WinButton_Click("Fn_MSO_ImportToTeamcenter", SwfWindow("Excel Import Wizard").Dialog("ConfirmationBox"),"Yes","","","")
													Else
														'Will be coded as required
													End If
											ElseIf Fn_UI_ObjectExist("Fn_MSO_ImportToTeamcenter", Dialog("ConfirmationBox")) Then
														If dicImportTeamcenter("sDialogError") <> "" Then
															If trim(Dialog("ConfirmationBox").Static("TextMessage").GetROProperty("text")) = trim(dicImportTeamcenter("sDialogError")) Then
																Fn_MSO_ImportToTeamcenter = True
																Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_ImportToTeamcenter: Successfully verified error message  [ " & dicImportTeamcenter("sDialogError") & " ].")
															End If
															Call Fn_UI_WinButton_Click("Fn_MSO_ImportToTeamcenter", Dialog("ConfirmationBox"),"OK","","","")
														End If
											ElseIf Fn_UI_ObjectExist("Fn_MSO_ImportToTeamcenter", Dialog("ExcelImport")) Then
														If dicImportTeamcenter("sDialogError") <> "" Then
															If trim(Dialog("ExcelImport").Static("TextMessage").GetROProperty("text")) = trim(dicImportTeamcenter("sDialogError")) Then
																Fn_MSO_ImportToTeamcenter = True
																Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_ImportToTeamcenter: Successfully verified error message  [ " & dicImportTeamcenter("sDialogError") & " ].")
															ElseIf trim(Dialog("ExcelImport").Static("TextMessage").GetROProperty("text")) = trim(dicImportTeamcenter("sDialogPermissionError"))  Then
																Fn_MSO_ImportToTeamcenter = True
																Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_ImportToTeamcenter: Successfully verified error message  [ " & dicImportTeamcenter("sDialogPermissionError") & " ].")
															End If
														Call Fn_UI_WinButton_Click("Fn_MSO_ImportToTeamcenter", Dialog("ExcelImport"),"OK","","","")
														End If
											End If
									'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
									Case "VerifyColumnMapping"
											SwfWindow("Excel Import Wizard").SwfLabel("FeedbackMessage").SetTOProperty "text", "Excel import process completed successfully."
											SwfWindow("Excel Import Wizard").SwfLabel("FeedbackMessage").WaitProperty "visible", true, 15000
									'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -									
								End Select
								Wait 5
								SwfWindow("Excel Import Wizard").SwfLabel("StatusLabel").WaitProperty "visible", true, 15000
								SwfWindow("Excel Import Wizard").SwfButton("Finish").WaitProperty "enabled",True,240000
								If  SwfWindow("Excel Import Wizard").SwfButton("Finish").Exist Then
									Call Fn_UI_SwfButtonClick("Fn_MSO_ImportToTeamcenter", SwfWindow("Excel Import Wizard"), "Finish")
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_ImportToTeamcenter: Successfully clicked on [ Next ] of Excel Import Wizard [ Step 4 ].")
								End If
						Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_ImportToTeamcenter: Failed to open Import Wizard.")
								Exit function
						End If
				' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
				'Case "VerifyInvalidLogin"
				'Extra case commented by sonal [06-07-2012]
				Case "VerifyInvalidLogin"
						If Fn_MSO_TeamcenterLogin("WpfExcelLogin", dicImportTeamcenter("UserID"), dicImportTeamcenter("Password"), dicImportTeamcenter("Group") , dicImportTeamcenter("Role") ) = False then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_ImportToTeamcenter: Failed to perform Import To Teamcer operation.")
							Exit function
						End IF
						' verify error message
						If Fn_UI_ObjectExist("Fn_MSO_ImportToTeamcenter", WpfWindow("Teamcenter Login").WpfObject("LoginError")) Then
								If dicImportTeamcenter("sErrorMessages") <> "" Then
									If trim(dicImportTeamcenter("sErrorMessages")) = trim(WpfWindow("Teamcenter Login").WpfObject("LoginError").GetROProperty("text")) Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_ImportToTeamcenter: Successfully verified error message[ " & trim(dicImportTeamcenter("sErrorMessages")) & " ].")
										Fn_MSO_ImportToTeamcenter = True
									End If
								Else
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_ImportToTeamcenter: Successfully verified existance of  error message.")
										Fn_MSO_ImportToTeamcenter = True
								End If
						End If
						' click on cancel
						Call Fn_UI_WpfButtonClick("Fn_MSO_ImportToTeamcenter", WpfWindow("Teamcenter Login"), "Cancel")

						' handle error dialog window  
						Call Fn_MSO_ExcelDialogHandler("Excel Import", "", "No")
				' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
				Case "VerifyValidationError"
						' Import to Teamcenter login window
						If Fn_MSO_TeamcenterLogin("WpfExcelLogin", dicImportTeamcenter("UserID"), dicImportTeamcenter("Password"), dicImportTeamcenter("Group") , dicImportTeamcenter("Role") ) = True Then
							' select control file
							If Window("MicrosoftExcel").Window("FileOpen").exist(120) Then
								If dicImportTeamcenter("ControlFilePath") <> "" Then
'									Window("MicrosoftExcel").Window("FileOpen").WinObject("FileName").Type dicImportTeamcenter("ControlFilePath") 
									'Added by Nilesh to handle Or change on Windows 7 on 1st June 2012
									If Window("MicrosoftExcel").Window("FileOpen").WinObject("FileName").Exist(5) Then
										Window("MicrosoftExcel").Window("FileOpen").WinObject("FileName").Type dicImportTeamcenter("ControlFilePath") 
									End If
									If Window("MicrosoftExcel").Window("FileOpen").WinEdit("FileName").Exist(5) Then
										Window("MicrosoftExcel").Window("FileOpen").WinEdit("FileName").Set dicImportTeamcenter("ControlFilePath") 
									End If
									'End 

									Set WshShell = CreateObject("WScript.Shell")
									WshShell.SendKeys "{ENTER}"
									Set WshShell = nothing
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_ImportToTeamcenter: Successfully opened Control file [ " & dicImportTeamcenter("ControlFilePath") & " ].")
								End If
							End If ' end of open control file file
						ElseIf Fn_UI_ObjectExist("Fn_MSO_ImportToTeamcenter", SwfWindow("AssociateControlFile")) Then
							If dicImportTeamcenter("ControlFilePath") <> "" Then
								'SwfWindow("AssociateControlFile").SwfEdit("FilePath").Type dicImportTeamcenter("ControlFilePath") 
								Call Fn_UI_SwfButtonClick("Fn_MSO_ImportToTeamcenter", SwfWindow("AssociateControlFile"), "Browse")
								If Fn_UI_ObjectExist("Fn_MSO_ImportToTeamcenter", SwfWindow("AssociateControlFile").Window("SelectControlFile")) Then
									If dicImportTeamcenter("ControlFilePath") <> "" Then
										SwfWindow("AssociateControlFile").Window("SelectControlFile").WinObject("FileName").Type dicImportTeamcenter("ControlFilePath") 
										Set WshShell = CreateObject("WScript.Shell")
										WshShell.SendKeys "{ENTER}"
										Set WshShell = nothing
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_ImportToTeamcenter: Successfully opened Control file [ " & dicImportTeamcenter("ControlFilePath") & " ].")
									End If
								End If ' end of open control file file
				'' ---------------Code added for changed hierarchy for 'SelectControlFile'-----------by Sagar 19 Nov 12
								If Fn_UI_ObjectExist("Fn_MSO_ImportToTeamcenter",SwfWindow("AssociateControlFile").Dialog("SelectControlFile") ) Then
									If dicImportTeamcenter("ControlFilePath") <> "" Then
										SwfWindow("AssociateControlFile").Dialog("SelectControlFile").WinEdit("FileName").Set dicImportTeamcenter("ControlFilePath") 
										Wait(1)
										SwfWindow("AssociateControlFile").Dialog("SelectControlFile").WinButton("Open").Click
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_ImportToTeamcenter: Successfully opened Control file [ " & dicImportTeamcenter("ControlFilePath") & " ].")
									End If
								End If ' end of open control file file
			''' ---------------End of Code added for changed hierarchy for 'SelectControlFile'-----------by Sagar
								Call Fn_UI_SwfButtonClick("Fn_MSO_ImportToTeamcenter", SwfWindow("AssociateControlFile"), "Continue")
							End If
						Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_ImportToTeamcenter: Failed to perform Import To Teamcer operation.")
						End IF
						' selecting control file sheet.
						If Fn_UI_ObjectExist("Fn_MSO_ImportToTeamcenter", SwfWindow("Select Sheet")) Then
							If dicImportTeamcenter("ControlFileSheet") <> "" Then
								SwfWindow("Select Sheet").SwfComboBox("SheetNames").Select dicImportTeamcenter("ControlFileSheet")
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_ImportToTeamcenter: Successfully selected Control sheet [ " & dicImportTeamcenter("ControlFileSheet") & " ].")
							Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_ImportToTeamcenter: Successfully opened Default Control sheet.")
							End If
							Call Fn_UI_SwfButtonClick("Fn_MSO_ImportToTeamcenter", SwfWindow("Select Sheet"), "OK")
						Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_ImportToTeamcenter: Failed to open Control sheet.")
						End IF

						' Control File parsing is complete!
						If Fn_UI_ObjectExist("Fn_MSO_ImportToTeamcenter", Window("MicrosoftExcel").Dialog("ExeclImportErrorDialog")) Then
								Call Fn_UI_WinButton_Click("Fn_MSO_ImportToTeamcenter", Window("MicrosoftExcel").Dialog("ExeclImportErrorDialog"), "OK","","","")
						Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_ImportToTeamcenter: Import Excel dialog is not present.")
						End IF
						Wait(2)
						' cliclicking OK of Column Mapping window.
						If Fn_UI_ObjectExist("Fn_MSO_ImportToTeamcenter", SwfWindow("Column Mappings")) Then
								Call Fn_UI_SwfButtonClick("Fn_MSO_ImportToTeamcenter", SwfWindow("Column Mappings"), "OK")
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_ImportToTeamcenter: Successfully clicked on [ OK ] of Column Mapping window.")
						Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_ImportToTeamcenter: Column Mapping window is not present.")
						End If

						'Excel Import Wizard
						If Fn_UI_ObjectExist("Fn_MSO_ImportToTeamcenter", SwfWindow("Excel Import Wizard")) Then
								' Step 1
								'setting number of validation error allowed
								If dicImportTeamcenter("ValidationErrorAllowed") <> "" Then
									SwfWindow("Excel Import Wizard").SwfSpin("NoOfErrorsAllowed").Set trim(dicImportTeamcenter("ValidationErrorAllowed")) 
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_ImportToTeamcenter: Successfully set Number of Validation Errors allowed as [ " & dicImportTeamcenter("ValidationErrorAllowed") & " ].")
								End If
								Call Fn_UI_SwfButtonClick("Fn_MSO_ImportToTeamcenter", SwfWindow("Excel Import Wizard"), "Next")
								SwfWindow("Excel Import Wizard").SwfLabel("StatusLabel").WaitProperty "visible", true, 5000
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_ImportToTeamcenter: Successfully clicked on [ Next ] of Excel Import Wizard [ Step 1 ].")
						End If
						wait 3
						' checking error / warning error message
						If Fn_UI_ObjectExist("Fn_MSO_ImportToTeamcenter",SwfWindow("Excel Import Wizard").Dialog("ConfirmationBox")) Then
							If dicImportTeamcenter("sDialogError") <> "" Then
								bFlag = False
								If trim(SwfWindow("Excel Import Wizard").Dialog("ConfirmationBox").Static("TextMessage").GetROProperty("text")) = dicImportTeamcenter("sDialogError") Then
										bFlag = True
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_ImportToTeamcenter: Successfully verified error message  [ " & dicImportTeamcenter("sDialogError") & " ].")
								End If
								If bFlag = False Then
										Fn_MSO_ImportToTeamcenter = False
										Exit function
								End If
							End If
							If SwfWindow("Excel Import Wizard").Dialog("ConfirmationBox").WinButton("OK").Exist(3) then
									Call Fn_UI_WinButton_Click("Fn_MSO_ImportToTeamcenter", SwfWindow("Excel Import Wizard").Dialog("ConfirmationBox"),"OK","","","")
							ElseIf SwfWindow("Excel Import Wizard").Dialog("ConfirmationBox").WinButton("No").Exist(3) then
									Call Fn_UI_WinButton_Click("Fn_MSO_ImportToTeamcenter", SwfWindow("Excel Import Wizard").Dialog("ConfirmationBox"),"No","","","")
							End If
						End If
						wait 3
						If Fn_UI_ObjectExist("Fn_MSO_ImportToTeamcenter", SwfWindow("Excel Import Wizard")) Then
								' Step 2
								Call Fn_UI_SwfButtonClick("Fn_MSO_ImportToTeamcenter", SwfWindow("Excel Import Wizard"), "Next")
								SwfWindow("Excel Import Wizard").SwfLabel("StatusLabel").WaitProperty "visible", true, 15000
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_ImportToTeamcenter: Successfully clicked on [ Next ] of Excel Import Wizard [ Step 2 ].")
						End If
						wait 3						
						If Fn_UI_ObjectExist("Fn_MSO_ImportToTeamcenter", SwfWindow("Excel Import Wizard")) Then
								' Step 3
								Call Fn_UI_SwfButtonClick("Fn_MSO_ImportToTeamcenter", SwfWindow("Excel Import Wizard"), "Next")
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_ImportToTeamcenter: Successfully clicked on [ Next ] of Excel Import Wizard [ Step 3 ].")
						End If
						wait 3						
						If Fn_UI_ObjectExist("Fn_MSO_ImportToTeamcenter", SwfWindow("Excel Import Wizard")) Then
								'Step 4
								Call Fn_UI_SwfButtonClick("Fn_MSO_ImportToTeamcenter", SwfWindow("Excel Import Wizard"), "Finish")
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_ImportToTeamcenter: Successfully clicked on [ Next ] of Excel Import Wizard [ Step 4 ].")
						End If
						'verification in validation error dialog
						Set objValidationError = SwfWindow("ValidationError")
						'verifying existance of validation error dialog
						If Fn_UI_ObjectExist("Fn_MSO_ImportToTeamcenter", objValidationError) Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_ImportToTeamcenter: Successfully opened Validation Error dialog.")
							If dicImportTeamcenter("sErrorMessagesCount") <> "" Then
								bFlag = False
								If cInt(dicImportTeamcenter("sErrorMessagesCount")) = cInt(objValidationError.SwfList("ErrorsList").GetItemsCount) then
									bFlag = True
								End IF
							End IF
							If dicImportTeamcenter("sErrorMessages") <> "" Then
								aErrors = split(dicImportTeamcenter("sErrorMessages"),"~")
								 iErrorCnt = cInt(objValidationError.SwfList("ErrorsList").GetItemsCount)
								For iCnt = 0 to UBound(aErrors)
									bFlag = False
									For iCount = 0 to iErrorCnt - 1
										If trim(objValidationError.SwfList("ErrorsList").GetItem(iCount)) = trim(aErrors(iCnt)) then
											bFlag = True
											objValidationError.SwfList("ErrorsList").Select iCount
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_ImportToTeamcenter: [ " & trim(aErrors(iCnt)) & " ] is present in Validation Error List.")
											Exit For
										End if
									Next
									If bFlag = False Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_ImportToTeamcenter: [ " & trim(aErrors(iCnt)) & " ] is not present in Validation Error List.")
										Set objValidationError = nothing
										Exit function
									End If
								Next
							End If
							'clicking on cancel button
							Call Fn_UI_SwfButtonClick("Fn_MSO_ImportToTeamcenter", objValidationError, "Cancel")
						Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_ImportToTeamcenter: Failed to open Validation Error dialog.")
						End If		
						If bFlag Then
							Fn_MSO_ImportToTeamcenter = True
						End If	
				' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
				Case "AssignID"
						if dicImportTeamcenter("UserID") <> "" Then
							If Fn_MSO_TeamcenterLogin("WpfExcelLogin", dicImportTeamcenter("UserID"), dicImportTeamcenter("Password"), dicImportTeamcenter("Group") , dicImportTeamcenter("Role") ) = True Then
								' Control File parsing is complete!
								If WpfWindow("Tc Status").Dialog("ConfirmationBox").Exist(15) then
									Set objDialog = WpfWindow("Tc Status").Dialog("ConfirmationBox")
								ElseIf Dialog("ConfirmationBox").Exist(15) then
									Set objDialog = Dialog("ConfirmationBox")
								ElseIf Window("MicrosoftExcel").Dialog("ExeclImportErrorDialog").Exist(15) Then
									Set objDialog = Window("MicrosoftExcel").Dialog("ExeclImportErrorDialog")
								Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_ImportToTeamcenter: Failed to handle Import teamcenter confirmation window.")
								End If
								Call Fn_UI_WinButton_Click("Fn_MSO_ImportToTeamcenter", objDialog, "OK","","","")
								
								' cliclicking OK of Column Mapping window.
								If SwfWindow("Column Mappings").exist(30) then
									Call Fn_UI_SwfButtonClick("Fn_MSO_ImportToTeamcenter", SwfWindow("Column Mappings"), "OK")
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_ImportToTeamcenter: Successfully clicked on [ OK ] of Column Mapping window.")
								Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_ImportToTeamcenter: Failed to open Column Mapping window.")
								End If
								Set objDialog = nothing
							Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_ImportToTeamcenter: Failed to login to teamcenter.")
							End IF
						End IF
						Fn_MSO_ImportToTeamcenter = True
				' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
				Case "VerifySelectSheet"
						' Import to Teamcenter login window
						If Fn_MSO_TeamcenterLogin("WpfExcelLogin", dicImportTeamcenter("UserID"), dicImportTeamcenter("Password"), dicImportTeamcenter("Group") , dicImportTeamcenter("Role") ) = True Then
							' select control file
							If Window("MicrosoftExcel").Window("FileOpen").exist(120) Then
								If dicImportTeamcenter("ControlFilePath") <> "" Then
'									Window("MicrosoftExcel").Window("FileOpen").WinObject("FileName").Type dicImportTeamcenter("ControlFilePath") 
									'Added by Nilesh to handle Object change on Windows 7
                                    If Window("MicrosoftExcel").Window("FileOpen").WinObject("FileName").Exist(5) Then
										Window("MicrosoftExcel").Window("FileOpen").WinObject("FileName").Type dicImportTeamcenter("ControlFilePath") 
									End If
									If Window("MicrosoftExcel").Window("FileOpen").WinEdit("FileName").Exist(5) Then
										Window("MicrosoftExcel").Window("FileOpen").WinEdit("FileName").Set dicImportTeamcenter("ControlFilePath") 
									End If
									'End
									Set WshShell = CreateObject("WScript.Shell")
									WshShell.SendKeys "{ENTER}"
									Set WshShell = nothing
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_ImportToTeamcenter: Successfully opened Control file [ " & dicImportTeamcenter("ControlFilePath") & " ].")
								End If
							Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_ImportToTeamcenter: Failed to open Control file.")
							End If ' end of open control file file
						ElseIf Fn_UI_ObjectExist("Fn_MSO_ImportToTeamcenter", SwfWindow("AssociateControlFile")) Then
							If dicImportTeamcenter("ControlFilePath") <> "" Then
								'SwfWindow("AssociateControlFile").SwfEdit("FilePath").Type dicImportTeamcenter("ControlFilePath") 
								Call Fn_UI_SwfButtonClick("Fn_MSO_ImportToTeamcenter", SwfWindow("AssociateControlFile"), "Browse")
'								If Fn_UI_ObjectExist("Fn_MSO_ImportToTeamcenter", SwfWindow("AssociateControlFile").Window("SelectControlFile")) Then
									If dicImportTeamcenter("ControlFilePath") <> "" Then
'										SwfWindow("AssociateControlFile").Window("SelectControlFile").WinObject("FileName").Type dicImportTeamcenter("ControlFilePath") 
										If SwfWindow("AssociateControlFile").Window("SelectControlFile").Exist(5) Then
											SwfWindow("AssociateControlFile").Window("SelectControlFile").WinObject("FileName").Type dicImportTeamcenter("ControlFilePath") 
										End If

										If SwfWindow("AssociateControlFile").Dialog("SelectControlFile").Exist(5) Then
											SwfWindow("AssociateControlFile").Dialog("SelectControlFile").WinEdit("FileName").Set dicImportTeamcenter("ControlFilePath") 
										End If
'                                   
										Set WshShell = CreateObject("WScript.Shell")
										WshShell.SendKeys "{ENTER}"
										Set WshShell = nothing
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_ImportToTeamcenter: Successfully opened Control file [ " & dicImportTeamcenter("ControlFilePath") & " ].")
									End If
'								End If ' end of open control file file
								Call Fn_UI_SwfButtonClick("Fn_MSO_ImportToTeamcenter", SwfWindow("AssociateControlFile"), "Continue")
							End If
						Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_ImportToTeamcenter: Failed to perform Import To Teamcer operation.")
						End IF
						' selecting control file sheet.
						If Fn_UI_ObjectExist("Fn_MSO_ImportToTeamcenter", SwfWindow("Select Sheet")) Then
							If dicImportTeamcenter("ControlFileSheet") <> "" Then
								aItems = split(dicImportTeamcenter("ControlFileSheet"),"~")
								iItemCount = cInt(SwfWindow("Select Sheet").SwfComboBox("SheetNames").GetItemsCount)
								
								For iCnt = 0 to uBound(aItems)
									bFlag = False
									For iCount = 0 to iItemCount -1
										If trim(SwfWindow("Select Sheet").SwfComboBox("SheetNames").GetItem(iCount)) = aItems(icnt) Then
											bFlag = True
											Exit For
										End If
									Next
									If bFlag = False Then 
										Exit For
									End If
								Next
								Fn_MSO_ImportToTeamcenter = bFlag
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_ImportToTeamcenter: Successfully selected Control sheet [ " & dicImportTeamcenter("ControlFileSheet") & " ].")
							Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_ImportToTeamcenter: Successfully opened Default Control sheet.")
							End If
							Call Fn_UI_SwfButtonClick("Fn_MSO_ImportToTeamcenter", SwfWindow("Select Sheet"), "Cancel")
						Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_ImportToTeamcenter: Failed to open Control sheet.")
						End IF
				' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - Added by Sandeep 
				Case "VerifyTabValues"
						Fn_MSO_ImportToTeamcenter = False
						' Import to Teamcenter login window
						If Fn_MSO_TeamcenterLogin("WpfExcelLogin", dicImportTeamcenter("UserID"), dicImportTeamcenter("Password"), dicImportTeamcenter("Group") , dicImportTeamcenter("Role") ) = True Then
							' select control file
							If Window("MicrosoftExcel").Window("FileOpen").exist(120) Then
								If dicImportTeamcenter("ControlFilePath") <> "" Then
'									Window("MicrosoftExcel").Window("FileOpen").WinObject("FileName").Type dicImportTeamcenter("ControlFilePath") 
									'Added by Nilesh to handle OR change on Windows 7
									If Window("MicrosoftExcel").Window("FileOpen").WinObject("FileName").Exist(5) Then
										Window("MicrosoftExcel").Window("FileOpen").WinObject("FileName").Type dicImportTeamcenter("ControlFilePath") 
									End If
									If Window("MicrosoftExcel").Window("FileOpen").WinEdit("FileName").Exist(5) Then
										Window("MicrosoftExcel").Window("FileOpen").WinEdit("FileName").Set dicImportTeamcenter("ControlFilePath") 
									End If
								'End
									Set WshShell = CreateObject("WScript.Shell")
									WshShell.SendKeys "{ENTER}"
									Set WshShell = nothing
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_ImportToTeamcenter: Successfully opened Control file [ " & dicImportTeamcenter("ControlFilePath") & " ].")
								End If
							End If ' end of open control file file
						'ElseIf SwfWindow("AssociateControlFile").Exist(10) Then
						ElseIf Fn_UI_ObjectExist("Fn_MSO_ImportToTeamcenter", SwfWindow("AssociateControlFile")) Then
							If dicImportTeamcenter("ControlFilePath") <> "" Then
								'SwfWindow("AssociateControlFile").SwfEdit("FilePath").Type dicImportTeamcenter("ControlFilePath") 
								Call Fn_UI_SwfButtonClick("Fn_MSO_ImportToTeamcenter", SwfWindow("AssociateControlFile"), "Browse")
								If Fn_UI_ObjectExist("Fn_MSO_ImportToTeamcenter", SwfWindow("AssociateControlFile").Window("SelectControlFile")) Then
									If dicImportTeamcenter("ControlFilePath") <> "" Then
										SwfWindow("AssociateControlFile").Window("SelectControlFile").WinObject("FileName").Type dicImportTeamcenter("ControlFilePath") 
										Set WshShell = CreateObject("WScript.Shell")
										WshShell.SendKeys "{ENTER}"
										Set WshShell = nothing
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_ImportToTeamcenter: Successfully opened Control file [ " & dicImportTeamcenter("ControlFilePath") & " ].")
									End If
								End If ' end of open control file file
								Call Fn_UI_SwfButtonClick("Fn_MSO_ImportToTeamcenter", SwfWindow("AssociateControlFile"), "Continue")
							End If
						Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_ImportToTeamcenter: Failed to perform Import To Teamcer operation.")
						End IF
						' selecting control file sheet.
						'If SwfWindow("Select Sheet").Exist(30)  then
						If Fn_UI_ObjectExist("Fn_MSO_ImportToTeamcenter", SwfWindow("Select Sheet")) Then
							If dicImportTeamcenter("ControlFileSheet") <> "" Then
								SwfWindow("Select Sheet").SwfComboBox("SheetNames").Select dicImportTeamcenter("ControlFileSheet")
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_ImportToTeamcenter: Successfully selected Control sheet [ " & dicImportTeamcenter("ControlFileSheet") & " ].")
							Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_ImportToTeamcenter: Successfully opened Default Control sheet.")
							End If
							Call Fn_UI_SwfButtonClick("Fn_MSO_ImportToTeamcenter", SwfWindow("Select Sheet"), "OK")
						Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_ImportToTeamcenter: Failed to open Control sheet.")
						End IF

						' Control File parsing is complete!
						'If Window("MicrosoftExcel").Dialog("ExeclImportErrorDialog").Exist(20) then
						If Fn_UI_ObjectExist("Fn_MSO_ImportToTeamcenter", Window("MicrosoftExcel").Dialog("ExeclImportErrorDialog")) Then
								'Window("MicrosoftExcel").Dialog("ExeclImportErrorDialog").WinButton("OK").Click 1,1,micLeftBtn
								Call Fn_UI_WinButton_Click("Fn_MSO_ImportToTeamcenter", Window("MicrosoftExcel").Dialog("ExeclImportErrorDialog"), "OK","","","")
						ElseIf Fn_UI_ObjectExist("Fn_MSO_ImportToTeamcenter", Dialog("ConfirmationBox")) Then
								Call Fn_UI_WinButton_Click("Fn_MSO_ImportToTeamcenter", Dialog("ConfirmationBox"), "OK","","","")
						End If

						' cliclicking OK of Column Mapping window.
						If Fn_UI_ObjectExist("Fn_MSO_ImportToTeamcenter",SwfWindow("Column Mappings")) Then
								Call Fn_UI_SwfButtonClick("Fn_MSO_ImportToTeamcenter", SwfWindow("Column Mappings"), "OK")
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_ImportToTeamcenter: Successfully clicked on [ OK ] of Column Mapping window.")
						Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_ImportToTeamcenter: Failed to open Column Mapping window.")
						End If

						'Excel Import Wizard
						If Fn_UI_ObjectExist("Fn_MSO_ImportToTeamcenter",SwfWindow("Excel Import Wizard")) Then
								' Step 1
								'Verifying Current value of Create new revision checkbox
								If dicImportTeamcenter("CreateNewRevision") <> "" Then
									bFlag=False
									currVal=SwfWindow("Excel Import Wizard").SwfCheckBox("Create new revision").Object.Checked()
									If cBool(currVal) = cBool(dicImportTeamcenter("CreateNewRevision"))  Then
										bFlag=True
									End If
									If bFlag=False Then
										Call Fn_UI_SwfButtonClick("Fn_MSO_ImportToTeamcenter", SwfWindow("Excel Import Wizard"),"Cancel")
										Exit Function
									End If
								End If

								'Verifying Current value of validation error allowed
								If dicImportTeamcenter("ValidationErrorAllowed") <> "" Then
									bFlag=False
									currVal=SwfWindow("Excel Import Wizard").SwfSpin("NoOfErrorsAllowed").GetROProperty("value")
									If trim(cstr(currVal))= trim(cstr(dicImportTeamcenter("ValidationErrorAllowed"))) Then
										bFlag=True
									End If
									If bFlag=False Then
										Call Fn_UI_SwfButtonClick("Fn_MSO_ImportToTeamcenter", SwfWindow("Excel Import Wizard"),"Cancel")
										Exit Function
									End If
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_ImportToTeamcenter: Successfully set Number of Validation Errors allowed as [ " & dicImportTeamcenter("ValidationErrorAllowed") & " ].")
								End If

								'Verifying Current value of  Logging Method
								If dicImportTeamcenter("LoggingMethod") <> "" Then
									bFlag=False
									currVal=SwfWindow("Excel Import Wizard").SwfComboBox("LoggingMethod").GetItem(0)
									If currVal=dicImportTeamcenter("LoggingMethod") Then
										bFlag=True
									End If
									If bFlag=False Then
										Call Fn_UI_SwfButtonClick("Fn_MSO_ImportToTeamcenter", SwfWindow("Excel Import Wizard"),"Cancel")
										Exit Function
									End If
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_ImportToTeamcenter: Successfully set Number of Validation Errors allowed as [ " & dicImportTeamcenter("ValidationErrorAllowed") & " ].")
								End If
								If dicImportTeamcenter("sButton") <> "" Then
									Call Fn_UI_SwfButtonClick("Fn_MSO_ImportToTeamcenter", SwfWindow("Excel Import Wizard"),dicImportTeamcenter("sButton"))
								End If
								Fn_MSO_ImportToTeamcenter = True
						End IF
								' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - By Shreyas [29-02-2012]  This is a special case
				Case "CancelOperation"
					If dicImportTeamcenter("CancelOperation") <> "" Then

						'login into teamcenter
						If Fn_MSO_TeamcenterLogin("WpfExcelLogin", dicImportTeamcenter("UserID"), dicImportTeamcenter("Password"), dicImportTeamcenter("Group") , dicImportTeamcenter("Role") ) = True then
							'wait while the login window dissapears
									Do
									Loop Until WpfWindow("Teamcenter Login").Exist = False
						End if

						'Check if the Exporrt Successful Dialog Exists
						If Dialog("ConfirmationBox").Exist then
								Dialog("ConfirmationBox").Close
						End if

						'Click on OK of the Column Mapping Dialog
						If SwfWindow("Column Mappings").Exist then
							SwfWindow("Column Mappings").SwfButton("OK").Click 5,5,micLeftBtn
						End if

						'Click on Cancel of the Column Mapping Dialog
						If SwfWindow("Excel Import Wizard").Exist then
							SwfWindow("Excel Import Wizard").SwfButton("Cancel").Click 5,5,micLeftBtn
							bFlag=true
						End if

						If bflag=true Then
							Fn_MSO_ImportToTeamcenter = True
						Else
							Fn_MSO_ImportToTeamcenter = False
						End If
					End if
				' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
				Case Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_ImportToTeamcenter: Invalid case [ " & sAction & " ].")
						Exit function
		End Select
		select case dicImportTeamcenter("bCloseExcel")
			Case "Yes", "True", True
				 Call Fn_MSO_FileOperations("FileClose" ,"Excel", "", "")
		End Select
		If  Fn_MSO_ImportToTeamcenter  Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_ImportToTeamcenter: Successfully executed with case [ " & sAction & " ].")
		End If
		Set objWMIService = nothing
		Set colItems = nothing
		Set objExcel = nothing
		Set objValidationError = nothing
End Function
'------------------------------------- Function to assign item ids in data excel template in MS Office ---------
'Function Name		  :	  Fn_MSO_DataFileAssignIDs

'Description			 :	 Function to assign item ids in data excel template files

'Parameters			   :		  sAction : Action to perform
'								sSourceFilePath : file path to open data excel file
'								sDestinationPath :  filepath with file name where new data excel sheet is to be saved
'								iWorksheets : worksheet number / name
'								sIDColumn : ID column name
'								sParentIDColumn : Parent ID column name

'Return Value		   : 		  True / False

'Examples				:	Call Fn_MSO_DataFileAssignIDs( "SingleLevel", "D:\DATA-MULTI_LEVEL-STRUCT.xls","C:\test.xls","", "ID", "Parent"&vbLf&"ID")

'History					 :					
'		Developer Name				Date						Rev. No.				Changes Done				Reviewed By
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------
'		Koustubh Watwe			16/06/2011			           1.0		
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------
'		Koustubh Watwe			19/10/2011			           1.0						concatinated ' character to IDs
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_MSO_DataFileAssignIDs(sAction, sSourceFilePath,sDestinationPath, iWorksheets, sIDColumn, sParentIDColumn)
	GBL_FAILED_FUNCTION_NAME="Fn_MSO_DataFileAssignIDs"
	Dim sID, aFields, iCount, sIDCol, sParentIDCol, objExcel, iCellCnt, iColCnt
	Dim aSampleIDs(), aNewIDs(), iIDCnt, iCounter, bFlag
	Fn_MSO_DataFileAssignIDs = False
	' creating id series
	sID = ""
	'extracting year, month, date
	aFields = split(date,"/")
	For iCount = UBound(aFields)-1 to 0 Step -1
		sID = sID & trim(aFields(iCount))
	Next
	'extracting mins and hours
	aFields = split(time,":")
	For iCount = UBound(aFields)-1 to 0 Step -1
		sID = sID & trim(aFields(iCount))
	Next
	' created ID series
	'converting id series into double datatype
	sID = CDbl(sID)

	If sSourceFilePath <> "" Then
		'Open excel
            SystemUtil.Run sSourceFilePath
		wait 10
	End If

	' get already opened excel object
      Set objExcel = GetObject(,"Excel.Application")
	objExcel.Visible = True
	If iWorksheets = "" Then iWorksheets = 1
	objExcel.Worksheets(iWorksheets).Activate
	Select Case sAction
			Case "SingleLevel", "MultiLevel"
				'get cell position of ID column
				sIDCol = Fn_MSO_ExcelEditOperations("GetCellPosition","", iWorksheets, "", sIDColumn, "")
				 If sIDCol = False Then
						' column not found.
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_DataFileAssignIDs:Specified column [ " & sIDColumn & " ] is not present in given excel sheet.")
					Exit function
				End If
				'row number	
				iCellCnt = Fn_MSO_ExtractExcelRowNumber(sIDCol)
				'column number
				iColCnt = Fn_MSO_ExcelColHeader("GetColumnHeaderNumber", Fn_MSO_ExtractExcelColumnName(sIDCol))
				'get all ids from ID column and store it in Array distinctly
				iIDCnt = 0
				For iCount = iCellCnt + 1 to objExcel.Worksheets(iWorksheets).UsedRange.Rows.Count
					bFlag = True
					If trim(objExcel.Worksheets(iWorksheets).cells(iCount, iColCnt).value) <> "" Then
						' checking for existance of id in array
						If IsArray(aSampleIDs) >= 0 Then
							For iCounter = 0 to UBound(aSampleIDs)
								If aSampleIDs(iIDCnt) = trim(objExcel.Worksheets(iWorksheets).cells(iCount, iColCnt).value) Then
									bFlag = False
									Exit For
								End If
							Next
						End If
						' if id is not present in array then add it to id pool
						If bFlag = True Then
							ReDim Preserve aSampleIDs(iIDCnt)
							aSampleIDs(iIDCnt) = trim(objExcel.Worksheets(iWorksheets).cells(iCount, iColCnt).value)
							iIDCnt = iIDCnt + 1
						End If
					End If
				Next
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_DataFileAssignIDs:Successfully fetched sample IDs.")
				'assign real id to each pseudo id
				ReDim Preserve aNewIDs(UBound(aSampleIDs))
				For iCount = 0 to UBound(aSampleIDs)
					aNewIDs(iCount) = sID
					sID = sID + 1
				Next
			
				'go through ID column and replace pseudo id with new one
				For iCount =  iCellCnt + 1 to objExcel.Worksheets(iWorksheets).UsedRange.Rows.Count
					If trim(objExcel.Worksheets(iWorksheets).cells(iCount, iColCnt).value) <> "" Then
						For iIDCnt = 0 to UBound(aSampleIDs)
							If Trim(cstr(aSampleIDs(iIDCnt))) =  trim(cstr(objExcel.Worksheets(iWorksheets).cells(iCount, iColCnt).value)) Then
								objExcel.Worksheets(iWorksheets).cells(iCount, iColCnt).value = "'" & aNewIDs(iIDCnt)
								Exit For
							End If
						Next
					End If
				Next
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_DataFileAssignIDs:Successfully assigned new IDs.")
				If sParentIDColumn <> "" Then
					'get cell position of Parent ID columns
					sParentIDCol = Fn_MSO_ExcelEditOperations("GetCellPosition","", iWorksheets, "", sParentIDColumn, "")
					If sParentIDCol = False Then
						' column not found.
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_DataFileAssignIDs:Specified column [ " & sParentIDColumn & " ] is not present in given excel sheet.")
						Exit function
					End If
					' row number	
					iCellCnt = Fn_MSO_ExtractExcelRowNumber(sParentIDCol)
					' column number
					iColCnt = Fn_MSO_ExcelColHeader("GetColumnHeaderNumber", Fn_MSO_ExtractExcelColumnName(sParentIDCol))
								
					'go through Parent ID column and replace pseudo id with new one
					For iCount =  iCellCnt + 1 to objExcel.Worksheets(iWorksheets).UsedRange.Rows.Count
						If trim(objExcel.Worksheets(iWorksheets).cells(iCount, iColCnt).value) <> "" Then
							For iIDCnt = 0 to UBound(aSampleIDs)
								If Trim(cstr(aSampleIDs(iIDCnt))) =  trim(cstr(objExcel.Worksheets(iWorksheets).cells(iCount, iColCnt).value)) Then
									objExcel.Worksheets(iWorksheets).cells(iCount, iColCnt).value = "'" & aNewIDs(iIDCnt)
									Exit For
								End If
							Next
						End If
					Next
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_DataFileAssignIDs:Successfully assigned new parent IDs.")
				End If
				Fn_MSO_DataFileAssignIDs = True
		Case Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_DataFileAssignIDs: Invalid case [ " & sAction & " ].")
	End Select

	' saving file
	If sDestinationPath <> "" Then
'            objExcel.Worksheets(iWorksheets).SaveAs sDestinationPath
'            Call Fn_MSO_FileOperations("FileClose" ,"Excel", "", "")
			objExcel.ActiveWorkbook.SaveCopyAs sDestinationPath
			Call Fn_MSO_ExcelEditOperations("VerifySaveAsDialog","","","","","No")
	End If

	If Fn_MSO_DataFileAssignIDs = True Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_DataFileAssignIDs: Executed successfully with case [ " & sAction & " ].")
	End If
	Set objExcel = nothing
End Function


''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''''/$$$$
''''/$$$$   FUNCTION NAME   :  Fn_MSO_DocumentMapOperation(sAction,sNode,sMenu,sInfo1,sInfo2)
''''/$$$$
''''/$$$$   DESCRIPTION        :  This function will verify Document Map Tree Nodes
''''/$$$$ 
''''/$$$$   PRE-REQUISITES        :  The Document mapping tree should be present
''''/$$$$
''''/$$$$  PARAMETERS   : 		sAction : Action to be performed
''''/$$$$										sNode : Valid Node name to be selected  { Note : (:) seperated part will contain the Parent node & (,) seperated part contains nodes to select
''''/$$$$										sMenu : Valid Menu Component
''''/$$$$										sInfo1: For Future Use
''''/$$$$										sInfo2 : For Future Use
''''/$$$$	
''''/$$$$		Return Value : 				True or False
''''/$$$$
''''/$$$$    Function Calls       :   Fn_WriteLogFile(), Fn_MenuOperation(), Fn_ReadyStatusSync
''''/$$$$										
''''/$$$$
''''/$$$$		HISTORY           :   		AUTHOR                 DATE        VERSION
''''/$$$$
''''/$$$$    CREATED BY     :   SHREYAS          16/12/2011         1.0
''''/$$$$
''''/$$$$    REVIWED BY     :  Shreyas			16/12/2011            1.0
''''/$$$$
''''/$$$$		How To Use :    bReturn=Fn_MSO_DocumentMapOperation("Exist","Effectivity:Description,Rationale,Analysis,Notes","","","")
''''/$$$$							
''''/$$$$	Modified by		: 		Sagar S. 		10.0 Porting 		29/11/12
''''/$$$$										Implementation of function changed.
''''/$$$$										Limitations: 
''''/$$$$										1. Sometimes not able to select first node in document map when root node having expand (+), collapse (-) signs.
''''/$$$$										2. If multiple /different parent nodes consist same child nodes, then it will verify child nodes of first parent node only.
''''/$$$$
''''/$$$$	Modified By:      Nilesh Gadekar		25-Dec-2012
''''/$$$$								Implemented API method for getting Tree Node of Document Map tree
''''/$$$$
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

'
'Public function  Fn_MSO_DocumentMapOperation(sAction,sNode,sMenu,sInfo1,sInfo2)
'		Dim objWord,sValue,xParentCord,yParentCord,achild,yChildCord,xChildCord,sLeft, Top, sRight, Bottom
'		Set objWord=Window("MicrosoftWord").WinObject("MicrosoftWordDocument")
'		Select Case sAction
'					Case "Select","Exist"
'								aArray=split(sNode,":",-1,1)
'								sValue= objWord.GetTextLocation(trim(aArray(0)),sLeft, Top, sRight, Bottom,false)
'								If sValue<>true Then 
'										Fn_MSO_DocumentMapOperation = False			 				
'										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail: Fn_MSO_DocumentMapOperation failed due to Invalid arguments ["+aArray(0)+"]")
'										Exit function
'								else 
'								   xParentCord = (sLeft+sRight) / 2 
'								   yParentCord =(Top+Bottom) / 2 
'									wait 1
'									objWord.Click xParentCord,yParentCord, micLeftBtn
'									If err.number<0 Then
'										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Select the WordTree Node  ["+aArray(0)+"]")
'										Fn_MSO_DocumentMapOperation=False
'										Exit Function
'									Else
'										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Selected the WordTree Node  ["+aArray(0)+"]")
'										Fn_MSO_DocumentMapOperation=True	
'									End If
'							End If 	
'							achild=split(aArray(1),",",-1,1)
''							sValue2=objWord.GetTextLocation(aArray(1),sLeft, Top, sRight, Bottom)
'							yChildCord = yParentCord
'								For iCount=0 to ubound(achild)
'								sValue= objWord.GetTextLocation(trim(achild(iCount)),0,0,0,0,False)
'								If sValue<>true Then 
'										Fn_MSO_DocumentMapOperation = False			 				
'										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail: Fn_MSO_DocumentMapOperation failed due to Invalid arguments ["+achild(iCount)+"]")
'										Exit function
'								else 
''										sValue2=objWord.GetTextLocation(trim(achild(iCount)),sLeft, Top, sRight, Bottom)
'										yChildCord = yChildCord +12
'										xChildCord = xParentCord
'										wait 1
'										objWord.Click xChildCord,yChildCord, micLeftBtn
'										If err.number<0 Then
'											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Select and verify the existence of the WordTree Node  ["+achild(iCount)+"]")
'											Fn_MSO_DocumentMapOperation=False
'											Exit Function
'										Else
'											Fn_MSO_DocumentMapOperation=True	
'											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Selected and verified the existence of the WordTree Node  ["+achild(iCount)+"]")
'										End If
'								End if
'							next
'				End Select
'				Set objWord=nothing
'End function

Public Function  Fn_MSO_DocumentMapOperation(sAction,sNode,sMenu,sInfo1,sInfo2)
	GBL_FAILED_FUNCTION_NAME="Fn_MSO_DocumentMapOperation"
	Dim aArray,iCount,sValue,xParentCord,yParentCord
	Dim sLeft,Top,sRight,Bottom
	Dim objWord,ShellObj', objWordWindow
	Dim objDocumentMap,HeaderList,aNode,iFlag,jCount
	Dim sExpected,sActual,aActual
	sLeft=-1
	Top=-1
	sRight=-1
	Bottom=-1
	If Instr(1,sNode,",")<>0 Then				' If string contain comma replace it with colon":"
		sNode=Replace(sNode,",",":")
	End If
	Wait(2)
	Fn_MSO_DocumentMapOperation=False
	'Set objWordWindow = Window("MicrosoftWord").WinObject("MicrosoftWordDocument")
	Set objWord=GetObject(,"Word.application")
	Set ShellObj=CreateObject("WScript.Shell")
	objWord.ActiveWindow.DocumentMap = True
	Wait(2)
	'objWordWindow.SetTOProperty "Location",0
	Select Case sAction					
					Case "Select","Exist"
								aArray=split(sNode,":",-1,1)
'								For iCount=0 to UBound(aArray)
'										objWordWindow.Highlight
'										If iCount=0 Then
'												ShellObj.SendKeys "{PGDN}"
'												Wait(2)
'												sValue= Window("MicrosoftWord").GetTextLocation(trim(aArray(iCount)),sLeft, Top, sRight, Bottom)
'												If sValue<>True And Err.number<0Then 
'														Exit function
'												Else 
'													   xParentCord = (sLeft+sRight) / 2 
'													   yParentCord =(Top+Bottom) / 2 
'														wait 1
'														Window("MicrosoftWord").Click xParentCord,yParentCord, micLeftBtn
'												End If
'										Else
'												ShellObj.SendKeys "{PGUP}"									
'												Wait(2)
'												objWordWindow.ClickOnText aArray(iCount)
'										End If
'										If Err.number<0 Then
'											Fn_MSO_DocumentMapOperation=False
'										Else
'											Fn_MSO_DocumentMapOperation=True
'										End If 
'								Next			
										'Added by Nilesh on 25-Dec-2012   APImethod implementation for Document Map Tree
										Set objDocumentMap=DotnetFactory.CreateInstance("WordDocumentMapTree.DocumentMapTree",Environment.Value("sPath")+"\Library\WordDocumentMapTree.dll")
										HeaderList=objDocumentMap.GetDocumentMapTreeNode()  'Returns the Tree Node seperated by ~ from Document map tree
										If HeaderList="" Then
											Fn_MSO_DocumentMapOperation=False
											Set objDocumentMap=Nothing
											Exit Function
										End If
										iFlag=0
										For jCount=0 to Ubound(aArray)
											aNode=Split(HeaderList,"~",-1,1)
											For iCount=1 to Ubound(aNode)
											sActual	=Replace(aNode(iCount)," ","")  'Remove Spaces from String
											'Remove the number
											If Instr(sActual,".")>0 Then
												aActual=Split(sActual,".",-1,1)   
												sActual=aActual(1)
											End If
												sExpected=Replace(aArray(jCount)," ","")			'Remove Spaces from String
												If Instr(LCase(sActual),LCase(sExpected))>0 Then
													Fn_MSO_DocumentMapOperation=True
													iFlag=iFlag+1
													Exit For
												End If
											Next
										Next

										If iFlag= Ubound(aArray)+1 Then
											Fn_MSO_DocumentMapOperation=True
										Else
											Fn_MSO_DocumentMapOperation=False
										End If
	End Select
			objWord.ActiveWindow.DocumentMap = False			' close Document Map
			Set objDocumentMap=Nothing
			Set objWord=nothing
			'Set objWordWindow=nothing
End Function

''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''''/$$$$
''''/$$$$   FUNCTION NAME   :   Fn_MSO_CreateAndFollow_HyperLink(sUrl,bFollowLink,sInfo1,sInfo2)
''''/$$$$
''''/$$$$   DESCRIPTION        :  This function will create Hyperlinks within a word Document
''''/$$$$ 
''''/$$$$   PRE-REQUISITES        :  Microsoft Word Wimdow should be Opened with the Text to be made as Hyperlink
''''/$$$$
''''/$$$$  PARAMETERS   : 		sUrl : Valid URL to be made as Hyperlink
''''/$$$$											bFollowLink : Boolean Parameter to follow the Hyperlink
''''/$$$$										sInfo1: For Future Use
''''/$$$$										sInfo2 : For Future Use
''''/$$$$	
''''/$$$$		Return Value : 				True or False
''''/$$$$
''''/$$$$    Function Calls       :   Fn_WriteLogFile(), Fn_MenuOperation(), Fn_ReadyStatusSync
''''/$$$$										
''''/$$$$
''''/$$$$		HISTORY           :   		AUTHOR                 DATE        VERSION
''''/$$$$
''''/$$$$    CREATED BY     :   SHREYAS          20/12/2011         1.0
''''/$$$$
''''/$$$$    REVIWED BY     :  Shreyas			20/12/2011            1.0
''''/$$$$
''''/$$$$		How To Use :   bReturn=Fn_MSO_CreateAndFollow_HyperLink("http://www.google.com","True","","")
''''/$$$$							
''''/$$$$	
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

Public Function Fn_MSO_CreateAndFollow_HyperLink(sUrl,bFollowLink,sInfo1,sInfo2)
	
	GBL_FAILED_FUNCTION_NAME="Fn_MSO_CreateAndFollow_HyperLink"
	   Dim objWord,sValue,xParentCord,yParentCord,objHyperLink,sLeft, Top, sRight, Bottom
	Set objWord=GetObject(,"Word.application")
	objWord.ActiveWindow.DocumentMap = False

	Set objWord=Window("MicrosoftWord").WinObject("MicrosoftWordDocument")
	'Set objHyperLink=Window("MicrosoftWord").Window("Insert Hyperlink")
	Fn_MSO_CreateAndFollow_HyperLink=false

									'Select the Text to be Converted into Hyperlink

									sValue= objWord.GetTextLocation("://",sLeft, Top, sRight, Bottom)
									If sValue<>true Then 
											Fn_MSO_Create_HyperLink = False			 				
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail: Fn_MSO_Create_HyperLink failed due to Invalid arguments ["+aArray(0)+"]")
											Exit function
									else 
									   xParentCord = (sLeft+sRight) / 2 
									   yParentCord =(Top+Bottom) / 2 
										wait 3
										objWord.Type micCtrlDwn
										objWord.Click xParentCord,yParentCord, micLeftBtn

										If err.number<0 Then
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Select the Text to make hyperlink")
											Fn_MSO_CreateAndFollow_HyperLink=False
											Exit Function
										Else
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Selected the Text to make hyperlink")
											Fn_MSO_CreateAndFollow_HyperLink=True	
										End If
								End If 
						wait 2


										'Invoke the  Create Hyperlink Dialog
										objWord.Type micCtrlDwn + "k" + micCtrlUp
										If err.number<0 Then
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Select the Text to make hyperlink")
											Fn_MSO_CreateAndFollow_HyperLink=False
											Exit Function
										Else
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Selected the Text to make hyperlink")
											Fn_MSO_CreateAndFollow_HyperLink=True	
										End If
										wait 3

										If Window("MicrosoftWord").Window("Insert Hyperlink").Exist(5) Then
                   							Set objHyperLink=Window("MicrosoftWord").Window("Insert Hyperlink")
                						ElseIf Window("MicrosoftWord").WinObject("Insert Hyperlink").Exist(5) Then
                    						Set objHyperLink=Window("MicrosoftWord").WinObject("Insert Hyperlink")
               						 	End If
               						 	Wait 1
										'Clear the Contents from the Url TextField
										objHyperLink.WinObject("URL").Type micShiftDwn +  micEnd  + micShiftUp
										wait 1
										objHyperLink.WinObject("URL").Type  micDel
										wait 1

										'Set the Url for Hyperlink
									If sUrl<>"" Then
											objHyperLink.WinObject("URL").Type sUrl
											If err.number<0 Then
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Set the URL for Hyperlink as ["+cstr(sUrl)+"]")
												Fn_MSO_CreateAndFollow_HyperLink=False
												Exit Function
											Else
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Set the URL for Hyperlink as ["+cstr(sUrl)+"]")
												Fn_MSO_CreateAndFollow_HyperLink=True	
											End If										
											wait 1
									End If

										'Hit OK on the Insert Hyperlink Dialog
										objHyperLink.WinObject("URL").Type  micReturn



										If cBool(bFollowLink)=true Then
											'Follow the Link  by Pressing Control Key and Clicking
											objWord.Type micCtrlDwn
											objWord.Click xParentCord,yParentCord, micLeftBtn
											objWord.Type micCtrlUp
												If err.number<0 Then
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to follow the Hyperlink ["+cstr(sUrl)+"]")
													Fn_MSO_CreateAndFollow_HyperLink=False
													Exit Function
												Else
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully followed the Hyperlink ["+cstr(sUrl)+"]")
													Fn_MSO_CreateAndFollow_HyperLink=True	
												End If
											wait 30
									End If

									'Check for any Errors / Information Popups
									If  Window("MicrosoftWord").Dialog("Information").Exist(40) Then
										Window("MicrosoftWord").Dialog("Information").WinButton("OK").Click 0,0,micLeftBtn
											If err.number<0 Then
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Handle the [Information Dialog]")
												Fn_MSO_CreateAndFollow_HyperLink=False
												Exit Function
											Else
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Handled the [Information Dialog]")
												Fn_MSO_CreateAndFollow_HyperLink=True	
											End If
									End If

						Set objWord=nothing
						Set objHyperLink=nothing

End Function



''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''''/$$$$
''''/$$$$   FUNCTION NAME   :  Fn_MSO_Find_And_ReplaceText(aLoginDetails,sTextToFind,sTextToReplace,bSave,sInfo1,sInfo2)
''''/$$$$
''''/$$$$   DESCRIPTION        :  This function will Find ,Replace & Save text in Ms-Word Document
''''/$$$$ 
''''/$$$$   PRE-REQUISITES        :  MS-Word Window Should be in Focus
''''/$$$$
''''/$$$$  PARAMETERS   : 		aLoginDetails :Login Details
''''/$$$$										sTextToFind : Valid Text to find
''''/$$$$										sTextToReplace : Valid Text to Replace
''''/$$$$										bSave : Boolean Parameter to perform Save Operation
''''/$$$$										sInfo1: For Future Use
''''/$$$$										sInfo2 : For Future Use
''''/$$$$	
''''/$$$$		Return Value : 				True or False
''''/$$$$
''''/$$$$    Function Calls       :   Fn_WriteLogFile(), Fn_MenuOperation(), Fn_ReadyStatusSync
''''/$$$$										
''''/$$$$
''''/$$$$		HISTORY           :   		AUTHOR                 DATE        VERSION
''''/$$$$
''''/$$$$    CREATED BY     :   SHREYAS          12/01/2012         1.0
''''/$$$$
''''/$$$$    REVIWED BY     :  Shreyas			12/01/2012           1.0
''''/$$$$
''''/$$$$		How To Use :    bReturn=Fn_MSO_Find_And_ReplaceText(Environment.Value("TcUser4"),"two","Shreyas","True","","")
 ''''/$$$$						bReturn=Fn_MSO_Find_And_ReplaceText(Environment.Value("TcUser4"),"two","Shreyas","True","RenameNode","")
''''/$$$$							
''''/$$$$	
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$


Public Function Fn_MSO_Find_And_ReplaceText(aLoginDetails,sTextToFind,sTextToReplace,bSave,sInfo1,sInfo2)
	GBL_FAILED_FUNCTION_NAME="Fn_MSO_Find_And_ReplaceText"

	Dim objWord,sValue,xParentCord,yParentCord,objHyperLink,sLeft, Top, sRight, Bottom,devicereplay,wshshell,bReturn,aLogin,WordSel
	Dim bFlag,aTextFind
	
	Set objWord=Window("MicrosoftWord").WinObject("MicrosoftWordDocument")
	Set devicereplay=CreateObject("Mercury.DeviceReplay")
	Set wshshell=Createobject("WScript.Shell")
	
	'Added By Nilesh on 17-Dec-2012
	If Instr(sTextToFind,".")>0 Then
		aTextFind=Split(sTextToFind,". ",-1,1)
		If Ubound(aTextFind)<>0 Then
			sTextToFind=aTextFind(1)
		Else
			aTextFind=Split(sTextToFind,".",-1,1)
			sTextToFind=aTextFind(1)
		End If
		
	End If
	Const ctrl=29
	'@Added by Nilesh on 15-Jul-2013  to remove focus from Word document
	If JavaWindow("DefaultWindow").Exist(5)=True And  JavaWindow("DefaultWindow").GetROProperty("enabled")=1 Then
        JavaWindow("DefaultWindow").Minimize
	End If
	'@End
	'Close the Login Window if it Exists
	
			wait 5
				'Login into Ms-Word addin Teamcenter
				If WpfWindow("Teamcenter Login").Exist(5)  Then
					aLogin=split(aLoginDetails,":",-1,1)
					bReturn=Fn_MSO_TeamcenterLogin("WpfExcelLogin", aLogin(0), aLogin(1),aLogin(2), aLogin(3) )
					If bReturn=true Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully logged in with User ["+aLogin(0)+"]")
								wait (3)
								Fn_MSO_Find_And_ReplaceText=True	
                                If  sTextToFind = "" And  sTextToReplace= ""  Then
									Fn_MSO_Find_And_ReplaceText=True
									wait 10
									If Dialog("ConfirmationBox").Exist(20) Then
                                        Dialog("ConfirmationBox").WinButton("OK").Click
										wait 2
									End If
									Exit Function
								End If
					Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to log in with User ["+aLogin(0)+"]")
								Fn_MSO_Find_And_ReplaceText=False

								
					End If
				End If
				wait 20		'it takes some time to load the document after login hence implemented wait after that
				If Dialog("ConfirmationBox").Exist(20) Then
					Dialog("ConfirmationBox").WinButton("OK").Click
					wait 2
				End If
				Set objWordWindow = Nothing
				JavaWindow("DefaultWindow").Maximize 'To remove the Focus From Word File
				Set objWordWindow = GetObject(,"Word.Application")
				objWordWindow.Visible = True
			'' Below code is added to close document Map
				objWordWindow.ActiveWindow.DocumentMap = false
							

			If sInfo1="" Then
					sInfo1="RenameNode"
			End If
			If sInfo1<>"" Then
				Select Case sInfo1
				
								Case "RenameNode"
	
	'Commented by Nilesh on 17-Dec-2012
	'															objWord.GetTextLocation sTextToFind,sLeft,Top,sRight,sBottom,True
	'															xParentCord = (sLeft+sRight) / 2
	'															yParentCord=(Top+sBottom) / 2
	'															objWord.Click xParentCord,yParentCord, micLeftBtn
	'									
	'															wait 2
	'									
	'													' press the  down arrow key twice
	'													For i= 1 to 2
	'															bReturn= Fn_KeyBoardOperation("SendKeys", "{DOWN}")
	'															If bReturn=true Then
	'																Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Successfully pressed the Down Arrow Key "+cstr(i)+" times")
	'																Fn_MSO_Find_And_ReplaceText=true
	'															Else
	'															   Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to click the Down arrow key")
	'															   Fn_MSO_Find_And_ReplaceText=False
	'															   Exit function	
	'															End If
	'													Next
	'													wait 2
	'
	'
	'											'Select the TextTo MArkup
	'
	'											bReturn= Fn_KeyBoardOperation("SendKeys", "{END}")
	'											wait 2
	'											bReturn= Fn_KeyBoardOperation("SendKeys", "+{HOME}")
	'											If bReturn=true Then
	'												Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Successfully selected the Text to Markup")
	'												Fn_MSO_Find_And_ReplaceText=true
	'											Else
	'											   Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to select the Text to Markup")
	'											   Fn_MSO_Find_And_ReplaceText=False
	'											   Exit function	
	'											End If
	
												'Type the Text to Replace
												'Added By Nilesh on 17-Dec-2012
												If Window("MicrosoftWord").exist(3) = True Then
													Window("MicrosoftWord").Activate
												End If
													Set WordSel=objWordWindow.Selection
													WordSel.Find.Text=sTextToFind
													WordSel.Find.Forward = TRUE
													WordSel.Find.MatchWholeWord = TRUE
													bFlag=WordSel.Find.Execute
													Wait 2
													If bFlag=True Then
														objWord.Type sTextToReplace									
														If err.number<0 Then
																Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Set the Text ["+sTextToReplace+"]")
																Fn_MSO_Find_And_ReplaceText=False
																Exit Function
														Else
																Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Set the Text ["+sTextToReplace+"]")
																Fn_MSO_Find_And_ReplaceText=True
														End If
		
																			'save the Details
															If bSave<>"" and cbool(bSave)=True Then
																	If Window("MicrosoftWord").exist(3) = True Then
																		Window("MicrosoftWord").Activate
																	End If
																	Set wshshell=Createobject("WScript.Shell")
																	wshshell.SendKeys "^(s)"
																	Wait 1
																	'WpfWindow("Save").WpfButton("OK").Click 5,5,micLeftBtn
																	'wait 2
																	'WpfWindow("Save").WpfButton("OK").highlight()
																	'wait 1
'                                                                    WpfWindow("Save").WpfButton("text:=OK").Click 
																	'WshShell.SendKeys "{ENTER}"
																	
																	
																If Window("MicrosoftWord").Dialog("Information").Exist(1) Then
																	bReturn = Window("MicrosoftWord").Dialog("Information").GetTextLocation("Yes",l, t, r, b, True)
																	If bReturn = true Then
																		x = (l+r)/2
																		y = (t+b)/2
																		Window("MicrosoftWord").Dialog("Information").Click x,y, micLeftBtn
																	End If
																End If
																
																' TC112-2015070100-17_07_2015-Porting-VivekA-Added code to handle Save dialog as discussed with Dhananjay
																If WpfWindow("Save").Exist(1) Then
	                                                    			WpfWindow("Save").WpfButton("OK").Click 5,5,micLeftBtn
	                                                    			Wait 2
																End If
																	
															End If
															Set objWord=nothing
															Set WordSel=Nothing
															Set devicereplay=nothing
															Set wshshell=nothing
					
															Fn_MSO_Find_And_ReplaceText=True
															Exit Function
													Else
															Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Find  ["+sTextToFind+"] in Word application")
															Fn_MSO_Find_And_ReplaceText=False
															Exit Function
													End If
				
											
				End Select
			End If
	
	'				sValue= objWord.GetTextLocation (sTextToFind,sLeft,Top,sRight,sBottom,True)
	'				If sValue=true Then
	'						xParentCord = (sLeft+sRight) / 2
	'						yParentCord=(Top+sBottom) / 2
	'				Else
	'						Fn_MSO_Find_And_ReplaceText = False
	'						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail: Fn_MSO_Create_HyperLink failed due to Invalid arguments ["+sTextToFind+"]")
	'						Exit function
	'				End If
	'
	''select the text to replace
	'			objWord.Click xParentCord,yParentCord,micLeftBtn
	'			If err.number<0 Then
	'					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Select the Text to to replace")
	'					Fn_MSO_Find_And_ReplaceText=False
	'					Exit Function
	'			Else
	'					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Selected the Text ["+sTextToFind+"] to replace")
	'					Fn_MSO_Find_And_ReplaceText=True
	'			End If
	'			wait 1
	'
	''now type the text to replace
	'
	'			devicereplay.KeyDown ctrl
	'			objWord.Click xParentCord,yParentCord,micLeftBtn
	'			devicereplay.KeyUp ctrl
	'			wait 1
	'			objWord.Type sTextToReplace
	'			wait 1
	'
	'			'save the Details
	'			If bSave<>"" and cbool(bSave)=True Then
	'			Set wshshell=Createobject("WScript.Shell")
	'			wshshell.SendKeys "^(s)"
	'			wait 1
	'			'WpfWindow("Save").WpfButton("OK").Click 5,5,micLeftBtn
	'			Call Fn_UI_WpfObjectClick("",WpfWindow("Save").WpfButton("OK"))
	'			End If
	'Added By Nilesh on 17-Dec-2012
													Set WordSel=objWordWindow.Selection
													WordSel.Find.Text=sTextToFind
													WordSel.Find.Forward = TRUE
													WordSel.Find.MatchWholeWord = TRUE
													bFlag=WordSel.Find.Execute
													Wait 2
													If bFlag=True Then
														objWord.Type sTextToReplace									
														If err.number<0 Then
																Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Set the Text ["+sTextToReplace+"]")
																Fn_MSO_Find_And_ReplaceText=False
																Exit Function
														Else
																Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Set the Text ["+sTextToReplace+"]")
																Fn_MSO_Find_And_ReplaceText=True
														End If
		
																			'save the Details
															If bSave<>"" and cbool(bSave)=True Then
																	Set wshshell=Createobject("WScript.Shell")
																	wshshell.SendKeys "^(s)"
																	wait 1
																	WpfWindow("Save").WpfButton("OK").Click 5,5,micLeftBtn
															End If
															Set objWord=nothing
															Set WordSel=Nothing
															Set devicereplay=nothing
															Set wshshell=nothing
					
															Fn_MSO_Find_And_ReplaceText=True
															Exit Function
													Else
															Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Find  ["+sTextToFind+"] in Word application")
															Fn_MSO_Find_And_ReplaceText=False
															Exit Function
													End If
	
	Set objWord=nothing
	Set devicereplay=nothing
	Set wshshell=nothing
End Function

''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''''/$$$$
''''/$$$$  FUNCTION NAME   :  Fn_MSO_ItemBasicCreate(aUser,sItemType,sItemName,sItemDesc,sUom,bOpenOnCreate,ByRef SItemRev,sInfo1,sInfo2)
''''/$$$$
''''/$$$$  DESCRIPTION        :  This function will create an Item In MS-Word/ Ms-Excel (Provided there is Teamcenter Addin For Ms-Office Installed on the Client)
''''/$$$$ 
''''/$$$$  PRE-REQUISITES   :  MS-Word Window Should be in Focus
''''/$$$$
''''/$$$$  PARAMETERS   	:  aUser : Login Details
''''/$$$$ 					   sItemName : valid Item Name
''''/$$$$ 					   sItemType : Item Type
''''/$$$$ 					   sItemDesc : Valid Item Description
''''/$$$$ 					   sUom : Valid Unit Of Measure
''''/$$$$ 					   bOpenOnCreate : Parameter to Open the Item On Create
''''/$$$$ 					   ByRef SItemRev : This function has a special last parameter SItemRev which is passed by Reference (ByRef) which will 
''''/$$$$									give you the provision to Retrieve the Details of the Item Revision
''''/$$$$ 					   sInfo1 : For Future Use
''''/$$$$ 					   sInfo2 : For Future Use
''''/$$$$	
''''/$$$$	Return Value 		:  True or False
''''/$$$$
''''/$$$$  Function Calls       	:  Fn_WriteLogFile(), Fn_MenuOperation(), Fn_ReadyStatusSync
''''/$$$$										
''''/$$$$	HISTORY           	:  	AUTHOR             	DATE            VERSION
''''/$$$$
''''/$$$$  CREATED BY     	:  SHREYAS          	16/02/2012        	1.0
''''/$$$$
''''/$$$$  REVIWED BY     	:  Shreyas		16/02/2012          1.0
''''/$$$$
''''/$$$$	How To Use 		:  bReturn=Fn_MSO_ItemBasicCreate(Environment.Value("TcUserDBA"),"Item","TestItem","Item WIth Office Addin","","off",sRevision,"","")
''''/$$$$					   bReturn=Fn_MSO_CreateFolder("","TestFolderName","TestFolder","Finish","Folder","")
''''/$$$$					  
''''/$$$$	Modified By 		: Vivek Ahirrao	10/06/2016		1.1				[TC1123-20160504-10_06_2016-VivekA-NewDevelopment]
''''/$$$$						Modified as per Design change, added code to set Object Type and then click on Next button
''''/$$$$						Modified for RM new development testcases					
''''/$$$$	
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Public Function Fn_MSO_ItemBasicCreate(aUser,sItemType,sItemName,sItemDesc,sUom,bOpenOnCreate,ByRef SItemRev,sInfo1,sInfo2)
	GBL_FAILED_FUNCTION_NAME="Fn_MSO_ItemBasicCreate"
   	Dim sValue,xParentCord1,yParentCord1,wshshell,objWord,objNewItem,xParentCord2,yParentCord2
   	Dim sItemId,sGeneratedItemName,sRev,iCount,aLogin,sCreateCounter,objExcel
   	Dim sLeft2,Top2,sRight2,sBottom2, objPowerPoint, sAppType

	'Check if the Object Exists for Word Or Excel or PowerPoint
	Set objNewItem = WpfWindow("New Item")
	Set objWord = Window("MicrosoftWord")
	Set objExcel = Window("MicrosoftExcel")
	Set objPowerPoint = Window("MicrosoftPowerPoint")

	Fn_MSO_ItemBasicCreate = False
	
	If Not objNewItem.Exist(5) Then
		If objWord.Exist(5) Then
			sAppType = "MSWord"
		ElseIf objExcel.Exist(5) Then
			sAppType = "MSExcel"
		ElseIf objPowerPoint.Exist(5) Then
			sAppType = "MSPowerPoint"
		Else
			Set objNewItem = Nothing
			Set objWord = Nothing
			Set objExcel = Nothing
			Set objPowerPoint = Nothing
			Exit Function
		End If
		
		Set wshshell = CreateObject("WScript.Shell")
		wshshell.SendKeys "%"
		wshshell.SendKeys "Y"
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Successfully Activated the Teamcenter Tab")
		Fn_MSO_ItemBasicCreate=true
		Set wshshell = Nothing
		wait 2
		
		bREturn = Fn_MSO_RibbonButton_Operations(sAppType,"Click","New:Item","")
		If bREturn=True Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Successfully performed Menu operation [New:Item]")
			Fn_MSO_ItemBasicCreate = True
		Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to perform Menu operation [New:Item]")
			Fn_MSO_ItemBasicCreate = False
			Set objNewItem = Nothing
			Set objWord = Nothing
			Set objExcel = Nothing
			Set objPowerPoint = Nothing
			Exit function
		End If
		wait 5
		
		'Login into Teamcenter
		If WpfWindow("Teamcenter Login").Exist(5) Then
			aLogin=split(aUser,":",-1,1)
			If sAppType = "MSWord" Then
				bReturn = Fn_MSO_TeamcenterLogin("WpfWordLogin", aLogin(0), aLogin(0), aLogin(2), aLogin(3) )
			ElseIf sAppType = "MSExcel" Then
				bReturn = Fn_MSO_TeamcenterLogin("WpfExcelLogin", aLogin(0), aLogin(0), aLogin(2), aLogin(3) )
			ElseIf sAppType = "MSPowerPoint" Then
				bReturn = Fn_MSO_TeamcenterLogin("WpfPowerPointLogin", aLogin(0), aLogin(0), aLogin(2), aLogin(3) )
			End If
			If bReturn=True Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Successfully logged in with user ["+aLogin(0)+"]")
				Fn_MSO_ItemBasicCreate=true
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to log in with user ["+aLogin(0)+"]")
				Fn_MSO_ItemBasicCreate=False
				Set objNewItem = Nothing
				Set objWord = Nothing
				Set objExcel = Nothing
				Set objPowerPoint = Nothing
				Exit function
			End If
		End If
	
		If Not objNewItem.Exist(8) Then
			wait 10
		End If
		wait 5
	End If
	
	'Select the Desired Item Type
	If sItemType<>"" Then
		objNewItem.WpfComboBox("ItemType").Select sItemType
		If Err.Number<0 Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to Select the Item Type ["+sItemType+"]")
			Fn_MSO_ItemBasicCreate=False
			Set objNewItem = Nothing
			Set objWord = Nothing
			Set objExcel = Nothing
			Set objPowerPoint = Nothing
			Exit function
		End If
		wait 2
	End If
	
'***********	'Click on Next Button	**************
	If objNewItem.WpfButton("Next").GetROProperty("enabled") = True Then
		objNewItem.WpfButton("Next").Click 5,5,micLeftBtn
		If Err.Number<0 Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to click on the next button")
			Fn_MSO_ItemBasicCreate=False
			Set objNewItem = Nothing
			Set objWord = Nothing
			Set objExcel = Nothing
			Set objPowerPoint = Nothing
			Exit function
		End If
		wait 3
	End If
'**********************************************

	'Click on Assign Button
	objNewItem.WpfButton("Assign").Click 5,5,micLeftBtn
	If Err.Number<0 Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to click on the assign button")
		Fn_MSO_ItemBasicCreate=False
		Set objNewItem = Nothing
		Set objWord = Nothing
		Set objExcel = Nothing
		Set objPowerPoint = Nothing
		Exit function
	End If
	wait 3
				
	'Set the name in the Name Edit Box
	If sItemName<>"" Then
		objNewItem.WpfEdit("object_name").Set sItemName
		If Err.Number<0 Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to set the name ["+sName+"]")
			Fn_MSO_ItemBasicCreate=False
			Set objNewItem = Nothing
			Set objWord = Nothing
			Set objExcel = Nothing
			Set objPowerPoint = Nothing
			Exit function
		End If
		wait 2
	End If
				
	'Set the Description
	If sItemDesc<>"" Then
		objNewItem.WpfObject("PropertyText").SetTOProperty "devname","Description"
		objNewItem.WpfEdit("PropertyEditBox").Set sItemDesc
		'objNewItem.WpfEdit("Description").Set sItemDesc
		If Err.Number<0 Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to set the Description as ["+sItemDesc+"]")
			Fn_MSO_ItemBasicCreate=False
			Set objNewItem = Nothing
			Set objWord = Nothing
			Set objExcel = Nothing
			Set objPowerPoint = Nothing
			Exit function
		End If
		wait 3
	End If

	'Set the Unit Of Measure
	If sUom<>"" Then
		objNewItem.WpfComboBox("UnitOfMeasure").Select sUom
		If Err.Number<0 Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to set the Unit of measure ["+sUom+"]")
			Fn_MSO_ItemBasicCreate=False
			Set objNewItem = Nothing
			Set objWord = Nothing
			Set objExcel = Nothing
			Set objPowerPoint = Nothing
			Exit function
		End If
		wait 3
	End If
				
	'Check the Unit of Measure Checkbox to either on or off
	If bOpenOnCreate<>"" Then
		If lCAse(bOpenOnCreate)="off" Then
			objNewItem.WpfCheckBox("OpenOnCreate").Set uCase(bOpenOnCreate)
		Elseif lCAse(bOpenOnCreate)="on" Then
			objNewItem.WpfCheckBox("OpenOnCreate").Set uCase(bOpenOnCreate)
		Else
			objNewItem.WpfCheckBox("OpenOnCreate").Set "OFF"	
		End If
		wait 3
	End if
	
	'save the Item name Details & assign it  to the Function
	sItemId = objNewItem.WpfEdit("item_id").GetROProperty ("text")
	sName = objNewItem.WpfEdit("object_name").GetROProperty ("text")
	sRev = objNewItem.WpfEdit("item_revision_id").GetROProperty ("text")
	
	sGeneratedItemName = sItemId+"-"+sName
	Fn_MSO_ItemBasicCreate = sGeneratedItemName
	sItemRev = sItemId & "/"& sRev & ";1-" & sName
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Successfully Created Item ["+sGeneratedItemName+"]")
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Successfully Generated Item Revision Name as ["+sItemRev+"]")

'***********	'Click on Finish Button	**************
	If objNewItem.WpfButton("Finish").GetROProperty("enabled") = True Then
		objNewItem.WpfButton("Finish").Click 5,5,micLeftBtn
		If Err.Number<0 Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to click on the Finish button")
			Fn_MSO_ItemBasicCreate=False
			Set objNewItem = Nothing
			Set objWord = Nothing
			Set objExcel = Nothing
			Set objPowerPoint = Nothing
			Exit function
		End If
		wait 3
	End If
'**********************************************

	'close the New Item Window if exists
	If objNewItem.Exist Then
		objNewItem.WpfButton("Close").Click 5,5,micLeftBtn
		If Err.Number<0 Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to click on the Close button")
			Fn_MSO_ItemBasicCreate=False
			Set objNewItem = Nothing
			Set objWord = Nothing
			Set objExcel = Nothing
			Set objPowerPoint = Nothing
			Exit function
		End If
		wait 3
	End If

	Set objNewItem = Nothing
	Set objWord = Nothing
	Set objExcel = Nothing
	Set objPowerPoint = Nothing
	
End Function

''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''''/$$$$
''''/$$$$  FUNCTION NAME	:  Fn_MSO_CreateFolder(aUser,sFolderName,sFolderDesc,sButtons,sInfo1,sInfo2)
''''/$$$$
''''/$$$$  DESCRIPTION   	:  This function will create an Item In MS-Word/ Ms-Excel (Provided there is Teamcenter Addin For Ms-Office Installed on the Client)
''''/$$$$ 
''''/$$$$  PRE-REQUISITES	:  MS-Word Window Should be in Focus
''''/$$$$
''''/$$$$  PARAMETERS   	:  aUser :Login Details
''''/$$$$ 					   sFolderName: valid Folder Name
''''/$$$$					   sFolderDesc : Valid Folder Description
''''/$$$$					   sButtons : Valid Combination or a single button to be clicked
''''/$$$$					   sInfo1: Object Type
''''/$$$$					   sInfo2 : Select Location in which you want to create object
''''/$$$$	
''''/$$$$	Return Value 		: True or False
''''/$$$$
''''/$$$$  Function Calls     	: Fn_WriteLogFile(), Fn_MenuOperation(), Fn_ReadyStatusSync
''''/$$$$										
''''/$$$$  HISTORY           	: AUTHOR      			DATE        VERSION
''''/$$$$
''''/$$$$  CREATED BY     	: SHREYAS    		17/02/2012        1.0
''''/$$$$
''''/$$$$  REVIWED BY     	: Shreyas		17/02/2012        1.0
''''/$$$$
''''/$$$$	How To Use 		: bReturn=Fn_MSO_CreateFolder(Environment.Value("TcUserDBA"),"NewFolder","Folder Created WIth Office Addin","Apply:OK","","")
''''/$$$$					  bReturn=Fn_MSO_CreateFolder("","TestFolderName","TestFolder","Finish:Close","Folder","")
''''/$$$$					  Set dicDetails = CreateObject("Scripting.Dictionary")
''''/$$$$					  	 dicDetails("VerifySelectedLocation") = "AutomatedTests"
''''/$$$$					  bReturn = Fn_MSO_CreateFolder("Verify$","","","Close","",dicDetails)
''''/$$$$					  
''''/$$$$	Modified By 		: Vivek Ahirrao	10/06/2016	    1.1					[TC1123-20160504-10_06_2016-VivekA-NewDevelopment]
''''/$$$$						Modified as per Design change, added code to set Object Type and then click on Next button
''''/$$$$						Modified for RM new development testcases
''''/$$$$					  Vivek Ahirrao	05/07/2016	    1.1	Added case "Verify", Modified for more customization.
''''/$$$$	
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Public Function Fn_MSO_CreateFolder(aUser,sFolderName,sFolderDesc,sButtons,sInfo1,sInfo2)
	GBL_FAILED_FUNCTION_NAME="Fn_MSO_CreateFolder"
	Dim sValue,xParentCord1,yParentCord1,wshshell,objWord,objNewFolder,xParentCord2,yParentCord2
   	Dim sItemId,sGeneratedItemName,sRev,iCount,aLogin,sCreateCounter,objExcel
    	Dim sLeft2,Top2,sRight2,sBottom2, objPowerPoint
	
	Fn_MSO_CreateFolder = False
	'Check if the Object Exists for Word Or Excel or PowerPoint
	Set objNewFolder = WpfWindow("CreateNewFolder")
	Set objWord = Window("MicrosoftWord")
	Set objExcel = Window("MicrosoftExcel")
	Set objPowerPoint = Window("MicrosoftPowerPoint")
	
	'Set Action name
	If Instr(aUser,"$")>0Then
		aaUser = Split(aUser,"$")
		sAction = aaUser(0)
		aUser = aaUser(1)
	Else
		sAction = "Create"
		aUser = aUser
	End If
	
	If Not objNewFolder.Exist(5)  Then
		If objWord.Exist(5) Then
			sAppType = "MSWord"
			Set objAppType = objWord
			Set objWord = Nothing
			Set objExcel = Nothing
			Set objPowerPoint = Nothing
		ElseIf objExcel.Exist(5) Then
			sAppType = "MSExcel"
			Set objAppType = objExcel
			Set objExcel = Nothing
			Set objWord = Nothing
			Set objPowerPoint = Nothing
		ElseIf objPowerPoint.Exist(5) Then
			sAppType = "MSPowerPoint"
			Set objAppType = objPowerPoint
			Set objPowerPoint = Nothing
			Set objWord = Nothing
			Set objExcel = Nothing
		Else
			Set objNewFolder = Nothing
			Set objWord = Nothing
			Set objExcel = Nothing
			Set objPowerPoint = Nothing
			Exit Function
		End If
				
		Set wshshell = CreateObject("WScript.Shell")
		wshshell.SendKeys "%"
		wshshell.SendKeys "Y"
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Successfully Activated the Teamcenter Tab")
		Fn_MSO_CreateFolder=true
		Set wshshell = Nothing
		wait 3
		
		bREturn = Fn_MSO_RibbonButton_Operations(sAppType,"Click","New:Folder","")
		If bREturn=True Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Successfully performed Menu operation [New:Folder]")
			Fn_MSO_CreateFolder = True
		Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to perform Menu operation [New:Folder]")
			Fn_MSO_CreateFolder = False
			Set objAppType = Nothing
			Set objNewFolder = Nothing
			Exit function
		End If
		wait 5
		'Login into Teamcenter
		If WpfWindow("Teamcenter Login").Exist(5) Then
			aLogin=split(aUser,":",-1,1)
			If sAppType = "MSWord" Then
				bReturn = Fn_MSO_TeamcenterLogin("WpfWordLogin", aLogin(0), aLogin(0), aLogin(2), aLogin(3) )
			ElseIf sAppType = "MSExcel" Then
				bReturn = Fn_MSO_TeamcenterLogin("WpfExcelLogin", aLogin(0), aLogin(0), aLogin(2), aLogin(3) )
			ElseIf sAppType = "MSPowerPoint" Then
				bReturn = Fn_MSO_TeamcenterLogin("WpfPowerPointLogin", aLogin(0), aLogin(0), aLogin(2), aLogin(3) )
			End If
			If bReturn=True Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Successfully logged in with user ["+aLogin(0)+"]")
				Fn_MSO_CreateFolder = True
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to log in with user ["+aLogin(0)+"]")
				Fn_MSO_CreateFolder = False
				Set objAppType = Nothing
				Set objNewFolder = Nothing
				Exit Function
			End If
		End If

		If not objNewFolder.Exist(8) Then
			wait 10
		End If
	End If
	wait 5
	
	Select Case sAction
		'Fn_MSO_CreateFolder(aUser,sFolderName,sFolderDesc,sButtons,sInfo1,sInfo2)
		Case "Verify"
				dicCount = sInfo2.Count
				dicItems = sInfo2.Items
				dicKeys = sInfo2.Keys
				
				For iCounter = 0 To dicCount - 1
					sSubAction = dicKeys(iCounter)
					sProperty = dicItems(iCounter)
					bFlag = False
					Select Case sSubAction
						'Case to Verify Selected Location
						Case "VerifySelectedLocation"
							If sProperty<>"" Then
								sAppText = WpfWindow("CreateNewFolder").WpfObject("SelectedLocation").GetROProperty("text")
								If Trim(sProperty)=Trim(sAppText) Then
									bFlag = True
								End If
							End If
						Case Else
							Fn_MSO_CreateFolder = False
							Set objAppType = Nothing
							Set objNewFolder = Nothing
							Exit Function
					End Select
					
					If bFlag = False Then
						Fn_MSO_CreateFolder = False
						Set objAppType = Nothing
						Set objNewFolder = Nothing
						Exit Function
					End If
				Next
				
				'Click on button provided
				If sButtons<>"" Then
					bFlag = Fn_MSO_WpfButton_Click("Fn_MSO_CreateFolder","Click",objSSTeam,sButtons)
					If bFlag = False Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Click on WPFButton ["+sButtons+"].")
						Fn_MSO_CreateFolder = False
						Set objAppType = Nothing
						Set objNewFolder = Nothing
						Exit Function
					End If
					Wait 1
				End If
				Fn_MSO_CreateFolder = True
		Case "Create"
				'Set the Object Type Folder or else, Used parameter sInfo1 for this
				If sInfo1<>"" Then
					objNewFolder.WpfComboBox("ObjectType").Select sInfo1
					If Err.Number<0 Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to select Object Type ["+sInfo1+"].")
						Fn_MSO_CreateFolder = False
						Set objAppType = Nothing
						Set objNewFolder = Nothing
						Exit Function
					End If
					Wait 2
				End If
				
				'Set the Select Location, Used parameter sInfo2 for this
				If sInfo2<>"" Then
					'Future Use
				End If
				
				'***********	'Click on Next Button	**************
				'Check Next button is enabled or not, if enabled then click on Next button
				If objNewFolder.WpfButton("Next").GetROProperty("enabled") = True Then
					objNewFolder.WpfButton("Next").Click 10,5,micLeftBtn
					If Err.Number<0 Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to Click on Next button.")
						Fn_MSO_CreateFolder = False
						Set objAppType = Nothing
						Set objNewFolder = Nothing
						Exit Function
					End If
					Wait 3
				End If
				'**********************************************
			
				 'Set the FolderName in the NAme Edit Box
				If sFolderName<>"" Then
					objNewFolder.WpfObject("PropertyText").SetTOProperty "devname","Folder Name:"
					objNewFolder.WpfEdit("PropertyEditBox").Set sFolderName
					'objNewFolder.WpfEdit("Name").Set sFolderName
					If Err.Number<0 Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to set the Folder Name as ["+sFolderName+"]")
						Fn_MSO_CreateFolder = False
						Set objAppType = Nothing
						Set objNewFolder = Nothing
						Exit Function
					End If
					wait 2
				End If
							
				'Set the Folder Description
				If sFolderDesc<>"" Then
					objNewFolder.WpfObject("PropertyText").SetTOProperty "devname","Description"
					objNewFolder.WpfEdit("PropertyEditBox").Set sFolderDesc
					'objNewFolder.WpfEdit("Description").Set sFolderDesc
					If Err.Number<0 Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to set the Folder Description as ["+sFolderDesc+"]")
						Fn_MSO_CreateFolder = False
						Exit Function
					End If
					wait 2
				End If
				
				If sButtons<>"" Then
					If Instr(sButtons,":")>0 Then
						aButtons=split(sButtons,":",-1,1)
						For iCount=0 to uBound(aButtons)
							objNewFolder.WpfButton(aButtons(iCount)).Click 10,5,micLeftBtn
							If Err.Number<0 Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to Click on the Button  ["+aButtons(iCount)+"]")
								Fn_MSO_CreateFolder = False
								Set objAppType = Nothing
								Set objNewFolder = Nothing
								Exit Function
							End If
							Wait 2
						Next
					Else
						objNewFolder.WpfButton(sButtons).Click 10,5,micLeftBtn
						If Err.Number<0 Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to Click on the Button  ["+sButtons+"]")
							Fn_MSO_CreateFolder = False
							Set objAppType = Nothing
							Set objNewFolder = Nothing
							Exit Function
						End If
						Wait 2
					End If
				Else
					objNewFolder.WpfButton("OK").Click 10,5,micLeftBtn
					If Err.Number<0 Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to Click on the Button  [OK]")
						Fn_MSO_CreateFolder = False
						Set objAppType = Nothing
						Set objNewFolder = Nothing
						Exit Function
					End If
				End If
			
				'Check if the New Folder Window Exists
				If objNewFolder.Exist(5) Then
					objNewFolder.WpfButton("Close").Click 10,5,micLeftBtn
				End If
			
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Successfully Created Folder with name ["+sFolderName+"]")			
				Fn_MSO_CreateFolder = True
	End Select
	
	Set objAppType = Nothing
	Set objNewFolder = Nothing
End Function

''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''''/$$$$
''''/$$$$   FUNCTION NAME   :  Fn_MSO_SpecificationOperations(sAction,aUser,sSpecType,sSpecName,sInfo1,sInfo2)
''''/$$$$
''''/$$$$   DESCRIPTION        :  This function will create an Requirement In MS-Word (Provided there is Teamcenter Addin For Ms-Office Installed on the Client)
''''/$$$$ 
''''/$$$$   PRE-REQUISITES        :  MS-Word Window Should be in Focus
''''/$$$$
''''/$$$$  PARAMETERS   : 		sAction : Valid Action Name
''''/$$$$												aUser :Login Details
''''/$$$$ 											sSpecName: valid Specification Name
''''/$$$$										sSpecType : valid specification type
''''/$$$$										sInfo1: For Future Use
''''/$$$$										sInfo2 : For Future Use
''''/$$$$	
''''/$$$$		Return Value : 				True or False
''''/$$$$
''''/$$$$    Function Calls       :   Fn_WriteLogFile(), Fn_MenuOperation(), Fn_ReadyStatusSync
''''/$$$$										
''''/$$$$
''''/$$$$		HISTORY           :   		AUTHOR                 DATE        VERSION
''''/$$$$
''''/$$$$    CREATED BY     :   SHREYAS          21/02/2012         1.0
''''/$$$$
''''/$$$$    REVIWED BY     :  Shreyas			21/02/2012          1.0
''''/$$$$
''''/$$$$		How To Use :    bReturn=Fn_MSO_SpecificationOperations("NewSpec",Environment.Value("TcUserDBA"),"RequirementSpec","TestSpec",sInfo1,sInfo2)
''''/$$$$							
''''/$$$$	
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

Public Function Fn_MSO_SpecificationOperations(sAction,aUser,sSpecType,sSpecName,sInfo1,sInfo2)
	GBL_FAILED_FUNCTION_NAME="Fn_MSO_SpecificationOperations"
   Dim sValue,xParentCord1,yParentCord1,wshshell,objWord,objNewSpec,xParentCord2,yParentCord2
   	Dim sItemId,sGeneratedItemName,sRev,iCount,aLogin,sCreateCounter
    dim sLeft2,Top2,sRight2,sBottom2

'Check if the Object Exists for Word Or Excel

	Set objNewSpec= WpfWindow("CreateSpecification")
	Set wshshell=CreateObject("WScript.Shell")
	Set objWord= Window("MicrosoftWord").WinObject("NetUIHWND")
	Fn_MSO_SpecificationOperations=false

		Select Case sAction
		
			Case "NewSpec"
				
		'					If objWord.Exist(5) Then
							
							'	 sValue=objWord.GetTextLocation ("Teamcenter",sLeft1,Top1,sRight1,sBottom1,True)
							'	 If sValue=True Then
							'			   xParentCord1 = (sLeft1+sRight1) / 2+120
							'			   yParentCord1=(Top1+sBottom1) / 2 
							'			   objWord.DblClick xParentCord1,yParentCord1,micLeftBtn
							'			   Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Successfully Activated the Teamcenter Tab")
							'			   Fn_MSO_ItemBasicCreate=true
							'	Else
							'			   Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to Activate the Teamcenter Tab")
							'			   Fn_MSO_ItemBasicCreate=False
							'			   Exit function
							'	 End If
							'	wait 3
							
							wshshell.SendKeys "%"
							wshshell.SendKeys "Y2"
                            wshshell.SendKeys "YE"
										   Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Successfully Activated the Teamcenter Tab")
										   Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Successfully Clicked on the Specifications Button")
										   Fn_MSO_SpecificationOperations=true
							wait 3'
								
'								 sValue= objWord.GetTextLocation ("Specifications",sLeft2,Top2,sRight2,sBottom2,True)
'								 If sValue=True Then
'										   xParentCord2 = (sLeft2+sRight2) / 2 -35
'										   yParentCord2 =(Top2+sBottom2) / 2 
'										   objWord.Click xParentCord2,yParentCord2,micLeftBtn
'										   Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Successfully Clicked on the Specifications Button")
'										   Fn_MSO_SpecificationOperations=true
'								Else
'										   Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to Click on the Specifications Button")
'										   Fn_MSO_SpecificationOperations=False
'										   Exit function
'								 End If
'									wait 1
							
								
							'Select the New Option
									For iCount=1 to 1
										wshshell.SendKeys "{TAB}"
										wait 1
									Next
									wait 1
									wshshell.SendKeys "{ENTER}"
									wait 3
								
								
							'Login into Teamcenter
							If aUser <> "" Then
								aLogin=split(aUser,":",-1,1)
								bReturn=Fn_MSO_TeamcenterLogin("WpfExcelLogin", aLogin(0), aLogin(0), aLogin(2), aLogin(3) )
								If bREturn=True Then
										   Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Successfully logged in with user ["+aLogin(0)+"]")
										   Fn_MSO_SpecificationOperations=true
								Else
										   Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to log in with user ["+aLogin(0)+"]")
										   Fn_MSO_SpecificationOperations=False
										   Exit function
								End If
							End If
								If not objNewSpec.Exist(8) Then
									wait 10
								End If
							
							wait 5
		
						'	Select the Desired Specification Type
						objNewSpec.WpfComboBox("List").Select sSpecType
						If err.number<0 Then
								   Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to Select the Specification Type ["+sSpecType+"]")
								   Fn_MSO_SpecificationOperations=False
								   Exit function
						End If
						wait 3
		
						'Set the Specification Name
						 If sSpecName<>"" Then
							objNewSpec.WpfEdit("Name").Set sSpecName
								If err.number<0 Then
								   Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to set the name ["+sSpecName+"]")
								   Fn_MSO_SpecificationOperations=False
								   Exit function
								End If
							 wait 3
						 End If

				'Click on OK Button
				 objNewSpec.WpfButton("OK").Click 5,5,micLeftBtn
				If err.number<0 Then
					   Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to click on the OK button")
					   Fn_MSO_SpecificationOperations=False
					   Exit function
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Successfully Created  ["+sSpecType+"] with name ["+sSpecName+"]")
				End If
				 wait 3
				 
				 
				 Case "NewReq"
							objNewSpec.SetTOProperty "regexpwndtitle","Create New"
							objNewSpec.SetTOProperty "devname","Create New"	
							WpfWindow("CreateSpecification").WpfComboBox("List").SetTOProperty "devname","cboSubType" 
							WpfWindow("CreateSpecification").WpfEdit("Name").SetTOProperty "devname","txtName" 
							
								If not objNewSpec.Exist(8) Then
									wait 10
								End If
								
						'	Select the Desired  Type
						objNewSpec.WpfComboBox("List").Select sSpecType
						If err.number<0 Then
								   Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to Select the Requirement Type ["+sSpecType+"]")
								   Fn_MSO_SpecificationOperations=False
								   Exit function
						End If
						wait 3
		
						'Set the Specification Name
						 If sSpecName<>"" Then
							objNewSpec.WpfEdit("Name").Set sSpecName
								If err.number<0 Then
								   Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to set the name ["+sSpecName+"]")
								   Fn_MSO_SpecificationOperations=False
								   Exit function
								End If
							 wait 3
						 End If

				'Click on OK Button
				 objNewSpec.WpfButton("OK").Click 5,5,micLeftBtn
				If err.number<0 Then
					   Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to click on the OK button")
					   Fn_MSO_SpecificationOperations=False
					   Exit function
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Successfully Created  ["+sSpecType+"] with name ["+sSpecName+"]")
					Fn_MSO_SpecificationOperations = True
					Exit Function
				End If

				Case "ImportSpec"
								'Will be coded as required
		
		End Select
		Set objNewSpec= nothing
		Set wshshell=nothing
		Set objWord=nothing
End function

''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''''/$$$$
''''/$$$$   FUNCTION NAME   :   Fn_MSO_Click_ImportToTeamCenter_Button()
''''/$$$$
''''/$$$$   DESCRIPTION        :  This function will Click on the Import to Teamcenter Button in the teamcenter Ribbon
''''/$$$$ 
''''/$$$$   PRE-REQUISITES        :  MS-Word Window Should be in Focus
''''/$$$$
''''/$$$$  PARAMETERS   : 		This Function has no parameters
''''/$$$$	
''''/$$$$		Return Value : 				True or False
''''/$$$$
''''/$$$$    Function Calls       :   Fn_WriteLogFile(), Fn_MenuOperation()
''''/$$$$										
''''/$$$$
'''/$$$$	Prerequsites:    "Import to Teamcenter menu should be adde in Customaize Toolbar option
'''/$$$$
''''/$$$$		HISTORY           :   		AUTHOR                 DATE        VERSION
''''/$$$$
''''/$$$$    CREATED BY     :   SHREYAS          01/03/2012         1.0
''''/$$$$
''''/$$$$    REVIWED BY     :  Shreyas			01/03/2012          1.0
''''/$$$$ 
''''/$$$$   Modified by :        Nilesh						01/05/2012			2.0
'''/$$$$  
'''/$$$$   Modified by :        Nilesh						10/12/2012			3.0                         'Added Import to Teamcenter button  click using DLL method
''''/$$$$
''''/$$$$		How To Use :    bReturn=Fn_MSO_Click_ImportToTeamCenter_Button()
''''/$$$$							
''''/$$$$		
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

Public function Fn_MSO_Click_ImportToTeamCenter_Button()
GBL_FAILED_FUNCTION_NAME="Fn_MSO_Click_ImportToTeamCenter_Button"
 On Error Resume Next
   Dim bReturn,dr,yAxis,xAxis,x,y
   Dim sLeft,Top,sRight,sBottom,bFlag
   Dim sFileName,sDir,sFilePath
   Dim objExcel,iAddinCount,iCount
   Dim objDll

   Fn_MSO_Click_ImportToTeamCenter_Button=False
   'Maximize Excel window 
'   If Window("MicrosoftExcel").GetRoProperty("maximized")=False Then
'	 Window("MicrosoftExcel").Maximize
'	End If
'
'   'Activate the Teamcenter Tab in the Ribbon
'		If Window("MicrosoftExcel").WinObject("NetUIHWND").Exist(5)	Then	' Code to focus ribbon control
'			Window("MicrosoftExcel").WinObject("NetUIHWND").Highlight
'		End If
'	 	bReturn= Fn_KeyBoardOperation("SendKeys","%~Y")
'		If bReturn=true Then
'			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Successfully Activated the Teamcenter Tab in the Ribbon")
'			Fn_MSO_Click_ImportToTeamCenter_Button=true
'		Else
'		   Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to Activate the Teamcenter Tab in the Ribbon")
'		   Fn_MSO_Click_ImportToTeamCenter_Button=False
'		   Exit function	
'		End If
'
''Code to avoid blurring of Excel
'	Window("MicrosoftExcel").Restore
'	Window("MicrosoftExcel").Maximize
'	Set objExcel=GetObject(,"Excel.Application")
'	sFileName=objExcel.ActiveWorkbook.Name
'	sDir=objExcel.ActiveWorkbook.Path
'	sFilePath=sDir+"\"+sFileName
''for Synchronisation
'		wait 2
'		sLeft=-1
'		Top=-1
'		sRight=-1
'		sBottom=-1
''				Logic to click on the Import to teamcenter button
'			If Window("MicrosoftExcel").WinObject("NetUIHWND").Exist(5)	Then	' Code to focus Ribbon Control.
'				Window("MicrosoftExcel").WinObject("NetUIHWND").Highlight
'			End If
'			 Window("MicrosoftExcel").WinObject("NetUIHWND").GetTextLocation "Import to",sLeft,Top,sRight,sBottom,True
'			 yAxis=(Top+sBottom)/2
'			xAxis=(sLeft+sRight)/2
'			 If yAxis<>0 OR xAxis<>0Then
'				 yAxis=(Top+sBottom)/2
'				 xAxis=(sLeft+sRight)/2
'				 wait 2
''				Set dr = CreateObject("Mercury.DeviceReplay")
'				Set dr=Window("MicrosoftExcel").WinObject("NetUIHWND")
'				wait 3
'				x=xAxis
'				y=Top
''				y=yAxis-25
''				dr.MouseClick x,y,LEFT_MOUSE_BUTTON
'				dr.Click x,y,micLeftBtn
'				Fn_MSO_Click_ImportToTeamCenter_Button=True
'				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Successfully Clicked on Import To Teamcenter Button")
'				bFlag=False
'				wait 3
'				
'			 End If
'' If Location property of an WinObject("NetUIHWND") is changed from 0 to 1 then
'			If yAxis=0 OR  xAxis=0 Then
'				Window("MicrosoftExcel").WinObject("NetUIHWND").SetToProperty "Location",1
'				 bFlag=Window("MicrosoftExcel").WinObject("NetUIHWND").GetTextLocation ("Import to",sLeft,Top,sRight,sBottom,True)
'				 If bFlag=True Then
'					 yAxis=(Top+sBottom)/2
'					 xAxis=(sLeft+sRight)/2
'					 wait 2
'
''					Set dr = CreateObject("Mercury.DeviceReplay")
'					Set dr=Window("MicrosoftExcel").WinObject("NetUIHWND")
'					wait 3
'					x=xAxis
''					y=yAxis-25
'					y=Top
''					dr.MouseClick x,y,LEFT_MOUSE_BUTTON
'					dr.Click x,y,micLeftBtn
'					Fn_MSO_Click_ImportToTeamCenter_Button=True
'					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Successfully Clicked on Import To Teamcenter Button")
'		
'					wait 3
'				Else
'					Fn_MSO_Click_ImportToTeamCenter_Button=False
'					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to Clickon Import To Teamcenter Button")
'				 End If	
'			End If
			'
            '*Added by Nilesh on 12-Dec-2012
			
				If Environment.Value("sPath")="" then
						Environment.Value("sPath") = Fn_GetEnvValue("User", "AutomationDir")  ''Added by Avinash J. 14-Feb-2013
				End If
				Set objDll=DotnetFactory.CreateInstance("MsRibbonMng.msRibbon.clickRibbon",Environment.Value("sPath")+"\Library\MsRibbonMng.dll")
				objDll.SetMsApplication "Excel"
				Wait 2
				objDll.ClickRibbonButton "Import to Teamcenter"
			'*End
			bReturn=False
			If Window("MicrosoftExcel").GetRoProperty("enabled") =False Then
				bReturn=True
			End If


'Check Import to Teamcenter click is successful or not 
			If WpfWindow("Teamcenter Login").Exist(5)=True OR bReturn=True Then
				Fn_MSO_Click_ImportToTeamCenter_Button=True
			Else
			'Check TcExcelAddin status (Enable /Disable)
				Set objExcel=CreateObject("Excel.Application")
				iAddinCount=objExcel.COMAddIns.Count
				For iCount=1 To iAddinCount
					If Instr(Lcase(objExcel.COMAddIns(iCount).ProgId),"tcexceladdin")>0 Then
						bFlag=objExcel.COMAddIns(iCount).Connect
					End If
				Next
			'If Addin is disabled then code to enable it
				If bFlag=False Then
					Call Fn_MSO_FileOperations("FileClose" ,"Excel", "", "")
					Call Fn_EnableTcExcelAddin()
					Call Fn_MSO_ExcelEditOperations("OpenExcel", sFilePath ,1,"","","")
					Call Fn_MSO_Click_ImportToTeamCenter_Button()
				End If
				Set objExcel=Nothing
		
			End If
			
	Set dr=Nothing
	Set objDll=Nothing
End Function

''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''''/$$$$
''''/$$$$   FUNCTION NAME   :  Fn_MSO_RenameSheetColumnHeader(sSheetIndex,aFindString,sColToFind,sNewColName,sInfo1,sInfo2)
''''/$$$$
''''/$$$$   DESCRIPTION        :  This function will replace the Column Header with the Desired New Header
''''/$$$$ 
''''/$$$$   PRE-REQUISITES        :  Valid Excel document on which operations need to be performed should exist 
''''/$$$$
''''/$$$$  PARAMETERS   : 		sSheetIndex : Index / Number of the Sheet to be activated
''''/$$$$												aFindString :Array of string needed to be found
''''/$$$$ 											sColToFind: Column name to be Found
''''/$$$$										sNewColName : New Column name to be assigned
''''/$$$$										sInfo1: For Future Use
''''/$$$$										sInfo2 : For Future Use
''''/$$$$	
''''/$$$$		Return Value : 				True or False
''''/$$$$
''''/$$$$    Function Calls       :   Fn_WriteLogFile(), Fn_MenuOperation(), Fn_ReadyStatusSync
''''/$$$$										
''''/$$$$
''''/$$$$		HISTORY           :   		AUTHOR                 DATE        VERSION
''''/$$$$
''''/$$$$    CREATED BY     :   SHREYAS          08/03/2012         1.0
''''/$$$$
''''/$$$$    REVIWED BY     :  Shreyas			08/03/2012           1.0
''''/$$$$
''''/$$$$		How To Use :    bReturn=Fn_MSO_RenameSheetColumnHeader(2,"object_string,last_mod_date,owning_user","Active","Ignored","","")
''''/$$$$							
''''/$$$$	
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

Public function Fn_MSO_RenameSheetColumnHeader(sSheetIndex,aFindString,sColToFind,sNewColName,sInfo1,sInfo2)
	GBL_FAILED_FUNCTION_NAME="Fn_MSO_RenameSheetColumnHeader"
	Dim objExcel,i,bReturn,sValue,ColNumber,s,aValues
	Fn_MSO_RenameSheetColumnHeader=False
	
	Set objExcel=getobject(,"Excel.Application")
	objExcel.visible=true
	set i= objExcel.Worksheets(2)

aValues=split(aFindString,",",-1,1)
For Counter=0 to uBound(aValues)
		bReturn= Fn_MSO_ExcelEditOperations("GetCellPosition","",sSheetIndex,"",aValues(Counter),"")
		If bReturn<>False Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Successfully Retrieved the Cell Position as ["+bReturn+"]")
				Fn_MSO_RenameSheetColumnHeader=True
		Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to Retrieve Cell Position For String ["+sFindString+"]")
				Fn_MSO_RenameSheetColumnHeader=False
				Exit Function
	   End If
	
		sValue = mid(bReturn,1,1)
		ColNumber= Fn_MSO_ExcelColHeader("GetColumnHeaderNumber",sValue)
		sValue=i.usedrange.rows.count
		For s=1 to sValue
				If objExcel.Cells(s,ColNumber).value=sColToFind Then
					objExcel.Cells(s,ColNumber).value=sNewColName
				End If
		Next
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"SuccessFully Replaced Column name ["+sColToFind+"] with ["+sNewColName+"]")
Next
End Function



''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''''/$$$$
''''/$$$$   FUNCTION NAME   : Fn_MSO_OutlineOperations(sAction,sInfo1,sInfo2)
''''/$$$$
''''/$$$$   DESCRIPTION        :  This function will perform Show & Hide Operations on the Outline at the moment
''''/$$$$ 
''''/$$$$   PRE-REQUISITES        :  The Excel Window Should be present in the syatem & should be in focus
''''/$$$$
''''/$$$$  PARAMETERS   : 		sAction : Action to be performed
''''/$$$$										sInfo1: For Future Use
''''/$$$$										sInfo2 : For Future Use
''''/$$$$	
''''/$$$$		Return Value : 				True or False
''''/$$$$
''''/$$$$    Function Calls       :   Fn_WriteLogFile(), Fn_KeyBoardOperation(),
''''/$$$$										
''''/$$$$
''''/$$$$		HISTORY           :   		AUTHOR                 DATE        VERSION
''''/$$$$
''''/$$$$    CREATED BY     :   SHREYAS          12/03/2012         1.0
''''/$$$$
''''/$$$$    REVIWED BY     :  Shreyas			12/03/2012            1.0
''''/$$$$
''''/$$$$		How To Use :    bReturn=Fn_MSO_OutlineOperations("ShowDetails","","")
''''/$$$$							
''''/$$$$	
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$


Public function Fn_MSO_OutlineOperations(sAction,sInfo1,sInfo2)
	GBL_FAILED_FUNCTION_NAME="Fn_MSO_OutlineOperations"
	Fn_MSO_OutlineOperations=false
   Dim bReturn,objExcel

   Set objExcel=Window("MicrosoftExcel")

'check if the Excel Window Exists
	   If objExcel.Exist(10) Then
		   Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"The Excel Window Exists on the System")
		   Fn_MSO_OutlineOperations=True
		Else
		   Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"The Excel Window does not Exist on the System")
		   Fn_MSO_OutlineOperations=False
			Exit Function
	   End If


		Select Case sAction

				Case "ShowDetails"
					set objExcel=GetObject(,"Excel.Application")
					objExcel.visible=True
					Wait 3

					'Now click on the Show Details Button
					bReturn= Fn_KeyBoardOperation("SendKeys","%~A~J")
					If bReturn=true Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Successfully Showed the Outline Details")
							Fn_MSO_OutlineOperations=true
					Else
						   Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to Show the Outline Details")
						   Fn_MSO_OutlineOperations=False
						   Exit function	
					End If


				Case "HideDetails"
					Set objExcel=GetObject(,"Excel.Application")
					objExcel.visible=True
					Wait 3

					'Now click on the Show Details Button
					bReturn= Fn_KeyBoardOperation("SendKeys","%~A~H")
					If bReturn=true Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Successfully Hide the Outline Details")
							Fn_MSO_OutlineOperations=true
					Else
						   Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to Hide the Outline Details")
						   Fn_MSO_OutlineOperations=False
						   Exit function	
					End If

			Case Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_OutlineOperations: Invalid case name [ " & sAction & " ].")
				   Fn_MSO_OutlineOperations=False
				   Exit function
		End Select

		Set ObjExcel=Nothing
End Function

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name	:	Fn_MSO_PPTOperations

'Description		:	Function Used to perform operations of Power point

'Parameters		:   1.StrAction: Action Name
'										2.dicPPTInfo: Power point information
'
'Return Value		: 	True or False

'Pre-requisite		:	

'Examples		:  	dicPPTInfo("PresentationName")="Presentation1"
'					dicPPTInfo("SlideNumber")="1"
'					Msgbox Fn_MSO_PPTOperations("GetTitleText",dicPPTInfo)
'					
'					dicPPTInfo("PresentationName")="Presentation1"
'					dicPPTInfo("SlideNumber")="1"
'					dicPPTInfo("Title")="Function Fn_MSO_PPTOperations Demo"
'					dicPPTInfo("BodyText")="Testing Function Fn_MSO_PPTOperations"
'					Msgbox Fn_MSO_PPTOperations("Write",dicPPTInfo)
'					
'					dicPPTInfo("PresentationName")="Presentation1"
'					dicPPTInfo("SlideNumber")="1"
'					Msgbox Fn_MSO_PPTOperations("GetBodyText",dicPPTInfo)
'					
'					dicPPTInfo("PresentationName")="Presentation1"
'					dicPPTInfo("SlideNumber")="1"
'					Msgbox Fn_MSO_PPTOperations("GetText",dicPPTInfo)
'					
'					dicPPTInfo("SlideNumber")="2"
'					Msgbox Fn_MSO_PPTOperations("TcGetText",dicPPTInfo)
'					
'					dicPPTInfo("SlideNumber")="1"
'					Msgbox Fn_MSO_PPTOperations("TcGetTitleText",dicPPTInfo)
'					
'					dicPPTInfo("SlideNumber")="2"
'					Msgbox Fn_MSO_PPTOperations("TcGetBodyText",dicPPTInfo)
'
'History		:			
'			Developer Name		Date			Rev. No.		Changes Done									Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'			Sandeep N		21-Mar-2012		1.0														Sunny R
'			Sandeep N		26-Mar-2012		1.1			Added Case "TcGetTitleText","TcGetBodyText","TcGetText"
'			Vivek A			10-Jun-2016		1.2			Added case "OpenPowerPoint" 					[TC1123-20160504-10_06_2016-VivekA-NewDevelopment]
'														Added for RM - Office Client new TC's development
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Public Function Fn_MSO_PPTOperations(StrAction,dicPPTInfo)
	GBL_FAILED_FUNCTION_NAME="Fn_MSO_PPTOperations"
	'Variable declaration
	Dim oPowerPoint,oPresentation,oSlide,sHWNDProperty
	Fn_MSO_PPTOperations=False
	'Creating object of [ PowerPoint ] applicaton
	Set oPowerPoint = CreateObject("PowerPoint.Application")
  	Select Case StrAction
  		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  		Case "OpenPowerPoint"
				'If filepath is blank then set default
				'If dicPPTInfo("PresentationName") = "" Then dicPPTInfo("PresentationName") = "POWERPNT.exe"
				Wait 10
				oPowerPoint.Visible = True
				If dicPPTInfo("PresentationName") <> "" Then 
					Set oObjPresentations = oPowerPoint.Presentations.Open(dicPPTInfo("PresentationName"))
				Else
					Set oObjPresentations = oPowerPoint.Presentations.Add
					oObjPresentations.Slides.Add 1, 11
				End If
				
				If Window("MicrosoftPowerPoint").Exist(5) Then
  					Window("MicrosoftPowerPoint").Maximize
  					Wait 1
  				End If
				Fn_MSO_PPTOperations = True
				Set oPowerPoint = Nothing
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "AddSlide"
				oPowerPoint.Visible = True
				If dicPPTInfo("PresentationName")<>"" Then
					Set oPresentation=oPowerPoint.Presentations(dicPPTInfo("PresentationName"))
				Else
					Set oPresentation=oPowerPoint.ActivePresentation
				End If
			
				If dicPPTInfo("SlideNumber") <> "" Then
					For i = 1 To dicPPTInfo.Item("SlideNumber")
						oPresentation.Slides.Add 1, 11
					Next					
				End If
				Fn_MSO_PPTOperations = True
				Set oPowerPoint = Nothing
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to read text from slides
		Case "GetTitleText","GetBodyText","GetText"
				oPowerPoint.Visible=True
				'Creating object of presentation
				If dicPPTInfo("PresentationName")<>"" Then
					Set oPresentation=oPowerPoint.Presentations(dicPPTInfo("PresentationName"))
				Else
					Set oPresentation=oPowerPoint.ActivePresentation
				End If
				'Creating object of slide
				If dicPPTInfo("SlideNumber")<>"" Then
					Set oSlide=oPresentation.Slides(CInt(dicPPTInfo("SlideNumber")))
				Else
					Set oSlide=oPresentation.Slides(1)
				End If
				oSlide.Select
				If StrAction="GetTitleText" Then
					Fn_MSO_PPTOperations=oSlide.Shapes(1).TextFrame.TextRange.Text
				ElseIf StrAction="GetBodyText" Then
					Fn_MSO_PPTOperations=oSlide.Shapes(2).TextFrame.TextRange.Text
				ElseIf StrAction="GetText" Then
					Fn_MSO_PPTOperations=oSlide.Shapes(1).TextFrame.TextRange.Text+"~"+oSlide.Shapes(2).TextFrame.TextRange.Text
				End If
				If LCase(dicPPTInfo("PresentationCloseFlag"))<>"false" Then
					oPresentation.Close
				End If
				If LCase(dicPPTInfo("PowerPointQuitFlag"))<>"false" Then
					oPowerPoint.Quit
				End If
				Set oPresentation=Nothing
				Set oSlide=Nothing
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to write text from slides
		Case "Write"
				oPowerPoint.Visible=True
				'Creating object of presentation
				If dicPPTInfo("PresentationName")<>"" Then
					Set oPresentation=oPowerPoint.Presentations(dicPPTInfo("PresentationName"))
				Else
					Set oPresentation=oPowerPoint.ActivePresentation
				End If
				'Creating object of slide
				If dicPPTInfo("SlideNumber")<>"" Then
					Set oSlide=oPresentation.Slides(CInt(dicPPTInfo("SlideNumber")))
				Else
					Set oSlide=oPresentation.Slides(1)
				End If
				oSlide.Select
				If dicPPTInfo("Title")<>"" Then
					oSlide.Shapes(1).TextFrame.TextRange.Text=dicPPTInfo("Title")
				End If
				If dicPPTInfo("BodyText")<>"" Then
					oSlide.Shapes(2).TextFrame.TextRange.Text=dicPPTInfo("BodyText")
				End If
				oPresentation.Save
				If LCase(dicPPTInfo("PresentationCloseFlag"))<>"false" Then
					oPresentation.Close
				End If
				If LCase(dicPPTInfo("PowerPointQuitFlag"))<>"false" Then
					oPowerPoint.Quit
				End If
				Fn_MSO_PPTOperations=True
				Set oPresentation=Nothing
				Set oSlide=Nothing
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to read text from slides from Teamcenter Tab
		Case "TcGetTitleText","TcGetBodyText","TcGetText"
				'Creating object of presentation	
				Set oPresentation=oPowerPoint.Presentations(1)
				'Creating object of slide
				If dicPPTInfo("SlideNumber")<>"" Then
					Set oSlide=oPresentation.Slides(CInt(dicPPTInfo("SlideNumber")))
				Else
					Set oSlide=oPresentation.Slides(1)
				End If
				If StrAction="TcGetTitleText" Then
					Fn_MSO_PPTOperations=oSlide.Shapes(1).TextFrame.TextRange.Text
				ElseIf StrAction="TcGetBodyText" Then
					Fn_MSO_PPTOperations=oSlide.Shapes(2).TextFrame.TextRange.Text
				ElseIf StrAction="TcGetText" Then
					Fn_MSO_PPTOperations=oSlide.Shapes(1).TextFrame.TextRange.Text+"~"+oSlide.Shapes(2).TextFrame.TextRange.Text
				End If
				Set oPresentation=Nothing
				Set oSlide=Nothing
		'---------------------------------------------------------------------------------
		Case "SetInstance"
			Set oPowerPoint = Window("MicrosoftPowerPoint")
			Set oPowerPoint = Window("text:=.*PowerPoint.*","hwnd:="&clng(dicPPTInfo("HWND")))
			If dicPPTInfo("Activate") = True and oPowerPoint.Exist(3) Then
				oPowerPoint.Activate
				Wait 1				
			End If
			Fn_MSO_PPTOperations = True
		Case "GetHWND"	
			Set oPowerPoint = Window("MicrosoftPowerPoint")
			sHWNDProperty = oPowerPoint.GetROProperty("hwnd")
			Fn_MSO_PPTOperations = sHWNDProperty
	End Select
	Set oPowerPoint =Nothing
End Function

''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''''/$$$$
''''/$$$$   FUNCTION NAME   :  Fn_MSO_MarkupOperations(sAction,sNodeToFind,sMarkupText,sLoginUserDetails,bSaveMarkup,sInfo1,sInfo2)
''''/$$$$
''''/$$$$   DESCRIPTION        :  This function will create an Markup on the Specified Node
''''/$$$$  
''''/$$$$   PRE-REQUISITES        :  Login to Teamcenter in Office Client Should be done
''''/$$$$
''''/$$$$  PARAMETERS   : 		sAction : Valid Action Name
''''/$$$$											sNodeToFind :Node to create Markup
''''/$$$$ 											sMarkupText: Text to be inserted while creating a markup
''''/$$$$										sLoginUserDetails : Logged in User name (upto 8 characters starting from the First character in lowercase)
''''/$$$$										bSaveMarkup : To save or not to save the markup
''''/$$$$										sInfo1: For Future Use
''''/$$$$										sInfo2 : For Future Use
''''/$$$$	
''''/$$$$		Return Value : 				True or False
''''/$$$$
''''/$$$$    Function Calls       :   Fn_WriteLogFile(), Fn_MenuOperation(), Fn_ReadyStatusSync
''''/$$$$										
''''/$$$$
''''/$$$$		HISTORY           :   		AUTHOR                 DATE        VERSION
''''/$$$$
''''/$$$$    CREATED BY     :   SHREYAS          26/03/2012         1.0
''''/$$$$
''''/$$$$    REVIWED BY     :  Shreyas			26/03/2012           1.0
''''/$$$$
''''/$$$$		How To Use :   bReturn=Fn_MSO_MarkupOperations("CreateMarkup","req2","New Markup","","ON","YES","","")
''''/$$$$									bReturn=Fn_MSO_MarkupOperations("VerifyMarkup","","amol","","","YES","","")
''''/$$$$	bReturn=Fn_MSO_MarkupOperations("ModifyMarkup","","NewQwerty,"",ON","YES","","")
''''/$$$$	bReturn=Fn_MSO_MarkupOperations("DeleteMarkup","","","",ON","YES","","")
''''/$$$$							
''''/$$$$	
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$


Public Function Fn_MSO_MarkupOperations(sAction,sNodeToFind,sMarkupText,sLoginUserDetails,bSaveMarkup,bCloseUI,sInfo1,sInfo2)
	GBL_FAILED_FUNCTION_NAME="Fn_MSO_MarkupOperations"
   Dim objWord,objWpf,bReturn,iCounter,objWindow,objWarning
   Dim sLeft,Top,sRight,sBottom,xParentCord,yParentCord,i,aDetails
   Dim sActual,cmntObj
	Fn_MSO_MarkupOperations=false

		Set objWord=Window("MicrosoftWord").WinObject("MicrosoftWordDocument")
		Set objWindow=Window("MicrosoftWord")
		set objWpf=Window("MicrosoftWord").SwfObject("ControlAxSourcingSite").SwfObject("SwfObject").WpfWindow("WpfWindow")
		Set objWarning=Window("MicrosoftWord").Dialog("Information")
		objWarning.SetTOProperty "text","Warning"
		objWindow.Maximize
		wait 1



		Select Case sAction

			Case "CreateMarkup"
			
								'Activate the MArkup UI in Word
					If objWpf.Exist = False Then	
							bReturn= Fn_KeyBoardOperation("SendKeys", "%Y2YB") ' Changes done as per MS Word 2016 (TC12.2 - 20190318)
							If bReturn=true Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Successfully Activated the Markup UI in the Teamcenter Ribbon")
								Fn_MSO_MarkupOperations=true
							Else
							   Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to Activate the Markup UI in the Teamcenter Ribbon")
							   Fn_MSO_MarkupOperations=False
							   Exit function	
							End If
					End If
						wait 2
						
					If objWpf.Exist(2) = False Then	
							bReturn= Fn_KeyBoardOperation("SendKeys", "%Y2YB") ' Changes done as per MS Word 2016 (TC12.2 - 20190318)
							If bReturn=true Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Successfully Activated the Markup UI in the Teamcenter Ribbon")
								Fn_MSO_MarkupOperations=true
							Else
							   Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to Activate the Markup UI in the Teamcenter Ribbon")
							   Fn_MSO_MarkupOperations=False
							   Exit function	
							End If
					End If
					
					wait 2

							objWord.GetTextLocation sNodeToFind,sLeft,Top,sRight,sBottom,True
							xParentCord = (sLeft+sRight) / 2
							yParentCord=(Top+sBottom) / 2
							objWord.Click xParentCord,yParentCord, micLeftBtn
	
							wait 2
	
					' press the  down arrow key twice
					For i= 1 to 2  ' TC112-2015070800-21_07_2015-Porting-VivekA-As per design change while exporting object to word
							bReturn= Fn_KeyBoardOperation("SendKeys", "{DOWN}")
							If bReturn=true Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Successfully pressed the Down Arrow Key "+iCounter+" times")
								Fn_MSO_MarkupOperations=true
							Else
							   Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to click the Down arrow key")
							   Fn_MSO_MarkupOperations=False
							   Exit function	
							End If
					Next
					wait 2
	
				'Select the TextTo MArkup
						bReturn= Fn_KeyBoardOperation("SendKeys", "{HOME}")
						If bReturn=true Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Successfully selected the Text to Markup")
							Fn_MSO_MarkupOperations=true
						Else
						   Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to select the Text to Markup")
						   Fn_MSO_MarkupOperations=False
						   Exit function	
						End If

						bReturn= Fn_KeyBoardOperation("SendKeys", "+{END}+{LEFT}+{LEFT}")
						If bReturn=true Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Successfully selected the Text to Markup")
							Fn_MSO_MarkupOperations=true
						Else
						   Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to select the Text to Markup")
						   Fn_MSO_MarkupOperations=False
						   Exit function	
						End If
				'Open the PopupMenu 
            	objWpf.WpfList("listView").ShowContextMenu
				wait 2
				call Fn_KeyBoardOperation("SendKeys", "{TAB}")
				wait 2
				call Fn_KeyBoardOperation("SendKeys", "{Enter}")
				wait 1
				'Select the Option to create markup
				call Fn_KeyBoardOperation("SendKeys", "{TAB}")
				wait 1
				call Fn_KeyBoardOperation("SendKeys", "{TAB}")
				wait 2	
	
			objWpf.WpfEdit("txtBox").Type sMarkupText
			If err.number<0 Then
					   Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to set the Markup Text ["+sMarkupText+"]")
					   Fn_MSO_MarkupOperations=False
					   Exit function
			Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Successfully set the Markup Text ["+sMarkupText+"]")
						Fn_MSO_MarkupOperations=true	
			End If
	
				'Click on the Check Icon to Create Markup
				objWpf.WpfImage("imgDone").Click 5,5,micLeftBtn
				If err.number<0 Then
						   Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to create Markup ["+sMarkupText+"]")
						   Fn_MSO_MarkupOperations=False
						   Exit function
				Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Successfully created Markup ["+sMarkupText+"]")
							Fn_MSO_MarkupOperations=true	
				End If

				If objWindow.Dialog("MarkupError").Exist Then
					objWindow.Dialog("MarkupError").WinButton("OK").Click 0,0,micLeftBtn
					   Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to create Markup ["+sMarkupText+"]")
					   Fn_MSO_MarkupOperations=False
					   Exit function
				End If

				If bSaveMarkup<>"" Then
					If lcase(bSaveMarkup)="on" Then
						Call Fn_UI_WpfButtonClick("Fn_MSO_MarkupOperations", objWpf, "Save/Extract")
							If err.number<0 Then
									   Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to save the markup Markup ["+sMarkupText+"]")
									   Fn_MSO_MarkupOperations=False
									   Exit function
							Else
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Successfully saved Markup ["+sMarkupText+"]")
										Fn_MSO_MarkupOperations=true	
							End If
					End If
				End If

				If objWarning.Exist then
					objWarning.WinButton("OK").SetTOProperty "text",sInfo1
					wait 2
					objWarning.WinButton("OK").Click 5,5,micLeftBtn
							If err.number<0 Then
									   Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to handle the [Warning] Dialog")
									   Fn_MSO_MarkupOperations=False
									   Exit function
							Else
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Successfully handled the [Warning] Dialog")
										Fn_MSO_MarkupOperations=true	
							End If
				End if

								'De-Activate the Markup UI
								If objWpf.Exist(1) = True Then	
									objWord.Click 2,2,micLeftBtn
									bReturn= Fn_KeyBoardOperation("SendKeys", "%YYA")
									If bReturn=true Then
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Successfully De-Activated the Markup UI in the Teamcenter Ribbon")
											Fn_MSO_MarkupOperations=true
									Else
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to De-Activate the Markup UI in the Teamcenter Ribbon")
											Fn_MSO_MarkupOperations=False
											Exit function	
									End If
								End If


						Case "VerifyMarkup"

											'De-Activate the Markup UI
'											bReturn= Fn_KeyBoardOperation("SendKeys", "%YY08")
'											If bReturn=true Then
'												Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Successfully De-Activated the Markup UI in the Teamcenter Ribbon")
'												Fn_MSO_MarkupOperations=true
'											Else
'											   Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to De-Activate the Markup UI in the Teamcenter Ribbon")
'											   Fn_MSO_MarkupOperations=False
'											   Exit function	
'											End If
'
'							sDetails=objWord.GetVisibleText
'							aDetails=split(sDetails,vbnewline,-1,1)
'							For i=0 to uBound(aDetails)
'										If instr(1,aDetails(i),trim(sMarkupText))>0 Then
'													bFlag=true
'													Exit For
'										Else
'													bFlag=False
'										End If
'							Next
											'Added by Nilesh on 4-March-2013
											Set objWord=GetObject(,"Word.Application")
                                            For Each cmntObj In objWord.ActiveDocument.Comments
												sActual=cmntObj.Range.Text
												If instr(1,Trim(sActual),trim(sMarkupText))>0 Then
															bFlag=true
															Exit For
												Else
															bFlag=False
												End If
											Next				

							If bFlag=true Then
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Successfully verified that the markup ["+sLoginUserDetails+"] is successfully created" )
											Fn_MSO_MarkupOperations=true
							Else
									   Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to verified that the markup ["+sLoginUserDetails+"] is successfully created" )
									   Fn_MSO_MarkupOperations=False
									   Exit function
							End If

							Case "ModifyMarkup"
				
												
								'Open the PopupMenu 
							
								objWpf.WpfList("listView").ShowContextMenu
								
								'Select the Option to create markup
								call Fn_KeyBoardOperation("SendKeys", "{TAB}")
								wait 1
								call Fn_KeyBoardOperation("SendKeys", "{TAB}")
								wait 1
								call Fn_KeyBoardOperation("SendKeys", "{Enter}")
								wait 1
					
					
							objWpf.WpfEdit("txtBox").Set sMarkupText
							If err.number<0 Then
									   Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to set the Markup Text ["+sMarkupText+"]")
									   Fn_MSO_MarkupOperations=False
									   Exit function
							Else
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Successfully set the Markup Text ["+sMarkupText+"]")
										Fn_MSO_MarkupOperations=true	
							End If
					
								'Click on the Check Icon to Create Markup
								objWpf.WpfImage("imgDone").Click 5,5,micLeftBtn
								If err.number<0 Then
										   Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to create Markup ["+sMarkupText+"]")
										   Fn_MSO_MarkupOperations=False
										   Exit function
								Else
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Successfully created Markup ["+sMarkupText+"]")
											Fn_MSO_MarkupOperations=true	
								End If
				
								If objWindow.Dialog("MarkupError").Exist Then
									objWindow.Dialog("MarkupError").WinButton("OK").Click 0,0,micLeftBtn
									   Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to create Markup ["+sMarkupText+"]")
									   Fn_MSO_MarkupOperations=False
									   Exit function
								End If
				
								If bSaveMarkup<>"" Then
									If lcase(bSaveMarkup)="on" Then
										objWpf.WpfButton("Save/Extract").Click 5,5,micLeftBtn
											If err.number<0 Then
													   Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to save the markup Markup ["+sMarkupText+"]")
													   Fn_MSO_MarkupOperations=False
													   Exit function
											Else
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Successfully saved Markup ["+sMarkupText+"]")
														Fn_MSO_MarkupOperations=true	
											End If
									End If
								End If
				
								If objWarning.Exist then
									objWarning.WinButton("OK").SetTOProperty "text",sInfo1
									wait 2
									objWarning.WinButton("OK").Click 5,5,micLeftBtn
											If err.number<0 Then
													   Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to handle the [Warning] Dialog")
													   Fn_MSO_MarkupOperations=False
													   Exit function
											Else
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Successfully handled the [Warning] Dialog")
														Fn_MSO_MarkupOperations=true	
											End If
								End if

'													Close Markup UI
													bReturn= Fn_KeyBoardOperation("SendKeys", "%YY08")
													If bReturn=true Then
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Successfully De-Activated the Markup UI in the Teamcenter Ribbon")
														Fn_MSO_MarkupOperations=true
													Else
													   Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to De-Activate the Markup UI in the Teamcenter Ribbon")
													   Fn_MSO_MarkupOperations=False
													   Exit function	
												   End If

							Case "DeleteMarkup"
				
												
								'Open the PopupMenu 
							
								objWpf.WpfList("listView").ShowContextMenu
								
								'Select the Option to create markup
								call Fn_KeyBoardOperation("SendKeys", "{TAB}")
								wait 1
								call Fn_KeyBoardOperation("SendKeys", "{TAB}")
								wait 1
								call Fn_KeyBoardOperation("SendKeys", "{TAB}")
								wait 1
								call Fn_KeyBoardOperation("SendKeys", "{Enter}")
								wait 1
					
					

									If objWindow.Dialog("MarkupError").Exist Then
									objWindow.Dialog("MarkupError").WinButton("OK").Click 0,0,micLeftBtn
									   Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to create Markup ["+sMarkupText+"]")
									   Fn_MSO_MarkupOperations=False
									   Exit function
								End If
				
								If bSaveMarkup<>"" Then
									If lcase(bSaveMarkup)="on" Then
										objWpf.WpfButton("Save/Extract").Click 5,5,micLeftBtn
											If err.number<0 Then
													   Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to save the markup Markup ["+sMarkupText+"]")
													   Fn_MSO_MarkupOperations=False
													   Exit function
											Else
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Successfully saved Markup ["+sMarkupText+"]")
														Fn_MSO_MarkupOperations=true	
											End If
									End If
								End If
				
								If objWarning.Exist then
									objWarning.WinButton("OK").SetTOProperty "text",sInfo1
									wait 2
									objWarning.WinButton("OK").Click 5,5,micLeftBtn
											If err.number<0 Then
													   Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to handle the [Warning] Dialog")
													   Fn_MSO_MarkupOperations=False
													   Exit function
											Else
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Successfully handled the [Warning] Dialog")
														Fn_MSO_MarkupOperations=true	
											End If
								End if

'													Close Markup UI
													bReturn= Fn_KeyBoardOperation("SendKeys", "%YY08")
													If bReturn=true Then
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Successfully De-Activated the Markup UI in the Teamcenter Ribbon")
														Fn_MSO_MarkupOperations=true
													Else
													   Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to De-Activate the Markup UI in the Teamcenter Ribbon")
													   Fn_MSO_MarkupOperations=False
													   Exit function	
												   End If

		End Select

		Set objWord=Nothing
		Set objWindow=Nothing
		set objWpf=Nothing
		Set objWarning=Nothing

End Function

''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''''/$$$$
''''/$$$$   FUNCTION NAME      :  Fn_MSO_Word_RequirementTreeOperations(sAction,sNode,sMenu,sInfo1,sInfo2,sInfo3)
''''/$$$$
''''/$$$$   DESCRIPTION        :  This function will Perform Desired Operations on the RM Tree in Ms-Word
''''/$$$$ 
''''/$$$$   PRE-REQUISITES     :  MS-Word Window Should be in Focus
''''/$$$$
''''/$$$$   PARAMETERS   	   : 		sAction :Valid Action Name
''''/$$$$										sNode : Valid Node Name
''''/$$$$										sMenu : Valid PopUpMenu
''''/$$$$										sInfo1: For Future Use
''''/$$$$										sInfo2: For Future Use
''''/$$$$										sInfo3 : For Future Use
''''/$$$$	
''''/$$$$	Return Value 	   : 				True or False
''''/$$$$
''''/$$$$   Function Calls     :   Fn_WriteLogFile(), Fn_MenuOperation(), Fn_ReadyStatusSync
''''/$$$$										
''''/$$$$
''''/$$$$	HISTORY            :   		AUTHOR                 DATE        VERSION
''''/$$$$
''''/$$$$   CREATED BY         :   SHREYAS          27/04/2012         1.0
''''/$$$$
''''/$$$$   REVIWED BY         :  Shreyas			27/04/2012           1.0
''''/$$$$
''''/$$$$	How To Use 		   :    bReturn=Fn_MSO_Word_RequirementTreeOperations("Expand","REQ-001115","","","","")
''''/$$$$				    		bReturn=Fn_MSO_Word_RequirementTreeOperations("PopUpMenuSelect","REQ-001360","New2","Requirement","Qwerty","")
''''/$$$$							
''''/$$$$	Modified by		   :  Vivek Ahirrao     [ TC112-20150715-31_07_2015-VivekA-Porting ]
''''/$$$$		Changes done   :  Modified function for operation on Req Tree elements.
''''/$$$$		How to use	   :  The node which
''''/$$$$						"Select","Verify" Case : First expand the Parent node under which Item is present (here REQ-003015).
''''/$$$$								Example : Fn_MSO_Word_RequirementTreeOperations("Select","000032:REQ-003015:REQ-003019","","","","")
''''/$$$$						"PopUpMenuSelect" Case : First expand the Parent node under which Item is present (here REQ-003015).
''''/$$$$								Example : Fn_MSO_Word_RequirementTreeOperations("PopUpMenuSelect","000032:REQ-003015:REQ-003019","Properties","","","")
''''/$$$$						"Expand", "Collapse" Case : First expand the Parent node under which Item is present (here REQ-003015).
''''/$$$$								Example : Fn_MSO_Word_RequirementTreeOperations("Expand","000032:REQ-003015:REQ-003019","","","","")
''''/$$$$	Modified by		   :  Vivek Ahirrao     [ TC1123-20160504-20_07_2016-VivekA-Maintenance ]
''''/$$$$	
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Public Function Fn_MSO_Word_RequirementTreeOperations(sAction,sNode,sMenu,sInfo1,sInfo2,sInfo3)
	GBL_FAILED_FUNCTION_NAME="Fn_MSO_Word_RequirementTreeOperations"
	Dim objAppType, objTree, objWPFWindow, objNode, obj, objMercury, wshshell, objMenu
	Dim aNode, xPos, yPos
	Dim nodetoBselected, bFlag, sProperties, sValues
	Dim itemcount, iCount, itemLocation, nodeCount

	Fn_MSO_Word_RequirementTreeOperations = False
	On error resume next

	Set objAppType = Window("MicrosoftWord")
	Set objTree = Window("MicrosoftWord").SwfObject("ControlAxSourcingSite").SwfObject("SwfObject").WpfWindow("WpfWindow").WpfObject("treeCtrl")
	Set objWPFWindow = Window("MicrosoftWord").SwfObject("ControlAxSourcingSite").SwfObject("SwfObject").WpfWindow("WpfWindow")
		
	If objWPFWindow.Exist(5) Then
		objAppType.Maximize
		wait 2
		
		Select Case sAction
			Case "Expand", "Collapse"
				aNode=Split(sNode,":")
				nodetoBselected = aNode(Ubound(aNode))
				itemcount= objTree.Object.Items.Count	
				For iCount = 0 To itemcount-1 Step 1
					Set objNode=objTree.Object.Items.GetItemAt(iCount)
					If objNode.component.DisplayName = nodetoBselected Then
						itemLocation=iCount
						Exit For
					End If
					Set objNode = nothing
				Next
				If sAction = "Collapse" Then
					If objTree.Object.Items.GetItemAt(itemLocation).IsExpanded <> False Then
						objTree.Object.Items.GetItemAt(itemLocation).IsExpanded = False
						If Err.Number < 0 Then
							Set objAppType = Nothing
							Set objTree = Nothing
							Set objWPFWindow = Nothing
							Exit Function
						End If
						Fn_MSO_Word_RequirementTreeOperations=True
						Wait 1
					Else
						Fn_MSO_Word_RequirementTreeOperations=True					
					End If
				Else	
					If objTree.Object.Items.GetItemAt(itemLocation).IsExpanded <> True Then
						objTree.Object.Items.GetItemAt(itemLocation).IsExpanded = True
						If Err.Number < 0 Then
							Set objAppType = Nothing
							Set objTree = Nothing
							Set objWPFWindow = Nothing
							Exit Function
						End If
						Fn_MSO_Word_RequirementTreeOperations=True
						Wait 2
					Else
						Fn_MSO_Word_RequirementTreeOperations=True					
					End If					
				End If
					
			Case "Select","Verify"
				aNode=Split(sNode,":")				
				For nodeCount = 0 To UBound(aNode)
					bFlag = False
					nodetoBselected = aNode(nodeCount)
					itemcount= objTree.Object.Items.Count	
					For iCount = 0 To itemcount-1 Step 1
						Set objNode=objTree.Object.Items.GetItemAt(iCount)
						If objNode.component.DisplayName = nodetoBselected Then
							If nodeCount = 0 Then
								If UBound(aNode) = 0 Then
									itemLocation=iCount								
									bFlag=True
									Exit For
								Else
									bFlag=True
									Exit For								
								End If
							Else
								If objNode.Parent.component.DisplayName = aNode(nodeCount-1) Then
									itemLocation=iCount
									bFlag=True
									Exit For
								End If
							End If
						End If
						Set objNode = nothing
					Next					
				Next
				
				If bFlag=True Then
					Set obj  = objTree.Object.Items.GetItemAt(itemLocation)
					Fn_MSO_Word_RequirementTreeOperations=objTree.Object.Items.GetItemAt(itemLocation).component.displayname
					Do while IsObject( obj.Parent)
						Set obj = obj.Parent
						Fn_MSO_Word_RequirementTreeOperations = obj.Component.DisplayName &":" &Fn_MSO_Word_RequirementTreeOperations
					Loop	
					If Trim(Fn_MSO_Word_RequirementTreeOperations) = Trim(sNode) Then
						If sAction = "Select" Then
							objTree.Object.SelectedIndex = itemLocation
						End If
						Fn_MSO_Word_RequirementTreeOperations = True
					Else
						Fn_MSO_Word_RequirementTreeOperations =False				
					End If								
				End If
			Case "PopUpMenuSelect"
				aNode=Split(sNode,":")				
				For nodeCount = 0 To UBound(aNode)
					bFlag = False
					nodetoBselected = aNode(nodeCount)
					itemcount= objTree.Object.Items.Count	
					For iCount = 0 To itemcount-1 Step 1
						Set objNode=objTree.Object.Items.GetItemAt(iCount)
						If objNode.component.DisplayName = nodetoBselected Then
							If nodeCount = 0 Then
								If UBound(aNode) = 0 Then
									itemLocation=iCount								
									bFlag=True
									Exit For
								Else
									bFlag=True
									Exit For								
								End If
							Else
								If objNode.Parent.component.DisplayName = aNode(nodeCount-1) Then
									itemLocation=iCount
									bFlag=True
									Exit For
								End If
							End If
						End If
						Set objNode = nothing
					Next					
				Next
			
				If bFlag=true Then
					objWPFWindow.WpfButton("SelectStartComponent").SetTOProperty "index",itemLocation
					Wait 1
					xPos = objWPFWindow.WpfButton("SelectStartComponent").GetROProperty("abs_x")
					yPos = objWPFWindow.WpfButton("SelectStartComponent").GetROProperty("abs_y")
					Set objMercury = CreateObject("Mercury.DeviceReplay")
					objMercury.MouseClick xPos+30,yPos+10,micLeftBtn
					Wait 0,200
					Set objMercury = Nothing
					Set wshshell = CreateObject("WScript.Shell")
					wshshell.SendKeys "+{F10}"
					Wait 2
					Set wshshell = Nothing
					sProperties ="Class Name~classname" 
					sValues = "ContextMenu~System.Windows.Controls.ContextMenu"
			
					Set objMenu = Fn_SISW_UI_Object_GetChildObjects("Fn_MSO_Word_RequirementTreeOperations",objWPFWindow, sProperties, sValues)
			
					For iCount = 0 To objMenu.count-1 
						If instr(objMenu(iCount).toString(), "WpfMenu") > 0 Then
							sMenu = Replace(sMenu,":",";")
							If sAction = "PopUpMenuSelect" Then
								objMenu(iCount).Select sMenu
							End If
							Exit For
						End If	
					Next
			 		If Err.Number<0 Then
				   		Fn_MSO_Word_RequirementTreeOperations=false
					Else
						If sMenu="New2" Then
							Set objNewSpec = WpfWindow("CreateSpecification")
							If objNewSpec.Exist(5) Then
							    'Select the Desired Specification Type
								objNewSpec.WpfComboBox("List").Select sInfo1
								wait 3
			
								'Set the Specification Name
								objNewSpec.WpfEdit("Name").Set sInfo2
								wait 3
	
								'Click on OK Button
					 			objNewSpec.WpfButton("OK").Click 5,5,micLeftBtn
					 			wait 3
					 		Else
					 			Fn_MSO_Word_RequirementTreeOperations = False
					 			Set objNewSpec = Nothing
					 			Set objAppType = Nothing
								Set objTree = Nothing
								Set objWPFWindow = Nothing
					 			Exit Function
					 		End If
					 		Set objNewSpec = Nothing
						End If
						Fn_MSO_Word_RequirementTreeOperations=true
					End If
				End If
			'TC 11.4 - Sandip C Maintenance Added case
			Case "PopUpMenuExist"
				aNode=Split(sNode,":")				
				For nodeCount = 0 To UBound(aNode)
					bFlag = False
					nodetoBselected = aNode(nodeCount)
					itemcount= objTree.Object.Items.Count	
					For iCount = 0 To itemcount-1 Step 1
						Set objNode=objTree.Object.Items.GetItemAt(iCount)
						If objNode.component.DisplayName = nodetoBselected Then
							If nodeCount = 0 Then
								If UBound(aNode) = 0 Then
									itemLocation=iCount								
									bFlag=True
									Exit For
								Else
									bFlag=True
									Exit For								
								End If
							Else
								If objNode.Parent.component.DisplayName = aNode(nodeCount-1) Then
									itemLocation=iCount
									bFlag=True
									Exit For
								End If
							End If
						End If
						Set objNode = nothing
					Next					
				Next
			
				If bFlag=true Then
					objWPFWindow.WpfButton("SelectStartComponent").SetTOProperty "index",itemLocation
					Wait 1
					xPos = objWPFWindow.WpfButton("SelectStartComponent").GetROProperty("abs_x")
					yPos = objWPFWindow.WpfButton("SelectStartComponent").GetROProperty("abs_y")
					Set objMercury = CreateObject("Mercury.DeviceReplay")
					objMercury.MouseClick xPos+30,yPos+10,micLeftBtn
					Wait 0,200
					
					Set wshshell = CreateObject("WScript.Shell")
					wshshell.SendKeys "+{F10}"
					Wait 2
					Set wshshell = Nothing
					sProperties ="Class Name~classname" 
					sValues = "ContextMenu~System.Windows.Controls.ContextMenu"
			
					Set objMenu = Fn_SISW_UI_Object_GetChildObjects("Fn_MSO_Word_RequirementTreeOperations",objWPFWindow, sProperties, sValues)
					If objMenu.count <> 0 Then
						Fn_MSO_Word_RequirementTreeOperations=True
						objMercury.MouseClick xPos,yPos-5,micLeftBtn
						Set objMercury = Nothing
					Else 
						Fn_MSO_Word_RequirementTreeOperations=false
					End If
				End If
				
			Case "GetTreeItem"
			
					itemcount= objTree.Object.Items.Count
					If iCount < 0 Then
						Fn_MSO_Word_RequirementTreeOperations = False
						Exit Function
					Else
						For iCount = 0 To itemcount-1 Step 1
							Set objNode=objTree.Object.Items.GetItemAt(iCount)
							If iCount = 0 Then
								bFlag = objNode.component.DisplayName
							else
								bFlag = bFlag&"~"&objNode.component.DisplayName
							End If
						Next
					End If					
					Fn_MSO_Word_RequirementTreeOperations = bFlag
					Exit Function
			
		End Select
	Else
		Fn_MSO_Word_RequirementTreeOperations = False
		Set objAppType = Nothing
		Set objTree = Nothing
		Set objWPFWindow = Nothing
		Exit Function
	End If			
	Set objAppType = Nothing
	Set objTree = Nothing
	Set objWPFWindow = Nothing			
End Function

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_SISW_MSO_FolderViewTreeOperations

'Description			 :	Function Used to Perform operation on Navigation Folder View Tree in MSWord

'Parameters			   :   '1.StrAction: Action Name
'										 2.StrNode: Node Name
'										 3.StrMenu: Popup menu
'										 4.StrMenuNumber: Popup menu number
'										 5.dicFolderViewInfo:extra parameter if required in future
'
'Return Value		   : 	True or False

'Examples				:   bReturn=Fn_SISW_MSO_FolderViewTreeOperations("Expand","Home:AutomatedTests:RMTrMTX_TLOccu_59633","","","")
'									 bReturn=Fn_SISW_MSO_FolderViewTreeOperations("Select","Home:AutomatedTests:RMTrMTX_TLOccu_59633","","","")
'									 bReturn=Fn_SISW_MSO_FolderViewTreeOperations("PopUpMenuSelect","Home:AutomatedTests:RMTrMTX_TLOccu_59633:000017-ReqSpec1","Open...","","")
'									 bReturn=Fn_SISW_MSO_FolderViewTreeOperations("PopUpMenuSelect","Home:AutomatedTests:RMTrMTX_TLOccu_59633","Create Item..","","")
'                       
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												16-May-2012								1.0																						Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_SISW_MSO_FolderViewTreeOperations(StrAction,StrNode,StrMenu,StrMenuNumber,dicFolderViewInfo)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_MSO_FolderViewTreeOperations"
	Fn_SISW_MSO_FolderViewTreeOperations=false
	'variable Declaration
	Dim objNodeObject,objChild
	Dim aNode,iCounter,bFlag,iCount,x,y,sLeft,Top,sRight,Bottom
	on error resume next
	'Maximizing [ MicrosoftWord ] window
	Window("MicrosoftWord").Maximize
	wait 2
	'Checking existance of [ FolderView ] 
	If not Window("MicrosoftWord").WinObject("FolderView").Exist(6) Then
		Call Fn_KeyBoardOperation("SendKeys","%~Y~C~1~{DOWN}~{ENTER}")
	End If
	Select Case StrAction
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to Expand and Collapse node
		Case "Expand","Collapse"
			Set objNodeObject=Description.Create
			objNodeObject("wpftypename").value="object"
			Set objChild=Window("MicrosoftWord").WinObject("FolderView").SwfObject("SwfObject").WpfWindow("WpfWindow").WpfObject("treeCtrl").ChildObjects(objNodeObject)
			aNode=Split(StrNode,":")
			For iCounter=0 to ubound(aNode)
				bFlag=false
				For iCount=0 to objChild.count-1
					If trim(objChild(iCount).GetVisibleText())=aNode(iCounter) Then
						bFlag=true
						Exit For
					End If
				Next
				If bFlag=false Then
					Exit For
				End If
			Next
			Set objNodeObject=nothing
			Set objChild=nothing
			If bFlag=true Then
				Window("MicrosoftWord").WinObject("FolderView").SwfObject("SwfObject").WpfWindow("WpfWindow").WpfButton("ExpandButton").SetTOProperty "index",iCount
				wait 1
				Window("MicrosoftWord").WinObject("FolderView").SwfObject("SwfObject").WpfWindow("WpfWindow").WpfButton("ExpandButton").Click 1,1,micLeftBtn
				If err.number<0 Then
				   Fn_SISW_MSO_FolderViewTreeOperations=false
				else
					Fn_SISW_MSO_FolderViewTreeOperations=true
				End if
			End If
			
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to Select node
		Case "Select"
			Set objNodeObject=Description.Create
			objNodeObject("wpftypename").value="object"
			Set objChild=Window("MicrosoftWord").WinObject("FolderView").SwfObject("SwfObject").WpfWindow("WpfWindow").WpfObject("treeCtrl").ChildObjects(objNodeObject)
			wait 2
			aNode=Split(StrNode,":")
			For iCounter=0 to ubound(aNode)
				bFlag=false
				For iCount=0 to objChild.count-1
					If trim(objChild(iCount).GetVisibleText())=aNode(iCounter) Then
						wait 2
						bFlag=true
						Exit For
					End If
				Next
				If bFlag=false Then
					Exit For
				End If
			Next
			Set objNodeObject=nothing
			Set objChild=nothing
			If bFlag=true Then
				If Window("MicrosoftWord").WinObject("FolderView").SwfObject("SwfObject").GetTextLocation (aNode(ubound(aNode)),sLeft,Top,sRight,Bottom)=True then
					x=(sLeft+sRight)/2
					y=(Top+Bottom)/2
					Window("MicrosoftWord").WinObject("FolderView").SwfObject("SwfObject").Click x,y
					If err.number<0 Then
					   Fn_SISW_MSO_FolderViewTreeOperations=false
					else
						Fn_SISW_MSO_FolderViewTreeOperations=true
					End if
				End If
			End if

		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to Select pop up menu/Context menu
		Case "PopUpMenuSelect"
			Set objNodeObject=Description.Create
			objNodeObject("wpftypename").value="object"
			Set objChild=Window("MicrosoftWord").WinObject("FolderView").SwfObject("SwfObject").WpfWindow("WpfWindow").WpfObject("treeCtrl").ChildObjects(objNodeObject)
			wait 2
			aNode=Split(StrNode,":")
			For iCounter=0 to ubound(aNode)
				bFlag=false
				For iCount=0 to objChild.count-1
					If trim(objChild(iCount).GetVisibleText())=aNode(iCounter) Then
						wait 2
						bFlag=true
						Exit For
					End If
				Next
				If bFlag=false Then
					Exit For
				End If
			Next
			Set objNodeObject=nothing
			Set objChild=nothing
			If bFlag=true Then
				If Window("MicrosoftWord").WinObject("FolderView").SwfObject("SwfObject").GetTextLocation(aNode(ubound(aNode)),sLeft,Top,sRight,Bottom)=True then
					x=(sLeft+sRight)/2
					y=(Top+Bottom)/2
					Window("MicrosoftWord").WinObject("FolderView").SwfObject("SwfObject").Click x,y,micRightBtn
					wait 2
					Select Case StrMenu
						Case "Open..."
							If StrMenuNumber<>"" Then
								iCount=CInt(StrMenuNumber)
							else
								iCount=9
							End If
							For iCounter= 1 to iCount
								bReturn= Fn_KeyBoardOperation("SendKeys", "{TAB}")
									If bReturn=true Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Successfully pressed the Down Arrow Key "+cstr(iCounter)+" times and selected pop up menu [ "+StrMenu+"]")
										Fn_SISW_MSO_FolderViewTreeOperations=true
									Else
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to click the Down arrow key")
										Fn_SISW_MSO_FolderViewTreeOperations=False
									End If
									bReturn= Fn_KeyBoardOperation("SendKeys", "{ENTER}")
							Next
						End Select
				End If
			End if
	End Select
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_MSO_ZipFileOperations

'Description			 :	Function Used to perform operations on zip files

'Parameters			   :   1.StrAction: Action name
'										2.StrZipFileLocation : Zip file location
'										3.StrExtractTo: Extract location
'
'Return Value		   : 	True or False

'Pre-requisite			:	

'Examples				:  	bReturn=Fn_MSO_ZipFileOperations("Unzip","C:\Documents and Settings\x_navgha\My Documents\Downloads\X76765_e5ae9bdd.zip","C:\Documents and Settings\x_navgha\My Documents\Downloads\X76765_e5ae9bdd")
'									in this case use parameter : StrZipFileLocation to give folder path
'									bReturn=Fn_MSO_ZipFileOperations("CreateFolder","C:\Documents and Settings\x_navgha\My Documents\Downloads\X76765_e5ae9bdd","")
'
'
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												17-Apr-2012								1.0																						Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_MSO_ZipFileOperations(StrAction,StrZipFileLocation,StrExtractTo)
	GBL_FAILED_FUNCTION_NAME="Fn_MSO_ZipFileOperations"
 	'variable declaration
	Dim fso,objShell,FilesInZip
	Fn_MSO_ZipFileOperations=true
	'creating object of FileSystem
	Set fso = CreateObject("Scripting.FileSystemObject")
	Select Case StrAction
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'case to extract zip file
		Case "Unzip"
			'checking existance of zip file
			If not fso.FileExists(StrZipFileLocation) Then
				Set fso =nothing
				Exit function
			End If
			If NOT fso.FolderExists(StrExtractTo) Then
				fso.CreateFolder(StrExtractTo)
			End If
			'Extract the contants of the zip file.
			set objShell = CreateObject("Shell.Application")
			set FilesInZip=objShell.NameSpace(StrZipFileLocation).items
			objShell.NameSpace(StrExtractTo).CopyHere(FilesInZip)
			Set objShell = Nothing
			Fn_MSO_ZipFileOperations=true
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'in this case use parameter : StrZipFileLocation to give folder path
		Case "CreateFolder"
			If fso.FolderExists(StrZipFileLocation) Then
				fso.DeleteFolder StrZipFileLocation,True
				wait 2
			End If
			fso.CreateFolder(StrZipFileLocation)
			wait 2
			If fso.FolderExists(StrZipFileLocation) Then
				Fn_MSO_ZipFileOperations=true
			else
				Fn_MSO_ZipFileOperations=false
			End If
	End Select
	Set fso = Nothing
End Function
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''''/$$$$
''''/$$$$   FUNCTION NAME   :   Fn_SISW_MSO_PropertyOperations(sAction,sProperty,sPropVal,bClose,sInfo1,sInfo2)
''''/$$$$
''''/$$$$   DESCRIPTION        :  This function will perform Various Operations on the Property Dialog
''''/$$$$ 
''''/$$$$   PRE-REQUISITES        :  MS-Word Window Should be in Focus
''''/$$$$
''''/$$$$  PARAMETERS  :	sAction :Error Message Contents
''''/$$$$					sProperty : Valid Property Name
''''/$$$$					sPropVal : Desired Property Value
''''/$$$$					bClose : Boolean parameter to Close the Properties Dialog
''''/$$$$					sInfo1: For Future Use
''''/$$$$					sInfo2: For Future Use
''''/$$$$	
''''/$$$$		Return Value : 				True or False
''''/$$$$
''''/$$$$    Function Calls       :   Fn_WriteLogFile(), Fn_MenuOperation(), Fn_ReadyStatusSync
''''/$$$$										
''''/$$$$
''''/$$$$		HISTORY           :   		AUTHOR                 DATE        VERSION
''''/$$$$
''''/$$$$    CREATED BY     :   SHREYAS         11/05/2012         1.0
''''/$$$$
''''/$$$$    REVIWED BY     :  Shreyas			 1/05/2012          1.0
''''/$$$$
''''/$$$$		How To Use :    bReturn= Fn_SISW_MSO_PropertyOperations("Verify","BOM View Revisions","REQ-001756/A-View","true","","")
''''/$$$$
''''/$$$$	
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Public Function Fn_SISW_MSO_PropertyOperations(sAction,sProperty,sPropVal,bClose,ByRef sInfo1,sInfo2)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_MSO_PropertyOperations"
	Dim objProperty,iCount,iList,sListVal,bFlag
	Fn_SISW_MSO_PropertyOperations=false
	bFlag=False
	Set objProperty=SwfWindow("PropertiesDisplay").SwfObject("SwfObject").WpfWindow("PropWindow")

	'Check if the Property Dialog Exists
	If SwfWindow("PropertiesDisplay").Exist(10) then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The property Dialog Exists")
		Fn_SISW_MSO_PropertyOperations=True
	Else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Cannot proceed as the property Dialog does not Exists")
		Fn_SISW_MSO_PropertyOperations=False
		Exit Function
	End if
	Wait 5

	'iList= SwfWindow("PropertiesDisplay").SwfObject("SwfObject").WpfWindow("PropWindow").WpfList("PropertyList").GetItemsCount
	iList= uBound( split(SwfWindow("PropertiesDisplay").SwfObject("SwfObject").WpfWindow("PropWindow").WpfList("PropertyList").GetROProperty("all items"), vblf )) + 1

	Select Case sAction
		Case "Verify"
			For iCount=0 to iList-1
				sValue=objProperty.WpfList("PropertyList").GetItem(iCount)
				If lCAse(sValue)=lCase(sProperty)Then
					objProperty.WpfEdit("PropertyBox").SetTOProperty "index",iCount
					wait 2
					If objProperty.WpfEdit("PropertyBox").Exist(5) Then
						sListVal=objProperty.WpfEdit("PropertyBox").GetROProperty("text")
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The Edit Box does not Exists")
						Fn_SISW_MSO_PropertyOperations=False
						Exit Function	
					End If
					sListVal=objProperty.WpfEdit("PropertyBox").GetROProperty("text")
					If instr(lcase(sListVal),lcase(sPropVal))>=0 Then
						bFlag=true
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The Property ["+sProperty+"] & its Specified Value ["+sPropVal+"] Matches")
						Exit For
					End If
				End If
			Next

		Case "Get"
			For iCount=0 to iList-1
				sValue=objProperty.WpfList("PropertyList").GetItem(iCount)
				If lCAse(sValue)=lCase(sProperty)Then
					objProperty.WpfEdit("PropertyBox").SetTOProperty "index",iCount
					wait 2
					If objProperty.WpfEdit("PropertyBox").Exist(5) Then
						sListVal=objProperty.WpfEdit("PropertyBox").GetROProperty("text")
						sInfo1=sListVal
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The Edit Box does not Exists")
						Fn_SISW_MSO_PropertyOperations=False
						Exit Function	
					End If
					bFlag=true
					Exit For
				End If
			Next
	End Select

	If cbool(bClose)=true Then
		SwfWindow("PropertiesDisplay").Close
	End If

	If bFlag=false then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The Function Failed")
		Fn_SISW_MSO_PropertyOperations=False
		If Fn_SISW_UI_Object_Operations("Fn_SISW_MSO_PropertyOperations","Exist", SwfWindow("PropertiesDisplay"), SISW_MIN_TIMEOUT) Then
			SwfWindow("PropertiesDisplay").Close
		End If
		Exit Function	
	Else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The Function Passed")
		Fn_SISW_MSO_PropertyOperations=True
	End If
	Set objProperty=nothing
End Function
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''''/$$$$
''''/$$$$   FUNCTION NAME   :   Fn_MSO_VerifyErrorAfterEdit(sMsg,sInfo1,sInfo2)
''''/$$$$
''''/$$$$   DESCRIPTION        :  This function will handle Error Msg after Performinmg Save Operation on Ms-Word
''''/$$$$ 
''''/$$$$   PRE-REQUISITES        :  MS-Word Window Should be in Focus
''''/$$$$
''''/$$$$  PARAMETERS   : 		sMsg :Error Message Contents
''''/$$$$										sInfo1: For Future Use
''''/$$$$										sInfo2: For Future Use
''''/$$$$	
''''/$$$$		Return Value : 				True or False
''''/$$$$
''''/$$$$    Function Calls       :   Fn_WriteLogFile(), Fn_MenuOperation(), Fn_ReadyStatusSync
''''/$$$$										
''''/$$$$
''''/$$$$		HISTORY           :   		AUTHOR                 DATE        VERSION
''''/$$$$
''''/$$$$    CREATED BY     :   SHREYAS         02/05/2012         1.0
''''/$$$$
''''/$$$$    REVIWED BY     :  Shreyas			 02/05/2012          1.0
''''/$$$$
''''/$$$$		How To Use :    bReturn=Fn_MSO_VerifyErrorAfterEdit("Error","","")
''''/$$$$							
''''/$$$$	
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

Public function Fn_MSO_VerifyErrorAfterEdit(sMsg,sInfo1,sInfo2)
	GBL_FAILED_FUNCTION_NAME="Fn_MSO_VerifyErrorAfterEdit"
	GBL_EXPECTED_MESSAGE=sMsg
   Dim sValue,i,bFlag,aValues
		sValue= WpfWindow("Errorsencounteredduring").WpfList("lvwErrors").GetVisibleText
		aValues=Split(sValue,vbnewline,-1,1)
		For i=0 to ubound(aValues)
			If instr(aValues(i),sMsg)>0 Then
				bFlag=true
				Exit For
			End If
		Next
	If bFlag=True Then
		WpfWindow("Errorsencounteredduring").WpfButton("OK").Click
		Fn_MSO_VerifyErrorAfterEdit=True
	Else
		GBL_ACTUAL_MESSAGE=sValue
		Fn_MSO_VerifyErrorAfterEdit=false
		Exit Function
	End If
End Function

''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''''/$$$$
''''/$$$$   FUNCTION NAME   :   Fn_SISW_MSO_RibbonbuttonClick(sAppName,sButton)
''''/$$$$
''''/$$$$   DESCRIPTION        :  This function will perform button click on ribbon button
''''/$$$$ 
''''/$$$$   PRE-REQUISITES        :  MSOffice application instance should be opened
''''/$$$$
''''/$$$$  PARAMETERS   : 		sAppName: MSApplication  Name i.e. Excel,Word
'''/$$$$										sButton: Ribbon button Name											
''''/$$$$	
''''/$$$$		Return Value : 				True or False
''''/$$$$
''''/$$$$    Function Calls       :   Fn_GetEnvValue()
''''/$$$$										
''''/$$$$
''''/$$$$		HISTORY           :   		AUTHOR                 DATE        VERSION
''''/$$$$
''''/$$$$    CREATED BY     :   Nilesh Gadekar         24/06/2013         1.0
''''/$$$$
''''/$$$$
''''/$$$$		How To Use :    bReturn= Fn_SISW_MSO_RibbonbuttonClick("Excel","Save")
''''/$$$$							
''''/$$$$	
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

Public Function Fn_SISW_MSO_RibbonbuttonClick(sAppName,sButton)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_MSO_RibbonbuttonClick"
   Dim objDll
	On Error Resume Next
	If Environment.Value("sPath")="" then
			Environment.Value("sPath") = Fn_GetEnvValue("User", "AutomationDir") 
	End If
	Set objDll=DotnetFactory.CreateInstance("MsRibbonMng.msRibbon.clickRibbon",Environment.Value("sPath")+"\Library\MsRibbonMng.dll")
	objDll.SetMsApplication sAppName
	Wait 2
	objDll.ClickRibbonButton sButton
	If Err.Number<0 Then
		Fn_SISW_MSO_RibbonbuttonClick=False
	Else
		Fn_SISW_MSO_RibbonbuttonClick=True
	End If
End Function
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''''/$$$$
''''/$$$$   FUNCTION NAME   :   Fn_RemoveHyperLink(sUrl,bFollowLink,sInfo1,sInfo2)
''''/$$$$
''''/$$$$   DESCRIPTION        :  This function will remove Hyperlinks within a word Document
''''/$$$$ 
''''/$$$$   PRE-REQUISITES        :  Microsoft Word Wimdow should be Opened with the Text to be made as Hyperlink
''''/$$$$
''''/$$$$  PARAMETERS   : 		sUrl : Valid URL to be made as Hyperlink
''''/$$$$											bFollowLink : Boolean Parameter to follow the Hyperlink
''''/$$$$										sInfo1: For Future Use
''''/$$$$										sInfo2 : For Future Use
''''/$$$$	
''''/$$$$		Return Value : 				True or False
''''/$$$$
''''/$$$$    Function Calls       :   Fn_WriteLogFile(), Fn_MenuOperation(), Fn_ReadyStatusSync
''''/$$$$										
''''/$$$$
''''/$$$$		HISTORY           :   		AUTHOR                 DATE        VERSION
''''/$$$$
''''/$$$$    CREATED BY     :   SHREYAS          20/12/2011         1.0
''''/$$$$
''''/$$$$    REVIWED BY     :  Shreyas			20/12/2011            1.0
''''/$$$$
''''/$$$$		How To Use :   bReturn=Fn_RemoveHyperLink("http://www.google.com","True","")
''''/$$$$							
''''/$$$$	
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Public Function Fn_RemoveHyperLink(sUrl,bFollowLink,sInfo1)
	GBL_FAILED_FUNCTION_NAME="Fn_RemoveHyperLink"
	Dim objWord,sValue,xParentCord,yParentCord,objHyperLink,sLeft, Top, sRight, Bottom
	Set objWord=GetObject(,"Word.application")
	objWord.ActiveWindow.DocumentMap = False

	Set objWord=Window("MicrosoftWord").WinObject("MicrosoftWordDocument")
	Set objHyperLink=Window("MicrosoftWord").Window("Insert Hyperlink")
	Fn_RemoveHyperLink = false

	'Select the Text to be Converted into Hyperlink
	sValue= objWord.GetTextLocation("://",sLeft, Top, sRight, Bottom)
	If sValue<>true Then 
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail: Fn_RemoveHyperLink failed due to Invalid arguments")
		Exit function
	else 
		xParentCord = (sLeft+sRight) / 2 
		yParentCord =(Top+Bottom) / 2 
		wait 3
		objWord.Click xParentCord,yParentCord, micLeftBtn
		If err.number<0 Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Select the Text to make hyperlink")
			Exit Function
		Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Selected the Text to make hyperlink")
			Fn_RemoveHyperLink=True	
		End If
	End If 
	wait 1
	objWord.Type micCtrlDwn + "a" + micCtrlUp
	wait 2
	'Invoke the  Create Hyperlink Dialog
	objWord.Type micCtrlDwn + "k" + micCtrlUp
	If err.number<0 Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Select the Text to make hyperlink")
		Exit Function
	Else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Selected the Text to make hyperlink")
	End If
	wait 3

	'Clear the Contents from the Url TextField
	objHyperLink.WinObject("URL").Type micAltDwn +  "r"  + micAltUp

	If err.number<0 Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Select the Text to make hyperlink")
		Exit Function
	Else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Selected the Text to make hyperlink")
		Fn_RemoveHyperLink =True	
	End If
	Set objWord=nothing
	Set objHyperLink=nothing
End Function

''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''''/$$$$
''''/$$$$   FUNCTION NAME   :   Fn_MSO_WordErrorDialogOperations(sTitle, sTextMessage, sBtnName)
''''/$$$$
''''/$$$$   DESCRIPTION        :  This function will verify the message in the error dialog in MSWORD & will perform button click in the dialog.
''''/$$$$ 
''''/$$$$   PRE-REQUISITES        :  Dialog should be present
''''/$$$$
''''/$$$$  PARAMETERS   : 		sTitle: Title of the dialog
'''/$$$$										sTextMessage: Message in the dialog 		
'''/$$$$										sButton: Name of button to be clicked										
''''/$$$$	
''''/$$$$		Return Value : 				True or False
''''/$$$$
''''/$$$$    Function Calls       :   Fn_UI_WinButton_Click()
''''/$$$$										
''''/$$$$
''''/$$$$		HISTORY           :   		AUTHOR                 DATE        VERSION
''''/$$$$
''''/$$$$    CREATED BY     :   Ashwini Kumar        11/09/2013         1.0
''''/$$$$
''''/$$$$
''''/$$$$		How To Use :    bReturn= Fn_MSO_WordErrorDialogOperations("Microsoft Office Word", "<filename> may contain features that are not compatible with Plain Text format. Do you want to save the document in this format?", "Yes")
''''/$$$$							
''''/$$$$	
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Public Function Fn_MSO_WordErrorDialogOperations(sTitle, sTextMessage, sBtnName)
	GBL_FAILED_FUNCTION_NAME="Fn_MSO_WordErrorDialogOperations"
	Dim strMsg
	Set objDialog=Dialog("Microsoft Word")
	Call objDialog.SetTOProperty("text",sTitle)
    If objDialog.Exist(5) Then
		If sTextMessage<>"" Then
			strMsg=objDialog.WinObject("WinObject").GetROProperty("text")
				If InStr(strMsg,sTextMessage)>0 Then
					Fn_MSO_WordErrorDialogOperations=True
					Call Fn_UI_WinButton_Click("Fn_MSO_WordErrorDialogOperations", objDialog, sBtnName,"","","")
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSO_WordErrorDialogOperations: Dialog verified successfully.")
				Else
					Fn_MSO_WordErrorDialogOperations=False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL:Fn_MSO_WordErrorDialogOperations: Message in the dialog doesn't match")
				End If
		Else
			Fn_MSO_WordErrorDialogOperations= Fn_UI_WinButton_Click("Fn_MSO_WordErrorDialogOperations", objDialog, sBtnName,"","","")			
		End If
	Else
		Fn_MSO_WordErrorDialogOperations=False
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL:Fn_MSO_WordErrorDialogOperations: Window does not exist")
			'ExitTest
	End If
	Set objDialog = nothing
End Function

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - Function to perform Excel operations in background - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''Function Name		:	Fn_MSO_HiddenExcelOperation
'
''Description		  	 :  	Function to perform Excel operations in background

''Parameters		   :	1. sAction : NAme of theoperation to perform
'							2. sFilePath: PAth of the excel file for which operation to be carried out
'							3. sSheetNm: Name of Excel sheet to work upon
'							4. sCellPosition: Position or location of the cell to work on
'							5. sString: Name of the new file or Value to be replaced in the excel if any. Additional parameter for requiered operation
								
''Return Value		   :  	Object \ Nothing
'
''Examples		     	:	Fn_MSO_HiddenExcelOperation()

'History:                
'								Developer Name							Date				Rev. No.		Reviewer		Changes Done	
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'									Archana D		 				09-Oct-2013					1.0				Vallari
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_MSO_HiddenExcelOperation(sAction, sFilePath, sSheetNm, sCellPosition, sString)
	GBL_FAILED_FUNCTION_NAME="Fn_MSO_HiddenExcelOperation"
Dim objExcel, objWorkbook, objWorksheet

	Fn_MSO_HiddenExcelOperation = False
	
	'**********************************************************************************
	'Close all open Excel sheet
	'**********************************************************************************
	bReturn = Fn_WindowsApplications("TerminateAll", "EXCEL.EXE")
	If bReturn = True Then
		Call Fn_UpdateLogFiles("[" + Cstr(time) + "] - ACTION - PASS | Successfully Closed all the Open Excel Sheets" , "")
	End If
	Wait(3)
	
	Set objExcel = CreateObject("Excel.Application")
	objExcel.Visible = False
	objExcel.AlertBeforeOverwriting = False
	objExcel.DisplayAlerts = False
    '--------------------------------------------------------------------------
	'Open the workBook  Here you will have to give the path of the Excel file.
	'--------------------------------------------------------------------------
	Set objWorkbook = objExcel.Workbooks.Open(sFilePath)
	Set objWorksheet = objWorkbook.Worksheets(sSheetNm)

	Select Case sAction

	Case "CopyCell"
				        'Application.CutCopyMode=False 
						Extern.EmptyClipboard()
						objWorksheet.range(sCellPosition).Copy

	Case "Paste"
						 'Paste the data	 
						objWorksheet.Activate
						objWorksheet.range(sCellPosition).Select
					    objWorksheet.Paste
						
	End Select
	
	If sAction <> "Paste" Then
		objWorkbook.close
    	ObjExcel.Quit
	End If

	Fn_MSO_HiddenExcelOperation = True
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - Function to perform Cancel Checkin and Checkout Operation - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''Function Name		:	Fn_MSO_CancelChckInChckOut
'
''Description		  	 :  	Function to perform Cancel Checkout Operation

''Parameters		   :	1. sAction : NAme of Action to perform
'							2. sButton: Which you need to pass after modification in Object(word,excel)
'			
'							3. sCheckbox: It should be either 'ON' or 'OFF'
'							4. sString: Name of the new file or Value to be replaced in the excel if any. Additional parameter for requiered operation
								
''Return Value		   :  	True/false
'
''Examples		     	:	Fn_MSO_CancelChckInChckOut("CancelCheckoutAfterModification","No","ON","")

'History:                
'								Developer Name							Date				Rev. No.		Reviewer		Changes Done						TcRelease
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'								Madhura Puranik	 					27-Jul-2015				  1.0			Ankit N			Migrated from TC1013			  TC11.2_2015071500
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Public Function Fn_MSO_CancelChckInChckOut(sAction,sButton,sCheckbox,sParameter)
	GBL_FAILED_FUNCTION_NAME="Fn_MSO_CancelChckInChckOut"
	'declare varibales
	Dim bResult, iCount
	Fn_MSO_CancelChckInChckOut=False
	Select Case sAction
	'Case for Cancel Checkout After Modification
			Case "CancelCheckoutAfterModification","CancelCheckout"
				If sAction="CancelCheckoutAfterModification" Then
					For iCount = 1 To 20 
						If Dialog("ConfirmationBox").Exist(2)=True Then
							Exit For
						End If	
						wait 3
					Next
		
					If Dialog("ConfirmationBox").Exist(2)=True Then 
						If sButton <> "" Then
							Call Fn_UI_WinButton_Click("",Dialog("ConfirmationBox"),sButton,"","","")
							Fn_MSO_CancelChckInChckOut=True
						End If
					ElseIf Window("MicrosoftWordWin").Dialog("Microsoft Office Word").Exist(2)=True Then
						If sButton <> "" Then
							Call Fn_UI_WinButton_Click("",Window("MicrosoftWordWin").Dialog("Microsoft Office Word"), sButton,"","","")	
							Fn_MSO_CancelChckInChckOut=True 
						End If
					Else
						Fn_MSO_CancelChckInChckOut=False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL:Fn_MSO_CancelChckInChckOut: ConfirmationBox Dailog does not exist")		
					End If 
				End If
			
				For iCount = 1 To 20 
					If WpfWindow("Close").Exist(2)=True Then
						Exit For
					End If	
					wait 3
				Next
					
				If WpfWindow("Close").Exist(2)=True Then
					If sCheckbox<>"" Then
						WpfWindow("Close").WpfCheckBox("CancelCheck-out").Set sCheckbox
						wait 3
						bResult = Fn_UI_WpfButtonClick("",WpfWindow("Close"),"OK")
						If bResult = True Then
							Fn_MSO_CancelChckInChckOut=True
						Else
							Fn_MSO_CancelChckInChckOut=False	
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: failed to Click on OK button of Cancel Checkout window")		
							Exit Function	
						End If
					End If	
				Else	
					Exit Function										
				End If
				
			Case Else
				Exit Function				
		End Select
End Function	

''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''''/$$$$
''''/$$$$  FUNCTION NAME   	:  Fn_MSO_WorkFlowProcess
''''/$$$$
''''/$$$$  DESCRIPTION      :  Function is used to Perform operations on Workflow dilaog
''''/$$$$ 
''''/$$$$  PRE-REQUISITES  	:  Office Client Window Should be in Focus
''''/$$$$
''''/$$$$  PARAMETERS   	:  sProcesName 		: 	Process name
''''/$$$$ 					   sProcessDesc 	: 	Process description
''''/$$$$ 					   sTmplFilter 		: 	Filter type
''''/$$$$ 					   sProcessTempl 	: 	Process template
''''/$$$$ 					   sButton 			: 	button name
''''/$$$$ 					   sInfo1 			: 	For future use
''''/$$$$ 					   sInfo2 			: 	For future use
''''/$$$$	
''''/$$$$	Return Value 	:  True or False
''''/$$$$										
''''/$$$$	How To Use 		:  bReturn = Fn_MSO_WorkFlowProcess("Test","TestingProcess","All","AutoDoDo","OK","","")
''''/$$$$	
''''/$$$$	HISTORY         :  
''''/$$$$					Developer Name	     Date		Version		Changes								Reviewer
''''/$$$$	Created by  :	 Vivek Ahirrao	 10/06/2016	  	  1.0		Created								[TC1123-20160504-10_06_2016-VivekA-NewDevelopment]
''''/$$$$  
''''/$$$$	Modified by : 
''''/$$$$	
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Public Function Fn_MSO_WorkFlowProcess(sProcesName,sProcessDesc,sTmplFilter,sProcessTempl,sButton,sInfo1,sInfo2)
	GBL_FAILED_FUNCTION_NAME="Fn_MSO_WorkFlowProcess"
	Dim wrkFlow
	Fn_MSO_WorkFlowProcess=False
	Set wrkFlow = WpfWindow("NewWorkflowProcess")

	If not wrkFlow.Exist(5) Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL:Fn_MSO_WordErrorDialogOperations: Window does not exist")
		Exit Function
	End If
	
	'Select process templete
	wrkFlow.WpfComboBox("processTemplate").Select sProcessTempl
	If Err.Number < 0 Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL:Unable to Select Process Templete")
		Exit Function
	End If
	
	'Enter Process name
	wrkFlow.WpfEdit("processName").SetTOProperty "devname","processName"
	wrkFlow.WpfEdit("processName").Set sProcesName
	If Err.Number < 0 Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL:Unable to Enter Process Name in New Worflow Process Window")
		Exit Function
	End If
	
	'Enter Description name
	If sProcessDesc <> "" Then
		wrkFlow.WpfEdit("processName").SetTOProperty "devname","processDesc"
		wait 1
		wrkFlow.WpfEdit("processName").Set sProcessDesc
		If Err.Number < 0 Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL:Unable to Enter Description in New Worflow Process Window")
			Exit Function
		End If		
	End If
	
	'Set Process Templete Filter
	If sTmplFilter <> "" Then
		wrkFlow.WpfRadioButton("Assigned").SetTOProperty "text",sTmplFilter
		wrkFlow.WpfRadioButton("Assigned").Set
		If Err.Number < 0 Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL:Unable to set Process Templete Filter")
			Exit Function
		End If
	End If

	'Button Click
	If sButton <> "" Then
		wrkFlow.WpfButton("OK").SetTOProperty "text",sButton
		wrkFlow.WpfButton("OK").Click	
		If Err.Number < 0 Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL:Failed to Click on Button")
			Exit Function
		End If		
	Else
		wrkFlow.WpfButton("OK").Click	
		If Err.Number < 0 Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL:Failed to Click on [OK] Button")
			Exit Function
		End If		
	End If	
	Set wrkFlow =nothing
	Fn_MSO_WorkFlowProcess=True

End Function

''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''''/$$$$
''''/$$$$  FUNCTION NAME   	:  Fn_MSO_CheckIn_CheckoutOperations
''''/$$$$
''''/$$$$  DESCRIPTION      :  Function is used to Perform operations check-in, check-out etc.
''''/$$$$ 
''''/$$$$  PRE-REQUISITES  	:  Office Client Window Should be in Focus
''''/$$$$
''''/$$$$  PARAMETERS   	:  sApplication 	: 	MSExcel, MSWord, MSPowerPoint
''''/$$$$ 					   sAction 			: 	Action name
''''/$$$$ 					   sCheckOutId 		: 	Id
''''/$$$$ 					   sCheckOutComment : 	comment
''''/$$$$ 					   sButton 			: 	button name
''''/$$$$ 					   sCheckoutUser 	: 	User name / Dictionary Object
''''/$$$$	
''''/$$$$	Return Value 	:  True or False
''''/$$$$										
''''/$$$$	How To Use 		:  bReturn = Fn_MSO_CheckIn_CheckoutOperations("MSExcel","CheckOut","","","OK","")
''''/$$$$					   bReturn = Fn_MSO_CheckIn_CheckoutOperations("MSWord","ConfirmCheckOut","","","No","")
''''/$$$$	
''''/$$$$					   Set dicCheckOutDetails = CreateObject("Scripting.Dictionary")
''''/$$$$					   	   dicCheckOutDetails("User") = "autotest3"
''''/$$$$					   	   dicCheckOutDetails("Activity") = "Check-Out"
''''/$$$$					   	   dicCheckOutDetails("DateTime") = "2016-06-30 13:23:29"
''''/$$$$					   	   dicCheckOutDetails("ChangeID") = "1234"
''''/$$$$					   	   dicCheckOutDetails("Comment") = "web"
''''/$$$$					   bReturn = Fn_MSO_CheckIn_CheckoutOperations("MSWord","VerifyCheckoutHistory","","","Close",dicCheckOutDetails)
''''/$$$$	HISTORY    :  
''''/$$$$				Developer Name	     Date		Version	Changes							Reviewer
''''/$$$$---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
''''/$$$$	Created by  :	 Vivek Ahirrao	 	10/06/2016	  1.0		Created							[TC1123-20160504-10_06_2016-VivekA-NewDevelopment]
''''/$$$$---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
''''/$$$$	Modified by : 	 Vivek Ahirrao	 	30/06/2016	  1.0		Added case "VerifyCheckoutHistory"
''''/$$$$				 Vivek Ahirrao	 	07/07/2016	  1.0		Added case "ConfirmCheckOut"
''''/$$$$	
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Public Function Fn_MSO_CheckIn_CheckoutOperations(sApplication,sAction,sCheckOutId,sCheckOutComment,sButton,sCheckoutUser)
	GBL_FAILED_FUNCTION_NAME="Fn_MSO_CheckIn_CheckoutOperations"
	Dim objWindow, objCheckDlg
	Dim bFlag, sAppText, iRowCount, sRowData, iCounter, sSubAction, sProperty
	Dim aAppText, dicCount, dicItems, dicKeys
	
	Fn_MSO_CheckIn_CheckoutOperations=False
	
	If sApplication = "MSWord" Then
		Set objWindow = Window("MicrosoftWord")
	ElseIf sApplication = "MSExcel" Then
		Set objWindow = Window("MicrosoftExcel")
	ElseIf sApplication = "MSPowerPoint" Then
		Set objWindow = Window("MicrosoftPowerPoint")
	End If

	Select Case sAction
		Case "CheckOut"
				Set objCheckDlg = WpfWindow("Check-Out")
				If objCheckDlg.Exist(5) Then
					'Enter Change ID
					If sCheckOutId <> "" Then
						objCheckDlg.WpfEdit("sChangeId").Set sCheckOutId
						wait 1
					End If
					'Enter Reason
					If sCheckOutComment <> "" Then
						objCheckDlg.WpfEdit("sReason").Set sCheckOutComment
						wait 1
					End If
					'Click on OK or Cancel button
					If sButton<>"" Then
						bFlag = Fn_MSO_WpfButton_Click("Fn_MSO_CheckIn_CheckoutOperations", "Click", objCheckDlg, sButton)
					Else
						bFlag = Fn_MSO_WpfButton_Click("Fn_MSO_CheckIn_CheckoutOperations", "Click", objCheckDlg, "OK")
					End If
					If bFlag = False Then
						Set objCheckDlg = Nothing
						Set objWindow = Nothing
						Exit Function
					End If
				End If
				Fn_MSO_CheckIn_CheckoutOperations=True
		Case "CheckIn"
				Set objCheckDlg = objWindow.Dialog("DialogInformation")
				objCheckDlg.SetTOProperty "text","Confirm Check-In"
				If objCheckDlg.Exist(5) Then
					'Click on Yes or No button
					If sButton<>"" Then
						bFlag = Fn_UI_WinButton_Click("Fn_MSO_CheckIn_CheckoutOperations",objCheckDlg,sButton,5,5,micLeftBtn)
					Else
						bFlag = Fn_UI_WinButton_Click("Fn_MSO_CheckIn_CheckoutOperations",objCheckDlg,"Yes",5,5,micLeftBtn)
					End If
					If bFlag = False Then
						Set objCheckDlg = Nothing
						Set objWindow = Nothing
						Exit Function
					End If
				End If
				Fn_MSO_CheckIn_CheckoutOperations=True
		Case "ConfirmCheckOut"
				Set objCheckDlg = objWindow.Dialog("DialogInformation")
				objCheckDlg.SetTOProperty "text","Confirm Check-Out"
				If objCheckDlg.Exist(5) Then
					'Click on Yes or No button
					If sButton<>"" Then
						bFlag = Fn_UI_WinButton_Click("Fn_MSO_CheckIn_CheckoutOperations",objCheckDlg,sButton,5,5,micLeftBtn)
					Else
						bFlag = Fn_UI_WinButton_Click("Fn_MSO_CheckIn_CheckoutOperations",objCheckDlg,"Yes",5,5,micLeftBtn)
					End If
					If bFlag = False Then
						Set objCheckDlg = Nothing
						Set objWindow = Nothing
						Exit Function
					End If
				Else
					Set objCheckDlg = Dialog("ConfirmationBox")
					If objCheckDlg.Exist(5) Then
					'Click on Yes or No button
						If sButton<>"" Then
							bFlag = Fn_UI_WinButton_Click("Fn_MSO_CheckIn_CheckoutOperations",objCheckDlg,sButton,5,5,micLeftBtn)
						Else
							bFlag = Fn_UI_WinButton_Click("Fn_MSO_CheckIn_CheckoutOperations",objCheckDlg,"Yes",5,5,micLeftBtn)
						End If
						If bFlag = False Then
							Set objCheckDlg = Nothing
							Set objWindow = Nothing
							Exit Function
						End If
					End If
				End If
				Fn_MSO_CheckIn_CheckoutOperations=True
		Case "CancelCheckOut"
				Set objCheckDlg = objWindow.Dialog("DialogInformation")
				objCheckDlg.SetTOProperty "text","Cancel Check-Out"
				If objCheckDlg.Exist(5) Then
					'Click on Yes or No button
					If sButton<>"" Then
						bFlag = Fn_UI_WinButton_Click("Fn_MSO_CheckIn_CheckoutOperations",objCheckDlg,sButton,5,5,micLeftBtn)
					Else
						bFlag = Fn_UI_WinButton_Click("Fn_MSO_CheckIn_CheckoutOperations",objCheckDlg,"Yes",5,5,micLeftBtn)
					End If
					If bFlag = False Then
						Set objCheckDlg = Nothing
						Set objWindow = Nothing
						Exit Function
					End If
				End If
				Fn_MSO_CheckIn_CheckoutOperations=True
		Case "TransferCheckOut"
				'Future Use
		'Case to verify check out history
		'Verify 1 Row only at a time
		Case "VerifyCheckoutHistory"
				Set objCheckDlg = WpfWindow("CheckOutHistory")
				If objCheckDlg.Exist(5) Then
					sAppText = objCheckDlg.WpfList("WpfList").GetVisibleText()
					aAppText = Split(sAppText,vblf)
					'Check from Row 1, as 0th row contains Column names
					For iRowCount = 1 To UBound(aAppText)
						'sCheckoutUser should be Dictionary Object
						If varType(sCheckoutUser)<>"9" Then
							Set objCheckDlg = Nothing
							Set objWindow = Nothing
							Fn_MSO_CheckIn_CheckoutOperations=False
							Exit Function
						End If
						dicCount = sCheckoutUser.Count
						dicItems = sCheckoutUser.Items
						dicKeys = sCheckoutUser.Keys
						sRowData = ""
						For iCounter = 0 To dicCount - 1
							sSubAction = dicKeys(iCounter)
							sProperty = dicItems(iCounter)
							If sRowData="" Then
								sRowData = sSubAction &":-:"& sProperty
							Else
								sRowData = sRowData & "~" & sSubAction &":-:"& sProperty
							End If
							bFlag = False
							Select Case sSubAction
								Case "ChangeID"
									If sProperty<>"" Then
										If Instr(Trim(aAppText(iRowCount)),sProperty)>0 Then
											bFlag = True
										End If
									End If
								Case "Comment"
									If sProperty<>"" Then
										If Instr(Trim(aAppText(iRowCount)),sProperty)>0 Then
											bFlag = True
										End If
									End If
								Case "User"
									If sProperty<>"" Then
										If Instr(Trim(aAppText(iRowCount)),sProperty)>0 Then
											bFlag = True
										End If
									End If
								Case "Activity"
									If sProperty<>"" Then
										If Instr(Trim(aAppText(iRowCount)),sProperty)>0 Then
											bFlag = True
										End If
									End If
								Case "DateTime"
									If sProperty<>"" Then
										If Instr(Trim(aAppText(iRowCount)),sProperty)>0 Then
											bFlag = True
										End If
									End If
							End Select
							
							If bFlag = False Then
								Exit For
							End If
						Next
						If bFlag = True Then
							Exit For
						End If
					Next
					If bFlag = False Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Verify Row data - ["+sRowData+"].")
						Fn_MSO_CheckIn_CheckoutOperations = False
						Set objCheckDlg = Nothing
						Set objWindow = Nothing
						Exit Function
					End If
				End If
				'Click on OK or Cancel button
				If sButton<>"" Then
					bFlag = Fn_MSO_WpfButton_Click("Fn_MSO_CheckIn_CheckoutOperations", "Click", objCheckDlg, sButton)
				Else
					bFlag = Fn_MSO_WpfButton_Click("Fn_MSO_CheckIn_CheckoutOperations", "Click", objCheckDlg, "Close")
				End If
				If bFlag = False Then
					Set objCheckDlg = Nothing
					Set objWindow = Nothing
					Exit Function
				End If
				Fn_MSO_CheckIn_CheckoutOperations=True
				
		Case "VerifyConfirmCheckInDialogTxt","VerifyConfirmCheckOutDialogTxt"
				Set objCheckDlg = objWindow.Dialog("DialogInformation")
				
				If sAction = "VerifyConfirmCheckOutDialogTxt" Then
					Set objCheckDlg = Dialog("ConfirmationBox")
				Else
					objCheckDlg.SetTOProperty "text","Confirm Check-In"
				End If
				
				If objCheckDlg.Exist(5) Then
					sAppText = objCheckDlg.Static("TextMessage").GetROProperty("text")
					If sAppText = sCheckOutComment Then
						Fn_MSO_CheckIn_CheckoutOperations = True
					Else
						Fn_MSO_CheckIn_CheckoutOperations = False
						Set objCheckDlg = Nothing
						Set objWindow = Nothing
						Exit Function
					End If
				End If
					
	End Select
	Set objCheckDlg = Nothing
	Set objWindow = Nothing
End Function

''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''''/$$$$
''''/$$$$  FUNCTION NAME   	:  Fn_MSO_WpfButton_Click
''''/$$$$
''''/$$$$  DESCRIPTION      :  Function is used to Click on wpf button
''''/$$$$ 
''''/$$$$  PRE-REQUISITES  	:  Office Client Window Should be in Focus
''''/$$$$
''''/$$$$  PARAMETERS   	:  sFunctionName 	: 	Function name
''''/$$$$ 					   sAction 			: 	Action name
''''/$$$$ 					   objWpfDialog 	: 	dialog 
''''/$$$$ 					   sWpfButton 		: 	Button name
''''/$$$$	
''''/$$$$	Return Value 	:  True or False
''''/$$$$										
''''/$$$$	How To Use 		:  bFlag1 = Fn_MSO_WpfButton_Click("","Click",objSession,sButton)
''''/$$$$	
''''/$$$$	HISTORY         :  
''''/$$$$					Developer Name	     Date		Version		Changes								Reviewer
''''/$$$$	Created by  :	 Vivek Ahirrao	 10/06/2016	  	  1.0		Created								[TC1123-20160504-10_06_2016-VivekA-NewDevelopment]
''''/$$$$  
''''/$$$$	Modified by : 
''''/$$$$	
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Public Function Fn_MSO_WpfButton_Click(sFunctionName, sAction, objWpfDialog, sWpfButton)
	GBL_FAILED_FUNCTION_NAME="Fn_MSO_WpfButton_Click"
	Dim objWpfButton,objDeviceReplay
	Fn_MSO_WpfButton_Click = False
	'Object Creation
	If sWpfButton <> "" Then
		Set objWpfButton = objWpfDialog.WpfButton(sWpfButton)
	Else
		Set objWpfButton = objWpfDialog
	ENd IF
	
	'Verify JavaButton object exists
	If not objWpfButton.Exist(5) Then
		Set objJavaButton = Nothing 
		Exit Function
	End If
	
	Select Case sAction
		Case "Click"
			objWpfButton.MakeVisible()
			Err.Clear
			objWpfButton.Click
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & "PASS : Sucessfully clicked on objWpfButton.")
			Fn_MSO_WpfButton_Click = True
'		Case "DeviceReplay.Click"
''			If sWpfButton <> "" Then
''				objWpfButton.Click 1, 1,micLeftBtn
''				wait SISW_MICRO_TIMEOUT
''			End If
'			Set objDeviceReplay = CreateObject("Mercury.DeviceReplay")
'			objDeviceReplay.MouseMove (sWpfButton.GetROProperty("abs_x") + 5), (objWpfButton.GetROProperty("abs_y") + 5)
'			objDeviceReplay.MouseClick  (sWpfButton.GetROProperty("abs_x") + 5), (objWpfButton.GetROProperty("abs_y") + 5), 0
'			Fn_MSO_WpfButton_Click = True
		Case Else
	End Select
	'Clear memory of JavaButton object.
'	Set objDeviceReplay = Nothing
	Set objWpfButton = Nothing 
End Function

''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''''/$$$$
''''/$$$$  FUNCTION NAME   	:  Fn_MSO_Navigate
''''/$$$$
''''/$$$$  DESCRIPTION      :  Function is used to open Folder or browse view
''''/$$$$ 
''''/$$$$  PRE-REQUISITES  	:  Office Client Window Should be in Focus
''''/$$$$
''''/$$$$  PARAMETERS   	:  sAction 		: 	Action name
''''/$$$$ 					   sAppType 	: 	MSExcel, MSWord, MSPowerPoint
''''/$$$$ 					   sUserName 	: 	User name 
''''/$$$$ 					   sPassword 	: 	Password
''''/$$$$ 					   sGroup 		: 	Group
''''/$$$$ 					   sRole 		: 	Role
''''/$$$$	
''''/$$$$	Return Value 	:  True or False
''''/$$$$										
''''/$$$$	How To Use 		:  bFlag1 = Fn_MSO_Navigate("FolderView", "MSExcel", "AutoTest1", "AutoTest1", "Engineering", "Designer")
''''/$$$$	
''''/$$$$	HISTORY         :  
''''/$$$$					Developer Name	     Date		Version		Changes								Reviewer
''''/$$$$	Created by  :	 Vivek Ahirrao	 10/06/2016	  	  1.0		Created								[TC1123-20160504-10_06_2016-VivekA-NewDevelopment]
''''/$$$$  
''''/$$$$	Modified by : 
''''/$$$$	
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Public Function Fn_MSO_Navigate(sAction, sAppType, sUserName, sPassword, sGroup, sRole)
	GBL_FAILED_FUNCTION_NAME="Fn_MSO_Navigate"
	Dim bFlag, sLoginCase, objAppWindow
	Fn_MSO_Navigate = False
	Select Case sAppType
		Case "MSExcel"
			Set objAppWindow = Window("MicrosoftExcel")
			sLoginCase = "WpfExcelLogin"
		Case "MSWord"
			Set objAppWindow = Window("MicrosoftWord")
			sLoginCase = "WpfWordLogin"
		Case "MSPowerPoint"
			Set objAppWindow = Window("MicrosoftPowerPoint")
			sLoginCase = "WpfPowerPointLogin"
	End Select
	
	Select Case sAction
		Case "FolderView"
				If not Fn_UI_ObjectExist("Fn_MSO_Navigate",objAppWindow.WinObject("FolderView")) Then
					'Call Fn_KeyBoardOperation("SendKeys","%~Y~Y~6~{DOWN}~{ENTER}")	
					bFlag = Fn_MSO_RibbonButton_Operations(sAppType,"Click","Navigate","")
					If bFlag = False Then
						Fn_MSO_Navigate = False
						Set objAppWindow = Nothing
						Exit Function
					End If
					wait 3
				End If
				
				If Fn_UI_ObjectExist("Fn_MSO_Navigate", WpfWindow("Teamcenter Login")) Then
					bFlag = Fn_MSO_TeamcenterLogin(sLoginCase,sUserName,sPassword,sGroup,sRole)
					If bFlag = False Then
						Fn_MSO_Navigate = False
						Set objAppWindow = Nothing
						Exit Function
					End If
				End If
				wait 3
				Fn_MSO_Navigate = True
		Case "BrowseView"
				If not Fn_UI_ObjectExist("Fn_MSO_Navigate",objAppWindow.WinObject("BrowseView")) Then	
					bFlag = Fn_MSO_RibbonButton_Operations(sAppType,"Click","NavigateDropDown:Browse","")
					If not Fn_UI_ObjectExist("Fn_MSO_Navigate",WpfWindow("Teamcenter Login")) Then
						bFlag = Fn_MSO_RibbonButton_Operations(sAppType,"Click","NavigateDropDown:Browse","")
					End If
					If bFlag = False Then
						Fn_MSO_Navigate = False
						Set objAppWindow = Nothing
						Exit Function
					End If
					wait 3
				End If
				
				If Fn_UI_ObjectExist("Fn_MSO_Navigate", WpfWindow("Teamcenter Login")) Then
					bFlag = Fn_MSO_TeamcenterLogin(sLoginCase,sUserName,sPassword,sGroup,sRole)
					If bFlag = False Then
						Fn_MSO_Navigate = False
						Set objAppWindow = Nothing
						Exit Function
					End If
				End If
				wait 3
				Fn_MSO_Navigate = True
	End Select
	Set objAppWindow = Nothing
End Function

''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''''/$$$$
''''/$$$$  FUNCTION NAME   	:  Fn_MSO_FolderViewTreeOperations
''''/$$$$
''''/$$$$  DESCRIPTION      :  Function is used to open Folder or browse view
''''/$$$$ 
''''/$$$$  PRE-REQUISITES  	:  Office Client Window Should be in Focus
''''/$$$$
''''/$$$$  PARAMETERS   	:  sApplication : 	MSExcel, MSWord, MSPowerPoint
''''/$$$$ 					   sAction 		: 	Action name
''''/$$$$ 					   StrNode 		: 	Node Path
''''/$$$$ 					   sMenu 		: 	Menu
''''/$$$$ 					   sInfo1 		: 	Future use
''''/$$$$ 					   sInfo2 		: 	Future use
''''/$$$$ 					   sInfo3 		: 	Future use
''''/$$$$	
''''/$$$$	Return Value 	:  True or False
''''/$$$$										
''''/$$$$	How To Use 		:  bReturn = Fn_MSO_FolderViewTreeOperations("MSExcel","Select","Home:AutomatedTests:TestFolder_31335:000560-TestItem56019:000560/A;1-TestItem56019","","","","")
''''/$$$$					   bReturn = Fn_MSO_FolderViewTreeOperations("MSExcel","PopupMenuExists","Home:Mailbox","Check-In/Out...:Check-In","","","")
''''/$$$$					   bReturn = Fn_MSO_FolderViewTreeOperations("MSExcel","PopupMenuEnabled","Home:Mailbox","Check-In/Out...:Check-In","","","")
''''/$$$$					   bReturn = Fn_MSO_FolderViewTreeOperations("MSPowerPoint","MultiSelect","Home:AutomatedTests~Home:Newstuff~Home:000304-Doc1","","","","")
''''/$$$$					   bReturn = Fn_MSO_FolderViewTreeOperations("MSPowerPoint","MultiSelectPopupMenuSelect","Home:AutomatedTests~Home:Newstuff~Home:000304-Doc1","New Workflow Process...","","","")
''''/$$$$	
''''/$$$$	HISTORY         :  
''''/$$$$					Developer Name	     Date		Version		Changes								Reviewer
''''/$$$$	Created by  :	 Vivek Ahirrao	 10/06/2016	  	  1.0		Created								[TC1123-20160504-10_06_2016-VivekA-NewDevelopment]
''''/$$$$  
''''/$$$$	Modified by : 	 Vivek Ahirrao	 28/06/2016		  1.0		Added new Cases "PopupMenuExists", "PopupMenuEnabled"
''''/$$$$					 Vivek Ahirrao	 18/07/2016		  1.0		Added new Cases "MultiSelect", "MultiSelectPopupMenuSelect"
''''/$$$$	
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Public Function Fn_MSO_FolderViewTreeOperations(sApplication,sAction,StrNode,sMenu,sInfo1,sInfo2,sInfo3)
	GBL_FAILED_FUNCTION_NAME="Fn_MSO_FolderViewTreeOperations"
	Dim objNodeObject, aNode, objChild, objSwf
	Dim iCounter,bFlag,xCord,yCord,height,width,iCount,iRowCounter
	Dim objTree,objNode,nodetoBselected,itemcount,itemLocation,obj,nodeCount
	Dim sProperties,sValues,objMenu, sSelectString, aStrNode, aSelectString
	Dim x,y,sLeft,Top,sRight,Bottom,sLocation
	Const VK_CONTROL = 29
	
	On Error Resume Next
	
	Fn_MSO_FolderViewTreeOperations=False
	
	If sApplication="MSExcel" Then
		Set objAppType = Window("MicrosoftExcel")
		Set objTree = Window("MicrosoftExcel").WinObject("FolderView").SwfObject("SwfObject").WpfWindow("WpfWindow").WpfObject("treeCtrl")
'		Set objStackPanel = Window("MicrosoftExcel").WinObject("FolderView").SwfObject("SwfObject").WpfWindow("WpfWindow").WpfObject("StackPanel")
		Set objWPFWindow = Window("MicrosoftExcel").WinObject("FolderView").SwfObject("SwfObject").WpfWindow("WpfWindow")
		'Maximizing [ MicrosoftExcel ] window
		objAppType.Maximize
	ElseIf sApplication="MSWord" Then	
		Set objAppType = Window("MicrosoftWord")
		Set objTree = Window("MicrosoftWord").WinObject("FolderView").SwfObject("SwfObject").WpfWindow("WpfWindow").WpfObject("treeCtrl")
'		Set objStackPanel = Window("MicrosoftWord").WinObject("FolderView").SwfObject("SwfObject").WpfWindow("WpfWindow").WpfObject("StackPanel")
		Set objWPFWindow = Window("MicrosoftWord").WinObject("FolderView").SwfObject("SwfObject").WpfWindow("WpfWindow")
		'Maximizing [ MicrosoftWord ] window
		 objAppType.Maximize		
	ElseIf sApplication="MSPowerPoint" Then
		Set objAppType = Window("MicrosoftPowerPoint")
		Set objTree = Window("MicrosoftPowerPoint").WinObject("FolderView").SwfObject("SwfObject").WpfWindow("WpfWindow").WpfObject("treeCtrl")
'		Set objStackPanel = Window("MicrosoftPowerPoint").WinObject("FolderView").SwfObject("SwfObject").WpfWindow("WpfWindow").WpfObject("StackPanel")
		Set objWPFWindow = Window("MicrosoftPowerPoint").WinObject("FolderView").SwfObject("SwfObject").WpfWindow("WpfWindow")
		'Maximizing [ MicrosoftPowerPoint ] window
		 objAppType.Maximize		
	Else
		Set objAppType = Window("MicrosoftExcel")
		Set objTree = Window("MicrosoftExcel").WinObject("FolderView").SwfObject("SwfObject").WpfWindow("WpfWindow").WpfObject("treeCtrl")
'		Set objStackPanel = Window("MicrosoftExcel").WinObject("FolderView").SwfObject("SwfObject").WpfWindow("WpfWindow").WpfObject("StackPanel")
		Set objWPFWindow = Window("MicrosoftExcel").WinObject("FolderView").SwfObject("SwfObject").WpfWindow("WpfWindow")
		'Maximizing [ MicrosoftExcel ] window
		objAppType.Maximize		
	End If
	wait 1
	
	If Fn_UI_ObjectExist("Fn_MSO_FolderViewTreeOperations",objAppType.WinObject("FolderView"))=False Then
		bFlag = Fn_MSO_RibbonButton_Operations(sApplication,"Click","Navigate","")
		If bFlag = False Then
			Set objAppType = Nothing
			Set objTree = Nothing
'			Set objStackPanel = Nothing
			Set objWPFWindow = Nothing
			Fn_MSO_FolderViewTreeOperations = False
			Exit Function
		End If
		Wait 5
	End If
	
	Select Case sAction
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "ExistPanel"
				If objTree.Exist = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : [ Navigation - Folder View ] panel does not Exist.")
					Set objAppType = Nothing
					Set objTree = Nothing
'					Set objStackPanel = Nothing
					Set objWPFWindow = Nothing
					Exit Function
				Else
					Fn_MSO_FolderViewTreeOperations = True
				End If
		'Case to Get Item location
		Case "GetItemLocation"
				aNode=Split(StrNode,":")
				iRowCounter = 0
				For nodeCount = 0 To UBound(aNode)
					bFlag = False
					nodetoBselected = aNode(nodeCount)
					itemcount = objTree.Object.Items.Count
					For iCount = iRowCounter To itemcount-1 Step 1
						Set objNode=objTree.Object.Items.GetItemAt(iCount)
						If objNode.component.DisplayName = nodetoBselected Then
							iRowCounter = iCount+1
							If nodeCount = 0 Then
								If UBound(aNode) = 0 Then
									itemLocation=iCount								
									bFlag=True
									Exit For
								Else
									bFlag=True
									Exit For								
								End If
							Else
								If objNode.Parent.component.DisplayName = aNode(nodeCount-1) Then
									itemLocation=iCount
									bFlag=True
									Exit For
								End If
							End If
						End If
						Set objNode = nothing
					Next					
				Next
				If bFlag=True Then
					Set obj  = objTree.Object.Items.GetItemAt(itemLocation)
					Fn_MSO_FolderViewTreeOperations = objTree.Object.Items.GetItemAt(itemLocation).component.displayname
					Do while IsObject(obj.Parent)
						Set obj = obj.Parent
						Fn_MSO_FolderViewTreeOperations = obj.Component.DisplayName &":" &Fn_MSO_FolderViewTreeOperations
					Loop
					If Trim(Fn_MSO_FolderViewTreeOperations) = Trim(StrNode) Then
						Fn_MSO_FolderViewTreeOperations = itemLocation
					Else
						Fn_MSO_FolderViewTreeOperations = False
					End If								
				End If
		'Case to Select node
		Case "Select"
				itemLocation = Fn_MSO_FolderViewTreeOperations(sApplication,"GetItemLocation",StrNode,"","","","")
				If CStr(itemLocation)<>CStr(False) Then
					objTree.Object.SelectedIndex = itemLocation
					Fn_MSO_FolderViewTreeOperations = True
				End If
		'Case to Verify whether Node Exist
		Case "Exists"
				itemLocation = Fn_MSO_FolderViewTreeOperations(sApplication,"GetItemLocation",StrNode,"","","","")
				If CStr(itemLocation)<>CStr(False) Then
					Fn_MSO_FolderViewTreeOperations = True
				End If
		'Case to Expand node
		Case "Expand"
				itemLocation = Fn_MSO_FolderViewTreeOperations(sApplication,"GetItemLocation",StrNode,"","","","")
				If CStr(itemLocation)<>CStr(False) Then
					If objTree.Object.Items.GetItemAt(itemLocation).IsExpanded <> True Then
						objTree.Object.Items.GetItemAt(itemLocation).IsExpanded = True
						If Err.Number < 0 Then
							Set objTree = Nothing							
							Exit Function
						End If
						Wait 2
					End If
					Fn_MSO_FolderViewTreeOperations = True
				End If
		'Case to Collapse node
		Case "Collapse"
				itemLocation = Fn_MSO_FolderViewTreeOperations(sApplication,"GetItemLocation",StrNode,"","","","")
				If CStr(itemLocation)<>CStr(False) Then
					If objTree.Object.Items.GetItemAt(itemLocation).IsExpanded <> False Then
						objTree.Object.Items.GetItemAt(itemLocation).IsExpanded = False
						If Err.Number < 0 Then
							Set objTree = Nothing
							Exit Function
						End If
						Wait 1
					End If
					Fn_MSO_FolderViewTreeOperations = True
				End If
		'Case for PopUpMenu Selection
		Case "PopupMenuSelect","PopupMenuEnabled","PopupMenuExists"
				itemLocation = Fn_MSO_FolderViewTreeOperations(sApplication,"GetItemLocation",StrNode,"","","","")
'				x=30
'				y=5
'				For sParentCount = 0 To UBound(aNode)-1
'					x=x+20
'				Next
				If CStr(itemLocation)<>CStr(False) Then
'					objStackPanel.SetTOProperty "index",itemLocation
'					objStackPanel.Click x,y,micRightBtn
'					wait 2
'					sProperties ="Class Name~classname" 
'					sValues = "ContextMenu~System.Windows.Controls.ContextMenu"

					objWPFWindow.WpfButton("SelectStartComponent").SetTOProperty "index",itemLocation
					Wait 1
					xPos = objWPFWindow.WpfButton("SelectStartComponent").GetROProperty("abs_x")
					yPos = objWPFWindow.WpfButton("SelectStartComponent").GetROProperty("abs_y")
					objWPFWindow.WpfButton("SelectStartComponent").Click 5,5,micLeftBtn
					Set objMercury = CreateObject("Mercury.DeviceReplay")
					objMercury.MouseClick xPos+30,yPos+10,micLeftBtn
					Wait 0,200
					Set objMercury = Nothing
					Set wshshell = CreateObject("WScript.Shell")
					wshshell.SendKeys "+{F10}"
					Wait 2
					Set wshshell = Nothing
					sProperties ="Class Name~classname" 
					sValues = "ContextMenu~System.Windows.Controls.ContextMenu"
			
					Set objMenu = Fn_SISW_UI_Object_GetChildObjects("Fn_MSO_FolderViewTreeOperations",objWPFWindow, sProperties, sValues)
			
					For iCount = 0 To objMenu.count-1 
						If instr(objMenu(iCount).toString(), "WpfMenu") > 0 Then
							sMenu = Replace(sMenu,":",";")
							If sAction = "PopupMenuSelect" Then
								objMenu(iCount).select sMenu
							ElseIf sAction = "PopupMenuEnabled" Then
								bFlag = False
								bFlag = objMenu(iCount).GetItemProperty(sMenu,"enabled")
								Wait 1
								Set wshshell = CreateObject("WScript.Shell")
								wshshell.SendKeys "{ESC}"
								Wait 1
								Set wshshell = Nothing
'								objStackPanel.Click x-1,y,micLeftBtn
								Wait 1
							ElseIf sAction = "PopupMenuExists" Then
								bFlag = False
								bFlag = objMenu(iCount).GetItemProperty(sMenu,"Exists")
								Wait 1
								Set wshshell = CreateObject("WScript.Shell")
								wshshell.SendKeys "{ESC}"
								Wait 1
								Set wshshell = Nothing
'								objStackPanel.Click x-1,y,micLeftBtn
								Wait 1
							End If
							Exit For
						End If	
					Next
			 		If sAction = "PopupMenuSelect" Then
			 			If Err.Number<0 Then
					   		Fn_MSO_FolderViewTreeOperations=false
						Else
							Fn_MSO_FolderViewTreeOperations=true
						End If
					ElseIf sAction = "PopupMenuEnabled" OR sAction = "PopupMenuExists" Then
						If bFlag = False Then
					   		Fn_MSO_FolderViewTreeOperations=false
						Else
							Fn_MSO_FolderViewTreeOperations=true
						End If
			 		End If
				End If
		'Case to Multi Select nodes and PopupMenuSelect on multiple nodes
		Case "MultiSelect","MultiSelectPopupMenuSelect"
				sSelectString=""
				aStrNode = Split(StrNode,"~")
				For iCounter = 0 To UBound(aStrNode)
					itemLocation = Fn_MSO_FolderViewTreeOperations(sApplication,"GetItemLocation",aStrNode(iCounter),"","","","")
					If CStr(itemLocation)<>CStr(False) Then
						'Store itemLocation indexes in string
						If sSelectString="" Then
							sSelectString = itemLocation
						Else
							sSelectString = sSelectString & "~" & itemLocation
						End If
					Else
						Fn_MSO_FolderViewTreeOperations = False
						Exit Function
					End If
				Next				
				aSelectString = Split(sSelectString,"~")
				If sSelectString<>"" AND UBound(aSelectString)=UBound(aStrNode) Then
					Set objDeviceReplay = CreateObject("Mercury.DeviceReplay")
					For iCount = 0 To UBound(aSelectString)						
						'to Get x, y co-ordinates of node to click
'						x=30
'						y=5
'						aNode=Split(aStrNode(iCount),":")
'						For sParentCount = 0 To UBound(aNode)-1
'							x=x+20
'						Next
						'to click on 1st node
						If iCount=0 Then
'							objStackPanel.SetTOProperty "index",aSelectString(iCount)
'							objStackPanel.Click x,y,micLeftBtn

							objWPFWindow.WpfButton("SelectStartComponent").SetTOProperty "index",aSelectString(iCount)
							Wait 0,100
							xPos = objWPFWindow.WpfButton("SelectStartComponent").GetROProperty("abs_x")
							yPos = objWPFWindow.WpfButton("SelectStartComponent").GetROProperty("abs_y")
							objWPFWindow.WpfButton("SelectStartComponent").Click 5,5,micLeftBtn
							
							objDeviceReplay.MouseClick xPos+30,yPos+10,micLeftBtn
							'Hold Down Control key after selecting 1st node
							objDeviceReplay.KeyDown VK_CONTROL
							Wait 0,100
						Else
							'to Select other nodes
'							objStackPanel.SetTOProperty "index",aSelectString(iCount)
'							objStackPanel.Click x,y,micLeftBtn
							objWPFWindow.WpfButton("SelectStartComponent").SetTOProperty "index",aSelectString(iCount)
							Wait 0,100
							xPos = objWPFWindow.WpfButton("SelectStartComponent").GetROProperty("abs_x")
							yPos = objWPFWindow.WpfButton("SelectStartComponent").GetROProperty("abs_y")
							objWPFWindow.WpfButton("SelectStartComponent").Click 5,5,micLeftBtn
							
							objDeviceReplay.MouseClick xPos+30,yPos+10,micLeftBtn
							Wait 0,100
						End If
					Next						
					
					'Release Control key after selecting all nodes
					objDeviceReplay.KeyUp VK_CONTROL
					Wait 1
					Set objDeviceReplay = Nothing
					
					If sAction = "MultiSelect" Then
						Fn_MSO_FolderViewTreeOperations = True
					ElseIf sAction = "MultiSelectPopupMenuSelect" Then
						'To Right click on multiple nodes
'						objStackPanel.SetTOProperty "index",aSelectString(iCount-1)
'						objStackPanel.Click x,y,micRightBtn
'						wait 2
'						sProperties ="Class Name~classname" 
'						sValues = "ContextMenu~System.Windows.Controls.ContextMenu"
				
'						objWPFWindow.WpfButton("SelectStartComponent").SetTOProperty "index",itemLocation
'						Wait 1
'						xPos = objWPFWindow.WpfButton("SelectStartComponent").GetROProperty("abs_x")
'						yPos = objWPFWindow.WpfButton("SelectStartComponent").GetROProperty("abs_y")
'						objWPFWindow.WpfButton("SelectStartComponent").Click 5,5,micLeftBtn
'						
'						objDeviceReplay.MouseClick xPos+30,yPos+10,micLeftBtn
'						Wait 0,200
											
						Set wshshell = CreateObject("WScript.Shell")
						wshshell.SendKeys "+{F10}"
						Wait 2
						Set wshshell = Nothing
						sProperties ="Class Name~classname" 
						sValues = "ContextMenu~System.Windows.Controls.ContextMenu"
				
						Set objMenu = Fn_SISW_UI_Object_GetChildObjects("Fn_MSO_FolderViewTreeOperations",objWPFWindow, sProperties, sValues)
				
						For iCount = 0 To objMenu.count-1 
							If Instr(objMenu(iCount).toString(), "WpfMenu") > 0 Then
								sMenu = Replace(sMenu,":",";")
								objMenu(iCount).Select sMenu
								Exit For
							End If	
						Next
				 		
			 			If Err.Number<0 Then
					   		Fn_MSO_FolderViewTreeOperations = False
						Else
							Fn_MSO_FolderViewTreeOperations = True
						End If
					End If
				Else
					Fn_MSO_FolderViewTreeOperations = False
				End If
	End Select	
	Set objSwf = nothing
	Set objChild = nothing
	Set objNodeObject = nothing
	Set objAppType = Nothing
	Set objTree = Nothing
'	Set objStackPanel = Nothing
	Set objWPFWindow = Nothing

End Function

''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''''/$$$$
''''/$$$$  FUNCTION NAME   	:  Fn_MSO_RibbonButton_Operations
''''/$$$$
''''/$$$$  DESCRIPTION      :  Function is used to Perform operation on Ribon buttons and Menus in Excel or Word application
''''/$$$$ 
''''/$$$$  PRE-REQUISITES  	:  MS-Word or MS-Excel Window Should be in Focus
''''/$$$$
''''/$$$$  PARAMETERS   	:  sAppType 	: 	MSExcel, MSWord, MSPowerPoint
''''/$$$$ 					   sAction 		: 	Action name
''''/$$$$ 					   sButtonMenu 	: 	New:Item or New:Folder or Import to Teamcenter or Current Settings:Login
''''/$$$$ 					   sReserve 	:	For future use
''''/$$$$	
''''/$$$$	Return Value 	:  True or False
''''/$$$$										
''''/$$$$	How To Use 		:  bReturn = Fn_MSO_RibbonButton_Operations("MSExcel","Click","New:Item","")
''''/$$$$					   bReturn = Fn_MSO_RibbonButton_Operations("MSWord","ClickWinToolbar","Undo","")
''''/$$$$					   bReturn = Fn_MSO_RibbonButton_Operations("MSWord","IsMenuExist","CheckInOut:Check-Out...","")
''''/$$$$					   bReturn = Fn_MSO_RibbonButton_Operations("MSWord","IsMenuEnabled","CheckInOut:Check-Out...","")
''''/$$$$	
''''/$$$$	HISTORY		:  
''''/$$$$					Developer Name		Date	 Version	 Changes								Reviewer
''''/$$$$	Created by  :  Vivek Ahirrao	10/06/2016	   1.0	 	 Created								[TC1123-20160504-10_06_2016-VivekA-NewDevelopment]
''''/$$$$------------------------------------------------------------------------------------------------------------------------------------------------------------------
''''/$$$$	Modified by :  Vivek Ahirrao	28/06/2016	   1.0		 Added case "ClickWinToolbar"			[TC1123-20160504-28_06_2016-VivekA-NewDevelopment]
''''/$$$$				   Vivek Ahirrao	30/06/2016	   1.0		 Added cases "IsMenuExist","IsMenuEnabled"
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Public Function Fn_MSO_RibbonButton_Operations(sAppType,sAction,sButtonMenu,sReserve)
	GBL_FAILED_FUNCTION_NAME="Fn_MSO_RibbonButton_Operations"
	Dim objRibbon
	Dim aButtonMenu, iCount, sMenu, sButton
	
	Fn_MSO_RibbonButton_Operations = False
	
	Select Case sAppType
		Case "MSExcel"
			If Window("MicrosoftExcel").WinObject("Ribbon").Exist(1) Then
				Set objRibbon = Window("MicrosoftExcel").WinObject("Ribbon")
			Else
				Fn_MSO_RibbonButton_Operations = False
				Exit Function
			End If
		Case "MSWord"
			If Window("MicrosoftWord").WinObject("Ribbon").Exist(1) Then
				Set objRibbon = Window("MicrosoftWord").WinObject("Ribbon")
			Else
				Fn_MSO_RibbonButton_Operations = False
				Exit Function
			End If
		Case "MSPowerPoint"
			If Window("MicrosoftPowerPoint").WinObject("Ribbon").Exist(1) Then
				Set objRibbon = Window("MicrosoftPowerPoint").WinObject("Ribbon")
			Else
				Fn_MSO_RibbonButton_Operations = False
				Exit Function
			End If
	End Select
	
	Select Case sAction
		'Click on Toolbar in Ribbon
		Case "ClickWinToolbar"
			sContent = objRibbon.WinToolbar("QuickAccessToolbar").GetContent()
			aContent = Split(Replace(sContent,vblf,"*"),"*")
			Select Case sButtonMenu
				Case "Undo"
					For iCount = 0 To UBound(aContent)
						If Instr(aContent(iCount),"Undo")>0 OR Instr(aContent(iCount),"VBA")>0 Then
							sButton = aContent(iCount)
							Exit For
						End If
					Next
				Case "Save"
					sButton = "Save"
			End Select

			objRibbon.WinToolbar("QuickAccessToolbar").Press sButton,micLeftBtn
			If Err.Number<0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to Click on ["+sButton+"] Button on Quick Access Toolbar Ribbon in ["+sAppType+"] Application.")
				Fn_MSO_RibbonButton_Operations=False
				Exit function
			End If
			Wait(1)
			Fn_MSO_RibbonButton_Operations = True
		Case "Click","IsMenuEnabled","IsMenuExist"
			aButtonMenu = Split(sButtonMenu,":")
			If UBound(aButtonMenu)>0 Then
				For iCount=1 To UBound(aButtonMenu)
					If iCount=1 Then
						sMenu = aButtonMenu(iCount)
					Else
						sMenu = sMenu +":"+ aButtonMenu(iCount)
					End If
				Next
			End If
			Select Case aButtonMenu(0)
				Case "New"
					sButton = "New"
				Case "Import to Teamcenter"
					sButton = "ImporttoTeamcenter"
					If sMenu<>"" Then
						sButton = "ImporttoTeamcenterDropDown"
					End If
				Case "Current Settings"
					sButton = "CurrentSettings"
				Case "Navigate"
					sButton = "Navigate"
				Case "NavigateDropDown"
					sButton = "NavigateDropDown"
				Case "Search"
					sButton = "Search"
				Case "Save As"
					sButton = "SaveAs"
				Case "My Worklist"
					sButton = "MyWorklist"
				Case "New Workflow Process"
					sButton = "NewWorkflowProcess"
				Case "Save"
					sButton = "Save"
				Case "Open"
					sButton = "Open"
				Case "CheckInOut"
					sButton = "CheckInOut"
				Case "Undo"
					sButton = "Undo"
			End Select
			'Click on Ribbon button
			wait 1
'			objRibbon.WinButton(sButton).Click 5,5,micLeftBtn
			bFlag = Fn_SISW_UI_WinButton_Operations("Fn_MSO_RibbonButton_Operations","Click",objRibbon,sButton,"","","")
			wait 2
'			If Err.Number<0 Then
			If bFlag = False Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to Click on ["+sButton+"] Button on Menu Ribbon in ["+sAppType+"] Application.")
				Fn_MSO_RibbonButton_Operations=False
				Exit function
			End If
		 	Wait(1)
		 	'Click on Menu
		 	If sMenu<>"" Then
'		 		If NOT objRibbon.WinMenu("WinMenu").Exist Then
'					objRibbon.WinButton(sButton).Click 5,5,micLeftBtn
'				End If
		 		If sAction="Click" Then
		 			objRibbon.WinMenu("WinMenu").Select sMenu
				 	If Err.Number<0 Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to Perform Menu operation ["+sMenu+"].")
						Fn_MSO_RibbonButton_Operations=False
						Exit Function
					End If
		 		ElseIf sAction="IsMenuEnabled" OR sAction="IsMenuExist" Then
		 			If sAction="IsMenuEnabled" Then
		 				bFlag = objRibbon.WinMenu("WinMenu").GetItemProperty(sMenu,"Enabled")
		 			ElseIf sAction="IsMenuExist" Then
		 				bFlag = objRibbon.WinMenu("WinMenu").GetItemProperty(sMenu,"Exists")
		 			End If
		 			objRibbon.WinButton(sButton).Click 5,5,micLeftBtn
		 			Wait 1
		 			If bFlag = False Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to Perform Action ["+sAction+"].")
						Fn_MSO_RibbonButton_Operations=False
						Exit Function
					End If
		 		End If
				Wait(1)
		 	End If
		 	Wait(1)
			Fn_MSO_RibbonButton_Operations = True
		Case Else
			Fn_MSO_RibbonButton_Operations = False
	End Select
	
	Set objRibbon = Nothing
End Function

''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''''/$$$$
''''/$$$$  FUNCTION NAME   	:  Fn_MSO_TeamcenterLogout
''''/$$$$
''''/$$$$  DESCRIPTION      :  Function is used to logout from Excel or Word application
''''/$$$$ 
''''/$$$$  PRE-REQUISITES  	:  Office Client Window Should be in Focus
''''/$$$$
''''/$$$$  PARAMETERS   	:  sAppType 	: 	MSExcel, MSWord, MSPowerPoint
''''/$$$$ 					   sReserve 	: 	For future use
''''/$$$$	
''''/$$$$	Return Value 	:  True or False
''''/$$$$										
''''/$$$$	How To Use 		:  bReturn = Fn_MSO_TeamcenterLogout("Excel","")
''''/$$$$	
''''/$$$$	HISTORY         :  
''''/$$$$					Developer Name	     Date		Version		Changes								Reviewer
''''/$$$$	Created by  :	 Vivek Ahirrao		10/06/2016	  1.0		Created								[TC1123-20160504-10_06_2016-VivekA-NewDevelopment]
''''/$$$$  
''''/$$$$	Modified by : 
''''/$$$$	
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Public Function Fn_MSO_TeamcenterLogout(sAppType,sReserve)
	GBL_FAILED_FUNCTION_NAME="Fn_MSO_TeamcenterLogout"
	Dim objAppWindow
	Dim bFlag, sFileType
	
	Fn_MSO_TeamcenterLogout = False
	
	Select Case sAppType
		Case "MSExcel"
			Set objAppWindow = Window("MicrosoftExcel")
			sFileType = "EXCEL.EXE"
		Case "MSWord"
			Set objAppWindow = Window("MicrosoftWord")
			sFileType = "WINWORD.EXE"
		Case "MSPowerPoint"
			Set objAppWindow = Window("MicrosoftPowerPoint")
			sFileType = "POWERPNT.EXE"
	End Select
	
	If objAppWindow.Exist(5) Then
		bFlag = Fn_MSO_RibbonButton_Operations(sAppType,"Click","Current Settings:Logout","")
		If bFlag = False Then
			Exit Function
		End If
		Wait 1
		If objAppWindow.Dialog("TeamcenterLogout").Exist(5) Then
			objAppWindow.Dialog("TeamcenterLogout").WinButton("Yes").Click 5,5,micLeftBtn
			Wait 2
		End If
		If WpfWindow("Teamcenter Login").Exist(3) Then
			'Clicking on Cancel button
			Call Fn_UI_WpfButtonClick("Fn_MSO_TeamcenterLogout", WpfWindow("Teamcenter Login"), "Cancel")
		End If
		Call Fn_WindowsApplications("TerminateAll",sFileType)
	End If
	Set objAppWindow = Nothing
	Fn_MSO_TeamcenterLogout = True
End Function

''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''''/$$$$
''''/$$$$  FUNCTION NAME   	:  Fn_MSO_SetFocusOnApplicationWindow
''''/$$$$
''''/$$$$  DESCRIPTION     	:  Function is used to set Focus on application
''''/$$$$ 
''''/$$$$  PRE-REQUISITES  	:  MS-Word or MS-Excel Window Should be opened
''''/$$$$
''''/$$$$  PARAMETERS   	:  sAppType 	: 	MSExcel, MSWord or MSPowerPoint
''''/$$$$ 					   sReserve 	: 	For future use
''''/$$$$	
''''/$$$$	Return Value 	:  True or False
''''/$$$$										
''''/$$$$	How To Use 		:  Call Fn_MSO_SetFocusOnApplicationWindow("MSExcel","")
''''/$$$$	
''''/$$$$	HISTORY         :  
''''/$$$$					Developer Name		Date	Version		Changes									Reviewer
''''/$$$$	Created by  :	Vivek Ahirrao	26/05/2016	  1.0		Created									[TC1122-20160504-25_05_2016-VivekA-NewDevelopment]
''''/$$$$  	
''''/$$$$	Modified by : 
''''/$$$$	
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Public Function Fn_MSO_SetFocusOnApplicationWindow(sAppType, sReserve)
	GBL_FAILED_FUNCTION_NAME="Fn_MSO_SetFocusOnApplicationWindow"
	Fn_MSO_SetFocusOnApplicationWindow = False
	
	Select Case sAppType
		Case "MSExcel"
			If Window("MicrosoftExcel").Exist = True Then
				Window("MicrosoftExcel").Maximize
				Wait 1
			End If
		Case "MSWord"
			If Window("MicrosoftWord").Exist = True Then
				Window("MicrosoftWord").Maximize
				Wait 1
			End If
		Case "MSPowerPoint"
			If Window("MicrosoftPowerPoint").Exist = True Then
				Window("MicrosoftPowerPoint").Maximize
				Wait 1
			End If
	End Select
	Fn_MSO_SetFocusOnApplicationWindow = True
End Function

''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''''/$$$$
''''/$$$$  FUNCTION NAME   	:  Fn_MSO_BasicTeamcenterPreferences_Ops
''''/$$$$
''''/$$$$  DESCRIPTION     	:  Function is used to perform operations on Basic Teamcenter Preferences dialog
''''/$$$$ 
''''/$$$$  PRE-REQUISITES  	:  Office Client (MS-Excel, MS-Word or MS-PowerPoint) Window Should be opened
''''/$$$$
''''/$$$$  PARAMETERS   	:  sAction 			: 	Action to be performed
''''/$$$$ 					   sAppType 		: 	MSExcel, MSWord or MSPowerPoint
''''/$$$$ 					   dicMSODetails 	: 	Dictionary object
''''/$$$$ 					   sButton 			: 	Button to be clicked
''''/$$$$ 					   sReserve 		: 	For future use
''''/$$$$	
''''/$$$$	Return Value 	:  True or False
''''/$$$$										
''''/$$$$	How To Use 		:  Set dicMSODetails = CreateObject("Scripting.Dictionary")
''''/$$$$						   dicMSODetails("TabName1") 				= "Miscellaneous"
''''/$$$$						   dicMSODetails("ChoicesCombo1") 			= "Change Session Culture Choices:en-US"
''''/$$$$						   dicMSODetails("ChoicesCombo2") 			= "Change Color Theme Choices:Blue"
''''/$$$$						   dicMSODetails("LinktoRichClient") 		= "ON"
''''/$$$$						   'dicMSODetails("AutomaticallyDelete") 	= 
''''/$$$$						   dicMSODetails("TabName2") 				= "Insert Data"
''''/$$$$						   dicMSODetails("InsertAsRadioBtns") 		= "Folders:Details:ON~Items:HyperLink:ON~Image Files:Embed:ON"
''''/$$$$						   dicMSODetails("ShowIconWithHyperlink") 	= "OFF"
''''/$$$$						   dicMSODetails("TabName3") 				= "Insert Details"
''''/$$$$						   dicMSODetails("InsertDetailsChkBox") 	= "object_name:ON~item_id:ON~object_type:ON"
''''/$$$$						   dicMSODetails("Button1") 				= "Apply"
''''/$$$$						   dicMSODetails("Button2") 				= "OK"
''''/$$$$						bReturn = Fn_MSO_BasicTeamcenterPreferences_Ops("Set","MSExcel",dicMSODetails,"OK","")
''''/$$$$	
''''/$$$$	HISTORY         :  
''''/$$$$					Developer Name		Date		Version		Changes							Reviewer
''''/$$$$	Created by  :	Vivek Ahirrao	 	10/06/2016	  1.0		Created							[TC1123-20160504-10_06_2016-VivekA-NewDevelopment]
''''/$$$$  
''''/$$$$	Modified by : 
''''/$$$$	
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Public Function Fn_MSO_BasicTeamcenterPreferences_Ops(sAction,sAppType,dicMSODetails,sButton,sReserve)
	GBL_FAILED_FUNCTION_NAME="Fn_MSO_BasicTeamcenterPreferences_Ops"
	On Error Resume Next
	
	Dim objBasicTCPref, dicCount, dicItems, dicKeys, aProperty, aFields
	Dim iCounter, sSubAction, sProperty, bFlag, iCount, bFlag1, sItemCount, iCount1, sListItem
	
	Fn_MSO_BasicTeamcenterPreferences_Ops = False
	
	If WpfWindow("BasicTeamcenterPreferences").Exist = False Then
		'Call for menu operation
		Call Fn_MSO_RibbonButton_Operations(sAppType,"Click","Current Settings:Basic Teamcenter Preferences","")
		Wait 2
		If WpfWindow("BasicTeamcenterPreferences").Exist Then
			Set objBasicTCPref = WpfWindow("BasicTeamcenterPreferences")
		Else
			Exit Function
		End If
	End If
	
	Select Case sAction
		Case "Set"
			dicCount = dicMSODetails.Count
			dicItems = dicMSODetails.Items
			dicKeys = dicMSODetails.Keys
			
			For iCounter = 0 To dicCount - 1
				If Instr(dicKeys(iCounter),"TabName")>0 Then
					sSubAction = "TabName"
				ElseIf Instr(dicKeys(iCounter),"ChoicesCombo")>0 Then
					sSubAction = "ChoicesCombo"
				ElseIf Instr(dicKeys(iCounter),"Button")>0 Then
					sSubAction = "Button"
				Else
					sSubAction = dicKeys(iCounter)
				End If
				sProperty = dicItems(iCounter)
				Select Case sSubAction
					Case "TabName"
						'Set Tab [Miscellaneous, Insert Data or Insert Details]
						If sProperty<>"" Then
							objBasicTCPref.WpfTabStrip("tabControl").Select sProperty
							If Err.Number < 0 Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Set Tab ["+sProperty+"].")
								Fn_MSO_BasicTeamcenterPreferences_Ops=False
								Set objBasicTCPref = Nothing
								Exit Function
							End If
							Wait 1
						End If
					Case "ChoicesCombo"
						'Select Value from Choices combo box. i.e. Session Culture or Color/Theme
						If sProperty<>"" Then
							aProperty = Split(sProperty,":")
							If aProperty(0) = "Change Session Culture Choices" Then
								objBasicTCPref.WpfComboBox("MiscSessionCultureCmb").Select aProperty(1)
							ElseIf aProperty(0) = "Change Color Theme Choices" Then
								objBasicTCPref.WpfComboBox("MiscColorThemeCmb").Select aProperty(1)
							End If
							If Err.Number < 0 Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Set ["+aProperty(1)+"] value in Combo Box ["+aProperty(0)+"].")
								Fn_MSO_BasicTeamcenterPreferences_Ops=False
								Set objBasicTCPref = Nothing
								Exit Function
							End If
							Wait 1
						End If
					Case "LinktoRichClient"
						'Check Advanced Settings button is expanded or not
						If objBasicTCPref.WpfButton("AdvancedSettings").Object.IsChecked = False Then
							objBasicTCPref.WpfButton("AdvancedSettings").Click 5,5,micLeftBtn
							Wait 2
						End If
						'Set the Link to Rich Client check box ON or OFF
						If sProperty<>"" Then
							objBasicTCPref.WpfObject("MiscPropText").SetTOProperty "devname","Link to Rich Client.*"
							Wait 1
							objBasicTCPref.WpfCheckBox("MiscChkBox").Set sProperty
							If Err.Number < 0 Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Set ["+sProperty+"] value of [Link to Rich Client] check box.")
								Fn_MSO_BasicTeamcenterPreferences_Ops=False
								Set objBasicTCPref = Nothing
								Exit Function
							End If
							Wait 1
						End If
					Case "AutomaticallyDelete"
						'Check Advanced Settings button is expanded or not
						If objBasicTCPref.WpfButton("AdvancedSettings").Object.IsChecked = False Then
							objBasicTCPref.WpfButton("AdvancedSettings").Click 5,5,micLeftBtn
							Wait 1
						End If
						'objBasicTCPref.WpfObject("MiscPropText").SetTOProperty "devname","Automatically delete all Log files.*"
						'Future Use
					Case "InsertAsRadioBtns"
						If sProperty<>"" Then
							'set Multiple Radio btns at a time
							aProperty = Split(sProperty,"~")
							For iCount = 0 To UBound(aProperty)
								aFields = Split(aProperty(iCount),":")
								'Set Row Name
								objBasicTCPref.WpfObject("InsertAsRowText").SetTOProperty "devname",aFields(0)
								'Set Column Name
								objBasicTCPref.WpfObject("InsertAsColumnText").SetTOProperty "devname",aFields(1)
								'Click on Radio button to set it ON
								objBasicTCPref.WpfRadioButton("InsertAsRadBtn").Set aFields(2)
								Wait 2
								If Err.Number < 0 Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Set ["+aFields(0)+":"+aFields(1)+"] Radio button ["+aFields(2)+"].")
									Fn_MSO_BasicTeamcenterPreferences_Ops=False
									Set objBasicTCPref = Nothing
									Exit Function
								End If
							Next
						End If
					Case "ShowIconWithHyperlink"
						If sProperty<>"" Then
							'Set Show icon with hyperlink checkbox ON or OFF
							objBasicTCPref.WpfCheckBox("ShowIconWithHyperlink").Set sProperty
							If Err.Number < 0 Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Set Check Box [Show icon with hyperlink] as ["+sProperty+"].")
								Fn_MSO_BasicTeamcenterPreferences_Ops=False
								Set objBasicTCPref = Nothing
								Exit Function
							End If
							Wait 1
						End If
					Case "InsertDetailsChkBox"
						If sProperty<>"" Then
							'Select multiple check boxes in Insert Details tab
							aProperty = Split(sProperty,"~")
							For iCount = 0 To UBound(aProperty)
								bFlag1 = False
								aFields = Split(aProperty(iCount),":")
								'Get count in List in Insert Details tab
								sItemCount = objBasicTCPref.WpfList("InsertDetailsList").GetItemsCount()
								For iCount1 = 0 To sItemCount-1
									sListItem = objBasicTCPref.WpfList("InsertDetailsList").GetItem(iCount1)
									If sListItem = aFields(0) Then
										objBasicTCPref.WpfCheckBox("InsertDetailsChkBox").SetTOProperty "Index",iCount1
										objBasicTCPref.WpfCheckBox("InsertDetailsChkBox").Set aFields(1)
										Wait 1
										bFlag1 = True
										Exit For
									End If
								Next
								If bFlag1 = False Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Set Check Box ["+aFields(0)+"] as ["+aFields(1)+"].")
									Fn_MSO_BasicTeamcenterPreferences_Ops=False
									Set objBasicTCPref = Nothing
									Exit Function
								End If
							Next
						End If
					Case "Button"
						If sProperty<>"" Then
							'Click on button provided
							bFlag1 = Fn_MSO_WpfButton_Click("Fn_MSO_BasicTeamcenterPreferences_Ops","Click",objBasicTCPref,sProperty)
							If bFlag1 = False Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Click on WPFButton ["+sProperty+"].")
								Fn_MSO_BasicTeamcenterPreferences_Ops=False
								Set objBasicTCPref = Nothing
								Exit Function
							End If
							Wait 1
						End If
				End Select
			Next
			Fn_MSO_BasicTeamcenterPreferences_Ops = True
			
		'To define sAction can not be empty
		Case Else
			Fn_MSO_BasicTeamcenterPreferences_Ops = False
			Set objBasicTCPref = Nothing
			Exit Function
	End Select
	Set objBasicTCPref = Nothing
End Function

''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''''/$$$$
''''/$$$$  FUNCTION NAME   	:  Fn_MSO_Session_Operations
''''/$$$$
''''/$$$$  DESCRIPTION     	:  Function is used to perform operations on Session dialog
''''/$$$$ 
''''/$$$$  PRE-REQUISITES   :  Office Client (MS-Excel, MS-Word or MS-PowerPoint) Window Should be opened
''''/$$$$
''''/$$$$  PARAMETERS   	:  sAction 			: 	Action to be performed
''''/$$$$ 					   sAppType 		: 	MSExcel, MSWord or MSPowerPoint
''''/$$$$ 					   dicMSODetails 	: 	Dictionary object
''''/$$$$ 					   sButton 			: 	Button to be clicked
''''/$$$$ 					   sReserve 		: 	For future use
''''/$$$$	
''''/$$$$	Return Value 	:  True or False
''''/$$$$										
''''/$$$$	How To Use 		:  Set dicMSODetails = CreateObject("Scripting.Dictionary")
''''/$$$$						   dicMSODetails("ComboList1") = "WorkContext:WC2"
''''/$$$$						   dicMSODetails("ComboList2") = "Group:Engineering"
''''/$$$$						   dicMSODetails("ComboList3") = "Role:Designer"
''''/$$$$						   'dicMSODetails("ComboList4") = "Project:"
''''/$$$$						bReturn = Fn_MSO_Session_Operations("Set","MSExcel",dicMSODetails,"OK","")
''''/$$$$	
''''/$$$$	HISTORY         :  
''''/$$$$					Developer Name		Date		Version		Changes								Reviewer
''''/$$$$	Created by  :	Vivek Ahirrao	 	10/06/2016	  1.0		Created								[TC1123-20160504-10_06_2016-VivekA-NewDevelopment]
''''/$$$$  
''''/$$$$	Modified by : 
''''/$$$$	
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Public Function Fn_MSO_Session_Operations(sAction,sAppType,dicMSODetails,sButton,sReserve)
	GBL_FAILED_FUNCTION_NAME="Fn_MSO_Session_Operations"
	
	On Error Resume Next
	Dim objSession, dicCount, dicItems, dicKeys, aProperty
	Dim iCounter, sSubAction, sProperty, sPropertyValue, sPropertyName
	Fn_MSO_Session_Operations = False
	
	If WpfWindow("Session").Exist = False Then
		'Call for menu operation
		Call Fn_MSO_RibbonButton_Operations(sAppType,"Click","Current Settings:Session","")
		Wait 2
		If WpfWindow("Session").Exist Then
			Set objSession = WpfWindow("Session")
		Else
			Exit Function
		End If
	End If
	
	Select Case sAction
		Case "Set"
			dicCount = dicMSODetails.Count
			dicItems = dicMSODetails.Items
			dicKeys = dicMSODetails.Keys
			
			For iCounter = 0 To dicCount - 1
				If Instr(dicKeys(iCounter),"ComboList")>0 Then
					sSubAction = "ComboList"
				Else
					sSubAction = dicKeys(iCounter)
				End If
				sProperty = dicItems(iCounter)
				Select Case sSubAction
					Case "ComboList"
						If sProperty<>"" Then
							'Select from List
							aProperty = Split(sProperty,":")
							sPropertyValue = aProperty(1)
							Select Case aProperty(0)
								Case "WorkContext"
									sPropertyName = "Work Context"
								Case "Group"
									sPropertyName = "Group"
								Case "Role"
									sPropertyName = "Role"
								Case "Project"
									sPropertyName = "Project"
							End Select
							
							objSession.WpfObject("WPFTextField").SetTOProperty "text",sPropertyName&":"
							objSession.WpfComboBox("WPFCmbList").Select sPropertyValue
							If Err.Number < 0 Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Select ["+sPropertyValue+"] from ["+sPropertyName+"] WPFCombo List.")
								Fn_MSO_Session_Operations=False
								Set objSession = Nothing
								Exit Function
							End If
						End If
				End Select
			Next
		Case "Verify"     'TC 11.5(2018030500) Sandip C. REG-Integration_MSOffice New Developement. Added Case for Current Group and Role Verification on Session Dialog.
			dicCount = dicMSODetails.Count
			dicItems = dicMSODetails.Items
			dicKeys = dicMSODetails.Keys
			
			For iCounter = 0 To dicCount - 1
				sSubAction = dicKeys(iCounter)
				sProperty = dicItems(iCounter)
				Select Case sSubAction
					Case "TextObject"
						If sProperty<>"" Then
							If not WpfWindow("Session").WpfObject("CurrentGroup/RolesText").GetROProperty("text") = sProperty Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Verify ["+sProperty+"] from Current Group/Role Text Object")
								Fn_MSO_Session_Operations=False
								Set objSession = Nothing
								Exit Function
							End If
						End If
				End Select
			Next
			
			If sButton<>"" Then
				'Click on button provided
				bFlag1 = Fn_MSO_WpfButton_Click("Fn_MSO_Session_Operations","Click",objSession,sButton)
				If bFlag1 = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Click on WPFButton ["+sButton+"].")
					Fn_MSO_Session_Operations=False
					Set objSession = Nothing
					Exit Function
				End If
				Wait 1
			End If
			Fn_MSO_Session_Operations = True
	End Select
	Set objSession = Nothing
End Function

''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''''/$$$$
''''/$$$$  FUNCTION NAME   	:  Fn_MSO_CalendarDialogOps
''''/$$$$
''''/$$$$  DESCRIPTION      :  Function is used to Perform operation on Calendar dialog present in Office client app
''''/$$$$ 
''''/$$$$  PRE-REQUISITES  	:  Calendar dialog should be opened, by clicking Date Button in Office cleint application
''''/$$$$
''''/$$$$  PARAMETERS   	:  sAction 		: 	Action name
''''/$$$$ 					   sAppType 	: 	MSExcel, MSWord or MSPowerPoint
''''/$$$$ 					   sDateTime 	: 	New:Item or New:Folder or Import to Teamcenter or Current Settings:Login
''''/$$$$ 					   sButton		:	OK or Cancel or Clear
''''/$$$$ 					   sReserve 	:	For future use
''''/$$$$	
''''/$$$$	Return Value 	:  True or False
''''/$$$$										
''''/$$$$	How To Use 		:  bFlag1 = Fn_MSO_CalendarDialogOps("SetDateAndTime","MSExcel","10-Jul-2016 15:40","OK","")
''''/$$$$	
''''/$$$$	HISTORY		:  
''''/$$$$				  Developer Name		Date	 Version	 Changes							Reviewer
''''/$$$$	Created by  :  Vivek Ahirrao	10/06/2016	   1.0	 	Created								[TC1123-20160504-10_06_2016-VivekA-NewDevelopment]
''''/$$$$  
''''/$$$$	Modified by : 
''''/$$$$	
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Public Function Fn_MSO_CalendarDialogOps(sAction,sAppType,sDateTime,sButton,sReserve)
	GBL_FAILED_FUNCTION_NAME="Fn_MSO_CalendarDialogOps"
	Dim objCalendar, WshShell
	Dim aDateTime, sDate, sTime, iWidth, iHeight, aTime, bFlag1
	
	Fn_MSO_CalendarDialogOps = False
	On Error Resume Next
	
	Set objCalendar = WpfWindow("Calendar")
	If objCalendar.Exist(5) = False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Calendar dialog does not exist.")
		Set objCalendar = Nothing
		Exit Function
	End If
	
	Select Case sAction
		Case "SetDateAndTime"
			If sDateTime<>"" Then
				aDateTime = Split(sDateTime," ")
				sDate = aDateTime(0)
				If UBound(aDateTime)>0 Then
					sTime = aDateTime(1)
				Else
					sTime = ""
				End If
			Else
				Set objCalendar = Nothing
				Exit Function
			End If
			'Set Date
			If sDate<>"" Then
				objCalendar.SwfObject("SwfObject").SwfCalendar("SwfCalendar").SetDate sDate
				If Err.Number < 0 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Set Date ["+sDate+"].")
					Set objCalendar = Nothing
					Exit Function
				End If
				Wait 1
			End If
			'Set Time
			If sTime<>"" Then			
				iWidth = CInt(objCalendar.SwfObject("SwfObject").SwfCalendar("SwfCalendar").GetROProperty("Width"))
				iHeight = CInt(objCalendar.SwfObject("SwfObject").SwfCalendar("SwfCalendar").GetROProperty("Height"))
				
				aTime = Split(sTime,":")

				'Click on Text field to set time
				objCalendar.SwfObject("SwfObject").SwfCalendar("SwfCalendar").Click (iWidth/2)+(iWidth/4),(iHeight/2)+(iHeight/4),micLeftBtn
				
				Set WshShell = CreateObject("WScript.Shell")
				
				'Set Hour
				WshShell.SendKeys "{LEFT}"
				Wait 0,50
				objCalendar.SwfObject("SwfObject").SwfCalendar("SwfCalendar").Type aTime(0)
				Wait 0,50
				'Set Minute
				WshShell.SendKeys "{RIGHT}"
				Wait 0,50
				objCalendar.SwfObject("SwfObject").SwfCalendar("SwfCalendar").Type aTime(1)
				
				Set WshShell = nothing
			End If
			If sButton<>"" Then
				bFlag1 = Fn_MSO_WpfButton_Click("Fn_MSO_CalendarDialogOps","Click",objCalendar,sButton)
				If bFlag1 = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Click on ["+sButton+"] button.")
					Set objCalendar = Nothing
					Exit Function
				End If
			End If
			Fn_MSO_CalendarDialogOps = True
		Case "VerifyDateAndTime"
			'Future Use
	End Select
	Set objCalendar = Nothing
End Function

''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''''/$$$$
''''/$$$$  FUNCTION NAME   	:  Fn_MSO_AdvancedSearchOps
''''/$$$$
''''/$$$$  DESCRIPTION     	:  Function is used to perform operations on Advanced Search dialog
''''/$$$$ 
''''/$$$$  PRE-REQUISITES  	:  Office Client (MS-Excel, MS-Word or MS-PowerPoint) Window Should be opened
''''/$$$$
''''/$$$$  PARAMETERS   	:  sAction 		: 	Action to be performed
''''/$$$$ 					   sAppType 		: 	MSExcel, MSWord or MSPowerPoint
''''/$$$$ 					   dicSrchDetails 	: 	Dictionary object
''''/$$$$ 					   sButton 		: 	Button to be clicked
''''/$$$$ 					   sReserve 		: 	For future use
''''/$$$$	
''''/$$$$	Return Value 	:  True or False
''''/$$$$										
''''/$$$$	How To Use 		:  Set dicSrchDetails = CreateObject("Scripting.Dictionary")
''''/$$$$						dicSrchDetails("TabName1") 		= "Queries"
''''/$$$$						dicSrchDetails("SelectQueryType1") 	= "All Sequences"
''''/$$$$						dicSrchDetails("EditBox1") 			= "Name:Test1"
''''/$$$$						dicSrchDetails("EditBox2") 			= "Item ID:000040"
''''/$$$$						dicSrchDetails("ComboBox1") 		= "Alias Type:Identifier Revision"
''''/$$$$						dicSrchDetails("ComboBox2") 		= "Type:Item Revision"
''''/$$$$						dicSrchDetails("SetDateTime1") 		= "Modified After~10-Jul-2016 15:40"
''''/$$$$						dicSrchDetails("SetDateTime2") 		= "Modified Before~10-Jul-2016 15:40"
''''/$$$$					bReturn = Fn_MSO_AdvancedSearchOps("Find","MSExcel",dicSrchDetails,"FindQueries","")
''''/$$$$	
''''/$$$$	HISTORY         :  
''''/$$$$					Developer Name		Date		Version		Changes								Reviewer
''''/$$$$	Created by  :	Vivek Ahirrao	 	10/06/2016	  1.0		Created								[TC1123-20160504-10_06_2016-VivekA-NewDevelopment]
''''/$$$$  
''''/$$$$	Modified by : 
''''/$$$$	
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Public Function Fn_MSO_AdvancedSearchOps(sAction,sAppType,dicSrchDetails,sButton,sReserve)
	GBL_FAILED_FUNCTION_NAME="Fn_MSO_AdvancedSearchOps"
	Dim objAdvancedSrch
	Dim dicCount, dicItems, dicKeys
	Dim bFlag, bFlag1, iCounter, sSubAction, sProperty
	Dim aProperty
	
	Fn_MSO_AdvancedSearchOps = False
	On Error Resume Next
	
	Set objAdvancedSrch = WpfWindow("AdvancedSearch")
	If objAdvancedSrch.Exist(5) = False Then
		bFlag = Fn_MSO_RibbonButton_Operations(sAppType,"Click","Search","")
		If bFlag = False Then
			Set objAdvancedSrch = Nothing
			Exit Function
		End If
	End If
	If objAdvancedSrch.Exist(5) = False Then
		Set objAdvancedSrch = Nothing
		Exit Function
	End If
	
	Select Case sAction
		Case "Find","FindWithClear","Save","SaveWithClear"
			If sAction = "FindWithClear" OR sAction = "SaveWithClear" Then
				'Click on Clear button
				bFlag1 = Fn_MSO_WpfButton_Click("Fn_MSO_AdvancedSearchOps","Click",objAdvancedSrch,"Clear")
				If bFlag1 = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Click on [Clear] button.")
					Set objAdvancedSrch = Nothing
					Exit Function
				End If
			End If
			
			dicCount = dicSrchDetails.Count
			dicItems = dicSrchDetails.Items
			dicKeys = dicSrchDetails.Keys
			
			For iCounter = 0 To dicCount - 1
				If Instr(dicKeys(iCounter),"TabName")>0 Then
					sSubAction = "TabName"
				ElseIf Instr(dicKeys(iCounter),"SelectQueryType")>0 Then
					sSubAction = "SelectQueryType"
				ElseIf Instr(dicKeys(iCounter),"EditBox")>0 Then
					sSubAction = "EditBox"
				ElseIf Instr(dicKeys(iCounter),"ComboBox")>0 Then
					sSubAction = "ComboBox"
				ElseIf Instr(dicKeys(iCounter),"SetDateTime")>0 Then
					sSubAction = "SetDateTime"
				Else
					sSubAction = dicKeys(iCounter)
				End If
				sProperty = dicItems(iCounter)
				bFlag = False
				Select Case sSubAction
					Case "TabName"
						If sProperty<>"" Then
							objAdvancedSrch.WpfTabStrip("TabAdvanceSearch").Select sProperty
							If Err.Number < 0 Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Select ["+sProperty+"] WPF Tab.")
								Set objAdvancedSrch = Nothing
								Exit Function
							End If
							Wait 1
							bFlag = True
						End If
					Case "SelectQueryType"
						If sProperty<>"" Then
							'First select 0th item
							objAdvancedSrch.WpfList("SelectQueryType").Select 0
							Wait 2
							Err.Clear
							objAdvancedSrch.WpfList("SelectQueryType").Type sProperty
							Wait 2
							Err.Clear
							objAdvancedSrch.WpfList("SelectQueryType").Select sProperty
							If Err.Number < 0 Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Select ["+sProperty+"] Select Query Type WPF List.")
								Set objAdvancedSrch = Nothing
								Exit Function
							End If
'							bFlag1 = False
'							sItemCount = objAdvancedSrch.WpfList("SelectQueryType").GetItemsCount()
'							For iCount = 0 To sItemCount-1
'								sAppItem = objAdvancedSrch.WpfList("SelectQueryType").getItem(iCount)
'								If sAppItem = sProperty Then
'									objAdvancedSrch.WpfList("SelectQueryType").Select iCount
'									bFlag1 = True
'									Exit For
'								End If
'							Next
'							If bFlag1 = False Then
'								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Select ["+sProperty+"] Select Query Type WPF List.")
'								Set objAdvancedSrch = Nothing
'								Exit Function
'							End If
							Wait 5
							bFlag = True
						End If
					'"Name:ItemName"
					Case "EditBox"
						If sProperty<>"" Then
							aProperty = Split(sProperty,":")
							'Set text property 
							objAdvancedSrch.WpfObject("PropertyName").SetTOProperty "text",aProperty(0)
							If objAdvancedSrch.WpfObject("PropertyName").Exist Then
								objAdvancedSrch.WpfEdit("PropertyEditBox").Set aProperty(1)
								If Err.Number < 0 Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Set ["+aProperty(1)+"] in ["+aProperty(0)+"] WPF Edit Box.")
									Set objAdvancedSrch = Nothing
									Exit Function
								End If
								Wait 1
								bFlag = True
							End If
						End If
					'"Type:Item Revision"
					Case "ComboBox"
						If sProperty<>"" Then
							aProperty = Split(sProperty,":")
							'Set text property 
							objAdvancedSrch.WpfObject("PropertyName").SetTOProperty "text",aProperty(0)
							Wait 1
							If objAdvancedSrch.WpfObject("PropertyName").Exist Then
								objAdvancedSrch.WpfComboBox("PropertyComboBox").Click 5,5,micLeftBtn
								Wait 1
								Err.Clear
								objAdvancedSrch.WpfComboBox("PropertyComboBox").Type aProperty(1)
								Wait 1
								Err.Clear
								objAdvancedSrch.WpfComboBox("PropertyComboBox").Select aProperty(1)
								If Err.Number < 0 Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Select ["+aProperty(1)+"] from ["+aProperty(0)+"] WPF Combo Box.")
									Set objAdvancedSrch = Nothing
									Exit Function
								End If
								Wait 1
								bFlag = True
							End If
						End If
					'"Modified After~10-Jul-2016 15:40"
					'"Modified After~10-Jul-2016"
					Case "SetDateTime"
						If sProperty<>"" Then
							aProperty = Split(sProperty,"~")
							'Set text property 
							objAdvancedSrch.WpfObject("PropertyName").SetTOProperty "text",aProperty(0)
							If objAdvancedSrch.WpfObject("PropertyName").Exist Then
								'Click on Date button
								bFlag1 = Fn_MSO_WpfButton_Click("Fn_MSO_AdvancedSearchOps","Click",objAdvancedSrch,"DateButton")
								If bFlag1 = False Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Click on [DateButton] button.")
									Set objAdvancedSrch = Nothing
									Exit Function
								End If
								Wait 2
								If WpfWindow("Calendar").Exist(2) Then
									'Set Date & Time using below function
									bFlag1 = Fn_MSO_CalendarDialogOps("SetDateAndTime",sAppType,aProperty(1),"OK","")
									If bFlag1 = False Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Set Date and Time ["+aProperty(1)+"].")
										Set objAdvancedSrch = Nothing
										Exit Function
									End If
									bFlag = True
								End If
							End If
						End If
				End Select
				
				If bFlag = False Then
					Fn_MSO_AdvancedSearchOps = False
					Set objAdvancedSrch = Nothing
					Exit Function
				End If
			Next
			
			'Click on button provided
			If sButton<>"" Then
				bFlag1 = Fn_MSO_WpfButton_Click("Fn_MSO_AdvancedSearchOps","Click",objAdvancedSrch,sButton)
				If bFlag1 = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Click on WPFButton ["+sButton+"].")
					Fn_MSO_AdvancedSearchOps = False
					Set objAdvancedSrch = Nothing
					Exit Function
				End If
				Wait 1
			End If
			
			Fn_MSO_AdvancedSearchOps = True
		Case "Verify"
			'Future Use
	End Select
	
	Set objAdvancedSrch = Nothing
End Function

''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''''/$$$$
''''/$$$$  FUNCTION NAME   	:  Fn_MSO_GetCursorType
''''/$$$$
''''/$$$$  DESCRIPTION      :  Function is used to get Cursor Handle type
''''/$$$$ 
''''/$$$$	Return Value 	:  Cursor Type ID
''''/$$$$										
''''/$$$$	How To Use 		:  bReturn = Fn_MSO_GetCursorType()
''''/$$$$	
''''/$$$$	HISTORY         :  
''''/$$$$					Developer Name		Date	 Version	Changes									Reviewer
''''/$$$$	Created by  :	 Vivek Ahirrao	10/06/2016	  1.0		Created									[TC1123-20160504-10_06_2016-VivekA-NewDevelopment]
''''/$$$$  															Added for RM - Office Client new TC's Development
''''/$$$$	Modified by : 
''''/$$$$	
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Public Function Fn_MSO_GetCursorType()
	GBL_FAILED_FUNCTION_NAME="Fn_MSO_GetCursorType"
	Dim hwnd,pid,thread_id
 
	extern.Declare micLong,"GetForegroundWindow","user32.dll","GetForegroundWindow"
	extern.Declare micLong,"AttachThreadInput","user32.dll","AttachThreadInput", micLong, micLong,micLong
	extern.Declare micLong,"GetWindowThreadProcessId","user32.dll","GetWindowThreadProcessId", micLong, micLong
	extern.Declare micLong,"GetCurrentThreadId","kernel32.dll","GetCurrentThreadId"
	extern.Declare micLong,"GetCursor","user32.dll","GetCursor"
 
	hwnd = extern.GetForegroundWindow()
 
 	pid = extern.GetWindowThreadProcessId(hWnd, NULL)
	thread_id = extern.GetCurrentThreadId()
	extern.AttachThreadInput pid,thread_id,True
 
	Fn_MSO_GetCursorType = Eval("extern.GetCursor")
 
	extern.AttachThreadInput pid,thread_id,False
End function

''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''''/$$$$
''''/$$$$  FUNCTION NAME  	:  Fn_MSO_MainTabOperations
''''/$$$$
''''/$$$$  DESCRIPTION      :  Function is used to set Main Tab such as Teamcenter
''''/$$$$ 
''''/$$$$	Return Value 	:  True/False
''''/$$$$	
''''/$$$$  PARAMETERS   	:  sAction 		: 	Action to be performed
''''/$$$$ 					   sAppType 		: 	MSExcel, MSWord or MSPowerPoint
''''/$$$$ 					   sTabName 		: 	Tab Name
''''/$$$$ 					   sReserve 		: 	For future use
''''/$$$$ 
''''/$$$$	How To Use 		:  bReturn = Fn_MSO_MainTabOperations("Activate","MSExcel","Teamcenter","")
''''/$$$$					   bReturn = Fn_MSO_MainTabOperations("Select","MSExcel","FileClose","")
''''/$$$$	
''''/$$$$	HISTORY         :  
''''/$$$$				Developer Name		Date	 	Version	Changes							Reviewer
''''/$$$$	Created by  :	 Vivek Ahirrao		10/06/2016	  1.0		Created							[TC1123-20160504-10_06_2016-VivekA-NewDevelopment]
''''/$$$$  												Added for RM - Office Client new TC's Development
''''/$$$$	Modified by :   Vivek Ahirrao		07/07/2016	  1.0		Added case "Select",Subcase "FileClose"
''''/$$$$	
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Public Function Fn_MSO_MainTabOperations(sAction,sAppType,sTabName,sReserve)
	GBL_FAILED_FUNCTION_NAME="Fn_MSO_MainTabOperations"
	Dim wshshell
	
	On Error Resume Next
	Fn_MSO_MainTabOperations = False
	
	Select Case sAppType
		Case "MSExcel"
			Set objAppWindow = Window("MicrosoftExcel")
		Case "MSWord"
			Set objAppWindow = Window("MicrosoftWord")
		Case "MSPowerPoint"
			Set objAppWindow = Window("MicrosoftPowerPoint")
	End Select
	
	If objAppWindow.Exist = False Then
		Set objAppWindow = Nothing
		Exit Function
	End If
	'Set focus on App window
	objAppWindow.Maximize
	Wait 1
	
	Select Case sAction
		Case "Activate"
			Select Case sTabName
				Case "Teamcenter"
					objAppWindow.WinObject("Ribbon").WinTab("RibbonTabs").Select("Teamcenter")
			End Select
			Fn_MSO_MainTabOperations = True
		Case "Exists"
		Case "Select"
			Select Case sTabName
				Case "FileClose"
					Set wshshell = CreateObject("WScript.Shell")
					wshshell.SendKeys "%"
					wshshell.SendKeys "F"
					Wait 0,500
					wshshell.SendKeys "C"
					Set wshshell = Nothing
					wait 2
			End Select
			Fn_MSO_MainTabOperations = True
		Case "ClickTab"
			If sTabName<>"" Then
				Set objTab = objAppWindow.WinObject("Ribbon").WinTab("RibbonTabs")
				objTab.Select sTabName,micLeftBtn
				If Err.Number<>0 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Select ["+sTabName+"] Ribbon Main tab.")
					Fn_MSO_MainTabOperations = False
					Set objTab = Nothing
				End If
				Set objTab = Nothing
				Fn_MSO_MainTabOperations = True
			End If
	End Select
End Function

''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''''/$$$$
''''/$$$$  FUNCTION NAME  	:  Fn_MSO_PerformWorkflowTaskOps
''''/$$$$
''''/$$$$  DESCRIPTION      	:  Function is used to Perform operations on Perform Workflow Task panel.
''''/$$$$ 
''''/$$$$	Return Value 		:  True/False
''''/$$$$	
''''/$$$$  PARAMETERS   	:  sAction 			: 	Action to be performed
''''/$$$$ 					   sAppType 			: 	MSExcel, MSWord or MSPowerPoint
''''/$$$$ 					   dicWorkflowDetails 	: 	Dic details
''''/$$$$ 					   sReserve 			: 	For future use
''''/$$$$ 
''''/$$$$	How To Use 		:  Set dicWorkflowDetails = CreateObject("Scripting.Dictionary")
''''/$$$$						dicWorkflowDetails("NodeName") = "My Worklist:AutoTest1 (autotest1) Inbox (7):Tasks To Perform:000041-Test (New Do Task 1)"
''''/$$$$						dicWorkflowDetails("Comments") = "Item Complted"
''''/$$$$						dicWorkflowDetails("Password") = "abcd"
''''/$$$$					   bReturn = Fn_MSO_PerformWorkflowTaskOps("Complete","MSExcel",dicWorkflowDetails,"")
''''/$$$$					   bReturn = Fn_MSO_PerformWorkflowTaskOps("CompleteWithErrorVerify","MSExcel",dicWorkflowDetails,"")
''''/$$$$					   bReturn = Fn_MSO_PerformWorkflowTaskOps("ExistPanel","MSExcel","","")
''''/$$$$					   	dicWorkflowDetails("NodeName") = "My Worklist:AutoTest6 (autotest6) Inbox (2):Tasks To Perform:000039-Item_123 (New Do Task 1)"
''''/$$$$					   	dicWorkflowDetails("TargetsList") = "000039-Item_123"
''''/$$$$					   bReturn = Fn_MSO_PerformWorkflowTaskOps("VerifyTargetList","MSExcel",dicWorkflowDetails,"")
''''/$$$$					   	dicWorkflowDetails("RadioButton") = "Go to Review Task"
''''/$$$$					   bReturn = Fn_MSO_PerformWorkflowTaskOps("RadioBtnSelect","MSPowerPoint",dicWorkflowDetails,"")
''''/$$$$					   	dicWorkflowDetails("RadioButton") = "Go to Do Task"
''''/$$$$					   bReturn = Fn_MSO_PerformWorkflowTaskOps("RadioBtnExist","MSPowerPoint",dicWorkflowDetails,"")
''''/$$$$					   	dicWorkflowDetails("RadioButton") = "Go to Acknowledge"
''''/$$$$					   	dicWorkflowDetails("Checked") = False
''''/$$$$					   bReturn = Fn_MSO_PerformWorkflowTaskOps("RadioBtnIsChecked","MSPowerPoint",dicWorkflowDetails,"")
''''/$$$$	
''''/$$$$	HISTORY         :  
''''/$$$$					Developer Name		Date	 	Version		Changes							Reviewer
''''/$$$$	Created by  :	 Vivek Ahirrao		10/06/2016	  1.0		Created							[TC1123-20160504-10_06_2016-VivekA-NewDevelopment]
''''/$$$$  												Added for RM - Office Client new TC's Development
''''/$$$$--------------------------------------------------------------------------------------------------------------------
''''/$$$$	Modified by : 	 Vivek Ahirrao		29/06/2016	  1.0		Added cases "SignoffTaskAndViewDecisionsExist","Refresh","RefreshExist"
''''/$$$$--------------------------------------------------------------------------------------------------------------------
''''/$$$$	Modified by : 	 Vivek Ahirrao		18/07/2016	  1.0		Added cases "RadioBtnSelect","RadioBtnExist","RadioBtnIsChecked"
''''/$$$$	
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Public Function Fn_MSO_PerformWorkflowTaskOps(sAction,sAppType,dicWorkflowDetails,sReserve)
	GBL_FAILED_FUNCTION_NAME="Fn_MSO_PerformWorkflowTaskOps"
	Dim objAppWindow, objPrfmWrkflowTskPanel
	Dim bFlag, sButtonName
	
	Fn_MSO_PerformWorkflowTaskOps = False
	bFlag = False
	
	Select Case sAppType
		Case "MSExcel"
			Set objAppWindow = Window("MicrosoftExcel")
			Set objPrfmWrkflowTskPanel = Window("MicrosoftExcel").WinObject("PerformWorkflowTask").SwfObject("SwfObject").WpfWindow("WpfWindow")
		Case "MSWord"
			Set objAppWindow = Window("MicrosoftWord")
			Set objPrfmWrkflowTskPanel = Window("MicrosoftWord").WinObject("PerformWorkflowTask").SwfObject("SwfObject").WpfWindow("WpfWindow")
		Case "MSPowerPoint"
			Set objAppWindow = Window("MicrosoftPowerPoint")
			Set objPrfmWrkflowTskPanel = Window("MicrosoftPowerPoint").WinObject("PerformWorkflowTask").SwfObject("SwfObject").WpfWindow("WpfWindow")
	End Select
	
	If objAppWindow.Exist = False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Office Client App Window ["+sAppType+"] does not Exist.")
		Set objAppWindow = Nothing
		Set objPrfmWrkflowTskPanel = Nothing
		Exit Function
	ElseIf objPrfmWrkflowTskPanel.Exist = False Then
		If varType(dicWorkflowDetails) = "9" Then
			If dicWorkflowDetails("NodeName")<>"" Then
				bFlag = Fn_MSO_MyWorkList_Operations(sAppType,"PopupMenuSelect",dicWorkflowDetails("NodeName"),"View/Perform task...","","","")
				If bFlag = False Then
					Set objAppWindow = Nothing
					Set objPrfmWrkflowTskPanel = Nothing
					Exit Function
				End If
				Wait 1
			End If
		End If
	End If
	
	Select Case sAction
		'Case to check existance of Perform Workflow Task panel
		Case "ExistPanel"
			If objPrfmWrkflowTskPanel.Exist = False Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : [Perform Workflow Task] panel does not Exist.")
				Set objAppWindow = Nothing
				Set objPrfmWrkflowTskPanel = Nothing
				Exit Function
			Else
				Fn_MSO_PerformWorkflowTaskOps = True
			End If
		'Case to Click on Complete button
		Case "Complete","CompleteWithErrorVerify","VerifyPasswordCharacterWithoutComplete"
			If varType(dicWorkflowDetails) = "9" Then
				If dicWorkflowDetails("Comments")<>"" Then
					objPrfmWrkflowTskPanel.WpfObject("PropertyName").SetTOProperty "text","Comments:"
					Wait 0,100
					objPrfmWrkflowTskPanel.WpfEdit("PropertyValue").Set dicWorkflowDetails("Comments")
					If Err.Number<0 Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Set text in [Comments] WPFEdit box.")
						Fn_MSO_PerformWorkflowTaskOps = False
						Set objAppWindow = Nothing
						Set objPrfmWrkflowTskPanel = Nothing
						Exit Function
					End If
				End If
			End If
			If varType(dicWorkflowDetails) = "9" Then
				If dicWorkflowDetails("Password")<>"" Then
					objPrfmWrkflowTskPanel.WpfObject("PropertyName").SetTOProperty "text","Password:"
					Wait 0,100
					objPrfmWrkflowTskPanel.WpfEdit("PropertyValue").Set dicWorkflowDetails("Password")
					If Err.Number<0 Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Set text in [Password] WPFEdit box.")
						Fn_MSO_PerformWorkflowTaskOps = False
						Set objAppWindow = Nothing
						Set objPrfmWrkflowTskPanel = Nothing
						Exit Function
					End If
				End If
			End If
			'Case to verify password is shown as all asterisks
			If sAction = "VerifyPasswordCharacterWithoutComplete" Then
				If varType(dicWorkflowDetails) = "9" Then
					If dicWorkflowDetails("PasswordChar")<>"" Then
						objPrfmWrkflowTskPanel.WpfObject("PropertyName").SetTOProperty "text","Password:"
						Wait 0,100
						sAppText = objPrfmWrkflowTskPanel.WpfEdit("PropertyValue").GetROProperty("text")
						If sAppText = dicWorkflowDetails("PasswordChar") Then
							bFlag = True
						End If
						If bFlag = False Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to match password character.")
							Fn_MSO_PerformWorkflowTaskOps = False
							Set objAppWindow = Nothing
							Set objPrfmWrkflowTskPanel = Nothing
							Exit Function
						End If
					End If
				End If
				Fn_MSO_PerformWorkflowTaskOps = True
				Exit Function
			End If
			'Click on Complete button
			objPrfmWrkflowTskPanel.WpfButton("Complete").SetToProperty "text","Complete"
			bFlag = Fn_MSO_WpfButton_Click("Fn_MSO_PerformWorkflowTaskOps","Click",objPrfmWrkflowTskPanel,"Complete")
			If bFlag = False Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Click on WPFButton [Complete].")
				Fn_MSO_PerformWorkflowTaskOps = False
				Set objAppWindow = Nothing
				Set objPrfmWrkflowTskPanel = Nothing
				Exit Function
			End If
			If sAction = "CompleteWithErrorVerify" Then
				If WpfWindow("PossibleIssues").Exist(5) Then
					If varType(dicWorkflowDetails) = "9" Then
						If dicWorkflowDetails("CompleteErrorMessage")<>"" Then
							sAppText = WpfWindow("PossibleIssues").WpfList("PartialErrorsList").GetVisibleText()
							If Instr(sAppText,dicWorkflowDetails("CompleteErrorMessage"))>0 Then
								Fn_MSO_PerformWorkflowTaskOps = True
								Call Fn_MSO_WpfButton_Click("Fn_MSO_PerformWorkflowTaskOps","Click",WpfWindow("PossibleIssues"),"Close")
							Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Verify Error Message ["+dicWorkflowDetails("CompleteErrorMessage")+"].")
								Fn_MSO_PerformWorkflowTaskOps = False
								Set objAppWindow = Nothing
								Set objPrfmWrkflowTskPanel = Nothing
								Exit Function
							End If
						End If
					End If
				End If
			Else
				Fn_MSO_PerformWorkflowTaskOps = True
			End If
		'to verify Target List items
		Case "VerifyTargetList"
			If objPrfmWrkflowTskPanel.WpfButton("Targets").Exist Then
				If objPrfmWrkflowTskPanel.WpfList("TargetsList").GetROProperty("visible") = False Then
					Call Fn_MSO_WpfButton_Click("Fn_MSO_PerformWorkflowTaskOps","Click",objPrfmWrkflowTskPanel,"Targets")
				End If
				
				If dicWorkflowDetails("TargetsList")<>"" Then
					aTargetList = Split(dicWorkflowDetails("TargetsList"),"~")
					iListCount = objPrfmWrkflowTskPanel.WpfList("TargetsList").GetItemsCount
					For iCount = 0 To UBound(aTargetList)
						bFlag = False
						For iCount1 = 0 To iListCount-1
							sAppNode = objPrfmWrkflowTskPanel.WpfList("TargetsList").GetItem(iCount1)
							If sAppNode = aTargetList(iCount) Then
								bFlag = True
								Exit For
							Else
								'[TC12.1_20180815.00_Maintenance_PoonamC_28Aug2018 : Added Case to get Node from TargetsList ]
								objPrfmWrkflowTskPanel.WpfList("TargetsList").Select(iCount1)
								Wait 1
								sAppNode = objPrfmWrkflowTskPanel.WpfList("TargetsList").GetROProperty("Selection")
								If sAppNode = aTargetList(iCount) Then
									bFlag = True
									Exit For
								End If
							End If
						Next
						If bFlag = False Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed as, Node ["+aTargetList(iCount)+"] does not Exist.")
							Fn_MSO_PerformWorkflowTaskOps = False
							Set objAppWindow = Nothing
							Set objPrfmWrkflowTskPanel = Nothing
							Exit Function
						End If
					Next
				End If
				Fn_MSO_PerformWorkflowTaskOps = True
			End If
		'Case to Click on "SelectSignoffTeam", "Signoff Task And View All Reviewers' Decisions...", "Refresh", "Reassign" or "OK" buttons
		Case "SelectSignoffTeam","SignoffTaskAndViewDecisions","Refresh","Reassign","OK"
			Select Case sAction
				Case "SelectSignoffTeam"
					sButtonName = "Select Signoff Team"
				Case "SignoffTaskAndViewDecisions"
					sButtonName = "Signoff Task And View All Reviewers' Decisions..."
				Case "Refresh"
					sButtonName = "Refresh"
				Case "Reassign"
					sButtonName = "Reassign"
				Case "OK"
					sButtonName = "OK"
			End Select
			objPrfmWrkflowTskPanel.WpfButton("Complete").SetToProperty "text",sButtonName
			bFlag = Fn_MSO_WpfButton_Click("Fn_MSO_PerformWorkflowTaskOps","Click",objPrfmWrkflowTskPanel,"Complete")
			If bFlag = False Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Click on WPFButton ["+sButtonName+"].")
				Fn_MSO_PerformWorkflowTaskOps = False
				Set objAppWindow = Nothing
				Set objPrfmWrkflowTskPanel = Nothing
				Exit Function
			End If
			Fn_MSO_PerformWorkflowTaskOps = True
		'Case to Verify existence of "Signoff Task And View All Reviewers' Decisions..." or "Refresh" buttons
		Case "SignoffTaskAndViewDecisionsExist","RefreshExist"
			Select Case sAction
				Case "SignoffTaskAndViewDecisionsExist"
					sButtonName = "Signoff Task And View All Reviewers' Decisions..."
				Case "RefreshExist"
					sButtonName = "Refresh"
			End Select
			objPrfmWrkflowTskPanel.WpfButton("Complete").SetToProperty "text",sButtonName
			bFlag = objPrfmWrkflowTskPanel.WpfButton("Complete").Exist
			If bFlag = False Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed as WPFButton ["+sButtonName+"] does not Exist.")
				Fn_MSO_PerformWorkflowTaskOps = False
				Set objAppWindow = Nothing
				Set objPrfmWrkflowTskPanel = Nothing
				Exit Function
			End If
			Fn_MSO_PerformWorkflowTaskOps = True
		Case "RadioBtnSelect","RadioBtnExist","RadioBtnIsChecked"
			If dicWorkflowDetails("RadioButton")<>"" Then
				objPrfmWrkflowTskPanel.WpfRadioButton("TaskResultRadioBtn").SetTOProperty "text",dicWorkflowDetails("RadioButton")
				If objPrfmWrkflowTskPanel.WpfRadioButton("TaskResultRadioBtn").Exist Then
					If sAction = "RadioBtnSelect" Then
						objPrfmWrkflowTskPanel.WpfRadioButton("TaskResultRadioBtn").Set
						If Err.Number<>0 Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Set ON WPFRadioButton ["+dicWorkflowDetails("RadioButton")+"].")
							Fn_MSO_PerformWorkflowTaskOps = False
							Set objAppWindow = Nothing
							Set objPrfmWrkflowTskPanel = Nothing
							Exit Function
						End If
					ElseIf sAction = "RadioBtnIsChecked" Then
						If dicWorkflowDetails("Checked")<>"" Then
							bFlag = objPrfmWrkflowTskPanel.WpfRadioButton("TaskResultRadioBtn").GetROProperty("checked")
							If bFlag <> dicWorkflowDetails("Checked") Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Verify WPFRadioButton ["+dicWorkflowDetails("RadioButton")+"] is ["+CStr(dicWorkflowDetails("Checked"))+"].")
								Fn_MSO_PerformWorkflowTaskOps = False
								Set objAppWindow = Nothing
								Set objPrfmWrkflowTskPanel = Nothing
								Exit Function
							End If
						End If
					End If
					Fn_MSO_PerformWorkflowTaskOps = True
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed as WPFRadioButton ["+dicWorkflowDetails("RadioButton")+"] does not Exist.")
					Fn_MSO_PerformWorkflowTaskOps = False
				End If
			End If
	End Select
	
	Set objAppWindow = Nothing
	Set objPrfmWrkflowTskPanel = Nothing
End Function

''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''''/$$$$
''''/$$$$  FUNCTION NAME  	:  Fn_MSO_VerifyPopupMessage
''''/$$$$
''''/$$$$  DESCRIPTION      	:  Function is used to Verify black color Popup message.
''''/$$$$ 
''''/$$$$	Return Value 		:  True/False
''''/$$$$	
''''/$$$$  PARAMETERS   	:  sAction 	: 	Action to be performed
''''/$$$$ 					   sAppType 	: 	MSExcel, MSWord or MSPowerPoint
''''/$$$$ 					   sMessage 	: 	Messgae to be verified
''''/$$$$ 					   sReserve 	: 	For future use
''''/$$$$ 
''''/$$$$	How To Use 		:  bReturn = Fn_MSO_VerifyPopupMessage("Verify","MSExcel","abcd","")
''''/$$$$	
''''/$$$$	HISTORY         :  
''''/$$$$				Developer Name		Date	 	Version	Changes							Reviewer
''''/$$$$	Created by  :	 Vivek Ahirrao		10/06/2016	  1.0		Created							[TC1123-20160504-10_06_2016-VivekA-NewDevelopment]
''''/$$$$  												Added for RM - Office Client new TC's Development
''''/$$$$	Modified by : 
''''/$$$$	
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Public Function Fn_MSO_VerifyPopupMessage(sAction,sAppType,sMessage,sReserve)
	GBL_FAILED_FUNCTION_NAME="Fn_MSO_VerifyPopupMessage"
	GBL_EXPECTED_MESSAGE=sMessage
	Dim sAppMsgeValue
	
	Fn_MSO_VerifyPopupMessage = False

	If WpfWindow("MessageWindow").Exist = False Then
		Exit Function
	End If
	Select Case sAction
		Case "Verify"
			sAppMsgeValue = WpfWindow("MessageWindow").WpfEdit("Message").GetROProperty("text")
			If sAppMsgeValue <> sMessage Then
				GBL_ACTUAL_MESSAGE=sAppMsgeValue
				Fn_MSO_VerifyPopupMessage = False
				Exit Function
			End If
			Fn_MSO_VerifyPopupMessage = True
	End Select
	
End Function

''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''''/$$$$
''''/$$$$  FUNCTION NAME   	:  Fn_MSO_MyWorkList_Operations
''''/$$$$
''''/$$$$  DESCRIPTION      :  Function is used to open Folder or browse view
''''/$$$$ 
''''/$$$$  PRE-REQUISITES  	:  Office Client Window Should be in Focus
''''/$$$$
''''/$$$$  PARAMETERS   	:  sApplication 	: 	MSExcel, MSWord, MSPowerPoint
''''/$$$$ 					   sAction 		: 	Action name
''''/$$$$ 					   StrNode 		: 	Node Path
''''/$$$$ 					   sMenu 			: 	Menu
''''/$$$$ 					   sInfo1 			: 	Future use
''''/$$$$ 					   sInfo2 			: 	Future use
''''/$$$$ 					   sInfo3 			: 	Future use
''''/$$$$	
''''/$$$$	Return Value 	:  True or False
''''/$$$$										
''''/$$$$	How To Use 		:  sNode = "My Worklist:AutoTest1 (autotest1) Inbox:Tasks To Perform:000118-Item_1 (select-signoff-team)"
''''/$$$$					   bReturn = Fn_MSO_MyWorkList_Operations("MSExcel","Select",sNode,"","","","")
''''/$$$$					   bReturn = Fn_MSO_MyWorkList_Operations("MSExcel","Exist",sNode,"","","","")
''''/$$$$					   bReturn = Fn_MSO_MyWorkList_Operations("MSExcel","Expand",sNode,"","","","")
''''/$$$$					   bReturn = Fn_MSO_MyWorkList_Operations("MSExcel","Collapse",sNode,"","","","")
''''/$$$$					   bReturn = Fn_MSO_MyWorkList_Operations("MSExcel","PopupMenuSelect",sNode,"View/Perform task...","","","")
''''/$$$$					   '---------- For 2nd Instance of node use below call -----------------------------------------------------
''''/$$$$					   sNode = "My Worklist:AutoTest3 (autotest3) Inbox:Tasks To Perform:000039-Vivek_1 (perform-signoffs)@2"
''''/$$$$					   bReturn = Fn_MSO_MyWorkList_Operations("MSExcel","Exist",sNode,"","","","")
''''/$$$$					   '--------------------------------------------------------------------------------------------------------
''''/$$$$					   
''''/$$$$	HISTORY     :  Developer Name	     Date		Version		Changes						Reviewer
''''/$$$$------------------------------------------------------------------------------------------------------------------------------------------------------------------
''''/$$$$	Created by  :	 Vivek Ahirrao	 	16/06/2016	  1.0		Created						[TC1123-20160504-16_06_2016-VivekA-NewDevelopment]
''''/$$$$------------------------------------------------------------------------------------------------------------------------------------------------------------------
''''/$$$$	Modified by : 	 Vivek Ahirrao		30/06/2016	  1.0		Modified All cases for Multiple instances of nodes
''''/$$$$	
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Public Function Fn_MSO_MyWorkList_Operations(sApplication,sAction,StrNode,sMenu,sInfo1,sInfo2,sInfo3)
	GBL_FAILED_FUNCTION_NAME="Fn_MSO_MyWorkList_Operations"
	Dim objAppType, objTree, objStackPanel, objWPFWindow, objNode, obj, objMenu
	Dim bFlag, nodeCount, itemcount, iCount, itemLocation, iIndexCounter, iInstance, iInstaCount
	Dim nodetoBselected, sAppText, sVisibleText, sAppTextParent, sVisibleTextParent
	Dim aNode, aVisibleText, aVisibleTextParent, aStrNode
	Dim x, y, iParentCount, sProperties, sValues, iInstaNode
	
	On Error Resume Next
	
	Fn_MSO_MyWorkList_Operations = False
	
	If sApplication="MSExcel" Then
		Set objAppType = Window("MicrosoftExcel")
		Set objTree = Window("MicrosoftExcel").WinObject("FolderView").SwfObject("SwfObject").WpfWindow("WpfWindow").WpfObject("treeCtrl")
'		Set objStackPanel = Window("MicrosoftExcel").WinObject("FolderView").SwfObject("SwfObject").WpfWindow("WpfWindow").WpfObject("StackPanel")
		Set objWPFWindow = Window("MicrosoftExcel").WinObject("FolderView").SwfObject("SwfObject").WpfWindow("WpfWindow")
		'Maximizing [ MicrosoftExcel ] window
		objAppType.Maximize
	ElseIf sApplication="MSWord" Then
		Set objAppType = Window("MicrosoftWord")
		Set objTree = Window("MicrosoftWord").WinObject("FolderView").SwfObject("SwfObject").WpfWindow("WpfWindow").WpfObject("treeCtrl")
'		Set objStackPanel = Window("MicrosoftWord").WinObject("FolderView").SwfObject("SwfObject").WpfWindow("WpfWindow").WpfObject("StackPanel")
		Set objWPFWindow = Window("MicrosoftWord").WinObject("FolderView").SwfObject("SwfObject").WpfWindow("WpfWindow")
		'Maximizing [ MicrosoftWord ] window
		 objAppType.Maximize
	ElseIf sApplication="MSPowerPoint" Then
		Set objAppType = Window("MicrosoftPowerPoint")
		Set objTree = Window("MicrosoftPowerPoint").WinObject("FolderView").SwfObject("SwfObject").WpfWindow("WpfWindow").WpfObject("treeCtrl")
'		Set objStackPanel = Window("MicrosoftPowerPoint").WinObject("FolderView").SwfObject("SwfObject").WpfWindow("WpfWindow").WpfObject("StackPanel")
		Set objWPFWindow = Window("MicrosoftPowerPoint").WinObject("FolderView").SwfObject("SwfObject").WpfWindow("WpfWindow")
		'Maximizing [ MicrosoftPowerPoint ] window
		 objAppType.Maximize
	End If
	wait 1
	If Fn_UI_ObjectExist("Fn_MSO_MyWorkList_Operations",objAppType.WinObject("FolderView")) = False OR objTree.Object.Items.GetItemAt(0).component.DisplayName <> "My Worklist" Then
		bFlag = Fn_MSO_RibbonButton_Operations(sApplication,"Click","My Worklist","")
		If bFlag = False Then
			Set objAppType = Nothing
			Set objTree = Nothing
'			Set objStackPanel = Nothing
			Set objWPFWindow = Nothing
			Fn_MSO_MyWorkList_Operations = False
			Exit Function
		End If
		Wait 5
	End If
	
	Select Case sAction
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to Select
		Case "Select","Exist"
				If Instr(StrNode,"@") Then
					aStrNode = Split(StrNode,"@")
					StrNode = aStrNode(0)
					iInstance = aStrNode(1)
					iInstaCount = 0
					aNode = Split(StrNode,":")
					iInstaNode = aNode(UBound(aNode))
				Else
					StrNode = StrNode
					iInstance = 1
					iInstaCount = 0
					aNode = Split(StrNode,":")
					iInstaNode = ""
				End If
				
				For nodeCount = 0 To UBound(aNode)
					bFlag = False
					nodetoBselected = aNode(nodeCount)
					itemcount= objTree.Object.Items.Count	
					For iCount = 0 To itemcount-1 Step 1
						Set objNode=objTree.Object.Items.GetItemAt(iCount)
						sAppText = objNode.component.DisplayName
						If sAppText = nodetoBselected Then
							If nodeCount = 0 Then
								If UBound(aNode) = 0 Then
									itemLocation=iCount
									If iInstaNode=nodetoBselected Then
										iInstaCount = iInstaCount + 1
										If CInt(iInstaCount)=CInt(iInstance) Then
											bFlag=True
											Exit For
										End If
									Else
										bFlag=True
										Exit For
									End If
								Else
									bFlag=True
									Exit For								
								End If
							Else
								If objNode.Parent.component.DisplayName = aNode(nodeCount-1) Then
									itemLocation=iCount
									If iInstaNode=nodetoBselected Then
										iInstaCount = iInstaCount + 1
										If CInt(iInstaCount)=CInt(iInstance) Then
											bFlag=True
											Exit For
										End If
									Else
										bFlag=True
										Exit For
									End If
								End If
							End If
						ElseIf Instr(nodetoBselected,sAppText)>0 AND Len(nodetoBselected)>Len(sAppText) Then
'							objStackPanel.SetTOProperty "Index",iCount
'							Wait 0,100
'							sVisibleText = objStackPanel.GetVisibleText()
'							If sVisibleText<>"" Then
'								aVisibleText = Split(sVisibleText,"(")
'								If Instr(nodetoBselected,Trim(aVisibleText(0)))>0 Then

'Alternate method for TC1123 to select MyWorklist nodes as StackPanel object is not present in TC1123-------------------------------
							sVisibleText = objWPFWindow.GetVisibleText
							aVisibleText = Split(Replace(sVisibleText,"Loading..."&vblf,""),vblf)
							aAdjustedText = Split(aVisibleText(iCount)," ")
							sAdjustedText = aAdjustedText(0)
							If sAdjustedText<>"" Then
								If Instr(nodetoBselected,Trim(sAdjustedText))>0 Then
'-----------------------------------------------------------------------------------------------------------------------------------
									If nodeCount = 0 Then
										If UBound(aNode) = 0 Then
											itemLocation=iCount								
											If iInstaNode=nodetoBselected Then
												iInstaCount = iInstaCount + 1
												If CInt(iInstaCount)=CInt(iInstance) Then
													bFlag=True
													Exit For
												End If
											Else
												bFlag=True
												Exit For
											End If
										Else
											bFlag=True
											Exit For								
										End If
									Else
										sAppTextParent = objNode.Parent.component.DisplayName
										If sAppTextParent = aNode(nodeCount-1) Then
											itemLocation=iCount
											If iInstaNode=nodetoBselected Then
												iInstaCount = iInstaCount + 1
												If CInt(iInstaCount)=CInt(iInstance) Then
													bFlag=True
													Exit For
												End If
											Else
												bFlag=True
												Exit For
											End If
										ElseIf Instr(aNode(nodeCount-1),sAppTextParent)>0 AND Len(aNode(nodeCount-1))>Len(sAppTextParent) Then
'											objStackPanel.SetTOProperty "Index",iCount-1
'											Wait 0,100
'											sVisibleTextParent = objStackPanel.GetVisibleText()
'											If sVisibleText<>"" Then
'												aVisibleTextParent = Split(sVisibleTextParent,"(")
'												If Instr(aNode(nodeCount-1),Trim(aVisibleTextParent(0)))>0 Then

'Alternate method for TC1123 to select MyWorklist nodes as StackPanel object is not present in TC1123-------------------------------
											sVisibleTextParent = objWPFWindow.GetVisibleText
											aVisibleTextParent = Split(Replace(sVisibleTextParent,"Loading..."&vblf,""),vblf)
											aAdjustedParentText = Split(aVisibleTextParent(iCount)," ")
											sAdjustedParentText = aAdjustedParentText(0)
											If sAdjustedParentText<>"" Then
												If Instr(nodetoBselected,Trim(sAdjustedParentText))>0 Then
'-----------------------------------------------------------------------------------------------------------------------------------
													itemLocation=iCount
													If iInstaNode=nodetoBselected Then
														iInstaCount = iInstaCount + 1
														If CInt(iInstaCount)=CInt(iInstance) Then
															bFlag=True
															Exit For
														End If
													Else
														bFlag=True
														Exit For
													End If
												End If
											End If
										End If
									End If
								End If
							End If
						End If
						Set objNode = Nothing
					Next					
				Next
				
				If bFlag=True Then
					Set obj  = objTree.Object.Items.GetItemAt(itemLocation)
					Fn_MSO_MyWorkList_Operations = objTree.Object.Items.GetItemAt(itemLocation).component.displayname
					
					nodetoBselected = aNode(nodeCount)
					If Fn_MSO_MyWorkList_Operations = nodetoBselected Then
						'Do Nothing
					ElseIf Instr(nodetoBselected,Fn_MSO_MyWorkList_Operations)>0 AND Len(nodetoBselected)>Len(Fn_MSO_MyWorkList_Operations) Then
'						objStackPanel.SetTOProperty "Index",itemLocation
'						Wait 0,100
'						sVisibleText = objStackPanel.GetVisibleText()
'						aVisibleText = Split(sVisibleText,"(")
'						If Instr(nodetoBselected,Trim(aVisibleText(0)))>0 Then

'Alternate method for TC1123 to select MyWorklist nodes as StackPanel object is not present in TC1123-------------------------------
						sVisibleText = objWPFWindow.GetVisibleText
						aVisibleText = Split(Replace(sVisibleText,"Loading..."&vblf,""),vblf)
						aAdjustedText = Split(aVisibleText(itemLocation)," ")
						sAdjustedText = aAdjustedText(0)
						If Instr(nodetoBselected,Trim(sAdjustedText))>0 Then
'-----------------------------------------------------------------------------------------------------------------------------------
							Fn_MSO_MyWorkList_Operations = nodetoBselected
						End If
					End If
					
					iIndexCounter = 0
					Do While IsObject(obj.Parent)
						iIndexCounter = iIndexCounter + 1
						Set obj = obj.Parent
						sApptext = obj.Component.DisplayName
						nodetoBselected = aNode(UBound(aNode)-iIndexCounter)
						If sApptext = nodetoBselected Then
							Fn_MSO_MyWorkList_Operations =  nodetoBselected &":" & Fn_MSO_MyWorkList_Operations
						ElseIf Instr(nodetoBselected,sApptext)>0 AND Len(nodetoBselected)>Len(sApptext) Then
'							objStackPanel.SetTOProperty "Index",itemLocation-iIndexCounter
'							Wait 0,100
'							sVisibleText = objStackPanel.GetVisibleText()
'							aVisibleText = Split(sVisibleText,"(")
'							If Instr(nodetoBselected,Trim(aVisibleText(0)))>0 Then

'Alternate method for TC1123 to select MyWorklist nodes as StackPanel object is not present in TC1123-------------------------------
						sVisibleText = objWPFWindow.GetVisibleText
						aVisibleText = Split(Replace(sVisibleText,"Loading..."&vblf,""),vblf)
						aAdjustedText = Split(aVisibleText(itemLocation)," ")
						sAdjustedText = aAdjustedText(0)
						If Instr(nodetoBselected,Trim(sAdjustedText))>0 Then
'-----------------------------------------------------------------------------------------------------------------------------------
								Fn_MSO_MyWorkList_Operations =  nodetoBselected &":" & Fn_MSO_MyWorkList_Operations
							End If
						End If
					Loop
					
					If Trim(Fn_MSO_MyWorkList_Operations) = Trim(StrNode) Then
						If sAction = "Select" Then
							objTree.Object.SelectedIndex = itemLocation
						End If
						Fn_MSO_MyWorkList_Operations = True
					Else
						Fn_MSO_MyWorkList_Operations = False			
					End If								
				End If
		
		'Case to Expand and Collapse node
		Case "Expand","Collapse"
				If Instr(StrNode,"@") Then
					aStrNode = Split(StrNode,"@")
					StrNode = aStrNode(0)
					iInstance = aStrNode(1)
					iInstaCount = 0
					aNode = Split(StrNode,":")
					iInstaNode = aNode(UBound(aNode))
				Else
					StrNode = StrNode
					iInstance = 1
					iInstaCount = 0
					aNode = Split(StrNode,":")
					iInstaNode = ""
				End If
				nodetoBselected = aNode(Ubound(aNode))
				itemcount= objTree.Object.Items.Count	
				For iCount = 0 To itemcount-1 Step 1
					Set objNode=objTree.Object.Items.GetItemAt(iCount)
					sAppText = objNode.component.DisplayName
					If sAppText = nodetoBselected Then
						itemLocation=iCount
						If iInstaNode=nodetoBselected Then
							iInstaCount = iInstaCount + 1
							If CInt(iInstaCount)=CInt(iInstance) Then
								Exit For
							End If
						Else
							Exit For
						End If
					ElseIf Instr(nodetoBselected,sAppText)>0 AND Len(nodetoBselected)>Len(sAppText) Then
'						objStackPanel.SetTOProperty "Index",iCount
'						Wait 0,100
'						sVisibleText = objStackPanel.GetVisibleText()
'						If sVisibleText<>"" Then
'							aVisibleText = Split(sVisibleText,"(")
'							If Instr(nodetoBselected,Trim(aVisibleText(0)))>0 Then

'Alternate method for TC1123 to select MyWorklist nodes as StackPanel object is not present in TC1123-------------------------------
						sVisibleText = objWPFWindow.GetVisibleText
						aVisibleText = Split(Replace(sVisibleText,"Loading..."&vblf,""),vblf)
						aAdjustedText = Split(aVisibleText(iCount)," ")
						sAdjustedText = aAdjustedText(0)
						If sAdjustedText<>"" Then
							If Instr(nodetoBselected,Trim(sAdjustedText))>0 Then
'-----------------------------------------------------------------------------------------------------------------------------------
								itemLocation=iCount
								If iInstaNode=nodetoBselected Then
									iInstaCount = iInstaCount + 1
									If CInt(iInstaCount)=CInt(iInstance) Then
										Exit For
									End If
								Else
									Exit For
								End If
							End If
						End If
					End If
					Set objNode = nothing
				Next
				If sAction = "Collapse" Then
					If objTree.Object.Items.GetItemAt(itemLocation).IsExpanded <> False Then
						objTree.Object.Items.GetItemAt(itemLocation).IsExpanded = False
						If Err.Number < 0 Then
							Set objAppType = Nothing
							Set objTree = Nothing
'							Set objStackPanel = Nothing
							Set objWPFWindow = Nothing
							Exit Function
						End If
						Fn_MSO_MyWorkList_Operations = True
						Wait 1
					Else
						Fn_MSO_MyWorkList_Operations = True					
					End If
				Else	
					If objTree.Object.Items.GetItemAt(itemLocation).IsExpanded <> True Then
						objTree.Object.Items.GetItemAt(itemLocation).IsExpanded = True
						If Err.Number < 0 Then
							Set objAppType = Nothing
							Set objTree = Nothing
'							Set objStackPanel = Nothing
							Set objWPFWindow = Nothing
							Exit Function
						End If
						Fn_MSO_MyWorkList_Operations = True
						Wait 2
					Else
						Fn_MSO_MyWorkList_Operations = True					
					End If					
				End If
				
		'Case for PopUpMenu Selection
		Case "PopupMenuSelect"
				If Instr(StrNode,"@") Then
					aStrNode = Split(StrNode,"@")
					StrNode = aStrNode(0)
					iInstance = aStrNode(1)
					iInstaCount = 0
					aNode = Split(StrNode,":")
					iInstaNode = aNode(UBound(aNode))
				Else
					StrNode = StrNode
					iInstance = 1
					iInstaCount = 0
					aNode = Split(StrNode,":")
					iInstaNode = ""
				End If				
				For nodeCount = 0 To UBound(aNode)
					bFlag = False
					nodetoBselected = aNode(nodeCount)
					itemcount= objTree.Object.Items.Count	
					For iCount = 0 To itemcount-1 Step 1
						Set objNode=objTree.Object.Items.GetItemAt(iCount)
						sAppText = objNode.component.DisplayName
						If sAppText = nodetoBselected Then
							If nodeCount = 0 Then
								If UBound(aNode) = 0 Then
									itemLocation=iCount								
									If iInstaNode=nodetoBselected Then
										iInstaCount = iInstaCount + 1
										If CInt(iInstaCount)=CInt(iInstance) Then
											bFlag=True
											Exit For
										End If
									Else
										bFlag=True
										Exit For
									End If
								Else
									bFlag=True
									Exit For								
								End If
							Else
								If objNode.Parent.component.DisplayName = aNode(nodeCount-1) Then
									itemLocation=iCount
									If iInstaNode=nodetoBselected Then
										iInstaCount = iInstaCount + 1
										If CInt(iInstaCount)=CInt(iInstance) Then
											bFlag=True
											Exit For
										End If
									Else
										bFlag=True
										Exit For
									End If
								End If
							End If
						ElseIf Instr(nodetoBselected,sAppText)>0 AND Len(nodetoBselected)>Len(sAppText) Then
'							objStackPanel.SetTOProperty "Index",iCount
'							Wait 0,100
'							sVisibleText = objStackPanel.GetVisibleText()
'							If sVisibleText<>"" Then
'								aVisibleText = Split(sVisibleText,"(")
'								If Instr(nodetoBselected,Trim(aVisibleText(0)))>0 Then

'Alternate method for TC1123 to select MyWorklist nodes as StackPanel object is not present in TC1123-------------------------------
							sVisibleText = objWPFWindow.GetVisibleText
							aVisibleText = Split(Replace(sVisibleText,"Loading..."&vblf,""),vblf)
							aAdjustedText = Split(aVisibleText(iCount)," ")
							sAdjustedText = aAdjustedText(0)
							If sAdjustedText<>"" Then
								If Instr(nodetoBselected,Trim(sAdjustedText))>0 Then
'-----------------------------------------------------------------------------------------------------------------------------------
									If nodeCount = 0 Then
										If UBound(aNode) = 0 Then
											itemLocation=iCount								
											If iInstaNode=nodetoBselected Then
												iInstaCount = iInstaCount + 1
												If CInt(iInstaCount)=CInt(iInstance) Then
													bFlag=True
													Exit For
												End If
											Else
												bFlag=True
												Exit For
											End If
										Else
											bFlag=True
											Exit For								
										End If
									Else
										sAppTextParent = objNode.Parent.component.DisplayName
										If sAppTextParent = aNode(nodeCount-1) Then
											itemLocation=iCount
											If iInstaNode=nodetoBselected Then
												iInstaCount = iInstaCount + 1
												If CInt(iInstaCount)=CInt(iInstance) Then
													bFlag=True
													Exit For
												End If
											Else
												bFlag=True
												Exit For
											End If
										ElseIf Instr(aNode(nodeCount-1),sAppTextParent)>0 AND Len(aNode(nodeCount-1))>Len(sAppTextParent) Then
'											objStackPanel.SetTOProperty "Index",iCount-1
'											Wait 0,100
'											sVisibleTextParent = objStackPanel.GetVisibleText()
'											If sVisibleTextParent<>"" Then
'												aVisibleTextParent = Split(sVisibleTextParent,"(")
'												If Instr(aNode(nodeCount-1),Trim(aVisibleTextParent(0)))>0 Then

'Alternate method for TC1123 to select MyWorklist nodes as StackPanel object is not present in TC1123-------------------------------
											sVisibleTextParent = objWPFWindow.GetVisibleText
											aVisibleTextParent = Split(Replace(sVisibleTextParent,"Loading..."&vblf,""),vblf)
											aAdjustedParentText = Split(aVisibleTextParent(iCount)," ")
											sAdjustedParentText = aAdjustedParentText(0)
											If sAdjustedParentText<>"" Then
												If Instr(nodetoBselected,Trim(sAdjustedParentText))>0 Then
'-----------------------------------------------------------------------------------------------------------------------------------
													itemLocation=iCount
													If iInstaNode=nodetoBselected Then
														iInstaCount = iInstaCount + 1
														If CInt(iInstaCount)=CInt(iInstance) Then
															bFlag=True
															Exit For
														End If
													Else
														bFlag=True
														Exit For
													End If
												End If
											End If
										End If
									End If
								End If
							End If
						End If
						Set objNode = nothing
					Next					
				Next

'				x=30
'				y=5
'				For iParentCount = 0 To UBound(aNode)-1
'					x=x+20
'				Next
			
				If bFlag=True Then
'					objStackPanel.SetTOProperty "index",itemLocation
'					objStackPanel.Click x,y,micRightBtn
'					Wait 2
'					sProperties ="Class Name~classname" 
'					sValues = "ContextMenu~System.Windows.Controls.ContextMenu"

'Alternate method for TC1123 to select MyWorklist nodes as StackPanel object is not present in TC1123-------------------------------
					objWPFWindow.WpfButton("SelectStartComponent").SetTOProperty "index",itemLocation
					Wait 1
					xPos = objWPFWindow.WpfButton("SelectStartComponent").GetROProperty("abs_x")
					yPos = objWPFWindow.WpfButton("SelectStartComponent").GetROProperty("abs_y")
					objWPFWindow.WpfButton("SelectStartComponent").Click 5,5,micLeftBtn
					Set objMercury = CreateObject("Mercury.DeviceReplay")
					objMercury.MouseClick xPos+30,yPos+10,micLeftBtn
					Wait 0,200
					Set objMercury = Nothing
					Set wshshell = CreateObject("WScript.Shell")
					wshshell.SendKeys "+{F10}"
					Wait 2
					Set wshshell = Nothing
					sProperties ="Class Name~classname" 
					sValues = "ContextMenu~System.Windows.Controls.ContextMenu"
'-----------------------------------------------------------------------------------------------------------------------------------
					Set objMenu = Fn_SISW_UI_Object_GetChildObjects("Fn_MSO_MyWorkList_Operations",objWPFWindow, sProperties, sValues)
			
					For iCount = 0 To objMenu.count-1 
						If Instr(objMenu(iCount).toString(), "WpfMenu") > 0 Then
							sMenu = Replace(sMenu,":",";")
							objMenu(iCount).Select sMenu
							Exit For
						End If	
					Next
			 		
			 		If Err.Number<0 Then
				   		Fn_MSO_MyWorkList_Operations = False
					Else
						Fn_MSO_MyWorkList_Operations = True
					End If
				End If
	End Select
	Set objAppType = Nothing
	Set objTree = Nothing
'	Set objStackPanel = Nothing
	Set objWPFWindow = Nothing
End Function

''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''''/$$$$
''''/$$$$  FUNCTION NAME  	:  Fn_MSO_DeleteOperations
''''/$$$$
''''/$$$$  DESCRIPTION      	:  Function is used to Perform Delete operations any node in Folder View Tree.
''''/$$$$ 
''''/$$$$	Return Value 		:  True/False
''''/$$$$	
''''/$$$$  PARAMETERS   	:  sAction 	: 	Action to be performed
''''/$$$$ 					   sAppType 	: 	MSExcel, MSWord or MSPowerPoint
''''/$$$$ 					   dicDetails 	: 	Dic details
''''/$$$$ 					   sButton 	: 	Button Name
''''/$$$$ 					   sReserve 	: 	For future use
''''/$$$$ 
''''/$$$$	How To Use 		:  Set dicDetails = CreateObject("Scripting.Dictionary")
''''/$$$$						  dicDetails("ItemsToDelete") = "Home:000041-Test"	(Not mandatory) (You can give call for Delete in Script also)
''''/$$$$						  dicDetails("DeleteDialogButton") = "Yes"		(Not mandatory)
''''/$$$$						  dicDetails("ErrorMessage") = "The selected object(s) cannot be deleted (e.g. they may be checked-out): "+vblf+"000041-Test"+vblf+""+vblf+""
''''/$$$$						  dicDetails("ErrorDialogButton") = "OK"		(Not mandatory)
''''/$$$$					   bReturn = Fn_MSO_DeleteOperations("DeleteWithErrorVerify","MSExcel",dicDetails,"","")
''''/$$$$	
''''/$$$$	HISTORY         :  
''''/$$$$				Developer Name		Date	 	Version	Changes							Reviewer
''''/$$$$	Created by  :	 Vivek Ahirrao		10/06/2016	  1.0		Created							[TC1123-20160504-10_06_2016-VivekA-NewDevelopment]
''''/$$$$  												Added for RM - Office Client new TC's Development
''''/$$$$	Modified by : 
''''/$$$$	
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Public Function Fn_MSO_DeleteOperations(sAction,sAppType,dicDetails,sButton,sReserve)
	GBL_FAILED_FUNCTION_NAME="Fn_MSO_DeleteOperations"
	Dim objAppType, objDeleteDialog, objErrorDialog
	Dim bFlag, sAppText
	
	Fn_MSO_DeleteOperations = False
	
	Select Case sAppType
		Case "MSExcel"
			Set objAppType = Window("MicrosoftExcel")
			Set objDeleteDialog = Window("MicrosoftExcel").Dialog("DialogInformation")
			objDeleteDialog.SetTOProperty "text","Delete Confirmation"
			objAppType.Maximize
		Case "MSWord"
			Set objAppType = Window("MicrosoftWord")
			Set objDeleteDialog = Window("MicrosoftWord").Dialog("DialogInformation")
			objDeleteDialog.SetTOProperty "text","Delete Confirmation"
			objAppType.Maximize
		Case "MSPowerPoint"
			Set objAppType = Window("MicrosoftPowerPoint")
			Set objDeleteDialog = Window("MicrosoftPowerPoint").Dialog("DialogInformation")
			objDeleteDialog.SetTOProperty "text","Delete Confirmation"
			objAppType.Maximize
	End Select
	Wait 1
	
	'PopupMenuSelect on Item which you want to delete
	If dicDetails("ItemsToDelete")<>"" AND objDeleteDialog.Exist(5) = False Then
		bFlag = Fn_MSO_FolderViewTreeOperations(sAppType,"PopupMenuSelect",dicDetails("ItemsToDelete"),"Delete","","","")
		If bFlag = False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to perform PopupMenuSelect [Delete] menu on ["+dicDetails("ItemsToDelete")+"] node.")
			Set objAppType = Nothing
			Set objDeleteDialog = Nothing
			Exit Function
		End If
		Wait 1
	End If
	
	Select Case sAction
		'Delete Item or Delete Item with Error verify
		Case "Delete", "DeleteWithErrorVerify"
			'Check Delete Confirmation dialog exist or not
			If objDeleteDialog.Exist(5) Then
				'Click on Yes if you eant to delete
				If dicDetails("DeleteDialogButton")<>"" Then
					bFlag = Fn_UI_WinButton_Click("Fn_MSO_DeleteOperations",objDeleteDialog,dicDetails("DeleteDialogButton"),5,5,micLeftBtn)
				Else
					bFlag = Fn_UI_WinButton_Click("Fn_MSO_DeleteOperations",objDeleteDialog,"Yes",5,5,micLeftBtn)
				End If
				If bFlag = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Click on Button on Delete Confirmation dialog.")
					Set objAppType = Nothing
					Set objDeleteDialog = Nothing
					Exit Function
				End If
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Delete Confirmation dialog does not exist.")
				Set objAppType = Nothing
				Set objDeleteDialog = Nothing
				Exit Function
			End If
			If sAction<>"Delete" Then
				If Fn_SISW_UI_Object_Operations("Fn_MSO_DeleteOperations", "Exist", Dialog("ConfirmationBox"), "") = True Then
					'Check Error message dialog exist or not
					Set objErrorDialog = Dialog("ConfirmationBox")
					If objErrorDialog.Exist(5) Then
						bFlag = False
						'Verify Error message
						If dicDetails("ErrorMessage")<>"" Then
							sAppText = objErrorDialog.Static("TextMessage").GetROProperty("text")
							If sAppText <> dicDetails("ErrorMessage") Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Verify Delete Error Message ["+dicDetails("ErrorMessage")+"].")
								Set objAppType = Nothing
								Set objDeleteDialog = Nothing
								Set objErrorDialog = Nothing
								Exit Function
							End If
						End If
						'Click on OK button on Error message dialog
						If dicDetails("ErrorDialogButton")<>"" Then
							bFlag = Fn_UI_WinButton_Click("Fn_MSO_DeleteOperations",objErrorDialog,dicDetails("ErrorDialogButton"),5,5,micLeftBtn)
						Else
							bFlag = Fn_UI_WinButton_Click("Fn_MSO_DeleteOperations",objErrorDialog,"OK",5,5,micLeftBtn)
						End If
						If bFlag = False Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Click on Button on Delete Erroe Message dialog.")
							Set objAppType = Nothing
							Set objDeleteDialog = Nothing
							Set objErrorDialog = Nothing
							Exit Function
						End If
					End If
				ElseIf Fn_SISW_UI_Object_Operations("Fn_MSO_DeleteOperations", "Exist", WpfWindow("PossibleIssues"), "") = True Then	
					Set objErrorDialog = WpfWindow("PossibleIssues")
					If dicDetails("ErrorMessage") <> "" Then
						If Not Instr(1, objErrorDialog.WpfList("PartialErrorsList").GetVisibleText(), dicDetails("ErrorMessage")) > 0 Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Verify Delete Error Message [" + dicDetails("ErrorMessage") + "].")
							Set objAppType = Nothing
							Set objDeleteDialog = Nothing
							Set objErrorDialog = Nothing
							Exit Function
						End If
	
						If Fn_MSO_WpfButton_Click("Fn_MSO_DeleteOperations", "Click", objErrorDialog.WpfButton("Close"), "") = False Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Click on Close Button of Possible Issuese dialog.")
							Set objAppType = Nothing
							Set objDeleteDialog = Nothing
							Set objErrorDialog = Nothing
							Exit Function
						End If
					End If
				End If
			End If
			Fn_MSO_DeleteOperations = True
			Set objErrorDialog = Nothing
			
		'sAction can not be empty
		Case Else
			Fn_MSO_DeleteOperations = False
			Set objAppType = Nothing
			Set objDeleteDialog = Nothing
	End Select
	
	Set objAppType = Nothing
	Set objDeleteDialog = Nothing
End Function

''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''''/$$$$
''''/$$$$  FUNCTION NAME   	:  Fn_MSO_NavTreeInWindow_Operations
''''/$$$$
''''/$$$$  DESCRIPTION      :  Function is used to perform Nav tree operations in Dialog, where Nav tree is present
''''/$$$$ 
''''/$$$$  PRE-REQUISITES  	:  The dialog should be opened.
''''/$$$$
''''/$$$$  PARAMETERS   	:  sAction 		: 	Action name
''''/$$$$ 					   sAppType		: 	MSExcel, MSWord, MSPowerPoint
''''/$$$$ 					   sWindow 		: 	Window dialog in which Nav tree is present
''''/$$$$ 					   StrNode 		: 	Node Path
''''/$$$$ 					   sInfo1 		: 	Future use
''''/$$$$ 					   sInfo2 		: 	Future use
''''/$$$$ 					   sInfo3 		: 	Future use
''''/$$$$	
''''/$$$$	Return Value 	:  True or False
''''/$$$$										
''''/$$$$	How To Use 		:  bReturn = Fn_MSO_NavTreeInWindow_Operations("Select","MSExcel","TeamcenterSaveAsDialog","Home:000038-TestItem","","","")
''''/$$$$					   bReturn = Fn_MSO_NavTreeInWindow_Operations("Expand","MSExcel","TeamcenterSaveAsDialog","Home:000038-TestItem","","","")
''''/$$$$					   bReturn = Fn_MSO_NavTreeInWindow_Operations("Exist","MSExcel","TeamcenterSaveAsDialog","Home:000038-TestItem:000038","","","")
''''/$$$$	
''''/$$$$	HISTORY         :  
''''/$$$$					Developer Name	     Date		Version		Changes						Reviewer
''''/$$$$	Created by  :	 Vivek Ahirrao	 16/06/2016	  	 1.0		Created						[TC1123-20160504-16_06_2016-VivekA-NewDevelopment]
''''/$$$$  
''''/$$$$	Modified by : 
''''/$$$$	
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Public Function Fn_MSO_NavTreeInWindow_Operations(sAction,sAppType,sWindow,StrNode,sInfo1,sInfo2,sInfo3)
	GBL_FAILED_FUNCTION_NAME="Fn_MSO_NavTreeInWindow_Operations"
	Dim objTree, objNode, obj
	Dim aNode, nodetoBselected
	Dim nodeCount, bFlag, itemcount, iCount, itemLocation, iRowCounter
	
	On Error Resume Next
	Fn_MSO_NavTreeInWindow_Operations = False
	
	Select Case sAppType
		Case "MSExcel"
			'Future Use
		Case "MSWord"
			'Future Use
		Case "MSPowerPoint"
			'Future Use
	End Select
	
	Select Case sWindow
		Case "TeamcenterSaveAsDialog"
			Set objTree = WpfWindow("TeamcenterSaveAs").WpfObject("FolderViewTree")
		Case "CreateNewItem"
			'Future Use
		Case "CreateNewFolder"
			'Future Use
		Case "SelectSignoffTeamOrgChart"
			Set objTree = WpfWindow("SelectSignoffTeam").WpfObject("OrganizationTree")
		Case "SelectDelegateUser"
			Set objTree = WpfWindow("Delegate").WpfObject("OrganizationTree")
		Case "TeamcenterOpenDialog"
			Set objTree = WpfWindow("TeamcenterOpen").WpfObject("FolderViewTree")
		Case Else
			Exit Function
	End Select
	
	If objTree.Exist(2) = False Then
		Set objTree = Nothing
		Exit Function
	End If
	
	Select Case sAction
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to get node Location or Index
		Case "GetItemLocation"
				aNode=Split(StrNode,":")
				iRowCounter = 0
				For nodeCount = 0 To UBound(aNode)
					bFlag = False
					nodetoBselected = aNode(nodeCount)
					itemcount = objTree.Object.Items.Count
					For iCount = iRowCounter To itemcount-1 Step 1
						Set objNode=objTree.Object.Items.GetItemAt(iCount)
						If objNode.component.DisplayName = nodetoBselected Then
							iRowCounter = iCount+1
							If nodeCount = 0 Then
								If UBound(aNode) = 0 Then
									itemLocation=iCount								
									bFlag=True
									Exit For
								Else
									bFlag=True
									Exit For								
								End If
							Else
								If objNode.Parent.component.DisplayName = aNode(nodeCount-1) Then
									itemLocation=iCount
									bFlag=True
									Exit For
								End If
							End If
						End If
						Set objNode = nothing
					Next					
				Next
				If bFlag=True Then
					Set obj  = objTree.Object.Items.GetItemAt(itemLocation)
					Fn_MSO_NavTreeInWindow_Operations=objTree.Object.Items.GetItemAt(itemLocation).component.displayname
					Do while IsObject(obj.Parent)
						Set obj = obj.Parent
						Fn_MSO_NavTreeInWindow_Operations = obj.Component.DisplayName &":" &Fn_MSO_NavTreeInWindow_Operations
					Loop
					If Trim(Fn_MSO_NavTreeInWindow_Operations) = Trim(StrNode) Then
						Fn_MSO_NavTreeInWindow_Operations = itemLocation
					Else
						Fn_MSO_NavTreeInWindow_Operations = False
					End If								
				End If
		'Case to Select node
		Case "Select"
				itemLocation = Fn_MSO_NavTreeInWindow_Operations("GetItemLocation",sAppType,sWindow,StrNode,"","","")
				If CStr(itemLocation)<>CStr(False) Then
					objTree.Object.SelectedIndex = itemLocation
					Fn_MSO_NavTreeInWindow_Operations = True
				End If
		'Case to Verify whether Node Exist
		Case "Exist"
				itemLocation = Fn_MSO_NavTreeInWindow_Operations("GetItemLocation",sAppType,sWindow,StrNode,"","","")
				If CStr(itemLocation)<>CStr(False) Then
					Fn_MSO_NavTreeInWindow_Operations = True
				End If
		'Case to Verify whether Node IsSelected
		Case "IsSelected"
				itemLocation = Fn_MSO_NavTreeInWindow_Operations("GetItemLocation",sAppType,sWindow,StrNode,"","","")
				If CStr(itemLocation)<>CStr(False) Then
					If objTree.Object.SelectedIndex <> itemLocation Then
						Fn_MSO_NavTreeInWindow_Operations = False
						Set objTree = Nothing
						Exit Function
					End If
					Fn_MSO_NavTreeInWindow_Operations = True
				End If
		'Case to Expand node
		Case "Expand"
				itemLocation = Fn_MSO_NavTreeInWindow_Operations("GetItemLocation",sAppType,sWindow,StrNode,"","","")
				If CStr(itemLocation)<>CStr(False) Then
					If objTree.Object.Items.GetItemAt(itemLocation).IsExpanded <> True Then
						objTree.Object.Items.GetItemAt(itemLocation).IsExpanded = True
						If Err.Number < 0 Then
							Set objTree = Nothing							
							Exit Function
						End If
						Wait 2
					End If
					Fn_MSO_NavTreeInWindow_Operations = True
				End If
		'Case to Collapse node
		Case "Collapse"
				itemLocation = Fn_MSO_NavTreeInWindow_Operations("GetItemLocation",sAppType,sWindow,StrNode,"","","")
				If CStr(itemLocation)<>CStr(False) Then
					If objTree.Object.Items.GetItemAt(itemLocation).IsExpanded <> False Then
						objTree.Object.Items.GetItemAt(itemLocation).IsExpanded = False
						If Err.Number < 0 Then
							Set objTree = Nothing
							Exit Function
						End If
						Wait 1
					End If
					Fn_MSO_NavTreeInWindow_Operations = True
				End If
	End Select
	Set objTree = Nothing
End Function

''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''''/$$$$
''''/$$$$  FUNCTION NAME   :  Fn_MSO_TeamcenterSaveAs_Operations
''''/$$$$
''''/$$$$  DESCRIPTION     	:  Function is used to perform operations on Teamcenter Save As dialog
''''/$$$$ 
''''/$$$$  PRE-REQUISITES  	:  Office Client (MS-Excel, MS-Word or MS-PowerPoint) Window Should be opened
''''/$$$$
''''/$$$$  PARAMETERS   	:  sAction 		: 	Action to be performed
''''/$$$$ 					   sAppType 		: 	MSExcel, MSWord or MSPowerPoint
''''/$$$$ 					   StrNode 		: 	Node Path
''''/$$$$ 					   dicSaveDetails 	: 	Dictionary object
''''/$$$$ 					   sButton 		: 	Button to be clicked
''''/$$$$ 					   sReserve 		: 	For future use
''''/$$$$	
''''/$$$$	Return Value 		:  True or False
''''/$$$$										
''''/$$$$	How To Use 		:  Set dicSaveDetails = CreateObject("Scripting.Dictionary")
''''/$$$$						  dicSaveDetails("EditBox1") 		= "File Name:Book1234"
''''/$$$$						  dicSaveDetails("EditBox2") 		= "Description:Excel Book"
''''/$$$$						  dicSaveDetails("ComboBox1") 	= "Relation:Rendering"
''''/$$$$						  dicSaveDetails("ComboBox2") 	= "Dataset type:MSExcelX"
''''/$$$$						  sNode = "Home:000038-TestItem:000038/A;1-TestItem"
''''/$$$$					   bReturn = Fn_MSO_TeamcenterSaveAs_Operations("Save","MSExcel",sNode,dicSaveDetails,"Save","")
''''/$$$$	
''''/$$$$	HISTORY         :  
''''/$$$$					Developer Name		Date		Version		Changes								Reviewer
''''/$$$$	Created by  :	Vivek Ahirrao	 	10/06/2016	  1.0		Created								[TC1123-20160504-10_06_2016-VivekA-NewDevelopment]
''''/$$$$  
''''/$$$$	Modified by : 
''''/$$$$	
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Public Function Fn_MSO_TeamcenterSaveAs_Operations(sAction,sAppType,StrNode,dicSaveDetails,sButton,sReserve)
	GBL_FAILED_FUNCTION_NAME="Fn_MSO_TeamcenterSaveAs_Operations"
	Dim objTCSaveAs
	Dim bFlag, iCount, sNode, iCounter, sSubAction, sProperty
	Dim aStrNode, dicCount, dicItems, dicKeys, aProperty
	
	Fn_MSO_TeamcenterSaveAs_Operations = False
	On Error Resume Next
	
	Set objTCSaveAs = WpfWindow("TeamcenterSaveAs")
	If objTCSaveAs.Exist(5) = False Then
		bFlag = Fn_MSO_RibbonButton_Operations(sAppType,"Click","Save As:Dataset","")
		If bFlag = False Then
			Set objTCSaveAs = Nothing
			Exit Function
		End If
	End If
	If objTCSaveAs.Exist(5) = False Then
		Set objTCSaveAs = Nothing
		Exit Function
	End If
	
	If StrNode<>"" Then
		aStrNode = Split(StrNode,":")
		For iCount = 0 To UBound(aStrNode)
			If iCount = 0 Then
				sNode = aStrNode(iCount)
			Else
				sNode = sNode + ":" + aStrNode(iCount)
			End If
			
			bFlag = Fn_MSO_NavTreeInWindow_Operations("Select",sAppType,"TeamcenterSaveAsDialog",sNode,"","","")
			If bFlag = False Then
				Set objTCSaveAs = Nothing
				Exit Function
			End If
			Wait 1
			If iCount <> UBound(aStrNode) Then
				bFlag = Fn_MSO_NavTreeInWindow_Operations("Expand",sAppType,"TeamcenterSaveAsDialog",sNode,"","","")
				If bFlag = False Then
					Set objTCSaveAs = Nothing
					Exit Function
				End If
				Wait 2
			End If
		Next
	End If
	
	Select Case sAction
		Case "Save"
			dicCount = dicSaveDetails.Count
			dicItems = dicSaveDetails.Items
			dicKeys = dicSaveDetails.Keys
			
			For iCounter = 0 To dicCount - 1
				If Instr(dicKeys(iCounter),"EditBox")>0 Then
					sSubAction = "EditBox"
				ElseIf Instr(dicKeys(iCounter),"ComboBox")>0 Then
					sSubAction = "ComboBox"
				Else
					sSubAction = dicKeys(iCounter)
				End If
				sProperty = dicItems(iCounter)
				bFlag = False
				Select Case sSubAction
					'"File Name:Book1"
					Case "EditBox"
						If sProperty<>"" Then
							aProperty = Split(sProperty,":")
							'Set text property 
							'objTCSaveAs.WpfObject("PropertyName").SetTOProperty "text",aProperty(0) & ":"
							If aProperty(0) = "File Name" Then
								objTCSaveAs.WpfEdit("PropertyValue").SetToProperty "devname","sDatasetName"
							ElseIf aProperty(0) = "Description" Then
								objTCSaveAs.WpfEdit("PropertyValue").SetToProperty "devname","sDatasetDescription"
							End If
							If objTCSaveAs.WpfEdit("PropertyValue").Exist Then
								objTCSaveAs.WpfEdit("PropertyValue").Set aProperty(1)
								If Err.Number < 0 Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Set ["+aProperty(1)+"] in ["+aProperty(0)+"] WPF Edit Box.")
									Set objTCSaveAs = Nothing
									Exit Function
								End If
								Wait 1
								bFlag = True
							End If
						End If
					'"Type:Item Revision"
					Case "ComboBox"
						If sProperty<>"" Then
							aProperty = Split(sProperty,":")
							If aProperty(0) = "Save as type" Then
								If objTCSaveAs.WpfComboBox("supportedFileTypes").Exist Then
									objTCSaveAs.WpfComboBox("supportedFileTypes").Click 5,5,micLeftBtn
									Wait 1
									Err.Clear
									objTCSaveAs.WpfComboBox("supportedFileTypes").Type aProperty(1)
									Wait 1
									Err.Clear
									objTCSaveAs.WpfComboBox("supportedFileTypes").Select aProperty(1)
									If Err.Number < 0 Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Select ["+aProperty(1)+"] from ["+aProperty(0)+"] WPF Combo Box.")
										Set objTCSaveAs = Nothing
										Exit Function
									End If
									Wait 1
									bFlag = True
								End If
							Else
								'Set text property 
								objTCSaveAs.WpfObject("PropertyName").SetTOProperty "devname",aProperty(0) & ":"
								Wait 1
								If objTCSaveAs.WpfObject("PropertyName").Exist Then
									objTCSaveAs.WpfComboBox("PropertyComboBox").Click 5,5,micLeftBtn
									Wait 1
									Err.Clear
									objTCSaveAs.WpfComboBox("PropertyComboBox").Type aProperty(1)
									Wait 1
									Err.Clear
									objTCSaveAs.WpfComboBox("PropertyComboBox").Select aProperty(1)
									If Err.Number < 0 Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Select ["+aProperty(1)+"] from ["+aProperty(0)+"] WPF Combo Box.")
										Set objTCSaveAs = Nothing
										Exit Function
									End If
									Wait 1
									bFlag = True
								End If
							End If
						End If
				End Select
				
				If bFlag = False Then
					Fn_MSO_TeamcenterSaveAs_Operations = False
					Set objTCSaveAs = Nothing
					Exit Function
				End If
			Next
			
			'Click on button provided
			If sButton<>"" Then
				bFlag = Fn_MSO_WpfButton_Click("Fn_MSO_AdvancedSearchOps","Click",objTCSaveAs,sButton)
				If bFlag = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Click on WPFButton ["+sButton+"].")
					Fn_MSO_TeamcenterSaveAs_Operations = False
					Set objTCSaveAs = Nothing
					Exit Function
				End If
				Wait 1
			End If
			
			Fn_MSO_TeamcenterSaveAs_Operations = True
	Case "Verify"
		
			dicCount = dicSaveDetails.Count
			dicItems = dicSaveDetails.Items
			dicKeys = dicSaveDetails.Keys
			
			For iCounter = 0 To dicCount - 1
				If Instr(dicKeys(iCounter),"EditBox")>0 Then
					sSubAction = "EditBox"
				ElseIf Instr(dicKeys(iCounter),"ComboBox")>0 Then
					sSubAction = "ComboBox"
				Else
					sSubAction = dicKeys(iCounter)
				End If
				sProperty = dicItems(iCounter)
				bFlag = False
				Select Case sSubAction
					'"File Name:Book1"
					Case "EditBox"
						If sProperty<>"" Then
							aProperty = Split(sProperty,":")
							'Set text property 
							'objTCSaveAs.WpfObject("PropertyName").SetTOProperty "text",aProperty(0) & ":"
							If aProperty(0) = "File Name" Then
								objTCSaveAs.WpfEdit("PropertyValue").SetTOProperty "devname","sDatasetName"
							ElseIf aProperty(0) = "Description" Then
								objTCSaveAs.WpfEdit("PropertyValue").SetTOProperty "devname","sDatasetDescription"
							End If
							If objTCSaveAs.WpfEdit("PropertyValue").Exist Then
								
								If trim(objTCSaveAs.WpfEdit("PropertyValue").GetROProperty("text"))=trim(aProperty(1)) Then
									Wait 1
									bFlag = True
								Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Set ["+aProperty(1)+"] in ["+aProperty(0)+"] WPF Edit Box.")
									Set objTCSaveAs = Nothing
									Exit Function
								End If
								
							End If
						End If
					'"Type:Item Revision"
					Case "ComboBox"
						If sProperty<>"" Then
							aProperty = Split(sProperty,":")
							If aProperty(0) = "Save as type" Then
								If objTCSaveAs.WpfComboBox("supportedFileTypes").Exist Then
									If trim(objTCSaveAs.WpfComboBox("supportedFileTypes").GetROProperty("selection"))=trim(aProperty(1)) Then
										Wait 1
										bFlag = True
									Else
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Select ["+aProperty(1)+"] from ["+aProperty(0)+"] WPF Combo Box.")
										Set objTCSaveAs = Nothing
										Exit Function
									End If
								End If
							Else
								'Set text property 
								objTCSaveAs.WpfObject("PropertyName").SetTOProperty "devname",aProperty(0) & ":"
								Wait 1
								If objTCSaveAs.WpfObject("PropertyName").Exist Then
									If trim(objTCSaveAs.WpfComboBox("PropertyComboBox").GetROProperty("selection"))=trim(aProperty(1)) Then
										Wait 1
										bFlag = True
									Else
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Select ["+aProperty(1)+"] from ["+aProperty(0)+"] WPF Combo Box.")
										Set objTCSaveAs = Nothing
										Exit Function
									End If
								End If
							End If
						End If
				End Select
				
				If bFlag = False Then
					Fn_MSO_TeamcenterSaveAs_Operations = False
					Set objTCSaveAs = Nothing
					Exit Function
				End If
			Next
			
			'Click on button provided
			If sButton<>"" Then
				bFlag = Fn_MSO_WpfButton_Click("Fn_MSO_AdvancedSearchOps","Click",objTCSaveAs,sButton)
				If bFlag = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Click on WPFButton ["+sButton+"].")
					Fn_MSO_TeamcenterSaveAs_Operations = False
					Set objTCSaveAs = Nothing
					Exit Function
				End If
				Wait 1
			End If
			
			Fn_MSO_TeamcenterSaveAs_Operations = True
	End Select
	Set objTCSaveAs = Nothing
End Function

''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''''/$$$$
''''/$$$$  FUNCTION NAME   :  Fn_MSO_SelectSignoffTeam_Ops
''''/$$$$
''''/$$$$  DESCRIPTION     	:  Function is used to perform operations on Select Signoff Team dialog
''''/$$$$ 
''''/$$$$  PRE-REQUISITES  	:  Select Signoff Team dialog should be opened
''''/$$$$
''''/$$$$  PARAMETERS   	:  sAction 		: 	Action to be performed
''''/$$$$ 					   sAppType 		: 	MSExcel, MSWord or MSPowerPoint
''''/$$$$ 					   dicSSTDetails 	: 	Dictionary object
''''/$$$$ 					   sButton 		: 	Button to be clicked
''''/$$$$ 					   sReserve 		: 	For future use
''''/$$$$	
''''/$$$$	Return Value 		:  True or False
''''/$$$$										
''''/$$$$	How To Use 		:  Set dicSSTDetails = CreateObject("Scripting.Dictionary")
''''/$$$$						  dicSSTDetails("SelectSignoffTeamNode") 	= "Signoff Team:Users"
''''/$$$$						  dicSSTDetails("EnterSearchString") 			= "SearchForGroup:Engineering"
''''/$$$$						  dicSSTDetails("SelectOrgTreeNode1") 		= "Engineering:Designer:AutoTest1 (autotest1)"
''''/$$$$						  dicSSTDetails("SelectAction1") 				= "Review"
''''/$$$$						  dicSSTDetails("Add1") 					= "AddUser"
''''/$$$$						  dicSSTDetails("SelectOrgTreeNode2") 		= "Engineering:Designer:AutoTest2 (autotest2)"
''''/$$$$						  dicSSTDetails("SelectAction2") 				= "Notify"
''''/$$$$						  dicSSTDetails("Add2") 					= "AddUser"
''''/$$$$						  dicSSTDetails("VerifySignoffTeamNode") 		= "Signoff Team:Users:AutoTest1 (autotest1) - Engineering/Designer" & "~" & "Signoff Team:Users:AutoTest2 (autotest2) - Engineering/Designer"
''''/$$$$						  dicSSTDetails("VerifySignoffTeamAction1") 	= "Signoff Team:Profiles:Engineering/Designer/1:AutoTest1 (autotest1) - Engineering/Designer$Review"
''''/$$$$						  dicSSTDetails("VerifySignoffTeamAction2") 	= "Signoff Team:Users:AutoTest2 (autotest2) - Engineering/Designer$Review"
''''/$$$$					   bReturn = Fn_MSO_SelectSignoffTeam_Ops("SelectSignoffTeam","MSExcel",dicSSTDetails,"Submit","")
''''/$$$$	
''''/$$$$	HISTORY         :  
''''/$$$$				Developer Name		Date		Version	Changes								Reviewer
''''/$$$$	Created by  :	Vivek Ahirrao	 	20/06/2016	  1.0		Created								[TC1123-20160504-20_06_2016-VivekA-NewDevelopment]
''''/$$$$  
''''/$$$$	Modified by : 
''''/$$$$	
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Public Function Fn_MSO_SelectSignoffTeam_Ops(sAction,sAppType,dicSSTDetails,sButton,sReserve)
	GBL_FAILED_FUNCTION_NAME="Fn_MSO_SelectSignoffTeam_Ops"
	Dim objSSTeam
	Dim dicCount, dicItems, dicKeys, aProperty, aAppVisibleText, aUser
	Dim iCounter, iCount, iItemCount, iCount1, iRow
	Dim sSubAction, sProperty, bFlag, sNode, sAppNode, sSrchBtnName, sBtnName
	Dim sAppVisibleText, sUser, sAAction, bFlag1
	
	Fn_MSO_SelectSignoffTeam_Ops = False
	On Error Resume Next
	
	Select Case sAppType
		Case "MSExcel"
			Set objSSTeam = WpfWindow("SelectSignoffTeam")
		Case "MSWord"
			Set objSSTeam = WpfWindow("SelectSignoffTeam")
		Case "MSPowerPoint"
			Set objSSTeam = WpfWindow("SelectSignoffTeam")
	End Select
	
	If sAction = "Reassign" Then
		objSSTeam.SetTOProperty "regexpwndtitle","Reassign"
		wait 1
		objSSTeam.SetTOProperty "devname","Reassign"
		wait 1
	End If
	
	If objSSTeam.Exist(5) = False Then
		Set objSSTeam = Nothing
		Exit Function
	End If
	
	Select Case sAction
		Case "SelectSignoffTeam","Reassign"
			dicCount = dicSSTDetails.Count
			dicItems = dicSSTDetails.Items
			dicKeys = dicSSTDetails.Keys
			
			For iCounter = 0 To dicCount - 1
				If Instr(dicKeys(iCounter),"SelectSignoffTeamNode")>0 Then
					sSubAction = "SelectSignoffTeamNode"
				ElseIf Instr(dicKeys(iCounter),"VerifySignoffTeamNode")>0 Then
					sSubAction = "VerifySignoffTeamNode"
				ElseIf Instr(dicKeys(iCounter),"VerifySignoffTeamAction")>0 Then
					sSubAction = "VerifySignoffTeamAction"
				ElseIf Instr(dicKeys(iCounter),"SelectOrgTreeNode")>0 Then
					sSubAction = "SelectOrgTreeNode"
				ElseIf Instr(dicKeys(iCounter),"VerifyOrgTreeNode")>0 Then
					sSubAction = "VerifyOrgTreeNode"
				ElseIf Instr(dicKeys(iCounter),"EnterSearchString")>0 Then
					sSubAction = "EnterSearchString"
				ElseIf Instr(dicKeys(iCounter),"SelectAction")>0 Then
					sSubAction = "SelectAction"
				ElseIf Instr(dicKeys(iCounter),"Add")>0 Then
					sSubAction = "Add"
				ElseIf Instr(dicKeys(iCounter),"Remove")>0 Then
					sSubAction = "Remove"
				Else
					sSubAction = dicKeys(iCounter)
				End If
				sProperty = dicItems(iCounter)
				bFlag = False
				Select Case sSubAction
					'"Signoff Team:Users"
					Case "SelectSignoffTeamNode"
						If sProperty<>"" Then
							aProperty = Split(sProperty,":")
							'Expand the Node
							For iCount = 0 To UBound(aProperty)
								If iCount = 0 Then
									sNode = aProperty(iCount)
								Else
									sNode = sNode + ";" + aProperty(iCount)
								End If
								If iCount <> UBound(aProperty) Then
									objSSTeam.WpfTreeView("SignoffTeamTree").Expand  sNode
									If Err.Number<0 Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Expand ["+sNode+"].")
										Set objSSTeam = Nothing
										Exit Function
									End If
									Wait 2
								End If
							Next
							
							'Select Node
							objSSTeam.WpfTreeView("SignoffTeamTree").Select sNode
							If Err.Number < 0 Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Select ["+sNode+"].")
								Set objSSTeam = Nothing
								Exit Function
							End If
							Wait 2
							bFlag = True
						End If
					'"Signoff Team:Users:AutoTest1 (autotest1) - Engineering/Designer" & "~" & "Signoff Team:Users:AutoTest2 (autotest2) - Engineering/Designer"
					Case "VerifySignoffTeamNode"
						If sProperty<>"" Then
							aProperty = Split(sProperty,"~")
							iItemCount = objSSTeam.WpfTreeView("SignoffTeamTree").GetItemsCount()
							For iCount = 0 To UBound(aProperty)
								bFlag = False
								sNode = Replace(aProperty(iCount),":",";")
								For iCount1 = 0 To iItemCount-1
									sAppNode = objSSTeam.WpfTreeView("SignoffTeamTree").GetItem(iCount1)
									If sAppNode = sNode Then
										bFlag = True
										Exit For
									End If
								Next
								If bFlag = False Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Node not found ["+aProperty(iCount)+"].")
									Set objSSTeam = Nothing
									Exit Function
								End If
							Next
							bFlag = True
						End If
					'"Signoff Team:Users:AutoTest2 (autotest2) - Engineering/Designer$Review"
					Case "VerifySignoffTeamAction"
						If sProperty<>"" Then
							sAppVisibleText = objSSTeam.WpfTreeView("SignoffTeamTree").GetVisibleText()
							aAppVisibleText = Split(Replace(sAppVisibleText,vblf,"*"),"*")
							aProperty = Split(sProperty,"$")
							'One row at a time
							sUser = aProperty(0)
							sAAction = aProperty(1)
							
							aUser = Split(sUser,":")
							iRow = 0
							For iCount = 0 To UBound(aUser)
								bFlag1 = False
								For iCount1 = iRow To UBound(aAppVisibleText)
									If aAppVisibleText(iCount1)<>"" AND aAppVisibleText(iCount1)<>" " Then
										If iCount = 0 Then
											If Instr(aAppVisibleText(iCount1),aUser(iCount))>0 Then
												iRow = iCount1+1
												bFlag1 = True
												Exit For
											End If
										Else
											If Trim(aUser(iCount))=Trim(aAppVisibleText(iCount1)) Then
												iRow = iCount1+1
												bFlag1 = True
												Exit For
											End If
										End If
									End If
								Next
								If bFlag1 = False Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Node not found ["+aUser(iCount)+"].")
									Set objSSTeam = Nothing
									Exit Function
								End If
							Next
							If Trim(aAppVisibleText(iRow)) = Trim(sAAction) Then
								bFlag = True
							Else
								bFlag = False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Action Value not found ["+sAAction+"] for node ["+sUser+"].")
							End If
						End If
					'"SearchForGroup:Engineering"
					Case "EnterSearchString"
						If sProperty<>"" Then
							'Click on Refresh button
							bFlag = Fn_MSO_WpfButton_Click("Fn_MSO_SelectSignoffTeam_Ops","Click",objSSTeam,"RefreshOrgChart")
							If bFlag = False Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Click on [RefreshOrgChart] WPFButton.")
								Set objSSTeam = Nothing
								Exit Function
							End If
							
							aProperty = Split(sProperty,":")
							Select Case aProperty(0)
								Case "SearchForGroup"
									sSrchBtnName = "SearchForGroup"
								Case "SearchForRole"
									sSrchBtnName = "SearchForRole"
								Case "SearchForUser"
									sSrchBtnName = "SearchForUser"
							End Select
							'Enter Search String
							objSSTeam.WpfEdit("EnterSearchString").Set aProperty(1)
							Wait 1
							'Click on Group, Role or User button to search
							bFlag = Fn_MSO_WpfButton_Click("Fn_MSO_SelectSignoffTeam_Ops","Click",objSSTeam,sSrchBtnName)
							If bFlag = False Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Click on ["+sSrchBtnName+"] WPFButton.")
								Set objSSTeam = Nothing
								Exit Function
							End If
							If aProperty(0) = "SearchForUser" Then
								If objSSTeam.Dialog("SearchforUser").Exist(5) Then
									objSSTeam.Dialog("SearchforUser").WinButton("Yes").Click 5,5,micLeftBtn
									If Err.Number<0 Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Click on [Yes] WinButton in Search for User dialog.")
										Set objSSTeam = Nothing
										Exit Function
									End If
								End If
							End If
							Wait 5
						End If
					'"Engineering:Designer:AutoTest1 (autotest1)"
					Case "SelectOrgTreeNode","VerifyOrgTreeNode"
						If sProperty<>"" Then					
							aProperty = Split(sProperty,":")
							'To Expand node
							For iCount = 0 To UBound(aProperty)
								If iCount = 0 Then
									sNode = aProperty(iCount)
								Else
									sNode = sNode + ":" + aProperty(iCount)
								End If
									
								If iCount <> UBound(aProperty) Then
									bFlag = Fn_MSO_NavTreeInWindow_Operations("Expand",sAppType,"SelectSignoffTeamOrgChart",sNode,"","","")
									If bFlag = False Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Expand ["+sNode+"].")
										Set objSSTeam = Nothing
										Exit Function
									End If
									Wait 2
								End If
							Next
							
							If sSubAction = "SelectOrgTreeNode" Then
								'To Select node
								bFlag = Fn_MSO_NavTreeInWindow_Operations("Select",sAppType,"SelectSignoffTeamOrgChart",sProperty,"","","")
							ElseIf sSubAction = "VerifyOrgTreeNode" Then
								'To Verify node
								bFlag = Fn_MSO_NavTreeInWindow_Operations("Exist",sAppType,"SelectSignoffTeamOrgChart",sProperty,"","","")
							End If
							If bFlag = False Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to ["+sSubAction+"] node ["+sProperty+"].")
								Set objSSTeam = Nothing
								Exit Function
							End If
							Wait 1
						End If
					'Case to Select Action:Review/Acknowledge/Notify
					Case "SelectAction"
						If sProperty<>"" Then
							objSSTeam.WpfList("ActionWpfList").Select sProperty
							If Err.Number<0 Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Select ["+sProperty+"] from Action WPFList.")
								Set objSSTeam = Nothing
								Exit Function
							End If
							bFlag = True
						End If
					'"AddUser"
					Case "Add","Remove"
						'Click on Add or Remove
						If sProperty<>"" Then
							sBtnName = sProperty
						End If
						bFlag = Fn_MSO_WpfButton_Click("Fn_MSO_SelectSignoffTeam_Ops","Click",objSSTeam,sBtnName)
						If bFlag = False Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Click on ["+sBtnName+"] WPFButton.")
							Set objSSTeam = Nothing
							Exit Function
						End If
						Wait 1
				End Select
				
				If bFlag = False Then
					Fn_MSO_SelectSignoffTeam_Ops = False
					Set objSSTeam = Nothing
					Exit Function
				End If
			Next
			
			'Click on button provided
			If sButton<>"" Then
				bFlag = Fn_MSO_WpfButton_Click("Fn_MSO_SelectSignoffTeam_Ops","Click",objSSTeam,sButton)
				If bFlag = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Click on WPFButton ["+sButton+"].")
					Fn_MSO_SelectSignoffTeam_Ops = False
					Set objSSTeam = Nothing
					Exit Function
				End If
				'Wait 1
			End If
			
			Fn_MSO_SelectSignoffTeam_Ops = True
	End Select
	
	'Reset the dialog properties in OR
	If sAction = "Reassign" Then
		objSSTeam.SetTOProperty "regexpwndtitle","Select Signoff Team"
		wait 1
		objSSTeam.SetTOProperty "devname","Select Signoff Team"
		wait 1
	End If
	Set objSSTeam = Nothing
End Function

''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''''/$$$$
''''/$$$$  FUNCTION NAME   	:  Fn_MSO_SignoffTaskAndReviewersDecisions_Ops
''''/$$$$
''''/$$$$  DESCRIPTION     	:  Function is used to perform operations on dialog comes on click "Signoff Task And View All Reviewers' Decisions" button in Perform Workflow Task panel
''''/$$$$ 
''''/$$$$  PRE-REQUISITES  	:  "Signoff Task And View All Reviewers' Decisions" button should be clicked and dialog should be opened.
''''/$$$$
''''/$$$$  PARAMETERS   	:  sAction 			: 	Action to be performed
''''/$$$$ 					   sAppType 		: 	MSExcel, MSWord or MSPowerPoint
''''/$$$$ 					   sDecision 		: 	Decision column value
''''/$$$$ 					   sReviewer 		: 	Reviewers name
''''/$$$$ 					   sComments 		: 	Cooments column value
''''/$$$$ 					   sDelegate 		: 	Delegate column
''''/$$$$ 					   dicSTRDDetails 	: 	Dictionary object for Future Use
''''/$$$$ 					   sButton 			: 	Button to be clicked
''''/$$$$ 					   sReserve 		: 	For future use
''''/$$$$	
''''/$$$$	Return Value 	:  True or False
''''/$$$$										
''''/$$$$	How To Use 		:  bReturn = Fn_MSO_SignoffTaskAndReviewersDecisions_Ops("GetRowCount","MSExcel","","","","","","","")
''''/$$$$					   bReturn = Fn_MSO_SignoffTaskAndReviewersDecisions_Ops("GetRowIndexCurrentUser","MSExcel","","","","","","","")
''''/$$$$					   bReturn = Fn_MSO_SignoffTaskAndReviewersDecisions_Ops("GetRowIndexCurrentUser","MSExcel","","AutoTest4 (autotest4)","","","","","")
''''/$$$$					   bReturn = Fn_MSO_SignoffTaskAndReviewersDecisions_Ops("GetRowIndexReviewerUser","MSExcel","","AutoTest3 (autotest3)","","","","","")
''''/$$$$					   bReturn = Fn_MSO_SignoffTaskAndReviewersDecisions_Ops("VerifyColumnValues","MSExcel","Approve~No Decision","AutoTest4 (autotest4)~AutoTest3 (autotest3)","","","","","")
''''/$$$$					   bReturn = Fn_MSO_SignoffTaskAndReviewersDecisions_Ops("SetDetails","MSExcel","Approve","","abcdefgh","","","","")
''''/$$$$					   bReturn = Fn_MSO_SignoffTaskAndReviewersDecisions_Ops("DelegateClick","MSExcel","","AutoTest4 (autotest4)","","","","","")
''''/$$$$					   Set dicSTRDDetails = CreateObject("Scripting.Dictionary")
''''/$$$$					   	   dicSTRDDetails("Password") = "abcd"
''''/$$$$					   bReturn = Fn_MSO_SignoffTaskAndReviewersDecisions_Ops("SetDetails","MSExcel","Approve","","abcdefgh","",dicSTRDDetails,"","")
''''/$$$$					   	   dicSTRDDetails("ErrorMessage") = "Incorrect password entered for secured task. Please enter the correct password."
''''/$$$$					   bReturn = Fn_MSO_SignoffTaskAndReviewersDecisions_Ops("ErrorVerify","MSExcel","","","","",dicSTRDDetails,"","")
''''/$$$$					   	   dicSTRDDetails("ColumnValues") = "Engineering~Designer"
''''/$$$$					   bReturn = Fn_MSO_SignoffTaskAndReviewersDecisions_Ops("VerifyOtherColValues","MSExcel","","AutoTest4 (autotest4)","","",dicSTRDDetails,"","")
''''/$$$$					   
''''/$$$$					   sMsg = "Your current group/role does not match the required group/role '*/Tc_QALead'. Would you like to change your current user setting to 'Tc_Enterprise/Tc_QALead'?"
''''/$$$$					   Set dicSTRDDetails = CreateObject("Scripting.Dictionary")
''''/$$$$					   	  dicSTRDDetails("ErrorMessage") = sMsg
''''/$$$$					   	  dicSTRDDetails("Button") = "Yes"
''''/$$$$					   bReturn = Fn_MSO_SignoffTaskAndReviewersDecisions_Ops("DecisionErrorVerify","MSExcel","Approve","","","",dicSTRDDetails,"","")
''''/$$$$	
''''/$$$$	HISTORY 	:  	Developer Name		Date	Version		Changes								Reviewer
''''/$$$$------------------------------------------------------------------------------------------------------------------------------------------------------------------
''''/$$$$	Created by  :	Vivek Ahirrao	 22/06/2016	  1.0		Created								[TC1123-20160504-22_06_2016-VivekA-NewDevelopment]
''''/$$$$------------------------------------------------------------------------------------------------------------------------------------------------------------------
''''/$$$$	Modified by : 	Vivek Ahirrao	 01/07/2016	  1.0		Added cases "VerifyOtherColValues", "ErrorVerify", "ExistTaskDialog"
''''/$$$$------------------------------------------------------------------------------------------------------------------------------------------------------------------
''''/$$$$	Modified by : 	Vivek Ahirrao	 11/07/2016	  1.0		Added case "DecisionErrorVerify"
''''/$$$$	
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Public Function Fn_MSO_SignoffTaskAndReviewersDecisions_Ops(sAction,sAppType,sDecision,sReviewer,sComments,sDelegate,dicSTRDDetails,sButton,sReserve)
	GBL_FAILED_FUNCTION_NAME="Fn_MSO_SignoffTaskAndReviewersDecisions_Ops"
	Dim objSTRDecisions, objSTRDList, objDesc, objChild
	Dim iReviewer, iRowCount, iCurrent, iCount, iCounter
	Dim sAppListText, sAppText, bFlag
	Dim aAppListText, aReviewer, aDecision
	
	Fn_MSO_SignoffTaskAndReviewersDecisions_Ops = False
	On Error Resume Next
	
	Select Case sAppType
		Case "MSExcel"
			Set objSTRDecisions = WpfWindow("SignoffAndReviewersDecisions")
			Set objSTRDList = WpfWindow("SignoffAndReviewersDecisions").WpfList("ReviewerDecisionsList")
		Case "MSWord"
			Set objSTRDecisions = WpfWindow("SignoffAndReviewersDecisions")
			Set objSTRDList = WpfWindow("SignoffAndReviewersDecisions").WpfList("ReviewerDecisionsList")
		Case "MSPowerPoint"
			Set objSTRDecisions = WpfWindow("SignoffAndReviewersDecisions")
			Set objSTRDList = WpfWindow("SignoffAndReviewersDecisions").WpfList("ReviewerDecisionsList")
	End Select
	
	If sAction <> "ErrorVerify" Then
		If objSTRDecisions.Exist(5) = False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Dialog [SignoffAndReviewersDecisions] does not exist.")
			Set objSTRDecisions = Nothing
			Set objSTRDList = Nothing
			Exit Function
		End If
	End If
	
	Select Case sAction
		Case "ExistTaskDialog"
			If objSTRDecisions.Exist(5)= True Then
				Set objSTRDecisions = Nothing
				Fn_MSO_SignoffTaskAndReviewersDecisions_Ops = True
				Exit Function
			End If
		'GetRowCount
		Case "GetRowCount"
			'Get Row index where Combobox is present
			Set objDesc = Description.Create
			objDesc("wpftypename").value = "edit"
			Set objChild = objSTRDList.ChildObjects(objDesc)
			Fn_MSO_SignoffTaskAndReviewersDecisions_Ops = objChild.Count
			If Fn_MSO_SignoffTaskAndReviewersDecisions_Ops <= 0 Then
				Fn_MSO_SignoffTaskAndReviewersDecisions_Ops = False
			End If
			Set objDesc = Nothing
			Set objChild = Nothing
		'Get Row Index for Current user
		Case "GetRowIndexCurrentUser"
			If sReviewer<>"" Then
				iReviewer = Fn_MSO_SignoffTaskAndReviewersDecisions_Ops("GetRowIndexReviewerUser",sAppType,"",sReviewer,"","","","","")
				If iReviewer = -1 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Reviewer ["+sReviewer+"] does not exist.")
					Set objSTRDecisions = Nothing
					Set objSTRDList = Nothing
					Exit Function
				End If
			End If
			'Get Row index where Combobox is present
			iRowCount = Fn_MSO_SignoffTaskAndReviewersDecisions_Ops("GetRowCount",sAppType,"","","","","","","")
			iCurrent = -1
			For iCount = 0 To iRowCount-1
				objSTRDList.WpfEdit("CommentsEditBox").SetTOProperty "Index",iCount
				If objSTRDList.WpfComboBox("DecisionComboBox").Exist(1) = True Then
					iCurrent = iCount
					Exit For
				End If
			Next
			If sReviewer<>"" Then
				If iReviewer = iCurrent Then
					Fn_MSO_SignoffTaskAndReviewersDecisions_Ops = iCurrent
				Else
					Fn_MSO_SignoffTaskAndReviewersDecisions_Ops = -1
				End If
			Else
				Fn_MSO_SignoffTaskAndReviewersDecisions_Ops = iCurrent				
			End If
			If Fn_MSO_SignoffTaskAndReviewersDecisions_Ops = -1 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Current User does not exist or match.")
				Set objSTRDecisions = Nothing
				Set objSTRDList = Nothing
				Exit Function
			End If
		'Get Row Index of Reviewer User provided
		'sReviewer = "AutoTest4 (autotest4)"
		Case "GetRowIndexReviewerUser"
			If sReviewer<>"" Then
				sAppListText = objSTRDList.GetVisibleText()
				aAppListText = Split(Replace(sAppListText,vblf,"*"),"*")
				iCounter = -1
				Fn_MSO_SignoffTaskAndReviewersDecisions_Ops = -1
				For iCount = 0 To UBound(aAppListText)
					If Instr(aAppListText(iCount),"(")>0 AND Instr(aAppListText(iCount),")")>0 Then
						iCounter = iCounter+1
						If Instr(aAppListText(iCount),sReviewer)>0 Then
							Fn_MSO_SignoffTaskAndReviewersDecisions_Ops = iCounter
							Exit For
						End If
					End If
				Next
				If Fn_MSO_SignoffTaskAndReviewersDecisions_Ops = -1 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Reviewer ["+sReviewer+"] does not exist.")
					Set objSTRDecisions = Nothing
					Set objSTRDList = Nothing
					Exit Function
				End If
			End If
		'sDecision = "No Decision~No Decision"
		'sReviewer = "AutoTest4 (autotest4)~AutoTest3 (autotest3)"
		Case "VerifyColumnValues"
			If sReviewer<>"" Then
				aReviewer = Split(sReviewer,"~")
				aDecision = Split(sDecision,"~")
				For iCount = 0 To UBound(aReviewer)
					iRowCount = Fn_MSO_SignoffTaskAndReviewersDecisions_Ops("GetRowIndexReviewerUser",sAppType,"",aReviewer(iCount),"","","","","")
					'Set index of edit box
					objSTRDList.WpfEdit("CommentsEditBox").SetTOProperty "Index",iRowCount
					If objSTRDList.WpfObject("DecisionObject").Exist(1) = True Then
						sAppText = objSTRDList.WpfObject("DecisionObject").GetROProperty("text")
						If sAppText <> aDecision(iCount) Then
							Fn_MSO_SignoffTaskAndReviewersDecisions_Ops = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Verify Column [Decision] : Value ["+aDecision(iCount)+"] for Reviewer ["+aReviewer(iCount)+"] .")
							Set objSTRDecisions = Nothing
							Set objSTRDList = Nothing
							Exit Function
						End If
					End If
				Next
				Fn_MSO_SignoffTaskAndReviewersDecisions_Ops = True
			End If
		'Verify one row at a time
		'sReviewer = "AutoTest4 (autotest4)"
		'dicSTRDDetails("ColumnValues") = "Engineering~Designer"
		Case "VerifyOtherColValues"
			If varType(dicSTRDDetails)<>"9" Then
				Fn_MSO_SignoffTaskAndReviewersDecisions_Ops = False
				Set objSTRDecisions = Nothing
				Set objSTRDList = Nothing
				Exit Function
			End If
			If sReviewer<>"" Then				
				sAppListText = objSTRDList.GetVisibleText()
				bFlag = False
				aAppListText = Split(Replace(sAppListText,vblf,"#"),"#")
				For iCount = 0 To UBound(aAppListText)
					If Instr(aAppListText(iCount),"(")>0 AND Instr(aAppListText(iCount),")")>0 OR Instr(aAppListText(iCount),"*")>0 Then
						If Instr(aAppListText(iCount),sReviewer)>0 Then
							aColValues = Split(dicSTRDDetails("ColumnValues"),"~")
							For iCount1 = 0 To UBound(aColValues)
								bFlag = False
								'Verify Column name exist in Row or not
								If Instr(aAppListText(iCount),aColValues(iCount1))>0 Then
									bFlag = True
								End If
								If bFlag = False Then
									Exit For
								End If
							Next
						End If
					End If
					If bFlag = True Then
						Exit For
					End If
				Next
				If bFlag = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Column Values ["+dicSTRDDetails("ColumnValues")+"] for Reviewer ["+sReviewer+"] does not exist.")
					Fn_MSO_SignoffTaskAndReviewersDecisions_Ops = False
					Set objSTRDecisions = Nothing
					Set objSTRDList = Nothing
					Exit Function
				End If
				Fn_MSO_SignoffTaskAndReviewersDecisions_Ops = True
			End If
		'Case to set Decision, comments
		'sDecision = "Approve"
		'sReviewer = "abcd"
		'dicSTRDDetails("Password")="abcd"
		Case "SetDetails"
			If sReviewer<>"" Then
				iRowCount = Fn_MSO_SignoffTaskAndReviewersDecisions_Ops("GetRowIndexCurrentUser",sAppType,"",sReviewer,"","","","","")
			Else
				iRowCount = Fn_MSO_SignoffTaskAndReviewersDecisions_Ops("GetRowIndexCurrentUser",sAppType,"","","","","","","")
			End If
			If iRowCount = -1 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Reviewer does not Exist.")
				Set objSTRDecisions = Nothing
				Set objSTRDList = Nothing
				Exit Function
			End If
			objSTRDList.WpfEdit("CommentsEditBox").SetTOProperty "Index",iRowCount
			If sComments<>"" Then
				objSTRDList.WpfEdit("CommentsEditBox").Set sComments
				If Err.Number < 0 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Set ["+sComments+"] in [Comments] WPF Edit Box.")
					Fn_MSO_SignoffTaskAndReviewersDecisions_Ops = False
					Set objSTRDecisions = Nothing
					Set objSTRDList = Nothing
					Exit Function
				End If
				Wait 1
			End If
			If sDecision<>"" Then
				objSTRDList.WpfComboBox("DecisionComboBox").Select sDecision
				If Err.Number < 0 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Select ["+sDecision+"] from [Decision] WPF Combo Box.")
					Fn_MSO_SignoffTaskAndReviewersDecisions_Ops = False
					Set objSTRDecisions = Nothing
					Set objSTRDList = Nothing
					Exit Function
				End If
				Wait 1
			End If
			'For Future use
			If varType(dicSTRDDetails)="9" Then
				'Enter Password
				If dicSTRDDetails("Password")<>"" Then
					objSTRDecisions.WpfEdit("PasswordEditBox").Set dicSTRDDetails("Password")
					If Err.Number < 0 Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Set ["+dicSTRDDetails("Password")+"] in [Password] WPF Edit Box.")
						Fn_MSO_SignoffTaskAndReviewersDecisions_Ops = False
						Set objSTRDecisions = Nothing
						Set objSTRDList = Nothing
						Exit Function
					End If
					Wait 1
				End If
			End If
			Fn_MSO_SignoffTaskAndReviewersDecisions_Ops = True
		'Case to Click on Delegate button
		Case "DelegateClick"
			If sReviewer<>"" Then
				iRowCount = Fn_MSO_SignoffTaskAndReviewersDecisions_Ops("GetRowIndexCurrentUser",sAppType,"",sReviewer,"","","","","")
			Else
				iRowCount = Fn_MSO_SignoffTaskAndReviewersDecisions_Ops("GetRowIndexCurrentUser",sAppType,"","","","","","","")
			End If
			If iRowCount = -1 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Reviewer does not Exist.")
				Set objSTRDecisions = Nothing
				Set objSTRDList = Nothing
				Exit Function
			End If
			objSTRDList.WpfEdit("CommentsEditBox").SetTOProperty "Index",iRowCount
			If objSTRDList.WpfButton("DelegateBtn").Exist(1) = True Then
				bFlag = Fn_MSO_WpfButton_Click("Fn_MSO_SignoffTaskAndReviewersDecisions_Ops","Click",objSTRDList,"DelegateBtn")
				If bFlag = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Click on WPFButton [DelegateBtn].")
					Fn_MSO_SignoffTaskAndReviewersDecisions_Ops = False
					Set objSTRDecisions = Nothing
					Set objSTRDList = Nothing
					Exit Function
				End If
				Wait 1
				Fn_MSO_SignoffTaskAndReviewersDecisions_Ops = True
			End If
		'Case to verify Possible issue error
		Case "ErrorVerify"
			If varType(dicSTRDDetails)="9" Then
				If Fn_SISW_UI_Object_Operations("Fn_MSO_SignoffTaskAndReviewersDecisions_Ops", "Exist", WpfWindow("PossibleIssues"), "") = True Then	
					Set objErrorDialog = WpfWindow("PossibleIssues")
					If dicSTRDDetails("ErrorMessage") <> "" Then
						If Not Instr(1, objErrorDialog.WpfList("PartialErrorsList").GetVisibleText(), dicSTRDDetails("ErrorMessage")) > 0 Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Verify Error Message [" + dicSTRDDetails("ErrorMessage") + "].")
							Set objSTRDecisions = Nothing
							Set objSTRDList = Nothing
							Set objErrorDialog = Nothing
							Exit Function
						End If

						If Fn_MSO_WpfButton_Click("Fn_MSO_SignoffTaskAndReviewersDecisions_Ops", "Click", objErrorDialog.WpfButton("Close"), "") = False Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Click on Close Button of Possible Issuese dialog.")
							Set objSTRDecisions = Nothing
							Set objSTRDList = Nothing
							Set objErrorDialog = Nothing
							Exit Function
						End If
						Fn_MSO_SignoffTaskAndReviewersDecisions_Ops = True
						Set objErrorDialog = Nothing
					End If
				End If
			End If
		Case "DecisionErrorVerify"
			If varType(dicSTRDDetails)="9" Then
				If sDecision<>"" Then
					objSTRDList.WpfComboBox("DecisionComboBox").Select sDecision
					If objSTRDecisions.Dialog("ChangeUserSetting").Exist(5) Then
						sAppText = objSTRDecisions.Dialog("ChangeUserSetting").Static("StaticText").GetROProperty("text")
						If dicSTRDDetails("ErrorMessage") = sAppText Then
							Fn_MSO_SignoffTaskAndReviewersDecisions_Ops = True
							If dicSTRDDetails("Button")<>"" Then
								Call Fn_UI_WinButton_Click("Fn_MSO_SignoffTaskAndReviewersDecisions_Ops",objSTRDecisions.Dialog("ChangeUserSetting"),dicSTRDDetails("Button"),5,5,micLeftBtn)
							Else
								Call Fn_UI_WinButton_Click("Fn_MSO_SignoffTaskAndReviewersDecisions_Ops",objSTRDecisions.Dialog("ChangeUserSetting"),"Yes",5,5,micLeftBtn)
							End If
						Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Error message not verified.")
							Set objSTRDecisions = Nothing
							Set objSTRDList = Nothing
							Exit Function
						End If
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Error dialog does not exist.")
						Set objSTRDecisions = Nothing
						Set objSTRDList = Nothing
						Exit Function
					End If
					Wait 1
				End If
			End If
	End Select
	
	If sButton<>"" Then
		bFlag = Fn_MSO_WpfButton_Click("Fn_MSO_SignoffTaskAndReviewersDecisions_Ops","Click",objSTRDecisions,sButton)
		If bFlag = False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Click on WPFButton ["+sButton+"].")
			Fn_MSO_SignoffTaskAndReviewersDecisions_Ops = False
			Set objSTRDecisions = Nothing
			Set objSTRDList = Nothing
			Exit Function
		End If
		'Wait 1
	End If
	
	Set objSTRDecisions = Nothing
	Set objSTRDList = Nothing
End Function

''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''''/$$$$
''''/$$$$  FUNCTION NAME   :  Fn_MSO_Delegate_Ops
''''/$$$$
''''/$$$$  DESCRIPTION     	:  Function is used to perform operations on Delegate dialog
''''/$$$$ 
''''/$$$$  PRE-REQUISITES  	:  Delegate dialog should be opened
''''/$$$$
''''/$$$$  PARAMETERS   	:  sAction 		: 	Action to be performed
''''/$$$$ 					   sAppType 		: 	MSExcel, MSWord or MSPowerPoint
''''/$$$$ 					   dicDDetails 		: 	Dictionary object
''''/$$$$ 					   sButton 		: 	Button to be clicked
''''/$$$$ 					   sReserve 		: 	For future use
''''/$$$$	
''''/$$$$	Return Value 		:  True or False
''''/$$$$										
''''/$$$$	How To Use 		:  Set dicDDetails = CreateObject("Scripting.Dictionary")
''''/$$$$						  dicDDetails("EnterSearchString") 	= "SearchForGroup:Engineering"
''''/$$$$						  dicDDetails("SelectOrgTreeNode") 	= "Engineering:Designer:AutoTest1 (autotest1)"
''''/$$$$					   bReturn = Fn_MSO_Delegate_Ops("SelectDelegateUser","MSExcel",dicDDetails,"Submit","")
''''/$$$$	
''''/$$$$	HISTORY         :  
''''/$$$$				Developer Name		Date		Version	Changes								Reviewer
''''/$$$$	Created by  :	Vivek Ahirrao	 	20/06/2016	  1.0		Created								[TC1123-20160504-20_06_2016-VivekA-NewDevelopment]
''''/$$$$  
''''/$$$$	Modified by : 
''''/$$$$	
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Public Function Fn_MSO_Delegate_Ops(sAction,sAppType,dicDDetails,sButton,sReserve)
	GBL_FAILED_FUNCTION_NAME="Fn_MSO_Delegate_Ops"
	Dim objDelegate
	Dim dicCount, dicItems, dicKeys, aProperty
	Dim iCounter, iCount
	Dim sSubAction, sProperty, bFlag, sNode, sSrchBtnName
	
	Fn_MSO_Delegate_Ops = False
	On Error Resume Next

	Select Case sAppType
		Case "MSExcel"
			Set objDelegate = WpfWindow("Delegate")
		Case "MSWord"
			Set objDelegate = WpfWindow("Delegate")
		Case "MSPowerPoint"
			Set objDelegate = WpfWindow("Delegate")
	End Select
	
	If objDelegate.Exist(5) = False Then
		Set objDelegate = Nothing
		Exit Function
	End If
	
	Select Case sAction
		Case "ExistDelegateDialog"
			If objDelegate.Exist(5) = True Then
				Set objDelegate = Nothing
				Fn_MSO_Delegate_Ops = True
				Exit Function
			End If
		Case "SelectDelegateUser"
			dicCount = dicDDetails.Count
			dicItems = dicDDetails.Items
			dicKeys = dicDDetails.Keys
			
			For iCounter = 0 To dicCount - 1
				If Instr(dicKeys(iCounter),"SelectOrgTreeNode")>0 Then
					sSubAction = "SelectOrgTreeNode"
				ElseIf Instr(dicKeys(iCounter),"VerifyOrgTreeNode")>0 Then
					sSubAction = "VerifyOrgTreeNode"
				ElseIf Instr(dicKeys(iCounter),"EnterSearchString")>0 Then
					sSubAction = "EnterSearchString"
				Else
					sSubAction = dicKeys(iCounter)
				End If
				sProperty = dicItems(iCounter)
				bFlag = False
				Select Case sSubAction
					'"SearchForGroup:Engineering"
					Case "EnterSearchString"
						If sProperty<>"" Then
							'Click on Refresh button
							bFlag = Fn_MSO_WpfButton_Click("Fn_MSO_Delegate_Ops","Click",objDelegate,"RefreshOrgChart")
							If bFlag = False Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Click on [RefreshOrgChart] WPFButton.")
								Set objDelegate = Nothing
								Exit Function
							End If
							
							aProperty = Split(sProperty,":")
							Select Case aProperty(0)
								Case "SearchForGroup"
									sSrchBtnName = "SearchForGroup"
								Case "SearchForRole"
									sSrchBtnName = "SearchForRole"
								Case "SearchForUser"
									sSrchBtnName = "SearchForUser"
							End Select
							'Enter Search String
							objDelegate.WpfEdit("EnterSearchString").Set aProperty(1)
							Wait 1
							'Click on Group, Role or User button to search
							bFlag = Fn_MSO_WpfButton_Click("Fn_MSO_Delegate_Ops","Click",objDelegate,sSrchBtnName)
							If bFlag = False Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Click on ["+sSrchBtnName+"] WPFButton.")
								Set objDelegate = Nothing
								Exit Function
							End If
							If aProperty(0) = "SearchForUser" Then
								If objDelegate.Dialog("SearchforUser").Exist(5) Then
									objDelegate.Dialog("SearchforUser").WinButton("Yes").Click 5,5,micLeftBtn
									If Err.Number<0 Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Click on [Yes] WinButton in Search for User dialog.")
										Set objDelegate = Nothing
										Exit Function
									End If
								End If
							End If
							Wait 5
						End If
					'"Engineering:Designer:AutoTest1 (autotest1)"
					Case "SelectOrgTreeNode","VerifyOrgTreeNode"
						If sProperty<>"" Then					
							aProperty = Split(sProperty,":")
							'To Expand node
							For iCount = 0 To UBound(aProperty)
								If iCount = 0 Then
									sNode = aProperty(iCount)
								Else
									sNode = sNode + ":" + aProperty(iCount)
								End If
									
								If iCount <> UBound(aProperty) Then
									bFlag = Fn_MSO_NavTreeInWindow_Operations("Expand",sAppType,"SelectDelegateUser",sNode,"","","")
									If bFlag = False Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Expand ["+sNode+"].")
										Set objDelegate = Nothing
										Exit Function
									End If
									Wait 2
								End If
							Next
							
							If sSubAction = "SelectOrgTreeNode" Then
								'To Select node
								bFlag = Fn_MSO_NavTreeInWindow_Operations("Select",sAppType,"SelectDelegateUser",sProperty,"","","")
							ElseIf sSubAction = "VerifyOrgTreeNode" Then
								'To Verify node
								bFlag = Fn_MSO_NavTreeInWindow_Operations("Exist",sAppType,"SelectDelegateUser",sProperty,"","","")
							End If
							If bFlag = False Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to ["+sSubAction+"] node ["+sProperty+"].")
								Set objDelegate = Nothing
								Exit Function
							End If
							Wait 1
						End If
				End Select
				
				If bFlag = False Then
					Fn_MSO_Delegate_Ops = False
					Set objDelegate = Nothing
					Exit Function
				End If
			Next
			
			'Click on button provided
			If sButton<>"" Then
				bFlag = Fn_MSO_WpfButton_Click("Fn_MSO_Delegate_Ops","Click",objDelegate,sButton)
				If bFlag = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Click on WPFButton ["+sButton+"].")
					Fn_MSO_Delegate_Ops = False
					Set objDelegate = Nothing
					Exit Function
				End If
				Wait 1
			End If
			
			Fn_MSO_Delegate_Ops = True
	End Select
	Set objDelegate = Nothing
End Function

''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''''/$$$$
''''/$$$$  FUNCTION NAME   	:  Fn_MSO_Revise_Ops
''''/$$$$
''''/$$$$  DESCRIPTION     	:  Function is used to perform operations on Revise dialog
''''/$$$$ 
''''/$$$$  PRE-REQUISITES  	:  Delegate dialog should be opened
''''/$$$$
''''/$$$$  PARAMETERS   	:  sAction 		: 	Action to be performed
''''/$$$$ 					   sAppType 	: 	MSExcel, MSWord or MSPowerPoint
''''/$$$$ 					   sNode 		: 	Node name path in Folder View Tree
''''/$$$$ 					   dicRevise 	: 	Dictionary object
''''/$$$$ 					   sButton 		: 	Button to be clicked
''''/$$$$ 					   sReserve 	: 	For future use
''''/$$$$	
''''/$$$$	Return Value 	:  True or False
''''/$$$$										
''''/$$$$	How To Use 		:  sNode = "Home:000387-VivekA:000387/A;1-VivekA"
''''/$$$$					   Set dicRevise = CreateObject("Scripting.Dictionary")
''''/$$$$					   	   dicRevise("ItemRevision") = "VivekA"
''''/$$$$					   	   dicRevise("Name") = "VivekARevise"
''''/$$$$					   	   dicRevise("Description") = "Revise Done"
''''/$$$$					   bReturn = Fn_MSO_Revise_Ops("Revise","MSExcel",sNode,dicRevise,"Cancel","")
''''/$$$$	
''''/$$$$	HISTORY         :  
''''/$$$$					Developer Name		Date		Version		Changes								Reviewer
''''/$$$$	Created by  :	Vivek Ahirrao	 	28/06/2016	  1.0		Created								[TC1123-20160504-28_06_2016-VivekA-NewDevelopment]
''''/$$$$  
''''/$$$$	Modified by : 
''''/$$$$	
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Public Function Fn_MSO_Revise_Ops(sAction,sAppType,sNode,dicRevise,sButton,sReserve)
	GBL_FAILED_FUNCTION_NAME="Fn_MSO_Revise_Ops"
	Dim objRevise, dicCount, dicItems, dicKeys
	Dim bFlag, iCounter, sSubAction, sProperty, sPropertyName
	
	Fn_MSO_Revise_Ops = False
	On Error Resume Next
	
	Select Case sAppType
		Case "MSExcel"
			Set objRevise = WpfWindow("ReviseItemRevision")
		Case "MSWord"
			Set objRevise = WpfWindow("ReviseItemRevision")
		Case "MSPowerPoint"
			Set objRevise = WpfWindow("ReviseItemRevision")
	End Select
	
	If objRevise.Exist(5) = False Then
		If sNode<>"" Then
			bFlag = Fn_MSO_FolderViewTreeOperations(sAppType,"PopupMenuSelect",sNode,"Revise...","","","")
			If bFlag = False Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Revise Window does not Exist.")
				Set objRevise = Nothing
				Exit Function
			End If
		End If
		If objRevise.Exist(5) = False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Revise Window does not Exist.")
			Set objRevise = Nothing
			Exit Function
		End If
	End If
	
	Select Case sAction
		Case "Revise"
			dicCount = dicRevise.Count
			dicItems = dicRevise.Items
			dicKeys = dicRevise.Keys
			
			For iCounter = 0 To dicCount - 1
				sSubAction = dicKeys(iCounter)
				sProperty = dicItems(iCounter)
				bFlag = False
				Select Case sSubAction
					'Set Item Revision
					Case "ItemRevision"
						If sProperty<>"" Then
							'Set text property 
							objRevise.WpfObject("PropertyName").SetTOProperty "text","Item Revision"
							Wait 0,100
							If objRevise.WpfObject("PropertyName").Exist Then
								objRevise.WpfComboBox("PropertyComboBox").Click 5,5,micLeftBtn
								Wait 1
								Err.Clear
								objRevise.WpfComboBox("PropertyComboBox").Type sProperty
								Wait 1
								Err.Clear
								objRevise.WpfComboBox("PropertyComboBox").Select sProperty
								If Err.Number < 0 Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Select ["+sProperty+"] from [Item Revision] WPF Combo Box.")
									Set objRevise = Nothing
									Exit Function
								End If
								Wait 1
								bFlag = True
							End If
						End If
					'"File Name:Book1"
					Case "Name","Description"
						If sProperty<>"" Then
							'Set text property
							If sSubAction = "Name" Then
								sPropertyName = "Name"
							ElseIf sSubAction = "Description" Then
								sPropertyName = "Description"
							End If
							objRevise.WpfObject("PropertyName").SetTOProperty "text",sPropertyName & ":"
							Wait 0,100
							If objRevise.WpfObject("PropertyName").Exist Then
								objRevise.WpfEdit("PropertyEditBox").Set sProperty
								If Err.Number < 0 Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Set ["+sProperty+"] in ["+sPropertyName+"] WPF Edit Box.")
									Set objRevise = Nothing
									Exit Function
								End If
								Wait 1
								bFlag = True
							End If
						End If
				End Select
				
				If bFlag = False Then
					Fn_MSO_Revise_Ops = False
					Set objRevise = Nothing
					Exit Function
				End If
			Next
			
			'Click on button provided
			If sButton<>"" Then
				bFlag = Fn_MSO_WpfButton_Click("Fn_MSO_AdvancedSearchOps","Click",objRevise,sButton)
				If bFlag = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Click on WPFButton ["+sButton+"].")
					Fn_MSO_Revise_Ops = False
					Set objRevise = Nothing
					Exit Function
				End If
				Wait 1
			End If
			Fn_MSO_Revise_Ops = True
	End Select
	Set objRevise = Nothing
End Function

''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''''/$$$$
''''/$$$$  FUNCTION NAME   	:  Fn_MSO_TeamcenterOpen_Operations
''''/$$$$
''''/$$$$  DESCRIPTION     	:  Function is used to perform operations on Teamcenter Open dialog
''''/$$$$ 
''''/$$$$  PRE-REQUISITES  	:  Teamcenter Open dialog should be opened or Open button should be present
''''/$$$$
''''/$$$$  PARAMETERS   	:  sAction 			: 	Action to be performed
''''/$$$$ 					   sAppType 		: 	MSExcel, MSWord or MSPowerPoint
''''/$$$$ 					   dicOpenDetails 	: 	Dictionary object
''''/$$$$ 					   sButton 			: 	Button to be clicked
''''/$$$$ 					   sReserve 		: 	For future use
''''/$$$$	
''''/$$$$	Return Value 	:  True or False
''''/$$$$										
''''/$$$$	How To Use 		:  Set dicOpenDetails = CreateObject("Scripting.Dictionary")
''''/$$$$						   dicOpenDetails("SelectNode") 		= "Home:AutomatedTests"
''''/$$$$						   dicOpenDetails("VerifyFileTable") 	= "DatasetName:MsExcelDt89"
''''/$$$$						   dicOpenDetails("SelectFileTable") 	= "DatasetName:MsExcelDt89"
''''/$$$$						   dicOpenDetails("EditBox1") 			= "Dataset Name:DatasetOpen"
''''/$$$$						   dicOpenDetails("ComboBox1") 			= "File of type:Microsoft Excel Worksheet (*.xlsx)"
''''/$$$$						   dicOpenDetails("VerifyComboboxItems")= "File of type:Microsoft Excel Worksheet (*.xlsx)~Microsoft Excel Macro-Enabled Worksheet (*.xlsm)"
''''/$$$$						   dicOpenDetails("VerifyEditBoxValue") = "Dataset Name:MsExcelDt89"
''''/$$$$						   dicOpenDetails("SelectOpenMenu") 	= "Open:Open File (Read-Only)"
''''/$$$$					   bReturn = Fn_MSO_TeamcenterOpen_Operations("Open","MSExcel",dicOpenDetails,"","")
''''/$$$$					   
''''/$$$$					   	   dicOpenDetails("DatasetName") 	= "Word11"
''''/$$$$					   	   dicOpenDetails("FileName") 		= "Word11.docx"
''''/$$$$					   bReturn = Fn_MSO_TeamcenterOpen_Operations("VerifyFileTypeTable","MSWord",dicOpenDetails,"","")
''''/$$$$	
''''/$$$$	HISTORY         :  
''''/$$$$				Developer Name		Date		Version	Changes								Reviewer
''''/$$$$	Created by  :	Vivek Ahirrao	 	29/06/2016	  1.0		Created								[TC1123-20160504-29_06_2016-VivekA-NewDevelopment]
''''/$$$$  
''''/$$$$	Modified by : Vivek Ahirrao	 	07/07/2016	  1.0		Added Case "VerifyFileTypeTable"
''''/$$$$	
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Public Function Fn_MSO_TeamcenterOpen_Operations(sAction,sAppType,dicOpenDetails,sButton,sReserve)
	GBL_FAILED_FUNCTION_NAME="Fn_MSO_TeamcenterOpen_Operations"
	Dim objTCOpen, objMenu
	Dim iCount, iCounter, iTotalRow, iRow
	Dim bFlag, sNode, sSubAction, sProperty, sAppText, sBtnOpen, sMenu, sProperties, sValues
	Dim aStrNode, dicCount, dicItems, dicKeys, aProperty
	
	Fn_MSO_TeamcenterOpen_Operations = False
	On Error Resume Next
	
	Set objTCOpen = WpfWindow("TeamcenterOpen")
	If objTCOpen.Exist(5) = False Then
		bFlag = Fn_MSO_RibbonButton_Operations(sAppType,"Click","Open","")
		If bFlag = False Then
			Set objTCOpen = Nothing
			Exit Function
		End If
	End If
	If objTCOpen.Exist(5) = False Then
		Set objTCOpen = Nothing
		Exit Function
	End If
	
	Select Case sAction
		'Case to Verify File type Table column values
		Case "VerifyFileTypeTable","GetRowNumber"
			'Get number of rows in Table
			iTotalRow = objTCOpen.WpfList("FileTypeTable").Object.Items.Count
			For iRow = 0 To iTotalRow-1
				iCount = 0
				iCounter = 0
				If dicOpenDetails("DatasetName")<>"" Then
					iCount = iCount+1
					sAppText = objTCOpen.WpfList("FileTypeTable").Object.Items.Item(iRow).Dataset.DisplayName
					If sAppText = dicOpenDetails("DatasetName") Then
						iCounter = iCounter+1
					End If
				End If
				If dicOpenDetails("FileName")<>"" Then
					iCount = iCount+1
					sAppText = objTCOpen.WpfList("FileTypeTable").Object.Items.Item(iRow).NamedReference
					If sAppText = dicOpenDetails("FileName") Then
						iCounter = iCounter+1
					End If
				End If
				'Type column
				'msgbox WpfWindow("Teamcenter Open").WpfList("FileTypeTable").Object.Items.Item(0).Dataset.MetaTypeInfo.Name
				If iCount = iCounter Then
					If sAction = "VerifyFileTypeTable" Then
						Fn_MSO_TeamcenterOpen_Operations = True
					ElseIf sAction = "GetRowNumber" Then
						Fn_MSO_TeamcenterOpen_Operations = iRow
					End If
					Exit For
				End If
			Next
			If Cint(iRow) = Cint(iTotalRow) Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Fail to verify Row in File Type Table")
				If sAction = "VerifyFileTypeTable" Then
					Fn_MSO_TeamcenterOpen_Operations = False
				ElseIf sAction = "GetRowNumber" Then
					Fn_MSO_TeamcenterOpen_Operations = -1
				End If
				Set objTCOpen = Nothing
				Exit Function
			End If
			
		'Case to perform Open operation
		Case "Open"
			dicCount = dicOpenDetails.Count
			dicItems = dicOpenDetails.Items
			dicKeys = dicOpenDetails.Keys
			
			For iCounter = 0 To dicCount - 1
				If Instr(dicKeys(iCounter),"VerifyEditBoxValue")>0 Then
					sSubAction = "VerifyEditBoxValue"
				ElseIf Instr(dicKeys(iCounter),"EditBox")>0 Then
					sSubAction = "EditBox"
				ElseIf Instr(dicKeys(iCounter),"ComboBox")>0 Then
					sSubAction = "ComboBox"
				Else
					sSubAction = dicKeys(iCounter)
				End If
				sProperty = dicItems(iCounter)
				bFlag = False
				Select Case sSubAction
					'Select node in Tree on left side of dialog
					Case "SelectNode"
						If sProperty<>"" Then
							aStrNode = Split(sProperty,":")
							For iCount = 0 To UBound(aStrNode)
								If iCount = 0 Then
									sNode = aStrNode(iCount)
								Else
									sNode = sNode + ":" + aStrNode(iCount)
								End If
								
								bFlag = Fn_MSO_NavTreeInWindow_Operations("Select",sAppType,"TeamcenterOpenDialog",sNode,"","","")
								If bFlag = False Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Select ["+sNode+"] in Tree in TeamcenterOpenDialog.")
									Set objTCOpen = Nothing
									Exit Function
								End If
								Wait 1
								If iCount <> UBound(aStrNode) Then
									bFlag = Fn_MSO_NavTreeInWindow_Operations("Expand",sAppType,"TeamcenterOpenDialog",sNode,"","","")
									If bFlag = False Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Expand ["+sNode+"] in Tree in TeamcenterOpenDialog.")
										Set objTCOpen = Nothing
										Exit Function
									End If
									Wait 2
								End If
							Next
							bFlag = True
						End If
					'"Dataset Name:DatasetOpen"
					Case "EditBox"
						If sProperty<>"" Then
							aProperty = Split(sProperty,":")
							'Set text property 
							objTCOpen.WpfObject("PropertyName").SetTOProperty "text",aProperty(0) & ":"
							If objTCOpen.WpfObject("PropertyName").Exist Then
								objTCOpen.WpfEdit("PropertyValue").Set aProperty(1)
								If Err.Number < 0 Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Set ["+aProperty(1)+"] in ["+aProperty(0)+"] WPF Edit Box.")
									Set objTCOpen = Nothing
									Exit Function
								End If
								Wait 1
								bFlag = True
							End If
						End If
					'"Dataset Name:DatasetOpen"
					Case "VerifyEditBoxValue"
						If sProperty<>"" Then
							aProperty = Split(sProperty,":")
							'Set text property 
							objTCOpen.WpfObject("PropertyName").SetTOProperty "text",aProperty(0) & ":"
							If objTCOpen.WpfObject("PropertyName").Exist Then
								sAppText = WpfWindow("TeamcenterOpen").WpfEdit("PropertyValue").GetROProperty("value")
								If aProperty(1)<>sAppText Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Verify value ["+aProperty(1)+"] in ["+aProperty(0)+"] WPF Edit Box.")
									Set objTCOpen = Nothing
									Exit Function
								End If
								Wait 1
								bFlag = True
							End If
						End If
					'"File of type:Item Revision"
					Case "ComboBox"
						If sProperty<>"" Then
							aProperty = Split(sProperty,":")
							'Set text property 
							objTCOpen.WpfObject("PropertyName").SetTOProperty "text",aProperty(0) & ":"
							Wait 1
							If objTCOpen.WpfObject("PropertyName").Exist Then
								objTCOpen.WpfComboBox("PropertyComboBox").Click 5,5,micLeftBtn
								Wait 1
								Err.Clear
								objTCOpen.WpfComboBox("PropertyComboBox").Type aProperty(1)
								Wait 1
								Err.Clear
								objTCOpen.WpfComboBox("PropertyComboBox").Select aProperty(1)
								If Err.Number < 0 Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Select ["+aProperty(1)+"] from ["+aProperty(0)+"] WPF Combo Box.")
									Set objTCOpen = Nothing
									Exit Function
								End If
								Wait 1
								bFlag = True
							End If
						End If
					'Case to verify Items in "File of type:" combobox list
					'"File of type:Microsoft Excel Worksheet (*.xlsx)~Microsoft Excel Macro-Enabled Worksheet (*.xlsm)"
					Case "VerifyComboboxItems"
						If sProperty<>"" Then
							aProperty = Split(sProperty,":")
							aItems = Split(aProperty(1),"~")
							'Set text property
							objTCOpen.WpfObject("PropertyName").SetTOProperty "text",aProperty(0) & ":"
							Wait 1
							If objTCOpen.WpfObject("PropertyName").Exist Then
								aAppItems = Split(objTCOpen.WpfComboBox("PropertyComboBox").GetContent(),vblf)
								For iCount = 0 To UBound(aItems)
									bFlag = False
									For iCount1 = 0 To UBound(aAppItems)
										If Trim(aAppItems(iCount1))=Trim(aItems(iCount)) Then
											bFlag = True
											Exit For
										End If
									Next
									If bFlag = False Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : File Type Item ["+aItems(iCount)+"] does not Exist in WPF Combobox.")
										Set objTCOpen = Nothing
										Exit Function
									End If
								Next
								Wait 1
								bFlag = True
							End If
						End If
					'Case to Select File Row from File Type Table
					'"DatasetName:MsExcelDt89~FileName:Word11.docx"
					Case "SelectFileTable"
						If sProperty<>"" Then
							aProperty = Split(sProperty,"~")
							'Get Row Number
							Set objDic = CreateObject("Scripting.Dictionary")
							For iCount = 0 To UBound(aProperty)
								aProp = Split(aProperty(iCount),":")
								objDic.Add aProp(0),aProp(1)
							Next
							iRowNumber = Fn_MSO_TeamcenterOpen_Operations("GetRowNumber","MSWord",objDic,"","")
							If iRowNumber = -1 Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Dataset ["+sProperty+"] does not Exist in ListTable.")
								Set objTCOpen = Nothing
								Exit Function
							End If

							'Select iRow
							objTCOpen.WpfList("FileTypeTable").object.SelectedIndex = iRowNumber
							If Err.Number<0 Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Select Dataset ["+sProperty+"] WPFList Table.")
								Set objTCOpen = Nothing
								Exit Function
							End If
							Wait 1
							bFlag = True
						End If
					'Case to Verify Dataset File Row from File Type Table
					'"DatasetName:MsExcelDt89"
					Case "VerifyFileTable"
						If sProperty<>"" Then
							aProperty = Split(sProperty,":")
							iTotalRow = objTCOpen.WpfList("FileTypeTable").Object.Items.Count
							bFlag = False
							For iRow = 0 To iTotalRow-1
								If aProperty(0) = "DatasetName" Then
									sAppText = objTCOpen.WpfList("FileTypeTable").Object.Items.Item(iRow).Dataset.DisplayName
								End If
								If sAppText = aProperty(1) Then
									bFlag = True
									Exit For
								End If
							Next
							If bFlag = False Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Dataset ["+aProperty(0)+" : "+aProperty(1)+"] does not Exist in ListTable.")
								Set objTCOpen = Nothing
								Exit Function
							End If
							Wait 1
							bFlag = True
						End If
					'"Open:Open and Check-Out File"
					'"Open:Open File (Read Only)"
					Case "SelectOpenMenu"
						If sProperty<>"" Then
							If Instr(sProperty,":")>0 Then
								aProperty = Split(sProperty,":")
								sBtnOpen = "OpenDropDown"
								sMenu = aProperty(1)
							Else
								sBtnOpen = "Open"
								sMenu = ""
							End If
							'Click on WPF button
							bFlag = Fn_MSO_WpfButton_Click("Fn_MSO_TeamcenterOpen_Operations","Click",objTCOpen,sBtnOpen)
							If bFlag = False Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Click on WPFButton ["+sButton+"].")
								Set objTCOpen = Nothing
								Exit Function
							End If
							Wait 1
						 	'Click on Menu
						 	If sMenu<>"" Then
						 		sProperties ="Class Name~classname" 
								sValues = "ContextMenu~System.Windows.Controls.ContextMenu"
								Set objMenu = Fn_SISW_UI_Object_GetChildObjects("Fn_MSO_TeamcenterOpen_Operations",objTCOpen,sProperties,sValues)
								For iCount = 0 To objMenu.count-1 
									If Instr(objMenu(iCount).toString(), "WpfMenu") > 0 Then
										'sMenu = Replace(sMenu,":",";")
										objMenu(iCount).Select sMenu
										Exit For
									End If	
								Next
							 	If Err.Number<0 Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to Perform Menu operation ["+sMenu+"].")
									Set objTCOpen = Nothing
									Set objMenu = Nothing
									Exit Function
								End If
								Wait 1
								Set objMenu = Nothing
						 	End If
						 	Wait 1
						 	bFlag = True
						End If
				End Select
				
				If bFlag = False Then
					Fn_MSO_TeamcenterOpen_Operations = False
					Set objTCOpen = Nothing
					Exit Function
				End If
			Next
			
			'Click on button provided
			If sButton<>"" Then
				bFlag = Fn_MSO_WpfButton_Click("Fn_MSO_TeamcenterOpen_Operations","Click",objTCOpen,sButton)
				If bFlag = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Click on WPFButton ["+sButton+"].")
					Fn_MSO_TeamcenterOpen_Operations = False
					Set objTCOpen = Nothing
					Exit Function
				End If
				Wait 1
			End If
			Fn_MSO_TeamcenterOpen_Operations = True
		Case Else
			Fn_MSO_TeamcenterOpen_Operations = False
	End Select
	
	Set objTCOpen = Nothing
End Function

''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''''/$$$$
''''/$$$$  FUNCTION NAME  	:  Fn_MSO_ErrorDialog_Ops
''''/$$$$
''''/$$$$  DESCRIPTION      	:  Function is used to Perform operations on any Warning dialog or error dialog.
''''/$$$$ 
''''/$$$$	Return Value 		:  True/False
''''/$$$$	
''''/$$$$  PARAMETERS   	:  sAction 	: 	Action to be performed
''''/$$$$ 					   sAppType 	: 	MSExcel, MSWord or MSPowerPoint
''''/$$$$ 					   dicErrorInfo : 	Dic details
''''/$$$$ 					   sButton 	: 	Button Name
''''/$$$$ 					   sReserve 	: 	For future use
''''/$$$$ 
''''/$$$$	How To Use 		:  Set dicErrorInfo = CreateObject("Scripting.Dictionary")
''''/$$$$						  dicErrorInfo("ErrorMessage") 	= "The file 'MSExcel.xlsx' is already open in Edit mode and has been checked-out."
''''/$$$$					   bReturn = Fn_MSO_ErrorDialog_Ops("VerifyError","MSExcel",dicErrorInfo,"OK","")
''''/$$$$	
''''/$$$$	HISTORY         :  
''''/$$$$				Developer Name		Date	 	Version	Changes							Reviewer
''''/$$$$	Created by  :	 Vivek Ahirrao		08/07/2016	  1.0		Created							[TC1123-20160504-08_07_2016-VivekA-NewDevelopment]
''''/$$$$  												Added for RM - Office Client new TC's Development
''''/$$$$	Modified by : 
''''/$$$$	
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Public Function Fn_MSO_ErrorDialog_Ops(sAction,sAppType,dicErrorInfo,sButton,sReserve)
	GBL_FAILED_FUNCTION_NAME="Fn_MSO_ErrorDialog_Ops"
	Dim objAppType, objError
	Dim sAppText, bFlag
	
	Fn_MSO_ErrorDialog_Ops = False
	
	Select Case sAppType
		Case "MSExcel"
			Set objAppType = Window("MicrosoftExcel")
		Case "MSWord"
			Set objAppType = Window("MicrosoftWord")
		Case "MSPowerPoint"
			Set objAppType = Window("MicrosoftPowerPoint")
	End Select

	'Set title if available
	If dicErrorInfo("Title")<>"" Then
		objAppType.Dialog("DialogInformation").SetTOProperty "text",dicErrorInfo("Title")
	End If
	
	If Dialog("ConfirmationBox").Exist Then
		Set objError = Dialog("ConfirmationBox")
	ElseIf objAppType.Dialog("DialogInformation").Exist Then
		Set objError = objAppType.Dialog("DialogInformation")
	ElseIf WpfWindow("PossibleIssues").Exist(3) Then
		Set objError = WpfWindow("PossibleIssues")
	End If
	
	Select Case sAction
		Case "VerifyError","VerifyPossibleIssuesError"
			'Verify error message
			If dicErrorInfo("ErrorMessage")<>"" Then
				If sAction = "VerifyPossibleIssuesError" Then
					sAppText = WpfWindow("PossibleIssues").WpfList("PartialErrorsList").GetVisibleText()
					If Instr(sAppText,dicErrorInfo("ErrorMessage")) <= 0 Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Error message does not match.")
						Set objError = Nothing
						Fn_MSO_ErrorDialog_Ops = False
						Exit Function
					End If
				Else
					sAppText = objError.Static("TextMessage").GetROProperty("text")
					If sAppText <> dicErrorInfo("ErrorMessage") Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Error message does not match.")
						Set objError = Nothing
						Fn_MSO_ErrorDialog_Ops = False
						Exit Function
					End If
				End If
			End If
			
			'Click on button
			If sButton<>"" Then
				If sAction = "VerifyPossibleIssuesError" Then
					bFlag = Fn_MSO_WpfButton_Click("Fn_MSO_ErrorDialog_Ops","Click",objError,sButton)
					wait 2
				Else
					bFlag = Fn_UI_WinButton_Click("Fn_MSO_ErrorDialog_Ops",objError,sButton,5,5,micLeftBtn)
				End If
				
			Else
				bFlag = Fn_UI_WinButton_Click("Fn_MSO_ErrorDialog_Ops",objError,"OK",5,5,micLeftBtn)
			End If
			Fn_MSO_ErrorDialog_Ops = True
	End Select
	Set objError = Nothing
End Function

''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''''/$$$$
''''/$$$$  Function Name 	:  Fn_MSO_BrowseTreeOperations
''''/$$$$
''''/$$$$  Description       	:  Function is used to perform operations on browse view
''''/$$$$ 
''''/$$$$  Pre-Req  		  	:  Office Client Window Should be in Focus
''''/$$$$
''''/$$$$  Parameters   	  	:  sAction 		: 	Action name
''''/$$$$ 				   	   sAppType 		: 	MSExcel, MSWord, MSPowerPoint
''''/$$$$ 					   sNode 			: 	Node Path
''''/$$$$ 					   sPopUpMenu 	: 	Menu
''''/$$$$ 					   sReserve1 		: 	Future use
''''/$$$$ 					   sReserve2 		: 	Future use
''''/$$$$	
''''/$$$$	Return Value 		:  True or False
''''/$$$$										
''''/$$$$	Examples 		:  bReturn = Fn_MSO_BrowseTreeOperations("SelectFromStartTrail","MSExcel","Start","","","")
''''/$$$$					   bReturn = Fn_MSO_BrowseTreeOperations("Select","MSExcel","Home","","","")
''''/$$$$					   bReturn = Fn_MSO_BrowseTreeOperations("Exist","MSExcel","AutomatedTests","","","")
''''/$$$$					   bReturn = Fn_MSO_BrowseTreeOperations("PopUpMenuSelect","MSExcel","000051-RootItem","Check-In/Out...:Check-In","","")
''''/$$$$	
''''/$$$$	History         :  
''''/$$$$				Developer Name	     Date			Version	Changes								Reviewer
''''/$$$$	Created by  :	 Vivek Ahirrao	 	11/07/2016	  	  1.0		Created								[TC1123-20160504-11_07_2016-VivekA-NewDevelopment]
''''/$$$$  
''''/$$$$	Modified by : 	 
''''/$$$$	
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Public Function Fn_MSO_BrowseTreeOperations(sAction, sAppType, sNode, sPopUpMenu, sReserve1, sReserve2)
	GBL_FAILED_FUNCTION_NAME="Fn_MSO_BrowseTreeOperations"
	Dim objAppType, objBrowseWindow, objDesc, objChilds, objMenu
	Dim bFlag, iCount, sAppText, sProperties, sValues
	
	Fn_MSO_BrowseTreeOperations = False
	On Error Resume Next
	
	Select Case sAppType
		Case "MSExcel"
			Set objAppType = Window("MicrosoftExcel")
			Set objBrowseWindow = Window("MicrosoftExcel").WinObject("BrowseView").SwfObject("SwfObject").WpfWindow("WpfWindow")
			objAppType.Maximize
		Case "MSWord"
			Set objAppType = Window("MicrosoftWord")
			Set objBrowseWindow = Window("MicrosoftWord").WinObject("BrowseView").SwfObject("SwfObject").WpfWindow("WpfWindow")
			objAppType.Maximize
		Case "MSPowerPoint"
			'Future Use
	End Select
	
	If Fn_UI_ObjectExist("Fn_MSO_BrowseTreeOperations",objAppType.WinObject("BrowseView"))=False Then
		bFlag = Fn_MSO_RibbonButton_Operations(sAppType,"Click","NavigateDropDown:Browse","")
		If bFlag = False Then
			Set objBrowseWindow = Nothing
			Exit Function
		End If
		Wait 5
	End If
	Set objAppType = Nothing
	
	Select Case sAction
		Case "SelectFromStartTrail"
			If sNode<>"" Then
				bFlag = False
				Set objDesc = Description.Create
				objDesc("wpftypename").value = "button"
				Set objChilds = objBrowseWindow.WpfObject("StartTrail").ChildObjects(objDesc)
				For iCount = 0 To objChilds.count - 1
					sAppText = objChilds(iCount).getVisibleText()
					If sAppText = sNode Then
						bFlag = True
						Exit For
					End If
				Next
				If bFlag = True Then
					objChilds(iCount).Click 10,5,micLeftBtn
					If Err.Number<0 Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail : Failed to Select ["+sNode+"] from Start Trail Ribbon")
						Set objBrowseWindow = Nothing
						Exit Function
					End If
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail : Node ["+sNode+"] does not Exist in Start Trail Ribbon")
					Set objBrowseWindow = Nothing
					Exit Function
				End If
				Fn_MSO_BrowseTreeOperations = True
			End If
		Case "Select", "Exist"
			If sNode<>"" Then
				objBrowseWindow.WpfObject("WpfButtonObject").SetTOProperty "text",sNode
				If objBrowseWindow.WpfObject("WpfButtonObject").Exist(2) Then
					If sAction = "Select" Then
						objBrowseWindow.WpfObject("WpfButtonObject").Click 10,5,micLeftBtn
						If Err.Number<0 Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail : Failed to Select ["+sNode+"] from Browse Tree")
							Set objBrowseWindow = Nothing
							Exit Function
						End If
					End If
					Fn_MSO_BrowseTreeOperations = True
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail : Node ["+sNode+"] does not Exist in Browse Tree")
					Set objBrowseWindow = Nothing
					Exit Function
				End If
			End If
		Case "PopUpMenuSelect"
			If sNode<>"" Then
				objBrowseWindow.WpfObject("WpfButtonObject").SetTOProperty "text",sNode
				If objBrowseWindow.WpfObject("WpfButtonObject").Exist(2) Then
					objBrowseWindow.WpfObject("WpfButtonObject").Click 10,5,micRightBtn
					If Err.Number<0 Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail : Failed to Right Click ["+sNode+"] in Browse Tree")
						Set objBrowseWindow = Nothing
						Exit Function
					End If
					Wait 1
					sProperties ="Class Name~classname" 
					sValues = "ContextMenu~System.Windows.Controls.ContextMenu"
			
					Set objMenu = Fn_SISW_UI_Object_GetChildObjects("Fn_MSO_BrowseTreeOperations",objBrowseWindow,sProperties,sValues)
			
					For iCount = 0 To objMenu.count-1 
						If Instr(objMenu(iCount).toString(), "WpfMenu") > 0 Then
							sPopUpMenu = Replace(sPopUpMenu,":",";")
							objMenu(iCount).Select sPopUpMenu
							If Err.Number<0 Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail : Failed to Perform PopupMenu ["+sPopUpMenu+"] operation on node ["+sNode+"] in Browse Tree")
								Set objBrowseWindow = Nothing
								Exit Function
							End If
							Exit For
						End If	
					Next
					Fn_MSO_BrowseTreeOperations = True
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail : Node ["+sNode+"] does not Exist in Browse Tree")
					Set objBrowseWindow = Nothing
					Exit Function
				End If
			End If
	End Select
	Set objBrowseWindow = Nothing
End Function	
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''''/$$$$
''''/$$$$  Function Name 	:  Fn_MSO_SwitchToTCTab
''''/$$$$
''''/$$$$  Description       	:  Function is used to perform operations to switch to Teamcenter tab under MS Office application
''''/$$$$ 
''''/$$$$  Pre-Req  		  	:  Office Client Window Should be in Focus
''''/$$$$
''''/$$$$  Parameters   	  	:  sAction 		: 	MS Office Application name
''''/$$$$						e.g "Word" /"Excel" /"PowerPoint"
''''/$$$$							sTabname : MS Application Tab Name	Ex: "Teamcenter"/"AddIns"
''''/$$$$
''''/$$$$	Return Value 		:  True or False
''''/$$$$										
''''/$$$$					Developer Name	     Date		Version	Changes								Reviewer
''''/$$$$	Created by  :	 Dhananjay Niwal	28-12-2018		1.0	
''''/$$$$  
''''/$$$$	Modified by : 	 
''''/$$$$	
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Public Function Fn_MSO_SwitchToTCTab(sAppName, sTabname)
dim objMSApp, bReturn

If instr(1,lcase(sAppName),"excel") > 0 Then
	set objMSApp = Window("MicrosoftExcel").WinObject("Ribbon").WinTab("RibbonTabs")	
ElseIf instr(1,lcase(sAppName),"word") > 0 Then
	set objMSApp = Window("MicrosoftWord").WinObject("Ribbon").WinTab("RibbonTabs")	
End if 
If instr(1,lcase(sTabname),"teamcenter") > 0 Then
	sTabname = "Teamcenter"
ElseIf instr(1,lcase(sTabname),"addins") > 0 Then
	sTabname = "Add-Ins"
End if	
	
	objMSApp.select(sTabname)	
	err.clear
	If err.number > 0 Then
		bReturn = False
	End If
	
	bReturn =True

Fn_MSO_SwitchToTCTab = bReturn
set objMSApp = nothing
End function
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''''/$$$$
''''/$$$$  Function Name 	:  Fn_MSO_Word_FlatFileImport_NewImport
''''/$$$$
''''/$$$$  Description       	:  Function is used to import new flat file
''''/$$$$  Pre-Req  		  	:  Office Client Window Should be in Focus
''''/$$$$
''''/$$$$  Parameters   	  	:  CreateLink 		: 	True/False
''''/$$$$
''''/$$$$	Return Value 		:  True or False
''''/$$$$										
''''/$$$$					Developer Name	     Date		 Version	Changes		
''''/$$$$	Created by  :	 Amruta Patil	   14-03-2022		1.0	
''''/$$$$  
''''/$$$$	Modified by : 	 
''''/$$$$	
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Public Function Fn_MSO_Word_FlatFileImport_NewImport(CreateLink,button,Saveandcheckin,sinfo2)
dim ObjImport
Fn_MSO_Word_FlatFileImport_NewImport=False
Set ObjImport = WpfWindow("WordFlatFileImport")
Set wshshell = CreateObject("WScript.Shell")
							wshshell.SendKeys "%"
							wshshell.SendKeys "Y2"
                            wshshell.SendKeys "YF"
										   Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Successfully Activated the Teamcenter Tab")
										   Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Successfully Clicked on the Specifications Button")
										   Fn_MSO_Word_FlatFileImport_NewImport=true
									wait 3
									wshshell.SendKeys "{ENTER}"
								
								If not ObjImport.Exist(10) Then
									Fn_MSO_Word_FlatFileImport_NewImport = False
									Exit Function
								End If
							wait 5
						ObjImport.WpfCheckBox("Create Links").Set CreateLink
						If err.number<0 Then
								   Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to set checkbox of createlink to ["+CreateLink+"]")
								   Fn_MSO_Word_FlatFileImport_NewImport=False
								   Exit function
						End If
						wait 3
		
						 If button<>"" Then
							ObjImport.WpfButton(button).Click 5,5,micLeftBtn
								If err.number<0 Then
								   Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to click button["+button+"]")
								   Fn_MSO_Word_FlatFileImport_NewImport=False
								   Exit function
								End If
							 wait 3
						 End If
						 
						If WpfWindow("Save").Exist(5) Then
							If Saveandcheckin <> "" Then
								WpfWindow("Save").WpfCheckBox("SaveAndCheckIn").Set Saveandcheckin
							End If
							WpfWindow("Save").WpfButton("OK").Click 5,5,micLeftBtn
							If err.number<0 Then
								   Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to click button["+button+"]")
								   Fn_MSO_Word_FlatFileImport_NewImport=False
								   Exit function
							Else
								wait 30
								ObjImport.WpfButton("Close").Click 5,5,micLeftBtn
								Fn_MSO_Word_FlatFileImport_NewImport=True
								 Exit function
							End If
							
						Else
							Fn_MSO_Word_FlatFileImport_NewImport=False
							Exit function
						End If
End Function
'------------------------------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_MSO_UnZipFile

'Description	   	:	Function Used to Unzip any file

'Parameters			:	1.SourceFilePath - file of the file to be unzip
'						2.TargetExtractFolderName - Path of the folder where to extract the file

'Return Value		: 	True or False

'Examples			:  	bReturn = Fn_WorkflowProcess_ResourceTreeSelect("Review With Profile","New")
'History			:			
'						Developer Name	      	Date		      Rev. No.	    Changes Done	
'------------------------------------------------------------------------------------------------------------------------------------------------------
'						Vaishali D             22-July-2022		 	1.0         	Created	
'------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function Fn_MSO_UnZipFile(SourceFilePath,TargetExtractFolderName)
    Dim OZipObject,OFile,fso
    Set OFile = DotNetFactory.CreateInstance("System.IO.File","System") 
    
    'checking existance of zip file
	If not OFile.Exists(SourceFilePath) Then
		Set OFile =nothing
		Exit function
	End If
	
	'delete folder before extract
	Set fso = CreateObject("scripting.filesystemobject")
	If  fso.FolderExists(TargetExtractFolderName) Then
		fso.DeleteFolder TargetExtractFolderName, True
	       wait 2
	 End If
	
    If OFile.Exists(SourceFilePath) Then
        Set OZipObject = DotNetFactory.CreateInstance("System.IO.Compression.ZipFile","System.IO.Compression.FileSystem")
                
      OZipObject.ExtractToDirectory SourceFilePath, TargetExtractFolderName  
    End IF
    wait 3
	Fn_MSO_UnZipFile = True    
End Function

