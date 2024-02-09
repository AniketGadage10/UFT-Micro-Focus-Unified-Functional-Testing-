
Dim objVis

Public Function FN_SISW_VISIO_GetObject()
   	Set objVis = GetObject(, "VISIO.Application")
End Function


Public Function Fn_SISW_VISIO_GetShape(sObjName)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_VISIO_GetShape"
' Traverse and find out particular shape & return 

	Dim  objActivePage, objShp 
	Dim sPropValue, iInst, arrObjName, iCnt

	Fn_SISW_VISIO_GetShape = False	
	Call FN_SISW_VISIO_GetObject
	Set objActivePage = objVis.ActivePage
	Err.Clear		
	Select Case sObjName

	Case "ThePage"
			Set Fn_SISW_VISIO_GetShape = objActivePage.PageSheet
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SISW_VISIO_GetShape ] sucessfully Retrieved object ["+sObjName+"]")	
	Case Else	
			If instr(sObjName, "@") > 0 Then
				arrObjName= Split(sObjName,"@")
				sObjName = arrObjName(0)
				iInst = arrObjName(1)
			Else
				iInst = 1
			End If
			iCnt = 1
			For Each objShp in objActivePage.Shapes				
					If objShp.Text = sObjName  Then
						If Cint(iCnt) = Cint(iInst) Then
							Set Fn_SISW_VISIO_GetShape = objShp
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SISW_VISIO_GetShape ] sucessfully captured object ["+sObjName+"]")	
							Exit For
						Else
							iCnt = iCnt + 1
						End If
					End If
			Next
	End Select
	If Err.Number < 0  Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail: Function [Fn_SISW_VISIO_GetShape] Failed to get Shape")
	End If

	Set objActivePage = Nothing
	Set objShp = Nothing

End Function

Public Function Fn_SISW_VISIO_DropShape(sObjType, sObjName, iXpos, iYpos)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_VISIO_DropShape"
' Return Dropped Obj
	Dim objVisDoc, objStncil, objVsoShape, objShp
	Dim sStencilFile, iCnt ,sObj, bReturn
	Fn_SISW_VISIO_DropShape = False
	Err.Clear

	Call FN_SISW_VISIO_GetObject
	

		Set objVisDoc =  objVis.Documents
		'Get Stencil File From Temp Folder	
		sStencilFile = Fn_SISW_VISIO_GetVISIOStencilFile()
		Set objStncil = objVisDoc.Open(sStencilFile)
							
		Set objVsoShape = objStncil.Masters(cStr(sObjType))
	
		'Drop the Shape at given location
		Set objShp =  objVis.ActivePage.Drop(objVsoShape,cInt(iXpos), cInt(iYpos))
		Wait(10)
		' Set text property for the droped shape
		If sObjName <> "" Then
			 bReturn = FN_SISW_VISIO_SetShapeProperty(objShp, "Text",sObjName)
		End if

		If Err.Number < 0  Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_VISIO_DropShape ] Fail to drop Shape ["+ sObjName+"] ")	
		Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SISW_VISIO_DropShape ] New Shape ["+sObjName +"] Dropped sucessfully")	
			' Return Shape Obj
			Set Fn_SISW_VISIO_DropShape = objShp	
		End If			
        		
	Set objVisDoc = Nothing
	Set objStncil = Nothing
	Set objVsoShape =Nothing
	Set objShp =  Nothing
End Function

Public Function Fn_SISW_VISIO_GetVISIOStencilFile( )
		GBL_FAILED_FUNCTION_NAME="Fn_SISW_VISIO_GetVISIOStencilFile"
		On error resume next
		Dim objFSO, objFolder, objFiles
		Dim sPath , sFile
		Dim sFileExt, sTypicalExt

		Fn_SISW_VISIO_GetVISIOStencilFile = False
		'Creates Objects for  File System
    	Set objFSO = CreateObject("Scripting.FileSystemObject")
		'Object of Temp path
		Set objFolder = objFSO.GetSpecialFolder(2)
		'Objects of file .vsx file  within Temp folder

		 For Each sFile In objFolder.Files
           sFileExt = ObjFSO.GetExtensionName(sFile)		
			If sFileExt  = "vsx" Then
			   Fn_SISW_VISIO_GetVISIOStencilFile = sFile.Path
			End If
         Next

	Set objFSO = Nothing
	Set objFolder =Nothing
	Set sFile = Nothing
End Function



Public Function Fn_SISW_VISIO_ConnectConnector(objConnShp, objFrmShape, objToShape) 
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_VISIO_ConnectConnector"
' Return True or false
    Dim iBeginX, iEndX   	
	Fn_SISW_VISIO_ConnectConnector = False
	Err.Clear
		
		Set iBeginX = objConnShp.Cells("BeginX") ' obtain a reference to the Connector's Begin X Cell
		Set iEndX = objConnShp.Cells("EndX") ' obtain a reference to the Connector's End X Cell
	
		'Connect Ends of Connector to Shapes
		Call iBeginX.GlueTo(objFrmShape.Cells("PinX"))
		Call iEndX.GlueTo(objToShape.Cells("PinX"))
		Wait(3)
		
		If Err.Number < 0 Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_VISIO_ConnectConnector ] Faile to Connect given shapes")	
		Else
			Fn_SISW_VISIO_ConnectConnector = true
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SISW_VISIO_ConnectConnector ] sucessfully Connected given shapes")	
		End If

    Set iBeginX = Nothing
	Set iEndX = Nothing

End Function

Public Function Fn_SISW_VISIO_ConnectPortsToShape(objPort, objShape)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_VISIO_ConnectPortsToShape"
' Return True or false
	Dim objGluePoint, objCellPrt, objBlockCell
	Dim  iRows, iRowIndx
	'To Drop  port Object
	Fn_SISW_VISIO_ConnectPortsToShape = False
	Err.Clear
	'  Actual connecting port to Block shapes
	iRows = objPort.RowCount(7)
	
	Set objGluePoint = Nothing
' Get connections points of the port
	For iRowIndx=0 to iRows -1		
			Set objCellPrt = objPort.CellsSRC( 7,irowIndx,4)		
			if(objCellPrt.Formula() =  "1") OR (objCellPrt.Formula()= "2") Then  
				  Set objGluePoint = objCellPrt
				  Exit For
			End if   		
	Next
	' Get connections points of the Block 
	Set objBlockCell =objShape.CellsSRC(7, 1,4)
	objGluePoint.GlueTo objBlockCell
	Wait(2)

	If Err.Number < 0 Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_VISIO_ConnectPortsToShape ] Faile to Connect port to given shapes")	
	Else
		Fn_SISW_VISIO_ConnectPortsToShape = True
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SISW_VISIO_ConnectPortsToShape ] sucessfully Connected port to given shapes")	
	End If
	
	Set objGluePoint = Nothing
	Set objCellPrt = Nothing
	Set objGluePoint = Nothing
	Set objBlockCell = Nothing
	
End Function

Public Function Fn_SISW_VISIO_GetShapeData(objShp, sShpProp )
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_VISIO_GetShapeData"
' Return Property of Shape
	Dim iCnt, sPropValue,arrProperty, sProp	
	Fn_SISW_VISIO_GetShapeData= False
		
        'Load Visio\ShapeData XML
		Environment.Value("sVis_ShpFile") = Fn_LogUtil_GetXMLPath("VISIO_ShapeData")

		arrProperty = Split(sShpProp,":")	
		Err.Clear			
		sPropValue =""

		For iCnt = 0 to UBound(arrProperty)
			sProp= Fn_GetXMLNodeValue( Environment.Value("sVis_ShpFile") , arrProperty(iCnt) )
			sPropValue = sPropValue + "," + objShp.Cells("Prop."& sProp).Formula 
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SISW_GetShapeData ] sucessfully captured Property Value for Property ["+ arrProperty(iCnt)+"]")	
	
			If Err.Number < 0  Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail: Function [Fn_SISW_GetShapeData] Failed to get property value")
				Exit For
			End if
		Next					
' Remove additional , from Value string
		If Len(sPropValue) > 2 Then
			sPropValue = Mid(sPropValue,2, Len(sPropValue))
			sPropValue = Replace(sPropValue, """","")
			Fn_SISW_VISIO_GetShapeData = sPropValue	
		End If
End Function

Public Function FN_SISW_VISIO_ActionMenuSelect(objShape, sMenu)
	GBL_FAILED_FUNCTION_NAME="FN_SISW_VISIO_ActionMenuSelect"
	Dim objCell, objSection
	Dim iRowNo,sShrtCtMenu, aMenu, sButtonName
	FN_SISW_VISIO_ActionMenuSelect = False
   'Select the requiered shape	
	objVis.ActiveWindow.DeselectAll
	If objShape.Name <>"ThePage" Then
		objVis.ActiveWindow.Select objShape,2
	End If
	Set objSection = objShape.Section(240)

	aMenu = Split(sMenu, "@")
	If aMenu(1) <> "" Then sButtonName = aMenu(1) 

	Select Case aMenu(0)
	Case "Delete in Teamcenter"
			For iRowNo = 0 To objSection.Count - 1
				' To get the name of the menu
					Set objCell = objShape.CellsSRC(240, iRowNo, 0)
					sShrtCtMenu = objCell.Formula()
					sShrtCtMenu = Replace(sShrtCtMenu, """", "")
					If sShrtCtMenu = aMenu(0) Then
						Set objCell = objShape.CellsSRC(240, iRowNo, 3)
						objCell.Trigger()       								
						JavaWindow("SystemsEngineering").JavaWindow("Confirmation").JavaButton("Delete").Click						
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Function [FN_SISW_VISIO_ActionMenuSelect] Sucessfully invoked Popup Menu")
						FN_SISW_VISIO_ActionMenuSelect = True
						Exit For
					End If	
			Next


	Case Else
			For iRowNo = 0 To objSection.Count - 1
				' To get the name of the menu
					Set objCell = objShape.CellsSRC(240, iRowNo, 0)
					sShrtCtMenu = objCell.Formula()
					sShrtCtMenu = Replace(sShrtCtMenu, """", "")
					If sShrtCtMenu = aMenu(0) Then
						Set objCell = objShape.CellsSRC(240, iRowNo, 3)
						objCell.Trigger()						
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Function [FN_SISW_VISIO_ActionMenuSelect] Sucessfully invoked Popup Menu")
						FN_SISW_VISIO_ActionMenuSelect = True
						Exit For
					End If	
			Next	

	End Select
	
	If Err.Number < 0 Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail: Function [FN_SISW_VISIO_ActionMenuSelect] Failed to invoke Popup Menu")
    End If

	'objVis.ActiveWindow.DeselectAll
	Set objSection = Nothing
	Set objCell = Nothing

End Function

Function Fn_SISW_VISIO_ClearTemp(sFiletype)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_VISIO_ClearTemp"
	On error resume next
		Dim objFSO, objFolder, objFiles
		Dim sPath , sFile
		Dim sFileExt, sTypicalExt
		'Creates Objects for  File System

		Set objFSO = CreateObject("Scripting.FileSystemObject")

		Select Case lcase(sFileType)

			Case "visio"
				sTypicalExt = "vsx:vdx"
		
			Case "all"
				sTypicalExt = "All"
		End Select
	
		'Object of Temp path
		Set objFolder = objFSO.GetSpecialFolder(2)
		'Objects of files within Temp folder

		 For Each sFile In objFolder.Files
           sFileExt = ObjFSO.GetExtensionName(sFile)
			If sTypicalExt = "All" Then
				 ObjFSO.DeleteFile sFile,True
			End If
			If Instr(1, sFileExt, sTypicalExt) > 1 Then
			   ObjFSO.DeleteFile sFile,True
			End If
         Next

		Fn_SISW_ClearTemp = True
'Clears out objects
Set objFSO = nothing
Set objFolder = nothing
End Function

Public Function FN_SISW_VISIO_SetShapeProperty(objShp, sProperty, sValue)
GBL_FAILED_FUNCTION_NAME="FN_SISW_VISIO_SetShapeProperty"
Dim arrProp, arrVal

FN_SISW_VISIO_SetShapeProperty = False
Err.Clear

	arrProp = Split(sProperty,",")
	arrVal = Split(sValue,",")
	For iCnt =0 to UBound(arrProp)
	
		Select Case Lcase(arrProp(iCnt))
	
		Case LCase("Text")
					objShp.Text = arrVal(iCnt)
					Wait(2)
		End Select
	Next

	If Err.Number < 0  Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail: Function [FN_SISW_VISIO_SetShapeProperty] Failed to set property")
	End If

FN_SISW_VISIO_SetShapeProperty = True
Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Function [FN_SISW_VISIO_SetShapeProperty] Sucessfully set property")	

End Function

Function FN_SISW_VISIO_GetShapeLocation(objShape, sLocation)
	GBL_FAILED_FUNCTION_NAME="FN_SISW_VISIO_GetShapeLocation"
Dim objCel
Dim iLoc
	FN_SISW_VISIO_GetShapeLocation = False
	Err.Clear

   Select Case LCase(sLocation)
   Case "xloc"
					Set objCel = objShape.Cells("pinx") 
					iLoc = objCel.Result("inches") 
	Case "yloc"
					Set objCel = objShape.Cells("piny") 
					iLoc = objCel.Result("inches") 
	Case "xloc:yloc"
					Set objCel = objShape.Cells("pinx") 
					iLoc = objCel.Result("inches") 
					Set objCel = objShape.Cells("piny") 
					iLoc = cStr(iLoc) +":"+ cStr(objCel.Result("inches"))
	Case "begin:end"
					Set objCel = objShape.Cells("beginx") 
					iLoc = objCel.Result("inches") 
					Set objCel = objShape.Cells("endy") 
					iLoc = cStr(iLoc) +":"+ cStr(objCel.Result("inches"))
	
    End Select

   If err.Number < 0  Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail: Function [FN_SISW_VISIO_GetShapeLocation] Failed to get Location")		
   Else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Function [FN_SISW_VISIO_GetShapeLocation] Sucessfully get Location")	
		FN_SISW_VISIO_GetShapeLocation = iLoc
   End If

Set objCel = Nothing
End Function
