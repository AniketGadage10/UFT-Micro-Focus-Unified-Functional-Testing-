Option Explicit

'*********************************************************	Function List		***********************************************************************
'1.  Fn_SISW_GetVISIOStencilFile( ) 'Returns  Stencil File Name
'2.  Fn_SISW_SE_DrawVisioDiagram(sAction, arrDropObjs, sConnector, sConnFrom, sConnTo )
'3.  Fn_SISW_GetShapeData(sShpName, sShpProp) 'Return the value of the property from VISIO diagram
'4.  
'5. 
'



'****************************************    Function to get Stencil File ***************************************
'
''Function Name		 	:	Fn_SISW_GetVISIOStencilFile
'
''Description		    :  	Function to get Stencil File From Temprary Location

''Parameters		    :	None
								
''Return Value		    :  	Full Path of the Stencil File
'
''Examples		     	:	Fn_SISW_GetVISIOStencilFile()

'History:
'	Developer Name			Date			Rev. No.		Reviewer		Changes Done	
'--------------------------------------------------------------------------------------------------------------------
'	Archana D		 27-June-2013		1.0				
'--------------------------------------------------------------------------------------------------------------------

Function Fn_SISW_GetVISIOStencilFile( )
		GBL_FAILED_FUNCTION_NAME="Fn_SISW_GetVISIOStencilFile"
		On error resume next
		Dim objFSO, objFolder, objFiles
		Dim sPath , sFile
		Dim sFileExt, sTypicalExt

		Fn_GetVISIOStencilFile = False
		'Creates Objects for  File System
    	Set objFSO = CreateObject("Scripting.FileSystemObject")
		'Object of Temp path
		Set objFolder = objFSO.GetSpecialFolder(2)
		'Objects of file .vsx file  within Temp folder

		 For Each sFile In objFolder.Files
           sFileExt = ObjFSO.GetExtensionName(sFile)		
			If sFileExt  = "vsx" Then
			   Fn_SISW_GetVISIOStencilFile = sFile.Path
			End If
         Next

Set objFSO = Nothing
Set objFolder =Nothing
Set sFile = Nothing
End Function


'****************************************    Function to get DrawVisioDiagram ***************************************
'
''Function Name		 	:	Fn_SISW_SE_DrawVisioDiagram
'
''Description		    :  	Function to draw the VISIO diagram as per the value in sAction Variable

''Parameters		    :	1.sAction: Action to be done with respect to VISIO Daigram
'                                 2. arrDropObjs : Array of Objects to be draw on  VISIO diagram used in Create Action
'								 3. sConnector : Used in Connect Action. It is Name of the Connector Object ( if you want to Specify Connector property use : seperated array)
'								 4. sConnFrom : Object Name to be connected
'								 5. sConnTo : Object Name to be connected 
								
''Return Value		    :  	Full Path of the Stencil File
'
''Examples		     	:	Fn_SISW_GetVISIOStencilFile()

'History:
'	Developer Name			Date			Rev. No.		Reviewer		Changes Done	
'--------------------------------------------------------------------------------------------------------------------
'	Archana D		 27-June-2013		1.0				
'--------------------------------------------------------------------------------------------------------------------


Function Fn_SISW_SE_DrawVisioDiagram(sAction, arrDropObjs, sConnector, sConnFrom, sConnTo )
		GBL_FAILED_FUNCTION_NAME="Fn_SISW_SE_DrawVisioDiagram"
		Dim objVis, objVisDoc, objStncil, objVsoShape, objShp
		Dim sStencilFile, iCnt ,sObj
		Fn_SISW_SE_DrawVisioDiagram = False

		Select Case sAction
		
			Case "Create"   								
				' Clear all Stencil FIles from Temp				

				Set objVis = GetObject(,"Visio.Application")				
				Set objVisDoc =  objVis.Documents
	
				'Get Stencil File From Temp Folder	
				sStencilFile = Fn_SISW_GetVISIOStencilFile()

				'sStencilFile ="C:\Temp\tmp1372069858214.vsx"
				Set objStncil = objVisDoc.Open(sStencilFile)

				
				For iCnt = 0 to UBound(arrDropObjs)
                    If Cstr(arrDropObjs(iCnt)) <> ""  Then
							sObj = Split(Cstr(arrDropObjs(iCnt)), ":")					

							'Create Object for Stencil Shape to drop
							Set objVsoShape = objStncil.Masters(Cstr(sObj(0)))

							'Drop the Shape at given location
							Set objShp =  objVis.ActivePage.Drop(objVsoShape,sObj(1), sObj(2)) 
							Wait(10)
							' Set properties for the droped shape
							objShp.Text = sObj(3)					
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SISW_SE_DrawVisioDiagram ] New Shape ["+ sObj(0) +"] Dropped sucessfully")	
					End If
							sObj = Empty
				Next

				Fn_SISW_SE_DrawVisioDiagram = True
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SISW_SE_DrawVisioDiagram ] sucessfully dropped all shapes")	

		Case "Connect"
					Set objVis = GetObject(,"Visio.Application")
					
					Set objVisDoc =  objVis.Documents
		
					'Get Stencil File From Temp Folder		
					sStencilFile = Fn_SISW_GetVISIOStencilFile()	
				'	sStencilFile ="C:\Temp\tmp1372310808696.vsx"

					Set objStncil = objVisDoc.Open(sStencilFile)

					Dim iBeginX, iEndX, objShpConn1, objShpConn2			

                    If NOT IsEmpty(sConnector)  Then
							sObj = Split(Cstr(sConnector), ":")					

							'Create Object for Stencil Shape ( connector Object)
							Set objVsoShape = objStncil.Masters(Cstr(sObj(0)))

							'Drop the Shape at given location
							Set objShp =  objVis.ActivePage.Drop(objVsoShape,0, 0) 
							
                            Set iBeginX = objShp.Cells("BeginX") ' obtain a reference to the Connector's Begin X Cell
							Set iEndX = objShp.Cells("EndX") ' obtain a reference to the Connector's End X Cell

							'Create Object for Shapes to be connected
							 Set objAllShp = objVis.ActivePage.Shapes
							For Each shpObj In objAllShp							
									  If  shpObj.Text = sConnFrom Then
											Set objShpConn1 = shpObj
									  ElseIf shpObj.Text = sConnTo Then
											Set objShpConn2 = shpObj
									  End If								
							Next					
						
							'Connect Ends of Connector to Shapes
							Call iBeginX.GlueTo(objShpConn1.Cells("PinX"))
							Call iEndX.GlueTo(objShpConn2.Cells("PinX"))
							Wait(5)
							' Set properties for the droped shape
							objShp.Text = Cstr(sObj(1))
							Fn_SISW_SE_DrawVisioDiagram = True
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SISW_SE_DrawVisioDiagram ] sucessfully Connected given shapes")	

					End If
                							
			
		End Select

Set objVis = Nothing
Set objStncil= Nothing
Set objVsoShape = Nothing
Set objShp = Nothing
Set objVisDoc = Nothing

End Function



'****************************************    Function to get Clear Temporary Location of Windows ***************************************
'
''Function Name		 	:	Fn_SISW_ClearTemp
'
''Description		    :  	Function to Clear ( delete) all files from Temp location of Windows.

''Parameters		    :	1.sFiletype: Specify the Type of the File 
'                                 2. arrDropObjs : Array of Objects to be draw on  VISIO diagram used in Create Action
'								 3. sConnector : Used in Connect Action. It is Name of the Connector Object ( if you want to Specify Connector property use : seperated array)
'								 4. sConnFrom : Object Name to be connected
'								 5. sConnTo : Object Name to be connected 
								
''Return Value		    :  	Full Path of the Stencil File
'
''Examples		     	:	Fn_SISW_GetVISIOStencilFile()

'History:
'	Developer Name			Date			Rev. No.		Reviewer		Changes Done	
'--------------------------------------------------------------------------------------------------------------------
'	Archana D		 27-June-2013		1.0				
'--------------------------------------------------------------------------------------------------------------------


Function Fn_SISW_ClearTemp(sFiletype)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_ClearTemp"
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

'****************************************    Function to get Stencil File ***************************************
'
''Function Name		 	:	Fn_SISW_GetShapeData
'
''Description		    :  	Function to get Shape Data from VISIO diagram

''Parameters		    :	sShpName: Name of the shape for which data to be extracted
'								 sShpProp: Name of the property whose value need to find
								
''Return Value		    :  Returns value of property 
'
''Examples		     	:	Fn_SISW_GetShapeData("Log1", "_VisDM_Primary_Object_ID")

'History:
'	Developer Name			Date			Rev. No.		Reviewer		Changes Done	
'--------------------------------------------------------------------------------------------------------------------
'	Archana D		 		07-July-2013		1.0				
'--------------------------------------------------------------------------------------------------------------------

Function Fn_SISW_GetShapeData(sShpName,sShpProp,iInstance)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_GetShapeData"
Dim objVis, objVisDoc, objActivePage, objShp 
Dim arrProperty, iCnt, sPropValue, iInstCnt

Fn_SISW_GetShapeData = False
iInstCnt =1
				Set objVis = GetObject(,"Visio.Application")				
				
				Set objActivePage = objVis.ActivePage

				arrProperty = Split(sShpProp,":")	
				Err.Clear			

				For Each objShp in objActivePage.Shapes
				
						If objShp.Text = sShpName  Then
							If Cint(iInstCnt) = Cint(iInstance) Then
								For iCnt = 0 to UBound(arrProperty)
									 sPropValue = sPropValue + "," + objShp.Cells("Prop."& arrProperty(iCnt) ).Formula 
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SISW_GetShapeData ] sucessfully captured Property Value for Property ["+ arrProperty(iCnt)+"]")	
								Next
								Exit For
							Else
								iInstCnt = iInstCnt + 1
							End If
						End If
				Next

				If Err.Number < 0  Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail: Function [Fn_SISW_GetShapeData] Failed to get property value")
					Set objVis = Nothing
                	Set objActivePage = Nothing
					Set objShp = Nothing
					Exit Function
				End If

' Remove additional : from Value string
	sPropValue = Mid(sPropValue,2, Len(sPropValue))
	Fn_SISW_GetShapeData = sPropValue

Set objVis = Nothing
Set objActivePage = Nothing
Set objShp = Nothing
End Function
