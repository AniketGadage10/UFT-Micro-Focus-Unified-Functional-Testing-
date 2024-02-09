Option Explicit 

'*********************************************************	Function List Start	***********************************************************************
'00 .Fn_SISW_CMViewer_GetObject
'01. Fn_CMViewer_TreeNodePath()
'02. Fn_CMViewer_TreeNodeOperation()
'*********************************************************	Function List	End 	********************************************************************
'****************************************    Function to get Object hierarchy ***************************************
'
''Function Name		 	:	Fn_SISW_CMViewer_GetObject
'
''Description		    :  	Function to get Object hierarchy

''Parameters		    :	1. sObjectName : Object Handle name
								
''Return Value		    :  	Object \ Nothing
'
''Examples		     	:	Fn_SISW_CMViewer_GetObject("CMVApple")

'History:
'	Developer Name			Date			Rev. No.		Reviewer		Changes Done	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Sachin Joshi		 27-June-2012		1.0	
'-----------------------------------------------------------------------------------------------------------------------------------
'	Ashwini Kumar		 25-Sept-2013		2.0								Externalized object hierarchies
'-----------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_CMViewer_GetObject(sObjectName)
	Dim sObjectXMLPath
	sObjectXMLPath = Fn_GetEnvValue("User", "AutomationDir") & "\TestData\AutomationXML\ObjectXML\CMViewer.xml"
	Set Fn_SISW_CMViewer_GetObject = Fn_SISW_Setup_GetObjectFromXML(sObjectXMLPath, sObjectName)
End Function
'*********************************************  Function to get Tree node path **************************************************************

'Function Name			:			        Fn_CMViewer_TreeNodePath

'Description			    :		 		  Function to get Tree node path 

'Parameters			   :	 			1. objTree : Tree Object
'									2. StrNodeName : " : " separated Node path

'Return Value		   	   : 				Tree node path / False

'Pre-requisite			    :		 		CMViewer winodw should be displayed.

'Examples				    :				Call Fn_CMViewer_TreeNodePath(objTree, "CR0048/A;1-ic1:Tasks to Perform @1")
'								         Call Fn_CMViewer_TreeNodePath(objTree, "CR0048/A;1-ic1:Tasks to Perform")

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'											Koustubh			  3-Nov-2010	   		     1.0
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_CMViewer_TreeNodePath(objTree, StrNodeName)
	GBL_FAILED_FUNCTION_NAME="Fn_CMViewer_TreeNodePath"
	Dim NodeLists, intNodeCount, intCount, StrExist, aMenuList, sTreeItem,sCmpItm
	Dim iCounter, iItemCount, aNodePath,  iInstance, instCount, aNodes
	Dim sPath, sEle ,iCnt, bFound
	
	If instr(StrNodeName, "@") > 0 Then
		'--- For selecting single node with instance ID --
		aNodePath = split(StrNodeName, "@",-1, 1)
		StrNodeName = trim(aNodePath(0))
		iInstance = cint(aNodePath(1))
	
		aNodes = split(StrNodeName,":")
		sPath = ""
		For iCounter = 0 to uBound(aNodes) - 1
		
			If sPath = "" Then
				sPath = aNodes(iCounter)
			Else
				sPath = sPath & ":" & aNodes(iCounter)
			End If
		Next
		sEle = aNodes( UBound(aNodes) )
		bFound = False
		iItemCount = cInt(objTree.GetROProperty("items count"))
		instCount = 0
		For iCounter = 0 to iItemCount - 1
			If objTree.GetItem(iCounter) = sPath then
				For iCnt = 0 to  iItemCount - 1 
					iCounter = iCounter +1
					If  iCounter >=  iItemCount Then
						Exit for
					End If
					If objTree.GetItem(iCounter) = ( sPath &":" & sEle ) Then
						instCount = instCount + 1
						If instCount = iInstance  Then
							sPath = sPath & ":#" & iCnt
							bFound = True
							Exit for
						End If
					End If
				Next
			End If
			If bFound Then Exit for
		Next
		If bFound = False Then
			Fn_CMViewer_TreeNodePath = False
		Else
			Fn_CMViewer_TreeNodePath = "" &cstr(sPath)
		End If
	Else
	'--- For selecting single node without instance ID--
		Fn_CMViewer_TreeNodePath = StrNodeName
	End If
End Function

'*********************************************  Function to perform operations on Tree node **************************************************************

'Function Name			:			        Fn_CMViewer_TreeNodeOperation

'Description			    :		 		  Function to perform operations on Tree node

'Parameters			   :	 			1. objTree : Tree Object
'									2. StrNodeName : " : " separated Node path
'									3. StrMenu : ( for future use )

'Return Value		   	   : 				Tree node path / False

'Pre-requisite			    :		 		CMViewer winodw should be displayed.

'Examples				    :				Call  Fn_CMViewer_TreeNodeOperation("Select","CR0048/A;1-ic1:Tasks to Perform @1","")
'								         Call  Fn_CMViewer_TreeNodeOperation("Expand","CR0048/A;1-ic1:Tasks to Perform","")

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'											Koustubh			  3-Nov-2010	   		     1.0
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public function Fn_CMViewer_TreeNodeOperation(sAction,StrNodeName,StrMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_CMViewer_TreeNodeOperation"
	Dim objTree, sPath
	Set objTree = Window("CMViewer").JavaApplet("CMVApplet").JavaTree("CMVTree")
	Fn_CMViewer_TreeNodeOperation = False
	If objTree.Exist(4) = False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_CMViewer_TreeNodeOperation ] CM Viewer tree does not exist.")	
		Set objTree = nothing
		Exit function
	End If

	Select Case sAction
		Case "Select"
			sPath = Fn_CMViewer_TreeNodePath(objTree, StrNodeName)
			If sPath <> False Then
				objTree.Select sPath
				Fn_CMViewer_TreeNodeOperation = True
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_CMViewer_TreeNodeOperation ] : [ " & StrNodeName & " ] is not exist in CM Viewer tree.")	
			End If

		Case "Expand"
			sPath = Fn_CMViewer_TreeNodePath(objTree, StrNodeName)
			If sPath <> False Then
				objTree.Expand sPath
				Fn_CMViewer_TreeNodeOperation = True
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_CMViewer_TreeNodeOperation ] : [ " & StrNodeName & " ] is not exist in CM Viewer tree.")
			End If
		Case Else
	End Select
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_CMViewer_TreeNodeOperation ] executed successfully with Action [ " & sAction & " ].")	
	Set objTree = nothing
End Function

