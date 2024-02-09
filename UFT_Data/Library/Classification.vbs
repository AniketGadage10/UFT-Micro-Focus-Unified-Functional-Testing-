Option Explicit
'*********************************************************	Function List  *********************************************************************************
'00. Fn_SISW_Classification_GetObject()
'01. Fn_Classification_HierarchyTreeOperations()
'02. Fn_Classification_ClassifyObject()
'03. Fn_Classification_AttributeValues()
'04. Fn_Classification_CreateClassificationObject()
'05. Fn_Classification_SearchItem()
'06. Fn_Classification_Attributes_ListOfValues()
'07. Fn_Classification_ChangeUnit()
'08. Fn_Classification_UnitClassOpearations()
'09. Fn_Classification_UnitConversionTable()
'10. Fn_Classification_SearchTableOpeartion()
'11. Fn_Classification_AttributeUnitText()
'12. Fn_Classification_TableResultPrintOperations()
'13. Fn_Classification_XMLExport()
'14. Fn_Classification_CheckClassificationProperty()
'15. Fn_Classification_SetValues()
'16. Fn_Classification_ImportObjects()
'17 Fn_Classification_SetRevisionRule()
'18. Fn_Classification_FavoritesOperations()
'19. Fn_Classification_FavoritesTree() 
'20. Fn_Classification_UnitSystemSearch()
'21.Fn_Classification_ErrorHandler()
'***********************************************************************************************************************************************************
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 			Function to get Object hierarchy  		- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''Function Name		:	Fn_SISW_Classification_GetObject
'
''Description		  	 :  	Function to get Object hierarchy

''Parameters		   :	1. sObjectName : Object Handle name
								
''Return Value		   :  	Object \ Nothing
'
''Examples		     	:	Fn_SISW_Classification_GetObject("ClassificationApplet")

'History:                
'								Developer Name							Date				Rev. No.		Reviewer		Changes Done	
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'									Pranav Ingle		 				6-June-2012				1.0					Sunny
'-----------------------------------------------------------------------------------------------------------------------------------
'	Ashwini Kumar		 25-Sept-2013		2.0								Externalized object hierarchies
'-----------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_Classification_GetObject(sObjectName)
	Dim sObjectXMLPath
	sObjectXMLPath = Fn_GetEnvValue("User", "AutomationDir") & "\TestData\AutomationXML\ObjectXML\Classification.xml"
	Set Fn_SISW_Classification_GetObject = Fn_SISW_Setup_GetObjectFromXML(sObjectXMLPath, sObjectName)
End Function
''****************************************    Function to perform operations on Hierarchy Tree ***************************************
'
''Function Name		      :	      Fn_Classification_HierarchyTreeOperations
'
''Description			  :  	  Function to perform operations on Classification Hierarchy Tree 
'
''Parameters			  :	  	  1. sAction : Action need to perform
''							  2. sNode : Tree Node path separated by :
'							  3. sMenu : Popup Menu item

''Return Value		      : 	  True \ False
'
''Pre-requisite			  :		 Classification tree should be visible.
''Examples				  :			  
''								Call Fn_Classification_HierarchyTreeOperations("Select", "Classification Root:Unit Definition Class  [0]", "")
''								Call Fn_Classification_HierarchyTreeOperations("Activate", "Classification Root:Unit Definition Class  [0]", "")
''								Call Fn_Classification_HierarchyTreeOperations("Expand", "Classification Root:Unit Definition Class  [0]", "")
''								Call Fn_Classification_HierarchyTreeOperations("PopupMenuSelect", "Classification Root:Unit Definition Class  [0]", "Collapse")
''								bReturn=Fn_Classification_HierarchyTreeOperations("VerifyToolTip", "Classification Root:str  [0]", "str  [0]:alias:alias-german")    (the tooltip to be verified should be passed in the sMenu Parameter)
''								bReturn=Fn_Classification_HierarchyTreeOperations("VerifyToolTip", "Classification Root:str  [0]", "alias-german")
'								bReturn=Fn_Classification_HierarchyTreeOperations("Select", "Classification Root:Unit Definition Class", "")
''History:
''						Developer Name			Date				Rev. No.			Changes Done			Reviewer	
''-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
''					 	   Koustubh W	 	17-Dec-2010				1.0
''						SHREYAS W       22-April-2011            1.1
''						Sandeep N       07-May-2013            1.2					Added case : GetNodePath		Sonal P
''-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public function Fn_Classification_HierarchyTreeOperations(sAction, sNode, sMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_Classification_HierarchyTreeOperations"
	Dim objApplet, iRowCounter,bFlag,sValue,aProperties
    Dim arrNode,iCount,StrNode,iCounter,arrCurrNode,arrNodeNumber

	Fn_Classification_HierarchyTreeOperations = False
	Set objApplet = JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow")
	If objApplet.JavaTree("Hierarchy").exist(10) = False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_Classification_HierarchyTreeOperations ] Failed Classification Hiararchy Tree is not present.")
		Set objApplet = nothing
		Exit function
	End If
	Select Case sAction
' - - - - - - - - - - -  - - - - - - - - - - -  - - - - - - - - - - -  - - - - - - - - - - -  - - - - - - - - - - -  - - - - - - - - - - -  - - - - - - - - - - -  - - - - - - - - - - -  - - - - - - - - - - - 
		Case "Select"
				'iRowCounter = Fn_JavaTree_NodeIndexExt("Fn_Classification_HierarchyTreeOperations", objApplet , "HierarchyTree", sNode, "", "")
				iRowCounter = Fn_JavaTree_NodeIndexExt("Fn_Classification_HierarchyTreeOperations", objApplet , "Hierarchy", sNode, "", "")
				If iRowCounter = -1 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_Classification_HierarchyTreeOperations ] Failed case [ " & sAction & " ] Specified node is not present in tree.")
					Set objApplet = nothing
					Exit function
				End If
				objApplet.JavaTree("Hierarchy").Object.setSelectionRow iRowCounter
				Fn_Classification_HierarchyTreeOperations = True
' - - - - - - - - - - -  - - - - - - - - - - -  - - - - - - - - - - -  - - - - - - - - - - -  - - - - - - - - - - -  - - - - - - - - - - -  - - - - - - - - - - -  - - - - - - - - - - -  - - - - - - - - - - - 
		Case "Activate"
				iRowCounter = Fn_JavaTree_NodeIndexExt("Fn_Classification_HierarchyTreeOperations", objApplet , "Hierarchy", sNode, "", "")
				If iRowCounter = -1 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_Classification_HierarchyTreeOperations ] Failed case [ " & sAction & " ] Specified node is not present in tree.")
					Set objApplet = nothing
					Exit function
				End If
				objApplet.JavaTree("Hierarchy").Object.setSelectionRow iRowCounter
				Wait 1
				objApplet.JavaTree("Hierarchy").OpenContextMenu(sNode)
				Wait 1
				objApplet.JavaMenu("MenuSelect").SetTOProperty "label", "Select"
				Wait 1
				objApplet.JavaMenu("MenuSelect").Select  
				Fn_Classification_HierarchyTreeOperations = True
' - - - - - - - - - - -  - - - - - - - - - - -  - - - - - - - - - - -  - - - - - - - - - - -  - - - - - - - - - - -  - - - - - - - - - - -  - - - - - - - - - - -  - - - - - - - - - - -  - - - - - - - - - - - 
		Case "Expand"
				iRowCounter = Fn_JavaTree_NodeIndexExt("Fn_Classification_HierarchyTreeOperations", objApplet , "Hierarchy", sNode, "", "")
				If iRowCounter = -1 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_Classification_HierarchyTreeOperations ] Failed case [ " & sAction & " ] Specified node is not present in tree.")
					Set objApplet = nothing
					Exit function
				End If
				objApplet.JavaTree("Hierarchy").Object.setSelectionRow iRowCounter
				objApplet.JavaTree("Hierarchy").Object.setExpandedState objApplet.JavaTree("Hierarchy").Object.getSelectionPath(), true
				Fn_Classification_HierarchyTreeOperations = True
' - - - - - - - - - - -  - - - - - - - - - - -  - - - - - - - - - - -  - - - - - - - - - - -  - - - - - - - - - - -  - - - - - - - - - - -  - - - - - - - - - - -  - - - - - - - - - - -  - - - - - - - - - - - 
		Case "PopupMenuSelect"
				iRowCounter = Fn_JavaTree_NodeIndexExt("Fn_Classification_HierarchyTreeOperations", objApplet , "Hierarchy", sNode, "", "")
				If iRowCounter = -1 Then
					'If Node name contains numeric index for any child nodes
					objApplet.JavaTree("Hierarchy").Select(sNode)
				Else
					objApplet.JavaTree("Hierarchy").Object.setSelectionRow iRowCounter
				End If

				If Err.Number < 0 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_Classification_HierarchyTreeOperations ] Failed case [ " & sAction & " ] Specified node is not present in tree.")
					Set objApplet = nothing
					Exit function
				End If
				
				objApplet.JavaTree("Hierarchy").OpenContextMenu(sNode)
				objApplet.JavaMenu("MenuSelect").SetTOProperty "label", sMenu
				objApplet.JavaMenu("MenuSelect").Select  
				Fn_Classification_HierarchyTreeOperations = True

		  Case "PopupMenuVerify"
				iRowCounter = Fn_JavaTree_NodeIndexExt("Fn_Classification_HierarchyTreeOperations", objApplet , "Hierarchy", sNode, "", "")
				If iRowCounter = -1 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_Classification_HierarchyTreeOperations ] Failed case [ " & sAction & " ] Specified node is not present in tree.")
					Set objApplet = nothing
					Exit function
				End If
				objApplet.JavaTree("Hierarchy").Object.setSelectionRow iRowCounter
				objApplet.JavaTree("Hierarchy").OpenContextMenu(sNode)
				objApplet.JavaMenu("MenuSelect").SetTOProperty "label", sMenu
				if objApplet.JavaMenu("Label:="+sMenu).Exist then
						Fn_Classification_HierarchyTreeOperations = True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: [ Fn_Classification_HierarchyTreeOperations ] Successful case [ " & sAction & " ] Specified node is not present in tree.")
				Else
						Fn_Classification_HierarchyTreeOperations = false
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_Classification_HierarchyTreeOperations ] Failed case [ " & sAction & " ] Specified node is not present in tree.")
						Set objApplet = nothing
						Exit function
				End if

		Case "SelectUnitDefClass"
			  iRowCounter = JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaTree("Hierarchy").GetROProperty( "items count")
			  'Select the Unit definition class	

			For iCounter = 0 to iRowCounter-1
					If  instr(1,JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaTree("Hierarchy").GetItem(iCounter),"Classification Root:Unit Definition Class",1) > 0 Then
							JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaTree("Hierarchy").Select "#0:#"&cint(iCounter)-1						
							JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaTree("Hierarchy").Activate "#0:#"&cint(iCounter)-1																		
					End If
			Next
'				'Select the Unit definition class	
'				  JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaTree("Hierarchy").OpenContextMenu "#0:#"&cint(iRowCounter)-2
'				  objApplet.JavaMenu("MenuSelect").SetTOProperty "label", "Select"
'				 objApplet.JavaMenu("MenuSelect").Select  
				  If Err.Number < 0 Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Unit Definition Class")
							Fn_Classification_HierarchyTreeOperations = false
							Exit function
				   Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected Unit Definition Class")
							Fn_Classification_HierarchyTreeOperations = True
							wait(1)
				  End If	
				  	
		Case "DoubleClick"
				iRowCounter = Fn_JavaTree_NodeIndexExt("Fn_Classification_HierarchyTreeOperations", objApplet , "Hierarchy", sNode, "", "")
				If iRowCounter = -1 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_Classification_HierarchyTreeOperations ] Failed case [ " & sAction & " ] Specified node is not present in tree.")
					Set objApplet = nothing
					Exit function
				End If
'				objApplet.JavaTree("Hierarchy").Select iRowCounter
				objApplet.JavaTree("Hierarchy").Object.setSelectionRow iRowCounter
				objApplet.JavaTree("Hierarchy").Activate "#0:#"&cint(iRowCounter)-1
				Fn_Classification_HierarchyTreeOperations = True


		Case "VerifyToolTip"

			bFlag=False
			'first select the node of which the tooltip has to be verified
            JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaTree("Hierarchy").Select sNode

			'Set focus on that node
            JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaTree("Hierarchy").Object.focusable(True)

			'Extract the tooltip of that string
			sValue=JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaTree("Hierarchy").Object.getJToolTip().toString
			If instr(1,sMenu,":",1)>0 Then
				aProperties=split(sMenu,":",-1,1)
				For iCount=0 to uBound(aProperties)
				If instr(1,sValue,aProperties(iCount),1)>0 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully verified the ToolTipText as "+aProperties(iCount)+" of the node "+sNode)
					bFlag=True
				End If
			Next
		Else
				If instr(1,sValue,sMenu,1)>0 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully verified the ToolTipText as "+sMenu+" of the node "+sNode)
					bFlag=True
				End If
	End If

			If bFlag=true Then
					Fn_Classification_HierarchyTreeOperations = True
			Else
					Fn_Classification_HierarchyTreeOperations = False
			End If
' - - - - - - - - - - -  - - - - - - - - - - -  - - - - - - - - - - -  - - - - - - - - - - -  - - - - - - - - - - -  - - - - - - - - - - -  - - - - - - - - - - -  - - - - - - - - - - -  - - - - - - - - - - - 
		Case "GetNodePath"
			arrNode=Split(sNode,":")
			For iCount=0 to ubound(arrNode)
				bFlag=False
				If iCount=0 Then
					StrNode=arrNode(0)
				Else
					StrNode=StrNode+":"+arrNode(iCount)
				End If
				For iCounter=0 to objApplet.JavaTree("Hierarchy").GetROProperty("items count")-1
					If instr(1,objApplet.JavaTree("Hierarchy").GetItem(iCounter),StrNode) Then
						If instr(1,objApplet.JavaTree("Hierarchy").GetItem(iCounter),"[") Then
							arrCurrNode=split(objApplet.JavaTree("Hierarchy").GetItem(iCounter),":")
							arrNodeNumber=Split(arrCurrNode(ubound(arrCurrNode)),"[")
							StrNode=StrNode+"  ["+arrNodeNumber(1)
						End If
						bFlag=True
						Exit for
					End If
				Next
				If bFlag=False Then
					Exit for
				End If
			Next
			If bFlag=True Then
				Fn_Classification_HierarchyTreeOperations=StrNode
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_Classification_HierarchyTreeOperations ] Failed case [ " & sAction & " ] Specified node is not present in tree.")
				Fn_Classification_HierarchyTreeOperations=False
			End If
' - - - - - - - - - - -  - - - - - - - - - - -  - - - - - - - - - - -  - - - - - - - - - - -  - - - - - - - - - - -  - - - - - - - - - - -  - - - - - - - - - - -  - - - - - - - - - - -  - - - - - - - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_Classification_HierarchyTreeOperations ] Invalid case [ " & sAction & " ].")
			Set objApplet = nothing
			Exit function
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	End Select
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: [ Fn_Classification_HierarchyTreeOperations ] Executed successfully with case [ " & sAction & " ].")
	Set objApplet = nothing
End Function

''****************************************    Function to classify Object in Classification perspective  ***************************************
'
''Function Name		      :	      Fn_Classification_ClassifyObject
'
''Description			  :  	  Function to classify Object in Classification perspective
'
''Parameters			:	  	  1. sAction : Action need to perform
''							  2. sNavTreeNode : Nav Tree Node path separated by :
'							  3. sValueToVerify : for future use
'							  4. sBtnName : for future use

''Return Value		      : 	  True \ False
'
''Pre-requisite		       : 
''Examples		       :	   Call Fn_Classification_ClassifyObject("ClassifyObject", sNavTreeNode, "", "")
'									  Dim message = "The object  ""000070 - ps1"" "+vblf+"is not classified - there is no ICO for it."+vblf+"Do you want to classify it?"	
'									  Call Fn_Classification_ClassifyObject("VerifyMessage", , message, "")	
''History:
''						Developer Name			Date				Rev. No.			Changes Done			Reviewer	
''-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
''					 	   Koustubh W	 		18-Dec-2010				1.0
''-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
''					 	   Meeta					01-Aug-2012				1.0						Added new object hierarchy for Classify Object dialog
''-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public function Fn_Classification_ClassifyObject(sAction, sNavTreeNode, sValueToVerify, sBtnName)
	GBL_FAILED_FUNCTION_NAME="Fn_Classification_ClassifyObject"
	Dim sMessage
	Fn_Classification_ClassifyObject = False

	If sNavTreeNode <> "" Then
		' call nav tree node operation
		bReturn = Fn_MyTc_NavTree_NodeOperation("PopupMenuSelect", sNavTreeNode,"Send To:Classification")
		If bReturn = FALSE Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_Classification_ClassifyObject ] Failed to send item revision to Classification.")
			exit function
		Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: [ Fn_Classification_ClassifyObject ] Item revision sent to Classification.")
		End If
		Wait(5)
	End If
	Select Case sAction
		Case "ClassifyObject"
				If JavaDialog("Classify Object").exist(10) then
					JavaDialog("Classify Object").JavaButton("Yes").Click micLeftBtn
					Fn_Classification_ClassifyObject = True
				Elseif JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaDialog("Classify Object").Exist(10) then
					JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaDialog("Classify Object").JavaButton("Yes").Click micLeftBtn
					Fn_Classification_ClassifyObject = True
				Elseif JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaDialog("Classify Object").Exist(5) then
					JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaDialog("Classify Object").JavaButton("Yes").Click micLeftBtn
					Fn_Classification_ClassifyObject = True
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_Classification_ClassifyObject ] Classify Object window does not exist.")
					Exit function
				End If 
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "VerifyMessage"
				If JavaDialog("Classify Object").exist(10) then
					sMessage = JavaDialog("Classify Object").JavaObject("MLabel").Object.getText()		
				Elseif JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaDialog("Classify Object").Exist(10) then
					sMessage = JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaDialog("Classify Object").JavaObject("MLabel").Object.getText()    					
				Elseif JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaDialog("Classify Object").Exist(5) then					
					sMessage = JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaDialog("Classify Object").JavaObject("MLabel").Object.getText()
				Elseif JavaWindow("ClassificationMainWin").JavaWindow("Save Instance").Exist(5) then					
					sMessage = JavaWindow("ClassificationMainWin").JavaWindow("Save Instance").JavaEdit("Text").Object.getText()
				Else		
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_Classification_ClassifyObject ] Classify Object window does not exist.")
					Exit function
				End If 
'
				If trim(sMessage) = trim(sValueToVerify) Then
						Fn_Classification_ClassifyObject = True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: [ Fn_Classification_ClassifyObject ] Message Verified Successfully .")
				else
						Fn_Classification_ClassifyObject = False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: [ Fn_Classification_ClassifyObject ] Failed to Verify Message.")
						Exit Function
				End If
				
				If sBtnName<>"" Then
					if JavaWindow("ClassificationMainWin").JavaWindow("Save Instance").Exist(5) then
						JavaWindow("ClassificationMainWin").JavaWindow("Save Instance").JavaButton("OK").Click micLeftBtn	
					End If
				End If
'				- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_Classification_ClassifyObject ] Invalid case [ " & sAction & " ].")
			Exit function
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	End Select
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: [ Fn_Classification_ClassifyObject ] Executed successfully with case [ " & sAction & " ].")
End Function

''****************************************    Function to Opeartions on Attribute Values***************************************
'
''Function Name		:	Fn_Classification_AttributeValues
'
''Description		:  	Function to Opeartions on Attribute Values 
'
''Parameters		:	1. sAction : Action need to perform
''						2. sAttributeType : Text, dropdown etc
'						3. sValue : Multiple values seperated with ',' and single entity seperated by ':' Single entity denotes 'name of attribute:value to add' 
'						4. sDetails : For feature use

''Return Value		:	True \ False
'
''Examples			:	
''						Call Fn_Classification_AttributeValues("Add", "Text", "IntegerAttrib:7","")
''						Call Call Fn_Classification_AttributeValues("Verify", "Text", "IntegerAttrib","")
''History			:
''			Developer Name		Date		Rev. No.	Changes Done																	Reviewer	
''-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
''			Prasanna 	 	29-Dec-2010		 1.0
''			Snehal S 	 	14-Jan-2016		 1.1		Added new case "TextMultipleValues" from TC10.1.5								[TC1122-2016010600-14_Jan_2016-VivekA-Maintenance]
''-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_Classification_AttributeValues(sAction,sAttributeType,sValue,sDetails)
	GBL_FAILED_FUNCTION_NAME="Fn_Classification_AttributeValues"
	Dim aValue,aSetValues,aGetValues,sTextVal
	Dim bFlag, iCounter, aSetAttribute, iCount, sIndex
	
	Select Case sAction
		
				Case "Add"              
						  Select Case sAttributeType
			  						  Case "Text", "Text_Ext"
											If instr(1,sValue,",") Then
													aValue = split(sValue,",",-1,1)
											Else
													aValue = Array(sValue)
											End If

											For iCounter = 0 to Ubound(aValue)
													 aSetValues = split(aValue(iCounter),":",-1,1)
													  'Addedby Sandeep Navghane : 3-May-2013
													 JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaStaticText("AttributeValue_Label").SetTOProperty "label",aSetValues(0)
												 If sAttributeType = "Text_Ext" Then
													 	Call Fn_SISW_UI_JavaEdit_Operations("Fn_Classification_AttributeValues", "setText",JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow"),"AttributeValue_Edit", aSetValues(1))
												 Else
													 	Call Fn_Edit_Box("Fn_Classification_AttributeValues",JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow"),"AttributeValue_Edit", aSetValues(1))													 
												 End If
												
													'Commented by Sandeep Navghane : 3-May-2013
'													 JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaEdit("ClassDicField").SetTOProperty "attached text",  aSetValues(0)
'													  wait(8)
'													 JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaEdit("ClassDicField").Set  aSetValues(1)
													 wait 3
													 If Err.Number < 0 Then
															Fn_Classification_AttributeValues = false
															Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Set value [" + aSetValues(1) + "] for Attribute ["+aSetValues(0)+"]" ) 
															Exit Function
													 Else
															Fn_Classification_AttributeValues = true
															Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Set value [" + aSetValues(1) + "] for Attribute ["+aSetValues(0)+"]" )
													 End If
											  Next
											  
									Case "Date" 
									
										aValue = split(sValue,":")
										
										Set objectDate = JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaEdit("AttributeMultipleValues_Edit")
										objectDate.SetTOProperty "attached text" , aValue(0)
										
										  If objectDate.Exist(1) Then
												objectDate.Click 1,1
												wait 5
												objectDate.Set aValue(1)
												wait 1
												objectDate.RefreshObject()
												wait 1
											Call  Fn_KeyBoardOperation("SendKey", "{Tab}")
											Fn_Classification_AttributeValues = true
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Set value [" + aValue(1) + "] for Attribute ["+aValue(0)+"]" )
										Else 
											Fn_Classification_AttributeValues = false
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Set value [" + aValue(1) + "] for Attribute ["+aValue(0)+"]" ) 
											Exit Function
										End If

									Case "List"
											If instr(1,sValue,",") Then
													aValue = split(sValue,",",-1,1)
											Else
													aValue = Array(sValue)
											End If

											For iCounter = 0 to Ubound(aValue)
													aSetValues = split(aValue(iCounter),":",-1,1)
													JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaList("ClassDicList").SetTOProperty "attached text",  aSetValues(0)
													If Fn_SISW_UI_Object_Operations("Fn_Classification_AttributeValues","Exist", JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaList("ClassDicList"), "") Then
														JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaList("ClassDicList").Select aSetValues(1)
													Else
														JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaStaticText("AttributeValue_Label").SetTOProperty "label",  aSetValues(0)
														wait 1
														JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaList("ClassDicList_Ext").Select aSetValues(1)
													End If

													wait(1)
													If Err.Number < 0 Then
															Fn_Classification_AttributeValues = false
															Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Set value [" + aSetValues(1) + "] for Attribute [" + aSetValues(0) + "] " )
															Exit Function
													Else
															Fn_Classification_AttributeValues = true
															Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Set value [" + aSetValues(1) + "] for Attribute [" + aSetValues(0) + "] " )
													End If
											Next

									 Case "ListWithTilda"
										If instr(1,sValue,"~") Then
													aValue = split(sValue,"~",-1,1)
											Else
													aValue = Array(sValue)
											End If

											For iCounter = 0 to Ubound(aValue)
													aSetValues = split(aValue(iCounter),":",-1,1)
													
													 JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaList("ClassDicList").SetTOProperty "attached text",  aSetValues(0)
													 JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaList("ClassDicList").Select aSetValues(1)
													 wait(1)
													 If Err.Number < 0 Then
															Fn_Classification_AttributeValues = false
															Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Set value [" + aSetValues(1) + "] for Attribute [" + aSetValues(0) + "] " )
															Exit Function
													 Else															
															Fn_Classification_AttributeValues = true
															Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Set value [" + aSetValues(1) + "] for Attribute [" + aSetValues(0) + "] " )
													 End If
											  Next
									'[TC1122-2016010600-14_Jan_2016-VivekA-Maintenance] - Added to Add single or multiple values in single or multiple attributes - By Snehal S Added from TC1015
									Case "TextMultipleValues"										
											bFlag = False
											If instr(1,sValue,",") Then
													aValue = split(sValue,",",-1,1)
											Else
													aValue = Array(sValue)
											End If
											
											For iCounter = 0 to Ubound(aValue)
												aSetAttribute = split(aValue(iCounter),":",-1,1)
												If Instr(aSetAttribute(1),"~")>0 Then
													JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaEdit("AttributeMultipleValues_Edit").SetTOProperty "attached text",aSetAttribute(0)
													
													aSetValues = Split(aSetAttribute(1),"~")
													Wait 1
													JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaEdit("AttributeMultipleValues_Edit").Set aSetValues(0)
													
													For iCount = 0 To 20 Step 1
														JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaEdit("AttributeMultipleValues1_Edit").SetTOProperty "index",iCount
														Wait 1
														If Trim(JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaEdit("AttributeMultipleValues1_Edit").GetROProperty("value"))=Trim(aSetValues(0)) Then
															sIndex = iCount
															bFlag = True
															Exit For
														Else
															bFlag = False
														End If
													Next
													If bFlag=True Then
														For iCount = 1 To UBound(aSetValues)
															JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaEdit("AttributeMultipleValues1_Edit").SetTOProperty "index", sIndex+iCount
															Wait 1
															JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaEdit("AttributeMultipleValues1_Edit").Set aSetValues(iCount)
														Next
														Fn_Classification_AttributeValues = True
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Set multiple values for Attribute [" + aSetAttribute(0) + "] " )
													Else
														Fn_Classification_AttributeValues = False
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Set multiple values for Attribute [" + aSetAttribute(0) + "] " )
														Exit Function
													End If												
												Else
													JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaEdit("AttributeMultipleValues_Edit").SetTOProperty "attached text",aSetAttribute(0)
													JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaEdit("AttributeMultipleValues_Edit").Set aSetAttribute(1)
													Fn_Classification_AttributeValues = True
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Set value for Attribute [" + aSetAttribute(0) + "] " )
												End If
											Next
								End Select		  
					Case "Verify"
								Select Case sAttributeType
			  						  Case "Text"
											If instr(1,sValue,",") Then
													aValue = split(sValue,",",-1,1)
											Else
													aValue = Array(sValue)
											End If

											For iCounter = 0 to Ubound(aValue)													 
													 'JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaEdit("ClassDicField").SetTOProperty "attached text",  aValue(iCounter)   -old code commented by prasanna
													 JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaStaticText("StAttributeText").SetTOProperty "Label",  aValue(iCounter)
													 
													 wait(1)
													 If JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaStaticText("StAttributeText").Exist Then
															Fn_Classification_AttributeValues = true
															Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Verify value [" + aValue(iCounter) + "] for Attribute " )                 
													 Else															
															Fn_Classification_AttributeValues = false
															Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Verify value [" + aValue(iCounter) + "] for Attribute " ) 
															Exit Function
													 End If
											  Next

									  Case "List"
											If instr(1,sValue,",") Then
													aValue = split(sValue,",",-1,1)
											Else
													aValue = Array(sValue)
											End If

											For iCounter = 0 to Ubound(aValue)													 
													 JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaList("ClassDicList").SetTOProperty "attached text",  aValue(iCounter)
													 
													 wait(1)
													 If JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaList("ClassDicList").Exist Then
															Fn_Classification_AttributeValues = true
															Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Verify value [" + aValue(iCounter) + "] for Attribute " )                 
													 Else															
															Fn_Classification_AttributeValues = false
															Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Verify value [" + aValue(iCounter) + "] for Attribute " ) 
															Exit Function
													 End If
											  Next
										Case "ListWithTilda"
													If instr(1,sValue,"~") Then												    
															aValue = split(sValue,"~",-1,1)
													Else
															aValue = Array(sValue)
													End If

													For iCounter = 0 to Ubound(aValue)
															aSetValues = split(aValue(iCounter),":",-1,1)
															
															 JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaList("ClassDicList").SetTOProperty "attached text",  aSetValues(0)
															 JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaList("ClassDicList").Select aSetValues(1)
															 wait(1)
															 If Err.Number < 0 Then
																	Fn_Classification_AttributeValues = false
																	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Set value [" + aSetValues(1) + "] for Attribute [" + aSetValues(0) + "] " )
																	Exit Function
															 Else															
																	Fn_Classification_AttributeValues = true
																	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Set value [" + aSetValues(1) + "] for Attribute [" + aSetValues(0) + "] " )
															 End If
													Next

								End Select		  
					Case "VerifyValue"
								Select Case sAttributeType
			  						  Case "Text"
											If instr(1,sValue,",") Then
													aValue = split(sValue,",",-1,1)
											Else
													aValue = Array(sValue)
											End If

											For iCounter = 0 to Ubound(aValue)													 
													  aGetValues = split(aValue(iCounter),":",-1,1)
													 JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaEdit("ClassDicField").SetTOProperty "attached text",  aGetValues(0)
													 sTextVal = JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaEdit("ClassDicField").GetROProperty("value")
													 wait(1)
													 If trim(cstr(aGetValues(1))) = trim(cstr(sTextVal))Then
															Fn_Classification_AttributeValues = true
															Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Verify value [" + aValue(iCounter) + "] for Attribute " )                 
													 Else															
															Fn_Classification_AttributeValues = false
															Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Verify value [" + aValue(iCounter) + "] for Attribute " ) 
															Exit Function
													 End If
											  Next
									Case "List"
											If instr(1,sValue,",") Then
													aValue = split(sValue,",",-1,1)
											Else
													aValue = Array(sValue)
											End If

											For iCounter = 0 to Ubound(aValue)
													aGetValues = split(aValue(iCounter),":",-1,1)
													 JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaList("ClassDicList").SetTOProperty "attached text",  aGetValues(0)
													 sTextVal = JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaList("ClassDicList").GetROProperty("value")
													 wait(1)
													 If trim(cstr(aGetValues(1))) = trim(cstr(sTextVal))Then
															Fn_Classification_AttributeValues = true
															Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Verify value [" + aValue(iCounter) + "] for Attribute " )                 
													 Else															
															Fn_Classification_AttributeValues = false
															Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Verify value [" + aValue(iCounter) + "] for Attribute " ) 
															Exit Function
													 End If
											  Next
									Case "ListItemExist"
											If instr(1,sValue,",") Then
													aValue = split(sValue,",",-1,1)
											Else
													aValue = Array(sValue)
											End If

											For iCounter = 0 to Ubound(aValue)
													aGetValues = split(aValue(iCounter),":",-1,1)
													JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaList("ClassDicList").SetTOProperty "attached text",  aGetValues(0)
													If Fn_SISW_UI_Object_Operations("Fn_Classification_AttributeValues","Exist", JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaList("ClassDicList"), "") Then
														sTextVal = Fn_UI_ListItemExist("Fn_Classification_AttributeValues", JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow"), "ClassDicList",trim(cstr(aGetValues(1))))
													Else
														JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaStaticText("AttributeValue_Label").SetTOProperty "label",  aGetValues(0)
														wait 2
														sTextVal = Fn_UI_ListItemExist("Fn_Classification_AttributeValues", JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow"), "ClassDicList_Ext",trim(cstr(aGetValues(1))))
													End If
													wait 2
													
													wait(1)
													If sTextVal = True Then
														Fn_Classification_AttributeValues = true
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Verify value [" + aValue(iCounter) + "] for Attribute " )
													Else
														Fn_Classification_AttributeValues = false
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Verify value [" + aValue(iCounter) + "] for Attribute " ) 
														Exit Function
													End If
											Next

										Case "ListWithTilda"
													If instr(1,sValue,"~") Then
															aValue = split(sValue,"~",-1,1)
													Else
															aValue = Array(sValue)
													End If
		
													For iCounter = 0 to Ubound(aValue)
															aGetValues = split(aValue(iCounter),":",-1,1)
															 JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaList("ClassDicList").SetTOProperty "attached text",  aGetValues(0)
															 sTextVal = JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaList("ClassDicList").GetROProperty("value")
															 wait(1)
															 If trim(cstr(aGetValues(1))) = trim(cstr(sTextVal))Then
																	Fn_Classification_AttributeValues = true
																	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Verify value [" + aValue(iCounter) + "] for Attribute " )                 
															 Else															
																	Fn_Classification_AttributeValues = false
																	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Verify value [" + aValue(iCounter) + "] for Attribute " ) 
																	Exit Function
															 End If
													  Next
										Case "TextMultipleValues"
												bFlag = False
												If instr(1,sValue,",") Then
														aValue = split(sValue,",",-1,1)
												Else
														aValue = Array(sValue)
												End If
												
												For iCounter = 0 to Ubound(aValue)
													aSetAttribute = split(aValue(iCounter),":",-1,1)
													If Instr(aSetAttribute(1),"~")>0 Then
														JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaEdit("AttributeMultipleValues_Edit").SetTOProperty "attached text",aSetAttribute(0)
														
														aGetValues = Split(aSetAttribute(1),"~")
														Wait 1
														sTextVal = JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaEdit("AttributeMultipleValues_Edit").GetROProperty("value")
														
														If trim(cstr(aGetValues(0))) = trim(cstr(sTextVal)) Then
															For iCount = 0 To 20 Step 1
																JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaEdit("AttributeMultipleValues1_Edit").SetTOProperty "index",iCount
																Wait 1
																If Trim(JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaEdit("AttributeMultipleValues1_Edit").GetROProperty("value"))=Trim(aGetValues(0)) Then
																	sIndex = iCount
																	bFlag = True
																	Exit For
																Else
																	bFlag = False
																End If
															Next
															If bFlag=True Then
																For iCount = 0 To UBound(aGetValues)
																	JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaEdit("AttributeMultipleValues1_Edit").SetTOProperty "index", sIndex+iCount
																	Wait 1
																	sTextVal1 = JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaEdit("AttributeMultipleValues1_Edit").GetROProperty("value")
																	If trim(cstr(aGetValues(iCount))) = trim(cstr(sTextVal1)) Then
																		Fn_Classification_AttributeValues = True
																		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Verify multiple values for Attribute [" + aSetAttribute(0) + "]" )                 
																	Else
																		Fn_Classification_AttributeValues = False
																		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Verify multiple values for Attribute [" + aSetAttribute(0) + "]" ) 
																		Exit Function
																	End If
																Next
															Else
																Fn_Classification_AttributeValues = False
																Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Verify multiple values for Attribute [" + aSetAttribute(0) + "] " )
																Exit Function
															End If		
														Else
															Fn_Classification_AttributeValues = False
															Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Verify multiple values for Attribute [" + aSetAttribute(0) + "]" ) 
															Exit Function
														End If																								
													Else
														JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaEdit("AttributeMultipleValues_Edit").SetTOProperty "attached text",aSetAttribute(0)
														sTextVal = JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaEdit("AttributeMultipleValues_Edit").GetROProperty("value")
														wait 1
													 	If trim(cstr(aSetAttribute(1))) = trim(cstr(sTextVal)) Then
															Fn_Classification_AttributeValues = True
															Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Verify value [" + aSetAttribute(1) + "] for Attribute [" + aSetAttribute(0) + "]" )                 
													 	Else															
															Fn_Classification_AttributeValues = False
															Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Verify value [" + aSetAttribute(1) + "] for Attribute [" + aSetAttribute(0) + "]" ) 
															Exit Function
													 	End If
													End If
												Next

								End Select
			End Select

End Function


''****************************************    Function to Create Classification Object ***************************************
'
''Function Name		      :	      Fn_Classification_CreateClassificationObject
'
''Description			  :  	    Function to Create Classification Object
'
''Parameters			  :	  	1. sAction : Action need to perform
''							2. sHierarchyPath : Hierarchy path to be selected ( optional )
'							3. sICOType : ICO type
'							4. sICOId : ICO Id
'							5. bCopyValues = Boolean value to set copy values check box

''Return Value		      : 	  True \ False
'
''Examples				  :			  
''								Call Fn_Classification_CreateClassificationObject("Create", "Classification Root:Unit Definition Class  [0]", "Create New", "ico007", "True")
''								Call Fn_Classification_CreateClassificationObject("Create", "", "Add Additional", "ico008", "False")
''History:
''						Developer Name			Date				Rev. No.			Changes Done			Reviewer	
''-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
''					 	   Koustubh W 	 		   3-Jan-2011				1.0
''-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public function Fn_Classification_CreateClassificationObject(sAction, sHierarchyPath, sICOType, sICOId, bCopyValues)
	GBL_FAILED_FUNCTION_NAME="Fn_Classification_CreateClassificationObject"
	Dim objCreateICO, bReturn
	Fn_Classification_CreateClassificationObject = False
	Set objCreateICO = JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaDialog("Create Classification")

	Select Case sAction
		Case "Create"
				If sHierarchyPath <> ""  Then
					bReturn = Fn_Classification_HierarchyTreeOperations("Activate", sHierarchyPath, "")
					If bReturn = false Then
						Set objCreateICO = nothing
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to select node  [" + sHierarchyPath + "]." ) 
						Exit function
					End If
					bReturn = Fn_ToolbatButtonClick("Add or create a new Instance")
					If bReturn = false Then
						Set objCreateICO = nothing
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to select [Add or create a new Instance] Toolbar Button" ) 
						Exit function
					End If
				End If
				' setting ICO type
				Select Case sICOType
					Case "Add Additional"
						If sHierarchyPath = "" and objCreateICO.Exist(1) = False Then
							bReturn = Fn_ToolbatButtonClick("Add or create a new Instance")
							If bReturn = false Then
								Set objCreateICO = nothing
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to select [Add or create a new Instance] Toolbar Button" ) 
								Exit function
							End If
						End If
						objCreateICO.JavaRadioButton("Add Additional").Set "ON"
					Case "Create New"
						objCreateICO.JavaRadioButton("Create New").Set "ON"
				End Select
				' setting ICO Id
				If sICOId <> "" Then
					objCreateICO.JavaEdit("ICO Id").Set sICOId
				End If
				' setting set values check box
				If bCopyValues <> "" Then
					If cbool(bCopyValues) then
							If cInt(objCreateICO.JavaCheckBox("Copy Values").GetROProperty("enabled")) = 1 Then
									objCreateICO.JavaCheckBox("Copy Values").Set "ON"
							End If
					Else
							If cInt(objCreateICO.JavaCheckBox("Copy Values").GetROProperty("enabled")) = 1 Then
									objCreateICO.JavaCheckBox("Copy Values").Set "OFF"
							End If
					End IF
				End If
				'clicking on OK
				objCreateICO.JavaButton("OK").Click micLeftBtn
				Fn_Classification_CreateClassificationObject = True
	End Select
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : [ Fn_Classification_CreateClassificationObject ] executed successfully with case [ " & sAction & " ]." ) 
	Set objCreateICO = nothing
End Function


''****************************************    Function to Opeartions on Attribute Values***************************************
'
''Function Name		      :	      Fn_Classification_SearchItem
'
''Description			  :  	   Function to Search item & Select 
'
''Parameters			  :	  	  1. sAction : Action need to perform
''							  			 2. sObjectIds : obj id 
'							  			 3. sObjectRevs : obj rev
'											
'										 4. sObjectNames : object name
'										  5.sDetails	

''Return Value		      : 	  True \ False
'
''Examples				  :			  
''								Call Fn_Classification_SearchItem("Select", "001077", "", "Item2", "")

''History:
''						Developer Name			Date				Rev. No.			Changes Done			Reviewer	
''-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
''					 	   Mahendra	 			07-Jan-2011				1.0
''-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function Fn_Classification_SearchItem(sAction, sObjectIds, sObjectRevs, sObjectNames, sDetails)
	GBL_FAILED_FUNCTION_NAME="Fn_Classification_SearchItem"

	Dim objDialog, sValue, aObjIds, aObjRevs, aObjNames, iCounter,aAction

	Set objDialog = JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow")

		
		If instr(1,sAction,":") Then
					aAction = split(sAction,":",-1,1)
					sAction = aAction(0)
		End If

		If Trim(sObjectIds) <> "" Then
						aObjIds =Split(sObjectIds,":", -1, 1)
						aObjRevs =Split(sObjectRevs,":", -1, 1)
						aObjNames =Split(sObjectNames,":", -1, 1)				
		End If

		Select Case sAction

						Case "Verify"

								If objDialog.JavaButton("Search").GetROProperty("enabled")="1" Then
										objDialog.JavaButton("Search").Click micLeftBtn
										If Err.Number < 0 Then
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: To Click on Search Button.")
												Fn_Classification_SearchItem = False
												Exit Function
										Else
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully Clicked on Search Button.")
										End If
										Wait 2
								End If

								For iCounter = 0 To UBound(aObjIds)

										If objDialog.JavaButton("PropRight").GetROProperty("enabled")="1" Then
												objDialog.JavaButton("PropRight").Click micLeftBtn
												wait(1)
												If Err.Number < 0 Then
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: To Click on PropRight Button.")
														Fn_Classification_SearchItem = False
														Exit Function
												Else
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully Clicked on PropRight Button.")
												End If
												Wait 2
										End If

										If sObjectIds <> "" Then
	
											sValue = objDialog.JavaEdit("Object ID").GetROProperty("value")
											If Trim(sValue) <> aObjIds(iCounter) Then
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: To Verify the Object Id value ["+aObjIds(iCounter)+"] with ["+sValue+"].")
														Fn_Classification_SearchItem = False
														Exit Function
											Else
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully Verified the Object Id value ["+aObjIds(iCounter)+"] with ["+sValue+"] .")
											End If

										End If

										If sObjectRevs <> "" Then
	
											sValue = objDialog.JavaEdit("Object RevID").GetROProperty("value")
											If Trim(sValue) <> aObjRevs(iCounter) Then
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: To Verify the Object Revision ID value ["+aObjRevs(iCounter)+"] with ["+sValue+"].")
														Fn_Classification_SearchItem = False
														Exit Function
											Else
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully Verified the Object Revision ID value ["+aObjRevs(iCounter)+"] with ["+sValue+"] .")
											End If

										End If

										If sObjectNames <> "" Then
	
											sValue = objDialog.JavaEdit("ObjectName").GetROProperty("value")
											If Trim(sValue) <> aObjNames(iCounter) Then
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: To Verify the Object Name value ["+aObjNames(iCounter)+"] with ["+sValue+"].")
														Fn_Classification_SearchItem = False
														Exit Function
											Else
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully Verified the Object Name value ["+aObjNames(iCounter)+"] with ["+sValue+"] .")
											End If

										End If
								Next

				Case "Select"

								For iCounter = 0 To UBound(aObjIds)

										If sObjectIds <> "" Then
	
											objDialog.JavaEdit("Object ID").Set aObjIds(iCounter)
											If Err.Number < 0  Then
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: To Set the Object Id with value ["+aObjIds(iCounter)+"].")
														Fn_Classification_SearchItem = False
														Exit Function
											Else
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully Set the Object Id with value ["+aObjIds(iCounter)+"].")
											End If

											objDialog.JavaButton("Find").Click micLeftBtn
											If Err.Number < 0 Then
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: To Click on Find Button.")
														Fn_Classification_SearchItem = False
														Exit Function
											Else
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully Clicked on Find Button.")
											End If
											wait(1)

											If sObjectRevs <> "" Then
													sValue = objDialog.JavaEdit("Object RevID").GetROProperty("value")
													If Trim(sValue) <> aObjRevs(iCounter) Then
																Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: To Verify the Object Revision ID value ["+aObjRevs(iCounter)+"] with ["+sValue+"].")
																Fn_Classification_SearchItem = False
																Exit Function
													Else
																Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully Verified the Object Revision ID value ["+aObjRevs(iCounter)+"] with ["+sValue+"] .")
													End If
												End If
		
												If sObjectNames <> "" Then
													sValue = objDialog.JavaEdit("ObjectName").GetROProperty("value")
													If Trim(sValue) <> aObjNames(iCounter) Then
																Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: To Verify the Object Name value ["+aObjNames(iCounter)+"] with ["+sValue+"].")
																Fn_Classification_SearchItem = False
																Exit Function
													Else
																Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully Verified the Object Name value ["+aObjNames(iCounter)+"] with ["+sValue+"] .")
													End If
												End If
										End If
								Next

					Case "VerifySearch"
											
'											objDialog.JavaEdit("Object ID").Set trim(aAction(1))
											wait(5)
											objDialog.JavaEdit("Object ID").object.settext trim(aAction(1))
											If Err.Number < 0  Then
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: To Set the Object Id with value ["+aObjIds(iCounter)+"].")
														Fn_Classification_SearchItem = False
														Exit Function
											Else
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully Set the Object Id with value ["+aObjIds(iCounter)+"].")
											End If

											objDialog.JavaButton("Find").Click micLeftBtn
											If Err.Number < 0 Then
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: To Click on Find Button.")
														Fn_Classification_SearchItem = False
														Exit Function
											Else
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully Clicked on Find Button.")
											End If

											wait(2)
											Call Fn_ReadyStatusSync(5)
											'Retrive the value of search results obtained
											iResults = JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaStaticText("ResultNoStText").GetRoProperty("label")
											iResults = split(iResults," ",-1,1)
											bFound = false

											For iCounter = 0 To iResults(2)-1
			
													If  iCounter  = 0 Then
														If sObjectIds <> "" Then	
																sValue = objDialog.JavaEdit("Object ID").GetROProperty("value")
																If Trim(sValue) =  aObjIds(0) Then                                                          												
																			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully Verified the Object Id value ["+aObjIds(0)+"] with ["+sValue+"] .")
																			bFound = true
																			Exit for
																End If
														End If
													End If
			
													If cInt(objDialog.JavaButton("PropRight").GetROProperty("enabled")) = 1 Then
																objDialog.JavaButton("PropRight").Click micLeftBtn
																If Err.Number < 0 Then
																			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: To Click on Right Button.")
																			Fn_Classification_SearchItem = False
																			Exit Function
																Else
																			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully Clicked on Right Button.")
																End If
													End If
			
													wait(2)
			
													If sObjectIds <> "" Then	
															sValue = objDialog.JavaEdit("Object ID").GetROProperty("value")
															If Trim(sValue) =  aObjIds(0) Then                                                          												
																		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully Verified the Object Id value ["+aObjIds(0)+"] with ["+sValue+"] .")
																		bFound = true
																		Exit for
															End If
													End If
										
											Next
											'Click on clear button
											If sDetails = "Clear" Then
														objDialog.JavaButton("Clear").Click micLeftBtn
											End If
											If cbool(bFound)  Then
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully Verified theSearch result.")
											 else
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Verified the Search result.")
														Fn_Classification_SearchItem = False
														Exit function
											End If

		End Select
	Fn_Classification_SearchItem = True
	Set objDialog = Nothing

End Function

''****************************************    Function to perform operations Attributes values list box  ***************************************
'
''Function Name		      :	      Fn_Classification_Attributes_ListOfValues
'
''Description			  :  	  Function to perform operations on attributes values list box 
'
''Parameters			  :	  	  1. sAction : Action need to perform
''							 			 2. sName : Name of attribute
'							  			3. sIndex : Index of object
'										4. sSelectValue = value to select from list
'										5. sBtnName = Name of button to be cliked 
'										6. sDetails = future use
'

''Return Value		      : 	  True \ False
'
''Pre-requisite			  :		 Classification tree node should be visible.
''Examples				  :			  
''								Call Fn_Classification_Attributes_ListOfValues("Clear", "","","","","")
''								Call Fn_Classification_HierarchyTreeOperations("PopupMenuSelect", "Classification Root:Unit Definition Class  [0]", "Collapse")
''				
''				Developer Name			Date				Rev. No.			Changes Done			Reviewer	
''-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
''					 	   Prasanna	 			13-Jan-2011				1.0
''-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public function Fn_Classification_Attributes_ListOfValues(sAction, sName,sIndex,sSelectValue,sBtnName,sDetails)
	GBL_FAILED_FUNCTION_NAME="Fn_Classification_Attributes_ListOfValues"
	Dim aValues

	'If index is mentioned 	
	If sIndex <> "" Then
			JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaCheckBox("ChbListofValues").SetTOProperty "Index",trim(sTagName)
	Else
			JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaCheckBox("ChbListofValues").SetTOProperty "Index",0			
	End If

	'Click on Check-box
	JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaCheckBox("ChbListofValues").Set "ON"
	If Err.Number < 0 Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Click on Checkbox" )    
			Fn_Classification_Attributes_ListOfValues = false
			Exit Function
	Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully to Clicked on Checkbox " )    			
	End If
	'Click on JavaStatic Text
	JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaStaticText("ListofValues").DblClick 0,0,"LEFT"
	If Err.Number < 0 Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Click on Static Text [List of Values]" )    
			Fn_Classification_Attributes_ListOfValues = false
			Exit Function
	Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully to Click on Static Text [List of Values]" )    			
	End If
	Wait(1)
		
 Select Case sAction
			 Case "Add"
						'Select the Value from list
						JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaDialog("ListofValuesDialog").JavaList("ValueList").Select trim(cstr(sSelectValue))		
						If Err.Number < 0 Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Value ["+cstr(sSelectValue)+"]" )    
									Fn_Classification_Attributes_ListOfValues = false
									Exit Function
						Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected Value ["+cstr(sSelectValue)+"]")    			
						End If		

                        	If sBtnName<>""  Then

									JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaDialog("ListofValuesDialog").JavaButton(sBtnName).Click micLeftBtn
									If Err.Number < 0 Then
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to click on button ["+sBtnName+"]" )    
												Fn_Classification_Attributes_ListOfValues = false
												Exit Function
									Else
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully clicked on button ["+sBtnName+"]" )     			
									End If

						End If

						Fn_Classification_Attributes_ListOfValues = true
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Added the Value from List Of Values box") 		
                        
		 Case "Clear"
						'Select the Value from list
						If trim(sSelectValue) <> "" Then
							JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaDialog("ListofValuesDialog").JavaList("ValueList").Select trim(sSelectValue)		
							If Err.Number < 0 Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Value ["+cstr(sSelectValue)+"]" )    
										Fn_Classification_Attributes_ListOfValues = false
										Exit Function
							Else
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected Value ["+cstr(sSelectValue)+"]")    			
							End If	
						End If

						'Click on 'Clear' button
						JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaDialog("ListofValuesDialog").JavaButton("Clear").Click micLeftBtn
						If Err.Number < 0 Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Click on [Clear] Button" )    
									Fn_Classification_Attributes_ListOfValues = false
									Exit Function
						Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Clicked on [Clear] Button" ) 
						End If
						
						Fn_Classification_Attributes_ListOfValues = true
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Cleared the Value from List Of Values box") 

			Case "Verify"
						If trim(sSelectValue) <> "" Then							
								If instr(1,sSelectValue,",") Then
											aValues = split(sSelectValue,",",-1,1)
								else
											aValues = Array(sSelectValue)

								End If

								For iCounter = 0 to Ubound(aValues)
										bReturn = Fn_UI_ListItemExist("Fn_Classification_Attributes_ListOfValues", JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaDialog("ListofValuesDialog"), "ValueList",trim(aValues(iCounter)))
										If bReturn = false Then
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Verify Value Present in AutoFilter Dialog" )    
													Fn_Classification_Attributes_ListOfValues = false
													Exit Function
										Else
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Verify Value Present in AutoFilter Dialog" ) 
										End If

								Next

                                'Click on 'Cancel' button
						JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaDialog("ListofValuesDialog").JavaButton("Cancel").Click micLeftBtn
						If Err.Number < 0 Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Click on [Cancel] Button" )    
									Fn_Classification_Attributes_ListOfValues = false
									Exit Function
						Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Clicked on [Cancel] Button" ) 
						End If					
		
								Fn_Classification_Attributes_ListOfValues = true
						End if
									
			End Select
End Function



''****************************************    Function to chnage active unit***************************************
'
''Function Name		      :	      Fn_Classification_ChangeUnit
'
''Description			  :  	  Function to perform operations on attributes values list box 
'
''Parameters			  :	  	  1. sLable : metric/non-metric
''							 			 2. sDetails : Other details if reuqired 
'
''Return Value		      : 	  True \ False
'
''Examples				  :			  
''								Call Fn_Classification_ChangeUnit("non-metric", "")
'
''				
''				Developer Name			Date				Rev. No.			Changes Done			Reviewer	
''-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
''					 	   Prasanna	 			14-Jan-2011				1.0
''-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public function Fn_Classification_ChangeUnit(sLabel,sDetails)
	GBL_FAILED_FUNCTION_NAME="Fn_Classification_ChangeUnit"
		Dim xCord,yCord,DeviceReplay
	   xCord= JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaObject("LblActiveUnit").GetROProperty ("abs_x")
		yCord=JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaObject("LblActiveUnit").GetROProperty ("abs_y")
		JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaMenu("MenuSelect").SetTOProperty "label", sLabel

		Set DeviceReplay = CreateObject("Mercury.DeviceReplay")
		DeviceReplay.MouseClick xCord,yCord,LEFT_MOUSE_BUTTON
		wait 3

		JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaMenu("MenuSelect").Select 	
		If JavaDialog("ChangingActiveSystem").Exist Then
		
					   JavaDialog("ChangingActiveSystem").JavaButton("OK").Click micLeftBtn
						If Err.Number < 0  Then
								JavaDialog("ChangingActiveSystem").JavaButton("Cancel").Click micLeftBtn
								Fn_Classification_ChangeUnit = false
								Exit function
						End If		
		End If

		Fn_Classification_ChangeUnit = true

ENd function



'*********************************************************  Function performs  to create or modify Unit Class*********************************************************************
'Function Name  :   Fn_Classification_UnitClassOpearations
'
' 
'Parameters      :     sAction: Add,Modify
'           				  		aValues : Array of vlues to set
'							  		sDetails : for feature use
'									bSave : If class need to be save

'Return Value     :   True/False

' Pre-requisite :  Unit Definition Class already selected
'
'Examples    :      
'								Dim aValues 
'								aValues  = Array("1","2","3","m Metric","Y Yes","4","5","Y Yes","6") 
'								call Fn_Classification_UnitClassOpearations("Modify",aValues,"","")
' 
'History:
'          Developer Name                          Date                                                       Rev. No.         Changes Done                   Reviewer 
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'           Prasanna    		                      18-Jan-2011   1.0                         
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function Fn_Classification_UnitClassOpearations(sAction,aValues,sDetails,bSave)
			GBL_FAILED_FUNCTION_NAME="Fn_Classification_UnitClassOpearations"
		   Fn_Classification_UnitClassOpearations = false

		   Dim iItemCount,iCounter
		
			Select Case sAction
							Case "Add","Modify"
										
					
                                             For iCounter = 0 to Ubound(aValues)
														If aValues(iCounter) <> "" Then														
														
																If iCounter < 3  Then      ' For Measure, Unit Name, Unit Display Name, 
																			JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaEdit("UnitClassText").SetTOProperty "Index",iCounter
																			Window("ClassificationWindow").JavaEdit("UnitClassText").SetTOProperty "Index",iCounter
																			If JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaEdit("UnitClassText").Exist < 0 Then
																				JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaEdit("UnitClassText").Set aValues(iCounter)
																			Else
																				Window("ClassificationWindow").JavaEdit("UnitClassText").Set aValues(iCounter)
																			End If
																ElseIf  iCounter = 5  Then    ' For Conversion Multiplication Factor, 
																			JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaEdit("UnitClassText").SetTOProperty "Index",3
																			'Window("ClassificationWindow").JavaWindow("ClassificationSubWindow").JavaEdit("UnitClassText").SetTOProperty "Index",7
																			If JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaEdit("UnitClassText").Exist <0 Then
																				JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaEdit("UnitClassText").Set aValues(iCounter)  
																			Else
																				Window("ClassificationWindow").JavaWindow("ClassificationSubWindow").JavaEdit("UnitClassText").Set aValues(iCounter)
																			End If																		
																ElseIf   iCounter = 6  Then    ' For Conversion Addition Factor, 
																			JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaEdit("UnitClassText").SetTOProperty "Index",4
																			Window("ClassificationWindow").JavaEdit("UnitClassText").SetTOProperty "Index",8
																			If JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaEdit("UnitClassText").Exist < 0 Then																				
																				JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaEdit("UnitClassText").Set aValues(iCounter)  
																			Else
																				Window("ClassificationWindow").JavaEdit("UnitClassText").Set aValues(iCounter)
																			End If
																ElseIf   iCounter = 8  Then   ' Number of Decimal Places
																			JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaEdit("UnitClassText").SetTOProperty "Index",5
																			Window("ClassificationWindow").JavaEdit("UnitClassText").SetTOProperty "Index",9
																			If JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaEdit("UnitClassText").Exist < 0 Then																				
																				JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaEdit("UnitClassText").Set aValues(iCounter)  
																			Else
																				Window("ClassificationWindow").JavaEdit("UnitClassText").Set aValues(iCounter)
																			End If
																ElseIf   iCounter = 3 Then       'System of Measurment
																			JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaList("UnitClassList").SetTOProperty "Index",0
																			Window("ClassificationWindow").JavaEdit("UnitClassText").SetTOProperty "Index",0
																			If JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaList("UnitClassList").Exist < 0 Then																				
																				JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaList("UnitClassList").Select trim(aValues(iCounter))
																			Else
																				Window("ClassificationWindow").JavaEdit("UnitClassText").Select trim(aValues(iCounter))
																			End If
																ElseIf   iCounter = 4 Then        ' Base Unit  
																			JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaList("UnitClassList").SetTOProperty "Index",1
																			Window("ClassificationWindow").JavaEdit("UnitClassText").SetTOProperty "Index",1
																			If JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaList("UnitClassList").Exist < 0 Then																				
																				JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaList("UnitClassList").Select trim(aValues(iCounter))
																			Else
																				Window("ClassificationWindow").JavaEdit("UnitClassText").Select trim(aValues(iCounter))
																			End If
																ElseIf   iCounter = 7 Then			'Ignore for optimization	
																			JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaList("UnitClassList").SetTOProperty "Index",2	
																			Window("ClassificationWindow").JavaEdit("UnitClassText").SetTOProperty "Index",2																			
																			If JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaList("UnitClassList").Exist < 0 Then																				
																				JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaList("UnitClassList").Select trim(aValues(iCounter))
																			Else
																				Window("ClassificationWindow").JavaEdit("UnitClassText").Select trim(aValues(iCounter))
																			End If
																Else
																			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : No Such Value Exist")
																			Fn_Classification_UnitClassOpearations = false
																			Exit function
																End If

																If Err.Number < 0 Then
																			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Set Value ["+aValues(iCounter)+" ] Unit Definition Class")
																			 Fn_Classification_UnitClassOpearations = false
																			Exit function	
																 Else
																			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Set Value ["+aValues(iCounter)+" ] Unit Definition Class")																			
																End If		
																 Fn_Classification_UnitClassOpearations = true
													End if 
											  Next
											  
								
			End Select

			If bSave  <> "" Then                                               ' Save if  required
		                bReturn= Fn_ClassAdmin_ToolbarOperations("Save")
						If bReturn = false Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Save  Unit Definition Class")
									 Fn_Classification_UnitClassOpearations = false
									Exit function	
						 Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Saved Unit Definition Class")
									wait(1)
						End If	
	
			End If

End Function


'*********************************************************  Function performs  to O peartion on Unit Change Table *********************************************************************
'Function Name  :   Fn_Classification_UnitConversionTable
'
' 
'Parameters      :     sAction: Rowcellexist/rowexist
'           				  		sUnitType : metrc/non-metric
'							  		sObjectName : name of object
'									sPropertyName : Column name
'									sExpectedValue : Value to verify
'									rowNumber : row number starts from zero
'									sOther : feature use

'Return Value     :   True/False
'
'Examples    :      						
'								call Fn_Classification_UnitConversionTable("Rowcellexist","non-metric","","non-metric","439.633",0,"")
' 
'History:
'          Developer Name                          Date                                                       Rev. No.         Changes Done                   Reviewer 
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'           Prasanna    		                      19-Jan-2011   1.0                         
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_Classification_UnitConversionTable(sAction,sUnitType ,sObjectName, sPropertyName,sExpectedValue,rowNumber,sOther)
		GBL_FAILED_FUNCTION_NAME="Fn_Classification_UnitConversionTable"
		Dim objDetailsTable,bReturn,bDoubleClickReturn,iCounter,aObjList,intItemCount,oCounter, rowIndex, bFlag, aMenuList, intCount, sMenu, sText, aMenuList1()
		Dim colCount, i, tab, textArr, columnNumber, columnFoundFlag, intObjectColumnNumber
		Dim colNameArr, bHeaderFoundFlag
		Dim xCord,yCord,DeviceReplay
		columnFoundFlag = False
		bHeaderFoundFlag = False
		intObjectColumnNumber = -1
		Fn_Classification_UnitConversionTable = False

		If sUnitType <> ""  Then
					
				   xCord= JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaObject("LblActiveUnit").GetROProperty ("abs_x")
					yCord=JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaObject("LblActiveUnit").GetROProperty ("abs_y")
			
					Set DeviceReplay = CreateObject("Mercury.DeviceReplay")
					DeviceReplay.MouseClick xCord,yCord,LEFT_MOUSE_BUTTON
					wait 3
		
					' create an object of the table

					'JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaMenu("MenuSelect").SetTOProperty "label", trim(sUnitType)
					JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaMenu(sUnitType).Select
					Set objDetailsTable =JavaDialog("ChangingActiveSystem").JavaTable("ConversionTable")

					If objDetailsTable.Exist(5) = false Then
								Fn_Classification_UnitConversionTable = false
								Exit Function
					End If
					colCount =  objDetailsTable.GetROProperty("cols")
					' Mapping Object column to column number.
			
					If  bHeaderFoundFlag = False Then
							For i = 0 to colCount - 1 
									textArr = split(objDetailsTable.GetColumnName(i),"text=")
									colNameArr = split(textArr(1),",")
									If  trim(sPropertyName) = "Attribute name" Then
												sPropertyName = "Atttribute name" 
									End If
									If trim(colNameArr(0)) = trim(sPropertyName)  then
												intObjectColumnNumber = i
												Exit for
									end if
							 Next
					End If
				
					 If intObjectColumnNumber = -1 Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Column does not exist")	
								JavaDialog("ChangingActiveSystem").JavaButton("Cancel").Click micLeftBtn
								Exit function
					 End If

					Select Case sAction
				
								Case "Rowexist"
										bFlag = false
										'Count number of rows of Table
										bReturn = objDetailsTable.GetROProperty("rows")	
										'Extract the index of row at which the object exist.
										For iCounter=0 to bReturn - 1
											sText = objDetailsTable.GetCellData(iCounter, intObjectColumnNumber )'	Object  column				
												If trim(cstr(sText)) = trim(cstr(sObjectName))  Then
														 bFlag = true
														 Exit for
												End If										
										Next
										If bFlag = false Then
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Fn_Classification_UnitConversionTable : Row with Object "&sObjectName&" does not exist")	
												JavaDialog("ChangingActiveSystem").JavaButton("Cancel").Click micLeftBtn
												Exit function
										Else 
												Fn_Classification_UnitConversionTable = True
										End If
								Case "Rowcellexist"
										If  sExpectedValue = ""  Then
												Fn_Classification_UnitConversionTable = FALSE	 
												JavaDialog("ChangingActiveSystem").JavaButton("Cancel").Click micLeftBtn
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_Classification_UnitConversionTable: Rowcellexist : Incorrect input parameters")
												Exit function
										End If
														
										sText = objDetailsTable.GetCellData(rowNumber, intObjectColumnNumber) ' Object column
										If trim(cstr(sText)) <> cstr(sExpectedValue) Then
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_Classification_UnitConversionTable: Rowcellexist : Expected value is not present")
													JavaDialog("ChangingActiveSystem").JavaButton("Cancel").Click micLeftBtn
													Fn_Classification_UnitConversionTable = False  
													Exit function
										 Else
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_Classification_UnitConversionTable: Rowcellexist : Expected value is present")
										End if 
										Fn_Classification_UnitConversionTable = True 
							End Select				

						JavaDialog("ChangingActiveSystem").JavaButton("OK").Click micLeftBtn
						If Err.Number < 0  Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_Classification_UnitConversionTable: Failed to click on Close Button")
								 Fn_Classification_UnitConversionTable = False  
								Exit function
						End If		
		 End if
End function



'*********************************************************  Function performs  to O peartion on Unit Change Table *********************************************************************
'Function Name  :   Fn_Classification_SearchTableOpeartion
'
' 
'Parameters      :     sAction: Rowcellexist/rowexist
'           				  		sUnitType : metrc/non-metric
'							  		sObjectName : name of object
'									sPropertyName : Column name
'									sExpectedValue : Value to verify
'									rowNumber : row number starts from zero
'									bClose : if  true then close the result panel
'									sOther : feature use

'Return Value     :   True/False
'
'Examples    :      						
'								call Fn_Classification_SearchTableOpeartion("RowSelect","Angle_degrees", "Object ID","","",false,"")
' 								call Fn_Classification_SearchTableOpeartion("ImageExist","000665", "Object ID","","",true,"")
'History:
'          Developer Name                          Date                                                       Rev. No.         Changes Done                   Reviewer 
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'           Prasanna    		                      19-Jan-2011   1.0                         
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function Fn_Classification_SearchTableOpeartion(sAction,sObjectName, sPropertyName,sExpectedValue,rowNumber,bClose,sOther)
	
		GBL_FAILED_FUNCTION_NAME="Fn_Classification_SearchTableOpeartion"
		Dim objDetailsTable,bReturn,bDoubleClickReturn,iCounter,aObjList,intItemCount,oCounter, rowIndex, bFlag, aMenuList, intCount, sMenu, sText, aMenuList1()
		Dim colCount, i, tab, textArr, columnNumber, columnFoundFlag, intObjectColumnNumber
		Dim colNameArr, bHeaderFoundFlag,iRowNumber,iColNumber,sCellData,objContextMenu, aMenu
		columnFoundFlag = False
		bHeaderFoundFlag = False
		intObjectColumnNumber = -1
		Fn_Classification_SearchTableOpeartion = False

		'Click on Search Button 
		If sOther = "" Then
			JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaTab("PropTableTab").Select "Properties"
			JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaButton("Search").WaitProperty "enabled","1",500
			If JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaButton("Search").Exist(3) and JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaButton("Search").GetROProperty("enabled")="1" Then
				JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaButton("Search").Click micLeftBtn
					If Err.Number < 0 Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: To Click on Search Button.")
							Fn_Classification_SearchTableOpeartion = False
							Exit Function
					Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully Clicked on Search Button.")
					End If
					
			End If
		End If
        wait(2)
		Call Fn_ReadyStatusSync(5) 

		'Click on Table pan
		JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaTab("PropTableTab").Select "Table"
		If Err.Number < 0  Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_Classification_SearchTableOpeartion: Failed to Set Table tab")
					 Fn_Classification_SearchTableOpeartion = False  
					Exit function
			 Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_Classification_SearchTableOpeartion: Successfully set Table tab")	
			End If	

		' create an object of the table
		Set objDetailsTable =JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaTable("SearchResultTable")

		If objDetailsTable.Exist(5) = false Then
					Fn_Classification_SearchTableOpeartion = false
					Exit Function
		End If
		wait(5)
		If  cInt(JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaButton("LoadAllSearchResult").GetROProperty ("enabled")) = 1 Then
					JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaButton("LoadAllSearchResult").Click micLeftBtn
					wait 5
		End If
    	colCount =  objDetailsTable.GetROProperty("cols")
		' Mapping Object column to column number.

'		If sPropertyName <> "RowSelect" Then
					If  bHeaderFoundFlag = False Then
						For i = 0 to colCount - 1
							If instr(objDetailsTable.GetColumnName(i), "text=") > 0 Then
								textArr = split(objDetailsTable.GetColumnName(i),"text=")
								colNameArr = split(textArr(1),",")
							End If
							If instr(colNameArr(0), "<nobr>") > 0 Then
								textArr = split(colNameArr(0),"nobr>")
								colNameArr = split(textArr(1),"<")
							End If
							If trim(colNameArr(0)) = trim(sPropertyName)  then
								intObjectColumnNumber = i
								Exit for
							end if
						 Next
					End If
				
					 If intObjectColumnNumber = -1 Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Column does not exist")	
								JavaDialog("ChangingActiveSystem").JavaButton("Cancel").Click micLeftBtn
								Exit function
					 End If
'		End If

		Select Case sAction
	
					Case "Rowexist"
							bFlag = false
							'Count number of rows of Table
							bReturn = objDetailsTable.GetROProperty("rows")	
							'Extract the index of row at which the object exist.
							For iCounter=0 to bReturn - 1
								sText = objDetailsTable.Object.getValueAt(Cint(iCounter), Cint(intObjectColumnNumber)).toString'	Object  column				
									If trim(cstr(sText)) = trim(cstr(sObjectName))  Then
											 bFlag = true
											 Exit for
									End If										
							Next
							If bFlag = false Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Fn_Classification_SearchTableOpeartion : Row with Object "&sObjectName&" does not exist")    									
									Exit function
							Else 
									Fn_Classification_SearchTableOpeartion = True
							End If
				'[TC1123-20161108-21_11_2016-VivekA-Maintenance] - Added from TC1017 - By Poonam C
				Case "RowCount"  ''   TC1015_20150721_VivekA_New Development added case a to get Row count
						'Count number of rows of Table
						bReturn = objDetailsTable.GetROProperty("rows")	
						If bReturn = false Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Fn_Classification_SearchTableOpeartion : failed to fetch  no of rows in  table.")    									
							Exit function
						Else 
						       Fn_Classification_SearchTableOpeartion = bReturn
						End If	
						
				Case "RowCountWithoutLoadAll"  		'   TC1015-2015091500-25_09_2015-AnkitN-New Development added case a to get Row count without Load All button call
							bReturn = objDetailsTable.GetROProperty("rows")
							If bReturn = False Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Fn_Classification_SearchTableOpeartion : failed to fetch  no of rows in  table.")    									
								Exit function
							Else 
							       Fn_Classification_SearchTableOpeartion = bReturn
							End If
				'----------------------------------------------------------------------------------
				Case "Rowcellexist"
							If  sExpectedValue = ""  Then
									Fn_Classification_SearchTableOpeartion = FALSE   									
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_Classification_SearchTableOpeartion: Rowcellexist : Incorrect input parameters")
									Exit function
							End If
											
							bReturn = objDetailsTable.GetROProperty("rows")	
							'Extract the index of row at which the object exist.
							For iCounter=0 to bReturn - 1
								sText = objDetailsTable.Object.getValueAt(Cint(iCounter), Cint(2)).toString'	'if row found
									If trim(cstr(sText)) = trim(cstr(sObjectName))  Then
	                                        sText = objDetailsTable.Object.getValueAt(Cint(iCounter), Cint(intObjectColumnNumber)).toString'	'if row found GET THE DESIRED VALUE
											 Exit for
									End If										
							Next		
									
							
							If trim(cstr(sText)) <> cstr(sExpectedValue) Then
											If cint(iCounter) = cint(bReturn) Then
													iCounter = 0 
											End If
											sText = objDetailsTable.GetCellData(iCounter, intObjectColumnNumber)
										If trim(cstr(sText)) <> cstr(sExpectedValue) Then
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_Classification_SearchTableOpeartion: Rowcellexist : Expected value is not present")										
													Fn_Classification_SearchTableOpeartion = False  
													Exit function
										 Else
													Fn_Classification_SearchTableOpeartion = true  	
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_Classification_SearchTableOpeartion: Rowcellexist : Expected value is present")
										 End if 				
							Else
													Fn_Classification_SearchTableOpeartion = true  	
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_Classification_SearchTableOpeartion: Rowcellexist : Expected value is present")										 
							End if 							
				  Case "RowSelect"
							bFlag = false
							'Count number of rows of Table
							bReturn = objDetailsTable.GetROProperty("rows")	
							'Extract the index of row at which the object exist.
							For iCounter=0 to bReturn - 1
								sText = objDetailsTable.Object.getValueAt(Cint(iCounter), Cint(intObjectColumnNumber)).toString'	Object  column				
									If trim(cstr(sText)) = trim(cstr(sObjectName))  Then
	                                        objDetailsTable.ClickCell iCounter,intObjectColumnNumber,"LEFT"
											 bFlag = true
											 Exit for
									End If										
							Next
							If bFlag = false Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Fn_Classification_SearchTableOpeartion : Row with Object "&sObjectName&" does not exist")    									
									Exit function
							Else 
									Fn_Classification_SearchTableOpeartion = True
							End If
				Case "RowDoubleClick"
							bFlag = false
							'Count number of rows of Table
							bReturn = objDetailsTable.GetROProperty("rows")	
							'Extract the index of row at which the object exist.
							For iCounter=0 to bReturn - 1
								sText = objDetailsTable.Object.getValueAt(Cint(iCounter), Cint(intObjectColumnNumber)).toString'	Object  column				
									If trim(cstr(sText)) = trim(cstr(sObjectName))  Then									
											objDetailsTable.DoubleClickCell iCounter,1,"LEFT"
									 bFlag = true
								Exit for
								End If										
						Next

							If bFlag = false Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Fn_Classification_SearchTableOpeartion : Row with Object "&sObjectName&" does not exist")    									
									Fn_Classification_SearchTableOpeartion = false
									Exit function
							Else 
									Fn_Classification_SearchTableOpeartion = True
							End If

					'- - - - -  - - - -  - - - -  - - - -  - - - - Added By  Pooja S-  7 March -2012  - - - -  - - - -  - - - -  - - - -  - - - -  - - - - 
					Case "VerifySortedVaues"
							Dim arrayVal()
							ReDim arrayVal(999)
							bFlag = false

							'Count number of rows of Table
							bReturn = objDetailsTable.GetROProperty("rows")	
							'Extract the index of row at which the object exist.
							For iCounter=0 to bReturn - 1
									sText = objDetailsTable.Object.getValueAt(Cint(iCounter), Cint(intObjectColumnNumber)).toString'	Object  column				
									arrayVal(iCounter)=sText						
							Next

							For iCounter=0 to Ubound(sExpectedValue)
									If 	arrayVal(iCounter)=sExpectedValue(iCounter) Then
												bFlag = true
									Else
												bFlag = false
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Fn_Classification_SearchTableOpeartion : Row with Object "&sObjectName&" does not exist")    									
												Exit function
									 End If
							Next

							If  bFlag = true Then
								Fn_Classification_SearchTableOpeartion = True
							End If
					'- - - - -  - - - -  - - - -  - - - -  - - - -  - - - -  - - - -  - - - -  - - - -  - - - -  - - - -  - - - -  - - - -  - - - -  - - - -  -  - - - -  -  - - - -  -  - - - - 
				Case "ImageExist"
											
										  	bFlag = false
											'Count number of rows of Table
											bReturn = objDetailsTable.GetROProperty("rows")	
											'Extract the index of row at which the object exist.
											For iCounter=0 to bReturn - 1
												sText = objDetailsTable.Object.getValueAt(Cint(iCounter), Cint(intObjectColumnNumber)).toString'	Object  column					
													If trim(cstr(sText)) = trim(cstr(sObjectName))  Then
															 iRowNumber = iCounter     ' Get the row number 
															 bFlag = true
															 Exit for
													End If										
											Next

											If bFlag = false Then
														Fn_Classification_SearchTableOpeartion = false
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Fn_Classification_SearchTableOpeartion : Row Does Not Exist for Value  Specified ["+cstr(sObjectName)+"]") 													
														Exit function
											 Else
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Fn_Classification_SearchTableOpeartion : Row Found for Value  Specified ["+cstr(sObjectName)+"]") 													
											End If
											'Get the Column name 
												colCount =  objDetailsTable.GetROProperty("cols")
											' 	Mapping Object column to column number.                                     												
												For i = 0 to colCount - 1 
														textArr = split(objDetailsTable.GetColumnName(i),"defaultIcon=")
														colNameArr = split(textArr(1),",")							
														If instr(1,lcase(colNameArr(0)),"classifying_16.png") then
																	iColNumber = i
																	Exit for
														end if
												 Next

'											Get the cell data - This returns the Image path exist in Cell
											sText = objDetailsTable.Object.getValueAt(Cint(iRowNumber), Cint(iColNumber)).toString'	Object  column					

											If sText = "" Then
													Fn_Classification_SearchTableOpeartion = false
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Fn_Classification_SearchTableOpeartion : The Item Image does not exist") 													
													Exit function
											Else 
													Fn_Classification_SearchTableOpeartion = True
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Fn_Classification_SearchTableOpeartion : The Item Image does not exist") 														
											End If
											
			    Case "SelectAll"
										Call Fn_ReadyStatusSync(3)
										JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaTable("SearchResultTable").SelectRow(0)
										JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaTable("SearchResultTable").Object.selectAll
										Fn_Classification_SearchTableOpeartion = True
			    
				Case "PopupMenuEnabled"									 
										If sExpectedValue <> "" Then
												aMenu = split(sExpectedValue,":")
												Select Case cInt(Ubound(aMenu))
													Case 0
														Set objContextMenu = JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaMenu("MenuSelect")
														objContextMenu.SetTOProperty "label", sExpectedValue
														objDetailsTable.ClickCell rowNumber,intObjectColumnNumber, "RIGHT","NONE"
														Wait SISW_MICRO_TIMEOUT
														If objContextMenu.CheckProperty (sExpectedValue, "Exists",10) Then
															Fn_PSE_BOMTable_NodeOperationExt = objContextMenu.CheckProperty (sExpectedValue, "Enabled",10)
														End IF
'														objDetailsTable.ClickCell rowNumber,intObjectColumnNumber, "LEFT","NONE"
														Set objContextMenu = nothing
												End Select
												Fn_Classification_SearchTableOpeartion = True
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [Fn_PSE_BOMTable_NodeOperationExt] Popup Menu ["+ sExpectedValue +"] Selected Sucessfully")
										Else
											Fn_Classification_SearchTableOpeartion = False
											Exit function
										End If
				Case "GetCelldata"
				
										sCellData=objDetailsTable.GetCellData(rowNumber,intObjectColumnNumber)
										If instr(1,sCellData,sObjectName)  Then
											Fn_Classification_SearchTableOpeartion = sCellData
											
										Else 
											Fn_Classification_SearchTableOpeartion = False
											Exit function																
										End If
																				   

		End Select				

		If bClose = true Then
			JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaTab("PropTableTab").Select "Properties"
			If Err.Number < 0  Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_Classification_SearchTableOpeartion: Failed to Set Properties tab")
					 Fn_Classification_SearchTableOpeartion = False  
					Exit function
			 Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_Classification_SearchTableOpeartion: Successfully set Properties tab")	
			End If		
		End If
End function

'*********************************************************  Fucntion to change the unit of attribute *********************************************************************
'Function Name  :   Fn_Classification_AttributeUnitText
'
' 
'Parameters      :     sAction: Change/Verify
'           				  		sTextSelect : Unit name to be clicked
'							  		sUnitName : Unit name to be Select from popup
'									sDetails : feature use

'Return Value     :   True/False
'
'Examples    :      						
'								Fn_Classification_AttributeUnitText("Change","<html><a href=''>m</a></html>","<html><body>mm</body></html>","")
' 
'History:
'          Developer Name                          Date                                                       Rev. No.         Changes Done                   Reviewer 
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'           Prasanna    		                      25-Jan-2011   1.0                         
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------


Public Function Fn_Classification_AttributeUnitText(sAction,sTextSelect,sUnitName,sDetails)

		GBL_FAILED_FUNCTION_NAME="Fn_Classification_AttributeUnitText"
		Dim objSelectTypeMenu,intNoOfObjects,bFound,objSelectType,iVal

		Select Case sAction
				
					 Case "Change"
		
					'Click on current unit 
					If sTextSelect <> "" Then
'								JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaStaticText("UnitStaticText").SetTOProperty "label",trim(sTextSelect)
'								JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaStaticText("UnitStaticText").SetTOProperty "attached text",trim(sTextSelect)
								'Click on Unit Name
								If sDetails="" then
										Set objSelectType = description.Create()
												objSelectType("Class Name").value = "JavaStaticText"
												Set objClassAdmin =JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").ChildObjects(objSelectType)
												
												iCount=JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").ChildObjects(objSelectType).Count
												For iVal=0 to iCount-1
																	sTextValue=objClassAdmin(iVal).GetRoProperty("attached text")
																	If instr(1,sTextValue,sTextSelect)>0 Then
																		wait 1
																		'objClassAdmin(iVal).SetToProperty "Index",iVal
																		objClassAdmin(iVal).Click 1,1
																	Exit For
																	End If
												Next
												
												If Err.Number < 0 then
															Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Click on Unit ["+cstr(sTextSelect)+"]")
															Fn_Classification_AttributeUnitText = false
															Exit Function
												Else
															Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully  Clicked on Unit ["+cstr(sTextSelect)+"]")
												End if
								  Else
												Set objSelectType = description.Create()
												objSelectType("Class Name").value = "JavaStaticText"
												Set objClassAdmin =JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").ChildObjects(objSelectType)
												iCount=JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").ChildObjects(objSelectType).Count
												
												For iVal=0 to iCount-1
																	sTextValue=objClassAdmin(iVal).GetRoProperty("attached text")
																	If instr(1,sTextValue,sTextSelect)>0 Then				
																				
																				iVal=cint(iVal)+cint(sDetails)
'																				objClassAdmin(iVal).SetToProperty "Index",iVal
																				objClassAdmin(iVal).Click 1,1
																				Exit for 																			 
																	Exit For
																	End If
												Next
												If Err.Number < 0 then
																		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Click on Unit ["+cstr(sTextSelect)+"]")
																		Fn_Classification_AttributeUnitText = false
																		Exit Function
															Else
																		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully  Clicked on Unit ["+cstr(sTextSelect)+"]")
															End if
								End If
					End If
			
					'Select the desired unit from pop-up menu
					bFound = false
					Set objSelectTypeMenu=Description.Create()
					objSelectTypeMenu("Class Name").value = "JavaMenu"
					 Set intNoOfObjects = JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").ChildObjects(objSelectTypeMenu)
					  For i = 0 to intNoOfObjects.count-1			   
						   if instr(1,intNoOfObjects(i).getROProperty("label"),sUnitName)   >   0 then
									intNoOfObjects(i).Select
									wait 1
									bFound = true
									Exit for
						   End If
					 Next			
			
					If bFound = true Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully  Selected Unit ["+cstr(sUnitName)+"]")
								Fn_Classification_AttributeUnitText = true
					Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Unit ["+cstr(sUnitName)+"]")					
								Fn_Classification_AttributeUnitText = false
								Exit Function
					End If
			Case "Verify"
							JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaStaticText("UnitStaticText").SetTOProperty "label",trim(sTextSelect)
							JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaStaticText("UnitStaticText").SetTOProperty "attached text",trim(sTextSelect)
							If JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaStaticText("UnitStaticText").Exist(5) Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully  Verified that Unit ["+cstr(sUnitName)+"] is present.")
									 Fn_Classification_AttributeUnitText = true
							Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Verify that Unit ["+cstr(sUnitName)+"] is not present.")
									 Fn_Classification_AttributeUnitText = false
									 Exit Function			
							End If		
		End Select
End Function

'''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'''/$$$$
'''/$$$$   FUNCTION NAME   : Fn_Classification_TableResultPrintOperations(sAction,bSave,bOpen,sFilePath,sInfo,aValues,bClose,bFileClose)
'''/$$$$
'''/$$$$   DESCRIPTION        :  This function will perform operations on the print dialog for the results generated in tabular format 
'''/$$$$
'''/$$$$    PARAMETERS      :   1.) sAction : Action to be performed
'''/$$$$ 									    2.) bSave : To save a file
'''/$$$$									   3.) sFilePath : Valid File Path with proper extension to save and to open as well
'''/$$$$									  4.) sInfo : For Future Use
'''/$$$$									 5.) aValues : array to verify values
'''/$$$$									 6.) bClose : a Boolean parameter to close the print dialog
'''/$$$$									 7.) bFileClose : Close the opened File
'''/$$$$
'''/$$$$	Return Value : 			Value of the static text if it exists and False if the static text does not exist
'''/$$$$
'''/$$$$    Function Calls       :   Fn_WriteLogFile()
'''/$$$$										
'''/$$$$
'''/$$$$	HISTORY           :   AUTHOR                 DATE        VERSION
'''/$$$$
'''/$$$$    CREATED BY     :   SHREYAS           29/01/2011         1.0
'''/$$$$
'''/$$$$    REVIWED BY     :  Prasanna			      29/01/2011         1.0
'''/$$$$
'''/$$$$    EXAMPLE          :  bReturn=Fn_Classification_TableResultPrintOperations("HTML",true,true,"D:\adas6","","ico_1234",true,true)
'''/$$$$									bReturn=Fn_Classification_TableResultPrintOperations("PasteContents&SaveFile","","","D:\Shreyas","","","","")
'''/$$$$									bReturn=Fn_Classification_TableResultPrintOperations("VerifyFileContents","","","D:\Shreyas","","ico_1234","","")
'''/$$$$									bReturn=Fn_Classification_TableResultPrintOperations("TEXT",true,true,"D:\adas6","","ico_1234",true,true)
'''/$$$$									bReturn=Fn_Classification_TableResultPrintOperations("PrintFormat","","","","text:Delimiter:on",",","",true)
'''/$$$$									
'''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Public Function Fn_Classification_TableResultPrintOperations(sAction,bSave,bOpen,sFilePath,sInfo,aValues,bClose,bFileClose)

	GBL_FAILED_FUNCTION_NAME="Fn_Classification_TableResultPrintOperations"
   Dim WshShell, iCounter, aProperties,bFlag,iRowCount,iCount,jCount,iColCount,strLine,sLength1,sLength2
   Dim sJEditorPaneText,arrJEditorPaneLine,sCompareText, objPrintDialog
   Fn_Classification_TableResultPrintOperations=false

		Set objPrintDialog = JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaDialog("Print")
		Set objPrintFormat = JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaDialog("Print Format")
		Set objSave = JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaDialog("Print").JavaDialog("Save")
	
		Err.Clear
				Select Case sAction

				Case "HTML"

					If  objPrintDialog.Exist(5)=false Then
						JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaButton("PrintSearchResult").Click micLeftBtn
					End if

				If objPrintDialog.JavaCheckBox("HTML").GetROProperty("value")=0 then
					
					objPrintDialog.JavaCheckBox("HTML").Set "ON"
				End If

				If bSave=true Then
					objPrintDialog.JavaButton("save_16").Click micLeftBtn
					wait 5
					objSave.JavaEdit("File name:").Set sFilePath+".html"
					objSave.JavaButton("Save").Click  
					wait(5)
				End If

				If bOpen=true Then
					SystemUtil.Run sFilePath+".html"
					wait (5)						
				End If

				If instr(1,aValues,",")>1 Then
						aProperties=split(aValues,",",-1,1)
						For iCounter=0 To Ubound(aProperties)
									bFlag = False
									iRowCount=Browser("Browser").Page("Page").WebTable("PropertyTable").RowCount
									For iCount=0 to iRowCount
										   iColCount=Browser("Browser").Page("Page").WebTable("PropertyTable").GetROProperty("cols")
											For jCount=0 to iColCount-1
												sValue=	Browser("Browser").Page("Page").WebTable("PropertyTable").GetCellData (iCount,jCount)
												If instr(1,sValue, aProperties(iCounter))>0 Then
													bFlag=True
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Verified that the node " + aProperties(iCounter)+" exists in the HTML Table")												
													Exit For
												End If
											Next
														
									 Next
									Exit For
						  Next

				else
						bFlag = False
						iRowCount=Browser("Browser").Page("Page").WebTable("PropertyTable").RowCount
							
						For iCount=0 to iRowCount
								iColCount=Browser("Browser").Page("Page").WebTable("PropertyTable").GetROProperty("cols")
								For jCount=1 to iColCount
									  sValue=	Browser("Browser").Page("Page").WebTable("PropertyTable").GetCellData (iCount,jCount)
									  If instr(1,sValue, aValues)<>0Then
												bFlag=True
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Verified that the node " + aValues+" exists in the HTML Table")
												Exit For
										End If
								Next
										
						 Next
				 End if

				If bFileClose=true Then
						Browser("Browser").Close
				 End If

				If bFlag=True Then
						Fn_Classification_TableResultPrintOperations=True
				Else
						Fn_Classification_TableResultPrintOperations=False
				End If

				Case "TEXT"

									If  objPrintDialog.Exist(5)=false Then
											JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaButton("PrintSearchResult").Click micLeftBtn
									End if
				
									If objPrintDialog.JavaCheckBox("Text").GetROProperty("value")=0 then
												objPrintDialog.JavaCheckBox("Text").Set "ON"
									End If
				
									If bSave=true Then
										objPrintDialog.JavaButton("save_16").Click micLeftBtn
										wait(5)
										objSave.JavaEdit("File name:").Set sFilePath+".txt"
										objSave.JavaButton("Save").Click
										 wait(5)
									End If

									If bOpen=true Then
										SystemUtil.Run sFilePath+".txt"
										Set objFSO = CreateObject("Scripting.FileSystemObject")
										Set objFile = objFSO.OpenTextFile(sFilePath+".txt")
										Do Until objFile.AtEndOfStream
												strLine = objFile.ReadLine
										Loop
										objFile.Close						
									End If
									If instr(1,aValues,",")>1 Then
										 aProperties=split(aValues,",",-1,1)
										 For iCounter=0 To Ubound(aProperties)
													bFlag = False
													If instr(1,strLine,aProperties(iCounter) ) >0 Then
																bFlag=True
																Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Verified that the node " +aValues+" exists in the file"+sFilePath+".txt")
													End If
													 Next
									Else
										  bFlag = False
									If instr(1,strLine,aValues ) >0 Then
											bFlag=True
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Verified that the node " +aValues+" exists in the file"+sFilePath+".txt")
											End If
									End if

									If bFileClose=true Then
											Window("Notepad").Close
									End If
				
									If bFlag=True Then
											Fn_Classification_TableResultPrintOperations=True
									Else
											Fn_Classification_TableResultPrintOperations=False
									End If


			Case "PasteContents&SaveFile"
									bFlag=False
				
									'click on the copy button
									JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaButton("CopyTableData").Click micLeftBtn

									If Window("Notepad").Exist then
												Set WshShell = CreateObject("WScript.Shell")
															WshShell.SendKeys "^(v)"
															WshShell.SendKeys "^(s)"
													Set WshShell = nothing
												If Window("Notepad").Dialog("Save As").Exist then
													'set the file name to be saved
												Window("Notepad").Dialog("Save As").WinEdit("File name:").Set sFilePath
									
												'click on save button
												Window("Notepad").Dialog("Save As").WinButton("Save").Click
									
												bFlag=True
																Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully saved the file at the location ["+sFilePath+"]")
												End if
									Else
												SystemUtil.Run "notepad"
												Window("Notepad").Maximize
												Set WshShell = CreateObject("WScript.Shell")
												WshShell.SendKeys "^(v)"
												WshShell.SendKeys "^(s)"
												Set WshShell = nothing
												If Window("Notepad").Dialog("Save As").Exist then
																'set the file name to be saved
															Window("Notepad").Dialog("Save As").WinEdit("File name:").Set sFilePath
																		'click on save button
														Window("Notepad").Dialog("Save As").WinButton("Save").Click
															bFlag=True
																			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully saved the file at the location ["+sFilePath+"]")
												End if
									End if

									If bFlag=True Then
													Fn_Classification_TableResultPrintOperations=True
										Else
													Fn_Classification_TableResultPrintOperations=False
										End If

		Case "VerifyFileContents"

					Set objFSO = CreateObject("Scripting.FileSystemObject")
					Set objFile = objFSO.OpenTextFile(sFilePath+".txt")
					Do Until objFile.AtEndOfStream
					 strLine = objFile.ReadLine
					Loop
					objFile.Close

						If instr(1,aValues,",")>1 Then
									aProperties=split(aValues,",",-1,1)
										For iCounter=0 To Ubound(aProperties)
													bFlag = False
											If instr(1,strLine,aProperties(iCounter) ) >0 Then
																		bFlag=True
																		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Verified that the node " + aValues+" exists in the file"+sFilePath+".txt")
																End If
																 Next
											Else
																bFlag = False
											If instr(1,strLine,aValues ) >0 Then
																		bFlag=True
																		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Verified that the node " +aValues+" exists in the file"+sFilePath+".txt")
																End If
											End if
						
											 If bFileClose=true Then
																Window("Notepad").Close
											 End If

											 If bFlag=True Then
													Fn_Classification_TableResultPrintOperations=True
											Else
													Fn_Classification_TableResultPrintOperations=False
											End If

						'==================================================================================================================================
						Case "VerifyHTMLContents"

						If bOpen=true Then
							objPrintDialog.JavaButton("openinbrowser_16").Click micLeftBtn
							wait(5)
						End If

''						Browser("browser").Page("Page").WebElement("HeadTitle").SetTOProperty "innertext",sInfo
''						Browser("browser").Page("Page").WebElement("HeadTitle").SetTOProperty "html tag","H1"
'						Browser("Classification").Page("Classification").WebElement("HTMLText").SetTOProperty "innertext",sInfo
'						wait(5)
'						
''						If Browser("browser").Page("Page").WebElement("HeadTitle").Exist(5) then
'						If Browser("Classification").Page("Classification").WebElement("HTMLText").Exist(5) then
''								Browser("Classification").Page("Classification").WebElement("HTMLText").Highlight
'								 Fn_Classification_TableResultPrintOperations=True
'								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Verified HTML Page Contents ["+sInfo+"] ")            
'						Else
'								Fn_Classification_TableResultPrintOperations=False
'						End If
'
'						If bFileClose=true Then
'							Browser("Browser").Close
'						End If
				'If instr(1,JavaWindow("ClassificationMainWin").JavaObject("Browser").Object.getText(),sInfo) > 0 then
				If instr(1,Browser("Classification").Page("Classification").WebElement("ClassifiedItemData").GetROProperty("innertext"),sInfo) > 0 then
							Fn_Classification_TableResultPrintOperations=true
                 else
							Fn_Classification_TableResultPrintOperations=false
				 end if



						
				' =================================================================================================================================
					Case "VerifyTEXTContents"
	
					sJEditorPaneText =objPrintDialog.JavaEdit("JEditorPane").Object.getText()
					arrJEditorPaneLine = split(trim(sJEditorPaneText),Chr(10),-1,1)     ' returns text split Linewise
					
	'				For iCounter=0 to 6		
							iCounter=aValues
							sCompareText=cstr(arrJEditorPaneLine(iCounter))
	
							If instr(1,sCompareText, sInfo )>0 Then
									Fn_Classification_TableResultPrintOperations=True
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Verified [ " + sText+" ] exists in the JEditorPane Text Panel")
	'								Exit For 
							Else
									Fn_Classification_TableResultPrintOperations=False			
							End If
	'				Next

				' =================================================================================================================================

						Case "PrintFormat"
						  Err.Clear	
						  bFlag=False
						  aCheck=split(sInfo,":",-1,1)
							  If lcase(aCheck(0))="html" Then
							   objPrintDialog.JavaCheckBox("HTML").Set "ON"
							  Elseif lcase(aCheck(0))="text" Then
							   objPrintDialog.JavaCheckBox("Text").Set "ON"
							  End If
						  'Invoke the Search type select window and convert it into a dialog 
						  objPrintDialog.JavaCheckBox("format_16").Set "OFF"
						  wait 2
						   objPrintDialog.JavaCheckBox("format_16").Set "ON"
						   Set objSelectType = description.Create()
							objSelectType("Class Name").value = "JavaStaticText"
							Set objClassAdmin =JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").ChildObjects(objSelectType)
							sStaticCount=JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").ChildObjects(objSelectType).count
							For iCount=0 to sStaticCount-1
							   sTextValue=objClassAdmin(iCount).GetRoProperty("attached text")
							   If sTextValue="Print Format" Then
								objClassAdmin(iCount).DblClick 5,5,"LEFT"
								Exit for
							   End If
							Next
						   If objPrintFormat.Exist Then
							
						   Select Case aCheck(1)
							Case "Title"
							  If lcase(aCheck(2))="on" and aValues<>""  Then
									   objPrintFormat.JavaCheckBox("Title").set "ON"
									   objPrintFormat.JavaEdit("Title:").Set aValues
									   If Err.Number < 0 then	  
											  bFlag=False
											  Exit Function
										Else
											  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully  Set The Value ["+(aValues)+"]")
											  bFlag=True
										End if
										'click on update button
										objPrintFormat.JavaButton("Update").Click micLeftBtn
										'verify if the changes are saved
		'								sValue=objPrintDialog.JavaEdit("JEditorPane").GetROProperty("value")
		'								 If instr(1,sValue, aValues)>0 Then
		'								 bFlag=True
		'								 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Verified that the node " + aValues+" exists and the formatting is successful")            
		'								End If
							  Else
									objPrintFormat.JavaCheckBox("Title").set "OFF"
								    bFlag=True
									objPrintFormat.JavaEdit("Title:").Set aValues
									If Err.Number < 0 then				  
										  bFlag=False
										  Exit Function
								   Else
										 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully  Set The Value ["+(aValues)+"]")
										  bFlag=True
								   End if
								'click on update button
								objPrintFormat.JavaButton("Update").Click micLeftBtn
								'verify if the changes are saved
'								sValue=objPrintDialog.JavaEdit("JEditorPane").GetROProperty("value")
'								 If instr(1,sValue, aValues)>0 Then
'								  bFlag=False
'								  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The Value" + aValues+" still exists and the formatting is not successful") 
'								 Else
'								  bFlag=True
'								  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verified that the formatting is successful for Title")       
'								End If
					 
							   End If
								If bFileClose=true Then
								objPrintFormat.Close
							   End If
							Case "Date"
							  If lcase(aCheck(2))="on" then
							   objPrintFormat.JavaCheckBox("Date").Set "ON"
								If Err.Number < 0 then
									  bFlag=False
									  Exit Function
								   End if
								'click on update button
								objPrintFormat.JavaButton("Update").Click micLeftBtn
								'verify if the changes are saved
'								sValue=objPrintDialog.JavaEdit("JEditorPane").GetROProperty("value")
'								aDate=split(now,"/",-1,1)
'								sMonthName=MonthName(aDate(0), True)
'								 If instr(1,sValue,sMonthName  )>0 Then
'								  bFlag=True
'								  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Verified that the node "+ cstr(year(now))+" exists and the formatting is successful")            
'								End If
							  Else
								  objPrintFormat.JavaCheckBox("Date").set "OFF"
								    bFlag=True
'							   If Err.Number < 0 then
'									  bFlag=False
'									  Exit Function
'								   End if
								'click on update button
								objPrintFormat.JavaButton("Update").Click micLeftBtn
								'verify if the changes are saved
								sValue=objPrintDialog.JavaEdit("JEditorPane").GetROProperty("value")
								aDate=split(now,"/",-1,1)
								sMonthName=MonthName(aDate(0), True)
'								 If instr(1,sValue,sMonthName  )>0 Then
'								  bFlag=False
'								  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The Value" +cstr(year(now))+" still exists and the formatting is not successful")
'								 Else
'								  bFlag=True
'								  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verified that the formatting is successful for Date")     
'								 End If
							   End If
								If bFileClose=true Then
								objPrintFormat.Close
							   End If
							 Case "ObjectCount"
							  If lcase(aCheck(2))="on" then
							   objPrintFormat.JavaCheckBox("Object Count").Set "ON"
								If Err.Number < 0 then
										bFlag=False
										Exit Function
										End if
										'click on update button
										objPrintFormat.JavaButton("Update").Click micLeftBtn
										'verify if the changes are saved
	'									sValue=objPrintDialog.JavaEdit("JEditorPane").GetROProperty("value")
	'									 If instr(1,sValue,"objects" )>0 Then
	'									 bFlag=True
	'									 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Verified that the value [objects] exists and the formatting is successful")            
	'									End If
								 Else
										objPrintFormat.JavaCheckBox("Object Count").set "OFF"
										bFlag=True
'							   
'							   If Err.Number < 0 then
'									  bFlag=False
'									  Exit Function
'								   End if
								'click on update button
								objPrintFormat.JavaButton("Update").Click micLeftBtn
								'verify if the changes are saved
								sValue=objPrintDialog.JavaEdit("JEditorPane").GetROProperty("value")
								 If instr(1,sValue, "objects" )>0 Then
									 bFlag=False
									 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The value [objects]  still exists and the formatting is not successful")    
								 Else
									  bFlag=True
									  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verified that the formatting is successful for Object Count")   
								End If
							   End If
							   If bFileClose=true Then
								objPrintFormat.Close
							   End If
							Case "Column Allignment" '(only applicable when the 'Text' is checked in the print dialog')
							 If lcase(aCheck(2))="on" then
							   objPrintFormat.JavaCheckBox("Column Alignment").Set "ON"
								If Err.Number < 0 then
									  bFlag=False
									  Exit Function
								   End if
								'click on update button
								objPrintFormat.JavaButton("Update").Click micLeftBtn
'								'verify if the changes are saved
'								sValue=objPrintDialog.JavaEdit("JEditorPane").GetROProperty("value")
'								sLength1=len(sValue)
'								'now uncheck the Column allignment checkbox and check the length
'								objPrintFormat.JavaCheckBox("Column Alignment").Set "OFF"
'								'click on update button
'								objPrintFormat.JavaButton("Update").Click micLeftBtn
'								sValue=objPrintDialog.JavaEdit("JEditorPane").GetROProperty("value")
'								sLength2=len(sValue)
'								If sLength1>sLength2 Then
'								  bFlag=True
'								  objPrintFormat.JavaCheckBox("Column Alignment").Set "ON"
'								  objPrintFormat.JavaButton("Update").Click micLeftBtn
'								  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Verified that the formatting is successful for Column Alignment")
'								Else
'								  bFlag=False
'								  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The formatting is not successful for Column Alignment")  
'								End If
							 Else
					   
								sValue=objPrintDialog.JavaEdit("JEditorPane").GetROProperty("value")
								sLength1=len(sValue)
								objPrintFormat.JavaCheckBox("Column Alignment").set "OFF"
								bFlag=True
'							   
'								If Err.Number < 0 then
'									  bFlag=False
'									  Exit Function
'								   End if
								'click on update button
								objPrintFormat.JavaButton("Update").Click micLeftBtn
								sValue=objPrintDialog.JavaEdit("JEditorPane").GetROProperty("value")
								sLength2=len(sValue)
								If sLength1<sLength2 Then
								  bFlag=False
								  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The formatting is not successful for Column Alignment") 
								Else
								  bFlag=True
								  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verified that the formatting is successful for Column alignment")  
								End If
							 End If
								If bFileClose=true Then
								 objPrintFormat.Close
								End If
						   Case "Delimiter" '(only applicable when the 'Text'  is checked in the print dialog')
							   If lcase(aCheck(2))="on" and aValues<>""  Then
							   
					'           objPrintFormat.JavaEdit("Delimiter").Activate
							   objPrintFormat.JavaEdit("Delimiter").Set (aValues)
								   If Err.Number < 0 then
									  bFlag=False
									  Exit Function
								   Else
									  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully  Set The Value ["+(aValues)+"]")
									  bFlag=True
								   End if
								'click on update button
								objPrintFormat.JavaButton("Update").Click micLeftBtn
								'verify if the changes are saved
'								sValue=objPrintDialog.JavaEdit("JEditorPane").GetROProperty("value")
'								 If instr(1,sValue, aValues)>0 Then
'								 bFlag=True
'								 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Verified that the node [" + aValues+"] exists and the formatting is successful")            
'								End If
'							  Else
'								bFlag=False
							  End If
					  
								If bFileClose=true Then
								 objPrintFormat.Close
								End If
							 End select
							  If bFlag=True Then
							   Fn_Classification_TableResultPrintOperations=True
							  Else
							   Fn_Classification_TableResultPrintOperations=False
							 End If
							End If
					  End Select
					  If  bClose=true Then
					   objPrintDialog.Close
					   Set objPrintDialog = Nothing
					   Set objPrintFormat = Nothing
					   Set objSave = Nothing
					  End If
End Function

 
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''''/$$$$
''''/$$$$   FUNCTION NAME   : Fn_Classification_XMLExport(sAction,sTab,sTargetApplication,sOutputFile,sInfo,sButton)
''''/$$$$
''''/$$$$   DESCRIPTION        :  This function will perform operations on the print dialog for the results generated in tabular format 
''''/$$$$
''''/$$$$    PARAMETERS      :   1.) sAction : Action to be performed
''''/$$$$ 									    2.) sTab : Tab To Be selected
''''/$$$$									   3.) sTargetApplication : Target appliction to be selected
''''/$$$$									  4.) sOutputFile : Output file for the input
''''/$$$$									 5.) sInfo : For future use
''''/$$$$									 6.) sButton : button to be clicked
''''/$$$$									
''''/$$$$
''''/$$$$	Return Value : 			To perform export operation in the Classification Application
''''/$$$$
''''/$$$$    Function Calls       :   Fn_WriteLogFile()
''''/$$$$										
''''/$$$$
''''/$$$$	HISTORY           :   AUTHOR                 DATE        VERSION
''''/$$$$
''''/$$$$    CREATED BY     :   SHREYAS           03/02/2011         1.0
''''/$$$$
''''/$$$$    REVIWED BY     :  Prasanna			    03/02/2011         1.0
''''/$$$$
''''/$$$$    Modify By                 Deepali               26/07/12               1.1
''''/$$$$
''''/$$$$
''''/$$$$    EXAMPLE          :  
''''/$$$$									
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

Public Function Fn_Classification_XMLExport(sAction,sTab,sTargetApplication,sOutputFile,sInfo,sButton)

	GBL_FAILED_FUNCTION_NAME="Fn_Classification_XMLExport"
   Dim WshShell, iCounter, aProperties,bFlag,iRowCount,iCount,jCount,iColCount,strLine,sStaticCount,aCheck,aDate,sMonthName
   Dim objXportDialog

   Fn_Classification_XMLExport=false

	Set objXportDialog=Fn_SISW_Classification_GetObject("XML Export")
   
   'check the existence of the XML export dialog
   If  objXportDialog.Exist(5) =False Then
	   bReturn = Fn_ToolbatButtonClick("Export Objects")
						 If Err.Number < 0 Then
									Fn_Classification_XMLExport = False
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to select tab [ " + sTab + "] of XML Import dialog" ) 
									Exit Function 
							End if
			Call Fn_ReadyStatusSync(3)
  End if


		 If  objXportDialog.Exist(5) Then

				Select Case sAction

								Case "Export"
									bFlag=False
						
											'activate the specified tab
						
											 objXportDialog.JavaTab("Tab").Select sTab						
											  If Err.Number < 0 Then
													Fn_Classification_XMLExport = False
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to select tab [ " + sTab + "] of XML Import dialog" ) 
													Exit Function 
											End if
						
											'slect the value from the target application list						
											objXportDialog.JavaList("TargetApplication").Select sTargetApplication						
											If Err.Number < 0 Then
													Fn_Classification_XMLExport = False
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to select tab [ " + sTab + "] of XML Import dialog" ) 
													Exit Function 
											End if
						
											'Set the output file						
											objXportDialog.JavaEdit("Output File").Set sOutputFile						
											If Err.Number < 0 Then
													Fn_Classification_XMLExport = False
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to select tab [ " + sTab + "] of XML Import dialog" ) 
													Exit Function 
											End if
											bFlag=True													
				End Select

				If lcase(sButton)="ok" Then
						objXportDialog.JavaButton("OK").Click micLeftBtn
						If Err.Number < 0 Then
								Fn_Classification_XMLExport = False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to select tab [ " + sTab + "] of XML Import dialog" ) 
								Exit Function 
						End if
				Elseif lcase(sButton)="cancel" Then
						objXportDialog.JavaButton("Cancel").Click micLeftBtn
						If Err.Number < 0 Then
								Fn_Classification_XMLExport = False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to select tab [ " + sTab + "] of XML Import dialog" ) 
								Exit Function 
						End if
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "No Button Needs To Be Clicked" ) 
				End If

				'handle the Export successful Dialog by clicking Ok button
				JavaWindow("ClassificationMainWin").JavaWindow("Export ICO").JavaButton("OK").Click micLeftBtn
				If Err.Number < 0 Then
								Fn_Classification_XMLExport = False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to select tab [ " + sTab + "] of XML Import dialog" ) 
								Exit Function 
				End if


				If bFlag=True Then
					Fn_Classification_XMLExport=True
				Else
					Fn_Classification_XMLExport=False
				End If

		End if

End function


''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''''/$$$$
''''/$$$$   FUNCTION NAME   : Fn_Classification_CheckClassificationProperty(sProperty)
''''/$$$$
''''/$$$$   DESCRIPTION        :  This function will check the classification properties
''''/$$$$
''''/$$$$    PARAMETERS      :   1.) sProperty : Property to be checked
''''/$$$$ 									 
''''/$$$$									
''''/$$$$
''''/$$$$	Return Value : 			True / False
''''/$$$$
''''/$$$$    Function Calls       :   Fn_WriteLogFile()
''''/$$$$										
''''/$$$$
''''/$$$$	HISTORY           :   AUTHOR                 DATE        VERSION
''''/$$$$
''''/$$$$    CREATED BY     :   SHREYAS           09/02/2011         1.0
''''/$$$$
''''/$$$$    REVIWED BY     :  Prasanna			    09/02/2011         1.0
''''/$$$$
''''/$$$$    EXAMPLE          :  
''''/$$$$									
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Public Function Fn_Classification_CheckClassificationProperty(sProperty)

	GBL_FAILED_FUNCTION_NAME="Fn_Classification_CheckClassificationProperty"
   Dim objSelectType,objClassAdmin,sValue,sStaticCount,bFlag
	bFlag=false
				Set objSelectType = description.Create()
				objSelectType("Class Name").value = "JavaStaticText"
				Set objClassAdmin =JavaWindow("MyTeamcenter").ChildObjects(objSelectType)
				sStaticCount=JavaWindow("MyTeamcenter").ChildObjects(objSelectType).count
				For iCount=0 to sStaticCount-1
						   sValue=objClassAdmin(iCount).GetRoProperty("attached text")
						   If Trim(cstr(sValue)) = Trim(cstr(sProperty)) Then
								bFlag=True
								Exit for
							Elseif isNumeric(sValue) and isNumeric(sProperty)  Then'And cint(sValue) = cint(sProperty) Then
								If cint(sValue) = cint(sProperty) Then
									bFlag=True
								Exit for
								End If
'								bFlag=True
'								Exit for
						   End If
				Next
				If bFlag=True Then
					Fn_Classification_CheckClassificationProperty=True
				Else
					Fn_Classification_CheckClassificationProperty=False
				End If
End Function

''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''''/$$$$
''''/$$$$   FUNCTION NAME   : Fn_Classification_SetValues(sEdit,sSet,sTableRow)
''''/$$$$
''''/$$$$   DESCRIPTION        :  This function will check the classification properties
''''/$$$$
''''/$$$$    PARAMETERS      :   1.) sProperty : Property to be checked
''''/$$$$ 									  2.) sSet : Value to be set
''''/$$$$ 									 3.)sTableRow : Datatable row to be set
''''/$$$$ 									 
''''/$$$$									
''''/$$$$
''''/$$$$	Return Value : 			True / False
''''/$$$$
''''/$$$$    Function Calls       :   Fn_WriteLogFile()
''''/$$$$										
''''/$$$$
''''/$$$$	HISTORY           :   AUTHOR                 DATE        VERSION
''''/$$$$
''''/$$$$    CREATED BY     :   SHREYAS           09/02/2011         1.0
''''/$$$$
''''/$$$$    REVIWED BY     :  Prasanna			    09/02/2011         1.0
''''/$$$$
''''/$$$$    EXAMPLE          :  
''''/$$$$									
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Public Function Fn_Classification_SetValues(sEdit,sSet,sTableRow)
	GBL_FAILED_FUNCTION_NAME="Fn_Classification_SetValues"
   Dim objSelectType,objClassAdmin,sValue,sStaticCount,bFlag
	bFlag=false
	Set WshShell = CreateObject("WScript.Shell")
				Set objSelectType = description.Create()
				objSelectType("Class Name").value = "JavaEdit"
				
				Set objClassAdmin =JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").ChildObjects(objSelectType)
				sStaticCount=JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").ChildObjects(objSelectType).count
				For iCount=0 to sStaticCount-1
						   sValue=objClassAdmin(iCount).GetRoProperty("attached text")
						   If sValue=sEdit Then
							   For sData=1 to sTableRow
								 Datatable.SetCurrentRow(sData)
								 objClassAdmin(iCount).SetFocus
								objClassAdmin(iCount).set DataTable("AttValue", dtGlobalSheet)
								WshShell.SendKeys "({TAB})"
								iCount=iCount+1
							bFlag=True
							Next
						   End If
				Next
				If bFlag=True Then
					 Fn_Classification_SetValues=True
				Else
					 Fn_Classification_SetValues=False
				End If
End Function

'''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'''''/$$$$
'''''/$$$$   FUNCTION NAME   : Fn_Classification_ImportObjects(sAction, sImportType, sImportObject, sTransferMode, bViewLog, aGlobalDictionary, sButtons)
'''''/$$$$
'''''/$$$$   DESCRIPTION        :  This function will create mapping and verify if the mapping has happened
'''''/$$$$
'''''/$$$$    PARAMETERS      :   1.) sAction : Action to be performed
'''''/$$$$                                      2.) sImportType : Import type to be specified
'''''/$$$$ 									   3.) sImportObject : Import file
'''''/$$$$ 									   4.) sTransferMode  :  Transfer mode to be selected
'''''/$$$$									   5.) bViewLog  :  To view or not to view the log
'''''/$$$$									  6.) sButtons    : buttons to be clicked
'''''/$$$$	
'''''/$$$$	Return Value : 			True or False
'''''/$$$$
'''''/$$$$    Function Calls       :   Fn_WriteLogFile()
'''''/$$$$										
'''''/$$$$
'''''/$$$$	HISTORY           :   AUTHOR                 DATE        VERSION
'''''/$$$$
'''''/$$$$    CREATED BY     :   SHREYAS           16/02/2011         1.0
'''''/$$$$
'''''/$$$$    REVIWED BY     :  Prasanna			 16/02/2011         1.0
'''''/$$$$
'''''/$$$$	 Note (09-11-2011) :  Restructured the function due to major UI changes	{By Shreyas}
'''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Public Function Fn_Classification_ImportObjects(sAction, sImportType, sImportObject, sTransferMode, bViewLog, aGlobalDictionary, sButtons)
	GBL_FAILED_FUNCTION_NAME="Fn_Classification_ImportObjects"
   'Declaring Variables
   Dim ObjImport, aButtons, iCount, bReturn,aValues
   'Setting bReturn to False
   bReturn=False
   'Initially Function returns False
   Fn_Classification_ImportObjects=False
   'Checking Existance of ImportToWord Window
   If Fn_UI_ObjectExist("Fn_Classification_ImportObjects",JavaWindow("PLMXML-TeamCenter").JavaWindow("PLMXMLImport"))=False Then
	   'Opening ImportObjects Window
		Call Fn_MenuOperation("Select","Tools:Import:From PLMXML")
   End If
   'Creating object of ImportToWord Window	
   Set ObjImport=JavaWindow("PLMXML-TeamCenter").JavaWindow("PLMXMLImport")
   Select Case sAction
	 	Case "SetPLMXMLImport"
					'Select the Import Type
'					ObjImport.JavaCheckBox("ImportType").SetTOProperty "attached text",sImportType
'					If sImportType<>"" Then
'						Call Fn_CheckBox_Select("Fn_Classification_ImportObjects", ObjImport, "ImportType")
'					End If
'					'SET property for Import dialog
'					ObjImport.SetTOProperty "title","PLM XML Import ..."
'					'Click on Browse button.
'					Call Fn_Button_Click("Fn_Classification_ImportObjects", ObjImport, "Browse")
					'Set Value for Import Directory.
					If sImportObject<>"" Then
'						ObjImport.JavaDialog("SelectObject").JavaEdit("FileName").SetTOProperty "attached text","File name:"
'						If ObjImport.JavaDialog("SelectObject").JavaEdit("FileName").Exist Then
'							Call Fn_Edit_Box("Fn_Classification_ImportObjects",ObjImport.JavaDialog("SelectObject"),"FileName",sImportObject)
'						Else
'							ObjImport.JavaDialog("SelectObject").JavaEdit("FileName").SetTOProperty "attached text","Directory name:"
'							If ObjImport.JavaDialog("SelectObject").JavaEdit("FileName").Exist Then
'								Call Fn_Edit_Box("Fn_Classification_ImportObjects",ObjImport.JavaDialog("SelectObject"),"FileName",sImportObject)
'							End If
'						End If						
			
						ObjImport.JavaEdit("ImportingXMLFile").Set sImportObject
						If err.number<0 Then
										 Fn_Classification_ImportObjects = False			 				
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail: Failed to set the Value "+sImportObject+" in the Import XML FIle Edit Box")
										Exit Function
						Else
										   Fn_Classification_ImportObjects = True			 				
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Pass: Successfully set the Value "+sImportObject+" in the Import XML FIle Edit Box")			
						End If
					End If
'
'					'Click on Select button.
'					Call Fn_Button_Click("Fn_Classification_ImportObjects", ObjImport.JavaDialog("SelectObject"), "Select")

					'check the existence of error Dialog

					If ObjImport.JavaWindow("Error").Exist(5) Then
					
								'JavaWindow("PLMXML-TeamCenter").JavaWindow("TcDefaultApplet").JavaDialog("Export").highlight
								Call Fn_Button_Click("Fn_Classification_ImportObjects", ObjImport.JavaWindow("Error"), "OK")
				Fn_Classification_ImportObjects = False			 				
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail: Error Encountered While Importing Objects")
				Exit Function

					else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "No Error Encountered")
					End If

					'Set the Transfer Mode Name
					If sTransferMode<>"" Then
						Call Fn_List_Select("Fn_Classification_ImportObjects", ObjImport, "TransferMode",sTransferMode)
					End If
					Fn_Classification_ImportObjects = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Pass: Successfully exceuted "& sAction &" case of Fn_Classification_ImportObjects")
		Case Else 
                Fn_Classification_ImportObjects = False			 				
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail: Fn_Classification_ImportObjects failed due to Invalid arguments")
				Exit Function
		End Select
		'Click on Buttons
		If sButtons<>"" Then
				aButtons = split(sButtons, ":",-1,1)				
				For iCount=0 to Ubound(aButtons)
					Call Fn_Button_Click("Fn_Classification_ImportObjects", ObjImport, aButtons(iCount))
				Next
		End If
		'View Lof for details
		If bViewLog<>"" Then
			JavaWindow("PLMXML-TeamCenter").JavaWindow("TcDefaultApplet").JavaDialog("Export Completed").SetTOProperty "title","Import Completed"
			'Do
			'Loop Until JavaWindow("PLMXML-TeamCenter").JavaWindow("TcDefaultApplet").JavaDialog("Export Completed").Exist = True
			Call Fn_ReadyStatusSync(10)
			if JavaWindow("PLMXML-TeamCenter").JavaWindow("TcDefaultApplet").JavaDialog("Export Completed").Exist = True then
				Call Fn_Button_Click("Fn_Classification_ImportObjects", JavaWindow("PLMXML-TeamCenter").JavaWindow("TcDefaultApplet").JavaDialog("Export Completed"), bViewLog)			
			End if 
		End If
		'Setting Object to Nothing
		Set ObjImport=Nothing
End Function



'''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'''/$$$$
'''/$$$$   FUNCTION NAME   : Fn_Classification_SetRevisionRule(sAction,sRule,sButtons)
'''/$$$$
'''/$$$$   DESCRIPTION        :  This function will set the revision rules
'''/$$$$
'''/$$$$    PARAMETERS      :   1.) sAction : Action to be performed
'''/$$$$ 									 2.) sRule : Rule to set
'''/$$$$									3.) sButtons : Buttons to be clicked
'''/$$$$								    4.) sInfo : For future use
'''/$$$$
'''/$$$$	Return Value : 			True or False
'''/$$$$
'''/$$$$    Function Calls       :   Fn_WriteLogFile()
'''/$$$$										
'''/$$$$
'''/$$$$	HISTORY           :   AUTHOR                 DATE        VERSION
'''/$$$$
'''/$$$$    CREATED BY     :   SHREYAS           12/04/2011         1.0
'''/$$$$
'''/$$$$    REVIWED BY     :  Prasanna			  12/04/2011         1.0
'''/$$$$
'''/$$$$    EXAMPLE          :  bReturn=Fn_Classification_SetRevisionRule("Set","Latest Working","Apply:OK","")
'''/$$$$									
'''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$


Public Function Fn_Classification_SetRevisionRule(sAction,sRule,sButtons,sInfo)

	GBL_FAILED_FUNCTION_NAME="Fn_Classification_SetRevisionRule"
Dim objRevisionRule,iCounter,aButtons,iCount
Fn_Classification_SetRevisionRule=false
	
	'invoke the View / Set Current Revision Rule dialog 
	Set objRevisionRule=JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaDialog("ViewSetCurrentRevision")

	If  Not(objRevisionRule.Exist(5)) Then
				If JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaButton("down_16").GetROProperty ("enabled")="1" then
					JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaButton("down_16").Click micLeftBtn
					If err.number<0 Then
						Fn_Classification_SetRevisionRule=false
						exit function
					End If
				End if
		
				JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaMenu("MenuSelect").SetTOProperty "label", "View / Set Current"
				JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaMenu("MenuSelect").Click 0,0,"LEFT"
			
				If err.number<0 Then
					Fn_Classification_SetRevisionRule=false
					exit function
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully invoked the [View / Set Current Revision Rule] dialog") 
				End If
	End If
			
				Select Case sAction
			
					Case "Set"	
			
							'Set the revision rule and click on buttons
							If objRevisionRule.Exist(3) Then
									objRevisionRule.JavaList("Rules").Select sRule
									If err.number<0 Then
										Fn_Classification_SetRevisionRule=false
										exit function
									Else
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully set the value ["+sRule+"]") 
										Fn_Classification_SetRevisionRule=True
									End If
							Else
									Fn_Classification_SetRevisionRule=false
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The Dialog [View / Set Current Revision Rule] does not exist") 
									exit function
					
							End If
				End Select
	
	 'Click on Buttons
	 If sButtons<>"" Then
			   aButtons = split(sButtons, ":",-1,1)
			   iCounter = Ubound(aButtons)
			   For iCount=0 to iCounter
					objRevisionRule.JavaButton(aButtons(iCount)).Click micLeftBtn
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully clicked on the button ["+aButtons(iCount)+"]") 
			   Next
					Fn_Classification_SetRevisionRule=True
	 End If

  Set objRevisionRule=nothing

End Function



Public function Fn_Classification_FavoritesOperations(sAction, sFolderName, sQuery,bExecute,sDetails)
	GBL_FAILED_FUNCTION_NAME="Fn_Classification_FavoritesOperations"
   Select Case sAction
		 	Case "AddFolder"
				If JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaTree("Favorites").Exist(5) = false Then
							bReturn =  Fn_ClassAdmin_ToolbarOperations("Favorite")
							If bReturn = false Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Click on [Favorite] Button.")
										Fn_Classification_FavoritesOperations = false
										Exit Function 	
							Else
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully Clicked on [Favorite] Button")									
							End If
							Call Fn_ReadyStatusSync(2) 
				End If

						'Click on Add button 
						JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaButton("AddFavorite").Click micLeftBtn
						If Err.Number < 0  Then
								Fn_Classification_FavoritesOperations = false
								Exit function
						End If

					'	If Javawindow("ClassificationMainWin").Javawindow("ClassificationSubWindow").Javadialog("AddFavoriteDialog").exist(5)  Then
						If Javawindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").Javadialog("AddFavoriteDialog").exist(5)  Then
								'Click on New folder button
								'Javawindow("ClassificationMainWin").Javawindow("ClassificationSubWindow").Javadialog("AddFavoriteDialog").JavaButton("New Folder").Click micLeftBtn
								Javawindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").Javadialog("AddFavoriteDialog").JavaButton("New Folder").Click micLeftBtn
								If Err.Number < 0  Then
										Fn_Classification_FavoritesOperations = false
										Exit function
								End If

'								If Javawindow("ClassificationMainWin").Javawindow("ClassificationSubWindow").Javadialog("Create new folder").exist(5)  Then
'										Javawindow("ClassificationMainWin").Javawindow("ClassificationSubWindow").Javadialog("Create new folder").JavaEdit("FolderName").Set sFolderName
								If Javawindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").Javadialog("Create new folder").exist(5)  Then
										Javawindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").Javadialog("Create new folder").JavaEdit("FolderName").Set sFolderName
										If Err.Number < 0  Then
												Fn_Classification_FavoritesOperations = false
												Exit function
										End If	
								End If

								'CLick on OK button
								'Javawindow("ClassificationMainWin").Javawindow("ClassificationSubWindow").Javadialog("Create new folder").JavaButton("OK").Click micLeftBtn
								Javawindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").Javadialog("Create new folder").JavaButton("OK").Click micLeftBtn
								If Err.Number < 0  Then
												Fn_Classification_FavoritesOperations = false
												Exit function
								End If		

								'Javawindow("ClassificationMainWin").Javawindow("ClassificationSubWindow").Javadialog("AddFavoriteDialog").Close
								Javawindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").Javadialog("AddFavoriteDialog").Close
								Fn_Classification_FavoritesOperations = true
						End If
						
			Case "AddClass"
						If JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaTree("Favorites").Exist(5) = false Then
								bReturn =  Fn_ClassAdmin_ToolbarOperations("Favorite")
								If bReturn = false Then
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Click on [Favorite] Button.")
											Fn_Classification_FavoritesOperations = false
											Exit Function 	
								Else
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully Clicked on [Favorite] Button")									
								End If
								Call Fn_ReadyStatusSync(2) 
						End If

						'Click on Add button 
						JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaButton("AddFavorite").Click micLeftBtn
						If Err.Number < 0  Then
								Fn_Classification_FavoritesOperations = false
								Exit function
						End If

					'	If Javawindow("ClassificationMainWin").Javawindow("ClassificationSubWindow").Javadialog("AddFavoriteDialog").exist(5)  Then
						If Javawindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").Javadialog("AddFavoriteDialog").exist(5)  Then
								'Select the Folder 
								'Javawindow("ClassificationMainWin").Javawindow("ClassificationSubWindow").Javadialog("AddFavoriteDialog").JavaTree("FolderList").Select "#0:"+sFolderName
								Javawindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").Javadialog("AddFavoriteDialog").JavaTree("FolderList").Select "#0:"+sFolderName
								If Err.Number < 0  Then
										Fn_Classification_FavoritesOperations = false
										Exit function
								End If	

						
								'Click on CreateIn button
								'Javawindow("ClassificationMainWin").Javawindow("ClassificationSubWindow").Javadialog("AddFavoriteDialog").JavaButton("CreateIn").Click micLeftBtn
								Javawindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").Javadialog("AddFavoriteDialog").JavaButton("CreateIn").Click micLeftBtn
								If Err.Number < 0  Then
										Fn_Classification_FavoritesOperations = false
										Exit function
								End If
						End if		

						wait 2
                        
						'Set the Query
						'If Javawindow("ClassificationMainWin").Javawindow("ClassificationSubWindow").Javadialog("CreateInDialog").exist(5)  Then
						If Javawindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").Javadialog("CreateInDialog").exist(5)  Then
									   If  sQuery <> "" Then
											'Javawindow("ClassificationMainWin").Javawindow("ClassificationSubWindow").Javadialog("CreateInDialog").JavaEdit("QueryName").Set sQuery
											Javawindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").Javadialog("CreateInDialog").JavaEdit("QueryName").Set sQuery
											If Err.Number < 0  Then
													Fn_Classification_FavoritesOperations = false
													Exit function
											End If	
									   End If

									'Javawindow("ClassificationMainWin").Javawindow("ClassificationSubWindow").Javadialog("CreateInDialog").JavaButton("OK").Click micLeftBtn
									Javawindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").Javadialog("CreateInDialog").JavaButton("OK").Click micLeftBtn
									If Err.Number < 0  Then
											Fn_Classification_FavoritesOperations = false
											Exit function
									End If	
									Fn_Classification_FavoritesOperations = true
						End If
								
   End Select
end function




'a = Fn_Classification_FavoritesTree("Select","NF","","")

Public function Fn_Classification_FavoritesTree(sAction, sNodeName, sMenu,sDetails)
	GBL_FAILED_FUNCTION_NAME="Fn_Classification_FavoritesTree"
   Dim objFavTree,iItemCount,iCounter,sTreeItem

   Set objFavTree = JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaTree("Favorites")
   Select Case sAction
		 	   Case "Select"
			             objFavTree.Select "#0:"+sNodeName
				Case "Expnad"
						 objFavTree.Expand "#0:"+sNodeName
				Case "Collapse"
						 objFavTree.Collapse "#0:"+sNodeName  
				Case "DoubleClick"
						 objFavTree.Activate "#0:"+sNodeName 
				 Case "Exist"
								iItemCount = objFavTree.GetROProperty( "items count")
								For iCounter=0 To (iItemCount-1)
										sTreeItem = objFavTree.GetItem(iCounter)
										If Trim (Lcase(sTreeItem)) = Trim(Lcase(sNodeName)) Then
												Fn_Classification_FavoritesTree = True
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully found node [" + sNodeName + "] of Favorites Tree." )	
												Exit For
										End If
								Next 

								If  Cint(iCounter) = Cint (iItemCount) Then
										Fn_Classification_FavoritesTree = False
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to  found node [" + sNodeName + "] of Favorites Tree." )	
										Set objSearchTree = Nothing
										Exit Function 
								End If		 							 
   End Select

   If Err.Number < 0  Then
			Fn_Classification_FavoritesTree = false
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Failed to Perform ["+sAction+"]")
			Exit function
   else
			 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully Performed ["+sAction+"]")
			 Fn_Classification_FavoritesTree = true
   End If
End function 


'''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'''/$$$$
'''/$$$$   FUNCTION NAME   :  Fn_Classification_UnitSystemSearch(sAction,sUnit,sInfo1,sInfo1)
'''/$$$$
'''/$$$$   DESCRIPTION        :  This function will select the appropriate menu after the UnitSystemSearch button click event
'''/$$$$
'''/$$$$    PARAMETERS      :   1.) sAction : Action to be performed
'''/$$$$ 									    2.) sUnit : Valid Menu name to be selected
'''/$$$$									   3.) sInfo1 : For Future Use
'''/$$$$									  4.) sInfo2 : For Future Use
'''/$$$$
'''/$$$$	Return Value : 			True / False
'''/$$$$
'''/$$$$    Function Calls       :   Fn_WriteLogFile()
'''/$$$$										
'''/$$$$
'''/$$$$	HISTORY           :   AUTHOR                 DATE        VERSION
'''/$$$$
'''/$$$$    CREATED BY     :   SHREYAS           23/05/2011         1.0
'''/$$$$
'''/$$$$    REVIWED BY     :  SHREYAS			    23/05/2011          1.0
'''/$$$$
'''/$$$$    EXAMPLE          :  bReturn= Fn_Classification_UnitSystemSearch("Set","metric","","")
'''/$$$$									
'''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

Public function Fn_Classification_UnitSystemSearch(sAction,sUnit,sInfo1,sInfo2)
		GBL_FAILED_FUNCTION_NAME="Fn_Classification_UnitSystemSearch"
		Fn_Classification_UnitSystemSearch=false

		Dim objUnit, xCord,yCord,DeviceReplay

		Set objUnit=JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow")

		Select Case sAction

				Case "Set"
							
						   xCord= objUnit.JavaButton("UnitSystemSearch").GetROProperty ("abs_x")
							yCord=objUnit.JavaButton("UnitSystemSearch").GetROProperty ("abs_y")
					
							Set DeviceReplay = CreateObject("Mercury.DeviceReplay")
							DeviceReplay.MouseClick xCord,yCord,LEFT_MOUSE_BUTTON
							wait 3

								'now click on the menu of choice
								objUnit.JavaMenu("MenuSelect").SetTOProperty "label",sUnit

								objUnit.JavaMenu("MenuSelect").Click 0,0,"RIGHT"
								If Err.Number < 0 then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Click on the menu ["+sUnit+"]")
										Fn_Classification_UnitSystemSearch = false
										Set objUnit=nothing
										Exit Function
								Else
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully  Clicked on the menu ["+sUnit+"]")
										Fn_Classification_UnitSystemSearch = True
								End if

		End Select

		Set objUnit=nothing

End Function

'********************************************************* Function to handle error dialogs ***********************************************************************

'Function Name		          :      Fn_Classification_ErrorHandler

'Description			    :	 Function to handle error dialogs

'Parameters			   :	1. sAction : Action to perform
'						2. sTitle : Title to set
'						3. sMsg : Message to verify
										
'Return Value		         : 	True / False

'Examples				:        call  Fn_Classification_ErrorHandler("ErrorDialog", "Constraints Errors", "The default value should be within the range specified by the minimum and maximum values.")
							        
'History:
'			Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'			Vidya 				  18-Aug-2011		   1.0
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
public function Fn_Classification_ErrorHandler(sAction, sTitle, sMsg)
	GBL_FAILED_FUNCTION_NAME="Fn_Classification_ErrorHandler"
	Dim objErrorWindow
	Fn_Classification_ErrorHandler = False
	Select Case sAction
			Case "FormatError"
				Set objErrorWindow = JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaDialog("Format error")
				If sTitle <> ""  Then
					objErrorWindow.SetTOProperty "title", trim(sTitle)
				End If
				' different error dialog
				If objErrorWindow.Exist(5) Then
					Fn_Classification_ErrorHandler = True
					If sMsg <> "" Then
						If instr(objErrorWindow.JavaStaticText("Msg").GetROProperty ("label"), sMsg) < 0 then
							Fn_Classification_ErrorHandler = false
						End if
					End If
				objErrorWindow.JavaButton("OK").SetTOProperty "Index",0
				objErrorWindow.JavaButton("OK").SetTOProperty "displayed", "1"	
				objErrorWindow.JavaButton("OK").Click micLeftBtn
				Fn_Classification_ErrorHandler = True
				Set objErrorWindow = nothing
				End If
			Case "DeleteConfirmation" '[TC1015-2015081100-27_08_2015-VivekA-NewDevelopment] - Added to handle Delete dialog
				Set objErrorWindow = JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaDialog("DeleteConfirmation")
				If sTitle <> ""  Then
					objErrorWindow.SetTOProperty "title", trim(sTitle)
				End If
				' different Delete dialog
				If objErrorWindow.Exist(5) Then
					'Click on Yes button 
					If Fn_Button_Click("Fn_Classification_ErrorHandler", objErrorWindow, "Yes") = true  Then
						Fn_Classification_ErrorHandler = True
						Set objErrorWindow = nothing
					Else
						Fn_Classification_ErrorHandler = False
						Set objErrorWindow = nothing
					End If
				End If
	End Select
End function





Function Fn_ClassAdmin_SearchExistingUnitName(sNametosearch)
Fn_ClassAdmin_SearchExistingUnitName = False
JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaEdit("UnitClassText").SetTOProperty "Index","1"
wait 2
JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaEdit("UnitClassText").Set sNametosearch
wait 2
JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaButton("Search").Click
wait 2
JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaTab("TabName").Select "Table"
wait 2
JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaTable("SearchResultTable").SelectRow("0")
wait 2
JavaWindow("ClassificationMainWin").JavaToolbar("AddorEditInstance").Press("Edit current Instance")

'JavaWindow("Classification - Teamcenter").JavaToolbar("EditCurrentInstance").Press("Edit current Instance")
Fn_ClassAdmin_SearchExistingUnitName = True
End Function
