Option Explicit

'1. Fn_SISW_BB_UI_TreeGetItemPath(sFunctionName,sAction,ObjTree,StrNode,sType)
'2. Fn_SISW_BB_UI_JavaTab_Operations(sAction, objJavaDialog, sTabObjectName, sTabName,sPopupMenu)
'3. Fn_SISW_BB_UI_JavaTree_Operations(sFunctionName,sAction,objJavaDialog,sJavaTree,sNode,sColName,sColValue,sColType,sRMB)
'4. Fn_SISW_BB_UI_XML_Operations(sFunctionName,sAction,discXmlPara)
'*********************************************************		Function to Synchronize on Application Response	***********************************************************************
'Function Name		:				Fn_SISW_BB_UI_TreeGetItemPath(sFunctionName,sAction,ObjTree,StrNode,sType)

'Description			 :		 		 to get path of javatree node

'Parameters			   :	 			1. sFunctionName : name of function from where it is calling
'										2. sAction: this is will be set to "Datasets" while fetching path of dataset tree otherwise keep it bank
'										3. ObjTree: tree object on which retriving the path of node
'										4. StrNode : full path of node
'										5. sType : this will be set in case of dataset tree, to define "type" of node we are retriving - where TYPE is Column name under tree
'
'Return Value		   : 				Path of tree \ FALSE

'Pre-requisite			:		 		Node should be selected on which dataset is created

'Examples				:				 Fn_SISW_BB_UI_TreeGetItemPath("Fn_SISW_BB_DataSetTab_Opeartions","Datasets",ObjTree,"DS1Word/A:Child1","MSWord")

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Shweta Rathod			25-Jul-2016			1.0										Koustubh Watwe
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_BB_UI_TreeGetItemPath(sFunctionName,sAction,ObjTree,StrNode,sType)
	Dim sItemPath,aStrNode,bFlag,i,iNodeItemsCount
	Dim oCurrentNode,eStrNode, iNodecnt
	
	bFlag = false
	aStrNode = split(StrNode,":")
	iRow = objTree.Object.getItemCount()
	If sAction = "Datasets" Then
		For iCnt = 0 to iRow - 1
			eStrNode = aStrNode(iNodecnt)
			sAppTypeCol = objTree.GetColumnValue(iCnt,"Type")
			sAppNameCol = objTree.GetColumnValue(iCnt,"Name")
			If sAppTypeCol = sType and sAppNameCol = Trim(eStrNode) then
				iRootIndex = iCnt
				bFlag=True
				Exit for
			End if
		next
	else
		bFlag=True
		iRootIndex = 0
	End If
	If bFlag = true then 
		Set oCurrentNode = ObjTree.Object.getItem(iRootIndex)
		sItemPath = "#" & iRootIndex
		
		If UBound(aStrNode) = 0 Then
			Fn_SISW_BB_UI_TreeGetItemPath = sItemPath
			If sItemPath=False Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Function " & sFunctionName & " Failed to find item [ " & StrNode & " ]"  )
			Else
				Set objNodeBounds = oCurrentNode.getBounds()
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Function " & sFunctionName & " executed successfully for item [ " & StrNode & " ]"  )
			End If
			Exit Function
		End If
		If bFlag Then
			bFlag = False
		Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Function " & sFunctionName & " Failed to find item [ " & StrNode & " ]"  )
			Exit function
		End If
		
		For iNodecnt = 1 to UBound(aStrNode)
			eStrNode = aStrNode(iNodecnt)
			iNodeItemsCount = oCurrentNode.getItemCount()
			For i = 0 to iNodeItemsCount - 1
				If Trim(oCurrentNode.getItem(i).getText()) = Trim(eStrNode) Then
					Set oCurrentNode = oCurrentNode.getItem(i)
					sItemPath = sItemPath & ":#" & i
					bFlag=True
					Exit For
				End if
			next
		next
	End If
	
	If bFlag=True Then
		'Function Returns Item Path
		Fn_SISW_BB_UI_TreeGetItemPath = sItemPath
		Set objNodeBounds = oCurrentNode.getBounds()
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Function " & sFunctionName & " executed successfully for item [ " & StrNode & " ]"  )
	Else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Function " & sFunctionName & " Failed to find item [ " & StrNode & " ]"  )
		Fn_SISW_BB_UI_TreeGetItemPath = False
	End If
	Set oCurrentNode =Nothing
End Function

'*************************************************  Function to perform various opeartions on the prefrence operation window ***********************************************************************
'Function Name		:				Fn_SISW_BB_UI_JavaTab_Operations(sAction, objJavaDialog, sTabObjectName, sTabName,sPopupMenu)

'Description			 :		 		to perform various opeartions on the java tab which is displaying in BB application

'Parameters			   :	 			1. sAction : select/exist/close
'										2. objJavaDialog : dilaog on which operation needs to perform / if it is balnk it will set to default window of BB
'										3. sTabObjectName : object of tab / if balnk default will be "CTabFolder"
'										4. sTabName : "Welcome"
'										4. sPopupMenu : popup menu to be performed
'
'Return Value		   : 				true \ false

'Pre-requisite			:		 		Briefcase browser or dilaog on which tab is displaying should be open 

'Examples				:				bReturn	= Fn_SISW_BB_UI_JavaTab_Operations("close","","","Welcome","","")

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Shweta Rathod			01-Aug-2016			1.0											Koustubh Watwe
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_BB_UI_JavaTab_Operations(sAction, objJavaDialog, sTabObjectName, sTabName,sPopupMenu) 'Fn_SISW_UI_RACTabFolderWidget_Operation
	Dim objTab,bFlag
	Fn_SISW_BB_UI_JavaTab_Operations = false
	if objJavaDialog = "" then set objJavaDialog = JavaWindow("BriefcaseBrowser")
	if sTabObjectName = "" and sTabName <> "Ownership Transfer" then 
		sTabObjectName = "CTabFolder" 
		Set objTab = objJavaDialog.JavaTab(sTabObjectName)
		objTab.SetTOProperty "value",sTabName
	else
		Set objTab = objJavaDialog.JavaTab(sTabObjectName)
		objTab.select sTabName
		wait 1
	End if
	if objTab.exist(1) = false then
		bFlag = false	
	else
		bFlag = true
	End if	
	Select Case lcase(sAction)
		Case "select"
				If bFlag = true Then
					objTab.Select sTabName
					Exit Function
				Else 
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "tab does not exist"+objTab.tostring())
				End If
		Case "exist"
				If bFlag = false Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "tab does not exist"+objTab.tostring())
					Exit Function
				End If			
		Case "close"
			if bFlag = true then 
				err.clear
				objTab.CloseTab sTabName
				If Err.Number < 0 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to close the tab " + objTab.tostring())
					Exit Function
				End if
			End if
		Case "isminimize"
			err.clear
			objTab.Minimize 
			If Err.Number < 0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to minimize the tab " + objTab.tostring())
				Exit Function
			End if
			bFlag = objTab.object.getMinimized()
			If bFlag = "false" Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to verify tab is in minimize state " + objTab.tostring())
				Exit Function
			End If
			wait 2
			objTab.Restore 
		Case "ismaximize"
			err.clear
			objTab.Maximize 
			If Err.Number < 0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Maximize the tab " + objTab.tostring())
				Exit Function
			End if
			wait 1
			bFlag = objTab.object.getMaximized()
			If bFlag = "false" Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to verify tab is in Maximize state " + objTab.tostring())
				Exit Function
			End If
			wait 2
			objTab.Restore
	End Select 
	Set objTab = Nothing	
	Fn_SISW_BB_UI_JavaTab_Operations = true
End Function

'*********************************************************		Function to Get the column index form the javatree 		***********************************************************************

'Function Name		:					Fn_SISW_BB_UI_JavaTree_Operations

'Description			 :		 		  This function is used to perfrom different opearations on javatree 

'Parameters			   :	 			1.  sFunctionName : name of function to called from 
'										2.  sAction:Name of the action to be perdformed
'										3. objJavaDialog:parent object of tree object or tree object
'										4. sJavaTree: name of tree object if any or keep it blank of provided in "objJavaDialog"
'										5. sNode: path of node to be select
'										6. sColName: column name to retrive the value of corresponding node 
'										7. sColValue: column value to be verfied corresponding to the node 
'										8. sColType	: used in dataset tree opeartions only to differentiate node by its "Type" coulmn name
'										9. sRMB: RMB menu to be performed
											
'Return Value		   : 				iPath/iCol_index/True/False

'Pre-requisite			:		 		Tree should be displayed in application

'Examples				:				  For Any other tree - 
'										 1. Fn_SISW_BB_UI_JavaTree_Operations("Fn_SISW_BB_DataSetTree_Opeartions","select",objTree,"","Top:Child1","","","","")
'										 For dataset tree same is applicable to all other following examples - 
'										 1. Fn_SISW_BB_UI_JavaTree_Operations("Fn_SISW_BB_DataSetTree_Opeartions","select",objTree,"","Top:Child1","","","MSWord","")
'										 2. Fn_SISW_BB_UI_JavaTree_Operations("Fn_SISW_BB_DataSetTree_Opeartions","exist",objTree,"","Top:Child1","","","MSWord","")
'										 3. Fn_SISW_BB_UI_JavaTree_Operations("Fn_SISW_BB_DataSetTree_Opeartions","expand",objTree,"","Top:Child1","","","MSWord","")
'										 4. Fn_SISW_BB_UI_JavaTree_Operations("Fn_SISW_BB_DataSetTree_Opeartions","collapse",objTree,"","Top:Child1","","","MSWord","")
'										 5. Fn_SISW_BB_UI_JavaTree_Operations("Fn_SISW_BB_DataSetTree_Opeartions","getcolumnval",objTree,"","Top:Child1","Part Name","","","")
'										 6. Fn_SISW_BB_UI_JavaTree_Operations("Fn_SISW_BB_DataSetTree_Opeartions","doubleclick",objTree,"","Top:Child1","","","MSWord","")
'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Shweta Rathod		 12-sep-2016   			1.0					created					Koustubh watwe
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_BB_UI_JavaTree_Operations(sFunctionName,sAction,objJavaDialog,sJavaTree,sNode,sColName,sColValue,sColType,sRMB)
Dim iPath,iCols,sAppColName,iCounter,objTree,objTemp,arrPath,iArrCnt,iRow,sAppColVal
Fn_SISW_BB_UI_JavaTree_Operations = false
If sJavaTree <> "" Then
	Set objTree = objJavaDialog.JavaTree(sJavaTree)
	sFuncLog = sFunctionName + " > Fn_SISW_BB_UI_JavaTree_Operations  : [ " &  objJavaDialog.toString & " ] : [ " +  sJavaTree + " ] : Action = " & sAction & " : "
Else
	Set objTree = objJavaDialog
	sFuncLog = sFunctionName + " > Fn_SISW_BB_UI_JavaTree_Operations  : [ " +  objTree.toString + " ] : Action = " & sAction & " : "
End IF

If lcase(sAction) <> "isempty" then
	If sFunctionName = "Fn_SISW_BB_DataSetTree_Opeartions" then
		If instr(sNode,"@")>0 then
			iPath = Fn_UI_JavaTreeGetItemPathExt("Fn_SISW_BB_UI_JavaTree_Operations",ObjTree,sNode,":","@")
		else
			iPath = Fn_SISW_BB_UI_TreeGetItemPath("Fn_SISW_BB_UI_JavaTree_Operations","Datasets",ObjTree,sNode,sColType)
		End if
	else
		iPath = Fn_UI_JavaTreeGetItemPathExt("Fn_SISW_BB_UI_JavaTree_Operations", objTree, sNode ,"","")
	End if

	If iPath = False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog&"Node does not exist " + sNode+" in javatree "+objTree.ToString())
		exit function
	End If
End if

Select Case lcase(sAction)
	Case "select"
			Err.clear
			objTree.Select iPath
			wait 1
			If Err.Number < 0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog&"Failed to select node  " + sNode +" in javatree " + objTree.ToString())
				Exit Function
			End if
	Case "expand"
			err.clear
			objTree.Expand iPath
			If Err.Number < 0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog&"Failed to expand node  " + dicOpenAssm("Node") +" in javatree " + objTree.ToString())
				Exit Function
			End if
			wait 1
	Case "collapse"
			err.clear
			objTree.Collapse iPath
			If Err.Number < 0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog&"Failed to expand node  " + dicOpenAssm("Node") +" in javatree " + objTree.ToString())
				Exit Function
			End if
			wait 1
	Case "exist"
			If iPath = false Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog&"Failed to exist node  " + sNode +" in javatree " + objTree.ToString())
				Exit Function
			End if
	Case "getcolumnindex"
		iCols = objTree.GetROProperty("columns_count")
		'Get the Col No. of required Column
		For iCounter = 0 to iCols -1
			sAppColName =objTree.GetColumnHeader("#"&iCounter)		  
			If Trim(sAppColName) = Trim(sColName) Then
				Fn_SISW_BB_UI_JavaTree_Operations = iCounter
				Exit Function
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog&"Fn_SISW_BB_UI_JavaTree_Operations:The Column Index for Column [" + sColName +"] is [" + iCounter + "]")	
				Exit For
			End If
		Next
		If Cint(iCounter) = Cint(iCols) Then
			Exit Function
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),sFuncLog& "Fn_SISW_BB_UI_JavaTree_Operations:The Column [" + sColName + "] dose not exist in BB  tree")	
		End If
	Case "verifytickmark"
		iColPos = Fn_SISW_BB_UI_JavaTree_Operations("Fn_SISW_BB_BriefCaseTree_Opearation","getcolumnindex",objTree,"",sNode,sColName,"","","")
		iColPos = Fn_SISW_BB_UI_JavaTree_Operations("Fn_SISW_BB_BriefCaseTree_Opearation","getcolumnindex",objTree,"",sNode,sColName,"","","")
        If cint(iColPos) <> 0 then            
            If iColPos = false Then
                Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to find column  " + sColName +" in javatree " + objTree.ToString())
                Exit Function
            End If
        End if 			
		iPath=Replace(iPath,"#","")
		arrPath=Split(iPath,":")
		If UBound(arrPath) <> 0 Then
			For iArrCnt = 0 to UBound(arrPath)
				arrPath(iArrCnt) = cInt(arrPath(iArrCnt))
				If iArrCnt = 0 Then
					Set objTemp = objTree.Object.getItem(arrPath(iArrCnt))
				Else
					Set objTemp = objTemp.getItem(arrPath(iArrCnt)) 
				End If			
			Next
		else
			Set objTemp = objTree.Object.getItem(arrPath(0))
		End If
		If isobject(objTemp.getImage(iColPos)) = False Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog&"Failed to Verify column [" + sColName +"] value in javatree " + objTree.ToString())
				Exit Function
		End If
	Case "getcolumnval"
		sAppColVal = objTree.GetColumnValue(iPath,sColName)
		Fn_SISW_BB_UI_JavaTree_Operations = sAppColVal
		Exit function
	Case "doubleclick"
		err.clear
		objTree.Select iPath
		wait 1
		Call Fn_KeyBoardOperation("SendKeys", "{ENTER}")
		wait 1
		If Err.Number < 0 Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog&"Failed to Double click node  " + sNode +" in javatree " + objTree.ToString())
			Exit Function
		End if
	Case "isempty"
		iRow = objTree.Object.getItemCount()
		If iRow <> 0 then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog&" Failed to verify that tree is empty "+objTree.tostring())
			Exit Function
		End if
	Case "popupmenu"
		bRet = Fn_UI_ClickJavaTreeCell("Fn_SISW_BB_UI_JavaTree_Operations",objJavaDialog,sJavaTree,sNode,sColName,"right")
		If bRet = false Then Exit Function
		wait 1
		bRet = Fn_UI_JavaMenu_Select("Fn_SISW_BB_UI_JavaTree_Operations",objJavaDialog,sRMB)
		If bRet = false Then Exit Function
	End select	
Fn_SISW_BB_UI_JavaTree_Operations = true
End Function

'****************************************    Function to perform opeartions on XML file ***************************************
'Function Name		 	:	Fn_SISW_BB_UI_XML_Operations
'
'Description		    :  	Function to perform opeartions on XML file
'
'Parameters		    	:	1. sFunctionName : function from this function is being called
'							2. sTagName : Tag name of the xml file
'							3. sAttributeName : attribute name to select data 
'							4. sAttributeVal : search Value to be modfied
'							5. sSubAttributeName : Sub attribute name 
'							6. sNewAttributeVal : Value to be modified 
'
'Return Value		    :  	TRUE / Attribute value
'
'Examples		     	:	Set discXmlPara = CreateObject("Scripting.Dictionary")
'							discXmlPara("XMLFilePath") = "C:\Tc11.2.3_2016060800_BB_Configured\bbworkspace\configurations\Unman\CustomMappings.xml"
'							discXmlPara("TagName") = "oem"
'							discXmlPara("AttributeName") = "name"
'							discXmlPara("AttributeValue") = "Teamcenter"
'							discXmlPara("SubAttributeName") = "site_id1"
'							bret = Fn_SISW_BB_UI_XML_Operations("","getxmlattributevalue",discXmlPara)
'
'							Set discXmlPara = CreateObject("Scripting.Dictionary")
'							discXmlPara("XMLFilePath") = "C:\Tc11.2.3_2016060800_BB_Configured\bbworkspace\configurations\Unman\visible_attributes.xml"
'							discXmlPara("TagName") = "group"
'							discXmlPara("AttributeName") = "name"
'							discXmlPara("AttributeValue") = "Item"
'							discXmlPara("SubAttributeName") = "name~name~name~name"
'							discXmlPara("SubAttributeValue")="item_id~object_type~object_name~object_desc"
'							bret = Fn_SISW_BB_UI_XML_Operations("","verifyattribute",discXmlPara)
'
'History:
'	Developer Name			Date			Rev. No.		Reviewer			Changes Done	
'----------------------------------------------------------------------------------------------------------------------------------
'	Shweta Rathod		 23-Sep-2016		1.0				
'----------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_BB_UI_XML_Operations(sFunctionName,sAction,discXmlPara)
	Dim objXMLDoc,objXMLNodeList,inumObjXMLNodeList,i,sXMLAttributeVal,bFlag
	Dim jCnt,aSubAttrName,aSubAttrVal,bRet,iFound
	Fn_SISW_BB_UI_XML_Operations = false
	
	set objXMLDoc = CreateObject("Microsoft.XMLDOM")
	bRet = fn_splm_util_file_operation("exist",discXmlPara("XMLFilePath"))
	If bRet = false then 
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFunctionName&" Failed : Failed dose not exist "+discXmlPara("XMLFilePath"))
		Exit Function
	else
		objXMLDoc.load(discXmlPara("XMLFilePath"))
	End if
	Select Case lcase(sAction)
		Case "getxmlattributevalue"	,"verifyattribute"
			Set objXMLNodeList = objXMLDoc.getElementsByTagName(discXmlPara("TagName"))
			inumObjXMLNodeList = objXMLNodeList.length
			If inumObjXMLNodeList = 0 then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFunctionName&" Failed : Failed find Tag name "+discXmlPara("TagName"))
				Exit function
			End if
			For i = 0 to inumObjXMLNodeList - 1
				sXMLAttributeVal = objXMLNodeList.item(i).getAttribute(discXmlPara("AttributeName"))
				If sXMLAttributeVal = discXmlPara("AttributeValue") then 
					bFlag = true
					exit for	
				End if
			next
			
			If bFlag = false then 
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFunctionName&" Failed : Failed to find TagName "+discXmlPara("TagName")+" or attribute name = "+discXmlPara("AttributeName")+" and attribute value "&discXmlPara("AttributeValue"))
				Exit Function
			End if
		Case else
			'do nothing
	End select
	Select Case lcase(sAction)
		Case "getxmlattributevalue"				
			If discXmlPara("SubAttributeName") <> "" then
				sXMLAttributeVal = objXMLNodeList.item(i).getAttribute(discXmlPara("SubAttributeName"))
				If isNull(sXMLAttributeVal) then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFunctionName&" Failed : Failed find sub attribute name "+discXmlPara("SubAttributeName")+" under attribute name "+discXmlPara("AttributeName"))
					Exit function
				else
					Fn_SISW_BB_UI_XML_Operations = sXMLAttributeVal
					Exit function
				End if
			End if
			
		Case "verifyattribute"
			If discXmlPara("SubAttributeName") <> "" then
				Set objChildNodes = objXMLNodeList.item(i).childNodes
				numObjXMLNodeList = objChildNodes.length
				If numObjXMLNodeList = 0 then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFunctionName&" Failed : Failed to find attribute "+discXmlPara("AttributeName") + " = " + discXmlPara("AttributeValue"))
					Exit function
				End if
				aSubAttrName = split(discXmlPara("SubAttributeName"),"~")
				aSubAttrVal = split(discXmlPara("SubAttributeValue"),"~")
				iFound = 0
				For jCnt = 0 to ubound(aSubAttrName)					
					For i = 0 to numObjXMLNodeList - 1
						If objChildNodes.item(i).getAttribute(aSubAttrName(jCnt)) = aSubAttrVal(jCnt) Then
							iFound = iFound + 1
							Exit for
						End If		
					next
				next
				If iFound <> ubound(aSubAttrName)+1 then 
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFunctionName&" Failed : failed to find sub attribute "+aSubAttrName(jCnt-1)+" = " + aSubAttrVal(jCnt-1) + " under "+discXmlPara("AttributeName") + " = " + discXmlPara("AttributeValue"))
					Exit function
				End if
			End if
	End Select
	Set objXMLDoc = nothing 
	Set objXMLNodeList = nothing
	Fn_SISW_BB_UI_XML_Operations = true
End Function
