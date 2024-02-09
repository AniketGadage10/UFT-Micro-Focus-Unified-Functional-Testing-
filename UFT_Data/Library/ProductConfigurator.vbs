Option Explicit

' Function List
'= = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'000.  Fn_SISW_PC_GetObject()
'001.  Fn_SISW_PC_NavTreeTableOperations()
'002.  Fn_SISW_CP_VariantNatTableOperations()
'003.  Fn_SISW_PC_VariantConstraintsViewTableOperations()
'004.  Fn_SISW_PC_VariantDefaultsViewTableOperations()
'005.  Fn_SISW_PC_DeleteVariants()
'006.  Fn_SISW_PC_VariantSaveRulesViewTableOperations()
'007.  Fn_SISW_PC_InclusionRulesViewTableOperations()
'= = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		 	:	Fn_SISW_PC_GetObject
'
'Description		    :  	Function to get specified Object hierarchy.
'
'Parameters		    :	1. sObjectName : Object Handle name
'								
'Return Value		    :  	Object \ Nothing
'
'Examples		     	:	Fn_SISW_PC_GetObject("ProductConfigurator")
'
'History:
'				Developer Name						Date					Rev. No.						Changes Done																				Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'				Sandeep N							17-May-2013				1.0																																					Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'-----------------------------------------------------------------------------------------------------------------------------------
'	Anurag Khera		 26-Sept-2013		2.0								Externalized object hierarchies
'-----------------------------------------------------------------------------------------------------------------------------------
Function Fn_SISW_PC_GetObject(sObjectName)
	Dim sObjectXMLPath
	sObjectXMLPath = Fn_GetEnvValue("User", "AutomationDir") & "\TestData\AutomationXML\ObjectXML\ProductConfigurator.xml"
	Set Fn_SISW_PC_GetObject = Fn_SISW_Setup_GetObjectFromXML(sObjectXMLPath, sObjectName)
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_SISW_PC_NavTreeTableOperations

'Description			 :	Function Used to perform operations on Nav Tree table

'Parameters			   :   1.StrAction: Action Name
'										2.StrNode: Node path
'										3.StrColumn: Column name
'										4.StrGroup: Group name
'										5.StrFamily: Family name
'										6.StrValue: Value name or Expected value
'										7.StrPopupMenu: Popup menu
'										8.StrToolBarOption: Toolbar option to click on toolbar buttons
'
'Return Value		   : 	True or False

'Pre-requisite			:	Nav tree table should be appear

'Examples				:   bReturn=Fn_SISW_PC_NavTreeTableOperations("AddGroup","","","MGroup1","","","","")
'										bReturn=Fn_SISW_PC_NavTreeTableOperations("AddFamily","Group3","","","Family1","","","no")
'										bReturn=Fn_SISW_PC_NavTreeTableOperations("AddValue","Group3:Family1","","","","Val2","","")
'										bReturn=Fn_SISW_PC_NavTreeTableOperations("ModifyCell","Group5:Family1","Comparison Mode","","","Text","","")
'										bReturn=Fn_SISW_PC_NavTreeTableOperations("CellExist","Group5:Family1","Comparison Mode","","","Text","","")
'										bReturn=Fn_SISW_PC_NavTreeTableOperations("VerifyNode","Group2@4","","","","","","")
'										bReturn=Fn_SISW_PC_NavTreeTableOperations("Collapse","Group5:Group2:Val2","","","","","","")
'										bReturn=Fn_SISW_PC_NavTreeTableOperations("Collapse","Group5","","","","","","")
'										bReturn=Fn_SISW_PC_NavTreeTableOperations("ModifyCell","MGroup1","Description","","","Test2","","")
'										bReturn=Fn_SISW_PC_NavTreeTableOperations("Multiselect","Group3:Family1~Group3:Family2","","","","","","")
'										bReturn=Fn_SISW_PC_NavTreeTableOperations("Popupmenuselect","Group3:Family1","","","","","Copy","")
'										bReturn=Fn_SISW_PC_NavTreeTableOperations("Multiselectpopupmenuselect",Group3:Family1~Group3:Family2","","","","","Copy","")
'										bReturn=Fn_SISW_PC_NavTreeTableOperations("Popupmenuenable","Group3","","","","","Add Group","")
'										bReturn=Fn_SISW_PC_NavTreeTableOperations("OpenSavedVariantRulesView","Group3","","","","","","")
'										bReturn=Fn_SISW_PC_NavTreeTableOperations("Popupmenuexist","TestGroup:Color56710:Blue","","","","","Copy","")
'										bReturn=Fn_SISW_PC_NavTreeTableOperations("VerifyObjectName","","","","","000209/A;1-ProductItem04645","","")
'										bReturn=Fn_SISW_PC_NavTreeTableOperations("VerifyValuesFromCellList","Test:Fam","Unit Of Measure","","","ft~gm~kg~km~ml","","")
'										bReturn=Fn_SISW_PC_NavTreeTableOperations("IsCellEditable","Grp2:Fam2","Comparison Mode","","","","","")
'										bReturn=Fn_SISW_PC_NavTreeTableOperations("IsCellEditable","Grp2:Fam2","Description","","","","","")
'										bReturn=Fn_SISW_PC_NavTreeTableOperations("GetNodeIndex","Grp1:Fam:Val","","","","","","")
'										bReturn=Fn_SISW_PC_NavTreeTableOperations("FilterContents","","","","","Family1","","")
'
'History					 :			
'				Developer Name						Date					Rev. No.						Changes Done																				Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'				Sandeep N							2-May-2013				1.0																																				Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'				Pranav Ingle						3-May-2013				1.1						Added Cases  "Multiselect", "Popupmenuselect"									Sandeep N
'																														"Multiselectpopupmenuselect"  & "Popupmenuenable" 
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'				Sandeep N						6-May-2013				1.1						Added Cases  "OpenVariantConstraintsView", "OpenVariantDefaultsView"	Sunny R
'																														"OpenSavedVariantRulesView"
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'				Sandeep N							15-May-2013				1.2						Added Case : Popupmenuexist																							Veena G
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'				Sandeep N							17-May-2013				1.3						Added Case : VerifyObjectName																							Veena G
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'				Sandeep N							17-May-2013				1.4						Added Case : VerifyValuesFromCellList																				Pranav I
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'				Sandeep N							17-May-2013				1.5						Added Case : IsCellEditable																							Anjali M
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'				Sandeep N							20-May-2013				1.6						Added Case : GetNodeIndex																							Anjali M
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'				Anjali M							04-Jul-2013				1.7						Added Case : FilterContents																							Sandeep N
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_SISW_PC_NavTreeTableOperations(StrAction,StrNode,StrColumn,StrGroup,StrFamily,StrValue,StrPopupMenu,StrToolBarOption)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_PC_NavTreeTableOperations"
 	'Declaring variables
	Dim bFlag,iCounter,StrNodePath,iWidth,iTempWidth,iRowNumber,iHieght,iLoopCounter
	Dim sPath,arrNode,iCount,iTempInstance,arrNode1,iInstance,iTempHieght,iCount1
	Dim ObjNavTreeTable,ObjTree,objTableColumn,ObjSubTree,objPopupMenu
	Dim aMenuList, sMenuVal,arrValue,aNode

	'creating object of [ Nav tree table ]
	Set ObjNavTreeTable=JavaWindow("ProductConfigurator").JavaTree("NavTreeTable")
	Fn_SISW_PC_NavTreeTableOperations=False
	Select Case StrAction
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to modify specific cell value
		Case "ModifyCell"
			'Click on cell
			bFlag=Fn_SISW_PC_NavTreeTableOperations("ClickCell",StrNode,StrColumn,"","",StrValue,"","")
			Wait(1)
			If bFlag=True Then
				'checking Edit box is exist
				If JavaWindow("ProductConfigurator").JavaEdit("NavTreeTableEdit").Exist(5) Then
					'Set value in edit box
					Fn_SISW_PC_NavTreeTableOperations=Fn_Edit_Box("Fn_SISW_PC_NavTreeTableOperations",JavaWindow("ProductConfigurator"),"NavTreeTableEdit", StrValue)
					If Fn_SISW_PC_NavTreeTableOperations Then
                    	call Fn_KeyBoardOperation("SendKey", "{TAB}")
                    End If
					Exit function
				'checking Existance of list
				End If

				if JavaWindow("ProductConfigurator").JavaList("NavTreeTableList").Exist(3) Then
					'Checking specific value available in list
                    Fn_SISW_PC_NavTreeTableOperations= Fn_List_Select("Fn_SISW_PC_NavTreeTableOperations", JavaWindow("ProductConfigurator"), "NavTreeTableList",StrValue)
				Else
					bFlag = Fn_SISW_PC_NavTreeTableOperations("ClickCell",StrNode,StrColumn,"","",StrValue,"","")
					wait 2
					if JavaWindow("ProductConfigurator").JavaList("NavTreeTableList").Exist(3) Then
						'Checking specific value available in list
	                    Fn_SISW_PC_NavTreeTableOperations= Fn_List_Select("Fn_SISW_PC_NavTreeTableOperations", JavaWindow("ProductConfigurator"), "NavTreeTableList",StrValue)
	                    If Fn_SISW_PC_NavTreeTableOperations Then
	                    	call Fn_KeyBoardOperation("SendKey", "{TAB}")
	                    End If
					Else
						Fn_SISW_PC_NavTreeTableOperations=False
					End If
				End If
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to add new group
		Case "AddGroup","AddGroup_Type"
			bFlag=False
			'Click on ( Add Group ) button
			If lcase(StrToolBarOption)="no" then
				bFlag=True
			Else
				bFlag=Fn_ToolbarOperation("Click", "Add Group (Ctrl+G)","")
			End If
			If bFlag=true then
				bFlag=False
'				For iCounter=0 to Cint(ObjNavTreeTable.GetROProperty("items count"))-1
'					If Trim(ObjNavTreeTable.GetItem(iCounter))="" Then
'						bFlag=Fn_SISW_PC_NavTreeTableOperations("ClickCell","","Name",StrGroup,"","","","")
'						Exit for
'					End If
'				Next
				For iCounter=0 to Cint(ObjNavTreeTable.GetROProperty("items count"))-1
				aNode = Split(Trim(ObjNavTreeTable.GetItem(iCounter)), ":")
					If aNode(uBound(aNode)) = "" Then
						'Modified by Sandeep : 14-May-2014
						If StrNode<>"" Then
							bFlag=Fn_SISW_PC_NavTreeTableOperations("ClickCell",StrNode+":","Object ID",StrGroup,"","","","")
						Else
							bFlag=Fn_SISW_PC_NavTreeTableOperations("ClickCell","","Object ID",StrGroup,"","","","")
						End If
						Exit for
					End If
				Next
				If bFlag=True Then
					'Setting group name
					If StrAction="AddGroup_Type" Then
						Call Fn_UI_EditBox_Type("Fn_SISW_PC_NavTreeTableOperations",JavaWindow("ProductConfigurator"),"NavTreeTableEdit",StrGroup)
					Else
						Call Fn_Edit_Box("Fn_SISW_PC_NavTreeTableOperations",JavaWindow("ProductConfigurator"),"NavTreeTableEdit", StrGroup)
					End If
					wait 2
					JavaWindow("ProductConfigurator").JavaEdit("NavTreeTableEdit").Activate
					Fn_SISW_PC_NavTreeTableOperations=True
				End If
			End if
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to add new Family
		Case "AddFamily"
			bFlag=False
			bFlag=Fn_SISW_PC_NavTreeTableOperations("Select",StrNode,"","",StrFamily,"","","")
			If bFlag=True Then
				'click on Add Family button
				If lcase(StrToolBarOption)="no" then
					bFlag=True
				Else
					bFlag=Fn_ToolbarOperation("Click", "Add Family (Ctrl+J)","")
				End If
				If bFlag=true then
					bFlag=False
					'bFlag=Fn_SISW_PC_NavTreeTableOperations("ClickCell",StrNode+":","Name","",StrFamily,"","","")
					bFlag=Fn_SISW_PC_NavTreeTableOperations("ClickCell",StrNode+":","Object ID","",StrFamily,"","","")
					If bFlag=true Then
						'Set family  name
						Call Fn_Edit_Box("Fn_SISW_PC_NavTreeTableOperations",JavaWindow("ProductConfigurator"),"NavTreeTableEdit", StrFamily)
						wait 2
						JavaWindow("ProductConfigurator").JavaEdit("NavTreeTableEdit").Activate
						Fn_SISW_PC_NavTreeTableOperations=True
					End If
				End if
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to add Model Family			
		Case "AddModelFamily"
			bFlag=False
			'Click on ( Add Model Family ) button
			If lcase(StrToolBarOption)="no" then
				bFlag=True
			Else
				bFlag=Fn_ToolbarOperation("Click", "Add Model Family (Ctrl+Shift+J)","")
			End If
			If bFlag=true then
				bFlag=False
				For iCounter=0 to Cint(ObjNavTreeTable.GetROProperty("items count"))-1
				aNode = Split(Trim(ObjNavTreeTable.GetItem(iCounter)), ":")
					If aNode(uBound(aNode)) = "" Then			
						If StrNode<>"" Then
							bFlag=Fn_SISW_PC_NavTreeTableOperations("ClickCell",StrNode+":","Object ID",StrGroup,"","","","")
						Else
							bFlag=Fn_SISW_PC_NavTreeTableOperations("ClickCell","","Object ID",StrGroup,"","","","")
						End If
						Exit for
					End If
				Next
				If bFlag=True Then
					'Setting model family name				
						Call Fn_Edit_Box("Fn_SISW_PC_NavTreeTableOperations",JavaWindow("ProductConfigurator"),"NavTreeTableEdit", StrGroup)
					wait 2
					JavaWindow("ProductConfigurator").JavaEdit("NavTreeTableEdit").Activate
					Fn_SISW_PC_NavTreeTableOperations=True
				End If
			End if
			
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to add new Value
		Case "AddValue"
			bFlag=False
			bFlag=Fn_SISW_PC_NavTreeTableOperations("Select",StrNode,"","","",StrValue,"","")
			If bFlag=True Then
				'click on Add Value button
				If lcase(StrToolBarOption)="no" then
					bFlag=True
				Else
					bFlag=Fn_ToolbarOperation("Click", "Add Value (Ctrl+L)","")
				End if
				If bFlag=true then
					bFlag=False
					bFlag=Fn_SISW_PC_NavTreeTableOperations("ClickCell",StrNode+":","Object ID","","",StrValue,"","")
					If bFlag=true Then
						'Setting value
						Call Fn_Edit_Box("Fn_SISW_PC_NavTreeTableOperations",JavaWindow("ProductConfigurator"),"NavTreeTableEdit", StrValue)
						wait 2
						JavaWindow("ProductConfigurator").JavaEdit("NavTreeTableEdit").Activate
						Fn_SISW_PC_NavTreeTableOperations=True
					End If
				End if
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to add product model			
		Case "AddProductModel"
			bFlag=False
			bFlag=Fn_SISW_PC_NavTreeTableOperations("Select",StrNode,"","",StrFamily,"","","")
			If bFlag=True Then
				'click on Add Family button
				If lcase(StrToolBarOption)="no" then
					bFlag=True
				Else
					'bFlag=Fn_ToolbarOperation("Click", "Add Product Model (Ctrl+Shift+P)","")
					bFlag=Fn_ToolbarOperation("Click", "Add Model (Ctrl+Shift+P)","")
				End If
				If bFlag=True then
					bFlag=False
					bFlag=Fn_SISW_PC_NavTreeTableOperations("ClickCell",StrNode+":","Object ID","",StrFamily,"","","")
					If bFlag=True Then
						'Set product model Name
						Call Fn_Edit_Box("Fn_SISW_PC_NavTreeTableOperations",JavaWindow("ProductConfigurator"),"NavTreeTableEdit", StrFamily)
						wait 2
						JavaWindow("ProductConfigurator").JavaEdit("NavTreeTableEdit").Activate
						Fn_SISW_PC_NavTreeTableOperations=True
					End If
				End if
			End If
			
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to check value of specific cell
		Case "CellExist"
			StrNodePath=Fn_SISW_PC_NavTreeTableOperations("GetPath",StrNode,"","","","","","")
			If Trim(ObjNavTreeTable.GetColumnValue(StrNodePath,StrColumn))=Trim(StrValue) Then
				Fn_SISW_PC_NavTreeTableOperations=True
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to varify specific node exist
		Case "VerifyNode"
			StrNodePath=Fn_SISW_PC_NavTreeTableOperations("GetPath",StrNode,"","","","","","")
			If StrNodePath=False Then
				Fn_SISW_PC_NavTreeTableOperations=False
			Else
				Fn_SISW_PC_NavTreeTableOperations=True
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to click on specific cell
		Case "ClickCell"
			iWidth=0
			bFlag=False
			For iCounter=0 to Cint(ObjNavTreeTable.GetROProperty("columns_count"))-1
				iWidth=iWidth+ObjNavTreeTable.Object.getColumn(iCounter).getWidth
				iTempWidth=Cint(ObjNavTreeTable.Object.getColumn(iCounter).getWidth)/2
				iTempWidth=Cint(iTempWidth)/2
				If cstr(StrColumn)=cstr(ObjNavTreeTable.GetColumnHeader(iCounter)) Then
					iWidth=iWidth-iTempWidth
					bFlag=True
					Set objTableColumn=ObjNavTreeTable.Object.getColumn(iCounter)
					ObjNavTreeTable.Object.showColumn objTableColumn
					wait 2
					Exit for
				End If
			Next
			If bFlag=True Then
				iRowNumber=Fn_SISW_PC_NavTreeTableOperations("GetRowNumber",StrNode,"","","","","","")
				iRowNumber=iRowNumber-1
				iHieght=ObjNavTreeTable.Object.getItemHeight
				iTempHieght=iHieght/2
				iHieght=iHieght*iRowNumber
				iHieght=iHieght+iTempHieght
				ObjNavTreeTable.Click iWidth,iHieght,"LEFT"
				If Err.Number < 0 Then
					Fn_SISW_PC_NavTreeTableOperations=False
				Else
					Fn_SISW_PC_NavTreeTableOperations=True
				End if
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to get node path or row number of specific node
		Case "GetPath"
			Set ObjTree=ObjNavTreeTable.Object
			sPath=""
			arrNode=Split(StrNode,":")
			For iCount=0 to ubound(arrNode)
				iTempInstance=1
				bFlag=False
				If arrNode(iCount)="" Then
                    arrNode1(0) = ""
				Else
					arrNode1=Split(arrNode(iCount),"@")
					If instr(1,arrNode(iCount),"@") Then
						iInstance=arrNode1(1)
					Else
						iInstance=1
					End If
				End If
                For iCounter=0 to Cint(ObjTree.getItemCount())-1
'				For iCounter=0 to Cint(ObjNavTreeTable.GetROProperty("items count"))-1
					If ObjTree.getItem(iCounter).getNameText()=arrNode1(0) Then
						If cint(iInstance)=iTempInstance Then
							bFlag=True
							If sPath="" Then
								sPath="#" & iCounter
							Else
								sPath=sPath & ":#"&iCounter
							End If
							Set ObjTree=ObjTree.getItem(iCounter)
							Exit for
						End If
						iTempInstance=iTempInstance+1
					End If
				Next
				If bFlag=False Then
					Exit for
				End If
			Next
			Set ObjTree=Nothing
			If bFlag=False Then
				Fn_SISW_PC_NavTreeTableOperations=False
			Else
				Fn_SISW_PC_NavTreeTableOperations=sPath				
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to get node path or row number of specific node
		Case "GetRowNumber"
			Set ObjTree=ObjNavTreeTable.Object
			sPath=""
			iRowNumber=0
			arrNode=Split(StrNode,":")
			If StrNode="" Then
				ReDim arrNode(0)
				arrNode(0)="@1"
				iLoopCounter=0
			Else	
				iLoopCounter=ubound(arrNode)
			End If
			For iCount=0 to iLoopCounter
				iTempInstance=1
				bFlag=False
				If arrNode(iCount)="" then
					arrNode(iCount)="@1"
				End if
				arrNode1=Split(arrNode(iCount),"@")
				If instr(1,arrNode(iCount),"@") Then
					iInstance=arrNode1(1)
				Else
					iInstance=1
				End If
				For iCounter=0 to Cint(ObjNavTreeTable.GetROProperty("items count"))-1
					iRowNumber=iRowNumber+1
					If ObjTree.getItem(iCounter).getNameText()=arrNode1(0) Then
						If cint(iInstance)=iTempInstance Then
							bFlag=True
							If sPath="" Then
								sPath="#" & iCounter
							Else
								sPath=sPath & ":#"&iCounter
							End If
							Set ObjTree=ObjTree.getItem(iCounter)
							Exit for
						End If
						iTempInstance=iTempInstance+1
					Elseif ObjTree.getItem(iCounter).getItemCount()>0 then
						
						If ObjTree.getItem(iCounter).getExpanded()="true" Then
								iRowNumber=iRowNumber+Cint(ObjTree.getItem(iCounter).getItemCount())
								Set ObjSubTree=ObjTree.getItem(iCounter)
								For iCount1=0 to Cint(ObjTree.getItem(iCounter).getItemCount())-1
									If ObjSubTree.getItem(iCount1).getExpanded()="true" Then
										iRowNumber=iRowNumber+Cint(ObjSubTree.getItem(iCount1).getItemCount())
									End if
								Next
						End If
					End If
				Next
				If bFlag=False Then
					Exit for
				End If
			Next
			Set ObjTree=Nothing
			Set ObjSubTree=Nothing
			If bFlag=False Then
				Fn_SISW_PC_NavTreeTableOperations=False
			Else
				Fn_SISW_PC_NavTreeTableOperations=iRowNumber
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to select node
		Case "Select"
			If inStr(1,StrNode,"@") Then
				StrNodePath=Fn_SISW_PC_NavTreeTableOperations("GetPath",StrNode,"","","","","","")
				If StrNodePath=False Then
					Fn_SISW_PC_NavTreeTableOperations=False
				Else
					ObjNavTreeTable.Select StrNodePath
					Fn_SISW_PC_NavTreeTableOperations=True
				End If
			Else
				ObjNavTreeTable.Select StrNode
				Fn_SISW_PC_NavTreeTableOperations=True
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to Expand node
		Case "Expand"
			If inStr(1,StrNode,"@") Then
				StrNodePath=Fn_SISW_PC_NavTreeTableOperations("GetPath",StrNode,"","","","","","")
				If StrNodePath=False Then
					Fn_SISW_PC_NavTreeTableOperations=False
				Else
					ObjNavTreeTable.Expand StrNodePath
					Fn_SISW_PC_NavTreeTableOperations=True
				End If
			Else
				ObjNavTreeTable.Expand StrNode
				Fn_SISW_PC_NavTreeTableOperations=True
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to Collapse node
		Case "Collapse"
			If inStr(1,StrNode,"@") Then
				StrNodePath=Fn_SISW_PC_NavTreeTableOperations("GetPath",StrNode,"","","","","","")
				If StrNodePath=False Then
					Fn_SISW_PC_NavTreeTableOperations=False
				Else
					ObjNavTreeTable.Collapse StrNodePath
					Fn_SISW_PC_NavTreeTableOperations=True
				End If
			Else
				ObjNavTreeTable.Collapse StrNode
				Fn_SISW_PC_NavTreeTableOperations=True
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to Select Multipale nodes
		Case "Multiselect"
				arrNode = split(StrNode,"~")
				For iCount = 0 to UBound(arrNode)
                    If inStr(1,arrNode(iCount),"@") Then
						StrNodePath=Fn_SISW_PC_NavTreeTableOperations("GetPath",arrNode(iCount),"","","","","","")
						If StrNodePath=False Then
							Fn_SISW_PC_NavTreeTableOperations=False
							Exit Function
						Else
							If iCount = 0 Then
								ObjNavTreeTable.Select StrNodePath
							Else
								ObjNavTreeTable.ExtendSelect StrNodePath
							End If
						End If
					Else
						If iCount = 0 Then
							ObjNavTreeTable.Select arrNode(iCount)
						Else
							ObjNavTreeTable.ExtendSelect arrNode(iCount)
						End If
					End If
        		Next
				Fn_SISW_PC_NavTreeTableOperations=True
			' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
			Case "Popupmenuselect"
				If inStr(1,StrNode,"@") Then
					StrNodePath=Fn_SISW_PC_NavTreeTableOperations("GetPath",StrNode,"","","","","","")
					If StrNodePath=False Then
							Fn_SISW_PC_NavTreeTableOperations=False
					Else
						ObjNavTreeTable.Select StrNodePath
						wait 1
						ObjNavTreeTable.OpenContextMenu StrNodePath
					End If
				Else
					ObjNavTreeTable.Select StrNode
					wait 1
					ObjNavTreeTable.OpenContextMenu StrNode
				End If
				wait 1

				 aMenuList = split(StrPopupMenu,":")
				'Select Menu action
				Select Case Ubound(aMenuList)
					Case "0"
						 StrPopupMenu = JavaWindow("ProductConfigurator").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0))
					Case "1"
						StrPopupMenu = JavaWindow("ProductConfigurator").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1))
					Case "2"
						StrPopupMenu = JavaWindow("ProductConfigurator").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1),aMenuList(2))
					Case Else
						Fn_SISW_PC_NavTreeTableOperations = False
						Exit Function
				End Select
				JavaWindow("ProductConfigurator").WinMenu("ContextMenu").Select StrPopupMenu
				Fn_SISW_PC_NavTreeTableOperations=True
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
			Case "Multiselectpopupmenuselect"
				arrNode = split(StrNode,"~")
				For iCount = 0 to UBound(arrNode)
					If inStr(1,arrNode(iCount),"@") Then
						arrNode(iCount)=Fn_SISW_PC_NavTreeTableOperations("GetPath",arrNode(iCount),"","","","","","")
						If arrNode(iCount)=False Then
							Fn_SISW_PC_NavTreeTableOperations=False
							Exit Function
						End If
					End If

					If iCount = 0 Then
						ObjNavTreeTable.Select arrNode(iCount)
					ElseIf iCount = UBound(arrNode) Then
						ObjNavTreeTable.ExtendSelect arrNode(iCount)
						wait 1
						ObjNavTreeTable.OpenContextMenu arrNode(iCount)
						wait 1
					Else
						ObjNavTreeTable.ExtendSelect arrNode(iCount)
					End If
        		Next

				 aMenuList = split(StrPopupMenu,":")
				'Select Menu action
				Select Case Ubound(aMenuList)
					Case "0"
						 StrPopupMenu = JavaWindow("ProductConfigurator").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0))
					Case "1"
						StrPopupMenu = JavaWindow("ProductConfigurator").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1))
					Case "2"
						StrPopupMenu = JavaWindow("ProductConfigurator").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1),aMenuList(2))
					Case Else
						Fn_SISW_PC_NavTreeTableOperations = False
						Exit Function
				End Select
				JavaWindow("ProductConfigurator").WinMenu("ContextMenu").Select StrPopupMenu
				Fn_SISW_PC_NavTreeTableOperations=True
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "Popupmenuenable","Popupmenuexist"
                If inStr(1,StrNode,"@") Then
					StrNodePath=Fn_SISW_PC_NavTreeTableOperations("GetPath",StrNode,"","","","","","")
					If StrNodePath=False Then
						Fn_SISW_PC_NavTreeTableOperations=False
					Else
						ObjNavTreeTable.Select StrNodePath
						wait 1
						ObjNavTreeTable.OpenContextMenu StrNodePath
					End If
				Else
					ObjNavTreeTable.Select StrNode
					wait 1
					ObjNavTreeTable.OpenContextMenu StrNode
				End If
				wait 1

				aMenuList = split(StrPopupMenu,":")
                Set objPopupMenu=JavaWindow("ProductConfigurator").JavaMenu("label:="&aMenuList(0)&"","index:=0")
				For iCounter=1 to Ubound(aMenuList)
					Set objPopupMenu=objPopupMenu.JavaMenu("label:="&aMenuList(iCounter)&"","index:=0")	
				Next

				Select Case StrAction
					Case "Popupmenuenable"
						If objPopupMenu.Exist(5) Then
							sMenuVal=cint(objPopupMenu.GetROProperty("enabled"))
						Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Popup menu ["+StrPopupMenu+"] not exist on node ["+StrNodePath+"]")
						End If
						If sMenuVal = 1 Then
							Fn_SISW_PC_NavTreeTableOperations = True
						Else
							Fn_SISW_PC_NavTreeTableOperations = False
						End If
					Case "Popupmenuexist"
						If objPopupMenu.Exist(5) then
							Fn_SISW_PC_NavTreeTableOperations = True
						Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Popup menu ["+StrPopupMenu+"] not exist on node ["+StrNodePath+"]")
						End if
				End Select
			
				Call Fn_KeyBoardOperation("SendKeys", "{ESC}")
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "OpenVariantConstraintsView","OpenVariantDefaultsView","OpenSavedVariantRulesView","OpenInclusionRulesView"
            If StrNode <> "" Then
				bFlag=Fn_SISW_PC_NavTreeTableOperations("Select",StrNode,"","","",StrValue,"","")
			End If
            Select Case StrAction
				Case "OpenVariantConstraintsView"
					bFlag="Open Variant Constraints view"
				Case "OpenVariantDefaultsView"
					bFlag="Open Variant Defaults view"
				Case "OpenSavedVariantRulesView"
					bFlag="Open Saved Variant Defaults Rules view"
				Case "OpenInclusionRulesView"
					bFlag="Open Inclusion Rules view"
			End Select
			Fn_SISW_PC_NavTreeTableOperations=Fn_ToolbarOperation("Click", bFlag,"")
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to verify for which object table is open
		Case "VerifyObjectName"
			 JavaWindow("ProductConfigurator").JavaStaticText("NavTreeTableObjectName").SetTOProperty "label",StrValue
			 If JavaWindow("ProductConfigurator").JavaStaticText("NavTreeTableObjectName").Exist(6) Then
				Fn_SISW_PC_NavTreeTableOperations=True
			 End If
	    ' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "VerifyValuesFromCellList"
			'Click on cell
			bFlag=Fn_SISW_PC_NavTreeTableOperations("ClickCell",StrNode,StrColumn,"","",StrValue,"","")
			If bFlag=True Then
				If JavaWindow("ProductConfigurator").JavaList("NavTreeTableList").Exist(10) Then
					arrValue=split(StrValue,"~")
					For iCounter=0 to ubound(arrValue)
						bFlag=False
						For iCount=0 to JavaWindow("ProductConfigurator").JavaList("NavTreeTableList").GetROProperty("items count")-1
							If trim(JavaWindow("ProductConfigurator").JavaList("NavTreeTableList").Object.getItem(iCount))=Trim(arrValue(iCounter)) then
								bFlag=True
								Exit for
							End if
						Next
						If bFlag=False Then
							Exit for
						End If
					Next
				End If
			End if
			If bFlag=True Then
				Fn_SISW_PC_NavTreeTableOperations=True
			End If
	    ' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to check specific cell is editable or not
		Case "IsCellEditable"
				bFlag=Fn_SISW_PC_NavTreeTableOperations("ClickCell",StrNode,StrColumn,"","","","","")
				If bFlag=True Then
					'cheking cell editable or not
					If JavaWindow("ProductConfigurator").JavaEdit("NavTreeTableEdit").Exist(5) Then
						JavaWindow("ProductConfigurator").JavaEdit("NavTreeTableEdit").Activate
						wait 1
						Fn_SISW_PC_NavTreeTableOperations=True
					End if
				End if
        ' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to get index of node
		Case "GetNodeIndex"
			bFlag=False
			For iCounter=0 to Cint(ObjNavTreeTable.GetROProperty("items count"))-1
				If StrNode=ObjNavTreeTable.GetItem(iCounter) Then
					bFlag=True
					Exit for
				End If
			Next
			If bFlag=False Then
				Fn_SISW_PC_NavTreeTableOperations=False
			Else
				Fn_SISW_PC_NavTreeTableOperations=iCounter+1
			End If
        ' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "FilterContents"
			If JavaWindow("ProductConfigurator").JavaEdit("NavTreeFilterTextEdit").Exist(5) Then
				JavaWindow("ProductConfigurator").JavaToolBar("NavTreeClearSearch").Press "Clear the Text Field"
				 wait 1
				Fn_SISW_PC_NavTreeTableOperations=Fn_Edit_Box("Fn_SISW_PC_NavTreeTableOperations",JavaWindow("ProductConfigurator"),"NavTreeFilterTextEdit",StrValue)
				JavaWindow("ProductConfigurator").JavaEdit("NavTreeFilterTextEdit").Activate
				wait 1
			End if
	End Select
	'Releasing object of Nav tree table
	Set ObjNavTreeTable=Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_SISW_CP_VariantNatTableOperations

'Description			 :	Function Used to perform operations on Variant Nat Tables

'Parameters			   :   1.StrAction: Action Name
'										2.StrType: Type Name { Applicibility or Mode}
'										3.StrTabName: Parent Tab name
'										4.StrNode: Node path
'										5.StrColumnName: Column name
'										6.StrValue: Expected value
'										7.StrMessage: Output message
'										8.StrPopupmenu: Popup menu
'
'Return Value		   : 	True or False

'Pre-requisite			:	Variant Nat table should be appear

'Examples				:   bReturn=Fn_SISW_CP_VariantNatTableOperations("SetFlag","Applicability","000057/A;1-TestItem (Variant Expression Editor)","Cooling:Air","Cooling=Air","Block","NOT Cooling = 'Air'","")
'										bReturn=Fn_SISW_CP_VariantNatTableOperations("Save","","","","","","","")
'										bReturn=Fn_SISW_CP_VariantNatTableOperations("VerifyOutput","Model","","","","","NOT Brake Type = 'Drum'","")
'										bReturn=Fn_SISW_CP_VariantNatTableOperations("PopupMenuSelectOnColumn","Applicability","","","Cooling=Air","","","Paste Expression")
'                                       bReturn=Fn_SISW_CP_VariantNatTableOperations("VerifyNode","Model","","Model:Sports~Model:Luxury","","","","")
'										bReturn=Fn_SISW_CP_VariantNatTableOperations("VerifyColumnExist","","","","Faml=Val1~Faml=Val2","","","")
'										bReturn=Fn_SISW_CP_VariantNatTableOperations("SelectColumn","Applicability","","","Color=Red","","","")
'										bReturn=Fn_SISW_CP_VariantNatTableOperations("PopupMenuExistOnColumn","Applicability","","","Color=Red","","","Copy Expression")
'										bReturn=Fn_SISW_CP_VariantNatTableOperations("ModifyNode","Model","","Date:","","10-May-2013","","")
'										bReturn=Fn_SISW_CP_VariantNatTableOperations("ModifyNode","Model","","Date:","","Today","","")
'										bReturn=Fn_SISW_CP_VariantNatTableOperations("ModifyNode","Model","","Date:","","11-May-2013~10-Jun-2013","","")
'										bReturn=Fn_SISW_CP_VariantNatTableOperations("ModifyNode","Model","","Edit:","","5","","")
'										bReturn=Fn_SISW_CP_VariantNatTableOperations("ModifyNode","Model","","Edit:5","","","","")
'										bReturn=Fn_SISW_CP_VariantNatTableOperations("ModifyNode","Model","","Date:5/14/13","","","","")
'                                       bReturn=Fn_SISW_CP_VariantNatTableOperations("SearchContents","Model","","","","Sports","","")
'
'History					 :			
'				Developer Name						Date					Rev. No.						Changes Done																				Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'				Sandeep N							6-May-2013				1.0																																				Sunny R
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'				Sandeep N							7-May-2013				1.1							Added Case : PopupMenuSelectOnColumn									   Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'				Pranav Ingle						7-May-2013				1.1							Modified Case : SetFlag																			   Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'				Sandeep N						8-May-2013				1.2							Added Case : SelectColumn																			 Sukhada B
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'				Sandeep N						9-May-2013				1.2							Added Case : PopupMenuExistOnColumn													Sukhada B
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'				Pranav Ingle					9-May-2013				1.3							Added Case : VerifyRowExist																		Sandeep N
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'				Sandeep N					9-May-2013				1.4							Added Case : ModifyNode																		Sandeep N
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'				Sandeep N					15-May-2013				1.4							Added Case : SelectVariantVariability													Veena G
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'				Sandeep N					15-May-2013				1.5							Added Case : VerifyNode													Rima P
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'				Sandeep N					15-May-2013				1.6							Added Case : SearchContents													Rima P
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_SISW_CP_VariantNatTableOperations(StrAction,StrType,StrTabName,StrNode,StrColumnName,StrValue,StrMessage,StrPopupmenu)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_CP_VariantNatTableOperations"
	 	'Declaring variables
		Dim ObjApplicabilityTable
		Dim iColIndex,iRowIndex,iX,iY,StrCurrentOutput,aNode,iCounter,arrMenu,aCol, aRow,bFlag
        Dim aVal,aSubVal,aColumnNum,iCol
        Dim objRow,iRow

		Fn_SISW_CP_VariantNatTableOperations=False
		'Creating object of [ Applicability ] table
		Set ObjApplicabilityTable=JavaWindow("ProductConfigurator").JavaObject("VariantNatTable")
		If StrType<>"" Then
			If JavaWindow("ProductConfigurator").JavaTab("GeneralInnerTab").Exist(5) Then
				'Click on [ Applicability ] tab
				Call Fn_UI_JavaTab_Select("Fn_SISW_CP_VariantNatTableOperations",JavaWindow("ProductConfigurator"),"GeneralInnerTab", StrType)
				wait 2
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Tab [ Applicability ] is not exist" )
			End If
		End If
		
	   Select Case StrAction
			
	 		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			Case "SetFlag"
				If StrTabName<>"" Then
					'Double click on tab
					If Fn_TabFolder_Operation("DoubleClickTab", StrTabName, "")=False then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Fail to double click on Tab [ "+StrTabName+" ]" )
						Set ObjApplicabilityTable=Nothing
						Exit function
					End if
					wait 2
				End If
                If instr(1,StrColumnName,"ColumnNumber:") Then
					aColumnNum=split(StrColumnName,":")
					iColIndex=Cint(aColumnNum(1))
					iRowIndex=Fn_SISW_RAC_NatTable_TreeTable_GetRowIndex(ObjApplicabilityTable, "",aColumnNum(1), StrNode, 1, "", "","")
				Else
					'Getting column index
					iColIndex = Fn_SISW_RAC_NatTable_GetColumnIndexExt(ObjApplicabilityTable, StrColumnName,"","","")
					'Getting row index
					If Environment.Value("ProductName") = sQTPProductName OR Environment.Value("ProductName") = sUFTProductName Then	'' added code by Ankit T(02-Jul-2014) to handle UFT related issue as some method not supported in UFT 
						iRowIndex=Fn_SISW_RAC_NatTable_TreeTable_GetRowIndex(ObjApplicabilityTable, StrColumnName, "",StrNode, 0, "", "","")
					Else
						iRowIndex=Fn_SISW_RAC_NatTable_TreeTable_GetRowIndex(ObjApplicabilityTable, StrColumnName, "",StrNode, 1, "", "","")
					End If
				End if
				If iColIndex <> -1 and iRowIndex <> -1 Then
                    iRow=Cint(iRowIndex)-1
'                    iCol=Cint(iColIndex)-1
					'Added New Code
					If instr(1,StrColumnName,"ColumnNumber:") Then
						iCol=Cint(iColIndex)-1
					Else
						For iCounter=0 to ObjApplicabilityTable.Object.getCellByPosition(iColIndex, iRowIndex).getSourceLayer().getColumnCount-1
							If Environment.Value("ProductName") = sQTPProductName OR Environment.Value("ProductName") = sUFTProductName Then	'' added code by Ankit T(02-Jul-2014) to handle UFT related issue as some method not supported in UFT 
							   If ObjApplicabilityTable.Object.getCellByPosition(iCounter, iRowIndex).getlayer.getColumnHeaderlayer.getDatavalueByPosition(iCounter,iRowIndex).tostring=CStr(StrColumnName) Then
							       iCol=iCounter-1
								   Exit For
							   End If
							Else
								If ObjApplicabilityTable.Object.getCellByPosition(iColIndex, iRowIndex).getSourceLayer().getDataProvider().getColumnHeaderDataValue(iCounter).toString()=CStr(StrColumnName) Then
									iCol=iCounter-1
									Exit For
								End If
							End If

						Next
					End If
					If Environment.Value("ProductName") = sQTPProductName OR Environment.Value("ProductName") = sUFTProductName Then'' added code by Ankit T(02-Jul-2014) to handle UFT related issue as some method not supported in UFT 
'						If StrType = "Applicability" Then
'								iX = ObjApplicabilityTable.Object.getStartXOfColumnPosition(iColIndex+ iColIndex+1) + 18
'						    iY = ObjApplicabilityTable.Object.getStartYOfRowPosition(iRowIndex) + 4
'                        ElseIf StrType = "Model" Then				         
						    iX = ObjApplicabilityTable.Object.getStartXOfColumnPosition(cdbl(ObjApplicabilityTable.Object.getCellByPosition(iColIndex, iRowIndex).getSourceLayer().getColumnCount)+iColIndex) + 18
						    iY = ObjApplicabilityTable.Object.getStartYOfRowPosition(iRowIndex) + 4
						    If cdbl(ObjApplicabilityTable.Object.getCellByPosition(iColIndex, iRowIndex).getSourceLayer().getRowCount) = iRowIndex Then 
							     iRowIndex = iRowIndex -1
							End if
'						End If
					Else
						iX = ObjApplicabilityTable.Object.getStartXOfColumnPosition(iColIndex) + 18
						iY = ObjApplicabilityTable.Object.getStartYOfRowPosition(iRowIndex) + 4
					End if

					Select Case StrValue
						'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
						Case "Check"
						For iCounter=0 to 2
								If Environment.Value("ProductName") = sQTPProductName OR Environment.Value("ProductName") = sUFTProductName Then'' added code by Ankit T(02-Jul-2014) to handle UFT related issue as some method not supported in UFT 
									Set objRow=ObjApplicabilityTable.Object.getCellByPosition(iColIndex, iRowIndex).getlayer.getRowHeaderLayer.getBaselayer.getCellByPosition(iColIndex,iRowIndex).getSourcelayer.getDataProvider().getRowObjects().get(iRow).getData()
									If ISEmpty(ObjApplicabilityTable.Object.getCellByPosition(iColIndex, iRowIndex).getlayer.getRowHeaderLayer.getBaselayer.getCellByPosition(iColIndex,iRowIndex).getSourcelayer.getDataProvider().getColumnObjects().get(iCol).get().get(objRow)) then
	                                   StrCurrentOutput=""
	                                Else
	                               	    StrCurrentOutput=ObjApplicabilityTable.Object.getCellByPosition(iColIndex, iRowIndex).getlayer.getRowHeaderLayer.getBaselayer.getCellByPosition(iColIndex,iRowIndex).getSourcelayer.getDataProvider().getColumnObjects().get(iCol).get().get(objRow).toString()
	                                End if
								Else
	                                Set objRow=ObjApplicabilityTable.Object.getCellByPosition(iColIndex, iRowIndex).getSourceLayer().getDataProvider().getRowObjects().get(iRow).getData()
									If ObjApplicabilityTable.Object.getCellByPosition(iColIndex, iRowIndex).getSourceLayer().getDataProvider().getColumnObjects().get(iCol).get().get(objRow) is Nothing then
	                                    StrCurrentOutput=""
	                                Else
	                               	    StrCurrentOutput=ObjApplicabilityTable.Object.getCellByPosition(iColIndex, iRowIndex).getSourceLayer().getDataProvider().getColumnObjects().get(iCol).get().get(objRow).toString()
	                                End if
                               End If
								If StrCurrentOutput="TRUE" Then
                                    Fn_SISW_CP_VariantNatTableOperations=True
									Exit for
								Else
									ObjApplicabilityTable.Click iX, iY, "LEFT"
									wait 1
								End If
								
							Next
						'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
						Case "None"
							For iCounter=0 to 2
								Set objRow=ObjApplicabilityTable.Object.getCellByPosition(iColIndex, iRowIndex).getSourceLayer().getDataProvider().getRowObjects().get(iRow).getData()
                                If ObjApplicabilityTable.Object.getCellByPosition(iColIndex, iRowIndex).getSourceLayer().getDataProvider().getColumnObjects().get(iCol).get().get(objRow) is Nothing then
                                    StrCurrentOutput=""
                                Else
								    StrCurrentOutput=ObjApplicabilityTable.Object.getCellByPosition(iColIndex, iRowIndex).getSourceLayer().getDataProvider().getColumnObjects().get(iCol).get().get(objRow).toString()
                                End if
								If StrCurrentOutput<>"" Then
									ObjApplicabilityTable.Click iX, iY, "LEFT"
									wait 1
								Else
                                    Fn_SISW_CP_VariantNatTableOperations=True
									Exit for
								End If
							Next
						'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
						Case "Block"
					For iCounter=0 to 2
					If Environment.Value("ProductName") = sQTPProductName OR Environment.Value("ProductName") = sUFTProductName Then'' added code by Ankit T(02-Jul-2014) to handle UFT related issue as some method not supported in UFT 
						Set objRow=ObjApplicabilityTable.Object.getCellByPosition(iColIndex, iRowIndex).getlayer.getRowHeaderLayer.getBaselayer.getCellByPosition(iColIndex,iRowIndex).getSourcelayer.getDataProvider().getRowObjects().get(iRow).getData()
						If ISEmpty(ObjApplicabilityTable.Object.getCellByPosition(iColIndex, iRowIndex).getlayer.getRowHeaderLayer.getBaselayer.getCellByPosition(iColIndex,iRowIndex).getSourcelayer.getDataProvider().getColumnObjects().get(iCol).get().get(objRow)) then
                           StrCurrentOutput=""
                        Else
                       	    StrCurrentOutput=ObjApplicabilityTable.Object.getCellByPosition(iColIndex, iRowIndex).getlayer.getRowHeaderLayer.getBaselayer.getCellByPosition(iColIndex,iRowIndex).getSourcelayer.getDataProvider().getColumnObjects().get(iCol).get().get(objRow).toString()
                        End if
					Else
                        Set objRow=ObjApplicabilityTable.Object.getCellByPosition(iColIndex, iRowIndex).getSourceLayer().getDataProvider().getRowObjects().get(iRow).getData()
						If ObjApplicabilityTable.Object.getCellByPosition(iColIndex, iRowIndex).getSourceLayer().getDataProvider().getColumnObjects().get(iCol).get().get(objRow) is Nothing then
                            StrCurrentOutput=""
                        Else
                       	    StrCurrentOutput=ObjApplicabilityTable.Object.getCellByPosition(iColIndex, iRowIndex).getSourceLayer().getDataProvider().getColumnObjects().get(iCol).get().get(objRow).toString()
                        End if
                       End If
					If StrCurrentOutput="FALSE" Then
                        Fn_SISW_CP_VariantNatTableOperations=True
						Exit for
					Else
						ObjApplicabilityTable.Click iX, iY, "LEFT"
						wait 1
					End If
					Next
					End Select
				End If
				If StrTabName<>"" Then
					Call Fn_TabFolder_Operation("DoubleClickTab", "*"&StrTabName, "")
					wait 2
				End If
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			Case "Save"
				'Save the changes
				Fn_SISW_CP_VariantNatTableOperations=Fn_ToolbatButtonClick("Save the current contents (Ctrl+S)")
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			Case "VerifyOutput"
				StrCurrentOutput=JavaWindow("ProductConfigurator").JavaEdit("VariantNatTableOutput").GetROProperty("value")
				If instr(1,StrCurrentOutput,StrMessage) Then
					Fn_SISW_CP_VariantNatTableOperations=True
				End if
            '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			Case "PopupMenuSelectOnColumn"
				If instr(1,StrColumnName,"ColumnNumber:") Then
					aColumnNum=split(StrColumnName,":")
					iColIndex=cint(aColumnNum(1))
				Else
					iColIndex = Fn_SISW_RAC_NatTable_GetColumnIndexExt(ObjApplicabilityTable, StrColumnName,"", "","")
				End If
                If iColIndex = -1 Then
					Fn_SISW_CP_VariantNatTableOperations = False
					exit function
				End If
                iX = ObjApplicabilityTable.Object.getStartXOfColumnPosition(iColIndex) + 10
				iY = 20
				wait 2
				ObjApplicabilityTable.Click iX, iY ,"RIGHT"
				Wait 2
				arrMenu=Split(StrPopupmenu,":") 
				Select Case cInt(ubound(arrMenu))
					Case 0				
						JavaWindow("ProductConfigurator").JavaMenu("label:="&arrMenu(0)&"","index:=0").Select
						Fn_SISW_CP_VariantNatTableOperations = True
					Case 1
						JavaWindow("ProductConfigurator").JavaMenu("label:="&arrMenu(0)&"","index:=0").JavaMenu("label:="&arrMenu(1)&"","index:=0").Select
						Fn_SISW_CP_VariantNatTableOperations = True
				End Select
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	 		Case "VerifyColumnExist"
				aCol=split(StrColumnName,"~")
				For iCounter=0 to ubound(aCol)
					bFlag=False
                    If instr(1,aCol(iCounter),"ColumnNumber:") Then
						aColumnNum=split(aCol(iCounter),":")
						iColIndex = Fn_SISW_RAC_NatTable_GetColumnIndexExt(ObjApplicabilityTable, "",aColumnNum(1),"","")
					Else
						iColIndex = Fn_SISW_RAC_NatTable_GetColumnIndexExt(ObjApplicabilityTable, aCol(iCounter),"","","")
					End If
					iColIndex = Fn_SISW_RAC_NatTable_GetColumnIndexExt(ObjApplicabilityTable, aCol(iCounter),"","","")
					If iColIndex <> -1 then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS : Column [ "+aCol(iCounter)+" ] exist in table" )
						bFlag=True
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Column [ "+aCol(iCounter)+" ] does not exist in table" )
						bFlag=False
						Exit for
					End if
				Next
				If bFlag=True Then
					Fn_SISW_CP_VariantNatTableOperations=true
				End If
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			Case "VerifyNode"
					If StrColumnName<>"" Then
						If Environment.Value("ProductName") = sQTPProductName OR Environment.Value("ProductName") = sUFTProductName Then
							iRowIndex=Fn_SISW_RAC_NatTable_TreeTable_GetRowIndex(ObjApplicabilityTable, StrColumnName, "",StrNode, 0, "", "","")
					    Else
					    	iRowIndex=Fn_SISW_RAC_NatTable_TreeTable_GetRowIndex(ObjApplicabilityTable, StrColumnName, "",StrNode, 1, "", "","")
					    End If
					Else
						iRowIndex = Fn_SISW_RAC_NatTable_Tree_GetRowIndex(ObjApplicabilityTable, "", "",StrNode, 1, "", "","")
					End If
					
					If iRowIndex = -1 OR iRowIndex = false then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : node [ "+StrNode+" ] does not exist in table" )
						Fn_SISW_CP_VariantNatTableOperations=False
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS : node [ "+StrNode+" ] exist in table" )
						Fn_SISW_CP_VariantNatTableOperations=True
					End if
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	 		Case "VerifyRowExist"
				aRow=split(StrNode,"~")
				For iCounter=0 to ubound(aRow)
					bFlag=False
					iRowIndex = Fn_SISW_RAC_NatTable_TreeTable_GetRowIndex(ObjApplicabilityTable, "", "",aRow(iCounter), 1, "", "","")
					If iRowIndex = -1 OR iRowIndex = false then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Column [ "+aRow(iCounter)+" ] does not exist in table" )
						bFlag=False
						Exit for
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS : Column [ "+aRow(iCounter)+" ] exist in table" )
						bFlag=True
					End if
				Next
				If bFlag=True Then
					Fn_SISW_CP_VariantNatTableOperations=true
				End If
            '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			Case "SelectColumn"
                If instr(1,StrColumnName,"ColumnNumber:") Then
					aColumnNum=split(StrColumnName,":")
					iColIndex = Fn_SISW_RAC_NatTable_GetColumnIndexExt(ObjApplicabilityTable, "",aColumnNum(1),"","")
				Else
					iColIndex = Fn_SISW_RAC_NatTable_GetColumnIndexExt(ObjApplicabilityTable, StrColumnName,"","","")
				End if
                If iColIndex = -1 Then
					Fn_SISW_CP_VariantNatTableOperations = False
					exit function
				End If
                iX = ObjApplicabilityTable.Object.getStartXOfColumnPosition(iColIndex) + 10
				iY = 20
				wait 2
				ObjApplicabilityTable.Click iX, iY ,"LEFT"
				Wait 2
				Fn_SISW_CP_VariantNatTableOperations = True
            '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
            Case "PopupMenuExistOnColumn"
				If instr(1,StrColumnName,"ColumnNumber:") Then
					aColumnNum=split(StrColumnName,":")
					iColIndex = Fn_SISW_RAC_NatTable_GetColumnIndexExt(ObjApplicabilityTable, "",aColumnNum(1),"","")
				Else
					iColIndex = Fn_SISW_RAC_NatTable_GetColumnIndexExt(ObjApplicabilityTable, StrColumnName,"","","")
				End if
                If iColIndex = -1 Then
					Fn_SISW_CP_VariantNatTableOperations = False
					exit function
				End If
                iX = ObjApplicabilityTable.Object.getStartXOfColumnPosition(iColIndex) + 10
				iY = 20
				wait 2
				ObjApplicabilityTable.Click iX, iY ,"RIGHT"
				Wait 2
				arrMenu=Split(StrPopupmenu,":") 
				Select Case cInt(ubound(arrMenu))
					Case 0				
						Fn_SISW_CP_VariantNatTableOperations=JavaWindow("ProductConfigurator").JavaMenu("label:="&arrMenu(0)&"","index:=0").Exist(5)
					Case 1
						Fn_SISW_CP_VariantNatTableOperations = JavaWindow("ProductConfigurator").JavaMenu("label:="&arrMenu(0)&"","index:=0").JavaMenu("label:="&arrMenu(1)&"","index:=0").Exist(5)
				End Select
                Call Fn_KeyBoardOperation("SendKeys", "{ESC}")
				wait 1
            '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			'Case to modify node value
			Case "ModifyNode"
				iRowIndex=Fn_SISW_RAC_NatTable_TreeTable_GetRowIndex(ObjApplicabilityTable, "","",StrNode , 1, "", "","")
				If iRowIndex <> -1 Then
					iX = 20
					iY = ObjApplicabilityTable.Object.getStartYOfRowPosition(iRowIndex) + 4
					ObjApplicabilityTable.Click iX, iY, "LEFT"
					wait 2
					'If value is blank then case make it empty
					If StrValue="" Then
						If JavaWindow("ProductConfigurator").JavaEdit("VariantNatTableEdit").Exist(5) Then
							Fn_SISW_CP_VariantNatTableOperations=Fn_Edit_Box("Fn_SISW_CP_VariantNatTableOperations",JavaWindow("ProductConfigurator"),"VariantNatTableEdit", "")
							wait 1
							JavaWindow("ProductConfigurator").JavaEdit("VariantNatTableEdit").Activate
						End If
					Else
						If JavaWindow("ProductConfigurator").JavaObject("VariantNatTableCalendarButton").Exist(3) Then
							JavaWindow("ProductConfigurator").JavaObject("VariantNatTableCalendarButton").Click 1,1,"LEFT"
							If JavaWindow("ProductConfigurator").JavaWindow("DateRangeControl").Exist(5) Then
								aVal=Split(StrValue,"~")
								Select Case ubound(aVal)
									'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
									'case to set single date
									Case 0
										If lcase(aVal(0))="today" Then
											Fn_SISW_CP_VariantNatTableOperations=Fn_Button_Click("Fn_SISW_CP_VariantNatTableOperations",JavaWindow("ProductConfigurator").JavaWindow("DateRangeControl"), "Today")
											wait 1
											JavaWindow("ProductConfigurator").JavaEdit("VariantNatTableEdit").Activate
										Else
											aSubVal=Split(aVal(0)," ")
											JavaWindow("ProductConfigurator").JavaWindow("DateRangeControl").JavaCalendar("DateTime").SetDate aSubVal(0)
											If ubound(aSubVal)=1 Then
												JavaWindow("ProductConfigurator").JavaWindow("DateRangeControl").JavaCalendar("DateTime").SetTime aSubVal(0)
											End If
											Fn_SISW_CP_VariantNatTableOperations=Fn_Button_Click("Fn_SISW_CP_VariantNatTableOperations",JavaWindow("ProductConfigurator").JavaWindow("DateRangeControl"), "OK")
											wait 1
											JavaWindow("ProductConfigurator").JavaEdit("VariantNatTableEdit").Activate
										End If
									'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
									'case to set From date to To date
									Case 1
										Call Fn_CheckBox_Set("Fn_SISW_CP_VariantNatTableOperations", JavaWindow("ProductConfigurator").JavaWindow("DateRangeControl"), "RangeDate","ON")
										Call Fn_UI_JavaRadioButton_SetON("Fn_SISW_CP_VariantNatTableOperations",JavaWindow("ProductConfigurator").JavaWindow("DateRangeControl"), "FromDate")
										aSubVal=Split(aVal(0)," ")
										JavaWindow("ProductConfigurator").JavaWindow("DateRangeControl").JavaCalendar("DateTime").SetDate aSubVal(0)
										If ubound(aSubVal)=1 Then
											JavaWindow("ProductConfigurator").JavaWindow("DateRangeControl").JavaCalendar("DateTime").SetTime aSubVal(0)
										End If
										wait 1
										Call Fn_UI_JavaRadioButton_SetON("Fn_SISW_CP_VariantNatTableOperations",JavaWindow("ProductConfigurator").JavaWindow("DateRangeControl"), "ToDate")
										aSubVal=Split(aVal(1)," ")
										JavaWindow("ProductConfigurator").JavaWindow("DateRangeControl").JavaCalendar("DateTime").SetDate aSubVal(0)
										If ubound(aSubVal)=1 Then
											JavaWindow("ProductConfigurator").JavaWindow("DateRangeControl").JavaCalendar("DateTime").SetTime aSubVal(0)
										End If
										Fn_SISW_CP_VariantNatTableOperations=Fn_Button_Click("Fn_SISW_CP_VariantNatTableOperations",JavaWindow("ProductConfigurator").JavaWindow("DateRangeControl"), "OK")
										wait 1
										JavaWindow("ProductConfigurator").JavaEdit("VariantNatTableEdit").Activate
								End Select
							End If
						Elseif JavaWindow("ProductConfigurator").JavaEdit("VariantNatTableEdit").Exist(3) then
							Fn_SISW_CP_VariantNatTableOperations=Fn_Edit_Box("Fn_SISW_CP_VariantNatTableOperations",JavaWindow("ProductConfigurator"),"VariantNatTableEdit", StrValue)
							JavaWindow("ProductConfigurator").JavaEdit("VariantNatTableEdit").Activate
						End If
					End If
				End If
            '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			Case "SelectVariantVariability"
					'Checking existance of [ VariantVariability ] list
					If JavaWindow("ProductConfigurator").JavaList("VariantVariability").Exist(5) Then
						If Fn_UI_ListItemExist("Fn_SISW_CP_VariantNatTableOperations", JavaWindow("ProductConfigurator"), "VariantVariability",StrValue) then
							Call Fn_List_Select("Fn_SISW_CP_VariantNatTableOperations",  JavaWindow("ProductConfigurator"), "VariantVariability",StrValue)
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS : Successfully Selected Variant Variability type [ "+StrValue+" ]")
							Fn_SISW_CP_VariantNatTableOperations=True
						Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Variant Variability type [ "+StrValue+" ] not exist in list" )
						End if
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Variant Variability list  not exist" )
					End If
            '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			Case "SearchContents"
                If JavaWindow("ProductConfigurator").JavaEdit("VariantSearch").Exist(5) Then
                    If Fn_ToolbarButtonClick_Ext(1,"Clear the Text Field") then
                        wait 1
                        Fn_SISW_CP_VariantNatTableOperations=Fn_Edit_Box("Fn_SISW_CP_VariantNatTableOperations",JavaWindow("ProductConfigurator"),"VariantSearch",StrValue)
                        JavaWindow("ProductConfigurator").JavaEdit("VariantSearch").Activate
                        wait 1
                   End if
                End if
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			Case "VerifyTabActivate"
				StrCurrentOutput=JavaWindow("ProductConfigurator").JavaTab("GeneralInnerTab").GetROProperty("value")
				If instr(1,StrCurrentOutput,StrTabName) Then
					Fn_SISW_CP_VariantNatTableOperations=True
				End if
		End Select

		'Releasing object of [ Applicability ] table
		Set ObjApplicabilityTable=Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_SISW_PC_VariantConstraintsViewTableOperations

'Description			 :	Function Used to perform operations on Variant Constraint view Table

'Parameters			   :   1.StrAction: Action Name
'										2.dicVariantConstraintsInfo: Variant Constraint Information
'
'Return Value		   : 	True or False

'Pre-requisite			:	Variant Constraint view Table should be appear

'Examples				:   Dim dicVariantConstraintsInfo
'										Set dicVariantConstraintsInfo = CreateObject( "Scripting.Dictionary" )
'										dicVariantConstraintsInfo("ConstraintType")="Exclusive"
'										bReturn=Fn_SISW_PC_VariantConstraintsViewTableOperations("SetConstraintType",dicVariantConstraintsInfo)
'										
'										dicVariantConstraintsInfo("ConstraintType")="Exclusive"
'										dicVariantConstraintsInfo("Instance")="1"
'										dicVariantConstraintsInfo("ColumnName")="Message Type"
'										dicVariantConstraintsInfo("Value")="Info"
'										bReturn=Fn_SISW_PC_VariantConstraintsViewTableOperations("SetCellData",dicVariantConstraintsInfo)
'										
'										dicVariantConstraintsInfo("ConstraintType")="Exclusive"
'										dicVariantConstraintsInfo("Instance")="1"
'										dicVariantConstraintsInfo("ColumnName")="Message"
'										dicVariantConstraintsInfo("Value")="Passible warnign"
'										bReturn=Fn_SISW_PC_VariantConstraintsViewTableOperations("SetCellData",dicVariantConstraintsInfo)
'										
'										dicVariantConstraintsInfo("ConstraintType")="Exclusive"
'										dicVariantConstraintsInfo("Instance")="1"
'										dicVariantConstraintsInfo("ColumnName")="Model"
'										bReturn=Fn_SISW_PC_VariantConstraintsViewTableOperations("DoubleClickCell",dicVariantConstraintsInfo)
'										
'										dicVariantConstraintsInfo("ConstraintType")="Exclusive"
'										dicVariantConstraintsInfo("Instance")="1"
'										dicVariantConstraintsInfo("ColumnName")="Message"
'										dicVariantConstraintsInfo("Value")="Passible warnign"
'										bReturn=Fn_SISW_PC_VariantConstraintsViewTableOperations("VerifyCellExist",dicVariantConstraintsInfo)
'
'										dicVariantConstraintsInfo("ConstraintType")="Exclusive~Exclusive~Exclusive"
'										dicVariantConstraintsInfo("Instance")="1~2~3"    			[ '::=>   Which instance you want select of Constraint  ]
'										bReturn=Fn_SISW_PC_VariantConstraintsViewTableOperations("SelectRow",dicVariantConstraintsInfo)
'
'										dicVariantConstraintsInfo("Info")="ON"
'										dicVariantConstraintsInfo("Warning")="ON"
'										dicVariantConstraintsInfo("Error")="Off"
'										dicVariantConstraintsInfo("Exclusive")="ON"
'										dicVariantConstraintsInfo("Inclusive")="ON"
'										dicVariantConstraintsInfo("Group")="Group1"
'										dicVariantConstraintsInfo("Family")="Family1"
'										dicVariantConstraintsInfo("Value")="Val1"
'										bReturn= Fn_SISW_PC_VariantConstraintsViewTableOperations("Search",dicVariantConstraintsInfo)
'
'                                        dicVariantConstraintsInfo("Value")="Info"
'                                        bReturn=Fn_SISW_PC_VariantConstraintsViewTableOperations("FilterContents",dicVariantConstraintsInfo)
'
'                                        dicVariantConstraintsInfo("ConstraintType")="Exclusive"
'                                        dicVariantConstraintsInfo("ColumnName")="Message Type"
'                                        dicVariantConstraintsInfo("Value")="Error~Info~Warning"
'                                        bReturn=Fn_SISW_PC_VariantConstraintsViewTableOperations("VerifyValuesFromCellList",dicVariantConstraintsInfo)
'
'History					 :			
'				Developer Name						Date					Rev. No.						Changes Done																				Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'				Sandeep N							6-May-2013				1.0																																				Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'				Pranav Ingle						9-May-2013				1.1							Added Case "SelectRow"																	Sandeep N
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'				Sandeep N							13-May-2013				1.2						Added Case "Search"																				Rima P
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'				Sandeep N							16-May-2013				1.3						Added Case "FilterContents"																    Rima P
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'				Sandeep N							18-May-2013				1.4						Added Case "VerifyValuesFromCellList"										    Rima P
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_SISW_PC_VariantConstraintsViewTableOperations(StrAction,dicVariantConstraintsInfo)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_PC_VariantConstraintsViewTableOperations"
	Dim ObjConstraintTable,WshShell
	Dim iTempInstance,iCounter,arrColName,arrValue,iCount,bFlag
	Dim arrRows, arrInstance, DeviceReplay
	Dim DictKeys,DictItems

	Fn_SISW_PC_VariantConstraintsViewTableOperations=False
 	'Creating object of [ Variant Constraint View ] Table
	Set ObjConstraintTable=JavaWindow("ProductConfigurator").JavaTable("VariantConstraintViewTable")

	Select Case StrAction
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "VerifyCellExist"
			arrColName=Split(dicVariantConstraintsInfo("ColumnName"),"~")
			arrValue=Split(dicVariantConstraintsInfo("Value"),"~")

			If dicVariantConstraintsInfo("Instance")="" then
				dicVariantConstraintsInfo("Instance")=1
			End if
			
			For iCount=0 to ubound(arrColName)
				iTempInstance=1
				bFlag=False
				For iCounter=0 to ObjConstraintTable.GetROProperty("rows")-1
					'verify [ Constraint Type ] match 
					If ObjConstraintTable.GetCellData(iCounter,"Constraint Type")=dicVariantConstraintsInfo("ConstraintType") then
                        If IsNumeric(arrValue(iCount)) Then
							arrValue(iCount)=Cint(arrValue(iCount))
						End If
						If Cint(dicVariantConstraintsInfo("Instance"))=iTempInstance Then
							If ObjConstraintTable.GetCellData(iCounter,arrColName(iCount))=arrValue(iCount) then
								bFlag=True
								Exit for
							Else
								Exit for
							End if
						End if
						iTempInstance=iTempInstance+1
					End if
				Next
				If bFlag=False Then
					Exit for
				End If
			Next
			If bFlag=True Then
				Fn_SISW_PC_VariantConstraintsViewTableOperations=True
			End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "DoubleClickCell"
			If dicVariantConstraintsInfo("Instance")="" then
				dicVariantConstraintsInfo("Instance")=1
			End if
			iTempInstance=1
			For iCounter=0 to ObjConstraintTable.GetROProperty("rows")-1
				If ObjConstraintTable.GetCellData(iCounter,"Constraint Type")=dicVariantConstraintsInfo("ConstraintType") then
					If Cint(dicVariantConstraintsInfo("Instance"))=iTempInstance Then
						'Double click on specific cell
						ObjConstraintTable.ActivateCell iCounter,dicVariantConstraintsInfo("ColumnName")
						wait 2
						Fn_SISW_PC_VariantConstraintsViewTableOperations=True
						Exit for
					End if
					iTempInstance=iTempInstance+1
				End if
			Next
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "SetCellData"
			arrColName=Split(dicVariantConstraintsInfo("ColumnName"),"~")
			arrValue=Split(dicVariantConstraintsInfo("Value"),"~")

			If dicVariantConstraintsInfo("Instance")="" then
				dicVariantConstraintsInfo("Instance")=1
			End if

			For iCount=0 to ubound(arrColName)
				iTempInstance=1
				bFlag=False
				For iCounter=0 to ObjConstraintTable.GetROProperty("rows")-1
					If ObjConstraintTable.GetCellData(iCounter,"Constraint Type")=dicVariantConstraintsInfo("ConstraintType") then
						If Cint(dicVariantConstraintsInfo("Instance"))=iTempInstance Then
							ObjConstraintTable.SelectCell iCounter,arrColName(iCount)
							If JavaWindow("ProductConfigurator").JavaList("VariantConstraintTableList").Exist(2) Then
								bFlag=Fn_List_Select("Fn_SISW_PC_VariantConstraintsViewTableOperations", JavaWindow("ProductConfigurator"), "VariantConstraintTableList",arrValue(iCount))
								wait 2
								Call Fn_KeyBoardOperation("SendKeys", "{ENTER}")
								Exit for
							Elseif JavaWindow("ProductConfigurator").JavaEdit("VariantConstraintTableEdit").Exist(2) then
								bFlag=Fn_Edit_Box("Fn_SISW_PC_VariantConstraintsViewTableOperations",JavaWindow("ProductConfigurator"),"VariantConstraintTableEdit", arrValue(iCount))
								wait 2
								JavaWindow("ProductConfigurator").JavaEdit("VariantConstraintTableEdit").Activate
								Exit for
							End If
						End If
						iTempInstance=iTempInstance+1
					End if
				Next
				If bFlag=False Then
					Exit for
				End If
			Next
			If bFlag=True Then
				Fn_SISW_PC_VariantConstraintsViewTableOperations=True
			End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "SetConstraintType"
			For iCounter=0 to ObjConstraintTable.GetROProperty("rows")-1
				If ObjConstraintTable.GetCellData(iCounter,"Constraint Type")="*" Then
                    ObjConstraintTable.Object.showItem iCounter
					wait 1
					ObjConstraintTable.SelectCell iCounter,"Constraint Type"
					wait 1
					If JavaWindow("ProductConfigurator").JavaList("VariantConstraintTableList").Exist(3) Then
						Fn_SISW_PC_VariantConstraintsViewTableOperations=Fn_List_Select("Fn_SISW_PC_VariantConstraintsViewTableOperations", JavaWindow("ProductConfigurator"), "VariantConstraintTableList",dicVariantConstraintsInfo("ConstraintType"))
						Call Fn_KeyBoardOperation("SendKeys", "{ENTER}")
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS : Successfully add Contraint Type [ "+dicVariantConstraintsInfo("ConstraintType")+" ]")
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Contraint Type list not exist")
					End If
				End If
			Next
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "SelectRow"
			arrRows=Split(dicVariantConstraintsInfo("ConstraintType"),"~")
			arrInstance=Split(dicVariantConstraintsInfo("Instance"),"~")

			Set DeviceReplay = CreateObject("Mercury.DeviceReplay")

            For iCount = 0 To ubound(arrRows)
				If arrInstance(iCount)="" then
					arrInstance(iCount)=1
				End if
				iTempInstance=1
				For iCounter=0 to ObjConstraintTable.GetROProperty("rows")-1
					If ObjConstraintTable.GetCellData(iCounter,"Constraint Type")=arrRows(iCount) then
						If Cint(arrInstance(iCount))=iTempInstance Then
							If iCount = 0 Then
								ObjConstraintTable.ClickCell iCounter, "Constraint"
							Else
								DeviceReplay.KeyDown 29
                                ObjConstraintTable.ClickCell iCounter, "Constraint"
                                DeviceReplay.KeyUp 29
							End If
							wait 2
							Exit for
						End if
						iTempInstance=iTempInstance+1
					End if
				Next
			Next
			Fn_SISW_PC_VariantConstraintsViewTableOperations=True
        '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Search"
			DictKeys = dicVariantConstraintsInfo.Keys
			DictItems = dicVariantConstraintsInfo.Items
			For iCount=0 to dicVariantConstraintsInfo.count-1
				Select Case DictKeys(iCount)
					Case "Info","Warning","Error","Exclusive","Inclusive"
						Call Fn_CheckBox_Set("Fn_SISW_PC_VariantConstraintsViewTableOperations", JavaWindow("ProductConfigurator"),DictKeys(iCount),DictItems(iCount))
					Case "Group","Family","Value"
						Call Fn_Edit_Box("Fn_SISW_PC_VariantConstraintsViewTableOperations", JavaWindow("ProductConfigurator"),DictKeys(iCount),DictItems(iCount))
				End Select
			Next
			Fn_SISW_PC_VariantConstraintsViewTableOperations=Fn_Button_Click("Fn_SISW_PC_VariantConstraintsViewTableOperations", JavaWindow("ProductConfigurator"), "Search")
        '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			Case "FilterContents"
                If JavaWindow("ProductConfigurator").JavaEdit("VariantViewTableFilter").Exist(5) Then
                    JavaWindow("ProductConfigurator").JavaToolbar("VariantViewTableClearSearch").Press "Clear the Text Field"
                     wait 1
                    Fn_SISW_PC_VariantConstraintsViewTableOperations=Fn_Edit_Box("Fn_SISW_PC_VariantConstraintsViewTableOperations",JavaWindow("ProductConfigurator"),"VariantViewTableFilter",dicVariantConstraintsInfo("Value"))
                    JavaWindow("ProductConfigurator").JavaEdit("VariantViewTableFilter").Activate
                    wait 1
                End if
            '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
            'Case to verify values fromcell List
            Case "VerifyValuesFromCellList"
                If dicVariantConstraintsInfo("Instance")="" then
                    dicVariantConstraintsInfo("Instance")=1
                End if
                iTempInstance=1
                bFlag=False
                'Click on specific cell
                For iCounter=0 to ObjConstraintTable.GetROProperty("rows")-1
                    If ObjConstraintTable.GetCellData(iCounter,"Constraint Type")=dicVariantConstraintsInfo("ConstraintType") then
                        If Cint(dicVariantConstraintsInfo("Instance"))=iTempInstance Then
                            ObjConstraintTable.SelectCell iCounter,dicVariantConstraintsInfo("ColumnName")
                            bFlag=True
                         End if
                         iTempInstance=iTempInstance+1
                        End if
                 Next
                 'Checking Existance of [ VariantConstraintTableList ]
                 If JavaWindow("ProductConfigurator").JavaList("VariantConstraintTableList").Exist(5) and bFlag=True Then
                    arrValue=split(dicVariantConstraintsInfo("Value"),"~")
                    For iCount=0 to ubound(arrValue)
                        bFlag=False
                        For iCounter=0 to JavaWindow("ProductConfigurator").JavaList("VariantConstraintTableList").GetROProperty("items count")-1
                            if trim(arrValue(iCount))= trim(JavaWindow("ProductConfigurator").JavaList("VariantConstraintTableList").Object.getItem(iCounter)) then
                                bFlag=True
                                Exit for
                            End if
                        Next
                        If bFlag=False then
                             Exit for
                        End if
                    Next
                 End if
                 Call Fn_KeyBoardOperation("SendKeys", "{ENTER}")
                 If bFlag=True then
                     Fn_SISW_PC_VariantConstraintsViewTableOperations=true
                 End if
				 
			Case "VerifyCellDataOrder"  'Added By Pritam Shikare

				'Collect the Data of the Column
				For iCounter=0 to ObjConstraintTable.GetROProperty("rows")-1
					If ObjConstraintTable.GetCellData(iCounter,dicVariantConstraintsInfo("ColumnName")) <> "" AND ObjConstraintTable.GetCellData(iCounter,dicVariantConstraintsInfo("ColumnName")) <> "*" Then
						sColumnData = CStr(ObjConstraintTable.GetCellData(iCounter,dicVariantConstraintsInfo("ColumnName")))
						Exit For
					End If
				Next

				For iCount=iCounter+1 to ObjConstraintTable.GetROProperty("rows")-1
					If ObjConstraintTable.GetCellData(iCount,dicVariantConstraintsInfo("ColumnName")) <> "*" Then
						If  ObjConstraintTable.GetCellData(iCount,dicVariantConstraintsInfo("ColumnName")) <> "" Then
							sColumnData = sColumnData+"~"+CStr(ObjConstraintTable.GetCellData(iCount,dicVariantConstraintsInfo("ColumnName")))
						End If
					Else
						Exit For
					End If
			Next

			'Split the Data into an array
			aColumnData = Split(sColumnData,"~",-1,1)
			'Verify the order
			Fn_SISW_PC_VariantConstraintsViewTableOperations = Fn_SISW_StringArraySort(aColumnData,dicVariantConstraintsInfo("Order"),"VerifyOrder")
	 '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "GetColumnData"  'Added by Pritam S.

			'Collect the Data of the Column
			For iCounter=0 to ObjConstraintTable.GetROProperty("rows")-1
				If ObjConstraintTable.GetCellData(iCounter,dicVariantConstraintsInfo("ColumnName")) <> "" AND ObjConstraintTable.GetCellData(iCounter,dicVariantConstraintsInfo("ColumnName")) <> "*" Then
					sColumnData = Cstr(ObjConstraintTable.GetCellData(iCounter,dicVariantConstraintsInfo("ColumnName")))
					Exit For
				End If
			Next

			For iCount=iCounter+1 to ObjConstraintTable.GetROProperty("rows")-1
				If ObjConstraintTable.GetCellData(iCount,dicVariantConstraintsInfo("ColumnName")) <> "*"  Then
					If  ObjConstraintTable.GetCellData(iCount,dicVariantConstraintsInfo("ColumnName")) <> "" Then
						sColumnData = sColumnData+"~"+CStr(ObjConstraintTable.GetCellData(iCount,dicVariantConstraintsInfo("ColumnName")))
					End If
				Else
					Exit For
				End If
			Next
			Fn_SISW_PC_VariantConstraintsViewTableOperations = sColumnData
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to verify for which object table is open
		Case "VerifyObjectName"
			 JavaWindow("ProductConfigurator").JavaStaticText("VariantViewTableObjectName").SetTOProperty "label",dicVariantConstraintsInfo("Value")
			 If JavaWindow("ProductConfigurator").JavaStaticText("VariantViewTableObjectName").Exist(6) Then
				Fn_SISW_PC_VariantConstraintsViewTableOperations=True
			 End If
	End Select
	Set ObjConstraintTable=Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_SISW_PC_VariantDefaultsViewTableOperations

'Description			 :	Function Used to perform operations on Variant Default view Table

'Parameters			   :   1.StrAction: Action Name
'										2.dicVariantDefaultsInfo: Variant Default Information
'
'Return Value		   : 	True or False

'Pre-requisite			:	Variant Default view Table should be appear

'Examples				:   Dim dicVariantDefaultsInfo,bReturn
'										Set dicVariantDefaultsInfo = CreateObject( "Scripting.Dictionary" )
'										dicVariantDefaultsInfo("OptionFamily")="Cooling"
'										bReturn=Fn_SISW_PC_VariantDefaultsViewTableOperations("SetOptionFamily",dicVariantDefaultsInfo)
'										
'										dicVariantDefaultsInfo("OptionFamily")="Cooling"
'										dicVariantDefaultsInfo("Instance")="1"
'										dicVariantDefaultsInfo("ColumnName")="Option Value"
'										dicVariantDefaultsInfo("Value")="Air"
'										bReturn=Fn_SISW_PC_VariantDefaultsViewTableOperations("SetCellData",dicVariantDefaultsInfo)
'										
'										dicVariantDefaultsInfo("OptionFamily")="Cooling"
'										dicVariantDefaultsInfo("Instance")="1"
'										dicVariantDefaultsInfo("ColumnName")="Model"
'										bReturn=Fn_SISW_PC_VariantDefaultsViewTableOperations("DoubleClickCell",dicVariantDefaultsInfo)
'										
'										dicVariantDefaultsInfo("OptionFamily")="Cooling"
'										dicVariantDefaultsInfo("Instance")="1"
'										dicVariantDefaultsInfo("ColumnName")="Option Value"
'										dicVariantDefaultsInfo("Value")="Air"
'										bReturn=Fn_SISW_PC_VariantDefaultsViewTableOperations("VerifyCellExist",dicVariantDefaultsInfo)

'										dicVariantDefaultsInfo("OptionFamily")="Cooling~Cooling~Accessories"
'										dicVariantDefaultsInfo("Instance")="1~2~1"
'										bReturn=Fn_SISW_PC_VariantDefaultsViewTableOperations("SelectRow",dicVariantDefaultsInfo)
'
'										dicVariantDefaultsInfo("Group")="Group1"
'										dicVariantDefaultsInfo("Family")="Family1"
'										dicVariantDefaultsInfo("Value")="Val1"
'										bReturn= Fn_SISW_PC_VariantDefaultsViewTableOperations("Search",dicVariantDefaultsInfo)
'
'										bReturn= Fn_SISW_PC_VariantDefaultsViewTableOperations("GetAllColumnName","")
'
'										dicVariantDefaultsInfo("Value")="000210/A;1-TestItem"
'										bReturn= Fn_SISW_PC_VariantDefaultsViewTableOperations("VerifyObjectName",dicVariantDefaultsInfo)
'
'										dicVariantDefaultsInfo("OptionFamily")="Color"
'										dicVariantDefaultsInfo("Value")="Color~Test"
'										bReturn=Fn_SISW_PC_VariantDefaultsViewTableOperations("CellListVerify",dicVariantDefaultsInfo)
'										
'										dicVariantDefaultsInfo("OptionFamily")="Color"
'										dicVariantDefaultsInfo("ColumnName")="Option Value"
'										dicVariantDefaultsInfo("Value")="Red~Green~Blue"
'										bReturn=Fn_SISW_PC_VariantDefaultsViewTableOperations("CellListVerify",dicVariantDefaultsInfo)
'
'History					 :			
'				Developer Name						Date					Rev. No.						Changes Done																				Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'				Sandeep N							7-May-2013				1.0																																			Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'				Pranav Ingle						7-May-2013				1.0							Added Case "SelectRow"																	Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'				Pranav Ingle						14-May-2013				1.0							Modified Case "VerifyCellExist"															Sandeep N
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'				Pranav Ingle						14-May-2013				1.1							Added Case "Search"															Pranav I
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'				Sandeep N							17-May-2013				1.2							Added Case :GetAllColumnName 									Veena G
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'				Veena G							17-May-2013				1.3							Added Case :VerifyObjectName 									Sandeep N
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'				Sandeep N							5-Jul-2013				1.4							Added Case :CellListVerify 									Anjali M
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_SISW_PC_VariantDefaultsViewTableOperations(StrAction,dicVariantDefaultsInfo)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_PC_VariantDefaultsViewTableOperations"
	Dim ObjDefaultTable,arrRows, arrInstance, DeviceReplay
	Dim iTempInstance,iCounter,arrColName,arrValue,iCount,bFlag,StrColumnName,bFlag1
	Dim DictKeys,DictItems,iCnt


	Fn_SISW_PC_VariantDefaultsViewTableOperations=False
 	'Creating object of [ Variant Default View ] Table
	Set ObjDefaultTable=JavaWindow("ProductConfigurator").JavaTable("VariantDefaultViewTable")

	Select Case StrAction
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "VerifyCellExist"
            arrColName=Split(dicVariantDefaultsInfo("ColumnName"),"~")
			arrValue=Split(dicVariantDefaultsInfo("Value"),"~")

			If dicVariantDefaultsInfo("Instance")="" then
				dicVariantDefaultsInfo("Instance")=1
			End if
			For iCount=0 to ubound(arrColName)
				iTempInstance=1
				bFlag=False
				For iCounter=0 to ObjDefaultTable.GetROProperty("rows")-1
					'verify [ Default Type ] match 
					If ObjDefaultTable.GetCellData(iCounter,"Option Family")=dicVariantDefaultsInfo("OptionFamily") then
                        If IsNumeric(arrValue(iCount)) Then
							arrValue(iCount)=Cint(arrValue(iCount))
						End If
						If Cint(dicVariantDefaultsInfo("Instance"))=iTempInstance Then
							If ObjDefaultTable.GetCellData(iCounter,arrColName(iCount))=arrValue(iCount) then
								bFlag=True
								Exit for
							Else
								Exit for
							End if
						End if
						iTempInstance=iTempInstance+1
					End if
				Next
                If bFlag=False Then
					Exit for
				End If
			Next
            If bFlag=True Then
				Fn_SISW_PC_VariantDefaultsViewTableOperations=True
			End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "DoubleClickCell"
			If dicVariantDefaultsInfo("Instance")="" then
				dicVariantDefaultsInfo("Instance")=1
			End if
			iTempInstance=1
			For iCounter=0 to ObjDefaultTable.GetROProperty("rows")-1
				If ObjDefaultTable.GetCellData(iCounter,"Option Family")=dicVariantDefaultsInfo("OptionFamily") then
					If Cint(dicVariantDefaultsInfo("Instance"))=iTempInstance Then
						'Double click on specific cell
						ObjDefaultTable.ActivateCell iCounter,dicVariantDefaultsInfo("ColumnName")
						wait 2
						Fn_SISW_PC_VariantDefaultsViewTableOperations=True
						Exit for
					End if
					iTempInstance=iTempInstance+1
				End if
			Next
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "SelectRow"
			arrRows=Split(dicVariantDefaultsInfo("OptionFamily"),"~")
			arrInstance=Split(dicVariantDefaultsInfo("Instance"),"~")

			Set DeviceReplay = CreateObject("Mercury.DeviceReplay")

            For iCount = 0 To ubound(arrRows)
				If arrInstance(iCount)="" then
					arrInstance(iCount)=1
				End if
				iTempInstance=1
				For iCounter=0 to ObjDefaultTable.GetROProperty("rows")-1
					If ObjDefaultTable.GetCellData(iCounter,"Option Family")=arrRows(iCount) then
						If Cint(arrInstance(iCount))=iTempInstance Then
							If iCount = 0 Then
								ObjDefaultTable.ClickCell iCounter, "Applicability Expression"
							Else
								DeviceReplay.KeyDown 29
                                ObjDefaultTable.ClickCell iCounter, "Applicability Expression"
                                DeviceReplay.KeyUp 29
							End If
							wait 2
							Fn_SISW_PC_VariantDefaultsViewTableOperations=True
							Exit for
						End if
						iTempInstance=iTempInstance+1
					End if
				Next
			Next
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "SetCellData"
            arrColName=Split(dicVariantDefaultsInfo("ColumnName"),"~")
			arrValue=Split(dicVariantDefaultsInfo("Value"),"~")

			If dicVariantDefaultsInfo("Instance")="" then
				dicVariantDefaultsInfo("Instance")=1
			End if
			iTempInstance=1
            For iCount=0 to ubound(arrColName)
				iTempInstance=1
				bFlag=False
				For iCounter=0 to ObjDefaultTable.GetROProperty("rows")-1
					If ObjDefaultTable.GetCellData(iCounter,"Option Family")=dicVariantDefaultsInfo("OptionFamily") then
						If Cint(dicVariantDefaultsInfo("Instance"))=iTempInstance Then
							ObjDefaultTable.SelectCell iCounter,dicVariantDefaultsInfo("ColumnName")
							wait 2
							If JavaWindow("ProductConfigurator").JavaList("VariantDefaultTableList").Exist(2) Then
								bFlag=Fn_List_Select("Fn_SISW_PC_VariantDefaultsViewTableOperations", JavaWindow("ProductConfigurator"), "VariantDefaultTableList",dicVariantDefaultsInfo("Value"))
								wait 2
								Call Fn_KeyBoardOperation("SendKeys", "{ENTER}")
								Exit for
							Elseif JavaWindow("ProductConfigurator").JavaEdit("VariantDefaultTableEdit").Exist(2) then
								bFlag=Fn_Edit_Box("Fn_SISW_PC_VariantDefaultsViewTableOperations",JavaWindow("ProductConfigurator"),"VariantDefaultTableEdit", dicVariantDefaultsInfo("Value"))
								wait 2
								JavaWindow("ProductConfigurator").JavaEdit("VariantDefaultTableEdit").Activate
								Exit for
							End If
						End If
						iTempInstance=iTempInstance+1
					End if
				Next
                If bFlag=False Then
					Exit for
				End If
			Next
            If bFlag=True Then
				Fn_SISW_PC_VariantDefaultsViewTableOperations=True
			End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "SetCellDataExt"
            arrColName=Split(dicVariantDefaultsInfo("ColumnName"),"~")
			arrValue=Split(dicVariantDefaultsInfo("Value"),"~")

			If dicVariantDefaultsInfo("Instance")="" then
				dicVariantDefaultsInfo("Instance")=1
			End if
			iTempInstance=1
            For iCount=0 to ubound(arrColName)
				iTempInstance=1
				bFlag=False
				For iCounter=0 to ObjDefaultTable.GetROProperty("rows")-1
'					If ObjDefaultTable.GetCellData(iCounter,"Option Family")=dicVariantDefaultsInfo("OptionFamily") then
					If Instr(1,ObjDefaultTable.GetCellData(iCounter,"Option Family"),dicVariantDefaultsInfo("OptionFamily")) then
						If Cint(dicVariantDefaultsInfo("Instance"))=iTempInstance Then
							ObjDefaultTable.ActivateCell iCounter,dicVariantDefaultsInfo("ColumnName")
							wait 1
							ObjDefaultTable.SelectCell iCounter,dicVariantDefaultsInfo("ColumnName")
							wait 2
							If JavaWindow("ProductConfigurator").JavaList("VariantDefaultTableList").Exist(2) Then

                                Set objDeviceReplay = CreateObject("Mercury.DeviceReplay")
								objDeviceReplay.SendString "*"
								wait 1
								Set objDeviceReplay =Nothing
								
								Call Fn_Button_Click("Fn_SISW_PC_VariantDefaultsViewTableOperations", JavaWindow("ProductConfigurator"), "VariantDefaultViewTable_ContentSearch")
								wait 3
								For iCnt=0 to JavaWindow("ProductConfigurator").JavaList("VariantDefaultTableList").GetROProperty("items count")-1
									If instr(1,JavaWindow("ProductConfigurator").JavaList("VariantDefaultTableList").Object.getItem(iCnt),dicVariantDefaultsInfo("Value")) Then
										JavaWindow("ProductConfigurator").JavaList("VariantDefaultTableList").Select "#" & iCnt
										wait 1
										bFlag=True
										Exit For
									End If
								Next
								wait 2
								Call Fn_KeyBoardOperation("SendKeys", "{ENTER}")
								Exit for
							Elseif JavaWindow("ProductConfigurator").JavaEdit("VariantDefaultTableEdit").Exist(2) then
								bFlag=Fn_Edit_Box("Fn_SISW_PC_VariantDefaultsViewTableOperations",JavaWindow("ProductConfigurator"),"VariantDefaultTableEdit", dicVariantDefaultsInfo("Value"))
								wait 2
								JavaWindow("ProductConfigurator").JavaEdit("VariantDefaultTableEdit").Activate
								Exit for
							End If
						End If
						iTempInstance=iTempInstance+1
					End if
				Next
                If bFlag=False Then
					Exit for
				End If
			Next
            If bFlag=True Then
				Fn_SISW_PC_VariantDefaultsViewTableOperations=True
			End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "SetOptionFamily"
			For iCounter=0 to ObjDefaultTable.GetROProperty("rows")-1
				If ObjDefaultTable.GetCellData(iCounter,"Option Family")="*" Then
                    ObjDefaultTable.Object.showItem iCounter
                    wait 1
					ObjDefaultTable.ActivateCell iCounter,"Option Family"
					wait 2
					ObjDefaultTable.SelectCell iCounter,"Option Family"
					wait 3
					If JavaWindow("ProductConfigurator").JavaList("VariantDefaultTableList").Exist(3) Then

						Set objDeviceReplay = CreateObject("Mercury.DeviceReplay")
						objDeviceReplay.SendString dicVariantDefaultsInfo("OptionFamily")
						wait 1
						Set objDeviceReplay =Nothing
						
						Call Fn_Button_Click("Fn_SISW_PC_VariantDefaultsViewTableOperations", JavaWindow("ProductConfigurator"), "VariantDefaultViewTable_ContentSearch")
						wait 3
						For iCount=0 to JavaWindow("ProductConfigurator").JavaList("VariantDefaultTableList").GetROProperty("items count")-1
							If instr(1,JavaWindow("ProductConfigurator").JavaList("VariantDefaultTableList").Object.getItem(iCount),dicVariantDefaultsInfo("OptionFamily")) Then
								JavaWindow("ProductConfigurator").JavaList("VariantDefaultTableList").Select "#" & iCount
								wait 1
								Fn_SISW_PC_VariantDefaultsViewTableOperations=True
								Exit For
							End If
						Next

						'Fn_SISW_PC_VariantDefaultsViewTableOperations=Fn_List_Select("Fn_SISW_PC_VariantDefaultsViewTableOperations", JavaWindow("ProductConfigurator"), "VariantDefaultTableList",dicVariantDefaultsInfo("OptionFamily"))
						Call Fn_KeyBoardOperation("SendKeys", "{ENTER}")
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS : Successfully add Option Family [ "+dicVariantDefaultsInfo("OptionFamily")+" ]")
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Option Family list not exist")
					End If
				End If
			Next
        '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Search"
			DictKeys = dicVariantDefaultsInfo.Keys
			DictItems = dicVariantDefaultsInfo.Items
			For iCount=0 to dicVariantDefaultsInfo.count-1
				Select Case DictKeys(iCount)
					Case "Group","Family","Value"
						Call Fn_Edit_Box("Fn_SISW_PC_VariantDefaultsViewTableOperations", JavaWindow("ProductConfigurator"),DictKeys(iCount),DictItems(iCount))
				End Select
			Next
			Fn_SISW_PC_VariantDefaultsViewTableOperations=Fn_Button_Click("Fn_SISW_PC_VariantDefaultsViewTableOperations", JavaWindow("ProductConfigurator"), "Search")
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	   Case "GetAllColumnName"
			StrColumnName=""
			StrColumnName= ObjDefaultTable.GetColumnName(0)
			For iCounter=1 to ObjDefaultTable.GetROProperty("cols")-1
				StrColumnName=StrColumnName+"~"+ObjDefaultTable.GetColumnName(iCounter)		
			Next
			If StrColumnName<>"" Then
				Fn_SISW_PC_VariantDefaultsViewTableOperations=StrColumnName
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to verify for which object table is open
		Case "VerifyObjectName"
			 JavaWindow("ProductConfigurator").JavaStaticText("VariantViewTableObjectName").SetTOProperty "label",dicVariantDefaultsInfo("Value")
			 If JavaWindow("ProductConfigurator").JavaStaticText("VariantViewTableObjectName").Exist(6) Then
				Fn_SISW_PC_VariantDefaultsViewTableOperations=True
			 End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "CellListVerify"
			arrValue=Split(dicVariantDefaultsInfo("Value"),"~")
			If dicVariantDefaultsInfo("ColumnName")="" Then
				dicVariantDefaultsInfo("ColumnName")="Option Family"
			End If
			If dicVariantDefaultsInfo("Instance")="" then
				dicVariantDefaultsInfo("Instance")=1
			End if
			iTempInstance=1
			bFlag1=False
			For iCounter=0 to ObjDefaultTable.GetROProperty("rows")-1
					bFlag=False
					If ObjDefaultTable.GetCellData(iCounter,"Option Family")=dicVariantDefaultsInfo("OptionFamily") then
						If Cint(dicVariantDefaultsInfo("Instance"))=iTempInstance Then
							bFlag1=True
							ObjDefaultTable.SelectCell iCounter,dicVariantDefaultsInfo("ColumnName")
							wait 1
							ObjDefaultTable.SelectCell iCounter,dicVariantDefaultsInfo("ColumnName")
							wait 1
							If JavaWindow("ProductConfigurator").JavaWindow("VariantDefaultTableShell").JavaList("VariantDefaultTableList").Exist(2) Then
								For iCount=0 to ubound(arrValue)
									bFlag=Fn_UI_ListItemExist("Fn_SISW_PC_VariantDefaultsViewTableOperations", JavaWindow("ProductConfigurator").JavaWindow("VariantDefaultTableShell"), "VariantDefaultTableList",arrValue(iCount))
									If bFlag=False Then
										Exit for
									End If
								Next
							End If
						End If
						iTempInstance=iTempInstance+1
					End if
					If bFlag1=True Then
						Exit for
					End If
			Next
			wait 1
            If bFlag=True Then
				Fn_SISW_PC_VariantDefaultsViewTableOperations=True
			End If
			Call Fn_KeyBoardOperation("SendKeys", "{ESC}")
		Case "GetID"
            arrValue=Split(dicVariantDefaultsInfo("OptionFamily"),"~")
			sID=""

			If dicVariantDefaultsInfo("Instance")="" then
				dicVariantDefaultsInfo("Instance")=1
			End if

			For iCount=0 to ubound(arrValue)
					iTempInstance=1
					bFlag=False
					For iCounter=0 to ObjDefaultTable.GetROProperty("rows")-1
						'verify [ Default Type ] match 
						If InStr(1,ObjDefaultTable.GetCellData(iCounter,"Option Family"),arrValue(iCount)) then
							If Cint(dicVariantDefaultsInfo("Instance"))=iTempInstance Then
									If sID="" Then
										sID=ObjDefaultTable.GetCellData(iCounter,"Object ID")
									Else
										sID=sID & "~" & ObjDefaultTable.GetCellData(iCounter,"Object ID")
									End If
									bFlag=True
									Exit for
							End if
							iTempInstance=iTempInstance+1
						End if
					Next
					If bFlag=False Then
						Exit For
					End If
			Next
            If bFlag=True Then
				Fn_SISW_PC_VariantDefaultsViewTableOperations=sID
			End If
	End Select
	Set ObjDefaultTable=Nothing
End Function

'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'    Function Name			:	Fn_SISW_PC_DeleteVariants
'
'    Description				:	Function Used to create Collaborative Design
'
'    Parameters			       :	1. sAction		: Action to be performed
'										       2. sDelMessage :  Send Messge only if you want verify 
'											   3. sBtnName		:  Button TO Click
'
'    Return Value		   	   	: 	True Or False
'
'    Examples					:	Call Fn_SISW_PC_DeleteVariants("Toolbar", "Do you want to delete the selected Variant object?", "Yes")
'
'	   History					:	
'				Developer Name				Date				Rev. No.		Changes Done								Reviewer
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'				Pranav Ingle				8-May-2013			1.0					Created												Sandeep
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_PC_DeleteVariants(sAction, sDelMessage, sBtnName)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_PC_DeleteVariants"
	Dim objDeleteVariants
	Fn_SISW_PC_DeleteVariants = False

	Set objDeleteVariants = JavaWindow("ProductConfigurator").JavaWindow("DeleteVariants")
    
	Select Case sAction
		Case "Menu"  
			' ************************************************************Case To delete Item from Menu option *****************************************************************
			Call Fn_MenuOperation("Select","Edit:Delete")
		Case "Toolbar" 
			'************************************* Case To Delete Item from Toolbar option and click delete icon************************************************************
			Call Fn_ToolbatButtonClick("Delete (Delete)")
			' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_PC_DeleteVariants ] Invalid case [ " & sAction & " ].")
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
	End Select

	'set sDelMessage
	If sDelMessage <> "" Then
		If sDelMessage <> objDeleteVariants.JavaStaticText("DoYouWantToDelete").GetROProperty("value") Then
			Exit Function
		End If
	End If

	' click on OK
	If sBtnName <> "" Then
		Call Fn_Button_Click("Fn_SISW_PC_DeleteVariants", objDeleteVariants,sBtnName)
	End If

	If  Fn_SISW_PC_DeleteVariants <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [ Fn_SISW_PC_DeleteVariants ] executed successfuly with case [ " & sAction & " ].")
	End If

	Fn_SISW_PC_DeleteVariants = True
	Set objDeleteVariants = Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_SISW_PC_VariantSaveRulesViewTableOperations

'Description			 :	Function Used to perform operations on Variant Save Rules view Table

'Parameters			   :   1.StrAction: Action Name
'										2.dicVariantSaveRulesInfo: Variant Save Rules Information
'
'Return Value		   : 	True or False

'Pre-requisite			:	Variant Save Rules view Table should be appear

'Examples				:   Dim dicVariantSaveRulesInfo
'										Set dicVariantSaveRulesInfo = CreateObject( "Scripting.Dictionary" )
'										dicVariantSaveRulesInfo("Name")="SaveRule1"
'										bReturn= Fn_SISW_PC_VariantSaveRulesViewTableOperations("AddName",dicVariantSaveRulesInfo)
'										
'										dicVariantSaveRulesInfo("Name")="SaveRule1"
'										dicVariantSaveRulesInfo("ColumnName")="Description"
'										dicVariantSaveRulesInfo("Value")="Desc"
'										bReturn= Fn_SISW_PC_VariantSaveRulesViewTableOperations("VerifyCellExist",dicVariantSaveRulesInfo)
'										
'										dicVariantSaveRulesInfo("Name")="SaveRule1"
'										dicVariantSaveRulesInfo("ColumnName")="Description"
'										dicVariantSaveRulesInfo("Value")="Desc"
'										bReturn= Fn_SISW_PC_VariantSaveRulesViewTableOperations("SetCellData",dicVariantSaveRulesInfo)
'										
'										dicVariantSaveRulesInfo("Name")="SaveRule1"
'										bReturn= Fn_SISW_PC_VariantSaveRulesViewTableOperations("SelectRow",dicVariantSaveRulesInfo)
'
'										bReturn= Fn_SISW_PC_VariantSaveRulesViewTableOperations("GetAllColumnName","")
'
'History					 :			
'				Developer Name						Date					Rev. No.						Changes Done																				Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'				Sandeep N								9-May-2013				1.0																																				Sunny R
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'				Veena G								2-Jul-2013				1.1							Added case : GetAllColumnName														Sandeep N
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_SISW_PC_VariantSaveRulesViewTableOperations(StrAction,dicVariantSaveRulesInfo)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_PC_VariantSaveRulesViewTableOperations"
	Dim ObjSaveRulesTable,arrRows, arrInstance, DeviceReplay,objConfigRuleTable
	Dim iTempInstance,iCounter,arrColName,arrValue,iCount,bFlag
	Fn_SISW_PC_VariantSaveRulesViewTableOperations=False
 	'Creating object of [ Variant Default View ] Table
	Set ObjSaveRulesTable=JavaWindow("ProductConfigurator").JavaTable("VariantSaveRulesViewTable")
	Set objConfigRuleTable = Fn_PC_GetObject("ConfiguratorRules")
	Select Case StrAction
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "VerifyCellExist"
            arrColName=Split(dicVariantSaveRulesInfo("ColumnName"),"~")
			arrValue=Split(dicVariantSaveRulesInfo("Value"),"~")

			If dicVariantSaveRulesInfo("Instance")="" then
				dicVariantSaveRulesInfo("Instance")=1
			End if
			For iCount=0 to ubound(arrColName)
				iTempInstance=1
				bFlag=False
				For iCounter=0 to ObjSaveRulesTable.GetROProperty("rows")-1
					'verify [ Name ] match 
					If ObjSaveRulesTable.GetCellData(iCounter,"Name")=dicVariantSaveRulesInfo("Name") then
						If IsNumeric(arrValue(iCount)) Then
							arrValue(iCount)=Cint(arrValue(iCount))
						End If
						If Cint(dicVariantSaveRulesInfo("Instance"))=iTempInstance Then
							If ObjSaveRulesTable.GetCellData(iCounter,arrColName(iCount))=arrValue(iCount) then
								bFlag=True
								Exit for
							Else
								Exit for
							End if
						End if
						iTempInstance=iTempInstance+1
					End if
				Next
                If bFlag=False Then
					Exit for
				End If
			Next
            If bFlag=True Then
				Fn_SISW_PC_VariantSaveRulesViewTableOperations=True
			End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "SelectRow"
			arrRows=Split(dicVariantSaveRulesInfo("Name"),"~")
			If dicVariantSaveRulesInfo("Instance")<>"" Then
				arrInstance=Split(dicVariantSaveRulesInfo("Instance"),"~")
			Else
				arrInstance=split(1,"~")
			End If

			Set DeviceReplay = CreateObject("Mercury.DeviceReplay")

            For iCount = 0 To ubound(arrRows)
				If arrInstance(iCount)="" then
					arrInstance(iCount)=1
				End if
				iTempInstance=1
				For iCounter=0 to ObjSaveRulesTable.GetROProperty("rows")-1
					If ObjSaveRulesTable.GetCellData(iCounter,"Name")=arrRows(iCount) then
						If Cint(arrInstance(iCount))=iTempInstance Then
							If iCount = 0 Then
								ObjSaveRulesTable.ClickCell iCounter, "Name"
							Else
								DeviceReplay.KeyDown 29
                                ObjSaveRulesTable.ClickCell iCounter, "Name"
                                DeviceReplay.KeyUp 29
							End If
							wait 2
							Fn_SISW_PC_VariantSaveRulesViewTableOperations=True
							Exit for
						End if
						iTempInstance=iTempInstance+1
					End if
				Next
			Next
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "SetCellData"
            arrColName=Split(dicVariantSaveRulesInfo("ColumnName"),"~")
			arrValue=Split(dicVariantSaveRulesInfo("Value"),"~")

			If dicVariantSaveRulesInfo("Instance")="" then
				dicVariantSaveRulesInfo("Instance")=1
			End if
			iTempInstance=1
            For iCount=0 to ubound(arrColName)
				iTempInstance=1
				bFlag=False
				For iCounter=0 to ObjSaveRulesTable.GetROProperty("rows")-1
					If ObjSaveRulesTable.GetCellData(iCounter,"Name")=dicVariantSaveRulesInfo("Name") then
						If Cint(dicVariantSaveRulesInfo("Instance"))=iTempInstance Then
							ObjSaveRulesTable.SelectCell iCounter,dicVariantSaveRulesInfo("ColumnName")
							wait 2
							If JavaWindow("ProductConfigurator").JavaEdit("VariantSaveRulesTableEdit").Exist(2) then
								bFlag=Fn_Edit_Box("Fn_SISW_PC_VariantSaveRulesViewTableOperations",JavaWindow("ProductConfigurator"),"VariantSaveRulesTableEdit", dicVariantSaveRulesInfo("Value"))
								wait 2
								JavaWindow("ProductConfigurator").JavaEdit("VariantSaveRulesTableEdit").Activate
								Exit for
							End If
						End If
						iTempInstance=iTempInstance+1
					End if
				Next
                If bFlag=False Then
					Exit for
				End If
			Next
            If bFlag=True Then
				Fn_SISW_PC_VariantSaveRulesViewTableOperations=True
			End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "AddName"
			For iCounter=0 to ObjSaveRulesTable.GetROProperty("rows")-1
				If ObjSaveRulesTable.GetCellData(iCounter,"Name")="*" or ObjSaveRulesTable.GetCellData(iCounter,"Name")=""  Then
					ObjSaveRulesTable.Object.showItem iCounter
					wait 1
					ObjSaveRulesTable.SelectCell iCounter,"Name"
					wait 1
					If JavaWindow("ProductConfigurator").JavaEdit("VariantSaveRulesTableEdit").Exist(2) Then
						Fn_SISW_PC_VariantSaveRulesViewTableOperations=Fn_Edit_Box("Fn_SISW_PC_VariantSaveRulesViewTableOperations",JavaWindow("ProductConfigurator"),"VariantSaveRulesTableEdit", dicVariantSaveRulesInfo("Name"))
						wait 2
						JavaWindow("ProductConfigurator").JavaEdit("VariantSaveRulesTableEdit").Activate
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS : Successfully add Name [ "+dicVariantSaveRulesInfo("Name")+" ]")
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Fail to set name")
					End If
				End If
			Next
        '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "GetAllColumnName"
			StrColumnName=""
			StrColumnName= ObjSaveRulesTable.GetColumnName(0)
			For iCounter=1 to ObjSaveRulesTable.GetROProperty("cols")-1
				StrColumnName=StrColumnName+"~"+ObjSaveRulesTable.GetColumnName(iCounter)		
			Next
			If StrColumnName<>"" Then
				Fn_SISW_PC_VariantSaveRulesViewTableOperations=StrColumnName
			End If
		'Case to check specific cell is editable or not
		Case "IsCellEditable"
			For iCounter=0 to ObjSaveRulesTable.GetROProperty("rows")-1
				If ObjSaveRulesTable.GetCellData(iCounter,"Name")= dicVariantSaveRulesInfo("Name") Then
						ObjSaveRulesTable.Object.showItem iCounter
						wait 1
						ObjSaveRulesTable.SelectCell iCounter,dicVariantSaveRulesInfo("ColumnName")
						wait 1
						If JavaWindow("ProductConfigurator").JavaEdit("VariantSaveRulesTableEdit").Exist(3) Then
							Fn_SISW_PC_VariantSaveRulesViewTableOperations = True
							JavaWindow("ProductConfigurator").JavaEdit("VariantSaveRulesTableEdit").Activate
						End If
						Exit For
				End If
			Next
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 	
		Case "CreateSVRule"
			iRowIndex = cint(objConfigRuleTable.Object.getItemCount())-1 'Get Current Row Index to create new rule
			 IF dicVariantSaveRulesInfo("ColumnName") <> "" And dicVariantSaveRulesInfo("ColumnValues") <> "" Then
				arrColumns = Split(dicVariantSaveRulesInfo("ColumnName"),"~")
				arrColValues = Split(dicVariantSaveRulesInfo("ColumnValues"),"~")
				For iCnt = 0 to UBound(arrColumns) 'loop to enter values to column
					sAppColName = Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\AutomationXML\ApplicationInformationXML\RAC_ConfiguratorTable_APL.xml",arrColumns(iCnt))
					If sAppColName = False Then 
						sAppColName = arrColumns(iCnt)
					End If
					objConfigRuleTable.ActivateCell iRowIndex,sAppColName
					wait 3
					JavaWindow("ProductConfigurator").JavaEdit("VariantSaveRulesTableEdit").Set dicVariantSaveRulesInfo("ColumnValues")
					wait 2
				
				Next
				objConfigRuleTable.ActivateRow iRowIndex
				Wait 1
				If dicVariantSaveRulesInfo("Save") <> "" Then
					sMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("RAC_Toolbar"),"Savethecurrentcontents")
					Call Fn_ToolBarOperation("Click",sMenu,"") 'Save Rule
	 				Call Fn_ReadyStatusSync(1)
				End If
	
				sAppColValue = objConfigRuleTable.GetCellData(iRowIndex,sAppColName)
				If IsEmpty(sAppColValue) or sAppColValue = "" Then
					Fn_SISW_PC_VariantSaveRulesViewTableOperations = False
				Else
					Fn_SISW_PC_VariantSaveRulesViewTableOperations = sAppColValue				
				End If
			 End IF	
	End Select
	Set ObjSaveRulesTable=Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_SISW_PC_InclusionRulesViewTableOperations

'Description			 :	Function Used to perform operations on Variant Default view Table

'Parameters			   :   1.StrAction: Action Name
'										2.dicInclusionRulesInfo: Variant Default Information
'
'Return Value		   : 	True or False

'Pre-requisite			:	Variant Default view Table should be appear

'Examples				:   Dim dicInclusionRulesInfo,bReturn
'										Set dicInclusionRulesInfo = CreateObject( "Scripting.Dictionary" )
'										dicInclusionRulesInfo("Severity")="Error"
'										bReturn=Fn_SISW_PC_InclusionRulesViewTableOperations("SetSeverity",dicInclusionRulesInfo)
'										
'										dicInclusionRulesInfo("Severity")="Information"
'										dicInclusionRulesInfo("Instance")="1"
'										dicInclusionRulesInfo("ColumnName")="Message"
'										dicInclusionRulesInfo("Value")="Air"
'										bReturn=Fn_SISW_PC_InclusionRulesViewTableOperations("SetCellData",dicInclusionRulesInfo)
'										
'										dicInclusionRulesInfo("Severity")="Information"
'										dicInclusionRulesInfo("Instance")="1"
'										dicInclusionRulesInfo("ColumnName")="Message"
'										bReturn=Fn_SISW_PC_InclusionRulesViewTableOperations("DoubleClickCell",dicInclusionRulesInfo)
'										
'										dicInclusionRulesInfo("Severity")="Error"
'										dicInclusionRulesInfo("Instance")="1"
'										dicInclusionRulesInfo("ColumnName")="Message"
'										dicInclusionRulesInfo("Value")="Air"
'										bReturn=Fn_SISW_PC_InclusionRulesViewTableOperations("VerifyCellExist",dicInclusionRulesInfo)

'										dicInclusionRulesInfo("Severity")="Error~Error~Information"
'										dicInclusionRulesInfo("Instance")="1~2~1"
'										bReturn=Fn_SISW_PC_InclusionRulesViewTableOperations("SelectRow",dicInclusionRulesInfo)
'
'										dicInclusionRulesInfo("Group")="Group1"
'										dicInclusionRulesInfo("Family")="Family1"
'										dicInclusionRulesInfo("Value")="Val1"
'										dicInclusionRulesInfo("Info")="ON"
'										dicInclusionRulesInfo("Warning")="OFF"
'										dicInclusionRulesInfo("Error")="ON"
'										bReturn= Fn_SISW_PC_InclusionRulesViewTableOperations("Search",dicInclusionRulesInfo)
'
'										bReturn= Fn_SISW_PC_InclusionRulesViewTableOperations("GetAllColumnName","")
'
'										dicInclusionRulesInfo("Value")="000210/A;1-TestItem"
'										bReturn= Fn_SISW_PC_InclusionRulesViewTableOperations("VerifyObjectName",dicInclusionRulesInfo)
'
'
'										dicInclusionRulesInfo("Severity")="Error~Error"
'										dicInclusionRulesInfo("Instance")="1~2"
'										bReturn=Fn_SISW_PC_InclusionRulesViewTableOperations("GetID",dicInclusionRulesInfo)
'
'History					 :			
'				Developer Name						Date					Rev. No.						Changes Done																				Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'				Ganesh B							23-May-2014													creat new Function
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_SISW_PC_InclusionRulesViewTableOperations(StrAction,dicInclusionRulesInfo)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_PC_InclusionRulesViewTableOperations"
	Dim ObjDefaultTable,arrRows, arrInstance, DeviceReplay
	Dim iTempInstance,iCounter,arrColName,arrValue,iCount,bFlag,StrColumnName,bFlag1
	Dim DictKeys,DictItems
	Dim objDeviceReplay
	Dim sID

	Fn_SISW_PC_InclusionRulesViewTableOperations=False
 	'Creating object of [ Variant Default View ] Table
	Set ObjDefaultTable=JavaWindow("ProductConfigurator").JavaTable("InclusionRulesViewTable")

	Select Case StrAction
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "SetCellData"
            arrColName=Split(dicInclusionRulesInfo("ColumnName"),"~")
			arrValue=Split(dicInclusionRulesInfo("Value"),"~")

			If dicInclusionRulesInfo("Instance")="" then
				dicInclusionRulesInfo("Instance")=1
			End if
			iTempInstance=1
            For iCount=0 to ubound(arrColName)
				iTempInstance=1
				bFlag=False
				For iCounter=0 to ObjDefaultTable.GetROProperty("rows")-1
					If Instr(1,ObjDefaultTable.GetCellData(iCounter,"Severity"),dicInclusionRulesInfo("Severity")) then
						If Cint(dicInclusionRulesInfo("Instance"))=iTempInstance Then
							ObjDefaultTable.SelectCell iCounter,dicInclusionRulesInfo("ColumnName")
							wait 2
							If JavaWindow("ProductConfigurator").JavaList("VariantDefaultTableList").Exist(2) Then

                                Set objDeviceReplay = CreateObject("Mercury.DeviceReplay")
								objDeviceReplay.SendString dicInclusionRulesInfo("Value")
								wait 1
								Set objDeviceReplay =Nothing
								
'								Call Fn_Button_Click("Fn_SISW_PC_InclusionRulesViewTableOperations", JavaWindow("ProductConfigurator"), "VariantDefaultViewTable_ContentSearch")
								wait 1

								bFlag=Fn_List_Select("Fn_SISW_PC_InclusionRulesViewTableOperations", JavaWindow("ProductConfigurator"), "VariantDefaultTableList",dicInclusionRulesInfo("Value"))
								wait 2
								Call Fn_KeyBoardOperation("SendKeys", "{ENTER}")
								Exit for
							Elseif JavaWindow("ProductConfigurator").JavaEdit("VariantDefaultTableEdit").Exist(2) then
								bFlag=Fn_Edit_Box("Fn_SISW_PC_InclusionRulesViewTableOperations",JavaWindow("ProductConfigurator"),"VariantDefaultTableEdit", dicInclusionRulesInfo("Value"))
								wait 2
								JavaWindow("ProductConfigurator").JavaEdit("VariantDefaultTableEdit").Activate
								Exit for
							End If
						End If
						iTempInstance=iTempInstance+1
					End if
				Next
                If bFlag=False Then
					Exit for
				End If
			Next
            If bFlag=True Then
				Fn_SISW_PC_InclusionRulesViewTableOperations=True
			End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "SetSeverity"
			For iCounter=0 to ObjDefaultTable.GetROProperty("rows")-1
				If ObjDefaultTable.GetCellData(iCounter,"Severity")="*" Then
'                    ObjDefaultTable.Object.showItem iCounter
					wait 1
					ObjDefaultTable.SelectCell iCounter,"Severity"
					wait 1
					If JavaWindow("ProductConfigurator").JavaList("VariantDefaultTableList").Exist(3) Then
						Set objDeviceReplay = CreateObject("Mercury.DeviceReplay")
						wait 1
						Set objDeviceReplay =Nothing
						
'						Call Fn_Button_Click("Fn_SISW_PC_InclusionRulesViewTableOperations", JavaWindow("ProductConfigurator"), "InclusionRulesViewTable_ContentSearch")
						wait 1
						For iCount=0 to JavaWindow("ProductConfigurator").JavaList("VariantDefaultTableList").GetROProperty("items count")-1
							If instr(1,JavaWindow("ProductConfigurator").JavaList("VariantDefaultTableList").Object.getItem(iCount),dicInclusionRulesInfo("Severity")) Then
								JavaWindow("ProductConfigurator").JavaList("VariantDefaultTableList").Select "#" & iCount
								wait 1
								Fn_SISW_PC_InclusionRulesViewTableOperations=True
								Exit For
							End If
						Next

						'Fn_SISW_PC_InclusionRulesViewTableOperations=Fn_List_Select("Fn_SISW_PC_InclusionRulesViewTableOperations", JavaWindow("ProductConfigurator"), "VariantDefaultTableList",dicInclusionRulesInfo("Severity"))
						Call Fn_KeyBoardOperation("SendKeys", "{ENTER}")
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS : Successfully add Severity  [ "+dicInclusionRulesInfo("Severity")+" ]")
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Severity list not exist")
					End If
				End If
			Next
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "GetID"
            arrValue=Split(dicInclusionRulesInfo("Severity"),"~")
			arrInstance=Split(dicInclusionRulesInfo("Instance"),"~")
			sID=""

			If dicInclusionRulesInfo("Instance")="" then
				dicInclusionRulesInfo("Instance")=1
			End if

			For iCount=0 to ubound(arrValue)
					iTempInstance=1
					bFlag=False
					For iCounter=0 to ObjDefaultTable.GetROProperty("rows")-1
						'verify [ Default Type ] match 
						If InStr(1,ObjDefaultTable.GetCellData(iCounter,"Severity"),arrValue(iCount)) then
							If Cint(arrInstance(iCount))=iTempInstance Then
									If sID="" Then
										sID=ObjDefaultTable.GetCellData(iCounter,"Object ID")
									Else
										sID=sID & "~" & ObjDefaultTable.GetCellData(iCounter,"Object ID")
									End If
									bFlag=True
									Exit for
							End if
							iTempInstance=iTempInstance+1
						End if
					Next
					If bFlag=False Then
						Exit For
					End If
			Next
            If bFlag=True Then
				Fn_SISW_PC_InclusionRulesViewTableOperations=sID
			End If
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	Case "Search"
		DictKeys = dicInclusionRulesInfo.Keys
		DictItems = dicInclusionRulesInfo.Items
		For iCount=0 to dicInclusionRulesInfo.count-1
			Select Case DictKeys(iCount)
				Case "Group","Family","Value"
					Call Fn_Edit_Box("Fn_SISW_PC_InclusionRulesViewTableOperations", JavaWindow("ProductConfigurator"),DictKeys(iCount),DictItems(iCount))
				Case "Info","Warning","Error"
					call Fn_SISW_UI_JavaCheckBox_Operations("Fn_SISW_PC_InclusionRulesViewTableOperations", "Set", JavaWindow("ProductConfigurator"), DictKeys(iCount), DictItems(iCount))
			End Select
	
		Next
		Fn_SISW_PC_InclusionRulesViewTableOperations=Fn_Button_Click("Fn_SISW_PC_InclusionRulesViewTableOperations", JavaWindow("ProductConfigurator"), "Search")
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "VerifyCellExist"
            arrColName=Split(dicInclusionRulesInfo("ColumnName"),"~")
			arrValue=Split(dicInclusionRulesInfo("Value"),"~")

			If dicInclusionRulesInfo("Instance")="" then
				dicInclusionRulesInfo("Instance")=1
			End if
			For iCount=0 to ubound(arrColName)
				iTempInstance=1
				bFlag=False
				For iCounter=0 to ObjDefaultTable.GetROProperty("rows")-1
					'verify [ Default Type ] match 
					If ObjDefaultTable.GetCellData(iCounter,"Severity")=dicInclusionRulesInfo("Severity") then
                        If IsNumeric(arrValue(iCount)) Then
							arrValue(iCount)=Cint(arrValue(iCount))
						End If
						If Cint(dicInclusionRulesInfo("Instance"))=iTempInstance Then
							If ObjDefaultTable.GetCellData(iCounter,arrColName(iCount))=arrValue(iCount) then
								bFlag=True
								Exit for
							Else
								Exit for
							End if
						End if
						iTempInstance=iTempInstance+1
					End if
				Next
                If bFlag=False Then
					Exit for
				End If
			Next
            If bFlag=True Then
				Fn_SISW_PC_InclusionRulesViewTableOperations=True
			End If
	End Select
	Set ObjDefaultTable=Nothing
End Function


