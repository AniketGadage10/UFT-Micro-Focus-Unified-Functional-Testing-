Option Explicit
'*********************************************************	Function List		***********************************************************************
'0.  Fn_SISW_ClassAdmin_GetObject
'1.  Fn_ClassAdmin_ToolbarOperations()
'2.  Fn_ClassAdmin_TreeNodeOperation()
'3.  Fn_ClassAdmin_ClassOperations()
'4.  Fn_CJavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("AddNewClass")()
'5.  Fn_ClassAdmin_TabOpeartions()
'6.  Fn_ClassAdmin_ExistingKeyLov_List()
'7.  Fn_ClassAdmin_FormatTypeSelect()
'8.  Fn_ClassAdmin_SearchAddAttributes()
'9.  Fn_ClassAdmin_CreateAttribute()
'10. Fn_ClassAdmin_ClassAttributesLists()
'11. Fn_ClassAdmin_DictionarySearchOperations()
'12. Fn_ClassAdmin_ICADictionaryTableOperations()
'13. Fn_ClassAdmin_KeyLOVTreeOperations()
'14. Fn_ClassAdmin_KeyLOVOperations()
'15. Fn_ClassAdmin_DateFormat()
'16. Fn_ClassAdmin_DeleteObjects()
'17. Fn_ClassAdmin_VerifyClassDetails()
'18. Fn_ClassAdmin_ClassAttributeOperations()
'19. Fn_ClassAdmin_ErrorHandler()
'20. Fn_ClassAdmin_AtrributeOperations()
'21. Fn_ClassAdmin_ViewOperations()
'22. Fn_ClassAdmin_ListofValues()
'23. Fn_ClassAdmin_UnitClassBasics()
'24. Fn_ClassAdmin_CreateVerifyMapping()
'25. Fn_ClassAdmin_StaticText()
'26. Fn_ClassAdmin_ClassSearchAndVerify 
'27. Fn_ClassAdmin_ConfigureReferencerAttribute()
'28. Fn_ClassAdmin_XMLExport()
'29 Fn_ClassAdmin_CheckBoxSet()
'30. Fn_ClassAdmin_Import()
'31. Fn_XML_File_Operations()
'32. Fn_ClassAdmin_ACLOperations()
'. *******************************************************************************************************************************************************
'****************************************    Function to get Object hierarchy ***************************************
'
''Function Name		 	:	Fn_SISW_ClassAdmin_GetObject
'
''Description		    :  	Function to get specified Object hierarchy.

''Parameters		    :	1. sObjectName : Object Handle name
								
''Return Value		    :  	Object \ Nothing
'
''Examples		     	:	Fn_SISW_ClassAdmin_GetObject("Remove")

'History:
'	Developer Name			Date			Rev. No.		Reviewer		Changes Done	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Koustubh Watwe		 7-June-2012		1.0
'-----------------------------------------------------------------------------------------------------------------------------------
'	Ashwini Kumar		 25-Sept-2013		2.0								Externalized object hierarchies
'-----------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_ClassAdmin_GetObject(sObjectName)
	Dim sObjectXMLPath
	sObjectXMLPath = Fn_GetEnvValue("User", "AutomationDir") & "\TestData\AutomationXML\ObjectXML\ClassAdmin.xml"
	Set Fn_SISW_ClassAdmin_GetObject = Fn_SISW_Setup_GetObjectFromXML(sObjectXMLPath, sObjectName)
End Function
''--------------------------------------------------------------------------------------------------------------------
'Function Name		:					Fn_ClassAdmin_ToolbarOperations

'Description			 :		 		    This function is used click on various buttons listed on Toolbar of Classification Admin.
'                                                                   
'Parameters			   :	 			   1. sOpeartion: Opearation to be performed
'                                                   										   
'Return Value		   : 			 		True/False

'Examples				:					Call Fn_ClassAdmin_ToolbarOperations("Cancel")
    
'' Function History
''--------------------------------------------------------------------------------------------------------------------
''	Developer Name		|	  Date		    |	Rev. No.	|		    Changes Done			|	Reviewer	|	Reviewed Date
''--------------------------------------------------------------------------------------------------------------------
'' 	Prasanna      		    |  07-Dec-2010	| 	1.0	    	   |									         |				       |  
''--------------------------------------------------------------------------------------------------------------------

Public Function Fn_ClassAdmin_ToolbarOperations(sOperation)
	GBL_FAILED_FUNCTION_NAME="Fn_ClassAdmin_ToolbarOperations"
		On Error Resume Next

		Dim bReturn,ObjClassAdminApplet

		Set ObjClassAdminApplet  =  Fn_SISW_ClassAdmin_GetObject("ClassAdminApplet") 
		
		Select Case sOperation
		
			Case "Save"
					bReturn = Fn_ToolbatButtonClick("Save current Instance")	
						if Fn_SISW_UI_Object_Operations("Fn_ClassAdmin_CreateAttribute","Exist", ObjClassAdminApplet.JavaDialog("Potential problem found"),"") then
						'If JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("Potential problem found").Exist(3) then
								JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("Potential problem found").JavaButton("OK").SetTOProperty "attached text","Yes"
								JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("Potential problem found").JavaButton("OK").Click micLeftBtn
						End if	
									
			Case "Cancel"
					bReturn = Fn_ToolbatButtonClick("Cancel edit")
			Case "Delete"
					bReturn = Fn_ToolbatButtonClick("Delete current instance")
					'Added by Vallari for Change in tooltip text od Delete in Tc9_0316 - 28-Mar-2011
					If cBool(bReturn) = False Then
						bReturn = Fn_ToolbatButtonClick("Delete")
					End If
			Case "Refresh"
					bReturn = Fn_ToolbatButtonClick("refresh privileges")
			Case "Edit"
					bReturn = Fn_ToolbatButtonClick("Edit current Instance")
			Case "Abort"
					bReturn = Fn_ToolbatButtonClick("Soft Abort")
			Case "NewInstance"
					bReturn = Fn_ToolbatButtonClick("Create a new Instance")
			Case "NewICO"
					bReturn = Fn_ToolbatButtonClick("Add or create a new Instance")				
			 Case "Favorite"
					bReturn = Fn_ToolbatButtonClick("show/hide Favorites")				
		End Select
		
		'wait(5)	
		If bReturn = false Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Perform ["+sOperation+"] Opeartion.")
				Fn_ClassAdmin_ToolbarOperations = false
				Exit Function 	
		Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully Performed ["+sOperation+"] Opeartion.")
				Fn_ClassAdmin_ToolbarOperations = true
		End If
End Function

'*********************************************************  Function do Operation on Classification Tree *********************************************************************

'Function Name		:					Fn_ClassAdmin_TreeNodeOperation

'Description			 :		 		    Action  performed :-
'																	1. Node Select
'																	2. Node Expand
'																	3. Node Collapse
'																	4. Exist
'																	5. DoubleClick
'																    

'Parameters			   :	 			1. StrAction: Action to be performed
'													2.StrNode: Fully qulified tree Path (delimiter as ':') [multiple node are separated by "," ] 
' 												   3. StrMenu: Context menu to be selected

'Return Value		   : 			 True/False

'Pre-requisite			:		 	 Classification  pane should be displayed.

'Examples				:			 Fn_ClassAdmin_TreeNodeOperation("Exist","SAM Classification Root","")
'											Fn_ClassAdmin_TreeNodeOperation("DoubleClick","SAM Classification Root","")
'

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Vidya Kulkarni				07-Dec-2010	       1.0														Prasanna
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function Fn_ClassAdmin_TreeNodeOperation(StrAction,StrNode,StrMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_ClassAdmin_TreeNodeOperation"
	 On Error Resume Next

	 'Declaration Of Variable
	   Dim objSearchTree,iItemCount,iCounter,sTreeItem,aMenuList,arrNodeList

	 'Initilisation Of Variable
	 Set objSearchTree =  Fn_SISW_ClassAdmin_GetObject("HierarchyTree")
   'Set objSearchTree = JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaTree("Hierarchy")
    if Fn_SISW_UI_Object_Operations("Fn_ClassAdmin_CreateAttribute","Exist", objSearchTree,"5") then
    'If  objSearchTree.Exist(5) Then
		'
	   iItemCount = objSearchTree.GetROProperty( "items count")

		arrNodeList = split(StrNode,":")
		 For iCounter=0 To (iItemCount-1)
						If Instr(1,arrNodeList(0),"SAM Classification Root") > 0 Then
									   'sTreeItem = JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaTree("Hierarchy").GetItem(iCounter)
									   sTreeItem = JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaTree("Hierarchy").GetItem(iCounter)
										 sTreeItem = Split(sTreeItem,":")

										 If UBOUND(sTreeItem) > 0 Then
														If  Instr(1,sTreeItem(UBOUND(sTreeItem)),arrNodeList(UBOUND(arrNodeList))) > 0 Then															
														  arrNodeList(UBOUND(arrNodeList)) =  sTreeItem(UBOUND(sTreeItem))
														  StrNode = Join(arrNodeList,":")
														End If
										 End If
						End If
		   Next

    	Err.Clear
      Select Case StrAction

       Case "Select"
					objSearchTree.Select StrNode					
					 If Err.Number < 0 Then
							Fn_ClassAdmin_TreeNodeOperation = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to select Node [ " + StrNode + "] of Classification Admin Tree." ) 
							Set objSearchTree = Nothing
							Exit Function 
					Else
							Fn_ClassAdmin_TreeNodeOperation = True
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully selected  Node [ " + StrNode + "] of Classification Admin Tree.") 
					End If

		Case "Expand"
				    objSearchTree.Expand StrNode
					If Err.Number < 0 Then
							Fn_ClassAdmin_TreeNodeOperation = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL :Failed to expand node   [" + StrNode + "] of Classification Admin Tree." )	
							Set objSearchTree = Nothing
							Exit Function 
					Else
							Fn_ClassAdmin_TreeNodeOperation = True
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully expanded node  [" + StrNode  + "] of Classification Admin Tree.")	
					End If

		Case "Collapse"
					objSearchTree.Collapse StrNode
					If Err.Number < 0 Then
							Fn_ClassAdmin_TreeNodeOperation = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Collapse node   [" + StrNode + "] of Classification Admin  Tree." )	
							Set objSearchTree = Nothing
							Exit Function 
					Else
							Fn_ClassAdmin_TreeNodeOperation = True
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Collapse node  [" + StrNode  + "] of Classification Admin Tree.")	
					End If

		Case "Exist"
					iItemCount = objSearchTree.GetROProperty( "items count")
					For iCounter=0 To (iItemCount-1)
							sTreeItem = objSearchTree.GetItem(iCounter)
							If Trim (Lcase(sTreeItem)) = Trim(Lcase(StrNode)) Then
									Fn_ClassAdmin_TreeNodeOperation = True
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully found node [" + StrNode + "] of Classification Admin Tree." )	
									Exit For
							End If
					Next 

					If  Cint(iCounter) = Cint (iItemCount) Then
							Fn_ClassAdmin_TreeNodeOperation = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to  found node [" + StrNode + "] of Classification Admin Tree." )	
							Set objSearchTree = Nothing
							Exit Function 
					End If

		Case  "DoubleClick"
				If Trim(StrNode) <> "" Then
							objSearchTree.Select StrNode
				End If
				objSearchTree.Activate StrNode

				If Err.Number < 0 Then
							Fn_ClassAdmin_TreeNodeOperation = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Double Clicked the Selected Node [" + StrNode + "] of Classification Admin Tree." )	
							Set objSearchTree = Nothing
							Exit Function 
				Else
							Fn_ClassAdmin_TreeNodeOperation = True
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Double Clicked on the Selected  Node [" + StrNode + " ] of Classification Admin Tree.")	
				End If

			Case "Deselect" 
					objSearchTree.Deselect StrNode
					'objSearchTree.Click 0,0
					 If Err.Number < 0 Then
							Fn_ClassAdmin_TreeNodeOperation = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Deselect Node [ " + StrNode + "] of Classification Admin Tree." ) 
							Set objSearchTree = Nothing
							Exit Function 
					Else
							Fn_ClassAdmin_TreeNodeOperation = True
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Deselected  Node [ " + StrNode + "] of Classification Admin Tree.") 
					End If

			   Case "RMB"
				
                JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaTree("Hierarchy").OpenContextMenu trim(StrNode)
				wait(2)	
				JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaMenu("Label:="+StrMenu).Select 
				If Err.Number < 0 Then
							Fn_ClassAdmin_TreeNodeOperation = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select the Menu  [" + StrMenu + "] Selected Node [" + StrNode + "] of Classification Admin Tree." )	
							Set objSearchTree = Nothing
							Exit Function 
				Else
							Fn_ClassAdmin_TreeNodeOperation = True
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected Menu  [" + StrMenu + "] on Node [" + StrNode + " ] of Classification Admin Tree.")	
				End if 

      End Select
   End If
   Set objSearchTree = Nothing
End Function 
 

'*********************************************************  Function performs Class Operations *********************************************************************
'Function Name  :   Fn_ClassAdmin_ClassOperations
'
'Description    :        Class Operations : Add, Edit, Remove
' 
'Parameters      :     sAction: Add
'           				 		dicClassOperations: Refer DictionaryDeclaration.vbs for the defination & keys included
' 
'Return Value     :   True/False
'
'Examples    :      
'								dicClassOperations.RemoveAll
'								dicClassOperations.Add "AssignID" , "ICM0102"
'								dicClassOperations.Add "ClassName" , "ABC"
'								dicClassOperations.Add "SysMeasurement" , "both(metric and non-metric)"
'								dicClassOperations.Add "Options_Abstract" , "OFF"
'								dicClassOperations.Add "Options_AllowsMultipleInstances" , "ON"
'								dicClassOperations.Add "Options_Assembly" , "ON"
'								dicClassOperations.Add "Options_PreventRemoteICOCreation" , "ON"
'								dicClassOperations.Add "SaveCurrentInstance" , "Save current Instance"
'								dicClassOperations.Add "Annotation" , "annotation text"
'								dicClassOperations.Add "ChkProperties_Check" , "Disable Auto Filter,User Defined Button,Propagated"
'								dicClassOperations.Add "ChkProperties_UnCheck" , "Mandatory,Unique,Protected"
'								
'           					Fn_ClassAdmin_ClassOperations(sAction,dicClassOperations)
' 
'History:
'          Developer Name		Date	Rev. No.     Changes Done                   														Reviewer 
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'           Pooja    		07-Dec-2010   1.0                                                                                               Prasanna 
'			Snehal S		14-Jan-2016	  1.1		Added new case "VerifyArraySize" from TC1015											[TC1122-2016010600-14_Jan_2016-VivekA-Maintenance]
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function  Fn_ClassAdmin_ClassOperations(sAction,dicClassOperations)
	GBL_FAILED_FUNCTION_NAME="Fn_ClassAdmin_ClassOperations"

	 On Error Resume Next
	 Dim dicCount , dicKeys , dicItems
	 Dim iCounter, bReturn ,iRanNum
	 Dim sNodeName,sChkProperties,sAnnotation,arrChkProperty,iOuterCount
	 Dim sAbstract,sAllowsMultipleInstances,sAssembly,sPreventRemoteICOCreation
	 Dim objSearchTree,sAssignIDText,sClassName,sSysOfMeasurement,sSaveCurrentInstance
	 Dim sAddImageUrl
	 Dim sAliasName,aAliasSet,iAliasCounter,aAliasSetVal,sLibrary,sDepAttribute
	 Dim objClassAdminApplet, bFlag, sClassAttributeswithValue, arrClassAttributewithValue, arrClassAttribute
	 Dim objAddClassDia
	 
	 dicCount  = dicClassOperations.Count
	 dicItems = dicClassOperations.Items
	 dicKeys = dicClassOperations.Keys

	Set objClassAdminApplet =  Fn_SISW_ClassAdmin_GetObject("ClassAdminApplet")
	Set objAddClassDia =  Fn_SISW_ClassAdmin_GetObject("AddClassDia")
	Err.Clear
	 Select Case sAction

	 Case "Add"
				'Select the node from the tree under which class has to be added
				sNodeName=dicClassOperations.Item("NodeName")
				bReturn=Fn_ClassAdmin_TreeNodeOperation("Select",sNodeName,"")
				 If bReturn = false Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Select ["+sNodeName+"] Node.")
						Fn_ClassAdmin_ClassOperations = false
						Exit Function 	
				Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully Selected ["+sNodeName+"] Node.")
						Fn_ClassAdmin_ClassOperations = true
				End If
				
				'Click on AddClass Button
				'objClassAdminApplet.JavaButton("Add Class").Click micLeftBtn
				 if Fn_SISW_UI_JavaButton_Operations("Fn_ClassAdmin_ClassOperations", "Click", objClassAdminApplet, "Add Class") = false then
				'If Err.Number < 0 Then
							Fn_ClassAdmin_ClassOperations = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed toClick on AddClass button" )
							Exit Function 
				Else
							Fn_ClassAdmin_ClassOperations = True
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Clicked on AddClass button")
				End If   

				Call Fn_ReadyStatusSync(3)
					
				'Check the existance of  'Add new Class' Dialog
				if Fn_SISW_UI_Object_Operations("Fn_ClassAdmin_CreateAttribute","Exist", objAddClassDia,"5") then
				'If objClassAdminApplet.JavaDialog("AddClassDia").Exist(5) = True Then
						 objAddClassDia.Activate
								 If Err.Number < 0 Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Add new Class Dialog Does not exist.")
										Fn_ClassAdmin_ClassOperations = False
										Exit Function
								Else
										wait(3)
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully verified the existence of Add new Class Dialog.")
										Fn_ClassAdmin_ClassOperations = True							
								End If
				End If

			   'If Assign Id is not provided then Click on Assign ID button
				sAssignIDText=dicClassOperations.Item("AssignID")
				If  sAssignIDText <> "" Then
						 objAddClassDia.JavaEdit("Class ID").Set sAssignIDText
						  If Err.Number < 0 Then
									Fn_ClassAdmin_ClassOperations = False
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to SET text in Class ID Edit Box" )								
									Exit Function 
							Else
									Fn_ClassAdmin_ClassOperations = True
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully SET text in Class ID Edit Box")								
							End If
				  Else
							iRanNum = Fn_Setup_RandNoGenerate(5)
							objAddClassDia.JavaEdit("Class ID").Set ""
							objAddClassDia.JavaEdit("Class ID").type "ICM" + iRanNum
							Call Fn_ReadyStatusSync(2)
							If Err.Number < 0 Then
									Fn_ClassAdmin_ClassOperations = False
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Click on Assign button" )
									Exit Function 
							Else
									Fn_ClassAdmin_ClassOperations = True
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Clicked on Assign button")
							End If    
				End If

				'Click on OK button
				'objAddClassDia.JavaButton("OK").Click micLeftBtn
				if Fn_SISW_UI_JavaButton_Operations("Fn_ClassAdmin_CreateAttribute", "Click", objAddClassDia,"OK") = false then
				'Wait 2
				'If Err.Number < 0 Then
							Fn_ClassAdmin_ClassOperations = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed toClick on OK button" )
							Exit Function 
				Else
							Fn_ClassAdmin_ClassOperations = True
							'Fetch the Class ID from UI : Parent Text box
							sAssignIDText=objClassAdminApplet.JavaEdit("Parent").GetROProperty ("value")
							sAddCase_AssignID=sAssignIDText
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Clicked on OK button")
				End If  

				'*******************************
				'Set Class Details Tab
				'*******************************
				bReturn = Fn_ClassAdmin_TabOpeartions("Activate","Subtab","Class Details","")
				If bReturn = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile")," Fail  | Failed to Activate [ Class Details ] Tab")
					Fn_ClassAdmin_ClassOperations = False
					Exit Function 
				Else
					Call Fn_ReadyStatusSync(5)
					Call Fn_WriteLogFile(Environment.Value("TestLogFile")," Pass | Successfully Activated [ Class Details ] Tab")
				End If

				'--------------------------------------------------------------------------------Delete Case--------------------------------------------------------------------------------------
                			Case "Delete"
						'Select the class node which  has to be deleted
						sNodeName=dicClassOperations.Item("NodeName")
						bReturn=Fn_ClassAdmin_TreeNodeOperation("Select",sNodeName,"")
						If bReturn = false Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Select ["+sNodeName+"] Node.")
								Fn_ClassAdmin_ClassOperations = false
								Exit Function 	
						Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully Selected ["+sNodeName+"] Node.")
								Fn_ClassAdmin_ClassOperations = true
						End If
						'Set  Edit Mode
						bReturn =  Fn_ClassAdmin_ToolbarOperations("Edit")
						If bReturn = false Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Perform [Edit] Opeartion.")
									Fn_ClassAdmin_ClassOperations = false
									Exit Function 	
						Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully Performed [Edit] Opeartion.")
									Fn_ClassAdmin_ClassOperations = true
						End If
						'Delete the selected Class
						bReturn =  Fn_ClassAdmin_ToolbarOperations("Delete")
						If bReturn = false Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Perform [Delete] Opeartion.")
									Fn_ClassAdmin_ClassOperations = false
									Exit Function 	
						Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully Performed [Delete] Opeartion.")
									Fn_ClassAdmin_ClassOperations = true
						End If
						
						if Fn_SISW_UI_Object_Operations("Fn_ClassAdmin_CreateAttribute","Exist", JavaDialog("ConfirmationMessage"),"5") then
						'If  JavaDialog("ConfirmationMessage").Exist(5) Then
								Set oConDlg = JavaDialog("ConfirmationMessage")
						'ElseIf objClassAdminApplet.JavaDialog("DeleteConfirmation").Exist(5) Then
						ElseIf Fn_SISW_UI_Object_Operations("Fn_ClassAdmin_CreateAttribute","Exist", objClassAdminApplet.JavaDialog("DeleteConfirmation"),"5") Then
							Set oConDlg = objClassAdminApplet.JavaDialog("DeleteConfirmation")
                        Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Confirmation Dialog Does not exist.")												
							Fn_ClassAdmin_ClassOperations = False
							Set oConDlg = Nothing
							Exit Function
						End If

						If oConDlg.Exist(5) = True Then
							oConDlg.Activate
							 If Err.Number < 0 Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Confirmation Dialog Does not exist.")												
									Fn_ClassAdmin_ClassOperations = False
									Set oConDlg = Nothing
									Exit Function
							Else
									wait(3)
									'Vallari - Ready is not Ready as Delete COnfirmation dialog in ON
									'Call Fn_ReadyStatusSync(3)
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully verified the existence of Confirmation Dialog.")								
							End If
	
							
							'Click on 'Yes' button
							oConDlg.JavaButton("Yes").Click micLeftBtn
							If Err.Number < 0 Then
										Fn_ClassAdmin_ClassOperations = False
										oConDlg.JavaButton("No").Click micLeftBtn
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed toClick on Yes button" )
										Set oConDlg = Nothing
										Exit Function 
							Else							
										wait(3)
										Call Fn_ReadyStatusSync(5)
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Clicked on Yes button")
							End If	
			End If
			Set oConDlg = Nothing
'			'--------------------------------------------------------------------------------End of Delete ------------------------------------------------------------------------------------------

				
				Case "RemoveImage"																									' Added on 13th Dec 2010
							'Select the Class node from the tree
							sNodeName=dicClassOperations.Item("NodeName")
							bReturn=Fn_ClassAdmin_TreeNodeOperation("Select",sNodeName,"")
							If bReturn = false Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Select ["+sNodeName+"] Node.")
									Fn_ClassAdmin_ClassOperations = false
									Exit Function 	
							Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully Selected ["+sNodeName+"] Node.")
									Fn_ClassAdmin_ClassOperations = true
							End If
							'Set  Edit Mode
							bReturn =  Fn_ClassAdmin_ToolbarOperations("Edit")
							If bReturn = false Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Perform [Edit] Opeartion.")
										Fn_ClassAdmin_ClassOperations = false
										Exit Function 	
							Else
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully Performed [Edit] Opeartion.")
										Fn_ClassAdmin_ClassOperations = true
							End If
							'Click on Delete Image Button
							objClassAdminApplet.JavaButton("DeleteImage").Click micLeftBtn
							 If Err.Number < 0 Then
										Fn_ClassAdmin_ClassOperations = False
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Click on [Delete Image] Button" )								
										Exit Function 
							Else
										Call Fn_ReadyStatusSync(3)
										wait(3)
										Fn_ClassAdmin_ClassOperations = True
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Clicked on [Delete Image] Button")							
							End If

				'---------------------------------------------------------------------------------Modify Case-------------------------------------------------------------------------------------------

					Case "Modify"
							'Select the Class node from the tree that has to be Modified
							sNodeName=dicClassOperations.Item("NodeName")
							bReturn=Fn_ClassAdmin_TreeNodeOperation("Select",sNodeName,"")
							If bReturn = false Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Select ["+sNodeName+"] Node.")
									Fn_ClassAdmin_ClassOperations = false
									Exit Function 	
							Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully Selected ["+sNodeName+"] Node.")
									Fn_ClassAdmin_ClassOperations = true
							End If
							'Set  Edit Mode
							bReturn =  Fn_ClassAdmin_ToolbarOperations("Edit")
							If bReturn = false Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Perform [Edit] Opeartion.")
										Fn_ClassAdmin_ClassOperations = false
										Exit Function 	
							Else
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully Performed [Edit] Opeartion.")
										Fn_ClassAdmin_ClassOperations = true
							End If
					'[TC1122-2016010600-14_Jan_2016-VivekA-Maintenance] - Added new case "VerifyArraySize" from tc 1015 to verify the array size of attributes - By Snehal S
					Case "VerifyArraySize"
							Set objClassAdminApplet       = objClassAdminApplet
							bFlag = False
							sNodeName 					  = dicClassOperations("NodeName")
							sClassAttributeswithValue     = dicClassOperations("Array")
							arrClassAttributewithValue    = split(sClassAttributeswithValue,"~")
							
							bReturn=Fn_ClassAdmin_TreeNodeOperation("Select",sNodeName,"")
							If bReturn = False Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Select ["+sNodeName+"] Node.")
								Fn_ClassAdmin_ClassOperations = False
								Set objClassAdminApplet = nothing
								Exit Function 	
							Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully Selected ["+sNodeName+"] Node.")
								Fn_ClassAdmin_ClassOperations = True
							End If
							
							Call  Fn_ClassAdmin_TabOpeartions("Activate","Subtab","Class Attributes","")
							
							For iCounter = 0 To UBound(arrClassAttributewithValue)
									bFlag = False
									arrClassAttribute = split(arrClassAttributewithValue(iCounter),":")
									bReturn = Fn_ClassAdmin_ClassAttributesLists("Select","",Array(arrClassAttribute(0)))
									If bReturn = false Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Select Class Attribute ["+arrClassAttribute(0)+"] .")
										Fn_ClassAdmin_ClassOperations = False
										Set objClassAdminApplet = nothing
										Exit Function 	
									Else
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully Selected Class Attribute ["+arrClassAttribute(0)+"] .")
										Fn_ClassAdmin_ClassOperations = True
									End If		
							
									If CInt(objClassAdminApplet.JavaEdit("ArrayLength").GetROProperty("text")) =  CInt(arrClassAttribute(1)) Then
										bFlag = True
									Else
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Verify the size [ " + arrClassAttribute(1) + "] of Class Attribute [ " & arrClassAttribute(0) & "] .")
										Fn_ClassAdmin_ClassOperations = False
										Set objClassAdminApplet = nothing
										Exit Function 					
									End If				
							Next
							
							If bFlag = True Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully Verified  the size of Class Attributes ." )
								Fn_ClassAdmin_ClassOperations = True
								Set objClassAdminApplet = nothing
								Exit Function 		
							End If
				'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

	 End Select

	' Dictionary  to handle - Name of Class, Sys of measurement, Options
	  For iCounter = 0 to dicCount - 1
          If  dicItems(iCounter) <> "" Then
			   Select Case dicKeys(iCounter)

			    Case "ClassName"
							sClassName=dicItems(iCounter)
							If  sClassName <> "" Then
									call Fn_SISW_UI_JavaEdit_Operations("Fn_ClassAdmin_ClassOperations", "Set",  objClassAdminApplet, "ClassName", sClassName)
									 If Err.Number < 0 Then
											Fn_ClassAdmin_ClassOperations = False
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to SET ["+sClassName+"] in ClassName Edit Box" )								
											Exit Function 
									Else
											Fn_ClassAdmin_ClassOperations = True
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully SET ["+sClassName+"]  inClassName Edit Box")								
									End If
							End If


				Case "SysMeasurement"
							sSysOfMeasurement=dicItems(iCounter)
							If  sSysOfMeasurement <> "" Then
									 objClassAdminApplet.JavaRadioButton("Measurement").SetTOProperty "Attached Text",lcase(sSysOfMeasurement)
									 objClassAdminApplet.JavaRadioButton("Measurement").Set "ON"
		
									 If Err.Number < 0 Then
											Fn_ClassAdmin_ClassOperations = False									
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to SET Radio Button ["+sSysOfMeasurement+"]" )								
											Exit Function 
									Else
											Fn_ClassAdmin_ClassOperations = True
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully SET Radio Button ["+sSysOfMeasurement+"]")								
									End If
							End If

				Case "Options_Abstract"
						sAbstract=dicItems(iCounter)
						If  sAbstract <> "" Then					      
								 objClassAdminApplet.JavaCheckBox("OptionsCheckBox").SetTOProperty "Attached Text","Abstract"
								 objClassAdminApplet.JavaCheckBox("OptionsCheckBox").Set sAbstract
								 If Err.Number < 0 Then
										Fn_ClassAdmin_ClassOperations = False
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to SET Options Abstract to ["+sAbstract+"]" )								
										Exit Function 
								Else
										Fn_ClassAdmin_ClassOperations = True
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully SET Options Abstract to ["+sAbstract+"]")								
								End If
						End If

				Case "Options_AllowsMultipleInstances"
						sAllowsMultipleInstances=dicItems(iCounter)
						If  sAllowsMultipleInstances <> "" Then						      
								 objClassAdminApplet.JavaCheckBox("OptionsCheckBox").SetTOProperty "Attached Text","Allows multiple Instances"
								 objClassAdminApplet.JavaCheckBox("OptionsCheckBox").Set sAllowsMultipleInstances
								 If Err.Number < 0 Then
										Fn_ClassAdmin_ClassOperations = False
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to SET Allows Multiple Instances CheckBox to ["+sAllowsMultipleInstances+"]" )								
										Exit Function 
								Else
										Fn_ClassAdmin_ClassOperations = True
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully SET Allows Multiple Instances CheckBox to ["+sAllowsMultipleInstances+"]")								
								End If
						End If

				Case "Options_Assembly"
						sAssembly=dicItems(iCounter)
						If  sAssembly <> "" Then						      
								 objClassAdminApplet.JavaCheckBox("OptionsCheckBox").SetTOProperty "Attached Text","Assembly"
								 objClassAdminApplet.JavaCheckBox("OptionsCheckBox").Set sAssembly
								 If Err.Number < 0 Then
										Fn_ClassAdmin_ClassOperations = False								
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to SET Assembly CheckBox to ["+sAssembly+"]" )								
										Exit Function 
								Else
										Fn_ClassAdmin_ClassOperations = True									
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully SET Assembly CheckBox to ["+sAssembly+"]")								
								End If
						End If


				Case "Options_PreventRemoteICOCreation"
						sPreventRemoteICOCreation=dicItems(iCounter)
						If  sPreventRemoteICOCreation <> "" Then						      
								 objClassAdminApplet.JavaCheckBox("OptionsCheckBox").SetTOProperty "Attached Text","Prevent remote ICO creation"
								 objClassAdminApplet.JavaCheckBox("OptionsCheckBox").Set sPreventRemoteICOCreation
								 If Err.Number < 0 Then
										Fn_ClassAdmin_ClassOperations = False								
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to SET Prevent Remote ICO Creation CheckBox ["+sPreventRemoteICOCreation+"]" )								
										Exit Function 
								Else
										Fn_ClassAdmin_ClassOperations = True								
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully SET Prevent Remote ICO Creation CheckBox ["+sPreventRemoteICOCreation+"]")												
								End If
						End If

				  Case "SaveCurrentInstance"
						sSaveCurrentInstance=dicItems(iCounter)
						If  sSaveCurrentInstance <> "" Then
								bReturn =  Fn_ClassAdmin_ToolbarOperations("Save")
								 If bReturn = false Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Perform ["+sSaveCurrentInstance+"] Opeartion.")
										Fn_ClassAdmin_ClassOperations = false
										Exit Function 	
								Else
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully Performed ["+sSaveCurrentInstance+"] Opeartion.")
										Fn_ClassAdmin_ClassOperations = true
								End If
						End If

				  Case "Annotation"
						sAnnotation=dicItems(iCounter)
						If  sAnnotation <> "" Then
								 objClassAdminApplet.JavaEdit("Annotation").Set sAnnotation
								 If Err.Number < 0 Then
										Fn_ClassAdmin_ClassOperations = False
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to SET text in sAnnotation Edit Box" )								
										Exit Function 
								Else
										Fn_ClassAdmin_ClassOperations = True
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully SET text in sAnnotation Edit Box")								
								End If
						End If

				Case "ChkProperties_Check"
						sChkProperties=dicItems(iCounter)
						arrChkProperty = Split(sChkProperties,",")
						For iOuterCount = 0 to  Ubound(arrChkProperty)
								If  arrChkProperty(iOuterCount) <> "" Then
									'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
										If arrChkProperty(iOuterCount) ="Application 1" Then
												Call Fn_ClassAdmin_TabOpeartions("Activate","ImageTab","Applicability/User Data","")
												wait 1, 500
												objClassAdminApplet.JavaCheckBox("ChkProperties").SetTOProperty "Attached Text","<html> Marks the attribute as important for Application 1 </html>" 
												objClassAdminApplet.JavaCheckBox("ChkProperties").Set "ON"
												 If Err.Number < 0 Then
													 	Fn_ClassAdmin_ClassOperations = False
														Exit Function
												Else
														Fn_ClassAdmin_ClassOperations = True
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully SET Properties CheckBox ["+sChkProperties+"]")		
												End If

										ElseIf arrChkProperty(iOuterCount) ="Application 2" Then
												Call Fn_ClassAdmin_TabOpeartions("Activate","ImageTab","Applicability/User Data","")
												wait 1, 500
												objClassAdminApplet.JavaCheckBox("ChkProperties").SetTOProperty "Attached Text","<html> Marks the attribute as important for Application 2 </html>" 
												objClassAdminApplet.JavaCheckBox("ChkProperties").Set "ON"
												 If Err.Number < 0 Then
													 	Fn_ClassAdmin_ClassOperations = False
														Exit Function
												Else
														Fn_ClassAdmin_ClassOperations = True
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully SET Properties CheckBox ["+sChkProperties+"]")		
												End If
	
										ElseIf arrChkProperty(iOuterCount) ="Application 3" Then
												Call Fn_ClassAdmin_TabOpeartions("Activate","ImageTab","Applicability/User Data","")
												wait 1, 500
												objClassAdminApplet.JavaCheckBox("ChkProperties").SetTOProperty "Attached Text","<html> Marks the attribute as important for Application 3 </html>" 
												objClassAdminApplet.JavaCheckBox("ChkProperties").Set "ON"
												 If Err.Number < 0 Then
													 	Fn_ClassAdmin_ClassOperations = False
														Exit Function
												Else
														Fn_ClassAdmin_ClassOperations = True
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully SET Properties CheckBox ["+sChkProperties+"]")		
												End If
										ElseIf arrChkProperty(iOuterCount) ="Application 4" Then
												Call Fn_ClassAdmin_TabOpeartions("Activate","ImageTab","Applicability/User Data","")
												wait 1, 500
												objClassAdminApplet.JavaCheckBox("ChkProperties").SetTOProperty "Attached Text","<html> Marks the attribute as important for Application 4 </html>" 
												objClassAdminApplet.JavaCheckBox("ChkProperties").Set "ON"
												 If Err.Number < 0 Then
													 	Fn_ClassAdmin_ClassOperations = False
														Exit Function
												Else
														Fn_ClassAdmin_ClassOperations = True
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully SET Properties CheckBox ["+sChkProperties+"]")		
												End If		
										ElseIf arrChkProperty(iOuterCount) ="Application 5" Then
												Call Fn_ClassAdmin_TabOpeartions("Activate","ImageTab","Applicability/User Data","")
												wait 1, 500
												objClassAdminApplet.JavaCheckBox("ChkProperties").SetTOProperty "Attached Text","<html> Marks the attribute as important for Application 5 </html>" 
												objClassAdminApplet.JavaCheckBox("ChkProperties").Set "ON"
												 If Err.Number < 0 Then
													 	Fn_ClassAdmin_ClassOperations = False
														Exit Function
												Else
														Fn_ClassAdmin_ClassOperations = True
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully SET Properties CheckBox ["+sChkProperties+"]")		
												End If		

										Else
												objClassAdminApplet.JavaCheckBox("ChkProperties").SetTOProperty "Attached Text",arrChkProperty(iOuterCount) 
												objClassAdminApplet.JavaCheckBox("ChkProperties").Set "ON"
												 If Err.Number < 0 Then
															Fn_ClassAdmin_ClassOperations = False
															Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to SET Properties CheckBox ["+sChkProperties+"]" )					
															Exit Function 
												Else
															Fn_ClassAdmin_ClassOperations = True
															Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully SET Properties CheckBox ["+sChkProperties+"]")			
												End If
									End If
								End If
								'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
						Next

				Case "ChkProperties_UnCheck"
						sChkProperties=dicItems(iCounter)
						arrChkProperty = Split(sChkProperties,",")
						For iOuterCount = 0 to  Ubound(arrChkProperty)
								If  arrChkProperty(iOuterCount) <> "" Then
										 objClassAdminApplet.JavaCheckBox("ChkProperties").SetTOProperty "Attached Text",arrChkProperty(iOuterCount) 
										objClassAdminApplet.JavaCheckBox("ChkProperties").Set "OFF"
										 If Err.Number < 0 Then
												Fn_ClassAdmin_ClassOperations = False
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Uncheck Properties CheckBox ["+sChkProperties+"]" )					
												Exit Function 
										Else
												Fn_ClassAdmin_ClassOperations = True
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Unchecked Properties CheckBox ["+sChkProperties+"]")			
										End If
								End If
						Next

				Case "Array"
						objClassAdminApplet.JavaCheckBox("ChkArray").Set "ON"
						If Err.Number < 0 Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Set Array Check Box")
								Fn_ClassAdmin_ClassOperations = false
								Exit Function 					
						 Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Successfully Set Array Check Box")
								Fn_ClassAdmin_ClassOperations = true
								wait(1)								
						End If

						objClassAdminApplet.JavaEdit("ArrayLength").Set cstr(dicItems(iCounter))
						If Err.Number < 0 Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Set Array Check Box")
								Fn_ClassAdmin_ClassOperations = false
								Exit Function 					
						 Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Successfully Set Array Check Box")
								Fn_ClassAdmin_ClassOperations = true
								wait(1)								
						End If


				Case "Image"
							'Select the Class node from the tree
							sNodeName=dicClassOperations.Item("NodeName")
							bReturn=Fn_ClassAdmin_TreeNodeOperation("Select",sNodeName,"")
							If bReturn = false Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Select ["+sNodeName+"] Node.")
									Fn_ClassAdmin_ClassOperations = false
									Exit Function 	
							Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully Selected ["+sNodeName+"] Node.")
									Fn_ClassAdmin_ClassOperations = true
							End If
							'Set  Edit Mode
							bReturn =  Fn_ClassAdmin_ToolbarOperations("Edit")
							If bReturn = false Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Perform [Edit] Opeartion.")
										Fn_ClassAdmin_ClassOperations = false
										Exit Function 	
							Else
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully Performed [Edit] Opeartion.")
										Fn_ClassAdmin_ClassOperations = true
							End If
							'Click on Add Image Button
							objClassAdminApplet.JavaButton("AddImage").Click micLeftBtn
							 If Err.Number < 0 Then
										Fn_ClassAdmin_ClassOperations = False
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Click on [Add Image] Button" )								
										Exit Function 
							Else
										Fn_ClassAdmin_ClassOperations = True
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Clicked on [Add Image] Button")							
							End If
							'Check existance of the Select Image Dialog Box
							Call Fn_ReadyStatusSync(3)
							If JavaDialog("SelectImageDialog").Exist(5) = True Then
										JavaDialog("SelectImageDialog").Activate
												 If Err.Number < 0 Then
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Select Image Dialog Does not exist.")
														Fn_ClassAdmin_ClassOperations = False
														Exit Function
												Else
														wait(3)
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully verified the existence of Select Image Dialog.")
														Fn_ClassAdmin_ClassOperations = True							
												End If
							End If
							'Paste the url in file name edit box
							sAddImageUrl=dicItems(iCounter)
							If  sAddImageUrl <> "" Then
											JavaDialog("SelectImageDialog").JavaEdit("FileName").Set sAddImageUrl
											 If Err.Number < 0 Then
													Fn_ClassAdmin_ClassOperations = False
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to SET url ["+sAddImageUrl+"] in File name Edit Box" )								
													Exit Function 
											Else
													Fn_ClassAdmin_ClassOperations = True
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully SET url ["+sAddImageUrl+"] in File name Edit Box")								
											End If
							End If
						   'Click on Import button
							'JavaDialog("SelectImageDialog").JavaButton("Import").Click micLeftBtn
							JavaDialog("SelectImageDialog").JavaButton("Add").Click micLeftBtn
								 If Err.Number < 0 Then
											Fn_ClassAdmin_ClassOperations = False
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Click on [ Import ] Button" )								
											Exit Function 
								Else
											Call Fn_ReadyStatusSync(3)
											wait(3)
											Fn_ClassAdmin_ClassOperations = True
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Clicked on [ Import ] Button")							
								End If

					   '===================================================================================================================================


		Case "AliasName"
				sAliasName=dicItems(iCounter)                   'Array("en:English,en:UKEnglish;de:Germen;de:UKGerman")
				If  sAliasName <> "" Then
						 If instr(1,sAliasName,",") > 0 Then
								 aAliasSet = split(sAliasName,",",-1,1)                                                                                                                                                     
						End If
				For iAliasCounter = 0 to UBOUND(aAliasSet) 

					aAliasSetVal = split(aAliasSet(iAliasCounter),":",-1,1)         
			 'Select value from Drop Down 
			 	objClassAdminApplet.JavaList("AliasNameList").SetTOProperty "Index",0
				Wait 1
				objClassAdminApplet.JavaList("AliasNameList").Select aAliasSetVal(0)
				 If Err.Number < 0 Then
						 Fn_ClassAdmin_ClassOperations = False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Select ["+aAliasSet(iAliasCounter)+"] from Alias Names List" )                                                                                                               
						Exit Function 
				Else
						Fn_ClassAdmin_ClassOperations = True
						 wait(1)
						 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully SET ["+aAliasSet(iAliasCounter)+"]  from Alias Names List")       
				End If  

				'Type in Edit Box
				objClassAdminApplet.JavaEdit("AliasName").Set trim(aAliasSetVal(1))                                                                                                                                                              
				 If Err.Number < 0 Then      
				     Fn_ClassAdmin_ClassOperations = False
					 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Set Value ["+aAliasSet(iAliasCounter)+"] " )     
					 Exit Function    
				Else
					     Fn_ClassAdmin_ClassOperations = True
						 wait(1)
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully SET ["+aAliasSet(iAliasCounter)+"]  ")       
				End If                                                                                                   
				'Click on Add Button
			   objClassAdminApplet.JavaButton("AddAlias").Click micLeftBtn
			   If Err.Number < 0 Then
						 Fn_ClassAdmin_ClassOperations = False
                         Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Add Value in Alias Names List" )
						  Exit Function 
				Else         
						   Fn_ClassAdmin_ClassOperations = True
						   wait(1)
						   Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Add Value in Alias Names List ")
				End If
				Next

				objClassAdminApplet.JavaEdit("ClassName").Set sClassName																												
				  If Err.Number < 0 Then
							Fn_ClassAdmin_ClassOperations = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to SET ["+sClassName+"] in ClassName Edit Box" )
							Exit Function
				 Else 
						  Fn_ClassAdmin_ClassOperations = True
						  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully SET ["+sClassName+"]  inClassName Edit Box")
						   End If
				 End If                

	'======================================Added By Pooja S   23-Jan-2012=======================================

		Case "Library"
				sLibrary=dicItems(iCounter)                  
				If  sLibrary <> "" Then
						 If instr(1,sLibrary,",") > 0 Then
								 aAliasSet = split(sLibrary,",",-1,1)                                                                                                                                                     
						End If
					For iAliasCounter = 0 to UBOUND(aAliasSet) 

						aAliasSetVal = split(aAliasSet(iAliasCounter),":",-1,1)         
						'Select value from Drop Down 
						objClassAdminApplet.JavaList("AliasNameList").SetTOProperty "Index",1
						wait(3)
						objClassAdminApplet.JavaList("AliasNameList").Select aAliasSetVal(0)
						If Err.Number < 0 Then
								 Fn_ClassAdmin_ClassOperations = False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Select ["+aAliasSet(iAliasCounter)+"] from Alias Names List" )                                                                                                               
								Exit Function 
						Else
								Fn_ClassAdmin_ClassOperations = True
								wait(1)
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully SET ["+aAliasSet(iAliasCounter)+"]  from Alias Names List")       
						End If  
						'Type in Edit Box
						objClassAdminApplet.JavaEdit("AliasName").SetTOProperty "Index",2
						wait(3)
						objClassAdminApplet.JavaEdit("AliasName").Set trim(aAliasSetVal(1))                                                                                                                                                              
						 If Err.Number < 0 Then      
								 Fn_ClassAdmin_ClassOperations = False
								 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Set Value ["+aAliasSet(iAliasCounter)+"] " )     
								 Exit Function    
						Else
								 Fn_ClassAdmin_ClassOperations = True
								 wait(1)
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully SET ["+aAliasSet(iAliasCounter)+"]  ")       
						End If    				                                                                                               
						'Click on Add Button
						objClassAdminApplet.JavaButton("AddAlias").SetTOProperty "Index",1
						wait(3)
						objClassAdminApplet.JavaButton("AddAlias").Click micLeftBtn
						 If Err.Number < 0 Then
								 Fn_ClassAdmin_ClassOperations = False
								 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Add Value in Alias Names List" )
								  Exit Function 
						Else         
								   Fn_ClassAdmin_ClassOperations = True
								   wait(1)
								   Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Add Value in Alias Names List ")
						End If
					Next
				Else
						objClassAdminApplet.JavaEdit("ClassName").Set sClassName																												
						 If Err.Number < 0 Then
								Fn_ClassAdmin_ClassOperations = False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to SET ["+sClassName+"] in ClassName Edit Box" )
								Exit Function
						Else 
							  Fn_ClassAdmin_ClassOperations = True
							  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully SET ["+sClassName+"]  inClassName Edit Box")
						End If
				 End If

		Case "Dependency Attribute"
				sDepAttribute = dicItems(iCounter)
				If sDepAttribute <> "" Then
					objClassAdminApplet.JavaList("DependencyAttribute").Select sDepAttribute
					If Err.Number < 0 Then
						Fn_ClassAdmin_ClassOperations = False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to SET Dependency Attribute to ["+sDepAttribute+"]" )
						Exit Function 
					Else
						Fn_ClassAdmin_ClassOperations = True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully SET Dependency Attribute to ["+sDepAttribute+"]")
					End If
				End If
		  '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
			End Select

		End If
	Next	

		'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		'Get the Class Assign ID as a Return value
		'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		If  sAction="Add" Then
				Fn_ClassAdmin_ClassOperations =sAddCase_AssignID
		Else
				Fn_ClassAdmin_ClassOperations = True
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Performed ["+sAction+"] Class Operation")
		End If

End Function


'*********************************************************  Function performs Group Operations *********************************************************************
'Function Name  :   Fn_ClassAdmin_GroupOperations
'
'Description    :        Group Operations: Add, Remove, Edit, Delete
' 
'Parameters      :     sAction: Add
'           				 		dicGroupOperations: Refer DictionaryDeclaration.vbs for the defination & keys included
' 
'Return Value     :   True/False
'
'Examples    :      
'								dicGroupOperations.RemoveAll
'								dicGroupOperations.Add "NodeName" , "SAM Classification Root:Classification Root"
'								dicGroupOperations.Add "AssignID" , "SAM09"
'								dicGroupOperations.Add "GroupName" , "Group123"
'								dicGroupOperations.Add "Check_ICOCreation" , "Prevent remote ICO creation"
'								
'           					Fn_ClassAdmin_GroupOperations(sAction,dicGroupOperations)
' 
'History:
'          Developer Name                          Date                                                       Rev. No.         Changes Done                   Reviewer 
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'           Pooja    		                                   07-Dec-2010   1.0                                                                                                              Prasanna 
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function  Fn_ClassAdmin_GroupOperations(sAction,dicGroupOperations)
	GBL_FAILED_FUNCTION_NAME="Fn_ClassAdmin_GroupOperations"
	 On Error Resume Next
	 Dim dicCount , dicKeys , dicItems
	 Dim iCounter, bReturn,sAddImageUrl
	 Dim sNodeName,sChkProperties,sAnnotation,arrChkProperty,iOuterCount
	 Dim sAssignIDText,sGroupName,sICOCreation,sSaveCurrentInstance,oConDlg
	  Dim objClassAdminApplet	
     dicCount  = dicGroupOperations.Count
	 dicItems = dicGroupOperations.Items
	 dicKeys = dicGroupOperations.Keys

	Set objClassAdminApplet =  Fn_SISW_ClassAdmin_GetObject("ClassAdminApplet")
	
	 Select Case sAction

	 Case "Add"
				'Select the node from the tree under which group has to be added
				sNodeName=dicGroupOperations.Item("NodeName")
'				bReturn=Fn_ClassAdmin_TreeNodeOperation("RMB",sNodeName,"Add Group")
'				If bReturn = false Then
'						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to RMB & Select ["+sNodeName+"] Node & invoke the New Group Dialog")
'						Fn_ClassAdmin_GroupOperations = false
'						Exit Function 	
'				Else
'						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully RMB & Selected ["+sNodeName+"] Node & invoked the New Group Dialog")
'						Fn_ClassAdmin_GroupOperations = true
'						wait(3)
'				End If

				'Select node
				bReturn=Fn_ClassAdmin_TreeNodeOperation("Select",sNodeName,"")
				wait(3)
'				'Click on Add Group Button
				objClassAdminApplet.JavaButton("Add Group").Click micLeftBtn
				If Err.Number < 0 Then
							Fn_ClassAdmin_GroupOperations = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed toClick on Add Group button" )
							Exit Function 
				Else
							Fn_ClassAdmin_GroupOperations = True
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Clicked on Add Group button")
				End If   

				'Check the existance of  'Add new Group' Dialog
				If objClassAdminApplet.JavaDialog("AddNewGroup").Exist(5) = True Then
                         								 If Err.Number < 0 Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Add new Group Dialog Does not exist.")
										Fn_ClassAdmin_GroupOperations = False
										Exit Function
								Else
										wait(3)
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully verified the existence of Add new Group.")
										Fn_ClassAdmin_GroupOperations = True							
								End If
				End If

			   'Click on Assign button
				sAssignIDText=dicGroupOperations.Item("AssignID")
				If  sAssignIDText <> "" Then
						 objClassAdminApplet.JavaDialog("AddNewGroup").JavaEdit("Group ID").Set sAssignIDText
						  If Err.Number < 0 Then
									Fn_ClassAdmin_GroupOperations = False
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to SET text in Group ID Edit Box" )								
									Exit Function 
							Else
									Fn_ClassAdmin_GroupOperations = True
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully SET text in Group ID Edit Box")								
							End If
				  Else
							objClassAdminApplet.JavaDialog("AddNewGroup").JavaButton("Assign").Click micLeftBtn
							Call Fn_ReadyStatusSync(2)
							If Err.Number < 0 Then
									Fn_ClassAdmin_GroupOperations = False
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed toClick on Assign button" )
									Exit Function 
							Else
									Fn_ClassAdmin_GroupOperations = True
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Clicked on Assign button")
							End If    
				End If

				'Click on OK button
				objClassAdminApplet.JavaDialog("AddNewGroup").JavaButton("OK").Click micLeftBtn
				If Err.Number < 0 Then
							Fn_ClassAdmin_GroupOperations = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed toClick on OK button" )
							Exit Function 
				Else
							Fn_ClassAdmin_GroupOperations = True
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Clicked on OK button")
				End If  

				'Remove the assigned image
				Case "RemoveImage"																									
							'Select the Class node from the tree
							sNodeName=dicGroupOperations.Item("NodeName")
							bReturn=Fn_ClassAdmin_TreeNodeOperation("Select",sNodeName,"")
							If bReturn = false Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Select ["+sNodeName+"] Node.")
									Fn_ClassAdmin_GroupOperations = false
									Exit Function 	
							Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully Selected ["+sNodeName+"] Node.")
									Fn_ClassAdmin_GroupOperations = true
							End If
							'Set  Edit Mode
							bReturn =  Fn_ClassAdmin_ToolbarOperations("Edit")
							If bReturn = false Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Perform [Edit] Opeartion.")
										Fn_ClassAdmin_GroupOperations = false
										Exit Function 	
							Else
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully Performed [Edit] Opeartion.")
										Fn_ClassAdmin_GroupOperations = true
							End If
							'Click on Delete Image Button
							objClassAdminApplet.JavaButton("DeleteImage").Click micLeftBtn
							 If Err.Number < 0 Then
										Fn_ClassAdmin_GroupOperations = False
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Click on [Delete Image] Button" )								
										Exit Function 
							Else
										Call Fn_ReadyStatusSync(3)
										wait(3)
										Fn_ClassAdmin_GroupOperations = True
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Clicked on [Delete Image] Button")							
							End If

                '---------------------------------------------------------------------------------Modify Case-------------------------------------------------------------------------------------------

					Case "Modify"
							'Select the Class node from the tree that has to be Modified
							sNodeName=dicGroupOperations.Item("NodeName")
							bReturn=Fn_ClassAdmin_TreeNodeOperation("Select",sNodeName,"")
							If bReturn = false Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Select ["+sNodeName+"] Node.")
									Fn_ClassAdmin_GroupOperations = false
									Exit Function 	
							Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully Selected ["+sNodeName+"] Node.")
									Fn_ClassAdmin_GroupOperations = true
							End If
							'Set  Edit Mode
							bReturn =  Fn_ClassAdmin_ToolbarOperations("Edit")
							If bReturn = false Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Perform [Edit] Opeartion.")
										Fn_ClassAdmin_GroupOperations = false
										Exit Function 	
							Else
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully Performed [Edit] Opeartion.")
										Fn_ClassAdmin_GroupOperations = true
							End If

				'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

                '--------------------------------------------------------------------------------Delete Case--------------------------------------------------------------------------------------
			Case "Delete"
						'Select the class node which  has to be deleted
						sNodeName=dicGroupOperations.Item("NodeName")
						bReturn=Fn_ClassAdmin_TreeNodeOperation("Select",sNodeName,"")
						If bReturn = false Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Select ["+sNodeName+"] Node.")
								Fn_ClassAdmin_GroupOperations = false
								Exit Function 	
						Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully Selected ["+sNodeName+"] Node.")
								Fn_ClassAdmin_GroupOperations = true
						End If
						'Set  Edit Mode
						bReturn =  Fn_ClassAdmin_ToolbarOperations("Edit")
						If bReturn = false Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Perform [Edit] Opeartion.")
									Fn_ClassAdmin_GroupOperations = false
									Exit Function 	
						Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully Performed [Edit] Opeartion.")
									Fn_ClassAdmin_GroupOperations = true
						End If
						'Delete the selected Class
						bReturn =  Fn_ClassAdmin_ToolbarOperations("Delete")
						If bReturn = false Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Perform [Delete] Opeartion.")
									Fn_ClassAdmin_GroupOperations = false
									Exit Function 	
						Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully Performed [Delete] Opeartion.")
									Fn_ClassAdmin_GroupOperations = true
						End If

						If  JavaDialog("ConfirmationMessage").Exist(5) Then
								Set oConDlg = JavaDialog("ConfirmationMessage")
						ElseIf objClassAdminApplet.JavaDialog("DeleteConfirmation").Exist(5) Then
							Set oConDlg = objClassAdminApplet.JavaDialog("DeleteConfirmation")
						ElseIf JavaDialog("Delete Preference(s)").Exist(5) Then
							Set oConDlg =JavaDialog("Delete Preference(s)")
						Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Confirmation Dialog Does not exist.")												
							Fn_ClassAdmin_DeleteObjects = False
							Set oConDlg = Nothing
							Exit Function
						End If

						If oConDlg.Exist(5) = True Then
							oConDlg.Activate
							 If Err.Number < 0 Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Confirmation Dialog Does not exist.")												
									Fn_ClassAdmin_DeleteObjects = False
									Set oConDlg = Nothing
									Exit Function
							Else
									wait(3)
									'Vallari - Ready is not Ready as Delete COnfirmation dialog in ON
									'Call Fn_ReadyStatusSync(3)
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully verified the existence of Confirmation Dialog.")								
							End If
	
							
							'Click on 'Yes' button
							oConDlg.JavaButton("Yes").Click micLeftBtn
							If Err.Number < 0 Then
										Fn_ClassAdmin_ClassOperations = False
										oConDlg.JavaButton("No").Click micLeftBtn
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed toClick on Yes button" )
										Set oConDlg = Nothing
										Exit Function 
							Else							
										wait(3)
										Call Fn_ReadyStatusSync(5)
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Clicked on Yes button")
							End If	
			End If
			Set oConDlg = Nothing
		
		
			'--------------------------------------------------------------------------------End of Delete ------------------------------------------------------------------------------------------


	 End Select
	wait(5)	
	 ' Dictionary  to handle - Name of Group, ICO Creation Chk Box
	  For iCounter = 0 to dicCount - 1
			If  dicItems(iCounter) <> "" Then
			   Select Case dicKeys(iCounter)

			    Case "GroupName"
						sGroupName=dicItems(iCounter)
						If  sGroupName <> "" Then
								 objClassAdminApplet.JavaEdit("GroupName").Set sGroupName
								 If Err.Number < 0 Then
										Fn_ClassAdmin_GroupOperations = False
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to SET text in Group Name Edit Box" )								
										Exit Function 
								Else
										Fn_ClassAdmin_GroupOperations = True
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully SET text in Group Name Edit Box")								
								End If
						End If


				Case "Check_ICOCreation"
						sICOCreation=dicItems(iCounter)
						If  sICOCreation <> "" Then						      
								 JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaCheckBox("ICO Creation").SetTOProperty "Attached Text",sICOCreation
								 objClassAdminApplet.JavaCheckBox("ICO Creation").Set "ON"
								 If Err.Number < 0 Then
										Fn_ClassAdmin_GroupOperations = False								
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to SET CheckBox ["+sICOCreation+"]" )								
										Exit Function 
								Else
										Fn_ClassAdmin_GroupOperations = True								
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully SET CheckBox ["+sICOCreation+"]")												
								End If
						End If

                    Case "SaveCurrentInstance"
						sSaveCurrentInstance=dicItems(iCounter)
						If  sSaveCurrentInstance <> "" Then
								bReturn =  Fn_ClassAdmin_ToolbarOperations("Save")
								 If bReturn = false Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Perform ["+sSaveCurrentInstance+"] Opeartion.")
										Fn_ClassAdmin_ClassOperations = false
										Exit Function 	
								Else
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully Performed ["+sSaveCurrentInstance+"] Opeartion.")
										Fn_ClassAdmin_ClassOperations = true
								End If
						End If

                     Case "Image"
							
							'Click on Add Image Button
							objClassAdminApplet.JavaButton("AddImage").Click micLeftBtn
							 If Err.Number < 0 Then
										Fn_ClassAdmin_GroupOperations = False
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Click on [Add Image] Button" )								
										Exit Function 
							Else
										Fn_ClassAdmin_GroupOperations = True
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Clicked on [Add Image] Button")							
							End If
							'Check existance of the Select Image Dialog Box
							Call Fn_ReadyStatusSync(3)
							If JavaDialog("SelectImageDialog").Exist(5) = True Then
										JavaDialog("SelectImageDialog").Activate
												 If Err.Number < 0 Then
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Select Image Dialog Does not exist.")
														Fn_ClassAdmin_GroupOperations = False
														Exit Function
												Else
														wait(3)
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully verified the existence of Select Image Dialog.")
														Fn_ClassAdmin_GroupOperations = True							
												End If
							End If
							'Paste the url in file name edit box
							sAddImageUrl=dicItems(iCounter)
							If  sAddImageUrl <> "" Then
											JavaDialog("SelectImageDialog").JavaEdit("FileName").Set sAddImageUrl
											 If Err.Number < 0 Then
													Fn_ClassAdmin_GroupOperations = False
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to SET url ["+sAddImageUrl+"] in File name Edit Box" )								
													Exit Function 
											Else
													Fn_ClassAdmin_GroupOperations = True
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully SET url ["+sAddImageUrl+"] in File name Edit Box")								
											End If
							End If
						   'Click on Import button
							JavaDialog("SelectImageDialog").JavaButton("Add").Click micLeftBtn
								 If Err.Number < 0 Then
											Fn_ClassAdmin_GroupOperations = False
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Click on [ Import ] Button" )								
											Exit Function 
								Else
											Call Fn_ReadyStatusSync(3)
											wait(3)
											Fn_ClassAdmin_GroupOperations = True
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Clicked on [ Import ] Button")							
								End If


				End Select

			End If
	Next

	Fn_ClassAdmin_GroupOperations = True
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Performed ["+sAction+"] Group Operation")
End Function


'*********************************************************  Function Operates on Tabs in All Windows*********************************************************************

'Function Name		:					Fn_ClassAdmin_TabOpeartions

'Description			 :		 		    Action  performed :-
'																	1. Node Select
'																	2. Node Expand
'																	3. Node Collapse
'																	4. Exist
'																	5. DoubleClick
'																    

'Parameters			   :	 			1. StrAction: Action to be performed
'												2.sTabType: 2 Types exists for tab : MainTab & SubTab
' 												3. sTabName: Name of tab to be operated
'												4. sDetails : Parameter to be used in feature 
												 

'Return Value		   : 			 True/False

'Pre-requisite			:		 	 Classification  pane should be displayed.

'Examples				:			 Fn_ClassAdmin_TabOpeartions("Activate","Maintab","Dictionary","")
'											Fn_ClassAdmin_TabOpeartions("Activate","Subtab","Class Attributes","")
'

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Prasanna 				10-Dec-2010	       1.0														Prasanna
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_ClassAdmin_TabOpeartions(sAction,sTabType,sTabName,sDetails)
	GBL_FAILED_FUNCTION_NAME="Fn_ClassAdmin_TabOpeartions"

	Dim bReturn,sTabObject,StrActivatedTab
	
	Dim objClassAdminApplet
	
	Set objClassAdminApplet =  Fn_SISW_ClassAdmin_GetObject("ClassAdminApplet")

	Select Case sAction
				Case "Activate"
						  Select Case sTabType 
			  						  Case "Maintab" 
												Set sTabObject = objClassAdminApplet.JavaTab("MainTab")
									  Case "Subtab"
												Set sTabObject = objClassAdminApplet.JavaTab("SubTab")
									  Case "ImageTab"
												Set sTabObject =objClassAdminApplet.JavaTab("ImageTab")												
						  End Select 

						'Check the tab name & activate the same						  
						if Fn_SISW_UI_JavaTab_Operations("","Select", sTabObject,"",sTabName ) = false then 
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_ClassAdmin_TabOpeartions : Failed to Activate the Tab [" + sTabType + "].")
							Fn_ClassAdmin_TabOpeartions = false
							Set objClassAdminApplet	= nothing						
						else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully [ "+sAction+"d ]  tab ["+sTabName+"].")
							Fn_ClassAdmin_TabOpeartions = true	
							Set objClassAdminApplet	= nothing						
						End if 
						 'sTabObject.Select  sTabName
						 'wait(2)

'						  If Err.Number <> 0 Then
'									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to ["+sAction+"]  tab ["+sTabName+"]")
'									Fn_ClassAdmin_TabOpeartions = false
'									Set sTabObject = Nothing
'									Exit Function 	
'						  Else
'									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully [ "+sAction+"d ]  tab ["+sTabName+"].")
'									Fn_ClassAdmin_TabOpeartions = true
'									Set sTabObject = Nothing
'						  End If
					' [TC1123-20161108-21_11_2016-JotibaT-Maintenance]	  
				Case "VerifyActivate"
				
						Select Case sTabType 
								Case "Maintab" 
									Set sTabObject =objClassAdminApplet.JavaTab("MainTab")
								Case "Subtab"
									Set sTabObject =objClassAdminApplet.JavaTab("SubTab")
								Case "ImageTab"
									Set sTabObject =objClassAdminApplet.JavaTab("ImageTab")												
						End Select 
							StrActivatedTab=sTabObject.GetROProperty("value")
							If Trim(sTabName)=Trim(StrActivatedTab) Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verified [" + sTabName + "] tab is currently activated")
								Fn_ClassAdmin_TabOpeartions=True
							Else 
								Fn_ClassAdmin_TabOpeartions = false
							End If
                End Select	 
End Function 

' *********************************************************  Function do Operation on Existing Key LOV List*********************************************************************

'Function Name		:					Fn_ClassAdmin_ExistingKeyLov_List

'Description			 :		 		    Action  performed :-
'																	1. Node Select
'																	2. Node Exist
'																	
'																    

'Parameters			   :	 			1. sAction: Action to be performed
'													2.sNode: Fully qulified tree Path (delimiter as ':') 
'

'Return Value		   : 			 True/False

'Pre-requisite			:		 	 Classification  pane should be displayed.

'Examples				:			' Fn_ClassAdmin_ExistingKeyLov_List("Select","-60009  Ignore for Optimization")
											'Fn_ClassAdmin_ExistingKeyLov_List("Exist","-60009  Ignore for Optimization")
'

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Vidya Kulkarni				13-Dec-2010	       1.0														Prasanna
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function Fn_ClassAdmin_ExistingKeyLov_List(sAction,sNode)
		GBL_FAILED_FUNCTION_NAME="Fn_ClassAdmin_ExistingKeyLov_List"
		''Declaration Of Variable
		Dim objExistingLOVList,iItemCount,iCounter,sTreeItem

		 'Initilisation Of Variable
	Set objExistingLOVList = JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaList("ExistingLOVList")

	 If  objExistingLOVList.Exist(5)Then
			Err.Clear
			Select Case sAction

			Case "Select"
						objExistingLOVList.Select sNode
						Wait(2)
						If Err.Number <0  Then
								Fn_ClassAdmin_ExistingKeyLov_List = False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to select Node [ " + sNode + "] of Classification Admin Tree." ) 
								Set objExistingLOVList = Nothing
								Exit Function 
						Else
								Fn_ClassAdmin_ExistingKeyLov_List = True
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully selected  Node [ " + sNode + "] of Classification Admin Tree.") 
						End If

			Case "Exist"
						iItemCount = objExistingLOVList.GetROProperty( "items count")
						For iCounter=0 To (iItemCount-1)
								sTreeItem = objExistingLOVList.GetItem(iCounter)
									If Trim (Lcase(sTreeItem)) = Trim(Lcase(sNode)) Then
											Fn_ClassAdmin_ExistingKeyLov_List = True
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully found node [" + StrNode + "] of Classification Admin Tree." )	
											Exit For
									End If
						Next 

						If  Cint(iCounter) = Cint (iItemCount) Then
								Fn_ClassAdmin_ExistingKeyLov_List = False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to  found node [" + StrNode + "] of Classification Admin Tree." )	
								Set objExistingLOVList = Nothing
								Exit Function 
  					 End If
			End Select
		End If
End Function

' *********************************************************  Function do Select the Metric/ Non-Metric Format*********************************************************************

'Function Name		:					Fn_ClassAdmin_FormatTypeSelect

'Description			 :		 		    Select the Fomrat Type based on Information passed in Array
'																														
'                                                                   
'Parameters			   :	 			1. aFormatInfo: informAtion to select the Format Details
'												aFormatInfo(0) = Format Type eg. Metric/Non-Metric							
'												aFormatInfo(1) = Format Type eg. Real/String/Integer								
'												aFormatInfo(2) = Format Type Length eg. 8							
'												aFormatInfo(3) = Format Type : To be Selected from List eg. Upper- and Lowercase
'														
'

'Return Value		   : 			 	True/False

'Pre-requisite			:		 	 	Classification  pane should be displayed.

'Examples				:			 	aFormatInfo = Array("Metric","String","8","Upper- and Lowercase")
'											   Fn_ClassAdmin_FormatTypeSelect(aFormatInfo)
'											  aFormatInfo = Array("NonMetric","Real","2:2","Force positive number")
'											  Fn_ClassAdmin_FormatTypeSelect(aFormatInfo)

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Prasanna				13-Dec-2010	       1.0														Prasanna
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'								 Jeevan Mutha				7-June-2012	       														
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_ClassAdmin_FormatTypeSelect(aFormatInfo)
	GBL_FAILED_FUNCTION_NAME="Fn_ClassAdmin_FormatTypeSelect"
   On error resume next
	Dim objApplet
	Set objApplet = JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow")
   If IsArray(aFormatInfo) = false Then		
		Fn_ClassAdmin_FormatTypeSelect = false	
		Exit Function 
   End If
	Err.Clear
		'Click on Checkbox 
	If lcase(trim(aFormatInfo(0))) = "metric" Then
			objApplet.JavaCheckBox("ChkMetricFormat").Set "ON"
	Else
			objApplet.JavaCheckBox("ChkFormatNonMetric").Set "ON"		
	End If

	If Err.Number < 0  Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Click on ["+aFormatInfo(iCounter)+"] Format Button")
			Fn_ClassAdmin_FormatTypeSelect = false					
			Exit Function 
	Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully Clicked on  ["+aFormatInfo(iCounter)+"] Format Button")
			Fn_ClassAdmin_FormatTypeSelect = true
			wait(3)		
	End If

	'Open the Dialog	
	 objApplet.JavaStaticText("FormatDialogStText").DblClick 0,0,"LEFT" 
	 wait (3)
	 If objApplet.JavaDialog("Format Dialog").Exist(5)=False Then
   			Err.Clear
			objApplet.JavaStaticText("FormatDialogStText").SetTOProperty "index",1
			objApplet.JavaStaticText("FormatDialogStText").DblClick 0,0,"LEFT" 
	 End If

	If objApplet.JavaDialog("Format Dialog").Exist(5)=False Then       ' Added for 0718 build onwards
   			Err.Clear 
			objApplet.JavaStaticText("FormatDialogStText").SetTOProperty "index", 0
			objApplet.JavaStaticText("FormatDialogStText").DblClick 0,0,"LEFT" 			
			wait (3)
	 End If

	If Err.Number < 0  Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Click on Format Dialog Java Static Text")
			Fn_ClassAdmin_FormatTypeSelect = false					
			Exit Function 
	Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully Clicked on Format Dialog Java Static Text")
			Fn_ClassAdmin_FormatTypeSelect = true
			wait(3)		
	End If

    'Check the Existance of Format Dialog
	If  objApplet.JavaDialog("Format Dialog").Exist(5) = false Then
			Fn_ClassAdmin_FormatTypeSelect = false					
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Format dialog does not exist")
			Exit Function 				
	End If

	'Set the tab
	if aFormatInfo(0) <> "" Then
			objApplet.JavaDialog("Format Dialog").JavaTab("FormatTypeTab").Select trim(aFormatInfo(1))
			If Err.Number < 0  Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Set  tab ["+aFormatInfo(1)+"]")
					Fn_ClassAdmin_FormatTypeSelect = false					
					Exit Function 
			Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully Set tab ["+aFormatInfo(1)+"]")
					Fn_ClassAdmin_FormatTypeSelect = true
					wait(3)		
			End If
	End If

	Wait 1
	If aFormatInfo(2) <> "" Then
	'set the value in text box
			If  instr(1, aFormatInfo(2),":") > 0 Then
					 aRealNoFormat = split(aFormatInfo(2),":",-1,1)  ' If the real no. is present then set values for Integer & Decimal edit box
		
					 'Set the value in Integer text box
					objApplet.JavaDialog("Format Dialog").JavaEdit("FormatLength").Set trim(aRealNoFormat(0))
					wait (2)
					If Err.Number < 0  Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Set value in Integer Format length text box ["+aRealNoFormat(0)+"]")
							Fn_ClassAdmin_FormatTypeSelect = false			
							objApplet.JavaDialog("Format Dialog").JavaButton("Cancel").Click micLeftBtn		
							Exit Function 
					Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully Set value in Integer Format length text box ["+aRealNoFormat(0)+"]")
							Fn_ClassAdmin_FormatTypeSelect = true
							wait(2)		
					End If	
										
					'Set the value in Decimal text box
					objApplet.JavaDialog("Format Dialog").JavaEdit("RealNumbers").Set trim(aRealNoFormat(1))
					If Err.Number < 0  Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Set value in Decimal Format length text box ["+aRealNoFormat(1)+"]")
							Fn_ClassAdmin_FormatTypeSelect = false			
							objApplet.JavaDialog("Format Dialog").JavaButton("Cancel").Click micLeftBtn		
							Exit Function 
					Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully Set value in Decimal Format length text box ["+aRealNoFormat(1)+"]")
							Fn_ClassAdmin_FormatTypeSelect = true
							wait(2)		
					End If	
			Else
					objApplet.JavaDialog("Format Dialog").JavaEdit("FormatLength").Set trim(aFormatInfo(2))
					If Err.Number < 0  Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Set value in Format length text box ["+aFormatInfo(2)+"]")
							Fn_ClassAdmin_FormatTypeSelect = false			
							objApplet.JavaDialog("Format Dialog").JavaButton("Cancel").Click micLeftBtn		
							Exit Function 
					Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully Set value in Format length text box ["+aFormatInfo(2)+"]")
							Fn_ClassAdmin_FormatTypeSelect = true
							wait(3)		
					End If
			End If
     End If           

	'set the value in List box
	If aFormatInfo(3) <>"" Then
			objApplet.JavaDialog("Format Dialog").JavaList("FormatTypeList").Select trim(aFormatInfo(3))
			wait (2)			
			If Err.Number < 0  Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Set value in Format type list box ["+aFormatInfo(3)+"]")
					Fn_ClassAdmin_FormatTypeSelect = false			
					objApplet.JavaDialog("Format Dialog").JavaButton("Cancel").Click micLeftBtn		
					Exit Function 
			Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully Set value in Format type list box ["+aFormatInfo(3)+"]")
					Fn_ClassAdmin_FormatTypeSelect = true
					wait(2)		
			End If
	End If

	'Click on OK Button
	objApplet.JavaDialog("Format Dialog").JavaButton("OK").Click micLeftBtn
		If Err.Number < 0  Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Click on OK Button")
			Fn_ClassAdmin_FormatTypeSelect = false			
			objApplet.JavaDialog("Format Dialog").JavaButton("Cancel").Click micLeftBtn		
			Exit Function 
	Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully Clicked on OK Button")
			Fn_ClassAdmin_FormatTypeSelect = true
			wait(3)		
	End If
End Function 


''*********************************************************  Function do Add the Atrributes in Class Atrribute Tab*********************************************************************

'Function Name		:					 Fn_ClassAdmin_SearchAddAttributes

'Description			 :		 		    Action  performed :-
'																	1. Select Search criteria																	
'																	2. Enter search text
'																	3. Select No. of entries : If kept blank will select the first result.
'																	4 Select search result																
'																   5 Add Attribute to class

'Parameters			   :	 			1. StrSearchType:Search criteria
'													2.StrSearchText:Enter search text
'													3. iSelectEntries No. of entries : If kept blank will select the first result.
'													4.aName:Select Multiple values 
' 												   5. aValue:values based on columns

'Return Value		   : 			 True/False

'Pre-requisite			:		 	 Class Attribute Tab should exist

'Examples				:			 Fn_ClassAdmin_SearchAddAttributes("Attribute ID","-60009","","","")

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										     Shobha				14-Dec-2010	       1.0
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function Fn_ClassAdmin_SearchAddAttributes(StrSearchType,StrSearchText,iSelectEntries,aName,aValue)
	GBL_FAILED_FUNCTION_NAME="Fn_ClassAdmin_SearchAddAttributes"
   On Error Resume Next
   'Variable Initialization

	Dim iCounterSelect, iCounter, aEntries, objResultTable, iRows,objAddAttribute
	Dim objClassAdminApplet 
	
	Set objClassAdminApplet =  Fn_SISW_ClassAdmin_GetObject("ClassAdminApplet")
	Err.Clear
	if Fn_SISW_UI_Object_Operations("Fn_ClassAdmin_SearchAddAttributes","Exist", ObjClassAdminApplet.JavaTab("SubTab"),"5") then
	'If JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaTab("SubTab").Exist(5) Then
				'Click on Add Attribute button
				JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaButton("Add Attribute").Click micLeftBtn
					If Err.Number < 0 Then
							Fn_ClassAdmin_SearchAddAttributes=False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Click on Add Attribute button " )
					Else
							Fn_ClassAdmin_SearchAddAttributes	=True					
							  wait(2)
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Clicked on Add Attribute button ")
                  			End if 
				if Fn_SISW_UI_Object_Operations("Fn_ClassAdmin_SearchAddAttributes","Exist", JavaDialog("Add Attribute"),"5") then
				'If  JavaDialog("Add Attribute").Exist(5) Then
						Set objAddAttribute=JavaDialog("Add Attribute")	
				Else
						Set objAddAttribute=objClassAdminApplet.JavaDialog("Add Attribute")
				End If

				Set objResultTable=objAddAttribute.JavaTable("ResultTable")
				
				'Click on Select Filter button	
				If trim(StrSearchType) <> ""  Then
							objAddAttribute.JavaCheckBox("ChkCriteria").Set "ON"
							If Err.Number < 0 Then
									Fn_ClassAdmin_SearchAddAttributes=False
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Click Search Criteria Button" )		
							Else
									Fn_ClassAdmin_SearchAddAttributes	=True		
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Clicked on Search Criteria Button")
							 End if             	
				
							'objSearchCriteriaStatic.Click 1, 1, "LEFT"
							wait(1)									
							objAddAttribute.JavaList("CriteriaList").Select StrSearchType
							wait(1)     
							If Err.Number < 0 Then
									Fn_ClassAdmin_SearchAddAttributes=False
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select Search Criteria" )						
							Else
									Fn_ClassAdmin_SearchAddAttributes	=True											
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully selected search Criteria"+ StrSearchType + " of Add Attributes")
							 End if 

							 If objAddAttribute.JavaList("CriteriaList").Exist(5) Then
									objAddAttribute.JavaCheckBox("ChkCriteria").DblClick 1,1,"LEFT"
									If Err.Number < 0 Then
											Fn_ClassAdmin_SearchAddAttributes=False
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Click Search Criteria Button" )		
									Else
											Fn_ClassAdmin_SearchAddAttributes	=True		
											wait(3)									
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Clicked on Search Criteria Button")
									 End if  
							 End If
				End If

					'Set text to Search
					objAddAttribute.JavaEdit("SearchText").object.setText(trim(StrSearchText))
					If Err.Number < 0 Then
							Fn_ClassAdmin_SearchAddAttributes=False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to enter search text " )
							
					Else
							Fn_ClassAdmin_SearchAddAttributes	=True
                        	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully entered search text ")							
					End if 

					'Click on Search button
					objAddAttribute.JavaButton("Search").Click micLeftBtn
					If Err.Number < 0 Then
							Fn_ClassAdmin_SearchAddAttributes=False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to click on Search button" )							
					Else
							Fn_ClassAdmin_SearchAddAttributes	=True
							Call Fn_ReadyStatusSync(1)
							'wait(2)					 
                            Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully click on Search button ")
                     End if 

					' If user has not given no. of entries to be selected then select the first entry only.
					If iSelectEntries = "" Then	 
							objResultTable.SelectRow 0							
					Else
						aEntries = split(iSelectEntries,"~")
						if cInt(objAddAttribute.JavaButton("ICADictionary_LoadAll").GetROProperty ("enabled")) = 1 Then
							objAddAttribute.JavaButton("ICADictionary_LoadAll").Click micLeftBtn
							Call Fn_ReadyStatusSync(10)
						End If
						Call Fn_ReadyStatusSync(1)
						iRows = cInt(objResultTable.GetROProperty("rows"))
						For iCounter = 0 to UBound(aEntries)
							For iCounterSelect = 0 to irows - 1
								If cstr(objResultTable.GetCellData(iCounterSelect,"Attribute ID")) =  aEntries(iCounter) Then
									objResultTable.ExtendRow iCounterSelect   
									wait(1)	
									Exit for
								End If
							Next
						Next
					End If
					If Err.Number < 0 Then
										Fn_ClassAdmin_SearchAddAttributes=False
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Select Row from Table" )
										objResultTable=Nothing										
					Else
										Fn_ClassAdmin_SearchAddAttributes	=True
										Call Fn_ReadyStatusSync(1)
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully selected Row from Table")						  
					End If
					'Click on OK button
					objAddAttribute.JavaButton("OK").Click micLeftBtn				
					If Err.Number < 0 Then
										Fn_ClassAdmin_SearchAddAttributes=False
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select Search result" )
										objResultTable=Nothing										
					Else
										Fn_ClassAdmin_SearchAddAttributes	=True
										Call Fn_ReadyStatusSync(1)
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully selected Search result ")						  
					End If

					'CLick on Yes button for confirmation

'					JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("AddAttributeConfirmation").JavaButton("Yes").Click micLeftBtn
					JavaDialog("Add Attribute").JavaButton("Yes").Click micLeftBtn
					If Err.Number < 0 Then
										Fn_ClassAdmin_SearchAddAttributes=False
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Click on Yes Button" )
										objResultTable=Nothing										
					Else
										Fn_ClassAdmin_SearchAddAttributes	=True
										Call Fn_ReadyStatusSync(1)
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Clicked on Yes Buttont ")						  
					End If
					
					'Click on Ok Button of the Error Dialog.  ' Added by Sneha 03-May-2011 - Updated by Vallari on 20May11
					Dim objErrDialog
					if Fn_SISW_UI_Object_Operations("Fn_ClassAdmin_CreateAttribute","Exist", JavaDialog("AddAttributeError"),"3") then
					'If JavaDialog("AddAttributeError").Exist(3) Then
						Set objErrDialog = JavaDialog("AddAttributeError")
					Else
						Set objErrDialog = JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("AddAttributeComplete").JavaDialog("Add AttributeError")
					End If
					if Fn_SISW_UI_Object_Operations("Fn_ClassAdmin_CreateAttribute","Exist", objErrDialog,"3") then
					'If  objErrDialog.Exist(3) Then
										 objErrDialog.JavaButton("OK").Click micLeftBtn 
										If Err.Number < 0 Then
													Fn_ClassAdmin_SearchAddAttributes=False
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Click on OK Button.Failed to Handle the Error Dialog" )
													objResultTable=Nothing										
										Else
													Fn_ClassAdmin_SearchAddAttributes	=True
													Call Fn_ReadyStatusSync(1)
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Clicked on OK Button. Error Dialog Handled Successfully")						  
										End If
										Set objErrDialog = Nothing
'										Click on the icon of an information bubble of the Add Attribute Dialog
					end if
							if Fn_SISW_UI_Object_Operations("Fn_ClassAdmin_CreateAttribute","Exist", JavaDialog("Add Attribute"),"3") then	
							'If JavaDialog("Add Attribute").Exist(3) Then
											Set objErrDialog = JavaDialog("Add Attribute")
							ElseIf Fn_SISW_UI_Object_Operations("Fn_ClassAdmin_CreateAttribute","Exist", JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("AddAttributeComplete"),"3") Then				
							'Elseif JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("AddAttributeComplete").Exist (3) then
											Set objErrDialog = JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("AddAttributeComplete")
							End If
							If Fn_SISW_UI_Object_Operations("Fn_ClassAdmin_CreateAttribute","Exist", objErrDialog,"3") then														
							'If objErrDialog.Exist (3) then											
											objErrDialog.JavaButton("InfoButton").Click micLeftBtn
									If Err.Number < 0  Then
													Fn_ClassAdmin_SearchAddAttributes=False
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Click on on Information Button" )
													objResultTable=Nothing											
													JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("AddAttributeComplete").JavaDialog("AddAttributeConfirmation").Close
									else
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Clicked on Information Button")						  
									End If
										Set objErrDialog = Nothing
										If JavaDialog("Add Attribute").JavaDialog("Handle error").Exist(3) Then
											Set objErrDialog = JavaDialog("Add Attribute").JavaDialog("Handle error")
										Else
											Set objErrDialog = JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("AddAttributeComplete").JavaDialog("Handle Error")
										End If
										objErrDialog.JavaButton("Yes").Click micLeftBtn
										If Err.Number < 0  Then
													Fn_ClassAdmin_SearchAddAttributes=False
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Click on Yes  Button" )
													objResultTable=Nothing											
													JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("AddAttributeComplete").JavaDialog("AddAttributeConfirmation").Close
										else
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Clicked on Yes  Button")						  
													Fn_ClassAdmin_SearchAddAttributes	=True
													Exit function
										End If
										Set objErrDialog = Nothing
                                        
end if
					'Click on Successful Completion button
					if Fn_SISW_UI_Object_Operations("Fn_ClassAdmin_CreateAttribute","Exist", JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("AddAttributeComplete"),"5") then
					'If JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("AddAttributeComplete").Exist(5) Then
						JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("AddAttributeComplete").JavaButton("OK").Click micLeftBtn
						If Err.Number < 0 Then
								Fn_ClassAdmin_SearchAddAttributes=False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Click on OK Button.Failed to Add Atrribute" )
								objResultTable=Nothing										
						Else
								Fn_ClassAdmin_SearchAddAttributes	=True
								Call Fn_ReadyStatusSync(1)
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Clicked on OK Button. Attribute Added Successfully")						  
						End If
	
					End If
	End If

End Function

'*********************************************************  Function do Creation of Attribute *********************************************************************

'Function Name		:					Fn_ClassAdmin_CreateAttribute

'Description			 :		 		    Action  performed :-
'																	1. Create new Attribute
'																	
'																    

'Parameters			   :	 			 1. sAttributeID
'													 2. aDictionaryValue : Array of Values splited with ":"   ex ->  Name:abc
'													3.aMetricFormat 
' 												   4. aNonMetricFormat 

'Return Value		   : 			 True/False

'Pre-requisite			:		 	 Dictionary Tab Selected  

'Examples				:			 aDictionaryValue = Array ( "Name:ABC" )
'												aMetricFormat = Array("Metric","String","8","Upper- and Lowercase","Default Value:abcd")
'										or 		aMetricFormat = Array("Metric","Integer","3","Force positive number","Default Value:150","Maximum Value:200","Minimum Value:100","Unit:Length~Meters")
'												Call Fn_ClassAdmin_CreateAttribute ( "", aDictionaryValue, aMetricFormat, "" )
'	                                              ---------------------------------------------------------------------------------------------------------------------
												'aDictionaryValue = Array ( "Name:ABC" )											
												'aMetricFormat = Array("Metric","KeyLOV","-2820","","Default Value:2 Downwards")

'												Call Fn_ClassAdmin_CreateAttribute ( "", aDictionaryValue, aMetricFormat, "" )


												
'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Harshal Tanpure			14-Dec-2010	       1.0														      Prasanna B.
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Jeevan Mutha			   7-June-2012	                  													    
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Sandeep N			   22-Mar-2013							Modified case "Unit"	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Sandeep N			   26-Mar-2013							Modified case "Unit@2"	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Sandeep N			   28-Mar-2013						Added condition to modify Tab name KeyLov to Key LOV
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function Fn_ClassAdmin_CreateAttribute ( sAttributeID , aDictionaryValue , aMetricFormat , aNonMetricFormat )
	GBL_FAILED_FUNCTION_NAME="Fn_ClassAdmin_CreateAttribute"
	 Dim bReturn, iCount, aTextSet
        Dim arrSelectUnitMenu,abs_x,abs_y,DeviceReplay,arrSelectUnitMenu1
        Dim objAddNewClass,ObjClassAdminApplet
		Err.Clear
''''''''TO rectify Spaces in aMetricFormat
''''''''' Ex: if KeyLOVhas been passed it would be converted to Key LOV
	If IsArray(aMetricFormat) Then             '''''''''to check if is it arrray
		If Ubound(aMetricFormat ) > 0 Then          ''''''''to check its ubounds value
				If trim(lcase(aMetricFormat(1))) = "keylov" Then
					aMetricFormat(1) = "Key LOV"
				End If
		End If
	End If

	If IsArray(aNonMetricFormat) Then             '''''''''to check if is it arrray
		If Ubound(aNonMetricFormat ) > 0 Then          ''''''''to check its ubounds value
				If trim(lcase(aNonMetricFormat(1))) = "keylov" Then
					aNonMetricFormat(1) = "Key LOV"
				End If
		End If
	End If

	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		' Click on Create New Instance
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	If  sAttributeID <> "Modify" Then	
		bReturn = Fn_ClassAdmin_ToolbarOperations("NewInstance")
	Else
		bReturn = true
	End if 
	Set objAddNewClass  =  Fn_SISW_ClassAdmin_GetObject("AddNewClass") 
	Set ObjClassAdminApplet  =  Fn_SISW_ClassAdmin_GetObject("ClassAdminApplet") 
	
		If bReturn = True Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully selected [ Create a new Instance ] from Toolbar")

				'''''''''''''''''''''''''''''''''''''''''
				' Assign Attribute ID
				''''''''''''''''''''''''''''''''''''''''
				
				If Fn_SISW_UI_Object_Operations("Fn_ClassAdmin_CreateAttribute","Exist", objAddNewClass,"") Then
						
						''''''''''''''''''''''''''''''''''''''''''''''''
						' Click on Assign button
						''''''''''''''''''''''''''''''''''''''''''''''''
						If sAttributeID = "" Then

'								JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("AddNewClass").JavaButton("Assign").Click
'								If Err.Number < 0 Then
'										Fn_ClassAdmin_CreateAttribute = False
'										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to click Assign button" ) 
'										Exit Function 
'								End If
								Wait(2)
					
								If Fn_SISW_UI_JavaButton_Operations("Fn_ClassAdmin_CreateAttribute", "Click", objAddNewClass, "Assign") = false Then
										Fn_ClassAdmin_CreateAttribute = False
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to click Assign button" ) 
										Exit Function 									
								Else
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully clicked on Assign button" ) 								
								End If
								
								wait(2) 
								
								If Fn_SISW_UI_JavaButton_Operations("Fn_ClassAdmin_CreateAttribute", "Click", objAddNewClass, "OK") = false Then
										Fn_ClassAdmin_CreateAttribute = False
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to click OK button" ) 
										Exit Function 									
								Else
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully clicked on OK button" ) 								
								End If

'								JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("AddNewClass").JavaButton("OK").Click
'								If Err.Number < 0 Then
'										Fn_ClassAdmin_CreateAttribute = False
'										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to click OK button" ) 
'										Exit Function 
'								End If
								Wait(2)
'
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Assigned Attribute ID")
						Else
								
								''''''''''''''''''''''''''''''''''''''''''''''''''''
								'write Attirbue ID in text box
								''''''''''''''''''''''''''''''''''''''''''''''''''''
								If objAddNewClass.exist(2)=True Then
									objAddNewClass.JavaEdit("Class ID").settoproperty "attached text","Attribute ID"
									If objAddNewClass.JavaEdit("Class ID").exist(2)=True Then
										objAddNewClass.JavaEdit("Class ID").Set sAttributeID
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Assign Attribute ID [ " + sAttributeID + "] " ) 
									Else
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Assign Attribute ID [ " + sAttributeID + "] " ) 
											Fn_ClassAdmin_CreateAttribute = False
											Exit Function 
									End If
									
								Else
									if Fn_UI_Object_SetTOProperty_ExistCheck("Fn_ClassAdmin_CreateAttribute",objAddNewClass,"attached text", "Attribute ID") then 
											if Fn_SISW_UI_JavaEdit_Operations("Fn_ClassAdmin_CreateAttribute", "Set", objAddNewClass, "Class ID", sAttributeID)	= false then 
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Assign Attribute ID [ " + sAttributeID + "] " ) 
												Fn_ClassAdmin_CreateAttribute = False
												Exit Function 
											Else
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Assign Attribute ID [ " + sAttributeID + "] " ) 
											End if 
										else
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Assign Attribute ID [ " + sAttributeID + "] " ) 
											Fn_ClassAdmin_CreateAttribute = False
											Exit Function 							
										End if 
								End If
'								JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("AddNewClass").JavaEdit("Class ID").SetTOProperty "attached text" , "Attribute ID"
'								JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("AddNewClass").JavaEdit("Class ID").Set sAttributeID 
'
'								If Err.Number < 0 Then
'										Fn_ClassAdmin_CreateAttribute = False
'										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Assign Attribute ID [ " + sAttributeID + "] " ) 
'										Exit Function 
'								End If
'								Wait(2)
'
								JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("AddNewClass").JavaButton("OK").Click
								If Err.Number < 0 Then
										Fn_ClassAdmin_CreateAttribute = False
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to click OK button" ) 
										Exit Function 
								End If
								Wait(2)

								''''''''''''''''''''''''''''''
								' Handle Error
								'''''''''''''''''''''''''''''
								if Fn_SISW_UI_Object_Operations("Fn_ClassAdmin_CreateAttribute","Exist", objAddNewClass,"") then 
									Call Fn_SISW_UI_JavaButton_Operations("Fn_ClassAdmin_CreateAttribute", "Click", objAddNewClass, "OK")
									Fn_ClassAdmin_CreateAttribute = False
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Assign Attribute ID [ " + sAttributeID + "] " ) 
									Exit Function 
								End if 
'								If  JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("AddNewClass").Exist Then
'										Fn_ClassAdmin_CreateAttribute = False
'										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Assign Attribute ID [ " + sAttributeID + "] " ) 
'
'										JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("AddNewClass").JavaButton("OK").Click
'										If Err.Number < 0 Then
'												Fn_ClassAdmin_CreateAttribute = False
'												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to click OK button" ) 
'												Exit Function 
'										End If
'										Exit Function 
'								End If

								'''''''''''''''''''''''''''''
								' Handle Error
								'''''''''''''''''''''''''''''
								
								If Fn_SISW_UI_Object_Operations("Fn_ClassAdmin_CreateAttribute","Exist", ObjClassAdminApplet.JavaDialog("NewAttributeID").JavaDialog("AttributeError"),"")  Then
										Fn_ClassAdmin_CreateAttribute = False
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Assign Attribute ID [ " + sAttributeID + "] " ) 

										Call Fn_SISW_UI_JavaButton_Operations("Fn_ClassAdmin_CreateAttribute", "Click", ObjClassAdminApplet.JavaDialog("NewAttributeID").JavaDialog("AttributeError"),"OK")
										Call Fn_SISW_UI_JavaButton_Operations("Fn_ClassAdmin_CreateAttribute", "Click", ObjClassAdminApplet.JavaDialog("NewAttributeID").JavaDialog("AttributeError"),"Cancel")
										'ObjClassAdminApplet.JavaDialog("AddNewClass").JavaButton("Cancel").Click
										'If Err.Number < 0 Then
												Fn_ClassAdmin_CreateAttribute = False
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to click OK button" ) 
												Exit Function 
										'End If
										Exit Function 								
								End If							
								
								
'								If JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("NewAttributeID").JavaDialog("AttributeError").Exist Then
'										Fn_ClassAdmin_CreateAttribute = False
'										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Assign Attribute ID [ " + sAttributeID + "] " ) 
'
'										JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("NewAttributeID").JavaDialog("AttributeError").JavaButton("OK").Click
'										If Err.Number < 0 Then
'												Fn_ClassAdmin_CreateAttribute = False
'												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to click OK button" ) 
'												Exit Function 
'										End If
'										JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("AddNewClass").JavaButton("Cancel").Click
'										If Err.Number < 0 Then
'												Fn_ClassAdmin_CreateAttribute = False
'												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to click OK button" ) 
'												Exit Function 
'										End If
'										Exit Function 								
'								End If							

								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Assigned Attribute ID")
						End If
    
				Else 
							If  sAttributeID <> "Modify" Then	
										Fn_ClassAdmin_CreateAttribute =FALSE
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : New Attribute ID dialog does not Exist" )
										Exit Function
							End If
				End If
		
				'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
				' Assign Values for Text Box
				''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
				
				For iCount = 0 to ubound(aDictionaryValue)
						aTextSet = Split ( aDictionaryValue(iCount), ":" , -1 , 1  )
						Select Case aTextSet(0)
								Case "Name","Short Name","Default Annotation"
										if Fn_UI_Object_SetTOProperty_ExistCheck("Fn_ClassAdmin_CreateAttribute",ObjClassAdminApplet.JavaEdit("AddAttributeEditBox"),"attached text", aTextSet(0)+":") then 
											Wait 2												
											if Fn_SISW_UI_JavaEdit_Operations("Fn_ClassAdmin_ClassOperations", "Set",  JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow"), "AddAttributeEditBox", aTextSet(1)) = false then 												
												Fn_ClassAdmin_CreateAttribute = False
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Set [ " + aTextSet(0) + "] = [ " + aTextSet(1) + "] " ) 										
											End If
										else
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Set [ " + aTextSet(0) + "] = [ " + aTextSet(1) + "] " ) 	
										End if	
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Set [ " + aTextSet(0) + "] = [ " + aTextSet(1) + "] " ) 							
'								Case "Name"
'										JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaEdit("AddAttributeEditBox").SetTOProperty "attached text" , aTextSet(0)+":" 
'										Call Fn_SISW_UI_JavaEdit_Operations("Fn_ClassAdmin_ClassOperations", "Set",  JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow"), "AddAttributeEditBox", aTextSet(1))
'		
'										If Err.Number < 0 Then
'												Fn_ClassAdmin_CreateAttribute = False
'												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Set [ " + aTextSet(0) + "] = [ " + aTextSet(1) + "] " ) 
'												Exit Function 
'										End If
'										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Set [ " + aTextSet(0) + "] = [ " + aTextSet(1) + "] " ) 
'
'								 Case "Short Name"
'										JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaEdit("AddAttributeEditBox").SetTOProperty "attached text" , aTextSet(0)+":" 
'										JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaEdit("AddAttributeEditBox").Set aTextSet(1)
'		
'										If Err.Number < 0 Then
'												Fn_ClassAdmin_CreateAttribute = False
'												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Set [ " + aTextSet(0) + "] = [ " + aTextSet(1) + "] " ) 
'												Exit Function 
'										End If
'										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Set [ " + aTextSet(0) + "] = [ " + aTextSet(1) + "] " ) 
'
'								Case "Default Annotation"
'										JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaEdit("AddAttributeEditBox").SetTOProperty "attached text" , aTextSet(0)+":" 
'										JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaEdit("AddAttributeEditBox").Set aTextSet(1)
'		
'										If Err.Number < 0 Then
'												Fn_ClassAdmin_CreateAttribute = False
'												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Set [ " + aTextSet(0) + "] = [ " + aTextSet(1) + "] " ) 
'												Exit Function 
'										End If
'										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Set [ " + aTextSet(0) + "] = [ " + aTextSet(1) + "] " ) 
'
						End Select				
				Next
				'''''''''''''''''''''''''''''''''''''''''''''
				'Set Value of Format
				''''''''''''''''''''''''''''''''''''''''''''
				If isArray ( aMetricFormat ) Then
						bReturn =  Fn_ClassAdmin_FormatTypeSelect(aMetricFormat)
						If bReturn = True Then
						'wait (3)
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Set Format " ) 
						Else
									Fn_ClassAdmin_CreateAttribute = False
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Set Format " ) 
									Exit Function
						End If				
				End If
				
				If Fn_SISW_UI_Object_Operations("Fn_ClassAdmin_CreateAttribute","Exist", ObjClassAdminApplet.JavaDialog("Potential problem found"),"") then
					Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_ClassAdmin_CreateAttribute",ObjClassAdminApplet.JavaDialog("Potential problem found").JavaButton("OK"),"attached text", "OK") 				
					Call Fn_SISW_UI_JavaButton_Operations("Fn_ClassAdmin_CreateAttribute", "Click", ObjClassAdminApplet.JavaDialog("Potential problem found"),"OK")
				End if	

'				If JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("Potential problem found").Exist(3) then
'					JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("Potential problem found").JavaButton("OK").SetTOProperty "attached text","OK"
'					JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("Potential problem found").JavaButton("OK").Click micLeftBtn
'				End if	
				
				''''''''''''''''''''''''''''''''''''''''''''''''''
				'Set Metric Unit Values
				'''''''''''''''''''''''''''''''''''''''''''''''''
				if Fn_SISW_UI_Object_Operations("Fn_ClassAdmin_CreateAttribute","Exist", ObjClassAdminApplet.JavaDialog("Potential problem found"),"") then 
					Call Fn_SISW_UI_JavaButton_Operations("Fn_ClassAdmin_CreateAttribute", "Click", ObjClassAdminApplet.JavaDialog("Potential problem found"),"OK")
				End if 
'				If JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("Potential problem found").Exist(5) then
'						JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("Potential problem found").JavaButton("OK").Click micLeftBtn
'				End if	
'
				If isArray ( aMetricFormat ) Then
						For iCount = 0 to ubound ( aMetricFormat )
							If aMetricFormat(iCount) <> ""  Then
						
									aTextSet = Split ( aMetricFormat(iCount), ":" , -1 , 1  )
									Select Case aTextSet(0)
			
											Case "Default Value"
                                                    	
													''''''''''''''''''''''''''''
													'For Text
													'''''''''''''''''''''''''''
		
													If aMetricFormat(1) <> "Key LOV"  and  aMetricFormat(1) <> "Date" Then
													
														If Fn_SISW_UI_Object_Operations("Fn_ClassAdmin_CreateAttribute","Exist", ObjClassAdminApplet.JavaEdit("AttributeDefaultValue"),"")  Then
																Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_ClassAdmin_CreateAttribute",ObjClassAdminApplet.JavaEdit("AttributeDefaultValue"),"attached text", aTextSet(0)+":") 
																Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_ClassAdmin_CreateAttribute",ObjClassAdminApplet.JavaEdit("AttributeDefaultValue"),"Index", "0") 
																if Fn_SISW_UI_JavaEdit_Operations("Fn_ClassAdmin_ClassOperations", "Set",  JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow"), "AttributeDefaultValue", aTextSet(1)) = false then			
																		Fn_ClassAdmin_CreateAttribute = False
																		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Set [ " + aTextSet(0) + "] = [ " + aTextSet(1) + "] " ) 
																		Exit Function 
																End If
																Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Set [ " + aTextSet(0) + "] = [ " + aTextSet(1) + "] " ) 		
														End If
'														If JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaEdit("AttributeDefaultValue").Exist Then
'																JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaEdit("AttributeDefaultValue").SetTOProperty "attached text",aTextSet(0)+":"
'																JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaEdit("AttributeDefaultValue").SetTOProperty "Index",0
'																JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaEdit("AttributeDefaultValue").Set aTextSet(1)
'																If Err.Number < 0 Then
'																		Fn_ClassAdmin_CreateAttribute = False
'																		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Set [ " + aTextSet(0) + "] = [ " + aTextSet(1) + "] " ) 
'																		Exit Function 
'																End If
'																Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Set [ " + aTextSet(0) + "] = [ " + aTextSet(1) + "] " ) 		
'														End If
				
													End If
		
													''''''''''''''''''''''
													'For Date
													'''''''''''''''''''''
													If aMetricFormat(1) = "Date" Then
													
															If Fn_SISW_UI_Object_Operations("Fn_ClassAdmin_CreateAttribute","Exist", ObjClassAdminApplet.JavaCheckBox("UnitDefaultDate"),"")  Then
																
																If Fn_UI_Object_SetTOProperty_ExistCheck("Fn_ClassAdmin_CreateAttribute",ObjClassAdminApplet.JavaCheckBox("UnitDefaultDate"),"attached text", aTextSet(1)) = false Then																															
																		Fn_ClassAdmin_CreateAttribute = False
																		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Set [ " + aTextSet(0) + "] = [ " + aTextSet(1) + "] " ) 
																		Exit Function 
																End If
																Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Set [ " + aTextSet(0) + "] = [ " + aTextSet(1) + "] " ) 		
														End If
'														If JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaCheckBox("UnitDefaultDate").Exist Then
'																JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaCheckBox("UnitDefaultDate").Object.setText aTextSet(1)
'																If Err.Number < 0 Then
'																		Fn_ClassAdmin_CreateAttribute = False
'																		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Set [ " + aTextSet(0) + "] = [ " + aTextSet(1) + "] " ) 
'																		Exit Function 
'																End If
'																Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Set [ " + aTextSet(0) + "] = [ " + aTextSet(1) + "] " ) 		
'														End If
'				
													End If
		
													''''''''''''''''''''''
													'For KeyLOV
													'''''''''''''''''''''	
													
													If  Fn_SISW_UI_Object_Operations("Fn_ClassAdmin_CreateAttribute","Exist", ObjClassAdminApplet.JavaList("AttributeDefaultValueList"),"")  Then
															if Fn_SISW_UI_JavaList_Operations("Fn_ClassAdmin_CreateAttribute", "Select", ObjClassAdminApplet, "AttributeDefaultValueList", aTextSet(1), "", "") = false then															
																	Fn_ClassAdmin_CreateAttribute = False
																	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Set [ " + aTextSet(0) + "] = [ " + aTextSet(1) + "] " ) 
																	Exit Function 
															End If
															Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Set [ " + aTextSet(0) + "] = [ " + aTextSet(1) + "] " ) 
													End If													
'													If  JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaList("AttributeDefaultValueList").Exist Then
'															JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaList("AttributeDefaultValueList").Select aTextSet(1)
'															If Err.Number < 0 Then
'																	Fn_ClassAdmin_CreateAttribute = False
'																	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Set [ " + aTextSet(0) + "] = [ " + aTextSet(1) + "] " ) 
'																	Exit Function 
'															End If
'															Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Set [ " + aTextSet(0) + "] = [ " + aTextSet(1) + "] " ) 
'													End If
		
											Case "Unit"
													Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_ClassAdmin_CreateAttribute",ObjClassAdminApplet.JavaObject("SelectUnit"),"attached text", "Metric Unit") 
													'JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaObject("SelectUnit").SetTOProperty "attached text", "Metric Unit"
													'wait 1
													'JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaStaticText("SelectUnitText").Click 0,0,"LEFT"
													Call Fn_UI_JavaStaticText_Click("Fn_ClassAdmin_CreateAttribute",ObjClassAdminApplet,"SelectUnitText", 1,1, "LEFT")
													
													'Call Fn_ReadyStatusSync(10)
													aTextSet(1) = replace(aTextSet(1),"~",":")
													arrSelectUnitMenu=Split(aTextSet(1),":")
													If Fn_SISW_UI_Object_Operations("Fn_ClassAdmin_CreateAttribute","Exist", ObjClassAdminApplet.JavaMenu("SelectUnitMenuScroller"),"")  = false Then
													'If not JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaMenu("SelectUnitMenuScroller").Exist(3) Then
															'wait(1)
															ObjClassAdminApplet.JavaStaticText("SelectUnitText").Click 1,1,"LEFT"
															Call Fn_ReadyStatusSync(2)
													End If
													'wait 1
													
													Select Case arrSelectUnitMenu(0)
														Case "Time"
															abs_x=JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaMenu("SelectUnitMenuScroller").GetROProperty("abs_x")
															abs_y=JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaMenu("SelectUnitMenuScroller").GetROProperty("abs_y")
															Set DeviceReplay = CreateObject("Mercury.DeviceReplay")
															DeviceReplay.MouseMove abs_x,abs_y
															wait 3
															Set DeviceReplay =Nothing
															bReturn = Fn_UI_JavaMenu_Select("Fn_ClassAdmin_CreateAttribute",JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow"),aTextSet(1))
														Case Else
															If Fn_SISW_UI_Object_Operations("Fn_ClassAdmin_CreateAttribute","Exist", JavaWindow("ClassAdminMainWin").JavaMenu("tagname:="&arrSelectUnitMenu(0)&"","index:=0","displayed:=1"),"")  Then																														
															'If JavaWindow("ClassAdminMainWin").JavaMenu("tagname:="&arrSelectUnitMenu(0)&"","index:=0","displayed:=1").Exist(3) Then
																'bReturn = Fn_UI_JavaMenu_Select("Fn_ClassAdmin_CreateAttribute",JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow"),aTextSet(1))
																JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaMenu("label:="&arrSelectUnitMenu(0)&"","index:=0", "displayed:=1").JavaMenu("label:="&arrSelectUnitMenu(1)&"","index:=0").Select
															
															ElseIf Fn_SISW_UI_Object_Operations("Fn_ClassAdmin_CreateAttribute","Exist", JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaMenu("tagname:="&arrSelectUnitMenu(0)&"","index:=0","displayed:=1"),"")  Then	
															'If JavaWindow("ClassAdminMainWin").JavaMenu("tagname:="&arrSelectUnitMenu(0)&"","index:=0","displayed:=1").Exist(3) Then
																'bReturn = Fn_UI_JavaMenu_Select("Fn_ClassAdmin_CreateAttribute",JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow"),aTextSet(1))
																JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaMenu("label:="&arrSelectUnitMenu(0)&"","index:=0", "displayed:=1").JavaMenu("label:="&arrSelectUnitMenu(1)&"","index:=0").Select
															
															Else
																abs_x=JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaMenu("SelectUnitMenuScroller").GetROProperty("abs_x")
																abs_y=JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaMenu("SelectUnitMenuScroller").GetROProperty("abs_y")
																Set DeviceReplay = CreateObject("Mercury.DeviceReplay")
																DeviceReplay.MouseMove abs_x,abs_y
																'wait 3
																Set DeviceReplay =Nothing
																bReturn = Fn_UI_JavaMenu_Select("Fn_ClassAdmin_CreateAttribute",JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow"),aTextSet(1))
															End If
													End Select

'													 bReturn = Fn_UI_JavaMenu_Select("Fn_ClassAdmin_CreateAttribute",JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow"),aTextSet(1))
													If bReturn = false Then
																	Fn_ClassAdmin_CreateAttribute = False
																	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Set [ " + aTextSet(0) + "] = [ " + aTextSet(1) + "] " ) 
																	Exit Function 
													End If
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Set [ " + aTextSet(0) + "] = [ " + aTextSet(1) + "] " ) 
													If Fn_SISW_UI_Object_Operations("Fn_ClassAdmin_CreateAttribute","Exist", ObjClassAdminApplet.JavaDialog("Potential problem found"),"")  Then													
													
													'If JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("Potential problem found").Exist(5) then
															JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("Potential problem found").JavaButton("OK").Click micLeftBtn
													end if
		
										  Case "Minimum Value"
													 bReturn = Fn_ClassAdmin_AtrributeOperations("SetValues", "", "Minimum Value:"+aTextSet(1), "", "","")
													 If bReturn = false Then
																	Fn_ClassAdmin_CreateAttribute = False
																	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Set [ " + aTextSet(0) + "] = [ " + aTextSet(1) + "] " ) 
																	Exit Function 
													End If
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Set [ " + aTextSet(0) + "] = [ " + aTextSet(1) + "] " ) 
													
										  Case "Maximum Value"
													bReturn = Fn_ClassAdmin_AtrributeOperations("SetValues", "", "Maximum Value:"+aTextSet(1), "", "","")
													 If bReturn = false Then
																	Fn_ClassAdmin_CreateAttribute = False
																	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Set [ " + aTextSet(0) + "] = [ " + aTextSet(1) + "] " ) 
																	Exit Function 
													End If
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Set [ " + aTextSet(0) + "] = [ " + aTextSet(1) + "] " )

										  Case "Dependency Configuration"
													'JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaList("DependencyConfiguration").Select aTextSet(1)
													if Fn_SISW_UI_JavaList_Operations("Fn_ClassAdmin_CreateAttribute", "Select", ObjClassAdminApplet, "DependencyConfiguration", aTextSet(1), "", "") = false then															
														'wait 1
													'If Err.Number < 0 Then
														Fn_ClassAdmin_CreateAttribute = False
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Set [ " + aTextSet(0) + "] = [ " + aTextSet(1) + "] " )
														Exit Function 
													End If
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Set [ " + aTextSet(0) + "] = [ " + aTextSet(1) + "] " )

									End Select

							End If

						Next
				End if

			'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
				'Set Value of Format - Non Metric
				'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
				If Fn_SISW_UI_Object_Operations("Fn_ClassAdmin_CreateAttribute","Exist", ObjClassAdminApplet.JavaDialog("Potential problem found"),"")  Then				
				'If JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("Potential problem found").Exist(3) then
						JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("Potential problem found").JavaButton("OK").Click micLeftBtn
				End if

				If isArray ( aNonMetricFormat ) Then
						bReturn =  Fn_ClassAdmin_FormatTypeSelect(aNonMetricFormat)
						If bReturn = True Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Set Format " ) 
						Else
							Fn_ClassAdmin_CreateAttribute = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Set Format " ) 
							Exit Function
						End If
				End If

                             If Fn_SISW_UI_Object_Operations("Fn_ClassAdmin_CreateAttribute","Exist", ObjClassAdminApplet.JavaDialog("Potential problem found"),"")  Then
				'If JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("Potential problem found").Exist(3) then
						JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("Potential problem found").JavaButton("OK").Click micLeftBtn
				End if

				''''''''''''''''''''''''''''''''''''''''''''''''''
				'Set Non Metric Metric Unit Values
				'''''''''''''''''''''''''''''''''''''''''''''''''
			
				If  ISArray(aNonMetricFormat) Then
					For iCount = 0 to ubound ( aNonMetricFormat )
						If aNonMetricFormat(iCount) <> ""  Then
					
								aTextSet = Split ( aNonMetricFormat(iCount), ":" , -1 , 1  )
								Select Case aTextSet(0)
		
										Case "Default Value"
	
												''''''''''''''''''''''''''''
												'For Text
												'''''''''''''''''''''''''''
	
												If aNonMetricFormat(1) <> "KeyLOV"  and  aNonMetricFormat(1) <> "Date" Then
				                                                                if Fn_SISW_UI_Object_Operations("Fn_ClassAdmin_CreateAttribute","Exist", ObjClassAdminApplet.JavaEdit("AttributeDefaultValue"),"") then
													'If JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaEdit("AttributeDefaultValue").Exist Then
															JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaEdit("AttributeDefaultValue").SetTOProperty "attached text",aTextSet(0)+":"
															JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaEdit("AttributeDefaultValue").SetTOProperty "Index",1
															JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaEdit("AttributeDefaultValue").Set aTextSet(1)
															If Err.Number < 0 Then
																	Fn_ClassAdmin_CreateAttribute = False
																	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Set [ " + aTextSet(0) + "] = [ " + aTextSet(1) + "] " ) 
																	Exit Function 
															End If
															Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Set [ " + aTextSet(0) + "] = [ " + aTextSet(1) + "] " ) 		
													End If
			
												End If
	
												''''''''''''''''''''''
												'For Date
												'''''''''''''''''''''
												If aNonMetricFormat(1) = "Date" Then
	                                                                                     if Fn_SISW_UI_Object_Operations("Fn_ClassAdmin_CreateAttribute","Exist", ObjClassAdminApplet.JavaCheckBox("UnitDefaultDate"),"") then
													'If JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaCheckBox("UnitDefaultDate").Exist Then
															JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaCheckBox("UnitDefaultDate").Object.setText aTextSet(1)
															If Err.Number < 0 Then
																	Fn_ClassAdmin_CreateAttribute = False
																	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Set [ " + aTextSet(0) + "] = [ " + aTextSet(1) + "] " ) 
																	Exit Function 
															End If
															Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Set [ " + aTextSet(0) + "] = [ " + aTextSet(1) + "] " ) 		
													End If
			
												End If
	
												''''''''''''''''''''''
												'For KeyLOV
												'''''''''''''''''''''
	                                                                             if Fn_SISW_UI_Object_Operations("Fn_ClassAdmin_CreateAttribute","Exist", ObjClassAdminApplet.JavaList("AttributeDefaultValueList"),"") then
												'If  JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaList("AttributeDefaultValueList").Exist Then
														JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaList("AttributeDefaultValueList").Select aTextSet(1)
														If Err.Number < 0 Then
																Fn_ClassAdmin_CreateAttribute = False
																Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Set [ " + aTextSet(0) + "] = [ " + aTextSet(1) + "] " ) 
																Exit Function 
														End If
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Set [ " + aTextSet(0) + "] = [ " + aTextSet(1) + "] " ) 
												End If
	
										Case "Unit"
												JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaObject("SelectUnit").SetTOProperty "index", "7"
												wait 2
												JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaObject("SelectUnit").SetTOProperty "attached text", "Maximum Value:"
												'wait 1
												JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaStaticText("SelectUnitText").Click 0,0,"LEFT"
												Call Fn_ReadyStatusSync(10)
												aTextSet(1) = replace(aTextSet(1),"~",":")
												arrSelectUnitMenu1=Split(aTextSet(1),":")
												If not JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaMenu("SelectUnitMenuScroller").Exist(3) Then
														wait(1)
														JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaStaticText("SelectUnitText").Click 0,0,"LEFT"
														Call Fn_ReadyStatusSync(2)
												End If
												                          													
												
												If JavaWindow("ClassAdminMainWin").JavaMenu("tagname:="&arrSelectUnitMenu1(0)&"","index:=0","displayed:=1").Exist(3) Then
													JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaMenu("label:="&arrSelectUnitMenu1(0)&"","index:=0", "displayed:=1").JavaMenu("label:="&arrSelectUnitMenu1(1)&"","index:=0").Select
													'bReturn = Fn_UI_JavaMenu_Select("Fn_ClassAdmin_CreateAttribute",JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow"),aTextSet(1))
												Else
													abs_x=JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaMenu("SelectUnitMenuScroller").GetROProperty("abs_x")
													abs_y=JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaMenu("SelectUnitMenuScroller").GetROProperty("abs_y")
													Set DeviceReplay = CreateObject("Mercury.DeviceReplay")
													DeviceReplay.MouseMove abs_x,abs_y
													wait 3
													Set DeviceReplay =Nothing
													bReturn = Fn_UI_JavaMenu_Select("Fn_ClassAdmin_CreateAttribute",JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow"),aTextSet(1))
												End If
'												 bReturn = Fn_UI_JavaMenu_Select("Fn_ClassAdmin_CreateAttribute",JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow"),aTextSet(1))
												If bReturn = false Then
																Fn_ClassAdmin_CreateAttribute = False
																Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Set [ " + aTextSet(0) + "] = [ " + aTextSet(1) + "] " ) 
																Exit Function 
												End If
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Set [ " + aTextSet(0) + "] = [ " + aTextSet(1) + "] " ) 
                                                                                     If Fn_SISW_UI_Object_Operations("Fn_ClassAdmin_CreateAttribute","Exist", ObjClassAdminApplet.JavaDialog("Potential problem found"),"")  Then
												'If JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("Potential problem found").Exist(5) then
														JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("Potential problem found").JavaButton("OK").Click micLeftBtn
												End if
	
									  Case "Minimum Value"
												 bReturn = Fn_ClassAdmin_AtrributeOperations("SetValues", "", "NonMetric_Minimum Value:"+aTextSet(1), "", "","")
												 If bReturn = false Then
																Fn_ClassAdmin_CreateAttribute = False
																Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Set [ " + aTextSet(0) + "] = [ " + aTextSet(1) + "] " ) 
																Exit Function 
												End If
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Set [ " + aTextSet(0) + "] = [ " + aTextSet(1) + "] " ) 
												
									  Case "Maximum Value"
												bReturn = Fn_ClassAdmin_AtrributeOperations("SetValues", "", "NonMetric_Maximum Value:"+aTextSet(1), "", "","")
												 If bReturn = false Then
																Fn_ClassAdmin_CreateAttribute = False
																Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Set [ " + aTextSet(0) + "] = [ " + aTextSet(1) + "] " ) 
																Exit Function 
												End If
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Set [ " + aTextSet(0) + "] = [ " + aTextSet(1) + "] " ) 
										
		
								End Select		
										
						End If
	
					Next
				End If

				'Set the Value for Optimize display if set
                For iCount = 0 to ubound(aDictionaryValue)
						aTextSet = Split ( aDictionaryValue(iCount), ":" , -1 , 1  )
						Select Case aTextSet(0)
								  Case "OptimizeDisplay"
										JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaCheckBox("OptimizeDisplay").Set aTextSet(1) 
										If Err.Number < 0 Then
												Fn_ClassAdmin_CreateAttribute = False
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Set [ " + aTextSet(0) + "] = [ " + aTextSet(1) + "] " ) 
												Exit Function 
										End If
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Set [ " + aTextSet(0) + "] = [ " + aTextSet(1) + "] " ) 
						End Select
				Next
                
		''''''''''''''''''''''''''''''''''''''''''''''''
		'for not saving attribute
		'''''''''''''''''''''''''''''''''''''''''''''''

		For iCount = 0 to ubound(aDictionaryValue)
				aTextSet = Split ( aDictionaryValue(iCount), ":" , -1 , 1  )
				Select Case aTextSet(0)
		
						Case "bSave"
								If lcase(aTextSet(1)) = "false" then

									''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
									'Get the Attribute ID as a Return value
									'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
									JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaEdit("AddAttributeEditBox").SetTOProperty "attached text" , "Attribute ID:"
									sAttributeID= JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaEdit("AddAttributeEditBox").GetROProperty("value")
							
									Fn_ClassAdmin_CreateAttribute =sAttributeID
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Attribute Created " ) 
									Exit Function

								End If
	
				End Select  		
		Next

				''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
				' Click on Save current Instance
				'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
				'wait 3
				bReturn = Fn_ClassAdmin_ToolbarOperations("Save")
				If bReturn = True Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Clicked on Save current Instance " ) 
				Else
							Fn_ClassAdmin_CreateAttribute = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Click on Save current Instance" ) 
							Exit Function
				End If
				if Fn_SISW_UI_Object_Operations("Fn_ClassAdmin_CreateAttribute","Exist", ObjClassAdminApplet.JavaDialog("Potential problem found"),"") then
			'If JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("Potential problem found").Exist(5) then
							JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("Potential problem found").JavaButton("Yes").Click micLeftBtn
			End if
		Else 
				Fn_ClassAdmin_CreateAttribute =FALSE
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to select [ Create a new Instance ] from Toolbar" )
				Exit Function
		End If

		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		'Get the Attribute ID as a Return value
		'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaEdit("AddAttributeEditBox").SetTOProperty "attached text" , "Attribute ID:"
		sAttributeID= JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaEdit("AddAttributeEditBox").GetROProperty("value")

		Fn_ClassAdmin_CreateAttribute =sAttributeID
        Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Attribute Created " ) 

End Function

'*********************************************************  Function for Select & Verify Items in CLass Attributes Tab's lists *********************************************************************

'Function Name		:					Fn_ClassAdmin_ClassAttributesLists

'Description			 :		 		    Action  performed :-
'												Select & Verify Items in CLass Attributes Tab's lists					
'																	
'																    

'Parameters			   :	 			 1. sAction -- Select /Verify
'												 2. sAttributeTypes : 'Class' is by default , Need to give mention for 'Inherited'
'												 3.aListItems : When  sAction = Select only 1 entry in the array 
'																	  sAction = Verify Then Multiple entries in Array Allowed 	
'                                                  
'Return Value		   : 			 True/False

'Pre-requisite			:		 	 Class Attributes Tab Selected  

'Examples				:			 aListItems = array("1004  Integer","1003  Real")
'												Call Fn_ClassAdmin_ClassAttributesLists("Exist","Inherited",aListItems)
'											aListItems = array("1004  Integer","1003  Real")	
'												Call Fn_ClassAdmin_ClassAttributesLists("Select","",aListItems)
'	                                              
'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Prasanna			14-Dec-2010	       1.0														      Prasanna B.
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function Fn_ClassAdmin_ClassAttributesLists(sAction,sAttributeTypes,aListItems)
	GBL_FAILED_FUNCTION_NAME="Fn_ClassAdmin_ClassAttributesLists"
	On error resume next

	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'  Variable Declaration
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim objList,iItemCount,iCounter,sTreeItem,sNode,iListCount

	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'  Set the Object for the list.
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	If trim(sAttributeTypes) = "Inherited" Then	
			Set objList = JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaList("InheritedAttributesList")
	Else
			Set objList = JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaList("ClassAttributesList")
			sAttributeTypes = "Class"
	End If
	Err.Clear
	If  objList.Exist(5)Then

			Select Case sAction

	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'Select the Entry from List
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			Case "Select"
						sNode = aListItems(0)
						objList.Select sNode
						Wait(2)
						If Err.Number < 0  Then
								Fn_ClassAdmin_ClassAttributesLists = False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to select [ " + sNode + "] of [ " +sAttributeTypes + "] Attributes List." ) 
								Set objList = Nothing
								Exit Function 
						Else								
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully selected  [ " + sNode + "] of [ " +sAttributeTypes + "] Attributes List.") 
						End If

	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'Cerify the List Entry .
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			Case "Exist"
						For iListCount = 0 to UBound(aListItems)
								sNode = aListItems(iListCount) 
								iItemCount = objList.GetROProperty( "items count")
								For iCounter=0 To (iItemCount-1)
										sTreeItem = objList.GetItem(iCounter)
										wait 2
											If Trim (Lcase(sTreeItem)) = Trim(Lcase(sNode)) Then													
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully found  [" + sNode + "] of [ " +sAttributeTypes + "] Attributes List." )	
													Exit For
											End If
								Next 
		
								If  Cint(iCounter) = Cint (iItemCount) Then
										Fn_ClassAdmin_ClassAttributesLists = False
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to  found  [" + sNode + "] of [ " +sAttributeTypes + "] Attributes List." )	
										Set objList = Nothing
										Exit Function 
								End If
						Next
			End Select
		End If

		Fn_ClassAdmin_ClassAttributesLists =  true
		Set objList = Nothing
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Performed Operation [" + sAction + "] on [ " +sAttributeTypes + "] Attributes List." )	
End Function

'*********************************************************  Function to perform Operations on Dictionary Search*********************************************************************

'Function Name		:	Fn_ClassAdmin_DictionarySearchOperations

'Description			 :	 Function to perform Operations on Dictionary Search
'																    

'Parameters			   :	1. sAction : Action  performed 
'							2. sSearchCriteria 
'							3. sTextToSearch
'							4. sValueToVerify -  for future use

'Return Value		   : 		True/False

'Pre-requisite			:	Classification  pane should be displayed.

'Examples				:	Call Fn_ClassAdmin_DictionarySearchOperations("Search", "Short Name", "*", "")
'

'History:
'			Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'			Koustubh W				15-Dec-2010	           1.0	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'			Koustubh W				17-Dec-2010	           1.0				Modified case Search
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'			Dipali K				12-Mar-2013	           1.1				Added Hierarchy : Window("TeamcenterWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Error")
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public function Fn_ClassAdmin_DictionarySearchOperations(sAction, sSearchCriteria, sTextToSearch, sValueToVerify)
	GBL_FAILED_FUNCTION_NAME="Fn_ClassAdmin_DictionarySearchOperations"
	Dim objApplet
	Fn_ClassAdmin_DictionarySearchOperations = False
	Set objApplet = JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow")
	' open dictionary search panel
	If objApplet.Exist(10) = False  Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_ClassAdmin_DictionarySearchOperations ] Class Admin Applet does not exist.")
		Set objApplet = nothing
		Exit function
	End If
	
	If objApplet.JavaTab("MainTab").GetROProperty ("value") <> "Dictionary" then
		objApplet.JavaTab("MainTab").Select "Dictionary"
	End If
	
	Select Case sAction
		Case "Search"
			If sSearchCriteria <> "" then
				objApplet.JavaCheckBox("SearchCriteria").Set "ON"
				wait 1
				If Fn_UI_ListItemExist("Fn_ClassAdmin_DictionarySearchOperations", objApplet,"SearchCriteria", sSearchCriteria) <> False then
					'selecting search criteria from criteria list
					objApplet.JavaList("SearchCriteria").Select sSearchCriteria
					wait 1
					If objApplet.JavaList("SearchCriteria").Exist(10) then		
						wait 5
						objApplet.JavaCheckBox("SearchCriteria").DblClick 1,1,"LEFT"	'Changed By Pritam Shikare	
					End if 
				Else
					'objApplet.JavaCheckBox("SearchCriteria").SetTOProperty "Index", 1
					objApplet.JavaCheckBox("SearchCriteria").DblClick 1,1,"LEFT"              'Changed By Pritam Shikare
					'objApplet.JavaCheckBox("SearchCriteria").SetTOProperty "Index", 0
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_ClassAdmin_DictionarySearchOperations ] Failed with invalid Search Criteria [ " & sSearchCriteria & " ].")
					Set objApplet = nothing
					Exit function
				End If
			End IF
			' setting text to search text box
			If sTextToSearch <> "" Then
				objApplet.JavaEdit("SearchText").Set sTextToSearch
			End If
			' clicking on search button
			objApplet.JavaButton("ICADictionary_Search").Click micLeftBtn
			Fn_ClassAdmin_DictionarySearchOperations = True

			JavaWindow("DefaultWindow").JavaWindow("ErrorJavaWindow").SetTOProperty "title","Search in Dictionary"
			JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("ErrorDialog").SetTOProperty "title","Search in Dictionary"
			Window("TeamcenterWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Error").SetTOProperty "title","Search in Dictionary"

			If JavaWindow("DefaultWindow").JavaWindow("ErrorJavaWindow").Exist(5) Then
				JavaWindow("DefaultWindow").JavaWindow("ErrorJavaWindow").JavaButton("OK").Click micLeftBtn
				Fn_ClassAdmin_DictionarySearchOperations = False
			Elseif JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("ErrorDialog").Exist(5) Then
				JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("ErrorDialog").JavaButton("OK").Click micLeftBtn
				Fn_ClassAdmin_DictionarySearchOperations = False
            Elseif  Window("TeamcenterWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Error").Exist(5) Then
				Window("TeamcenterWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Error").JavaButton("OK").Click micLeftBtn
				Fn_ClassAdmin_DictionarySearchOperations = False
			End If

			wait(3)
			If JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaButton("ButtonSearchClose").Exist(5) then
				JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaButton("ButtonSearchClose").Click micLeftBtn
			ElseIf JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaObject("SearchCloseButton").Exist(5) Then
				JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaObject("SearchCloseButton").Click 0,0,"LEFT"
			End If
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_ClassAdmin_DictionarySearchOperations ] Failed with invalid  case [ " & sAction & " ].")
			Set objApplet = nothing
			Exit function
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	End Select
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: [ Fn_ClassAdmin_DictionarySearchOperations ] Executed successfully with case [ " & sAction & " ].")
	Set objApplet = nothing
End Function

'*********************************************************  Function to perform Operations on ICA Dictionary Table *********************************************************************

'Function Name		:	Fn_ClassAdmin_ICADictionaryTableOperations

'Description			 :	 Function to perform Operations on Dictionary Search result i.e ICA Dictionary Table
'																    

'Parameters			   :	1. sAction : Action  performed 
'						2. sAttributeID 
'						3. sColName
'						4. sValueToVerify -  for future use

'Return Value		   : 		True/False

'Pre-requisite			:	Classification  pane should be displayed.

'Examples				:	'Call Fn_ClassAdmin_ICADictionaryTableOperations("CellVerify", "1002", "Name", "String")
'						Call Fn_ClassAdmin_ICADictionaryTableOperations("RowSelect", "1005", "", "")
'						Call Fn_ClassAdmin_ICADictionaryTableOperations("CellDoubleClick", "-2824", "Short Name", "")

'History:
'			Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'			Koustubh W				15-Dec-2010	           1.0	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public function Fn_ClassAdmin_ICADictionaryTableOperations(sAction, sAttributeID, sColName, sValueToVerify)
	GBL_FAILED_FUNCTION_NAME="Fn_ClassAdmin_ICADictionaryTableOperations"
	Dim objApplet, iRows, iRowCounter
	Fn_ClassAdmin_ICADictionaryTableOperations = False
	Set objApplet = JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow")
	
	If objApplet.JavaTab("MainTab").GetROProperty ("value") <> "Dictionary" then
		objApplet.JavaTab("MainTab").Select "Dictionary"
	End If
	
	' verifying ICA Dictionary Table
	If objApplet.JavaTable("ICADictionaryTable").Exist(10) = False  Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_ClassAdmin_ICADictionaryTableOperations ] ICA Dictionary Table does not exist.")
		Set objApplet = nothing
		Exit function
	End If
	
	Select Case sAction
		Case "CellVerify"
			'clicking on Load All if button is visible
			If  cInt(objApplet.JavaButton("ICADictionary_LoadAll").GetROProperty ("enabled")) = 1 Then
				objApplet.JavaButton("ICADictionary_LoadAll").Click micLeftBtn
				Call Fn_ReadyStatusSync(5) 
			End If
			iRows = cInt(objApplet.JavaTable("ICADictionaryTable").GetROProperty ("rows"))
			For iRowCounter = 0 to iRows - 1
				If cstr(objApplet.JavaTable("ICADictionaryTable").GetCellData(iRowCounter, "Attribute ID")) = sAttributeID then
					If cstr(objApplet.JavaTable("ICADictionaryTable").GetCellData(iRowCounter,sColName)) = sValueToVerify then
						Fn_ClassAdmin_ICADictionaryTableOperations = True
						Exit for
					End If
				End IF
			Next
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "RowSelect"
			'clicking on Load All if button is visible
			If  cInt(objApplet.JavaButton("ICADictionary_LoadAll").GetROProperty ("enabled")) = 1 Then
				objApplet.JavaButton("ICADictionary_LoadAll").Click micLeftBtn
				wait 10
			End If
			iRows = cInt(objApplet.JavaTable("ICADictionaryTable").GetROProperty ("rows"))
			For iRowCounter = 0 to iRows - 1
				If cstr(objApplet.JavaTable("ICADictionaryTable").GetCellData(iRowCounter, "Attribute ID")) = sAttributeID then
					objApplet.JavaTable("ICADictionaryTable").SelectRow iRowCounter
					Fn_ClassAdmin_ICADictionaryTableOperations = True
					Exit for
				End IF
			Next
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "CellDoubleClick"
			'clicking on Load All if button is visible
			If  cInt(objApplet.JavaButton("ICADictionary_LoadAll").GetROProperty ("enabled")) = 1 Then
				objApplet.JavaButton("ICADictionary_LoadAll").Click micLeftBtn
				wait 10
			End If
			iRows = cInt(objApplet.JavaTable("ICADictionaryTable").GetROProperty ("rows"))
			For iRowCounter = 0 to iRows - 1
				If cstr(objApplet.JavaTable("ICADictionaryTable").GetCellData(iRowCounter, "Attribute ID")) = sAttributeID then
					If sColName = "" Then sColName = "Attribute ID"
					objApplet.JavaTable("ICADictionaryTable").DoubleClickCell iRowCounter, sColName
					Fn_ClassAdmin_ICADictionaryTableOperations = True
					Exit for
				End IF
			Next
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_ClassAdmin_ICADictionaryTableOperations ] Failed with invalid  case [ " & sAction & " ].")
			Set objApplet = nothing
			Exit function
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	End Select
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: [ Fn_ClassAdmin_ICADictionaryTableOperations ] Executed successfully with case [ " & sAction & " ].")
	Set objApplet = nothing
End Function

'********************************************************* Function to perform operations on Key LOV Tree ***********************************************************************

'Function Name		          :      Fn_ClassAdmin_KeyLOVTreeOperations

'Description			    :	 Function to perform operations on Key LOV Tree.

'Parameters			   :	1. sAction - Java Tree Object 
'						2. sNode - Node to select
'								We should use seperator character '~' in case if node text contain ':'
'						3. sMenu - for future use, not yet implemented
											
'Return Value		         : 	True / False

'Pre-requisite			:	Key LOV Tree should be visible.

'Examples				:         'sNode =  "-995:?????~???:????? @3"
							         'msgbox Fn_ClassAdmin_KeyLOVTreeOperations("SelectWithTilda", sNode, "")
							         'msgbox Fn_ClassAdmin_KeyLOVTreeOperations("ExpandWithTilda", sNode, "")
							         'msgbox Fn_ClassAdmin_KeyLOVTreeOperations("ExistWithTilda", sNode, "")
'History:
'			Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'			Koustubh 				16-Dec-2010			   1.0
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public function Fn_ClassAdmin_KeyLOVTreeOperations(sAction, sNode, sMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_ClassAdmin_KeyLOVTreeOperations"
	Dim objLOVTree, iRows, iRowCounter, iInstance
	Set objLOVTree = JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaTree("LOVKeyTree")
	Fn_ClassAdmin_KeyLOVTreeOperations = False
	Select Case sAction
		Case "Select", "SelectWithTilda"
			'according to sepearator retrieving node index
			Select Case sAction
				Case "Select"
					iRowCounter = Fn_JavaTree_NodeIndexExt("Fn_ClassAdmin_KeyLOVTreeOperations",JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow"),"LOVKeyTree",  sNode, "", "")
				Case "SelectWithTilda"
					iRowCounter = Fn_JavaTree_NodeIndexExt("Fn_ClassAdmin_KeyLOVTreeOperations",JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow"),"LOVKeyTree",  sNode, "~", "")
			End Select
			wait 1
			If iRowCounter <> -1 then
				objLOVTree.Object.setSelectionRow iRowCounter
				Fn_ClassAdmin_KeyLOVTreeOperations = True
			End If
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Expand","ExpandWithTilda"
			'according to sepearator retrieving node index
			Select Case sAction
				Case "Expand"
					iRowCounter = Fn_JavaTree_NodeIndexExt("Fn_ClassAdmin_KeyLOVTreeOperations",JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow"),"LOVKeyTree",  sNode, "", "")
				Case "ExpandWithTilda"
					iRowCounter = Fn_JavaTree_NodeIndexExt("Fn_ClassAdmin_KeyLOVTreeOperations",JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow"),"LOVKeyTree",  sNode, "~", "")
			End Select
			If iRowCounter <> -1 then
				' expanding tree node by getting TreePath
				objLOVTree.Object.setSelectionRow iRowCounter
				objLOVTree.Object.setExpandedState objLOVTree.Object.getSelectionPath(), true
				Fn_ClassAdmin_KeyLOVTreeOperations = True
			End If
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Exist", "ExistWithTilda"
			'according to sepearator retrieving node index
			Select Case sAction
				Case "Exist"
					iRowCounter = Fn_JavaTree_NodeIndexExt("Fn_ClassAdmin_KeyLOVTreeOperations",JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow"),"LOVKeyTree",  sNode, "", "")
				Case "ExistWithTilda"
					iRowCounter = Fn_JavaTree_NodeIndexExt("Fn_ClassAdmin_KeyLOVTreeOperations",JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow"),"LOVKeyTree",  sNode, "~", "")
			End Select
			If iRowCounter <> -1 then
				Fn_ClassAdmin_KeyLOVTreeOperations = True
			End If
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "PopupMenuSelect"
			' for future use
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_ClassAdmin_KeyLOVTreeOperations ] Invalid case [ " & sAction & " ].")
			Set objLOVTree = nothing
			Exit function
	End Select
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: [ Fn_ClassAdmin_KeyLOVTreeOperations ] Executed successfully with case [ " & sAction & " ].")
	Set objLOVTree = nothing
End Function
'********************************************************* Function to perform operations on Key LOV ***********************************************************************

'Function Name		          :      Fn_ClassAdmin_KeyLOVOperations

'Description			    :	 Function to perform operations on Key LOV.

'Parameters			   :	1. sAction - Action to be performed
'						2. sKeyLOVID - to create new Key LOV Id
'						3. EntryValue - for future use, not yet implemented
'						4. bHideKeys - Boolean value of Hide Keys True or False
'						5. sKeyLOVDefinitions - ~ separated path of KeyLOVTree : eg. -8001:Value1~key1:Value1
'						6. sKeyValues - ~ separated pair of key and value eg.  Key1:Value1~key2:Value2
											
'Return Value		         : 	True / False

'Pre-requisite			:	Key LOV Tree should be visible.

'Examples				:         
							'msgbox Fn_ClassAdmin_KeyLOVOperations("Create", "-8001", "Lov2", "", "", "key1:Value1")
							'msgbox Fn_ClassAdmin_KeyLOVOperations("ActivateEditMode", "-8001  Lov2", "", "False", "", "")
							'msgbox Fn_ClassAdmin_KeyLOVOperations("SetValues", "", "8001", "True", "", "key1:Value1")
							'msgbox Fn_ClassAdmin_KeyLOVOperations("Add Entry", "", "", "", "-8001:Value1~key1:Value1", "key2:Value2")
							'msgbox Fn_ClassAdmin_KeyLOVOperations("Insert Entry", "", "", "", "-8001:Value1~key1:Value1", "key3:Value3")
							'msgbox Fn_ClassAdmin_KeyLOVOperations("Add Submenu", "", "", "", "-8001:Value1~key1:Value1", "Value1")
							'msgbox Fn_ClassAdmin_KeyLOVOperations("Insert Submenu", "", "", "", "-8001:Value1~key1:Value1", "")
							'msgbox Fn_ClassAdmin_KeyLOVOperations("Add Separator", "", "", "", "-8001:Value1~key1:Value1", "")
							'msgbox Fn_ClassAdmin_KeyLOVOperations("Insert Separator", "", "", "", "-8001:Value1~key1:Value1", "")
							'msgbox Fn_ClassAdmin_KeyLOVOperations("Delete", "", "", "", "-8001:Value1~key2:Value2", "")
							'msgbox Fn_ClassAdmin_KeyLOVOperations("Preview", "", "", "", "", "3 y")
							'msgbox Fn_ClassAdmin_KeyLOVOperations("VerifyPreview", "", "", "", "", "1 x~2 y~muru2:5 n")
							'msgbox Fn_ClassAdmin_KeyLOVOperations("Deprecate", "", "", "", "-8001:Value1~key1:Value1", "")
							'msgbox Fn_ClassAdmin_KeyLOVOperations("Deprecate", "", "", "", "", "")
							'msgbox Fn_ClassAdmin_KeyLOVOperations("Undeprecate", "", "", "", "-8001:Value1~key1:Value1", "")
							'msgbox Fn_ClassAdmin_KeyLOVOperations("Undeprecate", "", "", "", "", "")
'History:
'			Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'			Koustubh 				16-Dec-2010			   1.0
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'			Koustubh 				17-Dec-2010			   1.0				Added cases Preview adn VerifyPreview
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'			Koustubh 				18-Dec-2010			   1.0				Added cases Deprecate and Undeprecate
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public function Fn_ClassAdmin_KeyLOVOperations(sAction, sKeyLOVID, EntryValue, bHideKeys, sKeyLOVDefinitions, sKeyValues)
	GBL_FAILED_FUNCTION_NAME="Fn_ClassAdmin_KeyLOVOperations"
	Dim objNewKeyLov, objApplet, iCounter, aKeyValues, aPair, sNode, bReturn
	'Set objApplet = JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow")
	Set objApplet = Fn_SISW_ClassAdmin_GetObject("ClassAdminApplet")
	Fn_ClassAdmin_KeyLOVOperations = False 
	Select Case sAction
		Case "Create"
			if Fn_ToolbatButtonClick("Create a new Instance") = false then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_ClassAdmin_KeyLOVOperations ] Failed to open new Key LOV Definition.")
				Set objApplet = nothing
				Exit function
			end if 

			Set objNewKeyLov = Fn_UI_ObjectCreate("Fn_ClassAdmin_KeyLOVOperations", JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("New Key-LOV"))
			' open specified Keyh LOV
			If objNewKeyLov.Exist(10) then
				Call Fn_Edit_Box("Fn_ClassAdmin_KeyLOVOperations",objNewKeyLov,"Key-LOV ID", sKeyLOVID)
                Call Fn_Button_Click("Fn_ClassAdmin_KeyLOVOperations",objNewKeyLov, "OK")
			End IF
			Set objNewKeyLov = nothing
			If  EntryValue <> "" Then
				Call Fn_Edit_Box("Fn_ClassAdmin_KeyLOVOperations",objApplet,"LOVEntryValue", EntryValue)
				wait(2)
				objApplet.JavaEdit("LOVEntryValue").Activate
			End If
			If bHideKeys <> "" Then
				If cBool(bHideKeys) = True Then
					Call Fn_CheckBox_Set("Fn_ClassAdmin_KeyLOVOperations", objApplet, "Hide Keys", "ON")
				Else
					Call Fn_CheckBox_Set("Fn_ClassAdmin_KeyLOVOperations", objApplet, "Hide Keys", "OFF")
				End If
			End If
			If EntryValue = "" Then
				sNode =  sKeyLOVID & ":?????"
			Else
				sNode =  sKeyLOVID & ":" & EntryValue
			End If
			
			aKeyValues = split(sKeyValues,"~")
			For iCounter = 0 to UBound (aKeyValues)
				wait 1
				aPair =  Split(aKeyValues(iCounter),":")
				If  Fn_ClassAdmin_KeyLOVTreeOperations("SelectWithTilda", sNode &":???:????? @" & (iCounter + 1), "") = false then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_ClassAdmin_KeyLOVOperations ] Failed to select row from KeyLOV Tree.")
					Set objApplet = nothing
					Exit function
				end if 
				Call Fn_Edit_Box("Fn_ClassAdmin_KeyLOVOperations",objApplet,"LOVKeyID", aPair(0))
				wait(1)
				objApplet.JavaEdit("LOVKeyID").Activate
				wait(1)
				Call Fn_Edit_Box("Fn_ClassAdmin_KeyLOVOperations",objApplet,"LOVEntryValue", aPair(1))
				wait(1)
				objApplet.JavaEdit("LOVEntryValue").Activate

			Next
			wait(1)
			if Fn_ClassAdmin_ToolbarOperations("Save")  = false then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_ClassAdmin_KeyLOVOperations ] Failed to save created KeyLOV.")
				Set objApplet = nothing
				Exit function
			end if 
			Fn_ClassAdmin_KeyLOVOperations = True
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "ActivateEditMode"
				objApplet.JavaList("ExistingLOVList").Select sKeyLOVID
				wait(2)
				' click on edit
				if Fn_ClassAdmin_ToolbarOperations("Edit") = false then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_ClassAdmin_KeyLOVOperations ] Failed to Activate Edit mode.")
					Set objApplet = nothing
					Exit function
				end if 
				Fn_ClassAdmin_KeyLOVOperations = True
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "SetValues"
			If sKeyLOVDefinitions <> "" Then
				if Fn_ClassAdmin_KeyLOVTreeOperations("SelectWithTilda", sKeyLOVDefinitions, "") = false then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_ClassAdmin_KeyLOVOperations ] Failed to select row from KeyLOV Tree.")
					Set objApplet = nothing
					Exit function
				end if 
			End If
			'setting default Entry value
			If  EntryValue <> "" Then
				objApplet.JavaEdit("LOVEntryValue").Set EntryValue
				wait(2)
				objApplet.JavaEdit("LOVEntryValue").Activate
			End If
			If bHideKeys <> "" Then
				If cBool(bHideKeys) = True Then
					objApplet.JavaCheckBox("Hide Keys").Set "ON"
				Else
					objApplet.JavaCheckBox("Hide Keys").Set "OFF"
				End If
			End If
			aPair =  Split(sKeyValues,":")
			If trim(aPair(0)) <> "" Then
				objApplet.JavaEdit("LOVKeyID").Set trim(aPair(0))
				wait(2)
				objApplet.JavaEdit("LOVKeyID").Activate
			End If
			If trim(aPair(1)) <> ""Then
				objApplet.JavaEdit("LOVEntryValue").Set trim(aPair(1))
				wait(2)
				objApplet.JavaEdit("LOVEntryValue").Activate
			End If
			Fn_ClassAdmin_KeyLOVOperations = True
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Add Entry", "Insert Entry", "Add Submenu", "Insert Submenu", "Add Separator", "Insert Separator", "Delete"
			If sKeyLOVDefinitions <> "" Then
				if Fn_ClassAdmin_KeyLOVTreeOperations("SelectWithTilda", sKeyLOVDefinitions, "") = false then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_ClassAdmin_KeyLOVOperations ] Failed to select row from KeyLOV Tree.")
					Set objApplet = nothing
					Exit function
				end if 
			End If
			if cint(objApplet.JavaButton(sAction).GetROProperty ("enabled")) = 1 then
				objApplet.JavaButton(sAction).Click micLeftBtn
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_ClassAdmin_KeyLOVOperations ] Button [ " & sAction & " ] is disable.")
				Set objApplet = nothing
				Exit function
			End if 
			If sKeyValues <> "" Then
				Select Case sAction
					Case  "Add Submenu", "Insert Submenu"
						'Added From TC1017 -- TC1123-20161130-12_12_2016-SandipC-Maintenance- Added to set value LOVKeyID:LOVEntryValue after clicking Add Submenu button
						If Instr(sKeyValues,":") > 0 Then
							aPair =  Split(sKeyValues,":")
							If trim(aPair(0)) <> "" Then
								objApplet.JavaEdit("LOVKeyID").Set trim(aPair(0))
								wait(2)
								objApplet.JavaEdit("LOVKeyID").Activate
							End If
							If trim(aPair(1)) <> "" Then
								objApplet.JavaEdit("LOVEntryValue").Set trim(aPair(1))
								wait(2)
								objApplet.JavaEdit("LOVEntryValue").Activate
							End If
							'----------------------------------------------------------------------------------------------------------------------------							
						ElseIf trim(sKeyValues) <> ""Then
							objApplet.JavaEdit("LOVEntryValue").Set trim(sKeyValues)
							wait(2)
							objApplet.JavaEdit("LOVEntryValue").Activate
						End If

					Case Else
						aPair =  Split(sKeyValues,":")
						If trim(aPair(0)) <> "" Then
							objApplet.JavaEdit("LOVKeyID").Set trim(aPair(0))
							wait(2)
							objApplet.JavaEdit("LOVKeyID").Activate
						End If
						If trim(aPair(1)) <> ""Then
							objApplet.JavaEdit("LOVEntryValue").Set trim(aPair(1))
							wait(2)
							objApplet.JavaEdit("LOVEntryValue").Activate
						End If
				End Select
			End If
			Fn_ClassAdmin_KeyLOVOperations = True
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Preview"
			objApplet.JavaButton("Preview").Click micLeftBtn
			Fn_ClassAdmin_KeyLOVOperations = Fn_UI_JavaMenu_Select("Fn_ClassAdmin_KeyLOVOperations",objApplet,trim(sKeyValues))
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "VerifyPreview"
			Dim intX, intY
			aPair = split(sKeyValues,"~")
			For iCounter = 0 to UBound (aPair)
'				JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaButton("Preview").Object.setFocusable True
'				wait(1)
'				JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaButton("Preview").Click micLeftBtn
'				objApplet.JavaButton("Preview").Click micLeftBtn
				'objApplet.JavaButton("Preview").Object.doClick
				intX = JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaButton("Preview").GetROProperty("abs_x")
				intY = JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaButton("Preview").GetROProperty("abs_y")
				JavaWindow("ClassAdminMainWin").Click intX+20,intY+10,"LEFT"
				wait(3)
				'opening menu
				If Fn_UI_JavaMenu_Exist("Fn_ClassAdmin_KeyLOVOperations",objApplet,trim(aPair(iCounter))) = False Then
					Fn_ClassAdmin_KeyLOVOperations = False
					Exit for 
				Else
					Fn_ClassAdmin_KeyLOVOperations = True
				End If
				'closing menu
				JavaWindow("ClassAdminMainWin").Click intX+20,intY+10,"LEFT"
'				objApplet.JavaButton("Preview").Click micLeftBtn
'				JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaButton("Preview").Click micLeftBtn
				'objApplet.JavaButton("Preview").Object.doClick
			Next

			'JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaButton("Preview").FireEvent(micMouseClick,1)

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Deprecate", "Undeprecate"
			If sKeyLOVDefinitions <> "" Then
				if Fn_ClassAdmin_KeyLOVTreeOperations("SelectWithTilda", sKeyLOVDefinitions, "") = false then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_ClassAdmin_KeyLOVOperations ] Failed to select row from KeyLOV Tree.")
					Set objApplet = nothing
					Exit function
				end if 
			End If
			objApplet.JavaCheckBox("Deprecate").SetTOProperty "attached text", sAction
			If objApplet.JavaCheckBox("Deprecate").exist(5) = True  Then
				If cInt(objApplet.JavaCheckBox("Deprecate").GetROProperty("enabled")) = 1 Then
					objApplet.JavaCheckBox("Deprecate").Set "ON"
					wait(2)
					Fn_ClassAdmin_KeyLOVOperations = True
				End If
			End If			
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_ClassAdmin_KeyLOVOperations ] Invalid case [ " & sAction & " ].")
			Set objApplet = nothing
			Exit function
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	End Select
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: [ Fn_ClassAdmin_KeyLOVOperations ] Executed successfully with case [ " & sAction & " ].")
	Set objApplet = nothing
End Function

'********************************************************* Function to return different date formats ***********************************************************************

'Function Name		          :      Fn_ClassAdmin_DateFormat

'Description			    :	 Function to return different date formats.

'Parameters			   :	1. sAction - Format type
'						2. sNode - Date from which need to be in format	
											
'Return Value		         : 	Date format / False

'Examples				:        call Fn_ClassAdmin_DateFormat("DDMMYY",now)
							        
'History:
'			Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'			Prasanna 				17-Dec-2010			   1.0
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public function Fn_ClassAdmin_DateFormat(sType,sDate)

Dim sDay,sMonth,sYear
Fn_ClassAdmin_DateFormat = false

		Select Case sType
					Case "DD.MM.YY"
							sDay = day(sDate)
							If sDay < 10  Then
								sDay  = "0"+cstr(sDay)
							End If
							sMonth = month(sDate)	
							If sMonth < 10  Then
								sMonth  = "0"+cstr(sMonth)
							End If
							sYear = year(sDate)	
		
							sYear = right(sYear,2)
							Fn_ClassAdmin_DateFormat = cstr(sDay) +"."+cstr(sMonth)+"."+cstr(sYear)
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Date format is ["+cstr(sDay) +"."+cstr(sMonth)+"."+cstr(sYear)+"]")
		End Select

End Function

'********************************************************* Function to return different date formats ***********************************************************************

'Function Name		          :      Fn_ClassAdmin_DeleteObjects

'Description			    :	 Function to delete the objects.

'Parameters			   :	1. sInfo : To be used in feature
'						
											
'Return Value		         : 	true/ False

'Prerequisite 		         : 	Object need to be selected

'Examples				:        call Fn_ClassAdmin_DeleteObjects("")
							        
'History:
'			Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'			Prasanna 				18-Dec-2010			   1.0
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function Fn_ClassAdmin_DeleteObjects(sInfo)
		'GO to EDIT mode
		GBL_FAILED_FUNCTION_NAME="Fn_ClassAdmin_DeleteObjects"
		Dim bReturn, oConDlg,bFlag
		Err.Clear
		bReturn =  Fn_ClassAdmin_ToolbarOperations("Edit")
		If bReturn = false Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Perform [Edit] Opeartion.")
					Fn_ClassAdmin_DeleteObjects = false
					Exit Function 	
		Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully Performed [Edit] Opeartion.")
                    Call Fn_ReadyStatusSync(3)
					wait(3)					
		End If

		'Delete the selected object
		bReturn =  Fn_ClassAdmin_ToolbarOperations("Delete")
		If bReturn = false Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Perform [Delete] Opeartion.")
					Fn_ClassAdmin_DeleteObjects = false
					Exit Function 	
		Else
					wait(3)
                    Call Fn_ReadyStatusSync(1)
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully Performed [Delete] Opeartion.")					
		End If
		Err.Clear
		bFlag = False
		'Check the existance of  'Confirmation' Dialog
		If  JavaDialog("ConfirmationMessage").Exist(5) Then
			Set oConDlg = JavaDialog("ConfirmationMessage")
		'ElseIf Window("ClassificationWindow").JavaDialog("DeleteConfirmation").Exist(5) Then
		'	Set oConDlg = Window("ClassificationWindow").JavaDialog("DeleteConfirmation").JavaDialog("DeleteConfirmation")
		ElseIf JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("DeleteConfirmation").Exist(5) Then
			Set oConDlg = JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("DeleteConfirmation")
		ElseIf JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaDialog("DeleteConfirmation").Exist(5) Then
			Set oConDlg = JavaWindow("ClassificationMainWin").JavaWindow("ClassificationSubWindow").JavaDialog("DeleteConfirmation")
		ElseIf JavaDialog("Delete Preference(s)").Exist(5) Then
			Set oConDlg = JavaDialog("Delete Preference(s)")
		ElseIf JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("RemoveStorageClass").Exist(5)Then
			bFlag = True
		ElseIf Fn_UI_Object_SetTOProperty_ExistCheck("Fn_ClassAdmin_DeleteObjects",JavaDialog("Delete"),"title","Delete attribute") Then
			Set oConDlg = JavaDialog("Delete")
		Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Confirmation Dialog Does not exist.")												
			Fn_ClassAdmin_DeleteObjects = False
			Set oConDlg = Nothing
			Exit Function
		End If
		If bFlag = False Then
			If oConDlg.Exist(5) = True Then
							oConDlg.Activate
							 If Err.Number < 0 Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Confirmation Dialog Does not exist.")												
									Fn_ClassAdmin_DeleteObjects = False
									Set oConDlg = Nothing
									Exit Function
							Else
									wait(3)
									'Vallari - Ready is not Ready as Delete COnfirmation dialog in ON
									'Call Fn_ReadyStatusSync(3)
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully verified the existence of Confirmation Dialog.")								
							End If
	
							
							'Click on 'Yes' button
							oConDlg.JavaButton("Yes").Click micLeftBtn
							If Err.Number < 0 Then
										Fn_ClassAdmin_ClassOperations = False
										oConDlg.JavaButton("No").Click micLeftBtn
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed toClick on Yes button" )
										Set oConDlg = Nothing
										Exit Function 
							Else							
										wait(3)
										Call Fn_ReadyStatusSync(5)
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Clicked on Yes button")
							End If  
	
			End If
			Set oConDlg = Nothing
		End If

			'Check the existance of  'RemoveStorageClass' Dialog
		If JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("RemoveStorageClass").Exist(5)  Then
							JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("RemoveStorageClass").JavaCheckBox("RemoveHierachy").Set "ON"
							JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("RemoveStorageClass").JavaButton("Yes").Click micLeftBtn
							If Err.Number < 0 Then		
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : To Click on Yes Button.")									
										Fn_ClassAdmin_DeleteObjects = False
										JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("RemoveStorageClass").Close
										Exit function
						Else
									wait 2
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Clicked on Yes Button.")								
						End If			
		End If

		Fn_ClassAdmin_DeleteObjects = True							
		
End Function

'********************************************************* Function to verify different objects Classification admin***********************************************************************

'Function Name		          :      Fn_ClassAdmin_VerifyClassDetails

'Description			    :	 Function to delete the objects.

'Parameters			   :	1. sObject : object which need to be verified
'							2. sValue : value to be verified
'							3. sDetails : other info	
										
'Return Value		         : 	true/ False

'Examples				:        call Fn_ClassAdmin_VerifyClassDetails("DeprecateButtonLabels","Undeprecate","")
'										call Fn_ClassAdmin_VerifyClassDetails("VerifyPropertiesCheckbox","Auto Computed:1","")
'							        Call Fn_ClassAdmin_VerifyClassDetails("VerifyClassAttrListValues","123 abcde","")
'										bReturn=Fn_ClassAdmin_VerifyClassDetails("Modify&Verify","de:portugese","de:turkish")
'
'History:
'			Developer Name						Date				Rev. No.			Changes Done			Reviewer	
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'			Prasanna 								18-Dec-2010			   1.0
'           Mahendra Bhandarkar		07-Jan-2010
'			SHREYAS								22-04-2011				1.2
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function Fn_ClassAdmin_VerifyClassDetails(sObject,sValue,sDetails)
	GBL_FAILED_FUNCTION_NAME="Fn_ClassAdmin_VerifyClassDetails"

	Dim aValues,iCounter,objApplet,sCheckValue,sCheckValText ,aProperties,bFlag,sItemDetails,sItems
	Set objApplet = JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow")
	Select Case sObject
				Case "AliasName"
							If INSTR(1,sValue,",") Then
									aValues = split(sValue,",",-1,1)
							End If
	
							For iCounter = 0 to UBound(aValues)								   
								   If Fn_UI_ListItemExist("", objApplet,"AliasNamesSet", aValues(iCounter)) <> False then
											Fn_ClassAdmin_VerifyClassDetails = true
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Successfully Verified Value [" + aValues(iCounter) + "] in [ Alias Names ] List" ) 
									Else
											Fn_ClassAdmin_VerifyClassDetails = false
											objApplet = Nothing
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Verify Value [" + aValues(iCounter) + "] Exists in [ Alias Names ] List" ) 									
											Exit Function 
								   End if	
							Next

					Case "DepricateButtonLabels"
		
							 	objApplet.JavaCheckBox("Deprecate").SetTOProperty "attached text", sValue
								 If objApplet.JavaCheckBox("Deprecate").Exist(5) then
											Fn_ClassAdmin_VerifyClassDetails = true
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Successfully Verified Value [" + sValue + "]" ) 
								Else
											Fn_ClassAdmin_VerifyClassDetails = false
											objApplet = Nothing
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Verify Value [" + sValue + "] " ) 									
											Exit Function 
								End if		

					Case "AliasNameList"
								  If INSTR(1,sValue,",") Then
											aValues = split(sValue,",",-1,1)
								  End If						
								  For iCounter = 0 to UBound(aValues)			
								  
										 objApplet.JavaList("AliasNameList").SetTOProperty "Index",  0
										 If Fn_UI_ListItemExist("", objApplet,"AliasNameList", aValues(iCounter)) <> False then
													Fn_ClassAdmin_VerifyClassDetails = true
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Successfully Verified Value [" + aValues(iCounter) + "] in [ Alias Names ] Drop Down List" ) 
										 Else                                                                                                                                                                                                                                                      
													Fn_ClassAdmin_VerifyClassDetails = false
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Verify Value [" + aValues(iCounter) + "] Exists in [ Alias Names ] Drop Down List" ) 
													 Exit Function                                                                                                                                            
										   End if  
								  Next    					   

						'Case to verify Class Atrributes Prperties Checkbox values	
					   Case "VerifyPropertiesCheckbox"
								'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
								If INSTR(1,sValue,"Application 1") Then
											Call Fn_ClassAdmin_TabOpeartions("Activate","ImageTab","Applicability/User Data","")
											wait 1, 500								
											aValues = split(sValue,":",-1,1)
											 'Get the value from UI 
											JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaCheckBox("ChkProperties").SetTOProperty "attached text","<html> Marks the attribute as important for Application 1 </html>"				
											sCheckValue = JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaCheckBox("ChkProperties").GetROProperty("value")
		
											If cstr(sCheckValText) = "1" Then
													sCheckValText = "Checked"
											Else 
													sCheckValText = "UnCheck"
											End If
											If  trim(sCheckValue) = trim(aValues(1))Then
													Fn_ClassAdmin_VerifyClassDetails = true
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Verified Value [" + sCheckValText + "] for Checkbox ["+aValues(0)+"]" ) 
											Else
													Fn_ClassAdmin_VerifyClassDetails = false
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Successfully Verified Value [" + sCheckValText + "] for Checkbox ["+aValues(0)+"]" ) 
											End If

								ElseIf INSTR(1,sValue,"Application 2") Then
											Call Fn_ClassAdmin_TabOpeartions("Activate","ImageTab","Applicability/User Data","")
											wait 1, 500								
											aValues = split(sValue,":",-1,1)
											 'Get the value from UI 
											JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaCheckBox("ChkProperties").SetTOProperty "attached text","<html> Marks the attribute as important for Application 2 </html>"
											sCheckValue = JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaCheckBox("ChkProperties").GetROProperty("value")
		
											If cstr(sCheckValText) = "1" Then
													sCheckValText = "Checked"
											Else 
													sCheckValText = "UnCheck"
											End If
											If  trim(sCheckValue) = trim(aValues(1))Then
													Fn_ClassAdmin_VerifyClassDetails = true
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Verified Value [" + sCheckValText + "] for Checkbox ["+aValues(0)+"]" ) 
											Else
													Fn_ClassAdmin_VerifyClassDetails = false
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Successfully Verified Value [" + sCheckValText + "] for Checkbox ["+aValues(0)+"]" ) 
											End If

								ElseIf INSTR(1,sValue,"Application 3") Then
											Call Fn_ClassAdmin_TabOpeartions("Activate","ImageTab","Applicability/User Data","")
											wait 1, 500
											aValues = split(sValue,":",-1,1)
											 'Get the value from UI 
											JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaCheckBox("ChkProperties").SetTOProperty "attached text","<html> Marks the attribute as important for Application 3 </html>"
											sCheckValue = JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaCheckBox("ChkProperties").GetROProperty("value")
		
											If cstr(sCheckValText) = "1" Then
													sCheckValText = "Checked"
											Else 
													sCheckValText = "UnCheck"
											End If
											If  trim(sCheckValue) = trim(aValues(1))Then
													Fn_ClassAdmin_VerifyClassDetails = true
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Verified Value [" + sCheckValText + "] for Checkbox ["+aValues(0)+"]" ) 
											Else
													Fn_ClassAdmin_VerifyClassDetails = false
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Successfully Verified Value [" + sCheckValText + "] for Checkbox ["+aValues(0)+"]" ) 
											End If
								ElseIf INSTR(1,sValue,"Application 4") Then
											Call Fn_ClassAdmin_TabOpeartions("Activate","ImageTab","Applicability/User Data","")
											wait 1, 500
											aValues = split(sValue,":",-1,1)
											 'Get the value from UI 
											JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaCheckBox("ChkProperties").SetTOProperty "attached text","<html> Marks the attribute as important for Application 4 </html>"			
											sCheckValue = JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaCheckBox("ChkProperties").GetROProperty("value")
		
											If cstr(sCheckValText) = "1" Then
													sCheckValText = "Checked"
											Else 
													sCheckValText = "UnCheck"
											End If
											If  trim(sCheckValue) = trim(aValues(1))Then
													Fn_ClassAdmin_VerifyClassDetails = true
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Verified Value [" + sCheckValText + "] for Checkbox ["+aValues(0)+"]" ) 
											Else
													Fn_ClassAdmin_VerifyClassDetails = false
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Successfully Verified Value [" + sCheckValText + "] for Checkbox ["+aValues(0)+"]" ) 
											End If			
											
								ElseIf INSTR(1,sValue,"Application 5") Then
											Call Fn_ClassAdmin_TabOpeartions("Activate","ImageTab","Applicability/User Data","")
											wait 1, 500
											aValues = split(sValue,":",-1,1)
											 'Get the value from UI 
											JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaCheckBox("ChkProperties").SetTOProperty "attached text","<html> Marks the attribute as important for Application 5 </html>"			
											sCheckValue = JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaCheckBox("ChkProperties").GetROProperty("value")
		
											If cstr(sCheckValText) = "1" Then
													sCheckValText = "Checked"
											Else 
													sCheckValText = "UnCheck"
											End If
											If  trim(sCheckValue) = trim(aValues(1))Then
													Fn_ClassAdmin_VerifyClassDetails = true
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Verified Value [" + sCheckValText + "] for Checkbox ["+aValues(0)+"]" ) 
											Else
													Fn_ClassAdmin_VerifyClassDetails = false
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Successfully Verified Value [" + sCheckValText + "] for Checkbox ["+aValues(0)+"]" ) 
											End If	
											
							Else
							'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
										If INSTR(1,sValue,":") Then
											aValues = split(sValue,":",-1,1)
										 End If
										 'Get the value from UI 
										JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaCheckBox("ChkProperties").SetTOProperty "attached text",trim(aValues(0))						
										sCheckValue = JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaCheckBox("ChkProperties").GetROProperty("value")
		
										If cstr(sCheckValText) = "1" Then
												sCheckValText = "Checked"
										Else 
												sCheckValText = "UnCheck"
										End If
										If  trim(sCheckValue) = trim(aValues(1))Then
													Fn_ClassAdmin_VerifyClassDetails = true
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Verified Value [" + sCheckValText + "] for Checkbox ["+aValues(0)+"]" ) 
										Else
													Fn_ClassAdmin_VerifyClassDetails = false
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Successfully Verified Value [" + sCheckValText + "] for Checkbox ["+aValues(0)+"]" ) 
										End If		
							End If	

							'=================================Verify Optimize Display CheckBox================================================================
						Case "VerifyOptimizeDisplayCheckbox"
								If INSTR(1,sValue,":") Then
									aValues = split(sValue,":",-1,1)
								 End If
								 'Get the value from UI 
	   							sCheckValue = JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaCheckBox("OptimizeDisplay").GetROProperty("value")

								If cstr(sCheckValText) = "1" Then
										sCheckValText = "Checked"
								Else 
										sCheckValText = "UnCheck"
								End If
								If  trim(sCheckValue) = trim(aValues(1))Then
											Fn_ClassAdmin_VerifyClassDetails = true
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Verified Value [" + sCheckValText + "] for Checkbox ["+aValues(0)+"]" ) 
								Else
											Fn_ClassAdmin_VerifyClassDetails = false
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Successfully Verified Value [" + sCheckValText + "] for Checkbox ["+aValues(0)+"]" ) 
								End If		

					'Case to verify Class Atrributes Prperties Checkbox values			
					   Case "VerifyMeasurmentCheckbox"
								If INSTR(1,sValue,":") Then
									aValues = split(sValue,":",-1,1)
								 End If
								 'Get the value from UI 
								JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaRadioButton("Measurement").SetTOProperty "attached text",trim(aValues(0))						
	   							sCheckValue = JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaRadioButton("Measurement").GetROProperty("value")

								If cstr(sCheckValue) = "1" Then
										sCheckValText = "Checked"
								Else 
										sCheckValText = "UnCheck"
								End If
								If  trim(sCheckValue) = trim(aValues(1))Then
											Fn_ClassAdmin_VerifyClassDetails = true
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Verified Value [" + sCheckValText + "] for Checkbox ["+aValues(0)+"]" ) 
								Else
											Fn_ClassAdmin_VerifyClassDetails = false
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Successfully Verified Value [" + sCheckValText + "] for Checkbox ["+aValues(0)+"]" ) 
								End If		
					' Checked the Value Exist on Default value or not.
						Case "VerifyClassAttrListValues"
								' Get the value from UI
								JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaCheckBox("Default Value").Set "ON"
								Wait 2
								JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaList("AttributeDefaultValueList").SetTOProperty "attached text", "Default Value"								
								bReturn = Fn_UI_ListItemExist("Fn_ClassAdmin_VerifyClassDetails", JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow"), "AttributeDefaultValueList",sValue)
								If bReturn = False Then
										Fn_ClassAdmin_VerifyClassDetails = false
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : To Verify that Default Value List Contains [" + sValue + "]." ) 
								Else
										Fn_ClassAdmin_VerifyClassDetails = true
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Verified that Default Value List [" + sValue + "]." )
								End If
					' Checked the Value Exist on Default value or not.
						Case "VerifyDicListValues"
							' Get the value from UI
									JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaList("AttributeDefaultValueList").SetTOProperty "attached text", "Default Value:"
									If JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaList("AttributeDefaultValueList").Exist(5) then
										intCnt=JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaList("AttributeDefaultValueList").GetROProperty("items count")
										For i=0 to intCnt-1
												sItem=JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaList("AttributeDefaultValueList").GetItem(i)
													If  sItem <> "" Then
														If Trim(cStr(sItem))=Trim(cStr(sValue)) Then
															Fn_ClassAdmin_VerifyClassDetails = true
															Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Verified that Default Value List [" + sValue + "]." )						
															Exit for 
														Else
															Fn_ClassAdmin_VerifyClassDetails = false
															Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : To Verify that Default Value List Contains [" + sValue + "]." ) 
														End If
													End If
											Next
									Else
										Fn_ClassAdmin_VerifyClassDetails = false	
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : To Verify that Default Value List Contains [" + sValue + "]." ) 
									End if

					Case "Modify&Verify"
						bFlag=false

						aValues=split(sValue,":",-1,1)

'select the value from en list
						If aValues(0)<>"" Then
							JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaList("AliasNameList").Select  aValues(0)
						End If

						'add the value
						If aValues(1)<>"" Then
							JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaEdit("AliasName").Set  aValues(1)
						End If

'click on add button
JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaButton("AddAlias").Click micLeftBtn

'Now the modification will happen..

'							select the value from the en list

							If aValues(0)<>"" Then
								JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaList("AliasNameList").Select  aValues(0)						
							End If

							'select the value from the Alias names list which is previously added and modify its name
							If sDetails<>"" Then
								aProperties=split(sDetails,":",-1,1)
								JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaList("AliasNamesSet").Select aValues(1)
								
								'now modify the name in the alias name edit box
								JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaEdit("AliasName").set aProperties(1)

								'click on the modify button
								JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaButton("ModifyAlias").Click micLeftBtn

							End If


							'now verify if the value is saved after modifying
							sItems=JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaList("AliasNamesSet").GetROProperty ("items count")
							For iCounter=0 to sItems-1
								sItemDetails=JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaList("AliasNamesSet").GetItem(iCounter)
								If lCase(aProperties(1))=lCase(sItemDetails) Then
										bFlag=True
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Verified that value is present after modifying it")
								End If
							Next


If bFlag=True Then
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully implemented the function Fn_ClassAdmin_VerifyClassDetails")
	Fn_ClassAdmin_VerifyClassDetails = true
Else
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail : Failed to implement the function Fn_ClassAdmin_VerifyClassDetails")
	Fn_ClassAdmin_VerifyClassDetails = False
End If


	End Select
End Function
 
'********************************************************* Function to perform operations on Class Attributes ***********************************************************************

'Function Name		          :      Fn_ClassAdmin_ClassAttributeOperations

'Description			    :	 Function to perform operations on Class Attributes

'Parameters			   :	1. sAction : Action to perform
'						2. bActivateEditMode : Boolean value to activate Edit mode
'						3. sAttributeType : Class attribute list type
'						4. sClassAttribute : Class attribute to be selected and edited
'						5. sFieldValues : ~ separated list of sets fields:value 
'						6. sValueToVerify : value to be verified ( for future use )
'						7. bSave : boolean value to perform save operation
'						8. bActivateOnLastEntry : Boolean value to perform Activate on last entry
										
'Return Value		         : 	True / False

'Examples				:        call Fn_ClassAdmin_ClassAttributeOperations("SetValues","", "", "", "NonMetric_Minimum Value:True~Default Value:123~Local Value:False", "", "","")
'								 call Fn_ClassAdmin_ClassAttributeOperations("SetValues","", "", "", "Minimum Value From Dictionary:True~NonMetric_Minimum Value From Dictionary:False", "", "","True")
							        
'History:
'			Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'			Koustubh 				  21-Dec-2010			   1.0
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'			Koustubh 				  22-Dec-2010			   1.0			Added code to handle From Dictionary checkboxes
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public function Fn_ClassAdmin_ClassAttributeOperations(sAction, bActivateEditMode, sAttributeType, sClassAttribute, sFieldValues, sValueToVerify, bSave, bActivateOnLastEntry )
	GBL_FAILED_FUNCTION_NAME="Fn_ClassAdmin_ClassAttributeOperations"
	Dim aPairs, aFieldValue, iCounter, iCnt, objCheckProp, objPropEdit,sProperties, sApplicability
	Fn_ClassAdmin_ClassAttributeOperations = False
          If Fn_ClassAdmin_TabOpeartions("Activate","Subtab","Class Attributes","") = False then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_ClassAdmin_ClassAttributeOperations ] Failed to activate Class Attributes tab.")
		Exit function
	End IF
	if bActivateOnLastEntry = "" then bActivateOnLastEntry = False
	'activating Edit mode
	If bActivateEditMode <> "" Then
		If cBool(bActivateEditMode)  Then
			If Fn_ClassAdmin_ToolbarOperations("Edit") = False then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_ClassAdmin_ClassAttributeOperations ] Failed click on edit tool bar button.")
				Exit function
			End If
		End If
	End If
	'if class attribute is spefied...
	If sClassAttribute <> "" Then
		Select Case sAttributeType
			Case "Inherited"
				JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaList("InheritedAttributesList").Select sClassAttribute
			Case "", "Class"
				JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaList("ClassAttributesList").Select sClassAttribute
			Case Else
                                        Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_ClassAdmin_ClassAttributeOperations ]Invalid type of Attribute list.")
				Exit function
		End Select
	End If
	' setting values 
	Select Case sAction
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		' SetValues
		Case "SetValues"
			aPairs = split(sFieldValues, "~")
			For iCounter = 0 to UBound(aPairs)
				aFieldValue = split(aPairs(iCounter),":")
				Select Case UBound(aFieldValue)
					Case 0
						' do nothing... skip
					' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
					Case 1
						' field value pair
						' for Metric non Metric
						Select Case trim(aFieldValue(0))
							Case "Config"
							'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
							Case "NonMetric_Default Value From Dictionary",  "NonMetric_Minimum Value From Dictionary",  "NonMetric_Maximum Value From Dictionary", "Default Value From Dictionary",  "Minimum Value From Dictionary",  "Maximum Value From Dictionary"
							    Call Fn_ClassAdmin_TabOpeartions("Activate","ImageTab","Min/Max/Default Value","")
								wait 1, 500
								Set objCheckProp = JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaCheckBox("ChkProperties")
								objCheckProp.SetTOProperty "attached text",  "From dictionary"
								Select Case  trim(aFieldValue(0))
									Case "Default Value From Dictionary"
										objCheckProp.SetTOProperty "Index", 0
									Case "Minimum Value From Dictionary"
										objCheckProp.SetTOProperty "Index", 1
									Case "Maximum Value From Dictionary"
										objCheckProp.SetTOProperty "Index", 2
									Case "NonMetric_Default Value From Dictionary"
										objCheckProp.SetTOProperty "Index", 3
									Case "NonMetric_Minimum Value From Dictionary"
										objCheckProp.SetTOProperty "Index", 4
									Case "NonMetric_Maximum Value From Dictionary"
										objCheckProp.SetTOProperty "Index", 5
								End Select
								If objCheckProp.Exist(5) Then
									If cInt(objCheckProp.GetROProperty("enabled")) = 1 Then
										If cBool(trim(aFieldValue(1))) Then
											objCheckProp.set  "ON"
										Else
											objCheckProp.set   "OFF"
										End If
									Else
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_ClassAdmin_ClassAttributeOperations ] Object [ " & trim(aFieldValue(0)) & " ] is not enabled.")
										Set objCheckProp = nothing
										Set objPropEdit = nothing
										Exit function
									End If
								Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_ClassAdmin_ClassAttributeOperations ] Object [ " & trim(aFieldValue(0)) & " ] does not exist.")
									Set objCheckProp = nothing
									Set objPropEdit = nothing
									Exit function
								End If
								Set objCheckProp = nothing
							'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
							Case "NonMetric_Default Value",  "NonMetric_Minimum Value",  "NonMetric_Maximum Value", "Default Value",  "Minimum Value",  "Maximum Value"
								Call Fn_ClassAdmin_TabOpeartions("Activate","ImageTab","Min/Max/Default Value","")
								wait 1, 500
								Set objPropEdit = JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaEdit("PropertyEditBox")
								If instr(aFieldValue(0),"NonMetric_") > 0 Then
									aFieldValue(0) = replace(trim(aFieldValue(0)),"NonMetric_","")
									Set objCheckProp = JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaCheckBox(trim(aFieldValue(0)))
									objCheckProp.SetTOProperty "Index",  1
									objPropEdit.SetTOProperty "Index", 1
								Else
									Set objCheckProp = JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaCheckBox(trim(aFieldValue(0)))
									objCheckProp.SetTOProperty "Index", 0
									objPropEdit.SetTOProperty "Index", 0	
								End If
										
								If objCheckProp.Exist(5) Then
									If cInt(objCheckProp.GetROProperty("enabled")) = 1 Then
										If lcase(trim(aFieldValue(1))) = "true" OR lcase(trim(aFieldValue(1))) = "false" Then
											If cBool(trim(aFieldValue(1))) Then
												objCheckProp.set  "ON"
											Else
												objCheckProp.set   "OFF"
											End If
										Else
											objCheckProp.set  "ON"
											objPropEdit.SetTOProperty "attached text", trim(aFieldValue(0))
											wait 1
											objPropEdit.Set trim(aFieldValue(1))
											If iCounter = UBound(aPairs) AND cBool(bActivateOnLastEntry) = True then
												wait 1
												objPropEdit.Activate
											End If
										End If
									Else
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_ClassAdmin_ClassAttributeOperations ] Object [ " & trim(aFieldValue(0)) & " ] is not enabled.")
										Set objCheckProp = nothing
										Set objPropEdit = nothing
										Exit function
									End If
								Else
									Set objCheckProp = nothing
									Set objPropEdit = nothing
									Exit function
								End If
								Set objCheckProp = nothing
								Set objPropEdit = nothing
							' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
							Case Else
								' for remaining checkboxes and edit boxes.
								sProperties = "Mandatory-Local Value-Unique-Hidden-Protected-Auto Computed"
								sApplicability = "NX CAM-Graphics Creation-GCS Connection-Application 4-Application 5"
								If instr(1, sProperties , aFieldValue(0)) Then
									Call Fn_ClassAdmin_TabOpeartions("Activate","ImageTab","Properties","")
								ElseIf instr(1, sApplicability , aFieldValue(0)) Then
									Call Fn_ClassAdmin_TabOpeartions("Activate","ImageTab","Applicability/User Data","")	
								Else
									Call Fn_ClassAdmin_TabOpeartions("Activate","ImageTab","Min/Max/Default Value","")
								End If
								wait 1, 500
								Set objCheckProp = JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaCheckBox("ChkProperties")
								Set objPropEdit = JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaEdit("PropertyEditBox")
								
								
								If aFieldValue(0) = "NX CAM" Then
									aFieldValue(0) = "<html> Marks the attribute as important for Application 1 </html>"
								ElseIf aFieldValue(0) = "Graphics Creation" Then
									aFieldValue(0) = "<html> Marks the attribute as important for Application 2 </html>"
								ElseIf aFieldValue(0) = "GCS Connection" Then
									aFieldValue(0) = "<html> Marks the attribute as important for Application 3 </html>"
								ElseIf aFieldValue(0) = "Application 4" Then
									aFieldValue(0) = "<html> Marks the attribute as important for Application 4 </html>"
								ElseIf aFieldValue(0) = "Application 5" Then
									aFieldValue(0) = "<html> Marks the attribute as important for Application 5 </html>"
								End If
								
								
								objCheckProp.SetTOProperty "attached text",  trim(aFieldValue(0))
								objPropEdit.SetTOProperty "attached text", trim(aFieldValue(0)) & ":"
								objCheckProp.SetTOProperty "Index", 0
								objPropEdit.SetTOProperty "Index", 0
										
								If objCheckProp.Exist(5) Then
									If cInt(objCheckProp.GetROProperty("enabled")) = 1 Then
										If cInt(objCheckProp.GetROProperty("enabled")) = 1 Then
											' setting checkbox on off
											If lcase(trim(aFieldValue(1))) = "true" OR lcase(trim(aFieldValue(1))) = "false" Then
												If cBool(trim(aFieldValue(1))) Then
													objCheckProp.set  "ON"
												Else
													objCheckProp.set   "OFF"
												End If
											End If
										Else
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_ClassAdmin_ClassAttributeOperations ] Object [ " & trim(aFieldValue(0)) & " ] is not enabled.")
											Set objCheckProp = nothing
											Set objPropEdit = nothing
											Exit function
										End If									
									End If
								Elseif objPropEdit.Exist(5) then ' setting values to edit box
										objPropEdit.SetTOProperty "Index", 0
										objPropEdit.Set trim(aFieldValue(1))
								Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_ClassAdmin_ClassAttributeOperations ] Object [ " & trim(aFieldValue(0)) & " ] does not exist.")
									Set objCheckProp = nothing
									Set objPropEdit = nothing
									Exit function
								End If
								Set objCheckProp = nothing
								Set objPropEdit = nothing
						End Select
					' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
					Case else
						' field value pair with multiple values
						' this case is for future use.
				End Select
			Next
			Fn_ClassAdmin_ClassAttributeOperations = True
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Invalid case
		Case else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_ClassAdmin_ClassAttributeOperations ] Invalid case [ " & sAction & " ].")
			Exit function
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	End Select
	' performing save operation
	If bSave <> "" Then
		If cBool(bSave) Then
			If Fn_ClassAdmin_ToolbarOperations("Save") = False then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_ClassAdmin_ClassAttributeOperations ] Failed click on Save tool bar button.")
				Fn_ClassAdmin_ClassAttributeOperations = False
				Exit function
			End If
		End If
	End If
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: [ Fn_ClassAdmin_ClassAttributeOperations ] Executed successfully with case [ " & sAction & " ].")
 End Function

'********************************************************* Function to handle error dialogs ***********************************************************************

'Function Name		          :      Fn_ClassAdmin_ErrorHandler

'Description			    :	 Function to handle error dialogs

'Parameters			   :	1. sAction : Action to perform
'						2. sTitle : Title to set
'						3. sMsg : Message to verify
										
'Return Value		         : 	True / False

'Examples				:        call  Fn_ClassAdmin_ErrorHandler("ErrorDialog", "Constraints Errors", "The default value should be within the range specified by the minimum and maximum values.")
							        
'History:
'			Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'			Koustubh 				  21-Dec-2010			   1.0
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'			Koustubh 				  27-Mar-2010			   1.0
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
public function Fn_ClassAdmin_ErrorHandler(sAction, sTitle, sMsg)
	GBL_FAILED_FUNCTION_NAME="Fn_ClassAdmin_ErrorHandler"
	Dim objErrorWindow, objErrorWindow1, objErrorWindow2
	Fn_ClassAdmin_ErrorHandler = False
	Select Case sAction
		Case "ErrorDialog"
				Set objErrorWindow = JavaDialog("Error")
				Set objErrorWindow1 = JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("ErrorDialog")
				Set objErrorWindow2 = JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("ErrorDialog")
	
				If sTitle <> ""  Then
					objErrorWindow.SetTOProperty "title", trim(sTitle)
					objErrorWindow1.SetTOProperty "title", trim(sTitle)
					objErrorWindow2.SetTOProperty "title", trim(sTitle)
				End If
				' different error dialog
				If objErrorWindow.Exist(5) Then
					Fn_ClassAdmin_ErrorHandler = True
					If sMsg <> "" Then
						If instr(objErrorWindow.JavaStaticText("Msg").GetROProperty ("label"), sMsg) < 0 then
							Fn_ClassAdmin_ErrorHandler = false
						End if
					End If
					objErrorWindow.JavaButton("OK").SetTOProperty "Index",0
					objErrorWindow.JavaButton("OK").SetTOProperty "displayed", "1"	
					objErrorWindow.JavaButton("OK").Click micLeftBtn
				' different error dialog
				ElseIf objErrorWindow1.Exist(5) Then
					Fn_ClassAdmin_ErrorHandler = False
					If sMsg <> "" Then
						If instr (1,trim(objErrorWindow1.JavaEdit("Msg").GetROProperty("value")),trim(sMsg),1) > 0Then
							Fn_ClassAdmin_ErrorHandler = True
						End If
					End If
					objErrorWindow1.JavaButton("OK").Click micLeftBtn
				ElseIf objErrorWindow2.Exist(5) Then
					Fn_ClassAdmin_ErrorHandler = True
					If sMsg <> "" Then
						If trim(objErrorWindow2.JavaEdit("Msg").GetROProperty("value")) <> sMsg Then
							Fn_ClassAdmin_ErrorHandler = False
						End If
					End If
					objErrorWindow2.JavaButton("OK").Click micLeftBtn
				End If
				Set objErrorWindow = nothing
				Set objErrorWindow1 = nothing
				Set objErrorWindow2 = nothing
		Case "ErrorWindow"
				Set objErrorWindow = JavaWindow("DefaultWindow").JavaWindow("ErrorJavaWindow")
				If sTitle <> ""  Then
					objErrorWindow.SetTOProperty "title", trim(sTitle)
				End If
				If sMsg <> "" Then
						objErrorWindow.JavaStaticText("Details").SetTOProperty "label", sMsg
						If objErrorWindow.Exist = False Then
							Fn_ClassAdmin_ErrorHandler = False
						End If
				End If
				If objErrorWindow.Exist(10) Then
					objErrorWindow.JavaButton("OK").Click micLeftBtn
					Fn_ClassAdmin_ErrorHandler = True
				End If
				Set objErrorWindow = nothing

	End Select
End function

'********************************************************* Function to perform operations on Attributes ***********************************************************************

'Function Name		          :      Fn_ClassAdmin_AtrributeOperations

'Description			    :	 Function to perform operations on Attributes

'Parameters			   :	1. sAction : Action to perform
'						2. bActivateEditMode : Boolean value to activate Edit mode
'						3. sFieldValues : ~ separated list of sets fields:value 
'						6. sValueToVerify : value to be verified ( for future use )
'						7. bSave : boolean value to perform save operation
'						8. bActivateOnLastEntry : Boolean value to perform Activate on last entry
										
'Return Value		         : 	True / False

'Examples				:        call Fn_ClassAdmin_AtrributeOperations("SetValues", "", "Default Value:15~NonMetric_Default Value:asd", "", "","True")
'							     call Fn_ClassAdmin_AtrributeOperations("Verify", "", "Default Value:15~NonMetric_Default Value:asd", "", "","")
							        
'History:
'			Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'			Koustubh 				  22-Dec-2010			   1.0
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'			Koustubh 				  29-Dec-2010			   1.0			Added case Verify
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'			Sushma 				      18-Jan-2012			   1.0			Modified case Verify(Maintenance)
'___________________________________________________________________________________________________________________________________________________
''			Dipali 						31 July 2012              				Modified case Verify (Porting 10.0)			
'___________________________________________________________________________________________________________________________________________________
''			Sandeep N 						02 Apr 2013              				Modified case Verify ---> 1 (Porting 10.1)			
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public function Fn_ClassAdmin_AtrributeOperations(sAction, bActivateEditMode, sFieldValuePair, sValueToVerify, bSave, bActivateOnLastEntry)
	GBL_FAILED_FUNCTION_NAME="Fn_ClassAdmin_AtrributeOperations"
	Dim objEditBox, iCounter, iCnt, aPair, aFieldValues
	Fn_ClassAdmin_AtrributeOperations = False
	if bActivateOnLastEntry = "" then bActivateOnLastEntry = False
       	'activating Edit mode
	If bActivateEditMode <> "" Then
		If cBool(bActivateEditMode)  Then
			If Fn_ClassAdmin_ToolbarOperations("Edit") = False then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_ClassAdmin_AtrributeOperations ] Failed click on edit tool bar button.")
				Exit function
			End If
		End If
	End If
	Select Case sAction
' - - - - - - -  - - - - - - -  - - - - - - -  - - - - - - -  - - - - - - -  - - - - - - -  - - - - - - -  - - - - - - -  - - - - - - -  - - - - - - -  - - - - - - -  - - - - - - - 
		Case "SetValues"
			If sFieldValuePair <> "" Then
				aFieldValues = split(sFieldValuePair,"~")
				For iCounter = 0 to UBound(aFieldValues)
					aPair = Split(aFieldValues(iCounter),":")
					Select Case UBound(aPair)
						' - - - - - - -  - - - - - - -  - - - - - - -  
						Case 0
							' do nothing... skip
						' - - - - - - -  - - - - - - -  - - - - - - -  
						Case 1
							' field and its value
							Set objEditBox = JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaEdit("AttributeDefaultValue")
							If instr(trim(aPair(0)), "NonMetric_") > 0 then
								aPair(0) = replace(trim(aPair(0)), "NonMetric_","")
								objEditBox.SetTOProperty "Index", 1 
							Else
								objEditBox.SetTOProperty "Index", 0 
                            End if
							objEditBox.SetTOProperty "attached text", trim(aPair(0)) & ":"
							If objEditBox.Exist(10) Then
								If cInt(objEditBox.GetROProperty("enabled")) = 1 Then
									objEditBox.set trim(aPair(1))
									if iCounter = UBound(aFieldValues) AND cBool(bActivateOnLastEntry) = True then
										objEditBox.Activate
									End If
								Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_ClassAdmin_AtrributeOperations ] Object [ " & trim(aPair(0)) & " ] is not enabled.")
									Set objEditBox = nothing
									Exit function
								End If
							Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_ClassAdmin_AtrributeOperations ] Object [ " & trim(aPair(0)) & " ] does not exist.")
								Set objEditBox = nothing
								Exit function
							End If
							Set objEditBox = nothing
						' - - - - - - -  - - - - - - -  - - - - - - -  
						Case Else
							' for future use
					End Select
				Next
				Fn_ClassAdmin_AtrributeOperations = True
			End If
' - - - - - - -  - - - - - - -  - - - - - - -  - - - - - - -  - - - - - - -  - - - - - - -  - - - - - - -  - - - - - - -  - - - - - - -  - - - - - - -  - - - - - - -  - - - - - - - 
		Case "Verify"
			If sFieldValuePair <> "" Then
				Call Fn_ClassAdmin_TabOpeartions("Activate","ImageTab","Min/Max/Default Value","")
				wait 1	
				If Fn_ClassAdmin_TabOpeartions("VerifyActivate","ImageTab","Min/Max/Default Value","") = False Then
					Call Fn_ClassAdmin_TabOpeartions("Activate","ImageTab","Min/Max/Default Value","")
				End If
				aFieldValues = split(sFieldValuePair,"~")
				For iCounter = 0 to UBound(aFieldValues)
					aPair = Split(aFieldValues(iCounter),":")
					Select Case UBound(aPair)
						' - - - - - - -  - - - - - - -  - - - - - - -  
						Case 0
							' do nothing... skip
							' invalid input
							Set objEditBox = nothing
							Exit function
						' - - - - - - -  - - - - - - -  - - - - - - -  
						Case 1
							' field and its value
							Set objEditBox = JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaEdit("AttributeDefaultValue")
							If instr(trim(aPair(0)), "NonMetric_") > 0 then
								aPair(0) = replace(trim(aPair(0)), "NonMetric_","")
								objEditBox.SetTOProperty "Index", 1 
							Else
								objEditBox.SetTOProperty "Index", 0 
							End if

							objEditBox.SetTOProperty "attached text", trim(aPair(0))
							If objEditBox.Exist(5) Then
								If trim(aPair(1)) <> objEditBox.GetROProperty("value") then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_ClassAdmin_AtrributeOperations ] value of object [ " & trim(aPair(0)) & " ] did not match with [ " & trim(aPair(1)) & " ].")
									Set objEditBox = nothing
									Exit function
								End If
							Else
								objEditBox.SetTOProperty "attached text", trim(aPair(0))&":"
								If objEditBox.Exist(5) Then
									If trim(aPair(1)) <> objEditBox.GetROProperty("value") then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_ClassAdmin_AtrributeOperations ] value of object [ " & trim(aPair(0)) & " ] did not match with [ " & trim(aPair(1)) & " ].")
										Set objEditBox = nothing
										Exit function
									End If
								Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_ClassAdmin_AtrributeOperations ] Object [ " & trim(aPair(0)) & " ] does not exist.")
									Set objEditBox = nothing
									Exit function
								End If
								
							End If
							Set objEditBox = nothing
						' - - - - - - -  - - - - - - -  - - - - - - -  
						Case Else
							' for future use
							' do nothing... skip
							' invalid input
							Set objEditBox = nothing
							Exit function
					End Select
				Next
				Fn_ClassAdmin_AtrributeOperations = True
			End If
' - - - - - - -  - - - - - - -  - - - - - - -  - - - - - - -  - - - - - - -  - - - - - - -  - - - - - - -  - - - - - - -  - - - - - - -  - - - - - - -  - - - - - - -  - - - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_ClassAdmin_AtrributeOperations ] Invalid case [ " & sAction & " ].")
			Exit function
	End Select
	' performing save operation
	If bSave <> "" Then
		If cBool(bSave) Then
			If Fn_ClassAdmin_ToolbarOperations("Save") = False then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_ClassAdmin_AtrributeOperations ] Failed click on Save tool bar button.")
				Fn_ClassAdmin_AtrributeOperations = False
				Exit function
			End If
		End If
	End If
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: [ Fn_ClassAdmin_AtrributeOperations ] Executed successfully with case [ " & sAction & " ].")
 End Function

'********************************************************* Function to perofrm on view for storage class ***********************************************************************

'Function Name		          :      Fn_ClassAdmin_ViewOperations

'Description			    :	 Function to perofrm on  view for storage class

'Parameters			   :	1. sAction : Action to perform
'						2. sStorageClass : Storage class to be selected
'						3. dicViewOperations : Dictionary Object
'						4. bSave : Boolean value to perform save operation.

'Return Value		         : 	True / False

'Examples				:       

'							dicViewOperations( "sViewType" ) =  "User View"
'							dicViewOperations( "sViewID" ) = "123"
'							dicViewOperations( "sViewName" ) =  "testUserView"
'							dicViewOperations( "sAddImageUrl" ) = "c:\image.jpg" ' image url to add image
'							dicViewOperations("bRemoveImage")  = True 'To remove image
'							dicViewOperations( "bViewAttributes" ) =  "True" 
'							dicViewOperations( "sViewAttributes" ) =  "1036 Integer"  ' to select attribute from View Attribute
'							dicViewOperations("sFieldValues")= "Unique:True~Default Value:100" ' ~ separated list of fields:values
'							dicViewOperations("bActivateOnLastEntry") =  "True"
'							dicViewOperations( "sRemoveList" ) =  ""
'							dicViewOperations( "sAddToRightList" ) =  "1036 Integer ~1037 Real "
'												
'					       call Fn_ClassAdmin_ViewOperations("AddNewView", "SAM Classification Root:Classification Root:Storage_25694", dicViewOperations, True)
'					       call Fn_ClassAdmin_ViewOperations("Edit", "SAM Classification Root:Classification Root:Storage_25694", dicViewOperations, "")

'							dicViewOperations( "sViewType" ) =  "User View"
'							dicViewOperations( "sViewName" ) =  "testUserView"
'							dicViewOperations( "bViewAttributes" ) =  "True" 
'							dicViewOperations( "sViewAttributes" ) =  "1036  Integer  "  ' to select attribute from View Attribute 2 spaces in beteen ID and Name and at the end
'							dicViewOperations("sFieldValues")= "Unique:True~Default Value:100" ' ~ separated list of fields:values
' - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - - 
'					      NOTE : View Attribute can be verified 1 at a time using dicViewOperations( "bViewAttributes" ) 
'							Fields and values of 1 attribute can be verified at a time.
' - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - - 
'							call Fn_ClassAdmin_ViewOperations("Verify", "SAM Classification Root:Classification Root:Storage_25694", dicViewOperations, "")
'							call Fn_ClassAdmin_ViewOperations("SetValues", "", dicViewOperations, "")
							        
'History:
'			Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'			Koustubh 				  22-Dec-2010			   1.0
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'			Koustubh 				  22-Dec-2010			   1.0			Added code to remove image			
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'			Koustubh 				  27-Dec-2010			   1.0			Added case Verify, SetValues			
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_ClassAdmin_ViewOperations(sAction, sStorageClass, dicViewOperations, bSave)
	GBL_FAILED_FUNCTION_NAME="Fn_ClassAdmin_ViewOperations"
	Dim objAddView, objApplet, iCounter, aAttribArr, bLayOutTag, bClassAttribute
          Dim aLayOut, sLayOut, sName, aPairs, aFieldValue, objCheckProp, objPropEdit 
	Fn_ClassAdmin_ViewOperations = False
	Set objApplet = JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow")
	Select Case sAction
		' - - - - - - - - - - - Create / Edit view- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "AddNewView", "Edit", "SetValues"
			Select Case sAction
				' - - - - - - - - - - - Create new view- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
				Case "AddNewView"
					Set objAddView =objApplet.JavaDialog("Add View")
					If sStorageClass <> "" Then
						If Fn_ClassAdmin_TreeNodeOperation("Select",sStorageClass,"") = False then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_ClassAdmin_ViewOperations ] Failed to Select Storage class [ " & sStorageClass & " ].")
							Set objAddView = nothing
							Set objApplet = nothing
							Exit function
						End if
						If Fn_ClassAdmin_TabOpeartions("Activate","Subtab","Class Details","") = False Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_ClassAdmin_ViewOperations ] Failed to activate Tab [ Class Details ].")
							Fn_ClassAdmin_ViewOperations = False
							Exit function
						End If 
						objApplet.JavaButton("Add View").Click micLeftBtn
					End If
					If objAddView.Exist(10) = False Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_ClassAdmin_ViewOperations ] Add View dialog does not exist.")
						Set objAddView = nothing
						Set objApplet = nothing
						Exit function
					End If
					' setting unique View Type
					If dicViewOperations("sViewType") <> ""  Then
						objAddView.JavaList("View Type").Select dicViewOperations("sViewType")
						wait(3)
					End If
					' setting unique View ID
					 If dicViewOperations("sViewID") <> "" Then
						   If instr(1,dicViewOperations("sViewID"),":") > 0 then
								 aValues = split(dicViewOperations("sViewID"),":",-1,1)
								 objAddView.JavaEdit("View ID").SetToProperty "attached text",aValues(0)
								 objAddView.JavaEdit("View ID").Set aValues(1)
						  Else
								objAddView.JavaEdit("View ID").Set dicViewOperations("sViewID") 
						  End if 				
					 End If

					'clicking on OK
					if cInt(objAddView.JavaButton("OK").GetROProperty("enabled")) = 1 then
						objAddView.JavaButton("OK").Click micLeftBtn
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_ClassAdmin_ViewOperations ] OK Button is not activated.")
						Set objAddView = nothing
						Set objApplet = nothing
						Exit function
					End IF
					Set objAddView = nothing
				' - - - - - - - - - - - - Edit View - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
				Case "Edit"
					If sStorageClass <> "" Then
						If Fn_ClassAdmin_TreeNodeOperation("Select",sStorageClass,"") = False then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_ClassAdmin_ViewOperations ] Failed to Select Storage class [ " & sStorageClass & " ].")
							Set objAddView = nothing
							Set objApplet = nothing
							Exit function
						End if 
					End If
					' clicking on edit
					If Fn_ClassAdmin_ToolbarOperations("Edit") = False then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_ClassAdmin_ViewOperations ] Failed click on Edit tool bar button.")
						Fn_ClassAdmin_ViewOperations = False
						Exit function
					End If
					
				' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
				Case "SetValues"
						' do nothing 
						'View is in edit mode
			End select
			wait 5 
			'set name
			If  dicViewOperations("sViewName") <> "" Then
				call Fn_SISW_UI_JavaEdit_Operations("Fn_ClassAdmin_ClassOperations", "Set",  objApplet, "GroupName", dicViewOperations("sViewName"))
			End If
			wait 1
			'add Image
			If dicViewOperations("sAddImageUrl") <> ""  Then
				'Click on Add Image Button
				objApplet.JavaButton("AddImage").Click micLeftBtn
				' handle import dialog
                                        'Check existance of the Select Image Dialog Box
				Call Fn_ReadyStatusSync(3)
				If JavaDialog("SelectImageDialog").Exist(5) = True Then
					JavaDialog("SelectImageDialog").Activate
					'Paste the url in file name edit box
					wait 3
					call Fn_SISW_UI_JavaEdit_Operations("Fn_ClassAdmin_ClassOperations","Set",JavaDialog("SelectImageDialog"),"FileName",dicViewOperations("sAddImageUrl"))
					'Click on Import button
					 wait 2
					JavaDialog("SelectImageDialog").JavaButton("Add").Click micLeftBtn
				End If
			End If
			' remove Image
			If dicViewOperations("bRemoveImage") <> "" Then
				If cBool(dicViewOperations("bRemoveImage") ) then
					objApplet.JavaButton("DeleteImage").Click micLeftBtn
					Wait 5
				End if 
			End if
			'set View Details
			 If dicViewOperations("bViewDetails") <> "" then
				If cBool(dicViewOperations("bViewDetails")) Then
					'activate View Details tab
				         If Fn_ClassAdmin_TabOpeartions("Activate","Subtab","View Details","") = False Then
	 						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_ClassAdmin_ViewOperations ] Failed to activate Tab [ View Details ].")
							Fn_ClassAdmin_ViewOperations = False
							Exit function
					End If
					wait 1
					' set user 1
					If dicViewOperations("sUser1") <> "" then
						objApplet.JavaEdit("AttributeDefaultValue").SetTOProperty "attached text", "User 1:"
						If objApplet.JavaEdit("AttributeDefaultValue").exist(5) then
							If cInt(objApplet.JavaEdit("AttributeDefaultValue").GetROProperty("enabled") ) = 1 then
								objApplet.JavaEdit("AttributeDefaultValue").Set  dicViewOperations("sUser1")
							Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_ClassAdmin_ViewOperations ]  [ User 1: ] edit box is disabled.")
								Set objApplet = nothing
								Exit function
							end if
						Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_ClassAdmin_ViewOperations ]  [ User 1: ] edit box does not exist.")
							Set objApplet = nothing
							Exit function 
						end if
					End If
					' set user 2
					If dicViewOperations("sUser2") <> "" then
						objApplet.JavaEdit("AttributeDefaultValue").SetTOProperty "attached text", "2:"
						If objApplet.JavaEdit("AttributeDefaultValue").exist(5) then
							If cInt(objApplet.JavaEdit("AttributeDefaultValue").GetROProperty("enabled") ) = 1 then
								objApplet.JavaEdit("AttributeDefaultValue").Set  dicViewOperations("sUser2")
							Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_ClassAdmin_ViewOperations ]  [ User 2: ] edit box is disabled.")
								Set objApplet = nothing
								Exit function
							end if
						Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_ClassAdmin_ViewOperations ]  [ User 2: ] edit box does not exist.")
							Set objApplet = nothing
							Exit function 
						end if
					End If
					
				End If
			End If

			'Select the Class Attributes
			If dicViewOperations("bClassAttributes") <> "" then
						If Fn_ClassAdmin_TabOpeartions("Activate","Subtab","View Attributes","") = False Then
	 					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_ClassAdmin_ViewOperations ] Failed to activate Tab [ View Attributes ].")
						Fn_ClassAdmin_ViewOperations = False
						Exit function
					End If
					wait 1

					'Select the Attribute
					if dicViewOperations("sClassAttributesSelect") <> "" then
						aAttribArr = split(dicViewOperations("sClassAttributesSelect") ,"~")
						For iCounter = 0  to UBound(aAttribArr)
							wait 1
							If Fn_UI_ListItemExist("Fn_ClassAdmin_ViewOperations",objApplet,"ClassAttributesList", aAttribArr(iCounter)) Then
								wait 2
								objApplet.JavaList("ClassAttributesList").ExtendSelect trim(aAttribArr(iCounter))
							 End if 
						  Next
					 End if
			End if 

			'set View Attributes
			If dicViewOperations("bViewAttributes") <> "" then
				If cBool(dicViewOperations("bViewAttributes")) Then
					'activate View Attribute tab
					'JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaTab("SubTab").Select 
					If Fn_ClassAdmin_TabOpeartions("Activate","Subtab","View Attributes","") = False Then
	 					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_ClassAdmin_ViewOperations ] Failed to activate Tab [ View Attributes ].")
						Fn_ClassAdmin_ViewOperations = False
						Exit function
					End If
					wait 1
					'select AddList
					if dicViewOperations("sAddToRightList") <> "" then
						aAttribArr = split(dicViewOperations("sAddToRightList") ,"~")
						For iCounter = 0  to UBound(aAttribArr)
							wait 1
							If Fn_UI_ListItemExist("Fn_ClassAdmin_ViewOperations",objApplet,"ClassAttributesList", aAttribArr(iCounter)) Then
								wait 2
								objApplet.JavaList("ClassAttributesList").Select trim(aAttribArr(iCounter))
								' click on add
								wait 1
								objApplet.JavaButton("MoveRIght").Click micLeftBtn
							Else
								If instr(aAttribArr(iCounter),":") > 0 Then
									aLayOut = split(aAttribArr(iCounter),":") 
									sLayOut = aLayOut(0)
									sName =  aLayOut(1)
								Else
									sLayOut = aAttribArr(iCounter)
									sName =  ""
								End If
								If Fn_UI_ListItemExist("Fn_ClassAdmin_ViewOperations",objApplet,"ClassAttributesList", sLayOut) then
									objApplet.JavaList("ClassAttributesList").Select aAttribArr(iCounter)
									' click on add
									objApplet.JavaButton("MoveRIght").Click micLeftBtn
									' Layout Tag Parameter
									If objApplet.JavaDialog("Layout Tag Parameter").Exist(10) then
										objApplet.JavaDialog("Layout Tag Parameter").JavaEdit("EditBox").Set sName
										objApplet.JavaDialog("Layout Tag Parameter").JavaButton("OK").Click micLeftBtn
									End If 
								eLSE
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_ClassAdmin_ViewOperations ] Failed to select  [ " & aAttribArr(iCounter) & " ] in lists.")
									Set objApplet = nothing
									Exit function
								End If
							End IF
						Next
					End If
					' select remove list
					if dicViewOperations("sRemoveList") <> "" then
						aAttribArr = split(dicViewOperations("sRemoveList") ,"~")
                                                            For iCounter = 0  to UBound(aAttribArr)
							If Fn_UI_ListItemExist("Fn_ClassAdmin_ViewOperations",objApplet,"ViewAttributesList", aAttribArr(iCounter)) Then
									objApplet.JavaList("ViewAttributesList").ExtendSelect aAttribArr(iCounter)
							Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_ClassAdmin_ViewOperations ] Failed to select  [ " & aAttribArr(iCounter) & " ] in View Attributes List.")
									Set objApplet = nothing
									Exit function
							End If
						Next
						objApplet.JavaButton("MoveLeft").Click micLeftBtn
					End IF
					' selecting view attributes
					If dicViewOperations( "sViewAttributes") <> "" Then
							If Fn_UI_ListItemExist("Fn_ClassAdmin_ViewOperations",objApplet,"ViewAttributesList", dicViewOperations( "sViewAttributes") ) Then
								wait 1
								objApplet.JavaList("ViewAttributesList").Select dicViewOperations( "sViewAttributes") 
							Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_ClassAdmin_ViewOperations ] Failed to select  [ " & dicViewOperations( "sViewAttributes") & " ] in View Attributes List.")
								Set objApplet = nothing
								Exit function
							End If
					End If
					' setting values to check boxes. . .
					If dicViewOperations("sFieldValues") <> ""  Then
						If dicViewOperations("bActivateOnLastEntry") = "" then dicViewOperations("bActivateOnLastEntry") = False
						aPairs = split(dicViewOperations("sFieldValues"), "~")
						For iCounter = 0 to UBound(aPairs)
                                                                      aFieldValue = split(aPairs(iCounter),":")
                                                                      Select Case UBound(aFieldValue)
								Case 0
									' do nothing... skip
								' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
								Case 1
	                                                                                Select Case trim(aFieldValue(0))
										Case "Config"
										'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
										Case "NonMetric_Fixed"
											Set objCheckProp = JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaCheckBox("ChkProperties")
											objCheckProp.SetTOProperty "attached text", "Fixed"
											objCheckProp.SetTOProperty "Index", 1
											If objCheckProp.Exist(5) Then
												If cInt(objCheckProp.GetROProperty("enabled")) = 1 Then
													If cBool(trim(aFieldValue(1))) Then
														objCheckProp.set "ON"
													Else
														objCheckProp.set "OFF"
													End If
													
												Else
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_ClassAdmin_ViewOperations ] Object [ " & trim(aFieldValue(0)) & " ] is not enabled.")
													Set objCheckProp = nothing
													Set objPropEdit = nothing
													Exit function
												End If
											Else
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_ClassAdmin_ViewOperations ] Object [ " & trim(aFieldValue(0)) & " ] is not exist.")
												Set objCheckProp = nothing
												Set objPropEdit = nothing
												Exit function
											End If
											Set objCheckProp = nothing
											Set objPropEdit = nothing
										'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
										Case "NonMetric_Default Value From Class",  "NonMetric_Minimum Value From Class",  "NonMetric_Maximum Value From Class", "Default Value From Class",  "Minimum Value From Class",  "Maximum Value From Class"
											Call Fn_ClassAdmin_TabOpeartions("Activate","ImageTab","Min/Max/Default Value","")
											wait 1, 500											
											Set objCheckProp = JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaCheckBox("ChkProperties")
											objCheckProp.SetTOProperty "attached text",  "From Class"
											Select Case  trim(aFieldValue(0))
												Case "Default Value From Class"
													objCheckProp.SetTOProperty "Index", 0
												Case "Minimum Value From Class"
													objCheckProp.SetTOProperty "Index", 1
												Case "Maximum Value From Class"
													objCheckProp.SetTOProperty "Index", 2
												Case "NonMetric_Default Value From Class"
													objCheckProp.SetTOProperty "Index", 3
												Case "NonMetric_Minimum Value From Class"
													objCheckProp.SetTOProperty "Index", 4
												Case "NonMetric_Maximum Value From Class"
													objCheckProp.SetTOProperty "Index", 5
											End Select
											If objCheckProp.Exist(5) Then
												If cInt(objCheckProp.GetROProperty("enabled")) = 1 Then
													If cBool(trim(aFieldValue(1))) Then
														objCheckProp.set  "ON"
													Else
														objCheckProp.set   "OFF"
													End If
												Else
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_ClassAdmin_ViewOperations ] Object [ " & trim(aFieldValue(0)) & " ] is not enabled.")
													Set objCheckProp = nothing
													Exit function
												End If
											Else
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_ClassAdmin_ViewOperations ] Object [ " & trim(aFieldValue(0)) & " ] does not exist.")
												Set objCheckProp = nothing
												Exit function
											End If
											Set objCheckProp = nothing
										'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
										Case  "NonMetric_Default Value",  "NonMetric_Minimum Value",  "NonMetric_Maximum Value", "Default Value",  "Minimum Value",  "Maximum Value"
												Call Fn_ClassAdmin_TabOpeartions("Activate","ImageTab","Min/Max/Default Value","")
												wait 1, 500
												Set objPropEdit = JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaEdit("PropertyEditBox")
												If instr(aFieldValue(0),"NonMetric_") > 0 Then
													aFieldValue(0) = replace(trim(aFieldValue(0)),"NonMetric_","")
													Set objCheckProp = JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaCheckBox(trim(aFieldValue(0)))
													objCheckProp.SetTOProperty "Index",  1
													objPropEdit.SetTOProperty "Index", 1
												Else
													Set objCheckProp = JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaCheckBox(trim(aFieldValue(0)))
													objCheckProp.SetTOProperty "Index", 0
													objPropEdit.SetTOProperty "Index", 0	
												End If
														
												If objCheckProp.Exist(5) Then
													If cInt(objCheckProp.GetROProperty("enabled")) = 1 Then
														If lcase(trim(aFieldValue(1))) = "true" OR lcase(trim(aFieldValue(1))) = "false" Then
															If cBool(trim(aFieldValue(1))) Then
																objCheckProp.set "ON"
															Else
																objCheckProp.set "OFF"
															End If
														Else
															objCheckProp.set "ON"
															objPropEdit.SetTOProperty "attached text", trim(aFieldValue(0))
															wait 1
															objPropEdit.Set trim(aFieldValue(1))
															If iCounter = UBound(aPairs) AND cBool(dicViewOperations("bActivateOnLastEntry")) = True then
																wait 1
																objPropEdit.Activate
															End If
														End If
													Else
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_ClassAdmin_ViewOperations ] Object [ " & trim(aFieldValue(0)) & " ] is not enabled.")
														Set objCheckProp = nothing
														Set objPropEdit = nothing
														Exit function
													End If
												Else
													Set objCheckProp = nothing
													Set objPropEdit = nothing
													Exit function
												End If
												Set objCheckProp = nothing
												Set objPropEdit = nothing
										' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
										Case Else
	                                             ' for remaining checkboxes and edit boxes.
												Set objCheckProp = JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaCheckBox("ChkProperties")
												Set objPropEdit = JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaEdit("PropertyEditBox")
												If aFieldValue(0) = "NX CAM" Then
													aFieldValue(0) = "<html> Marks the attribute as important for Application 1 </html>"
												ElseIf aFieldValue(0) = "Graphics Creation" Then
													aFieldValue(0) = "<html> Marks the attribute as important for Application 2 </html>"
												ElseIf aFieldValue(0) = "GCS Connection" Then
													aFieldValue(0) = "<html> Marks the attribute as important for Application 3 </html>"
												ElseIf aFieldValue(0) = "Application 4" Then
													aFieldValue(0) = "<html> Marks the attribute as important for Application 4 </html>"
												ElseIf aFieldValue(0) = "Application 5" Then
													aFieldValue(0) = "<html> Marks the attribute as important for Application 5 </html>"
												End If
												objCheckProp.SetTOProperty "attached text",  trim(aFieldValue(0))
												objPropEdit.SetTOProperty "attached text", trim(aFieldValue(0)) & ":"
												objCheckProp.SetTOProperty "Index", 0
												objPropEdit.SetTOProperty "Index", 0
														
												If objCheckProp.Exist(5) Then
													If cInt(objCheckProp.GetROProperty("enabled")) = 1 Then
														' setting checkbox on off
														wait 1
														If lcase(trim(aFieldValue(1))) = "true" OR lcase(trim(aFieldValue(1))) = "false" Then
															If cBool(trim(aFieldValue(1))) Then
																objCheckProp.set  "ON"
															Else
																objCheckProp.set   "OFF"
															End If
														End If
													Else
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_ClassAdmin_ViewOperations ] Object [ " & trim(aFieldValue(0)) & " ] is not enabled.")
														Set objCheckProp = nothing
														Set objPropEdit = nothing
														Exit function
													End If									
													
												Elseif objPropEdit.Exist(5) then ' setting values to edit box
														objPropEdit.SetTOProperty "Index", 0
														objPropEdit.Set trim(aFieldValue(1))
												Else
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_ClassAdmin_ViewOperations ] Object [ " & trim(aFieldValue(0)) & " ] does not exist.")
													Set objCheckProp = nothing
													Set objPropEdit = nothing
													Exit function
												End If
												Set objCheckProp = nothing
												Set objPropEdit = nothing
									'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
									End Select
								'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
								Case Else
                                                                                          ' field value pair with multiple values
									' this case is for future use.
							End Select
						Next
					End If ' end of If dicViewOperations("sFieldValues") <> ""  Then
				End If ' end of if true
			End If 'end of if not empty
			Fn_ClassAdmin_ViewOperations = True
' - - - - - - Verify Details- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -	
		Case "Verify"
			'select view
			If sStorageClass <> "" Then
				If Fn_ClassAdmin_TreeNodeOperation("Select",sStorageClass,"") = False then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_ClassAdmin_ViewOperations ] Failed to Select view [ " & sStorageClass & " ].")
					Set objAddView = nothing
					Set objApplet = nothing
					Exit function
				End if 
			End If
			If dicViewOperations( "sViewType" ) <> "" Then
				If objApplet.JavaStaticText("ViewType").GetROProperty("value") <> dicViewOperations( "sViewType" ) then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_ClassAdmin_ViewOperations ] View Type is not matched.")
					Set objApplet = nothing
					Exit function
				End If
			End If
			If dicViewOperations( "sViewName" ) <> "" then
				If objApplet.JavaEdit("GroupName").getROProperty("value") <> dicViewOperations( "sViewName" ) then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_ClassAdmin_ViewOperations ] View Name is not matched.")
					Set objApplet = nothing
					Exit function
				End IF 
			End if
			If dicViewOperations("bViewAttributes") <> "" then
				If cBool(dicViewOperations("bViewAttributes")) Then
					If Fn_ClassAdmin_TabOpeartions("Activate","Subtab","View Attributes","") = False Then
	 					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_ClassAdmin_ViewOperations ] Failed to activate Tab [ View Attributes ].")
						Fn_ClassAdmin_ViewOperations = False
						Exit function
					End If
					wait 1
					If dicViewOperations("sViewAttributes" ) <> "" Then
						If Fn_UI_ListItemExist("Fn_ClassAdmin_ViewOperations",objApplet,"ViewAttributesList", dicViewOperations("sViewAttributes" )) = False Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_ClassAdmin_ViewOperations ] Failed to select  [ " & dicViewOperations("sViewAttributes" ) & " ] in View Attributes List.")
							Set objApplet = nothing
							Exit function
						End If
						If dicViewOperations("sFieldValues") <> "" Then
							objApplet.JavaList("ViewAttributesList").Select dicViewOperations("sViewAttributes" )
							aPairs = split(dicViewOperations("sFieldValues"), "~")
							For iCounter = 0 to UBound(aPairs)
								aFieldValue = split(aPairs(iCounter),":")
								Select Case UBound(aFieldValue)
									Case 0
										' do nothing... skip
									' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
									Case 1
										Select Case trim(aFieldValue(0))
											Case "Config"
											'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
											Case "NonMetric_Fixed"
												Set objCheckProp = JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaCheckBox("ChkProperties")
												objCheckProp.SetTOProperty "attached text", "Fixed"
												objCheckProp.SetTOProperty "Index", 1
												If objCheckProp.Exist(5) Then
													If lcase(trim(aFieldValue(1))) = "true" OR lcase(trim(aFieldValue(1))) = "false" Then
														If cBool(trim(aFieldValue(1))) Then
															If cInt(objCheckProp.GetROProperty ("value")) = 0 then
																Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_ClassAdmin_ViewOperations ] [ " & trim(aFieldValue(0)) &  " ] value did not match.")
																Set objApplet = nothing
																Exit function
															End IF
														Else
															If cInt(objCheckProp.GetROProperty ("value")) = 1 then
																Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_ClassAdmin_ViewOperations ] [ " & trim(aFieldValue(0)) &  " ] value did not match.")
																Set objApplet = nothing
																Exit function
															End IF
														End If
													Else
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_ClassAdmin_ViewOperations ] Object [ " & trim(aFieldValue(0)) & " ] with invalid value.")
														Set objCheckProp = nothing
														Set objPropEdit = nothing
														Exit function
													End If
												Else
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_ClassAdmin_ViewOperations ] Object [ " & trim(aFieldValue(0)) & " ] does not exist.")
													Set objCheckProp = nothing
													Set objPropEdit = nothing
													Exit function
												End If
												Set objCheckProp = nothing
												Set objPropEdit = nothing
										'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
										
											Case "NonMetric_Default Value From Class",  "NonMetric_Minimum Value From Class",  "NonMetric_Maximum Value From Class", "Default Value From Class",  "Minimum Value From Class",  "Maximum Value From Class"
												Call Fn_ClassAdmin_TabOpeartions("Activate","ImageTab","Min/Max/Default Value","")
												wait 1, 500	
												Set objCheckProp = JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaCheckBox("ChkProperties")
												objCheckProp.SetTOProperty "attached text",  "From Class"
												Select Case  trim(aFieldValue(0))
													Case "Default Value From Class"
														objCheckProp.SetTOProperty "Index", 0
													Case "Minimum Value From Class"
														objCheckProp.SetTOProperty "Index", 1
													Case "Maximum Value From Class"
														objCheckProp.SetTOProperty "Index", 2
													Case "NonMetric_Default Value From Class"
														objCheckProp.SetTOProperty "Index", 3
													Case "NonMetric_Minimum Value From Class"
														objCheckProp.SetTOProperty "Index", 4
													Case "NonMetric_Maximum Value From Class"
														objCheckProp.SetTOProperty "Index", 5
												End Select
												If objCheckProp.Exist(5) Then
													If cBool(trim(aFieldValue(1))) Then
														If cInt(objCheckProp.GetROProperty ("value")) = 0 then
															Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_ClassAdmin_ViewOperations ] [ " & trim(aFieldValue(0)) &  " ] value did not match.")
															Set objApplet = nothing
															Exit function
														End IF
													Else
														If cInt(objCheckProp.GetROProperty ("value")) = 1 then
															Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_ClassAdmin_ViewOperations ] [ " & trim(aFieldValue(0)) &  " ] value did not match.")
															Set objApplet = nothing
															Exit function
														End IF
													End If
												Else
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_ClassAdmin_ViewOperations ] Object [ " & trim(aFieldValue(0)) & " ] does not exist.")
													Set objCheckProp = nothing
													Exit function
												End If
												Set objCheckProp = nothing
											'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
											Case  "NonMetric_Default Value",  "NonMetric_Minimum Value",  "NonMetric_Maximum Value", "Default Value",  "Minimum Value",  "Maximum Value"
													Call Fn_ClassAdmin_TabOpeartions("Activate","ImageTab","Min/Max/Default Value","")
													wait 1, 500														
													Set objPropEdit = JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaEdit("PropertyEditBox")
													If instr(aFieldValue(0),"NonMetric_") > 0 Then
														aFieldValue(0) = replace(trim(aFieldValue(0)),"NonMetric_","")
														Set objCheckProp = JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaCheckBox(trim(aFieldValue(0)))
														objCheckProp.SetTOProperty "Index",  1
														objPropEdit.SetTOProperty "Index", 1
													Else
														Set objCheckProp = JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaCheckBox(trim(aFieldValue(0)))
														objCheckProp.SetTOProperty "Index", 0
														objPropEdit.SetTOProperty "Index", 0	
													End If
															
													If objCheckProp.Exist(5) Then
														If lcase(trim(aFieldValue(1))) = "true" OR lcase(trim(aFieldValue(1))) = "false" Then
															If cBool(trim(aFieldValue(1))) Then
																If cInt(objCheckProp.GetROProperty ("value")) = 0 then
																	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_ClassAdmin_ViewOperations ] [ " & trim(aFieldValue(0)) &  " ] value did not match.")
																	Set objApplet = nothing
																	Exit function
																End IF
															Else
																If cInt(objCheckProp.GetROProperty ("value")) = 1 then
																	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_ClassAdmin_ViewOperations ] [ " & trim(aFieldValue(0)) &  " ] value did not match.")
																	Set objApplet = nothing
																	Exit function
																End IF
															End If
														Else
															If cInt(objCheckProp.GetROProperty ("value")) = 0 then
																Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_ClassAdmin_ViewOperations ] [ " & trim(aFieldValue(0)) &  " ] value did not match.")
																Set objApplet = nothing
																Exit function
															End IF
															objPropEdit.SetTOProperty "attached text", trim(aFieldValue(0))
															wait 1
															If objPropEdit.GetROProperty ("value") <> trim(aFieldValue(1)) then
																Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_ClassAdmin_ViewOperations ] [ " & trim(aFieldValue(0)) &  " ] value did not match.")
																Set objApplet = nothing
																Exit function
															End If
														End If
													Else
														Set objCheckProp = nothing
														Set objPropEdit = nothing
														Exit function
													End If
													Set objCheckProp = nothing
													Set objPropEdit = nothing
											' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
											Case Else
													' for remaining checkboxes and edit boxes.
													Set objCheckProp = JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaCheckBox("ChkProperties")
													Set objPropEdit = JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaEdit("PropertyEditBox")
													If aFieldValue(0) = "NX CAM" Then
														aFieldValue(0) = "<html> Marks the attribute as important for Application 1 </html>"
													ElseIf aFieldValue(0) = "Graphics Creation" Then
														aFieldValue(0) = "<html> Marks the attribute as important for Application 2 </html>"
													ElseIf aFieldValue(0) = "GCS Connection" Then
														aFieldValue(0) = "<html> Marks the attribute as important for Application 3 </html>"
													ElseIf aFieldValue(0) = "Application 4" Then
														aFieldValue(0) = "<html> Marks the attribute as important for Application 4 </html>"
													ElseIf aFieldValue(0) = "Application 5" Then
														aFieldValue(0) = "<html> Marks the attribute as important for Application 5 </html>"
													End If
													objCheckProp.SetTOProperty "attached text",  trim(aFieldValue(0))
													objPropEdit.SetTOProperty "attached text", trim(aFieldValue(0)) & ":"
													objCheckProp.SetTOProperty "Index", 0
													objPropEdit.SetTOProperty "Index", 0															
													If objCheckProp.Exist(5) Then
														If lcase(trim(aFieldValue(1))) = "true" OR lcase(trim(aFieldValue(1))) = "false" Then
															If cBool(trim(aFieldValue(1))) Then
																If cInt(objCheckProp.GetROProperty ("value")) = 0 then
																	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_ClassAdmin_ViewOperations ] [ " & trim(aFieldValue(0)) &  " ] value did not match.")
																	Set objApplet = nothing
																	Exit function
																End IF
															Else
																If cInt(objCheckProp.GetROProperty ("value")) = 1 then
																	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_ClassAdmin_ViewOperations ] [ " & trim(aFieldValue(0)) &  " ] value did not match.")
																	Set objApplet = nothing
																	Exit function
																End IF
															End If
														End If	
													Elseif objPropEdit.Exist(5) then ' setting values to edit box
															objPropEdit.SetTOProperty "Index", 0
															If objPropEdit.GetROProperty ("value") <> trim(aFieldValue(1)) then
																Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_ClassAdmin_ViewOperations ] [ " & trim(aFieldValue(0)) &  " ] value did not match.")
																Set objApplet = nothing
																Exit function
															End If
													Else
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_ClassAdmin_ViewOperations ] Object [ " & trim(aFieldValue(0)) & " ] does not exist.")
														Set objCheckProp = nothing
														Set objPropEdit = nothing
														Exit function
													End If
													Set objCheckProp = nothing
													Set objPropEdit = nothing
										'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
										End Select
									'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
									Case Else
										' field value pair with multiple values
										' this case is for future use.
								End Select
							Next
						End If
					End If
					
				End If
			End If
			Fn_ClassAdmin_ViewOperations = true
' - - - - - - Default case - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -	
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_ClassAdmin_ViewOperations ] Invalid case [ " & sAction & " ].")
			Set objApplet = nothing
			Exit function
	End Select
          ' performing save operation
	If bSave <> "" Then
		If cBool(bSave) Then
			If Fn_ClassAdmin_ToolbarOperations("Save") = False then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_ClassAdmin_ViewOperations ] Failed click on Save tool bar button.")
				Fn_ClassAdmin_ViewOperations = False
				Exit function
			End If
		End If
	End If
     Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: [ Fn_ClassAdmin_ViewOperations ] Executed successfully with case [ " & sAction & " ].")
	Set objApplet = nothing
End Function


'*********************************************************  Function performs  Operations on All Attributes Values*********************************************************************
'Function Name  :   Fn_ClassAdmin_ListofValues
'
' 
'Parameters      :     sAction: Rowcellexist/Rowexist
'           				  sObjectName: Unique value in row
'							  sPropertyName : Name of column
'							  sExpectedValue : value to verify
'							  sOther : for future use	
' 
'Return Value     :   True/False
'
'Examples    :      
'						call Fn_ClassAdmin_ListofValues("Rowexist","abcd","","","") 
'						call Fn_ClassAdmin_ListofValues("Rowcellexist","abcd","Count","1","") 
' 
'History:
'          Developer Name                          Date                                                       Rev. No.         Changes Done                   Reviewer 
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'           Prasanna    		                      14-Jan-2011   1.0                         
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_ClassAdmin_ListofValues(sAction, sObjectName, sPropertyName,sExpectedValue,sOther)
		GBL_FAILED_FUNCTION_NAME="Fn_ClassAdmin_ListofValues"
		Dim objDetailsTable,bReturn,bDoubleClickReturn,iCounter,aObjList,intItemCount,oCounter, rowIndex, bFlag, aMenuList, intCount, sMenu, sText, aMenuList1()
		Dim colCount, i, tab, textArr, columnNumber, columnFoundFlag, intObjectColumnNumber
		Dim colNameArr, bHeaderFoundFlag,objDetailsDialog
		columnFoundFlag = False
		bHeaderFoundFlag = False
		intObjectColumnNumber = -1
		Fn_ClassAdmin_ListofValues = False
		Err.Clear
		' create an object of the table
		'Set objDetailsTable =JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("ListofAttriValues").JavaTable("ValuesTable")
		Set objDetailsDialog =Fn_SISW_ClassAdmin_GetObject("ListofAttriValues")
		Set objDetailsTable=objDetailsDialog.JavaTable("ValuesTable")

		If objDetailsTable.Exist(5) = false Then
				JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaButton("BtnLisofValues").Click micLeftBtn
				If Err.Number < 0 Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Column does not exist")	
				End If
		End If
		colCount =  objDetailsTable.GetROProperty("cols")
		wait(5)
		' Mapping Object column to column number.
		
'		For i = 0 to colCount - 1
'			If objDetailsTable.GetColumnName(i) = "Value" then
'						intObjectColumnNumber = i
'						bHeaderFoundFlag = True
'						Exit for
'			end if
'		next
		If  bHeaderFoundFlag = False Then
				For i = 0 to colCount - 1 
						textArr = split(objDetailsTable.GetColumnName(i),"text=")
						wait(2)
						colNameArr = split(textArr(1),",")
						If trim(colNameArr(0)) = trim(sPropertyName)  then
									intObjectColumnNumber = i
									Exit for
						end if
				 Next
		End If
	
		 If intObjectColumnNumber = -1 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Column does not exist")	
					Exit function
		 End If
	
		bHeaderFoundFlag = False
		columnNumber = -1
	' Mapping column name to number
		If  sPropertyName<>"" Then
				For i = 0 to colCount - 1
					If objDetailsTable.GetColumnName(i) = sPropertyName then
							columnNumber = i
							bHeaderFoundFlag = True
							Exit for
					end if
				next
				If  bHeaderFoundFlag = False Then
					 For i = 0 to colCount - 1 
							textArr = split(objDetailsTable.GetColumnName(i),"text=")
							wait(2)
							colNameArr = split(textArr(1),",")
							if instr(1,colNameArr(0),sPropertyName) then
							'If colNameArr(0) = sPropertyName then
										columnNumber = i
										columnFoundFlag = true
										Exit for
							end if
					 Next
					If columnFoundFlag = true Then
								columnNumber = i
					else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Column " + sPropertyName + " does not exist")	
								Exit function
					 End If
				 End If
		End If

		Select Case sAction
	
					Case "Rowexist"
							bFlag = false
							'Count number of rows of Table
							bReturn = objDetailsTable.GetROProperty("rows")	
							wait(2)
							'Extract the index of row at which the object exist.
							For iCounter=0 to bReturn - 1
								sText = objDetailsTable.GetCellData(iCounter, intObjectColumnNumber )'	Object  column		
									wait(2)		
									If trim(cstr(sText)) = trim(cstr(sObjectName))  Then
											 bFlag = true
											 Exit for
									End If										
							Next
							If bFlag = false Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Fn_ClassAdmin_ListofValues : Row with Object "&sObjectName&" does not exist")	
									Exit function
							Else 
									Fn_ClassAdmin_ListofValues = True
							End If
				Case "Rowcellexist"
					If  sExpectedValue = ""  Then
							Fn_ClassAdmin_ListofValues = FALSE	 
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_ClassAdmin_ListofValues: Rowcellexist : Incorrect input parameters")
							Exit function
					End If
					'Count number of rows of Table
					bReturn = objDetailsTable.GetROProperty("rows")
					intCount = 0
					For iCounter=0 to bReturn - 1
							sText = objDetailsTable.GetCellData(iCounter, intObjectColumnNumber) ' Object column
							If IsNumeric(sObjectName) Then
								 If cstr(sText) = cstr(cint(sObjectName))  Then
									ReDim Preserve aMenuList1(intCount+1)
									aMenuList1( intCount) = iCounter
									intCount = intCount + 1
								End If
							elseIf cstr(sText) = cstr(sObjectName)  Then
								ReDim Preserve aMenuList1(intCount+1)
									aMenuList1(intCount) = iCounter
									intCount = intCount + 1
							End If
					Next
					If intCount <> 0 Then
							For iCounter = 0 To UBound(aMenuList1) - 1
									   If objDetailsTable.GetCellData(aMenuList1(iCounter), columnNumber ) = sExpectedValue  Then
												 intCount = aMenuList1(iCounter)
												 Exit For
									   End If
							Next
 						   If cstr(objDetailsTable.GetCellData(intCount-1, columnNumber)) <> sExpectedValue Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_ClassAdmin_ListofValues: Rowcellexist : Expected value is not present")
									 Fn_ClassAdmin_ListofValues = False  
									Exit function
							End If  
					End If
					Fn_ClassAdmin_ListofValues = True 
		End Select				

		objDetailsDialog.JavaButton("Close").Click micLeftBtn
		If Err.Number < 0  Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_ClassAdmin_ListofValues: Failed to click on Close Button")
				 Fn_ClassAdmin_ListofValues = False  
				Exit function
		End If		
End function


'*********************************************************  Function performs  Operations on All Attributes Values*********************************************************************
'Function Name  :   Fn_ClassAdmin_UnitClassBasics
'
' 
'Parameters      :     sAction: MoveUnitClass
'           				  		strNode value in row
'							  		sPropertyName : Name of column

'Return Value     :   True/False
'
'Examples    :      
'						call Fn_ClassAdmin_UnitClassBasics("MoveUnitClass","","")
' 
'History:
'          Developer Name                          Date                                                       Rev. No.         Changes Done                   Reviewer 
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'           Prasanna    		                      17-Jan-2011   1.0                         
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'           Jeevan    		                            7-Jun-2012                                 ObjMoveDef= Window("ClassificationAdminWindow").JavaWindow("ClassSubWindow").JavaDialog("MoveClassGroup")
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_ClassAdmin_UnitClassBasics(strAction,strNode,strMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_ClassAdmin_UnitClassBasics"
   Dim bReturn 
   Dim ObjMoveDef
   Fn_ClassAdmin_UnitClassBasics = false
   Err.Clear
	Select Case strAction
					Case "MoveUnitClass"
'                   	bReturn = Fn_ClassAdmin_TreeNodeOperation("RMB","SAM Classification Root","ExpandAll")
					bReturn = Fn_ClassAdmin_TreeNodeOperation("Expand","SAM Classification Root","")
					bReturn = Fn_ClassAdmin_TreeNodeOperation("Expand","SAM Classification Root:Classification Root","")
					If bReturn=false Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Unit Definition Class")
						Exit function
					Else
						 bReturn =  Fn_ClassAdmin_TreeNodeOperation("Exist","SAM Classification Root:Classification Root:Unit Definition Class","")
							If bReturn = true Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Pass : Unit Defination Class Does not exist Under SAM Classification Root but exists under SAM Classification Root:Classification Root")
								 Fn_ClassAdmin_UnitClassBasics = True
								Exit function
							Else
								 bReturn =  Fn_ClassAdmin_TreeNodeOperation("Exist","SAM Classification Root:Unit Definition Class","")
										If bReturn=false Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Unit Definition Class")
									Exit function

								  bReturn =  Fn_ClassAdmin_TreeNodeOperation("Select","SAM Classification Root:Unit Definition Class","")
								  If bReturn = false Then
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Unit Definition Class")
											Exit function
								   Else
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected Unit Definition Class")
											wait(1)
								  End If
								End if
							End if

							'select the unit definition class
								 bReturn =  Fn_ClassAdmin_TreeNodeOperation("Select","SAM Classification Root:Unit Definition Class","")
										If bReturn=false Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Unit Definition Class")
									Exit function
									End if
									Wait 2
								 ' Cut the task 
								  bReturn =  Fn_ClassAdmin_TreeNodeOperation("RMB","SAM Classification Root:Unit Definition Class","Cut Class")
								  If bReturn = false Then
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Cut Unit Definition Class")
											Exit function
								   Else
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Cut Unit Definition Class")
											wait(2)
								  End If

								 	'Select the Classification Root class		
								  bReturn =  Fn_ClassAdmin_TreeNodeOperation("Select","SAM Classification Root:Classification Root","")
								  If bReturn = false Then
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Classification Root")
											Exit function
								   Else
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected Classification Root")
											wait(1)
								  End If

								' Paste the task 
								'the menu was observed as "Paste StorageClass" in all previous builds except the ones from 20110119 builds,hence commenting it for time bieng
								'a PR has been filed for the same
								'can be uncommennted if it reverts to its original value which is "Paste StorageClass"

'								  bReturn =  Fn_ClassAdmin_TreeNodeOperation("RMB","SAM Classification Root:Classification Root","Paste StorageClass")

								  bReturn =  Fn_ClassAdmin_TreeNodeOperation("RMB","SAM Classification Root:Classification Root","Paste Class")
								  If bReturn = false Then
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Paste Unit Definition Class")
											Exit function
								   Else
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Paste Unit Definition Class")
											wait(1)
											Call Fn_ReadyStatusSync(3)
								  End If

									Set ObjMoveDef = JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("MoveClassGroup")
								    If ObjMoveDef.Exist(5) Then
											If ObjMoveDef.JavaButton("Yes").Exist(3) Then
												 ObjMoveDef.JavaButton("Yes").Click micLeftBtn
											 wait 2
											ElseIf ObjMoveDef.JavaButton("OK").Exist(3) Then
												ObjMoveDef.JavaButton("OK").Click micLeftBtn
											 wait 2
											End If	
											 If Err.Number < 0 Then
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Filed to Click on Yes/OK button")				
													Exit Function 
											  else
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Clicked on Yes/Ok button")				
											 End If
								  End If

					End If
	End Select
	Fn_ClassAdmin_UnitClassBasics = true
End function

'''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''''/$$$$
''''/$$$$   FUNCTION NAME   : Fn_ClassAdmin_CreateVerifyMapping(sAction,sSourceId,sTargetId,sButtons,bSave)
''''/$$$$
''''/$$$$   DESCRIPTION        :  This function will create mapping and verify if the mapping has happened
''''/$$$$
''''/$$$$    PARAMETERS      :   1.) sAction : Action to be performed
''''/$$$$                                      2.) sSourceId : Source Class Id
''''/$$$$ 									   3.) sTargetId : Target Class ID
''''/$$$$ 									   4.) sButtons  :  Buttons to be clicked
''''/$$$$									   4.) bSave  :  Perform Save Operation (should be passed as true to perforf save operation)
''''/$$$$	
''''/$$$$	Return Value : 			True or False
''''/$$$$
''''/$$$$    Function Calls       :   Fn_WriteLogFile()
''''/$$$$										
''''/$$$$
''''/$$$$	HISTORY           :   AUTHOR                 DATE        VERSION
''''/$$$$
''''/$$$$    CREATED BY     :   SHREYAS           24/01/2011         1.0
''''/$$$$
''''/$$$$    REVIWED BY     :  Prasanna			      24/01/2011         1.0
''''/$$$$
''''/$$$$    EXAMPLE          : bReturn=Fn_ClassAdmin_CreateVerifyMapping("Map","1066","1066","Map","")
''''/$$$$								  bReturn=Fn_ClassAdmin_CreateVerifyMapping("VerifyTarget","","1066","","")
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

Public Function Fn_ClassAdmin_CreateVerifyMapping(sAction,sSourceId,sTargetId,sButtons,sColumnName,bSave)
	GBL_FAILED_FUNCTION_NAME="Fn_ClassAdmin_CreateVerifyMapping"
Dim sItemId, sRevId, objClassAdminDialog,bFlag,sCellData,iRowCnt,iCount

	Fn_ClassAdmin_CreateVerifyMapping=false

	If JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").Exist(5) Then
		'Activate the View Attributes tab
        bReturn = Fn_ClassAdmin_TabOpeartions("Activate","Subtab","View Attributes","")
		
				Select Case sAction
					Case "Map"
						bFlag=False						
						iRowCnt = JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaTable("SourceClass").GetROProperty ("rows")
						For iCount = 0 to iRowCnt - 1
								sCellData=JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaTable("SourceClass").GetCellData(iCount,0)
								If (LCase( Trim(sSourceId) ) = LCase( Trim( sCellData ) )) Then
									bFlag=True
									JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaTable("SourceClass").ClickCell iCount,"0","LEFT"
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Selected the node [" +sSourceId+"] at the row position "+Cstr(iCount))
									Exit for
								End If
						Next
		
						iRowCnt = JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaTable("TargetClass").GetROProperty ("rows")
						If iRowCnt=1 Then
							JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaTable("TargetClass").ClickCell "0","0","LEFT"
						Else
								For iCount = 0 to iRowCnt - 1
										sCellData=JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaTable("TargetClass").GetCellData(iCount,1)
										If (LCase( Trim(sTargetId) ) = LCase( Trim( sCellData))) Then
											JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaTable("TargetClass").ClickCell iCount,"0","LEFT"
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Selected the node [" +sTargetId+"] at the row position "+Cstr(iCount))
											bFlag=True
											Exit for
										End If
								Next
					End If

						If bFlag=True Then
							'click on the Map Button
								JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaButton("Map").Click micLeftBtn
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Performed mapping operation")
							Fn_ClassAdmin_CreateVerifyMapping=True
						Else
						   Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Perform mapping operation")
							Fn_ClassAdmin_CreateVerifyMapping=False
						End If
	

				Case "VerifyTarget"

						bFlag=False
						iRowCnt = JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaTable("TargetClass").GetROProperty ("rows")
						If iRowCnt=1 Then
							sCellData=JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaTable("TargetClass").GetCellData(iCount,0)
							If (LCase( Trim("#"+sTargetId) ) = LCase( Trim( sCellData ) )) Then
									bFlag=True
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Verified that the mapping operation has been performed")
									Fn_ClassAdmin_CreateVerifyMapping=True
							Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to verify that the mapping operation has been performed")
									Fn_ClassAdmin_CreateVerifyMapping=False
							End If
						Else
						For iCount = 0 to iRowCnt-1
								sCellData=JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaTable("TargetClass").GetCellData(iCount,0)
								If (LCase( Trim("#"+sTargetId) ) = LCase( Trim( sCellData ) )) Then
									bFlag=True
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Verified that the mapping operation has been performed")
									Fn_ClassAdmin_CreateVerifyMapping=True
									Exit for
								End If
						Next
					End If
				End Select
					If bSave=True Then
						bReturn = Fn_ToolbatButtonClick("Save current Instance")
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Saved The Instance")
					End If
	End If
End Function
'''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''''/$$$$
''''/$$$$   FUNCTION NAME   : Fn_ClassAdmin_StaticTextRetrieve(sAction,sValue)
''''/$$$$
''''/$$$$   DESCRIPTION        :  This function will create mapping and verify if the mapping has happened
''''/$$$$
''''/$$$$    PARAMETERS      :   1.) sAction : Action to be performed
''''/$$$$ 									   2.) sValue : Value of the static text to be retrieved
''''/$$$$
''''/$$$$	Return Value : 			Value of the static text if it exists and False if the static text does not exist
''''/$$$$
''''/$$$$    Function Calls       :   Fn_WriteLogFile()
''''/$$$$										
''''/$$$$
''''/$$$$	HISTORY           :   AUTHOR                 DATE        VERSION
''''/$$$$
''''/$$$$    CREATED BY     :   SHREYAS           24/01/2011         1.0
''''/$$$$
''''/$$$$    REVIWED BY     :  Prasanna			      24/01/2011         1.0
''''/$$$$
''''/$$$$    EXAMPLE          :bReturn=Fn_ClassAdmin_StaticTextRetrieve("Retrieve","") 
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

Public Function Fn_ClassAdmin_StaticText(sAction,sValue)
	GBL_FAILED_FUNCTION_NAME="Fn_ClassAdmin_StaticText"
Dim sItemId, sTextValue, objClassAdmin,bFlag,objSelectType

	Dim objClassAdminApplet 
	Fn_ClassAdmin_StaticText=false
	
	Set objClassAdminApplet =  Fn_SISW_ClassAdmin_GetObject("ClassAdminApplet")
	if Fn_SISW_UI_Object_Operations("Fn_ClassAdmin_CreateAttribute","Exist", ObjClassAdminApplet,"") then
	'If JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").Exist(5) then
		Select Case sAction
			Case "Retrieve"
					bFlag = False
					Set objSelectType = description.Create()
					objSelectType("Class Name").value = "JavaStaticText"
					Set objClassAdmin =JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").ChildObjects(objSelectType)
					sTextValue=objClassAdmin(0).GetRoProperty("attached text")
					If instr(1,sTextValue,"ICM")>0 Then
						   bFlag = True
					End If
						 If bFlag=false Then
								Fn_ClassAdmin_StaticText = False
								Set objClassAdminApplet = nothing
								Exit function
						Else
								Fn_ClassAdmin_StaticText = sTextValue
						 End If

			Case "Exist"
			'(for future use)
		
			End Select
	End If
	Set objClassAdminApplet = nothing
End Function

''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''''/$$$$
''''/$$$$   FUNCTION NAME   : Fn_ClassAdmin_ClassSearchAndVerify(sAction,sSearchType,sSearchValue,sInfo,bClose)
''''/$$$$
''''/$$$$   DESCRIPTION        :  This function will create mapping and verify if the mapping has happened
''''/$$$$
''''/$$$$    PARAMETERS      :   1.) sAction : Action To Be Performed
''''/$$$$ 									   2.) sSearchType : Search Type To Be Selected
''''/$$$$										3.) sSearchValue : Search Value to be entered
''''/$$$$										4.) sInfo : For Future USe
''''/$$$$										5.) bClose : Boolean parameter for closing the dialog
''''/$$$$
''''/$$$$	Return Value : 			Value of the static text if it exists and False if the static text does not exist
''''/$$$$
''''/$$$$    Function Calls       :   Fn_WriteLogFile()
''''/$$$$										
''''/$$$$
''''/$$$$	HISTORY           :   AUTHOR                 DATE        VERSION
''''/$$$$
''''/$$$$    CREATED BY     :   SHREYAS           24/01/2011         1.0
''''/$$$$
''''/$$$$    REVIWED BY     :  Prasanna			      24/01/2011         1.0
''''/$$$$
''''/$$$$    EXAMPLE          :bReturn=Fn_ClassAdmin_ClassSearchAndVerify("Search","Name","Target_36388","","")
''''/$$$$      							  bReturn=Fn_ClassAdmin_ClassSearchAndVerify("DoubleClick","","Target_36388","",True)
''''/$$$$    							  bReturn=Fn_ClassAdmin_ClassSearchAndVerify("Search","Name","t*","","")
''''/$$$$    							  bReturn=Fn_ClassAdmin_ClassSearchAndVerify("Verify","","Target_36388,TC Classification Root","",True)
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Public Function Fn_ClassAdmin_ClassSearchAndVerify(sAction,sSearchType,sSearchValue,sInfo,bClose)
	GBL_FAILED_FUNCTION_NAME="Fn_ClassAdmin_ClassSearchAndVerify"
   Dim sItemId,  objClassAdmin,objSelectType,sTextValue,sStaticCount

	Fn_ClassAdmin_ClassSearchAndVerify=false

	If JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").Exist(5) Then


Select Case sAction

			Case "Search"

		   			'Invoke the Search type select window and convert it into a dialog 
					JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaCheckBox("SearchCriteria").Set "ON"
				
					Set objSelectType = description.Create()
					objSelectType("Class Name").value = "JavaStaticText"
					Set objClassAdmin =JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").ChildObjects(objSelectType)
					sStaticCount=JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").ChildObjects(objSelectType).count
					For iCount=0 to sStaticCount-1
								sTextValue=objClassAdmin(iCount).GetRoProperty("attached text")
								If sTextValue="Search Class ..." Then
									objClassAdmin(iCount).DblClick 5,5,"LEFT"
									Exit for
								End If
				    Next

					'set the value in the SearchType list box
					JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("Search Class").JavaList("SearchType").Type sSearchType
									 If Err.Number < 0 Then
												Fn_ClassAdmin_ClassSearchAndVerify = False
												Exit Function 
										End If
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Set the value [" +sSearchType+"] in the Search Type Select List")
		
					'set the value for the class to be searched
					JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("Search Class").JavaEdit("SearchDetails").Set sSearchValue
									 If Err.Number < 0 Then
												Fn_ClassAdmin_ClassSearchAndVerify = False
												Exit Function 
										End If
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Set the value [" +sSearchValue+"] in the Search class Select List")
		
					'click on the search button
					JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("Search Class").JavaButton("Search").Click micLeftBtn
									 If Err.Number < 0 Then
												Fn_ClassAdmin_ClassSearchAndVerify = False
												Exit Function 
										End If
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully clicked on the search button")

					If bCLose=True Then
									'Close the dialog
							JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("Search Class").Close
					End If

					Fn_ClassAdmin_ClassSearchAndVerify = True

			Case "DoubleClick"

						bFlag=False
					'Select the value from the SearchResult JavaList and double click on it
					sItemCount=JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("Search Class").JavaList("ResultList").GetROProperty("items count")
					For iCount=0 to sItemCount-1
									sValue=	JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("Search Class").JavaList("ResultList").GetItem(iCount)
									If instr(1,sValue, sSearchValue)>1Then
										JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("Search Class").JavaList("ResultList").DblClick 5,5,"LEFT"
										 If Err.Number < 0 Then
												Fn_ClassAdmin_ClassSearchAndVerify = False
												Exit Function 
										End If
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully double clicked on the node [" +sValue+"] in the Search ResultList")
										bFlag=True
										Exit For
								End If
					Next
		
								Fn_ClassAdmin_ClassSearchAndVerify = True
					
						If bCLose=True Then
								'Close the dialog
						JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("Search Class").Close
						End If

					Case "Verify"
						If instr(1,sSearchValue,",")>1 Then
						aProperties=split(sSearchValue,",",-1,1)
						
						For iCounter=0 To Ubound(aProperties)
							bFlag = False
									sItemCount=JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("Search Class").JavaList("ResultList").GetROProperty("items count")
									For iCount = 0 to sItemCount - 1
										sValue=	JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("Search Class").JavaList("ResultList").GetItem(iCount)
										If instr(1,sValue, aProperties(iCounter))>1Then
												bFlag=True
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Verified that the node " + aProperties(iCounter)+" exists in the list")
												Exit For
											End If
										 Next
									If  bFlag = False Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Failed to verify that the node " + aProperties(iCounter)+" exists in the list.")
											'Close the dialog
											JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("Search Class").Close
										Exit For
									End If
					Next
						End If

						If bCLose=True Then
								'Close the dialog
						JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("Search Class").Close
						End If
						Fn_ClassAdmin_ClassSearchAndVerify = True
	End select
End if

End Function

'''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'''''/$$$$
'''''/$$$$   FUNCTION NAME   : Fn_ClassAdmin_ConfigureReferencerAttribute(sAction,sPOMTreeNode,sTab,sListValue,sButton,sInfo)
'''''/$$$$
'''''/$$$$   DESCRIPTION        :  This function will perform operations
'''''/$$$$
'''''/$$$$    PARAMETERS      :   1.) sAction : Action to be performed
'''''/$$$$ 									    2.) sPOMTreeNode : Node to be selected from the tree
'''''/$$$$									   3.) sTab : POM attribute or Type Property tab to be selected
'''''/$$$$									  4.) sListValue : Value to be selected from the List
'''''/$$$$									 5.) sButton : Button to be clicked (OK or Cancel)
'''''/$$$$									 6.) sInfo : Extra parameter for future use
'''''/$$$$								
'''''/$$$$
'''''/$$$$	Return Value : 			True / False
'''''/$$$$
'''''/$$$$    Function Calls       :   Fn_WriteLogFile()
'''''/$$$$										
'''''/$$$$
'''''/$$$$	HISTORY           :   AUTHOR                 DATE        VERSION
'''''/$$$$
'''''/$$$$    CREATED BY     :   SHREYAS           	31/01/2011         1.0
'''''/$$$$
'''''/$$$$    REVIWED BY     :  Prasanna			      31/01/2011         1.0
'''''/$$$$
'''''/$$$$    EXAMPLE          :  bReturn=Fn_ClassAdmin_ConfigureReferencerAttribute("RelatedFromAttribute","Types:Form Type:ItemRevision Master","Type Property","Name","Ok","")
'''''/$$$$ 									bReturn=Fn_ClassAdmin_ConfigureReferencerAttribute("MasterFromAttribute","Types:Item Type:Item","POM Attribute","serial_number","Ok","")
'''''/$$$$									
'''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

Public Function Fn_ClassAdmin_ConfigureReferencerAttribute(sAction,sPOMTreeNode,sTab,sListValue,sButton,sInfo)
	GBL_FAILED_FUNCTION_NAME="Fn_ClassAdmin_ConfigureReferencerAttribute"

   Dim  aProperties,bFlag
   Fn_ClassAdmin_ConfigureReferencerAttribute=false
	Err.clear
'first check the reference attribute checkbox to true..

If JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaCheckBox("ReferenceAttribute").GetROProperty("value")=0 then
			JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaCheckBox("ReferenceAttribute").Set "ON"
										If Err.Number < 0 Then
												Fn_ClassAdmin_ConfigureReferencerAttribute = False
												Exit Function 
										End If
			
			 'Click on the Configure button
			 JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaButton("Configure").Click micLeftBtn

									 If Err.Number < 0 Then
												Fn_ClassAdmin_ConfigureReferencerAttribute = False
												Exit Function 
										End If
			 Call Fn_ReadyStatusSync(5)
			 wait(2)
End If

'check the existence of the 'configure referencer attribute window and continue if exists

If JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("ConfigureReferenceAttribute").Exist(5) Then


		Select Case sAction


				Case "RelatedFromAttribute"
				bFlag=False
					'check the RelatedFromAttribute checkbox if not checked
					 If JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("ConfigureReferenceAttribute").JavaRadioButton("RelatedFormAttribute").GetROProperty("value")=0 Then
						JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("ConfigureReferenceAttribute").JavaRadioButton("RelatedFormAttribute").Set "ON"
										 If Err.Number < 0 Then
												Fn_ClassAdmin_ConfigureReferencerAttribute = False
												Exit Function 
										End If
										 Call Fn_ReadyStatusSync(3)
					End If

					'Select the node from the POM tree

													aProperties=split(sPOMTreeNode,":",-1,1)

												'Expand the node of the path specified
												JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("ConfigureReferenceAttribute").JavaTree("POMTree").Expand "Types:"+aProperties(1) 
												 Call Fn_ReadyStatusSync(5)
														JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("ConfigureReferenceAttribute").JavaTree("POMTree").Select (sPOMTreeNode)
														If Err.Number < 0 Then
																Fn_ClassAdmin_ConfigureReferencerAttribute = False
																Exit Function 
														End If
														bFlag=True
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully selected the node {"+sPOMTreeNode+"} in the POM Tree at the path ")
	

					'activate the specified tab
										JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("ConfigureReferenceAttribute").JavaTab("POMTab").Select sTab
										If Err.Number < 0 Then
												Fn_ClassAdmin_ConfigureReferencerAttribute = False
												Exit Function 
										End If
									   Call Fn_ReadyStatusSync(3)

					'Select the value from the values list
									JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("ConfigureReferenceAttribute").JavaList("ValueList").Select sListValue
									If Err.Number < 0 Then
												Fn_ClassAdmin_ConfigureReferencerAttribute = False
												Exit Function 
									End If
			
								If bFlag=True Then
									Fn_ClassAdmin_ConfigureReferencerAttribute=True
								 Else
									Fn_ClassAdmin_ConfigureReferencerAttribute=False
								End If

							Case "ClassifiedObject"

								'Two different trees are seen in both the tabs,hence select case is used

								Select Case sTab

								Case "POM Attribute"

										'activate the specified tab
										JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("ConfigureReferenceAttribute").JavaTab("POMTab").Select sTab
										If Err.Number < 0 Then
												Fn_ClassAdmin_ConfigureReferencerAttribute = False
												Exit Function 
										End If
										Call Fn_ReadyStatusSync(3)
												
															bFlag=False
															'check the ClassifiedObject checkbox if not checked
															JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("ConfigureReferenceAttribute").JavaRadioButton("RelatedFormAttribute").SetTOProperty "attached text","Classified Object"
															if JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("ConfigureReferenceAttribute").JavaRadioButton("RelatedFormAttribute").Exist(10) Then
																	 If JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("ConfigureReferenceAttribute").JavaRadioButton("RelatedFormAttribute").GetROProperty("value")=0 Then
																		JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("ConfigureReferenceAttribute").JavaRadioButton("RelatedFormAttribute").Set "ON"
																						 If Err.Number < 0 Then
																								Fn_ClassAdmin_ConfigureReferencerAttribute = False
																								Exit Function 
																						End If
																						 Call Fn_ReadyStatusSync(3)
																	End If
															End if
					
																		'Select the node from the POM tree
																		aProperties=split(sPOMTreeNode,":",-1,1)
					
															'Expand the node of the path specified
																JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("ConfigureReferenceAttribute").JavaTree("POMTree").Activate aProperties(0)
																 Call Fn_ReadyStatusSync(5)
																		JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("ConfigureReferenceAttribute").JavaTree("POMTree").Select (sPOMTreeNode)
																		If Err.Number < 0 Then
																				Fn_ClassAdmin_ConfigureReferencerAttribute = False
																				Exit Function 
																		End If
																		bFlag=True
																		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully selected the node {"+sPOMTreeNode+"} in the POM Tree at the path ")
						
					
														'Select the value from the values list
														JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("ConfigureReferenceAttribute").JavaList("ValueList").Select sListValue
														If Err.Number < 0 Then
																	Fn_ClassAdmin_ConfigureReferencerAttribute = False
																	Exit Function 
														End If

														'set the checkbox to on
														aValues=split(sInfo,":",-1,1)
														JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("ConfigureReferenceAttribute").JavaCheckBox("SelectAttributeFrom").SetTOProperty "attached text",aValues(0)
														JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("ConfigureReferenceAttribute").JavaCheckBox("SelectAttributeFrom").Set ucase(aValues(1))
														If Err.Number < 0 Then
																	Fn_ClassAdmin_ConfigureReferencerAttribute = False
																	Exit Function 
														End If

					
														If bFlag=True Then
															Fn_ClassAdmin_ConfigureReferencerAttribute=True
														 Else
															Fn_ClassAdmin_ConfigureReferencerAttribute=False
														End If
						
									Case "Type Property"

										bFlag=False
															'check the ClassifiedObject checkbox if not checked
															JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("ConfigureReferenceAttribute").JavaRadioButton("RelatedFormAttribute").SetTOProperty "attached text","Classified Object"
															JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("ConfigureReferenceAttribute").JavaRadioButton("RelatedFormAttribute").highlight
					
															 If JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("ConfigureReferenceAttribute").JavaRadioButton("RelatedFormAttribute").GetROProperty("value")=0 Then
																JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("ConfigureReferenceAttribute").JavaRadioButton("RelatedFormAttribute").Set "ON"
																				 If Err.Number < 0 Then
																						Fn_ClassAdmin_ConfigureReferencerAttribute = False
																						Exit Function 
																				End If
																				 Call Fn_ReadyStatusSync(3)
															End If
					
															'activate the specified tab
															JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("ConfigureReferenceAttribute").JavaTab("POMTab").Select sTab
															If Err.Number < 0 Then
																	Fn_ClassAdmin_ConfigureReferencerAttribute = False
																	Exit Function 
															End If
															Call Fn_ReadyStatusSync(3)
					
																		aProperties=split(sPOMTreeNode,":",-1,1)
					
																'Expand the node of the path specified
																JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("ConfigureReferenceAttribute").JavaTree("POMTree").Expand "Types:"+aProperties(1) 
																 Call Fn_ReadyStatusSync(5)
																		JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("ConfigureReferenceAttribute").JavaTree("POMTree").Select (sPOMTreeNode)
																		If Err.Number < 0 Then
																				Fn_ClassAdmin_ConfigureReferencerAttribute = False
																				Exit Function 
																		End If
																		bFlag=True
																		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully selected the node {"+sPOMTreeNode+"} in the POM Tree at the path ")
					
					
																'Select the value from the values list
																	JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("ConfigureReferenceAttribute").JavaList("ValueList").Select sListValue
																	If Err.Number < 0 Then
																				Fn_ClassAdmin_ConfigureReferencerAttribute = False
																				Exit Function 
																		End If
						
																If bFlag=True Then
																	Fn_ClassAdmin_ConfigureReferencerAttribute=True
																 Else
																	Fn_ClassAdmin_ConfigureReferencerAttribute=False
																End If
					
									End Select

						Case "Related Object"

							Select Case sTab

								Case "POM Attribute"
									aListValues=split(sListValue,":",-1,1)

										'activate the specified tab
										JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("ConfigureReferenceAttribute").JavaTab("POMTab").Select sTab
										If Err.Number < 0 Then
												Fn_ClassAdmin_ConfigureReferencerAttribute = False
												Exit Function 
										End If
										Call Fn_ReadyStatusSync(3)
												
															bFlag=False
															'check the RelatedObject checkbox if not checked
															JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("ConfigureReferenceAttribute").JavaRadioButton("RelatedFormAttribute").SetTOProperty "attached text","Related Object"
															JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("ConfigureReferenceAttribute").JavaRadioButton("RelatedFormAttribute").highlight
					
															 If JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("ConfigureReferenceAttribute").JavaRadioButton("RelatedFormAttribute").GetROProperty("value")=0 Then
																JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("ConfigureReferenceAttribute").JavaRadioButton("RelatedFormAttribute").Set "ON"
																				 If Err.Number < 0 Then
																						Fn_ClassAdmin_ConfigureReferencerAttribute = False
																						Exit Function 
																				End If
																				 Call Fn_ReadyStatusSync(3)
															End If
					
																		'Select the node from the POM tree
																		aProperties=split(sPOMTreeNode,":",-1,1)
					
															'Expand the node of the path specified
																JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("ConfigureReferenceAttribute").JavaTree("POMTree").Activate aProperties(0)
																 Call Fn_ReadyStatusSync(5)
																		JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("ConfigureReferenceAttribute").JavaTree("POMTree").Select (sPOMTreeNode)
																		If Err.Number < 0 Then
																				Fn_ClassAdmin_ConfigureReferencerAttribute = False
																				Exit Function 
																		End If
																		bFlag=True
																		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully selected the node {"+sPOMTreeNode+"} in the POM Tree at the path ")
						
																	'Select the value from the properties list
																	JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("ConfigureReferenceAttribute").JavaList("ValueList").Select aListValues(0)
																	If Err.Number < 0 Then
																				Fn_ClassAdmin_ConfigureReferencerAttribute = False
																				Exit Function 
																		End If
					
					
																'Select the value from the values list
																	JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("ConfigureReferenceAttribute").JavaList("ValueList").Select aListValues(1)
																	If Err.Number < 0 Then
																				Fn_ClassAdmin_ConfigureReferencerAttribute = False
																				Exit Function 
																		End If
					
														If bFlag=True Then
															Fn_ClassAdmin_ConfigureReferencerAttribute=True
														 Else
															Fn_ClassAdmin_ConfigureReferencerAttribute=False
														End If
						
									Case "Type Property"
										aListValues=split(sListValue,":",-1,1)
										bFlag=False
															'check the RelatedObject checkbox if not checked
															JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("ConfigureReferenceAttribute").JavaRadioButton("RelatedFormAttribute").SetTOProperty "attached text","Related Object"
															JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("ConfigureReferenceAttribute").JavaRadioButton("RelatedFormAttribute").highlight
					
															 If JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("ConfigureReferenceAttribute").JavaRadioButton("RelatedFormAttribute").GetROProperty("value")=0 Then
																JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("ConfigureReferenceAttribute").JavaRadioButton("RelatedFormAttribute").Set "ON"
																				 If Err.Number < 0 Then
																						Fn_ClassAdmin_ConfigureReferencerAttribute = False
																						Exit Function 
																				End If
																				 Call Fn_ReadyStatusSync(3)
															End If
					
															'activate the specified tab
															JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("ConfigureReferenceAttribute").JavaTab("POMTab").Select sTab
															If Err.Number < 0 Then
																	Fn_ClassAdmin_ConfigureReferencerAttribute = False
																	Exit Function 
															End If
															Call Fn_ReadyStatusSync(3)
					
																		aProperties=split(sPOMTreeNode,":",-1,1)
					
																'Expand the node of the path specified
																JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("ConfigureReferenceAttribute").JavaTree("POMTree").Expand "Types:"+aProperties(1) 
																 Call Fn_ReadyStatusSync(5)
																		JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("ConfigureReferenceAttribute").JavaTree("POMTree").Select (sPOMTreeNode)
																		If Err.Number < 0 Then
																				Fn_ClassAdmin_ConfigureReferencerAttribute = False
																				Exit Function 
																		End If
																		bFlag=True
																		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully selected the node {"+sPOMTreeNode+"} in the POM Tree at the path ")

																		'Select the value from the properties list
																	JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("ConfigureReferenceAttribute").JavaList("ValueList").Select aListValues(0)
																	If Err.Number < 0 Then
																				Fn_ClassAdmin_ConfigureReferencerAttribute = False
																				Exit Function 
																		End If
					
					
																'Select the value from the values list
																	JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("ConfigureReferenceAttribute").JavaList("Relations").Select aListValues(1)
																	If Err.Number < 0 Then
																				Fn_ClassAdmin_ConfigureReferencerAttribute = False
																				Exit Function 
																		End If
						
																If bFlag=True Then
																	Fn_ClassAdmin_ConfigureReferencerAttribute=True
																 Else
																	Fn_ClassAdmin_ConfigureReferencerAttribute=False
																End If
					
									End Select

									Case "MasterFromAttribute"
										'check the RelatedObject checkbox if not checked
															JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("ConfigureReferenceAttribute").JavaRadioButton("RelatedFormAttribute").SetTOProperty "attached text","Masterform Attribute"
															JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("ConfigureReferenceAttribute").JavaRadioButton("RelatedFormAttribute").highlight
				bFlag=False
					'check the RelatedFromAttribute checkbox if not checked
					 If JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("ConfigureReferenceAttribute").JavaRadioButton("RelatedFormAttribute").GetROProperty("value")=0 Then
						JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("ConfigureReferenceAttribute").JavaRadioButton("RelatedFormAttribute").Set "ON"
										 If Err.Number < 0 Then
												Fn_ClassAdmin_ConfigureReferencerAttribute = False
												Exit Function 
										End If
										 Call Fn_ReadyStatusSync(3)
					End If

					'Select the node from the POM tree

													aProperties=split(sPOMTreeNode,":",-1,1)

												'Expand the node of the path specified
												JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("ConfigureReferenceAttribute").JavaTree("POMTree").Expand "Types:"+aProperties(1) 
												 Call Fn_ReadyStatusSync(5)
														JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("ConfigureReferenceAttribute").JavaTree("POMTree").Select (sPOMTreeNode)
														If Err.Number < 0 Then
																Fn_ClassAdmin_ConfigureReferencerAttribute = False
																Exit Function 
														End If
														bFlag=True
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully selected the node {"+sPOMTreeNode+"} in the POM Tree at the path ")
	

					'activate the specified tab
										JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("ConfigureReferenceAttribute").JavaTab("POMTab").Select sTab
										If Err.Number < 0 Then
												Fn_ClassAdmin_ConfigureReferencerAttribute = False
												Exit Function 
										End If
									   Call Fn_ReadyStatusSync(3)

					'Select the value from the values list
									JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("ConfigureReferenceAttribute").JavaList("ValueList").Select sListValue
									If Err.Number < 0 Then
												Fn_ClassAdmin_ConfigureReferencerAttribute = False
												Exit Function 
									End If
			
								If bFlag=True Then
									Fn_ClassAdmin_ConfigureReferencerAttribute=True
								 Else
									Fn_ClassAdmin_ConfigureReferencerAttribute=False
								End If


		End Select

			If lcase(sButton)= "ok" Then
				JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("ConfigureReferenceAttribute").JavaButton("OK").Click micLeftBtn
				 Call Fn_ReadyStatusSync(3)
										If Err.Number < 0 Then
												Fn_ClassAdmin_ConfigureReferencerAttribute = False
												Exit Function 
										End If
			Elseif lcase(sButton)= "cancel" Then
				JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("ConfigureReferenceAttribute").JavaButton("Cancel").Click micLeftBtn
										If Err.Number < 0 Then
												Fn_ClassAdmin_ConfigureReferencerAttribute = False
												Exit Function 
										End If
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The button is not needed to be clicked ")
			End If

		End If
End function

''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''''/$$$$
''''/$$$$   FUNCTION NAME   : Fn_ClassAdmin_XMLImportExport(sAction,sTab,sTargetApplication,sOutputFile,sInfo,sButton)
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
''''/$$$$	Return Value : 			To perform export operation in the ClassAdmin Application
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
''''/$$$$    EXAMPLE          :  
''''/$$$$									
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

Public Function Fn_ClassAdmin_XMLImportExport(sAction,sTab,sTargetApplication,sOutputFile,sInfo,sButton)
	GBL_FAILED_FUNCTION_NAME="Fn_ClassAdmin_XMLImportExport"
'
   Dim WshShell, iCounter, aProperties,bFlag,iRowCount,iCount,jCount,iColCount,strLine,sStaticCount,aCheck,aDate,sMonthName
   Fn_ClassAdmin_XMLImportExport=false

		 

				Select Case sAction

								Case "Export"
											   'check the existence of the XML export dialog
											 If   JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("XML Export").Exist(5) =False Then
														bReturn = Fn_ToolbatButtonClick("Export Objects")
														If Err.Number < 0 Then
																	Fn_ClassAdmin_XMLImportExport = False
																	JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("XML Export").JavaButton("Cancel").Click micLeftBtn
																	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to select tab [ " + sTab + "] of XML Export dialog" ) 
																	Exit Function 
														End if
														Call Fn_ReadyStatusSync(3)
											 End if
		
											bFlag=False
						
											'activate the specified tab
						
											 JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("XML Export").JavaTab("Tab").Select sTab
						
															  If Err.Number < 0 Then
																	Fn_ClassAdmin_XMLImportExport = False
																	JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("XML Export").JavaButton("Cancel").Click micLeftBtn
																	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to select tab [ " + sTab + "] of XML Export dialog" ) 
																	Exit Function 
															End if
						
											'slect the value from the target application list
						
											 JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("XML Export").JavaList("TargetApplication").Select sTargetApplication
						
															If Err.Number < 0 Then
																	Fn_ClassAdmin_XMLImportExport = False
																	JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("XML Export").JavaButton("Cancel").Click micLeftBtn
																	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to select tab [ " + sTab + "] of XML Export dialog" ) 
																	Exit Function 
															End if
						
											'Set the output file
						
											 JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("XML Export").JavaEdit("Output File").Set sOutputFile
						
															If Err.Number < 0 Then
																	Fn_ClassAdmin_XMLImportExport = False
																	JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("XML Export").JavaButton("Cancel").Click micLeftBtn
																	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to select tab [ " + sTab + "] of XML Export dialog" ) 
																	Exit Function 
															End if

												If lcase(sButton)="export" Then
														   JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("XML Export").JavaButton("OK").SetTOProperty "label","Export"	
														 JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("XML Export").JavaButton("OK").Click micLeftBtn
														If Err.Number < 0 Then
																Fn_ClassAdmin_XMLImportExport = False
																Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to select tab [ " + sTab + "] of XML Import dialog" ) 
																Exit Function 
														End if
											Elseif lcase(sButton)="cancel" Then
														 JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("XML Export").JavaButton("Cancel").Click micLeftBtn
														If Err.Number < 0 Then
																Fn_Classification_XMLExport = False
																Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to select tab [ " + sTab + "] of XML Import dialog" ) 
																Exit Function 
														End if
											Else
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "No Button Needs To Be Clicked" ) 
											End If
											

												'handle the Export successful Dialog by clicking Ok button
											 JavaWindow("ClassAdminMainWin").JavaWindow("Export").JavaButton("OK").Click micLeftBtn
											If Err.Number < 0 Then
															Fn_ClassAdmin_XMLImportExport = False
															Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to select tab [ " + sTab + "] of XML Import dialog" ) 
															Exit Function 
											End if
											bFlag=True
							Case "Import"
											  JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("XML Import").SetTOProperty "title","XML Import"
											   'check the existence of the XML export dialog
											 If   JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("XML Import").Exist(5) =False Then
														bReturn = Fn_ToolbatButtonClick("Import Objects")
														If Err.Number < 0 Then
																	Fn_ClassAdmin_XMLImportExport = False
																	JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("XML Import").JavaButton("Cancel").Click micLeftBtn
																	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to select tab [ " + sTab + "] of XML Import dialog" ) 
																	Exit Function 
														End if
														Call Fn_ReadyStatusSync(3)
											 End if
		
											bFlag=False
						
											'activate the specified tab
						
											 JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("XML Import").JavaTab("Tab").Select sTab
						
															  If Err.Number < 0 Then
																	Fn_ClassAdmin_XMLImportExport = False
																	JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("XML Import").JavaButton("Cancel").Click micLeftBtn
																	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to select tab [ " + sTab + "] of XML Import dialog" ) 
																	Exit Function 
															End if
						
											'slect the value from the target application list
											  JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("XML Import").JavaList("TransferMode").SetTOProperty "attached text","Transfer Mode Name"	
											 JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("XML Import").JavaList("TransferMode").Select sTargetApplication
						
															If Err.Number < 0 Then
																	Fn_ClassAdmin_XMLImportExport = False
																	JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("XML Import").JavaButton("Cancel").Click micLeftBtn
																	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to select tab [ " + sTab + "] of XML Import dialog" ) 
																	Exit Function 
															End if
						
											'Set the output file
											JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("XML Import").JavaEdit("InputFile").SetTOProperty "attached text","Input File"	
											 JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("XML Import").JavaEdit("InputFile").Set sOutputFile
						
															If Err.Number < 0 Then
																	Fn_ClassAdmin_XMLImportExport = False
																	JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("XML Import").JavaButton("Cancel").Click micLeftBtn
																	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to select tab [ " + sTab + "] of XML Import dialog" ) 
																	Exit Function 
															End if
											
														
											If lcase(sButton)="import" Then
														   JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("XML Import").JavaButton("Import").SetTOProperty "label","Import"	
														 JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("XML Import").JavaButton("Import").Click micLeftBtn
														If Err.Number < 0 Then
																Fn_ClassAdmin_XMLImportExport = False
																Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to select tab [ " + sTab + "] of XML Import dialog" ) 
																Exit Function 
														End if
											Elseif lcase(sButton)="cancel" Then
														 JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("XML Import").JavaButton("Cancel").Click micLeftBtn
														If Err.Number < 0 Then
																Fn_Classification_XMLExport = False
																Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to select tab [ " + sTab + "] of XML Import dialog" ) 
																Exit Function 
														End if
											Else
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "No Button Needs To Be Clicked" ) 
											End If

											JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("Import Completed").JavaButton("No").Click micLeftBtn
											If Err.Number < 0 Then
																Fn_ClassAdmin_XMLImportExport = False
																Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Click on [ No ] button of Import dialog" ) 
																Exit Function 
											End if
											bFlag = true

				End Select

				
			

				If bFlag=True Then
					Fn_ClassAdmin_XMLImportExport=True
				Else
					Fn_ClassAdmin_XMLImportExport=False
				End If

		

End function

''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''''/$$$$
''''/$$$$   FUNCTION NAME   : Fn_ClassAdmin_XMLExport(sAction,sTab,sTargetApplication,sOutputFile,sInfo,sButton)
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
''''/$$$$    EXAMPLE          :  
''''/$$$$									
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

Public Function Fn_ClassAdmin_XMLExport(sAction,sTab,sTargetApplication,sOutputFile,aChbDetails,sButton)
	GBL_FAILED_FUNCTION_NAME="Fn_ClassAdmin_XMLExport"
'
   Dim WshShell, iCounter, aProperties,bFlag,iRowCount,iCount,jCount,iColCount,strLine,sStaticCount,aCheck,aDate,sMonthName,sText
   Fn_ClassAdmin_XMLExport=false
	Err.Clear
   'check the existence of the XML export dialog

   If JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("XML Export").Exist(5) =False Then
	   bReturn = Fn_ToolbatButtonClick("Export Objects")
						 If Err.Number < 0 Then
									Fn_ClassAdmin_XMLExport = False
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to select tab [ " + sTab + "] of XML Import dialog" ) 
									Exit Function 
							End if
			Call Fn_ReadyStatusSync(3)
  End if


		 If  JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("XML Export").Exist(5) Then

				Select Case sAction

								Case "Export"
									bFlag=False
						
											'activate the specified tab
											if sTab<>"" Then
													JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("XML Export").JavaTab("Tab").Select sTab
								
																	  If Err.Number < 0 Then
																			Fn_ClassAdmin_XMLExport = False
																			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to select tab [ " + sTab + "] of XML Import dialog" ) 
																			Exit Function 
																	End if
											End if

											'slect the value from the target application list
											if sTargetApplication<>"" Then
															JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("XML Export").JavaList("TargetApplication").Select sTargetApplication
						
																		If Err.Number < 0 Then
																				Fn_ClassAdmin_XMLExport = False
																				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to select tab [ " + sTab + "] of XML Import dialog" ) 
																				Exit Function 
																		End if
											End if
											'Set the output file
											if sOutputFile<>"" Then
													JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("XML Export").JavaEdit("Output File").Set sOutputFile
								
																	If Err.Number < 0 Then
																			Fn_ClassAdmin_XMLExport = False
																			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to select tab [ " + sTab + "] of XML Import dialog" ) 
																			Exit Function 
																	End if
											  End if
											bFlag=True

							'Click on More Options.
								If sTab<>"PLMXML" Then
											JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("XML Export").JavaButton("OK").SetTOProperty "label","More Options >>"
											 'JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("XML Export").JavaButton("OK").highlight
											If JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("XML Export").JavaButton("OK").Exist(3)=true Then
												JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("XML Export").JavaButton("OK").Click micLeftBtn
			
														'Set the check box value
														'First make all check box OFF
													Set objSelectType = description.Create()
													objSelectType("Class Name").value = "JavaCheckBox"
													Set objClassAdmin =JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("XML Export").ChildObjects(objSelectType)
													sStaticCount=JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("XML Export").ChildObjects(objSelectType).count
													For iCount=0 to sStaticCount-1
																sText=objClassAdmin(iCount).GetRoProperty("attached text")
																If sText="notpinned_16" Then
																	Exit for
																End if
																If objClassAdmin(iCount).GetRoProperty("enabled")="1" Then
																	objClassAdmin(iCount).set "OFF"
																end if
													  Next
												End if
			
															'Select the specific check box which is to be made ON
														If IsArray(aChbDetails) Then
																For iCounter = 0 to ubound(aChbDetails)
																		JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("XML Export").JavaCheckBox("ChbOptions").SetTOProperty "attached text",aChbDetails(iCounter)
																		JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("XML Export").JavaCheckBox("ChbOptions").Set "ON"
																		If Err.Number < 0 Then
																					Fn_ClassAdmin_XMLExport = False
																					JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("XML Export").JavaButton("Cancel").Click micLeftBtn 
																					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to turn off the checkbox ["+aChbDetails(iCounter)+"]" ) 
																					Exit Function 
																		Else
																					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully turned off the checkbox ["+aChbDetails(iCounter)+"]" ) 	
																		End If
																Next									
														End If
								end if					

				End Select

if sButton<>"" Then
	if sButton<>"" Then
				If lcase(sButton)="export" Then
					JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("XML Export").JavaButton("OK").SetTOProperty "label","Export"
						'JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("XML Export").JavaButton("OK").highlight
						JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("XML Export").JavaButton("OK").Click micLeftBtn
						If Err.Number < 0 Then
								Fn_ClassAdmin_XMLExport = False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to select tab [ " + sTab + "] of XML Import dialog" ) 
								Exit Function 
						End if
				Elseif lcase(sButton)="cancel" Then
						JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("XML Export").JavaButton("Cancel").Click micLeftBtn
						If Err.Number < 0 Then
								Fn_ClassAdmin_XMLExport = False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to select tab [ " + sTab + "] of XML Import dialog" ) 
								Exit Function 
						End if
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "No Button Needs To Be Clicked" ) 
				End If

				'handle the Export successful Dialog by clicking Ok button
				
				JavaWindow("ClassAdminMainWin").JavaWindow("Export ICO").SetTOProperty "title","Export"
				'JavaWindow("ClassAdminMainWin").JavaWindow("Export ICO").highlight
				JavaWindow("ClassAdminMainWin").JavaWindow("Export ICO").JavaButton("OK").Click micLeftBtn
				If Err.Number < 0 Then
								Fn_ClassAdmin_XMLExport = False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to click on OK button of XML Export dialog" ) 
								Exit Function 
				End if
	End if
End if

				If bFlag=True Then
					Fn_ClassAdmin_XMLExport=True
				Else
					Fn_ClassAdmin_XMLExport=False
				End If

		End if

End function



''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''/$$$$
''/$$$$   FUNCTION NAME   :  Fn_ClassAdmin_CheckBoxSet(objCheckbox,sCheckBox,sSetValue)
''/$$$$
''/$$$$   DESCRIPTION        :  Check the checkbox to ON
''/$$$$
''/$$$$    PARAMETERS      :   1.) sCheckBox : Valid CheckBox name
''/$$$$                                       2.) sSetValue : To Set On or OFF
''/$$$$
''/$$$$	   Return Value : 			True / False
''/$$$$
''/$$$$    Function Calls       :   Fn_WriteLogFile()
''/$$$$									  
''/$$$$
''/$$$$	 HISTORY           :   AUTHOR                 DATE        VERSION
''/$$$$
''/$$$$    CREATED BY     :   SHREYAS           11/04/2011         1.0
''/$$$$
''/$$$$    REVIWED BY     :  Shreyas			11/04/2011        1.0
''/$$$$
''/$$$$    EXAMPLE          : 	bReturn=Fn_ClassAdmin_CheckBoxSet("Sub-Class Instances","ON")
''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

Public Function Fn_ClassAdmin_CheckBoxSet(sCheckBox,sSetValue)
	GBL_FAILED_FUNCTION_NAME="Fn_ClassAdmin_CheckBoxSet"
  	Fn_ClassAdmin_CheckBoxSet=false	

   Dim sText, objClassAdmin,bFlag,sCellData,iCount,sStaticCount,objExport,objImport,aSetValue,bClickFlag
   Set objExport= JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("XML Export")
    


	
  If objExport.Exist(3) Then
     
					objExport.JavaButton("OK").SetTOProperty "label","More Options >>"
					  'objExport.JavaButton("OK").highlight
						If objExport.JavaButton("OK").Exist(3)=true Then
							objExport.JavaButton("OK").Click micLeftBtn
							Set objSelectType = description.Create()
										objSelectType("Class Name").value = "JavaCheckBox"
										Set objClassAdmin =JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("XML Export").ChildObjects(objSelectType)
										sStaticCount=JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("XML Export").ChildObjects(objSelectType).count
										For iCount=0 to sStaticCount-1
											sText=objClassAdmin(iCount).GetRoProperty("attached text")
											If sText="notpinned_16" Then
												Exit for
											End if
											If objClassAdmin(iCount).GetRoProperty("enabled")="1" Then
											objClassAdmin(iCount).set "OFF"
													If sText=sCheckBox Then
														objClassAdmin(iCount).set sSetValue
														Fn_ClassAdmin_CheckBoxSet=true
													End If
											End If
										Next

		
								End If
							Set objExport=nothing

		Else
		
				JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("XML Export").SetTOProperty "title","XML Import"
				 Set objImport= JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("XML Export")
				 'objImport.highlight

				 Set objSelectType = description.Create()
									objSelectType("Class Name").value = "JavaCheckBox"
									Set objClassAdmin =JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("XML Export").ChildObjects(objSelectType)
									sStaticCount=JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("XML Export").ChildObjects(objSelectType).count
									For iCount=0 to sStaticCount-1
										sText=objClassAdmin(iCount).GetRoProperty("attached text")
										If objClassAdmin(iCount).GetRoProperty("enabled")="1" Then
										objClassAdmin(iCount).set "OFF"
												If sText=sCheckBox Then
													objClassAdmin(iCount).set sSetValue
													Fn_ClassAdmin_CheckBoxSet=true
												End If
										End If
									Next
					Set objImport=nothing

		End If

End Function


''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''/$$$$
''/$$$$   FUNCTION NAME   :  Fn_ClassAdmin_Import
''/$$$$
''/$$$$   DESCRIPTION        :  To import various enties in Classification Admin
''/$$$$
''/$$$$    PARAMETERS      :   1.) sAction : XML 
''/$$$$                                     2.) sInputFile : Name of input file to be created
''/$$$$ 									3.) bUpdateExisting : True or False
''/$$$$										4.) aChbDetails : Checboc to be made OFF
''/$$$$										5.) sButton : button to be cliked at the end. If blank then click on Import button
''/$$$$										6.) sInfo : For feature use
''/$$$$
''/$$$$	   Return Value : 			True / False
''/$$$$
''/$$$$	 HISTORY           :   AUTHOR                 									DATE        										VERSION
''/$$$$
''/$$$$    CREATED BY     :   Prasanna,Shreyas(AddedCase PLMXML)        							13/04/2011         								1.0
''/$$$$
''/$$$$    REVIWED BY     :  
''/$$$$
''/$$$$    EXAMPLE          : 	bReturn=Fn_ClassAdmin_Import("XML","C:\mainline\TestData\Classification\TestXMLFile_32242.xml",true,"","","")
''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

Public Function Fn_ClassAdmin_Import(sAction,sInputFile,bUpdateExisting,aChbDetails,sButton,sInfo)
	GBL_FAILED_FUNCTION_NAME="Fn_ClassAdmin_Import"

   Dim iCounter,bReturn
   Fn_ClassAdmin_Import=false
	Err.Clear
   'check the existence of the XML export dialog

	   If JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("XML Import").Exist(5) =False Then
		   bReturn = Fn_ToolbatButtonClick("Import Objects")
							 If bReturn = false Then
										Fn_ClassAdmin_Import = False
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to select tab [ " + sTab + "] of XML Import dialog" ) 
										Exit Function 
								End if
				Call Fn_ReadyStatusSync(3)
	  End if


		 If  JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("XML Import").Exist(5) Then

				Select Case sAction

								Case "XML"
											'activate the specified tab
						
											JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("XML Import").JavaTab("Tab").Select "XML"						
												    If Err.Number < 0 Then
														Fn_ClassAdmin_Import = False
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to select tab [XML] of XML Import dialog" ) 
														Exit Function 
													End if
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully selected tab [XML] of XML Import dialog" ) 

                                        	'Set the input 
											JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("XML Import").JavaEdit("Input File").Set sInputFile						
													If Err.Number < 0 Then
																	Fn_ClassAdmin_Import = False
																	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to set file name [ " + sInputFile + "] of XML Import dialog" ) 
																	Exit Function 
													End if
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully set file name[ " + sInputFile + "] of XML Import dialog" ) 
											
											'Update the existing objects check box
											If cbool(bUpdateExisting) Then
													JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("XML Import").JavaCheckBox("ChbOptions").SetTOProperty "attached text","Update Existing Objects"
													JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("XML Import").JavaCheckBox("ChbOptions").Set "ON"
													If Err.Number < 0 Then
															Fn_ClassAdmin_Import = False
															JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("XML Import").JavaButton("Cancel").Click micLeftBtn 
															Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to set the checkbox Update Existing Objects to true" ) 
															Exit Function
													End If
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully set checkbox Update Existing Objects to true" ) 
											End If

											'Select the specific check box which is to be made off
											If IsArray(aChbDetails) Then
													For iCounter = 0 to ubound(aChbDetails)
															JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("XML Import").JavaCheckBox("ChbOptions").SetTOProperty "attached text",aChbDetails(iCounter)
															JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("XML Import").JavaCheckBox("ChbOptions").Set "OFF"
															If Err.Number < 0 Then
																		Fn_ClassAdmin_Import = False
																		JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("XML Import").JavaButton("Cancel").Click micLeftBtn 
																		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to turn off the checkbox ["+aChbDetails(iCounter)+"]" ) 
																		Exit Function 
															Else
																		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully turned off the checkbox ["+aChbDetails(iCounter)+"]" ) 	
															End If
													Next									
											End If

											'Click on Import Button
											JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("XML Import").JavaButton("Import").Click micLeftBtn         
											 If JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("Import Completed").Exist Then
														JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("Import Completed").JavaButton("OK").Click micLeftBtn 
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully handled the Import Completed dialog by clicking  OK button") 
											End If
											If Err.Number < 0 Then
															Fn_ClassAdmin_Import = False
															Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to click on Import Button" ) 															
															Exit Function
											Else
															Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully clicked on Import Button") 
															Fn_ClassAdmin_Import = true		
															Call Fn_ReadyStatusSync(5)
											End If

                                            'Handle the Complete Import dialog
										   If JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("Import Completed").Exist(3)  Then  
												 JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("Import Completed").JavaButton("Yes").SetTOProperty "label", "OK"
												 JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("Import Completed").JavaButton("Yes").Click micLeftBtn
												 If Err.Number < 0 Then
														Fn_ClassAdmin_Import = False
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Handle the Complete Import dialog" )                
														Exit Function
												 Else
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Handled the Complete Import dialog") 
														Fn_ClassAdmin_Import = true  
														Call Fn_ReadyStatusSync(5)
												 End If
										   End If



					Case "PLMXML"
											'activate the specified tab
						
											JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("XML Import").JavaTab("Tab").Select "PLMXML"
												    If Err.Number < 0 Then
														Fn_ClassAdmin_Import = False
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to select tab [PLMXML] of XML Import dialog" ) 
														Exit Function 
													End if
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully selected tab [PLMXML] of XML Import dialog" ) 

                                        	'Set the input 
                                        	JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("XML Import").JavaEdit("InputFile").Set sInputFile						
								'			JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("XML Import").JavaEdit("Input File").Set sInputFile						
													If Err.Number < 0 Then
																	Fn_ClassAdmin_Import = False
																	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to set file name [ " + sInputFile + "] of XML Import dialog" ) 
																	Exit Function 
													End if
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully set file name[ " + sInputFile + "] of XML Import dialog" ) 
											
										

									'Specify the  Transfer Mode name
										If instr(1,sInfo,":") > 0 Then
											aProperties=split(sInfo,":",-1,1)
										else
											aProperties = array(sInfo)
										End If

										If aProperties(0)<>"" Then
											JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("XML Import").JavaList("TransferMode").Select aProperties(0)
														If Err.Number < 0 Then
																		Fn_ClassAdmin_Import = False
																		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to select transfer mode [ " + aProperties(0) + "] of XML Import dialog" ) 
																		Exit Function 
														End if
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully selected transfer mode [ " + aProperties(0) + "] of XML Import dialog") 
										End If


											'Click on Import Button
											JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("XML Import").JavaButton("Import").Click micLeftBtn 		
											If Err.Number < 0 Then
															Fn_ClassAdmin_Import = False
															Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to click on Import Button" ) 															
															Exit Function
											Else
															Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully clicked on Import Button") 
															Fn_ClassAdmin_Import = true		
															Call Fn_ReadyStatusSync(5)
											End If

											'Handle the Complete Import dialog
											If JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("Import Completed").Exist(3)  Then		
													JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaDialog("Import Completed").JavaButton(aProperties(1)).Click micLeftBtn
													If Err.Number < 0 Then
																Fn_ClassAdmin_Import = False
																Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Handle the Complete Import dialog" ) 															
																Exit Function
													Else
																Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Handled the Complete Import dialog") 
																Fn_ClassAdmin_Import = true		
																Call Fn_ReadyStatusSync(5)
													End If
											End If
						
				End Select
		End if

End function




''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''/$$$$
''/$$$$   FUNCTION NAME   :  Fn_XML_File_Operations
''/$$$$
''/$$$$   DESCRIPTION        :  To import various enties in Classification Admin
''/$$$$
''/$$$$    PARAMETERS      :   1.) sAction : XML 
''/$$$$                                     2.) sFilePath : Name of input file to be created
''/$$$$ 									3.) sString : String To Be Replace
''/$$$$										4.) sValue : New string
''/$$$$										5.) sDetails : feature use
''/$$$$
''/$$$$	   Return Value : 			True / False
''/$$$$
''/$$$$	 HISTORY           :   AUTHOR                 									DATE        										VERSION
''/$$$$
''/$$$$    CREATED BY     :   Prasanna,      										15/04/2011         								1.0
''/$$$$
''/$$$$    REVIWED BY     :  
''/$$$$
''/$$$$    EXAMPLE          : 	bReturn=Fn_XML_File_Operations("Modify","C:\mainline\TestData\Classification\TestXMLFile_32242.xml","StrToBeReplace","NewString","")
''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

Public Function Fn_XML_File_Operations(sAction ,sFilePath, sString,sValue,sDetails)
	GBL_FAILED_FUNCTION_NAME="Fn_XML_File_Operations"
	Dim fso,MyFile,sTextLine, sContents,sChange
	Dim aString,objFSO,strText,strNewText,objFile,fObjWrite,bFoundFlag,iIncCounter,sActual
	Select Case sAction
				Case "ModifyLine"
				'Create an object
				 Set fso = CreateObject("Scripting.FileSystemObject")
				 'Check the exsitance of file
				 If  fso.FileExists(sFilePath)Then
						' Log the Result for the existance of file
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: File " + sFilePath + "  exist")
						Set MyFile = fso.GetFile(sFilePath)
					
						'Read All contents from the file
						Set sContents = MyFile.OpenAsTextStream()
						sTextLine =   sContents.ReadAll
						aText = split(sTextLine,vblf,-1,1)
						bFoundFlag = false
		
						Set fObjWrite = fso.OpenTextFile(sFilePath,2,true)
						For iCounter = 0 to Ubound(aText)
							If len(aText(iCounter)) > 1 Then
									If instr(1,trim(aText(iCounter)), sString) Then					
											fObjWrite.WriteLine sValue
											Fn_XML_File_Operations = true
											bFoundFlag = true
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS:The given text is replaced")
									else
											fObjWrite.WriteLine aText(iCounter)										
									End If
							End If
						Next
						fObjWrite.Close
						 If bFoundFlag = false Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL:The mentioned text has not been replaced.")
								Fn_XML_File_Operations=False	
								Exit function
						 End If
				Else
	   
						' Log the Result for the non-existance of file
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL:File does not exist")
						Fn_XML_File_Operations=False
			 End If

			  		 Case "VerifyData"   'Added by Prasanna 18-Apr-2011
				'Create an object
				 Set fso = CreateObject("Scripting.FileSystemObject")
				 'Check the exsitance of file
				 If  fso.FileExists(sFilePath)Then
						' Log the Result for the existance of file
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: File " + sFilePath + "  exist")
						Set MyFile = fso.GetFile(sFilePath)

						'split the string if it contains paramter with :
						If instr(1,sString,":") Then
								sString = split(sString,":",-1,1)
						else
								sString = Array(sString)
						End If

							sActual = sDetails
						'Read All contents from the file
						Set sContents = MyFile.OpenAsTextStream()
						sTextLine =   sContents.ReadAll
						aText = split(sTextLine,vblf,-1,1)
						
						iIncCounter  = 0 
						'iFlagcounter = 0 
						bFoundFlag = 0 
						sDetails = cint(sDetails) + cint(sDetails) ' Adding same no. of lines to handle blank lines
                        For iCounter = 0 to Ubound(aText)
								If len(aText(iCounter)) > 1 Then
											If instr(1,trim(aText(iCounter)), sString(0)) Then                  											
														If UBound(sString) > 0Then
																For iLineCounter = 1 to cint(sDetails) 
																			If  len(aText(iCounter+iLineCounter)) > 1 Then
																						If instr(1,trim(aText(iCounter+iLineCounter)), sString(iLineCounter-iIncCounter)) Then					
																							bFoundFlag = bFoundFlag + 1																				
																						end if			
																			Else																			
																						iIncCounter =  iIncCounter + 1
																			End If
																Next
																 If bFoundFlag = cint(sActual ) Then										
																				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS:The given text is found in the file")
																				Fn_XML_File_Operations=true		
																				Exit for
																 else
																				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL:The mentioned text is not present.")
																				Fn_XML_File_Operations=False
																				Exit function								
																 End If
														 else
																Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS:The given text is found in the file")
																Fn_XML_File_Operations=true		
																Exit for
														 End If                                                                              											
												End If
									End If							
						Next

						
				Else
	   
						' Log the Result for the non-existance of file
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL:File does not exist")
						Fn_XML_File_Operations=False
			 End If
		End Select
End function 


'''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'''''/$$$$
'''''/$$$$   FUNCTION NAME   : Fn_ClassAdmin_ACLOperations(sAction,sTreeNode,sACLName,sInfo1,sInfo2)
'''''/$$$$
'''''/$$$$   DESCRIPTION        :  This function will perform various ACL operations
'''''/$$$$
'''''/$$$$    PARAMETERS      :   1.) sAction : Action to be performed
'''''/$$$$                                     2.) sTreeNode : valid tree node
'''''/$$$$							          3.) sACLName : valid ACL name to be selected
'''''/$$$$									 4.) sInfo2 : For future Use
'''''/$$$$
''''''/$$$$
'''''/$$$$	Return Value : 			True or False
'''''/$$$$
'''''/$$$$    Function Calls       :   Fn_WriteLogFile()
'''''/$$$$										
'''''/$$$$
'''''/$$$$	HISTORY           :   AUTHOR                 DATE        VERSION
'''''/$$$$
'''''/$$$$    CREATED BY     :   SHREYAS           	05/05/2011         1.0
'''''/$$$$
'''''/$$$$    REVIWED BY     :  Prasanna			05/05/2011         1.0
'''''/$$$$
'''''/$$$$	How To Use : bReturn=Fn_ClassAdmin_ACLOperations("Add","Privileges on ICOs","Job","","")
'''''/$$$$						   bReturn=Fn_ClassAdmin_ACLOperations("VerifyTreeNode","no ACL set:Privileges on ICOs","","","")
'''''/$$$$						 
'''''/$$$$
'''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Public function  Fn_ClassAdmin_ACLOperations(sAction,sTreeNode,sACLName,sInfo1,sInfo2)
	GBL_FAILED_FUNCTION_NAME="Fn_ClassAdmin_ACLOperations"
   Dim sNode,objClassAdmin,sValue,sRows,iCount
   Dim aSetValue
   Set objClassAdmin= JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow")
	
'set the Access Control Tab
	bReturn = Fn_ClassAdmin_TabOpeartions("Activate","Subtab","Access Control","")
				If bReturn = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile")," Fail  | Failed to Activate [ Access Control ] Tab")
					Fn_ClassAdmin_ACLOperations = False
					Exit Function 
				Else
					Call Fn_ReadyStatusSync(5)
					Call Fn_WriteLogFile(Environment.Value("TestLogFile")," Pass | Successfully Activated [ Access Control ] Tab")
				End If
	Err.Clear
   Select Case sAction

 	Case "Add"


		if sTreeNode<>"" then

			'Select the node from the ACL tree
			objClassAdmin.JavaTree("ACL").Select "no ACL set:"+sTreeNode
			If err.number<0 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile")," Fail  | Failed to select the node "+sTreeNode)
					Fn_ClassAdmin_ACLOperations = False
					Exit Function 
			Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile")," Pass  | Successfully selected the node "+sTreeNode)
					Fn_ClassAdmin_ACLOperations = True
			End If
End if

'select the value from the ACL list

If sACLName<>"" then
	'objClassAdmin.JavaList("ACLName").Select sACLName
	aSetValue= objClassAdmin.JavaList("ACLName").GetItemIndex(sACLName)
   objClassAdmin.JavaList("ACLName").Object.setSelectedIndex aSetValue,true
	If err.number<0 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile")," Fail  | Failed to select the value "+sACLName)
					Fn_ClassAdmin_ACLOperations = False
					Exit Function 
			Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile")," Pass  | Successfully selected the value "+sACLName)
					Fn_ClassAdmin_ACLOperations = True
			End If
End if

'Click on the Add button
objClassAdmin.JavaButton("AddACL").Click micLeftBtn
			If err.number<0 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile")," Fail  | Failed to add the ACL "+sACLName)
					Fn_ClassAdmin_ACLOperations = False
					Exit Function 
			Else
					Call Fn_ReadyStatusSync(5)
					Call Fn_WriteLogFile(Environment.Value("TestLogFile")," Pass  | Successfully added the ACL "+sACLName)
					Fn_ClassAdmin_ACLOperations = True
			End If
	
Case "VerifyTreeNode"
	sRows=objClassAdmin.JavaTree("ACL").GetROProperty("items count")
	For iCount=0 to sRows-1
		sNode=objClassAdmin.JavaTree("ACL").GetItem(iCount)
		If Trim(lcase(sNode)) = Trim(Lcase(sTreeNode)) Then
							Fn_ClassAdmin_ACLOperations = True
							
							Exit For
						End If
					Next
					If cint(iCount) = cint(sRows) Then
						Fn_ClassAdmin_ACLOperations = FALSE
					End If
 End Select
End Function 

'''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'''''/$$$$
'''''/$$$$   FUNCTION NAME   : Fn_ClassAdmin_SetOptimizeDisplayForAtrribute(objChbox,sValue)
'''''/$$$$
'''''/$$$$   DESCRIPTION        :  This function will set check boxes during class creation/modification. By default will set OptimizeDisplay checkbox.
'''''/$$$$
'''''/$$$$    PARAMETERS      :   1.) objChbox : Checkbox to set
'''''/$$$$                                     2.) sValue : value to set "on" or "off"
''''''/$$$$
'''''/$$$$	Return Value : 			True or False
'''''/$$$$
'''''/$$$$   Function Calls       :   Fn_WriteLogFile()
'''''/$$$$										
'''''/$$$$
'''''/$$$$	HISTORY           :   AUTHOR                 DATE        VERSION
'''''/$$$$
'''''/$$$$    CREATED BY     :   Prasad           	03/01/2017         1.0
'''''/$$$$
'''''/$$$$    REVIWED BY     :  Prasad           	03/01/2017         1.0
'''''/$$$$
'''''/$$$$	How To Use : bReturn=Fn_ClassAdmin_SetOptimizeDisplayForAtrribute("","ON")
'''''/$$$$						   
'''''/$$$$						 
'''''/$$$$
'''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

Public Function Fn_ClassAdmin_SetOptimizeDisplayForAtrribute(objChbox, sValue)
	GBL_FAILED_FUNCTION_NAME="Fn_ClassAdmin_SetOptimizeDisplayForAtrribute"
		'Set the Value for Optimize display if set                
		JavaWindow("ClassAdminMainWin").JavaWindow("ClassSubWindow").JavaCheckBox("OptimizeDisplay").Set sValue 
		If Err.Number < 0 Then
				Fn_ClassAdmin_SetOptimizeDisplayForAtrribute = False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Set [ Optimize Display ] = [ " + sValue + "] " ) 
				Exit Function 
		End If
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Set [ Optimize Display ] = [ " + sValue + "] " ) 
		Fn_ClassAdmin_SetOptimizeDisplayForAtrribute = true
End Function
