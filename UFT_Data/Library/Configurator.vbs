Option Explicit

' Function List
'================================================================================================================================================================
'000.  Fn_PC_GetObject()
'001.  Fn_PC_NavTree_NodeOperation()
'002.  Fn_PC_ConfigurationContext_Operations()
'003.  Fn_PC_Dictionary_Operations	
'004.  Fn_PC_BreadcrumbOperations
'005.  Fn_SISW_PC_NavTreeTableOperations
'006.  Fn_PC_ErrorVerify
'007.  Fn_PC_CompnentTabOperations
'008.  Fn_PC_RevisionRuleOperations
'009.  Fn_PC_VariantNatTable_VariantExpressionEditor_Operations
'010.  Fn_PC_VariantConfigurationView_Operation
'011.  Fn_PC_ConfiguratorRules_Operation
'012.  Fn_PC_EffectivityOperations
'013.  Fn_PC_DateControl
'014.  Fn_PC_FreeFromRule_Operations
'015.  Fn_PC_Open_Context_Independent_Search_View_Operations'================================================================================================================================================================

'================================================================================================================================================================
'@@ Function Name	:	Fn_PC_GetObject
'@@ 
'@@ Description		:  	Function to get specified Object hierarchy.
'@@ 
'@@ Parameters		:	1. sObjectName : Object Handle name
'@@ 						
'@@ Return Value	:  	Object \ Nothing
'@@ 
'@@ Examples		:	Fn_PC_GetObject("ProductConfigurator")
'@@ 
'================================================================================================================================================================
'@@ History			:	Developer Name			Date			Rev. No.		Changes Done			Reviewer
'================================================================================================================================================================
'@@ Created By		:	Snehal Salunkhe			19-Oct-2015		1.0				
'================================================================================================================================================================
Function Fn_PC_GetObject(sObjectName)
	Dim sObjectXMLPath
	sObjectXMLPath = Fn_GetEnvValue("User", "AutomationDir") & "\TestData\AutomationXML\ObjectXML\Configurator.xml"
	Set Fn_PC_GetObject = Fn_SISW_Setup_GetObjectFromXML(sObjectXMLPath, sObjectName)
End Function

'================================================================================================================================================================
'@@ Function Name	:	Fn_PC_NavTree_NodeOperation

'@@ Description		:	Operations on Nodes in Nav Tree in Product Configurator perspective

'@@ Parameters		:	1. StrAction: Action to be performed
'@@ 					2. StrNodeName: Fully qulified tree Path (delimiter as ':') [multiple node are separated by "," ] 
'@@ 					3. StrMenu: Context menu to be selected

'@@ Return Value	: 	TRUE\FALSE

'@@ Pre-requisite	:	Product Configurator module window should be displayed

'@@ Examples		:	for Case "PopupMenuSelect" 	: Fn_PC_NavTree_NodeOperation("PopupMenuSelect","Home:Newstuff","Copy Ctrl+C")
'@@ 					for Case "Select" 			: Fn_PC_NavTree_NodeOperation("Select","Home:Newstuff:000032-CarModel_VI_LS1:000032","")
 					
'================================================================================================================================================================
'@@ History			:	Developer Name			Date			Rev. No.		Changes Done			Reviewer
'================================================================================================================================================================
'@@ Created By		:	Snehal Salunkhe			19-Oct-2015		1.0				Created					Vivek A	[TC1121-20152015101200-27_10_2015-VivekA-NewDevelopment]
'================================================================================================================================================================
Public Function Fn_PC_NavTree_NodeOperation(StrAction,StrNodeName,StrMenu)

	GBL_FAILED_FUNCTION_NAME="Fn_PC_NavTree_NodeOperation"
	Dim objConfigWin, objConfigNavTree
	Dim intCount, aMenuList, iCnt, iPath,arrNodes
	Dim aNodePath, oCurrentNode,objSelectType,intNoOfObjects

	Set objConfigWin = Fn_PC_GetObject("ProductConfigurator") 
	Set objConfigNavTree = Fn_PC_GetObject("ConfiguratorNavTreeTable") 

	Select Case StrAction

		'----------- For selecting single node
		Case "Select"
				iPath = Fn_UI_JavaTreeGetItemPathExt("Fn_PC_NavTree_NodeOperation", objConfigNavTree, StrNodeName , ":", "@")
				If iPath = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select Node [" + StrNodeName + "] of NavTree")
					Fn_PC_NavTree_NodeOperation = False
				Else
					objConfigNavTree.Select iPath
					Call Fn_ReadyStatusSync(1)
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Selected Node [" + StrNodeName + "] of NavTree")
					Fn_PC_NavTree_NodeOperation = True
				End If
		'----------- Deselect Node
		Case "Deselect"	
				iPath = Fn_UI_JavaTreeGetItemPathExt("Fn_PC_NavTree_NodeOperation", objConfigNavTree, StrNodeName , ":", "@")
				If iPath = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to DeSelect Node [" + StrNodeName + "] of NavTree")
					Fn_PC_NavTree_NodeOperation = False
				Else
					objConfigNavTree.Deselect iPath
					Call Fn_ReadyStatusSync(1)
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully DeSelected Node [" + StrNodeName + "] of NavTree")
					Fn_PC_NavTree_NodeOperation = True
				End If
		'----------- Expand Node
		Case "Expand"
				iPath = Fn_UI_JavaTreeGetItemPathExt("Fn_PC_NavTree_NodeOperation", objConfigNavTree, StrNodeName , ":", "@")
				If iPath=False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Expand Node [" + StrNodeName + "] of NavTree")
					Fn_PC_NavTree_NodeOperation = False
				Else
					objConfigNavTree.Expand iPath
					Call Fn_ReadyStatusSync(1)
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Expanded Node [" + StrNodeName + "] of NavTree")
					Fn_PC_NavTree_NodeOperation = True
				End If
		'----------- Collaplse Node
		Case "Collapse"
				iPath = Fn_UI_JavaTreeGetItemPathExt("Fn_PC_NavTree_NodeOperation", objConfigNavTree, StrNodeName , ":", "@")
				If iPath = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Collapse Node [" + StrNodeName + "] of NavTree")
					Fn_PC_NavTree_NodeOperation = False
				Else
					objConfigNavTree.Collapse iPath
					Call Fn_ReadyStatusSync(1)
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Collapse Node [" + StrNodeName + "] of NavTree")
					Fn_PC_NavTree_NodeOperation = True
				End If
		'----------- Pop Up Menu Select
		Case "PopupMenuSelect"
				'Build the Popup menu to be selected
				aMenuList = split(StrMenu, ":",-1,1)
				intCount = Ubound(aMenuList)

				'Select node
                iPath = Fn_UI_JavaTreeGetItemPathExt("Fn_PC_NavTree_NodeOperation", objConfigNavTree, StrNodeName , ":", "@")
				If iPath=False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Select Node [" + StrNodeName + "] of NavTree")
					Fn_PC_NavTree_NodeOperation = False
					Exit Function
				Else
					objConfigNavTree.Select iPath
					Call Fn_ReadyStatusSync(1)
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Selected Node [" + StrNodeName + "] of NavTree")
					Fn_PC_NavTree_NodeOperation = True
				End If

				'Open context menu
				Call Fn_UI_JavaTree_OpenContextMenu("Fn_PC_NavTree_NodeOperation",objConfigWin,"NavTreeTable",iPath)
                
				'Select Menu action
				Select Case intCount
					Case "0"
						StrMenu = objConfigWin.WinMenu("ContextMenu").BuildMenuPath(aMenuList(0))
					Case "1"
						StrMenu = objConfigWin.WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1))
					Case "2"
						StrMenu = objConfigWin.WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1),aMenuList(2))
					Case Else
						Fn_PC_NavTree_NodeOperation = FALSE
						Exit Function
				End Select

				Err.Clear
				objConfigWin.WinMenu("ContextMenu").Select StrMenu
				If Err.number < 0 Then
					Fn_PC_NavTree_NodeOperation = False
				Else
					Fn_PC_NavTree_NodeOperation = True
				End If
		'----------- Existance of Node
		Case "Exist"
				iPath = Fn_UI_JavaTreeGetItemPathExt("Fn_PC_NavTree_NodeOperation", objConfigNavTree, StrNodeName , ":", "@")
				If iPath = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Node [" + StrNodeName + "] Not exist in NavTree")
					Fn_PC_NavTree_NodeOperation = False
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Node [" + StrNodeName + "] Exist in NavTree")
					aNodePath = split(replace(iPath,"#",""),":")
					Fn_PC_NavTree_NodeOperation = True
					Set oCurrentNode = objConfigNavTree.Object
					For iCnt = 0 to UBound(aNodePath) -1
						Set oCurrentNode = oCurrentNode.GetItem(aNodePath(iCnt))
						If cBool(oCurrentNode.getExpanded()) = False Then
							Fn_PC_NavTree_NodeOperation = false
							Exit for
						End If
					Next
					If Fn_PC_NavTree_NodeOperation Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Node [" + StrNodeName + "] Exist in NavTree")
					End If
				End If
		'----------- Existance of Popup Menu
		Case "PopupMenuExist"
				aMenuList = split(StrMenu, ":",-1,1)
				intCount = Ubound(aMenuList)

				'Open context menu
				iPath = Fn_UI_JavaTreeGetItemPathExt("Fn_PC_NavTree_NodeOperation", objConfigNavTree, StrNodeName , ":", "@")
				Call Fn_UI_JavaTree_OpenContextMenu("Fn_PC_NavTree_NodeOperation",objConfigWin,"NavTree",iPath)
				Select Case intCount
					Case "0"
						StrMenu = objConfigWin.WinMenu("ContextMenu").BuildMenuPath(aMenuList(0))
                    Case "1"
						StrMenu = objConfigWin.WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1))
                    Case "2"
						StrMenu = objConfigWin.WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1),aMenuList(2))
                    Case Else
						Fn_PC_NavTree_NodeOperation = False
                    	Exit Function
				End Select
				If objConfigWin.WinMenu("ContextMenu").GetItemProperty (StrMenu,"Exists") = True Then
					Fn_PC_NavTree_NodeOperation = True
				Else
					Fn_PC_NavTree_NodeOperation = False
				End If
				JavaWindow("DefaultWindow").Click 150, 3, "LEFT"
		'----------- Checking State of Popup Menu		
		Case "PopupMenuEnabled"
				aMenuList = split(StrMenu, ":",-1,1)
				intCount = Ubound(aMenuList)

				'Open context menu
				iPath = Fn_UI_JavaTreeGetItemPathExt("Fn_PC_NavTree_NodeOperation", objConfigNavTree, StrNodeName , ":", "@")
				Call Fn_UI_JavaTree_OpenContextMenu("Fn_PC_NavTree_NodeOperation",objConfigWin,"NavTree",iPath)
				Select Case intCount
					Case "0"
						StrMenu = objConfigWin.WinMenu("ContextMenu").BuildMenuPath(aMenuList(0))
                    Case "1"
						StrMenu = objConfigWin.WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1))
                    Case "2"
						StrMenu = objConfigWin.WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1),aMenuList(2))
                    Case Else
						Fn_PC_NavTree_NodeOperation = FALSE
                    	Exit Function
				End Select
				If objConfigWin.WinMenu("ContextMenu").GetItemProperty (StrMenu,"Enabled") = True Then
					Fn_PC_NavTree_NodeOperation = True
				Else
					Fn_PC_NavTree_NodeOperation = False
			  	End If
		'-----------------
		Case "MultiSelect"  '[Tc12.1_20181213.00_NewDevelopment_PoonamC : Added New Case to select Multiple Nodes ]
				arrNodes = Split(StrNodeName,"~")
				For iCnt = 0 To UBound(arrNodes)
					iPath = Fn_UI_JavaTreeGetItemPathExt("Fn_PC_NavTree_NodeOperation", objConfigNavTree, arrNodes(iCnt) , ":", "@")
					If iPath = False Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select Node [" + arrNodes(iCnt) + "] of NavTree")
						Fn_PC_NavTree_NodeOperation = False
						Exit For
					Else
						If iCnt = 0 Then
							objConfigNavTree.Select iPath
						Else
							objConfigNavTree.ExtendSelect iPath			
						End If
						Call Fn_ReadyStatusSync(1)
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Selected Nodes [" + StrNodeName + "] of NavTree")
						Fn_PC_NavTree_NodeOperation = True
					End If	
				Next
		'---------------------------
		Case "ChangeConext","SetRevisionRule"
				Set objSelectType = Description.Create()
					objSelectType("Class Name").value = "JavaObject"
					objSelectType("toolkit class").value = "org.eclipse.ui.forms.widgets.ImageHyperlink"
				Set intNoOfObjects = objConfigWin.ChildObjects(objSelectType)
				For iCnt = 0 to intNoOfObjects.count-1
					If StrAction = "ChangeConext" Then
						If  lcase(trim( "" & intNoOfObjects(iCnt).Object.getToolTipText())) = lCase("Click to view and change the scope") Then
							intNoOfObjects(iCnt).Click 1,1, "LEFT"
							Exit for
						End If
					ElseIf StrAction = "SetRevisionRule" Then
						If  lcase(trim( "" & intNoOfObjects(iCnt).Object.getToolTipText())) = lCase("Click to view and change the current variant configuration") Then
							intNoOfObjects(iCnt).Click 1,1, "LEFT"
							Exit for
						End If
					End If
				Next
				Fn_PC_NavTree_NodeOperation = Fn_UI_JavaMenu_Select("Fn_PC_NavTree_NodeOperation",objConfigWin,StrNodeName)
				Set objSelectType = Nothing
				Set intNoOfObjects = Nothing
		Case Else
			Fn_PC_NavTree_NodeOperation = False
	End Select

	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), StrAction &" Sucessfully completed on Node [" + StrNodeName + "] of JavaTree of function Fn_PC_NavTree_NodeOperation")
	Set objConfigWin = nothing
	Set objConfigNavTree = nothing
End Function

'================================================================================================================================================================
'@@	Function Name		:	Fn_PC_ConfigurationContext_Operations
'@@
'@@ Description			:	Function Used to Create Business Object
'@@
'@@ Parameters			:   1.StrObjectName: Business Object  Name
'@@ 						2.dicBOInfo: Business Object information 
'@@	
'@@ Return Value		: 	(ItemId-ItemRevID) Or False
'@@
'@@ Pre-requisite		:	Select NewBusinessObject Should be present
'@@
'@@ Examples			:   Set dicBOInfo = CreateObject("Scripting.Dictionary")
'@@								dicBOInfo("ID")="123456"
'@@								dicBOInfo("Revision")="A"
'@@								dicBOInfo("VerifyPositiveAvailabilityDefaultVal")="False"
'@@								dicBOInfo("VerifyPositiveAvailabilityValues")="True~False"
'@@								dicBOInfo("ButtonState@1")="Finish:False"
'@@								dicBOInfo("Name")="Context113"
'@@								dicBOInfo("ButtonState@2")="Finish:True"
'@@								dicBOInfo("Description")="Context Created"
'@@								dicBOInfo("Open On Create")="ON"
'@@							bReturn = Fn_PC_ConfigurationContext_Operations("Create",dicBOInfo,"Close")
'================================================================================================================================================================
'History			:	Developer Name				Date				Rev. No.		Changes Done			Reviewer
'================================================================================================================================================================
'Created By 		:	Snehal Salunkhe				20-Oct-2015			1.0				Created					Vivek A	[TC1121-20152015101200-27_10_2015-VivekA-NewDevelopment]
'================================================================================================================================================================
'Modified By 		:	Poonam Chopade				12-Jan-2017			1.1				Modified function added Case "Create"
'================================================================================================================================================================
Public Function Fn_PC_ConfigurationContext_Operations(strAction,dicBOInfo,strButton)
	GBL_FAILED_FUNCTION_NAME="Fn_PC_ConfigurationContext_Operations"
	'Variable Declaration
	Dim ObjCCCreate,arrKeys,iCounter,arrItem, sItemId, sRevId, sMenu,arrFields,iCnt,sItemName
	
	'Function Returns False
	Fn_PC_ConfigurationContext_Operations = False
	
	'Creating Object of [ NewBusinessObject ] dialog
	Set ObjCCCreate = Fn_PC_GetObject("ConfiguratorContext")
	If Fn_UI_ObjectExist("Fn_PC_ConfigurationContext_Operations",ObjCCCreate) = False Then
			 'Calling Menu [ File:New:Configurator Context... ] to invoke [ NewBusinessObject ] dialog
	     	  sMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("Configurator_Menu"), "ConfiguratorContext")
		 	  Call Fn_MenuOperation("Select",sMenu)
		      Call Fn_ReadyStatusSync(1)
			  
			 ' Checking existence of Object
			If Fn_UI_ObjectExist("Fn_PC_ConfigurationContext_Operations",ObjCCCreate) = False Then
				Exit Function
			End if				  
	End if

  Select Case strAction
		Case "Create"
		    'Taking All keys from Dictionary
  			 arrKeys=dicBOInfo.Keys
		
			For iCounter=0 To dicBOInfo.Count-1
				Select Case arrKeys(iCounter)
					'------------ Case for ID, Revision
					Case "ID","Revision"
						If dicBOInfo(arrKeys(iCounter))<>"" Then
							arrItem=Split(dicBOInfo(arrKeys(iCounter)),"~")
							If Not ObjCCCreate.JavaStaticText("ObjectName").Exist(3) Then
								ObjCCCreate.JavaStaticText("ObjectName").SetTOProperty "label",arrKeys(iCounter)+":"
							End If
							If Fn_UI_Object_GetROProperty("Fn_PC_ConfigurationContext_Operations",ObjCCCreate.JavaEdit(arrKeys(iCounter)),"enabled") = 1 Then
							    Call Fn_Edit_Box("Fn_PC_ConfigurationContext_Operations",ObjCCCreate,arrKeys(iCounter),arrItem(0))
							End If 
						Else
							If Fn_UI_Object_GetROProperty("Fn_PC_ConfigurationContext_Operations",ObjCCCreate.JavaEdit(arrKeys(iCounter)),"enabled") = 1 Then
								Call Fn_Button_Click("Fn_PC_ConfigurationContext_Operations",ObjCCCreate, "Assign"&arrKeys(iCounter))
							End If	
						End If
						Call Fn_ReadyStatusSync(1)
						'Get values of Item ID and Item revision ID
						sItemId = Fn_Edit_Box_GetValue("Fn_PC_ConfigurationContext_Operations", ObjCCCreate,"ID")
					    sRevId = Fn_Edit_Box_GetValue("Fn_PC_ConfigurationContext_Operations", ObjCCCreate,"Revision")
					    
					'-------- Verify PositiveAvailability Field's Default value
					Case "VerifyPositiveAvailabilityDefaultVal"
						  ObjCCCreate.JavaRadioButton("PositiveAvailability").SetTOProperty "attached text",dicBOInfo(arrKeys(iCounter))
						  If Fn_UI_Object_GetROProperty("Fn_PC_ConfigurationContext_Operations",ObjCCCreate.JavaRadioButton("PositiveAvailability"),"value") <> 0 Then
						 	  Fn_PC_ConfigurationContext_Operations = False
						 	  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_PC_ConfigurationContext_Operations function failed to verify default value for  [ Positive Availability Biased ] as ["+dicBOInfo(arrKeys(iCounter)+"]."))
						      Exit For
						  End If	
					'-------- Verify PositiveAvailability Fields value
					Case "VerifyPositiveAvailabilityValues"	
							arrFields = Split(dicBOInfo(arrKeys(iCounter)),"~")
							For iCnt = 0 To UBound(arrFields)
								 ObjCCCreate.JavaRadioButton("PositiveAvailability").SetTOProperty "attached text",arrFields(iCnt)
								 If Fn_UI_ObjectExist("Fn_PC_ConfigurationContext_Operations",ObjCCCreate.JavaRadioButton("PositiveAvailability")) = False Then
								 	    Fn_PC_ConfigurationContext_Operations = False
									 	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_PC_ConfigurationContext_Operations function failed to verify option ["+arrFields(iCnt)+"] for [ Positive Availability Biased ].")
									    Set ObjCCCreate = Nothing	
									    Exit Function
								 End If
							Next
					'----- Check button button enable or not 
					Case "ButtonState@1","ButtonState@2"	
						  arrFields = Split(dicBOInfo(arrKeys(iCounter)),":")
						  ObjCCCreate.JavaButton("Button").SetTOProperty "label",arrFields(0)
						 If cbool(Fn_UI_Object_GetROProperty("Fn_PC_ConfigurationContext_Operations",ObjCCCreate.JavaButton("Button"),"enabled")) <> cbool(arrFields(1)) Then
						 	Fn_PC_ConfigurationContext_Operations = False
						 	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_PC_ConfigurationContext_Operations function failed to verify button [ "+arrFields(0)+" ] state as ["+arrFields(1)+"].")
						    Exit For
						 End If
					'----- Set Value for Positive Availability Biased
					Case "SetPositiveAvailabilityBiased"
					       If dicBOInfo(arrKeys(iCounter)) <> "" Then
					       		ObjCCCreate.JavaRadioButton("PositiveAvailability").SetTOProperty "attached text",dicBOInfo(arrKeys(iCounter))
					       		Call Fn_SISW_UI_JavaRadioButton_Operations("Fn_PC_ConfigurationContext_Operations","Set",ObjCCCreate,"PositiveAvailability","ON")
					       		Call Fn_ReadyStatusSync(1)
					       End If								
					'------------ Case for Name, Description			
					Case "Name","Description"
						If dicBOInfo(arrKeys(iCounter))<>"" Then
							arrItem=Split(dicBOInfo(arrKeys(iCounter)),"~")
		                    ObjCCCreate.JavaStaticText("ObjectName").SetTOProperty "label",arrKeys(iCounter)+":"
							If Not ObjCCCreate.JavaStaticText("ObjectName").Exist(3) Then
								ObjCCCreate.JavaStaticText("ObjectName").SetTOProperty "label",arrKeys(iCounter)+":"
							End If
							Call Fn_Edit_Box("Fn_PC_ConfigurationContext_Operations",ObjCCCreate,"ObjectName",arrItem(0))
		                    If UBound(arrItem)=1 Then
								Call Fn_Button_Click("Fn_PC_ConfigurationContext_Operations",ObjCCCreate, "Next")
							End If
						End If
						Call Fn_ReadyStatusSync(1)
						' Added code by Jotiba T as per design change- Discussed with Abhijit Patil 
						If arrKeys(iCounter)="Name" Then
							If ObjCCCreate.JavaRadioButton("PositiveAvailability").Exist(1) Then
								ObjCCCreate.JavaRadioButton("PositiveAvailability").SetTOProperty "attached text", "True"
								wait 1
								ObjCCCreate.JavaRadioButton("PositiveAvailability").Set "ON"
							End If
							sItemName = dicBOInfo(arrKeys(iCounter))
						End If
						
					'------------ Case for Open On Create	
					Case "Open On Create"
						If dicBOInfo(arrKeys(iCounter))<>"" Then
							Call Fn_SISW_UI_JavaCheckBox_Operations("","Set",ObjCCCreate,"OpenOnCreate",dicBOInfo(arrKeys(iCounter)))
						Else
							Call Fn_SISW_UI_JavaCheckBox_Operations("","Set",ObjCCCreate,"OpenOnCreate","OFF")
						End If
						Call Fn_ReadyStatusSync(1)
					'------------ Case Else
					Case Else
						Call Fn_Button_Click("Fn_PC_ConfigurationContext_Operations",ObjCCCreate, "Close")
						Fn_PC_ConfigurationContext_Operations=False
						Set ObjCCCreate=Nothing
						Exit Function
				End Select
			Next
			'Clicking On Finish button
			Call Fn_Button_Click("Fn_PC_ConfigurationContext_Operations",ObjCCCreate, "Finish")
			Call Fn_ReadyStatusSync(1)	
			Fn_PC_ConfigurationContext_Operations = "'"&sItemId & "-" & sRevId
 End Select
 	
 	If strButton <> "" Then	
		arrFields = Split(strButton,":")
		For iCounter = 0 To UBound(arrFields)
		    'Clicking On Close button
			Call Fn_Button_Click("Fn_PC_ConfigurationContext_Operations",ObjCCCreate, arrFields(iCounter))
			Call Fn_ReadyStatusSync(1)
		Next				
   End if 
 	For iCounter=0 To dicBOInfo.Count-1
	   If arrKeys(iCounter)="Open On Create" Then
	       
	       Call Fn_PC_CompnentTabOperations("Maximize",sItemId&"-"&sItemName&" (Variability Explorer)","") 
	   End If 
 	Next
  Set ObjCCCreate=Nothing
End Function
'================================================================================================================================================================
'@@	Function Name		:	Fn_PC_Dictionary_Operations
'@@
'@@ Description			:	Function Used to Create Business Object
'@@
'@@ Parameters			:   1.StrObjectName: Business Object  Name
'@@ 						2.dicBOInfo: Business Object information 
'@@
'@@ Return Value		: 	(ItemId-ItemRevID) Or False
'@@
'@@ Pre-requisite		:	Select NewBusinessObject Should be present
'@@
'@@ Examples			:   Set dicBOInfo = CreateObject("Scripting.Dictionary")
'@@								dicBOInfo("ID")="123456"
'@@								dicBOInfo("Revision")="A"
'@@								dicBOInfo("VerifyFieldExists")="Positive Availability Biased:True/False" 
'@@								dicBOInfo("ButtonState@1")="Finish:False" 
'@@								dicBOInfo("Name")="Dic113"
'@@								dicBOInfo("ButtonState@2")="Finish:True" 
'@@								dicBOInfo("Description")="Dictionary Created"
'@@								dicBOInfo("Open On Create")="ON"
'@@							bReturn = Fn_PC_Dictionary_Operations("Create",dicBOInfo,"Finish:Close")
'@@					   
'================================================================================================================================================================
'History			:	Developer Name				Date				Rev. No.		Changes Done			Reviewer
'================================================================================================================================================================
'Created By 		:	Poonam Chopade				05-Jan-2017			1.0				Created					
'================================================================================================================================================================
Public Function Fn_PC_Dictionary_Operations(strAction,dicBOInfo,strButton)
	GBL_FAILED_FUNCTION_NAME="Fn_PC_Dictionary_Operations"
	'Variable Declaration
	Dim ObjDicCreate,arrKeys,iCounter,arrItem, sItemId, sRevId,sMenu
	Dim arrFields
	'Function Returns False
	Fn_PC_Dictionary_Operations = False
	
	'Creating Object of [ NewBusinessObject ] dialog
	Set ObjDicCreate = Fn_PC_GetObject("ConfiguratorContext")
	If Fn_UI_ObjectExist("Fn_PC_Dictionary_Operations",ObjDicCreate) = False Then
			 'Calling Menu [ File:New:Dictionary... ] to invoke [ NewBusinessObject ] dialog
		 	 sMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("Configurator_Menu"), "FileNewDictionary")
			 Call Fn_MenuOperation("Select",sMenu)
			 Call Fn_ReadyStatusSync(1)	
			 
			' Checking existence of Object
			If Fn_UI_ObjectExist("Fn_PC_Dictionary_Operations",ObjDicCreate) = False Then
				Exit Function
			End if
	End if			 

  Select Case strAction
  	Case "Create"
			'Taking All keys from Dictionary
			arrKeys = dicBOInfo.Keys
		
			For iCounter = 0 To dicBOInfo.Count - 1
				Select Case arrKeys(iCounter)
					'------------ Case for ID, Revision
					Case "ID","Revision"
						If dicBOInfo(arrKeys(iCounter))<>"" Then
							arrItem=Split(dicBOInfo(arrKeys(iCounter)),"~")
							If Not ObjDicCreate.JavaStaticText("ObjectName").Exist(3) Then
								ObjDicCreate.JavaStaticText("ObjectName").SetTOProperty "label",arrKeys(iCounter)+":"
							End If
							If Fn_UI_Object_GetROProperty("Fn_PC_Dictionary_Operations",ObjDicCreate.JavaEdit(arrKeys(iCounter)),"enabled") = 1 Then
								Call Fn_Edit_Box("Fn_PC_Dictionary_Operations",ObjDicCreate,arrKeys(iCounter),arrItem(0))
							End If
						Else
						   If Fn_UI_Object_GetROProperty("Fn_PC_Dictionary_Operations",ObjDicCreate.JavaEdit(arrKeys(iCounter)),"enabled") = 1 Then	
							   Call Fn_Button_Click("Fn_PC_Dictionary_Operations",ObjDicCreate, "Assign"&arrKeys(iCounter))
						   End If		   
						End If
						Call Fn_ReadyStatusSync(1)
						'Get values of Item ID and Item revision ID
					    sItemId = Fn_Edit_Box_GetValue("Fn_PC_Dictionary_Operations", ObjDicCreate,"ID")
					    sRevId = Fn_Edit_Box_GetValue("Fn_PC_Dictionary_Operations", ObjDicCreate,"Revision")
				   '-------- Verify Filed is exists or not
					Case "VerifyPositiveAvailabilityFieldExists"		
						 If cbool(Fn_UI_ObjectExist("Fn_PC_Dictionary_Operations",ObjDicCreate.JavaStaticText("PositiveAvailability"))) <> cbool(dicBOInfo(arrKeys(iCounter))) Then
						 	Fn_PC_Dictionary_Operations = False
						 	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_PC_Dictionary_Operations function failed to verify existence of [ Positive Availability Biased ] as ["+dicBOInfo(arrKeys(iCounter))+"].")
							Exit For
						 End If
					'----- Check button button enable or not 
					Case "ButtonState@1","ButtonState@2"	
						  arrFields = Split(dicBOInfo(arrKeys(iCounter)),":")
						 ObjDicCreate.JavaButton("Button").SetTOProperty "label",arrFields(0)
						 If cbool(Fn_UI_Object_GetROProperty("Fn_PC_Dictionary_Operations",ObjDicCreate.JavaButton("Button"),"enabled")) <> cbool(arrFields(1)) Then
						 	Fn_PC_Dictionary_Operations = False
						 	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_PC_Dictionary_Operations function failed to verify button [ "+arrFields(0)+" ] state as ["+arrFields(1)+"].")
						    Exit For
						 End If		
					'------------ Case for Name, Description			
					Case "Name","Description"
						If dicBOInfo(arrKeys(iCounter))<>"" Then
							arrItem=Split(dicBOInfo(arrKeys(iCounter)),"~")
		                    ObjDicCreate.JavaStaticText("ObjectName").SetTOProperty "label",arrKeys(iCounter)+":"
							If Not ObjDicCreate.JavaStaticText("ObjectName").Exist(3) Then
								ObjDicCreate.JavaStaticText("ObjectName").SetTOProperty "label",arrKeys(iCounter)+":"
							End If
							Call Fn_Edit_Box("Fn_PC_Dictionary_Operations",ObjDicCreate,"ObjectName",arrItem(0))
		                    If UBound(arrItem)=1 Then
								Call Fn_Button_Click("Fn_PC_Dictionary_Operations",ObjDicCreate, "Next")
							End If
						End If
						Call Fn_ReadyStatusSync(1)					
					'------------ Case for Open On Create	
					Case "Open On Create"
						If dicBOInfo(arrKeys(iCounter))<>"" Then
							Call Fn_SISW_UI_JavaCheckBox_Operations("Fn_PC_Dictionary_Operations","Set",ObjDicCreate,"OpenOnCreate",dicBOInfo(arrKeys(iCounter)))
						Else
							Call Fn_SISW_UI_JavaCheckBox_Operations("Fn_PC_Dictionary_Operations","Set",ObjDicCreate,"OpenOnCreate","OFF")
						End If
						Call Fn_ReadyStatusSync(1)
					'------------ Case Else
					Case Else
						Call Fn_Button_Click("Fn_PC_Dictionary_Operations",ObjDicCreate, "Cancel")
						Fn_PC_Dictionary_Operations=False
						Set ObjDicCreate=Nothing
						Exit Function
				End Select
			Next
			Fn_PC_Dictionary_Operations = sItemId & "-" & sRevId
	
   End Select		   	
		
  If strButton <> "" Then	
		arrFields = Split(strButton,":")
		For iCounter = 0 To UBound(arrFields)
		    'Clicking On Finish/Close button
			Call Fn_Button_Click("Fn_PC_Dictionary_Operations",ObjDicCreate, arrFields(iCounter))
			Call Fn_ReadyStatusSync(1)
		Next				
  End if 
  Set ObjDicCreate = Nothing	    
	
End Function
'================================================================================================================================================================
'@@	Function Name		:	Fn_PC_BreadcrumbOperations
'@@
'@@ Description			:	Function Used to Set / Verify Revision Rule or Effectivity or Compile Rule Set
'@@
'@@ Parameters			:   1.sExistingRevRule : RuleName / EffName / CompileRuleSet
'@@ 						2.sNewRevRule: For future use
'@@
'@@ Return Value		: 	True Or False
'@@
'@@ Examples			:   Call Fn_PC_BreadcrumbOperations("Exist","AnyStatus; Working(Modified)~No Effectivity~14-Dec-2016 00:00","")
'@@					   
'================================================================================================================================================================
'History			:	Developer Name				Date				Rev. No.		Changes Done			Reviewer
'================================================================================================================================================================
'Created By 		:	Poonam Chopade				05-Jan-2017			1.0				Created					
'================================================================================================================================================================
Public Function Fn_PC_BreadcrumbOperations(strAction,sExistingRevRule,sNewRevRule)
	GBL_FAILED_FUNCTION_NAME="Fn_PC_BreadcrumbOperations"
	'Variable Declaration
	Dim ObjProdImgLink,iCounter
	
	'Function Returns False
	Fn_PC_BreadcrumbOperations=False
	
	'Creating Object of [ NewBusinessObject ] dialog	
	Set ObjProdImgLink = Fn_PC_GetObject("ConfiguratorImageHyperlink")
	
	Select Case strAction
		Case "Exist"
		   sExistingRevRule = Split(sExistingRevRule,"~")
		   For iCounter = 0 To UBound(sExistingRevRule)
		   	   ObjProdImgLink.SetTOProperty "text",sExistingRevRule(iCounter)
		       If Fn_UI_ObjectExist("Fn_PC_BreadcrumbOperations",ObjProdImgLink) = True Then
					Fn_PC_BreadcrumbOperations = True
			   Else
					Fn_PC_BreadcrumbOperations = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_PC_BreadcrumbOperations : Failed to verify Link in Breadcrumb [ " & sExistingRevRule(iCounter) & " ].") 	
			   		Exit For
			   End If 
		  Next		
   End Select
   
	Set ObjProdImgLink = Nothing
	
End Function
'================================================================================================================================================================
'@@	Function Name		:	Fn_SISW_PC_NavTreeTableOperations

'Description			 :	Function Used to perform operations on Nav Tree table

'Parameters			   :   	1.StrAction: Action Name
'							2.StrNode: Node path
'							3.StrColumn: Column name
'							4.StrGroup: Group name
'							5.StrFamily: Family name
'							6.StrValue: Value name or Expected value
'							7.StrPopupMenu: Popup menu
'							8.StrToolBarOption: Toolbar option to click on toolbar buttons
'							9.StrTabName - Provide the Tab Name
'
'Return Value		   : 	True or False

'Pre-requisite			:	Nav tree table should be appear

'Examples				:   bReturn=Fn_SISW_PC_NavTreeTableOperations("AddGroup","PConfig1","","MGroup1","","","","","")
'							bReturn=Fn_SISW_PC_NavTreeTableOperations("AddFamily","PConfig1:Group3","","","Family1","","","no","")
'							bReturn=Fn_SISW_PC_NavTreeTableOperations("ModifyCell","PConfig1:Group5:Family1","Comparison Mode","","","Text","","","*000069-PConfig1 (Variant Options)")
'							bReturn=Fn_SISW_PC_NavTreeTableOperations("CellVerify","PConfig1:Group5:Family1","Comparison Mode","","","Text","","","")					     
'							bReturn=Fn_SISW_PC_NavTreeTableOperations("VerifyValuesFromCellList","PConfig1:Group5:Family1","Unit Of Measure","","","ft~gm~kg~km~ml","","","*000069-PConfig1 (Variant Options)")
'							bReturn=Fn_SISW_PC_NavTreeTableOperations("VerifySetValueForCell","PConfig1:Group5:Family1","Comparison Mode","","","","Yes","","*000069-PConfig1 (Variant Options)")
'							bReturn=Fn_SISW_PC_NavTreeTableOperations("IsCellEditable","PConfig1:Group5:Family1","Comparison Mode","","","","","","*000069-PConfig1 (Variant Options)")
'							bReturn=Fn_SISW_PC_NavTreeTableOperations("GetTreeNodeNames","","","","","","","","")
'							bReturn=Fn_SISW_PC_NavTreeTableOperations("AddValue","PConfig1:Group5:Family1","","","","Value1","","","")
'
'History					 :			
'				Developer Name						Date					Rev. No.		Changes Done																				
'================================================================================================================================================================
'				Poonam Chopade						05-Jan-2017				1.0				Created																											
'================================================================================================================================================================
Function Fn_SISW_PC_NavTreeTableOperations(StrAction,StrNode,StrColumn,StrGroup,StrFamily,StrValue,StrPopupMenu,StrToolBarOption,StrTabName)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_PC_NavTreeTableOperations"
 	'Declaring variables
	Dim bFlag,iCounter,iWidth,iTempWidth,iRowNumber,iHieght,iLoopCounter
	Dim sPath,arrNode,iCount,iTempInstance,arrNode1,iInstance,iTempHieght,iCount1
	Dim ObjNavTreeTable,ObjTree,objTableColumn,ObjSubTree,sAppVal
	Dim arrValue,objConfigWin,sMenu,sNodeName,sTreeNodes

	'Creating object of [ Nav tree table ]
	Set objConfigWin = Fn_PC_GetObject("ProductConfigurator") 
	Set ObjNavTreeTable = Fn_PC_GetObject("ConfiguratorNavTreeTable")
	Fn_SISW_PC_NavTreeTableOperations=False	

	Select Case StrAction
		Case "ModifyTextCell"
			If StrTabName <> "" Then
				Call Fn_PC_CompnentTabOperations("Maximize",StrTabName,"") 
			    Call Fn_ReadyStatusSync(1)	
			End If
			
			'Get column value from TestData
			bFlag = Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\AutomationXML\ApplicationInformationXML\RAC_ConfiguratorTable_APL.xml",StrColumn)
			If bFlag <> False Then
				StrColumn = bFlag
			End If
			bFlag = Fn_SISW_PC_NavTreeTableOperations("ClickCell",StrNode,StrColumn,"","",StrValue,"","","")
			JavaWindow("ProductConfigurator").JavaEdit("Text").SetTOProperty "path","Text;Tree;Composite;ContributedPartRenderer\$1;Composite;CTabFolder;Composite;Composite;Composite;Composite;Composite;Shell;"

			If JavaWindow("ProductConfigurator").JavaEdit("Text").Exist(5) Then
				'Checking specific value available in list
				JavaWindow("ProductConfigurator").JavaEdit("Text").Set StrValue
				Call Fn_ReadyStatusSync(5)
				Call Fn_KeyBoardOperation("SendKey", "{ENTER}")
				Call Fn_ReadyStatusSync(5)
				Fn_SISW_PC_NavTreeTableOperations = bFlag
			Else
				Fn_SISW_PC_NavTreeTableOperations=False
			End If
		'================================================================================================================================================================
		'Case to modify specific cell value
		Case "ModifyCell","ModifyCellExt"
			
			'Maximise the Tab
			If StrTabName <> "" Then
				Call Fn_PC_CompnentTabOperations("Maximize",StrTabName,"") 
			    Call Fn_ReadyStatusSync(1)	
			End If
			
			'Get column value from TestData
			bFlag = Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\AutomationXML\ApplicationInformationXML\RAC_ConfiguratorTable_APL.xml",StrColumn)
			If bFlag <> False Then
				StrColumn = bFlag
			End If
			
			IF StrColumn = "Feature Data Type"OR StrColumn = "Type" OR  StrColumn = "Feature Data type" Then
				Fn_SISW_PC_NavTreeTableOperations = Fn_SISW_PC_NavTreeTableOperations("ModifyJavaListCell",StrNode,StrColumn,"","",StrValue,"","","")
			
			Else
				'Click on cell
			bFlag = Fn_SISW_PC_NavTreeTableOperations("ClickCell",StrNode,StrColumn,"","",StrValue,"","","")
			Call Fn_KeyBoardOperation("SendKey", "{ENTER}")
			Call Fn_KeyBoardOperation("SendKey", "{ENTER}")
			
			Call Fn_ReadyStatusSync(1)
			If bFlag=True Then
				'checking Edit box is exist
				If Fn_UI_ObjectExist("Fn_SISW_PC_NavTreeTableOperations",objConfigWin.JavaEdit("NavTreeTableEdit")) Then
					'Set value in edit box
					Fn_SISW_PC_NavTreeTableOperations = Fn_Edit_Box("Fn_SISW_PC_NavTreeTableOperations",objConfigWin,"NavTreeTableEdit", StrValue)
					Call Fn_ReadyStatusSync(1)
					Call Fn_KeyBoardOperation("SendKey", "{ENTER}")
					Fn_SISW_PC_NavTreeTableOperations=True
				End If
				
				'checking Existance of list
				If Fn_UI_ObjectExist("Fn_SISW_PC_NavTreeTableOperations",objConfigWin.JavaList("NavTreeTableList")) Then
					'Checking specific value available in list
                    Fn_SISW_PC_NavTreeTableOperations=Fn_List_Select("Fn_SISW_PC_NavTreeTableOperations",objConfigWin, "NavTreeTableList",StrValue)
                    Call Fn_KeyBoardOperation("SendKey", "{ENTER}")
                    Call Fn_ReadyStatusSync(1)
				Else
					bFlag = Fn_SISW_PC_NavTreeTableOperations("ClickCell",StrNode,StrColumn,"","",StrValue,"","","")
					Call Fn_KeyBoardOperation("SendKey", "{ENTER}")
					Call Fn_KeyBoardOperation("SendKey", "{ENTER}")
					
					Call Fn_ReadyStatusSync(1)
					If StrAction <> "ModifyCellExt" Then
						if Fn_UI_ObjectExist("Fn_SISW_PC_NavTreeTableOperations",objConfigWin.JavaList("NavTreeTableList")) Then
							'Checking specific value available in list
							Fn_SISW_PC_NavTreeTableOperations = Fn_List_Select("Fn_SISW_PC_NavTreeTableOperations", objConfigWin, "NavTreeTableList",StrValue)
							Call Fn_ReadyStatusSync(5)
							If Fn_SISW_PC_NavTreeTableOperations Then
								Call Fn_KeyBoardOperation("SendKey", "{ENTER}")
								Call Fn_ReadyStatusSync(5)
							End If
						Else
							Fn_SISW_PC_NavTreeTableOperations=False
						End If
					Else
						 Fn_SISW_PC_NavTreeTableOperations = bFlag
					End If
				End If
			End If
			
			
			
			
			End IF

		 'Minimize the Tab
		  If StrTabName <> "" Then
				Call Fn_PC_CompnentTabOperations("Minimize",StrTabName,"") 
				Call Fn_ReadyStatusSync(1)	
		   End If
		'================================================================================================================================================================
		'Case to add new group
		Case "AddGroup"
			bFlag=False
			'Click on ( Add Group ) button
			If lcase(StrToolBarOption)="no" then
				bFlag=True
			Else
			    sMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("Configurator_Toolbar"), "AddGroup")
				bFlag = Fn_ToolbarOperation("Click",sMenu,"")
				Call Fn_ReadyStatusSync(1)
			End If
			If bFlag=true then
				bFlag=False
				For iCounter=0 to Cint(ObjNavTreeTable.GetROProperty("items count"))-1
				aNode = Split(Trim(ObjNavTreeTable.GetItem(iCounter)), ":")
					If aNode(uBound(aNode)) = ""  Then
						If StrNode <> "" Then
							bFlag=Fn_SISW_PC_NavTreeTableOperations("ClickCell",StrNode+":","ID",StrGroup,"","","","","")
						Else
							bFlag=Fn_SISW_PC_NavTreeTableOperations("ClickCell","","ID",StrGroup,"","","","","")
						End If
						Exit for
					End If
				Next
				Call Fn_ReadyStatusSync(1)
				If bFlag=True Then
					'Setting group name
					If Fn_UI_ObjectExist("Fn_SISW_PC_NavTreeTableOperations",objConfigWin.JavaEdit("NavTreeTableEdit")) Then
						Call Fn_Edit_Box("Fn_SISW_PC_NavTreeTableOperations",objConfigWin,"NavTreeTableEdit", StrGroup)
						Call Fn_ReadyStatusSync(1)
						Call Fn_KeyBoardOperation("SendKey", "{ENTER}")
						Fn_SISW_PC_NavTreeTableOperations=True
					End if	
				End If
			End if
		'================================================================================================================================================================
		'Case to add new Family
		Case "AddFamily"
			bFlag=False
			bFlag = Fn_PC_NavTree_NodeOperation("Select",StrNode,"")
			If bFlag=True Then
				'click on Add Family button
				If lcase(StrToolBarOption)="no" then
					bFlag=True
				Else
				    sMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("Configurator_Toolbar"), "AddFamily")
					bFlag=Fn_ToolbarOperation("Click",sMenu,"")
					Call Fn_ReadyStatusSync(1)
				End If
				If bFlag=True then
					bFlag=False
					bFlag=Fn_SISW_PC_NavTreeTableOperations("ClickCell",StrNode+":","ID","",StrFamily,"","","","")
					If bFlag=true Then
						'Set family  name
						If Fn_UI_ObjectExist("Fn_SISW_PC_NavTreeTableOperations",objConfigWin.JavaEdit("NavTreeTableEdit")) Then
								Call Fn_Edit_Box("Fn_SISW_PC_NavTreeTableOperations",objConfigWin,"NavTreeTableEdit", StrFamily)
								Call Fn_ReadyStatusSync(1)
								Call Fn_KeyBoardOperation("SendKey", "{ENTER}")
								Fn_SISW_PC_NavTreeTableOperations=True
						End If
					End If
				End if
			End If
		'================================================================================================================================================================
		'Case to add new Value
		Case "AddValue"
			bFlag=False
			bFlag = Fn_PC_NavTree_NodeOperation("Select",StrNode,"")
			If bFlag=True Then
				'click on Add Family button
				If lcase(StrToolBarOption)="no" then
					bFlag=True
				Else
				    sMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("Configurator_Toolbar"), "AddValue")
					bFlag=Fn_ToolbarOperation("Click",sMenu,"")
					Call Fn_ReadyStatusSync(1)
				End If
				If bFlag=True then
					bFlag=False
					bFlag=Fn_SISW_PC_NavTreeTableOperations("ClickCell",StrNode+":","ID","",StrValue,"","","","")
					If bFlag=true Then
						'Set Value name
						If Fn_UI_ObjectExist("Fn_SISW_PC_NavTreeTableOperations",objConfigWin.JavaEdit("NavTreeTableEdit")) Then
								Call Fn_Edit_Box("Fn_SISW_PC_NavTreeTableOperations",objConfigWin,"NavTreeTableEdit", StrValue)
								Call Fn_ReadyStatusSync(1)
								Call Fn_KeyBoardOperation("SendKey", "{ENTER}")
								Fn_SISW_PC_NavTreeTableOperations=True
						End If
					End If
				End if
			End If			
		'================================================================================================================================================================
		Case "VerifyValuesFromCellList"
			'Maximise the Tab
			If StrTabName <> "" Then
				Call Fn_PC_CompnentTabOperations("Maximize",StrTabName,"") 
			    Call Fn_ReadyStatusSync(1)	
			End If
			'Click on cell
			bFlag=Fn_SISW_PC_NavTreeTableOperations("ClickCell",StrNode,StrColumn,"","",StrValue,"","","")
			Call Fn_ReadyStatusSync(1)
			If bFlag=True Then

				If Fn_UI_ObjectExist("Fn_SISW_PC_NavTreeTableOperations",objConfigWin.JavaList("NavTreeTableList")) Then
					arrValue = split(StrValue,"~")
					For iCounter = 0 to ubound(arrValue)
						bFlag = False
						For iCount=0 to objConfigWin.JavaList("NavTreeTableList").GetROProperty("items count")-1
							If trim(objConfigWin.JavaList("NavTreeTableList").Object.getItem(iCount)) = Trim(arrValue(iCounter)) then
								bFlag=True
								Exit for
							End if
						Next
						If bFlag=False Then
							Exit for
						End If
					Next
				 Else
				  For iCount1 = 1 to 20
				    objConfigWin.JavaWindow("Shell").SetTOProperty "index",iCount1
					 If Fn_UI_ObjectExist("Fn_SISW_PC_NavTreeTableOperations",objConfigWin.JavaWindow("Shell").JavaList("NavTreeTableList")) Then
					 	Exit for 
				End If
				  next
				  If Fn_UI_ObjectExist("Fn_SISW_PC_NavTreeTableOperations",objConfigWin.JavaWindow("Shell").JavaList("NavTreeTableList")) Then
				    arrValue = split(StrValue,"~")
					For iCounter = 0 to ubound(arrValue)
						bFlag = False
						For iCount=0 to objConfigWin.JavaWindow("Shell").JavaList("NavTreeTableList").GetROProperty("items count")-1
							If trim(objConfigWin.JavaWindow("Shell").JavaList("NavTreeTableList").Object.getItem(iCount)) = Trim(arrValue(iCounter)) then
								bFlag=True
								Exit for
			End if
						Next
						If bFlag=False Then
							Exit for
						End If
					Next
				  End if
				End If
			End if
			If bFlag=True Then
				Fn_SISW_PC_NavTreeTableOperations = True
			End If
			'Minimize the Tab
		  If StrTabName <> "" Then
			  Call Fn_PC_CompnentTabOperations("Minimize",StrTabName,"") 
			  Call Fn_ReadyStatusSync(1)	
		  End If
	   '================================================================================================================================================================
		'Case to check specific cell is editable or not
		Case "IsCellEditable"
		       'Maximise the Tab
				If StrTabName <> "" Then
					Call Fn_PC_CompnentTabOperations("Maximize",StrTabName,"") 
				    Call Fn_ReadyStatusSync(1)	
				End If
				bFlag=Fn_SISW_PC_NavTreeTableOperations("ClickCell",StrNode,StrColumn,"","","","","","")
				If bFlag=True Then
					'cheking cell editable or not
					If Fn_UI_ObjectExist("Fn_SISW_PC_NavTreeTableOperations",objConfigWin.JavaEdit("NavTreeTableEdit")) Then
						Fn_SISW_PC_NavTreeTableOperations=True
					ElseIf Fn_UI_ObjectExist("Fn_SISW_PC_NavTreeTableOperations",objConfigWin.JavaList("NavTreeTableList")) Then 
						Fn_SISW_PC_NavTreeTableOperations=True
					End if
				End if
				Call Fn_SISW_PC_NavTreeTableOperations("ClickCell",StrNode,"ID","","","","","","")
				Call Fn_ReadyStatusSync(1)
				'Minimize the Tab
			  If StrTabName <> "" Then
				  Call Fn_PC_CompnentTabOperations("Minimize",StrTabName,"") 
				  Call Fn_ReadyStatusSync(1)	
			  End If
		'================================================================================================================================================================
		Case "CellVerify"
				sPath = Fn_UI_JavaTreeGetItemPathExt("Fn_SISW_PC_NavTreeTableOperations",ObjNavTreeTable,StrNode,"","")
				If sPath <> False Then
				    StrColumn = Split(StrColumn,"~")
				    StrValue = Split(StrValue,"~")
				    For iCounter = 0 To UBound(StrColumn)
				    	If trim(ObjNavTreeTable.GetColumnValue(sPath,StrColumn(iCounter))) = trim(StrValue(iCounter)) Then
							Fn_SISW_PC_NavTreeTableOperations = True
						ElseIf cInt(ObjNavTreeTable.GetColumnValue(sPath,StrColumn(iCounter))) = cInt(StrValue(iCounter)) Then
							Fn_SISW_PC_NavTreeTableOperations = True
						Else
							Fn_SISW_PC_NavTreeTableOperations = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Fn_SISW_PC_NavTreeTableOperations : Failed to verify column value [ " & StrColumn(iCounter) & " = "& StrValue(iCounter) &" ].") 	
							Exit For	
						End If
				    Next
				End If
		'================================================================================================================================================================
		Case "VerifySetValueForCell"
				'Maximise the Tab
				If StrTabName <> "" Then
					Call Fn_PC_CompnentTabOperations("Maximize",StrTabName,"") 
				    Call Fn_ReadyStatusSync(1)	
				End If
			
				bFlag=Fn_SISW_PC_NavTreeTableOperations("ClickCell",StrNode,StrColumn,"","","","","","")
				If bFlag = True Then
					sAppVal = Fn_SISW_UI_JavaList_Operations("Fn_SISW_PC_NavTreeTableOperations","GetText",objConfigWin,"NavTreeTableList","","","")
			    	If trim(sAppVal) = trim(StrValue) Then
						Fn_SISW_PC_NavTreeTableOperations = True
					Else
						Fn_SISW_PC_NavTreeTableOperations = False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Fn_SISW_PC_NavTreeTableOperations : Failed to verify column value [ " & StrColumn & " = "& StrValue &" ].") 	
					End If
				End If
				
			  'Minimize the Tab
			  If StrTabName <> "" Then
					Call Fn_PC_CompnentTabOperations("Minimize",StrTabName,"") 
					Call Fn_ReadyStatusSync(1)	
			  End If		   
		'================================================================================================================================================================
		'Get Tree Node Names when it is autogenerated with IDs
		Case "GetTreeNodeNames"
		      For iCounter=0 to Cint(ObjNavTreeTable.GetROProperty("items count"))-1
		      	   sNodeName = ObjNavTreeTable.GetItem(iCounter)	
		       	   If iCounter = 0 Then 
		       	   		 sTreeNodes = ObjNavTreeTable.GetColumnValue(sNodeName,"ID")
		       	   Else
		       	   		sTreeNodes = sTreeNodes & "~" & ObjNavTreeTable.GetColumnValue(sNodeName,"ID")
		       	   End If
		       Next
			  Fn_SISW_PC_NavTreeTableOperations = sTreeNodes		
		'================================================================================================================================================================
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
				iRowNumber=Fn_SISW_PC_NavTreeTableOperations("GetRowNumber",StrNode,"","","","","","","")
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
		'================================================================================================================================================================
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
		'================================================================================================================================================================
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
		'================================================================================================================================================================
		Case "AddModelFamily" 'Case to Model Family
			bFlag=False
			If lcase(StrToolBarOption)="no" then 'Click on ( Add Model Family ) button
				bFlag=True
			Else
			    sMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("Configurator_Toolbar"), "AddModelFamily")
				bFlag = Fn_ToolbarOperation("Click",sMenu,"")
				Call Fn_ReadyStatusSync(1)
			End If
			If bFlag=true then
				bFlag=False
				For iCounter=0 to Cint(ObjNavTreeTable.GetROProperty("items count"))-1
				aNode = Split(Trim(ObjNavTreeTable.GetItem(iCounter)), ":")
					If aNode(uBound(aNode)) = ""  Then
						If StrNode <> "" Then
							bFlag=Fn_SISW_PC_NavTreeTableOperations("ClickCell",StrNode+":","ID",StrGroup,"","","","","")
						Else
							bFlag=Fn_SISW_PC_NavTreeTableOperations("ClickCell","","ID",StrGroup,"","","","","")
						End If
						Exit for
					End If
				Next
				Call Fn_ReadyStatusSync(1)
				If bFlag=True Then 'Setting Model Family Name
					If Fn_UI_ObjectExist("Fn_SISW_PC_NavTreeTableOperations",objConfigWin.JavaEdit("NavTreeTableEdit")) Then
						Call Fn_Edit_Box("Fn_SISW_PC_NavTreeTableOperations",objConfigWin,"NavTreeTableEdit", StrGroup)
						Call Fn_ReadyStatusSync(1)
						Call Fn_KeyBoardOperation("SendKey", "{ENTER}")
						Fn_SISW_PC_NavTreeTableOperations=True
					End if	
				End If
			End if
		'================================================================================================================================================================
		Case "AddModel" 'Case to add Model
			bFlag=False
			bFlag = Fn_PC_NavTree_NodeOperation("Select",StrNode,"")
			If bFlag=True Then
				'click on Add Model button
				If lcase(StrToolBarOption)="no" then
					bFlag=True
				Else
				    sMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("Configurator_Toolbar"), "AddModel")
					bFlag=Fn_ToolbarOperation("Click",sMenu,"")
					Call Fn_ReadyStatusSync(1)
				End If
				If bFlag=True then
					bFlag=False
					bFlag=Fn_SISW_PC_NavTreeTableOperations("ClickCell",StrNode+":","ID","",StrFamily,"","","","")
					If bFlag=true Then
						'Set Model  name
						If Fn_UI_ObjectExist("Fn_SISW_PC_NavTreeTableOperations",objConfigWin.JavaEdit("NavTreeTableEdit")) Then
								Call Fn_Edit_Box("Fn_SISW_PC_NavTreeTableOperations",objConfigWin,"NavTreeTableEdit", StrFamily)
								Call Fn_ReadyStatusSync(1)
								Call Fn_KeyBoardOperation("SendKey", "{ENTER}")
								Fn_SISW_PC_NavTreeTableOperations=True
						End If
					End If
				End if
			End If
		'================================================================================================================================================================
		Case "GetListValues"
				'Maximise the Tab
				If StrTabName <> "" Then
					Call Fn_PC_CompnentTabOperations("Maximize",StrTabName,"") 
					Call Fn_ReadyStatusSync(1)	
				End If
				'Get column value from TestData
				bFlag = Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\AutomationXML\ApplicationInformationXML\RAC_ConfiguratorTable_APL.xml",StrColumn)
				If bFlag <> False Then
					StrColumn = bFlag
				End If
				'Click on cell
				bFlag = Fn_SISW_PC_NavTreeTableOperations("ClickCell",StrNode,StrColumn,"","",StrValue,"","","")
				Call Fn_KeyBoardOperation("SendKey", "{ENTER}")
				Call Fn_KeyBoardOperation("SendKey", "{ENTER}")
				Call Fn_ReadyStatusSync(1)
				If bFlag=True Then
					If Fn_UI_ObjectExist("Fn_SISW_PC_NavTreeTableOperations",objConfigWin.JavaList("NavTreeTableList")) Then
					   Fn_SISW_PC_NavTreeTableOperations = Fn_SISW_UI_JavaList_Operations("Fn_SISW_PC_NavTreeTableOperations", "GetContents", objConfigWin, "NavTreeTableList", "", "", "")						   
					   Call Fn_KeyBoardOperation("SendKey", "{ENTER}")
					   Call Fn_ReadyStatusSync(1)
					End If
				End If	
			 'Minimize the Tab
			  If StrTabName <> "" Then
					Call Fn_PC_CompnentTabOperations("Minimize",StrTabName,"") 
					Call Fn_ReadyStatusSync(1)	
			  End If
	'================================================================================================================================================================   
	Case "ModifyJavaListCell"
			'Maximise the Tab
			If StrTabName <> "" Then
				Call Fn_PC_CompnentTabOperations("Maximize",StrTabName,"") 
			    Call Fn_ReadyStatusSync(1)	
			End If
			'Get column value from TestData
			bFlag = Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\AutomationXML\ApplicationInformationXML\RAC_ConfiguratorTable_APL.xml",StrColumn)
			If bFlag <> False Then
				StrColumn = bFlag
			End If
			'Click on cell
			bFlag = Fn_SISW_PC_NavTreeTableOperations("ClickCell",StrNode,StrColumn,"","",StrValue,"","","")
			Wait 2
			Call Fn_ReadyStatusSync(1)
			
			If bFlag=True Then
				For iCount1 = 0 to 10
					objConfigWin.JavaWindow("Shell").SetTOProperty "index",iCount1
					If Fn_UI_ObjectExist("Fn_SISW_PC_NavTreeTableOperations",objConfigWin.JavaWindow("Shell").JavaList("NavTreeTableList")) Then
						Fn_SISW_PC_NavTreeTableOperations = Fn_List_Select("Fn_SISW_PC_NavTreeTableOperations",objConfigWin.JavaWindow("Shell"), "NavTreeTableList",StrValue)
						Call Fn_KeyBoardOperation("SendKey", "{ENTER}")
						Call Fn_ReadyStatusSync(1)
					End If
				Next
			End If	
		 'Minimize the Tab
		  If StrTabName <> "" Then
				Call Fn_PC_CompnentTabOperations("Minimize",StrTabName,"") 
				Call Fn_ReadyStatusSync(1)	
		   End If
	'================================================================================================================================================================   
	End Select

	'Releasing object of Nav tree table
	Set ObjNavTreeTable = Nothing
	Set objConfigWin = Nothing
End Function
'================================================================================================================================================================
'@@	Function Name		:	Fn_PC_ErrorVerify
'@@
'@@ Description			:	Function Used to Verify Error Message
'@@
'@@ Parameters			:   1.dicErrorInfo : Dictionary Object
'@@
'@@ Return Value		: 	True Or False
'@@
'@@ Examples			:   Set dicErrorInfo = CreateObject("Scripting.Dictionary")
'@@							With dicErrorInfo 
'@@							   .Add "Title", "Failed to save variant options"
'@@							   .Add "Message", sErrMsg
'@@							   .Add "Action", "ErrorVerify"
'@@							  End with
'@@							bReturn = Fn_PC_ErrorVerify(dicErrorInfo)
'@@					   
'================================================================================================================================================================
'History			:	Developer Name				Date				Rev. No.		Changes Done			Reviewer
'================================================================================================================================================================
'Created By 		:	Poonam Chopade				10-Jan-2017			1.0				Created					
'================================================================================================================================================================
Public Function Fn_PC_ErrorVerify(dicErrorInfo)
			GBL_FAILED_FUNCTION_NAME="Fn_PC_ErrorVerify"
			Dim  dicKeys, dicItems, iCounter , sAction,sTitle, sErrorMsg,sButton,objError
			
			dicKeys = dicErrorInfo.Keys
			dicItems = dicErrorInfo.Items
			For  iCounter=0 to dicErrorInfo.Count-1
				Select Case dicKeys(iCounter)
					Case "Action"
							sAction = dicItems(iCounter)
					Case "Title"
							sTitle = dicItems(iCounter)
					Case "Message"
							sErrorMsg= dicItems(iCounter)
							GBL_EXPECTED_MESSAGE=sErrorMsg
					Case "Button"
							sButton = dicItems(iCounter)
				End Select
			Next		
			If sButton = "" Then
				sButton = "OK"
			End If

			On Error Resume Next
			Set objError = Fn_PC_GetObject("PCErrorWindow")
			Fn_PC_ErrorVerify = False

			Select Case sAction
				Case "ErrorVerify"
					objError.SetTOProperty "title",sTitle
					
					If trim(dicErrorInfo("EditAttText")) = "blank" Then
						objError.JavaEdit("DetailsMsg").SetTOProperty "attached text",""
					ElseIf trim(dicErrorInfo("EditAttText")) = "" Then
						objError.JavaEdit("DetailsMsg").SetTOProperty "attached text","Details"
					ElseIf trim(dicErrorInfo("EditAttText")) <> "" Then
						objError.JavaEdit("DetailsMsg").SetTOProperty "attached text",trim(dicErrorInfo("EditAttText"))
					End If
					
					If Fn_UI_ObjectExist("Fn_PC_ErrorVerify",objError) Then
						If dicErrorInfo("Message") <> "" Then
							If instr(1,trim(objError.JavaEdit("DetailsMsg").GetROProperty("text")),trim(dicErrorInfo("Message"))) Then
								Fn_PC_ErrorVerify = TRUE
							Else 
								GBL_ACTUAL_MESSAGE=objError.JavaEdit("DetailsMsg").GetROProperty("text")
							End If
						End If
						Call Fn_Button_Click("Fn_PC_ErrorVerify",objError,sButton)
						Call Fn_ReadyStatusSync(1)
					End If	
		End Select
		
Set objError = Nothing	
		
End Function
'================================================================================================================================================================
'@@	Function Name		:	Fn_PC_CompnentTabOperations
'@@
'@@ Description			:	Function Used to Perform operation on Tab
'@@
'@@ Parameters			:   1.sAction : Action Name
'@@							2.sTCComponentTabName = Tab Name
'@@							3.sPopupMenu = Popupmenu
'@@
'@@ Return Value		: 	True Or False
'@@
'@@ Examples			:   Call Fn_PC_CompnentTabOperations("Maximize","Variant Options","") 
'@@					   		Call Fn_PC_CompnentTabOperations("Minimize","Variant Options","") 
'================================================================================================================================================================
'History			:	Developer Name				Date				Rev. No.		Changes Done			Reviewer
'================================================================================================================================================================
'Created By 		:	Poonam Chopade				11-Jan-2017			1.0				Created					
'================================================================================================================================================================
Public Function Fn_PC_CompnentTabOperations(sAction,sTCComponentTabName, sPopupMenu) 
	GBL_FAILED_FUNCTION_NAME="Fn_PC_CompnentTabOperations"
	Fn_PC_CompnentTabOperations = False
	
	Select Case sAction
	
		Case "Maximize"
		     If Fn_SISW_UI_RACTabFolderWidget_Operation("IsMaximized",sTCComponentTabName,sPopupMenu) = False Then
				Fn_PC_CompnentTabOperations = Fn_SISW_UI_RACTabFolderWidget_Operation("DoubleClick",sTCComponentTabName,sPopupMenu)
				Call Fn_ReadyStatusSync(1)
			End If	
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "Minimize"
			If Fn_SISW_UI_RACTabFolderWidget_Operation("IsMaximized",sTCComponentTabName,sPopupMenu) Then
				Fn_PC_CompnentTabOperations = Fn_SISW_UI_RACTabFolderWidget_Operation("DoubleClick",sTCComponentTabName,sPopupMenu)
				Call Fn_ReadyStatusSync(1)
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Invalid ACTION  [" +sAction+ "] is Requested.")
	End Select
	
	IF Fn_PC_CompnentTabOperations <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_PC_CompnentTabOperations : Executed successfully with Case [" + sAction + "].")
	End If

End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_PC_SetRuleDate()
'@@
'@@    Description			:	Function to set Date Rule in configurator Perspective
'@@
'@@    Parameters			:	1.sAction	: Action Name
'@@ 							2.sDate : Date to set
'@@ 							3.sTime : Time to set
'@@ 							4.sButton : Button name
'@@
'@@    Return Value		   	: 	True Or False
'@@
'@@    Pre-requisite		:	Product Configurator Perspective Should be opened
'@@
'@@    Examples				:	bReturn = Fn_PC_SetRuleDate("SetRuleDate","","","No Date")
'@@ 							
'@@ 								
'@@-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@	   History				:	Developer Name				Date	  			Rev. No.		Changes Done		Reviewer
'@@-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@	   Created By 			:	Poonam Chopade			08-March-2019			  1.0		  	 Created		  [TC12.1(20181213.00)-08Mar2019-PoonamC-NewDevelopment]
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_PC_SetRuleDate(sAction,sDate,sTime,sButton)

	GBL_FAILED_FUNCTION_NAME = "Fn_PC_SetRuleDate"
	Dim objSetRuleDate,objPCWindow
	
	Fn_PC_SetRuleDate = False
	
	Set objPCWindow = Fn_PC_GetObject("ProductConfigurator")
	Set objSetRuleDate = objPCWindow.JavaWindow("Set Rule Date")
	
	Select Case sAction
		
		Case "SetRuleDate","SetRuleDateWithoutEmptyDateAndTime"
				'Check Existence of Date Rule set to system Default
				If Fn_UI_ObjectExist("Fn_PC_SetRuleDate",objPCWindow.JavaObject("DateTimeImageHyperlink_NoRuleDate")) = True Then
						Call Fn_UI_JavaObject_Click("Fn_PC_SetRuleDate",objPCWindow,"DateTimeImageHyperlink_NoRuleDate",5,5,"LEFT")
						Call Fn_ReadyStatusSync(1)
						'objPCWindow.WinMenu("ContextMenu").Select "Set Rule Date"
						objPCWindow.JavaMenu("Label:=Set Rule Date").Select
						Wait(1)
						If sAction = "SetRuleDateWithoutEmptyDateAndTime" Then
							Call Fn_SISW_UI_JavaEdit_Operations("Fn_PC_SetRuleDate","Set",objSetRuleDate,"Date","")
						End If
						'Check Date Text
						If sDate <> "" Then
							Call Fn_SISW_UI_JavaEdit_Operations("Fn_PC_SetRuleDate","Type",objSetRuleDate,"Date",sDate)
							Call Fn_ReadyStatusSync(1)
						End If
						'Check Time Text
						If sTime <> "" Then
							Call Fn_SISW_UI_JavaList_Operations("Fn_PC_SetRuleDate","Select",objSetRuleDate,"Time",sTime,"","")
							Call Fn_ReadyStatusSync(1)
						End If
						'Click on button OK / Cancel / No Date
						If sButton <> "" Then
							objSetRuleDate.JavaButton("Button").SetTOProperty "label",sButton
							Fn_PC_SetRuleDate = Fn_Button_Click("Fn_PC_SetRuleDate",objSetRuleDate,"Button")
							Call Fn_ReadyStatusSync(1)
						End If
				Else
					If Fn_UI_ObjectExist("Fn_PC_SetRuleDate",objPCWindow.JavaObject("DateTimeImageHyperlink")) = True Then
						Call Fn_UI_JavaObject_Click("Fn_PC_SetRuleDate",objPCWindow,"DateTimeImageHyperlink",5,5,"LEFT")
						Call Fn_ReadyStatusSync(1)
						'objPCWindow.WinMenu("ContextMenu").Select "Set Rule Date"
						objPCWindow.JavaMenu("Label:=Set Rule Date").Select
						Wait(1)
						
						'Click on button OK / Cancel / No Date
						If sButton <> "" Then
							objSetRuleDate.JavaButton("Button").SetTOProperty "label",sButton
							Fn_PC_SetRuleDate = Fn_Button_Click("Fn_PC_SetRuleDate",objSetRuleDate,"Button")
							Call Fn_ReadyStatusSync(1)
						End If
					End If
				End If
			
			Case "VerifyRuleDate"
			Fn_PC_SetRuleDate = FALSE
				'Check Existence of Date Rule
				If Fn_UI_ObjectExist("Fn_PC_SetRuleDate",objPCWindow.JavaObject("DateTimeImageHyperlink")) = True Then
					Call Fn_ReadyStatusSync(1)
					
					If instr(1,trim(objPCWindow.JavaObject("DateTimeImageHyperlink").getROProperty("text")),trim(sDate)) Then
						Fn_PC_SetRuleDate = TRUE
					Else 
						Fn_PC_SetRuleDate = FALSE
					End If
				End If				

		Case "VerifyNoRuleDate"
			Fn_PC_SetRuleDate = FALSE
				'Check Existence of Date Rule
				If Fn_UI_ObjectExist("Fn_PC_SetRuleDate",objPCWindow.JavaObject("DateTimeImageHyperlink_NoRuleDate")) = True Then
					Call Fn_ReadyStatusSync(1)
					
					If instr(1,trim(objPCWindow.JavaObject("DateTimeImageHyperlink_NoRuleDate").getROProperty("text")),trim(sDate)) Then
						Fn_PC_SetRuleDate = TRUE
					Else 
						Fn_PC_SetRuleDate = FALSE
					End If
				End If				

			End Select
	
	Set objPCWindow = Nothing
	Set objSetRuleDate = Nothing

End Function
'============================================================================================================================================
'Function Name		:	Fn_PC_RevisionRuleOperations
'
'Description		:	Function Used to Set Revision Rule
'
'Parameters			:  1.sAction			: Action Name
'					   2.sExistingRevRule	: Existing Revision Rule Name
'					   3.sNewRevRule		: New Revision Rule to be selected.
'
'Return Value		: 	True Or False
'
'Pre-requisite		:	Product Configurator perspective should be activated.
'
' Examples			:   Call Fn_PC_RevisionRuleOperations("Set", "Any Status; No Working", "Any Status; Working")
'
'History			:	Developer Name				Date			Changes Done		Reviewer
'============================================================================================================================================
'						Poonam Chopade				06-Aug-2019		Created			Tc12.3_2019070800_NewDevelopment_PoonamC_06Aug2019
'============================================================================================================================================
Public Function Fn_PC_RevisionRuleOperations(sAction, sExistingRevRule, sNewRevRule)
	GBL_FAILED_FUNCTION_NAME="Fn_PC_RevisionRuleOperations"
	Dim objConfigWin,aRevRule, iInstance,intNoOfObjects,iCnt,objSelectType,SIndex,SName,Flag
	Set objConfigWin = Fn_PC_GetObject("ProductConfigurator") 
	
	If sAction = "Set" Then	
			Flag = False
			Set objSelectType = Description.Create()
			objSelectType("Class Name").value = "JavaObject"
			objSelectType("toolkit class").value = "org.eclipse.ui.forms.widgets.ImageHyperlink"
			Set  intNoOfObjects = objConfigWin.ChildObjects(objSelectType)

			'Code to check if the New Revision rule is already set .
			For iCnt = 0 to intNoOfObjects.count-1
				 If Trim(intNoOfObjects(iCnt).GetROProperty("developer name")) = Trim(sNewRevRule) then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_PC_RevisionRuleOperations is already SET")
						Fn_PC_RevisionRuleOperations = True 
						Set objSelectType = Nothing
						Set intNoOfObjects = Nothing
						Exit Function
				End If
			Next

		' Code to check if the "sExistingRevRule" Rule that we pass is actually set	in the applciation
		   For iCnt = 0 to intNoOfObjects.count-1
				If Trim(intNoOfObjects(iCnt).GetROProperty("developer name")) = Trim(sExistingRevRule) then
					Flag = True
					Exit For
				End IF										

				If  Instr(lcase(Trim(intNoOfObjects(iCnt).GetROProperty("developer name"))),"working") > 0 then
					SName = Trim(intNoOfObjects(iCnt).GetROProperty("developer name"))
					SIndex = iCnt
				End If
			Next
			If Flag = False Then
				Fn_PC_RevisionRuleOperations = False
				sExistingRevRule = SName
				aRevRule = split(sExistingRevRule,"@")
				iInstance = SIndex
			Else
				Fn_PC_RevisionRuleOperations = False
				aRevRule = split(sExistingRevRule,"@")
				iInstance = 0
			End IF
	End IF
	
	Fn_PC_RevisionRuleOperations = False
	aRevRule = split(sExistingRevRule,"@")
	iInstance = 0
	
	aRevRule(0) = trim(aRevRule(0))
	
	if uBound(aRevRule) = 1 then
		iInstance = cInt(aRevRule(1))
	End If
	
	Select Case sAction
		Case "Set"
				'To resolve issues when QTP treats "(" and ")" as Special characters, thereby failing to work with Rules containing "(" or ")".
				If Instr( sNewRevRule , "(") <> 0 Or Instr( sNewRevRule , ")") <> 0 Then
					sNewRevRule = Replace(sNewRevRule , "(" , "\(")
					sNewRevRule = Replace(sNewRevRule , ")" , "\)")
				End If
	
				objConfigWin.JavaObject("RevRuleImageHyperlink").SetTOProperty "developer name", aRevRule(0)
				objConfigWin.JavaObject("RevRuleImageHyperlink").SetTOProperty "Index", iInstance
				
				objConfigWin.JavaObject("RevRuleImageHyperlink").Click 1,1,"LEFT"
				Wait 1
				'JavaWindow("ProductConfigurator").JavaMenu("Label:=" & sNewRevRule).Select 
				Fn_PC_RevisionRuleOperations = Fn_UI_JavaMenu_Select("Fn_PC_RevisionRuleOperations",objConfigWin,sNewRevRule)
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_PC_RevisionRuleOperations : successfully selected revision rule [ " & sNewRevRule & " ] ")	
	Case "ExistInMenu"
				Set objSelectType = Description.Create()
					objSelectType("Class Name").value = "JavaObject"
					objSelectType("toolkit class").value = "org.eclipse.ui.forms.widgets.ImageHyperlink"
				Set intNoOfObjects = objConfigWin.ChildObjects(objSelectType)
				For iCnt = 0 to intNoOfObjects.count-1
					If  lcase(trim( "" & intNoOfObjects(iCnt).Object.getToolTipText())) = lCase("Click to view and change the current variant configuration") Then
						intNoOfObjects(iCnt).Click 1,1, "LEFT"
						Exit for
					End If
				Next
				If Instr( sNewRevRule , "(") <> 0 Or Instr( sNewRevRule , ")") <> 0 Then
					sNewRevRule = Replace(sNewRevRule , "(" , "\(")
					sNewRevRule = Replace(sNewRevRule , ")" , "\)")
				End If
			  Fn_PC_RevisionRuleOperations = objConfigWin.JavaMenu("Label:=" & sNewRevRule).Exist(5)
			  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_PC_RevisionRuleOperations : successfully verified existence of revision rule [ " & sNewRevRule & " ] ")
			  Call Fn_KeyBoardOperation("SendKeys", "{ESC}")			  
	End Select
	
	If Fn_PC_RevisionRuleOperations <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_PC_RevisionRuleOperations : Executed successfully with Case [ " & sAction & " ] ")
	End IF
	
	objConfigWin.JavaObject("RevRuleImageHyperlink").SetTOProperty "developer name", ""
	objConfigWin.JavaObject("RevRuleImageHyperlink").SetTOProperty "Index",0
	
	Set objConfigWin = Nothing
	
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_PC_VariantNatTable_VariantExpressionEditor_Operations
'@@
'@@    Description			:	Function Used to Perform operation on Variant Expression Editor
'@@
'@@    Parameters			:	1. sAction		: Action to be performed
'@@							:	2. StrTabName	: Tab Name
'@@							:	3. dicDetails	: Dictionary object
'@@							:	4. Popupmenu	: Popup menu name
'@@							:	3. sTabClose	: Flag to close Tab
'@@
'@@    Return Value		   	: 	True Or False
'@@
'@@    Pre-requisite		:	Product Configurator Perspective should be opened 
'@@
'@@    Examples				:	Set dicDetails = CreateObject( "Scripting.Dictionary")
'@@									dicDetails("ColumnName") = "0004"
'@@									dicDetails("SubjectOption") = "a1~b1"
'@@									dicDetails("SubjectOptionFlag") = "Check~Check"
'@@									dicDetails("ApplicabilityOption") = "m1~a1"
'@@									dicDetails("ApplicabilityOptionFlag") = "Check~Check"
'@@							bReturn = Fn_PC_VariantNatTable_VariantExpressionEditor_OPerations("SetVariantExpEditorOptionAndSave","000202-PRoduct (Variant Expression Editor)",dicDetails,"","Yes")
'@@
'@@    Examples				:	Set dicDetails = CreateObject( "Scripting.Dictionary")
'@@									dicDetails("ColumnName") = "0004"
'@@									dicDetails("AvailabilityOption") = "a1~b1"
'@@									dicDetails("AvailabilityOptionFlag") = "Check~Check"
'@@							bReturn = Fn_PC_VariantNatTable_VariantExpressionEditor_OPerations("SetAvailabilityMatrixOptionAndSave","",dicDetails,"","Yes")
'@@
'@@-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@	   History				:	Developer Name				Date	  			Rev. No.		Changes Done		Reviewer
'@@-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@	   Created By 			:	Poonam Chopade			09-Aug-2019				1.0		  		 Created		  [TC12.3(2019070800)-09Aug2019-PoonamC-NewDevelopment]
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_PC_VariantNatTable_VariantExpressionEditor_Operations(sAction,StrTabName,dicDetails,Popupmenu,sTabClose)

	GBL_FAILED_FUNCTION_NAME = "Fn_PC_VariantNatTable_VariantExpressionEditor_Operations"
	
	Dim objPCWindow,arrOptionFlag,arrOptions,iSubJectInx,iApplicabilityInx,sMenu,bFlag
	Dim iCnt,StrBounds,iX,iY,iCnt1,iRowIndex,strRow,iRowsCount,iColIndex,strColumn,iColCount
	Dim myDeviceReplay,sAppMsg,iColPosition,aOption,iInstance,iCounter
	
	Fn_PC_VariantNatTable_VariantExpressionEditor_Operations = False
	On Error Resume Next
	If Fn_ToolBarOperation("IsSelected","Navigation Pane","") Then
		Call Fn_ToolBarOperation("Click","Navigation Pane","")
	End If
	
	Set objPCWindow = Fn_PC_GetObject("ProductConfigurator")
	If Fn_UI_ObjectExist("Fn_PC_VariantNatTable_VariantExpressionEditor_Operations",objPCWindow.JavaObject("VariantExpressionEditor")) = False Then
		sMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("Configurator_Toolbar"),"OpenAvailabilityMatrixview")
		Call Fn_ToolBarOperation("Click",sMenu,"")
		Call Fn_ReadyStatusSync(1)
		If Fn_UI_ObjectExist("Fn_PC_VariantNatTable_VariantExpressionEditor_Operations",objPCWindow.JavaObject("VariantExpressionEditor")) = False Then
			sMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("Configurator_Toolbar"),"OpenExpressionEditor")
			Call Fn_ToolBarOperation("Click",sMenu,"")
			Call Fn_ReadyStatusSync(1)
		End If
		
		If Fn_UI_ObjectExist("Fn_PC_VariantNatTable_VariantExpressionEditor_Operations",objPCWindow.JavaObject("VariantExpressionEditor")) = False Then
			Fn_PC_VariantNatTable_VariantExpressionEditor_Operations = False
			Set objPCWindow = Nothing
			Exit function
		End IF
	End If
	' Maximize tab
	 If StrTabName <> "" Then
	 	Call Fn_TabFolder_Operation("DoubleClickTab", StrTabName,"")
		Call Fn_ReadyStatusSync(1)
	 End If 
	
	Select Case sAction
		Case "SelectSplitColumnAndSetVarVariantExpEditor","SetVariantExpEditorOptionAndSave","SetVariantExpEditorOptionAndSave_Ext","SplitColumnAndSetVarVariantExpEditorAndSave","ModifySplitColumnAndSetVarVariantExpEditorAndSave","AddAdditionalValuesInVariantExpEditorAndSave","SetVEEForMultipleColumnsAndSave","SetVEEForMultipleColumnsAndSave_Ext","SplitColumnAndSetVarVariantExpEditorAndSave_Ext","Add_SplitColumn_AndSetVarVariantExpEditorAndSave"
		
				 '------------------- For Additional values ---------------------------------------------------------------------
				 If sAction = "AddAdditionalValuesInVariantExpEditorAndSave" or sAction = "SetVEEForMultipleColumnsAndSave" or sAction = "SetVEEForMultipleColumnsAndSave_Ext" or sAction = "SetVariantExpEditorOptionAndSave" or "Add_SplitColumn_AndSetVarVariantExpEditorAndSave" Then
				 	Call Fn_SISW_UI_JavaList_Operations("Fn_PC_VariantNatTable_VariantExpressionEditor_Operations","Select", objPCWindow,"CCombo","Show Features","","")
					Call Fn_ReadyStatusSync(1)
				 End IF
				 '---------------------------------------------------------------------------------------------------------------
				 
				 If dicDetails("ColumnName") <> "" Then  'Get column index
			 		iColCount = objPCWindow.JavaObject("VariantExpressionEditor").Object.getColumnCount()
			 		For iCnt = 0 To iColCount - 1
			 			iColPosition = objPCWindow.JavaObject("VariantExpressionEditor").Object.getstartXOfcolumnPosition(iCnt+3)
						strColumn =  CStr(objPCWindow.JavaObject("VariantExpressionEditor").Object.getCellByPosition(iCnt,0).getlayer.getColumnHeaderlayer.getDatavalueByPosition(iCnt,0).tostring)
						If trim(dicDetails("ColumnName")) = trim(strColumn) Then
								iColIndex = iCnt + 1
								Exit for
						End If
					Next
				 Else
					iColIndex = 2							 
				 End If
				
				 iRowsCount = objPCWindow.JavaObject("VariantExpressionEditor").Object.getRowCount()  'Get Index for Subject & Applicability sections
				 For iCnt = 1 To iRowsCount - 1
					If instr(objPCWindow.JavaObject("VariantExpressionEditor").Object.getCellByPosition(0,iCnt).getDataValue().getData().toString(),"Subject") > 0 Then
						iSubJectInx = iCnt
						iColCount = iCnt
						Exit For
					End IF
				 Next
				 For iCnt = iColCount To iRowsCount - 1
					If instr(objPCWindow.JavaObject("VariantExpressionEditor").Object.getCellByPosition(0,iCnt).getDataValue().getData().toString(),"Applicability") > 0 Then
						iApplicabilityInx = iCnt
						Exit For
					End IF
				 Next
				 '------------------- Code for Split column ---------------------------------------------------------------------
 				 If sAction = "SplitColumnAndSetVarVariantExpEditorAndSave" OR sAction = "SplitColumnAndSetVarVariantExpEditorAndSave_Ext" OR sAction = "Add_SplitColumn_AndSetVarVariantExpEditorAndSave" Then
 				 		 StrBounds = objPCWindow.JavaObject("VariantExpressionEditor").Object.getBoundsByPosition(iColIndex,0).tostring
						 StrBounds = Split(Replace(Replace(StrBounds,"Rectangle {",""),"}",""),",")
						 iX = cint(StrBounds(2))
						 iY = cint(StrBounds(3) - 40)
						 If IsEmpty(iX) or IsEmpty(iY) Then
							Fn_PC_VariantNatTable_VariantExpressionEditor_Operations = False
							Set objPCWindow = Nothing
							Exit Function
						 End If				
						 If sAction = "SplitColumnAndSetVarVariantExpEditorAndSave_Ext" Then
						 	 iX = iX + 204
							 iY = iY + 5
						 Else
							 iX = iX + 4
							 iY = iY + 5
						End If
						If dicDetails("GridNumber") <> "" Then
						 	 iX = dicDetails("GridNumber")
						 End If
						  
						objPCWindow.JavaObject("VariantExpressionEditor").Click iX, iY ,"RIGHT"
						 Wait 3
						 strPopupMenu = objPCWindow.WinMenu("ContextMenu").BuildMenuPath(Popupmenu)
						 objPCWindow.WinMenu("ContextMenu").Select strPopupMenu
						 Wait 3
						 '-----------------------------------------------------------------------
						 iColCount = objPCWindow.JavaObject("VariantExpressionEditor").Object.getColumnCount()
						 For iCnt = 0 To iColCount - 1
							strColumn =  CStr(objPCWindow.JavaObject("VariantExpressionEditor").Object.getCellByPosition(iCnt,0).getlayer.getColumnHeaderlayer.getDatavalueByPosition(iCnt,0).tostring)
							If instr(trim(strColumn) , trim("expressiongrid")) > 0 Then
								iColIndex = iCnt + 1
								Exit for
							End If
						 Next
 				 End If
 				 '------------------- ------------------------- ---------------------------------------------------------------------
 				 '------------------- Code for to modify Split column ---------------------------------------------------------------------
 				If sAction = "ModifySplitColumnAndSetVarVariantExpEditorAndSave" Then 
 				StrBounds = objPCWindow.JavaObject("VariantExpressionEditor").Object.getBoundsByPosition(iColIndex,0).tostring
						 StrBounds = Split(Replace(Replace(StrBounds,"Rectangle {",""),"}",""),",")
						 iX = cint(StrBounds(2))
						 iY = cint(StrBounds(3) - 40)
						 If IsEmpty(iX) or IsEmpty(iY) Then
							Fn_PC_VariantNatTable_VariantExpressionEditor_Operations = False
							Set objPCWindow = Nothing
							Exit Function
						 End If				
						 iX = iX + 4
						 iY = iY + 5
						 '-----------------------------------------------------------------------
						 iColCount = objPCWindow.JavaObject("VariantExpressionEditor").Object.getColumnCount()
						 For iCnt = 0 To iColCount - 1
							strColumn =  CStr(objPCWindow.JavaObject("VariantExpressionEditor").Object.getCellByPosition(iCnt,0).getlayer.getColumnHeaderlayer.getDatavalueByPosition(iCnt,0).tostring)
							If instr(trim(strColumn) , trim("expressiongrid")) > 0 Then
								iColIndex = iCnt + 1
								Exit for
							End If
						 Next
 				 End If
 				 '------------------- ------------------------- ---------------------------------------------------------------------
				 'Check Subject Option values 
				 If dicDetails("SubjectOption") <> "" AND dicDetails("SubjectOptionFlag") <> "" Then
					arrOptions = Split(dicDetails("SubjectOption"),"~")
					arrOptionFlag = Split(dicDetails("SubjectOptionFlag"),"~")
					For iCnt = 0 To UBound(arrOptions)
						'--------------- for multiple instance ------------------------------
				 		If instr(arrOptions(iCnt),"@") Then
				 			aOption = Split(arrOptions(iCnt),"@")
				 			arrOptions(iCnt) = aOption(0)
				 			iInstance = aOption(1)
				 		Else
							iInstance = 1				 		
				 		End If
				 		'--------------------------------------------------------------
				 		iCounter = 1
						For iCnt1 = 1 To iRowsCount
							If instr(CStr(objPCWindow.JavaObject("VariantExpressionEditor").Object.getCellByPosition(0,iCnt1).getDataValue().getData().toString()),arrOptions(iCnt)) Then
								iRowIndex = iCnt1
								'--------------------------------------------
								If cint(iCounter) = cint(iInstance) Then
					 				Exit for
					 			Else
									iCounter = iCounter + 1							 			
					 			End If
					 			'-------------------------------------------
							End If
						Next
						
						StrBounds = objPCWindow.JavaObject("VariantExpressionEditor").Object.getBoundsByPosition(iColIndex,iRowIndex).tostring
				 		StrBounds = Split(Replace(Replace(StrBounds,"Rectangle {",""),"}",""),",")
				 		iX = cint(StrBounds(2))
				 		iY = cint(StrBounds(1))
						If IsEmpty(iX) or IsEmpty(iY) Then
				 			Fn_PC_VariantNatTable_VariantExpressionEditor_Operations = False
							Set objPCWindow = Nothing
				 			Exit Function
				 		End If
				 		
						If sAction = "SplitColumnAndSetVarVariantExpEditorAndSave" Or sAction = "ModifySplitColumnAndSetVarVariantExpEditorAndSave" Or sAction = "Add_SplitColumn_AndSetVarVariantExpEditorAndSave" Then
							iX = iX + 29
						End if
						
						If sAction = "SplitColumnAndSetVarVariantExpEditorAndSave_Ext"  Then
								iX = 170
							For iCnt1 = 1 to dicDetails("GridNumber")-1
								iX = iX + 29
							Next
							End If
						
						If sAction = "SetVEEForMultipleColumnsAndSave" Then
							iX = iColPosition
							iX = iX + 4
							iY = iY + 5
						End if
						If sAction = "SetVariantExpEditorOptionAndSave" Then
							iX = iX + 4
							iY = iY + 5
						End if
						If sAction = "SetVEEForMultipleColumnsAndSave_Ext" Then
						StrBounds = objPCWindow.JavaObject("VariantExpressionEditor").Object.getBoundsByPosition(iColIndex,iRowIndex).tostring
							StrBounds = Split(Replace(Replace(StrBounds,"Rectangle {",""),"}",""),",")
				 			iX = cint(StrBounds(2))
				 			iY = cint(StrBounds(1))
'							iX = iColPosition
							iX = iX + 4
							iY = iY + 5
						End if
						If sAction = "SetVariantExpEditorOptionAndSave_Ext" Then  'For Boolean Value to check in subject section for same value
							StrBounds = objPCWindow.JavaObject("VariantExpressionEditor").Object.getBoundsByPosition(iColIndex,iRowIndex+1).tostring
							StrBounds = Split(Replace(Replace(StrBounds,"Rectangle {",""),"}",""),",")
				 			iX = cint(StrBounds(2))
				 			iY = cint(StrBounds(1))
							iX = iX + 4
							iY = iY + 5
						End if
						Select Case arrOptionFlag(iCnt)
				 			Case "Check" 'For "=" or "=Any" condition
				 				objPCWindow.JavaObject("VariantExpressionEditor").Click iX,iY,"LEFT"
				 				Wait 1
				 			Case "None" 'For "!=" or or "=NONE" condition
				 				objPCWindow.JavaObject("VariantExpressionEditor").dblClick iX,iY,"LEFT"
				 				Wait 1
				 			Case "Blank" 'to set blank
				 			   ' For future use
				 			Case else
				 				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Invalid case [ "&arrOptionFlag(iCnt)&" ]" )
								Fn_PC_VariantNatTable_VariantExpressionEditor_Operations = False
								Set objPCWindow = Nothing
								Exit function
				 		End Select		
					Next					
				 End IF
						 
				 'Check Applicability Option values 
				 If dicDetails("ApplicabilityOption") <> "" AND dicDetails("ApplicabilityOptionFlag") <> "" Then
					arrOptions = Split(dicDetails("ApplicabilityOption"),"~")
					arrOptionFlag = Split(dicDetails("ApplicabilityOptionFlag"),"~")
					For iCnt = 0 To UBound(arrOptions)
					    '--------------- for multiple instance ------------------------------
				 		If instr(arrOptions(iCnt),"@") Then
				 			aOption = Split(arrOptions(iCnt),"@")
				 			arrOptions(iCnt) = aOption(0)
				 			iInstance = aOption(1)
				 		Else
							iInstance = 1				 		
				 		End If
				 		'--------------------------------------------------------------
				 		iCounter = 1
						For iCnt1 = iApplicabilityInx To iRowsCount
							If instr(CStr(objPCWindow.JavaObject("VariantExpressionEditor").Object.getCellByPosition(0,iCnt1).getDataValue().getData().toString()),arrOptions(iCnt)) Then
								iRowIndex = iCnt1
								'--------------------------------------------
								If cint(iCounter) = cint(iInstance) Then
					 				Exit for
					 			Else
									iCounter = iCounter + 1							 			
					 			End If
					 			'-------------------------------------------
							End If
						Next
						StrBounds = objPCWindow.JavaObject("VariantExpressionEditor").Object.getBoundsByPosition(iColIndex,iRowIndex).tostring
				 		StrBounds = Split(Replace(Replace(StrBounds,"Rectangle {",""),"}",""),",")
				 		iX = cint(StrBounds(2))
				 		iY = cint(StrBounds(1))
						If IsEmpty(iX) or IsEmpty(iY) Then
				 			Fn_PC_VariantNatTable_VariantExpressionEditor_Operations = False
							Set objPCWindow = Nothing
				 			Exit Function
				 		End If
				 		
				 		If sAction = "SplitColumnAndSetVarVariantExpEditorAndSave" Or sAction = "ModifySplitColumnAndSetVarVariantExpEditorAndSave" Then
							iX = iX + 29
						End if
				 		
				 		If sAction = "SetVEEForMultipleColumnsAndSave" Then
							iX = iColPosition
							iX = iX + 4
							iY = iY + 5							
						End if
				 		If sAction = "SetVariantExpEditorOptionAndSave" Then
							iX = iX + 4
							iY = iY + 5
						End if
						If sAction = "SetVEEForMultipleColumnsAndSave_Ext" Then
							StrBounds = objPCWindow.JavaObject("VariantExpressionEditor").Object.getBoundsByPosition(iColIndex,iRowIndex).tostring
							StrBounds = Split(Replace(Replace(StrBounds,"Rectangle {",""),"}",""),",")
				 			iX = cint(StrBounds(2))
				 			iY = cint(StrBounds(1))
'							iX = iColPosition
							iX = iX + 4
							iY = iY + 5
						End if
						If sAction = "SetVariantExpEditorOptionAndSave_Ext" Then  'For Boolean Value to check in subject section for same value
							StrBounds = objPCWindow.JavaObject("VariantExpressionEditor").Object.getBoundsByPosition(iColIndex,iRowIndex+1).tostring
							StrBounds = Split(Replace(Replace(StrBounds,"Rectangle {",""),"}",""),",")
				 			iX = cint(StrBounds(2))
				 			iY = cint(StrBounds(1))
							iX = iX + 4
							iY = iY + 5
						End if
						
							 		
				 '------------------- ------------------------- ---------------------------------------------------------------------
 				 '------------------- Code for to Select Split column ---------------------------------------------------------------------
 				If sAction = "SelectSplitColumnAndSetVarVariantExpEditor" Then 
 				StrBounds = objPCWindow.JavaObject("VariantExpressionEditor").Object.getBoundsByPosition(iColIndex,0).tostring
						 StrBounds = Split(Replace(Replace(StrBounds,"Rectangle {",""),"}",""),",")
						 iX = cint(StrBounds(2))
						 iY = cint(StrBounds(3) - 40)
						 If IsEmpty(iX) or IsEmpty(iY) Then
							Fn_PC_VariantNatTable_VariantExpressionEditor_Operations = False
							Set objPCWindow = Nothing
							Exit Function
						 End If				
						 iX = iX + 200
						 iY = iY + 0
						 
						 If dicDetails("SplitColumnGrid") <> "" Then
						 	iX = iX + cint(dicDetails("SplitColumnGrid"))
						 End If
						 objPCWindow.JavaObject("VariantExpressionEditor").Click iX,iY,"LEFT"
				End If
				'-----------------------------------------------------------------------
				 	
						Select Case arrOptionFlag(iCnt)
				 			Case "Check" 'For "=" or "=Any" condition
				 				objPCWindow.JavaObject("VariantExpressionEditor").Click iX,iY,"LEFT"
				 				Wait 1
				 			Case "None" 'For "!=" or or "=NONE" condition
				 				objPCWindow.JavaObject("VariantExpressionEditor").dblClick iX,iY,"LEFT"
				 				Wait 1
				 			Case "SetBlankForChecked" 'to set blank
				 				 objPCWindow.JavaObject("VariantExpressionEditor").dblClick iX,iY,"LEFT"
				 				Wait 1
				 			Case "Blank" 'to set blank
				 			   ' For future use
				 			Case else
				 				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Invalid case [ "&arrOptionFlag(iCnt)&" ]" )
								Fn_PC_VariantNatTable_VariantExpressionEditor_Operations = False
								Set objPCWindow = Nothing
								Exit function
				 		End Select		
					Next						
				 End IF
				 Fn_PC_VariantNatTable_VariantExpressionEditor_Operations = True	
				 
				 sMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("RAC_Toolbar"),"Savethecurrentcontents")  'Save the Variant Condition
				 Call Fn_ToolBarOperation("Click",sMenu,"")
				 Call Fn_ReadyStatusSync(1)
		'===========================================================================================
		Case "SetAvailabilityMatrixOptionAndSave","SetAvailabilityMatrixOptionAndSave.ext"		 
				 If dicDetails("ColumnName") <> "" Then  'Get column index
			 		iColCount = objPCWindow.JavaObject("VariantExpressionEditor").Object.getColumnCount()
			 		For iCnt = 0 To iColCount - 1
						strColumn =  CStr(objPCWindow.JavaObject("VariantExpressionEditor").Object.getCellByPosition(iCnt,0).getlayer.getColumnHeaderlayer.getDatavalueByPosition(iCnt,0).tostring)
						If trim(dicDetails("ColumnName")) = trim(strColumn) Then
								iColIndex = iCnt + 1
								Exit for
						End If
					Next
				 Else
					iColIndex = 2							 
				 End If
				 Call Fn_SISW_UI_JavaList_Operations("Fn_PC_VariantNatTable_VariantExpressionEditor_Operations","Select", objPCWindow,"CCombo","Show Features","","")
				 Call Fn_ReadyStatusSync(1)
				 iRowsCount = objPCWindow.JavaObject("VariantExpressionEditor").Object.getRowCount()  'Get Index for Subject & Applicability sections
				 If dicDetails("AvailabilityOption") <> "" AND dicDetails("AvailabilityOptionFlag") <> "" Then
					arrOptions = Split(dicDetails("AvailabilityOption"),"~")
					arrOptionFlag = Split(dicDetails("AvailabilityOptionFlag"),"~")
					For iCnt = 0 To UBound(arrOptions)
						'--------------- for multiple instance ------------------------------
				 		If instr(arrOptions(iCnt),"@") Then
				 			aOption = Split(arrOptions(iCnt),"@")
				 			arrOptions(iCnt) = aOption(0)
				 			iInstance = aOption(1)
				 		Else
							iInstance = 1				 		
				 		End If
				 		'--------------------------------------------------------------
				 		iCounter = 1
						For iCnt1 = 1 To iRowsCount
							If instr(CStr(objPCWindow.JavaObject("VariantExpressionEditor").Object.getCellByPosition(0,iCnt1).getDataValue().getData().toString()),arrOptions(iCnt)) Then
								iRowIndex = iCnt1
								'--------------------------------------------
								If cint(iCounter) = cint(iInstance) Then
					 				Exit for
					 			Else
									iCounter = iCounter + 1							 			
					 			End If
					 			'-------------------------------------------
							End If
						Next
						
						If sAction = "SetAvailabilityMatrixOptionAndSave.ext" Then  'For Boolean Value to check in subject section for same value
							StrBounds = objPCWindow.JavaObject("VariantExpressionEditor").Object.getBoundsByPosition(iColIndex,iRowIndex+1).tostring
							StrBounds = Split(Replace(Replace(StrBounds,"Rectangle {",""),"}",""),",")
				 			iX = cint(StrBounds(2))
				 			iY = cint(StrBounds(1))
							iX = iX + 4
							iY = iY + 5
						Else
							StrBounds = objPCWindow.JavaObject("VariantExpressionEditor").Object.getBoundsByPosition(iColIndex,iRowIndex).tostring
				 			StrBounds = Split(Replace(Replace(StrBounds,"Rectangle {",""),"}",""),",")
				 			iX = cint(StrBounds(2))
				 			iY = cint(StrBounds(1))
							If IsEmpty(iX) or IsEmpty(iY) Then
					 			Fn_PC_VariantNatTable_VariantExpressionEditor_Operations = False
								Set objPCWindow = Nothing
					 			Exit Function
				 			End If
				 		End If
						iX = iX + 4
						iY = iY + 5	
						Select Case arrOptionFlag(iCnt)
				 			Case "Check" 'For "=" or "=Any" condition
				 				objPCWindow.JavaObject("VariantExpressionEditor").Click iX,iY,"LEFT"
				 				Wait 1
				 			Case "None" 'For "!=" or or "=NONE" condition
				 				objPCWindow.JavaObject("VariantExpressionEditor").dblClick iX,iY,"LEFT"
				 				Wait 1
				 			Case "Blank" 'to set blank
				 			   ' For future use
				 			Case else
				 				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Invalid case [ "&arrOptionFlag(iCnt)&" ]" )
								Fn_PC_VariantNatTable_VariantExpressionEditor_Operations = False
								Set objPCWindow = Nothing
								Exit function
				 		End Select		
					Next					
				 End If
				 
				 Fn_PC_VariantNatTable_VariantExpressionEditor_Operations = True	
				 
				 sMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("RAC_Toolbar"),"Savethecurrentcontents") 'Save the Variant Condition
				 Call Fn_ToolBarOperation("Click",sMenu,"")
				 Call Fn_ReadyStatusSync(1)
		'=========================================================================================================
		 Case "AddAdditionalValuesInAvailabilityMatrix"
		    If Fn_UI_ObjectExist("Fn_PC_VariantNatTable_VariantExpressionEditor_Operations",objPCWindow.JavaObject("VariantExpressionEditor")) = True Then
				 Call Fn_SISW_UI_JavaList_Operations("Fn_PC_VariantNatTable_VariantExpressionEditor_Operations","Select", objPCWindow,"CCombo","Show Features","","")
				 Call Fn_ReadyStatusSync(1)
				 
				 bFlag = Fn_PC_VariantNatTable_VariantExpressionEditor_Operations("SetAvailabilityMatrixOptionAndSave",StrTabName,dicDetails,"",sTabClose)	
				 Fn_PC_VariantNatTable_VariantExpressionEditor_Operations = bFlag
			End If
		'=========================================================================================================
		Case "AddFreeFormValues"
				 If dicDetails("ColumnName") <> "" Then  'Get column index
			 		iColCount = objPCWindow.JavaObject("VariantExpressionEditor").Object.getColumnCount()
			 		For iCnt = 0 To iColCount - 1
						strColumn =  CStr(objPCWindow.JavaObject("VariantExpressionEditor").Object.getCellByPosition(iCnt,0).getlayer.getColumnHeaderlayer.getDatavalueByPosition(iCnt,0).tostring)
						If Instr(trim(strColumn),trim(dicDetails("ColumnName"))) > 0  Then
							iColIndex = iCnt + 1
							Exit for
						End If
					Next
				 Else
					iColIndex = 1 'Default take Option column
				 End If
				'Subject Index
				 iRowsCount = objPCWindow.JavaObject("VariantExpressionEditor").Object.getRowCount()  'Get Index for Subject & Applicability sections
				 For iCnt = 1 To iRowsCount - 1
					If instr(objPCWindow.JavaObject("VariantExpressionEditor").Object.getCellByPosition(0,iCnt).getDataValue().getData().toString(),"Subject") > 0 Then
						iSubJectInx = iCnt
						iColCount = iCnt
						Exit For
					End IF
				 Next
				 'Applicability Index
				 For iCnt = iColCount To iRowsCount - 1
					If instr(objPCWindow.JavaObject("VariantExpressionEditor").Object.getCellByPosition(0,iCnt).getDataValue().getData().toString(),"Applicability") > 0 Then
						iApplicabilityInx = iCnt
						Exit For
					End IF
				 Next 
				 
				 'Create Free Form values in Subject Section
				 If dicDetails("SubjectOption") <> "" AND dicDetails("SubjectFreeFormValue") <> "" Then
				 	For iCnt1 = 1 To iApplicabilityInx
						If instr(CStr(objPCWindow.JavaObject("VariantExpressionEditor").Object.getCellByPosition(0,iCnt1).getDataValue().getData().toString()),dicDetails("SubjectOption")) Then
							iRowIndex = iCnt1
							Exit For
						End If
					Next
					'Split Free Form values to add under option provided
					arrOptions = Split(dicDetails("SubjectFreeFormValue"),"~")
					For iCnt = 0 To UBound(arrOptions)
						strRow = iRowIndex + iCnt + 1
						StrBounds = objPCWindow.JavaObject("VariantExpressionEditor").Object.getBoundsByPosition(iColIndex,strRow).tostring
			 			StrBounds = Split(Replace(Replace(StrBounds,"Rectangle {",""),"}",""),",")
			 			iX = cint(StrBounds(2))
			 			iY = cint(StrBounds(1))
						If IsEmpty(iX) or IsEmpty(iY) Then
				 			Fn_PC_VariantNatTable_VariantExpressionEditor_Operations = False
							Set objPCWindow = Nothing
				 			Exit Function
				 		End If
					 	iX = iX + 5
					 	iY = iY + 4
						objPCWindow.JavaObject("VariantExpressionEditor").Click iX,iY,"LEFT"
				 		Wait 1
						'----------------------- Added by Vinati, to set string of 128 characters long    25-6-21
						bFlag=Fn_SISW_UI_JavaEdit_Operations("Fn_PC_VariantNatTable_VariantExpressionEditor_Operations", "SetExt",  objPCWindow, "FreeFormTextInVEE",arrOptions(iCnt) )
						'bFlag = Fn_Edit_Box("Fn_PC_VariantNatTable_VariantExpressionEditor_Operations",objPCWindow,"FreeFormTextInVEE",arrOptions(iCnt))
						If bFlag = False Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to create Free Form value [ "&arrOptions(iCnt)&" ]" )
							Fn_PC_VariantNatTable_VariantExpressionEditor_Operations = False
							Set objPCWindow = Nothing
							Exit function
						End If
						Call Fn_KeyBoardOperation("SendKeys", "{ENTER}")
						Wait 1
					Next	
				 End IF
						 
				 'Create Free Form values in Subject Section
				 If dicDetails("ApplicabilityOption") <> "" AND dicDetails("ApplicabilityFreeFormValue") <> "" Then
				 	For iCnt1 = iApplicabilityInx To iRowsCount
						If instr(CStr(objPCWindow.JavaObject("VariantExpressionEditor").Object.getCellByPosition(0,iCnt1).getDataValue().getData().toString()),dicDetails("ApplicabilityOption")) Then
							iRowIndex = iCnt1
							Exit For
						End If
					Next
				 	'Split Free Form values to add under option provided
					arrOptions = Split(dicDetails("ApplicabilityFreeFormValue"),"~")
					For iCnt = 0 To UBound(arrOptions)
						strRow = iRowIndex + iCnt + 1
						StrBounds = objPCWindow.JavaObject("VariantExpressionEditor").Object.getBoundsByPosition(iColIndex,strRow).tostring
			 			StrBounds = Split(Replace(Replace(StrBounds,"Rectangle {",""),"}",""),",")
			 			iX = cint(StrBounds(2))
			 			iY = cint(StrBounds(1))
						If IsEmpty(iX) or IsEmpty(iY) Then
				 			Fn_PC_VariantNatTable_VariantExpressionEditor_Operations = False
							Set objPCWindow = Nothing
				 			Exit Function
				 		End If
					 	iX = iX + 5
					 	iY = iY + 4
						objPCWindow.JavaObject("VariantExpressionEditor").Click iX,iY,"LEFT"
				 		Wait 1
						bFlag = Fn_Edit_Box("Fn_PC_VariantNatTable_VariantExpressionEditor_Operations",objPCWindow,"FreeFormTextInVEE",arrOptions(iCnt))
						If bFlag = False Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to create Free Form value [ "&arrOptions(iCnt)&" ]" )
							Fn_PC_VariantNatTable_VariantExpressionEditor_Operations = False
							Set objPCWindow = Nothing
							Exit function
						End If
						Call Fn_KeyBoardOperation("SendKeys", "{ENTER}")
						Wait 1
					Next						
				 End IF
				 Fn_PC_VariantNatTable_VariantExpressionEditor_Operations = True	
		'=========================================================================================================
		Case "SelectColumnsInVariantExpressionEditor","VerifyErrorMessage"
			 If dicDetails("ColumnName") <> "" Then  'Get column index
					arrOptions = Split(dicDetails("ColumnName"),"~")
					For iCnt1 = 0 To UBound(arrOptions)
						iColCount = objPCWindow.JavaObject("VariantExpressionEditor").Object.getColumnCount()
						For iCnt = 0 To iColCount - 1
							iColPosition = objPCWindow.JavaObject("VariantExpressionEditor").Object.getstartXOfcolumnPosition(iCnt+3)
							strColumn =  CStr(objPCWindow.JavaObject("VariantExpressionEditor").Object.getCellByPosition(iCnt,0).getlayer.getColumnHeaderlayer.getDatavalueByPosition(iCnt,0).tostring)
							If instr(trim(strColumn),trim(arrOptions(iCnt1))) > 0 Then
								iColIndex = iCnt + 1
								Exit for
							End If
						Next
						StrBounds = objPCWindow.JavaObject("VariantExpressionEditor").Object.getBoundsByPosition(iColIndex,0).tostring
						StrBounds = Split(Replace(Replace(StrBounds,"Rectangle {",""),"}",""),",")
						iX = cint(StrBounds(2))
						iY = cint(StrBounds(3) - 40)
						If IsEmpty(iX) or IsEmpty(iY) Then
							Fn_PC_VariantNatTable_VariantExpressionEditor_Operations = False
							Set objPCWindow = Nothing
							Exit Function
						End If
						If iX = 0 Then
							iX = iColPosition
						End If
						iX = iX + 4
						iY = iY + 5
						If iCnt1 = 0 Then
							objPCWindow.JavaObject("VariantExpressionEditor").Click iX, iY ,"LEFT"
						Else
							iX = iColPosition
							Const LK_Ctrl = 29
							Set myDeviceReplay = CreateObject("Mercury.DeviceReplay")
							myDeviceReplay.KeyDown 29
							objPCWindow.JavaObject("VariantExpressionEditor").Click iX, iY ,"LEFT"
							myDeviceReplay.KeyUp 29
							Set myDeviceReplay = Nothing	
						End If
						Wait 1
					Next
					Fn_PC_VariantNatTable_VariantExpressionEditor_Operations = True
					'------------------- Added for Case Message verification ------------------
					If sAction = "VerifyErrorMessage" Then
						Fn_PC_VariantNatTable_VariantExpressionEditor_Operations = False
						If dicDetails("Message") <> "" Then
							iX = iX + 20
							iY = iY + 20
							objPCWindow.JavaObject("VariantExpressionEditor").Click iX, iY ,"LEFT"						
							Wait 1
							Set myDeviceReplay = CreateObject("Mercury.Clipboard")
							myDeviceReplay.Clear
							Call Fn_KeyBoardOperation("SendKeys", "^a") 'Select All
							Wait 1
							Call Fn_KeyBoardOperation("SendKeys", "^c") 'Copy text
							Wait 1
							sAppMsg = myDeviceReplay.GetText
							Set myDeviceReplay = Nothing
						   If instr(trim(sAppMsg),trim(dicDetails("Message")),1) > 0 Then
						   	   Fn_PC_VariantNatTable_VariantExpressionEditor_Operations = sAppMsg
						   Else
						   	   Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to verify error message [ "&dicDetails("Message")&" ]" )	
						   	   Fn_PC_VariantNatTable_VariantExpressionEditor_Operations = False
							   Set objPCWindow = Nothing
							   Exit Function
						   End If								
						End If
					End If
			End If	
		'=========================================================================================================
		Case "VerifyColumnInVariantExpressionEditor"
			 If dicDetails("ColumnName") <> "" Then  'Get column index
				arrOptions = Split(dicDetails("ColumnName"),"~")
				iColCount = objPCWindow.JavaObject("VariantExpressionEditor").Object.getColumnCount()
				For iCnt1 = 0 To UBound(arrOptions)
					bFlag = False
					For iCnt = 0 To iColCount - 1
						strColumn =  CStr(objPCWindow.JavaObject("VariantExpressionEditor").Object.getCellByPosition(iCnt,0).getlayer.getColumnHeaderlayer.getDatavalueByPosition(iCnt,0).tostring)
						If instr(trim(strColumn),trim(arrOptions(iCnt1))) > 0 Then
							bFlag = True	
							Exit for
						End If
					Next
					If bFlag = False Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed verify existence of column [ "&arrOptions(iCnt1)&" ]" )
						Fn_PC_VariantNatTable_VariantExpressionEditor_Operations = False
						Set objPCWindow = Nothing
						Exit function
					End If
				Next
				Fn_PC_VariantNatTable_VariantExpressionEditor_Operations = True				
			 End If
		'===========================================================================================
		Case "VerifyAvailabilityMatrixOption"	 
				 If dicDetails("ColumnName") <> "" Then  'Get column index
			 		iColCount = objPCWindow.JavaObject("VariantExpressionEditor").Object.getColumnCount()
			 		For iCnt = 0 To iColCount - 1
						strColumn =  CStr(objPCWindow.JavaObject("VariantExpressionEditor").Object.getCellByPosition(iCnt,0).getlayer.getColumnHeaderlayer.getDatavalueByPosition(iCnt,0).tostring)
						If trim(dicDetails("ColumnName")) = trim(strColumn) Then
								iColIndex = iCnt + 1
								Exit for
						End If
					Next
				 Else
					iColIndex = 2							 
				 End If
					
				 iRowsCount = objPCWindow.JavaObject("VariantExpressionEditor").Object.getRowCount()  'Get Index for Subject & Applicability sections
					If dicDetails("AvailabilityOption") <> "" Then
					arrOptions = Split(dicDetails("AvailabilityOption"),"~")
					For iCnt = 0 To UBound(arrOptions)
						bFlag = False
						For iCnt1 = 1 To iRowsCount-1
							If instr(CStr(objPCWindow.JavaObject("VariantExpressionEditor").Object.getCellByPosition(0,iCnt1).getDataValue().getData().toString()),arrOptions(iCnt)) Then
								iRowIndex = iCnt1
								bFlag = True
								Exit For
							End If
						Next
						
				 		If bFlag = False Then
				 			Fn_PC_VariantNatTable_VariantExpressionEditor_Operations = False
				 			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Fail to verify Variant condition [ "&arrOptions(iCnt)&" = "&arrOptionFlag(iCnt)&" ]" )
				 			Set objPCWindow = Nothing
							Exit function
				 			Exit For
				 		Else
				 			Fn_PC_VariantNatTable_VariantExpressionEditor_Operations = True
				 		End If			
					Next
				End IF 
		'===========================================================================================
		Case "PopupMenuSelectOnColumn"
			If dicDetails("ColumnName") <> "" Then  'Get column index
				iColPosition = 0
		 		iColCount = objPCWindow.JavaObject("VariantExpressionEditor").Object.getColumnCount()
		 		For iCnt = 0 To iColCount - 1
					iColPosition = objPCWindow.JavaObject("VariantExpressionEditor").Object.getstartXOfcolumnPosition(iCnt+3)
					strColumn =  CStr(objPCWindow.JavaObject("VariantExpressionEditor").Object.getCellByPosition(iCnt,0).getlayer.getColumnHeaderlayer.getDatavalueByPosition(iCnt,0).tostring)
					If trim(dicDetails("ColumnName")) = trim(strColumn) Then
							iColIndex = iCnt + 1
							Exit for
					End If
				Next
			Else
				iColIndex = 2							 
		    End If
		    StrBounds = objPCWindow.JavaObject("VariantExpressionEditor").Object.getBoundsByPosition(iColIndex,0).tostring
			StrBounds = Split(Replace(Replace(StrBounds,"Rectangle {",""),"}",""),",")
			iX = cint(StrBounds(2))
			iY = cint(StrBounds(3) - 40)
			If IsEmpty(iX) or IsEmpty(iY) Then
				Fn_PC_VariantNatTable_VariantExpressionEditor_Operations = False
				Set objPCWindow = Nothing
				Exit Function
			End If				
			If iColIndex > 2 Then
				iX = iColPosition
			End If
			iX = iX + 4
			iY = iY + 5
			objPCWindow.JavaObject("VariantExpressionEditor").Click iX, iY ,"RIGHT"
			Wait 2
			strPopupMenu = objPCWindow.WinMenu("ContextMenu").BuildMenuPath(Popupmenu)
			objPCWindow.WinMenu("ContextMenu").Select strPopupMenu
			Wait 3
			Fn_PC_VariantNatTable_VariantExpressionEditor_Operations = True	
	'=============================================================================================================
		Case "ClickOnRow"
			 If dicDetails("ColumnName") <> "" Then  'Get column index
		 		iColCount = objPCWindow.JavaObject("VariantExpressionEditor").Object.getColumnCount()
		 		For iCnt = 0 To iColCount - 1
					strColumn =  CStr(objPCWindow.JavaObject("VariantExpressionEditor").Object.getCellByPosition(iCnt,0).getlayer.getColumnHeaderlayer.getDatavalueByPosition(iCnt,0).tostring)
					If trim(dicDetails("ColumnName")) = trim(strColumn) Then
						iColIndex = iCnt + 1
						Exit for
					End If
				Next
			 Else
				iColIndex = 1 'Default take Option column
			 End If
			 
			 'get row count & Index of it
			 iRowsCount = objPCWindow.JavaObject("VariantExpressionEditor").Object.getRowCount()
			 If dicDetails("RowValue") <> "" Then
			 	For iCnt1 = 1 To iRowsCount
					If instr(CStr(objPCWindow.JavaObject("VariantExpressionEditor").Object.getCellByPosition(0,iCnt1).getDataValue().getData().toString()),dicDetails("RowValue")) Then
						iRowIndex = iCnt1
						Exit For
					End If
				Next
				StrBounds = objPCWindow.JavaObject("VariantExpressionEditor").Object.getBoundsByPosition(iColIndex,iRowIndex).tostring
				StrBounds = Split(Replace(Replace(StrBounds,"Rectangle {",""),"}",""),",")
				iX = cint(StrBounds(2))
				iY = cint(StrBounds(1))
				If IsEmpty(iX) or IsEmpty(iY) Then
					Fn_PC_VariantNatTable_VariantExpressionEditor_Operations = False
					Set objPCWindow = Nothing
					Exit Function
				End If
				iX = iX + 5
				iY = iY + 4
				objPCWindow.JavaObject("VariantExpressionEditor").Click iX,iY,"LEFT"
				Wait 1		
			 End IF
			 Fn_PC_VariantNatTable_VariantExpressionEditor_Operations = True
			'=============================================================================================================
		Case "PopupMenuSelectOnColumn_Ext"		'This case is used to select popup menu on columns in VEE if columns are with # no of splits and need to select popup menu columns 	
				If dicDetails("ColumnIndex") <> "" AND dicDetails("ColumnIndex")< 4 Then
					colindex=cint(dicDetails("ColumnIndex"))+1
					
				ElseIf dicDetails("ColumnIndex") <> "" AND dicDetails("ColumnIndex")>=4 Then
					colindex=dicDetails("ColumnIndex")
				Else
					colindex=3
				End If
				
				iColCount = objPCWindow.JavaObject("VariantExpressionEditor").Object.getColumnCount()
		 		For iCnt = 0 To iColCount 
			 			iColPosition = objPCWindow.JavaObject("VariantExpressionEditor").Object.getstartXOfcolumnPosition(iCnt+colindex)
						strColumn =  CStr(objPCWindow.JavaObject("VariantExpressionEditor").Object.getCellByPosition(iCnt,0).getlayer.getColumnHeaderlayer.getDatavalueByPosition(iCnt,0).tostring)
						If trim(dicDetails("ColumnName")) = trim(strColumn) Then
								iColIndex = iCnt + 1
								Exit for
						End If
					Next		
		
				 StrBounds = objPCWindow.JavaObject("VariantExpressionEditor").Object.getBoundsByPosition(iColIndex,0).tostring
				 StrBounds = Split(Replace(Replace(StrBounds,"Rectangle {",""),"}",""),",")
				 If iX = 0 Then
				 	  iX = iColPosition
				 	  iY = iY + 5
				 else
					 iX = iX + 4
					 iY = iY + 5
				 End If 	
				 objPCWindow.JavaObject("VariantExpressionEditor").Click iX, iY ,"RIGHT"
				 Wait 2
				 strPopupMenu = objPCWindow.WinMenu("ContextMenu").BuildMenuPath(Popupmenu)
				 objPCWindow.WinMenu("ContextMenu").Select strPopupMenu
				wait 2
				Fn_PC_VariantNatTable_VariantExpressionEditor_Operations = True	
  '===============================================================================================				
	Case "VerifyErrorMessage_Ext" 
						Fn_PC_VariantNatTable_VariantExpressionEditor_Operations = False
						If dicDetails("Message") <> "" Then
							iX = iX + 230
							iY = iY + 85
							objPCWindow.JavaObject("VariantExpressionEditor").Click iX, iY ,"LEFT"						
							Wait 1
							Set myDeviceReplay = CreateObject("Mercury.Clipboard")
							myDeviceReplay.Clear
							Call Fn_KeyBoardOperation("SendKeys", "^a") 'Select All
							Wait 1
							Call Fn_KeyBoardOperation("SendKeys", "^c") 'Copy text
							Wait 1
							sAppMsg = myDeviceReplay.GetText
							Set myDeviceReplay = Nothing
						   If instr(trim(sAppMsg),trim(dicDetails("Message")),1) > 0 Then
						   	   Fn_PC_VariantNatTable_VariantExpressionEditor_Operations = sAppMsg
						   Else
						   	   Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to verify error message [ "&dicDetails("Message")&" ]" )	
						   	   Fn_PC_VariantNatTable_VariantExpressionEditor_Operations = False
							   Set objPCWindow = Nothing
							   Exit Function
						   End If
						End If 
						
	End Select
	If Fn_ToolBarOperation("IsSelected","Navigation Pane","") = False Then
		Call Fn_ToolBarOperation("Click","Navigation Pane","")
	End If
	 If StrTabName <> "" Then ' Minimize Tab
		Call Fn_TabFolder_Operation("DoubleClickTab", StrTabName,"")
		Call Fn_ReadyStatusSync(1)
		If sTabClose = "Yes" Then ' Close Tab
			Call Fn_TabFolder_Operation("Close", StrTabName,"")
			Call Fn_ReadyStatusSync(1)
		End If	
	 End If
	 
	Set objPCWindow = Nothing
	
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_PC_VariantConfigurationView_Operation
'@@
'@@    Description			:	Function Used to Perform operation on Variant Configuration View
'@@
'@@    Parameters			:	1. sAction		: Action to be performed
'@@							:	2. StrTabName	: Tab Name
'@@							:	2. dicDetails	: Dictionary object
'@@							:	2. Popupmenu	: Popup menu name
'@@							:	3. sTabClose	: Flag to close Tab
'@@
'@@    Return Value		   	: 	True Or False
'@@
'@@    Pre-requisite		:	Product Configurator Perspective should be opened 
'@@
'@@    Examples				:	Set dicDetails = CreateObject( "Scripting.Dictionary")
'@@ 								dicDetails("Options") = "1200~1600~Arial~std"
'@@ 								dicDetails("OptionFlag") = "Check~Check~None~Check"
'@@									dicDetails("ToolBarButton") = "Expand"
'@@ 							bReturn = Fn_PC_VariantConfigurationView_Operation("SetVarOptionValue","Variant Configuration",dicDetails,"","Yes")
'@@ 								
'@@ 							Set dicDetails = CreateObject( "Scripting.Dictionary")
'@@ 								dicDetails("Options") = "1200~1600~Arial~std"
'@@									dicDetails("ToolBarButton") = "Expand"
'@@ 							bReturn = Fn_PC_VariantConfigurationView_Operation("ModifySetVarOptionValue","Variant Configuration",dicDetails,"","Yes")
'@@ 								
'@@ 							Set dicDetails = CreateObject( "Scripting.Dictionary")
'@@ 								dicDetails("Options") = "1200~1600~Arial~std"
'@@ 								dicDetails("OptionFlag") = "None~Check~None~Check"
'@@ 							bReturn = Fn_PC_VariantConfigurationView_Operation("VerifySetVarOptionValue","Variant Configuration",dicDetails,"","Yes")
'@@  
'@@-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@	   History				:	Developer Name				Date	  			Rev. No.		Changes Done		Reviewer
'@@-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@	   Created By 			:	Poonam Chopade			12-Aug-2019				1.0		  		 Created		  [TC12.3(2019070800)-12Aug2019-PoonamC-NewDevelopment]
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_PC_VariantConfigurationView_Operation(sAction,StrTabName,dicDetails,Popupmenu,sTabClose)

	GBL_FAILED_FUNCTION_NAME = "Fn_PC_VariantConfigurationView_Operation"
	Dim objPCWindow,arrOptionFlag,arrOptions
	Dim sMenu,bFlag,iCnt,iCnt1,iRowIndex,iRowsCount,iColIndex,iColCount,iCnt2,sOption,strFlag
	Dim sAppMsg,iX,iY,iHeight,iItemHeight,aOption,iInstance,iCounter
	
	Fn_PC_VariantConfigurationView_Operation = False
	
	Set objPCWindow = Fn_PC_GetObject("ProductConfigurator")
	On Error Resume Next
	
	'Check Existence of Variant Configuration Tab & click on Toolbar button
	If Fn_UI_ObjectExist("Fn_PC_VariantConfigurationView_Operation",objPCWindow.JavaObject("DateTimeImageHyperlink")) = False And Fn_UI_ObjectExist("Fn_PC_VariantConfigurationView_Operation",objPCWindow.JavaTable("VariantConfigTable")) = False Then
		sMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("Configurator_Toolbar"),"OpenVariantConfigurationview")
		Call Fn_ToolBarOperation("Click",sMenu,"")
		Call Fn_ReadyStatusSync(1)
	End If
	
	Select Case sAction
		Case "SetVarOptionValue","ModifySetVarOptionValue","SetVarOptionValue_Ext","SetVarOptionValue1","SetVarOptionValueWithoutRuleDate"
			   	
				If Fn_UI_ObjectExist("Fn_PC_VariantConfigurationView_Operation",objPCWindow.JavaTable("VariantConfigTable")) = True Then
						 ' Maximize tab
						 If StrTabName <> "" Then
						 		Call Fn_TabFolder_Operation("DoubleClickTab", StrTabName,"")
								Call Fn_ReadyStatusSync(1)
						 End If 
						 If sAction= "SetVarOptionValue" or sAction = "ModifySetVarOptionValue" Then
		       				Call Fn_PC_SetRuleDate("SetRuleDate","","","No Date") 'Clear date rule
			   				Call Fn_ReadyStatusSync(1)
		      	 		 End If
						 If sAction = "ModifySetVarOptionValue" Then ' clear set expression and then set new expression
						 		Call Fn_ToolBarOperation("ShowDropdownAndSelectWithInstance", "View Menu:1", "Clear Expression")
								Call Fn_ReadyStatusSync(1)
						 End If
						 
						 ' Added code to select Node ot group name in guided configuration Mode
						 If dicDetails("SelectNodeInActiveMode") <> "" Then
						 	Call Fn_UI_JavaStaticText_SetTOProperty("Fn_PC_VariantConfigurationView_Operation",objPCWindow,"StaticText","label",dicDetails("SelectNodeInActiveMode")) 
						 	Call Fn_UI_JavaStaticText_Click("Fn_PC_VariantConfigurationView_Operation",objPCWindow,"StaticText", 5, 5, "LEFT")
						 	Call Fn_ReadyStatusSync(1)
						 End If

						 arrOptions = Split(dicDetails("Options"),"~") ' Set Options as On or OFF
						 arrOptionFlag = Split(dicDetails("OptionFlag"),"~")
						 iRowsCount = objPCWindow.JavaTable("VariantConfigTable").GetROProperty("rows")
					     iColCount = objPCWindow.JavaTable("VariantConfigTable").GetROProperty("cols")			 
						 For iCnt = 0 To UBound(arrOptions)
						 		'--------------------- for multiple instance ----------------------------
						 		If instr(arrOptions(iCnt),"@") Then
						 			aOption = Split(arrOptions(iCnt),"@")
						 			arrOptions(iCnt) = aOption(0)
						 			iInstance = aOption(1)
						 		Else
									iInstance = 1				 		
						 		End If
						 		'--------------------- ---------------------- ----------------------------
							 	iCounter = 1
							 	For iCnt1 = 0 To iRowsCount - 1 'Get Row & col index as per option name	
							 		bFlag = False
							 		For iCnt2 = 0 To iColCount - 1
							 			sOption = objPCWindow.JavaTable("VariantConfigTable").GetCellData(iCnt1,iCnt2)
							 			If trim(arrOptions(iCnt)) = trim(sOption) Then
							 				'--------------------- ---------------------- ----------------------------
							 				If cint(iCounter) = cint(iInstance) Then
							 					bFlag = True
							 					Exit for
								 			Else
												iCounter = iCounter + 1							 			
								 			End If
							 				'--------------------- ---------------------- ----------------------------
							 			End If
							 		Next
							 		If bFlag = True Then
							 			iRowIndex = iCnt1
							 			If dicDetails("GridNumber") <> "" Then
							 				iColIndex = dicDetails("GridNumber")
							 			Else
							 				iColIndex = iCnt2-1
							 			End If
							 			Exit for
							 		End If	
							 	Next
							 	If sAction = "SetVarOptionValue_Ext" Then  'For Boolean Value to check  for same value
							 	    	iRowIndex=iRowIndex+1
							 	    	iColIndex=iColIndex+3
							 	       Call Fn_ReadyStatusSync(2)
							 	    End If
								Select Case arrOptionFlag(iCnt)
						 			Case "Check" 'For "=" or "=Any" condition
						 				objPCWindow.JavaTable("VariantConfigTable").SelectCell iRowIndex,iColIndex
						 				Wait 1
						 			Case "None" 'For "!=" or or "=NONE" condition
						 				objPCWindow.JavaTable("VariantConfigTable").SelectCell iRowIndex,iColIndex
						 				Wait 1
						 				objPCWindow.JavaTable("VariantConfigTable").SelectCell iRowIndex,iColIndex
						 				Wait 1
						 			Case Else
						 				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Fail to invalid case [ "&arrOptionFlag(iCnt)&" ]" )
										Fn_PC_VariantConfigurationView_Operation = False
										Set objPCWindow = Nothing
										Exit function
						 		End Select	
								Fn_PC_VariantConfigurationView_Operation = True	
						 Next
						 
						 If dicDetails("ToolBarButton") <> "" Then    ' Click toolbar button in Configuration view
						 		sOption = Split(dicDetails("ToolBarButton"),"~")
						 		For iCnt1 = 0 To UBound(sOption)
						 			Call Fn_ToolBarOperation("Click",sOption(iCnt1),"")
						 			Call Fn_ReadyStatusSync(1)
						 		Next
						 End If
						 
						 If StrTabName <> "" Then ' Minimize Tab
						 		Call Fn_TabFolder_Operation("DoubleClickTab", StrTabName,"")
								Call Fn_ReadyStatusSync(1)
								If sTabClose = "Yes" Then ' Close Tab
									Call Fn_TabFolder_Operation("Close", StrTabName,"")
									Call Fn_ReadyStatusSync(1)
								End If	
						 End If
				End If
		
		Case "VerifySetVarOptionValue","VerifySetVarOptionValueExt"
				If sAction <> "VerifySetVarOptionValueExt" Then
		        	Call Fn_PC_SetRuleDate("SetRuleDate","","","No Date") 'Clear date rule
			    Call Fn_ReadyStatusSync(1)
		        End If
		   
				If Fn_UI_ObjectExist("Fn_PC_VariantConfigurationView_Operation",objPCWindow.JavaTable("VariantConfigTable")) = True Then
						 If StrTabName <> "" Then  ' Maximize tab
						 		Call Fn_TabFolder_Operation("DoubleClickTab", StrTabName,"")
								Call Fn_ReadyStatusSync(1)
						 End If

						' Added code to select Node ot group name in guided configuration Mode
						 If dicDetails("SelectNodeInActiveMode") <> "" Then
						 	Call Fn_UI_JavaStaticText_SetTOProperty("Fn_PC_VariantConfigurationView_Operation",objPCWindow,"StaticText","label",dicDetails("SelectNodeInActiveMode")) 
						 	Call Fn_UI_JavaStaticText_Click("Fn_PC_VariantConfigurationView_Operation",objPCWindow,"StaticText", 5, 5, "LEFT")
						 	Call Fn_ReadyStatusSync(1)
						 End If
						
						 arrOptions = Split(dicDetails("Options"),"~") ' Set Options as On or OFF
						 arrOptionFlag = Split(dicDetails("OptionFlag"),"~")
						 iRowsCount = objPCWindow.JavaTable("VariantConfigTable").GetROProperty("rows")
					     iColCount = objPCWindow.JavaTable("VariantConfigTable").GetROProperty("cols")				 
						 For iCnt = 0 To UBound(arrOptions)  'Get Row & col index as per option name	
						 		'--------------------- for multiple instance ----------------------------
						 		If instr(arrOptions(iCnt),"@") Then
						 			aOption = Split(arrOptions(iCnt),"@")
						 			arrOptions(iCnt) = aOption(0)
						 			iInstance = aOption(1)
						 		Else
									iInstance = 1				 		
						 		End If
						 		'------------------------------------------------------------------------
						 		iCounter = 1
							 	For iCnt1 = 0 To iRowsCount - 1
							 		bFlag = False
							 		For iCnt2 = 0 To iColCount - 1
							 			sOption = objPCWindow.JavaTable("VariantConfigTable").GetCellData(iCnt1,iCnt2)
							 			If trim(arrOptions(iCnt)) = trim(sOption) Then
							 				'--------------------- ---------------------- ----------------------------
							 				If cint(iCounter) = cint(iInstance) Then
							 					bFlag = True
							 					Exit for
								 			Else
												iCounter = iCounter + 1							 			
								 			End If
							 				'--------------------- ---------------------- ----------------------------
							 			End If
							 		Next
							 		If bFlag = True Then
							 			iRowIndex = iCnt1
							 			iColIndex = iCnt2
							 			Exit for
							 		End If	
							 	Next
								' if multiple grid coulmn present to check condition ex. in case of Overlay option							 	
						 		If dicDetails("GridNumber") <> "" Then
						 			iColIndex = iColIndex - cint(dicDetails("GridNumber"))
						 		Else
									iColIndex = iColIndex - 1						 		
						 		End If
						 		Wait 1
						 		bFlag = False
								'Added below code to get Unique value of Image i.e Checked & Unchecked image data value
						 		strFlag = objPCWindow.JavaTable("VariantConfigTable").Object.getitem(iRowIndex).getimage(iColIndex).getimageData.getAlpha(4,5) 
						 		Select Case arrOptionFlag(iCnt)
						 			Case "Check" 'For checked condition
							 			If strFlag <> Empty Then	
							 				If CLng(strFlag) = 0  OR CLng(strFlag) = 63 Then
							 					bFlag = True
							 				End If
							 			End If	
						 			Case "None" 'For None condition
						 				 If strFlag <> Empty Then
						 					If CLng(strFlag) > 0 Then
							 					bFlag = True
							 				End If
							 			End If		
						 			Case "Blank" 'For blank condition
						 			   	    If strFlag = Empty Then
							 					bFlag = True
							 				End If
						 		End Select
						 		If bFlag = False Then
						 			Fn_PC_VariantConfigurationView_Operation = False
						 			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Fail to verify Variant condition [ "&arrOptions(iCnt)&" = "&arrOptionFlag(iCnt)&" ]" )
						 			Exit For
						 		Else
						 			Fn_PC_VariantConfigurationView_Operation = True
						 		End If
								strFlag = "" ' clear value from variable						 		
						 Next
						 
						 If StrTabName <> "" Then ' Minimize Tab
						 		Call Fn_TabFolder_Operation("DoubleClickTab", StrTabName,"")
								Call Fn_ReadyStatusSync(1)
								If sTabClose = "Yes" Then ' Close Tab
									Call Fn_TabFolder_Operation("Close", StrTabName,"")
									Call Fn_ReadyStatusSync(1)
								End If	
						 End If
				End If
		'==========================================================================================================
		Case "VerifyEnabledConfigHeaderLabel"
			 For iCnt = 0 To 2
			 	objPCWindow.JavaStaticText("ConfigHeaderLabel").SetTOProperty "index",iCnt
			 	If cstr(objPCWindow.JavaStaticText("ConfigHeaderLabel").Object.isEnabled()) = "true" Then
			 		If objPCWindow.JavaStaticText("ConfigHeaderLabel").Object.getToolTipText() = dicDetails("ConfigHeadertooltip") Then
			 			Fn_PC_VariantConfigurationView_Operation = True
			 		End If
			 	End If	
			 Next
		'==========================================================================================================
		Case "EditFreeFormValue"
			iRowsCount = objPCWindow.JavaTable("VariantConfigTable").GetROProperty("rows")
			iColCount = objPCWindow.JavaTable("VariantConfigTable").GetROProperty("cols")
			For iCnt1 = 0 To iRowsCount - 1
				bFlag = False
				For iCnt2 = 0 To iColCount - 1
					sOption = objPCWindow.JavaTable("VariantConfigTable").GetCellData(iCnt1,iCnt2)
					If trim(dicDetails("Options")) = trim(sOption) Then
						bFlag = True
						Exit for
					End If
				Next
				If bFlag = True Then
					iRowIndex = iCnt1
					Exit for
				End If	
			Next
			iRowIndex = iRowIndex + 1
			iColIndex = iColCount - 2
			objPCWindow.JavaTable("VariantConfigTable").SelectCell iRowIndex,iColIndex
			Wait 2
			bFlag = Fn_SISW_UI_JavaEdit_Operations("Fn_PC_VariantConfigurationView_Operation","Set",objPCWindow,"FreeFormText", dicDetails("FreeFormValue"))
			
			If dicDetails("ToolBarButton") <> "" Then    ' Click toolbar button in Configuration view
		 		sOption = Split(dicDetails("ToolBarButton"),"~")
		 		For iCnt1 = 0 To UBound(sOption)
		 			Call Fn_ToolBarOperation("Click",sOption(iCnt1),"")
		 			Call Fn_ReadyStatusSync(1)
		 		Next
			End If
			
			Fn_PC_VariantConfigurationView_Operation = bFlag	
	 '==========================================================================================================
	  Case "VerifyErrorMessage"	,"VerifyErrorMessage_Ext"
			iRowsCount = Fn_UI_Object_GetROProperty("Fn_PC_VariantConfigurationView_Operation",objPCWindow.JavaTable("VariantConfigTable"),"rows")
 			iColCount = Fn_UI_Object_GetROProperty("Fn_PC_VariantConfigurationView_Operation",objPCWindow.JavaTable("VariantConfigTable"),"cols")	
			For iCnt1 = 0 To iRowsCount - 1   'Get Row & col index as per option name	
				bFlag = False
				For iCnt2 = 0 To iColCount - 1
					sOption = objPCWindow.JavaTable("VariantConfigTable").GetCellData(iCnt1,iCnt2)
					If trim(dicDetails("Options")) = trim(sOption) Then
						bFlag = True
						Exit for
					End If
				Next
			  If bFlag = True Then
				iRowIndex = iCnt1
				iColIndex = iCnt2-2
				Exit for
			  End If	
			Next
			iHeight = objPCWindow.GetROProperty("height")
		 	iItemHeight = objPCWindow.JavaTable("VariantConfigTable").Object.getItemHeight()
			If iRowIndex <> 0  Then
			 	iItemHeight = iItemHeight * iRowIndex
			End If
			iX = cint(cint(iHeight)+38)
			iY = cint(320+cint(iItemHeight)) 
		 	objPCWindow.JavaTable("VariantConfigTable").ActivateCell iRowIndex,iColIndex
			Wait 2
			objPCWindow.Click iX,iY
			Wait 1
			'if Shell dialog not identified then
			'--------------------------------------------------
			If Fn_UI_ObjectExist("",objPCWindow.JavaWindow("Shell").JavaEdit("StyledText"))= False Then
				iColIndex = iColIndex - 1
				objPCWindow.JavaTable("VariantConfigTable").ActivateCell iRowIndex,iColIndex
				Wait 3
				objPCWindow.Click iX,iY
				Wait 1
			End IF
			
			'---------------------------- Added if Shell window is not identified -------------------------
			If Fn_UI_ObjectExist("",objPCWindow.JavaWindow("Shell").JavaEdit("StyledText")) = False Then
				For iCounter = 1 To 4
					objPCWindow.JavaWindow("Shell").SetTOProperty "index",iCounter
					If Fn_UI_ObjectExist("",objPCWindow.JavaWindow("Shell").JavaEdit("StyledText")) = True Then
						Exit For
					End IF 
				Next
			End IF 
			'---------------------------------------------------
			If Fn_UI_ObjectExist("",objPCWindow.JavaWindow("Shell")) Then
	 		 	sAppMsg = Fn_UI_Object_GetROProperty("Fn_PC_VariantConfigurationView_Operation",objPCWindow.JavaWindow("Shell").JavaEdit("StyledText"),"text")
	 		 	If sAppMsg = False AND sAction = "VerifyErrorMessage_Ext" Then
		 			Wait 1
					Set myDeviceReplay = CreateObject("Mercury.Clipboard")
					myDeviceReplay.Clear
					Call Fn_KeyBoardOperation("SendKeys", "^{HOME}") 
					wait 1
					Call Fn_KeyBoardOperation("SendKeys", "^+{END}") 
					Wait 1
					Call Fn_KeyBoardOperation("SendKeys", "^c") 'Copy text
					Wait 1
					sAppMsg = myDeviceReplay.GetText
					Set myDeviceReplay = Nothing
	 		 	End If
		 	 	If instr(sAppMsg,dicDetails("Message")) > 0 Then
		 			Fn_PC_VariantConfigurationView_Operation = True
		 	 	Else
		 			Fn_PC_VariantConfigurationView_Operation = False
		 			Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Actual Message ["+sAppMsg+"] does not match with expected message ["+dicDetails("Message")+"]")		
		 	 	End IF
			 Else
			 	Fn_PC_VariantConfigurationView_Operation = False
			 	Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Message window does not exists")		
			 End If
			 objPCWindow.JavaWindow("Shell").SetTOProperty "index",0
			 Call Fn_KeyBoardOperation("SendKeys", "{ESC}")
			 Call Fn_ReadyStatusSync(1)
		'==========================================================================================================	
		''[Tc13.2_20210401.00_NewDevelopment_AmrutaP : Added New Case to Verify options in config view]
		Case "VerifyOptionsInVarConfigView"
									
				If Fn_UI_ObjectExist("Fn_PC_VariantConfigurationView_Operation",objPCWindow.JavaTable("VariantConfigTable")) = True Then
						 ' Maximize tab
						 If StrTabName <> "" Then
						 		If Fn_TabFolder_Operation("DoubleClickTab", StrTabName,"") = False then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Fail to double click on Tab [ "&StrTabName&" ]" )
									Fn_PC_VariantConfigurationView_Operation = False
									Set objPCWindow = Nothing
									Exit function
								End if
						 End If 
						 ' Added code to select Node ot group name in guided configuration Mode
						 If dicDetails("SelectNodeInActiveMode") <> "" Then
						 	Call Fn_UI_JavaStaticText_SetTOProperty("Fn_PC_VariantConfigurationView_Operation",objPCWindow,"StaticText","label",dicDetails("SelectNodeInActiveMode")) 
						 	Call Fn_UI_JavaStaticText_Click("Fn_PC_VariantConfigurationView_Operation",objPCWindow,"StaticText", 5, 5, "LEFT")
						 	Call Fn_ReadyStatusSync(1)
						 End If
						
						' Set Options as On or OFF
						 arrOptions = Split(dicDetails("Options"),"~")
						 iRowsCount = objPCWindow.JavaTable("VariantConfigTable").GetROProperty("rows")
					     iColCount = objPCWindow.JavaTable("VariantConfigTable").GetROProperty("cols")
										 
						 For iCnt = 0 To UBound(arrOptions)
							    'Get Row & col index as per option name	
							    '--------------------- for multiple instance ----------------------------
						 		If instr(arrOptions(iCnt),"@") Then
						 			aOption = Split(arrOptions(iCnt),"@")
						 			arrOptions(iCnt) = aOption(0)
						 			iInstance = aOption(1)
						 		Else
									iInstance = 1				 		
						 		End If
						 		'------------------------------------------------------------------------
						 		iCounter = 1
							 	For iCnt1 = 0 To iRowsCount - 1
								 		bFlag = False
								 		For iCnt2 = 0 To iColCount - 1
											sOption = objPCWindow.JavaTable("VariantConfigTable").GetCellData(iCnt1,iCnt2)
								 			If trim(arrOptions(iCnt)) = trim(sOption) Then
								 				'--------------------- ---------------------- ----------------------------
								 				If cint(iCounter) = cint(iInstance) Then
								 					bFlag = True
								 					Exit for
									 			Else
													iCounter = iCounter + 1							 			
									 			End If
							 				'--------------------- ---------------------- ----------------------------
								 			End If
								 		Next
								 		If bFlag = True Then
								 			Exit for
								 		End If
							 	Next
							 	
						 		If bFlag = False Then
						 			Fn_PC_VariantConfigurationView_Operation = False
						 			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Fail to verify Variant Option [ "&arrOptions(iCnt)&" ] in Variant Configuration View.")
						 			Exit For
						 		Else
						 			Fn_PC_VariantConfigurationView_Operation = True
						 		End If		
						 Next
						 ' Minimize Tab
						 If StrTabName <> "" Then
						 		If Fn_TabFolder_Operation("DoubleClickTab", StrTabName,"") = False then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Fail to double click on Tab [ "&StrTabName&" ]" )
									Fn_PC_VariantConfigurationView_Operation = False
									Set objVarConfigTable = Nothing
									Set objPCWindow = Nothing
									Exit function
								End if
								' Close Tab
								If sTabClose = "Yes" Then
									Call Fn_TabFolder_Operation("Close", StrTabName,"")
									Call Fn_ReadyStatusSync(1)
								End If	
						 End If
				End If	
		'=============================================================================================================================
	End  Select	
	
	Set objPCWindow = Nothing
	
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_PC_ConfiguratorRules_Operation
'@@
'@@    Description			:	Function Used to Perform operation on Configurator Rules
'@@
'@@    Parameters			:	1. sAction		: Action to be performed
'@@							:	2. StrTabName	: Tab Name
'@@							:	2. dicDetails	: Dictionary object
'@@							:	2. Popupmenu	: Popup menu name
'@@							:	3. sTabClose	: Flag to close Tab
'@@
'@@    Return Value		   	: 	True Or False
'@@
'@@    Pre-requisite		:	Product Configurator Perspective should be opened 
'@@
'@@    Examples				:	Set dicDetails = CreateObject( "Scripting.Dictionary")
'@@ 								dicDetails("IDs") = "Group ID:0001~Family ID:0002~Feature ID:0003"
'@@ 								dicDetails("ConfigureCheckbox") = "Availability:ON~Default:ON~Exclusive:ON~Inclusive:ON"
'@@									dicDetails("ColumnName")="ID~Type~Severity"
'@@									dicDetails("ColumnValues")="0005~Exclusion Rule~Error"
'@@									dicDetails("ToolBarButton") = "Expand"
'@@ 							bReturn = Fn_PC_ConfiguratorRules_Operation("CreateConfigRule","000202-PRoduct (Configurator Rules)",dicDetails,"","Yes")
'@@  
'@@ 							Set dicDetails = CreateObject( "Scripting.Dictionary")
'@@ 								dicDetails("IDs") = "Group ID:0001~Family ID:0002~Feature ID:0003"	 
'@@									dicDetails("ToolBarButton") = "Search"
'@@ 							bReturn = Fn_PC_ConfiguratorRules_Operation("SearchConfigurationRules","000202-PRoduct (Configurator Rules)",dicDetails,"","Yes")
'@@  
'@@ 							Set dicDetails = CreateObject( "Scripting.Dictionary")
'@@									dicDetails("ColumnName")="ID~Type~Severity"
'@@									dicDetails("ColumnValues")="0005~Exclusion Rule~Error"
'@@ 							bReturn = Fn_PC_ConfiguratorRules_Operation("VerifyConfigurationRules","000202-PRoduct (Configurator Rules)",dicDetails,"","Yes")
'@@
'@@-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@	   History				:	Developer Name				Date	  			Rev. No.		Changes Done		Reviewer
'@@-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@	   Created By 			:	Poonam Chopade			12-Aug-2019				1.0		  		 Created		  [TC12.3(2019070800)-12Aug2019-PoonamC-NewDevelopment]
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_PC_ConfiguratorRules_Operation(sAction,StrTabName,dicDetails,Popupmenu,sTabClose)

	GBL_FAILED_FUNCTION_NAME = "Fn_PC_ConfiguratorRules_Operation"
	Dim objPCWindow,objConfigRuleTable,arrColValues,arrColumns,iX,iY,iRowCount
	Dim sMenu,bFlag,iCnt,iRowIndex,stoolbarBtn,iCnt1,sAppColName,sAppColValue
	Dim objJavaList,childObjects
	Dim ConfigurationProfileDialog,RButton,objdes,CBox,CBoxMode,CChecked
	
	Fn_PC_ConfiguratorRules_Operation = False
	
	Set objConfigRuleTable = Fn_PC_GetObject("ConfiguratorRules")
	Set objPCWindow = Fn_PC_GetObject("ProductConfigurator")
	On Error Resume Next
	
	'Check Existence of Configurator Rules Table
	If Fn_UI_ObjectExist("Fn_PC_ConfiguratorRules_Operation",objConfigRuleTable) = False Then
		sMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("Configurator_Toolbar"),"OpenConfiguratorRulesview")
		Call Fn_ToolBarOperation("Click",sMenu,"")
		Call Fn_ReadyStatusSync(1)
		If Fn_UI_ObjectExist("Fn_PC_ConfiguratorRules_Operation",objConfigRuleTable) = False Then ' Check Existence of Configurator Rules Table
			Fn_PC_ConfiguratorRules_Operation = False
			Set objConfigRuleTable = Nothing
			Set objPCWindow = Nothing
			Exit Function
		End  If
	End If
	If StrTabName <> "" Then  ' Maximize tab
 	   Call Fn_TabFolder_Operation("DoubleClickTab", StrTabName,"")
	   Call Fn_ReadyStatusSync(1)
	End If 
	If Fn_ToolBarOperation("IsSelected","Navigation Pane","") Then
		Call Fn_ToolBarOperation("Click","Navigation Pane","")
	End If
	Select Case sAction
		Case "CreateConfigRule","CreateSVRule"
			 iRowIndex = cint(objConfigRuleTable.Object.getItemCount())-1 'Get Current Row Index to create new rule
			 IF dicDetails("ColumnName") <> "" And dicDetails("ColumnValues") <> "" Then
				arrColumns = Split(dicDetails("ColumnName"),"~")
				arrColValues = Split(dicDetails("ColumnValues"),"~")
				For iCnt = 0 to UBound(arrColumns) 'loop to enter values to column
					sAppColName = Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\AutomationXML\ApplicationInformationXML\RAC_ConfiguratorTable_APL.xml",arrColumns(iCnt))
					If sAppColName = False Then 
						sAppColName = arrColumns(iCnt)
					End If
					If sAppColName= "ID" OR sAppColName= "Sequence Number" Then
					
					objConfigRuleTable.SelectCell iRowIndex,sAppColName
					Else
					objConfigRuleTable.ActivateCell iRowIndex,sAppColName
					End If
					wait 2
					If sAction <> "CreateSVRule"  AND sAppColName <>"ID" AND sAppColName <> "Sequence Number" Then
						Call Fn_KeyBoardOperation("SendKey", "{ENTER}")
						
					End If
					If sAction = "CreateSVRule" Then
						
					JavaWindow("ProductConfigurator").JavaEdit("Text").SetTOProperty "path","Text;Table;SavedVariantRulesPanel;Composite;Composite;Form;Composite;ContributedPartRenderer\$1;Composite;CTabFolder;Composite;Composite;Composite;Composite;Composite;Shell;"
					End If
					
					wait 5
					If Fn_UI_ObjectExist("Fn_PC_ConfiguratorRules_Operation",objPCWindow.JavaEdit("Text")) Then
						bFlag = Fn_Edit_Box("Fn_PC_ConfiguratorRules_Operation",objPCWindow,"Text",arrColValues(iCnt))
					Else
						Set objJavaList=Description.Create()
							objJavaList("Class Name").value = "JavaList"
						Set childObjects = objPCWindow.ChildObjects(objJavaList)
							childObjects(0).Select arrColValues(iCnt)
							wait 1
						Set objJavaList = Nothing
						Set childObjects = Nothing
						bFlag = True
					End IF
				
				    If bFlag = False Then
						Fn_PC_ConfiguratorRules_Operation = False
						Set objConfigRuleTable = Nothing
						Set objPCWindow = Nothing
						Exit Function
					End if			
				Next
				Wait 5
				objConfigRuleTable.ActivateRow iRowIndex
				Wait 2
				If dicDetails("Save") = "" Then
					sMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("RAC_Toolbar"),"Savethecurrentcontents")
					Call Fn_ToolBarOperation("Click",sMenu,"") 'Save Rule
	 				Call Fn_ReadyStatusSync(1)
				End If
	
 				If dicDetails("Save") <>""  Then
 					sAppColName = Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\AutomationXML\ApplicationInformationXML\RAC_ConfiguratorTable_APL.xml","Name")
 				Else
 				   sAppColName = Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\AutomationXML\ApplicationInformationXML\RAC_ConfiguratorTable_APL.xml","ID")
 				End If
				
				sAppColValue = objConfigRuleTable.GetCellData(iRowIndex,sAppColName)
				If IsEmpty(sAppColValue) or sAppColValue = "" Then
					Fn_PC_ConfiguratorRules_Operation = False
				Else
					Fn_PC_ConfiguratorRules_Operation = sAppColValue				
				End If
			 End IF
		'=====================================================================================================	 
		Case "SearchConfigurationRules"
			arrColumns = Split(dicDetails("IDs"),"~")
			For iCnt = 0 To UBound(arrColumns)
				arrColValues = Split(arrColumns(iCnt),":")
				objPCWindow.JavaStaticText("IDLabel").SetTOProperty "label",arrColValues(0)+":"
				bFlag = Fn_Edit_Box("Fn_PC_ConfiguratorRules_Operation",objPCWindow,"IDEditbox",arrColValues(1))
				If bFlag = False Then
					Fn_PC_ConfiguratorRules_Operation = False
					Set objConfigRuleTable = Nothing
					Set objPCWindow = Nothing
					Exit Function
				End if
			Next
			Fn_PC_ConfiguratorRules_Operation = True
		'=====================================================================================================
		Case "VerifyConfigurationRules"
			iRowIndex = cint(objConfigRuleTable.Object.getItemCount())-1 'get rows count
			If Isempty(iRowIndex) Then
				iRowIndex = cint(objConfigRuleTable.GetROProperty("rows"))
			End If
			arrColumns = Split(dicDetails("ColumnName"),"~")
			arrColValues = Split(dicDetails("ColumnValues"),"~")
			For iCnt = 0 to UBound(arrColumns)
			  	 sAppColName = Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\AutomationXML\ApplicationInformationXML\RAC_ConfiguratorTable_APL.xml",arrColumns(iCnt))
			  	 If sAppColName = False Then 
				 	 sAppColName = arrColumns(iCnt)
			   	 End If
			  	 bFlag = False
				 For iCnt1 = 0 to iRowIndex
				 	If cstr(objConfigRuleTable.GetCellData(iCnt1,sAppColName)) <> "" And cstr(objConfigRuleTable.GetCellData(iCnt1,sAppColName)) <> Empty Then
				 		If instr(cstr(objConfigRuleTable.GetCellData(iCnt1,sAppColName)),cstr(arrColValues(iCnt))) > 0 Then
				 			bFlag = True
				 			Exit For
				 		ElseIf IsNumeric(objConfigRuleTable.GetCellData(iCnt1,sAppColName)) And IsNumeric(arrColValues(iCnt)) Then
				 			If instr(cint(objConfigRuleTable.GetCellData(iCnt1,sAppColName)),cint(arrColValues(iCnt))) > 0 Then 
								bFlag = True
								Exit For	
							End If
						End If	
				 	End If	 
				 Next
			  	If bFlag = False Then
					Fn_PC_ConfiguratorRules_Operation = False
					Set objConfigRuleTable = Nothing
					Set objPCWindow = Nothing
					Exit Function
				End if 
			Next
			Fn_PC_ConfiguratorRules_Operation = True
		'=====================================================================================================
		 Case "DeleteConfiguratorRule"
			 iRowIndex = cint(objConfigRuleTable.Object.getItemCount())-1 'get rows count
			 IF dicDetails("ColumnName") <> "" And dicDetails("ColumnValues") <> "" Then
			  	 sAppColName = Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\AutomationXML\ApplicationInformationXML\RAC_ConfiguratorTable_APL.xml",dicDetails("ColumnName"))
			  	 If sAppColName = False Then 
				 	 sAppColName = dicDetails("ColumnName")
			   	 End If
			  	 bFlag = False
				 For iCnt1 = 0 to iRowIndex
				   	  If instr(cstr(objConfigRuleTable.GetCellData(iCnt1,sAppColName)),cstr(dicDetails("ColumnValues"))) > 0 Then
					   	 	bFlag = True
					   	 	iRowIndex = iCnt1
					   	 	Exit For
				   	  End If
				 Next
			  	If bFlag = False Then
					Fn_PC_ConfiguratorRules_Operation = False
					Set objConfigRuleTable = Nothing
					Set objPCWindow = Nothing
					Exit Function
				End if
				objConfigRuleTable.ActivateRow iRowIndex
				Call Fn_ReadyStatusSync(1)
				Fn_PC_ConfiguratorRules_Operation = Fn_TcObjectDelete(False,"","Toolbar")
				Call Fn_ReadyStatusSync(1)
			 End if
		'=====================================================================================================
		 Case "SelectConfiguratorRule"
			 iRowIndex = cint(objConfigRuleTable.Object.getItemCount())-1 'get rows count
			 IF dicDetails("ColumnName") <> "" And dicDetails("ColumnValues") <> "" Then
			  	 sAppColName = Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\AutomationXML\ApplicationInformationXML\RAC_ConfiguratorTable_APL.xml",dicDetails("ColumnName"))
			  	 If sAppColName = False Then 
				 	 sAppColName = dicDetails("ColumnName")
			   	 End If
			  	 bFlag = False
				 For iCnt1 = 0 to iRowIndex
				   	  If instr(cstr(objConfigRuleTable.GetCellData(iCnt1,sAppColName)),cstr(dicDetails("ColumnValues"))) > 0 Then
					   	 	bFlag = True
					   	 	iRowIndex = iCnt1
					   	 	Exit For
				   	  End If
				 Next
			  	If bFlag = False Then
					Fn_PC_ConfiguratorRules_Operation = False
					Set objConfigRuleTable = Nothing
					Set objPCWindow = Nothing
					Exit Function
				End if
				objConfigRuleTable.ActivateRow iRowIndex
				Call Fn_ReadyStatusSync(1)
				Fn_PC_ConfiguratorRules_Operation = True
			 End if
			 '=====================================================================================================
		 Case "PopupMenuSelect","SVRPopupMenuSelect"
		 	iRowIndex = cint(objConfigRuleTable.Object.getItemCount())-1 'get rows count
			arrColumns = Split(dicDetails("ColumnName"),"~")
			arrColValues = Split(dicDetails("ColumnValues"),"~")
			For iCnt = 0 to UBound(arrColumns)
			  	 sAppColName = Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\AutomationXML\ApplicationInformationXML\RAC_ConfiguratorTable_APL.xml",arrColumns(iCnt))
			  	 If sAppColName = False Then 
				 	 sAppColName = arrColumns(iCnt)
			   	 End If
			  	 bFlag = False
			  	 iX = 25
			  	 iY = 30
				 For iCnt1 = 0 to iRowIndex
				   	  If instr(cstr(objConfigRuleTable.GetCellData(iCnt1,sAppColName)),cstr(arrColValues(iCnt))) > 0 Then
					   	  	Wait 2
					   	  	objConfigRuleTable.SelectRow iCnt1
					   	  	Wait 1
					   	  	If sAction = "SVRPopupMenuSelect" Then
					   	  		iX = 70
					   	  		iY = 25
					   	  	End If
					   	  	objConfigRuleTable.Click iX,iY,"RIGHT"
					   	    Wait 8
						   	'Select Menu action
						   	aMenu = split(Popupmenu,":",-1,1)
							Select Case Ubound(aMenu)
								Case "0"
									StrMenu = objPCWindow.WinMenu("ContextMenu").BuildMenuPath(aMenu(0))
								Case "1"
									StrMenu = objPCWindow.WinMenu("ContextMenu").BuildMenuPath(aMenu(0),aMenu(1))
								Case "2"
									StrMenu = objPCWindow.WinMenu("ContextMenu").BuildMenuPath(aMenu(0),aMenu(1),aMenu(2))
								Case Else
									Fn_PC_ConfiguratorRules_Operation = FALSE
									Exit Function
							End Select
							Err.Clear
							objPCWindow.WinMenu("ContextMenu").Select StrMenu
							If Err.number < 0 Then
								Fn_PC_ConfiguratorRules_Operation = False
							Else
								bFlag = True
								Fn_PC_ConfiguratorRules_Operation = True
							End If
						   	Exit For
				   	  	Else
				   	  		iY = iY + 20
				   	  	End If
				 Next
			  	If bFlag = False Then
					Fn_PC_ConfiguratorRules_Operation = False
					Set objConfigRuleTable = Nothing
					Set objPCWindow = Nothing
					Exit Function
				End if 
			Next
			Fn_PC_ConfiguratorRules_Operation = True
		'=====================================================================================================
		 Case "GetID"
		 		iRowIndex = cint(objConfigRuleTable.Object.getItemCount())-1 'get rows count
				arrColumns = Split(dicDetails("ColumnName"),"~")
				arrColValues = Split(dicDetails("ColumnValues"),"~")
				For iCnt = 0 to UBound(arrColumns)
				  	 sAppColName = Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\AutomationXML\ApplicationInformationXML\RAC_ConfiguratorTable_APL.xml",arrColumns(iCnt))
				  	 If sAppColName = False Then 
					 	 sAppColName = arrColumns(iCnt)
				   	 End If
				  	 bFlag = False
					 For iCnt1 = 0 to iRowIndex
					   	  If instr(cstr(objConfigRuleTable.GetCellData(iCnt1,sAppColName)),cstr(arrColValues(iCnt))) > 0 Then
						   	 Wait 1
							 sAppColName = Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\AutomationXML\ApplicationInformationXML\RAC_ConfiguratorTable_APL.xml","ID")
							 sAppColValue = objConfigRuleTable.GetCellData(iCnt1,sAppColName)
							 If IsEmpty(sAppColValue) or sAppColValue = "" Then
								Fn_PC_ConfiguratorRules_Operation = False
							 Else
								Fn_PC_ConfiguratorRules_Operation = sAppColValue				
							 End If
							 Exit For
						  End If
					 Next
				Next
			'==========================================================================================================
			Case "VerifyColumnValuesForRuleID"
				iRowIndex = cint(objConfigRuleTable.Object.getItemCount())-1 'get rows count
				arrColumns = Split(dicDetails("ColumnName"),"~")
				arrColValues = Split(dicDetails("ColumnValues"),"~")
				sAppColName = Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\AutomationXML\ApplicationInformationXML\RAC_ConfiguratorTable_APL.xml","ID")
			 	If sAppColName = False Then 
				 	sAppColName = "ID"
			 	End If
				For iCnt = 0 to iRowIndex
					bFlag = False
					If cstr(objConfigRuleTable.GetCellData(iCnt,sAppColName)) <> "" or cstr(objConfigRuleTable.GetCellData(iCnt,sAppColName)) <> Empty Then
						If instr(cstr(objConfigRuleTable.GetCellData(iCnt,sAppColName)),cstr(dicDetails("RuleID"))) > 0 OR instr(cint(objConfigRuleTable.GetCellData(iCnt,sAppColName)),cint(dicDetails("RuleID"))) > 0 Then
							bFlag = True
							iCnt1 = iCnt
							Exit For
						End If
					End If	 
				Next
				If bFlag = False Then
					Fn_PC_ConfiguratorRules_Operation = False
					Set objConfigRuleTable = Nothing
					Set objPCWindow = Nothing
					Exit Function
				End if 
				
				For iCnt = 0 to UBound(arrColumns)  'loop through column verification for Particular Rule ID
				  	 sAppColName = Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\AutomationXML\ApplicationInformationXML\RAC_ConfiguratorTable_APL.xml",arrColumns(iCnt))
				  	 If sAppColName = False Then 
					 	 sAppColName = arrColumns(iCnt)
				   	 End If
				  	 bFlag = False
					 If cstr(objConfigRuleTable.GetCellData(iCnt1,sAppColName)) <> "" And cstr(objConfigRuleTable.GetCellData(iCnt1,sAppColName)) <> Empty Then
						If instr(cstr(objConfigRuleTable.GetCellData(iCnt1,sAppColName)),cstr(arrColValues(iCnt))) > 0 Then 
							bFlag = True
						ElseIf IsNumeric(objConfigRuleTable.GetCellData(iCnt1,sAppColName)) And IsNumeric(arrColValues(iCnt)) Then
							If instr(cint(objConfigRuleTable.GetCellData(iCnt1,sAppColName)),cint(arrColValues(iCnt))) > 0 Then 
								bFlag = True
							End If
						End If
					End If	
				  	If bFlag = False Then
						Fn_PC_ConfiguratorRules_Operation = False
						Set objConfigRuleTable = Nothing
						Set objPCWindow = Nothing
						Exit Function
					End if 
				Next
				Fn_PC_ConfiguratorRules_Operation = True
			'===========================================================================================================
			Case "MultiSelectConfiguratorRule"
				 iRowCount = cint(objConfigRuleTable.Object.getItemCount())-1 'get rows count
				 IF dicDetails("ColumnName") <> "" And dicDetails("ColumnValues") <> "" Then
					 	arrColumns = Split(dicDetails("ColumnName"),"~")
						arrColValues = Split(dicDetails("ColumnValues"),"~")
					    For iCnt = 0 to UBound(arrColValues)
						  	 sAppColName = Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\AutomationXML\ApplicationInformationXML\RAC_ConfiguratorTable_APL.xml",arrColumns(iCnt))
						  	 If sAppColName = False Then 
							 	 sAppColName = arrColumns(iCnt)
						   	 End If
						  	 bFlag = False
							 For iCnt1 = 0 to iRowCount
							   	  If instr(cstr(objConfigRuleTable.GetCellData(iCnt1,arrColumns(iCnt))),cstr(arrColValues(iCnt))) > 0 Then
								   	 	bFlag = True
								   	 	iRowIndex = iCnt1
								   	 	Exit For
							   	  End If
							 Next
						  	If bFlag = False Then
								Fn_PC_ConfiguratorRules_Operation = False
								Set objConfigRuleTable = Nothing
								Set objPCWindow = Nothing
								Exit Function
						    End if
						    If iCnt = 0 Then
						    	objConfigRuleTable.ActivateRow iRowIndex
						    Else
						    	objConfigRuleTable.ExtendRow iRowIndex 	 
						    End If
							Call Fn_ReadyStatusSync(1)
						Next	
					Fn_PC_ConfiguratorRules_Operation = True
				 End if
			'===========================================================================================================
		 Case "ConfigurationProfileSetting" 		
		 		Set ConfigurationProfileDialog=JavaWindow("ProductConfigurator").JavaWindow("ConfiguratorContext")
				wait 1
				ConfigurationProfileDialog.SetTOProperty "title","Configuration Profile Settings For New Variant Rules"
				CBox = Split(dicDetails("Option"),"~")
				CBoxMode = Split(dicDetails("OptionValues"),"~")
				
				If Fn_UI_ObjectExist("Fn_PC_ConfigurationContext_Operations",ConfigurationProfileDialog) = False Then
					call Fn_ToolBarOperation("ShowDropdownAndSelectWithInstance", "View Menu:1", "Configuration Profile Settings")
				End If
				
				If Fn_UI_ObjectExist("Fn_PC_ConfigurationContext_Operations",ConfigurationProfileDialog) = False Then
					Fn_PC_ConfiguratorRules_Operation = False
				Else
					RButton = dicDetails("Mode")
					ConfigurationProfileDialog.JavaRadioButton("Overlay").SetTOProperty "attached text", RButton
					wait 2
					ConfigurationProfileDialog.JavaRadioButton("Overlay").Set "ON"
					wait 1
					Set objdes = Nothing
					For iCnt = 1 To 5
						Set objdes = Nothing
						ConfigurationProfileDialog.JavaCheckBox("OpenOnCreate").SetTOProperty "attached text", CBox(iCnt-1)
						CChecked = Fn_UI_Object_GetROProperty("Fn_PC_ConfiguratorRules_Operation",ConfigurationProfileDialog.JavaCheckBox("OpenOnCreate"),"value")
						If CBoxMode(iCnt-1) = "OFF" Then
							If CChecked = 1 Then
								ConfigurationProfileDialog.JavaCheckBox("OpenOnCreate").Set "OFF"
							End If
						Else
							If CChecked = 0 Then
								ConfigurationProfileDialog.JavaCheckBox("OpenOnCreate").Set "ON"
							End If
						End If
						wait 1
					Next
					ConfigurationProfileDialog.JavaButton("Button").SetTOProperty "label" ,"OK"
					ConfigurationProfileDialog.JavaButton("Button").Click

					dicDetails.RemoveAll()
					Fn_PC_ConfiguratorRules_Operation = True
				End If
				
			'===========================================================================================================
		 Case "VerifyConfigurationProfileSetting" 		
		 		Set ConfigurationProfileDialog=Dialog("Information")
		 		ConfigurationProfileDialog.SetTOProperty "regexpwndtitle","Configuration Profile Settings For New Variant Rules"
				ConfigurationProfileDialog.SetTOProperty "text","Configuration Profile Settings For New Variant Rules"
				CBox = Split(dicDetails("Option"),"~")
				CBoxMode = Split(dicDetails("OptionValues"),"~")
				
				If Fn_UI_ObjectExist("Fn_PC_ConfigurationContext_Operations",ConfigurationProfileDialog) = False Then
					call Fn_ToolBarOperation("ShowDropdownAndSelectWithInstance", "View Menu:1", "Configuration Profile Settings")
				End If
				
				If Fn_UI_ObjectExist("Fn_PC_ConfigurationContext_Operations",ConfigurationProfileDialog) = False Then
					Fn_PC_ConfiguratorRules_Operation = False
				Else
					RButton = dicDetails("Mode")
					set objdes = Description.Create()
					objdes("Class Name").Value="WinRadioButton"
					objdes("text").Value=RButton
					dicDetails.RemoveAll()
					set dicDetails = ConfigurationProfileDialog.ChildObjects(objdes)
					CChecked = dicDetails(0).getROProperty("Checked")
					if CChecked <> "ON" Then
						Fn_PC_ConfiguratorRules_Operation = False
						Exit Function
					else
						Fn_PC_ConfiguratorRules_Operation = True
					End If
					Set objdes = Nothing
					For iCnt = 1 To 5
						Set objdes = Nothing
						set objdes = Description.Create()
						objdes("Class Name").Value="WinCheckBox"
						objdes("text").Value=CBox(iCnt-1)
						dicDetails.RemoveAll()
						set dicDetails = ConfigurationProfileDialog.ChildObjects(objdes)
						CChecked = dicDetails(0).getROProperty("Checked")
						If CBoxMode(iCnt-1) = CChecked Then
							Fn_PC_ConfiguratorRules_Operation = True
						Else
							Fn_PC_ConfiguratorRules_Operation = False
							Exit Function
						End If
					Next
					ConfigurationProfileDialog.WinButton("OK").Click
					dicDetails.RemoveAll()
					Fn_PC_ConfiguratorRules_Operation = True
				End If
			'==========================================================================================================					
	End Select
	
	 If dicDetails("ToolBarButton") <> "" Then    ' Click toolbar button in Configuration Rule tab
 		stoolbarBtn = Split(dicDetails("ToolBarButton"),"~")
 		For iCnt = 0 To UBound(stoolbarBtn)
 			Call Fn_ToolBarOperation("Click",stoolbarBtn(iCnt),"")
 			Call Fn_ReadyStatusSync(1)
 		Next
	 End If
	
	 If StrTabName <> "" Then ' Minimize Tab
 		Call Fn_TabFolder_Operation("DoubleClickTab", StrTabName,"")
		Call Fn_ReadyStatusSync(1)
		If sTabClose = "Yes" Then ' Close Tab
			Call Fn_TabFolder_Operation("Close", StrTabName,"")
			Call Fn_ReadyStatusSync(1)
		End If	
	 End If
	If Fn_ToolBarOperation("IsSelected","Navigation Pane","") = False Then
		Call Fn_ToolBarOperation("Click","Navigation Pane","")
	End If
'	  If StrTabName <> "" Then ' Minimize Tab
'    	  Call Fn_TabFolder_Operation("DoubleClickTab", StrTabName,"")
' 		Call Fn_ReadyStatusSync(1)
' 	End  If 	
 	
	Set objConfigRuleTable = Nothing
	Set objPCWindow = Nothing
	
End Function
'=====================================================================================================================================================================
'@@
'@@    Function Name			:	Fn_PC_EffectivityOperations
'@@
'@@    Description				:	Function Used to perform Effectivity Operations
'@@
'@@    Parameters			    :	1. sAction			: Action to be performed
'@@								:	2. sNode			: Node name with complete Path
'@@								:	3. sEffectivityNode : Node to Select in Effectivity tab
'@@								:	4. sValue			: Value to be verified / Row Number 
'@@								:	5. sFromUnit 		: Unit Starting Value (~ separated list of From Units)
'@@								:	6. sToUnit			: Unit End Value (~ separated list of To Units)
'@@								:	7. sInDate			: In Date value (~ separated list of Date strings eg. 02-Jan-2012~02-Jan-2012$12:30 )
'@@								:	8. sOutDates		: Out Date value (~ separated list of Date strings 02-Jan-2012~SO~UP )
'@@								: 
'@@    Return Value		   	   	: 	True Or False
'@@
'@@    Pre-requisite			:	Product Configurator module window should be displayed.	
'@@					
'@@    Examples					:	Call Fn_PC_EffectivityOperations("SetInEffectivityTab", "CD000010;1-CD:DE000001/001;1-de", "", "5", "SO", "", "")
'@@
'@@	   History					:	Developer Name				Date			Rev. No.	Changes Done		Reviewer
'@@-----------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@									Poonam Chopade			16-Aug-2019			1.0			Created				TC12.3(2019070800)_NewDevelopment_PoonamC_16Aug2019
'====================================================================================================================================================================
Public Function Fn_PC_EffectivityOperations(sAction,sNode,sEffectivityNode,sValue,sFromUnit,sToUnit,sInDate,sOutDate)
	GBL_FAILED_FUNCTION_NAME="Fn_PC_EffectivityOperations"
	Dim intNoOfObjects, i, bReturn,aFromUnit, aToUnit, aInDate, aOutDate,sMenu
	Dim iLimit,objEff,bStatus,arrDate,iKeyCnt,iCnt
	Dim objSelectType,objEffDialog,objEffectivityTable
	Dim objJavaList,childObjects
	
	Fn_PC_EffectivityOperations = False
	iLimit = 0

	aFromUnit = Split(sFromUnit,"~")
	iLimit = uBound(aFromUnit)
	
	aToUnit = Split(sToUnit,"~")
	If  uBound(aToUnit) > iLimit Then
		iLimit = uBound(aToUnit)
	End IF
	
	aInDate = Split(sInDate,"~")
	If  uBound(aInDate) > iLimit Then
		iLimit = uBound(aInDate)
	End IF
	
	aOutDate = Split(sOutDate,"~")
	If  uBound(aOutDate) > iLimit Then
		iLimit = uBound(aOutDate)
	End IF
	
	Set objEff = Fn_PC_GetObject("ProductConfigurator") 
	
	Select Case sAction
		'===============================================================================================================================
		Case "VerifyColumnExistInEffectivityTab","SetInEffectivityTab","DeleteInEffectivityTab","VerifyBlankInEffectivityTab"
			 If sNode <> "" Then 'Check Effectivity View tab 
				If Fn_TabFolder_Operation("Exist","Effectivity View","")=False And Fn_TabFolder_Operation("Exist","*Effectivity View","")=False Then
					sMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("RAC_PopupMenu"),"OpenwithEffectivityView")
					bReturn = Fn_PC_NavTree_NodeOperation("PopupMenuSelect",sNode,sMenu) 
					If bReturn = False Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_PC_EffectivityOperations : Failed to select popup menu of [ " & sNode & " ].")
						Set objEff = Nothing	
						Exit function
					End If
				Else
					bReturn = Fn_PC_NavTree_NodeOperation("Select",sNode,"") 
				End If
			End If
			Call Fn_ReadyStatusSync(1)	
			
			Select Case sAction
			   '-------------------------------------------------------------------------------------------------------------------------
				Case "VerifyColumnExistInEffectivityTab"
					intNoOfObjects = cInt(objEff.JavaTree("EffectivityTree").GetROProperty("columns_count"))
					For iCnt = 0 to intNoOfObjects - 1
						If objEff.JavaTree("EffectivityTree").GetColumnHeader(iCnt) = sValue Then
							Fn_PC_EffectivityOperations=True
							Exit For 
						End If
					Next
			    '-------------------------------------------------------------------------------------------------------------------------						
				Case "SetInEffectivityTab"
					For i = 0 to iLimit
						bStatus=False
						If Fn_PC_EffectivityOperations("VerifyColumnExistInEffectivityTab", "", "","To Unit", "", "", "", "") = True Then
							Call Fn_UI_ClickJavaTreeCell("Fn_PC_EffectivityOperations", objEff, "EffectivityTree", sEffectivityNode, "To Unit","LEFT")
						End If
						If sFromUnit <> "" Then
							If aFromUnit(i)<>"" Then 'setting From Unit
									Call Fn_UI_ClickJavaTreeCell("Fn_PC_EffectivityOperations", objEff, "EffectivityTree", sEffectivityNode, "From Unit","LEFT") 
									wait 1
									Set objShell = CreateObject("Wscript.Shell")
									If  i = 0 Then
										Call Fn_KeyBoardOperation("SendKeys", "{DOWN}")
										wait 1
										Call Fn_KeyBoardOperation("SendKeys", "{ENTER}")
									Else
										Call Fn_Button_Click("Fn_PC_EffectivityOperations", objEff,"AddButton")
										bStatus=True
										wait 1
										Call Fn_UI_ClickJavaTreeCell("Fn_PC_EffectivityOperations", objEff, "EffectivityTree", sEffectivityNode, "From Unit","LEFT")
										For iKeyCnt = 0 To i
											Call Fn_KeyBoardOperation("SendKeys", "{DOWN}")
											wait 1
										Next
										Call Fn_KeyBoardOperation("SendKeys", "{ENTER}")
									End If
									wait 1
									objShell.SendKeys "^A"
									wait 1
									objShell.SendKeys aFromUnit(i)
									wait 1
									Call Fn_KeyBoardOperation("SendKeys", "{ENTER}")
									Call Fn_KeyBoardOperation("SendKeys", "{RIGHT}")
									Call Fn_KeyBoardOperation("SendKeys", "{RIGHT}")
									Set objShell=Nothing
							End If
							
							If sToUnit<>"" Then  ' setting To Unit
								If aToUnit(i)<>"" Then	
									If bStatus=False and i > 0 Then
										Call Fn_Button_Click("Fn_PC_EffectivityOperations", objEff,"AddButton")
										bStatus=True
									End If 
									Call Fn_UI_ClickJavaTreeCell("Fn_PC_EffectivityOperations", objEff, "EffectivityTree", sEffectivityNode, "To Unit","LEFT")
									wait 1
									Set objShell = CreateObject("Wscript.Shell")
										For iKeyCnt = 0 To i
											Call Fn_KeyBoardOperation("SendKeys", "{DOWN}")
										Next
										wait 1
										Call Fn_KeyBoardOperation("SendKeys", "{ENTER}")
										wait 1
									If trim(aToUnit(i)) <> "" Then
										Set objJavaList=Description.Create()
											objJavaList("Class Name").value = "JavaList"
										Set childObjects = objEff.ChildObjects(objJavaList)
											childObjects(0).Type aToUnit(i)
										wait 1
										Set objJavaList = Nothing
										Set childObjects = Nothing
									End If
									Call Fn_KeyBoardOperation("SendKeys", "{ENTER}")
									Call Fn_KeyBoardOperation("SendKeys", "{RIGHT}")
									Set objShell=Nothing
								End If
							End If 
						End If
			
						If sInDate<>"" Then ' setting IN Date
							If Fn_PC_EffectivityOperations("VerifyColumnExistInEffectivityTab", "", "","From Unit", "", "", "", "")=True And sFromUnit="" Then
								Call Fn_UI_ClickJavaTreeCell("Fn_PC_EffectivityOperations", objEff, "EffectivityTree", sEffectivityNode, "From Unit","LEFT")
								Call Fn_KeyBoardOperation("SendKeys", "{RIGHT}")
								Call Fn_KeyBoardOperation("SendKeys", "{RIGHT}")
							End If
							If aInDate(i)<>"" Then
								If bStatus=False and i>0 Then
									Call Fn_Button_Click("Fn_PC_EffectivityOperations", objEff,"AddButton")
									bStatus=True
									Set objShell = CreateObject("Wscript.Shell")
									wait 1
									For iKeyCnt=0 To i
										Call Fn_KeyBoardOperation("SendKeys", "{TAB}")
									Next
									Set objShell=Nothing 
								End If
								
								Call Fn_SyncTCObjects()
								Call Fn_UI_ClickJavaTreeCell("Fn_PC_EffectivityOperations", objEff, "EffectivityTree", sEffectivityNode, "In Date","LEFT")
								wait 1
								Set objShell = CreateObject("Wscript.Shell")
								wait 1
								For iKeyCnt=0 To i
									Call Fn_KeyBoardOperation("SendKeys", "{DOWN}")
								Next
								Call Fn_KeyBoardOperation("SendKeys", "{ENTER}")
								wait 1
								Call Fn_KeyBoardOperation("SendKeys", "{DOWN}")
								If instr( aInDate(i),"$") > 0 Then
									arrDate = split(trim(aInDate(i)),"$")
									Call Fn_PC_DateControl("Set", arrDate(0), arrDate(1))
								Else
									Select Case lcase(trim(aInDate(i)))
										Case ""
											Call  Fn_PC_DateControl("Clear", "", "")
										Case "today"
											Call  Fn_PC_DateControl("Today", "", "")
										Case Else
											Call  Fn_PC_DateControl("Set", aInDate(i), "")
									End Select
								End If
							End If 
								Call Fn_KeyBoardOperation("SendKeys", "{ESC}")
								Call Fn_KeyBoardOperation("SendKeys", "{RIGHT}")								
							End If
							
							If sOutDate<>"" Then 'Set Out Date
								If aOutDate(i)<>"" Then
									If lcase(aOutDate(i)) <> "so" AND lcase(aOutDate(i)) <> "up" Then
											If bStatus=False and i>0  Then
												Call Fn_Button_Click("Fn_PC_EffectivityOperations", objEff,"AddButton")
												bStatus=True
											End If 
											Call Fn_SyncTCObjects()
											Call Fn_UI_ClickJavaTreeCell("Fn_PC_EffectivityOperations", objEff, "EffectivityTree", sEffectivityNode, "Out Date","LEFT")
											wait 1
											Set objShell = CreateObject("Wscript.Shell")
											For iKeyCnt=0 To i
												Call Fn_KeyBoardOperation("SendKeys", "{DOWN}")
											Next
											Call Fn_KeyBoardOperation("SendKeys", "{ENTER}")
											wait 1
											For iCnt = 1 To 3
												Call Fn_KeyBoardOperation("SendKeys", "{DOWN}")
											Next
											If instr( aOutDate(i),"$") > 0 Then
													arrDate = split(trim(aOutDate(i)),"$")
													Call  Fn_PC_DateControl("Set", arrDate(0), arrDate(1))
											Else
													Select Case lcase(trim(aOutDate(i)))
														Case ""
															Call  Fn_PC_DateControl("Clear", "", "")
														Case "today"
															Call  Fn_PC_DateControl("Today", "", "")
														Case Else
															Call  Fn_PC_DateControl("Set", aOutDate(i), "")
													End Select
											End If
									Else
											Call Fn_UI_ClickJavaTreeCell("Fn_PC_EffectivityOperations", objEff, "EffectivityTree", sEffectivityNode, "Out Date","LEFT")
											wait 1
											Set objShell = CreateObject("Wscript.Shell")
											For iKeyCnt=0 To i
												Call Fn_KeyBoardOperation("SendKeys", "{DOWN}")
											Next
											If lcase(aOutDate(i))="so" Then
												For iCnt = 1 To 2
													Call Fn_KeyBoardOperation("SendKeys", "{DOWN}")
												Next
											Elseif lcase(aOutDate(i))="up" Then
												Call Fn_KeyBoardOperation("SendKeys", "{DOWN}")
											End If
											Call Fn_KeyBoardOperation("SendKeys", "{ENTER}")
											wait 1
											Set objShell=Nothing
									End If
								End If
							End If
						Next
						wait 2
						sMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("RAC_Toolbar"),"Savetheselectedobject")
						Call Fn_ToolbarOperation("Click", sMenu,"")
						Call Fn_ReadyStatusSync(1)
						Fn_PC_EffectivityOperations = True
					'-------------------------------------------------------------------------------------------------------------------------						
					Case "DeleteInEffectivityTab"				
							For i = 0 to iLimit
								if (sFromUnit <> "" and sToUnit <> "" and sInDate <> "" and sOutDate <> "" AND uBound(aToUnit) <= iLimit) then
									Call Fn_UI_ClickJavaTreeCell("Fn_PC_EffectivityOperations", objEff, "EffectivityTree", sEffectivityNode, "From Unit","LEFT")
									wait 1
									Set objShell = CreateObject("Wscript.Shell")
									For iKeyCnt=0 To i
										Call Fn_KeyBoardOperation("SendKeys", "{DOWN}")
									Next
									Call Fn_KeyBoardOperation("SendKeys", "{ENTER}")
									wait 1
									Set objShell = Nothing
								ElseIf sFromUnit <> "" AND uBound(aToUnit) <= iLimit Then
									Call Fn_UI_ClickJavaTreeCell("Fn_PC_EffectivityOperations", objEff, "EffectivityTree", sEffectivityNode, "From Unit","LEFT")
									wait 1
									Set objShell = CreateObject("Wscript.Shell")
									For iKeyCnt=0 To i
										Call Fn_KeyBoardOperation("SendKeys", "{DOWN}")
									Next
									Call Fn_KeyBoardOperation("SendKeys", "{ENTER}")
									wait 1
									Set objShell = Nothing
								Elseif sInDate <> ""   AND uBound(aOutDate) <= iLimit Then
									Set objShell = CreateObject("Wscript.Shell")
									For iKeyCnt=0 To i
										Call Fn_KeyBoardOperation("SendKeys", "{DOWN}")
									Next
									Call Fn_KeyBoardOperation("SendKeys", "{ENTER}")
									wait 1
									Set objShell = Nothing
								End if
								Call Fn_Button_Click("Fn_PC_EffectivityOperations", objEff,"RemoveButton")
								wait 1
							Next
							wait 2
						sMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("RAC_Toolbar"),"Savetheselectedobject")
						Call Fn_ToolbarOperation("Click", sMenu,"")
						Call Fn_ReadyStatusSync(1)
						Fn_PC_EffectivityOperations = True
				
				' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
					Case "VerifyBlankInEffectivityTab"
							bFlag = True
							
						For i = 0 to iLimit
							' Verify From Unit
							If aFromUnit(0)="From Unit" Then
								Call Fn_ClickEffectivityTreeCell("Fn_PC_EffectivityOperations", objEff, "EffectivityTree", sEffectivityNode, "From Unit","LEFT")
								wait 1
								Set objShell = CreateObject("Wscript.Shell")
								For iKeyCnt=0 To i
									Call Fn_KeyBoardOperation("SendKeys", "{DOWN}")
								Next
								Call Fn_KeyBoardOperation("SendKeys", "{ENTER}")
								wait 1
								objShell.SendKeys "^A"
								wait 1 
								If cstr(JavaWindow("ProductConfigurator").JavaTab("Effectivity").JavaEdit("TableText").GetROProperty("text"))<>"" Then
									bFlag = False
								End If
								Call Fn_KeyBoardOperation("SendKeys", "{ESC}")
								Call Fn_KeyBoardOperation("SendKeys", "{RIGHT}")
								Set objShell=Nothing
							 End If 
													 
							 ' Verify To Unit
							 If aToUnit(0)="To Unit" Then
							 	Call Fn_ClickEffectivityTreeCell("Fn_PC_EffectivityOperations", objEff, "EffectivityTree", sEffectivityNode, "To Unit","LEFT")
								wait 1
								Set objShell = CreateObject("Wscript.Shell")
								For iKeyCnt=0 To i
									Call Fn_KeyBoardOperation("SendKeys", "{DOWN}")
								Next
								Call Fn_KeyBoardOperation("SendKeys", "{ENTER}")
								wait 1
'								If cstr(JavaWindow("Collaborative Product").JavaList("TableList").GetROProperty("value"))<>"" Then
'									bFlag = False
'								End If
								If cstr(objEff.WinEdit("WinEffTabText").GetROProperty("text"))<>"" Then
									bFlag = False
								End If
								Call Fn_KeyBoardOperation("SendKeys", "{ESC}")
								'Call Fn_KeyBoardOperation("SendKeys", "{TAB}")
								Call Fn_KeyBoardOperation("SendKeys", "{RIGHT}")
								Set objShell=Nothing
							 End If
							 ' Vrify In Date 
							 	If aInDate(0)="In Date" Then
							 		Call Fn_ClickEffectivityTreeCell("Fn_PC_EffectivityOperations", objEff, "EffectivityTree", sEffectivityNode, "In Date","LEFT")
									wait 1
									Set objShell = CreateObject("Wscript.Shell")
									For iKeyCnt=0 To i
										Call Fn_KeyBoardOperation("SendKeys", "{DOWN}")
									Next
									Call Fn_KeyBoardOperation("SendKeys", "{ENTER}")
									wait 1
'									If Trim(cstr(JavaWindow("Collaborative Product").JavaList("TableList").GetROProperty("value")))<>"" Then
'										bFlag = False
'									End If
									If cstr(objEff.WinEdit("WinEffTabText").GetROProperty("text"))<>"" Then
										bFlag = False
									End If
									Call Fn_KeyBoardOperation("SendKeys", "{ESC}")
								'Call Fn_KeyBoardOperation("SendKeys", "{TAB}")
								Call Fn_KeyBoardOperation("SendKeys", "{RIGHT}")
									Set objShell=Nothing
							 	End If
							 	'Verify Out Date
							 	If aOutDate(0)="Out Date" Then
							 		Call Fn_ClickEffectivityTreeCell("Fn_PC_EffectivityOperations", objEff, "EffectivityTree", sEffectivityNode, "Out Date","LEFT")
									wait 1
									Set objShell = CreateObject("Wscript.Shell")
									For iKeyCnt=0 To i
										Call Fn_KeyBoardOperation("SendKeys", "{DOWN}")
									Next
									Call Fn_KeyBoardOperation("SendKeys", "{ENTER}")
									wait 1
'									If Trim(cstr(JavaWindow("Collaborative Product").JavaList("TableList").GetROProperty("value")))<>"" Then
'										bFlag = False
'									End If
									If cstr(objEff.WinEdit("WinEffTabText").GetROProperty("text"))<>"" Then
										bFlag = False
									End If
									Call Fn_KeyBoardOperation("SendKeys", "{ESC}")
									Set objShell=Nothing
							 	End If
						Next
						If bFlag = False Then
							Fn_PC_EffectivityOperations = False
							 Exit Function
						Else
							Fn_PC_EffectivityOperations = bFlag
						End If
'									
			End  Select		
		'===============================================================================================================================
		Case "ViewChangeEffectivity"
			If sNode <> "" Then
				bReturn = Fn_PC_NavTree_NodeOperation("Select",sNode,"") 
				If bReturn = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_PC_EffectivityOperations : Failed to select Node [ " & sNode & " ].")
					Exit function
				End If
			End If
			If Fn_UI_ObjectExist("Fn_PC_EffectivityOperations",objEff.JavaWindow("ViewEditEffectivity")) = False Then
				Set objSelectType = Description.Create()
					objSelectType("Class Name").value = "JavaObject"
					objSelectType("toolkit class").value = "org.eclipse.ui.forms.widgets.ImageHyperlink"
				Set intNoOfObjects = objEff.ChildObjects(objSelectType)
				For i = 0 to intNoOfObjects.count-1
					If  lcase(trim( "" & intNoOfObjects(i).Object.getToolTipText())) = lCase("Click to view and change the current effectivity configuration") Then
						intNoOfObjects(i).Click 1,1, "LEFT"
						Exit for
					End If
				Next
			End IF
			Set objSelectType = Nothing
			Set intNoOfObjects = Nothing
			
			Set objEffDialog = objEff.JavaWindow("ViewEditEffectivity")
			Set objEffectivityTable = objEffDialog.JavaTable("EffectivityTable")
		 	Select Case sAction
				'---------------------------------------------------------------------------------------------
				Case "ViewChangeEffectivity"
					For i = 0 to iLimit
						If sFromUnit <> "" AND uBound(aToUnit) <= iLimit Then
							objEffectivityTable.ActivateCell i,"From Unit" ' setting From Unit
							wait 1
							If trim(aFromUnit(i)) = "" Then
								wait 1
								Call Fn_KeyBoardOperation("SendKeys", "{END}")
								For iKeyCnt = 0 to 10
									Call Fn_KeyBoardOperation("SendKeys", "{BKSP}")
								Next
							End If
							objEffDialog.JavaEdit("TableText").Set aFromUnit(i)
							objEffDialog.JavaEdit("TableText").Activate
							wait 1
							
							objEffectivityTable.ActivateCell i,"To Unit" ' setting To Unit
							wait 1
							Call Fn_KeyBoardOperation("SendKeys", "{END}")
							For iKeyCnt = 0 to 10
								Call Fn_KeyBoardOperation("SendKeys", "{BKSP}")
							Next
							If sToUnit<>"" Then	
								If trim(aToUnit(i)) <> "" Then
									objEffDialog.JavaList("TableList").Type aToUnit(i)
								End If
								Call Fn_KeyBoardOperation("SendKeys", "{ENTER}")
								wait 1
							End If
						End If
						
						If sInDate <> ""   AND uBound(aOutDate) <= iLimit Then
							objEffectivityTable.ActivateCell i,"In Date" ' setting IN Date
							wait 1
							Call Fn_Button_Click("Fn_PC_EffectivityOperations", objEffDialog, "TableDropDownList")
							wait 1
							If instr( aInDate(i),"$") > 0 Then
									arrDate = split(trim(aInDate(i)),"$")
									Call  Fn_PC_DateControl("Set", arrDate(0), arrDate(1))
							Else
									Select Case lcase(trim(aInDate(i)))
										Case ""
											Call  Fn_PC_DateControl("Clear", "", "")
										Case "today"
											Call  Fn_PC_DateControl("Today", "", "")
										Case Else
											Call  Fn_PC_DateControl("Set", aInDate(i), "")
									End Select
							End If
							
							objEffectivityTable.ActivateCell i, "Out Date" ' setting Out Date
							wait 1
							If lcase(aOutDate(i)) <> "so" AND lcase(aOutDate(i)) <> "up" Then
								For iCnt = 1 To 3
									Call Fn_KeyBoardOperation("SendKeys", "{DOWN}")
								Next
								If instr( aOutDate(i),"$") > 0 Then
										arrDate = split(trim(aOutDate(i)),"$")
										Call  Fn_PC_DateControl("Set", arrDate(0), arrDate(1))
								Else
										Select Case lcase(trim(aInDate(i)))
											Case ""
												Call  Fn_PC_DateControl("Clear", "", "")
											Case "today"
												Call  Fn_PC_DateControl("Today", "", "")
											Case Else
												Call  Fn_PC_DateControl("Set", aOutDate(i), "")
										End Select
								End If
							else
								If lcase(aOutDate(i))="so" Then
									For iCnt = 1 To 2
										Call Fn_KeyBoardOperation("SendKeys", "{DOWN}")
									Next
								Elseif lcase(aOutDate(i))="up" Then
									Call Fn_KeyBoardOperation("SendKeys", "{DOWN}")
								End If
								Call Fn_KeyBoardOperation("SendKeys", "{ENTER}")
								wait 1
							End If
						End If
					Next
					Call Fn_Button_Click("Fn_PC_EffectivityOperations", objEffDialog,"OK")
					Fn_PC_EffectivityOperations = True
					Set objEffDialog = Nothing
					Set objEffectivityTable = Nothing
			End  Select		
		'===============================================================================================================================
	End Select	
	Set objEff = Nothing
End Function
'============================================================================================================================================================
'@@
'@@    Function Name			:	Fn_PC_DateControl
'@@
'@@    Description				:	Function Used to set date control
'@@
'@@    Parameters			    :	1. sAction	: Action to be performed
'@@								:	2. sDate	: Date in format ( DD-MMM-YYYY eg. 04-Jan-2012 )
'@@								:	3. sTime 	: Time in 24 hrs format ( HH:MM:SS eg. 21:45:10 )
'@@
'@@    Return Value		   	   	: 	True Or False
'@@
'@@    Pre-requisite			:	PC perspective should be activated.						
'@@
'@@    Examples					:	Call  Fn_PC_DateControl("Set", "02-Jan-2012", "18:30:00")
'@@    Examples					:	Call  Fn_PC_DateControl("Today", "", "")
'@@    Examples					:	Call  Fn_PC_DateControl("Clear", "", "")
'@@
'@@	   History					:	Developer Name				Date		Rev. No.	Changes Done	 Reviewer
'============================================================================================================================================================
'@@									Poonam Chopade			16-Aug-2019		 1.0			Created		 Tc12.3(2019070800)_NewDevelopment_PoonamC_16Aug2019
'============================================================================================================================================================
Public Function Fn_PC_DateControl(sAction, sDate, sTime)
	GBL_FAILED_FUNCTION_NAME="Fn_PC_DateControl"
	Dim objDateCtrl, sDateStr
	Set objDateCtrl = Fn_PC_GetObject("DateControl")
	Fn_PC_DateControl = False
	If Fn_UI_ObjectExist("Fn_PC_DateControl",objDateCtrl) = False Then
        Call Fn_WriteLogFile(Environment.Value("TestLogFile"),  "FAIL : Fn_PC_DateControl : " & objDateCtrl.ToString &" does not Exist of Function Fn_PC_DateControl")
		Exit function
	End If
	Select Case sAction
		Case "Set"
				If sDate <> "" Then
                    sDateStr =  FormatDateTime(cDate(sDate),1)
					sDateStr = trim(mid( sDateStr, instr(sDateStr,",")+1, len(sDateStr)))
					objDateCtrl.JavaEdit("DateEditbox").Set sDateStr
				End If
				If sTime <> "" Then
					objDateCtrl.JavaCalendar("Time").SetTime sTime
				End If
				Call Fn_Button_Click("Fn_PC_DateControl",objDateCtrl,"OK")
				Fn_PC_DateControl = True
		Case "Clear"
				Call Fn_Button_Click("Fn_PC_DateControl",objDateCtrl,"Clear")
				Fn_PC_DateControl = True
		Case "Today"
				Call Fn_Button_Click("Fn_PC_DateControl",objDateCtrl,"Today")
				Fn_PC_DateControl = True
		Case Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Fn_PC_DateControl : Invalid Action [ " & sAction & " ] ")
	End Select
	Set objDateCtrl = Nothing
End Function
'======================================================================================================================================================================
'@@    Function Name		:	Fn_PC_SaveProductConfiguration_Operations
'@@
'@@    Description			:	Function Used to Perform operation on Save Product Configuration  dialog
'@@
'@@    Parameters			:	1. sAction		: Action to be performed
'@@							:	2. dicDetails	: Dictionary object
'@@
'@@    Return Value		   	: 	True Or False
'@@
'@@    Pre-requisite		:	Save Product Configuration dialog should be opened 
'@@
'@@    Examples				:	Set dicSaveDetails = CreateObject( "Scripting.Dictionary")
'@@ 								dicSaveDetails("Name") = "SVR01"
'@@									dicSaveDetails("Description") = "SVR01 description"
'@@									dicSaveDetails("Button") = "OK"
'@@ 							bReturn = Fn_PC_SaveProductConfiguration_Operations("Save",dicSaveDetails)
'@@ 								
'@@-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@	   History				:	Developer Name				Date	  			Rev. No.		Changes Done		Reviewer
'@@-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@	   Created By 			:	Poonam Chopade			18-Oct-2019				1.0		  		 Created		  [TC1123(20190930.00)-10Oct2019-PoonamC-NewDevelopment]
'=========================================================================================================================================================================
Public Function Fn_PC_SaveProductConfiguration_Operations(sAction,dicSaveDetails)

	GBL_FAILED_FUNCTION_NAME = "Fn_PC_SaveProductConfiguration_Operations"
	Dim objSavePC
	
	Fn_PC_SaveProductConfiguration_Operations = False
	Set objSavePC = Fn_PC_GetObject("SaveProductConfiguration")
	
	'Check Existence of Save As window
	If Fn_UI_ObjectExist("Fn_PC_SaveProductConfiguration_Operations",objSavePC) = False Then
		 Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Fail to check existence of [ Save Product Configuration ] window." )
		 Fn_PC_SaveProductConfiguration_Operations = False
		 Set objSavePC = Nothing
		 Exit Function
	End If
	
	Select Case sAction
		Case "Save"
				If dicSaveDetails("Name") <> "" Then  ' Enter Name
					 Call Fn_SISW_UI_JavaEdit_Operations("Fn_PC_SaveProductConfiguration_Operations", "Type", objSavePC, "Name", dicSaveDetails("Name"))
					 Call Fn_ReadyStatusSync(1)
				End If
				
				If dicSaveDetails("Description") <> "" Then 'Enter Description
					 Call Fn_Edit_Box("Fn_PC_SaveProductConfiguration_Operations",objSavePC,"Description",dicSaveDetails("Description"))
					 Call Fn_ReadyStatusSync(1)
				End If
				
				If dicSaveDetails("Button") <> "" Then  'Handle button OK or Cancel
					Call Fn_Button_Click("Fn_PC_SaveProductConfiguration_Operations",objSavePC,dicSaveDetails("Button"))
					Call Fn_ReadyStatusSync(1)
				End If
	End Select
	
	Fn_PC_SaveProductConfiguration_Operations = True
	Set objSavePC = Nothing
	
End Function
'@@=====================================================================================================================================================================
'@@
'@@    Function Name		:	Fn_PC_LoadVariantRule_Operations
'@@
'@@    Description			:	Function Used to Perform operation on Load Variant Rule dialog
'@@
'@@    Parameters			:	1. sAction				: Action to be performed
'@@							:	2. dicLoadVarDetails	: Dictionary object
'@@
'@@    Return Value		   	: 	True Or False
'@@
'@@    Pre-requisite		:	Load Variant Rule dialog should be opened 
'@@
'@@    Examples				:	Set dicLoadVarDetails = CreateObject( "Scripting.Dictionary")
'@@ 								dicLoadVarDetails("VariantRuleNames") = "SVR_CC_Car~SVR_IR_Car~SVR_CC_Utility_Vehicle"
'@@									dicLoadVarDetails("Button") = "Cancel"
'@@ 							bReturn = Fn_PC_LoadVariantRule_Operations("VerifyVariantRules",dicLoadVarDetails)	
'@@
'@@								Set dicLoadVarDetails = CreateObject( "Scripting.Dictionary")
'@@ 								dicLoadVarDetails("SearchCriteria") = "Name:SVR01~ID:0001~Description:TestDescription"
'@@ 							bReturn = Fn_PC_LoadVariantRule_Operations("SearchVariantRules",dicLoadVarDetails)	
'@@
'@@								Set dicLoadVarDetails = CreateObject( "Scripting.Dictionary")
'@@ 								dicLoadVarDetails("Message") = "No search criteria specified. At least one of the search criteria must be provided."
'@@									dicLoadVarDetails("Button") = "Cancel"
'@@ 							bReturn = Fn_PC_LoadVariantRule_Operations("VerifyErrorOnNoSearchCriteria",dicLoadVarDetails)						
'@@								
'@@-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@	   History				:	Developer Name				Date	  			Rev. No.		Changes Done		Reviewer
'@@-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@	   Created By 			:	Poonam Chopade			18-Oct-2019				1.0		  		 Created		  [TC1123(20190930.00)-10Oct2019-PoonamC-NewDevelopment]
'@@=======================================================================================================================================================================
Public Function Fn_PC_LoadVariantRule_Operations(sAction,dicLoadVarDetails)

	GBL_FAILED_FUNCTION_NAME = "Fn_PC_LoadVariantRule_Operations"
	Dim objLoadVarRuledialog,arrVarRules,iRowCnt,iCount,bFlag,iCount1
	Dim sVarRule,sAppMsg,arrSearchCriteria,aSearchValues,sCheckStatus
	
	Fn_PC_LoadVariantRule_Operations = False
	Set objLoadVarRuledialog = Fn_PC_GetObject("LoadVariantRule")
	
	'Check Existence of Load Variant Rule dialog
	If Fn_UI_ObjectExist("Fn_PC_LoadVariantRule_Operations",objLoadVarRuledialog) = False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Fail to check existence of [ Load Variant Rule ] dialog." )
		Fn_PC_LoadVariantRule_Operations = False
		Set objLoadVarRuledialog = Nothing
		Exit Function
	End If
	
	Select Case sAction
		'==============================================================================================================================		
		Case "VerifyVariantRules" 'Verify variant Rule Names
				If dicLoadVarDetails("VariantRuleNames") <> "" Then
					  arrVarRules = Split(dicLoadVarDetails("VariantRuleNames"),"~")
					  iRowCnt = objLoadVarRuledialog.JavaTable("VariantRules").GetROProperty("rows")
					  For iCount = 0 To UBound(arrVarRules)
					  		bFlag = False
					  		For iCount1 = 0 To iRowCnt - 1
					  			sVarRule = Fn_UI_JavaTable_GetCellData("Fn_PC_LoadVariantRule_Operations",objLoadVarRuledialog,"VariantRules",iCount1,"Name")
					  			If trim(sVarRule) = trim(arrVarRules(iCount)) Then
					  				bFlag = True
					  				Exit For
					  			End If
					  		Next
					  		If bFlag = False Then
					  			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Fail to check existence of [ Variant Rule = "&arrVarRules(iCount)&" ].")
					  			Fn_PC_LoadVariantRule_Operations = False
					  			Exit For
					  		Else
					  			Fn_PC_LoadVariantRule_Operations = bFlag
					  		End If
					  Next  
				End If
		'==============================================================================================================================		
		Case "SearchVariantRules" 'Verify variant Rule Names
				If dicLoadVarDetails("SearchCriteria") <> "" Then
					arrSearchCriteria = Split(dicLoadVarDetails("SearchCriteria"),"~")
					 For iCount = 0 To UBound(arrSearchCriteria)
							aSearchValues = Split(arrSearchCriteria(iCount),":")
							objLoadVarRuledialog.JavaEdit("SearchCriteria").SetTOProperty "attached text",aSearchValues(0)+":"
							bFlag = Fn_SISW_UI_JavaEdit_Operations("Fn_PC_LoadVariantRule_Operations", "Set", objLoadVarRuledialog, "SearchCriteria",aSearchValues(1))
							If bFlag = False Then
					  			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Fail to check existence of [ Search field = "&aSearchValues(0)&" ].")
					  			Fn_PC_LoadVariantRule_Operations = False
					  			Exit For
					  		End If
					 Next
					Fn_PC_LoadVariantRule_Operations = Fn_Button_Click("Fn_PC_LoadVariantRule_Operations",objLoadVarRuledialog,"Search")
					Call Fn_ReadyStatusSync(1)
			  End If
		'==============================================================================================================================				  
		Case "VerifyErrorOnNoSearchCriteria"
			Call Fn_Button_Click("Fn_PC_LoadVariantRule_Operations",objLoadVarRuledialog,"Clear")
			Call Fn_ReadyStatusSync(1)
			Call Fn_Button_Click("Fn_PC_LoadVariantRule_Operations",objLoadVarRuledialog,"Search")
			Call Fn_ReadyStatusSync(1)
			If dicLoadVarDetails("Message") <> "" Then				
				sAppMsg = objLoadVarRuledialog.JavaWindow("Search").JavaEdit("Text").GetROProperty("text")
				If instr(1,cstr(sAppMsg),cstr(dicLoadVarDetails("Message"))) > 0 Then
					Fn_PC_LoadVariantRule_Operations = True
				Else
					Fn_PC_LoadVariantRule_Operations = False
				End If
				'Click OK of Variant Rule Dialog
				If Fn_SISW_UI_Object_Operations("Fn_PC_LoadVariantRule_Operations", "Exist", objLoadVarRuledialog.JavaWindow("Search"), "") Then
					Call Fn_Button_Click("Fn_PC_LoadVariantRule_Operations",objLoadVarRuledialog.JavaWindow("Search"),"OK")
					Call Fn_ReadyStatusSync(1)
				End If	
			End if
	'==============================================================================================================================		
	 Case "SelectVariantRules"
	 	  If dicLoadVarDetails("VariantRuleNames") <> "" Then
			  arrVarRules = Split(dicLoadVarDetails("VariantRuleNames"),"~")
			  iRowCnt = objLoadVarRuledialog.JavaTable("VariantRules").GetROProperty("rows")
			  For iCount = 0 To UBound(arrVarRules)
			  		bFlag = False
			  		For iCount1 = 0 To iRowCnt - 1
			  			sVarRule = Fn_UI_JavaTable_GetCellData("Fn_PC_LoadVariantRule_Operations",objLoadVarRuledialog,"VariantRules",iCount1,"Name")
			  			If trim(sVarRule) = trim(arrVarRules(iCount)) Then
			  				bFlag = Fn_UI_JavaTable_ClickCell("Fn_PC_LoadVariantRule_Operations",objLoadVarRuledialog,"VariantRules",iCount1,"Select")
			  				Wait 1
			  				Exit For
			  			End If
			  		Next
			  		If bFlag = False Then
			  			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Fail to select [ Variant Rule = "&arrVarRules(iCount)&" ].")
			  			Fn_PC_LoadVariantRule_Operations = False
			  			Exit For
			  		Else
			  			Fn_PC_LoadVariantRule_Operations = bFlag
			  		End If
			  Next  
		   End If	
	 '==============================================================================================================================	
	 Case "ButtonClick"
		 Fn_PC_LoadVariantRule_Operations = Fn_Button_Click("Fn_PC_LoadVariantRule_Operations",objLoadVarRuledialog,dicLoadVarDetails("Button"))		 
		 Call Fn_ReadyStatusSync(1)	
		 dicLoadVarDetails("Button") = ""
	'==============================================================================================================================		
	Case "VerifyVariantRuleIsSelected"
		  If dicLoadVarDetails("VariantRuleNames") <> "" and dicLoadVarDetails("Status") <> "" Then
			  iRowCnt = objLoadVarRuledialog.JavaTable("VariantRules").GetROProperty("rows")
			  bFlag = False
			  For iCount1 = 0 To iRowCnt - 1
			  		sVarRule = Fn_UI_JavaTable_GetCellData("Fn_PC_LoadVariantRule_Operations",objLoadVarRuledialog,"VariantRules",iCount1,"Name")
			  		If trim(sVarRule) = trim(dicLoadVarDetails("VariantRuleNames")) Then
		  				sCheckStatus = objLoadVarRuledialog.JavaTable("VariantRules").Object.getItem(iCount1).getdata().isSelected()
		  				Select Case dicLoadVarDetails("Status")
		  					Case "Checked"
		  						If instr(sCheckStatus,"true") > 0 Then
		  							bFlag = True
		  						End If
		  					Case "UnChecked"
		  						If instr(sCheckStatus,"false") > 0 Then
		  							bFlag = True
		  						End If
		  					Case else
		  						bFlag = False
		  				End Select
						If bFlag = False Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Fail to verify state as selected for [ Variant Rule = "&dicLoadVarDetails("VariantRuleNames")&" ].")
							Fn_PC_LoadVariantRule_Operations = False
							Exit For
						Else
							Fn_PC_LoadVariantRule_Operations = bFlag
						End If
					 End If
			  	Next
		   End If
		'==============================================================================================================================	
		Case "CheckboxOperation"
			If dicLoadVarDetails("Appendonlyexpressions") <> "" Then
				bFlag = Fn_SISW_UI_JavaCheckBox_Operations("Fn_PC_LoadVariantRule_Operations", "Set", objLoadVarRuledialog, "Append only expressions", dicLoadVarDetails("Appendonlyexpressions"))
			End If
			If dicLoadVarDetails("Loadexpandedexpression") <> "" Then
				bFlag = Fn_SISW_UI_JavaCheckBox_Operations("Fn_PC_LoadVariantRule_Operations", "Set", objLoadVarRuledialog, "Load expanded expression", dicLoadVarDetails("Loadexpandedexpression"))
			End If
			Fn_PC_LoadVariantRule_Operations = bFlag
		'==============================================================================================================================
		Case "FilterVariantRule"
			Call Fn_Button_Click("Fn_PC_LoadVariantRule_Operations",objLoadVarRuledialog,"ClearFilterText")
			Wait 1
			If dicLoadVarDetails("FilterText") <> "" Then
				Fn_PC_LoadVariantRule_Operations = Fn_Edit_Box("Fn_PC_LoadVariantRule_Operations",objLoadVarRuledialog,"FilterEditBox",dicLoadVarDetails("FilterText")) 
			End If
	End Select
	
	If dicLoadVarDetails("Button") <> "" Then  'Click on Buttons
		 Call Fn_Button_Click("Fn_PC_LoadVariantRule_Operations",objLoadVarRuledialog,dicLoadVarDetails("Button"))	
		 Call Fn_ReadyStatusSync(1)	
	End If 
	
	Set objLoadVarRuledialog = Nothing
	
End Function
'@@    Function Name		:	Fn_PC_FreeFromRule_Operations
'@@
'@@    Description			:	Function Used to Perform Free-form expression opertion in editor
'@@
'@@    Parameters			:	1. sAction		: Action to be performed
'@@							:	2. dicDetails	: Dictionary object
'@@
'@@    Return Value		   	: 	True Or False
'@@
'@@    Pre-requisite		:	Need to Create CC and rules in Configuration rules.
'@@
'@@    Examples				:	Set dicDetails = CreateObject( "Scripting.Dictionary")
'@@								dicDetails("Expression") = "(assert(=>(= |[Teamcenter]10_"&iRanNo&"| false)(and(= |[Teamcenter]50_"&iRanNo&"| false)(= |[Teamcenter]60_"&iRanNo&"| false))))"
'@@								dicDetails("Button")= "Save"
'@@								bReturn = Fn_PC_FreeFromRule_Operations("SetExpression",dicDetails)
'@@ 								
'@@-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@	   History				:	Developer Name				Date	  			Rev. No.		Changes Done		Reviewer
'@@-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@	   Created By 			:	Gaurav Kasar		12-March-2021				1.0		  		 Created		  
'=========================================================================================================================================================================
Public Function Fn_PC_FreeFromRule_Operations(sAction,dicLoadVarDetails)
	GBL_FAILED_FUNCTION_NAME = "Fn_PC_FreeFromRule_Operations"
	Dim objConfigRuleTable,ObjFreeFormEditor,sMenu
	
	Fn_PC_FreeFromRule_Operations = False
	
	Set objConfigRuleTable = Fn_PC_GetObject("ConfiguratorRules")
    Set ObjFreeFormEditor = Fn_PC_GetObject("Free-formValueInput")
    	If Fn_ToolBarOperation("IsSelected","Navigation Pane","") Then
		Call Fn_ToolBarOperation("Click","Navigation Pane","")
	End If
	On Error Resume Next
	If Fn_UI_ObjectExist("Fn_PC_FreeFromRule_Operations",objConfigRuleTable) = True Then
		sMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("Configurator_Toolbar"),"OpenFreeFormExpressionEditor")
		Call Fn_ToolBarOperation("Click",sMenu,"")
		Call Fn_ReadyStatusSync(1)
		ObjFreeFormEditor.SetTOProperty "title","Free-form Expression Editor"
		If Fn_UI_ObjectExist("Fn_PC_FreeFromRule_Operations",ObjFreeFormEditor) = False Then ' Check Existence of Configurator Rules Table
			Fn_PC_FreeFromRule_Operations = False
			Set objConfigRuleTable = Nothing
			Set ObjFreeFormEditor = Nothing
			Exit Function
		End  If
	End If
	Err.Clear

Select Case sAction
 	   Case "SetExpression"
			ObjFreeFormEditor.JavaEdit("FromValue").SetTOProperty "toolkit class","org.eclipse.swt.custom.StyledText"
			ObjFreeFormEditor.JavaEdit("FromValue").SetTOProperty "attached text",""
			
			If Fn_UI_ObjectExist("Fn_PC_FreeFromRule_Operations",ObjFreeFormEditor.JavaEdit("FromValue")) = True Then
			    ObjFreeFormEditor.JavaButton("OK").SetTOProperty "label","Clear"
			    ObjFreeFormEditor.JavaButton("OK").Click
			    ObjFreeFormEditor.JavaEdit("FromValue").set dicLoadVarDetails("Expression")
			    Call Fn_ReadyStatusSync(1)
			    wait 3
			   If Err.number = 0 Then
			   	Fn_PC_FreeFromRule_Operations = True
			   End If	   
			End if
			If dicLoadVarDetails("Button") <> "" Then
				ObjFreeFormEditor.JavaButton("OK").SetTOProperty "label", dicLoadVarDetails("Button")
				ObjFreeFormEditor.JavaButton("OK").Click
				Call Fn_ReadyStatusSync(1)
				wait 3
			    If Err.number < 0 Then
			   	Fn_PC_FreeFromRule_Operations = False
			   	Else
			   	Fn_PC_FreeFromRule_Operations = True
			   	Set objConfigRuleTable = Nothing
				Set ObjFreeFormEditor = Nothing
			   	Exit Function
			    End If
			End If
End Select
	If Fn_ToolBarOperation("IsSelected","Navigation Pane","") = False Then
		Call Fn_ToolBarOperation("Click","Navigation Pane","")
	End If
Set objConfigRuleTable = Nothing
Set ObjFreeFormEditor = Nothing
End Function


'@@    Function Name		:	Fn_PC_Open_Context_Independent_Search_View_Operations
'@@
'@@    Description			:	Function Used to search thr rules in Context Independent search view in Product Configurator
'@@
'@@    Parameters			:	1. dicDetails	: Dictionary object
'@@						
'@@
'@@    Return Value		   	: 	True Or False
'@@
'@@    Pre-requisite		:	Need to Create rules in Configuration rules.
'@@
'@@    Examples				:	Set dicDetails = CreateObject( "Scripting.Dictionary")
'@@								dicDetails("Tab")="Context Independent Search"
'@@								dicDetails("ID") = "*"
'@@								dicDetails("Search Options")= "Free-form Rule"
'@@								bReturn = Fn_PC_FreeFromRule_Operations("SetExpression",dicDetails)
'@@ 								
'@@-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@	   History				:	Developer Name				Date	  			Rev. No.		Changes Done		Reviewer
'@@-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@	   Created By 			:	Gaurav Kasar		16-March-2021				1.0		  		 Created		  
'=========================================================================================================================================================================
Public Function Fn_PC_Open_Context_Independent_Search_View_Operations(dicLoadVarDetails)
	GBL_FAILED_FUNCTION_NAME = "Fn_PC_Open_Context_Independent_Search_View_Operations"
	Dim ObjProductConfigurator,sMenu
	
	Fn_PC_Open_Context_Independent_Search_View_Operations = False
	
    Set ObjProductConfigurator = Fn_PC_GetObject("ProductConfigurator")
	On Error Resume Next
	If Fn_TabFolder_Operation("Exist",dicLoadVarDetails("Tab"),"") = False Then
		sMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("Configurator_Toolbar"),"OpenContextIndependentSearchView")
		Call Fn_ToolBarOperation("Click",sMenu,"")
		Call Fn_ReadyStatusSync(1)
		If Fn_TabFolder_Operation("Exist",dicLoadVarDetails("Tab"),"") = False Then ' Check Existence of Configurator Rules Table
			Fn_PC_Open_Context_Independent_Search_View_Operations = False
			Set objContextIndSearch = Nothing
			Set ObjProductConfigurator = Nothing
			Exit Function
		End  If
	End If
	Err.Clear
	ObjProductConfigurator.JavaButton("Search").SetTOProperty "label","Clear"
	ObjProductConfigurator.JavaButton("Search").Click
	wait 1
	If dicLoadVarDetails("ID")<>"" Then
		ObjProductConfigurator.JavaStaticText("IDLabel").SetTOProperty "label","ID:"
		ObjProductConfigurator.JavaEdit("IDEditbox").Set dicLoadVarDetails("ID")
		wait 1
	End If
	If dicLoadVarDetails("Configurator Context")<>"" Then
		ObjProductConfigurator.JavaStaticText("IDLabel").SetTOProperty "label","Configurator Context:"
		ObjProductConfigurator.JavaEdit("IDEditbox").Set dicLoadVarDetails("Configurator Context")
		wait 1
	End If
	If dicLoadVarDetails("Feature")<>"" Then
		ObjProductConfigurator.JavaStaticText("IDLabel").SetTOProperty "label","Feature:"
		ObjProductConfigurator.JavaEdit("IDEditbox").Set dicLoadVarDetails("Feature")
		wait 1
	End If
	
	If dicLoadVarDetails("Search Options")<>"" Then
		ObjProductConfigurator.JavaList("NavTreeTableList").SetTOProperty "attached text","Search Options"
		ObjProductConfigurator.JavaList("NavTreeTableList").Select dicLoadVarDetails("Search Options")
		wait 1
	End If
	ObjProductConfigurator.JavaButton("Search").SetTOProperty "label","Search"
	ObjProductConfigurator.JavaButton("Search").Click
	wait 2
	Call Fn_ReadyStatusSync(1)
	If Err.Number < 0 Then
	    Fn_PC_Open_Context_Independent_Search_View_Operations = False
		Set objContextIndSearch = Nothing
		Set ObjProductConfigurator = Nothing
		Exit Function
	Else
		Fn_PC_Open_Context_Independent_Search_View_Operations = True
	End If
Set ObjProductConfigurator = Nothing
End Function


'@@    Function Name		:	Fn_PC_Save_Variant_Rule
'@@
'@@    Description			:	Function Used to Save the variant Rule.
'@@
'@@    Parameters			:	1. dicDetails	: Dictionary object
'@@						
'@@
'@@    Return Value		   	: 	True Or False
'@@
'@@    Pre-requisite		:	Need to Create rules in Configuration rules.
'@@
'@@    Examples				:	Set dicDetails = CreateObject( "Scripting.Dictionary")
'@@ 								dicDetails("SVR_Name") = DataTable("SVR_Name", dtGlobalSheet)+"_"+iRanNo
'@@ 							bReturn = Fn_PC_Save_Variant_Rule(dicDetails)
'@@ 								
'@@-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@	   History				:	Developer Name				Date	  			Rev. No.		Changes Done		Reviewer
'@@-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@	   Created By 			:	Gaurav Kasar		1-April-2021				1.0		  		 Created		  
'=========================================================================================================================================================================
Public Function Fn_PC_Save_Variant_Rule(dicLoadVarDetails)

Dim ObjSaveVarWindow

  Fn_PCA_Save_Variant_Rule = False 
    set ObjSaveVarWindow =JavaWindow("ProductConfigurator").JavaWindow("Set Rule Date")
    ObjSaveVarWindow.SetTOProperty "title","Save Variant Rule"
    Err.clear
    If Fn_UI_ObjectExist("Fn_PC_VariantConfigurationView_Operation",ObjSaveVarWindow) Then
      ObjSaveVarWindow.JavaEdit("Date").SetTOProperty "attached text","Name *"
      wait 2
       ObjSaveVarWindow.JavaEdit("Date").Type dicLoadVarDetails("SVR_Name")
       If Err.number < 0 Then
			Fn_PC_Save_Variant_Rule = False 
       Else
            sAppMsg = Fn_UI_Object_GetROProperty("Fn_PC_VariantConfigurationView_Operation",ObjSaveVarWindow.JavaEdit("Date"),"text")
			ObjSaveVarWindow.JavaButton("Button").SetTOProperty "label","OK"
			wait 1
			ObjSaveVarWindow.JavaButton("Button").Click
			Fn_PC_Save_Variant_Rule = sAppMsg
       End If
	 Else
	 	Fn_PC_Save_Variant_Rule = False
	 	Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Save variant Rule Dialog  does not exists")		
	 End If	  			 
Set ObjSaveVarWindow = Nothing
End Function

'-------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@    Function Name		:	Fn_PC_VariantConfigurationView_OptionsVerify
'@@
'@@    Description			:	Function Used to verify the option values in VEE table
'@@
'@@    Parameters			:	1. dicOptions	: Dictionary object
'@@						
'@@
'@@    Return Value		   	: 	True Or False
'@@
'@@   
'@@
'@@    Examples				:	Set dicOptions = CreateObject( "Scripting.Dictionary")
'@@ 								dicOptions("values")=DataTable("GFO_Name", dtGlobalSheet)
'@@									bReturn = Fn_PC_VariantConfigurationView_OptionsVerify(dicOptions)
'@@ 								
'@@-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@	   History				:	Developer Name				Date	  			Rev. No.		Changes Done		Reviewer
'@@-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@	   Created By 			:	Vsishali D	           7-April-2022				1.0		  		 Created		  
'=========================================================================================================================================================================

Function Fn_PC_VariantConfigurationView_OptionsVerify(dicOptions)

bFlag = False
Set objPCWindow = Fn_PC_GetObject("ProductConfigurator")
	 iRowsCount = objPCWindow.JavaObject("VariantExpressionEditor").Object.getRowCount() 

	arroptions = split(dicOptions("values"),"~")
	
            For iCount = 0 To ubound(arroptions) Step 1
            	
            	 For iCnt = 1 To iRowsCount - 1
				 
				 	'Get Index for Subject & Applicability sections
					If instr(objPCWindow.JavaObject("VariantExpressionEditor").Object.getCellByPosition(0,iCnt).getDataValue().getData().toString(),arroptions(iCount)) > 0 Then
						
						bFlag = True
						Exit For
					ElseIf iCnt = iRowsCount-1 and bFlag = False Then
						Exit function
						Fn_PC_VariantConfigurationView_OptionsVerify = False
					End IF
					
				 Next
			Next
			If bFlag = False Then
	 			Fn_PC_VariantConfigurationView_OptionsVerify = False
	 			Exit function
	 		Else
	 			Fn_PC_VariantConfigurationView_OptionsVerify = True
			End If
Set objPCWindow = Nothing			
End Function

