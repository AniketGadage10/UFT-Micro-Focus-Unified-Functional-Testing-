Option Explicit

'*************************** Function List ***************************************************************************************************************************************
'1. Fn_RsrcMgr_GetObject(sObjectName)
'2. Fn_RsrcMgr_NewResourceOperation(sAction,dicResourceDetails,sButton)
'3. Fn_RsrcMgr_NavTreeTableOperation(sAction,sNodeName,dicDetails,sReserve)
'4. Fn_RsrcMgr_ClassifyResourceObject(sAction,dicClassifyDetails,sReserve,sButton)
'5. Fn_RsrcMgr_AttributeValues(sAction,sAttributeType,sValue,sDetails,sReserve)
'*************************** Function List ***************************************************************************************************************************************

'*************************** Function to get Object hierarchy ********************************************************************************************************************
'
'Function Name		 	:	Fn_RsrcMgr_GetObject
'
'Description		    :  	Function to get specified Object hierarchy.

'Parameters		    	:	1. sObjectName : Object name
								
'Return Value		    :  	Object \ Nothing
'12
'Examples		     	:	Fn_RsrcMgr_GetObject("NewResourceDialog")

'History:
'		Developer Name		Date	   Rev. No.	    Reviewer			Changes Done	
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'		Vivek Ahirrao	 10-09-2015		1.0			Vivek Ahirrao		Created
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_RsrcMgr_GetObject(sObjectName)
	Dim sObjectXMLPath
	sObjectXMLPath = Fn_GetEnvValue("User", "AutomationDir") & "\TestData\AutomationXML\ObjectXML\ResourceManager.xml"
	Set Fn_RsrcMgr_GetObject = Fn_SISW_Setup_GetObjectFromXML(sObjectXMLPath, sObjectName)
End Function

'*************************** Function to perform operation on New Resource Dialog ********************************************************************************************************************
'
'Function Name		 	:	Fn_RsrcMgr_NewResourceOperation
'
'Description		    :  	Function to perform operation on New Resource Dialog
'
'Parameters		    	:	1. sAction 				: "Create"
'							2. dicResourceDetails 	: Dictionary Object
'							3. sButton 				: "OK"/Else
'
'Return Value		    :  	True/False  OR  ResourceName/RevisionID
'
'Prerequistes			:  	Resource Manager perspective should be opened
'
'Examples		     	:	Set dicResourceDetails = CreateObject("Scripting.Dictionary")
'								dicResourceDetails("ResourceType") = "Resource"
'								dicResourceDetails("ResourceID") = ""
'								dicResourceDetails("RevisionID") = ""
'								dicResourceDetails("ResourceName") = "ResourceTest"
'							bReturn = Fn_RsrcMgr_NewResourceOperation("Create",dicResourceDetails,"OK")
'
'History:
'		Developer Name		Date	   Rev. No.	    Reviewer			Changes Done	
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'		Vivek Ahirrao	 10-09-2015		1.0			Vivek Ahirrao		Created
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_RsrcMgr_NewResourceOperation(sAction,dicResourceDetails,sButton)
	GBL_FAILED_FUNCTION_NAME="Fn_RsrcMgr_NewResourceOperation"
	Dim objNewResourceDialog, bFlag, sRID, sRevID
	Fn_RsrcMgr_NewResourceOperation = False
	bFlag = False
	Set objNewResourceDialog = Fn_RsrcMgr_GetObject("NewResourceDialog")
	
	'Select Menu [File -> New -> Resource...]
	If Fn_UI_ObjectExist("Fn_RsrcMgr_NewResourceOperation",objNewResourceDialog)=False Then
		Call Fn_MenuOperation("Select","File:New:Resource...")
		Call Fn_ReadyStatusSync(3)
	End If
	Select Case sAction
		Case "Create"
				'Select Resource Type in list
				If dicResourceDetails("ResourceType")<>"" Then
					bFlag = Fn_List_Select("Fn_RsrcMgr_NewResourceOperation",objNewResourceDialog,"ResourceType",dicResourceDetails("ResourceType"))
					If bFlag = False Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to select Resource Type from List.")
						Fn_RsrcMgr_NewResourceOperation = False
						Set objNewResourceDialog = Nothing
						Exit Function
					End If	
				End If
				'Set ResourceId and RevisionId or Click Assign button
				If dicResourceDetails("ResourceID")<>"" AND dicResourceDetails("RevisionID")<>"" Then
					Call Fn_Edit_Box("Fn_RsrcMgr_NewResourceOperation",objNewResourceDialog,"ResourceID",dicResourceDetails("ResourceID"))
					Wait 1
					Call Fn_Edit_Box("Fn_RsrcMgr_NewResourceOperation",objNewResourceDialog,"RevisionID",dicResourceDetails("RevisionID"))
					Wait 1
				Else
					Call Fn_Button_Click("Fn_RsrcMgr_NewResourceOperation",objNewResourceDialog,"Assign")
					Wait 1
					sRID = Fn_Edit_Box_GetValue("Fn_RsrcMgr_NewResourceOperation",objNewResourceDialog,"ResourceID")
					sRevID = Fn_Edit_Box_GetValue("Fn_RsrcMgr_NewResourceOperation",objNewResourceDialog,"RevisionID")
				End If
				'Set ResourceName
				If dicResourceDetails("ResourceName")<>"" Then
					Call Fn_Edit_Box("Fn_RsrcMgr_NewResourceOperation",objNewResourceDialog,"ResourceName",dicResourceDetails("ResourceName"))
					Wait 1
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed du to Resource Name is blank.")
					Fn_RsrcMgr_NewResourceOperation = False
					Set objNewResourceDialog = Nothing
					Exit Function
				End If
				'Click OK button
				If sButton<>"" Then
					Call Fn_Button_Click("Fn_RsrcMgr_NewResourceOperation",objNewResourceDialog,sButton)
				Else
					Call Fn_Button_Click("Fn_RsrcMgr_NewResourceOperation",objNewResourceDialog,"OK")
				End If
				'Return True or "ResourceID/ResourceName"
				If dicResourceDetails("ResourceID")<>"" AND dicResourceDetails("RevisionID")<>"" Then
					Fn_RsrcMgr_NewResourceOperation = True
				Else
					Fn_RsrcMgr_NewResourceOperation = sRID + "/" + sRevID
				End If
		Case Else
				'Future Use		
	End Select
	Set objNewResourceDialog = Nothing
End Function

'*************************** Function to perform operation on NavTreeTable in Resource Manager ********************************************************************************************************************
'
'Function Name		 	:	Fn_RsrcMgr_NavTreeTableOperation
'
'Description		    :  	Function to perform operation on NavTreeTable in Resource Manager
'
'Parameters		    	:	1. sAction 		: "Create"
'							2. sNodeName 	: Node Name
'							3. dicDetails 	: Future use
'							4. sReserve 	: Future use
'
'Return Value		    :  	True/False
'
'Prerequistes			:  	Resource Manager perspective should be opened and NavTreeTable sould be opened
'
'Examples		     	:	bReturn = Fn_RsrcMgr_NavTreeTableOperation("Select", "000068/A;1-TestTest (View)", "", "")
'							bReturn = Fn_RsrcMgr_NavTreeTableOperation("Expand", "000068/A;1-TestTest (View)", "", "")
'							bReturn = Fn_RsrcMgr_NavTreeTableOperation("Exist", "000068/A;1-TestTest (View)", "", "")
'							bReturn = Fn_RsrcMgr_NavTreeTableOperation("PopupMenuSelect", "000068/A;1-TestTest (View)", "", "Send To:My Teamcenter")
'
'History:
'		Developer Name		Date	   Rev. No.	    Reviewer			Changes Done	
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'		Vivek Ahirrao	 10-09-2015		1.0			Vivek Ahirrao		Created
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_RsrcMgr_NavTreeTableOperation(sAction,sNodeName,dicDetails,sMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_RsrcMgr_NavTreeTableOperation"
	Dim objNevTreeTable, bFlag, objNevTree, iRowNo
	
	Fn_RsrcMgr_NavTreeTableOperation = False
	
	Set objNevTreeTable = Fn_RsrcMgr_GetObject("NavTreeTable")
	'Check Existance
	If Fn_UI_ObjectExist("Fn_RsrcMgr_NavTreeTableOperation",objNevTreeTable)=False Then
		Fn_RsrcMgr_NavTreeTableOperation = False
		Set objNevTreeTable = Nothing
		Exit Function
	End If
	
	Select Case sAction
		'Case to select node -----
		Case "Select"
				bFlag = Fn_SISW_UI_JavaTable_Operations("Fn_RsrcMgr_NavTreeTableOperation", "SelectRow", objNevTreeTable , "", "GetValueAt", "BOM Line", sNodeName, "", "", "", "")
				If bFlag = False Then
					Fn_RsrcMgr_NavTreeTableOperation = False
					Set objNevTreeTable = Nothing
					Exit Function
				Else 
					Fn_RsrcMgr_NavTreeTableOperation = True
					Set objNevTreeTable = Nothing
				End If
		Case "Expand"
				bFlag = Fn_SISW_UI_JavaTable_Operations("Fn_RsrcMgr_NavTreeTableOperation", "DoubleClickCell", objNevTreeTable , "", "GetValueAt", "BOM Line", sNodeName, "", "", "", "")
				If bFlag = False Then
					Fn_RsrcMgr_NavTreeTableOperation = False
					Set objNevTreeTable = Nothing
					Exit Function
				Else 
					Fn_RsrcMgr_NavTreeTableOperation = True
					Set objNevTreeTable = Nothing
				End If
		Case "Exist"  
				bFlag = Fn_SISW_UI_JavaTable_Operations("Fn_RsrcMgr_NavTreeTableOperation", "Exist", objNevTreeTable , "", "GetValueAt", "BOM Line", sNodeName, "", "", "", "")
				If bFlag = False Then
					Fn_RsrcMgr_NavTreeTableOperation = False
					Set objNevTreeTable = Nothing
					Exit Function
				Else 
					Fn_RsrcMgr_NavTreeTableOperation = True
					Set objNevTreeTable = Nothing
				End If
		Case "PopupMenuSelect"		'Pre-requisite = Row should be selected
				Set objNevTree = Fn_RsrcMgr_GetObject("ResourceMgrApplet")
				iRowNo = objNevTreeTable.Object.getSelectedRow()
				If iRowNo <> -1 Then
					Call Fn_UI_JavaTable_CellRightClick("Fn_RsrcMgr_NavTreeTableOperation",objNevTree,"NavTreeTable",iRowNo,"BOM Line","RIGHT","")
					wait 1
					Call Fn_UI_JavaMenu_Select("Fn_RsrcMgr_NavTreeTableOperation",JavaWindow("ResourceManagerWindow"),sMenu)
					Fn_RsrcMgr_NavTreeTableOperation=True
					Set objNevTreeTable = Nothing
					Set objNevTree = Nothing
				Else
					Fn_RsrcMgr_NavTreeTableOperation = False
					Set objNevTreeTable = Nothing
					Set objNevTree = Nothing
					Exit Function
				End if
		Case Else
				'Future Use
	End Select
End Function

'*************************** Function to perform operation on New Resource Dialog ********************************************************************************************************************
'
'Function Name		 	:	Fn_RsrcMgr_ClassifyResourceObject
'
'Description		    :  	Function to Classify the Resource object by Toolbar or by clicking Not Classified link or else
'
'Parameters		    	:	1. sAction 				: "Create"
'							2. dicClassifyDetails 	: Dictionary Object
'							3. sReserve 			: Future Use
'							4. sButton				: "Save"/else any other button
'
'Return Value		    :  	True/False
'
'Prerequistes			:  	Resource Manager perspective should be opened and a Resource object should be selected
'
'Examples		     	:	Set dicClassifyDetails = CreateObject("Scripting.Dictionary")
'								dicClassifyDetails("ClassifyObjectWith") = "NotClassifiedLink"
'								dicClassifyDetails("ClassPath") = "Classification Root:StorageClass_10086  [1]"
'							bReturn = Fn_RsrcMgr_ClassifyResourceObject("ClassifyObject",dicClassifyDetails,sReserve,"Save current Resource")
'
'History:
'		Developer Name		Date	   Rev. No.	    Reviewer			Changes Done	
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'		Vivek Ahirrao	 14-09-2015		1.0			Vivek Ahirrao		Created
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_RsrcMgr_ClassifyResourceObject(sAction,dicClassifyDetails,sReserve,sButton)
	GBL_FAILED_FUNCTION_NAME="Fn_RsrcMgr_ClassifyResourceObject"
	Dim bFlag, aClassPath, sClassPath, iCount
	Dim objApplet, objClsInfoDialog
	
	Fn_RsrcMgr_ClassifyResourceObject = False
	bFlag = False
	
	Select Case sAction
		Case "ClassifyObject"
				Select Case dicClassifyDetails("ClassifyObjectWith")
					Case "NotClassifiedLink"
							Set objApplet = Fn_RsrcMgr_GetObject("ResourceMgrApplet1")
							Set objClsInfoDialog = Fn_RsrcMgr_GetObject("ClassificationInformationDialog")
							
							'Select Classification Properties tab
							bFlag = Fn_TabFolder_Operation("Select", "Classification Properties", "")
							If bFlag = False Then
								Set objApplet = Nothing
								Set objClsInfoDialog = Nothing
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed To Select Classification Properties tab")
								Fn_RsrcMgr_ClassifyResourceObject = False
								Exit Function
							End If
							'Click on 'Not Classified' link
							Call Fn_UI_JavaStaticText_Click("Fn_RsrcMgr_ClassifyResourceObject", objApplet, "NotClassifiedLink", 2, 2, "LEFT")
							Call Fn_ReadyStatusSync(1)
							
							'Select Class and Activate it
							If dicClassifyDetails("ClassPath")<>"" AND objClsInfoDialog.Exist(2) Then
								aClassPath = Split(dicClassifyDetails("ClassPath") , ":")
								For iCount = 0 To Ubound(aClassPath)- 1 Step 1
									If iCount = 0 Then
										sClassPath = aClassPath(iCount)
									Else
										sClassPath = sClassPath &":" & aClassPath(iCount)
									End If
									If Fn_UI_JavaTree_Expand("Fn_RsrcMgr_ClassifyResourceObject", objClsInfoDialog, "Hierarchy",sClassPath)= False Then
										Set objApplet = Nothing
										Set objClsInfoDialog = Nothing
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed To EXPAND Class in Hierarchy Tree")
										Fn_RsrcMgr_ClassifyResourceObject = False
										Exit Function	
									End If
								Next
								If Fn_JavaTree_Node_Activate("Fn_RsrcMgr_ClassifyResourceObject",objClsInfoDialog,"Hierarchy",dicClassifyDetails("ClassPath")) = False then
									Set objApplet = Nothing
									Set objClsInfoDialog = Nothing
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Class is not found in Hierarchy Tree")
									Fn_RsrcMgr_ClassifyResourceObject = False
									Exit Function
								End if
							End If
							
							'Save Classified object
							If sButton<>"" Then
								Call Fn_ToolBarOperation("Click", sButton, "")
								Call Fn_ReadyStatusSync(5)
							End If
							
							Set objApplet = Nothing
							Set objClsInfoDialog = Nothing
							Fn_RsrcMgr_ClassifyResourceObject = True
					Case "Toolbar"
							'Future Use
					Case Else
							'Future Use
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed To Select Case [ "+dicClassifyDetails("ClassifyObjectWith")+" ]")
							Fn_RsrcMgr_ClassifyResourceObject = False
							Exit Function	
				End Select
		Case Else
				'Future Use
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed To Select Case [ "+sAction+" ]")
				Fn_RsrcMgr_ClassifyResourceObject = False
				Exit Function
	End Select
End Function

'*************************** Function to perform operation on Attribut Values ********************************************************************************************************************
'
'Function Name		 	:	Fn_RsrcMgr_AttributeValues
'
'Description		    :  	Function to Classify the Resource object by Toolbar or by clicking Not Classified link or else
'
'Parameters		    	:	1. sAction 			: "Add","Verify","VerifyValue"
'							2. sAttributeType 	: Dictionary Object
'							3. sValue 			: Future Use
'							4. sDetails			: "Save"/else any other button
'							5. sReserve			:
'
'Return Value		    :  	True/False
'
'Prerequistes			:  	Attribute Names and Values should be displayed
'
'Examples		     	:	Call Fn_RsrcMgr_AttributeValues("Add","Text","IntegerAttrib:7","","")
'							Call Fn_RsrcMgr_AttributeValues("VerifyValue","List","State:MH~Country:India","","")
'
'History:
'		Developer Name		Date	   Rev. No.	    Reviewer			Changes Done	
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'		Vivek Ahirrao	 14-09-2015		1.0			Vivek Ahirrao		Created
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_RsrcMgr_AttributeValues(sAction,sAttributeType,sValue,sDetails,sReserve)
	GBL_FAILED_FUNCTION_NAME="Fn_RsrcMgr_AttributeValues"
	Dim bFlag, aValue, iCount, aSetValues, aGetValues, sTextVal
	Dim objApplet
	
	Fn_RsrcMgr_AttributeValues = False
	bFlag = False
	
	Set objApplet = Fn_RsrcMgr_GetObject("ResourceMgrApplet1")
	'Select Classification Properties tab
	bFlag = Fn_TabFolder_Operation("Select", "Classification Properties", "")
	If bFlag = False Then
		Set objApplet = Nothing
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed To Select Classification Properties tab")
		Fn_RsrcMgr_AttributeValues = False
		Exit Function
	End If
	
	Select Case sAction
		Case "Add"
				Select Case sAttributeType
					Case "Text"
							If Instr(1,sValue,"~") Then
								aValue = Split(sValue,"~",-1,1)
							Else
								aValue = Array(sValue)
							End If
							For iCount = 0 to Ubound(aValue)
								aSetValues = Split(aValue(iCount),":",-1,1)
								objApplet.JavaStaticText("AttributeName_Label").SetTOProperty "label", aSetValues(0)
								Call Fn_Edit_Box("Fn_RsrcMgr_AttributeValues",objApplet,"AttributeValue_Edit", aSetValues(1))													 
							    Wait 3
								If Err.Number < 0 Then
									Fn_RsrcMgr_AttributeValues = False
									Set objApplet = Nothing
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Set value [" + aSetValues(1) + "] for Attribute ["+aSetValues(0)+"]" ) 
									Exit Function
								Else
									Fn_RsrcMgr_AttributeValues = True
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Set value [" + aSetValues(1) + "] for Attribute ["+aSetValues(0)+"]" ) 				
								End If
						  	Next
					Case "List"
							If Instr(1,sValue,"~") Then
								aValue = Split(sValue,"~",-1,1)
							Else
								aValue = Array(sValue)
							End If
							
							For iCount = 0 to Ubound(aValue)
								bFlag = False
								aSetValues = Split(aValue(iCount),":",-1,1)
								objApplet.JavaStaticText("AttributeName_Label").SetTOProperty "label", aSetValues(0)
								Wait 1
								bFlag = Fn_List_Select("Fn_RsrcMgr_AttributeValues",objApplet,"AttributeValueList",aSetValues(1))
								wait 1
								If bFlag = False Then
									Fn_RsrcMgr_AttributeValues = False
									Set objApplet = Nothing
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Set value [" + aSetValues(1) + "] for Attribute [" + aSetValues(0) + "] " )
									Exit Function
								End If
						  	Next
						  	
						  	If bFlag = True Then
						  		Set objApplet = Nothing
						  		Fn_RsrcMgr_AttributeValues = True
						  	End If
					Case "Date"
							'Future Use
				End Select
		Case "Verify"
				Select Case sAttributeType
					Case "Text"
							'Future Use
					Case "List"
							'Future Use
					Case "Date"
							'Future Use
				End Select
		Case "VerifyValue"
				Select Case sAttributeType
					Case "Text"
							If Instr(1,sValue,",") Then
								aValue = Split(sValue,",",-1,1)
							Else
								aValue = Array(sValue)
							End If

							For iCount = 0 to Ubound(aValue)													 
								aGetValues = split(aValue(iCount),":",-1,1)
								objApplet.JavaStaticText("AttributeName_Label").SetTOProperty "label", aGetValues(0)
								sTextVal = objApplet.JavaEdit("AttributeValue_Edit").GetROProperty("value")
								Wait 1
								If Trim(CStr(aGetValues(1))) = Trim(CStr(sTextVal))Then
									Fn_RsrcMgr_AttributeValues = True
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Verify value [" + aValue(iCount) + "] for Attribute " )                 
								Else															
									Fn_RsrcMgr_AttributeValues = False
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Verify value [" + aValue(iCount) + "] for Attribute " ) 
									Set objApplet = Nothing
									Exit Function
								End If
							  Next
					Case "List"
							If Instr(1,sValue,",") Then
								aValue = Split(sValue,"~",-1,1)
							Else
								aValue = Array(sValue)
							End If
							
							For iCount = 0 to Ubound(aValue)
								bFlag = False
								aGetValues = split(aValue(iCount),":",-1,1)
								objApplet.JavaStaticText("AttributeName_Label").SetTOProperty "label", aGetValues(0)
								Wait 1
'								sTextVal = objApplet.JavaList("AttributeValueList").GetROProperty("value")
								sTextVal = Fn_UI_Object_GetROProperty("Fn_RsrcMgr_AttributeValues",objApplet.JavaList("AttributeValueList"), "value")
								Wait 1
								If Trim(Cstr(aGetValues(1))) = Trim(Cstr(sTextVal)) Then
									bFlag = True
								Else															
									Fn_RsrcMgr_AttributeValues = False
									Set objApplet = Nothing
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Verify value [" + aGetValues(1) + "] for Attribute [" + aGetValues(0) + "] " )
									Exit Function
								End If
							Next
							
							If bFlag = True Then
								Set objApplet = Nothing
								Fn_RsrcMgr_AttributeValues = True
							End If
					Case "ListItemExist"
							If Instr(1,sValue,",") Then
								aValue = Split(sValue,"~",-1,1)
							Else
								aValue = Array(sValue)
							End If
							
							For iCount = 0 To Ubound(aValue)
								aGetValues = Split(aValue(iCount),":",-1,1)
								objApplet.JavaStaticText("AttributeName_Label").SetTOProperty "label", aGetValues(0)
								sTextVal = Fn_UI_ListItemExist("Fn_RsrcMgr_AttributeValues", objApplet, "AttributeValueList",Trim(CStr(aGetValues(1))))
								Wait 1
								If sTextVal = True Then
									bFlag = True
								Else															
									Fn_RsrcMgr_AttributeValues = False
									Set objApplet = Nothing
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Verify value [" + aValue(iCount) + "] for Attribute " ) 
									Exit Function
								End If
						  	Next
						  	If bFlag = True Then
								Set objApplet = Nothing
								Fn_RsrcMgr_AttributeValues = True
							End If
					Case "Date"
							'Future Use
				End Select
		Case Else
				'Future Use
	End Select	
End Function
