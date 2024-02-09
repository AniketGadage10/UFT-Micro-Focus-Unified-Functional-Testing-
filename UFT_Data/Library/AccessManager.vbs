'--------------------------'Global variables for Teamcenter Perspective Names------------------------------------------------------------
Public GBL_PERSPECTIVE_ACCESS_MANAGER
GBL_PERSPECTIVE_ACCESS_MANAGER="Access Manager"
'--------------------------'Global variables for Teamcenter Perspective Names------------------------------------------------------------
''@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
''@@
''@@ NAME			: AccessManager.vbs
''@@
''@@ DESCRIPTION	: Contains functions used for Access Manager Prespective
''@@  
''@@ REQUIRED FILES	:
''@@   				  1. AccessManager.tsr			- Access Manager Object Respository
''@@				  2. UI_Library.vbs				- UI Library Functions
''@@    
''@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
''======================================================================================================
'' List Of Functions
''------------------------------------------------------------------------------------------------------
'' 					Function Name					| 						Owner
''------------------------------------------------------------------------------------------------------
'' 0. Fn_SISW_AM_GetObject							|	Sukhada Bakshi (sukhada.bakshi.ext@siemens.com)
'' 1. Fn_AccMgr_DialogMsgVerify						| 	Samir Thosar (samir.thosar@siemens.com)    - Eliminated. Can be replaced by GeneralFunctions.vbs::Fn_SISW_ErrorVerify()
'' 2. Fn_AccMgr_ImportExportAMRules					|	Harshal Agrawal(harshal.agrawal.ext@siemens.com)
'' 3. Fn_AccMgr_TreeOpeartion				        |   Ketan Raje(Ketan.Raje.ext@siemens.com)
'' 4. Fn_AccMgr_AMRuleOperation						|	Ketan Raje(Ketan.Raje.ext@siemens.com)
'' 5. Fn_AccMgr_ACLOperation						|	Ketan Raje(Ketan.Raje.ext@siemens.com)
'' 6. Fn_AccMgr_AccessControlOperation  			|	Ketan Raje(Ketan.Raje.ext@siemens.com)
'' 7. Fn_AccMgr_SaveResources						|	Shreyas Waichal(Shreyas.Waichal.ext@siemens.com)
'' 8. Fn_AccMgr_AccessControlOperationExt			|	Vivek Ahirrao(vivek.ahirrao.ext@siemens.com)
'' 9. Fn_AccMgr_AMRuleCreate						|	Vivek Ahirrao(vivek.ahirrao.ext@siemens.com) - Depricated function.  please use function Fn_AccMgr_AMRuleCreateExt
''10. Fn_AccMgr_AMRuleCreateExt						|	Sandeep Talware
''=======================================================================================================
'****************************************    Function to get Object hierarchy ***************************************
'
''Function Name		 	:	Fn_SISW_AM_GetObject
'
''Description		    :  	Function to get specified Object hierarchy.

''Parameters		    :	1. sObjectName : Object Handle name
								
''Return Value		    :  	Object \ Nothing
'
''Examples		     	:	Fn_SISW_AM_GetObject("AccMgrApplet")

'History:
'	Developer Name			Date			Rev. No.		Reviewer		Changes Done	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Koustubh Watwe		 12-Sep-2012		1.0	
'-----------------------------------------------------------------------------------------------------------------------------------
'	Ashwini Kumar		 25-Sept-2013		2.0								Externalized object hierarchies
'-----------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_AM_GetObject(sObjectName)
	Dim sObjectXMLPath
	sObjectXMLPath = Fn_GetEnvValue("User", "AutomationDir") & "\TestData\AutomationXML\ObjectXML\AccessManager.xml"
	Set Fn_SISW_AM_GetObject = Fn_SISW_Setup_GetObjectFromXML(sObjectXMLPath, sObjectName)
End Function
''--------------------------------------------------------------------------------------------------------
'' Function Number   	: 2                                                                                
'' Function Name     	: Fn_AccMgr_ImportExportAMRules
'' Function Description : Function to Import Export AM Rules
'' Function Pre-req		: Access Manager Prespective should be open
'' Function Usage    	: bReturn = Fn_AccMgr_ImportExportAMRules(sAction, sWinFolderPath, sFileName)
''							sAction			- Export / Import
''							sWinFolderPath	- Path of the folder to export / import AM Rules
''							sFileName		- Export / Import file name
''                     		Return True on Success and False on Failuer
'' Function History
'----------------------------------------------------------------------------------------------------------------------
'	Developer Name		|	  Date		|Rev. No.|		    Changes Done			|	Reviewer	|	Reviewed Date
'----------------------------------------------------------------------------------------------------------------------
	' 	Harshal Agrawal		|  9-June-2010	| 	1.0	 |								|				|		Harshal
	' 	Ashok kakade		|  22-may-2012	| 	1.0	 |								|				|		Koustubh
	'	Sanjeet K.					07-06-2012			1.1															Sachin
	'   Sukhada B           |  10-10-2012              Added Case Import_WithOutDelete
''**********************************************************************************************************************
Function Fn_AccMgr_ImportExportAMRules(sAction,sWinFolderPath, sFileName)
	GBL_FAILED_FUNCTION_NAME="Fn_AccMgr_ImportExportAMRules"
	Dim objFSO, objImport
	Dim bReturn, sWinPath,ObjExport

	sWinPath = sWinFolderPath & "\" & sFileName

	Select Case sAction
		Case "Export"
			Set objFSO = CreateObject("Scripting.FileSystemObject")
			If objFSO.FolderExists(sWinFolderPath) then
				bReturn = Fn_MenuOperation("Select","File:Export...")
				bReturn = Fn_Edit_Box("Fn_AccMgr_ImportExportAMRules",JavaDialog("ExportRule"),"File name",sWinPath)
				bReturn = Fn_Button_Click("Fn_AccMgr_ImportExportAMRules",JavaDialog("ExportRule"),"Export")
'				JavaWindow("Default Teamcenter Window").JavaWindow("AMAcessErrorDialog").SetTOProperty "title","Export AMRule Tree"
'				bReturn = Fn_Button_Click("Fn_AccMgr_ImportExportAMRules",JavaWindow("Default Teamcenter Window").JavaWindow("AMAcessErrorDialog"),"OK")
'	Set ObjExport = Fn_UI_ObjectCreate("Fn_AccMgr_ImportExportAMRules",JavaWindow("Default Teamcenter Window").JavaWindow("AccMgrApplet").JavaDialog("ImportExportAMRuleTree"))
				'Set ObjExport = JavaWindow("Default Teamcenter Window").JavaWindow("AccMgrApplet").JavaDialog("ImportExportAMRuleTree")
				
				'Eclipse Upgrade 4.10 Changes
				Set ObjExport = JavaWindow("Default Teamcenter Window").JavaWindow("AMRules").JavaDialog("ImportExportAMRuleTree")
				
				ObjExport.SetTOProperty "title","Export AMRule Tree"
				wait 1
				bReturn = Fn_Button_Click("Fn_AccMgr_ImportExportAMRules",ObjExport,"OK")
				Fn_AccMgr_ImportExportAMRules =True
				bReturn = Fn_WriteLogFile(Environment.Value("TestLogFile"), "Sucessfully Exported to "+sWinPath)
			Else
					Fn_AccMgr_ImportExportAMRules =False
					bReturn =  Fn_WriteLogFile(Environment.Value("TestLogFile"), "Folder Does not Exist at "+sWinPath)
			End if
		Case "Import"
        	Set objFSO = CreateObject("Scripting.FileSystemObject")
			If objFSO.FileExists(sWinPath) then
					bReturn = Fn_MenuOperation("Select","File:Import...")
					'Set objImport = Fn_UI_ObjectCreate("Fn_AccMgr_ImportExportAMRules",JavaWindow("Default Teamcenter Window").JavaWindow("AccMgrApplet").JavaDialog("ImportRule"))
					Set objImport = Fn_UI_ObjectCreate("Fn_AccMgr_ImportExportAMRules",JavaWindow("Default Teamcenter Window").JavaWindow("AMRules").JavaDialog("ImportRule"))
                    bReturn = Fn_Edit_Box("Fn_AccMgr_ImportExportAMRules",objImport,"File name",sWinPath)
					bReturn = Fn_Button_Click("Fn_AccMgr_ImportExportAMRules",objImport,"Import")
					bReturn = Fn_ReadyStatusSync(3)
'					JavaWindow("Default Teamcenter Window").JavaWindow("AMAcessErrorDialog").SetTOProperty "title","Import AMRule Tree"
'					bReturn = Fn_Button_Click("Fn_AccMgr_ImportExportAMRules",JavaWindow("Default Teamcenter Window").JavaWindow("AMAcessErrorDialog"),"OK")
					Set objImport =  JavaWindow("Default Teamcenter Window").JavaWindow("AMRules").JavaDialog("ImportExportAMRuleTree")
					objImport.SetTOProperty "title","Import AMRule Tree"
					wait 1
					bReturn = Fn_Button_Click("Fn_AccMgr_ImportExportAMRules",objImport,"OK")
					Fn_AccMgr_ImportExportAMRules =True
					bReturn = Fn_WriteLogFile(Environment.Value("TestLogFile"), "Sucessfully Imported from "+sWinPath)
					objFSO.DeleteFile(sWinPath)
			Else
					Fn_AccMgr_ImportExportAMRules =False
					bReturn = Fn_WriteLogFile(Environment.Value("TestLogFile"), "File Does not Exist at "+sWinPath)
			End if
				Case "Import_WithOutDelete"
        	Set objFSO = CreateObject("Scripting.FileSystemObject")
			If objFSO.FileExists(sWinPath) then
					bReturn = Fn_MenuOperation("Select","File:Import...")
					'Set objImport = Fn_UI_ObjectCreate("Fn_AccMgr_ImportExportAMRules",JavaWindow("Default Teamcenter Window").JavaWindow("AccMgrApplet").JavaDialog("ImportRule"))
					Set objImport = Fn_UI_ObjectCreate("Fn_AccMgr_ImportExportAMRules",JavaWindow("Default Teamcenter Window").JavaWindow("AMRules").JavaDialog("ImportRule"))
                    bReturn = Fn_Edit_Box("Fn_AccMgr_ImportExportAMRules",objImport,"File name",sWinPath)
					bReturn = Fn_Button_Click("Fn_AccMgr_ImportExportAMRules",objImport,"Import")
					Call Fn_ReadyStatusSync(3)
					Set objImport =  JavaWindow("Default Teamcenter Window").JavaWindow("AMRules").JavaDialog("ImportExportAMRuleTree")
					objImport.SetTOProperty "title","Import AMRule Tree"
					wait 1
					bReturn = Fn_Button_Click("Fn_AccMgr_ImportExportAMRules",objImport,"OK")
					Fn_AccMgr_ImportExportAMRules =True
					bReturn = Fn_WriteLogFile(Environment.Value("TestLogFile"), "Sucessfully Imported from "+sWinPath)
			Else
					Fn_AccMgr_ImportExportAMRules =False
					bReturn = Fn_WriteLogFile(Environment.Value("TestLogFile"), "File Does not Exist at "+sWinPath)
			End if
	End Select 
	Set objFSO = Nothing
	Set objImport = Nothing
End Function

''*********************************************************		Function to perform action on AccessManager Tree	***********************************************************************
'Function Name		:				Fn_AccMgr_TreeOpeartion()

'Description			 :		 		 Actions performed in this function are:
'																	1. Node Select
'                                                                   2. Node Expand
'																	3. Node Collapse
'																	4. Node Exist
'																	5. GetIndex	

'Parameters			   :	 			1. sAction: Action to be performed
'													2. sNodeName: Fully qulified tree Path (delimiter as ':') 

'Return Value		   : 				TRUE / FALSE and Index Value in "GetIndex" case.

'Pre-requisite			:		 		AccessManager Prespective is Open.

'Examples				:				Case "Select" : Call Fn_AccMgr_TreeOpeartion("Select","Has Class( POM_object ):Owning Group Has Security( Internal ) -> Internal Data")
'													Case "Expand" : Call Fn_AccMgr_TreeOpeartion("Expand","Has Class( POM_object ):Owning Group Has Security( Internal ) -> Internal Data")
'													Case "Collapse" : Call Fn_AccMgr_TreeOpeartion("Collapse","Has Class( POM_object ):Owning Group Has Security( Internal ) -> Internal Data")
'													Case "Exist" : Call Fn_AccMgr_TreeOpeartion("Exist","Has Class( POM_object ):Owning Group Has Security( Internal ) -> Internal Data")
'													Case "GetIndex" : Call Fn_AccMgr_TreeOpeartion("GetIndex","Has Class( POM_object ):Owning Group Has Security( Internal ) -> Internal Data")
'													Call Fn_AccMgr_TreeOpeartion("MoveDown", "Has Class( POM_object ):Has Status(  ) -> Vault")
'													Call Fn_AccMgr_TreeOpeartion("MoveDown", "")
'													Call Fn_AccMgr_TreeOpeartion("MoveUp", "")
'													Call Fn_AccMgr_TreeOpeartion("IsSelected","Has Class( POM_object ):Has Class( POM_object ) -> System Objects")
  												
'History					 :		
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'												------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Ketan Raje										   			10/06/2010			              1.0										Created									Harshal
'												------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Function Fn_AccMgr_TreeOpeartion(sAction,sNodeName)
	GBL_FAILED_FUNCTION_NAME="Fn_AccMgr_TreeOpeartion"
	
	Dim objJavaWindowAcc, objJavaTreeAcc, intNodeCount, intCount, sTreeItem, sItemName, iRow

	'If JavaWindow("Default Teamcenter Window").Exist(5) Then
	If Fn_SISW_UI_Object_Operations("Fn_AccMgr_TreeOpeartion","Exist",JavaWindow("Default Teamcenter Window"),SISW_MICROLESS_TIMEOUT) Then
'			Set objJavaWindowAcc = Fn_UI_ObjectCreate( "Fn_AccMgr_TreeOpeartion",JavaWindow("Default Teamcenter Window"))
'	Else
		Set objJavaWindowAcc = Fn_UI_ObjectCreate("Fn_AccMgr_TreeOpeartion", JavaWindow("Default Teamcenter Window").JavaWindow("AMRules"))
	End If
	


	Select Case sAction
		'----------------------------------------------------------------------- For selecting single node -------------------------------------------------------------------------
		Case "Select"
                    Call Fn_JavaTree_Select("Fn_AccMgr_TreeOpeartion", objJavaWindowAcc, "AMRuleTree",sNodeName)
					Wait(3)
					Fn_AccMgr_TreeOpeartion = TRUE
		'----------------------------------------------------------------------- For expanding a particular  node-------------------------------------------------------------------------
		Case "Expand"
                    Call Fn_UI_JavaTree_Expand("Fn_AccMgr_TreeOpeartion",objJavaWindowAcc,"AMRuleTree",sNodeName)
					Fn_AccMgr_TreeOpeartion = TRUE
		'----------------------------------------------------------------------- For collapssing a particular  node-------------------------------------------------------------------------
		Case "Collapse"
                    Call Fn_UI_JavaTree_Collapse("Fn_AccMgr_TreeOpeartion", objJavaWindowAcc,"AMRuleTree",sNodeName)
					Fn_AccMgr_TreeOpeartion = TRUE
		'----------------------------------------------------------------------- For Checking existance of a particular  node-------------------------------------------------------------------------
		Case "Exist"
				Set objJavaTreeAcc = Fn_UI_ObjectCreate( "Fn_AccMgr_TreeOpeartion", objJavaWindowAcc.JavaTree("AMRuleTree"))
					intNodeCount = objJavaTreeAcc.GetROProperty ("items count") 
					For intCount = 0 to intNodeCount - 1
						sTreeItem = objJavaTreeAcc.GetItem(intCount)
						If Trim(lcase(sTreeItem)) = Trim(Lcase(sNodeName)) Then
							Fn_AccMgr_TreeOpeartion = TRUE	
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[" + sNodeName + "] of JavaTree exist")
							Exit For
						End If
					Next
					If Cstr(intCount) = intNodeCount Then
						Fn_AccMgr_TreeOpeartion = FALSE						
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[" + sNodeName + "] of JavaTree does not exist")
						Exit Function
					End If
		'----------------------------------------------------------------------- Get Index value of a particular node-------------------------------------------------------------------------
		Case "GetIndex"
				bFlag = False
				For intCount=0 to objJavaWindowAcc.JavaTree("AMRuleTree").GetROProperty ("items count")-1
					sTreeItem = objJavaWindowAcc.JavaTree("AMRuleTree").GetItem (intCount)
					If Trim(lcase(sTreeItem)) = Trim(Lcase(sNodeName)) Then
						Fn_AccMgr_TreeOpeartion = intCount
						bFlag = True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The index of the given node is "&intCount)
						Exit For
					End If
				Next
				If bFlag = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The given node does not exist")
					Fn_AccMgr_TreeOpeartion = FALSE
				End If
		Case "MoveUp"
					intNodeCount = Fn_UI_Object_GetROProperty("Fn_AccMgr_TreeOpeartion",objJavaWindowAcc.JavaTree("AMRuleTree"),"items count")
					' For Traversing the Position of an Current Item
					sItemName = objJavaWindowAcc.JavaTree("AMRuleTree").GetROProperty("value")
					For intCount = 0 To intNodeCount -1
						sTreeItem = objJavaWindowAcc.JavaTree("AMRuleTree").GetItem(intCount)
						If sTreeItem = sItemName  Then
							iRow = intCount
							Exit For
						End If
					Next
					' For Moving Upwards
					For intCount = iRow To 1 Step - 1
						sTreeItem = objJavaWindowAcc.JavaTree("AMRuleTree").GetItem(intCount-1)
						wait(1)
						If sTreeItem = sNodeName  Then
							Exit For
						Else
							Call Fn_ToolbatButtonClick("Moves the selected AM Rule Up one level")
						End If
					Next
					Call Fn_ToolbatButtonClick("Saves the AM Rule Tree changes. (Ctrl+S)")
					Fn_AccMgr_TreeOpeartion = TRUE
		Case "MoveDown"
					intNodeCount = Fn_UI_Object_GetROProperty("Fn_AccMgr_TreeOpeartion",objJavaWindowAcc.JavaTree("AMRuleTree"),"items count")
					' For Traversing the Position of an Current Item
					sItemName = objJavaWindowAcc.JavaTree("AMRuleTree").GetROProperty("value")
					For intCount = 0 To intNodeCount - 1
						sTreeItem = objJavaWindowAcc.JavaTree("AMRuleTree").GetItem(intCount)
						If sTreeItem = sItemName  Then
							iRow = intCount
							Exit For
						End If
					Next
					' For Moving Downwards
					For intCount = iRow To intNodeCount - 2 Step + 1
						sTreeItem = objJavaWindowAcc.JavaTree("AMRuleTree").GetItem(intCount+1)
						If sTreeItem = sNodeName  Then
							Exit For
						Else
							Call Fn_ToolbatButtonClick("Moves the selected AM Rule Down one level")
						End If
					Next
					Call Fn_ToolbatButtonClick("Saves the AM Rule Tree changes. (Ctrl+S)")
					Fn_AccMgr_TreeOpeartion = TRUE
		Case "IsSelected"
			wait(3)
			Set objJavaTreeAcc = Fn_UI_ObjectCreate( "Fn_AccMgr_TreeOpeartion", objJavaWindowAcc.JavaTree("AMRuleTree"))				
				If Trim(Lcase(objJavaTreeAcc.GetROProperty("value"))) = Trim(Lcase(sNodeName)) Then
				   'Writing Log
				   Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Java Tree Node ["+sNodeName+"] is Selected .")
				   Fn_AccMgr_TreeOpeartion = TRUE
				Else
				   'Writing Log
				   Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Java Tree Node ["+sNodeName+"] is Not Selected .")
				   Fn_AccMgr_TreeOpeartion = FALSE
			End If
		Case Else
						Fn_AccMgr_TreeOpeartion = FALSE
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_AccMgr_TreeOpeartion function failed")
						Exit Function
	End Select
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" Sucessfully completed on Node [" + sNodeName + "] of JavaTree of function Fn_AccMgr_TreeOpeartion")
	Set objJavaWindowAcc = nothing
	Set objJavaTreeAcc = nothing
End Function

'#########################################################################################################
'###
'###    FUNCTION NAME   :   Fn_AccMgr_AMRuleOperation(sAction, sCondition, sValue, sACLName)
'###
'###    DESCRIPTION        :   Add / Modify / Delete / Exist
'###	Prequisite 					:	1.AccessManager Prespective is Open.
'###
'###    PARAMETERS      :  1.sAction: Add / Modify / Delete
'###  											2.sCondition : 
'###  											3.sValue : 
'###                                         	4.sACLName : 
'###
'###    Function Calls       :   Fn_WriteLogFile() Fn_UI_ObjectCreate()
'###
'###	 HISTORY             :   AUTHOR                 			DATE        	VERSION
'###
'###    CREATED BY     :    Ketan Raje                    14/06/2010         1.0
'###
'###    REVIWED BY     :    Harshal
'###
'###    MODIFIED BY   :  
'###
'###    EXAMPLE          : 		Case "Add" : Call Fn_AccMgr_AMRuleOperation("Add", "Has Class", "POM_application_object", "Harshal")
'###										 Case "Modify" : Call Fn_AccMgr_AMRuleOperation("Modify", "Has Class", "POM_attribute", "Harshal")
'###										 Case "Delete" : Call Fn_AccMgr_AMRuleOperation("Delete", "", "", "")
'###										 Case "Verify" : Call Fn_AccMgr_AMRuleOperation("Verify", "Has Class", "SavedSearch", "Harshal")	
'#############################################################################################################
Public Function Fn_AccMgr_AMRuleOperation(sAction, sCondition, sValue, sACLName)
	GBL_FAILED_FUNCTION_NAME="Fn_AccMgr_AMRuleOperation"
	Dim objAMRule, iReturn
	Fn_AccMgr_AMRuleOperation = false
	'If JavaWindow("Default Teamcenter Window").JavaWindow("AccMgrApplet").Exist(5) Then
	If Fn_SISW_UI_Object_Operations("Fn_AccMgr_AMRuleOperation","Exist",JavaWindow("Default Teamcenter Window").JavaWindow("AccMgrApplet"),SISW_MINLESS_TIMEOUT) Then
			Set objAMRule = Fn_UI_ObjectCreate("Fn_AccMgr_AMRuleOperation", JavaWindow("Default Teamcenter Window").JavaWindow("AccMgrApplet"))
	ElseIf JavaWindow("Default Teamcenter Window").JavaWindow("AMRules").Exist(1)  Then	
		Set objAMRule = Fn_UI_ObjectCreate("Fn_AccMgr_AMRuleOperation", JavaWindow("Default Teamcenter Window").JavaWindow("AMRules"))
	Else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "AM Rules object now found")
		exit function
	End If

		Select Case sAction
			Case "Add","Modify"
						If sCondition<>"" Then
							'Set AMRule Condition
									iReturn = objAMRule.JavaList("Condition").GetItemIndex(sCondition)
									'Select ACL Name.
									objAMRule.JavaList("Condition").Object.setSelectedIndex iReturn,True
									Wait 2
						End If
						If sValue<>"" Then
							'Set AMRule Value
'							iReturn = objAMRule.JavaList("Value").GetItemIndex(sValue)
'							objAMRule.JavaList("Value").Object.setSelectedIndex iReturn,True
'- - - - - - - - Old - - - -
'							If objAMRule.JavaList("Value").Exist Then
'									objAMRule.JavaList("Value").Object.setSelectedItem "",True
'									
'									'iReturn = objAMRule.JavaList("Value").GetItemIndex(sValue)
'									'Select ACL Name.
'									'objAMRule.JavaList("Value").Object.setSelectedIndex iReturn,True
'							Else If objAMRule.JavaEdit("Value").Exist Then
'									objAMRule.JavaEdit("Value").Set sValue
'								End  If
'							End If
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  Added by Sandeep - - - - - - - - - - - - - - - - - - - 
								' To Handle Java List
								If objAMRule.JavaList("Value").Exist(6) Then
										iReturn = objAMRule.JavaList("Value").GetItemIndex(sValue)
										objAMRule.JavaList("Value").Object.setSelectedIndex iReturn,True
								'To Handle Edit box
								ElseIf objAMRule.JavaEdit("Value").Exist(6) Then
										objAMRule.JavaEdit("Value").Set sValue
                                Elseif Window("AccessManagerWindow").JavaApplet("AccMgrApplet").JavaEdit("Value").Exist(4) then
										Window("AccessManagerWindow").JavaApplet("AccMgrApplet").JavaEdit("Value").Set sValue
								End  If
						End If
						If sACLName<>"" Then
							'Set ACLName
									iReturn = objAMRule.JavaList("ACL Name").GetItemIndex(sACLName)
									'Select ACL Name.
									objAMRule.JavaList("ACL Name").Object.setSelectedIndex iReturn,True
						End If
						If sAction="Add" Then
							'Click on Add button
							Call Fn_Button_Click("Fn_AccMgr_AMRuleOperation", objAMRule, "Add")
						ElseIf sAction="Modify" Then
							'Click on Modify button
							Call Fn_Button_Click("Fn_AccMgr_AMRuleOperation", objAMRule, "Modify")
						End If
			Case "Delete"
						'Click on Delete button
						Call Fn_Button_Click("Fn_AccMgr_AMRuleOperation", objAMRule, "Delete")
			Case "Verify"
						'Verify Condition
						If sCondition<>"" Then
							If Trim(Lcase(objAMRule.JavaList("Condition").GetItem(objAMRule.JavaList("Condition").Object.getSelectedIndex))) <>Trim(Lcase(sCondition)) Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Condition does not match")
								Set objAMRule = nothing 
								Fn_AccMgr_AMRuleOperation = FALSE	
								Exit Function
							End If
						End If
						'Verify Value
						If sValue<>"" Then
							If Trim(Lcase(objAMRule.JavaList("Value").GetItem(objAMRule.JavaList("Value").Object.getSelectedIndex))) <>Trim(Lcase(sValue)) Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Value does not match")
								Set objAMRule = nothing 
								Fn_AccMgr_AMRuleOperation = FALSE	
								Exit Function
							End If
						End If
						'Verify ACL Name
						If sACLName<>"" Then
							If Trim(Lcase(objAMRule.JavaList("ACL Name").GetItem(objAMRule.JavaList("ACL Name").Object.getSelectedIndex))) <>Trim(Lcase(sACLName)) Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "ACL Name does not match")
								Set objAMRule = nothing 
								Fn_AccMgr_AMRuleOperation = FALSE	
								Exit Function
							End If
						End If
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Condition, Value and ACL Name matches with supplied data")
			Case Else						
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_AccMgr_AMRuleOperation function failed")
						Fn_AccMgr_AMRuleOperation = FALSE
						Exit Function											
		End Select
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" Sucessfully completed of the function Fn_AccMgr_AMRuleOperation")
	Fn_AccMgr_AMRuleOperation = TRUE
	Set objAMRule = nothing 
End Function

'#########################################################################################################
'###
'###    FUNCTION NAME   :   Fn_AccMgr_ACLOperation(sAction, sACLName, sExtra)
'###
'###    DESCRIPTION        :   Create / Delete ACL.
'###	Prequisite 					:	1.AccessManager Prespective is Open.
'###
'###    PARAMETERS      :  1.sAction:Create / Delete
'###  											2.sACLName
'###  											3.sExtra
'###                                         
'###    Function Calls       :   Fn_WriteLogFile(), Fn_Button_Click(), Fn_UI_ObjectCreate()
'###
'###	 HISTORY             :   AUTHOR                 			DATE        	VERSION
'###
'###    CREATED BY     :    Ketan Raje                    14/06/2010         1.0
'###
'###    REVIWED BY     :    Harshal
'###
'###    MODIFIED BY   :  
'###
'###    EXAMPLE          : 		Case "Create" : Call Fn_AccMgr_ACLOperation("Create", "Sam", "")
'###										 Case "Delete" : Call Fn_AccMgr_ACLOperation("Delete", "Sam", "")
'###										 Case "Select" : Call Fn_AccMgr_ACLOperation("Select", "Sam", "")
'#############################################################################################################
Public Function Fn_AccMgr_ACLOperation(sAction, sACLName, sExtra)
	GBL_FAILED_FUNCTION_NAME="Fn_AccMgr_ACLOperation"
	Dim objACL, iReturn
	Dim objShell, objListObject
	Dim iCount , i
	Fn_AccMgr_ACLOperation = false
	If JavaWindow("Default Teamcenter Window").JavaWindow("AccMgrApplet").Exist(5) Then
		Set objACL = Fn_UI_ObjectCreate("Fn_AccMgr_ACLOperation", JavaWindow("Default Teamcenter Window").JavaWindow("AccMgrApplet"))
	ElseIf  JavaWindow("Default Teamcenter Window").JavaWindow("AMRules").Exist(5) Then	
		Set objACL =  Fn_UI_ObjectCreate("Fn_AccMgr_ACLOperation",JavaWindow("Default Teamcenter Window").JavaWindow("AMRules"))
	Else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "AM Rules object now found")
		exit function
	End If

		Select Case sAction
			Case "Create"
								'Set the list to Blank
								objACL.JavaList("ACL Name").Object.setSelectedItem "",True
								'Set ACL Name.
								objACL.JavaList("ACL Name").Type sACLName
								'Click on Create button.
								Call Fn_Button_Click("Fn_AccMgr_ACLOperation", objACL, "CreateACL")
			Case "Delete"
								iReturn = objACL.JavaList("ACL Name").GetItemIndex(sACLName)
								'Select ACL Name.
								objACL.JavaList("ACL Name").Object.setSelectedIndex iReturn,True
								'Click on Delete button.
								Call Fn_Button_Click("Fn_AccMgr_ACLOperation", objACL, "DeleteACL")
			Case "Select"
								objACL.JavaList("ACL Name").Select ""
								Set objShell = createobject("wscript.shell")
								
								objACL.JavaList("ACL Name").Click 1,1,"LEFT"
								set objListObject = objACL.JavaList("ACL Name").Object
								iCount=objListObject.getItemCount()
								
								For i=0 to iCount-1
									objShell.sendkeys "{DOWN}"
									If objListObject.getItemAt(i).toString = sACLName Then
										objShell.sendkeys "{ENTER}"
										Exit For 
									End If
								Next

								If i = iCount Then
									Fn_AccMgr_ACLOperation = FALSE
								End If
			Case Else						
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_AccMgr_ACLOperation function failed")
						Fn_AccMgr_ACLOperation = FALSE
						Exit Function											
		End Select
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" Sucessfully completed of the function Fn_AccMgr_ACLOperation")
	Fn_AccMgr_ACLOperation = TRUE
	Set objShell = nothing
	Set objACL = nothing 
End Function

'#########################################################################################################
'###
'###    FUNCTION NAME   :  Fn_AccMgr_AccessControlOperation(sActionType, aACLPropertyValue, iRowIndex)
'###
'###    DESCRIPTION     :   
'###
'###    PARAMETERS      :   sActionType: Add/Modify/Delete
'###                                        	 aACLPropertyValue:
'###                                         	iRowIndex: Required only for modify case
'###			   							
'###	RETURNS					: True/False
'###
'###    Function Calls  			:  Fn_WriteLogFile(), Fn_Button_Click(), Fn_UI_ObjectCreate(), Fn_UI_ObjectExist()
'###
'###	 HISTORY         			:   AUTHOR                 							DATE        		VERSION
'###
'###    CREATED BY      	 :  Ketan Raje        		  					  15/06/2010      		   1.0
'###
'###    REVIWED BY     		 : Harshal
'###
'###    MODIFIED BY     	 : Koustubh Watwe							5-Aug-2011				1.0				Added code to transform input array of size 28 to 30 and vice versa
'###    MODIFIED BY     	 : Koustubh Watwe							9-Aug-2011				1.0				Modified code to handle error dialog.
'###
'###    MODIFIED BY     	 : Amit Talegaonkar							12-Dec-2011				1.0				Added extra code to trasfrom Array with Ubound=28 to Ubound=32 to handle scenario with 33 Columns. [ Build=1130 ]
'###
'###
'###    EXAMPLE        			: 	  Dim aACLPropertyValue(29)	
'###  												aACLPropertyValue(0) = "Approver(RIG)"
'###												aACLPropertyValue(1) = "DBA in dba"
'###												aACLPropertyValue(2) = "yes"
'###												aACLPropertyValue(3) = "yes"
'###												aACLPropertyValue(4) = ""
'###												aACLPropertyValue(5) = "yes"
'###												aACLPropertyValue(6) = "yes"
'###												aACLPropertyValue(7) = "no"
'###												aACLPropertyValue(8) = ""
'###												aACLPropertyValue(9) = "yes"
'###												aACLPropertyValue(10) = "unset"
'###												aACLPropertyValue(11) = "unset"
'###												aACLPropertyValue(12) = "yes"
'###												aACLPropertyValue(13) = "yes"
'###												aACLPropertyValue(14) = "no"
'###												aACLPropertyValue(15) = "unset"
'###												aACLPropertyValue(16) = "yes"
'###												aACLPropertyValue(17) = "yes"
'###												aACLPropertyValue(18) = "yes"
'###												aACLPropertyValue(19) = "yes"
'###												aACLPropertyValue(20) = "unset"
'###												aACLPropertyValue(21) = "yes"
'###												aACLPropertyValue(22) = "no"
'###												aACLPropertyValue(23) = "yes"
'###												aACLPropertyValue(24) = "yes"
'###												aACLPropertyValue(25) = ""
'###												aACLPropertyValue(26) = "no"
'###												aACLPropertyValue(27) = "yes"
'###												aACLPropertyValue(28) = "yes"
'###
'###												Case "Add" : Call Fn_AccMgr_AccessControlOperation("Add", aACLPropertyValue, "")
'###												Case "Modify" : Call Fn_AccMgr_AccessControlOperation("Modify", aACLPropertyValue, "0")
'###												Case "Delete" : Call Fn_AccMgr_AccessControlOperation("Delete", "", "0")
'###												Case "GetList" : Call Fn_AccMgr_AccessControlOperation("GetList", aACLPropertyVal, "4")
'###												Case "Verify" : Call Fn_AccMgr_AccessControlOperation("Verify", "World", "")
'#############################################################################################################
Public Function Fn_AccMgr_AccessControlOperation(sActionType, aACLPropertyValue, iRowIndex)
	GBL_FAILED_FUNCTION_NAME="Fn_AccMgr_AccessControlOperation"
Dim objAccMgr, iTotalRow, iTotalCol, iCounter, iCnt, bFlag,sAppValue,aDummyACL(), i
ReDim aDummyACL(32)

If JavaWindow("Default Teamcenter Window").JavaWindow("AccMgrApplet").Exist(5) Then
	Set objAccMgr = Fn_UI_ObjectCreate("Fn_AccMgr_AccessControlOperation", JavaWindow("Default Teamcenter Window").JavaWindow("AccMgrApplet"))
Else
	Set objAccMgr = Fn_UI_ObjectCreate("Fn_AccMgr_AccessControlOperation", JavaWindow("Default Teamcenter Window").JavaWindow("AMRules"))
End If

	Select Case sActionType
		Case "Add"
					'Click on "+" button.
					Call Fn_Button_Click("Fn_AccMgr_AccessControlOperation", objAccMgr, "Add_ACLTable")
					iTotalRow = objAccMgr.JavaTable("ACLTable").GetROProperty("rows")
					iTotalCol = objAccMgr.JavaTable("ACLTable").GetROProperty("cols")
                    ' transforming array fsize from 28 to 31.
'					If uBound(aACLPropertyValue) = 28 AND cint(iTotalCol) = 31 Then
						If uBound(aACLPropertyValue) = 32 AND cint(iTotalCol) = 33 Then
						For i = 0 to uBound(aDummyACL) - 3
							aDummyACL(i)=aACLPropertyValue(i)
						Next
						
						aDummyACL(28)="unset"
						aDummyACL(29)="unset"
						aDummyACL(30)= aACLPropertyValue(uBound(aACLPropertyValue))
						
						ElseIf uBound(aACLPropertyValue) = 28 AND cint(iTotalCol) = 34 Then 
						ReDim aDummyACL(33)
						
						'Dummy[0-27] = aACLPropertyValue[0-27]
						For i = 0 to uBound(aDummyACL)						
							Select Case i
								Case 28,29,30,31,33 'New columns
									 aDummyACL(i) = "unset"
								Case 32
									 aDummyACL(32) = aACLPropertyValue(uBound(aACLPropertyValue))
								Case Else
									 aDummyACL(i) = aACLPropertyValue(i)
							End Select
						Next
						aDummyACL(uBound(aDummyACL))= aACLPropertyValue(uBound(aACLPropertyValue))
						
					ElseIf uBound(aACLPropertyValue) = 30 AND cint(iTotalCol) = 29 Then
						ReDim aDummyACL(28)
						for i = 0 to uBound(aDummyACL) - 1
							aDummyACL(i)=aACLPropertyValue(i)
						Next
						aDummyACL(uBound(aDummyACL))= aACLPropertyValue(uBound(aACLPropertyValue))
						
					'Added code to handle 33 columns ; Build - 1130.
					ElseIf uBound(aACLPropertyValue) = 28 AND cint(iTotalCol) = 33 Then
						ReDim aDummyACL(32)
						
						'Dummy[0-27] = aACLPropertyValue[0-27]
						For i = 0 to uBound(aDummyACL)						
							Select Case i
								Case 28,29,30,31 'New columns
									 aDummyACL(i) = "unset"
								Case 32
									 aDummyACL(32) = aACLPropertyValue(uBound(aACLPropertyValue))
								Case Else
									 aDummyACL(i) = aACLPropertyValue(i)
							End Select
						Next
						
					Else
						ReDim aDummyACL(uBound(aACLPropertyValue))
						for i = 0 to uBound(aDummyACL)
							aDummyACL(i)=aACLPropertyValue(i)
						Next
					End IF

					For iCounter = 0 To iTotalCol - 1
								If iCounter = 1 Then
									If Trim(aDummyACL(iCounter)) <> "" Then
										objAccMgr.JavaTable("ACLTable").DoubleClickCell iTotalRow-1,iCounter
										Wait 3
										If Fn_UI_ObjectExist("Fn_AccMgr_AccessControlOperation", objAccMgr.JavaDialog("Select Accessor")) = True Then
												objAccMgr.JavaDialog("Select Accessor").JavaList("AccessorList").Select aDummyACL(iCounter)
													Call Fn_Button_Click("Fn_AccMgr_AccessControlOperation",objAccMgr.JavaDialog("Select Accessor"), "OK")
										End If
									End If
								ElseIf iCounter = 2 Then
									If Trim(aDummyACL(iCounter)) <> "" Then
										objAccMgr.JavaTable("ACLTable").DoubleClickCell  iTotalRow-1,iCounter, "LEFT"
										If Fn_UI_ObjectExist("Fn_AccMgr_AccessControlOperation", objAccMgr.JavaList("ACLTableValues")) = True  Then
												objAccMgr.JavaList("ACLTableValues").Select  Trim(aDummyACL(iCounter))
										End If
										If Fn_UI_ObjectExist("Fn_AccMgr_AccessControlOperation", JavaWindow("Default Teamcenter Window").Dialog("Warning")) = True Then
													iWinId = JavaWindow("Default Teamcenter Window").Dialog("Warning").GetROProperty("window id")
													JavaWindow("Default Teamcenter Window").Dialog("Warning").SetTOProperty "window id", iWinId
													Do
													 sLogText =  JavaWindow("Default Teamcenter Window").Dialog("Warning").GetROProperty("text")
													 If objAccMgr.JavaDialog("Warning").JavaButton("OK").exist(5) then
														Call Fn_Button_Click("Fn_AccMgr_AccessControlOperation", objAccMgr.JavaDialog("Warning"), "OK")
													 elseif JavaWindow("Default Teamcenter Window").Dialog("Warning").WinButton("OK").exist(5) then
														Call Fn_UI_WinButton_Click("Fn_AccMgr_AccessControlOperation",JavaWindow("Default Teamcenter Window").Dialog("Warning"),"OK","","","")
													 End If
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Warning Dialog for Message["+CStr(sLogText)+"]  handled Successfully.")
													Loop While Fn_UI_ObjectExist("Fn_AccMgr_AccessControlOperation", JavaWindow("Default Teamcenter Window").Dialog("Warning")) = True
										ElseIf objAccMgr.JavaDialog("Warning").Exist(5) then
										            Do
														objAccMgr.JavaDialog("Warning").JavaButton("OK").Click	
													Loop While Fn_UI_ObjectExist("Fn_AccMgr_AccessControlOperation", JavaWindow("Default Teamcenter Window").JavaWindow("AccMgrApplet").JavaDialog("Warning")) = True
										End If
									End If									
								Elseif Trim(aDummyACL(iCounter)) <> "" Then
									objAccMgr.JavaTable("ACLTable").SetCellData "#"+CStr(iTotalRow-1),"#"+CStr(iCounter), aDummyACL(iCounter)
								End If
					Next
					'Click on Save button
					Call Fn_Button_Click("Fn_AccMgr_AccessControlOperation", objAccMgr, "Save_ACLTable")			
		Case "Modify"
					If iRowIndex >= objAccMgr.JavaTable("ACLTable").GetROProperty("rows") Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_AccMgr_AccessControlOperation function failed as the Row does not exist")
							Fn_AccMgr_AccessControlOperation = FALSE
							Set objAccMgr = Nothing
							Exit Function						
					End If
					objAccMgr.JavaTable("ACLTable").SelectRow iRowIndex
					iTotalRow = objAccMgr.JavaTable("ACLTable").GetROProperty("rows")
					iTotalCol = objAccMgr.JavaTable("ACLTable").GetROProperty("cols")
                    ' transforming array fsize from 28 to 31.
					If uBound(aACLPropertyValue) = 28 AND cint(iTotalCol) = 31 Then
						for i = 0 to uBound(aDummyACL) - 3
							aDummyACL(i)=aACLPropertyValue(i)
						Next
						aDummyACL(28)="unset"
						aDummyACL(29)="unset"
						aDummyACL(30)= aACLPropertyValue(uBound(aACLPropertyValue))
					ElseIf uBound(aACLPropertyValue) = 30 AND cint(iTotalCol) = 29 Then
						ReDim aDummyACL(28)
						for i = 0 to uBound(aDummyACL) - 1
							aDummyACL(i)=aACLPropertyValue(i)
						Next
						aDummyACL(uBound(aDummyACL))= aACLPropertyValue(uBound(aACLPropertyValue))
					Else
						ReDim aDummyACL(uBound(aACLPropertyValue))
						for i = 0 to uBound(aDummyACL)-1
							aDummyACL(i)=aACLPropertyValue(i)
						Next
					End IF
					For iCounter = 1 To iTotalCol - 1
								objAccMgr.JavaTable("ACLTable").ClickCell iRowIndex, iCounter, "LEFT"
								If iCounter = 1 Then
									If  Trim(aDummyACL(iCounter)) <> "" Then
										If Fn_UI_ObjectExist("Fn_AccMgr_AccessControlOperation", objAccMgr.JavaDialog("Select Accessor")) = False Then
											objAccMgr.JavaTable("ACLTable").DoubleClickCell "#"+CStr(iRowIndex),"#"+CStr(iCounter), "LEFT"
										End If
										Wait 3
												objAccMgr.JavaDialog("Select Accessor").JavaList("AccessorList").Select aDummyACL(iCounter)
													Call Fn_Button_Click("Fn_AccMgr_AccessControlOperation", objAccMgr.JavaDialog("Select Accessor"), "OK")
									End If
								ElseIf iCounter = 2 Then
									If Trim(aDummyACL(iCounter)) <> "" Then
										objAccMgr.JavaTable("ACLTable").DoubleClickCell  iRowIndex,iCounter, "LEFT"
										If Fn_UI_ObjectExist("Fn_AccMgr_AccessControlOperation", JavaWindow("Default Teamcenter Window").Dialog("Warning")) = True Then
													iWinId = JavaWindow("Default Teamcenter Window").Dialog("Warning").GetROProperty("window id")
													JavaWindow("Default Teamcenter Window").Dialog("Warning").SetTOProperty "window id", iWinId
													Do
													 sLogText =  JavaWindow("Default Teamcenter Window").Dialog("Warning").GetROProperty("text")
													 If  JavaWindow("Default Teamcenter Window").Dialog("Warning").WinButton("OK").Exist(3) Then		'==== Added code to handle Java Button  [30-Jan-2012]
														Call Fn_UI_WinButton_Click("Fn_AccMgr_AccessControlOperation",JavaWindow("Default Teamcenter Window").Dialog("Warning"),"OK","","","")
													 Else
														Exit Do
													 End If
													Call Fn_UI_WinButton_Click("Fn_AccMgr_AccessControlOperation",JavaWindow("Default Teamcenter Window").Dialog("Warning"),"OK","","","")
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Warning Dialog for Message["+CStr(sLogText)+"]  handled Successfully.")
													Loop While Fn_UI_ObjectExist("Fn_AccMgr_AccessControlOperation", JavaWindow("Default Teamcenter Window").Dialog("Warning")) = True
										End If
										If Fn_UI_ObjectExist("Fn_AccMgr_AccessControlOperation", objAccMgr.JavaList("ACLTableValues")) = True  Then
												objAccMgr.JavaList("ACLTableValues").Select  Trim(aDummyACL(iCounter))
										End If
										If Fn_UI_ObjectExist("Fn_AccMgr_AccessControlOperation", objAccMgr.JavaDialog("Warning")) = True Then
'													iWinId = JavaWindow("Default Teamcenter Window").Dialog("Warning").GetROProperty("window id")
'													JavaWindow("Default Teamcenter Window").Dialog("Warning").SetTOProperty "window id", iWinId
													Do
'													 sLogText =  JavaWindow("Default Teamcenter Window").Dialog("Warning").GetROProperty("text")
'													Call Fn_UI_WinButton_Click("Fn_AccMgr_AccessControlOperation",JavaWindow("Default Teamcenter Window").Dialog("Warning"),"OK","","","")
													If  objAccMgr.JavaDialog("Warning").JavaButton("OK").Exist(5) Then
														Call Fn_Button_Click("Fn_AccMgr_AccessControlOperation",objAccMgr.JavaDialog("Warning"),"OK")
													ElseIf objAccMgr.JavaButton("OK").Exist(5) Then
														Call Fn_Button_Click("Fn_AccMgr_AccessControlOperation",JavaWindow("Default Teamcenter Window").JavaWindow("AccMgrApplet"),"OK")
													End If
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Warning Dialog for Message["+CStr(sLogText)+"]  handled Successfully.")
													Loop While Fn_UI_ObjectExist("Fn_AccMgr_AccessControlOperation", objAccMgr.JavaDialog("Warning")) = True
													objAccMgr.JavaTable("ACLTable").SetCellData "#"+CStr(iRowIndex),"#"+CStr(iCounter), aDummyACL(iCounter)					'==== To set the value of the cell after handling the window
										End If
									End If									
								Elseif Trim(aDummyACL(iCounter)) <> "" Then
									objAccMgr.JavaTable("ACLTable").SetCellData "#"+CStr(iRowIndex),"#"+CStr(iCounter), aDummyACL(iCounter)
								End If
					Next
					Call Fn_Button_Click("Fn_AccMgr_AccessControlOperation", objAccMgr, "Save_ACLTable")			
					Fn_AccMgr_AccessControlOperation = True
		Case "Delete"
					If iRowIndex >= objAccMgr.JavaTable("ACLTable").GetROProperty("rows") Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_AccMgr_AccessControlOperation function failed as the Row does not exist")
							Fn_AccMgr_AccessControlOperation = FALSE
							Set objAccMgr = Nothing
							Exit Function						
					End If
					objAccMgr.JavaTable("ACLTable").SelectRow iRowIndex
					'Click on Remove button
					Call Fn_Button_Click("Fn_AccMgr_AccessControlOperation", objAccMgr, "Remove_ACLTable")
					'Click on save button
					Call Fn_Button_Click("Fn_AccMgr_AccessControlOperation", objAccMgr, "Save_ACLTable")
		Case "GetList"
					If iRowIndex >= objAccMgr.JavaTable("ACLTable").GetROProperty("rows") Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_AccMgr_AccessControlOperation function failed as the Row does not exist")
							Fn_AccMgr_AccessControlOperation = FALSE
							Set objAccMgr = Nothing
							Exit Function						
					End If
					'Select given row from ACL Table
					objAccMgr.JavaTable("ACLTable").SelectRow iRowIndex
					'Count number of rows
					iTotalRow = objAccMgr.JavaTable("ACLTable").GetROProperty("rows")
					'Count number of cols
					iTotalCol = objAccMgr.JavaTable("ACLTable").GetROProperty("cols")
					For iCounter = 0 To iTotalCol - 1
						If iCounter > 1 Then
								objAccMgr.JavaTable("ACLTable").DoubleClickCell "#"+CStr(iRowIndex),"#"+CStr(iCounter),"LEFT","NONE"								
								aACLPropertyValue(iCounter) = objAccMgr.JavaList("ACLTableValues").GetItem(objAccMgr.JavaList("ACLTableValues").Object.getSelectedIndex)
						Else
								objAccMgr.JavaTable("ACLTable").ClickCell iRowIndex, iCounter, "LEFT"
								aACLPropertyValue(iCounter) = objAccMgr.JavaTable("ACLTable").GetCellData("#"+CStr(iRowIndex),"#"+CStr(iCounter))
						End If
					Next
					Fn_AccMgr_AccessControlOperation = aACLPropertyValue
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sActionType&" case sucessfully completed of function Fn_AccMgr_AccessControlOperation")
					Set objAccMgr = Nothing
					Exit Function
		Case "Verify"
					If objAccMgr.JavaTable("ACLTable").GetROProperty("rows") = 0 Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_AccMgr_AccessControlOperation function failed as the Row does not exist")
							Fn_AccMgr_AccessControlOperation = FALSE
							Set objAccMgr = Nothing
							Exit Function						
					End If
					'Count number of rows
					iTotalRow = objAccMgr.JavaTable("ACLTable").GetROProperty("rows")
					'Count number of cols
					iTotalCol = objAccMgr.JavaTable("ACLTable").GetROProperty("cols")
					For iCnt = 0 To iTotalRow - 1
						For iCounter = 0 To iTotalCol - 1
							If iCounter > 1 Then
									objAccMgr.JavaTable("ACLTable").DoubleClickCell "#"+CStr(iCnt),"#"+CStr(iCounter),"LEFT","NONE"								
									If  objAccMgr.JavaDialog("Warning").JavaButton("OK").Exist(0) Then		'-----Added code to handle the warning msg Build:627[10-Jul-2012]
										Call Fn_Button_Click("Fn_AccMgr_AccessControlOperation",objAccMgr.JavaDialog("Warning"),"OK")
									ElseIf objAccMgr.JavaButton("OK").Exist(0) Then
										Call Fn_Button_Click("Fn_AccMgr_AccessControlOperation",objAccMgr,"OK")
									End If
									If aACLPropertyValue = objAccMgr.JavaList("ACLTableValues").GetItem(objAccMgr.JavaList("ACLTableValues").Object.getSelectedIndex) Then
										If iCnt = 0 Then	' ------------in case the column value found at 0th row
											Fn_AccMgr_AccessControlOperation = "#0"
										Else
											Fn_AccMgr_AccessControlOperation = iCnt
										End If
                                            Call Fn_WriteLogFile(Environment.Value("TestLogFile"), aACLPropertyValue&" is at "&iCnt&" row in ACL Table.")
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sActionType&" case sucessfully completed of function Fn_AccMgr_AccessControlOperation")
											Set objAccMgr = Nothing
											Exit Function										
									End If
							Else
									objAccMgr.JavaTable("ACLTable").ClickCell iCnt, iCounter, "LEFT"
									If aACLPropertyValue = objAccMgr.JavaTable("ACLTable").GetCellData("#"+CStr(iCnt),"#"+CStr(iCounter)) Then		
											If iCnt = 0 Then    ' -----------------in case the column value found at 0th row
												Fn_AccMgr_AccessControlOperation =  "#0"
											Else
												Fn_AccMgr_AccessControlOperation = iCnt
											End If
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), aACLPropertyValue&" is at "&iCnt&" row in ACL Table.")
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sActionType&" case sucessfully completed of function Fn_AccMgr_AccessControlOperation")
											Set objAccMgr = Nothing
											Exit Function										
									End If
							End If
						Next
					Next
'					Steps is done to defocus the cell			Coded By Harshal on 24-Aug-2011.
					objAccMgr.JavaTable("ACLTable").ClickCell 0, 0, "LEFT"
					Fn_AccMgr_AccessControlOperation = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sActionType&" case sucessfully completed of function Fn_AccMgr_AccessControlOperation")
					Set objAccMgr = Nothing
					Exit Function
		Case Else						
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_AccMgr_AccessControlOperation function failed")
					Fn_AccMgr_AccessControlOperation = FALSE
					Set objAccMgr = Nothing
					Exit Function											
	End Select
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sActionType&" case sucessfully completed of function Fn_AccMgr_AccessControlOperation")
'	Steps is done to defocus the cell
	If objAccMgr.JavaTable("ACLTable").GetROproperty("rows")>0 Then
		objAccMgr.JavaTable("ACLTable").ClickCell 0, 0, "LEFT"
	End If
	Fn_AccMgr_AccessControlOperation = True
Set objAccMgr = Nothing
End Function

'###############################################################################################################################################################################
'###	Depricated function.  please use function Fn_AccMgr_AMRuleCreateExt
'###    FUNCTION NAME   :  	Fn_AccMgr_AMRuleCreate    -    Function to create new AM Rule along with new ACL 
'###    
'###    Modified By 	: 	Vivek Ahirrao [TC1015-2015092200-09_10_2015-VivekA-Maintenance] - Modified for Array changing, and called new Function.
'###    
'###############################################################################################################################################################################
Public Function Fn_AccMgr_AMRuleCreate(sCondition, sValue, sACLName, aACLPropertyValue)
	GBL_FAILED_FUNCTION_NAME="Fn_AccMgr_AMRuleCreate"
	Dim objAMRule, iReturn
    Dim iTotalRow, iTotalCol, iCounter, iCnt
	
	If JavaWindow("Default Teamcenter Window").JavaWindow("AccMgrApplet").Exist(5) Then
			Set objAMRule = Fn_UI_ObjectCreate("Fn_AccMgr_AMRuleOperation", JavaWindow("Default Teamcenter Window").JavaWindow("AccMgrApplet"))
	Else
		Set objAMRule = Fn_UI_ObjectCreate("Fn_AccMgr_AMRuleOperation", JavaWindow("Default Teamcenter Window").JavaWindow("AMRules"))
	End If
		
		
	If sCondition<>"" Then
		'Set AMRule Condition
				iReturn = objAMRule.JavaList("Condition").GetItemIndex(sCondition)
				'Select ACL Name.
				objAMRule.JavaList("Condition").Object.setSelectedIndex iReturn,True
	End If
	If sValue <> "" Then
		'Set the list to Blank
		objAMRule.JavaList("Value").Object.setSelectedItem "",True
		'Set ACL Name.
         objAMRule.JavaList("Value").Type sValue
		 'objAMRule.JavaList("Value").Activate
	End If
	'Set the list to Blank
	objAMRule.JavaList("ACL Name").Object.setSelectedItem "",True
	'Set ACL Name.
	objAMRule.JavaList("ACL Name").Type sACLName
	'Click on Create button.
	Call Fn_Button_Click("Fn_AccMgr_ACLOperation", objAMRule, "CreateACL")
	'call function Fn_AccMgr_AccessControlOperation to add ACL
    bReturn = Fn_AccMgr_AccessControlOperation("Add", aACLPropertyValue, "")
	If bReturn = False Then
       Fn_AccMgr_AMRuleCreate = False
	   Exit Function
	End If

	Call Fn_Button_Click("Fn_AccMgr_AMRuleOperation", objAMRule, "Add")
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" Sucessfully completed of the function Fn_AccMgr_AMRuleOperation")
	Fn_AccMgr_AMRuleCreate = TRUE
	Set objAMRule = nothing 
End Function
'###############################################################################################################################################################################
'###
'###    FUNCTION NAME   :  	Fn_AccMgr_AMRuleCreateExt
'###
'###    DESCRIPTION     :  	Function to create new AM Rule along with new ACL
'###
'###    PARAMETERS      :  	sCondition			: 	
'###                        sValue				:	
'###                        sACLName			:	
'###                        sACLProperties		:	"Type of Accessor~ID of Accessor~Read~Write"
'###                        sACLPropertyValue	:	"User~AutoTest2 (autotest2)~yes~no"
'###			   							
'###	RETURNS			: 	True/False/Array of ACL columns' values/Row number for specific column value
'###
'###	HISTORY        	:   AUTHOR             DATE        	VERSION				
'###
'###    CREATED BY      :  	Sandeep T		 01/06/2016		  1.0
'###
'###	ACL Properties 
'###	 					Type of Accessor, ID of Accessor, Read, Write, Delete, Change, Promote, Demote, Copy, Change Ownership, Publish, Subscribe, Export, Import, 
'###	 					Transfer Out, Transfer In, Write Classification ICOs, Assign to Project, Remove from Project, Remote Check-Out, Unmanage, IP Administrator
'###	 					ITAR Administrator, ITAR Classifier, IP Classifier, Check-In/Check-Out, Administer ADA Licenses, Translation, Markup, Batch Print, 
'###	 					Digitally Sign, Add Content, Remove Content, Void Digital Signature, Manage Variability
'###    
'###    EXAMPLEs       	: 	  
'###					
'###					 sACLProperties="Type of Accessor~Read~Write~Delete"
'###					 sACLPropertyValues="User ITAR Licensed~yes~yes~yes"
'###					 bReturn = Fn_AccMgr_AMRuleCreateExt("Has Government Classification", "", "ACLName123",sACLProperties,sACLPropertyValues)
'###						
'###############################################################################################################################################################################
Public Function Fn_AccMgr_AMRuleCreateExt(sCondition, sValue, sACLName, sACLProperties, sACLPropertyValue)
	GBL_FAILED_FUNCTION_NAME="Fn_AccMgr_AMRuleCreateExt"
	Dim objAMRule, iReturn
    Dim iTotalRow, iTotalCol, iCounter, iCnt
    Fn_AccMgr_AMRuleCreateExt = False

	'If JavaWindow("Default Teamcenter Window").JavaWindow("AccMgrApplet").Exist(5) Then
	If Fn_SISW_UI_Object_Operations("Fn_AccMgr_AMRuleCreateExt","Exist",JavaWindow("Default Teamcenter Window").JavaWindow("AccMgrApplet"),SISW_MINLESS_TIMEOUT) Then
		Set objAMRule = Fn_UI_ObjectCreate("Fn_AccMgr_AMRuleOperation", JavaWindow("Default Teamcenter Window").JavaWindow("AccMgrApplet"))
	Else
		Set objAMRule = Fn_UI_ObjectCreate("Fn_AccMgr_AMRuleOperation", JavaWindow("Default Teamcenter Window").JavaWindow("AMRules"))
	End If
		
		
	If sCondition<>"" Then
		'Set AMRule Condition
		iReturn = objAMRule.JavaList("Condition").GetItemIndex(sCondition)
		'Select ACL Name.
		objAMRule.JavaList("Condition").Object.setSelectedIndex iReturn,True
		Wait 1
	End If
	If sValue <> "" Then
		'Set the list to Blank
		objAMRule.JavaList("Value").Object.setSelectedItem "",True
		'Set ACL Name.
		objAMRule.JavaList("Value").Type sValue
		wait 1
	End If
	'Set the list to Blank
	objAMRule.JavaList("ACL Name").Object.setSelectedItem "",True
	'Set ACL Name.
	objAMRule.JavaList("ACL Name").Type sACLName
	'Click on Create button.
	Call Fn_Button_Click("Fn_AccMgr_ACLOperation", objAMRule, "CreateACL")
	'call function Fn_AccMgr_AccessControlOperation to add ACL
    bReturn = Fn_AccMgr_AccessControlOperationExt("Add", sACLProperties, sACLPropertyValue, "","")
	If bReturn = False Then
		Exit Function
	End If

	Call Fn_Button_Click("Fn_AccMgr_AMRuleOperation", objAMRule, "Add")
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" Sucessfully completed of the function Fn_AccMgr_AMRuleOperation")
	Fn_AccMgr_AMRuleCreateExt = TRUE
	Set objAMRule = nothing 
End Function
'$$$$$$$$$$$$$$$$$  Added  By shreyas $$$$$$$$$$$$$$$$$$$$$$
'This function is to handle the Save Resources dialog which occurs after closing the access manager perspective without saving the changes
 'Example bReturn=Fn_AccMgr_SaveResources("Yes")
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Public function Fn_AccMgr_SaveResources(sButton)
	GBL_FAILED_FUNCTION_NAME="Fn_AccMgr_SaveResources"
   Dim objResource

   Set objResource=JavaWindow("Default Teamcenter Window").JavaWindow("SaveResource")

	 	If objResource.Exist  Then
			objResource.JavaButton(sButton).Click micLeftBtn
			If err.number>0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile")," Failed to click on the button"+sButton)
				Exit Function				
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile")," Successfully clicked on the button"+sButton)
				Fn_AccMgr_SaveResources=True
				Set objResource = nothing 
			End If
		End If
End Function

'###############################################################################################################################################################################
'###
'###    FUNCTION NAME   :  	Fn_AccMgr_AccessControlOperationExt
'###
'###    DESCRIPTION     :  	Function used for ACL operations
'###
'###    PARAMETERS      :  	sActionType			: 	Case Name "Add", "Modify" etc
'###                        sACLProperties		:	"Type of Accessor~ID of Accessor~Read~Write"
'###                        sACLPropertyValue	:	"User~AutoTest2 (autotest2)~yes~no"
'###                        iRowIndex			:	index of row
'###                        sReserve			:	Future use
'###			   							
'###	RETURNS			: 	True/False/Array of ACL columns' values/Row number for specific column value
'###
'###	HISTORY        	:   AUTHOR             DATE        	VERSION				
'###
'###    CREATED BY      :  	Vivek A			01/10/2015		  1.0
'###
'###	ACL Column Headers 
'###	 					Type of Accessor, ID of Accessor, Read, Write, Delete, Change, Promote, Demote, Copy, Change Ownership, Publish, Subscribe, Export, Import, 
'###	 					Transfer Out, Transfer In, Write Classification ICOs, Assign to Project, Remove from Project, Remote Check-Out, Unmanage, IP Administrator
'###	 					ITAR Administrator, ITAR Classifier, IP Classifier, Check-In/Check-Out, Administer ADA Licenses, Translation, Markup, Batch Print, 
'###	 					Digitally Sign, Add Content, Remove Content, Void Digital Signature, Manage Variability
'###    
'###    EXAMPLEs       	: 	  
'###						Case "Add" 		: bReturn = Fn_AccMgr_AccessControlOperationExt("Add", sACLProperties, sACLPropertyValue, "", "")
'###						Case "Modify" 	: bReturn = Fn_AccMgr_AccessControlOperationExt("Modify", sACLProperties, sACLPropertyValue, "0", "")
'###						Case "Delete" 	: bReturn = Fn_AccMgr_AccessControlOperationExt("Delete", "", "", "0", "")
'###						Case "GetList" 	: bReturn = Fn_AccMgr_AccessControlOperationExt("GetList", "", "", "1", "")
'###						Case "Verify" 	: bReturn = Fn_AccMgr_AccessControlOperationExt("Verify", "", "User", "", "")
'###						Case "GetRowForColumnValue" 	: bReturn = Fn_AccMgr_AccessControlOperationExt("GetRowForColumnValue", "", "User", "", "")
'###						
'###############################################################################################################################################################################
Public Function Fn_AccMgr_AccessControlOperationExt(sActionType, sACLProperties, sACLPropertyValue, iRowIndex, sReserve)
	GBL_FAILED_FUNCTION_NAME="Fn_AccMgr_AccessControlOperationExt"
	
	Dim objAccMgr, iWinId, sLogText
	Dim iTotalRow, iTotalCol, iCounter, iColNum, iCount
	Dim aACLProperties, aACLPropertyValue
	
	Fn_AccMgr_AccessControlOperationExt = false
	
	'If JavaWindow("Default Teamcenter Window").JavaWindow("AccMgrApplet").Exist(5) Then
	If Fn_SISW_UI_Object_Operations("Fn_AccMgr_AccessControlOperationExt","Exist",JavaWindow("Default Teamcenter Window").JavaWindow("AccMgrApplet"),SISW_MINLESS_TIMEOUT) Then
		Set objAccMgr = Fn_UI_ObjectCreate("Fn_AccMgr_AccessControlOperationExt", JavaWindow("Default Teamcenter Window").JavaWindow("AccMgrApplet"))
	Else
		Set objAccMgr = Fn_UI_ObjectCreate("Fn_AccMgr_AccessControlOperationExt", JavaWindow("Default Teamcenter Window").JavaWindow("AMRules"))
	End If

	iTotalRow = cInt(objAccMgr.JavaTable("ACLTable").GetROProperty("rows"))
	iTotalCol = cInt(objAccMgr.JavaTable("ACLTable").GetROProperty("cols"))
	
	Select Case sActionType
		'Case to get Column index
		Case "GetColumnIndex"
			Fn_AccMgr_AccessControlOperationExt = -1
			For iCounter = 0 to iTotalCol - 1
				If sACLProperties = objAccMgr.JavaTable("ACLTable").Object.getColumnModel().getColumn(iCounter).getHeaderValue().getToolTipText() Then
					Fn_AccMgr_AccessControlOperationExt = iCounter
					Set objAccMgr = Nothing
					Exit function
				End If
			Next
			
		'Case to add or modify ACL Rule Row		
		Case "Add", "Modify"
			aACLProperties = split(sACLProperties,"~")
			aACLPropertyValue = split(sACLPropertyValue,"~")
			
			If sActionType = "Add" Then
				'Click on "+" button.
				Call Fn_Button_Click("Fn_AccMgr_AccessControlOperationExt", objAccMgr, "Add_ACLTable")
				iTotalRow = cInt(objAccMgr.JavaTable("ACLTable").GetROProperty("rows"))
				iTotalCol = cInt(objAccMgr.JavaTable("ACLTable").GetROProperty("cols"))
				iRowIndex =  iTotalRow-1
			ElseIf sActionType = "Modify" Then
				If iRowIndex >= objAccMgr.JavaTable("ACLTable").GetROProperty("rows") Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_AccMgr_AccessControlOperationExt function failed as the Row does not exist")
					Fn_AccMgr_AccessControlOperationExt = FALSE
					Set objAccMgr = Nothing
					Exit Function						
				End If
				objAccMgr.JavaTable("ACLTable").SelectRow iRowIndex
				If Err.Number <> 0 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_ACLOperationsExt function failed as select row is not performed on table")
					Fn_AccMgr_AccessControlOperationExt = False
					Set objAccMgr = Nothing
					Exit Function
				End If
				iTotalRow = objAccMgr.JavaTable("ACLTable").GetROProperty("rows")
				iTotalCol = objAccMgr.JavaTable("ACLTable").GetROProperty("cols")
			End If
								
			For iCounter = 0 To UBound(aACLProperties)
				Select Case Trim(aACLProperties(iCounter))
					Case "Type of Accessor"
						'iColNum = 0
						iColNum = Fn_AccMgr_AccessControlOperationExt("GetColumnIndex", Trim(aACLProperties(iCounter)), "", "", "")
						If iColNum <> -1 Then
							objAccMgr.JavaTable("ACLTable").DoubleClickCell iRowIndex, iColNum, "LEFT"
							If Err.Number <> 0 Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_ACLOperationsExt function failed as double click is not performed on table")
								Fn_AccMgr_AccessControlOperationExt = False
								Set objAccMgr = Nothing
								Exit Function
							End If
							If Fn_UI_ObjectExist("Fn_AccMgr_AccessControlOperationExt", objAccMgr.JavaList("ACLTableValues")) = True  Then
								objAccMgr.JavaList("ACLTableValues").Select  Trim(aACLPropertyValue(iCounter))
								If Err.Number <> 0 Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_ACLOperationsExt function failed as not able to select node from List")
									Fn_AccMgr_AccessControlOperationExt = False
									Set objAccMgr = Nothing
									Exit Function
								End If
								Wait 1
							End If
							If Fn_SISW_UI_Object_Operations("Fn_AccMgr_AccessControlOperationExt","Enabled", JavaWindow("Default Teamcenter Window").Dialog("Warning"),SISW_MICRO_TIMEOUT) Then
								iWinId = JavaWindow("Default Teamcenter Window").Dialog("Warning").GetROProperty("window id")
								JavaWindow("Default Teamcenter Window").Dialog("Warning").SetTOProperty "window id", iWinId
								Do
									sLogText =  JavaWindow("Default Teamcenter Window").Dialog("Warning").GetROProperty("text")
									If objAccMgr.JavaDialog("Warning").JavaButton("OK").exist(5) Then
										Call Fn_Button_Click("Fn_AccMgr_AccessControlOperationExt", objAccMgr.JavaDialog("Warning"), "OK")
									ElseIf JavaWindow("Default Teamcenter Window").Dialog("Warning").WinButton("OK").exist(5) Then
										Call Fn_UI_WinButton_Click("Fn_AccMgr_AccessControlOperationExt",JavaWindow("Default Teamcenter Window").Dialog("Warning"),"OK","","","")
									End If
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Warning Dialog for Message["+CStr(sLogText)+"]  handled Successfully.")
								Loop While Fn_UI_ObjectExist("Fn_AccMgr_AccessControlOperationExt", JavaWindow("Default Teamcenter Window").Dialog("Warning")) = True
							ElseIf Fn_SISW_UI_Object_Operations("Fn_AccMgr_AccessControlOperationExt","Enabled", objAccMgr.JavaDialog("Warning"),SISW_MICRO_TIMEOUT) then
								Do
									objAccMgr.JavaDialog("Warning").JavaButton("OK").Click	
								Loop While Fn_UI_ObjectExist("Fn_AccMgr_AccessControlOperationExt", JavaWindow("Default Teamcenter Window").JavaWindow("AccMgrApplet").JavaDialog("Warning")) = True
							End If
						End If
						
					Case "ID of Accessor"
						'iColNum = 1
						iColNum = Fn_AccMgr_AccessControlOperationExt("GetColumnIndex", Trim(aACLProperties(iCounter)), "", "", "")
						If iColNum <> -1 Then
							objAccMgr.JavaTable("ACLTable").DoubleClickCell iRowIndex, iColNum
							If Err.Number <> 0 Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_ACLOperationsExt function failed as double click is not performed on table")
								Fn_AccMgr_AccessControlOperationExt = False
								Set objAccMgr = Nothing
								Exit Function
							End If
							Wait 3
							If Fn_UI_ObjectExist("Fn_AccMgr_AccessControlOperationExt", objAccMgr.JavaDialog("Select Accessor")) = True Then
								objAccMgr.JavaDialog("Select Accessor").JavaList("AccessorList").Select aACLPropertyValue(iCounter)
								Call Fn_Button_Click("Fn_AccMgr_AccessControlOperationExt",objAccMgr.JavaDialog("Select Accessor"), "OK")
								Wait 1
							End If
						End If
					Case "Read"
						iColNum = Fn_AccMgr_AccessControlOperationExt("GetColumnIndex", Trim(aACLProperties(iCounter)), "", "", "")
						If iColNum <> -1 Then
							objAccMgr.JavaTable("ACLTable").ClickCell "#"&CStr(iRowIndex),"#"&CStr(iColNum),"LEFT"
							If Err.Number <> 0 Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_ACLOperationsExt function failed as click Cell is not performed on table")
								Fn_AccMgr_AccessControlOperationExt = False
								Set objAccMgr = Nothing
								Exit Function
							End If
							Wait 1
							objAccMgr.JavaTable("ACLTable").DoubleClickCell "#"&CStr(iRowIndex),"#"&CStr(iColNum)
							If Err.Number <> 0 Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_ACLOperationsExt function failed as double click is not performed on table")
								Fn_AccMgr_AccessControlOperationExt = False
								Set objAccMgr = Nothing
								Exit Function
							End If
								If Fn_SISW_UI_Object_Operations("Fn_AccMgr_AccessControlOperationExt","Enabled", objAccMgr.JavaDialog("Warning"),SISW_MICRO_TIMEOUT) Then
									Do
										objAccMgr.JavaDialog("Warning").JavaButton("OK").Click	
									Loop While Fn_UI_ObjectExist("Fn_AccMgr_AccessControlOperationExt", objAccMgr.JavaDialog("Warning")) = True
								End If	
							wait 1
							objAccMgr.JavaList("ACLTableValues").Select aACLPropertyValue(iCounter)
							If Err.Number <> 0 Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_ACLOperationsExt function failed as not able to select node from List")
								Fn_AccMgr_AccessControlOperationExt = False
								Set objAccMgr = Nothing
								Exit Function
							End If
							Wait 2
						End If
						If Fn_SISW_UI_Object_Operations("Fn_AccMgr_AccessControlOperationExt","Enabled", objAccMgr.JavaDialog("Warning"),SISW_MICRO_TIMEOUT) Then
							Do
								objAccMgr.JavaDialog("Warning").JavaButton("OK").Click	
							Loop While Fn_UI_ObjectExist("Fn_AccMgr_AccessControlOperationExt", objAccMgr.JavaDialog("Warning")) = True
						End If	
					Case Else
						iColNum = Fn_AccMgr_AccessControlOperationExt("GetColumnIndex", Trim(aACLProperties(iCounter)), "", "", "")
						If iColNum <> -1 Then
							objAccMgr.JavaTable("ACLTable").SetCellData "#" & CStr(iRowIndex),"#"&CStr(iColNum), aACLPropertyValue(iCounter)
							If Err.Number <> 0 Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_ACLOperationsExt function failed as not able to set Cell Data in table")
								Fn_AccMgr_AccessControlOperationExt = False
								Set objAccMgr = Nothing
								Exit Function
							End If
							wait 2
						End If
									
				End Select
			Next
			'Click on Save button
			Wait 1
			Call Fn_Button_Click("Fn_AccMgr_AccessControlOperationExt", objAccMgr, "Save_ACLTable")
			Wait 2
			Fn_AccMgr_AccessControlOperationExt = True
			Set objAccMgr = Nothing
			
		'Case to Delete ACL Rule Row
		Case "Delete"
			If iRowIndex >= objAccMgr.JavaTable("ACLTable").GetROProperty("rows") Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_AccMgr_AccessControlOperationExt function failed as the Row does not exist")
				Fn_AccMgr_AccessControlOperationExt = FALSE
				Set objAccMgr = Nothing
				Exit Function						
			End If
			objAccMgr.JavaTable("ACLTable").SelectRow iRowIndex
			If Err.Number <> 0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_ACLOperationsExt function failed as select row is not performed on table")
				Fn_AccMgr_AccessControlOperationExt = False
				Set objAccMgr = Nothing
				Exit Function
			End If
			'Click on Remove button
			Call Fn_Button_Click("Fn_AccMgr_AccessControlOperationExt", objAccMgr, "Remove_ACLTable")
			'Click on save button
			Fn_AccMgr_AccessControlOperationExt = Fn_Button_Click("Fn_AccMgr_AccessControlOperationExt", objAccMgr, "Save_ACLTable")
			Set objAccMgr = Nothing
		
		'Case to get all Column values for specific Row
		Case "GetList"
			If iRowIndex >= objAccMgr.JavaTable("ACLTable").GetROProperty("rows") Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_AccMgr_AccessControlOperationExt function failed as the Row does not exist")
				Fn_AccMgr_AccessControlOperationExt = FALSE
				Set objAccMgr = Nothing
				Exit Function						
			End If
			'Select given row from ACL Table
			objAccMgr.JavaTable("ACLTable").SelectRow iRowIndex
			If Err.Number <> 0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_ACLOperationsExt function failed as select row is not performed on table")
				Fn_AccMgr_AccessControlOperationExt = False
				Set objAccMgr = Nothing
				Exit Function
			End If
			'Count number of cols
			iTotalCol = objAccMgr.JavaTable("ACLTable").GetROProperty("cols")
			ReDim aACLPropertyValue(iTotalCol-1)
			
			For iCounter = 0 To iTotalCol - 1
				If iCounter > 1 Then
					objAccMgr.JavaTable("ACLTable").DoubleClickCell "#"+CStr(iRowIndex),"#"+CStr(iCounter),"LEFT","NONE"
					If Err.Number <> 0 Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_ACLOperationsExt function failed as Double Click is not performed on table")
						Fn_AccMgr_AccessControlOperationExt = False
						Set objAccMgr = Nothing
						Exit Function
					End If
					aACLPropertyValue(iCounter) = objAccMgr.JavaList("ACLTableValues").GetItem(objAccMgr.JavaList("ACLTableValues").Object.getSelectedIndex)
				Else
					objAccMgr.JavaTable("ACLTable").ClickCell iRowIndex, iCounter, "LEFT"
					If Err.Number <> 0 Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_ACLOperationsExt function failed as Click Cell is not performed on table")
						Fn_AccMgr_AccessControlOperationExt = False
						Set objAccMgr = Nothing
						Exit Function
					End If					
					aACLPropertyValue(iCounter) = objAccMgr.JavaTable("ACLTable").GetCellData("#"+CStr(iRowIndex),"#"+CStr(iCounter))
				End If
			Next
			Fn_AccMgr_AccessControlOperationExt = aACLPropertyValue
			'Step is done to defocus the cell
			If objAccMgr.JavaTable("ACLTable").GetROproperty("rows")>0 Then
				objAccMgr.JavaTable("ACLTable").ClickCell 0, 0, "LEFT"
				If Err.Number <> 0 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_ACLOperationsExt function failed as Click Cell is not performed on table")
					Fn_AccMgr_AccessControlOperationExt = False
					Set objAccMgr = Nothing
					Exit Function
				End If
			End If
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sActionType&" case successfully completed of function Fn_AccMgr_AccessControlOperationExt")
			Set objAccMgr = Nothing
			
		'Case to get Row number for specific Column value
		Case "GetRowForColumnValue"
			
			iColNum = Fn_AccMgr_AccessControlOperationExt("GetColumnIndex", Trim(sACLProperties), "", "", "")
			iTotalRow = cInt(objAccMgr.JavaTable("ACLTable").GetROProperty("rows"))
			For iCount = 0 To iTotalRow - 1
			
				If iColNum > 1 Then
					
					objAccMgr.JavaTable("ACLTable").DoubleClickCell "#"+CStr(iCount),"#"+CStr(iColNum),"LEFT","NONE"
					If Err.Number <> 0 Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_ACLOperationsExt function failed as Double Click is not performed on table")
						Fn_AccMgr_AccessControlOperationExt = False
						Set objAccMgr = Nothing
						Exit Function
					End If
							
					'Added code to handle the warning msg -----
					If objAccMgr.JavaDialog("Warning").JavaButton("OK").Exist(1) Then		
						Call Fn_Button_Click("Fn_AccMgr_AccessControlOperationExt",objAccMgr.JavaDialog("Warning"),"OK")
					ElseIf objAccMgr.JavaButton("OK").Exist(1) Then
						Call Fn_Button_Click("Fn_AccMgr_AccessControlOperationExt",objAccMgr,"OK")
					End If
										
					If sACLPropertyValue = objAccMgr.JavaList("ACLTableValues").GetItem(objAccMgr.JavaList("ACLTableValues").Object.getSelectedIndex) Then
						Fn_AccMgr_AccessControlOperationExt = iCount
	                    Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sACLPropertyValue&" is at "&iCount&" row in ACL Table.")
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sActionType&" case sucessfully completed of function Fn_AccMgr_AccessControlOperationExt")
						Set objAccMgr = Nothing
						Exit Function										
					End If
				Else
					If sACLPropertyValue = objAccMgr.JavaTable("ACLTable").GetCellData("#"+CStr(iCount),"#"+CStr(iColNum)) Then		
						Fn_AccMgr_AccessControlOperationExt = iCount
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sACLPropertyValue&" is at "&iCount&" row in ACL Table.")
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sActionType&" case sucessfully completed of function Fn_AccMgr_AccessControlOperationExt")
						Set objAccMgr = Nothing
						Exit Function
					End If
				End If
			Next
			
		Case "Verify"
			If objAccMgr.JavaTable("ACLTable").GetROProperty("rows") = 0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_AccMgr_AccessControlOperationExt function failed as the Row does not exist")
				Fn_AccMgr_AccessControlOperationExt = FALSE
				Set objAccMgr = Nothing
				Exit Function						
			End If
					
			For iCount = 0 To iTotalRow - 1
				For iCounter = 0 To iTotalCol - 1
					If iCounter > 1 Then
						objAccMgr.JavaTable("ACLTable").DoubleClickCell "#"+CStr(iCount),"#"+CStr(iCounter),"LEFT","NONE"
						If Err.Number <> 0 Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_ACLOperationsExt function failed as Double Click is not performed on table")
							Fn_AccMgr_AccessControlOperationExt = False
							Set objAccMgr = Nothing
							Exit Function
						End If
						
						'Added code to handle the warning msg -----
						If objAccMgr.JavaDialog("Warning").JavaButton("OK").Exist(1) Then		
							Call Fn_Button_Click("Fn_AccMgr_AccessControlOperationExt",objAccMgr.JavaDialog("Warning"),"OK")
						ElseIf objAccMgr.JavaButton("OK").Exist(1) Then
							Call Fn_Button_Click("Fn_AccMgr_AccessControlOperationExt",objAccMgr,"OK")
						End If
									
						If sACLPropertyValue = objAccMgr.JavaList("ACLTableValues").GetItem(objAccMgr.JavaList("ACLTableValues").Object.getSelectedIndex) Then
'							If iCount = 0 Then	' ------------in case the column value found at 0th row
'								Fn_AccMgr_AccessControlOperationExt = "#0"
'							Else
'								Fn_AccMgr_AccessControlOperationExt = iCount
'							End If
							Fn_AccMgr_AccessControlOperationExt = iCount
							
                            Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sACLPropertyValue&" is at "&iCount&" row in ACL Table.")
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sActionType&" case sucessfully completed of function Fn_AccMgr_AccessControlOperationExt")
							Set objAccMgr = Nothing
							Exit Function										
						End If
					Else
						objAccMgr.JavaTable("ACLTable").ClickCell iCount, iCounter, "LEFT"
						If Err.Number <> 0 Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_ACLOperationsExt function failed as Click Cell is not performed on table")
							Fn_AccMgr_AccessControlOperationExt = False
							Set objAccMgr = Nothing
							Exit Function
						End If
						If sACLPropertyValue = objAccMgr.JavaTable("ACLTable").GetCellData("#"+CStr(iCount),"#"+CStr(iCounter)) Then		
'							If iCount = 0 Then	' ------------in case the column value found at 0th row
'								Fn_AccMgr_AccessControlOperationExt = "#0"
'							Else
'								Fn_AccMgr_AccessControlOperationExt = iCount
'							End If
							Fn_AccMgr_AccessControlOperationExt = iCount
							
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sACLPropertyValue&" is at "&iCount&" row in ACL Table.")
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sActionType&" case sucessfully completed of function Fn_AccMgr_AccessControlOperationExt")
							Set objAccMgr = Nothing
							Exit Function
						End If
					End If
				Next
			Next
	End Select
End Function
