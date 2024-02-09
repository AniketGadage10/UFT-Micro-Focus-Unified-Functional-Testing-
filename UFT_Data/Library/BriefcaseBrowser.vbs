Option Explicit
SISW_DEFAULT_TIMEOUT = 240
SISW_DEFAULT_WAIT = 5

'1.Fn_SISW_BB_GetObject()
'2. Fn_BB_InvokeBriefcaseBrowser()
'3. Fn_BB_ResetBriefcaseBrowser()
'4. Fn_BB-UpdateCutomDatasetMappingXMLNode(XMLDataFile, sTagName,sAttributeName,sAttributeVal,sSubAttributeName,sNewAttributeVal,sReserved)
'5. Fn_BB_ReadyStatusSync(iIterations)
'6. Fn_BB_OpenAssembly(sAction,sPath)
'7. Fn_BB_BriefcaseBrowserTree_Opearation(sAction,sTabName,sNode,sReserve)
'8. Fn_BB_CreateDataset(sAction,sPath,sRelationName,sButton,sReserved)
'9. Fn_BB_DataSetTab_Opeartions(sAction,sNode,sColType,sColName,sColValue,sReserve)
'10. Fn_BB_TreeGetItemPath(sFunctionName,sAction,ObjTree,StrNode,sType)
'11. Fn_BB_BriefcaseBrowserExit(sAction,sPath,sReservered)
'12. Fn_BB_BriefcaseBrowserSave(sFilename,sReservered)
'13. Fn_BB_PrefrenceOpeartion(sAction,dicPreference)
'14. Fn_BB_JavaTab_Operations(sAction, objJavaDialog, sTabObjectName, sTabName,sPopupMenu,sReserved)
'15. Fn_BB_LoadBBXML()
'16. Fn_BB_ExtractContentBriefcaseBrowser(sFileName,sReserved)
'17. Fn_BB_VerifyBriefcaseBrowserContent(sFilePath,sFileName,sReserved)
'18. Fn_BriefcaseBrowser_TestcaseExit(bTcKill)
'19. Fn_BB_ToolBar_Opeartions(sAction,objBB,sButtonName,sReserved)
'20. Fn_BB_ItemPropertyPanelVerify(sPropName,sValue)
'21. Fn_BB_ErrorLog_Opeartions()
'22. Fn_BB_TabOperation()

'****************************************    Function to get Object hierarchy ***************************************
'Function Name		 	:	Fn_SISW_PSE_GetObject
'
'Description		    :  	Function to get Object hierarchy
'
'Parameters		    	:	1. sObjectName : Object Handle name
'
'Return Value		    :  	Object \ Nothing
'
'Examples		     	:	Fn_SISW_BB_GetObject("BBWindow")
'
'History:
'	Developer Name			Date			Rev. No.		Reviewer		Changes Done	
'----------------------------------------------------------------------------------------------------------------------------------
'	Shweta Rathod		 19-Jul-2016		1.0				Shweta
'----------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_BB_GetObject(sObjectName)
	Dim sObjectXMLPath
	sObjectXMLPath = Fn_GetEnvValue("User", "AutomationDir") & "\TestData\AutomationXML\ObjectXML\BriefcaseBrowser.xml"
	Set Fn_SISW_BB_GetObject = Fn_SISW_Setup_GetObjectFromXML(sObjectXMLPath, sObjectName)
End Function


'****************************************    Function to get Object hierarchy ***************************************
'Function Name		 	:	Fn_BB_InvokeBriefcaseBrowser
'
'Description		    :  	Function to Invoke BriefcaseBrowser application
'
'Return Value		    :  	TRUE \ FALSE
'
'Examples		     	:	Fn_SISW_BB_GetObject("BBWindow")
'
'History:
'	Developer Name			Date			Rev. No.		Reviewer		Changes Done	
'----------------------------------------------------------------------------------------------------------------------------------
'	Shweta Rathod		 19-Jul-2016		1.0				Shweta
'----------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_BB_InvokeBriefcaseBrowser()																																		 
   On Error resume next
	Dim sPath,sAutoDir
	SystemUtil.Run Environment.Value("BriefcaseBrowserPath")
	If JavaWindow("BriefcaseBrowser").Exist(240) Then						  							
		Fn_BB_InvokeBriefcaseBrowser = TRUE  																						
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Invoked Teamcenter Application from [" + Environment.Value("BriefcaseBrowserPath") + "]")
	Else
		 Fn_BB_InvokeBriefcaseBrowser = FALSE
		 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Invoke Teamcenter Application from [" + Environment.Value("BriefcaseBrowserPath") + "]")
		 Exit Function
	End If
End Function


'****************************************    Function to get Object hierarchy ***************************************
'Function Name		 	:	Fn_BB_ResetBriefcaseBrowser
'
'Description		    :  	Function to Fn_Reset Briefcase Browser application
'
'Return Value		    :  	TRUE \ FALSE
'
'Examples		     	:	Fn_BB_ResetBriefcaseBrowser()
'
'History:
'	Developer Name			Date			Rev. No.		Reviewer		Changes Done	
'----------------------------------------------------------------------------------------------------------------------------------
'	Shweta Rathod		 19-Jul-2016		1.0				Shweta
'----------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_BB_ResetBriefcaseBrowser()	
	Dim sMenu,bRet,objReset
	Fn_BB_ResetBriefcaseBrowser = false 
	set objReset = Fn_SISW_BB_GetObject("BBResetPerspective")
	Fn_BB_ResetBriefcaseBrowser = Fn_UI_ObjectExist("Fn_BB_ResetBriefcaseBrowser",objReset)
	If Fn_BB_ResetBriefcaseBrowser = false then
		sMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("BriefcaseBrowser_Menu"), "WindowResetPerspective")
		bRet = Fn_UI_JavaMenu_Select("Fn_MenuOperation",JavaWindow("BriefcaseBrowser"), sMenu)
		wait 1
		If bRet = false then
			 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to perform menu opeartion [" + sMenu + "]")
			 Exit Function
		End IF
		wait 3
		Fn_BB_ResetBriefcaseBrowser = Fn_UI_ObjectExist("Fn_BB_ResetBriefcaseBrowser",objReset)
		If Fn_BB_ResetBriefcaseBrowser = false then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: prefrence window does not exist.")	
			exit function
		End if		
	End if
	wait 1
	Fn_BB_ResetBriefcaseBrowser = Fn_Button_Click("Fn_BB_ResetBriefcaseBrowser",objReset,"Yes")
	If Fn_BB_ResetBriefcaseBrowser = false then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: falied to reset briefcase browser perspective.")	
		exit function
	End if
	wait 1
	Fn_BB_ResetBriefcaseBrowser = true
	Set objReset = Nothing
End Function

'****************************************    Function to get Object hierarchy ***************************************
'Function Name		 	:	Fn_BB_UpdateCutomDatasetMappingXMLNode
'
'Description		    :  	Function to update attributes of CutomDatasetMappingXML Node.
'
'Parameters		    	:	1. XMLDataFile : File name
'							2. sTagName : Tag name of the xml file
'							3. sAttributeName : attribute name to select data 
'							4. sAttributeVal : search Value to be modfied
'							5. sSubAttributeName : Sub attribute name 
'							6. sNewAttributeVal : Value to be modified 
'							7. sReserved : future use
'
'Return Value		    :  	TRUE 
'
'Examples		     	:	Fn_BB_UpdateCutomDatasetMappingXMLNode("C:\Tc11.2.3_2016060800_BB_Configured\bbworkspace\configurations\Unman\CustomDatasetMappings.xml","data_set_mapping","extension","docx","relation","XMLNewAttributeVal","")
'
'History:
'	Developer Name			Date			Rev. No.		Reviewer		Changes Done	
'----------------------------------------------------------------------------------------------------------------------------------
'	Shweta Rathod		 19-Jul-2016		1.0				Shweta
'----------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_BB_UpdateCutomDatasetMappingXMLNode(XMLDataFile, sTagName,sAttributeName,sAttributeVal,sSubAttributeName,sNewAttributeVal,sReserved)
	Dim objXMLDoc,objXMLNodeList,numObjXMLNodeList,i,sXMLAttribute
	set objXMLDoc = CreateObject("Microsoft.XMLDOM")
	objXMLDoc.load(XMLDataFile)
	Set objXMLNodeList = objXMLDoc.getElementsByTagName(sTagName)
	numObjXMLNodeList = objXMLNodeList.length
	For i = 0 to numObjXMLNodeList - 1
		sXMLAttributeVal = objXMLNodeList.item(i).getAttribute(sAttributeName)
		If sXMLAttributeVal = sAttributeVal then exit for	
	next
	Set objChildNodes = objXMLNodeList.item(i).childNodes
	numObjXMLNodeList = objChildNodes.length
	For i = 0 to numObjXMLNodeList - 1
		If objChildNodes.item(i).getAttribute(sSubAttributeName) <> sNewAttributeVal Then
			objChildNodes.item(i).setAttribute sSubAttributeName,sNewAttributeVal
		End If		
	next	
	objXMLDoc.Save(XMLDataFile)
	Set objXMLDoc = nothing 
	Set objXMLNodeList = nothing
	Fn_BB_UpdateCutomDatasetMappingXMLNode = True	
End Function


'*********************************************************		Function to Synchronize on Application Response	***********************************************************************
'Function Name		:				Fn_BB_ReadyStatusSync(iIterations)

'Description			 :		 		 This function waits till Application comes to Ready state

'Parameters			   :	 			1. iIterations: No. of times to be checked for Ready text
											
'Return Value		   : 				TRUE \ FALSE

'Pre-requisite			:		 		Briefcase browser application should be displayed

'Examples				:				 Fn_BB_ReadyStatusSync(2)

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Shweta Rathod			25-Jul-2016			1.0
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_BB_ReadyStatusSync(iIterations)
	Dim iCounter, bFound, iCnt, objQSearchEdit
	Fn_BB_ReadyStatusSync =  false
	JavaWindow("BriefcaseBrowser").JavaStaticText("ReadyStatus").SetTOProperty "label","Open \(\Finished at.*"
	For iCounter = 1 to iIterations
		If JavaWindow("BriefcaseBrowser").Exist(SISW_DEFAULT_TIMEOUT) Then
			JavaWindow("BriefcaseBrowser").JavaStaticText("ReadyStatus").WaitProperty "label", "ReadyStatus", 20000
			If JavaWindow("BriefcaseBrowser").JavaStaticText("ReadyStatus").Exist(1) Then
				exit for
			End If
		Else
			Fn_BB_ReadyStatusSync = FALSE
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Teamcenter window does not exist.")	
			exit function
		End If
	Next
	If JavaWindow("BriefcaseBrowser").JavaStaticText("ReadyStatus").Exist(1) Then
		For iCounter = 1 to iIterations
			For iCnt = 1 to 20
				wait 1
				If JavaWindow("BriefcaseBrowser").Exist(SISW_DEFAULT_TIMEOUT) Then
					' exit from inner loop if progressbar disappears
					If JavaWindow("BriefcaseBrowser").JavaObject("ProgressBar").exist(1) = FALSE Then
						Fn_BB_ReadyStatusSync = TRUE
						Exit for
					End If
				Else
					Fn_BB_ReadyStatusSync = FALSE
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: BriefcaseBrowser window does not exist.")	
					exit function
				End If
			Next
			' exit from main loop if progressbar disappears
			If Fn_BB_ReadyStatusSync Then Exit for
		Next
	end if
	If JavaWindow("BriefcaseBrowser").JavaStaticText("ReadyStatus").Exist(1) = FALSE OR Fn_BB_ReadyStatusSync = FALSE Then
		Fn_BB_ReadyStatusSync = FALSE
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: BriefcaseBrowser Not Ready after [" + CStr(iIterations) + "] sync iterations")		
	Else
		Fn_BB_ReadyStatusSync = TRUE
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: BriefcaseBrowser is Ready in [" + CStr(iIterations) + "] sync iterations")		
	End If
End Function


'*********************************************************		Function to Synchronize on Application Response	***********************************************************************
'Function Name		:				Fn_BB_OpenAssembly(sAction,sPath)

'Description			 :		 		 This function open the assembly from the given path

'Parameters			   :	 			1. sAction : type of assembly to be open CAD or BB 
											
'Return Value		   : 				TRUE \ FALSE

'Pre-requisite			:		 		Brief case browser application should be displayed

'Examples				:				 Fn_BB_OpenAssembly("OpenCAD","C:\mainline\Reports\NX\Add_Attachs_Default_Relat_88663\Top88663_A.prt")

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Shweta Rathod			25-Jul-2016			1.0
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_BB_OpenAssembly(sAction,sPath)
	Dim bRet,objOpen,sMenu
	Set objOpen =Dialog("Save")
	Fn_BB_OpenAssembly = FALSE
	
	Select Case sAction	
		Case "OpenCAD"
			objOpen.SetTOProperty "text","Select the CAD file."
			if not objOpen.Exist(1) then
				sMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("BriefcaseBrowser_Menu"), "OpenCADAssembly")
				bRet = Fn_UI_JavaMenu_Select("Fn_MenuOperation",JavaWindow("BriefcaseBrowser"), sMenu)
				wait 1
				If bRet = false then
					  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Invoke perform menu opeartion [" + sMenu + "]")
					  Fn_BB_OpenAssembly = false
					 Exit Function
				End IF
'				elseif dicOpenAss("ToolBarMenu") then
'					bRet = Fn_SISW_UI_JavaToolbar_Operations("Fn_BB_OpenAssembly", "Click", objOpen, "ToolBar", "", dicOpenAss("ToolBarMenu"), sMenu, "")
'				End if
			End if	
			
		Case "OpenBB"
			objOpen.SetTOProperty "text", "Select the file representing the briefcase"
			if not objOpen.Exist(10) then
				sMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("BriefcaseBrowser_Menu"), "OpenBriefcase")
				bRet = Fn_UI_JavaMenu_Select("Fn_MenuOperation",JavaWindow("BriefcaseBrowser"), sMenu)
				If bRet = false then
					 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Invoke perform menu opeartion [" + sMenu + "]")
					 Fn_BB_OpenAssembly = false
					 Exit Function
				End IF
'				elseif dicOpenAss("ToolBarMenu") then
'					bRet = Fn_SISW_UI_JavaToolbar_Operations("Fn_BB_OpenAssembly", "Click", objOpen, "ToolBar", "",dicOpenAss("ToolBarMenu"), sMenu, "")
'				End if	
			End if
			
		Case else
			'future use
	End Select
	
	If objOpen.Exist(10) Then
		Err.clear
		objOpen.WinEdit("Name").Set sPath
		wait 1
		If Err.Number < 0 Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to set the value in edit box " + objOpen.WinEdit("Name").ToString())
			Fn_BB_OpenAssembly = false
			Exit Function
		End if
		wait 1
		Err.clear
		objOpen.WinButton("Open").Click micLeftBtn
		If Err.Number < 0 Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to click on button " + objOpen.WinButton("Open").ToString())
			Fn_BB_OpenAssembly = false
			Exit Function
		End if
		wait 1		
	else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " dialog " + objOpen.ToString()+" does not exist.")
		Fn_BB_OpenAssembly = false
		Exit Function
	End If	
	Set objOpen = Nothing
	Fn_BB_OpenAssembly = True
End Function


'*********************************************************		Function to Synchronize on Application Response	***********************************************************************
'Function Name		:				Fn_BB_BriefcaseBrowserTree_Opearation(sAction,sTabName,sNode,sReserve)

'Description			 :		 		 This function perform variou operations on Briefcase browser tree

'Parameters			   :	 			1. sAction : type of assembly to be open CAD or BB 
'										2. sTabName: name of the tab under BB tree is displaying on which going to perform the operation
'										3. sNode: full path of the node on which performing the opeation
'										4. sReserve : for future use
											
'Return Value		   : 				TRUE \ FALSE

'Pre-requisite			:		 		Brief case browser application should be displayed

'Examples				:				 Fn_BB_BriefcaseBrowserTree_Opearation("select","*C:\mainline\Reports\NX\Add_Attachs_Default_Relat_88663\Top88663_A.prt","Top88663:child188663","")

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Shweta Rathod			25-Jul-2016			1.0
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_BB_BriefcaseBrowserTree_Opearation(sAction,sTabName,sNode,sReserve)
	Dim bRet,objTab,iRow
	Fn_BB_BriefcaseBrowserTree_Opearation = false
	set objTab = Fn_SISW_BB_GetObject("BriefcaseBrowser")
	Set objTree = objTab.JavaTree("BBTree_VisRel_CTabFolder")
	set objTab = objTab.JavaTab("CTabFolder")
	Fn_BB_BriefcaseBrowserTree_Opearation = Fn_UI_Object_SetTOProperty_ExistCheck("Fn_BB_BriefcaseBrowserTree_Opearation",objTab,"value",sTabName)
	If Fn_BB_BriefcaseBrowserTree_Opearation = true then
		Fn_BB_BriefcaseBrowserTree_Opearation = Fn_UI_JavaTab_Select("Fn_BB_BriefcaseBrowserTree_Opearation",JavaWindow("BriefcaseBrowser"),"CTabFolder", sTabName)
		wait 1
		If Fn_BB_BriefcaseBrowserTree_Opearation = false Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select tab  " + objTab.ToString())
			Exit Function
		End if		
	else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "tab does not exist " + objTab.ToString())
		Exit Function
	End if
	Select Case lcase(sAction)
		Case "select"
			Fn_BB_BriefcaseBrowserTree_Opearation = Fn_JavaTree_Select("Fn_BB_BriefcaseBrowserTree_Opearation", JavaWindow("BriefcaseBrowser"), "BBTree_VisRel_CTabFolder",sNode)
			If Fn_BB_BriefcaseBrowserTree_Opearation = false Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select node  " + sNode +" in javatree " + JavaWindow("BriefcaseBrowser").JavaTree("BBTree").ToString())
				Exit Function
			End if	
		Case "exist"
			Fn_BB_BriefcaseBrowserTree_Opearation = Fn_UI_JavaTree_NodeExist("Fn_BB_BriefcaseBrowserTree_Opearation",objTree,sNode)
			If Fn_BB_BriefcaseBrowserTree_Opearation = false Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to exist node  " + sNode +" in javatree " + JavaWindow("BriefcaseBrowser").JavaTree("BBTree_VisRel_CTabFolder").ToString())
				Exit Function
			End if
		Case "expand"
			Fn_BB_BriefcaseBrowserTree_Opearation = Fn_UI_JavaTree_Expand("Fn_BB_BriefcaseBrowserTree_Opearation",JavaWindow("BriefcaseBrowser"),"BBTree",sNode)
			If Fn_BB_BriefcaseBrowserTree_Opearation = false Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to expand node  " + dicOpenAssm("Node") +" in javatree " + JavaWindow("BriefcaseBrowser").JavaTree("BBTree").ToString())
				Exit Function
			End if
		Case "collapse"		
			Fn_BB_BriefcaseBrowserTree_Opearation = Fn_UI_JavaTree_Collapse("Fn_BB_BriefcaseBrowserTree_Opearation",JavaWindow("BriefcaseBrowser"),"BBTree",sNode)
			If Fn_BB_BriefcaseBrowserTree_Opearation = false Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to collapse node  " + dicOpenAssm("Node") +" in javatree " + JavaWindow("BriefcaseBrowser").JavaTree("BBTree").ToString())
				Exit Function
			End if
		Case "isempty"
			iRow = objTree.Object.getItemCount()
			If iRow <> 0 Then
				If Fn_BB_BriefcaseBrowserTree_Opearation = false Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to verify tree is empty" + JavaWindow("BriefcaseBrowser").JavaTree("BBTree").ToString())
				Exit Function
			End if
			End If
		
	End select
	Fn_BB_BriefcaseBrowserTree_Opearation = true
	set objTab = Nothing
	set objTree = Nothing
End Function

'*********************************************************		Function to Synchronize on Application Response	***********************************************************************
'Function Name		:				Fn_BB_BriefcaseBrowserTree_Opearation(sAction,sTabName,sNode,sReserve)

'Description			 :		 		 This function perform variou operations on Briefcase browser tree

'Parameters			   :	 			1. sAction : type of assembly to be open CAD or BB 
'										2. sTabName: name of the tab under BB tree is displaying on which going to perform the operation
'										3. sNode: full path of the node on which performing the opeation
'										4. sReserve : for future use
											
'Return Value		   : 				TRUE \ FALSE

'Pre-requisite			:		 		Node on which creating dataset should be selected

'Examples				:				 Fn_BB_CreateDataset("verify_defaultrelname","C:\mainline\Scripts\REG-BriefcaseBrowser\Add_Attachs_Default_Relation_Type\DSWord.docx","rendering","","")
' 										Fn_BB_CreateDataset("create","C:\mainline\Scripts\REG-BriefcaseBrowser\Add_Attachs_Default_Relation_Type\DSWord.docx","rendering","OK","")

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Shweta Rathod			25-Jul-2016			1.0
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_BB_CreateDataset(sAction,sPath,sRelationName,sButton,sReserved)
	Dim sRelName,sMenu
	Dim objBB,objOpen,objAddDt
	set objBB = Fn_SISW_BB_GetObject("BriefcaseBrowser")
	Set objOpen = Fn_SISW_BB_GetObject("OpenDatasetFile")
	set objAddDataset = Fn_SISW_BB_GetObject("AddDataset")
	
	Fn_BB_CreateDataset = false
	
	if sPath <> "" then
		objOpen.SetTOProperty "text","Select a file to attach to the selected CAD part or assembly" 
'		Fn_BB_CreateDataset = Fn_UI_Object_SetTOProperty_ExistCheck("Fn_BB_CreateDataset",objOpen,"text","Select a file to attach to the selected CAD part or assembly") 
'		Fn_BB_CreateDataset = Fn_UI_ObjectExist("Fn_BB_CreateDataset",objOpen) 
		If  objOpen.exist(1) = false Then
			Fn_BB_CreateDataset = Fn_UI_JavaMenu_Select("Fn_BB_CreateDataset",objBB, Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("BriefcaseBrowser_Menu"), "AddDataset"))
			If Fn_BB_CreateDataset = false Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to invoke menu  " + objTab.ToString())
				Exit Function
			End if
			wait 3
		End if	
		
		Fn_BB_CreateDataset = Fn_UI_ObjectExist("Fn_BB_CreateDataset",objOpen) 
		If Fn_BB_CreateDataset = true Then				
			Err.clear
			objOpen.WinEdit("Name").Set sPath
			If Err.Number < 0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to set the value in edit box " + objOpen.WinEdit("Name").ToString())
				Exit Function
			End if
			wait 1
			Err.clear
			objOpen.WinButton("Open").Click micLeftBtn
			If Err.Number < 0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to click on button " + objOpen.WinButton("Open").ToString())
				Exit Function
			End if	
			wait 1			
		End If
 	End if	
	
	Select Case lcase(sAction)
		Case "create"
			If sRelationName <> "" then
				Fn_BB_CreateDataset = Fn_SISW_UI_JavaList_Operations("Fn_BB_CreateDataset", "Select", objAddDataset, "RelationshipName", sRelationName, "", "")
				If Fn_BB_CreateDataset = false then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select item ["+sRelationName+"] from javalist " + objAddDataset.JavaList("RelationshipName").ToString())
					Exit Function
				End if
			End if
			wait 1
		Case "verify_defaultrelname"
			sRelName = Fn_UI_Object_GetROProperty("Fn_BB_CreateDataset",objAddDataset.JavaList("RelationshipName"),"text")
			If SRelName <> sRelationName then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to retrive text from javalist " + objAddDataset.JavaList("RelationshipName").ToString())
				Fn_BB_CreateDataset = false
				Exit Function
			End if	
			wait 1
	End select
	
	If sButton <> "" Then
		Fn_BB_CreateDataset = Fn_Button_Click("Fn_BB_CreateDataset", objAddDataset, sButton)
		If Fn_BB_CreateDataset = false Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to click on button " + objOpen.WinButton("Open").ToString())
			Exit Function
		End if
	End If
	wait 1
	Fn_BB_CreateDataset = True
	set objBB = nothing
	Set objOpen = nothing
	set objAddDataset = nothing
End Function

'*********************************************************		Function to Synchronize on Application Response	***********************************************************************
'Function Name		:				Fn_BB_DataSetTab_Opeartions(sAction,sNode,sColType,sColName,sColValue,sReserve)

'Description			 :		 		 This function perform variou operations on dataset tab and tree

'Parameters			   :	 			1. sAction : type of assembly to be open CAD or BB 
'										2. sNode: name of the tab under BB tree is displaying on which going to perform the operation
'										3. sColType: value of column "Type" which is displaying in the table.
'										4. sColName : name of the coulm from which needs to retrive/verify the value
'										5. sColValue : value to be verified against the column name
'										6. sReserve : for future use
'
'Return Value		   : 				TRUE \ FALSE

'Pre-requisite			:		 		Node should be selected on which dataset is created

'Examples				:				 Fn_BB_DataSetTab_Opeartions("select","DS1Word/A","MSWord","DS1Word/A","","")
'										 Fn_BB_DataSetTab_Opeartions("Expand","DS1Word/A","MSWord","DS1Word/A","","")
'										 Fn_BB_DataSetTab_Opeartions("exist","DS1Word/A:DSWord.docx","MSWord","","","")
'										 Fn_BB_DataSetTab_Opeartions("verify_columnval","DS1Word/A","MSWord","Relationship Name","rendering","")

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Shweta Rathod			25-Jul-2016			1.0
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_BB_DataSetTab_Opeartions(sAction,sNode,sColType,sColName,sColValue,sReserve)
	Dim bRet,objTab
	Dim iIndex,iRow,iCnt,sApp,sAppNode
	
	Fn_BB_DataSetTab_Opeartions = false
	
	set objBB = Fn_SISW_BB_GetObject("BriefcaseBrowser")
	set objTab = objBB.JavaTab("CTabFolder")
	Set objTree = objBB.JavaTree("BBTree_VisRel_CTabFolder")
	
	objTab.SetTOProperty "value","Datasets" 
	Fn_BB_DataSetTab_Opeartions = Fn_UI_ObjectExist("Fn_BB_DataSetTab_Opeartions",objTab)
	If Fn_BB_DataSetTab_Opeartions = true then
		Fn_BB_DataSetTab_Opeartions = Fn_UI_JavaTab_Select("Fn_BB_DataSetTab_Opeartions",objBB,"CTabFolder", "Datasets")
		If Fn_BB_DataSetTab_Opeartions = false Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select tab  " + objTab.ToString())
			Exit Function
		End if
		wait 1		
	else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "tab does not exist " + objTab.ToString())
		Exit Function
	End if
	
	Select Case lcase(sAction)
		Case "select"
			Fn_BB_DataSetTab_Opeartions = Fn_BB_TreeGetItemPath("Fn_BB_DataSetTab_Opeartions","Datasets",ObjTree,sNode,sColType)
			err.clear
			JavaWindow("BriefcaseBrowser").JavaTree("BBTree_VisRel_CTabFolder").Select Fn_BB_DataSetTab_Opeartions
			wait 1
			If Err.Number < 0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select node  " + sNode +" in javatree " + objBB.JavaTree("BBTree").ToString())
				Exit Function
			End if	
		Case "doubleclick"
			Fn_BB_DataSetTab_Opeartions = Fn_BB_TreeGetItemPath("Fn_BB_DataSetTab_Opeartions","Datasets",ObjTree,sNode,sColType)
			err.clear
			JavaWindow("BriefcaseBrowser").JavaTree("BBTree_VisRel_CTabFolder").Select Fn_BB_DataSetTab_Opeartions
			wait 1
			Call Fn_KeyBoardOperation("SendKeys", "{ENTER}")
			wait 1
			If Err.Number < 0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Double click node  " + sNode +" in javatree " + objBB.JavaTree("BBTree_VisRel_CTabFolder").ToString())
				Exit Function
			End if	
		Case "exist"
			Fn_BB_DataSetTab_Opeartions = Fn_BB_TreeGetItemPath("Fn_BB_DataSetTab_Opeartions","Datasets",ObjTree,sNode,sColType)
			If Fn_BB_DataSetTab_Opeartions = false Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select node  " + sNode +" in javatree " + objBB.JavaTree("BBTree_VisRel_CTabFolder").ToString())
				Exit Function
			End if
		Case "expand"  'Fn_BB_TreeGetItemPath(sFunctionName,sAction,ObjTree,StrNode,sType)
			Fn_BB_DataSetTab_Opeartions = Fn_BB_TreeGetItemPath("Fn_BB_DataSetTab_Opeartions","Datasets",ObjTree,sNode,sColType)
			err.clear
			JavaWindow("BriefcaseBrowser").JavaTree("BBTree_VisRel_CTabFolder").Expand Fn_BB_DataSetTab_Opeartions
			If Err.Number < 0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to expand node  " + dicOpenAssm("Node") +" in javatree " + objBB.JavaTree("BBTree").ToString())
				Exit Function
			End if
			wait 1
		Case "collapse"		
			Fn_BB_DataSetTab_Opeartions = Fn_BB_TreeGetItemPath("Fn_BB_DataSetTab_Opeartions","Datasets",ObjTree,sNode,sColType)
			err.clear
			JavaWindow("BriefcaseBrowser").JavaTree("BBTree_VisRel_CTabFolder").Collapse Fn_BB_DataSetTab_Opeartions
			If Err.Number < 0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to expand node  " + dicOpenAssm("Node") +" in javatree " + objBB.JavaTree("BBTree_VisRel_CTabFolder").ToString())
				Exit Function
			End if
			wait 1
		Case "verify_columnval"
			iIndex = Fn_BB_TreeGetItemPath("Fn_BB_DataSetTab_Opeartions","Datasets",ObjTree,sNode,sColType)
			Fn_BB_DataSetTab_Opeartions = objTree.GetColumnValue(iIndex,sColName)
			If Fn_BB_DataSetTab_Opeartions <> sColValue Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to retrive column ["+sColName+"] value from " + sNode +" in javatree " + objBB.JavaTree("BBTree_2").ToString())
				Exit Function
			End If	
			wait 1			
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_BB_DataSetTab_Opeartions >> Invalid case")
			exit function
	End select	
	Fn_BB_DataSetTab_Opeartions = true
	set objTab = Nothing
	set objTree = Nothing
End Function

'*********************************************************		Function to Synchronize on Application Response	***********************************************************************
'Function Name		:				Fn_BB_TreeGetItemPath(sFunctionName,sAction,ObjTree,StrNode,sType)

'Description			 :		 		 to get path of javatree node

'Parameters			   :	 			1. sFunctionName : name of function from where it is calling
'										2. sAction: this is will be set to "Datasets" while fetching path of dataset tree otherwise keep it bank
'										3. ObjTree: tree object on which retriving the path of node
'										4. StrNode : full path of node
'										5. sType : this will be set in case of dataset tree, to define "type" of node we are retriving - where TYPE is Column name under tree
'
'Return Value		   : 				Path of tree \ FALSE

'Pre-requisite			:		 		Node should be selected on which dataset is created

'Examples				:				 Fn_BB_TreeGetItemPath("Fn_BB_DataSetTab_Opeartions","Datasets",ObjTree,"DS1Word/A:Child1","MSWord")

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Shweta Rathod			25-Jul-2016			1.0
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_BB_TreeGetItemPath(sFunctionName,sAction,ObjTree,StrNode,sType)
	Dim sItemPath,aStrNode,bFlag,i,iNodeItemsCount
	Dim oCurrentNode,eStrNode, iNodecnt
	
	bFlag = false
	aStrNode = split(StrNode,":")
	iRow = objTree.GetROProperty ("items count")
	If sAction = "Datasets" Then
		For iCnt = 0 to iRow - 1
			eStrNode = aStrNode(iNodecnt)
			sAppTypeCol = objTree.GetColumnValue(iCnt,"Type")
			sAppNameCol = objTree.GetColumnValue(iCnt,"Name")
			If sAppTypeCol = sType and sAppNameCol = Trim(eStrNode) then
				iRootIndex = iCnt
				bFlag=True
				Exit for
			ElseIf iCnt = iRow - 1 Then
				Fn_BB_TreeGetItemPath = False
				Exit Function
			End If
		next
	else
		bFlag=True
		iRootIndex = 0
	End If
	
	Set oCurrentNode = ObjTree.Object.getItem(iRootIndex)
	sItemPath = "#" & iRootIndex
	
	If UBound(aStrNode) = 0 Then
		Fn_BB_TreeGetItemPath = sItemPath
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
	If bFlag=True Then
		'Function Returns Item Path
		Fn_BB_TreeGetItemPath = sItemPath
		Set objNodeBounds = oCurrentNode.getBounds()
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Function " & sFunctionName & " executed successfully for item [ " & StrNode & " ]"  )
	Else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Function " & sFunctionName & " Failed to find item [ " & StrNode & " ]"  )
		Fn_BB_TreeGetItemPath = False
	End If
	Set oCurrentNode =Nothing
End Function

'*********************************************************		Function to Synchronize on Application Response	***********************************************************************
'Function Name		:				Fn_BB_BriefcaseBrowserExit(sAction,sPath,sReservered)

'Description			 :		 		exit form briefcase browser application

'Parameters			   :	 			1. sAction: There are two ways to exit from application accordingly need to set this value 
'										2. sPath : path to save the briefcase browser
'										3. sReservered : reserved for future use
'
'Return Value		   : 				Path of tree \ FALSE

'Pre-requisite			:		 		Briefcase browser should be open and all other dialoues inside application should close if open

'Examples				:				 Fn_BB_BriefcaseBrowserExit("withsave","C:\Temp\BBAssembly\TestCaseName_IrandNum","")
'										Fn_BB_BriefcaseBrowserExit("withoutsave","","")

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Shweta Rathod			25-Jul-2016			1.0
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function Fn_BB_BriefcaseBrowserExit(sAction,sPath,sReservered)
	Dim objJavaWindowExit
	Fn_BB_BriefcaseBrowserExit = false
	Set objJavaWindowExit =JavaWindow("BriefcaseBrowser").JavaWindow("SaveResource")
	if objJavaWindowExit.Exist(1) = false then
		Fn_BB_BriefcaseBrowserExit = Fn_UI_JavaMenu_Select("Fn_BB_BriefcaseBrowserExit",JavaWindow("BriefcaseBrowser"), Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("BriefcaseBrowser_Menu"), "BBExit"))
		wait 1
		If  objJavaWindowExit.Exist(5) = false Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Function >> Fn_BB_BriefcaseBrowserExit object does not exist [ " & StrNode & " ]"  )
		End if
	End if 
	Select Case sAction
	 	Case "withoutsave"
	 		objJavaWindowExit.JavaButton("Yes").SetTOProperty "label","No"
	 		err.clear
	 		objJavaWindowExit.JavaButton("Yes").Click
	 		If Err.Number < 0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to click on [ No] button of dialog " + objJavaWindowExit.ToString())
			End if			
	 	Case "withsave"
	 End Select
	 
	 Call Fn_KillProcess("BriefcaseBrowser.exe")
	 Set objJavaWindowExit = nothing
	 Fn_BB_BriefcaseBrowserExit = true
End Function


'*********************************************************		Function to Synchronize on Application Response	***********************************************************************
'Function Name		:				Fn_BB_BriefcaseBrowserSave(sFilename,sReservered)

'Description			 :		 		Save the briefcase browser file 

'Parameters			   :	 			1. sFilename : fileName (should not be the full path)
'										3. sReservered : reserved for future use
'
'Return Value		   : 				Path of tree \ FALSE

'Pre-requisite			:		 		Briefcase browser should be open and all other dialoues inside application should close if open

'Examples				:				 Fn_BB_BriefcaseBrowserSave("fileName","")

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Shweta Rathod			25-Jul-2016			1.0
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function Fn_BB_BriefcaseBrowserSave(sFilename,sReservered)
Dim objbbSave,bRet
Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
Dim stempFolderPath: stempFolderPath = fso.GetSpecialFolder(2)

bRet = false
Fn_BB_BriefcaseBrowserSave = "false"	
Set objbbSave =Dialog("Save")
objbbSave.SetTOProperty "text", "Specify the name of the file in which to save briefcase"
objbbSave.WinButton("Open").SetTOProperty "text","&Save"
objbbSave.WinEdit("Name").SetTOProperty "attached text","File name:"

strFilePath = Environment.Value("BBAssemblyPath")
If NOT fso.FolderExists(strFilePath) Then
	FSO.CreateFolder(strFilePath)
End If
If NOT fso.FolderExists(strFilePath) Then
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Function >> Fn_BB_BriefcaseBrowserSave failed to create folder [ " & stempFolderPath+"\"+strFilePath & " ]"  )
	Exit Function
End If
strFilePath = strFilePath+"\"+sFilename+".bcz"

if objbbSave.Exist(1) = false then
	bRet = Fn_UI_JavaMenu_Select("Fn_BB_BriefcaseBrowserSave",JavaWindow("BriefcaseBrowser"), Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("BriefcaseBrowser_Menu"), "FileSaveBriefcase"))
	wait 1
	bRet = Fn_UI_ObjectExist("Fn_BB_BriefcaseBrowserSave",objbbSave) 
	If  bRet = false Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Function >> Fn_BB_BriefcaseBrowserSave object does not exist [ " & objbbSave.tostring() & " ]"  )
		Exit Function
	End if	
End if 
If objbbSave.Exist(5) Then
	Err.clear
	objbbSave.WinEdit("Name").Set strFilePath
	wait 1
	If Err.Number < 0 Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to set the value in edit box " + objbbSave.WinEdit("Name").ToString())
		Exit Function
	End if
	wait 1
	Err.clear
	objbbSave.WinButton("Open").Click micLeftBtn
	If Err.Number < 0 Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to click on button " + objbbSave.WinButton("Open").ToString())
		Exit Function
	End if
	wait 5		
else
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " dialog " + objbbSave.ToString()+" does not exist.")
	Exit Function
End If

If fso.FileExists(strFilePath) = false then
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  SVerified that the exsitence of briefcase browser ["+strFilePath+"] file")
	Set fso = Nothing
	Exit Function
End if
	
set objbbSave = Nothing
If bRet = true then 
	Fn_BB_BriefcaseBrowserSave = strFilePath
End if
End Function

'*************************************************  Function to perform various opeartions on the prefrence operation window ***********************************************************************
'Function Name		:				Fn_BB_PrefrenceOpeartion(sAction,dicPreference)

'Description			 :		 		to perform various opeartions on the prefrence operation window

'Parameters			   :	 			1. sAction : set_configuration - set the configuration profile (ex-unman,default etc..)
'										2. dicPreference : parameters to set the configuration 
'
'Return Value		   : 				in all other case [ true ] and in case of "get_configuration" [ Profile name which is currently set] \ FALSE

'Pre-requisite			:		 		Briefcase browser should be open and all other dialoues inside application should close if open

'Examples				:				1. Set dicPref = CreateObject("Scripting.Dictionary")
'										dicPref("PrefrenceName") = DataTable("PrefrenceName",dtGlobalSheet)
'										dicPref("PrefrenceFullPath") = DataTable("PrefrenceFullPath",dtGlobalSheet)
'										dicPref("ConfigurationName") = DataTable("ConfigurationName",dtGlobalSheet)
'										bReturn = Fn_BB_PrefrenceOpeartion("set_configuration",dicPref)
'										2. bReturn = Fn_BB_PrefrenceOpeartion("get_configuration","")

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Shweta Rathod			25-Jul-2016			1.0
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function Fn_BB_PrefrenceOpeartion(sAction,dicPreference)
Dim sMenu,objPrefrence,sAppText
	Fn_BB_PrefrenceOpeartion = false
	If sAction <> "get_configuration" then
		sMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("BriefcaseBrowser_Menu"), "WindowPrefrences")
		set objPrefrence = Fn_SISW_BB_GetObject("DlgPreferences")
		If objPrefrence.exist(1) = false then
			call Fn_UI_JavaMenu_Select("Fn_BB_CreateDataset",JavaWindow("BriefcaseBrowser"),sMenu)
			wait 1
			if objPrefrence.exist(3) = false then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to invoke Prefrence dialog")
				Exit Function
			End if	
		End if
	End if
	Select Case sAction
		Case "set_configuration"
			call Fn_Edit_Box("Fn_BB_PrefrenceOpeartion",objPrefrence,"FilterText",dicPreference("PrefrenceName"))
			wait 1
			Err.clear
			objPrefrence.JavaTree("Tree").Select dicPreference("PrefrenceFullPath")
			If Err.Number < 0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select node  " + dicPreference("PrefrenceFullPath") +" in javatree " + objPrefrence.JavaTree("Tree").ToString())
				Exit Function
			End if
			wait 1
			err.clear
			objPrefrence.JavaList("ConfigurationName").Select dicPreference("ConfigurationName")
			If Err.Number < 0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to set configuration name  " + dicPreference("ConfigurationName") +" in javalist " + objPrefrence.JavaList("ConfigurationName").ToString())
				Exit Function
			End if
			wait 1
			objPrefrence.JavaButton("OK").SetTOProperty "label","Apply"
			objPrefrence.JavaButton("OK").click
			wait 1
			objPrefrence.JavaButton("OK").SetTOProperty "label","OK"
			objPrefrence.JavaButton("OK").click
		Case "get_configuration"
			Fn_BB_PrefrenceOpeartion = JavaWindow("BriefcaseBrowser").GetROProperty("title")
			Exit Function
		Case "verify_releasestatus"
			If dicPreference("TCMRelease") <> "" then
				objPrefrence.JavaList("ConfigurationName").SetTOProperty "attached text","Release Status Name"
				sAppText = objPrefrence.JavaList("ConfigurationName").GetROProperty("text")
				If trim(lcase(dicPreference("TCMRelease"))) <> trim(lcase(sAppText)) then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to verify default value  " + dicPreference("TCMRelease") + " in javalist " + objPrefrence.JavaList("ConfigurationName").ToString())
					Exit Function
				End if
			else
				objPrefrence.JavaButton("OK").SetTOProperty "label","Cancel"
				objPrefrence.JavaButton("OK").click
				Fn_BB_PrefrenceOpeartion = false
				Exit function
			End if
			objPrefrence.JavaButton("OK").SetTOProperty "label","Cancel"
			objPrefrence.JavaButton("OK").click			
	End Select	
	Fn_BB_PrefrenceOpeartion = true
	Set objPrefrence = nothing
End Function


'*************************************************  Function to perform various opeartions on the prefrence operation window ***********************************************************************
'Function Name		:				Fn_BB_JavaTab_Operations(sAction, objJavaDialog, sTabObjectName, sTabName,sPopupMenu,sReserved)

'Description			 :		 		to perform various opeartions on the java tab which is displaying in BB application

'Parameters			   :	 			1. sAction : select/exist/close
'										2. objJavaDialog : dilaog on which operation needs to perform / if it is balnk it will set to default window of BB
'										3. sTabObjectName : object of tab / if balnk default will be "CTabFolder"
'										4. sTabName : "Welcome"
'										4. sPopupMenu : popup menu to be performed
'										5. sReserved : future use
'
'Return Value		   : 				true \ false

'Pre-requisite			:		 		Briefcase browser or dilaog on which tab is displaying should be open 

'Examples				:				bReturn	= Fn_BB_JavaTab_Operations("close","","","Welcome","","")

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Shweta Rathod			01-Aug-2016			1.0
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_BB_JavaTab_Operations(sAction, objJavaDialog, sTabObjectName, sTabName,sPopupMenu,sReserved)
	Dim objTab,bFlag
	Fn_BB_JavaTab_Operations = false
	if objJavaDialog = "" then set objJavaDialog = JavaWindow("BriefcaseBrowser")
	if sTabObjectName = "" then sTabObjectName = "CTabFolder"
	Set objTab = objJavaDialog.JavaTab(sTabObjectName)
	objTab.SetTOProperty "value",sTabName
	if objTab.exist(1) = false then
		bFlag = false	
	else
		bFlag = true
	End if	
	Select Case sAction
		Case "select"
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
			objTab.Minimize() 
			If Err.Number < 0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to minimize the tab " + objTab.tostring())
				Exit Function
			End if
			bFlag = objTab.object.getMinimized()
			If bFlag = "false" Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to verify tab is in minimize state " + objTab.tostring())
				Exit Function
			End If
			objTab.Minimize()
		Case "maximize"
			err.clear
			objTab.Maximize 
			If Err.Number < 0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Maximize the tab " + objTab.tostring())
				Exit Function
			End if
			bFlag = objTab.object.getMaximized()
			If bFlag = "false" Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to verify tab is in Maximize state " + objTab.tostring())
				Exit Function
			End If
			objTab.Restore
	End Select 
	Set objTab = Nothing	
	Fn_BB_JavaTab_Operations = true
End Function


'*************************************************  Function to perform various opeartions on the prefrence operation window ***********************************************************************
'Function Name		:					Fn_BB_LoadBBXML()

'Description			 :		 		this will load all the enviornment level variable which are given in configuration file BBEnvVar.xml

'
'Return Value		   : 				nothing

'Pre-requisite			:		 		File "BBEnvVar.xml" should be present with all the valid values

'Examples				:				Fn_BB_LoadBBXML()

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Shweta Rathod			01-Aug-2016			1.0
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_BB_LoadBBXML()
	Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
	Environment.Value("BBCustomDatasetMappings") = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("BriefcaseBrowser_Envvar"), "BBCustomDatasetMappings")
	Environment.Value("BriefcaseBrowserPath") = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("BriefcaseBrowser_Envvar"), "BriefcaseBrowserPath")
	Environment.Value("BBNX_AssemblyPath") = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("BriefcaseBrowser_Envvar"), "BBNX_AssemblyPath")
	Environment.Value("BBAssemblyPath") = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("BriefcaseBrowser_Envvar"), "BBAssemblyPath")
	
	'**********************************************************************************
	'creating folder to save NX assembly
	'**********************************************************************************
	If Not fso.FolderExists(Environment.Value("BBNX_AssemblyPath")) then
		fso.CreateFolder(Environment.Value("BBNX_AssemblyPath"))
	End if
	Set fso = nothing
End Function



'*************************************************  Function to extract the content of ".bcz" file into given new directory ***********************************************************************
'Function Name		:				Fn_BB_ExtractContentBriefcaseBrowser(sFileName,sReserved)

'Description			 :		 		to perform various opeartions on the java tab which is displaying in BB application

'Parameters			   :	 			1. sFileName : required file name in which extracting the content
'										2. sReserved : future use
'
'Return Value		   : 				FullPath of directory / "false"

'Pre-requisite			:		 		.bcz file should be present into the repective directory 

'Examples				:				bReturn = Fn_BB_ExtractContentBriefcaseBrowser("TestCaseName_iRandNo","")

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Shweta Rathod			01-Aug-2016			1.0
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function Fn_BB_ExtractContentBriefcaseBrowser(sFileName,sReserved)
	Dim objShell: set objShell = CreateObject("Shell.Application")
	Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
	Dim stempFolderPath: stempFolderPath = fso.GetSpecialFolder(2)
	Dim objFolder,bRet
	Dim sSource,strRename,sDestination
	On Error Resume Next
	Fn_BB_ExtractContentBriefcaseBrowser = "false"
	sSource = Environment.Value("BBAssemblyPath")+"\"+sFilename+".bcz"
	strRename = Environment.Value("BBAssemblyPath")+"\"+sFilename+".zip"
	If NOT fso.FolderExists(stempFolderPath+"\"+sFilename) Then
		FSO.CreateFolder(stempFolderPath+"\"+sFilename)
	End If
	sDestination = stempFolderPath+"\"+sFilename
	fso.CopyFile sSource,sDestination+"\",True
	wait 5
	If not fso.FileExists(sDestination+"\"+sFilename+".bcz") then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  SVerified that the File ["+sFilename+"] is Copied at the Path ["+sDestination+"]")
		Set fso = Nothing
		Exit Function
	End if
	FSO.MoveFile sDestination+"\"+sFilename+".bcz", strRename
	
	If NOT fso.FolderExists(stempFolderPath+"\"+sFilename) Then
		FSO.CreateFolder(stempFolderPath+"\"+sFilename)
	End If
	'Extract the contants of the zip file.
	set FilesInZip=objShell.NameSpace(strRename).items
	objShell.NameSpace(stempFolderPath+"\"+sFilename).CopyHere(FilesInZip)
	wait 10
	Set objFolder = fso.GetFolder(stempFolderPath+"\"+sFilename)
	If not(objFolder.Files.Count = 0 or objFolder.SubFolders.Count = 0) then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Verified that the File ["+sFilename+"] is Copied at the Path ["+sDestination+"]")
		Set fso = Nothing
		Exit Function
	End if
	
	Set fso = Nothing
	Fn_BB_ExtractContentBriefcaseBrowser = stempFolderPath+"\"+sFilename
End Function


'*************************************************  Function to verify briefcase browser content ***********************************************************************
'Function Name		:				Fn_BB_VerifyBriefcaseBrowserContent(sFilePath,sFileName,sReserved)

'Description			 :		 		verifying content of extracting directory

'Parameters			   :	 			1. sFilePath : required file name in which extracting the content
'										2. sFileName : 
'										3. sReserved : future use
'
'Return Value		   : 				FullPath of directory / "false"

'Pre-requisite			:		 		.bcz file should be extracted into the new directory 

'Examples				:				bReturn = Fn_BB_VerifyBriefcaseBrowserContent("C:\Temp\TestCaseName_iRandNo","child1.prt","")

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Shweta Rathod			01-Aug-2016			1.0
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function Fn_BB_VerifyBriefcaseBrowserContent(sFilePath,sFileName,sReserved)
	Dim aFileName
	Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
	Fn_BB_VerifyBriefcaseBrowserContent = false
	if sFileName <> "" then aFileName = split(sFileName,"~")
	If NOT fso.FolderExists(sFilePath) Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed - folder dose not exist at location ["+sFilePath+"]")
		Set fso = Nothing
		Exit Function
	End if
	For iCnt = 0 to ubound(aFileName)
		If not fso.FileExists(sFilePath+"\"+aFileName(iCnt)) then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Verified existence of File [ "+aFileName(iCnt)+" ] in the BB folder at location ["+sFilePath+"]")
			Set fso = Nothing
			Exit Function
		End if
		wait 1
	next
	Set fso = nothing
	Fn_BB_VerifyBriefcaseBrowserContent = true
End Function


'Function Fn_BB_BriefcaseBrowserExit()
' Dim objJavaWindowDefault, objJavaWindowExit, oDesc, intNoOfObjects
' Dim sWinTitle
' Set objJavaWindowDefault = Fn_UI_ObjectCreate("Fn_BB_BriefcaseBrowserExit",JavaWindow("BriefcaseBrowser"))
'
'   'Select menu [File -> Exit]
'  bReturn = Fn_BB_BriefcaseBrowserExit("withoutsave","","")
'   
'    'If Teamceneter window exists after log out
'	'If JavaWindow("DefaultWindow").Exist(10) Then
'	If Fn_UI_ObjectExist("Fn_TeamcenterExit",JavaWindow("DefaultWindow")) = True Then
'		sWinTitle = JavaWindow("DefaultWindow").GetROProperty("title")
'		If instr(sWinTitle, "Business Modeler IDE") > 0 Then
'			Fn_TeamcenterExit = True
'			Set objJavaWindowDefault =Nothing
'			Set objJavaWindowExit =Nothing
'			Exit Function
'		Else
'			Fn_TeamcenterExit = FALSE  
'		End If
'	Else
'		Fn_TeamcenterExit = TRUE  
'	End If
'End Function

'*********************************************  Function Componentise Actiob1 in the Script ..**************************************************************

'Function Name		:					Fn_BriefcaseBrowser_TestcaseExit

'Description			 :		 		  The function handles test script end part

'Parameters			   :	 			

'Return Value		   : 				True/False

'Pre-requisite			:		 		None

'Examples				:
'

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										shweta Rathod			05-Oct-2016	   		1.0
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function Fn_BriefcaseBrowser_TestcaseExit(bTcKill)
	Dim bReturn,filePath,objFSO,bCaptureImgVP

	On Error Resume Next
	call Fn_UI_JavaMenu_Select("Fn_BB_BriefcaseBrowserExit",JavaWindow("BriefcaseBrowser"), Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("BriefcaseBrowser_Menu"), "BBExit"))
	Call Fn_UpdateLogFiles("[" + Cstr(Time) + "] - ACTION - PASS | Successfully performed menu opeartion [ "+Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("BriefcaseBrowser_Menu"), "BBExit")+" ].", "")
	Set objJavaWindowExit =JavaWindow("BriefcaseBrowser").JavaWindow("SaveResource")
	If objJavaWindowExit.Exist(1) = true Then
		objJavaWindowExit.JavaButton("Yes").SetTOProperty "label","No"
 		objJavaWindowExit.JavaButton("Yes").Click
 		Call Fn_UpdateLogFiles("[" + Cstr(Time) + "] - ACTION - PASS | Successfully clicked on [ No ] button of [ Save Resource ] window.", "")
		Call Fn_UpdateLogFiles("[" + Cstr(Time) + "] - ACTION - PASS | Successfully exit from [ Briefcase Browser ] application.", "")
	End If
	 
	Call Fn_KillProcess("BriefcaseBrowser.exe")
	'**********************************************************************************
	'Log Test Result
	''**********************************************************************************
	Call Fn_UpdateLogFiles("[" + Cstr(time) + "] - QTP [" + Environment.Value("ActionName") + "] - End", "")
	Call Fn_UpdateLogFiles("-----------------------------------------------------------------------------------------------", "")
	If bCaptureImgVP = True Then
		Call Fn_UpdateLogFiles("[" + Cstr(time) + "] - Final - Pass | Test Execution Result: PASS without comparing images", "PASS: All VP Pass without comparing images")
	Else
		Call Fn_UpdateLogFiles("[" + Cstr(time) + "] - Final - Pass | Test Execution Result: PASS", "PASS: All VP Pass")
	End If
	Call Fn_UpdateLogFiles("-----------------------------------------------------------------------------------------------", "")
	'**********************************************************************************
	'Deleting Snapshot Image file, snapshot is not needed if ALL VP PASS
	'**********************************************************************************
	filePath = Environment.Value("BatchFldName") + "\" + Environment.Value("TestName") + ".png"
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	if objFSO.FileExists(filePath) then
		objFSO.DeleteFile filePath, True
	End if
	Set objFSO = Nothing
End Function



'*************************************************  Function to perform various opeartions on toolbar ***********************************************************************
'Function Name		:				Fn_BB_ToolBar_Opeartions(sAction,objBB,sButtonName,sReserved)

'Description			 :		 		to perform various opeartions on toolbar

'Parameters			   :	 			1. sAction : isenabled - check toolbar button is enabled
'										2. objBB : if balnk it will refer defaultwindow of BB / object on which needs to perfrom toolbar opeartion
'										3. sButtonName : toolbar button name
'										4. sReserved : future use
'
'Return Value		   : 				true / false

'Pre-requisite			:		 		briefcase browser should be open

'Examples				:				Fn_BB_ToolBar_Opeartions("isenabled","","Open Briefcase","")

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Shweta Rathod			01-Aug-2016			1.0
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function Fn_BB_ToolBar_Opeartions(sAction,objBB,sButtonName,sReserved)
	Dim objToolBtnList,iToolCnt,iCounter,sContents
	If  objBB = "CTabFolder" Then
		Set objBB = JavaWindow("BriefcaseBrowser").JavaTab("CTabFolder")
	ElseIf objBB = "OwnershipTransfer" Then
		Set objBB = JavaWindow("BriefcaseBrowser").JavaTab("OwnershipTransfer")
	Else
		Set objBB = JavaWindow("BriefcaseBrowser")
	End If
	
	
	
	If objBB.Exist(1) Then
		'Create Toolbar object
		Set ObjDesc = Description.Create() 
		ObjDesc("to_class").Value = "JavaToolbar" 
		ObjDesc("enabled").Value = 1
		
		'Get the total of Toolbar objects
		Set objToolBtnList =objBB.ChildObjects(ObjDesc)
		iToolCnt = objBB.ChildObjects(ObjDesc).count
	Else
		Fn_BB_ToolBar_Opeartions = FALSE
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: JavaWindow [DefaultWindow] Does not exist.")
	End If
	
	Select Case sAction
		Case "isenabled"			
			For iCounter = 0 to iToolCnt-1
				sContents = objToolBtnList(iCounter).GetContent()
				If instr(sContents, sButtonName) > 0 Then
					If  "1" = objToolBtnList(iCounter).GetItemProperty (sButtonName, "enabled")  Then
			            Fn_BB_ToolBar_Opeartions = TRUE
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Object Java Button["+sButtonName+"] Is Enabled.")
					Else
						Fn_BB_ToolBar_Opeartions = FALSE
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Object Java Button["+sButtonName+"] Is Not Enabled.")
					End if
					Exit For
				End If
			Next			
		Case "isselected"
			For iCounter = 0 to iToolCnt-1
				sContents = objToolBtnList(iCounter).GetContent()
				If instr(sContents, sButtonName) > 0 Then
					If  "1" = objToolBtnList(iCounter).GetItemProperty (sButtonName, "selected")  Then
                        Fn_BB_ToolBar_Opeartions = TRUE
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Object Java Button["+sButtonName+"] Is Selected.")
					Else
						Fn_BB_ToolBar_Opeartions = FALSE
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Object Java Button["+sButtonName+"] Is Not Selected.")
					End if
					Exit For
				End If
			Next	
		Case "click"
			For iCounter = 0 to iToolCnt-1
				sContents = objToolBtnList(iCounter).GetContent()
				If instr(lcase(sContents), lcase(sButtonName)) > 0 Then
					iItmCount = objToolBtnList(iCounter).GetROProperty("toolbar items")
	                aContents = split(sContents, ";", -1, 1)
					For iCounter1 = 1 to Ubound(aContents)+1
						sItmText = objToolBtnList(iCounter).GetItemProperty(iCounter1, "name")
						If Trim(lcase(sItmText) )=trim( lcase(sButtonName)) Then
							objToolBtnList(iCounter).Press iCounter1
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Clicked on Toolbar Button " + sButtonName)
							Fn_BB_ToolBar_Opeartions = TRUE
							Exit Function
						else
							Fn_BB_ToolBar_Opeartions = FALSE
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: to Clicked on Toolbar Button " + sButtonName)
						End If
					Next
				End If
			Next
	End Select
	
If iCounter = iToolCnt Then					
	Fn_BB_ToolBar_Opeartions = FALSE
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Object Java Button["+sButtonName+"] Not Found.")
End If
set objToolBtnList = Nothing
Set objBB = Nothing
End Function


'*************************************************  Function to verify property of BB object from Item property tab ***********************************************************************
'Function Name		:				Fn_BB_ItemPropertyPanelVerify("Item Name","Item123")

'Description			 :		 		to perform various opeartions on toolbar

'Parameters			   :	 			1. sPropName : name of the property
'										2. sValue : value of property to be verify
'
'Return Value		   : 				true / false

'Pre-requisite			:		 		briefcase browser should be open

'Examples				:				Fn_BB_ItemPropertyPanelVerify("Properties:PropertyName1;Properties:PropertyName2","PropertyValue1;PropertyValue2")

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Shweta Rathod			01-Aug-2016			1.0
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function Fn_BB_ItemPropertyPanelVerify(sPropName,sValue)
	Dim objBB,objTab,objPropertyTree,aPropNameArr,aPropValueArr
	Dim iCount,sAppPropValue,bRet,bCounter
	
	set objBB = Fn_SISW_BB_GetObject("BriefcaseBrowser")
	set objTab = objBB.JavaTab("CTabFolder")
	objTab.SetTOProperty "value","Item Properties"
	
	Set objPropertyTree = objBB.JavaTree("BBTree_VisRel_CTabFolder")
	If objPropertyTree.exist(1) = false then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Object does not exist "+objPropertyTree.tostring)
		Fn_BB_ItemPropertyPanelVerify = false
		Exit function
	End if
	
	bRet = Fn_BB_ToolBar_Opeartions("isselected","","Show Advanced Properties","")
	If bRet = false then
		bRet = Fn_BB_ToolBar_Opeartions("click","","Show Advanced Properties","")
		wait 1
	End if
	
	bRet = Fn_BB_ToolBar_Opeartions("isselected","","Show Categories","")
	If bRet = true then
		bRet = Fn_BB_ToolBar_Opeartions("click","","Show Categories","")
		wait 1
	End if
	
	bCounter = 0
	aPropNameArr = split(sPropName,"~")
	aPropValueArr = split(sValue,"~")
	For jCnt = 0 to ubound(aPropNameArr)	
		For iCount=0 to objPropertyTree.GetROProperty("count_all_items")-1
			If aPropNameArr(jCnt) = objPropertyTree.GetItem(iCount) Then
				sAppPropValue=Cstr(objPropertyTree.GetColumnValue("#"+cstr(iCount),"Value"))
				If aPropValueArr(jCnt) = sAppPropValue then
					bret = true
					bCounter = bCounter + 1
					Exit for
				End if
			End If
		next
	next
	
	If cint(ubound(aPropNameArr))+1 <> bCounter then
		bret = false
	else
		bret = true
	End if
	
	set objBB = nothing
	set objTab = nothing
	Set objPropertyTree = nothing
	Fn_BB_ItemPropertyPanelVerify = bret
End Function

'*******************************************************************************************************************************
'Function Name		:				Fn_BB_ErrorLog_Opeartions(sAction,sTabName,sNode,sReserve)

'Description			 :		 		 This function perform various operations on Briefcase browser error tree

'Parameters			   :	 			1. sAction : type of assembly to be open CAD or BB 
'										2. sTabName: name of the tab under BB tree is displaying on which going to perform the operation
'										3. sNode: full path of the node on which performing the opeation
'										4. sReserve : for future use
											
'Return Value		   : 				TRUE \ FALSE

'Pre-requisite			:		 		Brief case browser application should be displayed

'Examples				:				 Fn_BB_ErrorLog_Opeartions("isempty","*C:\mainline\Reports\NX\Add_Attachs_Default_Relat_88663\Top88663_A.prt","Top88663:child188663","")

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Amruta Patil			07-Jan-2021			1.0
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_BB_ErrorLog_Opeartions(sAction,sTabName,sNode,sReserve)
	Dim objTab,iRow
	Fn_BB_ErrorLog_Opeartions = false
	set objTab = Fn_SISW_BB_GetObject("BriefcaseBrowser")
	Set objTree = objTab.JavaTree("BBTree_VisRel_CTabFolder")
	set objTab = objTab.JavaTab("CTabFolder")
	Fn_BB_ErrorLog_Opeartions = Fn_UI_Object_SetTOProperty_ExistCheck("Fn_BB_ErrorLog_Opeartions",objTab,"value",sTabName)
	If Fn_BB_ErrorLog_Opeartions = true then
		Fn_BB_ErrorLog_Opeartions = Fn_UI_JavaTab_Select("Fn_BB_ErrorLog_Opeartions",JavaWindow("BriefcaseBrowser"),"CTabFolder", sTabName)
		wait 1
		If Fn_BB_ErrorLog_Opeartions = false Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select tab  " + objTab.ToString())
			Exit Function
		End if		
	else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "tab does not exist " + objTab.ToString())
		Exit Function
	End if
	Select Case lcase(sAction)
		Case "isempty"
			iRow = objTree.GetROProperty ("items count")
			If iRow <> 0 Then
				If Fn_BB_ErrorLog_Opeartions = false Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to verify tree is empty" + JavaWindow("BriefcaseBrowser").JavaTree("BBTree").ToString())
					Exit Function
				End if
			End If
			
			Case "close"
			if Fn_BB_ErrorLog_Opeartions = true then 
				err.clear
				objTab.CloseTab sTabName
				If Err.Number < 0 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to close the tab " + objTab.tostring())
					Exit Function
				End if
			End if
		Case "exist"
				If Fn_BB_ErrorLog_Opeartions = false Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "tab does not exist"+objTab.tostring())
					Exit Function
				End If	
	End select
	Fn_BB_ErrorLog_Opeartions = true
	set objTab = Nothing
	set objTree = Nothing
End Function
Public Function Fn_BB_TabOperation(sAction,sTabName)
	
End Function
