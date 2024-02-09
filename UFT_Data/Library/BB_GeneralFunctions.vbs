Option Explicit
'1. Fn_SISW_BB_ResetBriefcaseBrowser()
'2. Fn_SISW_BB_UpdateCutomDatasetMappingXMLNode(XMLDataFile, sTagName,sAttributeName,sAttributeVal,sSubAttributeName,sNewAttributeVal)
'3. Fn_SISW_BB_OpenAssembly(sAction,sPath)
'4. Fn_SISW_BB_CreateDataset(sAction,sPath,sRelationName,sButton)
'5. Fn_SISW_BB_DataSetTree_Opeartions(sAction,sNode,sColType,sColName,sColValue,sRMBMenu)
'6. Fn_SISW_BB_Save(sFilename)
'7. Fn_SISW_BB_PrefrenceOpeartion(sAction,dicPreference)
'8. Fn_SISW_BB_ExtractContentBriefcaseBrowser(sFileName)
'9. Fn_SISW_BB_VerifyBriefcaseBrowserContent(sFilePath,sFileName)
'10. Fn_SISW_BB_ItemPropertyPanelVerify(sPropName,sValue)
'11. Fn_SISW_BB_BriefCaseTree_Opearation(sAction,sTabName,sNode, sColumn, sValue, sRMBMenu) 
'12. Fn_SISW_BB_ErrorLogTree_Opeartions(sAction,discError)
'13. Fn_SISW_BB_OwnershipTransferTree_Opeartions(sAction,sPartNumber,sColName,sColVal,sMenu)
'14. Fn_SISW_BB_OwnershipTransferNATTable_Operations(sAction,sColNames,sColValues,dicDetails,sButton,sReserve)		[vivek.ahirrao.ext@Siemens.com]
'15. Fn_SISW_BB_ErrorDialogVerify(sAction,objErr,dicErrorInfo)

'****************************************    Function to reset BB application ***************************************
'Function Name		 	:	Fn_SISW_BB_ResetBriefcaseBrowser
'
'Description		    :  	Function to Fn_Reset Briefcase Browser application
'
'Return Value		    :  	TRUE \ FALSE
'
'Examples		     	:	Fn_SISW_BB_ResetBriefcaseBrowser()
'
'History:
'	Developer Name			Date			Rev. No.		Reviewer		Changes Done	
'----------------------------------------------------------------------------------------------------------------------------------
'	Shweta Rathod		 19-Jul-2016		1.0				Koustubh Watwe
'----------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_BB_ResetBriefcaseBrowser()	
	Dim sMenu,bRet,objReset
	
	Fn_SISW_BB_ResetBriefcaseBrowser = false 
	set objReset = Fn_SISW_BB_GetObject("BBResetPerspective")
	bRet =  Fn_SISW_UI_Object_Operations("Fn_SISW_BB_ResetBriefcaseBrowser","Exist",objReset,SISW_MICRO_TIMEOUT)
	If bRet = false then
		sMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("BriefcaseBrowser_Menu"), "WindowResetPerspective")
		bRet = Fn_UI_JavaMenu_Select("Fn_MenuOperation",JavaWindow("BriefcaseBrowser"), sMenu)
		wait 1
		If bRet = false then
			 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Function Name >> Fn_SISW_BB_ResetBriefcaseBrowser : Failed to perform menu opeartion [" + sMenu + "]")
			 Exit Function
		End IF
		bRet = Fn_SISW_UI_Object_Operations("Fn_SISW_BB_ResetBriefcaseBrowser","Exist",objReset,SISW_DEFAULT_TIMEOUT) 						  							
		If bRet = false then exit function
	End if
	
	wait 1
	bRet = Fn_SISW_UI_JavaButton_Operations("Fn_SISW_BB_ResetBriefcaseBrowser", "Click", objReset,"Yes") 
	If bRet = false Then Exit function
	
	Fn_SISW_BB_ResetBriefcaseBrowser = true
	Set objReset = Nothing
End Function

'****************************************    Function to Update CutomDatasetMappingXML Node ***************************************
'Function Name		 	:	Fn_SISW_BB_UpdateCutomDatasetMappingXMLNode
'
'Description		    :  	Function to update attributes of CutomDatasetMappingXML Node.
'
'Parameters		    	:	1. XMLDataFile : File name
'							2. sTagName : Tag name of the xml file
'							3. sAttributeName : attribute name to select data 
'							4. sAttributeVal : search Value to be modfied
'							5. sSubAttributeName : Sub attribute name 
'							6. sNewAttributeVal : Value to be modified 
'
'Return Value		    :  	TRUE 
'
'Examples		     	:	Fn_SISW_BB_UpdateCutomDatasetMappingXMLNode("C:\Tc11.2.3_2016060800_BB_Configured\bbworkspace\configurations\Unman\CustomDatasetMappings.xml","data_set_mapping","extension","docx","relation","XMLNewAttributeVal","")
'
'History:
'	Developer Name			Date			Rev. No.		Reviewer			Changes Done	
'----------------------------------------------------------------------------------------------------------------------------------
'	Shweta Rathod		 19-Jul-2016		1.0				Koustubh Watwe
'----------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_BB_UpdateCutomDatasetMappingXMLNode(XMLDataFile, sTagName,sAttributeName,sAttributeVal,sSubAttributeName,sNewAttributeVal)
	Dim objXMLDoc,objXMLNodeList,numObjXMLNodeList,i,sXMLAttribute
	Fn_SISW_BB_UpdateCutomDatasetMappingXMLNode = false
	set objXMLDoc = CreateObject("Microsoft.XMLDOM")
	objXMLDoc.load(XMLDataFile)
	Set objXMLNodeList = objXMLDoc.getElementsByTagName(sTagName)
	numObjXMLNodeList = objXMLNodeList.length
	If numObjXMLNodeList = 0 then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFunctionName&" Failed : Failed to find tag name "&sTagName)
		Exit function
	End if
	For i = 0 to numObjXMLNodeList - 1
		sXMLAttributeVal = objXMLNodeList.item(i).getAttribute(sAttributeName)
		If isNull(sXMLAttributeVal) then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFunctionName&" Failed : Failed find sub attribute value "& sXMLAttributeVal &" for attribute name " & sAttributeName)
			Exit function
		End if
		If sXMLAttributeVal = sAttributeVal then exit for	
	next
	Set objChildNodes = objXMLNodeList.item(i).childNodes
	numObjXMLNodeList = objChildNodes.length
	If numObjXMLNodeList = 0 then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFunctionName&" Failed : Failed to find elements uber attribute "&sAttributeName)
		Exit function
	End if
	For i = 0 to numObjXMLNodeList - 1
		If objChildNodes.item(i).getAttribute(sSubAttributeName) <> sNewAttributeVal Then
			objChildNodes.item(i).setAttribute sSubAttributeName,sNewAttributeVal
		End If		
	next	
	objXMLDoc.Save(XMLDataFile)
	Set objXMLDoc = nothing 
	Set objXMLNodeList = nothing
	Fn_SISW_BB_UpdateCutomDatasetMappingXMLNode = True	
End Function


'*********************************************************		Function to open assembly into the BB application***********************************************************************
'Function Name		:				Fn_SISW_BB_OpenAssembly(sAction,sPath)

'Description		:		 		 This function open the assembly from the given path

'Parameters			:	 			1. sAction : type of assembly to be open CAD or BB 
'										2. sPath : path form assembly to be open
											
'Return Value		: 				TRUE \ FALSE

'Pre-requisite		:		 		Brief case browser application should be displayed

'Examples			:				 Fn_SISW_BB_OpenAssembly("OpenCAD","C:\mainline\Reports\NX\Add_Attachs_Default_Relat_88663\Top88663_A.prt")

'History			:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Shweta Rathod			25-Jul-2016			1.0											Koustubh Watwe
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_BB_OpenAssembly(sAction,sPath)
	Dim bRet,objOpen,sMenu,ObjBB
	Dim sFileName
	Set objOpen =Dialog("OpenCADBB")
	Fn_SISW_BB_OpenAssembly = FALSE
	
	Select Case sAction	
		Case "OpenCAD"
			objOpen.SetTOProperty "text","Select the CAD file."
			bRet = Fn_SISW_UI_Object_Operations("Fn_SISW_BB_OpenAssembly","Exist",objOpen,SISW_MICRO_TIMEOUT) 
			if bRet = false then
				sMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("BriefcaseBrowser_Menu"), "OpenCADAssembly")
				bRet = Fn_UI_JavaMenu_Select("Fn_SISW_BB_OpenAssembly",JavaWindow("BriefcaseBrowser"), sMenu)
				If bRet = false then
					  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Function Name >> Fn_SISW_BB_OpenAssembly >> Failed to Invoke perform menu opeartion [" + sMenu + "]")
					 Exit Function
				End IF
			End if	
			
		Case "OpenBB","OpenBB_FrmFolder"	
			If sAction =  "OpenBB_FrmFolder" Then
				bRet = fn_SISW_util_folder_operation("exist",Environment.Value("BBAssemblyPath"))
				If bRet = false Then
					bRet = fn_SISW_util_folder_operation("createfolder",Environment.Value("BBAssemblyPath")) 
					if bRet = false then
						 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Function Name >> Fn_SISW_BB_OpenAssembly >> Failed to create folder [" + Environment.Value("BBAssemblyPath") + "]")
					 	Exit Function
					End if					
				End If	
				sDestination = Environment.Value("BBAssemblyPath")+"\"
				bRet = Fn_Local_File_Operations("CopyFile" ,sPath, sDestination)
				if bRet = false then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Function Name >> Fn_SISW_BB_OpenAssembly >> Failed to create folder [" + Environment.Value("BBAssemblyPath") + "]")
				 	Exit Function
				 else
				 	sFileName=mid(sPath,instrrev(sPath,"\")+1,len(sPath)-1)
				 	sPath = sDestination&sFileName
				End if
			End If		
			objOpen.SetTOProperty "text", "Select the file representing the briefcase"
			bRet = Fn_SISW_UI_Object_Operations("Fn_SISW_BB_OpenAssembly","Exist",objOpen,SISW_MICRO_TIMEOUT)
			if bRet = false then
				sMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("BriefcaseBrowser_Menu"), "OpenBriefcase")
				bRet = Fn_UI_JavaMenu_Select("Fn_SISW_BB_OpenAssembly",JavaWindow("BriefcaseBrowser"), sMenu)
				If bRet = false then
					 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Function Name >> Fn_SISW_BB_OpenAssembly >>Failed to Invoke perform menu opeartion [" + sMenu + "]")
					 Exit Function
				End IF
			End if
	End Select
	wait 6
	bRet = Fn_SISW_UI_Object_Operations("Fn_SISW_BB_OpenAssembly","Exist",objOpen,SISW_DEFAULT_TIMEOUT)
	If bRet = true Then
		objOpen.WinEdit("FileName").Set sPath
		wait 1
		objOpen.WinButton("Open").Click micLeftBtn
		wait 3
'		bRet = Fn_SISW_UI_WinEdit_Operations("Fn_SISW_BB_OpenAssembly","Set",objOpen,"FileName",sPath)
'		If bRet = false Then Exit Function
'		wait 2
'		bRet = Fn_SISW_UI_WinButton_Operations("Fn_SISW_BB_OpenAssembly","click",objOpen,"Open","","","")
'		If bRet = false Then Exit Function
		If sAction = "OpenBB" or sAction = "OpenBB_FrmFolder" Then
			Set ObjBB = JavaWindow("BriefcaseBrowser").JavaWindow("OpenBriefcase")
			bRet = Fn_SISW_UI_Object_Operations("Fn_SISW_BB_OpenAssembly","Exist",ObjBB,SISW_DEFAULT_TIMEOUT)
			If bRet = true Then	
				bRet = Fn_SISW_UI_JavaButton_Operations("Fn_SISW_BB_OpenAssembly", "Click", ObjBB,"OK")
				if bRet = false then exit function
			End If
		End if
		wait 2		
	else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " dialog " + objOpen.ToString()+" does not exist.")
		Exit Function
	End If	
	Set objOpen = Nothing
	Set ObjBB = nothing
	Fn_SISW_BB_OpenAssembly = True
End Function


'*********************************************************		Function to create dataset	***********************************************************************
'Function Name		:				Fn_SISW_BB_CreateDataset(sAction,sPath,sRelationName,sButton)

'Description			 :		 		 This function create the dataset on the selected item

'Parameters			   :	 			1. sAction : type of assembly to be open CAD or BB 
'										2. sTabName: name of the tab under BB tree is displaying on which going to perform the operation
'										3. sNode: full path of the node on which performing the opeation
											
'Return Value		   : 				TRUE \ FALSE

'Pre-requisite			:		 		Node on which creating dataset should be selected

'Examples				:				 Fn_SISW_BB_CreateDataset("verify_defaultrelname","C:\mainline\Scripts\REG-BriefcaseBrowser\Add_Attachs_Default_Relation_Type\DSWord.docx","rendering","","")
' 										Fn_SISW_BB_CreateDataset("create","C:\mainline\Scripts\REG-BriefcaseBrowser\Add_Attachs_Default_Relation_Type\DSWord.docx","rendering","OK","")

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Shweta Rathod			25-Jul-2016			1.0										Koustubh Watwe
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_BB_CreateDataset(sAction,sPath,sRelationName,sButton)
	Dim sRelName,sMenu
	Dim objBB,objOpen,objAddDt
	set objBB = Fn_SISW_BB_GetObject("BriefcaseBrowser")
	Set objOpen = Fn_SISW_BB_GetObject("OpenDatasetFile")
	set objAddDataset = Fn_SISW_BB_GetObject("AddDataset")
	
	Fn_SISW_BB_CreateDataset = false
	
	if sPath <> "" then
		objOpen.SetTOProperty "text","Select a file to attach to the selected CAD part or assembly" 
		bGblFuncRetVal = Fn_SISW_UI_Object_Operations("Fn_SISW_BB_CreateDataset","Exist",objOpen,SISW_MICRO_TIMEOUT)
		If  bGblFuncRetVal = false Then
			bGblFuncRetVal = Fn_UI_JavaMenu_Select("Fn_SISW_BB_CreateDataset",objBB, Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("BriefcaseBrowser_Menu"), "AddDataset"))
			If bGblFuncRetVal = false Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to invoke menu  " + objTab.ToString())
				Exit Function
			End if
			wait 3
		End if	
		wait 1
		bGblFuncRetVal = Fn_SISW_UI_Object_Operations("Fn_SISW_BB_CreateDataset","Exist",objOpen,SISW_MICRO_TIMEOUT)
		If bGblFuncRetVal = true Then
			objOpen.WinEdit("FileName").Set sPath
			wait 1
			objOpen.WinButton("Open").Click micLeftBtn
'			bGblFuncRetVal = Fn_SISW_UI_WinEdit_Operations("Fn_SISW_BB_CreateDataset","Set",objOpen,"FileName",sPath)		
'			If bGblFuncRetVal = false Then Exit Function
'			wait 2
'			bGblFuncRetVal = Fn_SISW_UI_WinButton_Operations("Fn_SISW_BB_CreateDataset","click",objOpen,"Open","","","")
'			If bGblFuncRetVal = false Then Exit Function				
			wait 2		
		End If
 	End if	
	
	bGblFuncRetVal = Fn_SISW_UI_Object_Operations("Fn_SISW_BB_CreateDataset","Exist",objAddDataset,SISW_MIN_TIMEOUT)
	If bGblFuncRetVal = false Then Exit Function
		
	Select Case lcase(sAction)
		Case "create"
			If sRelationName <> "" then
				bGblFuncRetVal = Fn_SISW_UI_JavaList_Operations("Fn_SISW_BB_CreateDataset", "Select", objAddDataset, "RelationshipName", sRelationName, "", "")
				If bGblFuncRetVal = false then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select item ["+sRelationName+"] from javalist " + objAddDataset.JavaList("RelationshipName").ToString())
					Exit Function
				End if
			End if
			wait 1
		Case "verify_defaultrelname"
			sRelName = Fn_UI_Object_GetROProperty("Fn_SISW_BB_CreateDataset",objAddDataset.JavaList("RelationshipName"),"text")
			If sRelName <> sRelationName then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to retrive text from javalist " + objAddDataset.JavaList("RelationshipName").ToString())
				Exit Function
			End if	
			wait 1
	End select
	
	If sButton <> "" Then
		bGblFuncRetVal = Fn_SISW_UI_JavaButton_Operations("Fn_SISW_BB_CreateDataset", "Click", objAddDataset,sButton)
		If bGblFuncRetVal = false Then exit function			
	End If
	
	wait 1
	Fn_SISW_BB_CreateDataset = True
	set objBB = nothing
	Set objOpen = nothing
	set objAddDataset = nothing
End Function

'*********************************************************		Function to perform operations on dataset tab and tree***********************************************************************
'Function Name		:				Fn_SISW_BB_DataSetTree_Opeartions(sAction,sNode,sColType,sColName,sColValue)

'Description			 :		 		 This function perform variou operations on dataset tab and tree

'Parameters			   :	 			1. sAction : type of assembly to be open CAD or BB 
'										2. sNode: name of the tab under BB tree is displaying on which going to perform the operation
'										3. sColType: value of column "Type" which is displaying in the table.
'										4. sColName : name of the coulm from which needs to retrive/verify the value
'										5. sColValue : value to be verified against the column name
'
'Return Value		   : 				TRUE \ FALSE

'Pre-requisite			:		 		Node should be selected on which dataset is created

'Examples				:				 Fn_SISW_BB_DataSetTree_Opeartions("select","DS1Word/A","MSWord","DS1Word/A","","")
'										 Fn_SISW_BB_DataSetTree_Opeartions("Expand","DS1Word/A","MSWord","DS1Word/A","","")
'										 Fn_SISW_BB_DataSetTree_Opeartions("exist","DS1Word/A:DSWord.docx","MSWord","","","")
'										 Fn_SISW_BB_DataSetTree_Opeartions("verify_columnval","DS1Word/A","MSWord","Relationship Name","rendering","")
'										Fn_SISW_BB_DataSetTree_Opeartions("expand","DS1Word/A@4","","","")
'										Fn_SISW_BB_DataSetTree_Opeartions("exist","DS1Word/A@4:DS1_Word.docx","","","")
'										Fn_SISW_BB_DataSetTree_Opeartions("verify_columnval","child122982/A@4","MSWord","Relationship Name","manifestation")
'
'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Shweta Rathod			25-Jul-2016			1.0										Koustubh Watwe
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_BB_DataSetTree_Opeartions(sAction,sNode,sColType,sColName,sColValue,sRMBMenu)
	Dim bRet,objTab
	Dim iIndex,iRow,iCnt,sApp,sAppNode
	
	Fn_SISW_BB_DataSetTree_Opeartions = false
	
	set objBB = Fn_SISW_BB_GetObject("BriefcaseBrowser")
	set objTab = objBB.JavaTab("CTabFolder")
	Set objTree = objBB.JavaTree("BBTree2_VisRel_CTabFolder")
	
	objTab.SetTOProperty "value","Datasets" 
	bGblFuncRetVal = Fn_SISW_UI_Object_Operations("Fn_SISW_BB_DataSetTree_Opeartions","Exist",objTab,SISW_MICRO_TIMEOUT)
	If bGblFuncRetVal = true then
		bGblFuncRetVal = Fn_SISW_UI_JavaTab_Operations("Fn_SISW_BB_DataSetTree_Opeartions", "Select", objBB, "CTabFolder", "Datasets")
		If bGblFuncRetVal = false Then Exit Function
		wait 1		
	else
		Exit Function
	End if
	
	Select Case lcase(sAction)
		Case "select"
			bGblFuncRetVal = Fn_SISW_BB_UI_JavaTree_Operations("Fn_SISW_BB_DataSetTree_Opeartions","select",objTree,"",sNode,"","",sColType,"")
			If bGblFuncRetVal = false then exit function	
		Case "exist"
			bGblFuncRetVal = Fn_SISW_BB_UI_JavaTree_Operations("Fn_SISW_BB_DataSetTree_Opeartions","exist",objTree,"",sNode,"","",sColType,"")
			If bGblFuncRetVal = false then exit function	
		Case "expand"
			bGblFuncRetVal = Fn_SISW_BB_UI_JavaTree_Operations("Fn_SISW_BB_DataSetTree_Opeartions","expand",objTree,"",sNode,"","",sColType,"")
			If bGblFuncRetVal = false then exit function
			wait 1
		Case "collapse"		
			bGblFuncRetVal = Fn_SISW_BB_UI_JavaTree_Operations("Fn_SISW_BB_DataSetTree_Opeartions","collapse",objTree,"",sNode,"","",sColType,"")
			If bGblFuncRetVal = false then exit function
			wait 1
		Case "verify_columnval"
			bGblFuncRetVal = Fn_SISW_BB_UI_JavaTree_Operations("Fn_SISW_BB_DataSetTree_Opeartions","getcolumnval",objTree,"",sNode,sColName,sColValue,sColType,"")
			If bGblFuncRetVal <> sColValue Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to retrive column ["+sColName+"] value from " + sNode +" in javatree " + objTree.ToString())
				Exit Function
			End If	
			wait 1
		Case "doubleclick"
			bGblFuncRetVal = Fn_SISW_BB_UI_JavaTree_Operations("Fn_SISW_BB_DataSetTree_Opeartions","doubleclick",objTree,"",sNode,"","",sColType,"")
			If bGblFuncRetVal = false then exit function		
	End select	
	Fn_SISW_BB_DataSetTree_Opeartions = true
	set objTab = Nothing
	set objTree = Nothing
End Function

'*********************************************************		Function to Synchronize on Application Response	***********************************************************************
'Function Name		:				Fn_SISW_BB_Save(sFilename)

'Description			 :		 		Save the briefcase browser file 

'Parameters			   :	 			1. sFilename : fileName (should not be the full path)
'
'Return Value		   : 				Path of tree \ FALSE

'Pre-requisite			:		 		Briefcase browser should be open and all other dialoues inside application should close if open

'Examples				:				 Fn_SISW_BB_Save("fileName")

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Shweta Rathod			25-Jul-2016			1.0										Koustubh Watwe
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function Fn_SISW_BB_Save(sFilename)
Dim objbbSave,bRet

Fn_SISW_BB_Save = "false"
Set objbbSave =Dialog("SaveBriefcaseBrowser")

objbbSave.SetTOProperty "text", "Specify the name of the file in which to save briefcase"

strFilePath = Environment.Value("BBAssemblyPath")

bRet = fn_SISW_util_folder_operation("exist",strFilePath)
If bRet = false Then
	bRet = fn_SISW_util_folder_operation("createfolder",strFilePath)
	If bRet = false Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Function >> Fn_SISW_BB_Save failed to create folder [ " & strFilePath & " ]"  )
		Exit Function
	End if
End If

strFilePath = strFilePath+"\"+sFilename+".bcz"
bRet = Fn_SISW_UI_Object_Operations("Fn_SISW_BB_Save","Exist",objbbSave,SISW_MICRO_TIMEOUT)
if bRet = false then
	call Fn_UI_JavaMenu_Select("Fn_SISW_BB_Save",JavaWindow("BriefcaseBrowser"), Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("BriefcaseBrowser_Menu"), "FileSaveBriefcase"))
	wait 1
	bRet = Fn_SISW_UI_Object_Operations("Fn_SISW_BB_Save","Exist",objbbSave,SISW_MICRO_TIMEOUT)
	If  bRet = false Then Exit Function
End if 

bRet = Fn_SISW_UI_Object_Operations("Fn_SISW_BB_Save","Exist",objbbSave,SISW_MIN_TIMEOUT)
If bRet = true Then
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	strFilePath1 = objFSO.GetParentFolderName(strFilePath)
	objbbSave.WinEdit("FileName").Type strFilePath1
	Call Fn_KeyBoardOperation ("SendKeys", "{Enter}")
	objbbSave.WinEdit("FileName").Type objFSO.GetFileName(strFilePath)
	wait 1
	objbbSave.WinButton("Save").Click micLeftBtn
	wait 1
'	bRet = Fn_SISW_UI_WinEdit_Operations("Fn_SISW_BB_Save","Set",objbbSave,"FileName",strFilePath)
'	If bRet = false Then Exit Function
'	wait 2
'	
'	bRet = Fn_SISW_UI_WinButton_Operations("Fn_SISW_BB_Save","click",objbbSave,"Save","","","")
'	If bRet = false Then Exit Function
'	wait 5		
else
	Exit Function
End If

set objbbSave = Nothing
If bRet = true then Fn_SISW_BB_Save = strFilePath

End Function

'*************************************************  Function to perform various opeartions on the prefrence operation window ***********************************************************************
'Function Name		:				Fn_SISW_BB_PrefrenceOpeartion(sAction,dicPreference)

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
'										bReturn = Fn_SISW_BB_PrefrenceOpeartion("set_configuration",dicPref)
'										2. bReturn = Fn_SISW_BB_PrefrenceOpeartion("get_configuration","")

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Shweta Rathod			25-Jul-2016			1.0											Koustubh Watwe
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function Fn_SISW_BB_PrefrenceOpeartion(sAction,dicPreference)
Dim sMenu,objPrefrence,sAppText
	Fn_SISW_BB_PrefrenceOpeartion = false
	If sAction <> "get_configuration" then
		sMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("BriefcaseBrowser_Menu"), "WindowPrefrences")
		set objPrefrence = Fn_SISW_BB_GetObject("DlgPreferences")
		bGblFuncRetVal = Fn_SISW_UI_Object_Operations("Fn_SISW_BB_PrefrenceOpeartion","Exist",objPrefrence,SISW_MICRO_TIMEOUT)
		If bGblFuncRetVal = false then 
			call Fn_UI_JavaMenu_Select("Fn_SISW_BB_CreateDataset",JavaWindow("BriefcaseBrowser"),sMenu)
			wait 1
			bGblFuncRetVal = Fn_SISW_UI_Object_Operations("Fn_SISW_BB_PrefrenceOpeartion","Exist",objPrefrence,SISW_MIN_TIMEOUT)
			if bGblFuncRetVal = false then Exit Function
		End if
	End if
	Select Case lcase(sAction)
		Case "set_configuration"
			bGblFuncRetVal = Fn_SISW_UI_JavaEdit_Operations("Fn_SISW_BB_PrefrenceOpeartion", "Set", objPrefrence, "FilterText", dicPreference("PrefrenceName"))
			if bGblFuncRetVal = false then exit function
			wait 1
			bGblFuncRetVal = Fn_SISW_UI_JavaList_Operations("Fn_SISW_BB_PrefrenceOpeartion", "Select", objPrefrence, "ConfigurationName", dicPreference("ConfigurationName"), "", "")
			If bGblFuncRetVal = false Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to set configuration name  " + dicPreference("ConfigurationName") +" in javalist " + objPrefrence.JavaList("ConfigurationName").ToString())
				Exit Function
			End if
			wait 1
			bGblFuncRetVal = Fn_SISW_UI_JavaButton_Operations("Fn_SISW_BB_PrefrenceOpeartion", "Click", objPrefrence,"Apply")
			If bGblFuncRetVal = false Then Exit Function
			wait 1
			
			bGblFuncRetVal = Fn_SISW_UI_JavaButton_Operations("Fn_SISW_BB_PrefrenceOpeartion", "Click", objPrefrence,"OK")
			If bGblFuncRetVal = false Then Exit Function
			
		Case "get_configuration"
			Fn_SISW_BB_PrefrenceOpeartion = JavaWindow("BriefcaseBrowser").GetROProperty("title")
			Exit Function
			
		Case "verify_releasestatus"
			If dicPreference("TCMRelease") <> "" then
				objPrefrence.JavaList("ConfigurationName").SetTOProperty "attached text","Release Status Name"
				sAppText = objPrefrence.JavaList("ConfigurationName").GetROProperty("text")
				If trim(lcase(dicPreference("TCMRelease"))) = trim(lcase(sAppText)) then
					JavaWindow("BriefcaseBrowser").JavaWindow("Preferences").JavaButton("OK").SetTOProperty "label","Cancel"
					bGblFuncRetVal = Fn_SISW_UI_JavaButton_Operations("Fn_SISW_BB_PrefrenceOpeartion", "Click", objPrefrence,"OK")
					Fn_SISW_BB_PrefrenceOpeartion = True
				Else	
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to verify default value  " + dicPreference("TCMRelease") + " in javalist " + objPrefrence.JavaList("ConfigurationName").ToString())
					Exit Function
				End if
			else
				bGblFuncRetVal = Fn_SISW_UI_JavaButton_Operations("Fn_SISW_BB_PrefrenceOpeartion", "Click", objPrefrence,"Cancel")
				If bGblFuncRetVal = false Then Exit Function
			End if
			bGblFuncRetVal = Fn_SISW_UI_JavaButton_Operations("Fn_SISW_BB_PrefrenceOpeartion", "Click", objPrefrence,"OK")
			If bGblFuncRetVal = false Then Exit Function			
	End Select	
	Fn_SISW_BB_PrefrenceOpeartion = true
	Set objPrefrence = nothing
End Function


'*************************************************  Function to extract the content of ".bcz" file into given new directory ***********************************************************************
'Function Name		:				Fn_SISW_BB_ExtractContentBriefcaseBrowser(sFileName)

'Description			 :		 		to perform various opeartions on the java tab which is displaying in BB application

'Parameters			   :	 			1. sFileName : required file name in which extracting the content
'
'Return Value		   : 				FullPath of directory / "false"

'Pre-requisite			:		 		.bcz file should be present into the repective directory 

'Examples				:				bReturn = Fn_SISW_BB_ExtractContentBriefcaseBrowser("TestCaseName_iRandNo","")

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Shweta Rathod			01-Aug-2016			1.0										Koustubh Watwe
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function Fn_SISW_BB_ExtractContentBriefcaseBrowser(sFileName)
	Dim objShell: set objShell = CreateObject("Shell.Application")
	Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
	Dim stempFolderPath: stempFolderPath = Environment.Value("BBReportFolderPath")
	Dim objFolder,bRet
	Dim sSource,strRename,sDestination
	
	Fn_SISW_BB_ExtractContentBriefcaseBrowser = "false"
	sSource = Environment.Value("BBAssemblyPath")+"\"+sFilename+".bcz"
	strRename = Environment.Value("BBAssemblyPath")+"\"+sFilename+".zip"
	
	If NOT fso.FolderExists(stempFolderPath+"\"+sFilename) Then
		FSO.CreateFolder(stempFolderPath+"\"+sFilename)
		sDestination = stempFolderPath+"\"+sFilename
		fso.CopyFile sSource,sDestination+"\",True
		wait 5
	End If
	
	If not fso.FileExists(sDestination+"\"+sFilename+".bcz") then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Verified that the File ["+sFilename+"] is Copied at the Path ["+sDestination+"]")
		Set fso = Nothing
		Exit Function
	End if
	FSO.MoveFile sDestination+"\"+sFilename+".bcz", strRename
	
	call fn_SISW_util_folder_operation("deletefolder",sDestination)
	
	If NOT fso.FolderExists(stempFolderPath+"\ExtractedFiles") Then
		FSO.CreateFolder(stempFolderPath+"\ExtractedFiles")
	End If
	
	'Extract the contants of the zip file.
	set FilesInZip=objShell.NameSpace(strRename).items
	objShell.NameSpace(stempFolderPath+"\ExtractedFiles").CopyHere(FilesInZip)
	wait 10
	Set objFolder = fso.GetFolder(stempFolderPath+"\ExtractedFiles")
	If not(objFolder.Files.Count = 0 or objFolder.SubFolders.Count = 0) then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Verified that the File ["+sFilename+"] is Copied at the Path ["+sDestination+"]")
		Set fso = Nothing
		Exit Function
	End if
	
	Call fn_splm_util_file_operation("delete",strRename)
	Set fso = Nothing
	Fn_SISW_BB_ExtractContentBriefcaseBrowser = stempFolderPath+"\ExtractedFiles"
End Function


'*************************************************  Function to verify briefcase browser content ***********************************************************************
'Function Name		:				Fn_SISW_BB_VerifyBriefcaseBrowserContent(sFilePath,sFileName)

'Description			 :		 		verifying content of extracting directory

'Parameters			   :	 			1. sFilePath : required file name in which extracting the content
'										2. sFileName : 
'
'Return Value		   : 				FullPath of directory / "false"

'Pre-requisite			:		 		.bcz file should be extracted into the new directory 

'Examples				:				bReturn = Fn_SISW_BB_VerifyBriefcaseBrowserContent("C:\Temp\TestCaseName_iRandNo","child1.prt","")

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Shweta Rathod			01-Aug-2016			1.0										Koustubh Watwe
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function Fn_SISW_BB_VerifyBriefcaseBrowserContent(sFilePath,sFileName)
	Dim aFileName
	Fn_SISW_BB_VerifyBriefcaseBrowserContent = false
	
	if sFileName <> "" then aFileName = split(sFileName,"~")
	
	bGblFuncRetVal = fn_SISW_util_folder_operation("exist",sFilePath)
	If bGblFuncRetVal = false Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed - folder dose not exist at location ["+sFilePath+"]")
		Exit Function
	End if
	
	For iCnt = 0 to ubound(aFileName)
		bGblFuncRetVal = fn_splm_util_file_operation("exist",sFilePath+"\"+aFileName(iCnt))
		If bGblFuncRetVal = false then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Verified existence of File [ "+aFileName(iCnt)+" ] in the BB folder at location ["+sFilePath+"]")
			Exit Function
		End if
		wait 1
	next
	Fn_SISW_BB_VerifyBriefcaseBrowserContent = true
End Function

'*************************************************  Function to verify property of BB object from Item property tab ***********************************************************************
'Function Name		:				Fn_SISW_BB_ItemPropertyPanelVerify("Item Name","Item123")

'Description			 :		 		to perform various opeartions on toolbar

'Parameters			   :	 			1. sPropName : name of the property
'										2. sValue : value of property to be verify
'
'Return Value		   : 				true / false

'Pre-requisite			:		 		briefcase browser should be open

'Examples				:				Fn_SISW_BB_ItemPropertyPanelVerify("Item ID :~Object Name :~POM_imc","Top7186012121~Top71860~12345")))

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Shweta Rathod			01-Aug-2016			1.0											Koustubh Watwe
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function Fn_SISW_BB_ItemPropertyPanelVerify(sPropName,sValue)
	Dim objBB,objTab,objPropertyTree,aPropNameArr,aPropValueArr
	Dim iCount,sAppPropValue,bRet,bCounter
	
	set objBB = Fn_SISW_BB_GetObject("BriefcaseBrowser")
	set objTab = objBB.JavaTab("CTabFolder")
	objTab.SetTOProperty "value","Item Properties"
	
	Set objPropertyTree = objBB.JavaTree("BBTree2_VisRel_CTabFolder")
	bGblFuncRetVal = Fn_SISW_UI_Object_Operations("Fn_SISW_BB_ItemPropertyPanelVerify","Exist",objPropertyTree,SISW_MICRO_TIMEOUT)
	If bGblFuncRetVal = false then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Object does not exist "+objPropertyTree.tostring)
		Fn_SISW_BB_ItemPropertyPanelVerify = false
		Exit function
	End if
	
	bRet = Fn_SISW_UI_JavaToolbar_Operations("Fn_SISW_BB_ItemPropertyPanelVerify","isselected",JavaWindow("BriefcaseBrowser"),"","Show Advanced Properties","","","")
	If bRet = false then
		bRet = Fn_SISW_UI_JavaToolbar_Operations("Fn_SISW_BB_ItemPropertyPanelVerify","Click",JavaWindow("BriefcaseBrowser"),"","Show Advanced Properties","","","")
		wait 1
	End if
	
	bRet = Fn_SISW_UI_JavaToolbar_Operations("Fn_SISW_BB_ItemPropertyPanelVerify","isselected",JavaWindow("BriefcaseBrowser"),"","Show Categories","","","")
	If bRet = true then
		bRet = Fn_SISW_UI_JavaToolbar_Operations("Fn_SISW_BB_ItemPropertyPanelVerify","Click",JavaWindow("BriefcaseBrowser"),"","Show Categories","","","")
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
	Fn_SISW_BB_ItemPropertyPanelVerify = bret
End Function


'*********************************************************		Function to perfrom various opearations on BriefCase browser Tree	***********************************************************************
'Function Name		:				Fn_SISW_BB_BriefCaseTree_Opearation(sAction,sTabName,sNode)

'Description			 :		 		 This function perform variou operations on Briefcase browser tree

'Parameters			   :	 			1. sAction : type of assembly to be open CAD or BB 
'										2. sTabName: name of the tab under BB tree is displaying on which going to perform the operation
'										3. sNode: full path of the node on which performing the opeation
											
'Return Value		   : 				TRUE \ FALSE

'Pre-requisite			:		 		Brief case browser application should be displayed

'Examples				:				 Fn_SISW_BB_BriefCaseTree_Opearation("select","*C:\mainline\Reports\NX\Add_Attachs_Default_Relat_88663\Top88663_A.prt","Top88663:child188663","")
'										Fn_SISW_BB_BriefCaseTree_Opearation("verifycelldata","*C:\mainline\Reports\NX\Add_Attachs_Default_Relat_88663\Top88663_A.prt","Top88663:child188663","CAD Attached","","")
'Fn_SISW_BB_BriefCaseTree_Opearation("verifycelldata","*C:\mainline\Reports\NX\Add_Attachs_Default_Relat_88663\Top88663_A.prt","Top88663:child188663","Name","Top123","")

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Shweta Rathod			25-Jul-2016			1.0										Koustubh Watwe
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_BB_BriefCaseTree_Opearation(sAction,sTabName,sNode, sColumn, sValue, sRMBMenu)
	Dim bRet,objTab,iRow,iIndex
	Dim iColPos,iPath,arrPath,iArrCnt,objTemp,sColVal
	Dim objTree,objDialog
	Fn_SISW_BB_BriefCaseTree_Opearation = false
	set objDialog = Fn_SISW_BB_GetObject("BriefcaseBrowser")
	Set objTree = objDialog.JavaTree("BBTree_VisRel_CTabFolder")
	set objTab = objDialog.JavaTab("CTabFolder")
	
	bRet = Fn_UI_Object_SetTOProperty_ExistCheck("Fn_SISW_BB_BriefCaseTree_Opearation",objTab,"value",sTabName)
	If bRet = true then
		bRet = Fn_UI_JavaTab_Select("Fn_SISW_BB_BriefCaseTree_Opearation",JavaWindow("BriefcaseBrowser"),"CTabFolder", sTabName)
		wait 1
		If bRet = false Then Exit Function
	else
		Exit Function
	End if
	
	Select Case lcase(sAction)
		Case "select"
			bRet = Fn_SISW_BB_UI_JavaTree_Operations("Fn_SISW_BB_BriefCaseTree_Opearation","select",objTree,"",sNode,"","","","")
			If bRet = false then exit function		
		Case "exist"
			bRet = Fn_SISW_BB_UI_JavaTree_Operations("Fn_SISW_BB_BriefCaseTree_Opearation","exist",objTree,"",sNode,"","","","")
			If bRet = false then exit function			
		Case "expand"
			bRet = Fn_SISW_BB_UI_JavaTree_Operations("Fn_SISW_BB_BriefCaseTree_Opearation","expand",objTree,"",sNode,"","","","")
			If bRet = false then exit function
			wait 1
		Case "collapse"		
			bRet = Fn_SISW_BB_UI_JavaTree_Operations("Fn_SISW_BB_BriefCaseTree_Opearation","collapse",objTree,"",sNode,"","","","")
			If bRet = false then exit function
			wait 1
		Case "verifycelldata"
			Select Case sColumn
				Case "CAD Attached", "JT Attached", "Read Only"
					'If isobject(objTemp.getImage(iColPos)) = False Then
					bRet = Fn_SISW_BB_UI_JavaTree_Operations("Fn_SISW_BB_BriefCaseTree_Opearation","verifytickmark",objTree,"",sNode,sColumn,"","","")
					If bRet = false then exit function
				Case else
					bRet = Fn_SISW_BB_UI_JavaTree_Operations("Fn_SISW_BB_BriefCaseTree_Opearation","getcolumnval",objTree,"",sNode,sColumn,"","","")
					If sValue <> bRet then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to verify column value " + sColumn +" = " + sColVal +" in javatree " + objTree.ToString())
						Exit Function
					End if
			End Select	
		Case "popupmenu"
			bRet = Fn_SISW_BB_UI_JavaTree_Operations("Fn_SISW_BB_BriefCaseTree_Opearation","popupmenu",objDialog,"BBTree_VisRel_CTabFolder",sNode,sColumn,"","",sRMBMenu)
			If bRet = false then exit function
	End select
	Fn_SISW_BB_BriefCaseTree_Opearation = true
	set objTab = Nothing
	set objTree = Nothing
	Set objDialog = nothing
End Function

'*********************************************************	Function to perfrom various opearations on Error log tree ***********************************************************************
'Function Name		:				Fn_SISW_BB_ErrorLogTree_Opeartions(sAction,discError)

'Description			 :		 		 This function perform to perfrom various opearations on Error log tree

'Parameters			   :	 			1. sAction : case to be execute
'										2. discError: parameters to perform opeartion
											
'Return Value		   : 				TRUE \ FALSE

'Pre-requisite			:		 		Brief case browser application should be displayed

'Examples				:				bReturn = Fn_SISW_BB_ErrorLogTree_Opeartions("isEmpty","")

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Shweta Rathod			25-Jul-2016			1.0										Koustubh Watwe
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_BB_ErrorLogTree_Opeartions(sAction,discError)
Dim iRow,bret,objTab,objTree
	Fn_SISW_BB_ErrorLogTree_Opeartions = false
	set objTab = Fn_SISW_BB_GetObject("BriefcaseBrowser")
	Set objTree = objTab.JavaTree("BBTree2_VisRel_CTabFolder")
	set objTab = objTab.JavaTab("CTabFolder")
	bret =  Fn_UI_Object_SetTOProperty_ExistCheck("Fn_SISW_BB_ErrorLogTree_Opeartions",objTab,"value","Error Log")
	If bret = false then Exit Function
	
	Select Case sAction
		Case "isEmpty"
			bRet = Fn_SISW_BB_UI_JavaTree_Operations("Fn_SISW_BB_ErrorLogTree_Opeartions","isempty",objTree,"","","","","","")
			If bRet = false then exit function
	End Select
	
	set objTab = nothing
	set objTree = nothing
	Fn_SISW_BB_ErrorLogTree_Opeartions = true
End function




'*********************************************************	Function to perfrom various opearations on ownership transfer tree  ***********************************************************************
'Function Name		:				Fn_SISW_BB_OwnershipTransferTree_Opeartions(sAction,sPartNumber,sColName,sColVal,sRMBMenu)

'Description			 :		 		 This function perform to perfrom various opearations on ownership transfer tree

'Parameters			   :	 			1. sAction : case to be execute
'										2. sPartNumber: Part number (multiple partname should be tilda (~) seperated)
'										3. sColName: clumn to be verified (this should be only one)
'										4. sColVal: value to be verify against column and partnumber (multiple partname should be tilda (~) seperated
'										5. sMenu: menu opeartion if any 
											
'Return Value		   : 				TRUE \ FALSE

'Pre-requisite			:		 		Brief case browser application should be displayed

'Examples				:				bReturn =  Fn_SISW_BB_OwnershipTransferTree_Opeartions("verifyownershiptrans","assembly1~Sub4Ch22~Sub1Ch1~Ch1","Name","assembly1~Sub4Ch2kk~Sub1Ch1~Ch1","")

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Shweta Rathod			21-Sep-2016			1.0										
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_BB_OwnershipTransferTree_Opeartions(sAction,sPartNumber,sColName,sColVal,sMenu)
	Dim objJavaDlg,objTree,objTab,iNodeItemsCount
	Dim bret,aPartNumArr,bFlag,iCnt,jCnt,sAppColVal,aColVal
	Fn_SISW_BB_OwnershipTransferTree_Opeartions = false
	set objJavaDlg = Fn_SISW_BB_GetObject("BriefcaseBrowser")
	Set objTree = objJavaDlg.JavaTree("OT_VisRel_OwnershipTransfer")
	set objTab = objJavaDlg.JavaTab("OwnershipTransfer")
	bret = Fn_SISW_UI_JavaTab_Operations("Fn_SISW_BB_OwnershipTransferTree_Opeartions", "Select", JavaWindow("BriefcaseBrowser"), "OwnershipTransfer","Ownership Transfer")
	wait 1
	If bret = false Then Exit Function
	
	Select Case lcase(sAction)
		Case "verifyownershiptrans"
			aPartNumArr = split(sPartNumber,"~")
			aColVal = split(sColVal,"~")
			For iCnt = 0 to ubound(aPartNumArr)	
				bFlag = false
				iNodeItemsCount = objTree.Object.getItemCount()
				For jCnt=0 to iNodeItemsCount-1
					If aPartNumArr(iCnt) = objTree.GetItem(jCnt) Then
						bFlag = true
						sAppColVal = Cstr(objTree.GetColumnValue("#"+cstr(jCnt),sColName))
						If aColVal(iCnt) <> sAppColVal then
							Exit function
						End if
						Exit for
					End If
				next
				if bFlag = false then 
					Exit function
				End if
			next			
	End Select
	
	set objTab = nothing
	set objTree = nothing
	set objJavaDlg = nothing
	Fn_SISW_BB_OwnershipTransferTree_Opeartions = true
End function

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name	:	Fn_SISW_BB_OwnershipTransferNATTable_Operations
'@@
'@@    Description		:	Function Used to Check or Uncheck or Verify row in Ownership Transfer window in BB
'@@
'@@    Parameters		:	1. sAction		: Action name
'@@						:	2. sColNames	: Column names
'@@						:	3. sColValues 	: Column values
'@@						:	4. dicDetails 	: Dictionary obejct
'@@						:	5. sButton 		: OK / Cancel
'@@						:	6. sReserve 	: Future use
'@@
'@@    Return Value		: 	True Or False or RowNumber or CellValue or -1
'@@
'@@    Examples			:   bReturn = Fn_SISW_BB_OwnershipTransferNATTable_Operations("VerifyRow","Transfer~Part Number~Name","false~Top69353~Top69353","","OK","")
'@@    						bReturn = Fn_SISW_BB_OwnershipTransferNATTable_Operations("GetRowNumber","Transfer~Part Number~Name","false~Top69353~Top69353","","Cancel","")
'@@    						Set dicDetails = CreateObject("Scripting.Dictionary")
'@@    							dicDetails("Checked") = "True"
'@@    						bReturn = Fn_SISW_BB_OwnershipTransferNATTable_Operations("SetChecked","Part Number~Name","Top69353~Top69353",dicDetails,"","")
'@@    						bReturn = Fn_SISW_BB_OwnershipTransferNATTable_Operations("VerifyChecked","Part Number~Name","Top69353~Top69353",dicDetails,"","")
'@@    							
'@@	   History			:	
'@@			Developer Name		Date	  		Rev. No.	Changes Done										Reviewer
'@@-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@			Vivek Ahirrao		26-Sep-2016		1.0		  	Created												[BB1123-20160608-26_09_2016-VivekA-NewDevelopment]
'@@			shweta Rathod		21-oct-2016		1.1			Modified										[BB1123-20160608-21_10_2016-ShwetaR-NewDevelopment]
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_SISW_BB_OwnershipTransferNATTable_Operations(sAction,sColNames,sColValues,dicDetails,sButton,sReserve)
	Dim objOwnerTransWindow, objNATTable
	Dim bFlag, sRectangle, sColIntName, sAppText
	Dim aBounds, aColName, aColValue
	Dim iRowNumber, iColNumber, iRowCount, iColCount, iCount, iCount1,iCnt
	
	On Error Resume Next
	Fn_SISW_BB_OwnershipTransferNATTable_Operations = False
	
	Set objOwnerTransWindow = JavaWindow("Shell").JavaWindow("OwnershipTransfer")
	For iCnt = 0 to 10
		JavaWindow("Shell").SetTOProperty "Index",iCnt
		if objOwnerTransWindow.Exist(1) then Exit for
	next
	If objOwnerTransWindow.Exist(5) = False Then
		'Node in BB Tree should be selected to perform Menu operation
		bFlag = Fn_UI_JavaMenu_Select("Fn_SISW_BB_OwnershipTransferNATTable_Operations",JavaWindow("BriefcaseBrowser"),"Tools:Ownership Transfer")
		If bFlag = False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [Fn_SISW_BB_OwnershipTransferNATTable_Operations] Failed to Perform menu operation [Tools:Ownership Transfer].")
			Set objOwnerTransWindow = Nothing
			Exit Function
		End If
	End If
	For iCnt = 0 to 10
		JavaWindow("Shell").SetTOProperty "Index",iCnt
		if objOwnerTransWindow.Exist(1) then Exit for
	next
	'Maximize window
	If objOwnerTransWindow.Exist(5) Then
		If objOwnerTransWindow.GetROProperty("maximized") <> "1" Then
			objOwnerTransWindow.Maximize
			Wait SISW_MICRO_TIMEOUT
		End If
	else
		Exit function
	End If
	
	Set objNATTable = objOwnerTransWindow.JavaObject("OTNatTable")
	
	Select Case sAction
		Case "GetCellValue"
				iRowNumber = sColNames
				iColNumber = sColValues
				If iRowNumber<>"" AND iColNumber<>"" Then
					Fn_SISW_BB_OwnershipTransferNATTable_Operations = objNATTable.Object.getCellByPosition(iColNumber,iRowNumber).getDataValue.tostring
				End If
				Set objNATTable = Nothing				
		Case "SetChecked","VerifyChecked","SetCheckedAll"
				If (sColNames<>"" AND sColValues<>"") OR sAction = "SetCheckedAll" Then
					'GetRowNumber of provided column values
					If sAction = "SetCheckedAll" Then
						iRowNumber = 0
					Else
						iRowNumber = Fn_SISW_BB_OwnershipTransferNATTable_Operations("GetRowNumber",sColNames,sColValues,"","","")
					End If
					If iRowNumber=-1 Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Row not Found.")
						Set objNATTable = Nothing
						Set objOwnerTransWindow = Nothing
						Exit Function
					End If
					'GetColumnNumber of "Transfer" Column to set Checkbox ON or OFF
					iColNumber = Fn_SISW_BB_OwnershipTransferNATTable_Operations("GetColumnNumber","Transfer","","","","")
					If iColNumber = -1 Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Column [ Transfer ] not Found.")
						Set objNATTable = Nothing
						Set objOwnerTransWindow = Nothing
						Exit Function
					End If
					
					'Check the checkbox required is already checked or not
					If dicDetails("Checked")<>"" Then
						If sAction = "SetCheckedAll" Then
							iRowCount = objNATTable.Object.getrowcount
							For iCount = 2 To CInt(iRowCount)-1
								sAppText = Fn_SISW_BB_OwnershipTransferNATTable_Operations("GetCellValue",iCount,iColNumber,"","","")
								If LCase(CStr(dicDetails("Checked"))) <> LCase(sAppText) Then
									bFlag = False
									Exit For
								Else
									bFlag = True
								End If
							Next
						Else
							sAppText = Fn_SISW_BB_OwnershipTransferNATTable_Operations("GetCellValue",iRowNumber,iColNumber,"","","")
							If LCase(CStr(dicDetails("Checked"))) = LCase(sAppText) Then
								bFlag = True
							Else
								bFlag = False
							End If
						End If
					End If
					
					If bFlag = False AND (sAction = "SetChecked" OR sAction = "SetCheckedAll") Then
						'Set checked or unchecked
						sRectangle = objNATTable.Object.getBoundsByPosition(iColNumber,iRowNumber).tostring
						sRectangle = Right(sRectangle, (Len(sRectangle) - Instr(1, sRectangle, "{", 1)))	
						sRectangle = Replace(sRectangle,"}","")
						sRectangle = Replace(sRectangle," ","")
						aBounds = Split(sRectangle,",")
						
						If sAction = "SetCheckedAll" Then
							objNATTable.Click cInt(aBounds(0)) + (cInt(aBounds(2))-10),  cInt(aBounds(1)) + (cInt(aBounds(3))/2)  , "LEFT"
						Else
							objNATTable.Click cInt(aBounds(0)) + (cInt(aBounds(2))/2),  cInt(aBounds(1)) + (cInt(aBounds(3))/2)  , "LEFT"
						End If
						If Err.Number<0 Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Click on Row to Set CheckBox.")
							Set objNATTable = Nothing
							Set objOwnerTransWindow = Nothing
							Exit Function
						End If
						Wait SISW_MICRO_TIMEOUT
						bFlag = True
					End If
					If bFlag = True Then
						Fn_SISW_BB_OwnershipTransferNATTable_Operations = True
					End If
				End If
				Set objNATTable = Nothing
		Case "GetColumnNumber"
				If sColNames<>"" Then
					iColNumber = -1
					iColCount = objNATTable.Object.getColumncount
					For iCount = 0 To CInt(iColCount)-1
						If IsObject(objNATTable.Object.getCellByPosition(iCount,0)) Then
							If IsObject(objNATTable.Object.getCellByPosition(iCount,0).getDataValue) Then
								sAppText = objNATTable.Object.getCellByPosition(iCount,0).getDataValue.tostring
								If sAppText = sColNames Then
									iColNumber = iCount
									Exit For
								End If
							End If
						End If
					Next
					If iColNumber<>-1 Then
						Fn_SISW_BB_OwnershipTransferNATTable_Operations = iColNumber
					Else
						Fn_SISW_BB_OwnershipTransferNATTable_Operations = -1
					End If
				End If
				Set objNATTable = Nothing
		Case "VerifyRow","GetRowNumber"
				If sAction="GetRowNumber" Then
					Fn_SISW_BB_OwnershipTransferNATTable_Operations = -1
				End If
				If sColNames<>"" AND sColValues<>"" Then
					iRowNumber = -1
					iRowCount = objNATTable.Object.getrowcount
					aColName = Split(sColNames,"~")
					aColValue = Split(sColValues,"~")
					For iCount = 2 To CInt(iRowCount)-1
						For iCount1 = 0 To UBound(aColName)
							bFlag = False
							sColIntName = aColName(iCount1)
							If sColIntName<>"" Then
								'GetColumnIndex of column
								iColNumber = Fn_SISW_BB_OwnershipTransferNATTable_Operations("GetColumnNumber",sColIntName,"","","","")
								If iColNumber = -1 Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Column ["+sColIntName+"] not Found.")
									Set objNATTable = Nothing
									Set objOwnerTransWindow = Nothing
									Exit Function
								End If
								'Get Cell value
								If IsObject(objNATTable.Object.getCellByPosition(iColNumber,iCount)) Then
									If IsObject(objNATTable.Object.getCellByPosition(iColNumber,iCount).getDataValue) Then
										sAppText = objNATTable.Object.getCellByPosition(iColNumber,iCount).getDataValue.tostring
										If sColIntName = "Transfer" Then
											If LCase(sAppText) <> LCase(aColValue(iCount1)) Then
												bFlag = False
												Exit For
											End If
										Else
											If sAppText <> aColValue(iCount1) Then
												bFlag = False
												Exit For
											End If
										End If
									End If
								End If
								bFlag = True
							End If
						Next
						'If Found Row with column values
						If bFlag = True Then
							iRowNumber = iCount
							Exit For
						End If
					Next
					
					If bFlag = False Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Row not Found.")
						Set objNATTable = Nothing
						Set objOwnerTransWindow = Nothing
						Exit Function
					End If
					If sAction = "GetRowNumber" Then
						Fn_SISW_BB_OwnershipTransferNATTable_Operations = iRowNumber
					Else
						Fn_SISW_BB_OwnershipTransferNATTable_Operations = True
					End If
				End If
				Set objNATTable = Nothing
	End Select
	If sButton<>"" Then
		Call Fn_Button_Click("Fn_SISW_BB_OwnershipTransferNATTable_Operations",objOwnerTransWindow,sButton)
	End If
	Set objOwnerTransWindow = Nothing
End Function


'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
'
'    Function Name	:	Fn_SISW_BB_ErrorVerify
'
'    Description		:	Function Used to verify errors in BB
'
'    Parameters		:	1. sAction		: Action name
'						:	2. objErr		: error dialog
'						:	3. dicDetails 	: Dictionary obejct
'
'    Return Value		: 	True Or False 
'
'    Examples			:   
'    						Set dicErrorInfo = CreateObject("Scripting.Dictionary")
'    							dicErrorInfo("DialogTitle") = "Ownership Transfer"
'								dicErrorInfo("ButtonName") = "OK"
'								dicErrorInfo("ErrMessage") = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("BB_ErrorMessage"),"OwnershipTransErr")
'    						bReturn = Fn_SISW_BB_ErrorVerify("verifyownershiptranserr",objErr,dicErrorInfo)
'    							
'	   History			:	
'		Developer Name		Date	  		Rev. No.	Changes Done										Reviewer
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
'			Shweta Rathod		30-Sep-2016		1.0		  	Created												[BB1123-20160608-30_09_2016-ShwetaR-NewDevelopment]
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function Fn_SISW_BB_ErrorDialogVerify(sAction,objErr,dicErrorInfo)
Dim  objStaticText,intNoOfObjects,bFound
	Fn_SISW_BB_ErrorDialogVerify = false
	bFound = false
	'set title of error dialog
	If dicErrorInfo("DialogTitle") <> "" then
		objErr.SetTOProperty "title", dicErrorInfo("DialogTitle")
	End if
	
	bReturn =  Fn_SISW_UI_Object_Operations("Fn_SISW_BB_ErrorVerify","Exist",objErr,SISW_MICRO_TIMEOUT) 
	If bReturn = False Then exit function
	
	Select Case lcase(sAction)
		Case "verifyownershiptranserr" 'this case added because for static text message contain crlf  object is not able identify by SETTO method so used descriptive programming
			Set objStaticText=description.Create()
			objStaticText("Class Name").value = "JavaStaticText"
			Set  intNoOfObjects = objErr.ChildObjects(objStaticText)
			For i = 0 to intNoOfObjects.count-1
			   If  trim(intNoOfObjects(i).getROProperty("text")) = trim(dicErrorInfo("ErrMessage")) Then				
					bFound = true
					Exit for
			   End If
			Next
			If bFound = false then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to verify error message [" + dicErrorInfo("ErrMessage") + "] with expected value ["+sAppMsg+"]")      									
			End If
	End Select
	
	If dicErrorInfo("ButtonName") <> "" then
		bRet = Fn_SISW_UI_JavaButton_Operations("Fn_SISW_BB_ErrorDialogVerify", "Click", objErr,dicErrorInfo("ButtonName")) 
		If bRet = false Then Exit function
	End if
	
	if bFound = false then
		Fn_SISW_BB_ErrorDialogVerify = true	
		Exit function
	End if
Fn_SISW_BB_ErrorDialogVerify = true	
Set intNoOfObjects = nothing
Set objStaticText = nothing
End Function
