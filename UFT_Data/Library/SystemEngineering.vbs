Option Explicit
'Public  iTimeOut
iTimeOut=180
'-------------------------------'Global variables for Teamcenter Perspective Names-------------------------------------------------------
Public GBL_PERSPECTIVE_SYSTEMENGINERRING
GBL_PERSPECTIVE_SYSTEMENGINERRING = "Systems Engineering"
'-------------------------------'Global variables for Teamcenter Perspective Names-------------------------------------------------------

'0. Fn_SISW_SE_GetObject()
'1.	Fn_SE_RequirementSpecCreate()
'2.	Fn_SE_NewRequirementCreate()
'3.	Fn_SE_NewParagraphCreate()
'4.	Fn_SE_BOMTable_RowIndex()
'5.	Fn_SE_BOMTableNodeOpeations()
'6.	Fn_SE_TraceabilityReportOperations()
'7.	Fn_SE_ImportReqSpec()
'8.	Fn_SE_DialogMsgVerify()
'9.	Fn_SE_ViewerTabOperations()
'10.Fn_SE_RandNoGenerate()
'11.Fn_SE_CustomRequirementCreate()
'12. Fn_SE_MSWordTabOperations()
'13.Fn_SE_ParamatricValueOperation()
'14.Fn_SE_QuickPanelOperation()
'15.Fn_SE_CustomizeIWantTo()
'16.Fn_SE_TraceabilityReportColumnOperations()
'17.Fn_SE_NavTree_NodeOperation()
'18.Fn_SE_DetailTableOperation()
'19.Fn_SE_DetailTableSort()
'20.Fn_SE_SaveColumnConfiguration()
'21.Fn_SE_ApplyColumnConfiguration()
'22.Fn_SE_DetailsTableFilterManagement()
'23.Fn_SE_CAEItemBasicCreate()
'24.Fn_SE_DetailTableConfigOperation()
'25.Fn_SE_ExportToExcel()
'26.Fn_SE_TraceLinkOpeartions()
'27.Fn_SE_OpenByNameOperations()
'28.Fn_SE_MSWordTabOperationsExt()
'29.Fn_SE_ItemFromTemplateOperations()
'30.Fn_SE_ErrorMessageVerify()
'31.Fn_SE_RightPanelTabOperations()
'32.Fn_SE_CustomNoteCreate()
'33. Fn_SE_AttachmentTableNodeOperation()
'34. Fn_SE_CreateNewRequirementWithProject()
'35. Fn_SE_AccountiblityCheck()
'36. Fn_SE_DetachedTraceabilityOperations()
'37. Fn_SE_SaveViewConfiguration()
'38. Fn_SE_ReportGenerationWizard()
'39 . Fn_SE_NewDerivedRequirementCreate()
'40. Fn_SE_SaveColumnConfigurationFromBomLine()
'41. Fn_SE_ApplyViewConfiguration()
'42. Fn_SE_DiagramOperations()
'43. Fn_SE_DataDictionarySearchDialogOperations()
'44. Fn_SE_DataDictionarySearchTreeOperations()
'45. Fn_SE_ShowTraceabilityMatrix()
'46. Fn_SE_TraceabilityMatrixPanelOperations()
'45. Fn_SE_NotesOperations()
'46. Fn_SE_RevRuleCreateWithExitingOne()
'47. Fn_SE_FilterSetting()
'48. Fn_SISW_SE_DetailsTable_GetCellData()
'49. Fn_SISW_SE_SetFilterDescription()
'50. Fn_SISW_SE_BomCompareStructureOperation
'51. Fn_SISW_SE_Create_And_OpenWebPage()
'52. Fn_SE_ComponentAndSEBOMTableNodeOpeations()
'53. Fn_SISW_SE_PropertiesOperations()
'54. Fn_SISW_SE_SourceTargetTableOperations()
'55. Fn_SISW_SE_SourceTargetTable_GetCellData()
'56. Fn_SISW_SE_ErrorVerify()
'57. Fn_SISW_SE_RequirementDetailsCreate()
'58. Fn_SISW_SE_BudgetDefinitionDetailsCreate()
'59. Fn_SISW_SE_EditBudgetOperations()
'60. Fn_SISW_SE_BudgetsTableOperations()
'61. Fn_SISW_SE_OpenDiagram()
'62. Fn_SISW_SE_ShapeDeleteConfirmation()
'63. Fn_SISW_SE_TraceLinkTabOperations()
'64. Fn_SE_WarningMsgVerify()
'65  Fn_SISW_SE_TraceLinkCriteriaOperations()
'66. Fn_SISW_SE_ColumnConfigurationOperation()
'67. Fn_SISW_SE_ColumnManagementOperation()
'68. Fn_SISW_SE_ShowTraceabilityMatrixOperation()
'69. Fn_SE_SnapshotCreate()
'70. Fn_SISW_SE_DeleteTraceLinks()
'71. Fn_SE_ReplaceOperation
'72. Fn_SE_RemoveLevel
'73. Fn_SE_ProcessHistoryTabOperations
'74. Fn_SE_SelectConfigurationOperations
'75. Fn_SE_BOMCompareReportTableOperations
'76. Fn_SE_ConfigurationInformationOperations
'77. Fn_SE_TraceabilityTabOperations
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 			Function to get Object hierarchy  		- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''Function Name		:	Fn_SISW_SE_GetObject
'
''Description		  	 :  	Function to get Object hierarchy

''Parameters		   :	1. sObjectName : Object Handle name
								
''Return Value		   :  	Object \ Nothing
'
''Examples		     	:	Fn_SISW_SE_GetObject("CMEBOMTreeTable")

'History:                
'								Developer Name							Date				Rev. No.		Reviewer		Changes Done	
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'									Sonal Padmawar		 				12June-2012				1.0					Sunny
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'-----------------------------------------------------------------------------------------------------------------------------------
'	Anurag Khera		 25-Sept-2013		2.0								Externalized object hierarchies
'-----------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_SE_GetObject(sObjectName)
	Dim sObjectXMLPath
	sObjectXMLPath = Fn_GetEnvValue("User", "AutomationDir") & "\TestData\AutomationXML\ObjectXML\SystemEngineering.xml"
	Set Fn_SISW_SE_GetObject = Fn_SISW_Setup_GetObjectFromXML(sObjectXMLPath, sObjectName)
		
End Function
'*********************************************************		Function to Create Specification in SE ***********************************************************************
'Function Name		:				Fn_SE_RequirementSpecCreate

'Description			 :		 		 This function is used to Create the Specification in System Engineering

'Parameters			   :	 			1. strNodeName: Select the Requirement Spec
'													2. strSpecID: ID of the Specification
'												   3. strSpecRev: Revision of the Spec
'												  4. strSpecName: Name of the Spec
'												  5.strSpecDesc: Description of the Spec

'Return Value		   : 				TRUE \ FALSE

'Pre-requisite			:		 		should be in SE Prespective

'Examples				:				 Call Fn_SE_RequirementSpecCreate("RequirementSpec","","","NewSpec","Description")

'History:
'										Developer Name			Date			Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Amol								24.01.2011																			Tushar
'									Harshal Agrawal				14.02.2011
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function Fn_SE_RequirementSpecCreate(strNodeName,strSpecID,strSpecRev,strSpecName,strSpecDesc)
	GBL_FAILED_FUNCTION_NAME="Fn_SE_RequirementSpecCreate"
	Dim strNodePath,intNodeCount,sTreeItem,intCount,strNodePathC
   Dim ObjSpecWnd,bFlag,objTemp,iRowIndex
   
   Set objTemp = Fn_SISW_SE_GetObject("RequirementSpec")
   If strNodeName="RequirementSpec" Then
		strNodeName="Requirement Specification"
   End If
	Fn_SE_RequirementSpecCreate=False
'	If Fn_UI_ObjectExist("Fn_SE_RequirementSpecCreate",JavaWindow("SystemsEngineering").JavaWindow("RequirementSpec"))=False Then
	If Fn_SISW_UI_Object_Operations("Fn_SE_RequirementSpecCreate", "Exist", objTemp, SISW_MICRO_TIMEOUT)=False then
		Call Fn_MenuOperation("Select",Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("RM_Menu"),"FileNewRequirementsSpec"))
	End If
  	Call Fn_ReadyStatusSync(2)
	Set ObjSpecWnd=Fn_UI_ObjectCreate("Fn_SE_RequirementSpecCreate",objTemp)
	Call Fn_UI_JavaTree_Expand("Fn_SE_RequirementSpecCreate", ObjSpecWnd, "RequirementSpecification","Complete List")
	ObjSpecWnd.JavaTree("RequirementSpecification").WaitProperty "items count" , micGreaterThan(1)
	If Fn_UI_JavaTree_NodeExist("Fn_SE_RequirementSpecCreate",ObjSpecWnd.JavaTree("RequirementSpecification"),"Complete List:"+strNodeName) Then
			strNodePath="Complete List:"+strNodeName
	Else
			strNodePath="Most Recently Used:"+strNodeName
	End If

   Call Fn_JavaTree_Select("Fn_SE_RequirementSpecCreate", ObjSpecWnd, "RequirementSpecification",strNodePath)
'   Call Fn_JavaTree_Select("Fn_SE_RequirementSpecCreate", ObjSpecWnd, "RequirementSpecification","Complete List")
'   Call Fn_JavaTree_Select("Fn_SE_RequirementSpecCreate", ObjSpecWnd, "RequirementSpecification",strNodePath)
   Call Fn_Button_Click("Fn_SE_RequirementSpecCreate",ObjSpecWnd,"Next")
   'Added Sync by Nilesh on 4-March-2013
   Call Fn_ReadyStatusSync(2)
   wait 2
	If strSpecID<>"" Then
		'Setting Id
		'Call Fn_Edit_Box("Fn_SE_RequirementSpecCreate",ObjSpecWnd,"ID",strSpecID)
	End If
	If strSpecRev<>"" Then
		'Setting Revision
		'Call Fn_Edit_Box("Fn_SE_RequirementSpecCreate",ObjSpecWnd,"Revision",strSpecRev)
	End If
	'Setting Name
    Call Fn_Edit_Box("Fn_SE_RequirementSpecCreate",ObjSpecWnd,"Name",strSpecName)
	'Setting Description
	Call Fn_Edit_Box("Fn_SE_RequirementSpecCreate",ObjSpecWnd,"Description",strSpecDesc)
	'Clicking On Finish Button To finish the Operation
	Call Fn_Button_Click("Fn_SE_RequirementSpecCreate",ObjSpecWnd,"Finish")

	Call Fn_ReadyStatusSync(1)
	'Click on Cancel Button
	If ObjSpecWnd.Exist(1) = True Then
	     ObjSpecWnd.JavaButton("Cancel").Click micLeftBtn
	End If
	Call Fn_ReadyStatusSync(2)
	iRowIndex = DataTable.GetCurrentRow()
	iRowNo = Fn_SE_BOMTable_RowIndex(strSpecName)
	DataTable.GetSheet("Global").AddParameter "NewReqSpecID",""
	DataTable.GetSheet("Global").AddParameter "NewReqSpecRev",""
	sFirstChildReq=Fn_SE_BOMTableNodeOpeations("GetCellData",0,0,"","")
	DataTable.SetCurrentRow(iRowIndex)
	If sFirstChildReq <> "" Then
		aItmInfo = split(sFirstChildReq, "/", -1, 1)
		DataTable("NewReqSpecID",dtGlobalSheet) = "'"&aItmInfo(0)
		aItmInfo1= split(aItmInfo(1), ";", -1, 1)
		DataTable("NewReqSpecRev",dtGlobalSheet) = aItmInfo1(0)
	End If
	'function Return True
	Fn_SE_RequirementSpecCreate=True
	'Releasing "New Specification" window's object
	Set ObjChangeWnd=Nothing
	Set objTemp=Nothing
End Function
'*********************************************************		Function to Create Requirement in SE ***********************************************************************
'Function Name		:				Fn_SE_NewRequirementCreate

'Description			 :		 		 This function is used to Create the Requirement in System Engineering

'Parameters			   :	 			1. strNodeName: Select the Requirement 
'													2. strSpecID: ID of the Requirement
'												   3. strSpecRev: Revision of the Requirement
'												  4. strSpecName: Name of the Requirement
'												  5.strSpecDesc: Description of the Requirement

'Return Value		   : 				TRUE \ FALSE

'Pre-requisite			:		 		should be in SE Prespective

'Examples				:				 Call Fn_SE_NewRequirementCreate("Requirement","","","Requirement","Description")

'History:
'										Developer Name			Date			Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'											Amol								24.01.2011																			Tushar
'										Harshal Agrawal				15.02.2011
'									
'										Avinash Jagdale				 27-May-2012					Set the Visual Identifier/Relations 	 																				
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function Fn_SE_NewRequirementCreate(strNodeName,strReqID,strReqRev,strReqName,strReqDesc)
		GBL_FAILED_FUNCTION_NAME="Fn_SE_NewRequirementCreate"
		Dim strNodePath,intNodeCount,sTreeItem,intCount,strNodePathC
		Dim ObjReqWnd,bFlag,iRowIndex
		Fn_SE_NewRequirementCreate=False
		If Fn_UI_ObjectExist("Fn_SE_NewRequirementCreate",JavaWindow("SystemsEngineering").JavaWindow("NewRequirement"))=False Then
			Call Fn_MenuOperation("Select","File:New:Requirement...")
			Call Fn_ReadyStatusSync(1)
		End If
		Set ObjReqWnd=Fn_UI_ObjectCreate("Fn_SE_NewRequirementCreate",JavaWindow("SystemsEngineering").JavaWindow("NewRequirement"))

		Call Fn_UI_JavaTree_Expand("Fn_SE_NewRequirementCreate", ObjReqWnd, "RequirementTree","Complete List")

		JavaWindow("SystemsEngineering").JavaWindow("NewRequirement").JavaTree("RequirementTree").WaitProperty "items count" , micGreaterThan(1)  
		If Fn_UI_JavaTree_NodeExist("Fn_SE_RequirementSpecCreate",ObjReqWnd.JavaTree("RequirementTree"),"Complete List:"+strNodeName) Then
				strNodePathC="Complete List:"+strNodeName
		Else
				strNodePathC="Most Recently Used:"+strNodeName
'				Call Fn_UI_JavaTree_Expand("Fn_SE_NewRequirementCreate", ObjReqWnd, "RequirementTree","Most Recently Used")
		End If
        Call Fn_JavaTree_Select("Fn_SE_NewRequirementCreate", ObjReqWnd, "RequirementTree",strNodePathC)
'		Call Fn_JavaTree_Select("Fn_SE_NewRequirementCreate", ObjReqWnd, "RequirementTree","Complete List")
'		Call Fn_JavaTree_Select("Fn_SE_NewRequirementCreate", ObjReqWnd, "RequirementTree",strNodePathC)
		Call Fn_Button_Click("Fn_SE_NewRequirementCreate",ObjReqWnd,"Next")
		wait(1)
		wait 2
		If strReqID<>"" Then
			'Call Fn_Edit_Box("Fn_SE_NewRequirementCreate",ObjReqWnd,"ID",strReqID)	
		End If
		If strReqRev<>"" Then
			'Call Fn_Edit_Box("Fn_SE_NewRequirementCreate",ObjReqWnd,"Revision",strReqRev)
		End If
		
		Call Fn_Edit_Box("Fn_SE_NewRequirementCreate",ObjReqWnd,"Name",strReqName)
		
		If strReqDesc<>"" Then
			Call Fn_Edit_Box("Fn_SE_NewRequirementCreate",ObjReqWnd,"Description",strReqDesc)
		End If
		Call Fn_Button_Click("Fn_SE_NewRequirementCreate",ObjReqWnd,"Finish")
        'Click on Cancel Button
		If JavaWindow("SystemsEngineering").JavaWindow("NewRequirement").Exist = True Then
            JavaWindow("SystemsEngineering").JavaWindow("NewRequirement").JavaButton("Cancel").Click micLeftBtn
		End If
		Call Fn_ReadyStatusSync(2)
		Call Fn_MenuOperation("Select", "View:Expand Options:Expand")
		Call Fn_ReadyStatusSync(1)
		iRowIndex = DataTable.GetCurrentRow()
		iRowNo = Fn_SE_BOMTable_RowIndex(strReqName)
		DataTable.GetSheet("Global").AddParameter "NewReqID",""
		DataTable.GetSheet("Global").AddParameter "NewReqRev",""
		sReq=Fn_SE_BOMTableNodeOpeations("GetCellData",iRowNo,0,"","")
		DataTable.SetCurrentRow(iRowIndex)
		If sReq <> "" Then
			sReqFull=mid(sReq,instr(1,sReq,":")+1,len(sReq))
			sChdReqName1=mid(sReqFull,instr(1,sReqFull,":")+1,len(sReqFull))
			aItmInfo3 = split(sChdReqName1, "/", -1, 1)
			DataTable("NewReqID",dtGlobalSheet)= aItmInfo3(0)
			aItmInfo4 = split(aItmInfo3(1), ";", -1, 1)
			DataTable("NewReqRev",dtGlobalSheet)= aItmInfo4(0)
		End If
		Fn_SE_NewRequirementCreate=True
		Set ObjReqWnd=Nothing
End Function
'*********************************************************		Function to Create New Paragraph in SE ***********************************************************************
'Function Name		:				Fn_SE_NewParagraphCreate

'Description			 :		 		 This function is used to Create the Paragraph in System Engineering

'Parameters			   :	 			1. strNodeName: Select the Paragraph 
'													2. strSpecID: ID of the Paragraph
'												   3. strSpecRev: Revision of the Paragraph
'												  4. strSpecName: Name of the Paragraph
'												  5.strSpecDesc: Description of the Paragraph

'Return Value		   : 				TRUE \ FALSE

'Pre-requisite			:		 		should be in SE Prespective

'Examples				:				 Call Fn_SE_NewParagraphCreate("Paragraph","223347","A","Paragraph","Description")

'History:
'										Developer Name			Date			Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Amol								24.01.2011																			Tushar
'								Harshal Agrawal					15.01.2011
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function Fn_SE_NewParagraphCreate(strNodeName,strParaID,strParaRev,strParaName,strParaDesc)
		GBL_FAILED_FUNCTION_NAME="Fn_SE_NewParagraphCreate"
		Dim strNodePath,intNodeCount,sTreeItem,intCount,strNodePathC
		Dim ObjParaWnd,bFlag,iRowIndex
		Set ObjParaWnd=JavaWindow("SystemsEngineering").JavaWindow("NewParagraph")
			Fn_SE_NewParagraphCreate=False
		'Verifying "New Paragraph" window's existance
		If Not ObjParaWnd.Exist(8) Then
			'Invoking "New Paragraph" Window
			Call Fn_MenuOperation("Select","File:New:Paragraph...")
		End If
		Call Fn_UI_JavaTree_Expand("Fn_SE_NewParagraphCreate", ObjParaWnd, "ParagraphTree","Complete List")
		JavaWindow("SystemsEngineering").JavaWindow("NewParagraph").JavaTree("ParagraphTree").WaitProperty "items count" , micGreaterThan(1)  
		If Fn_UI_JavaTree_NodeExist("Fn_SE_RequirementSpecCreate",ObjParaWnd.JavaTree("ParagraphTree"),"Complete List:"+strNodeName) Then
				strNodePathC="Complete List:"+strNodeName
		Else
				strNodePathC="Most Recently Used:"+strNodeName
'				Call Fn_JavaTree_Select("Fn_SE_NewParagraphCreate", ObjParaWnd, "ParagraphTree","Most Recently Used")
		End If

        Call Fn_JavaTree_Select("Fn_SE_NewParagraphCreate", ObjParaWnd, "ParagraphTree",strNodePathC)
'		Call Fn_JavaTree_Select("Fn_SE_NewParagraphCreate", ObjParaWnd, "ParagraphTree","Complete List")
'		Call Fn_JavaTree_Select("Fn_SE_NewParagraphCreate", ObjParaWnd, "ParagraphTree",strNodePathC)
		Call Fn_Button_Click("Fn_SE_NewParagraphCreate",ObjParaWnd,"Next")
		wait 2
		If strParaID<>"" Then
			'Setting Id
            wait SISW_MIN_TIMEOUT
			'Call Fn_Edit_Box("Fn_SE_NewParagraphCreate",ObjParaWnd,"ID",strParaID)
		'		JavaWindow("SystemsEngineering").JavaWindow("NewParagraph").JavaEdit("ID").Type strSpecID
		End If
		If strParaRev<>"" Then
			'Setting Revision
			'Call Fn_Edit_Box("Fn_SE_NewParagraphCreate",ObjParaWnd,"Revision",strParaRev)
		'			JavaWindow("SystemsEngineering").JavaWindow("NewParagraph").JavaEdit("Revision").Type strSpecRev
		End If
		'Setting Name
		Call Fn_Edit_Box("Fn_SE_NewParagraphCreate",ObjParaWnd,"Name",strParaName)
		'Setting Description
		Call Fn_Edit_Box("Fn_SE_NewParagraphCreate",ObjParaWnd,"Description",strParaDesc)
		'Clicking On Finish Button To finish the Operation
		Call Fn_Button_Click("Fn_SE_NewParagraphCreate",ObjParaWnd,"Finish")

	    'Click on Cancel Button
		If JavaWindow("SystemsEngineering").JavaWindow("NewParagraph").Exist = True Then
            JavaWindow("SystemsEngineering").JavaWindow("NewParagraph").JavaButton("Cancel").Click micLeftBtn
		End If
		Call Fn_ReadyStatusSync(1)
		Call Fn_MenuOperation("Select", "View:Expand Options:Expand")
		Call Fn_ReadyStatusSync(1)
		iRowIndex = DataTable.GetCurrentRow()
		iRowNo = Fn_SE_BOMTable_RowIndex(strParaName)
		DataTable.GetSheet("Global").AddParameter "NewParaID",""
		DataTable.GetSheet("Global").AddParameter "NewParaRev",""
		sReq=Fn_SE_BOMTableNodeOpeations("GetCellData",iRowNo,0,"","")
		DataTable.SetCurrentRow(iRowIndex)
		If sReq <> "" Then
			sReqFull=mid(sReq,instr(1,sReq,":")+1,len(sReq))
			sChdReqName1=mid(sReqFull,instr(1,sReqFull,":")+1,len(sReqFull))
			aItmInfo3 = split(sChdReqName1, "/", -1, 1)
			DataTable("NewParaID",dtGlobalSheet)= "'"&aItmInfo3(0)
			aItmInfo4 = split(aItmInfo3(1), ";", -1, 1)
			DataTable("NewParaRev",dtGlobalSheet)= aItmInfo4(0)
		End If
		Set ObjReqWnd=Nothing
		'function Return True
		Fn_SE_NewParagraphCreate=True
		'Releasing "New Paragraph" window's object
		Set ObjParaWnd=Nothing
End Function

'*********************************************************	Function  ***********************************************************************
'Function Name		:				Fn_SE_BOMTable_RowIndex

'Description			 :		 		 This function is used to Get the index of that node

'Parameters			   :	 			1. strNodeName: Select the Paragraph 

'Return Value		   : 				Index

'Pre-requisite			:		 		should be in SE Prespective

'Examples				:				 Call Fn_SE_BOMTable_RowIndex("000024/A;1-Spec (View)")

'History:
'										Developer Name			Date			Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Amol								25.01.2011																			Tushar
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Ashwini				20-Mar-2014		25.01.2011				Code modified to identify the QTP version					Ganesh
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function Fn_SE_BOMTable_RowIndex(byval StrNodeName)
	GBL_FAILED_FUNCTION_NAME="Fn_SE_BOMTable_RowIndex"
	On Error Resume Next
	Dim IntRows ,StrNodePath, IntCounter, ObjTable, StrIndex, ArrNode,arrNode1
	Dim aNodePath,sNodeName,iInstance,iGlbCnt,objComponent,StrNodePath1

	Fn_SE_BOMTable_RowIndex="FAIL"

	'If JavaWindow("SystemsEngineering").JavaWindow("WEmbeddedFrame").JavaTable("CMEBOMTreeTable").Exist(iTimeOut) Then  : Swapnil
	If  Window("SEWindow").JavaWindow("WEmbeddedFrame").JavaTable("CMEBOMTreeTable").Exist(iTimeOut) Then
		'Get the No. of rows present in the Bom Table
		'IntRows = JavaWindow("SystemsEngineering").JavaWindow("WEmbeddedFrame").JavaTable("CMEBOMTreeTable").GetROProperty("rows") :Swapnil 
		IntRows = cInt(Window("SEWindow").JavaWindow("WEmbeddedFrame").JavaTable("CMEBOMTreeTable").GetROProperty("rows"))
		'Set ObjTable =JavaWindow("SystemsEngineering").JavaWindow("WEmbeddedFrame").JavaTable("CMEBOMTreeTable").Object :Swapnil
		Set ObjTable =Window("SEWindow").JavaWindow("WEmbeddedFrame").JavaTable("CMEBOMTreeTable").Object
		'Format the Inout as per Table Default Nodes
        'Format the Inout as per Table Default Nodes
				'Get the No. of rows present in the BOM Table

	If instr(StrNodeName, "@") > 0 Then
		aNodePath = split(StrNodeName, "@",-1, 1)
		StrNodeName = aNodePath(0)
		if isNumeric(Trim(aNodePath(1))) = False then 
			Fn_SE_BOMTable_RowIndex = -1
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_SE_BOMTable_RowIndex:Failed to Get  Row Index of [" + sNodeName +"]")	
			exit function
		end if
		iInstance = cint(aNodePath(1))
	Else
		'sNodeName = StrNodeName
		iInstance = 1
	End If

		StrNodeName = Replace(StrNodeName, ":", ", ")
		'Get the Row No. of required Node
        iGlbCnt=0

		'*Commented by Anjali on 24-Dec-2012
'		For IntCounter = 0 to IntRows -1
'			StrNodePath = ObjTable.getPathForRow(IntCounter).toString
'			StrNodePath = Right(StrNodePath, (Len(StrNodePath)-Instr(1, StrNodePath, ",", 1)))					
'			StrNodePath = Left(StrNodePath, Len(StrNodePath)-1)
'	
'			If Trim(StrNodePath) = Trim(StrNodeName) Then
'				StrIndex = Cstr(IntCounter)
'				Fn_SE_BOMTable_RowIndex = StrIndex
'				Exit For
'			End If
'		Next
'*End

'#Added by Anjali on 24-Dec-2012  for TC10.1 Changes

        For iLoop = 1 to iInstance
			For IntCounter = iGlbCnt to IntRows -1
					set objComponent = ObjTable.getComponentForRow(IntCounter)
					StrNodePath = ""
					Do while NOT (objComponent Is Nothing)
						If StrNodePath = "" Then
								StrNodePath = objComponent.getProperty("bl_indented_title")
								If StrNodePath = "" Then
									StrNodePath = objComponent.getProperty("me_cl_display_string")
								End If
						Else
							If objComponent.getProperty("bl_indented_title")  = "" Then
								StrNodePath = objComponent.getProperty("me_cl_display_string") & ", " & StrNodePath
							Else
								StrNodePath =objComponent.getProperty("bl_indented_title") & ", " & StrNodePath
							End If
						End If
'------------------------------Code modified to identify the QTP version----------------------						
							If Environment.Value("ProductName") = sQTPProductName OR Environment.Value("ProductName") = sUFTProductName Then
								If IsObject(objComponent.parent()) = True Then
									set objComponent = objComponent.parent()
								Else
									Exit Do
								End If
							Else
								set objComponent = objComponent.parent()
								If objComponent Is Nothing Then
									Exit do
								End If
							End IF
							
						Loop
						set objComponent = Nothing
						'If Trim(StrNodePath) = Trim(StrNodeName) Then
							If Instr(1,Trim(StrNodePath),Trim(StrNodeName)) > 0 Then
								arrNode1 = Split(StrNodePath,"-",-1,1)
								If len(arrNode1(uBound(arrNode1))) <= 1 or IsNumeric(arrNode1(uBound(arrNode1))) Then
								arrNode1(uBound(arrNode1)) = arrNode1(uBound(arrNode1)-1) + "-" +arrNode1(uBound(arrNode1))
							End If
								StrNodePath1 = arrNode1(uBound(arrNode1))
							If Trim(StrNodePath1) = Trim(StrNodeName) Then
								StrIndex = IntCounter
								iGlbCnt = StrIndex +1 
								Fn_SE_BOMTable_RowIndex = StrIndex
								Exit For
							End If
						End If
			Next
		 Next
'#End
		If IntCounter = IntRows Then
			Fn_SE_BOMTable_RowIndex = "FAIL:Node Not Found"
		End If
		'Release the Table object
	   set ObjTable = Nothing
	Else
        Fn_SE_BOMTable_RowIndex="FAIL"
	End If
End Function


'*********************************************************	Function  ***********************************************************************
'Function Name		:				Fn_SE_BOMTableNodeOpeations

'Description			 :		 		 This function is used to Get the index of that node

'Parameters			   :	 			1. strAction: Select the Paragraph 
'													3. strNodeName: Node to select the tree
'														3. strColName: Column name of the table
'														4. strColValue: Value of the column
'															5. strPopupMenu: Popup Menu

'Return Value		   : 				True False

'Pre-requisite			:		 		should be in SE Prespective

'Examples				:				 Call Fn_SE_BOMTableNodeOpeations("Select","000024/A;1-Spec (View)","","","")
'								Case "EditNumber" : Call Fn_SE_BOMTableNodeOpeations("EditNumber","000335/A;1-Requierment_Data_File (View):REQ-000071/A;1-Program Context (View)","Number", "3", "")
'								Case "ColumnExists" : Call Fn_SE_BOMTableNodeOpeations("ColumnExists","","Item Type", "", "")
'								Case "AddColumns" : Call Fn_SE_BOMTableNodeOpeations("AddColumns","","Item Type~Item Description", "", "")
'								Case "GetRowCount" : Call Fn_SE_BOMTableNodeOpeations("GetRowCount","","", "", "")
'								Case "ExpandBelow" : Call Fn_SE_BOMTableNodeOpeations("ExpandBelow"," 000114/A;1-Spec (View):REQ-000001/A;1-Req (View)","","","")
'								Case "AllColumnNames" : Call Fn_SE_BOMTableNodeOpeations("AllColumnNames","","", "", "")
'								Case " ConfigureColumn"	 : bReturn=Fn_SE_BOMTableNodeOpeations("ConfigureColumn","","APN UID:All Notes", "", "")
'								bReturn=Fn_SE_BOMTableNodeOpeations("GetNodePathByName","ParaChild1234","", "", "")
'								Case "Collapse" : Call Fn_SE_BOMTableNodeOpeations("Collapse","000024/A;1-Spec (View)","", "", "")
'History:
'		Developer Name		Date		                     Rev. No.	Changes Done													Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'		Amol			25.01.2011																							Tushar
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'		Koustubh		25.05.2011				       Added case : AddColumns, ColumnExists										Tushar
'		Ketan Raje		16.09.2011				       Added Caes : IsSelected
'		Sandeep			27.09.2011				       Added Case : GetRowCount
'		Sandeep			14.10.2011				       Added Case : ExpandBelow
'		Amit T			19.10.2011				       Added Case : SelectAll
'		Ketan Raje		20.10.2011				       Added Case : PopupMenuEnabled
'		Sandeep N		21.10.2011				       Added Case : AllColumnNames													
'		Sandeep N		16.11.2011				       Added Case : RemoveColumn
'		Shreyas			10-01-2012				       Added Case : ConfigureColumn
'		Sandeep N		16.11.2011				       Added Case : GetNodePathByName
'		Sachin J.		        30 - 05 - 2012			       Added Case : Collapse
'		Snehal			20-May-2015				       Added case : SelectExpRow
'		Jotiba			24-Feb-2016				       Added Case : "SelectRowsRange" and "Remove" from TC1015 to mainline
'		Shraddha J		26-Apr-2016				       Added Case : "AddColumnsWithoutClose" from TC1015							[TC1122-20160413-26_04_2016-VivekA-Maintenance]
'		Chaitali R		05-June-2017				       Added Case : "MoveColumnUp" & "MoveColumnDown" 							[TC113-20170509-05_06_2017-ShwetaR-Development]
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function Fn_SE_BOMTableNodeOpeations(strAction,strNodeName,strColName, strColValue, strPopupMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_SE_BOMTableNodeOpeations"
	on Error Resume Next
	Dim iRowNo, sMenu, iNodeNo, iColNo, iStart,strName, objContextMenu
	Dim objTable, iCnt, aColumns, objChangeCol, sColName
	Dim bFlag,sValues,icount1
	Dim sColValue,aValues,iCounter,i,stabname
	Dim iRows,iPathCount,StrNodePath,iCount,aName,iOccurence,strNodeName1
	Dim objSelectType, objIntNoOfObjects, objItem
	Dim objTabFld,StrTabName,objComponent
    Dim ObjAppletWin,objLov,objChild,iCnt1
    Dim arrNode,arrNode1,strNodeName2
         
	If strColValue="RequirementSpec" Then
		strColValue="Requirement Specification"
	End If
	bFlag=False
	Fn_SE_BOMTableNodeOpeations=False
	'Set objTable = JavaWindow("SystemsEngineering").JavaWindow("WEmbeddedFrame").JavaTable("CMEBOMTreeTable") swapnil :Parent class of JavaApplet changed.
	Set objTable = Window("SEWindow").JavaWindow("WEmbeddedFrame").JavaTable("CMEBOMTreeTable") 
	'Verify ReqMgr Table
'	If objTable.Exist(iTimeOut) then
'	   iRowNo = objTable.GetROProperty("rows")   [ Commented this code since it expanded the Top node, thereby causing the VP to Fail ]
'		 If iRowNo = 1 Then
'			 call Fn_menuOperation("Select","View:Expand Options:Expand")
'			 Wait 1
'		 End If

		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Code to handle Applet as applet index is dyanamically changes
'		Set objSelectType = description.Create()
'		objSelectType("Class Name").value = "JavaObject"
'		objSelectType("toolkit class").value = "com.teamcenter.rac.presentation.RACTabFolderWidget"						
'		Set  objIntNoOfObjects = JavaWindow("SystemsEngineering").ChildObjects(objSelectType)
'		
'		bFlag = False
'	
'			For iCount = 0 to (objIntNoOfObjects.Count -1)
'				i = objIntNoOfObjects(iCount).Object.getSelectedTabIndex
'				Set objItem = objIntNoOfObjects(iCount).Object.getItem(i)
'				StrTabName=objItem.text()
'				StrTabName=Split(StrTabName,"-")
'				If StrTabName(0)="REQ" Then
'					StrTabName(0)=StrTabName(1)
'				End If
'				For iCounter=0 to 8
'					Window("SEWindow").JavaWindow("WEmbeddedFrame").SetTOProperty "index",iCounter
'				    If Window("SEWindow").JavaWindow("WEmbeddedFrame").JavaTable("CMEBOMTreeTable").Exist(1) Then
'					   If InStr(1,trim(Window("SEWindow").JavaWindow("WEmbeddedFrame").JavaTable("CMEBOMTreeTable").Object.getValueAt(0,0).toString()),trim(StrTabName(0)))>0 Then 
''							If InStr(1,trim(strNodeName), trim(Window("SEWindow").JavaWindow("WEmbeddedFrame").JavaTable("CMEBOMTreeTable").Object.getValueAt(0,0).toString()))>0 Then  '' added to to compare first row value of table with strNodeName
'								bFlag = True
'								Exit For
''							End If
'						End If
'					End If
'				Next
'				If bFlag Then
'					Exit For
'				End If
'			Next
'			
'		Set objTabFld = JavaWindow("SystemsEngineering").JavaObject("RACTabFolderWidget")
'		i = objTabFld.Object.getSelectedTabIndex
'		StrTabName=objTabFld.Object.getItem(i).text()
'		StrTabName=Split(StrTabName,"-")
''		- - - - - - - Added Code to handle Requirement Opened in BOM table
'		If StrTabName(0)="REQ" Then
'			StrTabName(0)=StrTabName(1)
'		End If
'      For iCounter=0 to 12
'			 Window("SEWindow").JavaWindow("WEmbeddedFrame").SetTOProperty "index",iCounter
'			 If Window("SEWindow").JavaWindow("WEmbeddedFrame").JavaTable("CMEBOMTreeTable").Exist(2) Then
'				'If InStr(1,trim(Window("SEWindow").JavaWindow("WEmbeddedFrame").JavaTable("CMEBOMTreeTable").Object.getValueAt(0,0).toString()),trim(StrTabName(0))) Then
'				If InStr(1,trim(Window("SEWindow").JavaWindow("WEmbeddedFrame").JavaTable("CMEBOMTreeTable").Object.getComponentForRow(0).getProperty("bl_indented_title")),trim(StrTabName(0))) Then
'					bFlag=True
'					Exit for
'				End If
'			 End If
'		Next
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'[Mainline -(20170726.00)-10_8_2017-JotibaT]   Added Code to handle Requirement Opened in BOM table
		Set objSelectType = description.Create()
		objSelectType("Class Name").value = "JavaTab"
		objSelectType("toolkit class").value = "org.eclipse.swt.custom.CTabFolder"						
		Set  objIntNoOfObjects = JavaWindow("DefaultWindow").ChildObjects(objSelectType)
		
		For icount1 = 0 To objIntNoOfObjects.Count-1 Step 1	
			iIndex=objIntNoOfObjects(icount1).Object.getSelectionIndex
			Set objItem=objIntNoOfObjects(icount1).Object.getItem(iIndex)
			StrTabName=trim(objItem.text)
			StrTabName=Split(StrTabName,"-")
			
			If StrTabName(0)="REQ" Then
				StrTabName(0)=StrTabName(1)
			End If
			
			For iCounter=0 to 12
				Window("SEWindow").JavaWindow("WEmbeddedFrame").SetTOProperty "index",iCounter
				If Window("SEWindow").JavaWindow("WEmbeddedFrame").JavaTable("CMEBOMTreeTable").Exist(2) Then
				stabname = trim(Window("SEWindow").JavaWindow("WEmbeddedFrame").JavaTable("CMEBOMTreeTable").Object.getComponentForRow(0).getProperty("bl_indented_title"))
					If stabname<>"" Then
							If InStr(1,stabname,trim(StrTabName(0))) Then
								bFlag=True
								Exit for
							End If
					End If
				End If
			Next
			If bFlag=True Then
				Exit for
			End If
		Next	
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 	 
		If bFlag=false Then
			Exit function
		End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 

		Select Case StrAction

			Case "Select"		'("Select"," 000452/A;1-ReqSpec (View):REQ-000047/A;1-Req2 (View):000454/A;1-P2 (View):REQ-000048/A;1-Req3","","","")
				arrNode = Split(strNodeName,"-",-1,1)
				If len(arrNode(uBound(arrNode))) <= 1 or IsNumeric(arrNode(uBound(arrNode))) Then
					arrNode(uBound(arrNode)) = arrNode(uBound(arrNode)-1) + "-" +arrNode(uBound(arrNode))
				End If
				strNodeName1 = arrNode(uBound(arrNode))
				iRowNo = Fn_SE_BOMTable_RowIndex(strNodeName1)
				If isNumeric(iRowNo) Then
					objTable.SelectRow iRowNo
					Fn_SE_BOMTableNodeOpeations=True
				End if
				If JavaWindow("SystemsEngineering").JavaWindow("ErrorWindow").Exist(2) Then
					Call Fn_SISW_UI_JavaButton_Operations("Fn_SISW_SE_TraceLinkTabOperations", "Click", JavaWindow("SystemsEngineering").JavaWindow("ErrorWindow"),"OK")
				End If	
			Case "SelectExpRow"		'("SelectExpRow"," 000452/A;1-ReqSpec (View):REQ-000047/A;1-Req2 (View):000454/A;1-P2 (View):REQ-000048/A;1-Req3","","","")
				arrNode = Split(strNodeName,"-",-1,1)
				If len(arrNode(uBound(arrNode))) <= 1 or IsNumeric(arrNode(uBound(arrNode))) Then
					arrNode(uBound(arrNode)) = arrNode(uBound(arrNode)-1) + "-" +arrNode(uBound(arrNode))
				End If
				strNodeName1 = arrNode(uBound(arrNode))
				iRowNo = Fn_SE_BOMTable_RowIndex(strNodeName1)
				objTable.Object.getColumnModel().getColumn(0).setPreferredWidth("180")
				If isNumeric(iRowNo) Then
					objTable.SelectRow iRowNo
					Fn_SE_BOMTableNodeOpeations=True
				End if
				
			Case "IsSelected"		'("IsSelected"," 000040/A;1-Spec1 (View):REQ-000004/A;1-Req2","","","")
				arrNode = Split(strNodeName,"-",-1,1)
				If len(arrNode(uBound(arrNode))) <= 1 or IsNumeric(arrNode(uBound(arrNode))) Then
					arrNode(uBound(arrNode)) = arrNode(uBound(arrNode)-1) + "-" +arrNode(uBound(arrNode))
				End If
				strNodeName1 = arrNode(uBound(arrNode))
				iRowNo = Fn_SE_BOMTable_RowIndex(strNodeName1)				
				If isNumeric(iRowNo) Then
					If Cint(objTable.GetROProperty("SelectedRow")) = Cint(iRowNo) Then						
						Fn_SE_BOMTableNodeOpeations=True
					End If
				End if

			Case "Deselect"		'("Select"," 000452/A;1-ReqSpec (View):REQ-000047/A;1-Req2 (View):000454/A;1-P2 (View):REQ-000048/A;1-Req3","","","")
				arrNode = Split(strNodeName,"-",-1,1)
				If len(arrNode(uBound(arrNode))) <= 1 or IsNumeric(arrNode(uBound(arrNode))) Then
					arrNode(uBound(arrNode)) = arrNode(uBound(arrNode)-1) + "-" +arrNode(uBound(arrNode))
				End If
				strNodeName1 = arrNode(uBound(arrNode))			
				iRowNo = Fn_SE_BOMTable_RowIndex(strNodeName1)
				If isNumeric(iRowNo) Then
					objTable.DeselectRow iRowNo
					Fn_SE_BOMTableNodeOpeations=True
				End if

			Case "VerifyNode"		'("Verify"," 000452/A;1-ReqSpec (View):REQ-000047/A;1-Req2 (View):000454/A;1-P2 (View)","","","")		
        			'Verify Node Exist
        			arrNode = Split(strNodeName,"-",-1,1)
'        				if instr(arrNode(uBound(arrNode)),"@") then
'							  
'        				End If
        			If len(arrNode(uBound(arrNode))) <= 1 or IsNumeric(arrNode(uBound(arrNode)))Then
					arrNode(uBound(arrNode)) = arrNode(uBound(arrNode)-1) + "-" +arrNode(uBound(arrNode))
					End If
					strNodeName1 = arrNode(uBound(arrNode))
					iRowNo = Fn_SE_BOMTable_RowIndex(strNodeName1)
					If isNumeric(iRowNo) then
						Fn_SE_BOMTableNodeOpeations=True
					End if
        
			Case "getNodeIndex"	'("getNodeIndex"," 000452/A;1-ReqSpec (View):REQ-000047/A;1-Req2 (View):000454/A;1-P2 (View)","","","")
				arrNode = Split(strNodeName,"-",-1,1)
				If len(arrNode(uBound(arrNode))) <= 1 or IsNumeric(arrNode(uBound(arrNode))) Then
					arrNode(uBound(arrNode)) = arrNode(uBound(arrNode)-1) + "-" +arrNode(uBound(arrNode))
				End If
				strNodeName1 = arrNode(uBound(arrNode))				
				iRowNo = Fn_SE_BOMTable_RowIndex(strNodeName1)
				If isNumeric(iRowNo) then
					Fn_SE_BOMTableNodeOpeations=iRowNo
				End if

			Case "Expand"	'("Expand"," 000452/A;1-ReqSpec (View):REQ-000047/A;1-Req2 (View)","","","")
				arrNode = Split(strNodeName,"-",-1,1)
				If len(arrNode(uBound(arrNode))) <= 1 or IsNumeric(arrNode(uBound(arrNode))) Then
					arrNode(uBound(arrNode)) = arrNode(uBound(arrNode)-1) + "-" +arrNode(uBound(arrNode))
				End If
				strNodeName1 = arrNode(uBound(arrNode))
				iRowNo = Fn_SE_BOMTable_RowIndex(strNodeName1)
				If isNumeric(iRowNo) then
					objTable.SelectRow iRowNo
					'Code modified By Ketan on 06/09/2011 as Expand call is not working.				
					call Fn_menuOperation("Select","View:Expand Options:Expand Below...")
'					Call Fn_Edit_Box("Fn_SE_BOMTableNodeOpeations",JavaWindow("SystemsEngineering").JavaWindow("WEmbeddedFrame").JavaDialog("ExpandToLevel"),"Level","1")
			        JavaWindow("SystemsEngineering").JavaWindow("ExpandToLevel").JavaSpin("Spinner").Set "1"
					  Call Fn_Button_Click("Fn_SE_BOMTableNodeOpeations", JavaWindow("SystemsEngineering").JavaWindow("ExpandToLevel"), "OK")
					Fn_SE_BOMTableNodeOpeations=True
				End if

			Case "Collapse"		'("Collapse"," 000040/A;1-Spec1 (View)","","","")
				arrNode = Split(strNodeName,"-",-1,1)
				If len(arrNode(uBound(arrNode))) <= 1 or IsNumeric(arrNode(uBound(arrNode))) Then
					arrNode(uBound(arrNode)) = arrNode(uBound(arrNode)-1) + "-" +arrNode(uBound(arrNode))
				End If
				strNodeName1 = arrNode(uBound(arrNode))
				iRowNo = Fn_SE_BOMTable_RowIndex(strNodeName1)
				If isNumeric(iRowNo) Then
					objTable.SelectRow iRowNo
					Fn_SE_BOMTableNodeOpeations=True
					If Err.Number < 0 Then						
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail: Function [Fn_SE_BOMTableNodeOpeations] Failed to Select SE BOM Table Node [" + StrNodeName + "]")
						Fn_SE_BOMTableNodeOpeations = FALSE			
					Else
						'Operate Collapse Below Menu if Node selected Sucessfully
						StrReturn = Fn_MenuOperation("Select", "View:Collapse Below")
						If StrReturn = True Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Pass: Function [Fn_SE_BOMTableNodeOpeations] Sucessfully CollaSEd SE BOM Table Node [" + StrNodeName + "]")							
							Fn_SE_BOMTableNodeOpeations = TRUE
						Else							
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail: Function [Fn_SE_BOMTableNodeOpeations] Failed to CollaSE SE BOM Table Node [" + StrNodeName + "]")
							Fn_SE_BOMTableNodeOpeations = FALSE
						End If						
					End If
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail: Function [Fn_SE_BOMTableNodeOpeations] Failed to Get SE BOM Table Node [" + StrNodeName + "]")
					Fn_SE_BOMTableNodeOpeations = FALSE
				End If
				
			Case "ClickCell"		'("ClickCell"," 000452/A;1-ReqSpec (View):REQ-000047/A;1-Req2 (View):000454/A;1-P2 (View):REQ-000048/A;1-Req3","","","")
				arrNode = Split(strNodeName,"-",-1,1)
				If len(arrNode(uBound(arrNode))) <= 1 or IsNumeric(arrNode(uBound(arrNode))) Then
					arrNode(uBound(arrNode)) = arrNode(uBound(arrNode)-1) + "-" +arrNode(uBound(arrNode))
				End If
				strNodeName1 = arrNode(uBound(arrNode))
					iRowNo = Fn_SE_BOMTable_RowIndex(strNodeName1)
					iColNo = Fn_SISW_UI_JavaTable_Operations("","GetColumnIndex",objTable, "", "", strColName, "", "", "", "", "")
					If isNumeric(iRowNo) AND iColNo <> -1 Then
						objTable.ClickCell iRowNo,iColNo
						If Err.number < 0 Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail: Function [Fn_SE_BOMTableNodeOpeations] Failed to Click on TableCell.")
							Fn_SE_BOMTableNodeOpeations = False
						Else
							Fn_SE_BOMTableNodeOpeations = True
						End If
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail: Function [Fn_SE_BOMTableNodeOpeations] Failed to get Row and Column index of Node.")
						Fn_SE_BOMTableNodeOpeations=False
					End if

			Case "PopupMenuSelect"	'("PopupMenuSelect","","","","Trace Link:Start Trace Link")
				'Pre-requisite = Row should be selected
				strPopupMenu=Replace(strPopupMenu,":",";")
				iRowNo = objTable.Object.getSelectedRow()
				If isNumeric(iRowNo) then
					'JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet").JavaTable("RMTable").SelectRow iRowNo
					objTable.ClickCell iRowNo,0,"RIGHT" 
					wait 1
					sMenu = JavaWindow("SystemsEngineering").WinMenu("ContextMenu").BuildMenuPath(strPopupMenu)
					wait 3
					JavaWindow("SystemsEngineering").WinMenu("ContextMenu").Select sMenu
					
					Fn_SE_BOMTableNodeOpeations=True
				End if

			Case "MultiSelect"		'("MultiSelect","REQ-000049/A;1-Req1 (View):000455/A;1-P1 (View):000575/A;1-P23~REQ-000049/A;1-Req1 (View):REQ-000148/A;1-Req2","","","")

				strNodeName=split(strNodeName,"~")
				For iNodeNo=0 to Ubound(strNodeName)
					arrNode1 = Split(strNodeName(iNodeNo),":",-1,1)
					arrNode = Split(arrNode1(uBound(arrNode1)),"-",-1,1)
					'strNodeName(iNodeNo) = strNodeName+":"+arrNode(uBound(arrNode))
					iRowNo = Fn_SE_BOMTable_RowIndex(arrNode(uBound(arrNode)))
					If isNumeric(iRowNo) Then
						If iNodeNo=0 Then
							objTable.SelectRow iRowNo
							Fn_SE_BOMTableNodeOpeations=True
						Else
							objTable.ExtendRow "#"&iRowNo
							Wait 2
							Fn_SE_BOMTableNodeOpeations=True
						End If					
					End if
				Next
				
			Case "VerifyColValue"	'("VerifyColValue"," REQ-000049/A;1-Req1 (View):000455/A;1-P1 (View)","Item Type","Requirement","")
				arrNode = Split(strNodeName,"-",-1,1)
				If len(arrNode(uBound(arrNode))) <= 1 or IsNumeric(arrNode(uBound(arrNode))) Then
					arrNode(uBound(arrNode)) = arrNode(uBound(arrNode)-1) + "-" +arrNode(uBound(arrNode))
				End If
				strNodeName1 = arrNode(uBound(arrNode))
				iRowNo = Fn_SE_BOMTable_RowIndex(strNodeName1)
            	If isNumeric(iRowNo) then
					'Get column Rows
					iColNo = objTable.GetROProperty("cols")

					For iStart=0 to iColNo-1
						''Verify the Column name is similar to required column name
						If objTable.GetColumnName(iStart)=strColName Then
							If instr(strColValue,",") Then
								aVal = Split(strColValue,",")
								For iCnt = 0 To Ubound(aVal) 
									Fn_SE_BOMTableNodeOpeations = False
									If instr(trim(objTable.GetCellData(iRowNo,iStart)),trim(cstr(aVal(iCnt)))) then
										Fn_SE_BOMTableNodeOpeations=True
									Else
										Exit For
									End if	
								Next
							Else
								'Verify the Column value is similar to required value
								If cstr(objTable.GetCellData(iRowNo,iStart))=cstr(strColValue) then
									Fn_SE_BOMTableNodeOpeations=True
								End if							
							End If
							Exit For
						End If
					Next
				Else
					Fn_SE_BOMTableNodeOpeations=False
				End if

			Case "EditColValue"		'("EditColValue"," REQ-000049/A;1-Req1 (View):000455/A;1-P1 (View)","Find No.1","50","")	
				arrNode = Split(strNodeName,"-",-1,1)
				If len(arrNode(uBound(arrNode))) <= 1 or IsNumeric(arrNode(uBound(arrNode))) Then
					arrNode(uBound(arrNode)) = arrNode(uBound(arrNode)-1) + "-" +arrNode(uBound(arrNode))
				End If
				strNodeName1 = arrNode(uBound(arrNode))
				iRowNo = Fn_SE_BOMTable_RowIndex(strNodeName1)
            	If isNumeric(iRowNo) then
					
					'Get column Rows
					iColNo=objTable.GetROProperty("cols")
					For iStart=0 to iColNo-1
						''Verify the Column name is similar to required column name
						If objTable.GetColumnName(iStart)=strColName Then
								'Verify the Column value is similar to required value
								If strColName = "IP Classification" Then
	'								objTable.SetCellData iRowNo,iStart,strColValue
									objTable.ClickCell iRowNo,strColName
									wait 1
									Window("SEWindow").JavaWindow("WEmbeddedFrame").JavaEdit("LOVSelectionDisplayView").Set strColValue+vblf 
								
								ElseIf strColName = "Content Type" Then   '[TC11.3(20160131_NewDevelopment_PoonamC_31Mar2017):Added case to edit colvalue for Content Type Lov]
								     Set ObjAppletWin = Window("SEWindow").JavaWindow("WEmbeddedFrame")
									 objTable.ClickCell iRowNo,strColName
									 wait 2
									 If ObjAppletWin.Exist(2) = False Then
									  ObjAppletWin.SetTOProperty "index",2
									 End if
									 ObjAppletWin.JavaButton("dropdown_16").Click
									  wait 2
									  'get LOV Table object
									  Set objLov=Description.Create()
										  objLov("Class Name").value="JavaTable"
										  'objLov("tagname").value="LOVTreeTable"
										
									 Set objChild = ObjAppletWin.ChildObjects(objLov)
								     Wait 1
									'Traverse through the value list 
									For iCnt = 0 To objChild.Count 	
											bFlag = False								
	                  						If trim(objChild(1).Object.getValueAt(iCnt,0).getDisplayableValue())=trim(strColValue) Then
												objChild(1).ClickCell iCnt,0
												wait 3
												bFlag=true
												Exit For
											End If
									Next
									Set objChild = Nothing
									Set ObjAppletWin = Nothing
									Set objLov = Nothing
									
									 If bFlag=true Then
									 	Fn_SE_BOMTableNodeOpeations=True
									 Else
									 	Fn_SE_BOMTableNodeOpeations=False
									 	Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [Fn_SE_BOMTableNodeOpeations] Failed to select [ "&strColValue&" ]")
									 	Exit Function
									 End If 
							Else
									objTable.SetCellData iRowNo,iStart,strColValue
							End If
							Fn_SE_BOMTableNodeOpeations=True
							Exit For
						End If
					Next
				Else
					Fn_SE_BOMTableNodeOpeations=False
				End if
			Case "GetCellData" '("GetCellData",1,0,"","")
					''modify case GetCellData:- to use "bl_indented_title" method to get cell data if column name is Function Or Requirement.
					'(IMP NOTE)' "strNodeName" - This parameter is use as Row number in this Case
					'strColName - This parameter is use as column number in this case
					Select Case strColName
						Case "Requirement", "Function"
							strColName = 0
					End Select
					If strcomp(strColName,0) = 0 Then
						set ObjTable = Window("SEWindow").JavaWindow("WEmbeddedFrame").JavaTable("CMEBOMTreeTable").Object
							set objComponent = ObjTable.getComponentForRow(strNodeName)
							strName = ""
							Do while NOT (objComponent Is Nothing)
								If strName = "" Then
										strName = objComponent.getProperty("bl_indented_title")
										If strName = "" Then
											strName = objComponent.getProperty("me_cl_display_string")
										End If
								Else
									If objComponent.getProperty("bl_indented_title")  = "" Then
										strName = objComponent.getProperty("me_cl_display_string") & ", " & strName
									Else
										strName =objComponent.getProperty("bl_indented_title") & ", " & strName
									End If
								End If
								If Environment.Value("ProductName") = sQTPProductName OR Environment.Value("ProductName") = sUFTProductName Then
									If IsObject(objComponent.parent()) = True Then
										set objComponent = objComponent.parent()
									Else
										Exit Do
									End If
								Else
									set objComponent = objComponent.parent()
									If objComponent Is Nothing Then
										Exit do
									End If
								End IF
									'set objComponent = objComponent.parent()
									If objComponent Is Nothing Then
										Exit do
									End If
								Loop
								set objComponent = Nothing
									aName = Split(strName, ", ")
									strName = Join(aName, ":")
					Else
						objTable.SelectRow strNodeName
						wait(3)
						strName = objTable.GetCellData(strNodeName,strColName)
						strName = MId(strName,instr(1,strName,":")+1 , Len(strName))
					End If

					If Err.number < 0 Then
						Fn_SE_BOMTableNodeOpeations=False
					Else
						Fn_SE_BOMTableNodeOpeations=strName
					End If

			Case "PopupMenuExist"		
						strPopupMenu=Replace(strPopupMenu,":",";")
						iRowNo = objTable.Object.getSelectedRow()
						If isNumeric(iRowNo) then
							objTable.ClickCell iRowNo,0,"RIGHT" 
							wait 1
							sMenu = JavaWindow("SystemsEngineering").WinMenu("ContextMenu").BuildMenuPath(strPopupMenu)
							If JavaWindow("SystemsEngineering").WinMenu("ContextMenu").GetItemProperty (sMenu,"Exists") = True Then
									Fn_SE_BOMTableNodeOpeations = TRUE
							Else
									Fn_SE_BOMTableNodeOpeations = FALSE
							End If
						End If
		Case "DoubleClickCell"		'("DoubleClickCell"," 000452/A;1-ReqSpec (View):REQ-000047/A;1-Req2 (View):000454/A;1-P2 (View):REQ-000048/A;1-Req3","Has Attached Notes","","")
				arrNode = Split(strNodeName,"-",-1,1)
				If len(arrNode(uBound(arrNode))) <= 1 or IsNumeric(arrNode(uBound(arrNode))) Then
					arrNode(uBound(arrNode)) = arrNode(uBound(arrNode)-1) + "-" +arrNode(uBound(arrNode))
				End If
				strNodeName1 = arrNode(uBound(arrNode))				
				iRowNo = Fn_SE_BOMTable_RowIndex(strNodeName1)
				If isNumeric(iRowNo) Then
					objTable.SelectRow iRowNo
					Fn_SE_BOMTableNodeOpeations=True
				End if
				objTable.DoubleClickCell iRowNo,strColName
		Case "EditNumber"
				arrNode = Split(strNodeName,"-",-1,1)
				If len(arrNode(uBound(arrNode))) <= 1 or IsNumeric(arrNode(uBound(arrNode))) Then
					arrNode(uBound(arrNode)) = arrNode(uBound(arrNode)-1) + "-" +arrNode(uBound(arrNode))
				End If
				strNodeName1 = arrNode(uBound(arrNode))				
				iRowNo = Fn_SE_BOMTable_RowIndex(strNodeName1)
				If IsNumeric(iRowNo) Then
					objTable.SelectRow iRowNo					
					'Open the Edit Number Dialog.
					objTable.DoubleClickCell iRowNo,strColName
					wait 2
					If Not JavaWindow("SystemsEngineering").JavaWindow("EditNumber").Exist(5) Then
						Call Fn_MenuOperation("Select","Edit:Edit Number...")
						wait 1
					End If
					'Set the New number value
					Call Fn_Edit_Box("Fn_SE_BOMTableNodeOpeations",JavaWindow("SystemsEngineering").JavaWindow("EditNumber"),"NewNumber","")	
					Call Fn_UI_EditBox_Type("Fn_SE_BOMTableNodeOpeations",JavaWindow("SystemsEngineering").JavaWindow("EditNumber"),"NewNumber",strColValue)
					'Click on Ok button
					Call Fn_Button_Click("Fn_SE_BOMTableNodeOpeations", JavaWindow("SystemsEngineering").JavaWindow("EditNumber"), "OK")
					Fn_SE_BOMTableNodeOpeations=True
				End If
		Case "ColumnExists"
					iColNo = cInt(objTable.GetROProperty("cols"))
					For iCnt = 0 to iColNo -1
						If trim(objTable.GetColumnName(iCnt)) = strColName then
							Fn_SE_BOMTableNodeOpeations = True
							Exit for
						end if
					Next
		Case "AddColumn", "AddColumns", "AddColumnsWithoutClose", "AddColumnsWithSaveOnly","AddColumnExt"
					' checking existance of columns 
					iColNo = cInt(objTable.GetROProperty("cols"))
					aColumns = split(strColName,"~")
					For iColumnCnt = 0 to UBound(aColumns)
						For iCnt = 0 to iColNo -1
							' if exists then ignoring that column
							If trim(objTable.GetColumnName(iCnt)) = aColumns(iColumnCnt) then
								aColumns(iColumnCnt) = ""
								Exit for
							end if
						Next
					Next
					
					'[TC_11.4_NewDevelopment_PoonamC_08Aug2017 : Added New case to add column]
					If StrAction = "AddColumnExt" Then
							objTable.SelectColumnHeader 0
							Wait 3
							objTable.SelectColumnHeader 0,"RIGHT"
							wait 3
					Else
							'' Change parameter from 1 to 0 by dipali
							If  iColNo > 1 Then
								'objTable.SelectColumnHeader 1,"LEFT"
								'wait 3
								objTable.SelectColumnHeader 1,"RIGHT"
								wait 2
							Else
								'objTable.SelectColumnHeader 0,"LEFT"
								'wait 3
								objTable.SelectColumnHeader 0,"RIGHT"
								wait 2
							End If
					End If	
					Window("SEWindow").JavaWindow("WEmbeddedFrame").JavaMenu("label:=Insert column\(s\) \.\.\.").Select

					If Window("SEWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Change Columns").Exist(5) Then
						Set objChangeCol = Window("SEWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Change Columns")
						For iColumnCnt = 0 to UBound(aColumns)
							If aColumns(iColumnCnt) <> "" Then
								Call Fn_List_Select("Fn_SE_BOMTableNodeOpeations", objChangeCol,"ListAvailableCols",aColumns(iColumnCnt))
								Wait 2
								Call Fn_Button_Click("Fn_SE_BOMTableNodeOpeations", objChangeCol,"Add")
								wait 1
							End If
						Next
						
						If strAction <> "AddColumnsWithoutClose"  Then
							If strAction <> "AddColumnsWithSaveOnly" Then
								If cInt(objChangeCol.JavaButton("Apply").getROProperty("enabled")) = 1 then
									Call Fn_Button_Click("Fn_SE_BOMTableNodeOpeations", objChangeCol,"Apply")
								End if
							Else
								If CInt(objChangeCol.JavaButton("Save").getROProperty("enabled")) = 1 Then
									Call Fn_Button_Click("Fn_SE_BOMTableNodeOpeations", objChangeCol,"Save")
								End If
								Set ObjSaveConfig = Window("SEWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Save Column Configuration")
								If ObjSaveConfig.Exist(2) Then
									strColValue = Split(strColValue,"~")
									'Set Config Name
									Call Fn_Edit_Box("Fn_SE_BOMTableNodeOpeations",ObjSaveConfig,"Name",strColValue(0)) 
									wait 1
									'Set Config Desc
									Call Fn_Edit_Box("Fn_SE_BOMTableNodeOpeations",ObjSaveConfig,"Description",strColValue(1)) 
									wait 1
									Call Fn_Button_Click("Fn_SE_BOMTableNodeOpeations", ObjSaveConfig,"Save")
								End If
							End If

							Call Fn_Button_Click("Fn_SE_BOMTableNodeOpeations", objChangeCol,"Cancel")
						End If
						Fn_SE_BOMTableNodeOpeations = True
					End if
			' GetRowCount - Case will return total number of row count in Table
			Case "GetRowCount"
					Fn_SE_BOMTableNodeOpeations=Fn_UI_Object_GetROProperty("Fn_SE_BOMTableNodeOpeations",objTable,"rows")

			Case "ExpandBelow"	'("ExpandBelow"," 000452/A;1-ReqSpec (View):REQ-000047/A;1-Req2 (View)","","","")
				arrNode = Split(strNodeName,"-",-1,1)
				If len(arrNode(uBound(arrNode))) <= 1 or IsNumeric(arrNode(uBound(arrNode))) Then
					arrNode(uBound(arrNode)) = arrNode(uBound(arrNode)-1) + "-" +arrNode(uBound(arrNode))
				End If
				strNodeName1 = arrNode(uBound(arrNode))				
				iRowNo = Fn_SE_BOMTable_RowIndex(strNodeName1)
				If isNumeric(iRowNo) then
					objTable.SelectRow iRowNo
					Call Fn_menuOperation("Select","View:Expand Options:Expand Below")
					If JavaWindow("SystemsEngineering").JavaWindow("SEWindow").JavaDialog("ExpandBelow").exist(3) = False Then
						Call Fn_KeyBoardOperation("SendKeys","{ESC}")
						Call Fn_menuOperation("WinMenuSelect","View:Expand Options:Expand Below")
					End If
					Call Fn_Button_Click("Fn_SE_BOMTableNodeOpeations", JavaWindow("SystemsEngineering").JavaWindow("SEWindow").JavaDialog("ExpandBelow"), "Yes")
					Fn_SE_BOMTableNodeOpeations=True
				End if
				
			Case "SelectAll"
				'Clear previously selected Nodes
				objTable.Object.clearSelection
				iRows = cInt(objTable.GetROProperty ("rows"))
				For iCounter = 0 to iRows - 1
					objTable.ExtendRow iCounter
				Next
				Fn_SE_BOMTableNodeOpeations = True		


			Case "CompareColValue"   'Added by Pooja S :  2/2/2012
					bFlag=false
					sColValue = Fn_SE_BOMTableNodeOpeations("GetCellData",strNodeName,strColName, "", "")
					wait(2)
					aValues = split(sColValue,",",-1,1)
					For iCounter = 0 to Ubound(aValues)				
									For i=0 to Ubound(strColValue)
										If 	lCase(trim(aValues(iCounter)))=lCase(trim(strColValue(i))) Then
												bFlag=true
												Exit for 
										End If	
									Next
						Next	
						If bFlag=true Then
									Fn_SE_BOMTableNodeOpeations=True
						Else
									Fn_SE_BOMTableNodeOpeations=False
						End If
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		Case "PopupMenuEnabled"
			If strNodeName <> "" Then
				arrNode = Split(strNodeName,"-",-1,1)
				If len(arrNode(uBound(arrNode))) <= 1 or IsNumeric(arrNode(uBound(arrNode))) Then
					arrNode(uBound(arrNode)) = arrNode(uBound(arrNode)-1) + "-" +arrNode(uBound(arrNode))
				End If
				strNodeName1 = arrNode(uBound(arrNode))
				iRowNo = Fn_SE_BOMTable_RowIndex(strNodeName1)
				If iRowNo <> -1 Then
					'Split Context menu to Build Path Accordingly
					sMenu = split(strPopupMenu,":",-1,1)
					If sColName = "" Then
						objTable.ClickCell iRowNo ,0, "RIGHT","NONE"
					Else
						objTable.ClickCell iRowNo ,sColName, "RIGHT","NONE"
					End If
					Select Case cInt(Ubound(sMenu))
						Case 0
							Set objContextMenu = JavaWindow("SystemsEngineering").WinMenu("ContextMenu")
							objTable.ClickCell iRowNo,"Requirement", "RIGHT","NONE"
							Wait(2)
							If objContextMenu.CheckItemProperty (strPopupMenu, "Exists",true,10) Then
								Fn_SE_BOMTableNodeOpeations = objContextMenu.CheckItemProperty (strPopupMenu, "Enabled",true,10)
							End IF
							objTable.ClickCell iRowNo,"Requirement", "LEFT","NONE"
							Set objContextMenu = nothing
					End Select
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [Fn_SE_BOMTableNodeOpeations] Popup Menu ["+ strPopupMenu +"] Selected Sucessfully")
				Else
					Fn_SE_BOMTableNodeOpeations = False
				End If
			Else
				Fn_SE_BOMTableNodeOpeations = False
			End If
	'- - - - - - -  Added Case by Sandeep: Case to return All Column names currently exist in BOM Table
			Case "AllColumnNames"
					'Returning All column Names present in BOM Table
					Fn_SE_BOMTableNodeOpeations =Fn_UI_TableOperations("Fn_SE_BOMTableNodeOpeations","GetAllColumnNames",objTable,"","")
		'- - - - - - - - - - - - - Added Case to Remove Column from Table
		Case "RemoveColumn"
					'checking existance of columns 
					bFlag=False
					iColNo = cInt(objTable.GetROProperty("cols"))
					aColumns = split(strColName,"~")
					For iColumnCnt = 0 to UBound(aColumns)
						For iCnt = 0 to iColNo -1
							' if exists then ignoring that column
							If trim(objTable.GetColumnName(iCnt)) = aColumns(iColumnCnt) then
								bFlag=True
								Exit for
							end if
						Next
					Next
					'-------------------------------Modified Case by Chaitali----------------------------------------------
					
					'Change parameter from 1 to 0 by dipali
					'objTable.SelectColumnHeader 1,"RIGHT"
					
					If  iColNo > 1 Then
						objTable.SelectColumnHeader 1,"RIGHT"
						wait 4
					Else
						objTable.SelectColumnHeader 0,"RIGHT"
						wait 4
					End If
					'------------------------------------------------------------------------------------------------------------------
					
					Window("SEWindow").JavaWindow("WEmbeddedFrame").JavaMenu("label:=Insert column\(s\) \.\.\.").Select
					wait 1
					If Window("SEWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Change Columns").Exist(5) Then
						Set objChangeCol = Window("SEWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Change Columns")
						For iColumnCnt = 0 to UBound(aColumns)
							If bFlag=True Then
								Call Fn_List_Select("Fn_SE_BOMTableNodeOpeations", objChangeCol,"ListDisplayedCols",aColumns(iColumnCnt))
								Call Fn_Button_Click("Fn_SE_BOMTableNodeOpeations", objChangeCol,"Remove")
								wait 1
							End If
						Next
						if cInt(objChangeCol.JavaButton("Apply").getROProperty("enabled")) = 1 then
							Call Fn_Button_Click("Fn_SE_BOMTableNodeOpeations", objChangeCol,"Apply")
						end if
						Call Fn_Button_Click("Fn_SE_BOMTableNodeOpeations", objChangeCol,"Cancel")
						Fn_SE_BOMTableNodeOpeations = True
					End if

				Case "ConfigureColumn", "ConfigureColumns"
					' checking existance of columns 
					iColNo = cInt(objTable.GetROProperty("cols"))
					aColumns = split(strColName,"~")
					For iColumnCnt = 0 to UBound(aColumns)
						For iCnt = 0 to iColNo -1
							' if exists then ignoring that column
							If trim(objTable.GetColumnName(iCnt)) = aColumns(iColumnCnt) then
								aColumns(iColumnCnt) = ""
								Exit for
							end if
						Next
					Next

					  objTable.SelectColumnHeader 1,"RIGHT"
					Window("SEWindow").JavaWindow("WEmbeddedFrame").JavaMenu("label:=Insert column\(s\) \.\.\.").Select
					If Window("SEWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Change Columns").Exist(5) Then
						Set objChangeCol = Window("SEWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Change Columns")
			'Delete all the Existing Columns from the Displayed Columns list
				sValues=objChangeCol.JavaList("ListDisplayedCols").GetROProperty("items count")
				For iCount=0 to cInt(sValues)-1
					sNode=objChangeCol.JavaList("ListDisplayedCols").GetItem(iCount)
					objChangeCol.JavaList("ListDisplayedCols").ExtendSelect  sNode
				
				Next
				objChangeCol.JavaButton("Remove").Click micLeftBtn

			'Now add new Columns
			aColumns=split(strColName,":",-1,1)
						For iColumnCnt = 0 to UBound(aColumns)
							If aColumns(iColumnCnt) <> "" Then
								Call Fn_List_Select("Fn_SE_BOMTableNodeOpeations", objChangeCol,"ListAvailableCols",aColumns(iColumnCnt))
								Call Fn_Button_Click("Fn_SE_BOMTableNodeOpeations", objChangeCol,"Add")
								wait 1
							End If
						Next
						if cInt(objChangeCol.JavaButton("Apply").getROProperty("enabled")) = 1 then
							Call Fn_Button_Click("Fn_SE_BOMTableNodeOpeations", objChangeCol,"Apply")
						end if
						Call Fn_Button_Click("Fn_SE_BOMTableNodeOpeations", objChangeCol,"Cancel")
						Fn_SE_BOMTableNodeOpeations = True
					End if
					
		'-------------------------------Added Case by Chaitali : To Move Column Up & Down in BOM Table
		Case "MoveColumnUp","MoveColumnDown"
					'checking existance of columns 
					bFlag=False
					iColNo = cInt(objTable.GetROProperty("cols"))
					aColumns = split(strColName,"~")
					For iColumnCnt = 0 to UBound(aColumns)
						For iCnt = 0 to iColNo -1
							' If exists then ignoring that column
							If trim(objTable.GetColumnName(iCnt)) = aColumns(iColumnCnt) then
								bFlag=True
								Exit for
							end if
						Next
					Next
					
					If  iColNo > 1 Then
						objTable.SelectColumnHeader 1,"RIGHT"
						wait 2
					Else
						objTable.SelectColumnHeader 0,"RIGHT"
						wait 2
					End If
					Window("SEWindow").JavaWindow("WEmbeddedFrame").JavaMenu("label:=Insert column\(s\) \.\.\.").Select
					
					If Window("SEWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Change Columns").Exist(5) Then
						Set objChangeCol = Window("SEWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Change Columns")
						For iColumnCnt = 0 to UBound(aColumns)
							If bFlag=True Then
								Call Fn_List_Select("Fn_SE_BOMTableNodeOpeations", objChangeCol,"ListDisplayedCols",aColumns(iColumnCnt))
								For iCount = 0 To strColValue - 1
									If strAction = "MoveColumnUp" Then
											Call Fn_Button_Click("Fn_SE_BOMTableNodeOpeations", objChangeCol, "Up")
											bFlag = True
									ElseIf strAction = "MoveColumnDown" Then
											Call Fn_Button_Click("Fn_SE_BOMTableNodeOpeations", objChangeCol, "Down")
											bFlag = True
									End If
								Next
								wait 1
							End If
						Next
						if cInt(objChangeCol.JavaButton("Apply").getROProperty("enabled")) = 1 then
							Call Fn_Button_Click("Fn_SE_BOMTableNodeOpeations", objChangeCol,"Apply")
						end if
						Call Fn_Button_Click("Fn_SE_BOMTableNodeOpeations", objChangeCol,"Cancel")
						Fn_SE_BOMTableNodeOpeations = True
					End if			
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			Case "GetNodePathByName"
					iRows=objTable.GetROProperty("rows")
					bFlag=False
					For iCounter=0 to iRows-1	
						If InStr(1,trim(objTable.Object.getPathForRow(iCounter).toString()),strNodeName) then
							iPathCount=objTable.Object.getPathForRow(iCounter).getPathCount()
							StrNodePath=objTable.Object.getPathForRow(iCounter).getPathComponent(1).toString()
							For iCount=2 to iPathCount-1
								StrNodePath=StrNodePath+":"+objTable.Object.getPathForRow(iCounter).getPathComponent(iCount).toString()
							Next
							bFlag=true
							Exit for
						end if
					Next
					If bFlag=True Then
						Fn_SE_BOMTableNodeOpeations=StrNodePath
					else
						Fn_SE_BOMTableNodeOpeations=False
					End If
					
		Case "GetNodePathByName_Ext"  ' Added by Anurag K to get path of more than 1 occurence of node with same name is present in BOM table
			
				iRows=objTable.GetROProperty("rows")
				iCnt=1
				iOccurence = 1
			    bFlag=False
				If instr( strNodeName, "@") > 0 Then
				    aValues = split (strNodeName, "@")
				    iOccurence = aValues(1)
				    strNodeName = aValues(0)
				End IF
				For iCounter=0 to iRows-1	
					 aName = Split(objTable.Object.getPathForRow(iCounter).toString(),",")
					 strName = aName(uBound(aName))
					 strName = Trim(Replace(strName,"]",""))
					
					If InStr(1,strName,strNodeName) > 0 then
						If iCnt = CInt(iOccurence) Then
						    iPathCount=objTable.Object.getPathForRow(iCounter).getPathCount()
						    StrNodePath=objTable.Object.getPathForRow(iCounter).getPathComponent(1).toString()
						    For iCount=2 to iPathCount-1
							   StrNodePath=StrNodePath+":"+objTable.Object.getPathForRow(iCounter).getPathComponent(iCount).toString()
						    Next
					
							bFlag=true
							Exit for 
						Else
						    iCnt= iCnt + 1
						End if
					End If
				Next
				If bFlag=True Then
					Fn_SE_BOMTableNodeOpeations = StrNodePath
				Else
					Fn_SE_BOMTableNodeOpeations=False
				End If

			Case "VerifyForegroundColour", "VerifyBackgroundColour","VerifyForegroundColourExt"

				If strNodeName <> "" Then
					arrNode = Split(strNodeName,"-",-1,1)
					If len(arrNode(uBound(arrNode))) <= 1 or IsNumeric(arrNode(uBound(arrNode))) Then
					arrNode(uBound(arrNode)) = arrNode(uBound(arrNode)-1) + "-" +arrNode(uBound(arrNode))
				End If
					strNodeName1 = arrNode(uBound(arrNode))
					iRowCounter = Fn_SE_BOMTable_RowIndex(strNodeName1)
					If cint(iRowCounter) = -1 Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [Fn_SE_BOMTableNodeOpeations] Couldnt find  BOM Table Node [" + StrNodeName + "]")
						Exit function
					End If
					iRows = iRowCounter +1
					iCount = iRowCounter
				Else
					iRows = objTable.GetROProperty("rows")
					iCount = 0
				End If

				Do While cint(iCount) < cint(iRows)
					Set  objNodeForRow =  objTable.Object.getNodeForRow(cint(iCount))
					' if background colour
					If StrAction = "VerifyBackgroundColour" Then
						sColour = objTable.Object.getBackground(objNodeForRow,False).toString()
					ElseIf StrAction = "VerifyForegroundColourExt" Then
						sColour = objTable.Object.getBackground(objNodeForRow,True).toString()
					Else
					' if foreground colour
						sColour = objTable.Object.getForeground(objNodeForRow,False).toString()
					End If
	
					sColour =  mid(sColour ,instr(sColour ,"[")  ,instr(sColour ,"]") )
					' Comparing colour codes RGB
					Select Case cstr(strPopupMenu)
						Case "BLACK"
							sColourCode = "[r=0,g=0,b=0]"
						Case "LIGHTPINK"
							sColourCode = "[r=253,g=217,b=220]"
						Case "WHITE"
							sColourCode =  "[r=255,g=255,b=255]"
						Case "GRAY"
							sColourCode = "[r=178,g=180,b=191]" 
						'-------------------------------Added Case by Chaitali ------------------------------------------------------------
						Case "LIGHTGRAY"
							sColourCode = "[r=192,g=192,b=192]" 	
						'-------------------------------	---------------------------------------------------------------------------------------------
						Case "DARKGRAY"
							sColourCode = "[r=128,g=128,b=128]"
						Case "DARKBLUE"
							sColourCode = "[r=0,g=0,b=255]" 
						Case "LIGHTBLUE"
							sColourCode = "[r=183,g=219,b=255]"
						Case "GREEN"
							sColourCode = "[r=80,g=176,b=128]"
						Case "DARKGREEN"
							sColourCode = "[r=0,g=255,b=0]"
						Case "LIGHTGREEN"
							sColourCode = "[r=159,g=255,b=159]"
						Case "ORANGE"
							sColourCode = "[r=255,g=200,b=0]"
						Case "RED"
							sColourCode = "[r=255,g=0,b=0]" 
						Case "LIGHTRED"
							sColourCode = "[r=255,g=121,b=121]" 
						Case "YELLOW"
							sColourCode = "[r=255,g=255,b=0]"
						Case "YELLOWISHORANGE"
							sColourCode = "[r=254,g=190,b=95]"
						Case "PAROTGREEN"
							sColourCode = "[r=51,g=153,b=255]"	
						Case Else
							Exit function
					End Select
					
					If sColour = sColourCode  Then
						Fn_SE_BOMTableNodeOpeations = True
					Else
						Fn_SE_BOMTableNodeOpeations = False
						Exit function
					End If
					iCount = iCount +1
				Loop
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [Fn_SE_BOMTableNodeOpeations] Successfully verified colour code [ " & sColourCode & " ] for case [" & strPopupMenu & "]")
				Set objNodeForRow = nothing

		Case "GetNodeFromUID"
				Dim objNode
				'Treat strColValue Variable as UIDof the Node
				'Get Node Object from UID
				strNodeName = Trim(Replace(strNodeName,""""," "))
				Set objNode =  objTable.Object.getNodeBasedOnUID(StrNodeName)
				'Get Row Index from the Path
				If isEmpty(ObjNode)  Then
					Fn_SE_BOMTableNodeOpeations=False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [Fn_SE_BOMTableNodeOpeations] Failed to retrieve BOMLine Node for Shape")
				Else
					Fn_SE_BOMTableNodeOpeations=objNode.toString()
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [Fn_SE_BOMTableNodeOpeations] Sucessfully retrieve BOMLine Node for Shape")
				End If		
	'==========================================================================================================
		Case "VerifyColumnConfigurationValues","SelectColumnConfigurationValues"			'TC1015-2015090100-18_09_2015-VivekA-NewDevelopment-Added by Kaveri P.-Added Case "SelectColumnConfigurationValues" to select Column Configuration
					aColumns = split(strColName,"~")
					'JavaWindow("SystemsEngineering").JavaList("Column Configurations").SetTOProperty "attached text",""
					Wait 1
					JavaWindow("SystemsEngineering").JavaList("Column Configurations").DblClick 10,10,"LEFT"
					wait 1
					For iCounter = 0 to UBound(aColumns)
						bFlag=False
						bFlag = Fn_UI_ListItemExist("Fn_SE_BOMTableNodeOpeations", JavaWindow("SystemsEngineering"),"Column Configurations",aColumns(iCounter))
						If bFlag=false Then
							Fn_SE_BOMTableNodeOpeations=False
							Exit Function
						End If
					Next
					If bFlag=True Then
						If StrAction = "SelectColumnConfigurationValues" Then
							bFlag = Fn_SISW_UI_JavaList_Operations("Fn_SE_BOMTableNodeOpeations","Select",JavaWindow("SystemsEngineering"),"Column Configurations",strColName,"","")
							If bFlag = False Then
								Fn_SE_BOMTableNodeOpeations=False
								Exit Function
							End If							
						End If
						Fn_SE_BOMTableNodeOpeations=True
					End If
	'==========================================================================================================
        Case "GetNodePathByName_Ext"  ' Added by Anurag K to get path of more than 1 occurence of node with same name is present in BOM table
					iRows=objTable.GetROProperty("rows")
					iCnt=1
					iOccurence = 1
				    bFlag=False
					If instr( strNodeName, "@") > 0 Then
					    aValues = split (strNodeName, "@")
					    iOccurence = aValues(1)
					    strNodeName = aValues(0)
					End IF
					For iCounter=0 to iRows-1	
						 aName = Split(objTable.Object.getPathForRow(iCounter).toString(),",")
						 strName = aName(uBound(aName))
						 strName = Trim(Replace(strName,"]",""))
						
						If InStr(1,strName,strNodeName) > 0 then
							If iCnt = CInt(iOccurence) Then
							    iPathCount=objTable.Object.getPathForRow(iCounter).getPathCount()
							    StrNodePath=objTable.Object.getPathForRow(iCounter).getPathComponent(1).toString()
							    For iCount=2 to iPathCount-1
								   StrNodePath=StrNodePath+":"+objTable.Object.getPathForRow(iCounter).getPathComponent(iCount).toString()
							    Next
						
								bFlag=true
								Exit for 
							Else
							    iCnt= iCnt + 1
							End if
						End If
					Next
					If bFlag=True Then
						Fn_SE_BOMTableNodeOpeations = StrNodePath
					Else
						Fn_SE_BOMTableNodeOpeations=False
					End If
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------			
			Case "SetColumnWidth"	'[TC11.3-2017050900-15_06_2017-JotibaT-NewDevlopment] 
					If strColName<>"" Then
						iColNo = Fn_SISW_UI_JavaTable_Operations("Fn_SE_BOMTableNodeOpeations","GetColumnIndex",objTable, "", "", strColName, "", "", "", "", "")
						objTable.Object.getColumnModel().getColumn(iColNo).setPreferredWidth(strColValue)
						Fn_SE_BOMTableNodeOpeations=True
					End If
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------					
			Case "GetColumnWidth"	'[TC11.3-2017050900-15_06_2017-JotibaT-NewDevlopment] 
					If strColName<>"" Then
						iColNo = Fn_SISW_UI_JavaTable_Operations("Fn_SE_BOMTableNodeOpeations","GetColumnIndex",objTable, "", "", strColName, "", "", "", "", "")
						Fn_SE_BOMTableNodeOpeations=objTable.Object.getColumnModel().getColumn(iColNo).getWidth()
					End If
'--------------------------------------------------------------------------------------------------------------------------------------------------------------				
			'[TC1015-2015090100-18_09_2015-AnkitN-NewDevlopment] - To select all rows between given two nodes --------------------------------------------------
			Case "SelectRowsRange"		'("SelectRowsRange","REQ-000049/A;1-Req1 (View):000455/A;1-P1 (View):000575/A;1-P23~REQ-000049/A;1-Req1 (View):REQ-000148/A;1-Req2","","","")
				strNodeName=split(strNodeName,"~")
				arrNode = Split(strNodeName(0),"-",-1,1)
				If len(arrNode(uBound(arrNode))) <= 1 or IsNumeric(arrNode(uBound(arrNode))) Then
					arrNode(uBound(arrNode)) = arrNode(uBound(arrNode)-1) + "-" +arrNode(uBound(arrNode))
				End If
				strNodeName1 = arrNode(uBound(arrNode))
				iStartRowNo = Fn_SE_BOMTable_RowIndex(strNodeName1)
				arrNode1 = Split(strNodeName(1),"-",-1,1)
				If len(arrNode1(uBound(arrNode1))) <= 1 or IsNumeric(arrNode1(uBound(arrNode1))) Then
					arrNode1(uBound(arrNode1)) = arrNode1(uBound(arrNode1)-1) + "-" +arrNode1(uBound(arrNode1))
				End If
				strNodeName2 = arrNode1(uBound(arrNode1))
				iEndRowNo   = Fn_SE_BOMTable_RowIndex(strNodeName2)
				If isNumeric(iStartRowNo) AND isNumeric(iEndRowNo) Then
					objTable.SelectRowsRange iStartRowNo,iEndRowNo
					Fn_SE_BOMTableNodeOpeations=True
				End If
			Case "Delete"	
				arrNode = Split(strNodeName,"-",-1,1)
				If len(arrNode(uBound(arrNode))) <= 1 or IsNumeric(arrNode(uBound(arrNode))) Then
					arrNode(uBound(arrNode)) = arrNode(uBound(arrNode)-1) + "-" +arrNode(uBound(arrNode))
				End If
				strNodeName1 = arrNode(uBound(arrNode))
				iRowNo = Fn_SE_BOMTable_RowIndex(strNodeName1)			
					'iRowNo = Fn_SE_BOMTable_RowIndex(strNodeName)
					If isNumeric(iRowNo) then
						objTable.SelectRow iRowNo	
						call Fn_MenuOperation("Select","Edit:Delete")
						JavaDialog("Delete").JavaCheckBox("DeleteAllSequences").Set "OFF"						
						Call Fn_Button_Click("Fn_TcObjectRemove", JavaDialog("Delete"),"Yes")
						Fn_SE_BOMTableNodeOpeations=True							
					End if		
			Case "Remove"							'Tc10.1.5-2015090100-14_Sep_2015-AnkitN-Added Case "Remove" to Remove a node
				arrNode = Split(strNodeName,"-",-1,1)
				If len(arrNode(uBound(arrNode))) <= 1 or IsNumeric(arrNode(uBound(arrNode))) Then
					arrNode(uBound(arrNode)) = arrNode(uBound(arrNode)-1) + "-" +arrNode(uBound(arrNode))
				End If
				strNodeName1 = arrNode(uBound(arrNode))
				iRowNo = Fn_SE_BOMTable_RowIndex(strNodeName1)
				'iRowNo = Fn_SE_BOMTable_RowIndex(strNodeName)
				If isNumeric(iRowNo) then
					objTable.SelectRow iRowNo						
					Call Fn_MenuOperation("Select","Edit:Remove")
					Call Fn_Button_Click("Fn_TcObjectRemove", JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Remove"),"Yes")
					Fn_SE_BOMTableNodeOpeations=True					
				End if
			'[TC1122-20160413-25_04_2016-VivekA-Maintenance] - Added from TC1016
			Case "GetAllInternalColumnNames"
				aColumns = cInt(objTable.getROProperty("cols"))
				For iCounter = 0 to aColumns-1
					If iCounter = 0 then
						Fn_SE_BOMTableNodeOpeations = ObjTable.Object.getColumnIdentify(iCounter)
					Else
						Fn_SE_BOMTableNodeOpeations = Fn_SE_BOMTableNodeOpeations & "~" & ObjTable.Object.getColumnIdentify(iCounter)
					End If
				Next
							
				'---------------------------------------------------------------------------------
				Case "ExpandAndSelect"
					'Initial Item Path
					aStrNode = Split (StrNodeName, ":")
					For i = 0 to UBound(aStrNode)-1
						If sParentPath = "" Then
							sParentPath  = aStrNode(i)
						Else
							sParentPath  = sParentPath + ":" + aStrNode(i)
						End If
						Call Fn_SE_BOMTableNodeOpeations("Expand", sParentPath, "","","")

					Next
				
					call Fn_SE_BOMTableNodeOpeations("Select",strNodeName,"","","")
					Fn_SE_BOMTableNodeOpeations=True
'------------------------------------------------------------------------------------------------------------------------------------------------------------------
				Case "VerifyLOVValues" '[TC11.3(20160131_NewDevelopment_PoonamC_31Mar2017):Added case to verify values for Content Type Lov]
				arrNode = Split(strNodeName,"-",-1,1)
				If len(arrNode(uBound(arrNode))) <= 1 or IsNumeric(arrNode(uBound(arrNode))) Then
					arrNode(uBound(arrNode)) = arrNode(uBound(arrNode)-1) + "-" +arrNode(uBound(arrNode))
				End If
				strNodeName1 = arrNode(uBound(arrNode))					
					iRowNo = Fn_SE_BOMTable_RowIndex(strNodeName1)
	            	If isNumeric(iRowNo) then
						'Get column Rows
						iColNo = objTable.GetROProperty("cols")
	
						For iStart=0 to iColNo-1
							''Verify the Column name is similar to required column name
							If objTable.GetColumnName(iStart)=strColName Then
								If strColName = "Content Type" Then 
									  Set ObjAppletWin = Window("SEWindow").JavaWindow("WEmbeddedFrame")
									  objTable.ClickCell iRowNo,strColName
									   wait 2
									   ObjAppletWin.SetTOProperty "index",2
									   ObjAppletWin.JavaButton("dropdown_16").Click
									   wait 2
									   'get LOV Table object
									   Set objLov=Description.Create()
											objLov("Class Name").value="JavaTable"
											objLov("tagname").value="LOVTreeTable"
										Set objChild = ObjAppletWin.ChildObjects(objLov)
									         Wait 1
										'Traverse through the value list
										strColValue = Split(strColValue,"~")
										For iCnt1 = 0 To UBound(strColValue)
												bFlag = False									
												For iCnt = 0 To objChild.Count	
														If trim(objChild(1).Object.getValueAt(iCnt,0).getDisplayableValue())=trim(strColValue(iCnt1)) Then
															bFlag=true
															Exit For
														End If
												Next											
		                  						 If bFlag=true Then
												 	Fn_SE_BOMTableNodeOpeations=True
												 Else
												 	Fn_SE_BOMTableNodeOpeations=False
												 	Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [Fn_SE_BOMTableNodeOpeations] Failed to verify [ "&strColValue(iCnt1)&" ]")
												 	Exit Function
												 End If 
										Next
										Call Fn_KeyBoardOperation("SendKeys","{ESC}")
										Call Fn_ReadyStatusSync(1)
										
										Set objChild = Nothing
										Set ObjAppletWin = Nothing
										Set objLov = Nothing
								End If		
							End If
						Next
					Else
						Fn_SE_BOMTableNodeOpeations=False
					End if
	'----------------------------------------------------------------------------------------------------------------------------------------					
		 '[TC11.4_2017080100_NewDevelopment_PoonamC_23Aug2017: Added new case to select multiple Nodes & select popup menu on it ]
		 Case "MultiSelectPopupMenuSelect"
				strNodeName=split(strNodeName,"~")
				For iNodeNo=0 to Ubound(strNodeName)
					arrNode1 = Split(strNodeName(iNodeNo),":",-1,1)
					arrNode = Split(arrNode1(uBound(arrNode1)),"-",-1,1)
			
					iRowNo = Fn_SE_BOMTable_RowIndex(arrNode(uBound(arrNode)))
					If isNumeric(iRowNo) Then
						If iNodeNo=0 Then
							objTable.SelectRow iRowNo
							Fn_SE_BOMTableNodeOpeations=True
						Else
							objTable.ExtendRow "#"&iRowNo
							Fn_SE_BOMTableNodeOpeations=True
						End If					
					End if
				Next
				
				strPopupMenu=Replace(strPopupMenu,":",";")
				iRowNo = objTable.Object.getSelectedRow()
				If isNumeric(iRowNo) then
					objTable.ClickCell iRowNo,0,"RIGHT" 
					wait 1
					sMenu = JavaWindow("SystemsEngineering").WinMenu("ContextMenu").BuildMenuPath(strPopupMenu)
					wait 1
					JavaWindow("SystemsEngineering").WinMenu("ContextMenu").Select sMenu
					Fn_SE_BOMTableNodeOpeations=True
				End if
'----------------------------------------------------------------------------------------------------------------------------------------		
		 '[TC11.4_2017080100_NewDevelopment_PoonamC_23Aug2017]
		 Case "CompareColValueExt"
		 			If strNodeName<>"" Then
			 			arrNode1 = Split(strNodeName,":",-1,1)
						arrNode = Split(arrNode1(uBound(arrNode1)),"-",-1,1)
						iRowNo = Fn_SE_BOMTable_RowIndex(arrNode(uBound(arrNode)))
						If isNumeric(iRowNo) Then
								sColValue = objTable.GetCellData(iRowNo,strColName)	
								Wait 2
								aValues = split(sColValue,",",-1,1)
								strColValue = split(strColValue,",",-1,1)
								For iCounter = 0 to Ubound(strColValue)
										bFlag=false								
										For i=0 to Ubound(aValues)
											If 	lCase(trim(aValues(i)))=lCase(trim(strColValue(iCounter))) Then
													bFlag=true
													Exit for 
											End If	
										Next
										If bFlag=false Then
											Fn_SE_BOMTableNodeOpeations=False
											Exit Function
										End If
								Next
								Fn_SE_BOMTableNodeOpeations=True															
						 End if
		 			End If 		
'----------------------------------------------------------------------------------------------------------------------------------------
		 '[TC11.4_2017080100_NewDevelopment_PoonamC_23Aug2017]
		 Case "GetColumnValue"
		 			If strNodeName<>"" Then
			 			arrNode1 = Split(strNodeName,":",-1,1)
						arrNode = Split(arrNode1(uBound(arrNode1)),"-",-1,1)
						iRowNo = Fn_SE_BOMTable_RowIndex(arrNode(uBound(arrNode)))
						If isNumeric(iRowNo) Then
								Fn_SE_BOMTableNodeOpeations = objTable.GetCellData(iRowNo,strColName)	
								Wait 2															
						 End if
		 			End If
'----------------------------------------------------------------------------------------------------------------------------------------		
		End Select
'	Else
'		'RMTable not displayed in Requirement Manager!
'		Fn_SE_BOMTableNodeOpeations=False
'	End if
	Set objTable = nothing
End Function


 '*********************************************************		Function to Report Operation	***********************************************************************
'Function Name		:				Fn_SE_TraceabilityReportOperations

'Description			 :		 		 Perform Operations on "Traceability Report" Dialog

'Parameters			   :	 			1.sAction: DefiningTable:Properties (First is Table Name Compulsory)
'													 2.sNodeName: Node on which we have to perform operation
'													 3.sNewName:New Name in Property
'													4.sColName:Column Name
'													5.sCellValue:Cell Value  												

'Return Value		   : 				True or False

'Pre-requisite			:		 	Must be Selected Trace Link Node

'Examples				:			'Fn_SE_TraceabilityReportOperations("DefiningTable:Properties","000494-Test3:Test4->Test3","NewName3","","") -->This will change the Name property  and press OK on TraceabilityReport
'												'Fn_SE_TraceabilityReportOperations("DefiningTable:Expand","000494-Test3:Test4->Test3","","","") --->This Case Expand the tree node but it will not  press OK on TraceabilityReport
'												'Fn_SE_TraceabilityReportOperations("DefiningTable:Select","000494-Test3:Test4->Test3:000495-Test4","","","") --->This Case Select the tree node but it will not  press OK on TraceabilityReport
'												'Fn_SE_TraceabilityReportOperations("DefiningTable:Verify","000494-Test3:Test4->Test3:000495-Test4","","","") ---->This Case Verify the tree node but it will not  press OK on TraceabilityReport
'												'Fn_SE_TraceabilityReportOperations("DefiningTable:Properties","000494-Test3:Test4->Test3","Change Name","","") 
'												'Fn_SE_TraceabilityReportOperations("DefiningTable:Go To Object","Change Name","","","") 
'												'Fn_SE_TraceabilityReportOperations("DefiningTable:Delete Trace Link","Change Name","","","") 
'												'Fn_SE_TraceabilityReportOperations("DefiningTable:CellVerify","","","Relation Type","Trace Link") 
'												'Fn_SE_TraceabilityReportOperations("DefiningTable:Refresh Report","","","","")--->This will refresh the TraceabilityReport  window and close the TraceabilityReport
'												'Fn_SE_TraceabilityReportOperations("DefiningTable:DescriptionProperties","000494-Test3:Test4->Test3","New Description","","") 
'												'Fn_SE_TraceabilityReportOperations("ComplyingTable:VerifyDescriptionProperties","REQ-000043-req2","Sonal","","")
'												'Fn_SE_TraceabilityReportOperations("ComplyingTable:Expand","000257-Item:Item->Req1@2","","","")
'												'Fn_SE_TraceabilityReportOperations("TraceabilityReport_Defining:FindInView","000432-Process3","000430/A;1-Process1 (View)","","")
'												'Fn_SE_TraceabilityReportOperations("ComplyingTable:GoToObject_DropdownSelect","","","","Manufacturing Process Planner")
'History					 :		
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer			Build
'												-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'												Harshal	 Agrawal								   31/01/2011			              1.0										Created														20110119
'												Harshal	 Agrawal								   04/05/2011			              1.0										All OR and Logic Modified 				 20110330
'												Koustubh Watwe								   	   01/08/2011			              1.0										Added Object hierarchy for Refresh Window
'												Ketan Raje												12/09/2011							1.0										Added code to handle Instance of Node.
'												Ketan Raje												16/09/2011							1.0										Added Case : "FindInView"
'												Sagar Shivade										27/09/11							1.0										Added case :- ''GoToObject_DropdownSelect''
'												Koustubh										27/09/11							1.0										Added case :- ''handle delete dialog while deleting tracability''
'												Sonal										26/03/13							1.0										Added case :- GoToObject_DropdownSelectWithError
'												-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SE_TraceabilityReportOperations(sAction,sNodeName,sNewName,sColName,sCellValue)
	GBL_FAILED_FUNCTION_NAME="Fn_SE_TraceabilityReportOperations"
	'Declaring All Varaibles
	Dim aAction, sTableName, iRows, iCounter, sNodePath, sIndex, bFlag, ArrLists, iToolCnt, sContents, sCellData
	'Declaring All Object's
	Dim objJavaDialogReport, ObjDesc, iMatch, iInstance, sInstance, ObjTable, aMenuList, sMenu,scrollMax,sTitle
	iMatch = 0
		
		'Code for Trace link operations in System Engineering Perspective
'------------------------------------------------------------------------------------------------------------------------------------
	sTitle = JavaWindow("DefaultWindow").GetROProperty("title")
	If (instr (sTitle, "Systems Engineering") > 0) Then
		aAction = Split(sAction, ":")
		If strComp(aAction(0),"TraceabilityReport_Defining")=0  Then
			sAction = Replace (sAction, "TraceabilityReport_Defining", "DefiningTable")
		Elseif strComp(aAction(0),"TraceabilityReport_Complying")=0  Then
			sAction = Replace (sAction, "TraceabilityReport_Complying","ComplyingTable")
		End If
		If strComp(aAction(1), "CellVerify")= 0 Then
			sAction = Replace (sAction, "CellVerify","VerifyCellValue")
		ElseIf strComp(aAction(1), "Verify")= 0 Then
			sAction= Replace (sAction, "Verify","NodeVerify")
		ElseIf strComp(aAction(1), "Refresh Report")= 0 Then
			Fn_SE_TraceabilityReportOperations = True
			Exit Function	
			sAction= Replace (sAction, "Refresh Report","RefreshReport")
		ElseIf strComp(aAction(1), "Delete Trace Link")= 0 Then
			sAction = Replace (sAction, "Delete Trace Link","DeleteTraceLink")			
		End If
		bReturn = Fn_SE_TraceLinkOpeartions(sAction,sNodeName,sNewName,sColName,sCellValue)
		Fn_SE_TraceabilityReportOperations =bReturn 
		Exit Function
	End If
'------------------------------------------------------------------------------------------------------------------------------------
		sInstance = False
		'Spliting sAction To retriewe Table name
		aAction=Split(sAction,":")
		sTableName=aAction(0)
		'Setting bFlag
		bFlag=False
		'Creating Object "Traceability Report" Dialog
		Select Case sTableName
				Case "ComplyingTable"
					JavaWindow("Traceability").SetTOProperty "title","Complying Traceability Report"
					Set objJavaDialogReport=Fn_UI_ObjectCreate("Fn_SE_TraceabilityReportOperations", JavaWindow("Traceability"))
				Case "DefiningTable"
					JavaWindow("Traceability").SetTOProperty "title","Defining Traceability Report"
					Set objJavaDialogReport=Fn_UI_ObjectCreate("Fn_SE_TraceabilityReportOperations", JavaWindow("Traceability"))
				Case "TraceabilityReport_Defining"
					Set objJavaDialogReport=Fn_UI_ObjectCreate("Fn_SE_TraceabilityReportOperations", JavaWindow("Traceability"))
					JavaWindow("Traceability").JavaTable("TraceabilityReportPanel").SetTOProperty "index",0
				Case "TraceabilityReport_Complying"
					Set objJavaDialogReport=Fn_UI_ObjectCreate("Fn_SE_TraceabilityReportOperations", JavaWindow("Traceability"))
					JavaWindow("Traceability").JavaTable("TraceabilityReportPanel").SetTOProperty "index",1
		End Select
		If aAction(1) <> "RMBSelectPopupMenu" Then
			Set ObjDesc = Description.Create() 
			ObjDesc("to_class").Value = "JavaToolbar" 
			ObjDesc("enabled").Value = 1
			'Get the total of Toolbar objects
			Set ArrLists =objJavaDialogReport.ChildObjects(ObjDesc)
			iToolCnt = objJavaDialogReport.ChildObjects(ObjDesc).count
			'Checking "Show Trace Link" button present or not
			For iCounter = 0 to iToolCnt-1
					sContents = ArrLists(iCounter).GetContent()
					If instr(sContents, "Show Trace Link") > 0 Then
					'Clicking "Show Trace Link" button
					ArrLists(iCounter).Press "Show Trace Link"
					bFlag=True
					Exit For
					End If
			Next
			If sNodeName <> "" Then
				If Instr(1,sNodeName,"@") <> 0 Then
					sInstance = True
					aNode = Split(sNodeName,"@")
					iInstance = aNode(1)
					sNodeName = aNode(0)
				End If
			End If
	
	
			'Retriwing No of rows
			iRows = Fn_Table_GetRowCount("Fn_SE_TraceabilityReportOperations",objJavaDialogReport,"TraceabilityReportPanel")
	
			If sNodeName<>"" And sInstance = False Then
							For iCounter = 0 to iRows -1
								objJavaDialogReport.JavaTable("TraceabilityReportPanel").SelectRowsRange iCounter,iCounter
								sNodePath=objJavaDialogReport.JavaTable("TraceabilityReportPanel").GetCellData(iCounter,0)
								'Checking "sNodeName" present in table or not
								If Trim(sNodePath) = Trim(sNodeName) Then
									sIndex = Cstr(iCounter)
									bFlag=True
									Exit For
								End If
							Next
							If iCounter = Cint(iRows) Then
								Fn_SE_TraceabilityReportOperations =False
								Exit Function
							End If
			ElseIf sNodeName <> "" And sInstance = True Then
							For iCounter = 0 to iRows -1
								objJavaDialogReport.JavaTable("TraceabilityReportPanel").SelectRowsRange iCounter,iCounter
								sNodePath=objJavaDialogReport.JavaTable("TraceabilityReportPanel").GetCellData(iCounter,0)
								'Checking "sNodeName" present in table or not
								If Trim(sNodePath) = Trim(sNodeName) Then
									sIndex = Cstr(iCounter)
									iMatch = iMatch + 1
									If CInt(Cstr(iMatch)) = Cint(Cstr( iInstance)) Then
										bFlag=True
										Exit For
									End If
								End If
							Next
							If iCounter = iRows Then
								Fn_SE_TraceabilityReportOperations =False
								Exit Function
							End If
			End If
		End If
        Select Case aAction(1)
			'To Expand Tree Node of table
			Case "Expand"
					objJavaDialogReport.JavaTable("TraceabilityReportPanel").SelectRow sIndex
					'objJavaDialogReport.JavaTable("TraceabilityReportPanel").DoubleClickCell sIndex,0	Commented by Ketan on 12/09/2011 and added below line.
                    objJavaDialogReport.JavaTable("TraceabilityReportPanel").ActivateRow sIndex
					Fn_SE_TraceabilityReportOperations =True
					Exit Function
			Case "ExpandBelow"		
					objJavaDialogReport.JavaTable("TraceabilityReportPanel").SelectRow sIndex
					objJavaDialogReport.JavaTable("TraceabilityReportPanel").Object.ExpandBelow
                    Fn_SE_TraceabilityReportOperations =True
					Exit Function
			'To Select Tree Node of table		
			Case "Select"
					objJavaDialogReport.JavaTable("TraceabilityReportPanel").SelectRow sIndex
					Fn_SE_TraceabilityReportOperations =True
					Exit Function
			'Verifying Node is present or not	
			Case "Verify"
					If bFlag=True Then
						Fn_SE_TraceabilityReportOperations =True
						Exit Function
					Else
						Fn_SE_TraceabilityReportOperations =False
						Exit Function
					End If
			'To Verifying Node is present or not and close the TraceabilityReport.
			'Added case "VerifyWithClose" by Ujwal N
			Case "VerifyWithClose"
				If bFlag=True Then
					Fn_SE_TraceabilityReportOperations =True					
				Else
					Fn_SE_TraceabilityReportOperations =False
					Exit Function
				End If
			'To Select Tree Node and perform FindInView operation.
			Case "FindInView"
					objJavaDialogReport.JavaTable("TraceabilityReportPanel").SelectRow sIndex
					'Select value from FindInView list.
					Call Fn_List_Select("Fn_SE_TraceabilityReportOperations", JavaWindow("Traceability"), "FindInView",sNewName)
					'Click on FindInView button
					Call Fn_Button_Click("Fn_SE_TraceabilityReportOperations",JavaWindow("Traceability"),"FindInView")
					Fn_SE_TraceabilityReportOperations =True

			  Case "RMBSelectPopupMenu"
					If sTableName="ComplyingTable" Then
						Set ObjTable = objJavaDialogReport.JavaTable("ComplyingTable")
					ElseIf sTableName="DefiningTable" Then
						Set ObjTable = objJavaDialogReport.JavaTable("DefiningTable")
					Else
						Set ObjTable = objJavaDialogReport.JavaTable("TraceabilityReportPanel")
					End If
					ObjTable.SelectColumnHeader 0,"RIGHT"
									aMenuList = split(sNodeName, ":",-1,1)
									iCounter = cstr(Ubound(aMenuList))
									Select Case iCounter
										Case "0"
											JavaWindow("Traceability").JavaMenu("MainMenu").SetTOProperty "label",aMenuList(0)
											JavaWindow("Traceability").JavaMenu("MainMenu").Select
										Case "1"
											JavaWindow("Traceability").JavaMenu("MainMenu").SetTOProperty "label",aMenuList(0)
											JavaWindow("Traceability").JavaMenu("MainMenu").JavaMenu("LevelOne").SetTOProperty "label",aMenuList(1)
											JavaWindow("Traceability").JavaMenu("MainMenu").JavaMenu("LevelOne").Select
										Case Else											
											Fn_AuditLogTableRMB = FALSE
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_AuditLogTableRMB Failed to select Menu "&sNodeName)											
									End Select								
					Fn_SE_TraceabilityReportOperations = True
					Exit Function

			'To modify properties	
			Case "Properties"
					If Fn_UI_ObjectExist("Fn_SE_TraceabilityReportOperations",objJavaDialogReport)=True Then
						'Pressing "Properties" button to change "Name" property
						For iCounter = 0 to iToolCnt-1
							sContents = ArrLists(iCounter).GetContent()
								If instr(sContents, "Properties") > 0 Then
									ArrLists(iCounter).Press "Properties"
									wait(5)
									'Changing the "Name"
								'Call Fn_Edit_Box("Fn_SE_TraceabilityReportOperations",JavaWindow("Traceability").JavaDialog("Properties"),"Name",sNewName)
								JavaWindow("Traceability").JavaDialog("Properties").JavaEdit("Name").Set ""
								JavaWindow("Traceability").JavaDialog("Properties").JavaEdit("Name").Type sNewName
								Call Fn_Button_Click("Fn_SE_TraceabilityReportOperations",JavaWindow("Traceability").JavaDialog("Properties"),"OK")
									Fn_SE_TraceabilityReportOperations = True
									Exit For
								End If
						Next
			
						If iCounter = iToolCnt Then
							Fn_SE_TraceabilityReportOperations = FALSE
						End If
					Else
						Fn_SE_TraceabilityReportOperations = FALSE
					End If
					'Refrefing the report
					For iCounter = 0 to iToolCnt-1
						sContents = ArrLists(iCounter).GetContent()
							If instr(sContents, "Refresh Report") > 0 Then
								ArrLists(iCounter).Press "Refresh Report"
								bFlag=True
								Exit For
							End If
					Next
     	Case "VerifyProperties"                  'Added by Avinash j.     16/july/2012
		   If Fn_UI_ObjectExist("Fn_SE_TraceabilityReportOperations",objJavaDialogReport)=True Then
			'Pressing "Properties" button to verify "Description" property
			For iCounter = 0 to iToolCnt-1
			 sContents = ArrLists(iCounter).GetContent()
			  If instr(sContents, "Properties") > 0 Then
			   ArrLists(iCounter).Press "Properties"
			   wait(5)
			   JavaWindow("Traceability").JavaDialog("Properties").JavaStaticText("EmptyProperties").SetTOProperty "label","Hide empty properties..."
			   wait(5)
			   If Fn_UI_ObjectExist("Fn_SE_TraceabilityReportOperations",JavaWindow("Traceability").JavaDialog("Properties").JavaStaticText("EmptyProperties"))=False Then
				JavaWindow("Traceability").JavaDialog("Properties").JavaStaticText("EmptyProperties").SetTOProperty "label","Show empty properties..."
				JavaWindow("Traceability").JavaDialog("Properties").JavaStaticText("EmptyProperties").Click 1,1,"LEFT"
				wait(5)
			   End If
			   'Changing the "Name"
				sAppValue= Fn_Edit_Box_GetValue("Fn_SE_TraceabilityReportOperations",JavaWindow("Traceability").JavaDialog("Properties"),"Name")
			   If Trim(Lcase(sAppValue)) =Trim(Lcase(sNewName))Then
				  Call Fn_Button_Click("Fn_SE_TraceabilityReportOperations",JavaWindow("Traceability").JavaDialog("Properties"),"OK")
				  Fn_SE_TraceabilityReportOperations = True
			   Else
				  Call Fn_Button_Click("Fn_SE_TraceabilityReportOperations",JavaWindow("Traceability").JavaDialog("Properties"),"OK")
				  Fn_SE_TraceabilityReportOperations = False
			   End If
			   
						   Exit For
			  End If
			Next
		 
			If iCounter = iToolCnt Then
				Fn_SE_TraceabilityReportOperations = FALSE
			End If
		   Else
			Fn_SE_TraceabilityReportOperations = FALSE
		   End If

			'To modify Description property
			Case "DescriptionProperties"
					If Fn_UI_ObjectExist("Fn_SE_TraceabilityReportOperations",objJavaDialogReport)=True Then
						'Pressing "Properties" button to change "Name" property
						For iCounter = 0 to iToolCnt-1
							sContents = ArrLists(iCounter).GetContent()
								If instr(sContents, "Properties") > 0 Then
									ArrLists(iCounter).Press "Properties"
'									Call Fn_SISW_UI_JavaToolbar_Operations("Fn_SE_TraceabilityReportOperations", "DeviceReplay.Click", JavaWindow("Traceability"), "ToolBar", "Properties", "", "", "")
									wait(5)
									If JavaWindow("Traceability").JavaDialog("Properties").JavaSlider("JScrollPane").Exist(3) Then
										scrollMax=JavaWindow("Traceability").JavaDialog("Properties").JavaSlider("JScrollPane").GetROProperty("max")
										JavaWindow("Traceability").JavaDialog("Properties").JavaSlider("JScrollPane").Drag scrollMax
										wait 1
									End If
									JavaWindow("Traceability").JavaDialog("Properties").JavaStaticText("EmptyProperties").SetTOProperty "label","Hide empty properties..."
									wait(5)
									If Fn_UI_ObjectExist("Fn_SE_TraceabilityReportOperations",JavaWindow("Traceability").JavaDialog("Properties").JavaStaticText("EmptyProperties"))=False Then
										JavaWindow("Traceability").JavaDialog("Properties").JavaStaticText("EmptyProperties").SetTOProperty "label","Show empty properties..."
										JavaWindow("Traceability").JavaDialog("Properties").JavaStaticText("EmptyProperties").Click 1,1,"LEFT"
										wait(5)
									End If
								'Changing the "Name"
									Call Fn_Edit_Box("Fn_SE_TraceabilityReportOperations",JavaWindow("Traceability").JavaDialog("Properties"),"Description",sNewName)
									Call Fn_Button_Click("Fn_SE_TraceabilityReportOperations",JavaWindow("Traceability").JavaDialog("Properties"),"OK")
									Fn_SE_TraceabilityReportOperations = True
'									Exit For
								End If
						Next
			
'						If iCounter = iToolCnt Then
'							Fn_SE_TraceabilityReportOperations = FALSE
'						End If
					Else
						Fn_SE_TraceabilityReportOperations = FALSE
					End If
					'Refrefing the report
					Call Fn_SISW_UI_JavaToolbar_Operations("Fn_SE_TraceabilityReportOperations", "DeviceReplay.Click", JavaWindow("Traceability"), "ToolBar", "Refresh Report", "", "", "")
'					For iCounter = 0 to iToolCnt-1
'						sContents = ArrLists(iCounter).GetContent()
'							If instr(sContents, "Refresh Report") > 0 Then
'								ArrLists(iCounter).Press "Refresh Report"
'								bFlag=True
'								Exit For
'							End If
'					Next

	
	Case "VerifyDescriptionProperties"
		If Fn_UI_ObjectExist("Fn_SE_TraceabilityReportOperations",objJavaDialogReport)=True Then
			'Pressing "Properties" button to verify "Description" property
'			For iCounter = 0 to iToolCnt-1
'				sContents = ArrLists(iCounter).GetContent()
'				If instr(sContents, "Properties") > 0 Then
'					ArrLists(iCounter).Press "Properties"
					Call Fn_SISW_UI_JavaToolbar_Operations("Fn_SE_TraceabilityReportOperations", "DeviceReplay.Click", JavaWindow("Traceability"), "ToolBar", "Properties", "", "", "")
					wait 5
					JavaWindow("Traceability").JavaDialog("Properties").JavaStaticText("EmptyProperties").SetTOProperty "label","Hide empty properties..."
					If Fn_UI_ObjectExist("Fn_SE_TraceabilityReportOperations",JavaWindow("Traceability").JavaDialog("Properties").JavaStaticText("EmptyProperties"))=False Then
						JavaWindow("Traceability").JavaDialog("Properties").JavaStaticText("EmptyProperties").SetTOProperty "label","Show empty properties..."
						JavaWindow("Traceability").JavaDialog("Properties").JavaStaticText("EmptyProperties").Click 1,1,"LEFT"
						If JavaWindow("Traceability").JavaDialog("Properties").JavaEdit("Description").Exist(15) = False Then
							Fn_SE_TraceabilityReportOperations = False
							Exit function
						End If
					End If
					'Changing the "Name"
					sAppValue= Fn_Edit_Box_GetValue("Fn_SE_TraceabilityReportOperations",JavaWindow("Traceability").JavaDialog("Properties"),"Description")
					If Trim(Lcase(sAppValue)) =Trim(Lcase(sNewName))Then
						Call Fn_Button_Click("Fn_SE_TraceabilityReportOperations",JavaWindow("Traceability").JavaDialog("Properties"),"OK")
						Fn_SE_TraceabilityReportOperations = True
					Else
						Call Fn_Button_Click("Fn_SE_TraceabilityReportOperations",JavaWindow("Traceability").JavaDialog("Properties"),"OK")
						Fn_SE_TraceabilityReportOperations = False
					End If
'					Exit For
'				End If
'			Next
'			If iCounter = iToolCnt Then
'				Fn_SE_TraceabilityReportOperations = FALSE
'			End If
		Else
			Fn_SE_TraceabilityReportOperations = FALSE
		End If
		
			'To Delete Trace Link		
			Case "Delete Trace Link"
				'Pressing "Delete Trace Link" button to Delete Trace Link
				For iCounter = 0 to iToolCnt-1
					sContents = ArrLists(iCounter).GetContent()
					If instr(sContents, "Delete Trace Link") > 0 Then
							ArrLists(iCounter).Press "Delete Trace Link"
							bFlag=True
							wait(2)
							' clicking on Yes button of delete dialog.
							If JavaWindow("Traceability").JavaDialog("Delete").Exist(5) Then
								JavaWindow("Traceability").JavaDialog("Delete").JavaButton("Yes").Click micLeftBtn
							End If
		

							If Fn_UI_ObjectExist("Fn_SE_TraceabilityReportOperations",JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Refresh Window")) Then
								Call Fn_Button_Click("Fn_SE_TraceabilityReportOperations",JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Refresh Window"),"Yes")
							ElseIf Fn_UI_ObjectExist("Fn_SE_TraceabilityReportOperations",JavaWindow("Defining Traceability").JavaDialog("Refresh Window")) Then
								Call Fn_Button_Click("Fn_SE_TraceabilityReportOperations",JavaWindow("Defining Traceability").JavaDialog("Refresh Window"),"Yes")
						    End If
							Exit For
					End If
				Next
					If  bFlag=True Then
						Fn_SE_TraceabilityReportOperations=True
					Else
						Fn_SE_TraceabilityReportOperations=False
					End If
			'To go Object		
			Case "Go To Object"
				'Pressing "Go To Object" button to Delete Trace Link
				For iCounter = 0 to iToolCnt-1
					sContents = ArrLists(iCounter).GetContent()
						If instr(sContents, "Go To Object") > 0 Then
							ArrLists(iCounter).Press "Go To Object"
							bFlag=True
							Exit For
						End If
				Next
					If  bFlag=True Then
						Fn_SE_TraceabilityReportOperations=True
					Else
						Fn_SE_TraceabilityReportOperations=False
					End If
			  Case "CellVerify"
				 If sTableName="ComplyingTable" Then
						iRows=Fn_Table_GetRowCount("Fn_SE_TraceabilityReportOperations",objJavaDialogReport,sTableName)
								For iCounter=0 to iRows-1
									sCellData=objJavaDialogReport.JavaTable("ComplyingTable").GetCellData(iCounter,sColName)
										If sCellData=sCellValue Then
											bReturn=True
											Exit For
									   End If
								Next
				  ElseIf sTableName = "DefiningTable" Then
						iRows=Fn_Table_GetRowCount("Fn_SE_TraceabilityReportOperations",objJavaDialogReport,sTableName)
							For iCounter=0 to iRows-1
								   sCellData=objJavaDialogReport.JavaTable("DefiningTable").GetCellData(iCounter,sColName)
									If sCellData=sCellValue Then
											bReturn=True
											Exit For
								   End If
							Next
				  Else
						iRows=Fn_Table_GetRowCount("Fn_SE_TraceabilityReportOperations",objJavaDialogReport,"TraceabilityReportPanel")
							For iCounter=0 to iRows-1
								   sCellData=objJavaDialogReport.JavaTable("TraceabilityReportPanel").GetCellData(iCounter,sColName)
									If sCellData=sCellValue Then
											bReturn=True
											Exit For
								   End If
							Next
				  End If
					
					If bReturn=True Then
						Fn_SE_TraceabilityReportOperations=True
					Else
						Fn_SE_TraceabilityReportOperations=False
					End If
				Case "Refresh Report"
					'Refrefing the report
					For iCounter = 0 to iToolCnt-1
						sContents = ArrLists(iCounter).GetContent()
							If instr(sContents, "Refresh Report") > 0 Then
								ArrLists(iCounter).Press "Refresh Report"
								bReturn=True
								Exit For
							End If
					Next
					If bReturn=True Then
						Fn_SE_TraceabilityReportOperations=True
					Else
						Fn_SE_TraceabilityReportOperations=False
					End If
                				'To go Object	Dropdown	
			Case "GoToObject_DropdownSelect"
				Dim objMenu
				'Pressing "Go To Object" Dropdown button and Select application from dropdownlist. ( Node must be selected)
				For iCounter = 0 to iToolCnt-1
						sContents = ArrLists(iCounter).GetContent()
						If instr(sContents, "down_16") > 0 Then
								'Clicking "Show Trace Link" button
								ArrLists(iCounter).Press "down_16"
								wait(2)
								'Added work around code to highlight java menu : Sandeep Navghane
								ArrLists(iCounter).Press "down_16"
								wait(2)
								ArrLists(iCounter).Press "down_16"
								wait(2)
								' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	'							JavaWindow("Traceability").JavaMenu("label:="&sCellValue,"index:=0").Select   			
	
'								'By Shreyas  [12-03-2012]
'								 JavaWindow("Traceability").JavaMenu("MainMenu").SetTOProperty "label",sCellValue
'								  JavaWindow("Traceability").JavaMenu("MainMenu").SetTOProperty "index",0
'								  wait 2
'								  JavaWindow("Traceability").JavaMenu("MainMenu").Select
'								bFlag=True
'								*Added by Nilesh on 25-Feb-2013
								If JavaMenu("Menu").Exist(5)=True Then
									Set objMenu=JavaMenu("Menu")
								Else
									Set objMenu=JavaWindow("Traceability").JavaMenu("MainMenu")
								End If

								 objMenu.SetTOProperty "label",sCellValue
								  objMenu.SetTOProperty "index",0
								  wait 2
								  objMenu.Select
								bFlag=True
								Exit For
							'*End
						End If
				Next
					If  bFlag=True Then
						Fn_SE_TraceabilityReportOperations=True
					Else
						Fn_SE_TraceabilityReportOperations=False
					End If
					Set objMenu=Nothing
			Case "GoToObject_DropdownSelectWithError"			
				'Pressing "Go To Object" Dropdown button and Select application from dropdownlist. ( Node must be selected)
				For iCounter = 0 to iToolCnt-1
						sContents = ArrLists(iCounter).GetContent()
						If instr(sContents, "down_16") > 0 Then
								'Clicking "Show Trace Link" button
								ArrLists(iCounter).Press "down_16"
								wait(2)
								'Added work around code to highlight java menu : Sandeep Navghane
								ArrLists(iCounter).Press "down_16"
								wait(2)
								ArrLists(iCounter).Press "down_16"
								wait(2)
								' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
								If JavaMenu("Menu").Exist(5)=True Then
									Set objMenu=JavaMenu("Menu")
								Else
									Set objMenu=JavaWindow("Traceability").JavaMenu("MainMenu")
								End If

								 objMenu.SetTOProperty "label",sCellValue
								  objMenu.SetTOProperty "index",0
								  wait 2
								  objMenu.Select
								bFlag=True
								Exit For
							'*End
						End If
				Next
					If  bFlag=True Then
						Fn_SE_TraceabilityReportOperations=True
					Else
						Fn_SE_TraceabilityReportOperations=False
					End If
					Set objMenu=Nothing
					Exit Function
		Case "PopupSelect"
				'Pre-requisite = Row should be selected
					objJavaDialogReport.JavaTable("TraceabilityReportPanel").ClickCell sIndex,0,"RIGHT" 
					wait 1
					JavaWindow("Traceability").JavaMenu("MainMenu").SetTOProperty "label",sNewName
					JavaWindow("Traceability").JavaMenu("MainMenu").Select
					Fn_SE_TraceabilityReportOperations=True
					Exit Function
					
		Case "VerifyListBoxProperties"
				If Fn_UI_ObjectExist("Fn_SE_TraceabilityReportOperations",objJavaDialogReport.JavaDialog("Properties"))=False Then
					objJavaDialogReport.highlight
					Call Fn_SISW_UI_JavaToolbar_Operations("Fn_SE_TraceabilityReportOperations", "DeviceReplay.Click", JavaWindow("Traceability"), "ToolBar", "Properties", "", "", "")
					wait 5
				End If
				
				If Fn_UI_ObjectExist("Fn_SE_TraceabilityReportOperations",objJavaDialogReport.JavaDialog("Properties"))=True Then
					
					Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_SISW_Prop_VerifyProperties",objJavaDialogReport.JavaDialog("Properties").JavaList("Property"),"attached text",sColName+":")
					If objJavaDialogReport.JavaDialog("Properties").JavaList("Property").Exist(SISW_MICRO_TIMEOUT) Then
							aValues=Split(sCellValue,"~")
							For iCounter=0 to uBound(aValues)
								bFlag=false
								'Verifying value exist in list or not
								'taking item count from list
								iEleCount=Fn_UI_Object_GetROProperty("Fn_SISW_Prop_VerifyProperties",objJavaDialogReport.JavaDialog("Properties").JavaList("Property"), "items count")
								For iCount=0 to iEleCount-1
									If objJavaDialogReport.JavaDialog("Properties").JavaList("Property").GetItem(iCount)=aValues(iCounter) Then
										bFlag=true
										Fn_SE_TraceabilityReportOperations=True
										Call Fn_Button_Click("Fn_SE_TraceabilityReportOperations",objJavaDialogReport.JavaDialog("Properties"),"OK")
										Exit for
									End If
								Next
							Next
							If bFlag=False Then
								Call Fn_Button_Click("Fn_SE_TraceabilityReportOperations",objJavaDialogReport.JavaDialog("Properties"),"OK")
								Fn_SE_TraceabilityReportOperations=False
							End If
					Else
							Call Fn_Button_Click("Fn_SE_TraceabilityReportOperations",objJavaDialogReport.JavaDialog("Properties"),"OK")
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"List box [ " & sColName & " ] is not exist on dialog")
							Fn_SE_TraceabilityReportOperations= False
					End If
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Properties Dialog does not exist on dialog")
					Fn_SE_TraceabilityReportOperations=False
				End If
		End Select
	'Clicking on "OK" Button of "Report" Dialog
	Call Fn_Button_Click("Fn_SE_TraceabilityReportOperations",objJavaDialogReport,"OK")
	
	Set ObjDesc = Nothing
	Set ArrLists = Nothing
	Set objJavaDialogReport=Nothing
End Function
'*********************************************************		Function for Importing Requirement Specification ***********************************************************************
'Function Name		:			Fn_SE_ImportReqSpec

'Description		:			This function is used to Import Specification

'Parameters			:			1.	sFileName:sFileName is complete path of file with file name and its extension
'											 2.	sSpecType:Specification Type
'											3.	sDescription:Desription Of Impoerted specification
'											4.sOption: "Import "  OR "Keyword"
'											5.sKeywords:
'											6.sSubType:Specification Sub Type

'											Imp Note:sOption is Either  "Import "  OR "Keyword"
'Return Value		:		True/False

'Pre-requisite		:		Requirement Manager window should be displayed .

'Examples			:				
									'Call Fn_SE_ImportReqSpec("D:\mainline\TestData\Requierment Data File.docx","RequirementSpec","Test","Import","","Requirement")
									'Call Fn_SE_ImportReqSpec("D:\mainline\TestData\Requierment Data File.docx","RequirementSpec","Test","Keyword","Program","Requirement")
									'Call Fn_SE_ImportReqSpec("D:\mainline\TestData\Req_RM037\Requierment_Data_File.docx","RequirementSpec","","Keyword","~Test","")-To Verify value of Keiwords edit box
									'Call Fn_SE_ImportReqSpec("D:\mainline\TestData\Requierment Data File.docx","RequirementSpec","Test~True","Keyword","Program","Requirement") - to set import as child of selected element.
'History:					
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	Build
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Harshal 				31/01/2011			1.0														20110119
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Amit T. 				20/10/2011			1.0					Added code to set import as child of selected element checkbox.
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'sFileName is complete path of file with file name and its extension
Public Function Fn_SE_ImportReqSpec(sFileName,sSpecType,sDescription,sOption,sKeywords,sSubType)
	GBL_FAILED_FUNCTION_NAME="Fn_SE_ImportReqSpec"
	On Error Resume Next
	'Declaring Variables 
	Dim iItemCnt,iCnt,bFlag,strSource,strVerify , aDescription
	'Declaring Objects
	Dim objImportDialog,objImportSpecDialog
	
	If sSpecType="RequirementSpec" Then
		sSpecType="Requirement Specification"
	End If
	Fn_SE_ImportReqSpec=False
	bFlag=False
	'Setting Object of  "Import Spec" Dilog
	Set objImportDialog=JavaWindow("SystemsEngineering").JavaWindow("SEWindow").JavaDialog("Import Spec")
	'Checking "Import Spec" Dialog Is exist or Not 
		If Fn_UI_ObjectExist("Fn_SE_ImportReqSpec",objImportDialog)=False Then
			'Calling Menuoperation to Open Import Spec Dialog
			Call Fn_MenuOperation("Select","File:Import Spec...")
		End If
	'Creating object of  "Import Spec" 
	Set objImportSpecDialog=Fn_UI_ObjectCreate("Fn_SE_ImportReqSpec",JavaWindow("SystemsEngineering").JavaWindow("SEWindow").JavaDialog("Import Spec"))
		'Checking File Name is pass or not
		'File Name is Compulsory Parameter if not pass then function will exit
		If sFileName<>"" Then
            Call Fn_Edit_Box("Fn_SE_ImportReqSpec",objImportSpecDialog,"FileName",sFileName)
		Else
			Set objImportDialog=Nothing
			Set objImportSpecDialog=Nothing
			Exit Function
		End If
		'Checking Specification Type  pass or not
		If sSpecType<>"" Then
			'Retriwing Items Count from "SpecType" Java List
			iItemCnt= Fn_UI_Object_GetROProperty("Fn_SE_ImportReqSpec",objImportSpecDialog.JavaList("SpecType"),"items count")
			For iCnt=0 To iItemCnt-1
				If  objImportSpecDialog.JavaList("SpecType").GetItem(iCnt)=sSpecType Then
					objImportSpecDialog.JavaList("SpecType").Select(sSpecType)
					'If item is present in list then select it and changing bFlag to True
					bFlag=True
					Exit For
				End If
			Next
		End If
	'If Wrong Item Name is pass then function will exit and return False
	If bFlag=False Then
		Exit Function
	End If
	
	' Split [ sDescription ] to Obtain Description and CheckBox value
	
	If sDescription <> "" Then
		If Instr( sDescription , "~" ) = 0 Then 'Directly set the Description as it is
			Call Fn_Edit_Box("Fn_SE_ImportReqSpec",objImportSpecDialog,"Description",sDescription)
		Else
			'Split sDescription to get Description and CheckBox value
			aDescription = Split( sDescription , "~" )

			'1. Set Description
			Call Fn_Edit_Box("Fn_SE_ImportReqSpec",objImportSpecDialog,"Description",aDescription(0))
			Call Fn_ReadyStatusSync(2)
			
			'2. CheckBox [ ON/OFF ]
			If cBool(aDescription(1)) Then
				 Call Fn_CheckBox_Set("Fn_SE_ImportReqSpec",objImportSpecDialog,"ImportAsChildOfSelected", "ON" )
			Else
				Call Fn_CheckBox_Set("Fn_SE_ImportReqSpec",objImportSpecDialog,"ImportAsChildOfSelected", "OFF" )
			End IF
			Call Fn_ReadyStatusSync(2)
		End If
	End If

			'aDescription = Split( sDescription , "~" )
			
			'If sDescription<>"" Then
				'Setting Description
			'	Call Fn_Edit_Box("Fn_SE_ImportReqSpec",objImportSpecDialog,"Description",sDescription)
		'	End If
			
	'End If
			
	'Clicking On Next Button to go on Next Window
	Call Fn_Button_Click("Fn_SE_ImportReqSpec", objImportSpecDialog, "Next")
	'selection Import Option
	If Ucase(sOption)="IMPORT" Then
		'Setting "ImportAsSingleSubtype" Radion Button To "ON"
		objImportSpecDialog.JavaRadioButton("ImportAsSingleSubtype").Set "ON"
		bFlag=False
		If sSubType<>"" Then
			'Selecting Item From "ImportSubType" List
			iItemCnt= Fn_UI_Object_GetROProperty("Fn_SE_ImportReqSpec",objImportSpecDialog.JavaList("ImportSubType"),"items count")
			For iCnt=0 To iItemCnt-1
				If objImportSpecDialog.JavaList("ImportSubType").GetItem(iCnt)=sSubType Then
					objImportSpecDialog.JavaList("ImportSubType").Select(sSubType)
					bFlag=True
					Exit For
				End If
			Next
		
				If bFlag=False Then
					Call Fn_Button_Click("Fn_SE_ImportReqSpec",objImportSpecDialog, "Close")
					Set objImportDialog=Nothing
					Set objImportSpecDialog=Nothing
					Exit Function
				End If
		End If
	End If
	'selection Keyword Option
	If Ucase(sOption)="KEYWORD" Then
		'Setting "ImportAsSingleSubtype" Radion Button To "ON"
		objImportSpecDialog.JavaRadioButton("UseKwdsForImport").Set "ON"
		bFlag=False
		If sKeywords<>"" Then
				If inStr(1,sKeywords,"~")>0 Then
				strSource=objImportSpecDialog.JavaEdit("Keywords").GetROProperty("value")
				strVerify=split(sKeywords,"~")(1)
				 If  inStr(1,strSource,strVerify)>0 then
					 Fn_SE_ImportReqSpec=True
					 Call Fn_Button_Click("Fn_SE_ImportReqSpec",objImportSpecDialog, "Close")
                    Set objImportSpecDialog=Nothing
					Set objImportDialog=Nothing
					Exit Function
				 Else
					'Clicking on "OK" To save
					Call Fn_Button_Click("Fn_SE_ImportReqSpec",objImportSpecDialog, "Close")

				 End if
 					Fn_SE_ImportReqSpec=False
					Set objImportSpecDialog=Nothing
					Set objImportDialog=Nothing
					Exit Function
			 Else
			'Setting KeyWords
			Call Fn_Edit_Box("Fn_SE_ImportReqSpec",objImportSpecDialog,"Keywords",sKeywords)
			End If
		End If
		'Selecting Item From "KwdSubType" List
		If  sSubType<>"" Then
			iItemCnt= Fn_UI_Object_GetROProperty("Fn_SE_ImportReqSpec",objImportSpecDialog.JavaList("KwdSubType"),"items count")
			For iCnt=0 To iItemCnt-1
				If objImportSpecDialog.JavaList("KwdSubType").GetItem(iCnt)=sSubType Then
					objImportSpecDialog.JavaList("KwdSubType").Select(sSubType)
					bFlag=True
					Exit For
				End If
			Next

			If bFlag=False Then
				Call Fn_Button_Click("Fn_SE_ImportReqSpec",objImportSpecDialog, "Close")
				Set objImportDialog=Nothing
				Set objImportSpecDialog=Nothing
				Exit Function
			End If
		End If
	End If
	'Clicking On finish Button
	If objImportSpecDialog.JavaButton("Finish").GetROProperty("enabled")  Then
		objImportSpecDialog.JavaButton("Finish").Object.doClick 1
	End If
    Fn_SE_ImportReqSpec=True
	Set objImportDialog=Nothing
	Set objImportSpecDialog=Nothing
End Function
'*********************************************************Function For Dialog Verification ***********************************************************************
'Function Name		:			Fn_SE_DialogMsgVerify

'Description		:			Function Used to Handle Dialog under System Engineering Prespective

'Parameters			:			sDialogTitle,sMsg,sButton
'											
'Return Value		:		True/False

'Pre-requisite		:		Subject Dialog Should Exist

'Examples			:				sMsg = "Importing Specification by keyword requires Teamcenter Extension for Microsoft to be installed. Please install Teamcenter Extension for Microsoft and try the operation again."
'												MsgBox Fn_SE_DialogMsgVerify("Import Spec",sMsg,"OK")
'												MsgBox  Fn_SE_DialogMsgVerify(MSWordSaveMessageVerify,"None of the changes you make to the Rich Content for the selected Teamcenter Item","")
'History:					
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	Build
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Harshal 				01/02/2011			1.0																								20110119
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function Fn_SE_DialogMsgVerify(sDialogTitle,sMsg,sButton)

	Dim dicErrorInfo
	Set dicErrorInfo = CreateObject("Scripting.Dictionary")
	dicErrorInfo.Add "Message", sMsg
	dicErrorInfo.Add "Button", sButton
    dicErrorInfo.Add "Action", sDialogTitle
	dicErrorInfo.Add "Title", sDialogTitle
	Fn_SE_DialogMsgVerify = Fn_SISW_SE_ErrorVerify(dicErrorInfo)

End Function 
'*********************************************************Function For Viewer Tab ***********************************************************************
'Function Name		:			Fn_SE_ViewerTabOperations

'Description		:			Function Used to Handle operation in viewer Tab

'Parameters			:			sAction,sProperty,sValue
'											
'Return Value		:		True/False

'Pre-requisite		:		Viewer tab should be open in System Prespective

'Examples			:				MsgBox Fn_SE_ViewerTabOperations("Verify","Object","REQ-000017/A;1-rewq")
'												MsgBox Fn_SE_ViewerTabOperations("Verify","Name","rewq")
'												MsgBox Fn_SE_ViewerTabOperations("Verify","ID","REQ-000017")
'												MsgBox Fn_SE_ViewerTabOperations("Verify","Revision","A")
'												Msgbox Fn_SE_ViewerTabOperations("ModifyEditBox","Intel Owner","auto_rw23821")
'												Msgbox Fn_SE_ViewerTabOperations("ModifyMultiOptionList","HW_SW~Req Category","SW:HW:Other~Technology:BIOS")
'												Msgbox Fn_SE_ViewerTabOperations("VerifyEditBox","Intel Owner~Revision","auto_rw23821~A")
'												Msgbox Fn_SE_ViewerTabOperations("VerifyJavaList","HW_SW~Req Category","SW:HW:Other~Technology:BIOS"

'History:					
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	Build
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Harshal 				01/02/2011			1.0																								20110119
'										Sandeep 			  04/06/2012		  1.1			Added Case : ModifyEditBox,ModifyMultiOptionList,VerifyEditBox,VerifyJavaList
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function Fn_SE_ViewerTabOperations(sAction,sProperty,sValue)
	GBL_FAILED_FUNCTION_NAME="Fn_SE_ViewerTabOperations"
   Dim aProperty,aValue,bFlag,iCounter,iCount,arrValues,sAppMsg
   Select Case sAction
			'- - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - 
		 	'Case to Modify value of Edit Box
		 	Case "ModifyEditBox"
				aProperty=Split(sProperty,"~")
				aValue=Split(sValue,"~")
				For iCounter=0 to ubound(aProperty)
					JavaWindow("SystemsEngineering").JavaStaticText("PropertyName").SetTOProperty "label",aProperty(iCounter)+":"
                    If JavaWindow("SystemsEngineering").JavaEdit("PropertyEdit").Exist(2) Then
						Call Fn_Edit_Box("Fn_SE_ViewerTabOperations",JavaWindow("SystemsEngineering"),"PropertyEdit", aValue(iCounter))
						bFlag=true
					else
						bFlag=false
						Exit for
					End If
				Next
				If bFlag=True Then
					Fn_SE_ViewerTabOperations=true
				else
					Fn_SE_ViewerTabOperations=false
				End If
			'- - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - 
		 	'Case to Verify value of Edit Box
		 	Case "VerifyEditBox"
				aProperty=Split(sProperty,"~")
				aValue=Split(sValue,"~")
				For iCounter=0 to ubound(aProperty)
					JavaWindow("SystemsEngineering").JavaStaticText("PropertyName").SetTOProperty "label",aProperty(iCounter)+":"
                    If JavaWindow("SystemsEngineering").JavaEdit("PropertyEdit").Exist(2) Then
						If Fn_Edit_Box_GetValue("Fn_SE_ViewerTabOperations",JavaWindow("SystemsEngineering"),"PropertyEdit")=aValue(iCounter) Then
							bFlag=true
						Else
							bFlag=false
							Exit for
						End If
					else
						bFlag=false
						Exit for
					End If
				Next
				If bFlag=True Then
					Fn_SE_ViewerTabOperations=true
				else
					Fn_SE_ViewerTabOperations=false
				End If
			'- - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - 
		 	'Case to Modify value of Multi Option List
			Case "VerifyJavaList"
				aProperty=Split(sProperty,"~")
				aValue=Split(sValue,"~")
				For iCounter=0 to ubound(aProperty)
					JavaWindow("SystemsEngineering").JavaStaticText("PropertyName").SetTOProperty "label",aProperty(iCounter)+":"
                    If Window("SEWindow").JavaWindow("WEmbeddedFrame").Exist(2) Then
						arrValues=Split(aValue(iCounter),":")
						For iCount=0 to ubound(arrValues)
							bFlag=Fn_UI_ListItemExist("Fn_SE_ViewerTabOperations", Window("SEWindow").JavaWindow("WEmbeddedFrame"), "PropertyValueList",arrValues(iCount))
							If bFlag=False Then
								Exit for
							End If
						Next
					else
						bFlag=false
						Exit for
					End If
					If bFlag=False Then
						Exit for
					End If
				Next
				If bFlag=True Then
					Fn_SE_ViewerTabOperations=true
				else
					Fn_SE_ViewerTabOperations=false
				End If

			'- - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - 
		 	'Case to Modify value of Multi Option List
			Case "ModifyMultiOptionList"
				aProperty=Split(sProperty,"~")
				aValue=Split(sValue,"~")
				For iCounter=0 to ubound(aProperty)
					JavaWindow("SystemsEngineering").JavaStaticText("PropertyName").SetTOProperty "label",aProperty(iCounter)+":"
                    If Window("SEWindow").JavaWindow("WEmbeddedFrame").Exist(2) Then
						Call Fn_CheckBox_Set("Fn_SE_ViewerTabOperations", Window("SEWindow").JavaWindow("WEmbeddedFrame"), "edit_16","on")
						arrValues=Split(aValue(iCounter),":")
						For iCount=0 to ubound(arrValues)
							Window("SEWindow").JavaWindow("WEmbeddedFrame").JavaEdit("ComboBoxEdit").Object.setText arrValues(iCount)
							Call Fn_Button_Click("Fn_SE_ViewerTabOperations", Window("SEWindow").JavaWindow("WEmbeddedFrame"), "add_16")
						Next
						Call Fn_CheckBox_Set("Fn_SE_ViewerTabOperations", Window("SEWindow").JavaWindow("WEmbeddedFrame"), "edit_16","off")
						bFlag=true
					else
						bFlag=false
						Exit for
					End If
				Next
				If bFlag=True Then
					Fn_SE_ViewerTabOperations=true
				else
					Fn_SE_ViewerTabOperations=false
				End If
			'- - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - 
			Case "Verify"
				Select Case sProperty
						Case "Object"
							sAppMsg =  Fn_UI_Object_GetROProperty("Fn_SE_ViewerTabOperations",JavaWindow("SystemsEngineering").JavaWindow("WEmbeddedFrame").JavaEdit("ViewerObject"), "value")
						Case "Name"
							sAppMsg =  Fn_UI_Object_GetROProperty("Fn_SE_ViewerTabOperations",JavaWindow("SystemsEngineering").JavaWindow("WEmbeddedFrame").JavaEdit("ViewerName"), "value")
						Case "ID"
							sAppMsg =  Fn_UI_Object_GetROProperty("Fn_SE_ViewerTabOperations",JavaWindow("SystemsEngineering").JavaWindow("WEmbeddedFrame").JavaEdit("ViewerID"), "value")
						Case "Revision"
							sAppMsg =  Fn_UI_Object_GetROProperty("Fn_SE_ViewerTabOperations",JavaWindow("SystemsEngineering").JavaWindow("WEmbeddedFrame").JavaEdit("ViewerRevision"), "value")
						Case "Description"
							sAppMsg =  Fn_UI_Object_GetROProperty("Fn_SE_ViewerTabOperations",JavaWindow("SystemsEngineering").JavaWindow("WEmbeddedFrame").JavaEdit("ViewerDescription"), "value")
						Case Else
							Fn_SE_ViewerTabOperations = False
'							Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Fail: "+sProperty+" not found")
							Exit function
				End Select
				If trim(lcase(sAppMsg)) = trim(lcase(sValue)) Then
					Fn_SE_ViewerTabOperations = True
							Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Pass: Verification sucessful")
				End If
			Case "VerifyHiperLink"
				aProperty=Split(sProperty,"~")
				aValue=Split(sValue,"~")
				For iCounter=0 to ubound(aProperty)
					JavaWindow("SystemsEngineering").JavaStaticText("PropertyName").SetTOProperty "label",aProperty(iCounter)+":"
                    If JavaWindow("SystemsEngineering").JavaObject("ImageHyperlink").Exist(2) Then
							HText=JavaWindow("SystemsEngineering").JavaObject("ImageHyperlink").Object.getText
							If Trim(HText)=Trim(aValue(iCounter)) Then
									bFlag=True
                         else
								bFlag=false
						Exit for
						End If
					End If
				Next
				If bFlag=True Then
					Fn_SE_ViewerTabOperations=true
				else
					Fn_SE_ViewerTabOperations=false
				End If

	Case "CheckOut"
			Call Fn_Button_Click("Fn_SE_ViewerTabOperations",JavaWindow("SystemsEngineering").JavaWindow("WEmbeddedFrame"),"Check-Out and Edit")
			If JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Check-Out").Exist Then
				Call Fn_Button_Click("Fn_SE_ViewerTabOperations",JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Check-Out"),"Yes")
			End If
			Fn_SE_ViewerTabOperations = True
			Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Pass: Click Check-Out and Edit Button sucessful")
	Case Else
			Fn_SE_ViewerTabOperations = False
			Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Fail: "+sAction+" not found")
   End Select
End Function

'-------------------------------------------------------------------Function Used to Genarate 6 Digit Random Number------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_SE_RandNoGenerate

'Description			 :	Function Used to Genarate 6 Digit Random Number
										
'Return Value		   : 	Random Number

'Examples				:	Fn_SE_RandNoGenerate

'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done					Build
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Harshal Agrawal					   						01/02/2011						1.0																					20110119
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SE_RandNoGenerate()
	 Dim iNumber,iScnd,iNum,iHr,iDay
	 Randomize
	 iScnd=Second(Now)
	iHr=Hour(Now)
	iDay=Day(Now)
	iNum=987123+iScnd+iHr+iDay
	 iNumber = Int((iNum * Rnd) + 1)
	 If Len(Cstr(iNumber)) < 5 Then
			Fn_SE_RandNoGenerate = "0" + Cstr(iNumber)+"6"
	 ElseIf Len(Cstr(iNumber)) < 6 Then
			Fn_SE_RandNoGenerate = "0" + Cstr(iNumber)
	 ElseIf Len(Cstr(iNumber)) > 6 Then
			iNumber = Int((900000 * Rnd) + 1)
				If Len(Cstr(iNumber)) < 5 Then
					Fn_SE_RandNoGenerate = "0" + Cstr(iNumber)+"6"
				ElseIf Len(Cstr(iNumber)) < 6 Then
					Fn_SE_RandNoGenerate = "0" + Cstr(iNumber)
				Else
					Fn_SE_RandNoGenerate = Cstr(iNumber)
				End If
	 Else
			Fn_SE_RandNoGenerate = Cstr(iNumber)
	 End If
End Function
 '*********************************************************		Function to Create Specification in SE ***********************************************************************
'Function Name		:				Fn_SE_CustomRequirementCreate

'Description			 :		 		 This function is used to Create the Custom Requirement in System Engineering

'Parameters			   :	 			1. strNodeName: Select the Requirement Spec
'													2. strCustID: ID of the Specification
'												   3. strCustRev: Revision of the Spec
'												  4. strCusrName: Name of the Spec
'												  5.strCustDesc: Description of the Spec

'Return Value		   : 				TRUE \ FALSE

'Pre-requisite			:		 		should be in SE Prespective

'Examples				:				 Call Fn_SE_CustomRequirementCreate("RequirementSpec","","","NewSpec","Description")

'History:
'										Developer Name			Date			Rev. No.			Changes Done			Reviewer	      Build
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Harshal Agrawal			02/02/2011																								    20110119
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function Fn_SE_CustomRequirementCreate(strNodeName,strCustID,strCustRev,strCustName,strCustDesc)
	GBL_FAILED_FUNCTION_NAME="Fn_SE_CustomRequirementCreate"
	Dim strNodePath,intNodeCount,sTreeItem,intCount,strNodePathC
	Dim ObjSpecWnd,bFlag

		Fn_SE_CustomRequirementCreate=False
	'Verifying "New Specification" window's existance
	If Fn_UI_ObjectExist("Fn_SE_CustomRequirementCreate",JavaWindow("SystemsEngineering").JavaWindow("NewCustomRequirement"))=False Then
		'Invoking "New Specification" Window
		Call Fn_MenuOperation("Select","File:New:Custom Note")
	End If
	'Creating Object of "New Specification" window
	Set ObjSpecWnd=Fn_UI_ObjectCreate("Fn_SE_CustomRequirementCreate",JavaWindow("SystemsEngineering").JavaWindow("NewCustomRequirement"))
	Call Fn_UI_JavaTree_Expand("Fn_SE_CustomRequirementCreate", ObjSpecWnd, "CustomRequirementTree","Complete List")
	JavaWindow("SystemsEngineering").JavaWindow("NewCustomRequirement").JavaTree("CustomRequirementTree").WaitProperty "items count" , micGreaterThan(1)
	If Fn_UI_JavaTree_NodeExist("Fn_SE_RequirementSpecCreate",ObjSpecWnd.JavaTree("CustomRequirementTree"),"Complete List:"+strNodeName) Then
			strNodePathC="Complete List:"+strNodeName
	Else
			strNodePathC="Most Recently Used:"+strNodeName
	End If
			Call Fn_JavaTree_Select("Fn_SE_CustomRequirementCreate", ObjSpecWnd, "CustomRequirementTree",strNodePathC)
		   Call Fn_JavaTree_Select("Fn_SE_CustomRequirementCreate", ObjSpecWnd, "CustomRequirementTree","Complete List")
		   Call Fn_JavaTree_Select("Fn_SE_CustomRequirementCreate", ObjSpecWnd, "CustomRequirementTree",strNodePathC)
		   Call Fn_Button_Click("Fn_SE_CustomRequirementCreate",ObjSpecWnd,"Next")

	If strCustID<>"" Then
		'Setting Id
		Call Fn_UI_EditBox_Type("Fn_SE_CustomRequirementCreate",ObjSpecWnd,"RequirnmentID",strCustID)
	End If
	If strCustRev<>"" Then
		'Setting Revision
		Call Fn_UI_EditBox_Type("Fn_SE_CustomRequirementCreate",ObjSpecWnd,"Revision",strCustRev)
	End If
	'Setting Name
    Call Fn_UI_EditBox_Type("Fn_SE_CustomRequirementCreate",ObjSpecWnd,"Name",strCustName)
	'Setting Description
	Call Fn_UI_EditBox_Type("Fn_SE_CustomRequirementCreate",ObjSpecWnd,"Description",strCustDesc)
	'Clicking On Finish Button To finish the Operation
	Call Fn_Button_Click("Fn_SE_CustomRequirementCreate",ObjSpecWnd,"Finish")
	'function Return True
	Fn_SE_CustomRequirementCreate=True
	'Releasing "New Specification" window's object
	Set ObjChangeWnd=Nothing
End Function
'-------------------------------------------------------------------------This Function is used to to perform operation on MS Word tab---------------------------------------------------------------
'Function Name		:	Fn_SE_MSWordTabOperations

'Description			:	This Function is used to to perform operation on MS Word tab

'Parameters			:			1.	strAction:Action Name
'											 2.	 strValue:Value to set in text Box Or to verify the value
											'3.	strParameterName:Parameter Name

'Return Value		:	True/False

'Pre-requisite		:	Object should be selected
'											
'Examples			:	Fn_SE_MSWordTabOperations("SetValue","[parameter1:value1,value2,value3..]"+VbCrlf+"[parameter2: value1,value2,value3..]","")
'									Fn_SE_MSWordTabOperations("VerifyValue","value1","parameter1")
'							Case "VerifyInStr"	 : Fn_SE_MSWordTabOperations("VerifyInStr","[humidity:9,8,7]","")	'Case Added By Ketan Raje on 15-Nov-2010
'History:
'									Developer Name												Date								Rev. No.						Changes Done				Reviewer	 Build
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'									Amol								   02-Feb-2011		              1.0										Created							Tushar B     20110119
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SE_MSWordTabOperations(strAction,strValue,strParameterName)
	GBL_FAILED_FUNCTION_NAME="Fn_SE_MSWordTabOperations"
	'Variable declaration
   Dim bReturn,bFlag,iRowCnt,sCellData,sValueName,iCount
   	Dim sAttachedText,sText
   'Function Return False
   Fn_SE_MSWordTabOperations=False
   bReturn=False
   bFlag=False
'   Verifying MS Word tab is activated or not
   bFlag=Fn_TabFolder_Operation("VerifyActivate", "MS Word","")
   If bFlag=False Then
	   'Activating MS Word tab
	   Call Fn_SetView("Teamcenter:MS Word")
   End If
	Select Case strAction
		'"SetValue" this action set the value 
		Case "SetValue"		'Fn_SE_MSWordTabOperations("SetValue","[parameter1:value1,value2,value3..]"+VbCrlf+"[parameter2: value1,value2,value3..]","")
				'Setting Value in text box
                Call Fn_Edit_Box("Fn_SE_MSWordTabOperations",JavaWindow("MyTeamcenter"),"MSWordTabText",strValue)
				'Saving the changes
                Call Fn_ToolbatButtonClick("Save")
				Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Pass:Successfully Set  value"& strValue)   
				'Function Return True
				Fn_SE_MSWordTabOperations=True
		 'VerifyValue Case to verify paameter values
		Case "VerifyValue"	'Fn_SE_MSWordTabOperations("VerifyValue","value1","parameter1")
				'Taking no of rows from ParametricValues table
				iRowCnt=Fn_UI_Object_GetROProperty("Fn_SE_MSWordTabOperations",JavaWindow("MyTeamcenter").JavaTable("ParametricValues"),"rows")
				For iCount=0 To iRowCnt-1
						'Taking data from ParametricValues table
						sCellData=Fn_UI_JavaTable_GetCellData("Fn_SE_MSWordTabOperations",JavaWindow("MyTeamcenter"), "ParametricValues",iCount,0)
						If strParameterName=sCellData Then
								sValueName=Fn_UI_JavaTable_GetCellData("Fn_SE_MSWordTabOperations",JavaWindow("MyTeamcenter"), "ParametricValues",iCount,1)
								If strValue=sValueName Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Pass:Successfully verify  value"& strValue )   
									bReturn=True
									Exit For
								End If
						End If
				Next
				If bReturn=False Then
                    Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Fail:"& strParameterName & "Parameter is not present")   
					Exit Function
				Else
					  Fn_SE_MSWordTabOperations=True
				End If
		Case "VerifyInStr"		
				'Getting Value in MSWord text box
				sValueName = Fn_Edit_Box_GetValue("Fn_SE_MSWordTabOperations",JavaWindow("MyTeamcenter"),"MSWordTabText")
				If Instr(1, sValueName, strValue)<>0 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Pass:"& strValue &"Successfully found")
					Fn_SE_MSWordTabOperations=True
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Fail:"& strValue &"Not Found.")
					Fn_SE_MSWordTabOperations=False
				End If		
		Case "SetValueWithoutSave"		'Fn_SE_MSWordTabOperations("SetValue","[parameter1:value1,value2,value3..]"+VbCrlf+"[parameter2: value1,value2,value3..]","")
				'Setting Value in text box
                Call Fn_Edit_Box("Fn_SE_MSWordTabOperations",JavaWindow("MyTeamcenter"),"MSWordTabText",strValue)
				'Saving the changes
     			Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Pass:Successfully Set  value"& strValue)   
				'Function Return True
				Fn_SE_MSWordTabOperations=True		
		Case "GetValue"		'Fn_SE_MSWordTabOperations("GetValue","","")
		
					sAttachedText = JavaWindow("SystemsEngineering").JavaEdit("QuickCreatePnlName").GetToProperty("attached text")
				'Saving the changes
				JavaWindow("SystemsEngineering").JavaEdit("QuickCreatePnlName").SetTOProperty "attached text","Note Text:"
				sText =  JavaWindow("SystemsEngineering").JavaEdit("QuickCreatePnlName").GetROProperty("value")
				JavaWindow("SystemsEngineering").JavaEdit("QuickCreatePnlName").SetTOProperty "attached text",sAttachedText
				Fn_SE_MSWordTabOperations=sText						
	End Select
	'Activating summary tab
	'Call Fn_MyTc_TabOperation("Activate", "Summary")
End Function

'-------------------------------------------------------------------------This Function is used to to perform operation on Input Parametric Values Window-----------------------------------------
'Function Name		:	Fn_SE_ParamatricValueOperation

'Description			:	This Function is used to to perform operation on Input Parametric Values Window

'Parameters			:			1.	strAction:Action Name
'											 2.	 strValue:Value to select from Table
											'3.	 strNoteText:

'Return Value		:	True/False

'Pre-requisite		:	 Object should be selected
'											
'Examples			:	Fn_SE_ParamatricValueOperation("SetParametricValue","value1:value2","")
'									
'									
'History:
'									Developer Name												Date								Rev. No.						Changes Done				Reviewer	Build
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'									Amol							   02-Feb-2011		              1.0										Created							Tushar B     				  20110119
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SE_ParamatricValueOperation(strAction,strValue,strNoteText)
	GBL_FAILED_FUNCTION_NAME="Fn_SE_ParamatricValueOperation"
   'Declaring variables
   Dim iCounter,strVal
   Dim objSelectType,objDialog
	'Function Return False
    Fn_SE_ParamatricValueOperation=False
	'Verifying existance of Input Parametric Values window
	If Fn_UI_ObjectExist("Fn_SE_ParamatricValueOperation",JavaWindow("MyTeamcenter").JavaWindow("Input Parametric Values"))=False Then
		'Invoking Input Parametric Values window
		Call Fn_MenuOperation("Select","Edit:Attach Requirements/Notes:Parametric Requirement")
	End If
	
	Select Case strAction
			'SetParametricValue Case set values
			Case "SetParametricValue"	'Fn_SE_ParamatricValueOperation("SetParametricValue","value1:value2","")
                    strVal=Split(strValue,":")
                    For iCounter=0 To Ubound(strVal)
						If strVal(iCounter)<>"" Then
								Set objSelectType=description.Create()
								objSelectType("Class Name").value = "JavaList"
								Set objDialog =JavaWindow("MyTeamcenter").JavaWindow("Input Parametric Values").JavaTable("Table").ChildObjects(objSelectType)
                                objDialog(iCounter).Select strVal(iCounter)
								Fn_SE_ParamatricValueOperation=True
								Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Pass:Successfully Set  value"& strVal(iCounter))
						End If
					Next				
		End Select
	 'Click on OK button
	Call Fn_Button_Click("Fn_SE_ParamatricValueOperation", JavaWindow("MyTeamcenter").JavaWindow("Input Parametric Values"), "OK")
	'Releasing objects
	Set objSelectType=Nothing
	Set objDialog =Nothing
End Function
'-------------------------------------------------------------------------------------------------------------Function to create Para Or Req from quick panel--------------------------------------------------------------------
'Function Name		:			Fn_SE_QuickPanelOperation

'Description			 :		 This function is used to create Requirement or Paragraph Through Quick Panel

'Parameters			   :	 		1.sName:Name of Requirement or Paragraph (Mandetory Parameter)
'												 2.sType: Type(Mandetory Parameter)
													'1.Requirement
													'2.Paragraph
'												 3.sChildOpt:"ON" OR "OFF"

'Return Value		   : 	True/False

'Pre-requisite			:		 Should be logged in & present on System Engineering perspective and Quick panel has to display

'Examples				:		Fn_SE_QuickPanelOperation("Test","Paragraph","OFF")
'									 Fn_SE_QuickPanelOperation("Test","Requirement","ON")
'History					 :		
'													Developer Name												Date						Rev. No.						Changes Done	
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Ketan Raje.									   				04/02/2011			              1.0								Created									
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SE_QuickPanelOperation(sName,sType,sChildOpt)
	GBL_FAILED_FUNCTION_NAME="Fn_SE_QuickPanelOperation"
	Dim ObjSEWindow
	Dim arrDetails, iCount
	Fn_SE_QuickPanelOperation=False
	If Fn_UI_ObjectExist("Fn_SE_QuickPanelOperation",JavaWindow("SystemsEngineering"))=False Then
        Exit Function
	End If
	Set ObjSEWindow=Fn_UI_ObjectCreate("Fn_SE_QuickPanelOperation",JavaWindow("SystemsEngineering"))
	'Setting Name to Req Or Para
	Call Fn_Edit_Box("Fn_SE_QuickPanelOperation",ObjSEWindow,"QuickCreatePnlName","")
	Call Fn_UI_EditBox_Type("Fn_SE_QuickPanelOperation",ObjSEWindow,"QuickCreatePnlName",sName)
	'Added code when type method not works correctly - by PoonamC_29Jun2017_NewDevelopment
	If Fn_UI_Object_GetROProperty("Fn_SE_QuickPanelOperation",ObjSEWindow.JavaEdit("QuickCreatePnlName"),"text") = ""  Then
		Call Fn_Edit_Box("Fn_SE_QuickPanelOperation",ObjSEWindow,"QuickCreatePnlName",sName)
	End If

	'Selecting Type
	Call Fn_Edit_Box("Fn_SE_QuickPanelOperation",ObjSEWindow,"Type","")
	Call Fn_UI_EditBox_Type("Fn_SE_QuickPanelOperation",ObjSEWindow,"Type",sType)
	'Added code when type method not works correctly - by PoonamC_29Jun2017_NewDevelopment
	If Fn_UI_Object_GetROProperty("Fn_SE_QuickPanelOperation",ObjSEWindow.JavaEdit("Type"),"text") = ""  Then
		Call Fn_Edit_Box("Fn_SE_QuickPanelOperation",ObjSEWindow,"Type",sType)
	End If
	
     'Call Fn_List_Select("Fn_SE_QuickPanelOperation",ObjSEWindow,"QuickCreatePnlType",sType)
    '[TC1122-20160316-30_03_2016-VivekA-Maintenance] - Added from TC1015, as index of Shell window changes at runtime.
	If Instr(sChildOpt,"Details:")>0 Then
		arrDetails=Split(sChildOpt,":")
	    JavaWindow("SystemsEngineering").JavaLink("Link").Click "5","5","LEFT"
	    For iCount = 0 To 10
			JavaWindow("SystemsEngineering").JavaWindow("Shell").SetTOProperty "index", iCount
			Wait 1
			If JavaWindow("SystemsEngineering").JavaWindow("Shell").JavaEdit("DetailsLink_Edit").Exist(1) Then
				JavaWindow("SystemsEngineering").JavaWindow("Shell").JavaEdit("DetailsLink_Edit").Set arrDetails(1)
				'Call Fn_Edit_Box("",JavaWindow("SystemsEngineering").JavaWindow("Shell"),"DetailsLink_Edit",arrDetails(1))
				Wait 1
				Exit For
			End If
		Next
	ElseIf sChildOpt<>"" Then
		'Setting Child option "ON" Or "OFF"
'		JavaWindow("SystemsEngineering").JavaCheckBox("QuickCreatePnlChild").Object.setEnabled True
		Call Fn_CheckBox_Set("Fn_SE_QuickPanelOperation",ObjSEWindow,"QuickCreatePnlChild",sChildOpt)
	End If
    'Clicking on create button to create Para Or Req
	Call Fn_Button_Click("Fn_SE_QuickPanelOperation",ObjSEWindow,"QuickCreatePnlCreate")
	Fn_SE_QuickPanelOperation=True
	Set ObjSEWindow=Nothing
End Function
'-------------------------------------------------------------------------------------------------------------Function to Customize Menu's--------------------------------------------------------------------
'Function Name		:			Fn_SE_CustomizeIWantTo

'Description			 :		 This function is used to Customize Menu's

'Parameters			   :	 		1.strAction:Action Name
													'1."Add"
													'2."Remove"
													'3.VerifyEntries
'												 2.strEntryNode: Node Name
													'1.In "Add" Case Node Name is From "Available Entries" Tree (:) Seperated
													'2.In "Remove" And "VerifyEntries" Case Node Name is From "Selected Entries" Table

											'Compulsory
'											IMP Note : To Use this function Clear the cache first and Launch New Application	
											'Fn_ReUserTcSession(True, True, Environment.Value("TcUser1"))
'Return Value		   : 	True/False

'Pre-requisite			:	Should be logged in & present on System Engineering perspective

											'IMP Note : To Use this function Clear the cache first and Launch New Application	
											'Fn_ReUserTcSession(True, True, Environment.Value("TcUser1"))

'Examples				:		Fn_SE_CustomizeIWantTo("Add","View:Show Data Panel")
'											Fn_SE_CustomizeIWantTo("Remove","Show Data Panel")
'											Fn_SE_CustomizeIWantTo("VerifyEntries","Show Data Panel")
'History					 :		
'					Developer Name					Date				Rev. No.				Changes Done		TC-Build
' ------------------------------------------------------------------------------------------------------------------------------------------------------------
' 					Ketan Raje.					04/02/2011		            1.0						Created					2011011900
'-------------------------------------------------------------------------------------------------------------------------------------------------------------
Public function Fn_SE_CustomizeIWantTo(strAction,strEntryNode)
	GBL_FAILED_FUNCTION_NAME="Fn_SE_CustomizeIWantTo"
   'Declaring Variables
	Dim ObjSEIWantTo,bFlag,sEntryNames,iCount,iItemCount,iCnt,sCellData
	'Initially Setting function to False
	Fn_SE_CustomizeIWantTo=False
  
	'Checking Existance "SEIWantTo" Window
	If Fn_UI_ObjectExist("Fn_SE_CustomizeIWantTo",JavaWindow("SystemsEngineering").JavaWindow("SEIWantTo"))=False Then
		'Changing Label of "IWantTo" static text to "History"
		Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_SE_CustomizeIWantTo",JavaWindow("SystemsEngineering").JavaStaticText("IWantTo..."),"label","History")
		'Clicking on "History"
		Call Fn_UI_JavaStaticText_Click("Fn_SE_CustomizeIWantTo", JavaWindow("SystemsEngineering"),"IWantTo...",1,1, "")

		'Changing Label of "IWantTo" static text to "Open Items"
		Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_SE_CustomizeIWantTo",JavaWindow("SystemsEngineering").JavaStaticText("IWantTo..."),"label","Open Items")
		'Clicking on "Open Items"
		Call Fn_UI_JavaStaticText_Click("Fn_SE_CustomizeIWantTo", JavaWindow("SystemsEngineering"),"IWantTo...",1,1, "")
		
		'Changing Label of "IWantTo" static text to "Favorites"
		Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_SE_CustomizeIWantTo",JavaWindow("SystemsEngineering").JavaStaticText("IWantTo..."),"label","Favorites")
		'Clicking on "Favorites"
		Call Fn_UI_JavaStaticText_Click("Fn_SE_CustomizeIWantTo", JavaWindow("SystemsEngineering"),"IWantTo...",1,1, "")
		'Clicking on Customize Toolbar button to Open "SEIWantTo" Window
		Call Fn_ToolbarButtonClick_Ext(2,"Customize")
	End If
	'Creating Object "SEIWantTo" Window
	Set ObjSEIWantTo=Fn_UI_ObjectCreate("Fn_SE_CustomizeIWantTo",JavaWindow("SystemsEngineering").JavaWindow("SEIWantTo"))
   Select Case strAction
    Case "Add" 'Case to Add Entries
		'Selecting Item from Available Entries Tree
		Call Fn_JavaTree_Select("Fn_SE_CustomizeIWantTo", ObjSEIWantTo,"Available Entries",strEntryNode)
		'Click on Plus button to Add Entry
        Call Fn_Button_Click("Fn_SE_CustomizeIWantTo", ObjSEIWantTo, "Plus")
	Case "Remove"  'Case to Remove Entries
		'Spliting Node Name
		sEntryNames=Split(strEntryNode,":")
		For iCount=0 To Ubound(sEntryNames)
			bFlag=False
			'Taking Total item Count of Selected Entries Table
            iItemCount=Fn_UI_Object_GetROProperty("Fn_SE_CustomizeIWantTo",ObjSEIWantTo.JavaTable("Selected Entries"),"rows")
			For iCnt=0 To iItemCount-1
				'Taking Data from Selected Entries Table
				sCellData=Fn_UI_JavaTable_GetCellData("Fn_SE_CustomizeIWantTo",ObjSEIWantTo, "Selected Entries",iCnt,0)
				'Checking Entry with table data
				If  sEntryNames(iCount)=sCellData Then
					'Selecting Data from table
                    Call Fn_UI_JavaTable_SelectCell("Fn_SE_CustomizeIWantTo", ObjSEIWantTo, "Selected Entries",iCnt,0)
					'Clicking on Remove button to remove Entry
					Call Fn_Button_Click("Fn_SE_CustomizeIWantTo", ObjSEIWantTo, "Minus")
					bFlag=True
					Exit For					
				End If
			Next
				If bFlag=False Then
					'If data is not present in table then exit the function
					Call Fn_Button_Click("Fn_SE_CustomizeIWantTo", ObjSEIWantTo, "OK")
					Set ObjSEIWantTo=Nothing
					Exit Function
				End If
		Next
	
	Case "VerifyEntries"
		sEntryNames=Split(strEntryNode,":")
		For iCount=0 To Ubound(sEntryNames)
			bFlag=False
			'Taking Total item Count of Selected Entries Table
            iItemCount=Fn_UI_Object_GetROProperty("Fn_SE_CustomizeIWantTo",ObjSEIWantTo.JavaTable("Selected Entries"),"rows")
			For iCnt=0 To iItemCount-1
				'Taking Data from Selected Entries Table
				sCellData=Fn_UI_JavaTable_GetCellData("Fn_SE_CustomizeIWantTo",ObjSEIWantTo, "Selected Entries",iCnt,0)
				If  sEntryNames(iCount)=sCellData Then
                    bFlag=True
					Exit For					
				End If
			Next
		Next
		If bFlag=False Then
			'If data is not present in table then exit the function
			Call Fn_Button_Click("Fn_SE_CustomizeIWantTo", ObjSEIWantTo, "OK")
			Set ObjSEIWantTo=Nothing
			Exit Function
		End If
   End Select
   'Setting Function to True
   Fn_SE_CustomizeIWantTo=True
   'Clicking on OK button 
  Call Fn_Button_Click("Fn_SE_CustomizeIWantTo", ObjSEIWantTo, "OK")
  'Setting object to Nothing
  Set ObjSEIWantTo=Nothing
End Function
'*********************************************************		Function to perform operations on TraceabilityReport Table Column*************************************
'Function Name		:			Fn_SE_TraceabilityReportColumnOperations

'Description			 :		 	  Function to perform operations on TraceabilityReport Table Column

'Parameters			   :	 			sTableName,sAction,sColName,sCategoryAndType,bUseDisplayableName,sButton
															
'Return Value		   : 			True /False

'Pre-requisite			:		 	Traceability Report Dialog should be open

'Examples				:			'MsgBox Fn_SE_TraceabilityReportColumnOperations("Defining Traceability Report","RemoveColumn","Name","","","")
												'MsgBox Fn_SE_TraceabilityReportColumnOperations("Defining Traceability Report","InsertColumn","Name","","","")
												'MsgBox Fn_SE_TraceabilityReportColumnOperations("Defining Traceability Report","ColumnExist","Name","","","")
'History					 :			
'													Developer Name								Date						Rev. No.						Changes Done						Build
'												------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Harshal Agrawal						02-Feb-2011					1.0										Developed								20110119
'													Sandeep N								29-Jul-2011					1.1										Modify All Cases						20100707
'													Vivek A								17-July-2015				1.1				 TC112-2015070100-17_07_2015-Porting-VivekA-Added check for new object heirarchy from TC1014				
'												------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SE_TraceabilityReportColumnOperations(sTraceabilityReport,sAction,sColName,sCategoryAndType,bUseDisplayableName,sButton)
	GBL_FAILED_FUNCTION_NAME="Fn_SE_TraceabilityReportColumnOperations"
   'Declaring Variables
	Dim objTable,iAppColCount,iCount,sAppColName,objChangeCol,sHeader,objReport,objTcApplet
	Dim bFlag
	bFlag =  True
	'Function Returning False
	Fn_SE_TraceabilityReportColumnOperations=False
   Select Case sTraceabilityReport
	 	Case "Defining Traceability Report"		   
		   If JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Traceability Report").JavaTable("DefiningTable").Exist(6) Then
				Set objTable = JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Traceability Report").JavaTable("DefiningTable")
				Set objReport=JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Traceability Report")
				Set objTcApplet=JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet")
			ElseIf JavaWindow("Traceability").JavaTable("DefiningTable").Exist(5) Then
                Set objTable = JavaWindow("Traceability").JavaTable("DefiningTable")
				Set objReport=JavaWindow("Traceability")
				Set objTcApplet=JavaWindow("Traceability")
			ElseIf JavaWindow("Defining Traceability").JavaTable("TraceabilityReportPanel").Exist(5) Then
				Set objTable = JavaWindow("Defining Traceability").JavaTable("TraceabilityReportPanel")
				Set objReport=JavaWindow("Defining Traceability")
				Set objTcApplet=JavaWindow("Defining Traceability")
		    End If

		Case "Complying Traceability Report"
			If JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Traceability Report").JavaTable("ComplyingTable").Exist(6) Then
				Set objTable = JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Traceability Report").JavaTable("ComplyingTable")
				Set objReport=JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Traceability Report")
				Set objTcApplet=JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet")
			ElseIf JavaWindow("Traceability").JavaTable("ComplyingTable").Exist(5) Then
                Set objTable = JavaWindow("Traceability").JavaTable("ComplyingTable")
				Set objReport=JavaWindow("Traceability")
				Set objTcApplet=JavaWindow("Traceability")
			ElseIf JavaWindow("Defining Traceability").JavaTable("TraceabilityReportPanel").Exist(5) Then
				Set objTable = JavaWindow("Defining Traceability").JavaTable("TraceabilityReportPanel")
				Set objReport=JavaWindow("Defining Traceability")
				Set objTcApplet=JavaWindow("Defining Traceability")
			ElseIf JavaWindow("SystemsEngineering").JavaWindow("WEmbeddedFrame").JavaTable("ComplyingObject").Exist(5) Then ' TC112-2015070100-17_07_2015-Porting-VivekA-Added check for new object heirarchy from TC1014
				Set objTable = JavaWindow("SystemsEngineering").JavaWindow("WEmbeddedFrame").JavaTable("ComplyingObject")
				Set objReport=JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame")
				'Set objTcApplet=JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame")
				Set objTcApplet=JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet")
				bFlag =  False
		    End If

		Case Else
			Exit Function
   End Select
   Select Case sAction
		Case  "RemoveColumn"
            		objTable.SelectColumnHeader sColName,"RIGHT"
					objReport.JavaMenu("label:=Remove this column").Select
                	'Clicking on yes button to confirm remove of column
					Call Fn_Button_Click("Fn_SE_TraceabilityReportColumnOperations",JavaWindow("Traceability").JavaDialog("RemoveColumn"),"Yes")
                	Fn_SE_TraceabilityReportColumnOperations=True
        Case  "InsertColumn"
					'Selecting Insert column\(s\) menu
						iAppColCount = objTable.GetROProperty("cols")
						iAppColCount = iAppColCount - 1
					 sHeader = objTable.GetColumnName(iAppColCount)
					objTable.SelectColumnHeader sHeader,"RIGHT"
					objReport.JavaMenu("label:=Insert column\(s\) \.\.\.").Select
					If objTcApplet.JavaDialog("ChangeColumns").Exist Then
						Set objChangeCol = objTcApplet.JavaDialog("ChangeColumns")
						If bUseDisplayableName<>"" Then
                    		objChangeCol.JavaCheckBox("UseDisplayableName").Set bUseDisplayableName
						End If
						If sCategoryAndType<>""  Then
							objChangeCol.JavaTree("CategoryAndType").Select sCategoryAndType
						End If
						objChangeCol.JavaList("AvailableCol").Select sColName
						Wait 1
						If objChangeCol.JavaButton("Add").GetROProperty("enabled") <> "1" Then
							objChangeCol.JavaEdit("Available Col").Type sColName
							Wait 1
						End If
						objChangeCol.JavaButton("Add").Click micLeftBtn
						objChangeCol.JavaButton("Apply").Click micLeftBtn
						If objChangeCol.JavaButton("Cancel").Exist(4) Then
							objChangeCol.JavaButton("Cancel").Click micLeftBtn
						Else
							objChangeCol.JavaButton("Close").Click micLeftBtn
						End If
						Fn_SE_TraceabilityReportColumnOperations = True
                End If
	Case "ColumnExist"
			iAppColCount = objTable.GetROProperty("cols")
			For iCount = 0 to cint(iAppColCount)-1
				sAppColName = objTable.GetColumnName(iCount)
				If trim(lcase(sAppColName)) = trim(lcase(sColName)) Then
						Fn_SE_TraceabilityReportColumnOperations =  True
						Exit For
				End If
			Next
    Case  Else
		Exit Function
	End Select
	'Closing TraceabilityReport Report
	If bFlag <> False Then
		Call Fn_Button_Click("Fn_SE_TraceabilityReportColumnOperations",objReport,"Ok")
	End If
	Set objTable = Nothing
	Set objChangeCol = Nothing
End Function
''*********************************************************		Function to action perform on NavTree	***********************************************************************
'Function Name		:				Fn_SE_NavTree_NodeOperation

'Description			 :		 		 Actions performed in this function are:
'																	1. Node Select
'																	2. Node multi-select
'																	3. Node Expand
'																	4. Node Collapse
'																	5. Node Popup menu select
'																	6. Node double-click
'																	7. Node MultiSelect Cntxt Menu
'																	8. Node Exist
'																	9. MultiSelectContextMenuExist

'Parameters			   :	 			1. StrAction: Action to be performed
'													2. StrNodeName: Fully qulified tree Path (delimiter as ':') [multiple node are separated by "," ] 
'												   3. StrMenu: Context menu to be selected

'Return Value		   : 				TRUE \ FALSE

'Pre-requisite			:		 		System Engineering module window should be displayed

'Examples				:				  Fn_SE_NavTree_NodeOperation("PopupMenuSelect","Home:Newstuff","Copy Ctrl+C")
'													EXAMPLE for Case "Select" : Call Fn_SE_NavTree_NodeOperation( "Select" ,  "Home:Newstuff:000032-CarModel_VI_LS1:000032 @2" , "" ) 
' 												   Call Fn_SE_NavTree_NodeOperation( "Select" ,  "Home:Newstuff:000032-CarModel_VI_LS1:000032" , "" ) 
'													EXAMPLE for Case "GetSelected"::  Fn_SE_NavTree_NodeOperation( "GetSelected" , "Home:Mailbox,Home:Newstuff,Home:000039-1,Home:Kavan_Shah" , "" ) 
'													EXAMPLE for Case "GetChildItemCount"::  Fn_SE_NavTree_NodeOperation( "GetChildItemCount" , "Home:Mailbox" , "" ) 
'													EXAMPLE for Case "GetChildInstances"::  Fn_SE_NavTree_NodeOperation("GetChildInstances","Home:AutomatedTests:sonal:000112-top:000112/A;1-top:View:000114-sub2","") 		Added By Ketan On 11-Jan-2011
'History					 :		
'									Developer Name				Date						Rev. No.			Changes Done			Reviewer
'								-------------------------------------------------------------------------------------------------------------------------------
'									Ketan Raje				  08/02/2011			            1.0				   Created						Harshal
'									Sandeep N				16/11/2011			            1.1				   Created						  Sunny R
'								-------------------------------------------------------------------------------------------------------------------------------


Function Fn_SE_NavTree_NodeOperation(StrAction,StrNodeName,StrMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_SE_NavTree_NodeOperation"
	Dim NodeLists, intNodeCount, intCount, StrExist, aMenuList, sTreeItem,sCmpItm
	Dim objJavaWindowSE, objJavaTreeNav,ArrNodeName
	Dim ArrStrcomp, sArrStr1,sArrStr2, iCounter
	Dim iRows, colonCnt,arriPath,iVal,oCurrentNode,echStrNode
	Dim iItemCount, aNodePath,  iInstance, instCount, aNodes
	Dim sPath, sEle ,iCnt, bFound,sNodePath,sSrcPath,sDestPath
	Set objJavaWindowSE = Fn_UI_ObjectCreate( "Fn_SE_NavTree_NodeOperation",JavaWindow("SystemsEngineering"))
	Select Case StrAction
		'----------------------------------------------------------------------- For selecting single node -------------------------------------------------------------------------
		Case "Select"
				'--- For selecting single node without instance ID--
					Set objJavaTreeNav = Fn_UI_ObjectCreate( "Fn_SE_NavTree_NodeOperation", JavaWindow("SystemsEngineering").JavaTree("NavTree"))
					sNodePath = Fn_UI_JavaTreeGetItemPathExt("Fn_SE_NavTree_NodeOperation", objJavaTreeNav, StrNodeName, "", "")
					objJavaTreeNav.Select sNodePath
					Fn_SE_NavTree_NodeOperation = TRUE
			
		'----------------------------------------------------------------------- For selecting single node -------------------------------------------------------------------------
		Case "Deselect"
			Set objJavaTreeNav = Fn_UI_ObjectCreate( "Fn_SE_NavTree_NodeOperation", JavaWindow("SystemsEngineering").JavaTree("NavTree"))
			 sNodePath = Fn_UI_JavaTreeGetItemPathExt("Fn_SE_NavTree_NodeOperation", objJavaTreeNav, StrNodeName, "", "")
			JavaWindow("SystemsEngineering").JavaTree("NavTree").Deselect sNodePath
			If Err.Number < 0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Deselect Node [" + StrNodeName + "] of NavTree")
				Fn_SE_NavTree_NodeOperation = False
			End If
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Deselected Node [" + StrNodeName + "] of NavTree")
			Fn_SE_NavTree_NodeOperation = True
		'----------------------------------------------------------------------- For selecting multiple node at a time -------------------------------------------------------------------------
		Case "Multiselect"
			'Split the string where "'," exist
			Set objJavaTreeNav = Fn_UI_ObjectCreate( "Fn_SE_NavTree_NodeOperation", JavaWindow("SystemsEngineering").JavaTree("NavTree"))
					NodeLists=Split(StrNodeName,",")
					For iCounter=0 To UBound(NodeLists)
						sNodePath = Fn_UI_JavaTreeGetItemPathExt("Fn_SE_NavTree_NodeOperation", objJavaTreeNav, NodeLists(iCounter), "", "")
						objJavaTreeNav.ExtendSelect sNodePath
					Next
					Fn_SE_NavTree_NodeOperation = TRUE
		'----------------------------------------------------------------------- For expanding a particular  node-------------------------------------------------------------------------
		Case "Expand"
			Set objJavaTreeNav = Fn_UI_ObjectCreate( "Fn_SE_NavTree_NodeOperation", JavaWindow("SystemsEngineering").JavaTree("NavTree"))
					sNodePath = Fn_UI_JavaTreeGetItemPathExt("Fn_SE_NavTree_NodeOperation", objJavaTreeNav, StrNodeName, "", "")
                    Call Fn_UI_JavaTree_Expand("Fn_SE_NavTree_NodeOperation",objJavaWindowSE,"NavTree",sNodePath)
					Fn_SE_NavTree_NodeOperation = TRUE
		'----------------------------------------------------------------------- For collapssing a particular  node-------------------------------------------------------------------------
		Case "Collapse"
			Set objJavaTreeNav = Fn_UI_ObjectCreate( "Fn_SE_NavTree_NodeOperation", JavaWindow("SystemsEngineering").JavaTree("NavTree"))
					sNodePath = Fn_UI_JavaTreeGetItemPathExt("Fn_SE_NavTree_NodeOperation", objJavaTreeNav, StrNodeName, "", "")
                    Call Fn_UI_JavaTree_Collapse("Fn_SE_NavTree_NodeOperation", objJavaWindowSE,"NavTree",sNodePath)
					Fn_SE_NavTree_NodeOperation = TRUE
		'----------------------------------------------------------------------- For selecting popup menu of  a particular  node-------------------------------------------------------------------------
		Case "PopupMenuSelect"
			Set objJavaTreeNav = Fn_UI_ObjectCreate( "Fn_SE_NavTree_NodeOperation", JavaWindow("SystemsEngineering").JavaTree("NavTree"))
					'Build the Popup menu to be selected
					aMenuList = split(StrMenu, ":",-1,1)
					intCount = Ubound(aMenuList)
					'Select node
					sNodePath = Fn_UI_JavaTreeGetItemPathExt("Fn_SE_NavTree_NodeOperation", objJavaTreeNav, StrNodeName, "", "")
                    Call Fn_JavaTree_Select("Fn_SE_NavTree_NodeOperation",objJavaWindowSE,"NavTree",sNodePath)
					'Open context menu
					Call Fn_UI_JavaTree_OpenContextMenu("Fn_SE_NavTree_NodeOperation",objJavaWindowSE,"NavTree",sNodePath)
					'Select Menu action
					Select Case intCount
						Case "0"
							 StrMenu = JavaWindow("SystemsEngineering").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0))
							'Implementation
						Case "1"
							StrMenu = JavaWindow("SystemsEngineering").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1))
							'Implementation
						Case "2"
							StrMenu = JavaWindow("SystemsEngineering").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1),aMenuList(2))
							'Implementation
						Case Else
							Fn_SE_NavTree_NodeOperation = FALSE
							Exit Function
					End Select
					JavaWindow("SystemsEngineering").WinMenu("ContextMenu").Select StrMenu
					'Implementation
					Fn_SE_NavTree_NodeOperation = TRUE
        '----------------------------------------------------------------------- For Checking Existing in Context menu for Multiselected Item node------------------------------------------------
		Case "MultiSelectContextMenuExist"
				NodeLists = Split(StrNodeName,",")
				Set objJavaTreeNav = Fn_UI_ObjectCreate( "Fn_SE_NavTree_NodeOperation", JavaWindow("SystemsEngineering").JavaTree("NavTree"))
				Call Fn_SE_NavTree_NodeOperation("Multiselect", StrNodeName, "")
				sNodePath = Fn_UI_JavaTreeGetItemPath(objJavaTreeNav,NodeLists(0))
				Call Fn_UI_JavaTree_OpenContextMenu("Fn_SE_NavTree_NodeOperation",objJavaWindowSE,"NavTree",sNodePath)
				If JavaWindow("SystemsEngineering").WinMenu("ContextMenu").GetItemProperty (StrMenu,"Exists") = True Then
					Fn_SE_NavTree_NodeOperation = TRUE
				Else
					Fn_SE_NavTree_NodeOperation = FALSE
			  	End If
		'----------------------------------------------------------------------- For doble clicking on a particular  node-------------------------------------------------------------------------
		Case "DoubleClick"
'			Set objJavaTreeNav = Fn_UI_ObjectCreate( "Fn_SE_NavTree_NodeOperation", JavaWindow("SystemsEngineering").JavaTree("NavTree"))
'					sNodePath = Fn_UI_JavaTreeGetItemPath(objJavaTreeNav,StrNodeName)
'					JavaWindow("SystemsEngineering").JavaTree("NavTree").Activate sNodePath
					Call Fn_SE_NavTree_NodeOperation( "Select" ,StrNodeName, "" )
                    wait 1					
					Call Fn_KeyBoardOperation("SendKeys", "{ENTER}")
					wait 2
					Fn_SE_NavTree_NodeOperation = TRUE
		'----------------------------------------------------------------------- For doble clicking on a particular  node-------------------------------------------------------------------------
		Case "MultiSelectContextMenu"
			NodeLists = Split(StrNodeName,",")
			aMenuList = split(StrMenu, ":",-1,1)
			intCount = Ubound(aMenuList)
			Set objJavaTreeNav = Fn_UI_ObjectCreate( "Fn_SE_NavTree_NodeOperation", JavaWindow("SystemsEngineering").JavaTree("NavTree"))
					'Select multiple node
					Call Fn_SE_NavTree_NodeOperation("Multiselect", StrNodeName, "")
					'Open context menu
					sNodePath = Fn_UI_JavaTreeGetItemPath(objJavaTreeNav,NodeLists(0))
                	Call Fn_UI_JavaTree_OpenContextMenu("Fn_SE_NavTree_NodeOperation",objJavaWindowSE,"NavTree",sNodePath)
					Select Case intCount
						Case "0"
							 StrMenu = JavaWindow("SystemsEngineering").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0))
						Case "1"
							StrMenu = JavaWindow("SystemsEngineering").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1))
						Case "2"
							StrMenu = JavaWindow("SystemsEngineering").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1),aMenuList(2))
							'Implementation
						Case Else
							Fn_SE_NavTree_NodeOperation = FALSE
							Exit Function
					End Select
					JavaWindow("SystemsEngineering").WinMenu("ContextMenu").Select StrMenu
					'Select Menu action
					Fn_SE_NavTree_NodeOperation = TRUE
		'----------------------------------------------------------------------- For doble clicking on a particular  node-------------------------------------------------------------------------
		Case "Exist"
					sNodePath = Fn_UI_JavaTreeGetItemPath(JavaWindow("SystemsEngineering").JavaTree("NavTree"),StrNodeName)
					If sNodePath = False Then
						Fn_SE_NavTree_NodeOperation = FALSE
					Else
						Fn_SE_NavTree_NodeOperation = TRUE
					End If
		Case "PopupMenuExist"
			aMenuList = split(StrMenu, ":",-1,1)
					intCount = Ubound(aMenuList)
					Set objJavaTreeNav = Fn_UI_ObjectCreate( "Fn_SE_NavTree_NodeOperation", JavaWindow("SystemsEngineering").JavaTree("NavTree"))
					Call Fn_SE_NavTree_NodeOperation("Select", StrNodeName, "")
					'Open context menu
					sNodePath = Fn_UI_JavaTreeGetItemPath(objJavaTreeNav,StrNodeName)
                	Call Fn_UI_JavaTree_OpenContextMenu("Fn_SE_NavTree_NodeOperation",objJavaWindowSE,"NavTree",sNodePath)
					Select Case intCount
						Case "0"
							 StrMenu = JavaWindow("SystemsEngineering").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0))
                        Case "1"
							StrMenu = JavaWindow("SystemsEngineering").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1))
                        Case "2"
							StrMenu = JavaWindow("SystemsEngineering").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1),aMenuList(2))
                        Case Else
							Fn_SE_NavTree_NodeOperation = FALSE
                        Exit Function
					End Select
				If JavaWindow("SystemsEngineering").WinMenu("ContextMenu").GetItemProperty (StrMenu,"Exists") = True Then
					Fn_SE_NavTree_NodeOperation = TRUE
				Else
					Fn_SE_NavTree_NodeOperation = FALSE
			  	End If
		'------------------- Checks That item is inactively focused Or Not for single node OR Multiple Node(comma "," SEPERATED)---------------
		Case "GetSelected"
		wait 5
			Set objJavaTreeNav = Fn_UI_ObjectCreate( "Fn_SE_NavTree_NodeOperation", JavaWindow("SystemsEngineering").JavaTree("NavTree"))	
				
				ArrStrcomp = Split(objJavaTreeNav.GetROProperty("value") ,"",-1,1)
				sArrStr2 = ArrStrcomp(0)
				For iCounter = 1 To ubound(ArrStrcomp)
					sArrStr2 = sArrStr2 & "," & ArrStrcomp(iCounter)
				Next
				If sArrStr2 = StrNodeName Then
				   'Writing Log
				   Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Java Tree Multiple Node ["+StrNodeName+"] is Selected .")
				   Fn_SE_NavTree_NodeOperation = TRUE
				Else
				   'Writing Log
				   Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Java Tree Multiple  Node ["+StrNodeName+"] is Not Selected .")
				   Fn_SE_NavTree_NodeOperation = FALSE
			End If

		Case "GetIndex"
			'Index of Item1
			ArrNodeName=Split(StrNodeName,":")
			If UBound(ArrNodeName)=0 And Lcase(ArrNodeName(0))="home" Then
				Fn_SE_NavTree_NodeOperation=0
			Else
				sNodePath = Fn_UI_JavaTreeGetItemPathExt("Fn_SE_NavTree_NodeOperation", objJavaTreeNav, StrNodeName, "", "")
				sCmpItm=Replace(sNodePath,"#","")
				sCmpItm=Replace(sCmpItm,"0","1")
				arriPath=Split(sCmpItm,":")
				iVal=0
				For iCounter=0 To UBound(arriPath)
					iVal=iVal+CInt(arriPath(iCounter))
				Next
				If sNodePath=False Then
					Fn_SE_NavTree_NodeOperation=False
				Else
					Fn_SE_NavTree_NodeOperation=iVal
				End If
			End If

		Case "GetChildItemCount"
				If Fn_SE_NavTree_NodeOperation("Expand",StrNodeName,"")=True Then
					ArrNodeName = Split (StrNodeName, ":")
					Set oCurrentNode = JavaWindow("SystemsEngineering").JavaTree("NavTree").Object.getItem(0)
					intNodeCount=0
					For each echStrNode In ArrNodeName
						iRows = oCurrentNode.getItemCount()
						For iCounter = 0 to iRows - 1
							If oCurrentNode.getItem(iCounter).getData().toString() = echStrNode Then
								intNodeCount = oCurrentNode.getItem(iCounter).getItemCount()
								Exit For
							End If
						Next
					Next 
					Fn_SE_NavTree_NodeOperation = intNodeCount
					Set oCurrentNode=Nothing
				Else
					Fn_SE_NavTree_NodeOperation = False
				End If

		Case "SelectRange"
			ReDim ArrNodeName(2)
					ArrNodeName = Split(StrNodeName,"|")
					sSrcPath =  Fn_UI_JavaTreeGetItemPath(JavaWindow("SystemsEngineering").JavaTree("NavTree"),ArrNodeName(0))
					sDestPath =  Fn_UI_JavaTreeGetItemPath(JavaWindow("SystemsEngineering").JavaTree("NavTree"),ArrNodeName(1))
					JavaWindow("SystemsEngineering").JavaTree("NavTree").SelectRange sSrcPath,sDestPath
					Fn_SE_NavTree_NodeOperation = True

		Case "GetChildInstances"
               	iInstance = 0
				aNodePath = Split(StrNodeName,":",-1,1)
				sPath = ""
				For intCount = 0 to Ubound(aNodePath)-1
					If sPath = "" Then
						sPath = aNodePath(intCount)
					Else
						sPath = sPath & ":" & aNodePath(intCount)
					End If
				Next 
				'Get Index of Parent Node
				iCnt = Fn_SE_NavTree_NodeOperation( "GetIndex" , sPath , "")
				iItemCount = Fn_SE_NavTree_NodeOperation( "GetChildrenList" , sPath , "" )
				For iCounter=0 To UBound(iItemCount)
					If Trim(iItemCount(iCounter))=Trim(aNodePath( UBound(aNodePath))) Then
						iInstance = iInstance+1
					End If
				Next
				Fn_SE_NavTree_NodeOperation = iInstance
				
				Case "ExpandAndSelect"
					'Initial Item Path
					aStrNode = Split (StrNodeName, ":")
					For i = 0 to UBound(aStrNode)-1
						If sParentPath = "" Then
							sParentPath  = aStrNode(i)
						Else
							sParentPath  = sParentPath + ":" + aStrNode(i)
						End If
						Call Fn_SE_NavTree_NodeOperation("Expand", sParentPath, "")

					Next
				
					call Fn_SE_NavTree_NodeOperation("Select",strNodeName,"")
					Fn_SE_NavTree_NodeOperation = True
				
		'****************************************************************************************	
		Case Else
				Fn_SE_NavTree_NodeOperation = FALSE
	End Select
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), StrAction &" Sucessfully completed on Node [" + StrNodeName + "] of JavaTree of function Fn_SE_NavTree_NodeOperation")
	Set objJavaWindowSE = nothing
	Set objJavaTreeNav = nothing
End Function

'-------------------------------------------------------------------------This Function is used to to perform operation on IDetails Table In RM----------------------------------------------------
'Function Name		:	Fn_SE_DetailTableOperation

'Description			:	This Function is used to to perform operation on IDetails Table In RM

'Parameters			:	   1.) sAction: Action string to navigate to appropriate case
'									    2.) sObjectName: Name of the object under Details Table
'  										3.) sColumnName: Name of the column under Details Table
'										4.) sExpectedValue: Expected value of the object property under Details Table

'Return Value		:	True/False/ColumnCount

'Pre-requisite		:	 Should be present on RM perspective And If details table is open then Good
'											
'Examples			:'MsgBox Fn_SE_DetailTableOperation("ColumnCount", "", "", "","")
								'MsgBox Fn_SE_DetailTableOperation("Rowmultiselect","001466-Spec1:001546-Spec1", "", "","")
								'MsgBox Fn_SE_DetailTableOperation("PopUpMenuSelect","", "", "","Apply Column Configuration...")
								'MsgBox Fn_SE_DetailTableOperation("GetCellData","1", "0", "","")
								'MsgBox Fn_SE_DetailTableOperation("GetIndex","001466-Spec1", "", "","")
								'MsgBox Fn_SE_DetailTableOperation("VerifyCell","001466-Spec1", "Relation", "Contents","")
								'Msgbox Fn_SE_DetailTableOperation("Rowcellupdate", "AutomatedTests", "Description", "Abxy","")
								'MsgBox Fn_SE_DetailTableOperation("MultiRowPopupMenuSelect","AutomatedTests:Newstuff", "", "","View Properties	Alt+Enter")
								'Msgbox Fn_SE_DetailTableOperation("AllColumnNames", "", "", "","")
'History:
'									Developer Name												Date								Rev. No.						Changes Done										Build	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'								Harshal Agrawal											09/Feb/2011								1.0																											20110119
'								Sandeep Nl													13/Ocy/2011								1.1								Added Case "AllColumnNames"				20110119
'								Vrushali Wani                                             30/April/2013                              1.0                              Added Case "RowSelect" 
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function Fn_SE_DetailTableOperation(sAction, sObjectName, sColumnName, sExpectedValue,sPopUpMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_SE_DetailTableOperation"
	Dim bReturn,oCounter,aObjList,intItemCount,sText,iCounter,iRowCnt,sObjName,sCellValue
	Dim ObjDetailsTable, strName, TabWidth, aMenuList, StrMenu, intCount, WshShell
	Fn_SE_DetailTableOperation=False
	'Verifying Existance of Details Table
	Set ObjDetailsTable=Fn_UI_ObjectCreate("Fn_SE_DetailTableOperation",JavaWindow("SystemsEngineering").JavaTable("DetailsTable"))
	If Fn_UI_Object_GetROProperty("Fn_SE_DetailTableOperation",ObjDetailsTable,"rows") = 0 Then
		Call Fn_UI_Object_SetTOProperty("Fn_SE_DetailTableOperation",ObjDetailsTable,"index","1")
	End If
	Select Case sAction
		 Case "ColumnCount"		'Return Number of columns currently present in Details Table
					'Returning Number of columns present in Details Table
					Fn_SE_DetailTableOperation =Fn_UI_Object_GetROProperty("Fn_SE_DetailTableOperation",ObjDetailsTable,"cols")			
		 Case "Rowmultiselect"
					'Split the string where " : " exist
					aObjList = Split(sObjectName,":")
					intItemCount =ubound(aObjList)
					'Count number of rows of Table
					bReturn=Fn_UI_Object_GetROProperty("Fn_SE_DetailTableOperation",ObjDetailsTable,"rows")
					'Extract the index of row at which the object exist.
					For oCounter=0 to intItemCount
							For iCounter=0 to bReturn-1
							sText = Fn_SISW_SE_DetailsTable_GetCellData(ObjDetailsTable,iCounter, "Object")					
							If IsNumeric(aObjList(oCounter)) Then
								 If cstr(sText) = cstr(cint(aObjList(oCounter)))  Then
									 ObjDetailsTable.ClickCell iCounter, "Object","LEFT","CONTROL"
									 Exit for
								End If
							ElseIf cstr(sText) = cstr(aObjList(oCounter))  Then
									 ObjDetailsTable.ClickCell iCounter, "Object","LEFT","CONTROL"
									 Fn_SE_DetailTableOperation=True
									 Exit for
							End If									
							Next
					Next
		Case "RowSelect"
					Fn_SE_DetailTableOperation = Fn_SISW_UI_JavaTable_Operations("Fn_SE_DetailTableOperation", "SelectRow", ObjDetailsTable , "", "GetProperty", "Object", sObjectName, "", "", "", "")
		Case "SelectCell"
					Fn_SE_DetailTableOperation = Fn_SISW_UI_JavaTable_Operations("Fn_SE_DetailTableOperation", "SelectCell", ObjDetailsTable , "", "", sColumnName, sObjectName, "", "", "", "")
		Case "SelectAllRows"
                    'Count number of rows of Table
					bReturn=Fn_UI_Object_GetROProperty("Fn_SE_DetailTableOperation",ObjDetailsTable,"rows")
                    For iCounter=0 to bReturn-1
                    	 ObjDetailsTable.ClickCell iCounter, "Object","LEFT","CONTROL"
					Next
					 Fn_SE_DetailTableOperation=True
		   Case "PopUpMenuSelect"
'					Call Fn_ToolbatButtonClick("View Menu")
					Call Fn_ToolbarButtonClick_Ext(3,"View Menu")
					JavaWindow("SystemsEngineering").WinMenu("ContextMenu").Select sPopUpMenu
					 Fn_SE_DetailTableOperation=True
		Case "GetCellData" '("GetCellData",1,0,"",")					
						JavaWindow("SystemsEngineering").JavaTable("DetailsTable").SelectRow sObjectName
						strName=JavaWindow("SystemsEngineering").JavaTable("DetailsTable").GetCellData(sObjectName,sColumnName)						
					If Err.number < 0 Then
                		Fn_SE_DetailTableOperation=False
					Else
						Fn_SE_DetailTableOperation =MId(strName,instr(1,strName,":")+1 , Len(strName))
					End If
		Case "GetCellData_Ext" 
					Fn_SE_DetailTableOperation = Fn_SISW_UI_JavaTable_Operations("Fn_SE_DetailTableOperation", "GetCellData", ObjDetailsTable , "", "GetProperty", "Object", sObjectName, sColumnName,"", "", "")
		Case "GetIndex" '("GetIndex",1,0,"","")
					Fn_SE_DetailTableOperation = Fn_SISW_UI_JavaTable_Operations("Fn_SE_DetailTableOperation", "GetRowIndex", ObjDetailsTable , "", "GetProperty", "Object", sObjectName, "", "", "", "")
		 Case "VerifyCell"
					Fn_SE_DetailTableOperation = Fn_SISW_UI_JavaTable_Operations("Fn_SE_DetailTableOperation", "VerifyCellData", ObjDetailsTable , "", "GetProperty", "Object", sObjectName, sColumnName, sExpectedValue, "", "")

   		 Case "Rowcellupdate"
					'Count number of rows of Table
					bReturn = objDetailsTable.GetROProperty("rows")	
					'Extract the index of row of which relation is to be changed
					For iCounter=0 to bReturn - 1
							'sText = objDetailsTable.GetCellData(iCounter,"Object")
							sText = Fn_SISW_SE_DetailsTable_GetCellData(ObjDetailsTable,iCounter, "Object")							
							If IsNumeric(sObjectName) Then
								 If cstr(sText) = cstr(cint(sObjectName)) Then
										TabWidth = JavaWindow("SystemsEngineering").JavaObject("SEComponentTab").GetROProperty("width")
										JavaWindow("SystemsEngineering").JavaObject("SEComponentTab").Click TabWidth-5,5,"LEFT"
										objDetailsTable.SelectCell iCounter,sColumnName
										JavaWindow("SystemsEngineering").JavaEdit("Text").Type sExpectedValue
										TabWidth = JavaWindow("SystemsEngineering").JavaObject("SEComponentTab").GetROProperty("width")
										JavaWindow("SystemsEngineering").JavaObject("SEComponentTab").Click TabWidth-5,5,"LEFT"
										Fn_SE_DetailTableOperation=True
										Exit for
								End If
							elseIf cstr(sText) = cstr(sObjectName)  Then
										objDetailsTable.SelectCell iCounter,sColumnName
										Wait 1
										objDetailsTable.SelectCell iCounter,sColumnName
                                        ObjDetailsTable.ClickCell iCounter, sColumnName,"LEFT","CONTROL"
										wait 3
										'JavaWindow("SystemsEngineering").JavaEdit("Text").Type sExpectedValue
										 Call Fn_KeyBoardOperation("SendKey", "^(A)")
										Call Fn_KeyBoardOperation("SendKeys", sExpectedValue)
										Call Fn_KeyBoardOperation("SendKeys", "{ENTER}")
										Fn_SE_DetailTableOperation=True
										Exit for
							End If								
					Next
		 Case "MultiRowPopupMenuSelect"
					aMenuList = split(sPopUpMenu, ":",-1,1)
					intCount = Ubound(aMenuList)
					'Split the string where " : " exist
					aObjList = Split(sObjectName,":")
					intItemCount =ubound(aObjList)
					'Count number of rows of Table
					bReturn=Fn_UI_Object_GetROProperty("Fn_SE_DetailTableOperation",ObjDetailsTable,"rows")
					'Extract the index of row at which the object exist.
					For oCounter=0 to intItemCount
							For iCounter=0 to bReturn-1
							sText = ObjDetailsTable.GetCellData(iCounter,"Object")						
							If IsNumeric(aObjList(oCounter)) Then
								 If cstr(sText) = cstr(cint(aObjList(oCounter)))  Then
									 ObjDetailsTable.ClickCell iCounter, "Object","LEFT","CONTROL"
									 Exit for
								End If
							ElseIf cstr(sText) = cstr(aObjList(oCounter))  Then
									 ObjDetailsTable.ClickCell iCounter, "Object","LEFT","CONTROL"									 
									 Exit for
							End If									
							Next
					Next
					'ObjDetailsTable.ClickCell bReturn-1, "Object","RIGHT","CONTROL"
						Set WshShell = CreateObject("WScript.Shell")
						WAIT(3)
						WshShell.SendKeys "+{F10}"
						Set WshShell = Nothing		
					Select Case intCount
						Case "0"
							 StrMenu = JavaWindow("SystemsEngineering").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0))
						Case "1"
							StrMenu = JavaWindow("SystemsEngineering").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1))
						Case "2"
							StrMenu = JavaWindow("SystemsEngineering").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1),aMenuList(2))
						Case Else
							Fn_SE_DetailTableOperation = False
							Exit Function
					End Select
					JavaWindow("SystemsEngineering").WinMenu("ContextMenu").Select StrMenu
					Fn_SE_DetailTableOperation=True
			'- - - - - - -  Added Case by Sandeep: Case to return All Column names currently exist in Details Table
			Case "AllColumnNames"
					'Returning All column Names present in Details Table
					Fn_SE_DetailTableOperation =Fn_UI_TableOperations("Fn_SE_DetailTableOperation","GetAllColumnNames",ObjDetailsTable,"","")
		End Select
        		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_SE_DetailTableOperation :  passed with case "&sAction&" on Object "&sObjectName)	
				Set ObjDetailsTable=Nothing 
End Function
'-------------------------------------------------------------------------This Function is used to Sort Details Table Containt-----------------------------------------------------------------------------------------------
'Function Name		:	Fn_SE_DetailTableSort

'Description			:	This Function is used to Sort Details Table Containt

'Parameters			:	 sSortOrder,strSortBy,strThenBy1,strThenBy2

'Return Value		:	True/False

'Pre-requisite		:	Details Table Should be Present
'											
'Examples			:	MsgBox Fn_SE_DetailTableSort( "Select default order","Name:Desc","Name:Desc","Name:Desc")
'									
'History:
'									Developer Name												Date								Rev. No.						Changes Done				Build	
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'									Harshal Agrawal									09/Feb/2011									1.0																				20110119
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SE_DetailTableSort(sSortOrder,strSortBy,strThenBy1,strThenBy2)
	GBL_FAILED_FUNCTION_NAME="Fn_SE_DetailTableSort"
   'Decalring Variables
   Dim ObjSortWnd,bFlag,sSortCriteria,sSortCriteria1,sSortCriteria2
   'Setting False to bFlag And Function
   bFlag=False
   Fn_SE_DetailTableSort=False
	'Ceating object of Sort window
	Set ObjSortWnd=Fn_UI_ObjectCreate("Fn_SE_DetailTableSort", JavaWindow("SystemsEngineering").JavaWindow("Sort"))

	'Select Sort Order
	Select Case sSortOrder
	Case "Select default order"
				Call Fn_UI_JavaRadioButton_SetON("Fn_SE_DetailTableSort",ObjSortWnd,"SelectDefaultOrder")
	Case "Select below criteria"
				Call Fn_UI_JavaRadioButton_SetON("Fn_SE_DetailTableSort",ObjSortWnd,"SelectCriteria")
	End Select


	'Spliting first sort parameter
	sSortCriteria=Split(strSortBy,":")
	'Apllying first sort criteria
	If sSortCriteria(0)<>"" Then
		'Checking existance of Item in Sort item list
		bFlag=Fn_UI_ListItemExist("Fn_SE_DetailTableSort", ObjSortWnd, "SortBy",sSortCriteria(0))
		If bFlag=False Then
            Call Fn_Button_Click("Fn_SE_DetailTableSort", ObjSortWnd,"Cancel")
            Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Fail: Wrong Search Criteria"& sSortCriteria(0) &"Pass")
			Set ObjSortWnd=Nothing
			Exit Function
		End If
		'Selecting item
		Call Fn_List_Select("Fn_SE_DetailTableSort", ObjSortWnd, "SortBy",sSortCriteria(0))
	End If
	'Selecting Sort type Ascending Or Descending
	If sSortCriteria(1)<>"" Then
		If sSortCriteria(1)="Asc" Then
			Call Fn_UI_JavaRadioButton_SetON("Fn_SE_DetailTableSort",ObjSortWnd,"Ascending")
		ElseIf	sSortCriteria(1)="Desc" Then
			Call Fn_UI_JavaRadioButton_SetON("Fn_SE_DetailTableSort",ObjSortWnd,"Descending")
		Else
			Call Fn_Button_Click("Fn_SE_DetailTableSort", ObjSortWnd,"Cancel")
			Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Fail: Invalid criteria pass")
			Set ObjSortWnd=Nothing
			Exit Function
		End If
	End If
	'Spliting Second sort parameter
    sSortCriteria1=Split(strThenBy1,":")
	'Apllying Second sort criteria
	If sSortCriteria1(0)<>"" Then
		'Checking existance of Item in Sort item list
		bFlag=Fn_UI_ListItemExist("Fn_SE_DetailTableSort", ObjSortWnd, "ThenBy_1",sSortCriteria1(0))
		If bFlag=False Then
            Call Fn_Button_Click("Fn_SE_DetailTableSort", ObjSortWnd,"Cancel")
            Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Fail: Wrong Search Criteria"& sSortCriteria1(0) &"Pass")
			Set ObjSortWnd=Nothing
			Exit Function
		End If
		'Selecting item
		Call Fn_List_Select("Fn_SE_DetailTableSort", ObjSortWnd, "ThenBy_1",sSortCriteria1(0))
	End If
	'Selecting Sort type Ascending Or Descending
	If sSortCriteria1(1)<>"" Then
		If sSortCriteria1(1)="Asc" Then
			Call Fn_UI_JavaRadioButton_SetON("Fn_SE_DetailTableSort",ObjSortWnd,"Ascending_2")
		ElseIf	sSortCriteria1(1)="Desc" Then
			Call Fn_UI_JavaRadioButton_SetON("Fn_SE_DetailTableSort",ObjSortWnd,"Descending_2")
		Else
			Call Fn_Button_Click("Fn_SE_DetailTableSort", ObjSortWnd,"Cancel")
			Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Fail: Invalid criteria pass")
			Set ObjSortWnd=Nothing
			Exit Function
		End If
	End If
	'Spliting Third sort parameter
	sSortCriteria2=Split(strThenBy2,":")
	'Apllying Third sort criteria
	If sSortCriteria2(0)<>"" Then
		'Checking existance of Item in Sort item list
		bFlag=Fn_UI_ListItemExist("Fn_SE_DetailTableSort", ObjSortWnd, "ThenBy_2",sSortCriteria2(0))
		If bFlag=False Then
            Call Fn_Button_Click("Fn_SE_DetailTableSort", ObjSortWnd,"Cancel")
            Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Fail: Wrong Search Criteria"& sSortCriteria2(0) &"Pass")
			Set ObjSortWnd=Nothing
			Exit Function
		End If
		'Selecting item
		Call Fn_List_Select("Fn_SE_DetailTableSort", ObjSortWnd, "ThenBy_2",sSortCriteria2(0))
	End If
	'Selecting Sort type Ascending Or Descending
	If sSortCriteria2(1)<>"" Then
		If sSortCriteria2(1)="Asc" Then
			Call Fn_UI_JavaRadioButton_SetON("Fn_SE_DetailTableSort",ObjSortWnd,"Ascending_3")
		ElseIf	sSortCriteria2(1)="Desc" Then
			Call Fn_UI_JavaRadioButton_SetON("Fn_SE_DetailTableSort",ObjSortWnd,"Descending_3")
		Else
			Call Fn_Button_Click("Fn_SE_DetailTableSort", ObjSortWnd,"Cancel")
			Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Fail: Invalid criteria pass")
			Set ObjSortWnd=Nothing
			Exit Function
		End If
	End If
	'Function Return True
	Fn_SE_DetailTableSort=True
	Call Fn_Button_Click("Fn_SE_DetailTableSort", ObjSortWnd,"OK")
	'Releasing Sort window object
	Set ObjSortWnd=Nothing
End Function
'-------------------------------------------------------------------------This Function is used to to perform operation on Save Column Application menu--------------------------------------------------------
'Function Name		:	Fn_SE_SaveColumnConfiguration

'Description			:	This Function is used to to perform operation on Save Column Application menu

'Parameters			:	   1.) strConfigName: Name of Configuration (It should be unique in ConfigurationSaveAs and Add case )
'										2.)strConfigDesc:Description of Configuration

'Return Value		:	True/False

'Pre-requisite		:	Should be details table present
'											
'Examples			:	Fn_SE_SaveColumnConfiguration("TestConfig","Test Configuration")
'								
'History:
'									Developer Name												Date								Rev. No.						Changes Done				Build
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'								Harshal Agrawal											09/Feb/2011								1.0																				20110119
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SE_SaveColumnConfiguration(strConfigName,strConfigDesc)
	GBL_FAILED_FUNCTION_NAME="Fn_SE_SaveColumnConfiguration"
  'Declaring Variables
   Dim ObjColumnConfigWnd
   Fn_SE_SaveColumnConfiguration=False
   'verifying existance of ApplyColumnConfiguration window
	If Fn_UI_ObjectExist("Fn_SE_SaveColumnConfiguration",JavaWindow("SystemsEngineering").JavaWindow("SaveColumnConfiguration"))=False Then
		'Invoking ApplyColumnConfiguration window
		Call Fn_SE_DetailTableOperation("PopUpMenuSelect","", "", "","Save Column Configuration...")
	End If
	'Creating objects 	
	Set ObjColumnConfigWnd=JavaWindow("SystemsEngineering").JavaWindow("SaveColumnConfiguration")
	'Setting Configuration Name
    Call Fn_UI_EditBox_Type("Fn_SE_SaveColumnConfiguration",ObjColumnConfigWnd,"Name",strConfigName)
	If strConfigDesc<>"" Then
		'Setting Description of configuration 
		Call Fn_Edit_Box("Fn_SE_SaveColumnConfiguration",ObjColumnConfigWnd,"Description",strConfigDesc)
	End If
	'Clicking on save button to create configuration
	Call Fn_Button_Click("Fn_SE_SaveColumnConfiguration",ObjColumnConfigWnd,"Save")
    Fn_SE_SaveColumnConfiguration=True
    'Releasing all objects
	Set ObjColumnConfigWnd=Nothing
End Function 
'-------------------------------------------------------------------------This Function is used to to perform operation on Apply Column Application menu--------------------------------------------------------
'Function Name		:	Fn_SE_ApplyColumnConfiguration

'Description			:	This Function is used to to perform operation on Apply Column Application menu

'Parameters			:	   1.) sAction: Action string to navigate to appropriate case
'									    2.) strConfigName: Name of Configuration (It should be unique in ConfigurationSaveAs and Add case )
'  										3.) arrAvailableProp: Avaiable Propeties array
'										4.) bShowIntPropName: Show Internal names of Properties option
'										5.)strConfigDesc:Description of Configuration

'Return Value		:	True/False

'Pre-requisite		:	Should be details table present
'											
'Examples			:	Fn_SE_ApplyColumnConfiguration("ConfigurationSaveAs","Demo4","","","")
'									Fn_SE_ApplyColumnConfiguration("ColumnAdd","Demo5",columnName,"","")
'History:
'									Developer Name												Date								Rev. No.						Changes Done				Build
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'									Harshal Agrawal											09/Feb/2011							1.0																				20110119
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SE_ApplyColumnConfiguration(strAction,strConfigName,arrAvailableProp,bShowIntPropName,strConfigDesc)
	GBL_FAILED_FUNCTION_NAME="Fn_SE_ApplyColumnConfiguration"
   'Declaring variables
   Dim bFlag,iCounter,iRowCount,iCnt,sDsplColName,iAvlRowCount,intCount,avlColName
   'Declaring Object
   bFlag=False
   Dim ObjColumnWnd,ObjColumnMngmntWnd,ObjColumnConfigWnd
   Fn_SE_ApplyColumnConfiguration=False
   'verifying existance of ApplyColumnConfiguration window
	If Fn_UI_ObjectExist("Fn_SE_ApplyColumnConfiguration",JavaWindow("SystemsEngineering").JavaWindow("ApplyColumnConfiguration"))=False Then
		If strAction = "VerifyNameInApplyColumnCofiguration" Then
			Call Fn_ToolBarOperation("ShowDropdownAndSelectWithInstance", "View Menu:3", "Apply Column Configuration...")
		Else
			Call Fn_SE_DetailTableOperation("PopUpMenuSelect","", "", "","Apply Column Configuration...")	
		End If
	End If
	'Creating objects 
	Set ObjColumnWnd=Fn_UI_ObjectCreate("Fn_SE_ApplyColumnConfiguration",JavaWindow("SystemsEngineering").JavaWindow("ApplyColumnConfiguration"))
	Set ObjColumnMngmntWnd=JavaWindow("SystemsEngineering").JavaWindow("ApplyColumnConfiguration").JavaWindow("ColumnManagement")
	Set ObjColumnConfigWnd=JavaWindow("SystemsEngineering").JavaWindow("ApplyColumnConfiguration").JavaWindow("ColumnManagement").JavaWindow("SaveColumnConfiguration")
	Select Case strAction
		Case "ConfigurationSaveAs" 'This Case use to create same as default configuration
			Call Fn_Button_Click("Fn_SE_ApplyColumnConfiguration",ObjColumnWnd,"Add")
			'Verifying existance Column Management window
			If Fn_UI_ObjectExist("Fn_SE_ApplyColumnConfiguration",ObjColumnMngmntWnd)=False Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail:Failed to Invoke Column Management Window")	
				Set ObjColumnWnd=Nothing
				Set ObjColumnMngmntWnd=Nothing
				Set ObjColumnConfigWnd=Nothing
				Exit Function
			End If
			'Clicking on save button to open SaveColumnConfiguration
			Call Fn_Button_Click("Fn_SE_ApplyColumnConfiguration",ObjColumnMngmntWnd,"Save")
			'Setting name of configuration 
            Call Fn_UI_EditBox_Type("Fn_SE_ApplyColumnConfiguration",ObjColumnConfigWnd,"Name",strConfigName)
			If strConfigDesc<>"" Then
				'Setting Description of configuration 
				Call Fn_Edit_Box("Fn_SE_ApplyColumnConfiguration",ObjColumnConfigWnd,"Description",strConfigDesc)
			End If
			'Clicking on save button to create configuration
			Call Fn_Button_Click("Fn_SE_ApplyColumnConfiguration",ObjColumnConfigWnd,"Save")
			'Closing the window
			Call Fn_Button_Click("Fn_SE_ApplyColumnConfiguration",ObjColumnMngmntWnd,"Close")
			'Checking existance of Configuration in ColumnConfigurations list
            bFlag=Fn_UI_ListItemExist("Fn_SE_ApplyColumnConfiguration",ObjColumnWnd, "Column Configurations",strConfigName)
			If bFlag=True Then
				'Selecting Confugarion in ColumnConfigurations list
				Call Fn_List_Select("Fn_SE_ApplyColumnConfiguration",ObjColumnWnd, "Column Configurations",strConfigName)
				'Applying the changes
				Call Fn_Button_Click("Fn_SE_ApplyColumnConfiguration",ObjColumnWnd,"Apply")
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail:Failed to create confugaration of name" & strConfigName)	
				Set ObjColumnWnd=Nothing
				Set ObjColumnMngmntWnd=Nothing
				Set ObjColumnConfigWnd=Nothing
				Exit Function
			End If
	Case "ColumnAdd"
			Call Fn_Button_Click("Fn_SE_ApplyColumnConfiguration",ObjColumnWnd,"Add")
            Call Fn_ReadyStatusSync(1)
			'Verifying existance Column Management window
			arrAvailableProp=split(arrAvailableProp,":")
			'If  IsArray(arrAvailableProp) Then
				For iCounter=0 To Ubound(arrAvailableProp)
					 If arrAvailableProp(iCounter)<>"" Then
						 iRowCount=Fn_UI_Object_GetROProperty("Fn_ReqMgr_DetailTableOperation",ObjColumnMngmntWnd.JavaTable("DisplayedColumns"),"rows")
						 For iCnt=0 To iRowCount-1
							 bFlag=False
							sDsplColName=ObjColumnMngmntWnd.JavaTable("DisplayedColumns").GetCellData(iCnt,"0")
								If sDsplColName=arrAvailableProp(iCounter) Then
									bFlag=True
									Exit For
								End If
						Next
						If bFlag=False Then
							iAvlRowCount=Fn_UI_Object_GetROProperty("Fn_ReqMgr_DetailTableOperation",ObjColumnMngmntWnd.JavaTable("AvailableProperties"),"rows")
							'Selecting Confugarion in ColumnConfigurations list
							For intCount=0 To iAvlRowCount-1
									avlColName=ObjColumnMngmntWnd.JavaTable("AvailableProperties").GetCellData(intCount,"Property")
									If avlColName=arrAvailableProp(iCounter) Then
										ObjColumnMngmntWnd.JavaTable("AvailableProperties").SelectCell intCount,"Property"
										Call Fn_Button_Click("Fn_SE_ApplyColumnConfiguration",ObjColumnMngmntWnd,"Add")
										Exit For
								End If
							Next                     					
						End If
					 End If
				Next
			'End If
			'Clicking on save button to open SaveColumnConfiguration
			Call Fn_Button_Click("Fn_SE_ApplyColumnConfiguration",ObjColumnMngmntWnd,"Save")
			'Setting name of configuration 
            Call Fn_UI_EditBox_Type("Fn_SE_ApplyColumnConfiguration",ObjColumnConfigWnd,"Name",strConfigName)
			If strConfigDesc<>"" Then
				'Setting Description of configuration 
				Call Fn_Edit_Box("Fn_SE_ApplyColumnConfiguration",ObjColumnConfigWnd,"Description",strConfigDesc)
			End If
			'Clicking on save button to create configuration
			Call Fn_Button_Click("Fn_SE_ApplyColumnConfiguration",ObjColumnConfigWnd,"Save")
			'Closing the window
			Call Fn_Button_Click("Fn_SE_ApplyColumnConfiguration",ObjColumnMngmntWnd,"Close")
			'Checking existance of Configuration in ColumnConfigurations list
            bFlag=Fn_UI_ListItemExist("Fn_SE_ApplyColumnConfiguration",ObjColumnWnd, "Column Configurations",strConfigName)
			If bFlag=True Then
				'Selecting Confugarion in ColumnConfigurations list
				Call Fn_List_Select("Fn_SE_ApplyColumnConfiguration",ObjColumnWnd, "Column Configurations",strConfigName)
				'Applying the changes
				wait 2 
				Call Fn_Button_Click("Fn_SE_ApplyColumnConfiguration",ObjColumnWnd,"Apply")
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail:Failed to create confugaration of name" & strConfigName)	
				Set ObjColumnWnd=Nothing
				Set ObjColumnMngmntWnd=Nothing
				Set ObjColumnConfigWnd=Nothing
				Exit Function
			End If
	 Case "Apply"
			bFlag=Fn_UI_ListItemExist("Fn_SE_ApplyColumnConfiguration",ObjColumnWnd, "Column Configurations",strConfigName)
			If bFlag=True Then
				'Selecting Confugarion in ColumnConfigurations list
				Call Fn_List_Select("Fn_SE_ApplyColumnConfiguration",ObjColumnWnd, "Column Configurations",strConfigName)
				'Applying the changes
				Call Fn_Button_Click("Fn_SE_ApplyColumnConfiguration",ObjColumnWnd,"Apply")
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail:Failed to create confugaration of name" & strConfigName)	
				Set ObjColumnWnd=Nothing
				Set ObjColumnMngmntWnd=Nothing
				Set ObjColumnConfigWnd=Nothing
				Exit Function
			End If
			
	Case "VerifyNameInApplyColumnCofiguration"
			Call Fn_ReadyStatusSync(1)
			'Verifying in list the names
			avlColName = split(strConfigName,":")
			For iCounter = 0 To uBound(avlColName)
				bFlag = Fn_SISW_UI_JavaList_Operations("Fn_SE_ApplyColumnConfiguration", "Exist", ObjColumnWnd,"Column Configurations",avlColName(iCounter), "", "")
				If bFlag=False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail: The Column Name does not exist")
					Call Fn_Button_Click("Fn_SE_ApplyColumnConfiguration", ObjColumnWnd, "Close")
					Set ObjColumnWnd=Nothing
					Set ObjColumnMngmntWnd=Nothing
					Set ObjColumnConfigWnd=Nothing
					Exit Function
				End If
			Next
			Call Fn_ReadyStatusSync(1)
		
	End Select
	Fn_SE_ApplyColumnConfiguration=True
	Call Fn_Button_Click("Fn_SE_ApplyColumnConfiguration",ObjColumnWnd,"Close")
	'Releasing all objects
	Set ObjColumnWnd=Nothing
	Set ObjColumnMngmntWnd=Nothing
	Set ObjColumnConfigWnd=Nothing
End Function
'-------------------------------------------------------------------------This Function is used to to Filter content of Details Table-------------------------------------------------------------------------------------
'Function Name		:	Fn_SE_DetailsTableFilterManagement

'Description			:	This Function is used to to Filter content of Details Table

'Parameters			:	   1.) sAction: Action Name
'									    2.) strConditionName: Condition Name ("Object == REQ-001140/A;1-Req")
'  										3.) strColName: Column Name to set condition (Object,Type,Group ID.................)
'										4.) strOperator: Operator to Set Condition (==,=,!=,<>.......................)
'										5.) strColValue:Column Velue to Set Condition (REQ-001140/A;1-Req..............)
'										6.)strLogicalType:Logical Type (And or OR....)

'Return Value		:	True/False

'Pre-requisite		:	Should be details table present
'											
'Examples			:	Msgbox Fn_SE_DetailsTableFilterManagement("AddCondition","Type!=RequirementSpec Revision Master","Type","!=","RequirementSpec Revision Master","And")
'									Msgbox Fn_SE_DetailsTableFilterManagement("ApplyCondition","Type!=RequirementSpec Revision Master","","","","")
'									
'History:
'									Developer Name												Date								Rev. No.						Changes Done		
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'									Ketan Raje									   				10/02/2011			              1.0										Created					
'									Sushma	P								   				24/10/2011			              1.1										Added Case "ApplyCondition"
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SE_DetailsTableFilterManagement(strAction,strConditionName,strColName,strOperator,strColValue,strLogicalType)
	GBL_FAILED_FUNCTION_NAME="Fn_SE_DetailsTableFilterManagement"
   'Veriable Declaration
	Dim bFlag,iRowCnt,iCounter,strCellData
	Dim ObjAutoFilterWnd
	bFlag=False
	Fn_SE_DetailsTableFilterManagement=False
   'verifying existance of AutoFilter window
	If Fn_UI_ObjectExist("Fn_SE_DetailsTableFilterManagement",JavaWindow("SystemsEngineering").JavaWindow("AutoFilter"))=False Then
		'Invoking AutoFilter window
		Call Fn_ToolbatButtonClick("Filter Management")
	End If
	'Creating objects 
	Set ObjAutoFilterWnd=Fn_UI_ObjectCreate("Fn_SE_DetailsTableFilterManagement",JavaWindow("SystemsEngineering").JavaWindow("AutoFilter"))
	
	Select Case strAction
		'Case to Add New Condiation
		Case "AddCondition" 
					'Clicking plus button
		            Call Fn_Button_Click("Fn_SE_DetailsTableFilterManagement", ObjAutoFilterWnd, "PlusButton")
					'Setting Column Name Condition
					If strColName<>"" Then
						bFlag=Fn_UI_ListItemExist("Fn_SE_DetailsTableFilterManagement", ObjAutoFilterWnd.JavaWindow("FilterConditionEditor"),"ColumnList",strColName)
						If bFlag=False Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Fail: "& strColName &"Column is not present in List" ) 
							Set ObjAutoFilterWnd=Nothing
							Exit Function
						End If
							Call Fn_List_Select("Fn_SE_DetailsTableFilterManagement", ObjAutoFilterWnd.JavaWindow("FilterConditionEditor"),"ColumnList",strColName)
					End If
					bFlag=False
					'Setting Operator Condiation
					If strOperator<>"" Then
						bFlag=Fn_UI_ListItemExist("Fn_SE_DetailsTableFilterManagement", ObjAutoFilterWnd.JavaWindow("FilterConditionEditor"),"OperatorList",strOperator)
						If bFlag=False Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Fail: "& strOperator &"Column is not present in List" ) 
							Set ObjAutoFilterWnd=Nothing
							Exit Function
						End If
							Call Fn_List_Select("Fn_SE_DetailsTableFilterManagement", ObjAutoFilterWnd.JavaWindow("FilterConditionEditor"),"OperatorList",strOperator)
					End If
					bFlag=False
					'Setting Column Value Condition
					If strColValue<>"" Then
						bFlag=Fn_UI_ListItemExist("Fn_SE_DetailsTableFilterManagement", ObjAutoFilterWnd.JavaWindow("FilterConditionEditor"),"ObjectNameList",strColValue)
						If bFlag=False Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Fail: "& strColValue &"Column is not present in List" ) 
							Set ObjAutoFilterWnd=Nothing
							Exit Function
						End If
							Call Fn_List_Select("Fn_SE_DetailsTableFilterManagement", ObjAutoFilterWnd.JavaWindow("FilterConditionEditor"),"ObjectNameList",strColValue)
					End If
					'Clicking Plus button to add condition into Table
					Call Fn_Button_Click("Fn_SE_DetailsTableFilterManagement", ObjAutoFilterWnd.JavaWindow("FilterConditionEditor"), "plusButton")
					bFlag=False
                    iRowCnt=Fn_UI_Object_GetROProperty("Fn_SE_DetailsTableFilterManagement", ObjAutoFilterWnd.JavaWindow("FilterConditionEditor").JavaTable("Table"),"rows")
					For iCounter=0 To iRowCnt-1
							strCellData=Fn_UI_JavaTable_GetCellData("Fn_SE_DetailsTableFilterManagement",ObjAutoFilterWnd.JavaWindow("FilterConditionEditor"), "Table",iCounter,0)
							If Trim(Replace(strCellData," ",""))=Trim(Replace(strConditionName," ","")) Then
								iRowCnt=iCounter
								bFlag=True
								Exit For
							End If
					Next
					If bFlag=True Then
                        Call Fn_UI_JavaTable_SelectRow("Fn_SE_DetailsTableFilterManagement",ObjAutoFilterWnd.JavaWindow("FilterConditionEditor"), "Table",iRowCnt)
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Fail: Invalid parameter pass" ) 
						Set ObjAutoFilterWnd=Nothing
						Exit Function
					End If
					'Clicking ok to Apply condition
					Call Fn_Button_Click("Fn_SE_DetailsTableFilterManagement",ObjAutoFilterWnd.JavaWindow("FilterConditionEditor"),"OK")

					 iRowCnt=Fn_UI_Object_GetROProperty("Fn_SE_DetailsTableFilterManagement", ObjAutoFilterWnd.JavaTable("ConditionTable"),"rows")
					For iCounter=0 To iRowCnt-1
							strCellData=Fn_UI_JavaTable_GetCellData("Fn_SE_DetailsTableFilterManagement",ObjAutoFilterWnd,"ConditionTable",iCounter,0)
							If Trim(Replace(strCellData," ",""))=Trim(Replace(strConditionName," ","")) Then
                                JavaWindow("SystemsEngineering").JavaWindow("AutoFilter").JavaTable("ConditionTable").ActivateCell iCounter,0
								Call Fn_UI_JavaTable_SelectRow("Fn_SE_DetailsTableFilterManagement",ObjAutoFilterWnd, "ConditionTable",iCounter)
								Exit For
							End If
					Next
			Case "ApplyCondition" 
					 iRowCnt=Fn_UI_Object_GetROProperty("Fn_SE_DetailsTableFilterManagement", ObjAutoFilterWnd.JavaTable("ConditionTable"),"rows")
					For iCounter=0 To iRowCnt-1
							strCellData=Fn_UI_JavaTable_GetCellData("Fn_SE_DetailsTableFilterManagement",ObjAutoFilterWnd,"ConditionTable",iCounter,0)
							If Trim(Replace(strCellData," ",""))=Trim(Replace(strConditionName," ","")) Then
                                JavaWindow("SystemsEngineering").JavaWindow("AutoFilter").JavaTable("ConditionTable").ActivateCell iCounter,0
								Call Fn_UI_JavaTable_SelectRow("Fn_SE_DetailsTableFilterManagement",ObjAutoFilterWnd, "ConditionTable",iCounter)
								Exit For
							End If
					Next
	End Select
	Fn_SE_DetailsTableFilterManagement=True
	'Closing Window
	Call Fn_Button_Click("Fn_SE_DetailsTableFilterManagement", ObjAutoFilterWnd, "Close")
	'Releasing object Auto Filter Window
	Set ObjAutoFilterWnd=Nothing
End Function
'*********************************************************		Function to create basic Item		***********************************************************************
'Function Name		:				Fn_SE_CAEItemBasicCreate

'Description			 :		 		 Creats an Requirement Specification with basic information

'Parameters			   :	 			1.sSpecType: Type of the item.(e.g. Requirement Specification)
'													 2.sConfItem: True or False
'													 2.sSpecID: ID of the Item it should be unique.
'													3.sSpecRevID:Revision ID of the Item.
'													4.sSpecName:Name of Item.
'													5.sSpecDesc: Description of the Item.
'													6:sSpecUOM: Unit of measure of Item. ( not handling this part)

'Return Value		   : 				Item Id  / Revision Id

'Pre-requisite			:		 		should be logged in & present on System Engineering perspective

'Examples				:				"RequirmentSpecification"  : Msgbox Fn_SE_CAEItemBasicCreate("RequirementSpec","OFF","","","ReqSpecTest1","","")
'												"Paragraph" : Msgbox Fn_SE_CAEItemBasicCreate("Paragraph","OFF","","","TestPara1","","")
'												"Requirment" : Msgbox Fn_SE_CAEItemBasicCreate("Requirement","OFF","","","ReqTest1","","")

'History					 :		
'								Developer Name					Date				Rev. No.		Changes Done		
'								--------------------------------------------------------------------------------------------------
'								Ketan Raje							21/02/2011	        1.0					Created					
'								--------------------------------------------------------------------------------------------------
Public Function Fn_SE_CAEItemBasicCreate(sSpecType,sConfItem,sSpecID,sSpecRevID,sSpecName,sSpecDesc,sSpecUOM)
	GBL_FAILED_FUNCTION_NAME="Fn_SE_CAEItemBasicCreate"
	on Error Resume Next
	Dim sSpecificationId, sRevId
	Dim objDialogNewSpec,objSelectType,objDialog

	If Fn_UI_ObjectExist("Fn_SE_CAEItemBasicCreate",Window("SEWindow").JavaDialog("NewCAEItem"))=False Then
         Call Fn_MenuOperation("Select","Edit:Insert Level")
	End If
	Wait(2)
	'Check the existence of "NewCAEItem" window
	Set objDialogNewSpec=Fn_UI_ObjectCreate("Fn_SE_CAEItemBasicCreate",Window("SEWindow").JavaDialog("NewCAEItem"))
		'Select Item Type
		Call Fn_List_Select("Fn_SE_CAEItemBasicCreate", objDialogNewSpec,"CAEItemType",sSpecType)
		'checked Configuration RequirementSpec or not
		If sConfItem <> "" Then
		 Call Fn_CheckBox_Set("Fn_SE_CAEItemBasicCreate", objDialogNewSpec,"Configuration Item",sConfItem)
		End If
		Wait(2)
		'Click on "Next" button
		 Call Fn_Button_Click("Fn_SE_CAEItemBasicCreate", objDialogNewSpec,"Next")
		
		If sSpecID <> "" Then
			'Set  Item Id
			 Call Fn_Edit_Box("Fn_SE_CAEItemBasicCreate", objDialogNewSpec,"CAEItemID", sSpecID)
		End If
		
		If sSpecRevID <> "" Then
			'Set Revision ID
			Call Fn_Edit_Box("Fn_SE_CAEItemBasicCreate",objDialogNewSpec,"RevID", sSpecRevID)
		End If
		Wait(2)
		If  sSpecID = "" or sSpecRevID = "" Then
			'click on assign button
			  Call Fn_Button_Click("Fn_SE_CAEItemBasicCreate", objDialogNewSpec, "Assign")
		End If
		
		'Extract Creation data
		sSpecificationId =Fn_Edit_Box_GetValue("Fn_SE_CAEItemBasicCreate", objDialogNewSpec,"CAEItemID")
		sRevId = Fn_Edit_Box_GetValue("Fn_SE_CAEItemBasicCreate", objDialogNewSpec,"RevID")
		
		'Set RequirementSpec name
		 Call Fn_Edit_Box("Fn_SE_CAEItemBasicCreate", objDialogNewSpec,"SpecName",sSpecName)
		'Set description
		Call Fn_Edit_Box("Fn_SE_CAEItemBasicCreate", objDialogNewSpec,"Description",sSpecDesc)
		'Set UOM
			If sSpecUOM <> "" Then
				 Set objSelectType=description.Create()
				objSelectType("Class Name").value = "JavaStaticText"
				objSelectType("label").value = sSpecUOM
				objDialogNewSpec.JavaButton("UnitOfMeasureDrpDwn").Click
				Set objDialog =objDialogNewSpec.ChildObjects(objSelectType)
				objDialog(0).Click 5, 5, "LEFT"
		   End If
		
		 wait(2)
			objDialogNewSpec.JavaButton("Finish").WaitProperty "enabled", 1, 20000
        			
			Call Fn_Button_Click("Fn_SE_CAEItemBasicCreate", objDialogNewSpec, "Finish") 
			Fn_SE_CAEItemBasicCreate = sSpecificationId & "-" & sRevId
			Call Fn_ReadyStatusSync(1)

			If Fn_UI_ObjectExist("Fn_SE_CAEItemBasicCreate",Window("SEWindow").JavaDialog("NewCAEItem"))=True Then		
					Call Fn_Button_Click("Fn_SE_CAEItemBasicCreate", objDialogNewSpec, "Close")
			End If
		
        Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Created Specification of ID [" + CStr(sItemId) + "]")
		Set objDialogNewSpec=Nothing
		Set objSelectType=Nothing
		Set objDialog=Nothing
End Function
'#########################################################################################################
'###
'###    FUNCTION NAME   :   Fn_SE_DetailTableConfigOperation(aColumnNames, sTableConfigName, sTableConfigNewName, sTableConfigDescription, sAction)
'###
'###    DESCRIPTION        :  	System Engineering utility function which allpied to Details Table Config Operation
'###
'###    PARAMETERS      :     1.  sAction: Action string to navigate to appropriate case
'### 								       2.  aColumnNames: Array of Column Names          
'###								       3.  aColumnType: Array of column Type 
'###									   4.  sTableConfigName: Table Configuration Name
'### 									   5.  sTableConfigNewName: Table Configuration New Name
'### 									   6.  sTableConfigDescription: Table Configuration Description
'###
'###  	HISTORY             :   AUTHOR                 DATE        	VERSION
'###
'### 	CREATED BY     :   Ketan Raje 	              7-Mar-2011       	1.0
'###
'###    EXAMPLES
'###         Case 1 :       ColumnAdd :  Msgbox Fn_SE_DetailTableConfigOperation("ColumnAdd", "Relation", "", "", "", "")
'###         Case 2:        ColumnRemove : Msgbox Fn_SE_DetailTableConfigOperation("ColumnRemove", "Relation", "", "", "", "")
'###         Case 3:        ColumnValidateExists  : Msgbox Fn_SE_DetailTableConfigOperation("ColumnValidateExists", "", "", "", "", "Object")
'###         Case 4: 		MoveColumnUp		:Msgbox Fn_SE_DetailTableConfigOperation("MoveColumnUp", "Description", "4", "", "", "")
'###         Case 5: 		MoveColumnDown	:Msgbox Fn_SE_DetailTableConfigOperation("MoveColumnDown", "Description", "4", "", "", "")
'############################################################################################################# 

Public Function Fn_SE_DetailTableConfigOperation(sAction, aColumnNames, aColumnType, sTableConfigName, sTableConfigNewName, sTableConfigDescription)
		GBL_FAILED_FUNCTION_NAME="Fn_SE_DetailTableConfigOperation"
		Dim  iCounter, sMenu, ObjDialog, bReturn, bFlag 
		 Dim  bReturn1,hCount
		Fn_SE_DetailTableConfigOperation = False
		bFlag = False
		Set ObjDetails = JavaWindow("SystemsEngineering").JavaTable("DetailsTable")
		
		If Not JavaWindow("SystemsEngineering").JavaWindow("ColumnManagement").Exist(2) Then
			ObjDetails.SelectColumnHeader 0,"LEFT"
		End If

		Select Case sAction
				Case "ColumnAdd"
							If  Fn_SE_DetailTableConfigOperation("ColumnValidateExists", aColumnNames, "", "", "", "") = True Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), aColumnNames&" column already exist in Details Table.")
									Fn_SE_DetailTableConfigOperation = True
							Else
									If Not JavaWindow("SystemsEngineering").JavaWindow("ColumnManagement").JavaTable("AvailableProp").Exist(4) Then
'										Call Fn_ToolbarButtonClick_Ext(1, "View Menu")
										Call Fn_ToolbarButtonClick_Ext(3, "View Menu")
										sMenu =  JavaWindow("SystemsEngineering").WinMenu("ContextMenu").BuildMenuPath("Column...")
										JavaWindow("SystemsEngineering").WinMenu("ContextMenu").Select sMenu   
									End If
									Set ObjDialog = JavaWindow("SystemsEngineering").JavaWindow("ColumnManagement")
									'Count number of rows of Table
									bReturn = ObjDialog.JavaTable("AvailableProp").GetROProperty("rows")	
									'Extract the index of row at which the object exist.
									For iCounter=0 to bReturn - 1
										If Trim(Lcase(ObjDialog.JavaTable("AvailableProp").GetCellData(iCounter,"Property"))) = Trim(Lcase(aColumnNames)) And Trim(Lcase(ObjDialog.JavaTable("AvailableProp").GetCellData(iCounter,"Type"))) = Trim(Lcase(aColumnType)) then
											ObjDialog.JavaTable("AvailableProp").ClickCell iCounter,0 
											bFlag = True
											Exit For
										End If
									Next
									'Click on Add column Button
									Call Fn_Button_Click("Fn_SE_DetailTableConfigOperation", ObjDialog, "AddColumn")
									'Click on Apply Button
									wait 2
									Call Fn_Button_Click("Fn_SE_DetailTableConfigOperation", ObjDialog, "Apply")
									'Click on Close Button
									Call Fn_Button_Click("Fn_SE_DetailTableConfigOperation", ObjDialog, "Close")
									If bFlag=True Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_SE_DetailTableConfigOperation passed with case "&sAction)
										Fn_SE_DetailTableConfigOperation = True
									Else
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_SE_DetailTableConfigOperation passed with case "&sAction)
										Fn_SE_DetailTableConfigOperation = False
									End If									
							End If
				Case "ColumnRemove"
							If  Fn_SE_DetailTableConfigOperation("ColumnValidateExists", "", "", "", "", "aColumnNames") = False Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), aColumnNames&" column does not exist in Details Table.")
									Fn_SE_DetailTableConfigOperation = True
							Else
'									Call Fn_ToolbarButtonClick_Ext(2, "View Menu")
									Call Fn_ToolbarButtonClick_Ext(3, "View Menu")
									sMenu =  JavaWindow("SystemsEngineering").WinMenu("ContextMenu").BuildMenuPath("Column...")
									JavaWindow("SystemsEngineering").WinMenu("ContextMenu").Select sMenu   	
									Set ObjDialog = JavaWindow("SystemsEngineering").JavaWindow("Column Management")
									'Count number of rows of Table
									bReturn = ObjDialog.JavaTable("DisplayedColumns").GetROProperty("rows")	
									'Extract the index of row at which the object exist.
									For iCounter=0 to bReturn - 1
										If ObjDialog.JavaTable("DisplayedColumns").GetCellData(iCounter,0) = aColumnNames then
											ObjDialog.JavaTable("DisplayedColumns").ClickCell iCounter,0
											bFlag = True
											Exit For
										End If
									Next
									'Click on Add column Button
									Call Fn_Button_Click("Fn_SE_DetailTableConfigOperation", ObjDialog, "RemoveColumn")
									'Click on Apply Button
									Call Fn_Button_Click("Fn_SE_DetailTableConfigOperation", ObjDialog, "Apply")
									'Click on Close Button
									Call Fn_Button_Click("Fn_SE_DetailTableConfigOperation", ObjDialog, "Close")
									If bFlag=True Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_SE_DetailTableConfigOperation passed with case "&sAction)
										Fn_SE_DetailTableConfigOperation = True
									Else
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_SE_DetailTableConfigOperation failed with case "&sAction)
										Fn_SE_DetailTableConfigOperation = False
									End If
							End If
				Case "ColumnValidateExists"
							bFlag = False
							bReturn = ObjDetails.GetROProperty("cols")
							For iCounter = 0 to bReturn - 1 
									'checking existance of column
									If Trim(Lcase(ObjDetails.GetColumnName(iCounter))) = Trim(Lcase(aColumnNames)) Then
											bFlag = True
											Exit For
									End If
							Next
							If bFlag=True Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Fn_SE_DetailTableConfigOperation passed with case "&sAction)
									Fn_SE_DetailTableConfigOperation = True
							Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Fn_SE_DetailTableConfigOperation failed with case "&sAction)
									Fn_SE_DetailTableConfigOperation = False											
							End If
				Case "MoveColumnUp","MoveColumnDown"                                  'Added by Avinash  J.[19/07/2012]
                    				If Not JavaWindow("SystemsEngineering").JavaWindow("ColumnManagement").JavaTable("AvailableProp").Exist(4) Then
							              'Call Fn_ToolbarButtonClick_Ext(1, "View Menu")
										Call Fn_ToolbarButtonClick_Ext(3, "View Menu")
										sMenu =  JavaWindow("SystemsEngineering").WinMenu("ContextMenu").BuildMenuPath("Column...")
									JavaWindow("SystemsEngineering").WinMenu("ContextMenu").Select sMenu   	
									End If
									Set ObjDialog = JavaWindow("SystemsEngineering").JavaWindow("ColumnManagement")
									'Count number of rows of Table
                            		bReturn = ObjDialog.JavaTable("DisplayedColumns").GetROProperty("rows")	
									bReturn1=bReturn
									'Extract the index of row at which the object exist.
									For iCounter=0 to bReturn - 1
										If ObjDialog.JavaTable("DisplayedColumns").GetCellData(iCounter,0) = aColumnNames then
											ObjDialog.JavaTable("DisplayedColumns").ClickCell iCounter,0
											bFlag = True
											Exit For
										End If
									Next                            

                          hCount= CInt(aColumnType)
						For iCounter=0 to hCount - 1                     ''It will click on specifed button[UP/Down] on specified number of times i.e aColumnType value
									    If sAction=Trim("MoveColumnUp")  Then
												      If  ObjDialog.JavaButton("MoveUp").GetROProperty("enabled")="1"  then
													      'Click on MoveUP  column Button
												           Call Fn_Button_Click("Fn_SE_DetailTableConfigOperation", ObjDialog, "MoveUp")
												    	    bFlag = True
												      End If
								    	ElseIf  sAction=Trim("MoveColumnDown") Then
												If  ObjDialog.JavaButton("MoveDown").GetROProperty("enabled")="1"   then
													  'Click on MoveDown  column Button
												       Call Fn_Button_Click("Fn_SE_DetailTableConfigOperation", ObjDialog, "MoveDown")
													   bFlag = True
												End If
									End If
				          Next
						
                    				'Click on Apply Button
										Call Fn_Button_Click("Fn_SE_DetailTableConfigOperation", ObjDialog, "Apply")
										'	Click on Close Button
										Call Fn_Button_Click("Fn_SE_DetailTableConfigOperation", ObjDialog, "Close")
									If bFlag=True  Then
                                    	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_SE_DetailTableConfigOperation passed with case "&sAction)
										Fn_SE_DetailTableConfigOperation = True
									Else
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_SE_DetailTableConfigOperation failed with case "&sAction)
										 Fn_SE_DetailTableConfigOperation = False
									 End If
    		End Select
		Set ObjDialog = nothing
End Function
'----------------------------------------------------------------------------------------------Function to Export Object to Excel----------------------------------------------------------------------------------------------------------------------------
'Function Name		:			Fn_SE_ExportToExcel

'Description			 :		 	Function to Export Object to Word

'Return Value		   : 			True Or False

'Pre-requisite			:			Object which is to be exported should be selected

'Examples				:			Msgbox Fn_SE_ExportToExcel("ExportToExcel", "Static Snapshot", "Use Excel Template", "REQ_default_excel_template", "http:///tc/reporturlclient?TcObjectId=Q5LAAABpoIOaJD&TcExportMode=Static&TcTemplateName=REQ_default_excel_template&TcReportType=BOMReport&TcReportFormat=ExcelReport", "Generate URL:OK")
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - -
'											 As Column Configuration parameter is not there so [ sOutputTemplate ] parameter use to pass vlues wich has to verify in Case "VerifyColumnConfigurations"
'											 Msgbox Fn_SE_ExportToExcel("VerifyColumnConfigurations", "", "* sss~* Config1", "", "", "")

'											Call Fn_SE_ExportToExcel("ExportToExcel", "Live integration with Excel (Interactive)~ON", "Use Excel Template", "REQ_default_excel_template", "", "OK")
'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done	
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Ketan Raje									   				10/03/2011			           1.0															
'													Sandeep N									   			  21/10/2011			          1.1						Added Case "VerifyColumnConfigurations"		
'													pranav Ingle									   			 16/11/2011			          1.1						Modified Case "ExportToExcel"								
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public function Fn_SE_ExportToExcel(sAction, sOutput, sOutputTemplate, sExcelTemplate, sURL, sButtons)
		GBL_FAILED_FUNCTION_NAME="Fn_SE_ExportToExcel"
		Dim bReturn, objExportDialog, aButtons,arrOutput
		arrOutput=Split(sOutput,"~")
		aButtons = Split(sButtons,":")
		Fn_SE_ExportToExcel = False
		Set objExportDialog = JavaWindow("SystemsEngineering").JavaWindow("ExportToExcel")
		If Fn_UI_ObjectExist("Fn_SE_ExportToExcel",objExportDialog) = False Then
				Call Fn_MenuOperation("Select","Tools:Export:Objects To Excel")
		End If
		Select Case sAction
				Case "ExportToExcel"
						If Ubound(arrOutput) = 1 Then 
								Call Fn_UI_Object_SetTOProperty("Fn_SE_ExportToExcel",objExportDialog.JavaRadioButton("Output"), "attached text",  arrOutput(0))
								Call Fn_UI_JavaRadioButton_SetON("Fn_SE_ExportToExcel", objExportDialog,"Output")
								wait(2)
                                If arrOutput(1) = "ON" Then
												Call Fn_CheckBox_Set("Fn_SE_ExportToExcel" ,JavaWindow("SystemsEngineering").JavaWindow("ExportToExcel"),"CheckOutObjs", "ON") 
								End If
						Else
								Call Fn_UI_Object_SetTOProperty("Fn_SE_ExportToExcel",objExportDialog.JavaRadioButton("Output"), "attached text", sOutput )
								Call Fn_UI_JavaRadioButton_SetON("Fn_SE_ExportToExcel", objExportDialog,"Output")
						End If
						'Set the value for Output Template Radio button.
						If sOutputTemplate <> "" Then
								Call Fn_UI_Object_SetTOProperty("Fn_SE_ExportToExcel",objExportDialog.JavaRadioButton("OutputTemplate"), "attached text",  sOutputTemplate)
								Call Fn_UI_JavaRadioButton_SetON("Fn_SE_ExportToExcel", objExportDialog,"OutputTemplate")								
						End If
						'Set the value for Output Template List.
						If sExcelTemplate <> "" Then
							If sOutputTemplate = "Use Excel Template" Then
								Call Fn_List_Select("Fn_SE_ExportToExcel",objExportDialog,"ExcelTemplate", sExcelTemplate)
							ElseIf sOutputTemplate = "Column Configurations" Then
								Call Fn_List_Select("Fn_SE_ExportToExcel",objExportDialog,"ColumnConfig", sExcelTemplate)
							End If								
						End If
						'Click on given Button
						Call Fn_Button_Click("Fn_SE_ExportToExcel", objExportDialog, aButtons(0))
						'Verify the generated URL.
						If aButtons(0) = "Generate URL" Then
							If objExportDialog.JavaWindow("URLGenerated").Exist = True Then
								'Get the URL value.
								If Trim(Lcase(sURL)) = Trim(Lcase(Fn_Edit_Box_GetValue("Fn_SE_ExportToExcel", objExportDialog.JavaWindow("URLGenerated"), "Details"))) Then
									'Click on OK button of URL Generated Dialog.
									Call Fn_Button_Click("Fn_SE_ExportToExcel", objExportDialog.JavaWindow("URLGenerated"), "OK")
									'Click on OK button of ExportToExcel Dialog.
									Call Fn_Button_Click("Fn_SE_ExportToExcel", objExportDialog, aButtons(1))
									Call Fn_WriteLogFile(Environment.Value("TestLogFile") , sURL &" URL value matches successfully.") 
								Else 
									Call Fn_WriteLogFile(Environment.Value("TestLogFile") , sURL &" URL value does not match.") 
									Fn_SE_ExportToExcel = False
									Set objExportDialog = nothing
									Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL : Function [ Fn_SE_ExportToExcel ] Invalid Option [ " & sAction & " ] ")   	
									Exit function
								End If
							End If
						End If
						Fn_SE_ExportToExcel = True
				'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				'As Column Configuration parameter is not there so [ sOutputTemplate ] parameter use to pass vlues wich has to verify in Case "VerifyColumnConfigurations"
				Case "VerifyColumnConfigurations"
							bReturn=False
							Call Fn_UI_JavaRadioButton_SetON("Fn_SE_ExportToExcel",objExportDialog, "ColumnConfigurations")
							arrColConfig=Split(sOutputTemplate,"~")
							For iCounter=0 To UBound(arrColConfig)
								bReturn=False
								bReturn=Fn_UI_ListItemExist("Fn_SE_ExportToExcel", objExportDialog, "ColumnConfig",arrColConfig(iCounter))
								If bReturn=False Then
									Exit For
								End If
							Next
							Call Fn_Button_Click("Fn_SE_ExportToExcel", objExportDialog, "Cancel")
							If bReturn Then
								Fn_SE_ExportToExcel = True
							End If
				Case Else
						Fn_SE_ExportToExcel = False
						Set objExportDialog = nothing
						Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL : Function [ Fn_SE_ExportToExcel ] Invalid Option [ " & sAction & " ] ")   	
						Exit function
		End Select
		Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS : Function [ Fn_SE_ExportToExcel ] executed successfully with case [ " & sAction & " ] ")   	
		Set objExportDialog = nothing
End Function
''*********************************************************		Function to Perform SE Panel operation in System Engineering	***********************************************************************
'
''Function Name		:				Fn_SE_TraceLinkOpeartions
'
''Description			 :		 		This function is used to get the SE Table Node Index.
'
''Parameters			   :				1.	sAction = "Select"
''												2.   sNodeName:Name of the Node. 
''												3. sNewName
''												4. sColName
''												5. sColValue
'			  										
''Return Value		   : 				True/ False
'
''Pre-requisite			:				System Engineering window should be displayed .
'
'
''Examples				:		  		Fn_SE_TraceLinkOpeartions("ComplyingTable:DeleteTraceLink","REQ-000148/A;1-Req2:Req2->Req3","","","")
'										Fn_SE_TraceLinkOpeartions("ComplyingTable:NodeVerify","REQ-000148/A;1-Req2:Req2->Req3","","","")
'										Fn_SE_TraceLinkOpeartions("ComplyingTable:Select","REQ-000148/A;1-Req2:Req2->Req3","","","")
'										Fn_SE_TraceLinkOpeartions("ComplyingTable:Expand","REQ-000148/A;1-Req2:Req2->Req3","","","")
'										Fn_SE_TraceLinkOpeartions("ComplyingTable:VerifyCellValue","REQ-000148/A;1-Req2:Req2->Req3","","Relation Type","Trace Link")
'									    Fn_SE_TraceLinkOpeartions("DefiningTable:GoToObject","000015/A;1-Paragraph:REQ-000001/A;1-Requiremnent","","","")
'
'									Note: - For [ PopupMenuSelect ] case use paramaeter [ sColValue ] to pass Popup menu name
'									bReturn=Fn_SE_TraceLinkOpeartions("ComplyingTable:PopupMenuSelect","REQ-000191-Requirement1"+":"+"Requirement1"+"->"+"Requirement2","","","Properties...")
'									bReturn=Fn_SE_TraceLinkOpeartions("ComplyingTable:PopupMenuSelect","REQ-000191-Requirement1"+":"+"Requirement1"+"->"+"Requirement2","","Type","Properties...")
''History:
''										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
''-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
''										Ketan Raje				21-Mar-2011			1.0											Harshal
''-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
''										Sonal P					18-Aug-2011			1.0					
''-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
''										Mohammad A				4-Jan-2012			1.0					Added code to handle Delete confirmation box
''-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
''										Sandeep N				10-Sep-2012			1.0					Modified case : Expand
''-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
''										Sonal P						19-Mar-2013			1.1					Added case : PopupMenuSelect
''-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SE_TraceLinkOpeartions(sAction,sNodeName,sNewName,sColName,sColValue)
	GBL_FAILED_FUNCTION_NAME="Fn_SE_TraceLinkOpeartions"
	On error resume next
	'------------------------------------------------------------------------------------------------------------------------------------------------------
	'Declare variables
	Dim dicTraceLinkObjects, aAction
	Fn_SE_TraceLinkOpeartions = False
	Set dicTraceLinkObjects=CreateObject("Scripting.Dictionary")
	'Map the paramenters of Fn_SE_TraceLinkOpeartions with Fn_SISW_SE_TraceLinkTabOperations
	aAction = Split(sAction,":")
	If aAction(0) = "DefiningTable" Then
		dicTraceLinkObjects("Defining Objects")=sNodeName
	Else
		dicTraceLinkObjects("Complying Objects")=sNodeName
	End If
	'Change Action name from VerifyCellValue to Verify
	If aAction(1) = "VerifyCellValue" Then
		aAction(1) = "Verify"
	ElseIf aAction(1) = "DeleteTraceLink" Then
		aAction(1) = "Delete Trace Link"
	End If
	'Set value to property field
	If sNewname <> "" Then
		dicTraceLinkObjects("PropertyValue") = sNewname
	End If
	'Set column names for verification
	If sColName <> "" Then
		dicTraceLinkObjects("ColumnNames") = sColName
	End If
	'Set column values for verification
	If sColValue <> "" Then
		dicTraceLinkObjects("ColumnValues") = sColValue
	End If
	'Call another function due to designed change	
	Fn_SE_TraceLinkOpeartions = Fn_SISW_SE_TraceLinkTabOperations(aAction(1), "Trace Links", "", "", dicTraceLinkObjects, "",sColValue)
	'---------------------------------------------------------------------------------------------------------------------------------------------------------
End Function
'*********************************************************		Function to perform operations on Open By Names dialogs ***********************************************************************
'Function Name		      :		Fn_SE_OpenByNameOperations

'Description			 	 :	    This function is used to perform operations on Open By Names dialog 
											
'Return Value		   : 	  True/False

'Pre-requisite			:	   System Engineering prespective should be open.

'Examples				:	   Dim dicOpenByName
'										Set dicOpenByName = CreateObject( "Scripting.Dictionary" )
'										With dicOpenByName  
'												.Add "SearchCriteria","Attributes"
'												.Add "Name", "TestItem"
'												.Add "ID", "000033"
'												.Add "Object", "000033-TestItem"  
'										End with
'										Call Fn_SE_OpenByNameOperations("OpenProduct", "FindAndDblClick", dicOpenByName)	
'										Call Fn_SE_OpenByNameOperations("OpenProduct", "FindAndVerify", dicOpenByName)	
'History:
'								Developer Name			Date				Rev. No.			Changes Done			
'---------------------------------------------------------------------------------------------------------------------------
'								 Ketan Raje				12-May-2011	          1.0													
'								 Anjali M				05-Feb-2013	          1.1			Added function call Fn_SyncTCObjects()	before setting Name												
'---------------------------------------------------------------------------------------------------------------------------
Public function Fn_SE_OpenByNameOperations(sAction, sSubAction, dicOpenByName)
		GBL_FAILED_FUNCTION_NAME="Fn_SE_OpenByNameOperations"
		Dim bReturn, objOpenByName, iRowCount, iCount, objTable
		Dim sItem
		Fn_SE_OpenByNameOperations = False

		Select Case sAction
					Case "OpenProduct"
								Set objOpenByName = JavaWindow("SystemsEngineering").JavaWindow("SEWindow").JavaDialog("OpenProduct")										
								If Fn_UI_ObjectExist("Fn_SE_OpenByNameOperations", objOpenByName) = False Then
									'Open the Open Product window.
									Call Fn_ToolBarOperation("ShowDropdownAndSelect", "Open By Name", "Open Product")
								End If
								Select Case sSubAction
										Case "FindAndDblClick", "FindAndVerify"
												'Select Seacrh criteria.
												If dicOpenByName("SearchCriteria") <> "" Then
													If Trim(Lcase(dicOpenByName("SearchCriteria"))) = "association" Then
														Call Fn_UI_JavaRadioButton_SetON("Fn_SE_OpenByNameOperations",objOpenByName, "Association")
													ElseIf Trim(Lcase(dicOpenByName("SearchCriteria"))) = "attributes" Then
														Call Fn_UI_JavaRadioButton_SetON("Fn_SE_OpenByNameOperations",objOpenByName, "Attributes")
													End If
												End If
												' typeing value in Name edit box
												If dicOpenByName("Name") <> ""  Then
														Call Fn_SyncTCObjects()
														objOpenByName.JavaEdit("Name").Set ""
														Call Fn_Edit_Box("Fn_SE_OpenByNameOperations", objOpenByName,"Name",dicOpenByName("Name"))
												End If
												' typeing value in Id edit box
												If dicOpenByName("ID") <> ""  Then
														objOpenByName.JavaEdit("ID").Set ""
														Call Fn_Edit_Box("Fn_SE_OpenByNameOperations", objOpenByName,"ID",dicOpenByName("ID"))
												End If
												' clicking on Search button
												objOpenByName.JavaButton("Search").DblClick  1,1,"LEFT"
												Call Fn_ReadyStatusSync(2)
												Set objTable = objOpenByName.JavaTable("OpenProduct")
												iRowCount = cint(objTable.GetROProperty("rows"))
												For iCount = 0 to iRowCount -1
															If Trim(Lcase(objTable.GetCellData(iCount,"Object"))) = Trim(Lcase(dicOpenByName("Object")))  Then
																 Call Fn_UI_JavaTable_SelectRow("Fn_SE_OpenByNameOperations", objOpenByName,"OpenProduct",iCount)
																Exit for
															End If
												Next
												iCount = cint(objTable.GetROProperty("SelectedRow"))
												If iCount <> -1 Then
													If sSubAction = "FindAndDblClick" Then
														Wait(5)
														objTable.DoubleClickCell iCount,"Object","LEFT"
														Call Fn_ReadyStatusSync(2)
													End If
												Else
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SE_OpenByNameOperations ] Case [ " & sAction  & " ] No Item is selected.")
													 Fn_SE_OpenByNameOperations = False
													 If objOpenByName.Exist Then
														 objOpenByName.Close
													 End If
													Exit function
												End If
												Fn_SE_OpenByNameOperations = True
												 If objOpenByName.Exist Then
													 objOpenByName.Close
												 End If
								End Select
		End Select

		Set objOpenByName = nothing
		Set objTable = nothing
		Fn_SE_OpenByNameOperations = True
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SE_OpenByNameOperations ] executed successfully with case [ " & sAction &" ].")

End Function
'-------------------------------------------------------------------------This Function is used to to perform operation on MS Word tab---------------------------------------------------------------
'Function Name		:	Fn_SE_MSWordTabOperationsExt

'Description			:	This Function is used to to perform operation on MS Word tab

'Parameters			:			1.	strAction:Action Name
'											 2.	 strValue:Value to set in text Box Or to verify the value
											'3.	strParameterName:Parameter Name

'Return Value		:	True/False

'Pre-requisite		:	Object should be selected
'											
'Examples			:	Case "VerifyValue" : Fn_SE_MSWordTabOperationsExt("VerifyValue","value1","parameter1")
'							  Case "SetvalueWithoutSave" : Call Fn_SE_MSWordTabOperationsExt("SetvalueWithoutSave","23","temp")
'									
'History:
'									Developer Name												Date								Rev. No.						Changes Done				Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'									Sushma Pagare									   12/5/2010			                  1.0										Created							Harshal Agrawal
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SE_MSWordTabOperationsExt(strAction,strValue,strParameterName)
	GBL_FAILED_FUNCTION_NAME="Fn_SE_MSWordTabOperationsExt"
	'Variable declaration
   Dim bReturn,bFlag,iRowCnt,sCellData,sValueName,iCount
   'Function Return False
   Fn_SE_MSWordTabOperationsExt=False
   bReturn=False
   Call Fn_SetView("Teamcenter:MS Word")

	Select Case strAction
		Case "VerifyValue"	'Fn_SE_MSWordTabOperationsExt("VerifyValue","value1","parameter1")
				'Taking no of rows from ParametricValues table
				iRowCnt=Fn_UI_Object_GetROProperty("Fn_SE_MSWordTabOperationsExt",JavaWindow("MyTeamcenter").JavaTable("ParametricValues"),"rows")
				For iCount=0 To iRowCnt-1
						'Taking data from ParametricValues table
						sCellData=Fn_UI_JavaTable_GetCellData("Fn_SE_MSWordTabOperationsExt",JavaWindow("MyTeamcenter"), "ParametricValues",iCount,0)
						If strParameterName=sCellData Then
								sValueName=Fn_UI_JavaTable_GetCellData("Fn_SE_MSWordTabOperationsExt",JavaWindow("MyTeamcenter"), "ParametricValues",iCount,1)
								If strValue=sValueName Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Pass:Successfully verify  value"& strValue )   
									bReturn=True
									Exit For
								End If
						End If
				Next
				If bReturn=False Then
                    Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Fail:"& strParameterName & "Parameter is not present")   
					Exit Function
				Else
					  Fn_SE_MSWordTabOperationsExt=True
				End If
				wait(2)
			'	JavaWindow("MyTeamcenter").JavaTab("MSWord").CloseTab "MS Word"	
		Case "SetvalueWithoutSave","SetvalueAndSave"	'Fn_SE_MSWordTabOperationsExt("VerifyValue","value1","parameter1")
				'Taking no of rows from ParametricValues table
				iRowCnt=Fn_UI_Object_GetROProperty("Fn_SE_MSWordTabOperationsExt",JavaWindow("MyTeamcenter").JavaTable("ParametricValues"),"rows")
				For iCount=0 To iRowCnt-1
						'Taking data from ParametricValues table
						sCellData=Fn_UI_JavaTable_GetCellData("Fn_SE_MSWordTabOperationsExt",JavaWindow("MyTeamcenter"), "ParametricValues",iCount,0)
						If Trim(Lcase(strParameterName)) = Trim(Lcase(sCellData)) Then								
							'JavaWindow("RequirementsManager").JavaTable("Parametric Values:").SelectCell iCount,"VALUE"
							Call Fn_UI_JavaTable_SelectCell("Fn_SE_MSWordTabOperationsExt", JavaWindow("MyTeamcenter"), "ParametricValues",iCount,"VALUE")
'							 JavaWindow("MyTeamcenter").JavaTable("ParametricValues").DoubleClickCell iCount,"VALUE"
							JavaWindow("MyTeamcenter").JavaTable("ParametricValues").ActivateCell iCount,"VALUE"
							iLen = len(strValue)
								Set WshShell = CreateObject("WScript.Shell")
									For i = 1 to iLen
										WshShell.SendKeys mid(strValue,i,1)
									Next
								Set WshShell = Nothing
							bReturn = True
							Exit For
						End If
				Next
				If strAction="SetvalueAndSave" Then
					'Saving the changes
					JavaWindow("MyTeamcenter").JavaToolbar("ParametricValueToolBar").Press "Save"
				End If
				If bReturn = False Then
                    Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Fail:"& strParameterName & "Parameter is not present")   
					Exit Function
				Else
					  Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Pass:Successfully Set value"& strValue )   
					  Fn_SE_MSWordTabOperationsExt=True
				End If
	End Select
End Function
'-------------------------------------------------------------------------This Function is used to to perform operation on MS Word tab---------------------------------------------------------------
'Function Name		:	Fn_SE_ItemFromTemplateOperations

'Description			:	This Function is used to to perform operation on MS Word tab

'Parameters			:			StrAction,dicItemFromTemplate

'Return Value		:	True/False

'Pre-requisite		:	Object should be selected
'											
'History:
'									Developer Name												Date								Rev. No.						Changes Done				Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'							Sonal P								   17/5/2010			                  1.0										Created							Harshal Agrawal
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function Fn_SE_ItemFromTemplateOperations(StrAction,dicItemFromTemplate)
	GBL_FAILED_FUNCTION_NAME="Fn_SE_ItemFromTemplateOperations"
   Dim objTemplt
   Dim itemID,itemRev
   Set objTemplt=JavaWindow("SystemsEngineering").JavaWindow("SEWindow").JavaDialog("NewItemFromTemplate")
   If Not objTemplt.Exist(6) Then
		'Invoking "New Item From Template" Window
		Call Fn_MenuOperation("Select","File:New:Item From Template...")
   End If
	Select Case StrAction
		Case "Add"
			If dicItemFromTemplate("TemplateID")<>"" Then
				'Setting Template ID
				Call Fn_UI_EditBox_Type("Fn_SE_ItemFromTemplateOperations",objTemplt,"TemplateID",dicItemFromTemplate("TemplateID"))
			End If
			If dicItemFromTemplate("ItemID")<>"" Then
				'Setting Item ID
				Call Fn_UI_EditBox_Type("Fn_SE_ItemFromTemplateOperations",objTemplt,"ItemID",dicItemFromTemplate("ItemID"))
			End If
			If dicItemFromTemplate("Revision")<>"" Then
				'Setting Revision
				Call Fn_UI_EditBox_Type("Fn_SE_ItemFromTemplateOperations",objTemplt,"Revision",dicItemFromTemplate("Revision"))
			End If
			
			If dicItemFromTemplate("ItemID")="" OR dicItemFromTemplate("Revision")="" Then
				'Clicking Assign Button To Assign Item ID And Revision
				Call Fn_Button_Click("Fn_SE_ItemFromTemplateOperations",objTemplt,"Assign")
			End If
			If dicItemFromTemplate("Name")<>"" Then
				'Setting Revision
				Call Fn_UI_EditBox_Type("Fn_SE_ItemFromTemplateOperations",objTemplt,"Name",dicItemFromTemplate("Name"))
			End If
			If dicItemFromTemplate("NumberOfObject")<>"" Then
				'Setting Number Of Objects
				Call Fn_UI_EditBox_Type("Fn_SE_ItemFromTemplateOperations",objTemplt,"NumberOfObjects",dicItemFromTemplate("NumberOfObject"))
			End If
			If dicItemFromTemplate("Description")<>"" Then
				'SettingDescription
				Call Fn_UI_EditBox_Type("Fn_SE_ItemFromTemplateOperations",objTemplt,"Description",dicItemFromTemplate("Description"))
			End If
			If dicItemFromTemplate("RootItem")<>"" Then
				'Setting Set As New Root Item Option
				Call  Fn_CheckBox_Set("Fn_SE_ItemFromTemplateOperations", objTemplt, "ShowAsNewRootItem", dicItemFromTemplate("RootItem"))
			End If
			itemID=Fn_Edit_Box_GetValue("Fn_SE_ItemFromTemplateOperations",objTemplt,"ItemID")
			itemRev=Fn_Edit_Box_GetValue("Fn_SE_ItemFromTemplateOperations",objTemplt,"Revision")
			Call Fn_Button_Click("Fn_SE_ItemFromTemplateOperations",objTemplt,"OK")
			wait(3)
			Fn_SE_ItemFromTemplateOperations=itemID+"/"+itemRev
	End Select
End Function
'*********************************************************		Function to handle error dialog / window in SE ***********************************************************************
'Function Name		:				Fn_SE_ErrorMessageVerify

'Description			 :		 		 This function is used to handle error dialog / window in System Engineering

'Parameters			   :	 			1. sAction: Select the Requirement Spec
'									    2. sTitle: error dialog / window title
'									    3. sErrorMessage: error message to verify
'									    4. sBtnName: Name of the button

'Return Value		   : 				TRUE \ FALSE

'Pre-requisite			:		 		should be in SE Prespective

'Examples				:				 Call Fn_SE_ErrorMessageVerify("VerifyErrorWindow","Teamcenter","Do you want to save your modifications to the Rich Content?","")

'History:
'										Developer Name			Date			Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Koustubh					19.05.2011		1.0					New
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public function Fn_SE_ErrorMessageVerify(sAction, sTitle, sErrorMessage, sBtnName)
	Dim dicErrorInfo
	Set dicErrorInfo = CreateObject("Scripting.Dictionary")
	dicErrorInfo.Add "Action", sAction
	dicErrorInfo.Add "Title", sTitle
	dicErrorInfo.Add "Message", sErrorMessage
	dicErrorInfo.Add "Button", sBtnName    
	Fn_SE_ErrorMessageVerify = Fn_SISW_SE_ErrorVerify(dicErrorInfo)
	
End Function
'*********************************************************		Function to perform operations on Right Panel Tabsin SE ***********************************************************************
'Function Name		:				Fn_SE_RightPanelTabOperations

'Description			 :		 		 This function is used to perform operations on Right Panel Tabs

'Parameters			   :	 			1. sAction: Action to perform on Tab
'									    2. sTab: Tab name
'									    3. sMenu: for future use.

'Return Value		   : 				TRUE \ FALSE

'Pre-requisite			:		 		should be in SE Prespective

'Examples				:				 msgbox  Fn_SE_RightPanelTabOperations("Select", "Attachments", "")
'Examples				:				 msgbox  Fn_SE_RightPanelTabOperations("DoubleClick", "Attachments", "")
'Examples				:				 msgbox  Fn_SE_RightPanelTabOperations("Close", "Attachments", "")

'History:
'										Developer Name			Date			Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Koustubh					24.05.2011		1.0					New
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public function Fn_SE_RightPanelTabOperations(sAction, sTab, sMenu)
		GBL_FAILED_FUNCTION_NAME="Fn_SE_RightPanelTabOperations"
		Dim objTabWidget, objTab, iCounter, iTabsCount, iXposition, iYposition, sBounds, aBounds
		Dim iIndexCounter,icount,iItemCount,iIndex,X,H,sxLen,syLen
	
		Fn_SE_RightPanelTabOperations = False
		iXposition = 0
		iIndexCounter = 0
		bFlag = False
'		Set objTabWidget = JavaWindow("SystemsEngineering").JavaObject("SEComponentTab")
		
		Set objSelectType = description.Create()
		objSelectType("Class Name").value = "JavaTab"
		objSelectType("toolkit class").value = "org.eclipse.swt.custom.CTabFolder"						
		Set  objIntNoOfObjects = JavaWindow("DefaultWindow").ChildObjects(objSelectType)
		
		Select Case sAction
				' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				Case "Select"    ' [TC12 - 20172600-11_8_2017-JotibaT-Maintenace]- Added changes as per object change from JavaObject to JavaTab
					For iIndexCounter = 0 To objIntNoOfObjects.Count-1 Step 1			
						iItemCount = cInt(objIntNoOfObjects(iIndexCounter).Object.getItemCount())
						For iCounter = 0 To iItemCount- 1 Step 1
							If trim(sTab) = trim(objIntNoOfObjects(iIndexCounter).Object.getItems().mic_arr_get(iCounter).getText()) Then
								objIntNoOfObjects(iIndexCounter).Select sTab
								bFlag=True
								Exit For 
							End IF
						Next
						If bFlag=True Then Exit For 
					Next
'					For iIndexCounter = 0 to 2 
'						objTabWidget.setTOProperty "Index", iIndexCounter
'						iTabsCount = cInt(objTabWidget.Object.getTabItemCount)
'						For iCounter = 0 to iTabsCount -1
'								Set objTab = objTabWidget.Object.getItem(iCounter)
'								iXposition = iXposition + (objTab.getWidth)
'								If  trim(objTab.Text) = sTab Then
'										iXposition = iXposition - (objTab.getWidth/2)
'										objTabWidget.Click iXposition, (objTab.getHeight/2), "LEFT"
'										Fn_SE_RightPanelTabOperations = True
'										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_SE_RightPanelTabOperations : Successfully clicked on [ " & sTab & " ] tab.")
'										bFlag = True
'										Exit for
'								End If
'						Next
'						If bFlag = True Then exit for
'					Next
					
				' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				Case "DoubleClick"
'						For iIndexCounter = 0 to 2 
'								objTabWidget.setTOProperty "Index", iIndexCounter
'								iTabsCount = cInt(objTabWidget.Object.getTabItemCount)
'								For iCounter = 0 to iTabsCount -1
'										Set objTab = objTabWidget.Object.getItem(iCounter)
'										iXposition = iXposition + (objTab.getWidth)
'										If  trim(objTab.Text) = sTab Then
'												iXposition = iXposition - (objTab.getWidth/2)
'												objTabWidget.DblClick iXposition, (objTab.getHeight/2), "LEFT"
'												Fn_SE_RightPanelTabOperations = True
'												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_SE_RightPanelTabOperations : Successfully double clicked on [ " & sTab & " ] tab.")
'												bFlag = true
'												Exit for
'										End If
'								Next
'								If bFlag = true Then exit for
'						Next
'						
						For icount = 0 To objIntNoOfObjects.Count-1 Step 1			
						iItemCount = cInt(objIntNoOfObjects(icount).Object.getItemCount())
						For iCounter = 0 To iItemCount- 1 Step 1
							If trim(sTab) = trim(objIntNoOfObjects(icount).Object.getItems().mic_arr_get(iCounter).getText()) Then
								objIntNoOfObjects(icount).Select sTab
								iIndex=objIntNoOfObjects(icount).Object.getSelectionIndex
								Set objItem=objIntNoOfObjects(icount).Object.getItem(iIndex)
										sBounds = objItem.getBounds().toString()
										sBounds = mid(sBounds,instr(sBounds,"{")+1, len(sBounds) -instr(sBounds,"{")-1)
										aBounds = split(sBounds,",")
										X = cInt(trim(aBounds(0)))
										H = cInt(trim(aBounds(3)))
										sxLen = X + 15
										syLen = (H/2)
									objIntNoOfObjects(icount).DblClick sxLen,syLen,"LEFT"
									wait 2
									bFlag=True
									Exit For 
							End IF
						Next
						If bFlag=True Then Exit For 
					Next
				' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				Case "Close" '' [TC12 - 20172600-11_8_2017-JotibaT-Maintenace]- Added changes as per object change from JavaObject to JavaTab
						For iIndexCounter = 0 To objIntNoOfObjects.Count-1 Step 1			
							iItemCount = cInt(objIntNoOfObjects(iIndexCounter).Object.getItemCount())
							For iCounter = 0 To iItemCount- 1 Step 1
								If trim(sTab) = trim(objIntNoOfObjects(iIndexCounter).Object.getItems().mic_arr_get(iCounter).getText()) Then
									objIntNoOfObjects(iIndexCounter).Select sTab
									wait 1
									objIntNoOfObjects(iIndexCounter).CloseTab sTab
									bFlag=True
									Exit For 
								End IF
							Next
							If bFlag=True Then Exit For 
						Next
	
'						For iIndexCounter = 0 to 2 
'								objTabWidget.setTOProperty "Index", iIndexCounter
'								iTabsCount = cInt(objTabWidget.Object.getTabItemCount)
'								For iCounter = 0 to iTabsCount -1
'										Set objTab = objTabWidget.Object.getItem(iCounter)
'										iXposition = iXposition + (objTab.getWidth)
'										If  trim(objTab.Text) = sTab Then
'												' selecting tab
'												iXposition = iXposition - (objTab.getWidth/2)
'												objTabWidget.Click iXposition, (objTab.getHeight/2), "LEFT"
'												' fetching coordinates of close button
'												sBounds = objTab.getCloseButtonBounds.toString()
'												sBounds = right(sBounds, Len(sBounds)-instr(sBounds, "{"))
'												aBounds = split(sBounds, ",", -1, 1)
'												iXposition = Cint(trim(aBounds(0))) + 5
'												iYposition = Cint(trim(aBounds(1))) + 5
'												' clicking on close X 
'												objTabWidget.Click iXposition, iYposition, "LEFT"
'												Fn_SE_RightPanelTabOperations = True
'												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_SE_RightPanelTabOperations : Successfully Double clicked on [ " & sTab & " ] tab.")
'												bFlag = True
'												Exit for
'										End If
'								Next
'								If bFlag = True Then exit for
'						Next
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				Case Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_SE_RightPanelTabOperations : Invalid case [ " & sAction & " ].")
				' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		End Select
		If bFlag=True Then
			Fn_SE_RightPanelTabOperations=True
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_SE_RightPanelTabOperations : executed successfully with case [ " & sAction & " ].")
		End If
		Set objTabWidget = nothing
End Function
 '*********************************************************		Function to Create Specification in SE ***********************************************************************
'Function Name		:				Fn_SE_CustomNoteCreate

'Description			 :		 		 This function is used to Create the Custom Note in System Engineering

'Parameters			   :	 			1. strNodeName: Select the Note Spec
'													2. strCustID: ID of the Specification
'												   3. strCustRev: Revision of the Spec
'												  4. strCusrName: Name of the Spec
'												  5.strCustDesc: Description of the Spec

'Return Value		   : 				TRUE \ FALSE

'Pre-requisite			:		 		should be in SE Prespective

'Examples				:				 Call Fn_SE_CustomNoteCreate("Custom Note","CSTMNOTE-000001","","CustNoteName","Description")

'History:
'										Developer Name			Date			Rev. No.			Reviewer			Build			Changes Done
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Sushma P.		
'	Modified By 				Ketan Raje					12/09/2011		1.0					Harshal A.			20110824		Handled JavaEdit by using descriptive code.	
'                                        Sanjeet k.                         05/01/12                                                                 9.1-1219                         Handled JavaEdit by Set method
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function Fn_SE_CustomNoteCreate(strNodeName,strCustID,strCustRev,strCustName,strCustDesc)
	GBL_FAILED_FUNCTION_NAME="Fn_SE_CustomNoteCreate"
	Dim strNodePath,intNodeCount,sTreeItem,intCount,strNodePathC,objCustNote
	Dim ObjSpecWnd,bFlag, objElement, intNoOfObjects, innerCntr, intCnt

		Fn_SE_CustomNoteCreate=False
		set objCustNote = Fn_SISW_SE_GetObject("NewCustomNote")
		sNewMenu= Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("RAC_Menu"),"NewCustomNote")
		
	'Verifying "New Specification" window's existance
	If Fn_UI_ObjectExist("Fn_SE_CustomNoteCreate",objCustNote)=False Then
		'Invoking "New Specification" Window
		Call Fn_MenuOperation("Select",sNewMenu)
		Call Fn_ReadyStatusSync(1)
	End If
	
	'Creating Object of "New Specification" window
	Set ObjSpecWnd=Fn_UI_ObjectCreate("Fn_SE_CustomNoteCreate",objCustNote)
	'Check for existance of Custom Note Tree 
	'If Fn_UI_ObjectExist("Fn_SE_CustomNoteCreate",ObjSpecWnd.JavaTree("CustomNoteTree")) Then
	If Fn_SISW_UI_Object_Operations("Fn_SE_CustomNoteCreate", "Exist", ObjSpecWnd.JavaTree("CustomNoteTree"), SISW_MICRO_TIMEOUT) then
	    Call Fn_UI_JavaTree_Expand("Fn_SE_CustomNoteCreate", ObjSpecWnd, "CustomNoteTree","Complete List")
		JavaWindow("SystemsEngineering").JavaWindow("NewCustomNote").JavaTree("CustomNoteTree").WaitProperty "items count" , micGreaterThan(1)
		If Fn_UI_JavaTree_NodeExist("Fn_SE_NoteSpecCreate",ObjSpecWnd.JavaTree("CustomNoteTree"),"Complete List:"+strNodeName) Then
				strNodePathC="Complete List:"+strNodeName
		Else
				strNodePathC="Most Recently Used:"+strNodeName
		End If
		Call Fn_JavaTree_Select("Fn_SE_CustomNoteCreate", ObjSpecWnd, "CustomNoteTree",strNodePathC)
		Call Fn_JavaTree_Select("Fn_SE_CustomNoteCreate", ObjSpecWnd, "CustomNoteTree","Complete List")
	    Call Fn_JavaTree_Select("Fn_SE_CustomNoteCreate", ObjSpecWnd, "CustomNoteTree",strNodePathC)
		Call Fn_Button_Click("Fn_SE_CustomNoteCreate",ObjSpecWnd,"Next")
	End If
	'Setting Id
	If strCustID<>"" Then
'			Set objElement = Description.Create()
'			objElement("Class Name").value = "JavaEdit"														
'			Set  intNoOfObjects = ObjSpecWnd.JavaObject("Composite").ChildObjects(objElement)
'					intNoOfObjects(0).Type strCustID
                			  Wait 5
					    Call Fn_ReadyStatusSync(2)
			Call Fn_Edit_Box("Fn_SE_CustomNoteCreate",ObjSpecWnd,"ID",strCustID)
	End If
	'Setting Revision
	If strCustRev<>"" Then
'			Set objElement = Description.Create()
'			objElement("Class Name").value = "JavaEdit"											
'			Set  intNoOfObjects = ObjSpecWnd.JavaObject("Composite").ChildObjects(objElement)
'					intNoOfObjects(1).Type strCustRev
				       Call Fn_ReadyStatusSync(2)
			Call Fn_Edit_Box("Fn_SE_CustomNoteCreate",ObjSpecWnd,"RevisionName",strCustRev)
	End If
	'Setting Name
	If strCustName <> "" Then
'			Set objElement = Description.Create()
'			objElement("Class Name").value = "JavaEdit"						
'			Set  intNoOfObjects = ObjSpecWnd.JavaObject("Composite").ChildObjects(objElement)
'					intNoOfObjects(2).Type strCustName
				       Call Fn_ReadyStatusSync(2)
			Call Fn_Edit_Box("Fn_SE_CustomNoteCreate",ObjSpecWnd,"Name",strCustName)
	End If
	'Setting Description
	If strCustDesc <> "" Then
'			Set objElement = Description.Create()
'			objElement("Class Name").value = "JavaEdit"											
'			Set  intNoOfObjects = ObjSpecWnd.JavaObject("Composite").ChildObjects(objElement)
'					intNoOfObjects(3).Type strCustDesc
				       Call Fn_ReadyStatusSync(2)
			Call Fn_Edit_Box("Fn_SE_CustomNoteCreate",ObjSpecWnd,"Description",strCustDesc)
	End If
	'Clicking On Finish Button To finish the Operation
	Call Fn_Button_Click("Fn_SE_CustomNoteCreate",ObjSpecWnd,"Finish")
    Call Fn_ReadyStatusSync(5)
	Call Fn_Button_Click("Fn_SE_CustomNoteCreate",ObjSpecWnd,"Cancel")
	'function Return True
	Fn_SE_CustomNoteCreate=True
	'Releasing "New Specification" window's object
	Set ObjChangeWnd=Nothing
	set objCustNote = Nothing
End Function
'****************************************		Function to perform operations on Attachments Table.  ***************************************
'Function Name		      :			  Fn_SE_AttachmentTableNodeOperation  

'Description			     :  	      Function to perform operations on Attachments Table.

'Parameters			   		:	    1.  String - sAction
'						          		  2.  String - sAttchLine
' 							      	  3.  String - sColumn
'									  4.  String - sData - For future use
'									  5.  String - sMenu - For future use
											
'Return Value		       : 			 True / False

'Pre-requisite			    :		 	  should be in SE Prespective, Attachment window should be displayed .
'								   		

'Examples				    :			'Call  Fn_SE_AttachmentTableNodeOperation("Exist", "CN0004/A;1-Ic1", "", "", "")
										'Call  Fn_SE_AttachmentTableNodeOperation("Expand", "CN0004/A;1-Ic1", "", "", "")
										'Call  Fn_SE_AttachmentTableNodeOperation("Select", "CN0004/A;1-Ic1:Addressed By", "", "", "")
										'Call Fn_SE_AttachmentTableNodeOperation("DoubleClick", "CN0004/A;1-Ic", "", "", "")
										'Call Fn_SE_AttachmentTableNodeOperation("DoubleClick", "CN0004/A;1-Ic", "Item Type", "", "")
										'Call Fn_SE_AttachmentTableNodeOperation("GetCellData", "REQ-123458/A;1-Req3:REQ-123458/A", "Relation","Item Masters", "")
										'Call Fn_SE_AttachmentTableNodeOperation("VerifyColumn","","Relation~Description", "", "")
										'Call Fn_SE_AttachmentTableNodeOperation("AddColumn","","Acl Bits~Archive Date", "", "")
'History:
'						Developer Name			Date				Rev. No.			Changes Done							Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'					 	   Koustubh W			  25-May-2011		1.0
'					 	   Sandeep N			  18-Oct-2011		  1.1				Added Case "GetCellData"			Sunny R
'					 	   Sandeep N			  18-Oct-2011		  1.2				Added Case "VerifyColumn"			Sunny R
'																											  Added Case "AddColumn"
'					 	   Sandeep N			  17-Jul-2012		  1.3			Modified Case : Select 							Anjali M
'																											  Added new JavaApplet  Hierarchy
'																											Old : JavaWindow("SystemsEngineering").JavaWindow("WEmbeddedFrame")
'																											New : Window("SEWindow").JavaWindow("WEmbeddedFrame")
'					 	   Sandeep N			  18-Jul-2012		  1.4			Modified function to handle Attachement tab Table and BOM table wich opens in Nav tree
'					 	   Sandeep N			  27-Jun-2013		  1.5		Added Case : PopupMenuSelect
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public function Fn_SE_AttachmentTableNodeOperation(sAction, sAttchLine, sColumn, sData, sMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_SE_AttachmentTableNodeOperation"
	Dim objTable, iRowCnt, iRowNum, aSubElement, jCounter, bFlag
	Dim iInstanceCnt, aNode, sNode, iCnt, iColNum, bColumnFlag , arrNode, strNodeName1
	Dim iCol,aColumns,currColName,iCount,iColumnCnt,objChangeCol
	Dim objTabFld,i,StrTabName,iRowNo,bFlag1
	Dim iColNo,iStart,StrMenu
	bFlag = False
	bFlag1=False

'	Set objTabFld = JavaWindow("SystemsEngineering").JavaObject("RACTabFolderWidget")
'	i = objTabFld.Object.getSelectedTabIndex
'	StrTabName=objTabFld.Object.getItem(i).text()

    '- - - - - - - - - - - - - - - -  Modified by Chaitali R.- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	If JavaWindow("SystemsEngineering").JavaTab("RACTabFolderWidget").Exist(2) Then
		StrTabName=JavaWindow("SystemsEngineering").JavaTab("RACTabFolderWidget").GetROProperty("value")
	Else
		JavaWindow("SystemsEngineering").JavaTab("RACTabFolderWidget").SetTOProperty "orig_logical_location", "X_BIG__Y_SMALL" 
	 	If JavaWindow("SystemsEngineering").JavaTab("RACTabFolderWidget").Exist(2) Then
	 		StrTabName=JavaWindow("SystemsEngineering").JavaTab("RACTabFolderWidget").GetROProperty("value")
	 	Else
	 		JavaWindow("SystemsEngineering").JavaTab("RACTabFolderWidget").SetTOProperty "orig_logical_location", "X_BIG__Y_FULL" 
	 		StrTabName=Fn_UI_Object_GetROProperty("Fn_SignalBasicCreate", JavaWindow("SystemsEngineering").JavaTab("RACTabFolderWidget"),"value")
	 		'StrTabName=JavaWindow("SystemsEngineering").JavaTab("RACTabFolderWidget").GetROProperty("value")
	 	End If
	 	JavaWindow("SystemsEngineering").JavaTab("RACTabFolderWidget").SetTOProperty "orig_logical_location", "X_UNK__Y_FULL" 
	End If
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	For iCounter=0 to 12
		Window("SEWindow").JavaWindow("WEmbeddedFrame").SetTOProperty "index",iCounter
		If Window("SEWindow").JavaWindow("WEmbeddedFrame").JavaTable("CMEBOMTreeTable").Exist(2) Then
			if StrTabName="Attachments" then
				If Window("SEWindow").JavaWindow("WEmbeddedFrame").JavaTable("CMEBOMTreeTable").GetROProperty("tagname")="CMEBOMTreeTable" Then
					bFlag=true
					Exit for
				ElseIf Window("SEWindow").JavaWindow("WEmbeddedFrame").JavaTable("CMEBOMTreeTable").GetROProperty("tagname")="AttachmentsTreeTable" Then
					bFlag=true
					Exit for	
				End If
			Else
				If Window("SEWindow").JavaWindow("WEmbeddedFrame").JavaTable("CMEBOMTreeTable").GetROProperty("tagname")="AttachmentsTreeTable" Then
					bFlag=true
					Exit for
				else
					Set objTable =  Window("SEWindow").JavaWindow("WEmbeddedFrame").JavaTable("CMEBOMTreeTable")
					aSubElement = Split(sAttchLine, ":", -1, 1)
					jCounter = 0
					iRowCnt = cInt(objTable.GetROProperty("rows"))
					sNode = ""
					For iRowNum = 0 to iRowCnt -1
							If sNode = "" then
								iInstanceCnt = 1
								aNode = split(Trim(aSubElement(jCounter)),"@")
								If UBound(aNode) = 1 Then
									iInstanceCnt = cInt(aNode(1))
								End If
								sNode = trim( aNode(0) )
							End if
							If Trim(objTable.Object.getValueAt(iRowNum, 0).toString() ) = sNode Then
								iInstanceCnt = iInstanceCnt - 1
								If iInstanceCnt = 0 Then
								If jCounter = UBound(aSubElement) Then
									bFlag1 = True
									Exit for
								End if
								jCounter = jCounter + 1
								sNode = ""
							End If
						End If
					Next
					If bFlag1 = True Then
						bFlag1=False
						bFlag = True
						Exit for
					End if
				End If
			End if
		End if
	Next
	If bFlag=false Then
		Exit function
	End If
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	'Set objTable = JavaWindow("SystemsEngineering").JavaWindow("WEmbeddedFrame").JavaTable("CMEBOMTreeTable") Swapnil :Parent class of javaApplet changed.
	Set objTable =  Window("SEWindow").JavaWindow("WEmbeddedFrame").JavaTable("CMEBOMTreeTable")
	Fn_SE_AttachmentTableNodeOperation = False
	bFlag = False
	bColumnFlag = False
	' fetching row number
	If sAttchLine <> "" Then
		If bFlag1=True Then
			bFlag=True
			iRowNum=iRowNo
		else
			aSubElement = Split(sAttchLine, ":", -1, 1)
			jCounter = 0
			iRowCnt = cInt(objTable.GetROProperty("rows"))
			sNode = ""
			For iRowNum = 0 to iRowCnt -1
				' For the Node Hierarchy of an Element
				If sNode = "" then
					iInstanceCnt = 1
					aNode = split(Trim(aSubElement(jCounter)),"@")
					If UBound(aNode) = 1 Then
						iInstanceCnt = cInt(aNode(1))
					End If
					sNode = trim( aNode(0) )
				end if
				If Trim(objTable.Object.getValueAt(iRowNum, 0).toString() ) = sNode Then
					iInstanceCnt = iInstanceCnt - 1
					If iInstanceCnt = 0 Then
						' For Last Node of an Element
						If jCounter = UBound(aSubElement) Then
											bFlag = True
							Exit for
						end if
						jCounter = jCounter + 1
						sNode = ""
					End If
				End If
			Next
		end if
	End If

	If InStr(1,sColumn,"~") Then
		sColumn=sColumn
	Else
		' fetching column number
		If sColumn = "" Then sColumn = "Line"
		iCnt = cInt(objTable.GetROProperty("cols"))
		For iColNum = 0 to iCnt -1
			If Trim(objTable.GetColumnName(iColNum)) = sColumn Then
				bColumnFlag = True
				Exit for
			End If
		Next
	End If

	Select Case sAction
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "Select"
				If bFlag Then
					Call Fn_UI_JavaTable_SelectRow("Fn_SE_AttachmentTableNodeOperation", Window("SEWindow").JavaWindow("WEmbeddedFrame"),"CMEBOMTreeTable",iRowNum) 
					Fn_SE_AttachmentTableNodeOperation = True
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_SE_AttachmentTableNodeOperation : Failed to find Row [ " & sAttchLine & " ].")
				End IF
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "Expand"
			If bFlag Then 
				objTable.Object.expandNode(objTable.Object.getNodeForRow(iRowNum))
				Fn_SE_AttachmentTableNodeOperation = True
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_SE_AttachmentTableNodeOperation : Failed to find Row [ " & sAttchLine & " ].")
			End IF
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "Exist", "Exists"
				If bFlag Then Fn_SE_AttachmentTableNodeOperation = True
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_SE_AttachmentTableNodeOperation : Successfully executed case [ " & sAction & " ] for Row [ " & sAttchLine & " ].")
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "DoubleClick"
			If bFlag Then
				If bColumnFlag Then
					Call Fn_UI_JavaTable_DoubleClickCell("Fn_SE_AttachmentTableNodeOperation", JavaWindow("SystemsEngineering").JavaWindow("WEmbeddedFrame"),"CMEBOMTreeTable",iRowNum,iColNum, "LEFT", "NONE")
					wait 2 
					Fn_SE_AttachmentTableNodeOperation = True
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_SE_AttachmentTableNodeOperation : Failed to find Column [ " & sColumn & " ].")
				End If
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_SE_AttachmentTableNodeOperation : Failed to find Row [ " & sAttchLine & " ].")
			End IF
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "GetCellData"
				If bFlag Then 
					'Fn_SE_AttachmentTableNodeOperation = JavaWindow("SystemsEngineering").JavaWindow("WEmbeddedFrame").JavaTable("CMEBOMTreeTable").GetCellData(iRowNum,sColumn) :Swapnil
					Fn_SE_AttachmentTableNodeOperation = Window("SEWindow").JavaWindow("WEmbeddedFrame").JavaTable("CMEBOMTreeTable").GetCellData(iRowNum,sColumn)
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_SE_AttachmentTableNodeOperation : Failed to find Row [ " & sAttchLine & " ].")
				End If
				Window("SEWindow").JavaWindow("WEmbeddedFrame").SetTOProperty "index",0
				Exit Function
		Case "SetCellData"		
				iRowNo = Fn_SE_BOMTable_RowIndex(sAttchLine)
            	If isNumeric(iRowNo) then
								'Get column Rows
					iColNo=objTable.GetROProperty("cols")
					For iStart=0 to iColNo-1
						''Verify the Column name is similar to required column name
						If objTable.GetColumnName(iStart)=sColumn Then
							'Verify the Column value is similar to required value
							objTable.SetCellData iRowNo,iStart,sData
							Fn_SE_AttachmentTableNodeOperation=True
							Exit For
						End If
                Next
				Else
					Fn_SE_AttachmentTableNodeOperation=False
				End if				
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "AddColumn", "AddColumns"            ''Modfied by Avinash J on 12-March-2013  for Hirarchy Chanage 
			aColumns = split(sColumn,"~")
			If  Not  Window("SEWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Change Columns").Exist Then ''insert column dialog is allready open   Added by Avinash J.
								objTable.SelectColumnHeader 1,"RIGHT"
								Call Fn_ReadyStatusSync(2)
								Window("SEWindow").JavaWindow("WEmbeddedFrame").JavaMenu("label:=Insert column\(s\) \.\.\.").Select
			 End IF
			
				If  Window("SEWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Change Columns").Exist Then
					Set objChangeCol = Window("SEWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Change Columns")
					For iColumnCnt = 0 to UBound(aColumns)
						If aColumns(iColumnCnt) <> "" Then
							Call Fn_List_Select("Fn_SE_AttachmentTableNodeOperation", objChangeCol,"ListAvailableCols",aColumns(iColumnCnt))
							Call Fn_Button_Click("Fn_SE_AttachmentTableNodeOperation", objChangeCol,"Add")
							wait 1
						End If
					Next
					If cInt(objChangeCol.JavaButton("Apply").getROProperty("enabled")) = 1 then
						Call Fn_Button_Click("Fn_SE_AttachmentTableNodeOperation", objChangeCol,"Apply")
					End if
					Call Fn_Button_Click("Fn_SE_AttachmentTableNodeOperation", objChangeCol,"Cancel")
					Fn_SE_AttachmentTableNodeOperation = True
				End if
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "VerifyColumn", "VerifyColumns"
					'iCol=JavaWindow("SystemsEngineering").JavaWindow("WEmbeddedFrame").JavaTable("CMEBOMTreeTable").GetROProperty("cols") :Swapnil
					iCol=Window("SEWindow").JavaWindow("WEmbeddedFrame").JavaTable("CMEBOMTreeTable").GetROProperty("cols")
					aColumns = split(sColumn,"~")
					For iColumnCnt = 0 To UBound(aColumns)
						For iCount=0 To iCol-1
								bFlag=False
								'currColName=JavaWindow("SystemsEngineering").JavaWindow("WEmbeddedFrame").JavaTable("CMEBOMTreeTable").GetColumnName(iCount) :Swapnil
								currColName=Window("SEWindow").JavaWindow("WEmbeddedFrame").JavaTable("CMEBOMTreeTable").GetColumnName(iCount)
								If Trim(currColName)=aColumns(iColumnCnt) Then
									bFlag=True
									Exit For
								End If
						Next
						If bFlag=False Then
							Exit For
						End If
					Next
					If bFlag=True Then
						Fn_SE_AttachmentTableNodeOperation=True
					End If
        ' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	    Case "PopupMenuSelect"
				Call Fn_UI_JavaTable_SelectRow("Fn_SE_AttachmentTableNodeOperation", Window("SEWindow").JavaWindow("WEmbeddedFrame"),"CMEBOMTreeTable",iRowNum) 
				wait 1
				If isNumeric(iRowNum) then
					objTable.ClickCell iRowNum,0,"RIGHT" 
					wait 1
					StrMenu = JavaWindow("SystemsEngineering").WinMenu("ContextMenu").BuildMenuPath(sMenu)
					wait 1
					JavaWindow("SystemsEngineering").WinMenu("ContextMenu").Select StrMenu
					wait 1
					Fn_SE_AttachmentTableNodeOperation=True
				End if
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case Else
			Exit function
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	End Select
	If Fn_SE_AttachmentTableNodeOperation  Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_SE_AttachmentTableNodeOperation : executed successfully with case [ " & sAction & " ].")
	End If
	Window("SEWindow").JavaWindow("WEmbeddedFrame").SetTOProperty "index",0
End Function

'*********************************************************		Function to Create Requirement with project in SE ***********************************************************************
'Function Name		:				Fn_SE_CreateNewRequirementWithProject

'Description			 :		 		 This function is used to Create the Requirement in System Engineering

'Parameters			   :	 		  1. sAction : Action tobe performed
'									2. strNodeName : Requirement Type
'									3. strReqID  : Requirement ID
'									4. strReqRev :  : Requirement Rev Id
'									5. strReqName :  : Requirement Name
'									6. strReqDesc : : Requirement Description
'									7. sOwningProject  : Owning User Name
'									8. sSelectionType  : Selection Type  : MoveLeftAll / MoveRightAll / MoveLeft / MoveRight
'									9. sProjectsForSelection : ~ separated list of projects
'									10. sSelectedProjects : ~ separated list of projects

'Return Value		   : 		TRUE \ FALSE

'Pre-requisite		    :	  should be in SE Prespective

'Examples			    :	  Call Fn_SE_CreateNewRequirementWithProject("Create", "Requirement","","","Req1","Desc", "", "MoveLeftAll", "", "")
'Examples			    :	  Call Fn_SE_CreateNewRequirementWithProject("Create", "Requirement","","","Req1","Desc", "", "MoveLeft", "", "Prj1~Prj2")
'Examples			    :	  Call Fn_SE_CreateNewRequirementWithProject("Create", "Requirement","","","Req1","Desc", "", "MoveRightAll", "", "")
'Examples			    :	  Call Fn_SE_CreateNewRequirementWithProject("Create", "Requirement","","","Req1","Desc", "", "MoveRight", "Prj1~Prj2", "")


'History:
'				Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'				Koustubh Watwe			28.01.2011
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public function Fn_SE_CreateNewRequirementWithProject(sAction, strNodeName, strReqID, strReqRev, strReqName, strReqDesc, sOwningProject, sSelectionType, sProjectsForSelection, sSelectedProjects )
		GBL_FAILED_FUNCTION_NAME="Fn_SE_CreateNewRequirementWithProject"
		Dim strNodePath,intNodeCount,sTreeItem,intCount,strNodePathC
		Dim ObjReqWnd,bFlag, aNodes
		Fn_SE_CreateNewRequirementWithProject=False
		' checking for existance of create requirement dialog
		If Fn_UI_ObjectExist("Fn_SE_CreateNewRequirementWithProject",JavaWindow("SystemsEngineering").JavaWindow("NewRequirement"))=False Then
				Call Fn_MenuOperation("Select","File:New:Requirement...")
				Call Fn_ReadyStatusSync(3)
		End If
		Set ObjReqWnd=Fn_UI_ObjectCreate("Fn_SE_CreateNewRequirementWithProject",JavaWindow("SystemsEngineering").JavaWindow("NewRequirement"))
		Select Case sAction
				' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
				Case "Create"
						' setting requirement type
						Call Fn_UI_JavaTree_Expand("Fn_SE_CreateNewRequirementWithProject", ObjReqWnd, "RequirementTree","Complete List")
						JavaWindow("SystemsEngineering").JavaWindow("NewRequirement").JavaTree("RequirementTree").WaitProperty "items count" , micGreaterThan(1)  
						If Fn_UI_JavaTree_NodeExist("Fn_SE_RequirementSpecCreate",ObjReqWnd.JavaTree("RequirementTree"),"Complete List:"+strNodeName) Then
								strNodePathC="Complete List:"+strNodeName
						Else
								strNodePathC="Most Recently Used:"+strNodeName
						End If
						Call Fn_JavaTree_Select("Fn_SE_CreateNewRequirementWithProject", ObjReqWnd, "RequirementTree",strNodePathC)
						wait(2)
						Call Fn_Button_Click("Fn_SE_CreateNewRequirementWithProject",ObjReqWnd,"Next")
						wait(2)
						'setting requirement ID
						If strReqID<>"" Then
								wait(2)
								'Call Fn_UI_EditBox_Type("Fn_SE_CreateNewRequirementWithProject",ObjReqWnd,"ID",strReqID)
						End If
						'setting requirement revision id
						If strReqRev<>"" Then
								wait(2)
								'Call Fn_UI_EditBox_Type("Fn_SE_CreateNewRequirementWithProject",ObjReqWnd,"Revision",strReqRev)
						End If
						'setting requirement name
						wait(2)
						Call  Fn_Edit_Box("Fn_SE_CreateNewRequirementWithProject",ObjReqWnd,"Name",strReqName)
						'Call Fn_UI_EditBox_Type("Fn_SE_CreateNewRequirementWithProject",ObjReqWnd,"Name",strReqName)
						' setting requirement description
						If strReqDesc <> "" Then
							wait(2)
							Call Fn_UI_EditBox_Type("Fn_SE_CreateNewRequirementWithProject",ObjReqWnd,"Description",strReqDesc)
						End If
						' clicking on Next
						wait(2)
						Call Fn_Button_Click("Fn_SE_CreateNewRequirementWithProject",ObjReqWnd,"Next")
						Wait(2)	
						' clicking on Next
						Call Fn_Button_Click("Fn_SE_CreateNewRequirementWithProject",ObjReqWnd,"Next")
						wait(2)
						
						'setting owning project
						If sOwningProject <> "" Then
							Call Fn_UI_EditBox_Type("Fn_SE_CreateNewRequirementWithProject",ObjReqWnd,"OwningProject",sOwningProject)
						End If

						Select Case sSelectionType
							Case "MoveLeftAll", "MoveRightAll"
								Call Fn_Button_Click("Fn_SE_CreateNewRequirementWithProject",ObjReqWnd, sSelectionType)
							' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
							Case "MoveRight"
                                If sProjectsForSelection <> "" Then
									aNodes = split(sProjectsForSelection,"~")
									For intCount = 0 to UBound(aNodes)
										wait 2
										If Fn_UI_ListItemExist("Fn_SE_CreateNewRequirementWithProject",ObjReqWnd,"ProjectsForSelectionList",aNodes(intCount)) then
												Call Fn_List_Select("Fn_SE_CreateNewRequirementWithProject",ObjReqWnd,"ProjectsForSelectionList",aNodes(intCount))
												wait 1
												Call Fn_Button_Click("Fn_SE_CreateNewRequirementWithProject",ObjReqWnd,"MoveRight")
										Else 
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_SE_CreateNewRequirementWithProject: Project [ " & aNodes(intCount) & " ] is not present in [ Projects For Selection ] list.")
											Exit Function
										End IF										
									Next
								End If
							' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
							Case "MoveLeft"
								If sSelectedProjects <> "" Then
									aNodes = split(sSelectedProjects,"~")
									For intCount = 0 to UBound(aNodes)
										If Fn_UI_ListItemExist("Fn_SE_CreateNewRequirementWithProject",ObjReqWnd,"SelectedProjectsList",aNodes(intCount)) then
												Call Fn_List_Select("Fn_SE_CreateNewRequirementWithProject",ObjReqWnd,"SelectedProjectsList",aNodes(intCount))
												Call Fn_Button_Click("Fn_SE_CreateNewRequirementWithProject",ObjReqWnd,"MoveLeft")
										Else 
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_SE_CreateNewRequirementWithProject: Project [ " & aNodes(intCount) & " ] is not present in [ Selected Projects ] list.")
											Exit Function
										End IF
									Next
								End If
							' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
						End Select

						' clicking on Finish button
						Call Fn_Button_Click("Fn_SE_CreateNewRequirementWithProject", ObjReqWnd, "Finish")
						Call Fn_ReadyStatusSync(3)

						'Click on Cancel Button
						If JavaWindow("SystemsEngineering").JavaWindow("NewRequirement").Exist = True Then
								Call Fn_Button_Click("Fn_SE_CreateNewRequirementWithProject",ObjReqWnd,"Cancel")
								wait(2)
						End If
						iRowNo = Fn_SE_BOMTable_RowIndex(strReqName)
						DataTable.GetSheet("Global").AddParameter "NewReqID",""
						DataTable.GetSheet("Global").AddParameter "NewReqRev",""
						sReq=Fn_SE_BOMTableNodeOpeations("GetCellData",iRowNo,0,"","")
						If sReq <> "" Then
							sReqFull=mid(sReq,instr(1,sReq,":")+1,len(sReq))
							sChdReqName1=mid(sReqFull,instr(1,sReqFull,":")+1,len(sReqFull))
							aItmInfo3 = split(sChdReqName1, "/", -1, 1)
							DataTable("NewReqID",dtGlobalSheet)= aItmInfo3(0)
							aItmInfo4 = split(aItmInfo3(1), ";", -1, 1)
							DataTable("NewReqRev",dtGlobalSheet)= aItmInfo4(0)
						End If
						Fn_SE_CreateNewRequirementWithProject = True
						Set ObjReqWnd=Nothing
				' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
				Case Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_SE_CreateNewRequirementWithProject: Invalid case [ " & sAction & " ].")
	End Select

	If Fn_SE_CreateNewRequirementWithProject = True Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_SE_CreateNewRequirementWithProject: Executed successfully with case [ " & sAction & " ].")
	End If
End Function
'******************************************************************************************************************************************************************************************************************
'Function Name		:				Fn_SE_AccountiblityCheck

'Description			 :		 		 This function is used for accountiblity check in System Engineering

'Parameters			   :	 		  1. sAction : Action to be performed

'Return Value		   : 		TRUE \ FALSE

'Pre-requisite		    :	  should be in SE Prespective

'Examples			    :	  
'									Set dicAccCheck = CreateObject( "Scripting.Dictionary" )
'										dicAccCheck.RemoveAll
'										dicAccCheck("Check") = "yes"
'										dicAccCheck("SelectedSrcObjs") = "000119/A;1-Spec1 (View)"
'										dicAccCheck("SelectedTargetObjs") = "000112/A;1-Proc1 (View)"
'										dicAccCheck("DisplayOptions") = "yes"
'										dicAccCheck("ReportOccurrenceGroups") = "OFF"
'										dicAccCheck("ReportSelectedCheck") = "ON"
'										dicAccCheck("CompareOptions") = "yes"	
'										dicAccCheck("PartialMatchOptions") = "yes"	
'										dicAccCheck("Button") = "OK"
'										Msgbox Fn_SE_AccountiblityCheck("Verify", dicAccCheck)	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		'								dicAccCheck("AddSrcObjs") = True
		'								dicAccCheck("RemoveSrcObjs") = ""
		'								dicAccCheck("SwitchSourceTarget") = False
		'								dicAccCheck("AddTargetObjs") = True
		'								dicAccCheck("RemoveTargetObjs") = ""
		'								
		'								dicAccCheck("SearchCurrentlyExpandedSourceLines") = "Yes"
		'								dicAccCheck("CompareLowestVisibleOfSource") = True
		'								
		'								dicAccCheck("SearchLinesPerFilteringRule") = "Yes"
		'								dicAccCheck("SourceFilteringRule") = ""
		'								dicAccCheck("SourceLevel") = "3" 
		'								dicAccCheck("TargetFilteringRule") = False
		'								dicAccCheck("TargetLevel") = ""
'										dicAccCheck("Button") = ""
'										msgbox Fn_SE_AccountiblityCheck("ModifyScopeTab", dicAccCheck)
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		'								dicAccCheck("EquivalentTraceLink") = True
		'								dicAccCheck("LogicalDesignator") = ""
		'								dicAccCheck("PublishLinkConnection") = False
'										dicAccCheck("Button") = ""
'										msgbox Fn_SE_AccountiblityCheck("ModifyEquivalenceTab", dicAccCheck)
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										dicAccCheck("Properties") = True
'										dicAccCheck("ConsiderValuesOfProperties") = True
'										dicAccCheck("AddAvailableProperties") = "APN UID~Allocated Time"
'										dicAccCheck("RemoveSelectedProperties") = "Quantity~Revision"
'										dicAccCheck("MoveUpSelectedProperties") = "Allocated Time:2"
'										dicAccCheck("MoveDownSelectedProperties") = "Allocated Time:2"
'										dicAccCheck("Button") = "OK"

'										msgbox Fn_SE_AccountiblityCheck("ModifyPartialMatchTab", dicAccCheck)
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'									Set dicAccCheck = Nothing

'History:
'	Developer Name			Date		  Rev.No.	Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Ketan Raje				05/09/2011		1.0
'   Sanjeet kumar     		04/01/2011				Changes in Menu Hierarchy Tools:Compare:Accountability Check:Accountability Check...
'   Koustubh Watwe     		10/05/2012		1.0		Added cases ModifyScopeTab, ModifyEquivalenceTab, ModifyPartialMatchTab
'   Koustubh Watwe     		03/08/2012		1.0		Modifeid cases ModifyReportingTab,"ModifyEquivalenceTab"
' 	Sagar S. 							03/12/12					Added code to check existance of Button, resize dialog and click button. As On 919b build, buttons does not displayed until resize of dialog. dependable on screen resolution
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public function Fn_SE_AccountiblityCheck(sAction, dicAccCheck)
	GBL_FAILED_FUNCTION_NAME="Fn_SE_AccountiblityCheck"
	Dim iCounter, iCount, aValues, iRows, intCnt, iTotal
	Dim aObjects
	Fn_SE_AccountiblityCheck=False
	' checking for existance of create requirement dialog
	If Fn_UI_ObjectExist("Fn_SE_AccountiblityCheck",JavaWindow("SystemsEngineering").JavaWindow("AccountabilityCheck"))=False Then
		Call Fn_MenuOperation("Select","Tools:Compare:Accountability Check:Accountability Check...")
	End If
	Set ObjAccCheck = Fn_UI_ObjectCreate("Fn_SE_AccountiblityCheck",JavaWindow("SystemsEngineering").JavaWindow("AccountabilityCheck"))
	Select Case sAction		
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Verify"
			iCounter = 0
			iCount = 0
			'Activate Check Tab
			If Trim(Lcase(dicAccCheck("Check"))) = "yes" Then
				ObjAccCheck.JavaTab("AccountiblityTab").Select "Check"
				'Verify Selected Source Objects.
				If dicAccCheck("SelectedSrcObjs") <> "" Then
					aValues = Split(dicAccCheck("SelectedSrcObjs"),":")
					iRows = ObjAccCheck.JavaList("SelectedSource").GetROProperty("items count")
					For intCnt = 0 to Ubound(aValues)
						iCounter = iCounter + 1
						For iTotal = 0 to iRows-1
							If Trim(Lcase(aValues(intCnt))) = Trim(Lcase(ObjAccCheck.JavaList("SelectedSource").GetItem(iTotal))) Then
								iCount = iCount + 1
								Exit For
							End If
						Next
					Next
				End If
				'Verify Selected Target Objects.
				If dicAccCheck("SelectedTargetObjs") <> "" Then
						aValues = Split(dicAccCheck("SelectedTargetObjs"),":")
						iRows = ObjAccCheck.JavaList("SelectedTarget").GetROProperty("items count")
						For intCnt = 0 to Ubound(aValues)
							iCounter = iCounter + 1
							For iTotal = 0 to iRows-1
								If Trim(Lcase(aValues(intCnt))) = Trim(Lcase(ObjAccCheck.JavaList("SelectedTarget").GetItem(iTotal))) Then
									iCount = iCount + 1
									Exit For
								End If
							Next
						Next
				End If
			End If
			'Activate Display Options Tab.
			If Trim(Lcase(dicAccCheck("DisplayOptions"))) = "yes" Then
				ObjAccCheck.JavaTab("AccountiblityTab").Select "Display Options"
				'Verify status of ReportOccurrenceGroups
				If dicAccCheck("ReportOccurrenceGroups") <> "" Then
					iCounter = iCounter + 1
					ObjAccCheck.JavaRadioButton("RadioButton").SetTOProperty "attached text","Report in occurrence groups"
					If Trim(Lcase(dicAccCheck("ReportOccurrenceGroups"))) = "on" And ObjAccCheck.JavaRadioButton("RadioButton").GetROProperty("value") = 1 Then
						iCount = iCount + 1
					ElseIf Trim(Lcase(dicAccCheck("ReportOccurrenceGroups"))) = "off" And ObjAccCheck.JavaRadioButton("RadioButton").GetROProperty("value") = 0 Then
						iCount = iCount + 1
					End If
				End If
				'Verify status of ReportSelectedCheck
				If dicAccCheck("ReportSelectedCheck") <> "" Then
					iCounter = iCounter + 1
					ObjAccCheck.JavaRadioButton("RadioButton").SetTOProperty "attached text","Report the selected check criteria"
					If Trim(Lcase(dicAccCheck("ReportSelectedCheck"))) = "on" And ObjAccCheck.JavaRadioButton("RadioButton").GetROProperty("value") = 1 Then
						iCount = iCount + 1
					ElseIf Trim(Lcase(dicAccCheck("ReportSelectedCheck"))) = "off" And ObjAccCheck.JavaRadioButton("RadioButton").GetROProperty("value") = 0 Then
						iCount = iCount + 1
					End If
				End If
			End If
			'Activate Compare Options Tab.
			If Trim(Lcase(dicAccCheck("CompareOptions"))) = "yes" Then
				ObjAccCheck.JavaTab("AccountiblityTab").Select "Compare Options"
			End If
			'Activate Partial Match Options Tab.
			If Trim(Lcase(dicAccCheck("PartialMatchOptions"))) = "yes" Then
				ObjAccCheck.JavaTab("AccountiblityTab").Select "Partial Match Options"
			End If
			If iCounter = iCount Then
				Fn_SE_AccountiblityCheck = True
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "ModifyScopeTab"
			' select Scope tab
			ObjAccCheck.JavaTab("AccountiblityTab").Select "Scope"
			' addSourceObject
			If dicAccCheck("AddSrcObjs") <> "" Then
				If cBool(dicAccCheck("AddSrcObjs")) Then
					If  ObjAccCheck.JavaButton("AddSource").Exist(5) Then
						Call Fn_Button_Click("Fn_SE_AccountiblityCheck", ObjAccCheck, "AddSource")
					End if 
					If  ObjAccCheck.JavaButton("AddSetSourceBtn").Exist(5) Then
						Call Fn_Button_Click("Fn_SE_AccountiblityCheck", ObjAccCheck, "AddSetSourceBtn")
					End if 
				End If
			End If
			' select Source Object(s)
			If dicAccCheck("RemoveSrcObjs") <> "" Then
				aObjects = split(dicAccCheck("RemoveSrcObjs"), "~")
				For iCount = 0 to uBound(aObjects)
					If Fn_List_Select("Fn_SE_AccountiblityCheck", ObjAccCheck,"SelectedSource",aObjects(iCount)) = False Then
						exit function
					End If
					' click on remove
					Call Fn_Button_Click("Fn_SE_AccountiblityCheck", ObjAccCheck, "RemoveSource")
				Next
			End If
			
			' click on Switch Source Target
			If dicAccCheck("SwitchSourceTarget") <> "" Then
				If cBool(dicAccCheck("SwitchSourceTarget")) Then
					Call Fn_Button_Click("Fn_SE_AccountiblityCheck", ObjAccCheck, "Flip")
				End If
			End If
			' addTargetObject
			If dicAccCheck("AddTargetObjs") <> "" Then
				If cBool(dicAccCheck("AddTargetObjs")) Then
					If  ObjAccCheck.JavaButton("AddTarget").Exist(5) Then
						Call Fn_Button_Click("Fn_SE_AccountiblityCheck", ObjAccCheck, "AddTarget")
					End if 
					If  ObjAccCheck.JavaButton("AddSetTargetBtn").Exist(5) Then
						Call Fn_Button_Click("Fn_SE_AccountiblityCheck", ObjAccCheck, "AddSetTargetBtn")
					End if 
				End If
			End If
			' select Target Object(s)
			If dicAccCheck("RemoveTargetObjs") <> "" Then
				aObjects = split(dicAccCheck("RemoveTargetObjs"), "~")
				For iCount = 0 to uBound(aObjects)
					If Fn_List_Select("Fn_SE_AccountiblityCheck", ObjAccCheck,"SelectedTarget",aObjects(iCount)) = False Then
						exit function
					End If
					' click on remove
					Call Fn_Button_Click("Fn_SE_AccountiblityCheck", ObjAccCheck, "RemoveTarget")
				Next
			End If

			' select radiobox of Search currently expanded source lines
			If dicAccCheck("SearchCurrentlyExpandedSourceLines") <> "" Then
				ObjAccCheck.JavaRadioButton("RadioButton").SetTOProperty "attached text","Search currently expanded source lines"
				Call Fn_UI_JavaRadioButton_SetON("Fn_SE_AccountiblityCheck",ObjAccCheck, "RadioButton")
			
			' select checkbox of compare lowest visible
				ObjAccCheck.JavaCheckBox("CheckBox").SetTOProperty "attached text","Compare lowest visible level of source"
				If cBool(dicAccCheck("CompareLowestVisibleOfSource")) Then
					Call Fn_CheckBox_Set("Fn_SE_AccountiblityCheck", ObjAccCheck, "CheckBox", "ON")
				Else
					Call Fn_CheckBox_Set("Fn_SE_AccountiblityCheck", ObjAccCheck, "CheckBox", "OFF")
				End If
			End If
			
			' select searhc line per filter radio button
			If dicAccCheck("SearchLinesPerFilteringRule") <> "" Then
				ObjAccCheck.JavaRadioButton("RadioButton").SetTOProperty "attached text","Search lines per filtering rule"
				Call Fn_UI_JavaRadioButton_SetON("Fn_SE_AccountiblityCheck",ObjAccCheck, "RadioButton")
				' select source filtering rule
				If dicAccCheck("SourceFilteringRule") <> "" Then
					If Fn_List_Select("Fn_SE_AccountiblityCheck", ObjAccCheck,"SearchSourceFilter",dicAccCheck("SourceFilteringRule")) = False Then
						exit function
					End If
				End If
				' select checkbox Limit Search in Source to first ... levels
				If dicAccCheck("SourceLevel") <> "" Then
					ObjAccCheck.JavaCheckBox("CheckBox").SetTOProperty "attached text","Limit search in source to first"
					If lcase(cStr(dicAccCheck("SourceLevel"))) = "false" Then
						Call Fn_CheckBox_Set("Fn_SE_AccountiblityCheck", ObjAccCheck, "CheckBox", "OFF")
					Else
						Call Fn_CheckBox_Set("Fn_SE_AccountiblityCheck", ObjAccCheck, "CheckBox", "ON")
						Call Fn_Edit_Box("Fn_SE_AccountiblityCheck",ObjAccCheck, "SourceLevel", dicAccCheck("SourceLevel"))
					End IF
				End If
				' select target filtering rule
				If dicAccCheck("TargetFilteringRule") <> "" Then
					If Fn_List_Select("Fn_SE_AccountiblityCheck", ObjAccCheck,"SearchTargetFilter",dicAccCheck("TargetFilteringRule")) = False Then
						exit function
					End If
				End If
				' select checkbox Limit Search in Target to first ... levels
				If dicAccCheck("TargetLevel") <> "" Then
					ObjAccCheck.JavaCheckBox("CheckBox").SetTOProperty "attached text","Limit search in target to first"
					If lcase(cStr(dicAccCheck("TargetLevel"))) = "false" Then
						Call Fn_CheckBox_Set("Fn_SE_AccountiblityCheck", ObjAccCheck, "CheckBox", "OFF")
					Else
						Call Fn_CheckBox_Set("Fn_SE_AccountiblityCheck", ObjAccCheck, "CheckBox", "ON")
						Call Fn_Edit_Box("Fn_SE_AccountiblityCheck",ObjAccCheck, "TargetLevel", dicAccCheck("TargetLevel"))
					End If
				End If
			End If
			Fn_SE_AccountiblityCheck = True
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "ModifyReportingTab"
			ObjAccCheck.JavaTab("AccountiblityTab").Select "Reporting"
			If dicAccCheck("ReportInOccurrenceGroup") <> "" Then
				If cBool(dicAccCheck("ReportInOccurrenceGroup")) Then
					Call Fn_UI_JavaRadioButton_SetON("Fn_SE_AccountiblityCheck",ObjAccCheck,"ReportInOccurrenceGroups")
				Else
					' do nothing
				End If
			End If
			
			If dicAccCheck("ReportTheSelectedCheckCriteria") <> "" Then
				Call Fn_UI_JavaRadioButton_SetON("Fn_SE_AccountiblityCheck",ObjAccCheck,"ReportTheSelectedCheck")
			End If
			
			If dicAccCheck("ColourTheComparedObjects") <> "" Then
				If cBool(dicAccCheck("ColourTheComparedObjects")) Then
					Call Fn_CheckBox_Set("Fn_SE_AccountiblityCheck", ObjAccCheck, "ColorTheComparedObjects", "ON")
				Else
					Call Fn_CheckBox_Set("Fn_SE_AccountiblityCheck", ObjAccCheck, "ColorTheComparedObjects", "OFF")
				End If
			End If
			If dicAccCheck("PrintableReportName") <> "" Then
				Call Fn_CheckBox_Set("Fn_SE_AccountiblityCheck", ObjAccCheck, "PrintableReportNameCheckBox", "ON")
				Call Fn_Edit_Box("Fn_SE_AccountiblityCheck",ObjAccCheck, "PrintableReportName", dicAccCheck("PrintableReportName"))
			End If
			If dicAccCheck("DisplayOptions") <> "" Then
				aDisplayOptn = split(dicAccCheck("DisplayOptions"),"~")
				aDisplayOptnVal = split(dicAccCheck("DisplayOptionValues"),"~")
				For iCnt = 0 to uBound(aDisplayOptn)
					ObjAccCheck.JavaCheckBox("MatchCheckBox").SetTOProperty "attached text", aDisplayOptn(iCnt)
					If ObjAccCheck.JavaCheckBox("MatchCheckBox").Exist(5) Then
						If cBool(aDisplayOptnVal(iCnt)) Then
							Call Fn_CheckBox_Set("Fn_SE_AccountiblityCheck", ObjAccCheck, "MatchCheckBox", "ON")
						Else
							Call Fn_CheckBox_Set("Fn_SE_AccountiblityCheck", ObjAccCheck, "MatchCheckBox", "OFF")
						End If
					Else
						Exit Function
					End If
				Next
			End IF
			Fn_SE_AccountiblityCheck = True
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "ModifyEquivalenceTab"
			ObjAccCheck.JavaTab("AccountiblityTab").Select "Equivalence"

			If dicAccCheck("LogicalDesignator") <> "" Then
				' select checkbox of compare lowest visible
					ObjAccCheck.JavaCheckBox("CheckBox").SetTOProperty "attached text","Logical Designator"
					If cBool(dicAccCheck("LogicalDesignator")) Then
						Call Fn_CheckBox_Set("Fn_SE_AccountiblityCheck", ObjAccCheck, "CheckBox", "ON")
					Else
						Call Fn_CheckBox_Set("Fn_SE_AccountiblityCheck", ObjAccCheck, "CheckBox", "OFF")
					End If
			End If

			If dicAccCheck("EquivalentTraceLink") <> "" Then
				' select checkbox of compare lowest visible
					ObjAccCheck.JavaCheckBox("CheckBox").SetTOProperty "attached text","Trace Link"
					If cBool(dicAccCheck("EquivalentTraceLink")) Then
						Call Fn_CheckBox_Set("Fn_SE_AccountiblityCheck", ObjAccCheck, "CheckBox", "ON")
					Else
						Call Fn_CheckBox_Set("Fn_SE_AccountiblityCheck", ObjAccCheck, "CheckBox", "OFF")
					End If
			End If

			If dicAccCheck("PublishLinkConnection") <> "" Then
				' select checkbox of compare lowest visible
					ObjAccCheck.JavaCheckBox("CheckBox").SetTOProperty "attached text","PublishLink Connection"
					If cBool(dicAccCheck("PublishLinkConnection")) Then
						Call Fn_CheckBox_Set("Fn_SE_AccountiblityCheck", ObjAccCheck, "CheckBox", "ON")
					Else
						Call Fn_CheckBox_Set("Fn_SE_AccountiblityCheck", ObjAccCheck, "CheckBox", "OFF")
					End If
			End If
			Fn_SE_AccountiblityCheck = True
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "ModifyPartialMatchTab"
			ObjAccCheck.JavaTab("AccountiblityTab").Select "Partial Match"
			'modified on 12-Aug-13
			'remove all properties from selected property ist
			If 	lcase(dicAccCheck("RemoveAllProperties"))="yes"		 Then                
				'select properties and click on move left			
				iRows = cInt(ObjAccCheck.JavaTable("SelectedProperties").GetROProperty("rows"))
				For iCnt = 0 to iRows - 1
					ObjAccCheck.JavaTable("SelectedProperties").ClickCell 0, 0 
					Call Fn_Button_Click("Fn_SE_AccountiblityCheck", ObjAccCheck, "MoveLeft")
				Next
			End If
			
			If dicAccCheck("Properties") = "" Then dicAccCheck("Properties") = False
			If cBool(dicAccCheck("Properties")) Then
				'ObjAccCheck.JavaTab("PartialMatchTab").Select "BOM Properties"
				' Consider Values Of Properties
				If dicAccCheck("ConsiderValuesOfProperties") <> "" Then
					If cBool(dicAccCheck("ConsiderValuesOfProperties") ) Then
						Call Fn_CheckBox_Set("Fn_SE_AccountiblityCheck", ObjAccCheck, "ConsiderValuesOfProperties", "ON")
					Else
						Call Fn_CheckBox_Set("Fn_SE_AccountiblityCheck", ObjAccCheck, "ConsiderValuesOfProperties", "OFF")
						Fn_SE_AccountiblityCheck = True
					End if
				End If
				' select Available proeprties and click on Add
				Dim iRowCount
				If dicAccCheck("AddProperties") <> "" Then
					aValues = split(dicAccCheck("AddProperties"),"~")
					For iCount = 0 to UBound(aValues)
						' verifying property present in SelectedProperties Table
						iRows = cInt(ObjAccCheck.JavaTable("SelectedProperties").GetROProperty("rows"))
						For iCnt = 0 to iRows - 1
							If ObjAccCheck.JavaTable("SelectedProperties").GetCellData(iCnt,0) = aValues(iCount) Then
									Exit for
							End If
						Next
						' if proeprty is not present in SelectedProperties
						If iCnt = iRows Then
							iRows = cInt(ObjAccCheck.JavaTable("AvailableProperties").GetROProperty("rows"))
							For iCnt = 0 to iRows - 1
								If ObjAccCheck.JavaTable("AvailableProperties").GetCellData(iCnt,"Property") = aValues(iCount) Then
										'ObjAccCheck.JavaTable("AvailableProperties").SelectRow iCnt
										ObjAccCheck.JavaTable("AvailableProperties").ClickCell iCnt,"Property" ' Ankit T, | 11.2 Porting | 5 Mar 15 | Added code to properly select row.
										wait 1,500
										Exit For
								End If
							Next
							If iCnt = iRows Then
								'aValues(iCount) is not present
								Exit function
							Else
								Call Fn_Button_Click("Fn_SE_AccountiblityCheck", ObjAccCheck, "MoveRight")
							End If
						End If
					Next
				Fn_SE_AccountiblityCheck = True
				End If
				' select Selected proeprties and click on remove
				If dicAccCheck("RemoveProperties") <> "" Then
					aValues = split(dicAccCheck("RemoveProperties"),"~")
					For iCount = 0 to UBound(aValues)
						' verifying property present in AvailableProperties Table
						iRows = cInt(ObjAccCheck.JavaTable("AvailableProperties").GetROProperty("rows"))
						For iCnt = 0 to iRows - 1
							If ObjAccCheck.JavaTable("AvailableProperties").GetCellData(iCnt,"Property") = aValues(iCount) Then
									Exit for
							End If
						Next
						' if proeprty is not present in AvailableProperties
						If iCnt = iRows Then
							iRows = cInt(ObjAccCheck.JavaTable("SelectedProperties").GetROProperty("rows"))
							For iCnt = 0 to iRows - 1
								If ObjAccCheck.JavaTable("SelectedProperties").GetCellData(iCnt,"") = aValues(iCount) Then
										ObjAccCheck.JavaTable("SelectedProperties").SelectCell iCnt,""
										Exit for
								End If
							Next
						End If
						If iCnt = iRows Then
							'aValues(iCount) is not present
							Exit function
						Else
							Call Fn_Button_Click("Fn_SE_AccountiblityCheck", ObjAccCheck, "MoveLeft")
						End If
					Next
				Fn_SE_AccountiblityCheck = True
				End If
				' select Selected proeprties and click on MoveUp
				If dicAccCheck("MoveUpProperties") <> "" Then
					aValues = split(dicAccCheck("MoveUpProperties"),"~")
					iRows = cInt(ObjAccCheck.JavaTable("SelectedProperties").GetROProperty("rows"))
					For iCount = 0 to UBound(aValues)
						For iCnt = 0 to iRows - 1
							aObjects = split(aValues(iCount),":")
							If ObjAccCheck.JavaTable("SelectedProperties").GetCellData(iCnt,"Property") = aObjects(0) Then
									ObjAccCheck.JavaTable("SelectedProperties").SelectRow iCnt
									Exit for
							End If
						Next
					Next
					If iCnt = iRows Then
						'aValues(iCount) is not present
						Exit function
					Else
						iCounter = 1
						If UBound(aObjects) > 0 Then
							iCounter = cInt(aObjects(1))
						End If
						For iCnt = 1 to iCounter
							If cInt(ObjAccCheck.JavaButton("MoveUp").GetROProperty("enabled")) = 1 then
								Call Fn_Button_Click("Fn_SE_AccountiblityCheck", ObjAccCheck, "MoveUp")
							End If
						Next
					End If
					Fn_SE_AccountiblityCheck = True
				End If
				'select Selected proeprties and click on MoveDown
				If dicAccCheck("MoveDownProperties") <> "" Then
					aValues = split(dicAccCheck("MoveDownProperties"),"~")
					iRows = cInt(ObjAccCheck.JavaTable("SelectedProperties").GetROProperty("rows"))
					For iCount = 0 to UBound(aValues)
						For iCnt = 0 to iRows - 1
							aObjects = split(aValues(iCount),":")
							If ObjAccCheck.JavaTable("SelectedProperties").GetCellData(iCnt,"Property") = aObjects(0) Then
									ObjAccCheck.JavaTable("SelectedProperties").SelectRow iCnt
									Exit for
							End If
						Next
					Next
					If iCnt = iRows Then
						'aValues(iCount) is not present
						Exit function
					Else
						iCounter = 1
						If UBound(aObjects) > 0 Then
							iCounter = cInt(aObjects(1))
						End If
						For iCnt = 1 to iCounter
							If cInt(ObjAccCheck.JavaButton("MoveDown").GetROProperty("enabled")) = 1 then
								Call Fn_Button_Click("Fn_SE_AccountiblityCheck", ObjAccCheck, "MoveDown")
							End If
						Next
					End If
				Fn_SE_AccountiblityCheck = True
				End If
			End If
		'---------------------------------------------------------------------------------
		'[TC11.3(20170403)_NewDevelopment_PoonamC_28Apr2017 : Added case addSource & Target & verify it]
		Case "AddSourceTargetAndVerify"
			' select Scope tab
			 ObjAccCheck.JavaTab("AccountiblityTab").Select "Scope"
			 If dicAccCheck("SelectedSrcObjs") <> "" Then
						' Click on Add source
						Call Fn_Button_Click("Fn_SE_AccountiblityCheck", ObjAccCheck, "AddSetSourceBtn")
						Call Fn_ReadyStatusSync(1)					
					   'Verify Added source node
					   Fn_SE_AccountiblityCheck = Fn_SISW_UI_JavaList_Operations("Fn_SE_AccountiblityCheck","Exist",ObjAccCheck,"SelectedSource",dicAccCheck("SelectedSrcObjs"),"","")
					   If Fn_SE_AccountiblityCheck = False Then
					  		Set ObjAccCheck = Nothing
					  		Exit Function
					   End If
			 End If	
			' Select target Tab
			If dicAccCheck("SelectedTargetTab") <> "" Then	
					'Check Accountability Window state
						If Fn_UI_ObjectExist("Fn_SE_AccountiblityCheck",ObjAccCheck) Then
							If ObjAccCheck.GetROProperty("minimized") <> 1 Then
								ObjAccCheck.Minimize
								Wait 1
							End If
						End If
					' Select Source Tab
					Call Fn_TabFolder_Operation("Select",dicAccCheck("SelectedTargetTab"),"")
					Call Fn_ReadyStatusSync(1)	
					'Maximise the Accountability Window
					ObjAccCheck.Restore
					Wait 1	
			End if

			 If dicAccCheck("SelectedTargetObjs") <> "" Then
					' Click on Add Target
					Call Fn_Button_Click("Fn_SE_AccountiblityCheck", ObjAccCheck, "AddSetTargetBtn")
					Call Fn_ReadyStatusSync(1)
				  'Verify Added source node
				   Fn_SE_AccountiblityCheck = Fn_SISW_UI_JavaList_Operations("Fn_SE_AccountiblityCheck","Exist",ObjAccCheck,"SelectedTarget",dicAccCheck("SelectedTargetObjs"),"","")
				   If Fn_SE_AccountiblityCheck = False Then
				  		Set ObjAccCheck = Nothing
				  		Exit Function
				   End If
		    End If			
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_SE_AccountiblityCheck: Invalid case [ " & sAction & " ].")
	End Select
	'Click on button.
	If dicAccCheck("Button") <> "" Then
		Call Fn_Button_Click("Fn_SE_AccountiblityCheck", ObjAccCheck, dicAccCheck("Button"))
		Wait(3)
		If ObjAccCheck.Exist(3) AND Trim(LCase(dicAccCheck("Button")))="ok"  Then			' Sagar, | 10.0 Porting | 3 Dec 12 | Added code to check existance of Button, resize dialog and click button.
			JavaWindow("SystemsEngineering").JavaWindow("AccountabilityCheck").Maximize										' On 919b build, buttons does not displayed until resize of dialog. dependable on screen resolution.
			JavaWindow("SystemsEngineering").JavaWindow("AccountabilityCheck").JavaButton("OK").Click
		End If 
	End IF
	If Fn_SE_AccountiblityCheck = True Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_SE_AccountiblityCheck: Executed successfully with case [ " & sAction & " ].")
	End If
End Function
 '*********************************************************		Function to Report Operation	***********************************************************************
'Function Name		:				Fn_SE_DetachedTraceabilityOperations

'Description			 :		 		 Perform Operations on "Traceability Report" Dialog

'Parameters			   :	 			1.sAction: DefiningTable:Properties (First is Table Name Compulsory)
'													 2.sNodeName: Node on which we have to perform operation
'													 3.sNewName:New Name in Property
'													4.sColName:Column Name
'													5.sCellValue:Cell Value  												

'Return Value		   : 				True or False

'Pre-requisite			:		 	Must be Selected Trace Link Node

'Examples				:	
'												'Fn_SE_DetachedTraceabilityOperations("DefiningTable:Expand","REQ-000001/A;1-One:REQ-000002/A;1-Two","","","") --->This Case Expand the tree node but it will not  press OK on TraceabilityReport
'												'Fn_SE_DetachedTraceabilityOperations("ComplyingTable:Select","REQ-000001/A;1-One:REQ-000002/A;1-Two","","","") --->This Case Select the tree node but it will not  press OK on TraceabilityReport
'												'Fn_SE_DetachedTraceabilityOperations("DefiningTable:Verify","REQ-000001/A;1-One:REQ-000002/A;1-Two","","","") ---->This Case Verify the tree node but it will not  press OK on TraceabilityReport
'History					 :		
'													Developer Name				Date						Rev. No.						Changes Done			Reviewer			Build
'												-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'														Ketan Raje			   08/09/2011			            1.0									Created					Harshal	A.		20110824
'												-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SE_DetachedTraceabilityOperations(sAction,sNodeName,sNewName,sColName,sCellValue)
	GBL_FAILED_FUNCTION_NAME="Fn_SE_DetachedTraceabilityOperations"
        'Declaring All Varaibles
		Dim aAction,sTableName,iRows,iCounter,sNodePath,sIndex,bFlag, ArrLists, iToolCnt,  sContents,sCellData
		'Declaring All Object's
		Dim objJavaDialogReport,ObjDesc
		'Spliting sAction To retriewe Table name
		aAction=Split(sAction,":")
		sTableName=aAction(0)
		'Setting bFlag
		bFlag=False
		Set objJavaDialogReport=Fn_UI_ObjectCreate("Fn_SE_DetachedTraceabilityOperations", JavaWindow("DefaultWindow").JavaWindow("Shell"))
		If sNodeName<>"" Then
			'Identifying Table
			Select Case sTableName
				   Case "ComplyingTable"
                        'Checking Existance of "ComplyingTable" Table
						If Fn_UI_ObjectExist("Fn_SE_DetachedTraceabilityOperations",objJavaDialogReport.JavaTable("ComplyingTable"))=True Then
							'Retriwing No of rows
							iRows = Fn_Table_GetRowCount("Fn_SE_DetachedTraceabilityOperations",objJavaDialogReport,"ComplyingTable")
							For iCounter = 0 to iRows -1
								objJavaDialogReport.JavaTable("ComplyingTable").SelectRow iCounter
								sNodePath=objJavaDialogReport.JavaTable("ComplyingTable").GetCellData(iCounter,0)
								'Checking "sNodeName" present in table or not
								If Trim(sNodePath) = Trim(sNodeName) Then
									sIndex = Cstr(iCounter)
									bFlag=True
									Exit For
								End If
							Next
							If Cint(iCounter) = Cint(iRows) Then
								Fn_SE_DetachedTraceabilityOperations =False
								Exit Function
							End If
						End If
					Case "DefiningTable"
                        'Checking Existance of "ComplyingTable" Table
						If Fn_UI_ObjectExist("Fn_SE_DetachedTraceabilityOperations",objJavaDialogReport.JavaTable("DefiningTable"))=True Then
							'Retriwing No of rows
							iRows = Fn_Table_GetRowCount("Fn_SE_DetachedTraceabilityOperations",objJavaDialogReport,"DefiningTable")
							For iCounter = 0 to iRows -1
									objJavaDialogReport.JavaTable("DefiningTable").SelectRow iCounter
									sNodePath=objJavaDialogReport.JavaTable("DefiningTable").GetCellData(iCounter,0)
									'Checking "sNodeName" present in table or not
									If Trim(sNodePath) = Trim(sNodeName) Then
												sIndex = Cstr(iCounter)
												bFlag=True
												Exit For
									End If
							Next
							If Cint(iCounter) = Cint(iRows) Then
									Fn_SE_DetachedTraceabilityOperations =False
									Exit Function
									End If
						End If
				End Select 
		End If
        Select Case aAction(1)
			'To Expand Tree Node of table
			Case "Expand"
				If sTableName="ComplyingTable" Then
					objJavaDialogReport.JavaTable("ComplyingTable").SelectRow sIndex
					objJavaDialogReport.JavaTable("ComplyingTable").DoubleClickCell sIndex,0
					Fn_SE_DetachedTraceabilityOperations =True
					Exit Function
				Else
					objJavaDialogReport.JavaTable("DefiningTable").SelectRow sIndex
					objJavaDialogReport.JavaTable("DefiningTable").DoubleClickCell sIndex,0
					Fn_SE_DetachedTraceabilityOperations =True
					Exit Function
				End If
			'To Select Tree Node of table
			Case "Select"
				If sTableName="ComplyingTable" Then
					objJavaDialogReport.JavaTable("ComplyingTable").SelectRow sIndex
					Fn_SE_DetachedTraceabilityOperations =True
					Exit Function
				Else
					objJavaDialogReport.JavaTable("DefiningTable").SelectRow sIndex
					Fn_SE_DetachedTraceabilityOperations =True
					Exit Function
				End If
			'Verifying Node is present or not	
			Case "Verify"
					If bFlag=True Then
						Fn_SE_DetachedTraceabilityOperations =True
						Exit Function
					Else
						Fn_SE_DetachedTraceabilityOperations =False
						Exit Function
					End If
		End Select
	'Clicking on "OK" Button of "Report" Dialog
	Call Fn_Button_Click("Fn_SE_DetachedTraceabilityOperations",objJavaDialogReport,"Ok")
	
Set ObjDesc = Nothing
Set ArrLists = Nothing
Set objJavaDialogReport=Nothing
End Function
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_SE_SaveViewConfiguration

'Description			 :	This Function is used to Save View Configuration in System Engineering

'Parameters			   :  1. StrName : View Configuration Name
'									2. StrDescription : View Configuration Description

'Return Value		   : True \ False

'Pre-requisite		    : Should be in SE Prespective

'Examples			    : Call Fn_SE_SaveViewConfiguration("Config1","Test View Configuration")

'History:
'								   Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'									Sandeep N				21.10.2011				1.0															Sunny R
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SE_SaveViewConfiguration(StrName,StrDescription)
	GBL_FAILED_FUNCTION_NAME="Fn_SE_SaveViewConfiguration"
    'Declaring Variables
	Dim ObjViewConfig,bFlag
	bFlag=False
	Fn_SE_SaveViewConfiguration=False
	'Creating Object of [ SaveViewConfiguration ] Window
	Set ObjViewConfig=JavaWindow("SystemsEngineering").JavaWindow("SaveViewConfiguration")
	If Not ObjViewConfig.Exist(8) Then
		bFlag=Fn_ToolBarOperation("ShowDropdownAndSelectWithInstance", "View Menu:2", "Save View Configuration")
		If bFlag=False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail: Fail to select View Menu [ Save View Configuration ]")
			'Releasing Object of [ SaveViewConfiguration ] Window
			Set ObjViewConfig=Nothing
			Exit Function
		End If
	End If
	'Entering Name for Save View Configuration :- Its mandetory field
	Call Fn_UI_EditBox_Type("Fn_SE_SaveViewConfiguration",ObjViewConfig,"Name",StrName)
	'Entering Description for Save View Configuration
	If StrDescription<>"" Then
		Call Fn_Edit_Box("Fn_SE_SaveViewConfiguration",ObjViewConfig,"Description",StrDescription)
	End If
	bFlag=False
	bFlag=Fn_Button_Click("Fn_SE_SaveViewConfiguration", ObjViewConfig, "Save")
	If bFlag Then
		Fn_SE_SaveViewConfiguration=True
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Pass:Successfully Save View Configuration of Name [ "+StrName+" ]")
	End If
	'Releasing Object of [ SaveViewConfiguration ] Window
	Set ObjViewConfig=Nothing
End Function
'*********************************************************		Function for Report Generation wizard.		***********************************************************************
'Function Name		:        Fn_SE_ReportGenerationWizard  

'Description	    	:        Creates an Part with detail information

'	Example	 		:						  
'								Set dicWizard = CreateObject( "Scripting.Dictionary" )
'										dicWizard.RemoveAll
'										dicWizard("sName") = "TL Complying & Defining Report"
'										dicWizard("sSource") = "Office Template"
'										dicWizard("FillInCriteria") = "yes"
'										dicWizard("sDisplayLocale") = ""
'										dicWizard("sStyleSheets") = "REQ_TraceLink_complying_template"
'										dicWizard("sDatasetName") = ""
'										dicWizard("bLiveIntegration") = "ON"
'										dicWizard("bCreateDataset") = ""
'										dicWizard("sButtons") = "Finish"
'										Environment.Value("TestLogFile") = "D:\Log.txt"
'										Msgbox Fn_SE_ReportGenerationWizard("Set", dicWizard)
'								Set dicWizard = Nothing

'Return Value		: 			True / False

'Pre-requisite	    :		 	Should be logged in

'History		    :		
'													Developer Name				Date						Rev. No.			
'--------------------------------------------------------------------------------------------------------------------------
'													Ketan Raje				     21/10/2011			           1.0								
'--------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SE_ReportGenerationWizard(sAction, dicWizard)
		GBL_FAILED_FUNCTION_NAME="Fn_SE_ReportGenerationWizard"
		Dim iRows, iCount, aButtons
		On Error Resume Next
		Fn_SE_ReportGenerationWizard = False
		'Creating Object for New Part window.

		Set ObjWizard = JavaWindow("SystemsEngineering").JavaWindow("SEWindow").JavaDialog("ReportGenerationWizard")
		 'Creating Object of links on the left side of the window
		Set ObjStaticText =ObjWizard.JavaStaticText("Steps")
        
	   'Check the existence of the "Wizard" Window
		If ObjWizard.Exist (20)  Then		
				Select Case sAction
				Case "Set"
						'Select row from table.
						If dicWizard("sName") <> "" And dicWizard("sSource") <> "" Then
							iRows = ObjWizard.JavaTable("TCTable").GetROProperty("rows")
							If iRows > 0 Then
									For iCount = 0 to iRows-1
										If Trim(Lcase(ObjWizard.JavaTable("TCTable").GetCellData(iCount,"Name"))) = Trim(Lcase(dicWizard("sName"))) Then
												If Trim(Lcase(ObjWizard.JavaTable("TCTable").GetCellData(iCount,"Source"))) = Trim(Lcase(dicWizard("sSource"))) Then
													'Select the row
													ObjWizard.JavaTable("TCTable").SelectRow iCount
													Exit For
												End If
										End If
									Next
									If iCount = iRows Then
											Set ObjWizard = Nothing
											Set ObjStaticText = Nothing
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Name and Source not found in the Table")
											Exit Function
									End If
							Else
								Set ObjWizard = Nothing
								Set ObjStaticText = Nothing
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "No rows found in the table.")
								Exit Function
							End If
						End If
						'Click on "Fill In Criteria" Static Text.
						If Trim(Lcase(dicWizard("FillInCriteria"))) = "yes" Then
								ObjStaticText.SetTOProperty "label","Fill in Criteria"
								ObjStaticText.WaitProperty "enabled" , 1, 20000
								ObjStaticText.Click 1, 1
						End If
						'Enter Report Display Locale
						If dicWizard("sDisplayLocale") <> "" Then
							 Call Fn_Edit_Box("Fn_SE_ReportGenerationWizard",ObjWizard,"ReportDisplayLocale",dicWizard("sDisplayLocale"))
						End If
						'Enter Report Stylesheets
						If dicWizard("sStyleSheets") <> "" Then
							Call Fn_Edit_Box("Fn_SE_ReportGenerationWizard",ObjWizard,"ReportStylesheets",dicWizard("sStyleSheets"))
						End If
						'Set the Dataset Name
						If dicWizard("sDatasetName") <> "" Then
							Call Fn_Edit_Box("Fn_SE_ReportGenerationWizard",ObjWizard,"DatasetName",dicWizard("sDatasetName"))
						End If
						'Set the Live Intergration checkbox
						If dicWizard("bLiveIntegration") <> "" Then
							Call Fn_CheckBox_Set("Fn_SE_ReportGenerationWizard",ObjWizard,"LiveIntegration", dicWizard("bLiveIntegration"))
						End If
						'Set the Create Dataset checkbox
						If dicWizard("bCreateDataset") <> "" Then
							Call Fn_CheckBox_Set("Fn_SE_ReportGenerationWizard",ObjWizard,"CreateDataset", dicWizard("bCreateDataset"))
						End If	
						 'Click on Buttons
						 If dicWizard("sButtons")<>"" Then
							   aButtons = split(dicWizard("sButtons"), ":",-1,1)					   
							   For iCount=0 to Ubound(aButtons)
										'Click on Add Button
										Call Fn_Button_Click("Fn_SE_ReportGenerationWizard", ObjWizard, aButtons(iCount))
										Call Fn_ReadyStatusSync(2)
							   Next
						End If
						Fn_SE_ReportGenerationWizard = True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Function Fn_SE_ReportGenerationWizard executed succesfully with case "&sAction)
				End Select
		End If
	Set ObjWizard = Nothing
	Set ObjStaticText = Nothing
End Function

'*********************************************************		Function to Create Requirement in SE ***********************************************************************
'Function Name		:				Fn_SE_NewDerivedRequirementCreate

'Description			 :		 		 This function is used to Create the Derived Requirement in System Engineering

'Parameters			   :	 			1. strNodeName: Select the Requirement 
'													2. strSpecID: ID of the Requirement
'												   3. strSpecRev: Revision of the Requirement
'												  4. strSpecName: Name of the Requirement
'												  5.strSpecDesc: Description of the Requirement
'                                                 6.sTracelinkType : TraceLink Type                 
'Return Value		   : 				TRUE \ FALSE

'Pre-requisite			:		 		should be in SE Prespective

'Examples				:				 Call Fn_SE_NewDerivedRequirementCreate("Requirement","","","Requirement","Description","")

'History:
'										Developer Name			Date			Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'											Vrushali Wani	    31.10.2011																			
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function Fn_SE_NewDerivedRequirementCreate(strNodeName,strReqID,strReqRev,strReqName,strReqDesc,sTracelinkType)
		GBL_FAILED_FUNCTION_NAME="Fn_SE_NewDerivedRequirementCreate"
		Dim strNodePath,intNodeCount,sTreeItem,intCount,strNodePathC
		Dim ObjReqWnd,bFlag,iRowNo,sReq,sReqFull,sChdReqName1,aItmInfo3,aItmInfo4

		Fn_SE_NewDerivedRequirementCreate=False

		If Fn_UI_ObjectExist("Fn_SE_NewDerivedRequirementCreate",JavaWindow("SystemsEngineering").JavaWindow("NewDerivedRequirement"))=False Then
			Call Fn_MenuOperation("Select","File:New:Derived Requirement...")
		End If
		Set ObjReqWnd=Fn_UI_ObjectCreate("Fn_SE_NewDerivedRequirementCreate",JavaWindow("SystemsEngineering").JavaWindow("NewDerivedRequirement"))

		Call Fn_UI_JavaTree_Expand("Fn_SE_NewDerivedRequirementCreate", ObjReqWnd, "Derived Requirement","Complete List")

		JavaWindow("SystemsEngineering").JavaWindow("NewDerivedRequirement").JavaTree("Derived Requirement").WaitProperty "items count" , micGreaterThan(1)  
		If Fn_UI_JavaTree_NodeExist("Fn_SE_NewDerivedRequirementCreate",ObjReqWnd.JavaTree("Derived Requirement"),"Complete List:"+strNodeName) Then
				strNodePathC="Complete List:"+strNodeName
		Else
				strNodePathC="Most Recently Used:"+strNodeName
		End If
        Call Fn_JavaTree_Select("Fn_SE_NewDerivedRequirementCreate", ObjReqWnd, "Derived Requirement",strNodePathC)
		Call Fn_Button_Click("Fn_SE_NewDerivedRequirementCreate",ObjReqWnd,"Next")

		If strReqID<>"" Then
			'Call Fn_UI_EditBox_Type("Fn_SE_NewDerivedRequirementCreate",ObjReqWnd,"ID",strReqID)
		End If
		If strReqRev<>"" Then
			'Call Fn_UI_EditBox_Type("Fn_SE_NewDerivedRequirementCreate",ObjReqWnd,"Revision",strReqRev)
		End If

		Call Fn_UI_EditBox_Type("Fn_SE_NewDerivedRequirementCreate",ObjReqWnd,"Name",strReqName)
		Call Fn_UI_EditBox_Type("Fn_SE_NewDerivedRequirementCreate",ObjReqWnd,"Description",strReqDesc)
		if sTracelinkType <> "" Then
			Call Fn_List_Select("Fn_SE_NewDerivedRequirementCreate", ObjReqWnd, "Tracelink Type",sTracelinkType)
		End If
		Call Fn_Button_Click("Fn_SE_NewDerivedRequirementCreate",ObjReqWnd,"Finish")
      			Call Fn_ReadyStatusSync(2)
					iRowNo = Fn_SE_BOMTable_RowIndex(strReqName)
					DataTable.GetSheet("Global").AddParameter "NewReqID",""
					DataTable.GetSheet("Global").AddParameter "NewReqRev",""
					sReq=Fn_SE_BOMTableNodeOpeations("GetCellData",iRowNo,0,"","")
					If sReq <> "" Then
						sReqFull=mid(sReq,instr(1,sReq,":")+1,len(sReq))
						sChdReqName1=mid(sReqFull,instr(1,sReqFull,":")+1,len(sReqFull))
						aItmInfo3 = split(sChdReqName1, "/", -1, 1)
						DataTable("NewReqID",dtGlobalSheet)= aItmInfo3(0)
						aItmInfo4 = split(aItmInfo3(1), ";", -1, 1)
						DataTable("NewReqRev",dtGlobalSheet)= aItmInfo4(0)
					End If
		Fn_SE_NewDerivedRequirementCreate=True
		Set ObjReqWnd=Nothing
End Function

'-------------------------------------------------------------------------This Function is used to to perform operation on Save Column Configuration menu--------------------------------------------------------
'Function Name		:	Fn_SE_SaveColumnConfigurationFromBomLine

'Description			:	This Function is used to to perform operation on Save Column Configuration menu

'Parameters			:	   1)sAction:action to be performed
'										2.) strConfigName: Name of Configuration (It should be unique in ConfigurationSaveAs and Add case )
'										3.)strConfigDesc:Description of Configuration

'Return Value		:	True/False

'Pre-requisite		:	Should be details table present
'											
'Examples			:	 Fn_SE_SaveColumnConfigurationFromBomLine("SaveColumnConfiguration",sConfigName,"Test Configuration")
'								
'History:
'									Developer Name												Date								Rev. No.						Changes Done				Build
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'								Sonal Padmawar										14/Nov/2011								1.0																				2011102600
'								Sandeep Navghane								22/06/2012							1.2							Added New object Hierarchy as per 10.0 and added code to handle dyanamic Applet Index
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SE_SaveColumnConfigurationFromBomLine(sAction, strConfigName, strConfigDesc)
	GBL_FAILED_FUNCTION_NAME="Fn_SE_SaveColumnConfigurationFromBomLine"
	  'Declaring Variables
	   Dim ObjColumnConfigWnd
	   Dim objTabFld,StrTabName,i,bFlag,iCounter,objTable
	   Dim objSelectType,objIntNoOfObjects,objItem,icount1,iIndex,stabname

	   Fn_SE_SaveColumnConfigurationFromBomLine=False

	   '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Code to handle Applet as applet index is dyanamically changes
'        bFlag=False		
'		Set objTabFld = JavaWindow("SystemsEngineering").JavaObject("RACTabFolderWidget")
'		i = objTabFld.Object.getSelectedTabIndex
'		StrTabName=objTabFld.Object.getItem(i).text()
'		StrTabName=Split(StrTabName,"-")
'        For iCounter=0 to 12
'			 Window("SEWindow").JavaWindow("WEmbeddedFrame").SetTOProperty "index",iCounter
'			 If Window("SEWindow").JavaWindow("WEmbeddedFrame").JavaTable("CMEBOMTreeTable").Exist(2) Then
'				If InStr(1,trim(Window("SEWindow").JavaWindow("WEmbeddedFrame").JavaTable("CMEBOMTreeTable").Object.getValueAt(0,0).toString()),trim(StrTabName(0))) Then
'					bFlag=True
'					Exit for
'				End If
'			 End If
'		Next
'		If bFlag=false Then
'			Window("SEWindow").JavaWindow("WEmbeddedFrame").SetTOProperty "index",0
'			Exit function
'		End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 

		'Added Code to handle Requirement Opened in BOM table
		Set objSelectType = description.Create()
		objSelectType("Class Name").value = "JavaTab"
		objSelectType("toolkit class").value = "org.eclipse.swt.custom.CTabFolder"						
		Set  objIntNoOfObjects = JavaWindow("DefaultWindow").ChildObjects(objSelectType)
		
		For icount1 = 0 To objIntNoOfObjects.Count-1 Step 1	
			iIndex=objIntNoOfObjects(icount1).Object.getSelectionIndex
			Set objItem=objIntNoOfObjects(icount1).Object.getItem(iIndex)
			StrTabName=trim(objItem.text)
			StrTabName=Split(StrTabName,"-")
			
			If StrTabName(0)="REQ" Then
				StrTabName(0)=StrTabName(1)
			End If
			
			For iCounter=0 to 12
				Window("SEWindow").JavaWindow("WEmbeddedFrame").SetTOProperty "index",iCounter
				If Window("SEWindow").JavaWindow("WEmbeddedFrame").JavaTable("CMEBOMTreeTable").Exist(2) Then
				stabname = trim(Window("SEWindow").JavaWindow("WEmbeddedFrame").JavaTable("CMEBOMTreeTable").Object.getComponentForRow(0).getProperty("bl_indented_title"))
					If stabname<>"" Then
							If InStr(1,stabname,trim(StrTabName(0))) Then
								bFlag=True
								Exit for
							End If
					End If
				End If
			Next
			If bFlag=True Then
				Exit for
			End If
		Next	
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 	 
		If bFlag=false Then
			Exit function
		End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 

	   'verifying existance of SaveColumnConfiguration window
		Set ObjColumnConfigWnd=Window("SEWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Save Column Configuration")

		If not ObjColumnConfigWnd.Exist(5) Then
			'Invoking ApplyColumnConfiguration window				
					'Set objTable = JavaWindow("SystemsEngineering").JavaWindow("WEmbeddedFrame").JavaTable("CMEBOMTreeTable")
					Set objTable = Window("SEWindow").JavaWindow("WEmbeddedFrame").JavaTable("CMEBOMTreeTable")

					objTable.SelectColumnHeader 1,"RIGHT"
					Window("SEWindow").JavaWindow("WEmbeddedFrame").JavaMenu("label:=Save Column Configuration").Select
					If Fn_UI_ObjectExist("Fn_SE_SaveColumnConfigurationFromBomLine",Window("SEWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Save Column Configuration"))=False Then
							Set ObjColumnConfigWnd= nothing
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_SE_SaveColumnConfigurationFromBomLine : Failed to open Save Column configuration dialog.")
							Window("SEWindow").JavaWindow("WEmbeddedFrame").SetTOProperty "index",0
							Exit function
					End If
		End If
		Select Case sAction
				Case "SaveColumnConfiguration"

						'Setting Configuration Name
						Call Fn_UI_EditBox_Type("Fn_SE_SaveColumnConfigurationFromBomLine",ObjColumnConfigWnd,"Name",strConfigName)
						If strConfigDesc<>"" Then
								'Setting Description of configuration 
								Call Fn_Edit_Box("Fn_SE_SaveColumnConfigurationFromBomLine",ObjColumnConfigWnd,"Description",strConfigDesc)
						End If

						'Clicking on save button to create configuration
						Call Fn_Button_Click("Fn_SE_SaveColumnConfigurationFromBomLine",ObjColumnConfigWnd,"Save")
						Fn_SE_SaveColumnConfigurationFromBomLine=True
			
			Case Else
						Fn_SE_SaveColumnConfigurationFromBomLine = False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL :  Fn_SE_SaveColumnConfigurationFromBomLine : Invalid case "&sAction)
	End Select

	If Fn_SE_SaveColumnConfigurationFromBomLine=True Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_SE_SaveColumnConfigurationFromBomLine executed succesfully with case "&sAction)
	End If
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_SE_SaveColumnConfigurationFromBomLine executed succesfully with case "&sAction)
	Window("SEWindow").JavaWindow("WEmbeddedFrame").SetTOProperty "index",0
	Set ObjColumnConfigWnd=Nothing	
End Function
'-------------------------------------------------------------------------This Function is used to to perform operation on Save Column Configuration menu--------------------------------------------------------
'Function Name		:	Fn_SE_ApplyViewConfiguration

'Description			:	This Function is used to to perform operation on Apply View Configuration menu

'Parameters			:	   1) sAction : action to be performed
'						   2) sConfigName: Name of Configuration
'						   3) sConfigDesc: For future Use
'						   4) sAvalableProperties: For future Use
'						   5) sDisplayedColumns: For future Use
'						   6) bShowInternalNames: For future Use
'						   7) sNewName: For future Use
'						   8) sNewDesc: For future Use

'Return Value		:	True/False

'Pre-requisite		:   Details table should be present
'											
'Examples			:	Call Fn_SE_ApplyViewConfiguration("Apply", "* asd", "", "", "", "", "", "")
'Examples			:	Call Fn_SE_ApplyViewConfiguration("Verify", "* asd", "", "", "", "", "", "")
'								
'History:
'				Developer Name			Date			Rev. No.	  Changes Done		
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'				KOustubh Watwe			06/Dec/2011		1.0			  Created																	2011102600
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
 Public Function Fn_SE_ApplyViewConfiguration(sAction, sConfigName, sConfigDesc, sAvalableProperties, sDisplayedColumns, bShowInternalNames, sNewName, sNewDesc)
 	GBL_FAILED_FUNCTION_NAME="Fn_SE_ApplyViewConfiguration"
	Dim objApplyColumnConfig, arrConfig, iItemCnt, iCnt, bReturn 
	Fn_SE_ApplyViewConfiguration = False
	Set objApplyColumnConfig = JavaWindow("SystemsEngineering").JavaWindow("ApplyViewConfiguration")

	' checking existence of  Apply config dialog / window
	If NOT(Fn_UI_ObjectExist("Fn_SE_ApplyViewConfiguration",objApplyColumnConfig)) Then
		bReturn = Fn_ToolBarOperation("ShowDropdownAndSelectWithInstance", "View Menu:2", "Apply View Configuration")
		If NOT(bReturn) Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Fn_SE_ApplyViewConfiguration : Failed to perform Menu operation [ View Menu >> Apply View Configuration ]")
		End If
	End If

	If NOT(Fn_UI_ObjectExist("Fn_SE_ApplyViewConfiguration",objApplyColumnConfig)) Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Fn_SE_ApplyViewConfiguration : Failed to verify existence of  [ Apply View Configuration ] window.")
		Exit function
	End IF
	Select Case sAction
		' - - - - -  - - - - -  - - - - -  - - - - -  - - - - -  - - - - -  - - - - -  - - - - -  - - - - -  - - - - -  - - - - -  - - - - -  - - - - - 
		Case "Apply"
			If Fn_UI_ListItemExist("Fn_SE_ApplyViewConfiguration",objApplyColumnConfig,"Column Configurations",sConfigName) then
				Call Fn_List_Select("Fn_SE_ApplyViewConfiguration",objApplyColumnConfig,"Column Configurations",sConfigName)
				wait 2
				objApplyColumnConfig.Activate
				wait 1
                Call Fn_Button_Click("Fn_SE_ApplyViewConfiguration",objApplyColumnConfig, "Apply")
                If Fn_UI_ObjectExist("Fn_SE_ApplyViewConfiguration",objApplyColumnConfig) Then
					Call Fn_Button_Click("Fn_SE_ApplyViewConfiguration",objApplyColumnConfig, "Close")
				End If
				Fn_SE_ApplyViewConfiguration = True
			End If
			
		' - - - - -  - - - - -  - - - - -  - - - - -  - - - - -  - - - - -  - - - - -  - - - - -  - - - - -  - - - - -  - - - - -  - - - - -  - - - - - 
		Case "Verify"
			arrConfig = Split(sConfigName, "~")
			For iCnt = 0 to UBound(arrConfig)
				Fn_SE_ApplyViewConfiguration = Fn_UI_ListItemExist("Fn_SE_ApplyViewConfiguration",objApplyColumnConfig,"Column Configurations",arrConfig(iCnt))
				If Fn_SE_ApplyViewConfiguration = False then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Fn_SE_ApplyViewConfiguration : Configuration [ " & arrConfig(iCnt) & " ] is not present in the list.")
						Exit for
				Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Fn_SE_ApplyViewConfiguration : Configuration [ " & arrConfig(iCnt) & " ] is present in the list.")
				End If
			Next
			Call Fn_Button_Click("Fn_SE_ApplyViewConfiguration",objApplyColumnConfig, "Close")
		' - - - - -  - - - - -  - - - - -  - - - - -  - - - - -  - - - - -  - - - - -  - - - - -  - - - - -  - - - - -  - - - - -  - - - - -  - - - - - 
		Case "Add"
			' Not yet Implemented
		' - - - - -  - - - - -  - - - - -  - - - - -  - - - - -  - - - - -  - - - - -  - - - - -  - - - - -  - - - - -  - - - - -  - - - - -  - - - - - 
		Case "Modify"
			' Not yet Implemented
		' - - - - -  - - - - -  - - - - -  - - - - -  - - - - -  - - - - -  - - - - -  - - - - -  - - - - -  - - - - -  - - - - -  - - - - -  - - - - - 
		Case "Delete"
				If Fn_UI_ListItemExist("Fn_SE_ApplyViewConfiguration",objApplyColumnConfig,"Column Configurations",sConfigName) Then
					Call Fn_List_Select("Fn_SE_ApplyViewConfiguration",objApplyColumnConfig,"Column Configurations",sConfigName)
					wait 2
					objApplyColumnConfig.Activate
					wait 1
						bReturn=Fn_UI_Object_GetROProperty("Fn_SE_ApplyViewConfiguration",objApplyColumnConfig.JavaButton("DeleteConfig"), "enabled")
						If bReturn="1" Then
							Call Fn_SISW_UI_JavaButton_Operations("Fn_SE_ApplyViewConfiguration", "Click", objApplyColumnConfig, "DeleteConfig")
							If Fn_SISW_UI_Object_Operations("Fn_SE_RequirementSpecCreate", "Exist", JavaWindow("MyTeamcenter").JavaWindow("Delete Column Configuration") , SISW_MICRO_TIMEOUT)=True Then
								Call Fn_SISW_UI_JavaButton_Operations("Fn_SE_ApplyViewConfiguration", "Click", JavaWindow("MyTeamcenter").JavaWindow("Delete Column Configuration"), "Yes")
							End If 
							Fn_SE_ApplyViewConfiguration = True
						Else
							Fn_SE_ApplyViewConfiguration=bReturn
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Fn_SE_ApplyViewConfiguration : Button [ DeleteConfig ] is disabled.")
							Call Fn_SISW_UI_JavaButton_Operations("Fn_SE_ApplyViewConfiguration", "Click", objApplyColumnConfig, "Close")
							Exit Function 
						End If 
				End If 
				Call Fn_SISW_UI_JavaButton_Operations("Fn_SE_ApplyViewConfiguration", "Click", objApplyColumnConfig, "Close")

		' - - - - -  - - - - -  - - - - -  - - - - -  - - - - -  - - - - -  - - - - -  - - - - -  - - - - -  - - - - -  - - - - -  - - - - -  - - - - - 
		Case Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Fn_SE_ApplyViewConfiguration : Invalid case [ " & sAction & " ] ")
		' - - - - -  - - - - -  - - - - -  - - - - -  - - - - -  - - - - -  - - - - -  - - - - -  - - - - -  - - - - -  - - - - -  - - - - -  - - - - - 
	End Select
	If Fn_SE_ApplyViewConfiguration = True Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Fn_SE_ApplyViewConfiguration : Executed successfully with Case [ " & sAction & " ] ")
	End If
	Set objApplyColumnConfig = nothing
End Function


'''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''''/$$$$
''''/$$$$   FUNCTION NAME   :  Fn_SE_DiagramOperations(sAction,sId,sName,sDescription,bOpenOnCreate,sAppDomain,sTemplateType,sTemplate,sButton,sInfo1,sInfo2)
''''/$$$$
''''/$$$$   DESCRIPTION        :  This function will  perform Operations on the Create New Diagram Dialog
''''/$$$$ 
''''/$$$$   PRE-REQUISITES        :  The Document mapping tree should be present
''''/$$$$
''''/$$$$  PARAMETERS   : 		sAction : Action to be performed
''''/$$$$										sId : Valid ID
''''/$$$$										sName : Valid Name to Be entered
''''/$$$$										sDescription : Valid Description to be entered
'''/$$$$										bOpenOnCreate : to determine wheteg=her to open after creation
'''/$$$$										sAppDomain : Valid Application Domain
'''/$$$$										sTemplateType : Valid Template Type
'''/$$$$										sTemplate : Valid Template Name
'''/$$$$										sButton :  Valid Button Name to be clicked
''''/$$$$										sInfo1: For Future Use
''''/$$$$										sInfo2 : For Future Use
''''/$$$$	
''''/$$$$		Return Value : 				True or False
''''/$$$$
''''/$$$$    Function Calls       :   Fn_WriteLogFile(), Fn_MenuOperation(), Fn_ReadyStatusSync
''''/$$$$										
''''/$$$$
''''/$$$$		HISTORY           :   		AUTHOR                 DATE        VERSION
''''/$$$$
''''/$$$$    CREATED BY     :   SHREYAS          02/02/2012         1.0
''''/$$$$
''''/$$$$    REVIWED BY     :  Shreyas			02/02/2012            1.0
''''/$$$$
''''/$$$$		How To Use :    bReturn=Fn_SE_DiagramOperations("Create","","Qwerty","","","FunctionalModeling","Visio Template","FUNCTIONAL_MODELING_TEMPLATE","OK","","")
''''/$$$$									  bReturn=Fn_SE_DataDictionarySearchDialogOperations("VerifyTableValues","Classification Root:MyLib_40586  [2]*Select","","false","","","Object Name",aVal)
''''/$$$$							
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Public Function Fn_SE_DiagramOperations(sAction,sId,sName,sDescription,bOpenOnCreate,sAppDomain,sTemplateType,sTemplate,sButton,sInfo1,sInfo2)
GBL_FAILED_FUNCTION_NAME="Fn_SE_DiagramOperations"
Dim sValues,bReturn,iCounter,objDiagram

Fn_SE_DiagramOperations=false
Set objDiagram= JavaWindow("SystemsEngineering").JavaWindow("CreateDiagram")
If objDiagram.Exist(5)=false Then
	bReturn=Fn_MenuOperation("Select","File:New:Create Diagram")
	If bReturn=false Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SE_DiagramOperations ] New Diagram Dialog Not invoked")
				 Fn_SE_DiagramOperations = False
				 Exit function
	Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SE_DiagramOperations ] New Diagram Dialog Successfully invoked")	
				Call Fn_ReadyStatusSync(2)
	End If
End If

Select Case sAction

		Case "Create"

'set the Id
			If sId<>"" Then
				objDiagram.JavaEdit("ID").Set sId
					If Err.number < 0 Then
                		Fn_SE_DiagramOperations=False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SE_DiagramOperations ] Case [ " & sAction  & " ] Id Not Set")
						Exit function
					Else
						Fn_SE_DiagramOperations=true
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SE_DiagramOperations ] Case [ " & sAction  & " ] Successfully Set the Id as ["+sId+"]")
					End If
			End If

'Set the Description
			If sDescription<>"" Then				 
				objDiagram.JavaEdit("Description").Set sDescription
					If Err.number < 0 Then
                		Fn_SE_DiagramOperations=False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SE_DiagramOperations ] Case [ " & sAction  & " ] Description Not Set")
						Exit function
					Else
						Fn_SE_DiagramOperations=true
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SE_DiagramOperations ] Case [ " & sAction  & " ] Successfully Set the Description as ["+sDescription+"]")
					End If
			End If

'Set the Name
			If sName<>"" Then
				objDiagram.JavaEdit("Name").Set sName
					If Err.number < 0 Then
                		Fn_SE_DiagramOperations=False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SE_DiagramOperations ] Case [ " & sAction  & " ] Name Not Set")
						Exit function
					Else
						Fn_SE_DiagramOperations=true
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SE_DiagramOperations ] Case [ " & sAction  & " ] Successfully Set the Name as ["+sName+"]")
					End If
			Else
                		Fn_SE_DiagramOperations=False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SE_DiagramOperations ] Case [ " & sAction  & " ] Name is a Mandatory Parameter. "+vbnewline+"Cannot Proceed Further")
						Exit function	
			End If

'To open or to not open on Create
		If bOpenOnCreate<>"" Then
				objDiagram.JavaCheckBox("Open").Set bOpenOnCreate
        	End if

'Set the Application Domain

			If sAppDomain<>"" Then
			
				objDiagram.JavaList("ApplicationDomain").Select sAppDomain
				Wait(2)
					If Err.number < 0 Then
                		Fn_SE_DiagramOperations=False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SE_DiagramOperations ] Case [ " & sAction  & " ] Application Domain Not Set")
						Exit function
					Else
						Fn_SE_DiagramOperations=true
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SE_DiagramOperations ] Case [ " & sAction  & " ] Successfully Set the Application Domain as ["+sAppDomain+"]")
					End If
			End If

'Set the Template Type

			If sTemplateType<>"" Then
				objDiagram.JavaList("TemplateType").Select sTemplateType
					Wait(2)
					If Err.number < 0 Then
                		Fn_SE_DiagramOperations=False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SE_DiagramOperations ] Case [ " & sAction  & " ] Template Type Not Set")
						Exit function
					Else
						Fn_SE_DiagramOperations=true
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SE_DiagramOperations ] Case [ " & sAction  & " ] Successfully Set the Template Typeas ["+sTemplateType+"]")
					End If
			End If

'Set the Template

			If sTemplate<>"" Then
				objDiagram.JavaList("SelectTemplate").Select sTemplate
					Wait(2)
					If Err.number < 0 Then
                		Fn_SE_DiagramOperations=False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SE_DiagramOperations ] Case [ " & sAction  & " ] Template Not Set")
						Exit function
					Else
						Fn_SE_DiagramOperations=true
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SE_DiagramOperations ] Case [ " & sAction  & " ] Successfully Set the Template as ["+sTemplate+"]")
					End If
			End If

'Click on the Required Button
				Err.Clear
			objDiagram.JavaButton(sButton).WaitProperty "enabled", "1", 1000
			objDiagram.JavaButton(sButton).Click micLeftBtn
					If Err.number < 0 Then
                		Fn_SE_DiagramOperations=False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SE_DiagramOperations ] Case [ " & sAction  & " ] Failed to Click the Desired Button")
						Exit function
					Else
						Fn_SE_DiagramOperations=true
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SE_DiagramOperations ] Case [ " & sAction  & " ] Successfully Clicked the Desired Button ["+sButton+"]")
					End If
End Select
Set objDiagram=nothing
End function




'''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''''/$$$$
''''/$$$$   FUNCTION NAME   :  Fn_SE_DataDictionarySearchDialogOperations(sAction,sHeirarchyTreeNode,sSearchCreiteria,bSearchObjectId,sObjectId,sButton,sInfo1,sInfo2)
''''/$$$$
''''/$$$$   DESCRIPTION        :  This function will  perform Operations on the DataDictionarySearchDialog
''''/$$$$ 
''''/$$$$   PRE-REQUISITES        :  System Engineering Perspective Should Be Open
''''/$$$$
''''/$$$$  PARAMETERS   : 		sAction : Action to be performed
''''/$$$$										sHeirarchyTreeNode : Valid Tree Node
''''/$$$$										sSearchCreiteria : Valid Search Crieteria
''''/$$$$										bSearchObjectId : Boolean Parameter to Enter Search Crieteria
'''/$$$$										sObjectId : Object Id To Search
'''/$$$$										sButton : Valid Button To Be Clicked
''''/$$$$										sInfo1: For Future Use
''''/$$$$										sInfo2 : For Future Use
''''/$$$$	
''''/$$$$		Return Value : 				True or False
''''/$$$$
''''/$$$$    Function Calls       :   Fn_WriteLogFile(), Fn_MenuOperation(), Fn_ReadyStatusSync
''''/$$$$										
''''/$$$$
''''/$$$$		HISTORY           :   		AUTHOR                 DATE        VERSION
''''/$$$$
''''/$$$$    CREATED BY     :   SHREYAS          31/01/2012         1.0
''''/$$$$
''''/$$$$    REVIWED BY     :  Shreyas			31/01/2012            1.0
''''/$$$$
''''/$$$$		How To Use :    bReturn=Fn_SE_DataDictionarySearchDialogOperations("AddObject","Classification Root:MyLib  [1]*Select","","false","000045","OK","","")
''''/$$$$							
''''/$$$$	
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$


Public Function Fn_SE_DataDictionarySearchDialogOperations(sAction,sHeirarchyTreeNode,sSearchCreiteria,bSearchObjectId,sObjectId,sButton,sInfo1,sInfo2)
	GBL_FAILED_FUNCTION_NAME="Fn_SE_DataDictionarySearchDialogOperations"
   Fn_SE_DataDictionarySearchDialogOperations=false
   Dim bReturn,sValue,objDialog,iCounter,aNodes,i,bFlag,sColIndex,sCols
   Dim Counter,aDetails,aObjectId,iCnt
   Set objDialog= JavaWindow("SystemsEngineering").JavaWindow("Search").JavaDialog("DataDictionarySearch")
bFlag=False

If not objDialog.Exist Then
	   'invoke the DataDictionarySearchDialog
	   bReturn=Fn_ToolbatButtonClick("Add Signals From Library")
	   Call Fn_ReadyStatusSync(1)
	   If bReturn=true Then
				Fn_SE_DataDictionarySearchDialogOperations=true
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Successfully invoked the [DataDictionarySearchDialog] Dialog")
		Else
				Fn_SE_DataDictionarySearchDialogOperations=False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fail to invoke the [DataDictionarySearchDialog] Dialog")
				Exit function
	   End If
End If

   Select Case sAction
	 
		Case "AddObject"

			'Expand the Parent Node in the Heirarchy Tree

			objDialog.JavaTree("HierarchyTree").Select "Classification Root"
			If err.number<0 Then
					Fn_SE_DataDictionarySearchDialogOperations=False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fail to select the Parent Node [Classification Root]")
					objDialog.Close
					Exit function
			Else
					Fn_SE_DataDictionarySearchDialogOperations=true
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Successfully selected the Parent Node [Classification Root]")
					Call Fn_ReadyStatusSync(1)
			End If

'			Expand the Parent Node
			objDialog.JavaTree("HierarchyTree").OpenContextMenu("Classification Root")
			objDialog.JavaMenu("MenuSelect").SetTOProperty "label", "ExpandAll"
			objDialog.JavaMenu("MenuSelect").Select  
			wait(2)
			objDialog.JavaTree("HierarchyTree").OpenContextMenu("Classification Root")
			objDialog.JavaMenu("MenuSelect").SetTOProperty "label", "ExpandAll"
			wait(2) 'Added by Manish Singh on 12-Dec-2012
			objDialog.JavaMenu("MenuSelect").Select  
			If err.number<0 Then
					Fn_SE_DataDictionarySearchDialogOperations=False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fail to Expand the Parent Node [Classification Root]")
					objDialog.Close
					Exit function
			Else
					Fn_SE_DataDictionarySearchDialogOperations=true
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Successfully Expanded the Parent Node [Classification Root]")
					Call Fn_ReadyStatusSync(1)
			End If

		'Select the Required Node From the 
		If sHeirarchyTreeNode<>"" Then
			aNodes=split(sHeirarchyTreeNode,"*",-1,1)
					'Select the Node First
						objDialog.JavaTree("HierarchyTree").Select aNodes(0)
						If err.number<0 Then
								Fn_SE_DataDictionarySearchDialogOperations=False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fail to select the  Node ["+aNodes(0)+"]")
								objDialog.Close
								Exit function
						Else
								Fn_SE_DataDictionarySearchDialogOperations=true
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Successfully selected the Node ["+aNodes(0)+"]")
								Call Fn_ReadyStatusSync(1)
						End If
					objDialog.JavaTree("HierarchyTree").OpenContextMenu(aNodes(0))
					objDialog.JavaMenu("MenuSelect").SetTOProperty "label", aNodes(1)
					objDialog.JavaMenu("MenuSelect").Select  
					Call Fn_ReadyStatusSync(1)
		End If
		
			'Click on Clear Button
			If objDialog.JavaButton("Clear").GetROProperty("enabled")=1 Then
				objDialog.JavaButton("Clear").Click micLeftBtn
					If err.number<0 Then
							Fn_SE_DataDictionarySearchDialogOperations=False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fail to Click on the Clear Button")
							objDialog.Close
							Exit function
					Else
							Fn_SE_DataDictionarySearchDialogOperations=true
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Successfully Clicked on the Clear Button")
							Call Fn_ReadyStatusSync(1)
					End If
			End If

			'Enter the Search Creiteria in the Object Id Field
			If sObjectId<>""  Then
				'Enter the Value in the Object Id Field
				If cBool(bSearchObjectId)=true Then
					objDialog.JavaEdit("Object ID").Set sObjectId
						If err.number<0 Then
								Fn_SE_DataDictionarySearchDialogOperations=False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fail to set the value ["+sObjectId+"] in the Object Id Field")
								objDialog.Close
								Exit function
						Else
								Fn_SE_DataDictionarySearchDialogOperations=true
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS:  Successfully set the value ["+sObjectId+"] in the Object Id Field")
								Call Fn_ReadyStatusSync(1)
						End If
				End If

				'Click on the Search Button
				objDialog.JavaButton("Search").Click micLeftBtn
				If err.number<0 Then
						Fn_SE_DataDictionarySearchDialogOperations=False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fail to Click on the Search Button")
						objDialog.Close
						Exit function
				Else
						Fn_SE_DataDictionarySearchDialogOperations=true
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Successfully Clicked on the Search Button")
						Call Fn_ReadyStatusSync(1)
				End If
			End If



			'Activate the table Tab
			objDialog.JavaTab("SearchCriteria").Select "Table"
			If err.number<0 Then
					Fn_SE_DataDictionarySearchDialogOperations=False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fail to Activate the table Tab")
					objDialog.Close
					Exit function
			Else
					Fn_SE_DataDictionarySearchDialogOperations=true
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Successfully Activated the table Tab")
					Call Fn_ReadyStatusSync(1)
			End If

			'Select the Search Result From the Table
			If instr(1,sObjectId,",")>0 Then
					aDetails=split(sObjectId,",",-1,1)
					For Counter=0 to uBound(aDetails)
'							sRows=objDialog.JavaTable("ResultTable").GetROProperty ("rows")
							sRows=Fn_UI_Object_GetROProperty("Fn_SE_DataDictionarySearchDialogOperations",objDialog.JavaTable("ResultTable"), "rows")
							For iCounter= 0 to sRows-1
								sValue=objDialog.JavaTable("ResultTable").GetCellData (iCounter,2)
								If instr(1,aDetails(Counter),sValue)>0 Then
										objDialog.JavaTable("ResultTable").ExtendRow (iCounter)
										If err.number<0 Then
												Fn_SE_DataDictionarySearchDialogOperations=False
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fail to add the Value ["+sObjectId+"]")
												objDialog.Close
												Exit function
										Else
												Fn_SE_DataDictionarySearchDialogOperations=true
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Successfully added the Value ["+sObjectId+"]")
												Call Fn_ReadyStatusSync(1)
									End If
								End If
							Next
					Next

					'Click on Add Button
					objDialog.JavaButton("Add").Click micLeftBtn
					Call Fn_ReadyStatusSync(1)
		Else
'					sRows=objDialog.JavaTable("ResultTable").GetROProperty ("rows")
					sRows=Fn_UI_Object_GetROProperty("Fn_SE_DataDictionarySearchDialogOperations",objDialog.JavaTable("ResultTable"), "rows")
					For iCounter= 0 to sRows-1
						sValue=objDialog.JavaTable("ResultTable").GetCellData (iCounter,2)
						If instr(1,sObjectId,sValue)>0 Then
								objDialog.JavaTable("ResultTable").SelectRow iCounter
								'Click on Add Button
								objDialog.JavaButton("Add").Click micLeftBtn
								If err.number<0 Then
										Fn_SE_DataDictionarySearchDialogOperations=False
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fail to add the Value ["+sObjectId+"]")
										objDialog.Close
										Exit function
								Else
										Fn_SE_DataDictionarySearchDialogOperations=true
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Successfully added the Value ["+sObjectId+"]")
										Call Fn_ReadyStatusSync(1)
								End If
						End If
					Next
		End If

			If  sButton<>"" Then
					'Click on the Required Button Of the Dialog
					objDialog.JavaButton(sButton).Click micLeftBtn
					If err.number<0 Then
							Fn_SE_DataDictionarySearchDialogOperations=False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fail to Click the Button ["+sButton+"]")
							objDialog.Close
							Exit function
					Else
							Fn_SE_DataDictionarySearchDialogOperations=true
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Successfully Clicked the Button ["+sButton+"]")
							Call Fn_ReadyStatusSync(1)
					End If
			End If


		Case "SetFilterCriteria"
					'Click on drop Down button
					objDialog.JavaObject("JLabelDropDown").Click 3,6,"LEFT"
					If err.number<0 Then
								Fn_SE_DataDictionarySearchDialogOperations=False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fail to Click on JLabelDropDown button")
								Exit function
					End If
					'Select from Java Menu "Filter Classification Hierarchy"
					objDialog.JavaMenu("MenuSelect").SetTOProperty "label", "Filter Classification Hierarchy"
					objDialog.JavaMenu("MenuSelect").Select
					wait(3)
					'Select  "Select All"  from the 'SetFilterCriteria'  Dialog 
					objDialog.JavaDialog("SetFilterCriteria").JavaList("SelectLibrarytypes").Select sInfo1
					wait(3)
					If err.number<0 Then
								Fn_SE_DataDictionarySearchDialogOperations=False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fail to Select [ Select All ] from [ SetFilterCriteria]  Dialog")
								Exit function
					End If
					'Click on 'OK' button
					objDialog.JavaDialog("SetFilterCriteria").JavaButton("OK").Click
					wait(3)
					If err.number<0 Then
								Fn_SE_DataDictionarySearchDialogOperations=False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fail to Click on [ OK ] from [ SetFilterCriteria]  Dialog")
								Exit function
					Else
							Fn_SE_DataDictionarySearchDialogOperations=True
							Call Fn_ReadyStatusSync(1)
					End If
			
		Case "VerifyTableValues"

			'Expand the Parent Node in the Heirarchy Tree

			objDialog.JavaTree("HierarchyTree").Select "Classification Root"
			If err.number<0 Then
					Fn_SE_DataDictionarySearchDialogOperations=False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fail to select the Parent Node [Classification Root]")
					Exit function
			Else
					Fn_SE_DataDictionarySearchDialogOperations=true
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Successfully selected the Parent Node [Classification Root]")
					Call Fn_ReadyStatusSync(1)
			End If

'			Expand the Parent Node
			objDialog.JavaTree("HierarchyTree").OpenContextMenu("Classification Root")
			objDialog.JavaMenu("MenuSelect").SetTOProperty "label", "ExpandAll"
			objDialog.JavaMenu("MenuSelect").Select  
			wait(2)
			objDialog.JavaTree("HierarchyTree").OpenContextMenu("Classification Root")
			objDialog.JavaMenu("MenuSelect").SetTOProperty "label", "ExpandAll"
			objDialog.JavaMenu("MenuSelect").Select  
			If err.number<0 Then
					Fn_SE_DataDictionarySearchDialogOperations=False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fail to Expand the Parent Node [Classification Root]")
					objDialog.Close
					Exit function
			Else
					Fn_SE_DataDictionarySearchDialogOperations=true
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Successfully Expanded the Parent Node [Classification Root]")
					Call Fn_ReadyStatusSync(1)
			End If

		'Select the Required Node From the 
		If sHeirarchyTreeNode<>"" Then
			aNodes=split(sHeirarchyTreeNode,"*",-1,1)
					'Select the Node First
						objDialog.JavaTree("HierarchyTree").Select aNodes(0)
						If err.number<0 Then
								Fn_SE_DataDictionarySearchDialogOperations=False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fail to select the  Node ["+aNodes(0)+"]")
								objDialog.Close
								Exit function
						Else
								Fn_SE_DataDictionarySearchDialogOperations=true
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Successfully selected the Node ["+aNodes(0)+"]")
								Call Fn_ReadyStatusSync(1)
						End If
					objDialog.JavaTree("HierarchyTree").OpenContextMenu(aNodes(0))
					objDialog.JavaMenu("MenuSelect").SetTOProperty "label", aNodes(1)
					objDialog.JavaMenu("MenuSelect").Select  
					Call Fn_ReadyStatusSync(1)
		End If
		
			'Click on Clear Button
			If objDialog.JavaButton("Clear").GetROProperty("enabled")=1 Then
				objDialog.JavaButton("Clear").Click micLeftBtn
					If err.number<0 Then
							Fn_SE_DataDictionarySearchDialogOperations=False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fail to Click on the Clear Button")
							objDialog.Close
							Exit function
					Else
							Fn_SE_DataDictionarySearchDialogOperations=true
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Successfully Clicked on the Clear Button")
							Call Fn_ReadyStatusSync(1)
					End If
			End If

			'Enter the Search Creiteria in the Object Id Field
			If sObjectId<>""  Then
				'Enter the Value in the Object Id Field
				If cBool(bSearchObjectId)=true Then
					objDialog.JavaEdit("Object ID").Set sObjectId
						If err.number<0 Then
								Fn_SE_DataDictionarySearchDialogOperations=False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fail to set the value ["+sObjectId+"] in the Object Id Field")
								objDialog.Close
								Exit function
						Else
								Fn_SE_DataDictionarySearchDialogOperations=true
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS:  Successfully set the value ["+sObjectId+"] in the Object Id Field")
								Call Fn_ReadyStatusSync(1)
						End If
				End If

				'Click on the Search Button
				objDialog.JavaButton("Search").Click micLeftBtn
				If err.number<0 Then
						Fn_SE_DataDictionarySearchDialogOperations=False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fail to Click on the Search Button")
						objDialog.Close
						Exit function
				Else
						Fn_SE_DataDictionarySearchDialogOperations=true
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Successfully Clicked on the Search Button")
						Call Fn_ReadyStatusSync(1)
				End If
			End If



			'Activate the table Tab
			objDialog.JavaTab("SearchCriteria").Select "Table"
			If err.number<0 Then
					Fn_SE_DataDictionarySearchDialogOperations=False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fail to Activate the table Tab")
					objDialog.Close
					Exit function
			Else
					Fn_SE_DataDictionarySearchDialogOperations=true
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Successfully Activated the table Tab")
					Call Fn_ReadyStatusSync(1)
			End If

			'get the Column name
			sCols=objDialog.JavaTable("ResultTable").GetROProperty ("cols")
			For iCounter= 0 to cInt( sCols)-1
				sValue=objDialog.JavaTable("ResultTable").GetColumnName(iCounter) 
				If instr(1,sValue,sInfo1)>0 Then
					sColIndex=iCounter
					bFlag=true
					Exit for
				End If
			Next

			If bFlag=true Then
					Fn_SE_DataDictionarySearchDialogOperations=true
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Successfully retrieved Column Index as ["+cstr(sColIndex)+"]")
			Else
					Fn_SE_DataDictionarySearchDialogOperations=False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fail to retrieve Column Index")
					objDialog.Close
					Exit function
					
			End If


			'Verify the Search Creiteria

				For i =0 to ubound(sInfo2)
					sRows=objDialog.JavaTable("ResultTable").GetROProperty ("rows")
					For iCounter= 0 to sRows-1
							sValue=objDialog.JavaTable("ResultTable").GetCellData(i,sColIndex)
							If lCAse(trim(sValue))=Lcase(trim(sInfo2(i))) Then
								bFlag=True
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Successfully Verified the Value ["+sInfo2(i)+"]")
								Exit For
							End If
					Next
			Next

			If bFlag=true Then
					Fn_SE_DataDictionarySearchDialogOperations=true
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Verifications Complete")
			Else
					Fn_SE_DataDictionarySearchDialogOperations=False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fail to Verify the Value")
					objDialog.Close
					Exit function
			End If

		Case "SearchandVerify_AttrValue"                                            'Added By Pooja S :  8-Feb-2012

			'Expand the Parent Node in the Heirarchy Tree
			objDialog.JavaTree("HierarchyTree").Select "Classification Root"
			Call Fn_ReadyStatusSync(1)
			If err.number<0 Then
					Fn_SE_DataDictionarySearchDialogOperations=False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fail to select the Parent Node [Classification Root]")
					Exit function
			End If

'			Expand the Parent Node
			objDialog.JavaTree("HierarchyTree").OpenContextMenu("Classification Root")
			objDialog.JavaMenu("MenuSelect").SetTOProperty "label", "ExpandAll"
			objDialog.JavaMenu("MenuSelect").Select  
			wait(2)
			objDialog.JavaTree("HierarchyTree").OpenContextMenu("Classification Root")
			objDialog.JavaMenu("MenuSelect").SetTOProperty "label", "ExpandAll"
			objDialog.JavaMenu("MenuSelect").Select  
			Call Fn_ReadyStatusSync(1)
			If err.number<0 Then
					Fn_SE_DataDictionarySearchDialogOperations=False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fail to Expand the Parent Node [Classification Root]")
					objDialog.Close
					Exit function
			End If

		'Select the Required Node From the 
		If sHeirarchyTreeNode<>"" Then
			aNodes=split(sHeirarchyTreeNode,"*",-1,1)
						'Select the Node First
						objDialog.JavaTree("HierarchyTree").Select aNodes(0)
						Call Fn_ReadyStatusSync(1)
						If err.number<0 Then
								Fn_SE_DataDictionarySearchDialogOperations=False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fail to select the  Node ["+aNodes(0)+"]")
								objDialog.Close
								Exit function
						End If
						objDialog.JavaTree("HierarchyTree").OpenContextMenu(aNodes(0))
						objDialog.JavaMenu("MenuSelect").SetTOProperty "label", aNodes(1)
						objDialog.JavaMenu("MenuSelect").Select  
						Call Fn_ReadyStatusSync(2)
						wait(3)
		End If
		
			'Click on Clear Button
			If objDialog.JavaButton("Clear").GetROProperty("enabled")=1 Then
						objDialog.JavaButton("Clear").Click micLeftBtn
						Call Fn_ReadyStatusSync(2)
						If err.number<0 Then
									Fn_SE_DataDictionarySearchDialogOperations=False
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fail to Click on the Clear Button")
									objDialog.Close
									Exit function
						End If
			End If

			'Enter the Search Creiteria in the Object Id Field
			If sObjectId<>""  Then
					'Enter the Value in the Object Id Field
					If bSearchObjectId="AttrSearch" Then
								iCnt=4
								If instr(1,sObjectId,"~") Then
										aObjectId = split(sObjectId,"~",-1,1)
								Else
										aObjectId = Array(sObjectId)
								End If

							    For iCounter=0 To Ubound(aObjectId)
											objDialog.JavaEdit("SearchCreiteria").SetTOProperty "Index",iCnt
											wait(2)
											objDialog.JavaEdit("SearchCreiteria").Set aObjectId(iCounter)
											Call Fn_ReadyStatusSync(2)
											If err.number<0 Then
														Fn_SE_DataDictionarySearchDialogOperations=False
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fail to set the value ["+sObjectId+"] in the Object Id Field")
														objDialog.Close
														Exit function
											End If
											iCnt=iCnt+1
								Next
					End If

					'Click on the Search Button
					objDialog.JavaButton("Search").Click micLeftBtn
					Call Fn_ReadyStatusSync(2)
					If err.number<0 Then
							Fn_SE_DataDictionarySearchDialogOperations=False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fail to Click on the Search Button")
							objDialog.Close
							Exit function
					End If
			End If

'				- - - - - - - - - - - - - - - - - - - -For Future Use - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			'Activate the Properties Tab    
			objDialog.JavaTab("SearchCriteria").Select "Properties"
			Call Fn_ReadyStatusSync(2)
			If err.number<0 Then
					Fn_SE_DataDictionarySearchDialogOperations=False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fail to Activate the table Tab")
					objDialog.Close
					Exit function
			End If
'			- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 

			'Activate the table Tab
			objDialog.JavaTab("SearchCriteria").Select "Table"
			Call Fn_ReadyStatusSync(2)
			If err.number<0 Then
					Fn_SE_DataDictionarySearchDialogOperations=False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fail to Activate the table Tab")
					objDialog.Close
					Exit function
			End If

			'get the Column name
			sCols=objDialog.JavaTable("ResultTable").GetROProperty ("cols")
			For iCounter= 0 to cInt( sCols)-1
				sValue=objDialog.JavaTable("ResultTable").GetColumnName(iCounter) 
				If instr(1,sValue,sInfo1)>0 Then
					sColIndex=iCounter
					bFlag=true
					Exit for
				End If
			Next

			If bFlag=true Then
					Fn_SE_DataDictionarySearchDialogOperations=true
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Successfully retrieved Column Index as ["+cstr(sColIndex)+"]")
			Else
					Fn_SE_DataDictionarySearchDialogOperations=False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fail to retrieve Column Index")
					objDialog.Close
					Exit function			
			End If


			'Verify the Search Creiteria
				For i =0 to ubound(sInfo2)
					sRows=objDialog.JavaTable("ResultTable").GetROProperty ("rows")
					For iCounter= 0 to sRows-1
							sValue=objDialog.JavaTable("ResultTable").GetCellData(i,sColIndex)
							If lCAse(trim(sValue))=Lcase(trim(sInfo2(i))) Then
								bFlag=True
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Successfully Verified the Value ["+sInfo2(i)+"]")
								Exit For
							End If
					Next
			Next

			If bFlag=true Then
					Fn_SE_DataDictionarySearchDialogOperations=true
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Verifications Complete")
			Else
					Fn_SE_DataDictionarySearchDialogOperations=False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fail to Verify the Value")
					objDialog.Close
					Exit function
			End If

  End Select

Set objDialog=Nothing
End Function

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_SE_DataDictionarySearchTreeOperations

'Description			 :	Function Used to perform operations on Data Dictionary Search Tree

'Parameters			   :   '1.StrAction: Action Name
'										 2.StrNodeName: Node Path
'									3.StrMenu : Pop up menu name
'
'Return Value		   : 	True / False

'Pre-requisite			:	Data Dictionary Search dialog should be open

'Examples				:   Call Fn_SE_DataDictionarySearchTreeOperations("Expand","Classification Root:My_Group_05217~2","")
'									Call Fn_SE_DataDictionarySearchTreeOperations("Select","Classification Root:My_Group_05217~2:MyLibs_05217  [5]","")
'									Call Fn_SE_DataDictionarySearchTreeOperations("PopupMenuSelect","Classification Root:My_Group_05217","ExpandAll")
'									Call Fn_SE_DataDictionarySearchTreeOperations("Exist","Classification Root:My_Group_05217","")
'										
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												08-Feb-2012								1.0																						Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Public Function Fn_SE_DataDictionarySearchTreeOperations(StrAction,StrNodeName,StrMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_SE_DataDictionarySearchTreeOperations"
    Dim aNodes, iCount, arrNodes, iNodes, iIndex, iChildInstance, iHashCounter, sRealPath,sPath,iChildCount,aMenuList,intCount
	Dim intNodeCount,sTreeItem
   Fn_SE_DataDictionarySearchTreeOperations=False
	If InStr(StrNodeName,"~")>0 Then
		sPath = ""
		sRealPath = ""
		aNodes = Split(StrNodeName,":")
		For iCount = 0 to Ubound(aNodes)
			If InStr(1,aNodes(iCount),"~") <> 0 Then
				arrNodes = Split(aNodes(iCount),"~")
					iNodes = JavaWindow("SystemsEngineering").JavaWindow("Search").JavaDialog("DataDictionarySearch").JavaTree("HierarchyTree").GetROProperty("items count")-1
					For iIndex = 0 to iNodes
						If Trim(Lcase(JavaWindow("SystemsEngineering").JavaWindow("Search").JavaDialog("DataDictionarySearch").JavaTree("HierarchyTree").GetItem(iIndex))) = Trim(Lcase(sPath)) Then
							Exit For
						End If
					Next
					iChildInstance = 0
					iHashCounter = 0
					For iChildCount = iIndex+1 to iNodes
						iHashCounter = iHashCounter + 1
						If Trim(Lcase(JavaWindow("SystemsEngineering").JavaWindow("Search").JavaDialog("DataDictionarySearch").JavaTree("HierarchyTree").GetItem(iChildCount))) = Trim(Lcase(sPath&":"&arrNodes(0))) Then
							iChildInstance = iChildInstance + 1
							If Cint(iChildInstance) = Cint(arrNodes(1)) Then
								sRealPath = sRealPath&":#"&iHashCounter-1
								sPath = sPath&":"&arrNodes(0)
								Exit For
							End If
						End If
					Next		
			Else
				If iCount = 0 Then
					sRealPath = aNodes(iCount)
					sPath = aNodes(iCount)
				Else
					sRealPath = sRealPath &":"& aNodes(iCount)
					sPath = sPath &":"& aNodes(iCount)
				End If
			End If
		Next
		StrNodeName = sRealPath
	End If
	
	Select Case StrAction
		Case "Select"
					Call Fn_JavaTree_Select("Fn_SE_DataDictionarySearchTreeOperations",JavaWindow("SystemsEngineering").JavaWindow("Search").JavaDialog("DataDictionarySearch"),"HierarchyTree",StrNodeName)
					Fn_SE_DataDictionarySearchTreeOperations = True
		Case "Expand"
		             Call Fn_UI_JavaTree_Expand("Fn_SE_DataDictionarySearchTreeOperations",JavaWindow("SystemsEngineering").JavaWindow("Search").JavaDialog("DataDictionarySearch"),"HierarchyTree",StrNodeName)
					Fn_SE_DataDictionarySearchTreeOperations = True
		Case "PopupMenuSelect"
					aMenuList = split(StrMenu, ":",-1,1)
					intCount = Ubound(aMenuList)
					'Select node
                    Call Fn_JavaTree_Select("Fn_SE_DataDictionarySearchTreeOperations",JavaWindow("SystemsEngineering").JavaWindow("Search").JavaDialog("DataDictionarySearch"),"HierarchyTree",StrNodeName)
					'Open context menu
					Call Fn_UI_JavaTree_OpenContextMenu("Fn_SE_DataDictionarySearchTreeOperations",JavaWindow("SystemsEngineering").JavaWindow("Search").JavaDialog("DataDictionarySearch"),"HierarchyTree",StrNodeName)
					'Select Menu action
					Select Case intCount
						Case "0"
							 JavaWindow("SystemsEngineering").JavaWindow("Search").JavaDialog("DataDictionarySearch").JavaMenu("label:="&aMenuList(0)&"","index:=0").Select
						Case "1"
							JavaWindow("SystemsEngineering").JavaWindow("Search").JavaDialog("DataDictionarySearch").JavaMenu("label:="&aMenuList(0)&"","index:=0").JavaMenu("label:="&aMenuList(1)&"","index:=1").Select
					End Select
					If err.number < 0 Then
						Fn_SE_DataDictionarySearchTreeOperations=False
					Else
						Fn_SE_DataDictionarySearchTreeOperations = True
					End If
	        Case "DoubleClick"
					 JavaWindow("SystemsEngineering").JavaWindow("Search").JavaDialog("DataDictionarySearch").JavaTree("HierarchyTree").Activate StrNodeName
					If err.number < 0 Then
						Fn_SE_DataDictionarySearchTreeOperations = false
					Else
						Fn_SE_DataDictionarySearchTreeOperations = True
					End If
			Case "Exist"
					intNodeCount =  JavaWindow("SystemsEngineering").JavaWindow("Search").JavaDialog("DataDictionarySearch").JavaTree("HierarchyTree").GetROProperty ("items count") 
					For intCount = 0 to intNodeCount - 1
						sTreeItem =  JavaWindow("SystemsEngineering").JavaWindow("Search").JavaDialog("DataDictionarySearch").JavaTree("HierarchyTree").GetItem(intCount)
						If Trim(lcase(sTreeItem)) = Trim(Lcase(StrNodeName)) Then
							Fn_SE_DataDictionarySearchTreeOperations = True
							Exit For
						End If
					Next
					If intCount = intNodeCount Then
						Fn_SE_DataDictionarySearchTreeOperations = False
					End If
	End Select
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name			:	Fn_SE_ShowTraceabilityMatrix
'@@
'@@    Description				:	Function Used to perform operations on  Show Traceability Matrix window in System Engineering
'@@
'@@    Parameters			   	:	1. sAction			: Action [Type of Attribute Group]
'@@								:	2. bAddSource		: boolean flag to click on Set / Add Source button
'@@							 	:	3. bRemoveSource 	: boolean flag to click on Remove Source button
'@@							 	:	4. bAddTarget 		: boolean flag to click on Set / Add Target button
'@@							 	:	5. bRemoveTarget 	: boolean flag to click on Remove Target button
'@@							 	:	6. bSwitchSourceAndTarget : boolean flag to click on Switch Source And Target button
'@@							 	:	7. sSource 			: for future use
'@@							 	:	8. sTarget 			: for future use
'@@
'@@    Return Value		   	   	: 	True Or False
'@@
'@@    Pre-requisite			:	System Engineering perspective should be activated
'@@
'@@    Examples					:	Call Fn_SE_ShowTraceabilityMatrix("Set", "", "", True, "", True, "", "")
'@@
'@@	History				:	
'@@	Developer Name			Date			 Rev. No.	   Changes Done		 Reviewer
'@@------------------------------------------------------------------------------------------------------------------------------
'@@	Koustubh Watwe			2-May-2012		  1.0			created
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public function Fn_SE_ShowTraceabilityMatrix(sAction, bAddSource, bRemoveSource, bAddTarget, bRemoveTarget, bSwitchSourceAndTarget, sSource, sTarget )
	GBL_FAILED_FUNCTION_NAME="Fn_SE_ShowTraceabilityMatrix"
	Dim objDialog
	Fn_SE_ShowTraceabilityMatrix = False
	Set objDialog = JavaWindow("SystemsEngineering").JavaWindow("ShowTraceabilityMatrix")
	If bRemoveSource = "" Then bRemoveSource = False
	If bAddSource = "" Then bAddSource = False
	If bRemoveTarget = "" Then bRemoveTarget = False
	If bAddTarget = "" Then bAddTarget = False
	If bSwitchSourceAndTarget = "" Then bSwitchSourceAndTarget = False

	If Fn_UI_ObjectExist("Fn_SE_ShowTraceabilityMatrix", objDialog) = False Then
		Call Fn_MenuOperation("Select","View:Show Traceability Matrix...")
		Call Fn_ReadyStatusSync(2)
		If Fn_UI_ObjectExist("Fn_SE_ShowTraceabilityMatrix", objDialog) = False  Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SE_ShowTraceabilityMatrix ] Failed to open Show Traceability Matrix window.")
			Exit function
		End If
	End If


	Select Case sAction
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "Set","SetWithOutClose"
			If cBool(bRemoveSource) Then 
				Call Fn_Button_Click("Fn_SE_ShowTraceabilityMatrix", objDialog, "RemoveSource")
			End If
			If sSource <> "" Then
				sSource = Split(sSource,"~")
				Call Fn_TabFolder_Operation("Select",sSource(0),"")
				Call Fn_SE_BOMTableNodeOpeations("Select",sSource(1),"", "", "")
			End If
			If cBool(bAddSource) Then 
				If Fn_UI_ObjectExist("Fn_SE_ShowTraceabilityMatrix",objDialog.JavaButton("AddSourceNew")) Then
					Call Fn_Button_Click("Fn_SE_ShowTraceabilityMatrix", objDialog, "AddSourceNew")
				Else
					Call Fn_Button_Click("Fn_SE_ShowTraceabilityMatrix", objDialog, "AddSource")			
				End If
			End If

			If cBool(bRemoveTarget) Then 
				Call Fn_Button_Click("Fn_SE_ShowTraceabilityMatrix", objDialog, "RemoveTarget")
			End If
			If sTarget <> "" Then 
				sTarget = Split(sTarget,"~")
				Call Fn_TabFolder_Operation("Select",sTarget(0),"")
				Call Fn_SE_BOMTableNodeOpeations("Select",sTarget(1),"", "", "")
			End If
			If cBool(bAddTarget) Then 
				If Fn_UI_ObjectExist("Fn_SE_ShowTraceabilityMatrix",objDialog.JavaButton("AddTargetNew")) Then
					Call Fn_Button_Click("Fn_SE_ShowTraceabilityMatrix", objDialog, "AddTargetNew")
				Else
					Call Fn_Button_Click("Fn_SE_ShowTraceabilityMatrix", objDialog, "AddTarget")				
				End If
			End If

			If cBool(bSwitchSourceAndTarget) Then 
				Call Fn_Button_Click("Fn_SE_ShowTraceabilityMatrix", objDialog, "SwitchSourceAndTarget")
			End If
			
			If sAction <> "SetWithOutClose" Then
				Call Fn_Button_Click("Fn_SE_ShowTraceabilityMatrix", objDialog, "OK")
			End If
			Fn_SE_ShowTraceabilityMatrix = True
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "DialogExist"
			If Fn_UI_ObjectExist("Fn_SE_ShowTraceabilityMatrix", objDialog) = True Then
				Fn_SE_ShowTraceabilityMatrix =True
			Else
				Fn_SE_ShowTraceabilityMatrix = False
			End If
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "VerifySourceTarget","VerifySourceTargetWithoutClose"  '[TC1015-2015072100-20_08_2015-VivekA-NewDevlopment] - Added new cases to verfy Source and Target values
			If sSource <> "" Then
				sGetSource = Fn_SISW_UI_JavaEdit_Operations("Fn_SE_ShowTraceabilityMatrix", "GetText", objDialog, "Source", "")
				If Trim(sGetSource) = Trim(sSource) Then
					Fn_SE_ShowTraceabilityMatrix = True
				Else
					Fn_SE_ShowTraceabilityMatrix = False
				End If
			End If
			If sTarget <> "" Then 
				sGetTarget = Fn_SISW_UI_JavaEdit_Operations("Fn_SE_ShowTraceabilityMatrix", "GetText", objDialog, "Target", "")
				If Trim(sGetTarget) = Trim(sTarget) Then
					Fn_SE_ShowTraceabilityMatrix = True
				Else
					Fn_SE_ShowTraceabilityMatrix = False
				End If
			End If
			If sAction <> "VerifySourceTargetWithoutClose"  Then
				Call Fn_Button_Click("Fn_SE_ShowTraceabilityMatrix", objDialog, "OK")
			End If
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SE_ShowTraceabilityMatrix ] Invalid case [ " & sAction & " ].")
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
	End Select
	If  Fn_SE_ShowTraceabilityMatrix <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [ Fn_SE_ShowTraceabilityMatrix ] executed successfuly with case [ " & sAction & " ].")
	End If
	Set objDialog = Nothing
End Function

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name			:	Fn_SE_TraceabilityMatrixPanelOperations
'@@
'@@    Description				:	Function Used to perform operations on  Show Traceability Matrix window in System Engineering
'@@
'@@    Parameters			   	:	1. sAction		: Action [Type of Attribute Group]
'@@								:	2. sRow			: Row text
'@@							 	:	3. sCol 		: for future use
'@@							 	:	4. sValue 		: for future use
'@@							 	:	5. sPopupMenu 	: Popup Menu select
'@@
'@@    Return Value		   	   	: 	True Or False
'@@
'@@    Pre-requisite			:	SYstem Engineering perspective should be activated and Traceability Matrix Tab is Selected.						
'@@
'@@    Examples					:	Call Fn_SE_TraceabilityMatrixPanelOperations("Select", "001860/A;1-f2", "", "", "")
'@@    Examples					:	Call Fn_SE_TraceabilityMatrixPanelOperations("PopupMenuSelect", "001860/A;1-f2", "", "", "Refresh Report")
'@@    Examples					:	Call Fn_SE_TraceabilityMatrixPanelOperations("GetColumnIndex", "", "Total", "", "")
'@@    Examples					:	Call Fn_SE_TraceabilityMatrixPanelOperations("RowVerify", "Total", "Total~001765/A;1-rspec (View)~REQ-000001/A;1-req", "~0~0", "")
'@@
'@@	History				:	
'@@	Developer Name			Date			 			Rev. No.	   Changes Done		
'@@----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@	Koustubh Watwe			2-May-2012		  		1.0					created
'@@----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@	Pranav Ingle				2-May-2012		  			1.0				Added Cases "ColPopupMenuSelect", "GetRowIndex", "VerifyTraceLink"
'																											"CellSelect", "CellPopupMenuSelect", "CellPopupMenuExist", "CreateTraceLink", "DeleteTraceLink"
'																											"CellMultiSelect", "CellMultiPopupSelect", "CreateMultiTraceLink", "CellMultiSelectPopupMenuExist"
'																											"TabVerify"
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'  Note  :   For "CreateTraceLink"  & "CreateMultiTraceLink" Cases sPopupMenu  paramater will work as Trace Link Type  

Public Function Fn_SE_TraceabilityMatrixPanelOperations(sAction, sRow, sCol, sValue, sPopupMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_SE_TraceabilityMatrixPanelOperations"
	Dim objNatTable, iRowCnt, iColCnt, iCnt, sColindex, bResult
	Dim bFound, aBounds, aMenu, aValues, iCount, iColNumber
	Dim objTable, objChild
	Dim arrRow, arrCol, btn

	Set objNatTable = JavaWindow("SystemsEngineering").JavaObject("TraceabilityMatrixNatTable")
	Fn_SE_TraceabilityMatrixPanelOperations = False
	' add call to select Traceability Matrix Tab
	If Fn_UI_ObjectExist("Fn_SE_TraceabilityMatrixPanelOperations",objNatTable ) = False  Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SE_TraceabilityMatrixPanelOperations ] Nat Table object of [ Traceability Matrix ] is not visible.")
		Exit function
	End If

	Select Case sAction
		Case "Select"
			sCol = 0
			iRowCnt = cInt(objNatTable.Object.getRowCount())
			For iCnt =1 to iRowCnt -1
				If sRow = objNatTable.Object.getCellByPosition(sCol,iCnt).getDataValue().toString() then
					Exit for
				End If
			Next
			If iCnt < iRowCnt Then
				sRectangle = cStr(objNatTable.Object.getCellByPosition(sCol,iCnt).getBounds().toString())
				sRectangle =  Right(sRectangle, (Len(sRectangle) - Instr(1, sRectangle, "{", 1)))	
				sRectangle = replace(sRectangle,"}","")
				sRectangle = replace(sRectangle," ","")
				aBounds = split(sRectangle,",")
				objNatTable.Click cInt(aBounds(0)) + (cInt(aBounds(2))/2),  cInt(aBounds(1)) + (cInt(aBounds(3))/2)  , "LEFT"
				Fn_SE_TraceabilityMatrixPanelOperations = True
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SE_TraceabilityMatrixPanelOperations ] Row not found [ " & sRow & " ].")
			End If
	' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "PopupMenuSelect"
			sCol = 0
			iRowCnt = cInt(objNatTable.Object.getRowCount())
			For iCnt = 1 to iRowCnt -1
				If sRow = objNatTable.Object.getCellByPosition(sCol,iCnt).getDataValue().toString() then
					Exit for
				End If
			Next
			If iCnt < iRowCnt Then
				sRectangle = cStr(objNatTable.Object.getCellByPosition(sCol,iCnt).getBounds().toString())
				sRectangle =  Right(sRectangle, (Len(sRectangle) - Instr(1, sRectangle, "{", 1)))	
				sRectangle = replace(sRectangle,"}","")
				sRectangle = replace(sRectangle," ","")
				aBounds = split(sRectangle,",")
				objNatTable.Click cInt(aBounds(0)) + (cInt(aBounds(2))/2),  cInt(aBounds(1)) + (cInt(aBounds(3))/2)  , "RIGHT"
				wait 2
				Fn_SE_TraceabilityMatrixPanelOperations = Fn_UI_JavaMenu_Select("Fn_SE_TraceabilityMatrixPanelOperations",JavaWindow("SystemsEngineering"),sPopupMenu)
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SE_TraceabilityMatrixPanelOperations ] Row not found [ " & sRow & " ].")
			End If
			Call Fn_SyncTCObjects()
	' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "ColPopupMenuSelect", "ColPopupMenuSelectExt"
			sRow = 0
			iColCnt = cInt(objNatTable.Object.getColumnCount())
			For iCnt = 1 to iColCnt -1
				If sCol = objNatTable.Object.getCellByPosition(iCnt,sRow).getDataValue().toString() then
					Exit for
				End If
			Next
			If iCnt < iColCnt Then
				sRectangle = cStr(objNatTable.Object.getCellByPosition(iCnt,sRow).getBounds().toString())
				sRectangle =  Right(sRectangle, (Len(sRectangle) - Instr(1, sRectangle, "{", 1)))	
				sRectangle = replace(sRectangle,"}","")
				sRectangle = replace(sRectangle," ","")
				aBounds = split(sRectangle,",")
				objNatTable.Click cInt(aBounds(0)) + (cInt(aBounds(2))/2),  cInt(aBounds(1)) + (cInt(aBounds(3))/2)  , "RIGHT"
				wait 2
				Fn_SE_TraceabilityMatrixPanelOperations = Fn_UI_JavaMenu_Select("Fn_SE_TraceabilityMatrixPanelOperations",JavaWindow("SystemsEngineering"),sPopupMenu)
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SE_TraceabilityMatrixPanelOperations ] Column not found [ " & sCol & " ].")
			End If
			
			If sAction ="ColPopupMenuSelect" Then
				Call Fn_SyncTCObjects()
			End If
	' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "RowVerify"
			aValues	= split(sValue,"~")
			aCols = split(sCol,"~")
			
			iRowCnt = cInt(objNatTable.Object.getRowCount())
			For iCnt = 1 to iRowCnt -1
				If sRow = objNatTable.Object.getCellByPosition(0,iCnt).getDataValue().toString() then
					Exit for
				End If
			Next
			If iCnt < iRowCnt Then
				For iCount = 0 to uBound(aValues)
					iColNumber = Fn_SE_TraceabilityMatrixPanelOperations("GetColumnIndex", "", aCols(iCount), "", "")
					If iColNumber <> -1 Then 
						Fn_SE_TraceabilityMatrixPanelOperations = True
						If trim(aValues(iCount)) <> trim(objNatTable.Object.getCellByPosition(iColNumber, iCnt).getDataValue().toString()) then
							Fn_SE_TraceabilityMatrixPanelOperations = False
							Exit for
						End If
					Else
						Fn_SE_TraceabilityMatrixPanelOperations = False
						exit for
					End If
					
				Next
			End IF
	' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "GetColumnIndex"
			Fn_SE_TraceabilityMatrixPanelOperations = -1
			iCount = cInt(objNatTable.Object.getColumnCount)
			For iCnt = 0 to iCount - 1
				If trim(sCol) = trim(objNatTable.Object.getCellByPosition(iCnt,0).getDataValue().toString()) Then
					Fn_SE_TraceabilityMatrixPanelOperations = iCnt
					exit for
				End If

			Next
	 ' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "GetRowIndex"
			sColindex = 0
			Fn_SE_TraceabilityMatrixPanelOperations = -1
			iRowCnt = cInt(objNatTable.Object.getRowCount())
			For iCnt = 1 to iRowCnt -1
				If Trim(sRow) = objNatTable.Object.getCellByPosition(sColindex,iCnt).getDataValue().toString() then
						Fn_SE_TraceabilityMatrixPanelOperations = iCnt
						Exit For
				End If
			Next
	'----------------------------------------------------------------------------------------------------------------------------------------------
		Case "VerifyTraceLink"		
			iRowCnt=0
			If sPopupMenu <> "" Then
					sPopupMenu	= split(sPopupMenu,"~")
					bFound = False
					Call Fn_SISW_UI_JavaButton_Operations("Fn_SE_TraceabilityMatrixPanelOperations", "Click", JavaWindow("SystemsEngineering"),"PanelTraceLinkType")
					Set objTable=Description.Create()
					objTable("Class Name").value="JavaTable"
					objTable("path").value="Table;Shell;Shell;"
					Set objChild=JavaWindow("SystemsEngineering").ChildObjects(objTable)

					For iCount=0 to Ubound(sPopupMenu)
						For iCnt=0 to (objChild(0).GetROProperty("rows")-1)
							If trim(sPopupMenu(iCount))=trim(objChild(0).GetCellData(iCnt,0)) Then		
								iRowCnt=iRowCnt+1
							End If
						Next
					Next
					Set objTable=Nothing
					Set objChild=Nothing
					If iRowCnt = (Ubound(sPopupMenu)+1) Then
						Fn_SE_TraceabilityMatrixPanelOperations = True  
					Else
						 Exit Function
					End If
			End If
  	' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "VerifyBtnEnable"	
			JavaWindow("SystemsEngineering").JavaButton("PanelTraceLinkDelete").SetTOProperty"label",sValue
			set btn =JavaWindow("SystemsEngineering").JavaButton("PanelTraceLinkDelete")
			If True =btn.CheckProperty("enabled",1) Then
				Fn_SE_TraceabilityMatrixPanelOperations=True
		Else
				Fn_SE_TraceabilityMatrixPanelOperations=False
			Exit Function
			End If
	' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "CellSelect", "CellPopupMenuSelect", "CellPopupMenuExist", "CreateTraceLink", "DeleteTraceLink"
			iRowCnt = Fn_SE_TraceabilityMatrixPanelOperations("GetRowIndex",sRow, "","","")
			If CInt(iRowCnt) = -1 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SE_TraceabilityMatrixPanelOperations ] Row Not Found[ " & sRow & " ].")
				Exit Function
			End If
		
			iColCnt = Fn_SE_TraceabilityMatrixPanelOperations("GetColumnIndex","",sCol ,"","")
			If CInt(iColCnt) = -1 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SE_TraceabilityMatrixPanelOperations ] Column Not Found[ " & sCol & " ].")
				Exit Function
			End If

			sRectangle = cStr(objNatTable.Object.getCellByPosition(iColCnt,iRowCnt).getBounds().toString())
			sRectangle =  Right(sRectangle, (Len(sRectangle) - Instr(1, sRectangle, "{", 1)))	
			sRectangle = replace(sRectangle,"}","")
			sRectangle = replace(sRectangle," ","")
			aBounds = split(sRectangle,",")

            If sAction = "CellSelect" Then
				objNatTable.Click cInt(aBounds(0)) + (cInt(aBounds(2))/2),  cInt(aBounds(1)) + (cInt(aBounds(3))/2)  , "LEFT"
				Fn_SE_TraceabilityMatrixPanelOperations = True

			ElseIf sAction = "CreateTraceLink" Then
                objNatTable.Click cInt(aBounds(0)) + (cInt(aBounds(2))/2),  cInt(aBounds(1)) + (cInt(aBounds(3))/2)  , "LEFT"
				wait 1
				If sPopupMenu <> "" Then
						bFound = False
						Call Fn_SISW_UI_JavaButton_Operations("Fn_SE_TraceabilityMatrixPanelOperations", "Click", JavaWindow("SystemsEngineering"),"PanelTraceLinkType")
						Set objTable=Description.Create()
						objTable("Class Name").value="JavaTable"
						objTable("path").value="Table;Shell;Shell;"
                        Set objChild=JavaWindow("SystemsEngineering").JavaWindow("Shell").ChildObjects(objTable)
						'Set objChild=JavaWindow("SystemsEngineering").ChildObjects(objTable)
						
						For iCnt=0 to objChild(0).GetROProperty("rows")
							If trim(sPopupMenu)=trim(objChild(0).GetCellData(iCnt,0)) Then
								'objChild(0).ClickCell iCnt,0
								objChild(0).SelectCell iCnt,0
								bFound = True
								Exit for
							End If
						Next
						Set objTable=Nothing
						Set objChild=Nothing
						If bFound = False Then
							   Exit Function
						End If
				End If
				Fn_SE_TraceabilityMatrixPanelOperations = Fn_SISW_UI_JavaButton_Operations("Fn_SE_TraceabilityMatrixPanelOperations", "Click", JavaWindow("SystemsEngineering"),"PanelTraceLinkCreate")
			ElseIf sAction = "DeleteTraceLink" Then
				objNatTable.Click cInt(aBounds(0)) + (cInt(aBounds(2))/2),  cInt(aBounds(1)) + (cInt(aBounds(3))/2)  , "LEFT"
				wait 1
				Fn_SE_TraceabilityMatrixPanelOperations = Fn_SISW_UI_JavaButton_Operations("Fn_SE_TraceabilityMatrixPanelOperations", "Click", JavaWindow("SystemsEngineering"),"PanelTraceLinkDelete")
			ElseIf sAction = "CellPopupMenuSelect" Then
				objNatTable.Click cInt(aBounds(0)-10) + (cInt(aBounds(2))/2),  cInt(aBounds(1)) + (cInt(aBounds(3))/2)  , "LEFT"
                wait 1
				objNatTable.Click cInt(aBounds(0)-20) + (cInt(aBounds(2))/2),  cInt(aBounds(1)) + (cInt(aBounds(3))/2)  , "RIGHT"
				wait 2
				Fn_SE_TraceabilityMatrixPanelOperations = Fn_UI_JavaMenu_Select("Fn_SE_TraceabilityMatrixPanelOperations",JavaWindow("SystemsEngineering"),sPopupMenu)


            Elseif sAction = "CellPopupMenuExist" then
				objNatTable.Click cInt(aBounds(0)-10) + (cInt(aBounds(2))/2),  cInt(aBounds(1)) + (cInt(aBounds(3))/2)  , "LEFT"
                wait 1
				objNatTable.Click cInt(aBounds(0)-20) + (cInt(aBounds(2))/2),  cInt(aBounds(1)) + (cInt(aBounds(3))/2)  , "RIGHT"
				wait 2
				aMenu=Split(sPopupMenu,":")
				Set objPopupMenu=JavaWindow("SystemsEngineering").JavaMenu("label:="&aMenu(0))
				For iCount=1 to ubound(aMenu)
					Set objPopupMenu=objPopupMenu.JavaMenu("label:="&aMenu(iCount))
				Next
				Fn_SE_TraceabilityMatrixPanelOperations=objPopupMenu.Exist(5)
				
				wait 1
				Set WshShell = CreateObject("WScript.Shell")
				wait 1
				WshShell.SendKeys "{ESC}"
				Set WshShell =Nothing
			End If

	Case "CellMultiSelect", "CellMultiPopupSelect", "CreateMultiTraceLink", "CellMultiSelectPopupMenuExist"
			arrRow = Split(sRow, "~")
			arrCol = Split(sCol, "~")
			For iCount = 0 To UBound(arrRow)
				iRowCnt = Fn_SE_TraceabilityMatrixPanelOperations("GetRowIndex",arrRow(iCount), "","","")
				If CInt(iRowCnt) = -1 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SE_TraceabilityMatrixPanelOperations ] Row Not Found[ " & arrRow(iCount) & " ].")
					Exit Function
				End If

				iColCnt = Fn_SE_TraceabilityMatrixPanelOperations("GetColumnIndex","", arrCol(iCount),"","")
				If CInt(iColCnt) = -1 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SE_TraceabilityMatrixPanelOperations ] Column Not Found[ " & arrCol(iCount) & " ].")
					Exit Function
				End If
                
				sRectangle = cStr(objNatTable.Object.getCellByPosition(iColCnt,iRowCnt).getBounds().toString())
				sRectangle =  Right(sRectangle, (Len(sRectangle) - Instr(1, sRectangle, "{", 1)))	
				sRectangle = replace(sRectangle,"}","")
				sRectangle = replace(sRectangle," ","")
				aBounds = split(sRectangle,",")
	
				Set myDeviceReplay = CreateObject("Mercury.DeviceReplay")
				If iCount = 1 Then
					myDeviceReplay.KeyDown 29
				End If
				objNatTable.Click cInt(aBounds(0)) + (cInt(aBounds(2))/2),  cInt(aBounds(1)) + (cInt(aBounds(3))/2)  , "LEFT"
			Next

			Fn_SE_TraceabilityMatrixPanelOperations = True
            myDeviceReplay.KeyUp 29

			If sAction = "CreateMultiTraceLink" Then
				wait 1
				If sPopupMenu <> "" Then
						bFound = False
						Call Fn_SISW_UI_JavaButton_Operations("Fn_SE_TraceabilityMatrixPanelOperations", "Click", JavaWindow("SystemsEngineering"),"PanelTraceLinkType")
						Set objTable=Description.Create()
						objTable("Class Name").value="JavaTable"
						objTable("path").value="Table;Shell;Shell;"
'                       Set objChild=JavaWindow("SystemsEngineering").JavaWindow("Shell").ChildObjects(objTable)
						Set objChild=JavaWindow("SystemsEngineering").ChildObjects(objTable)
						
						For iCnt=0 to objChild(0).GetROProperty("rows")
							If trim(sPopupMenu)=trim(objChild(0).GetCellData(iCnt,0)) Then
								objChild(0).SelectCell iCnt,0
								bFound = True
								Exit for
							End If
						Next
						Set objTable=Nothing
						Set objChild=Nothing
						If bFound = False Then
							   Exit Function
						End If
				End If
				Fn_SE_TraceabilityMatrixPanelOperations = Fn_SISW_UI_JavaButton_Operations("Fn_SE_TraceabilityMatrixPanelOperations", "Click", JavaWindow("SystemsEngineering"),"PanelTraceLinkCreate")

			ElseIf sAction = "CellMultiPopupSelect" Then
				objNatTable.Click cInt(aBounds(0)) + (cInt(aBounds(2))/2),  cInt(aBounds(1)) + (cInt(aBounds(3))/2)  , "RIGHT"
				wait 2
				Fn_SE_TraceabilityMatrixPanelOperations = Fn_UI_JavaMenu_Select("Fn_SE_TraceabilityMatrixPanelOperations",JavaWindow("SystemsEngineering"),sPopupMenu)
			Elseif sAction = "MultiSelectPopupMenuExist" then
				objNatTable.Click cInt(aBounds(0)) + (cInt(aBounds(2))/2),  cInt(aBounds(1)) + (cInt(aBounds(3))/2)  , "RIGHT"
				wait 2

				aMenu=Split(sPopupMenu,":")
				Set objPopupMenu=JavaWindow("SystemsEngineering").JavaMenu("label:="&aMenu(0))
				For iCount=1 to ubound(aMenu)
					Set objPopupMenu=objPopupMenu.JavaMenu("label:="&aMenu(iCount))
				Next
				Fn_SE_TraceabilityMatrixPanelOperations=objPopupMenu.Exist(5)
				
				wait 1
				Set WshShell = CreateObject("WScript.Shell")
				wait 1
				WshShell.SendKeys "{ESC}"
				Set WshShell =Nothing
			End If
	' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "TabVerify"
				If Fn_UI_ObjectExist("Fn_SE_TraceabilityMatrixPanelOperations",objNatTable ) = False  Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SE_TraceabilityMatrixPanelOperations ] Nat Table object of [ Traceability Matrix ] is not visible.")
					Exit function
				Else
					Fn_SE_TraceabilityMatrixPanelOperations = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SE_TraceabilityMatrixPanelOperations ] Nat Table object of [ Traceability Matrix ] is visible.")
				End If
				
		Case "IsSelected"
			arrRow = Split(sRow, "~")
			arrCol = Split(sCol, "~")
			bResult=False
			For iCount = 0 To UBound(arrRow)
				iRowCnt = Fn_SE_TraceabilityMatrixPanelOperations("GetRowIndex",arrRow(iCount), "","","")
				If CInt(iRowCnt) = -1 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SE_TraceabilityMatrixPanelOperations ] Row Not Found[ " & arrRow(iCount) & " ].")
					Exit Function
				End If

				iColCnt = Fn_SE_TraceabilityMatrixPanelOperations("GetColumnIndex","", arrCol(iCount),"","")
				If CInt(iColCnt) = -1 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SE_TraceabilityMatrixPanelOperations ] Column Not Found[ " & arrCol(iCount) & " ].")
					Exit Function
				End If
                
				sDisplayMode = objNatTable.Object.getCellByPosition(iColCnt,iRowCnt).getDisplayMode()
				If sDisplayMode ="SELECT" Then
					bResult=True
				Else	
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SE_TraceabilityMatrixPanelOperations ] Nat Table Cell is not selected.")
					Fn_SE_TraceabilityMatrixPanelOperations=False
					Exit Function
				End If
			Next	
			
			If bResult=True Then
				Fn_SE_TraceabilityMatrixPanelOperations=bResult
			End If
			
		Case "VerifyRowStructure"
					'no of rows in table
					flag=False
					MatchCounter=-1
					iRowCnt=objNatTable.Object.getrowcount
					tempRowCount=Split(sRow,"~")
							For iCnt = 0 To iRowCnt-1
								If Trim(tempRowCount(0))=Trim(objNatTable.Object.getCellByPosition(0,iCnt).getDataValue().toString()) Then
									flag=True
									iRandom=icnt
								End If
								If flag=True Then
									If Trim(tempRowCount(iCnt-iRandom))=Trim(objNatTable.Object.getCellByPosition(0,iCnt).getDataValue().toString()) Then
										MatchCounter=MatchCounter + 1
									End If
								End If
							Next 
							If MatchCounter = UBound(tempRowCount) Then
								Fn_SE_TraceabilityMatrixPanelOperations = True
							End If
					
			Case "VerifyColStructure"
					'no of columns in table
					flag=False
					MatchCounter=-1					
					iColCnt=objNatTable.Object.getColumncount
					tempColCount=Split(sCol,"~")
							For iCnt = 0 To iColCnt-1
								If Trim(tempColCount(0))=Trim(objNatTable.Object.getCellByPosition(iCnt,0).getDataValue().toString()) Then
									flag=True
									iRandom=icnt
								End If
								If flag=True Then
									If Trim(tempColCount(iCnt-iRandom))=Trim(objNatTable.Object.getCellByPosition(iCnt,0).getDataValue().toString()) Then
										MatchCounter=MatchCounter + 1
									End If
								End If
							Next 
							If MatchCounter = UBound(tempColCount) Then
								Fn_SE_TraceabilityMatrixPanelOperations = True
							End If	
	'[TC11.4_NewDevelopment_PoonamC_19July2017 : Added Case to select Trace Link type ]
	Case "SelectTraceLinkTypes"
				JavaWindow("SystemsEngineering").JavaEdit("Type").SetTOProperty "attached text", "Trace Link Types:"
				If Fn_UI_ObjectExist("Fn_SE_TraceabilityMatrixPanelOperations",JavaWindow("SystemsEngineering").JavaEdit("Type") ) = False  Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SE_TraceabilityMatrixPanelOperations ] Trace Link Types object of [ Traceability Matrix ] is not visible.")
					Exit function
				Else
					JavaWindow("SystemsEngineering").JavaEdit("Type").Click 1,1
					Call Fn_SISW_UI_JavaEdit_Operations("Fn_SE_TraceabilityMatrixPanelOperations","Type",JavaWindow("SystemsEngineering"),"Type",sValue)
					call Fn_KeyBoardOperation("SendKeys", "{TAB}")
					Fn_SE_TraceabilityMatrixPanelOperations = True	
				End If
	' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SE_TraceabilityMatrixPanelOperations ] Invalid case [ " & sAction & " ].")
	' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
	End Select
	If  Fn_SE_TraceabilityMatrixPanelOperations <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [ Fn_SE_TraceabilityMatrixPanelOperations ] executed successfuly with case [ " & sAction & " ].")
	End If
	Set objNatTable = Nothing
End Function

'''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''''/$$$$
''''/$$$$   FUNCTION NAME   :  Fn_SE_NotesOperations(sAction,strNodeName,strColName, NoteName,AttachedLov,sInfo1)
''''/$$$$
''''/$$$$   DESCRIPTION        :  This function will  perform Operations on Note For Specified Requirement
''''/$$$$ 
''''/$$$$   PRE-REQUISITES        :  System Engineering Perspective Should Be Open and 
''''/$$$$
''''/$$$$  PARAMETERS   : 		StrAction : Action to be performed
''''/$$$$										StrNode : Node name in BOMTable
''''/$$$$										StrColName : ColumnName  to double click
''''/$$$$										StrNoteName :Note type  in Create list [ e.g "Ignore Partial match ]
'''/$$$$										StrValue :  Value in List of Values list  i.e true/false
'''/$$$$									
''''/$$$$										StrButton: Button name
''''/$$$$									
''''/$$$$	
''''/$$$$		Return Value : 				True or False
''''/$$$$
''''/$$$$    Function Calls       :   Fn_WriteLogFile(), Fn_MenuOperation(), Fn_ReadyStatusSync
''''/$$$$										
''''/$$$$
''''/$$$$		HISTORY           :   		AUTHOR                 DATE        VERSION
''''/$$$$
''''/$$$$    CREATED BY     :   Avinash          27/004/2012         1.0
''''/$$$$
''''/$$$$    REVIWED BY     :  Sandeep N.	27/004/2012           1.0
''''/$$$$
''''/$$$$		How To Use :    bReturn=Fn_SE_NotesOperations("SetNote","001624/A;1-reqspec1 (View):REQ-001020/A;1-req1","All Notes","Ignore Partial Match","true","OK")
''''/$$$$							
''''/$$$$	
''''/$$$$			Developer Name				Date						 Rev. No.	   	Changes Done		 														Reviewer
''''/$$$$		-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
''''/$$$$			Pranav Ingle					12-Feb-2013		  			1.1				Modified code for AttachedLOV property
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Public Function Fn_SE_NotesOperations(StrAction,StrNode,StrColName,StrNoteName,StrAttachedLov,StrButton)
	GBL_FAILED_FUNCTION_NAME="Fn_SE_NotesOperations"
	Dim ObjNoteDialog,bFlag
	Fn_SE_NotesOperations=False 
	'creating object of [ Notes ] dialog
	Set ObjNoteDialog=JavaWindow("SystemsEngineering").JavaWindow("SEWindow").JavaDialog("Notes")
	'Chcking existane of  [  Notes ] dialog
	If  Not ObjNoteDialog.Exist(5)  Then
			'Double click on  All Notes cell to open NoteFor Requirement dialog
			bFlag=Fn_SE_BOMTableNodeOpeations("DoubleClickCell",StrNode,StrColName, "", "")
			Call Fn_ReadyStatusSync(2)
			If bFlag=false or Not ObjNoteDialog.Exist(6) Then
				Set ObjNoteDialog=nothing
				Exit function
			End If
			'Double click on  cell to open NoteFor Req dialog
	End If
	   Select Case StrAction
			Case "SetNote"
			  'If Remove button is enabled then click on remove button
				If Fn_UI_Object_GetROProperty("Fn_SE_NotesOperations",ObjNoteDialog.JavaButton("Remove"), "enabled")="1"  Then
					call Fn_Button_Click("Fn_SE_NotesOperations",ObjNoteDialog,"Remove")
				End If
				'Checking Note type exist in Create list
				bFlag=Fn_UI_ListItemExist("Fn_SE_NotesOperations",ObjNoteDialog,"CreateList",StrNoteName)
				If bFlag=true Then
					bFlag= Fn_List_Select("Fn_SE_NotesOperations",ObjNoteDialog,"CreateList",StrNoteName)
				else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Note type [ "+StrNoteName+" ] is not exist in Create list")
					Set ObjNoteDialog=nothing
					Exit function
				End If
				'Setting Attached LOV option on to open [ List of values ] list
				If StrAttachedLov <> "" Then
					'-------------------------------------------------------------------------------------------------------------------------------------------    By: Pranav Ingle 12-Feb-2013
'					Call Fn_CheckBox_Set("Fn_SE_NotesOperations" ,ObjNoteDialog,"AttachedLOV", "on")
					Call Fn_Edit_Box("Fn_SE_NotesOperations", ObjNoteDialog, "AttachedLOV",StrAttachedLov+vblf )
					
					'Checking values exist in Liat of Values list
'					bFlag=Fn_UI_ListItemExist("Fn_SE_NotesOperations",ObjNoteDialog,"ListOfValues", StrAttachedLov)
'					bFlag=JavaWindow("SystemsEngineering").JavaWindow("SEWindow").JavaDialog("Notes").JavaEdit("ExistingNotes").CheckProperty("value","true")
'					If bFlag=true Then
'						ObjNoteDialog.JavaList("ListOfValues").Activate(StrAttachedLov)
'						wait 1
'					else
'						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: value [ "+StrAttachedLov+" ] is not exist in List of Values")
'						Set ObjNoteDialog=nothing
'						Exit function
'					End If	
				End If
				'--------------------------------------------------------------------------------------------------------------------------------------------   

				'clik on button
				 Call  Fn_Button_Click("Fn_SE_NotesOperations",ObjNoteDialog,StrButton)
				 If Err.number < 0 Then
                	Fn_SE_NotesOperations=false
				Else
					Fn_SE_NotesOperations=true
				End If  
	  End Select
	'releasing object of [ Notes ] dialog
	 Set ObjNoteDialog=Nothing
End Function

'************************************		Function to Invokes the Create New Rev Rule dialog and fill in basic details as Name, Desciption etc	****************************************************
'Function Name			:		       Fn_SE_RevRuleCreateWithExitingOne

'Description			    :		 	 Invokes the Create New Rev Rule dialog and fill in basic details as Name, Desciption etc.

'Parameters			   :	 		1. sAction --> Action to perform
'								2. sRevName --> Variable to insert name for the revision rule
'								3. sRevDesc --> Variable to provide revision rule description
'								4. blNestEffect --> ON/OFF Toggle for setting Nested Effectivity on "Create New Revision Rule window/Dialog Box - 0 to Uncheck and 1 to Check.

'Return Value		   	   :            		True / False

'Pre-requisite			    :		 	 PE window should be displayed / open

'Examples				    :			 Call Fn_PE_RevRuleModify("ModifyCurrent","","My second test", "ON")
'								Call Fn_PE_RevRuleModify("Modify","Test1","My second test", "ON")

'History				        :		
'							Developer Name			Date			Rev. No.			Changes Done			Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'							Avinash Jagdale	        04/04/2012	           1.0				Created				
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SE_RevRuleCreateWithExitingOne(sAction, sExitingRev,sRevName, sRevDesc, blNestEffect)
	GBL_FAILED_FUNCTION_NAME="Fn_SE_RevRuleCreateWithExitingOne"
	Dim objCreateRevisionRuleDialog, objRevisionRuleDialog, bReturn
	Fn_SE_RevRuleCreateWithExitingOne = False
	Set objCreateRevisionRuleDialog = JavaWindow("StructureManager").JavaWindow("PSEWindow").JavaDialog("Create New Revision Rule")

	Select Case sAction
        '  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Create"
				'Verify that Revision rules dialog box is displayed	
				Set objRevisionRuleDialog =  JavaWindow("StructureManager").JavaWindow("PSEWindow").JavaDialog("Revision Rules")
				If NOT(objRevisionRuleDialog.Exist(5))  Then
					'Operate Tools>>Revision Rule>>Create/Edit... menu to invoke required dialog
					Call Fn_MenuOperation("Select","Tools:Revision Rule:Create/Edit...")
				End If
				' selecting rev ruld from the list
				Call Fn_List_Select("Fn_SE_RevRuleCreateWithExitingOne",objRevisionRuleDialog,"RevisionRules", sExitingRev)
				' clicking on modify
				bReturn = Fn_Button_Click("Fn_SE_RevRuleCreateWithExitingOne",objRevisionRuleDialog,"Create")
				If bReturn  = True Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SE_RevRuleCreateWithExitingOne ]  Successfully Clicked on Create Revision Rule button. ")
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SE_RevRuleCreateWithExitingOne ] Failed to click on Create  Revision Rule button.")
					Fn_SE_RevRuleCreateWithExitingOne = FALSE
					Exit function
				End If
'  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case Else
			Fn_SE_RevRuleCreateWithExitingOne = FALSE
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SE_RevRuleCreateWithExitingOne ] Invalid Action [ " & sAction & " ]")
			Exit function
	End Select
			
	'Verify that Modify Revision Rule dialog is displayed
	If objCreateRevisionRuleDialog.Exist(5) Then
		' setting name
		If sRevName <> "" Then
			Call Fn_Edit_Box("Fn_SE_RevRuleCreateWithExitingOne", objCreateRevisionRuleDialog, "Name", sRevName)
		End If
		' setting description
		If sRevDesc <>"" Then
			Call Fn_Edit_Box("Fn_SE_RevRuleCreateWithExitingOne", objCreateRevisionRuleDialog, "Description",sRevDesc)
		End If
		'To check the Nested Effectivity ON/OFF
		If blNestEffect<>"" AND (UCase(blNestEffect) = "ON" OR UCase(blNestEffect) = "TRUE" )Then
			objCreateRevisionRuleDialog.JavaCheckBox("Nested Effectivity").Set "ON"
		Else 
			objCreateRevisionRuleDialog.JavaCheckBox("Nested Effectivity").Set "OFF"
		End If

     		Fn_SE_RevRuleCreateWithExitingOne = TRUE
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SE_RevRuleCreateWithExitingOne ] executed successfully with case [ " & sAction & " ]")
	Else
		Fn_SE_RevRuleCreateWithExitingOne = FALSE
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SE_RevRuleCreateWithExitingOne ] execution failed with case [ " & sAction & " ]")
	End If

	Set objCreateRevisionRuleDialog = nothing
	Set objRevisionRuleDialog = nothing
End Function

'************************************		Function to Invokes the Create New Rev Rule dialog and fill in basic details as Name, Desciption etc	****************************************************
'Function Name			:		       Fn_SE_FilterSetting

'Description			    :		 	Apply the filter conditions 

'Parameters			   :	 	1.sAction: Action type	
'								2. sColumn --> Column name to filtered	
'								3. sFilterType --> filter type ie. Operator and Attribute
'								4. sValue --> value to be set
'								5. sReserve --> Reserve for future use

'Return Value		   	   :            		True / False

'Pre-requisite			    :		 	 PE window should be displayed / open

'Examples				    :			 Call Fn_SE_FilterSetting("Item Type","Attribute","M3SEItem_Auto")

'History				        :		
'							Developer Name			Date			Rev. No.			Changes Done			Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'							Nilesh Gadekar    		16/05/2012	           1.0				Created				
'							Nilesh Gadekar			11/6/2012				2.0			 Implemented DP for filter select		
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public  Function Fn_SE_FilterSetting(sAction,sColumn,sFilterType,sValue)
	GBL_FAILED_FUNCTION_NAME="Fn_SE_FilterSetting"
	Dim iCols,iCounter,sColumnName,iColumn_no,WshShell,objTable,sGetValue,sNode,aNode,ObjEdit,FiletrEdit
	Set objTable=JavaWindow("SystemsEngineering").JavaWindow("WEmbeddedFrame").JavaTable("CMEBOMTreeTable")
	sNode=Fn_SE_BOMTableNodeOpeations("GetCellData",0,0, "", "")
	aNode=Split(sNode,"(",-1,1)
'	If Instr(aNode(1),"-")>0 Then
'			If   Instr(aNode(1),"(")>0 Then
'				aNode=Split(
'			End If
'	End If
	If Instr(sNode,"/A")>0 Then
		sNode=Replace(Trim(aNode(0)),"/A;1","")
	End If

	If Instr(sNode,"/B")>0 Then
		sNode=Replace(Trim(aNode(0)),"/B;1","")
	End If
	
	Call Fn_TabFolder_Operation("DoubleClickTab", sNode,"")
	wait 2
	iCols= objTable.GetROProperty("cols")
	For iCounter=0 To iCols-1
		sColumnName=objTable.GetColumnName(iCounter)
		If sColumnName=sColumn Then
			iColumn_no= iCounter-1
			Exit For
		End If
	Next
		Set ObjEdit=Description.Create()
		ObjEdit("Class Name").Value="JavaEdit"
		Set FiletrEdit=JavaWindow("SystemsEngineering").JavaObject("SEBOMHeader").ChildObjects(ObjEdit)
	Select Case sAction
		Case "Set"
		
			If sFilterType="Operator" Then
				'JavaWindow("SystemsEngineering").JavaEdit("FilterObjectEdit").SetTOProperty"index",(iColumn_no)
				FiletrEdit(iColumn_no).Set sValue
			End If
			
			If sFilterType="Attribute" Then
				FiletrEdit(iColumn_no+iCols-1).Set sValue
'				JavaWindow("SystemsEngineering").JavaEdit("FilterObjectEdit").SetTOProperty "index",(Cint(iColumn_no+iCols))
'				JavaWindow("SystemsEngineering").JavaEdit("FilterObjectEdit").Set sValue
			End If 
			Wait 2
			Set WshShell = CreateObject("WScript.Shell")
			WshShell.SendKeys "{ENTER}"
			Set WshShell = nothing
			If Err.Number<0 Then
				Fn_SE_FilterSetting=False
			Else
				Fn_SE_FilterSetting=True
			End If

		Case "Verify"
				If sFilterType="Operator" Then
'					JavaWindow("SystemsEngineering").JavaEdit("FilterObjectEdit").SetTOProperty"index",(iColumn_no)
					sGetValue=Cstr(FiletrEdit(iColumn_no).GetROProperty("value"))
				End If
			
				If sFilterType="Attribute" Then
					iColumn_no=iColumn_no-1
'					JavaWindow("SystemsEngineering").JavaEdit("FilterObjectEdit").SetTOProperty "index",(iColumn_no+iCols)
					sGetValue=Cstr(FiletrEdit(iColumn_no+iCols-1).GetROProperty("value"))
				End If 
				If sValue=sGetValue Then
					Fn_SE_FilterSetting=True
				Else
					Fn_SE_FilterSetting=False
				End If

	End Select
	Call Fn_TabFolder_Operation("DoubleClickTab", sNode,"")
	Set objTable=Nothing
	
End Function 

'*********************************************************		Function to  get Cell values of Detail table	***********************************************************************

'Function Name		:					Fn_SISW_SE_DetailsTable_GetCellData

'Description			 :		 		  This function is used to get Cell values of Detail table

'Parameters			   :	 			1.  objTable:Table Object
'												  2. iRow : Row Index
'												  3.. iCol : Column Index or Column Name									
											
'Return Value		   : 				 CellValue/false

'Pre-requisite			:		 		Detail Table is visible

'Examples				:				 Fn_SISW_SE_DetailsTable_GetCellData(JavaWindow("My Teamcenter - Teamcenter").JavaTable("DetailsTable"), 1, "Type")

'History:
'	Developer Name			Date				Rev. No.			Changes Done							Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Koustubh 				3-May-2012			1.0					Created									Koustubh
'	Sachin					22-May-2012			1.1					Added Case "Fnd0ListsCustomNotes"
'	Sachin					28-May-2012			1.2					Code Commented
'	Koustubh				29-Jun-2012			1.2					Modified code to get non Rlation column values
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_SE_DetailsTable_GetCellData(objTable, iRow, iCol)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SE_DetailsTable_GetCellData"
	Dim sPropName, sColName
	Dim sOutVal

	Fn_SISW_SE_DetailsTable_GetCellData = False

		'Code Commented by Sachin as discussed with vallari on 28-May-2012.
	'	sOutVal = objTable.GetCellData(iRow, iCol)

	'	If trim(sOutVal) = "" Then
		If IsNumeric(iCol) Then
			sColName = objTable.Object.getColumn(iCol).getText()
		Else
			sColName = iCol
		End If
	
		Select Case trim(sColName)
			Case "Type"
				sPropName = "object_type"
			Case "Owner"
				sPropName = "owning_user"
			Case "Object"
				sPropName = "object_string"
			Case "Group ID"
				sPropName = "owning_group"
			Case "Last Modified Date", "Date Modified"
				sPropName = "last_mod_date"
			Case "Checked-Out"
				sPropName = "checked_out"
			Case "Release Status"
				sPropName = "release_status_list"
			Case "Checked-Out By"
				sPropName = "checked_out_user"
			Case "Project IDs"
				sPropName = "project_ids"
			Case "Checked-Out Date"
				sPropName = "checked_out_date"			
			Case "Classified"				
				sPropName = "ics_classified"			
			Case "Classified in"				
				sPropName = "ics_subclass_name"	
			Case "Description"
				sPropName = "object_desc"	
		End Select
	
		If trim(sColName) = "Relation" Then
			sOutVal = objTable.Object.getItem(iRow).getData().getContext().toString()
			Select Case trim(sOutVal)
				Case "IMAN_reference"
					Fn_SISW_SE_DetailsTable_GetCellData = "References"
				Case "IMAN_specification"
					Fn_SISW_SE_DetailsTable_GetCellData = "Specifications"
				Case "contents"
					Fn_SISW_SE_DetailsTable_GetCellData = "Contents"
				Case "revision_list"
					Fn_SISW_SE_DetailsTable_GetCellData = "Revisions"
				Case "IMAN_master_form"
					' aaded s at the end - snehal salunkhe - 3-Apr-12 
					Fn_SISW_SE_DetailsTable_GetCellData = "Item Masters"
				Case "IMAN_classification"
					Fn_SISW_SE_DetailsTable_GetCellData = "Classification"
				Case "TC_Attaches"
					Fn_SISW_SE_DetailsTable_GetCellData = "Attaches"
				Case "IMAN_Rendering"
					Fn_SISW_SE_DetailsTable_GetCellData = "Rendering"
				Case "IMAN_manifestation"
					Fn_SISW_SE_DetailsTable_GetCellData = "Manifestations"
				Case "IMAN_aliasid"
					Fn_SISW_SE_DetailsTable_GetCellData = "Alias IDs"
				Case "release_status_list"
					Fn_SISW_SE_DetailsTable_GetCellData = "Release Status"
				Case "Fnd0ListsCustomNotes"
					Fn_SISW_SE_DetailsTable_GetCellData = "Custom Requirements Lists"
				Case Else
					Fn_SISW_SE_DetailsTable_GetCellData = sOutVal
			End Select
		Else
			On Error Resume Next
			sOutVal = ""
			sOutVal = objTable.Object.getItem(iRow).getData().getComponent().getProperty(sPropName)
			If sOutVal = "" Then
				sOutVal = objTable.Object.getItem(iRow).getData().getStringProperty(sPropName)
			End If
			Fn_SISW_SE_DetailsTable_GetCellData = sOutVal
		End If	
End Function
'*********************************************************		Function to  Set Filter  	***********************************************************************

'Function Name		:					Fn_SISW_SE_SetFilterDescription

'Description			 :		 		  This function is used to set filter.

'Parameters			   :	 			 1. sAction	:	Action to perform
'												  	  2. sFillterType : Up/Down
'												  	  3. sValues :  Values	to set [seperated by ~]   '[ needs to set all columns Values. ]
											
'Return Value		   : 				 true/false

'Pre-requisite			:		 		Filter should be Visible.

'Examples				:				 Fn_SISW_SE_SetFilterDescription("set","up","=~=~Contains~=~=")
'													 Fn_SISW_SE_SetFilterDescription("set","down","Analysis~All~Requirement~All~All")
'History:
'
'	Developer Name			Date				Rev. No.			Changes Done			  						Reviewer			Description
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Sachin 					28-May-2012			1.0				Created		
'	Chaitali R 				05-Jun-2012			1.0				Added Case :"select" 	         					Shweta R            Added Subcase :  "select"  in Case "Set" -  To select filter type from dropdown list
'	Shweta Rathod			04-Jul-2017			1.0				Added Case :"getfilterval_pos",getfilteraval_all 	Shweta R            Added case : To get filter data from given dropdown list
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_SE_SetFilterDescription(sAction, sFillterType, sValues)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SE_SetFilterDescription"
	Dim objSelectType,objTable,aValues,WshShell,intNoOfObjects,i,iCount,objSEWindow
    
	Fn_SISW_SE_SetFilterDescription = False
	
    Select Case lcase(sAction)
			Case "set"
					
                    If lcase(sFillterType) = "up" Then
							aValues = Split(sValues,"~")
							Set objTable  = JavaWindow("SystemsEngineering").JavaTable("FilterTable")
							Set WshShell = CreateObject("WScript.Shell")
							objTable.SetTOProperty "Index","1"
							Set objSelectType = Description.Create()
							objSelectType("to_class").value = "JavaEdit"
							Set  intNoOfObjects = objTable.ChildObjects(objSelectType)
							For i = 0 to intNoOfObjects.count-1
								'intNoOfObjects(i).highlight
								intNoOfObjects(i).setfocus
								wait 1
							   intNoOfObjects(i).Set aValues(i)
							   wait 2
							   WshShell.SendKeys "{ENTER}"
							   wait 3
							    If UBound(aValues)=i Then
								   Fn_SISW_SE_SetFilterDescription = True
								   Exit Function
							    End If
							Next
							wait 2
							Fn_SISW_SE_SetFilterDescription = True
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SISW_SE_SetFilterDescription ] executed successfully with case [ "+sFillterType+"]")

					ElseIf lcase(sFillterType) = "down"  Then
							aValues = Split(sValues,"~")
							Set objTable  = JavaWindow("SystemsEngineering").JavaTable("FilterTable")
							Set WshShell = CreateObject("WScript.Shell")
'							objTable.SetTOProperty "Index","2"
							objTable.SetTOProperty "Index","1" 'added this index to overcome OR changes in tc14.2
							Set objSelectType = Description.Create()
							objSelectType("to_class").value = "JavaEdit"
							Set  intNoOfObjects = objTable.ChildObjects(objSelectType)
							For i = 0 to intNoOfObjects.count-1
								'intNoOfObjects(i).highlight
								intNoOfObjects(i).setfocus
								wait 1
								intNoOfObjects(i).Set aValues(i)
							   wait 2
							   WshShell.SendKeys "{ENTER}"
							   wait 3
							     If UBound(aValues)=i Then
									Fn_SISW_SE_SetFilterDescription = True
									Exit Function
								 End If
							Next
							wait 2
							Fn_SISW_SE_SetFilterDescription = True
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SISW_SE_SetFilterDescription ] executed successfully with case [ "+sFillterType+"]")
					 
				'-------------------------------Added Case by Chaitali :  To select filter type from dropdown list------------------------------------------------------------					
					ElseIf lcase(sFillterType) = "select"  Then
							aValues = Split(sValues,"~")
							Set objTable  = JavaWindow("SystemsEngineering").JavaTable("FilterTable")
							Set WshShell = CreateObject("WScript.Shell")
							objTable.SetTOProperty "Index","1"
							Set objSelectType = Description.Create()
							objSelectType("to_class").value = "JavaButton"
							Set  intNoOfObjects = objTable.ChildObjects(objSelectType)
							
							For i = 0 to intNoOfObjects.count-1
								intNoOfObjects(i).highlight
								wait 1
								intNoOfObjects(i).click
							   	wait 2
							  
								Set objSEWindow = JavaWindow("SystemsEngineering")
	                             				objSEWindow.JavaWindow("Shell").SetTOProperty "index", i
					             		If objSEWindow.JavaWindow("Shell").JavaTable("TypeTable").Exist(1) Then
						         		bFlag = True
								Else
									bFlag = False				
								 End If
	
								If bFlag = True Then
									bFlag = False
									For iCount = 0 To objSEWindow.JavaWindow("Shell").JavaTable("TypeTable").GetROProperty("rows")
										If objSEWindow.JavaWindow("Shell").JavaTable("TypeTable").GetCellData(iCount, 0) = sValues Then
											objSEWindow.JavaWindow("Shell").JavaTable("TypeTable").ClickCell iCount, 0
											wait 1
											bFlag = True
											Exit For
										End If
									Next
								End If
								 wait 2
								If bFlag = True Then 
									Fn_SISW_SE_SetFilterDescription = True
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SISW_SE_SetFilterDescription ] executed successfully with case [ "+sFillterType+"]")					
									Exit For
								Else
									Fn_SISW_SE_SetFilterDescription = False
                       				 			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SE_SetFilterDescription ] execution failed with case [ "+sFillterType+"]")
								End If
								 
							Next	 
						  	 
					'---------------------------------------------------------------------------------------------------------------------------------------------------------------	----------------------------------								
					
					Else
						Fn_SISW_SE_SetFilterDescription = False
                        Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SE_SetFilterDescription ] execution failed with case [ "+sFillterType+"]")
					End If
			'-------------------------------Added Case by Shweta :  To get filter data from given dropdown list------------------------------------------------------------					
			Case "getfilterval_pos","getfilteraval_all" 
					Set objTable  = JavaWindow("SystemsEngineering").JavaTable("FilterTable")
					Set WshShell = CreateObject("WScript.Shell")
					
                    If lcase(sFillterType) = "up" Then
						objTable.SetTOProperty "Index","0" 'added this index to overcome OR changes in tc14.2
					ElseIf lcase(sFillterType) = "down"  Then							
							objTable.SetTOProperty "Index","1" 'added this index to overcome OR changes in tc14.2
					End if
					Set objSelectType = Description.Create()
					objSelectType("to_class").value = "JavaEdit"
					Set  intNoOfObjects = objTable.ChildObjects(objSelectType)
					
					If sAction = "getfilterval_pos" then
						i = sValues
						intNoOfObjects(i).highlight
						wait 1
						sRetValues = intNoOfObjects(i).object.getText()
						wait 1
						WshShell.SendKeys "{ENTER}"
						wait 1
						Fn_SISW_SE_SetFilterDescription = sRetValues
						Exit Function
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SISW_SE_SetFilterDescription ] executed successfully with case [ "+sFillterType+"]")
					End if
					If sAction = "getfilteraval_all" then
							For i = 0 to intNoOfObjects.count-1
								intNoOfObjects(i).highlight
								wait 1
								If i = 1 then
									sRetValues = intNoOfObjects(i).object.getText()
								else								
									sRetValues = sRetValues + "~"+intNoOfObjects(i).object.getText()
								End if
							   wait 2
							   WshShell.SendKeys "{ENTER}"
							   wait 3
							 Next
							 Fn_SISW_SE_SetFilterDescription = sRetValues
							Exit Function
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SISW_SE_SetFilterDescription ] executed successfully with case [ "+sFillterType+"]")
					End if
			
			Case Else
				'	Wrong case passed to the Function
				Fn_SISW_SE_SetFilterDescription = False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: the case [ "+sAction+"] not valid for Function [ Fn_SISW_SE_SetFilterDescription ] ")
		End Select

	'Clear allocated Memory
	Set objTable = Nothing
	Set objSelectType = Nothing
	Set WshShell = nothing
	Set objSEWindow = nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name			:	Fn_SISW_SE_BomCompareStructureOperation
'@@
'@@    Description				:	Function Used to Copare the BOM Structure 
'@@
'@@    Parameters			   	:	1. sAction			: Action [Type of Attribute Group]
'@@								:	2. bAddSource		: boolean flag to click on Set / Add Source button
'@@							 	:	3. bRemoveSource 	: boolean flag to click on Remove Source button
'@@							 	:	4. bAddTarget 		: boolean flag to click on Set / Add Target button
'@@							 	:	5. bRemoveTarget 	: boolean flag to click on Remove Target button
'@@							 	:	6. sMode			 : 	Mode Level/Name to select from Mode list 
'@@							 	:	7. sReport 			: ON/OFF  -Check Box Value 
'@@								:  8 : sButton            : Button Name to Click on
'@@							 	
'@@
'@@    Return Value		   	   	: 	True Or False
'@@
'@@    Pre-requisite			:	System Engineering perspective should be activated
'@@
'@@    Examples					:	Call  Fn_SISW_SE_BomCompareStructureOperation("Set", "", "", "true", "","Multi Level","on" ,"Ok~Cancel")
'@@
'@@		History				:	
'@@		Developer Name			Date			 Rev. No.	   Changes Done		 Reviewer
'@@		------------------------------------------------------------------------------------------------------------------------------
'@@		Avinash Jagale  		31-May-2012		  1.0			created                   Kausthubh W.
'@@	
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

Public function Fn_SISW_SE_BomCompareStructureOperation(sAction, bAddSource, bRemoveSource, bAddTarget, bRemoveTarget,sMode,sReport ,sButton)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SE_BomCompareStructureOperation"
	Dim objDialog,sButton1,iCount,arrMode
	Fn_SISW_SE_BomCompareStructureOperation = False
	Set objDialog =JavaWindow("SystemsEngineering").JavaWindow("BOMCompareStructure")
	If bRemoveSource = "" Then bRemoveSource = False
	If bAddSource = "" Then bAddSource = False
	If bRemoveTarget = "" Then bRemoveTarget = False
	If bAddTarget = "" Then bAddTarget = False

	If Fn_UI_ObjectExist("Fn_SISW_SE_BomCompareStructureOperation", objDialog) = False Then
		Call Fn_MenuOperation("Select","Tools:Compare:Structure Compare:BOM Compare...")
		Call Fn_ReadyStatusSync(2)
	If Fn_UI_ObjectExist("Fn_SISW_SE_BomCompareStructureOperation", objDialog) = False  Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_SE_BomCompareStructureOperation ] Failed to open Show Traceability Matrix window.")
			Set objDialog =nothing
			Exit function
		End If
	End If

	Select Case sAction
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "Set"
			If cBool(bRemoveSource) Then
				Call Fn_Button_Click("Fn_SISW_SE_BomCompareStructureOperation", objDialog, "RemoveSource")
			End If

			If cBool(bAddSource) Then 
				Call Fn_Button_Click("Fn_SISW_SE_BomCompareStructureOperation", objDialog, "AddSource")
			End If

			If cBool(bRemoveTarget) Then 
				Call Fn_Button_Click("Fn_SISW_SE_BomCompareStructureOperation", objDialog, "RemoveTarget")
			End If

			If cBool(bAddTarget) Then 
				Call Fn_Button_Click("Fn_SISW_SE_BomCompareStructureOperation", objDialog, "AddTarget")
			End If
			'Selecting Structure compare made
			If sMode <> "" Then 
				Call  Fn_List_Select("Fn_SISW_SE_BomCompareStructureOperation", objDialog,"ModeList",sMode)
			End If
			'Setting Report Option
			If  sReport  <> "" Then 
				 Call Fn_CheckBox_Set("Fn_SISW_SE_BomCompareStructureOperation" ,objDialog,"Report",sReport) 
			End If
			'Click on button
			If sButton <> "" Then
				Call Fn_Button_Click("Fn_SISW_SE_BomCompareStructureOperation", objDialog, sButton)
				If objDialog.Exist(3) Then
					Call Fn_Button_Click("Fn_SISW_SE_BomCompareStructureOperation", objDialog,"Cancel")
				End If
			End IF
			Fn_SISW_SE_BomCompareStructureOperation = True
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - -
		'[TC_11.4_NewDevelopment_PoonamC_08Aug2017 : Added Case to verify compare modes from list.]
		Case "VerifyCompareModes"
			arrMode = Split(sMode,"~")
			For iCount = 0 To UBound(arrMode)
				If Fn_SISW_UI_JavaList_Operations("Fn_SISW_SE_BomCompareStructureOperation","Exist",objDialog,"ModeList",arrMode(iCount),"","") = False Then
					Fn_SISW_SE_BomCompareStructureOperation = False	
					Exit For
				Else
					Fn_SISW_SE_BomCompareStructureOperation = True
				End If
			Next
			'Click on button
			If sButton <> "" Then
				Call Fn_Button_Click("Fn_SISW_SE_BomCompareStructureOperation", objDialog, sButton)
				If objDialog.Exist(3) Then
					Call Fn_Button_Click("Fn_SISW_SE_BomCompareStructureOperation", objDialog,"Cancel")
				End If
			End IF
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_SE_BomCompareStructureOperation ] Invalid case [ " & sAction & " ].")
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
	End Select
	If  Fn_SISW_SE_BomCompareStructureOperation <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [ Fn_SISW_SE_BomCompareStructureOperation ] executed successfuly with case [ " & sAction & " ].")
	End If
	Set objDialog = Nothing
End Function



''''''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''''''''/$$$$
''''''''/$$$$   FUNCTION NAME   :    Fn_SISW_SE_Create_And_OpenWebPage(sFilePath,sInfo1,sInfo2)
''''''''/$$$$
''''''''/$$$$   DESCRIPTION        :  This Function will create an HTML File and Open it as a webpage
''''''''/$$$$
''''''''/$$$$	PRE-REQUISITERS :  Teamcenter Browser should be in focus
''''''''/$$$$
''''''''/$$$$  PARAMETERS   : 		sFilePath : File Path  for the file
''''''''/$$$$										sInfo1 : For Future Use
''''''''/$$$$										sInfo2:	For Future Use
''''''''/$$$$	
''''''''/$$$$		Return Value : 				True or False
''''''''/$$$$
''''''''/$$$$    Function Calls       :   Fn_WriteLogFile()
''''''''/$$$$
''''''''/$$$$		HISTORY           :   		AUTHOR                 DATE        VERSION
''''''''/$$$$  
''''''''/$$$$    CREATED BY     :   SHREYAS          13/09/2012         1.0
''''''''/$$$$
''''''''/$$$$    REVIWED BY     :  Shreyas			13/09/2012           1.0
''''''''/$$$$ 
''''''''/$$$$    Modified BY     :   Anjali M         05/02/2013         1.1		Modified code to right click on [ WebSite ] java object
''''''''/$$$$
''''''''/$$$$		How To Use :      				Example #1
'''''''/$$$$																
'''''''/$$$$							bReturn=Fn_SISW_SE_Create_And_OpenWebPage("D:\Shree","","")
'''''''/$$$$
'''''''/$$$$
''''''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
public Function Fn_SISW_SE_Create_And_OpenWebPage(sFilePath,sInfo1,sInfo2)

GBL_FAILED_FUNCTION_NAME="Fn_SISW_SE_Create_And_OpenWebPage"
Fn_SISW_SE_Create_And_OpenWebPage=false
Dim shell,objWindow,iCount
Set objWindow=Fn_SISW_SE_GetObject("SystemsEngineering")
If Window("Notepad").Exist Then   'Added by Avinash
	Window("Notepad").Close
End If
objWindow.JavaObject("WebSite").Click 10,10, "LEFT"
If err.number<0 Then
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed Right Click on System Engineering Window")
	Fn_SISW_SE_Create_And_OpenWebPage=false
	Exit function
Else
	Fn_SISW_SE_Create_And_OpenWebPage=True
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Right Clicked on System Engineering Window")
End If

Set shell=CreateObject("Wscript.Shell")
'shell.SendKeys "+{F10}"
objWindow.JavaObject("WebSite").Click "50","50","RIGHT"

wait 1
For iCount=1 to 6
shell.SendKeys "{UP}"
Next
shell.SendKeys "{ENTER}"
wait 1
 Window("Notepad").Click 0,0,micLeftBtn
wait 1
shell.SendKeys "^(a)"
wait 1
shell.SendKeys "^(c)"
wait 1
shell.SendKeys "^(n)"
wait 1
shell.SendKeys "^(v)"
wait 1
shell.SendKeys "^(s)"
wait 1
Window("Notepad").Dialog("Save As").WinEdit("File name:").Set sFilePath+".html"
sFilePath=sFilePath+".html"								''Added by Avinash J.		
If err.number<0 Then
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed To set value ["+sFilePath+"]")
	Fn_SISW_SE_Create_And_OpenWebPage=false
	Exit function
Else
	Fn_SISW_SE_Create_And_OpenWebPage=True
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully set value ["+sFilePath+"]")	
End If
wait 1
Window("Notepad").Dialog("Save As").WinButton("Save").Click 0,0,micLeftBtn
If err.number<0 Then
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed Click Save Button")
	Fn_SISW_SE_Create_And_OpenWebPage=false
	Exit function
Else
	Fn_SISW_SE_Create_And_OpenWebPage=True
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Clicked Save Button")
End If
Fn_SISW_SE_Create_And_OpenWebPage=True
wait 1
If Window("Notepad").Exist Then
	Window("Notepad").Close
End If

Systemutil.Run sFilePath
Wait 5

End Function
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Function Name		:				Fn_SE_ComponentAndSEBOMTableNodeOpeations

'Description			 :		 		 This function is used to Perform operations on all BOM tables ( Component , SE )

'Parameters			   :	 			1. StrAction: Select the Paragraph 
'										2. StrTableType: Table type ( Component , SE )
'										3. iTableInstance: Table instance as 2 table of same name opens
'										4. StrNodeName: Node to select the tree
'										5. StrColName: Column name of the table
'										6. StrColValue: Value of the column
'										7. StrPopupMenu: Popup Menu

'Return Value		   : 				True False

'Pre-requisite			:		 		Should be in SE Prespective

'Examples				:				Fn_SE_ComponentAndSEBOMTableNodeOpeations("Select","ComponentBOMTable","2","",DataTable("SMPath", dtGlobalSheet),"", "", "")
'										Fn_SE_ComponentAndSEBOMTableNodeOpeations("ExpandBelow","ComponentBOMTable","2","",DataTable("SMPath", dtGlobalSheet),"", "", "")
'										Fn_SE_ComponentAndSEBOMTableNodeOpeations("AddColumn","ComponentBOMTable","2","","","Number", "", "")
'										Fn_SE_ComponentAndSEBOMTableNodeOpeations("CompareColValue","ComponentBOMTable","2","",DataTable("SMPath", dtGlobalSheet),"Number", Cstr(iCnt), "")
'History:
'										Developer Name			Date			Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Amol					04.03.2013													Sandeep N
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function Fn_SE_ComponentAndSEBOMTableNodeOpeations(StrAction,StrTableType,iTableInstance,StrTableTabName,StrNodeName,StrColName, StrColValue, StrPopupMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_SE_ComponentAndSEBOMTableNodeOpeations"
	on Error Resume Next
	Dim iRowNo, sMenu, iNodeNo, iColNo, iStart,strName, objContextMenu
	Dim objTable, iCnt, aColumns, objChangeCol, sColName
	Dim bFlag,sValues
	Dim sColValue,aValues,iCounter,i
	Dim iRows,iPathCount,StrNodePath,iCount
	Dim objTabFld,StrTabName

	bFlag=False
	Fn_SE_ComponentAndSEBOMTableNodeOpeations=False
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	'On which table want to perform operations
	If StrTableType="SEBOMTable" Then
		Set objTabFld =JavaWindow("SystemsEngineering").JavaTab("SEComponentTab")
	Elseif StrTableType="ComponentBOMTable" then
		Set objTabFld =JavaWindow("SystemsEngineering").JavaTab("RACTabFolderWidget")
	Else
        Set objTabFld =JavaWindow("SystemsEngineering").JavaTab("SEComponentTab")
	End If
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	Set objTable = Window("SEWindow").JavaWindow("WEmbeddedFrame").JavaTable("CMEBOMTreeTable") 
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	'Code to handle Applet as applet index is dyanamically changes
	If StrTableTabName="" Then
		i=objTabFld.Object.getSelectionIndex()   '[Tc112-2017091400-10_10_2017-JotibaT]--Added code as per object changed.
		StrTabName=objTabFld.Object.getItem(i).text()
		StrTabName=Split(StrTabName,"-")
	Else
		StrTabName=StrTableTabName
	End If
	'- - - - - - - Added Code to handle Requirement Opened in BOM table
	If StrTabName(0)="REQ" Then
		StrTabName(0)=StrTabName(1)
	End If
	If iTableInstance="" Then
		iTableInstance=1
	End If
	i=0
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	For iCounter=0 to 12
		 Window("SEWindow").JavaWindow("WEmbeddedFrame").SetTOProperty "index",iCounter
		 If Window("SEWindow").JavaWindow("WEmbeddedFrame").JavaTable("CMEBOMTreeTable").Exist(2) Then
			If InStr(1,trim(Window("SEWindow").JavaWindow("WEmbeddedFrame").JavaTable("CMEBOMTreeTable").Object.getValueAt(0,0).toString()),trim(StrTabName(0))) Then
				i=i+1
				If i=CInt(iTableInstance) Then
					bFlag=True
					Exit for
				End If
			End If
		 End If
	Next
	If bFlag=false Then
		Exit function
	End If
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 

		Select Case StrAction

			Case "Select"		'("Select"," 000452/A;1-ReqSpec (View):REQ-000047/A;1-Req2 (View):000454/A;1-P2 (View):REQ-000048/A;1-Req3","","","")
				StrNodeName = Split(strNodeName,"-",-1,1)
				If len(StrNodeName(uBound(StrNodeName))) <= 1 or IsNumeric(StrNodeName(uBound(StrNodeName))) Then
					StrNodeName(uBound(StrNodeName)) = StrNodeName(uBound(StrNodeName)-1) + "-" +StrNodeName(uBound(StrNodeName))
				End If
				iRowNo = Fn_SE_BOMTable_RowIndex(StrNodeName(uBound(StrNodeName)))
				If isNumeric(iRowNo) Then
					objTable.Object.clearSelection()   'Added by pratap for focus issues during selection build tc 12.1 20181017
					objTable.SelectRow iRowNo
					Fn_SE_ComponentAndSEBOMTableNodeOpeations=True
				End if

			Case "IsSelected"		'("IsSelected"," 000040/A;1-Spec1 (View):REQ-000004/A;1-Req2","","","")
				iRowNo = Fn_SE_BOMTable_RowIndex(StrNodeName)				
				If isNumeric(iRowNo) Then
					If Cint(objTable.GetROProperty("SelectedRow")) = Cint(iRowNo) Then						
						Fn_SE_ComponentAndSEBOMTableNodeOpeations=True
					End If
				End if

			Case "Deselect"		'("Select"," 000452/A;1-ReqSpec (View):REQ-000047/A;1-Req2 (View):000454/A;1-P2 (View):REQ-000048/A;1-Req3","","","")
				iRowNo = Fn_SE_BOMTable_RowIndex(StrNodeName)
				If isNumeric(iRowNo) Then
					objTable.DeselectRow iRowNo
					Fn_SE_ComponentAndSEBOMTableNodeOpeations=True
				End if

			Case "VerifyNode"		'("Verify"," 000452/A;1-ReqSpec (View):REQ-000047/A;1-Req2 (View):000454/A;1-P2 (View)","","","")		
        			'Verify Node Exist
					StrNodeName = Split(strNodeName,"-",-1,1)
				If len(StrNodeName(uBound(StrNodeName))) <= 1 or IsNumeric(StrNodeName(uBound(StrNodeName))) Then
					StrNodeName(uBound(StrNodeName)) = StrNodeName(uBound(StrNodeName)-1) + "-" +StrNodeName(uBound(StrNodeName))
				End If
				iRowNo = Fn_SE_BOMTable_RowIndex(StrNodeName(uBound(StrNodeName)))
					'iRowNo = Fn_SE_BOMTable_RowIndex(StrNodeName)
					If isNumeric(iRowNo) then
						Fn_SE_ComponentAndSEBOMTableNodeOpeations=True
					End if
        
			Case "getNodeIndex"	'("getNodeIndex"," 000452/A;1-ReqSpec (View):REQ-000047/A;1-Req2 (View):000454/A;1-P2 (View)","","","")
				iRowNo = Fn_SE_BOMTable_RowIndex(StrNodeName)
				If isNumeric(iRowNo) then
					Fn_SE_ComponentAndSEBOMTableNodeOpeations=iRowNo
				End if

			Case "Expand"	'("Expand"," 000452/A;1-ReqSpec (View):REQ-000047/A;1-Req2 (View)","","","")
				iRowNo = Fn_SE_BOMTable_RowIndex(StrNodeName)
				If isNumeric(iRowNo) then
					objTable.SelectRow iRowNo
					'Code modified By Ketan on 06/09/2011 as Expand call is not working.				
					call Fn_menuOperation("Select","View:Expand Options:Expand Below...")
'					Call Fn_Edit_Box("Fn_SE_ComponentAndSEBOMTableNodeOpeations",JavaWindow("SystemsEngineering").JavaWindow("WEmbeddedFrame").JavaDialog("ExpandToLevel"),"Level","1")
			        JavaWindow("SystemsEngineering").JavaWindow("ExpandToLevel").JavaSpin("Spinner").Set "1"
					  Call Fn_Button_Click("Fn_SE_ComponentAndSEBOMTableNodeOpeations", JavaWindow("SystemsEngineering").JavaWindow("ExpandToLevel"), "OK")
					Fn_SE_ComponentAndSEBOMTableNodeOpeations=True
				End if

			Case "Collapse"		'("Collapse"," 000040/A;1-Spec1 (View)","","","")
				iRowNo = Fn_SE_BOMTable_RowIndex(StrNodeName)
				If isNumeric(iRowNo) Then
					objTable.SelectRow iRowNo
					Fn_SE_ComponentAndSEBOMTableNodeOpeations=True
					If Err.Number < 0 Then						
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail: Function [Fn_SE_ComponentAndSEBOMTableNodeOpeations] Failed to Select SE BOM Table Node [" + StrNodeName + "]")
						Fn_SE_ComponentAndSEBOMTableNodeOpeations = FALSE			
					Else
						'Operate Collapse Below Menu if Node selected Sucessfully
						StrReturn = Fn_MenuOperation("Select", "View:Collapse Below")
						If StrReturn = True Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Pass: Function [Fn_SE_ComponentAndSEBOMTableNodeOpeations] Sucessfully CollaSEd SE BOM Table Node [" + StrNodeName + "]")							
							Fn_SE_ComponentAndSEBOMTableNodeOpeations = TRUE
						Else							
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail: Function [Fn_SE_ComponentAndSEBOMTableNodeOpeations] Failed to CollaSE SE BOM Table Node [" + StrNodeName + "]")
							Fn_SE_ComponentAndSEBOMTableNodeOpeations = FALSE
						End If						
					End If
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail: Function [Fn_SE_ComponentAndSEBOMTableNodeOpeations] Failed to Get SE BOM Table Node [" + StrNodeName + "]")
					Fn_SE_ComponentAndSEBOMTableNodeOpeations = FALSE
				End If

			Case "PopupMenuSelect"	'("PopupMenuSelect","","","","Trace Link:Start Trace Link")
				'Pre-requisite = Row should be selected
				StrPopupMenu=Replace(StrPopupMenu,":",";")
				iRowNo = objTable.Object.getSelectedRow()
				If isNumeric(iRowNo) then
					'JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet").JavaTable("RMTable").SelectRow iRowNo
					objTable.ClickCell iRowNo,0,"RIGHT" 
					wait 1
					sMenu = JavaWindow("SystemsEngineering").WinMenu("ContextMenu").BuildMenuPath(StrPopupMenu)
					wait 1
					JavaWindow("SystemsEngineering").WinMenu("ContextMenu").Select sMenu
					Fn_SE_ComponentAndSEBOMTableNodeOpeations=True
				End if

			Case "MultiSelect"		'("MultiSelect","REQ-000049/A;1-Req1 (View):000455/A;1-P1 (View):000575/A;1-P23~REQ-000049/A;1-Req1 (View):REQ-000148/A;1-Req2","","","")

				StrNodeName=split(StrNodeName,"~") 
				For iNodeNo=0 to Ubound(StrNodeName)
					iRowNo = Fn_SE_BOMTable_RowIndex(StrNodeName(iNodeNo))
					If isNumeric(iRowNo) Then
						If iNodeNo=0 Then
							objTable.SelectRow iRowNo
							Fn_SE_ComponentAndSEBOMTableNodeOpeations=True
						Else
							objTable.ExtendRow "#"&iRowNo
							Fn_SE_ComponentAndSEBOMTableNodeOpeations=True
						End If					
					End if
				Next
				
			Case "VerifyColValue"	'("VerifyColValue"," REQ-000049/A;1-Req1 (View):000455/A;1-P1 (View)","Item Type","Requirement","")
				iRowNo = Fn_SE_BOMTable_RowIndex(StrNodeName)
            	If isNumeric(iRowNo) then
					'Get column Rows
					iColNo = objTable.GetROProperty("cols")

					For iStart=0 to iColNo-1
						''Verify the Column name is similar to required column name
						If objTable.GetColumnName(iStart)=StrColName Then
							'Verify the Column value is similar to required value
							If cstr(objTable.GetCellData(iRowNo,iStart))=cstr(StrColValue) then
								Fn_SE_ComponentAndSEBOMTableNodeOpeations=True
							End if
							Exit For
						End If
					Next
				Else
					Fn_SE_ComponentAndSEBOMTableNodeOpeations=False
				End if

			Case "EditColValue"		'("EditColValue"," REQ-000049/A;1-Req1 (View):000455/A;1-P1 (View)","Find No.1","50","")	

				iRowNo = Fn_SE_BOMTable_RowIndex(StrNodeName)
            	If isNumeric(iRowNo) then
					
					'Get column Rows
					iColNo=objTable.GetROProperty("cols")
					For iStart=0 to iColNo-1
						''Verify the Column name is similar to required column name
						If objTable.GetColumnName(iStart)=StrColName Then
							'Verify the Column value is similar to required value
							If StrColName = "IP Classification" Then
'								objTable.SetCellData iRowNo,iStart,StrColValue
								objTable.ClickCell iRowNo,StrColName
								wait 1
								Window("SEWindow").JavaWindow("WEmbeddedFrame").JavaEdit("LOVSelectionDisplayView").Set StrColValue+vblf 
							Else
								objTable.SetCellData iRowNo,iStart,StrColValue
							End If
							Fn_SE_ComponentAndSEBOMTableNodeOpeations=True
							Exit For
						End If
					Next
				Else
					Fn_SE_ComponentAndSEBOMTableNodeOpeations=False
				End if
				Case "GetCellData" '("GetCellData",1,0,"","")
					
					
					'(IMP NOTE)' "StrNodeName" - This parameter is use as Row number in this Case
					'StrColName - This parameter is use as column nuber in this case
					
						objTable.SelectRow StrNodeName
						wait(3)
						strName = objTable.GetCellData(StrNodeName,StrColName)
						
					If Err.number < 0 Then
                		Fn_SE_ComponentAndSEBOMTableNodeOpeations=False
					Else
						Fn_SE_ComponentAndSEBOMTableNodeOpeations = MId(strName,instr(1,strName,":")+1 , Len(strName))
					End If

			Case "PopupMenuExist"		
						StrPopupMenu=Replace(StrPopupMenu,":",";")
						iRowNo = objTable.Object.getSelectedRow()
						If isNumeric(iRowNo) then
							objTable.ClickCell iRowNo,0,"RIGHT" 
							wait 1
							sMenu = JavaWindow("SystemsEngineering").WinMenu("ContextMenu").BuildMenuPath(StrPopupMenu)
							If JavaWindow("SystemsEngineering").WinMenu("ContextMenu").GetItemProperty (sMenu,"Exists") = True Then
									Fn_SE_ComponentAndSEBOMTableNodeOpeations = TRUE
							Else
									Fn_SE_ComponentAndSEBOMTableNodeOpeations = FALSE
							End If
						End If
		Case "DoubleClickCell"		'("DoubleClickCell"," 000452/A;1-ReqSpec (View):REQ-000047/A;1-Req2 (View):000454/A;1-P2 (View):REQ-000048/A;1-Req3","Has Attached Notes","","")
				iRowNo = Fn_SE_BOMTable_RowIndex(StrNodeName)
				If isNumeric(iRowNo) Then
					objTable.SelectRow iRowNo
					Fn_SE_ComponentAndSEBOMTableNodeOpeations=True
				End if
				objTable.DoubleClickCell iRowNo,StrColName
		Case "EditNumber"
				StrNodeName = Split(strNodeName,"-",-1,1)
				If len(StrNodeName(uBound(StrNodeName))) <= 1 or IsNumeric(StrNodeName(uBound(StrNodeName))) Then
					StrNodeName(uBound(StrNodeName)) = StrNodeName(uBound(StrNodeName)-1) + "-" +StrNodeName(uBound(StrNodeName))
				End If
				iRowNo = Fn_SE_BOMTable_RowIndex(StrNodeName(uBound(StrNodeName)))'
				
				
				'iRowNo = Fn_SE_BOMTable_RowIndex(StrNodeName)
				If IsNumeric(iRowNo) Then
					objTable.SelectRow iRowNo					
					'Open the Edit Number Dialog.
					objTable.DoubleClickCell iRowNo,StrColName
					'Set the New number value.
					Call Fn_Edit_Box("Fn_SE_ComponentAndSEBOMTableNodeOpeations",JavaWindow("SystemsEngineering").JavaWindow("EditNumber"),"NewNumber","")	
					Call Fn_UI_EditBox_Type("Fn_SE_ComponentAndSEBOMTableNodeOpeations",JavaWindow("SystemsEngineering").JavaWindow("EditNumber"),"NewNumber",StrColValue)
					'Click on Ok button
					Call Fn_Button_Click("Fn_SE_ComponentAndSEBOMTableNodeOpeations", JavaWindow("SystemsEngineering").JavaWindow("EditNumber"), "OK")
					Fn_SE_ComponentAndSEBOMTableNodeOpeations=True
				End If
		Case "ColumnExists"
					iColNo = cInt(objTable.GetROProperty("cols"))
					For iCnt = 0 to iColNo -1
						If trim(objTable.GetColumnName(iCnt)) = StrColName then
							Fn_SE_ComponentAndSEBOMTableNodeOpeations = True
							Exit for
						end if
					Next
		Case "AddColumn", "AddColumns"
					' checking existance of columns 
					iColNo = cInt(objTable.GetROProperty("cols"))
					aColumns = split(StrColName,"~")
					For iColumnCnt = 0 to UBound(aColumns)
						For iCnt = 0 to iColNo -1
							' if exists then ignoring that column
							If trim(objTable.GetColumnName(iCnt)) = aColumns(iColumnCnt) then
								aColumns(iColumnCnt) = ""
								Exit for
							end if
						Next
					Next
					'' Change parameter from 1 to 0 by dipali
					If  iColNo > 1 Then
						objTable.SelectColumnHeader 1,"LEFT"
						wait 3
						objTable.SelectColumnHeader 1,"RIGHT"
						wait 2
					Else
						objTable.SelectColumnHeader 0,"LEFT"
						wait 3
						objTable.SelectColumnHeader 0,"RIGHT"
						wait 2
					End If
					Window("SEWindow").JavaWindow("WEmbeddedFrame").JavaMenu("label:=Insert column\(s\) \.\.\.").Select

					If Window("SEWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Change Columns").Exist(5) Then
						Set objChangeCol = Window("SEWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Change Columns")
						For iColumnCnt = 0 to UBound(aColumns)
							If aColumns(iColumnCnt) <> "" Then
								Call Fn_List_Select("Fn_SE_ComponentAndSEBOMTableNodeOpeations", objChangeCol,"ListAvailableCols",aColumns(iColumnCnt))
								Call Fn_Button_Click("Fn_SE_ComponentAndSEBOMTableNodeOpeations", objChangeCol,"Add")
								wait 1
							End If
						Next
						if cInt(objChangeCol.JavaButton("Apply").getROProperty("enabled")) = 1 then
							Call Fn_Button_Click("Fn_SE_ComponentAndSEBOMTableNodeOpeations", objChangeCol,"Apply")
						end if
						Call Fn_Button_Click("Fn_SE_ComponentAndSEBOMTableNodeOpeations", objChangeCol,"Cancel")
						Fn_SE_ComponentAndSEBOMTableNodeOpeations = True
					End if
			' GetRowCount - Case will return total number of row count in Table
			Case "GetRowCount"
					Fn_SE_ComponentAndSEBOMTableNodeOpeations=Fn_UI_Object_GetROProperty("Fn_SE_ComponentAndSEBOMTableNodeOpeations",objTable,"rows")

			Case "ExpandBelow"	'("ExpandBelow"," 000452/A;1-ReqSpec (View):REQ-000047/A;1-Req2 (View)","","","")
				StrNodeName = Split(strNodeName,"-",-1,1)
				If len(StrNodeName(uBound(StrNodeName))) <= 1 or IsNumeric(StrNodeName(uBound(StrNodeName))) Then
					StrNodeName(uBound(StrNodeName)) = StrNodeName(uBound(StrNodeName)-1) + "-" +StrNodeName(uBound(StrNodeName))
				End If
				iRowNo = Fn_SE_BOMTable_RowIndex(StrNodeName(uBound(StrNodeName)))
				If isNumeric(iRowNo) then
					objTable.Object.clearselection()    ' Added by Pratap For focus issue on tc 12.1 build 201801017
					objTable.SelectRow iRowNo
					Call Fn_menuOperation("Select","View:Expand Options:Expand Below")
					If JavaWindow("SystemsEngineering").JavaWindow("SEWindow").JavaDialog("ExpandBelow").exist(3) = False Then
						Call Fn_KeyBoardOperation("SendKeys","{ESC}")
						Call Fn_menuOperation("WinMenuSelect","View:Expand Options:Expand Below")
					End If
					Call Fn_Button_Click("Fn_SE_ComponentAndSEBOMTableNodeOpeations", JavaWindow("SystemsEngineering").JavaWindow("SEWindow").JavaDialog("ExpandBelow"), "Yes")
					Fn_SE_ComponentAndSEBOMTableNodeOpeations=True
				End if
				
			Case "SelectAll"
				'Clear previously selected Nodes
				objTable.Object.clearSelection
				iRows = cInt(objTable.GetROProperty ("rows"))
				For iCounter = 0 to iRows - 1
					objTable.ExtendRow iCounter
				Next
				Fn_SE_ComponentAndSEBOMTableNodeOpeations = True		


			Case "CompareColValue"   'Added by Pooja S :  2/2/2012
					bFlag=false
					sColValue = Fn_SE_ComponentAndSEBOMTableNodeOpeations("GetCellData",StrTableType,iTableInstance,StrTableTabName,StrNodeName,StrColName, "", "")
					wait(2)
					aValues = split(sColValue,",",-1,1)
					For iCounter = 0 to Ubound(aValues)				
									For i=0 to Ubound(StrColValue)
										If 	lCase(trim(aValues(iCounter)))=lCase(trim(StrColValue(i))) Then
												bFlag=true
												Exit for 
										End If	
									Next
						Next	
						If bFlag=true Then
									Fn_SE_ComponentAndSEBOMTableNodeOpeations=True
						Else
									Fn_SE_ComponentAndSEBOMTableNodeOpeations=False
						End If
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		Case "PopupMenuEnabled"
			If StrNodeName <> "" Then
				iRowNo = Fn_SE_BOMTable_RowIndex(StrNodeName)
				If iRowNo <> -1 Then
					'Split Context menu to Build Path Accordingly
					sMenu = split(StrPopupMenu,":",-1,1)
					If sColName = "" Then
						objTable.ClickCell iRowNo ,"Requirement", "RIGHT","NONE"
					Else
						objTable.ClickCell iRowNo ,sColName, "RIGHT","NONE"
					End If
					Select Case cInt(Ubound(sMenu))
						Case 0
							Set objContextMenu = JavaWindow("SystemsEngineering").WinMenu("ContextMenu")
							objTable.ClickCell iRowNo,"Requirement", "RIGHT","NONE"
							Wait(2)
							If objContextMenu.CheckItemProperty (StrPopupMenu, "Exists",true,10) Then
								Fn_SE_ComponentAndSEBOMTableNodeOpeations = objContextMenu.CheckItemProperty (StrPopupMenu, "Enabled",true,10)
							End IF
							objTable.ClickCell iRowNo,"Requirement", "LEFT","NONE"
							Set objContextMenu = nothing
					End Select
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [Fn_SE_ComponentAndSEBOMTableNodeOpeations] Popup Menu ["+ StrPopupMenu +"] Selected Sucessfully")
				Else
					Fn_SE_ComponentAndSEBOMTableNodeOpeations = False
				End If
			Else
				Fn_SE_ComponentAndSEBOMTableNodeOpeations = False
			End If
	'- - - - - - -  Added Case by Sandeep: Case to return All Column names currently exist in BOM Table
			Case "AllColumnNames"
					'Returning All column Names present in BOM Table
					Fn_SE_ComponentAndSEBOMTableNodeOpeations =Fn_UI_TableOperations("Fn_SE_ComponentAndSEBOMTableNodeOpeations","GetAllColumnNames",objTable,"","")
		'- - - - - - - - - - - - - Added Case to Remove Column from Table
		Case "RemoveColumn"
					'checking existance of columns 
					bFlag=False
					iColNo = cInt(objTable.GetROProperty("cols"))
					aColumns = split(StrColName,"~")
					For iColumnCnt = 0 to UBound(aColumns)
						For iCnt = 0 to iColNo -1
							' if exists then ignoring that column
							If trim(objTable.GetColumnName(iCnt)) = aColumns(iColumnCnt) then
								bFlag=True
								Exit for
							end if
						Next
					Next
					'' Change parameter from 1 to 0 by dipali
					objTable.SelectColumnHeader 1,"RIGHT"
					Window("SEWindow").JavaWindow("WEmbeddedFrame").JavaMenu("label:=Insert column\(s\) \.\.\.").Select
					If Window("SEWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Change Columns").Exist(5) Then
						Set objChangeCol = Window("SEWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Change Columns")
						For iColumnCnt = 0 to UBound(aColumns)
							If bFlag=True Then
								Call Fn_List_Select("Fn_SE_ComponentAndSEBOMTableNodeOpeations", objChangeCol,"ListDisplayedCols",aColumns(iColumnCnt))
								Call Fn_Button_Click("Fn_SE_ComponentAndSEBOMTableNodeOpeations", objChangeCol,"Remove")
								wait 1
							End If
						Next
						if cInt(objChangeCol.JavaButton("Apply").getROProperty("enabled")) = 1 then
							Call Fn_Button_Click("Fn_SE_ComponentAndSEBOMTableNodeOpeations", objChangeCol,"Apply")
						end if
						Call Fn_Button_Click("Fn_SE_ComponentAndSEBOMTableNodeOpeations", objChangeCol,"Cancel")
						Fn_SE_ComponentAndSEBOMTableNodeOpeations = True
					End if

				Case "ConfigureColumn", "ConfigureColumns"
					' checking existance of columns 
					iColNo = cInt(objTable.GetROProperty("cols"))
					aColumns = split(StrColName,"~")
					For iColumnCnt = 0 to UBound(aColumns)
						For iCnt = 0 to iColNo -1
							' if exists then ignoring that column
							If trim(objTable.GetColumnName(iCnt)) = aColumns(iColumnCnt) then
								aColumns(iColumnCnt) = ""
								Exit for
							end if
						Next
					Next

					  objTable.SelectColumnHeader 1,"RIGHT"
					Window("SEWindow").JavaWindow("WEmbeddedFrame").JavaMenu("label:=Insert column\(s\) \.\.\.").Select
					If Window("SEWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Change Columns").Exist(5) Then
						Set objChangeCol = Window("SEWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Change Columns")
			'Delete all the Existing Columns from the Displayed Columns list
				sValues=objChangeCol.JavaList("ListDisplayedCols").GetROProperty("items count")
				For iCount=0 to cInt(sValues)-1
					sNode=objChangeCol.JavaList("ListDisplayedCols").GetItem(iCount)
					objChangeCol.JavaList("ListDisplayedCols").ExtendSelect  sNode
				
				Next
				objChangeCol.JavaButton("Remove").Click micLeftBtn

'Now add new Columns
aColumns=split(StrColName,":",-1,1)
						For iColumnCnt = 0 to UBound(aColumns)
							If aColumns(iColumnCnt) <> "" Then
								Call Fn_List_Select("Fn_SE_ComponentAndSEBOMTableNodeOpeations", objChangeCol,"ListAvailableCols",aColumns(iColumnCnt))
								Call Fn_Button_Click("Fn_SE_ComponentAndSEBOMTableNodeOpeations", objChangeCol,"Add")
								wait 1
							End If
						Next
						if cInt(objChangeCol.JavaButton("Apply").getROProperty("enabled")) = 1 then
							Call Fn_Button_Click("Fn_SE_ComponentAndSEBOMTableNodeOpeations", objChangeCol,"Apply")
						end if
						Call Fn_Button_Click("Fn_SE_ComponentAndSEBOMTableNodeOpeations", objChangeCol,"Cancel")
						Fn_SE_ComponentAndSEBOMTableNodeOpeations = True
					End if
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			Case "GetNodePathByName"
					iRows=objTable.GetROProperty("rows")
					bFlag=False
					For iCounter=0 to iRows-1	
						If InStr(1,trim(objTable.Object.getPathForRow(iCounter).toString()),StrNodeName) then
							iPathCount=objTable.Object.getPathForRow(iCounter).getPathCount()
							StrNodePath=objTable.Object.getPathForRow(iCounter).getPathComponent(1).toString()
							For iCount=2 to iPathCount-1
								StrNodePath=StrNodePath+":"+objTable.Object.getPathForRow(iCounter).getPathComponent(iCount).toString()
							Next
							bFlag=true
							Exit for
						end if
					Next
					If bFlag=True Then
						Fn_SE_ComponentAndSEBOMTableNodeOpeations=StrNodePath
					else
						Fn_SE_ComponentAndSEBOMTableNodeOpeations=False
					End If

			Case "VerifyForegroundColour", "VerifyBackgroundColour"

				If StrNodeName <> "" Then
				
					StrNodeName = Split(strNodeName,"-",-1,1)
					If len(StrNodeName(uBound(StrNodeName))) <= 1 or IsNumeric(StrNodeName(uBound(StrNodeName))) Then
						StrNodeName(uBound(StrNodeName)) = StrNodeName(uBound(StrNodeName)-1) + "-" +StrNodeName(uBound(StrNodeName))
					End If
					iRowCounter = Fn_SE_BOMTable_RowIndex(StrNodeName(uBound(StrNodeName)))
'					iRowCounter = Fn_SE_BOMTable_RowIndex(StrNodeName)
					If cint(iRowCounter) = -1 Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [Fn_SE_ComponentAndSEBOMTableNodeOpeations] Couldnt find  BOM Table Node [" + StrNodeName + "]")
						Exit function
					End If
					iRows = iRowCounter +1
					iCount = iRowCounter
				Else
					iRows = objTable.GetROProperty("rows")
					iCount = 0
				End If

				Do While cint(iCount) < cint(iRows)
					Set  objNodeForRow =  objTable.Object.getNodeForRow(cint(iCount))
					' if background colour
					If StrAction = "VerifyBackgroundColour" Then
						sColour = objTable.Object.getBackground(objNodeForRow,False).toString()
					Else
					' if foreground colour
						sColour = objTable.Object.getForeground(objNodeForRow,False).toString()
					End If
	
					sColour =  mid(sColour ,instr(sColour ,"[")  ,instr(sColour ,"]") )
					' Comparing colour codes RGB
					Select Case cstr(StrPopupMenu)
						Case "BLACK"
							sColourCode = "[r=0,g=0,b=0]"
						Case "LIGHTPINK"
							sColourCode = "[r=253,g=217,b=220]"
						Case "WHITE"
							sColourCode =  "[r=255,g=255,b=255]"
						Case "GRAY"
							sColourCode = "[r=178,g=180,b=191]" 
						Case "DARKGRAY"
							sColourCode = "[r=128,g=128,b=128]"
						Case "DARKBLUE"
							sColourCode = "[r=0,g=0,b=255]" 
						Case "LIGHTBLUE"
							sColourCode = "[r=183,g=219,b=255]"
						Case "GREEN"
							sColourCode = "[r=80,g=176,b=128]"
						Case "DARKGREEN"
							sColourCode = "[r=0,g=255,b=0]"
						Case "LIGHTGREEN"
							sColourCode = "[r=159,g=255,b=159]"
						Case "ORANGE"
							sColourCode = "[r=255,g=200,b=0]"
						Case "RED"
							sColourCode = "[r=255,g=0,b=0]" 
						Case "LIGHTRED"
							sColourCode = "[r=255,g=121,b=121]" 
						Case "YELLOW"
							sColourCode = "[r=255,g=255,b=0]"
						Case "YELLOWISHORANGE"
							sColourCode = "[r=254,g=190,b=95]"
						Case Else
							Exit function
					End Select
					
					If sColour = sColourCode  Then
						Fn_SE_ComponentAndSEBOMTableNodeOpeations = True
					Else
						Fn_SE_ComponentAndSEBOMTableNodeOpeations = False
						Exit function
					End If
					iCount = iCount +1
				Loop
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [Fn_SE_ComponentAndSEBOMTableNodeOpeations] Successfully verified colour code [ " & sColourCode & " ] for case [" & StrPopupMenu & "]")
				Set objNodeForRow = nothing

		End Select
	Set objTable = nothing
End Function
'***********************************************************************************************************************************************************************************
'Function Name		         :	       Fn_SISW_SE_PropertiesOperations()

'Description			   :		   This function is used to Verify Property

'Parameters			  :	 			sAction,sProperty,sValue,sButtons
'											    										
'Return Value		            :		True / False

'Pre-requisite			   :		  Properties Dialog Should Exist
'
'Examples				   :	
'									Call Fn_SISW_SE_PropertiesOperations("RetrieveValue_JavaList", "Custem Notes", "","Cancel")
						 
'History:
'						Developer Name				Date				Rev. No.			Changes Done					Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'					   	Sonal Padmawar			20-Mar-20103															Sonal P
'***********************************************************************************************************************************************************************************
Public Function Fn_SISW_SE_PropertiesOperations(sAction,sProperty,sValue,sButtons)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SE_PropertiesOperations"
  Dim OjProp,iCounter,aButtons, ObjJavaList , aValues, iItemsCount,iCount  
  Window("SEWindow").JavaWindow("WEmbeddedFrame").SetTOProperty "Index",1

   Select Case sAction 	

	Case "RetrieveValue_JavaList"	
				'Set OjProp = Window("SEWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Properties")
								'	Code to handle properties dialog of tracelinks tab as hierarchy is diffrent for trace links tab and traceability tab
				 If Window("SEWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Properties").Exist(2) Then
					Set OjProp = Window("SEWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Properties")
					Set ObjJavaList = OjProp.JavaList("CustomNotes")
				ElseIf JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Properties").Exist(2) Then
					Set OjProp = JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Properties")
					Set ObjJavaList = OjProp.JavaList("NotesList")
				Else 
					 Set OjProp=JavaWindow("SystemsEngineering").JavaWindow("WEmbeddedFrame").JavaDialog("Properties")
					 Set ObjJavaList = OjProp.JavaList("CustomNotes")
				End If			
				If OjProp.Exist  Then						
 ' 					    Set ObjJavaList = OjProp.JavaList("CustomNotes")
						ObjJavaList.SetTOProperty "attached text", sProperty& ":"
						 iItemsCount = ObjJavaList.GetROProperty("items count")
						 sValue = ""
						For iCounter = 0 to iItemsCount-1
								sValue = sValue & ObjJavaList.GetItem(iCounter)
								If  iCounter <  iItemsCount-1 Then
										sValue = sValue & ","
								 End If					
						Next
						Fn_SISW_SE_PropertiesOperations = sValue
				Else
					Fn_SISW_SE_PropertiesOperations = False
					Exit Function
				End If
	
   End Select

   'Click on Buttons
	 If sButtons<>"" Then
	   aButtons = split(sButtons, ":",-1,1)
	   iCounter = Ubound(aButtons)
	   For iCount=0 to iCounter
		'Click on Add Button
		Call Fn_Button_Click("Fn_SISW_SE_PropertiesOperations", OjProp, aButtons(iCount))
        Call Fn_ReadyStatusSync(2)
	   Next
	End If
   Set OjProp = Nothing

End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_SISW_SE_SourceTargetTableOperations

'Description			 :	Function Used to perform operations on Source or Target table

'Parameters			   :   1.StrAction: Action Name
'										2.StrTableName: Table Name
'										3.StrTab: Tab name
'										4.StrBOMLine: Object Name
'										5.iInstance: Object instance
'										6.StrColName: Column name
'										7.StrValue: Expected value
'										8.StrPopupMenu: Popup menu
'
'Return Value		   : 	True or False

'Pre-requisite			:	Source or Target table should be present

'Examples				:  	bReturn=Fn_SISW_SE_SourceTargetTableOperations("RowExist","Target","","000056/A;1-TopFunc (View)","","","","")
'										bReturn=Fn_SISW_SE_SourceTargetTableOperations("RowCellExist","Target","","000056/A;1-TopFunc (View)","","Item Type","Function","")
'										bReturn=Fn_SISW_SE_SourceTargetTableOperations("RowExist","source","","000055/A;1-ReqSpec1 (View)","","","","")
'										bReturn=Fn_SISW_SE_SourceTargetTableOperations("RowCellExist","source","","000055/A;1-ReqSpec1 (View)","2","Item Type","RequirementSpec","")
'
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												18-Apr-2013								1.0																						Sunny R
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_SISW_SE_SourceTargetTableOperations(StrAction,StrTableName,StrTab,StrBOMLine,iInstance,StrColName,StrValue,StrPopupMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SE_SourceTargetTableOperations"
 	'Declaring variables
	Dim bFlag,iTemp,iCounter,sColName
	Dim ObjBOMTable
	If StrTableName="" Then
		StrTableName="source"
	End If
	'Checking on which table need to perform operations
 	If lcase(StrTableName)="source" Then
		Set ObjBOMTable=JavaWindow("SystemsEngineering").JavaTable("SourceBOMTable")
		If StrTab="" Then
			StrTab="Full Results"
		End If
		sColName="BOM Line Name"
		'Selecting tab of Source table
		Call Fn_UI_JavaTab_Select("Fn_SISW_SE_SourceTargetTableOperations",JavaWindow("SystemsEngineering"),"SourceBOMTableTab", StrTab)
	Else
		Set ObjBOMTable=JavaWindow("SystemsEngineering").JavaTable("TargetBOMTable")
		If StrTab="" Then
			StrTab="Full Results"
		End If
		sColName="BOM Line"
		'Selecting tab of Target table
		Call Fn_UI_JavaTab_Select("Fn_SISW_SE_SourceTargetTableOperations",JavaWindow("SystemsEngineering"),"TargetBOMTableTab", StrTab)
	End If
	If iInstance="" Then
		iInstance=1
	End If
	Select Case StrAction
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to check specific BOM Line exist or not
		Case "RowExist"
			bFlag=False
			iTemp=1
			For iCounter=0 to ObjBOMTable.GetROProperty("rows")-1
				If Trim(StrBOMLine)=Trim(Fn_SISW_SE_SourceTargetTable_GetCellData(ObjBOMTable,iCounter,sColName)) Then
					If iInstance=iTemp Then
						bFlag=True
						Exit for
					End If
					iTemp=iTemp+1
				End If
			Next
			If bFlag=True Then
				Fn_SISW_SE_SourceTargetTableOperations=True
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: BOM Line " &StrBOMLine& " exist in table")
			Else
				Fn_SISW_SE_SourceTargetTableOperations=False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: BOM Line " &StrBOMLine& " not exist in table")
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to check specific value against BOM Line
		Case "RowCellExist"
			bFlag=False
			iTemp=1
			For iCounter=0 to ObjBOMTable.GetROProperty("rows")-1
				If Trim(StrBOMLine)=Trim(Fn_SISW_SE_SourceTargetTable_GetCellData(ObjBOMTable,iCounter,sColName)) Then
					If iInstance=iTemp Then
						If trim(Fn_SISW_SE_SourceTargetTable_GetCellData(ObjBOMTable,iCounter,StrColName))=trim(StrValue) Then
							bFlag=True
						End If
						Exit for
					End if
					iTemp=iTemp+1
				End If
			Next
			If bFlag=True Then
				Fn_SISW_SE_SourceTargetTableOperations=True
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Value " &StrValue& " exist against BMO Line " &StrBOMLine)
			Else
				Fn_SISW_SE_SourceTargetTableOperations=False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Value " &StrValue& " does not exist against BMO Line " &StrBOMLine)
			End If
	End Select
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_SISW_SE_SourceTargetTable_GetCellData

'Description			 :	Function Used to Get cell data from Target and Source table

'Parameters			   :   1.ObjTable: Table Object
'										2.iRow: Row number
'										3.iCol: Column number or column name
'
'Return Value		   : 	Data or False

'Examples				:  	Call Fn_SISW_SE_SourceTargetTable_GetCellData(JavaWindow("SystemsEngineering").JavaTable("TargetBOMTable"),1,"BOM Line")
'
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												18-Apr-2013								1.0																						Sunny R
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_SISW_SE_SourceTargetTable_GetCellData(ObjTable,iRow,iCol)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SE_SourceTargetTable_GetCellData"
	'Declaring variables
	Dim sColName,sPropName
    Fn_SISW_SE_SourceTargetTable_GetCellData=False
	If IsNumeric(iCol) Then
		sColName = ObjTable.Object.getColumn(iCol).getText()
	Else
		sColName = iCol
	End If
		Select Case trim(sColName)
			Case "BOM Line Name"
				sPropName = "bl_line_name"
			Case "BOM Line"
				sPropName = "bl_indented_title"
			Case "Item Type"
				sPropName ="bl_item_object_type"
			Case "Item Description"
				sPropName ="bl_item_object_desc"
		End Select
		Fn_SISW_SE_SourceTargetTable_GetCellData=ObjTable.Object.getItem(iRow).getData().getSelectedComponents().get(0).getProperty(sPropName)
End Function


'*********************************************************		Generic function to handle Error dialogs in SE  	***********************************************************************
'Function Name		:				Fn_SISW_SE_ErrorVerify()

'Description			 :		 		 The function is generic function to handle error dialogs. It is created after combining error dialog functions from GeneralFunctions.vbs
'										Fn_SE_DialogMsgVerify
'										Fn_SE_ErrorMessageVerify

'Parameters			   :	 			1.  dicErrorInfo
											
'Return Value		   : 				True/False

'Pre-requisite			:		 		NA.

'Examples				:				
'									Set dicErrorInfo = CreateObject("Scripting.Dictionary")
'									With dicErrorInfo	
'										.Add "Title", "Error"
'										.Add "Message", "Error..."
'										.Add "Button", "OK"
'										.Add "",
'									End with
'									bReturn = Fn_SISW_SE_ErrorVerify(dicErrorInfo)

'History:
'										Developer Name			Date			Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Sushma Pagare          10-Jun-2013
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function Fn_SISW_SE_ErrorVerify(dicErrorInfo)
			GBL_FAILED_FUNCTION_NAME="Fn_SISW_SE_ErrorVerify"
			Dim  dicKeys, dicItems, iCounter
			Dim sAction, sTitle, sErrorMsg,sButton, sAppMsg, sLine1, sLine2, sLine3
			Dim objErrorDialog
			
			dicKeys = dicErrorInfo.Keys
			dicItems = dicErrorInfo.Items
			For  iCounter=0 to dicErrorInfo.Count-1
				Select Case dicKeys(iCounter)
					Case "Action"
							sAction = dicItems(iCounter)
                    Case "Title"
							sTitle= dicItems(iCounter)
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
			Fn_SISW_SE_ErrorVerify = FALSE
            On Error Resume Next

			Select Case sAction
			
				''This covers Fn_SE_ErrorMessageVerify(sAction, sTitle, sErrorMessage, sBtnName)
				Case "VerifyErrorWindow"
					Set objErrorDialog = JavaWindow("SystemsEngineering").JavaWindow("ErrorWindow")
					objErrorDialog.SetTOProperty "title", sTitle
					If objErrorDialog.Exist(5)  Then
						if sErrorMsg <> "" then
							If instr(trim(objErrorDialog.JavaStaticText("ErrorMessage").GetROProperty("label")),sErrorMsg) > 0 then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Pass: Msg Verification Successful.")
								Fn_SISW_SE_ErrorVerify = True
							Else
								GBL_ACTUAL_MESSAGE=objErrorDialog.JavaStaticText("ErrorMessage").GetROProperty("label")
								Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Pass: Msg Verification Failed.")
								Exit Function
							end if
						Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Pass: Msg Verification Successful.")
							Fn_SISW_SE_ErrorVerify = True
						End If
						' clicking on button
						If  sButton <> "" Then
							objErrorDialog.JavaButton("Yes").SetTOProperty "label", sButton							
						End If
						call Fn_Button_Click("Fn_SISW_SE_ErrorVerify", objErrorDialog, "Yes")
					else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Function [ Fn_SISW_SE_ErrorVerify ] Error window does not exists.")	
					End If
					Set objErrorDialog = nothing
					Exit Function
	
				Case "Import Spec"
					If JavaWindow("SystemsEngineering").JavaWindow("SEWindow").JavaDialog("ImportSpecWarning").Exist Then
							 sAppMsg = JavaWindow("SystemsEngineering").JavaWindow("SEWindow").JavaDialog("ImportSpecWarning").JavaEdit("JTextArea").GetROProperty("value")
							 If instr(1,sAppMsg,sErrorMsg)<>0 Then
								Fn_SISW_SE_ErrorVerify = True
								Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Pass: Msg Verification Sucessful.")
								If  sButton <> "" Then
									JavaWindow("SystemsEngineering").JavaWindow("SEWindow").JavaDialog("ImportSpecWarning").JavaButton("OK").SetTOProperty "label", sButton							
								End If
								JavaWindow("SystemsEngineering").JavaWindow("SEWindow").JavaDialog("ImportSpecWarning").JavaButton("OK").Click micLeftBtn
							Else
								GBL_ACTUAL_MESSAGE=sAppMsg
								Fn_SISW_SE_ErrorVerify = False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Fail: Msg Verification UnSucessful.")
							End If
					Else
							Fn_SISW_SE_ErrorVerify = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Fail: Dialog Not Found")
					End If
					Exit Function					

				''This covers Fn_SE_DialogMsgVerify( sDialogTitle,sMsg,sButton)  -> Case "MSWordSaveMessageVerify"
				Case "MSWordSaveMessageVerify"
						If JavaDialog("TeamcenterMsg").Exist Then
								 sLine1 = JavaDialog("TeamcenterMsg").JavaStaticText("MsgLine3").GetROProperty("attached text") 
								 sLine2 = JavaDialog("TeamcenterMsg").JavaStaticText("MsgLine2").GetROProperty("attached text") 
								 sLine3 = JavaDialog("TeamcenterMsg").JavaStaticText("MsgLine1").GetROProperty("attached text")
								sAppMsg = sLine1+" "+sLine2+" "+sLine3
								If instr(1,sAppMsg,sErrorMsg)<>0 Then
									JavaDialog("TeamcenterMsg").JavaButton("OK").Click micLeftBtn
									Fn_SISW_SE_ErrorVerify = True
									Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Pass: Msg Verification Successful.")
								Else
									GBL_ACTUAL_MESSAGE=sAppMsg
									Fn_SISW_SE_ErrorVerify = False
									Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Fail: Msg Verification UnSucessful.")
								End If
						Else
								Fn_SISW_SE_ErrorVerify = False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Fail: Dialog Not Found")
						End If
						Exit Function
	
				'This covers Fn_SE_DialogMsgVerify( sDialogTitle,sMsg,sButton)  -> Case "MSExcelImportErrorVerify"
				Case "MSExcelImportErrorVerify"
						If Window("MicrosoftExcelWin").Dialog("ExeclImportErrorDialog").Exist Then
								sAppMsg = Window("MicrosoftExcelWin").Dialog("ExeclImportErrorDialog").Static("ErrorMsg").GetROProperty("text")
								If instr(1,sAppMsg,sErrorMsg)<>0 Then
									Window("MicrosoftExcelWin").Dialog("ExeclImportErrorDialog").WinButton("OK").Click
									Fn_SISW_SE_ErrorVerify = True
									Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Pass: Msg Verification Successful.")
								Else
									GBL_ACTUAL_MESSAGE=sAppMsg
									Fn_SISW_SE_ErrorVerify = False
									Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Fail: Msg Verification UnSucessful.")
								End If
						Else
									Fn_SISW_SE_ErrorVerify = False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Fail: Dialog Not Found")
						End If							
						Exit Function
	
					''This covers Fn_SE_DialogMsgVerify( sDialogTitle,sMsg,sButton)  -> Case "Import Spec" and Else Case
				 Case Else
							If sTitle <> "" Then
									JavaWindow("SystemsEngineering").JavaWindow("SEWindow").JavaDialog("ImportSpecWarning").SetTOProperty "title",sTitle
							End If							
							If JavaWindow("SystemsEngineering").JavaWindow("SEWindow").JavaDialog("ImportSpecWarning").Exist Then
								Set objErrorDialog = JavaWindow("SystemsEngineering").JavaWindow("SEWindow").JavaDialog("ImportSpecWarning")								
							End If						
   						    sAppMsg = objErrorDialog.JavaEdit("JTextArea").GetROProperty("value")													 		
							If instr(1,sAppMsg,sErrorMsg) <>0 Then
								If  sButton <> "" Then
									objErrorDialog.JavaButton("OK").SetTOProperty "label", sButton							
								End If
								objErrorDialog.JavaButton("OK").Click micLeftBtn
								Fn_SISW_SE_ErrorVerify = True
								Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Pass: Msg Verification Successful.")
								Set objErrorDialog =Nothing
								Exit Function
							Else
								GBL_ACTUAL_MESSAGE=sAppMsg
								Fn_SISW_SE_ErrorVerify = False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Fail: Msg Verification UnSucessful.")
							End If						 
							Set objErrorDialog =Nothing
		End Select

End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_SISW_SE_RequirementDetailsCreate

'Description			 :	Function Used to create Requirement in detail

'Parameters			   :   1.StrRequirementType: Requirement Type
'										2.dicRequirementDetailInfo: Requirement information
'										3.StrButton: Button name
'
'Return Value		   : 	Requirement Id-Revision or False

'Pre-requisite			:	Should be log in teamcenter

'Examples				:  	Dim dicRequirementDetailInfo
'									Set dicRequirementDetailInfo=CreateObject("Scripting.Dictionary")
'									dicRequirementDetailInfo("Requirement ID")="Assign"
'									dicRequirementDetailInfo("Name")="Req1"
'									dicRequirementDetailInfo("Description")="Req Desc"
'									dicRequirementDetailInfo("Unit of Measure")="gm"
'									dicRequirementDetailInfo("Next")=""
'									dicRequirementDetailInfo("Revision")="Assign"
'									dicRequirementDetailInfo("Next@1")=""
'									dicRequirementDetailInfo("AddVariables")=""
'									dicRequirementDetailInfo("VariableName")="Var1"
'									dicRequirementDetailInfo("VariableType")="Double"
'									dicRequirementDetailInfo("VariableMeasure")="Length"
'									dicRequirementDetailInfo("VariableUnit")="cm"
'									dicRequirementDetailInfo("VariableDescription")="Var desc"
'									dicRequirementDetailInfo("VariableValue")="12"
'									bReturn=Fn_SISW_SE_RequirementDetailsCreate("Design Requirement",dicRequirementDetailInfo,"Finish")
'									
'									dicRequirementDetailInfo("Requirement ID")="Assign"
'									dicRequirementDetailInfo("Name")="Req1"
'									dicRequirementDetailInfo("Description")="Req Desc"
'									dicRequirementDetailInfo("Unit of Measure")="gm"
'									dicRequirementDetailInfo("Next")=""
'									dicRequirementDetailInfo("Revision")="Assign"
'									dicRequirementDetailInfo("Requirement Formula")="KT"
'									dicRequirementDetailInfo("Severity")="1"
'									bReturn=Fn_SISW_SE_RequirementDetailsCreate("Validation Requirement",dicRequirementDetailInfo,"Finish")
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Pritam S												6-Jul-2013								1.0																					Sonal P
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_SISW_SE_RequirementDetailsCreate(StrRequirementType,dicRequirementDetailInfo,StrButton)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SE_RequirementDetailsCreate"
	Dim objDialogNewRequirement,WshShell
	Dim iItemCount,iCount,crrItem,bFlag,iCounter,sItemId,sRevId,iRowNumber,iRowNo
	Dim arrVariableName,arrType,arrMeasure,arrUnit,arrDescription,arrValue,aItmInfo3,aItmInfo4
	Dim DictItems,DictKeys
	Dim  sReq,sReqFull,sChdReqName1,strReqName
	Fn_SISW_SE_RequirementDetailsCreate=False

	Set objDialogNewRequirement = JavaWindow("SystemsEngineering").JavaWindow("NewRequirement")
	If StrRequirementType<>"" Then
		'Checking existance [ New Requirement ] dialog
		If Fn_UI_ObjectExist("Fn_SISW_SE_RequirementDetailsCreate",objDialogNewRequirement)=False Then
			'Select menu [	File -> New -> Requirement...	]
			Call Fn_MenuOperation("Select","File:New:Requirement...")
			Call Fn_ReadyStatusSync(2)
		End If
		'Selecting "Business Object" from list
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'[TC11.4_NewDevelopment_PoonamC_20July2017:Added Code if directly opens design creation page]
		If Fn_UI_ObjectExist("Fn_SISW_SE_RequirementDetailsCreate",objDialogNewRequirement.JavaTree("RequirementTree")) Then
			bFlag=False
			iItemCount=Fn_UI_Object_GetROProperty("Fn_SISW_SE_RequirementDetailsCreate",objDialogNewRequirement.JavaTree("RequirementTree"), "items count")
			For iCount=0 To iItemCount-1
				crrItem=objDialogNewRequirement.JavaTree("RequirementTree").GetItem(iCount)
				If Trim(crrItem)="Most Recently Used:"+Trim(StrRequirementType) Then
					bFlag=True
					Exit For
				ElseIf Trim(crrItem)="Complete List" Then
					Exit For
				End If
			Next
		
			If bFlag=True Then
				Call Fn_JavaTree_Select("Fn_SISW_SE_RequirementDetailsCreate", objDialogNewRequirement, "RequirementTree","Most Recently Used")
				Call Fn_JavaTree_Select("Fn_SISW_SE_RequirementDetailsCreate", objDialogNewRequirement, "RequirementTree","Most Recently Used:"+StrRequirementType)
			Else
				Call Fn_UI_JavaTree_Expand("Fn_SISW_SE_RequirementDetailsCreate", objDialogNewRequirement, "RequirementTree","Complete List")
				Call Fn_JavaTree_Select("Fn_SISW_SE_RequirementDetailsCreate", objDialogNewRequirement, "RequirementTree","Complete List")
				Call Fn_JavaTree_Select("Fn_SISW_SE_RequirementDetailsCreate", objDialogNewRequirement, "RequirementTree","Complete List:"+StrRequirementType)	
			End If
			wait 3
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			'Clicking On Next button
			objDialogNewRequirement.JavaButton("Next").WaitProperty "enabled", 1, 60000
			Call Fn_Button_Click("Fn_SISW_SE_RequirementDetailsCreate",objDialogNewRequirement, "Next")
		End If
	End If

	Set WshShell = CreateObject("WScript.Shell")
	DictItems = dicRequirementDetailInfo.Items
	DictKeys = dicRequirementDetailInfo.Keys
	For iCounter=0 to dicRequirementDetailInfo.count-1
			Select Case DictKeys(iCounter)
				' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				Case "AddVariables"
					arrVariableName=Split(dicRequirementDetailInfo("VariableName"),"~")
					For iCount = 0 to ubound(arrVariableName)
						Call Fn_Button_Click("Fn_SISW_SE_RequirementDetailsCreate",objDialogNewRequirement, "AddVariable")
						wait 1
						iRowNumber=Fn_UI_Object_GetROProperty("Fn_SISW_SE_RequirementDetailsCreate",objDialogNewRequirement.JavaTable("VariableTable"),"rows")-1
						'iRowNumber=Cint(objDialogNewRequirement.JavaTable("VariableTable").GetROProperty("rows"))-1
                        'Setting variable name
						objDialogNewRequirement.JavaTable("VariableTable").SetCellData iRowNumber,"Name",arrVariableName(iCount)
						'Selecting variable type
						If dicRequirementDetailInfo("VariableType")<>"" Then
							objDialogNewRequirement.JavaTable("VariableTable").ClickCell iRowNumber,"Type","LEFT"
							arrType=Split(dicRequirementDetailInfo("VariableType"),"~")
							If Fn_SISW_UI_JavaList_Operations("Fn_SISW_SE_RequirementDetailsCreate","Select",objDialogNewRequirement,"VariableTableList",arrType(iCount),"","")=False then
								Set WshShell = Nothing
								Set objDialogNewRequirement=Nothing
								Exit function
							End if
						End If
						'Selecting variable Measure
						If dicRequirementDetailInfo("VariableMeasure")<>"" Then
							objDialogNewRequirement.JavaTable("VariableTable").ClickCell iRowNumber,"Measure","LEFT"
							arrMeasure=Split(dicRequirementDetailInfo("VariableMeasure"),"~")
							If Fn_SISW_UI_JavaList_Operations("Fn_SISW_SE_RequirementDetailsCreate","Select",objDialogNewRequirement,"VariableTableList",arrMeasure(iCount),"","")=False then
								Set WshShell = Nothing
								Set objDialogNewRequirement=Nothing
								Exit function
							End if
						End If
						'Setting variable Unit
						If dicRequirementDetailInfo("VariableUnit")<>"" Then
							objDialogNewRequirement.JavaTable("VariableTable").ClickCell iRowNumber,"Unit","LEFT"
							arrUnit=Split(dicRequirementDetailInfo("VariableUnit"),"~")
							wait 1
							Call Fn_SISW_UI_JavaEdit_Operations("Fn_SISW_SE_RequirementDetailsCreate", "Set",  objDialogNewRequirement, "VariableTableEdit",arrUnit(iCount))
						End If
						'Setting variable Description
						If dicRequirementDetailInfo("VariableDescription")<>"" Then
							arrDescription=Split(dicRequirementDetailInfo("VariableDescription"),"~")
							objDialogNewRequirement.JavaTable("VariableTable").SetCellData iRowNumber,"Description",arrDescription(iCount)
						End If
						'Setting variable Value
						If dicRequirementDetailInfo("VariableValue")<>"" Then
							objDialogNewRequirement.JavaTable("VariableTable").DoubleClickCell iRowNumber,"Value","LEFT"
							arrValue=Split(dicRequirementDetailInfo("VariableValue"),"~")
							If objDialogNewRequirement.JavaEdit("VariableTableEdit").Exist(1) Then
								Call Fn_SISW_UI_JavaEdit_Operations("Fn_SISW_SE_RequirementDetailsCreate", "Set",  objDialogNewRequirement, "VariableTableEdit",arrValue(iCount))
							Else
								objDialogNewRequirement.JavaTable("VariableTable").ClickCell 0,0
								bFlag=false
								For iItemCount=0 to Cint(objDialogNewRequirement.JavaTable("VariableTable").GetROProperty("cols"))-1
									If trim(objDialogNewRequirement.JavaTable("VariableTable").GetColumnName(iItemCount))="Value" Then
										objDialogNewRequirement.JavaTable("VariableTable").Object.setValueAt arrValue(iCount),iRowNumber,iItemCount
										bFlag=True
										Exit for
									End If
								Next
								If bFlag=false Then
									Set WshShell = Nothing
									Set objDialogNewRequirement=Nothing
									Exit function
								End If
							End If
						End If
					Next
    			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				Case "Unit of Measure","Severity"
					bFlag=false
					objDialogNewRequirement.JavaStaticText("RequirementLabel").SetTOProperty "label",DictKeys(iCounter)+":"
					Call Fn_Button_Click("Fn_SISW_SE_RequirementDetailsCreate",objDialogNewRequirement, "LOVDropDownButton")
					wait 1
					WshShell.SendKeys "{TAB}"
					wait 1
					WshShell.SendKeys "{DOWN}"
					wait 1
					If objDialogNewRequirement.JavaWindow("TreeShell").JavaTree("Tree").Exist(1) Then
						objDialogNewRequirement.JavaWindow("TreeShell").JavaTree("Tree").Activate  DictItems(iCounter)
						wait 2
						bFlag=true
						If objDialogNewRequirement.JavaWindow("TreeShell").JavaTree("Tree").Exist(1) Then
							bFlag=False
						End If
					Else
						bFlag=False
					End If
					If bFlag=False Then
						Set WshShell = Nothing
						Set objDialogNewRequirement=Nothing
						Exit Function
					End If
				' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				'Assign or Enter ID & Revision
				Case "Revision","Requirement ID"
						If StrRequirementType = "Requirement" or StrRequirementType = "Paragraph" Then
							'Do Nothing
						Else
							objDialogNewRequirement.JavaStaticText("RequirementLabel").SetTOProperty "label",DictKeys(iCounter)+":"
							wait 1
							If LCase(DictItems(iCounter))="assign" Then
								If objDialogNewRequirement.JavaButton("Assign").GetROProperty("enabled")=1 Then
									Call Fn_Button_Click("Fn_SISW_SE_RequirementDetailsCreate",objDialogNewRequirement, "Assign")
									Call Fn_ReadyStatusSync(1)
								End If
							Else
								Call Fn_Edit_Box("Fn_SISW_SE_RequirementDetailsCreate",objDialogNewRequirement,"RequirementEdit", DictItems(iCounter))
							End If
							If DictKeys(iCounter)="Requirement ID" Then
								sItemId = Fn_Edit_Box_GetValue("Fn_SISW_SE_RequirementDetailsCreate", objDialogNewRequirement,"ID")
							Else
								sRevId = Fn_Edit_Box_GetValue("Fn_SISW_SE_RequirementDetailsCreate", objDialogNewRequirement,"Revision")
							End If
						End If
            	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				Case "Name","Description","Requirement Formula"
						objDialogNewRequirement.JavaStaticText("RequirementLabel").SetTOProperty "label",DictKeys(iCounter)+":"
						Call Fn_Edit_Box("Fn_SISW_SE_RequirementDetailsCreate",objDialogNewRequirement,"RequirementEdit", DictItems(iCounter))
    			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				'Click on Next Button
				Case "Next","Next@1","Next@2","Next@3"
						Call Fn_Button_Click("Fn_SISW_SE_RequirementDetailsCreate",objDialogNewRequirement, "Next")
			End Select
	Next
	wait(2)
	If StrButton<>"" Then
		objDialogNewRequirement.JavaButton(StrButton).WaitProperty "enabled", 1, 20000
		Call Fn_Button_Click("Fn_SISW_SE_RequirementDetailsCreate", objDialogNewRequirement, StrButton)
		Call Fn_ReadyStatusSync(1)
	End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	If sItemId<>"" Then
		If sRevId="" Then
			sRevId="A"
		End If
		Fn_SISW_SE_RequirementDetailsCreate = sItemId & "-" & sRevId
	End If

	If StrButton="Finish" Then
		If Fn_UI_ObjectExist("Fn_SISW_SE_RequirementDetailsCreate",objDialogNewRequirement)=True Then
			'Click on Close button
			Call Fn_Button_Click("Fn_SISW_SE_RequirementDetailsCreate", objDialogNewRequirement, "Cancel") 
			Call Fn_ReadyStatusSync(1)
		End If	
	End If
	If StrRequirementType = "Requirement" or StrRequirementType = "Paragraph" Then
	    Call Fn_MenuOperation("Select", "View:Expand Options:Expand")
		strReqName  =  DictItems(1)
		Call Fn_ReadyStatusSync(2)
		iRowNo = Fn_SE_BOMTable_RowIndex(strReqName)
		sReq=Fn_SE_BOMTableNodeOpeations("GetCellData",iRowNo,0,"","")
		If sReq <> "" Then
           While instr(1,sReq,":") <> 0
           	    sReq=mid(sReq,instr(1,sReq,":")+1,len(sReq))
           Wend
'			sReqFull=mid(sReq,instr(1,sReq,":")+1,len(sReq))
'			sChdReqName1=mid(sReqFull,instr(1,sReqFull,":")+1,len(sReqFull))
'			aItmInfo3 = split(sChdReqName1, "/", -1, 1)
			aItmInfo3 = split(sReq, "/", -1, 1)
			sItemId= aItmInfo3(0)
			aItmInfo4 = split(aItmInfo3(1), ";", -1, 1)
			sRevId= aItmInfo4(0)
			If sItemId<>"" Then
				If sRevId="" Then
					sRevId="A"
				End If
				Fn_SISW_SE_RequirementDetailsCreate = sItemId & "-" & sRevId
			End If
		End If
	End If
	Set WshShell = Nothing
	Set objDialogNewRequirement=Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_SISW_SE_BudgetDefinitionDetailsCreate

'Description			 :	Function Used to create Budget Definition in detail

'Parameters			   :   1.StrBudgetDefinitionType: Budget Definition Type
'										2.dicBudgetDefinitionDetailInfo: Budget Definition information
'										3.StrButton: Button name
'
'Return Value		   : 	True or False

'Pre-requisite			:	Should be log in teamcenter [ SE or RM perspective ]

'Examples				:	Dim dicBudgetDefinitionDetailInfo
'										Set dicBudgetDefinitionDetailInfo=CreateObject("Scripting.Dictionary")
'										dicBudgetDefinitionDetailInfo("Name")="B Def2"
'										dicBudgetDefinitionDetailInfo("Unit")="Test"
'										dicBudgetDefinitionDetailInfo("Roll-Up")="MAX"
'										dicBudgetDefinitionDetailInfo("Description")="Budget Desc"
'										dicBudgetDefinitionDetailInfo("Excel Template")="SE_BUDGET_TEMPLATE"
'										bReturn=Fn_SISW_SE_BudgetDefinitionDetailsCreate("Budget Definition",dicBudgetDefinitionDetailInfo,"Finish")
'
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Pritam S												7-Jul-2013								1.0																					Sonal P
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_SISW_SE_BudgetDefinitionDetailsCreate(StrBudgetDefinitionType,dicBudgetDefinitionDetailInfo,StrButton)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SE_BudgetDefinitionDetailsCreate"
	Dim objDialogNewBudgetDefinition,WshShell
	Dim iCount,bFlag,iCounter
	Dim DictItems,DictKeys
	Fn_SISW_SE_BudgetDefinitionDetailsCreate=False
	'Creating Object of [ NewBudgetDefinition ] dialog
	Set objDialogNewBudgetDefinition= JavaWindow("SystemsEngineering").JavaWindow("NewBudgetDefinition")
	If StrBudgetDefinitionType<>"" Then
		'Checking existance [ New Budget Definition ] dialog
		If objDialogNewBudgetDefinition.Exist(6)=False Then
			'Select menu [	File -> New -> Budget Definition...	]
			Call Fn_MenuOperation("Select","File:New:Budget Definition...")
			Call Fn_ReadyStatusSync(1)
		End If
		'Selecting "Budget Definition Type" from list
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		' TC112-2015070100-17_07_2015-Porting-VivekA-Added code to check existance of JavaTree("BudgetDefinitionType") as per design change
		If objDialogNewBudgetDefinition.JavaTree("BudgetDefinitionType").Exist(2) = True Then
			bFlag=False
	        Err.Clear
			bFlag = Fn_UI_JavaTree_NodeExist("Fn_SISW_SE_BudgetDefinitionDetailsCreate",objDialogNewBudgetDefinition.JavaTree("BudgetDefinitionType"),"Most Recently Used:"+StrBudgetDefinitionType)
			If bFlag = True Then
				objDialogNewBudgetDefinition.JavaTree("BudgetDefinitionType").Select  "Most Recently Used:"+StrBudgetDefinitionType
				If Err.Number < 0  Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_SISW_SE_BudgetDefinitionDetailsCreate ] Failed to Select Budget Definition Type [ "+StrBudgetDefinitionType+" ]") 
					Set objDialogNewBudgetDefinition = Nothing
					Exit Function
				End If
			End If
	
			bFlag = Fn_UI_JavaTree_NodeExist("Fn_SISW_SE_BudgetDefinitionDetailsCreate",objDialogNewBudgetDefinition.JavaTree("BudgetDefinitionType"),"Complete List:"+StrBudgetDefinitionType)
			If bFlag = True Then
				objDialogNewBudgetDefinition.JavaTree("BudgetDefinitionType").Select "Complete List:"+StrBudgetDefinitionType
				If Err.Number < 0  Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_SISW_SE_BudgetDefinitionDetailsCreate ] Failed to Select Budget Definition Type [ "+StrBudgetDefinitionType+" ]") 
					Set objDialogNewBudgetDefinition = Nothing
					Exit Function
				End If
			End If
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			'Clicking On Next button
			objDialogNewBudgetDefinition.JavaButton("Next").WaitProperty "enabled", 1, 60000
			Call Fn_Button_Click("Fn_SISW_SE_BudgetDefinitionDetailsCreate",objDialogNewBudgetDefinition, "Next")
		End If
	End If

	Set WshShell = CreateObject("WScript.Shell")
	DictItems = dicBudgetDefinitionDetailInfo.Items
	DictKeys = dicBudgetDefinitionDetailInfo.Keys
	For iCounter=0 to dicBudgetDefinitionDetailInfo.count-1
			Select Case DictKeys(iCounter)
    			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				Case "Excel Template"
					bFlag=false
					objDialogNewBudgetDefinition.JavaStaticText("BudgetDefinitionLabel").SetTOProperty "label",DictKeys(iCounter)+":"
					Call Fn_Button_Click("Fn_SISW_SE_BudgetDefinitionDetailsCreate",objDialogNewBudgetDefinition, "LOVDropDownButton")
					wait 1
					WshShell.SendKeys "{TAB}"
					wait 1
					WshShell.SendKeys "{DOWN}"
					wait 1
					If objDialogNewBudgetDefinition.JavaWindow("TreeShell").JavaTree("Tree").Exist(1) Then
						objDialogNewBudgetDefinition.JavaWindow("TreeShell").JavaTree("Tree").Activate  DictItems(iCounter)
						wait 1
						bFlag=true
						If objDialogNewBudgetDefinition.JavaWindow("TreeShell").JavaTree("Tree").Exist(1) Then
							bFlag=False
						End If
					Else
						bFlag=False
					End If
					If bFlag=False Then
						Set WshShell = Nothing
						Set objDialogNewBudgetDefinition=Nothing
						Exit Function
					End If
            	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				Case "Name","Description","Unit"
						objDialogNewBudgetDefinition.JavaStaticText("BudgetDefinitionLabel").SetTOProperty "label",DictKeys(iCounter)+":"
						Call Fn_Edit_Box("Fn_SISW_SE_BudgetDefinitionDetailsCreate",objDialogNewBudgetDefinition,"BudgetDefinitionEdit", DictItems(iCounter))
				' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				Case "Roll-Up"
					bFlag=False
					objDialogNewBudgetDefinition.JavaStaticText("BudgetDefinitionLabel").SetTOProperty "label",DictKeys(iCounter)+":"
					Call Fn_Button_Click("Fn_SISW_SE_BudgetDefinitionDetailsCreate",objDialogNewBudgetDefinition, "LOVDropDownButton")
					wait 2
					If objDialogNewBudgetDefinition.JavaWindow("TreeShell").JavaTable("LOVTable").Exist(5) Then
						For iCount=0 to Cint(objDialogNewBudgetDefinition.JavaWindow("TreeShell").JavaTable("LOVTable").GetROProperty("rows"))-1
							If objDialogNewBudgetDefinition.JavaWindow("TreeShell").JavaTable("LOVTable").GetCellData(iCount,0)=DictItems(iCounter) Then
								objDialogNewBudgetDefinition.JavaWindow("TreeShell").JavaTable("LOVTable").SelectCell iCount,0
								wait 1
								bFlag=True
								Exit for
							End If
						Next
					End If
					 If bFlag=False Then
						 Set objDialogNewBudgetDefinition=Nothing
						 Exit Function
					 End If
    			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				'Click on Next Button
				Case "Next","Next@1","Next@2","Next@3"
						Call Fn_Button_Click("Fn_SISW_SE_BudgetDefinitionDetailsCreate",objDialogNewBudgetDefinition, "Next")
			End Select
	Next
	wait 2
	If StrButton<>"" Then
		objDialogNewBudgetDefinition.JavaButton(StrButton).WaitProperty "enabled", 1, 20000
		Fn_SISW_SE_BudgetDefinitionDetailsCreate=Fn_Button_Click("Fn_SISW_SE_BudgetDefinitionDetailsCreate", objDialogNewBudgetDefinition, StrButton)
		Call Fn_ReadyStatusSync(1)
	Else
		Fn_SISW_SE_BudgetDefinitionDetailsCreate=True
	End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	If StrButton="Finish" Then
		If Fn_UI_ObjectExist("Fn_SISW_SE_BudgetDefinitionDetailsCreate",objDialogNewBudgetDefinition)=True Then
			'Click on Close button
			Call Fn_Button_Click("Fn_SISW_SE_BudgetDefinitionDetailsCreate", objDialogNewBudgetDefinition, "Cancel") 
			Call Fn_ReadyStatusSync(1)
		End If	
	End If
	'Releasing Objects
	Set WshShell = Nothing
	Set objDialogNewBudgetDefinition=Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_SISW_SE_EditBudgetOperations

'Description			 :	Function Used to perform operations on Edit Budget

'Parameters			   :   1.StrAction: Action Name
'										2.StrBudgetDefination: Budget Definition name
'										3.StrExcelTemplate: Excel Template name
'										4.StrButton: Button name
'
'Return Value		   : 	True or False

'Pre-requisite			:	Edit Budget Dialog should be appear

'Examples				:	bReturn=Fn_SISW_SE_EditBudgetOperations("Edit","B Def1","SE_BUDGET_TEMPLATE","")
'
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Pritam S												7-Jul-2013								1.0																					Sonal P
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_SISW_SE_EditBudgetOperations(StrAction,StrBudgetDefination,StrExcelTemplate,StrButton)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SE_EditBudgetOperations"
 	'Declaring variables
	Dim objEditBudget

	Fn_SISW_SE_EditBudgetOperations=False
	'Checking Existance of [ Edit Budget ] dialog
 	If not JavaWindow("SystemsEngineering").JavaWindow("EditBudget").Exist(6) Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: [ Edit Budget ] dialog not appear") 
		Exit function
	End If
	'Creating object of [ Edit Budget ] dialog
	Set objEditBudget=JavaWindow("SystemsEngineering").JavaWindow("EditBudget")
	Select Case StrAction
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to Edit Budget
		Case "Edit"
			'Selecting Budget Defination From [ Budget Defination ] list
			If Fn_SISW_UI_JavaList_Operations("Fn_SISW_SE_EditBudgetOperations", "Exist", objEditBudget,"BudgetDefination",StrBudgetDefination, "", "") then
				Call Fn_SISW_UI_JavaList_Operations("Fn_SISW_SE_EditBudgetOperations", "Select", objEditBudget,"BudgetDefination",StrBudgetDefination, "", "")
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Budget Defination [ "+StrBudgetDefination+" ] not found in [ Budget Defination ] list")
				Set objEditBudget=Nothing
				Exit Function
			End if
			If StrExcelTemplate<>"" Then
				'Selecting Excel Template From [ Excel Template ] list
				If Fn_SISW_UI_JavaList_Operations("Fn_SISW_SE_EditBudgetOperations", "Exist", objEditBudget,"ExcelTemplate",StrExcelTemplate, "", "") then
					Call Fn_SISW_UI_JavaList_Operations("Fn_SISW_SE_EditBudgetOperations", "Select", objEditBudget,"ExcelTemplate",StrExcelTemplate, "", "")
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Excel Template [ "+StrExcelTemplate+" ] not found in [ Excel Template ] list")
					Set objEditBudget=Nothing
					Exit Function
				End if
			End If
			Fn_SISW_SE_EditBudgetOperations=Fn_Button_Click("Fn_SISW_SE_EditBudgetOperations",objEditBudget, "OK")
			Call Fn_ReadyStatusSync(1)
	End Select
	'Releasing object of [ Edit Budget ] dialog
	Set objEditBudget=Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_SISW_SE_BudgetsTableOperations

'Description			 :	Function Used to perform operations on Budgets Table

'Parameters			   :   1.StrAction: Action Name
'										2.StrName: Budget name
'										3.StrColumn: Column name
'										4.StrValue: Expected value
'										5.StrBudgetDefination: Budget Defination name
'										6.StrExcelTemplate: Excel Template name
' 
'Return Value		   : 	True or False

'Pre-requisite			:	 Budgets table should be appear

'Examples				:	bReturn=Fn_SISW_SE_BudgetsTableOperations("Select","B Def1","","","","")
'										bReturn=Fn_SISW_SE_BudgetsTableOperations("VerifyCellExist","B Def1","Expression","MAX","","")
'										bReturn=Fn_SISW_SE_BudgetsTableOperations("Edit","B Def1","","","b1","SE_BUDGET_TEMPLATE")
'										bReturn=Fn_SISW_SE_BudgetsTableOperations("Remove","b1","","","","")
'										bReturn=Fn_SISW_SE_BudgetsTableOperations("CellListSelect","B1","Expression","MAX","","")
'
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Pritam S														8-Jul-2013								1.0																						  Sonal P
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sonal P														13-Aug-2013								1.1					Added Case : CellListSelect							  Pritam S
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_SISW_SE_BudgetsTableOperations(StrAction,StrName,StrColumn,StrValue,StrBudgetDefination,StrExcelTemplate)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SE_BudgetsTableOperations"
 	'Declaring variables
	Dim objBudgetsTable
	Dim iCounter,iCount,arrName,arrValue,bFlag
	Fn_SISW_SE_BudgetsTableOperations=False
	'Checking Existance of [ Budgets ] table
	If JavaWindow("SystemsEngineering").JavaTable("BudgetsTable").Exist(5) Then
		'Do Nothing
	Elseif Fn_SISW_UI_RACTabFolderWidget_Operation("Select","Budgets","") Then
		'Checking Existance of [ Budgets ] table
		If not JavaWindow("SystemsEngineering").JavaTable("BudgetsTable").Exist(5) Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: [ Budgets ] Table not found")
			Exit Function
		End If
	Else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: [ Budgets ] Table not found")
		Exit Function
	End if
	'Creating object of [ Budgets ] table
	Set objBudgetsTable=JavaWindow("SystemsEngineering").JavaTable("BudgetsTable")
	Select Case StrAction
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to select [ Budget ] from table
		Case "Select"
			For iCounter=0 to Cint(objBudgetsTable.GetROProperty("rows"))-1
				If Trim(StrName)=Trim(objBudgetsTable.GetCellData(iCounter,"Name")) Then
					objBudgetsTable.SelectCell iCounter,"Name"
					wait 1
					Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Successfully selected Budget [ "+StrName+" ] from [ Budgets ] Table")
					Fn_SISW_SE_BudgetsTableOperations=True
					Exit For
				End If
			Next
			If Fn_SISW_SE_BudgetsTableOperations=False Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Budget [ "+StrName+" ] not found in [ Budgets ] Table")
			End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to verify specific value Exist against specific Budget in [ Budgets ] table
		Case "VerifyCellExist"
			arrName=Split(StrName,"~")
			arrValue=Split(StrValue,"~")
			For iCounter=0 to Ubound(arrName)
				bFlag=False
				For iCount=0 to Cint(objBudgetsTable.GetROProperty("rows"))-1
					If Trim(arrName(iCounter))=Trim(objBudgetsTable.GetCellData(iCount,"Name")) Then
						If Trim(arrValue(iCounter))=Trim(objBudgetsTable.GetCellData(iCount,StrColumn)) Then
							bFlag=True
							Exit For
						End If	
					End If
				Next
				If bFlag=False Then
				End If
			Next
			If bFlag=True Then
				Fn_SISW_SE_BudgetsTableOperations=True
			End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to Edit Budget from [ Budgets ] table
		Case "Edit"
			'Selecting Budget from table
			If Fn_SISW_SE_BudgetsTableOperations("Select",StrName,"","","","") Then
				'Clicking on [ Edit ] Button
				Call Fn_Button_Click("Fn_SISW_SE_BudgetsTableOperations",JavaWindow("SystemsEngineering"), "EditBudget")
				Fn_SISW_SE_BudgetsTableOperations=Fn_SISW_SE_EditBudgetOperations("Edit",StrBudgetDefination,StrExcelTemplate,"")
			End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to Remove Budget from [ Budgets ] table
		Case "Remove"
			If Fn_SISW_SE_BudgetsTableOperations("Select",StrName,"","","","") Then
				Call Fn_Button_Click("Fn_SISW_SE_BudgetsTableOperations",JavaWindow("SystemsEngineering"), "RemoveBudget")
				'Checking existance of [ Confirm ] dialog
				If JavaWindow("SystemsEngineering").JavaWindow("Confirm").Exist(5) Then
					Fn_SISW_SE_BudgetsTableOperations=Fn_Button_Click("Fn_SISW_SE_BudgetsTableOperations",JavaWindow("SystemsEngineering").JavaWindow("Confirm"), "Yes")
					Call Fn_ReadyStatusSync(1)
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Fail to Remove Budget [ "+StrName+" ]")
				End If
			End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "CellListSelect"
			For iCounter=0 to Cint(objBudgetsTable.GetROProperty("rows"))-1
				If Trim(StrName)=Trim(objBudgetsTable.GetCellData(iCounter,"Name")) Then
					objBudgetsTable.SelectCell iCounter,StrColumn
					For iCount = 0 to 5
						objBudgetsTable.SelectCell 0,StrColumn
						wait 1
						If JavaWindow("SystemsEngineering").JavaList("BudgetsTableList").Exist(5) Then
							Exit For
						End If
					Next
					If JavaWindow("SystemsEngineering").JavaList("BudgetsTableList").Exist(5) Then
						Fn_SISW_SE_BudgetsTableOperations=Fn_SISW_UI_JavaList_Operations("Fn_SISW_SE_BudgetsTableOperations", "Select",JavaWindow("SystemsEngineering"),"BudgetsTableList",StrValue, "", "")
						wait 1
						Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Successfully selected value [ "+StrValue+" ] from list for Budget [ "+StrName+" ] from [ Budgets ] Table")
						objBudgetsTable.SelectCell iCounter,"Name"
						wait 1
					End If		
					Exit For
				End If
			Next
			If Fn_SISW_SE_BudgetsTableOperations=False Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Fail to select value [ "+StrValue+" ] from list for Budget [ "+StrName+" ] from [ Budgets ] Table")
			End If
	End Select
	'Releasing Object
	Set objBudgetsTable=Nothing
End Function


'/$$$$   FUNCTION NAME   :  Fn_SE_OpenDiagram(sAction, sDiagName)
'/$$$$
'/$$$$   DESCRIPTION        :  This function will  perform Operations on the Create New Diagram Dialog
'/$$$$ 
'/$$$$   PRE-REQUISITES        :  The Document mapping tree should be present
'/$$$$
'/$$$$  PARAMETERS   : 		sAction : Action to be performed Open or Verify
'/$$$$								sDiagName:Name of the diagram to open or Verify
'/$$$$	
'/$$$$		Return Value : 				True or False
'/$$$$
'/$$$$    Function Calls       :   Fn_WriteLogFile(), Fn_MenuOperation(), Fn_ReadyStatusSync
'/$$$$										
'/$$$$
'/$$$$		HISTORY           :   		AUTHOR                 DATE        VERSION
'/$$$$
'/$$$$    CREATED BY     :   Archana          08/12/2013         1.0
'/$$$$
'/$$$$   	
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

Function Fn_SISW_SE_OpenDiagram(sAction, sDiagName)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SE_OpenDiagram"
Dim objDiagram,bReturn,iIndex
Fn_SISW_SE_OpenDiagram=false

		Set objDiagram= JavaWindow("SystemsEngineering").JavaWindow("SelectDiagramtoOpen")
		If objDiagram.Exist(5)=false Then
			bReturn=Fn_MenuOperation("Select","File:Open Diagram")
			If bReturn=false Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SE_OpenDiagram ] New Diagram Dialog Not invoked")
						 Fn_SISW_SE_OpenDiagram = False
						 Exit function
			Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SISW_SE_OpenDiagram ] New Open Diagram Dialog Successfully invoked")	
						Call Fn_ReadyStatusSync(2)
			End If
		End If

		Select Case sAction
			Case "Open"
						'Select the given diagrm from list.								
					Err.clear	
					objDiagram.JavaList("DiagramList").Select sDiagName
					If Err.Number < 0 Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SE_OpenDiagram ] Unabe to select ["+ sDiagName +"] from Diagram list")
							 Fn_SISW_SE_OpenDiagram = False
							 Exit function									
					End If	
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SISW_SE_OpenDiagram ] Sucessfully select ["+ sDiagName +"] from Diagram list")
					Call Fn_ReadyStatusSync(2)
					objDiagram.JavaButton("OK").Click micLeftBtn
					If Err.Number < 0 Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SE_OpenDiagram ] Unabe to Click Ok Button")
							 Fn_SISW_SE_OpenDiagram = False
							 Exit function									
					End If	
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SISW_SE_OpenDiagram ] Sucessfully Click OK Button")	

              Case "Verify"
					Err.clear	
					iIndex = objDiagram.JavaList("DiagramList").GetItemIndex(sDiagName)
					If iIndex > 0 Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SISW_SE_OpenDiagram ] Sucessfully verified that ["+ sDiagName +"] is listed in Diagram list")	
					End If
					If Err.Number < 0 Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SE_OpenDiagram ] Unabe to verify exsitance of ["+ sDiagName +"] from Diagram list")
						Fn_SISW_SE_OpenDiagram = False
						Exit function									
					End If	
					objDiagram.JavaButton("OK").Click micLeftBtn
					If Err.Number < 0 Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SE_OpenDiagram ] Unabe to Click Ok Button")
							 Fn_SISW_SE_OpenDiagram = False
							 Exit function									
					End If	
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SISW_SE_OpenDiagram ] Sucessfully Click OK Button")
					
		End Select
Fn_SISW_SE_OpenDiagram = True
Set objDiagram= Nothing
End Function
 
'/$$$$   FUNCTION NAME   :  Fn_SISW_SE_ShapeDeleteConfirmation(sButtonName)
'/$$$$
'/$$$$   DESCRIPTION        :  This function will  perform Operations on the Confirmation Dialog
'/$$$$ 
'/$$$$   PRE-REQUISITES        :  Dialog exist
'/$$$$
'/$$$$  PARAMETERS   : 		sButtonName : Name of the button to hit
'/$$$$								
'/$$$$		Return Value : 				True or False
'/$$$$
'/$$$$    Function Calls       :   Fn_WriteLogFile(), Fn_MenuOperation(), Fn_ReadyStatusSync
'/$$$$										
'/$$$$
'/$$$$		HISTORY           :   		AUTHOR                 DATE        VERSION
'/$$$$
'/$$$$    CREATED BY     :   Archana          08/27/2013         1.0
'/$$$$
'/$$$$   	
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Public Function Fn_SISW_SE_ShapeDeleteConfirmation(sButtonName)
GBL_FAILED_FUNCTION_NAME="Fn_SISW_SE_ShapeDeleteConfirmation"
Dim objDialog

Set objDialog =  Fn_SISW_SE_GetObject("Confirmation")
Fn_SISW_SE_ShapeDeleteConfirmation = False

  Err.Clear
  objDialog.JavaButton(sButtonName).Click
	
	If Err. Number < 0  Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SE_ShapeDeleteConfirmation ] Unabe to Click ["+sButtonName+"] Button")		
	Else
			Fn_SISW_SE_ShapeDeleteConfirmation = TRUE
	End If
Set objDialog =  Nothing

End Function

''*********************************************************		Function to Perform SE Trace Link Panel operation in System Engineering	***********************************************************************
'
''Function Name		:		Fn_SISW_SE_TraceLinkTabOperations
'
''Description			 :		 This function is used to get the SE Table Node Index.
'
''Parameters		:		   1.	sAction = Action To Perform
'							   			2.  sTabName = "Trace Link",  pass this value only if you want select "tracelink" if tab already selected pass ""
'										3. bMaximiseTab = "Yes", Pass this value only if tracelink tab is not maximised, otherwise it will restore tab
'										4. sColumnConfig = Column Configuration Value to Select
'										5. dicTraceLinkObjects = Dictionary Object
'														Containg info of Trees [Defining Objects / Complying Objects]
'										6. sReserve = For Future Use
'										7. StrMenu = Popup menu

'			  										
''Return Value		   : 				True/ False
'
''Pre-requisite			:				Trace Link Panel should be displayed .
'
'
''Examples				:		  	Dim dicTraceLinkObjects
'											Set dicTraceLinkObjects=CreateObject("Scripting.Dictionary")
'											dicTraceLinkObjects("Defining Objects")=""
'											dicTraceLinkObjects("Complying Objects")=""
'											dicTraceLinkObjects("ColumnNames")=""
'											dicTraceLinkObjects("ColumnValues")=""
'											dicTraceLinkObjects("TreeName") = ""
'											dicTraceLinkObjects("ExpandLevel") = ""

'											Case "Select", "Expand"
		'										dicTraceLinkObjects("Defining Objects")="REQ-000014/A;1-Description:Description->Rationale:REQ-000015/A;1-Rationale"
		'										Fn_SISW_SE_TraceLinkTabOperations("Select","Trace Links", "Yes", "", dicTraceLinkObjects,"","")
		
'											Case "PopupMenuSelect"
		'										dicTraceLinkObjects("Defining Objects")="REQ-000014/A;1-Description:Description->Rationale:REQ-000015/A;1-Rationale"
		'										Fn_SISW_SE_TraceLinkTabOperations("Select","Trace Links", "Yes", "", dicTraceLinkObjects,"","Properties")

'											Case "Verify"
		'										dicTraceLinkObjects("Defining Objects")="REQ-000014/A;1-Description:Description->Rationale:REQ-000015/A;1-Rationale"
'												dicTraceLinkObjects("ColumnNames")="Type~Relation Type"
'												dicTraceLinkObjects("ColumnValues")="Requirement Revision~Trace Link"
		'										Fn_SISW_SE_TraceLinkTabOperations("Select","Trace Links", "Yes", "", dicTraceLinkObjects,"","")

'											Case  "GetColumnList", "GetColCount", "ColumnHeaderPopupMenuSelect"
		'										dicTraceLinkObjects("TreeName")="Complying Objects"
		'										Fn_SISW_SE_TraceLinkTabOperations("GetColumnList","", "", "", dicTraceLinkObjects,"","")
		
'											Case  "VerifyExpandLevel" - Note:If we want to verify Expandlevel 2, then pass tree hierarchy of expanded (Node1&node2)node and also 3rd (Node3)node which not to be expanded.
		'										dicTraceLinkObjects("Complying Objects")="Node1:Node2:Node3"
'												dicTraceLinkObjects("ExpandLevel") = "2"
		'										Fn_SISW_SE_TraceLinkTabOperations("VerifyExpandLevel","", "", "", dicTraceLinkObjects,"","")

'                                            Case "VerifySortColumnValue"	- Verify Column Value against node --- Added By Jotiba NewDev
		'											dicTraceLinkObjects("TreeName")="Complying Objects"
		'											dicTraceLinkObjects("Root Node")="REQ-936381/A;1-ReqNew"
		'											dicTraceLinkObjects("Node Names")="000404/A;1-bb~000403/A;1-aa"
		'											dicTraceLinkObjects("ColumnNames")="Type~Owner:Type"
		'											dicTraceLinkObjects("ColumnValues")="Requirement Specification Revision~AutoTest2 (autotest2):Requirement Specification Revision"


'History:
'										Developer Name				Date						Rev. No.							Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Pranav Ingle					3-Mar-2014					1.0									
'										Jotiba T						1-Jun-17					1.0						added case "IsEditable"         Shweta Rathod
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_SE_TraceLinkTabOperations(sAction, sTabName, bMaximiseTab, sColumnConfig, dicTraceLinkObjects, sReserve,StrMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SE_TraceLinkTabOperations"
	On error resume next
	'Declare Variables
	Dim DictItems,DictKeys, ObjTraceLinkTree, ObjTree, aNodePath, bResult
	Dim iCounter,iCount, sValue, aColNames, aColValues, iOccurence, icnt, aOccurence
	Dim aMenuList, intCount, sColName,ObjTraceLinkCriteria, dicProperties
	Dim aColumnName,ObjTraceLinktabTree,icnt1, aNode, iOuter, aSubColName,aSubColValues

	Fn_SISW_SE_TraceLinkTabOperations=False
	If JavaWindow("SystemsEngineering").JavaWindow("ErrorWindow").Exist(2) Then
		Call Fn_SISW_UI_JavaButton_Operations("Fn_SISW_SE_TraceLinkTabOperations", "Click", JavaWindow("SystemsEngineering").JavaWindow("ErrorWindow"),"OK")
		Call Fn_ReadyStatusSync(1)
	End If
	
	'Select Tracelink tab in right side panel
	If sTabName<> "" Then
		Call Fn_SetView ("Other:Trace Links")
        	Call Fn_ReadyStatusSync(1)
		If JavaWindow("SystemsEngineering").JavaWindow("ErrorWindow").Exist(2) Then
				Call Fn_SISW_UI_JavaButton_Operations("Fn_SISW_SE_TraceLinkTabOperations", "Click", JavaWindow("SystemsEngineering").JavaWindow("ErrorWindow"),"OK")
		End If
		If sAction <> "VerifyWthtShowTraceLinks" Then
			If Fn_ToolBarOperation("ButtonExist", "Show Trace Link", "" ) Then					'Tc1015-19_08_2015-AnkitN-Development-Checked Existance of Show Trace Link
				Call Fn_ToolBarOperation("Click", "Show Trace Link", "")
				Call Fn_ReadyStatusSync(1)
			End If
		Else
			If Fn_ToolBarOperation("ButtonExist", "Hide Trace Link", "" ) Then					'[TC1015-2015091500-29_09_2015-VivekA-Development]-Added by Snehal to Check Existance of Hide Trace Link
				Call Fn_ToolBarOperation("Click", "Hide Trace Link", "")
				Call Fn_ReadyStatusSync(1)
			End If
		End IF
	End If
	
	'Maximize the tab if provided Yes
	If bMaximiseTab = "Yes" Then
		Call Fn_SE_RightPanelTabOperations("DoubleClick", "Trace Links", "")
		Call Fn_ReadyStatusSync(1)
	End If

	If sColumnConfig <> "" Then
		JavaWindow("SystemsEngineering").JavaList("Column Configurations").Select sColumnConfig
		Wait 1
	End If
	
	If sAction <> "VerifyTraceLinkCriteriaWindow" Then
		'Taking Items & Keys from dictionary
		DictItems = dicTraceLinkObjects.Items
		DictKeys = dicTraceLinkObjects.Keys
	End IF
	
   	Select Case sAction
		Case "Select","Expand","Verify","PopupMenuSelect","NodeVerify","Properties","VerifyProperties","DescriptionProperties","VerifyDescriptionProperties","Delete Trace Link","VerifyExpandLevel","ExportToExcel","GoToObject","FindInView","FindInViewNodeExist","VerifyWthtShowTraceLinks","Collapse","FindInViewNodeEnabled","IsEditable"
			For iCounter=0 to dicTraceLinkObjects.count-1
					Select Case DictKeys(iCounter)
	 '----------------------------------------------------------------------------------------------------------------------------------------
						Case "Complying Objects", "Defining Objects"
							' Set Object of Tree in Variable
							Set ObjTraceLinkTree = JavaWindow("SystemsEngineering").JavaTree(DictKeys(iCounter)).Object
							Set ObjTree = JavaWindow("SystemsEngineering").JavaTree(DictKeys(iCounter))
							' Split Node values to get path
							arrNodes = Split(DictItems(iCounter), ":")
							For itemCount = 0 To UBound(arrNodes)
								bResult = False
								iOccurence=1
								icnt = 1

								' Get Rowcount for perticular level of tree
								iRowCount = ObjTraceLinkTree.getItemCount
								If instr(arrNodes(itemCount),"@") Then
									aOccurence = Split(arrNodes(itemCount),"@")
									arrNodes(itemCount) = aOccurence(0)
									iOccurence = cint(aOccurence(1))
								End If

								For iCount = 0 To iRowCount - 1
									' Get node name from tree
									sNodeName = ObjTraceLinkTree.getItem(iCount).getData().getDisplayedObject().toString()
									If arrNodes(itemCount) = sNodeName Then
										If icnt=iOccurence Then
											Set ObjTraceLinkTree = ObjTraceLinkTree.getItem(iCount)
											If itemCount = 0 Then
												sNodePath = "#"& iCount
											Else
												sNodePath = sNodePath &":#"& iCount
											End If
											bResult = True
											Exit For	
										Else
											icnt = icnt + 1
										End If
									End If
								Next
								'Exit Function if perticular node is not found
								If bResult = False Then
									Exit Function
								End If
							Next
						 '----------------------------------------------------------------------------------------------------------------------------------------
							If sAction="Select" Then						'Select Tracelink node
								ObjTree.Select sNodePath
						 '----------------------------------------------------------------------------------------------------------------------------------------		
							ElseIf sAction = "NodeVerify" Then			'Verify tracelink node
								If bResult = False Then
									Exit Function
								End If
						 '----------------------------------------------------------------------------------------------------------------------------------------			
							ElseIf sAction="Expand" Then				'Expand Tracelink node
								ObjTree.Expand sNodePath
							ElseIf sAction="Collapse" Then		'[TC1122-20160203-15_02_2016-VivekA-Maintenance] - Added from TC10.1.5
								ObjTree.Collapse sNodePath
						 '----------------------------------------------------------------------------------------------------------------------------------------		
							ElseIf sAction="Verify" OR sAction="CellVerify" OR sAction="VerifyWthtShowTraceLinks" Then
								aColNames = Split(dicTraceLinkObjects("ColumnNames"), "~")
								aColValues = Split(dicTraceLinkObjects("ColumnValues"), "~")
								For iCount = 0 To UBound(aColNames)
									bResult = False
									If aColNames(iCount)="Defining Context" Then
										sValue = ObjTraceLinkTree.getData().getSourceContextString()
									ElseIf aColNames(iCount)="Complying Context" Then
										sValue = ObjTraceLinkTree.getData().getTargetContextString()
'									ElseIf aColNames(iCount)="Type" Then
'										sValue = ObjTraceLinkTree.getData().getDisplayedObject().getProperty("Type")
									Else
										aColNames(iCount)  = Fn_GetXMLNodeValue(Environment.Value("sPath")+ "\TestData\AutomationXML\ObjectRealNames\SystemsEngineering.xml" , Trim(aColNames(iCount)))
										sValue = ObjTraceLinkTree.getData().getDisplayedObject().getProperty(aColNames(iCount))
									End If
									If Instr(1, sValue,aColValues(iCount))> 0  Then
										bResult=true
									End If
									If bResult = False Then
										Exit Function
									End If
								Next
						 '----------------------------------------------------------------------------------------------------------------------------------------			
							ElseIf sAction = "PopupMenuSelect"  Then	'Pop-up menu Select
								If StrMenu="Create Custom Note" Then
									StrMenu="Custom Note"
								ElseIf StrMenu="Properties..." Then
									StrMenu="View Properties	Alt+P"
								End If
								wait 1
								JavaWindow("SystemsEngineering").JavaTree(DictKeys(iCounter)).Select sNodePath
								wait 1
								Call Fn_UI_JavaTree_OpenContextMenu("Fn_SISW_SE_TraceLinkTabOperations",JavaWindow("SystemsEngineering"), DictKeys(iCounter),sNodePath)
								Wait 2
								'Select Menu action
								aMenuList = split(StrMenu, ":",-1,1)
								intCount = Ubound(aMenuList)
								Select Case intCount
									Case "0"
										 StrMenu =JavaWindow("SystemsEngineering").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0))
									Case "1"
										StrMenu =JavaWindow("SystemsEngineering").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1))
									Case "2"
										StrMenu =JavaWindow("SystemsEngineering").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1),aMenuList(2))
									Case Else
										Exit Function
								End Select
								JavaWindow("SystemsEngineering").WinMenu("ContextMenu").Select StrMenu
								wait 3
								If StrMenu = "Go To Object" OR StrMenu = "Expand All" OR StrMenu = "Expand Below" Then
									If JavaWindow("SystemsEngineering").JavaWindow("ErrorWindow").Exist(5) Then
											Call Fn_SISW_UI_JavaButton_Operations("Fn_SISW_SE_TraceLinkTabOperations", "Click", JavaWindow("SystemsEngineering").JavaWindow("ErrorWindow"),"OK")
									End If
								End If
						 '----------------------------------------------------------------------------------------------------------------------------------------	
							ElseIf sAction="Properties" or sAction="DescriptionProperties" Then
									JavaWindow("SystemsEngineering").JavaTree(DictKeys(iCounter)).Select sNodePath	
									wait 1
									Call Fn_ToolBarOperation("ShowDropdownAndSelectWithInstance", "View Menu:3", "Properties	Alt+P")
									Call Fn_ReadyStatusSync(1)
									If dicTraceLinkObjects("PropertyValue") <>"" Then
										Set dicProperties = CreateObject( "Scripting.Dictionary" )
										If sAction="Properties" Then
											dicProperties("PropertyName")="Name"
										Else
											dicProperties("PropertyName")="Description"
										End If
										dicProperties("Value")=dicTraceLinkObjects("PropertyValue")
										If Fn_SISW_VerifyProperties("ModifyEditBox","Show empty properties...",dicProperties,"OK") = False Then
											Exit Function
										End If 
									End If
						 '----------------------------------------------------------------------------------------------------------------------------------------			
							ElseIf sAction="VerifyProperties" or sAction="VerifyDescriptionProperties" Then
									JavaWindow("SystemsEngineering").JavaTree(DictKeys(iCounter)).Select sNodePath	
									wait 1
									Call Fn_ToolBarOperation("ShowDropdownAndSelectWithInstance", "View Menu:3", "Properties	Alt+P")
									Call Fn_ReadyStatusSync(1)
									If dicTraceLinkObjects("PropertyValue") <>"" Then
										Set dicProperties = CreateObject( "Scripting.Dictionary" )
										If sAction="VerifyProperties" Then
											dicProperties("PropertyName")="Name"
										Else
											dicProperties("PropertyName")="Description"
										End If
										dicProperties("Value")=dicTraceLinkObjects("PropertyValue")
										If Fn_SISW_VerifyProperties("EditBox","",dicProperties,"Cancel") = False Then
											Exit Function
										End If 
									End If
						 '----------------------------------------------------------------------------------------------------------------------------------------			
							ElseIf sAction = "GoToObject" or sAction = "FindInView" Then
							    	JavaWindow("SystemsEngineering").Click 150,15
								wait 1
								JavaWindow("SystemsEngineering").JavaTree(DictKeys(iCounter)).Select sNodePath	
								wait 1
								'As discussed with dhananjay the buttons are not available so changed to Rmb operation  -Pratap 
								Call Fn_UI_JavaTree_OpenContextMenu("Fn_SISW_SE_TraceLinkTabOperations",JavaWindow("SystemsEngineering"), DictKeys(iCounter),sNodePath)
								If sAction = "GoToObject" Then
									'Call Fn_SISW_UI_JavaToolbar_Operations("Fn_SISW_SE_TraceLinkTabOperations","ClickPackBtn",JavaWindow("SystemsEngineering"),"","Go To Object", "",StrMenu,"")
								     StrMenu="Go To Object"
								ElseIf sAction = "FindInView" Then
									'Call Fn_SISW_UI_JavaToolbar_Operations("Fn_SISW_SE_TraceLinkTabOperations","ClickPackBtn",JavaWindow("SystemsEngineering"),"","Find In", "",StrMenu,"")								
								    StrMenu= "Find in All Visible Views"
								End If
								wait 2
								StrMenu =JavaWindow("SystemsEngineering").WinMenu("ContextMenu").BuildMenuPath(StrMenu)
								JavaWindow("SystemsEngineering").WinMenu("ContextMenu").Select StrMenu
						 '----------------------------------------------------------------------------------------------------------------------------------------			
						 	ElseIf sAction = "FindInViewNodeExist" OR sAction = "FindInViewNodeEnabled" Then				
								JavaWindow("SystemsEngineering").Click 150,15
						    	wait 1
								JavaWindow("SystemsEngineering").JavaTree(DictKeys(iCounter)).Select sNodePath	
								wait 1
								If sAction = "FindInViewNodeExist" Then
									Call Fn_SISW_UI_JavaToolbar_Operations("Fn_SISW_SE_TraceLinkTabOperations","ClickPackBtn",JavaWindow("SystemsEngineering"),"","Find In", "","","")
									Wait 2
									Fn_SISW_SE_TraceLinkTabOperations = Fn_MenuOperation("Exist",StrMenu)
									Exit Function
								ElseIf sAction = "FindInViewNodeEnabled" Then
									bFlag = Fn_SISW_UI_JavaToolbar_Operations("Fn_SISW_SE_TraceLinkTabOperations","ClickPackBtn",JavaWindow("SystemsEngineering"),"","Find In", "","","")
									wait 2
									StrMenu = JavaWindow("SystemsEngineering").WinMenu("ContextMenu").BuildMenuPath(StrMenu)
									Fn_SISW_SE_TraceLinkTabOperations = JavaWindow("SystemsEngineering").WinMenu("ContextMenu").CheckItemProperty(StrMenu,"Enabled","1")
									Wait 1
									Call Fn_KeyBoardOperation("SendKeys", "{ESC}")
									Exit Function
								End If
						'----------------------------------------------------------------------------------------------------------------------------------------			
							ElseIf sAction="Delete Trace Link" Then
								JavaWindow("SystemsEngineering").JavaTree(DictKeys(iCounter)).Select sNodePath	
								wait 1
								Call Fn_ToolBarOperation("ShowDropdownAndSelectWithInstance", "View Menu:3", "Delete Trace Link	Delete")
								If Fn_SISW_UI_Object_Operations("Fn_SISW_SE_TraceLinkTabOperations","Exist", JavaWindow("SystemsEngineering").JavaWindow("ErrorWindow"),"") = True Then
									Call Fn_SISW_UI_JavaButton_Operations("Fn_SISW_SE_TraceLinkTabOperations", "Click",  JavaWindow("SystemsEngineering").JavaWindow("ErrorWindow"),"OK")
								End If
								If Fn_SISW_UI_Object_Operations("Fn_SISW_SE_TraceLinkTabOperations","Exist",  JavaWindow("DefaultWindow").JavaWindow("Refresh Window"),"") = True Then
									Call Fn_SISW_UI_JavaButton_Operations("Fn_SISW_SE_TraceLinkTabOperations", "Click",  JavaWindow("DefaultWindow").JavaWindow("Refresh Window"),"OK")
								End If
						 '----------------------------------------------------------------------------------------------------------------------------------------		
							ElseIf sAction="VerifyExpandLevel" Then
								aNodePath = Split(Replace(sNodePath,"#",""),":")
								For icnt=0 To Ubound(aNodePath)
									If icnt = 0 Then
										Set ObjTraceLinkTree = ObjTree.Object.getItem(aNodePath(icnt))
										sValue = ObjTraceLinkTree.getExpanded()
									Else
										Set ObjTraceLinkTree = ObjTraceLinkTree.getItem(aNodePath(icnt))
										If ObjTraceLinkTree.getItemCount() = 0 Then
											sValue = sValue+":"+"false"
										Else
											sValue = sValue+":"+ObjTraceLinkTree.getExpanded()
										End If
									End If
								Next
								aValue = Split(sValue,":")
								For iCount = 0 To cint(dicTraceLinkObjects("ExpandLevel"))
									If iCount <  cint(dicTraceLinkObjects("ExpandLevel")) Then
										If aValue(iCount) <> "true" Then
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_SE_TraceLinkTabOperations ] Failed to verify Expand level-"+cint(dicTraceLinkObjects("ExpandLevel"))+" in TraceLink tree.")
											Exit Function
										End If
									Else
										If aValue(iCount) <> "false" Then
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_SE_TraceLinkTabOperations ] Failed to verify Expand level-"+cint(dicTraceLinkObjects("ExpandLevel"))+" in TraceLink tree.")
											Exit Function
										End If
									End If
								Next
							ElseIf sAction="ExportToExcel" Then  						'[TC1015-2015081100-28_08_2015-VivekA-NewDevelopment] - Added to perform operation on Export To Excel dialog
								Dim objExportToExcel, aReserve, aContentSelection, aTraceabilityViewtoExport
								If Not JavaWindow("Shell").JavaWindow("Export To Excel").Exist Then
									JavaWindow("SystemsEngineering").Click 150,15
						    		wait 1
									JavaWindow("SystemsEngineering").JavaTree(DictKeys(iCounter)).Select sNodePath	
									wait 1
									Call Fn_SISW_UI_JavaToolbar_Operations("Fn_SISW_SE_TraceLinkTabOperations","ClickPackBtn",JavaWindow("SystemsEngineering"),"","Objects To Excel", "","","")
									wait 2
								End If
								
								If JavaWindow("SystemsEngineering").JavaWindow("ExportToExcel").Exist(3) = False Then
								 Call Fn_MenuOperation("Select","Tools:Export:Objects To Excel")
                                 Call Fn_ReadyStatusSync(3)
                                 
									Set objExportToExcel = JavaWindow("SystemsEngineering").JavaWindow("ExportToExcel")
								ElseIf JavaWindow("Shell").JavaWindow("Export To Excel").Exist(3) Then
									Set objExportToExcel = JavaWindow("Shell").JavaWindow("Export To Excel")
								End If
								
								If sReserve <> "" Then
									aReserve = Split(sReserve,"~")
									'Set Content Seletion Radio Button "Export All Visible Columns" or "Use Excel Template" ON
									If aReserve(0)<>"" Then
										aContentSelection = Split(aReserve(0),":")
										objExportToExcel.JavaRadioButton("OutputTemplate").SetTOProperty "attached text",aContentSelection(1)
										Wait 1
										objExportToExcel.JavaRadioButton("OutputTemplate").Set "ON"
										Wait 1
										'Select from list if "Use Excel Template" is ON
										If aContentSelection(1) = "Use Excel Template" AND UBound(aContentSelection) = 2 Then
											objExportToExcel.JavaList("ExcelTemplate").Select aContentSelection(2)
											Wait 1
										End If
									End If
									'Set Traceability View to Export ['"Complying Objects View:ON:Defining Objects View:ON"] 
									If UBound(aReserve) = 1 Then
										If aReserve(1)<>"" Then
											aTraceabilityViewtoExport = Split(aReserve(1),":")
											
											objExportToExcel.JavaCheckBox("CheckOutObjs").SetTOProperty "attached text",aTraceabilityViewtoExport(0)
											Wait 1
											If objExportToExcel.JavaCheckBox("CheckOutObjs").GetROProperty("enabled") = "1" Then
												objExportToExcel.JavaCheckBox("CheckOutObjs").Set aTraceabilityViewtoExport(1)
												Wait 2
											End If
											objExportToExcel.JavaCheckBox("CheckOutObjs").SetTOProperty "attached text",aTraceabilityViewtoExport(2)
											Wait 1
											If objExportToExcel.JavaCheckBox("CheckOutObjs").GetROProperty("enabled") = "1" Then
												objExportToExcel.JavaCheckBox("CheckOutObjs").Set aTraceabilityViewtoExport(3)
												Wait 2
											End If											
										End If
									End IF
									'------------------------------------------------------------------
									'Verify Output Section - TC11.3_20170502a_NewDevelopment_PoonamC_15May2017
									If UBound(aReserve) = 2 Then
											If aReserve(2)<>"" Then
												aTraceabilityViewtoExport = Split(aReserve(2),":")
												if Fn_UI_Object_SetTOProperty_ExistCheck("Fn_SISW_SE_TraceLinkTabOperations",objExportToExcel.JavaStaticText("Output"),"label",Trim(aTraceabilityViewtoExport(1))) = False Then
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to verify [" + aTraceabilityViewtoExport(0) + "] = [" + aTraceabilityViewtoExport(1) + "]")       								
													Fn_SISW_SE_TraceLinkTabOperations = False
													Exit Function
												End if
											End if
									End if
								'----------------------------------------------------------------------	
								End If
								Call Fn_SISW_UI_JavaButton_Operations("Fn_SISW_SE_TraceLinkTabOperations", "Click", objExportToExcel,"OK")
								'------------'[TC11.3-20170509-01_jun_2017-ShwetaR-Development]-Added by Jotiba to Check objects are editable---------------
							ElseIf sAction="IsEditable" Then
								intY = 0
								intX = 0
								Set objTree = objDialog.JavaTree(sTree)
								If Fn_UI_ObjectExist("Fn_SISW_SE_TraceLinkTabOperations", ObjTree) = False Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_SISW_SE_TraceLinkTabOperations : FAIL : " & sUIFail & " Doesn not exist.")	
									Exit function
								End If
								If dicTraceLinkObjects("ColumnNames")<>"" Then
									For iCnt = 0 to objTree.GetROProperty("columns_count")-1
										sColName = objTree.GetColumnHeader(iCnt)
										If trim(LCase(sColName)) = trim(LCase(dicTraceLinkObjects("ColumnNames"))) Then
											Exit For
										End If
									Next
								End If
								Call Fn_SISW_SE_TraceLinkTabOperations("ColumnHeaderPopupMenuSelect",sTabName, "", "", dicTraceLinkObjects,"","Modify Column(s)...")
								 wait 1
							    Call Fn_SISW_SE_ColumnManagementOperation("MoveColumnUp",dicTraceLinkObjects("ColumnNames"),iCnt-1, "", "", "Close")
							    wait 2
							    If dicTraceLinkObjects("ColumnNames")<>"" Then 'To set Index
									For iCnt = 0 to objTree.GetROProperty("columns_count")-1
										sColName = objTree.GetColumnHeader(iCnt)
										If trim(LCase(sColName)) = trim(LCase(dicTraceLinkObjects("ColumnNames"))) Then
											Exit For
										End If
									Next
								End If
								
								For iIterate = 0 to iCnt
									iColWidth = objTree.Object.getColumn(iIterate).getWidth()
									intX = intX + iColWidth
								Next
								intX = intX - iColWidth/2
							    
								For iIterate = 0 to iCnt
									iItmHeight = objTree.Object.getItemHeight()
									intY = intY + iItmHeight
								Next
								intY = intY - iItmHeight/2
							
								objTree.Click intX, intY,"LEFT"
								wait 1
								Set descEdit = Description.Create
								descEdit("to_class").value = "JavaEdit"
								Set objEditChild = ObjTree.ChildObjects(descEdit)
								If objEditChild.count > 0 Then
									 Fn_SISW_SE_TraceLinkTabOperations=False
									 Exit Function 
								End If
							Else
								Exit Function
							End If
					End Select
			Next
	'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		Case "GetColumnList", "GetColCount"
			iColCount = JavaWindow("SystemsEngineering").JavaTree(dicTraceLinkObjects("TreeName")).GetROProperty("columns_count")
			If sAction = "GetColCount" Then
				Fn_SISW_SE_TraceLinkTabOperations =	iColCount
			ElseIf sAction = "GetColumnList" Then
				For iCounter = 0 To iColCount - 1
					If iCounter = 0 Then
						sColName = JavaWindow("SystemsEngineering").JavaTree(dicTraceLinkObjects("TreeName")).GetColumnHeader("#"&iCounter)
					Else
						sColName = sColName &"~" & JavaWindow("SystemsEngineering").JavaTree(dicTraceLinkObjects("TreeName")).GetColumnHeader("#"&iCounter)
					End If
				Next
				Fn_SISW_SE_TraceLinkTabOperations =	sColName
			End If
			Exit Function
	'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		Case "ColumnHeaderPopupMenuSelect"
			JavaWindow("SystemsEngineering").JavaTree(dicTraceLinkObjects("TreeName")).Click 20, -10,"RIGHT"
			wait 1
			aMenuList = split(StrMenu, ":",-1,1)
			intCount = Ubound(aMenuList)
			Select Case intCount
				Case "0"
					 StrMenu =JavaWindow("SystemsEngineering").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0))
				Case "1"
					StrMenu =JavaWindow("SystemsEngineering").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1))
				Case "2"
					StrMenu =JavaWindow("SystemsEngineering").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1),aMenuList(2))
				Case Else
			'		Exit Function
			End Select
			JavaWindow("SystemsEngineering").WinMenu("ContextMenu").Select StrMenu
	'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		Case "VerifyTraceLinkCriteriaWindow"
			Set ObjTraceLinkCriteria=JavaWindow("SystemsEngineering").JavaList("Traceability Scope")
				bReturn=Fn_SISW_UI_Object_Operations("Fn_SISW_SE_TraceLinkTabOperations","Exist",ObjTraceLinkCriteria,"")
				If bReturn=False Then
					Fn_SISW_SE_TraceLinkTabOperations=bReturn
					Exit Function
				End If
				
	'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		Case "VerifySortColumnValue"	
				Set ObjTraceLinktabTree=JavaWindow("SystemsEngineering").JavaTree(dicTraceLinkObjects("TreeName"))
					If dicTraceLinkObjects("TreeName")<> "" Then
							iCount=ObjTraceLinktabTree.GetROProperty("items count")
						
							' Find out Root Node
							For icnt = 0 To iCount-1
								bResult = False
								If Trim(ObjTraceLinktabTree.Object.getItem(icnt).getData().getDisplayedObject().tostring)=Trim(dicTraceLinkObjects("Root Node")) Then
									bResult = True
									Exit For 	
								End If
							Next
							If bResult = False Then
								Exit Function 
							End If
							
							iChildCount=ObjTraceLinktabTree.Object.getItem(icnt).getItemCount
							' Find Out Child Node
							aNodePath=Split(dicTraceLinkObjects("Node Names"),"~")
							For iOuter = 0 To UBound(aNodePath)
									aNode=""
								If aNodePath(iOuter)<> "" Then
										For icnt1 = iOuter To UBound(aNodePath)
											bResult = False
											If Trim(ObjTraceLinktabTree.Object.getItem(icnt).getItem(icnt1).getData().getDisplayedObject().tostring)= Trim(aNodePath(icnt1))Then
												aNode=Trim(aNodePath(icnt1))
												bResult = True
												Exit For 
											Else
												bResult = False
												Exit For 
											End If
										Next
										
										If bResult = False Then
											Exit Function 
										End If
									
										For iCounter=iOuter To iChildCount-1
											bResult = False
											If Trim(ObjTraceLinktabTree.Object.getItem(icnt).getItem(iCounter).getData().getDisplayedObject().tostring)=Trim(aNode) Then
												bResult = True
												Exit For 	
											End If
										Next
										If bResult = False Then
											Exit Function 
										End If
									
										'... to check Column values against column names
										If dicTraceLinkObjects("ColumnNames")<>"" OR dicTraceLinkObjects("ColumnValues") <> "" Then
										
											aColNames = Split(dicTraceLinkObjects("ColumnNames"), ":")
											aSubColName=Split(aColNames(iOuter),"~")
											
											aColValues = Split(dicTraceLinkObjects("ColumnValues"), ":")
											aSubColValues=Split(aColValues(iOuter),"~")
											
												For iCount = 0 To UBound(aSubColName)
														bResult = False
														aColumnName  = Fn_GetXMLNodeValue(Environment.Value("sPath")+ "\TestData\AutomationXML\ObjectRealNames\SystemsEngineering.xml" , Trim(aSubColName(iCount)))
														sColumnValue=ObjTraceLinktabTree.Object.getItem(icnt).getItem(iCounter).getData().getDisplayedObject().getProperty(aColumnName)
													If Trim(aSubColValues(iCount)) =Trim(sColumnValue) Then
														Fn_SISW_SE_TraceLinkTabOperations=True
														bResult=True
													End If
													
													If bResult = False Then
														Exit Function 
													End If
												Next
										End If
										
								End If										
							Next
						End If
	'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		Case "ClickColumnHeader"	
				JavaWindow("SystemsEngineering").JavaTree(dicTraceLinkObjects("TreeName")).Click 5, -5,"LEFT"
			
	'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		Case Else
			Fn_SISW_SE_TraceLinkTabOperations=False
			Exit Function
	End Select
	Fn_SISW_SE_TraceLinkTabOperations=True
End Function

'*******************************************************************************************************************
''Function Name		 	:	Fn_SE_WarningMsgVerify
'
''Description		    :  	This function is use to perform operations on Trace link Warning Window .

''Parameters		    :	1. sAction : Action to be perform
''					2. dicWarningInfo	: Warning information dictionary
''					3. sButton: Button click on Warning Window 
'                           4. Reserve for further use

''Return Value		    :  	true \ false
'
''Examples		     	:		Dim dicWarningInfo
						'	Set dicWarningInfo = CreateObject("Scripting.Dictionary")
						'	dicWarningInfo.Add "Title", ""
						'	dicWarningInfo.Add "WarningMessage", "The Trace Link direction"
						'	dicWarningInfo.Add "ReverseMessage", "Reverse the Trace Link direction."
						'	dicWarningInfo.Add "ProceedMessage", "Create the Trace Link as defined"
						'	dicWarningInfo.Add "RedBtnMessage", "A "default_relation" preference sets the default Trace Link type "
						'	dicWarningInfo.Add "DirectionButton", "Reverse~Proceed"
						'	bReturn = Fn_SE_WarningMsgVerify("VerifyWarningMsg",dicWarningInfo,"Cancel","")
						'	bReturn = Fn_SE_WarningMsgVerify("VerifyTraceLinkDirectionButton",dicWarningInfo,"","")


'History:
'	Developer Name				Date			Rev. No.	Reviewer		Changes Done	
'*******************************************************************************************************************
'	Paresh Dhake		        11-Sept-2014	    1.0			Self	
''	Anurag K					25-Nov-2014								Added case : VerifyTraceLinkDirectionButton
'*******************************************************************************************************************
Function Fn_SE_WarningMsgVerify(sAction,dicWarningInfo,sButton,sReserve)
	GBL_FAILED_FUNCTION_NAME="Fn_SE_WarningMsgVerify"
	'Declare variables
	Dim sAppMsg, aButtonName, iCnt
	Fn_SE_WarningMsgVerify = False
			
	'Set Warning window Object hierarchy 
	Set objWarning = JavaWindow("SystemsEngineering").JavaWindow("WarningMessage")
	'Set title for warning window if provided
	If dicWarningInfo("Title")<>"" Then
		objWarning.SetTOProperty "title", dicWarningInfo("Title")
	End If
	
	'Select action 
	Select Case sAction
	'case for verifyng complete warning messages
		Case "VerifyTraceLinkWarningMsg"
	
	'check existence of warning window	
				If objWarning.Exist(4)=True Then
	'verify main warning message			
					If dicWarningInfo("WarningMessage")<>"" Then
						GBL_EXPECTED_MESSAGE=dicWarningInfo("WarningMessage")
						sAppMsg = objWarning.JavaStaticText("TracelinkMsg").GetROProperty("label")
						If Instr(sAppMsg, dicWarningInfo("WarningMessage")) > 0 Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully verified error message [" + dicWarningInfo("WarningMessage") + "] against Actual Message [" + sAppMsg + "]")     								
						Else
							GBL_ACTUAL_MESSAGE=sAppMsg
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to verify error message [" + dicWarningInfo("WarningMessage") + "] against Actual Message [" + sAppMsg + "]")       								
							Exit Function
						End If
					End If
			
	'Verify Reverse message 			
					If dicWarningInfo("ReverseMessage")<>"" Then
						GBL_EXPECTED_MESSAGE=dicWarningInfo("ReverseMessage")
						If objWarning.JavaButton("Reverse").Exist(2) Then
							'sAppMsg = objWarning.JavaEdit("ReverseMsg").GetROProperty("value")
							sAppMsg = Fn_UI_Object_GetROProperty("Fn_SE_WarningMsgVerify",objWarning.JavaEdit("ReverseMsg"), "value")
							If lcase(dicWarningInfo("ReverseMessage")) = lcase(sAppMsg) Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully verified error message [" + dicWarningInfo("ReverseMessage") + "] against Actual Message [" + sAppMsg + "]")     								
							Else
								GBL_ACTUAL_MESSAGE=sAppMsg
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to verify error message [" + dicWarningInfo("ReverseMessage") + "] against Actual Message [" + sAppMsg + "]")       								
								Exit Function
							End If
						Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to verify existence of 'Reverse' button on Warning window.")    
							Exit Function
						End If
					End If
					
	'Verify Proceed message
					If dicWarningInfo("ProceedMessage")<>"" Then
						GBL_EXPECTED_MESSAGE=dicWarningInfo("ProceedMessage")
						If objWarning.JavaButton("Proceed").Exist(2) Then
							'sAppMsg = objWarning.JavaEdit("ProceedMsg").GetROProperty("value")
							sAppMsg = Fn_UI_Object_GetROProperty("Fn_SE_WarningMsgVerify",objWarning.JavaEdit("ProceedMsg"), "value")
							If lcase(dicWarningInfo("ProceedMessage")) = lcase(sAppMsg) Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully verified error message [" + dicWarningInfo("ProceedMessage")+ "] against Actual Message [" + sAppMsg + "]")     								
							Else
								GBL_ACTUAL_MESSAGE=sAppMsg
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to verify error message [" + dicWarningInfo("ProceedMessage") + "] against Actual Message [" + sAppMsg + "]")       								
								Exit Function
							End If
						Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to verify existence of 'Proceed' button on Warning window.")       								
							Exit Function
						End If
					End If
					
	'Verify Red button message
					If dicWarningInfo("RedBtnMessage")<>"" Then
						GBL_EXPECTED_MESSAGE=dicWarningInfo("RedBtnMessage")
						'Click on Red button for details message
						Call Fn_SISW_UI_JavaButton_Operations("Fn_SE_WarningMsgVerify", "Click", objWarning, "RedButton")
						If objWarning.JavaWindow("Details").Exist(4) = True Then
							sAppMsg = objWarning.JavaWindow("Details").JavaEdit("DetailsMsg").GetROProperty("value")
							If Instr(sAppMsg, dicWarningInfo("RedBtnMessage")) > 0 Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully verified error message [" + dicWarningInfo("RedBtnMessage") + "] against Actual Message [" + sAppMsg + "]")     								
								Call Fn_SISW_UI_JavaButton_Operations("Fn_SE_WarningMsgVerify", "Click", objWarning.JavaWindow("Details"), "OK")
							Else
								GBL_ACTUAL_MESSAGE=sAppMsg
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to verify error message [" + dicWarningInfo("RedBtnMessage") + "] against Actual Message [" + sAppMsg + "]")       								
								Exit Function
							End If
						Else
							Fn_SE_WarningMsgVerify = False
							Exit Function
						End If
					End If
				Else
	'Non-Existence of Warning
					Fn_SE_WarningMsgVerify = False
					Exit Function
				End If
				
				If sReserve<>"" Then
					'Do Nothing
				End If
				
		'case for verifyng TraceLink direction Button
		Case "VerifyTraceLinkDirectionButton"
				If dicWarningInfo("DirectionButton")<>"" Then 
					If Instr( dicWarningInfo("DirectionButton") , "~") > 0 Then
						aButtonName = Split( dicWarningInfo("DirectionButton"),"~")
					   	For iCnt = 0 to Ubound(aButtonName)
					   	   If Fn_SISW_UI_Object_Operations("Fn_SE_WarningMsgVerify","Exist", objWarning.JavaButton(aButtonName(iCnt)),"") = False Then
					   	    	 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to verifiy direction button [" +aButtonName(iCnt)+ "] in warning message window")     								
					   	     	 Exit Function
					   	   Else
					   		     Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully verified direction button [" +aButtonName(iCnt)+ "] in warning message window")     								
					   	   End If
					   	Next
					Else
						If Fn_SISW_UI_Object_Operations("Fn_SE_WarningMsgVerify","Exist", objWarning.JavaButton(dicWarningInfo("DirectionButton")),"") = False Then
				   	     	 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to verifiy direction button [" +dicWarningInfo("DirectionButton")+ "] in warning message window")     								
				   	    	 Exit Function
				   	    Else
				   	       	 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully verified direction button [" +dicWarningInfo("DirectionButton")+ "] in warning message window")     								
				  	    End If
					End If
						
			   Else
			       	Fn_SE_WarningMsgVerify = False
					Exit Function
			   End If
			
	Case Else
	'Else case when sAction not provided
			Fn_SE_WarningMsgVerify = False
			Exit Function
	End Select
	
	'Click on button on Warning window
	If sButton<>"" Then
		Call Fn_SISW_UI_JavaButton_Operations("Fn_SE_WarningMsgVerify", "Click", objWarning, sButton)
	End If
	
	Fn_SE_WarningMsgVerify = True
	Set objWarning = nothing
End Function
'*********************************************************		Function to  perform Operation on different tab of Tracelink Criteria***********************************************************************

'Function Name		:					Fn_SISW_SE_TraceLinkCriteriaOperations

'Description			 :		 		  Operation on different tab of Tracelink Criteria

'Parameters			   :	 			1.  sAction: Action to be performed
'									 2. sTabName : Tab to be select
'									 3.. bMaximiseTab : "Yes": want to maximize the tracelink tab
'									4. sButton	: pass "Apply" if want to click on Apply button
' 									5.dicTraceLinkObjects:
'										Set dicTraceLinkObjects = CreateObject("Scripting.Dictionary")
'										dicTraceLinkObjects("TracelinkName") = "BZ_FND_Tracelink"
'									sReserve, sObjectsCheck for future use
											
'Return Value		   : 				 True/False

'Pre-requisite		:		 	Trace Link Tab should be selected

'Examples		:			Fn_SISW_SE_TraceLinkCriteriaOperations("AddTraceLink","Link Type Filters","","Apply",dicTraceLinkObjects,"","")

'History:
'	Developer Name			Date				Rev. No.			Changes Done							Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Snehal Salunkhe			19-May-2015			1.0		
'	Shweta Rathod			30-Jun-2017			1.0			      added case "AddAll","RemoveAll" 
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_SE_TraceLinkCriteriaOperations(sAction, sTabName, bMaximiseTab, sButton, dicTraceLinkObjects, sReserve, sObjectsCheck)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SE_TraceLinkCriteriaOperations"
	'Declare Veriables
	Dim DictItems,DictKeys,objTab,bResult, sNodeName, iCnt, iItemsCount, bFlag
	Dim iCounter,sValue,objTable,objJavaList
	
	
	Fn_SISW_SE_TraceLinkCriteriaOperations = False
	
	If JavaWindow("SystemsEngineering").JavaWindow("ErrorWindow").Exist(2) Then
		Call Fn_SISW_UI_JavaButton_Operations("Fn_SISW_SE_TraceLinkTabOperations", "Click", JavaWindow("SystemsEngineering").JavaWindow("ErrorWindow"),"OK")
		Call Fn_ReadyStatusSync(1)
	End If
	'Select traceLink tab
	'[TC1123-20161115a-23_11_2016-VivekA-Maintenance] - Added from TC1017 - Added by Archana D
	If sAction <> "AddtoScope" Then
		Call Fn_SISW_UI_RACTabFolderWidget_Operation("Select","Trace Links", "")
		wait 1
	End If
	If JavaWindow("SystemsEngineering").JavaWindow("ErrorWindow").Exist(2) Then
		Call Fn_SISW_UI_JavaButton_Operations("Fn_SISW_SE_TraceLinkTabOperations", "Click", JavaWindow("SystemsEngineering").JavaWindow("ErrorWindow"),"OK")
	End If
	Call Fn_ReadyStatusSync(2)
	'--------------------------------------------------
	
	Set objTab = JavaWindow("SystemsEngineering").JavaTab("CTabFolder")
	'Maximise tab
	If bMaximiseTab = "Yes" Then
		'Call Fn_SE_RightPanelTabOperations("DoubleClick", "Trace Links", "")
		Call Fn_TabFolder_Operation("DoubleClickTab","Trace Links", "")
		wait 1
	End If
	
	Select Case sTabName
		'[TC1122-20160203-15_02_2016-VivekA-Maintenance] - Added from TC10.1.5
		Case "Property Filters"
			If Fn_SISW_UI_Object_Operations("Fn_SISW_SE_TraceLinkCriteriaOperations","Exist", objTab.JavaTable("Filter by Properties"),"") <> True Then
				Call Fn_SISW_UI_JavaTab_Operations("", "Select", JavaWindow("SystemsEngineering"), "CTabFolder", sTabName)	
			End If
			Call Fn_ReadyStatusSync(1)
			'===================================
			Select Case sAction
				Case "Add"								
					arrPropertyName   = Split(dicTraceLinkObjects("PropertyName"),"~")
					
					For iCounter = 0 To UBound(arrPropertyName)
						Call Fn_SISW_UI_JavaButton_Operations("Fn_SISW_SE_TraceLinkCriteriaOperations", "Object.setFocus", objTab.JavaButton("Plus"),"")
						wait 1
						If arrPropertyName(iCounter) <> "" Then							'Add Value in Property Value
							objTab.JavaTable("Filter by Properties").ActivateCell iCounter,"Property Name"
							wait 1
							objTab.JavaTable("Filter by Properties").SelectCell iCounter , "Property Name"
							wait 1
							Set dicProperty = Description.Create
							dicProperty("to_class").Value = "JavaList"
							set objList = JavaWindow("SystemsEngineering").JavaWindow("Shell").ChildObjects(dicProperty)
							For iIterator1 = 0 To objList.Count - 1
								bFlag = False
								For Iterator = 0 To objList(iIterator1).getROProperty("items count") - 1
									If trim(objList(Iterator1).getItem(Iterator)) = trim(arrPropertyName(iCounter)) Then
										bFlag = True
										Exit For
									End If									
								Next
								If bFlag = True Then
									objList(iIterator1).Select arrPropertyName(iCounter)	
									Exit For
								Else 
									Fn_SISW_SE_TraceLinkCriteriaOperations = False
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Select Property Values in Filter by Property Table ")     								
									Exit Function	
								End If									
							Next	
							Set dicProperty = nothing 							
						End If
						If dicTraceLinkObjects("Condition") <> "" Then
							If iCounter = 0 Then
								arrCondition = Split(dicTraceLinkObjects("Condition"),"~")
							End If
							If arrCondition(iCounter) <> "" AND iCounter > 0 Then				'Add Value in Logical FIlter
								objTab.JavaTable("Filter by Properties").ActivateCell iCounter , 0
								wait 1 
								objTab.JavaTable("Filter by Properties").SelectCell iCounter , 0
								wait 1
								Set dicProperty = Description.Create
								dicProperty("to_class").Value = "JavaList"
								set objList = JavaWindow("SystemsEngineering").JavaWindow("Shell").ChildObjects(dicProperty)
								For iIterator1 = 0 To objList.Count - 1
									bFlag = False
									For Iterator = 0 To objList(iIterator1).getROProperty("items count") - 1
										If trim(objList(Iterator1).getItem(Iterator)) = trim(arrCondition(iCounter)) Then
											bFlag = True
											Exit For
										End If									
									Next
									If bFlag = True Then
										objList(iIterator1).Select arrCondition(iCounter)
										Exit For
									Else 
										Fn_SISW_SE_TraceLinkCriteriaOperations = False	
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Select Condition Values in Filter by Property Table ")     								
										Exit Function									
									End If								
								Next
								Set dicProperty = nothing 
							End If
						End If
						
						If dicTraceLinkObjects("EqualCondition") <> "" Then
							If iCounter = 0 Then
								arrEqualCondition = Split(dicTraceLinkObjects("EqualCondition"),"~")
							End If								
							If arrEqualCondition(iCounter) <> "" Then						''Add Value in Equal Condition
								objTab.JavaTable("Filter by Properties").ActivateCell iCounter , 2
								wait 1
								objTab.JavaTable("Filter by Properties").SelectCell iCounter , 2
								Set dicProperty = Description.Create
								dicProperty("to_class").Value = "JavaList"
								set objList = JavaWindow("SystemsEngineering").JavaWindow("Shell").ChildObjects(dicProperty)
								For iIterator1 = 0 To objList.Count - 1
									bFlag = False
									For Iterator = 0 To objList(iIterator1).getROProperty("items count") - 1
										If trim(objList(Iterator1).getItem(Iterator)) = trim(arrEqualCondition(iCounter)) Then
											bFlag = True
											Exit For
										End If									
									Next
									If bFlag = True Then
										objList(iIterator1).Select arrEqualCondition(iCounter)
										Exit For
									Else 
										Fn_SISW_SE_TraceLinkCriteriaOperations = False
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Select Equal Field in Filter by Property Table ")     								
										Exit Function									
									End If									
								Next	
								Set dicProperty = nothing 								
							End If							
						End If
						If dicTraceLinkObjects("SearchValue") <> ""  Then
							If iCounter = 0 Then
								arrSearchValue = Split(dicTraceLinkObjects("SearchValue"),"~")
							End If	
							If arrSearchValue(iCounter) <> "" Then						'Added Value in Search Value
								objTab.JavaTable("Filter by Properties").SetCellData iCounter ,"Searching Value",arrSearchValue(iCounter) 		
								If Err.number < 0 Then
									Fn_SISW_SE_TraceLinkCriteriaOperations = False
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Set Searching Value in Filter by Property Table ")     								
									Exit Function
								End If							
							End If								
						End If	
					Next
				Case "Clear"
					If Fn_UI_Object_GetROProperty("Fn_SISW_SE_TraceLinkCriteriaOperations",objTab.JavaButton("Clear"),"enabled") = "1" Then
						Call Fn_SISW_UI_JavaButton_Operations("Fn_SISW_SE_TraceLinkCriteriaOperations", "Object.setFocus", objTab.JavaButton("Clear"),"")
						If Err.number < 0 Then
							bFlag = False
						Else 
							bFlag = True						
						End If
					Else 
						bFlag = False					
					End If
					If bFlag = False Then
						If bMaximiseTab = "Yes" Then
							Call Fn_SISW_UI_RACTabFolderWidget_Operation("DoubleClick", "Trace Links", "")
							wait 1
						End If
						Fn_SISW_SE_TraceLinkCriteriaOperations = False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to click Clear Filter button in Property Table ")     								
						Exit Function
					End If
				Case "GetPropertyNames"		
					Fn_SISW_SE_TraceLinkCriteriaOperations = ""
					
					Call Fn_SISW_UI_JavaButton_Operations("Fn_SISW_SE_TraceLinkCriteriaOperations", "Object.setFocus", objTab.JavaButton("Plus"),"")
					wait 1					
					iCount = Fn_UI_Object_GetROProperty("Fn_SISW_SE_TraceLinkCriteriaOperations",objTab.JavaTable("Filter by Properties"),"rows")
					objTab.JavaTable("Filter by Properties").ActivateCell (iCount - 1), "Property Name"
					wait 1
					objTab.JavaTable("Filter by Properties").SelectCell (iCount - 1), "Property Name"
					wait 1		
					
					Set dicProperty = Description.Create
					dicProperty("to_class").Value = "JavaList"
					set objList = JavaWindow("SystemsEngineering").JavaWindow("Shell").ChildObjects(dicProperty)							
					For iIterator1 = 0 To objList.Count - 1
						For Iterator = 0 To objList(iIterator1).getROProperty("items count") - 1
							If Iterator = objList(iIterator1).getROProperty("items count") - 1 Then
								Fn_SISW_SE_TraceLinkCriteriaOperations = Fn_SISW_SE_TraceLinkCriteriaOperations & trim(objList(Iterator1).getItem(Iterator))
							Else 
								Fn_SISW_SE_TraceLinkCriteriaOperations = Fn_SISW_SE_TraceLinkCriteriaOperations & trim(objList(Iterator1).getItem(Iterator)) & "~"									
							End If							
						Next	
						Exit For							
					Next	
					Set dicProperty = nothing 
					
					Call Fn_SISW_UI_JavaButton_Operations("Fn_SISW_SE_TraceLinkCriteriaOperations", "Object.setFocus", objTab.JavaButton("Minus"),"")
					
					If bMaximiseTab = "Yes" Then
						Call Fn_SISW_UI_RACTabFolderWidget_Operation("DoubleClick", "Trace Links", "")
						wait 1
					End If
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully fetched value of property names in Property Table ")     								
					Exit Function
				Case Else
					Exit Function
			End Select
			
			'===================================		
		Case "Link Type Filters", "Object Type Filters"  '[TC1122-20160203-15_02_2016-VivekA-Maintenance] - Added from TC10.1.5
			Call Fn_SISW_UI_JavaTab_Operations("", "Select", JavaWindow("SystemsEngineering"), "CTabFolder", sTabName)
			Call Fn_ReadyStatusSync(1)
			
			Select Case sAction
				Case "AddTraceLink","AddMultipleTraceLink"
					'Selecting Tracelink from available List
				      	sTracelink=Split(dicTraceLinkObjects("TracelinkName"),"~","-1")
				      	If objTab.JavaTable("AvailableObjectTypes").Exist(2) = True Then
				      		For iCounter = 0 To UBound(sTracelink) 	
				      			If iCounter=0 Then
					      			bResult= Fn_Table_Select_Cell("", objTab, "AvailableObjectTypes",sTracelink(iCounter),0)
				      			Else
						      		bResult= Fn_UI_JavaTable_ExtendRow("",objTab,"AvailableObjectTypes",sTracelink(iCounter))
							End If	
							wait 1
					      		If bResult = False Then
						      		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to select row ["+sTracelink(iCounter)+"]")     								
								Exit Function
							End If		
						Next
				      	End If
				      	'Click on Add button
					bResult  = Fn_SISW_UI_JavaButton_Operations("Fn_SISW_SE_TraceLinkCriteriaOperations", "Click", objTab, "Add")
					If bResult = False Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to click on [Add] button")     								
						Exit Function
					End If
					Call Fn_ReadyStatusSync(1)

				Case "RemoveAll"
					    'Click on Remove all button
					bResult  = Fn_SISW_UI_JavaButton_Operations("Fn_SISW_SE_TraceLinkCriteriaOperations", "Click", objTab, "Remove All")
					If bResult = False Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to click on [Remove all] button")     								
						Exit Function
					End If
					Call Fn_ReadyStatusSync(1)
					
				Case "RemoveTraceLink", "RemoveSelectedObjectType"
					'Remove Tracelink
					sTracelink=Split(dicTraceLinkObjects("TracelinkName"),"~")
				      	If objTab.JavaTable("SelectedObjectTypes").Exist(2) = True Then
				      		For iCounter = 0 To UBound(sTracelink) 	
				      			If iCounter=0 Then
					      			bResult= Fn_Table_Select_Cell("", objTab, "SelectedObjectTypes",sTracelink(iCounter),0)
				      			Else
						      		bResult= Fn_UI_JavaTable_ExtendRow("",objTab,"SelectedObjectTypes",sTracelink(iCounter))
							End If	
							wait 1
					      		If bResult = False Then
						      		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to select row ["+sTracelink(iCounter)+"]")     								
								Exit Function
							End If		
						Next
				      	End If
				      	'Click on Remove button
					bResult  = Fn_SISW_UI_JavaButton_Operations("Fn_SISW_SE_TraceLinkCriteriaOperations", "Click", objTab, "Remove")
					If bResult = False Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to click on [Remove] button")     								
						Exit Function
					End If
					Call Fn_ReadyStatusSync(1)			
					
				'[TC1122-20160309-28_03_2016-VivekA-Maintenance] - Added by Chaitali R - Added from TC1016
				Case "VerifySelectedObjectType"
					sTracelink=Split(dicTraceLinkObjects("TracelinkName"),"~","-1")
					If objTab.JavaTable("SelectedObjectTypes").Exist(2) = True Then
						iItemsCount = Fn_Table_GetRowCount("Fn_SISW_SE_TraceLinkCriteriaOperations",objTab, "SelectedObjectTypes")
						For iCount = 0 To UBound(sTracelink)
							bResult = False
							For iCounter = 0 To iItemsCount-1
								If Trim(sTracelink(iCount)) = Trim(objTab.JavaTable("SelectedObjectTypes").Object.getItem(iCounter).gettext()) Then
									bResult = True
									Exit For
								End if
							Next	
							If bResult = False Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to verify node ["+sTracelink(iCount)+"].")
								Exit Function
							End If
						Next
					End If
				'------------------------------------------------------------------------					
				Case Else
					Exit Function
			End Select
			
		Case "Scope"
			Call Fn_SISW_UI_JavaTab_Operations("", "Select", JavaWindow("SystemsEngineering"), "CTabFolder", "Scope")
			Call Fn_ReadyStatusSync(1)
			
			Select Case sAction
				Case "VerifyScope"
					'Verify object in "Traceability Scope"
					sNodeName = Split(dicTraceLinkObjects("TraceabilityScope"),"~")
					Set objJavaList=JavaWindow("SystemsEngineering").JavaList("Traceability Scope")	
					iItemsCount = objJavaList.GetROProperty("items count")	
					For iCnt = 0 To UBound(sNodeName)	
						bFlag = False
						For iCounter = 0 To iItemsCount-1 
							if sNodeName(iCnt) = objJavaList.GetItem(iCounter) then
								bFlag = True
								Exit For
							End if
						Next
						If bFlag = False Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Verify object added in Scope.")
							Exit Function
						End If
					Next
				'Case to add selected element in Scope tab and then apply
				'[TC1123-20161115a-23_11_2016-VivekA-Maintenance] - Added from TC1017 - Added by Archana D
				Case "AddtoScope"
					If sButton <> "" Then
						'[TC1015-2015091500-29_09_2015-VivekA-NewDevelopment] - Modified by Snehal
						If Instr(sButton,"~") > 0 Then
							aButton = Split(sButton,"~")
							JavaWindow("SystemsEngineering").JavaButton("RemoveBudget").SetTOProperty "label",aButton(0)
							sButton = aButton(1)
						Else
							aButton = sButton
							sButton = ""
							JavaWindow("SystemsEngineering").JavaButton("RemoveBudget").SetTOProperty "label",aButton
						End If
						
						Wait 1
						If JavaWindow("SystemsEngineering").JavaButton("RemoveBudget").GetROProperty("enabled") = "1" Then
							JavaWindow("SystemsEngineering").JavaButton("RemoveBudget").Click micLeftBtn
						End If														
					End If
				'[TC1123-20161115a-23_11_2016-VivekA-Maintenance] - Added from TC1017 - Added by Archana D
				Case "RemoveFromScope"
					If sButton <> "" Then
						sNodeName = dicTraceLinkObjects("TraceabilityScope")
						Set objJavaList = JavaWindow("SystemsEngineering").JavaList("Traceability Scope")	
						iItemsCount = objJavaList.GetROProperty("items count")	
						bFlag = False
						For iCounter = 0 To iItemsCount-1 
							if sNodeName = objJavaList.GetItem(iCounter) then
								objJavaList.Select sNodeName
								bFlag = True
								Exit For
							End if
						Next
						If bFlag = False Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: to Select object in Scope.")
							Exit Function
						End If
						If Instr(sButton,"~") > 0 Then
							aButton = Split(sButton,"~")
							JavaWindow("SystemsEngineering").JavaButton("RemoveBudget").SetTOProperty "label",aButton(0)
							sButton = aButton(1)
						Else
							aButton = sButton
							sButton = ""
							JavaWindow("SystemsEngineering").JavaButton("RemoveBudget").SetTOProperty "label",aButton
						End If
						Wait 1
						If JavaWindow("SystemsEngineering").JavaButton("RemoveBudget").GetROProperty("enabled") = "1" Then
							JavaWindow("SystemsEngineering").JavaButton("RemoveBudget").Click micLeftBtn
						End If														
					End If
					Case "AddAll"
					If sButton <> "" Then
						'[TC113-20170509d-20_06_2017-shweta-NewDevelopment] 
						If Instr(sButton,"~") > 0 Then
							aButton = Split(sButton,"~")
							JavaWindow("SystemsEngineering").JavaButton("RemoveBudget").SetTOProperty "label",aButton(0)
							sButton = aButton(1)
						Else
							aButton = sButton
							sButton = ""
							JavaWindow("SystemsEngineering").JavaButton("RemoveBudget").SetTOProperty "label",aButton
						End If
						
						Wait 1
						If JavaWindow("SystemsEngineering").JavaButton("RemoveBudget").GetROProperty("enabled") = "1" Then
							JavaWindow("SystemsEngineering").JavaButton("RemoveBudget").Click micLeftBtn
						End If														
					End If
					Case "RemoveAll"
					If sButton <> "" Then
						'[TC113-20170509d-20_06_2017-shweta-NewDevelopment] 
						If Instr(sButton,"~") > 0 Then
							aButton = Split(sButton,"~")
							JavaWindow("SystemsEngineering").JavaButton("RemoveBudget").SetTOProperty "label",aButton(0)
							sButton = aButton(1)
						Else
							aButton = sButton
							sButton = ""
							JavaWindow("SystemsEngineering").JavaButton("RemoveBudget").SetTOProperty "label",aButton
						End If
						
						Wait 1
						If JavaWindow("SystemsEngineering").JavaButton("RemoveBudget").GetROProperty("enabled") = "1" Then
							JavaWindow("SystemsEngineering").JavaButton("RemoveBudget").Click micLeftBtn
						End If														
					End If
				'--------------------------
			End Select
			
		Case Else
			Exit Function
	End Select
	'For Click on Apply
	If sButton <> "" Then
		Call Fn_SISW_UI_JavaButton_Operations("Fn_SISW_SE_TraceLinkTabOperations", "Click", JavaWindow("SystemsEngineering"), sButton)
	End If	
	
	If bMaximiseTab = "Yes" Then
		Call Fn_SE_RightPanelTabOperations("DoubleClick", "Trace Links", "")
		wait 1
	End If
	
	Fn_SISW_SE_TraceLinkCriteriaOperations = True
	Set objTab = nothing
End Function
'*********************************	Function to handle different sceanarios for column configuration ***************************************************************

'Function Name			:			Fn_SISW_SE_ColumnConfigurationOperation

'Description			:		 	To perform Operations on Trace Link tab and Bom View tab

'Parameters			   	:	 		1. sAction: Action to be performed
'									2. sConfigName : Name need to Publish or MarkAsPublish
'									3. sTab : BOMLineView or TraceLinks
'									4. sMenu	: Menu to Select
' 									5. sButtonName : Button to click
'									6. sExtra : For future purpose		
'Return Value		   	: 			True/False

'Pre-requisite			:	 		Trace Link Tab or BOMLineView  should be selected.

'Examples				:			Fn_SISW_SE_ColumnConfigurationOperation("MarkAsPublishable","ConfigName","BOMLineView","Mark As Publish","OK","")

'Developer Name			Date				Rev. No.			Changes Done							Reviewer	
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Snehal Salunkhe			19-May-2015			        1.0		
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_SE_ColumnConfigurationOperation(sAction,sConfigName, sTab,sMenu, sButtonName, sExtra)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SE_ColumnConfigurationOperation"
	Dim  bFlag,objMarkCoulmnConfig,objPublishCoulmnConfig
	Fn_SISW_SE_ColumnConfigurationOperation = False
	bFlag = False
	
	'Create Object of publishable Window
	Set objMarkCoulmnConfig = JavaWindow("SystemsEngineering").JavaWindow("MarkAsPublishable")
	'Create Object of Publish Column Configuration
	Set objPublishCoulmnConfig = JavaWindow("SystemsEngineering").JavaWindow("PublishColumnConfiguration")
	
	If sTab = "BOMLineView" Then
			Call Fn_ToolBarOperation("ShowDropdownAndSelectWithInstance", "View Menu:2", sMenu)
	ElseIf sTab = "TraceLinks" Then
			Call Fn_ToolBarOperation("ShowDropdownAndSelectWithInstance", "View Menu:3", sMenu)
	End If 
	Call Fn_ReadyStatusSync(2)	

Select Case sAction
		'Case for Select Mark as Publish
		Case "MarkAsPublishable"	
			'Verifying in list the names and publish it
			If objMarkCoulmnConfig.Exist(2) Then
					If sConfigName<>"" Then
							bFlag = Fn_SISW_UI_JavaList_Operations("Fn_SISW_SE_ColumnConfigurationOperation", "Exist", objMarkCoulmnConfig,"Column Views",sConfigName, "", "")
							If bFlag = True Then
								Call Fn_SISW_UI_JavaList_Operations("Fn_SISW_SE_ColumnConfigurationOperation", "Select", objMarkCoulmnConfig,"Column Views",sConfigName, "", "")
								Call Fn_ReadyStatusSync(1)
								'Click on Button
								Call Fn_Button_Click("Fn_SISW_SE_ColumnConfigurationOperation",objMarkCoulmnConfig,"OK")
								Fn_SISW_SE_ColumnConfigurationOperation = True
							Else
								'Click on Cancel Button
								Call Fn_Button_Click("Fn_SISW_SE_ColumnConfigurationOperation",objMarkCoulmnConfig,"Cancel")
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Select "+sConfigName+" from Mark as publish list.")
								Fn_SISW_SE_ColumnConfigurationOperation = False
							End If
					End IF
			End If
		'Case for Add public column configuration
		Case "PublishColumnConfigurationAdd"
			If objPublishCoulmnConfig.Exist(2) Then
					'Verifying in list the names and Add to make Public		
					If sConfigName<>"" Then
							bFlag = Fn_SISW_UI_JavaList_Operations("Fn_SISW_SE_ColumnConfigurationOperation", "Exist", objPublishCoulmnConfig,"Available",sConfigName, "", "")
							If bFlag = True Then
								Call Fn_SISW_UI_JavaList_Operations("Fn_SISW_SE_ColumnConfigurationOperation", "Select", objPublishCoulmnConfig,"Available",sConfigName, "", "")
								Call Fn_ReadyStatusSync(1)
								'Click on Add Button
								Call Fn_Button_Click("Fn_SISW_SE_ColumnConfigurationOperation",objPublishCoulmnConfig,"AddColumn")
								Call Fn_ReadyStatusSync(1)
								'Click on OK Button
								Call Fn_Button_Click("Fn_SISW_SE_ColumnConfigurationOperation",objPublishCoulmnConfig,"OK")
								Fn_SISW_SE_ColumnConfigurationOperation = True
							Else
								'Click on Cancel Button
								Call Fn_Button_Click("Fn_SISW_SE_ColumnConfigurationOperation",objPublishCoulmnConfig,"Cancel")
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to add "+sConfigName+" from Publish column list.")
								Fn_SISW_SE_ColumnConfigurationOperation = False
							End If
					End If
			End If
			
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Invalid Case.")
			Fn_SISW_SE_ColumnConfigurationOperation = False
	End Select
	
	Set objMarkCoulmnConfig = nothing
	Set objPublishCoulmnConfig = nothing
End Function

'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'###    FUNCTION NAME   :   Fn_SISW_SE_ColumnManagementOperation
'###
'###    DESCRIPTION        :  	System Engineering utility function which allpied to Column Management Dialog Operation
'###
'###    PARAMETERS      :    1.  sAction: Action string to navigate to appropriate case
'### 								       	2.  aColumnNames: Array of Column Names          
'###								        3.  iColPos: No of Position to Move Column
'###									    4.  sColConfigName: New Column configuration Name
'### 									    5.  sColConfigDescription: New Column  Configuration Description
'### 									    6.  sButtonName:  Button Name to Click On
'###    EXAMPLES
'###         Case 1 :       ColumnAdd :  Msgbox Fn_SISW_SE_ColumnManagementOperation("ColumnAdd", "Relation", "", "", "", "")
'###         Case 2:        ColumnRemove : Msgbox Fn_SISW_SE_ColumnManagementOperation("ColumnRemove", "Relation", "", "", "", "")
'###         Case 3:        ColumnValidateExists  : Msgbox Fn_SISW_SE_ColumnManagementOperation("ColumnValidateExists", "Relation", "", "", "", "")
'###         Case 4: 		MoveColumnUp		:Msgbox Fn_SISW_SE_ColumnManagementOperation("MoveColumnUp", "Description", "4", "", "", "")
'###         Case 5: 		MoveColumnDown	:Msgbox Fn_SISW_SE_ColumnManagementOperation("MoveColumnDown", "Description", "4", "", "", "")
'History					 :			
'					Developer Name				Date						Rev. No.						Changes Done	
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'					Snehal Salunkhe		           19-May-2015			        1.0
'					Nishigandha J				   23-Nov-2016			             1.1								Migrated case "ColumnAdd_SaveViewConfig" from TC1017
'					Kaveri Parab				 23-May-2016			           1.1								Added case - ColumnAdd_SaveViewConfigWithoutApply
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_SE_ColumnManagementOperation(sAction, aColumnNames, iColPos, sColConfigName, sColConfigDescription, sButtonName)
GBL_FAILED_FUNCTION_NAME="Fn_SISW_SE_ColumnManagementOperation"
Dim  iCounter, ObjDialog, bReturn, bFlag, aColNames, iCount,objSaveViewConfig
Fn_SISW_SE_ColumnManagementOperation = False
bFlag = False

'Creating Object of [ SaveViewConfiguration ] Window
Set objSaveViewConfig = JavaWindow("SystemsEngineering").JavaWindow("SaveViewConfiguration")
Set ObjDialog = JavaWindow("SystemsEngineering").JavaWindow("ColumnManagement")

	Select Case sAction
		Case "ColumnAdd", "ColumnAdd_SaveConfig", "ColumnAdd_SaveViewConfig" , "ColumnAdd_SaveViewConfigWithoutApply"
				If aColumnNames <> "" Then
					aColNames = Split(aColumnNames, "~")
					For iCounter = 0 To UBound(aColNames)
						bFlag = False
						If  Fn_SISW_SE_ColumnManagementOperation("ColumnValidateExists", aColNames(iCounter), "", "", "", "") = True Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), aColumnNames&" column already exist in Details Table.")
							bFlag = True
							'[TC1122-20160309-28_03_2016-VivekA-Maintenance] - Added by Shriram K
							If UBound(aColNames)=0 Then			
								Fn_SISW_SE_ColumnManagementOperation = True
								If sAction = "ColumnAdd_SaveConfig" Then
									'Click on Apply Button
									Call Fn_Button_Click("Fn_SISW_SE_ColumnManagementOperation", ObjDialog, "Save")
									wait 1
									bFlag = Fn_SE_SaveColumnConfiguration(sColConfigName,sColConfigDescription)
								End If
								If sButtonName <> "" Then
									Call Fn_Button_Click("Fn_SISW_SE_ColumnManagementOperation", ObjDialog, sButtonName)
								End If
								Exit Function
							End If 
							'------------------------------------------------------------------------								
						Else
							'Count number of rows of Table
							bReturn = ObjDialog.JavaTable("AvailableProp").GetROProperty("rows")	
							'Extract the index of row at which the object exist.
							For iCount=0 to bReturn - 1
								If Trim(Lcase(ObjDialog.JavaTable("AvailableProp").GetCellData(iCount,"Property"))) = Trim(Lcase(aColNames(iCounter))) then
									'ObjDialog.JavaTable("AvailableProp").ExtendRow iCount
									If iCounter > 0 Then
										ObjDialog.JavaTable("AvailableProp").ExtendRow iCount
									Else
										ObjDialog.JavaTable("AvailableProp").SelectCell iCount, 0
									End If
									bFlag = True
									Exit For
								End If
							Next
							If bFlag = False Then
								Exit Function
							End If
						End If
					Next
					'Click on Add column Button
					Call Fn_Button_Click("Fn_SISW_SE_ColumnManagementOperation", ObjDialog, "AddColumn")
					If sAction <> "ColumnAdd_SaveViewConfigWithoutApply" Then
						'Click on Apply Button
						Call Fn_Button_Click("Fn_SISW_SE_ColumnManagementOperation", ObjDialog, "Apply")
					End If	
					
				End If
				If sAction = "ColumnAdd_SaveConfig" Then
					'Click on Apply Button
					Call Fn_Button_Click("Fn_SISW_SE_ColumnManagementOperation", ObjDialog, "Save")
					wait 1
					bFlag = Fn_SE_SaveColumnConfiguration(sColConfigName,sColConfigDescription)
				End If
				
				'[TC1123-20161115a-23_11_2016-VivekA-Maintenance] - Added from TC1017 - Added case to save view configuration dialog - By Nishigandha J
				If sAction = "ColumnAdd_SaveViewConfig" or sAction= "ColumnAdd_SaveViewConfigWithoutApply" Then
					'Click on Apply Button
					Call Fn_Button_Click("Fn_SISW_SE_ColumnManagementOperation", ObjDialog, "Save")
					wait 1
					bFlag = Fn_SE_SaveViewConfiguration(sColConfigName, sColConfigDescription)
				End If
				'-----------------------------------------------------------------------
				
				'Click on Close Button
				If sButtonName <> "" Then
					Call Fn_Button_Click("Fn_SISW_SE_ColumnManagementOperation", ObjDialog, sButtonName)
				End If
				wait 1
				If bFlag=True Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_SISW_SE_ColumnManagementOperation passed with case "&sAction)
					Fn_SISW_SE_ColumnManagementOperation = True
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_SISW_SE_ColumnManagementOperation passed with case "&sAction)
					Fn_SISW_SE_ColumnManagementOperation = False
				End If									

		Case "ColumnRemove", "ColumnRemove_SaveConfig"
				If aColumnNames <> "" Then
					aColNames = Split(aColumnNames, "~")
					For iCounter = 0 To UBound(aColNames)
						bFlag = False
						If  Fn_SISW_SE_ColumnManagementOperation("ColumnValidateExists", aColNames(iCounter), "", "", "", "") = False Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), aColumnNames&" column already removed from Details Table.")
							bFlag = True
						Else
							'Count number of rows of Table
							bReturn = ObjDialog.JavaTable("DisplayedColumns").GetROProperty("rows")	
							'Extract the index of row at which the object exist.
							For iCount=0 to bReturn - 1
								If Trim(Lcase(ObjDialog.JavaTable("DisplayedColumns").GetCellData(iCount,"Property"))) = Trim(Lcase(aColNames(iCounter))) then
									'ObjDialog.JavaTable("DisplayedColumns").ExtendRow iCount
									If iCounter > 0 Then
										ObjDialog.JavaTable("DisplayedColumns").ExtendRow iCount
									Else
										ObjDialog.JavaTable("DisplayedColumns").SelectCell iCount, 0
									End If
									bFlag = True
									Exit For
								End If
							Next
							If bFlag = False Then
								Exit Function
							End If
						End If
					Next
					'Click on Add column Button
					Call Fn_Button_Click("Fn_SISW_SE_ColumnManagementOperation", ObjDialog, "RemoveColumn")
					'Click on Apply Button
					Call Fn_Button_Click("Fn_SISW_SE_ColumnManagementOperation", ObjDialog, "Apply")
				End If

				If sAction = "ColumnRemove_SaveConfig" Then
					'Click on Apply Button
					Call Fn_Button_Click("Fn_SISW_SE_ColumnManagementOperation", ObjDialog, "Save")
					wait 1
					bFlag = Fn_SE_SaveColumnConfiguration(sColConfigName,sColConfigDescription)
				End If

				'Click on Close Button
				If sButtonName <> "" Then
					Call Fn_Button_Click("Fn_SISW_SE_ColumnManagementOperation", ObjDialog, sButtonName)
				End If
				wait 1

				If bFlag=True Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_SISW_SE_ColumnManagementOperation passed with case "&sAction)
					Fn_SISW_SE_ColumnManagementOperation = True
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_SISW_SE_ColumnManagementOperation passed with case "&sAction)
					Fn_SISW_SE_ColumnManagementOperation = False
				End If	

		Case "ColumnValidateExists"
				'Count number of rows of Table
				bReturn = ObjDialog.JavaTable("DisplayedColumns").GetROProperty("rows")	
				'Extract the index of row at which the object exist.
				For iCounter=0 to bReturn - 1
					If Trim(Lcase(ObjDialog.JavaTable("DisplayedColumns").GetCellData(iCounter,"Property"))) = Trim(Lcase(aColumnNames)) then
						Fn_SISW_SE_ColumnManagementOperation = True
						Exit For
					End If
				Next
				'Click on Close Button
				If sButtonName <> "" Then
					Call Fn_Button_Click("Fn_SISW_SE_ColumnManagementOperation", ObjDialog, sButtonName)
				End If
				wait 1

		Case "MoveColumnUp","MoveColumnDown","MoveColumnToFirstPosition"
				If aColumnNames <> "" Then
					bFlag = False
					'Count number of rows of Table
					bReturn = ObjDialog.JavaTable("DisplayedColumns").GetROProperty("rows")	
					'Extract the index of row at which the object exist.
					For iCounter=0 to bReturn - 1
						If Trim(Lcase(ObjDialog.JavaTable("DisplayedColumns").GetCellData(iCounter,"Property"))) = Trim(Lcase(aColumnNames)) then
							'ObjDialog.JavaTable("DisplayedColumns").SelectRow iCounter
							ObjDialog.JavaTable("DisplayedColumns").SelectCell iCounter, 0
							If sAction ="MoveColumnToFirstPosition" Then
								iColPos = iCounter
							End If
							bFlag = True
							Exit For
						End If
					Next
					If bFlag = False Then
						Exit Function
					End If
					For iCount = 0 To iColPos - 1
						If sAction = "MoveColumnUp" or sAction ="MoveColumnToFirstPosition" Then
								Call Fn_Button_Click("Fn_SISW_SE_ColumnManagementOperation", ObjDialog, "MoveUp")
								bFlag = True
						ElseIf sAction = "MoveColumnDown" Then
								Call Fn_Button_Click("Fn_SISW_SE_ColumnManagementOperation", ObjDialog, "MoveDown")
								bFlag = True
						End If
					Next
				End If

				'Click on Apply Button
				Call Fn_Button_Click("Fn_SISW_SE_ColumnManagementOperation", ObjDialog, "Apply")
				'Click on Close Button
				If sButtonName <> "" Then
					Call Fn_Button_Click("Fn_SISW_SE_ColumnManagementOperation", ObjDialog, sButtonName)
				End If
				wait 1
				If bFlag=True  Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_SISW_SE_ColumnManagementOperation passed with case "&sAction)
					Fn_SISW_SE_ColumnManagementOperation = True
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_SISW_SE_ColumnManagementOperation failed with case "&sAction)
					 Fn_SISW_SE_ColumnManagementOperation = False
				End If
			
		Case "SaveColumnConfiguration"

			If Not objSaveViewConfig.Exist(5) Then
				Call Fn_ToolBarOperation("ShowDropdownAndSelectWithInstance", "View Menu:3", "Save Column Configuration...")
				Call Fn_ReadyStatusSync(1)
			End If
				
			If objSaveViewConfig.Exist(1) Then		
				'Entering Name for Save View Configuration :- Its mandetory field
				If sColConfigName<>"" Then
					Call Fn_UI_EditBox_Type("Fn_SISW_SE_ColumnManagementOperation",objSaveViewConfig,"Name",sColConfigName)
				End If
		
				'Entering Description for Save View Configuration
				If sColConfigDescription<>"" Then
					Call Fn_Edit_Box("Fn_SISW_SE_ColumnManagementOperation",objSaveViewConfig,"Description",sColConfigDescription)
				End If
				
				'Click on Save Button
				bFlag=Fn_Button_Click("Fn_SISW_SE_ColumnManagementOperation", objSaveViewConfig, "Save")
				wait 1
				If bFlag Then
					Fn_SISW_SE_ColumnManagementOperation=True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Pass:Successfully Save Column Configuration of Name [ "+strName+" ]")
				Else
					Fn_SISW_SE_ColumnManagementOperation=False
				End If						
			End If
		End Select
Set ObjDialog = nothing
End Function


'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name			:	Fn_SISW_SE_ShowTraceabilityMatrixOperation
'@@
'@@    Description				:	Function Used to perform operations on  Show Traceability Matrix window in System Engineering
'@@
'@@    Parameters			   	:	1. sAction			: Action [Type of Attribute Group]
'@@								:	2. bAddSource		: boolean flag to click on Set / Add Source button
'@@							 	:	3. bRemoveSource 	: boolean flag to click on Remove Source button
'@@							 	:	4. bAddTarget 		: boolean flag to click on Set / Add Target button
'@@							 	:	5. bRemoveTarget 	: boolean flag to click on Remove Target button
'@@							 	:	6. bSwitchSourceAndTarget : boolean flag to click on Switch Source And Target button
'@@							 	:	7. sSource 			: used to select Source Element( Source Tab Name and Source Node name is seperated by ~)
'@@							 	:	8. sTarget 			: used to select Target Element( Target Tab Name and Target Node name is seperated by ~)
'@@							 	:	9. sTraceLinkType : Trace Link Type want to select from the list
'@@							 	:	10. bIncludeSubtype : toi include Include Subtype ("ON" OR "OFF")
'@@
'@@    Return Value		   	   	: 	True Or False
'@@
'@@    Pre-requisite			:	System Engineering perspective should be activated
'@@
'@@    Examples					:	Call Fn_SISW_SE_ShowTraceabilityMatrixOperation("Set", "", "", True, "", True, "", "")
'@@
'@@	History				:	
'@@	Developer Name			Date			 Rev. No.	   Changes Done		 Reviewer
'@@------------------------------------------------------------------------------------------------------------------------------
'@@	Ganesh Bhosale			24-Feb-2014		  1.0			created
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public function Fn_SISW_SE_ShowTraceabilityMatrixOperation(sAction, bAddSource, bRemoveSource, bAddTarget, bRemoveTarget, bSwitchSourceAndTarget, sSource, sTarget,sTraceLinkType,bIncludeSubtype )
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SE_ShowTraceabilityMatrixOperation"
	Dim objDialog,rowValues
	Fn_SISW_SE_ShowTraceabilityMatrixOperation = False
	Set objDialog = JavaWindow("SystemsEngineering").JavaWindow("ShowTraceabilityMatrix")

	If Fn_UI_ObjectExist("Fn_SISW_SE_ShowTraceabilityMatrixOperation", objDialog) = False Then
		Call Fn_MenuOperation("Select","View:Show Traceability Matrix...")
		Call Fn_ReadyStatusSync(2)
		If Fn_UI_ObjectExist("Fn_SISW_SE_ShowTraceabilityMatrixOperation", objDialog) = False  Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_SE_ShowTraceabilityMatrixOperation ] Failed to open Show Traceability Matrix window.")
			Exit function
		End If
	End If
	Select Case sAction
		Case "Set"
			If sTraceLinkType <> "" Then
				If inStr(sTraceLinkType, "~")> 0 Then
					call Fn_UI_JavaList_ExtendSelect("Fn_SISW_SE_ShowTraceabilityMatrixOperation", objDialog, "TraceLinkType",sTraceLinkType)
				Else
					call Fn_List_Select("Fn_SISW_SE_ShowTraceabilityMatrixOperation", objDialog, "TraceLinkType",sTraceLinkType)
				End If
			End If
				If bIncludeSubtype<> "" Then
					call Fn_CheckBox_Set("Fn_SISW_SE_ShowTraceabilityMatrixOperation", objDialog, "IncludeSubTypes", bIncludeSubtype)
			End If
			Fn_SISW_SE_ShowTraceabilityMatrixOperation =  Fn_SE_ShowTraceabilityMatrix(sAction, bAddSource, bRemoveSource, bAddTarget, bRemoveTarget, bSwitchSourceAndTarget, sSource, sTarget )
		'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		Case "GetTraceLinkTypes"
				rowValues=Fn_SISW_UI_JavaList_Operations("Fn_SISW_SE_ShowTraceabilityMatrixOperation", "GetContents", objDialog,"TraceLinkType","", "", "")
				If rowValues<> "" Then
					Fn_SISW_SE_ShowTraceabilityMatrixOperation=rowValues
				Else
					Fn_SISW_SE_ShowTraceabilityMatrixOperation=False
				End If
		'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		Case "VerifySourceTarget", "VerifySourceTargetWithoutClose"		'[TC1015-2015072100-20_08_2015-VivekA-NewDevlopment] - Added new cases to verfy Source and Target values
			Fn_SISW_SE_ShowTraceabilityMatrixOperation =  Fn_SE_ShowTraceabilityMatrix(sAction, bAddSource, bRemoveSource, bAddTarget, bRemoveTarget, bSwitchSourceAndTarget, sSource, sTarget )
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_SE_ShowTraceabilityMatrixOperation ] Invalid case [ " & sAction & " ].")
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
	End Select
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [ Fn_SISW_SE_ShowTraceabilityMatrixOperation ] executed successfuly with case [ " & sAction & " ].")
	Set objDialog = Nothing
End Function

'*********************************************************		Function to Create the Snapshot of the assembly	***********************************************************************
'Function Name			: 		Fn_SE_SnapshotCreate

'Description		    :		Create the Snapshot of the assembly

'Parameters			    :	 	1. sAction         :  Name of Action
'								2. dicSnapshotInfo :  Dictionary Object to send Values										

'Return Value		    :      	True \ False

'Pre-requisite		    :		 Assembly should be loded in SE

'Examples				:		 Fn_SE_SnapshotCreate("CreateSnapshot",dicSnapshotInfo)
'													dicSnapshotInfo("SnapshotName")        = "Snap1"
'													dicSnapshotInfo("SnapshotDescription") = "Snapshot Description"
'													dicSnapshotInfo("Button")              = "OK"
'
'History		        :		
'	Developer Name			Date			Rev. No.			Changes Done					Reviewer					Tc Release
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Ankit Nigam		     09-Sep-2015         1.0				Created               		    Vivek A.				TC1015-2015082000-AnkitN-23_09_2015-NewDevelopment
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SE_SnapshotCreate(sAction,dicSnapshotInfo)
	GBL_FAILED_FUNCTION_NAME="Fn_SE_SnapshotCreate"
	Dim strPath , sMenu
	Dim objCreateSnapshot
	
	Fn_SE_SnapshotCreate = False
	 'Select menu "File --> New --> Snapshot"
	 Set objCreateSnapshot = JavaWindow("SystemsEngineering").JavaWindow("SEWindow").JavaDialog("CreateSnapshot")
	 
	If Not objCreateSnapshot.Exist(5) Then
	    strPath = Fn_LogUtil_GetXMLPath("RAC_Menu")
    	sMenu = Fn_GetXMLNodeValue(strPath,"FileNewSnapshot@1")
		Call Fn_MenuOperation("Select",sMenu)
	End If
	
	If objCreateSnapshot.Exist(10) Then
		Select Case sAction
        	'Case for Creating snapshot
	        Case "CreateSnapshot"	            
                If dicSnapshotInfo("SnapshotName") <> "" Then
                	'Input Name
                    Call Fn_Edit_Box("Fn_SE_SnapshotCreate", objCreateSnapshot,"Name",dicSnapshotInfo("SnapshotName"))
               	Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Function [ Fn_SE_SnapshotCreate ] Invalid parameter [ Name ].")
					Fn_PSE_SnapshotCreate = False
					Set objCreateSnapshot = nothing
					Exit Function
				End IF
                wait 1
                If dicSnapshotInfo("SnapshotDescription") <> "" Then
                	'Input Description
                    Call Fn_Edit_Box("Fn_SE_SnapshotCreate",objCreateSnapshot,"Description",dicSnapshotInfo("SnapshotDescription"))
                End IF
                wait 1
                If dicSnapshotInfo("Button") <> "" Then
                    Call Fn_Button_Click("Fn_SE_SnapshotCreate",objCreateSnapshot,dicSnapshotInfo("Button"))
                Else
                    Call Fn_Button_Click("Fn_SE_SnapshotCreate",objCreateSnapshot,"OK")
                End IF	        
    	End Select
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SE_SnapshotCreate ] Successfully created snapshot [" + sName + "]")
		Fn_SE_SnapshotCreate = True
	Else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SE_SnapshotCreate ] Failed to Create snapshot [" + sName + "]")
		Fn_SE_SnapshotCreate = False
	End If	
	Set objCreateSnapshot = nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name			:	Fn_SISW_SE_DeleteTraceLinks
'@@
'@@    Description				:	Function Used to perform operations on  Delete Trace links Dialog in System Engineering
'@@
'@@    Parameters			   	:	1. sAction			: Action [Type of Attribute Group]
'@@								:	2. sButton	    	: Button to be clicked	
'@@								:	3. sTracelink		: Type of Tracelink
'@@								:	4. sCheck   		: verify checkbox(ON/OFF)
'@@							 	:	5. sReserv 			: for future use
'@@							 	         
'@@
'@@    Return Value		   	   	: 	True Or False
'@@
'@@    Pre-requisite			:	Delete Tracelink Dialog must exists.
'@@
'@@    Examples					:	Call Fn_SISW_SE_DeleteTraceLinks("Exists", "No", "","","")
'@@    Examples					:	Call Fn_SISW_SE_DeleteTraceLinks("SelectAll", "", "","","")
'@@    Examples					:	Call Fn_SISW_SE_DeleteTraceLinks("VerifyAllCheckBox", "No", "","ON","")
'@@    Examples					:	Call Fn_SISW_SE_DeleteTraceLinks("Set", "Yes", "Trace Link~BZ_FND_Tracelink","","")
'@@
'@@	History				:	
'@@	Developer Name			Date			 Rev. No.	   Changes Done		 Reviewer
'@@------------------------------------------------------------------------------------------------------------------------------
'@@	Ankit Tewari			4-Mar-2014		  1.0			created
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public function Fn_SISW_SE_DeleteTraceLinks(sAction,sButton,sTracelink,sCheck,sReserv )
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SE_DeleteTraceLinks"
	Dim objDialog,rowCount,objTable,iCounter,iCount
	Dim bFlag,sItemName,iNum

	iCounter=0
	iCount=0
	Fn_SISW_SE_DeleteTraceLinks = False
	Set objDialog = JavaWindow("SystemsEngineering").JavaWindow("DeleteTraceLinks")
	Set objTable=JavaWindow("SystemsEngineering").JavaWindow("DeleteTraceLinks").JavaTable("SelectTraceLinktypes")

	If Fn_UI_ObjectExist("Fn_SISW_SE_DeleteTraceLinks", objDialog) = False  Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_SE_DeleteTraceLinks ] Failed to open Fn_SISW_SE_DeleteTraceLinks window.")
		Exit function
	End If

	Select Case sAction
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "Exists"
			If Fn_UI_ObjectExist("Fn_SISW_SE_DeleteTraceLinks",objDialog) <> False Then
				Fn_SISW_SE_DeleteTraceLinks=True
			Else
				Fn_SISW_SE_DeleteTraceLinks=false
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_SE_DeleteTraceLinks ] Failed to verify Fn_SISW_SE_DeleteTraceLinks window exists.")
				Exit function
			End If			
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "SelectAll","DeselectAll"
			If sAction="SelectAll" Then
				Call Fn_Button_Click("Fn_SISW_SE_DeleteTraceLinks", objDialog, "SelectAll")
			Elseif sAction="DeselectAll" then
				Call Fn_Button_Click("Fn_SISW_SE_DeleteTraceLinks", objDialog, "DeselectAll")
			End If
			Fn_SISW_SE_DeleteTraceLinks = True		
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "VerifyAllCheckBox"
			If Fn_UI_ObjectExist("Fn_SISW_SE_DeleteTraceLinks",objTable) <> False Then
				rowCount=objTable.Object.getItemCount()
				If sCheck="ON" Then
					For iCount=0 to rowCount-1 
						bFlag=objTable.Object.getItem(iCount).getChecked()
						If CBool(bFlag)= true then
							iCounter=iCounter+1
						end if
					Next
					If Cstr(iCounter)=rowCount Then
						Fn_SISW_SE_DeleteTraceLinks = True	
					End If
				Elseif sCheck="OFF" then
					For iCount=0 to rowCount-1 
						bFlag=objTable.Object.getItem(iCount).getChecked()
						If CBool(bFlag)=false then
							iCounter=iCounter+1
						end if
					Next
					If Cstr(iCounter)=rowCount Then
						Fn_SISW_SE_DeleteTraceLinks = True	
					End If
				End IF				
			Else
				Fn_SISW_SE_DeleteTraceLinks=false
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_SE_DeleteTraceLinks ] Failed to verify Fn_SISW_SE_DeleteTraceLinks window exists.")
				Exit function
			End If	
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "Set"
			arrTracelink=split(sTracelink,"~")
			rowCount=objTable.Object.getItemCount()
			For iCount=0 to rowCount-1
				If iCount=0 Then
					sItemName=objTable.Object.getItem(iCount).getText()
				Else
					sItemName=sItemName+"~"+objTable.Object.getItem(iCount).getText()
				End If
			Next
			arrItemName=split(sItemName,"~")
			For iCount=0 to UBound(arrItemName)
				arrItemName(iCount)=trim(replace(arrItemName(iCount),"Count=1",""))
			Next
			For iCount=0 to UBound(arrTracelink)
				For iNum=0 to UBound(arrItemName)
					If arrTracelink(iCount)=arritemName(iNum) Then
						objTable.Object.getItem(iNum).setchecked(true)
						Fn_SISW_SE_DeleteTraceLinks=true
						Exit for
					End If
				Next
			Next
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 		
		Case "ClickButton"
			If Fn_UI_ObjectExist("Fn_SISW_SE_DeleteTraceLinks",objDialog)=True Then
			Call Fn_Button_Click("Fn_SISW_SE_DeleteTraceLinks", objDialog, sButton) 
			Fn_SISW_SE_DeleteTraceLinks=true
		End If

		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_SE_DeleteTraceLinks ] Invalid case [ " & sAction & " ].")
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
	End Select

	If sButton<>"" Then
		If Fn_UI_ObjectExist("Fn_SISW_SE_DeleteTraceLinks",objDialog)=True Then
			Call Fn_Button_Click("Fn_SISW_SE_DeleteTraceLinks", objDialog, sButton) 
			Call Fn_ReadyStatusSync(1)
		End If
	End If

	If  Fn_SISW_SE_DeleteTraceLinks <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [ Fn_SISW_SE_DeleteTraceLinks ] executed successfuly with case [ " & sAction & " ].")
	else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_SE_DeleteTraceLinks ] Failed to executed with case [ " & sAction & " ].")
		Exit function
	End If
	Set objDialog = Nothing
End Function

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name			:	Fn_SE_ReplaceOperation
'@@
'@@    Description				:	Function Used to perform Replace operations on Replace Dialog in System Engineering
'@@
'@@    Parameters			   	:	1. sAction			: Action [ActionName]
'@@								:	2. sBOMLine	    	: BOMLine for selection	
'@@								:	3. dicInfo			: Information for Repace operation 
'@@								:	4. sButton   		: Button Name to click						 	         
'@@
'@@    Return Value		   	   	: 	True Or False
'@@
'@@    							:	Set dicInfo = CreateObject("Scripting.Dictionary")
'@@										dicInfo("SearchCriteria") = "REQ-000401~Reqirement1~REQ-000401-Reqirement1"
'@@										dicInfo("Replace") = "All occurrences in 000571/A;1-Spec1 (View)"
'@@    Examples					:	Call Fn_SE_ReplaceOperation("SearchAndReplace","000335/A;1-ReqSpec(View):REQ-000071/A;1-Req1",dicInfo,"OK")
'@@
'@@		History					:	
'@@ ------------------------------------------------------------------------------------------------------------------------------
'@@		Developer Name			Date			 Rev. No.	   Changes Done		 Reviewer
'@@ ------------------------------------------------------------------------------------------------------------------------------
'@@		Poonam Chopade			18-May-2017		  1.0			created			TC11.3(20170502a)_NewDevelopment_PoonamC_18May2017
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public function Fn_SE_ReplaceOperation(sAction,sReqNode,dicInfo,sButton)
	
	GBL_FAILED_FUNCTION_NAME="Fn_SE_ReplaceOperation"
	Dim objReplaceDialog,sMenu,arrSearch
	
	Fn_SE_ReplaceOperation = False
	
	' Select Requirement node
	If sReqNode <> "" Then
			bFlag = Fn_SE_BOMTableNodeOpeations("Select",sReqNode,"","","")
			If bFlag = False Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SE_ReplaceOperation ] Failed to select Node [ "&sReqNode&" ] in SE.")
				Exit function
			End If
	End If
	
	Set objReplaceDialog = Fn_SISW_SE_GetObject("Replace")
    'check existence of Repace Dialog & perform menu operation to open it
    If Fn_UI_ObjectExist("Fn_SE_ReplaceOperation", objReplaceDialog) = False  Then
		sMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("RAC_Menu"),"EditReplaceExt")
		Call Fn_MenuOperation("Select",sMenu)
		Call Fn_ReadyStatusSync(1)
		
		If Fn_UI_ObjectExist("Fn_SE_ReplaceOperation", objReplaceDialog) = False  Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SE_ReplaceOperation ] Failed to open Replace window.")
				Set objReplaceDialog = Nothing
				Exit function
		End if	
	End If
	
	Select Case sAction
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "SearchAndReplace"
				If dicInfo("SearchCriteria") <> "" Then
						'Click Serach checkbox
						Call Fn_SISW_UI_JavaCheckBox_Operations("Fn_SE_ReplaceOperation","Set",objReplaceDialog,"SearchCheckbox","ON")
						Call Fn_ReadyStatusSync(1)
						
						'Enter serach criteria and selected searched object
						arrSearch = Split(dicInfo("SearchCriteria"),"~")
						Call Fn_OpenByNameOperations("CellDoubleClick",arrSearch(1), arrSearch(0), "","", arrSearch(2))
						Call Fn_ReadyStatusSync(1)
			    End If		
			    
			    ' select Replace Type
			    If dicInfo("Replace") <> "" Then
				    	If instr(dicInfo("Replace"),"All occurrences") > 0 Then
				    		 Call Fn_SISW_UI_JavaRadioButton_Operations("Fn_SE_ReplaceOperation","Set",objReplaceDialog,"ReplaceOptionAll","ON")
				    		 Call Fn_ReadyStatusSync(1)
				    	Else
				    		objReplaceDialog.JavaRadioButton("ReplaceOption_SingleComp").SetTOProperty "attached text",dicInfo("Replace")
				    		Call Fn_SISW_UI_JavaRadioButton_Operations("Fn_SE_ReplaceOperation","Set",objReplaceDialog,"ReplaceOption_SingleComp","ON")
				    		Call Fn_ReadyStatusSync(1)
				    	End If
			    End If
			    Fn_SE_ReplaceOperation = True
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SE_ReplaceOperation ] Invalid case [ " & sAction & " ].")
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
	End Select

	If sButton <> "" Then
		 Call Fn_Button_Click("Fn_SE_ReplaceOperation",objReplaceDialog,sButton) 
		 Call Fn_ReadyStatusSync(1)
	End If
	
	Set objReplaceDialog = Nothing
	
End Function

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name			:	Fn_SE_StructureSearchPanelOperations
'@@
'@@    Description				:	Function Used to perform operations on Structure Search panel in System Engineering
'@@
'@@    Parameters			   	:	1. sAction			: Action [ActionName]
'@@								:	2. sBOMLine	    	: BOMLine for selection	
'@@								:	3. dicInfo			: Information for Structure Search
'@@								:	4. sTab   		    : Tab to Close			 	         
'@@
'@@    Return Value		   	   	: 	True Or False
'@@    							:	Set dicSearchCriteria = CreateObject("Scripting.Dictionary")
'@@										dicInfo("Item ID")="REQ-001301"
'@@
'@@    							:	Set dicInfo = CreateObject("Scripting.Dictionary")
'@@										dicInfo("RMBMenuOption")="Find in other Views:Find in Source View"
'@@										
'@@		Examples				:	Call Fn_SE_StructureSearchPanelOperations("Search","000580/A;1-Spec_01 (View)","dicSearchCriteria","","Close")
'@@    Examples					:	Call Fn_SE_StructureSearchPanelOperations("StructureSearchResultsOpt","REQ-001301/A;1-Req1_01","",dicInfo,"Close")
'@@
'@@		History					:	
'@@ ------------------------------------------------------------------------------------------------------------------------------
'@@		Developer Name			Date			 Rev. No.	   Changes Done		 Reviewer
'@@ ------------------------------------------------------------------------------------------------------------------------------
'@@		Poonam Chopade			18-May-2017		  1.0			created			TC11.3(20170502a)_NewDevelopment_PoonamC_18May2017
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

Public function Fn_SE_StructureSearchPanelOperations(sAction,sReqNode,dicSearchCriteria,dicInfo,sTab)
	GBL_FAILED_FUNCTION_NAME="Fn_SE_StructureSearchPanelOperations"
	Dim objReplaceDialog,sMenu,arrSearch, strMenuFile, strMenu, bFlag, iRowCount, iCnt, arrMenu, sContents, sText
	Fn_SE_StructureSearchPanelOperations=False
	bFlag=False
	
	Set objSE=Fn_SISW_SE_GetObject("SystemsEngineering")
	Set objStructureSearchTable= Fn_SISW_SE_GetObject("StructureSearchResults")
		Select Case sAction
			'---------------------------------------------------------------------------------------
			Case "Search"
								' Select Requirement node
						If sReqNode <> "" Then
								bFlag = Fn_SE_BOMTableNodeOpeations("Select",sReqNode,"","","")
								If bFlag = False Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SE_StructureSearchPanelOperations ] Failed to select Node [ "&sReqNode&" ] in SE.")
									Exit function
								End If
						End If
						
						If Fn_TabFolder_Operation("Exist", "Structure Search", "")=False And Fn_TabFolder_Operation("Exist", "*Structure Search", "")=False Then
							strMenu=Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("RM_Menu"),"ToolsStructureSearch")
							bFlag= Fn_MenuOperation("Select",strMenu)
							If bFlag = False Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SE_StructureSearchPanelOperations ] Failed to select menu [ "&strMenu&" ] in SE.")
								Exit function
							End If
						End If 
						Call Fn_TabFolder_Operation( "DoubleClickTab" , "Structure Search" , "" )
						Wait 1
						
						For Each Elem in dicSearchCriteria
							Select Case Elem
								'----------------------------------------------------------------------------------------------------
								Case "Item ID","Item name"
										objSE.JavaStaticText("StructureSearchItemID").SetTOProperty "label",Elem&":"
										Call Fn_SISW_UI_JavaEdit_Operations("Fn_SE_StructureSearchPanelOperations","Set",objSE,"StructureSearchItemID",dicSearchCriteria(Elem))
										Call Fn_KeyBoardOperation("SendKeys", "{TAB}")
										Wait 2
								'----------------------------------------------------------------------------------------------------
								Case Else
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SE_SteuctureSearchPanelOperations ] Invalid case [ " & Elem & " ].")
							End Select
						Next	
						If sAction="Search" Then
							bFlag=Fn_SISW_UI_JavaButton_Operations("Fn_SE_StructureSearchPanelOperations", "Click", objSE, "Search")
							If bFlag=False Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SE_SteuctureSearchPanelOperations ] Failed to click on Search button.")
						 			Exit Function
						 	End If
						End If
						Call Fn_TabFolder_Operation( "DoubleClickTab" , "*Structure Search" , "" )
						Wait 1
						
					' Close 
						If sTab="Close" Then
						 	bFlag= Fn_TabFolder_Operation("Close", "*Structure Search", "")
						 	If bFlag=False Then
						 			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SE_SteuctureSearchPanelOperations ] Failed to Close tab [ Structure Search].")
						 			Exit Function
						 	End If
						End If
				
				
			Case "StructureSearchResultsOpt"
					bFlag = Fn_TabFolder_Operation("Select", "Structure Search Results", "")
					Call Fn_ReadyStatusSync(1)
					If bFlag = False Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SE_StructureSearchPanelOperations ] Failed to select tab [ Structure Search Results ] in SE.")
						Exit function
					End IF 
					
'					Call Fn_TabFolder_Operation( "DoubleClickTab" , "Structure Search Results" , "" )
'					Wait 1
					
					If dicInfo("Action")="Select" Then
						If sReqNode <> "" Then
							iRowCount = cInt(objStructureSearchTable.GetROProperty("rows"))
								For iCnt = 0 to iRowCount -1
									sText = objStructureSearchTable.object.getItem(iCnt).getData().toString()
									If IsNumeric(sReqNode) Then
										If cstr(sText) = cstr(cint(sReqNode))  Then
											objStructureSearchTable.DeselectRow iCnt
											wait 1
											objStructureSearchTable.selectRow iCnt
											wait 1
											Exit For 
										End If
									ElseIf cstr(sText) = cstr(sReqNode)  Then
											objStructureSearchTable.DeselectRow iCnt
											wait 1
											objStructureSearchTable.selectRow iCnt
											wait 1
											Exit For 
									End If 
								Next  
						End If 
					End If
				
					If dicInfo("Action")="Exist" Then
						If sReqNode <> "" Then
							iRowCount = cInt(objStructureSearchTable.GetROProperty("rows"))
								For iCnt = 0 to iRowCount -1
									sText = objStructureSearchTable.object.getItem(iCnt).getData().toString()
									If cstr(sReqNode)=cstr(sText) Then
									       bFlag=True
										Exit For 
									End If
								Next  
								If bFlag=False Then
									Exit function
								End If
						End If 
					End If
					
					If dicInfo("RMBMenuOption")<>"" Then
						If sReqNode <> "" Then
							iRowCount = cInt(objStructureSearchTable.GetROProperty("rows"))
							For iCnt = 0 to iRowCount -1
								sText = objStructureSearchTable.object.getItem(iCnt).getData().toString()
								If IsNumeric(sReqNode) Then
								 	If cstr(sText) = cstr(cint(sReqNode))  Then
										objStructureSearchTable.DeselectRow iCnt
										wait 1
										objStructureSearchTable.ClickCell iCnt, "BOM Line Name", "RIGHT"
										wait 1
										 Exit for
									End If
								'TC12 20180214 [16-March-2018]Modified by Zaid S. right click operation call as ClickCell was not working - 
								ElseIf cstr(sText) = cstr(sReqNode)  Then
										objStructureSearchTable.DeselectRow iCnt
										wait 1
										objStructureSearchTable.SelectRow iCnt
										wait 1
										objStructureSearchTable.ActivateCell iCnt, "BOM Line Name"
										wait 1
										Call Fn_KeyBoardOperation("SendKey","+{F10}")   'Right click
										wait 1
										 Exit for
								End If			
							Next  
						Else 
							Exit function
						End If 
						
						arrMenu=Split(dicInfo("RMBMenuOption"),":")
						Select Case UBound(arrMenu)
							Case "0"
								sContents = JavaWindow("SystemsEngineering").WinMenu("ContextMenu").BuildMenuPath(arrMenu(0))
							Case "1"
								sContents = JavaWindow("SystemsEngineering").WinMenu("ContextMenu").BuildMenuPath(arrMenu(0),arrMenu(1))
						End Select
						Wait 1
						JavaWindow("SystemsEngineering").WinMenu("ContextMenu").Select sContents
					End If 
				
'					Call Fn_TabFolder_Operation( "DoubleClickTab" , "Structure Search Results" , "" )
'					Wait 1

				' Close 
					If sTab="Close" Then
					 	bFlag= Fn_TabFolder_Operation("Close", "Structure Search Results", "")
					 	If bFlag=False Then
					 			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SE_SteuctureSearchPanelOperations ] Failed to Close tab [ Structure Search].")
					 			Exit Function
					 	End If
					End If
				
			'---------------------------------------------------------------------------------------
			Case Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SE_SteuctureSearchPanelOperations ] Invalid case [ " & sAction & " ].")
			'---------------------------------------------------------------------------------------
		End Select 
		Fn_SE_StructureSearchPanelOperations=True
End Function
'**************************************** Function to remove level of the BOMLine. ***************************************
'
'Function Name	:		   Fn_SE_RemoveLevel  
'
'Description	:  	   Remove level of the BOMLine.
'
'Parameters		 :	   	   1.  String - sAction - Action name
'						   2.  String - sBOMLine
'						   3. String - sKeepSubTree ( ON  /  OFF )
'											
'Return Value	 :	   True / False
'
'Pre-requisite	 :	   Systems Engineering window should be displayed .
'
'Examples		:	  Fn_SE_RemoveLevel("Remove","12345/A;1-ReqSpec (view):REQ-001111/A;1-Req1", "ON")
'
'History		:	  Developer Name		Date			Rev. No.		Changes Done		Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'					  Kaveri P			  14-Jun-2017		 1.0			Created			TC11.3_20170509_NewDevelopment_PoonamC_19Jun2017
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SE_RemoveLevel(sAction,sBOMLine, sKeepSubTree)
GBL_FAILED_FUNCTION_NAME="Fn_SE_RemoveLevel"
	Dim objRemove
	Dim bFlag, sMenu
	Fn_SE_RemoveLevel = False
		
	If sBOMLine <> "" Then
		' selecting BOM Line from BOM Table
		bFlag = Fn_SE_BOMTableNodeOpeations("Select",sBOMLine ,"","","")
		If bFlag = True Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SE_RemoveLevel ] BOM Line [ "+ sBOMLine +" ] selected successfully") 
		Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SE_RemoveLevel ]  Failed to select BOMLine[ "+ sBOMLine +" ]") 
			Exit function
		End If
	End If
	
	Set objRemove = Fn_SISW_SE_GetObject("Remove")
	If Fn_SISW_UI_Object_Operations("Fn_SE_RemoveLevel", "Exist", objRemove, SISW_MINLESS_TIMEOUT) = False Then
			sMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("RAC_Menu"), "EditRemove")
			Call Fn_MenuOperation("Select", sMenu)
			If Fn_SISW_UI_Object_Operations("Fn_SE_RemoveLevel", "Exist", objRemove, SISW_MINLESS_TIMEOUT) = False Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SE_RemoveLevel ]  Failed to open Remove Dialogbox.") 
				Fn_SE_RemoveLevel = False
				Set objRemove = Nothing
				Exit function
			End if
	End If
	
	Select Case sAction
		Case "RemoveLevel"
				'Select checkbox On or OFF to keep subtree while removing node
				if sKeepSubTree <> "" Then
					Call Fn_SISW_UI_JavaCheckBox_Operations("Fn_SE_RemoveLevel", "Set", objRemove.JavaCheckBox("KeepSubtree"), "", sKeepSubTree)
				End if
				'Click on Yes to remove
				Call Fn_Button_Click("Fn_SE_RemoveLevel",objRemove,"Yes")
				Call Fn_ReadyStatusSync(1)
				'click OK 
				If Fn_SISW_UI_Object_Operations("Fn_SE_RemoveLevel", "Exist", objRemove.JavaButton("OK"), SISW_MINLESS_TIMEOUT) Then
					Call Fn_Button_Click("Fn_SE_RemoveLevel",objRemove,"OK")
					Call Fn_ReadyStatusSync(1)
				End if
				Fn_SE_RemoveLevel = True
		Case Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SE_RemoveLevel ] Invalid case.") 
	End Select
	
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SE_RemoveLevel ] executed successfully.") 
	Set objRemove = nothing
	
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name			:	Fn_SE_ProcessHistoryTabOperations
'@@
'@@    Description				:	Function used to verify node existence at Process History view in System Engineering
'@@
'@@    Parameters			   	:	1. sAction			: Action [ActionName]
'@@								:	2. dicInfo			: Information for process history operation 
'@@								:						 	         
'@@
'@@    Return Value		   	   	: 	True Or False
'@@
'@@    							:	Set dicInfo = CreateObject("Scripting.Dictionary")
'@@									dicInfo("SearchCriteria") = "REQ-000401~Reqirement1~REQ-000401-Reqirement1"									
'@@    Examples					:	Call Fn_SE_ProcessHistoryTabOperations("VerifyNodes",dicInfo)
'@@
'@@		History					:	
'@@ ------------------------------------------------------------------------------------------------------------------------------
'@@		Developer Name			Date			 Rev. No.	   Changes Done		 Reviewer
'@@ ------------------------------------------------------------------------------------------------------------------------------
'@@		Poonam Chopade			06-June-2017	 1.0			created			TC11.3(20170509)_NewDevelopment_PoonamC_18May2017
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public function Fn_SE_ProcessHistoryTabOperations(sAction,dicInfo)
	GBL_FAILED_FUNCTION_NAME="Fn_SE_ProcessHistoryTabOperations"
	Dim iCount,iCount1,bFlag,aNode1,sNodeName,objProcessTree,aNode,objTree,sTreeItem
	
	Fn_SE_ProcessHistoryTabOperations = False
	
	Set objProcessTree = Fn_SISW_SE_GetObject("SystemsEngineering")
	Set objProcessTree = objProcessTree.JavaTree("ProcessHistoryTree")
	
    'check existence of ProcessHistoryTree & perform menu operation to open it
    If Fn_TabFolder_Operation("Exist","Process History", "") = False  Then    
		Call Fn_SetView("Teamcenter:Process History")
		Call Fn_ReadyStatusSync(1)
	End If
	
	Select Case sAction
	
		Case "VerifyNodes"
		
				If dicInfo("Nodes") <> "" Then
					aNode = Split(dicInfo("Nodes"),"~")
					If Fn_UI_ObjectExist("Fn_SE_ProcessHistoryTabOperations",objProcessTree)  Then
						Set objTree = objProcessTree.Object
						For iCount = 0 to UBound(aNode)
							bFlag = False
							aNode1 = Split(aNode(iCount),":")
							For iCount1 = 0 To UBound(aNode1)
								Set objTree = objTree.GetItem(iCount1)	
							Next
							sNodeName = aNode1(UBound(aNode1))
							sTreeItem = objTree.getdata().getcomponent().tostring()
							If InStr(trim(sTreeItem),trim(sNodeName)) > 0 Then
								bFlag = True
							End If
							
							If bFlag = False Then
								Fn_SE_ProcessHistoryTabOperations = False
								Set objProcessTree = Nothing
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SE_ProcessHistoryTabOperations ] Node [ " & sNodeName & " dose not exists ].")
								Exit Function
							Else
								Fn_SE_ProcessHistoryTabOperations = True							
							End If
						Next
					End if    
				End If
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SE_ProcessHistoryTabOperations ] Invalid case [ " & sAction & " ].")
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
	End Select
	
	Set objProcessTree = Nothing
	
End Function

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name			:	Fn_SE_SelectConfigurationOperations
'@@
'@@    Description				:	Function used to perform operation on select configuration dialog in sE
'@@
'@@    Parameters			   	:	1. sAction			: Action [ActionName]
'@@							:	2. dicInfo			: Information for  select Configuration Operation
'@@														 	         
'@@
'@@    Return Value		   	   	: 	True Or False
'@@
'@@    							:	Set dicInfo = CreateObject("Scripting.Dictionary")
'@@									dicInfo("RevisionRule") = "Latest Working"
'@@									dicInfo("VariantRule") =  Name of Saved Variant Rule
'@@									dicInfo("Buttton") = "Apply~OK~Cancel"
'@@    Examples					:	Call Fn_SE_SelectConfigurationOperations("Set",dicInfo)
'@@
'@@    History					:	
'@@ ------------------------------------------------------------------------------------------------------------------------------
'@@		Developer Name			Date			 Rev. No.	   Changes Done		 Reviewer
'@@ ------------------------------------------------------------------------------------------------------------------------------
'@@		Archana Dhadiwal	26-July-2017	 		1.0			created			TC11.4(20170626)_NewDevelopment_PoonamC_26July2017
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_SE_SelectConfigurationOperations(sAction,dicInfo)
	GBL_FAILED_FUNCTION_NAME="Fn_SE_SelectConfigurationOperations"
	Dim objConfig,sButton,iCount
	Fn_SE_SelectConfigurationOperations = False
	
	Set objConfig = Fn_SISW_SE_GetObject("SelectConfiguration")
	
	'check existence of ProcessHistoryTree & perform menu operation to open it
	If Fn_UI_ObjectExist("Fn_SE_SelectConfigurationOperations",objConfig)  = False  Then    
		Set objConfig = Nothing
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SE_SelectConfigurationOperations ] Select configuration dialog dose not exists ].")
		Exit Function
	End If
	
	Select Case sAction
		'Case to set configuration settings from dialog.
		Case "Set"
  			' Select Revision Rule		
			If dicInfo("RevisionRule") <> "" Then
				Call Fn_SISW_UI_JavaList_Operations("Fn_SE_SelectConfigurationOperations", "Select", objConfig, "SelectRevisionRule", dicInfo("RevisionRule"), "", "")
			End If
			' Select variant Rule	
			If dicInfo("VariantRule") <> "" Then
				Call Fn_SISW_UI_JavaList_Operations("Fn_SE_SelectConfigurationOperations", "Select", objConfig, "SelectVariantRule", dicInfo("VariantRule"), "", "")
			End If
			'Click on button Apply or OK or Cancel
			If dicInfo("Buttton")<>"" Then
				sButton = Split(dicInfo("Buttton"),"~")
				For iCount = 0 To Ubound(sButton)
					If sButton(iCount) = "Cancel" Then
						If Fn_UI_ObjectExist("Fn_SE_SelectConfigurationOperations",objConfig)  = True	Then
							Call Fn_SISW_UI_JavaButton_Operations("Fn_SE_SelectConfigurationOperations", "Click", objConfig, sButton(iCount))
						End If
					Else
						Call Fn_SISW_UI_JavaButton_Operations("Fn_SE_SelectConfigurationOperations", "Click", objConfig, sButton(iCount))
					End If	
				Next
				Fn_SE_SelectConfigurationOperations = True
			End If
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SE_SelectConfigurationOperations ] Invalid case [ " & sAction & " ].")
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
	End Select
	Set objConfig = Nothing
End Function

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name			:	Fn_SE_BOMCompareReportTableOperations
'@@
'@@    Description				:	his function allows to perform	operations on BOM Compare Report Table.						 	         
'@@
'@@    Return Value		   	   	: 	True Or False
'@@								
'@@    Examples					:	Fn_SE_BOMCompareReportTableOperations("VerifyCellData","REQ-000109","","Part1~A~2->0", "")
'@@
'@@		History					:	
'@@ ------------------------------------------------------------------------------------------------------------------------------
'@@		Developer Name			Date			 Rev. No.	   Changes Done		 Reviewer
'@@ ------------------------------------------------------------------------------------------------------------------------------
'@@		Poonam Chopade			08-Aug-2017	 	  1.0			created			TC11.4(20170626)_NewDevelopment_PoonamC_08Aug2017
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public function Fn_SE_BOMCompareReportTableOperations(sAction, sNodeName, sColName, sValue, sPopupMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_SE_BOMCompareReportTableOperations"
	
	Dim objBOMCompTable,ObjWEmbeddedFrame,iRow,aValues,iCnt,iCnt1,bFlag,sAppData,iRowIndex,iInstance
	Fn_SE_BOMCompareReportTableOperations = False
	
	Set objBOMCompTable = Fn_SISW_SE_GetObject("BOMCompareReportTable")
	Set ObjWEmbeddedFrame = Fn_SISW_SE_GetObject("WEmbeddedFrame")	
	
    'check existence of BOM Compare Table
	ObjWEmbeddedFrame.SetTOProperty "index",1
    If Fn_UI_ObjectExist("Fn_SE_BOMCompareReportTableOperations",objBOMCompTable) = False  Then    
		Set objBOMCompTable = Nothing
		Set ObjWEmbeddedFrame = Nothing
		Fn_SE_BOMCompareReportTableOperations = False
		Exit Function
	End If
	
	Select Case sAction
	
		Case "VerifyCellData"
				If instr(sNodeName,"@") Then
					aValues = Split(sNodeName,"@")
					iInstance = aValues(1)
					sNodeName = aValues(0)
				Else
					iInstance = 1
				End If
				
				iCnt1 = 0
				iRow = objBOMCompTable.GetROProperty("rows")
				For iCnt = 0 To iRow - 1
					bFlag = False
					If instr(objBOMCompTable.GetCellData(iCnt,0),sNodeName) > 0 Then
						iCnt1 = iCnt1 + 1
						If cint(iCnt1) = cint(iInstance) Then
								bFlag = True
								iRowIndex = iCnt
								Exit For									
						End If
					End If  
				Next
				
				'IF Item Id is not found then exit
				If bFlag = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SE_BOMCompareReportTableOperations ] [ Item Id - " & sNodeName & " ] is not found.")
					Set objBOMCompTable = Nothing
					Set ObjWEmbeddedFrame = Nothing
					Fn_SE_BOMCompareReportTableOperations = False
					Exit Function
				End If
				
				'Check Item Name,Qty , Rev according to Item Id 
				aValues = Split(sValue,"~")
				sAppData = objBOMCompTable.GetCellData(iRowIndex,0)
				For iCnt = 0 To UBound(aValues)
						bFlag = False
						If instr(sAppData,aValues(iCnt)) > 0 Then
							bFlag = True
						Else
							bFlag = False
							Exit For	
						End If 
				Next
				If bFlag = False Then
					Fn_SE_BOMCompareReportTableOperations = False
				Else
					Fn_SE_BOMCompareReportTableOperations = True
				End If
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SE_BOMCompareReportTableOperations ] Invalid case [ " & sAction & " ].")
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
	End Select
	
	Set objBOMCompTable = Nothing
	Set ObjWEmbeddedFrame = Nothing
	
End Function

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name			:	Fn_SE_ConfigurationInformationOperations
'@@
'@@    Description				:	Function used to perform operation on Configuration Information in SE
'@@
'@@    Parameters			   	:	1. sAction			: Action [ActionName]
'@@								:	2. dicInfo			: Information for Configuration Operation
'@@														 	         
'@@
'@@    Return Value		   	   	: 	True Or False
'@@
'@@    							:	Set dicInfo = CreateObject("Scripting.Dictionary")
'@@									dicInfo("Revision rule") = "Precise Only"
'@@							
'@@    Examples					:	Call Fn_SE_ConfigurationInformationOperations("Verify",dicInfo,"Close")
'@@
'@@    History					:	
'@@ ------------------------------------------------------------------------------------------------------------------------------
'@@		Developer Name			Date			 Rev. No.	   Changes Done		 Reviewer
'@@ ------------------------------------------------------------------------------------------------------------------------------
'@@		Poonam Chopade	 	26-July-2017	 		1.0			created			TC11.4(20170626)_NewDevelopment_PoonamC_26July2017
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_SE_ConfigurationInformationOperations(sAction,dicInfo,sButton)
	GBL_FAILED_FUNCTION_NAME="Fn_SE_ConfigurationInformationOperations"
	Dim objConfigInfo,sMenu
	Fn_SE_ConfigurationInformationOperations=False
	
	Set objConfigInfo = Fn_SISW_SE_GetObject("ConfigurationInformation")
	
	'check existence of ProcessHistoryTree & perform menu operation to open it
	If Fn_UI_ObjectExist("Fn_SE_ConfigurationInformationOperations",objConfigInfo)  = False  Then
		sMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("RM_Toolbar"),"ShowConfigurationInformation")
		Call Fn_ToolBarOperation("Click", sMenu,"")	
		Call Fn_ReadyStatusSync(2)
		If Fn_UI_ObjectExist("Fn_SE_ConfigurationInformationOperations",objConfigInfo)  = False  Then
			Set objConfigInfo = Nothing
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SE_ConfigurationInformationOperations ] configuration information dialog dose not exists ].")
			Exit Function
		End If
	End If

	Select Case sAction
		'Case to set configuration settings from dialog.
		Case "Verify"
				' Verify Revision Rule	
				If dicInfo("Revision rule") <>"" Then
						' Get Configuration Value
						If trim(dicInfo("Revision rule")) = trim(objConfigInfo.JavaEdit("ConfigurationRule").GetROProperty("value")) Then
							Fn_SE_ConfigurationInformationOperations=True
						Else
							Fn_SE_ConfigurationInformationOperations=False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SE_ConfigurationInformationOperations ] Revision Rule is not equal to ["&trim(dicInfo("Revision rule"))&"]")
						End If
				End If  			
				'Click on button
				If sButton <>"" Then
					Call Fn_SISW_UI_JavaButton_Operations("Fn_SE_ConfigurationInformationOperations", "Click", objConfigInfo,sButton)
				End If
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SE_ConfigurationInformationOperations ] Invalid case [ " & sAction & " ].")
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
	End Select
	Set objConfigInfo = Nothing
End Function
 '*********************************************************		Function to Perform Operation on Traceability Tab	***********************************************************************
'Function Name		:				Fn_SE_TraceabilityTabOperations

'Description		:		 	Perform Operations on "Traceability Tab" 

'Parameters			   :	 			1.sAction: DefiningTable:Properties (First is Table Name Compulsory)
'													 2.sNodeName: Node on which we have to perform operation
'													 3.sNewName:New Name in Property
'													4.sColName:Column Name
'													5.sCellValue:Cell Value  												

'Return Value		   : 				True or False


'Examples				:			            'Fn_SE_TraceabilityTabOperations("ComplyingTable:Select","Traceability","Yes","REQ-000415/A;1-Req2_02:Req2_02->Req3_03","","") 
'												'Fn_SE_TraceabilityTabOperations("ComplyingTable:PopupMenuSelect","Traceability","Yes","REQ-000415/A;1-Req2_02:Req2_02->Req3_03","","Properties...") 
'												'Fn_SE_TraceabilityTabOperations("ComplyingTable:Expand","Traceability","Yes","REQ-000415/A;1-Req2_02:Req2_02->Req3_03","","") 
'												'Fn_SE_TraceabilityTabOperations("DefiningTable:VerifyNode","Traceability","Yes","REQ-000415/A;1-Req2_02:Req2_02->Req3_03","","") 
'												'Fn_SE_TraceabilityTabOperations("Properties:EditCustomeNoteProperty","","","","","")
'												Fn_SE_TraceabilityTabOperations("Properties:VerifyListBoxProperties","","","Custom Notes","CSTMNOTE-731041/A;1-cust1","")
'History			 :		
'								Developer Name							Date						Rev. No.				Build
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'							Jotiba Takkekar						   26/07/2017			              1.0				 TC11.4(20170626.00)
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SE_TraceabilityTabOperations(sAction, sTabName, bMaximiseTab, sNodeName, sReserve,StrMenu)
	Dim iCounter, aMenuList, aAction, iRows, bFlag, sIndex, sNodePath,iEleCount
	Dim objTraceabilityTab, objTraceabilityTabTable,objProperty,ObjCustNoteAddObjectButton,ObjCustNoteEditButton,objCustomnotesProperty
		sIndex=0
		iCounter=0
		iRows=0
		bFlag=False
		
		If sTabName<> "" Then
			Call Fn_SetView ("Systems Engineering:Traceability")
	        	Call Fn_ReadyStatusSync(1)
		End If 
		
		'Maximize the tab if provided Yes
		If bMaximiseTab = "Yes" Then
			Call Fn_SE_RightPanelTabOperations("DoubleClick", "Traceability", "")
			Call Fn_ReadyStatusSync(1)
		End If
		
		'Spliting sAction To retriewe Table name
		aAction=Split(sAction,":")
		
		Set objTraceabilityTab=JavaWindow("SystemsEngineering").JavaWindow("WEmbeddedFrame")
			Select Case aAction(0)
				Case "ComplyingTable"
					Set objTraceabilityTabTable=objTraceabilityTab.JavaTable("ComplyingObject")
					iRows = Fn_Table_GetRowCount("Fn_SE_TraceabilityTabOperations",objTraceabilityTab,"ComplyingObject")
				Case "DefiningTable"
					Set objTraceabilityTabTable=objTraceabilityTab.JavaTable("DefiningObject")
					iRows = Fn_Table_GetRowCount("Fn_SE_TraceabilityTabOperations",objTraceabilityTab,"DefiningObject")
				Case "Properties"
					Set objProperty=objTraceabilityTab.JavaDialog("Properties")	
			End Select 
			
			'Added code if Want to verifyNodes with full TraceLink or only with Rev name by Poonam C
			If sReserve <> "" Then
				If sReserve = "Show Trace Link" Then
				 	 If Fn_UI_ObjectExist("Fn_SE_TraceabilityTabOperations",objTraceabilityTab.JavaButton("ShowTraceLinkBtn")) = True Then
				  		Call Fn_SISW_UI_JavaButton_Operations("Fn_SE_TraceabilityTabOperations", "Click", objTraceabilityTab,"ShowTraceLinkBtn")
						Call Fn_ReadyStatusSync(2)
					 End If
				ElseIf sReserve = "Hide Trace Link" Then
					If Fn_UI_ObjectExist("Fn_SE_TraceabilityTabOperations",objTraceabilityTab.JavaButton("HideTraceLinkBtn")) = True Then
				  		Call Fn_SISW_UI_JavaButton_Operations("Fn_SE_TraceabilityTabOperations", "Click", objTraceabilityTab,"HideTraceLinkBtn")
						Call Fn_ReadyStatusSync(2)
					 End If
				End If
			End If
			
			If sNodeName<>"" Then
				For iCounter = 0 to iRows -1
					objTraceabilityTabTable.SelectRowsRange iCounter,iCounter
					sNodePath=objTraceabilityTabTable.GetCellData(iCounter,0)
					'Checking "sNodeName" present in table or not
					If Trim(sNodePath) = Trim(sNodeName) Then
						sIndex = Cstr(iCounter)
						bFlag=True
						Exit For
					End If
				Next
				If bFlag = False Then
					Fn_SE_TraceabilityTabOperations =False
					Exit Function
				End If
			End If 
		 
		 Select Case aAction(1)
		 		
			Case "Expand"
					objTraceabilityTabTable.SelectRow sIndex
                    objTraceabilityTabTable.ActivateRow sIndex
					Fn_SE_TraceabilityTabOperations =True
					Exit Function
					
			Case "Select"
					objTraceabilityTabTable.SelectRow sIndex
					Fn_SE_TraceabilityTabOperations =True
					Exit Function
					
			Case "VerifyNode"
			     If bFlag=True Then
			     	Fn_SE_TraceabilityTabOperations =True
			     	Exit Function
			     End If
			
			
			 Case "PopupMenuSelect"
			 		objTraceabilityTabTable.SelectRow sIndex
			 		wait 1
			 		objTraceabilityTabTable.ClickCell sIndex,0,"RIGHT"
			 		wait 1
					aMenuList = split(StrMenu, ":",-1,1)
					iCounter = cstr(Ubound(aMenuList))
					Select Case iCounter
						Case "0"
							objTraceabilityTab.JavaMenu("MainMenu").SetTOProperty "label",aMenuList(0)
							objTraceabilityTab.JavaMenu("MainMenu").Select
						Case "1"
							objTraceabilityTab.JavaMenu("MainMenu").SetTOProperty "label",aMenuList(0)
							objTraceabilityTab.JavaMenu("MainMenu").JavaMenu("LevelOne").SetTOProperty "label",aMenuList(1)
							objTraceabilityTab.JavaMenu("MainMenu").JavaMenu("LevelOne").Select
						Case Else											
							Fn_SE_TraceabilityTabOperations = FALSE
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_SE_TraceabilityTabOperations Failed to select Menu "&sNodeName)											
					End Select								
						Fn_SE_TraceabilityTabOperations = True
					Exit Function
					
					
			Case "EditCustomeNoteProperty"	

				If Fn_UI_ObjectExist("Fn_SE_TraceabilityTabOperations",objProperty.JavaStaticText("EmptyProperties"))=True Then
					objProperty.JavaStaticText("EmptyProperties").SetTOProperty "label","Show empty properties..."
					objProperty.JavaStaticText("EmptyProperties").Click 1,1,"LEFT"
					If objProperty.JavaEdit("Description").Exist(15) = False Then
						Fn_SE_TraceabilityTabOperations = False
						Exit function
					End If
				End If
				
				Set objCustomnotesProperty=objProperty.JavaList("CustomNotes")
				Set ObjCustNoteEditButton=objProperty.JavaCheckBox("EditButton")
'				Set ObjCustNoteAddObjectButton=objProperty.JavaButton("Cancel")
				
				If Fn_UI_ObjectExist("Fn_SE_TraceabilityTabOperations",objCustomnotesProperty)=False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_SE_TraceabilityTabOperations Failed to find [Custom Note] property box in Properties Dialog.")											
					Fn_SE_TraceabilityTabOperations=False	
					Exit Function 				
				End If
				
				objProperty.JavaButton("Cancel").SetTOProperty "label","add_16"
				If Fn_UI_ObjectExist("Fn_SE_TraceabilityTabOperations",objProperty.JavaButton("Cancel"))=False Then
					ObjCustNoteEditButton.Click 1,1,"LEFT"	
					If err.number>0 Then
						Fn_SE_TraceabilityTabOperations=False					
					End If						
				End If
				
				If Fn_UI_ObjectExist("Fn_SE_TraceabilityTabOperations",objProperty.JavaButton("Cancel"))=False Then					
						Fn_SE_TraceabilityTabOperations=False
						Exit Function	
				Else
					if Fn_SISW_UI_JavaButton_Operations("Fn_SE_TraceabilityTabOperations", "Click", objProperty,"Cancel")=False Then 
						Fn_SE_TraceabilityTabOperations=False	
						Exit Function
					Else
						Fn_SE_TraceabilityTabOperations=True	
					End If				
				End If				
				
				if Fn_SISW_UI_JavaButton_Operations("Fn_SE_TraceabilityTabOperations", "Click", objProperty,"OK")=False Then 
					Fn_SE_TraceabilityTabOperations=False											
				End If	
				
				Exit Function
				
		Case "VerifyListBoxProperties"
				If Fn_UI_ObjectExist("Fn_SE_TraceabilityTabOperations",objProperty.JavaStaticText("EmptyProperties"))=True Then
					objProperty.JavaStaticText("EmptyProperties").SetTOProperty "label","Show empty properties..."
					objProperty.JavaStaticText("EmptyProperties").Click 1,1,"LEFT"
					If objProperty.JavaEdit("Description").Exist(15) = False Then
						Fn_SE_TraceabilityTabOperations = False
						Exit function
					End If
				End If
				sPropertyName=StrMenu
				objProperty.JavaList("CustomNotes").SetTOProperty "attached text",sPropertyName+":"
				If objProperty.JavaList("CustomNotes").Exist(SISW_MICRO_TIMEOUT) Then
					aValues=Split(sReserve,"~")
					For iCounter=0 to uBound(aValues)
						bFlag=false
						'Verifying value exist in list or not
						'taking item count from list
						iEleCount=Fn_UI_Object_GetROProperty("Fn_SE_TraceabilityTabOperations",objProperty.JavaList("CustomNotes"), "items count")
						For iCount=0 to iEleCount-1
							If objProperty.JavaList("CustomNotes").GetItem(iCount)=aValues(iCounter) Then
								bFlag=true
								Exit for
							End If
						Next
						If bFlag=false Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Value [ " & aValues(iCounter) & " ] is not exist in [ " & sColName & " ] List")
							Exit for
						End If
					Next
					If bFlag=True Then
						Call Fn_Button_Click("Fn_SE_TraceabilityTabOperations",objProperty,"OK")
						Fn_SE_TraceabilityTabOperations=true
					End If
				Else
					Call Fn_Button_Click("Fn_SE_TraceabilityTabOperations",objProperty,"OK")
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"List box [ " & sColName & " ] is not exist on dialog")
				End If				
		 End Select 
		 
		Set objTraceabilityTab=Nothing
		Set objTraceabilityTabTable= Nothing
		Set objCustomnotesProperty=Nothing
		Set ObjCustNoteEditButton=Nothing
		Set ObjCustNoteAddObjectButton=Nothing
				
End Function 
