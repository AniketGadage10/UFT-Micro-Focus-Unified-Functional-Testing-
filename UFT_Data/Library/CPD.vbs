Option Explicit
'Function List
'************************************************************************************************************************************************************************************************************
'000. Fn_SISW_CPD_GetObject(sObjectName)
'001. Fn_CPD_CollaborativeDesignCreate()
'002. Fn_CPD_ContentExplorer()
'003. Fn_CPD_CompnentTabOperations()
'004. Fn_CPD_ItemBasicCreate() - Private function
'005. Fn_CPD_NewPartitionItemCreate()
'006. Fn_CPD_NewWorksetCreate()
'007. Fn_CPD_DesignElementCreate()
'008. Fn_CPD_SubsetDefinitionCreate()
'009. Fn_CPD_RecipeOperations()
'010. Fn_CPD_EffectivityOperations()
'011. Fn_CPD_NavTree_NodeOperation()
'012. Fn_CPD_DateControl()
'013. Fn_CPD_CreatePartitionScheme()
'014. Fn_CPD_CreatePartition()
'015. Fn_CPD_RevisionRuleOperations()
'016. Fn_CPD_ContentSearchOperations()
'017. Fn_CPD_SearchResultTreeOperations()
'018. Fn_CPD_SummaryTabOperations()
'019. Fn_CPD_ColumnManagementOperation()
'020. Fn_CPD_CreateDesignElementWhilePaste()
'021. Fn_CPD_CreatePartitionWhilePaste()
'022. Fn_CPD_Revise()
'023. Fn_CPD_UpdateDesignElement()
'024. Fn_CPD_AttributeGroupCreate()
'025. Fn_SISW_CPD_SubsetDefaultsCreate()
'026. Fn_SISW_CPD_CreatePartitionTemplate()
'027. Fn_SISW_CPD_TargetModelCarryoverOptions()
'028. Fn_SISW_CPD_UpdatePartition()
'029. Fn_SISW_CPD_CreateSubset()
'030. Fn_SISW_CPD_NatTable_TreeTable_GetRowIndex()
'031. Fn_SISW_CPD_VariantNatTableOperations()
'032. Fn_SISW_CPD_ModelCloneandRealizationOptions
'033. Fn_CPD_MarkupSpaceOperations()
'034. Fn_CPD_NewMarkupSpaceCreate()
'035. Fn_SISW_CPD_ModelCloneandRealization_SourcePartitions()
'036. Fn_CPD_CompositeViewMenuOperations()
'037. Fn_CPD_TargetPropertiesOperations()
'038. Fn_CPD_UpdateInstantiationOfModelContentOperations()
'039. Fn_CPD_ModelContentCloneAndInstantiationOperations()
'040. Fn_CPD_SourceTargetTables_Operations()
'041. Fn_CPD_AdvancedAccountabilityCheck_Ops()
'042. Fn_CPD_ViewEditMappings_Operations()
'043. Fn_CPD_Preview4GDModel_Operations()
'044. Fn_CPD_PCA_VariantConfiguration_Operation()
'045. Fn_CPD_PCA_SetRuleDateInConfigurationView()
'046. Fn_CPD_PCA_VariantNatTable_VariantConditions_Operation()
'047. Fn_CPD_LoadVariantRule_Operations()
'048. Fn_CPD_PCA_Save_Variant_Rule()
'*********************************************************	Function List		***********************************************************************
'****************************************    Function to get Object hierarchy ***************************************
'
''Function Name		 	:	Fn_SISW_CPD_GetObject
'
''Description		    :  	Function to get specified Object hierarchy.

''Parameters		    :	1. sObjectName : Object Handle name
								
''Return Value		    :  	Object \ Nothing
'
''Examples		     	:	Fn_SISW_CPD_GetObject("NewBusinessObject")

'History:
'	Developer Name			             Date			          Rev. No.		Reviewer		Changes Done	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Koustubh Watwe		            7-June-2012	                	1.0				
'-----------------------------------------------------------------------------------------------------------------------------------
'	Shrikant Narkhede	          21-June-2012		           1.0				                           Added Case "IncludeChildPartitions"
'-----------------------------------------------------------------------------------------------------------------------------------------------------
'	Ashwini Kumar		 25-Sept-2013		2.0								Externalized object hierarchies
'-----------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_CPD_GetObject(sObjectName)
	Dim sObjectXMLPath
	sObjectXMLPath = Fn_GetEnvValue("User", "AutomationDir") & "\TestData\AutomationXML\ObjectXML\CPD.xml"
	Set Fn_SISW_CPD_GetObject = Fn_SISW_Setup_GetObjectFromXML(sObjectXMLPath, sObjectName)
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name			:	Fn_CPD_CollaborativeDesignCreate
'@@
'@@    Description				:	Function Used to create Collaborative Design
'@@
'@@    Parameters			    :	1. sAction		: Action to be performed
'@@								:	2. sModelID		: Model ID
'@@								:	3. sName		: Name
'@@								:	4. sDescription : description
'@@								:	5. bOpenOnCreate: Boolean value to set Open On Create checkbox  True / False
'@@
'@@    Return Value		   	   	: 	True Or False /  ModelID or False
'@@
'@@    Pre-requisite			:	CPD perspective should be activated.						
'@@
'@@    Examples					:	Call Fn_CPD_CollaborativeDesignCreate("Create", "", "CD1", "CD1", "")
'@@    Examples					:	Call Fn_CPD_CollaborativeDesignCreate("GetErrorMessageOnCreate", "", "CD1", "CD1", "")
'@@
'@@	   History					:	
'@@				Developer Name				Date			Rev. No.	Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe			16-Jan-2012			1.0			Created
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe			20-Feb-2012			1.0			Added case GetErrorMessageOnCreate
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_CPD_CollaborativeDesignCreate(sAction, sModelID, sName, sDescription, bOpenOnCreate)
	GBL_FAILED_FUNCTION_NAME="Fn_CPD_CollaborativeDesignCreate"
	Dim objCDCreate, sType
	Fn_CPD_CollaborativeDesignCreate = False
	Set objCDCreate = JavaWindow("Collaborative Product").JavaWindow("NewBusinessObject")
	
	'If Fn_UI_ObjectExist("Fn_CPD_CollaborativeDesignCreate", objCDCreate) = False Then
	If Fn_SISW_UI_Object_Operations("Fn_CPD_CollaborativeDesignCreate","Exist",objCDCreate,SISW_MICRO_TIMEOUT) = False Then
		Select Case sAction
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - \
			Case "CreateByToolbarButton"
				Call Fn_ToolbarOperation("Click", "Create a new Application Model","")
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
			Case Else
				Call Fn_MenuOperation("Select","File:New:Application Model...")
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		End Select
		
		'If Fn_UI_ObjectExist("Fn_CPD_CollaborativeDesignCreate", objCDCreate) = False Then
		If Fn_SISW_UI_Object_Operations("Fn_CPD_CollaborativeDesignCreate","Exist",objCDCreate,SISW_MIN_TIMEOUT) = False Then
            Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_CPD_CollaborativeDesignCreate ] Failed to opn Collaborative Design window.")
			Set objCDCreate = Nothing
		End IF
	End IF

	Select Case sAction
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "Create", "CreateByToolbarButton", "GetErrorMessageOnCreate"
			wait 2
			
			sType = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("CPD_DisplayName"),"Collaborative Design")
	
			If objCDCreate.JavaTree("BusinessObjectType").exist(1) Then
				' select collaborative design from tree
				objCDCreate.JavaTree("BusinessObjectType").Expand "Complete List"
				wait 1
				objCDCreate.JavaTree("BusinessObjectType").Select "Complete List:" & sType
	
				' click on next
				Call Fn_Button_Click("Fn_CPD_CollaborativeDesignCreate",objCDCreate,"Next" )
				wait(2)
			End If
			' if ModelD is empty
			objCDCreate.JavaStaticText("Field").SetTOProperty "label", "Model ID:"
			If sModelID = "" Then
			     wait (2)
				'	then click on assign
				Call Fn_Button_Click("Fn_CPD_CollaborativeDesignCreate",objCDCreate,"Assign" )
				Call Fn_ReadyStatusSync(5)
				wait (2)
'				Fn_CPD_CollaborativeDesignCreate = objCDCreate.JavaEdit("Field").GetROProperty("value")
				Fn_CPD_CollaborativeDesignCreate=Fn_UI_Object_GetROProperty("",objCDCreate.JavaEdit("Field"), "value")
				sModelID = Fn_CPD_CollaborativeDesignCreate
			Else
				Call Fn_Edit_Box("Fn_CPD_CollaborativeDesignCreate",objCDCreate,"Field",sModelID)
				Fn_CPD_CollaborativeDesignCreate = True
			End If

			'set name
			If sName <> "" Then
				objCDCreate.JavaStaticText("Field").SetTOProperty "label", "Name:"
				'objCDCreate.JavaEdit("Field").Type sName
				Call Fn_Edit_Box("Fn_CPD_CollaborativeDesignCreate",objCDCreate,"Field",sName)
				Call Fn_ReadyStatusSync(5)
			End If

			' set description
			If sDescription <> "" Then
				objCDCreate.JavaStaticText("Field").SetTOProperty "label", "Description:"
				Call Fn_Edit_Box("Fn_CPD_CollaborativeDesignCreate",objCDCreate,"Field",sDescription)
			End If

			' click on next
			'Call Fn_Button_Click("Fn_CPD_CollaborativeDesignCreate",objCDCreate,"Next" )
			'wait(3)
			' if open on create is not empty then click on next
			If bOpenOnCreate <> "" Then
				' set open on create
				If cBool(bOpenOnCreate) Then
					Call Fn_CheckBox_Set("Fn_CPD_CollaborativeDesignCreate",objCDCreate, "OpenOnCreate","ON")
				Else
					Call Fn_CheckBox_Set("Fn_CPD_CollaborativeDesignCreate",objCDCreate, "OpenOnCreate","OFF")
				End If
			End If

			' click on finish
			Call Fn_Button_Click("Fn_CPD_CollaborativeDesignCreate",objCDCreate,"Finish" )
			Call Fn_ReadyStatusSync(5)
			
			If sAction = "GetErrorMessageOnCreate" Then
				Fn_CPD_CollaborativeDesignCreate = False
				If objCDCreate.JavaWindow("Error").Exist(15) Then
'					Fn_CPD_CollaborativeDesignCreate = objCDCreate.JavaWindow("Error").JavaStaticText("ErrorMsg").GetROProperty("value")
					Fn_CPD_CollaborativeDesignCreate = objCDCreate.JavaWindow("Error").JavaEdit("ErrorMsg").GetROProperty("value")
					Call Fn_Button_Click("Fn_CPD_CollaborativeDesignCreate",objCDCreate.JavaWindow("Error"),"OK" )
				End If
			End If

			Call Fn_Button_Click("Fn_CPD_CollaborativeDesignCreate",objCDCreate,"Cancel" )
			IF cBool( bOpenOnCreate) Then
				Call Fn_CPD_CompnentTabOperations("Close", "Content Search","")
				Call Fn_ReadyStatusSync(1)
                Call Fn_CPD_CompnentTabOperations("Close", sName,"")
                Call Fn_ReadyStatusSync(1)
            End If 
			
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_CPD_CollaborativeDesignCreate ] Invalid case [ " & sAction & " ].")
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
	End Select
	
	Wait 5 
	Call Fn_CPD_CompnentTabOperations("Activate",sModelID+";1-"+ sName+" (Content Explorer)","")
	
	If  Fn_CPD_CollaborativeDesignCreate <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [ Fn_CPD_CollaborativeDesignCreate ] executed successfuly with case [ " & sAction & " ].")
	End If
	Set objCDCreate = Nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name			:	Fn_CPD_ContentExplorer
'@@
'@@    Description				:	Function Used to perform operations on Content Explorer
'@@
'@@    Parameters			    :	1. sAction		: Action to be performed
'@@								:	2. sNode		: Node Name
'@@								:	3. sColumn		: Column Name
'@@								:	4. sValue 		: for future use
'@@								:	5. sPopupMenu	: Popup menu
'@@
'@@    Return Value		   	   	: 	True Or False
'@@
'@@    Pre-requisite			:	CPD perspective should be activated.						
'@@
'@@    Examples					:	Call Fn_CPD_ContentExplorer("Select", "CD000015;1-CD1", "", "", "")
'@@    Examples					:	Call Fn_CPD_ContentExplorer("GetChildItems", "CD000015;1-CD1", "", "", "")  - returns ~ separated list of Child Items
'@@    Examples					:	Call Fn_CPD_ContentExplorer("Exist", "CD000015;1-CD1", "", "", "")
'@@    Examples					:	Call Fn_CPD_ContentExplorer("DeSelect", "CD000015;1-CD1", "", "", "")
'@@    Examples					:	Call Fn_CPD_ContentExplorer("MultiSelect", "CD000015;1-CD1~CD000017;1-CD2", "", "", "")
'@@    Examples					:	Call Fn_CPD_ContentExplorer("Expand", "CD000015;1-CD1", "", "", "")
'@@    Examples					:	Call Fn_CPD_ContentExplorer("DoubleClick", "CD000015;1-CD1", "", "", "")
'@@    Examples					:	Call Fn_CPD_ContentExplorer("PopupMenuSelect", "CD000015;1-CD1", "", "", "Copy	Ctrl+C")
'@@    Examples					:	Call Fn_CPD_ContentExplorer("PopupMenuExist", "CD000015;1-CD1", "", "", "Copy	Ctrl+C")
'@@    Examples					:	Call Fn_CPD_ContentExplorer("PopupMenuEnable", "CD000015;1-CD1", "", "", "Copy	Ctrl+C")
'@@    Examples					:	Call Fn_CPD_ContentExplorer("FindCollaborativeDesign", "CD000015;1-CD1", "", "", "")
'@@    Examples					:	Call Fn_CPD_ContentExplorer("CellVerify", "CD000015;1-CD1", "Type", "Collaborative Design", "")
'@@    Examples					:	Call Fn_CPD_ContentExplorer("MultiSelectPopupMenuSelect", "CD000015;1-CD1~CD000017;1-CD2", "", "", "Copy	Ctrl+C")
'@@    Examples					:	Call Fn_CPD_ContentExplorer("GetSelectedNodePath", "", "", "", "")
'@@    Examples					:	Call Fn_CPD_ContentExplorer("Exist_Label", "", "", "CD000005;1-cd", "")
'@@    Examples					:	Call Fn_CPD_ContentExplorer("ClickOn_Label", "", "", "CD000005;1-cd", "")
'@@    Examples					:	Call Fn_CPD_ContentExplorer("PopupMenuSelectOn_Label", "", "", "Partition Scheme Spatial", "Partition Scheme Functional")
'@@    Examples					:	Call Fn_CPD_ContentExplorer("VerifyExistPopupMenuOn_Label", "", "", "Partition Scheme Spatial", "Partition Scheme Functional")
'@@    Examples					:	Call Fn_CPD_ContentExplorer("VerifyEnabledPopupMenOon_Label", "", "", "Partition Scheme Spatial", "Partition Scheme Functional")
'@@    Examples					:	Call Fn_CPD_ContentExplorer("VerifyColumnExists", "", "Type", "", "")
'@@    Examples					:	Call Fn_CPD_ContentExplorer("getfullnodebynodename_ext", "CD000005;1-cd:ItemName", "", "", "")
'@@    Examples					:	Call Fn_CPD_ContentExplorer("exists_contentexplorerlabelpath", "", "", "CD000006;1-SourceCD_70590~000313/A;1-test~wwwww", "")
'@@    Examples					:	Call Fn_CPD_ContentExplorer("exists_contentexplorerlabelpath", "", "","WorksetRev~sBaseRevision~(SUBSET)DesignSubsetName, "")		
'@@	   									 -- use this case if header contains "SUBSET" ,add (SUBSET) before subset name ,i.e. (SUBSET)DesignSubsetName
'@@    Examples					:	Call Fn_CPD_ContentExplorer("celledit", "CD000015;1-CD1", "Logical Designator", "test", "")
'@@
'@@	   History					:	
'@@				Developer Name				Date			Rev. No.	Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe			17-Jan-2012			1.0			Created
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe			19-Jan-2012			1.0			Added case MultiSelectPopupMenuSelect
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe			06-Feb-2012			1.0			Added cases Exist_Label, ClickOn_Label
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe			09-Feb-2012			1.0			Added cases VerifyExistPopupMenuOn_Label, VerifyEnabledPopupMenOon_Label
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Sneha Chavan			15-Feb-2012			1.0			Added cases DoubleClick						Koustubh
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe			08-Mar-2012			1.0			Added cases VerifyColumnExists
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe			08-Mar-2012			1.0			Added cases GetSelectedNodePath
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe			20-Apr-2012			1.0			Added case "popupmenuenable"
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Amit T					31-Jul-2012			1.0			Added Cases "select_basedonsourceobjectname" , "expand_basedonsourceobjectname"
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe			08-Aug-2012			1.0			Modified cases Exist_Label, ClickOn_Label, VerifyExistPopupMenuOn_Label
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Sandeep N			22-Mar-2013			1.0			Modified cases popupmenuenable : Added index property to JavaMenu
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Sandeep N			14-May-2013			1.1			Added Case : cellverifyext
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Sandeep N			23-Jul-2013			1.2			Added Case : getfullnodebynodename
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Ashwini P			29-Apr-2014			1.3			Code modified to handle UFT specific issue for "Nothing" keyword
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Vivek A				06-Nov-2015			1.4			Added case "getfullnodebynodename_ext" to get full node by only Item name				[TC1121-2015102600-06_11_2015-VivekA-NewDevelopment]
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Ankit Nigam			09-Dec-2015			1.4			Added case "exists_contentexplorerlabelpath" to verify the header of context explorer	[TC1121-20151116a00-09_12_2015-AnkitN-NewDevelopment]
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Neelima P			11-Dec-2015			1.4			Added case "verifynumberinnode" Verify Node which contains "( 12 )" in Node name		[TC1121-20151116b-11_12_2015-VivekA-NewDevelopment]
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Vivek A				28-Dec-2015			1.4			Added case "getfirstnodedisplayname_ext" to get first node display name					[TC1122-20151116d-28_12_2015-VivekA-NewDevelopment] 
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Vivek A				19-Aug-2016			1.4			Added case "verifybackgroundcolor" to Verify Background color of row						[TC1123-20160729a-19_08_2016-VivekA-NewDevelopment] 
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Ankit Nigam				18-Aug-2016			1.0			Added Case : celledit
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_CPD_ContentExplorer(sAction, sNode, sColumn, sValue, sPopupMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_CPD_ContentExplorer"
	Dim iTreeIndex, aMenuList, iCnt, oCurrentNode,objTree
	Dim objSelectType, intNoOfObjects, arrStrNode
	Dim aValue, iInstanceCnt, iCounter , sMenuVal
	Dim iInstance,iTempInstance, objShell
	Dim iNodeItemsCount,bFlag
	Dim aSubCondition,  aValueSubCond, aCellDataSubCond
	Dim sVal, sVarCond, sSubVal, sSubVarCond ,sCellData,sNewNode, sCDNode, aCDNode
	Dim arr, sNodepath, appNodeName
	Dim sDisplayNode, aDisplayNode, sVerifyNode, aTreeIndex, sAppText, sNumber
	Fn_CPD_ContentExplorer = False
	iTreeIndex = False
	
	
	If GBL_4GD_EXTRA_TAB_CLOSER=True Then ' Added by Jotiba T
		If Fn_CPD_CompnentTabOperations("Exists", "Content Search","") = True Then
			Call Fn_CPD_CompnentTabOperations("Close","Content Search","")
		End If
		If Instr(sNode,":")>0 Then   
			aCDNode=split(sNode,":")
			sCDNode=aCDNode(lbound(aCDNode)) ' Modified by Chaitali R.
		Else
		 	sCDNode=sNode
		End If
		
		If Fn_CPD_CompnentTabOperations("Exists", sCDNode+" (Content Explorer)","") = True Then
			Call Fn_CPD_CompnentTabOperations("Activate",sCDNode+" (Content Explorer)","")
		End If
		Call Fn_ReadyStatusSync(1)
		GBL_4GD_EXTRA_TAB_CLOSER=False
	End If
	
	Select Case  lcase(sAction)
	Case "findcollaborativedesign", "existde_basedonsourceobjectname", "expand_basedonsourceobjectname", "getdepath_basedonsourceobjectname", "getde_basedonsourceobjectname","getptn_basedontypepartitiondesign","getfullnodebynodename","getfullnodebynodename_ext"
		' do nothing
	Case Else
		If sNode <> "" Then
			If LCase(sAction) = "verifynumberinnode" Then
				sDisplayNode = sNode
			End If
			'---------------------------------------
			' Note : Temporary Solution B'coz Tree is not visible
			' [TC12-20170630.00-17_7_2017-Maintenance]  - Modified Code By JotibaT-JavaObject("TCComponentTab") changed to JavaTab("TCComponentTab")
			If Fn_SISW_UI_Object_Operations("Fn_CPD_ContentExplorer","Exist",JavaWindow("Collaborative Product").JavaTab("TCComponentTab"),SISW_MICRO_TIMEOUT) = True Then
				Dim iHeight, aRect
				aRect = split(JavaWindow("Collaborative Product").JavaTab("TCComponentTab").Object.getBounds().toString() ,",")
				iHeight = cInt( mid(aRect(UBound(aRect)), 1, instr(aRect(UBound(aRect)),"}") -1))
				do While JavaWindow("Collaborative Product").JavaTree("NavTree").Exist(1) = False
					For iCnt = 0 To 10
						JavaWindow("Collaborative Product").JavaTab("TCComponentTab").Click 35, iHeight - 15
					Next 
				loop
			End If
			'---------------------------------------
			if instr(sNode, " ( ")>0 and instr(sNode, " )")>0 Then
				iLen = Cint(instr(sNode, " )"))-cInt(instr(sNode, " ( "))
				sCount = mid(sNode, cInt(instr(sNode, " ( "))+3,iLen-3)	
'				sCount = mid(sNode, cInt(instr(sNode, " )"))+1,Cint(instr(sNode, " ( "))+1)
				sNode = Replace(sNode, " ( " & sCount & " )", "")
			End If
			iTreeIndex = Fn_UI_JavaTreeGetItemPathExt("Fn_CPD_ContentExplorer", JavaWindow("Collaborative Product").JavaTree("NavTree"), sNode, "", "")
		End If
	End Select
	Select Case lcase(sAction)
		Case "verifybackgroundcolor"
				If iTreeIndex <> False Then
					arrStrNode = Split(Replace(iTreeIndex,"#",""), ":")
					Set oCurrentNode = JavaWindow("Collaborative Product").JavaTree("NavTree").Object.getItem(arrStrNode(0))
					For iCnt = 1 To UBound(arrStrNode)
						Set oCurrentNode = oCurrentNode.getItem(arrStrNode(iCnt)) 
					Next
					sRowBackColor = oCurrentNode.getData.getBackgroundColor.tostring
					sRowBackColor = Mid(sRowBackColor, Instr(sRowBackColor,"{"), Instr(sRowBackColor,"}"))
					
					Select Case CStr(UCase(sPopupMenu))
						Case "LIGHTGREEN"
							sColorCode = "{159, 255, 159, 255}"
						Case "LIGHTYELLOW"
							sColorCode = "{255, 255, 128, 255}"
						Case "LIGHTRED"
							sColorCode = "{255, 121, 121, 255}"
						Case "LIGHTBLUE"
							sColorCode = "{183, 219, 255, 255}"
						Case "YELLOWISHORANGE"
							sColorCode = "{254, 190, 95, 255}"
						Case "LIGHTPINK"
							sColorCode = "{253, 217, 220}"
					End Select
					
					If sRowBackColor = sColorCode  Then
						Fn_CPD_ContentExplorer = True
					Else
						Fn_CPD_ContentExplorer = False
					End If
				End If
		Case "select"
				If iTreeIndex <> False Then
					JavaWindow("Collaborative Product").JavaTree("NavTree").select iTreeIndex
					Fn_CPD_ContentExplorer = True
				End If
		Case "getchilditems"
				If iTreeIndex <> False Then
					JavaWindow("Collaborative Product").JavaTree("NavTree").Expand iTreeIndex
					arrStrNode = Split (replace(iTreeIndex,"#",""), ":")
					Set oCurrentNode = JavaWindow("Collaborative Product").JavaTree("NavTree").Object.getItem(arrStrNode(0))	
					For iCnt = 1 to uBound(arrStrNode)
						Set oCurrentNode = oCurrentNode.getItem(arrStrNode(iCnt)) 
					Next
					intNoOfObjects = oCurrentNode.getItemCount()
					
					for iCnt = 0 to intNoOfObjects -1
						If iCnt = 0 then
							Fn_CPD_ContentExplorer = oCurrentNode.getItem(iCnt).getData().toString()
						Else
							Fn_CPD_ContentExplorer = Fn_CPD_ContentExplorer & "~" & oCurrentNode.getItem(iCnt).getData().toString()
						End IF
					Next
				End If
		Case "exist","exists"
				If iTreeIndex <> False Then
					Fn_CPD_ContentExplorer = True
				End If
		Case "deselect"
				If iTreeIndex <> False Then
					JavaWindow("Collaborative Product").JavaTree("NavTree").Deselect iTreeIndex
					Fn_CPD_ContentExplorer = True
				End If
		Case "multiselect"
				aNodes = split(sNode,"~")
				For iCnt = 0 to UBound(aNodes)
					iTreeIndex = Fn_UI_JavaTreeGetItemPathExt("Fn_CPD_ContentExplorer", JavaWindow("Collaborative Product").JavaTree("NavTree"),  aNodes(iCnt), "", "")
					If iTreeIndex <> False Then
						JavaWindow("Collaborative Product").JavaTree("NavTree").ExtendSelect iTreeIndex
						Fn_CPD_ContentExplorer = True
					Else
						Fn_CPD_ContentExplorer = False
						Exit For
					End If
				Next
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
   		Case "multiselectpopupmenuselect"
				Set objTree = JavaWindow("Collaborative Product").JavaTree("NavTree")
				aNodes = split(sNode,"~")
				For iCnt = 0 to UBound(aNodes)
					iTreeIndex = Fn_UI_JavaTreeGetItemPathExt("Fn_CPD_ContentExplorer", JavaWindow("Collaborative Product").JavaTree("NavTree"),  aNodes(iCnt), "", "")
					If iTreeIndex <> False Then
						If iCnt = 0 then
							objTree.Select iTreeIndex
						ElseIf iCnt = UBound(aNodes) Then
							objTree.ExtendSelect iTreeIndex
							wait 1
							objTree.OpenContextMenu iTreeIndex
							wait 1
						Else
							objTree.ExtendSelect iTreeIndex
						End If
						Fn_CPD_ContentExplorer = True
					Else
						Fn_CPD_ContentExplorer = False
						Exit For
					End If
				Next
				
				aMenuList = split(sPopupMenu,":")
				'Select Menu action
				Select Case Ubound(aMenuList)
					Case "0"
						 sPopupMenu = JavaWindow("Collaborative Product").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0))
					Case "1"
						sPopupMenu = JavaWindow("Collaborative Product").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1))
					Case "2"
						sPopupMenu = JavaWindow("Collaborative Product").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1),aMenuList(2))
					Case Else
						Fn_CPD_ContentExplorer = FALSE
						Exit Function
				End Select
				JavaWindow("Collaborative Product").WinMenu("ContextMenu").Select sPopupMenu
				Fn_CPD_ContentExplorer = True
		'------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		' Editable column should be at 2nd position .
		Case "celledit"
				Set objShell = CreateObject("Wscript.Shell")
				Call Fn_UI_ClickJavaTreeCell("Fn_CPD_ContentExplorer", JavaWindow("Collaborative Product"), "NavTree", sNode, sColumn,"LEFT")
				wait 1
				objShell.SendKeys "^A"
				wait 1
				objShell.SendKeys sValue
				wait 2
				JavaWindow("Collaborative Product").JavaTree("NavTree").Click 3,3,"LEFT"
				wait 2
				If LCase(Trim(JavaWindow("Collaborative Product").JavaTree("NavTree").GetColumnValue(iTreeIndex, sColumn))) = LCase(Trim(sValue)) Then
					Set objShell = Nothing
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS : [ Fn_CPD_ContentExplorer ] passes in case [ " & sAction & " ], able to Edit Cell.")
					Fn_CPD_ContentExplorer = True				
				Else	
					Set objShell = Nothing
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : [ Fn_CPD_ContentExplorer ] failed in case [ " & sAction & " ], unable to Edit Cell.")
				End If
		'------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		Case "cellverify"
				If iTreeIndex <> False Then
					sCellData=Trim(JavaWindow("Collaborative Product").JavaTree("NavTree").GetColumnValue(iTreeIndex,sColumn))

					aValue = replace(sValue, "(", "")
					aValue = replace(aValue, ")", "")
					aValue = split(trim(aValue), "&")
					
					aCellData = replace(sCellData, "(", "")
					aCellData = replace(aCellData, ")", "")
					aCellData = split(trim(aCellData), "&")

					For each sVal in aValue
						bFlag = False
						sVal = replace(sVal, " ", "")
						For each sVarCond in aCellData
							sVarCond = replace(sVarCond, " ", "")
							If instr(sVal,"|") > 0 Then
								If instr(sVarCond,"|") > 0 Then
									aValueSubCond  = split(sVal,"|")
									aCellDataSubCond = split(sVarCond,"|")
									For each sSubVal in aValueSubCond
										bFlag = False 
										For each sSubVarCond in aCellDataSubCond
											If sSubVal = sSubVarCond Then
												bFlag = True
												Exit for
											End If
										Next
										If bFlag = False Then Exit for
									Next
									If bFlag Then Exit for
								Else
									'do nothing
								End If
							Else
								If instr(sVarCond,"|") > 0 Then
									'do nothing 
								Else
									If sVal = sVarCond Then
										bFlag = True
										Exit for
									End If
								End If
							End If
						Next
						If bFlag = False then exit for
					Next
					Fn_CPD_ContentExplorer= bFlag
				End If
		'------------------- Get Selected Node with Path---------------
		'[TC1121-20151116a-08_12_2015-VivekA-NewDevelopment] - Added to handle Baseline node which returns "( Open )" or "( Closed )" in selected name
		Case "getselectednodepath","getselectednodepath_ext"
			If  JavaWindow("Collaborative Product").JavaTree("NavTree").Object.getSelectionCount() <> 0 Then 
				Set oCurrentNode = JavaWindow("Collaborative Product").JavaTree("NavTree").Object.getFocusItem()
	'			If lcase(typename(oCurrentNode)) <> "nothing" then
	'-------------------------Modifications for UFT issue----------------------------------
				If Environment.Value("ProductName") = sQTPProductName OR Environment.Value("ProductName") = sUFTProductName Then
					If IsObject(oCurrentNode) then
						If lCase(sAction) = "getselectednodepath_ext" Then
							Fn_CPD_ContentExplorer = Fn_SISW_UI_JavaTree_GetSanitizedNodeName(oCurrentNode)
						Else
							Fn_CPD_ContentExplorer = oCurrentNode.getData().toString() 
						End If
						
						Do while IsObject(oCurrentNode.getParentItem())
							Set oCurrentNode = oCurrentNode.getParentItem()
							Fn_CPD_ContentExplorer = oCurrentNode.getData().toString() & ":" & Fn_CPD_ContentExplorer 
						Loop
					End If
				Else
					If not oCurrentNode is Nothing then
						Fn_CPD_ContentExplorer = oCurrentNode.getData().toString() 
						Do while lcase(typename(oCurrentNode.getParentItem())) <> "nothing"
							Set oCurrentNode = oCurrentNode.getParentItem()
							Fn_CPD_ContentExplorer = oCurrentNode.getData().toString() & ":" & Fn_CPD_ContentExplorer 
						Loop
					End If
				End If
			Else
				Fn_CPD_ContentExplorer =False
			End If
		'--------------------------End-----------------------------------------------------------			
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "expand"
				If iTreeIndex <> False Then
					JavaWindow("Collaborative Product").JavaTree("NavTree").Expand iTreeIndex
					Fn_CPD_ContentExplorer = True
				End If
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "getindex"
				Fn_CPD_ContentExplorer = -1
				If iTreeIndex <> False Then
					Fn_CPD_ContentExplorer = Fn_UI_getJavaTreeIndex(JavaWindow("Collaborative Product").JavaTree("NavTree"), sNode)
				End If
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "doubleclick"
			If iTreeIndex <> False Then
				'JavaWindow("Collaborative Product").JavaTree("NavTree").Select iTreeIndex
				JavaWindow("Collaborative Product").JavaTree("NavTree").Activate iTreeIndex
				wait 5
				Call Fn_KeyBoardOperation("SendKeys", "{ENTER}")
				Call Fn_ReadyStatusSync(1)
				'-----Modified Chaitali R.------
				If Instr(sNode,":")>0 Then   ' Added by Jotiba T
					sNewNode=split(sNode,":")
					sNode=sNewNode(ubound(sNewNode))
				End If
								
				If Fn_CPD_CompnentTabOperations("Exists", sNode+" (Content Explorer)","") = True Then
					Call Fn_CPD_CompnentTabOperations("Activate",sNode+" (Content Explorer)","")
				End If
				Call Fn_ReadyStatusSync(1)
				Fn_CPD_ContentExplorer = True
			End If
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "popupmenuselect"
				If iTreeIndex <> False Then
					JavaWindow("Collaborative Product").JavaTree("NavTree").Select iTreeIndex
					wait 2
					JavaWindow("Collaborative Product").JavaTree("NavTree").OpenContextMenu iTreeIndex
					wait 2
                    aMenuList = split(sPopupMenu,":")
					'Select Menu action
					Select Case Ubound(aMenuList)
						Case "0"
							 sPopupMenu = JavaWindow("Collaborative Product").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0))
						Case "1"
							sPopupMenu = JavaWindow("Collaborative Product").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1))
						Case "2"
							sPopupMenu = JavaWindow("Collaborative Product").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1),aMenuList(2))
						Case Else
							Fn_CPD_ContentExplorer = FALSE
							Exit Function
					End Select
					JavaWindow("Collaborative Product").WinMenu("ContextMenu").Select sPopupMenu
					Fn_CPD_ContentExplorer = True
				End If
		'----------------------------------------------------------------------------------------------------------
		Case "popupmenuexist"
				If iTreeIndex <> False Then
					JavaWindow("Collaborative Product").JavaTree("NavTree").Select iTreeIndex
					wait 1
					JavaWindow("Collaborative Product").JavaTree("NavTree").OpenContextMenu iTreeIndex
					wait 1
                    aMenuList = split(sPopupMenu,":")
					'Select Menu action
					Select Case Ubound(aMenuList)
						Case "0"
							 sPopupMenu = JavaWindow("Collaborative Product").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0))
						Case "1"
							sPopupMenu = JavaWindow("Collaborative Product").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1))
						Case "2"
							sPopupMenu = JavaWindow("Collaborative Product").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1),aMenuList(2))
						Case Else
							Fn_CPD_ContentExplorer = FALSE
							Exit Function
					End Select

					If JavaWindow("Collaborative Product").WinMenu("ContextMenu").GetItemProperty (sPopupMenu,"Exists") = True Then
						Fn_CPD_ContentExplorer = True
					Else
						Fn_CPD_ContentExplorer = False
					End If
				End If
				Call Fn_KeyBoardOperation("SendKeys", "{ESC}")
		'----------------------------------------------------------------------------------------------------------
		Case "popupmenuenable" 
				If iTreeIndex <> False Then
					JavaWindow("Collaborative Product").JavaTree("NavTree").Select iTreeIndex
					Wait 1
					JavaWindow("Collaborative Product").JavaTree("NavTree").OpenContextMenu iTreeIndex
					Wait 1

                    aMenuList = split(sPopupMenu,":")
					'Select Menu action
					Select Case Ubound(aMenuList)
						Case "0"
'							 sPopupMenu = JavaWindow("Collaborative Product").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0))
							 sMenuVal = Cint(JavaWindow("Collaborative Product").JavaMenu("label:="&aMenuList(0)&"","index:=0").GetROProperty("enabled"))
						Case "1"
'							sPopupMenu = JavaWindow("Collaborative Product").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1))
							sMenuVal = Cint(JavaWindow("Collaborative Product").JavaMenu("label:="&aMenuList(0)&"","index:=0").JavaMenu("label:="&aMenuList(1)&"","index:=0").GetROProperty("enabled"))
						Case "2"
'							sPopupMenu = JavaWindow("Collaborative Product").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1),aMenuList(2))
							sMenuVal = Cint(JavaWindow("Collaborative Product").JavaMenu("label:="&aMenuList(0)&"","index:=0").JavaMenu("label:="&aMenuList(1)&"","index:=0").JavaMenu("label:="&aMenuList(2)&"","index:=0").GetROProperty("enabled"))
						Case Else
							Fn_CPD_ContentExplorer = FALSE
							Exit Function
					End Select

'					If cInt(JavaWindow("Collaborative Product").WinMenu("ContextMenu").GetItemProperty (sPopupMenu,"Enabled")) = 1 Then
'						Fn_CPD_ContentExplorer = True
'					Else
'						Fn_CPD_ContentExplorer = False
'					End If

					If sMenuVal = 1 Then
						Fn_CPD_ContentExplorer = True
					Else
						Fn_CPD_ContentExplorer = False
					End If
				
				End If
				Call Fn_KeyBoardOperation("SendKeys", "{ESC}")

		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "multiselectpopupmenuselect"
				aNodes = split(sNode,"~")
				For iCnt = 0 to UBound(aNodes) -1
					iTreeIndex = Fn_UI_JavaTreeGetItemPathExt("Fn_CPD_ContentExplorer", JavaWindow("Collaborative Product").JavaTree("NavTree"),  aNodes(iCnt), "", "")
					If iTreeIndex <> False Then
						JavaWindow("Collaborative Product").JavaTree("NavTree").ExtendSelect iTreeIndex
						Fn_CPD_ContentExplorer = True
					Else
						Fn_CPD_ContentExplorer = False
						Exit Function
					End If
				Next
				iTreeIndex = Fn_UI_JavaTreeGetItemPathExt("Fn_CPD_ContentExplorer", JavaWindow("Collaborative Product").JavaTree("NavTree"),  aNodes(iCnt), "", "")
				If iTreeIndex <> False Then
					JavaWindow("Collaborative Product").JavaTree("NavTree").ExtendSelect iTreeIndex
					wait 1
					JavaWindow("Collaborative Product").JavaTree("NavTree").OpenContextMenu iTreeIndex
					wait 1
					aMenuList = split(sPopupMenu,":")
					'Select Menu action
					Select Case Ubound(aMenuList)
						Case "0"
							 sPopupMenu = JavaWindow("Collaborative Product").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0))
						Case "1"
							sPopupMenu = JavaWindow("Collaborative Product").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1))
						Case "2"
							sPopupMenu = JavaWindow("Collaborative Product").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1),aMenuList(2))
						Case Else
							Fn_CPD_ContentExplorer = FALSE
							Exit Function
					End Select
					JavaWindow("Collaborative Product").WinMenu("ContextMenu").Select sPopupMenu
					Fn_CPD_ContentExplorer = True
				End If
				
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "findcollaborativedesign"
				'swapnil : 24-Jan-13 - added activate call.
				If cInt(JavaWindow("Collaborative Product").JavaTree("NavTree").Object.getItemCount()) > 0 Then
					Wait 1
					JavaWindow("Collaborative Product").JavaTree("NavTree").Deselect cint(iRootindex)
				End If
				JavaWindow("Collaborative Product").Activate
				Call Fn_Edit_Box("Fn_CPD_ContentExplorer", JavaWindow("Collaborative Product"), "FindCollaborativeDesign",sNode)
				wait 1 
				JavaWindow("Collaborative Product").JavaEdit("FindCollaborativeDesign").Click 1,1
				wait 1
				Call Fn_KeyBoardOperation("SendKeys", "{ENTER}")
				JavaWindow("Collaborative Product").JavaEdit("FindCollaborativeDesign").Activate
				wait 1
				'Swapnil : 29-Jan-2013 -added the call to maximize the window.
				JavaWindow("Collaborative Product").Maximize
				Fn_CPD_ContentExplorer = True
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "exist_label","exists_label"
					aValue = split(sValue,"@")
					if uBound(aValue) = 1 then
						JavaWindow("Collaborative Product").JavaObject("ImageHyperlink").SetTOProperty "Index",(cInt(aValue(1)) -1)
					End If
					JavaWindow("Collaborative Product").JavaObject("ImageHyperlink").SetTOProperty "developer name", trim(aValue(0))
					Fn_CPD_ContentExplorer = JavaWindow("Collaborative Product").JavaObject("ImageHyperlink").Exist(10)
					JavaWindow("Collaborative Product").JavaObject("ImageHyperlink").SetTOProperty "developer name", ""
					IF NOT(Fn_CPD_ContentExplorer) Then
						Set objSelectType = Description.Create()
						objSelectType("Class Name").value = "JavaObject"
						objSelectType("toolkit class").value = "org.eclipse.ui.forms.widgets.ImageHyperlink"
						Set  intNoOfObjects = JavaWindow("Collaborative Product").ChildObjects(objSelectType)
						For iCnt = 0 to intNoOfObjects.count-1
								IF intNoOfObjects(iCnt).Object.getText() = trim(aValue(0)) Then
									Fn_CPD_ContentExplorer =True
									Exit for
								End IF
						Next
					End If
					If Fn_CPD_ContentExplorer = False Then
						JavaWindow("Collaborative Product").JavaStaticText("Search_Type").SetTOProperty "label",aValue(0)
						If JavaWindow("Collaborative Product").JavaStaticText("Search_Type").Exist(5) Then
							Fn_CPD_ContentExplorer = True
						End If
					End If
					JavaWindow("Collaborative Product").JavaObject("ImageHyperlink").SetTOProperty "Index", 0
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "clickon_label"
					aValue = split(sValue,"@")
					if uBound(aValue) = 1 then
						JavaWindow("Collaborative Product").JavaObject("ImageHyperlink").SetTOProperty "Index",(cInt(aValue(1)) -1)
					End If
					JavaWindow("Collaborative Product").JavaObject("ImageHyperlink").SetTOProperty "developer name", trim(aValue(0))
					If JavaWindow("Collaborative Product").JavaObject("ImageHyperlink").Exist(10) Then
						xAxis = JavaWindow("Collaborative Product").JavaObject("ImageHyperlink").GetROProperty( "abs_x")
						JavaWindow("Collaborative Product").JavaObject("ImageHyperlink").Click 1, 1 ,"LEFT" 
						Fn_CPD_ContentExplorer = True
					End If
					JavaWindow("Collaborative Product").JavaObject("ImageHyperlink").SetTOProperty "developer name", ""
					
					If Fn_CPD_ContentExplorer = False Then
						Set objSelectType = Description.Create()
						objSelectType("Class Name").value = "JavaObject"
						objSelectType("toolkit class").value = "org.eclipse.ui.forms.widgets.ImageHyperlink"
						Set  intNoOfObjects = JavaWindow("Collaborative Product").ChildObjects(objSelectType)
						For iCnt = 0 to intNoOfObjects.count-1
								IF intNoOfObjects(iCnt).Object.getText() = trim(aValue(0)) Then
									intNoOfObjects(iCnt).Click 1, 1 ,"LEFT" 
									Fn_CPD_ContentExplorer =True
									Exit for
								End IF
						Next
					End If
					JavaWindow("Collaborative Product").JavaObject("ImageHyperlink").SetTOProperty "Index", 0
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "popupmenuselecton_label"
						aValue = split(sValue,"@")
						iCounter = 0
						iInstanceCnt = 0
						If uBound(aValue) = 1 then
							iInstanceCnt =  cInt(aValue(1)) -1
						End If
						Set objSelectType = Description.Create()
						'objSelectType("Class Name").value = "JavaObject"
						objSelectType("to_class").value = "JavaObject"
						objSelectType("toolkit class").value = "org.eclipse.ui.forms.widgets.ImageHyperlink"
						Set  intNoOfObjects = JavaWindow("Collaborative Product").ChildObjects(objSelectType)
						For iCnt = 0 to intNoOfObjects.count-1
								If intNoOfObjects(iCnt).Object.getText() = trim(aValue(0)) Then
									If iCounter = iInstanceCnt Then
										Call Fn_SISW_UI_DeviceReplayObjectClick("Fn_CPD_ContentExplorer", intNoOfObjects(iCnt))
										'intNoOfObjects(iCnt).Click 3, 3 ,"LEFT" 
										wait 3
										sPopupMenu = JavaWindow("Collaborative Product").WinMenu("ContextMenu").BuildMenuPath(sPopupMenu)
										JavaWindow("Collaborative Product").WinMenu("ContextMenu").Select sPopupMenu
										Fn_CPD_ContentExplorer =True
										Exit for
									Else
										iCounter = iCounter + 1
									End If
								End IF
						Next
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "verifyexistpopupmenuon_label", "verifyenabledpopupmenuon_label"
					Dim sProperty
					aValue = split(sValue,"@")
					iCounter = 0
					iInstanceCnt = 0
					If uBound(aValue) = 1 then
						iInstanceCnt =  cInt(aValue(1)) -1
					End If
					Set objSelectType = Description.Create()
					objSelectType("Class Name").value = "JavaObject"
					objSelectType("toolkit class").value = "org.eclipse.ui.forms.widgets.ImageHyperlink"
					Set  intNoOfObjects = JavaWindow("Collaborative Product").ChildObjects(objSelectType)
					For iCnt = 0 to intNoOfObjects.count-1
							If intNoOfObjects(iCnt).Object.getText() = trim(aValue(0)) Then
								If iCounter = iInstanceCnt Then
									intNoOfObjects(iCnt).Click 1, 1 ,"LEFT" 
									wait 1
									sPopupMenu = JavaWindow("Collaborative Product").WinMenu("ContextMenu").BuildMenuPath(sPopupMenu)
									Select Case lcase(sAction)
										Case "verifyexistpopupmenuon_label"
											sProperty = "Exists"
										Case "verifyenabledpopupmenuon_label"
											sProperty = "Enabled"
									End Select
									Fn_CPD_ContentExplorer =  JavaWindow("Collaborative Product").WinMenu("ContextMenu").CheckItemProperty(sPopupMenu, sProperty, True, 20)
									Call Fn_KeyBoardOperation("SendKeys", "{ESC}")
									Exit for
								Else
									iCounter = iCounter + 1
								End If
							End IF
					Next
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "verifycolumnexists"
				
				intNoOfObjects = cInt(JavaWindow("Collaborative Product").JavaTree("NavTree").GetROProperty("columns_count"))
				Fn_CPD_ContentExplorer = false
				For iCnt = 0 to intNoOfObjects - 1
					If JavaWindow("Collaborative Product").JavaTree("NavTree").GetColumnHeader(iCnt) = sColumn Then
						Fn_CPD_ContentExplorer = True
						Exit for
					End If
				Next

		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "select_basedonsourceobjectname" , "expand_basedonsourceobjectname"
		
'			Set objTree = JavaWindow("Collaborative Product").JavaTree("NavTree")
'
'			For iCnt = 1 to cint(objTree.GetROProperty("items count")) - 1
'				'Create node
'				sRetNodePath = objTree.GetItem( iCnt )
'				'Get value at this node and column"source Object Name"
'				sTreeVal = objTree.GetColumnValue( sRetNodePath , "Source Object Name" )
'				If sNode = sTreeVal Then
'					'Select/Expand node at this index
'					sRetNodePath = Fn_UI_JavaTreeGetItemPathExt("Fn_CPD_ContentExplorer", objTree, sRetNodePath, "", "")
'					If sRetNodePath <> False Then
'						Select Case sAction
'							Case "select_basedonsourceobjectname"
'								objTree.Select sRetNodePath
'							Case "expand_basedonsourceobjectname"
'								objTree.Expand sRetNodePath
'						End Select
'						Fn_CPD_ContentExplorer = True
'						Exit For
'					End If
'				End If
'			Next
			Set objTree = JavaWindow("Collaborative Product").JavaTree("NavTree").Object.getItem(0)
			sNodepath =  Fn_CPD_ContentExplr_GetNodePathFromColumnValue(objTree, "Source Object Name", sNode,"#0")
			sNodepath = replace(sNodepath, "True", "")
			set objTree  = JavaWindow("Collaborative Product").JavaTree("NavTree")
			If sNodepath <> False Then
				Select Case sAction
					Case "select_basedonsourceobjectname"
						objTree.Select sNodepath
					Case "expand_basedonsourceobjectname"
						objTree.Expand sNodepath
				End Select
				Fn_CPD_ContentExplorer = True
			End If
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "getde_basedonsourceobjectname" , "existde_basedonsourceobjectname" , "getdepath_basedonsourceobjectname"  'Parent node should be EXPANDED
		
'			Set objTree = JavaWindow("Collaborative Product").JavaTree("NavTree")
'
'			For iCnt = 1 to cint(objTree.GetROProperty("items count")) - 1
'				'Create node
'				sRetNodePath = objTree.GetItem( iCnt )
'				'Get value at this node and column"source Object Name"
'				sTreeVal = objTree.GetColumnValue( sRetNodePath , "Source Object Name" )
'				If sNode = sTreeVal Then
'					Select Case sAction
'						Case "existde_basedonsourceobjectname"
'								Fn_CPD_ContentExplorer = True
'						Case "getde_basedonsourceobjectname"
'								'Split node on ":"
'								aValue = Split( sRetNodePath , ":" )
'								Fn_CPD_ContentExplorer = aValue( Ubound(aValue) )
'						Case "getdepath_basedonsourceobjectname"
'								Fn_CPD_ContentExplorer = sRetNodePath
'						End Select
'					Exit For
'				End If
'			Next

			Set objTree = JavaWindow("Collaborative Product").JavaTree("NavTree").Object.getItem(0)
			
			bFlag =  Fn_CPD_ContentExplr_GetNodePathFromColumnValue(objTree, "Source Object Name", sNode,"#0")
			'[TC1121-2015101900-04_11_2015-VivekA-Maintenance] - Added by Priyanka K, if the node note present in tree
			If bFlag = False AND sAction = "getde_basedonsourceobjectname" Then
				Fn_CPD_ContentExplorer = False
				set objTree = nothing
				Exit Function 
			End If
			'--------------------------------------------------
			bFlag = replace(bFlag, "#", "")
			bFlag = replace(bFlag, "True", "")
			arr = split(bFlag, ":")
			set objTree  = JavaWindow("Collaborative Product").JavaTree("NavTree").Object
			For icounter = 0 To ubound(arr)
				Set objTree = objTree.getItem(cint(arr(iCounter)))
				If icounter = 0 Then
					sNodepath = objTree.getData.tostring()
				 else
				 	sNodepath =sNodepath + ":" + objTree.getData.tostring()
				End If
			Next
				Select Case sAction
					Case "existde_basedonsourceobjectname"
							Fn_CPD_ContentExplorer = True
					Case "getde_basedonsourceobjectname"
							'Split node on ":"
							aValue = Split( sNodepath , ":" )
							Fn_CPD_ContentExplorer = aValue( Ubound(aValue) )
					Case "getdepath_basedonsourceobjectname"
							Fn_CPD_ContentExplorer = sNodepath
				End Select
        ' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
        Case "getfirstnodedisplayname","getfirstnodedisplayname_ext"				'[TC1122-20151116d-28_12_2015-VivekA-NewDevelopment] - Added case to get first node display name
        	If sAction = "getfirstnodedisplayname_ext" Then
        		Fn_CPD_ContentExplorer = Fn_SISW_UI_JavaTree_GetSanitizedNodeName(JavaWindow("Collaborative Product").JavaTree("NavTree").Object.getItem(0))
        	Else
        		Fn_CPD_ContentExplorer = JavaWindow("Collaborative Product").JavaTree("NavTree").Object.getItem(0).getText()
        	End If            
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "cellverifyext"
				If iTreeIndex <> False Then
					If instr(1,JavaWindow("Collaborative Product").JavaTree("NavTree").GetColumnValue(iTreeIndex,sColumn),sValue) Then
						Fn_CPD_ContentExplorer = True
					End If
				End If
	    ' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "getptn_basedontypepartitiondesign"

'				Set objTree = JavaWindow("Collaborative Product").JavaTree("NavTree")
'				sValue=split(sValue,"~")
'				If ubound(sValue)=1 Then
'					iInstance=cint(sValue(1))
'				Else
'					iInstance=1
'				End If
'				iTempInstance=1
'				For iCnt = 1 to cint(objTree.GetROProperty("items count")) - 1
'					'Create node
'					sRetNodePath = objTree.GetItem( iCnt )
'					sNode=split(sRetNodePath,":")
'					If instr(1,sNode(ubound(sNode)),sValue(0)) Then
'						'Get value at this node and column"source Object Name"
'						sTreeVal = objTree.GetColumnValue( sRetNodePath , "Type" )
'						If "Partition Design" = sTreeVal Then
'							If iInstance=iTempInstance Then
'								Fn_CPD_ContentExplorer = sNode(ubound(sNode))		
'								Exit For
'							End If
'							iTempInstance=iTempInstance+1
'						End If
'					End If
'				Next

				Dim ispCnt, iPtnCnt, objCDNode, objPtnNode, objSubPtnCnt
				Set objTree = JavaWindow("Collaborative Product").JavaTree("NavTree")
                		sValue=split(sValue,"~")
				If ubound(sValue)=1 Then
					iInstance=cint(sValue(1))
				Else
					iInstance=1
				End If
				iTempInstance=1
				
				sRetNodePath = "#0"
				bFlag = false
				Set objTree = JavaWindow("Collaborative Product").JavaTree("NavTree")
				Set objCDNode = objTree.Object.getItem(0)
				iPtnCnt = cint(objCDNode.getItemCount())
				For iCnt = 0 to iPtnCnt -1
					Set objPtnNode = objCDNode.getItem(iCnt)
					objSubPtnCnt = cInt(objPtnNode.getItemCount())
				
					sVal = Fn_SISW_UI_JavaTree_GetSanitizedNodeName(objPtnNode)	
					If instr(1,sVal ,sValue(0)) Then
						sTreeVal = objTree.GetColumnValue( sRetNodePath & ":#" & iCnt , "Type" )
						If "Design Partition" = sTreeVal Then
							If iInstance=iTempInstance Then
								bFlag = true
								Exit for
							End If
							iTempInstance=iTempInstance+1
						End If
					End If
					If objSubPtnCnt > 0 Then
						For ispCnt = 0 to objSubPtnCnt -1
							objPtnNode.getItem(ispCnt)
							sVal = Fn_SISW_UI_JavaTree_GetSanitizedNodeName(objPtnNode.getItem(ispCnt))
							If instr(1,sVal ,sValue(0)) Then
								sTreeVal = objTree.GetColumnValue( sRetNodePath & ":#" & iCnt & ":#" & ispCnt , "Type" )
								If "Design Partition" = sTreeVal Then
									If iInstance=iTempInstance Then
										bFlag = true
										Exit for
									End If
                                    iTempInstance=iTempInstance+1
								End If
							End If
						Next
					End If
					If bFlag = true then exit for
				Next
				
				If bFlag Then
					Fn_CPD_ContentExplorer = sVal
				Else
					Fn_CPD_ContentExplorer = false
				End If
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "getfullnodebynodename", "getfullnodebynodename_ext"
			Fn_CPD_ContentExplorer=False
			arrStrNode=SPlit(sNode,":")
        	Set oCurrentNode=JavaWindow("Collaborative Product").JavaTree("NavTree").Object.getItem(0)
			For iCnt = 1 to UBound(arrStrNode)
				bFlag=False
				iNodeItemsCount = oCurrentNode.getItemCount()
        		For iCounter = 0 to iNodeItemsCount - 1
        			If sAction = "getfullnodebynodename_ext" Then
    					appNodeName = oCurrentNode.getItem(iCounter).getData().toString()
    				Else
    					appNodeName = oCurrentNode.getItem(iCounter).getText()
    				End If
					If UBound(arrStrNode)=iCnt Then
						If instr(1,Trim(appNodeName), Trim(arrStrNode(iCnt))) Then
							Fn_CPD_ContentExplorer=Trim(appNodeName)
							bFlag=True
							Exit For
						End If
					Else
						If Trim(appNodeName) = Trim(arrStrNode(iCnt)) Then
							Set oCurrentNode = oCurrentNode.getItem(iCounter)
							bFlag=True
							Exit For
						End If
					End If
				Next
				If bFlag=False Then
					Exit For
				End If
			Next 
			Set oCurrentNode=Nothing
			' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "getcolumnvaluebynodeandcolumnname" '[TC1121-2015102600-19_11_2015-VivekA-NewDevelopment] - Added by Pallavi C
			If iTreeIndex <> False Then
				sCellData=Trim(JavaWindow("Collaborative Product").JavaTree("NavTree").GetColumnValue(iTreeIndex,sColumn))
				Fn_CPD_ContentExplorer = sCellData
			Else
				Fn_CPD_ContentExplorer = False
			End If
			'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		Case "exists_contentexplorerlabelpath"		'[TC1121-20151116a00-09_12_2015-AnkitN-NewDevelopment] - Added by Ankit Nigam
				sVal = ""
				Set objSelectType = Description.Create()
				objSelectType("Class Name").value = "JavaObject"
				objSelectType("toolkit class").value = "org.eclipse.ui.forms.widgets.ImageHyperlink"
				Set  intNoOfObjects = JavaWindow("Collaborative Product").JavaObject("CPDStructureInfoPanel").ChildObjects(objSelectType)
				For iCnt = 1 to intNoOfObjects.count-1
					If iCnt = 1 Then
						sVal = intNoOfObjects(iCnt).Object.getText()
					Else
						sVal = sVal + "~" + intNoOfObjects(iCnt).Object.getText()
					End If 
					If iCnt = (intNoOfObjects.count-1) Then
						If instr(sValue,"(SUBSET)")>0 Then
							JavaWindow("Collaborative Product").JavaStaticText("SearchOptions").SetTOProperty "label", split(sValue,"(SUBSET)")(1)
							If Fn_SISW_UI_Object_Operations("Fn_CPD_ContentExplorer","Exist", JavaWindow("Collaborative Product").JavaStaticText("SearchOptions"),"") = True Then
								 sVal = sVal + "~(SUBSET)" + split(sValue,"(SUBSET)")(1)
							End If
						End If
					End If
					If instr(sVal , sValue) > 0 Then
						Fn_CPD_ContentExplorer =True
						Exit For
					End If
				Next
		'[TC1121-20151116b-10_12_2015-VivekA-NewDevelopment] - Verify Node which contains "( 12 )" in Node name
		Case "verifynumberinnode"
				If iTreeIndex <> False Then
                    Set objTree = JavaWindow("Collaborative Product").JavaTree("NavTree")                    
                    If Instr(iTreeIndex,":")>0 Then 
                    	aDisplayNode = Split(sDisplayNode,":")
                    	sVerifyNode = aDisplayNode(UBound(aDisplayNode))
						
						aTreeIndex = Split(iTreeIndex,":")
                        For iCounter=0 To UBound(aTreeIndex)
                        	If iCounter=0 Then
                            	Set oCurrentNode = objTree.Object.GetItem(Replace(aTreeIndex(iCounter),"#",""))
                            Else
                            	Set oCurrentNode = oCurrentNode.GetItem(Replace(aTreeIndex(iCounter),"#",""))
                            End If
                        Next
					Else
						sVerifyNode = sDisplayNode
						Set oCurrentNode = objTree.Object.GetItem(Replace(iTreeIndex,"#",""))
                    End If
                    'Get applcation Node name
                    If oCurrentNode.getText() <> "" Then
                    	sAppText = oCurrentNode.getText()
                    ElseIf oCurrentNode.getData().toString() <> "" Then
                    	sAppText = oCurrentNode.getData().toString()
                    End If
                    sNumber = oCurrentNode.getData().getNameDisplayText()
                    'Verify App Node Name with Script Node Name
                    If sVerifyNode = sAppText+""+sNumber Then
                    	Fn_CPD_ContentExplorer = True
					Else
						Fn_CPD_ContentExplorer = False
                    End If
                    			
                    Set oCurrentNode = Nothing
                Else
                	Fn_CPD_ContentExplorer = False
                End If
        '[TC1122-20151116d-18_12_2015-VivekA-NewDevelopment] - Added to verify blank cell
       	Case "cellverifyblank"
				If iTreeIndex <> False Then
					If JavaWindow("Collaborative Product").JavaTree("NavTree").GetColumnValue(iTreeIndex,sColumn)="" Then
						Fn_CPD_ContentExplorer = True
					Else
						Fn_CPD_ContentExplorer = False
					End If
				End If
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "popupmenuselecton_on_changecontext" '[Tc11.5(2018616b.00)_NewDevelopment_PoonamC_21Sept2018 : Added Case to select ENC from content header]
						Dim stab
						aValue = split(sValue,"@")
						iCounter = 0
						iInstanceCnt = 0
						If uBound(aValue) = 1 then
							iInstanceCnt =  cInt(aValue(1)) -1
						End If
						stab = JavaWindow("Collaborative Product").JavaTab("TCComponentTab").GetROProperty("value")
						If stab <> "" Then
							If  Fn_CPD_CompnentTabOperations("IsMaximized",stab, "")  = False Then
								bflag=1
								Call Fn_CPD_CompnentTabOperations("DoubleClick",stab, "") 
							End If
						End If
						Set objSelectType = Description.Create()
						objSelectType("Class Name").value = "JavaObject"
						objSelectType("toolkit class").value = "org.eclipse.ui.forms.widgets.ImageHyperlink"
						Set  intNoOfObjects = JavaWindow("Collaborative Product").ChildObjects(objSelectType)
						For iCnt = 0 to intNoOfObjects.count-1
								If intNoOfObjects(iCnt).Object.getText() = trim(aValue(0)) Then
									If iCounter = iInstanceCnt Then
										Call Fn_SISW_UI_DeviceReplayObjectClick("Fn_CPD_ContentExplorer", intNoOfObjects(iCnt))
										wait 3
										sNumber = JavaWindow("Collaborative Product").WinMenu("ContextMenu").GetItemProperty("","SubMenuCount")
										'loop through Menu's
										For iCounter = 1 To sNumber
											sAppText = JavaWindow("Collaborative Product").WinMenu("ContextMenu").GetItemProperty("<Item "&iCounter&">","Label")
											If trim(sAppText) = trim(sPopupMenu) Then
												sVal = iCounter
												Exit For
											End If
										Next
										For iCounter = 1 To sVal
										 	Call Fn_KeyBoardOperation("SendKeys","{DOWN}") 
										Next
										Call Fn_KeyBoardOperation("SendKeys", "{ENTER}")
										Fn_CPD_ContentExplorer = True
										Exit for
									Else
										iCounter = iCounter + 1
									End If
								End IF
						Next		
						If  bflag=1 Then
							Call Fn_CPD_CompnentTabOperations("DoubleClick",stab, "") 
						End If	
		' - - - - -  -  -- -- -  -- - - - - -- - - - --  - --  - - -- - -- - - -- -- - - - -- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - -  - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 					
		Case Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_CPD_ContentExplorer ] Invalid case [ " & sAction & " ].")
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
	End Select
	If  Fn_CPD_ContentExplorer <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [ Fn_CPD_ContentExplorer ] executed successfuly with case [ " & sAction & " ].")
	End If
	set objTree = nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@	Function to select  the Component Tab into Collaborative Product Development	@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@	   Function Name		:			Fn_CPD_CompnentTabOperations(sAction,sTCComponentTabName) 
'@@	   Description			:		 	Verify the Tab Activated tab title
'@@	   									This is to Activate the Required tab
'@@	   									Select Tab To activate
'@@	   									Close Tab if open
'@@	   
'@@	   
'@@	   Parameters			:			1) sAction: Action to be performed on the Tab
'@@	   									2) sTCComponentTabName: ComponentTab to be selected.
'@@	   
'@@	   Return Value		   	: 			TRUE \ FALSE
'@@	   
'@@	   Pre-requisite		:		 	Required Item Should be Double Clicked in home Tab
'@@	   
'@@	   Examples				:			Fn_CPD_CompnentTabOperations("VerifyActivate", "Home","")
'@@	   Examples				:			Fn_CPD_CompnentTabOperations("Exists", "Home","")
'@@	   Examples				:			Fn_CPD_CompnentTabOperations("Exists", "Home @2","")
'@@	   Examples				:			Fn_CPD_CompnentTabOperations("Activate", "Home","")
'@@	   Examples				:			Fn_CPD_CompnentTabOperations("Close", "Home","")
'@@	   Examples				:			Fn_CPD_CompnentTabOperations("TabRMBMenuSelect", "Home","Close") 
'@@	   Examples				:			Fn_CPD_CompnentTabOperations("VerifyActivateBelow", "Effectivity","")
'@@	   Examples				:			Fn_CPD_CompnentTabOperations("ExistsBelow", "Home","")
'@@	   Examples				:			Fn_CPD_CompnentTabOperations("ActivateBelow", "Effectivity","")
'@@	   Examples				:			Fn_CPD_CompnentTabOperations("CloseBelow", "Effectivity","")
'@@	   Examples				:			Fn_CPD_CompnentTabOperations("TabRMBMenuSelectBelow", "Home","Close") 
'@@	   
'@@	   History				:	
'@@				Developer Name				Date			Rev. No.	Reviewer				Changes Done								
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe			17-Jan-2012			1.0									Created
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe			18-Jan-2012			1.0									Added new cases "VerifyActivateBelow", "ActivateBelow", "CloseBelow", "TabRMBMenuSelectBelow"
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe			18-Jan-2012			1.0									Modified cases  "TabRMBMenuSelectBelow", "TabRMBMenuSelect"
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe			07-Mar-2012			1.0									Added case  "ExistsBelow", "Exists"
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Veena Gurjar			28-Feb-2013			1.0			Koustubh Watwe			Modified function.
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_CPD_CompnentTabOperations(sAction,sTCComponentTabName, sPopupMenu) 
	GBL_FAILED_FUNCTION_NAME="Fn_CPD_CompnentTabOperations"
	Dim sTabVal, bTabActive, iCount,objItem,objTabFld,i, aTab, iInstanceCnt
	Dim sxLen,syLen,sBounds,aBounds, aMenuList, StrMenu
	Fn_CPD_CompnentTabOperations = False
'	If inStr(sAction,"Below") > 0 then
'		Set objTabFld = JavaWindow("Collaborative Product").JavaObject("DownCompTab")
'	Else
'		Set objTabFld = JavaWindow("Collaborative Product").JavaObject("TCComponentTab")
'	End If

	Select Case sAction
		Case "VerifyActivate", "VerifyActivateBelow"
			Fn_CPD_CompnentTabOperations=Fn_SISW_UI_RACTabFolderWidget_Operation("VerifyActivate", sTCComponentTabName, "")
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "Exists", "ExistsBelow", "Exist", "ExistBelow"
			Fn_CPD_CompnentTabOperations=Fn_SISW_UI_RACTabFolderWidget_Operation("Verify", sTCComponentTabName, "")

		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "Activate", "ActivateBelow"
			Fn_CPD_CompnentTabOperations = Fn_SISW_UI_RACTabFolderWidget_Operation("Select", sTCComponentTabName, "")
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "Close", "CloseBelow"
			Fn_CPD_CompnentTabOperations = Fn_SISW_UI_RACTabFolderWidget_Operation("Close", sTCComponentTabName, sPopupMenu)
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "TabRMBMenuSelect", "TabRMBMenuSelectBelow" 
			Call Fn_CPD_CompnentTabOperations("Activate", sTCComponentTabName,"")
			Fn_CPD_CompnentTabOperations = Fn_SISW_UI_RACTabFolderWidget_Operation("RMBMenuSelect", sTCComponentTabName, sPopupMenu)
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "DoubleClick"
			Fn_CPD_CompnentTabOperations = Fn_SISW_UI_RACTabFolderWidget_Operation("DoubleClick", sTCComponentTabName, sPopupMenu)
		Case "IsMaximized"
			Fn_CPD_CompnentTabOperations = Fn_SISW_UI_RACTabFolderWidget_Operation("IsMaximized", sTCComponentTabName, sPopupMenu)
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Invalid ACTION  [" +sAction+ "] is Requested.")
	End Select
	IF Fn_CPD_CompnentTabOperations <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_CPD_CompnentTabOperations : Executed successfully with Case [" + sAction + "].")
	End If
	Set objItem = Nothing
	Set objTabFld = Nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name			:	Private Function Fn_CPD_ItemBasicCreate
'@@
'@@    Description				:	Function Used to create Basic Item
'@@
'@@    Parameters			    :	1. StrItemType		: Item type
'@@								:	2. StrConfItem		: Configure Iteam
'@@								:	3. StrItemID		: Item ID
'@@								:	4. StrItemRevID 	: Item Rev ID
'@@								:	5. StrItemName		: Item Name
'@@								:	6. StrItemDesc		: Item Description
'@@								:	7. StrItemUOM		: Unit of Measurement
'@@								:	8. sItemDetailsSet	: Detailed information required for Item creation ( for future use )
'@@
'@@    Return Value		   	   	: 	Item ID-RevID or False
'@@
'@@    Pre-requisite			:	CPD perspective should be activated.						
'@@
'@@    Examples					:	Call Fn_CPD_ItemBasicCreate("Workset", "ON", "", "", "Name", "Desc", "", "")
'@@
'@@	   History					:	
'@@				Developer Name				Date			Rev. No.	Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe			17-Jan-2012			1.0			Created
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe			25-May-2012			1.1			Modified object hierarchy
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Private Function Fn_CPD_ItemBasicCreate(StrItemType, StrConfItem, StrItemID, StrItemRevID, StrItemName, StrItemDesc, StrItemUOM, dicItemDetials)
		GBL_FAILED_FUNCTION_NAME="Fn_CPD_ItemBasicCreate"
		Dim sItemId, sRevId
		Dim objDialogNewItem
		Fn_CPD_ItemBasicCreate = False
		
		Set objDialogNewItem = Window("TeamcenterWindow").JavaDialog("New Item")
		If Fn_UI_ObjectExist("Fn_CPD_ItemBasicCreate", objDialogNewItem )=False Then
			Exit Function
		End If
		
		'Select Item Type
		' To handle Application change (Funcitonality is changed to Funciton from Tc 09 0119 Build) Code added by Archana 
		Call Fn_List_Select("Fn_CPD_ItemBasicCreate", objDialogNewItem,"ItemType",StrItemType)
		
		' Wait till  Button is Enabled
		objDialogNewItem.JavaButton("Next").WaitProperty "enabled", 1, 60000
		
		'Click on "Next" button
		Call Fn_Button_Click("Fn_CPD_ItemBasicCreate", objDialogNewItem,"Next")
		wait(2)
		If StrItemID <> "" Then
			'Set  Item Id
			Call Fn_Edit_Box("Fn_CPD_ItemBasicCreate",objDialogNewItem,"ItemID", StrItemID)
		End If
		
		If StrItemRevID <> "" Then
			'Set Revision ID
			Call Fn_Edit_Box("Fn_CPD_ItemBasicCreate",objDialogNewItem,"RevisionID", StrItemRevID)
		End If
		
		If  StrItemID = "" or StrItemRevID = "" Then
			'click on assign button
			Call Fn_Button_Click("Fn_CPD_ItemBasicCreate", objDialogNewItem, "Assign")
		End If
		Call Fn_ReadyStatusSync(5)
	
		'Extract Creation data
		sItemId = Fn_Edit_Box_GetValue("Fn_CPD_ItemBasicCreate", objDialogNewItem,"ItemID")
		sRevId = Fn_Edit_Box_GetValue("Fn_CPD_ItemBasicCreate", objDialogNewItem,"RevisionID")
		
		'*****************************************************************
		'Added by Tushar B, In case ItemId and rev field are blank
		If  sItemId = "" or sRevId = "" Then
			'click on assign button
			Call Fn_UpdateLogFiles(Time() & " - " & "WARNING - Assign button need to click again.", "")
			call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Item ID not shown in ItemId Textbox[" + CStr(sItemId) + "]")
			Call Fn_Button_Click("Fn_CPD_ItemBasicCreate", objDialogNewItem, "Assign")
			sItemId = Fn_Edit_Box_GetValue("Fn_CPD_ItemBasicCreate", objDialogNewItem,"ItemID")
			sRevId = Fn_Edit_Box_GetValue("Fn_CPD_ItemBasicCreate", objDialogNewItem,"RevisionID")
		End If
		'*****************************************************************
		
		'Set Item name
		Call Fn_Edit_Box("Fn_CPD_ItemBasicCreate", objDialogNewItem,"ItemName",StrItemName)
		'Set description
		Call Fn_Edit_Box("Fn_CPD_ItemBasicCreate", objDialogNewItem,"Description",StrItemDesc)
		'Set UOM
		If StrItemUOM <> "" Then
			Call Fn_Edit_Box("Fn_CPD_ItemBasicCreate", objDialogNewItem,"Unit of Measure",StrItemUOM)
		End If
		''**********************************************************
		'Radio button is not present in application and OR. - Commented by Koustubh - 25th Jan 2013
		'Click on "Next" button
'			checked Configuration item or not
'		If StrConfItem <> "" Then
'			Call Fn_Button_Click("Fn_CPD_ItemBasicCreate", objDialogNewItem,"Next")
'			If objDialogNewItem.JavaRadioButton("Radiobtn").Exist(5) Then
'				If lcase(StrConfItem) = "true" OR lcase(StrConfItem) = "on" Then
'	'				 set ON
'					objDialogNewItem.JavaRadioButton("Radiobtn").SetTOProperty "attached text", "True"
'				Else
'	'				 set OFF
'					objDialogNewItem.JavaRadioButton("Radiobtn").SetTOProperty "attached text", "False"
'				End If
'				Call Fn_UI_JavaRadioButton_SetON("Fn_CPD_ItemBasicCreate", objDialogNewItem , "Radiobtn")
'			End If
'		End If
''******************************************************************************************************

		'For Detail Creation...
				If TypeName(dicItemDetials) <> "String" Then
		
					'Click on [ Next ] button thrice [ To go to checkbox "Show as new root" ]
					Call Fn_Button_Click("Fn_CPD_ItemBasicCreate", objDialogNewItem,"Next")
					wait 1
					Call Fn_Button_Click("Fn_CPD_ItemBasicCreate", objDialogNewItem,"Next")
					wait 1
					
					Call Fn_Button_Click("Fn_CPD_ItemBasicCreate", objDialogNewItem,"Next")
					wait 1
					
					'Handle Dialog "Dont show this message again"			
					If JavaWindow("DefaultWindow").JavaWindow("Shell").JavaWindow("NewItemErrorDialog").Exist(8) Then
							'Click on OK button
							Call Fn_Button_Click("Fn_CPD_ItemBasicCreate", JavaWindow("DefaultWindow").JavaWindow("Shell").JavaWindow("NewItemErrorDialog") ,"OK")
							Call Fn_ReadyStatusSync(1)
					End If
					
					Call Fn_Button_Click("Fn_CPD_ItemBasicCreate", objDialogNewItem,"Next")
					Call Fn_ReadyStatusSync(1)
					
					'Check/Uncheck "Show as new root"
					If dicItemDetials("ShowAsNewRoot") <> "" Then
							If Cbool(dicItemDetials("ShowAsNewRoot")) Then
									Call Fn_CheckBox_Set("Fn_CPD_ItemBasicCreate",objDialogNewItem, "ShowAsNwRt","ON")
							Else
									Call Fn_CheckBox_Set("Fn_CPD_ItemBasicCreate",objDialogNewItem, "ShowAsNwRt","OFF")
							End If
							Call Fn_ReadyStatusSync(1)
						End If
				 End If
		'Detail Create END

		Wait(2)
		objDialogNewItem.JavaButton("Finish").WaitProperty "enabled", 1, 20000
		
		'Click on "Finish" butto
		'            Call Fn_Button_Click("Fn_CPD_ItemBasicCreate", objDialogNewItem, "Finish") 
		'			   = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = 
		'				Sandeep : Added Code to handle Negative scenario for BMIDE test cases
		If objDialogNewItem.JavaButton("Finish").GetROProperty("enabled")="0" Then
			If StrItemDesc="" Then
				Call Fn_Edit_Box("Fn_CPD_ItemBasicCreate", objDialogNewItem,"Description","Test")
			End If
			objDialogNewItem.JavaButton("Finish").WaitProperty "enabled", 1, 20000
			Call Fn_Button_Click("Fn_CPD_ItemBasicCreate", objDialogNewItem, "Finish")
		Else
			Call Fn_Button_Click("Fn_CPD_ItemBasicCreate", objDialogNewItem, "Finish")
		End If
		'= = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = 
		Fn_CPD_ItemBasicCreate = sItemId & "-" & sRevId
		Call Fn_ReadyStatusSync(1)
		
		'Click on Close button
		'Call Fn_Button_Click("Fn_CPD_ItemBasicCreate", objDialogNewItem, "Close") 
		If Fn_UI_ObjectExist("Fn_CPD_ItemBasicCreate",objDialogNewItem)=True Then
			'Click on Close button
			Call Fn_Button_Click("Fn_CPD_ItemBasicCreate", objDialogNewItem, "Close") 
		End If
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Created an Item of ID [" + CStr(sItemId) + "]")
		Set objDialogNewItem=Nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name			:	Fn_CPD_NewPartitionItemCreate
'@@
'@@    Description				:	Function Used to create Partition Item
'@@
'@@    Parameters			    :	1. sAction			: Action to be performed
'@@								:	2. StrItemType		: Item type
'@@								:	3. StrConfItem		: Configure Iteam
'@@								:	4. StrItemID		: Item ID
'@@								:	5. StrItemRevID 	: Item Rev ID
'@@								:	6. StrItemName		: Item Name
'@@								:	7. StrItemDesc		: Item Description
'@@								:	8. StrItemUOM		: Unit of Measurement
'@@								:	9. sItemDetailsSet	: Detailed information required for Item creation ( for future use )
'@@
'@@    Return Value		   	   	: 	Item ID-RevID or False
'@@
'@@    Pre-requisite			:	CPD perspective should be activated.						
'@@
'@@    Examples					:	Call Fn_CPD_NewPartitionItemCreate("Create", "Partition Design Item", "ON", "", "", "Name", "Desc", "", "")
'@@
'@@	   History					:	
'@@				Developer Name				Date			Rev. No.	Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe			17-Jan-2012			1.0			Created
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe			25-May-2012			1.1			Modifeid object hierarchy
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Veena Gurjar			01-Mar-2013			2.0			Modifeid object hierarchy
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_CPD_NewPartitionItemCreate(sAction, StrItemType, StrConfItem, StrItemID, StrItemRevID, StrItemName, StrItemDesc, StrItemUOM, sItemDetailsSet)
	GBL_FAILED_FUNCTION_NAME="Fn_CPD_NewPartitionItemCreate"
	Dim objCreateItem,sMenu,StrOldItemType
	StrOldItemType = sSchemeType
    StrItemType = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("CPD_DisplayName"),StrItemType)
    If StrItemType = False Then
        StrItemType = StrOldItemType
    End If 
	Set objCreateItem = JavaWindow("DefaultWindow").JavaWindow("NewItem")												
	'If Fn_UI_ObjectExist("Fn_CPD_NewPartitionItemCreate",objCreateItem)=False Then
	If Fn_SISW_UI_Object_Operations("Fn_CPD_NewPartitionItemCreate","Exist", objCreateItem, SISW_MICRO_TIMEOUT) = false then
		Select Case sAction
			Case "CreateByToolbar"
				' toolbar operations
			Case Else
				sMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("CPD_Menu"), "FileNewPartitionItem") 
				Call Fn_MenuOperation("Select",sMenu)
		End Select
		Call  Fn_ReadyStatusSync(3)	
	End If
	Fn_CPD_NewPartitionItemCreate = Fn_ItemBasicCreate(StrItemType,StrConfItem,StrItemID,StrItemRevID,StrItemName,StrItemDesc,StrItemUOM)
	Set objCreateItem = Nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name			:	Fn_CPD_NewWorksetCreate
'@@
'@@    Description				:	Function Used to create Workset
'@@
'@@    Parameters			    :	1. sAction			: Action to be performed
'@@								:	2. StrItemType		: Item type
'@@								:	3. StrConfItem		: Configure Iteam
'@@								:	4. StrItemID		: Item ID
'@@								:	5. StrItemRevID 	: Item Rev ID
'@@								:	6. StrItemName		: Item Name
'@@								:	7. StrItemDesc		: Item Description
'@@								:	8. StrItemUOM		: Unit of Measurement
'@@								:	9. sItemDetailsSet	: Detailed information required for Item creation ( for future use )
'@@
'@@    Return Value		   	   	: 	Item ID-RevID or False
'@@
'@@    Pre-requisite			:	CPD perspective should be activated.						
'@@
'@@    Examples					:	Call Fn_CPD_NewWorksetCreate("Create", "Workset", "ON", "", "", "Name", "Desc", "", "")
'@@    Examples					:	Call Fn_CPD_NewWorksetCreate("CreateByToolbar", "Workset", "ON", "", "", "Name", "Desc", "", "")
'@@
'@@	   History					:	
'@@				Developer Name				Date			Rev. No.	Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe			17-Jan-2012			1.0			Created
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe			25-May-2012			1.1			Modifeid object hierarchy
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Veena Gurjar			01-Mar-2013			2.0			Modifeid object hierarchy
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_CPD_NewWorksetCreate(sAction, StrItemType, StrConfItem, StrItemID, StrItemRevID, StrItemName, StrItemDesc, StrItemUOM, sOpenOnCreate)
	GBL_FAILED_FUNCTION_NAME="Fn_CPD_NewWorksetCreate"
	Dim objCreateItem, iItemCount
	Dim sItemId, sRevId
	Dim objDialogNewItem
	Dim iCount,crrItem,bFlag
	Dim WshShell
	Set objCreateItem = Fn_SISW_GetObject("New Item")
	
	If not Fn_ToolBarOperation("ButtonExist","Save Working Context","") Then         ' Modified function by Chaitali R.
	
		If Fn_ToolbarOperation("IsEnabled", "Search within current content","") = false Then
			If Fn_CPD_CollaborativeDesignCreate("Create", "", "CD_Test", "CD_Test", "True") <> False Then
				Call Fn_ToolbarOperation("Click", "Search within current content","")
			End If
		End If
		Wait 1
 		Call Fn_ToolbarOperation("Click", "Search within current content","")
 		Set dicContentSearch = CreateObject("Scripting.Dictionary")
 		dicContentSearch("SearchCriteria") = "ID=*"
		Call Fn_CPD_ContentSearchOperations("Search","","Attribute",dicContentSearch,"Search")
		dicContentSearch.RemoveAll
		Wait 2
	End If
	Wait 2
	Call Fn_ReadyStatusSync(5)
	If Fn_SISW_UI_Object_Operations("Fn_CPD_NewWorksetCreate","Exist",objCreateItem,SISW_MICRO_TIMEOUT) = False Then
		Wait(3)
		Select Case sAction
			Case "CreateByToolbar"
				' toolbar operations
				'Call Fn_ToolbarOperation("Click", "Create a new Workset","")
				
				Call Fn_ToolbarOperation("Click", "Save Working Context","")

			Case Else
				'Call Fn_MenuOperation("Select","File:New:Workset...")
				Call Fn_ToolbarOperation("Click", "Save Working Context","")
		End Select
		Call  Fn_ReadyStatusSync(3)	
	End If
	
	If StrItemType = "Workset" OR StrItemType = "Subset Definition" Then
		StrItemType = "4G " & StrItemType
	End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	If objCreateItem.JavaTree("ItemType").Exist(1) Then
		'Select Item type
		iItemCount=Fn_UI_Object_GetROProperty("Fn_CPD_NewWorksetCreate",objCreateItem.JavaTree("ItemType"), "items count")
		For iCount=0 To iItemCount-1
			crrItem=objCreateItem.JavaTree("ItemType").GetItem(iCount)
			If Trim(crrItem)="Most Recently Used:"+Trim(StrItemType) Then
				bFlag=True
				Exit For
			ElseIf Trim(crrItem)="Complete List" Then
				Exit For
			End If
		Next
	
		If bFlag=True Then
			Call Fn_JavaTree_Select("Fn_CPD_NewWorksetCreate", objCreateItem, "ItemType","Most Recently Used")
			Call Fn_JavaTree_Select("Fn_CPD_NewWorksetCreate", objCreateItem, "ItemType","Most Recently Used:"+StrItemType)
		Else
			Call Fn_UI_JavaTree_Expand("Fn_CPD_NewWorksetCreate", objCreateItem, "ItemType","Complete List")
			Call Fn_JavaTree_Select("Fn_CPD_NewWorksetCreate", objCreateItem, "ItemType","Complete List")
			Call Fn_JavaTree_Select("Fn_CPD_NewWorksetCreate", objCreateItem, "ItemType","Complete List:"+StrItemType)	
		End If
		wait 2
	
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Clicking On Next button
		objCreateItem.JavaButton("Next").WaitProperty "enabled", 1, 60000
	    Call Fn_Button_Click("Fn_CPD_NewWorksetCreate",objCreateItem, "Next")
		wait(2)
	End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	'Set Item ID
	If StrItemID <> "" Then
		'Set Item Id
		 Call Fn_Edit_Box("Fn_CPD_NewWorksetCreate",objCreateItem,"ItemID", StrItemID)
	Else
		Call Fn_Button_Click("Fn_CPD_NewWorksetCreate", objCreateItem, "AssignID")
		Call  Fn_ReadyStatusSync(1)
		wait(1)
	End If
	sItemId = Fn_Edit_Box_GetValue("Fn_CPD_NewWorksetCreate", objCreateItem,"ItemID")

	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	'Set Item Name
	Call Fn_Edit_Box("Fn_CPD_NewWorksetCreate", objCreateItem,"ItemName",StrItemName)

	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	'Set Item Description
	If  StrItemDesc="<SKIP>" Then
		'do nothing
	ElseIf StrItemDesc <> "" Then
		Call Fn_Edit_Box("Fn_CPD_NewWorksetCreate", objCreateItem,"Description",StrItemDesc)
	Else
		Call Fn_Edit_Box("Fn_CPD_NewWorksetCreate", objCreateItem,"Description","Test")
	End If

	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	'Set Item RevID
	If StrItemRevID <> ""  Then
		If objCreateItem.JavaEdit("RevisionID").Exist(5) Then
			Call Fn_Edit_Box("Fn_CPD_NewWorksetCreate",objCreateItem,"RevisionID", StrItemRevID)
		End If
	Else
		Call Fn_Button_Click("Fn_CPD_NewWorksetCreate", objCreateItem, "AssignRevID")
		wait(2)
    End If
	sRevId = Fn_Edit_Box_GetValue("Fn_CPD_NewWorksetCreate", objCreateItem,"RevisionID")

	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	'Set Item Unit of Measure
	If StrItemUOM <> "" Then
		objCreateItem.JavaButton("UnitOfMeasure").Click
		wait(2)
		objCreateItem.JavaWindow("TreeShell").JavaTree("Tree").Activate StrItemUOM
	End If

	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	'Check uncheck Open on Create
	
'	If sOpenOnCreate<>"" AND sOpenOnCreate<> "OFF" Then
'			'Call Fn_CheckBox_Set("Fn_CPD_CreatePartition",objCDCreate, "OpenOnCreate",dicPartitionInfo("OpenOnCreate"))
'			Call Fn_SISW_UI_JavaCheckBox_Operations("Fn_CPD_NewWorksetCreate", "Set", objCreateItem, "OpenOnCreate", "ON")
'	ElseIf sOpenOnCreate = "OFF" Then
'			Call Fn_SISW_UI_JavaCheckBox_Operations("Fn_CPD_NewWorksetCreate", "Set", objCreateItem, "OpenOnCreate", "OFF")
'			'Call Fn_CheckBox_Set("Fn_CPD_CreatePartition",objCDCreate, "OpenOnCreate",dicPartitionInfo("OpenOnCreate"))
'	End If
	
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	'Click Finish button
	wait(2)
	objCreateItem.JavaButton("Finish").WaitProperty "enabled", 1, 20000
	Call Fn_Button_Click("Fn_CPD_NewWorksetCreate", objCreateItem, "Finish")
	Call Fn_ReadyStatusSync(3)

	'= = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
	Fn_CPD_NewWorksetCreate = "'"&sItemId & "-" & sRevId

	If Fn_UI_ObjectExist("Fn_CPD_NewWorksetCreate",objCreateItem)=True Then
		'Click on Close button
		Call Fn_Button_Click("Fn_CPD_NewWorksetCreate", objCreateItem, "Cancel") 
	End If
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Created an Item of ID [" + CStr(sItemId) + "]")
	
	If Fn_CPD_CompnentTabOperations("Exist",sItemId &"/"& sRevId &";1-"& StrItemName &" (Content Explorer)", "") = True Then
		Call Fn_CPD_CompnentTabOperations("Activate",sItemId &"/"& sRevId &";1-"& StrItemName &" (Content Explorer)", "") 
		If Fn_CPD_ContentExplorer("Exist", sItemId &"/"& sRevId &";1-"& StrItemName &":CD_Test", "", "", "") = True Then
			Call Fn_CPD_ContentExplorer("Select", sItemId &"/"& sRevId &";1-"& StrItemName &":CD_Test", "", "", "")
			Call Fn_TcObjectDelete("False",sItemId &"/"& sRevId &";1-"& StrItemName &":CD_Test","Menu")
    	End If
    End If
    
Set objCreateItem = Nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name			:	Fn_CPD_DesignElementCreate
'@@
'@@    Description				:	Function Used to create Design Element
'@@
'@@    Parameters			    :	1. sAction				: Action to be performed
'@@								:	2. sDesignEleID			: Design Element ID
'@@								:	3. sName				: Name
'@@								:	4. sDescription 		: description
'@@								:	5. sLogicalDesignator	: Logical Designator
'@@								:	6. bCopyEffectivity		: Boolean value to set Copy Effectivity checkbox  True /False / ""
'@@								:	7. bCheckoutOnCreate	: Boolean value to set Checkout On Create checkbox  True / False
'@@
'@@    Return Value		   	   	: 	True Or False /  ModelID or False
'@@
'@@    Pre-requisite			:	CPD perspective should be activated.						
'@@
'@@    Examples					:	Call Fn_CPD_DesignElementCreate("Create", "", "nam", "desc", "", "", "", "")
'@@    Examples					:	Call Fn_CPD_DesignElementCreate("GetErrorMessageOnCreate", "", "nam", "desc", "", "", "", "")
'@@
'@@	   History					:	
'@@				Developer Name				Date			Rev. No.	Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe			16-Jan-2012			1.0			Created
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe			20-Feb-2012			1.0			Added case GetErrorMessageOnCreate
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe			29-Feb-2012			1.0			Modified case Create
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_CPD_DesignElementCreate(sAction, sDesignEleID, sName, sDescription, sLogicalDesignator, sShapeDesign, bCopyEffectivity, bCheckoutOnCreate)
	GBL_FAILED_FUNCTION_NAME="Fn_CPD_DesignElementCreate"
	Dim objDesignEle, bReturn, sType
	Fn_CPD_DesignElementCreate = False
	Set objDesignEle = JavaWindow("Collaborative Product").JavaWindow("NewBusinessObject")

	If Fn_UI_ObjectExist("Fn_CPD_DesignElementCreate",objDesignEle) = False Then
			bReturn =  Fn_ToolbarButtonClick_Ext(1,"Create Component")
			If bReturn =  False OR Fn_UI_ObjectExist("Fn_CPD_DesignElementCreate",objDesignEle) = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_CPD_DesignElementCreate ] Failed to open Design Element window.")
					Exit function
			End If
	End If
	Select Case sAction
		Case "Create",  "GetErrorMessageOnCreate" 
				
				sType = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("CPD_DisplayName"),"Design Element")
				
				If objDesignEle.JavaTree("BusinessObjectType").Exist(2) Then
					' select collaborative design from tree
						objDesignEle.JavaTree("BusinessObjectType").Expand "Complete List"
						wait 1
						objDesignEle.JavaTree("BusinessObjectType").Select "Complete List:" & sType
			
						' click on next
						Call Fn_Button_Click("Fn_CPD_DesignElementCreate",objDesignEle,"Next" )
						wait(2)
				End If
				' if ModelD is empty
				objDesignEle.JavaStaticText("Field").SetTOProperty "label", "ID:"
				If sDesignEleID = "" Then
					'	then click on assign
					Call Fn_Button_Click("Fn_CPD_DesignElementCreate",objDesignEle,"Assign" )
					Call Fn_ReadyStatusSync(5)
					'Fn_CPD_DesignElementCreate = objDesignEle.JavaEdit("Field").GetROProperty("value")
					Fn_CPD_DesignElementCreate = Fn_UI_Object_GetROProperty("Fn_CPD_DesignElementCreate",objDesignEle.JavaEdit("Field"), "value")
				
				Else
					Call Fn_Edit_Box("Fn_CPD_DesignElementCreate",objDesignEle,"Field",sDesignEleID)
					Fn_CPD_DesignElementCreate = True
				End If
				
				'set name
				If sName <> "" Then
					objDesignEle.JavaStaticText("Field").SetTOProperty "label", "Name:"
					'objDesignEle.JavaEdit("Field").Type sName
					Call Fn_Edit_Box("Fn_CPD_DesignElementCreate",objDesignEle,"Field",sName)
					Call Fn_ReadyStatusSync(5)
				End If
	
				' set description
				If sDescription <> "" Then
					objDesignEle.JavaStaticText("Field").SetTOProperty "label", "Description:"
					Call Fn_Edit_Box("Fn_CPD_DesignElementCreate",objDesignEle,"Field",sDescription)
				End If

				' set sLogicalDesignator
				If sLogicalDesignator <> "" Then
					objDesignEle.JavaStaticText("Field").SetTOProperty "label", "Logical Designator:"
					Call Fn_Edit_Box("Fn_CPD_DesignElementCreate",objDesignEle,"Field",sLogicalDesignator)
				End If

'				If sShapeDesign <> "" OR bCopyEffectivity <> ""  OR bCheckoutOnCreate <> ""  Then
					Call Fn_Button_Click("Fn_CPD_DesignElementCreate",objDesignEle,"Next" )
					wait(2)
					If sShapeDesign <> "" Then
						If lcase(cstr(trim(sShapeDesign))) <> "false" Then	
							Call Fn_CheckBox_Select("Fn_CPD_DesignElementCreate", objDesignEle, "CreateShapeDesign" )
							Call Fn_List_Select("Fn_CPD_DesignElementCreate", objDesignEle, "CreateOptions", sShapeDesign)
						Else
							Call Fn_CheckBox_Set("Fn_CPD_DesignElementCreate", objDesignEle, "CreateShapeDesign","OFF")
						End If
					End If

					If bCopyEffectivity <> "" Then
						If cBool(bCopyEffectivity) Then
							wait(1)
							Call Fn_CheckBox_Set("Fn_CPD_DesignElementCreate", objDesignEle, "CopyEffectivity","ON")
						Else
							wait(1)
							Call Fn_CheckBox_Set("Fn_CPD_DesignElementCreate", objDesignEle, "CopyEffectivity","OFF")
						End If
					End If

					If bCheckoutOnCreate <> "" Then
						If cBool(bCheckoutOnCreate) Then
							wait(1)
							Call Fn_CheckBox_Set("Fn_CPD_DesignElementCreate", objDesignEle, "CheckOutOnCreate","ON")
						Else
							wait(1)
							Call Fn_CheckBox_Set("Fn_CPD_DesignElementCreate", objDesignEle, "CheckOutOnCreate","OFF")
						End If
					End If

'				End If
                
			' click on finish
			Call Fn_Button_Click("Fn_CPD_DesignElementCreate",objDesignEle,"Finish" )
			Call Fn_ReadyStatusSync(5)

			If sAction = "GetErrorMessageOnCreate" Then
				Fn_CPD_DesignElementCreate = False
				If objDesignEle.JavaWindow("Error").Exist(15) Then
					If objDesignEle.JavaWindow("Error").JavaEdit("DetailsMsg").Exist(1) Then
						Fn_CPD_DesignElementCreate = objDesignEle.JavaWindow("Error").JavaEdit("DetailsMsg").GetROProperty("value")
						Call Fn_Button_Click("Fn_CPD_DesignElementCreate",objDesignEle.JavaWindow("Error"),"OK" )
					Else

						Fn_CPD_DesignElementCreate = objDesignEle.JavaWindow("Error").JavaEdit("ErrorMsg").GetROProperty("value")
						Call Fn_Button_Click("Fn_CPD_DesignElementCreate",objDesignEle.JavaWindow("Error"),"OK" )
					End If
				End If
			End If
			If objDesignEle.JavaWindow("Error").Exist(10) Then
				If objDesignEle.JavaWindow("Error").getROProperty("title") = "Paste" Then
					' do nothing
				Else
					Fn_CPD_DesignElementCreate = False
				End If
				Call Fn_Button_Click("Fn_CPD_DesignElementCreate",objDesignEle.JavaWindow("Error"),"OK" )
			End If
			
			Call Fn_Button_Click("Fn_CPD_DesignElementCreate",objDesignEle,"Cancel" )
			Call Fn_CPD_CompnentTabOperations("Close", sName,"")
			Call Fn_ReadyStatusSync(1)
	' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_CPD_DesignElementCreate ] Invalid case [ " & sAction & " ].")
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
	End Select
	If  Fn_CPD_DesignElementCreate <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [ Fn_CPD_DesignElementCreate ] executed successfuly with case [ " & sAction & " ].")
	End If
	Set objDesignEle = Nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name			:	Fn_CPD_SubsetDefinitionCreate
'@@
'@@    Description				:	Function Used to create Subset Definition
'@@
'@@    Parameters			    :	1. sAction		: Action to be performed
'@@								:	2. sName		: Name
'@@								:	3. sDescription : description
'@@								:	4. bOpenOnCreate: Boolean value to set Open On Create checkbox  True / False
'@@
'@@    Return Value		   	   	: 	True Or False
'@@
'@@    Pre-requisite			:	CPD perspective should be activated.						
'@@
'@@    Examples					:	Call  Fn_CPD_SubsetDefinitionCreate("Create", "subsetDef", "desc", True)
'@@	
'@@								    Set dicContentSearch = CreateObject("Scripting.Dictionary")
'@@									dicContentSearch.RemoveAll
'@@									dicContentSearch("SearchCriteria") ="Name=Jotiba Takkekar~ID=*"
'@@								  	Call  Fn_CPD_SubsetDefinitionCreate("Create:SearchCriteria", "subsetDef", "desc", True)
'@@                               	Note : Associate DictionaryDeclaration.vbs to tset case  for this case
'@@
'@@	   History					:	
'@@				Developer Name				Date			Rev. No.	Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe			18-Jan-2012			1.0			Created
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Jotiba Takkekar 		31-Aug-2017						Case : "Create:SearchCriteria"  Note: Associate DictionaryDeclaration.vbs to tset case 
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@ 
Public Function Fn_CPD_SubsetDefinitionCreate(sAction, sName, sDescription, bOpenOnCreate)
	GBL_FAILED_FUNCTION_NAME="Fn_CPD_SubsetDefinitionCreate"
   Dim objSubsetDef, bReturn
	Fn_CPD_SubsetDefinitionCreate = False
	Set objSubsetDef = JavaWindow("Collaborative Product").JavaWindow("NewBusinessObject")
	
	aAction=Split(sAction,":")
	sAction=aAction(0)
	
	If not Fn_ToolBarOperation("IsEnabled","Save Working Context","") Then
			'------------------------------------------------- [TC11.4-20170815.00-31_8_2017-JotibaT-Maintenance] - Added code to specify SearchCriteria other than ID=*
			If ubound(aAction)>0 Then                         ' Note: Associate DictionaryDeclaration.vbs to tset case 
				If aAction(1)="SearchCriteria" Then
			   		Call Fn_CPD_ContentSearchOperations("Search","","Attribute",dicContentSearch,"Search")
			   		dicContentSearch.RemoveAll
					wait 1
				End If
			'-------------------------------------------------
		   Else
				Set dicContentSearch = CreateObject("Scripting.Dictionary")
		 		dicContentSearch("SearchCriteria") = "ID=*"
				Call Fn_CPD_ContentSearchOperations("Search","","Attribute",dicContentSearch,"Search")
				dicContentSearch.RemoveAll
				wait 1
		   End If
	 	
    End If
    	Call Fn_ReadyStatusSync(5)
'	If Fn_UI_ObjectExist("Fn_CPD_SubsetDefinitionCreate",objSubsetDef) = False Then
	If Fn_SISW_UI_Object_Operations("Fn_CPD_SubsetDefinitionCreate","Exist",objSubsetDef,SISW_MICRO_TIMEOUT) = False Then
			'bReturn =  Fn_ToolbarButtonClick_Ext(1,"Create Subset Definition")
			wait 3
			bReturn =  Fn_ToolbarButtonClick_Ext(1,"Save Working Context")
			If bReturn =  False OR Fn_UI_ObjectExist("Fn_CPD_SubsetDefinitionCreate",objSubsetDef) = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_CPD_SubsetDefinitionCreate ] Failed to open Create Subset Definition window.")
					Exit function
			End If
	End If
	Wait 2
	Select Case sAction
		Case "Create"
			'added condition to check existence of object tree , for single object tree does not appear.
			If objSubsetDef.JavaTree("BusinessObjectType").Exist(SISW_MICRO_TIMEOUT) = True Then
				' select collaborative design from tree
				objSubsetDef.JavaTree("BusinessObjectType").Expand "Complete List"
				wait 1
				objSubsetDef.JavaTree("BusinessObjectType").Select "Complete List:4G Subset Definition"
				Wait 1
				' click on next
				Call Fn_Button_Click("Fn_CPD_SubsetDefinitionCreate",objSubsetDef,"Next" )
				wait 3
			End If
				
				' Modified by Chaitali R				
								
				'set name
				If sName <> "" Then
					objSubsetDef.JavaStaticText("Field").SetTOProperty "label", "Name:"
					'objSubsetDef.JavaEdit("Field").Type sName
					objSubsetDef.JavaEdit("Field").Set(sName)
					Call Fn_ReadyStatusSync(5)
				End If
	
				' set description
				If sDescription <> "" Then
					objSubsetDef.JavaStaticText("Field").SetTOProperty "label", "Description:"
					Call Fn_Edit_Box("Fn_CPD_SubsetDefinitionCreate",objSubsetDef,"Field",sDescription)
				End If
								
				If Fn_SISW_UI_Object_Operations("Fn_CPD_SubsetDefinitionCreate", "Enabled", objSubsetDef.JavaButton("Next"), "") = True Then
					Call Fn_Button_Click("Fn_CPD_SubsetDefinitionCreate",objSubsetDef,"Next" )
					' if open on create is not empty then click on next
					If bOpenOnCreate <> "" Then
						' set open on create
						If cBool(bOpenOnCreate) Then
							Call Fn_CheckBox_Set("Fn_CPD_SubsetDefinitionCreate",objSubsetDef, "OpenOnCreate","ON")
						Else
							Call Fn_CheckBox_Set("Fn_CPD_SubsetDefinitionCreate",objSubsetDef, "OpenOnCreate","OFF")
						End If
					End If
				End If
			' click on finish
			Call Fn_Button_Click("Fn_CPD_SubsetDefinitionCreate",objSubsetDef,"Finish" )
			Call Fn_ReadyStatusSync(5)

			Call Fn_Button_Click("Fn_CPD_SubsetDefinitionCreate",objSubsetDef,"Cancel" )
			Fn_CPD_SubsetDefinitionCreate = True
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_CPD_SubsetDefinitionCreate ] Invalid case [ " & sAction & " ].")
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
	End Select
	If  Fn_CPD_SubsetDefinitionCreate <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [ Fn_CPD_SubsetDefinitionCreate ] executed successfuly with case [ " & sAction & " ].")
	End If
	Set objSubsetDef = Nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name			:	Fn_CPD_RecipeOperations
'@@
'@@    Description				:	Function Used to perform operations on Recipe Explorer
'@@
'@@    Parameters			    :	1. sAction		: Action to be performed
'@@								:	2. sRecipeTab	: Component tab name
'@@								:	3. sNode		: Node Name
'@@								:	4. sColumn		: Column Name
'@@								:	5. sValue 		: value to be verified
'@@								:	6. sPopupMenu	: Popup menu
'@@
'@@    Return Value		   	   	: 	True Or False
'@@
'@@    Pre-requisite			:	CPD perspective should be activated.						
'@@
'@@    Examples					:	Call Fn_CPD_RecipeOperations("Select", "subdef Recipe", "Group:DE000003/001;1-de", "", "", "")
'@@    Examples					:	Call Fn_CPD_RecipeOperations("Exist", "subdef Recipe", "Group:DE000003/001;1-de", "", "", "")
'@@    Examples					:	Call Fn_CPD_RecipeOperations("DeSelect", "subdef Recipe", "Group:DE000003/001;1-de", "", "", "")
'@@    Examples					:	Call Fn_CPD_RecipeOperations("MultiSelect", "", "Group~Group:DE000003/001;1-de", "", "", "")
'@@    Examples					:	Call Fn_CPD_RecipeOperations("Expand", "", "Group:DE000003/001;1-de", "", "", "")
'@@    Examples					:	Call Fn_CPD_RecipeOperations("PopupMenuSelect", "", "Group:DE000003/001;1-de", "", "", "Copy	Ctrl+C")
'@@    Examples					:	Call Fn_CPD_RecipeOperations("PopupMenuVerifyProperty", "", "Group:DE000003/001;1-de", "", "Exists", "Copy	Ctrl+C")
'@@    Examples					:	Call Fn_CPD_RecipeOperations("PopupMenuVerifyProperty", "", "Group:DE000003/001;1-de", "", "Enabled", "Copy	Ctrl+C")
'@@    Examples					:	Call Fn_CPD_RecipeOperations("CellVerify", "", "Group:DE000003/001;1-de", "Type", "Collaborative Design", "")
'@@    Examples					:	Call Fn_CPD_RecipeOperations("CellEdit", "", "DE000176/001;1-d2", "Logic", "Include", "")
'@@    Examples					:	Call Fn_CPD_RecipeOperations("MultiSelectPopupMenuSelect", "", "DE000176/001;1-d2~DE000177/001;1-d3", "", "", "Group")
'@@    Examples					:	Call Fn_CPD_RecipeOperations("CellListVerify", "", "Design Element ("+dicContentSearch("SearchCriteria")+")" , "Logic", DataTable("LogColVal",dtGlobalSheet) , "")
'@@    Examples					:	Call Fn_CPD_RecipeOperations("GetIndex", "", "Group:DE000003/001;1-de", "", "", "")
'@@    Examples					:	Call Fn_CPD_RecipeOperations("MoveRecipeUP", "", "Group:DE000003/001;1-de", "", "", "")
'@@    Examples					:	Call Fn_CPD_RecipeOperations("MoveRecipeDown", "", "Group:DE000003/001;1-de", "", "", "")
'@@    Examples					:	Call Fn_CPD_RecipeOperations("getChildrenCount", "", "", "", "", "")
'@@
'@@	   History					:	
'@@				Developer Name				Date			Rev. No.	Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe			18-Jan-2012			1.0			Created
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe			11-Apr-2012			1.0			Added case CellEdit
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe			17-Apr-2012			1.0			Added case CellEdit
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Sandeep N				02-May-2013			1.0			Added case GetIndex,MoveRecipeUP,MoveRecipeDown
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Poonam C				26-Nov-2015			1.1			Added case DeleteRecipe					[TC1121-2015110900-26_11_2015-VivekA-NewDevelopment]
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Chaitali R				30-Nov-2015			1.1			Added case VerifyEnabledSearchOptions	[TC1121-2015110900-30_11_2015-VivekA-NewDevelopment]
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Ankit N					02-Dec-2015			1.1			Added case getChildrenCount				[TC1121-2015110900-02_12_2015-AnkitN-NewDevelopment]
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_CPD_RecipeOperations(sAction, sRecipeTab, sNode, sColumn, sValue, sPopupMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_CPD_RecipeOperations"
	Dim bReturn, objTree, aMenuList, aNodes, iCnt
	Dim iItemCnt, iCount, sListValue, bFlag
	Dim objSelectType, intNoOfObjects
	Dim objSearchOptions
	Fn_CPD_RecipeOperations = False
	bFlag=False
	If sRecipeTab <> "" Then
		If Fn_CPD_CompnentTabOperations("ActivateBelow",sRecipeTab,"") = False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_CPD_RecipeOperations ] failed to select component tab [ " & sRecipeTab & " ].")
			Exit Function
		End If
	End If

	Set objTree = JavaWindow("Collaborative Product").JavaTree("ComponentTree")
	If Fn_SISW_UI_Object_Operations("Fn_CPD_RecipeOperations","Exist", objTree,SISW_MINLESS_TIMEOUT) = False Then
		Call Fn_UI_JavaStaticText_Click("Fn_CPD_RecipeOperations", JavaWindow("Collaborative Product"), "Recipe", "1", "1", "")
		bFlag =True
		If Fn_SISW_UI_Object_Operations("Fn_CPD_RecipeOperations","Exist", objTree,SISW_MINLESS_TIMEOUT) = False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_CPD_RecipeOperations ] failed to find Recipe Component Tree.") 
			Set objTree = Nothing
			Exit Function
		End If
	End If
	

	bReturn = False
	If sNode <> "" Then
		bReturn = Fn_UI_JavaTreeGetItemPathExt("Fn_CPD_RecipeOperations",JavaWindow("Collaborative Product").JavaTree("ComponentTree"), sNode, "", "")
	End If

	Select Case sAction
		Case "getChildrenCount"											'Added case getChildrenCount	[TC1121-2015110900-02_12_2015-AnkitN-NewDevelopment]
				Fn_CPD_RecipeOperations = objTree.Object.getItemCount()
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "Select", "SelectAndShowResult", "SelectAndAddSearchTerm"
			If bReturn <> False Then
				objTree.Select bReturn
				Select Case sAction
					Case "SelectAndShowResult"
							Call Fn_Button_Click("Fn_CPD_RecipeOperations",JavaWindow("Collaborative Product"), "ShowResults")
					Case "SelectAndAddSearchTerm"
							Call Fn_Button_Click("Fn_CPD_RecipeOperations",JavaWindow("Collaborative Product"), "AddSearchTerm")
				End Select
				Fn_CPD_RecipeOperations = True
			End If
		'Case for clicking 'Add Search Term' button without selecting node in "objTree"		
		Case "AddSearchTermWithoutSelect"	
				Fn_CPD_RecipeOperations = Fn_Button_Click("Fn_CPD_RecipeOperations",JavaWindow("Collaborative Product"), "AddSearchTerm")
		'Case for clicking 'Show Results' button without selecting node in "objTree"						
		Case "ShowResultWithoutSelect"
				JavaWindow("Collaborative Product").JavaButton("ShowResults").SetTOProperty "label" , "Replay"
				Fn_CPD_RecipeOperations = Fn_Button_Click("Fn_CPD_RecipeOperations",JavaWindow("Collaborative Product"), "ShowResults")
		'Case for clicking 'Add to Recipe' button without selecting node in "objTree"						
		Case "AddToRecipe"
				Fn_CPD_RecipeOperations = Fn_Button_Click("Fn_CPD_RecipeOperations",JavaWindow("Collaborative Product"), "AddToRecipe")
		Case "ExistRecipeTab"
				If bFlag=True Then
					Call Fn_UI_JavaStaticText_Click("Fn_CPD_RecipeOperations", JavaWindow("Collaborative Product"), "Recipe", "1", "1", "")
					Fn_CPD_RecipeOperations=True
				End If
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
   		Case "MultiSelect"
				aNodes = split(sNode,"~")
				For iCnt = 0 to UBound(aNodes)
					bReturn = Fn_UI_JavaTreeGetItemPathExt("Fn_CPD_RecipeOperations",JavaWindow("Collaborative Product").JavaTree("ComponentTree"), aNodes(iCnt), "", "")
					If bReturn <> False Then
						If iCnt=0 Then
							objTree.Select bReturn
						Else
							objTree.ExtendSelect bReturn
						End If
						Fn_CPD_RecipeOperations = True
					Else
						Fn_CPD_RecipeOperations = False
						Exit For
					End If
				Next
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
   		Case "MultiSelectPopupMenuSelect"
				aNodes = split(sNode,"~")
				For iCnt = 0 to UBound(aNodes)
					bReturn = Fn_UI_JavaTreeGetItemPathExt("Fn_CPD_RecipeOperations",JavaWindow("Collaborative Product").JavaTree("ComponentTree"), aNodes(iCnt), "", "")
					If bReturn <> False Then
						If iCnt = 0 then
							objTree.Select bReturn
						ElseIf iCnt = UBound(aNodes) Then
							objTree.ExtendSelect bReturn
							wait 1
							objTree.OpenContextMenu bReturn
							wait 1
						Else
							objTree.ExtendSelect bReturn
						End If
						Fn_CPD_RecipeOperations = True
					Else
						Fn_CPD_RecipeOperations = False
						Exit For
					End If
				Next
				aMenuList = split(sPopupMenu,":")
				'Select Menu action
				Select Case Ubound(aMenuList)
					Case "0"
						 sPopupMenu = JavaWindow("Collaborative Product").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0))
					Case "1"
						sPopupMenu = JavaWindow("Collaborative Product").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1))
					Case "2"
						sPopupMenu = JavaWindow("Collaborative Product").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1),aMenuList(2))
					Case Else
						Fn_CPD_RecipeOperations = FALSE
						Exit Function
				End Select
				JavaWindow("Collaborative Product").WinMenu("ContextMenu").Select sPopupMenu
				Fn_CPD_RecipeOperations = True
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "Expand"
			If bReturn <> False Then
				objTree.Expand bReturn
				Fn_CPD_RecipeOperations = True
			End If
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "Collapse"
			If bReturn <> False Then
				objTree.Collapse bReturn
				Fn_CPD_RecipeOperations = True
			End If
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "Exist"
			If bReturn <> False Then
				Fn_CPD_RecipeOperations = True
			End If
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "DeSelect"
				If bReturn <> False Then
					objTree.Deselect bReturn
					Fn_CPD_RecipeOperations = True
				End If
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "CellVerify"
				If bReturn <> False Then
					If trim(objTree.GetColumnValue(bReturn, sColumn)) = sValue Then
						Fn_CPD_RecipeOperations = True
					End If
				End If
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "CellEdit"
				If bReturn <> False Then
					bReturn = Fn_UI_ClickJavaTreeCell("Fn_CPD_RecipeOperations",JavaWindow("Collaborative Product"), "ComponentTree", sNode ,sColumn, "LEFT")
					If bReturn <> False Then
						wait 1
						Set objSelectType=description.Create()
						objSelectType("Class Name").value = "JavaList"
						objSelectType("toolkit class").value = "org.eclipse.swt.custom.CCombo"
						'objSelectType("tagname").value = "List"
						Set  intNoOfObjects = JavaWindow("Collaborative Product").ChildObjects(objSelectType)
						For iCnt = 0 to intNoOfObjects.count-1
							intNoOfObjects(iCnt).Select sValue
							Call Fn_KeyBoardOperation("SendKeys", "{TAB}")
							Fn_CPD_RecipeOperations = True
						Next
					End If
				End If
				Set objSelectType = Nothing
				Set intNoOfObjects = Nothing
		 ' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
  Case "CellListVerify" ' [ Include , Exclude and Filter ]
		If bReturn <> False Then
			bReturn = Fn_UI_ClickJavaTreeCell("Fn_CPD_RecipeOperations",JavaWindow("Collaborative Product"), "ComponentTree", sNode ,sColumn, "LEFT")
			If bReturn <> False Then
				Wait 1
				Set objSelectType=description.Create()
				objSelectType("Class Name").value = "JavaList"
				objSelectType("toolkit class").value = "org.eclipse.swt.custom.CCombo"
				'objSelectType("tagname").value = "List"
				Set intNoOfObjects = JavaWindow("Collaborative Product").ChildObjects(objSelectType)

				For iCnt = 0 to intNoOfObjects.count-1
					
					iItemCnt = cInt(intNoOfObjects(iCnt).getROProperty("items count"))
					For iCount = 0 to iItemCnt - 1
						sListValue = intNoOfObjects(iCnt).Object.getItem(iCount)
						If trim(sListValue) = trim(sValue)  then
							Fn_CPD_RecipeOperations = True
							Exit for
						End If
					Next
					If Fn_CPD_RecipeOperations = True Then
						Call Fn_KeyBoardOperation("SendKeys", "{ESC}")
						Exit for
					End If
				Next
			End If
		End If
		Set objSelectType = Nothing
		Set intNoOfObjects = Nothing
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "PopupMenuSelect"
			If bReturn <> False Then
				objTree.Select bReturn
                wait 1
				objTree.OpenContextMenu bReturn
				wait 1
				aMenuList = split(sPopupMenu,":")
				'Select Menu action
				Select Case Ubound(aMenuList)
					Case "0"
						 sPopupMenu = JavaWindow("Collaborative Product").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0))
					Case "1"
						sPopupMenu = JavaWindow("Collaborative Product").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1))
					Case "2"
						sPopupMenu = JavaWindow("Collaborative Product").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1),aMenuList(2))
					Case Else
						Fn_CPD_RecipeOperations = FALSE
						Exit Function
				End Select
				JavaWindow("Collaborative Product").WinMenu("ContextMenu").Select sPopupMenu
				Fn_CPD_RecipeOperations = True
			End If
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "PopupMenuVerifyProperty"
				aNodes = split(sNode,"~")
				For iCnt = 0 to UBound(aNodes)
					bReturn = Fn_UI_JavaTreeGetItemPathExt("Fn_CPD_RecipeOperations",JavaWindow("Collaborative Product").JavaTree("ComponentTree"), aNodes(iCnt), "", "")
					If bReturn <> False Then
						If iCnt = 0 then
							objTree.ExtendSelect bReturn
							If iCnt = UBound(aNodes) Then
								wait 1
								objTree.OpenContextMenu bReturn
								wait 1
							End If
						ElseIf iCnt = UBound(aNodes) Then
							objTree.ExtendSelect bReturn
							wait 1
							objTree.OpenContextMenu bReturn
							wait 1
						Else
							objTree.ExtendSelect bReturn
						End If
						Fn_CPD_RecipeOperations = True
					Else
						Fn_CPD_RecipeOperations = False
						Exit For
					End If
				Next
				aMenuList = split(sPopupMenu,":")
				'Select Menu action
				Select Case Ubound(aMenuList)
					Case "0"
						 sPopupMenu = JavaWindow("Collaborative Product").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0))
					Case "1"
						sPopupMenu = JavaWindow("Collaborative Product").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1))
					Case "2"
						sPopupMenu = JavaWindow("Collaborative Product").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1),aMenuList(2))
					Case Else
						Fn_CPD_RecipeOperations = FALSE
						Exit Function
				End Select
				Fn_CPD_RecipeOperations = JavaWindow("Collaborative Product").WinMenu("ContextMenu").CheckItemProperty(sPopupMenu, sValue, True, 20)
				Call Fn_KeyBoardOperation("SendKeys", "{ESC}")
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		'this case return index number of specific child under its imidiate parent
		Case "GetIndex"
			bReturn=split(bReturn,"#")
			Fn_CPD_RecipeOperations=Cint(bReturn(ubound(bReturn)))
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "MoveRecipeUP","MoveRecipeDown","DeleteRecipe"
			If bReturn <> False Then
				If Fn_CPD_CompnentTabOperations("IsMaximized","Content Search", "") = False Then  
					Call Fn_CPD_CompnentTabOperations("DoubleClick","Content Search", "")  
					Call Fn_ReadyStatusSync(2)
				End If 
				objTree.Select bReturn
				Select Case sAction
					Case "MoveRecipeUP"
							Fn_CPD_RecipeOperations= Fn_Button_Click("Fn_CPD_RecipeOperations",JavaWindow("Collaborative Product"), "MoveRecipeUP")
					Case "MoveRecipeDown"
							Fn_CPD_RecipeOperations= Fn_Button_Click("Fn_CPD_RecipeOperations",JavaWindow("Collaborative Product"), "MoveRecipeDown")
					Case "DeleteRecipe"
						    Fn_CPD_RecipeOperations= Fn_Button_Click("Fn_CPD_RecipeOperations",JavaWindow("Collaborative Product"), "DeleteRecipe")
				End Select
				If Fn_CPD_CompnentTabOperations("IsMaximized","Content Search", "") = True Then  
					Call Fn_CPD_CompnentTabOperations("DoubleClick","Content Search", "")  
					Call Fn_ReadyStatusSync(2)
				End If 
			End If
		Case "VerifyEnabledSearchOptions"
			Call Fn_SISW_UI_Twistie_Operations("Fn_CPD_RecipeOperations", "Expand", JavaWindow("Collaborative Product"), "Twistie", "Search Options","SearchOptions")
			Wait 1
			Set objSearchOptions = JavaWindow("Collaborative Product").JavaList("OptionsCombo")
			If objSearchOptions.GetROProperty("enabled") Then
				Fn_CPD_RecipeOperations = True
				Set objSearchOptions = Nothing
			Else
				Fn_CPD_RecipeOperations = False
				Set objSearchOptions = Nothing
			End If
		'[TC11.4(20171201.00)_NewDevelopment_PoonamC_28Dec2017 : Added Case to verify state of buttons in recipe tab ]	
		Case "VerifyEnabledMoveRecipeUP","VerifyEnabledMoveRecipeDown","VerifyEnabledDeleteRecipe"
			 If Fn_CPD_CompnentTabOperations("IsMaximized","Content Search", "") = False Then  
					Call Fn_CPD_CompnentTabOperations("DoubleClick","Content Search", "")  
					Call Fn_ReadyStatusSync(2)
			 End If
			 Select Case sAction
				Case "VerifyEnabledMoveRecipeUP"
						Fn_CPD_RecipeOperations= Fn_UI_Object_GetROProperty("Fn_CPD_RecipeOperations",JavaWindow("Collaborative Product").JavaButton("MoveRecipeUP"),"enabled")
				Case "VerifyEnabledMoveRecipeDown"
						Fn_CPD_RecipeOperations= Fn_UI_Object_GetROProperty("Fn_CPD_RecipeOperations",JavaWindow("Collaborative Product").JavaButton("MoveRecipeDown"),"enabled") 
				Case "VerifyEnabledDeleteRecipe"
					    Fn_CPD_RecipeOperations= Fn_UI_Object_GetROProperty("Fn_CPD_RecipeOperations",JavaWindow("Collaborative Product").JavaButton("DeleteRecipe"),"enabled") 
			End Select
			If Fn_CPD_CompnentTabOperations("IsMaximized","Content Search", "") = True Then  
				 Call Fn_CPD_CompnentTabOperations("DoubleClick","Content Search", "")  
				 Call Fn_ReadyStatusSync(2)
			 End If			
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		'[TC11.5(20180616b.00)_NewDevelopment_PoonamC_24Sept2018 : Added Case to verify Existence of buttons in recipe tab ]	
		Case "IsExistsMoveRecipeUP","IsExistsMoveRecipeDown","IsExistsDeleteRecipe"
			 If Fn_CPD_CompnentTabOperations("IsMaximized","Content Search", "") = False Then  
					Call Fn_CPD_CompnentTabOperations("DoubleClick","Content Search", "")  
					Call Fn_ReadyStatusSync(2)
			 End If
			 Select Case sAction
				Case "IsExistsMoveRecipeUP"
						Fn_CPD_RecipeOperations = Fn_SISW_UI_Object_Operations("Fn_CPD_RecipeOperations", "Exist", JavaWindow("Collaborative Product").JavaButton("MoveRecipeUP"),"")
				Case "IsExistsMoveRecipeDown"
						Fn_CPD_RecipeOperations = Fn_SISW_UI_Object_Operations("Fn_CPD_RecipeOperations", "Exist", JavaWindow("Collaborative Product").JavaButton("MoveRecipeDown"),"") 
				Case "IsExistsDeleteRecipe"
					    Fn_CPD_RecipeOperations = Fn_SISW_UI_Object_Operations("Fn_CPD_RecipeOperations", "Exist", JavaWindow("Collaborative Product").JavaButton("DeleteRecipe"),"")
			End Select
			If Fn_CPD_CompnentTabOperations("IsMaximized","Content Search", "") = True Then  
				 Call Fn_CPD_CompnentTabOperations("DoubleClick","Content Search", "")  
				 Call Fn_ReadyStatusSync(2)
			 End If			
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_CPD_RecipeOperations ] Invalid case [ " & sAction & " ].")
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
	End Select
	If  Fn_CPD_RecipeOperations <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [ Fn_CPD_RecipeOperations ] executed successfuly with case [ " & sAction & " ].")
	End If
	Set objTree = Nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name			:	Fn_CPD_EffectivityOperations
'@@
'@@    Description				:	Function Used to create Collaborative Design
'@@
'@@    Parameters			    :	1. sAction		: Action to be performed
'@@								:	2. sNode		: Node name with complete Path
'@@								:	3. sValue		: Value to be verified / Row Number 
'@@								:	4. sFromUnit 	: Unit Starting Value (~ separated list of From Units)
'@@								:	5. sToUnit		: Unit End Value (~ separated list of To Units)
'@@								:	6. sInDate		: In Date value (~ separated list of Date strings eg. 02-Jan-2012~02-Jan-2012$12:30 )
'@@								:	7. sOutDates	: Out Date value (~ separated list of Date strings 02-Jan-2012~SO~UP )
'@@
'@@    Return Value		   	   	: 	True Or False
'@@
'@@    Pre-requisite			:	CPD perspective should be activated.						
'@@
'@@    Examples					:	Call Fn_CPD_EffectivityOperations("ViewChangeEffectivity", "CD000006;1-dszg:DE000006/001;1-asq", "", "5", "10", "02-Jan-2012", "SO")
'@@    Examples					:	Call Fn_CPD_EffectivityOperations("SetInEffectivityTab", "CD000010;1-CD:DE000001/001;1-de", "", "5", "SO", "", "")
'@@    Examples					:	Call Fn_CPD_EffectivityOperations("VerifyInViewChangeEffectivity", "CD000006;1-dszg:DE000006/001;1-asq", "", "5~1", "10~16", "", "")
'@@    Examples					:	Call Fn_CPD_EffectivityOperations("VerifyInEffectivityTab", "CD000010;1-CD:DE000001/001;1-de", "", "5", "SO", "", "")
'@@    Examples					:	Call Fn_CPD_EffectivityOperations("VerifyEffectivityConfiguration", "", "Unit=1..5", "", "", "", "")
'@@    Examples					:	Call Fn_CPD_EffectivityOperations("VerifyColumnExistInViewChangeEffectivity", "", "Out Date", "", "", "", "")
'@@    Examples					:	Call Fn_CPD_EffectivityOperations("GetColumnCountInViewChangeEffectivity", "", "", "", "", "", "")
'@@    Examples					:	Call Fn_CPD_EffectivityOperations("VerifyColumnExistInEffectivityTab", "", "Out Date", "", "", "", "")
'@@    Examples					:	Call Fn_CPD_EffectivityOperations("GetColumnCountInEffectivityTab", "", "", "", "", "", "")
'@@    Examples					:	Call Fn_CPD_EffectivityOperations("SetInEffectivityTab", "", "0", "5", "SO", "", "")
'@@    Examples					:	Call Fn_CPD_EffectivityOperations("VerifyListContents", "", "0", "", "SO", "", "UP")
'@@    Examples					:	Call Fn_CPD_EffectivityOperations("DeleteInEffectivityTab","CD000006;1-dszg:DE000006/001;1-asq" , "", "10", "20", "02-Jan-2015", "03-Jan-2015")
'@@
'@@	   History					:	
'@@				Developer Name				Date			Rev. No.	Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe			16-Jan-2012			1.0			Created
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe			31-Jan-2012			1.1			Added cases VerifyInViewChangeEffectivity, VerifyInEffectivityTab
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe			03-Feb-2012			1.1			Added case VerifyEffectivityConfiguration
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe			03-Feb-2012			1.1			Added case VerifyColumnExistInViewChangeEffectivity, GetColumnCountInViewChangeEffectivity, VerifyColumnExistInEffectivityTab, GetColumnCountInEffectivityTab
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe			05-Feb-2012			1.1			Added case VerifyListContents
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Shweta Rathore			03-Dec-2015			1.1			Added case DeleteInEffectivityTab      		[TC1121-2015110900-03_12_2015-AnkitN-NewDevelopment]       
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_CPD_EffectivityOperations(sAction, sNode, sValue, sFromUnit, sToUnit, sInDate, sOutDate)
	GBL_FAILED_FUNCTION_NAME="Fn_CPD_EffectivityOperations"
	Dim objSelectType, intNoOfObjects, i, bReturn
	Dim aFromUnit, aToUnit, aInDate, aOutDate, aNode
	Dim objEffectivityTable, iLimit, bFlag, objEff,bStatus
	Dim arrDate,tFlag
	Dim iKeyCnt
	Dim iItemCount, iCnt
	Fn_CPD_EffectivityOperations = False
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
	
	
				
	Select Case sAction
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "VerifyEffectivityConfiguration"
			Set objSelectType = Description.Create()
			objSelectType("Class Name").value = "JavaObject"
			objSelectType("toolkit class").value = "org.eclipse.ui.forms.widgets.ImageHyperlink"
			Set  intNoOfObjects = JavaWindow("Collaborative Product").ChildObjects(objSelectType)
			For i = 0 to intNoOfObjects.count-1
				If  lcase(trim( "" & intNoOfObjects(i).Object.getToolTipText())) = lCase("Click to view and change the current effectivity configuration") Then
					IF intNoOfObjects(i).Object.getText() = sValue Then
						Fn_CPD_EffectivityOperations =True
						Exit for
					End IF
				End If
			Next
			
		tab = JavaWindow("Collaborative Product").JavaTab("TCComponentTab").GetROProperty("value")
        If tab="Content Search" Then
        	Call Fn_CPD_CompnentTabOperations("Close",tab, "") 
        	tab = JavaWindow("Collaborative Product").JavaTab("TCComponentTab").GetROProperty("value")
        End If
                    
		If Fn_CPD_EffectivityOperations = False And tab <> "" Then
			If  Fn_CPD_CompnentTabOperations("IsMaximized",tab, "")  = False Then
				tflag=1
				Call Fn_CPD_CompnentTabOperations("DoubleClick",tab, "") 
			End If
			
			Set objSelectType = Description.Create()
			objSelectType("Class Name").value = "JavaObject"
			objSelectType("toolkit class").value = "org.eclipse.ui.forms.widgets.ImageHyperlink"
			Set  intNoOfObjects = JavaWindow("Collaborative Product").ChildObjects(objSelectType)
						
			For iCnt = 0 to intNoOfObjects.count-1	
				IF  intNoOfObjects(iCnt).Object.getText() = sValue Then
					Fn_CPD_EffectivityOperations =True
					Exit for
				End IF
			Next
				
				If  tflag=1 Then
					tflag= 0 
					Call Fn_CPD_CompnentTabOperations("DoubleClick",tab, "") 
				End If				
		End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	    Case "ViewChangeEffectivity", "VerifyInViewChangeEffectivity", "VerifyInViewChangeEffectivity_Ext","VerifyColumnExistInViewChangeEffectivity", "ActivateAndTypeInViewChangeEffectivity" , "GetColumnCountInViewChangeEffectivity","ViewChangeEffectivity_Ext"
				If sNode <> "" Then
					bReturn = Fn_CPD_ContentExplorer("Select", sNode, "", "", "")
					If bReturn = False Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_CPD_EffectivityOperations : Failed to select Node [ " & sNode & " ].")
						Exit function
					End If
				End If
	
				If Fn_UI_ObjectExist("Fn_CPD_EffectivityOperations",JavaWindow("Collaborative Product").JavaWindow("ViewEditEffectivity")) = False Then
						Set objSelectType = Description.Create()
						objSelectType("Class Name").value = "JavaObject"
						objSelectType("toolkit class").value = "org.eclipse.ui.forms.widgets.ImageHyperlink"
						Set  intNoOfObjects = JavaWindow("Collaborative Product").ChildObjects(objSelectType)
						For i = 0 to intNoOfObjects.count-1
							If  lcase(trim( "" & intNoOfObjects(i).Object.getToolTipText())) = lCase("Click to view and change the current effectivity configuration") Then
								intNoOfObjects(i).Click 1,1, "LEFT"
								Exit for
							End If
						Next
						
					'modified by pratap
						If Fn_UI_ObjectExist("Fn_CPD_EffectivityOperations",JavaWindow("Collaborative Product").JavaWindow("ViewEditEffectivity")) = False  Then
							tab = JavaWindow("Collaborative Product").JavaTab("TCComponentTab").GetROProperty("value")
					        If tab="Content Search" Then
					        	Call Fn_CPD_CompnentTabOperations("Close",tab, "") 
					        	tab = JavaWindow("Collaborative Product").JavaTab("TCComponentTab").GetROProperty("value")
					        End If

							If tab <> "" Then
								If  Fn_CPD_CompnentTabOperations("IsMaximized",tab, "")  = False Then
									tflag=1
									Call Fn_CPD_CompnentTabOperations("DoubleClick",tab, "") 
								End If
									
								Set objSelectType = Description.Create()
								objSelectType("Class Name").value = "JavaObject"
								objSelectType("toolkit class").value = "org.eclipse.ui.forms.widgets.ImageHyperlink"
								Set  intNoOfObjects = JavaWindow("Collaborative Product").ChildObjects(objSelectType)
						
									For iCnt = 0 to intNoOfObjects.count-1	
										IF instr(1,trim(intNoOfObjects(iCnt).Object.getText()),"Unit")>0 OR instr(1,trim(intNoOfObjects(iCnt).Object.getText()),"Date")>0 OR instr(1,trim(intNoOfObjects(iCnt).Object.getText()),"Effectivity")>0 Then
											intNoOfObjects(iCnt).object.setFocus
											intNoOfObjects(iCnt).Click 1, 1 ,"LEFT" 
											wait 1
											If JavaWindow("Collaborative Product").JavaWindow("ViewEditEffectivity").Exist(3) = False Then
												intNoOfObjects(iCnt).Click 20, 0 ,"LEFT" 
											End If
											Exit for
										End IF
									Next
						
							End If
					End If
						If Fn_UI_ObjectExist("Fn_CPD_EffectivityOperations",JavaWindow("Collaborative Product").JavaWindow("ViewEditEffectivity")) = False Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_CPD_EffectivityOperations : Failed to find View / Edit Effectivity window.")
							Exit function
						End If
				End If
	
				Set objEff = JavaWindow("Collaborative Product").JavaWindow("ViewEditEffectivity")
				Set objEffectivityTable = objEff.JavaTable("EffectivityTable")
				
				Select Case sAction
					' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
					Case "ViewChangeEffectivity"
						For i = 0 to iLimit
							If sFromUnit <> "" AND uBound(aToUnit) <= iLimit Then
								' setting From Unit
								objEffectivityTable.ActivateCell i,"From Unit"
								wait 1
								If trim(aFromUnit(i)) = "" Then
									wait 1
									Call Fn_KeyBoardOperation("SendKeys", "{END}")
									For iKeyCnt = 0 to 10
										Call Fn_KeyBoardOperation("SendKeys", "{BKSP}")
									Next
								End If
								objEff.JavaEdit("TableText").Set aFromUnit(i)
								objEff.JavaEdit("TableText").Activate
								wait 1

								' setting To Unit
								objEffectivityTable.ActivateCell i,"To Unit"
								wait 1
								Call Fn_KeyBoardOperation("SendKeys", "{END}")
								For iKeyCnt = 0 to 10
									Call Fn_KeyBoardOperation("SendKeys", "{BKSP}")
								Next
								If sToUnit<>"" Then	
									If trim(aToUnit(i)) <> "" Then
										objEff.JavaList("TableList").Type aToUnit(i)
									End If
									Call Fn_KeyBoardOperation("SendKeys", "{ENTER}")
									wait 1
								End If
							End If
							If sInDate <> ""   AND uBound(aOutDate) <= iLimit Then
								' setting IN Date
								objEffectivityTable.ActivateCell i,"In Date"
								wait 1
								Call Fn_Button_Click("Fn_CPD_EffectivityOperations", objEff, "TableDropDownList")
								wait 1
								If instr( aInDate(i),"$") > 0 Then
										arrDate = split(trim(aInDate(i)),"$")
										Call  Fn_CPD_DateControl("Set", arrDate(0), arrDate(1))
								Else
										Select Case lcase(trim(aInDate(i)))
											Case ""
												Call  Fn_CPD_DateControl("Clear", "", "")
											Case "today"
												Call  Fn_CPD_DateControl("Today", "", "")
											Case Else
												Call  Fn_CPD_DateControl("Set", aInDate(i), "")
										End Select
								End If

								' setting Out Date
								objEffectivityTable.ActivateCell i, "Out Date"
								wait 1
								If lcase(aOutDate(i)) <> "so" AND lcase(aOutDate(i)) <> "up" Then
									Call Fn_List_Select("Fn_CPD_EffectivityOperations",objEff,"TableList","Select Date...")
									If instr( aOutDate(i),"$") > 0 Then
											arrDate = split(trim(aOutDate(i)),"$")
											Call  Fn_CPD_DateControl("Set", arrDate(0), arrDate(1))
									Else
											Select Case lcase(trim(aOutDate(i)(i)))
												Case ""
													Call  Fn_CPD_DateControl("Clear", "", "")
												Case "today"
													Call  Fn_CPD_DateControl("Today", "", "")
												Case Else
													Call  Fn_CPD_DateControl("Set", aOutDate(i), "")
											End Select
									End If
								else
									Call Fn_List_Select("Fn_CPD_EffectivityOperations",objEff,"TableList",aOutDate(i))
								End If
							End If
						Next
						Call Fn_Button_Click("Fn_CPD_EffectivityOperations", JavaWindow("Collaborative Product").JavaWindow("ViewEditEffectivity"),"OK")
						Fn_CPD_EffectivityOperations = True
					' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
					Case "ViewChangeEffectivity_Ext"
						For i = 0 to iLimit
							If sFromUnit <> "" OR uBound(aToUnit) <= iLimit Then
								
								If sFromUnit <> "" Then
									' setting From Unit
									Call Fn_ReadyStatusSync(2)
									objEffectivityTable.ActivateCell i,"From Unit"
									wait 1
									If trim(aFromUnit(i)) = "" Then
										wait 1
										Call Fn_KeyBoardOperation("SendKeys", "{END}")
										For iKeyCnt = 0 to 10
											Call Fn_KeyBoardOperation("SendKeys", "{BKSP}")
										Next
									End If
									objEff.JavaEdit("TableText").Set aFromUnit(i)
									objEff.JavaEdit("TableText").Activate
									wait 1									
								End If
								
								If sToUnit <> "" Then
									' setting To Unit
									Call Fn_ReadyStatusSync(2)
									objEffectivityTable.ActivateCell i,"To Unit"
									wait 1
									Call Fn_KeyBoardOperation("SendKeys", "{END}")
									For iKeyCnt = 0 to 10
										Call Fn_KeyBoardOperation("SendKeys", "{BKSP}")
									Next
									If sToUnit<>"" Then	
										If trim(aToUnit(i)) <> "" Then
											objEff.JavaList("TableList").Type aToUnit(i)
										End If
										Call Fn_KeyBoardOperation("SendKeys", "{ENTER}")
										wait 1
									End If
								End If								
							End If
							
							If sInDate <> ""   OR uBound(aOutDate) <= iLimit Then
								' setting IN Date
								If sInDate<>"" Then
								    Call Fn_ReadyStatusSync(2)
									objEffectivityTable.ActivateCell i,"In Date"
									wait 1
									Call Fn_Button_Click("Fn_CPD_EffectivityOperations", objEff, "TableDropDownList")
									wait 1
									If instr( aInDate(i),"$") > 0 Then
										arrDate = split(trim(aInDate(i)),"$")
										Call  Fn_CPD_DateControl("Set", arrDate(0), arrDate(1))
									Else
										Select Case lcase(trim(aInDate(i)))
											Case "Clear"
												Call  Fn_CPD_DateControl("Clear", "", "")
											Case "today"
												Call  Fn_CPD_DateControl("Today", "", "")
											Case Else
												Call  Fn_CPD_DateControl("Set", aInDate(i), "")
										End Select
									End If
								End If
								
								If sOutDate<>"" Then
									' setting Out Date
									Call Fn_ReadyStatusSync(2)
									objEffectivityTable.ActivateCell i, "Out Date"
									wait 1
									If lcase(aOutDate(i)) <> "so" OR lcase(aOutDate(i)) <> "up" Then
										Call Fn_List_Select("Fn_CPD_EffectivityOperations",objEff,"TableList","Select Date...")
										If instr( aOutDate(i),"$") > 0 Then
											arrDate = split(trim(aOutDate(i)),"$")
											Call  Fn_CPD_DateControl("Set", arrDate(0), arrDate(1))
										Else
											Select Case lcase(trim(aOutDate(i)))
												Case "Clear"
													Call  Fn_CPD_DateControl("Clear", "", "")
												Case "today"
													Call  Fn_CPD_DateControl("Today", "", "")
												Case Else
													Call  Fn_CPD_DateControl("Set", aOutDate(i), "")
											End Select
										End If
									Else
										Call Fn_List_Select("Fn_CPD_EffectivityOperations",objEff,"TableList",aOutDate(i))
									End If									
								End If								
							End If
						Next
						Call Fn_Button_Click("Fn_CPD_EffectivityOperations", JavaWindow("Collaborative Product").JavaWindow("ViewEditEffectivity"),"OK")
						Fn_CPD_EffectivityOperations = True
										
					' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
					Case "VerifyInViewChangeEffectivity"
						bFlag = True
						For i = 0 to iLimit
							If sFromUnit <> "" AND uBound(aToUnit) <= iLimit Then
								objEffectivityTable.ActivateCell i,"From Unit"
								wait 1
								
								If cstr(objEffectivityTable.GetCellData(i, "From Unit")) <> cstr(aFromUnit(i)) Then
									bFlag = False
								End If
								
								objEffectivityTable.ActivateCell i,"To Unit"
								wait 1
								
								If cstr(objEffectivityTable.GetCellData(i, "To Unit")) <> cstr(aToUnit(i)) Then
									bFlag = False
								End If
							End If
							If sInDate <> ""   AND uBound(aOutDate) <= iLimit Then
								If cstr(objEffectivityTable.GetCellData(i, "In Date")) <> cstr(aInDate(i)) Then
									bFlag = False
								End If
								If cstr(objEffectivityTable.GetCellData(i, "Out Date")) <> cstr(aOutDate(i)) Then
									bFlag = False
								End If
							End If
							If bFlag = False Then 
								Exit For
							End If
						Next
						Call Fn_Button_Click("Fn_CPD_EffectivityOperations", JavaWindow("Collaborative Product").JavaWindow("ViewEditEffectivity"),"Cancel")
						Fn_CPD_EffectivityOperations = bFlag
					' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
					Case "VerifyInViewChangeEffectivity_Ext"
						bFlag = True
						For i = 0 to iLimit
							If sFromUnit <> "" OR uBound(aToUnit) <= iLimit Then
								
								If sFromUnit <> "" Then
									' setting From Unit
									Call Fn_ReadyStatusSync(2)
									objEffectivityTable.ActivateCell i,"From Unit"
									wait 1
									If cstr(objEffectivityTable.GetCellData(i, "From Unit")) <> cstr(aFromUnit(i)) Then
									 	bFlag = False
								    End If
									wait 1									
								End If
								
								If sToUnit <> "" Then
									' setting To Unit
									Call Fn_ReadyStatusSync(2)
									objEffectivityTable.ActivateCell i,"To Unit"
									wait 1
									If cstr(objEffectivityTable.GetCellData(i, "To Unit")) <> cstr(aToUnit(i)) Then
										bFlag = False
									End If
								End If								
							End If
							
							If sInDate <> ""   OR uBound(aOutDate) <= iLimit Then
								' setting IN Date
								If sInDate<>"" Then
								    Call Fn_ReadyStatusSync(2)
									objEffectivityTable.ActivateCell i,"In Date"
									wait 1
									If cstr(objEffectivityTable.GetCellData(i, "In Date")) <> cstr(aInDate(i)) Then
										bFlag = False
									End If
								End If
								
								If sOutDate<>"" Then
									' setting Out Date
									Call Fn_ReadyStatusSync(2)
									objEffectivityTable.ActivateCell i, "Out Date"
									wait 1
									If cstr(objEffectivityTable.GetCellData(i, "Out Date")) <> cstr(aOutDate(i)) Then
										bFlag = False
									End If					
							    End If
						    End If
						Next
						Call Fn_Button_Click("Fn_CPD_EffectivityOperations", JavaWindow("Collaborative Product").JavaWindow("ViewEditEffectivity"),"Cancel")
						Fn_CPD_EffectivityOperations = bFlag
					' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
					Case "VerifyColumnExistInViewChangeEffectivity"
						If objEffectivityTable.Exist(5) Then
							If inStr(objEffectivityTable.GetROProperty("column names"), sValue) > 0 Then
								Fn_CPD_EffectivityOperations = True
							End If
						End If
						Call Fn_Button_Click("Fn_CPD_EffectivityOperations", JavaWindow("Collaborative Product").JavaWindow("ViewEditEffectivity"),"Cancel")
					' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
					Case "GetColumnCountInViewChangeEffectivity"
						Fn_CPD_EffectivityOperations = -1
						If objEffectivityTable.Exist(5) Then
							Fn_CPD_EffectivityOperations = objEffectivityTable.GetROProperty("cols")
						End If
						Call Fn_Button_Click("Fn_CPD_EffectivityOperations", JavaWindow("Collaborative Product").JavaWindow("ViewEditEffectivity"),"Cancel")
					' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
					Case "ActivateAndTypeInViewChangeEffectivity"
						If sFromUnit <> "" Then
							objEffectivityTable.ActivateCell sValue, "From Unit"
							If trim(sFromUnit) = "" Then
								wait 1
								Call Fn_KeyBoardOperation("SendKeys", "{END}")
								For iKeyCnt = 0 to 10
									Call Fn_KeyBoardOperation("SendKeys", "{BKSP}")
								Next
							End If
							objEff.JavaEdit("TableText").Set sFromUnit
						End If 
						If sToUnit <> "" Then
							objEffectivityTable.ActivateCell sValue, "To Unit"
							objEff.JavaList("TableList").Select sToUnit
						End If 
						If sInDate <> "" Then
							objEffectivityTable.ActivateCell sValue, "In Date"
							If trim(sInDate) = "" Then
								wait 1
								Call Fn_KeyBoardOperation("SendKeys", "{END}")
								For iKeyCnt = 0 to 10
									Call Fn_KeyBoardOperation("SendKeys", "{BKSP}")
								Next
							End If
							objEff.JavaEdit("TableText").Set sInDate
						End If
						If sOutDate <> "" Then
							objEffectivityTable.ActivateCell sValue, "Out Date"
							objEff.JavaList("TableList").Select sOutDate
						End If
						Fn_CPD_EffectivityOperations = True
					' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				End Select
				
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "SetInEffectivityTab_WithoutSave" , "SetInEffectivityTab", "VerifyInEffectivityTab", "VerifyColumnExistInEffectivityTab","GetColumnCountInEffectivityTab", "ActivateAndTypeInEffectivityTab", "VerifyListContents","DeleteInEffectivityTab","DeleteInEffectivityTab_Ext","CheckEffectivityEmpty","ClearcellData", "VerifyBlankInEffectivityTab","SetInEffectivityForMultipleNode","VerifyEffectivityForMultipleNode"
				Wait 1
				If sNode <> "" Then
					If Fn_TabFolder_Operation("Exist","Effectivity","")=False Then
						If Fn_TabFolder_Operation("Exist","*Effectivity","")=False Then
							bReturn = Fn_CPD_ContentExplorer("PopupMenuSelect", sNode, "", "", "Open with:Effectivity")
							If bReturn = False Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_CPD_EffectivityOperations : Failed to select popup menu of [ " & sNode & " ].")
								Exit function
							End If
						End If
					End If
				End If
				Wait 5
				Set objEff = JavaWindow("Collaborative Product")
				aNode=Split(sNode,":")
				Select Case sAction
					' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
					Case "SetInEffectivityTab" , "SetInEffectivityTab_WithoutSave"      '[TC1123-20170410-14_04_2017-JotibaT-Maintenance] - Updated By Jotiba (Changed object JavaTable to JavaTree)
					For i = 0 to iLimit
						bStatus=False
							If Fn_CPD_EffectivityOperations("VerifyColumnExistInEffectivityTab", "", "To Unit", "", "", "", "")=True Then
								Call Fn_UI_ClickJavaTreeCell("Fn_CPD_EffectivityOperations", JavaWindow("Collaborative Product"), "EffectivityTree", aNode(Ubound(aNode)), "To Unit","LEFT")
							End If
						If sFromUnit<>"" Then
							If aFromUnit(i)<>"" Then
								'setting From Unit
									Call Fn_UI_ClickJavaTreeCell("Fn_CPD_EffectivityOperations", JavaWindow("Collaborative Product"), "EffectivityTree", aNode(Ubound(aNode)), "From Unit","LEFT")
									wait 1
									Set objShell = CreateObject("Wscript.Shell")
									If  i = 0 Then
										Call Fn_KeyBoardOperation("SendKeys", "{DOWN}")
										wait 1
										Call Fn_KeyBoardOperation("SendKeys", "{ENTER}")
									Else
										Call Fn_Button_Click("Fn_CPD_EffectivityOperations", JavaWindow("Collaborative Product"),"AddButton")
										bStatus=True
										wait 1
										Call Fn_UI_ClickJavaTreeCell("Fn_CPD_EffectivityOperations", JavaWindow("Collaborative Product"), "EffectivityTree", aNode(Ubound(aNode)), "From Unit","LEFT")
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
							' setting To Unit
							If sToUnit<>"" Then
								If aToUnit(i)<>"" Then	
									If bStatus=False and i>0 Then
										Call Fn_Button_Click("Fn_CPD_EffectivityOperations", JavaWindow("Collaborative Product"),"AddButton")
										bStatus=True
									End If 
									Call Fn_UI_ClickJavaTreeCell("Fn_CPD_EffectivityOperations", JavaWindow("Collaborative Product"), "EffectivityTree", aNode(Ubound(aNode)), "To Unit","LEFT")
									wait 1
									Set objShell = CreateObject("Wscript.Shell")
										For iKeyCnt = 0 To i
											Call Fn_KeyBoardOperation("SendKeys", "{DOWN}")
										Next
										wait 1
										Call Fn_KeyBoardOperation("SendKeys", "{ENTER}")
										wait 1
										objShell.SendKeys "^A"
										wait 1
									If trim(aToUnit(i)) <> "" Then
										objEff.JavaList("TableList").Type aToUnit(i)
									End If
									wait 1
									Call Fn_KeyBoardOperation("SendKeys", "{ENTER}")
									Call Fn_KeyBoardOperation("SendKeys", "{RIGHT}")
									Set objShell=Nothing
								End If
							End If 
						End If
							' setting IN Date
							If sInDate<>"" Then
						
									If Fn_CPD_EffectivityOperations("VerifyColumnExistInEffectivityTab", "", "From Unit", "", "", "", "")=True And sFromUnit="" Then
										Call Fn_UI_ClickJavaTreeCell("Fn_CPD_EffectivityOperations", JavaWindow("Collaborative Product"), "EffectivityTree", aNode(Ubound(aNode)), "From Unit","LEFT")
										Call Fn_KeyBoardOperation("SendKeys", "{RIGHT}")
										Call Fn_KeyBoardOperation("SendKeys", "{RIGHT}")
									End If
								If aInDate(i)<>"" Then
									If bStatus=False and i>0 Then
										Call Fn_Button_Click("Fn_CPD_EffectivityOperations", JavaWindow("Collaborative Product"),"AddButton")
										bStatus=True
										Set objShell = CreateObject("Wscript.Shell")
										wait 1
										For iKeyCnt=0 To i
											Call Fn_KeyBoardOperation("SendKeys", "{TAB}")
										Next
										Set objShell=Nothing 
									End If
									
									Call Fn_SyncTCObjects()
									Call Fn_UI_ClickJavaTreeCell("Fn_CPD_EffectivityOperations", JavaWindow("Collaborative Product"), "EffectivityTree", aNode(Ubound(aNode)), "In Date","LEFT")
									wait 1
									Set objShell = CreateObject("Wscript.Shell")
									wait 1
									For iKeyCnt=0 To i
										Call Fn_KeyBoardOperation("SendKeys", "{DOWN}")
									Next
									Call Fn_KeyBoardOperation("SendKeys", "{ENTER}")
									wait 1
									Call Fn_List_Select("",objEff,"TableList","Select Date...")
									If instr( aInDate(i),"$") > 0 Then
										arrDate = split(trim(aInDate(i)),"$")
										Call  Fn_CPD_DateControl("Set", arrDate(0), arrDate(1))
									Else
										Select Case lcase(trim(aInDate(i)))
											Case ""
												Call  Fn_CPD_DateControl("Clear", "", "")
											Case "today"
												Call  Fn_CPD_DateControl("Today", "", "")
											Case Else
												Call  Fn_CPD_DateControl("Set", aInDate(i), "")
										End Select
									End If
								End If 
									Call Fn_KeyBoardOperation("SendKeys", "{ESC}")
									Call Fn_KeyBoardOperation("SendKeys", "{RIGHT}")								
							End If
							'Set Out Date
							If sOutDate<>"" Then
								If aOutDate(i)<>"" Then
									If lcase(aOutDate(i)) <> "so" AND lcase(aOutDate(i)) <> "up" Then
											If bStatus=False and i>0  Then
												Call Fn_Button_Click("Fn_CPD_EffectivityOperations", JavaWindow("Collaborative Product"),"AddButton")
												bStatus=True
											End If 
											Call Fn_SyncTCObjects()
											Call Fn_UI_ClickJavaTreeCell("Fn_CPD_EffectivityOperations", JavaWindow("Collaborative Product"), "EffectivityTree", aNode(Ubound(aNode)), "Out Date","LEFT")
											wait 1
											Set objShell = CreateObject("Wscript.Shell")
											For iKeyCnt=0 To i
												Call Fn_KeyBoardOperation("SendKeys", "{DOWN}")
											Next
											Call Fn_KeyBoardOperation("SendKeys", "{ENTER}")
											wait 1
											Call Fn_List_Select("Fn_CPD_EffectivityOperations",objEff,"TableList","Select Date...")
											If instr( aOutDate(i),"$") > 0 Then
													arrDate = split(trim(aOutDate(i)),"$")
													Call  Fn_CPD_DateControl("Set", arrDate(0), arrDate(1))
											Else
													Select Case lcase(trim(aOutDate(i)))
														Case ""
															Call  Fn_CPD_DateControl("Clear", "", "")
														Case "today"
															Call  Fn_CPD_DateControl("Today", "", "")
														Case Else
															Call  Fn_CPD_DateControl("Set", aOutDate(i), "")
													End Select
											End If
									Else
											Call Fn_UI_ClickJavaTreeCell("Fn_CPD_EffectivityOperations", JavaWindow("Collaborative Product"), "EffectivityTree", aNode(Ubound(aNode)), "Out Date","LEFT")
											wait 1
											Set objShell = CreateObject("Wscript.Shell")
											For iKeyCnt=0 To i
												Call Fn_KeyBoardOperation("SendKeys", "{DOWN}")
											Next
											Call Fn_KeyBoardOperation("SendKeys", "{ENTER}")
											wait 1
											Call Fn_List_Select("Fn_CPD_EffectivityOperations",objEff,"TableList",aOutDate(i))
											Set objShell=Nothing
									End If
								End If
							End If
						Next 	

						If sAction = "SetInEffectivityTab" Then
'							Call Fn_Button_Click("Fn_CPD_EffectivityOperations", JavaWindow("Collaborative Product"),"Save")
							Call Fn_ToolbarOperation("Click", "Save (Ctrl+S)","")
						End If
						Fn_CPD_EffectivityOperations = True
					'------------------------------------------------------------------
					Case "SetInEffectivityForMultipleNode"    	 '[TC1123-20170410-24_05_2017-JotibaT-Maintenance] - Updated By Jotiba (Changed object JavaTable to JavaTree)
					For i = 0 to iLimit
						bStatus=False
							If Fn_CPD_EffectivityOperations("VerifyColumnExistInEffectivityTab", "", "To Unit", "", "", "", "")=True Then
								Call Fn_ClickEffectivityTreeCell("Fn_CPD_EffectivityOperations", JavaWindow("Collaborative Product"), "EffectivityTree", aNode(Ubound(aNode)), "To Unit","LEFT")
							End If
						If sFromUnit<>"" Then
							If aFromUnit(i)<>"" Then
								'setting From Unit
									Call Fn_ClickEffectivityTreeCell("Fn_CPD_EffectivityOperations", JavaWindow("Collaborative Product"), "EffectivityTree", aNode(Ubound(aNode)), "From Unit","LEFT")
									wait 1
									Set objShell = CreateObject("Wscript.Shell")
									If  i = 0 Then
										Call Fn_KeyBoardOperation("SendKeys", "{DOWN}")
										wait 1
										Call Fn_KeyBoardOperation("SendKeys", "{ENTER}")
									Else
										Call Fn_Button_Click("Fn_CPD_EffectivityOperations", JavaWindow("Collaborative Product"),"AddButton")
										bStatus=True
										wait 1
										Call Fn_ClickEffectivityTreeCell("Fn_CPD_EffectivityOperations", JavaWindow("Collaborative Product"), "EffectivityTree", aNode(Ubound(aNode)), "From Unit","LEFT")
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
							' setting To Unit
							If sToUnit<>"" Then
								If aToUnit(i)<>"" Then	
									If bStatus=False and i>0 Then
										Call Fn_Button_Click("Fn_CPD_EffectivityOperations", JavaWindow("Collaborative Product"),"AddButton")
										bStatus=True
									End If 
									Call Fn_ClickEffectivityTreeCell("Fn_CPD_EffectivityOperations", JavaWindow("Collaborative Product"), "EffectivityTree", aNode(Ubound(aNode)), "To Unit","LEFT")
									wait 1
									Set objShell = CreateObject("Wscript.Shell")
										For iKeyCnt = 0 To i
											Call Fn_KeyBoardOperation("SendKeys", "{DOWN}")
										Next
										wait 1
										Call Fn_KeyBoardOperation("SendKeys", "{ENTER}")
										wait 1
										objShell.SendKeys "^A"
										wait 1
									If trim(aToUnit(i)) <> "" Then
										objEff.JavaList("TableList").Type aToUnit(i)
									End If
									wait 1
									Call Fn_KeyBoardOperation("SendKeys", "{ENTER}")
									Call Fn_KeyBoardOperation("SendKeys", "{RIGHT}")
									Set objShell=Nothing
								End If
							End If 
						End If
							' setting IN Date
							If sInDate<>"" Then
								If aInDate(i)<>"" Then
									If bStatus=False and i>0 Then
										Call Fn_Button_Click("Fn_CPD_EffectivityOperations", JavaWindow("Collaborative Product"),"AddButton")
										bStatus=True
										Set objShell = CreateObject("Wscript.Shell")
										wait 1
										For iKeyCnt=0 To i
											Call Fn_KeyBoardOperation("SendKeys", "{TAB}")
										Next
										Set objShell=Nothing 
									End If
									
									Call Fn_SyncTCObjects()
									Call Fn_ClickEffectivityTreeCell("Fn_CPD_EffectivityOperations", JavaWindow("Collaborative Product"), "EffectivityTree", aNode(Ubound(aNode)), "In Date","LEFT")
									wait 1
									Set objShell = CreateObject("Wscript.Shell")
									wait 1
									For iKeyCnt=0 To i
										Call Fn_KeyBoardOperation("SendKeys", "{DOWN}")
									Next
									Call Fn_KeyBoardOperation("SendKeys", "{ENTER}")
									wait 1
									Call Fn_List_Select("",objEff,"TableList","Select Date...")
									If instr( aInDate(i),"$") > 0 Then
										arrDate = split(trim(aInDate(i)),"$")
										Call  Fn_CPD_DateControl("Set", arrDate(0), arrDate(1))
									Else
										Select Case lcase(trim(aInDate(i)))
											Case ""
												Call  Fn_CPD_DateControl("Clear", "", "")
											Case "today"
												Call  Fn_CPD_DateControl("Today", "", "")
											Case Else
												Call  Fn_CPD_DateControl("Set", aInDate(i), "")
										End Select
									End If
								End If 
								Call Fn_KeyBoardOperation("SendKeys", "{ENTER}")
								Call Fn_KeyBoardOperation("SendKeys", "{RIGHT}")
							End If
							'Set Out Date
							If sOutDate<>"" Then
								If aOutDate(i)<>"" Then
									If lcase(aOutDate(i)) <> "so" AND lcase(aOutDate(i)) <> "up" Then
											If bStatus=False and i>0  Then
												Call Fn_Button_Click("Fn_CPD_EffectivityOperations", JavaWindow("Collaborative Product"),"AddButton")
												bStatus=True
											End If 
											Call Fn_SyncTCObjects()
											Call Fn_ClickEffectivityTreeCell("Fn_CPD_EffectivityOperations", JavaWindow("Collaborative Product"), "EffectivityTree", aNode(Ubound(aNode)), "Out Date","LEFT")
											wait 1
											Set objShell = CreateObject("Wscript.Shell")
											For iKeyCnt=0 To i
												Call Fn_KeyBoardOperation("SendKeys", "{DOWN}")
											Next
											Call Fn_KeyBoardOperation("SendKeys", "{ENTER}")
											wait 1
											Call Fn_List_Select("Fn_CPD_EffectivityOperations",objEff,"TableList","Select Date...")
											If instr( aOutDate(i),"$") > 0 Then
													arrDate = split(trim(aOutDate(i)),"$")
													Call  Fn_CPD_DateControl("Set", arrDate(0), arrDate(1))
											Else
													Select Case lcase(trim(aOutDate(i)))
														Case ""
															Call  Fn_CPD_DateControl("Clear", "", "")
														Case "today"
															Call  Fn_CPD_DateControl("Today", "", "")
														Case Else
															Call  Fn_CPD_DateControl("Set", aOutDate(i), "")
													End Select
											End If
									Else
											Call Fn_ClickEffectivityTreeCell("Fn_CPD_EffectivityOperations", JavaWindow("Collaborative Product"), "EffectivityTree", aNode(Ubound(aNode)), "Out Date","LEFT")
											wait 1
											Set objShell = CreateObject("Wscript.Shell")
											For iKeyCnt=0 To i
												Call Fn_KeyBoardOperation("SendKeys", "{DOWN}")
											Next
											Call Fn_KeyBoardOperation("SendKeys", "{ENTER}")
											wait 1
											Call Fn_List_Select("Fn_CPD_EffectivityOperations",objEff,"TableList",aOutDate(i))
											Set objShell=Nothing
									End If
								End If
							End If
						Next 	

						If sAction = "SetInEffectivityForMultipleNode" Then
							Call Fn_ToolbarOperation("Click", "Save (Ctrl+S)","")
						End If
						Fn_CPD_EffectivityOperations = True					
					' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
					Case "VerifyInEffectivityTab"							'[TC1123-20170410-14_04_2017-JotibaT-Maintenance] - Updated By Jotiba (Changed object JavaTable to JavaTree)
						bFlag = True
						For i = 0 to iLimit
							' Verify From Unit
							If sFromUnit <> "" Then
								If aFromUnit(i)<>"" Then
								Call Fn_UI_ClickJavaTreeCell("Fn_CPD_EffectivityOperations", JavaWindow("Collaborative Product"), "EffectivityTree", aNode(Ubound(aNode)), "From Unit","LEFT")
								wait 1
								Set objShell = CreateObject("Wscript.Shell")
								For iKeyCnt=0 To i
									Call Fn_KeyBoardOperation("SendKeys", "{DOWN}")
								Next
								Call Fn_KeyBoardOperation("SendKeys", "{ENTER}")
								wait 1
								objShell.SendKeys "^A"
								wait 1
								If cstr(aFromUnit(i)) <> cstr(JavaWindow("Collaborative Product").JavaEdit("TableText").GetROProperty("text")) Then
									bFlag = False
								End If
								Call Fn_KeyBoardOperation("SendKeys", "{ESC}")
								Call Fn_KeyBoardOperation("SendKeys", "{RIGHT}")
								Set objShell=Nothing
								End If 
							 End If 
													 
							 ' Verify To Unit
							 If sToUnit<> "" Then
								 If sFromUnit = ""  Then
								 	Call Fn_UI_ClickJavaTreeCell("Fn_CPD_EffectivityOperations", JavaWindow("Collaborative Product"), "EffectivityTree", aNode(Ubound(aNode)), "From Unit","LEFT")
								 	Call Fn_KeyBoardOperation("SendKeys", "{RIGHT}")
								 End If
							 	If aToUnit(i)<>"" Then	
							 	Call Fn_UI_ClickJavaTreeCell("Fn_CPD_EffectivityOperations", JavaWindow("Collaborative Product"), "EffectivityTree", aNode(Ubound(aNode)), "To Unit","LEFT")
								wait 1
								Set objShell = CreateObject("Wscript.Shell")
								For iKeyCnt=0 To i
									Call Fn_KeyBoardOperation("SendKeys", "{DOWN}")
								Next
								Call Fn_KeyBoardOperation("SendKeys", "{ENTER}")
								wait 1
								If cstr(aToUnit(i)) <> cstr(JavaWindow("Collaborative Product").JavaList("TableList").GetROProperty("value")) Then
									bFlag = False
								End If
								Call Fn_KeyBoardOperation("SendKeys", "{ESC}")
								'Call Fn_KeyBoardOperation("SendKeys", "{TAB}")
								Call Fn_KeyBoardOperation("SendKeys", "{RIGHT}")
								Set objShell=Nothing
								End If 
							 End If
							 ' Vrify In Date 
							 	If sInDate <> ""  Then
									If Fn_CPD_EffectivityOperations("VerifyColumnExistInEffectivityTab", "", "From Unit", "", "", "", "")=True And sFromUnit="" Then
										Call Fn_UI_ClickJavaTreeCell("Fn_CPD_EffectivityOperations", JavaWindow("Collaborative Product"), "EffectivityTree", aNode(Ubound(aNode)), "From Unit","LEFT")
										Call Fn_KeyBoardOperation("SendKeys", "{RIGHT}")
										Call Fn_KeyBoardOperation("SendKeys", "{RIGHT}")
									End If
							 		If aInDate(i)<>"" Then
							 		Call Fn_UI_ClickJavaTreeCell("Fn_CPD_EffectivityOperations", JavaWindow("Collaborative Product"), "EffectivityTree", aNode(Ubound(aNode)), "In Date","LEFT")
									wait 1
									Set objShell = CreateObject("Wscript.Shell")
									For iKeyCnt=0 To i
										Call Fn_KeyBoardOperation("SendKeys", "{DOWN}")
									Next
									Call Fn_KeyBoardOperation("SendKeys", "{ENTER}")
									wait 1
									If Trim(cstr(aInDate(i))) <> Trim(cstr(JavaWindow("Collaborative Product").JavaList("TableList").GetROProperty("value"))) Then
										bFlag = False
									End If
									Call Fn_KeyBoardOperation("SendKeys", "{ESC}")
									Call Fn_KeyBoardOperation("SendKeys", "{RIGHT}")
									Set objShell=Nothing
									End If 
							 	End If
							 	'Verify Out Date
							 	If sOutDate<>"" Then
							 		If aOutDate(i)<>"" Then
							 		Call Fn_UI_ClickJavaTreeCell("Fn_CPD_EffectivityOperations", JavaWindow("Collaborative Product"), "EffectivityTree", aNode(Ubound(aNode)), "Out Date","LEFT")
									wait 1
									Set objShell = CreateObject("Wscript.Shell")
									For iKeyCnt=0 To i
										Call Fn_KeyBoardOperation("SendKeys", "{DOWN}")
									Next
									Call Fn_KeyBoardOperation("SendKeys", "{ENTER}")
									wait 1
									If Trim(cstr(aOutDate(i))) <> Trim(cstr(JavaWindow("Collaborative Product").JavaList("TableList").GetROProperty("value"))) Then
										bFlag = False
									End If
									Call Fn_KeyBoardOperation("SendKeys", "{ESC}")
									Set objShell=Nothing
									End If 
							 	End If
						Next
						If bFlag = False Then
							Fn_CPD_EffectivityOperations = False
							 Exit Function
						Else
							Fn_CPD_EffectivityOperations = bFlag
						End If
					'---------------------------------------------------------------------------
					Case "VerifyEffectivityForMultipleNode"							'[TC1123-20170410-14_04_2017-JotibaT-Maintenance] - Updated By Jotiba (Changed object JavaTable to JavaTree)
						bFlag = True
						For i = 0 to iLimit
							' Verify From Unit
							If sFromUnit <> "" Then
								If aFromUnit(i)<>"" Then
								Call Fn_ClickEffectivityTreeCell("Fn_CPD_EffectivityOperations", JavaWindow("Collaborative Product"), "EffectivityTree", aNode(Ubound(aNode)), "From Unit","LEFT")
								wait 1
								Set objShell = CreateObject("Wscript.Shell")
								For iKeyCnt=0 To i
									Call Fn_KeyBoardOperation("SendKeys", "{DOWN}")
								Next
								Call Fn_KeyBoardOperation("SendKeys", "{ENTER}")
								wait 1
								objShell.SendKeys "^A"
								wait 1
								If cstr(aFromUnit(i)) <> cstr(JavaWindow("Collaborative Product").JavaEdit("TableText").GetROProperty("text")) Then
									bFlag = False
								End If
								Call Fn_KeyBoardOperation("SendKeys", "{ESC}")
								Call Fn_KeyBoardOperation("SendKeys", "{RIGHT}")
								Set objShell=Nothing
								End If 
							 End If 
													 
							 ' Verify To Unit
							 If sToUnit<> "" Then
							 	If aToUnit(i)<>"" Then	
							 	Call Fn_ClickEffectivityTreeCell("Fn_CPD_EffectivityOperations", JavaWindow("Collaborative Product"), "EffectivityTree", aNode(Ubound(aNode)), "To Unit","LEFT")
								wait 1
								Set objShell = CreateObject("Wscript.Shell")
								For iKeyCnt=0 To i
									Call Fn_KeyBoardOperation("SendKeys", "{DOWN}")
								Next
								Call Fn_KeyBoardOperation("SendKeys", "{ENTER}")
								wait 1
								If cstr(aToUnit(i)) <> cstr(JavaWindow("Collaborative Product").JavaList("TableList").GetROProperty("value")) Then
									bFlag = False
								End If
								'Call Fn_KeyBoardOperation("SendKeys", "{TAB}")
								Call Fn_KeyBoardOperation("SendKeys", "{ESC}")
								Call Fn_KeyBoardOperation("SendKeys", "{RIGHT}")
								Set objShell=Nothing
								End If 
							 End If
							 ' Vrify In Date 
							 	If sInDate <> ""  Then
							 		If aInDate(i)<>"" Then
							 		Call Fn_ClickEffectivityTreeCell("Fn_CPD_EffectivityOperations", JavaWindow("Collaborative Product"), "EffectivityTree", aNode(Ubound(aNode)), "In Date","LEFT")
									wait 1
									Set objShell = CreateObject("Wscript.Shell")
									For iKeyCnt=0 To i
										Call Fn_KeyBoardOperation("SendKeys", "{DOWN}")
									Next
									Call Fn_KeyBoardOperation("SendKeys", "{ENTER}")
									wait 1
									If Trim(cstr(aInDate(i))) <> Trim(cstr(JavaWindow("Collaborative Product").JavaList("TableList").GetROProperty("value"))) Then
										bFlag = False
									End If
									Call Fn_KeyBoardOperation("SendKeys", "{ESC}")
									Call Fn_KeyBoardOperation("SendKeys", "{RIGHT}")
									Set objShell=Nothing
									End If 
							 	End If
							 	'Verify Out Date
							 	If sOutDate<>"" Then
							 		If Fn_CPD_EffectivityOperations("VerifyColumnExistInEffectivityTab", "", "From Unit", "", "", "", "")=True And sFromUnit="" Then
										Call Fn_KeyBoardOperation("SendKeys", "{RIGHT}")
										Call Fn_KeyBoardOperation("SendKeys", "{RIGHT}")
									End If
							 		If aOutDate(i)<>"" Then
							 		Call Fn_ClickEffectivityTreeCell("Fn_CPD_EffectivityOperations", JavaWindow("Collaborative Product"), "EffectivityTree", aNode(Ubound(aNode)), "Out Date","LEFT")
									wait 1
									Set objShell = CreateObject("Wscript.Shell")
									For iKeyCnt=0 To i
										Call Fn_KeyBoardOperation("SendKeys", "{DOWN}")
									Next
									Call Fn_KeyBoardOperation("SendKeys", "{ENTER}")
									wait 1
									If Trim(cstr(aOutDate(i))) <> Trim(cstr(JavaWindow("Collaborative Product").JavaList("TableList").GetROProperty("value"))) Then
										bFlag = False
									End If
									Set objShell=Nothing
									End If 
							 	End If
						Next
						If bFlag = False Then
							Fn_CPD_EffectivityOperations = False
							 Exit Function
						Else
							Fn_CPD_EffectivityOperations = bFlag
						End If
					'---------------------------------------------------------------------------	
					Case "VerifyBlankInEffectivityTab"
							bFlag = True
						For i = 0 to iLimit
							' Verify From Unit
							If aFromUnit(0)="From Unit" Then
								Call Fn_ClickEffectivityTreeCell("Fn_CPD_EffectivityOperations", JavaWindow("Collaborative Product"), "EffectivityTree", aNode(Ubound(aNode)), "From Unit","LEFT")
								wait 1
								Set objShell = CreateObject("Wscript.Shell")
								For iKeyCnt=0 To i
									Call Fn_KeyBoardOperation("SendKeys", "{DOWN}")
								Next
								Call Fn_KeyBoardOperation("SendKeys", "{ENTER}")
								wait 1
								objShell.SendKeys "^A"
								wait 1
								If cstr(JavaWindow("Collaborative Product").JavaEdit("TableText").GetROProperty("text"))<>"" Then
									bFlag = False
								End If
								Call Fn_KeyBoardOperation("SendKeys", "{ESC}")
								Call Fn_KeyBoardOperation("SendKeys", "{RIGHT}")
								Set objShell=Nothing
							 End If 
													 
							 ' Verify To Unit
							 If aToUnit(0)="To Unit" Then
							 	Call Fn_ClickEffectivityTreeCell("Fn_CPD_EffectivityOperations", JavaWindow("Collaborative Product"), "EffectivityTree", aNode(Ubound(aNode)), "To Unit","LEFT")
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
								If cstr(JavaWindow("Collaborative Product").WinEdit("WinEffTabText").GetROProperty("text"))<>"" Then
									bFlag = False
								End If
								Call Fn_KeyBoardOperation("SendKeys", "{ESC}")
								'Call Fn_KeyBoardOperation("SendKeys", "{TAB}")
								Call Fn_KeyBoardOperation("SendKeys", "{RIGHT}")
								Set objShell=Nothing
							 End If
							 ' Vrify In Date 
							 	If aInDate(0)="In Date" Then
							 		Call Fn_ClickEffectivityTreeCell("Fn_CPD_EffectivityOperations", JavaWindow("Collaborative Product"), "EffectivityTree", aNode(Ubound(aNode)), "In Date","LEFT")
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
									If cstr(JavaWindow("Collaborative Product").WinEdit("WinEffTabText").GetROProperty("text"))<>"" Then
										bFlag = False
									End If
									Call Fn_KeyBoardOperation("SendKeys", "{ESC}")
								'Call Fn_KeyBoardOperation("SendKeys", "{TAB}")
								Call Fn_KeyBoardOperation("SendKeys", "{RIGHT}")
									Set objShell=Nothing
							 	End If
							 	'Verify Out Date
							 	If aOutDate(0)="Out Date" Then
							 		Call Fn_ClickEffectivityTreeCell("Fn_CPD_EffectivityOperations", JavaWindow("Collaborative Product"), "EffectivityTree", aNode(Ubound(aNode)), "Out Date","LEFT")
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
									If cstr(JavaWindow("Collaborative Product").WinEdit("WinEffTabText").GetROProperty("text"))<>"" Then
										bFlag = False
									End If
									Call Fn_KeyBoardOperation("SendKeys", "{ESC}")
									Set objShell=Nothing
							 	End If
						Next
						If bFlag = False Then
							Fn_CPD_EffectivityOperations = False
							 Exit Function
						Else
							Fn_CPD_EffectivityOperations = bFlag
						End If
					' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
					Case "VerifyColumnExistInEffectivityTab"					'[TC1123-20170410-14_04_2017-JotibaT-Maintenance] - Updated By Jotiba (Changed object JavaTable to JavaTree)
							intNoOfObjects = cInt(JavaWindow("Collaborative Product").JavaTree("EffectivityTree").GetROProperty("columns_count"))
							For iCnt = 0 to intNoOfObjects - 1
								If JavaWindow("Collaborative Product").JavaTree("EffectivityTree").GetColumnHeader(iCnt) = sValue Then
									Fn_CPD_EffectivityOperations=True
									Exit For 
								End If
							Next
					' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
					Case "GetColumnCountInEffectivityTab"			'[TC1123-20170410-14_04_2017-JotibaT-Maintenance] - Updated By Jotiba (Changed object JavaTable to JavaTree)
						Fn_CPD_EffectivityOperations = -1
						If JavaWindow("Collaborative Product").JavaTree("EffectivityTree").Exist(2) Then
							Fn_CPD_EffectivityOperations = cInt(JavaWindow("Collaborative Product").JavaTree("EffectivityTree").GetROProperty("columns_count"))
						End If
					' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
					Case "VerifyListContents"				'[TC1123-20170410-14_04_2017-JotibaT-Maintenance] - Updated By Jotiba (Changed object JavaTable to JavaTree)
						Fn_CPD_EffectivityOperations = False
'						sFromUnit
						If sToUnit <> "" Then
'							objEffectivityTable.ActivateCell cInt(sValue),"To Unit"
							Call Fn_UI_ClickJavaTreeCell("Fn_CPD_EffectivityOperations", JavaWindow("Collaborative Product"), "EffectivityTree", aNode(Ubound(aNode)), "To Unit","LEFT")
							wait 1
							Set objShell = CreateObject("Wscript.Shell")
							Call Fn_KeyBoardOperation("SendKeys", "{DOWN}")
							wait 1
							Call Fn_KeyBoardOperation("SendKeys", "{ENTER}")
							wait 1
							iItemCount = cInt(objEff.javaList("TableList").GetROProperty("items count"))
							For iCnt = 0 to iItemCount - 1
								If sToUnit = objEff.javaList("TableList").Object.getItem(iCnt) Then
									Fn_CPD_EffectivityOperations = True
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully to Found [ " & sToUnit & " ] in [ To Unit ]")
									exit for
								End If
							Next
							Call Fn_KeyBoardOperation("SendKeys", "{TAB}")
						End If
						If Fn_CPD_EffectivityOperations = False Then 
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Find [ " & sToUnit & " ] in [ To Unit ]")
							Exit Function
						End If
'						sInDate,
						If sOutDate <> "" Then
							Fn_CPD_EffectivityOperations = False
'							objEffectivityTable.ActivateCell cInt(sValue),"Out Date"
							Call Fn_UI_ClickJavaTreeCell("Fn_CPD_EffectivityOperations", JavaWindow("Collaborative Product"), "EffectivityTree", aNode(Ubound(aNode)), "Out Date","LEFT")
							wait 1
							Set objShell = CreateObject("Wscript.Shell")
							Call Fn_KeyBoardOperation("SendKeys", "{DOWN}")
							wait 1
							Call Fn_KeyBoardOperation("SendKeys", "{ENTER}")
							wait 1
							iItemCount = cInt(objEff.javaList("TableList").GetROProperty("items count"))
							For iCnt = 0 to iItemCount - 1
								If sOutDate = objEff.javaList("TableList").Object.getItem(iCnt) Then
									Fn_CPD_EffectivityOperations = True
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully to Found [ " & sOutDate & " ] in [ Out Date ]")
									exit for
								End If
							Next
						End If
						If Fn_CPD_EffectivityOperations = False Then 
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Find [ " & sOutDate  & " ] in [ Out Date ]")
							Exit Function
						End If
					' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
					Case "ActivateAndTypeInEffectivityTab"				'[TC1123-20170410-14_04_2017-JotibaT-Maintenance] - Updated By Jotiba (Changed object JavaTable to JavaTree)
						If sFromUnit <> ""  Then 
							Call Fn_UI_ClickJavaTreeCell("Fn_CPD_EffectivityOperations", JavaWindow("Collaborative Product"), "EffectivityTree", aNode(Ubound(aNode)), "From Unit","LEFT")
							wait 1
							Set objShell = CreateObject("Wscript.Shell")
							Call Fn_KeyBoardOperation("SendKeys", "{DOWN}")
							wait 1
							Call Fn_KeyBoardOperation("SendKeys", "{ENTER}")
							objShell.SendKeys "^A"
							wait 1
							objShell.SendKeys sFromUnit
							wait 1
							Call Fn_KeyBoardOperation("SendKeys", "{ENTER}")
							Set objShell=Nothing
						End If 
					
						If sToUnit<>"" Then	
								If Fn_CPD_EffectivityOperations("VerifyColumnExistInEffectivityTab", "", "From Unit", "", "", "", "")=True Then
									Call Fn_ClickEffectivityTreeCell("Fn_CPD_EffectivityOperations", JavaWindow("Collaborative Product"), "EffectivityTree", aNode(Ubound(aNode)), "From Unit","LEFT")
									Set objShell = CreateObject("Wscript.Shell")
										Call Fn_KeyBoardOperation("SendKeys", "{RIGHT}")
									Set objShell=Nothing
								End If
								Call Fn_UI_ClickJavaTreeCell("Fn_CPD_EffectivityOperations", JavaWindow("Collaborative Product"), "EffectivityTree", aNode(Ubound(aNode)), "To Unit","LEFT")
								wait 1
								Set objShell = CreateObject("Wscript.Shell")
								Call Fn_KeyBoardOperation("SendKeys", "{DOWN}")
								wait 1
								Call Fn_KeyBoardOperation("SendKeys", "{ENTER}")
								wait 1
								objShell.SendKeys "^A"
								wait 1
								If trim(aToUnit(i)) <> "" Then
									objEff.JavaList("TableList").Type sToUnit
								End If
								wait 1
								Call Fn_KeyBoardOperation("SendKeys", "{ENTER}")
								Call Fn_KeyBoardOperation("SendKeys", "{RIGHT}")
								Set objShell=Nothing
							End If
							
							If sInDate <> "" Then
								Call Fn_SyncTCObjects()
								Call Fn_UI_ClickJavaTreeCell("Fn_CPD_EffectivityOperations", JavaWindow("Collaborative Product"), "EffectivityTree", aNode(Ubound(aNode)), "In Date","LEFT")
								wait 1
								Set objShell = CreateObject("Wscript.Shell")
								wait 1
								Call Fn_KeyBoardOperation("SendKeys", "{DOWN}")
								wait 1
								Call Fn_KeyBoardOperation("SendKeys", "{ENTER}")
								
								Call Fn_List_Select("",objEff,"TableList","Select Date...")
								If instr(sInDate,"$") > 0 Then
									arrDate = split(trim(sInDate,"$"))
									Call  Fn_CPD_DateControl("Set", arrDate(0), arrDate(1))
								Else
									Select Case lcase(trim(aInDate(i)))
										Case ""
											Call  Fn_CPD_DateControl("Clear", "", "")
										Case "today"
											Call  Fn_CPD_DateControl("Today", "", "")
										Case Else
											Call  Fn_CPD_DateControl("Set", sInDate, "")
									End Select
								End If
							End If 
					
							If sOutDate<>"" Then
								If lcase(aOutDate(i)) <> "so" AND lcase(aOutDate(i)) <> "up" Then
										Call Fn_SyncTCObjects()
										Call Fn_UI_ClickJavaTreeCell("Fn_CPD_EffectivityOperations", JavaWindow("Collaborative Product"), "EffectivityTree", aNode(Ubound(aNode)), "Out Date","LEFT")
										wait 1
										Set objShell = CreateObject("Wscript.Shell")
										wait 1
										Call Fn_KeyBoardOperation("SendKeys", "{DOWN}")
										wait 1
										Call Fn_KeyBoardOperation("SendKeys", "{ENTER}")
										Call Fn_List_Select("Fn_CPD_EffectivityOperations",objEff,"TableList","Select Date...")
										If instr(sOutDate,"$") > 0 Then
												arrDate = split(trim(sOutDate,"$"))
												Call  Fn_CPD_DateControl("Set", arrDate(0), arrDate(1))
										Else
												Select Case lcase(trim(sOutDate))
													Case ""
														Call  Fn_CPD_DateControl("Clear", "", "")
													Case "today"
														Call  Fn_CPD_DateControl("Today", "", "")
													Case Else
														Call  Fn_CPD_DateControl("Set",sOutDate, "")
												End Select
										End If
								Else
										Call Fn_List_Select("Fn_CPD_EffectivityOperations",objEff,"TableList",sOutDate)
								End If
							End If
					
							Fn_CPD_EffectivityOperations = True
							
					Case "ClearcellData"
							If sFromUnit= "From Unit"  Then 
								Call Fn_UI_ClickJavaTreeCell("Fn_CPD_EffectivityOperations", JavaWindow("Collaborative Product"), "EffectivityTree", aNode(Ubound(aNode)), "From Unit","LEFT")
								wait 1
								Set objShell = CreateObject("Wscript.Shell")
								Call Fn_KeyBoardOperation("SendKeys", "{DOWN}")
								wait 1
								Call Fn_KeyBoardOperation("SendKeys", "{ENTER}")
								objShell.SendKeys "^A"
								wait 1
								For iKeyCnt = 0 to 10
									Call Fn_KeyBoardOperation("SendKeys", "{BKSP}")
								Next
								wait 1
								Call Fn_KeyBoardOperation("SendKeys", "{ENTER}")
								Set objShell=Nothing
							End If 
							If sToUnit="To Unit" Then	
								Call Fn_UI_ClickJavaTreeCell("Fn_CPD_EffectivityOperations", JavaWindow("Collaborative Product"), "EffectivityTree", aNode(Ubound(aNode)), "To Unit","LEFT")
								wait 1
								Set objShell = CreateObject("Wscript.Shell")
								Call Fn_KeyBoardOperation("SendKeys", "{DOWN}")
								wait 1
								Call Fn_KeyBoardOperation("SendKeys", "{ENTER}")
								wait 1
								objShell.SendKeys "^A"
								wait 1
								If trim(aToUnit(i)) <> "" Then
									objEff.JavaList("TableList").Type ""
								End If
								For iKeyCnt = 0 to 10
									Call Fn_KeyBoardOperation("SendKeys", "{BKSP}")
								Next
								wait 1
								Call Fn_KeyBoardOperation("SendKeys", "{ENTER}")
								Call Fn_KeyBoardOperation("SendKeys", "{RIGHT}")
								Set objShell=Nothing
							End If
							Fn_CPD_EffectivityOperations = True
							
					Case "DeleteInEffectivityTab","DeleteInEffectivityTab_Ext"					'[TC1123-20170410-14_04_2017-JotibaT-Maintenance] - Updated By Jotiba (Changed object JavaTable to JavaTree)
'						bFlag =Fn_CPD_EffectivityOperations("VerifyInEffectivityTab",sNode,sValue,sFromUnit,sToUnit,sInDate,sOutDate)
'						If bFlag = True then
							bFlag = True 
							For i = 0 to iLimit
								if (sFromUnit <> "" and sToUnit <> "" and sInDate <> "" and sOutDate <> "" AND uBound(aToUnit) <= iLimit) then
									Call Fn_UI_ClickJavaTreeCell("Fn_CPD_EffectivityOperations", JavaWindow("Collaborative Product"), "EffectivityTree", aNode(Ubound(aNode)), "From Unit","LEFT")
									wait 1
									Set objShell = CreateObject("Wscript.Shell")
									For iKeyCnt=0 To i
										Call Fn_KeyBoardOperation("SendKeys", "{DOWN}")
									Next
									Call Fn_KeyBoardOperation("SendKeys", "{ENTER}")
									wait 1
									Set objShell = Nothing
								ElseIf sFromUnit <> "" AND aFromUnit(i) <> "" AND uBound(aToUnit) <= iLimit Then
									Call Fn_UI_ClickJavaTreeCell("Fn_CPD_EffectivityOperations", JavaWindow("Collaborative Product"), "EffectivityTree", aNode(Ubound(aNode)), "From Unit","LEFT")
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
								Call Fn_Button_Click("Fn_CPD_EffectivityOperations", JavaWindow("Collaborative Product"),"DeleteButton")
								wait 1
								If sAction="DeleteInEffectivityTab" Then
									Call Fn_ToolbarOperation("Click", "Save (Ctrl+S)","")
								End If
								Fn_CPD_EffectivityOperations = True
							Next
'						End if	
						Fn_CPD_EffectivityOperations = bFlag
						
				Case "CheckEffectivityEmpty"
						If JavaWindow("Collaborative Product").JavaTree("EffectivityTree").Exist(2) Then
							JavaWindow("Collaborative Product").JavaTree("EffectivityTree").Object.selectAll()
							Fn_CPD_EffectivityOperations=JavaWindow("Collaborative Product").JavaTree("EffectivityTree").Object.getSelectionCount 
						Else 
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_CPD_EffectivityOperations : Invalid action [ " +sAction+ " ] is requested.")
						End If 
						
					'------------------------------------------------------
				End Select
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_CPD_EffectivityOperations : Invalid action [ " +sAction+ " ] is requested.")
	End Select

	If  tflag=1 Then
		Call Fn_CPD_CompnentTabOperations("DoubleClick",tab, "") 
	End If
	IF Fn_CPD_EffectivityOperations <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_CPD_EffectivityOperations : Executed successfully with case [" + sAction + "].")
	End If
	Set objEffectivityTable = Nothing
	Set intNoOfObjects = Nothing
	Set objEff = Nothing
End Function
''*********************************************************		Function to action perform on NavTree	***********************************************************************
'Function Name			:				Fn_CPD_NavTree_NodeOperation

'Description			:		 		 Actions performed in this function are:
'																	1. Node Select
'																	2. Node multi-select
'																	3. Node Expand
'																	4. Node Collapse
'																	5. Node Popup menu select
'																	6. Node double-click
'																	7. Node MultiSelect Cntxt Menu
'																	8. Node Exist
'																	9. MultiSelectContextMenuExist

'Parameters			    :	 	1. StrAction: Action to be performed
'								2. StrNodeName: Fully qulified tree Path (delimiter as ':') [multiple node are separated by "," ] 
'								3. StrMenu: Context menu to be selected

'Return Value		    : 		TRUE \ FALSE

'Pre-requisite			:		Collaborative Product Development module window should be displayed

'Examples				:		Fn_CPD_NavTree_NodeOperation("PopupMenuSelect","Home:Newstuff","Copy Ctrl+C")
'								EXAMPLE for Case "Select" : Call Fn_CPD_NavTree_NodeOperation( "Select" ,  "Home:Newstuff:000032-CarModel_VI_LS1:000032 @2" , "" ) 
'															Call Fn_CPD_NavTree_NodeOperation( "Select" ,  "Home:Newstuff:000032-CarModel_VI_LS1:000032" , "" ) 
'								EXAMPLE for Case "GetSelected"::  Fn_CPD_NavTree_NodeOperation( "GetSelected" , "Home:Mailbox,Home:Newstuff,Home:000039-1,Home:Kavan_Shah" , "" ) 
'								EXAMPLE for Case "GetChildItemCount"::  Fn_CPD_NavTree_NodeOperation( "GetChildItemCount" , "Home:Mailbox" , "" ) 
'								EXAMPLE for Case "GetChildInstances"::  Fn_CPD_NavTree_NodeOperation("GetChildInstances","Home:AutomatedTests:sonal:000112-top:000112/A;1-top:View:000114-sub2","") 		Added By Ketan On 11-Jan-2011
'History				:		
'	Developer Name		Date		Rev. No.		Changes Done								Reviewer
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Koustubh W		20/01/2011		1.0				Created
'	Reema W			07/12/2015		1.1				Added Case "MultiSelectContextMenuEnabled"	[TC1121-20151116a-07_12_2015-VivekA-NewDevelopment]
'	Chaitali R		11/12/2015		1.1				Added Case "VerifySelectedNode"				[TC1121-20151116b-11_12_2015-VivekA-NewDevelopment]
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_CPD_NavTree_NodeOperation(StrAction,StrNodeName,StrMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_CPD_NavTree_NodeOperation"
	Dim NodeLists, intNodeCount, intCount, StrExist, aMenuList, sTreeItem,sCmpItm
	Dim objJavaWindowCPD, objJavaTreeNav,ArrNodeName
	Dim ArrStrcomp, sArrStr1,sArrStr2, iCounter
	Dim iRows, colonCnt
	Dim iItemCount, aNodePath, iInstance, instCount, aNodes
	Dim sPath, sEle ,iCnt, bFound
	Dim iLen,iIndex,iTotal,iCount,sReturn,iReturn,arr
	Dim iPath,iVal,iPath1,arrNode
	Dim arrStrNode,echStrNode,oCurrentNode
	Dim sParentPath,sVerifyNode

         'Variable Declaration
	Dim sItemPath,aStrNode, i
	Dim iInstanceCnt, iOccCnt

	Set objJavaWindowCPD = JavaWindow("Collaborative Product")
	
	aMenuList = split(StrMenu, ":",-1,1)
	intCount = Ubound(aMenuList)

	'Swapnil:Made changes for CPD
	If intCount  > "0" Then
		If aMenuList (1) = "Collaborative Product Development"  Then
			StrMenu = Replace(StrMenu,trim(aMenuList(1)),"4G Designer",1,1,1)
			aMenuList = split(StrMenu, ":",-1,1)
		End If
	End If

	Select Case StrAction
		'POC shweta rathod **********************************************************************************************
		'To Expand nodes upto Last child's parent node in StrNodeName hierarchy and then select the last node
		Case "ExpandAndSelect"
				'Initial Item Path
				arrStrNode = Split (StrNodeName, ":")
				For i = 0 to UBound(arrStrNode)-1
					If sParentPath = "" Then
						sParentPath  = arrStrNode(i)
					Else
						sParentPath  = sParentPath + ":" + arrStrNode(i)
					End If
					'Call Fn_MyTc_NavTree_NodeOperation("Expand", sParentPath, "")
					iPath = Fn_UI_JavaTreeGetItemPathExt("Fn_CPD_NavTree_NodeOperation", objJavaWindowCPD.JavaTree("NavTree"), StrNodeName , ":", "@")
					objJavaWindowCPD.JavaTree("NavTree").Expand iPath
					If arrStrNode(i) = "AutomatedTests" Then
						Call Fn_ReadyStatusSync(SISW_MICROLESS_TIMEOUT)
					End If
					Call Fn_ReadyStatusSync(SISW_MICRO_TIMEOUT)
				Next

				''iPath = Fn_UI_JavaTreeGetItemPathExt("Fn_MyTc_NavTree_NodeOperation", JavaWindow("MyTeamcenter").JavaTree("NavTree"), StrNodeName , sDelimiter, "@")
				iPath = Fn_UI_JavaTreeGetItemPathExt("Fn_CPD_NavTree_NodeOperation", objJavaWindowCPD.JavaTree("NavTree"), StrNodeName , ":", "@")
				If iPath=False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select Node [" + StrNodeName + "] of NavTree")
					Fn_CPD_NavTree_NodeOperation = False
				Else
					objJavaWindowCPD.JavaTree("NavTree").Select iPath
					Call Fn_ReadyStatusSync(SISW_MICRO_TIMEOUT)
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Selected Node [" + StrNodeName + "] of NavTree")
					Fn_CPD_NavTree_NodeOperation = True
				End If
		'To expand all the nodes present in StrNodeName
		Case "ExpandAll"
				'Initial Item Path
				arrStrNode = Split (StrNodeName,":")
				For i = 0 to UBound(arrStrNode)-1
					If sParentPath = "" Then
						sParentPath  = arrStrNode(i)
					Else
						sParentPath  = sParentPath + ":" + arrStrNode(i)
					End If
					iPath = Fn_UI_JavaTreeGetItemPathExt("Fn_CPD_NavTree_NodeOperation", objJavaWindowCPD.JavaTree("NavTree"), StrNodeName , ":", "@")
					objJavaWindowCPD.JavaTree("NavTree").Expand iPath
					If arrStrNode(i) = "AutomatedTests" Then
						Call Fn_ReadyStatusSync(SISW_MICROLESS_TIMEOUT)
					End If
					Call Fn_ReadyStatusSync(SISW_MICRO_TIMEOUT)
				Next

				iPath = Fn_UI_JavaTreeGetItemPathExt("Fn_CPD_NavTree_NodeOperation", objJavaWindowCPD.JavaTree("NavTree"), StrNodeName , ":", "@")
				If iPath=False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select Node [" + StrNodeName + "] of NavTree")
					Fn_CPD_NavTree_NodeOperation = False
				Else
					objJavaWindowCPD.JavaTree("NavTree").Expand iPath
					Call Fn_ReadyStatusSync(SISW_MICRO_TIMEOUT)
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Selected Node [" + StrNodeName + "] of NavTree")
					Fn_CPD_NavTree_NodeOperation = True
				End If
		'POC shweta Rathod **********************************************************************************************

		'----------------------------------------------------------------------- For selecting single node -------------------------------------------------------------------------
		Case "Select"
				iPath = Fn_UI_JavaTreeGetItemPathExt("Fn_CPD_NavTree_NodeOperation", objJavaWindowCPD.JavaTree("NavTree"), StrNodeName , ":", "@")
				If iPath = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select Node [" + StrNodeName + "] of NavTree")
					Fn_CPD_NavTree_NodeOperation = False
				Else
					objJavaWindowCPD.JavaTree("NavTree").Select iPath
					Call Fn_ReadyStatusSync(1)
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Selected Node [" + StrNodeName + "] of NavTree")
					Fn_CPD_NavTree_NodeOperation = True
				End If
		
		' - - - - - - - - - - Deselect Node
		Case "Deselect"	
				iPath = Fn_UI_JavaTreeGetItemPathExt("Fn_CPD_NavTree_NodeOperation", objJavaWindowCPD.JavaTree("NavTree"), StrNodeName , ":", "@")
				If iPath = False Then
					 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to DeSelect Node [" + StrNodeName + "] of NavTree")
					  Fn_CPD_NavTree_NodeOperation = False
				Else
					objJavaWindowCPD.JavaTree("NavTree").Deselect iPath
					Call Fn_ReadyStatusSync(1)
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully DeSelected Node [" + StrNodeName + "] of NavTree")
					Fn_CPD_NavTree_NodeOperation = True
				End If
		' - - - - - - - - - - Multi Select Nodes
		Case "Multiselect"
				arrNode = Split(StrNodeName,",")
				For iCount = 0 To UBound(arrNode)
					iPath = Fn_UI_JavaTreeGetItemPathExt("Fn_CPD_NavTree_NodeOperation", objJavaWindowCPD.JavaTree("NavTree"), arrNode(iCount) , ":", "@")
					If iPath = False Then
						 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Multi Select Node [" + StrNodeName + "] of NavTree")
						  Fn_CPD_NavTree_NodeOperation = False
						  Exit Function
					Else
						objJavaWindowCPD.JavaTree("NavTree").ExtendSelect iPath
						Call Fn_ReadyStatusSync(1)
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Multi Selected Node [" + StrNodeName + "] of NavTree")
						Fn_CPD_NavTree_NodeOperation = True
					End If
				Next
		' - - - - - - - - - - Expand Node
		Case "Expand"
			iPath = Fn_UI_JavaTreeGetItemPathExt("Fn_CPD_NavTree_NodeOperation", objJavaWindowCPD.JavaTree("NavTree"), StrNodeName , ":", "@")
			If iPath=False Then
				 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Expand Node [" + StrNodeName + "] of NavTree")
				  Fn_CPD_NavTree_NodeOperation = False
			Else
				objJavaWindowCPD.JavaTree("NavTree").Expand iPath
				Call Fn_ReadyStatusSync(1)
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Expanded Node [" + StrNodeName + "] of NavTree")
				Fn_CPD_NavTree_NodeOperation = True
			End If

		' - - - - - - - - - - Collaplse Node
		Case "Collapse"
			iPath = Fn_UI_JavaTreeGetItemPathExt("Fn_CPD_NavTree_NodeOperation", objJavaWindowCPD.JavaTree("NavTree"), StrNodeName , ":", "@")
			If iPath = False Then
				 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Collapse Node [" + StrNodeName + "] of NavTree")
				  Fn_CPD_NavTree_NodeOperation = False
			Else
				objJavaWindowCPD.JavaTree("NavTree").Collapse iPath
				Call Fn_ReadyStatusSync(1)
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Collapse Node [" + StrNodeName + "] of NavTree")
				Fn_CPD_NavTree_NodeOperation = True
			End If
		' - - - - - - - - - - Pop Up Menu Select
		Case "PopupMenuSelect"
			Set objJavaTreeNav = objJavaWindowCPD.JavaTree("NavTree")
					'Build the Popup menu to be selected
					aMenuList = split(StrMenu, ":",-1,1)
					intCount = Ubound(aMenuList)

					'Select node
                    iPath = Fn_UI_JavaTreeGetItemPathExt("Fn_CPD_NavTree_NodeOperation", objJavaWindowCPD.JavaTree("NavTree"), StrNodeName , ":", "@")
					If iPath=False Then
						 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Select Node [" + StrNodeName + "] of NavTree")
						  Fn_CPD_NavTree_NodeOperation = False
						  Exit Function
					Else
						objJavaWindowCPD.JavaTree("NavTree").Select iPath
						Call Fn_ReadyStatusSync(1)
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Selected Node [" + StrNodeName + "] of NavTree")
						Fn_CPD_NavTree_NodeOperation = True
					End If

					'Open context menu
					Call Fn_UI_JavaTree_OpenContextMenu("Fn_CPD_NavTree_NodeOperation",objJavaWindowCPD,"NavTree",iPath)
                    
					'Select Menu action
					Select Case intCount
						Case "0"
							 StrMenu = objJavaWindowCPD.WinMenu("ContextMenu").BuildMenuPath(aMenuList(0))
						Case "1"
							StrMenu = objJavaWindowCPD.WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1))
						Case "2"
							StrMenu = objJavaWindowCPD.WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1),aMenuList(2))
						Case Else
							Fn_CPD_NavTree_NodeOperation = FALSE
							Exit Function
					End Select

					objJavaWindowCPD.WinMenu("ContextMenu").Select StrMenu
					If Err.number < 0 Then
						Fn_CPD_NavTree_NodeOperation = False
					Else
						Fn_CPD_NavTree_NodeOperation = True
					End If
        ' - - - - - - - - - - PopUp Menu Existance on multi Select
		Case "MultiSelectContextMenuExist"
				NodeLists = split(StrNodeName,",",-1,1)
				Set objJavaTreeNav = objJavaWindowCPD.JavaTree("NavTree")
				Call Fn_CPD_NavTree_NodeOperation("Multiselect",StrNodeName,"")
				iPath=Fn_UI_JavaTreeGetItemPath(objJavaWindowCPD.JavaTree("NavTree"),NodeLists(0))
				iPath = Fn_UI_JavaTreeGetItemPathExt("Fn_CPD_NavTree_NodeOperation", objJavaWindowCPD.JavaTree("NavTree"), NodeLists(0) , ":", "@")
				Call Fn_UI_JavaTree_OpenContextMenu("Fn_CPD_NavTree_NodeOperation",objJavaWindowCPD,"NavTree",iPath)
				If objJavaWindowCPD.WinMenu("ContextMenu").GetItemProperty (StrMenu,"Exists") = True Then
					Fn_CPD_NavTree_NodeOperation = True
				Else
					Fn_CPD_NavTree_NodeOperation = False
			  	End If
		'Double Click on Node
		Case "DoubleClick"
			iPath = Fn_UI_JavaTreeGetItemPathExt("Fn_CPD_NavTree_NodeOperation", objJavaWindowCPD.JavaTree("NavTree"), StrNodeName , ":", "@")
			If iPath <> False Then
				objJavaWindowCPD.JavaTree("NavTree").Select iPath 
				Call Fn_KeyBoardOperation("SendKeys", "{ENTER}")
				Call Fn_ReadyStatusSync(1)
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully DoubleClick Node [" + StrNodeName + "] of NavTree")
				Fn_CPD_NavTree_NodeOperation = True
			End If
			If Fn_CPD_CompnentTabOperations("Exists", "Content Search","") = True Then ' [TC11.4-2017071700-4_8_2017-JotibaT-Maintenace]- Added code as per design change.
				Call Fn_CPD_CompnentTabOperations("Close","Content Search","")
			End If
			
'				Dim intX, intY, intWidth, intHeight, strComputer, sOSName, objWMIService, oss, os
'				intX = 0
'				intY = 0
'				intWidth = 0
'				intHeight = 0
'				strComputer = "." 
'				Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
'				Set oss = objWMIService.ExecQuery ("Select * from Win32_OperatingSystem") 
'				For Each os in oss 
'					sOSName = os.Caption
'				Next
'	
'				If Instr(ucase(sOSName),"XP") > 0 then
'					wait 1
'					objJavaWindowCPD.JavaTree("NavTree").Activate iPath
'					Set objWMIService = Nothing
'					Set oss = Nothing
'					Fn_CPD_NavTree_NodeOperation = True
'				Else
'					intX = objNodeBounds.x
'					intY = objNodeBounds.y
'					intWidth = objNodeBounds.width
'					intHeight = objNodeBounds.height
'					Set objNodeBounds = nothing
'					wait 1
'					objJavaWindowCPD.JavaTree("NavTree").DblClick cInt(intX + intWidth/2), cInt(intY + intHeight/2), "LEFT"
'					If Err.number < 0 Then
'						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to DoubleClick Node [" + StrNodeName + "] of NavTree")
'						Fn_CPD_NavTree_NodeOperation = False
'						Exit Function
'					End If
'				End If
		' - - - - - - - - - - Popup Menu operation on Multi Selected Nodes
		Case "MultiSelectContextMenu"
					NodeLists = split(StrNodeName,",",-1,1)
					aMenuList = split(StrMenu, ":",-1,1)
					intCount = Ubound(aMenuList)
					Set objJavaTreeNav = objJavaWindowCPD.JavaTree("NavTree")

					'Select multiple node
					Call Fn_CPD_NavTree_NodeOperation("Multiselect", StrNodeName, "")
					'Open context menu
					iPath = Fn_UI_JavaTreeGetItemPathExt("Fn_CPD_NavTree_NodeOperation", objJavaWindowCPD.JavaTree("NavTree"), NodeLists(0) , ":", "@")
					Call Fn_UI_JavaTree_OpenContextMenu("Fn_CPD_NavTree_NodeOperation",objJavaWindowCPD,"NavTree",iPath)
					Select Case intCount
						Case "0"
							 StrMenu = objJavaWindowCPD.WinMenu("ContextMenu").BuildMenuPath(aMenuList(0))
						Case "1"
							StrMenu = objJavaWindowCPD.WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1))
						Case "2"
							StrMenu = objJavaWindowCPD.WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1),aMenuList(2))
						Case Else
							Fn_CPD_NavTree_NodeOperation = False
							Exit Function
					End Select
					objJavaWindowCPD.WinMenu("ContextMenu").Select StrMenu
					If Err.number < 0 Then
						Fn_CPD_NavTree_NodeOperation = False
					else
						Fn_CPD_NavTree_NodeOperation = True
					End If	
		' - - - - - - - - - - Existance of Node
		Case "Exist"
				Set objJavaTreeNav = objJavaWindowCPD.JavaTree("NavTree")
				iPath = Fn_UI_JavaTreeGetItemPathExt("Fn_CPD_NavTree_NodeOperation", objJavaWindowCPD.JavaTree("NavTree"), StrNodeName , ":", "@")
				If iPath = False Then
				'iPath = Fn_UI_getJavaTreeIndex(objJavaWindowCPD.JavaTree("NavTree"),StrNodeName)
				'If iPath = -1 Then
					 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Node [" + StrNodeName + "] Not exist in NavTree")
					  Fn_CPD_NavTree_NodeOperation = False
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Node [" + StrNodeName + "] Exist in NavTree")
					aNodePath = split(replace(iPath,"#",""),":")
					Fn_CPD_NavTree_NodeOperation = True
					Set oCurrentNode = objJavaWindowCPD.JavaTree("NavTree").Object
					For iCnt = 0 to UBound(aNodePath) -1
						Set oCurrentNode = oCurrentNode.GetItem(aNodePath(iCnt))
						If cBool(oCurrentNode.getExpanded()) = False Then
							Fn_CPD_NavTree_NodeOperation = false
							Exit for
						End If
					Next
					If Fn_CPD_NavTree_NodeOperation Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Node [" + StrNodeName + "] Exist in NavTree")
					End If
				End If
		' - - - - - - - - - - Existance of Popup Menu
		Case "PopupMenuExist"
					aMenuList = split(StrMenu, ":",-1,1)
					intCount = Ubound(aMenuList)
					Set objJavaTreeNav = objJavaWindowCPD.JavaTree("NavTree")
					'Open context menu
					iPath = Fn_UI_JavaTreeGetItemPathExt("Fn_CPD_NavTree_NodeOperation", objJavaWindowCPD.JavaTree("NavTree"), StrNodeName , ":", "@")
					Call Fn_UI_JavaTree_OpenContextMenu("Fn_CPD_NavTree_NodeOperation",objJavaWindowCPD,"NavTree",iPath)
					Select Case intCount
						Case "0"
							 StrMenu = objJavaWindowCPD.WinMenu("ContextMenu").BuildMenuPath(aMenuList(0))
                        Case "1"
							StrMenu = objJavaWindowCPD.WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1))
                        Case "2"
							StrMenu = objJavaWindowCPD.WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1),aMenuList(2))
                        Case Else
							Fn_CPD_NavTree_NodeOperation = False
                        Exit Function
					End Select
					If objJavaWindowCPD.WinMenu("ContextMenu").GetItemProperty (StrMenu,"Exists") = True Then
						Fn_CPD_NavTree_NodeOperation = True
					Else
						Fn_CPD_NavTree_NodeOperation = False
					End If
					Call Fn_KeyBoardOperation("SendKeys", "{ESC}")
		' - - - - - - - - - - Checking State of Popup Menu		
		Case "PopupMenuEnabled"
					aMenuList = split(StrMenu, ":",-1,1)
					intCount = Ubound(aMenuList)
					Set objJavaTreeNav = objJavaWindowCPD.JavaTree("NavTree")
					'Open context menu
					iPath = Fn_UI_JavaTreeGetItemPathExt("Fn_CPD_NavTree_NodeOperation", objJavaWindowCPD.JavaTree("NavTree"), StrNodeName , ":", "@")
					Call Fn_UI_JavaTree_OpenContextMenu("Fn_CPD_NavTree_NodeOperation",objJavaWindowCPD,"NavTree",iPath)
					Select Case intCount
						Case "0"
							 StrMenu = objJavaWindowCPD.WinMenu("ContextMenu").BuildMenuPath(aMenuList(0))
                        Case "1"
							StrMenu = objJavaWindowCPD.WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1))
                        Case "2"
							StrMenu = objJavaWindowCPD.WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1),aMenuList(2))
                        Case Else
							Fn_CPD_NavTree_NodeOperation = FALSE
                        Exit Function
					End Select
				If objJavaWindowCPD.WinMenu("ContextMenu").GetItemProperty (StrMenu,"Enabled") = True Then
					Fn_CPD_NavTree_NodeOperation = True
				Else
					Fn_CPD_NavTree_NodeOperation = False
			  	End If
		'------------------- Checks That item is inactively focused Or Not for single node OR Multiple Node(comma "," SEPERATED)---------------
		Case "GetSelected"
		wait 5
			Set objJavaTreeNav = objJavaWindowCPD.JavaTree("NavTree")
				
				ArrStrcomp = Split(objJavaTreeNav.GetROProperty("value") ,"",-1,1)
				sArrStr2 = ArrStrcomp(0)
				For iCounter = 1 To ubound(ArrStrcomp)
					sArrStr2 = sArrStr2 & "," & ArrStrcomp(iCounter)
				Next
				If sArrStr2 = StrNodeName Then
				   'Writing Log
				   Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Java Tree Multiple Node ["+StrNodeName+"] is Selected .")
				   Fn_CPD_NavTree_NodeOperation = TRUE
				Else
				   'Writing Log
				   Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Java Tree Multiple  Node ["+StrNodeName+"] is Not Selected .")
				   Fn_CPD_NavTree_NodeOperation = FALSE
			End If
		' - - - - - - - - - - Getting Index of Node
		Case "GetIndex"
			'Index of Item1
			arrStrNode=Split(StrNodeName,":")
			If UBound(arrStrNode)=0 And Lcase(arrStrNode(0))="home" Then
				Fn_CPD_NavTree_NodeOperation=0
			Else
				iPath = Fn_UI_JavaTreeGetItemPathExt("Fn_CPD_NavTree_NodeOperation", objJavaWindowCPD.JavaTree("NavTree"), StrNodeName , ":", "@")
				sCmpItm=Replace(iPath,"#","")
				sCmpItm=Replace(sCmpItm,"0","1")
				arriPath=Split(sCmpItm,":")
				iVal=0
				For iCounter=0 To UBound(arriPath)
					iVal=iVal+CInt(arriPath(iCounter))
				Next
				If iPath=False Then
					Fn_CPD_NavTree_NodeOperation=False
				Else
					Fn_CPD_NavTree_NodeOperation=iVal
				End If
			End If
		' - - - - - - - - - - Getting Child Item Count
		Case "GetChildItemCount"
				iPath = Fn_UI_JavaTreeGetItemPathExt("Fn_CPD_NavTree_NodeOperation", objJavaWindowCPD.JavaTree("NavTree"), StrNodeName , ":", "@")
				Fn_CPD_NavTree_NodeOperation = -1
				If iPath <> False Then
					objJavaWindowCPD.JavaTree("NavTree").Expand iPath 
					arrStrNode = Split (replace(iPath,"#",""), ":")
					Set  oCurrentNode = objJavaWindowCPD.JavaTree("NavTree").Object.getItem(cInt(arrStrNode(0)))
					For iCount = 1 to UBound(arrStrNode)
						Set  oCurrentNode = oCurrentNode.getItem(cInt(arrStrNode(iCount))) 
					Next
					Fn_CPD_NavTree_NodeOperation = cInt(oCurrentNode.getItemCount())
				End If
				Set oCurrentNode=Nothing
		' - - - - - - - - - - Selecting Range of Nodes
		Case "SelectRange"
					ReDim ArrNodeName(2)
					ArrNodeName = Split(StrNodeName,"|")
				
					iPath = Fn_UI_JavaTreeGetItemPathExt("Fn_CPD_NavTree_NodeOperation", objJavaWindowCPD.JavaTree("NavTree"), ArrNodeName(0) , ":", "@")
					iPath1 = Fn_UI_JavaTreeGetItemPathExt("Fn_CPD_NavTree_NodeOperation", objJavaWindowCPD.JavaTree("NavTree"), ArrNodeName(1) , ":", "@")
					If iPath=False Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Select Node [" + StrNodeName + "] of NavTree")
						Fn_CPD_NavTree_NodeOperation = False
					Else
						objJavaWindowCPD.JavaTree("NavTree").SelectRange iPath,iPath1
						Call Fn_ReadyStatusSync(1)
						Fn_CPD_NavTree_NodeOperation = True
					End If

		'- - - - - - - - - - - -  To Get child Instances
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
				iCnt = Fn_CPD_NavTree_NodeOperation( "GetIndex" , sPath , "")
				iItemCount = Fn_CPD_NavTree_NodeOperation( "GetChildrenList" , sPath , "" )
				For iCounter=0 To UBound(iItemCount)
					If Trim(iItemCount(iCounter))=Trim(aNodePath( UBound(aNodePath))) Then
						iInstance = iInstance+1
					End If
				Next
				Fn_CPD_NavTree_NodeOperation = iInstance
				'- - - - - - - - - - - -  Retruns All Childs of any given Node in the tree in form of an array - - - - - - - - - - - - - - -
				Case "GetChildrenList"
						sReturn=""
						If Fn_CPD_NavTree_NodeOperation("Expand",StrNodeName,"")=True Then
							arrStrNode = Split (StrNodeName, ":")
							If UBound(arrStrNode)=0 And  lCase(arrStrNode(0))="home" Then
									Set oCurrentNode = objJavaWindowCPD.JavaTree("NavTree").Object.getItem(0)
									intNodeCount = oCurrentNode.getItemCount()
									For iCount=0 To intNodeCount-1
										If iCount=0 Then
											sReturn=oCurrentNode.getItem(iCount).getData().toString()
										Else
											sReturn=sReturn+","+oCurrentNode.getItem(iCount).getData().toString()
										End If
									Next
										arr = Split(sReturn,",")
										Fn_CPD_NavTree_NodeOperation = arr
										Set oCurrentNode=Nothing
										Exit Function
							Else
									Set oCurrentNode = objJavaWindowCPD.JavaTree("NavTree").Object.getItem(0)
									intNodeCount=0
									For each echStrNode In arrStrNode
										iRows = oCurrentNode.getItemCount()
										For iCounter = 0 to iRows - 1
											If oCurrentNode.getItem(iCounter).getData().toString() = echStrNode Then
												Set oCurrentNode=oCurrentNode.getItem(iCounter)
												intNodeCount = oCurrentNode.getItemCount()
												Exit For
											End If
										Next
									Next 
									For iCount=0 To intNodeCount-1
										If iCount=0 Then
											sReturn=oCurrentNode.getItem(iCount).getData().toString()
										Else
											sReturn=sReturn+","+oCurrentNode.getItem(iCount).getData().toString()
										End If
									Next
									arr = Split(sReturn,",")
									Fn_CPD_NavTree_NodeOperation = arr
									Set oCurrentNode=Nothing
							End If
						Else
							Fn_CPD_NavTree_NodeOperation = False
						End If
		'PopUp Menu Enabled or not on multi Select - [TC1121-20151116a-07_12_2015-VivekA-NewDevelopment] - Added by Reema W
		Case "MultiSelectContextMenuEnabled"
				NodeLists = split(StrNodeName,",",-1,1)
				Set objJavaTreeNav = objJavaWindowCPD.JavaTree("NavTree")
				Call Fn_CPD_NavTree_NodeOperation("Multiselect",StrNodeName,"")
				iPath = Fn_UI_JavaTreeGetItemPath(objJavaWindowCPD.JavaTree("NavTree"),NodeLists(0))
				iPath = Fn_UI_JavaTreeGetItemPathExt("Fn_CPD_NavTree_NodeOperation", objJavaWindowCPD.JavaTree("NavTree"), NodeLists(0) , ":", "@")
				Call Fn_UI_JavaTree_OpenContextMenu("Fn_CPD_NavTree_NodeOperation",objJavaWindowCPD,"NavTree",iPath)
				If objJavaWindowCPD.WinMenu("ContextMenu").GetItemProperty(StrMenu,"Enabled") = True Then
					Fn_CPD_NavTree_NodeOperation = True
				Else
					Fn_CPD_NavTree_NodeOperation = False
			  	End If
		'[TC1121-20151116b-10_12_2015-VivekA-NewDevelopment] - Added to verify selected node in Nav Tree
		Case "VerifySelectedNode"
				Set objJavaTreeNav = objJavaWindowCPD.JavaTree("NavTree")
				If objJavaTreeNav.Object.getSelectionCount() <> 0 Then 
					Set oCurrentNode = objJavaTreeNav.Object.getFocusItem()
	
					If IsObject(oCurrentNode) then
						sVerifyNode = oCurrentNode.getData().toString() 
											
						Do while IsObject(oCurrentNode.getParentItem())
							Set oCurrentNode = oCurrentNode.getParentItem()
							sVerifyNode = oCurrentNode.getData().toString() & ":" & sVerifyNode 
						Loop
					End If
					
					If Trim(StrNodeName) = Trim(sVerifyNode) Then
						Fn_CPD_NavTree_NodeOperation = True
					Else
						Fn_CPD_NavTree_NodeOperation = False
					End If
					Set oCurrentNode = Nothing	
				Else
					Fn_CPD_NavTree_NodeOperation = False
				End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case Else
			Fn_CPD_NavTree_NodeOperation = False
	End Select

	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), StrAction &" Sucessfully completed on Node [" + StrNodeName + "] of JavaTree of function Fn_CPD_NavTree_NodeOperation")
	Set objJavaWindowCPD = nothing
	Set objJavaTreeNav = nothing
End Function
''*********************************************************		Function to action perform Content Search	***********************************************************************
'Function Name			:		Fn_CPD_ContentSearchOperations

'Description			:		1. sAction - Action to be performed
'								2. sSearchScope : Scope name
'								3. sSearchType : Search Type
'								4. dicContentSearch : Dictionary object
'								5. sBtnName = Button Name

'Return Value		    : 		TRUE \ FALSE

'Pre-requisite			:		Content Search panel should be displayed

'Examples				:		dicContentSearch("Scheme") = ""
'								dicContentSearch("SearchCriteria") = "Design Element Name=de~Category=Shape~Owning User Name=Koustubh Watwe"
'								msgbox  Fn_CPD_ContentSearchOperations("Search", "sub", "Attribute", dicContentSearch, "")
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------								
'								dicContentSearch("DistanceFromOrigin") = 1.0
'								dicContentSearch("DefinePlaneAxis") = "along Y axis"
'								dicContentSearch("Option") = "Above or Intersecting"
'								dicContentSearch("TrueShapeFiltering")  = True
'								msgbox  Fn_CPD_ContentSearchOperations("Search", "sub", "Plane Zone", dicContentSearch, "")
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'								dicContentSearch("SearchCriteria") = "Volume Extent Minimum Y=2.2~Volume Extent Maximum Y=2.2"
'								dicContentSearch("DistanceFromOrigin") = 1.0
'								dicContentSearch("DefinePlaneAxis") = "along Y axis"
'								dicContentSearch("Option") = "Above or Intersecting"
'								dicContentSearch("TrueShapeFiltering")  = True
'								dicContentSearch("IncludeChildPartitions") = True / False
'								msgbox  Fn_CPD_ContentSearchOperations("Search", "sub", "Box Zone", dicContentSearch, "")

'History				:		
'	Developer Name				Date						Rev. No.			Changes Done																							Reviewer
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Koustubh W				   21/01/2012			         1.0				Created
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Koustubh W				   31/01/2012			         1.0				Modified case Partition, added code to select partition
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Koustubh W				   10/04/2012			         1.0				Added case VerifyDesignElementDetails
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Koustubh W				   03/05/2012			         1.0				Modified code to set data in Attribute Case
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Koustubh W				   06/08/2012			         1.0				Modified case "Box Zone" ,added temparary code for build 0718
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Jotiba					   08/07/2015				  	 1.0				Modified cases "Plane Zone" and "Box Zone" as per Archana's comment
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Chaitali R				   30/11/2015				  	 1.0				Modified Case "Search" : Case "Attribute" - 										[TC1121-2015110900-30_11_2015-VivekA-NewDevelopment]
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Nitish B.				   24/12/2015				  	 1.0				Added Case "Search_Edit"  															[TC1122-20151116d00-24_12_2015-AnkitN-NewDevelopment]
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_CPD_ContentSearchOperations(sAction, sSearchScope, sSearchType, dicContentSearch, sBtnName)
	GBL_FAILED_FUNCTION_NAME="Fn_CPD_ContentSearchOperations"
	Dim objSearch, bReturn, iCount, aSearchCriterias
	Dim aSearchData, objSelectType, objIntNoOfObjects, innerCntr
	Dim iIndex, sLabel
	iIndex = 0

	Set objSearch = JavaWindow("Collaborative Product")
	objSearch.JavaStaticText("Search_Type").SetTOProperty "Index", iIndex
	Fn_CPD_ContentSearchOperations = False

	bReturn = objSearch.JavaList("SearchScope").Exist(5)
	If bReturn = False Then
		bReturn = Fn_CPD_CompnentTabOperations("Activate","Content Search","")
		If bReturn = False Then
			' failed to open content search tab
			If JavaWindow("Collaborative Product").JavaTree("NavTree").Exist(2) Then
				JavaWindow("Collaborative Product").JavaTree("NavTree").Click 0, 0,"LEFT"	
			End If
			Call Fn_ToolbarButtonClick_Ext(1,"Search within current content")
			If objSearch.JavaList("SearchScope").Exist(15) = False Then
				' Content Search panel is not opened
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_CPD_ContentSearchOperations : Content Search Panel is not opened for case [ " +sAction+ " ].")
				Exit function
			End If
		End If
	End If
	
	If Fn_CPD_CompnentTabOperations("IsMaximized","Content Search", "") = False Then  ' Modified by Chaitali Rane
		Call Fn_CPD_CompnentTabOperations("DoubleClick","Content Search", "")  
		Call Fn_ReadyStatusSync(2)
	End If 
	
	Select Case sAction
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Search","Search_Edit"											'[TC1122-20151116d00-24_12_2015-AnkitN-NewDevelopment] - Added case by Nitish B.
			If sSearchScope <> "" Then
				Call Fn_List_Select("Fn_CPD_ContentSearchOperations",objSearch,"SearchScope", sSearchScope)
			End If

			If sSearchType <> "" And  sAction <> "Search_Edit" Then
				Call Fn_List_Select("Fn_CPD_ContentSearchOperations",objSearch,"SearchType", sSearchType)
			End If

			Select Case sSearchType
				' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				Case "Partition"
					If dicContentSearch("Scheme") <> "" Then
						Call Fn_List_Select("Fn_CPD_ContentSearchOperations",objSearch,"Scheme", dicContentSearch("Scheme"))
					End If

					'Checkbox for child Partition
					If dicContentSearch("IncludeChildPartitions") <> "" Then
							If cBool(dicContentSearch("IncludeChildPartitions") ) Then
								Call Fn_CheckBox_Set("Fn_CPD_ContentSearchOperations",objSearch,"IncludeChildPartitions", "ON")
							Else
								Call Fn_CheckBox_Set("Fn_CPD_ContentSearchOperations",objSearch,"IncludeChildPartitions", "OFF")
							End If
					End If

					If dicContentSearch("PartitionNode") <> "" Then
                        Call Fn_ReadyStatusSync(1)
						iTreeIndex = Fn_UI_JavaTreeGetItemPathExt("Fn_CPD_ContentSearchOperations", JavaWindow("Collaborative Product").JavaTree("NavTree"), dicContentSearch("PartitionNode"), "", "")
						If iTreeIndex <> False Then
							JavaWindow("Collaborative Product").JavaTree("NavTree").select iTreeIndex
						Else
							Fn_CPD_ContentSearchOperations = False
							Exit function
						End If
					End If

				' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				Case "Attribute"
						If dicContentSearch("AttributeType")<> "" Then
								'Call Fn_CPD_CompnentTabOperations("DoubleClick", "Content Search","")
								Call Fn_List_Select("Fn_CPD_ContentSearchOperations",objSearch,"AttributeCombo", dicContentSearch("AttributeType"))
								'Call Fn_CPD_CompnentTabOperations("DoubleClick", "Content Search","")
						End If
						aSearchCriterias = split( dicContentSearch("SearchCriteria") , "~")
						For iCount = 0 to UBound(aSearchCriterias)
							aSearchData = split(aSearchCriterias(iCount), "=")
							Select Case aSearchData(0)
								Case "Design Element Name"
									aSearchData(0)= Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("CPD_DisplayName"),aSearchData(0))
								Case "Design Element ID"
									aSearchData(0)= Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("CPD_DisplayName"),aSearchData(0))
							End Select
							
							'Call Fn_CPD_CompnentTabOperations("DoubleClick", "Content Search","")
							objSearch.JavaStaticText("Search_Type").SetTOProperty "label", aSearchData(0) & ":"
							For iIndex = 0 to 1
								objSearch.JavaStaticText("Search_Type").SetTOProperty "Index", "1"
								If objSearch.JavaButton("Search_DropDown").Exist(3) Then
										Call Fn_Button_Click( "Fn_CPD_ContentSearchOperations", objSearch,"Search_DropDown" )
										wait 3
										Set objSelectType=Description.Create()
										objSelectType("Class Name").value = "JavaStaticText"					
										Set  objIntNoOfObjects = objSearch.ChildObjects(objSelectType)
										For  innerCntr = 0 to objIntNoOfObjects.count - 1
											   If objIntNoOfObjects(innerCntr).getROProperty("label") = aSearchData(1) Then
													objIntNoOfObjects(innerCntr).Click 2,2
													bReturn = TRUE
													Exit for
											   End If
										Next
										If  bReturn = True Then
											Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Value for [" & aSearchData(1) & "] Successfully set.  ")   	
										End If
								ElseIf objSearch.JavaEdit("Search_Editbox").Exist(3) Then
										Call Fn_Edit_Box("Fn_CPD_ContentSearchOperations",objSearch,"Search_Editbox", aSearchData(1))
										bReturn = true
										Exit for
								ElseIf objSearch.JavaList("Search_List").Exist(3)  Then								        
								        	Call Fn_List_Select("Fn_CPD_ContentSearchOperations",objSearch,"Search_List", aSearchData(1))
										bReturn = true
										Exit for
								End If
							Next
							If bReturn = False Then
									Fn_CPD_ContentSearchOperations = False
									Exit function
							End If
						Next
						'[TC1121-2015110900-30_11_2015-VivekA-NewDevelopment] - Added by Chaitali R
						If dicContentSearch("Option") <> "" Then
							Call Fn_SISW_UI_Twistie_Operations("Fn_CPD_ContentSearchOperations", "Expand", JavaWindow("Collaborative Product"), "Twistie", "Search Options","SearchOptions")
			            	Call Fn_List_Select("Fn_CPD_ContentSearchOperations",objSearch,"AttrSearchOptionsCombo", dicContentSearch("Option"))
						End If
						'------------------------------------------------------
				' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				Case "Proximity"
					If dicContentSearch("SearchCriteria") <> "" Then
						aSearchData = split( dicContentSearch("SearchCriteria") , "=")
						objSearch.JavaStaticText("Search_Type").SetTOProperty "label", aSearchData(0) 
						If objSearch.JavaEdit("Search_Editbox").Exist(3) Then
								Call Fn_Edit_Box("Fn_CPD_ContentSearchOperations",objSearch,"Search_Editbox", aSearchData(1))
						Else
							Fn_CPD_ContentSearchOperations = False
							Exit function
						End If
					End If

					If dicContentSearch("TrueShapeFiltering") <> "" Then
						If cBool(dicContentSearch("TrueShapeFiltering") ) Then
							Call Fn_CheckBox_Set("Fn_CPD_ContentSearchOperations",objSearch,"TrueShapeFiltering", "ON")
						Else
							Call Fn_CheckBox_Set("Fn_CPD_ContentSearchOperations",objSearch,"TrueShapeFiltering", "OFF")
						End If
					End If
				' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				Case "Box Zone"
					aSearchCriterias = split( dicContentSearch("SearchCriteria") , "~")
					For iCount = 0 to UBound(aSearchCriterias)
						aSearchData = split(aSearchCriterias(iCount), "=")

						Select Case aSearchData(0)
							Case "Volume Extent Minimum X"
								sLabel = "X"
							Case "Volume Extent Minimum Y"
								sLabel = "Y"
							Case "Volume Extent Minimum Z"
								sLabel = "Z"
							Case "Volume Extent Maximum X"
								sLabel = "X"
								iIndex = 1
							Case "Volume Extent Maximum Y"
								sLabel = "Y"
								iIndex = 1
							Case "Volume Extent Maximum Z"
								sLabel = "Z"
								iIndex = 1
						End Select

						objSearch.JavaStaticText("Search_Type").SetTOProperty "label", sLabel &":"
						objSearch.JavaStaticText("Search_Type").SetTOProperty "Index", iIndex
						' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
						' Temparary solution for Build 718 - Koustubh [ 6-Aug-2012]
						Set objSelectType=description.Create()
						objSelectType("Class Name").value = "JavaStaticText"
						objSelectType("label").value = sLabel &":"
						Set  intNoOfObjects = JavaWindow("Collaborative Product").ChildObjects(objSelectType)

						If instr(aSearchData(0),"Maximum") > 0 Then
							objSearch.JavaStaticText("Search_Type").SetTOProperty "Index", intNoOfObjects.count - 1
						Else
							objSearch.JavaStaticText("Search_Type").SetTOProperty "Index", intNoOfObjects.count - 2
						End If
						Set objSelectType = Nothing
						Set intNoOfObjects = Nothing

						' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
						If objSearch.JavaEdit("Search_Editbox").Exist(3) Then
								Call Fn_Edit_Box("Fn_CPD_ContentSearchOperations",objSearch,"Search_Editbox", aSearchData(1))
						Else
								Exit function
						End If
					Next
					Wait 4
					
					'code added to click on search option if it is not expanded
					'As per Archana's comment updating funciton"
					'Call Fn_SISW_UI_Twistie_Operations("Fn_CPD_ContentSearchOperations", "Expand", JavaWindow("Collaborative Product"), "Twistie", "Search Options","SearchOptions")
					
					' Modified by Chaitali Rane [ TC11.4 - Maintenance ]
					If JavaWindow("Collaborative Product").JavaObject("Twistie").Exist = True Then
						Call Fn_SISW_UI_Twistie_Operations("Fn_CPD_ContentSearchOperations", "Expand", JavaWindow("Collaborative Product"), "Twistie", "Search","SearchOptions")
					End If
					
					'Amit T - Changed below code -> First Checked/Unchecked "True Shape Filtering" then Selected Item from "Option"

					If dicContentSearch("TrueShapeFiltering") <> "" Then
						If cBool(dicContentSearch("TrueShapeFiltering") ) Then
							Call Fn_CheckBox_Set("Fn_CPD_ContentSearchOperations",objSearch,"TrueShapeFiltering", "ON")
						Else
							Call Fn_CheckBox_Set("Fn_CPD_ContentSearchOperations",objSearch,"TrueShapeFiltering", "OFF")
						End If
					End If
					
					If dicContentSearch("Option") <> "" Then
						Call Fn_List_Select("Fn_CPD_ContentSearchOperations",objSearch,"OptionsCombo", dicContentSearch("Option"))
					End If

				' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				Case "Plane Zone"
					objSearch.JavaStaticText("Search_Type").SetTOProperty "label",  "Distance From Origin:"
					If objSearch.JavaEdit("Search_Editbox").Exist(3) Then
							Call Fn_Edit_Box("Fn_CPD_ContentSearchOperations",objSearch,"Search_Editbox", dicContentSearch("DistanceFromOrigin"))
					Else
						Fn_CPD_ContentSearchOperations = False
					End If

					If dicContentSearch("DefinePlaneAxis") <> "" Then
						Call Fn_List_Select("Fn_CPD_ContentSearchOperations",objSearch,"DefinePlaneCombo", dicContentSearch("DefinePlaneAxis"))	
						If dicContentSearch("DefinePlaneAxis") = "along other axis..." Then
							aSearchCriterias = split( dicContentSearch("SearchCriteria") , "~")
							For iCount = 0 to UBound(aSearchCriterias)
								aSearchData = split(aSearchCriterias(iCount), "=")
								objSearch.JavaStaticText("Search_Type").SetTOProperty "label", aSearchData(0)  &":"
								If objSearch.JavaEdit("Search_Editbox").Exist(3) Then
										Call Fn_Edit_Box("Fn_CPD_ContentSearchOperations",objSearch,"Search_Editbox", aSearchData(1))
								Else
									Exit function
								End If
							Next
						End If
					End If
					
					'code added to click on search option if it is not expanded
					'As per Archana's comment updating funciton"
					If JavaWindow("Collaborative Product").JavaObject("Twistie").Exist = True Then
					Call Fn_SISW_UI_Twistie_Operations("Fn_CPD_ContentSearchOperations", "Expand", JavaWindow("Collaborative Product"), "Twistie", "Search Options","SearchOptions")
					End If
					If dicContentSearch("TrueShapeFiltering") <> "" Then
						If cBool(dicContentSearch("TrueShapeFiltering") ) Then
							Call Fn_CheckBox_Set("Fn_CPD_ContentSearchOperations",objSearch,"TrueShapeFiltering", "ON")
						Else
							Call Fn_CheckBox_Set("Fn_CPD_ContentSearchOperations",objSearch,"TrueShapeFiltering", "OFF")
						End If
					End If

					If dicContentSearch("Option") <> "" Then
					Call Fn_List_Select("Fn_CPD_ContentSearchOperations",objSearch,"OptionsComboPlaneZone", dicContentSearch("Option"))
					End If

					
							
				' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				Case Else
					If dicContentSearch("PartitionNode") <> "" Then
						iTreeIndex = Fn_UI_JavaTreeGetItemPathExt("Fn_CPD_ContentSearchOperations", JavaWindow("Collaborative Product").JavaTree("NavTree"), dicContentSearch("PartitionNode"), "", "")
						If iTreeIndex <> False Then
							JavaWindow("Collaborative Product").JavaTree("NavTree").select iTreeIndex
						Else
							Fn_CPD_ContentSearchOperations = False
							Exit function
						End If
					End If
			End Select
			Fn_CPD_ContentSearchOperations = True
			
			If Fn_CPD_CompnentTabOperations("IsMaximized","Content Search", "") = True Then  ' Modified by Chaitali Rane
				Call Fn_CPD_CompnentTabOperations("DoubleClick","Content Search", "")  
				Call Fn_ReadyStatusSync(2)
			End If 
			If sBtnName <> "" Then
'				Call Fn_Button_Click( "Fn_CPD_ContentSearchOperations", objSearch, sBtnName)
				If sBtnName = "Filter" or sBtnName = "Include" or sBtnName = "Replay" Then ' Modified by Chaitali Rane
					objSearch.JavaButton("Search").SetTOProperty "label", sBtnName
					objSearch.JavaButton("Search").WaitProperty "enabled", "1"
					objSearch.JavaButton("Search").Object.click
				Else
					objSearch.JavaButton(sBtnName).WaitProperty "enabled", "1"
					objSearch.JavaButton(sBtnName).Object.click
				End If
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "VerifyDesignElementDetails"
				If sSearchScope <> "" Then
					Fn_CPD_ContentSearchOperations = Fn_UI_ListItemExist("Fn_CPD_ContentSearchOperations",objSearch,"SearchScope", sSearchScope)
					If Fn_CPD_ContentSearchOperations = False Then
						Exit function
					End If
				End If
	
				If sSearchType <> "" Then
					Fn_CPD_ContentSearchOperations = Fn_UI_ListItemExist("Fn_CPD_ContentSearchOperations",objSearch,"SearchType", sSearchType)
					If Fn_CPD_ContentSearchOperations = False Then
						Exit function
					End If
				End If

				aSearchCriterias = split( dicContentSearch("SearchCriteria") , "~")
					For iCount = 0 to UBound(aSearchCriterias)
						aSearchData = split(aSearchCriterias(iCount), "=")
						iIndex = 0
						Select Case aSearchData(0)
'							Case "Volume Extent Minimum X"
'								sLabel = "X"
'							Case "Volume Extent Minimum Y"
'								sLabel = "Y"
'							Case "Volume Extent Minimum Z"
'								sLabel = "Z"
'							Case "Volume Extent Maximum X"
'								sLabel = "X"
'								iIndex = 1
'							Case "Volume Extent Maximum Y"
'								sLabel = "Y"
'								iIndex = 1
'							Case "Volume Extent Maximum Z"
'								sLabel = "Z"
'								iIndex = 1
									Case Else
										sLabel = trim(aSearchData(0))
						End Select

						objSearch.JavaStaticText("Search_Type").SetTOProperty "label", sLabel &":"
						objSearch.JavaStaticText("Search_Type").SetTOProperty "Index", iIndex
						If objSearch.JavaEdit("Search_Editbox").Exist(3) Then
								If trim(objSearch.JavaEdit("Search_Editbox").getROProperty("value"))  = trim(aSearchData(1)) Then
									Fn_CPD_ContentSearchOperations = True
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_CPD_ContentSearchOperations : Successfully verified [ " + aSearchData(0) + " = " & trim(aSearchData(1)) & " ].")
								Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_CPD_ContentSearchOperations : Successfully verified [ " + aSearchData(0) + " <> " & trim(aSearchData(1)) & " ].")
									Fn_CPD_ContentSearchOperations = False
									Exit for
								End IF
						Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_CPD_ContentSearchOperations : Failed to find field[ " + aSearchData(0) & " ].")
								Fn_CPD_ContentSearchOperations = False
								Exit function
						End If
					Next
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_CPD_ContentSearchOperations : Invalid action [ " +sAction+ " ] is requested.")
	End Select
	IF Fn_CPD_ContentSearchOperations <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_CPD_ContentSearchOperations : Executed successfully with case [" + sAction + "].")
	End If
	Set objSearch = Nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name			:	Fn_CPD_DateControl
'@@
'@@    Description				:	Function Used to set date control
'@@
'@@    Parameters			    :	1. sAction	: Action to be performed
'@@								:	2. sDate	: Date in format ( DD-MMM-YYYY eg. 04-Jan-2012 )
'@@								:	3. sTime 	: Time in 24 hrs format ( HH:MM:SS eg. 21:45:10 )
'@@
'@@    Return Value		   	   	: 	True Or False
'@@
'@@    Pre-requisite			:	CPD perspective should be activated.						
'@@
'@@    Examples					:	Call  Fn_CPD_DateControl("Set", "02-Jan-2012", "18:30:00")
'@@    Examples					:	Call  Fn_CPD_DateControl("Today", "", "")
'@@    Examples					:	Call  Fn_CPD_DateControl("Clear", "", "")
'@@
'@@	   History					:	
'@@				Developer Name				Date			Rev. No.	Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe			25-Jan-2012			1.0			Created
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_CPD_DateControl(sAction, sDate, sTime)
	GBL_FAILED_FUNCTION_NAME="Fn_CPD_DateControl"
	Dim objDateCtrl, sDateStr
	Set objDateCtrl = JavaWindow("Collaborative Product").JavaWindow("DateControl")
	Fn_CPD_DateControl = False
	If Fn_UI_ObjectExist("Fn_CPD_DateControl",objDateCtrl) = False Then
        Call Fn_WriteLogFile(Environment.Value("TestLogFile"),  "FAIL : Fn_CPD_DateControl : " & objDateCtrl.ToString &" does not Exist of Function Fn_CPD_DateControl")
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
				Call Fn_Button_Click("Fn_CPD_DateControl",objDateCtrl,"OK")
				Fn_CPD_DateControl = True
		Case "Clear"
				Call Fn_Button_Click("Fn_CPD_DateControl",objDateCtrl,"Clear")
				Fn_CPD_DateControl = True
		Case "Today"
				Call Fn_Button_Click("Fn_CPD_DateControl",objDateCtrl,"Today")
				Fn_CPD_DateControl = True
		Case Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Fn_CPD_DateControl : Invalid Action [ " & sAction & " ] ")
	End Select
	If Fn_CPD_DateControl = True Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Fn_CPD_DateControl : Executed successfully with Action [ " & sAction & " ] ")
	End If
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_CPD_CreatePartitionScheme
'
'Description			 :	Function Used to Create Partition Scheme
'
'Parameters			   :   '1.sAction: Action Name
'										 2.sSchemeType: Partition Scheme Type
'										3.sName : Partition Name
'										4.sDescription : Partition Description
'
'Return Value		   : 	True Or False
'
'Pre-requisite			:	CPD perspective should be activated.
'
'Examples				:   Call Fn_CPD_CreatePartitionScheme("Create", "Partition Scheme Functional","PSFunctional3", "")
'										Call Fn_CPD_CreatePartitionScheme("Create", "Partition Scheme Physical","PSPhysical1", "Partition Scheme Physical Description")
'										Call Fn_CPD_CreatePartitionScheme("Create", "Partition Scheme Spatial","PSSpatial1", "Partition Scheme Spatial Description")
'
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												27-Jan-2012								1.0																						Sachin Joshi
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Public Function Fn_CPD_CreatePartitionScheme(sAction, sSchemeType,sName, sDescription)
	GBL_FAILED_FUNCTION_NAME="Fn_CPD_CreatePartitionScheme"
	Dim objCDCreate,sOldSchemeType
	Fn_CPD_CreatePartitionScheme = False
	
	sOldSchemeType = sSchemeType
	sSchemeType = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("CPD_DisplayName"),sSchemeType)
	If sSchemeType = False Then
		sSchemeType = sOldSchemeType
	End If
	
	Set objCDCreate = JavaWindow("Collaborative Product").JavaWindow("NewBusinessObject")
	
	'If Not objCDCreate.Exist(6) Then
	If Fn_SISW_UI_Object_Operations("Fn_CPD_CreatePartitionScheme","Exist",objCDCreate,SISW_MIN_TIMEOUT) = False Then
		Call Fn_ToolBarOperation("ShowDropdownAndSelectWithInstance", "View Menu:1", "Create Partition Scheme")
		Call Fn_ReadyStatusSync(3)
		'Added condition as the index of tool tip is changing in Tc 12 20171025 build   by pratap
		If Fn_SISW_UI_Object_Operations("Fn_CPD_CreatePartitionScheme","Exist",objCDCreate,SISW_MIN_TIMEOUT) = False Then
			Call Fn_ToolBarOperation("ShowDropdownAndSelectWithInstance", "View Menu:2", "Create Partition Scheme")
	    End If
    
    End If
	
	'If Fn_UI_ObjectExist("Fn_CPD_CreatePartitionScheme", objCDCreate) = False Then
	If Fn_SISW_UI_Object_Operations("Fn_CPD_CreatePartitionScheme","Exist",objCDCreate,SISW_MIN_TIMEOUT) = False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_CPD_CreatePartitionScheme ] Failed to opn Create Partition Scheme window.")
		Set objCDCreate = Nothing
	End If
    
	Select Case sAction
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "Create"
			' select collaborative design from tree
			objCDCreate.JavaTree("BusinessObjectType").Expand "Complete List"
			wait 1
			objCDCreate.JavaTree("BusinessObjectType").Select "Complete List:"+sSchemeType

			' click on next
			Call Fn_Button_Click("Fn_CPD_CreatePartitionScheme",objCDCreate,"Next" )
			'set name
			If sName <> "" Then
				objCDCreate.JavaStaticText("Field").SetTOProperty "label", "Name:"
				Call Fn_Edit_Box("Fn_CPD_CreatePartitionScheme",objCDCreate,"Field","")
				objCDCreate.JavaEdit("Field").Type sName
				Call Fn_ReadyStatusSync(1)
			End If

			' set description
			If sDescription <> "" Then
				objCDCreate.JavaStaticText("Field").SetTOProperty "label", "Description:"
				Call Fn_Edit_Box("Fn_CPD_CreatePartitionScheme",objCDCreate,"Field",sDescription)
			End If
			' click on finish
			Call Fn_Button_Click("Fn_CPD_CreatePartitionScheme",objCDCreate,"Finish" )
			Call Fn_ReadyStatusSync(1)
			If objCDCreate.Exist(5) Then
				Call Fn_Button_Click("Fn_CPD_CreatePartitionScheme",objCDCreate,"Cancel" )
			End If

			Fn_CPD_CreatePartitionScheme=True
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_CPD_CreatePartitionScheme ] Invalid case [ " & sAction & " ].")
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
	End Select
	If  Fn_CPD_CreatePartitionScheme <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [ Fn_CPD_CreatePartitionScheme ] executed successfuly with case [ " & sAction & " ].")
	End If
	Set objCDCreate = Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_CPD_CreatePartition
'
'Description			 :	Function Used to Create Partition
'
'Parameters			   :   '1.sAction: Action Name
'										 2.dicPartitionInfo: Partition Information
'
'Return Value		   : 	True Or False
'
'Pre-requisite			:	CPD perspective should be activated.
'
'Examples				:   dicPartitionInfo("PartitionType")="Partition Design"
'										dicPartitionInfo("Name")="PartitionDesign1"
'										dicPartitionInfo("Description")="PartitionDesignDesc"
'										Call Fn_CPD_CreatePartition("Create", dicPartitionInfo)
'
'										dicPartitionInfo("PartitionType")="Partition Manufacturing"
'										dicPartitionInfo("Name")="PartitionManufacturing1"
'										dicPartitionInfo("Description")="PartitionManufacturingDesc"
'										Call Fn_CPD_CreatePartition("Create", dicPartitionInfo)
'										Call Fn_CPD_CreatePartition("GetErrorMessageOnCreate", dicPartitionInfo)
'
'History					 :			
'			Developer Name				Date				Rev. No.	Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'			Sandeep N					27-Jan-2012			1.0												Sachin Joshi
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'			KOustubh W					20-Feb-2012			1.0			Added case GetErrorMessageOnCreate
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Public Function Fn_CPD_CreatePartition(sAction, dicPartitionInfo)
	GBL_FAILED_FUNCTION_NAME="Fn_CPD_CreatePartition"
	Dim objCDCreate, sOldPartitionType
	Fn_CPD_CreatePartition = False
	Set objCDCreate = JavaWindow("Collaborative Product").JavaWindow("NewBusinessObject")

	sOldPartitionType = dicPartitionInfo("PartitionType")
	dicPartitionInfo("PartitionType") = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("CPD_DisplayName"),dicPartitionInfo("PartitionType"))
	If dicPartitionInfo("PartitionType") = False Then
		dicPartitionInfo("PartitionType") = sOldPartitionType
	End If
	If Not objCDCreate.Exist(6) Then
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - \
		Call Fn_ToolbarOperation("Click", "Create Partition","")
		'Checking Partion Creation Dialog Open or not		
		Call Fn_ReadyStatusSync(3)
		If Fn_UI_ObjectExist("Fn_CPD_CreatePartition", objCDCreate) = False Then
            Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_CPD_CreatePartition ] Failed to open Partition window.")
			Set objCDCreate = Nothing
		End IF
	End IF

	Select Case sAction
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "Create", "GetErrorMessageOnCreate"
			' select collaborative design from tree
			If objCDCreate.JavaTree("BusinessObjectType").exist(2) then
				objCDCreate.JavaTree("BusinessObjectType").Expand "Complete List"
				wait 1
				objCDCreate.JavaTree("BusinessObjectType").Select "Complete List:"+dicPartitionInfo("PartitionType")
	
				' click on next
				Call Fn_Button_Click("Fn_CPD_CreatePartition",objCDCreate,"Next" )
				wait(2)
			End If 
			' if ModelD is empty
			objCDCreate.JavaStaticText("Field").SetTOProperty "label", "ID:"
			If dicPartitionInfo("PartitionID") = "" Then
				'	then click on assign
				Call Fn_Button_Click("Fn_CPD_CreatePartition",objCDCreate,"Assign" )
				Call Fn_ReadyStatusSync(1)
				wait 3
				Fn_CPD_CreatePartition=Fn_UI_Object_GetROProperty("",objCDCreate.JavaEdit("Field"), "value")
'				Fn_CPD_CreatePartition = objCDCreate.JavaEdit("Field").GetROProperty("value")
               wait 2
			Else
				Call Fn_Edit_Box("Fn_CPD_CreatePartition",objCDCreate,"Field",dicPartitionInfo("PartitionID"))
				Fn_CPD_CreatePartition = True
			End If

			'set name
			If dicPartitionInfo("Name") <> "" Then
				objCDCreate.JavaStaticText("Field").SetTOProperty "label", "Name:"
				'objCDCreate.JavaEdit("Field").Type dicPartitionInfo("Name")
				Call Fn_Edit_Box("Fn_CPD_CreatePartition",objCDCreate,"Field",dicPartitionInfo("Name"))
				Call Fn_ReadyStatusSync(1)
			End If

			' set description
			If dicPartitionInfo("Description") <> "" Then
				objCDCreate.JavaStaticText("Field").SetTOProperty "label", "Description:"
				Call Fn_Edit_Box("Fn_CPD_CreatePartition",objCDCreate,"Field",dicPartitionInfo("Description"))
			End If

			If dicPartitionInfo("OpenOnCreate")<>"" AND dicPartitionInfo("OpenOnCreate")<> "OFF" Then
				'Call Fn_CheckBox_Set("Fn_CPD_CreatePartition",objCDCreate, "OpenOnCreate",dicPartitionInfo("OpenOnCreate"))
				wait(1)
				Call Fn_SISW_UI_JavaCheckBox_Operations("Fn_CPD_CreatePartition", "Set", objCDCreate, "OpenOnCreate", "ON")
			ElseIf dicPartitionInfo("OpenOnCreate")= "OFF" Then
				wait(1)
				Call Fn_SISW_UI_JavaCheckBox_Operations("Fn_CPD_CreatePartition", "Set", objCDCreate, "OpenOnCreate", "OFF")
				'Call Fn_CheckBox_Set("Fn_CPD_CreatePartition",objCDCreate, "OpenOnCreate",dicPartitionInfo("OpenOnCreate"))
			End If

			' click on next
			Call Fn_Button_Click("Fn_CPD_CreatePartition",objCDCreate,"Next" )
			wait(2)
			'Setting Create Partition Item Option
			If dicPartitionInfo("CreatePartitionItem")<>""  Then
				Call Fn_CheckBox_Set("Fn_CPD_CreatePartition",objCDCreate, "CreatePartitionItem",dicPartitionInfo("CreatePartitionItem"))
			End If
			'Selecting Partition Item Type
			If dicPartitionInfo("PartitionItemType")<>""  Then
				Select Case dicPartitionInfo("PartitionItemType")
					Case "Partition Functional Item"
						dicPartitionInfo("PartitionItemType") = "Functional Partition Item"
				End Select	
				Call Fn_List_Select("Fn_CPD_CreatePartition", objCDCreate, "CreateOptions",dicPartitionInfo("PartitionItemType"))
			End If
			'Setting Copy Effectivity Option
			If dicPartitionInfo("CopyEffectivity")<>""  Then
				Call Fn_CheckBox_Set("Fn_CPD_CreatePartition",objCDCreate, "CopyEffectivity",dicPartitionInfo("CopyEffectivity"))
			End If
			'Setting Check out on create Option
			If dicPartitionInfo("CheckOutOnCreate")<>""  Then
				Call Fn_CheckBox_Set("Fn_CPD_CreatePartition",objCDCreate, "CheckOutOnCreate",dicPartitionInfo("CheckOutOnCreate"))
			End If
			'Setting open on create Option
			'If dicPartitionInfo("OpenOnCreate")<>"" AND dicPartitionInfo("OpenOnCreate")<> "OFF" Then
			'	'Call Fn_CheckBox_Set("Fn_CPD_CreatePartition",objCDCreate, "OpenOnCreate",dicPartitionInfo("OpenOnCreate"))
			'	wait(1)
			'	Call Fn_SISW_UI_JavaCheckBox_Operations("Fn_CPD_CreatePartition", "Set", objCDCreate, "OpenOnCreate", "ON")
			'ElseIf dicPartitionInfo("OpenOnCreate")= "OFF" Then
			'	wait(1)
			'	Call Fn_SISW_UI_JavaCheckBox_Operations("Fn_CPD_CreatePartition", "Set", objCDCreate, "OpenOnCreate", "OFF")
			'	'Call Fn_CheckBox_Set("Fn_CPD_CreatePartition",objCDCreate, "OpenOnCreate",dicPartitionInfo("OpenOnCreate"))
			'End If

			' click on finish
			Call Fn_Button_Click("Fn_CPD_CreatePartition",objCDCreate,"Finish" )
			Call Fn_ReadyStatusSync(1)
			
			If sAction = "GetErrorMessageOnCreate" Then
				Fn_CPD_CreatePartition = False
				If objCDCreate.JavaWindow("Error").Exist(15) Then
					'Fn_CPD_CreatePartition = objCDCreate.JavaWindow("Error").JavaStaticText("ErrorMsg").GetROProperty("value")
					wait 2
					Fn_CPD_CreatePartition = objCDCreate.JavaWindow("Error").JavaEdit("ErrorMsg").GetROProperty("value")
					wait 2
					Call Fn_Button_Click("Fn_CPD_CreatePartition",objCDCreate.JavaWindow("Error"),"OK" )
					wait 2
				End If
			End If
			
			Call Fn_Button_Click("Fn_CPD_CreatePartition",objCDCreate,"Cancel" )
			If cBool(bOpenOnCreate) Then
				Call Fn_CPD_CompnentTabOperations("Close", "Content Search","")
				Call Fn_ReadyStatusSync(1)
    			Call Fn_CPD_CompnentTabOperations("Close", dicPartitionInfo("Name"),"")
    			Call Fn_ReadyStatusSync(1)
			End If
			
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_CPD_CreatePartition ] Invalid case [ " & sAction & " ].")
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
	End Select
	If  Fn_CPD_CreatePartition <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [ Fn_CPD_CreatePartition ] executed successfuly with case [ " & sAction & " ].")
	End If
	Set objCDCreate = Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_CPD_RevisionRuleOperations
'
'Description		:	Function Used to Create Partition
'
'Parameters			:  '1.sAction			: Action Name
'						2.sExistingRevRule	: Existing Revision Rule Name
'						3.sNewRevRule		: New Revision Rule to be selected.
'
'Return Value		: 	True Or False
'
'Pre-requisite		:	CPD perspective should be activated.
'
'Examples			:   Call Fn_CPD_RevisionRuleOperations("Set", "Any Status; No Working", "Any Status; Working")
'Examples			:   Call Fn_CPD_RevisionRuleOperations("Exist", "Any Status; No Working", "")
'
'History			:			
'			Developer Name				Date			Rev. No.	Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'			Koustubh Watwe			30-Jan-2012			1.0			Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'			Koustubh Watwe			28-Mar-2012			1.0			Added case Exist
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'			Swapnil Gore			14-DEC-2012			1.1			Added Generic code
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_CPD_RevisionRuleOperations(sAction, sExistingRevRule, sNewRevRule)
	GBL_FAILED_FUNCTION_NAME="Fn_CPD_RevisionRuleOperations"
	Dim objRevRule, aRevRule, iInstance,intNoOfObjects,iCnt,objSelectType,SIndex,SName,Flag

	Set objRevRule = JavaWindow("Collaborative Product").JavaObject("ImageHyperlink")

'Swapnil : 12-Sep-2012 : Added code to check the Working status:

	If   sAction = "Set" Then
						
						Flag = False
						
						Set objSelectType = Description.Create()
						objSelectType("Class Name").value = "JavaObject"
						objSelectType("toolkit class").value = "org.eclipse.ui.forms.widgets.ImageHyperlink"
						Set  intNoOfObjects = JavaWindow("Collaborative Product").ChildObjects(objSelectType)

						'Code to check if the New Revision rule is already set .
						
						For iCnt = 0 to intNoOfObjects.count-1
							 If Trim(intNoOfObjects(iCnt).GetROProperty("developer name")) = Trim(sNewRevRule) then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_CPD_RevisionRuleOperations is already SET")
									Fn_CPD_RevisionRuleOperations = True 
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
			Fn_CPD_RevisionRuleOperations = False
			sExistingRevRule = SName
			aRevRule = split(sExistingRevRule,"@")
			iInstance = SIndex
		Else
			Fn_CPD_RevisionRuleOperations = False
			aRevRule = split(sExistingRevRule,"@")
			iInstance = 0
		End IF
	
	End IF
	
	Fn_CPD_RevisionRuleOperations = False
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
	
				objRevRule.SetTOProperty "developer name", aRevRule(0)
				objRevRule.SetTOProperty "Index", iInstance
				
				objRevRule.Click 1,1,"LEFT"
				Wait 1
				'JavaWindow("Collaborative Product").JavaMenu("Label:=" & sNewRevRule).Select 
				Fn_CPD_RevisionRuleOperations = Fn_UI_JavaMenu_Select("Fn_CPD_RevisionRuleOperations",JavaWindow("Collaborative Product"),sNewRevRule)
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_CPD_RevisionRuleOperations : successfully selected revision rule [ " & sNewRevRule & " ] ")
				

		Case "Exist"
				objRevRule.SetTOProperty "developer name", aRevRule(0)
				objRevRule.SetTOProperty "Index", iInstance
				Fn_CPD_RevisionRuleOperations = objRevRule.Exist(5)
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_CPD_RevisionRuleOperations : successfully verified revision rule [ " & sNewRevRule & " ] ")
			
		Case "ExistInMenu"
				objRevRule.SetTOProperty "developer name", aRevRule(0)
				objRevRule.SetTOProperty "Index", iInstance
				
				objRevRule.Click 1,1,"LEFT"
				Wait 1
				
				'To resolve issues when QTP treats "(" and ")" as Special characters, thereby failing to work with Rules containing "(" or ")".
				If Instr( sNewRevRule , "(") <> 0 Or Instr( sNewRevRule , ")") <> 0 Then
					sNewRevRule = Replace(sNewRevRule , "(" , "\(")
					sNewRevRule = Replace(sNewRevRule , ")" , "\)")
				End If
	
				Fn_CPD_RevisionRuleOperations = JavaWindow("Collaborative Product").JavaMenu("Label:=" & sNewRevRule).Exist(5)
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_CPD_RevisionRuleOperations : successfully verified existence of revision rule [ " & sNewRevRule & " ] ")
				Call Fn_KeyBoardOperation("SendKeys", "{ESC}")
			
	End Select
	
	If Fn_CPD_RevisionRuleOperations <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_CPD_RevisionRuleOperations : Executed successfully with Case [ " & sAction & " ] ")
	End IF
	JavaWindow("Collaborative Product").JavaObject("ImageHyperlink").SetTOProperty "developer name", ""
	JavaWindow("Collaborative Product").JavaObject("ImageHyperlink").SetTOProperty "Index",0	
	Set objRevRule = Nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name			:	Fn_CPD_SearchResultTreeOperations
'@@
'@@    Description				:	Function Used to perform operations on Content Explorer
'@@
'@@    Parameters			    :	1. sAction		: Action to be performed
'@@								:	2. sNode		: Node Name
'@@								:	3. sColumn		: Column Name
'@@								:	4. sValue 		: for future use
'@@								:	5. sPopupMenu	: Popup menu
'@@
'@@    Return Value		   	   	: 	True Or False
'@@
'@@    Pre-requisite			:	CPD perspective should be activated.						
'@@
'@@    Examples					:	Call Fn_CPD_SearchResultTreeOperations("Select", "Search Results (5 found):DE000023/001;1-ddd", "", "", "")
'@@    Examples					:	Call Fn_CPD_SearchResultTreeOperations("DoubleClick", "Search Results (5 found):DE000023/001;1-ddd", "", "", "")
'@@    Examples					:	Call Fn_CPD_SearchResultTreeOperations("Exist", "Search Results (5 found):DE000023/001;1-ddd", "", "", "")
'@@    Examples					:	Call Fn_CPD_SearchResultTreeOperations("DeSelect", "Search Results (5 found):DE000023/001;1-ddd", "", "", "")
'@@    Examples					:	Call Fn_CPD_SearchResultTreeOperations("MultiSelect", "CD000015;1-CD1~CD000017;1-CD2", "", "", "")
'@@    Examples					:	Call Fn_CPD_SearchResultTreeOperations("Expand", "CD000015;1-CD1", "", "", "")
'@@    Examples					:	Call Fn_CPD_SearchResultTreeOperations("PopupMenuSelect", "CD000015;1-CD1", "", "", "Copy	Ctrl+C")
'@@    Examples					:	Call Fn_CPD_SearchResultTreeOperations("FindCollaborativeDesign", "CD000015;1-CD1", "", "", "")
'@@    Examples					:	Call Fn_CPD_SearchResultTreeOperations("CellVerify", "CD000015;1-CD1", "Type", "Collaborative Design", "")
'@@    Examples					:	Call Fn_CPD_SearchResultTreeOperations("MultiSelectPopupMenuSelect", "CD000015;1-CD1~CD000017;1-CD2", "", "", "Copy	Ctrl+C")
'@@    Examples					:	Call Fn_CPD_SearchResultTreeOperations("getfullnodebynodename_ext","DE1","","","") 
'@@
'@@	   History					:	
'@@				Developer Name				Date			Rev. No.	Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe			1-Feb-2012			1.0			Created
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe			9-Mar-2012			1.0			Modifeid code to get Node path, added code to set Top node in node path
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Amit T.					31-Jul-2012			1.0			Added cases "exist_basedonsourceobjectname", "select_basedonsourceobjectname" , existwithdename_basedonsourceobjectname
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe			03-Jul-2012			1.0			Added case DoubleClick
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Vivek Ahirrao			14-Jul-2015			1.0			Case "select_basedonsourceobjectname" Modified case to select the node based on Column name (Source Object Name) [Tc11.2 Porting : Build(2015062400) : By Vivek Ahirrao]
'@@																		Case "existwithdename_basedonsourceobjectname" Modified case to check the node and return it based on Column name (Source Object Name) [Tc11.2 Porting : Build(2015062400) : By Vivek Ahirrao]
'@@																		Case "exist_basedonsourceobjectname" Modified case to check existance of node based on Column name (Source Object Name) [Tc11.2 Porting : Build(2015062400) : By Vivek Ahirrao]
'@@																		Case "exist_countintree" Added to check total count of child nodes in tree [TC11.2 Maintenence : Build(2015062400) : By Vivek Ahirrao]
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Shweta Rathod			05-Nov-2015			1.0			Added case "getchilditems"														[TC1121-2015101900-19_11_2015-VivekA-NewDevelopment]
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Ankit Nigam				02-Dec-2015			1.0			Added case "getfullnodebynodename","getfullnodebynodename_ext""					[TC1121-2015110900-02_12_2015-AnkitN-NewDevelopment]
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Vivek Ahirrao			28-Dec-2015			1.0			Added case "Select_TopNode","verifyselectednode"								[TC1122-20151116d00-28_12_2015-AnkitN-NewDevelopment]
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_CPD_SearchResultTreeOperations(sAction, strNode, sColumn, sValue, sPopupMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_CPD_SearchResultTreeOperations"
	Dim sRetNodePath, aMenuList, iCnt, aNode , sTreeVal
	Dim objTree, sNode, sPath
	Dim iCount, strTopNode
	Dim oCurrentNode , arrStrNode , appNodeName , bFlag , iNodeItemsCount
	sNode = strNode
	Fn_CPD_SearchResultTreeOperations = False
	sRetNodePath = False
	Set objTree = JavaWindow("Collaborative Product").JavaTree("SearchResultTree")
'		If Not objTree.Exist(2) Then
'		  Set objTree=JavaWindow("Collaborative Product").JavaTree("SearchResultContentExplorerTree") 
'		   If Not objTree.Exist(2) Then
'		   	 Exit Function
'		   End If 
'		End If
		
		If Not objTree.Exist(2) Then     ' Modified by Chaitali R.
		  Set objTree=JavaWindow("Collaborative Product").JavaTree("NavTree")
		   If Not objTree.Exist(2) Then
		   	 Exit Function
		   End If 
		End If
		
	If Instr( sAction , "BasedOnSourceObjectName" ) = 0 And Instr( sAction , "basedonsourceobjectname") = 0 Then
		IF sNode <> "" Then
			If sAction<>"Select_TopNode" Then
				aNode = Split(sNode,":")
				If instr(lcase(aNode(0)), "search result") > 0 Then
					aNode(0) = objTree.getItem(0)
					For iCnt = 0 to uBound(aNode)
							If iCnt = 0 then
								sNode = aNode(0)
							Else
								sNode = sNode & ":" & aNode(iCnt)
							End IF
					Next
				Else
					sNode = objTree.getItem(0) & ":" & sNode
				End If
			End If
	   End If
	End If

	Select Case lcase(sAction)
		'[TC1122-20151116b-28_12_2015-VivekA-NewDevelopment] - Added to verify selected node in Nav Tree
        Case "verifyselectednode"
                If objTree.Object.getSelectionCount() <> 0 Then 
                    Set oCurrentNode = objTree.Object.getFocusItem()
    
                    If IsObject(oCurrentNode) then
                        sVerifyNode = oCurrentNode.getData().toString() 
                                            
                        Do while IsObject(oCurrentNode.getParentItem())
                            Set oCurrentNode = oCurrentNode.getParentItem()
                            sVerifyNode = oCurrentNode.getData().toString() & ":" & sVerifyNode 
                        Loop
                    End If
                    
                    If Trim(strNode) = Trim(sVerifyNode) Then
                        Fn_CPD_SearchResultTreeOperations = True
                    Else
                        Fn_CPD_SearchResultTreeOperations = False
                    End If
                    Set oCurrentNode = Nothing    
                Else
                    Fn_CPD_SearchResultTreeOperations = False
                End If 
	
		Case "select", "select_topnode"
				If lcase(sAction)="select_topnode" Then
					objTree.Select "#0"
					Fn_CPD_SearchResultTreeOperations = True
				Else
					sRetNodePath = Fn_UI_JavaTreeGetItemPathExt("Fn_CPD_SearchResultTreeOperations", objTree, sNode, "", "")
					If sRetNodePath <> False Then
						objTree.select sRetNodePath
						Fn_CPD_SearchResultTreeOperations = True
					Else
						Fn_CPD_SearchResultTreeOperations = Fn_JavaTree_Select("Fn_CPD_SearchResultTreeOperations", JavaWindow("Collaborative Product"),"SearchResultTree", sNode)
					End If
				End If
	
		Case "doubleclick"
				sRetNodePath = Fn_UI_JavaTreeGetItemPathExt("Fn_CPD_SearchResultTreeOperations", objTree, sNode, "", "")
				If sRetNodePath <> False Then
'					objTree.Activate sRetNodePath
					objTree.select sRetNodePath
					wait 1
                    Call Fn_KeyBoardOperation("SendKeys", "{ENTER}")
					Fn_CPD_SearchResultTreeOperations = True
				Else
					Fn_CPD_SearchResultTreeOperations = Fn_JavaTree_Select("Fn_CPD_SearchResultTreeOperations", JavaWindow("Collaborative Product"),"SearchResultTree", sNode)
				End If
	
		Case "exist","exists"
					Fn_CPD_SearchResultTreeOperations = False
					sRetNodePath = Fn_UI_JavaTree_NodeExist("Fn_CPD_SearchResultTreeOperations", objTree, sNode)
					If sRetNodePath = False Then
							sRetNodePath = Fn_UI_JavaTreeGetItemPathExt("Fn_CPD_SearchResultTreeOperations", objTree, sNode, "", "")
							If sRetNodePath <> False Then
								Fn_CPD_SearchResultTreeOperations = True
							End If
					Else
						Fn_CPD_SearchResultTreeOperations = True
					End If
					
		Case "deselect"
				sRetNodePath = Fn_UI_JavaTreeGetItemPathExt("Fn_CPD_SearchResultTreeOperations", objTree, sNode, "", "")
				If sRetNodePath <> False Then
					objTree.Deselect sRetNodePath
					Fn_CPD_SearchResultTreeOperations = True
				End If
		Case "multiselect"
				aNodes = split(sNode,"~")
				For iCnt = 0 to UBound(aNodes)
					sRetNodePath = Fn_JavaTree_NodeIndexExt("Fn_CPD_SearchResultTreeOperations", JavaWindow("Collaborative Product"), "SearchResultTree",  aNodes(iCnt), "", "")
					If sRetNodePath <> False Then
						objTree.ExtendSelect aNodes(iCnt)
						Fn_CPD_SearchResultTreeOperations = True
					Else
						Fn_CPD_SearchResultTreeOperations = False
						Exit For
					End If
				Next
		Case "cellverify"
				sRetNodePath = Fn_UI_JavaTreeGetItemPathExt("Fn_CPD_SearchResultTreeOperations", objTree, sNode, "", "")
				If sRetNodePath <> False Then
					If trim(objTree.GetColumnValue(sRetNodePath,sColumn)) = sValue Then
						Fn_CPD_SearchResultTreeOperations = True
					End If
				End If
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "expand"
				sRetNodePath = Fn_UI_JavaTreeGetItemPathExt("Fn_CPD_SearchResultTreeOperations", objTree, sNode, "", "")
				If sRetNodePath <> False Then
					objTree.Expand sRetNodePath
					Fn_CPD_SearchResultTreeOperations = True
				End If
		Case "popupmenuselect"
					sRetNodePath = Fn_UI_JavaTreeGetItemPathExt("Fn_CPD_SearchResultTreeOperations", objTree, sNode, "", "")
					If sRetNodePath = False Then
						sRetNodePath = sNode
					End IF
					objTree.Select sRetNodePath
					wait 1
					objTree.OpenContextMenu sRetNodePath
					wait 1
                    aMenuList = split(sPopupMenu,":")
					'Select Menu action
					Select Case Ubound(aMenuList)
						Case "0"
							 sPopupMenu = JavaWindow("Collaborative Product").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0))
						Case "1"
							sPopupMenu = JavaWindow("Collaborative Product").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1))
						Case "2"
							sPopupMenu = JavaWindow("Collaborative Product").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1),aMenuList(2))
						Case Else
							Fn_CPD_SearchResultTreeOperations = FALSE
							Exit Function
					End Select
					JavaWindow("Collaborative Product").WinMenu("ContextMenu").Select sPopupMenu
					Fn_CPD_SearchResultTreeOperations = True
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "multiselectpopupmenuselect"
				aNodes = split(sNode,"~")
				For iCnt = 0 to UBound(aNodes) -1
					sRetNodePath = Fn_JavaTree_NodeIndexExt("Fn_CPD_SearchResultTreeOperations", JavaWindow("Collaborative Product"), "SearchResultTree",  aNodes(iCnt), "", "")
					If sRetNodePath <> False Then
						objTree.ExtendSelect aNodes(iCnt)
						Fn_CPD_SearchResultTreeOperations = True
					Else
						Fn_CPD_SearchResultTreeOperations = False
						Exit Function
					End If
				Next
				sRetNodePath = Fn_JavaTree_NodeIndexExt("Fn_CPD_SearchResultTreeOperations", JavaWindow("Collaborative Product"), "SearchResultTree",  aNodes(iCnt), "", "")
				If sRetNodePath <> False Then
					objTree.ExtendSelect aNodes(iCnt)
					wait 1
					objTree.OpenContextMenu aNodes(iCnt)
					wait 1
                    aMenuList = split(sPopupMenu,":")
					'Select Menu action
					Select Case Ubound(aMenuList)
						Case "0"
							 sPopupMenu = JavaWindow("Collaborative Product").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0))
						Case "1"
							sPopupMenu = JavaWindow("Collaborative Product").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1))
						Case "2"
							sPopupMenu = JavaWindow("Collaborative Product").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1),aMenuList(2))
						Case Else
							Fn_CPD_SearchResultTreeOperations = FALSE
							Exit Function
					End Select
					JavaWindow("Collaborative Product").WinMenu("ContextMenu").Select sPopupMenu
					Fn_CPD_SearchResultTreeOperations = True
				End If
				
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "exist_basedonsourceobjectname"  ' Modified case to check existance of node based on Column name (Source Object Name) [Tc11.2 Porting : Build(2015062400) : By Vivek Ahirrao]
				For iCnt = 0 to cint(objTree.GetROProperty("items count")) - 1
					' set node path	
					sPath = "#0:#"& iCnt
					' get value of node based on Source Object Name
		            sTreeVal = objTree.GetColumnValue(sPath,"Source Object Name")
					If Trim(sTreeVal) = Trim(sNode) Then
				         Fn_CPD_SearchResultTreeOperations = True
				         Exit For
					End If
				Next
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "select_basedonsourceobjectname" ' Modified case to select the node based on Column name (Source Object Name) [Tc11.2 Porting : Build(2015062400) : By Vivek Ahirrao]
				For iCnt = 0 to cint(objTree.Object.getItem(0).getItemCount()) - 1
					' set node path				
					sRetNodePath = "#0:#"& iCnt
					' get value of node based on Source Object Name
					sTreeVal = objTree.GetColumnValue(sRetNodePath,"Source Object Name")
					If sNode = sTreeVal Then
						'Select node at this index
						If sRetNodePath <> "" Then
							objTree.Select sRetNodePath
							wait 1
							Fn_CPD_SearchResultTreeOperations = True
							Exit For
						End If
					End If
				Next
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "getchildrencount"
				Fn_CPD_SearchResultTreeOperations = objTree.Object.getItem(0).getData().getChildrenCount()
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "getrootnodename"
				Fn_CPD_SearchResultTreeOperations = objTree.Object.getItem(0).getData().tostring()
				
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "existwithdename_basedonsourceobjectname" ' Modified case to check the node and return it based on Column name (Source Object Name) [Tc11.2 Porting : Build(2015062400) : By Vivek Ahirrao]
				For iCnt = 0 to cint(objTree.Object.getItem(0).getItemCount())-1
					' set node path	
					sPath = "#0:#"& iCnt
					' get value of node based on Source Object Name
		            sTreeVal = objTree.GetColumnValue(sPath,"Source Object Name")
					If Trim(sTreeVal) = Trim(sNode) Then
				         Fn_CPD_SearchResultTreeOperations = objTree.Object.getItem(0).getItem(iCnt).getdata().tostring()
				         Exit For
					End If
				Next
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "exist_countintree"	' Added new Case "exist_countintree" [ Maintenence : TC11.2 : Build(2015062400) : 09/July/2015] by Vivek Ahirrao
				Fn_CPD_SearchResultTreeOperations = False
				iCount = objTree.Object.getItem(0).getItemCount()
				strTopNode = objTree.Object.getTopItem().getdata().tostring()
				If Instr(1,strNode,strTopNode) > 0 AND Instr(1,strNode,iCount) > 0 Then
					Fn_CPD_SearchResultTreeOperations = True
				End If
		Case "getchilditems"	'[TC1121-2015101900-19_11_2015-VivekA-NewDevelopment] - Added by Shweta R
				For iCnt = 0 to cint(objTree.Object.getItem(0).getData().getChildrenCount())-1
					If iCnt = 0 then
						Fn_CPD_SearchResultTreeOperations = objTree.Object.getItem(0).getItem(iCnt).getData().tostring()
					Else
						Fn_CPD_SearchResultTreeOperations = Fn_CPD_SearchResultTreeOperations & "~" & objTree.Object.getItem(0).getItem(iCnt).getData().tostring()
					End IF
				next
		Case "getfullnodebynodename","getfullnodebynodename_ext"       '[TC1121-2015110900-02_12_2015-AnkitN-NewDevelopment] - Added by Ankit N.
				Fn_CPD_SearchResultTreeOperations=False
				arrStrNode=SPlit(strNode,":")
	        	Set oCurrentNode = objTree.Object.getItem(0)
				For iCnt = 0 to UBound(arrStrNode)
					bFlag=False
					iNodeItemsCount = oCurrentNode.getItemCount()
	        		For iCount = 0 to iNodeItemsCount - 1
	        			If sAction = "getfullnodebynodename_ext" Then
	    					appNodeName = oCurrentNode.getItem(iCount).getData().toString()
	    				Else
	    					appNodeName = oCurrentNode.getItem(iCount).getText()
	    				End If
						If UBound(arrStrNode)=iCnt Then
							If instr(1,Trim(appNodeName), Trim(arrStrNode(iCnt))) Then
								Fn_CPD_SearchResultTreeOperations=Trim(appNodeName)
								bFlag=True
								Exit For
							End If
						Else
							If Trim(appNodeName) = Trim(arrStrNode(iCnt)) Then
								Set oCurrentNode = oCurrentNode.getItem(iCount)
								bFlag=True
								Exit For
							End If
						End If
					Next
					If bFlag=False Then
						Exit For
					End If
				Next 
				Set oCurrentNode=Nothing				
		Case Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_CPD_SearchResultTreeOperations ] Invalid case [ " & sAction & " ].")
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
	End Select
	If  Fn_CPD_SearchResultTreeOperations <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [ Fn_CPD_SearchResultTreeOperations ] executed successfuly with case [ " & sAction & " ].")
	End If
	Set objTree = Nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name			:	Fn_CPD_SummaryTabOperations
'@@
'@@    Description				:	Function Used to perform operations on summary tab
'@@
'@@    Parameters			    :	1. sAction			: Action to be performed
'@@								:	2. sSummaryField	: Summary field label
'@@								:	3. sSummaryValue	: Value to be verified
'@@								:	4. sPopupMenu 		: for future use.
'@@
'@@    Return Value		   	   	: 	True / Tab Name Or False
'@@
'@@    Pre-requisite			:	CPD perspective should be activated.						
'@@
'@@    Examples					:	Call Fn_CPD_SummaryTabOperations("Verify", "Name~Type", "cd1~Collaborative Design", "")
'@@    Examples					:	Call Fn_CPD_SummaryTabOperations("Verify", "WherePartitioned", "PTN000172/001;1-pt1", "")
'@@    Examples					:	Call Fn_CPD_SummaryTabOperations("SelectToobarButton", "", "Cut", "")
'@@    Examples					:	Call Fn_CPD_SummaryTabOperations("SelectTab", "", "Overview", "")
'@@    Examples					:	Call Fn_CPD_SummaryTabOperations("GetSelectedTabName", "", "", "")
'@@	   Examples					:	Call Fn_CPD_SummaryTabOperations("ToobarButtonExist", "","Add New...", "")
'@@	   Examples					:	Call Fn_CPD_SummaryTabOperations("VerifyInRelatedAttributeGroups", "a1","Type=Subset Defaults", "")
'@@	   Examples					:	Call Fn_CPD_SummaryTabOperations("SelectInRelatedAttributeGroups", "a1","", "")
'@@	   History					:	
'@@				Developer Name				Date			Rev. No.	Changes Done																								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe			06-Feb-2012			1.0			Created
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe			15-Feb-2012			1.0			Added case VerifyHeaderFields
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe			15-Feb-2012			1.0			Added case SelectToobarButton, SelectTab, GetSelectedTabName
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe			15-Feb-2012			1.0			Modifeid case Verify added case for WherePartitioned
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe			18-Apr-2012			1.0			Added case EditFields
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Pallavi Patil			03-May-2012			1.0			Added case ToobarButtonExist
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe			23-Aug-2012			1.0			Added cases VerifyInRelatedAttributeGroups, SelectInRelatedAttributeGroups
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Vivek Ahirrao			24-Dec-2014			1.0			Added cases PopupMenuSelectInRelatedAttributeGroups											[TC1122-20151116d-24_12_2015-VivekA-NewDevelopment]
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_CPD_SummaryTabOperations(sAction, sSummaryField, sSummaryValue, sPopupMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_CPD_SummaryTabOperations"
	Dim aFields, aValues, objApplet, sPropertyName
	Dim iCnt, sActVal,iItmCount
	Dim iCount, iInnerCnt, bFlag
	Fn_CPD_SummaryTabOperations = False
	Set objApplet = JavaWindow("Collaborative Product")
	Fn_CPD_SummaryTabOperations = False

	If sAction = "VerifyInRelatedAttributeGroups" Then
		' do something
		aValues = split(sSummaryValue,"=")
		Select Case aValues(0) 
			Case "Object"
				sPropertyName = "object_string"
			Case "Type"
				sPropertyName = "object_type"
			Case "Last Modified Date"
				sPropertyName = "last_mod_date"
			Case "Date Modified"
				sPropertyName = "last_mod_date"
			Case "Last Modifying User"
				sPropertyName = "ast_mod_user"
			Case "Checked-Out By"
				sPropertyName = "checked_out_user"
			Case "Viewed Partition Scheme"
				sPropertyName = "ptn0viewed_partition_scheme"
		End Select
	End If

	Select Case sAction
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "EditFields"
			aFields = Split(sSummaryField, "~")
			aValues = Split(sSummaryValue, "~")
			Fn_CPD_SummaryTabOperations = True
			objApplet.JavaStaticText("Summary_PropertyName").SetTOProperty  "Index", 1
			
			For iCnt = 0 to uBound(aFields)
				objApplet.JavaStaticText("Summary_PropertyName").SetTOProperty "label", aFields(iCnt) & ":"
				If objApplet.JavaEdit("Summary_PropertyValue").Exist(3) Then
					objApplet.JavaEdit("Summary_PropertyValue").Set trim(aValues(iCnt))
					Fn_CPD_SummaryTabOperations = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [ Fn_CPD_SummaryTabOperations ] Successfully set [ " & aFields(iCnt) & " = " & aValues(iCnt) & " ].")
				Else
					Fn_CPD_SummaryTabOperations = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed: [ Fn_CPD_SummaryTabOperations ] Failed to set [ " & aFields(iCnt) & " = " & aValues(iCnt) & " ].")	
				End If
			Next
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "Verify"
			aFields = Split(sSummaryField, "~")
			aValues = Split(sSummaryValue, "~")
			Fn_CPD_SummaryTabOperations = True
			objApplet.JavaStaticText("Summary_PropertyName").SetTOProperty  "Index", 1
			
			For iCnt = 0 to uBound(aFields)
				Select Case aFields(iCnt)
					Case "WherePartitioned"
						bFlag = False
						'msgbox JavaWindow("Collaborative Product").JavaObject("Gallery").Object.getItem(0).getitem(0).getText()
						If objApplet.JavaObject("Summary_WherePartitionedGallery").Exist(5) = False Then
							objApplet.JavaStaticText("Summary_WherePartitioned_label").Click 1, 1,"LEFT"
							wait(2)
						End If
						If objApplet.JavaObject("Summary_WherePartitionedGallery").Exist(5) Then
							For iCount = 0 to objApplet.JavaObject("Summary_WherePartitionedGallery").Object.getItemCount() -1
								For iInnerCnt = 0 to objApplet.JavaObject("Summary_WherePartitionedGallery").Object.getItem(iCount).getItemCount() -1
									If objApplet.JavaObject("Summary_WherePartitionedGallery").Object.getItem(iCount).getitem(iInnerCnt).getText() = aValues(iCnt) Then
										bFlag = True
										Exit for
									End If
								Next
								If bFlag Then
									Exit for
								End If
							Next
						End If

						If bFlag = False Then
							Fn_CPD_SummaryTabOperations = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_CPD_SummaryTabOperations ] to verify [ " & aFields(iCnt) & " = " & aValues(iCnt) & " ].")
							Exit for
						End If

					Case Else
						objApplet.JavaStaticText("Summary_PropertyName").SetTOProperty "label", aFields(iCnt) & ":"
						If objApplet.JavaEdit("Summary_PropertyValue").Exist(3) Then
							sActVal = objApplet.JavaEdit("Summary_PropertyValue").GetROProperty("value")
							If trim(sActVal) <> trim(aValues(iCnt)) Then
								Fn_CPD_SummaryTabOperations = False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_CPD_SummaryTabOperations ] to verify [ " & aFields(iCnt) & " = " & aValues(iCnt) & " ].")
								Exit for
							End If
						ElseIf objApplet.JavaObject("Summary_PropertyValueHyperlink").Exist(5) Then
							sActVal = objApplet.JavaObject("Summary_PropertyValueHyperlink").Object.getText() 
							If trim(sActVal) <> trim(aValues(iCnt)) Then
								Fn_CPD_SummaryTabOperations = False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_CPD_SummaryTabOperations ] to verify [ " & aFields(iCnt) & " = " & aValues(iCnt) & " ].")
								Exit for
							End If
						Else
							Fn_CPD_SummaryTabOperations = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_CPD_SummaryTabOperations ] to verify [ " & aFields(iCnt) & " = " & aValues(iCnt) & " ].")
							Exit for
						End If
				End Select
			Next
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "VerifyHeaderFields"
			aFields = Split(sSummaryField, "~")
			aValues = Split(sSummaryValue, "~")
			Fn_CPD_SummaryTabOperations = True
			
			objApplet.JavaStaticText("Summary_PropertyName").SetTOProperty  "Index", 0
			For iCnt = 0 to uBound(aFields)
				'[TC1123-20160518-27_05_2016-VivekA-Maintenance] - Added as per Design change, and dicussed with Akshay J
				If aFields(iCnt) = "Last Modified Date" Then
					objApplet.JavaStaticText("Summary_PropertyName").SetTOProperty "label", aFields(iCnt) & ":"
					If objApplet.JavaStaticText("Summary_PropertyName").Exist = False Then
						objApplet.JavaStaticText("Summary_PropertyName").SetTOProperty "label", "Date Modified" & ":"
					End If
				Else
					objApplet.JavaStaticText("Summary_PropertyName").SetTOProperty "label", aFields(iCnt) & ":"
				End If
				'-------------------------------------------------------------
				'objApplet.JavaStaticText("Summary_PropertyName").SetTOProperty "label", aFields(iCnt) & ":"
				If objApplet.JavaEdit("SummaryHeader_PropertyValue").Exist(3) Then
					sActVal = objApplet.JavaEdit("SummaryHeader_PropertyValue").GetROProperty("value")
				ElseIf objApplet.JavaObject("SummaryHeader_PropertyValueHyperlink").Exist(5) Then
					sActVal = objApplet.JavaObject("SummaryHeader_PropertyValueHyperlink").Object.getText() 
				Else
					Fn_CPD_SummaryTabOperations = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_CPD_SummaryTabOperations ] to verify [ " & aFields(iCnt) & " = " & aValues(iCnt) & " ].")
					Exit for
				End If
				Select case aFields(iCnt)
						Case "Last Modified Date"
							If trim(left(sActVal,instr(sActVal," "))) <> trim(left(aValues(iCnt),instr(aValues(iCnt)," "))) Then
								Fn_CPD_SummaryTabOperations = False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_CPD_SummaryTabOperations ] to verify [ " & aFields(iCnt) & " = " & aValues(iCnt) & " ].")
								Exit for
							End If
						Case Else
							If trim(sActVal) <> trim(aValues(iCnt)) Then
								Fn_CPD_SummaryTabOperations = False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_CPD_SummaryTabOperations ] to verify [ " & aFields(iCnt) & " = " & aValues(iCnt) & " ].")
								Exit for
							End If
					End Select
			Next
			objApplet.JavaStaticText("Summary_PropertyName").SetTOProperty  "Index", 1
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "SelectToobarButton"
			'Dim iItmCount
			objApplet.JavaButton("SummaryButton").SetTOProperty "label", sSummaryValue			
			If objApplet.JavaButton("SummaryButton").Exist(2) Then
				objApplet.JavaButton("SummaryButton").Click micLeftBtn
				Fn_CPD_SummaryTabOperations = TRUE
			End If
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "ToobarButtonExist"
			iItmCount = cInt(objApplet.JavaToolbar("SummaryTabToolbar").GetROProperty("toolbar items"))
			For iCnt = 1 to iItmCount 
				If lcase(sSummaryValue) = lcase(objApplet.JavaToolbar("SummaryTabToolbar").GetItemProperty(iCnt, "name")) Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verified Toolbar Button " & sSummaryValue)
					Fn_CPD_SummaryTabOperations = TRUE
					Exit For
				End If
			Next
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "SelectTab"
			objApplet.JavaTab("SummaryTab").Select sSummaryValue
			Fn_CPD_SummaryTabOperations = True
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "GetSelectedTabName"
			Fn_CPD_SummaryTabOperations = JavaWindow("Collaborative Product").JavaTab("SummaryTab").GetROProperty("value")
	' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "VerifyInRelatedAttributeGroups"
			objApplet.JavaTab("SummaryTab").Select "Related Attribute Groups"

			iItmCount = cInt(objApplet.JavaTable("RelatedDesignElement").GetROProperty("rows"))
			For iCount = 0 to iItmCount - 1
				If objApplet.JavaTable("RelatedDesignElement").Object.getItem(iCount).getData().toString() = sSummaryField Then
					If objApplet.JavaTable("RelatedDesignElement").Object.getItem(iCount).getData().getComponent().getProperty(sPropertyName) = aValues(1) Then
						Fn_CPD_SummaryTabOperations = TRUE
					End IF
					Exit for
				End If
			Next

	' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "SelectInRelatedAttributeGroups"
			objApplet.JavaTab("SummaryTab").Select "Related Attribute Groups"
			iItmCount = cInt(objApplet.JavaTable("RelatedDesignElement").GetROProperty("rows"))
			aFields = Split(sSummaryValue, "~")
			For iCount = 0 to iItmCount - 1
				If objApplet.JavaTable("RelatedDesignElement").Object.getItem(iCount).getData().toString() = sSummaryField Then
					objApplet.JavaTable("RelatedDesignElement").DeselectRow iCount
					wait 1
					objApplet.JavaTable("RelatedDesignElement").ClickCell iCount, 0
					Fn_CPD_SummaryTabOperations = TRUE
					Exit for
				End If
			Next
		'[TC1122-20151116d-18_12_2015-VivekA-NewDevelopment] - Added to Select Popup menu on hyperlink property value
		Case "PropertyValueMenuSelect"
			objApplet.JavaTab("SummaryTab").Select "Overview"
			objApplet.JavaStaticText("Summary_PropertyName").SetTOProperty "label",sSummaryField
			bFlag = Fn_UI_JavaStaticText_Click("Fn_CPD_SummaryTabOperations",objApplet,"Summary_PropertyValue",1,1,"LEFT")
			If bFlag <> True Then
				Fn_CPD_SummaryTabOperations = False
				Set objApplet = Nothing
				Exit Function
			End If
			'bFlag = Fn_UI_JavaMenu_Select("Fn_CPD_SummaryTabOperations",JavaWindow("DefaultWindow"),sPopupMenu)
			bFlag = Fn_UI_JavaMenu_Select("Fn_CPD_SummaryTabOperations",objApplet,sPopupMenu)
			If bFlag <> True Then
				Fn_CPD_SummaryTabOperations = False
				Set objApplet = Nothing
				Exit Function
			Else
				Fn_CPD_SummaryTabOperations = True
				Set objApplet = Nothing
			End If
		'[TC1122-20151116d-18_12_2015-VivekA-NewDevelopment] - Added to get Hyperlink property value
		Case "GetPropertyValueHyperlink"
			objApplet.JavaTab("SummaryTab").Select "Overview"
			objApplet.JavaStaticText("Summary_PropertyName").SetTOProperty "label",sSummaryField
			Wait 1
			bFlag = Fn_UI_Object_GetROProperty("",objApplet.JavaObject("Summary_PropertyValueHyperlink"),"text")
			If bFlag = False Then
				Fn_CPD_SummaryTabOperations = False
				Set objApplet = Nothing
				Exit Function
			Else
				Fn_CPD_SummaryTabOperations = bFlag
				Set objApplet = Nothing
			End If
		'bReturn = Fn_CPD_SummaryTabOperations("PopupMenuSelectInRelatedAttributeGroups", "Object", "Test", "View Properties	Alt+P")
		'[TC1122-20151116d-24_12_2015-VivekA-NewDevelopment] - Added to Select Popup menu on Related Attribute Groups
		Case "PopupMenuSelectInRelatedAttributeGroups"
			objApplet.JavaTab("SummaryTab").Select "Related Attribute Groups"
			iItmCount = cInt(objApplet.JavaTable("RelatedDesignElement").GetROProperty("rows"))
			For iCount = 0 to iItmCount - 1
				If Trim(objApplet.JavaTable("RelatedDesignElement").Object.getItem(iCount).getData().toString()) = Trim(sSummaryValue) Then
					objApplet.JavaTable("RelatedDesignElement").DeselectRow iCount
					Wait 1
					objApplet.JavaTable("RelatedDesignElement").ClickCell iCount,sSummaryField,"RIGHT"
					Wait 1
					sContents = JavaWindow("Collaborative Product").WinMenu("ContextMenu").BuildMenuPath(sPopupMenu)
					objApplet.WinMenu("ContextMenu").Select sContents
					Fn_CPD_SummaryTabOperations = TRUE
					Exit for
				End If
			Next
			Set objApplet = Nothing			
	' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_CPD_SummaryTabOperations ] Invalid case [ " & sAction & " ].")
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
	End Select
	If  Fn_CPD_SummaryTabOperations <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [ Fn_CPD_SummaryTabOperations ] executed successfuly with case [ " & sAction & " ].")
	End If
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name			:	Fn_CPD_ColumnManagementOperation
'@@
'@@    Description				:	Function Used to perform operations on summary tab
'@@
'@@    Parameters			    :	1. sAction			: Action to be performed
'@@								:	2. sAailableCols
'@@								:	3. sDisplayedCols
'@@								:	4. bShowInternalNames
'@@								:	5. sTableConfigName
'@@								:	6. sTableConfigNewName
'@@								:	7. sTableConfigDescription
'@@								:	8. bCloseWIndow
'@@
'@@    Return Value		   	   	: 	True Or False
'@@
'@@    Pre-requisite			:	CPD perspective should be activated., Content Explorer should be opened.						
'@@
'@@    Examples					:	Call  Fn_CPD_ColumnManagementOperation("AddRemoveColumns", "Object~Type", "", True, "", "", "", "")
'@@    Examples					:	Call  Fn_CPD_ColumnManagementOperation("VerifyAvailableProp", "Type", "", "", "", "", "", "")
'@@    Examples					:	Call  Fn_CPD_ColumnManagementOperation("VerifyDisplayedProp", "", "Type", "", "", "", "", False)
'@@    Examples					:	Call  Fn_CPD_ColumnManagementOperation("MoveIndexUpDown", "", "Type:3", "", "", "", "", "")
'@@
'@@	   History					:	
'@@				Developer Name				Date			Rev. No.	Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe			07-Mar-2012			1.0			Created
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Archana Dhadiwal		23-Aug-2016			1.0			Created								[TC1123-20160729-25_08_2016-VivekA-NewDevelopment] - Added case "MoveIndexUpDown" for 4GD New Tc's
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_CPD_ColumnManagementOperation(sAction, sAailableCols, sDisplayedCols, bShowInternalNames, sTableConfigName, sTableConfigNewName, sTableConfigDescription, bCloseWIndow)
	GBL_FAILED_FUNCTION_NAME="Fn_CPD_ColumnManagementOperation"
	Dim objDialog, iItemCnt, iCnt, aCols, iColCnt, bFlag, aColumns, iColIndex
	Set objDialog = JavaWindow("MyTeamcenter").JavaWindow("Column Management")
	Fn_CPD_ColumnManagementOperation = False
	If Fn_UI_ObjectExist("Fn_CPD_ColumnManagementOperation",JavaWindow("MyTeamcenter").JavaWindow("Column Management")) = False Then
		If JavaWindow("Collaborative Product").JavaToolbar("CompositeTabVeiwMenu").Exist(5) then
			JavaWindow("Collaborative Product").JavaToolbar("CompositeTabVeiwMenu").Press "View Menu"
			wait 2
			sContents = JavaWindow("DefaultWindow").WinMenu("ContextMenu").BuildMenuPath("Column...")
			JavaWindow("DefaultWindow").WinMenu("ContextMenu").Select sContents
		Else
			Call Fn_ToolBarOperation("ShowDropdownAndSelectWithInstance", "View Menu:1", "Column...")
			wait 2
			If Fn_UI_ObjectExist("Fn_CPD_ColumnManagementOperation",JavaWindow("MyTeamcenter").JavaWindow("Column Management"))  = False  Then
				Call Fn_ToolBarOperation("ShowDropdownAndSelectWithInstance", "View Menu:2", "Column...")	
				wait 2
			End If
		End if
		If Fn_UI_ObjectExist("Fn_CPD_ColumnManagementOperation",JavaWindow("MyTeamcenter").JavaWindow("Column Management"))  = False Then
			exit function
		End If
	End If
	Call Fn_ReadyStatusSync(5)
	Select Case sAction
	' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "AddRemoveColumns", "AddRemoveColumnsWithType"
				bFlag = False
				If bShowInternalNames <> "" Then
					If CBool(bShowInternalNames) Then
						objDialog.JavaCheckBox("Show internal names").Set "ON"
					Else
						objDialog.JavaCheckBox("Show internal names").Set "OFF"
					End If
				End If
				If sAailableCols <> "" Then
					iCnt = cInt(objDialog.JavaTable("AvailableProp").Object.getSelectionIndex)
					If iCnt <> -1  Then
						objDialog.JavaTable("AvailableProp").DeselectRow iCnt
					End If
					iItemCnt = objDialog.JavaTable("AvailableProp").GetROProperty("rows") 
					aCols = split(sAailableCols,"~")
					If sAction = "AddRemoveColumns" Then
						' selection without type
						For iColCnt = 0 to UBound(aCols)
							For iCnt = 0 to iItemCnt -1
								If objDialog.JavaTable("AvailableProp").GetCellData(iCnt,0 ) = aCols(iColCnt) Then
									If iColCnt = 0 Then
										objDialog.JavaTable("AvailableProp").SelectCell iCnt,0
										wait 2
									Else
										objDialog.JavaTable("AvailableProp").ExtendRow iCnt
									End If
									bFlag = True
									Exit for
								End IF
							Next
						Next
					Else
						' selection with Type
						For iColCnt = 0 to UBound(aCols)
						aProperties = split(aCols(iColCnt),"$")
						For iCnt = 0 to iItemCnt -1
							If objDialog.JavaTable("AvailableProp").GetCellData(iCnt,0 ) = aProperties(0) Then
								If objDialog.JavaTable("AvailableProp").GetCellData(iCnt,1 ) = aProperties(1) Then
									If iColCnt = 0 Then
										objDialog.JavaTable("AvailableProp").SelectRow iCnt
									Else
										objDialog.JavaTable("AvailableProp").ExtendRow iCnt
									End If
									bFlag = True
									Exit for
								End IF
							End IF
						Next
					Next
					End IF
					If bFlag = True Then
						'Click on Add column Button
						Call Fn_Button_Click("Fn_CPD_ColumnManagementOperation", ObjDialog, "AddCol")
					End If
				End If
				If sDisplayedCols <> "" Then
					iCnt = cInt(objDialog.JavaTable("DisplayedColumns").Object.getSelectionIndex)
					If iCnt <> -1  Then
						objDialog.JavaTable("DisplayedColumns").DeselectRow iCnt
					End If
					iItemCnt = objDialog.JavaTable("DisplayedColumns").GetROProperty("rows") 
					aCols = split(sDisplayedCols,"~")
					For iColCnt = 0 to UBound(aCols)
						For iCnt = 0 to iItemCnt -1
							If objDialog.JavaTable("DisplayedColumns").GetCellData(iCnt,0 ) = aCols(iColCnt) Then
'								If iColCnt = 0 Then
									objDialog.JavaTable("DisplayedColumns").ActivateCell iCnt, 0
'								Else
'									objDialog.JavaTable("DisplayedColumns").ExtendRow iCnt
'								End If
								bFlag = True
								Exit for
							End IF
						Next
					Next
'					'Click on Remove column Button
'					If bFlag = True Then
'						Call Fn_Button_Click("Fn_CPD_ColumnManagementOperation", ObjDialog,  "RemoveCol")
'					End If
				End If
				If bFlag = True Then
					'Click on Apply Button
					Call Fn_Button_Click("Fn_CPD_ColumnManagementOperation", ObjDialog, "Apply")
				End If
				Fn_CPD_ColumnManagementOperation = bFlag 
	' - - - - - -  - - - - - - - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "VerifyAvailableProp"
					iItemCnt = objDialog.JavaTable("AvailableProp").GetROProperty("rows") 
					aCols = split(sAailableCols,"~")
					For iColCnt = 0 to UBound(aCols)
						Fn_CPD_ColumnManagementOperation = False 
						For iCnt = 0 to iItemCnt -1
							If objDialog.JavaTable("AvailableProp").GetCellData(iCnt,0 ) = aCols(iColCnt) Then
								Fn_CPD_ColumnManagementOperation = True
								Exit for
							End IF
						Next
						If Fn_CPD_ColumnManagementOperation = False Then
							Exit for
						End If
					Next
	' - - - - - -  - - - - - - - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "VerifyAvailablePropWithType"
					Dim aProperties
					iItemCnt = objDialog.JavaTable("AvailableProp").GetROProperty("rows") 
					aCols = split(sAailableCols,"~")
					For iColCnt = 0 to UBound(aCols)
						Fn_CPD_ColumnManagementOperation = False 
						aProperties = split(aCols(iColCnt),"$")
						For iCnt = 0 to iItemCnt -1
							If objDialog.JavaTable("AvailableProp").GetCellData(iCnt,0 ) = aProperties(0) Then
								If objDialog.JavaTable("AvailableProp").GetCellData(iCnt,1 ) = aProperties(1) Then
									Fn_CPD_ColumnManagementOperation = True
									Exit for
								End If
							End IF
						Next
						If Fn_CPD_ColumnManagementOperation = False Then
							Exit for
						End If
					Next
	' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "VerifyDisplayedProp"
					iItemCnt = objDialog.JavaTable("DisplayedColumns").GetROProperty("rows") 
					aCols = split(sDisplayedCols,"~")
					For iColCnt = 0 to UBound(aCols)
						Fn_CPD_ColumnManagementOperation = False
						For iCnt = 0 to iItemCnt -1
							If objDialog.JavaTable("DisplayedColumns").GetCellData(iCnt,0 ) = aCols(iColCnt) Then
								Fn_CPD_ColumnManagementOperation = True
								Exit for
							End IF
						Next
						If Fn_CPD_ColumnManagementOperation = False Then
							Exit for
						End If
					Next
		'Case to move Column Index Up or Down
		Case "MoveIndexUpDown"
				aColumns = Split(sDisplayedCols,":",-1,1)
				iItemCnt = objDialog.JavaTable("DisplayedColumns").GetROProperty("rows")
				For iCnt = 0 To iItemCnt-1
	 				If objDialog.JavaTable("DisplayedColumns").GetCellData(iCnt, 0) = aColumns(0) then
						objDialog.JavaTable("DisplayedColumns").SelectCell iCnt, 0
						Exit for
				 	End IF
				Next
				iColIndex = cInt(aColumns(1)) - cInt(iCnt)
				If iColIndex > 0 Then
					For iIndex = 0 To iColIndex -2	
						Call Fn_Button_Click("Fn_CPD_ColumnManagementOperation",objDialog,"MoveDown")
					Next
				ElseIf iColIndex < 0 Then
					iColIndex = iColIndex * ( -1 )
					For iIndex = 0 To iColIndex
						Call Fn_Button_Click("Fn_CPD_ColumnManagementOperation",objDialog,"MoveUp")
					Next
				End If
				' Hit  Apply Button 
				Call Fn_Button_Click("Fn_CPD_ColumnManagementOperation",objDialog,"Apply")
				Fn_CPD_ColumnManagementOperation = True
	' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_CPD_ColumnManagementOperation ] Invalid case [ " & sAction & " ].")
	' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
	End Select

	If bCloseWIndow = "" Then bCloseWIndow = True
	If cBool(bCloseWIndow) Then
		'Click on Close Button
		Call Fn_Button_Click("Fn_CPD_ColumnManagementOperation", ObjDialog, "Close")
	End If

	If  Fn_CPD_ColumnManagementOperation Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [ Fn_CPD_ColumnManagementOperation ] executed successfuly with case [ " & sAction & " ].")
	End If
End Function

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name			:	Fn_CPD_CreateDesignElementWhilePaste
'@@
'@@    Description				:	Use to handle paste Dialog to Paste DE in CD.
'@@
'@@    Parameters			    :	1. sAction		: Action to be performed
'@@								:	2. sTargetType		: Target Type
'@@								:	3. sCheckout : Boolean value to set Check out on Create or Not. valid values = True / False
'@@								:	4. sCopyEffectivity: Boolean value to Copy Effectivity On checkbox  True / False
'@@								:	5. sButton: Button Name.
'@@
'@@    Return Value		   	   	: 	True Or False
'@@
'@@    Pre-requisite			:	Create Design Element Dialog is Opened.
'@@
'@@    Examples					:	Call Fn_CPD_CreateDesignElementWhilePaste("Create","Design Element",True,"","Finish")
'@@    							:	Call Fn_CPD_CreateDesignElementWhilePaste("Create","TargetType:Design Element~SavedVariantRule:any value",True,"Applyeffectivitybasedontargetrevisionrule:True~Applyvariantconditionsbasedonsourceobject:False","Finish")
'@@
'@@	   History					:	
'@@	   
'@@		Developer Name			Date			Rev. No.	Changes Done													Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@		Sachin Joshi			15-March-2012	1.0			 Created
'@@		Veena Gurjar			04-Jan-2013						1. Modified Call Fn_Button_Click() for finish button		Koustubh watwe
'@@																2. Added Call Fn_Button_Click() for Next button
'@@		Vivek Ahirrao			17-Dec-2015		1.1			 Modified case "create" 										[TC1122-20151116d-17_12_2015-VivekA-NewDevelopment]
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_CPD_CreateDesignElementWhilePaste(sAction,sTargetType,sCheckout,sCopyEffectivity,sButton)
	GBL_FAILED_FUNCTION_NAME="Fn_CPD_CreateDesignElementWhilePaste"
	Dim objDialog
	Dim aPropertyArr, iCount, aProperty, sPropertyName, sOldTargetType
	Set objDialog = JavaWindow("Collaborative Product").JavaWindow("CreateDesignElement")
	Fn_CPD_CreateDesignElementWhilePaste = False
	If Fn_SISW_UI_Object_Operations("Fn_CPD_CreateDesignElementWhilePaste","Exist",objDialog,SISW_MIN_TIMEOUT) = False Then
	'If Fn_UI_ObjectExist("Fn_CPD_CreateDesignElementWhilePaste",objDialog) = False Then
			Exit Function
	End If

   Select Case lcase(sAction)
		Case "create"
				' Select Desing Element Type
				
				sOldTargetType = sTargetType
			  	sTargetType = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("CPD_DisplayName"),"Design Element")
			    If sTargetType = False Then
			        sTargetType = sOldTargetType
			    End If 
				If sTargetType <> "" Then
					If Instr(sTargetType,"~")>0 OR Instr(sTargetType,":")>0 Then
						aPropertyArr = Split(sTargetType,"~")
						For iCount = 0 To UBound(aPropertyArr)
							aProperty = Split(aPropertyArr(iCount),":")
							Select Case aProperty(0)
								Case "TargetType"
									sPropertyName = "Target Type:"
								Case "SavedVariantRule"
									sPropertyName = "Saved Variant Rule:"
							End Select
							objDialog.JavaStaticText("Target Design Element").SetTOProperty "label",sPropertyName
							Wait 1
							Call Fn_List_Select("Fn_CPD_CreateDesignElementWhilePaste", objDialog, "TargetDesignElement",aProperty(1))
						Next
						Wait 1
						objDialog.JavaStaticText("Target Design Element").SetTOProperty "label","Target Type:"
						Set aPropertyArr = Nothing
						Set aProperty = Nothing
						Set sPropertyName = Nothing
					Else
						Call Fn_List_Select("Fn_CPD_CreateDesignElementWhilePaste", objDialog, "TargetDesignElement",sTargetType)
					End If
				End If

				'Click on Button
				Call Fn_Button_Click("Fn_CPD_CreateDesignElementWhilePaste",objDialog,"Next")
				' Set Check Box ON/OFF
				If sCheckout <> "" Then
					If cBool(sCheckout) Then
						Call Fn_CheckBox_Set("Fn_CPD_CreateDesignElementWhilePaste", objDialog, "CheckoutOnCreate","ON")
					Else
						Call Fn_CheckBox_Set("Fn_CPD_CreateDesignElementWhilePaste", objDialog, "CheckoutOnCreate","OFF")
					End If
				End If

				'Set Check Box ON/OFF
				If sCopyEffectivity <> "" Then
					If Instr(sCopyEffectivity,"~")>0 OR Instr(sCopyEffectivity,":")>0 Then
						aPropertyArr = Split(sCopyEffectivity,"~")
						For iCount = 0 To UBound(aPropertyArr)
							aProperty = Split(aPropertyArr(iCount),":")
							Select Case aProperty(0)
								Case "Applyeffectivitybasedontargetrevisionrule"
									sPropertyName = "Apply effectivity based on target revision rule"
								Case "Applyeffectivitybasedonsourceobject"
									sPropertyName = "Apply effectivity based on source object"
								Case "Applyvariantconditionsbasedonsourceobject"
									sPropertyName = "Apply variant conditions based on source object"
							End Select
							objDialog.JavaCheckBox("CopyEffectivityData").SetTOProperty "attached text",sPropertyName
							Wait 1
							If cBool(aProperty(1)) Then
								Call Fn_CheckBox_Set("Fn_CPD_CreateDesignElementWhilePaste", objDialog, "CopyEffectivityData","ON")
							Else
								Call Fn_CheckBox_Set("Fn_CPD_CreateDesignElementWhilePaste", objDialog, "CopyEffectivityData","OFF")
							End If
						Next
						Wait 1
						objDialog.JavaCheckBox("CopyEffectivityData").SetTOProperty "attached text", "Apply effectivity based on target revision rule"
						Set aPropertyArr = Nothing
						Set aProperty = Nothing
						Set sPropertyName = Nothing
					Else
						If cBool(sCopyEffectivity) Then
							Call Fn_CheckBox_Set("Fn_CPD_CreateDesignElementWhilePaste", objDialog, "CopyEffectivityData","ON")
						Else
							Call Fn_CheckBox_Set("Fn_CPD_CreateDesignElementWhilePaste", objDialog, "CopyEffectivityData","OFF")
						End If						
					End If
				End If

				'Click on Button
				'If sButton <> "" Then
				'	Call Fn_Button_Click("Fn_CPD_CreateDesignElementWhilePaste",objDialog,sButton)
				'Else
				'	Call Fn_Button_Click("Fn_CPD_CreateDesignElementWhilePaste",objDialog,"Finish")
				'End If
				'Swapnil : 28-01-13 : Design Change : 
				
				wait 1
				
				Call Fn_Button_Click("Fn_CPD_CreateDesignElementWhilePaste",objDialog,"Finish")
				
				Fn_CPD_CreateDesignElementWhilePaste = True
	' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_CPD_CreateDesignElementWhilePaste ] Invalid case [ " & sAction & " ].")
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
	End Select
	If  Fn_CPD_CreateDesignElementWhilePaste <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [ Fn_CPD_CreateDesignElementWhilePaste ] executed successfully with case [ " & sAction & " ].")
	End If
	Set objDialog = Nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name			:	Fn_CPD_CreatePartitionWhilePaste
'@@
'@@    Description				:	Function Used to perform operations when Pasting Partition Item onto a Partition.
'@@
'@@    Parameters			    :	1. sAction			: Action to be performed
'@@								:	2. sPartitionType
'@@								:	3. sCopyEffData
'@@								:	4. sButton
'@@
'@@    Return Value		   	   	: 	True Or False
'@@
'@@    Pre-requisite			:	[ Create Partition in CD ] dialog should be OPENED.
'@@
'@@    Examples					:	Call  Fn_CPD_CreatePartitionWhilePaste("Create", "Partition Funcntional", "" , "OK" )
'@@    Examples					:	Call  Fn_CPD_CreatePartitionWhilePaste("ListVerify", "Partition Funcntional~Partition System", "", "" )
'@@
'@@	   History					:	
'@@									Developer Name				Date			 Rev. No.	   Changes Done		 Reviewer
'@@------------------------------------------------------------------------------------------------------------------------------
'@@								  Amit Talegaonkar			20 - Mar - 2012		  1.0			 Created
'@@------------------------------------------------------------------------------------------------------------------------------
'@@								  Koustubh Watwe			21 - Mar - 2012		  1.0			 Added case IsEnabledEffectivityConfigCheckbox
'@@------------------------------------------------------------------------------------------------------------------------------
'@@								  Pranav Ingle					29 - Apr - 2013		  1.0			 Added Code To handle CreatePartitionInPartitionTemplate Dialog
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_CPD_CreatePartitionWhilePaste( sAction , sPartitionType , sCopyEffData , sButton )
	
	GBL_FAILED_FUNCTION_NAME="Fn_CPD_CreatePartitionWhilePaste"
	Dim objDialog , ArrList , intg , sListExist
		
	Fn_CPD_CreatePartitionWhilePaste = False
	
	If Fn_UI_ObjectExist("Fn_CPD_CreatePartitionWhilePaste",JavaWindow("Collaborative Product").JavaWindow("CreatePartitionInCD")) Then
		Set objDialog = JavaWindow("Collaborative Product").JavaWindow("CreatePartitionInCD")
	ElseIf Fn_UI_ObjectExist("Fn_CPD_CreatePartitionWhilePaste",JavaWindow("Collaborative Product").JavaWindow("CreatePartitionInPartitionTemplate")) Then
		Set objDialog = JavaWindow("Collaborative Product").JavaWindow("CreatePartitionInPartitionTemplate")
	Else
    	Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_CPD_CreatePartitionWhilePaste ] failed since [ CreatePartitionInCD ] dialog does not Exist.")
		Exit Function
	End If

   Select Case sAction
   
		Case "Create"

				' Select Partition Type
				If sPartitionType <> "" Then
                    Call Fn_List_Select("Fn_CPD_CreatePartitionWhilePaste", objDialog, "TargetPartitionType",sPartitionType)
				End If

				' Set Check Box [ Copy effectivity Data ] ON/OFF
				If sCopyEffData <> "" Then
					If cBool(sCopyEffData) Then
						Call Fn_CheckBox_Set("Fn_CPD_CreatePartitionWhilePaste", objDialog, "CopyEffData","ON")
					Else
						Call Fn_CheckBox_Set("Fn_CPD_CreatePartitionWhilePaste", objDialog, "CopyEffData","OFF")
					End If
				End If

				'Click on Button
				If sButton <> "" Then
                    Call Fn_Button_Click("Fn_CPD_CreatePartitionWhilePaste",objDialog,sButton)
				Else
					Call Fn_Button_Click("Fn_CPD_CreatePartitionWhilePaste",objDialog,"OK")
				End If
				Fn_CPD_CreatePartitionWhilePaste = True
				
	' 	- - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - 
	
		Case "ListVerify"
		
				ArrList = Split( sPartitionType , "~" )
				
				For intg = 0 to UBound(ArrList) - 1
					sListExist = Fn_UI_ListItemExist( "Fn_CPD_CreatePartitionWhilePaste" , objDialog , "TargetPartitionType" , ArrList(intg))
					If sListExist = False Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_CPD_CreatePartitionWhilePaste ] failed in Case - [ " & sAction & " ]. Partition Type - [ "&ArrList(intg)&" ] does not exist.")
						Exit function
					End If
				Next
				
				'Click on Button
				If sButton <> "" Then
                    Call Fn_Button_Click("Fn_CPD_CreatePartitionWhilePaste",objDialog,sButton)
				End If
				
				Fn_CPD_CreatePartitionWhilePaste = True				
				
	' 	- - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - 
		Case "IsEnabledEffectivityConfigCheckbox"
				If cInt(objDialog.JavaCheckBox("CopyEffData").getROPRoperty("enabled")) = 1 then
					Fn_CPD_CreatePartitionWhilePaste = True				
				End If
				'Click on Button
				If sButton <> "" Then
                    Call Fn_Button_Click("Fn_CPD_CreatePartitionWhilePaste",objDialog,sButton)
				End If
	' 	- - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_CPD_CreatePartitionWhilePaste ] Invalid case [ " & sAction & " ].")
		
	'	 - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - -
	End Select
	
	If  Fn_CPD_CreatePartitionWhilePaste <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [ Fn_CPD_CreatePartitionWhilePaste ] executed successfully with Case [ " & sAction & " ].")
	End If
	
	Set objDialog = Nothing

End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name			:	Fn_CPD_Revise
'@@
'@@    Description				:	Function Used to perform revise operations.
'@@
'@@    Parameters			    :	1. sAction			: Action to be performed
'@@								:	2. bCheckOut		: True / False value to set Check-Out The New Objects checkbox
'@@								:	3. sBtnName			: Button Name
'@@
'@@    Return Value		   	   	: 	True Or False
'@@
'@@    Pre-requisite			:	[ Create Partition in CD ] dialog should be OPENED.
'@@
'@@    Examples					:	Call  Fn_CPD_Revise("Revise", True, "No")
'@@
'@@	   History					:	
'@@									Developer Name				Date			 Rev. No.	   Changes Done		 Reviewer
'@@------------------------------------------------------------------------------------------------------------------------------
'@@								  Koustubh Watwe			29 - Mar - 2012		  1.0			created
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_CPD_Revise(sAction, bCheckOut, sBtnName )
	GBL_FAILED_FUNCTION_NAME="Fn_CPD_Revise"
	Dim objRevise
	Fn_CPD_Revise = False
	Set objRevise = JavaWindow("Collaborative Product").JavaWindow("Revise")

	If Fn_UI_ObjectExist("Fn_CPD_Revise", objRevise) = False  Then
		' perform menu operation
		Call Fn_MenuOperation("Select","File:Revise...")
		Call Fn_ReadyStatusSync(2)
		If Fn_UI_ObjectExist("Fn_CPD_Revise", objRevise) = False  Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_CPD_Revise ] Failed to open Revise window.")
			Exit function
		End If
	End If
	Select Case sAction
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "Revise"
			If bCheckOut <> "" Then
				If cBool(bCheckOut) Then
					Call Fn_CheckBox_Set("Fn_CPD_Revise",objRevise,"CheckOutTheNewObjects","ON")
				Else
					Call Fn_CheckBox_Set("Fn_CPD_Revise",objRevise,"CheckOutTheNewObjects","OFF")
				End If
				If sBtnName = "" Then sBtnName = "Yes"
				Call Fn_Button_Click("Fn_CPD_Revise", objRevise, sBtnName )
				Fn_CPD_Revise = True
			End IF
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_CPD_Revise ] Invalid case [ " & sAction & " ].")
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
	End Select
	If  Fn_CPD_Revise <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [ Fn_CPD_Revise ] executed successfuly with case [ " & sAction & " ].")
	End If
	Set objRevise = Nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@	Function Name		:	Fn_CPD_UpdateDesignElement
'@@
'@@	Description			:	Function used to perform Update Design Element.
'@@
'@@	Parameters			:	1. sAction			: Action to be performed
'@@						:	2. sMsg				: error message / static text
'@@						:	3. sItemRevision	: Revision ID
'@@						:	4. sBtnName			: Button Name
'@@
'@@	Return Value		: 	True Or False
'@@
'@@	Pre-requisite		:	[ Update Design Element ] dialog should be OPENED.
'@@
'@@	Examples			:	Call  Fn_CPD_UpdateDesignElement("verify", "", "A", "Cancel" )
'@@	Examples			:	Call  Fn_CPD_UpdateDesignElement("Set", "", "A", "OK" )
'@@
'@@	History				:	
'@@	Developer Name			Date			 Rev. No.	   Changes Done		 Reviewer
'@@------------------------------------------------------------------------------------------------------------------------------
'@@	Koustubh Watwe			4-Apr-2012		  1.0			created
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_CPD_UpdateDesignElement(sAction, sMsg, sItemRevision, sBtnName )
	GBL_FAILED_FUNCTION_NAME="Fn_CPD_UpdateDesignElement"
	Dim objUpdateDE
	Fn_CPD_UpdateDesignElement = False
	Set objUpdateDE = JavaWindow("Collaborative Product").JavaWindow("UpdateDesignElement")
	'If Fn_UI_ObjectExist("Fn_CPD_UpdateDesignElement", objUpdateDE) = False Then
	If Fn_SISW_UI_Object_Operations("Fn_CPD_UpdateDesignElement","Exist",objUpdateDE,SISW_MICRO_TIMEOUT) = False Then
		Exit function
	End If
	Select Case sAction
				Case "SeletcObject"
'					For iCounter=0 to objUpdateDE.JavaTable("UpdateUsingItemRevision").GetROProperty("rows")-1
					For iCounter=0 to Fn_UI_Object_GetROProperty("Fn_CPD_UpdateDesignElement",objUpdateDE.JavaTable("UpdateUsingItemRevision"), "rows")-1
						Wait 1
						If cStr(objUpdateDE.JavaTable("UpdateUsingItemRevision").Object.getItem(iCounter).getdata().getProperty("object_string")) = sItemRevision Then
							Wait 1
							objUpdateDE.JavaTable("UpdateUsingItemRevision").SelectCell iCounter,"Object"
							Wait 3
							Fn_CPD_UpdateDesignElement=True
							Exit for
						End If
					Next
		Case "VerifyObject"
				' verifying Item revision
				If sItemRevision <> "" Then
					Fn_CPD_UpdateDesignElement=False
'					For iCounter=0 to objUpdateDE.JavaTable("UpdateUsingItemRevision").GetROProperty("rows")-1
					For iCounter=0 to Fn_UI_Object_GetROProperty("Fn_CPD_UpdateDesignElement",objUpdateDE.JavaTable("UpdateUsingItemRevision"), "rows")-1
							If cStr(objUpdateDE.JavaTable("UpdateUsingItemRevision").Object.getItem(iCounter).getdata().getProperty("object_string")) = sItemRevision Then
							Fn_CPD_UpdateDesignElement=True
							Exit for
						End If
					Next
				End If
		Case "Verify"
				' verifying Item revision
				If sItemRevision <> "" Then
					Fn_CPD_UpdateDesignElement = Fn_UI_ListItemExist("Fn_CPD_UpdateDesignElement", objUpdateDE, "ItemRevision", sItemRevision)
				End If

				' verifying message
				If sMsg <> "" Then
					If lcase(objUpdateDE.JavaStaticText("Msg").GetROProperty("label")) = LCase(sMsg) Then
						Fn_CPD_UpdateDesignElement = True
					Else
						Fn_CPD_UpdateDesignElement = False
					End If
				End If
	
		Case "Set"
			If sItemRevision <> "" Then
				Fn_CPD_UpdateDesignElement = Fn_UI_ListItemExist("Fn_CPD_UpdateDesignElement", objUpdateDE, "ItemRevision", sItemRevision)
				If Fn_CPD_UpdateDesignElement Then
					Call Fn_List_Select("Fn_CPD_UpdateDesignElement", objUpdateDE, "ItemRevision", sItemRevision)
				End If
			End If

		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_CPD_UpdateDesignElement ] Invalid case [ " & sAction & " ].")
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
	End Select

	If sBtnName <> "" Then
		Call Fn_Button_Click("Fn_CPD_UpdateDesignElement", objUpdateDE,sBtnName)
	End If

	If  Fn_CPD_UpdateDesignElement <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [ Fn_CPD_UpdateDesignElement ] executed successfuly with case [ " & sAction & " ].")
	End If
	Set objRevise = Nothing
End Function

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_CPD_AttributeGroupCreate
'@@
'@@    Description				:	Function Used to create Managed Attribute Group
'@@
'@@    Parameters			    :	1. sAction		: Action [Type of Attribute Group]
'@@								             : 	 2. sName		: Name
'@@											:	3. sDescription : description
'@@
'@@    Return Value		   	   	: 	True Or False /  ModelID or False
'@@
'@@    Pre-requisite			:	CPD perspective should be activated and Partition is Selected.						
'@@
'@@    Examples					:	Call Fn_CPD_AttributeGroupCreate("CreateAttributeGroup",  "Attr1", "Att_Desc")
'@@    Examples					:	Call Fn_CPD_AttributeGroupCreate("CreateManagedAttributeGroup", "ManagedAttr1", "ManagedAtt_Desc")
'@@
'@@	   History					:	
'@@
'@@					Developer Name				Date			Rev. No.	Changes Done								Reviewer
'@@
'@@					Sachin Joshi			19-April-2012			1.0			Developed
'@@
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_CPD_AttributeGroupCreate(sAction, sName, sDescription)
	GBL_FAILED_FUNCTION_NAME="Fn_CPD_AttributeGroupCreate"
	Dim objAttributeCreate
	Fn_CPD_AttributeGroupCreate = False
	Set objAttributeCreate = JavaWindow("Collaborative Product").JavaWindow("NewBusinessObject")
	
	If Fn_UI_ObjectExist("Fn_CPD_AttributeGroupCreate", objAttributeCreate) = False Then
		Call Fn_MenuOperation("Select","File:New:Attribute Group...")

		If Fn_UI_ObjectExist("Fn_CPD_AttributeGroupCreate", objAttributeCreate) = False Then
            Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_CPD_AttributeGroupCreate ] Failed to opn Collaborative Design window.")
			Set objAttributeCreate = Nothing
		End IF
	End IF

	Select Case sAction
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "CreateAttributeGroup","CreateManagedAttributeGroup"
			' select Attribute Tree
			objAttributeCreate.JavaTree("BusinessObjectType").Expand "Complete List"
			wait 1
			If sAction = "CreateAttributeGroup" Then
				objAttributeCreate.JavaTree("BusinessObjectType").Select "Complete List:Attribute Group Custom"
			Elseif sAction = "CreateManagedAttributeGroup" Then
				objAttributeCreate.JavaTree("BusinessObjectType").Select "Complete List:Managed Attribute Group Custom"
			End If

			' click next Button
			Call Fn_Button_Click("Fn_CPD_AttributeGroupCreate",objAttributeCreate,"Next" )

			'set name
			If sName <> "" Then
				objAttributeCreate.JavaStaticText("Field").SetTOProperty "label", "Name:"
				objAttributeCreate.JavaEdit("Field").Type sName
				Call Fn_ReadyStatusSync(5)
			End If

			' set description
			If sDescription <> "" Then
				objAttributeCreate.JavaStaticText("Field").SetTOProperty "label", "Description:"
				Call Fn_Edit_Box("Fn_CPD_AttributeGroupCreate",objAttributeCreate,"Field",sDescription)
			End If

			' click finish Button
			Call Fn_Button_Click("Fn_CPD_AttributeGroupCreate",objAttributeCreate,"Finish" )
			Call Fn_ReadyStatusSync(5)
			
			Call Fn_Button_Click("Fn_CPD_AttributeGroupCreate",objAttributeCreate,"Cancel" )
			Fn_CPD_AttributeGroupCreate = True
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_CPD_AttributeGroupCreate ] Invalid case [ " & sAction & " ].")
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
	End Select
	If  Fn_CPD_AttributeGroupCreate <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [ Fn_CPD_AttributeGroupCreate ] executed successfuly with case [ " & sAction & " ].")
	End If
	Set objAttributeCreate = Nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_SISW_CPD_SubsetDefaultsCreate
'@@
'@@    Description			:	 Function to create SubsetDefaults in CPD 
'@@
'@@    Parameters			:   1. sAction : Action need to perform
'@@  				   		:	 2. sName : Name of subset defaults
'@@  					    :   3. sDescription : Description for subset defaults
'@@    				   		:   4. sViewedPartitionScheme : Scheme name
'@@
'@@    Return Value		   	: 	True Or False
'@@
'@@    Pre-requisite		:	CPD perspective should be activated and Related Attribute Groups tab from summary view should be Selected.						
'@@
'@@    Examples				:	Call Fn_SISW_CPD_SubsetDefaultsCreate("Create", "ABC", "aaasfh", "Functional")
'@@
'@@	   History				:	
'@@
'@@	   Developer Name			Date			Rev. No.	Changes Done			Reviewer
'@@
'@@		Veena Gurjar		17-August-2012		 1.0		 Created               Kaustubh Watwe
'@@
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_SISW_CPD_SubsetDefaultsCreate(sAction, sName, sDescription, sViewedPartitionScheme)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_CPD_SubsetDefaultsCreate"
	Dim objSubsetDefaults, objSelectType, objDialog, iItemCnt, iCnt, bFlag
	Fn_SISW_CPD_SubsetDefaultsCreate = False
	Set objSubsetDefaults = JavaWindow("Collaborative Product").JavaWindow("NewBusinessObject")
	
	If Fn_UI_ObjectExist("Fn_SISW_CPD_SubsetDefaultsCreate", objSubsetDefaults) = False Then
		If instr(sAction,"_ByMenu")>0 Then
			Call Fn_MenuOperation("Select","File:New:Attribute Group...")
		Else
			Call Fn_ToolbatButtonClick("Add New...")
		End If
		Call Fn_ReadyStatusSync(1)
		If Fn_UI_ObjectExist("Fn_SISW_CPD_SubsetDefaultsCreate", objSubsetDefaults) = False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_CPD_SubsetDefaultsCreate ] Failed to find [ New Business Object ] window.")
			Exit function
		End IF
	End IF
	
	Select Case sAction
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "Create","Create_ByMenu"
			objSubsetDefaults.JavaTree("BusinessObjectType").Expand "Complete List"
			wait 1
			objSubsetDefaults.JavaTree("BusinessObjectType").Select "Complete List:4G Subset Defaults"
			wait 1
			' click on next
			Call Fn_Button_Click("Fn_SISW_CPD_SubsetDefaultsCreate",objSubsetDefaults,"Next" )
			
			'set value
			If sName <> "" Then
				objSubsetDefaults.JavaStaticText("Field").SetTOProperty "label", "Name:"
				Call Fn_UI_EditBox_Type("Fn_SISW_CPD_SubsetDefaultsCreate", objSubsetDefaults, "Field", sName)
			End If
			
			'set value
			If sDescription <> "" Then
				objSubsetDefaults.JavaStaticText("Field").SetTOProperty "label", "Description:"
				Call Fn_UI_EditBox_Type("Fn_SISW_CPD_SubsetDefaultsCreate", objSubsetDefaults, "Field", sDescription)
			End If
			
			If sViewedPartitionScheme <> "" Then
				objSubsetDefaults.JavaStaticText("Field").SetTOProperty "label", "Viewed Partition Scheme:"
				Call Fn_Button_Click("Fn_SISW_CPD_SubsetDefaultsCreate",objSubsetDefaults,"DropDownBtn" )
					
				Set objSelectType=description.Create()
				objSelectType("Class Name").value = "JavaTable"
				Set objDialog =objSubsetDefaults.ChildObjects(objSelectType)
				iItemCnt = cInt(objDialog(0).GetROProperty("rows"))
				bFlag = False
				For iCnt = 0 to iItemCnt - 1
					If objDialog(0).GetCellData(iCnt,0) =  sViewedPartitionScheme Then
						objDialog(0).ClickCell iCnt,0
						bFlag = True
					    Exit For
					End If
				Next
				If bFlag = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_CPD_SubsetDefaultsCreate ] Failed to select [ Viewed Partition Scheme = " & sViewedPartitionScheme & " ] in [ New Business Object ] window.")
					Exit function
				End If
			End If

			' click on Finish
			Call Fn_Button_Click("Fn_SISW_CPD_SubsetDefaultsCreate",objSubsetDefaults,"Finish" )

			' click on Cancel
			Call Fn_Button_Click("Fn_SISW_CPD_SubsetDefaultsCreate",objSubsetDefaults,"Cancel" )

			Fn_SISW_CPD_SubsetDefaultsCreate = True
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "Verify","Verify_ByMenu"
				
				objSubsetDefaults.JavaTree("BusinessObjectType").Expand "Complete List"
				wait 1
				objSubsetDefaults.JavaTree("BusinessObjectType").Select "Complete List:4G Subset Defaults"
				wait 1
				' click on next
				Call Fn_Button_Click("Fn_SISW_CPD_SubsetDefaultsCreate",objSubsetDefaults,"Next" )
			
				'set value
'				not yet implemented
				If sName <> "" Then
					objSubsetDefaults.JavaStaticText("Field").SetTOProperty "label", "Name:"
				End If
			
				'not yet implemented
				If sDescription <> "" Then
					objSubsetDefaults.JavaStaticText("Field").SetTOProperty "label", "Description:"
				End If
			
				If sViewedPartitionScheme <> "" Then
					objSubsetDefaults.JavaStaticText("Field").SetTOProperty "label", "Viewed Partition Scheme:"
					Call Fn_Button_Click("Fn_SISW_CPD_SubsetDefaultsCreate",objSubsetDefaults,"DropDownBtn" )
					
					Set objSelectType=description.Create()
					objSelectType("Class Name").value = "JavaTable"
					Set objDialog =objSubsetDefaults.ChildObjects(objSelectType)
					iItemCnt = cInt(objDialog(0).GetROProperty("rows"))
					bFlag = False
					For iCnt = 0 to iItemCnt - 1
						If objDialog(0).GetCellData(iCnt,0) =  sViewedPartitionScheme Then
							Call Fn_KeyBoardOperation("SendKeys", "{ESC}")
							bFlag = True
							Exit For
						End If
					Next
					If bFlag = False Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_CPD_SubsetDefaultsCreate ] Failed to verify [ Viewed Partition Scheme = " & sViewedPartitionScheme & " ] in [ New Business Object ] window.")
						Exit function
					End If
				End If

			' click on Cancel
			Call Fn_Button_Click("Fn_SISW_CPD_SubsetDefaultsCreate",objSubsetDefaults,"Cancel" )
			Fn_SISW_CPD_SubsetDefaultsCreate = True

		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_CPD_SubsetDefaultsCreate ] Invalid case [ " & sAction & " ].")
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
	End Select
	If  Fn_SISW_CPD_SubsetDefaultsCreate <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [ Fn_SISW_CPD_SubsetDefaultsCreate ] executed successfuly with case [ " & sAction & " ].")
	End If
	Set objSubsetDefaults = Nothing
	Set objSelectType = Nothing
	Set objDialog = Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_SISW_CPD_CreatePartitionTemplate
'
'Description			 :	Function Used to Create Partition template
'
'Parameters			   :   '1.sAction: Action Name
'										 2.dicPartitionInfo: Partition Information
'
'Return Value		   : 	True Or False
'
'Pre-requisite			:	CPD perspective should be activated.
'
'Examples				:     dicPartitionInfo("PartitionType")="Partition Template"
'                                        dicPartitionInfo("Name")="PartitionTemplate"
'                                        dicPartitionInfo("Description")="PartitionTempDesc"
'                                        dicPartitionInfo("OpenOnCreate")="ON"
'                                        bReturn= Fn_SISW_CPD_CreatePartitionTemplate("Create", dicPartitionInfo)
'
'History					 :			
'			Developer Name				Date				Rev. No.	Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'			Rima P					15-Apr-2013			1.0												Sandeep N
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Public Function Fn_SISW_CPD_CreatePartitionTemplate(sAction, dicPartitionInfo)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_CPD_CreatePartitionTemplate"
	Dim objCDCreate
	Dim sNewItemMenu
	Fn_SISW_CPD_CreatePartitionTemplate = False
	Set objCDCreate = JavaWindow("Collaborative Product").JavaWindow("NewBusinessObject")
	
	sNewItemMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("RAC_Menu"),"FileNewPartitionTemplate")
		
	If Not objCDCreate.Exist(SISW_MIN_TIMEOUT) Then
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - \
		Call Fn_MenuOperation("Select",sNewItemMenu)
		Call Fn_ReadyStatusSync(2)
		'Checking Partion Creation Dialog Open or not		
		If Fn_UI_ObjectExist("Fn_SISW_CPD_CreatePartitionTemplate", objCDCreate) = False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_CPD_CreatePartition ] Failed to open Partition window.")
			Set objCDCreate = Nothing
		End IF
	End IF

	Select Case sAction
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "Create","GetErrorMessageOnCreate"
			' select collaborative design from tree
			If objCDCreate.JavaTree("BusinessObjectType").Exist(5) Then			
				wait 2
				objCDCreate.JavaTree("BusinessObjectType").Expand "Complete List"
				wait 1
				objCDCreate.JavaTree("BusinessObjectType").Select "Complete List:"+dicPartitionInfo("PartitionType")
				' click on next
				Call Fn_Button_Click("Fn_SISW_CPD_CreatePartitionTemplate",objCDCreate,"Next" )
				wait(2)
			End If
			' if ModelD is empty
			objCDCreate.JavaStaticText("Field").SetTOProperty "label", "Partition Template ID:"
			If dicPartitionInfo("PartitionID") = "" Then
				'	then click on assign
				Call Fn_Button_Click("Fn_SISW_CPD_CreatePartitionTemplate",objCDCreate,"Assign" )
				Call Fn_ReadyStatusSync(1)
				Fn_SISW_CPD_CreatePartitionTemplate=Fn_UI_Object_GetROProperty("",objCDCreate.JavaEdit("Field"), "value")
				'Fn_SISW_CPD_CreatePartitionTemplate = objCDCreate.JavaEdit("Field").GetROProperty("value")
			Else
				Call Fn_Edit_Box("Fn_SISW_CPD_CreatePartitionTemplate",objCDCreate,"Field",dicPartitionInfo("PartitionID"))
				Fn_SISW_CPD_CreatePartitionTemplate = True
			End If

			'set name
			If dicPartitionInfo("Name") <> "" Then
				objCDCreate.JavaStaticText("Field").SetTOProperty "label", "Name:"
				'objCDCreate.JavaEdit("Field").Type dicPartitionInfo("Name")
				Call Fn_Edit_Box("Fn_SISW_CPD_CreatePartitionTemplate",objCDCreate,"Field",dicPartitionInfo("Name"))
				Call Fn_ReadyStatusSync(1)
			End If

			' set description
			If dicPartitionInfo("Description") <> "" Then
				objCDCreate.JavaStaticText("Field").SetTOProperty "label", "Description:"
				Call Fn_Edit_Box("Fn_SISW_CPD_CreatePartitionTemplate",objCDCreate,"Field",dicPartitionInfo("Description"))
			End If
			' click on next
		'	Call Fn_Button_Click("Fn_SISW_CPD_CreatePartitionTemplate",objCDCreate,"Next" )
		'	wait(2)

			'Setting open on create Option
			If dicPartitionInfo("OpenOnCreate")<>""  Then
				Call Fn_CheckBox_Set("Fn_SISW_CPD_CreatePartitionTemplate",objCDCreate, "OpenOnCreate",dicPartitionInfo("OpenOnCreate"))
			End If

			' click on finish
			Call Fn_Button_Click("Fn_SISW_CPD_CreatePartitionTemplate",objCDCreate,"Finish" )
			Call Fn_ReadyStatusSync(1)	
            If sAction = "GetErrorMessageOnCreate" Then
				Fn_SISW_CPD_CreatePartitionTemplate = False
				If objCDCreate.JavaWindow("Error").Exist(15) Then
					Fn_SISW_CPD_CreatePartitionTemplate = objCDCreate.JavaWindow("Error").JavaStaticText("ErrorMsg").GetROProperty("value")
					Call Fn_Button_Click("Fn_SISW_CPD_CreatePartitionTemplate",objCDCreate.JavaWindow("Error"),"OK" )
				End If
			End If
			Call Fn_Button_Click("Fn_SISW_CPD_CreatePartitionTemplate",objCDCreate,"Cancel" )
			
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_CPD_CreatePartition ] Invalid case [ " & sAction & " ].")
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
	End Select
	If  Fn_SISW_CPD_CreatePartitionTemplate <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [ Fn_CPD_CreatePartition ] executed successfuly with case [ " & sAction & " ].")
	End If
	Set objCDCreate = Nothing
End Function

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		 :	Fn_SISW_CPD_TargetModelCarryoverOptions
'
'Description			 :	Function Used to Set the target model carry-over options
'
'Parameters			   :   sAction - > Action to be done ( Set, Verify etc)
'								 sActionToPerform -> Action to Perform (eg. Clone of Partition Breakdown, Realization of full Partition Breakdown etc )		
'							     sSelectPartionSchemes - > Names of the Partition Schemes to be selected Separated by ~
'								 sOtherOptions  -> Other options to be specified 
'                                sExtra1            -> Reserved for future use
'
'Return Value		   : 	True Or False
'
'Pre-requisite			:	 The Target Model Carry-over Options dialog should be envoked
'
'Examples				:    breturn = Fn_SISW_CPD_TargetModelCarryoverOptions("Set","Clone of Partition Breakdown","Partition Scheme Functional","","")
'                                 breturn = Fn_SISW_CPD_TargetModelCarryoverOptions("Set","Clone of Partition Breakdown~","Partition Scheme Functional~Partiton Scheme Physical","Copy associated group information","")                                                                                                          
'
'								Pranav - 24-Apr-2013
'								breturn = Fn_SISW_CPD_TargetModelCarryoverOptions("Set","Realization of partial Partition Breakdown","PSFunctional-05228","PTN000198/001;1-PartFunctional-05228/PartFunctional-05228","")
'
'History					 :			
'			Developer Name				Date				Rev. No.	                 Changes Done																		Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'			 Pritam S					17-Apr-2013			  1.0								Developed			              														Sandeep N
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'			 Pranav Ingle			24-Apr-2013			  1.1						Added Code to handle sOtherOptions paramter			              Sandeep N
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'			 Sandeep N			24-Apr-2013			  1.1						Modified Code to handle sOtherOptions paramter			              Rima p
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'			 Pranav Ingle			24-Apr-2013			  1.1						Added Code to handle sCarryOverOptions paramter			              Sandeep N
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'			 Ganesh B				22-May-2014			  1.2						modified Case "Set", "Verify" as per Design changes
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Public Function Fn_SISW_CPD_TargetModelCarryoverOptions(sAction,sActionToPerform,sSelectPartionSchemes,sOtherOptions,sCarryOverOptions)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_CPD_TargetModelCarryoverOptions"
    Dim objTrgModOptDlg, bReturn, aSelectPartionSchemes, iCount
	Dim sExpand, sNodeName, iCounter
    Dim i,aOtherOptions
   Fn_SISW_CPD_TargetModelCarryoverOptions = False
   Set objTrgModOptDlg =Fn_SISW_CPD_GetObject("TargetModelCarryoverOptions")

	'Check the Existence of the TargetModelCarryoverOptions window, if not Exist the Function will be Terminated
	If Not objTrgModOptDlg.Exist(6) Then
            Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_CPD_CreatePartition ] Target Model Carry-over Options Dialog not Exists")
			Set objTrgModOptDlg = Nothing
			Exit Function
	End IF

	Select Case sAction
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "Set"

			'Select the RadioButton of the Action to Perform
			If sActionToPerform<>"" Then
				'Modify the attached text of the JavaRadioButton("ActionToPerform") to the option to be selected
				objTrgModOptDlg.JavaRadioButton("ActionToPerform").SetTOProperty "attached text", sActionToPerform
				' Set the Radiobutton ON
				bReturn = Fn_UI_JavaRadioButton_SetON("Fn_SISW_CPD_TargetModelCarryoverOptions",objTrgModOptDlg,"ActionToPerform")
				If bReturn = False Then
					 Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_CPD_CreatePartition ] Failed to Set The "+sActionToPerform+ "Radio Button")
					 Set objTrgModOptDlg = Nothing
					 Exit Function
				End If
			End If
			Call Fn_ReadyStatusSync(1)
			If objTrgModOptDlg.GetRoProperty("title") = "Model content clone and instantiation" Then  '' added code as per design Changes on Tc111(20140514)
				If  objTrgModOptDlg.JavaButton("Next").GetROProperty("enabled") Then
					bReturn = Fn_Button_Click("Fn_SISW_CPD_TargetModelCarryoverOptions",objTrgModOptDlg,"Next" )
					If bReturn = False Then
						 Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_CPD_CreatePartition ] Failed to Click [ Next ] button ")
						 Set objTrgModOptDlg = Nothing
						 Exit Function
					End If
					Call Fn_ReadyStatusSync(2)
				End If
			End If
			'Select the Partition Schemes, MAke the Checkboxes ON
			If sSelectPartionSchemes<>"" Then
				'Split the multiple PartitionSchemes
				aSelectPartionSchemes = Split(sSelectPartionSchemes,"~",-1,1)
				For iCount =  0 to UBound(aSelectPartionSchemes)
					'Modify the attached text Property of the JavaCheckBox("SelectPartitionSchemes") to the Schemes to be selected
					objTrgModOptDlg.JavaCheckBox("SelectPartitionSchemes").SetTOProperty "attached text", aSelectPartionSchemes(iCount)
					If objTrgModOptDlg.JavaCheckBox("SelectPartitionSchemes").Exist(3) Then
								' Set the Checkboxes ON
								bReturn = Fn_CheckBox_Set("Fn_SISW_CPD_TargetModelCarryoverOptions",objTrgModOptDlg, "SelectPartitionSchemes","ON")
								If bReturn = False Then
									 Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_CPD_CreatePartition ] Failed to Set The "+sActionToPerform+ "Radio Button")
									 Set objTrgModOptDlg = Nothing
									 Exit Function
								End If
					Else
							objTrgModOptDlg.JavaRadioButton("ActionToPerform").SetTOProperty "attached text", aSelectPartionSchemes(iCount)
							' Set the Radiobutton ON
							bReturn = Fn_UI_JavaRadioButton_SetON("Fn_SISW_CPD_TargetModelCarryoverOptions",objTrgModOptDlg,"ActionToPerform")
							If bReturn = False Then
								 Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_CPD_CreatePartition ] Failed to Set The "+sActionToPerform+ "Radio Button")
								 Set objTrgModOptDlg = Nothing
								 Exit Function
							End If
					End If
			     Next
			End If
			Call Fn_ReadyStatusSync(1)

			'Select the Other Options if Required
			If  sOtherOptions <> "" Then
    				'For the Other Options goto Next page by Clicking Next button
					If  objTrgModOptDlg.JavaButton("Next").GetROProperty("enabled") Then
						bReturn = Fn_Button_Click("Fn_SISW_CPD_TargetModelCarryoverOptions",objTrgModOptDlg,"Next" )
						If bReturn = False Then
							 Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_CPD_CreatePartition ] Failed to Click [ Next ] button ")
							 Set objTrgModOptDlg = Nothing
							 Exit Function
						End If
						Call Fn_ReadyStatusSync(2)
					End If

					aOtherOptions=Split(sOtherOptions,"~")
					For iCounter=0 to ubound(aOtherOptions)
						sNodeName = split(aOtherOptions(iCounter),":",-1,1)
						For i=0 to ubound(sNodeName)-1
							If iCounter=0 Then
								sExpand=sNodeName(0)
							else
								sExpand=sExpand+":"+sNodeName(iCounter)
							End If
							bReturn=Fn_UI_JavaTreeGetItemPathExt("Fn_SISW_CPD_TargetModelCarryoverOptions", objTrgModOptDlg.JavaTree("ObjectTree"), sExpand, "", "")
							If bReturn=False then
								   Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_CPD_TargetModelCarryoverOptions ] Failed expand node [ "+sExpand+" ]")
								   Set objTrgModOptDlg = Nothing
								   Exit function
							 Else
								Call Fn_UI_JavaTree_Expand("",JavaWindow("Collaborative Product").JavaWindow("TargetModelCarryoverOptions"),"ObjectTree",bReturn)
							 End if
						Next
						Call Fn_ReadyStatusSync(2)
						bReturn=Fn_UI_JavaTreeGetItemPathExt("Fn_SISW_CPD_TargetModelCarryoverOptions", objTrgModOptDlg.JavaTree("ObjectTree"), aOtherOptions(iCounter), "", "")
						 If bReturn=False then
							   Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_CPD_TargetModelCarryoverOptions ] Failed check node [ "+aOtherOptions(iCounter)+" ]")
							   Set objTrgModOptDlg = Nothing
							   Exit function
						 End if
						 objTrgModOptDlg.JavaTree("ObjectTree").SetItemState CStr(bReturn),micChecked
						 Call Fn_ReadyStatusSync(1)
					Next
               End if

			' CarryOver Options  :   Include child Partition(s) ,  Copy associated attribute group information ,   Apply variant conditions based on source Partition
			If sCarryOverOptions<>"" Then
					'For the CarryOverOptions Options goto Next page by Clicking Next button
    				bReturn = Fn_Button_Click("Fn_SISW_CPD_TargetModelCarryoverOptions",objTrgModOptDlg,"Next" )
    				If bReturn = False Then
    					 Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_CPD_CreatePartition ] Failed to Click [ Next ] button ")
    					 Set objTrgModOptDlg = Nothing
    					 Exit Function
    				End If
    				Call Fn_ReadyStatusSync(2)

					'Split the multiple CarryOver Options
					asCarryOverOptions = Split(sCarryOverOptions,"~",-1,1)
					For iCount =  0 to UBound(asCarryOverOptions)
							'Modify the attached text Property of the JavaCheckBox("SelectPartitionSchemes") to the Schemes to be selected
							objTrgModOptDlg.JavaCheckBox("SelectPartitionSchemes").SetTOProperty "attached text", asCarryOverOptions(iCount)
							' Set the Checkboxes ON
							bReturn = Fn_CheckBox_Set("Fn_SISW_CPD_TargetModelCarryoverOptions",objTrgModOptDlg, "SelectPartitionSchemes","ON")
							If bReturn = False Then
								 Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_CPD_CreatePartition ] Failed to Set The "+sActionToPerform+ "Radio Button")
								 Set objTrgModOptDlg = Nothing
								 Exit Function
							End If
					Next
			End If
			Call Fn_ReadyStatusSync(1)

			'Complete the Action by Clicking the Finish button
			bReturn = Fn_Button_Click("Fn_SISW_CPD_TargetModelCarryoverOptions",objTrgModOptDlg,"Finish" )
			If bReturn = False Then
				 Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_CPD_CreatePartition ] Failed to Click [ Finish ] button ")
				 Set objTrgModOptDlg = Nothing
				 Exit Function
			End If
			Call Fn_ReadyStatusSync(1)

	 '- - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "Verify"

			'Select the RadioButton of the Action to Perform
			If sActionToPerform<>"" Then
				'Modify the attached text of the JavaRadioButton("ActionToPerform") to the option to be selected
				objTrgModOptDlg.JavaRadioButton("ActionToPerform").SetTOProperty "attached text", sActionToPerform
				' Set the Radiobutton ON
				bReturn = Fn_UI_JavaRadioButton_SetON("Fn_SISW_CPD_TargetModelCarryoverOptions",objTrgModOptDlg,"ActionToPerform")
				If bReturn = False Then
					 Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_CPD_CreatePartition ] Failed to Set The "+sActionToPerform+ "Radio Button")
					 Set objTrgModOptDlg = Nothing
					 Exit Function
				End If
			End If
			Call Fn_ReadyStatusSync(1)
			If objTrgModOptDlg.GetRoProperty("title") = "Model content clone and instantiation" Then  '' added code as per design Changes on Tc111(20140514)
				If  objTrgModOptDlg.JavaButton("Next").GetROProperty("enabled") Then
					bReturn = Fn_Button_Click("Fn_SISW_CPD_TargetModelCarryoverOptions",objTrgModOptDlg,"Next" )
					If bReturn = False Then
						 Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_CPD_CreatePartition ] Failed to Click [ Next ] button ")
						 Set objTrgModOptDlg = Nothing
						 Exit Function
					End If
					Call Fn_ReadyStatusSync(2)
				End If
			End If
			'Select the Partition Schemes, MAke the Checkboxes ON
			If sSelectPartionSchemes<>"" Then
				'Split the multiple PartitionSchemes
				aSelectPartionSchemes = Split(sSelectPartionSchemes,"~",-1,1)
				For iCount =  0 to UBound(aSelectPartionSchemes)
					'Modify the attached text Property of the JavaCheckBox("SelectPartitionSchemes") to the Schemes to be selected
					objTrgModOptDlg.JavaCheckBox("SelectPartitionSchemes").SetTOProperty "attached text", aSelectPartionSchemes(iCount)
					If objTrgModOptDlg.JavaCheckBox("SelectPartitionSchemes").Exist(3) Then
								' Set the Checkboxes ON
								bReturn = Fn_CheckBox_Set("Fn_SISW_CPD_TargetModelCarryoverOptions",objTrgModOptDlg, "SelectPartitionSchemes","ON")
								If bReturn = False Then
									 Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_CPD_CreatePartition ] Failed to Set The "+sActionToPerform+ "Radio Button")
									 Set objTrgModOptDlg = Nothing
									 Exit Function
								End If
					Else
							objTrgModOptDlg.JavaRadioButton("ActionToPerform").SetTOProperty "attached text", aSelectPartionSchemes(iCount)
							' Set the Radiobutton ON
							bReturn = Fn_UI_JavaRadioButton_SetON("Fn_SISW_CPD_TargetModelCarryoverOptions",objTrgModOptDlg,"ActionToPerform")
							If bReturn = False Then
								 Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_CPD_CreatePartition ] Failed to Set The "+sActionToPerform+ "Radio Button")
								 Set objTrgModOptDlg = Nothing
								 Exit Function
							End If
					End If
			     Next
			End If
			Call Fn_ReadyStatusSync(1)

			'Select the Other Options if Required
			If  sOtherOptions <> "" Then
    				'For the Other Options goto Next page by Clicking Next button
    				bReturn = Fn_Button_Click("Fn_SISW_CPD_TargetModelCarryoverOptions",objTrgModOptDlg,"Next" )
    				If bReturn = False Then
    					 Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_CPD_CreatePartition ] Failed to Click [ Next ] button ")
    					 Set objTrgModOptDlg = Nothing
    					 Exit Function
    				End If
    				Call Fn_ReadyStatusSync(2)

				    aOtherOptions=Split(sOtherOptions,"~")
                    For iCounter=0 to ubound(aOtherOptions)
                        sNodeName = split(aOtherOptions(iCounter),":",-1,1)
                        For i=0 to ubound(sNodeName)-1
                            If iCounter=0 Then
        						sExpand=sNodeName(0)
        					else
        						sExpand=sExpand+":"+sNodeName(iCounter)
        					End If
                            bReturn=Fn_UI_JavaTreeGetItemPathExt("Fn_SISW_CPD_TargetModelCarryoverOptions", objTrgModOptDlg.JavaTree("ObjectTree"), sExpand, "", "")
                            If bReturn=False then
                                   Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_CPD_TargetModelCarryoverOptions ] Failed expand node [ "+sExpand+" ]")
                                   Set objTrgModOptDlg = Nothing
                                   Exit function
                             Else
                                Call Fn_UI_JavaTree_Expand("",JavaWindow("Collaborative Product").JavaWindow("TargetModelCarryoverOptions"),"ObjectTree",bReturn)
                             End if
                        Next
                        Call Fn_ReadyStatusSync(2)
                         bReturn=Fn_UI_JavaTreeGetItemPathExt("Fn_SISW_CPD_TargetModelCarryoverOptions", objTrgModOptDlg.JavaTree("ObjectTree"), aOtherOptions(iCounter), "", "")
                         If bReturn=False then
                               Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_CPD_TargetModelCarryoverOptions ] Failed check node [ "+aOtherOptions(iCounter)+" ]")
                               Set objTrgModOptDlg = Nothing
                               Exit function
                         End if
                         Call Fn_ReadyStatusSync(1)
                    Next
               End if

			' CarryOver Options  :   Include child Partition(s) ,  Copy associated attribute group information ,   Apply variant conditions based on source Partition
			If sCarryOverOptions<>"" Then
					'For the CarryOverOptions Options goto Next page by Clicking Next button
    				bReturn = Fn_Button_Click("Fn_SISW_CPD_TargetModelCarryoverOptions",objTrgModOptDlg,"Next" )
    				If bReturn = False Then
    					 Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_CPD_CreatePartition ] Failed to Click [ Next ] button ")
    					 Set objTrgModOptDlg = Nothing
    					 Exit Function
    				End If
    				Call Fn_ReadyStatusSync(2)

					'Split the multiple CarryOver Options
					asCarryOverOptions = Split(sCarryOverOptions,"~",-1,1)
					For iCount =  0 to UBound(asCarryOverOptions)
							'Modify the attached text Property of the JavaCheckBox("SelectPartitionSchemes") to the Schemes to be selected
							objTrgModOptDlg.JavaCheckBox("SelectPartitionSchemes").SetTOProperty "attached text", asCarryOverOptions(iCount)

							' Verify the Checkboxes 
                            bReturn=Fn_UI_ObjectExist("",objTargetModelCarryoverOptions.JavaCheckBox("SelectPartitionSchemes"))
                            If bReturn = False Then
								 Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_CPD_CreatePartition ] Failed to Verify The "+asCarryOverOptions(iCount)+ "Check Box")
								 Set objTrgModOptDlg = Nothing
								 Exit Function
							End If
					Next
			End If
			Call Fn_ReadyStatusSync(1)

			'Complete the Action by Clicking the Finish button
			bReturn = Fn_Button_Click("Fn_SISW_CPD_TargetModelCarryoverOptions",objTrgModOptDlg,"Cancel" )
			If bReturn = False Then
				 Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_CPD_CreatePartition ] Failed to Click [ Cancel ] button ")
				 Set objTrgModOptDlg = Nothing
				 Exit Function
			End If
			Call Fn_ReadyStatusSync(1)
				
	End Select

	Fn_SISW_CPD_TargetModelCarryoverOptions = True
	Set objTrgModOptDlg = Nothing
End Function


'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name			:	Fn_SISW_CPD_UpdatePartition
'@@
'@@    Description				:	Function Used to create Collaborative Design
'@@
'@@    Parameters			    :	1. sAction		: Action to be performed
'@@								:	2. sModelID		: Model ID
'@@								:	3. sCheckBox	: To set Partition template checkbox
'@@
'@@    Return Value		   	   	: 	True Or False
'@@
'@@    Examples					:	Call Fn_SISW_CPD_UpdatePartition("Set", "Working(Current User); Any Status", "Include child Partition(s)", "OK")
'@@
'@@	   History					:	
'@@				Developer Name				Date				Rev. No.		Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Pranav Ingle				25-Apr-2013			1.0					Created												Sandeep
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function Fn_SISW_CPD_UpdatePartition(sAction, sRevRule, sCheckBox, sBtnName)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_CPD_UpdatePartition"
	Dim objUpdatePartitionFromPartitionTemplate
	Fn_SISW_CPD_UpdatePartition = False

	Set objUpdatePartitionFromPartitionTemplate = JavaWindow("Collaborative Product").JavaWindow("UpdatePartitionFromPartitionTemplate")
    
	Select Case sAction
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "Set"

            'set sRevRule
			If sRevRule <> "" Then
				Call Fn_List_Select("Fn_SISW_CPD_UpdatePartition", objUpdatePartitionFromPartitionTemplate, "UpdateUsingRevisionRule", sRevRule)
            End If
            
            'set checkbox 
	    If sCheckBox <> "" Then
			aCheckbox = Split(sCheckBox,"~")
			For iCount = 0 To Ubound(aCheckbox)
				objUpdatePartitionFromPartitionTemplate.JavaCheckBox("OverrideTargetCarryOver").SetTOProperty "attached text", aCheckbox(iCount)	
				bResult = Fn_SISW_UI_JavaCheckBox_Operations("Fn_SISW_CPD_UpdatePartition", "Set", objUpdatePartitionFromPartitionTemplate.JavaCheckBox("OverrideTargetCarryOver"),"" , "ON")
				wait 1
				If bResult = False Then
					Exit Function
				End If
			Next
		End If
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_CPD_UpdatePartition ] Invalid case [ " & sAction & " ].")
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
	End Select

	' click on OK
	If sBtnName <> "" Then
		Call Fn_Button_Click("Fn_SISW_CPD_UpdatePartition", objUpdatePartitionFromPartitionTemplate,sBtnName)
	End If

	If  Fn_SISW_CPD_UpdatePartition <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [ Fn_SISW_CPD_UpdatePartition ] executed successfuly with case [ " & sAction & " ].")
	End If

	Fn_SISW_CPD_UpdatePartition = True
	Set objUpdatePartitionFromPartitionTemplate = Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_SISW_CPD_CreateSubset
'
'Description			 :	Function Used to Create Subset
'
'Parameters			   :   '1.sAction: Action Name
'										 2.sInvokeOption : Dialog invoke option
'										 3.dicSubsetInfo: Subset Information
'
'Return Value		   : 	True Or False
'
'Pre-requisite			:	CPD perspective should be activated.
'
'Examples				:   Dim dicSubsetInfo
'										Set dicSubsetInfo = CreateObject( "Scripting.Dictionary")
'										dicSubsetInfo("SubsetType")="Subset"
'										dicSubsetInfo("Name")="Subset1"
'										dicSubsetInfo("Model")="Paste"
'										dicSubsetInfo("Description")="Subset Desc"
'                                        bReturn= Fn_SISW_CPD_CreateSubset("Create","ToolbarButton",dicSubsetInfo)
'
'History					 :			
'			Developer Name				Date				Rev. No.	Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'			Sandeep N					02-May-2013			1.0																	Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Public Function Fn_SISW_CPD_CreateSubset(sAction,sInvokeOption,dicSubsetInfo)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_CPD_CreateSubset"
	Dim objCDCreate
	Dim bFlag,ObjDesc,ArrLists,iToolCnt,iCounter,iItmCount,aContents,sItmText,sButtonName,sContents,iCounter1
	Fn_SISW_CPD_CreateSubset = False
	Set objCDCreate = JavaWindow("Collaborative Product").JavaWindow("NewBusinessObject")
	
	If Not objCDCreate.Exist(6) Then
		Select Case sInvokeOption
			' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - \
			Case "ToolbarButton",""
				Call Fn_ToolbarOperation("Click", "Create a new Subset","")
				Call Fn_ReadyStatusSync(1)
		End Select
		'Checking Subset Creation Dialog Open or not		
		If Fn_UI_ObjectExist("Fn_SISW_CPD_CreateSubset", objCDCreate) = False Then
            Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_CPD_CreateSubset ] Failed to open Subset window.")
			Set objCDCreate = Nothing
		End IF
	End IF

	Select Case sAction
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "Create"
			' select collaborative design from tree
			If objCDCreate.JavaTree("BusinessObjectType").Exist(2) = True Then
				objCDCreate.JavaTree("BusinessObjectType").Expand "Complete List"
				wait 1
				objCDCreate.JavaTree("BusinessObjectType").Select "Complete List:"+dicSubsetInfo("SubsetType")
				' click on next
				Call Fn_Button_Click("Fn_SISW_CPD_CreateSubset",objCDCreate,"Next" )
				wait(1)
			End If
			
			'Setting Subset name
			If dicSubsetInfo("Name") <> "" Then
				objCDCreate.JavaStaticText("Field").SetTOProperty "label", "Name:"
				Call Fn_Edit_Box("Fn_SISW_CPD_CreateSubset",objCDCreate,"Field",dicSubsetInfo("Name"))		
			End If
			
			'set Model
			objCDCreate.JavaStaticText("Field").SetTOProperty "label", "Model:"
			Select Case dicSubsetInfo("Model")
				Case "Paste"
					objCDCreate.JavaToolbar("FieldToolBar").Press "Paste the Model from Clipboard"
			End Select
            
			' set description
			If dicSubsetInfo("Description") <> "" Then
				objCDCreate.JavaStaticText("Field").SetTOProperty "label", "Description:"
				Call Fn_Edit_Box("Fn_SISW_CPD_CreateSubset",objCDCreate,"Field",dicSubsetInfo("Description"))
			End If
			
			' set Include In Parts List
			If dicSubsetInfo("IncludeInPartsList") <> "" Then
				If lcase(dicSubsetInfo("IncludeInPartsList")) = "on" Or lcase(dicSubsetInfo("IncludeInPartsList")) = "true" Then
					objCDCreate.JavaRadioButton("IncludeInPartsList").SetTOProperty "attached text","True"
					Call Fn_SISW_UI_JavaRadioButton_Operations("Fn_SISW_CPD_CreateSubset", "Set", objCDCreate, "IncludeInPartsList", "ON")
				ElseIf lcase(dicSubsetInfo("IncludeInPartsList")) = "off" Or lcase(dicSubsetInfo("IncludeInPartsList")) = "false" Then
					objCDCreate.JavaRadioButton("IncludeInPartsList").SetTOProperty "attached text","False"
					Call Fn_SISW_UI_JavaRadioButton_Operations("Fn_SISW_CPD_CreateSubset", "Set", objCDCreate, "IncludeInPartsList", "ON")
				End If
			End If
			
			' set Report In Where Used
			If dicSubsetInfo("ReportInWhereUsed") <> "" Then
				If lcase(dicSubsetInfo("ReportInWhereUsed")) = "on" Or lcase(dicSubsetInfo("ReportInWhereUsed")) = "true" Then
					objCDCreate.JavaRadioButton("ReportInWhereUsed").SetTOProperty "attached text","True"
					Call Fn_SISW_UI_JavaRadioButton_Operations("Fn_SISW_CPD_CreateSubset", "Set", objCDCreate, "ReportInWhereUsed", "ON")
				ElseIf lcase(dicSubsetInfo("ReportInWhereUsed")) = "off" Or lcase(dicSubsetInfo("ReportInWhereUsed")) = "false" Then
					objCDCreate.JavaRadioButton("ReportInWhereUsed").SetTOProperty "attached text","False"
					Call Fn_SISW_UI_JavaRadioButton_Operations("Fn_SISW_CPD_CreateSubset", "Set", objCDCreate, "ReportInWhereUsed", "ON")
				End If
			End If
			
			' click on finish
			Fn_SISW_CPD_CreateSubset=Fn_Button_Click("Fn_SISW_CPD_CreateSubset",objCDCreate,"Finish" )
			Call Fn_ReadyStatusSync(1)
			
			If Fn_UI_ObjectExist("Fn_SISW_CPD_CreateSubset", objCDCreate) = True Then
				Call Fn_Button_Click("Fn_SISW_CPD_CreateSubset",objCDCreate,"Cancel" )
				Call Fn_ReadyStatusSync(1)
			End If	
			
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_CPD_CreateSubset ] Invalid case [ " & sAction & " ].")
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
	End Select
	If  Fn_SISW_CPD_CreateSubset <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [ Fn_CPD_CreateSubset ] executed successfuly with case [ " & sAction & " ].")
	End If
	Set objCDCreate = Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		 	:	Fn_SISW_CPD_NatTable_TreeTable_GetRowIndex

'Description		    :  	Function to return row number of specified value in given column from NatTable.

'Parameters		    :	   1. objNatTable : Object name
'										2. StrNode : Node Path
'										3. sDelimiter : Delimiter
'										4. sInstanceHandler : Instance handler 
'										5. sColumnIndex : Column Index
'										6. sRowStartingIndex : starting Row Index
'										7. iRowIndexIncrementor : Row index incrementor
								
''Return Value		    :  	Row Number \ -1
'
''Examples		     	:	bReturn=Fn_SISW_CPD_NatTable_TreeTable_GetRowIndex(objNatTable, "PTN000127/001;1-ZonePT", "Body:Sedan", "", "", "","")

'History					 :			
'				Developer Name						Date					Rev. No.						Changes Done																				Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'				Sandeep N							7-May-2013				1.0																																				Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Public Function Fn_SISW_CPD_NatTable_TreeTable_GetRowIndex(objNatTable, sColumnIndex, StrNode, sRowStartingIndex, sDelimiter, sInstanceHandler,iRowIndexIncrementor)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_CPD_NatTable_TreeTable_GetRowIndex"
   'Variable Declaration
	Dim sItemPath,aStrNode,bFlag,i,iNodeItemsCount, sFuncLog
	Dim oCurrentNode,eStrNode, iCount, iNodecnt
	Dim iInstanceCnt, aNode,iOccCnt
	Dim sTreeNodeStr
	Dim objRootObjects, objDataProvider, iRowIndex, iColIndex
	
	sFuncLog = "Fn_SISW_CPD_NatTable_TreeTable_GetRowIndex : on [ " & objNatTable.toString() & " ] : "

	If sColumnIndex = "" Then
		iColIndex = 1
	Else
		iColIndex = Fn_SISW_RAC_NatTable_GetColumnIndexExt(objNatTable, sColumnIndex,"","","" )
		If iColIndex = -1 Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & " FAIL : Failed to find column [ " & sColumnIndex & " ].")
			Exit function
		End If
	End If
	If sRowStartingIndex = "" Then
		sRowStartingIndex = 2
	End If
	If Environment.Value("ProductName") = sQTPProductName OR Environment.Value("ProductName") = sUFTProductName Then '' added code by Vivek A(16-Dec-2014) to handle UFT related issue as some method not supported in UFT 
		Set objDataProvider = objNatTable.Object.getCellByPosition(iColIndex, sRowStartingIndex).getlayer.getRowHeaderLayer.getBaselayer.getCellByPosition(iColIndex,sRowStartingIndex).getSourcelayer().getDataProvider()
	Else
		Set objDataProvider = objNatTable.Object.getCellByPosition(iColIndex, sRowStartingIndex).getSourceLayer().getDataProvider()
	End If
	'Set objDataProvider = objNatTable.Object.getCellByPosition(iColIndex, sRowStartingIndex).getSourceLayer().getDataProvider()
	Set objRootObjects = objDataProvider.getTreeList().getRoots()

	If sDelimiter = "" Then sDelimiter = ":"
	If sInstanceHandler = "" Then sInstanceHandler = "@"
	Fn_SISW_CPD_NatTable_TreeTable_GetRowIndex = False
	sTreeNodeStr = ""
	'Initial Item Path
	sItemPath= -1
	aStrNode = Split (StrNode, sDelimiter)
	bFlag=False
	
	'To handle the situation where operation needs to be performed on Root Node
	iOccCnt = 1
	For iCount = 0 to cInt(objRootObjects.size()) - 1
'	For iCount = 0 to cInt(objDataProvider.getRowCount) - 1
		If Instr(aStrNode(0), sInstanceHandler) > 0 Then
			aNode = split(aStrNode(0),sInstanceHandler)
			eStrNode = trim(aNode(0))
			iInstanceCnt = cInt(aNode(1) )
		Else
			eStrNode = trim(aStrNode(0))
			iInstanceCnt = 1
		End If
		If iRowIndexIncrementor="" Then
			iRowIndexIncrementor=1
		End If
		' Index of the Row then get Row text
 		iRowIndex = objDataProvider.indexOfRowObject(objRootObjects.get(iCount).getElement())
		iRowIndex = cInt(iRowIndex) 
		iRowIndex = cInt(iRowIndex) + iRowIndexIncrementor
		
		'sTreeNodeStr = objNatTable.Object.getCellByPosition(0, iRowIndex).getDataValue().getData().getName()
		'Added by Vivek A as per design change
		sTreeNodeStr = objNatTable.Object.getCellByPosition(0, iRowIndex).getDataValue().getData().toString()
		If sTreeNodeStr<>eStrNode Then
			sTreeNodeStr = objNatTable.Object.getCellByPosition(0, iRowIndex).getDataValue().getData().getID()
		End If

        If instr(eStrNode, "[") > 0 AND  instr(eStrNode, "]") > 0   Then
'        If sTreeNodeStr <> eStrNode Then
		    If typename(objNatTable.Object.getCellByPosition(0, iRowIndex).getDataValue().getData().getNameSpace() )<> "Null" Then
				sTreeNodeStr = objNatTable.Object.getCellByPosition(0, iRowIndex).getDataValue().getData().getName()
				sTreeNodeStr = "["+objNatTable.Object.getCellByPosition(0, iRowIndex).getDataValue().getData().getNameSpace()+"]"+sTreeNodeStr
				If sTreeNodeStr <> eStrNode Then
					If TypeName(objNatTable.Object.getCellByPosition(0, iRowIndex).getDataValue().getData().getUOMSymbol()) <> "Null" Then
						sTreeNodeStr = sTreeNodeStr +", ("+objNatTable.Object.getCellByPosition(0, iRowIndex).getDataValue().getData().getUOMSymbol()+")"
					End If
				End If
			End If
		End If
		
		If sTreeNodeStr = eStrNode Then
			If  iOccCnt = iInstanceCnt Then
				Set oCurrentNode = objRootObjects.get(iCount)
				sItemPath = iRowIndex 
				bFlag = True
				Exit For
			else
				iOccCnt = iOccCnt + 1
			End If
		End If
	Next
	If UBound(aStrNode) = 0 Then
		Fn_SISW_CPD_NatTable_TreeTable_GetRowIndex = sItemPath
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & " PASS : executed successfully for item [ " & StrNode & " ]"  )
		Exit Function
	End If
	If bFlag Then
		bFlag = False
	Else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & "Failed to find item [ " & StrNode & " ]"  )
		Exit function
	End If
		'To Select first Occurance of Node
'		For each eStrNode In 
		For iNodecnt = 1 to UBound(aStrNode)
			eStrNode = aStrNode(iNodecnt)
			iNodeItemsCount = cInt(oCurrentNode.getChildren().size())
'			iNodeItemsCount = cInt(oCurrentNode.size())
			bFlag=False
			iOccCnt = 1
			If Instr(eStrNode, sInstanceHandler) > 0 Then
				aNode = split(eStrNode,sInstanceHandler)
				eStrNode = trim(aNode(0))
				iInstanceCnt = cInt(aNode(1) )
			Else
				iInstanceCnt = 1
			End If
			For i = 0 to iNodeItemsCount - 1
				' get text from table
				iRowIndex = objDataProvider.indexOfRowObject(oCurrentNode.getChildren().get(i).getElement())
				iRowIndex = cInt(iRowIndex) + 1
				sTreeNodeStr = objNatTable.Object.getCellByPosition(0, iRowIndex).getDataValue().getData().toString()	'Added by Vivek A as per design change
				If Trim(sTreeNodeStr) <> Trim(eStrNode) Then
					sTreeNodeStr = objNatTable.Object.getCellByPosition(0, iRowIndex).getDataValue().getData().getId()
				End If
				If Trim(sTreeNodeStr) = Trim(eStrNode) Then
					If  iOccCnt = iInstanceCnt Then
						sItemPath = objDataProvider.indexOfRowObject(oCurrentNode.getChildren().get(i).getElement())
						Set oCurrentNode = oCurrentNode.getChildren().get(i)
						bFlag=True
						Exit For
					else
						iOccCnt = iOccCnt + 1
					End If
				End If
			Next
			If bFlag=False Then
				Exit For
			End If
		Next 
	If bFlag=True Then
		'Function Returns Item Path
		Fn_SISW_CPD_NatTable_TreeTable_GetRowIndex = (sItemPath + 1)
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & " PASS : executed successfully for item [ " & StrNode & " ]"  )
	Else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & "Failed to find item [ " & StrNode & " ]"  )
		Fn_SISW_CPD_NatTable_TreeTable_GetRowIndex = False
	End If
	Set oCurrentNode =Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_SISW_CPD_VariantNatTableOperations

'Description			 :	Function Used to perform operations on Variant Nat Tables

'Parameters			   :   1.StrAction: Action Name
'										2.StrInnerTabName: Inner tab Name
'										3.StrTabName: Parent Tab name
'										4.StrNode: Node path
'										5.StrColumnName: Column name
'										6.StrValue: Expected value
'										7.StrMessage: Output message
'										8.StrPopupmenu: Popup menu
'
'Return Value		   : 	True or False

'Pre-requisite			:	Variant Nat table should be appear

'Examples				:   bReturn=Fn_SISW_CPD_VariantNatTableOperations("SetFlag","","","Body:Sedan","PTN000127/001;1-ZonePT","Check","Body = 'Sedan'","")
'										bReturn=Fn_SISW_CPD_VariantNatTableOperations("Save","","","","","","","")
'										bReturn=Fn_SISW_CPD_VariantNatTableOperations("VerifyOutput","","","","","","NOT Body = 'Sedan'","")
'										bReturn=Fn_SISW_CPD_VariantNatTableOperations("VerifyNode","","","Body48624:HatchBack~Body48624:Sedan~Engine48624:180 HP","PTN000143/001;1-ZonePartition148624","","","")
'
'History					 :			
'				Developer Name						Date					Rev. No.						Changes Done																				Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'				Sandeep N							7-May-2013				1.0																																				Sunny R
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'				Sandeep N							7-May-2013				1.1						Added case : VerifyNode																		Veena G
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_SISW_CPD_VariantNatTableOperations(StrAction,StrInnerTabName,StrTabName,StrNode,StrColumnName,StrValue,StrMessage,StrPopupmenu)
		GBL_FAILED_FUNCTION_NAME="Fn_SISW_CPD_VariantNatTableOperations"
	 	'Declaring variables
		Dim ObjVariantNatTable
		Dim iColIndex,iRowIndex,iX,iY,StrCurrentOutput,aNode,iCounter
		Dim iRow, iCol, objRow
		
		Fn_SISW_CPD_VariantNatTableOperations=False
		'Creating object of [ Applicability ] table
		Set ObjVariantNatTable=JavaWindow("Collaborative Product").JavaObject("VariantNatTable")
	
	   Select Case StrAction
	 		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			Case "SetFlag"
				If StrTabName<>"" Then
					'Double click on tab
					If Fn_TabFolder_Operation("DoubleClickTab", StrTabName, "")=False then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Fail to double click on Tab [ "+StrTabName+" ]" )
						Set ObjVariantNatTable=Nothing
						Exit function
					End if
					wait 2
				End If
				'Getting column index
				iColIndex = Fn_SISW_RAC_NatTable_GetColumnIndexExt(ObjVariantNatTable, StrColumnName,"","","")
				'Getting row index
				If Environment.Value("ProductName") = sQTPProductName OR Environment.Value("ProductName") = sUFTProductName Then	'' added code by Vivek(16-dec-2014) to handle UFT related issue as some method not supported in UFT 
					iRowIndex=Fn_SISW_CPD_NatTable_TreeTable_GetRowIndex(ObjVariantNatTable, StrColumnName, StrNode, 0, "", "","")
				Else
					iRowIndex=Fn_SISW_CPD_NatTable_TreeTable_GetRowIndex(ObjVariantNatTable, StrColumnName, StrNode, 1, "", "","")
				End If
				'iRowIndex=Fn_SISW_CPD_NatTable_TreeTable_GetRowIndex(ObjVariantNatTable, StrColumnName, StrNode, 1, "", "","")
				
				If iColIndex <> -1 and iRowIndex <> -1 Then
                    iRow=Cint(iRowIndex)-1
'                    iCol=Cint(iColIndex)-1
					'Added New Code
					For iCounter=0 to ObjVariantNatTable.Object.getCellByPosition(iColIndex, iRowIndex).getSourceLayer().getColumnCount-1
						If Environment.Value("ProductName") = sQTPProductName OR Environment.Value("ProductName") = sUFTProductName Then	'' added code by Vivek (16-dec-2014) to handle UFT related issue as some method not supported in UFT 
						   If ObjVariantNatTable.Object.getCellByPosition(iCounter, iRowIndex).getlayer.getColumnHeaderlayer.getDatavalueByPosition(iCounter,iRowIndex).tostring=CStr(StrColumnName) Then
						       iCol=iCounter-1
							   Exit For
						   End If
						Else
							If ObjVariantNatTable.Object.getCellByPosition(iColIndex, iRowIndex).getSourceLayer().getDataProvider().getColumnHeaderDataValue(iCounter).toString()=CStr(StrColumnName) Then
								iCol=iCounter-1
								Exit For
							End If
						End If
					Next

					If Environment.Value("ProductName") = sQTPProductName OR Environment.Value("ProductName") = sUFTProductName Then'' added code by Vivek(16-dec-2014) to handle UFT related issue as some method not supported in UFT 
					    iX = ObjVariantNatTable.Object.getStartXOfColumnPosition(cdbl(ObjVariantNatTable.Object.getCellByPosition(iColIndex, iRowIndex).getSourceLayer().getColumnCount)+iColIndex) + 18
					    iY = ObjVariantNatTable.Object.getStartYOfRowPosition(iRowIndex) + 4
					    If cdbl(ObjVariantNatTable.Object.getCellByPosition(iColIndex, iRowIndex).getSourceLayer().getRowCount) = iRowIndex Then 
						     iRowIndex = iRowIndex -1
						End if
					Else				
						iX = ObjVariantNatTable.Object.getStartXOfColumnPosition(iColIndex) + 18
						iY = ObjVariantNatTable.Object.getStartYOfRowPosition(iRowIndex) + 4
					End If
					Select Case StrValue
						'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
						Case "Check"
							For iCounter=0 to 2
								If Environment.Value("ProductName") = sQTPProductName OR Environment.Value("ProductName") = sUFTProductName Then'' added code by Vivek(16-dec-2014) to handle UFT related issue as some method not supported in UFT 
									Set objRow=ObjVariantNatTable.Object.getCellByPosition(iColIndex, iRowIndex).getlayer.getRowHeaderLayer.getBaselayer.getCellByPosition(iColIndex,iRowIndex).getSourcelayer.getDataProvider().getRowObjects().get(iRow).getData()
									If ISEmpty(ObjVariantNatTable.Object.getCellByPosition(iColIndex, iRowIndex).getlayer.getRowHeaderLayer.getBaselayer.getCellByPosition(iColIndex,iRowIndex).getSourcelayer.getDataProvider().getColumnObjects().get(iCol).get().get(objRow)) then
	                                   StrCurrentOutput=""
	                                Else
	                               	    StrCurrentOutput=ObjVariantNatTable.Object.getCellByPosition(iColIndex, iRowIndex).getlayer.getRowHeaderLayer.getBaselayer.getCellByPosition(iColIndex,iRowIndex).getSourcelayer.getDataProvider().getColumnObjects().get(iCol).get().get(objRow).toString()
	                                End if
								Else
	                                Set objRow=ObjVariantNatTable.Object.getCellByPosition(iColIndex, iRowIndex).getSourceLayer().getDataProvider().getRowObjects().get(iRow).getData()
									If ObjVariantNatTable.Object.getCellByPosition(iColIndex, iRowIndex).getSourceLayer().getDataProvider().getColumnObjects().get(iCol).get().get(objRow) is Nothing then
	                                    StrCurrentOutput=""
	                                Else
	                               	    StrCurrentOutput=ObjVariantNatTable.Object.getCellByPosition(iColIndex, iRowIndex).getSourceLayer().getDataProvider().getColumnObjects().get(iCol).get().get(objRow).toString()
	                                End if
                               End If
								If StrCurrentOutput="TRUE" Then
                                    Fn_SISW_CPD_VariantNatTableOperations=True
									Exit for
								Else
									ObjVariantNatTable.Click iX, iY, "LEFT"
									wait 1
								End If
							Next
						'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
						Case "None"
							aNode=Split(StrNode,":")
							For iCounter=0 to 2
								StrCurrentOutput=JavaWindow("Collaborative Product").JavaEdit("VariantNatTableOutput").GetROProperty("value")
								If instr(1,StrCurrentOutput,aNode(ubound(aNode))) Then
									ObjVariantNatTable.Click iX, iY, "LEFT"
									wait 1
								Else
									Exit for
								End If
							Next
							If instr(1,StrCurrentOutput,aNode(ubound(aNode))) Then
							Else
								Fn_SISW_CPD_VariantNatTableOperations=True
							End If
						'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
						Case "Block"
							For iCounter=0 to 2
								StrCurrentOutput=JavaWindow("Collaborative Product").JavaEdit("VariantNatTableOutput").GetROProperty("value")
								If instr(1,StrCurrentOutput,StrMessage) Then
									Exit for
								Else
									ObjVariantNatTable.Click iX, iY, "LEFT"
									wait 1
								End If
							Next
							If instr(1,StrCurrentOutput,StrMessage) Then
								Fn_SISW_CPD_VariantNatTableOperations=True
							End If
					End Select
				End If
				If StrTabName<>"" Then
					Call Fn_TabFolder_Operation("DoubleClickTab", StrTabName, "")
					wait 2
				End If
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			Case "Save"
				'Save the changes
				Fn_SISW_CPD_VariantNatTableOperations=Fn_ToolbatButtonClick("Save the current contents (Ctrl+S)")
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			Case "VerifyOutput"
				StrCurrentOutput=JavaWindow("Collaborative Product").JavaEdit("VariantNatTableOutput").GetROProperty("value")
				If instr(1,StrCurrentOutput,StrMessage) Then
					Fn_SISW_CPD_VariantNatTableOperations=True
				End if
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			Case "VerifyNode"
				aNode=Split(StrNode,"~")
				For iCounter=0 to ubound(aNode)
					If Environment.Value("ProductName") = sQTPProductName OR Environment.Value("ProductName") = sUFTProductName Then	' added code by Vivek(16-dec-2014) to handle UFT related issue as some method not supported in UFT 
						iRowIndex=Fn_SISW_CPD_NatTable_TreeTable_GetRowIndex(ObjVariantNatTable, StrColumnName, aNode(iCounter), 0, "", "","")
					Else
						iRowIndex=Fn_SISW_CPD_NatTable_TreeTable_GetRowIndex(ObjVariantNatTable, StrColumnName, aNode(iCounter), 1, "", "","")
					End If
					If iRowIndex = -1 OR iRowIndex = False Then
						Fn_SISW_CPD_VariantNatTableOperations=False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Node [ "+aNode(iCounter)+" ] not found in tree" )
						Exit for
					Else
						Fn_SISW_CPD_VariantNatTableOperations=True
					End If
				Next
		End Select
		'Releasing object of [ Applicability ] table
		Set ObjVariantNatTable=Nothing
End Function
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------


'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		 :	Fn_SISW_CPD_ModelCloneandRealizationOptions
'
'Description			 :	Function Used to perform operations on Model clone and realization dialog
'
'Parameters			   :   		 sAction - > Action to be done ( Set, Verify etc)
'								 sActionToPerform -> Action to Perform (eg. Clone of Partition Breakdown, Realization of Partition Breakdown etc )		
'							     sSelectPartionSchemes - > Names of the Partition Schemes to be selected Separated by ~'		-> Also name with Description of ModelReuse Design Element seperated by ~													
'								 sPartitionRevRule - > Select Partition Revision Rule
'								 sSelectSourceContent -> Name of Source Content (eg. Select Partitions in Partition Breakdown etc )	
'								 sOtherOptions  -> Other options to be specified 
'                                sReserv        -> Reserved for future use
'
'Return Value		   : 	True Or False
'
'Pre-requisite			:	 The Model clone and realization dialog should be envoked
'
'Examples				:    breturn = Fn_SISW_CPD_ModelCloneandRealizationOptions("Set","Realization of Partition Breakdown","Select Partitions in Partition Breakdown","Partition Scheme Functional","Working(Current User); Any Status","","")
'							 bReturn = Fn_SISW_CPD_ModelCloneandRealizationOptions("OpenByName","Clone of Design Elements~Select a Subset Definition~Open the Subset Definition by Name","SubsetDefinition123","","","Apply effectivity based on source Design Elements","")
'							 bReturn = Fn_SISW_CPD_ModelCloneandRealizationOptions("OpenByName","Open the Subset Definition by Name","SubsetDefinition123","","","Apply effectivity based on source Design Elements","")
'							 bReturn = Fn_SISW_CPD_ModelCloneandRealizationOptions("ModelCloneandInstantiation","Instantiation of Design Elements","ModelReuseDE_Name~ModelReuseDE_Description","","","","")
'History					 :			
'		Developer Name		Date		Rev. No.	 Changes Done																Reviewer
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'		Ankit Tewari	02-Sept-2014	  1.0		Developed			              												
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'		Nilima Pandit	06-Nov-2015		  1.1		Added Case "ModelCloneandInstantiation" as per design change	[TC1121-2015102600-06_11_2015-VivekA-NewDevelopment]
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'		Shweta Rathod	18-Nov-2015		  1.1		Added Case "OpenByName" 										[TC1121-2015102600-18_11_2015-VivekA-NewDevelopment]
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'		Ankit Nigam  	30-Nov-2015		  1.1		Modified Case "ModelCloneandInstantiation" for description		[TC1121-2015110900-30_11_2015-AnkitN-NewDevelopment]
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Public Function Fn_SISW_CPD_ModelCloneandRealizationOptions(sAction,sActionToPerform,sSelectSourceContent,sSelectPartionSchemes,sPartitionRevRule,sOtherOptions,sSourcePartition)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_CPD_ModelCloneandRealizationOptions"
    Dim objModClneRealDlg,bReturn,aSelectPartionSchemes,iCount
	Dim iCounter,bFlag, Iterator, sPartition,aOtherOptions,aSourcePartition
    Dim sRevRule, objOpenByName, objTable, iRowCount, aActionToPerform
    Dim sName, sDescription, sOldActionToPerform
    bFlag = false
   	Fn_SISW_CPD_ModelCloneandRealizationOptions = False
   	Set objModClneRealDlg =Fn_SISW_CPD_GetObject("Modelcontentcloneandinstantiation")
   	'Check the existance of dialog
	If objModClneRealDlg.Exist(4) Then
		bFlag =  True
	Else
		Set objModClneRealDlg =Fn_SISW_CPD_GetObject("Modelcloneandrealization")
		If objModClneRealDlg.Exist(2) Then
			bFlag =  True
		End If		
	End If

	'Check the Existence of the TargetModelCarryoverOptions window, if not Exist the Function will be Terminated
	If bFlag = False Then
            Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_CPD_ModelCloneandRealizationOptions ] Model Clone Dialog not Exists")
			Set objModClneRealDlg = Nothing
			Exit Function
	End IF

	' ----------Modified by Chaitali R.----------------	
	If sActionToPerform <> "" Then
		sOldActionToPerform = sActionToPerform
		sActionToPerform = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("CPD_DisplayName"),sActionToPerform)
		If sActionToPerform = False Then
			sActionToPerform = sOldActionToPerform
		End If
	End If

	Select Case sAction
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "Set"

			'Select the RadioButton of the Action to Perform
			If sActionToPerform<>"" Then
				'Modify the attached text of the JavaRadioButton("ActionToPerform") to the option to be selected
				objModClneRealDlg.JavaRadioButton("ActionToPerform").SetTOProperty "attached text", sActionToPerform
				Wait 1
				' Set the Radiobutton ON
				bReturn = Fn_UI_JavaRadioButton_SetON("Fn_SISW_CPD_ModelCloneandRealizationOptions",objModClneRealDlg,"ActionToPerform")
				If bReturn = False Then
					 Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_CPD_ModelCloneandRealizationOptions ] Failed to Set The "+sActionToPerform+ "Radio Button")
					 Set objModClneRealDlg = Nothing
					 Exit Function
				End If
			End If
			Call Fn_ReadyStatusSync(1)

			If  objModClneRealDlg.JavaButton("Next").GetROProperty("enabled") Then
				bReturn = Fn_Button_Click("Fn_SISW_CPD_ModelCloneandRealizationOptions",objModClneRealDlg,"Next" )
				If bReturn = False Then
					 Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_CPD_ModelCloneandRealizationOptions ] Failed to Click [ Next ] button ")
					 Set objModClneRealDlg = Nothing
					 Exit Function
				End If
				Call Fn_ReadyStatusSync(2)
			End If
		
		If sSelectSourceContent<>"" Then
			'Modify the attached text Property of the JavaCheckBox("SelectPartitionSchemes") to the Schemes to be selected
			objModClneRealDlg.JavaRadioButton("SelectSourceContent").SetTOProperty "attached text", sSelectSourceContent
			Wait 1
			If objModClneRealDlg.JavaRadioButton("SelectSourceContent").Exist(3) Then
				' Set the Checkboxes ON
				bReturn = Fn_UI_JavaRadioButton_SetON("Fn_SISW_CPD_ModelCloneandRealizationOptions",objModClneRealDlg,"SelectSourceContent")
				If bReturn = False Then
					 Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_CPD_ModelCloneandRealizationOptions ] Failed to Set The "+sSelectSourceContent+ "Radio Button")
					 Set objModClneRealDlg = Nothing
					 Exit Function
				End If
			End If
			    ' Next
		End If
					
		If  objModClneRealDlg.JavaButton("Next").GetROProperty("enabled") Then
			bReturn = Fn_Button_Click("Fn_SISW_CPD_ModelCloneandRealizationOptions",objModClneRealDlg,"Next" )
			If bReturn = False Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_CPD_ModelCloneandRealizationOptions ] Failed to Click [ Next ] button ")
				Set objModClneRealDlg = Nothing
				Exit Function
				End If
			Call Fn_ReadyStatusSync(2)
		End If
		
		If sPartitionRevRule<>"" Then
			iCount=objModClneRealDlg.JavaList("PartitionRevisionRule").GetROProperty("items count")
			For iCounter=0 to iCount-1 
				sRevRule = objModClneRealDlg.JavaList("PartitionRevisionRule").GetItem("#"+cStr(iCounter))
				If Trim(sRevRule) = Trim(sPartitionRevRule) Then
					objModClneRealDlg.JavaList("PartitionRevisionRule").Select sRevRule
					Exit For
				End If
			Next
			Wait 1
		End If
		
		If  objModClneRealDlg.JavaButton("Next").GetROProperty("enabled") Then
			bReturn = Fn_Button_Click("Fn_SISW_CPD_ModelCloneandRealizationOptions",objModClneRealDlg,"Next" )
			If bReturn = False Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_CPD_ModelCloneandRealizationOptions ] Failed to Click [ Next ] button ")
				Set objModClneRealDlg = Nothing
				Exit Function
				End If
			Call Fn_ReadyStatusSync(2)
		End If
		
		'Select the Partition Schemes, MAke the Checkboxes ON
		If sSelectPartionSchemes<>"" Then
			'Split the multiple PartitionSchemes
			aSelectPartionSchemes = Split(sSelectPartionSchemes,"~",-1,1)
			For iCount =  0 to UBound(aSelectPartionSchemes)
				'Modify the attached text Property of the JavaCheckBox("SelectPartitionSchemes") to the Schemes to be selected
				objModClneRealDlg.JavaCheckBox("SelectPartitionSchemes").SetTOProperty "attached text", aSelectPartionSchemes(iCount)
				Wait 1
				If objModClneRealDlg.JavaCheckBox("SelectPartitionSchemes").Exist(3) Then
					' Set the Checkboxes ON
					Wait 1
					bReturn = Fn_CheckBox_Set("Fn_SISW_CPD_ModelCloneandRealizationOptions",objModClneRealDlg, "SelectPartitionSchemes","ON")
					If bReturn = False Then
						 Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_CPD_ModelCloneandRealizationOptions ] Failed to Set The "+sActionToPerform+ "Radio Button")
						 Set objModClneRealDlg = Nothing
						 Exit Function
					End If
				Else
					objModClneRealDlg.JavaRadioButton("ActionToPerform").SetTOProperty "attached text", aSelectPartionSchemes(iCount)
					Wait 1
					' Set the Radiobutton ON
					bReturn = Fn_UI_JavaRadioButton_SetON("Fn_SISW_CPD_ModelCloneandRealizationOptions",objModClneRealDlg,"ActionToPerform")
					If bReturn = False Then
						 Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_CPD_ModelCloneandRealizationOptions ] Failed to Set The "+sActionToPerform+ "Radio Button")
						 Set objModClneRealDlg = Nothing
						 Exit Function
					End If
				End If
		     Next
		End If
		Call Fn_ReadyStatusSync(4)
		wait 4
		
		If sSourcePartition<>"" Then
			If  objModClneRealDlg.JavaButton("Next").GetROProperty("enabled") Then
				bReturn = Fn_Button_Click("Fn_SISW_CPD_ModelCloneandRealizationOptions",objModClneRealDlg,"Next" )
				If bReturn = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_CPD_ModelCloneandRealizationOptions ] Failed to Click [ Next ] button ")
					Set objModClneRealDlg = Nothing
					Exit Function
				End If
				Call Fn_ReadyStatusSync(5)
			End If
			wait 6
			aSourcePartition = Split(sSourcePartition,"~",-1,1)
			For iCount = 0 To UBound(aSourcePartition) 
				bFlag = False
				For Iterator = 0 To objModClneRealDlg.JavaTree("SourcePartitions").Object.itemCount
					sPartition = objModClneRealDlg.JavaTree("SourcePartitions").Object.getItem(Iterator).getData().toString()
					If sPartition= aSourcePartition(iCount) Then
						wait 4
						objModClneRealDlg.JavaTree("SourcePartitions").object.getitem(Iterator).setChecked(true)
						wait 4
						Call Fn_UI_ClickJavaTreeCell("Fn_SISW_CPD_ModelCloneandRealizationOptions", objModClneRealDlg, "SourcePartitions", sPartition, "Object", "LEFT")
						If err.number < 0 Then
							 Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_CPD_ModelCloneandRealizationOptions ] Failed to Set The "+sPartition+ "CheckBox to ON")
							 Set objModClneRealDlg = Nothing						
							 Exit Function
						End If
						bFlag = True
						Exit for
					End If   
				Next
				If bFlag = False Then
			 		Set objModClneRealDlg = Nothing
				    	Exit Function
			 	End If
			Next				
		End If

		'Select the Other Options if Required
		If  sOtherOptions <> "" Then
			'To click on the next button
			If  objModClneRealDlg.JavaButton("Next").GetROProperty("enabled") Then
				bReturn = Fn_Button_Click("Fn_SISW_CPD_ModelCloneandRealizationOptions",objModClneRealDlg,"Next" )
				If bReturn = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_CPD_ModelCloneandRealizationOptions ] Failed to Click [ Next ] button ")
					Set objModClneRealDlg = Nothing
					Exit Function
					End If
				Call Fn_ReadyStatusSync(2)
			End If	
			aOtherOptions = Split(sOtherOptions,"~",-1,1)
			For iCount = 0 To UBound(aOtherOptions)		
				'To set the check box from Targetr model configurration	
	    		objModClneRealDlg.JavaCheckBox("SelectPartitionSchemes").SetTOProperty "attached text", aOtherOptions(iCount)
	    		Wait 1
				If objModClneRealDlg.JavaCheckBox("SelectPartitionSchemes").Exist(3) Then
					bReturn = Fn_CheckBox_Set("Fn_SISW_CPD_ModelCloneandRealizationOptions",objModClneRealDlg, "SelectPartitionSchemes","ON")
					If bReturn=False Then
						 Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_CPD_ModelCloneandRealizationOptions ] Failed to Set The "+sActionToPerform+ "CheckBox to ON")
						 Set objModClneRealDlg = Nothing						
						 Exit Function
					End If
				End If   
			Next		
		End if

		Call Fn_ReadyStatusSync(1)

		'Complete the Action by Clicking the Finish button
		bReturn = Fn_Button_Click("Fn_SISW_CPD_ModelCloneandRealizationOptions",objModClneRealDlg,"Finish" )
		If bReturn = False Then
			 Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_CPD_ModelCloneandRealizationOptions ] Failed to Click [ Finish ] button ")
			 Set objModClneRealDlg = Nothing
			 Exit Function
		End If
		Call Fn_ReadyStatusSync(1)
		
		'[TC1121-2015102600-06_11_2015-VivekA-NewDevelopment] - Added by Nilima Pandit
		Case "ModelCloneandInstantiation"
			'Select the RadioButton of the Action to Perform
			If sActionToPerform<>"" Then
				'Modify the attached text of the JavaRadioButton("ActionToPerform") to the option to be selected
				objModClneRealDlg.JavaRadioButton("ActionToPerform").SetTOProperty "attached text", sActionToPerform
				Wait 1
				'Set the Radiobutton ON
				bReturn = Fn_UI_JavaRadioButton_SetON("Fn_SISW_CPD_ModelCloneandRealizationOptions",objModClneRealDlg,"ActionToPerform")
				If bReturn = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_CPD_ModelCloneandRealizationOptions ] Failed to Set The "+sActionToPerform+ "Radio Button")
					Set objModClneRealDlg = Nothing
					Exit Function
				End If
			End If
			Call Fn_ReadyStatusSync(1)

			If objModClneRealDlg.JavaButton("Next").GetROProperty("enabled") Then
				bReturn = Fn_Button_Click("Fn_SISW_CPD_ModelCloneandRealizationOptions",objModClneRealDlg,"Next" )
				If bReturn = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_CPD_ModelCloneandRealizationOptions ] Failed to Click [ Next ] button ")
					Set objModClneRealDlg = Nothing
					Exit Function
				End If
				Call Fn_ReadyStatusSync(2)
			End If
		
			'Select the Revision rule from the JavaList
			If sPartitionRevRule<>"" Then
				iCount=objModClneRealDlg.JavaList("PartitionRevisionRule").GetROProperty("items count")
				For iCounter=0 to iCount-1 
					sRevRule = objModClneRealDlg.JavaList("PartitionRevisionRule").GetItem("#"+cStr(iCounter))
					If Trim(sRevRule) = Trim(sPartitionRevRule) Then
						objModClneRealDlg.JavaList("PartitionRevisionRule").Select sRevRule
						Exit For
					End If
				Next
				Wait 1
			End If
		
			If objModClneRealDlg.JavaButton("Next").GetROProperty("enabled") Then
				bReturn = Fn_Button_Click("Fn_SISW_CPD_ModelCloneandRealizationOptions",objModClneRealDlg,"Next" )
				If bReturn = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_CPD_ModelCloneandRealizationOptions ] Failed to Click [ Next ] button ")
					Set objModClneRealDlg = Nothing
					Exit Function
				End If
				Call Fn_ReadyStatusSync(2)
			End If
		
			If sSelectSourceContent<>"" Then
				If InStr( sSelectSourceContent ,"~" ) > 0 Then
					sSelectSourceContent = Split(sSelectSourceContent , "~" )
					sName = sSelectSourceContent(0)
					sDescription = sSelectSourceContent(1)
				Else
					sName = sSelectSourceContent
					sDescription = ""					
				End If
				'Enter the Name of the ModelReuse Design Element
				If objModClneRealDlg.JavaEdit("Name").Exist(3) Then
					objModClneRealDlg.JavaEdit("Name").Type sName
					Call Fn_ReadyStatusSync(1)
				End If
				'Enter the Description of the ModelReuse Design Element
				If sDescription <> "" Then
					If objModClneRealDlg.JavaEdit("Description").Exist(3) Then
						objModClneRealDlg.JavaEdit("Description").Type sDescription
						Call Fn_ReadyStatusSync(1)
					End If					
				End If
				
				'Modify the attached text Property of the JavaCheckBox("SelectPartitionSchemes") to the Schemes to be selected
				objModClneRealDlg.JavaRadioButton("SelectSourceContent").SetTOProperty "attached text", sName
				Wait 1
				If objModClneRealDlg.JavaRadioButton("SelectSourceContent").Exist(3) Then
					' Set the Checkboxes ON
					bReturn = Fn_UI_JavaRadioButton_SetON("Fn_SISW_CPD_ModelCloneandRealizationOptions",objModClneRealDlg,"SelectSourceContent")
					If bReturn = False Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_CPD_ModelCloneandRealizationOptions ] Failed to Set The "+sName+ "Radio Button")
						Set objModClneRealDlg = Nothing
						Exit Function
					End If
				End If
			    ' Next
			End If
	
			If objModClneRealDlg.JavaButton("Next").GetROProperty("enabled") Then
				bReturn = Fn_Button_Click("Fn_SISW_CPD_ModelCloneandRealizationOptions",objModClneRealDlg,"Next" )
				If bReturn = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_CPD_ModelCloneandRealizationOptions ] Failed to Click [ Next ] button ")
					Set objModClneRealDlg = Nothing
					Exit Function
				End If
				Call Fn_ReadyStatusSync(2)
			End If
		
			'Select the Partition Schemes, MAke the Checkboxes ON
			If sSelectPartionSchemes<>"" Then
				'Split the multiple PartitionSchemes
				aSelectPartionSchemes = Split(sSelectPartionSchemes,"~",-1,1)
				For iCount =  0 to UBound(aSelectPartionSchemes)
					'Modify the attached text Property of the JavaCheckBox("SelectPartitionSchemes") to the Schemes to be selected
					objModClneRealDlg.JavaCheckBox("SelectPartitionSchemes").SetTOProperty "attached text", aSelectPartionSchemes(iCount)
					Wait 1
					If objModClneRealDlg.JavaCheckBox("SelectPartitionSchemes").Exist(3) Then
						' Set the Checkboxes ON
						bReturn = Fn_CheckBox_Set("Fn_SISW_CPD_ModelCloneandRealizationOptions",objModClneRealDlg, "SelectPartitionSchemes","ON")
						If bReturn = False Then
						 	Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_CPD_ModelCloneandRealizationOptions ] Failed to Set The "+sActionToPerform+ "Radio Button")
						 	Set objModClneRealDlg = Nothing
						 	Exit Function
						End If
					Else
						objModClneRealDlg.JavaRadioButton("ActionToPerform").SetTOProperty "attached text", aSelectPartionSchemes(iCount)
						Wait 1
						' Set the Radiobutton ON
						bReturn = Fn_UI_JavaRadioButton_SetON("Fn_SISW_CPD_ModelCloneandRealizationOptions",objModClneRealDlg,"ActionToPerform")
						If bReturn = False Then
						 	Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_CPD_ModelCloneandRealizationOptions ] Failed to Set The "+sActionToPerform+ "Radio Button")
						 	Set objModClneRealDlg = Nothing
						 	Exit Function
						End If
					End If
		     	Next
			End If
			Call Fn_ReadyStatusSync(4)
			wait 4
		
			If sSourcePartition<>"" Then
				If objModClneRealDlg.JavaButton("Next").GetROProperty("enabled") Then
					bReturn = Fn_Button_Click("Fn_SISW_CPD_ModelCloneandRealizationOptions",objModClneRealDlg,"Next" )
					If bReturn = False Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_CPD_ModelCloneandRealizationOptions ] Failed to Click [ Next ] button ")
						Set objModClneRealDlg = Nothing
						Exit Function
					End If
					Call Fn_ReadyStatusSync(5)
				End If
				wait 6
				aSourcePartition = Split(sSourcePartition,"~",-1,1)
				For iCount = 0 To UBound(aSourcePartition) 
					bFlag = False
					For Iterator = 0 To objModClneRealDlg.JavaTree("SourcePartitions").Object.itemCount
						sPartition = objModClneRealDlg.JavaTree("SourcePartitions").Object.getItem(Iterator).getData().toString()
						If sPartition= aSourcePartition(iCount) Then
							wait 2
							objModClneRealDlg.JavaTree("SourcePartitions").object.getitem(Iterator).setChecked(true)
							wait 2
							Call Fn_UI_ClickJavaTreeCell("Fn_SISW_CPD_ModelCloneandRealizationOptions", objModClneRealDlg, "SourcePartitions", sPartition, "Object", "LEFT")
							If err.number < 0 Then
							 	Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_CPD_ModelCloneandRealizationOptions ] Failed to Set The "+sPartition+ "CheckBox to ON")
							 	Set objModClneRealDlg = Nothing						
							 	Exit Function
							End If
							bFlag = True
							Exit for
						End If   
					Next
					If bFlag = False Then
				 		Set objModClneRealDlg = Nothing
				   		Exit Function
				 	End If
				Next				
			End If

			'Select the Other Options if Required
			If sOtherOptions <> "" Then
				''To click on the next button
				If objModClneRealDlg.JavaButton("Next").GetROProperty("enabled") Then
					bReturn = Fn_Button_Click("Fn_SISW_CPD_ModelCloneandRealizationOptions",objModClneRealDlg,"Next" )
					If bReturn = False Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_CPD_ModelCloneandRealizationOptions ] Failed to Click [ Next ] button ")
						Set objModClneRealDlg = Nothing
						Exit Function
					End If
					Call Fn_ReadyStatusSync(2)
				End If	
				aOtherOptions = Split(sOtherOptions,"~",-1,1)
				For iCount = 0 To UBound(aOtherOptions)
					''To set the check box 	from Targetr model configurration	
	    			objModClneRealDlg.JavaCheckBox("SelectPartitionSchemes").SetTOProperty "attached text", aOtherOptions(iCount)
	    			Wait 1
					If objModClneRealDlg.JavaCheckBox("SelectPartitionSchemes").Exist(3) Then
						bReturn = Fn_CheckBox_Set("Fn_SISW_CPD_ModelCloneandRealizationOptions",objModClneRealDlg, "SelectPartitionSchemes","ON")
						If bReturn=False Then
							 Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_CPD_ModelCloneandRealizationOptions ] Failed to Set The "+sActionToPerform+ "CheckBox to ON")
							 Set objModClneRealDlg = Nothing						
							 Exit Function
						End If
					End If   
				Next		
			End if

			Call Fn_ReadyStatusSync(1)
			'Complete the Action by Clicking the Finish button
			bReturn = Fn_Button_Click("Fn_SISW_CPD_ModelCloneandRealizationOptions",objModClneRealDlg,"Finish" )
			If bReturn = False Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_CPD_ModelCloneandRealizationOptions ] Failed to Click [ Finish ] button ")
				Set objModClneRealDlg = Nothing
			 	Exit Function
			End If
			Call Fn_ReadyStatusSync(1)	
		'- - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		'[TC1121-2015102600-18_11_2015-VivekA-NewDevelopment] - Added by Shweta Rathod
		Case "OpenByName"
			If Instr(sActionToPerform,"~")>0 Then
				aActionToPerform = Split(sActionToPerform,"~")
				'sActionToPerform = aActionToPerform(0)
				'Select the RadioButton of the Action to Perform
				If aActionToPerform(0)<>"" Then
					'Modify the attached text of the JavaRadioButton("ActionToPerform") to the option to be selected
					objModClneRealDlg.JavaRadioButton("ActionToPerform").SetTOProperty "attached text", aActionToPerform(0)
					Wait 1
					'Set the Radiobutton ON
					bReturn = Fn_UI_JavaRadioButton_SetON("Fn_SISW_CPD_ModelCloneandRealizationOptions",objModClneRealDlg,"ActionToPerform")
					If bReturn = False Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_CPD_ModelCloneandRealizationOptions ] Failed to Set The "+aActionToPerform(0)+ "Radio Button")
						Set objModClneRealDlg = Nothing
						Exit Function
					End If
				End If
				Wait 2
	
				If objModClneRealDlg.JavaButton("Next").GetROProperty("enabled") Then
					bReturn = Fn_Button_Click("Fn_SISW_CPD_ModelCloneandRealizationOptions",objModClneRealDlg,"Next" )
					If bReturn = False Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_CPD_ModelCloneandRealizationOptions ] Failed to Click [ Next ] button ")
						Set objModClneRealDlg = Nothing
						Exit Function
					End If
					Wait 2
				End If
			End If
			
			If Instr(sActionToPerform,"~")>0 Then
				If UBound(aActionToPerform)>1 Then
					If aActionToPerform(0)="Clone of Design Components" Then
						'Modify the attached text of the JavaRadioButton("ActionToPerform") to the option to be selected
						objModClneRealDlg.JavaRadioButton("ActionToPerform").SetTOProperty "attached text", aActionToPerform(1)
						Wait 1
						'Set the Radiobutton ON
						bReturn = Fn_UI_JavaRadioButton_SetON("Fn_SISW_CPD_ModelCloneandRealizationOptions",objModClneRealDlg,"ActionToPerform")
						If bReturn = False Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_CPD_ModelCloneandRealizationOptions ] Failed to Set The "+aActionToPerform(1)+ "Radio Button")
							Set objModClneRealDlg = Nothing
							Exit Function
						End If
						Wait 1
					End If
					sActionToPerform = aActionToPerform(2)
				Else
					sActionToPerform = aActionToPerform(1)
				End If
			End If
			objModClneRealDlg.JavaToolbar("ToolBar").Press sActionToPerform
			Wait 5
			Set objOpenByName = objModClneRealDlg.JavaWindow("Shell").JavaWindow("OpenByName")
			' typeing value in Name edit box
			If sSelectSourceContent<> ""  Then
				objOpenByName.JavaEdit("Name").Set ""
				Call Fn_UI_EditBox_Type("Fn_SISW_CPD_ModelCloneandRealizationOptions",objOpenByName,"Name",sSelectSourceContent)
				Wait 1
			End If
		
			Call  Fn_Button_Click("Fn_SISW_CPD_ModelCloneandRealizationOptions",objOpenByName,"Find")
			wait(5)
			Set objTable = objOpenByName.JavaTable("IDTable")
			iRowCount = cint(objTable.GetROProperty("rows"))
			
			if iRowCount = 0 then
				objOpenByName.Activate
				wait 2
				Call  Fn_Button_Click("Fn_SISW_CPD_ModelCloneandRealizationOptions",objOpenByName,"Find")
			End If
			
			wait 2
			iRowCount = cint(objTable.GetROProperty("rows"))
			For iCount = 0 to iRowCount -1
				bReturn = False
				If sSelectSourceContent <> "" Then
					If instr(objTable.GetCellData(iCount,"Object"), sSelectSourceContent) > 0 Then bReturn = True
				ElseIf objTable.GetCellData(iCount,"Object") = sSelectSourceContent  Then
					 bReturn = True
				End If
				If bReturn = True Then
					Call Fn_UI_JavaTable_SelectRow("Fn_SISW_CPD_ModelCloneandRealizationOptions", objOpenByName,"IDTable",iCount)
					Exit for
				End If
				Wait 1
			Next
			iCount = cint(objTable.GetROProperty("SelectedRow"))
			If iCount <> -1 Then
				Wait(5)
				objTable.DoubleClickCell iCount,"Object","LEFT"
				Call Fn_ReadyStatusSync(5)
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_CPD_ModelCloneandRealizationOptions ] Case [ " & sAction  & " ] No Item is selected.")
				Fn_SISW_CPD_ModelCloneandRealizationOptions = False
				Exit function
			End If
			
			If sSelectPartionSchemes<>"" Then
				If objModClneRealDlg.JavaButton("Next").GetROProperty("enabled") Then
					bReturn = Fn_Button_Click("Fn_SISW_CPD_ModelCloneandRealizationOptions",objModClneRealDlg,"Next" )
					If bReturn = False Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_CPD_ModelCloneandRealizationOptions ] Failed to Click [ Next ] button ")
						Set objModClneRealDlg = Nothing
						Exit Function
					End If
					Call Fn_ReadyStatusSync(2)
				End If
				
				If Instr(sSelectPartionSchemes,"~")>0 Then
					aSelectPartionSchemes = Split(sSelectPartionSchemes,"~")
					sName = aSelectPartionSchemes(0)
					sDescription = aSelectPartionSchemes(1)
				Else
					sName = aSelectPartionSchemes
					sDescription = ""
				End If
				'Enter the Name of the ModelReuse Design Element
				If objModClneRealDlg.JavaEdit("Name").Exist(3) Then
					objModClneRealDlg.JavaEdit("Name").Type sName
					Call Fn_ReadyStatusSync(1)
				End If
				If sDescription <> "" Then
					If objModClneRealDlg.JavaEdit("Description").Exist(3) Then
						objModClneRealDlg.JavaEdit("Description").Type sDescription
						Call Fn_ReadyStatusSync(1)
					End If					
				End If
			End If

			'Select the Other Options if Required
			If sOtherOptions <> "" Then
				''To click on the next button
				If objModClneRealDlg.JavaButton("Next").GetROProperty("enabled") Then
					bReturn = Fn_Button_Click("Fn_SISW_CPD_ModelCloneandRealizationOptions",objModClneRealDlg,"Next" )
					If bReturn = False Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_CPD_ModelCloneandRealizationOptions ] Failed to Click [ Next ] button ")
						Set objModClneRealDlg = Nothing
						Exit Function
					End If
					Call Fn_ReadyStatusSync(2)
				End If
				If sOtherOptions<>"Next" Then
					aOtherOptions = Split(sOtherOptions,"~",-1,1)
					For iCount = 0 To UBound(aOtherOptions)
						''To set the check box 	from Targetr model configurration	
						objModClneRealDlg.JavaCheckBox("SelectPartitionSchemes").SetTOProperty "attached text", aOtherOptions(iCount)
						Wait 1
						If objModClneRealDlg.JavaCheckBox("SelectPartitionSchemes").Exist(3) Then
							bReturn = Fn_CheckBox_Set("Fn_SISW_CPD_ModelCloneandRealizationOptions",objModClneRealDlg, "SelectPartitionSchemes","ON")
							If bReturn=False Then
								 Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_CPD_ModelCloneandRealizationOptions ] Failed to Set The "+sActionToPerform+ "CheckBox to ON")
								 Set objModClneRealDlg = Nothing						
								 Exit Function
							End If
						End If   
					Next
				End If
			End If
			If sOtherOptions<>"Next" Then
				bReturn = Fn_Button_Click("Fn_SISW_CPD_ModelCloneandRealizationOptions",objModClneRealDlg,"Finish" )
				If bReturn = False Then
					 Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_CPD_ModelCloneandRealizationOptions ] Failed to Click [ Finish ] button ")
					 Set objModClneRealDlg = Nothing
					 Exit Function
				End If
				Call Fn_ReadyStatusSync(1)	
			End If				
			Fn_SISW_CPD_ModelCloneandRealizationOptions = True
	 '- - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
	Case Else
		Fn_SISW_CPD_ModelCloneandRealizationOptions = False
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_CPD_ModelCloneandRealizationOptions ] Invalid case [ " & sAction & " ].")
		Set objModClneRealDlg = Nothing
		Exit Function		
	End Select

	Fn_SISW_CPD_ModelCloneandRealizationOptions = True
	Set objModClneRealDlg = Nothing
End Function

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_CPD_ContentExplr_GetNodePathFromColumnValue

'Description			 :	Function Used to perform operations on Variant Nat Tables

'Parameters			   :   1.objTree: Root Node object  
'						   2.sColumnName: Column Name
'						   3.sValue: Column value
'						   4.path: path of root node 

'Return Value		    : 	path of node

'Pre-requisite			:	Content Explorer tree should be appear

'Examples				:   bReturn=Fn_CPD_ContentExplr_GetNodePathFromColumnValue(JavaWindow("Collaborative Product").JavaTree("NavTree").Object.getItem(0),"Source Object Name", "OFFICE","#0")				
'							
'History			    :			
'							Developer Name						Date					Rev. No.		Reviewer Name				Changes Done	
'
' 							Shweta Rathod					10-Sep-2014					1.0											Developed
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_CPD_ContentExplr_GetNodePathFromColumnValue(objTree,sColumnName, sValue,path)
	GBL_FAILED_FUNCTION_NAME="Fn_CPD_ContentExplr_GetNodePathFromColumnValue"
	
	Dim Iterator

	For Iterator = 0 To cInt(objTree.getItemCount())-1
	'create path and compare
		
		If cInt(objTree.getItem(Iterator).getItemCount())>0 Then
			Set objTree = objTree.getItem(Iterator)	
			path = path &":#" & Cstr(Iterator)

			If JavaWindow("Collaborative Product").JavaTree("NavTree").GetColumnValue(path, sColumnName) = sValue Then
				path = path & "True"
				Fn_CPD_ContentExplr_GetNodePathFromColumnValue = path
				Exit Function
			End If
			call Fn_CPD_ContentExplr_GetNodePathFromColumnValue(objTree,sColumnName, svalue, path)
		End If
		If instr(path, "True") = 0 Then
			If Iterator = cInt(objTree.getItemCount() - 1) Then
				If JavaWindow("Collaborative Product").JavaTree("NavTree").GetColumnValue(path +":#" +cstr(Iterator), sColumnName) = sValue Then
					path = path &":#" & Cstr(Iterator) & "True"
					Fn_CPD_ContentExplr_GetNodePathFromColumnValue = path
					Exit Function
				Else
					'[TC1121-2015101900-04_11_2015-VivekA-Maintenance] - Added by Priyanka K, if the node note present in tree
					If JavaWindow("Collaborative Product").JavaTree("NavTree").Object.getTopItem().getData().tostring() = objTree.getData().toString() Then
						Fn_CPD_ContentExplr_GetNodePathFromColumnValue = False					
						Exit Function
						'-----------------------------------------------
					Else
						Set objTree = objTree.getParentItem()
						path = left(path, (len(path)-3))
						Exit Function					
					End If
				End If
			End If
		Else
			Fn_CPD_ContentExplr_GetNodePathFromColumnValue = path
			Exit Function	
		End If
		If instr(path, "True") = 0 Then 
			If JavaWindow("Collaborative Product").JavaTree("NavTree").GetColumnValue(path +":#" +cstr(Iterator), sColumnName) = sValue Then
				path = path &":#" & Cstr(Iterator)& "True"
				Fn_CPD_ContentExplr_GetNodePathFromColumnValue = path
				Exit Function
			End If
		Else
			Fn_CPD_ContentExplr_GetNodePathFromColumnValue = path
			Exit Function
		End If
	Next
'	Fn_CPD_ContentExplr_GetNodePathFromColumnValue = path
End Function

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_CPD_MarkupSpaceOperations
'
'Description		:	Function Used to Set and Get Markup Space
'
'Parameters			:  	1.sAction				: Action Name
'						2.sExistingMarkupSpace	: Existing Markup Space Name
'						3.sNewMarkupSpace		: New Markup Space to be selected.
'						4.sReserve				: Future Use
'
'Return Value		: 	True Or False
'
'Pre-requisite		:	CPD perspective should be activated.
'
'Examples			:   Call Fn_CPD_MarkupSpaceOperations("Set", "No Markup Space", "MS000001-MK1", "")
'Examples			:   Call Fn_CPD_MarkupSpaceOperations("Exist", "MS000001-MK1", "", "")
'Examples			:   Call Fn_CPD_MarkupSpaceOperations("ExistInMenu", "MS000001-MK1", "", "")
'Examples			:   Call Fn_CPD_MarkupSpaceOperations("Get", "", "", "")
'
'History			:			
'			Developer Name				Date			Rev. No.	  Changes Done			Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'			Vivek Ahirrao			08-Dec-2014			 1.0			Created
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Public Function Fn_CPD_MarkupSpaceOperations(sAction, sExistingMarkupSpace, sNewMarkupSpace, sReserve)
	GBL_FAILED_FUNCTION_NAME="Fn_CPD_MarkupSpaceOperations"
	'Declare Variables
	Dim objMarkupSpace, aMarkupSpace, iInstance,intNoOfObjects,iCnt,objSelectType,SIndex,SName,bFlag
	Set objMarkupSpace = JavaWindow("Collaborative Product").JavaObject("MarkupSpaceHyperlink")
	Fn_CPD_MarkupSpaceOperations = False
	
	'check which MarkupSpace is set
	If sAction = "Set" OR sAction = "Get" OR sAction = "ExistInMenu" Then
		bFlag = False
		Set objSelectType = Description.Create()
		objSelectType("Class Name").value = "JavaObject"
		objSelectType("toolkit class").value = "org.eclipse.ui.forms.widgets.ImageHyperlink"
		Set intNoOfObjects = JavaWindow("Collaborative Product").ChildObjects(objSelectType)
	'Code to check if the New Markup Space is already set.
		If sAction = "Set" Then
			For iCnt = 0 to intNoOfObjects.count-1
				If Trim(intNoOfObjects(iCnt).GetROProperty("text")) = Trim(sNewMarkupSpace) then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_CPD_MarkupSpaceOperations is already SET")
					Fn_CPD_MarkupSpaceOperations = True 
					Set objMarkupSpace = Nothing
					Set objSelectType = Nothing
					Set intNoOfObjects = Nothing
					Exit Function
				End If
			Next
		End If

	'Code to check if the "sExistingMarkupSpace" that we pass is actually set in the applciation
	'if not set then get the actual MarkupSpace set
		For iCnt = 0 to intNoOfObjects.count-1
			If Trim(intNoOfObjects(iCnt).GetROProperty("text")) = Trim(sExistingMarkupSpace) then
				bFlag = True
				Exit For
			End IF										
			If Instr(LCase(Trim(intNoOfObjects(iCnt).GetROProperty("text"))),LCase("MS")) > 0 then
				SName = Trim(intNoOfObjects(iCnt).GetROProperty("text"))
			ElseIf LCase(Trim(intNoOfObjects(iCnt).GetROProperty("text"))) = LCase(Trim("No Markup Space")) Then
				SName = Trim(intNoOfObjects(iCnt).GetROProperty("text"))
			End If
		Next
		If bFlag = False Then
			sExistingMarkupSpace = SName
		End IF
	End IF

	iInstance = 0
	Select Case sAction
	'Case for Set Markup space
		Case "Set"
				objMarkupSpace.SetTOProperty "text", sExistingMarkupSpace
				objMarkupSpace.SetTOProperty "Index", iInstance
				objMarkupSpace.Click 1,1,"LEFT"
				Wait 1
				Fn_CPD_MarkupSpaceOperations = Fn_UI_JavaMenu_Select("Fn_CPD_MarkupSpaceOperations",JavaWindow("Collaborative Product"),sNewMarkupSpace)
	'Case for getting selected markup space
		Case "Get"
				Fn_CPD_MarkupSpaceOperations = sExistingMarkupSpace		
	'Case for check existence of Markup
		Case "Exist"
				objMarkupSpace.SetTOProperty "text", sExistingMarkupSpace
				objMarkupSpace.SetTOProperty "Index", iInstance
				Fn_CPD_MarkupSpaceOperations = objMarkupSpace.Exist(5)
	''Case for check existence of Markup in menu 		
		Case "ExistInMenu"
				objMarkupSpace.SetTOProperty "text", sExistingMarkupSpace
				objMarkupSpace.SetTOProperty "Index", iInstance
				objMarkupSpace.Click 1,1,"LEFT"
				Wait 1
				Fn_CPD_MarkupSpaceOperations = JavaWindow("Collaborative Product").JavaMenu("Label:=" & sNewMarkupSpace).Exist(5)
				Call Fn_KeyBoardOperation("SendKeys", "{ESC}")
	'Invalid case
		Case Else
			Exit Function
	End Select
	
	If Fn_CPD_MarkupSpaceOperations <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_CPD_MarkupSpaceOperations : Executed successfully with Case [ " & sAction & " ] ")
	End IF
	'Reset objects
	objMarkupSpace.SetTOProperty "text", "No Markup Space"
	objMarkupSpace.SetTOProperty "Index",0	
	Set objMarkupSpace = Nothing
End Function

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_CPD_NewMarkupSpaceCreate
'
'Description		:	Function Used to Create New Markup Space
'
'Parameters			:  	1.sAction			: Action Name
'						2.sMarkupSpaceID	: Markup Space ID
'						3.sMarkupSpaceName	: Markup Space Name
'						4.sDescription		: Markup Space Description
'						5.sReserve			: Future Use
'
'Return Value		: 	MarkupSpaceID Or False
'
'Pre-requisite		:	CPD perspective should be activated.
'
'Examples			:   Call Fn_CPD_NewMarkupSpaceCreate("Create", "MS000022", "MarkupSpace1", "Markup Space Description")
'
'History			:			
'			Developer Name				Date			Rev. No.	  Changes Done			Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'			Vivek Ahirrao			08-Dec-2014			 1.0			Created
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Public Function Fn_CPD_NewMarkupSpaceCreate(sAction, sMarkupSpaceID, sMarkupSpaceName, sDescription, sReserve)
	GBL_FAILED_FUNCTION_NAME="Fn_CPD_NewMarkupSpaceCreate"
	'Decalre Variables
	Dim objMarkupSpace, bReturn, sExistingMarkupSpace
	Fn_CPD_NewMarkupSpaceCreate = False
	Set objMarkupSpace = JavaWindow("Collaborative Product").JavaWindow("NewBusinessObject")
	
	'Get selected Markup space 
	sExistingMarkupSpace =  Fn_CPD_MarkupSpaceOperations("Get", "No Change Context", "", "")
	'Open New Bussiness dialog if not exists
	If Fn_UI_ObjectExist("Fn_CPD_NewMarkupSpaceCreate",objMarkupSpace) = False Then
		bReturn =  Fn_CPD_MarkupSpaceOperations("Set", sExistingMarkupSpace, "New...", "")
		If bReturn =  False OR Fn_UI_ObjectExist("Fn_CPD_NewMarkupSpaceCreate",objMarkupSpace) = False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_CPD_NewMarkupSpaceCreate ] Failed to open Markup Space window.")
			Set objMarkupSpace = Nothing
			Exit function
		End If
	End If
	
	Select Case sAction
	'Case for create markup space
		Case "Create"
	'Select Markup Space from tree
				If objMarkupSpace.JavaTree("BusinessObjectType").exist(4) Then
					If Fn_UI_JavaTreeGetItemPathExt("", objMarkupSpace.JavaTree("BusinessObjectType"), "Complete List:Change Proposal", "", "") <> False Then
						objMarkupSpace.JavaTree("BusinessObjectType").Select "Complete List:Change Proposal"
					Else
						objMarkupSpace.JavaTree("BusinessObjectType").Select "Most Recently Used:Change Proposal"
					End IF
					Call Fn_Button_Click("Fn_CPD_NewMarkupSpaceCreate",objMarkupSpace,"Next" )
				End If
				
	'If MarkupSpaceID is empty
				objMarkupSpace.JavaStaticText("Field").SetTOProperty "label", "ID:"
				If sMarkupSpaceID = "" Then
					Call Fn_Button_Click("Fn_CPD_NewMarkupSpaceCreate",objMarkupSpace,"Assign")
					Call Fn_ReadyStatusSync(5)
					sMarkupSpaceID = objMarkupSpace.JavaEdit("Field").GetROProperty("value")
					If sMarkupSpaceID <> "" Then
						Fn_CPD_NewMarkupSpaceCreate = sMarkupSpaceID
					Else
						Exit Function
					End If
				Else
					Call Fn_Edit_Box("Fn_CPD_NewMarkupSpaceCreate",objMarkupSpace,"Field",sMarkupSpaceID)
					Fn_CPD_NewMarkupSpaceCreate = True
				End If
	'Set Markup space Name
				If sMarkupSpaceName <> "" Then
					objMarkupSpace.JavaStaticText("Field").SetTOProperty "label", "Name:"
					objMarkupSpace.JavaEdit("Field").Type sMarkupSpaceName
					Call Fn_ReadyStatusSync(5)
				End If
	'Set markup Description
				If sDescription <> "" Then
					objMarkupSpace.JavaStaticText("Field").SetTOProperty "label", "Description:"
					Call Fn_Edit_Box("Fn_CPD_NewMarkupSpaceCreate",objMarkupSpace,"Field",sDescription)
				End If
	'Click on Finish button
				Call Fn_Button_Click("Fn_CPD_NewMarkupSpaceCreate",objMarkupSpace,"Finish")
				Call Fn_ReadyStatusSync(5)
	'Click on Cancel button
				Call Fn_Button_Click("Fn_CPD_NewMarkupSpaceCreate",objMarkupSpace,"Cancel" )
		Case Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_CPD_NewMarkupSpaceCreate ] Invalid case [ " & sAction & " ].")
	End Select
	
	If Fn_CPD_NewMarkupSpaceCreate <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [ Fn_CPD_NewMarkupSpaceCreate ] executed successfuly with case [ " & sAction & " ].")
	End If
	Set objMarkupSpace = Nothing
End Function

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_SISW_CPD_ModelCloneandRealization_SourcePartitions

'Description			 :	Function Used to perform operations on JavaTree in ModelCloneDialog

'Parameters			   :   1.sAction: Action to perform
'						   2.sNodeNameWithPath: Full path of node in tree
'						   3.sButton : Click on provided button
'						   4.sFutureUse: For future use
'						

'Return Value		    : 	True or False

'Pre-requisite			:	ModelClone Dialog tree should be visible

'Examples				:  bReturn =  Fn_SISW_CPD_ModelCloneandRealization_SourcePartitions("SetCheckBox","PTN000027/001;1-B:PTN000028/001;1-C~PTN000027/001;1-B:PTN000028/001;1-C:PTN000029/001;1-C~PTN000031/001;1-v","Finish","")
'							
'History			    :			
'							Developer Name						Date					Rev. No.		Reviewer Name				Changes Done	
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' 							Anurag Khera					19-Dec-2014					1.0											Developed
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_SISW_CPD_ModelCloneandRealization_SourcePartitions(sAction,sNodeNameWithPath,sButton,sFutureUse)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_CPD_ModelCloneandRealization_SourcePartitions"
	'Declare Variables
	Dim intNodeCount,iCount,aNodePath,strTreenode,iCnt, iTreeIndex, aSourcePartition
	Dim objModelCloneandRealization, objJavaTree,bFlag
	
	bFlag = false
	Fn_SISW_CPD_ModelCloneandRealization_SourcePartitions = False
	'Set Object of Model Clone Dialog
	 Set objModelCloneandRealization =Fn_SISW_CPD_GetObject("Modelcontentcloneandinstantiation")
   	'Check the existance of dialog
	If objModelCloneandRealization.Exist(4) Then
		bFlag =  True
	Else
		Set objModelCloneandRealization =Fn_SISW_CPD_GetObject("Modelcloneandrealization")
		If objModelCloneandRealization.Exist(2) Then
			bFlag =  True
		End If		
	End If
	
	Set objJavaTree = objModelCloneandRealization.JavaTree("SourcePartitions")
	'Check the Existence of the TargetModelCarryoverOptions window, if not Exist the Function will be Terminated
	If bFlag = False Then
            Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_CPD_ModelCloneandRealizationOptions ] Model Clone Dialog not Exists")
			Set objModClneRealDlg = Nothing
			Exit Function
	End IF

	Select Case sAction
		'case to check existence of partition template
		Case "Exist"
			If objJavaTree.Exist Then
				iTreeIndex= Fn_UI_JavaTreeGetItemPathExt("Fn_SISW_CPD_ModelCloneandRealization_SourcePartitions", objJavaTree, sNodeNameWithPath , "", "")
				If iTreeIndex <> False Then
					Fn_SISW_CPD_ModelCloneandRealization_SourcePartitions = True
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_CPD_ModelCloneandRealization_SourcePartitions ] Failed to Exist "+sNodeNameWithPath+ "node in Source partition tree.")
				End If
			End If
		'case for Expand partition template
		Case "Expand"
			If objJavaTree.Exist Then
				iTreeIndex = Fn_UI_JavaTreeGetItemPathExt("Fn_SISW_CPD_ModelCloneandRealization_SourcePartitions", objJavaTree, sNodeNameWithPath, "", "")
				If iTreeIndex <> False Then
					objJavaTree.Expand iTreeIndex
					Fn_SISW_CPD_ModelCloneandRealization_SourcePartitions = True
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_CPD_ModelCloneandRealization_SourcePartitions ] Failed to Expand "+sNodeNameWithPath+ "node in Source Partition tree.")
				End iF
			End If	
	'Case for set Partition template checkbox			
		Case "SetCheckBox"
			 If objJavaTree.Exist Then
			 	aSourcePartition = Split(sNodeNameWithPath,"~",-1,1)
				For iCnt = 0 To UBound(aSourcePartition) 
	'Retrive tree index of node for selecting checkbox
					iTreeIndex = Fn_UI_JavaTreeGetItemPathExt("Fn_SISW_CPD_ModelCloneandRealization_SourcePartitions", objJavaTree, aSourcePartition(iCnt), "", "")
					If iTreeIndex <> False Then
						aNodePath = split(Replace(iTreeIndex,"#",""),":")
						For iCount = 0 To UBound(aNodePath)
							If iCount =0  Then
								Set strTreenode = objJavaTree.Object.getItem(aNodePath(iCount))
							Else
								Set strTreenode = strTreenode.getItem(aNodePath(iCount))
							End If
						Next
						strTreenode.setChecked(true)
						wait 1
						objJavaTree.Activate iTreeIndex
						Fn_SISW_CPD_ModelCloneandRealization_SourcePartitions = True
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_CPD_ModelCloneandRealization_SourcePartitions ] Failed to set "+sNodeNameWithPath+ "node ON in Source Partition tree.")
						Fn_SISW_CPD_ModelCloneandRealization_SourcePartitions = False
						Exit For
					End If
				Next
			End If
	'Case does not exist
		Case Else
			Exit Function
	End Select
	'Click on provided button
	If sButton <> "" Then
		Call Fn_SISW_UI_JavaButton_Operations("Fn_SISW_CPD_ModelCloneandRealization_SourcePartitions", "Click", objModelCloneandRealization, sButton)
	End If
	
	Set objModelCloneandRealization = nothing
	Set objJavaTree = nothing
End Function

'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@
'@@    Function Name			:	Fn_CPD_CompositeViewMenuOperations
'@@
'@@    Description				:	Function Used to perform operations on View Menu (Context Menu)
'@@
'@@    Parameters			    :	1. sAction		: Action to be performed
'@@								:	2. sMenuLabel	: Menu to be Selected
'@@								:	3. sProperty    : Name of menu property 
'@@								:	4. sPropValue	: Value of property to verify
'@@    Return Value		   	   	: 	True Or False
'@@
'@@    Pre-requisite			:	CPD perspective should be activated., Content Explorer should be opened.						
'@@
'@@    Examples					:	Call  Fn_CPD_CompositeViewMenuOperations("WinMenuSelect", "Exclude Inactive Partitions", "","","")
'@@    Examples					:	Call  Fn_CPD_CompositeViewMenuOperations("CheckItemProperty","Exclude Inactive Partitions,"checked",True,"")
'@@    Examples					:	Call  Fn_CPD_CompositeViewMenuOperations("WinMenuExist","Exclude Inactive Partitions","","","")
'@@
'@@	   History					:	
'@@				Developer Name				Date			Rev. No.	Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Ankit Tewari			23-Dec-2014			1.0			Created
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_CPD_CompositeViewMenuOperations(sAction,sMenuLabel,sProperty,sPropValue,sReserve)
	GBL_FAILED_FUNCTION_NAME="Fn_CPD_CompositeViewMenuOperations"
	'declare variables
	Dim objMenu,strMenu
	Err.clear
	'Set menu object
	Set objMenu = JavaWindow("DefaultWindow").WinMenu("ContextMenu")
	Fn_CPD_CompositeViewMenuOperations = False

	Call Fn_ToolbarButtonClick_Ext(1,"View Menu")
	Wait 2
	
	StrMenu=Replace(sMenuLabel,":",";")

	Select Case sAction
 	'===============Select WinMenu=======================
		Case "WinMenuSelect"
			objMenu.Select(strMenu)
			If Err.Number < 0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Select Menu Item  [" + strMenu + "]")
				Fn_CPD_CompositeViewMenuOperations = False
			Else
				Fn_CPD_CompositeViewMenuOperations = True
			End If
			wait 1
	'===============Check WinMenu Item Property value=======================
		Case "CheckItemProperty"
			Fn_CPD_CompositeViewMenuOperations = objMenu.CheckItemProperty(strMenu,sProperty,sPropValue,5)
	'===============Check Existence of WinMenu Item =======================
		Case "WinMenuExist"
			Fn_CPD_CompositeViewMenuOperations = objMenu.CheckItemProperty(strMenu,"exists",True,5)
	'===============Invalid Case =======================
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_CPD_CompositeViewMenuOperations ] Invalid case [ " & sAction & " ].")
	End Select

	If  Fn_CPD_CompositeViewMenuOperations = False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to execute function [ Fn_CPD_CompositeViewMenuOperations ] with case [ " & sAction & " ].")
	End If
	Set objMenu = Nothing
End Function

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name			:	Fn_CPD_TargetPropertiesOperations
'@@
'@@    Description				:	Function Used for operations on Target Properties window
'@@
'@@    Parameters			    :	1. sAction		: Action to be performed
'@@								:	2. sTabName		: Partitions/Effectivity tab 
'@@								:	3. dicDetails	: Dictionary object
'@@								:	4. sButton 		: OK/Cancel
'@@
'@@    Return Value		   	   	: 	True Or False
'@@
'@@    Pre-requisite			:	CPD perspective should be activated and node should be selected.						
'@@
'@@    Examples					:	Set dicDetails = CreateObject( "Scripting.Dictionary")
'@@    									dicDetails.RemoveAll
'@@    									dicDetails("PopupMenuNodeName") = "000031/A;1-WS123:SubSet123 (0)"
'@@    									dicDetails("SelectTargetPartitions") = "CD000102;1-CD1:Partition Scheme Spatial:PTN000098/001;1-Partition Zone"
'@@    								bReturn = Fn_CPD_TargetPropertiesOperations("Add","Partitions",dicDetails,"OK")
'@@    							:	Set dicDetails = CreateObject( "Scripting.Dictionary")
'@@    									dicDetails.RemoveAll
'@@    									dicDetails("PopupMenuNodeName") = "000031/A;1-WS123:SubSet123 (0)"
'@@    									dicDetails("TargetPartitionsListNode") = "PTN000098/001;1-Partition Zone"
'@@    								bReturn = Fn_CPD_TargetPropertiesOperations("Remove","Partitions",dicDetails,"OK")
'@@    							:	Set dicDetails = CreateObject( "Scripting.Dictionary")
'@@    									dicDetails.RemoveAll
'@@    									dicDetails("PopupMenuNodeName") = "000031/A;1-WS123:SubSet123 (0)"
'@@    									dicDetails("TargetPartitionsListNode") = "PTN000098/001;1-Partition Zone"
'@@    									dicDetails("FromUnit") = 7
'@@    									dicDetails("ToUnit") = 10
'@@    								bReturn = Fn_CPD_TargetPropertiesOperations("SetEffectivity","Effectivity",dicDetails,"OK")
'@@    							:	Set dicDetails = CreateObject( "Scripting.Dictionary")
'@@    									dicDetails.RemoveAll
'@@    									dicDetails("PopupMenuNodeName") = "000031/A;1-WS123:SubSet123 (0)"
'@@    									dicDetails("TargetPartitionsListNode") = "PTN000098/001;1-Partition Zone"
'@@    									dicDetails("FromUnit") = 4~3~~2
'@@    									dicDetails("ToUnit") = 7~~10~12
'@@    								bReturn = Fn_CPD_TargetPropertiesOperations("VerifyEffectivity","Effectivity",dicDetails,"Cancel")
'@@
'@@	   History					:	
'@@			Developer Name		Date		Rev. No.	Changes Done											Reviewer
'@@-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@			Vivek Ahirrao	09-Nov-2015		1.0			Created													[TC1121-2015102600-09_11_2015-VivekA-NewDevelopment]
'@@			Vivek Ahirrao	14-Dec-2015		1.0			Added Case "Effectivity" in this case 					[TC1121-20151116d-18_12_2015-VivekA-NewDevelopment]
'@@														- added Cases "ModifyEffectivity", "SetEffectivity", "VerifyEffectivity"
'@@			Vivek Ahirrao	14-Dec-2015		1.0			Modified Case "Partitions" in this case 				[TC1121-20151116d-18_12_2015-VivekA-NewDevelopment]
'@@														- added Cases "VerifyTargetPartitionsCount", "VerifyTargetPartitionTreeNode", "ExpandTargetPartitionTree"
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_CPD_TargetPropertiesOperations(sAction,sTabName,dicDetails,sButton)
	GBL_FAILED_FUNCTION_NAME="Fn_CPD_TargetPropertiesOperations"
	Dim objTPWindow, bReturn, aPartition, iCount, sPartition
	Dim bFlag, iRowNum, iKeyCnt, aFromUnit, aToUnit, iCount1, sAppFromUnit, sAppToUnit, sItemCount
	Dim arrDate

	Fn_CPD_TargetPropertiesOperations = False
	
	Set objTPWindow = JavaWindow("Collaborative Product").JavaWindow("TargetProperties")
	If Not objTPWindow.Exist(1) Then
		bReturn = Fn_CPD_ContentExplorer("PopupMenuSelect",dicDetails("PopupMenuNodeName"),"","","Target Properties...")
		If bReturn = False Then
			Fn_CPD_TargetPropertiesOperations = False
			Set objTPWindow = Nothing
			Exit Function
		End If
		Call Fn_ReadyStatusSync(3)
	End If
	
	Select Case sTabName
		'Partitions tab operations
		Case "Partitions"
			'Select tab Partitions
			If sTabName<>"" Then
				bReturn = Fn_SISW_UI_JavaTab_Operations("Fn_CPD_TargetPropertiesOperations","Select",objTPWindow,"TargetPartitionsTab",sTabName)
				If bReturn = False Then
					Fn_CPD_TargetPropertiesOperations = False
					Set objTPWindow = Nothing
					Exit Function
				End If
			End If
			
			Select Case sAction
				'Case to Add Partition to Target Partitions List
				Case "Add"
					'Select Target Partitions
					If dicDetails("SelectTargetPartitions")<>"" Then
						aPartition = Split(dicDetails("SelectTargetPartitions"),":")
						For iCount = 0 To UBound(aPartition)
							If iCount = 0 Then
								sPartition = aPartition(iCount)
							Else
								sPartition = sPartition+":"+aPartition(iCount)
							End If
							Call Fn_UI_JavaTree_Expand("Fn_CPD_TargetPropertiesOperations",objTPWindow,"SelectTargetPartitionsTree",sPartition)
						Next
						bReturn = Fn_JavaTree_Select("Fn_CPD_TargetPropertiesOperations",objTPWindow,"SelectTargetPartitionsTree",sPartition)
						If bReturn = False Then
							Fn_CPD_TargetPropertiesOperations = False
							Set objTPWindow = Nothing
							Exit Function
						End If
					End If
					'Click on Add button
					Call Fn_Button_Click("Fn_CPD_TargetPropertiesOperations", objTPWindow, "Add")
				'Case to Remove Partition from Target Partitions List	
				Case "Remove"
					If dicDetails("TargetPartitionsListNode")<>"" Then
						bReturn = Fn_List_Select("Fn_CPD_TargetPropertiesOperations",objTPWindow,"TargetPartitionsList",dicDetails("TargetPartitionsListNode"))
						If bReturn = False Then
							Fn_CPD_TargetPropertiesOperations = False
							Set objTPWindow = Nothing
							Exit Function
						End If
					End If
					'Click on Remove button
					Call Fn_Button_Click("Fn_CPD_TargetPropertiesOperations", objTPWindow, "Remove")
				'Case to verify Partition is present in Target Partitions List or not
				Case "VerifyTargetPartitionsList"
					If dicDetails("TargetPartitionsListNode")<>"" Then
						bReturn = Fn_UI_ListItemExist("Fn_CPD_TargetPropertiesOperations",objTPWindow,"TargetPartitionsList",dicDetails("TargetPartitionsListNode"))
						If bReturn = False Then
							Fn_CPD_TargetPropertiesOperations = False
							Set objTPWindow = Nothing
							Exit Function
						End If
					End If
				'Case to verify count in Target Partitions List
				Case "VerifyTargetPartitionsCount"
					If dicDetails("TargetPartitionsCount")<>"" Then
						sItemCount = objTPWindow.JavaList("TargetPartitionsList").GetROProperty("items count")
						If CInt(dicDetails("TargetPartitionsCount")) <> CInt(sItemCount) Then
							Fn_CPD_TargetPropertiesOperations = False
							Set objTPWindow = Nothing
							Exit Function
						End If
					End If
				'Case to Verify Partition nodes to Target Partitions Tree
				Case "VerifyTargetPartitionTreeNode"
					'Select Target Partitions
					If dicDetails("SelectTargetPartitions")<>"" Then
						aPartition = Split(dicDetails("SelectTargetPartitions"),"~")
						For iCount = 0 To UBound(aPartition)
							bReturn = Fn_UI_JavaTree_NodeExist("Fn_CPD_TargetPropertiesOperations",objTPWindow.JavaTree("SelectTargetPartitionsTree"),aPartition(iCount))
							If bReturn = False Then
								Fn_CPD_TargetPropertiesOperations = False
								Set objTPWindow = Nothing
								Exit Function
							End If
						next
					End If
				'Case to expand Partition in Target Partitions Tree
				Case "ExpandTargetPartitionTree"
					If dicDetails("SelectTargetPartitions")<>"" Then
						aPartition = Split(dicDetails("SelectTargetPartitions"),":")
						For iCount = 0 To UBound(aPartition)
							If iCount = 0 Then
								sPartition = aPartition(iCount)
							Else
								sPartition = sPartition+":"+aPartition(iCount)
							End If
							Call Fn_UI_JavaTree_Expand("Fn_CPD_TargetPropertiesOperations",objTPWindow,"SelectTargetPartitionsTree",sPartition)
							wait 2
							If bReturn = False Then
								Fn_CPD_TargetPropertiesOperations = False
								Set objTPWindow = Nothing
								Exit Function
							End If
						Next
					End If
			End Select
			
			'Click on sButton
			If sButton<>"" Then
				Call Fn_Button_Click("Fn_CPD_TargetPropertiesOperations", objTPWindow, sButton)
			End If
			Fn_CPD_TargetPropertiesOperations = True
			Set objTPWindow = Nothing
			
		'Effectivity tab operations	
		Case "Effectivity"
			Select Case sAction
				Case "Delete a row"
					'Future Use
				Case "ModifyEffectivity"
					If dicDetails("TargetPartitionsListNode")<>"" Then
						'Select Partitions tab
						Call Fn_SISW_UI_JavaTab_Operations("Fn_CPD_TargetPropertiesOperations","Select",objTPWindow,"TargetPartitionsTab","Partitions")
						Wait 1
						bReturn = Fn_List_Select("Fn_CPD_TargetPropertiesOperations", objTPWindow, "TargetPartitionsList",dicDetails("TargetPartitionsListNode"))
						If bReturn = False Then
							Fn_CPD_TargetPropertiesOperations = False
							Set objTPWindow = Nothing
							Exit Function
						End If
					End If
					
					'Select tab Effectivity
					If sTabName<>"" Then
						bReturn = Fn_SISW_UI_JavaTab_Operations("Fn_CPD_TargetPropertiesOperations","Select",objTPWindow,"TargetPartitionsTab",sTabName)
						If bReturn = False Then
							Fn_CPD_TargetPropertiesOperations = False
							Set objTPWindow = Nothing
							Exit Function
						End If
					End If
					
'					iRowNum = objTPWindow.JavaTable("EffectivityTable").GetROProperty("rows")-1
					iRowNum = Fn_UI_Object_GetROProperty("Fn_CPD_TargetPropertiesOperations",objTPWindow.JavaTable("EffectivityTable"), "rows")-1
					
					aFromUnit = Split(dicDetails("FromUnit"),"~")
					aToUnit = Split(dicDetails("ToUnit"),"~")
					sOldFromUnit = aFromUnit(0)
					sOldToUnit = aToUnit(0)
					sNewFromUnit = aFromUnit(1)
					sNewToUnit = aToUnit(1)
					
					For iCount = 0 To iRowNum-1
						'Get application data
						sAppFromUnit = objTPWindow.JavaTable("EffectivityTable").GetCellData(iCount,"From Unit")
						sAppToUnit = objTPWindow.JavaTable("EffectivityTable").GetCellData(iCount,"To Unit")
						If cstr(sAppFromUnit) = sOldFromUnit AND cstr(sAppToUnit) = sOldToUnit Then
							'Set From Unit
							objTPWindow.JavaTable("EffectivityTable").ActivateCell iCount,"From Unit"
							Wait 1
							Call Fn_KeyBoardOperation("SendKeys", "{END}")
							For iKeyCnt = 0 to 10
								Call Fn_KeyBoardOperation("SendKeys", "{BKSP}")
							Next
							JavaWindow("Collaborative Product").JavaWindow("TargetProperties").JavaEdit("UnitText").Set sNewFromUnit
							JavaWindow("Collaborative Product").JavaWindow("TargetProperties").JavaEdit("UnitText").Activate
							Wait 1
							
							'Set To Unit
							objTPWindow.JavaTable("EffectivityTable").ActivateCell iCount,"To Unit"
							Wait 1
							Call Fn_KeyBoardOperation("SendKeys", "{END}")
							For iKeyCnt = 0 to 10
								Call Fn_KeyBoardOperation("SendKeys", "{BKSP}")
							Next
							JavaWindow("Collaborative Product").JavaWindow("TargetProperties").JavaList("ToUnitList").Type sNewToUnit
							Call Fn_KeyBoardOperation("SendKeys", "{ENTER}")
							wait 1
							
							Exit For
						End If
					Next					
					
				Case "SetEffectivity"
					If dicDetails("TargetPartitionsListNode")<>"" Then
						'Select Partitions tab
						Call Fn_SISW_UI_JavaTab_Operations("Fn_CPD_TargetPropertiesOperations","Select",objTPWindow,"TargetPartitionsTab","Partitions")
						Wait 1
						bReturn = Fn_List_Select("Fn_CPD_TargetPropertiesOperations", objTPWindow, "TargetPartitionsList",dicDetails("TargetPartitionsListNode"))
						If bReturn = False Then
							Fn_CPD_TargetPropertiesOperations = False
							Set objTPWindow = Nothing
							Exit Function
						End If
					End If
					
					'Select tab Effectivity
					If sTabName<>"" Then
						bReturn = Fn_SISW_UI_JavaTab_Operations("Fn_CPD_TargetPropertiesOperations","Select",objTPWindow,"TargetPartitionsTab",sTabName)
						If bReturn = False Then
							Fn_CPD_TargetPropertiesOperations = False
							Set objTPWindow = Nothing
							Exit Function
						End If
					End If
					
					iRowNum = objTPWindow.JavaTable("EffectivityTable").GetROProperty("rows")-1
					
					If dicDetails("FromUnit")<>"" Then
						objTPWindow.JavaTable("EffectivityTable").ActivateCell iRowNum,"From Unit"
						Wait 1
						Call Fn_KeyBoardOperation("SendKeys", "{END}")
						For iKeyCnt = 0 to 10
							Call Fn_KeyBoardOperation("SendKeys", "{BKSP}")
						Next
						JavaWindow("Collaborative Product").JavaWindow("TargetProperties").JavaEdit("UnitText").Set dicDetails("FromUnit")
						JavaWindow("Collaborative Product").JavaWindow("TargetProperties").JavaEdit("UnitText").Activate
						Wait 1
					End If
					
					If dicDetails("ToUnit")<>"" Then
						objTPWindow.JavaTable("EffectivityTable").ActivateCell iRowNum,"To Unit"
						Wait 1
						Call Fn_KeyBoardOperation("SendKeys", "{END}")
						For iKeyCnt = 0 to 10
							Call Fn_KeyBoardOperation("SendKeys", "{BKSP}")
						Next
						JavaWindow("Collaborative Product").JavaWindow("TargetProperties").JavaList("ToUnitList").Type dicDetails("ToUnit")
						Call Fn_KeyBoardOperation("SendKeys", "{ENTER}")
						wait 1
					End If
					
					If dicDetails("InDate")<>"" Then
						objTPWindow.JavaTable("EffectivityTable").ActivateCell iRowNum,"In Date"
						Wait 1
						Call Fn_Button_Click("Fn_CPD_TargetPropertiesOperations", objTPWindow, "DateList")
						Wait 1
						If instr( dicDetails("InDate"),"$") > 0 Then
							arrDate = split(trim(dicDetails("InDate")),"$")
							Call  Fn_CPD_DateControl("Set", arrDate(0), arrDate(1))
						Else
							Select Case lcase(trim(dicDetails("InDate")))
								Case ""
									Call  Fn_CPD_DateControl("Clear", "", "")
								Case "today"
									Call  Fn_CPD_DateControl("Today", "", "")
								Case Else
									Call  Fn_CPD_DateControl("Set", dicDetails("InDate"), "")
							End Select
						End If
					End If
					
					If dicDetails("OutDate")<>"" Then
						objTPWindow.JavaTable("EffectivityTable").ActivateCell iRowNum,"Out Date"
						Wait 1
						Call Fn_List_Select("Fn_CPD_TargetPropertiesOperations",objTPWindow,"ToUnitList","Select Date...")
						Wait 1
						If instr( dicDetails("OutDate"),"$") > 0 Then
							arrDate = split(trim(dicDetails("OutDate")),"$")
							Call  Fn_CPD_DateControl("Set", arrDate(0), arrDate(1))
						Else
							Select Case lcase(trim(dicDetails("OutDate")))
								Case ""
									Call  Fn_CPD_DateControl("Clear", "", "")
								Case "today"
									Call  Fn_CPD_DateControl("Today", "", "")
								Case Else
									Call  Fn_CPD_DateControl("Set", dicDetails("OutDate"), "")
							End Select
						End If
					End If
					
				Case "VerifyEffectivity"
					If dicDetails("TargetPartitionsListNode")<>"" Then
						'Select Partitions tab
						Call Fn_SISW_UI_JavaTab_Operations("Fn_CPD_TargetPropertiesOperations","Select",objTPWindow,"TargetPartitionsTab","Partitions")
						Wait 1
						bReturn = Fn_List_Select("Fn_CPD_TargetPropertiesOperations", objTPWindow, "TargetPartitionsList",dicDetails("TargetPartitionsListNode"))
						If bReturn = False Then
							Fn_CPD_TargetPropertiesOperations = False
							Set objTPWindow = Nothing
							Exit Function
						End If
						Wait 1
					End If
					
					'Select tab Effectivity
					If sTabName<>"" Then
						bReturn = Fn_SISW_UI_JavaTab_Operations("Fn_CPD_TargetPropertiesOperations","Select",objTPWindow,"TargetPartitionsTab",sTabName)
						If bReturn = False Then
							Fn_CPD_TargetPropertiesOperations = False
							Set objTPWindow = Nothing
							Exit Function
						End If
					End If
					
'					iRowNum = objTPWindow.JavaTable("EffectivityTable").GetROProperty("rows")-1
					iRowNum = Fn_UI_Object_GetROProperty("Fn_CPD_TargetPropertiesOperations",objTPWindow.JavaTable("EffectivityTable"), "rows")-1
					If dicDetails("FromUnit")<>"" OR dicDetails("ToUnit")<>"" Then
						aFromUnit = Split(dicDetails("FromUnit"),"~")
						aToUnit = Split(dicDetails("ToUnit"),"~")
						For iCount1 = 0 To UBound(aFromUnit)
							bFlag = False
							For iCount = 0 To iRowNum-1
								'Get application data
								sAppFromUnit = objTPWindow.JavaTable("EffectivityTable").GetCellData(iCount,"From Unit")
								sAppToUnit = objTPWindow.JavaTable("EffectivityTable").GetCellData(iCount,"To Unit")
								If cstr(sAppFromUnit) = aFromUnit(iCount1) AND cstr(sAppToUnit) = aToUnit(iCount1) Then
									bFlag = True
									Exit For
								End If
							Next
							If bFlag = False Then
								Fn_CPD_TargetPropertiesOperations = False
								Set objTPWindow = Nothing
								Exit Function
							End If
						Next
					End If
			End Select
			
			'Click on sButton
			If sButton<>"" Then
				Call Fn_Button_Click("Fn_CPD_TargetPropertiesOperations", objTPWindow, sButton)
			End If
			Fn_CPD_TargetPropertiesOperations = True
			Set objTPWindow = Nothing
			
		Case Else
			Fn_CPD_TargetPropertiesOperations = False
			Set objTPWindow = Nothing
			Exit Function
	End Select
End Function

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name			:	Fn_CPD_UpdateInstantiationOfModelContentOperations
'@@
'@@    Description				:	Function Used for operations on Target Properties window
'@@
'@@    Parameters			    :	1. sAction		: Action to be performed
'@@								:	2. dicDetails	: Dictionary object
'@@								:	4. sButton 		: OK/Cancel
'@@
'@@    Return Value		   	   	: 	True Or False
'@@
'@@    Pre-requisite			:	CPD perspective should be activated and node should be selected.						
'@@
'@@    Examples					:	Set dicDetails = CreateObject( "Scripting.Dictionary")
'@@    									dicDetails.RemoveAll
'@@    									dicDetails("PopupMenuNodeName") = "CD000007;1-TargetCD:DE000002/001;1-Test"
'@@    									dicDetails("SelectSubsetDefinition") = "Replay Existing Recipe"/"Use Subset Definition"
'@@    								bReturn = Fn_CPD_UpdateInstantiationOfModelContentOperations("Set",dicDetails,"OK")
'@@   							:	Set dicDetails = CreateObject( "Scripting.Dictionary")
'@@    									dicDetails.RemoveAll
'@@    									dicDetails("PopupMenuNodeName") = "CD000007;1-TargetCD:DE000002/001;1-Test"
'@@    									dicDetails("Replay Existing Recipe") = "ON"
'@@    									dicDetails("Use Subset Definition") = "OFF"
'@@    									dicDetails("Revision Rule") = "Working(Current User); Any Status"
'@@    								bReturn = Fn_CPD_UpdateInstantiationOfModelContentOperations("Verify",dicDetails,"OK")
'@@
'@@	   History					:	
'@@			Developer Name		Date		Rev. No.	Changes Done											Reviewer
'@@-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@			Vivek Ahirrao	16-Nov-2015		1.0			Created													[TC1121-2015102600-16_11_2015-VivekA-NewDevelopment]
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_CPD_UpdateInstantiationOfModelContentOperations(sAction,dicDetails,sButton)
	GBL_FAILED_FUNCTION_NAME="Fn_CPD_UpdateInstantiationOfModelContentOperations"
	Dim objUIMCWindow, bReturn, sValue, sText
	Fn_CPD_UpdateInstantiationOfModelContentOperations = False
	
	Set objUIMCWindow = JavaWindow("Collaborative Product").JavaWindow("UpdateInstantiationOfModelContent")
	If Not objUIMCWindow.Exist(1) Then
		bReturn = Fn_CPD_ContentExplorer("PopupMenuSelect",dicDetails("PopupMenuNodeName"),"","","Update Model Instantiation")
		If bReturn = False Then
			Fn_CPD_UpdateInstantiationOfModelContentOperations = False
			Set objUIMCWindow = Nothing
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Window ["+objUIMCWindow.toString+"] does not Exist.")
			Exit Function
		End If
		Call Fn_ReadyStatusSync(3)
	End If
	
	Select Case sAction
		Case "Set"
			If dicDetails("SelectSubsetDefinition")<>"Replay Existing Recipe" Then
				objUIMCWindow.JavaRadioButton("SubsetDefinitionRadioBtn").SetTOProperty "attached text",dicDetails("SelectSubsetDefinition")
				Wait 1
				objUIMCWindow.JavaRadioButton("SubsetDefinitionRadioBtn").Set "ON"
				Wait 1
			ElseIf dicDetails("SelectSubsetDefinition")<>"Use Subset Definition" Then
'				objUIMCWindow.JavaRadioButton("SubsetDefinitionRadioBtn").SetTOProperty "attached text",dicDetails("SelectSubsetDefinition")
'				Wait 1
'				objUIMCWindow.JavaRadioButton("SubsetDefinitionRadioBtn").Set "ON"
'				Wait 1
				'Open by Name dialog operations
				'Future Use
			Else
				Fn_CPD_UpdateInstantiationOfModelContentOperations = False
				Set objUIMCWindow = Nothing
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Radio Button does not Exist.")
				Exit Function
			End If
		
		Case "Verify"
			'Verify "Replay Existing Recipe" Radio Button
			If dicDetails("Replay Existing Recipe")<>"" Then
				sValue = ""
				objUIMCWindow.JavaRadioButton("SubsetDefinitionRadioBtn").SetTOProperty "attached text","Replay Existing Recipe"
				Wait 1
				sValue = objUIMCWindow.JavaRadioButton("SubsetDefinitionRadioBtn").GetROProperty("value")
				If sValue = "1" Then
					sValue = "ON"
				ElseIf sValue = "0" Then
					sValue = "OFF"
				End If
				If dicDetails("Replay Existing Recipe")<>sValue Then
					Fn_CPD_UpdateInstantiationOfModelContentOperations = False
					Set objUIMCWindow = Nothing
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Radio Button is not set ["+dicDetails("Replay Existing Recipe")+"].")
					Exit Function
				End If
			End If
			'Verify "Use Subset Definition" Radio Button
			If dicDetails("Use Subset Definition")<>"" Then
				sValue = ""
				objUIMCWindow.JavaRadioButton("SubsetDefinitionRadioBtn").SetTOProperty "attached text","Use Subset Definition"
				Wait 1
				sValue = objUIMCWindow.JavaRadioButton("SubsetDefinitionRadioBtn").GetROProperty("value")
				If sValue = "1" Then
					sValue = "ON"
				ElseIf sValue = "0" Then
					sValue = "OFF"
				End If
				If dicDetails("Use Subset Definition")<>sValue Then
					Fn_CPD_UpdateInstantiationOfModelContentOperations = False
					Set objUIMCWindow = Nothing
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Radio Button is not set ["+dicDetails("Use Subset Definition")+"].")
					Exit Function
				End If
			End If
			'Verify "Revision Rule"
			If dicDetails("Revision Rule")<>"" Then
				sText = ""
				objUIMCWindow.JavaStaticText("SrcConfigPropName").SetTOProperty "label", "Revision Rule:"
				Wait 1
				If objUIMCWindow.JavaStaticText("SrcConfigPropValue").Exist(1) Then
					sText = objUIMCWindow.JavaStaticText("SrcConfigPropValue").GetROProperty("label")
				Else
					Fn_CPD_UpdateInstantiationOfModelContentOperations = False
					Set objUIMCWindow = Nothing
					Exit Function
				End If
				If Trim(dicDetails("Revision Rule")) <> Trim(sText) Then
					Fn_CPD_UpdateInstantiationOfModelContentOperations = False
					Set objUIMCWindow = Nothing
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Revision Rule value is not correct.")
					Exit Function
				End If
			End If
			'Verify "Effectivity Formula"
			If dicDetails("Effectivity Formula")<>"" Then
				sText = ""
				objUIMCWindow.JavaStaticText("SrcConfigPropName").SetTOProperty "label", "Effectivity Formula:"
				Wait 1
				If objUIMCWindow.JavaStaticText("SrcConfigPropValue").Exist(1) Then
					sText = objUIMCWindow.JavaStaticText("SrcConfigPropValue").GetROProperty("label")
				Else
					Fn_CPD_UpdateInstantiationOfModelContentOperations = False
					Set objUIMCWindow = Nothing
					Exit Function
				End If
				If Trim(dicDetails("Effectivity Formula")) <> Trim(sText) Then
					Fn_CPD_UpdateInstantiationOfModelContentOperations = False
					Set objUIMCWindow = Nothing
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Effectivity Formula value is not correct.")
					Exit Function
				End If
			End If
			'Verify "Saved Variant Rule"
			If dicDetails("Saved Variant Rule")<>"" Then
				sText = ""			
				objUIMCWindow.JavaStaticText("SrcConfigPropName").SetTOProperty "label", "Saved Variant Rule:"
				Wait 1
				If objUIMCWindow.JavaStaticText("SrcConfigPropValue").Exist(1) Then
					sText = objUIMCWindow.JavaStaticText("SrcConfigPropValue").GetROProperty("label")
				Else
					Fn_CPD_UpdateInstantiationOfModelContentOperations = False
					Set objUIMCWindow = Nothing
					Exit Function
				End If
				If Trim(dicDetails("Saved Variant Rule")) <> Trim(sText) Then
					Fn_CPD_UpdateInstantiationOfModelContentOperations = False
					Set objUIMCWindow = Nothing
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Saved Variant Rule value is not correct.")
					Exit Function
				End If
			End If
		Case Else
			Fn_CPD_UpdateInstantiationOfModelContentOperations = False
			Set objUIMCWindow = Nothing
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Case is not valid.")
			Exit Function
	End Select
	'Click on sButton
	If sButton<>"" Then
		Call Fn_Button_Click("Fn_CPD_UpdateInstantiationOfModelContentOperations", objUIMCWindow, sButton)
	End If
	Fn_CPD_UpdateInstantiationOfModelContentOperations = True
	Set objUIMCWindow = Nothing
End Function

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name			:	Fn_CPD_ModelContentCloneAndInstantiationOperations
'@@
'@@    Description				:	Function Used for operations on "Model Content Clone And Instantiation" dialog
'@@
'@@    Parameters			    :	1. sAction		: Action to be performed
'@@								:	2. dicDetails	: Dictionary object
'@@								:	4. sButton 		: OK/Cancel
'@@
'@@    Return Value		   	   	: 	True Or False
'@@
'@@    Pre-requisite			:	"Model Content Clone And Instantiation" dialog should be opened.						
'@@
'@@    Examples					:	Set dicDetails = CreateObject("Scripting.Dictionary")
'@@    								With dicDetails 
'@@    									''''Page number 1
'@@    									.Add "RadioButton1", "CloneofDesignElements:ON"
'@@    									.Add "RadioButton2", "InstantiationofDesignElements:OFF~Next"
'@@    									''''Page number 2
'@@    									.Add "StaticText1", "Name"
'@@    									.Add "StaticText2", "Description"
'@@    									.Add "EditBox1", "Name:Test123:Set"
'@@    									.Add "EditBox2", "Description::Get~Next"
'@@    								End with
'@@    								bReturn = Fn_CPD_ModelContentCloneAndInstantiationOperations("Verify",dicDetails,"Finish")
'@@	   Use						:	First perform operation on one page and send Next or Cancel Or Finish button
'@@									Radio Button : Use "Set" or "Get" as in example to set or verify Radio Button
'@@									Edit Box 	 : Use "Set" or "Get" as in example to set or verify Edit Box Value
'@@									Check Box 	 : Use "Set" or "Get" as in example to set or verify Check Box Value
'@@									Static Text	 : As in example to verify Static Text Value
'@@	   History					:	
'@@			Developer Name		Date	  		Rev. No.	Changes Done										Reviewer
'@@-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@			Vivek Ahirrao		02-Dec-2015		1.0		  	Created												[TC1121-2015110900-02_12_2015-VivekA-NewDevelopment]
'@@         shweta rathod		10-Dec-2015					Added case "IsNameMandatory" 						[TC1121-20151116a-10_12_2015-VivekA-NewDevelopment]
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_CPD_ModelContentCloneAndInstantiationOperations(sAction,dicDetails,sButton)
	GBL_FAILED_FUNCTION_NAME="Fn_CPD_ModelContentCloneAndInstantiationOperations"
	Dim bFlag, iCounter, bReturn
	Dim aParameter, sProperty, aProperty, sPropertyName, sPropertyValue, sPropertyAction, sObjectCase, sOldPropertyName
	Dim objDialog, objPageButton
	Dim dicCount, dicItems, dicKeys
	
	Fn_CPD_ModelContentCloneAndInstantiationOperations = False
	Wait 2
	bFlag = False
   	Set objDialog =Fn_SISW_CPD_GetObject("Modelcontentcloneandinstantiation")
   	'Check the existance of dialog
	If objDialog.Exist(4) Then
		bFlag = True
	Else
		Set objDialog = Fn_SISW_CPD_GetObject("Modelcloneandrealization")
		If objDialog.Exist(2) Then
			bFlag = True
		End If		
	End If

	'Check the Existence of the ModelContentCloneAndInstantiationOperations window, if not Exist the Function will be Terminated
	If bFlag = False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : [ Fn_CPD_ModelContentCloneAndInstantiationOperations ] Dialog does not Exist.")
		Set objDialog = Nothing
		Exit Function
	End If
	
	Select Case sAction
		Case "Set"
			'Future Use
			
		Case "Verify"
			If IsObject(dicDetails) Then
				dicCount = dicDetails.Count
				dicItems = dicDetails.Items
				dicKeys = dicDetails.Keys
				For iCounter = 0 to dicCount - 1
					If dicItems(iCounter) <> "" Then
						If Instr(dicItems(iCounter),"~")>0 Then
							aParameter = Split(dicItems(iCounter),"~")
							sProperty = aParameter(0)
							objPageButton = aParameter(1)
						Else
							sProperty = dicItems(iCounter)
						End If
						
						If Instr(sProperty,":")>0 Then
							aProperty = Split(sProperty,":")
							sPropertyName = aProperty(0)
							sPropertyValue = aProperty(1)
							If UBound(aProperty)>1 Then
								sPropertyAction = aProperty(2)
							End If
						Else
							sPropertyName = sProperty
						End If
						If Instr(dicKeys(iCounter),"RadioButton")>0 Then
							sObjectCase = "RadioButton"
						ElseIf Instr(dicKeys(iCounter),"EditBox")>0 Then
							sObjectCase = "EditBox"
						ElseIf Instr(dicKeys(iCounter),"CheckBox")>0 Then
							sObjectCase = "CheckBox"
						ElseIf Instr(dicKeys(iCounter),"StaticText")>0 Then
							sObjectCase = "StaticText"
						End If
						Select Case sObjectCase
							Case "RadioButton"
								'Select the RadioButton
								If sPropertyName<>"" Then
									Select Case sPropertyName
									
										Case "CloneofDesignElements"
											sPropertyName = "Clone of Design Components"						
										Case "InstantiationofDesignElements"
											sPropertyName = "Instantiation of Design Components"
									End Select
									' ----------Modified by Chaitali R.----------------									
									sOldPropertyName = sPropertyName
									sPropertyName = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("CPD_DisplayName"),sPropertyName)
									If sPropertyName = False Then
										sPropertyName = sOldPropertyName
									End If
									
									objDialog.JavaRadioButton("ActionToPerform").SetTOProperty "attached text", sPropertyName
									Wait 1
									If NOT objDialog.JavaRadioButton("ActionToPerform").Exist(2) Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : [ Fn_CPD_ModelContentCloneAndInstantiationOperations ] Radio Button [ "+sPropertyName+" ] does not Exist.")
										Set objDialog = Nothing
										Fn_CPD_ModelContentCloneAndInstantiationOperations = False
										Exit Function
									End If
									'Set the Radiobutton ON or OFF
									If sPropertyValue<>"" AND sPropertyAction<>"" Then
										Select Case sPropertyAction
											Case "Set"
												If sPropertyValue = "ON" Then
													bReturn = Fn_UI_JavaRadioButton_SetON("Fn_CPD_ModelContentCloneAndInstantiationOperations",objDialog,"ActionToPerform")
												ElseIf sPropertyValue = "OFF" Then
													bReturn = Fn_UI_JavaRadioButtont_setOff("Fn_CPD_ModelContentCloneAndInstantiationOperations",objDialog,"ActionToPerform")
												End If
												Wait 1
											Case "Get"
												bFlag = Fn_UI_Object_GetROProperty("Fn_CPD_ModelContentCloneAndInstantiationOperations",objDialog.JavaRadioButton("ActionToPerform"),"value")
												If bFlag = sPropertyValue Then
													bReturn = True
												Else
													bReturn = False
												End If
												Wait 1
										End Select
										If bReturn = False Then
											 Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_CPD_ModelContentCloneAndInstantiationOperations ] Failed to "+sPropertyAction+" The [ "+sPropertyName+" ] Radio Button.")
											 Set objDialog = Nothing
											 Exit Function
										End If									
									End If
								End If						
								Wait 1
								
							Case "EditBox"
								'Set or Verify the Edit Box value
								If sPropertyName<>"" Then
									Select Case sPropertyName
										Case "IsNameMandatory"
											'Enter the Name of the ModelReuse Design Element
											bReturn = Fn_UI_Object_GetROProperty("",objDialog.JavaButton("Next"),"enabled")
											If cbool(bReturn) <> true Then
												If objDialog.JavaEdit("Name").Exist(3) Then
													objDialog.JavaEdit("Name").Set sPropertyValue
													Wait 1
													bReturn = Fn_UI_Object_GetROProperty("",objDialog.JavaButton("Next"),"enabled")
													If cbool(bReturn) <> true Then
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_CPD_ModelContentCloneAndInstantiationOperations ] [ NEXT ] button is not enabled. Hence failed to verify edit box [ Name ] is Mandatory.")
														Set objDialog = Nothing
														Fn_CPD_ModelContentCloneAndInstantiationOperations = False
														Exit Function
													End If
												End If
											End If
										Case "Name"
											sPropertyName = "Name"
											If NOT objDialog.JavaEdit("Name").Exist(2) Then
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : [ Fn_CPD_ModelContentCloneAndInstantiationOperations ] Edit Box [ "+sPropertyName+" ] does not Exist.")
												Set objDialog = Nothing
												Fn_CPD_ModelContentCloneAndInstantiationOperations = False
												Exit Function
											End If
										Case "Description"
											sPropertyName = "Description"
											If NOT objDialog.JavaEdit("Description").Exist(2) Then
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : [ Fn_CPD_ModelContentCloneAndInstantiationOperations ] Edit Box [ "+sPropertyName+" ] does not Exist.")
												Set objDialog = Nothing
												Fn_CPD_ModelContentCloneAndInstantiationOperations = False
												Exit Function
											End If
									End Select
									
									If sPropertyValue<>"" AND sPropertyAction<>"" Then
										Select Case sPropertyAction
											Case "Set"
												bReturn = Fn_Edit_Box("Fn_CPD_ModelContentCloneAndInstantiationOperations",objDialog,sPropertyName,sPropertyValue)
												Wait 1
											Case "Get"
												bFlag = Fn_Edit_Box_GetValue("Fn_CPD_ModelContentCloneAndInstantiationOperations",objDialog,sPropertyName)
												If Trim(bFlag) = Trim(sPropertyValue) Then
													bReturn = True
												Else
													bReturn = False
												End If
										End Select
										If bReturn = False Then
											 Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_CPD_ModelContentCloneAndInstantiationOperations ] Failed to "+sPropertyAction+" The [ "+sPropertyName+" ] Edit Box.")
											 Set objDialog = Nothing
											 Exit Function
										End If										
									End If
								End If
								
							Case "CheckBox"
								'Set or Verify the Check Box value
								If sPropertyName<>"" Then
									Select Case sPropertyName
										Case "Check-outoncreate"
											sPropertyName = "Check-out on create"
										Case "Applyeffecitivitybasedontargetrevisionrule"
											sPropertyName = "Apply effecitivity based on target revision rule"
										Case "ApplyeffecitivitybasedonsourceDesignElements"
											sPropertyName = "Apply effecitivity based on source Design Elements"
										Case "ApplyvariantconditionsbasedonsourceDesignElements"
											sPropertyName = "Apply variant conditions based on source Design Elements"
									End Select
									
									' ----------Modified by Chaitali R.----------------	
									sOldPropertyName = sPropertyName
									sPropertyName = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("CPD_DisplayName"),sPropertyName)
									If sPropertyName = False Then
										sPropertyName = sOldPropertyName
									End If
									
									objDialog.JavaCheckBox("SelectPartitionSchemes").SetTOProperty "attached text", sPropertyName
									Wait 1
									If NOT objDialog.JavaCheckBox("SelectPartitionSchemes").Exist(3) Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : [ Fn_CPD_ModelContentCloneAndInstantiationOperations ] Check Box [ "+sPropertyName+" ] does not Exist.")
										Set objDialog = Nothing
										Fn_CPD_ModelContentCloneAndInstantiationOperations = False
										Exit Function
									End If 
	
									'Set the Check Box ON or OFF
									If sPropertyValue<>"" AND sPropertyAction<>"" Then
										Select Case sPropertyAction
											Case "Set"
												bReturn = Fn_CheckBox_Set("Fn_CPD_ModelContentCloneAndInstantiationOperations",objDialog,"SelectPartitionSchemes",sPropertyValue)
												Wait 1
											Case "Get"
												bFlag = Fn_UI_Object_GetROProperty("Fn_CPD_ModelContentCloneAndInstantiationOperations",objDialog.JavaCheckBox("SelectPartitionSchemes"),"value")
												If bFlag = sPropertyValue Then
													bReturn = True
												Else
													bReturn = False
												End If
										End Select
										If bReturn = False Then
											 Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_CPD_ModelContentCloneAndInstantiationOperations ] Failed to "+sPropertyAction+" The [ "+sPropertyName+" ] Check Box.")
											 Set objDialog = Nothing
											 Fn_CPD_ModelContentCloneAndInstantiationOperations = False
											 Exit Function
										End If									
									End If								
								End If
								Wait 1
							Case "StaticText"
								'Set or Verify the Check Box value
								If sPropertyName<>"" Then
									Select Case sPropertyName
										Case "Name"
											sPropertyName = "Name"
										Case "Description"
											sPropertyName = "Description"
										Case "Actiontoperform"
											sPropertyName = "Action to perform"
										Case "ModelReuseElementInformation"
											sPropertyName = "Model Reuse Element Information"
										Case "TargetModelConfigurationOptions"
											sPropertyName = "Target Model Configuration Options"
										Case "SourceModelConfigurationOptions"
											sPropertyName = "Source Model Configuration Options"
									End Select
									
									' ----------Modified by Chaitali R.----------------	
									sOldPropertyName = sPropertyName
									sPropertyName = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("CPD_DisplayName"),sPropertyName)
									If sPropertyName = False Then
										sPropertyName = sOldPropertyName
									End If
									
									objDialog.JavaStaticText("SourceModelContent").SetTOProperty "label", sPropertyName
									Wait 1
									If NOT objDialog.JavaStaticText("SourceModelContent").Exist(3) Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : [ Fn_CPD_ModelContentCloneAndInstantiationOperations ] Static Text [ "+sPropertyName+" ] does not Exist.")
										Set objDialog = Nothing
										Fn_CPD_ModelContentCloneAndInstantiationOperations = False
										Exit Function
									End If 
								End If
						End Select
					End If
					'Click Next button as provided
					If objPageButton<>"" Then
						bReturn = Fn_Button_Click("Fn_CPD_ModelContentCloneAndInstantiationOperations",objDialog,objPageButton)
						If bReturn = False Then
							 Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_CPD_ModelContentCloneAndInstantiationOperations ] Failed to Click [ "+objPageButton+" ] button ")
							 Set objDialog = Nothing
							 Fn_CPD_ModelContentCloneAndInstantiationOperations = False
							 Exit Function
						End If
						Wait 2
						objPageButton = ""
					End If
				Next
			End If
			'Click Cancel or Finish button as provided
			If sButton<>"" Then
				bReturn = Fn_Button_Click("Fn_CPD_ModelContentCloneAndInstantiationOperations",objDialog,sButton)
				If bReturn = False Then
					 Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_CPD_ModelContentCloneAndInstantiationOperations ] Failed to Click [ "+sButton+" ] button ")
					 Set objDialog = Nothing
					 Fn_CPD_ModelContentCloneAndInstantiationOperations = False
					 Exit Function
				End If
				Wait 5
			End If
		Case Else
			'Future Use
	End Select
	Set objDialog = Nothing
	Fn_CPD_ModelContentCloneAndInstantiationOperations = True
End Function

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name	:	Fn_CPD_SourceTargetTables_Operations
'@@
'@@    Description		:	Function Used to perform operations on "Maintenance Actions" window
'@@
'@@    Parameters		:	1. sAction		: Action to be performed
'@@						:	2. sTabName		: Tab name to be selected [Full Results or Equivalent Lines]
'@@						:	3. sTableName 	: Table name [Source ( PTN000001/001;1-static_ptn )]  or [Target ( PTN000002/001;1-Dynamin_ptn )]
'@@						:	4. sColNames	: Column names
'@@						:	5. sColValues	: Column values
'@@						:	6. dicDetails	: Dictionary object
'@@						:	7. sPopupMenu	: Popup Menu to select
'@@
'@@    Return Value		: 	True Or False Or Row number in integer format
'@@
'@@    Examples			:	bReturn = Fn_CPD_SourceTargetTables_Operations("GetRowNumber","Full Results","Source ( PTN000001/001;1-static_ptn )","Object~Category~Type","DE000003/001;1-DE3~Promissory~Design Element","","")
'@@    					:	bReturn = Fn_CPD_SourceTargetTables_Operations("VerifyRow","Full Results","Target ( PTN000002/001;1-Dynamin_ptn )","Object~Category~Type","DE000006/001;1-DE4~Promissory~Design Element","","")
'@@    					:	bReturn = Fn_CPD_SourceTargetTables_Operations("SelectRow","Full Results","Source ( PTN000001/001;1-static_ptn )","Object~Category~Type","DE000003/001;1-DE3~Promissory~Design Element","","")
'@@    					:	bReturn = Fn_CPD_SourceTargetTables_Operations("DoubleClickRow","Full Results","Source ( src_sdfn_01 )","Object~Category","DE000100/001;1-flipfone_antenna_03427~Subordinate","","")
'@@    							
'@@	   History			:	
'@@	  	Developer Name		Date	   Rev. No.		Changes Done										Reviewer
'@@-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@	  	Vivek Ahirrao	 22-Aug-2016	1.0			Created - Added for 4GD new TC's development		[TC1123-20160729-22_08_2016-VivekA-NewDevelopment]
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_CPD_SourceTargetTables_Operations(sAction,sTabName,sTableName,sColNames,sColValues,dicDetails,sPopupMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_CPD_SourceTargetTables_Operations"
	Dim objWindow, objTable
	Dim iRowNumber, iRowCount, iCount, sColIntName, iCount1, bFlag, sAppText
	Dim aColName, aColValue
	
	Fn_CPD_SourceTargetTables_Operations = False
	On Error Resume Next
	
	Set objWindow = JavaWindow("Collaborative Product")
	Set objTable = objWindow.JavaTable("SourceTargetTable")
	
	'Maximise the 4GD Compare tab-----
	If Fn_CPD_CompnentTabOperations("IsMaximized","4GD Compare","")=False Then
		bFlag = Fn_CPD_CompnentTabOperations("DoubleClick","4GD Compare","")
		If bFlag = False Then
			Set objWindow = Nothing
			Set objTable = Nothing
			Exit Function
		End If
	End If
	'---------------------------------
	'Give Table name to work on Source or Target table
	If sTableName<>"" Then
		'Set Tab if not empty
		If sTabName<>"" Then
			objWindow.JavaStaticText("SourceTargetTable").SetTOProperty "label",sTableName
'			objWindow.JavaTab("SourceTagetTab").SetTOProperty "attached text",sTableName
			objWindow.JavaTab("SourceTagetTab").Select sTabName
			If Err.Number<>0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [Fn_CPD_SourceTargetTables_Operations] Failed to Select ["+sTabName+"] Tab.")
				Set objWindow = Nothing
				Set objTable = Nothing
				Exit Function
			End If
		End If
		
		If Instr(sTableName,"Source")>0 Then
			objTable.SetTOProperty "logical_location","X_BIG__Y_UNK"
		ElseIf Instr(sTableName,"Target")>0 Then
			objTable.SetTOProperty "logical_location","X_BIG__Y_BIG"
		ElseIf Instr(sTableName,"Partial")>0 Then     'added by piyush - added for Partial table
			objTable.SetTOProperty "logical_location","X_BIG__Y_BIG"
		End If
		
		objTable.SetTOProperty "attached text",sTableName
		If objTable.Exist(1)=False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [Fn_CPD_SourceTargetTables_Operations] Table ["+sTableName+"] does noe Exist!!!")
			Set objWindow = Nothing
			Set objTable = Nothing
			Exit Function
		End If
	End If
	Set objWindow = Nothing
	
	Select Case sAction
		'Case to Select or DoubleClick particular Row in Source or Target Table
		'Case to Select or DoubleClick particular Row in Partial Matach Details Table
		Case "SelectRow", "DoubleClickRow","SelectRowPartial", "DoubleClickRowPartial"
				If sColNames<>"" AND sColValues<>"" Then
						'Get row number on which PopupMenu Select operation wants to perform
						If sAction = "SelectRow" OR sAction = "DoubleClickRow" Then
							iRowNumber = Fn_CPD_SourceTargetTables_Operations("GetRowNumber","",sTableName,sColNames,sColValues,"","")
						ElseIf sAction = "SelectRowPartial" OR sAction = "DoubleClickRowPartial" Then
							iRowNumber = Fn_CPD_SourceTargetTables_Operations("GetRowNumberPartial","",sTableName,sColNames,sColValues,"","")
						End If
						If iRowNumber<>-1 Then
							If sAction = "SelectRow" Then
								JavaWindow("Collaborative Product").JavaTable("SourceTargetTable").SelectRow iRowNumber
							ElseIf sAction = "DoubleClickRow" Then
							    If Fn_CPD_CompnentTabOperations("IsMaximized","4GD Compare","")=False  Then
									bFlag = Fn_CPD_CompnentTabOperations("DoubleClick","4GD Compare","")
									If bFlag = False Then
										Set objWindow = Nothing
										Set objTable = Nothing
										Exit Function
									End If
								End If  
								JavaWindow("Collaborative Product").JavaTable("SourceTargetTable").DoubleClickCell iRowNumber,"Object"
							End If
							If Err.Number<>0 Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [Fn_CPD_SourceTargetTables_Operations] Failed to ["+sAction+"] in ["+sTableName+"] Table.")
								Set objTable = Nothing
								Exit Function
							End If
						Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [Fn_CPD_SourceTargetTables_Operations] Failed to Find row in ["+sTableName+"] Table.")
							Set objTable = Nothing
							Exit Function
						End If
						Set objTable = Nothing
						Fn_CPD_SourceTargetTables_Operations = True
				End If
		'Case to verify Existence or to get a Row number in Source or Target Table
		'Case to verify Existence or to get a Row number in Partial Matach Details Table
		Case "VerifyRow","GetRowNumber","VerifyRowPartial","GetRowNumberPartial"
				If sColNames<>"" AND sColValues<>"" Then
						iRowNumber = -1
						If sAction = "GetRowNumber" OR sAction = "GetRowNumberPartial" Then
							Fn_CPD_SourceTargetTables_Operations = iRowNumber
						End If
						iRowCount = objTable.GetROProperty("rows")
						aColName = Split(sColNames,"~")
						aColValue = Split(sColValues,"~")
						For iCount = 0 To CInt(iRowCount)-1
							sColIntName = ""
							For iCount1 = 0 To UBound(aColName)
								bFlag = False
								If sAction = "VerifyRow" OR sAction = "GetRowNumber" Then
									'Use [ objTable.Object.getItem(0).getData.getSelectedComponents.get(0).getProperties.tostring ]
									'to get all Internal names for column for this table
									Select Case aColName(iCount1)
										Case "Object"
											sColIntName = "object_string"
										Case "Category"
											sColIntName = "cpd0category"
										Case "Type"
											sColIntName = "object_type"
										Case "NX Entity Handle"
											sColIntName = "cpd0UG_ENTITY_HANDLE"
										Case "Revision ID"
											sColIntName = "fnd0RevisionId"
										Case Else
											sColIntName = ""
									End Select
								End If
								
								If sAction = "VerifyRow" OR sAction = "GetRowNumber" Then
									sAppText = objTable.Object.getItem(iCount).getData.getSelectedComponents.get(0).getProperty(sColIntName)
								ElseIf sAction = "VerifyRowPartial" OR sAction = "GetRowNumberPartial" Then
									sAppText = Fn_SISW_UI_JavaTable_Operations("Fn_CPD_SourceTargetTables_Operations","GetCellData",objTable,"","","",iCount,aColName(iCount1),"","","")
								Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [Fn_CPD_SourceTargetTables_Operations] Column ["+aColName(iCount1)+"] not Found.")
									Set objTable = Nothing
									Exit Function
								End If
								If sAppText<>aColValue(iCount1) Then
									bFlag = False
									Exit For
								End If
								bFlag = True
							Next
							'If Found Row with column values
							If bFlag = True Then
								iRowNumber = iCount
								Exit For
							End If
						Next
						If bFlag = False Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [Fn_CPD_SourceTargetTables_Operations] Row not Found.")
							If Fn_CPD_CompnentTabOperations("IsMaximized","4GD Compare","") Then
								call Fn_CPD_CompnentTabOperations("DoubleClick","4GD Compare","")
							End If
							Set objTable = Nothing
							Exit Function
						End If
						
						If sAction = "GetRowNumber" OR sAction = "GetRowNumberPartial" Then
							Fn_CPD_SourceTargetTables_Operations = iRowNumber
						ElseIf sAction = "VerifyRow" OR sAction = "VerifyRowPartial" Then
							Fn_CPD_SourceTargetTables_Operations = True
						End If
				End If
		Case Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [Fn_CPD_SourceTargetTables_Operations] Enter valid case.")
	End Select
	Set objTable = Nothing
	'Maximise the 4GD Compare tab-----
	If Fn_CPD_CompnentTabOperations("IsMaximized","4GD Compare","") Then
		bFlag = Fn_CPD_CompnentTabOperations("DoubleClick","4GD Compare","")
		If bFlag = False Then
			Set objWindow = Nothing
			Set objTable = Nothing
			Exit Function
		End If
	End If
	'---------------------------------
End Function

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name	:	Fn_CPD_AdvancedAccountabilityCheck_Ops
'@@
'@@    Description		:	Function Used to perform operations on "Advanced Accountability Check" window
'@@
'@@    Parameters		:	1. sAction		: Action to be performed
'@@						:	2. dicDetails	: Dictionary object
'@@						:	3. sButton		: OK / Cancel button
'@@
'@@    Return Value		: 	True Or False
'@@
'@@    Examples			:	Set dicDetails = CreateObject("Scripting.Dictionary")
'@@    							dicDetails("SelectSourceObject") = "PTN000001/001;1-static_ptn"
'@@    							dicDetails("SelectTargetObject") = "PTN000002/001;1-Dynamin_ptn"
'@@    							dicDetails("SelectMainTab1") = "Reporting"
'@@    							dicDetails("SetColorCheckBox") = "SetAllON"
'@@    							dicDetails("SelectMainTab2") = "Partial Match"
'@@    							dicDetails("SelectInternalTab") = "4GD DE/DF Properties"
'@@    							dicDetails("AddProperties") = "Absolute Transform~Axis Direction X~Type~Axis Direction Y~Axis Direction Z~Volume"
'@@    							dicDetails("RemoveProperties") = "Axis Direction Z~Volume"
'@@    							dicDetails("VerifySourceObject") = "PTN000001/001;1-static_ptn"
'@@    							dicDetails("VerifyTargetObject") = "PTN000002/001;1-Dynamin_ptn"
'@@    						bReturn = Fn_CPD_AdvancedAccountabilityCheck_Ops("Compare",dicDetails,"OK")
'@@    							
'@@	   History			:	
'@@	  	Developer Name		Date	   Rev. No.		Changes Done										Reviewer
'@@-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@	  	Vivek Ahirrao	 24-Aug-2016	1.0			Created - Added for 4GD new TC's development		[TC1123-20160729-24_08_2016-VivekA-NewDevelopment]
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_CPD_AdvancedAccountabilityCheck_Ops(sAction,dicDetails,sButton)
	GBL_FAILED_FUNCTION_NAME="Fn_CPD_AdvancedAccountabilityCheck_Ops"
	Dim objAACWindow, objDevReplay
	Dim dicCount, dicItems, dicKeys, aProperty, aChkBox, asProperty
	Dim iCounter, iCount, bFlag, iBtnIndex
	Dim sSubAction, sProperty, sTab, sTableName
	
	Const VK_CONTROL = 29
	
	Fn_CPD_AdvancedAccountabilityCheck_Ops = False
	'On Error Resume Next
	
	Set objAACWindow = JavaWindow("Collaborative Product").JavaWindow("AdvancedAccountabilityCheck")
	
	Select Case sAction
		Case "Compare"
				dicCount = dicDetails.Count
				dicItems = dicDetails.Items
				dicKeys = dicDetails.Keys
				For iCounter = 0 To dicCount - 1
					If Instr(dicKeys(iCounter),"SelectMainTab")>0 Then
						sSubAction = "SelectMainTab"
					ElseIf Instr(dicKeys(iCounter),"SelectInternalTab")>0 Then
						sSubAction = "SelectInternalTab"
					Else
						sSubAction = dicKeys(iCounter)
					End If
					
					sProperty = dicItems(iCounter)
					bFlag = False
					Select Case sSubAction
						'Select MainTab or InternalTab in Partial Match tab
						Case "SelectMainTab","SelectInternalTab"
							If sProperty<>"" Then
								If sSubAction = "SelectMainTab" Then
									sTab = "MainTab"
								ElseIf sSubAction = "SelectInternalTab" Then
									sTab = "InternalTab"
								End If
								bFlag = Fn_UI_JavaTab_Select("Fn_CPD_AdvancedAccountabilityCheck_Ops",objAACWindow,sTab,sProperty)
								If bFlag = False Then
									Call Fn_WriteLogFile("","FAIL: [Fn_CPD_AdvancedAccountabilityCheck_Ops]-Action-["+sAction+"]-SubAction-["+sSubAction+"] : Failed to Select ["+sTab+"] Tab in [AdvancedAccountabilityCheck] window.")
									Set objAACWindow = Nothing
									Exit Function
								End If
							End If
						Case "SelectSelectedTargetObject"
							If objAACWindow.Exist(1) Then
								Call Fn_UI_JavaTab_Select("Fn_CPD_AdvancedAccountabilityCheck_Ops",objAACWindow,"MainTab","Scope")
							End If
						
							If Fn_Button_Click("Fn_CPD_AdvancedAccountabilityCheck_Ops", objAACWindow, "SetAddCurrentSelection") = True Then
								bFlag = True
							End If 
							
						Case "SelectSourceObject","SelectTargetObject"
							If sProperty<>"" Then
								If objAACWindow.Exist(1) Then
									Call Fn_UI_JavaTab_Select("Fn_CPD_AdvancedAccountabilityCheck_Ops",objAACWindow,"MainTab","Scope")
								End If
								If sSubAction="SelectTargetObject" Then
									'Check whether Target Object list is empty or not
									objAACWindow.JavaList("SourceTargetObjectList").SetTOProperty "attached text","Target Object:"
									Wait 0,100
									If objAACWindow.JavaList("SourceTargetObjectList").GetROProperty("items count")="0" Then
										bFlag = True
									Else
										If objAACWindow.JavaList("SourceTargetObjectList").GetItem("0")<>sProperty Then
											bFlag = True
										Else
											bFlag = False
										End If
									End If
								ElseIf sSubAction="SelectSourceObject" Then
									bFlag = True
								End If
								
								If bFlag = True Then
									'To set Focus on Content Explorer tab first minimize the [ AdvancedAccountabilityCheck ] window
									If objAACWindow.Exist(1) Then
										If objAACWindow.GetROProperty("minimized")<>1 Then
											objAACWindow.Minimize
											Wait 1
										End If
									End If
									'Activate Source/Target Partition Tab in Content Explorer View
									bFlag = Fn_CPD_CompnentTabOperations("Activate",sProperty&" (Content Explorer)","")
									If bFlag = False Then
										Call Fn_WriteLogFile("","FAIL: [Fn_CPD_AdvancedAccountabilityCheck_Ops]-Action-["+sAction+"]-SubAction-["+sSubAction+"] : Failed to Select ["+sProperty+" (Content Explorer)] Tab in [Content Explorer View].")
										Set objAACWindow = Nothing
										Exit Function
									End If
									Wait 0,500
									'Select Source/Target Partition in Content Explorer View
									bFlag = Fn_CPD_ContentExplorer("Select",sProperty,"","","")
									If bFlag = False Then
										Call Fn_WriteLogFile("","FAIL: [Fn_CPD_AdvancedAccountabilityCheck_Ops]-Action-["+sAction+"]-SubAction-["+sSubAction+"] : Failed to Select ["+sProperty+"] Node in [Content Explorer View].")
										Set objAACWindow = Nothing
										Exit Function
									End If
									Wait 0,500
									'Again Restore the window
									If objAACWindow.Exist(1) Then
										If objAACWindow.GetROProperty("minimized")=1 Then
											objAACWindow.Restore
											Wait 1
										End If
									End If
									
									If sSubAction="SelectTargetObject" Then
										'Click on Set/Add Current selection button to add Target Object in list
										bFlag = Fn_Button_Click("Fn_CPD_AdvancedAccountabilityCheck_Ops",objAACWindow,"SetAddCurrentSelection")
										If bFlag=False Then
											Call Fn_WriteLogFile("","FAIL: [Fn_CPD_AdvancedAccountabilityCheck_Ops]-Action-["+sAction+"]-SubAction-["+sSubAction+"] : Failed to Click on [SetAddCurrentSelection] button in [AdvancedAccountabilityCheck] window.")
											Set objAACWindow = Nothing
											Exit Function
										End If
									End If
								End If
								
								If sSubAction="SelectSourceObject" Then
									'Select Menu call for [ Tools:Accountability Check...:Compare... ] menu
									bFlag = Fn_MenuOperation("Select","Tools:Accountability Check...:Compare...")
									If bFlag = False Then
										Call Fn_WriteLogFile("","FAIL: [Fn_CPD_AdvancedAccountabilityCheck_Ops]-Action-["+sAction+"]-SubAction-["+sSubAction+"] : Failed to Select Menu [Tools:Accountability Check...:Compare...].")
										Set objAACWindow = Nothing
										Exit Function
									End If
									Wait 0,500
									If objAACWindow.Exist(5) = False Then
										Call Fn_WriteLogFile("","FAIL: [Fn_CPD_AdvancedAccountabilityCheck_Ops]-Action-["+sAction+"]-SubAction-["+sSubAction+"] : Failed as [AdvancedAccountabilityCheck] window does not Exist.")
										Set objAACWindow = Nothing
										Exit Function
									End If
								ElseIf sSubAction="SelectTargetObject" Then
									bFlag = True
								End If
							End If
						'Case to Verify Source Object & Target Object list node
						Case "VerifySourceObject","VerifyTargetObject"
							If sProperty<>"" Then
								If objAACWindow.Exist(1) Then
									Call Fn_UI_JavaTab_Select("Fn_CPD_AdvancedAccountabilityCheck_Ops",objAACWindow,"MainTab","Scope")
								End If
								If sSubAction="VerifySourceObject" Then
									sListName = "Source Object:"
								ElseIf sSubAction="VerifyTargetObject" Then
									sListName = "Target Object:"
								End If
								objAACWindow.JavaList("SourceTargetObjectList").SetTOProperty "attached text",sListName
								If objAACWindow.JavaList("SourceTargetObjectList").Exist(1) = False Then
									Call Fn_WriteLogFile("","FAIL: [Fn_CPD_AdvancedAccountabilityCheck_Ops]-Action-["+sAction+"]-SubAction-["+sSubAction+"] : Failed to List ["+sListName+"] does not Exist in [AdvancedAccountabilityCheck] window.")
									Set objAACWindow = Nothing
									Exit Function
								End If
								
								'Verify Node in list
								iItemCount = CInt(objAACWindow.JavaList("SourceTargetObjectList").GetROProperty("items count"))
								For iCount = 0 To iItemCount-1
									sAppNode = objAACWindow.JavaList("SourceTargetObjectList").GetItem(iCount)
									If sAppNode<>sProperty Then
										Call Fn_WriteLogFile("","FAIL: [Fn_CPD_AdvancedAccountabilityCheck_Ops]-Action-["+sAction+"]-SubAction-["+sSubAction+"] : Failed as Node ["+sProperty+"] does not Exist in ["+sListName+"] list.")
										Set objAACWindow = Nothing
										Exit Function
									End If
								Next
								bFlag = True
							End If
						'Case to Set Color check boxes values as ON or OFF in Reporting tab
						Case "SetColorCheckBox"
							If sProperty<>"" Then
								Call Fn_UI_JavaTab_Select("Fn_CPD_AdvancedAccountabilityCheck_Ops",objAACWindow,"MainTab","Reporting")
								If sProperty = "SetAllON" Then
									sProperty = "Color the compared objects:ON~Full match:ON~Partial match:ON~Missing target:ON~Missing source:ON~Multiple match:ON~Multiple partial match:ON"
									aProperty = Split(sProperty,"~")
								Else
									aProperty = Split(sProperty,"~")
								End If
								'Set color check boxes values as ON or OFF
								For iCount = 0 To UBound(aProperty)
									aChkBox = Split(aProperty(iCount),":")
									objAACWindow.JavaCheckBox("ColorCheckBox").SetTOProperty "attached text",aChkBox(0)
									bFlag = Fn_CheckBox_Set("Fn_CPD_AdvancedAccountabilityCheck_Ops",objAACWindow,"ColorCheckBox",aChkBox(1))
									If bFlag=False Then
										Call Fn_WriteLogFile("","FAIL: [Fn_CPD_AdvancedAccountabilityCheck_Ops]-Action-["+sAction+"]-SubAction-["+sSubAction+"] : Failed to Set Checkbox ["+aChkBox(0)+"] value as ["+aChkBox(1)+"] in [AdvancedAccountabilityCheck] window.")
										Set objAACWindow = Nothing
										Exit Function
									End If
									Wait 0,500
								Next
							End If
						'Case to Add properties in 
						Case "AddProperties","RemoveProperties"
							If sProperty<>"" Then
								'Set Checkbox "Consider values of properties when searching for a partial match" value as ON
								objAACWindow.JavaCheckBox("ColorCheckBox").SetTOProperty "attached text","Consider values of properties when searching for a partial match"
								bFlag = Fn_CheckBox_Set("Fn_CPD_AdvancedAccountabilityCheck_Ops",objAACWindow,"ColorCheckBox","ON")
								If bFlag=False Then
									Call Fn_WriteLogFile("","FAIL: [Fn_CPD_AdvancedAccountabilityCheck_Ops]-Action-["+sAction+"]-SubAction-["+sSubAction+"] : Failed to Set Checkbox [Consider values of properties when searching for a partial match] value as [ON] in [AdvancedAccountabilityCheck] window.")
									Set objAACWindow = Nothing
									Exit Function
								End If
								
								If sSubAction="AddProperties" Then
									sTableName = "Available Properties:"	'or it could be "Available Attribute Group properties:"
									iBtnIndex = 0
								ElseIf sSubAction="RemoveProperties" Then
									sTableName = "Selected Properties:"	'or it could be "Selected Attribute Group properties:"
									iBtnIndex = 1
								End If
								
								Set objDevReplay = CreateObject("Mercury.DeviceReplay")
								'Add Properties from Available Properties Table to Selected Properties Table
								objAACWindow.JavaTable("PropertiesTable").SetTOProperty "attached text",sTableName
								aProperty = Split(sProperty,"~")
								iClickCounter = 0
								For iCount = 0 To UBound(aProperty)
									bFlag = Fn_SISW_UI_JavaTable_Operations("Fn_CPD_AdvancedAccountabilityCheck_Ops","Exist",objAACWindow,"PropertiesTable","",0,aProperty(iCount),"","","","")
									If bFlag=True Then
										bFlag = Fn_SISW_UI_JavaTable_Operations("Fn_CPD_AdvancedAccountabilityCheck_Ops","ClickCell",objAACWindow,"PropertiesTable","",0,aProperty(iCount),"","","","")
										If bFlag=False Then
											Call Fn_WriteLogFile("","FAIL: [Fn_CPD_AdvancedAccountabilityCheck_Ops]-Action-["+sAction+"]-SubAction-["+sSubAction+"] : Failed to Select ["+aProperty(iCount)+"] in ["+sTableName+"] table in [AdvancedAccountabilityCheck] window.")
											Set objAACWindow = Nothing
											Set objDevReplay = Nothing
											Exit Function
										End If
										iClickCounter = iClickCounter+1
									End If
									If iClickCounter=1 Then
										objDevReplay.KeyDown VK_CONTROL
									End If
								Next
								If iClickCounter>0 Then
									objDevReplay.KeyUp VK_CONTROL
									Set objDevReplay = Nothing
									'Click on Add or Remove button as per SubCase
									objAACWindow.JavaButton("AddRemoveProperties").SetTOProperty "Index",iBtnIndex
									bFlag = Fn_Button_Click("Fn_CPD_AdvancedAccountabilityCheck_Ops",objAACWindow,"AddRemoveProperties")
									If bFlag=False Then
										Call Fn_WriteLogFile("","FAIL: [Fn_CPD_AdvancedAccountabilityCheck_Ops]-Action-["+sAction+"]-SubAction-["+sSubAction+"] : Failed to Click on [AddRemoveProperties] button in [AdvancedAccountabilityCheck] window.")
										Set objAACWindow = Nothing
										Exit Function
									End If
								End If
								bFlag = True
							End If
						'Case to Verify properties in Available or Selected Properties Table
						'dicDetails("VerifyProperties") = "AvailableProperties#NX Entity Handle~Name~Resistance"
						Case "VerifyProperties"
							If sProperty<>"" Then
								'Set Checkbox "Consider values of properties when searching for a partial match" value as ON
								objAACWindow.JavaCheckBox("ColorCheckBox").SetTOProperty "attached text","Consider values of properties when searching for a partial match"
								bFlag = Fn_CheckBox_Set("Fn_CPD_AdvancedAccountabilityCheck_Ops",objAACWindow,"ColorCheckBox","ON")
								If bFlag=False Then
									Call Fn_WriteLogFile("","FAIL: [Fn_CPD_AdvancedAccountabilityCheck_Ops]-Action-["+sAction+"]-SubAction-["+sSubAction+"] : Failed to Set Checkbox [Consider values of properties when searching for a partial match] value as [ON] in [AdvancedAccountabilityCheck] window.")
									Set objAACWindow = Nothing
									Exit Function
								End If
								
								asProperty = Split(sProperty,"#")
								If asProperty(0)="AvailableProperties" Then
									sTableName = "Available Properties:"	'or it could be "Available Attribute Group properties:"
								ElseIf asProperty(0)="SelectedProperties" Then
									sTableName = "Selected Properties:"	'or it could be "Selected Attribute Group properties:"
								End If
								
								'Verify Properties from Available or Selected Properties Table
								objAACWindow.JavaTable("PropertiesTable").SetTOProperty "attached text",sTableName
								aProperty = Split(asProperty(1),"~")
								For iCount = 0 To UBound(aProperty)
									bFlag = Fn_SISW_UI_JavaTable_Operations("Fn_CPD_AdvancedAccountabilityCheck_Ops","Exist",objAACWindow,"PropertiesTable","",0,aProperty(iCount),"","","","")
									If bFlag=False Then
										Call Fn_WriteLogFile("","FAIL: [Fn_CPD_AdvancedAccountabilityCheck_Ops]-Action-["+sAction+"]-SubAction-["+sSubAction+"] : Failed to Verify ["+aProperty(iCount)+"] in ["+sTableName+"] table in [AdvancedAccountabilityCheck] window.")
										Set objAACWindow = Nothing
										Exit Function
									End If
								Next
								bFlag = True
							End If
					End Select
					
					If bFlag = False Then
						Fn_CPD_AdvancedAccountabilityCheck_Ops = False
						Set objAACWindow = Nothing
						Call Fn_WriteLogFile("","FAIL : Function [Fn_CPD_AdvancedAccountabilityCheck_Ops] Failed to Perform Case ["&sAction&"] SubCase ["+sSubAction+"].")
						Exit Function
					End If
				Next
				If sButton<>"" Then
					bFlag = Fn_Button_Click("Fn_CPD_AdvancedAccountabilityCheck_Ops",objAACWindow,sButton)
					If bFlag=False Then
						Call Fn_WriteLogFile("","FAIL: [Fn_CPD_AdvancedAccountabilityCheck_Ops]-Action-["+sAction+"]-SubAction-["+sSubAction+"] : Failed to Click on ["+sButton+"] button in [AdvancedAccountabilityCheck] window.")
						Set objAACWindow = Nothing
						Exit Function
					End If
				End If
				Fn_CPD_AdvancedAccountabilityCheck_Ops = True
		Case Else
				Set objAACWindow = Nothing
				Exit Function
	End Select
	Set objAACWindow = Nothing
End Function

'@@------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@
'@@    Function Name		:	Fn_ClickEffectivityTreeCell
'@@
'@@    Description			:	Function Used to Click on particular Cell of TreeTable
'@@
'@@    Parameters			:	1. sCallerFunctionName = Caller Function's name
'@@								2. objDialog = parent Java Dialog / Window / Applet 
'@@								3. sTree = Tree Object Name 
'@@								4. sNode = Full Node Path
'@@								5. sColumn = Column Name 
'@@								6. sButton = Button Name ( "LEFT" / "RIGHT" )
'@@
'@@    Return Value		   	: 	True / False
'@@
'@@    Pre-requisite		:	Tree Should Exist							
'@@
'@@    Examples				:	Call Fn_ClickEffectivityTreeCell("Fn_CPD_RecipeOperations",JavaWindow("Collaborative Product"), "EffectivityTree","DE000176/001;1-d2","From Unit","LEFT")
'@@    Examples				:	Call Fn_ClickEffectivityTreeCell("Fn_CPD_RecipeOperations",JavaWindow("Collaborative Product"), "EffectivityTree","DE000176/001;1-d2","TO Unit","RIGHT")
'@@
'@@	   History				:	
'@@					Developer Name					Date				Rev. No.			Changes Done								Reviewer
'@@----------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@					Jotiba T				     24-May-2012			  1.0				Created											
'@@----------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function Fn_ClickEffectivityTreeCell(sCallerFunctionName, objDialog, sTree, sNode, sColumn, sButton)
	Dim intX, intY, iCnt, iIterate, objTree, i
	Dim sColName, sItemName, sUIFail
	Dim iColWidth, iItmHeight, iCount
	bGblFailedFunctionName = "Fn_ClickEffectivityTreeCell"
	sUIFail = sCallerFunctionName + ">> Fn_Menu_Select >> " +  objDialog.toString + " >>  Tree[ " & sTree & " ] "
	Fn_ClickEffectivityTreeCell = False
	intY = 0
	intX = 0
	Set objTree = objDialog.JavaTree(sTree)
	If Fn_UI_ObjectExist("Fn_UI_ClickJavaTreeCell", objTree) = False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_UI_ClickJavaTreeCell : FAIL : " & sUIFail & " Doesn not exist.")	
		Exit function
	End If
	If isNumeric(trim(sColumn)) = False Then
		For iCnt = 0 to objTree.GetROProperty("columns_count")-1
			sColName = objTree.GetColumnHeader(iCnt)
			If trim(LCase(sColName)) = trim(LCase(sColumn)) Then
				Exit For
			End If
		Next
	Else
		iCnt = Cint(trim(sColumn))
	End If
	For iIterate = 0 to iCnt
		iColWidth = objTree.Object.getColumn(iIterate).getWidth()
		intX = intX + iColWidth
	Next
	intX = intX - iColWidth/2

	If isNumeric(trim(sNode)) = Flase Then
			iCount = objTree.GetROProperty("count_all_items")
			For i=0 To iCount 
				If objTree.GetItem(i)=sNode Then
					iCnt=i
					Exit for 
				End If	
			Next
		If iCnt = -1 Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_UI_ClickJavaTreeCell : FAIL : Failed to find node " & sNode)	
			Exit function 
		End If
	Else
		iCnt = Cint(trim(sNode))
	End If
	
	For iIterate = 0 to iCnt
		iItmHeight = objTree.Object.getItemHeight()
		intY = intY + iItmHeight
	Next
	intY = intY - iItmHeight/2
	
	Select Case lcase(sButton)
		Case "left"
			objTree.Click intX, intY,"LEFT"
			'Added by Sandeep : Some times its not work with single click so added this workaround : 20-May-2013
			wait 1
		Case "right"
			objTree.Click intX, intY,"RIGHT"
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_UI_ClickJavaTreeCell : FAIL : Failed  with [ " & sButton & " Click].")
			Exit Function
	End Select
	If Err.Number < 0 Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_UI_ClickJavaTreeCell : FAIL : Failed  with [ " & sButton & " Click].")
		Exit Function
	End If
	Fn_ClickEffectivityTreeCell = True
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_UI_ClickJavaTreeCell : PASS : Executed successfully with [ " & sButton & " Click].")	
End Function

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name			:	Fn_CPD_ViewEditMappings_Operations
'@@
'@@    Description				:	Function Used Perform operation on View/Edit Mappings
'@@
'@@    Parameters			    :	1. sAction		: Action to be performed
'@@								:	2. dicDetails	: Dictionary object
'@@								:	3. sClose 		: Yes/no
'@@
'@@    Return Value		   	   	: 	True Or False
'@@
'@@    Pre-requisite			:	CPD perspective should be activated and Preview 4GD tab should be opened						
'@@
'@@    Examples					:	Set dicDetails = CreateObject("Scripting.Dictionary") 
'@@										dicDetails("MapDatasetName") = "Test_Map7"
'@@										dicDetails("AssemblyTypeMappingProperties") = "bl_item_object_type:Item:Cpd0DesignElement"
'@@										dicDetails("SynchronizeVariant") = "OFF"
'@@										dicDetails("SynchronizeEffectivity") = "OFF"
'@@									bReturn = Fn_CPD_ViewEditMappings_Operations("Create",dicDetails,"Yes")
'@@	
'@@									Set dicDetails = CreateObject("Scripting.Dictionary") 
'@@										dicDetails("MapDatasetName") = "Test_Map7"
'@@										dicDetails("SynchronizeVariant") = "ON"
'@@										dicDetails("SynchronizeEffectivity") = "ON"
'@@									bReturn = Fn_CPD_ViewEditMappings_Operations("OpenMapDatasetAndModify",dicDetails,"Yes")
'@@
'@@									Set dicDetails = CreateObject("Scripting.Dictionary") 
'@@										dicDetails("MapDatasetName") = "Test_Map7"
'@@										dicDetails("AssemblyTypeMappingProperties") = "bl_item_object_type:Item:Cpd0DesignElement"
'@@										dicDetails("SynchronizeVariant") = "OFF"
'@@										dicDetails("SynchronizeEffectivity") = "OFF"
'@@									bReturn = Fn_CPD_ViewEditMappings_Operations("OpenMapDatasetAndVerify",dicDetails,"Yes")
'@@
'@@	   History					:	Developer Name		Date			Rev. No.	Changes Done		Reviewer
'@@-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@									Poonam Chopade		03-Jan-2018		1.0			Created			 TC11.4(2017120100)_NewDevelopment_PoonamC_03Jan2018
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_CPD_ViewEditMappings_Operations(sAction,dicDetails,sClose)
	GBL_FAILED_FUNCTION_NAME="Fn_CPD_ViewEditMappings_Operations"
	
	Dim ObjCPDWin,aOrgTypeDtls,sProp,sPropVal,sPropType,objTable,iRORows
	Dim iVpCnt,bFlag,iCnt
	
	Fn_CPD_ViewEditMappings_Operations = False
	Set ObjCPDWin = Fn_SISW_CPD_GetObject("Collaborative Product")
	
	'Check View/Edit Mappings tab opened and open it
	If Fn_CPD_CompnentTabOperations("Exists","View/Edit Mappings", "") = False Then 
		sMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("4GD_Toolbar"),"ViewEditMappings")
		Call Fn_ToolbarButtonClick_Ext(1,sMenu)
		Call Fn_ReadyStatusSync(3)
	Else
	    'Activate Tab
		Call Fn_CPD_CompnentTabOperations("Activate","View/Edit Mappings", "")
		Call Fn_ReadyStatusSync(1)
	End If 
	
	'Maximize tab
	If Fn_CPD_CompnentTabOperations("IsMaximized","View/Edit Mappings", "") = False Then  
		Call Fn_CPD_CompnentTabOperations("DoubleClick","View/Edit Mappings", "")  
		Call Fn_ReadyStatusSync(2)
	End If 
	
	Select Case sAction
		Case "Create","OpenMapDatasetAndModify","CreateWithoutSave"
			If sAction = "Create" or sAction = "CreateWithoutSave" Then
				If dicDetails("MapDatasetName") <> "" Then
					Call Fn_SISW_UI_JavaEdit_Operations("Fn_CPD_ViewEditMappings_Operations", "Type", ObjCPDWin, "MappingDataset", dicDetails("MapDatasetName"))
					Wait 1	
					'Click on Create
			   		 Call Fn_Button_Click("Fn_CPD_ViewEditMappings_Operations",ObjCPDWin, "CreateMappingDataset")
			   		 Wait 1
				End If
			ElseIf sAction = "OpenMapDatasetAndModify" Then
				'Open Mapping dataset
				If dicDetails("MapDatasetName") <> "" Then
					Call Fn_Button_Click("Fn_CPD_ViewEditMappings_Operations",ObjCPDWin, "MappingDataset")
					Wait 1
					Call Fn_OpenByNameOperations("CellDoubleClick",dicDetails("MapDatasetName"), "","0","Object",dicDetails("MapDatasetName"))
					wait 1
				End If
			End If
			
		    'Select Organization Type Mapping
		    If dicDetails("OrgTypeMapping") <> "" Then
		    	Call Fn_SISW_UI_JavaList_Operations("Fn_CPD_ViewEditMappings_Operations", "Select", ObjCPDWin, "Organization Type Mapping", dicDetails("OrgTypeMapping"), "", "")
		    End If
			
			'Select Organization Type Mapping
		    If dicDetails("OrgTypeMappingName") <> "" Then
		    	Call Fn_SISW_UI_JavaEdit_Operations("Fn_CPD_ViewEditMappings_Operations","Set",ObjCPDWin,"SchemeName",dicDetails("OrgTypeMappingName"))
		    End If
		    
		    'Add Organization Type Mapping details
		    If dicDetails("OrgTypeMappingProperties") <> "" Then
					'Click on Add button
					Call Fn_Button_Click("Fn_CPD_ViewEditMappings_Operations",ObjCPDWin, "OrgTypeMapAdd")
		    		Wait 1

			    	aOrgTypeDtls = Split(dicDetails("OrgTypeMappingProperties"),":")
			    	sProp = aOrgTypeDtls(0)
			    	sPropVal = aOrgTypeDtls(1)
			    	sPropType = aOrgTypeDtls(2)
		    	
			    	'Select Source Property
			    	Set objTable = ObjCPDWin.JavaTable("Organization Type Mapping")
					iRORows = objTable.GetROProperty("rows") - 1
					objTable.SelectCell iRORows,"Source Property"
					Wait(3)
					objTable.SelectCell iRORows,"Source Property"
					ObjCPDWin.JavaList("PropValTypeList").Select sProp
			    	
			    	'Select Source Propety Value
					objTable.SelectCell iRORows,"Source Property Value"
					Wait(3)
					objTable.SelectCell iRORows,"Source Property Value"
					ObjCPDWin.JavaList("PropValTypeList").Select sPropVal
		    	   
		    	    'Select Target 4GD Types
					objTable.SelectCell iRORows,"Target 4GD Types"
					Wait(3)
					objTable.SelectCell iRORows,"Target 4GD Types"
					ObjCPDWin.JavaList("PropValTypeList").Select sPropType
		    End If
		    
		     'Add Assembly Type Mapping details
		    If dicDetails("AssemblyTypeMappingProperties") <> "" Then
					'Click on Add button
					Call Fn_Button_Click("Fn_CPD_ViewEditMappings_Operations",ObjCPDWin, "AssmTypeMapAdd")
		    		Wait 1

			    	aOrgTypeDtls = Split(dicDetails("AssemblyTypeMappingProperties"),":")
			    	sProp = aOrgTypeDtls(0)
			    	sPropVal = aOrgTypeDtls(1)
			    	sPropType = aOrgTypeDtls(2)
		    	
			    	'Select Source Property
			    	Set objTable = ObjCPDWin.JavaTable("Assembly Type Mapping")
					iRORows = objTable.GetROProperty("rows") - 1
					objTable.SelectCell iRORows,"Source Property"
					Wait(3)
					objTable.SelectCell iRORows,"Source Property"
					ObjCPDWin.JavaList("PropValTypeList").Select sProp
			    	
			    	'Select Source Propety Value
					objTable.SelectCell iRORows,"Source Property Value"
					Wait(3)
					objTable.SelectCell iRORows,"Source Property Value"
					ObjCPDWin.JavaList("PropValTypeList").Select sPropVal
		    	   
		    	    'Select Target 4GD Types
					objTable.SelectCell iRORows,"Target 4GD Types"
					Wait(3)
					objTable.SelectCell iRORows,"Target 4GD Types"
					ObjCPDWin.JavaList("PropValTypeList").Select sPropType
		    End If
		    
		    'Set Synchronize Variant option
		    If dicDetails("SynchronizeVariant") <> "" Then
		    	Call Fn_SISW_UI_JavaCheckBox_Operations("Fn_CPD_ViewEditMappings_Operations", "Set", ObjCPDWin, "SynchronizeVariant", dicDetails("SynchronizeVariant"))
		    End If
		    
		    'Set Synchronize Effectivity option
		    If dicDetails("SynchronizeEffectivity") <> "" Then
		    	Call Fn_SISW_UI_JavaCheckBox_Operations("Fn_CPD_ViewEditMappings_Operations", "Set", ObjCPDWin, "SynchronizeEffectivity", dicDetails("SynchronizeEffectivity"))
		    End If
		    
		    If sAction <> "CreateWithoutSave" Then
		    	'Save to create Mapping dataset
		   		 sMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("4GD_Toolbar"),"SaveTheCurrentContents")
				Call Fn_ToolbarButtonClick_Ext(1,sMenu)
				Call Fn_ReadyStatusSync(3)
			End If				
			
		Case "OpenMapDatasetAndVerify"
				iVpCnt = 0
				'Open Mapping dataset
				If dicDetails("MapDatasetName") <> "" Then
					Call Fn_Button_Click("Fn_CPD_ViewEditMappings_Operations",ObjCPDWin, "MappingDataset")
					Wait 1
					Call Fn_OpenByNameOperations("CellDoubleClick",dicDetails("MapDatasetName"), "","0","Object",dicDetails("MapDatasetName"))
					wait 1
					iVpCnt = iVpCnt + 1
				End If
				
				 'Select Organization Type Mapping
			    If dicDetails("OrgTypeMapping") <> "" Then
			    	If dicDetails("OrgTypeMapping") = Fn_SISW_UI_JavaList_Operations("Fn_CPD_ViewEditMappings_Operations", "GetText", ObjCPDWin, "Organization Type Mapping", "", "", "") Then
			    		iVpCnt = iVpCnt + 1
			    	End If 
			    End If
			    
			    'Verify Organization Type Mapping details
			    If dicDetails("OrgTypeMappingProperties") <> "" Then
			    
						Set objTable = ObjCPDWin.JavaTable("Organization Type Mapping")
				    	aOrgTypeDtls = Split(dicDetails("OrgTypeMappingProperties"),":")
			
			    		For iCnt = 0 To objTable.GetROProperty("rows") - 1
			    			bFlag = False
			    			sProp = objTable.GetCellData(iCnt,"Source Property") 
				    		sPropVal = objTable.GetCellData(iCnt,"Source Property Value")
				    		sPropType = objTable.GetCellData(iCnt,"Target 4GD Types")
							If sProp = aOrgTypeDtls(0) and sPropVal = aOrgTypeDtls(1) and sPropType = aOrgTypeDtls(2)  Then
								bFlag = True
								Exit For
							End If			    				
			    		Next
			    		If bFlag = True Then
			    			iVpCnt = iVpCnt + 1
			    		End If	
			    End If
			    
			     'Add Assembly Type Mapping details
			    If dicDetails("AssemblyTypeMappingProperties") <> "" Then
						
						Set objTable = ObjCPDWin.JavaTable("Assembly Type Mapping")
				    	aOrgTypeDtls = Split(dicDetails("AssemblyTypeMappingProperties"),":")
				    	For iCnt = 0 To objTable.GetROProperty("rows") - 1
			    			bFlag = False
			    			sProp = objTable.GetCellData(iCnt,"Source Property") 
				    		sPropVal = objTable.GetCellData(iCnt,"Source Property Value")
				    		sPropType = objTable.GetCellData(iCnt,"Target 4GD Types")
							If sProp = aOrgTypeDtls(0) and sPropVal = aOrgTypeDtls(1) and sPropType = aOrgTypeDtls(2)  Then
								bFlag = True
								Exit For
							End If			    				
			    		Next
			    		If bFlag = True Then
			    			iVpCnt = iVpCnt + 1
			    		End If
			    End If
			    
			    'Verify Synchronize Variant option
			    If dicDetails("SynchronizeVariant") <> "" Then
			    	If dicDetails("SynchronizeVariant") = "ON" Then
			    		If cbool(ObjCPDWin.JavaCheckBox("SynchronizeVariant").GetROProperty("value")) = true then
			    			iVpCnt = iVpCnt + 1
			    		End If
			    	ElseIf dicDetails("SynchronizeVariant") = "OFF" Then
			    		If cbool(ObjCPDWin.JavaCheckBox("SynchronizeVariant").GetROProperty("value")) = false then
			    			iVpCnt = iVpCnt + 1
			    		End If
			    	End If  	
			    End If
			    
			    'Verify Synchronize Effectivity option
			    If dicDetails("SynchronizeEffectivity") <> "" Then
			    	If dicDetails("SynchronizeEffectivity") = "ON" Then
			    		If cbool(ObjCPDWin.JavaCheckBox("SynchronizeEffectivity").GetROProperty("value")) = true then
			    			iVpCnt = iVpCnt + 1
			    		End If
			    	ElseIf dicDetails("SynchronizeEffectivity") = "OFF" Then
			    		If cbool(ObjCPDWin.JavaCheckBox("SynchronizeEffectivity").GetROProperty("value")) = false then
			    			iVpCnt = iVpCnt + 1
			    		End If
			    	End If 
			    End If
			    
				If cint(dicDetails.Count) <> cint(iVpCnt) Then
					Set ObjCPDWin = Nothing
					Set objTable = Nothing
					Fn_CPD_ViewEditMappings_Operations = False
					Exit Function
				End If
	End Select
	
	'Minimize tab
	If Fn_CPD_CompnentTabOperations("IsMaximized","View/Edit Mappings", "") = True Then  
		Call Fn_CPD_CompnentTabOperations("DoubleClick","View/Edit Mappings", "")  
		Call Fn_ReadyStatusSync(2)
	End If 
	
	'close Tab
	If sClose = "Yes" Then
		Call Fn_CPD_CompnentTabOperations("Close","View/Edit Mappings", "")
		Call Fn_ReadyStatusSync(3)
	End If
	
	Fn_CPD_ViewEditMappings_Operations = True
	Set ObjCPDWin = Nothing
	Set objTable = Nothing
	
End Function

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name			:	Fn_CPD_Preview4GDModel_Operations
'@@
'@@    Description				:	Function Used Perform operation on Preview 4G Model
'@@
'@@    Parameters			    :	1. sAction		: Action to be performed
'@@								:	2. dicViewEditMappingDtls	: Dictionary object with Mapping dataset creation
'@@								:	3. dicCreateModelDtls 		: Dictionary object with 4GD Model creation 
'@@								:	4. sClose 					: Yes/No 
'@@	
'@@    Return Value		   	   	: 	True Or False
'@@
'@@    Pre-requisite			:	CPD perspective should be activated 
'@@
'@@    Examples					:	Set dicViewEditMappingDtls = CreateObject("Scripting.Dictionary") 
'@@										dicViewEditMappingDtls("MapDatasetName") = "Test_Map7"
'@@										dicViewEditMappingDtls("AssemblyTypeMappingProperties") = "bl_item_object_type:Item:Cpd0DesignElement"
'@@										dicViewEditMappingDtls("SynchronizeVariant") = "OFF"
'@@										dicViewEditMappingDtls("SynchronizeEffectivity") = "OFF"
'@@										
'@@									Set dicCreateModelDtls = CreateObject("Scripting.Dictionary") 	
'@@										dicCreateModelDtls("SrcAssemblyRevision") = "DEMO_CAR"
'@@										dicCreateModelDtls("MappingDataset") = "Test_Map7"
'@@										dicCreateModelDtls("TargetModelName") = "Test_CD1"
'@@										
'@@									bReturn = Fn_CPD_Preview4GDModel_Operations("CreatePreview",dicViewEditMappingDtls,dicCreateModelDtls,"Yes")
'@@	
'@@								    Set dicCreateModelDtls = CreateObject("Scripting.Dictionary") 	
'@@										dicCreateModelDtls("SrcAssemblyRevision") = "DEMO_CAR"
'@@										dicCreateModelDtls("MappingDataset") = "Test_Map7"
'@@										dicCreateModelDtls("TargetModelName") = "Test_CD1"
'@@										
'@@									bReturn = Fn_CPD_Preview4GDModel_Operations("GeneratePreview",dicViewEditMappingDtls,dicCreateModelDtls,"Yes")
'@@
'@@	   History					:	Developer Name		Date		Rev. No.	Changes Done		Reviewer
'@@-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@									Poonam Chopade		03-Jan-2018		1.0			Created			 TC11.4(2017120100)_NewDevelopment_PoonamC_03Jan2018
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_CPD_Preview4GDModel_Operations(sAction,dicViewEditMappingDtls,dicCreateModelDtls,sClose)
	GBL_FAILED_FUNCTION_NAME="Fn_CPD_Preview4GDModel_Operations"
	
	Dim ObjCPDWin,bFlag,sAppValue,sMenu
	
	Fn_CPD_Preview4GDModel_Operations = False
	Set ObjCPDWin = Fn_SISW_CPD_GetObject("Collaborative Product")
	
	'Check View/Edit Mappings tab opened and open it
	If Fn_CPD_CompnentTabOperations("Exists","Preview 4G Model", "") = False Then 
		sMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("CPD_Menu"),"4GPopulate")
		Call Fn_MenuOperation("Select",sMenu)
		Call Fn_ReadyStatusSync(3)
	Else
	    'Activate Tab
		Call Fn_CPD_CompnentTabOperations("Activate","Preview 4G Model", "")
		Call Fn_ReadyStatusSync(1)
	End If 
	
	'Maximize tab
	If Fn_CPD_CompnentTabOperations("IsMaximized","Preview 4G Model", "") = False Then  
		Call Fn_CPD_CompnentTabOperations("DoubleClick","Preview 4G Model", "")  
		Call Fn_ReadyStatusSync(2)
	End If 
	
	Select Case sAction
		Case "CreatePreview"
			If vartype(dicViewEditMappingDtls) = "9" Then
				 bFlag = Fn_CPD_ViewEditMappings_Operations("Create",dicViewEditMappingDtls,"Yes")
				 Wait 2
			Else
				bFlag = True
			End If
			
		    If bFlag = True Then
		    
		    	'Maximize tab
				If Fn_CPD_CompnentTabOperations("IsMaximized","Preview 4G Model", "") = False Then  
					Call Fn_CPD_CompnentTabOperations("DoubleClick","Preview 4G Model", "")  
					Call Fn_ReadyStatusSync(2)
				End If	
		    
				'Check Create PreView radio button
				Call Fn_SISW_UI_JavaRadioButton_Operations("Fn_CPD_Preview4GDModel_Operations", "Set", ObjCPDWin, "CreatePreview", "ON")			
				Wait 1
				
				'Search Assembly revision
				If dicCreateModelDtls("SrcAssemblyRevision") <> "" Then
					'Click on checkbox
					Call Fn_Button_Click("Fn_CPD_Preview4GDModel_Operations",ObjCPDWin, "MappingDataset")
					Wait 1
					'Search revision and select it
					Call Fn_OpenByNameOperations("CellDoubleClick",dicCreateModelDtls("SrcAssemblyRevision"), "","0","Object",dicCreateModelDtls("SrcAssemblyRevision"))
					wait 1
				ElseIf dicCreateModelDtls("SrcAssemblyID") <> "" Then
					 'Click on checkbox
					 Call Fn_Button_Click("Fn_CPD_Preview4GDModel_Operations",ObjCPDWin, "MappingDataset")
					 Wait 1
					 'Search revision and select it
					 Call Fn_OpenByNameOperations("CellDoubleClick","",dicCreateModelDtls("SrcAssemblyID"),"0","Object",dicCreateModelDtls("SrcAssemblyID"))
					 wait 1
				End If
				
				'Search Mapping dataset
				If dicCreateModelDtls("MappingDataset") <> "" Then
					'Click on checkbox
					Call Fn_Button_Click("Fn_CPD_Preview4GDModel_Operations",ObjCPDWin, "OpenMappingDataset")
					Wait 1
					'Search revision and select it
					Call Fn_OpenByNameOperations("CellDoubleClick",dicCreateModelDtls("MappingDataset"), "","0","Object",dicCreateModelDtls("MappingDataset"))
					wait 1
				End If
				
				'Enter Target Model Name
				If dicCreateModelDtls("TargetModelName") <> "" Then
					Call Fn_SISW_UI_JavaEdit_Operations("Fn_CPD_Preview4GDModel_Operations", "Type", ObjCPDWin, "TargetModelName", dicCreateModelDtls("TargetModelName"))
					Wait 1
				End if
				
				'select Revision Rule
				If dicCreateModelDtls("RevisionRule") <> "" Then
			    	Call Fn_SISW_UI_JavaList_Operations("Fn_CPD_Preview4GDModel_Operations", "Select", ObjCPDWin, "RevisionRule", dicCreateModelDtls("RevisionRule"), "", "")
			    End If
			    
				'Select Target Model Type
				If dicCreateModelDtls("TargetModelType") <> "" Then
			    	Call Fn_SISW_UI_JavaList_Operations("Fn_CPD_Preview4GDModel_Operations", "Select", ObjCPDWin, "TargetModelType", dicCreateModelDtls("TargetModelType"), "", "")
			    End If
				
				'Click on Generate Mode Preview
				Call Fn_Button_Click("Fn_CPD_Preview4GDModel_Operations",ObjCPDWin, "GenerateModelPreview")
				Call Fn_ReadyStatusSync(2)
				Wait 2
				
				'Select generated Target Model 
				Call Fn_JavaTree_Select("",ObjCPDWin,"PreviewModelTree",dicCreateModelDtls("TargetModelName"))
				Wait 2
				
				'Click on Create 4GD Model
				Call Fn_Button_Click("Fn_CPD_Preview4GDModel_Operations",ObjCPDWin, "Create4GDModel")
				Call Fn_ReadyStatusSync(10)
				Call Fn_ReadyStatusSync(5)
				
				'Handle Message of creation
				If Fn_UI_ObjectExist("Fn_CPD_Preview4GDModel_Operations",ObjCPDWin.JavaWindow("Success")) = True Then
					sAppValue = Fn_UI_Object_GetROProperty("Fn_CPD_Preview4GDModel_Operations",ObjCPDWin.JavaWindow("Success").JavaEdit("Text"),"value")
					If instr(sAppValue,"Model created successfully.") > 0 Then
						Fn_CPD_Preview4GDModel_Operations = True
						Call Fn_Button_Click("Fn_CPD_Preview4GDModel_Operations",ObjCPDWin.JavaWindow("Success"), "OK")
						Wait 1
					End If
				Else
					Fn_CPD_Preview4GDModel_Operations = False
				End if
				
		 End If
		 '---------------- Generate Preview Only --------------------------------
		 Case "GeneratePreview","GeneratePreviewNoMinimize"
		 
			'Check Create PreView radio button
			Call Fn_SISW_UI_JavaRadioButton_Operations("Fn_CPD_Preview4GDModel_Operations", "Set", ObjCPDWin, "CreatePreview", "ON")			
			Wait 1
			
			'Search Assembly revision
			If dicCreateModelDtls("SrcAssemblyRevision") <> "" Then
				'Click on checkbox
				Call Fn_Button_Click("Fn_CPD_Preview4GDModel_Operations",ObjCPDWin, "MappingDataset")
				Wait 1
				'Search revision and select it
				Call Fn_OpenByNameOperations("CellDoubleClick",dicCreateModelDtls("SrcAssemblyRevision"), "","0","Object",dicCreateModelDtls("SrcAssemblyRevision"))
				wait 1
			ElseIf dicCreateModelDtls("SrcAssemblyID") <> "" Then
					 'Click on checkbox
					 Call Fn_Button_Click("Fn_CPD_Preview4GDModel_Operations",ObjCPDWin, "MappingDataset")
					 Wait 1
					 'Search revision and select it
					 Call Fn_OpenByNameOperations("CellDoubleClick","",dicCreateModelDtls("SrcAssemblyID"),"0","Object",dicCreateModelDtls("SrcAssemblyID"))
					 wait 1
				End If
			
			'Search Mapping dataset
			If dicCreateModelDtls("MappingDataset") <> "" Then
				'Click on checkbox
				Call Fn_Button_Click("Fn_CPD_Preview4GDModel_Operations",ObjCPDWin, "OpenMappingDataset")
				Wait 1
				'Search revision and select it
				Call Fn_OpenByNameOperations("CellDoubleClick",dicCreateModelDtls("MappingDataset"), "","0","Object",dicCreateModelDtls("MappingDataset"))
				wait 1
			End If
			
			'Enter Target Model Name
			If dicCreateModelDtls("TargetModelName") <> "" Then
				Call Fn_SISW_UI_JavaEdit_Operations("Fn_CPD_Preview4GDModel_Operations", "Type", ObjCPDWin, "TargetModelName", dicCreateModelDtls("TargetModelName"))
				Wait 1
			End if
			
			'select Revision Rule
			If dicCreateModelDtls("RevisionRule") <> "" Then
		    	Call Fn_SISW_UI_JavaList_Operations("Fn_CPD_Preview4GDModel_Operations", "Select", ObjCPDWin, "RevisionRule", dicCreateModelDtls("RevisionRule"), "", "")
		    End If
		    
			'Select Target Model Type
			If dicCreateModelDtls("TargetModelType") <> "" Then
		    	Call Fn_SISW_UI_JavaList_Operations("Fn_CPD_Preview4GDModel_Operations", "Select", ObjCPDWin, "TargetModelType", dicCreateModelDtls("TargetModelType"), "", "")
		    End If
			
			'Click on Generate Mode Preview
			Fn_CPD_Preview4GDModel_Operations = Fn_Button_Click("Fn_CPD_Preview4GDModel_Operations",ObjCPDWin, "GenerateModelPreview")
			Call Fn_ReadyStatusSync(2)
			Wait 2
			
	End Select
	
	'Minimize tab
	If sAction <> "GeneratePreviewNoMinimize" Then 
		If Fn_CPD_CompnentTabOperations("IsMaximized","Preview 4G Model", "") = True Then  
			Call Fn_CPD_CompnentTabOperations("DoubleClick","Preview 4G Model", "")  
			Call Fn_ReadyStatusSync(2)
		End If 
	End If
	
	'close Tab
	If sClose = "Yes" Then
		Call Fn_CPD_CompnentTabOperations("Close","Preview 4G Model", "")
		Call Fn_ReadyStatusSync(3)
	End If
	
	Set ObjCPDWin = Nothing
		
End Function
'@'=========================================================================================================================================================================
'@@
'@@    Function Name		:	Fn_CPD_PCA_VariantNatTable_VariantConditions_Operation
'@@
'@@    Description			:	Function Used to Perform operation on Variant Conditions tab
'@@
'@@    Parameters			:	1. sAction		: Action to be performed
'@@							:	2. sBOMLine		: Node to select from Content Explorer
'@@							:	2. StrTabName	: Tab Name
'@@							:	2. dicDetails	: Dictionary object
'@@							:	2. Popupmenu	: Popup menu name
'@@							:	3. sTabClose	: Flag to close Tab
'@@
'@@    Return Value		   	: 	True Or False
'@@
'@@    Pre-requisite		:	4G Perspective should be opened 
'@@
'@@    Examples				:	Set dicDetails = CreateObject( "Scripting.Dictionary")
'@@ 							dicDetails("ColumnName") = "000060/A;1-1200"
'@@ 							dicDetails("Options") = "1200~1600~Arial~std"
'@@ 							dicDetails("OptionFlag") = "Check~Check~None~Check"
'@@ 							bReturn = Fn_CPD_PCA_VariantNatTable_VariantConditions_Operation("SetVarConditionAndSave","000058;1-CD:000059/001;1-DE","000065-PC (Variant Conditions)",dicDetails,"","Yes")
'@@ 							
'@@ 							Set dicDetails = CreateObject( "Scripting.Dictionary")
'@@ 							dicDetails("ColumnName") = "Option"
'@@ 							dicDetails("Options") = "1200~1600~Arial~std"
'@@ 							bReturn = Fn_CPD_PCA_VariantNatTable_VariantConditions_Operation("VerifyOptionsInVariantCondition","","000065-PC (Variant Conditions)",dicDetails,"","Yes")
'@@
'@@-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@	   History				:	Developer Name				Date	  			Rev. No.		Changes Done		Reviewer
'@@-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@	   Created By 			:	Poonam Chopade			26-Aug-2019				1.0		  		 Created		  [Tc12.3(2019081900)-26Aug2019-PoonamC-NewDevelopment]
'@'=========================================================================================================================================================================
Public Function Fn_CPD_PCA_VariantNatTable_VariantConditions_Operation(sAction,sNodeName,StrTabName,dicDetails,Popupmenu,sTabClose)

	GBL_FAILED_FUNCTION_NAME = "Fn_CPD_PCA_VariantNatTable_VariantConditions_Operation"
	
	Dim obCPDWindow,arrOptionFlag,arrOptions,iColIndex,strColumn,iColCount,strPopupMenu
	Dim sMenu,bFlag,iCnt,StrBounds,iX,iY,iCnt1,iRowIndex,iRowsCount,aOption,iInstance
		
	Fn_CPD_PCA_VariantNatTable_VariantConditions_Operation = False
	Set obCPDWindow = Fn_SISW_CPD_GetObject("Collaborative Product")
	
	On Error Resume Next
	If Fn_UI_ObjectExist("Fn_CPD_PCA_VariantNatTable_VariantConditions_Operation",obCPDWindow.JavaObject("VariantNatTable")) = False	Then
		If sNodeName <> "" Then 'Select Node From Content explorer
				sMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("CPD_PopupMenu"),"OpenWithVariantConditions")
				bFlag = Fn_CPD_ContentExplorer("popupmenuselect", sNodeName, "", "", sMenu) 
				If bFlag = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_CPD_PCA_VariantNatTable_VariantConditions_Operation ]  Failed to select Node [ "& sNodeName &" ]") 
					Fn_CPD_PCA_VariantNatTable_VariantConditions_Operation = False
					Set obCPDWindow = Nothing
					Exit function
				End If
				Call Fn_ReadyStatusSync(1)
		End If
		If Fn_UI_ObjectExist("Fn_CPD_PCA_VariantNatTable_VariantConditions_Operation",obCPDWindow.JavaObject("VariantNatTable")) = False	Then
			Fn_CPD_PCA_VariantNatTable_VariantConditions_Operation = False
			Set obCPDWindow = Nothing
			Exit function
		End If
	Else
		If sNodeName <> "" Then
			bFlag = Fn_CPD_ContentExplorer("select", sNodeName, "", "", "")
			Call Fn_ReadyStatusSync(1)			
			If bFlag = False Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_CPD_PCA_VariantNatTable_VariantConditions_Operation ]  Failed to select Node [ "& sNodeName &" ]") 
				Fn_CPD_PCA_VariantNatTable_VariantConditions_Operation = False
				Set obCPDWindow = Nothing
				Exit function
			End If
		End If	
	End If
	
	If StrTabName <> "" Then ' Maximize tab
		If Fn_CPD_CompnentTabOperations("IsMaximized",StrTabName, "") = False Then
			Call Fn_CPD_CompnentTabOperations("DoubleClick",StrTabName, "")
			Call Fn_ReadyStatusSync(1)
		End If
	End If
	
	Select Case sAction
		'========================================================================================================================================
		Case "SetVarConditionAndSave","SetVarConditionAndSave_Ext"
			 If dicDetails("ColumnName") <> "" Then  'Get column index
				iColCount = obCPDWindow.JavaObject("VariantNatTable").Object.getColumnCount()
				For iCnt = 0 To iColCount - 1
					strColumn =  CStr(obCPDWindow.JavaObject("VariantNatTable").Object.getCellByPosition(iCnt,0).getlayer.getColumnHeaderlayer.getDatavalueByPosition(iCnt,0).tostring)
					If trim(dicDetails("ColumnName")) = trim(strColumn) Then
							iColIndex = iCnt + 1
							Exit for
					End If
				Next
			 Else
				iColIndex = 2							 
			 End If
			 
			 arrOptions = Split(dicDetails("Options"),"~") ' Set Options as On or OFF
			 arrOptionFlag = Split(dicDetails("OptionFlag"),"~")
			 iRowsCount = obCPDWindow.JavaObject("VariantNatTable").Object.getRowCount()						 
			 For iCnt = 0 To UBound(arrOptions)
		 		If instr(arrOptions(iCnt),"@") Then  'for multiple instance
		 			aOption = Split(arrOptions(iCnt),"@")
		 			arrOptions(iCnt) = aOption(0)
		 			iInstance = aOption(1)
		 		Else
					iInstance = 1				 		
		 		End If
		 		iCounter = 1	
				For iCnt1 = 1 To iRowsCount - 1
					If instr(CStr(obCPDWindow.JavaObject("VariantNatTable").Object.getCellByPosition(0,iCnt1).getDataValue().getData().toString()),arrOptions(iCnt)) Then
						iRowIndex = iCnt1
						If cint(iCounter) = cint(iInstance) Then
			 				Exit For
			 			Else
							iCounter = iCounter + 1							 			
			 			End If
					End If	
				Next
				
				StrBounds = obCPDWindow.JavaObject("VariantNatTable").Object.getBoundsByPosition(iColIndex,iRowIndex).tostring
				StrBounds = Split(Replace(Replace(StrBounds,"Rectangle {",""),"}",""),",")
				iX = cint(StrBounds(2))
				iY = cint(StrBounds(1))
				
				If IsEmpty(iX) or IsEmpty(iY) Then
					Fn_CPD_PCA_VariantNatTable_VariantConditions_Operation = False
					Set obCPDWindow = Nothing
					Exit Function
				End If
				iX = iX + 4
				iY = iY + 5
				If sAction = "SetVarConditionAndSave_Ext" Then  'For Boolean Value to check in subject section for same value
							StrBounds = obCPDWindow.JavaObject("VariantNatTable").Object.getBoundsByPosition(iColIndex,iRowIndex+1).tostring
							StrBounds = Split(Replace(Replace(StrBounds,"Rectangle {",""),"}",""),",")
				 			iX = cint(StrBounds(2))
				 			iY = cint(StrBounds(1))
							iX = iX + 4
							iY = iY + 5
				End if
				Select Case arrOptionFlag(iCnt)
					Case "Check" 'For "=" or "=Any" condition
						obCPDWindow.JavaObject("VariantNatTable").Click iX,iY,"LEFT"
						Wait 1
					Case "None" 'For "!=" or or "=NONE" condition
						obCPDWindow.JavaObject("VariantNatTable").dblClick iX,iY,"LEFT"
						Wait 1
					Case "Blank" 'to set blank
					   ' For future use
					Case else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Invalid case [ "&arrOptionFlag(iCnt)&" ]" )
						Fn_CPD_PCA_VariantNatTable_VariantConditions_Operation = False
						Set obCPDWindow = Nothing
						Exit function
				End Select
			 Next	 
			 Fn_CPD_PCA_VariantNatTable_VariantConditions_Operation = True
			 sMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("4GD_Toolbar"),"SaveTheCurrentContents")  ' Save the Variant Condition
			 Call Fn_ToolBarOperation("Click",sMenu,"")
			 Call Fn_ReadyStatusSync(1)
		'========================================================================================================================================
		Case "VerifyOptionsInVariantCondition"	
			 If dicDetails("ColumnName") <> "" Then 'Get column index
				iColCount = obCPDWindow.JavaObject("VariantNatTable").Object.getColumnCount()
				For iCnt = 0 To iColCount - 1
					strColumn =  CStr(obCPDWindow.JavaObject("VariantNatTable").Object.getCellByPosition(iCnt,0).getlayer.getColumnHeaderlayer.getDatavalueByPosition(iCnt,0).tostring)
					If trim(dicDetails("ColumnName")) = trim(strColumn) Then
							iColIndex = iCnt + 1
							Exit for
					End If
				Next
			 Else
				iColIndex = 2							 
			 End If
			 arrOptions = Split(dicDetails("Options"),"~") ' Set Options as On or OFF
			 iRowsCount = obCPDWindow.JavaObject("VariantNatTable").Object.getRowCount()						 
			 For iCnt = 0 To UBound(arrOptions)
				bFlag = False
				For iCnt1 = 1 To iRowsCount - 1
					If instr(CStr(obCPDWindow.JavaObject("VariantNatTable").Object.getCellByPosition(0,iCnt1).getDataValue().getData().toString()),arrOptions(iCnt)) Then
						bFlag = True
						Exit for
					End If	
				Next	
				If bFlag = False Then
					Fn_CPD_PCA_VariantNatTable_VariantConditions_Operation = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Fail to verify Variant Option [ "&arrOptions(iCnt)&" ] in Variant Condition View.")
					Exit For
				Else
					Fn_CPD_PCA_VariantNatTable_VariantConditions_Operation = True
				End If	
			 Next
		'========================================================================================================================================
		Case "SplitColumnAndSetVarConditionAndSave"
			 If dicDetails("ColumnName") <> "" Then  'Get column index
				iColCount = obCPDWindow.JavaObject("VariantNatTable").Object.getColumnCount()
				For iCnt = 0 To iColCount - 1
					strColumn =  CStr(obCPDWindow.JavaObject("VariantNatTable").Object.getCellByPosition(iCnt,0).getlayer.getColumnHeaderlayer.getDatavalueByPosition(iCnt,0).tostring)
					If trim(dicDetails("ColumnName")) = trim(strColumn) Then
							iColIndex = iCnt + 1
							Exit for
					End If
				Next
			 Else
				iColIndex = 2							 
			 End If
			 '------------------- Split column ---------------------------------------------------------------------
			 StrBounds = obCPDWindow.JavaObject("VariantNatTable").Object.getBoundsByPosition(iColIndex,0).tostring
			 StrBounds = Split(Replace(Replace(StrBounds,"Rectangle {",""),"}",""),",")
			 iX = cint(StrBounds(2))
			 iY = cint(StrBounds(3) - 40)
			 If IsEmpty(iX) or IsEmpty(iY) Then
				Fn_CPD_PCA_VariantNatTable_VariantConditions_Operation = False
				Set obCPDWindow = Nothing
				Exit Function
			 End If				
			 iX = iX + 4
			 iY = iY + 5
			 obCPDWindow.JavaObject("VariantNatTable").Click iX, iY ,"RIGHT"
			 Wait 2
			 strPopupMenu = obCPDWindow.WinMenu("ContextMenu").BuildMenuPath(Popupmenu)
			 obCPDWindow.WinMenu("ContextMenu").Select strPopupMenu
			 Wait 3
			 '-----------------------------------------------------------------------
			 iColCount = obCPDWindow.JavaObject("VariantNatTable").Object.getColumnCount()
			 For iCnt = 0 To iColCount - 1
				strColumn =  CStr(obCPDWindow.JavaObject("VariantNatTable").Object.getCellByPosition(iCnt,0).getlayer.getColumnHeaderlayer.getDatavalueByPosition(iCnt,0).tostring)
				If instr(trim(strColumn) , trim("expressiongrid")) > 0 Then
					iColIndex = iCnt + 1
					Exit for
				End If
			 Next
			 ' Check for Row value to check
			 arrOptions = Split(dicDetails("Options"),"~") ' Set Options as On or OFF
			 arrOptionFlag = Split(dicDetails("OptionFlag"),"~")
			 iRowsCount = obCPDWindow.JavaObject("VariantNatTable").Object.getRowCount()						 
			 For iCnt = 0 To UBound(arrOptions)
		 		If instr(arrOptions(iCnt),"@") Then  'for multiple instance
		 			aOption = Split(arrOptions(iCnt),"@")
		 			arrOptions(iCnt) = aOption(0)
		 			iInstance = aOption(1)
		 		Else
					iInstance = 1				 		
		 		End If
		 		iCounter = 1	
				For iCnt1 = 1 To iRowsCount - 1
					If instr(CStr(obCPDWindow.JavaObject("VariantNatTable").Object.getCellByPosition(0,iCnt1).getDataValue().getData().toString()),arrOptions(iCnt)) Then
						iRowIndex = iCnt1
						If cint(iCounter) = cint(iInstance) Then
			 				Exit For
			 			Else
							iCounter = iCounter + 1							 			
			 			End If
					End If	
				Next
				
				StrBounds = obCPDWindow.JavaObject("VariantNatTable").Object.getBoundsByPosition(iColIndex,iRowIndex).tostring
				StrBounds = Split(Replace(Replace(StrBounds,"Rectangle {",""),"}",""),",")
				iX = cint(StrBounds(2))
				iY = cint(StrBounds(1))
				
				If IsEmpty(iX) or IsEmpty(iY) Then
					Fn_CPD_PCA_VariantNatTable_VariantConditions_Operation = False
					Set obCPDWindow = Nothing
					Exit Function
				End If
				iX = iX + 29
				Select Case arrOptionFlag(iCnt)
					Case "Check" 'For "=" or "=Any" condition
						obCPDWindow.JavaObject("VariantNatTable").Click iX,iY,"LEFT"
						Wait 1
					Case "None" 'For "!=" or or "=NONE" condition
						obCPDWindow.JavaObject("VariantNatTable").dblClick iX,iY,"LEFT"
						Wait 1
					Case "Blank" 'to set blank
					   ' For future use
					Case else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Invalid case [ "&arrOptionFlag(iCnt)&" ]" )
						Fn_CPD_PCA_VariantNatTable_VariantConditions_Operation = False
						Set obCPDWindow = Nothing
						Exit function
				End Select
			 Next	 
			 Fn_CPD_PCA_VariantNatTable_VariantConditions_Operation = True
			 sMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("4GD_Toolbar"),"SaveTheCurrentContents")  ' Save the Variant Condition
			 Call Fn_ToolBarOperation("Click",sMenu,"")
			 Call Fn_ReadyStatusSync(1)			 
		'================================================================================================================================================	
		Case "VerifyColumnInVariantConditionTab"
			 If dicDetails("ColumnName") <> "" Then  'Get column index
				arrOptions = Split(dicDetails("ColumnName"),"~")
				iColCount = obCPDWindow.JavaObject("VariantNatTable").Object.getColumnCount()
				For iCnt1 = 0 To UBound(arrOptions)
					bFlag = False
					For iCnt = 0 To iColCount - 1
						strColumn =  CStr(obCPDWindow.JavaObject("VariantNatTable").Object.getCellByPosition(iCnt,0).getlayer.getColumnHeaderlayer.getDatavalueByPosition(iCnt,0).tostring)
						If instr(trim(strColumn),trim(arrOptions(iCnt1))) > 0 Then
							bFlag = True	
							Exit for
						End If
					Next
					If bFlag = False Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed verify existence of column [ "&arrOptions(iCnt1)&" ]" )
						Fn_CPD_PCA_VariantNatTable_VariantConditions_Operation = False
						Set obCPDWindow = Nothing
						Exit function
					End If
				Next
				Fn_CPD_PCA_VariantNatTable_VariantConditions_Operation = True				
			 End If
	 '================================================================================================================================================			 
	End Select
	
	If StrTabName <> "" Then ' Minimize tab
		If Fn_CPD_CompnentTabOperations("IsMaximized",StrTabName, "") = True Then
			Call Fn_CPD_CompnentTabOperations("DoubleClick",StrTabName, "")
			Call Fn_ReadyStatusSync(1)
		End If
		If lcase(sTabClose) = "yes" Then ' Close Tab
			Call Fn_CPD_CompnentTabOperations("Close",StrTabName, "")
			Call Fn_ReadyStatusSync(1)
		End If
	End If
	
	Set obCPDWindow = Nothing
	
End Function
'========================================================================================================================================================================
'@@
'@@    Function Name		:	Fn_CPD_PCA_SetRuleDateInConfigurationView()
'@@
'@@    Description			:	Function to set Date Rule in configuration View
'@@
'@@    Parameters			:	1. sAction	: Action Name
'@@ 							2.sDate : Date to set
'@@ 							3.sTime : Time to set
'@@ 							4.sButton : Button name
'@@
'@@    Return Value		   	: 	True Or False
'@@
'@@    Pre-requisite		:	Configuration View Should be opened
'@@
'@@    Examples				:	bReturn = Fn_CPD_PCA_SetRuleDateInConfigurationView("SetRuleDate","","","No Date")
'@@ 														
'@@-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@	   History				:	Developer Name				Date	  			Rev. No.		Changes Done		Reviewer
'@@-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@	   Created By 			:	Poonam Chopade			27-Aug-2019				1.0		  		 Created		  [Tc12.3(2019081900)-27Aug2019-PoonamC-NewDevelopment]
'========================================================================================================================================================================
Public Function Fn_CPD_PCA_SetRuleDateInConfigurationView(sAction,sDate,sTime,sButton)

	GBL_FAILED_FUNCTION_NAME = "Fn_CPD_PCA_SetRuleDateInConfigurationView"
	Dim objSetRuleDate,objCPDWindow
	
	Fn_CPD_PCA_SetRuleDateInConfigurationView = False
	
	Set objCPDWindow = Fn_SISW_CPD_GetObject("Collaborative Product")
	Set objSetRuleDate = objCPDWindow.JavaWindow("Set Rule Date")
	
	Select Case sAction
		Case "SetRuleDate"
				'Check Existence of Date Rule set to system Default
				If Fn_UI_ObjectExist("Fn_CPD_PCA_SetRuleDateInConfigurationView",objCPDWindow.JavaObject("DateTimeImageHyperlink")) = True Then 
						Call Fn_UI_JavaObject_Click("Fn_CPD_PCA_SetRuleDateInConfigurationView",objCPDWindow,"DateTimeImageHyperlink",5,5,"LEFT")
						Call Fn_ReadyStatusSync(1)
						objCPDWindow.WinMenu("ContextMenu").Select "Set Rule Date"
						Wait(1)
						If sDate <> "" Then  'Check Date Text
							Call Fn_SISW_UI_JavaEdit_Operations("Fn_CPD_PCA_SetRuleDateInConfigurationView","Type",objSetRuleDate,"Date",sDate)
							Call Fn_ReadyStatusSync(1)
						End If
						If sTime <> "" Then  'Check Time Text
							Call Fn_SISW_UI_JavaList_Operations("Fn_CPD_PCA_SetRuleDateInConfigurationView","Select",objSetRuleDate,"Time",sTime,"","")
							Call Fn_ReadyStatusSync(1)
						End If
						If sButton <> "" Then  'Click on button OK / Cancel / No Date
							objSetRuleDate.JavaButton("Button").SetTOProperty "label",sButton
							Fn_CPD_PCA_SetRuleDateInConfigurationView = Fn_Button_Click("Fn_CPD_PCA_SetRuleDateInConfigurationView",objSetRuleDate,"Button")
							Call Fn_ReadyStatusSync(1)
						End If
				End If
	End Select
	
	Set objCPDWindow = Nothing
	Set objSetRuleDate = Nothing

End Function
'=========================================================================================================================================================================
'@@
'@@    Function Name		:	Fn_CPD_PCA_VariantConfiguration_Operation
'@@
'@@    Description			:	Function Used to Perform operation on Variant Configuration view
'@@
'@@    Parameters			:	1. sAction		: Action to be performed
'@@							:	2. StrTabName	: Tab Name
'@@							:	2. dicDetails	: Dictionary object
'@@							:	2. Popupmenu	: Popup menu name
'@@							:	3. sTabClose	: Flag to close Tab
'@@
'@@    Return Value		   	: 	True Or False
'@@
'@@    Pre-requisite		:	4GD Designer Perspective should be opened 
'@@
'@@    Examples				:	Set dicDetails = CreateObject( "Scripting.Dictionary")
'@@ 								dicDetails("Options") = "1200~1600~Arial~std"
'@@ 								dicDetails("OptionFlag") = "Check~Check~None~Check"
'@@									dicDetails("ToolBarButton") = "Expand"
'@@ 							bReturn = Fn_CPD_PCA_VariantConfiguration_Operation("SetVarOptionValue","Variant Configuration",dicDetails,"","Yes")
'@@ 								
'@@ 							Set dicDetails = CreateObject( "Scripting.Dictionary")
'@@ 								dicDetails("Options") = "1200~1600~Arial~std"
'@@									dicDetails("ToolBarButton") = "Validate~Expand"
'@@ 							bReturn = Fn_CPD_PCA_VariantConfiguration_Operation("ModifySetVarOptionValue","Variant Configuration",dicDetails,"","Yes")
'@@ 								
'@@ 							Set dicDetails = CreateObject( "Scripting.Dictionary")
'@@ 								dicDetails("Options") = "1200~1600~Arial~std"
'@@ 								dicDetails("OptionFlag") = "None~Check~None~Check"
'@@ 							bReturn = Fn_CPD_PCA_VariantConfiguration_Operation("VerifySetVarOptionValue","Variant Configuration",dicDetails,"","Yes")
'@@
'@@-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@	   History				:	Developer Name				Date	  			Rev. No.		Changes Done		Reviewer
'@@-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@	   Created By 			:	Poonam Chopade			27-Aug-2019				1.0		  		 Created		  [Tc12.3(2019081900)-27Aug2019-PoonamC-NewDevelopment]
'=========================================================================================================================================================================
Public Function Fn_CPD_PCA_VariantConfiguration_Operation(sAction,StrTabName,dicDetails,Popupmenu,sTabClose)

	GBL_FAILED_FUNCTION_NAME = "Fn_CPD_PCA_VariantConfiguration_Operation"
	
	Dim objCPDWindow,arrOptionFlag,arrOptions
	Dim sMenu,bFlag,iCnt,iCnt1,iRowIndex,iRowsCount,iColIndex,iColCount,iCnt2,sOption,strFlag
	Dim sAppMsg,iX,iY,iHeight,iItemHeight,aOption,iInstance,iCounter
	
	Fn_CPD_PCA_VariantConfiguration_Operation = False
	Set objCPDWindow = Fn_SISW_CPD_GetObject("Collaborative Product")
	
	On Error Resume Next 
	'Check Existence of Variant Configuration Tab & click on Toolbar button
	'If Fn_UI_ObjectExist("Fn_CPD_PCA_VariantConfiguration_Operation",objCPDWindow.JavaObject("DateTimeImageHyperlink")) = False And Fn_UI_ObjectExist("Fn_CPD_PCA_VariantConfiguration_Operation",objCPDWindow.JavaTable("VariantConfigurationTable")) = False Then
	If not  objCPDWindow.JavaObject("DateTimeImageHyperlink").Exist(5) and not objCPDWindow.JavaTable("VariantConfigurationTable").Exist(1) Then
		sMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("4GD_Toolbar"),"ViewVariantConfiguration")
		Call Fn_ToolBarOperation("Click",sMenu,"")
		Call Fn_ReadyStatusSync(1)
		
		If Fn_UI_ObjectExist("Fn_CPD_PCA_VariantConfiguration_Operation",objCPDWindow.JavaTable("VariantConfigurationTable")) = False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Fail to check existence of Variant Configuration Table" )
			Fn_CPD_PCA_VariantConfiguration_Operation = False
			Set objCPDWindow = Nothing
			Exit function
		End If
		
	End If
    If StrTabName <> "" Then ' Maximize tab
		If Fn_CPD_CompnentTabOperations("IsMaximized",StrTabName, "") = False Then
			Call Fn_CPD_CompnentTabOperations("DoubleClick",StrTabName, "")
			Call Fn_ReadyStatusSync(1)
		End If
	End If 
	
	Select Case sAction
		'==============================================================================================================================================	
		Case "SetVarOptionValue","ModifySetVarOptionValue","SetVarOptionValue_Ext"
			 If sAction= "SetVarOptionValue" or sAction= "ModifySetVarOptionValue" Then
			 	Call Fn_CPD_PCA_SetRuleDateInConfigurationView("SetRuleDate","","","No Date") 'Clear date rule
			    Call Fn_ReadyStatusSync(1)
			 End If 
			 
			 If sAction = "ModifySetVarOptionValue" Then ' clear set expression and then set new expression
					Call Fn_ToolBarOperation("ShowDropdownAndSelectWithInstance", "View Menu:1", "Clear Expression")
					Call Fn_ReadyStatusSync(1)
			 End If
			 arrOptions = Split(dicDetails("Options"),"~")  ' Set Options as On or OFF
			 arrOptionFlag = Split(dicDetails("OptionFlag"),"~")
			 iRowsCount = objCPDWindow.JavaTable("VariantConfigurationTable").GetROProperty("rows")
			 iColCount = objCPDWindow.JavaTable("VariantConfigurationTable").GetROProperty("cols")				 
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
				 	
					For iCnt1 = 0 To iRowsCount - 1  'Get Row & col index as per option name	
						bFlag = False
						For iCnt2 = 0 To iColCount - 1
							sOption = objCPDWindow.JavaTable("VariantConfigurationTable").GetCellData(iCnt1,iCnt2)
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
					Select Case arrOptionFlag(iCnt)
						Case "Check" 'For "=" or "=Any" condition
							objCPDWindow.JavaTable("VariantConfigurationTable").SelectCell iRowIndex,iColIndex
							Wait 1
						Case "None" 'For "!=" or or "=NONE" condition
							objCPDWindow.JavaTable("VariantConfigurationTable").SelectCell iRowIndex,iColIndex
							Wait 1
							objCPDWindow.JavaTable("VariantConfigurationTable").SelectCell iRowIndex,iColIndex
							Wait 1
						Case Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Fail to invalid case [ "&arrOptionFlag(iCnt)&" ]" )
							Fn_CPD_PCA_VariantConfiguration_Operation = False
							Set objCPDWindow = Nothing
							Exit function
					End Select	
					Fn_CPD_PCA_VariantConfiguration_Operation = True	
			 Next
		'==============================================================================================================================================			
		Case "VerifySetVarOptionValue","VerifySetVarOptionValue_Ext"
		         If sAction = "VerifySetVarOptionValue" Then
		         	 Call Fn_CPD_PCA_SetRuleDateInConfigurationView("SetRuleDate","","","No Date") 'Clear date rule
			         Call Fn_ReadyStatusSync(1)
		         End If
				 
				 arrOptions = Split(dicDetails("Options"),"~")  ' Set Options as On or OFF
				 arrOptionFlag = Split(dicDetails("OptionFlag"),"~")
				 iRowsCount = objCPDWindow.JavaTable("VariantConfigurationTable").GetROProperty("rows")
				 iColCount = objCPDWindow.JavaTable("VariantConfigurationTable").GetROProperty("cols")			 
				 For iCnt = 0 To UBound(arrOptions)  
						For iCnt1 = 0 To iRowsCount - 1   'Get Row & col index as per option name	
							bFlag = False
							For iCnt2 = 0 To iColCount - 1
								sOption = objCPDWindow.JavaTable("VariantConfigurationTable").GetCellData(iCnt1,iCnt2)
								If trim(arrOptions(iCnt)) = trim(sOption) Then
									bFlag = True
									Exit for
								End If
							Next
							If bFlag = True Then
								iRowIndex = iCnt1
								iColIndex = iCnt2
								Exit for
							End If	
						Next					
						If dicDetails("GridNumber") <> "" Then  ' if multiple grid coulmn present to check condition ex. in case of Overlay option		
							iColIndex = iColIndex - cint(dicDetails("GridNumber"))
						Else
							iColIndex = iColIndex - 1						 		
						End If
						Wait 1
						bFlag = False
						'Added below code to get Unique value of Image i.e Checked & Unchecked image data value
						strFlag = objCPDWindow.JavaTable("VariantConfigurationTable").Object.getitem(iRowIndex).getimage(iColIndex).getimageData.getAlpha(3,4) 
						
						Select Case arrOptionFlag(iCnt)
							Case "Check" 'For checked condition
									If CLng(strFlag) = 0 Then
										bFlag = True
									End If
							Case "None" 'For None condition
									If CLng(strFlag) > 0 Then
										bFlag = True
									End If
							Case "Blank" 'For blank condition
									If strFlag = Empty Then
										bFlag = True
									End If
						End Select
						If bFlag = False Then
							Fn_CPD_PCA_VariantConfiguration_Operation = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Fail to verify Variant condition [ "&arrOptions(iCnt)&" = "&arrOptionFlag(iCnt)&" ]" )
							Exit For
						Else
							Fn_CPD_PCA_VariantConfiguration_Operation = True
						End If
						strFlag = "" ' clear value from variable						 		
				 Next	
			'==============================================================================================================================================			
			Case "VerifyErrorMessage","GetErrorMessage"
				 iRowsCount = Fn_UI_Object_GetROProperty("Fn_CPD_PCA_VariantConfiguration_Operation",objCPDWindow.JavaTable("VariantConfigurationTable"),"rows")
 				 iColCount = Fn_UI_Object_GetROProperty("Fn_CPD_PCA_VariantConfiguration_Operation",objCPDWindow.JavaTable("VariantConfigurationTable"),"cols")				 
 				
				 For iCnt1 = 0 To iRowsCount - 1   'Get Row & col index as per option name	
					 bFlag = False
					 For iCnt2 = 0 To iColCount - 1
						sOption = objCPDWindow.JavaTable("VariantConfigurationTable").GetCellData(iCnt1,iCnt2)
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
				 
				 iHeight = objCPDWindow.GetROProperty("height")
				 iItemHeight = objCPDWindow.JavaTable("VariantConfigurationTable").Object.getItemHeight()
				 If iRowIndex <> 0  Then
				 	iItemHeight = iItemHeight * iRowIndex
				 End If
				 
				 iX = cint(cint(iHeight)+38)
				 iY = cint(320+cint(iItemHeight))
				 
				 objCPDWindow.JavaTable("VariantConfigurationTable").ActivateCell iRowIndex,iColIndex
				 Wait 2
				 objCPDWindow.Click iX+10,iY+10
				 Wait 1
				 'if Shell dialog not identified then
				 '--------------------------------------------------
				 If Fn_UI_ObjectExist("",objCPDWindow.JavaWindow("Shell"))= False Then
				 	 iY = iY + 30
					 iColIndex = iColIndex - 1
					 objCPDWindow.JavaTable("VariantConfigurationTable").ActivateCell iRowIndex,iColIndex
					 Wait 3
					 objCPDWindow.Click iX,iY
					 Wait 1
				 End IF
				'---------------------------------------------------
				 If Fn_UI_ObjectExist("",objCPDWindow.JavaWindow("Shell")) Then
			 		 sAppMsg = Fn_UI_Object_GetROProperty("Fn_CPD_PCA_VariantConfiguration_Operation",objCPDWindow.JavaWindow("Shell").JavaEdit("StyledText"),"text")
				 	 If sAction = "GetErrorMessage" Then
				 	 	Fn_CPD_PCA_VariantConfiguration_Operation = sAppMsg
				 	 Else
					 	 If instr(sAppMsg,dicDetails("Message")) > 0 Then
					 		Fn_CPD_PCA_VariantConfiguration_Operation = True
					 	 Else
					 		Fn_CPD_PCA_VariantConfiguration_Operation = False
					 		Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Actual Message ["+sAppMsg+"] does not match with expected message ["+dicDetails("Message")+"]")		
					 	 End IF
					 End If	 
				 Else
				 	Fn_CPD_PCA_VariantConfiguration_Operation = False
				 	Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Message window does not exists")		
				 End If
				'==============================================================================================================================================				
			 Case "VerifyEnabledConfigHeaderLabel"
				 For iCnt = 0 To 2
				 	objCPDWindow.JavaStaticText("ConfigHeaderLabel").SetTOProperty "index",iCnt
				 	If cstr(objCPDWindow.JavaStaticText("ConfigHeaderLabel").Object.isEnabled()) = "true" And cbool(objCPDWindow.JavaStaticText("ConfigHeaderLabel").Object.isEnabled()) = true Then
				 		If Instr(objCPDWindow.JavaStaticText("ConfigHeaderLabel").Object.getToolTipText(),dicDetails("ConfigHeadertooltip")) > 0 Then
				 			Fn_CPD_PCA_VariantConfiguration_Operation = True
				 		End If
				 	End If	
				 Next
		'==============================================================================================================================================					
			End Select
			
	 If dicDetails("ToolBarButton") <> "" Then ' Click toolbar button in Configuration view
			sOption = Split(dicDetails("ToolBarButton"),"~")
			For iCnt1 = 0 To UBound(sOption)
				Call Fn_ToolBarOperation("Click",sOption(iCnt1),"")
				Call Fn_ReadyStatusSync(1)
			Next
	 End If
	If StrTabName <> "" Then ' Minimize Tab
		If Fn_CPD_CompnentTabOperations("IsMaximized",StrTabName, "") = True Then
			Call Fn_CPD_CompnentTabOperations("DoubleClick",StrTabName, "")
			Call Fn_ReadyStatusSync(1)
		End If
		If lcase(sTabClose) = "yes" Then ' Close Tab
			Call Fn_CPD_CompnentTabOperations("Close",StrTabName, "")
			Call Fn_ReadyStatusSync(1)
		End If
	End If
	
	Set objCPDWindow = Nothing
End Function

'@@=====================================================================================================================================================================
'@@
'@@    Function Name		:	Fn_CPD_LoadVariantRule_Operations
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
'@@    Examples				:
'@@								Set dicLoadVarDetails = CreateObject( "Scripting.Dictionary")
'@@ 								dicLoadVarDetails("SearchCriteria") = "Name:SVR01~ID:0001~Description:TestDescription"
'@@ 							bReturn = Fn_CPD_LoadVariantRule_Operations("SearchVariantRules",dicLoadVarDetails)				
'@@								
'@@-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@	   History				:	Developer Name				Date	  			Rev. No.		Changes Done		
'@@-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@	   Created By 			:Neha Patil			      25-June 2021			1.0		  		 Created		  
'@@=======================================================================================================================================================================
Public Function Fn_CPD_LoadVariantRule_Operations(sAction,dicLoadVarDetails)

	GBL_FAILED_FUNCTION_NAME = "Fn_CPD_LoadVariantRule_Operations"
	Dim objLoadVarRuledialog,arrVarRules,iRowCnt,iCount,bFlag,iCount1
	Dim sVarRule,sAppMsg,arrSearchCriteria,aSearchValues,sCheckStatus
	
	Fn_CPD_LoadVariantRule_Operations = False
	Set objLoadVarRuledialog = Fn_SISW_CPD_GetObject("LoadVariantRule")
	
	'Check Existence of Load Variant Rule dialog
	If Fn_UI_ObjectExist("Fn_CPD_LoadVariantRule_Operations",objLoadVarRuledialog) = False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Fail to check existence of [ Load Variant Rule ] dialog." )
		Fn_PC_LoadVariantRule_Operations = False
		Set objLoadVarRuledialog = Nothing
		Exit Function
	End If
	
	Select Case sAction
		'==============================================================================================================================			
		Case "SearchVariantRules" 
				If dicLoadVarDetails("SearchCriteria") <> "" Then
					arrSearchCriteria = Split(dicLoadVarDetails("SearchCriteria"),"~")
					 For iCount = 0 To UBound(arrSearchCriteria)
							aSearchValues = Split(arrSearchCriteria(iCount),":")
							objLoadVarRuledialog.JavaEdit("SearchCriteria").SetTOProperty "attached text",aSearchValues(0)+":"
							bFlag = Fn_SISW_UI_JavaEdit_Operations("Fn_CPD_LoadVariantRule_Operations", "Set", objLoadVarRuledialog, "SearchCriteria",aSearchValues(1))
							If bFlag = False Then
					  			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Fail to check existence of [ Search field = "&aSearchValues(0)&" ].")
					  			Fn_CPD_LoadVariantRule_Operations = False
					  			Exit For
					  		End If
					 Next
					Fn_CPD_LoadVariantRule_Operations = Fn_Button_Click("Fn_CPD_LoadVariantRule_Operations",objLoadVarRuledialog,"Search")
					Call Fn_ReadyStatusSync(1)
			  End If
		'==============================================================================================================================				  	
	 Case "SelectVariantRules"
	 	  If dicLoadVarDetails("VariantRuleNames") <> "" Then
			  arrVarRules = Split(dicLoadVarDetails("VariantRuleNames"),"~")
			  iRowCnt = objLoadVarRuledialog.JavaTable("VariantRules").GetROProperty("rows")
			  For iCount = 0 To UBound(arrVarRules)
			  		bFlag = False
			  		For iCount1 = 0 To iRowCnt - 1
			  			sVarRule = Fn_UI_JavaTable_GetCellData("Fn_CPD_LoadVariantRule_Operations",objLoadVarRuledialog,"VariantRules",iCount1,"Name")
			  			If trim(sVarRule) = trim(arrVarRules(iCount)) Then
			  				bFlag = Fn_UI_JavaTable_ClickCell("Fn_CPD_LoadVariantRule_Operations",objLoadVarRuledialog,"VariantRules",iCount1,"Select")
			  				Wait 1
			  				Exit For
			  			End If
			  		Next
			  		If bFlag = False Then
			  			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Fail to select [ Variant Rule = "&arrVarRules(iCount)&" ].")
			  			Fn_CPD_LoadVariantRule_Operations = False
			  			Exit For
			  		Else
			  			Fn_CPD_LoadVariantRule_Operations = bFlag
			  		End If
			  Next  
		   End If	
	    '==============================================================================================================================	
		Case "CheckboxOperation"
			If dicLoadVarDetails("Appendonlyexpressions") <> "" Then
				bFlag = Fn_SISW_UI_JavaCheckBox_Operations("Fn_CPD_LoadVariantRule_Operations", "Set", objLoadVarRuledialog, "Append only expressions", dicLoadVarDetails("Appendonlyexpressions"))
			End If
			If dicLoadVarDetails("Loadexpandedexpression") <> "" Then
				bFlag = Fn_SISW_UI_JavaCheckBox_Operations("Fn_CPD_LoadVariantRule_Operations", "Set", objLoadVarRuledialog, "Load expanded expression", dicLoadVarDetails("Loadexpandedexpression"))
			End If
			Fn_CPD_LoadVariantRule_Operations = bFlag
		'==============================================================================================================================
	
	End Select
	
	If dicLoadVarDetails("Button") <> "" Then  'Click on Buttons
		 Call Fn_Button_Click("Fn_CPD_LoadVariantRule_Operations",objLoadVarRuledialog,dicLoadVarDetails("Button"))	
		 Call Fn_ReadyStatusSync(1)	
	End If 
	
	Set objLoadVarRuledialog = Nothing
	
End Function


'=========================================================================================================================================================================
'@@    Function Name		:	Fn_CPD_PCA_Save_Variant_Rule
'@@
'@@    Description			:	Function Used to Save the variant Rule.
'@@
'@@    Parameters			:	1. dicLoadVarDetails	: Dictionary object
'@@						
'@@
'@@    Return Value		   	: 	True Or False
'@@
'@@    Pre-requisite		:	Need to Create rules in Configuration rules.
'@@
'@@    Examples				:	Set dicLoadVarDetails = CreateObject( "Scripting.Dictionary")
'@@ 								dicLoadVarDetails("Name") = DataTable("SVR_Name", dtGlobalSheet)+"_"+iRanNo
'@@ 							bReturn = Fn_CPD_PCA_Save_Variant_Rule(dicDetails)
'@@ 								
'@@-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@	   History				:	Developer Name				Date	  			Rev. No.		Changes Done		Reviewer
'@@-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@	   Created By 			:	Radha Mane 		      25-June-2021				1.0		  		 Created		  
'=========================================================================================================================================================================
Public Function Fn_CPD_PCA_Save_Variant_Rule(dicLoadVarDetails)
	Dim ObjSaveVarWindow
	
  	Fn_CPD_PCA_Save_Variant_Rule = False 
    set ObjSaveVarWindow = JavaWindow("Collaborative Product").JavaWindow("Set Rule Date")
    ObjSaveVarWindow.SetTOProperty "title","Save Variant Rule"
    Err.clear
    If Fn_UI_ObjectExist("Fn_PC_VariantConfigurationView_Operation",ObjSaveVarWindow) Then
    	ObjSaveVarWindow.JavaEdit("Date").SetTOProperty "attached text","ID"
    	If dicLoadVarDetails("ID") <> "" Then
    		ObjSaveVarWindow.JavaEdit("Date").Type dicLoadVarDetails("ID")
    	End If
    	 
    	If ObjSaveVarWindow.JavaEdit("Date").Exist(1) Then
    		Call Fn_KeyBoardOperation("SendKey","{TAB}")
    	End If
    	
      	ObjSaveVarWindow.JavaEdit("Date").SetTOProperty "attached text","Name *"
		wait 2
        ObjSaveVarWindow.JavaEdit("Date").Type dicLoadVarDetails("Name")
        If Err.number < 0 Then
			Fn_CPD_PCA_Save_Variant_Rule = False 
        Else
            sAppMsg = Fn_UI_Object_GetROProperty("Fn_PC_VariantConfigurationView_Operation",ObjSaveVarWindow.JavaEdit("Date"),"text")
			ObjSaveVarWindow.JavaButton("Button").SetTOProperty "label","OK"
			wait 1
			ObjSaveVarWindow.JavaButton("Button").Click
			Fn_CPD_PCA_Save_Variant_Rule = sAppMsg
        End If
	 Else
	 	Fn_CPD_PCA_Save_Variant_Rule = False
	 	Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Save variant Rule Dialog  does not exists")		
	 End If	  			 
	Set ObjSaveVarWindow = Nothing
End Function
