
'********************************************************************************************
' Dictionary Definitions
'********************************************************************************************

' Defining a dictionary for Extended Scheduke Creation
Dim dicScheduleInfo
' Defining a dictionary for Schedule property
Dim dicScheduleProperty
' Defining a dictionary for Task property
Dim dicTaskProperty
' Defining a dictionary for Extended Task Creation
Dim dicTaskInfo
' Defining a dictionary for Specific Search Criteria Creation
Dim dicSearchCriteria
'Defining a dictionary for Object properties
Dim dicItemProperty
'Defining a dictionary for Schedule Calendar
Dim dicSchCalendar
'Defining a dictionary for Schedule Cost 
Dim dicSchCost
'Defining a dictionary for Select signoff 
Dim dicSelectSignOff
'Defining a dictionary for VariantRule
Dim dicVariantRule
'Defining a dictionary for Configure Operation
Dim dicConfiguration
'Defining a dictionary for Effectivity Mapping
Dim dicEffectivityMapping
' Defining a Dictionary for Process View Attributes for function 
Dim dicProcessViewAttributes
' Defining a Dictionary for Incremental Change 
Dim dicIncrementalChange
' Defining a Dictionary for Sub-Process Properties
Dim dicNewSubProcess
' Defining a Dictionary for Edit Properties
Dim dicItemPropDictonary
'Defining a Dictionary for Report Creation Wizard
Dim dicReportDesignerWizard
'Defining a dictionary for ViewerTab Dialog
Dim dicViewerTab
'Defining a dictionary for New Change Dialog
Dim dicNewChange
'Defining a dictionary for Technical Document
Dim dicTechDoc
'Defining a dictionary for Vendor and Other Item Types.
Dim dicItemInfo

Dim dicSearchCriteriaOne
' Defining Dictionay Object to Set Date / Unit / End Item / Intent 
Dim dicRevRuleInfo
'Defining Dictionay Object  to handle - Name of Class, Sys of measurement, Options while creating Class in Classification Admin
Dim dicClassOperations
'Defining Dictionay Object  to handle - Name of Group, AssignID while creating Group in Classification Admin
Dim dicGroupOperations
'Defining Dictionay Object  to handle - View Operations in Classification Admin
Dim dicViewOperations
'Process Assignment List
Dim dicNewCreateEditAssign
'Defining Dictionary object to handle - Business object creation.
Dim dicBusinessObj
'Defining Dictionary object to handle - Referencer tab in PSE
Dim dReferencersDict
'Defining Dictionary object to handle - New Option Dialog in PSE
Dim dictNewOption
'Defining Dictionary object For Web Search
Dim dicWebSearch
'Defining Dictionary object For Vendor
Dim dicVendor
'Defining Dictionary object For Bid Package
Dim dicBidPackage
'Defining Dictionary object For dicCommercialPart
Dim dicCommercialPart
'Defining Dictionary object For Company Contact
Dim dicCompanyContact
'Defining Dictionary object For Company Location
Dim dicCompanyLocation
'Defining Dictionary object For Vendor Part
Dim dicVendorPart
'Defining Dictionary object For Item Details Create
Dim dicItemDetailsCreate
'Defining Dictionary object For Overview Tab Contents
Dim dicOverviewTabContent
'Defining Dictionary object For Object Properties
Dim dicProperties
'Defining Dictionary object For ID Display Rules
Dim dicIDDisplayRules
'Defining Dictonary object for Audit Log.
Dim dicAuditLog
'Defining Dictonary object forItem From Template
Dim dicItemFromTemplate
'Defining Dictonary object for Save As Object
Dim dicSaveAsObject
'Defining Dictonary object for Import to Teamceter functionality
Dim dicImportTeamcenter
'Declaring Dictionary for Search Preference
Dim dicSearchPreferences
'Declaring Dictionary for Lot Operations
Dim dicLotOperations
'Declaring Dictionary for Details Dataset Creation
Dim dicDatasetInfo
'Declaring Dictionary for Cacheless Search. ( RDV )
Dim dicItemIDSearch
'Declaring Dictionary for New IRDC creation ( BMIDE )
Dim dicIRDC
'Declaring Dictionary for New Functionality creation ( BMIDE )
Dim dicFunctionality
'Declaring Dictionary for Id Display Rule Information
Dim dicIdDisplayRulesInfo
'Declaring Dictionary for Deep Copy Rule Information ( BMIDE )
Dim dicDeepCopyRuleInfo
'Declaring Dictionary for Status Indicators Image file
Dim dicStatusIndicator
'Declaring Dictionary for Component Information
Dim dicComponentInfo
'Declaring Dictionary for Replace Component Information
Dim dicReplaceComponentInfo
'Declaring Dictionary for Content Search in CPD perspective
Dim dicContentSearch
Dim dicBOInfo
'Declaring Dictionary for Web Project Assign Information
Dim dicProjectInfo
'Declaring Dictionary for Partition Information in CPD perspective
Dim dicPartitionInfo
'Declaring Dictionary for BMIDE
Dim dicOpsInputPropInfo
'Declaring Dictionary for Web Project Remove Information
Dim dicRemoveProjectInfo
'Declaring Dictionary for Web Business Object Information
Dim dicWebBOInfo
'Declaring Dictionary for Bat file Details
Dim dicBatFileDetails
Dim dicWebCompanyInfo
'Declaring Dictionary for Parameter Definition Group Information
Dim dicParameterDefinitionGroupInfo
Dim dicDuplicateInfo
' Added for function Fn_SISW_GetHierarchy (Setup.vbs) used for frequently changing menus, tree items etc.
Dim dicGetHierarchy
Dim dicLoadVarDetails

'Setting up the object for Dictionary
Set dicScheduleInfo = CreateObject( "Scripting.Dictionary" )
Set dicScheduleProperty = CreateObject( "Scripting.Dictionary" )
Set dicTaskProperty = CreateObject( "Scripting.Dictionary" )
Set dicTaskInfo = CreateObject( "Scripting.Dictionary" )
Set dicSearchCriteria = CreateObject( "Scripting.Dictionary" )
Set dicItemProperty = CreateObject("Scripting.Dictionary")
Set dicSchCalendar = CreateObject("Scripting.Dictionary")
Set dicSchCost = CreateObject("Scripting.Dictionary")
Set dicSelectSignOff = CreateObject( "Scripting.Dictionary")
Set dicVariantRule = CreateObject( "Scripting.Dictionary")
Set dicConfiguration = CreateObject( "Scripting.Dictionary")
Set dicEffectivityMapping =  CreateObject( "Scripting.Dictionary")
Set dicProcessViewAttributes = CreateObject("Scripting.Dictionary")
Set dicIncrementalChange  = CreateObject("Scripting.Dictionary")
Set dicNewSubProcess = CreateObject("Scripting.Dictionary")
Set dicItemPropDictonary = CreateObject("Scripting.Dictionary")
Set dicReportDesignerWizard = CreateObject("Scripting.Dictionary")
Set dicViewerTab = CreateObject("Scripting.Dictionary")
Set dicNewChange = CreateObject("Scripting.Dictionary")
Set dicSearchCriteriaOne = CreateObject("Scripting.Dictionary")
Set dicRevRuleInfo = CreateObject("Scripting.Dictionary")
Set dicClassOperations=CreateObject("Scripting.Dictionary")
Set dicGroupOperations=CreateObject("Scripting.Dictionary")
Set dicViewOperations = CreateObject("Scripting.Dictionary")
Set dicNewCreateEditAssign = CreateObject("Scripting.Dictionary")
Set dicTechDoc = CreateObject("Scripting.Dictionary")
Set dicItemInfo = CreateObject("Scripting.Dictionary")
Set dicBusinessObj = CreateObject("Scripting.Dictionary")
Set dReferencersDict = CreateObject("Scripting.Dictionary")
Set dictNewOption = CreateObject("Scripting.Dictionary")
Set dicWebSearch=CreateObject("Scripting.Dictionary")
Set dicVendor=CreateObject("Scripting.Dictionary")
Set dicBidPackage = CreateObject( "Scripting.Dictionary" )
Set dicCommercialPart=CreateObject("Scripting.Dictionary")
Set dicCompanyContact = CreateObject( "Scripting.Dictionary")
Set dicCompanyLocation = CreateObject( "Scripting.Dictionary")
Set dicVendorPart=CreateObject("Scripting.Dictionary")
Set dicItemDetailsCreate = CreateObject("Scripting.Dictionary")
Set dicOverviewTabContent = CreateObject("Scripting.Dictionary")
Set dicProperties = CreateObject("Scripting.Dictionary")
Set dicIDDisplayRules = CreateObject( "Scripting.Dictionary" )
Set dicAuditLog = CreateObject( "Scripting.Dictionary" )
Set dicItemFromTemplate = CreateObject( "Scripting.Dictionary" )
Set dicSaveAsObject = CreateObject( "Scripting.Dictionary" )
Set dicImportTeamcenter = CreateObject( "Scripting.Dictionary" )
Set dicSearchPreferences = CreateObject( "Scripting.Dictionary" )
Set dicLotOperations  = CreateObject( "Scripting.Dictionary" )
Set dicDatasetInfo = CreateObject( "Scripting.Dictionary" )
Set dicItemIDSearch = CreateObject( "Scripting.Dictionary" )
Set dicIRDC =CreateObject( "Scripting.Dictionary")
Set dicFunctionality = CreateObject( "Scripting.Dictionary")
Set dicIdDisplayRulesInfo = CreateObject("Scripting.Dictionary")
Set dicDeepCopyRuleInfo = CreateObject("Scripting.Dictionary")
Set dicStatusIndicator = CreateObject( "Scripting.Dictionary" )
Set dicComponentInfo = CreateObject( "Scripting.Dictionary" )
Set dicReplaceComponentInfo = CreateObject( "Scripting.Dictionary" )
Set dicContentSearch = CreateObject( "Scripting.Dictionary" )
Set dicBOInfo = CreateObject( "Scripting.Dictionary" )
Set dicProjectInfo = CreateObject( "Scripting.Dictionary" )
Set dicPartitionInfo = CreateObject("Scripting.Dictionary")
Set dicOpsInputPropInfo=CreateObject("Scripting.Dictionary")
Set dicRemoveProjectInfo = CreateObject("Scripting.Dictionary")
Set dicWebBOInfo=CreateObject("Scripting.Dictionary")
Set dicBatFileDetails = CreateObject("Scripting.Dictionary")
Set dicWebCompanyInfo=CreateObject("Scripting.Dictionary")
Set dicParameterDefinitionGroupInfo = CreateObject( "Scripting.Dictionary")
Set dicDuplicateInfo=CreateObject("Scripting.Dictionary")
Set dicGetHierarchy=CreateObject("Scripting.Dictionary")
Set dicLoadVarDetails = CreateObject("Scripting.Dictionary")

'Declaration of the object as structure.
With dicScheduleInfo  
			 .Add "TemplateName",""	
			.Add "ShiftDate", ""	                		 'Formate for  start date15-May-2010 18:30"		  
			.Add "Customer", ""
			.Add "CustomerNumber", ""     
			.Add "TimeZone", ""
			.Add "StartDate", ""							'Formate for  start date15-May-2010 18:30"		  
			.Add "FinishDate", ""						'Formate for  start date15-May-2010 18:30"		  
            .Add "ScheduleTemplate", ""		'	Pass boolean value
			.Add "SchedulePublic", ""			 '	Pass boolean value
			.Add "PercentLinked", ""				'	Pass boolean value
			.Add "Published", ""						'	Pass boolean value
			.Add "NotificationEnabled", ""		'   Pass boolean value
			.Add "FinishDateScheduling", ""  '  Pass boolean value   
			.Add "OwningProject", ""
			.Add "ProjectsSelection", ""			'This value should be : seprated for  multiple selection. and if we need to select all the value of list then set value as ALL
            .Add "DefineFormat",""             'this value should be passed as "dicScheduleInfo("DefineFormat")=True" to click on the define format button
End with

'Declaration of the object  for schedule property
With dicScheduleProperty         
			.Add "Name", ""
            .Add "Description", ""     
            .Add "StartDate", ""							  'Formate for  start date15-May-2010 18:30"
			.Add "FinishDate", ""							'Formate for finish date15-May-2010 18:30
			.Add "CustomerName", ""
			.Add "CustomerNumber", ""
			.Add "StatusDropDown", ""            'This is to modify the value of status
			.Add "TskStatusDrpDwn", ""			'This is to modify the value of  Task status
			.Add "PriorityDropDown", ""			'This is to modify the value of  Priority
			.Add "SchPriority", ""						'This is to verify the value of priority
			.Add "SchStatus", ""				'This is to verify  the value of  schedule status
			.Add "State", ""                     'This Is to verify the state of schedule
			.Add "TaskStatus", ""					 'This is to verify the value of  Task status
			.Add "IsTemplate", ""					'	Pass boolean value
			.Add "IsPercentLinked", ""				'Pass boolean value
			.Add "IsPublic", ""							 	'Pass boolean value
			.Add "Published", ""						'Pass boolean value
			.Add "NotificationsEnabled", ""		'Pass boolean value
			.Add "FinishDateSchedul", ""
			.Add "ScheduleMembers", ""		'Multiple schedule member should be pass , (Comma) separated.
			.Add "ScheduleDeliverable",""	 'this value should be passed as "dicScheduleProperty("ScheduleDeliverable")=True" to click on the ScheduleDeliverable button
			.Add "TaskDeliverables", ""	
			.Add "WorkflowTrigger", ""		
			.Add "WorkflowTaskTemplate", ""	
End with

'Declaration of the object  for  Task  Property
With dicTaskProperty         
			.Add "Name", ""
            .Add "Description", ""     
            .Add "StartDate", ""							  'Formate for  start date15-May-2010 18:30"
			.Add "ActualStartDate", ""				
			.Add "FinishDate", ""							'Formate for finish date15-May-2010 18:30
			.Add "ActualFinishDate", ""
			.Add "Duration", ""            
			.Add "WorkEstimate", ""			
			.Add "WorkComplete", ""			
			.Add "WorkCompletePercent", ""						
			.Add "Constraint", ""						
			.Add "Priority", ""	
			.Add "State",""
			.Add "Status", ""				
			.Add "TaskType", ""				
			.Add "FixedType", ""							 	
			.Add "AutoComplete", ""						'Pass boolean value
			.Add "WorkflowTrigger", ""		
			.Add "WorkflowTaskTemplate", ""	
			.Add "ResourceAssignments", ""			'Multiple  resource assignment  should be pass , (Comma) separated.
			.Add "TaskDeliverable",""
End with

'Declaration of the object as Task Info.
With dicTaskInfo  
			 .Add "CreatePhaseGate",""            'Pass boolean value
			 .Add "AdminTask",""            		  'Pass boolean value	
			 .Add "Category",""            				'String value
			 .Add "Complexity",""            		   'String value
			 .Add "ImpactAssReq",""            		'Pass boolean value	
			 .Add "ProposedTask",""            		'Pass boolean value
			.Add "StartDate", ""	                		 'Formate for  start date15-May-2010 18:30"		  
			.Add "FinishDate", ""							'Formate for  start date15-May-2010 18:30"		  
            .Add "FixedType", ""     
			.Add "CreateWorkflowTask", ""      'Pass boolean value
			.Add "WorkTrigger", ""							
			.Add "WorkFlowTaskTemplate", ""		
			.Add "AssignMember", ""	                 'Pass fully path which is : seprated.If multiple user need to pass seprate it with , (Comma) 
			.Add "Deliverable", ""	                 
End with

'Declaration of the object as structure.
With dicSearchCriteria  
'Java Edit Box
			.Add "SrchPersonName", ""
			.Add "AllSeqName", ""
			.Add "AllSeqOwningUser", ""
			.Add "CurrentTask",""	
			.Add "DatasetID", ""	                
			.Add "Description", ""	 
			.Add "ItmItemID", ""	 
			.Add "ItmName", ""	 
			.Add "ItmOwningOrgID", ""	 
			.Add "ItmRevAliasID", ""	 
            .Add "ItmRevAliasIDCntxtNm", ""	
			.Add "ItmRevAltID", ""	 
			.Add "ItmRevAltRev", ""
			.Add "ItmRevAltRevIDCntxtNm", ""	 
			.Add "ItmRevCurrentTask", ""	 		
			.Add "ItmRevDesc", ""  
			.Add "ItmRevItemID", ""
			.Add "ItmRevName", ""
			.Add "ItmRevRelStatus", ""
			.Add "ItmRevRevision", ""
			.Add "Keywords", ""
			.Add "Name", ""
			.Add "ReleaseStatus", ""
			.Add "Revision", ""
			.Add "PersonNm", ""
			.Add "ObjectID", ""
			.Add "ProjectID", ""
			.Add "UserData1", ""
			.Add "UserData2", ""
			.Add "ExcludeStatus", ""
			.Add "Partition ID:" , ""
			.Add "Partition Name:" , ""
			.Add "Model Name:" , ""
			.Add "Design Element Name:" , ""
			.Add "Design Element ID:" , ""
			.Add "Public ID:" , ""
            .Add "DocumentTitle",""
			.Add "IssuingAuthority",""
			.Add "ID",""
            .Add "License ID:",""
            .Add "Dataset Name:",""
			.Add "ProjectUser",""
			.Add "Project Name",""
'Java Check Box ->> Date Button
			.Add "CreatedAfterDt",""	
			.Add "CreatedBeforeDt",""	
			.Add "ModifiedAfterDt",""	
			.Add "ModifiedBeforeDt",""	
			.Add "ReleasedAfterDt",""	
			.Add "ReleasedBeforeDt",""	
			.Add "DateCreatedDt",""
' Java Button(DropDown)
			.Add "DsOwnGrpDrpDwn",""	
			.Add "DsOwnUsrDrpDwn",""	
			.Add "DsTypDrpDwn", ""	
			.Add "ItmRevAliasTypDrpDwn", ""	
			.Add "ItmRevAltRevTypDrpDwn", ""
			.Add "ItmRevOwningGrpDrpDwn",""	
			.Add "ItmRevOwningUsrDrpDwn",""	
			.Add "ItmRevTypDrpDwn", ""	
			.Add "OwningGrpDrpDwn", ""	
			.Add "OwningUsrDrpDwn", ""
			.Add "UsrID", ""
			.Add "Requestor", ""
			.Add "OwnUsrNameDrpDwn", ""
			.Add "Apply Class Name:", ""
			'Content management
			.Add "Name:", ""
			.Add "Is This A Template:", ""
			.Add "Content Version",""
			.Add "StyleSheetType",""
			.Add "StyleSheetResourceContentType",""
End with


'Declaration of the object property.
With dicItemProperty  
			 .Add "Description",""	
			.Add "Name", ""	                		 
			.Add "Version Limit", ""
			.Add "GovClassification", ""     
			.Add "Contract Pricing Model", ""  
			.Add "Note Text", ""  
End with


'Declaration of the object as structure.
With dicSchCalendar  
			 .Add "BaseCalendar", ""
			 .Add "TimeZone", ""
			 .Add "ExceptionDate", ""
			 .Add "WorkingOption", ""
			 .Add "OnWeekDays", ""					'Provide ~ separated Day Names
			 .Add "OffWeekDays", ""					'Provide ~ separated Day Names
			 .Add "DayHrDetails", ""				'Provide ~ separated values pairs for From-To. For multiple rows use , as separator
			 .Add "WorkingHrs",""                   'Provide ~ separated Day working hours"
End with

'Declaration of  the Schedule Cost 
With dicSchCost
			 .Add "TotalEstimatedCost", ""
			 .Add "TotalAccruedCost", ""
			 .Add "TotalEstimatedWork", ""
			 .Add "TotalAccruedWord", ""
			 .Add "BillCode", ""					
			 .Add "BillSub-code", ""					
			 .Add "BillType", ""	
			 .Add "RateModifier", ""
			 .Add "Rollup", ""                             'Set boolean value 
			 .Add "DrillDown", ""                       'Set boolean value
			 .Add "Name", ""                              
			 .Add "EstimatedHours", ""
			 .Add "AccruedHours", ""
			 .Add "EstimatedCost", ""
			 .Add "AccruedCost", ""
			 .Add "BD Name",""
			 .Add "BD EstimatedHours",""	
			 .Add "BD AccruedHours",""	
			 .Add "BD EstimatedCost",""	
			 .Add "BD AccruedCost",""
			 .Add "FC CostName",""
			 .Add "FC EstimatedCost",""
			 .Add "FC AccuredCost",""
			 .Add "CancelButtonClick",""

End with

'Declaration of  the Select signoff 
With dicSelectSignOff
		.Add "WorkListTreeNode",""
		.Add "SignOffTeamSelect",""         ' To be use in case of Profiles Case
		.Add "UsersName",""
		.Add "ProjectName",""
		.Add "ProjectUsers",""
		.Add "MemberOption",""
		.Add "GroupOption",""
		.Add "Quorum",""     ' Combination of QuorumType:Quorumvalue		
		.Add "Wait",""		
		.Add "ProcessDescription",""
		.Add "Comments",""
		.Add "Adhoc",""
		.Add "Action",""
End with


'Declaration of Variant Rule
' Use '~' as a separator for multiple values
With dicVariantRule
		.Add "Item",""
		.Add "Option",""         
		.Add "Description",""
		.Add "Value",""
		.Add "State",""
		.Add "SaveName",""
		.Add "SaveDesc",""
		.Add "SaveRelationType",""    
		.Add "ColName",""
End with

'Declaration for Configuration Operation 
With dicConfiguration
		.Add "sBOMLineNode",""
		.Add "sName",""
		.Add "sSavedName",""		
		.Add "sDescription",""         
		.Add "sConfiguration",""
		.Add "sInModule",""
		.Add "sItem",""             'Multiple value seprated by ~
	    .Add "sOption",""		'Multiple value seprated by ~
		.Add "sTabDesc",""
		.Add "sValue",""
		.Add "Message",""
End with


'Declaration for Effectivity Mapping Operation 
With dicEffectivityMapping
		 .Add "sColName","" 'Multiple value seprated by ~
		 .Add "sValue","" 'Multiple value seprated by ~
		 .Add "aRows","" 'Multiple value seprated by ~
		 .Add "bPackEffectivities",""    ' -  True / False
		 .Add "bUsedSharedEffectivity",""    ' -  True / False
		 .Add "bCreateNew",""    ' -  True / False
		 .Add "sEffectivityId",""
         .Add "bEffectivityProtection",""   ' -  True / False
		 .Add "sEndItem",""
		 .Add "sEndItemSelectType",""
		 .Add "sEndItemMRU",""
		 .Add "sEndItemName",""
		 .Add "sEndItemRev",""
		 .Add "bUnit",""    ' -  True / False
		 .Add "sUnit",""
		 .Add "bDate",""    ' -  True / False
		 .Add "sStartDates",""
		 .Add "sEndDates",""
		 .Add "sSubEffectivityEndItem",""
		 .Add "sSubEffectivityEndItemSelectType",""
		 .Add "sSubEffectivityEndItemMRU",""
		 .Add "sSubEffectivityEndItemName",""
		 .Add "sSubEffectivityUnit",""
		 .Add "sSubEffectivityDate",""
		 .Add "bUseLastReleaseDate",""    ' -  True / False
		
End with

' Declaration for Process View Attributes
With dicProcessViewAttributes
	.Add "WorkListTreeNode", ""           ' Mandatory field
	.Add "ProcessTree", ""
	.Add "Attributes", "" 
	.Add "State", ""
	.Add "ResParty", ""
	.Add "NameACL", ""
	.Add "SignOffsQuorum", ""
	.Add "DueDate", ""
	.Add "Duration", ""
	.Add "ReleaseStatus", ""
End With



' Declaration for Incremental Change
With dicIncrementalChange  
	.Add "sICId", "" 
	.Add "sICRevID", "" 
	.Add "sICName", ""
	.Add "sICDesc", ""
	.Add "sICType", "" 
	.Add "sICSelectionType", "" 
	.Add "sCol", "" 
	.Add "sValue", "" 
End With

' Declaration for Worklist Sub-Process Dialog
With dicNewSubProcess
	.Add "ProcessName", ""
	.Add "Description", ""
	.Add "ProcessTemplate", ""
	.Add "InheritTargets", ""
	.Add "Attachments", ""
End With

' Declaration for Edit Properties Dialog
With dicItemPropDictonary
	.Add "IsFastTrack", ""
	.Add "RecurringCost", ""
End With

' Declaration for Report Creation Wiard
With dicReportDesignerWizard
	.Add "JavaList:ReportDesign", ""
	.Add "JavaButton:Next-1", False
	.Add "JavaEdit:ObjectId", ""
	.Add "JavaEdit:ObjectName", ""
	.Add "JavaEdit:ObjectRevision", ""
	.Add "JavaButton:Next-2", False
	.Add "JavaList:ReportFormat", ""
	.Add "JavaButton:Finish", False
End With

'Declaration of ViewerTab Dialog
With dicViewerTab
		.Add "ReOccurence",""
		.Add "Offset",""         
		.Add "EventName",""
		.Add "EventName_EditBox",""
		.Add "RelativeTo",""
		.Add "ButtonName",""
		.Add "UIColumns",""
		.Add "ColsData",""
		.Add "StartDate",""
		.Add "EndDate",""
		.Add "RecurrenceEndDate",""
		.Add "Recurrence",""
		.Add "Save",""
End with

'Declaration of New Change Dialog (Also used in Web for function Fn_Web_CreateChange.)
With dicNewChange
	.Add"Menu",""
	.Add"Action",""
	.Add"Type",""
	.Add"ChangeID",""
	.Add"NodeName",""
	.Add"Filter",""
	.Add"ECRNo",""
	.Add"Revision",""
	.Add"Synopsis",""
	.Add"Desc",""
	.Add"ChangeType",""
	.Add"ButtonName",""
End With

With dicSearchCriteriaOne
.Add"Check-OutAndEdit",""
.Add"Cancel",""
.Add"Close",""
.Add"Cancel Check-Out",""
.Add"SaveAndCheck-In",""
.Add"Save",""
.Add"ItemComment",""
.Add"PreviousID",""
.Add"ProjectID",""
.Add"UserData1",""
.Add"UserData2",""
.Add"UserData3",""
End With

With dicRevRuleInfo
	.Add"sEffectivityDate","" 
	.Add"bUseToday",""
	.Add"sUnitNumber","" 
	.Add"sEndItem", ""
	.Add"sEndItemBy",""
	.Add"bAnyIntent", ""
	.Add"sIntentName", ""
	.Add"sIntentDesc", ""
	.Add"sAddIntentBy", ""
	.Add"sRemoveIntents","" ' - for future use
	.Add"sEffectivityGrpAction","" '
	.Add"sEffectivityGrpEntry","" '
	.Add"sEffectivityGrpSearchBy","" '
	.Add"sEffectivityGrp","" '
	.Add"sEffectivityGrpRev","" '
End With

'Used in ClassAdmin.vbs to Create Class
With dicClassOperations
		.Add "NodeName" , ""   'Node in the Hierarchy Tree
		.Add"AssignID",""            'Id to be assign to class if blank then Click on Assign button
		.Add"ClassName",""      ' Name of class
		.Add"SysMeasurement","" '
		.Add"Options_Abstract",""
		.Add"Options_AllowsMultipleInstances",""
		.Add"Options_Assembly",""
		.Add"Options_PreventRemoteICOCreation",""
		.Add"SaveCurrentInstance",""
		.Add"ChkProperties_Check",""
		.Add"ChkProperties_UnCheck",""
		.Add"Image",""
End With

'Used in ClassAdmin.vbs to Create Group
With dicGroupOperations
		.Add "NodeName" , ""
		.Add"AssignID",""
		.Add"GroupName",""
		.Add"Check_ICOCreation",""
		.Add"SaveCurrentInstance",""		
End With

'Used in ClassAdmin.vbs to Create View
With dicViewOperations
	.Add "sViewType", ""
	.Add "sViewID", ""
	.Add "sViewName", ""
	.Add "bViewDetails", ""
	.Add "sUser1", ""
	.Add "sUser2", ""
	.Add "bViewAttributes", ""
	.Add "sFieldValues", ""
	.Add "sViewAttributes", ""
	.Add "sAddImageUrl", ""
	.Add "bRemoveImage", ""
	.Add "bActivateOnLastEntry", ""
	.Add "sRemoveList", ""
	.Add "sAddToRightList", ""
	.Add "bClassAttributes", ""
	.Add "sClassAttributesSelect", ""
End With
'Used in function Fn_MyTc_CreateEditAssignmentList
With dicNewCreateEditAssign
	.Add "Description", ""
	.Add "ResourceTree", ""
	.Add "OrganizationTree", ""
	.Add "Member", ""
	.Add "Group", ""
	.Add "Action", ""
	.Add "ReviewQuorum", ""
	.Add "WaitForUndecidedReviewers", ""
	.Add "ResourceAdd", ""
	.Add "Projects", ""
	.Add "ProjectTeamTree", ""
	.Add "Process", ""
End With 
'Used in ADS.vbs to create ADS Technical Document.
With dicTechDoc
	.Add"TechType","" 
	.Add"TechIDPattern",""
	.Add"TechRevID","" 
	.Add"TechName", ""
	.Add"Categories",""
	.Add"SourceDocID",""
	.Add"CategoryList",""
	.Add"SrcDocCategories",""
	.Add"SrcTecDocCategories",""
	.Add"TechDocList", ""
	.Add"ButtonName", ""
End With
'Used in MyTeamcenter.vbs to create vendor and other Item types.
 With dicItemInfo
	.Add"ItemType","" 
	.Add"ConfigItem",""
	.Add"ItemID","" 
	.Add"ItemRev", ""
	.Add"ItemName",""
	.Add"ItemDesc",""
	.Add"ItemUOM",""
	.Add"VendorPartNo",""
	.Add"VendorPartName",""
	.Add"VendorID",""
	.Add"VendorName",""
	.Add"Title","" 
	.Add"FirstName",""
	.Add"LastName","" 
	.Add"Suffix", ""
	.Add"Email",""
	.Add"Fax",""
	.Add"Pager",""
	.Add"PhoneBusiness",""
	.Add"PhoneHome",""
	.Add"PhoneMobile",""
	.Add"Buttons",""
	.Add"ItemAddInfo",""
	.Add"OrgCageCode",""
End With
'Used in MyTeamcenter.vbs to create New Business Objects.
 With dicBusinessObj
	.Add"Name","" 
	.Add"LocationCode",""
	.Add"LocationType","" 
	.Add"Street", ""
	.Add"City",""
	.Add"State",""
	.Add"PostalCode",""
	.Add"Country",""
	.Add"Region",""
	.Add"URL",""
	.Add"Description",""
	.Add"Buttons",""
	.Add"CAGECode",""
	.Add"BriefcaseBrowserLicense",""
	.Add"DesignDataExchangeLicense",""
	.Add"FirstName",""
	.Add"LastName","" 
	.Add"CompContactForVendor",""
End With

'Used in StructureMananger.vbs to perform Referencers tab operations.
with dReferencersDict
	.Add"sWhere","" 
	.Add"sRule",""
	.Add"sDisplay","" 
	.Add"sItemType", ""
	.Add"sAt",""
	.Add"bIncludeSubtype",""
	.Add"sDepth",""
	.Add"sItem",""
	.Add"sType",""
	.Add"sRelation",""
end with

'Used in StructureMananger.vbs to perform operations on New Option dialog.
with dictNewOption
	.Add"sCreationType" , ""  ' "AnyWordOrPhrase",  "OneWordOrPhraseFromAFixedList", "AnyNumber", "NumberWithRestrictionsOnItsValue",  "LogicalValue",  "SameTypeAndRestrictions"
	.Add"sVisibility" , "" ' "Public", "Private"
	.Add"sName" , ""
	.Add"sDescription" , "" 
	.Add"sDerivedOptionType" , "" ' "Search for existing options", "Select from global options"
	.Add"sDefault" , ""
	.Add"bRestartWizard" , "" ' True , False
	.Add"sModuleName" , ""
	.Add"sOptionName" , ""
	.Add"bUseAsAnExternal" , "" ' True , False
end with

'Used in Web.vbs to perform Search Operation
'Used In Function Fn_Web_SearchOperation
With dicWebSearch
	.Add "EditBox:Name",""
	.Add "EditBox:Description",""
	.Add "EditBox:Owning User",""
	.Add "EditBox:Owning Group",""
	.Add "EditBox:Created After",""
	.Add "EditBox:Created Before",""
	.Add "EditBox:Item ID",""
	.Add "EditBox:ECR No",""
	.Add "EditBox:PR No",""
	.Add "EditBox:ECN No",""
	.Add "EditBox:Analyst",""
	.Add "Button:Type",""
	.Add "Button:Maturity",""
    .Add "EditBox:Requestor",""
	.Add "EditBox:Change Specialist I",""
	.Add "Button:Closure",""
End With

'Declaration of the object as Vendor
With dicVendor  
			 .Add "ID","" 
			.Add "Revision", ""            		 
			.Add "Name", ""    'Mandetory Parameter
			.Add "Description", ""
			.Add "UOM", ""
			.Add "CreateAlternateID", ""     'Pass "ON" Or "OFF" Or ""
			.Add "CheckOut", ""       'Pass "ON" Or "OFF" Or ""
            .Add "Contact", ""	
			.Add "Address", ""		
			.Add "WebSite", ""
			.Add "Phone", ""			
			.Add "Email", ""
			 .Add "VendorRole",""
			 .Add "VendorStatus",""
			 .Add "CertificationStatus",""
			 .Add "sAction",""
End with


'Declaration of the object as Bid Package
With dicBidPackage  
			 .Add "ID","" 
			.Add "Revision", ""            		 
			.Add "Name", ""    'Mandetory Parameter
			.Add "Description", ""
			.Add "UOM", ""
			.Add "CreateAlternateID", ""     'Pass "ON" Or "OFF" Or ""
			.Add "CheckOut", ""       'Pass "ON" Or "OFF" Or ""
            .Add "RequestedDate", ""	'Pass Date Like "12-Apr-2011 16:41"
			.Add "RequiredPurpose", ""		
End with

'Declaration of the object as dicCommercialPart
With dicCommercialPart  
			 .Add "ID","" 
			.Add "Revision", ""            		 
			.Add "Name", ""    'Mandetory Parameter
			.Add "Description", ""
			.Add "UOM", ""
			.Add "CreateAlternateID", ""     'Pass "ON" Or "OFF" Or ""
			.Add "CheckOut", ""       'Pass "ON" Or "OFF" Or ""
            .Add "DesignRequired", ""	'Pass "ON" Or "OFF" Or ""
			.Add "MakeOrBuy", ""		
End with

'Defination of the object as dicCompanyContact
With dicCompanyContact
	.Add "Title",""
	.Add "FirstName",""
	.Add "LastName",""
	.Add "Suffix",""
	.Add "BusinessPhone",""
	.Add "HomePhone",""
	.Add "Mobile",""
	.Add "Fax",""
	.Add "Pager",""
	.Add "Email",""
	.Add "Description",""
End With

'Defination of the object as dicCompanyLocation
With dicCompanyLocation
	.Add "Name",""
	.Add "LocationCode",""
	.Add "LocationType",""
	.Add "Street",""
	.Add "City",""
	.Add "State",""
	.Add "PostalCode",""
	.Add "Country",""
	.Add "Region",""
	.Add "URL",""
	.Add "Description",""
End With

'Defination of the object dicVendorPart
With dicVendorPart
	.Add "PartNumber",""
	.Add "PartName",""
	.Add "ID",""
	.Add "VendorName",""
	.Add "Location",""
	.Add "Description",""
	.Add "Type",""
	.Add "UOM",""
	.Add "CheckOuItem",""
	.Add "CreateAlternateID",""
	.Add "DesignRequired",""
	.Add "MakeOrBuy",""
End With

'Declaring Dictionary for Detail Item
With dicItemDetailsCreate         
			.Add "Type", ""
            .Add "ID", ""     
            .Add "Revision", ""     
			.Add "Name", ""     
			.Add "Description", ""     
			.Add "UOM", ""     
			.Add "CreateAlternateID", ""     
			.Add "CheckOutOnCreate", ""     
			.Add "ProjectID", ""     
			.Add "PreviousID", ""     
			.Add "SerialNumber", ""     
			.Add "ItemComment", ""     
			.Add "UserData1", ""
			.Add "UserData2", ""
			.Add "UserData3", ""
			.Add "RevProjectID", ""     
			.Add "PreviousVersionID", ""     
			.Add "RevSerialNumber", ""     
			.Add "RevItemComment", ""     
			.Add "RevUserData1", ""
			.Add "RevUserData2", ""
			.Add "RevUserData3", ""
			.Add "ItemAddInfo", ""
			.Add "OrgCageCode", ""			
			.Add "Category", ""			
			.Add "TecDocCategory", ""			
			.Add "NoteCategory", ""			
			.Add "SrcDocID", ""						
			.Add "WA11int", ""	
			.Add "OriginalCageCode",""
			.Add "PartCategory",""
			.Add "FinishItems",""
			.Add "ContractCategory",""
			.Add "DesignCategory",""
			.Add "WorkPkgComplexity",""
			.Add "WorkPkgSecurity",""
			.Add "WorkPkgType",""
            .Add "AvailableProjects",""		
            .Add "SrcDocCategory",""		
            .Add "SrcTechDocCategory",""		
End with

'Declaring Dictionary for Overview Tab Contents
With dicOverviewTabContent
	.Add "ObjectName",""
	.Add "ObjectDescription",""
	.Add "Action",""
End With

'Declaring Dictionary for Properties
With dicProperties
	.Add "EditBox:Description",""
	.Add "EditBox:Name",""
	.Add "EditBox:Website URL",""
	.Add "CheckOut",""
    .Add "CheckBox:Is Shared",""
End With

'Declaring Dictionary for Id Display Rules
With dicIDDisplayRules
	.Add "RuleName",""
End With

'Declaring Dictionary for Audit Log.
With dicAuditLog
	.Add "ObjectID",""
	.Add "ObjectName",""
	.Add "ObjectRevision",""
	.Add "ObjectType",""
	.Add "ObjSequenceNo",""
	.Add "EventType",""
	.Add "Project",""
	.Add "DateCreatedBefore",""
	.Add "DateCreatedAfter",""
	.Add "SecObjectID",""
	.Add "SecObjectName",""
	.Add "SecObjectRevision",""
	.Add "SecObjectType",""
	.Add "SecObjSequenceNo",""
	.Add "GroupName",""
	.Add "UserID",""
	.Add "ErrorCode",""
	.Add "sButton",""
End With

'Declaring Dictionary for Item From Template
With dicItemFromTemplate
	.Add "TemplateID",""
	.Add "ItemID",""
	.Add "Revision",""
	.Add "Name",""
	.Add "NumberOfObject",""
	.Add "Description",""
	.Add "RootItem",""
End With

'Declaring Dictionary for Object Save As
With dicSaveAsObject
	.Add "ID",""
	.Add "RevID",""
	.Add "Description",""
	.Add "Name",""
End With
'Declaring Dictionary for Import To Teamcenter functionality
With dicImportTeamcenter
	.Add "UserID",""
	.Add "Password",""
	.Add "Group",""
	.Add "Role",""
	.Add "ControlFilePath",""
	.Add "ControlFileSheet",""
	.Add "bCloseExcel",""
	.Add "ValidationErrorAllowed",""
	.Add "LoggingMethod",""
	.Add "sErrorMessages",""
	.Add "CreateNewRevision",""
	.Add "ImportLogFileName", ""
	.Add "ImportLogMessage", "" 
	.Add "sErrorMessagesCount", ""
	.Add "sDialogError", ""
	.Add "sButton", ""
	.Add "CancelOperation",""
End With


'Declaring Dictionary for Search Preference
With dicSearchPreferences
	.Add "CaseSensitive",""   ' Value should be "ON" Or "OFF"
	.Add "LatestDatasetVersion",""   ' Value should be "ON" Or "OFF"
	.Add "SearchClassification",""		' Value should be "ON" Or "OFF"
	.Add "EnableHierarchicalTypeSearch",""	' Value should be "ON" Or "OFF"
	.Add "WildcardOption",""          'Value should be Wildcard option radio button name [ SQL Style or Unix Style or Windows Style ]
	.Add "DelimitingCharacter",""
	.Add "EscapeCharacter",""
	.Add "DefaultSearch",""
	.Add "DefaultBOType",""        ' BOType :- Business Object Type
	.Add "SearchLocale",""			
	.Add "FavoriteBOTypeAction",""  'Eg. "Add","Remove","Up","Down"
	.Add "FavoriteBOType",""			 'Favorite Bussiness Object Type to Add or Remove or shift Up Or Shift Down
	.Add "ShiftingCount",""						'How many possition Favorite Bussiness Object Type shift Up Or Down if Dont Pass then byDefault its 1
	.Add "LoadingPageSize",""			
	.Add "OpenSearchResultLimit",""
	.Add "LoadAllLimit",""
End With

'Declaring Dictionary for Lot Operations
With dicLotOperations
	.Add "LotNumber",""   
	.Add "ManufacturersID",""   
	.Add "LotSize",""   
End With

'Declaring Dictionary for Details Dataset Creation
With dicDatasetInfo
	.Add "DatasetType",""   
	.Add "DatasetID",""   
	.Add "Revision",""   
	.Add "DatasetName",""   
	.Add "Descirption",""   
	.Add "ToolUsed",""   
	.Add "ImportFile",""   
	.Add "OpenOnCreate",""  
    .Add "Relation","" 
End With

'Declaring Dictionary for Item Attribute Search ( RDV )
With dicItemIDSearch  
		.Add "bClear","False"	
		.Add "bChangeSearch",""	
		.Add "AdvancedDefaultSearchType", ""	    
		.Add "bClearHistory", ""
		.Add "RememberMyLastSearches", ""     
		.Add "SearchType", ""
		.Add "SearchCriteria", ""
		.Add "bClickOnSearchButton","True"
		.Add "BOMLine",""
End with

'Declaring Dictionary for New IRDC creation ( BMIDE )
With dicIRDC  
		.Add "Name",""	
		.Add "Description", ""
		.Add "AppliesToBusinessObject", ""     
		.Add "Condition", ""
End with

'Declaring Dictionary for New Functionality creation ( BMIDE )
With dicFunctionality  
		.Add "Name",""	
		.Add "DisplayName", ""
		.Add "Description", ""     
		.Add "EnableForVerificationRules", "" 
		.Add "BusinessObjectScope", "" 				
		.Add "SupportedConditionSignature", "" 
		.Add "SubGroupLOV", "" 
End with

'Declaring Dictionary for Id Display Rule Information
With dicIdDisplayRulesInfo  
			 .Add "RuleName",""	
			.Add "SelectedContexts", ""
			.Add "UseDefault",""
End with

'Declaring Dictionary for Deep Copy Rule Information
With dicDeepCopyRuleInfo  
			 .Add "OperationType",""	
			.Add "PropertyType", ""
			.Add "RelationType",""
			.Add "ObjectType",""	
			.Add "Condition", ""
			.Add "ActionType",""
			.Add "TargetPrimary",""	
			.Add "CopyPropertiesOnRelation", ""
			.Add "Required",""
			.Add "Secured",""
			.Add "ReferenceProperty",""
End with
'Declaring Dictionary for Status Indicators Image Files
With dicStatusIndicator
		.Add "Not Started",""
		.Add "In Progress",""
		.Add "Needs Attention",""
		.Add "Late",""
		.Add "Complete",""
		.Add "Abandoned",""
End With

'Declaring Dictionary for Component Information
With dicComponentInfo  
			 .Add "OpenByNameColumns",""	
			 .Add "ItemId",""	
			 .Add "UniqueItem",""	
End with

'Declaring Dictionary for Replace Component Information
With dicReplaceComponentInfo  
			 .Add "ItemId",""	
			 .Add "UniqueItem",""	
End with

'Declaring Dictionary for Content Search in CPD perspective
With dicContentSearch
		.Add "Scheme",""	
		.Add "SearchCriteria",""	
		.Add "TrueShapeFiltering",""	
		.Add "Option",""	
		.Add "DistanceFromOrigin",""	
		.Add "DefinePlaneAxis",""	
		.Add "PartitionNode",""	
End With

'Declaring Dictionary for Business Object Information
With dicBOInfo  
			 .Add "ID",""	
			  .Add "Revision",""
			 .Add "Name",""	
			 .Add "MFK Key1",""
			 .Add "MFK Key2",""
End with

'Declaring Dictionary for Project Assign Information [ Web ]
With dicProjectInfo  
			 .Add "AvailableProjects",""	
			.Add "SelectedProjetcs",""
End with

'Declaring Dictionary for Partition Information in CPD perspective
With dicPartitionInfo
		.Add "PartitionType",""	
		.Add "PartitionID",""	
		.Add "Name",""	
		.Add "Description",""	
		.Add "CreatePartitionItem",""	
		.Add "PartitionItemType",""	
		.Add "CopyEffectivity",""	
		.Add "CheckOutOnCreate",""
		.Add "OpenOnCreate",""
End With

'Declaring Dictionary for BMIDE [ Fn_BMIDE_OperationInputPropertyTableOperations ]
With dicOpsInputPropInfo
		.Add "PropertyName",""
		.Add "Required",""
		.Add "Visible",""
		.Add "Usage",""
		.Add "CompoundObjectType",""
		.Add "CompoundObjectConstant",""
		.Add "Description",""
		.Add "DisplayName",""
		.Add "AttributeType",""
		.Add "StringLength",""
		.Add "ReferenceBusinessObject",""
		.Add "Array",""
		.Add "Unlimited",""
		.Add "MaxLength",""
		.Add "CopyFromOriginal",""
End With

'Declaring Dictionary for Project Remove Information [ Web ]
With dicRemoveProjectInfo  
			 .Add "AvailableProjects",""	
			.Add "SelectedProjetcs",""
End with

'Declaring Dictionary for Web Business Object Information
With dicWebBOInfo  
			 .Add "ID",""		
			  .Add "Revision",""
			 .Add "Name",""
			 .Add "Description",""	
			 .Add "MFK Key1",""
			 .Add "MFK Key2",""
			 .Add "CreateAlternateID",""
			 .Add "CheckOutItemRevisionOnCreate",""
End with

'Declaring Dictionary for Bat file details
With dicBatFileDetails
			 .Add "BatFilePath",""	'Full bat file path with extension ( .bat )
			 .Add "TC_ROOT",""	'TC Root path	
			 .Add "TC_DATA","" 'TC Data path
			 .Add "cdTC_DATA",""  ' True of False
			 .Add "Calltc_profilevarsBat",""	' True of False
			 .Add "cdTC_ROOT",""  ' True of False
			 .Add "Command",""  'Full command
			 .Add "BatFilePastePath",""	'Full bat file path with extension ( .bat )
			 .Add "PsExecPath",""	'Full file path with of [ PsExec.exe ]
			 .Add "CPDCommand","" ' Set Command Value
			 .Add "cdCPDCommand","" ' to Execute Command
End with

'Declaring Dictionary for Web Business Object Information ( Company)
With dicWebCompanyInfo  
			 .Add "Name",""
			 .Add "Description",""	
			 .Add "Location Type",""
End with

'Declaring Dictionary for Parameter Definition Group Information
With dicParameterDefinitionGroupInfo  
			 .Add "ConfigurationItem",""
			 .Add "ID",""
			 .Add "Revision",""
			 .Add "Name",""
			 .Add "Description",""	
			 .Add "GenericComponentID",""	
			 .Add "Represents",""
End with


With dicDuplicateInfo
			 .Add "BOMNode",""
			 .Add "DuplicateAllItems",""
			 .Add "DrawFromRevision",""
			 .Add "RequiredDependencies",""
			 .Add "AllDependencies",""
			 .Add "PartFamilyMasters",""
			 .Add "AssignNewDefaultID",""	
			 .Add "DefaultIDPrefix",""	
			 .Add "DefaultIDSuffix",""
			 .Add "DefaultIDReplace",""
			 .Add "DefaultIDWith",""
			 .Add "ReturnTypeFormat",""
End with

With dicGetHierarchy
			 .Add "Project ID",""			
End with
