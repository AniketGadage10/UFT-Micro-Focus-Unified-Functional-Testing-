'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Dictionary Definitions for Content Management
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'Declaring Dictionary for Publication Type Information
Dim dicPublicationTypeInfo
'Declaring Dictionary for Topic Type Information
Dim dicTopicTypeInfo
'Declaring Dictionary for  S1000D Data Module Information
Dim dicS1000DDataModuleInfo
'Declaring Dictionary for S1000D Publication Module 4.0 Information
Dim dicS1000DPublicationModule4Info
'Declaring Dictionary for DITA Object Information
Dim dicDITAObjectInfo
'Declaring Dictionary for Import Graphic Options Information
Dim dicImportGraphicOptionsInfo
'Declaring Dictionary for S1000D Data Dispatch Note Information
Dim dicS1000DDataDispatchNote4Info
'Declaring Dictionary for Window Preferences Information
Dim dicWindowPreferencesInfo
'Declaring Dictionary for S1000D Data Module List 4.0 Information
Dim dicS1000DDataModuleList4Info
'Declaring Dictionary for S1000D Data Module 4.0 Information
Dim dicS1000DDataModule4Info
'Declaring Dictionary for Commentry 4.0 Information
Dim dicCommentary4info
'Declaring Dictionary for Translation Office Information
Dim dicTranslationOfficeInfo
'Declaring Dictionary for Import DITA Map Information
Dim dicImportDITAMapInfo
'Declaring Dictionary for XML Attribute Map Table Entry
Dim dicXMLAttributeMapInfo
'Declaring Dictionary for Content Publish Information
Dim dicPublishInfo

Set dicPublicationTypeInfo=CreateObject("Scripting.Dictionary")
Set dicTopicTypeInfo=CreateObject("Scripting.Dictionary")
Set dicS1000DDataModuleInfo = CreateObject("Scripting.Dictionary")
Set dicS1000DPublicationModule4Info = CreateObject("Scripting.Dictionary")
Set dicDITAObjectInfo = CreateObject("Scripting.Dictionary")
Set dicImportGraphicOptionsInfo=CreateObject("Scripting.Dictionary")
Set dicS1000DDataDispatchNote4Info=CreateObject("Scripting.Dictionary")
Set dicWindowPreferencesInfo=CreateObject("Scripting.Dictionary")
Set dicS1000DDataModuleList4Info=CreateObject("Scripting.Dictionary")
Set dicS1000DDataModule4Info=CreateObject("Scripting.Dictionary")
Set dicCommentary4info=CreateObject("Scripting.Dictionary")
Set dicTranslationOfficeInfo=CreateObject("Scripting.Dictionary")
Set dicImportDITAMapInfo=CreateObject("Scripting.Dictionary")
Set dicXMLAttributeMapInfo=CreateObject("Scripting.Dictionary")
Set dicPublishInfo=CreateObject("Scripting.Dictionary")

'Declaring Dictionary for Publication Type Information
With dicPublicationTypeInfo  
			 .Add "Name",""
			 .Add "Local Tag Name",""
			 .Add "System Usage",""
			 .Add "Validate Incoming On Parse",""
			 .Add "Validate Outgoing On Parse",""	
			 .Add "Validate Example Content On Parse",""	
			 .Add "Transfer Mode",""
			 .Add "File Extension",""
			 .Add "Apply Classname",""
			 .Add "Namespace URI",""
			 .Add "Default Namespace Prefix",""
End with

'Declaring Dictionary for Topic Type Information
With dicTopicTypeInfo
			 .Add "TopicType",""  
			 .Add "Name",""
			 .Add "Local Tag Name",""
			 .Add "System Usage",""
			 .Add "Validate Incoming On Parse",""
			 .Add "Validate Outgoing On Parse",""	
			 .Add "Validate Example Content On Parse",""	
			 .Add "Transfer Mode",""
			 .Add "File Extension",""
			 .Add "Apply Classname",""
			 .Add "Namespace URI",""
			 .Add "Default Namespace Prefix",""
			 .Add "Reference Type",""
			 .Add "Variant",""
			 .Add "Fragment Tag Names",""
End with

'Declaring Dictionary for S1000D Data Module Information
With dicS1000DDataModuleInfo  
	.Add "TopicType",""
	.Add "Revision",""
	.Add "Name",""
	.Add "MasterLanguageReference",""
	.Add "DocumentTitle",""
	.Add "ModelIdentificationCode",""
	.Add "SystemDifferenceCode",""
	.Add "ChapterNumber",""
	.Add "SectionNumber",""
	.Add "Subsection",""
	.Add "DisassemblyCode",""
	.Add "DisassemblyCodeVariant",""
	.Add "InformationCode",""
	.Add "InformationCodeVariant",""
	.Add "ItemLocationCode",""
	.Add "SupportEquipmentVariantCode",""
	.Add "EquipmentCategoryAndSubCategoryCode",""
	.Add "EquipmentIdentificationCode",""
	.Add "ComponentIdentificationCode",""
	.Add "ExtensionCode",""
	.Add "ExtensionProducer",""
	.Add "ExportFileName",""
	.Add "IsThisATemplate",""
	.Add "ReferenceOnly",""
	.Add "TechnicalName",""
	.Add "InformationName",""
	.Add "IssueNumber",""
	.Add "IssueType",""
	.Add "IssueDay",""
	.Add "IssueMonth",""
	.Add "IssueYear",""
	.Add "SecurityClass",""
	.Add "ResponsiblePartnerCompany",""
	.Add "Originator",""
	.Add "ApplicabilityOfTheMaterial",""
	.Add "ApplicabilityType",""
	.Add "QualityAssurance",""
	.Add "SystemBreakdownCode",""
	.Add "Skill",""
	.Add "ReasonForUpdate",""
	.Add "Remarks",""
	.Add "Level",""	
End with

'Declaring Dictionary for S1000D Publication Module 4.0 Information
With dicS1000DPublicationModule4Info  
	.Add "AuthorClass",""
	.Add "TopicType",""
	.Add "Revision",""
	.Add "Name",""
	.Add "MasterLanguageReference",""
	.Add "DocumentTitle",""
	.Add "ModelIdentificationCode",""
	.Add "IssuingAuthority",""
	.Add "Number",""
	.Add "VolumeOfPublication",""
	.Add "ExtensionCode",""
	.Add "ExtensionProducer",""
	.Add "IssueNumber",""
	.Add "IssueType",""
	.Add "IssueDay",""
	.Add "IssueMonth",""
	.Add "IssueYear",""
	.Add "SecurityClass",""
	.Add "ResponsiblePartnerCompany",""
	.Add "Originator",""
	.Add "QualityAssurance",""
	.Add "SystemBreakdownCode",""
	.Add "Remarks",""
	.Add "Effectivity",""
	.Add "Media",""
	.Add "MediaType",""
	.Add "MediaCode",""
	.Add "FunctionalItemCode",""	
	.Add "ReasonForUpdate",""	
	.Add "InWorkNumber",""
	.Add "ExportFileName",""
	.Add "IsThisATemplate",""
	.Add "ReferenceOnly",""
End with

'Declaring Dictionary for DITA Object Information
With dicDITAObjectInfo  
	.Add "TopicType",""
	.Add "ID",""
	.Add "Revision",""
	.Add "Name",""
	.Add "DocumentTitle",""
	.Add "MasterLanguageReference",""
	.Add "IsThisATemplate",""
	.Add "ReferenceOnly",""
	.Add "DITAAudience",""
	.Add "DITAImportance",""
	.Add "DITAOtherProperties",""
	.Add "DITAPlatform",""
	.Add "DITAProduct",""
End with

'Declaring Dictionary for Import Graphic Options Information
With dicImportGraphicOptionsInfo  
			 .Add "FromDirectory",""
			 .Add "FileNames",""                  'If want to select all files then pass value : [ SelectAll ] else Pass file names by tilda [ ~ ] separated for multiple file selection not all 
			 .Add "GraphicUsage",""			 'if want to select Use Graphic Usages from Graphics Mapping checkbox then pass value : [ Use Graphic Usages from Graphics Mapping OR UseGraphicUsagesfromGraphicsMapping ]	
														   				'if want to select multiple Graphic Usages then pass Graphic Usages type by tilda [ ~ ] separated : eg -  PDF~ICON
			 .Add "GraphicAttributeMapping",""
			 .Add "GraphicClassname",""	
			 .Add "Language",""	
			 .Add "OverwriteMode",""		  'Eg- Skip Existing , Overwrite existing || Name of radio button
			 .Add "Usages",""						  'Eg- Keep , Merge , Overwrite || Name of radio button	
End with

'Declaring Dictionary for S1000D Data Dispatch Note 4.0 Information
With dicS1000DDataDispatchNote4Info  
	.Add "AuthorClass",""
	.Add "TopicType",""
	.Add "Revision",""
	.Add "Name",""
	.Add "MasterLanguageReference",""
	.Add "DocumentTitle",""
	.Add "ModelIdentificationCode",""
	.Add "Originator",""
	.Add "ReceiverIdentification",""
	.Add "YearOfDispatch",""
	.Add "SequenceNumber",""
	.Add "IssueNumber",""
	.Add "IssueType",""
	.Add "IssuedDay",""
	.Add "IssuedMonth",""
	.Add "IssuedYear",""
	.Add "SecurityClass",""
	.Add "AuthorizationIdentification",""
	.Add "MediaIdentification",""
	.Add "Remarks",""
	.Add "InWorkNumber",""
	.Add "ExportFileName",""
	.Add "IsThisATemplate",""
	.Add "ReferenceOnly",""
	.Add "DispatchToEnterpriseName",""
	.Add "DispatchToCity",""
	.Add "DispatchToCountry",""
	.Add "DispatchFromCompanyName",""
	.Add "DispatchFromCity",""
	.Add "DispatchFromCountry",""
End with

'Declaring Dictionary for Window Preferences Information
With dicWindowPreferencesInfo  
			 .Add "Editor",""
End with

'Declaring Dictionary for S1000D Data Module List 4.0 Information
With dicS1000DDataModuleList4Info  
	.Add "AuthorClass",""
	.Add "TopicType",""
	.Add "Revision",""
	.Add "Name",""
	.Add "MasterLanguageReference",""
	.Add "DocumentTitle",""
	.Add "ModelIdentificationCode",""
	.Add "Originator",""
	.Add "TypeOfDataModuleList",""
	.Add "YearOfDispatch",""
	.Add "SequenceNumber",""
	.Add "IssueNumber",""
	.Add "IssueType",""
	.Add "IssuedDay",""
	.Add "IssuedMonth",""
	.Add "IssuedYear",""
	.Add "Remarks",""
	.Add "InWorkNumber",""
	.Add "SecurityClass",""
	.Add "ExportFileName",""
	.Add "IsThisATemplate",""
	.Add "ReferenceOnly",""
	.Add "ID",""
End with

'Declaring Dictionary for S1000D Data Module 4.0 Information
With dicS1000DDataModule4Info  
	.Add "TopicType",""
	.Add "Revision",""
	.Add "Name",""
	.Add "MasterLanguageReference",""
	.Add "DocumentTitle",""
	.Add "ModelIdentifier",""
	.Add "SystemDifferenceCode",""
	.Add "SystemCode",""
	.Add "SubSystemCode",""
	.Add "SubSubSystemCode",""
	.Add "AssemblyCode",""
	.Add "DisassemblyCode",""
	.Add "DisassemblyCodeVariant",""
	.Add "InformationCode",""
	.Add "InformationCodeVariant",""
	.Add "ItemLocationCode",""
	.Add "ExtensionCode",""
	.Add "ExtensionProducer",""
	.Add "InWorkNumber",""
	.Add "ExportFileName",""
	.Add "IsThisATemplate",""
	.Add "ReferenceOnly",""
	.Add "TechnicalName",""
	.Add "InformationName",""
	.Add "IssueNumber",""
	.Add "IssueType",""
	.Add "IssueDay",""
	.Add "IssueMonth",""
	.Add "IssueYear",""
	.Add "SecurityClass",""
	.Add "ResponsiblePartnerCompany",""
	.Add "ResponsiblePartnerCompanyEnterpriseCode",""
	.Add "OriginatorName",""
	.Add "OriginatorEnterpriseCode",""
	.Add "QualityAssurance",""
	.Add "SystemBreakdownCode",""
	.Add "SkillLevel",""
	.Add "ReasonForUpdate",""
	.Add "Remarks",""
End with

'Declaring Dictionary for Commentary 4.0 Information
With dicCommentary4info
		.Add "TopicType",""
		.Add "Revision",""
		.Add "Name",""
		.Add "MasterLanguageReference",""
		.Add "DocumentTitle",""
		.Add "ModelIdentifier",""
		.Add "SenderIdentificationCode",""
		.Add "YearOfDataIssue",""
		.Add "SequentialNumber",""
		.Add "CommentType",""
		.Add "IssueType",""
		.Add "IssueDay",""
		.Add "IssueMonth",""
		.Add "IssueYear",""
		.Add "IssueNumber",""
		.Add "CommentPriorityCode",""
		.Add "Remarks",""
		.Add "InWorkNumber",""
		.Add "CommentResponceType",""
		.Add "SecurityClass",""
		.Add "ExportFileName",""
		.Add "Isthisatemplate",""
		.Add "Referenceonly",""
		.Add "DispatchPersonFirstName",""
		.Add "DispatchPersonSurname",""
		.Add "DispatchPersonJobTitle",""
		.Add "OriginatorEmailAddress",""
		.Add "OriginatorDispatchAddressPhone",""
		.Add "OriginatorDispatchAddressFax",""
		.Add "OriginatorInternetAddress",""
		.Add "OriginatorDispatchAddressDepartment",""
		.Add "OriginatorDispatchAddressBuilding",""
		.Add "OriginatorDispatchAddressRoom",""
		.Add "OriginatorDispatchAddressStreet",""
		.Add "OriginatorDispatchAddressPostOfficeBox",""
		.Add "OriginatorDispatchAddressCity",""
		.Add "OriginatorDispatchAddressState",""
		.Add "OriginatorDispatchAddressZipCode",""
		.Add "OriginatorDispatchAddressProvince",""
		.Add "OriginatorDispatchAddressPostCode",""
		.Add "OriginatorDispatchAddressCountry",""
End with

'Declaring Dictionary for Translation Office Information
With dicTranslationOfficeInfo  
			 .Add "ID",""
			 .Add "Revision",""
			 .Add "Name",""
			 .Add "TranslationOfficeTitle",""
			 .Add "Address",""	
			 .Add "ContactName",""	
			 .Add "Phone",""
			 .Add "Website",""
			 .Add "EmailInbox",""
			 .Add "DeliverComposedContent",""
			 .Add "DeliverDecomposedContent",""
			 .Add "IncludeSupportingData",""
End with

'Declaring Dictionary for Import DITA Map Information
With dicImportDITAMapInfo  
			 .Add "FileNames",""                  
			 .Add "TopicTypeName",""			 'Select Topic Type Name as DITA Dynamic Map or DITA Static Map	
			 .Add "GraphicAttributeMapping",""
			 .Add "GraphicMode",""				  'Eg- Import Original Name, XML Number etc || Name of radio button
			 .Add "ReuseExistingTopic",""		  'Eg- Find by XML Number, Overwrite existing
End with

'Declaring Dictionary for XML Attribute Map Table Entry
With dicXMLAttributeMapInfo
			 .Add "AttributeName",""  
			 .Add "ConstantValue",""
			 .Add "FieldSeparator",""
			 .Add "FixFieldLength",""
			 .Add "Function",""
			 .Add "OmitEmptyAttribute",""	
			 .Add "Path",""	
			 .Add "XMLProcedure",""
End with

'Declaring Dictionary for Content Publish Information
With dicPublishInfo
   	.Add "Tool",""
	.Add "StyleType",""
	.Add "Language",""
	.Add "ComposeVersionSelection",""
	.Add "TranslationVersionSelection",""
	.Add "ResultingFileFolder",""
	.Add "ResultingFileName",""
	.Add "RegisterResult",""
	.Add "DitaFilterValue",""
	.Add "View",""
End With
