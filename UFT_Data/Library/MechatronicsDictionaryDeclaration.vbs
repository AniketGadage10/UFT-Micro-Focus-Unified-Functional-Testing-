Dim dicParameterDefRevGeneralInfo
Dim dicAdditionalParameterDefInfo
Dim dicProperties
Dim dicEditProperties
Dim dicParaGrpDefRevisionInfo
Dim dicViewerTabInfo
Dim dicAvailableParameters
Dim dicOverrideConversionRule

Set dicParameterDefRevGeneralInfo = CreateObject("Scripting.Dictionary")
Set dicAdditionalParameterDefInfo = CreateObject("Scripting.Dictionary")
Set dicProperties = CreateObject("Scripting.Dictionary")
Set dicEditProperties = CreateObject("Scripting.Dictionary")
Set dicParaGrpDefRevisionInfo = CreateObject("Scripting.Dictionary")
Set dicViewerTabInfo = CreateObject("Scripting.Dictionary")
Set dicAvailableParameters = CreateObject("Scripting.Dictionary")
Set dicOverrideConversionRule = CreateObject("Scripting.Dictionary")

'Declaring [ dicParameterDefRevGeneralInfo ] for Parameter Defination revision General Information
With dicParameterDefRevGeneralInfo
   	.Add "Comment",""
	.Add "ParameterDefinitionDescriptor",""
	.Add "SizeUnits",""
	.Add "Size",""
	.Add "ControlEngineer",""
	.Add "IsSigned",""
	.Add "Tolerance",""
	.Add "ResolutionNumerator",""
	.Add "ResolutionDenominator",""
	.Add "Precision",""
	.Add "SizeInByte",""
End With

'Declaring [ dicAdditionalParameterDefInfo ] for additional Parameter Defination Information
With dicAdditionalParameterDefInfo
   	.Add "Comment",""
	.Add "ParameterType",""
	.Add "ButtonName",""
End With

'Declaring [ dicProperties ] for Object properties information
With dicProperties
	.Add "PropertyName",""				'Set of properties	
	.Add "Value",""								   'Set of property values
	.Add "Row",""                  					'Row number
	.Add "Column",""							 'Column nuber or column name [ E.g.  in case of maximum value or initial value table its column number in case Constants table its Column name ]
	.Add "ConstantName",""
	.Add "ConstantValue",""
	.Add "DomainElementName",""
	.Add "Description",""
	.Add "ParameterName",""
	.Add "ColumnName",""
	.Add "Byte",""
	.Add "BitNumber",""			
	.Add "PropertyState",""	
	.Add "CheckPropertyName",""
End With

'Declaring [ dicEditProperties ] for Object editi properties information
With dicEditProperties
	.Add "PropertyName",""				'Set of properties	
	.Add "Value",""								   'Set of property values
	.Add "Row",""                  					'Row number
	.Add "Column",""							 'Column nuber or column name [ E.g.  in case of maximum value or initial value table its column number in case Constants table its Column name ]
	.Add "ConstantName",""
	.Add "ConstantValue",""
	.Add "DomainElementName",""			
	.Add "ParameterName",""
	.Add "Byte",""
	.Add "BitNumber",""
	.Add "Collapse",""
	.Add "Color",""
	.Add "CellErrorMessage",""
	.Add "PropertyState",""
	.Add "DomainElementValue",""
	.Add "CheckPropertyName",""
End With

'Declaring [ dicParaGrpDefRevisionInfo ] for additional Parameter Defination Group Revision Information
With dicParaGrpDefRevisionInfo
   	.Add "Comment",""
	.Add "ControlEngineer",""
	.Add "ParameterGroupDescriptor",""
	.Add "Specialist",""
	.Add "ButtonName",""
End With

'Declaring [ dicViewerTabInfo ] for Viewer Tab Information
With dicViewerTabInfo
	.Add "Row",""                  					'Row number
	.Add "Column",""							 'Column nuber or column name [ E.g.  in case of maximum value or initial value table its column number in case Constants table its Column name ]
	.Add "ShowValueDescription",""
	.Add "Collapse",""
	.Add "CellErrorMessage",""
	.Add "ButtonName",""
	.Add "Value",""
	.Add "PropertyName",""
	.Add "Byte",""
	.Add "BitNumber",""
	.Add "Color",""
    .Add "PageLink",""
	.Add "PropertyState",""
	.Add "DomainElementName",""
	.Add "DomainElementValue",""
	.Add "DomainElementDescription",""
	.Add "ParameterName",""
End With

'Declaring [ dicAvailableParameters ] for Available Parameters Tab     -  [Pranav Ingle : 1-Nov-2012 ]
With dicAvailableParameters
	.Add "TabName",""
	.Add "TableHeader",""
   	.Add "Object",""
	.Add "ColName",""
	.Add "Value",""
	.Add "ShowAllAvailableParameters",""
	.Add "ShowAllAssignedParameters", ""
End With

'Declaring [ dicOverrideConversionRule ] for Creating Override Conversion Rule    -  [Pranav Ingle : 2-Nov-2012 ]
With dicOverrideConversionRule
	.Add  "Action",""
	.Add "Name",""
	.Add "Description",""
	.Add "Type",""
	.Add "Expression","" 
	.Add "ConstantsName",""
	.Add "ConstantsValue",""
	.Add "ShortErrMsg",""
	.Add "DetailErrMsg",""
End With
