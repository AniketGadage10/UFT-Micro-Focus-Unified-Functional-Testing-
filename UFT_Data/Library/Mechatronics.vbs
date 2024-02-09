Option Explicit

' Function List
'= = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'0.  Fn_SISW_Mechatronics_GetObject 
'1.  Fn_Mech_RemoveSignalAssociationOperations
'2.  Fn_Mech_RemoveRealizedByOperations
'3.  Fn_Mech_RemoveImplementedByOperations
'4.  Fn_Mech_RemoveProcessorAssociationOperations
'5.  Fn_Mech_FixInStructureAssociationOperation
'6.  Fn_Mech_ParameterDefinationBasicCreate
'7.  Fn_Mech_AdditionalParameterDefInfo
'8.  Fn_Mech_ParameterDefinationRevisionInfoOperations
'9.  Fn_Mech_ParameterDefRevGeneralInfo
'10. Fn_Mech_TableDefination
'11. Fn_Mech_ConstantsTableOperations
'12. Fn_Mech_ConversionRule
'13. Fn_Mech_MaximumValueTable
'14. Fn_Mech_MinimumValueTable
'15. Fn_Mech_InitialValueTable
'16. Fn_Mech_VerifyProperties
'17. Fn_Mech_EditProperties
'18. Fn_Mech_ValidValuesTable
'19. Fn_Mech_SEDInitialValueTable
'20. Fn_Mech_BitDefination
'21. Fn_Mech_ValidationErrorHandle
'22. Fn_Mech_ParameterDefinationGroupBasicCreate
'23. Fn_Mech_ParameterDefinationGroupRevisionInfo
'25. Fn_SISW_Mech_ParameterValuesTableOperations
'26. Fn_SISW_Mech_EnterActualValueForParameterOperations
'27. Fn_SISW_Mech_ParameterValueProductSearchCriteria
'28. Fn_SISW_Mech_ParameterValueChooseCategoryForProduct
'29. Fn_SISW_Mech_ParameterValueBasicCreate
'30. Fn_SISW_Mech_ParameterValueAdditionalInformation
'31. Fn_SISW_Mech_EnterActualValueForBitParameterOperations
'32. Fn_SISW_Mech_ViewerTabOperations
'33. Fn_SISW_Mech_ViewerTabErrorHandle
'34. Fn_SISW_Mechatronics_SPMNavigationTree
'35. Fn_SISW_Mechatronics_CreateNewBlock
'36 .Fn_SISW_Mech_DeleteOverrideRecord
'37 .Fn_SISW_Mech_AvailableParam_BreakdownOperations
'38. Fn_SISW_Mechatronics_DefineAdditionalMemoryLayoutInfo
'39. Fn_SISW_Mech_OverrideConversionRule
'40.Fn_SISW_Mechatronics_MemoryLayoutBasicCreate
'41. Fn_SISW_Mech_ColumnChooser
'42. Fn_SISW_Mechatronics_DefineMemoryLayoutRevisionInfo 
'43. Fn_SISW_Mech_AvailableParam_DictionaryOperations
'44. Fn_SISW_Mech_ParameterDefRevisionInfoVerifyProperties
'45. Fn_SISW_Mech_NewItemForInsertLevel
'46. Fn_SISW_Mech_SoftwareDesignComponentBasicCreate
'47. Fn_SISW_Mech_ParameterDefinitionBasicCreate
'48. Fn_SISW_Mech_ParameterDefinitionGroupDetailsCreate
'= = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = 
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 			Function to get Object hierarchy  		- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''Function Name		:	Fn_SISW_Mechatronics_GetObject
'
''Description		  	 :  	Function to get Object hierarchy

''Parameters		   :	1. sObjectName : Object Handle name
								
''Return Value		   :  	Object \ Nothing
'
''Examples		     	:	Fn_SISW_Mechatronics_GetObject("ParameterDefinition")

'History:                
'								Developer Name							Date				Rev. No.		Reviewer		Changes Done	
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'									Pranav Ingle		 				4-June-2012				1.0					Sunny
'-----------------------------------------------------------------------------------------------------------------------------------
'	Ashwini Kumar		 25-Sept-2013		2.0								Externalized object hierarchies
'-----------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_Mechatronics_GetObject(sObjectName)
	Dim sObjectXMLPath
	sObjectXMLPath = Fn_GetEnvValue("User", "AutomationDir") & "\TestData\AutomationXML\ObjectXML\Mechatronics.xml"
	Set Fn_SISW_Mechatronics_GetObject = Fn_SISW_Setup_GetObjectFromXML(sObjectXMLPath, sObjectName)
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_Mech_RemoveSignalAssociationOperations

'Description			 :	Function Used to Perform operation Remove Signal Association

'Parameters			   :  1.StrType : Association Type : Source , Target etc. 
'									2.StrAction: Action Name
'								 	3.StrAssociatedBOMLine: BOMLine to remove
'								    4.StrErrorDialogName: Error Dialog caption/Name
'								    5.StrErrorMsg: Error Message
'
'Return Value		   : 	True or False

'Pre-requisite			:	Structure should be selected

'Examples				:   Fn_Mech_RemoveSignalAssociationOperations("Source","Remove","PC-ECU_56789/A;1 (View)","","")
'								     Fn_Mech_RemoveSignalAssociationOperations("Terget","Remove","PC-ECU_56789/A;1 (View)","","")
'								     Fn_Mech_RemoveSignalAssociationOperations("Transmitter","Remove","PC-ECU_56789/A;1 (View)","","")
'									 Fn_Mech_RemoveSignalAssociationOperations("Redundant Signal","Remove","PC-ECU_56789/A;1 (View)","","")
'						Fn_Mech_RemoveSignalAssociationOperations("ProcessVariable","RemoveProcessVariable","","","Do you want to remove all associations with Process Variable")
'                       
'History					 :			
'										Developer Name							Date						Rev. No.				Changes Done											Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'										Sandeep N								19-Mar-2012					1.0																								Pranav Ingle
'										Pranav Ingle							  20-Mar-2012				  1.1				Added Case "RemoveProcessVariable"			   
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Public Function Fn_Mech_RemoveSignalAssociationOperations(StrType,StrAction,StrAssociatedBOMLine,StrErrorDialogName,StrErrorMsg)
	GBL_FAILED_FUNCTION_NAME="Fn_Mech_RemoveSignalAssociationOperations"
 	'Variable Declaration
	Dim objMechDiaolg,StrMenu,bFlag,iRow,iCounter,crrBOMLine
	Fn_Mech_RemoveSignalAssociationOperations=False

	If StrAction="RemoveProcessVariable" Then
		Set objMechDiaolg=JavaDialog("RemovalConfirmation")
	Else
		Set objMechDiaolg=Window("MechatronicsWindow").JavaWindow("WEmbeddedFrame").JavaDialog("RemoveSignalAssociation")
'		Set objMechDiaolg=JavaWindow("Mechatronics").JavaWindow("TcDefaultApplet").JavaDialog("RemoveSignalAssociation")
	End If
	'Creating object of [ Remove Signal Association ] dialog
	If Not objMechDiaolg.Exist(6) Then
		 ' to call menu
		'Selecting Signal Association type : - Source , Target , Transmitter , Redundant Signal
		Select Case StrType
			Case "Source","source"
				StrMenu=Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\Mechatronics\Mechatronics_Menu.xml","RemoveSignalAssociationSource")
			Case "Terget","target"
				StrMenu=Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\Mechatronics\Mechatronics_Menu.xml","RemoveSignalAssociationTarget")
			Case "Transmitter","transmitter"
				StrMenu=Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\Mechatronics\Mechatronics_Menu.xml","RemoveSignalAssociationTransmitter")
			Case "RedundantSignal","redundantsignal"
				StrMenu=Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\Mechatronics\Mechatronics_Menu.xml","RemoveSignalAssociationRedundantSignal")
			Case "ProcessVariable","processvariable"
				StrMenu=Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\Mechatronics\Mechatronics_Menu.xml","RemoveSignalAssociationProcessVariable")
			Case Else
				Set objMechDiaolg=Nothing
				Exit Function
		End Select
		'Calling Menu :
		Call Fn_MenuOperation("Select",StrMenu)
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Successfully called menu [ " & StrMenu & " ]")
	End If
	
	Select Case StrAction
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Remove"
			'Retriving number of rows exist in table
			iRow=Fn_UI_Object_GetROProperty("Fn_Mech_RemoveSignalAssociationOperations",objMechDiaolg.JavaTable("AssociatedBOMLines"),"rows")
			For iCounter=0 To iRow-1
				crrBOMLine=objMechDiaolg.JavaTable("AssociatedBomLines").GetCellData(iCounter,"Associated BOMLines")
				If Trim(crrBOMLine)=Trim(StrAssociatedBOMLine) Then
					Call Fn_Table_Select_Cell("Fn_Mech_RemoveSignalAssociationOperations", objMechDiaolg, "AssociatedBOMLines",iCounter,"Associated BOMLines")
					'Clicking Remove button to remove BOMLine
					Call Fn_Button_Click("Fn_Mech_RemoveSignalAssociationOperations", objMechDiaolg,"Remove")
					Fn_Mech_RemoveSignalAssociationOperations=True
				End If
			Next
			wait 3
			If objMechDiaolg.Exist(6) Then
				Call Fn_Button_Click("Fn_Mech_RemoveSignalAssociationOperations", objMechDiaolg,"Cancel")
			End If

		Case "RemoveProcessVariable"
			If StrErrorMsg<>"" Then
				crrBOMLine= JavaDialog("RemovalConfirmation").JavaObject("ErrorMsg").Object.getText
				If Instr(1,Trim(crrBOMLine),Trim(StrErrorMsg))<=0 Then
					Fn_Mech_RemoveSignalAssociationOperations=False
					Exit Function
				End If
			End If
			Call Fn_Button_Click("Fn_Mech_RemoveSignalAssociationOperations", JavaDialog("RemovalConfirmation"),"Yes")
			Fn_Mech_RemoveSignalAssociationOperations=True

		Case Else
			Fn_Mech_RemoveSignalAssociationOperations=False
	End Select
	'Releasing object of [ Remove Signal Association ] dialog
	Set objMechDiaolg=Nothing
End Function

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_Mech_RemoveRealizedByOperations

'Description			 :	Function Used to Perform operation Remove Realized By

'Parameters			   :    1.StrAction: Action Name
'								 2.StrAssociatedBOMLine: BOMLine to remove
'								 3.StrErrorDialogName: Error Dialog caption/Name
'								 4.StrErrorMsg: Error Message
'
'Return Value		   : 	True or False

'Pre-requisite			:	Structure should be selected

'Examples				:   Fn_Mech_RemoveRealizedByOperations("RemoveErrorVerify","NWPort_44446.1","","You do not have write access to object Root_Item_44446/A.001-View")
'								Fn_Mech_RemoveRealizedByOperations("RemoveErrorVerify","NWPort_44446.1","Realized By","You do not have write access to object Root_Item_44446/A.001-View")
'								Fn_Mech_RemoveRealizedByOperations("Remove","NWPort_44446.1","","")
'                       
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												14-Mar-2012								1.0																						Sunny R
'													Pranav Ingle											11-Jul-2012								   1.1																						Sandeep
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Public Function Fn_Mech_RemoveRealizedByOperations(StrAction,StrAssociatedBOMLine,StrErrorDialogName,StrErrorMsg)
	GBL_FAILED_FUNCTION_NAME="Fn_Mech_RemoveRealizedByOperations"
 	'Variable Declaration
	Dim objMechDiaolg,StrMenu,bFlag,iRow,iCounter,crrBOMLine
	Fn_Mech_RemoveRealizedByOperations=False

'	Modified To handle Dialog   Window("MechatronicsWindow").JavaApplet("JApplet").JavaDialog("RemoveReallizedBy")  - Pranav 

	'Creating object of [ Remove Realized By ] dialog
	If Not JavaWindow("Mechatronics").JavaWindow("TcDefaultApplet").JavaDialog("RemoveReallizedBy").Exist(5) And Not Window("MechatronicsWindow").JavaWindow("WEmbeddedFrame").JavaDialog("RemoveReallizedBy").Exist(5) Then
			StrMenu=Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\Mechatronics\Mechatronics_Menu.xml","RemoveRealizedBy")
			'Calling Menu : Tools:Implemented By:Remove Realized By
			Call Fn_MenuOperation("Select",StrMenu)
	End If

	If Fn_UI_ObjectExist("Fn_Mech_RemoveRealizedByOperations",JavaWindow("Mechatronics").JavaWindow("TcDefaultApplet").JavaDialog("RemoveReallizedBy")) Then
		Set objMechDiaolg = JavaWindow("Mechatronics").JavaWindow("TcDefaultApplet").JavaDialog("RemoveReallizedBy")
	ElseIf Fn_UI_ObjectExist("Fn_Mech_RemoveRealizedByOperations",Window("MechatronicsWindow").JavaWindow("WEmbeddedFrame").JavaDialog("RemoveReallizedBy")) Then
		Set objMechDiaolg = Window("MechatronicsWindow").JavaWindow("WEmbeddedFrame").JavaDialog("RemoveReallizedBy")
	End If

	Select Case StrAction
	    ' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Remove"
			'Retriving number of rows exist in table
			iRow=Fn_UI_Object_GetROProperty("Fn_Mech_RemoveRealizedByOperations",objMechDiaolg.JavaTable("AssociatedBomLines"),"rows")
			For iCounter=0 To iRow-1
				crrBOMLine=objMechDiaolg.JavaTable("AssociatedBomLines").GetCellData(iCounter,"Associated BOMLines")
				If Trim(crrBOMLine)=Trim(StrAssociatedBOMLine) Then
					Call Fn_Table_Select_Cell("Fn_Mech_RemoveRealizedByOperations", objMechDiaolg, "AssociatedBomLines",iCounter,"Associated BOMLines")
					'Clicking Remove button to remove BOMLine
					Call Fn_Button_Click("Fn_Mech_RemoveRealizedByOperations", objMechDiaolg,"Remove")
					Fn_Mech_RemoveRealizedByOperations=True
				End If
			Next
			wait 3
			If objMechDiaolg.Exist(6) Then
				Call Fn_Button_Click("Fn_Mech_RemoveRealizedByOperations", objMechDiaolg,"Cancel")
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "RemoveErrorVerify"
			'Retriving number of rows exist in table
			iRow=Fn_UI_Object_GetROProperty("Fn_Mech_RemoveRealizedByOperations",objMechDiaolg.JavaTable("AssociatedBomLines"),"rows")
			For iCounter=0 To iRow-1
				crrBOMLine=objMechDiaolg.JavaTable("AssociatedBomLines").GetCellData(iCounter,"Associated BOMLines")
				If Trim(crrBOMLine)=Trim(StrAssociatedBOMLine) Then
					Call Fn_Table_Select_Cell("Fn_Mech_RemoveRealizedByOperations", objMechDiaolg, "AssociatedBomLines",iCounter,"Associated BOMLines")
					'Clicking Remove button to remove BOMLine
					Call Fn_Button_Click("Fn_Mech_RemoveRealizedByOperations", objMechDiaolg,"Remove")
					If StrErrorDialogName="" Then
						StrErrorDialogName="Realized By"
					End If
					bFlag=Fn_ErrorDialogMessageVerify(StrErrorDialogName, StrErrorMsg, "EditBoxErrMsgVerify_New")
					If bFlag=True Then
						Fn_Mech_RemoveRealizedByOperations=True
					End If
					Call Fn_Button_Click("Fn_Mech_RemoveRealizedByOperations", objMechDiaolg,"Cancel")
				End If
			Next
		Case Else
			Fn_Mech_RemoveRealizedByOperations=False
	End Select
	Set objMechDiaolg=Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_Mech_RemoveImplementedByOperations

'Description			 :	Function Used to Perform operation Remove Implemented By

'Parameters			   :    1.StrAction: Action Name
'								 2.StrAssociatedBOMLine: BOMLine to remove
'								 3.StrErrorDialogName: Error Dialog caption/Name
'								 4.StrErrorMsg: Error Message
'
'Return Value		   : 	True or False

'Pre-requisite			:	Structure should be selected

'Examples				:   Fn_Mech_RemoveImplementedByOperations("RemoveErrorVerify","HRN_Gen_Wire_44446/A;1","","You do not have write access to object Root_Item_44446/A.001-View")
'								Fn_Mech_RemoveImplementedByOperations("RemoveErrorVerify","HRN_Gen_Wire_44446/A;1","Implemented By","You do not have write access to object Root_Item_44446/A.001-View")
'								Fn_Mech_RemoveImplementedByOperations("Remove","HRN_Gen_Wire_44446/A;1","","")
'                       
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												15-Mar-2012								1.0																	Priyanka B
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Public Function Fn_Mech_RemoveImplementedByOperations(StrAction,StrAssociatedBOMLine,StrErrorDialogName,StrErrorMsg)
	GBL_FAILED_FUNCTION_NAME="Fn_Mech_RemoveImplementedByOperations"
 	'Variable Declaration
	Dim objMechDiaolg,StrMenu,bFlag,iRow,iCounter,crrBOMLine
	Fn_Mech_RemoveImplementedByOperations=False

	'Creating object of [ Remove Implemented By ] dialog
	If Not JavaWindow("Mechatronics").JavaWindow("TcDefaultApplet").JavaDialog("RemoveImplementedBy").Exist(5) Then
		If Not Window("MechatronicsWindow").JavaWindow("WEmbeddedFrame").JavaDialog("RemoveImplementedBy").Exist(5) Then
			StrMenu=Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\Mechatronics\Mechatronics_Menu.xml","RemoveImplementedBy")
			'Calling Menu : Tools:Implemented By:Remove Implemented By
			Call Fn_MenuOperation("Select",StrMenu)
		End If
	End If

	If Fn_UI_ObjectExist("Fn_Mech_RemoveImplementedByOperations",JavaWindow("Mechatronics").JavaWindow("TcDefaultApplet").JavaDialog("RemoveImplementedBy")) Then
		Set objMechDiaolg = JavaWindow("Mechatronics").JavaWindow("TcDefaultApplet").JavaDialog("RemoveImplementedBy")
	ElseIf Fn_UI_ObjectExist("Fn_Mech_RemoveImplementedByOperations",Window("MechatronicsWindow").JavaWindow("WEmbeddedFrame").JavaDialog("RemoveImplementedBy")) Then
		Set objMechDiaolg = Window("MechatronicsWindow").JavaWindow("WEmbeddedFrame").JavaDialog("RemoveImplementedBy")
	End If

	Select Case StrAction
	    ' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Remove"
			'Retriving number of rows exist in table
			iRow=Fn_UI_Object_GetROProperty("Fn_Mech_RemoveImplementedByOperations",objMechDiaolg.JavaTable("AssociatedBomLines"),"rows")
			For iCounter=0 To iRow-1
				crrBOMLine=objMechDiaolg.JavaTable("AssociatedBomLines").GetCellData(iCounter,"Associated BOMLines")
				If Trim(crrBOMLine)=Trim(StrAssociatedBOMLine) Then
					Call Fn_Table_Select_Cell("Fn_Mech_RemoveImplementedByOperations", objMechDiaolg, "AssociatedBomLines",iCounter,"Associated BOMLines")
					'Clicking Remove button to remove BOMLine
					Call Fn_Button_Click("Fn_Mech_RemoveImplementedByOperations", objMechDiaolg,"Remove")
					Fn_Mech_RemoveImplementedByOperations=True
				End If
			Next
			wait 3
			If objMechDiaolg.Exist(6) Then
				Call Fn_Button_Click("Fn_Mech_RemoveImplementedByOperations", objMechDiaolg,"Cancel")
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "RemoveErrorVerify"
			'Retriving number of rows exist in table
			iRow=Fn_UI_Object_GetROProperty("Fn_Mech_RemoveImplementedByOperations",objMechDiaolg.JavaTable("AssociatedBomLines"),"rows")
			For iCounter=0 To iRow-1
				crrBOMLine=objMechDiaolg.JavaTable("AssociatedBomLines").GetCellData(iCounter,"Associated BOMLines")
				If Trim(crrBOMLine)=Trim(StrAssociatedBOMLine) Then
					Call Fn_Table_Select_Cell("Fn_Mech_RemoveImplementedByOperations", objMechDiaolg, "AssociatedBomLines",iCounter,"Associated BOMLines")
					'Clicking Remove button to remove BOMLine
					Call Fn_Button_Click("Fn_Mech_RemoveImplementedByOperations", objMechDiaolg,"Remove")
					If StrErrorDialogName="" Then
						StrErrorDialogName="Implemented By"
					End If
					bFlag=Fn_ErrorDialogMessageVerify(StrErrorDialogName, StrErrorMsg, "EditBoxErrMsgVerify_New")
					If bFlag=True Then
						Fn_Mech_RemoveImplementedByOperations=True
					End If
					Call Fn_Button_Click("Fn_Mech_RemoveImplementedByOperations", objMechDiaolg,"Cancel")
				End If
			Next
		Case Else
			Fn_Mech_RemoveImplementedByOperations=False
	End Select
	'Releasing object of [ Remove Implemented By ] dialog
	Set objMechDiaolg=Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_Mech_RemoveProcessorAssociationOperations

'Description			 :	Function Used to Perform operation Remove Processor Association

'Parameters			   :  1.StrType : Association Type : Software , Gateway etc. 
'									2.StrAction: Action Name
'								 	3.StrAssociatedBOMLine: BOMLine to remove
'								    4.StrErrorDialogName: Error Dialog caption/Name
'								    5.StrErrorMsg: Error Message
'
'Return Value		   : 	True or False

'Pre-requisite			:	Structure should be selected

'Examples				:   Fn_Mech_RemoveProcessorAssociationOperations("Software","Remove","ABS-Processor1_23456/A;1","","")
'								     Fn_Mech_RemoveProcessorAssociationOperations("Gateway","Remove","ABS-Processor1_23456/A;1","","")
'                       
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												19-Mar-2012								1.0																				Sonal P
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Public Function Fn_Mech_RemoveProcessorAssociationOperations(StrType,StrAction,StrAssociatedBOMLine,StrErrorDialogName,StrErrorMsg)
	GBL_FAILED_FUNCTION_NAME="Fn_Mech_RemoveProcessorAssociationOperations"
 	'Variable Declaration
	Dim objMechDiaolg,StrMenu,bFlag,iRow,iCounter,crrBOMLine
	Fn_Mech_RemoveProcessorAssociationOperations=False
	'Creating object of [ Remove Processor Association ] dialog
'	Set objMechDiaolg=JavaWindow("Mechatronics").JavaWindow("TcDefaultApplet").JavaDialog("RemoveProcessorAssociation")
	Set objMechDiaolg=Window("MechatronicsWindow").JavaWindow("WEmbeddedFrame").JavaDialog("RemoveProcessorAssociation")
	If Not objMechDiaolg.Exist(6) Then
		 ' to call menu
		'Selecting Processor Association type : - Software , Gateway
		Select Case StrType
			Case "Software","software"
				StrMenu=Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\Mechatronics\Mechatronics_Menu.xml","RemoveProcessorAssociationSoftware")
			Case "Gateway","gateway"
				StrMenu=Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\Mechatronics\Mechatronics_Menu.xml","RemoveProcessorAssociationGateway")
			Case Else
				Set objMechDiaolg=Nothing
				Exit Function
		End Select
		'Calling Menu :
		Call Fn_MenuOperation("Select",StrMenu)
        Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Successfully called menu [ " & StrMenu & " ]")
	End If
	
	Select Case StrAction
	    ' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Remove"
			'Retriving number of rows exist in table
			iRow=Fn_UI_Object_GetROProperty("Fn_Mech_RemoveProcessorAssociationOperations",objMechDiaolg.JavaTable("AssociatedBOMLines"),"rows")
			For iCounter=0 To iRow-1
				crrBOMLine=objMechDiaolg.JavaTable("AssociatedBomLines").GetCellData(iCounter,"Associated BOMLines")
				If Trim(crrBOMLine)=Trim(StrAssociatedBOMLine) Then
					Call Fn_Table_Select_Cell("Fn_Mech_RemoveProcessorAssociationOperations", objMechDiaolg, "AssociatedBOMLines",iCounter,"Associated BOMLines")
					'Clicking Remove button to remove BOMLine
					Call Fn_Button_Click("Fn_Mech_RemoveProcessorAssociationOperations", objMechDiaolg,"Remove")
					Fn_Mech_RemoveProcessorAssociationOperations=True
				End If
			Next
			wait 3
			If objMechDiaolg.Exist(6) Then
				Call Fn_Button_Click("Fn_Mech_RemoveProcessorAssociationOperations", objMechDiaolg,"Cancel")
			End If
		Case Else
			Fn_Mech_RemoveProcessorAssociationOperations=False
	End Select
	'Releasing object of [ Remove Processor Association ] dialog
	Set objMechDiaolg=Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_Mech_FixInStructureAssociationOperation

'Description			 :	Function Used to Perform operation Remove Processor Association

'Parameters			   :  1.StrLevel : Level Type : Eg- Current Level , All Level 
'									2.StrAction: Action Name
'								 	3.StrPrimaryObject: Primary Object Name [ Seperated by "~" if more than 1 name]
'									4.StrAssociationType: Association Type [ Seperated by "~" if more than 1 name] 
'								    5.StrErrorMsg: Error Message [ Seperated by "~" if more than 1 name]  
'									6.StrButtonName : Cancel Button if you want to close the dialog
'
'Return Value		   : 	True or False

'Pre-requisite			:	Structure should be selected

'Examples				:   1. Case "ValidAllLevelMessageVerify"
'										Fn_Mech_FixInStructureAssociationOperation("AllLevel","ValidAllLevelMessageVerify","","","No invalid associations are found in the selected level","")
'                       
'								:   2. Case "ShowAssociation"
'										Fn_Mech_FixInStructureAssociationOperation("AllLevel","ShowAssociation","CT_Network_1_40832/A;1","","","")
'
'								:   3. Case "VerifyReason"
'										a.  Single Value
'											Fn_Mech_FixInStructureAssociationOperation("AllLevel","VerifyReason", "CT_Network_1_40832/A;1","TC_Connected_To","Context for the relation is changed.","")
'										b.  Multi Value
'											Fn_Mech_FixInStructureAssociationOperation("AllLevel","VerifyReason", "CT_Network_1_40832/A;1~ConnectionTerminal2_40832","TC_Connected_To~TC_Realized_By","Context for the relation is changed~Ancestral hierarchy violated","Cancel")
'									4. Case "RemoveAssociation"
'										a.  Single Value
'											Fn_Mech_FixInStructureAssociationOperation("AllLevel","RemoveAssociation", "CT_Network_1_40832/A;1","","","")
'										b.  Multi Value
'											Fn_Mech_FixInStructureAssociationOperation("AllLevel","RemoveAssociation", "CT_Network_1_40832/A;1~ConnectionTerminal2_40832","","","Cancel")
										
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Pranav Ingle											28-Mar-2012								1.0																				Sonal P
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Public Function Fn_Mech_FixInStructureAssociationOperation(StrLevel,StrAction,StrPrimaryObject,StrAssociationType,StrErrorMsg,StrButtonName)
	GBL_FAILED_FUNCTION_NAME="Fn_Mech_FixInStructureAssociationOperation"
   'Variable Declaration
	Dim objMechDiaolg,StrMenu,bFlag,sAppMsg
   	Dim iCounter,iRow,strPrimary,arrPrimary,intCount
	Dim arrAssoType,arrReason, strAssoType,strReason
   Fn_Mech_FixInStructureAssociationOperation=False

	'Create object of FixInStructureAssociations JavaWindow
    Set objMechDiaolg=JavaWindow("Mechatronics").JavaWindow("FixInStructureAssociations")

		'Click yes Button of All levels Confirmation Dialog 
    If Window("MechatronicsWindow").JavaWindow("WEmbeddedFrame").JavaDialog("AllLevels").Exist(3) Then 'Added by Mohit Mishra
	   Call Fn_Button_Click("Fn_Mech_FixInStructureAssociationOperation",Window("MechatronicsWindow").JavaWindow("WEmbeddedFrame").JavaDialog("AllLevels"),"Yes")
	End If

	If Not objMechDiaolg.Exist(6) Then
		Select Case StrLevel
		Case "CurrentLevel"
			StrMenu=Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\Mechatronics\Mechatronics_Menu.xml","ToolsFixInStructureCurrentLevel")
		Case "AllLevel"
			StrMenu=Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\Mechatronics\Mechatronics_Menu.xml","ToolsFixInStructureAllLevels")
		Case Else
			Set objMechDiaolg=Nothing
			Exit Function
		End Select
		'Calling Menu :
		Call Fn_MenuOperation("Select",StrMenu)
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Successfully called menu [ " & StrMenu & " ]")
	End If

	'Click yes Button of All levels Confirmation Dialog
	If StrLevel="AllLevel" Then
	  If Window("MechatronicsWindow").JavaWindow("WEmbeddedFrame").JavaDialog("AllLevels").Exist(3) Then ' Added by Mohit Mishra
	   Call Fn_Button_Click("Fn_Mech_FixInStructureAssociationOperation",Window("MechatronicsWindow").JavaWindow("WEmbeddedFrame").JavaDialog("AllLevels"),"Yes")
	End If
	End If
	
	'As Relation display name has been changed need to update
	StrAssociationType = Replace(StrAssociationType,"TC_Realized_By","TC Realized By")
	StrAssociationType = Replace(StrAssociationType,"TC_Implemented_By","TC Implemented By")	
	StrAssociationType = Replace(StrAssociationType,"SIG_asystem_target","SIG Asystem Target")
    StrAssociationType = Replace(StrAssociationType,"SIG_asystem_source","SIG Asystem Source")
    StrAssociationType = Replace(StrAssociationType,"TC_Connected_To","TC Connected To")
    StrAssociationType = Replace(StrAssociationType,"SIG_pvariable","SIG P-variable")					'' TC112-2015071500-29_07_2015-HCMaintenance-PiyushP-Changed Value as per design change
    StrAssociationType = Replace(StrAssociationType,"SIG_redundant","SIG Redundant")
	StrAssociationType = Replace(StrAssociationType,"SIG_asystem_transmitter","SIG Asystem Transmitter")
	
   Select Case StrAction
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
 	Case "ValidAllLevelMessageVerify"
      	sAppMsg = trim(objMechDiaolg.JavaStaticText("ErrMsg").GetROProperty("label"))
		If Instr(1,trim(StrErrorMsg), trim(sAppMsg))>0 then							
			Fn_Mech_FixInStructureAssociationOperation = True						
		End if
	If objMechDiaolg.JavaButton("OK").Exist(5) Then
	   Call Fn_Button_Click("Fn_Mech_FixInStructureAssociationOperation",objMechDiaolg,"OK")
	End If

	'Check existance in case of Fix-In Structure fails
	If objMechDiaolg.Exist(3) Then
		Call Fn_Button_Click("Fn_Mech_FixInStructureAssociationOperation", objMechDiaolg,"Cancel")
	End If

	Case "ShowAssociation"
        iRow=Fn_UI_Object_GetROProperty("Fn_Mech_FixInStructureAssociationOperation",objMechDiaolg.JavaTable("InvalidAssociationsTable"),"rows")
		For iCounter=0 To iRow-1
			strPrimary=objMechDiaolg.JavaTable("InvalidAssociationsTable").GetCellData(iCounter,"Primary")
			If Trim(strPrimary)=Trim(StrPrimaryObject) Then
				objMechDiaolg.JavaTable("InvalidAssociationsTable").ActivateRow iCounter
				'Clicking Remove button to remove BOMLine
				Call Fn_Button_Click("Fn_Mech_FixInStructureAssociationOperation", objMechDiaolg,"ShowAssociation")
				Fn_Mech_FixInStructureAssociationOperation=True
			End If
		Next
		Call Fn_ReadyStatusSync(1)
		Call Fn_Button_Click("Fn_Mech_FixInStructureAssociationOperation", objMechDiaolg,"Cancel")

	Case "RemoveAssociation"
		arrPrimary=Split(StrPrimaryObject,"~")
		iRow=Fn_UI_Object_GetROProperty("Fn_Mech_FixInStructureAssociationOperation",objMechDiaolg.JavaTable("InvalidAssociationsTable"),"rows")
		For intCount=0 To UBound(arrPrimary)
			bFlag=False
			For iCounter=0 To iRow-1
				strPrimary=objMechDiaolg.JavaTable("InvalidAssociationsTable").GetCellData(iCounter,"Primary")
				If Trim(strPrimary)=Trim(arrPrimary(intCount)) Then
					If intCount=0 Then
						objMechDiaolg.JavaTable("InvalidAssociationsTable").ActivateRow iCounter
					Else
						objMechDiaolg.JavaTable("InvalidAssociationsTable").ExtendRow iCounter
					End If
					bFlag=True
					Exit For
				End If
			Next
			If bFlag=False Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [Fn_Mech_FixInStructureAssociationOperation] Failed to find "+arrPrimary(intCount)+" in Fn_Mech_FixInStructureAssociationOperation table")
				Exit Function
			End If
		Next

		'Clicking Remove button
		Call Fn_Button_Click("Fn_Mech_FixInStructureAssociationOperation", objMechDiaolg,"Remove")
		Fn_Mech_FixInStructureAssociationOperation=True
		Call Fn_ReadyStatusSync(1)

	Case "VerifyReason"

		arrPrimary=Split(StrPrimaryObject,"~")
		If StrAssociationType<>"" Then
			arrAssoType=Split(StrAssociationType,"~")
		End If
		If StrErrorMsg<>"" Then
			arrReason=Split(StrErrorMsg,"~")
		End If

		iRow=Fn_UI_Object_GetROProperty("Fn_Mech_FixInStructureAssociationOperation",objMechDiaolg.JavaTable("InvalidAssociationsTable"),"rows")
		For intCount=0 To UBound(arrPrimary)
			bFlag=False
			For iCounter=0 To iRow-1
				strPrimary=objMechDiaolg.JavaTable("InvalidAssociationsTable").GetCellData(iCounter,"Primary")
				If Trim(strPrimary)=Trim(arrPrimary(intCount)) Then
					strAssoType=objMechDiaolg.JavaTable("InvalidAssociationsTable").GetCellData(iCounter,"Association Type")
					strReason=objMechDiaolg.JavaTable("InvalidAssociationsTable").GetCellData(iCounter,"Reason")
					If Trim(strAssoType)=Trim(arrAssoType(intCount)) And Instr(1,Trim(strReason),Trim(arrReason(intCount)))>0 Then
						Fn_Mech_FixInStructureAssociationOperation=True
						bFlag=True
						Exit For
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [Fn_Mech_FixInStructureAssociationOperation] Failed to match Reason and Asso type Properties for "+arrPrimary(intCount))
						Fn_Mech_FixInStructureAssociationOperation=False
						Exit Function
					End If
				End If
			Next
			If bFlag=False Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [Fn_Mech_FixInStructureAssociationOperation] Failed to find "+arrPrimary(intCount)+" in Fn_Mech_FixInStructureAssociationOperation table")
				Fn_Mech_FixInStructureAssociationOperation=False
				Exit Function
			End If
		Next
   End Select

	If StrButtonName="Cancel" Then
		If objMechDiaolg.Exist(3) Then
			Call Fn_Button_Click("Fn_Mech_FixInStructureAssociationOperation", objMechDiaolg,"Cancel")
		End If
	End If	  
   Set objMechDiaolg=Nothing	
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_Mech_ParameterDefinationBasicCreate

'Description			 :	Function Used to create basic Parameter  Defenation

'Parameters			   :   1.StrParaDefType: Parameter  Defenation type
'										2.bConfigItem: Configuration Item
'										3.StrID: Parameter  Defenation ID
'										4.StrRevision: Parameter  Defenation Revision
'										5.StrName: Parameter  Defenation Name
'										6.StrDescription: Parameter  Defenation Description
'										7.UOM: Unit of measure
'										8.StrButtonName: Button Name
'
'Return Value		   : 	Item Id - revision or False

'Pre-requisite			:	Should be log in RAC

'Examples				:   bReturn=Fn_Mech_ParameterDefinationBasicCreate("ParmDefBCD","","","","ParaDefBCD6","Parameter defination for BDC","","Next")
'
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												20-Apr-2012								1.0																						Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_Mech_ParameterDefinationBasicCreate(StrParaDefType,bConfigItem,StrID,StrRevision,StrName,StrDescription,UOM,StrButtonName)
	GBL_FAILED_FUNCTION_NAME="Fn_Mech_ParameterDefinationBasicCreate"
 	'declaring variables
	Dim objParameterDefinitionDialog
	Dim bFlag,crrID,crrRevision,hieght,width, sParameterDefinationmenu
	StrParaDefType = Fn_SISW_MechCurrentobjName(StrParaDefType)
	Fn_Mech_ParameterDefinationBasicCreate=false
	'Checking existance of [ NewParameterDefinition ] dialog
	sParameterDefinationMenu=Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("Mechatronics_Menu"), "NewParameterDefinition")
    If Environment.Value("ProductName") = sUFTProductName Then
		If not Window("MechatronicsWindow").JavaDialog("NewParameterDefinition").Exist(SISW_MINLESS_TIMEOUT) Then
			bFlag = Fn_MenuOperation("WinMenuSelect",sParameterDefinationMenu)
			Call  Fn_ReadyStatusSync(1)
			If bFlag=false Then
				exit function
			End If
		End If
	Else
	   If not Window("MechatronicsWindow").JavaDialog("NewParameterDefinition").Exist(SISW_MINLESS_TIMEOUT) Then
			bFlag = Fn_MenuOperation("Select",sParameterDefinationMenu)
			Call  Fn_ReadyStatusSync(1)
			If bFlag=false Then
				exit function
			End If
		End If
	End If
	
	'Creating object of [ NewParameterDefinition ] dialog
	set objParameterDefinitionDialog=Window("MechatronicsWindow").JavaDialog("NewParameterDefinition")
	'selecting Parameter Defination Type
	Call Fn_List_Select("Fn_Mech_ParameterDefinationBasicCreate", objParameterDefinitionDialog,"ParameterDefinitionList",StrParaDefType)
	 ' Wait till  Button is Enabled
	objParameterDefinitionDialog.JavaButton("Next").WaitProperty "enabled", 1, 60000
	'Click on "Next" button
	objParameterDefinitionDialog.JavaButton("Next").Click micLeftBtn

	'setting Parameter Definition ID
	If StrID<>"" Then
		Call Fn_Edit_Box("Fn_Mech_ParameterDefinationBasicCreate",objParameterDefinitionDialog,"ID",StrID)
	End If
	'setting Parameter Definition Revision
	If StrRevision<>"" Then
		Call Fn_Edit_Box("Fn_Mech_ParameterDefinationBasicCreate",objParameterDefinitionDialog,"Revision",StrRevision)
	End If
	'clicking on assign button to assign ID and Revision
	If StrID="" or StrRevision="" Then
		Call Fn_Button_Click("Fn_Mech_ParameterDefinationBasicCreate", objParameterDefinitionDialog, "Assign")
	End If
	'retriving ID and Revision
	crrID=Fn_Edit_Box_GetValue("Fn_Mech_ParameterDefinationBasicCreate",objParameterDefinitionDialog,"ID")
	crrRevision=Fn_Edit_Box_GetValue("Fn_Mech_ParameterDefinationBasicCreate",objParameterDefinitionDialog,"Revision")
	If crrID="" Then
		Call Fn_Button_Click("Fn_Mech_ParameterDefinationBasicCreate", objParameterDefinitionDialog, "Assign")
		crrID=Fn_Edit_Box_GetValue("Fn_Mech_ParameterDefinationBasicCreate",objParameterDefinitionDialog,"ID")
		crrRevision=Fn_Edit_Box_GetValue("Fn_Mech_ParameterDefinationBasicCreate",objParameterDefinitionDialog,"Revision")
	End If
	'setting Parameter Definition Name
	If StrName<>"" Then
		Call Fn_Edit_Box("Fn_Mech_ParameterDefinationBasicCreate",objParameterDefinitionDialog,"Name",StrName)
	End If
	'setting Parameter Definition Description
	If StrDescription<>"" Then
		Call Fn_Edit_Box("Fn_Mech_ParameterDefinationBasicCreate",objParameterDefinitionDialog,"Description",StrDescription)
	End If
	Fn_Mech_ParameterDefinationBasicCreate="'"&crrID+"-"+crrRevision
	If StrButtonName<>"" Then
		If lcase(StrButtonName)="next" Then
			Call Fn_Button_Click("Fn_Mech_ParameterDefinationBasicCreate", objParameterDefinitionDialog, "Next")
			wait 2
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
			'Resizing window
			hieght=JavaWindow("Mechatronics").GetROProperty("height")
			width=JavaWindow("Mechatronics").GetROProperty("width")
			objParameterDefinitionDialog.Move 0,0
			wait 2
			objParameterDefinitionDialog.Resize width-5,hieght-5
			wait 2
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		else
			Call Fn_Button_Click("Fn_Mech_ParameterDefinationBasicCreate", objParameterDefinitionDialog, StrButtonName)
		End If
	End If
	'releasing object of [ NewParameterDefinition ] dialog
	set objParameterDefinitionDialog=nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_Mech_AdditionalParameterDefInfo

'Description			 :	Function Used to enter Additional Parameter  Defenation information

'Parameters			   :   1.dicAdditionalParameterDefInfo: Additional Parameter  Defenation information
'
'Return Value		   : 	True or False

'Pre-requisite			:	Parameter Defination dialog should be appear on screen on user should be on Define Parameter Defination Information dialog

'Examples				:   dicAdditionalParameterDefInfo("Comment")="Comment1"
'										dicAdditionalParameterDefInfo("ParameterType")="Calibration"
'										dicAdditionalParameterDefInfo("ButtonName")="Next"
'										bReturn=Fn_Mech_AdditionalParameterDefInfo(dicAdditionalParameterDefInfo)
'
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												20-Apr-2012								1.0																						Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_Mech_AdditionalParameterDefInfo(dicAdditionalParameterDefInfo)
	GBL_FAILED_FUNCTION_NAME="Fn_Mech_AdditionalParameterDefInfo"
 	'variable declaration
	Dim objParameterDefinitionDialog,objTable,objChild

	Fn_Mech_AdditionalParameterDefInfo=False
	If Fn_SISW_UI_Object_Operations("Fn_Mech_AdditionalParameterDefInfo","Exist",Window("MechatronicsWindow").JavaDialog("NewParameterDefinition"),SISW_MIN_TIMEOUT) = False Then
'	if not Window("MechatronicsWindow").JavaDialog("NewParameterDefinition").Exist(6) then
		Exit function
	else
		'Creating object of [ NewParameterDefinition ] dialog
		set objParameterDefinitionDialog=Window("MechatronicsWindow").JavaDialog("NewParameterDefinition")
	end if
	'Setting comment
	If dicAdditionalParameterDefInfo("Comment")<>"" Then
		Call Fn_Edit_Box("Fn_Mech_AdditionalParameterDefInfo",objParameterDefinitionDialog,"Comment",dicAdditionalParameterDefInfo("Comment"))
	End If
	'Selecting Parameter Type
	If dicAdditionalParameterDefInfo("ParameterType")<>"" Then
		objParameterDefinitionDialog.JavaStaticText("ParameterDef_text").SetTOProperty "label","Parameter Type:"
		objParameterDefinitionDialog.JavaButton("ParameterDef_DropDown").Click
		wait 2
		Set objTable=Description.Create()
		objTable("Class Name").value="JavaTable"
        objTable("toolkit class").value="com.teamcenter.rac.common.lov.view.components.LOVTreeTable"
		'objTable("tagname").value="LOVTreeTable"
		Set objChild=objParameterDefinitionDialog.ChildObjects(objTable)
		For iCounter=0 to objChild(0).GetROProperty("rows")
			If trim(dicAdditionalParameterDefInfo("ParameterType"))=trim(objChild(0).Object.getValueAt(iCounter,0).getDisplayableValue()) Then
				objChild(0).DoubleClickCell iCounter,0
				Exit for
			End If
		Next
		Set objTable=Nothing
		Set objChild=Nothing
'		Call Fn_Edit_Box("Fn_Mech_AdditionalParameterDefInfo",objParameterDefinitionDialog,"ParameterType",dicAdditionalParameterDefInfo("ParameterType"))
		End If
	'Clicking on button
	If dicAdditionalParameterDefInfo("ButtonName")<>"" Then
'		Call Fn_Button_Click("Fn_Mech_AdditionalParameterDefInfo", objParameterDefinitionDialog,dicAdditionalParameterDefInfo("ButtonName"))
		Call Fn_SISW_UI_JavaButton_Operations("Fn_Mech_AdditionalParameterDefInfo","Click", objParameterDefinitionDialog,dicAdditionalParameterDefInfo("ButtonName"))
	End If
	If Err.Number < 0 Then
		Fn_Mech_AdditionalParameterDefInfo=False
	Else
		Fn_Mech_AdditionalParameterDefInfo=True
	End If
	set objParameterDefinitionDialog=nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_Mech_ParameterDefinationRevisionInfoOperations

'Description			 :	Function Used to perform operations on Parameter Defination Revision Information

'Parameters			   :   1.StrAction: Action name
'										2.dicRevisionInfo: Parameter Defination Revision Information
'
'Return Value		   : 	True or False

'Pre-requisite			:	Parameter Defination dialog should be appear on screen on user should be on Define Parameter Defination Revision Information dialog

'Examples				:   Dim dicRevisionInfo
'										Set dicRevisionInfo = CreateObject("Scripting.Dictionary")
'										With dicRevisionInfo
'											.Add "Comment",""
'											.Add "ParameterDefinitionDescriptor",""
'										End With
'										dicRevisionInfo("Comment")="Parameter Def revision info comment"
'										dicRevisionInfo("ParameterDefinitionDescriptor")="Descriptor"
'										bReturn= Fn_Mech_ParameterDefinationRevisionInfoOperations("ParameterDefinationRevisionGeneralInfo",dicRevisionInfo)
'
'										dicRevisionInfo.RemoveAll
'										With dicRevisionInfo
'											.Add "RowCount",""
'											.Add "ColumnCount",""
'											.Add "RowLabels",""
'											.Add "ColumnLabels",""
'										End With
'										dicRevisionInfo("RowCount")=3
'										dicRevisionInfo("ColumnCount")=3
'										dicRevisionInfo("RowLabels")="Row1~Row2~Row3"
'										dicRevisionInfo("ColumnLabels")="Column1~Column2~Column3"
'										bReturn= Fn_Mech_ParameterDefinationRevisionInfoOperations("TableDefination",dicRevisionInfo)
'
'										dicRevisionInfo.RemoveAll
'										With dicRevisionInfo
'											.Add "ColumnNumber",""
'											.Add "RowNumber",""
'											.Add "Values",""
'										End With
'										
'										dicRevisionInfo("ColumnNumber")="1~2~3~1~2~3~1~2~3"
'										dicRevisionInfo("RowNumber")="1~1~1~2~2~2~3~3~3"
'										dicRevisionInfo("Values")="150~90~268~28~36~42~52~68~88"/dicRevisionInfo("Values")="10:Ten~20:Twenty~30:Thirty~40:Four"
'										bReturn= Fn_Mech_ParameterDefinationRevisionInfoOperations("MaximunValues",dicRevisionInfo)

'										dicRevisionInfo.RemoveAll
'										With dicRevisionInfo
'											.Add "ColumnNumber",""
'											.Add "RowNumber",""
'											.Add "Values",""
'										End With
'										
'										dicRevisionInfo("ColumnNumber")="1~2~3~1~2~3~1~2~3"
'										dicRevisionInfo("RowNumber")="1~1~1~2~2~2~3~3~3"
'										dicRevisionInfo("Values")="14-5-2012~15-5-2012~16-5-2012~17-5-2012~18-5-2012~19-5-2012~20-5-2012~21-5-2012~22-5-2012"
'										bReturn= Fn_Mech_ParameterDefinationRevisionInfoOperations("DateMaximunValues",dicRevisionInfo)
'
'
'										dicRevisionInfo.RemoveAll
'										With dicRevisionInfo
'											.Add "ColumnNumber",""
'											.Add "RowNumber",""
'											.Add "Values",""
'										End With
'										
'										dicRevisionInfo("ColumnNumber")="1~2~3~1~2~3~1~2~3"
'										dicRevisionInfo("RowNumber")="1~1~1~2~2~2~3~3~3"
'										dicRevisionInfo("Values")="135~65~215~22~32~41~51~65~81"/dicRevisionInfo("Values")="10:Ten~20:Twenty~30:Thirty~40:Four"
'										bReturn= Fn_Mech_ParameterDefinationRevisionInfoOperations("InitialValues",dicRevisionInfo)
'
'										dicRevisionInfo.RemoveAll
'										With dicRevisionInfo
'											.Add "ColumnNumber",""
'											.Add "RowNumber",""
'											.Add "Values",""
'										End With
'										
'										dicRevisionInfo("ColumnNumber")="1~2~3~1~2~3~1~2~3"
'										dicRevisionInfo("RowNumber")="1~1~1~2~2~2~3~3~3"
'										dicRevisionInfo("Values")="13-5-2012~14-5-2012~15-5-2012~16-5-2012~17-5-2012~18-5-2012~19-5-2012~20-5-2012~21-5-2012"
'										bReturn= Fn_Mech_ParameterDefinationRevisionInfoOperations("DateInitialValues",dicRevisionInfo)
'
'
'										dicRevisionInfo.RemoveAll
'										With dicRevisionInfo
'											.Add "ColumnNumber",""
'											.Add "RowNumber",""
'											.Add "Values",""
'										End With
'										
'										dicRevisionInfo("ColumnNumber")="1~2~3~1~2~3~1~2~3"
'										dicRevisionInfo("RowNumber")="1~1~1~2~2~2~3~3~3"
'										dicRevisionInfo("Values")="120~56~212~15~26~38~48~62~78"/dicRevisionInfo("Values")="10:Ten~20:Twenty~30:Thirty~40:Four"
'										bReturn= Fn_Mech_ParameterDefinationRevisionInfoOperations("MinimunValues",dicRevisionInfo)
'
'										dicRevisionInfo.RemoveAll
'										With dicRevisionInfo
'											.Add "ColumnNumber",""
'											.Add "RowNumber",""
'											.Add "Values",""
'										End With
'										
'										dicRevisionInfo("ColumnNumber")="1~2~3~1~2~3~1~2~3"
'										dicRevisionInfo("RowNumber")="1~1~1~2~2~2~3~3~3"
'										dicRevisionInfo("Values")="13-5-2012~14-5-2012~15-5-2012~16-5-2012~17-5-2012~18-5-2012~19-5-2012~20-5-2012~21-5-2012"
'										bReturn= Fn_Mech_ParameterDefinationRevisionInfoOperations("DateMinimunValues",dicRevisionInfo)
'
'
'										dicRevisionInfo.RemoveAll
'										With dicRevisionInfo
'											.Add "ConversionName",""
'											.Add "ConversionDescription",""
'											.Add "ConversionType",""
'										End With
'										
'										dicRevisionInfo("ConversionName")="Rule1"
'										dicRevisionInfo("ConversionDescription")="Rule 1 Description"
'										dicRevisionInfo("ConversionType")="Linear"
'										bReturn= Fn_Mech_ParameterDefinationRevisionInfoOperations("ConversionRule",dicRevisionInfo)
'
'										dicRevisionInfo.RemoveAll
'										With dicRevisionInfo
'											.Add "ConstantName",""
'											.Add "ConstantColName",""
'											.Add "ConstantValue",""
'										End With
'										
'										dicRevisionInfo("ConstantName")="A~B"
'										dicRevisionInfo("ConstantColName")="Constant Value"
'										dicRevisionInfo("ConstantValue")="10~20"
'										bReturn= Fn_Mech_ParameterDefinationRevisionInfoOperations("ConstantsTable",dicRevisionInfo)
'
'										dicRevisionInfo.RemoveAll
'										With dicRevisionInfo
'											.Add "ButtonName",""
'										End With
'										
'										dicRevisionInfo("ButtonName")="Finish"
'										bReturn= Fn_Mech_ParameterDefinationRevisionInfoOperations("ClickButton",dicRevisionInfo)
'										
'										dicRevisionInfo.RemoveAll
'										With dicRevisionInfo
'											.Add "ButtonName",""
'										End With
'										
'										dicRevisionInfo("ButtonName")="Close"
'										bReturn= Fn_Mech_ParameterDefinationRevisionInfoOperations("ClickButton",dicRevisionInfo)
'
'								dicRevisionInfo("ButtonName")="Finish"
'								dicRevisionInfo("PropertyState")="enabled"
'								bReturn=Fn_Mech_ParameterDefinationRevisionInfoOperations("Button_PropertyState",dicRevisionInfo)
'
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												20-Apr-2012								1.0																						Sunny R
'													Sandeep N												20-Apr-2012								1.1																						Sunny R
'																																													Added Cases : DateMaximunValues,DateMinimunValues,DateInitialValues
'													Sandeep N												08-Aug-2012								1.2					added case : Button_PropertyState		Anjali M
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_Mech_ParameterDefinationRevisionInfoOperations(StrAction,dicRevisionInfo)
		GBL_FAILED_FUNCTION_NAME="Fn_Mech_ParameterDefinationRevisionInfoOperations"
		'variable declaration
		Dim objDialog,bFlag
		Fn_Mech_ParameterDefinationRevisionInfoOperations=false
	   'checking existance of [ NewParameterDefinition ] dialog
	   if not Window("MechatronicsWindow").JavaDialog("NewParameterDefinition").Exist(6) then
			Exit function
		else
			'Creating object of [ NewParameterDefinition ] dialog
			set objDialog=Window("MechatronicsWindow").JavaDialog("NewParameterDefinition")
		end if
        Select Case StrAction
				'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
				'case to enter general parameter defination information
				Case "ParameterDefinationRevisionGeneralInfo"
					dicParameterDefRevGeneralInfo("Comment")=dicRevisionInfo("Comment")
					dicParameterDefRevGeneralInfo("ParameterDefinitionDescriptor")=dicRevisionInfo("ParameterDefinitionDescriptor")
					dicParameterDefRevGeneralInfo("SizeUnits")=dicRevisionInfo("SizeUnits")
					dicParameterDefRevGeneralInfo("ControlEngineer")=dicRevisionInfo("ControlEngineer")
					dicParameterDefRevGeneralInfo("Size")=dicRevisionInfo("Size")
					dicParameterDefRevGeneralInfo("IsSigned")=dicRevisionInfo("IsSigned")
					dicParameterDefRevGeneralInfo("ResolutionNumerator")=dicRevisionInfo("ResolutionNumerator")
					dicParameterDefRevGeneralInfo("ResolutionDenominator")=dicRevisionInfo("ResolutionDenominator")
					dicParameterDefRevGeneralInfo("Precision")=dicRevisionInfo("Precision")
					dicParameterDefRevGeneralInfo("Tolerance")=dicRevisionInfo("Tolerance")

					bFlag=Fn_Mech_ParameterDefRevGeneralInfo(dicParameterDefRevGeneralInfo)
					If bFlag=False Then
						set objDialog=nothing
						Exit function
					End If
				'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
				'case to enter general parameter defination information
				Case "ClickButton"
					bFlag=Fn_Button_Click("Fn_Mech_ParameterDefinationRevisionInfoOperations", objDialog, dicRevisionInfo("ButtonName"))
					If bFlag=False Then
						set objDialog=nothing
						Exit function
					End If
				'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
				'case to enter table defination
				Case "TableDefination"
					bFlag=Fn_Mech_TableDefination(dicRevisionInfo("RowCount"),dicRevisionInfo("ColumnCount"),dicRevisionInfo("RowLabels"),dicRevisionInfo("ColumnLabels"))
					If bFlag=False Then
						set objDialog=nothing
						Exit function
					End If
				'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
				'case to enter values in Maximum table
				Case "MaximunValues","DateMaximunValues"
					Select Case StrAction
							Case "MaximunValues"
								bFlag=Fn_Mech_MaximumValueTable("SetData",dicRevisionInfo("ColumnNumber"),dicRevisionInfo("RowNumber"),dicRevisionInfo("Values"))
							Case "DateMaximunValues"
								bFlag=Fn_Mech_MaximumValueTable("SetDateData",dicRevisionInfo("ColumnNumber"),dicRevisionInfo("RowNumber"),dicRevisionInfo("Values"))
					End Select
						
					If bFlag=False Then
						set objDialog=nothing
						Exit function
					End If
				'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
				'case to enter values in Minimum table
				Case "MinimunValues","DateMinimunValues"
					Select Case StrAction
						Case "MinimunValues"
							bFlag=Fn_Mech_MinimumValueTable("SetData",dicRevisionInfo("ColumnNumber"),dicRevisionInfo("RowNumber"),dicRevisionInfo("Values"))
						Case "DateMinimunValues"
							bFlag=Fn_Mech_MinimumValueTable("SetDateData",dicRevisionInfo("ColumnNumber"),dicRevisionInfo("RowNumber"),dicRevisionInfo("Values"))
					End Select
					If bFlag=False Then
						set objDialog=nothing
						Exit function
					End If
				'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
				'case to enter values in Initial table
				Case "InitialValues","DateInitialValues"
					Select Case StrAction
						Case "InitialValues"
							bFlag=Fn_Mech_InitialValueTable("SetData",dicRevisionInfo("ColumnNumber"),dicRevisionInfo("RowNumber"),dicRevisionInfo("Values"))
						Case "DateInitialValues"
							bFlag=Fn_Mech_InitialValueTable("SetDateData",dicRevisionInfo("ColumnNumber"),dicRevisionInfo("RowNumber"),dicRevisionInfo("Values"))
					End Select
					If bFlag=False Then
						set objDialog=nothing
						Exit function
					End If
				'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
				'case for conversion rule
				Case "ConversionRule"
					bFlag=Fn_Mech_ConversionRule(dicRevisionInfo("ConversionAction"),dicRevisionInfo("ConversionName"),dicRevisionInfo("ConversionDescription"),dicRevisionInfo("ConversionType"),dicRevisionInfo("ConversionExpression"))						
					If bFlag=False Then
						set objDialog=nothing
						Exit function
					End If
				'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
				'case for conversion rule
				Case "ConstantsTable"
					bFlag=Fn_Mech_ConstantsTableOperations("SetCellData",dicRevisionInfo("ConstantName"),dicRevisionInfo("ConstantColName"),dicRevisionInfo("ConstantValue"))
					If bFlag=False Then
						set objDialog=nothing
						Exit function
					End If
	            ' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				'Case to get specific property state of specific Button : e.g { Label , enabled state }
				Case "Button_PropertyState"
					If objDialog.JavaButton(dicRevisionInfo("ButtonName")).Exist(3) Then
						Fn_Mech_ParameterDefinationRevisionInfoOperations=objDialog.JavaButton(dicRevisionInfo("ButtonName")).GetROProperty(dicRevisionInfo("PropertyState"))
					else
						Fn_Mech_ParameterDefinationRevisionInfoOperations=false
					End if
					set objDialog=nothing
					Exit function
				Case else
					set objDialog=nothing
					Exit function
	   End Select
	   set objDialog=nothing
	   Fn_Mech_ParameterDefinationRevisionInfoOperations=true
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_Mech_ParameterDefRevGeneralInfo

'Description			 :	Function Used to enter general information for Parameter Defination Revision

'Parameters			   :   1.dicParameterDefRevGeneralInfo: Parameter Defination Revision general information
'
'Return Value		   : 	True or False

'Pre-requisite			:	Parameter Defination dialog should be appear on screen on user should be on Define Parameter Defination Revision Information dialog

'Examples				:  	dicParameterDefRevGeneralInfo("Comment")="Comment1"
'										dicParameterDefRevGeneralInfo("ParameterDefinitionDescriptor")="Def1"
'										dicParameterDefRevGeneralInfo("SizeUnits")="Bit\(s\)"
'										bReturn=Fn_Mech_ParameterDefRevGeneralInfo(dicParameterDefRevGeneralInfo)
'
'										dicParameterDefRevGeneralInfo("IsSigned")="True"
'										dicParameterDefRevGeneralInfo("ResolutionNumerator")="1"
'										dicParameterDefRevGeneralInfo("Precision")="2"
'										dicParameterDefRevGeneralInfo("Tolerance")="1"
'										bReturn=Fn_Mech_ParameterDefRevGeneralInfo(dicParameterDefRevGeneralInfo)
'
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												20-Apr-2012								1.0																						Sunny R
'													Sandeep N												26-Apr-2012								1.1																						Sunny R
'		Snehal Salunkhe		10-Dec-2012		1.1				Koustubh W			Modified code to set values from Dropdown / TableCombo. for TC 10.1
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_Mech_ParameterDefRevGeneralInfo(dicParameterDefRevGeneralInfo)
	GBL_FAILED_FUNCTION_NAME="Fn_Mech_ParameterDefRevGeneralInfo"
	Dim objParameterDefinitionDialog,objStaticText,objChild,scrollMax
	Dim WshShell,i
	Dim iCounter,objTable

	Fn_Mech_ParameterDefRevGeneralInfo=false
	If Fn_SISW_UI_Object_Operations("Fn_Mech_ParameterDefRevGeneralInfo","Exist",Window("MechatronicsWindow").JavaDialog("NewParameterDefinition"),SISW_MIN_TIMEOUT) = False Then
'	If not Window("MechatronicsWindow").JavaDialog("NewParameterDefinition").Exist(6) Then
		Exit function
	Else
		set objParameterDefinitionDialog=Window("MechatronicsWindow").JavaDialog("NewParameterDefinition")
	End If
	If dicParameterDefRevGeneralInfo("Comment")<>"" Then
		'Setting comment
		Call Fn_Edit_Box("Fn_Mech_ParameterDefRevGeneralInfo",objParameterDefinitionDialog,"Comment",dicParameterDefRevGeneralInfo("Comment"))
	End If
	If  dicParameterDefRevGeneralInfo("ControlEngineer")<>"" Then

		objParameterDefinitionDialog.JavaStaticText("ParameterDef_text").SetTOProperty "label","Control Engineer:"
        	
       ' Call Fn_Button_Click("Fn_Mech_ParameterDefRevGeneralInfo", objParameterDefinitionDialog, "ParameterDef_DropDown")
        Call Fn_SISW_UI_JavaButton_Operations("Fn_Mech_ParameterDefRevGeneralInfo","Click", objParameterDefinitionDialog,"ParameterDef_DropDown")
		wait 2
		
		For i=0 to 2
			set WshShell = CreateObject("WScript.Shell")
			wait 1
			WshShell.SendKeys "{TAB}"
			WshShell.SendKeys "^{END}"
			set WshShell =nothing
		Next
        
		Set objTable=Description.Create()
		objTable("Class Name").value="JavaTable"

		objTable("toolkit class").value="com\.teamcenter\.rac\.common\.lov\.view\.components\.LOVTreeTable"
        objTable("displayed").value="1"
		objTable("enabled").value="1"
		objTable("focused").value="1"
		Set objChild=objParameterDefinitionDialog.ChildObjects(objTable)


		For iCounter=0 to objChild(0).GetROProperty("rows")

			If trim(dicParameterDefRevGeneralInfo("ControlEngineer"))=trim(objChild(0).Object.getValueAt(iCounter,0).getDisplayableValue()) Then
				objChild(0).DoubleClickCell iCounter,0
				Exit for
			End If
		Next
		Set objTable=Nothing
		Set objChild=Nothing
	End If
	wait(7)
	If dicParameterDefRevGeneralInfo("ParameterDefinitionDescriptor")<>"" Then
		'Setting Parameter Definition Descriptor
		objParameterDefinitionDialog.JavaStaticText("PropertyLabel").SetTOProperty "label","Parameter Definition Descriptor:"
		Call Fn_Edit_Box("Fn_Mech_ParameterDefRevGeneralInfo",objParameterDefinitionDialog,"ParameterDefinitionDescriptor",dicParameterDefRevGeneralInfo("ParameterDefinitionDescriptor"))
	End If
	If  dicParameterDefRevGeneralInfo("SizeUnits")<>"" Then

		dicParameterDefRevGeneralInfo("SizeUnits")=Replace(dicParameterDefRevGeneralInfo("SizeUnits"),"\","")
		objParameterDefinitionDialog.JavaStaticText("ParameterDef_text").SetTOProperty "label","Size Units:"
        Call Fn_Button_Click("Fn_Mech_ParameterDefRevGeneralInfo", objParameterDefinitionDialog, "ParameterDef_DropDown")
		wait 2
		Set objTable=Description.Create()
		objTable("Class Name").value="JavaTable"

		objTable("toolkit class").value="com\.teamcenter\.rac\.common\.lov\.view\.components\.LOVTreeTable"
        objTable("displayed").value="1"
		objTable("enabled").value="1"
		objTable("focused").value="1"
		Set objChild=objParameterDefinitionDialog.ChildObjects(objTable)
		For iCounter=0 to objChild(0).GetROProperty("rows")

			If trim(dicParameterDefRevGeneralInfo("SizeUnits"))=trim(objChild(0).Object.getValueAt(iCounter,0).getDisplayableValue()) Then
				objChild(0).DoubleClickCell iCounter,0
				Exit for
			End If
		Next
		Set objTable=Nothing
		Set objChild=Nothing
	End If
	If dicParameterDefRevGeneralInfo("Size")<>"" Then
		'Setting Size
		Call Fn_Edit_Box("Fn_Mech_ParameterDefRevGeneralInfo",objParameterDefinitionDialog,"Size",dicParameterDefRevGeneralInfo("Size"))
	End If
	If dicParameterDefRevGeneralInfo("SizeInByte")<>"" Then
		'Setting Size in Byte
		Call Fn_Edit_Box("Fn_Mech_ParameterDefRevGeneralInfo",objParameterDefinitionDialog,"SizeInByte",dicParameterDefRevGeneralInfo("SizeInByte")+ vbLf + "")
	End If
	
	If Fn_SISW_UI_Object_Operations("Fn_Mech_ParameterDefRevGeneralInfo","Exist",objParameterDefinitionDialog.JavaSlider("JScrollPane"),SISW_MINLESS_TIMEOUT) = True Then
'	If objParameterDefinitionDialog.JavaSlider("JScrollPane").Exist(3) Then
		scrollMax=objParameterDefinitionDialog.JavaSlider("JScrollPane").GetROProperty("max")
		objParameterDefinitionDialog.JavaSlider("JScrollPane").Drag scrollMax
		wait 1
	End If
	If dicParameterDefRevGeneralInfo("IsSigned")<>"" Then
		'Setting [ Is Signed ] option
		objParameterDefinitionDialog.JavaRadioButton("IsSigned").SetTOProperty "attached text",dicParameterDefRevGeneralInfo("IsSigned")
		Call Fn_UI_JavaRadioButton_SetON("Fn_Mech_ParameterDefRevGeneralInfo",objParameterDefinitionDialog, "IsSigned")
	End If
	If dicParameterDefRevGeneralInfo("ResolutionNumerator")<>"" Then
		'Setting Resolution Numerator
'		Call Fn_Edit_Box("Fn_Mech_ParameterDefRevGeneralInfo",objParameterDefinitionDialog,"ResolutionNumerator",dicParameterDefRevGeneralInfo("ResolutionNumerator")+ vbTab + "")
 	    objParameterDefinitionDialog.JavaEdit("ResolutionNumerator").set dicParameterDefRevGeneralInfo("ResolutionNumerator")+ vbTab + ""
	End If
	If dicParameterDefRevGeneralInfo("ResolutionDenominator")<>"" Then
		'Setting Resolution Denominator
		Call Fn_Edit_Box("Fn_Mech_ParameterDefRevGeneralInfo",objParameterDefinitionDialog,"ResolutionDenominator",dicParameterDefRevGeneralInfo("ResolutionDenominator")+ vbTab + "")
	End If
	If dicParameterDefRevGeneralInfo("Precision")<>"" Then
		'Setting Precision
		Call Fn_Edit_Box("Fn_Mech_ParameterDefRevGeneralInfo",objParameterDefinitionDialog,"Precision",dicParameterDefRevGeneralInfo("Precision"))
	End If
	If dicParameterDefRevGeneralInfo("Tolerance")<>"" Then
		'Setting Tolerance
		Call Fn_Edit_Box("Fn_Mech_ParameterDefRevGeneralInfo",objParameterDefinitionDialog,"Tolerance",dicParameterDefRevGeneralInfo("Tolerance"))
	End If
	Fn_Mech_ParameterDefRevGeneralInfo=true
	set objParameterDefinitionDialog=nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_Mech_TableDefination

'Description			 :	Function Used to Define Table Defination of Parameter Defination Object

'Parameters			   :   1.iRows: Number of rows
'										2.iColumns: Number of columns
'										3.StrRowLabels: Row Labels
'										4.StrColumnLabels : Column Labels
'
'Return Value		   : 	True or False

'Pre-requisite			:	Parameter Defination dialog should be appear on screen on user should be on Define Parameter Defination Revision Information dialog

'Examples				:   bReturn=Fn_Mech_TableDefination(3,3,"R1~R2~R3","C1~C2~C3")
'
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												20-Apr-2012								1.0																						Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_Mech_TableDefination(iRows,iColumns,StrRowLabels,StrColumnLabels)
	GBL_FAILED_FUNCTION_NAME="Fn_Mech_TableDefination"
 	'Variable declaration
	Dim scrollMax,arrRowLabels,iCounter,arrColLabels,objParameterDefinitionDialog

	Fn_Mech_TableDefination=false
	If Fn_SISW_UI_Object_Operations("Fn_Mech_TableDefination","Exist",Window("MechatronicsWindow").JavaDialog("NewParameterDefinition"),SISW_MIN_TIMEOUT) = False Then
'	if not Window("MechatronicsWindow").JavaDialog("NewParameterDefinition").Exist(6) then
		Exit function
	else
		set objParameterDefinitionDialog=Window("MechatronicsWindow").JavaDialog("NewParameterDefinition")
	end if
	If Fn_SISW_UI_Object_Operations("Fn_Mech_TableDefination","Exist",objParameterDefinitionDialog.JavaSlider("JScrollPane"),SISW_MINLESS_TIMEOUT) = True Then
'	If objParameterDefinitionDialog.JavaSlider("JScrollPane").Exist(3) Then
		'Scrolling till the end of panel
		scrollMax=objParameterDefinitionDialog.JavaSlider("JScrollPane").GetROProperty("max")
		objParameterDefinitionDialog.JavaSlider("JScrollPane").Drag scrollMax
		wait 1
	End If

	If iRows<>"" Then
		'Setting table rows
		'Call Fn_Edit_Box("Fn_Mech_TableDefination",objParameterDefinitionDialog,"Rows",iRows+ vbLf + "")
		objParameterDefinitionDialog.JavaEdit("Rows").Set iRows+ vbLf + ""
		wait 1
	End If
	If iColumns<>"" Then
		'Setting table columns
		'Call Fn_Edit_Box("Fn_Mech_TableDefination",objParameterDefinitionDialog,"Columns",iColumns+ vbLf + "")
		objParameterDefinitionDialog.JavaEdit("Columns").set iColumns+ vbLf + ""
		wait 1
	End If
	
	If StrRowLabels<>"" Then
		'Setting row labels
		arrRowLabels=Split(StrRowLabels,"~")
		Call Fn_CheckBox_Set("Fn_Mech_TableDefination", objParameterDefinitionDialog, "RowLabels","on")
		For iCounter=0 to ubound(arrRowLabels)
			'Call Fn_Edit_Box("Fn_Mech_TableDefination",objParameterDefinitionDialog,"Labels",arrRowLabels(iCounter))
			objParameterDefinitionDialog.JavaEdit("Labels").set arrRowLabels(iCounter)
			Call Fn_Button_Click("Fn_Mech_TableDefination", objParameterDefinitionDialog, "AddLabels")
		Next
		Call Fn_CheckBox_Set("Fn_Mech_TableDefination", objParameterDefinitionDialog, "RowLabels","off")
	End If
	If StrColumnLabels<>"" Then
		'Setting column labels
		arrColLabels=Split(StrColumnLabels,"~")
		Call Fn_CheckBox_Set("Fn_Mech_TableDefination", objParameterDefinitionDialog, "ColumnLabels","on")
		For iCounter=0 to ubound(arrColLabels)
			'Call Fn_Edit_Box("Fn_Mech_TableDefination",objParameterDefinitionDialog,"Labels",arrColLabels(iCounter))
			objParameterDefinitionDialog.JavaEdit("Labels").set arrColLabels(iCounter)
			Call Fn_Button_Click("Fn_Mech_TableDefination", objParameterDefinitionDialog, "AddLabels")
		Next
		Call Fn_CheckBox_Set("Fn_Mech_TableDefination", objParameterDefinitionDialog, "ColumnLabels","off")
	End If
	Fn_Mech_TableDefination=true
	set objParameterDefinitionDialog=nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_Mech_ConstantsTableOperations

'Description			 :	Function Used to perform operations of Constants Table of parameter Defination Revision Information

'Parameters			   :   1.StrAction: Action Name
'										2.StrConstantName: Constant Name
'										3.StrColName: Column Name
'										4.StrValue : Cell value
'
'Return Value		   : 	True or False

'Pre-requisite			:	Parameter Defination dialog should be appear on screen on user should be on Define Parameter Defination Revision Information dialog

'Examples				:   bReturn=Fn_Mech_ConstantsTableOperations("SetCellData","B~C~D","Constant Value","8~9~10")
'
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												20-Apr-2012								1.0																						Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_Mech_ConstantsTableOperations(StrAction,StrConstantName,StrColName,StrValue)
	GBL_FAILED_FUNCTION_NAME="Fn_Mech_ConstantsTableOperations"
 	'declaring variables
	Dim arrConstantName,arrValue,iCounter,iRowCount,iCount,bFlag,crrConstantName,objParameterDefinitionDialog

	Fn_Mech_ConstantsTableOperations=false
	if not Window("MechatronicsWindow").JavaDialog("NewParameterDefinition").Exist(6) then
		Exit function
	else
		set objParameterDefinitionDialog=Window("MechatronicsWindow").JavaDialog("NewParameterDefinition")
	end if
	Select Case StrAction
		'Case to set cell data against respective [ Constant Name ]
		Case "SetCellData"
				arrConstantName=Split(StrConstantName,"~")
				arrValue=Split(StrValue,"~")
				For iCounter=0 to ubound(arrConstantName)
					iRowCount=objParameterDefinitionDialog.JavaTable("Constants").GetROProperty("rows")
					For iCount=0 to iRowCount-1
						bFlag=false
						crrConstantName=objParameterDefinitionDialog.JavaTable("Constants").GetCellData(iCount,"Constant Name")
						If trim(crrConstantName)=trim(arrConstantName(iCounter)) Then
							If StrColName="" Then
								StrColName="Constant Value"
							End If
							objParameterDefinitionDialog.JavaTable("Constants").ClickCell iCount,StrColName,"LEFT"
							wait 10
							Call Fn_Edit_Box("Fn_Mech_TableDefination",objParameterDefinitionDialog,"ConversionRuleConstants",arrValue(iCounter))
							bFlag=true
							Exit for
						End If
					Next
					If bFlag=false Then
						set objParameterDefinitionDialog=nothing
						Exit function
					End If
				Next
				Fn_Mech_ConstantsTableOperations=true
	End Select
	set objParameterDefinitionDialog=nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_Mech_ConversionRule

'Description			 :	Function Used to define Conversion Rule of parameter Defination Revision Information

'Parameters			   :   1.StrAction: Action Name
'										2.StrName: Rule name
'										3.StrDescription: Description
'										4.StrType : Conversion rule Type
'										5.StrExpression : Conversion rule expression
'
'Return Value		   : 	True or False

'Pre-requisite			:	Parameter Defination dialog should be appear on screen on user should be on Define Parameter Defination Revision Information dialog

'Examples				:   bReturn=Fn_Mech_ConversionRule("","CR1","Conversion rule one","Quadratic","")
'
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												20-Apr-2012								1.0																						Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_Mech_ConversionRule(StrAction,StrName,StrDescription,StrType,StrExpression)
	GBL_FAILED_FUNCTION_NAME="Fn_Mech_ConversionRule"
 	'declaring variables
	Dim objParameterDefinitionDialog,scrollMax
	Fn_Mech_ConversionRule=false
	'checking existance of [ NewParameterDefinition ] dialog
	if not Window("MechatronicsWindow").JavaDialog("NewParameterDefinition").Exist(6) then
		Exit function
	else
		'Creating object of [ NewParameterDefinition ] dialog
		set objParameterDefinitionDialog=Window("MechatronicsWindow").JavaDialog("NewParameterDefinition")
	end if
	If objParameterDefinitionDialog.JavaSlider("JScrollPane").Exist(3) Then
		'Scrolling till the end of panel
		scrollMax=objParameterDefinitionDialog.JavaSlider("JScrollPane").GetROProperty("max")
		objParameterDefinitionDialog.JavaSlider("JScrollPane").Drag scrollMax
	End If
	If StrAction<>"" Then
        If objParameterDefinitionDialog.JavaStaticText("ActionLinkDownButton").Exist(1) Then
			objParameterDefinitionDialog.JavaStaticText("ActionLinkDownButton").Click 1,1
		ElseIf objParameterDefinitionDialog.JavaObject("ActionLinkDownButton").Exist(1) Then			
			objParameterDefinitionDialog.JavaObject("ActionLinkDownButton").Click 1,1
		End If        
		wait 1
        objParameterDefinitionDialog.JavaMenu("index:=0","label:="&StrAction).Select
	End If
	'setting conversion rule name
	If StrName<>"" Then
		Call Fn_Edit_Box("Fn_Mech_ConversionRule",objParameterDefinitionDialog,"ConversionRuleName",StrName)
	End If
	'setting conversion rule Description
	If StrDescription<>"" Then
		Call Fn_Edit_Box("Fn_Mech_ConversionRule",objParameterDefinitionDialog,"ConversionRuleDescription",StrDescription)
	End If
	'selecting Conversion rule type
	If StrType<>"" Then
		If Fn_UI_ListItemExist("Fn_Mech_ConversionRule", objParameterDefinitionDialog, "ConversionRuleType",StrType) Then
			Call Fn_List_Select("Fn_Mech_ConversionRule", objParameterDefinitionDialog,"ConversionRuleType",StrType)
		Else
			'releasing object of [ NewParameterDefinition ] dialog
			set objParameterDefinitionDialog=false
			Exit function
		End If
	End If
    If Err.Number < 0 Then
		Fn_Mech_ConversionRule=False
	Else
		Fn_Mech_ConversionRule=True
	End If
'	Fn_Mech_ConversionRule=true
	'releasing object of [ NewParameterDefinition ] dialog
	set objParameterDefinitionDialog=nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_Mech_MaximumValueTable

'Description			 :	Function Used to perform operation on MaximumValueTable of parameter Defination Revision Information

'Parameters			   :   1.StrAction: Action Name
'										2.iCol: column number
'										3.iRow: row number
'										4.StrValue : values
'
'Return Value		   : 	True or False

'Pre-requisite			:	Parameter Defination dialog should be appear on screen on user should be on Define Parameter Defination Revision Information dialog

'Examples				:   bReturn=Fn_Mech_MaximumValueTable("SetData","1~2~3~1~2~3~1~2~3","1~1~1~2~2~2~3~3~3","1~2~3~4~5~6~7~8~9")
'										bReturn=Fn_Mech_MaximumValueTable("SetDateData","1~2~3~1~2~3~1~2~3","1~1~1~2~2~2~3~3~3","14-5-2012~15-5-2012~16-5-2012~17-5-2012~18-5-2012~19-5-2012~20-5-2012~21-5-2012~22-5-2012")
'										bReturn=Fn_Mech_MaximumValueTable("SetData","1~2~1~2","1~1~2~2","10:Ten~20:Twenty~30:Thirty~40:Four")
'										"10:Ten~20:Twenty~30:Thirty~40:Four"="Value:Value Description~Value:Value Description"
'
'										bReturn=Fn_Mech_MaximumValueTable("VerifyCellData","1~2~1~2","1~1~2~2","100:Hundred~200:Two Hundred~300:Three Hundred~400:Four Hundred")
'										bReturn=Fn_Mech_MaximumValueTable("RowColumnCount","","","")
'										bReturn=Fn_Mech_MaximumValueTable("PasteData","","1","")
'										bReturn=Fn_Mech_MaximumValueTable("PasteData","2","1","")
'										bReturn=Fn_Mech_MaximumValueTable("PasteData_Keyboard","","","")
'										bReturn=Fn_Mech_MaximumValueTable("PasteData_Keyboard","2","2","")
'										bReturn=Fn_Mech_MaximumValueTable("GetAllContextMenu","","","")
'										bReturn=Fn_Mech_MaximumValueTable("GetAllContextMenu","","2","")
'										bReturn=Fn_Mech_MaximumValueTable("GetAllColumnNames","","2","")
'
'										IMPNote : For Cases PasteData, PasteData_Keyboard use [ StrValue ] parameter to set Show Value Description Option
'
'										bReturn=Fn_Mech_MaximumValueTable("Collapse_SetON","","","")
'										bReturn=Fn_Mech_MaximumValueTable("Collapse_SetOFF","","","")
'										bReturn=Fn_Mech_MaximumValueTable("UndoChanges","","","")
'										bReturn=Fn_Mech_MaximumValueTable("GetAllDisableContextMenu","","","")
'										bReturn=Fn_Mech_MaximumValueTable("GetAllDisableContextMenu","","1","")
'										bReturn=Fn_Mech_MaximumValueTable("VerifyDateData","1","1","13-Jun-2012:Test")
'
'										bReturn=Fn_Mech_MaximumValueTable("PasteRowsCols","","1~2~3~4","")
'										bReturn=Fn_Mech_MaximumValueTable("PasteRowsCols","StartColumn~EndColumn","StartRow~EndRow","")
'										bReturn=Fn_Mech_MaximumValueTable("PasteRowsCols","1~2","1~2","")
'										bReturn=Fn_Mech_MaximumValueTable("PasteRowsCols","1~4","1~1","")
'
'										bReturn=Fn_Mech_MaximumValueTable("GetColumnName","1","","")
'										bReturn=Fn_Mech_MaximumValueTable("GetRowName","","1","")
'
'											bReturn=Fn_Mech_MaximumValueTable("IsCellCurrentValueCorrect","1~2~2","1~1~2","false~true~false")
'											bReturn=Fn_Mech_MaximumValueTable("VerifyHeaderForegroundColour","1","","Red")
'
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												20-Apr-2012								1.0																						Sunny R
'													Sandeep N												26-Apr-2012								1.1						added case SetDateData			   Sunny R
'													Sandeep N												24-May-2012							   1.2						added case VerifyCellData			   Sunny R
'													Sandeep N												05-Jun-2012							     1.3						added case RowColumnCount,PasteData,PasteData_Keyboard,GetAllContextMenu,GetAllColumnNames			   Sunny R
'													Sandeep N												07-Jun-2012							     1.4						added case Collapse_SetON,Collapse_SetOFF			   Sunny R
'													Sandeep N												12-Jun-2012							     1.5						added case UndoChanges,GetAllDisableContextMenu			   Sunny R
'													Sandeep N												13-Jun-2012							   1.6						added case VerifyDateData			   Sunny R
'													Sandeep N												02-Jul-2012							   1.6						added case PasteRowsCols			   Sonal P
'													Sandeep N												30-Jul-2012							   1.7						added GetColumnName & GetRowName			   Sonal P
'													Sandeep N												08-Aug-2012							   1.8						added cases IsCellCurrentValueCorrect & VerifyHeaderForegroundColour			   Anjali M
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_Mech_MaximumValueTable(StrAction,iCol,iRow,StrValue)
	GBL_FAILED_FUNCTION_NAME="Fn_Mech_MaximumValueTable"
 	'declaring variables
	Dim objParameterDefinitionDialog,objTable,objChld,objDate
	Dim aCol,aRow,aValue,scrollMax,iCounter,aDate,aValDesc,bFlag,cellval
	Dim iRows,iCols,objMenu,crrMenu,StrLabel
	Dim sColourCode,sColour

	Fn_Mech_MaximumValueTable=False
	If Fn_SISW_UI_Object_Operations("Fn_Mech_MaximumValueTable", "Exist", Window("MechatronicsWindow").JavaDialog("NewParameterDefinition"),SISW_MIN_TIMEOUT)=False Then 
'	if not Window("MechatronicsWindow").JavaDialog("NewParameterDefinition").Exist(6) then
		Exit function
	else
		'Creating object of [ NewParameterDefinition ] dialog
		set objParameterDefinitionDialog=Window("MechatronicsWindow").JavaDialog("NewParameterDefinition")
	end if
	If Fn_SISW_UI_Object_Operations("Fn_Mech_MaximumValueTable", "Exist", objParameterDefinitionDialog.JavaSlider("JScrollPane"),SISW_MINLESS_TIMEOUT)=True Then 
'	If objParameterDefinitionDialog.JavaSlider("JScrollPane").Exist(3) Then
		'Scrolling till the mid of panel
		scrollMax=objParameterDefinitionDialog.JavaSlider("JScrollPane").GetROProperty("max")
		objParameterDefinitionDialog.JavaSlider("JScrollPane").Drag scrollMax/2
	End If

'	iRowCount=objParameterDefinitionDialog.JavaTable("MaximumValues").GetROProperty("rows")
	Select Case StrAction
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			Case "SetData"
				aCol=Split(iCol,"~")
				aRow=Split(iRow,"~")
				aValue=Split(StrValue,"~")
				If inStr(1,StrValue,":") Then
					Call Fn_CheckBox_Set("Fn_Mech_MaximumValueTable", objParameterDefinitionDialog,"MaxValueDescriptionCell","on")
				End If

				objParameterDefinitionDialog.JavaTable("MaximumValues").SelectRow 0
				For iCounter=0 to UBound(aRow)
					objParameterDefinitionDialog.JavaTable("MaximumValues").ActivateCell CInt(aRow(iCounter))-1,CInt(aCol(iCounter))-1
					wait 1
					aValDesc=Split(aValue(iCounter),":")

					Set objTable=Description.Create
					objTable("Class Name").value="JavaTable"
					Set objChld=objParameterDefinitionDialog.JavaTable("MaximumValues").ChildObjects(objTable)
					If aValDesc(0)<>"" Then
						objChld(0).SetCellData 0,0,aValDesc(0)
					End If
					If uBound(aValDesc)=1 Then
						objChld(0).SetCellData 1,0,aValDesc(1)
					End If

					Set objChld=nothing
					Set objTable=nothing
				Next
				Fn_Mech_MaximumValueTable=true
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			Case "SetDateData"
				If objParameterDefinitionDialog.JavaSlider("JScrollPane").Exist(3) Then
					scrollMax=objParameterDefinitionDialog.JavaSlider("JScrollPane").GetROProperty("max")
					objParameterDefinitionDialog.JavaSlider("JScrollPane").Drag scrollMax
					wait 1
				End If
				aCol=Split(iCol,"~")
				aRow=Split(iRow,"~")
				aValue=Split(StrValue,"~")
				If inStr(1,StrValue,":") Then
					Call Fn_CheckBox_Set("Fn_Mech_MaximumValueTable", objParameterDefinitionDialog,"MaxValueDescriptionCell","on")
				End If
				For iCounter=0 to UBound(aRow)
					objParameterDefinitionDialog.JavaTable("MaximumValues").ClickCell CInt(aRow(iCounter))-1,CInt(aCol(iCounter))-1,"LEFT"
					objParameterDefinitionDialog.JavaTable("MaximumValues").DoubleClickCell CInt(aRow(iCounter))-1,CInt(aCol(iCounter))-1,"LEFT"
					wait 1
					If JavaDialog("SelectDate").Exist(2) then
						aValDesc=Split(aValue(iCounter),":")
						aDate=Split(aValDesc(0),"-")

						Set objDate=JavaDialog("SelectDate").JavaObject("CalendarPanel").Object.getTime()

						objDate.setYear(Cint(aDate(2))-1900)
						objDate.setMonth(Cint(aDate(1))-1)
						objDate.setDate(aDate(0))
                        						
						JavaDialog("SelectDate").JavaObject("CalendarPanel").Object.setTime(objDate)
						JavaDialog("SelectDate").JavaButton("Ok").Click

						If ubound(aValDesc)=1 Then
							objParameterDefinitionDialog.JavaTable("MaximumValues").Object.getValueAt(CInt(aRow(iCounter))-1,CInt(aCol(iCounter))-1).setDesc(aValDesc(1))
'							Set objTable=Description.Create
'							objTable("Class Name").value="JavaTable"
'							Set objChld=objParameterDefinitionDialog.JavaTable("MaximumValues").ChildObjects(objTable)
'							objChld(0).SetCellData 1,0,aValDesc(1)
'							objChld(0).Object.setFocusable False
						End If
					else
						set objParameterDefinitionDialog=nothing
						Exit function
					end if
				Next
				Fn_Mech_MaximumValueTable=true
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
			Case "VerifyCellData"
				'Spliting row numbers
				aCol=Split(iCol,"~")
				aRow=Split(iRow,"~")
				aValue=Split(StrValue,"~")
				For iCounter=0 to ubound(aValue)
					bFlag=False
					cellval=objParameterDefinitionDialog.JavaTable("MaximumValues").Object.getCellData(Cint(aRow(iCounter))-1,Cint(aCol(iCounter))-1).toString()
					If instr(1,cellval,"[") Then cellval=Replace(cellval,"[","") end if
					If instr(1,cellval,"]") Then cellval=Replace(cellval,"]","") end if
					cellval=Split(cellval,",")
					aValDesc=Split(aValue(iCounter),":")

					If trim(cellval(UBound(cellval)))=trim(aValDesc(0)) Then
						If ubound(aValDesc)=1 Then
							If trim(cellval(2))=aValDesc(1) Then
								bFlag=True		
							End If
						Else
							bFlag=True
						End If
					else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Value [ " & aValue(iCounter) & " ] is not exist in [ Maximum Values ] table")
						Exit for
					End If
				Next
				If bFlag=True Then
					Fn_Mech_MaximumValueTable=true
				End If
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			'Case to get All Column name Exist in Table
			Case "GetAllColumnNames"
                	For iCounter=0 to objParameterDefinitionDialog.JavaObject("MaximumValuesTableHeader").Object.getColumnModel().getColumnCount()-1
						If iCounter=0 Then
							StrLabel=objParameterDefinitionDialog.JavaObject("MaximumValuesTableHeader").Object.getColumnModel().getColumn(0).getHeaderRenderer().getColName()
						else
							StrLabel=StrLabel+"~"+objParameterDefinitionDialog.JavaObject("MaximumValuesTableHeader").Object.getColumnModel().getColumn(iCounter).getHeaderRenderer().getColName()
						End If
					Next
					If Err.Number < 0 Then
						Fn_Mech_MaximumValueTable=False
					Else
						Fn_Mech_MaximumValueTable=StrLabel
					End If
		    '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			'Case to get Row ~ Column count currently exist in [ Maximum Values ] table
			Case "RowColumnCount"
'					iRows=objParameterDefinitionDialog.JavaTable("MaximumValues").GetROProperty("rows")
'					iCols=objParameterDefinitionDialog.JavaTable("MaximumValues").GetROProperty("cols")
					iCols =Fn_UI_Object_GetROProperty("Fn_PLM_PropertySetOperations",objParameterDefinitionDialog.JavaTable("MaximumValues"), "cols")
					iRows =Fn_UI_Object_GetROProperty("Fn_PLM_PropertySetOperations",objParameterDefinitionDialog.JavaTable("MaximumValues"),"rows")
					Fn_Mech_MaximumValueTable=iRows+"~"+iCols
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			'Case to paste data in [ Maximum Values ] table
			Case "PasteData"
'						For this case [ StrValue ] use to set Show Value Description Option
						If StrValue<>"" Then
							Call Fn_CheckBox_Set("Fn_Mech_MaximumValueTable", objParameterDefinitionDialog,"MaxValueDescriptionCell",StrValue)
						End If
						If iRow<>"" and iCol<>"" Then
							objParameterDefinitionDialog.JavaTable("MaximumValues").ClickCell Cint(iRow)-1,Cint(iCol)-1,"RIGHT"
						elseif iRow<>"" then
							objParameterDefinitionDialog.JavaTable("MaximumValuesRows").SelectRow Cint(iRow)-1
							objParameterDefinitionDialog.JavaTable("MaximumValuesRows").ClickCell Cint(iRow)-1,0,"RIGHT"
						else
							objParameterDefinitionDialog.JavaTable("MaximumValues").ClickCell 0,0,"RIGHT"
						End If
						objParameterDefinitionDialog.JavaMenu("index:=0","label:=Paste").Select
						If Err.Number < 0 Then
							Fn_Mech_MaximumValueTable=False
						Else
							Fn_Mech_MaximumValueTable=True
						End If
						If iRow<>"" and iCol="" Then
							objParameterDefinitionDialog.JavaTable("MaximumValues").DeselectRow Cint(iRow)-1
						End If
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			'Case to get all available Context menu for [ Maximum Values ] table
			Case "GetAllContextMenu"
						If iRow<>"" Then
							objParameterDefinitionDialog.JavaTable("MaximumValuesRows").SelectRow Cint(iRow)-1
							objParameterDefinitionDialog.JavaTable("MaximumValuesRows").ClickCell Cint(iRow)-1,0,"RIGHT"
						else
							objParameterDefinitionDialog.JavaTable("MaximumValues").ClickCell 0,0,"RIGHT"
						End If
						Set objMenu=Description.Create
						objMenu("Class Name").value="JavaMenu"
						Set objChld=objParameterDefinitionDialog.ChildObjects(objMenu)
						crrMenu=""
						For iCounter=0 to objChld.count-1
								If iCounter=0 Then
									crrMenu=objChld(0).GetROProperty("label")
								else
									crrMenu=crrMenu+"~"+objChld(iCounter).GetROProperty("label")
								End If
						Next
						If Err.Number < 0 Then
							Fn_Mech_MaximumValueTable=False
						Else
							If crrMenu<>"" Then
								Fn_Mech_MaximumValueTable=crrMenu
							else
								Fn_Mech_MaximumValueTable=False
							End If
						End If
						If iRow<>"" Then
							objParameterDefinitionDialog.JavaTable("MaximumValuesRows").ClickCell Cint(iRow)-1,0
						else
							objParameterDefinitionDialog.JavaTable("MaximumValues").ClickCell 0,0
						End If
						Set objChld=Nothing
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			'Case to paste data in [ Maximum Values ] table using Keyboard Ctrl+v
			Case "PasteData_Keyboard"
'						For this case [ StrValue ] use to set Show Value Description Option
						If StrValue<>"" Then
							Call Fn_CheckBox_Set("Fn_Mech_MaximumValueTable", objParameterDefinitionDialog,"MaxValueDescriptionCell",StrValue)
						End If
						If iRow<>"" and iCol<>"" Then
							objParameterDefinitionDialog.JavaTable("MaximumValues").SelectCell CInt(iRow)-1,Cint(iCol)-1
							wait 1
							objParameterDefinitionDialog.JavaTable("MaximumValues").PressKey "V",micCtrl
						Elseif iRow<>"" then
							objParameterDefinitionDialog.JavaTable("MaximumValuesRows").SelectRow Cint(iRow)-1
							wait 1
							objParameterDefinitionDialog.JavaTable("MaximumValuesRows").PressKey "V",micCtrl
						else
							objParameterDefinitionDialog.JavaTable("MaximumValues").SelectCell 0,0
							wait 1
							objParameterDefinitionDialog.JavaTable("MaximumValues").PressKey "V",micCtrl
						End If
						If Err.Number < 0 Then
							Fn_Mech_MaximumValueTable=False
						Else
							Fn_Mech_MaximumValueTable=True
						End If
						If iRow<>"" and iCol="" Then
							objParameterDefinitionDialog.JavaTable("MaximumValues").DeselectRow Cint(iRow)-1
						End If
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			'Case to Set Collapse checkbox On or Off
			Case "Collapse_SetON","Collapse_SetOFF"			
				If StrAction="Collapse_SetON" Then
					Fn_Mech_MaximumValueTable=Fn_CheckBox_Set("Fn_Mech_MaximumValueTable", objParameterDefinitionDialog, "CollapseMaximumValues","on")
				Else
					Fn_Mech_MaximumValueTable=Fn_CheckBox_Set("Fn_Mech_MaximumValueTable", objParameterDefinitionDialog, "CollapseMaximumValues","off")
				End If
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			'Case to Undo Changes in [ Maximum Values ] table
			Case "UndoChanges"
				objParameterDefinitionDialog.JavaTable("MaximumValues").ClickCell 0,0,"RIGHT"
				wait 1
				objParameterDefinitionDialog.JavaMenu("index:=0","label:=Undo").Select
				If Err.Number < 0 Then
					Fn_Mech_MaximumValueTable=False
				Else
					Fn_Mech_MaximumValueTable=True
				End If
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			'Case to get all Disable Context menu for [ Maximum Values ] table
			Case "GetAllDisableContextMenu"
						If iRow<>"" Then
							objParameterDefinitionDialog.JavaTable("MaximumValuesRows").SelectRow Cint(iRow)-1
							objParameterDefinitionDialog.JavaTable("MaximumValuesRows").ClickCell Cint(iRow)-1,0,"RIGHT"
						else
							objParameterDefinitionDialog.JavaTable("MaximumValues").ClickCell 0,0,"RIGHT"
						End If
						Set objMenu=Description.Create
						objMenu("Class Name").value="JavaMenu"
						Set objChld=objParameterDefinitionDialog.ChildObjects(objMenu)
						StrLabel=""
						For iCounter=0 to objChld.count-1
								crrMenu=objChld(iCounter).GetROProperty("label")
								If objChld(iCounter).CheckProperty("enabled",1,1)=false then
									If StrLabel="" Then
										StrLabel=objChld(iCounter).GetROProperty("label")
									else
										StrLabel=StrLabel+"~"+objChld(iCounter).GetROProperty("label")
									End If
								end if
						Next

						If Err.Number < 0 Then
							Fn_Mech_MaximumValueTable=False
						Else
							If StrLabel<>"" Then
								Fn_Mech_MaximumValueTable=StrLabel
							else
								Fn_Mech_MaximumValueTable=False
							End If
						End If
						If iRow<>"" Then
							objParameterDefinitionDialog.JavaTable("MaximumValuesRows").ClickCell Cint(iRow)-1,0
						else
							objParameterDefinitionDialog.JavaTable("MaximumValues").ClickCell 0,0
						End If
						Set objChld=Nothing
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
			'Case to Verify Date Data
			Case "VerifyDateData"
				'Spliting row numbers
				aCol=Split(iCol,"~")
				aRow=Split(iRow,"~")
				aValue=Split(StrValue,"~")
				For iCounter=0 to ubound(aValue)
					bFlag=False
					cellval=objParameterDefinitionDialog.JavaTable("MaximumValues").Object.getCellData(Cint(aRow(iCounter))-1,Cint(aCol(iCounter))-1).get(3).get(1).toString()
					aValDesc=Split(aValue(iCounter),":")
                    aDate=Split(cellval)

					If aDate(2)+"-"+aDate(1)+"-"+aDate(5)=aValDesc(0) or aDate(2)+"-0"+aDate(1)+"-"+aDate(5)=aValDesc(0) Then
						If ubound(aValDesc)=1 Then
							If trim(objParameterDefinitionDialog.JavaTable("MaximumValues").Object.getCellData(Cint(aRow(iCounter))-1,Cint(aCol(iCounter))-1).get(2).toString())=aValDesc(1) Then
								bFlag=True		
							End If
						Else
							bFlag=True
						End If
					else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Value [ " & aValue(iCounter) & " ] is not exist in [ Maximum Values ] table")
						Exit for
					End If
				Next
				If bFlag=True Then
					Fn_Mech_MaximumValueTable=true
				End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "PasteRowsCols"           
			If iRow<>"" and iCol<>"" Then
				aRow=Split(iRow,"~")
				aCol=Split(iCol,"~")
				objParameterDefinitionDialog.JavaTable("MaximumValues").SelectCellsRange cint(aRow(0))-1,cint(aCol(0))-1,cint(aRow(1))-1,Cint(aCol(1))-1
				objParameterDefinitionDialog.JavaTable("MaximumValues").ClickCell Cint(aRow(0))-1,Cint(aCol(0))-1,"RIGHT"
			elseif iRow<>"" then
				aRow=Split(iRow,"~")
				objParameterDefinitionDialog.JavaTable("MaximumValuesRows").SelectRow Cint(aRow(0))-1
				For iCounter=1 to ubound(aRow)
					objParameterDefinitionDialog.JavaTable("MaximumValuesRows").ExtendRow CInt(aRow(iCounter))-1
				Next
				objParameterDefinitionDialog.JavaTable("MaximumValues").ClickCell Cint(aRow(UBound(aRow)))-1,0,"RIGHT"
			End If
			objParameterDefinitionDialog.JavaMenu("index:=0","label:=Paste").Select
			If Err.Number < 0 Then
				Fn_Mech_MaximumValueTable=False
			Else
				Fn_Mech_MaximumValueTable=True
			End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to get specific Column name Exist in Table
		Case "GetColumnName"
				If iCol<>"" Then
					StrLabel=objParameterDefinitionDialog.JavaObject("MaximumValuesTableHeader").Object.getColumnModel().getColumn(CInt(iCol)-1).getHeaderRenderer().getColName()
					Fn_Mech_MaximumValueTable=StrLabel
				Else
					Fn_Mech_MaximumValueTable=false
				end if
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to get specific row name Exist in Table
		Case "GetRowName"
				If iRow<>"" Then
					StrLabel=objParameterDefinitionDialog.JavaTable("MaximumValuesRows").GetCellData(Cint(iRow)-1,0)
					Fn_Mech_MaximumValueTable=StrLabel
				Else
					Fn_Mech_MaximumValueTable=false
				end if
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			'Case to verify current value of cell is correct or not
			Case "IsCellCurrentValueCorrect"
				aCol=Split(iCol,"~")
				aRow=Split(iRow,"~")
				aValue=Split(StrValue,"~")
                For iCounter=0 to ubound(aCol)
						bFlag=false
						bFlag=objParameterDefinitionDialog.JavaTable("MaximumValues").Object.getValueAt(cint(aRow(iCounter))-1,cint(aCol(iCounter))-1).isValueCorrect()
						If bFlag<>lcase(cstr(aValue(iCounter))) Then
							bFlag=false
							Exit for
						End If
				Next
				 If bFlag=false Then
					Fn_Mech_MaximumValueTable=true
				else
					Fn_Mech_MaximumValueTable=false
				End If
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			'Case to verify table foreground color
			Case "VerifyHeaderForegroundColour"
                aCol=Split(iCol,"~")
				aValue=Split(StrValue,"~")
                For iCounter=0 to ubound(aCol)
						bFlag=false
						sColourCode=""
						sColour=objParameterDefinitionDialog.JavaObject("MaximumValuesTableHeader").Object.getColumnModel().getColumn(cint(aCol(iCounter))-1).getHeaderRenderer().getForeground().toString()
						sColour =  mid(sColour ,instr(sColour ,"[")  ,instr(sColour ,"]") )
						Select Case lCase(aValue(iCounter))
							Case "red"
								sColourCode="[r=255,g=0,b=0]"
						End Select
						If sColour=sColourCode Then
							bFlag=true
						else
							Exit for
						End if
				Next
                If bFlag=true Then
					Fn_Mech_MaximumValueTable=true
				else
					Fn_Mech_MaximumValueTable=false
				End If
	End Select
	set objParameterDefinitionDialog=nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_Mech_MinimumValueTable

'Description			 :	Function Used to perform operation on MinimumValueTable of parameter Defination Revision Information

'Parameters			   :   1.StrAction: Action Name
'										2.iCol: column number
'										3.iRow: row number
'										4.StrValue : values
'
'Return Value		   : 	True or False

'Pre-requisite			:	Parameter Defination dialog should be appear on screen on user should be on Define Parameter Defination Revision Information dialog

'Examples				:   bReturn=Fn_Mech_MinimumValueTable("SetData","1~2~3~1~2~3~1~2~3","1~1~1~2~2~2~3~3~3","1~2~3~4~5~6~7~8~9")
'										bReturn=Fn_Mech_MinimumValueTable("SetDateData","1~2~3~1~2~3~1~2~3","1~1~1~2~2~2~3~3~3","13-5-2012~14-5-2012~15-5-2012~16-5-2012~17-5-2012~18-5-2012~19-5-2012~20-5-2012~21-5-2012")
'										bReturn=Fn_Mech_MinimumValueTable("SetData","1~2~1~2","1~1~2~2","10:Ten~20:Twenty~30:Thirty~40:Four")
'										"10:Ten~20:Twenty~30:Thirty~40:Four"="Value~Value Description~Value~Value Description"
'										bReturn=Fn_Mech_MinimumValueTable("RowColumnCount","","","")
'										bReturn=Fn_Mech_MinimumValueTable("PasteData","","1","")
'										bReturn=Fn_Mech_MinimumValueTable("PasteData","2","1","")
'										bReturn=Fn_Mech_MinimumValueTable("PasteData_Keyboard","","","")
'										bReturn=Fn_Mech_MinimumValueTable("PasteData_Keyboard","2","2","")
'										bReturn=Fn_Mech_MinimumValueTable("GetAllContextMenu","","","")
'										bReturn=Fn_Mech_MinimumValueTable("GetAllContextMenu","","2","")
'										bReturn=Fn_Mech_MinimumValueTable("GetAllColumnNames","","2","")
'
'									   IMPNote : For Cases PasteData, PasteData_Keyboard use [ StrValue ] parameter to set Show Value Description Option
'
'										bReturn=Fn_Mech_MinimumValueTable("UndoChanges","","","")
'										bReturn=Fn_Mech_MinimumValueTable("GetAllDisableContextMenu","","","")
'										bReturn=Fn_Mech_MinimumValueTable("GetAllDisableContextMenu","","1","")
'										bReturn=Fn_Mech_MinimumValueTable("VerifyDateData","1","1","13-Jun-2012:Test")
'
'										bReturn=Fn_Mech_MinimumValueTable("GetColumnName","1","","")
'										bReturn=Fn_Mech_MinimumValueTable("GetRowName","","1","")
'
'												bReturn=Fn_Mech_MinimumValueTable("IsCellCurrentValueCorrect","1~2~2","1~1~2","false~true~false")
'												bReturn=Fn_Mech_MinimumValueTable("VerifyHeaderForegroundColour","1","","Red")
'
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												20-Apr-2012								1.0																						Sunny R
'													Sandeep N												26-Apr-2012								1.1						added case SetDateData			   Sunny R
'													Sandeep N												30-Jul-2012							   1.2						added GetColumnName & GetRowName			   Sonal P
'													Sandeep N												08-Aug-2012							   1.3						added IsCellCurrentValueCorrect & VerifyHeaderForegroundColour			   Sonal P
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_Mech_MinimumValueTable(StrAction,iCol,iRow,StrValue)
	GBL_FAILED_FUNCTION_NAME="Fn_Mech_MinimumValueTable"
 	'declaring variables
	Dim objParameterDefinitionDialog,objTable,objChld,objDate
	Dim aCol,aRow,aValue,scrollMax,iCounter,aDate,aValDesc
	Dim iRows,iCols,objMenu,crrMenu,StrLabel
	Dim sColourCode,sColour

	Fn_Mech_MinimumValueTable=False
	If Fn_SISW_UI_Object_Operations("Fn_Mech_MinimumValueTable","Exist",Window("MechatronicsWindow").JavaDialog("NewParameterDefinition"),SISW_MIN_TIMEOUT) = False Then
'	if not Window("MechatronicsWindow").JavaDialog("NewParameterDefinition").Exist(6) then
		Exit function
	else
		'Creating object of [ NewParameterDefinition ] dialog
		set objParameterDefinitionDialog=Window("MechatronicsWindow").JavaDialog("NewParameterDefinition")
	end if
	If Fn_SISW_UI_Object_Operations("Fn_Mech_MinimumValueTable","Exist",objParameterDefinitionDialog.JavaSlider("JScrollPane"),SISW_MINLESS_TIMEOUT) = True Then
'	If objParameterDefinitionDialog.JavaSlider("JScrollPane").Exist(3) Then
		'Scrolling till the mid of panel
		scrollMax=objParameterDefinitionDialog.JavaSlider("JScrollPane").GetROProperty("max")
		objParameterDefinitionDialog.JavaSlider("JScrollPane").Drag scrollMax/3
	End If

'	iRowCount=objParameterDefinitionDialog.JavaTable("MaximumValues").GetROProperty("rows")
	Select Case StrAction
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			Case "SetData"
				aCol=Split(iCol,"~")
				aRow=Split(iRow,"~")
				aValue=Split(StrValue,"~")
				If inStr(1,StrValue,":") Then
					Call Fn_CheckBox_Set("Fn_Mech_MinimumValueTable", objParameterDefinitionDialog,"MinValueDescriptionCell","on")
				End If

				objParameterDefinitionDialog.JavaTable("MinimumValues").SelectRow 0
				For iCounter=0 to UBound(aRow)
					objParameterDefinitionDialog.JavaTable("MinimumValues").ActivateCell CInt(aRow(iCounter))-1,CInt(aCol(iCounter))-1
					wait 1
					aValDesc=Split(aValue(iCounter),":")

					Set objTable=Description.Create
					objTable("Class Name").value="JavaTable"
					Set objChld=objParameterDefinitionDialog.JavaTable("MinimumValues").ChildObjects(objTable)
					If aValDesc(0)<>"" Then
						objChld(0).SetCellData 0,0,aValDesc(0)
					End If
					If uBound(aValDesc)=1 Then
						objChld(0).SetCellData 1,0,aValDesc(1)
					End If

					Set objChld=nothing
					Set objTable=nothing
				Next
				Fn_Mech_MinimumValueTable=true
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			Case "SetDateData"
				aCol=Split(iCol,"~")
				aRow=Split(iRow,"~")
				aValue=Split(StrValue,"~")
				If inStr(1,StrValue,":") Then
					Call Fn_CheckBox_Set("Fn_Mech_MinimumValueTable", objParameterDefinitionDialog,"MinValueDescriptionCell","on")
				End If

				For iCounter=0 to UBound(aRow)
					objParameterDefinitionDialog.JavaTable("MinimumValues").ClickCell CInt(aRow(iCounter))-1,CInt(aCol(iCounter))-1,"LEFT"
					objParameterDefinitionDialog.JavaTable("MinimumValues").DoubleClickCell CInt(aRow(iCounter))-1,CInt(aCol(iCounter))-1,"LEFT"
					wait 1
					If JavaDialog("SelectDate").Exist(2) then
						aValDesc=Split(aValue(iCounter),":")
						aDate=Split(aValDesc(0),"-")

						Set objDate=JavaDialog("SelectDate").JavaObject("CalendarPanel").Object.getTime()
						objDate.setYear(Cint(aDate(2))-1900)
						objDate.setMonth(Cint(aDate(1))-1)
						objDate.setDate(aDate(0))
                        						
						JavaDialog("SelectDate").JavaObject("CalendarPanel").Object.setTime(objDate)
						JavaDialog("SelectDate").JavaButton("Ok").Click
						If ubound(aValDesc)=1 Then
                            objParameterDefinitionDialog.JavaTable("MinimumValues").Object.getValueAt(CInt(aRow(iCounter))-1,CInt(aCol(iCounter))-1).setDesc(aValDesc(1))
'							Set objTable=Description.Create
'							objTable("Class Name").value="JavaTable"
'							Set objChld=objParameterDefinitionDialog.JavaTable("MinimumValues").ChildObjects(objTable)
'							objChld(0).SetCellData 1,0,aValDesc(1)
'							objChld(0).Object.setFocusable False
						End If
					else
						set objParameterDefinitionDialog=nothing
						Exit function
					end if
				Next
				Fn_Mech_MinimumValueTable=true
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
			Case "VerifyCellData"
				'Spliting row numbers
				aCol=Split(iCol,"~")
				aRow=Split(iRow,"~")
				aValue=Split(StrValue,"~")
				For iCounter=0 to ubound(aValue)
					bFlag=False
					cellval=objParameterDefinitionDialog.JavaTable("MinimumValues").Object.getCellData(Cint(aRow(iCounter))-1,Cint(aCol(iCounter))-1).toString()
					If instr(1,cellval,"[") Then cellval=Replace(cellval,"[","") end if
					If instr(1,cellval,"]") Then cellval=Replace(cellval,"]","") end if
					cellval=Split(cellval,",")
					aValDesc=Split(aValue(iCounter),":")

					If trim(cellval(UBound(cellval)))=trim(aValDesc(0)) Then
						If ubound(aValDesc)=1 Then
							If trim(cellval(2))=aValDesc(1) Then
								bFlag=True		
							End If
						Else
							bFlag=True
						End If
					else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Value [ " & aValue(iCounter) & " ] is not exist in [ Maximum Values ] table")
						Exit for
					End If
				Next
				If bFlag=True Then
					Fn_Mech_MinimumValueTable=true
				End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			'Case to get All Column name Exist in Table
			Case "GetAllColumnNames"
                	For iCounter=0 to objParameterDefinitionDialog.JavaObject("MinimumValuesTableHeader").Object.getColumnModel().getColumnCount()-1
						If iCounter=0 Then
							StrLabel=objParameterDefinitionDialog.JavaObject("MinimumValuesTableHeader").Object.getColumnModel().getColumn(0).getHeaderRenderer().getColName()
						else
							StrLabel=StrLabel+"~"+objParameterDefinitionDialog.JavaObject("MinimumValuesTableHeader").Object.getColumnModel().getColumn(iCounter).getHeaderRenderer().getColName()
						End If
					Next
					If Err.Number < 0 Then
						Fn_Mech_MinimumValueTable=False
					Else
						Fn_Mech_MinimumValueTable=StrLabel
					End If
		    '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			'Case to get Row ~ Column count currently exist in [ Minimum Values ] table
			Case "RowColumnCount"
					iRows=objParameterDefinitionDialog.JavaTable("MinimumValues").GetROProperty("rows")
					iCols=objParameterDefinitionDialog.JavaTable("MinimumValues").GetROProperty("cols")
					Fn_Mech_MinimumValueTable=iRows+"~"+iCols
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			'Case to paste data in [ Minimum Values ] table
			Case "PasteData"
						'For this case use [ StrValue ] parameter to Set Show Value Description option
						If StrValue<>"" Then
							Call Fn_CheckBox_Set("Fn_Mech_MinimumValueTable", objParameterDefinitionDialog,"MinValueDescriptionCell",StrValue)
						End If
						If iRow<>"" and iCol<>"" Then
							objParameterDefinitionDialog.JavaTable("MinimumValues").ClickCell Cint(iRow)-1,Cint(iCol)-1,"RIGHT"
						elseif iRow<>"" then
							objParameterDefinitionDialog.JavaTable("MinimumValuesRows").SelectRow Cint(iRow)-1
							objParameterDefinitionDialog.JavaTable("MinimumValuesRows").ClickCell Cint(iRow)-1,0,"RIGHT"
						else
							objParameterDefinitionDialog.JavaTable("MinimumValues").ClickCell 0,0,"RIGHT"
						End If
						objParameterDefinitionDialog.JavaMenu("index:=0","label:=Paste").Select
						If Err.Number < 0 Then
							Fn_Mech_MinimumValueTable=False
						Else
							Fn_Mech_MinimumValueTable=True
						End If
						If iRow<>"" and iCol="" Then
							objParameterDefinitionDialog.JavaTable("MinimumValues").DeselectRow Cint(iRow)-1
						End If
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			'Case to get all available Context menu for [ Minimum Values ] table
			Case "GetAllContextMenu"
						If iRow<>"" Then
							objParameterDefinitionDialog.JavaTable("MinimumValuesRows").SelectRow Cint(iRow)-1
							objParameterDefinitionDialog.JavaTable("MinimumValuesRows").ClickCell Cint(iRow)-1,0,"RIGHT"
						else
							objParameterDefinitionDialog.JavaTable("MinimumValues").ClickCell 0,0,"RIGHT"
						End If
						Set objMenu=Description.Create
						objMenu("Class Name").value="JavaMenu"
						Set objChld=objParameterDefinitionDialog.ChildObjects(objMenu)
						crrMenu=""
						For iCounter=0 to objChld.count-1
								If iCounter=0 Then
									crrMenu=objChld(0).GetROProperty("label")
								else
									crrMenu=crrMenu+"~"+objChld(iCounter).GetROProperty("label")
								End If
						Next
						If Err.Number < 0 Then
							Fn_Mech_MinimumValueTable=False
						Else
							If crrMenu<>"" Then
								Fn_Mech_MinimumValueTable=crrMenu
							else
								Fn_Mech_MinimumValueTable=False
							End If
						End If
						If iRow<>"" Then
							objParameterDefinitionDialog.JavaTable("MinimumValuesRows").ClickCell Cint(iRow)-1,0
						else
							objParameterDefinitionDialog.JavaTable("MinimumValues").ClickCell 0,0
						End If
						Set objChld=Nothing
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			'Case to paste data in [ Minimum Values ] table using Keyboard Ctrl+v
			Case "PasteData_Keyboard"
						'For this case use [ StrValue ] parameter to Set Show Value Description option
						If StrValue<>"" Then
							Call Fn_CheckBox_Set("Fn_Mech_MinimumValueTable", objParameterDefinitionDialog,"MinValueDescriptionCell",StrValue)
						End If
						If iRow<>"" and iCol<>"" Then
							objParameterDefinitionDialog.JavaTable("MinimumValues").SelectCell CInt(iRow)-1,Cint(iCol)-1
							wait 1
							objParameterDefinitionDialog.JavaTable("MinimumValues").PressKey "V",micCtrl
						Elseif iRow<>"" then
							objParameterDefinitionDialog.JavaTable("MinimumValuesRows").SelectRow Cint(iRow)-1
							wait 1
							objParameterDefinitionDialog.JavaTable("MinimumValuesRows").PressKey "V",micCtrl
						else
							objParameterDefinitionDialog.JavaTable("MinimumValues").SelectCell 0,0
							wait 1
							objParameterDefinitionDialog.JavaTable("MinimumValues").PressKey "V",micCtrl
						End If
						If Err.Number < 0 Then
							Fn_Mech_MinimumValueTable=False
						Else
							Fn_Mech_MinimumValueTable=True
						End If
						If iRow<>"" and iCol="" Then
							objParameterDefinitionDialog.JavaTable("MinimumValues").DeselectRow Cint(iRow)-1
						End If
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			'Case to Set Collapse checkbox On or Off
			Case "Collapse_SetON","Collapse_SetOFF"			
				If StrAction="Collapse_SetON" Then
					Fn_Mech_MinimumValueTable=Fn_CheckBox_Set("Fn_Mech_MinimumValueTable", objParameterDefinitionDialog, "CollapseMinimumValues","on")
				Else
					Fn_Mech_MinimumValueTable=Fn_CheckBox_Set("Fn_Mech_MinimumValueTable", objParameterDefinitionDialog, "CollapseMinimumValues","off")
				End If
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			'Case to Undo Changes in [ Minimum Values ] table
			Case "UndoChanges"
						objParameterDefinitionDialog.JavaTable("MinimumValues").ClickCell 0,0,"RIGHT"
						wait 1
						objParameterDefinitionDialog.JavaMenu("index:=0","label:=Undo").Select
						If Err.Number < 0 Then
							Fn_Mech_MinimumValueTable=False
						Else
							Fn_Mech_MinimumValueTable=True
						End If
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			'Case to get all Disable Context menu for [ Minimum Values ] table
			Case "GetAllDisableContextMenu"
						If iRow<>"" Then
							objParameterDefinitionDialog.JavaTable("MinimumValuesRows").SelectRow Cint(iRow)-1
							objParameterDefinitionDialog.JavaTable("MinimumValuesRows").ClickCell Cint(iRow)-1,0,"RIGHT"
						else
							objParameterDefinitionDialog.JavaTable("MinimumValues").ClickCell 0,0,"RIGHT"
						End If
						Set objMenu=Description.Create
						objMenu("Class Name").value="JavaMenu"
						Set objChld=objParameterDefinitionDialog.ChildObjects(objMenu)
						StrLabel=""
						For iCounter=0 to objChld.count-1
								crrMenu=objChld(iCounter).GetROProperty("label")
								If objChld(iCounter).CheckProperty("enabled",1,1)=false then
									If StrLabel="" Then
										StrLabel=objChld(iCounter).GetROProperty("label")
									else
										StrLabel=StrLabel+"~"+objChld(iCounter).GetROProperty("label")
									End If
								end if
						Next

						If Err.Number < 0 Then
							Fn_Mech_MinimumValueTable=False
						Else
							If StrLabel<>"" Then
								Fn_Mech_MinimumValueTable=StrLabel
							else
								Fn_Mech_MinimumValueTable=False
							End If
						End If
						If iRow<>"" Then
							objParameterDefinitionDialog.JavaTable("MinimumValuesRows").ClickCell Cint(iRow)-1,0
						else
							objParameterDefinitionDialog.JavaTable("MinimumValues").ClickCell 0,0
						End If
						Set objChld=Nothing
				'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
			'Case to Verify Date Data
			Case "VerifyDateData"
				'Spliting row numbers
				aCol=Split(iCol,"~")
				aRow=Split(iRow,"~")
				aValue=Split(StrValue,"~")
				For iCounter=0 to ubound(aValue)
					bFlag=False
					cellval=objParameterDefinitionDialog.JavaTable("MinimumValues").Object.getCellData(Cint(aRow(iCounter))-1,Cint(aCol(iCounter))-1).get(3).get(1).toString()
					aValDesc=Split(aValue(iCounter),":")
                    aDate=Split(cellval)

					If aDate(2)+"-"+aDate(1)+"-"+aDate(5)=aValDesc(0) or aDate(2)+"-0"+aDate(1)+"-"+aDate(5)=aValDesc(0) Then
						If ubound(aValDesc)=1 Then
							If trim(objParameterDefinitionDialog.JavaTable("MinimumValues").Object.getCellData(Cint(aRow(iCounter))-1,Cint(aCol(iCounter))-1).get(2).toString())=aValDesc(1) Then
								bFlag=True		
							End If
						Else
							bFlag=True
						End If
					else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Value [ " & aValue(iCounter) & " ] is not exist in [ Minimum Values ] table")
						Exit for
					End If
				Next
				If bFlag=True Then
					Fn_Mech_MinimumValueTable=true
				End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to get specific Column name Exist in Table
		Case "GetColumnName"
				If iCol<>"" Then
					StrLabel=objParameterDefinitionDialog.JavaObject("MinimumValuesTableHeader").Object.getColumnModel().getColumn(CInt(iCol)-1).getHeaderRenderer().getColName()
					Fn_Mech_MinimumValueTable=StrLabel
				Else
					Fn_Mech_MinimumValueTable=false
				end if
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to get specific row name Exist in Table
		Case "GetRowName"
				If iRow<>"" Then
					StrLabel=objParameterDefinitionDialog.JavaTable("MinimumValuesRows").GetCellData(Cint(iRow)-1,0)
					Fn_Mech_MinimumValueTable=StrLabel
				Else
					Fn_Mech_MinimumValueTable=false
				end if
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "PasteRowsCols"           
			If iRow<>"" and iCol<>"" Then
				aRow=Split(iRow,"~")
				aCol=Split(iCol,"~")
				objParameterDefinitionDialog.JavaTable("MinimumValues").SelectCellsRange cint(aRow(0))-1,cint(aCol(0))-1,cint(aRow(1))-1,Cint(aCol(1))-1
				objParameterDefinitionDialog.JavaTable("MinimumValues").ClickCell Cint(aRow(0))-1,Cint(aCol(0))-1,"RIGHT"
			elseif iRow<>"" then
				aRow=Split(iRow,"~")
				objParameterDefinitionDialog.JavaTable("MinimumValuesRows").SelectRow Cint(aRow(0))-1
				For iCounter=1 to ubound(aRow)
					objParameterDefinitionDialog.JavaTable("MinimumValuesRows").ExtendRow CInt(aRow(iCounter))-1
				Next
				objParameterDefinitionDialog.JavaTable("MinimumValues").ClickCell Cint(aRow(UBound(aRow)))-1,0,"RIGHT"
			End If
			objParameterDefinitionDialog.JavaMenu("index:=0","label:=Paste").Select
			If Err.Number < 0 Then
				Fn_Mech_MinimumValueTable=False
			Else
				Fn_Mech_MinimumValueTable=True
			End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			'Case to verify current value of cell is correct or not
			Case "IsCellCurrentValueCorrect"
				aCol=Split(iCol,"~")
				aRow=Split(iRow,"~")
				aValue=Split(StrValue,"~")
                For iCounter=0 to ubound(aCol)
						bFlag=false
						bFlag=objParameterDefinitionDialog.JavaTable("MinimumValues").Object.getValueAt(cint(aRow(iCounter))-1,cint(aCol(iCounter))-1).isValueCorrect()
						If bFlag<>lcase(cstr(aValue(iCounter))) Then
							bFlag=false
							Exit for
						End If
				Next
				 If bFlag=false Then
					Fn_Mech_MinimumValueTable=true
				else
					Fn_Mech_MinimumValueTable=false
				End If
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			'Case to verify table foreground color
			Case "VerifyHeaderForegroundColour"
                aCol=Split(iCol,"~")
				aValue=Split(StrValue,"~")
                For iCounter=0 to ubound(aCol)
						bFlag=false
						sColourCode=""
						sColour=objParameterDefinitionDialog.JavaObject("MinimumValuesTableHeader").Object.getColumnModel().getColumn(cint(aCol(iCounter))-1).getHeaderRenderer().getForeground().toString()
						sColour =  mid(sColour ,instr(sColour ,"[")  ,instr(sColour ,"]") )
						Select Case lCase(aValue(iCounter))
							Case "red"
								sColourCode="[r=255,g=0,b=0]"
						End Select
						If sColour=sColourCode Then
							bFlag=true
						else
							Exit for
						End if
				Next
                If bFlag=true Then
					Fn_Mech_MinimumValueTable=true
				else
					Fn_Mech_MinimumValueTable=false
				End If
	End Select
	set objParameterDefinitionDialog=nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_Mech_InitialValueTable

'Description			 :	Function Used to perform operation on InitialValueTable of parameter Defination Revision Information

'Parameters			   :   1.StrAction: Action Name
'										2.iCol: column number
'										3.iRow: row number
'										4.StrValue : values
'
'Return Value		   : 	True or False

'Pre-requisite			:	Parameter Defination dialog should be appear on screen on user should be on Define Parameter Defination Revision Information dialog

'Examples				:   bReturn=Fn_Mech_InitialValueTable("SetData","1~2~3~1~2~3~1~2~3","1~1~1~2~2~2~3~3~3","10~20~30~40~50~60~70~80~90")
'										bReturn=Fn_Mech_InitialValueTable("SetDateData","1~2~3~1~2~3~1~2~3","1~1~1~2~2~2~3~3~3","13-5-2012~14-5-2012~15-5-2012~16-5-2012~17-5-2012~18-5-2012~19-5-2012~20-5-2012~21-5-2012")
'										bRetrun=Fn_Mech_InitialValueTable("SetData","1~2~1~2","1~1~2~2","10:Ten~20:Twenty~30:Thirty~40:Four")
'										"10:Ten~20:Twenty~30:Thirty~40:Four"="Value:Value Description~Value:Value Description"
'										bReturn=Fn_Mech_InitialValueTable("SetBoolData","1~2~1~2","1~1~2~2","false:Desc24~false:Descfalse~false:Desc235~false:Desc467")
'										bReturn=Fn_Mech_InitialValueTable("RowColumnCount","","","")
'										bReturn=Fn_Mech_InitialValueTable("PasteData","","1","")
'										bReturn=Fn_Mech_InitialValueTable("PasteData","2","1","")
'										bReturn=Fn_Mech_InitialValueTable("PasteData_Keyboard","","","")
'										bReturn=Fn_Mech_InitialValueTable("PasteData_Keyboard","2","2","")
'										bReturn=Fn_Mech_InitialValueTable("GetAllContextMenu","","","")
'										bReturn=Fn_Mech_InitialValueTable("GetAllContextMenu","","2","")
'										bReturn=Fn_Mech_InitialValueTable("GetAllColumnNames","","2","")
'
''									   IMPNote : For Cases PasteData, PasteData_Keyboard use [ StrValue ] parameter to set Show Value Description Option
'
'										bReturn=Fn_Mech_InitialValueTable("UndoChanges","","","")
'										bReturn=Fn_Mech_InitialValueTable("GetAllDisableContextMenu","","","")
'										bReturn=Fn_Mech_InitialValueTable("GetAllDisableContextMenu","","1","")
'										bReturn=Fn_Mech_InitialValueTable("VerifyDateData","1","1","13-Jun-2012:Test")
'
'										bReturn=Fn_Mech_InitialValueTable("GetColumnName","1","","")
'										bReturn=Fn_Mech_InitialValueTable("GetRowName","","1","")
'
'
'										bReturn=Fn_Mech_InitialValueTable("PasteRowsCols","1~2","1~2","")
'
'												bReturn=Fn_Mech_InitialValueTable("IsCellCurrentValueCorrect","1~2~2","1~1~2","false~true~false")
'												bReturn=Fn_Mech_InitialValueTable("VerifyHeaderForegroundColour","1","","Red")
'
'										bReturn=Fn_Mech_InitialValueTable("SelectCell","2","3","off")
'										bReturn=Fn_Mech_InitialValueTable("DoubleClickCell","2","3","off")
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												20-Apr-2012								1.0																						Sunny R
'													Sandeep N												26-Apr-2012								1.1						added case SetDateData			   Sunny R
'													Sandeep N												28-Apr-2012								1.2						added case SetBoolData			   Sunny R
'													Sandeep N												30-Jul-2012							   1.3						added GetColumnName & GetRowName			   Sonal P
'													Sandeep N												31-Jul-2012							   1.4						added case PasteRowsCols		Anjali M
'													Sandeep N												08-Aug-2012							   1.5						added case IsCellCurrentValueCorrect & VerifyHeaderForegroundColour		Anjali M
'													Sandeep N												28-Aug-2012							   1.6						added case SelectCell & DoubleClickCell		Sachin J
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_Mech_InitialValueTable(StrAction,iCol,iRow,StrValue)
	GBL_FAILED_FUNCTION_NAME="Fn_Mech_InitialValueTable"
 	'declaring variables
	Dim objParameterDefinitionDialog,objTable,objChld
	Dim aCol,aRow,aValue,scrollMax,iCounter,aValDesc
	Dim iRows,iCols,objMenu,crrMenu,StrLabel
	Dim sColourCode,sColour

	Fn_Mech_InitialValueTable=False
	If Fn_SISW_UI_Object_Operations("Fn_Mech_InitialValueTable","Exist",Window("MechatronicsWindow").JavaDialog("NewParameterDefinition"),SISW_MIN_TIMEOUT) = False Then
'	if not Window("MechatronicsWindow").JavaDialog("NewParameterDefinition").Exist(6) then
		Exit function
	else
		'Creating object of [ NewParameterDefinition ] dialog
		set objParameterDefinitionDialog=Window("MechatronicsWindow").JavaDialog("NewParameterDefinition")
	end if
	If Fn_SISW_UI_Object_Operations("Fn_Mech_InitialValueTable","Exist",objParameterDefinitionDialog.JavaSlider("JScrollPane"),SISW_MINLESS_TIMEOUT) = True Then
'	If objParameterDefinitionDialog.JavaSlider("JScrollPane").Exist(3) Then
		'Scrolling till the mid of panel
		scrollMax=objParameterDefinitionDialog.JavaSlider("JScrollPane").GetROProperty("max")
		objParameterDefinitionDialog.JavaSlider("JScrollPane").Drag scrollMax/1
	End If

'	iRowCount=objParameterDefinitionDialog.JavaTable("MaximumValues").GetROProperty("rows")
	Select Case StrAction
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
			Case "SetData"
				aCol=Split(iCol,"~")
				aRow=Split(iRow,"~")
				aValue=Split(StrValue,"~")
				If inStr(1,StrValue,":") Then
					Call Fn_CheckBox_Set("Fn_Mech_InitialValueTable", objParameterDefinitionDialog,"IntValueDescriptionCell","on")
				End If
				objParameterDefinitionDialog.JavaTable("InitialValues").SelectRow 0
				For iCounter=0 to UBound(aRow)
					objParameterDefinitionDialog.JavaTable("InitialValues").ActivateCell CInt(aRow(iCounter))-1,CInt(aCol(iCounter))-1
					wait 1
					aValDesc=Split(aValue(iCounter),":")

					Set objTable=Description.Create
					objTable("Class Name").value="JavaTable"
					Set objChld=objParameterDefinitionDialog.JavaTable("InitialValues").ChildObjects(objTable)
					If aValDesc(0)<>"" Then
						objChld(0).SetCellData 0,0,aValDesc(0)
					End If
					If uBound(aValDesc)=1 Then
						objChld(0).SetCellData 1,0,aValDesc(1)
					End If
					Set objChld=nothing
					Set objTable=nothing
				Next
				Fn_Mech_InitialValueTable=true
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
			Case "SetDateData"
				aCol=Split(iCol,"~")
				aRow=Split(iRow,"~")
				aValue=Split(StrValue,"~")
				If inStr(1,StrValue,":") Then
					Call Fn_CheckBox_Set("Fn_Mech_InitialValueTable", objParameterDefinitionDialog,"IntValueDescriptionCell","on")
				End If
				For iCounter=0 to UBound(aRow)
					objParameterDefinitionDialog.JavaTable("InitialValues").ClickCell CInt(aRow(iCounter))-1,CInt(aCol(iCounter))-1,"LEFT"
					objParameterDefinitionDialog.JavaTable("InitialValues").DoubleClickCell CInt(aRow(iCounter))-1,CInt(aCol(iCounter))-1,"LEFT"
			
					wait 1
					If JavaDialog("SelectDate").Exist(2) then
						aValDesc=Split(aValue(iCounter),":")
						aDate=Split(aValDesc(0),"-")

						Set objDate=JavaDialog("SelectDate").JavaObject("CalendarPanel").Object.getTime()
						objDate.setYear(Cint(aDate(2))-1900)
						objDate.setMonth(Cint(aDate(1))-1)
						objDate.setDate(aDate(0))				

						JavaDialog("SelectDate").JavaObject("CalendarPanel").Object.setTime(objDate)
						JavaDialog("SelectDate").JavaButton("Ok").Click
						If ubound(aValDesc)=1 Then
                            objParameterDefinitionDialog.JavaTable("InitialValues").Object.getValueAt(CInt(aRow(iCounter))-1,CInt(aCol(iCounter))-1).setDesc(aValDesc(1))
'							Set objTable=Description.Create
'							objTable("Class Name").value="JavaTable"
'							Set objChld=objParameterDefinitionDialog.JavaTable("InitialValues").ChildObjects(objTable)
'							objChld(0).SetCellData 1,0,aValDesc(1)
'							objChld(0).Object.setFocusable False
						End If
					else
						set objParameterDefinitionDialog=nothing
						Exit function
					end if
				Next
				Fn_Mech_InitialValueTable=true
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
			Case "VerifyCellData"
				'Spliting row numbers
				aCol=Split(iCol,"~")
				aRow=Split(iRow,"~")
				aValue=Split(StrValue,"~")
				For iCounter=0 to ubound(aValue)
					bFlag=False
					aValDesc=Split(aValue(iCounter),":")
					cellval=objParameterDefinitionDialog.JavaTable("InitialValues").Object.getCellData(Cint(aRow(iCounter))-1,Cint(aCol(iCounter))-1).get(3).get(1).toString()
					If trim(cellval)=trim(aValDesc(0)) Then
						If ubound(aValDesc)=1 Then
							cellval=objParameterDefinitionDialog.JavaTable("InitialValues").Object.getCellData(Cint(aRow(iCounter))-1,Cint(aCol(iCounter))-1).get(2).toString()
							If trim(cellval)=aValDesc(1) Then
								bFlag=True		
							End If
						Else
							bFlag=True
						End If
					else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Value [ " & aValue(iCounter) & " ] is not exist in [ Maximum Values ] table")
						Exit for
					End If
				Next
				If bFlag=True Then
					Fn_Mech_InitialValueTable=true
				End If
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
			Case "SetBoolData"
				aCol=Split(iCol,"~")
				aRow=Split(iRow,"~")
				aValue=Split(StrValue,"~")
				If objParameterDefinitionDialog.JavaCheckBox("IntValueDescriptionCell").GetROProperty("enabled")="1" Then
					Call Fn_CheckBox_Set("Fn_Mech_InitialValueTable", objParameterDefinitionDialog,"IntValueDescriptionCell","on")
				End If

				objParameterDefinitionDialog.JavaTable("InitialValues").SelectRow 0
                If objParameterDefinitionDialog.JavaCheckBox("CollapseInitialValues").GetROProperty("value")="1" Then
					If aValue(0)<>objParameterDefinitionDialog.JavaTable("InitialValues").Object.getCellData(0,0).get(3).get(1).toString() Then
						objParameterDefinitionDialog.JavaTable("InitialValues").DoubleClickCell CInt(aRow(iCounter))-1,CInt(aCol(iCounter))-1
					End If
				else
					For iCounter=0 to UBound(aRow)
						objParameterDefinitionDialog.JavaTable("InitialValues").DoubleClickCell CInt(aRow(iCounter))-1,CInt(aCol(iCounter))-1
						
						wait 1
						aValDesc=Split(aValue(iCounter),":")
	
						Set objTable=Description.Create
						objTable("Class Name").value="JavaTable"
						Set objChld=objParameterDefinitionDialog.JavaTable("InitialValues").ChildObjects(objTable)
						If aValDesc(0)<>"" Then
							If trim(objChld(0).getCellData(0,0))<>aValDesc(0) Then
								objChld(0).SelectRow 0
								objChld(0).ClickCell 0,0
							End If
						End If
						If uBound(aValDesc)=1 Then
							objChld(0).SelectRow 1
							objChld(0).SetCellData 1,0,aValDesc(1)
						End If
	
					'	objChld(0).Object.setFocusable False
						Set objChld=nothing
						Set objTable=nothing
					Next
				end if
				If Err.Number < 0 Then
					Fn_Mech_InitialValueTable=False
				Else
					Fn_Mech_InitialValueTable=True
				End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			'Case to get All Column name Exist in Table
			Case "GetAllColumnNames"
                	For iCounter=0 to objParameterDefinitionDialog.JavaObject("InitialValuesTableHeader").Object.getColumnModel().getColumnCount()-1
						If iCounter=0 Then
							StrLabel=objParameterDefinitionDialog.JavaObject("InitialValuesTableHeader").Object.getColumnModel().getColumn(0).getHeaderRenderer().getColName()
						else
							StrLabel=StrLabel+"~"+objParameterDefinitionDialog.JavaObject("InitialValuesTableHeader").Object.getColumnModel().getColumn(iCounter).getHeaderRenderer().getColName()
						End If
					Next
					If Err.Number < 0 Then
						Fn_Mech_InitialValueTable=False
					Else
						Fn_Mech_InitialValueTable=StrLabel
					End If
		    '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			'Case to get Row ~ Column count currently exist in [ Initial Values ] table
			Case "RowColumnCount"
					iRows=objParameterDefinitionDialog.JavaTable("InitialValues").GetROProperty("rows")
					iCols=objParameterDefinitionDialog.JavaTable("InitialValues").GetROProperty("cols")
					Fn_Mech_InitialValueTable=iRows+"~"+iCols
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			'Case to paste data in [ Initial Values ] table
			Case "PasteData"
						'For this Case use [ StrValue ] parameter to set show ValueDescription option
						If StrValue<>"" Then
							Call Fn_CheckBox_Set("Fn_Mech_InitialValueTable", objParameterDefinitionDialog,"IntValueDescriptionCell",StrValue)
						End If
						If iRow<>"" and iCol<>"" Then
							objParameterDefinitionDialog.JavaTable("InitialValues").ClickCell Cint(iRow)-1,Cint(iCol)-1,"RIGHT"
						elseif iRow<>"" then
							objParameterDefinitionDialog.JavaTable("InitialValuesRows").SelectRow Cint(iRow)-1
							objParameterDefinitionDialog.JavaTable("InitialValuesRows").ClickCell Cint(iRow)-1,0,"RIGHT"
						elseif iCol<>"" Then
							objParameterDefinitionDialog.JavaTable("InitialValues").SelectColumn Cint(iCol)-1
							objParameterDefinitionDialog.JavaTable("InitialValues").SelectColumnHeader Cint(iCol)-1,"RIGHT"
						else
							objParameterDefinitionDialog.JavaTable("InitialValues").ClickCell 0,0,"RIGHT"
						End If
						objParameterDefinitionDialog.JavaMenu("index:=0","label:=Paste").Select
						If Err.Number < 0 Then
							Fn_Mech_InitialValueTable=False
						Else
							Fn_Mech_InitialValueTable=True
						End If
						If iRow<>"" and iCol="" Then
							objParameterDefinitionDialog.JavaTable("InitialValues").DeselectRow Cint(iRow)-1
						End If
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			'Case to get all available Context menu for [ Initial Values ] table
			Case "GetAllContextMenu"
						If iRow<>"" Then
							objParameterDefinitionDialog.JavaTable("InitialValuesRows").SelectRow Cint(iRow)-1
							objParameterDefinitionDialog.JavaTable("InitialValuesRows").ClickCell Cint(iRow)-1,0,"RIGHT"
						else
							objParameterDefinitionDialog.JavaTable("InitialValues").ClickCell 0,0,"RIGHT"
						End If
						Set objMenu=Description.Create
						objMenu("Class Name").value="JavaMenu"
						Set objChld=objParameterDefinitionDialog.ChildObjects(objMenu)
						crrMenu=""
						For iCounter=0 to objChld.count-1
								If iCounter=0 Then
									crrMenu=objChld(0).GetROProperty("label")
								else
									crrMenu=crrMenu+"~"+objChld(iCounter).GetROProperty("label")
								End If
						Next
						If Err.Number < 0 Then
							Fn_Mech_InitialValueTable=False
						Else
							If crrMenu<>"" Then
								Fn_Mech_InitialValueTable=crrMenu
							else
								Fn_Mech_InitialValueTable=False
							End If
						End If
						If iRow<>"" Then
							objParameterDefinitionDialog.JavaTable("InitialValuesRows").ClickCell Cint(iRow)-1,0
						else
							objParameterDefinitionDialog.JavaTable("InitialValues").ClickCell 0,0
						End If
						Set objChld=Nothing
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			'Case to paste data in [ Initial Values ] table using Keyboard Ctrl+v
			Case "PasteData_Keyboard"
						'For this Case use [ StrValue ] parameter to set show ValueDescription option
						If StrValue<>"" Then
							Call Fn_CheckBox_Set("Fn_Mech_InitialValueTable", objParameterDefinitionDialog,"IntValueDescriptionCell",StrValue)
						End If
						If iRow<>"" and iCol<>"" Then
							objParameterDefinitionDialog.JavaTable("InitialValues").SelectCell CInt(iRow)-1,Cint(iCol)-1
							wait 1
							objParameterDefinitionDialog.JavaTable("InitialValues").PressKey "V",micCtrl
						Elseif iRow<>"" then
							objParameterDefinitionDialog.JavaTable("InitialValuesRows").SelectRow Cint(iRow)-1
							wait 1
							objParameterDefinitionDialog.JavaTable("InitialValuesRows").PressKey "V",micCtrl
						else
							objParameterDefinitionDialog.JavaTable("InitialValues").SelectCell 0,0
							wait 1
							objParameterDefinitionDialog.JavaTable("InitialValues").PressKey "V",micCtrl
						End If
						If Err.Number < 0 Then
							Fn_Mech_InitialValueTable=False
						Else
							Fn_Mech_InitialValueTable=True
						End If
						If iRow<>"" and iCol="" Then
							objParameterDefinitionDialog.JavaTable("InitialValues").DeselectRow Cint(iRow)-1
						End If
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			'Case to Set Collapse checkbox On or Off
			Case "Collapse_SetON","Collapse_SetOFF"			
				If StrAction="Collapse_SetON" Then
					Fn_Mech_InitialValueTable=Fn_CheckBox_Set("Fn_Mech_InitialValueTable", objParameterDefinitionDialog, "CollapseInitialValues","on")
				Else
					Fn_Mech_InitialValueTable=Fn_CheckBox_Set("Fn_Mech_InitialValueTable", objParameterDefinitionDialog, "CollapseInitialValues","off")
				End If
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			'Case to Undo Changes in [ Initial Values ] table
			Case "UndoChanges"
						objParameterDefinitionDialog.JavaTable("InitialValues").ClickCell 0,0,"RIGHT"
						wait 1
						objParameterDefinitionDialog.JavaMenu("index:=0","label:=Undo").Select
						If Err.Number < 0 Then
							Fn_Mech_InitialValueTable=False
						Else
							Fn_Mech_InitialValueTable=True
						End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			'Case to get all Disable Context menu for [ Maximum Values ] table
			Case "GetAllDisableContextMenu"
						If iRow<>"" Then
							objParameterDefinitionDialog.JavaTable("InitialValuesRows").SelectRow Cint(iRow)-1
							objParameterDefinitionDialog.JavaTable("InitialValuesRows").ClickCell Cint(iRow)-1,0,"RIGHT"
						else
							objParameterDefinitionDialog.JavaTable("InitialValues").ClickCell 0,0,"RIGHT"
						End If
						Set objMenu=Description.Create
						objMenu("Class Name").value="JavaMenu"
						Set objChld=objParameterDefinitionDialog.ChildObjects(objMenu)
						StrLabel=""
						For iCounter=0 to objChld.count-1
								crrMenu=objChld(iCounter).GetROProperty("label")
								If objChld(iCounter).CheckProperty("enabled",1,1)=false then
									If StrLabel="" Then
										StrLabel=objChld(iCounter).GetROProperty("label")
									else
										StrLabel=StrLabel+"~"+objChld(iCounter).GetROProperty("label")
									End If
								end if
						Next

						If Err.Number < 0 Then
							Fn_Mech_InitialValueTable=False
						Else
							If StrLabel<>"" Then
								Fn_Mech_InitialValueTable=StrLabel
							else
								Fn_Mech_InitialValueTable=False
							End If
						End If
						If iRow<>"" Then
							objParameterDefinitionDialog.JavaTable("InitialValuesRows").ClickCell Cint(iRow)-1,0
						else
							objParameterDefinitionDialog.JavaTable("InitialValues").ClickCell 0,0
						End If
						Set objChld=Nothing
				'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
			'Case to Verify Date Data
			Case "VerifyDateData"
				'Spliting row numbers
				aCol=Split(iCol,"~")
				aRow=Split(iRow,"~")
				aValue=Split(StrValue,"~")
				For iCounter=0 to ubound(aValue)
					bFlag=False
					cellval=objParameterDefinitionDialog.JavaTable("InitialValues").Object.getCellData(Cint(aRow(iCounter))-1,Cint(aCol(iCounter))-1).get(3).get(1).toString()
					aValDesc=Split(aValue(iCounter),":")
                    aDate=Split(cellval)

					If aDate(2)+"-"+aDate(1)+"-"+aDate(5)=aValDesc(0) or aDate(2)+"-0"+aDate(1)+"-"+aDate(5)=aValDesc(0) Then
						If ubound(aValDesc)=1 Then
							If trim(objParameterDefinitionDialog.JavaTable("InitialValues").Object.getCellData(Cint(aRow(iCounter))-1,Cint(aCol(iCounter))-1).get(2).toString())=aValDesc(1) Then
								bFlag=True		
							End If
						Else
							bFlag=True
						End If
					else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Value [ " & aValue(iCounter) & " ] is not exist in [ Initial Values ] table")
						Exit for
					End If
				Next
				If bFlag=True Then
					Fn_Mech_InitialValueTable=true
				End If
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to get specific Column name Exist in Table
		Case "GetColumnName"
				If iCol<>"" Then
					StrLabel=objParameterDefinitionDialog.JavaObject("InitialValuesTableHeader").Object.getColumnModel().getColumn(CInt(iCol)-1).getHeaderRenderer().getColName()
					Fn_Mech_InitialValueTable=StrLabel
				Else
					Fn_Mech_InitialValueTable=false
				end if
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to get specific row name Exist in Table
		Case "GetRowName"
				If iRow<>"" Then
					StrLabel=objParameterDefinitionDialog.JavaTable("InitialValuesRows").GetCellData(Cint(iRow)-1,0)
					Fn_Mech_InitialValueTable=StrLabel
				Else
					Fn_Mech_InitialValueTable=false
				end if
        '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "PasteRowsCols"           
			If iRow<>"" and iCol<>"" Then
				aRow=Split(iRow,"~")
				aCol=Split(iCol,"~")
				objParameterDefinitionDialog.JavaTable("InitialValues").SelectCellsRange cint(aRow(0))-1,cint(aCol(0))-1,cint(aRow(1))-1,Cint(aCol(1))-1
				objParameterDefinitionDialog.JavaTable("InitialValues").ClickCell Cint(aRow(0))-1,Cint(aCol(0))-1,"RIGHT"
			elseif iRow<>"" then
				aRow=Split(iRow,"~")
				objParameterDefinitionDialog.JavaTable("InitialValuesRows").SelectRow Cint(aRow(0))-1
				For iCounter=1 to ubound(aRow)
					objParameterDefinitionDialog.JavaTable("InitialValuesRows").ExtendRow CInt(aRow(iCounter))-1
				Next
				objParameterDefinitionDialog.JavaTable("InitialValues").ClickCell Cint(aRow(UBound(aRow)))-1,0,"RIGHT"
			End If
			wait 1
			objParameterDefinitionDialog.JavaMenu("index:=0","label:=Paste").Select
			wait 1
			If Err.Number < 0 Then
				Fn_Mech_InitialValueTable=False
			Else
				Fn_Mech_InitialValueTable=True
			End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			'Case to verify current value of cell is correct or not
			Case "IsCellCurrentValueCorrect"
				aCol=Split(iCol,"~")
				aRow=Split(iRow,"~")
				aValue=Split(StrValue,"~")
                For iCounter=0 to ubound(aCol)
						bFlag=false
						bFlag=objParameterDefinitionDialog.JavaTable("InitialValues").Object.getValueAt(cint(aRow(iCounter))-1,cint(aCol(iCounter))-1).isValueCorrect()
						If bFlag<>lcase(cstr(aValue(iCounter))) Then
							bFlag=false
							Exit for
						End If
				Next
				 If bFlag=false Then
					Fn_Mech_InitialValueTable=true
				else
					Fn_Mech_InitialValueTable=false
				End If
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			'Case to verify table foreground color
			Case "VerifyHeaderForegroundColour"
                aCol=Split(iCol,"~")
				aValue=Split(StrValue,"~")
                For iCounter=0 to ubound(aCol)
						bFlag=false
						sColourCode=""
						sColour=objParameterDefinitionDialog.JavaObject("InitialValuesTableHeader").Object.getColumnModel().getColumn(cint(aCol(iCounter))-1).getHeaderRenderer().getForeground().toString()
						sColour =  mid(sColour ,instr(sColour ,"[")  ,instr(sColour ,"]") )
						Select Case lCase(aValue(iCounter))
							Case "red"
								sColourCode="[r=255,g=0,b=0]"
						End Select
						If sColour=sColourCode Then
							bFlag=true
						else
							Exit for
						End if
				Next
                If bFlag=true Then
					Fn_Mech_InitialValueTable=true
				else
					Fn_Mech_InitialValueTable=false
				End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to select / Double click specific cell from tabel
		Case "SelectCell","DoubleClickCell"
				If StrValue<>"" Then
					Call Fn_CheckBox_Set("Fn_Mech_InitialValueTable", objParameterDefinitionDialog,"IntValueDescriptionCell",StrValue)
				End If
				Select Case StrAction
					Case "SelectCell"
						objParameterDefinitionDialog.JavaTable("InitialValues").ClickCell Cint(iCol)-1,Cint(iRow)-1
					Case "DoubleClickCell"
						objParameterDefinitionDialog.JavaTable("InitialValues").DoubleClickCell Cint(iCol)-1,Cint(iRow)-1
				End Select
				wait 2
				If Err.Number < 0 Then
					Fn_Mech_InitialValueTable=False
				Else
					Fn_Mech_InitialValueTable=True
				End If
	End Select
	set objParameterDefinitionDialog=nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_Mech_VerifyProperties

'Description			 :	Function Used to verify object properties

'Parameters			   :   1.StrAction: Action Name
'										2.StrLink: Property link [ All | General ]
'										3.dicProperties: Properties information
'										4.StrButtonName: Button Name
'
'Return Value		   : 	True / False

'Pre-requisite			:	Object should be selected

'Examples				:   dicProperties("Column")="1~2~3~1~2~3~1~2~3"
'										dicProperties("Row")="1~1~1~2~2~2~3~3~3"
'										dicProperties("Value")="150~90~268~28~36~42~52~68~88"/dicProperties("Value")="10.0:Ten~20.0:Twenty~30.0:Thirty~40.0:Fourty"
'										bReturn=Fn_Mech_VerifyProperties("MaximumValues","General",dicProperties,"Cancel")
'										
'										dicProperties("Column")="1~2~3~1~2~3~1~2~3"
'										dicProperties("Row")="1~1~1~2~2~2~3~3~3"
'										dicProperties("Value")="126~78~215~28~36~42~52~68~88"/dicProperties("Value")="10.0:Ten~20.0:Twenty~30.0:Thirty~40.0:Fourty"
'										bReturn= Fn_Mech_VerifyProperties("MinimumValues","General",dicProperties,"Cancel")
'										
'										dicProperties("Column")="1~2~3~1~2~3~1~2~3"
'										dicProperties("Row")="1~1~1~2~2~2~3~3~3"
'										dicProperties("Value")="126~78~215~28~36~42~52~68~88"/dicProperties("Value")="10.0:Ten~20.0:Twenty~30.0:Thirty~40.0:Fourty"
'										bReturn= Fn_Mech_VerifyProperties("InitialValues","General",dicProperties,"Cancel")
'										
'										dicProperties("PropertyName")="Row Labels"
'										dicProperties("Value")="ROW1~ROW2~ROW3"
'										bReturn= Fn_Mech_VerifyProperties("ListBox","General",dicProperties,"Close")
'										
'										dicProperties("PropertyName")="Current ID~Current Revision~Name"
'										dicProperties("Value")="000030~A~BCD"
'										bReturn= Fn_Mech_VerifyProperties("EditBox","All",dicProperties,"Cancel")
'
'										dicProperties("PropertyName")="Current ID~Current Revision~Name"
'										dicProperties("Value")="000030~A~BCD"
'										bReturn= Fn_Mech_VerifyProperties("StaticText","General",dicProperties,"Cancel")
'
'										dicProperties("Value")="ConRule1"
'										bReturn=Fn_Mech_VerifyProperties("ConversionRuleName","",dicProperties,"")
'
'										dicProperties("Value")="ConRuleDescription"
'										bReturn=Fn_Mech_VerifyProperties("ConversionRuleDescription","",dicProperties,"")
'
'										dicProperties("ConstantName")="A~B~C"
'										dicProperties("ConstantValue")="7~8~9"
'										bReturn=Fn_Mech_VerifyProperties("ConstantsTable","",dicProperties,"")
'
'										dicProperties("PropertyName")="Is Signed"
'										dicProperties("Value")="True"
'										bReturn=Fn_Mech_VerifyProperties("RadioButton","General",dicProperties,"")
'
'										bReturn=Fn_Mech_VerifyProperties("InitialValuesRowNames","","","")
'										bReturn=Fn_Mech_VerifyProperties("MaximumValuesRowNames","","","")
'										bReturn=Fn_Mech_VerifyProperties("MinimumValuesRowNames","","","")
'										bReturn=Fn_Mech_VerifyProperties("MinimumValuesRowNumbers","","","")
'										bReturn=Fn_Mech_VerifyProperties("InitialValuesRowNumbers","","","")
'										bReturn=Fn_Mech_VerifyProperties("MaximumValuesRowNumbers","","","")
'
'										dicProperties("Row")="1~2"
'										dicProperties("Column")="1~1"
'										dicProperties("Value")="04-Apr-2012~13-Apr-2012:Date3"
'										bReturn=Fn_Mech_VerifyProperties("DateMinimumValues","",dicProperties,"")
'
'										dicProperties("Value")="Ele1:36:Domain 1~Ele2:24:Domain 2"
'										dicProperties("Value")="Domain Element Name:Value:Description~Domain Element Name:Value:Description"
'										bRetrun=Fn_Mech_VerifyProperties("ValidValues","",dicProperties,"")
'													
'										dicProperties("DomainElementName")="Ele1"
'										dicProperties("Value")="0x36"
'										dicProperties("Description")="Domain 1"
'										bRetrun=Fn_Mech_VerifyProperties("SEDInitialValues","",dicProperties,"")
'
'										 to call [ BitDefination ] case Object should be Check Out
'										dicProperties("Value")="1:0:B1N8:0:1~2:7:B2N1:0:1~3:4:B3N4:0:1"
'										dicProperties("Value")="Byte number:Bit number:Bit name:Bit 0 meaning:Bit 1 meaning~Byte number:Bit number:Bit name:Bit 0 meaning:Bit 1 meaning~Byte number:Bit number:Bit name:Bit 0 meaning:Bit 1 meaning"
'										dicProperties("Value")="1:0:B1N8:0:1~2:7:B2N1:0:1~3:4:B3N4:0:1"
'										bReturn=Fn_Mech_VerifyProperties("BitDefination","",dicProperties,"")
'
'										dicProperties("ParameterName")="Intsingle~Int1DArr~Intsingle~Int1DArr"
'										dicProperties("ColumnName")="Parameter Values~Parameter Values~Type~Maximum Values"
'										dicProperties("Value")="27~{18,26,37}~ParmDefIntRevision~{20,30,40}"
'										bReturn=Fn_Mech_VerifyProperties("ParameterValuesTable","",dicProperties,"")
'
'									dicProperties("Row")="1"
'									dicProperties("Column")="2"
'									bReturn=Fn_Mech_VerifyProperties("MaximumValues_CopyData","General",dicProperties,"")
'									
'									bReturn=Fn_Mech_VerifyProperties("MinimumValues_CopyData","General",dicProperties,"")
'									
'									dicProperties("Row")="1"
'									bReturn=Fn_Mech_VerifyProperties("ValidValues_CopyData","General",dicProperties,"")
'									
'									bReturn=Fn_Mech_VerifyProperties("ValidValues_CopyData","General",dicProperties,"")
'
'									dicProperties("Byte")="1"
'									bReturn=Fn_Mech_VerifyProperties("BitDefination_CopyData","General",dicProperties,"")
'
'									dicProperties("Byte")="1~1~1~1"
'									dicProperties("BitNumber")="1~2~4~5"
'									bReturn=Fn_Mech_VerifyProperties("BitDefination_CopyData","General",dicProperties,"")
'
'									dicProperties("Row")= "1~3"
'									dicProperties("Row")= "Start row~End row"	
'									dicProperties("Column")= "1~3"
'									dicProperties("Column")= "Start column~End column"
'									bReturn=Fn_Mech_VerifyProperties("MaximumValues_CopyRowsCols","General",dicProperties,"")
'
'												dicProperties("DomainElementName")="Ele1~Ele2"
'												dicProperties("DomainElementValue")="0x1A~0x7F"
'												bReturn=Fn_Mech_VerifyProperties("SEDInitialValue_VerifyListData","",dicProperties,"")
'
'												dicEditProperties("PropertyName")="Action"
'												bReturn=Fn_Mech_VerifyProperties("PasteLink","",dicProperties,"")
'												
'												dicEditProperties("PropertyName")="Action"
'												bReturn=Fn_Mech_VerifyProperties("CopyLink","",dicProperties,"")
'												
'												dicEditProperties("PropertyName")="Action"
'												dicEditProperties("Value")="Copy~Paste~Clear"
'												bReturn=Fn_Mech_VerifyProperties("VerifyLinkMenu","",dicProperties,"")
'
'												dicProperties("PropertyName")="Type"
'												dicProperties("Value")=DataTable("ConRuleType", dtGlobalSheet)
'												dicProperties("CheckPropertyName")="value"
'												bReturn = Fn_Mech_VerifyProperties("ListBoxCheckProperty","General",dicProperties,"")
'
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												25-Apr-2012								1.0																						Sunny R
'													Sandeep N												27-Apr-2012								1.1					Added case : RadioButton				Sunny R
'													Sandeep N												30-Apr-2012								1.2					Added case : 										  Sunny R
'																																																	 InitialValuesRowNames,MaximumValuesRowNames,MinimumValuesRowNames
'																																																	 MinimumValuesRowNumbers,InitialValuesRowNumbers,MaximumValuesRowNumbers
'													Sandeep N												02-May-2012								1.3					Added case : 										  Sunny R
'																																																	 DateMinimumValues,DateMaximumValues,DateInitialValues		
'													Sandeep N												03-May-2012								1.4					Added case : BitDefination				   Sunny R
'													Sandeep N												14-May-2012								1.5					Added case : ParameterValuesTable				   Sunny R
'													Sandeep N												26-June-2012								1.6					Added case : BitDefination_CopyData				   Sunny R
'													Sandeep N												31-Jul-2012								1.7					Added case : InitialValues_CopyRowsCols				  Pranav I
'													Sandeep N												10-Aug-2012								1.8					Added case : SEDInitialValue_VerifyListData			Anjali M	  
'													pranav Ingle											05-Nov-2012									1.9				Added Case "StaticText"	 
'													pranav Ingle											07-Nov-2012									1.9				Added Case "DropDownMenu" & "EditBox_Edit"	 
'													Sandeep N												06-Dec-2012									1.10				Added Case "PasteLink" ,"CopyLink","ClearLink","VerifyLinkMenu"
'													Sandeep N												10-Dec-2012									1.11				Added Case "ListBoxCheckProperty"
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_Mech_VerifyProperties(StrAction,StrLink,dicProperties,StrButtonName)
	GBL_FAILED_FUNCTION_NAME="Fn_Mech_VerifyProperties"
 	'Declaring variables
	Dim objPropetiesDialog
	Dim aRowNumber,aColumnNumber,aValues,iCounter,bFlag,cellval,iEleCount,iCount,aProperty,iRows
	Dim aValidVal,aConstantName,aValDesc,aDate,aParameterName,aColumnName,aRow,aCol
	Dim aAction,tableName,max
	Dim aBitNumber,aByte
	Dim aDomainEleName,aDomainEleValue

	Fn_Mech_VerifyProperties=False
 	'Checking existance of [ Properties ] dialog
	If not JavaWindow("Mechatronics").JavaWindow("TcDefaultApplet").JavaDialog("Properties").Exist(6) Then
		Call Fn_MenuOperation("Select","View:Properties")
	End If
	Call Fn_ReadyStatusSync(1)
	'Creating object of [ Properties ] dialog
	Set objPropetiesDialog=JavaWindow("Mechatronics").JavaWindow("TcDefaultApplet").JavaDialog("Properties")
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Clicking on page link
	If StrLink="" Then
		objPropetiesDialog.JavaStaticText("PageLink").SetTOProperty "label","General"
		objPropetiesDialog.JavaStaticText("PageLink").Click 1,1,"LEFT"
	Else
		objPropetiesDialog.JavaStaticText("PageLink").SetTOProperty "label",StrLink
		objPropetiesDialog.JavaStaticText("PageLink").Click 1,1,"LEFT"
		If StrLink="All" Then
			'to click on [ Show empty properties... ] link
			wait 1
			objPropetiesDialog.JavaStaticText("PageLink").SetTOProperty "label","Show empty properties..."
			If objPropetiesDialog.JavaStaticText("PageLink").Exist(2) Then
				objPropetiesDialog.JavaStaticText("PageLink").Click 1,1,"LEFT"
			End If
		End If
	End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Select Case StrAction
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to verify values of Maximum Value Table
		Case "InitialValues","MaximumValues","MinimumValues"
				'Spliting row numbers
				aRowNumber=Split(dicProperties("Row"),"~")
				aColumnNumber=Split(dicProperties("Column"),"~")
				aValues=Split(dicProperties("Value"),"~")
				For iCounter=0 to ubound(aRowNumber)
					bFlag=False
'					cellval=objPropetiesDialog.JavaTable(StrAction).Object.getCellData(Cint(aRowNumber(iCounter))-1,Cint(aColumnNumber(iCounter))-1).toString()
'					If instr(1,cellval,"[") Then cellval=Replace(cellval,"[","") end if
'					If instr(1,cellval,"]") Then cellval=Replace(cellval,"]","") end if
'					cellval=Split(cellval,",")
					aValDesc=Split(aValues(iCounter),":")
					If aValDesc(0) = "0.0" Then
						aValDesc(0) = Cint(aValDesc(0))
						cellval = CInt(trim(objPropetiesDialog.JavaTable(StrAction).Object.getCellData(Cint(aRowNumber(iCounter))-1,Cint(aColumnNumber(iCounter))-1).get(3).get(1).toString()))
					Else
						cellval = objPropetiesDialog.JavaTable(StrAction).Object.getCellData(Cint(aRowNumber(iCounter))-1,Cint(aColumnNumber(iCounter))-1).get(3).get(1).toString()
					End If
'					If trim(cellval(ubound(cellval)))=aValDesc(0) Then
					If trim(cellval)=trim(aValDesc(0)) Then
						If ubound(aValDesc)=1 Then
'							If trim(cellval(2))=aValDesc(1) Then
							If trim(objPropetiesDialog.JavaTable(StrAction).Object.getCellData(Cint(aRowNumber(iCounter))-1,Cint(aColumnNumber(iCounter))-1).get(2).toString())=aValDesc(1) Then
								bFlag=True		
							End If
						Else
							bFlag=True
						End If
					else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Value [ " & aValues(iCounter) & " ] is not exist in [ " & StrAction & " ] table")
						Exit for
					End If
				Next
				If bFlag=True Then
					Fn_Mech_VerifyProperties=true
				End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to verify values from list Box
		'to verify values of List box list box should be enabled so if want to verify values from List box object should be check out
		Case "ListBox"
				objPropetiesDialog.JavaStaticText("PropertyLabel").SetTOProperty "label",dicProperties("PropertyName")+":"
				If objPropetiesDialog.JavaList("ListBox").Exist(2) Then
					aValues=Split(dicProperties("Value"),"~")
					For iCounter=0 to uBound(aValues)
						bFlag=false
						'Verifying value exist in list or not
						'taking item count from list
						iEleCount=Fn_UI_Object_GetROProperty("Fn_Mech_VerifyProperties",objPropetiesDialog.JavaList("ListBox"), "items count")
						For iCount=0 to iEleCount-1
							If objPropetiesDialog.JavaList("ListBox").GetItem(iCount)=aValues(iCounter) Then
								bFlag=true
								Exit for
							End If
						Next
						If bFlag=false Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Value [ " & aValues(iCounter) & " ] is not exist in [ " & dicProperties("PropertyName") & " ] List")
							Exit for
						End If
					Next
					If bFlag=True Then
						Fn_Mech_VerifyProperties=true
					End If
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"List box [ " & dicProperties("PropertyName") & " ] is not exist on dialog")
				End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to verify property value from edit boxes
		Case "EditBox"
				aProperty=Split(dicProperties("PropertyName"),"~")
				aValues=Split(dicProperties("Value"),"~")
				For iCounter=0 to UBound(aProperty)
					bFlag=False
					objPropetiesDialog.JavaStaticText("PropertyLabel").SetTOProperty "label",aProperty(iCounter)+":"
					If objPropetiesDialog.JavaEdit("EditBox").Exist(3) Then
						If Fn_Edit_Box_GetValue("Fn_Mech_VerifyProperties",objPropetiesDialog,"EditBox")=aValues(iCounter) Then
							bFlag=True
						End If
					else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Property [ " & aProperty(iCounter) & " ] is not exist on dialog")
						Exit for
					End If
					If bFlag=False Then
						Exit for
					End If
				Next
				If bFlag=True Then
					Fn_Mech_VerifyProperties=true
				End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to modify property value from edit boxes
		Case "EditBox_Edit"															' 	Added For Editing edit box values from properties Dialog - Pranav Ingle
				aProperty=Split(dicProperties("PropertyName"),"~")
				aValues=Split(dicProperties("Value"),"~")

				For iCounter=0 to UBound(aProperty)
					objPropetiesDialog.JavaStaticText("PropertyLabel").SetTOProperty "label",aProperty(iCounter)+":"

					If objPropetiesDialog.JavaEdit("EditBox").Exist(3) Then
						Call Fn_Edit_Box("Fn_Mech_VerifyProperties",objPropetiesDialog,"EditBox",aValues(iCounter)+ vbLf + "")					
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Property [ " & aProperty(iCounter) & " ] is not exist on dialog")
						Exit function
					End If
				Next
				Fn_Mech_VerifyProperties=true
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to verify property value from edit boxes
		Case "DropDownMenu"  									'-   Added To Modify Header ObjectAnd Trailer Object Propeties  - Pranav Ingle
				aProperty=Split(dicProperties("PropertyName"),"~")
				aValues=Split(dicProperties("Value"),"~")

				bFlag=true
				For iCounter=0 to UBound(aProperty)

					objPropetiesDialog.JavaStaticText("PropertyLabel").SetTOProperty "label",aProperty(iCounter)+":"
					'objPropetiesDialog.JavaObject("HeaderObject").Click 1,1
					objPropetiesDialog.JavaStaticText("DropDownButton").Click 1,1
					wait 1
										
					objPropetiesDialog.JavaMenu("label:=" & aValues(iCounter)).Click 1,1

					If Err.Number < 0 Then
						bFlag=false
					End If
				Next

				If bFlag=True Then
					Fn_Mech_VerifyProperties=true
				End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to verify property value from edit boxes
		Case "StaticText"  									'-   Added To Verify Header ObjectAnd Trailer Object Propeties  - Pranav Ingle
				aProperty=Split(dicProperties("PropertyName"),"~")
				aValues=Split(dicProperties("Value"),"~")
				For iCounter=0 to UBound(aProperty)
					bFlag=False
					objPropetiesDialog.JavaStaticText("PropertyLabel").SetTOProperty "label",aProperty(iCounter)+":"
					If objPropetiesDialog.JavaStaticText("HeaderObjectValue").Exist(3) Then
						If objPropetiesDialog.JavaStaticText("HeaderObjectValue").GetROProperty("label") = aValues(iCounter) Then
							bFlag=True
						End If
					else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Property [ " & aProperty(iCounter) & " ] is not exist on dialog")
						Exit for
					End If
					If bFlag=False Then
						Exit for
					End If
				Next

				If bFlag=True Then
					Fn_Mech_VerifyProperties=true
				End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to verify conversion rule name,Description
		Case "ConversionRuleName","ConversionRuleDescription"
				bFlag=false
				If Fn_Edit_Box_GetValue("Fn_Mech_VerifyProperties",objPropetiesDialog,StrAction)=dicProperties("Value") Then
					bFlag=True
				End If
				If bFlag=True Then
					Fn_Mech_VerifyProperties=true
				End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to verify values from Constants Table
		Case "ConstantsTable"
				aConstantName=Split(dicProperties("ConstantName"),"~")
				aValues=Split(dicProperties("ConstantValue"),"~")
				iRows=objPropetiesDialog.JavaTable("Constants").GetROProperty("rows")
				For iCounter=0 to ubound(aConstantName)
					bFlag=False
					For iCount=0 to iRows-1
						cellval=objPropetiesDialog.JavaTable("Constants").GetCellData(iCount,"Constant Value")
						If trim(cellval)=aValues(iCounter) and objPropetiesDialog.JavaTable("Constants").GetCellData(iCount,"Constant Name")=aConstantName(iCounter) Then
							bFlag=True
							Exit for
						End If
					Next
					If bFlag=false Then
						Exit for
					End If
				Next
				If bFlag=True Then
					Fn_Mech_VerifyProperties=true
				End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to verify Radio Buttons
		Case "RadioButton"
				aProperty=Split(dicProperties("PropertyName"),"~")
				aValues=Split(dicProperties("Value"),"~")
				For iCounter=0 to UBound(aProperty)
					bFlag=False
					objPropetiesDialog.JavaStaticText("PropertyLabel").SetTOProperty "label",aProperty(iCounter)+":"
					objPropetiesDialog.JavaRadioButton("RadioButton").SetTOProperty "attached text",aValues(iCounter)

					If objPropetiesDialog.JavaRadioButton("RadioButton").Exist(3) Then
						If Fn_UI_Object_GetROProperty("Fn_Mech_VerifyProperties",objPropetiesDialog.JavaRadioButton("RadioButton"), "value")=1 Then
							bFlag=True
						End If
					else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Property [ " & aProperty(iCounter) & " ] is not exist on dialog")
						Exit for
					End If
					If bFlag=False Then
						Exit for
					End If
				Next
				If bFlag=True Then
					Fn_Mech_VerifyProperties=true
				End If
        ' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get all Row names of Tables"InitialValues","MaximumValues","MinimumValues"
		Case "InitialValuesRowNames","MaximumValuesRowNames","MinimumValuesRowNames"
				Select Case StrAction
					Case "InitialValuesRowNames"
						cellval="Initial Values:"
					Case "MaximumValuesRowNames"
						cellval="Maximum Values:"
					Case "MinimumValuesRowNames"
						cellval="Minimum Values:"
				End Select
				objPropetiesDialog.JavaStaticText("PropertyLabel").SetTOProperty "label",cellval
				If objPropetiesDialog.JavaTable("RowNames").Exist(3) Then
					iEleCount=Fn_UI_Object_GetROProperty("Fn_Mech_VerifyProperties",objPropetiesDialog.JavaTable("RowNames"), "rows")
					For iCount=0 to iEleCount-1
						If iCount=0 Then
							aValues=objPropetiesDialog.JavaTable("RowNames").GetCellData(iCount,0)
						else
							aValues=aValues+"~"+objPropetiesDialog.JavaTable("RowNames").GetCellData(iCount,0)
						End If
					Next
					Fn_Mech_VerifyProperties=aValues
				End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get all Row numbers of Tables"InitialValues","MaximumValues","MinimumValues"
		Case "InitialValuesRowNumbers","MaximumValuesRowNumbers","MinimumValuesRowNumbers"
				Select Case StrAction
					Case "InitialValuesRowNumbers"
						cellval="Initial Values:"
					Case "MaximumValuesRowNumbers"
						cellval="Maximum Values:"
					Case "MinimumValuesRowNumbers"
						cellval="Minimum Values:"
				End Select
				objPropetiesDialog.JavaStaticText("PropertyLabel").SetTOProperty "label",cellval
				If objPropetiesDialog.JavaList("RowNumbers").Exist(3) Then
					iEleCount=Fn_UI_Object_GetROProperty("Fn_Mech_VerifyProperties",objPropetiesDialog.JavaList("RowNumbers"),"items count")
					For iCount=0 to iEleCount-1
						If iCount=0 Then
							aValues=CStr(objPropetiesDialog.JavaList("RowNumbers").GetItem(iCount))
						else
							aValues=CStr(aValues)+"~"+CStr(objPropetiesDialog.JavaList("RowNumbers").GetItem(iCount))
						End If
					Next
					Fn_Mech_VerifyProperties=aValues
				End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to verify Date values of Maximum Value Table
		Case "DateInitialValues","DateMaximumValues","DateMinimumValues"
				Select Case StrAction
					Case "DateInitialValues"
						StrAction="InitialValues"
					Case "DateMaximumValues"
						StrAction="MaximumValues"
					Case "DateMinimumValues"
						StrAction="MinimumValues"
				End Select
				'Spliting row numbers
				aRowNumber=Split(dicProperties("Row"),"~")
				aColumnNumber=Split(dicProperties("Column"),"~")
				aValues=Split(dicProperties("Value"),"~")
				For iCounter=0 to ubound(aRowNumber)
					bFlag=False
					cellval=objPropetiesDialog.JavaTable(StrAction).Object.getCellData(Cint(aRowNumber(iCounter))-1,Cint(aColumnNumber(iCounter))-1).get(3).toString()
					
					If instr(1,cellval,"[") Then cellval=Replace(cellval,"[","") end if
					If instr(1,cellval,"]") Then cellval=Replace(cellval,"]","") end if
					
					aValDesc=Split(aValues(iCounter),":")
					aDate=Split(cellval)

					If aDate(3)+"-"+aDate(2)+"-"+aDate(6)=aValDesc(0) or aDate(3)+"-0"+aDate(2)+"-"+aDate(6)=aValDesc(0) Then
						If ubound(aValDesc)=1 Then
							If trim(objPropetiesDialog.JavaTable(StrAction).Object.getCellData(Cint(aRowNumber(iCounter))-1,Cint(aColumnNumber(iCounter))-1).get(2).toString())=aValDesc(1) Then
								bFlag=True		
							End If
						Else
							bFlag=True
						End If
					else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Value [ " & aValues(iCounter) & " ] is not exist in [ " & StrAction & " ] table")
						Exit for
					End If
				Next
				If bFlag=True Then
					Fn_Mech_VerifyProperties=true
				End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to verify Initial values from SED Initial values table
		Case "SEDInitialValues"
			bFlag=False
			'verifying Domain Element Name 
			If dicProperties("DomainElementName")<>"" Then
				If dicProperties("DomainElementName")=objPropetiesDialog.JavaTable("SEDInitialValue").GetCellData(0,"Domain Element Name") Then
					bFlag=true
				End If
			Else
				bFlag=true
			End If
			If bFlag=False Then
				Exit function
			End If
			bFlag=False
			'verifying Value 
			If dicProperties("Value")<>"" Then
				If dicProperties("Value")=objPropetiesDialog.JavaTable("SEDInitialValue").GetCellData(0,"Value") Then
					bFlag=true
				End If
			Else
				bFlag=true
			End If
			If bFlag=False Then
				Exit function
			End If
			bFlag=False
			'verifying Description 
			If dicProperties("Description")<>"" Then
				If dicProperties("Description")=objPropetiesDialog.JavaTable("SEDInitialValue").GetCellData(0,"Description") Then
					bFlag=true
				End If
			Else
				bFlag=true
			End If
			If bFlag=true Then
				Fn_Mech_VerifyProperties=true
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to verify Valid values from SED Valid values table
		Case "ValidValues"
				aValues=Split(dicProperties("Value"),"~")
				For iCounter=0 to ubound(aValues)
                	aValDesc=Split(aValues(iCounter),":")
					bFlag=False
					For iCount=0 to cint(objPropetiesDialog.JavaTable("ValidValues").GetROProperty("rows"))-1
						If aValDesc(0)=objPropetiesDialog.JavaTable("ValidValues").Object.getCellData(iCount,0).get(2).toString() Then
							bFlag=true
							If aValDesc(1)<>"" Then
								If aValDesc(1)<>objPropetiesDialog.JavaTable("ValidValues").Object.getCellData(iCount,0).get(3).get(1).toString() Then
									bFlag=false
									Exit for
								end if
							End If
							If aValDesc(2)<>"" Then
								If aValDesc(2)<>objPropetiesDialog.JavaTable("ValidValues").Object.getCellData(iCount,0).get(4).get(1).toString() Then
									bFlag=false
									Exit for
								end if
							End If
						End If
					Next
					If bFlag=false Then
						Exit for
					End If
				Next
				If bFlag=true Then
					Fn_Mech_VerifyProperties=true
				End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'to verify values from this Bit Defination table Object should be Check Out
		'Case to verify values from Bit Defination table
		Case "BitDefination"
			'dicProperties("Value")="Byte number:Bit number:Bit name:Bit 0 meaning:Bit 1 meaning~Byte number:Bit number:Bit name:Bit 0 meaning:Bit 1 meaning"
			aValues=Split(dicProperties("Value"),"~")
			For iCounter=0 to ubound(aValues)
				bFlag=false
				aValDesc=Split(aValues(iCounter),":")
				
				If objPropetiesDialog.JavaCheckBox("BitDefinitionTableCollapse").Exist(2) Then
					If objPropetiesDialog.JavaCheckBox("BitDefinitionTableCollapse").GetROProperty("enabled")=1 Then
						Call Fn_CheckBox_Set("Fn_Mech_VerifyProperties", objPropetiesDialog, "BitDefinitionTableCollapse", "on")
					End If
				End If
				iRows=Cint(aValDesc(0))*8-CInt(aValDesc(1))-1
				'verifing Bit name
                If aValDesc(2)<>"" Then
                    If trim(objPropetiesDialog.JavaTable("CCDMBitDefTable").GetCellData(iRows,"Name"))=aValDesc(2) then
						bFlag=true
					end if
				else
					bFlag=true
				End If
				If bFlag=false Then
					Exit for
				End If
				bFlag=false
				'verifing Bit 0 meaning
                If aValDesc(3)<>"" Then
					If trim(objPropetiesDialog.JavaTable("CCDMBitDefTable").GetCellData(iRows,"""0"" Meaning"))=aValDesc(3) then
						bFlag=true
					end if
				else
					bFlag=true
				End If
				If bFlag=false Then
					Exit for
				End If
				bFlag=false
				'verifing Bit 1 meaning
                If aValDesc(4)<>"" Then
					If trim(objPropetiesDialog.JavaTable("CCDMBitDefTable").GetCellData(iRows,"""1"" Meaning"))=aValDesc(4) then
						bFlag=true
					end if
				else
					bFlag=true
				End If
				If bFlag=false Then
					Exit for
				End If
			Next
			If bFlag=true Then
				Fn_Mech_VerifyProperties=true
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to verify values from Parameter Values table
		 Case "ParameterValuesTable"
				aParameterName=Split(dicProperties("ParameterName"),"~")
				aValues=Split(dicProperties("Value"),"~")
				aColumnName=Split(dicProperties("ColumnName"),"~")

				For iCount=0 to uBound(aParameterName)
					bFlag=false
					If not isNumeric(aParameterName(iCount)) Then
						bFlag=false
						iRows=Fn_UI_Object_GetROProperty("Fn_Mech_VerifyProperties",objPropetiesDialog.JavaTable("ParameterValues"),"rows")
						For iCounter=0 to iRows-1
							cellval=objPropetiesDialog.JavaTable("ParameterValues").GetCellData(iCounter,"Name")
							If trim(cellval)=trim(aParameterName(iCount)) Then
								If trim(objPropetiesDialog.JavaTable("ParameterValues").GetCellData(iCounter,aColumnName(iCount)))=trim(aValues(iCount)) Then
									bFlag=true
									Exit for
								End If
							End If
						Next
					else
						iCounter=cInt(aParameterName(iCount))-1
						If trim(objPropetiesDialog.JavaTable("ParameterValues").GetCellData(iCounter,aColumnName(iCount)))=trim(aValues(iCount)) Then
							bFlag=true
						End if
					End If
					If bFlag=false Then
						Exit for
					End If
				Next
				If bFlag=true Then
					Fn_Mech_VerifyProperties=true
				End If
		 '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to Copy data from [ InitialValues , Maximum Values , MinimumValues ] table
		Case "InitialValues_CopyData","MaximumValues_CopyData","MinimumValues_CopyData"
				If objPropetiesDialog.JavaSlider("JScrollPane_2").exist(4) Then
					max=objPropetiesDialog.JavaSlider("JScrollPane_2").GetROProperty("max")
					objPropetiesDialog.JavaSlider("JScrollPane_2").Drag max
					wait 1
				else
					max=objPropetiesDialog.JavaSlider("JScrollPane").GetROProperty("max")
					objPropetiesDialog.JavaSlider("JScrollPane").Drag max
					wait 1
				End If

				aAction=Split(StrAction,"_")
				tableName=aAction(0)
				objPropetiesDialog.JavaTable(tableName).Object.setEnabled True	'changed by shweta rathod
				'Setting [ ShowValueDescription ] option
				If dicProperties("ShowValueDescription")<>"" Then
					Call Fn_CheckBox_Set("Fn_Mech_VerifyProperties", objPropetiesDialog, tableName&"ValueDescription",dicProperties("ShowValueDescription"))
				End If
				' - - - - - Selecting Specified Cell or Row- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				If dicProperties("Row")<>"" and dicProperties("Column")<>"" Then
					objPropetiesDialog.JavaTable(tableName).ClickCell Cint(dicProperties("Row"))-1,Cint(dicProperties("Column"))-1,"RIGHT"
				elseif dicViewerTabInfo("Row")<>"" then
					objPropetiesDialog.JavaTable(tableName&"Rows").SelectRow Cint(dicProperties("Row"))-1
					objPropetiesDialog.JavaTable(tableName&"Rows").ClickCell Cint(dicProperties("Row"))-1,0,"RIGHT"
				else
					objPropetiesDialog.JavaTable(tableName).PressKey "A",micCtrl
					wait 1
					objPropetiesDialog.JavaTable(tableName).ClickCell 0,0,"RIGHT"
				End If
				'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				'Selecting menu [ Copy ]  data from Table
				objPropetiesDialog.JavaMenu("index:=0","label:=Copy").Select
				If Err.Number < 0 Then
					Fn_Mech_VerifyProperties=False
				Else
					Fn_Mech_VerifyProperties=True
				End If
				If dicProperties("Row")<>"" and dicProperties("Column")="" Then
					objPropetiesDialog.JavaTable(tableName).DeselectRow Cint(dicProperties("Row"))-1
				End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to Copy data from [ ISED Valid Values ] table
		Case "ValidValues_CopyData"
			max=objPropetiesDialog.JavaSlider("JScrollPane").GetROProperty("max")
				objPropetiesDialog.JavaSlider("JScrollPane").Drag max
				wait 1
				objPropetiesDialog.JavaTable("ValidValues").Object.setEnabled True
				aAction=Split(StrAction,"_")
				tableName=aAction(0)
				' - - - - - Selecting Specified Cell or Row- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				If dicProperties("Row")<>"" Then
					aRow = Split(dicProperties("Row"),"~")
                    objPropetiesDialog.JavaTable(tableName).SelectRow Cint(aRow(0))-1
					For iCounter = 1 To Ubound(aRow)
						objPropetiesDialog.JavaTable(tableName).ExtendRow Cint(aRow(iCounter))-1
					Next
					objPropetiesDialog.JavaTable(tableName).ClickCell Cint(aRow(0))-1,0,"RIGHT"

				else
					objPropetiesDialog.JavaTable(tableName).PressKey "A",micCtrl
					wait 1
					objPropetiesDialog.JavaTable(tableName).ClickCell 0,0,"RIGHT"
				End If
				'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				'Selecting menu [ Paste ]  to paste copied data
				objPropetiesDialog.JavaMenu("index:=0","label:=Copy").Select
				If Err.Number < 0 Then
					Fn_Mech_VerifyProperties=False
				Else
					Fn_Mech_VerifyProperties=True
				End If
				If dicProperties("Row")<>"" Then
					objPropetiesDialog.JavaTable(tableName).SelectRow Cint(aRow(0))-1
					objPropetiesDialog.JavaTable(tableName).DeselectRow Cint(aRow(0))-1
				End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to Copy data from [ BitDefination ] table
		Case "BitDefination_CopyData"
				max=objPropetiesDialog.JavaSlider("JScrollPane").GetROProperty("max")
				objPropetiesDialog.JavaSlider("JScrollPane").Drag max
				wait 1
				' - - - - - Selecting Specified Cell or Row- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				If dicProperties("Byte")<>"" Then

					objPropetiesDialog.JavaTable("CCDMBitDefRowHeaderTable").SelectRow Cint(dicProperties("Byte"))-1
					objPropetiesDialog.JavaTable("CCDMBitDefRowHeaderTable").ClickCell Cint(dicProperties("Byte"))-1,0,"RIGHT"

				elseif dicProperties("BitNumber")<>"" and  dicProperties("Byte")<>"" then

					aBitNumber=Split(dicProperties("BitNumber"),"~")
					aByte=Split(dicProperties("Byte"),"~")
					iRows=Cint(aByte(0))*8-CInt(aBitNumber(1))-1
					objPropetiesDialog.JavaTable("CCDMBitDefTable").SelectRow iRows
					For iCount=1 to ubound(aByte)
						iRows=Cint(aByte(iCount))*8-CInt(aBitNumber(iCount))-1
						objPropetiesDialog.JavaTable("CCDMBitDefTable").ExtendRow iRows	
					Next
					objPropetiesDialog.JavaTable("CCDMBitDefTable").Click 1,1,"RIGHT"

				End If
				'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				'Selecting menu [ Paste ]  to paste copied data
				objPropetiesDialog.JavaMenu("index:=0","label:=Copy").Select
				If Err.Number < 0 Then
					Fn_Mech_VerifyProperties=False
				Else
					Fn_Mech_VerifyProperties=True
				End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to Copy specific data from [ InitialValues , Maximum Values , MinimumValues ] table
		Case "InitialValues_CopyRowsCols","MaximumValues_CopyRowsCols","MinimumValues_CopyRowsCols"
				max=objPropetiesDialog.JavaSlider("JScrollPane").GetROProperty("max")
				objPropetiesDialog.JavaSlider("JScrollPane").Drag max
				wait 1						
				aAction=Split(StrAction,"_")
				tableName=aAction(0)
				objPropetiesDialog.JavaTable(tableName).Object.setEnabled True	'changed by shweta rathod
				'Setting [ ShowValueDescription ] option
				If dicProperties("ShowValueDescription")<>"" Then
					Call Fn_CheckBox_Set("Fn_Mech_VerifyProperties", objPropetiesDialog, tableName&"ValueDescription",dicProperties("ShowValueDescription"))
				End If

				If dicProperties("Row")<>"" and dicProperties("Column")<>"" Then
					aRow=Split(dicProperties("Row"),"~")
					aCol=Split(dicProperties("Column"),"~")
					objPropetiesDialog.JavaTable(tableName).SelectCellsRange cint(aRow(0))-1,cint(aCol(0))-1,cint(aRow(1))-1,Cint(aCol(1))-1
					objPropetiesDialog.JavaTable(tableName).ClickCell Cint(aRow(0))-1,Cint(aCol(0))-1,"RIGHT"
				elseif dicProperties("Row")<>"" then
					aRow=Split(dicProperties("Row"),"~")
					objPropetiesDialog.JavaTable(tableName&"Rows").SelectRow Cint(aRow(0))-1
				For iCounter=1 to ubound(aRow)
					objPropetiesDialog.JavaTable(tableName&"Rows").ExtendRow CInt(aRow(iCounter))-1
				Next
				objPropetiesDialog.JavaTable(tableName).ClickCell Cint(aRow(UBound(aRow)))-1,0,"RIGHT"
			End If
            objPropetiesDialog.JavaMenu("index:=0","label:=Copy").Select
			If Err.Number < 0 Then
				Fn_Mech_VerifyProperties=False
			Else
				Fn_Mech_VerifyProperties=True
			End If
		 '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to verify data from [ SED Initial Values ] List
        Case "SEDInitialValue_VerifyListData"
			Fn_Mech_VerifyProperties=False
			max=objPropetiesDialog.JavaSlider("JScrollPane_2").GetROProperty("max")
			objPropetiesDialog.JavaSlider("JScrollPane_2").Drag max
			wait 1

			If dicProperties("DomainElementName")<>"" Then
					aDomainEleName=Split(dicProperties("DomainElementName"),"~")
					objPropetiesDialog.JavaTable("SEDInitialValue").ClickCell 0,"Domain Element Name"
					wait 1
					For iCounter=0 to ubound(aDomainEleName)
						bFlag=Fn_UI_ListItemExist("Fn_Mech_VerifyProperties", objPropetiesDialog, "InitialValuesList",aDomainEleName(iCounter))
						If bFlag=false Then
							objPropetiesDialog.JavaTable("SEDInitialValue").ClickCell 0,"Description"
							Exit function
							'Releasing Object of Java Applet [ JApplet ]
							Set objPropetiesDialog=Nothing
						End If
					Next
			End If
			objPropetiesDialog.JavaTable("SEDInitialValue").ClickCell 0,"Description"
			wait 1
			If dicProperties("DomainElementValue")<>"" Then
					aDomainEleValue=Split(dicProperties("DomainElementValue"),"~")
					objPropetiesDialog.JavaTable("SEDInitialValue").ClickCell 0,"Value"
					wait 1
					For iCounter=0 to ubound(aDomainEleValue)
						bFlag=Fn_UI_ListItemExist("Fn_Mech_VerifyProperties", objPropetiesDialog, "InitialValuesList",aDomainEleValue(iCounter))
						If bFlag=false Then
							objPropetiesDialog.JavaTable("SEDInitialValue").ClickCell 0,"Description"
							Exit function
							'Releasing Object of Java Applet [ JApplet ]
							Set objPropetiesDialog=Nothing
						End If
					Next
			End if
			Fn_Mech_VerifyProperties=True
			objPropetiesDialog.JavaTable("SEDInitialValue").ClickCell 0,"Description"
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to get specific property state of specific Edit box : e.g { current value, editable state, enabled state }
		Case "EditBox_GetPropertyState"
				objPropetiesDialog.JavaStaticText("PropertyLabel").SetTOProperty "label",dicProperties("PropertyName")+":"
				If objPropetiesDialog.JavaEdit("EditBox").Exist(3) Then
					Fn_Mech_VerifyProperties=objPropetiesDialog.JavaEdit("EditBox").GetROProperty(dicProperties("PropertyState"))
				else
					Fn_Mech_VerifyProperties=false
				End if
		 '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "VerifyLinkMenu"
				max=objPropetiesDialog.JavaSlider("JScrollPane").GetROProperty("max")
				objPropetiesDialog.JavaSlider("JScrollPane").Drag max
				wait 1
				objPropetiesDialog.JavaStaticText("PropertyLabel").SetTOProperty "label",dicProperties("PropertyName")+":"
				If objPropetiesDialog.JavaStaticText("DropDownButton").Exist(2) Then
					objPropetiesDialog.JavaStaticText("DropDownButton").Click 1, 1
				Else
					objPropetiesDialog.JavaObject("LinkOptionDropDown").Click 1,1	
				End If
				wait 1
                aValues=Split(dicProperties("Value"),"~")
				For iCounter=0 to UBound(aValues)
					bFlag=True
					If not objPropetiesDialog.JavaMenu("index:=0","label:=" & aValues(iCounter)).Exist(2) Then
						bFlag=False
						Exit for
					End If
				Next
				If bFlag=True Then
					Fn_Mech_VerifyProperties=True
				Else
					Fn_Mech_VerifyProperties=False
				End If
				If objPropetiesDialog.JavaStaticText("DropDownButton").Exist(2) Then
					objPropetiesDialog.JavaStaticText("DropDownButton").Click 1, 1
				Else
					objPropetiesDialog.JavaObject("LinkOptionDropDown").Click 1,1	
				End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "CopyLink","PasteLink","ClearLink"
				max=objPropetiesDialog.JavaSlider("JScrollPane").GetROProperty("max")
				objPropetiesDialog.JavaSlider("JScrollPane").Drag max
				wait 1
				objPropetiesDialog.JavaStaticText("PropertyLabel").SetTOProperty "label",dicProperties("PropertyName")+":"
				
				If objPropetiesDialog.JavaStaticText("DropDownButton").Exist(1) Then
					objPropetiesDialog.JavaStaticText("DropDownButton").Click 10, 5,"LEFT"
				ElseIf objPropetiesDialog.JavaObject("LinkOptionDropDown").Exist(1) Then	
					objPropetiesDialog.JavaObject("LinkOptionDropDown").Click 1,1				
				End If
				
				
				wait 1
				' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 					
				Select Case StrAction
					' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
					Case "CopyLink"
						objPropetiesDialog.JavaMenu("index:=0","label:=Copy").Select
					' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
					Case "PasteLink"
						objPropetiesDialog.JavaMenu("index:=0","label:=Paste").Select
					' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
					Case "ClearLink"
						objPropetiesDialog.JavaMenu("index:=0","label:=Clear").Select
					' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				End Select

				If Err.Number < 0 Then
					Fn_Mech_VerifyProperties=False
				Else
					Fn_Mech_VerifyProperties=True
				End If
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			Case "ListBoxCheckProperty"
				objPropetiesDialog.JavaStaticText("PropertyLabel").SetTOProperty "label",dicProperties("PropertyName")+":"
				If objPropetiesDialog.JavaList("ListBox").Exist(2) Then
					Fn_Mech_VerifyProperties = objPropetiesDialog.JavaList("ListBox").CheckProperty(dicProperties("CheckPropertyName"),dicProperties("Value"))
				End IF
	End	Select
	'Clicking on button
	If StrButtonName<>"" Then
		Call Fn_Button_Click("Fn_Mech_VerifyProperties", objPropetiesDialog,StrButtonName)
	End If
	'Releasing object of [ Properties ] dialog
	Set objPropetiesDialog=nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_Mech_EditProperties

'Description			 :	Function Used to verify object properties

'Parameters			   :   1.StrAction: Action Name
'										2.StrLink: Property link [ All | General ]
'										3.dicEditProperties: Properties information
'										4.StrButtonName: Button Name
'
'Return Value		   : 	True / False

'Pre-requisite			:	Object should be selected

'Examples				:   dicEditProperties("Row")="1~1~1~2~2~2~3~3~3"
'										dicEditProperties("Column")="1~2~3~1~2~3~1~2~3"
'										dicEditProperties("Value")="600~700~800~900~1000~1100~1200~1300~1400"/dicEditProperties("Value")="10.0:M_Ten~20.0:M_Twenty~:M_Thirty~50.0:Fifty"
'										bReturn= Fn_Mech_EditProperties("MaximumValues","",dicEditProperties,"Save")
'										
'										dicEditProperties("Row")="1~1~1~2~2~2~3~3~3"
'										dicEditProperties("Column")="1~2~3~1~2~3~1~2~3"
'										dicEditProperties("Value")="200~300~400~500~600~700~800~900~1000"/dicEditProperties("Value")="10.0:M_Ten~20.0:M_Twenty~:M_Thirty~50.0:Fifty"
'										bReturn= Fn_Mech_EditProperties("MinimumValues","",dicEditProperties,"Save")
'										
'										dicEditProperties("Row")="1~1~1~2~2~2~3~3~3"
'										dicEditProperties("Column")="1~2~3~1~2~3~1~2~3"
'										dicEditProperties("Value")="200~300~400~500~600~700~800~900~1000"/dicEditProperties("Value")="10.0:M_Ten~20.0:M_Twenty~:M_Thirty~50.0:Fifty"
'										bReturn= Fn_Mech_EditProperties("InitialValues","",dicEditProperties,"Save")
'										
'										dicEditProperties("Value")="ROW1~ROW2~ROW3"
'										bReturn= Fn_Mech_EditProperties("RemoveRowLabels","",dicEditProperties,"")
'										
'										dicEditProperties("Value")="M_ROW`~M_ROW2~M_ROW3"
'										bReturn= Fn_Mech_EditProperties("AddRowLabels","",dicEditProperties,"")
'										
'										dicEditProperties("PropertyName")="Type"
'										dicEditProperties("Value")="Rational"
'										bReturn= Fn_Mech_EditProperties("DropDownList","",dicEditProperties,"")
'										
'										dicEditProperties("PropertyName")="Comment~Rows"
'										dicEditProperties("Value")="new comment~4"
'										bReturn= Fn_Mech_EditProperties("EditBox","",dicEditProperties,"")
'										
'										dicEditProperties("ConstantName")="A~B"
'										dicEditProperties("ConstantValue")="10~20"
'										bReturn= Fn_Mech_EditProperties("ConstantsTable","",dicEditProperties,"")
'										
'										dicEditProperties("PropertyName")="Is Signed"
'										dicEditProperties("Value")="True"
'										bReturn= Fn_Mech_EditProperties("RadioButton","",dicEditProperties,"")
'
'										dicEditProperties("DomainElementName")="Ele2~Ele1"
'										dicEditProperties("Value")="M_Ele2:38:M_Domain 1~Ele1:24:M_Dom2"
'										bReturn=Fn_Mech_EditProperties("ValidValues","",dicEditProperties,"")
'										
'										dicEditProperties("DomainElementName")="M_Ele2"
'										bReturn=Fn_Mech_EditProperties("SEDInitialValue","",dicEditProperties,"")
'
'										dicEditProperties("Value")="1:7:M_B1N7:1:2~2:3:M_B3N8:3:4~3:4:M_B4N8:5:6"
'										dicEditProperties("Value")="Byte number:Bit number:new Bit name:new Bit 0 meaning:new Bit 1 meaning~Byte number:Bit number:new Bit name:new Bit 0 meaning:new Bit 1 meaning~Byte number:Bit number:new Bit name:new Bit 0 meaning:new Bit 1 meaning"
'										bReturn=Fn_Mech_EditProperties("BitDefination","",dicEditProperties,"")
'
'										dicEditProperties("Row")="1~1~2~2"
'										dicEditProperties("Column")="1~2~1~2"
'										dicEditProperties("Value")="8-5-2012:NewDate1~5-6-2012:NewDate2~10-3-2012:NewDate3~31-12-2012:NewDate4"
'										bReturn=Fn_Mech_EditProperties("DateInitialValues","",dicEditProperties,"")
'
'										dicEditProperties("Row")="1~1~1~2~2~2"
'										dicEditProperties("Column")="1~2~3~1~2~3"
'										dicEditProperties("Value")="false:First~true:Second~false~true~true~true:Last"
'										bReturn=Fn_Mech_EditProperties("BoolInitialValues","",dicEditProperties,"")
'
'										dicEditProperties("ParameterName")="FETDisable_Cfg5459"
'										bReturn=Fn_Mech_EditProperties("ParameterValuesCellDoubleClick","",dicEditProperties,"")
'
'									dicEditProperties("Value")="Byte Number:Bit Number:Column Name~Byte Number:Bit Number:Column Name"
'									dicEditProperties("Value")="1:7:Bit Number~1:6:Bit Number"
'									bReturn=Fn_Mech_EditProperties("VerifyBitDefinationCellEditable","",dicEditProperties,"")
'
'									dicEditProperties("Row")="1"
'									dicEditProperties("Column")="1"
'									dicEditProperties("ShowValueDescription")="off"
'									bReturn=Fn_Mech_EditProperties("MaximumValues_PasteData","",dicEditProperties,"")
'									
'									dicEditProperties("Row")="2"
'									dicEditProperties("ShowValueDescription")="off"
'									bReturn=Fn_Mech_EditProperties("InitialValues_PasteData","",dicEditProperties,"")
'									
'									dicEditProperties("ShowValueDescription")="off"
'									bReturn=Fn_Mech_EditProperties("MinimumValues_PasteData","",dicEditProperties,"")
'
'									dicEditProperties("Byte")="1"
'									bReturn=Fn_Mech_EditProperties("BitDefination_PasteData","",dicEditProperties,"")
'									
'									dicEditProperties("Byte")="2~2"
'									dicEditProperties("BitNumber")="0~1"
'									bReturn=Fn_Mech_EditProperties("BitDefination_PasteData","",dicEditProperties,"")
'									
'									bReturn=Fn_Mech_EditProperties("BitDefination_PasteData","",dicEditProperties,"")
'
'									dicEditProperties("Row")="1~1~1~2~2~2~3~3~3"
'									dicEditProperties("Column")="1~2~3~1~2~3~1~2~3"
'									dicEditProperties("Value")="600~700~800~900~1000~1100~1200~1300~1400"
'									dicEditProperties("Collapse")="off"
'									bReturn=Fn_Mech_EditProperties("MaximumValues","",dicEditProperties,"")
'
'									dicEditProperties("Row")="1~1"
'									dicEditProperties("Column")="1~4"
'									bReturn=Fn_Mech_EditProperties("MaximumValues_PasteRowsCols","",dicEditProperties,"")'
'
'									bReturn=Fn_Mech_EditProperties("ValidValues_PasteData","",dicEditProperties,"")
'									
'									dicEditProperties("Row")="2~3~4"
'									bReturn=Fn_Mech_EditProperties("ValidValues_PasteData","",dicEditProperties,"")
'
'												dicEditProperties("Column")="1"
'												dicEditProperties("Color")="Red"
'												bReturn= Fn_Mech_EditProperties("InitialValues_VerifyHeaderForegroundColour","",dicEditProperties,"")
'												
'												dicEditProperties("Row")="1"
'												dicEditProperties("Column")="1"
'												dicEditProperties("CellErrorMessage")="Property = Initial Values[1][1], Reason = Value is not within min-max limit [Maximum: 100, Minimum: 40]"
'												bReturn= Fn_Mech_EditProperties("InitialValues_VerifyCellErrorMessage","",dicEditProperties,"")
'
'												dicEditProperties("PropertyName")="Checked-Out"
'														dicEditProperties("PropertyState")="value"
'														bReturn=Fn_Mech_EditProperties("EditBox_GetPropertyState","",dicEditProperties,"")
'
'												dicEditProperties("Row")="3"
'												dicEditProperties("Column")="3"
'												dicEditProperties("Collapse")="Off"
'												bReturn=Fn_Mech_EditProperties("InitialValues_SelectCell","",dicEditProperties,"")
'												
'												dicEditProperties("Row")="3"
'												dicEditProperties("Column")="3"
'												dicEditProperties("Collapse")="Off"
'												bReturn=Fn_Mech_EditProperties("InitialValues_DoubleClickCell","",dicEditProperties,"")
'
'												dicEditProperties("Row")="3"
'												bReturn=Fn_Mech_EditProperties("SEDValidValues_IsNameAndValueCorrect","",dicEditProperties,"")
'
'												dicEditProperties("PropertyName")="Action"
'												bReturn=Fn_Mech_EditProperties("PasteLink","",dicEditProperties,"")
'												
'												dicEditProperties("PropertyName")="Action"
'												bReturn=Fn_Mech_EditProperties("CopyLink","",dicEditProperties,"")
'												
'												dicEditProperties("PropertyName")="Action"
'												dicEditProperties("Value")="Copy~Paste~Clear"
'												bReturn=Fn_Mech_EditProperties("VerifyLinkMenu","",dicEditProperties,"")
'
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												26-Apr-2012								1.0																						Sunny R
'													Sandeep N												27-Apr-2012								1.1					Added Case : RadioButton				Sunny R
'													Sandeep N												02-May-2012								1.2					Added Case : ValidValues,SEDInitialValue				Sunny R
'													Sandeep N												03-May-2012								1.3					Added Case : BitDefination				Sunny R
'													Sandeep N												04-May-2012								1.4					Added Case : DateInitialValues		 Sunny R
'													Sandeep N												28-May-2012								1.5					Added Case : BoolInitialValues		 Sunny R
'													Sandeep N												31-May-2012								1.6					Added Case : VerifyBitDefinationCellEditable		 Sunny R
'													Sandeep N												01-Jun-2012								1.7					Added Case : ParameterValuesCellDoubleClick		 Sunny R
'													Sandeep N												08-Jun-2012								1.8					Added Case : "InitialValues_PasteData","MaximumValues_PasteData","MinimumValues_PasteData"		 Sunny R
'													Sandeep N												13-Jun-2012								1.9					Added Case : BitDefination_PasteData		 Sunny R
'													Sandeep N												02-Jul-2012								1.10					Added Case : MaximumValues_PasteRowsCols		 Sunny R
'													Sandeep N												02-Aug-2012								1.11					Added Case : ValidValues_PasteData		 Anjali M
'													Sandeep N												06-Aug-2012								1.12					Added Case : InitialValues_VerifyHeaderForegroundColour,CellErrorMessage		 Anjali M
'													Sandeep N												28-Aug-2012								1.13					Added Case : InitialValues_SelectCell,InitialValues_DoubleClickCell		Sachin J
'													Sandeep N												28-Aug-2012								1.14					Added Case : SEDValidValues_IsNameAndValueCorrect       		Anjali M
' 													Pranav Ingle											07-Nov-2012								1.15					Added Case :  DropDownMenu
' 													Sandeep N												04-Dec-2012								1.16					Added Case :  VerifyLinkMenu,CopyLink,PasteLink
' 													Pranav Ingle											19-Nov-2013								1.17					Modified Case :    ValidValues
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_Mech_EditProperties(StrAction,StrLink,dicEditProperties,StrButtonName)
	GBL_FAILED_FUNCTION_NAME="Fn_Mech_EditProperties"
 	'Declaring variables
	Dim objPropetiesDialog,objTable,objChld,objCheckOut,objCheckIn
	Dim aRowNumber,aColumnNumber,aValues,iCounter,bFlag,iCount,aProperty,iRows,iDragCounter
	Dim aEleName,aConstantName,max,aValDesc,aDate,objDate,iRow
	Dim aByte,aBitNumber,aAction,tableName,crrErrMsg,sColour,aCol,sColourCode,sFirstChar
	Dim WshShell
	Dim iHieght,iTempHieght,iWidth,iX
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Handling New Check Out dialog for [ My Teamcenter ] and [ Change Manager ] perspective
	Dim StrTitle
	StrTitle=JavaWindow("DefaultWindow").GetROProperty("title")
	If Instr(1,StrTitle,"My Teamcenter") or Instr(1,StrTitle,"Change Manager") Then
		Set objCheckOut= JavaWindow("DefaultWindow").JavaWindow("Check-Out")
	Else
		Set objCheckOut=JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Check-Out")
	End if
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

	Fn_Mech_EditProperties=False
 	'Checking existance of [ Properties ] dialog
	If not JavaWindow("Mechatronics").JavaWindow("TcDefaultApplet").JavaDialog("EditProperties").Exist(3) Then
		'checking existance of Check Out dialog
		If not  objCheckOut.Exist(3) Then
			'calling menu [ Edit:Properties ]
			Call Fn_MenuOperation("Select","Edit:Properties")
		else
			'ckicking on Yes button to checkout object
			Call Fn_Button_Click("Fn_Mech_EditProperties",JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Check-Out"),"Yes")
		End If
	End If
	Call Fn_ReadyStatusSync(1)
	'Creating object of [ EditProperties ] dialog
	Set objPropetiesDialog=JavaWindow("Mechatronics").JavaWindow("TcDefaultApplet").JavaDialog("EditProperties")
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Clicking on page link
	'    Added By Pranav Ingle  -  14-Aug-2012  To To Skip call to 'General' and 'All' Links in Edit Properties Dialog after Invalid Scenarios
	If StrAction<>"MaximumValues_VerifyCellErrorMessage" And StrAction<>"MinimumValues_VerifyCellErrorMessage" And StrAction<>"InitialValues_VerifyCellErrorMessage" And StrAction<>"InitialValues_VerifyHeaderForegroundColour" And StrAction<>"MaximumValues_VerifyHeaderForegroundColour" And StrAction <> "MinimumValues_VerifyHeaderForegroundColour" Then
		If StrLink="" Then
			objPropetiesDialog.JavaStaticText("PageLink").SetTOProperty "label","General"
			objPropetiesDialog.JavaStaticText("PageLink").Click 1,1,"LEFT"
		Else
			objPropetiesDialog.JavaStaticText("PageLink").SetTOProperty "label",StrLink
			objPropetiesDialog.JavaStaticText("PageLink").Click 1,1,"LEFT"
			If StrLink="All" Then
				'to click on [ Show empty properties... ] link
				wait 1
				objPropetiesDialog.JavaStaticText("PageLink").SetTOProperty "label","Show empty properties..."
				If objPropetiesDialog.JavaStaticText("PageLink").Exist(2) Then
					objPropetiesDialog.JavaStaticText("PageLink").Click 1,1,"LEFT"
				End If
			End If
		End If
	End If		
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Select Case StrAction
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to modify values of Maximum ,minimum , Initial Value Table
		Case "InitialValues","MaximumValues","MinimumValues"
				
				If objPropetiesDialog.JavaSlider("JScrollPane").Exist(2) Then
					max=objPropetiesDialog.JavaSlider("JScrollPane").GetROProperty("max")
					objPropetiesDialog.JavaSlider("JScrollPane").Drag max
				End If
				wait 1
				If dicEditProperties("Collapse")<>"" Then
					Call Fn_CheckBox_Set("Fn_Mech_EditProperties", objPropetiesDialog,StrAction&"Collapse",dicEditProperties("Collapse"))
					wait 1
				End If
				If objPropetiesDialog.JavaCheckBox(StrAction&"ShowValueDescription").CheckProperty ("value",0) Then
					Call Fn_CheckBox_Set("Fn_Mech_EditProperties", objPropetiesDialog,StrAction&"ShowValueDescription","ON")
					wait 1
				End If
				'Spliting row numbers
				aRowNumber=Split(dicEditProperties("Row"),"~")
				aColumnNumber=Split(dicEditProperties("Column"),"~")
				aValues=Split(dicEditProperties("Value"),"~")
				objPropetiesDialog.JavaTable(StrAction).SelectRow 0
				For iCounter=0 to ubound(aRowNumber)
                    objPropetiesDialog.JavaTable(StrAction).DeselectRow Cint(aRowNumber(iCounter))-1
                    objPropetiesDialog.JavaTable(StrAction).ActivateCell Cint(aRowNumber(iCounter))-1,Cint(aColumnNumber(iCounter))-1
					wait 1
					aValDesc=Split(aValues(iCounter),":")

					Set objTable=Description.Create
					objTable("Class Name").value="JavaTable"
					Set objChld=objPropetiesDialog.JavaTable(StrAction).ChildObjects(objTable)
					If aValDesc(0)<>"" Then
						objChld(0).SetCellData 0,0,aValDesc(0)
					End If
					If ubound(aValDesc)=1 Then
						objChld(0).SetCellData 1,0,aValDesc(1)
					End If
'					objChld(0).Object.setFocusable False
					Set objChld=nothing
					Set objTable=nothing
				Next
				If Err.Number < 0 Then
					Fn_Mech_EditProperties=false
				else
					Fn_Mech_EditProperties=true
				End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to modify values of Row Labels and Column Lables list Box
		Case "AddRowLabels","AddColumnLabels"
	
				If objPropetiesDialog.JavaSlider("JScrollPane").Exist(2) Then
					max=objPropetiesDialog.JavaSlider("JScrollPane").GetROProperty("max")
					objPropetiesDialog.JavaSlider("JScrollPane").Drag max
				End If
				wait 1
				If StrAction="AddRowLabels" Then
					Call Fn_CheckBox_Set("Fn_Mech_EditProperties", objPropetiesDialog, "RowLabels","on")
				Else
					Call Fn_CheckBox_Set("Fn_Mech_EditProperties", objPropetiesDialog, "ColumnLabels","on")
				End If
				aValues=Split(dicEditProperties("Value"),"~")
				For iCounter=0 to uBound(aValues)
					Call Fn_Edit_Box("Fn_Mech_EditProperties",objPropetiesDialog,"Labels",aValues(iCounter))
					Call Fn_Button_Click("Fn_Mech_EditProperties", objPropetiesDialog, "add_16")
				Next
				If StrAction="AddRowLabels" Then
					Call Fn_CheckBox_Set("Fn_Mech_EditProperties", objPropetiesDialog, "RowLabels","off")
				Else
					Call Fn_CheckBox_Set("Fn_Mech_EditProperties", objPropetiesDialog, "ColumnLabels","off")
				End If
				If Err.Number < 0 Then
					Fn_Mech_EditProperties=false
				else
					Fn_Mech_EditProperties=true
				End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to remove values of Row Labels and Column Lables list Box
		Case "RemoveRowLabels","RemoveColumnLabels"
				If objPropetiesDialog.JavaSlider("JScrollPane").Exist(2) Then
					max=objPropetiesDialog.JavaSlider("JScrollPane").GetROProperty("max")
					objPropetiesDialog.JavaSlider("JScrollPane").Drag max
				End If
				wait 1

				If StrAction="RemoveRowLabels" Then
					Call Fn_CheckBox_Set("Fn_Mech_EditProperties", objPropetiesDialog, "RowLabels","on")
					objPropetiesDialog.JavaStaticText("PropertyLabel").SetTOProperty "label","Row Lables:"
				Else
					Call Fn_CheckBox_Set("Fn_Mech_EditProperties", objPropetiesDialog, "ColumnLabels","on")
					objPropetiesDialog.JavaStaticText("PropertyLabel").SetTOProperty "label","Column Lables:"
				End If
				aValues=Split(dicEditProperties("Value"),"~")
				For iCounter=0 to uBound(aValues)
					Call Fn_List_Select("Fn_Mech_EditProperties", objPropetiesDialog,"ListBox",aValues(iCounter))
					Call Fn_Button_Click("Fn_Mech_EditProperties", objPropetiesDialog,"remove_16")
				Next
				If StrAction="RemoveRowLabels" Then
					Call Fn_CheckBox_Set("Fn_Mech_EditProperties", objPropetiesDialog, "RowLabels","off")
				Else
					Call Fn_CheckBox_Set("Fn_Mech_EditProperties", objPropetiesDialog, "ColumnLabels","off")
				End If
				If Err.Number < 0 Then
					Fn_Mech_EditProperties=false
				else
					Fn_Mech_EditProperties=true
				End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to modify values from list Box
		Case "DropDownList"
				objPropetiesDialog.JavaStaticText("PropertyLabel").SetTOProperty "label",dicEditProperties("PropertyName")+":"
				If objPropetiesDialog.JavaList("DropDownList").Exist(2) Then
					For iCounter=0 to 0
						Call Fn_List_Select("Fn_Mech_EditProperties", objPropetiesDialog,"DropDownList",dicEditProperties("Value"))			
					Next
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"List box [ " & dicEditProperties("PropertyName") & " ] is not exist on dialog")
				End If
				If Err.Number < 0 Then
					Fn_Mech_EditProperties=false
				else
					Fn_Mech_EditProperties=true
				End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to modify property value from edit boxes
		Case "EditBox"
				
				aProperty=Split(dicEditProperties("PropertyName"),"~")
				aValues=Split(dicEditProperties("Value"),"~")
				'Added code to handle hidden Edit box
				'21-03-2013
				If objPropetiesDialog.JavaSlider("JScrollPane").Exist(2) Then
					max=objPropetiesDialog.JavaSlider("JScrollPane").GetROProperty("max")
					objPropetiesDialog.JavaSlider("JScrollPane").Drag max
				End If
				For iCounter=0 to UBound(aProperty)

                        bFlag=False
						objPropetiesDialog.JavaStaticText("PropertyLabel").SetTOProperty "label",aProperty(iCounter)+":"
						If objPropetiesDialog.JavaButton("ParameterDef_DropDown").Exist(2) Then

								If instr(1,aValues(iCounter),"\") Then
									aValues(iCounter)=Replace(aValues(iCounter),"\","")
								End If

								If objPropetiesDialog.JavaSlider("JScrollPane").Exist(2) Then
									max=objPropetiesDialog.JavaSlider("JScrollPane").GetROProperty("max")
									objPropetiesDialog.JavaSlider("JScrollPane").Drag max
								End If

								wait 1
								objPropetiesDialog.JavaButton("ParameterDef_DropDown").Click
								wait 4
								Set objTable=Description.Create()
								objTable("Class Name").value="JavaTable"
								objTable("toolkit class").value="com.teamcenter.rac.common.lov.view.components.LOVTreeTable"
								objTable("displayed").value="1"
								objTable("enabled").value="1"
								objTable("focused").value="1"
								
								Set objChild=objPropetiesDialog.ChildObjects(objTable)

								sFirstChar=Mid(aValues(iCounter),1,1)
								bFlag=False
								For iDragCounter = 0 To 50
									Set WshShell = CreateObject("WScript.Shell")
									WshShell.SendKeys "^{END}"
'									wait 1
									If aValues(iCounter) < Cint(Asc(Mid(trim(objChild(0).Object.getValueAt(objChild(0).GetROProperty("rows")-1,0).getDisplayableValue()),1,1))) Then
										wait 1
										WshShell.SendKeys "{TAB}"
										WshShell.SendKeys "^{END}"
										wait 1
										bFlag=True
										Set WshShell =Nothing
										Exit For	
									End If
									Set WshShell =Nothing
								Next
'								If bFlag=False Then
'									Exit Function
'								Else
'									wait 1
'								End If

								For iCount=0 to objChild(0).GetROProperty("rows")-1
									If trim(aValues(iCounter))=trim(objChild(0).Object.getValueAt(iCount,0).getDisplayableValue()) Then
										'objChild(0).DoubleClickCell iCount,0
										objChild(0).SelectRow iCount
										wait 2
										Fn_Mech_EditProperties=true
										Exit for
									Else									
										Fn_Mech_EditProperties=false
									End If
									
								Next

								Set objTable=Nothing
								Set objChild=Nothing
								Set WshShell =Nothing
'								
'
'								If  aProperty(iCounter) = "Control Engineer" OR aProperty(iCounter) = "Size Units" Then
'									objPropetiesDialog.JavaEdit("EditBox").Object.setText aValues(iCounter)
'								Else
'									Call Fn_Edit_Box("Fn_Mech_EditProperties",objPropetiesDialog,"EditBox",aValues(iCounter))
'									wait 2
'									Set WshShell = CreateObject("WScript.Shell")
'									WshShell.SendKeys "{ENTER}"
'									Set WshShell =nothing
'								End If		
'								wait 1
						Else
								If objPropetiesDialog.JavaEdit("EditBox").Exist(3) Then
								
									If aProperty(iCounter)="Rows" or aProperty(iCounter)="Columns" or aProperty(iCounter)="Size in Byte(s)" Then
'											Call Fn_Edit_Box("Fn_Mech_EditProperties",objPropetiesDialog,"EditBox",aValues(iCounter)+ vbLf + "")
                                            objPropetiesDialog.JavaEdit("EditBox").Set aValues(iCounter)+ vbLf + ""
											wait 1
									ElseIf aProperty(iCounter)="Resolution Numerator" Then
											'Call Fn_Edit_Box("Fn_Mech_EditProperties",objPropetiesDialog,"EditBox",aValues(iCounter)+ vbTab + "")
											objPropetiesDialog.JavaEdit("EditBox").Set aValues(iCounter)+ vbTab + ""
											wait 1
									Else
											Call Fn_Edit_Box("Fn_Mech_EditProperties",objPropetiesDialog,"EditBox",aValues(iCounter))
											
									End If
											
								else
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Property [ " & aProperty(iCounter) & " ] is not exist on dialog")
											Exit function
								End If
								
						End If
				Next
				If Err.Number < 0 Then
					Fn_Mech_EditProperties=false
				else
					Fn_Mech_EditProperties=true
				End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to verify property value from edit boxes
		Case "DropDownMenu"  									'-   Added To Modify Header ObjectAnd Trailer Object Propeties  - Pranav Ingle
				aProperty=Split(dicEditProperties("PropertyName"),"~")
				aValues=Split(dicEditProperties("Value"),"~")

				bFlag=true
				For iCounter=0 to UBound(aProperty)

					objPropetiesDialog.JavaStaticText("PropertyLabel").SetTOProperty "label",aProperty(iCounter)+":"
					'objPropetiesDialog.JavaObject("HeaderObject").Click 1,1
					objPropetiesDialog.JavaStaticText("DropDownButton").Click 1,1
					wait 1
					
					err.Clear
					objPropetiesDialog.JavaMenu("label:=" & aValues(iCounter)).Click 1,1

					If Err.Number < 0 Then
						bFlag=false
					End If

				Next
				If bFlag=True Then
					Fn_Mech_EditProperties=true
				End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to verify conversion rule name,Description
		Case "ConversionRuleName","ConversionRuleDescription"
				Call Fn_Edit_Box("Fn_Mech_EditProperties",objPropetiesDialog,StrAction,aValues(iCounter))
				If Err.Number < 0 Then
					Fn_Mech_EditProperties=false
				else
					Fn_Mech_EditProperties=true
				End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to modify values from Constants Table
		Case "ConstantsTable"
				aConstantName=Split(dicEditProperties("ConstantName"),"~")
				aValues=Split(dicEditProperties("ConstantValue"),"~")
				iRows=objPropetiesDialog.JavaTable("Constants").GetROProperty("rows")
				For iCounter=0 to ubound(aConstantName)
					bFlag=False
					For iCount=0 to iRows-1
						If objPropetiesDialog.JavaTable("Constants").GetCellData(iCount,"Constant Name")=aConstantName(iCounter) Then
							If dicEditProperties("Column")="" Then
								dicEditProperties("Column")="Constant Value"
							End If
							objPropetiesDialog.JavaTable("Constants").ClickCell iCount,dicEditProperties("Column"),"LEFT"
							wait 1
							Call Fn_Edit_Box("Fn_Mech_EditProperties",objPropetiesDialog,"ConstantsValue",aValues(iCounter))
							bFlag=True
							Exit for
						End If
					Next
					If bFlag=false Then
						Exit for
					End If
				Next
				If bFlag=True Then
					Fn_Mech_EditProperties=true
				End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to modify Radio Buttons
		Case "RadioButton"
				aProperty=Split(dicEditProperties("PropertyName"),"~")
				aValues=Split(dicEditProperties("Value"),"~")
				For iCounter=0 to UBound(aProperty)
					objPropetiesDialog.JavaStaticText("PropertyLabel").SetTOProperty "label",aProperty(iCounter)+":"
					objPropetiesDialog.JavaRadioButton("RadioButton").SetTOProperty "attached text",aValues(iCounter)

					If objPropetiesDialog.JavaRadioButton("RadioButton").Exist(3) Then
						Call Fn_UI_JavaRadioButton_SetON("Fn_Mech_EditProperties",objPropetiesDialog, "RadioButton")
					else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Property [ " & aProperty(iCounter) & " ] is not exist on dialog")
						Exit for
					End If
				Next
				If Err.Number < 0 Then
					Fn_Mech_EditProperties=false
				else
					Fn_Mech_EditProperties=true
				End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to modify Initial Values of SED Initial Values table
		Case "SEDInitialValue"
			If dicEditProperties("DomainElementName")<>"" Then
				objPropetiesDialog.JavaTable("SEDInitialValue").SetCellData 0,"Domain Element Name",dicEditProperties("DomainElementName")
			End If
			If dicEditProperties("Value")<>"" Then
				objPropetiesDialog.JavaTable("SEDInitialValue").SetCellData 0,"Value",dicEditProperties("Value")
			End If
			if Err.Number < 0 Then
					Fn_Mech_EditProperties=false
			else
					Fn_Mech_EditProperties=true
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to modify Valid Values of SED Valid Values table :  
		'		Modified Case to handle  Inner Table  - Pranav 19-Nov-2013
		Case "ValidValues"
			bFlag=false
			aEleName=Split(dicEditProperties("DomainElementName"),"~")
			aValues=Split(dicEditProperties("Value"),"~")
			For iCounter=0 to ubound(aEleName)
				aValDesc=Split(aValues(iCounter),":")
				For iCount=0 to Cint(objPropetiesDialog.JavaTable("ValidValues").GetROProperty("rows"))-1
					If trim(objPropetiesDialog.JavaTable("ValidValues").Object.getCellData(iCount,0).get(2).toString())=aEleName(iCounter) Then
'						objPropetiesDialog.JavaTable("ValidValues").SelectRow iCount
						objPropetiesDialog.JavaTable("ValidValues").Click 1,1
						wait 2
						'setting new Domain Element Name
						If aValDesc(0)<>"" Then
							objPropetiesDialog.JavaTable("ValidValues").DoubleClickCell iCount,0
							objPropetiesDialog.JavaTable("ValidValuesInnerTable").SetCellData 0,0,aValDesc(0)
							wait 1
						End If
						'setting new Value
						If aValDesc(1)<>"" Then
'							objPropetiesDialog.JavaTable("ValidValues").DoubleClickCell iCount,0
'							wait 1
'							objPropetiesDialog.JavaTable("ValidValuesInnerTable").SetCellData 0,1,aValDesc(1)
'							wait 1

							iHieght= objPropetiesDialog.JavaTable("ValidValues").Object.getRowHeight(0)
							iTempHieght=iHieght/2
							iHieght=iTempHieght
							objPropetiesDialog.JavaTable("ValidValues").SelectRow iCount
							If iCounter<>0 Then
								iHieght=objPropetiesDialog.JavaTable("ValidValues").Object.getRowHeight(0)*iCount+iTempHieght
							End If

							iWidth= objPropetiesDialog.JavaObject("ValidValuesTableHeader").Object.getTable().getTableHeader().getHeaderRect(1).getWidth()/2
							iX= objPropetiesDialog.JavaObject("ValidValuesTableHeader").Object.getTable().getTableHeader().getHeaderRect(1).getX()
							iWidth=iWidth+iX
							objPropetiesDialog.JavaTable("ValidValues").Click iWidth,iHieght,"LEFT"
							wait 1
							objPropetiesDialog.JavaTable("ValidValues").DblClick iWidth,iHieght,"LEFT"
							If objPropetiesDialog.JavaEdit("ValidValuesTableEdit").Exist(3) Then
								Call Fn_Edit_Box("Fn_Mech_ValidValuesTable",objPropetiesDialog,"ValidValuesTableEdit", aValDesc(1))
								objPropetiesDialog.JavaEdit("ValidValuesTableEdit").Activate
							Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail to set Domain value")
								Exit function
							End If
							wait 1
						End If
						'setting new Description
						If aValDesc(2)<>"" Then
	'							objPropetiesDialog.JavaTable("ValidValues").DoubleClickCell iCount,0
	'							objPropetiesDialog.JavaTable("ValidValuesInnerTable").SetCellData 0,2,aValDesc(2)
	'							wait 1
								iWidth= objPropetiesDialog.JavaObject("ValidValuesTableHeader").Object.getTable().getTableHeader().getHeaderRect(2).getWidth()/2
								iX= objPropetiesDialog.JavaObject("ValidValuesTableHeader").Object.getTable().getTableHeader().getHeaderRect(2).getX()
								iWidth=iWidth+iX
								objPropetiesDialog.JavaTable("ValidValues").Click iWidth,iHieght,"LEFT"
								wait 1
								objPropetiesDialog.JavaTable("ValidValues").DblClick iWidth,iHieght,"LEFT"
								If objPropetiesDialog.JavaEdit("ValidValuesTableDescEdit").Exist(3)  Then
									Call Fn_Edit_Box("Fn_Mech_ValidValuesTable",objPropetiesDialog,"ValidValuesTableDescEdit",  aValDesc(2)+vbLf)
									wait 1
                                Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail to set Domain Description")
									Exit function
								End If
						End If
						bFlag=true
						Exit for
					End If
				Next
				If bFlag=false Then
					Exit for
				End If
			Next
			If bFlag=true Then
				Fn_Mech_EditProperties=true
			else
				Fn_Mech_EditProperties=false
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to modify Bit Defination
		Case "BitDefination"
			If objPropetiesDialog.JavaCheckBox("BitDefinitionTableCollapse").Exist(2) Then
				Call Fn_CheckBox_Set("Fn_Mech_BitDefination", objPropetiesDialog, "BitDefinitionTableCollapse", "on")
			End If
			aValues=Split(dicEditProperties("Value"),"~")
			For iCounter=0 to ubound(aValues)
				aValDesc=Split(aValues(iCounter),":")
				iRow=Cint(aValDesc(0))*8-CInt(aValDesc(1))-1

				'Modifing Bit name
                If aValDesc(2)<>"" Then
					objPropetiesDialog.JavaTable("CCDMBitDefTable").SetCellData iRow,"Name",aValDesc(2)
				End If
				'Modifing 0 Meaning
                If aValDesc(3)<>"" Then
					objPropetiesDialog.JavaTable("CCDMBitDefTable").SetCellData iRow,"""0"" Meaning",aValDesc(3)
				End If
				'Modifing 1 Meaning
                If aValDesc(4)<>"" Then
					objPropetiesDialog.JavaTable("CCDMBitDefTable").SetCellData iRow,"""1"" Meaning",aValDesc(4)
				End If
			Next
			if Err.Number < 0 Then
					Fn_Mech_EditProperties=false
			else
					Fn_Mech_EditProperties=true
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to modify values of Date Maximum ,minimum , Initial Value Table
		Case "DateInitialValues","DateMaximumValues","DateMinimumValues"
				Select Case StrAction
					Case "DateInitialValues"
							StrAction="InitialValues"
					Case "DateMaximumValues"
							StrAction="MaximumValues"
					Case "DateMinimumValues"
							StrAction="MinimumValues"
				End Select
				If objPropetiesDialog.JavaSlider("JScrollPane").Exist(2) Then
					max=objPropetiesDialog.JavaSlider("JScrollPane").GetROProperty("max")
					objPropetiesDialog.JavaSlider("JScrollPane").Drag max
				End If
				wait 1
				If objPropetiesDialog.JavaCheckBox(StrAction&"ShowValueDescription").CheckProperty ("value",0) Then
					Call Fn_CheckBox_Set("Fn_Mech_EditProperties", objPropetiesDialog,StrAction&"ShowValueDescription","ON")
					wait 1
				End If
				'Spliting row numbers
				aRowNumber=Split(dicEditProperties("Row"),"~")
				aColumnNumber=Split(dicEditProperties("Column"),"~")
				aValues=Split(dicEditProperties("Value"),"~")
				For iCounter=0 to ubound(aRowNumber)
                    objPropetiesDialog.JavaTable(StrAction).DoubleClickCell Cint(aRowNumber(iCounter))-1,Cint(aColumnNumber(iCounter))-1
					wait 1
					aValDesc=Split(aValues(iCounter),":")
					If aValDesc(0)<>"" Then
						aDate=Split(aValDesc(0),"-")
						Set objDate=JavaDialog("SelectDate").JavaObject("CalendarPanel").Object.getTime()
	
						objDate.setYear(Cint(aDate(2))-1900)
						objDate.setMonth(Cint(aDate(1))-1)
						objDate.setDate(aDate(0))
						JavaDialog("SelectDate").JavaObject("CalendarPanel").Object.setTime(objDate)
						JavaDialog("SelectDate").JavaButton("Ok").Click
					End If
					If ubound(aValDesc)=1 Then
						objPropetiesDialog.JavaTable(StrAction).Object.getValueAt(Cint(aRowNumber(iCounter))-1,Cint(aColumnNumber(iCounter))-1).setDesc(aValDesc(1))
'						Set objTable=Description.Create
'						objTable("Class Name").value="JavaTable"
'						Set objChld=objPropetiesDialog.JavaTable(StrAction).ChildObjects(objTable)	
'						objChld(0).SetCellData 1,0,aValDesc(1)
					End If

					Set objChld=nothing
					Set objTable=nothing
				Next
				If Err.Number < 0 Then
					Fn_Mech_EditProperties=false
				else
					Fn_Mech_EditProperties=true
				End If
		'Case to modify values of nitial Value Tables Boolean value
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "BoolInitialValues"
				If objPropetiesDialog.JavaCheckBox("InitialValuesCollapse").Exist(2) Then
					Call Fn_CheckBox_Set("Fn_Mech_InitialValueTable", objPropetiesDialog,"InitialValuesCollapse","off")
				End If
				If objPropetiesDialog.JavaCheckBox("InitialValuesShowValueDescription").Exist(2) Then
					Call Fn_CheckBox_Set("Fn_Mech_InitialValueTable", objPropetiesDialog,"InitialValuesShowValueDescription","on")
				End If

				If objPropetiesDialog.JavaSlider("JScrollPane").Exist(2) Then
					max=objPropetiesDialog.JavaSlider("JScrollPane").GetROProperty("max")
					objPropetiesDialog.JavaSlider("JScrollPane").Drag max
				End If
				wait 1
				'Spliting row numbers
				aRowNumber=Split(dicEditProperties("Row"),"~")
				aColumnNumber=Split(dicEditProperties("Column"),"~")
				aValues=Split(dicEditProperties("Value"),"~")
				objPropetiesDialog.JavaTable("InitialValues").SelectRow 0
				For iCounter=0 to ubound(aRowNumber)
					objPropetiesDialog.JavaTable("InitialValues").DoubleClickCell Cint(aRowNumber(iCounter))-1,Cint(aColumnNumber(iCounter))-1
					wait 1
					aValDesc=Split(aValues(iCounter),":")
					Set objTable=Description.Create
					objTable("Class Name").value="JavaTable"
					Set objChld=objPropetiesDialog.JavaTable("InitialValues").ChildObjects(objTable)
					If aValDesc(0)<>"" Then
						If trim(objChld(0).getCellData(0,0))<>aValDesc(0) Then
							objChld(0).SelectRow 0
							objChld(0).ClickCell 0,0
						End If
					End If
					If uBound(aValDesc)=1 Then
						objChld(0).SelectRow 1
						objChld(0).SetCellData 1,0,aValDesc(1)
					End If

					'objChld(0).Object.setFocusable False
					Set objChld=nothing
					Set objTable=nothing
					objPropetiesDialog.JavaTable("InitialValues").DeselectRow Cint(aRowNumber(iCounter))-1
				Next
				If Err.Number < 0 Then
					Fn_Mech_EditProperties=false
				else
					Fn_Mech_EditProperties=true
				End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to Verify Bit Defination Cell editable or not
		Case "VerifyBitDefinationCellEditable"
			If objPropetiesDialog.JavaCheckBox("BitDefinitionTableCollapse").Exist(2) Then
				Call Fn_CheckBox_Set("Fn_Mech_BitDefination", objPropetiesDialog, "BitDefinitionTableCollapse", "on")
			End If
			aValues=Split(dicEditProperties("Value"),"~")
			For iCounter=0 to ubound(aValues)
				aValDesc=Split(aValues(iCounter),":")
				iRow=Cint(aValDesc(0))*8-CInt(aValDesc(1))-1
				Select Case aValDesc(2)
					Case "Bit Number"
						aColumnNumber=0
					Case "Name"
						aColumnNumber=1
				End Select
				If iCounter=0 Then
					aProperty=objPropetiesDialog.JavaTable("CCDMBitDefTable").Object.isCellEditable(iRow,aColumnNumber)
				else
					aProperty=aProperty+"~"+objPropetiesDialog.JavaTable("CCDMBitDefTable").Object.isCellEditable(iRow,aColumnNumber)
				End If
			Next
			if Err.Number < 0 Then
					Fn_Mech_EditProperties=false
			else
					Fn_Mech_EditProperties=aProperty
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to Double click on Parameter Values Cell Double Click
		Case "ParameterValuesCellDoubleClick"
			If objPropetiesDialog.JavaSlider("JScrollPane").Exist(2) Then
				max=objPropetiesDialog.JavaSlider("JScrollPane").GetROProperty("max")
				objPropetiesDialog.JavaSlider("JScrollPane").Drag max
				wait 1
			End If
			If dicEditProperties("Column")="" Then
				dicEditProperties("Column")="Parameter Values"
			End If
'			iRows=objPropetiesDialog.JavaTable("ParameterValues").GetROProperty("rows")
			iRows=Fn_UI_Object_GetROProperty("",JavaWindow("Mechatronics").JavaWindow("TcDefaultApplet").JavaDialog("EditProperties").JavaTable("ParameterValues"), "rows")
			bFlag=false
			For iCounter=0 to iRows-1
				If dicEditProperties("ParameterName")=objPropetiesDialog.JavaTable("ParameterValues").GetCellData(iCounter,"Name") Then
					objPropetiesDialog.JavaTable("ParameterValues").DoubleClickCell iCounter,dicEditProperties("Column")
					wait 2
					bFlag=true
					Exit for
				End If
			Next
			If bFlag=true Then
				Fn_Mech_EditProperties=true
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case toPaste Data in Maximum ,minimum , Initial Value Table
		Case "InitialValues_PasteData","MaximumValues_PasteData","MinimumValues_PasteData"
				If objPropetiesDialog.JavaSlider("JScrollPane").Exist(2) Then
					max=objPropetiesDialog.JavaSlider("JScrollPane").GetROProperty("max")
					objPropetiesDialog.JavaSlider("JScrollPane").Drag max
					wait 1
				End If
				StrAction=Split(StrAction,"_")
				If dicEditProperties("Collapse")<>"" Then
					Call Fn_CheckBox_Set("Fn_Mech_EditProperties", objPropetiesDialog,StrAction(0)&"Collapse",dicEditProperties("Collapse"))
					wait 1
				End If
				If dicEditProperties("ShowValueDescription")<>"" Then
					Call Fn_CheckBox_Set("Fn_Mech_EditProperties", objPropetiesDialog,StrAction(0)&"ShowValueDescription",dicEditProperties("ShowValueDescription"))
				End If
				If dicEditProperties("Row")<>"" and dicEditProperties("Column")<>"" Then
					objPropetiesDialog.JavaTable(StrAction(0)).ClickCell Cint(dicEditProperties("Row"))-1,Cint(dicEditProperties("Column"))-1,"RIGHT"
				elseif dicEditProperties("Row")<>"" then
					objPropetiesDialog.JavaTable(StrAction(0)&"Rows").SelectRow Cint(dicEditProperties("Row"))-1
					objPropetiesDialog.JavaTable(StrAction(0)&"Rows").ClickCell Cint(dicEditProperties("Row"))-1,0,"RIGHT"
				else
					objPropetiesDialog.JavaTable(StrAction(0)).ClickCell 0,0,"RIGHT"
				End If
				objPropetiesDialog.JavaMenu("index:=0","label:=Paste").Select
				If Err.Number < 0 Then
					Fn_Mech_EditProperties=False
				Else
					Fn_Mech_EditProperties=True
				End If
				If dicEditProperties("Row")<>"" and dicEditProperties("Column")="" Then
					objPropetiesDialog.JavaTable(StrAction(0)).DeselectRow Cint(dicEditProperties("Row"))-1
				End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to copy & paste data in [ Bit Defination ] table
		Case "BitDefination_CopyData","BitDefination_PasteData","BitDefination_CopyData_Keyboard","BitDefination_PasteData_Keyboard"
			If dicEditProperties("Byte")<>"" and dicEditProperties("BitNumber")<>"" Then
				'Spliting Byte information in array
				aByte=Split(dicEditProperties("Byte"),"~")
				aBitNumber=Split(dicEditProperties("BitNumber"),"~")
				iRow=Cint(aByte(0))*8-CInt(aBitNumber(0))-1

				objPropetiesDialog.JavaTable("CCDMBitDefTable").SelectRow iRow
				For iCounter=1 to ubound(aByte)
					iRow=Cint(aByte(iCounter))*8-CInt(aBitNumber(iCounter))-1
					objPropetiesDialog.JavaTable("CCDMBitDefTable").ExtendRow iRow
				Next
				objPropetiesDialog.JavaTable("CCDMBitDefTable").ClickCell iRow,"Bit Number","RIGHT"
			elseif dicEditProperties("Byte")<>"" then
				iRow=CInt(dicEditProperties("Byte"))-1
				objPropetiesDialog.JavaTable("CCDMBitDefRowHeaderTable").SelectRow iRow
				objPropetiesDialog.JavaTable("CCDMBitDefRowHeaderTable").ClickCell iRow,"Byte","RIGHT"
			else
				objPropetiesDialog.JavaTable("CCDMBitDefTable").SelectCell 0,"Bit Number"
                objPropetiesDialog.JavaTable("CCDMBitDefRowHeaderTable").ClickCell 0,0,"RIGHT"
			End If		
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 					
			Select Case StrAction
				' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				Case "BitDefination_CopyData"
					objPropetiesDialog.JavaMenu("index:=0","label:=Copy").Select
				' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				Case "BitDefination_PasteData"
					objPropetiesDialog.JavaMenu("index:=0","label:=Paste").Select
				' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			End Select
			If Err.Number < 0 Then
				Fn_Mech_EditProperties=False
			Else
				Fn_Mech_EditProperties=True
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case toPaste Data in Maximum ,minimum , Initial Value Table
		Case "InitialValues_PasteRowsCols","MaximumValues_PasteRowsCols","MinimumValues_PasteRowsCols"
				If objPropetiesDialog.JavaSlider("JScrollPane").Exist(2) Then
					max=objPropetiesDialog.JavaSlider("JScrollPane").GetROProperty("max")
					objPropetiesDialog.JavaSlider("JScrollPane").Drag max
					wait 1
				End If
				StrAction=Split(StrAction,"_")
				If dicEditProperties("Collapse")<>"" Then
					Call Fn_CheckBox_Set("Fn_Mech_EditProperties", objPropetiesDialog,StrAction(0)&"Collapse",dicEditProperties("Collapse"))
					wait 1
				End If
				If dicEditProperties("ShowValueDescription")<>"" Then
					Call Fn_CheckBox_Set("Fn_Mech_EditProperties", objPropetiesDialog,StrAction(0)&"ShowValueDescription",dicEditProperties("ShowValueDescription"))
				End If

				If dicEditProperties("Row")<>"" and dicEditProperties("Column")<>"" Then
					aRowNumber=Split(dicEditProperties("Row"),"~")
					aColumnNumber=Split(dicEditProperties("Column"),"~")
					'bjPropetiesDialog.JavaTable(StrAction(0)).SelectCellsRange cint(aRowNumber(0))-1,cint(aColumnNumber(0))-1,cint(aRowNumber(1))-1,Cint(aColumnNumber(1))-1
					objPropetiesDialog.JavaTable(StrAction(0)).SelectCellsRange cint(aRowNumber(0))-1,cint(aColumnNumber(0))-1,cint(aRowNumber(Ubound(aRowNumber)))-1,Cint(aColumnNumber(Ubound(aColumnNumber)))-1
					objPropetiesDialog.JavaTable(StrAction(0)).ClickCell Cint(aRowNumber(0))-1,Cint(aColumnNumber(0))-1,"RIGHT"
				elseif dicEditProperties("Row")<>"" then
					aRowNumber=Split(dicEditProperties("Row"),"~")
					objPropetiesDialog.JavaTable(StrAction(0)&"Rows").SelectRow Cint(aRowNumber(0))-1
					For iCounter=1 to ubound(aRowNumber)
						objPropetiesDialog.JavaTable(StrAction(0)&"Rows").ExtendRow CInt(aRowNumber(iCounter))-1
					Next
					objPropetiesDialog.JavaTable(StrAction(0)).ClickCell Cint(aRowNumber(UBound(aRowNumber)))-1,0,"RIGHT"
				End If
				objPropetiesDialog.JavaMenu("index:=0","label:=Paste").Select
				If Err.Number < 0 Then
					Fn_Mech_EditProperties=False
				Else
					Fn_Mech_EditProperties=True
				End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to Paste data in [ ISED Valid Values ] table
		Case "ValidValues_PasteData","ValidValues_PasteDataExt"
				If objPropetiesDialog.JavaSlider("JScrollPane").Exist(2) Then
					max=objPropetiesDialog.JavaSlider("JScrollPane").GetROProperty("max")
					objPropetiesDialog.JavaSlider("JScrollPane").Drag max
				End If
				
				wait 1
				' - - - - - Selecting Specified Cell or Row- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				If dicEditProperties("Row")<>"" Then
					iRows=Split(dicEditProperties("Row"),"~")
					objPropetiesDialog.JavaTable("ValidValues").SelectRow Cint(iRows(0))-1
					For iCounter=1 to ubound(iRows)
						objPropetiesDialog.JavaTable("ValidValues").ExtendRow Cint(iRows(iCounter))-1	
					Next
					objPropetiesDialog.JavaTable("ValidValues").ClickCell Cint(iRows(ubound(iRows)))-1,0,"RIGHT"
				else
					objPropetiesDialog.JavaTable("ValidValues").PressKey "A",micCtrl
					wait 1
					objPropetiesDialog.JavaTable("ValidValues").ClickCell 0,0,"RIGHT"
				End If
				'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				'Selecting menu [ Paste ]  to paste copied data
				objPropetiesDialog.JavaMenu("index:=0","label:=Paste").Select
				If Err.Number < 0 Then
					Fn_Mech_EditProperties=False
				Else
					Fn_Mech_EditProperties=True
				End If
				If StrAction<>"ValidValues_PasteDataExt" Then
					If dicEditProperties("Row")<>"" Then
						objPropetiesDialog.JavaTable("ValidValues").DeselectRow Cint(iRows(0))-1
					End If
				End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to verify Foreground color of Table header
		Case "InitialValues_VerifyHeaderForegroundColour","MaximumValues_VerifyHeaderForegroundColour","MinimumValues_VerifyHeaderForegroundColour"
			aAction=Split(StrAction,"_")
			tableName=aAction(0)
			aCol=Split(dicEditProperties("Column"),"~")
			For iCounter=0 to ubound(aCol)
				bFlag=false
				sColourCode=""
				sColour=objPropetiesDialog.JavaObject(tableName&"TableHeader").Object.getColumnModel().getColumn(cint(aCol(iCounter))-1).getHeaderRenderer().getForeground().toString()
				sColour =  mid(sColour ,instr(sColour ,"[")  ,instr(sColour ,"]") )
				Select Case lcase(dicEditProperties("Color"))
					Case "red"
						sColourCode="[r=255,g=0,b=0]"
				End Select
				If sColour=sColourCode Then
					bFlag=true
				else
					Exit for
				End if

			Next
			If bFlag=true Then
					Fn_Mech_EditProperties=true
			else
					Fn_Mech_EditProperties=false
			End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to verify cell error message in [ InitialValues , Maximum Values , MinimumValues ] table
		Case "InitialValues_VerifyCellErrorMessage","MaximumValues_VerifyCellErrorMessage","MinimumValues_VerifyCellErrorMessage"
				aAction=Split(StrAction,"_")
				tableName=aAction(0)
				'Setting [ ShowValueDescription ] option
				If dicEditProperties("ShowValueDescription")<>"" Then
					Call Fn_CheckBox_Set("Fn_Mech_EditProperties", objPropetiesDialog, tableName&"ValueDescription",dicEditProperties("ShowValueDescription"))
				End If
				'Setting [ Collapse ] option
				If dicEditProperties("Collapse")<>"" Then
					Call Fn_CheckBox_Set("Fn_Mech_EditProperties", objPropetiesDialog, tableName&"Collapse",dicEditProperties("Collapse"))
				End If
				'Getting current error message from Specific cell
				crrErrMsg=objPropetiesDialog.JavaTable(tableName).Object.getValueAt(Cint(dicEditProperties("Row"))-1,Cint(dicEditProperties("Column"))-1).getErrMsg()
				'Comparing user pass error message with actual error message
				If InStr(1,crrErrMsg,dicEditProperties("CellErrorMessage")) Then
					Fn_Mech_EditProperties=true
				else
					Fn_Mech_EditProperties=false
				End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to get specific property state of specific Edit box : e.g { current value, editable state, enabled state }
		Case "EditBox_GetPropertyState"
				objPropetiesDialog.JavaStaticText("PropertyLabel").SetTOProperty "label",dicEditProperties("PropertyName")+":"
				If objPropetiesDialog.JavaEdit("EditBox").Exist(3) Then
					Fn_Mech_EditProperties=objPropetiesDialog.JavaEdit("EditBox").GetROProperty(dicEditProperties("PropertyState"))
				else
					Fn_Mech_EditProperties=false
				End if
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to select / double click specific cell from [ InitialValues , MaximumValues  , MinimumValues ] tables
		Case "InitialValues_SelectCell","MaximumValues_SelectCell","MinimumValues_SelectCell","InitialValues_DoubleClickCell","MaximumValues_DoubleClickCell","MinimumValues_DoubleClickCell"
				If objPropetiesDialog.JavaSlider("JScrollPane").Exist(2) Then
					max=objPropetiesDialog.JavaSlider("JScrollPane").GetROProperty("max")
					objPropetiesDialog.JavaSlider("JScrollPane").Drag max
					wait 1
				End If
				StrAction=Split(StrAction,"_")
				If dicEditProperties("Collapse")<>"" Then
					Call Fn_CheckBox_Set("Fn_Mech_EditProperties", objPropetiesDialog,StrAction(0)&"Collapse",dicEditProperties("Collapse"))
					wait 1
				End If
				If dicEditProperties("ShowValueDescription")<>"" Then
					Call Fn_CheckBox_Set("Fn_Mech_EditProperties", objPropetiesDialog,StrAction(0)&"ShowValueDescription",dicEditProperties("ShowValueDescription"))
				End If
				Select Case StrAction(1)
					Case "SelectCell"
						objPropetiesDialog.JavaTable(StrAction(0)).ClickCell Cint(dicEditProperties("Row"))-1,Cint(dicEditProperties("Column"))-1
					Case "DoubleClickCell"
						objPropetiesDialog.JavaTable(StrAction(0)).DoubleClickCell Cint(dicEditProperties("Row"))-1,Cint(dicEditProperties("Column"))-1
				End Select
				wait 2
				If Err.Number < 0 Then
					Fn_Mech_EditProperties=False
				Else
					Fn_Mech_EditProperties=True
				End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
        'Case to check current value and name is correct or not
        Case "SEDValidValues_IsNameAndValueCorrect"
            If objPropetiesDialog.JavaTable("ValidValues").Object.getValueAt(Cint(dicEditProperties("Row"))-1,0).isNameCorrect()="true" and objPropetiesDialog.JavaTable("ValidValues").Object.getValueAt(Cint(dicEditProperties("Row"))-1,0).isValueCorrect()="true" then
                Fn_Mech_EditProperties=true
            else
                Fn_Mech_EditProperties=false
            end if
        '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "VerifyLinkMenu"
				If objPropetiesDialog.JavaSlider("JScrollPane").Exist(2) Then
					max=objPropetiesDialog.JavaSlider("JScrollPane").GetROProperty("max")
					objPropetiesDialog.JavaSlider("JScrollPane").Drag max
				End If
				wait 1
				objPropetiesDialog.JavaStaticText("PropertyLabel").SetTOProperty "label",dicEditProperties("PropertyName")+":"
				If objPropetiesDialog.JavaStaticText("DropDownButton").Exist(2) Then
					objPropetiesDialog.JavaStaticText("DropDownButton").Click 1, 1
				Else
					objPropetiesDialog.JavaObject("LinkOptionDropDown").Click 1,1	
				End If
				wait 1
                aValues=Split(dicEditProperties("Value"),"~")
				For iCounter=0 to UBound(aValues)
					bFlag=True
					If not objPropetiesDialog.JavaMenu("index:=0","label:=" & aValues(iCounter)).Exist(2) Then
						bFlag=False
						Exit for
					End If
				Next
				If bFlag=True Then
					Fn_Mech_EditProperties=True
				Else
					Fn_Mech_EditProperties=False
				End If
				If objPropetiesDialog.JavaStaticText("DropDownButton").Exist(2) Then
					objPropetiesDialog.JavaStaticText("DropDownButton").Click 1, 1
				Else
					objPropetiesDialog.JavaObject("LinkOptionDropDown").Click 1,1	
				End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "CopyLink","PasteLink","ClearLink"
				If objPropetiesDialog.JavaSlider("JScrollPane").Exist(2) Then
					max=objPropetiesDialog.JavaSlider("JScrollPane").GetROProperty("max")
					objPropetiesDialog.JavaSlider("JScrollPane").Drag max
				End If
				wait 1
				objPropetiesDialog.JavaStaticText("PropertyLabel").SetTOProperty "label",dicEditProperties("PropertyName")+":"
				'objPropetiesDialog.JavaObject("LinkOptionDropDown").Click 1,1
				objPropetiesDialog.JavaStaticText("DropDownButton").Click 10, 5,"LEFT"
				wait 1
				' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 					
				Select Case StrAction
					' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
					Case "CopyLink"
						objPropetiesDialog.JavaMenu("index:=0","label:=Copy").Select
					' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
					Case "PasteLink"
						objPropetiesDialog.JavaMenu("index:=0","label:=Paste").Select
					' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
					Case "ClearLink"
						objPropetiesDialog.JavaMenu("index:=0","label:=Clear").Select
					' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				End Select

				If Err.Number < 0 Then
					Fn_Mech_EditProperties=False
				Else
					Fn_Mech_EditProperties=True
				End If
	End Select
	'Clicking on button
	If StrButtonName<>"" Then
		Call Fn_Button_Click("Fn_Mech_EditProperties", objPropetiesDialog,StrButtonName)
		'saving changes and checking in
		If StrButtonName="SaveAndCheckIn" Then
			Set objCheckIn = Fn_SISW_GetObject("Check-In@2")
            If Not (objCheckIn.Exist(10)) Then
				Set objCheckIn = JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Check-In")
			    If  Not (objCheckIn.Exist(10)) Then
                   JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").SetTOProperty "Index", 0
                   Set objCheckIn = JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Check-In")         
                End If
			End if

			Call Fn_Button_Click("Fn_Mech_EditProperties",objCheckIn,"Yes")			
			Call Fn_ReadyStatusSync(1)
		End If
	JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").SetTOProperty "Index", 1
	End If
	'Releasing object of [ Properties ] dialog
	Set objPropetiesDialog=nothing
	Set objCheckIn = Nothing
	Set objCheckOut = Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_Mech_ValidValuesTable

'Description			 :	Function Used to perform operation on ValidValuesTable of parameter Defination [ ParmDefSED ] Revision Information

'Parameters			   :   1.StrAction: Action Name
'									 2.StrDomainElementName: Domain Element name
'									 3.StrValue: Values
'									 4.StrDescription : Desription
'
'Return Value		   : 	True or False

'Pre-requisite			:	Parameter Defination dialog should be appear on screen on user should be on Define Parameter Defination Revision Information dialog

'Examples				:   bReturn=Fn_Mech_ValidValuesTable("SetData","Ele1~Ele2","36~24","Domain 1~Domain 2")
'										bReturn=Fn_Mech_ValidValuesTable("VerifyCellData","D1","A2","Desc1")
'
'										bReturn=Fn_Mech_ValidValuesTable("GetAllContextMenu","","","")
'										For case [ GetAllContextMenu ] use StrDomainElementName parameter to pass row number if need to get available popup menu on specific row
'										bReturn=Fn_Mech_ValidValuesTable("GetAllContextMenu","2","","")
'										bReturn=Fn_Mech_ValidValuesTable("RowColumnCount","","","")
'										
'										For case [ PasteData ] use StrDomainElementName parameter to pass row number on which have to paste the data
'										bReturn=Fn_Mech_ValidValuesTable("PasteData","2","","")
'										
'										For case [ PasteData_Keyboard ] use StrDomainElementName parameter to pass row number on which have to paste the data
'										bReturn=Fn_Mech_ValidValuesTable("PasteData_Keyboard","3","","")
'
'										bReturn=Fn_Mech_ValidValuesTable("UndoChanges","","","")
'										bReturn=Fn_Mech_ValidValuesTable("GetAllDisableContextMenu","","","")
'
'										bReturn=Fn_Mech_ValidValuesTable("PasteRowData","1~2~3","","")
'										bReturn=Fn_Mech_ValidValuesTable("CopyData","1~2~3","","")
'										bReturn=Fn_Mech_ValidValuesTable("VerifyRowForegroundColor","1~2~3","red~red~red","")
'										bReturn=Fn_Mech_ValidValuesTable("IsNameAndValueCorrect","7","","")
'										
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												2-May-2012								1.0																						Sunny R
'													Sandeep N												25-May-2012								1.1				Added Case : VerifyCellData				 Sunny R
'													Sandeep N												07-May-2012								1.2				Added Case : GetAllContextMenu,RowColumnCount,PasteData,PasteData_Keyboard				 Sunny R
'													Sandeep N												12-Jun-2012								1.3				Added Case : UndoChanges,GetAllDisableContextMenu				 Sunny R
'													Sandeep N												27-Jul-2012								1.4				Added Case : PasteRowData,CopyData,VerifyRowForegroundColor,IsNameAndValueCorrect				 Anjali M
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_Mech_ValidValuesTable(StrAction,StrDomainElementName,StrValue,StrDescription)
	GBL_FAILED_FUNCTION_NAME="Fn_Mech_ValidValuesTable"
	'Declaring variables
	Dim objParameterDefinitionDialog
	Dim aDomainElementName,aValue,aDescription,scrollMax,iCounter
	Dim bFlag,cellval,iCount,iRow
	Dim objMenu,objChld,StrLabel,crrMenu
	Dim sColourCode,sColour
	Dim iHieght,iTempHieght,iWidth,iX

	Fn_Mech_ValidValuesTable=false
	'checking existance of [ NewParameterDefinition ] dialog
	If Fn_SISW_UI_Object_Operations("Fn_Mech_ValidValuesTable","Exist", Window("MechatronicsWindow").JavaDialog("NewParameterDefinition"), SISW_MICRO_TIMEOUT) = False Then
'	if not Window("MechatronicsWindow").JavaDialog("NewParameterDefinition").Exist(6) then
		Exit function
	else
		'Creating object of [ NewParameterDefinition ] dialog
		set objParameterDefinitionDialog=Window("MechatronicsWindow").JavaDialog("NewParameterDefinition")
	end if
	
	If Fn_SISW_UI_Object_Operations("Fn_Mech_ValidValuesTable","Exist", objParameterDefinitionDialog.JavaSlider("JScrollPane"), SISW_MICRO_TIMEOUT) = True Then
'	If objParameterDefinitionDialog.JavaSlider("JScrollPane").Exist(2) Then
		'Scrolling till end of panel
		scrollMax=objParameterDefinitionDialog.JavaSlider("JScrollPane").GetROProperty("max")
		objParameterDefinitionDialog.JavaSlider("JScrollPane").Drag scrollMax
	End If
	Select Case StrAction
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to set Valid values
		Case "SetData"
			aDomainElementName=Split(StrDomainElementName,"~")
			aValue=Split(StrValue,"~")
			aDescription=Split(StrDescription,"~")
			'Setting DomainElementName | Value | Description
			iHieght= objParameterDefinitionDialog.JavaTable("ValidValues").Object.getRowHeight(0)
			iTempHieght=iHieght/2
			iHieght=iTempHieght

			For iCounter=0 to ubound(aDomainElementName)
				objParameterDefinitionDialog.JavaTable("ValidValues").SelectRow iCounter
				If iCounter<>0 Then
					iHieght=objParameterDefinitionDialog.JavaTable("ValidValues").Object.getRowHeight(0)*iCounter+iTempHieght
				End If

				iWidth= objParameterDefinitionDialog.JavaObject("ValidValuesTableHeader").Object.getTable().getTableHeader().getHeaderRect(0).getWidth()/2
				iX= objParameterDefinitionDialog.JavaObject("ValidValuesTableHeader").Object.getTable().getTableHeader().getHeaderRect(0).getX()
				iWidth=iWidth+iX
				objParameterDefinitionDialog.JavaTable("ValidValues").Click iWidth,iHieght,"LEFT"
				wait 1
				objParameterDefinitionDialog.JavaTable("ValidValues").DblClick iWidth,iHieght,"LEFT"
				If objParameterDefinitionDialog.JavaEdit("ValidValuesTableEdit").Exist(3) Then
				Elseif JavaDialog("NewParameterDefinition").JavaEdit("ValidValuesTableEdit").Exist(2) then
					Set objParameterDefinitionDialog=JavaDialog("NewParameterDefinition")
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail to set Domain name")
					Exit function
				End If
				'Setting Domain Element Name
				objParameterDefinitionDialog.JavaEdit("ValidValuesTableEdit").Set aDomainElementName(iCounter)
				objParameterDefinitionDialog.JavaEdit("ValidValuesTableEdit").Activate
				wait 1
				set objParameterDefinitionDialog=Window("MechatronicsWindow").JavaDialog("NewParameterDefinition")
				wait 1

				'Setting Value
				If aValue(iCounter)<>"" Then
					iWidth= objParameterDefinitionDialog.JavaObject("ValidValuesTableHeader").Object.getTable().getTableHeader().getHeaderRect(1).getWidth()/2
					iX= objParameterDefinitionDialog.JavaObject("ValidValuesTableHeader").Object.getTable().getTableHeader().getHeaderRect(1).getX()
					iWidth=iWidth+iX
					objParameterDefinitionDialog.JavaTable("ValidValues").Click iWidth,iHieght,"LEFT"
					wait 1
					objParameterDefinitionDialog.JavaTable("ValidValues").DblClick iWidth,iHieght,"LEFT"
					If objParameterDefinitionDialog.JavaEdit("ValidValuesTableEdit").Exist(3) Then
					Elseif JavaDialog("NewParameterDefinition").JavaEdit("ValidValuesTableEdit").Exist(2) then
						Set objParameterDefinitionDialog=JavaDialog("NewParameterDefinition")
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail to set Domain value")
						Exit function
					End If
					objParameterDefinitionDialog.JavaEdit("ValidValuesTableEdit").Set aValue(iCounter)
					objParameterDefinitionDialog.JavaEdit("ValidValuesTableEdit").Activate
					wait 1
					set objParameterDefinitionDialog=Window("MechatronicsWindow").JavaDialog("NewParameterDefinition")  
				End If
				'Setting Description
				If aDescription(iCounter)<>"" Then
					iWidth= objParameterDefinitionDialog.JavaObject("ValidValuesTableHeader").Object.getTable().getTableHeader().getHeaderRect(2).getWidth()/2
					iX= objParameterDefinitionDialog.JavaObject("ValidValuesTableHeader").Object.getTable().getTableHeader().getHeaderRect(2).getX()
					iWidth=iWidth+iX
					objParameterDefinitionDialog.JavaTable("ValidValues").Click iWidth,iHieght,"LEFT"
					wait 1
					objParameterDefinitionDialog.JavaTable("ValidValues").DblClick iWidth,iHieght,"LEFT"
					If objParameterDefinitionDialog.JavaEdit("ValidValuesTableDescEdit").Exist(3) Then
					Elseif JavaDialog("NewParameterDefinition").JavaEdit("ValidValuesTableDescEdit").Exist(3) then
						Set objParameterDefinitionDialog=JavaDialog("NewParameterDefinition")
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail to set Domain Description")
						Exit function
					End If
					objParameterDefinitionDialog.JavaEdit("ValidValuesTableDescEdit").Set aDescription(iCounter)+vbLf
					wait 1
					set objParameterDefinitionDialog=Window("MechatronicsWindow").JavaDialog("NewParameterDefinition")
				End If
			Next
			Fn_Mech_ValidValuesTable=true
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "VerifyCellData"
			aDomainElementName=Split(StrDomainElementName,"~")
			aValue=Split(StrValue,"~")
			aDescription=Split(StrDescription,"~")
			For iCounter=0 to ubound(aDomainElementName)
				bFlag=false
				iRow=objParameterDefinitionDialog.JavaTable("ValidValues").GetROProperty("rows")
				For iCount=0 to iRow-1
					cellval=objParameterDefinitionDialog.JavaTable("ValidValues").Object.getCellData(iCount,0).toString()
					If instr(1,cellval,"[") Then cellval=Replace(cellval,"[","") end if
					If instr(1,cellval,"]") Then cellval=Replace(cellval,"]","") end if
					cellval=Split(cellval,",")
					If trim(cellval(2))=aDomainElementName(iCounter) Then
						bFlag=true
						If StrValue<>"" Then
							'Verifing Domain value
							If aValue(iCounter)<>"" Then
								If trim(cellval(4))=aValue(iCounter) Then
									bFlag=true
								else
									bFlag=false
									Exit for		
								end if
							end if
						End If
						If StrDescription<>"" Then
							'Verifing Domain description
							If aDescription(iCounter)<>"" Then
									If trim(cellval(6))=aDescription(iCounter) Then
										bFlag=true
									else
										bFlag=false		
										Exit for
									end if
							End If
						End If
						Exit for
					End If
				Next
				If bFlag=false Then
					Exit for
				End If
			Next
			If bFlag=true Then
				Fn_Mech_ValidValuesTable=true
			End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			'Case to paste data in [ ValidValues ] table
			Case "PasteData","PasteData_Ext"
				If StrDomainElementName<>"" Then
					objParameterDefinitionDialog.JavaTable("ValidValues").SelectRow Cint(StrDomainElementName)-1
					wait 1
					objParameterDefinitionDialog.JavaTable("ValidValues").ClickCell Cint(StrDomainElementName)-1,0,"RIGHT"
					wait 1
				else
					objParameterDefinitionDialog.JavaTable("ValidValues").ClickCell 0,0,"RIGHT"
					wait 1
				End If
				objParameterDefinitionDialog.JavaMenu("index:=0","label:=Paste").Select
				If Err.Number < 0 Then
					Fn_Mech_ValidValuesTable=False
				Else
					Fn_Mech_ValidValuesTable=True
				End If
				If StrAction<>"PasteData_Ext" Then
					If StrDomainElementName<>"" Then
						objParameterDefinitionDialog.JavaTable("ValidValues").ClickCell Cint(StrDomainElementName)-1,0
						wait 1
						objParameterDefinitionDialog.JavaTable("ValidValues").DeselectRow Cint(StrDomainElementName)-1
					else
						objParameterDefinitionDialog.JavaTable("ValidValues").ClickCell 0,0
						wait 1
						objParameterDefinitionDialog.JavaTable("ValidValues").DeselectRow 0
					End If
				End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			'Case to paste data in [ ValidValues ] table
			Case "PasteData_Keyboard"
				If StrDomainElementName<>"" Then
					objParameterDefinitionDialog.JavaTable("ValidValues").SelectRow Cint(StrDomainElementName)-1
				else
					objParameterDefinitionDialog.JavaTable("ValidValues").SelectRow 0
				End If
				objParameterDefinitionDialog.JavaTable("ValidValues").PressKey "V",micCtrl
				If Err.Number < 0 Then
					Fn_Mech_ValidValuesTable=False
				Else
					Fn_Mech_ValidValuesTable=True
				End If

				If StrDomainElementName<>"" Then
					objParameterDefinitionDialog.JavaTable("ValidValues").ClickCell Cint(StrDomainElementName)-1,0
					wait 1
					objParameterDefinitionDialog.JavaTable("ValidValues").DeselectRow Cint(StrDomainElementName)-1
				else
					objParameterDefinitionDialog.JavaTable("ValidValues").ClickCell 0,0
					wait 1
					objParameterDefinitionDialog.JavaTable("ValidValues").DeselectRow 0
				End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to get all available Context menu for [ Valid Values ] table
		Case "GetAllContextMenu"
				If StrDomainElementName<>"" Then
					objParameterDefinitionDialog.JavaTable("ValidValuesRows").ClickCell Cint(StrDomainElementName)-1,0,"RIGHT"
				else
					objParameterDefinitionDialog.JavaTable("ValidValues").ClickCell 0,0,"RIGHT"
				End If
				Set objMenu=Description.Create
				objMenu("Class Name").value="JavaMenu"
				Set objChld=objParameterDefinitionDialog.ChildObjects(objMenu)
				crrMenu=""
				For iCounter=0 to objChld.count-1
						If iCounter=0 Then
							crrMenu=objChld(0).GetROProperty("label")
						else
							crrMenu=crrMenu+"~"+objChld(iCounter).GetROProperty("label")
						End If
				Next
				If Err.Number < 0 Then
					Fn_Mech_ValidValuesTable=False
				Else
					If crrMenu<>"" Then
						Fn_Mech_ValidValuesTable=crrMenu
					else
						Fn_Mech_ValidValuesTable=False
					End If
				End If
				If StrDomainElementName<>"" Then
					objParameterDefinitionDialog.JavaTable("ValidValues").ClickCell Cint(StrDomainElementName)-1,0
					wait 1
					objParameterDefinitionDialog.JavaTable("ValidValues").DeselectRow Cint(StrDomainElementName)-1
				else
					objParameterDefinitionDialog.JavaTable("ValidValues").ClickCell 0,0
					wait 1
					objParameterDefinitionDialog.JavaTable("ValidValues").DeselectRow 0
				End If
				Set objChld=Nothing
		 '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to get Row ~ Column count currently exist in [ Maximum Values ] table
		Case "RowColumnCount"
				iRows=objParameterDefinitionDialog.JavaTable("ValidValues").GetROProperty("rows")
				iCols=objParameterDefinitionDialog.JavaObject("InitialValuesTableHeader").Object.getColumnModel().getColumnCount()
				Fn_Mech_ValidValuesTable=iRows+"~"+iCols
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			'Case to Undo Changes in [ Valid Values ] table
			Case "UndoChanges"
					objParameterDefinitionDialog.JavaTable("ValidValues").ClickCell 0,0,"RIGHT"
					wait 1
					objParameterDefinitionDialog.JavaMenu("index:=0","label:=Undo").Select
					If Err.Number < 0 Then
						Fn_Mech_ValidValuesTable=False
					Else
						Fn_Mech_ValidValuesTable=True
					End If
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			'Case to get all Disable Context menu for [ Valid Values ] table
			Case "GetAllDisableContextMenu"
						objParameterDefinitionDialog.JavaTable("ValidValues").ClickCell 0,0,"RIGHT"
						wait 1
						Set objMenu=Description.Create
						objMenu("Class Name").value="JavaMenu"
						Set objChld=objParameterDefinitionDialog.ChildObjects(objMenu)
						StrLabel=""
						For iCounter=0 to objChld.count-1
								crrMenu=objChld(iCounter).GetROProperty("label")
								If objChld(iCounter).CheckProperty("enabled",1,1)=false then
									If StrLabel="" Then
										StrLabel=objChld(iCounter).GetROProperty("label")
									else
										StrLabel=StrLabel+"~"+objChld(iCounter).GetROProperty("label")
									End If
								end if
						Next

						If Err.Number < 0 Then
							Fn_Mech_ValidValuesTable=False
						Else
							If StrLabel<>"" Then
								Fn_Mech_ValidValuesTable=StrLabel
							else
								Fn_Mech_ValidValuesTable=False
							End If
						End If
						wait 1
						objParameterDefinitionDialog.JavaTable("ValidValues").DeselectRow 0
						Set objChld=Nothing
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to check current value and name is correct or not
		Case "IsNameAndValueCorrect"
			If objParameterDefinitionDialog.JavaTable("ValidValues").Object.getValueAt(Cint(StrDomainElementName)-1,0).isNameCorrect()="true" and objParameterDefinitionDialog.JavaTable("ValidValues").Object.getValueAt(Cint(StrDomainElementName)-1,0).isValueCorrect()="true" then
				Fn_Mech_ValidValuesTable=true
			else
				Fn_Mech_ValidValuesTable=false
			end if
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to verify Foreground color of specific cell text
		Case "VerifyRowForegroundColor"
            aDomainElementName=Split(StrDomainElementName,"~")
			aValue=Split(StrValue,"~")
			For iCounter=0 to ubound(aDomainElementName)
				bFlag=false
				objParameterDefinitionDialog.JavaTable("ValidValues").SelectRow cint(aDomainElementName(iCounter))-1
				wait 1
				sColourCode=""
				sColour=objParameterDefinitionDialog.JavaTable("ValidValues").Object.getInnerTableCellRenderer().getForeground().toString() 
				sColour =  mid(sColour ,instr(sColour ,"[")  ,instr(sColour ,"]") )
				Select Case lCase(aValue(iCounter))
					Case "red"
						sColourCode="[r=255,g=0,b=0]"
				End Select
				If sColour=sColourCode Then
					bFlag=true
				else
					Exit for
				End if
			Next
			 If bFlag=true Then
				Fn_Mech_ValidValuesTable=true
			else
				Fn_Mech_ValidValuesTable=false
			End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			'Case to Copy data in multiple rows of [ ValidValues ] table
			Case "CopyData"
				If StrDomainElementName<>"" Then
					aDomainElementName=split(StrDomainElementName,"~")
					objParameterDefinitionDialog.JavaTable("ValidValues").SelectRow Cint(aDomainElementName(0))-1
					For iCounter=1 to ubound(aDomainElementName)
						objParameterDefinitionDialog.JavaTable("ValidValues").ExtendRow Cint(aDomainElementName(iCounter))-1
						wait 1
					Next
					objParameterDefinitionDialog.JavaTable("ValidValues").ClickCell Cint(aDomainElementName(ubound(aDomainElementName)))-1,0,"RIGHT"
						wait 1
				End If
				objParameterDefinitionDialog.JavaMenu("index:=0","label:=Copy").Select
				If Err.Number < 0 Then
					Fn_Mech_ValidValuesTable=False
				Else
					Fn_Mech_ValidValuesTable=True
				End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			'Case to paste data in multiple rows of [ ValidValues ] table
			Case "PasteRowData"
				If StrDomainElementName<>"" Then
					aDomainElementName=split(StrDomainElementName,"~")
					objParameterDefinitionDialog.JavaTable("ValidValues").SelectRow Cint(aDomainElementName(0))-1
					For iCounter=1 to ubound(aDomainElementName)
						objParameterDefinitionDialog.JavaTable("ValidValues").ExtendRow Cint(aDomainElementName(iCounter))-1
						wait 1
					Next
					objParameterDefinitionDialog.JavaTable("ValidValues").ClickCell Cint(aDomainElementName(ubound(aDomainElementName)))-1,0,"RIGHT"
						wait 1
				End If
				objParameterDefinitionDialog.JavaMenu("index:=0","label:=Paste").Select
				If Err.Number < 0 Then
					Fn_Mech_ValidValuesTable=False
				Else
					Fn_Mech_ValidValuesTable=True
				End If
	End Select
	'Releasing object of [ NewParameterDefinition ] dialog
	set objParameterDefinitionDialog=nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_Mech_SEDInitialValueTable

'Description			 :	Function Used to perform operation on InitialValueTable of parameter Defination [ ParmDefSED ] Revision Information

'Parameters			   :   1.StrAction: Action Name
'									 2.StrDomainElementName: Domain Element name
'									 3.StrValue: Values
'
'Return Value		   : 	True or False

'Pre-requisite			:	Parameter Defination dialog should be appear on screen on user should be on Define Parameter Defination Revision Information dialog

'Examples				:   bReturn=Fn_Mech_SEDInitialValueTable("SetData","Ele1","0x36")
'												bReturn=Fn_Mech_SEDInitialValueTable("SetData","","0x24")
'												bReturn=Fn_Mech_SEDInitialValueTable("SetData","Ele2","")
'
'										bReturn=Fn_Mech_SEDInitialValueTable("VerifyData","Element 8","0x1249B~A new Description for Element 8")
'										bReturn=Fn_Mech_SEDInitialValueTable("VerifyData","Element 8","0x1249B")
'
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												2-May-2012								1.0																						Sunny R
'													Sandeep N												19-June-2012								1.1				Added Case : VerifyData																		Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_Mech_SEDInitialValueTable(StrAction,StrDomainElementName,StrValue)
	GBL_FAILED_FUNCTION_NAME="Fn_Mech_SEDInitialValueTable"
	'Declaring variables
	Dim objParameterDefinitionDialog
	Dim scrollMax,bFlag,aValue

	Fn_Mech_SEDInitialValueTable=false
	'checking existance of [ NewParameterDefinition ] dialog
	If Fn_SISW_UI_Object_Operations("Fn_Mech_SEDInitialValueTable","Exist", Window("MechatronicsWindow").JavaDialog("NewParameterDefinition"), SISW_MINLESS_TIMEOUT) = False Then
'	if not Window("MechatronicsWindow").JavaDialog("NewParameterDefinition").Exist(6) then
		Exit function
	else
		'Creating object of [ NewParameterDefinition ] dialog
		set objParameterDefinitionDialog=Window("MechatronicsWindow").JavaDialog("NewParameterDefinition")
	end if
	If Fn_SISW_UI_Object_Operations("Fn_Mech_SEDInitialValueTable","Exist", objParameterDefinitionDialog.JavaSlider("JScrollPane"), SISW_MINLESS_TIMEOUT) = True Then
'	If objParameterDefinitionDialog.JavaSlider("JScrollPane").Exist(3) Then
		'Scrolling till end of panel
		scrollMax=objParameterDefinitionDialog.JavaSlider("JScrollPane").GetROProperty("max")
		objParameterDefinitionDialog.JavaSlider("JScrollPane").Drag scrollMax
	End If
	Select Case StrAction
		'Case to set Valid values
		Case "SetData"
				'Selecting Domain Element Name
				If StrDomainElementName<>"" Then
					objParameterDefinitionDialog.JavaTable("SEDInitialValue").SetCellData 0,"Domain Element Name",StrDomainElementName
					wait 1
				End If
                'Selecting Value
				If StrValue<>"" Then
					objParameterDefinitionDialog.JavaTable("SEDInitialValue").SetCellData 0,"Value",StrValue
					wait 1
				End If
				Fn_Mech_SEDInitialValueTable=true
		'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		'case to verify Initial values
		Case "VerifyData"
			bFlag=true
			If StrDomainElementName<>"" Then
				bFlag=false
				If objParameterDefinitionDialog.JavaTable("SEDInitialValue").GetCellData(0,"Domain Element Name")=StrDomainElementName then
					bFlag=true
				else
					set objParameterDefinitionDialog=nothing
					Exit function
				End if
			End If
			aValue=Split(StrValue,"~")
			If aValue(0)<>"" Then
				bFlag=false
				If objParameterDefinitionDialog.JavaTable("SEDInitialValue").GetCellData(0,"Value")=aValue(0) then
					bFlag=true
				else
					set objParameterDefinitionDialog=nothing
					Exit function
				End if
			End If
			If ubound(aValue)=1 Then
				bFlag=false
				If objParameterDefinitionDialog.JavaTable("SEDInitialValue").GetCellData(0,"Description")=aValue(1) then
					bFlag=true
				else
					set objParameterDefinitionDialog=nothing
					Exit function
				End if
			End If
			Fn_Mech_SEDInitialValueTable=true
	End Select
	'Releasing object of [ NewParameterDefinition ] dialog
	set objParameterDefinitionDialog=nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_Mech_BitDefination

'Description			 :	Function Used to perform operation on Bit Defination table to define Bits of parameter Defination [ ParmDefBitDef ] Revision Information

'Parameters			   :   1.StrAction: Action Name
'									 	2.StrByte: Byte number
'									 	3.StrBitNumber: Bit number of specific Byte
'									 	4.StrName: Bit name of specific Byte
'									 	5.Str0Meaning: Bit 0 Meaning of specific Byte
'									 	6.Str1Meaning: Bit 1 Meaning of specific Byte
'
'Return Value		   : 	True or False

'Pre-requisite			:	Parameter Defination dialog should be appear on screen on user should be on Define Parameter Defination Revision Information dialog

'Examples				:   ByteInfo="1~1~1~1~1~1~1~1~2~2~2~2~2~2~2~2~3~3~3~3~3~3~3~3"
'										BitNumberInfo="7~6~5~4~3~2~1~0~7~6~5~4~3~2~1~0~7~6~5~4~3~2~1~0"
'										NameInfo="B1N1~B1N2~B1N3~B1N4~B1N5~B1N6~B1N7~B1N8~B2N1~B2N2~B2N3~B2N4~B2N5~B2N6~B2N7~B2N8~B3N1~B3N2~B3N3~B3N4~B3N5~B3N6~B3N7~B3N8"
'										Meaning0Info="0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0"
'										Meaning1Info="1~1~1~1~1~1~1~1~1~1~1~1~1~1~1~1~1~1~1~1~1~1~1~1"
'										
'										ByteInfo="Byte Number~Byte Number"
'										BitNumberInfo="Bit Number~Bit Number"
'										NameInfo="Bit Name~Bit Name"
'										Meaning0Info="Bit 0 Meaning~Bit 0 Meaning"
'										Meaning1Info="Bit 1 Meaning~Bit 1 Meaning"
'										bReturn=Fn_Mech_BitDefination("SetData",ByteInfo,BitNumberInfo,NameInfo,Meaning0Info,Meaning1Info)
'
'										ColName="Bit Number~Bit Number~Bit Number~Bit Number~Bit Number~Bit Number~Bit Number~Bit Number~Bit Number~Bit Number~Bit Number~Bit Number~Bit Number~Bit Number~Bit Number~Bit Number~Bit Number~Bit Number~Bit Number~Bit Number~Bit Number~Bit Number~Bit Number~Bit Number"
'										ByteInfo="1~1~1~1~1~1~1~1~2~2~2~2~2~2~2~2~3~3~3~3~3~3~3~3"
'										BitNumberInfo="7~6~5~4~3~2~1~0~7~6~5~4~3~2~1~0~7~6~5~4~3~2~1~0"
'										bReturn= Fn_Mech_BitDefination("VerifyCellEditable",ByteInfo,BitNumberInfo,ColName,"","")
'										IMP Note : For case VerifyCellEditable Use StrName parameter to pass column names
'
'										bReturn=Fn_Mech_BitDefination("VerifyData","1~1~1","7~6~5","","off~on~dim","no~yes~active")
'
'										bReturn=Fn_Mech_BitDefination("CopyRows","1~1~1~1","7~6~5~4","","","")
'										bReturn=Fn_Mech_BitDefination("PasteRows","2","7","","","")
'										bReturn=Fn_Mech_BitDefination("CopyRows_Keyboard","1~1~1~1","7~6~5~4","","","")
'										bReturn=Fn_Mech_BitDefination("PasteRows_Keyboard","2","7","","","")
'
'										bReturn=Fn_Mech_BitDefination("GetAllContextMenu","","","","","")
'										bReturn=Fn_Mech_BitDefination("GetAllContextMenu","2","3","","","")
'										bReturn=Fn_Mech_BitDefination("UndoChanges","","","","","")
'										bReturn=Fn_Mech_BitDefination("GetAllDisableContextMenu","","","","","")
'
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												3-May-2012								1.0																						Sunny R
'													Sandeep N												31-May-2012							   1.1																						Sunny R
'													Sandeep N												01-Jun-2012							   1.2						Add Case : VerifyData						Sunny R
'													Sandeep N												06-Jun-2012							   1.3						Add Case : "CopyRows","PasteRows","CopyRows_Keyboard","PasteRows_Keyboard"						Sunny R
'													Sandeep N												07-Jun-2012							   1.4						Add Case : GetAllContextMenu						Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'For Case VerifyCellEditable : Use [ StrName ] parameter to pass column name
Function Fn_Mech_BitDefination(StrAction,StrByte,StrBitNumber,StrName,Str0Meaning,Str1Meaning)
	GBL_FAILED_FUNCTION_NAME="Fn_Mech_BitDefination"
	'Declaring variables
	Dim objParameterDefinitionDialog
	Dim scrollMax,iCounter,aByte,aBitNumber,aName,a0Meaning,a1Meaning,iRow,PrevRowNumber
	Dim aValue,aCol,iCol,crrName,crr0Meaning,crr1Meaning,bFlag
	Dim objMenu,objChld,crrMenu,StrLabel

	Fn_Mech_BitDefination=false
	'checking existance of [ NewParameterDefinition ] dialog
	if not Window("MechatronicsWindow").JavaDialog("NewParameterDefinition").Exist(6) then
		Exit function
	else
		'Creating object of [ NewParameterDefinition ] dialog
		set objParameterDefinitionDialog=Window("MechatronicsWindow").JavaDialog("NewParameterDefinition")
	end if
	If objParameterDefinitionDialog.JavaSlider("JScrollPane").Exist(3) Then
		'Scrolling till end of panel
		scrollMax=objParameterDefinitionDialog.JavaSlider("JScrollPane").GetROProperty("max")
		objParameterDefinitionDialog.JavaSlider("JScrollPane").Drag scrollMax
	End If
	Select Case StrAction
		Case "SetData"
			If objParameterDefinitionDialog.JavaCheckBox("BitDefinitionTableCollapse").Exist(2) Then
				Call Fn_CheckBox_Set("Fn_Mech_BitDefination", objParameterDefinitionDialog, "BitDefinitionTableCollapse", "on")
			End If
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			'Spliting Byte information in array
			aByte=Split(StrByte,"~")
			aBitNumber=Split(StrBitNumber,"~")
			aName=Split(StrName,"~")
			a0Meaning=Split(Str0Meaning,"~")
			a1Meaning=Split(Str1Meaning,"~")
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 			
			For iCounter=0 to ubound(aByte)
				iRow=Cint(aByte(iCounter))*8-CInt(aBitNumber(iCounter))-1
				'setting Bit name
				If aName(iCounter)<>"" Then
					objParameterDefinitionDialog.JavaTable("CCDMBitDefTable").SetCellData iRow,"Name",aName(iCounter)
				End If
				'setting 0 Meaning
				If a0Meaning(iCounter)<>"" Then
					objParameterDefinitionDialog.JavaTable("CCDMBitDefTable").SetCellData iRow,"""0"" Meaning",a0Meaning(iCounter)
				End If
				'setting 1 Meaning
				If a1Meaning(iCounter)<>"" Then
					objParameterDefinitionDialog.JavaTable("CCDMBitDefTable").SetCellData iRow,"""1"" Meaning",a1Meaning(iCounter)
				End If
			Next
			if Err.Number < 0 Then
					Fn_Mech_BitDefination=false
			else
					Fn_Mech_BitDefination=true
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "VerifyCellEditable"
			If objParameterDefinitionDialog.JavaCheckBox("BitDefinitionTableCollapse").Exist(2) Then
				Call Fn_CheckBox_Set("Fn_Mech_BitDefination", objParameterDefinitionDialog, "BitDefinitionTableCollapse", "on")
			End If
			aCol=Split(StrName,"~")
			'Spliting Byte information in array
			aByte=Split(StrByte,"~")
			aBitNumber=Split(StrBitNumber,"~")
			For iCounter=0 to ubound(aByte)
				Select Case aCol(iCounter)
					Case "Bit Number"
						iCol=0
					Case "Name"
						iCol=1
				End Select
				iRow=Cint(aByte(iCounter))*8-CInt(aBitNumber(iCounter))-1
				If iCounter=0 Then
					aValue=objParameterDefinitionDialog.JavaTable("CCDMBitDefTable").Object.isCellEditable(iRow,iCol)
				else
					aValue=aValue+"~"+objParameterDefinitionDialog.JavaTable("CCDMBitDefTable").Object.isCellEditable(iRow,iCol)
				End If
			Next
			Fn_Mech_BitDefination=aValue
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "VerifyData"
			If objParameterDefinitionDialog.JavaCheckBox("BitDefinitionTableCollapse").Exist(2) Then
				Call Fn_CheckBox_Set("Fn_Mech_BitDefination", objParameterDefinitionDialog, "BitDefinitionTableCollapse", "on")
			End If
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			'Spliting Byte information in array
			aByte=Split(StrByte,"~")
			aBitNumber=Split(StrBitNumber,"~")
			If StrName<>"" Then
				aName=Split(StrName,"~")
			End If
			If Str0Meaning<>"" Then
				a0Meaning=Split(Str0Meaning,"~")
			End If
			If Str1Meaning<>"" Then
				a1Meaning=Split(Str1Meaning,"~")
			End If
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 		
			For iCounter=0 to ubound(aByte)
				bFlag=false
				iRow=Cint(aByte(iCounter))*8-CInt(aBitNumber(iCounter))-1
				'getting Bit name
				If StrName<>"" Then
					bFlag=false
					crrName=objParameterDefinitionDialog.JavaTable("CCDMBitDefTable").GetCellData(iRow,"Name")
					If CStr(crrName)=CStr(aName(iCounter)) Then
						bFlag=true
					else
						Exit for
					End If
				End If
				'getting 0 Meaning
				If Str0Meaning<>"" Then
					bFlag=false
					crr0Meaning=objParameterDefinitionDialog.JavaTable("CCDMBitDefTable").GetCellData(iRow,"""0"" Meaning")
					If crr0Meaning=a0Meaning(iCounter) Then
						bFlag=true
					else
						Exit for
					End If
				End If
				'getting 1 Meaning
				If Str1Meaning<>"" Then
					bFlag=false
					crr1Meaning=objParameterDefinitionDialog.JavaTable("CCDMBitDefTable").GetCellData(iRow,"""1"" Meaning")
					If crr1Meaning=a1Meaning(iCounter) Then
						bFlag=true
					else
						Exit for
					End If
				End If
				If bFlag=false Then
					Exit for
				End If
			Next
			If bFlag=true Then
				Fn_Mech_BitDefination=true
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "CopyRows","PasteRows","CopyRows_Keyboard","PasteRows_Keyboard"
			If objParameterDefinitionDialog.JavaCheckBox("BitDefinitionTableCollapse").Exist(2) Then
				Call Fn_CheckBox_Set("Fn_Mech_BitDefination", objParameterDefinitionDialog, "BitDefinitionTableCollapse", "on")
			End If
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			'Spliting Byte information in array
			aByte=Split(StrByte,"~")
			aBitNumber=Split(StrBitNumber,"~")
			aName=Split(StrName,"~")
			a0Meaning=Split(Str0Meaning,"~")
			a1Meaning=Split(Str1Meaning,"~")
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 			
			iRow=Cint(aByte(0))*8-CInt(aBitNumber(0))-1
			objParameterDefinitionDialog.JavaTable("CCDMBitDefTable").SelectRow iRow
			For iCounter=1 to ubound(aByte)
				iRow=Cint(aByte(iCounter))*8-CInt(aBitNumber(iCounter))-1
				objParameterDefinitionDialog.JavaTable("CCDMBitDefTable").ExtendRow iRow
			Next
			If objParameterDefinitionDialog.JavaCheckBox("BitDefinitionTableExpand").Exist(2) Then
				Call Fn_CheckBox_Set("Fn_Mech_BitDefination", objParameterDefinitionDialog, "BitDefinitionTableExpand", "off")
			End If
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 					
			Select Case StrAction
				' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				Case "CopyRows"
					objParameterDefinitionDialog.JavaTable("CCDMBitDefTable").Click 1,1,"RIGHT"
					objParameterDefinitionDialog.JavaMenu("index:=0","label:=Copy").Select
				' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				Case "PasteRows"
					iRow=Cint(aByte(0))*8-CInt(aBitNumber(0))-1
					If objParameterDefinitionDialog.JavaSlider("JScrollPane").Exist(3) Then
						'Scrolling till end of panel
						scrollMax=objParameterDefinitionDialog.JavaSlider("JScrollPane").GetROProperty("max")
						objParameterDefinitionDialog.JavaSlider("JScrollPane").Drag scrollMax
					End If
					objParameterDefinitionDialog.JavaTable("CCDMBitDefTable").ClickCell iRow,"Bit Number","RIGHT"
					objParameterDefinitionDialog.JavaMenu("index:=0","label:=Paste").Select
				' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				Case "CopyRows_Keyboard"
					objParameterDefinitionDialog.JavaTable("CCDMBitDefTable").PressKey "C",micCtrl
				' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				Case "PasteRows_Keyboard"
					iRow=Cint(aByte(0))*8-CInt(aBitNumber(0))-1
					If objParameterDefinitionDialog.JavaSlider("JScrollPane").Exist(3) Then
						'Scrolling till end of panel
						scrollMax=objParameterDefinitionDialog.JavaSlider("JScrollPane").GetROProperty("max")
						objParameterDefinitionDialog.JavaSlider("JScrollPane").Drag scrollMax
					End If
					objParameterDefinitionDialog.JavaTable("CCDMBitDefTable").SelectRow iRow
					objParameterDefinitionDialog.JavaTable("CCDMBitDefTable").PressKey "V",micCtrl
				' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			End Select
			If Err.Number < 0 Then
				Fn_Mech_BitDefination=False
			Else
				Fn_Mech_BitDefination=True
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get All available Context menu
		Case "GetAllContextMenu"
			If StrByte<>"" and  StrBitNumber<>"" Then
				iRow=Cint(StrByte)*8-CInt(StrBitNumber)-1
			End If
			If StrByte<>"" and StrBitNumber<>"" Then
				objParameterDefinitionDialog.JavaTable("CCDMBitDefTable").SelectRow iRow
				objParameterDefinitionDialog.JavaTable("CCDMBitDefTable").ClickCell iRow,"Name","RIGHT"
			else
				objParameterDefinitionDialog.JavaTable("CCDMBitDefTable").ClickCell 0,"Name","RIGHT"
			End If
			Set objMenu=Description.Create
			objMenu("Class Name").value="JavaMenu"
			Set objChld=objParameterDefinitionDialog.ChildObjects(objMenu)
			crrMenu=""
			For iCounter=0 to objChld.count-1
					If iCounter=0 Then
						crrMenu=objChld(0).GetROProperty("label")
					else
						crrMenu=crrMenu+"~"+objChld(iCounter).GetROProperty("label")
					End If
			Next
			If Err.Number < 0 Then
				Fn_Mech_BitDefination=False
			Else
				If crrMenu<>"" Then
					Fn_Mech_BitDefination=crrMenu
				else
					Fn_Mech_BitDefination=False
				End If
			End If
			If StrByte<>"" and StrBitNumber<>"" Then
				objParameterDefinitionDialog.JavaTable("CCDMBitDefTable").ClickCell 0,"Bit Number"
				objParameterDefinitionDialog.JavaTable("CCDMBitDefTable").DeselectRow 0
			else
				objParameterDefinitionDialog.JavaTable("CCDMBitDefTable").ClickCell 0,"Bit Number"
			end if
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			'Case to Undo Changes in [ Bit Defination ] table
			Case "UndoChanges"
						objParameterDefinitionDialog.JavaTable("CCDMBitDefTable").ClickCell 0,"Bit Number","RIGHT"
						wait 1
						objParameterDefinitionDialog.JavaMenu("index:=0","label:=Undo").Select
						If Err.Number < 0 Then
							Fn_Mech_BitDefination=False
						Else
							Fn_Mech_BitDefination=True
						End If
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			'Case to get all Disable Context menu for [ Bit Defination ] table
			Case "GetAllDisableContextMenu"
						objParameterDefinitionDialog.JavaTable("CCDMBitDefTable").ClickCell 0,"Bit Number","RIGHT"
						wait 1
						Set objMenu=Description.Create
						objMenu("Class Name").value="JavaMenu"
						Set objChld=objParameterDefinitionDialog.ChildObjects(objMenu)
						StrLabel=""
						For iCounter=0 to objChld.count-1
								crrMenu=objChld(iCounter).GetROProperty("label")
								If objChld(iCounter).CheckProperty("enabled",1,1)=false then
									If StrLabel="" Then
										StrLabel=objChld(iCounter).GetROProperty("label")
									else
										StrLabel=StrLabel+"~"+objChld(iCounter).GetROProperty("label")
									End If
								end if
						Next

						If Err.Number < 0 Then
							Fn_Mech_BitDefination=False
						Else
							If StrLabel<>"" Then
								Fn_Mech_BitDefination=StrLabel
							else
								Fn_Mech_BitDefination=False
							End If
						End If
						objParameterDefinitionDialog.JavaTable("CCDMBitDefTable").ClickCell 0,"Name"
						Set objChld=Nothing
	End Select
	'Releasing object of [ NewParameterDefinition ] dialog
	 set objParameterDefinitionDialog=nothing
End Function

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_Mech_ValidationErrorHandle

'Description			 :	Function Used to handle validation errors

'Parameters			   :   1.StrAction: Action Name
'									 	2.StrErrorMsg : Error Message
'									 	3.btnName: Button N
'
'Return Value		   : 	True or False

'Pre-requisite			:	Validation dialog should be open

'Examples				:   bReturn=Fn_Mech_ValidationErrorHandle("Resolution_Error","Value entered results in validation violations","ClearInvalidTableCells")

'										For Direct Value
'										bReturn=Fn_Mech_ValidationErrorHandle("Table_Error","Value entered results in validation violations Invalid values will be cleared and need to be entered again.","OK")
'
'										After Clicking on More Button
'										bReturn= Fn_Mech_ValidationErrorHandle("Table_Error","The table can be collapsed only when all the values in the table~Error","OK")
'										StrErrorMsg = Err msg ~ Dialog name
'
'										For Detail Value By Clicking on More
'										bReturn=Fn_Mech_ValidationErrorHandle("Table_Error","Maximum Values[2][1],Reason = Resolution based validation failed  Maximum Values[3][1],Reason = Resolution based validation failed~ Detail","OK")
'										bReturn=Fn_Mech_ValidationErrorHandle("ValidationConfirmation_Error","Following domain element name entered already exists: GG","Cancel")
'										bReturn=Fn_Mech_ValidationErrorHandle("ValidationConfirmation_Error_Ext","Following domain element name entered already exists: GG","Cancel")
'										bReturn = Fn_Mech_ValidationErrorHandle("ValidationConfirmation_DoubleDialog","2~Validation Confirmation","Yes")
'
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done															Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Pranav Ingle												7-May-2012								1.0																																														Sunny R
'													Sandeep N												   22-May-2012							   1.1					Added Case : ValidationConfirmation_Error					Sunny R
'													Pranav Ingle											   24-May-2012							   1.2					Added Case : ValidationConfirmation_Error_Ext			Sunny R
'													Pranav Ingle											   15-Jun-2012							   1.3					Added Case : ValidationConfirmation_DoubleDialog			Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_Mech_ValidationErrorHandle(StrAction,StrErrorMsg,btnName)
	GBL_FAILED_FUNCTION_NAME="Fn_Mech_ValidationErrorHandle"
	GBL_EXPECTED_MESSAGE=StrErrorMsg
   Dim sDispMsg,bReturn,objErrorDialog,arrErrorMsg,iCount
	Select Case StrAction
			Case "Resolution_Error"
				Set objErrorDialog=JavaDialog("ValidationError")
                objErrorDialog.JavaStaticText("ErrMsg").SetTOProperty "label",StrErrorMsg

				If  objErrorDialog.JavaStaticText("ErrMsg").Exist(3) Then
					sDispMsg = objErrorDialog.JavaStaticText("ErrMsg").GetROProperty("value")
				Else
					sDispMsg = objErrorDialog.JavaObject("ErrorMsg").Object.getText
				End If

				If btnName<>"" Then
					Call Fn_Button_Click("Fn_Mech_ValidationErrorHandle", objErrorDialog,btnName)	
				End If

				If  Instr(1,sDispMsg,StrErrorMsg)>0 Then
					Fn_Mech_ValidationErrorHandle=True
				Else
					GBL_ACTUAL_MESSAGE=sDispMsg
					Fn_Mech_ValidationErrorHandle=False
				End If

			Case "Table_Error"
				If Fn_SISW_UI_Object_Operations("Fn_Mech_ValidationErrorHandle", "Exist", Window("MechatronicsWindow").JavaDialog("NewParameterDefinition"), SISW_MICRO_TIMEOUT) Then
					Set objErrorDialog=Window("MechatronicsWindow").JavaDialog("NewParameterDefinition").JavaDialog("ValidationError")
				Else
					Set objErrorDialog=JavaDialog("NewParameterDefinition").JavaDialog("ValidationError")
				End If
				arrErrorMsg= Split(StrErrorMsg,"~")
				If  UBound(arrErrorMsg)=1 Then
					objErrorDialog.SetTOProperty "title",arrErrorMsg(1)
'					objErrorDialog.JavaCheckBox("More").Click 1,1
					objErrorDialog.JavaEdit("ErrMsg").SetTOProperty "index",1
				End If
				sDispMsg = objErrorDialog.JavaEdit("ErrMsg").GetROProperty("value")
				If btnName<>"" Then
					Call Fn_Button_Click("Fn_Mech_ValidationErrorHandle", objErrorDialog,btnName)	
				End If

				If  Instr(1,sDispMsg,arrErrorMsg(0))>0 Then
					Fn_Mech_ValidationErrorHandle=True
				Else
					GBL_ACTUAL_MESSAGE=sDispMsg
					Fn_Mech_ValidationErrorHandle=False
				End If
				objErrorDialog.JavaEdit("ErrMsg").SetTOProperty "index",0

			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			Case "ValidationConfirmation_Error"
				Set objErrorDialog=JavaDialog("RemovalConfirmation")
				objErrorDialog.SetTOProperty "Title","Validation Confirmation"
				If objErrorDialog.Exist(6) Then
                	objErrorDialog.JavaCheckBox("More").Set "ON"
					sDispMsg =objErrorDialog.JavaEdit("JTextArea").GetROProperty("value")
					If  Instr(1,sDispMsg,StrErrorMsg)>0 Then
						Fn_Mech_ValidationErrorHandle=True
					Else
						GBL_ACTUAL_MESSAGE=sDispMsg
						Fn_Mech_ValidationErrorHandle=False
					End If
					If btnName<>"" Then
						objErrorDialog.JavaButton("Yes").SetTOProperty "label",btnName
						wait 1
						Call Fn_Button_Click("Fn_Mech_ValidationErrorHandle", objErrorDialog,"Yes")	
					End If		
				End If
				wait 1
				objErrorDialog.SetTOProperty "Title","Removal Confirmation"
				wait 1
				objErrorDialog.JavaButton("Yes").SetTOProperty "label","Yes"
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			Case "ValidationConfirmation_Error_Ext"
				Dim objDesc, objChild, DeviceReplay
				Set objDesc=Description.Create()
				objDesc("Class Name").value="JavaCheckBox"
				objDesc("label").value="More..."
				Set  objChild = JavaDialog("label:=Validation Confirmation").ChildObjects(objDesc)
				xCord=objChild(0).getROProperty("abs_x")
				yCord=objChild(0).getROProperty("abs_y")
				Set DeviceReplay = CreateObject("Mercury.DeviceReplay")
				DeviceReplay.MouseClick xCord,yCord,LEFT_MOUSE_BUTTON
				Set objChild= Nothing
				Set objDesc = Nothing
				Set DeviceReplay = Nothing
				wait 1
				
				Set objDesc=Description.Create()
				objDesc("Class Name").value="JavaEdit"
				Set  objChild = JavaDialog("label:=Validation Confirmation").ChildObjects(objDesc)
				sDispMsg=objChild(0).getROProperty("value")
				Set objChild= Nothing
				Set objDesc = Nothing
				wait 1
				
				Set objDesc=Description.Create()
				objDesc("Class Name").value="JavaButton"
				objDesc("label").value=btnName
				Set  objChild = JavaDialog("label:=Validation Confirmation").ChildObjects(objDesc)
				xCord=objChild(0).getROProperty("abs_x")
				yCord=objChild(0).getROProperty("abs_y")
				Set DeviceReplay = CreateObject("Mercury.DeviceReplay")
				DeviceReplay.MouseClick xCord,yCord,LEFT_MOUSE_BUTTON
				Set objChild= Nothing
				Set objDesc = Nothing
				Set DeviceReplay = Nothing
				wait 1

				If  Instr(1,sDispMsg,StrErrorMsg)>0 Then
					Fn_Mech_ValidationErrorHandle=True
				Else
					GBL_ACTUAL_MESSAGE=sDispMsg
					Fn_Mech_ValidationErrorHandle=False
				End If
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			 Case "ValidationConfirmation_DoubleDialog"
					arrErrorMsg= Split(StrErrorMsg,"~")
					Set objErrorDialog=JavaDialog("RemovalConfirmation")
					objErrorDialog.SetTOProperty "title",arrErrorMsg(1)
					
					For iCount=arrErrorMsg(0) To 1 Step -1 
						objErrorDialog.SetTOProperty "index",iCount-1
						If objErrorDialog.Exist(6) Then
								objErrorDialog.JavaButton("Yes").SetTOProperty "label",btnName
								Call Fn_Button_Click("Fn_Mech_ValidationErrorHandle", objErrorDialog,"Yes")	
						End If
					Next
					Fn_Mech_ValidationErrorHandle=True
					objErrorDialog.SetTOProperty "index","0"

	End Select

	Set objErrorDialog=Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_Mech_ParameterDefinationGroupBasicCreate

'Description			 :	Function Used to create basic Parameter  Defenation group

'Parameters			   :   1.StrType: Parameter  Defenation group type
'										2.bConfigItem: Configuration Item
'										3.StrID: Parameter  Defenation group ID
'										4.StrRevision: Parameter  Defenation group Revision
'										5.StrName: Parameter  Defenation group Name
'										6.StrDescription: Parameter  Defenation group Description
'										7.StrGenCompID: Generic Component ID
'										8.StrRepresents: Parameter  Defenation group Represents
'										9.StrButtonName: Button Name
'
'Return Value		   : 	Item Id - revision or False

'Pre-requisite			:	Should be log in RAC

'Examples				:   bReturn=Fn_Mech_ParameterDefinationGroupBasicCreate("ParmGrpDef","off","","","Group1","","2486","Organizational Group","Finish")
'										bReturn=Fn_Mech_ParameterDefinationGroupBasicCreate("ParmGrpDef","","","","Group2","Parm Grp Desc","6","Represents Group","Next")
'
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												08-May-2012								1.0																						Sunny R
'											Snehal Salunkhe												10-Dec-2012								1.1							Koustubh W			Modified code to set values from Dropdown / TableCombo. for TC 10.1
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_Mech_ParameterDefinationGroupBasicCreate(StrType,bConfigItem,StrID,StrRevision,StrName,StrDescription,StrGenCompID,StrRepresents,StrButtonName)
	GBL_FAILED_FUNCTION_NAME="Fn_Mech_ParameterDefinationGroupBasicCreate"
 	'declaring variables
	Dim objParameterDefinitionGroupDialog,objStaticText,objChild
	Dim bFlag,crrID,crrRevision, sParameterDefinitionGroup
    StrType = Fn_SISW_MechCurrentobjName(StrType)
	Fn_Mech_ParameterDefinationGroupBasicCreate=false
	sParameterDefinitionGroup=Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("Mechatronics_Menu"), "NewParameterDefinitionGroup")
	'Checking existance of [ NewParameterDefinition ] dialog
	if not Window("MechatronicsWindow").JavaDialog("NewParameterDefinitionGroup").Exist(SISW_MIN_TIMEOUT) then
	   	If Environment.Value("ProductName") = sUFTProductName Then
	   		bFlag = Fn_MenuOperation("WinMenuSelect",sParameterDefinitionGroup)
			Call  Fn_ReadyStatusSync(2)
			If bFlag=false Then
				exit function
			End If
		Else	
			bFlag = Fn_MenuOperation("Select",sParameterDefinitionGroup)
			Call  Fn_ReadyStatusSync(2)
			If bFlag=false Then
				exit function
			End If
		End if
	End if
	
	'Creating object of [ NewParameterDefinitionGroup ] dialog
	set objParameterDefinitionGroupDialog=Window("MechatronicsWindow").JavaDialog("NewParameterDefinitionGroup")
	'selecting Parameter Defination Group Type
	Call Fn_List_Select("Fn_Mech_ParameterDefinationGroupBasicCreate", objParameterDefinitionGroupDialog,"ParameterDefinitionGroupList",StrType)
     'Wait till  Button is Enabled
	objParameterDefinitionGroupDialog.JavaButton("Next").WaitProperty "enabled", 1, 60000
	'Click on "Next" button
	objParameterDefinitionGroupDialog.JavaButton("Next").Click micLeftBtn
	wait 1
	'setting Parameter Definition Group ID
	If StrID<>"" Then
		Call Fn_Edit_Box("Fn_Mech_ParameterDefinationGroupBasicCreate",objParameterDefinitionGroupDialog,"ID",StrID)
		wait(1)
	End If
	'setting Parameter Definition group Revision
	If StrRevision<>"" Then
		Call Fn_Edit_Box("Fn_Mech_ParameterDefinationGroupBasicCreate",objParameterDefinitionGroupDialog,"Revision",StrRevision)
		wait(1)
	End If
	'clicking on assign button to assign ID and Revision
	If StrID="" or StrRevision="" Then
		Call Fn_Button_Click("Fn_Mech_ParameterDefinationGroupBasicCreate", objParameterDefinitionGroupDialog, "Assign")
		wait(2)
	End If
	'retriving ID and Revision
	crrID=Fn_Edit_Box_GetValue("Fn_Mech_ParameterDefinationGroupBasicCreate",objParameterDefinitionGroupDialog,"ID")
	crrRevision=Fn_Edit_Box_GetValue("Fn_Mech_ParameterDefinationGroupBasicCreate",objParameterDefinitionGroupDialog,"Revision")
	'setting Parameter Definition group Name
	If StrName<>"" Then
		Call Fn_Edit_Box("Fn_Mech_ParameterDefinationGroupBasicCreate",objParameterDefinitionGroupDialog,"Name","")
		wait(1)
		Call Fn_Edit_Box("Fn_Mech_ParameterDefinationGroupBasicCreate",objParameterDefinitionGroupDialog,"Name",StrName)
		wait(1)
	End If
	'setting Parameter Definition group Description
	If StrDescription<>"" Then
		Call Fn_Edit_Box("Fn_Mech_ParameterDefinationGroupBasicCreate",objParameterDefinitionGroupDialog,"Description","")
		wait(1)
		Call Fn_Edit_Box("Fn_Mech_ParameterDefinationGroupBasicCreate",objParameterDefinitionGroupDialog,"Description",StrDescription)
		wait(1)
	End If
	'setting Parameter Definition group Generic Component ID
	If StrGenCompID<>"" Then
		Call Fn_Edit_Box("Fn_Mech_ParameterDefinationGroupBasicCreate",objParameterDefinitionGroupDialog,"GenericComponentID",StrGenCompID)
		wait(1)
	End If
	Fn_Mech_ParameterDefinationGroupBasicCreate="'"&crrID+"-"+crrRevision
	'Clicking Next button
	Call Fn_Button_Click("Fn_Mech_ParameterDefinationGroupBasicCreate", objParameterDefinitionGroupDialog, "Next")
	Call  Fn_ReadyStatusSync(1)
	'Setting Parameter Definition group Represents
	If StrRepresents<>"" Then
'		objParameterDefinitionGroupDialog.JavaStaticText("ParaDefGroup_text").SetTOProperty "label","Represents:"
'		Call Fn_Button_Click("Fn_Mech_ParameterDefinationGroupBasicCreate", objParameterDefinitionGroupDialog, "ParaDefGroup_DropDown")
'		wait 2
'		Set objStaticText=Description.Create
'		objStaticText("Class Name").value="JavaStaticText"
'		objStaticText("label").value=StrRepresents
'		Set objChild=objParameterDefinitionGroupDialog.ChildObjects(objStaticText)
'		objChild(0).Click 1,1
'		wait 2
'		Set objStaticText=nothing
'		Set objChild=nothing

		set objParameterDefinitionGroupDialog=Window("MechatronicsWindow").JavaDialog("NewParameterDefinitionGroup")
		objParameterDefinitionGroupDialog.JavaStaticText("ParaDefGroup_text").SetTOProperty "label","Represents:"


		Set objStaticText=Description.Create
		objStaticText("Class Name").value="JavaTable"
		objStaticText("path").value=".*NewParmDefGrpDialog.*"
		objStaticText("path").RegularExpression = true
		Call Fn_Button_Click("Fn_Mech_ParameterDefinationGroupBasicCreate", objParameterDefinitionGroupDialog, "ParaDefGroup_DropDown")

		wait 2
		Set objChild=objParameterDefinitionGroupDialog.ChildObjects(objStaticText)
		Dim iRows
		iRows = cInt(objChild(0).getROProperty("rows"))
		For iCnt = 0 to iRows - 1
			If objChild(0).object.getValueAt(iCnt,0).getDisplayableValue() = StrRepresents Then
				objChild(0).ClickCell iCnt, 0
				Exit for
			End If
		Next
		Set objStaticText=nothing
		Set objChild=nothing
	End If
	wait 2
	If StrButtonName<>"" Then
        Call Fn_Button_Click("Fn_Mech_ParameterDefinationGroupBasicCreate", objParameterDefinitionGroupDialog, StrButtonName)
        Call  Fn_ReadyStatusSync(1)
        wait 1
		If lcase(StrButtonName)="finish" Then
			Call  Fn_ReadyStatusSync(1)
			Call Fn_Button_Click("Fn_Mech_ParameterDefinationGroupBasicCreate", objParameterDefinitionGroupDialog,"Close")
			Call  Fn_ReadyStatusSync(1)
		end if
	End If
	'releasing object of [ NewParameterDefinitionGroup ] dialog
	set objParameterDefinitionGroupDialog=nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_Mech_ParameterDefinationGroupRevisionInfo

'Description			 :	Function Used to fill Parameter Defenation group revision information

'Parameters			   :   1.dicParaGrpDefRevisionInfo: Parameter Defenation group revision information
'
'Return Value		   : 	true or False

'Pre-requisite			:	New Parameter Definition Group dialog shouls be appear and user should present on Parameter Defenation group revision information page

'Examples				:   dicParaGrpDefRevisionInfo("Comment")="Comment 1"
'										dicParaGrpDefRevisionInfo("ControlEngineer")="Analyst1"
'										dicParaGrpDefRevisionInfo("ParameterGroupDescriptor")="Descriptor 1"
'										dicParaGrpDefRevisionInfo("Specialist")="Analyst2"
'										dicParaGrpDefRevisionInfo("ButtonName")="Finish"
'										bReturn=Fn_Mech_ParameterDefinationGroupRevisionInfo(dicParaGrpDefRevisionInfo)
'
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												08-May-2012								1.0																						Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Pranav Ingle											13-Nov-2013								1.0																						Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_Mech_ParameterDefinationGroupRevisionInfo(dicParaGrpDefRevisionInfo)
	GBL_FAILED_FUNCTION_NAME="Fn_Mech_ParameterDefinationGroupRevisionInfo"
	'variable declaration
	Dim objParameterDefinitionGroupDialog,objStaticText,objChild
  	  Dim WshShell
	Dim iCounter,objTable

	'Checking existance of [ NewParameterDefinition ] dialog
	If Fn_SISW_UI_Object_Operations("Fn_Mech_ParameterDefinationGroupRevisionInfo","Exist",Window("MechatronicsWindow").JavaDialog("NewParameterDefinitionGroup"),SISW_MIN_TIMEOUT) = False Then
'	if not Window("MechatronicsWindow").JavaDialog("NewParameterDefinitionGroup").Exist(6) then
        exit function
	end if
	'Creating object of [ NewParameterDefinitionGroup ] dialog
	set objParameterDefinitionGroupDialog=Window("MechatronicsWindow").JavaDialog("NewParameterDefinitionGroup")
	'setting Parameter Definition group Revision comment
	If dicParaGrpDefRevisionInfo("Comment")<>"" Then
		Call Fn_Edit_Box("Fn_Mech_ParameterDefinationGroupRevisionInfo",objParameterDefinitionGroupDialog,"Comment",dicParaGrpDefRevisionInfo("Comment"))
		wait 1
	End If
	'selecting Control Engineer
	If dicParaGrpDefRevisionInfo("ControlEngineer")<>"" Then
		objParameterDefinitionGroupDialog.JavaStaticText("ParaDefGroup_text").SetTOProperty "label","Control Engineer:"
		Call Fn_Button_Click("Fn_Mech_ParameterDefinationGroupRevisionInfo", objParameterDefinitionGroupDialog, "ParaDefGroup_DropDown")
		wait 2
        For iCounter=0 to 2
			set WshShell = CreateObject("WScript.Shell")
			wait 1
			WshShell.SendKeys "{TAB}"
			WshShell.SendKeys "^{END}"
			set WshShell =nothing
		Next
        
		Set objTable=Description.Create()
		objTable("Class Name").value="JavaTable"

		objTable("toolkit class").value="com\.teamcenter\.rac\.common\.lov\.view\.components\.LOVTreeTable"
        objTable("displayed").value="1"
		objTable("enabled").value="1"
		objTable("focused").value="1"
		Set objChild=objParameterDefinitionGroupDialog.ChildObjects(objTable)


		For iCounter=0 to objChild(0).GetROProperty("rows")

			If trim(dicParaGrpDefRevisionInfo("ControlEngineer"))=trim(objChild(0).Object.getValueAt(iCounter,0).getDisplayableValue()) Then
				objChild(0).DoubleClickCell iCounter,0
				Exit for
			End If
		Next
		Set objTable=Nothing
		Set objChild=Nothing
		wait 2
    End If
	'setting Parameter Definition group Revision Parameter Group Descriptor
	If dicParaGrpDefRevisionInfo("ParameterGroupDescriptor")<>"" Then
	    objParameterDefinitionGroupDialog.JavaStaticText("ParaDefGroup_text").SetTOProperty "label","Parameter Group Descriptor:"
		Call Fn_Edit_Box("Fn_Mech_ParameterDefinationGroupRevisionInfo",objParameterDefinitionGroupDialog,"ParameterGroupDescriptor",dicParaGrpDefRevisionInfo("ParameterGroupDescriptor"))
		wait 1
	End If
	'selecting Specialist
	If dicParaGrpDefRevisionInfo("Specialist")<>"" Then
		objParameterDefinitionGroupDialog.JavaStaticText("ParaDefGroup_text").SetTOProperty "label","Specialist:"
		Call Fn_Button_Click("Fn_Mech_ParameterDefinationGroupRevisionInfo", objParameterDefinitionGroupDialog, "ParaDefGroup_DropDown")
		wait 2

		For iCounter=0 to 2
			set WshShell = CreateObject("WScript.Shell")
			wait 1
			WshShell.SendKeys "{TAB}"
			WshShell.SendKeys "^{END}"
			set WshShell =nothing
		Next
        
		Set objTable=Description.Create()
		objTable("Class Name").value="JavaTable"
		objTable("toolkit class").value="com.teamcenter.rac.common.lov.view.components.LOVTreeTable"
'		objTable("tagname").value="LOVTreeTable"
        objTable("displayed").value="1"
		objTable("enabled").value="1"
		objTable("focused").value="1"
		Set objChild=objParameterDefinitionGroupDialog.ChildObjects(objTable)


		For iCounter=0 to objChild(0).GetROProperty("rows")

			If trim(dicParaGrpDefRevisionInfo("Specialist"))=trim(objChild(0).Object.getValueAt(iCounter,0).getDisplayableValue()) Then
				objChild(0).DoubleClickCell iCounter,0
				Exit for
			End If
		Next
		Set objTable=Nothing
		Set objChild=Nothing
		wait 2
	End If
	'Clicking on button
	If dicParaGrpDefRevisionInfo("ButtonName")<>"" Then
		Call Fn_Button_Click("Fn_Mech_ParameterDefinationGroupRevisionInfo", objParameterDefinitionGroupDialog,dicParaGrpDefRevisionInfo("ButtonName"))
		Call  Fn_ReadyStatusSync(1)
		wait 1
		If lCase(dicParaGrpDefRevisionInfo("ButtonName"))="finish" Then
			Call  Fn_ReadyStatusSync(1)
			Call Fn_Button_Click("Fn_Mech_ParameterDefinationGroupRevisionInfo", objParameterDefinitionGroupDialog,"Close")
			Call  Fn_ReadyStatusSync(1)
		End If
	End If
	If Err.Number < 0 Then
		Fn_Mech_ParameterDefinationGroupRevisionInfo=False
	Else
		Fn_Mech_ParameterDefinationGroupRevisionInfo=True
	End If
	Set objParameterDefinitionGroupDialog=nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_SISW_Mech_ParameterValuesTableOperations

'Description			 :	Function Used to perform operation on Parameter Values table which is appear on New Parameter Values dialog

'Parameters			   :   1.StrAction: Action Name
'										2.StrName: Parameter Name or Row number
'										3.StrColName: Column name
'										4.StrValue: Cell value
'										5.StrButtonName: Button Name
'
'Return Value		   : 	true or false or "Not Editable"

'Pre-requisite			:	Parameter Values Table should exist

'Examples				:   bReturn=Fn_SISW_Mech_ParameterValuesTableOperations("DoubleClickCell","IntArr","Parameter Values","","")
'										bReturn=Fn_SISW_Mech_ParameterValuesTableOperations("DoubleClickCell","2","Parameter Values","","")
'										bReturn=Fn_SISW_Mech_ParameterValuesTableOperations("Verify","IntSingle","Type","ParmDefIntRevision","")
'										bReturn=Fn_SISW_Mech_ParameterValuesTableOperations("Verify","1","Type","ParmDefIntRevision","Finish")
'										bReturn=Fn_SISW_Mech_ParameterValuesTableOperations("isCellEditable","AutolampEnable_Cfg8604","Type","","")
'										Return Type of Case { isCellEditable } = True / False / "Not Editable"
'										bReturn=Fn_SISW_Mech_ParameterValuesTableOperations("VerifyCellBackgroundColour","AutolampEnable_Cfg8604","Minimum Values","LIGHTGREY","")
'
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done													Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												14-May-2012								1.0																																			Sunny R
'													Sandeep N												16-May-2012								1.1								Case : isCellEditable											Sunny R
'													Sandeep N												16-May-2012								1.2								Case : VerifyCellBackgroundColour				Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_SISW_Mech_ParameterValuesTableOperations(StrAction,StrName,StrColName,StrValue,StrButtonName)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_Mech_ParameterValuesTableOperations"
	'declaring variables
	Dim objParaTable
	Dim iRows,iCounter,bFlag,crrName,iCols,iColPos,sColour,sColourCode
	Fn_SISW_Mech_ParameterValuesTableOperations=false

	If StrColName="Type" AND StrValue<>"" Then
		StrValue=trim(replace(StrValue,"Revision",""))
		StrValue = Fn_SISW_MechCurrentobjName(StrValue)
		'StrValue=StrValue+" Revision"
	End If

    'checking existance of [ ParameterValues ] table
    If Fn_SISW_UI_Object_Operations("Fn_SISW_Mech_ParameterValuesTableOperations","Exist",Window("MechatronicsWindow").JavaDialog("NewParameterValue").JavaTable("ParameterValues"), SISW_MIN_TIMEOUT)= False Then 
'    If not Window("MechatronicsWindow").JavaDialog("NewParameterValue").JavaTable("ParameterValues").Exist(5) Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail: [ ParameterValues ] table not exist")
		Exit function
	else
		'creating object of [ ParameterValues ] table
		Set objParaTable=Window("MechatronicsWindow").JavaDialog("NewParameterValue").JavaTable("ParameterValues")
	End If
	
	Select Case StrAction
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to double click on specific cell
		Case "DoubleClickCell"
			'Checking "StrName" parameter is Numeric or not
			If not isNumeric(StrName) Then
				bFlag=false
				'Taking number of row from [ ParameterValues ] table
				iRows=Fn_UI_Object_GetROProperty("Fn_SISW_Mech_ParameterValuesTableOperations",objParaTable,"rows")
				For iCounter=0 to iRows-1
					'Taking Name row by row
					crrName=objParaTable.GetCellData(iCounter,"Name")
					If trim(crrName)=trim(StrName) Then
						bFlag=true
						Exit for
					End If
				Next
			else
				iCounter=cInt(StrName)-1
				bFlag=true
			End If
			If bFlag=true Then
				objParaTable.DoubleClickCell iCounter,StrColName,"LEFT","NONE"
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Successfully Double click on row [ "+CStr(iCounter)+ "] under column [ "+StrColName+" ]")
				Fn_SISW_Mech_ParameterValuesTableOperations=true			
			End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to Verify value
		Case "Verify"
			If not isNumeric(StrName) Then
				bFlag=false
				iRows=Fn_UI_Object_GetROProperty("Fn_SISW_Mech_ParameterValuesTableOperations",objParaTable,"rows")
				For iCounter=0 to iRows-1
					crrName=objParaTable.GetCellData(iCounter,"Name")
					If trim(crrName)=trim(StrName) Then
						If trim(objParaTable.GetCellData(iCounter,StrColName))=trim(StrValue) Then
							bFlag=true
							Exit for
						End If
					End If
				Next
			else
				iCounter=cInt(StrName)-1
				If trim(objParaTable.GetCellData(iCounter,StrColName))=trim(StrValue) Then
					bFlag=true
				End if
			End If
			If bFlag=true Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Successfully verify value [ "+StrValue+ "] appear under column [ "+StrColName+" ] on row [ "+CStr(iCounter)+" ]")
				Fn_SISW_Mech_ParameterValuesTableOperations=true
			else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail to verify value [ "+StrValue+ "] appear under column [ "+StrColName+" ] on row [ "+CStr(iCounter)+" ]")
			End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "isCellEditable"
			iCols=Fn_UI_Object_GetROProperty("Fn_SISW_Mech_ParameterValuesTableOperations",objParaTable,"cols")
			bFlag=false
			For iCounter=0 to iCols-1
				If trim(StrColName)=objParaTable.GetColumnName(iCounter) Then
					iColPos=iCounter
					bFlag=true
					Exit for
				End If
			Next
			If bFlag=false Then
				Exit function
			End If
			If not isNumeric(StrName) Then
				iRows=Fn_UI_Object_GetROProperty("Fn_SISW_Mech_ParameterValuesTableOperations",objParaTable,"rows")
				bFlag=false
				For iCounter=0 to iRows-1
					crrName=objParaTable.GetCellData(iCounter,"Name")
					If trim(crrName)=trim(StrName) Then
						bFlag=objParaTable.Object.isCellEditable(iCounter,iColPos)
						Exit for
					end if
				Next
			else
				iCounter=cInt(StrName)-1
				bFlag=objParaTable.Object.isCellEditable(iCounter,iColPos)
			End If
			If lcase(bFlag)="false" Then
					Fn_SISW_Mech_ParameterValuesTableOperations="Not Editable"
			else
					Fn_SISW_Mech_ParameterValuesTableOperations=true
			End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "VerifyCellBackgroundColour"
			iCols=Fn_UI_Object_GetROProperty("Fn_SISW_Mech_ParameterValuesTableOperations",objParaTable,"cols")
			bFlag=false
			For iCounter=0 to iCols-1
				If trim(StrColName)=objParaTable.GetColumnName(iCounter) Then
					iColPos=iCounter
					bFlag=true
					Exit for
				End If
			Next
			If bFlag=false Then
				Exit function
			End If
			If not isNumeric(StrName) Then
				iRows=Fn_UI_Object_GetROProperty("Fn_SISW_Mech_ParameterValuesTableOperations",objParaTable,"rows")
				bFlag=false
				For iCounter=0 to iRows-1
					crrName=objParaTable.GetCellData(iCounter,"Name")
					If trim(crrName)=trim(StrName) Then
						sColour=objParaTable.Object.getCellRenderer(iCounter,iColPos).getBackground().toString()
						Exit for
					end if
				Next
			else
				iCounter=cInt(StrName)-1
				sColour=objParaTable.Object.getCellRenderer(iCounter,iColPos).getBackground().toString()
			End If
			sColour =  mid(sColour ,instr(sColour ,"[")  ,instr(sColour ,"]") )
			Select Case StrValue
				Case "lightgrey","LIGHTGREY"
					sColourCode="[r=192,g=192,b=192]"
				Case "WHITE","white"
						sColourCode = "[r=255,g=255,b=255]"
			End Select
			If sColour=sColourCode Then
					Fn_SISW_Mech_ParameterValuesTableOperations=true
			else
					Fn_SISW_Mech_ParameterValuesTableOperations=false
			End If
	End Select
	'Clicking on Button as per user requirement
	If StrButtonName<>"" Then
		Call Fn_Button_Click("Fn_SISW_Mech_ParameterValuesTableOperations", Window("MechatronicsWindow").JavaDialog("NewParameterValue"),StrButtonName)
		If lcase(StrButtonName)="finish" Then
			Call Fn_ReadyStatusSync(1)
			Call Fn_Button_Click("Fn_SISW_Mech_ParameterValuesTableOperations",Window("MechatronicsWindow").JavaDialog("NewParameterValue"),"Close")
		End If
	End If
	'releasing object of [ ParameterValues ] table
	Set objParaTable=nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_SISW_Mech_EnterActualValueForParameterOperations

'Description			 :	Function Used to perform operation on Enter actual Value for Parameter Dialog

'Parameters			   :   1.StrAction: Action Name
'										2.bUseInitialValues: Use Initial Values option
'										3.bShowValueDescription: Show Value Description discription
'										4.iRow: row number
'										5.iCol: column numbers
'										6.StrValue: values
'										7.StrActualValueDialogButton: EnterActualValueForParameter Dialog Button name to click
'										8.StrButtonName: Parameter Values Dialog Button name to click
'
'Return Value		   : 	true or false

'Pre-requisite			:	Enter actual Value for Parameter Dialog should exist

'Examples				:   bReturn=Fn_SISW_Mech_EnterActualValueForParameterOperations("SetData","on","on","1~1~1","1~2~3","17:Value1~25:Value2~36:Value3","OK","Finish")
'										bReturn=Fn_SISW_Mech_EnterActualValueForParameterOperations("GetAllColumnNames","","","","","","","")
'										bReturn=Fn_SISW_Mech_EnterActualValueForParameterOperations("GetAllRowNames","","","","","","","")
'										bReturn=Fn_SISW_Mech_EnterActualValueForParameterOperations("SetSEDData","","","1~1","1~2","ele1~0xA2","OK","")
'										bReturn=Fn_SISW_Mech_EnterActualValueForParameterOperations("SetSEDData","","","","Domain Element Name~Value","ele1~0xA2","OK","")
'										bReturn=Fn_SISW_Mech_EnterActualValueForParameterOperations("SetSEDData","","","","Value","0xA2","OK","")
'										bReturn=Fn_SISW_Mech_EnterActualValueForParameterOperations("SetDateData","","","1","1","24-5-2012:Tomm","OK","")
'										bReturn=Fn_SISW_Mech_EnterActualValueForParameterOperations("SetBoolData","","","1","1","true:Value is true","","")
'										bReturn=Fn_SISW_Mech_EnterActualValueForParameterOperations("VerifyData","","","1~1~1~1~1","1~2~3~4~5","41~2D:Int2~36~43~0:Int5","","")
'										bReturn=Fn_SISW_Mech_EnterActualValueForParameterOperations("GetAllDomainElementName","","","","","","","")
'										bReturn=Fn_SISW_Mech_EnterActualValueForParameterOperations("GetAllDomainElementValue","","","","","","","")
'										bReturn=Fn_SISW_Mech_EnterActualValueForParameterOperations("VerifySEDData","","","","Domain Element Name~Value~Description","Hardware_levelB~0x10~Domain 2","","")
'										bReturn=Fn_SISW_Mech_EnterActualValueForParameterOperations("VerifySEDData","","","","1~2~3","Hardware_levelB~0x10~Domain 2","","")
'
'										bReturn=Fn_SISW_Mech_EnterActualValueForParameterOperations("PasteData","","off","1","1","","","")
'										bReturn=Fn_SISW_Mech_EnterActualValueForParameterOperations("PasteData","","off","1","5","","","")
'										bReturn=Fn_SISW_Mech_EnterActualValueForParameterOperations("PasteData","","on","1","","","","")
'										bReturn=Fn_SISW_Mech_EnterActualValueForParameterOperations("PasteData","","on","1","3","","","")
'
'										- - - - - - - to copy single cell
'										bReturn=Fn_SISW_Mech_EnterActualValueForParameterOperations("CopyData","","","1~1","1~1","","","")
'										
'										- - - - - - - to copy multiple cell range
'										bReturn=Fn_SISW_Mech_EnterActualValueForParameterOperations("CopyData","","","1~3","1~4","","","")
'										
'										- - - - - - - to copy Single row
'										bReturn=Fn_SISW_Mech_EnterActualValueForParameterOperations("CopyData","","","2","","","","")
'										
'										- - - - - - - to copy multiple rows
'										bReturn=Fn_SISW_Mech_EnterActualValueForParameterOperations("CopyData","","","2~3~4","","","","")
'
'										bReturn=Fn_SISW_Mech_EnterActualValueForParameterOperations("VerifyDateData","","","1~2~3","1~2~3","16-Mar-2012~16-Mar-2045~02-Oct-1954","","")
'
'										bReturn= Fn_SISW_Mech_EnterActualValueForParameterOperations("SetUseInitialValuesState","ON","","","","","","")
'										bReturn= Fn_SISW_Mech_EnterActualValueForParameterOperations("PasteRowColumnData","","","1~2","1~3","","","")
'										bReturn= Fn_SISW_Mech_EnterActualValueForParameterOperations("VerifyCellErrorMessage","","","1","1","Property = Parameter Values[1][1], Reason = Value is not within min-max limit [Maximum: 60, Minimum: 10]","","")
'										bReturn= Fn_SISW_Mech_EnterActualValueForParameterOperations("VerifyHeaderForegroundColour","","","","1","Red","","")'
'
'										If value is correct in cell then pass True and if value is incorrect then pass False
'										bReturn=Fn_SISW_Mech_EnterActualValueForParameterOperations("IsCellCurrentValueCorrect","","","1~2~2~4~3","1~1~2~2~4","True~True~True~True~False","","")
'
'										bReturn=Fn_SISW_Mech_EnterActualValueForParameterOperations("SelectCell","","","2","4","","","")
'										bReturn=Fn_SISW_Mech_EnterActualValueForParameterOperations("DoubleClickCell","","","4","4","","","")
'										bReturn=Fn_SISW_Mech_EnterActualValueForParameterOperations("Collapse_ON","","","","","","","")
'										bReturn=Fn_SISW_Mech_EnterActualValueForParameterOperations("Collapse_OFF","","","","","","","")
'
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												14-May-2012								1.0																						 Sunny R
'													Sandeep N												16-May-2012								1.1					Case : GetAllColumnNames			  Sunny R
'													Sandeep N												16-May-2012								1.2					Case : 	GetAllRowNames					 Sunny R
'													Sandeep N												23-May-2012								1.3					Case : 	SetSEDData					 		   Sunny R
'													Sandeep N												23-May-2012								1.4					Case : 	SetDateData					 			Sunny R
'													Sandeep N												30-May-2012								1.5					Case : 	SetBoolData					 			Sunny R
'													Sandeep N												01-Jun-2012								 1.6					Case : 	VerifyData					 			  Sunny R
'													Sandeep N												04-Jun-2012								 1.7					Case : 	GetAllDomainElementValue,GetAllDomainElementName					 			Sunny R
'													Sandeep N												05-Jun-2012								 1.8					Case : 	VerifySEDData					 	Sunny R
'													Sandeep N												03-Jul-2012								 1.9					Case : 	PasteData					 				Anjali M
'													Sandeep N												21-Aug-2012								 10.0					Case : 	CopyData					 			Pranav I
'													Anjali M												21-Jul-2012								 10.1					Case : 	VerifyDateData					 		Sandeep N
'													Sandeep N												23-Jul-2012								 10.2					Case : 	SetUseInitialValuesState,PasteRowColumnData,VerifyCellErrorMessage,VerifyHeaderForegroundColour					 		Sachin J
'													Sandeep N												23-Jul-2012								 10.3					Case : 	IsCellCurrentValueCorrect					 		Sachin J
'													Sandeep N												28-Aug-2012								 10.4					Case : 	SelectCell,DoubleClickCell					 		Sachin J
'													Sandeep N												29-Aug-2012								 10.5					Case : 	Collapse_OFF,Collapse_ON					 	Anjali M
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_SISW_Mech_EnterActualValueForParameterOperations(StrAction,bUseInitialValues,bShowValueDescription,iRow,iCol,StrValue,StrActualValueDialogButton,StrButtonName)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_Mech_EnterActualValueForParameterOperations"
 	'Declaring variables
	Dim objParaValueDialog,objTable,objChld
	Dim aCol,aRow,aValue,iCounter,aValDesc,StrLabel
	Dim colNumber,rowNumber,objDate,aDate,bFlag,iListCount,cellval,crrErrMsg
	Dim scrollMax
	Fn_SISW_Mech_EnterActualValueForParameterOperations=false
	'Checking existance of [ EnterActualValueForParameter ] dialog 
	If Fn_SISW_UI_Object_Operations("Fn_SISW_Mech_EnterActualValueForParameterOperations","Exist",JavaWindow("Mechatronics").JavaWindow("TcDefaultApplet").JavaDialog("EnterActualValueForParameter"), SISW_MIN_TIMEOUT)= False Then 
'	If not JavaWindow("Mechatronics").JavaWindow("TcDefaultApplet").JavaDialog("EnterActualValueForParameter").Exist(5) Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail: [ EnterActualValueForParameter ] dialog not exist")
		Exit function
	else
		'Creating object of [ EnterActualValueForParameter ] dialog
		Set objParaValueDialog=JavaWindow("Mechatronics").JavaWindow("TcDefaultApplet").JavaDialog("EnterActualValueForParameter")
	End If

	'Setting [ Use Initial Values ] option
	If bUseInitialValues<>"" Then
		Call Fn_CheckBox_Set("Fn_SISW_Mech_EnterActualValueForParameterOperations", objParaValueDialog, "UseInitialValues", bUseInitialValues)
	End If

	If objParaValueDialog.JavaCheckBox("Collapse").Exist(3) Then
		Call Fn_CheckBox_Set("Fn_SISW_Mech_EnterActualValueForParameterOperations", objParaValueDialog, "Collapse", "OFF")
	End If

	'Setting [ Show Value Description ] option
	If bShowValueDescription<>"" Then
		Call Fn_CheckBox_Set("Fn_SISW_Mech_EnterActualValueForParameterOperations", objParaValueDialog, "ValueDescriptionCell", bShowValueDescription)
	End If
	Select Case StrAction
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "SetData"
			aCol=Split(iCol,"~")
			aRow=Split(iRow,"~")
			aValue=Split(StrValue,"~")
			objParaValueDialog.JavaTable("Parameter").SelectRow 0
			For iCounter=0 to UBound(aValue)
				objParaValueDialog.JavaTable("Parameter").DeselectRow CInt(aRow(iCounter))-1
				objParaValueDialog.JavaTable("Parameter").ActivateCell CInt(aRow(iCounter))-1,CInt(aCol(iCounter))-1
				wait 1
				aValDesc=Split(aValue(iCounter),":")
				Set objTable=Description.Create
				objTable("Class Name").value="JavaTable"
				Set objChld=objParaValueDialog.JavaTable("Parameter").ChildObjects(objTable)
				If aValDesc(0)<>"" Then
					objChld(0).SetCellData 0,0,aValDesc(0)
				End If
				If uBound(aValDesc)=1 Then

					objChld(0).SetCellData 1,0,aValDesc(1)
				End If
			'	objChld(0).Object.setFocusable False
				Set objChld=nothing
				Set objTable=nothing
			next
			If Err.Number < 0 Then
				Fn_SISW_Mech_EnterActualValueForParameterOperations=False
			Else
				Fn_SISW_Mech_EnterActualValueForParameterOperations=True
			End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "GetAllRowNames"
			If objParaValueDialog.JavaCheckBox("Collapse").GetROProperty("value")=1 Then
				objParaValueDialog.JavaCheckBox("Collapse").Set "OFF"
				wait 1
			End If

			For iCounter=0 to objParaValueDialog.JavaTable("RowNames").GetROProperty("rows")-1
				If iCounter=0 Then
					StrLabel=objParaValueDialog.JavaTable("RowNames").GetCellData(0,0)
				else
					StrLabel=StrLabel+"~"+objParaValueDialog.JavaTable("RowNames").GetCellData(iCounter,0)
				End If
			Next
			If Err.Number < 0 Then
				Fn_SISW_Mech_EnterActualValueForParameterOperations=False
			Else
				Fn_SISW_Mech_EnterActualValueForParameterOperations=StrLabel
			End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "GetAllColumnNames","GetAllColumnNamesSED"

			If objParaValueDialog.JavaCheckBox("Collapse").Exist(2) Then
				If objParaValueDialog.JavaCheckBox("Collapse").GetROProperty("value")=1 Then
					objParaValueDialog.JavaCheckBox("Collapse").Set "OFF"
					wait 1
				End If
			End If
			
			For iCounter=0 to objParaValueDialog.JavaObject("JTableHeader").Object.getColumnModel().getColumnCount()-1
				If StrAction="GetAllColumnNames" Then
					If iCounter=0 Then
						StrLabel=objParaValueDialog.JavaObject("JTableHeader").Object.getColumnModel().getColumn(0).getHeaderRenderer().getColName()
					else
						StrLabel=StrLabel+"~"+objParaValueDialog.JavaObject("JTableHeader").Object.getColumnModel().getColumn(iCounter).getHeaderRenderer().getColName()
					End If
				Else
					If iCounter=0 Then
						StrLabel=objParaValueDialog.JavaObject("JTableHeader").Object.getColumnModel().getColumn(0).getHeaderValue().toString()
					else
						StrLabel=StrLabel+"~"+objParaValueDialog.JavaObject("JTableHeader").Object.getColumnModel().getColumn(iCounter).getHeaderValue().toString()
					End If
				End If
			Next
           
			If Err.Number < 0 Then
				Fn_SISW_Mech_EnterActualValueForParameterOperations=False
			Else
				Fn_SISW_Mech_EnterActualValueForParameterOperations=StrLabel
			End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "SetDateData"
			aCol=Split(iCol,"~")
			aRow=Split(iRow,"~")
			aValue=Split(StrValue,"~")
			objParaValueDialog.JavaTable("Parameter").SelectRow 0
			For iCounter=0 to UBound(aValue)
				objParaValueDialog.JavaTable("Parameter").ClickCell CInt(aRow(iCounter))-1,CInt(aCol(iCounter))-1,"LEFT"
				objParaValueDialog.JavaTable("Parameter").DoubleClickCell CInt(aRow(iCounter))-1,CInt(aCol(iCounter))-1,"LEFT"
				wait 1
				aValDesc=Split(aValue(iCounter),":")
				If JavaDialog("SelectDate").Exist(2) then
						aValDesc=Split(aValue(iCounter),":")
						aDate=Split(aValDesc(0),"-")

						Set objDate=JavaDialog("SelectDate").JavaObject("CalendarPanel").Object.getTime()

						objDate.setYear(Cint(aDate(2))-1900)
						objDate.setMonth(Cint(aDate(1))-1)
						objDate.setDate(aDate(0))
                        						
						JavaDialog("SelectDate").JavaObject("CalendarPanel").Object.setTime(objDate)
						JavaDialog("SelectDate").JavaButton("Ok").Click

						If ubound(aValDesc)=1 Then
							Set objTable=Description.Create
							objTable("Class Name").value="JavaTable"
							Set objChld=objParaValueDialog.JavaTable("Parameter").ChildObjects(objTable)
							objChld(0).SetCellData 1,0,aValDesc(1)
							objChld(0).Object.setFocusable False
							Set objChld=nothing
							Set objTable=nothing
						End If
					else
						set objParaValueDialog=nothing
						Exit function
					end if
				Next
				If Err.Number < 0 Then
					Fn_SISW_Mech_EnterActualValueForParameterOperations=False
				Else
					Fn_SISW_Mech_EnterActualValueForParameterOperations=True
				End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "SetSEDData"
			aCol=Split(iCol,"~")
			If aRow<>"" Then
				aRow=Split(iRow,"~")
			End If
			aValue=Split(StrValue,"~")
			For iCounter=0 to UBound(aValue)
				If not isNumeric(aCol(iCounter)) Then
					Select Case aCol(iCounter)
						Case "Domain Element Name"
							colNumber=0
						Case "Value"
							colNumber=1
					end Select
				else
					colNumber=cInt(aCol(iCounter))-1
				End If
				If aRow="" Then
					rowNumber=0
				else
					rowNumber=cInt(aRow(iCounter))-1
				End If
				objParaValueDialog.JavaTable("SEDTable").SetCellData rowNumber,colNumber,aValue(iCounter)
			next
			If Err.Number < 0 Then
				Fn_SISW_Mech_EnterActualValueForParameterOperations=False
			Else
				Fn_SISW_Mech_EnterActualValueForParameterOperations=True
			End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "SetBoolData"
			If objParaValueDialog.JavaCheckBox("ValueDescriptionCell").GetROProperty("enabled")=0 then
				objParaValueDialog.JavaCheckBox("Collapse").Set "OFF"
			end if
			Call Fn_CheckBox_Set("Fn_SISW_Mech_EnterActualValueForParameterOperations", objParaValueDialog, "ValueDescriptionCell","on")

			aCol=Split(iCol,"~")
			aRow=Split(iRow,"~")
			aValue=Split(StrValue,"~")
			objParaValueDialog.JavaTable("Parameter").SelectRow 0
			For iCounter=0 to UBound(aValue)
				objParaValueDialog.JavaTable("Parameter").ClickCell CInt(aRow(iCounter))-1,CInt(aCol(iCounter))-1,"LEFT"
				objParaValueDialog.JavaTable("Parameter").DoubleClickCell CInt(aRow(iCounter))-1,CInt(aCol(iCounter))-1,"LEFT"
				wait 1
				aValDesc=Split(aValue(iCounter),":")
				Set objTable=Description.Create
				objTable("Class Name").value="JavaTable"
				Set objChld=objParaValueDialog.JavaTable("Parameter").ChildObjects(objTable)
				If aValDesc(0)<>"" Then
						If trim(objChld(0).getCellData(0,0))<>aValDesc(0) Then
							objChld(0).SelectRow 0
							objChld(0).ClickCell 0,0
						End If
					End If
					If uBound(aValDesc)=1 Then
						objChld(0).SelectRow 1
						objChld(0).SetCellData 1,0,aValDesc(1)
					End If
	
				'	objChld(0).Object.setFocusable False
					Set objChld=nothing
					Set objTable=nothing
			Next
			If Err.Number < 0 Then
				Fn_SISW_Mech_EnterActualValueForParameterOperations=False
			Else
				Fn_SISW_Mech_EnterActualValueForParameterOperations=True
			End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "VerifyData"
			aCol=Split(iCol,"~")
			aRow=Split(iRow,"~")
			aValue=Split(StrValue,"~")
			For iCounter=0 to UBound(aValue)
				bFlag=false
				aValDesc=Split(aValue(iCounter),":")
				If aValDesc(0)=trim(objParaValueDialog.JavaTable("Parameter").Object.getCellData(CInt(aRow(iCounter))-1,CInt(aCol(iCounter))-1).get(3).get(1).toString()) Then
					bFlag=true
					If UBound(aValDesc)=1 Then
						If aValDesc(1)=trim(objParaValueDialog.JavaTable("Parameter").Object.getCellData(CInt(aRow(iCounter))-1,CInt(aCol(iCounter))-1).get(2).toString()) Then
							bFlag=true
						else
							bFlag=false
						End If
					End If
				End If
				If bFlag=false Then
					Exit for
				End If
			Next
			If bFlag=False Then
				Fn_SISW_Mech_EnterActualValueForParameterOperations=False
			Else
				Fn_SISW_Mech_EnterActualValueForParameterOperations=True
			End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "GetAllDomainElementName","GetAllDomainElementValue"
			Select Case StrAction
				Case "GetAllDomainElementName"
					objParaValueDialog.JavaTable("SEDTable").ClickCell 0,"Domain Element Name"
					wait 1
				Case "GetAllDomainElementValue"
					objParaValueDialog.JavaTable("SEDTable").ClickCell 0,"Value"
					wait 1
			End Select
			iListCount=objParaValueDialog.JavaList("SEDList").GetROProperty("items count")
			For iCounter=0 to iListCount-1
				If iCounter=0 Then
					aValue=objParaValueDialog.JavaList("SEDList").GetItem(iCounter)
				else
					aValue=aValue+"~"+objParaValueDialog.JavaList("SEDList").GetItem(iCounter)
				End If
			Next
			If Err.Number < 0 Then
				Fn_SISW_Mech_EnterActualValueForParameterOperations=False
			Else
				Fn_SISW_Mech_EnterActualValueForParameterOperations=aValue
			End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "VerifySEDData"
			aCol=Split(iCol,"~")
			aValue=Split(StrValue,"~")
			For iCounter=0 to UBound(aValue)
				bFlag=false
				If isNumeric(aCol(iCounter)) Then
					aCol(iCounter)=Int(aCol(iCounter))-1
				else
					Select Case aCol(iCounter)
						Case "Domain Element Name"
							aCol(iCounter)=0
						Case "Value"
							aCol(iCounter)=1
						Case "Description"
							aCol(iCounter)=2
					End Select
				End If

				If aValue(iCounter)=trim(objParaValueDialog.JavaTable("SEDTable").GetCellData(0,aCol(iCounter))) Then
					bFlag=true
				End If
				If bFlag=false Then
					Exit for
				End If
			Next
			If bFlag=False Then
				Fn_SISW_Mech_EnterActualValueForParameterOperations=False
			Else
				Fn_SISW_Mech_EnterActualValueForParameterOperations=True
			End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to Paste copied Data in Table
		Case "PasteData"
			If bShowValueDescription<>"" Then
				Call Fn_CheckBox_Set("Fn_SISW_Mech_EnterActualValueForParameterOperations", objParaValueDialog,"ValueDescriptionCell",bShowValueDescription)
			End If
			If iRow<>"" and iCol<>"" Then
				objParaValueDialog.JavaTable("Parameter").ClickCell Cint(iRow)-1,Cint(iCol)-1,"RIGHT"
			elseif iRow<>"" then
				aRow=Split(iRow,"~")
				objParaValueDialog.JavaTable("RowNames").SelectRow Cint(aRow(0))-1
				For iCounter=1 to ubound(aRow) 
					objParaValueDialog.JavaTable("RowNames").ExtendRow CInt(aRow(iCounter))-1
				Next
				objParaValueDialog.JavaTable("Parameter").ClickCell Cint(aRow(UBound(aRow)))-1,0,"RIGHT"
			else
				objParaValueDialog.JavaTable("Parameter").ClickCell 0,0,"RIGHT"
			End If
			objParaValueDialog.JavaMenu("index:=0","label:=Paste").Select
			If Err.Number < 0 Then
				Fn_SISW_Mech_EnterActualValueForParameterOperations=False
			Else
				Fn_SISW_Mech_EnterActualValueForParameterOperations=True
			End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to Copy Data from Table
		Case "CopyData"
			objParaValueDialog.JavaTable("Parameter").ClickCell 0,0,"LEFT"
			wait 1
			If bShowValueDescription<>"" Then
				Call Fn_CheckBox_Set("Fn_SISW_Mech_EnterActualValueForParameterOperations", objParaValueDialog,"ValueDescriptionCell",bShowValueDescription)
			End If
			If iRow<>"" and iCol<>"" Then
				aRow=Split(iRow,"~")
				aCol=Split(iCol,"~")
				objParaValueDialog.JavaTable("Parameter").SelectCellsRange cint(aRow(0))-1,cint(aCol(0))-1,cint(aRow(1))-1,Cint(aCol(1))-1
				objParaValueDialog.JavaTable("Parameter").ClickCell Cint(aRow(0))-1,Cint(aCol(0))-1,"RIGHT"
			elseif iRow<>"" then
				aRow=Split(iRow,"~")
				objParaValueDialog.JavaTable("RowNames").SelectRow Cint(aRow(0))-1
				For iCounter=1 to ubound(aRow) 
					objParaValueDialog.JavaTable("RowNames").ExtendRow CInt(aRow(iCounter))-1
				Next
				objParaValueDialog.JavaTable("Parameter").ClickCell Cint(aRow(UBound(aRow)))-1,0,"RIGHT"
			else
				objParaValueDialog.JavaTable("Parameter").PressKey "A",micCtrl
				wait 1
				objParaValueDialog.JavaTable("Parameter").ClickCell 0,0,"RIGHT"
			End If
			objParaValueDialog.JavaMenu("index:=0","label:=Copy").Select
			If Err.Number < 0 Then
				Fn_SISW_Mech_EnterActualValueForParameterOperations=False
			Else
				Fn_SISW_Mech_EnterActualValueForParameterOperations=True
			End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to Verify Date Data
		Case "VerifyDateData"
				'Spliting row numbers
				aCol=Split(iCol,"~")
				aRow=Split(iRow,"~")
				aValue=Split(StrValue,"~")
				For iCounter=0 to ubound(aValue)
					bFlag=False
					cellval=objParaValueDialog.JavaTable("Parameter").Object.getCellData(Cint(aRow(iCounter))-1,Cint(aCol(iCounter))-1).get(3).get(1).toString()
					aValDesc=Split(aValue(iCounter),":")
                    aDate=Split(cellval)

					If aDate(2)+"-"+aDate(1)+"-"+aDate(5)=aValDesc(0) or aDate(2)+"-0"+aDate(1)+"-"+aDate(5)=aValDesc(0) Then
						If ubound(aValDesc)=1 Then
							If trim(objParaValueDialog.JavaTable("Parameter").Object.getCellData(Cint(aRow(iCounter))-1,Cint(aCol(iCounter))-1).get(2).toString())=aValDesc(1) Then
								bFlag=True		
							End If
						Else
							bFlag=True
						End If
					else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Value [ " & aValue(iCounter) & " ] is not exist in [ Parameter ] table")
						Exit for
					End If
				Next
				If bFlag=True Then
					Fn_SISW_Mech_EnterActualValueForParameterOperations=true
				End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to Paste Data in Table with specific row column range
		Case "PasteRowColumnData"
			objParaValueDialog.JavaTable("Parameter").ClickCell 0,0,"LEFT"
			wait 1
			If bShowValueDescription<>"" Then
				Call Fn_CheckBox_Set("Fn_SISW_Mech_EnterActualValueForParameterOperations", objParaValueDialog,"ValueDescriptionCell",bShowValueDescription)
			End If
			If iRow<>"" and iCol<>"" Then
				aRow=Split(iRow,"~")
				aCol=Split(iCol,"~")
				objParaValueDialog.JavaTable("Parameter").SelectCellsRange cint(aRow(0))-1,cint(aCol(0))-1,cint(aRow(1))-1,Cint(aCol(1))-1
				objParaValueDialog.JavaTable("Parameter").ClickCell Cint(aRow(0))-1,Cint(aCol(0))-1,"RIGHT"
        	End If
			objParaValueDialog.JavaMenu("index:=0","label:=Paste").Select
			If Err.Number < 0 Then
				Fn_SISW_Mech_EnterActualValueForParameterOperations=False
			Else
				Fn_SISW_Mech_EnterActualValueForParameterOperations=True
			End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to Set state of [ Use Initial Values ] option ( check or uncheck )
		Case "SetUseInitialValuesState"
			Call Fn_CheckBox_Set("Fn_SISW_Mech_EnterActualValueForParameterOperations", objParaValueDialog, "UseInitialValues", bUseInitialValues)
			If Err.Number < 0 Then
				Fn_SISW_Mech_EnterActualValueForParameterOperations=False
			Else
				Fn_SISW_Mech_EnterActualValueForParameterOperations=True
			End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to verify cell error message in table
		Case "VerifyCellErrorMessage"
    			'Getting current error message from Specific cell
				crrErrMsg=objParaValueDialog.JavaTable("Parameter").Object.getValueAt(Cint(iRow)-1,Cint(iCol)-1).getErrMsg()
				'Comparing user pass error message with actual error message
				If StrValue<>"" Then
					If InStr(1,crrErrMsg,StrValue) Then
						Fn_SISW_Mech_EnterActualValueForParameterOperations=true
					else
						Fn_SISW_Mech_EnterActualValueForParameterOperations=false
					End If
				else
						Fn_SISW_Mech_EnterActualValueForParameterOperations=false
				End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			'Case to verify table foreground color
			Case "VerifyHeaderForegroundColour"
				If objParaValueDialog.JavaSlider("JScrollPane").Exist(1) Then
					scrollMax=objParaValueDialog.JavaSlider("JScrollPane").GetROProperty("max")
					objParaValueDialog.JavaSlider("JScrollPane").Drag scrollMax
					wait 1
				End If
                aCol=Split(iCol,"~")
				aValue=Split(StrValue,"~")
                For iCounter=0 to ubound(aCol)
						bFlag=false
						sColourCode=""
						sColour=objParaValueDialog.JavaObject("JTableHeader").Object.getColumnModel().getColumn(cint(aCol(iCounter))-1).getHeaderRenderer().getForeground().toString()
						sColour =  mid(sColour ,instr(sColour ,"[")  ,instr(sColour ,"]") )
						Select Case lCase(aValue(iCounter))
							Case "red"
								sColourCode="[r=255,g=0,b=0]"
						End Select
						If sColour=sColourCode Then
							bFlag=true
						else
							Exit for
						End if
				Next
                If bFlag=true Then
					Fn_SISW_Mech_EnterActualValueForParameterOperations=true
				else
					Fn_SISW_Mech_EnterActualValueForParameterOperations=false
				End If
         '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to verify current value of cell is correct or not
		Case "IsCellCurrentValueCorrect"
			aCol=Split(iCol,"~")
			aRow=Split(iRow,"~")
			aValue=Split(StrValue,"~")
			For iCounter=0 to ubound(aCol)
					bFlag=false
					bFlag=objParaValueDialog.JavaTable("Parameter").Object.getValueAt(cint(aRow(iCounter))-1,cint(aCol(iCounter))-1).isValueCorrect()
					If bFlag<>lcase(cstr(aValue(iCounter))) Then
						bFlag=false
						Exit for
					else
						bFlag=true
					End If
			Next
			 If bFlag=true Then
				Fn_SISW_Mech_EnterActualValueForParameterOperations=true
			else
				Fn_SISW_Mech_EnterActualValueForParameterOperations=false
			End If
	    ' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to select / Double click specific cell from tabel
		Case "SelectCell","DoubleClickCell"
				Select Case StrAction
					Case "SelectCell"
						objParaValueDialog.JavaTable("Parameter").ClickCell Cint(iCol)-1,Cint(iRow)-1
					Case "DoubleClickCell"
						objParaValueDialog.JavaTable("Parameter").DoubleClickCell Cint(iCol)-1,Cint(iRow)-1
				End Select
				wait 2
				If Err.Number < 0 Then
					Fn_SISW_Mech_EnterActualValueForParameterOperations=False
				Else
					Fn_SISW_Mech_EnterActualValueForParameterOperations=True
				End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to set state of [ Collapse ] option
		Case "Collapse_ON","Collapse_OFF"
			If StrAction="Collapse_ON" Then
				Fn_SISW_Mech_EnterActualValueForParameterOperations=Fn_CheckBox_Set("Fn_SISW_Mech_EnterActualValueForParameterOperations", objParaValueDialog, "Collapse", "ON")
			else
				Fn_SISW_Mech_EnterActualValueForParameterOperations=Fn_CheckBox_Set("Fn_SISW_Mech_EnterActualValueForParameterOperations", objParaValueDialog, "Collapse", "OFF")
			End If
	End Select
	'Clicking on [ EnterActualValueForParameter ] dialogs Button
	If StrActualValueDialogButton<>"" Then
		Call Fn_Button_Click("Fn_SISW_Mech_EnterActualValueForParameterOperations", objParaValueDialog,StrActualValueDialogButton)
	End If
	'Clicking on Button
	If StrButtonName<>"" Then
		Call Fn_ReadyStatusSync(1)
		Call Fn_Button_Click( "Fn_SISW_Mech_EnterActualValueForParameterOperations", Window("MechatronicsWindow").JavaDialog("NewParameterValue"),StrButtonName)
		If lCase(StrButtonName)="finish" Then
			Call Fn_ReadyStatusSync(1)
			Call Fn_Button_Click( "Fn_SISW_Mech_EnterActualValueForParameterOperations", Window("MechatronicsWindow").JavaDialog("NewParameterValue"),"Close")
		End If
	End If
	'Releasing object of [ EnterActualValueForParameter ] dialog
	Set objParaValueDialog=nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_SISW_Mech_ParameterValueProductSearchCriteria

'Description			 :	Function Used to perform search  Product on Parameter Value Dialog

'Parameters			   :   1.dicProductSearchDetails: Product search details
'
'Return Value		   : 	true or false

'Pre-requisite			:	Product Search Page should exist

'Examples				:   Dim dicProductSearchDetails
'										Set dicProductSearchDetails=CreateObject("Scripting.Dictionary")
'										With dicProductSearchDetails
'											.Add "Item ID",""
'										End With
'										dicProductSearchDetails("Item ID")="000379"
'										bReturn=Fn_SISW_Mech_ParameterValueProductSearchCriteria(dicProductSearchDetails)
'
'										With dicProductSearchDetails
'											.Add "Item ID",""
'											.Add "ButtonName",""
'										End With
'										dicProductSearchDetails("Item ID")="000379"
'										dicProductSearchDetails("ButtonName")="Next"
'										bReturn=Fn_SISW_Mech_ParameterValueProductSearchCriteria(dicProductSearchDetails)
'
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												14-May-2012								1.0																						Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_SISW_Mech_ParameterValueProductSearchCriteria(dicProductSearchDetails)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_Mech_ParameterValueProductSearchCriteria"
	Fn_SISW_Mech_ParameterValueProductSearchCriteria=False
 	'checking existance of [ NewParameterValue ] dialog
 	
 	If Fn_SISW_UI_Object_Operations("Fn_SISW_Mech_ParameterValueProductSearchCriteria","Exist",Window("MechatronicsWindow").JavaDialog("NewParameterValue"), SISW_MIN_TIMEOUT)= False Then 
'	If not Window("MechatronicsWindow").JavaDialog("NewParameterValue").Exist(5) Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail: [ New Parameter Value ] dialog not exist")
		Exit function
	End If
	'Variable Declaration
	Dim objParaValueDialog,bFlag,DictItems,DictKeys,iCounter,StrAction,i
	'Creating object of [ NewParameterValue ] dialog 
	Set objParaValueDialog=Window("MechatronicsWindow").JavaDialog("NewParameterValue")
	'Clearing All previous search criteria
    Call Fn_Button_Click( "Fn_SISW_Mech_ParameterValueProductSearchCriteria", objParaValueDialog,"Clear")
   
   For i = 1 To 21
    	Call Fn_KeyBoardOperation("SendKeys","{UP}")
   Next
    
	bFlag=false
	'taking the keys & items count from data dictionary
	DictItems = dicProductSearchDetails.Items
	DictKeys = dicProductSearchDetails.Keys
	
	For iCounter=0 to dicProductSearchDetails.count-1
        If DictItems(iCounter)<>"" Then
			StrAction=DictKeys(iCounter)
			Select Case StrAction
				'- - - - - - - - - - - - - -  Edit Box
				Case "Item ID","Name","Revision"
					'objParaValueDialog.JavaEdit("Search_Text").SetTOProperty "attached text",StrAction+":"
					objParaValueDialog.JavaStaticText("Name").SetTOProperty "label",StrAction+":"
					If objParaValueDialog.JavaEdit("Search_Text").Exist(3) Then
						Call Fn_Edit_Box("Fn_SISW_Mech_ParameterValueProductSearchCriteria", objParaValueDialog,"Search_Text",DictItems(iCounter))
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail: Search criteria [ "+StrAction+" ] not exist")
						Set objParaValueDialog=nothing
						Exit function
					End If
			End Select
		End if
	Next
	'Click on find button to Invoke search
	bFlag=Fn_Button_Click( "Fn_SISW_Mech_ParameterValueProductSearchCriteria", objParaValueDialog,"Find")
	If bFlag=true Then
		Call Fn_ReadyStatusSync(2)
		Fn_SISW_Mech_ParameterValueProductSearchCriteria=true
		'Clicking on Button
		If dicProductSearchDetails("ButtonName")<>"" Then
			Call Fn_Button_Click( "Fn_SISW_Mech_ParameterValueProductSearchCriteria", objParaValueDialog,dicProductSearchDetails("ButtonName"))
			If lCase(dicProductSearchDetails("ButtonName"))="finish" Then
				Call Fn_ReadyStatusSync(1)
				Call Fn_Button_Click( "Fn_SISW_Mech_ParameterValueProductSearchCriteria", objParaValueDialog,"Close")
			End If
		End If
	End If
	'Releasing object of [ NewParameterValue ] dialog
	Set objParaValueDialog=nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_SISW_Mech_ParameterValueChooseCategoryForProduct

'Description			 :	Function Used to Choose category for product from Parameter Value Dialog

'Parameters			   :   1.StrArchitecture: Architecture Name
'										2.StrRevisionRule: Revision Rule Name
'										3.StrElementID: Architecture Element ID
'										4.StrCategoryNodeName: Category Node Name
'										5.StrButtonName: Button name to click
'
'Return Value		   : 	true or false

'Pre-requisite			:	Choose category for product page should exist

'Examples				:   bReturn=Fn_SISW_Mech_ParameterValueChooseCategoryForProduct("A-Arch_Break","Latest Working","116655","A-Arch_Break (View):116655/A-ParaDefGrp (View)","Next")
'
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												14-May-2012								1.0																						Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_SISW_Mech_ParameterValueChooseCategoryForProduct(StrArchitecture,StrRevisionRule,StrElementID,StrCategoryNodeName,StrButtonName)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_Mech_ParameterValueChooseCategoryForProduct"
	Fn_SISW_Mech_ParameterValueChooseCategoryForProduct=False
	'checking existance of [ NewParameterValue ] dialog
	If Fn_SISW_UI_Object_Operations("Fn_SISW_Mech_ParameterValueChooseCategoryForProduct","Exist",Window("MechatronicsWindow").JavaDialog("NewParameterValue"), SISW_MIN_TIMEOUT)= False Then 
	'If not Window("MechatronicsWindow").JavaDialog("NewParameterValue").Exist(5) Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail: [ New Parameter Value ] dialog not exist")
		Exit function
	End If
	'variable declaration
	Dim objParaValueDialog
	'Creating object of [ NewParameterValue ] dialog 
	Set objParaValueDialog=Window("MechatronicsWindow").JavaDialog("NewParameterValue")
	'Selecting Architecture
	If StrArchitecture<>"" Then
		Call Fn_List_Select("Fn_SISW_Mech_ParameterValueChooseCategoryForProduct", objParaValueDialog, "Architecture",StrArchitecture)
	End If
	'Selecting Revision Rule
	If StrRevisionRule<>"" Then
		Call Fn_List_Select("Fn_SISW_Mech_ParameterValueChooseCategoryForProduct", objParaValueDialog, "RevisionRule",StrRevisionRule)
	End If
	'Setting Architecture Element ID
	If StrElementID<>"" Then
		Call Fn_Edit_Box("Fn_SISW_Mech_ParameterValueChooseCategoryForProduct", objParaValueDialog,"ArchitectureElementID",StrElementID)
	End If
	'Click on Find button
	Call Fn_CheckBox_Set("Fn_SISW_Mech_ParameterValueChooseCategoryForProduct", objParaValueDialog, "FindCategory","ON")
	'Selecting Category from products tree
	If StrCategoryNodeName<>"" Then
		wait 2
		objParaValueDialog.JavaTree("ProductTree").Activate StrCategoryNodeName
		Call Fn_ReadyStatusSync(2)
	End If
	'Clicking on Button
	If StrButtonName<>"" Then
		Call Fn_Button_Click( "Fn_SISW_Mech_ParameterValueChooseCategoryForProduct", objParaValueDialog,StrButtonName)
		If lCase(StrButtonName)="finish" Then
			Call Fn_ReadyStatusSync(1)
			Call Fn_Button_Click( "Fn_SISW_Mech_ParameterValueChooseCategoryForProduct", objParaValueDialog,"Close")
		End If
	End If
	If Err.Number < 0 Then
		Fn_SISW_Mech_ParameterValueChooseCategoryForProduct=False
	Else
		Fn_SISW_Mech_ParameterValueChooseCategoryForProduct=True
	End If
	'Releasing object of [ NewParameterValue ] dialog 
	Set objParaValueDialog=nothing			
End function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_SISW_Mech_ParameterValueBasicCreate

'Description			 :	Function Used to create basic Parameter Value

'Parameters			   :   1.StrParaValueType: Parameter Value Type
'										2.bConfigItem: Configuration Item Option
'										3.StrID: Parameter Value ID
'										4.StrRevision: Parameter Value Revision
'										5.StrName: Parameter Value Name
'										6.StrDescription: Parameter Value Description
'										7.StrButtonName: Button name to click
'
'Return Value		   : 	true or false

'Pre-requisite			:	should be log in RAC

'Examples				:   bReturn=Fn_SISW_Mech_ParameterValueBasicCreate("ParmGrpVal","off","","","Val1","Para Value1","Next")
'
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												14-May-2012								1.0																						Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Public Function Fn_SISW_Mech_ParameterValueBasicCreate(StrParaValueType,bConfigItem,StrID,StrRevision,StrName,StrDescription,StrButtonName)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_Mech_ParameterValueBasicCreate"
 	'Variable Declaration
	Dim objParaValueDialog
	Dim bFlag,crrID,crrRevision,hieght,width, sNewParameterDefinitionGroupMenu
    StrParaValueType = Fn_SISW_MechCurrentobjName(StrParaValueType)
	Fn_SISW_Mech_ParameterValueBasicCreate=False
	sNewParameterDefinitionGroupMenu= Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("Mechatronics_Menu"), "NewParameterValueGroup")
   	'Checking existance of 	[ New Parameter Value ] dialog
	If not Window("MechatronicsWindow").JavaDialog("NewParameterValue").Exist(SISW_MIN_TIMEOUT) Then
		'Calling Menu : File:New:Parameter Management:Parameter Value Group...
		bFlag = Fn_MenuOperation("Select",sNewParameterDefinitionGroupMenu)
		Call  Fn_ReadyStatusSync(1)
		If bFlag=false Then
			exit function
		End If
	End If
	
	'Creating object of [ New Parameter Value ] dialog
	Set objParaValueDialog=Window("MechatronicsWindow").JavaDialog("NewParameterValue")
	'selecting Parameter Defination Type
	Call Fn_List_Select("Fn_SISW_Mech_ParameterValueBasicCreate", objParaValueDialog,"ParameterValueList",StrParaValueType)
	'Selecting Configuration Item option
	If bConfigItem<>"" Then
		Call Fn_CheckBox_Set("Fn_SISW_Mech_ParameterValueBasicCreate", objParaValueDialog, "ConfigurationItem", bConfigItem)
	End If
	 ' Wait till  Button is Enabled
	objParaValueDialog.JavaButton("Next").WaitProperty "enabled", 1, 60000
	'Click on "Next" button
	objParaValueDialog.JavaButton("Next").Click micLeftBtn
	'setting Parameter Value ID
	If StrID<>"" Then
		Call Fn_Edit_Box("Fn_SISW_Mech_ParameterValueBasicCreate",objParaValueDialog,"ID",StrID)
	End If
	'setting Parameter Value Revision
	If StrRevision<>"" Then
		Call Fn_Edit_Box("Fn_SISW_Mech_ParameterValueBasicCreate",objParaValueDialog,"Revision",StrRevision)
	End If
	'clicking on assign button to assign ID and Revision
	If StrID="" or StrRevision="" Then
		Call Fn_Button_Click("Fn_SISW_Mech_ParameterValueBasicCreate", objParaValueDialog, "Assign")
	End If
	'retriving ID and Revision
	crrID=Fn_Edit_Box_GetValue("Fn_SISW_Mech_ParameterValueBasicCreate",objParaValueDialog,"ID")
	crrRevision=Fn_Edit_Box_GetValue("Fn_SISW_Mech_ParameterValueBasicCreate",objParaValueDialog,"Revision")
	'setting Parameter Value Name
	If StrName<>"" Then
		Call Fn_Edit_Box("Fn_SISW_Mech_ParameterValueBasicCreate",objParaValueDialog,"Name",StrName)
	End If
	'setting Parameter Value Description
	If StrDescription<>"" Then
		Call Fn_Edit_Box("Fn_SISW_Mech_ParameterValueBasicCreate",objParaValueDialog,"Description",StrDescription)
	End If
	Fn_SISW_Mech_ParameterValueBasicCreate="'"&crrID+"-"+crrRevision
	If StrButtonName<>"" Then
		If lcase(StrButtonName)="next" Then
			Call Fn_Button_Click("Fn_SISW_Mech_ParameterValueBasicCreate", objParaValueDialog, "Next")
'			wait 2
'			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'			'Resizing window
'			hieght=JavaWindow("Mechatronics").GetROProperty("height")
'			width=JavaWindow("Mechatronics").GetROProperty("width")
'			objParaValueDialog.Move 0,0
'			wait 2
'			objParaValueDialog.Resize width-5,hieght-5
'			wait 2
'			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		else
			Call Fn_Button_Click("Fn_SISW_Mech_ParameterValueBasicCreate", objParaValueDialog,StrButtonName)
			If lcase(StrButtonName)="finish" Then
				Call Fn_ReadyStatusSync(1)
				Call Fn_Button_Click("Fn_SISW_Mech_ParameterValueBasicCreate", objParaValueDialog,"Close")
			End If
		End If
	End If
	'releasing object of [ New Parameter Value ] dialog
	set objParaValueDialog=nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_SISW_Mech_ParameterValueAdditionalInformation

'Description			 :	Function Used to Define Parameter Value and Parameter Value group revision Additional Information

'Parameters			   :   1.StrUserData3: User Data 3 value
'										2.StrUserData2: User Data 2 value
'										3.StrUserData1: User Data 1 value
'										4.StrSerialNumber: Serial Number
'										5.StrPreviousID: Previous ID
'										6.StrItemComment: Item Comment
'										7.StrProjectID: Project ID
'										8.StrPreviousVersionID: Previous Version ID
'										9.StrButtonName: Button name to click
'
'Return Value		   : 	true or false

'Pre-requisite			:	Additional Information page should exist

'Examples				:   bReturn=Fn_SISW_Mech_ParameterValueAdditionalInformation("Data3","Data2","Data1","1111","ID1","Item 1","Project24","","Next")
'										bReturn=Fn_SISW_Mech_ParameterValueAdditionalInformation("RevData3","RevData2","RevData1","2211","","Item Rev 1","Project24","PrevID1","Next")
'
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												14-May-2012								1.0																						Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_SISW_Mech_ParameterValueAdditionalInformation(StrUserData3,StrUserData2,StrUserData1,StrSerialNumber,StrPreviousID,StrItemComment,StrProjectID,StrPreviousVersionID,StrButtonName)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_Mech_ParameterValueAdditionalInformation"
    'Variable Declaration 
	Dim objParaValueDialog

	Fn_SISW_Mech_ParameterValueAdditionalInformation=False
	'checking existance of [ NewParameterValue ] dialog
	If Fn_SISW_UI_Object_Operations("Fn_SISW_Mech_ParameterValueAdditionalInformation","Exist",Window("MechatronicsWindow").JavaDialog("NewParameterValue"), SISW_MIN_TIMEOUT)= False Then 
'	If not Window("MechatronicsWindow").JavaDialog("NewParameterValue").Exist(5) Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail: [ New Parameter Value ] dialog not exist")
		Exit function
	End If
	'Creating object of [ NewParameterValue ]
	Set objParaValueDialog=Window("MechatronicsWindow").JavaDialog("NewParameterValue")
	'Setting User Data 3
	If StrUserData3<>"" Then
		Call Fn_Edit_Box("Fn_SISW_Mech_ParameterValueAdditionalInformation",objParaValueDialog,"UserData3",StrUserData3)
	End If
	'Setting User Data 2
	If StrUserData2<>"" Then
		Call Fn_Edit_Box("Fn_SISW_Mech_ParameterValueAdditionalInformation",objParaValueDialog,"UserData2",StrUserData2)
	End If
	'Setting User Data 1
	If StrUserData1<>"" Then
		Call Fn_Edit_Box("Fn_SISW_Mech_ParameterValueAdditionalInformation",objParaValueDialog,"UserData1",StrUserData1)
	End If
	'Setting Serial Number
	If StrSerialNumber<>"" Then
		Call Fn_Edit_Box("Fn_SISW_Mech_ParameterValueAdditionalInformation",objParaValueDialog,"SerialNumber",StrSerialNumber)
	End If
	'Setting Previous ID
	If StrPreviousID<>"" Then
		Call Fn_Edit_Box("Fn_SISW_Mech_ParameterValueAdditionalInformation",objParaValueDialog,"PreviousID",StrPreviousID)
	End If
	'Setting Previous version ID
	If StrPreviousVersionID<>"" Then
		Call Fn_Edit_Box("Fn_SISW_Mech_ParameterValueAdditionalInformation",objParaValueDialog,"PreviousVersionID",StrPreviousVersionID)
	End If
	'Setting Item Comment
	If StrItemComment<>"" Then
		Call Fn_Edit_Box("Fn_SISW_Mech_ParameterValueAdditionalInformation",objParaValueDialog,"ItemComment",StrItemComment)
	End If
	'Setting Project ID
	If StrProjectID<>"" Then
		Call Fn_Edit_Box("Fn_SISW_Mech_ParameterValueAdditionalInformation",objParaValueDialog,"ProjectID",StrProjectID)
	End If
	'Clicking on button
	If StrButtonName<>"" Then
		Call Fn_Button_Click("Fn_SISW_Mech_ParameterValueAdditionalInformation", objParaValueDialog,StrButtonName)
		If lcase(StrButtonName)="finish" Then
			Call Fn_ReadyStatusSync(1)
			Call Fn_Button_Click("Fn_SISW_Mech_ParameterValueAdditionalInformation", objParaValueDialog,"Close")
		End If
	End If
	Fn_SISW_Mech_ParameterValueAdditionalInformation=True
	'releasing object of [ New Parameter Value ] dialog
	set objParaValueDialog=nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_SISW_Mech_EnterActualValueForBitParameterOperations

'Description			 :	Function Used to perform operation on Enter actual Value for Parameter Dialog for Bit Defination

'Parameters			   :   1.StrAction: Action Name
'										2.bShowDescriptor: Show Descriptor option
'										3.StrByte: Byte Number
'										4.StrBitName: Bit Name
'										5.StrValue: Values
'										6.StrDescriptor: Descriptor
'										7.StrActualValueDialogButton: EnterActualValueForParameter Dialog Button name to click
'										8.StrButtonName: Parameter Values Dialog Button name to click
'
'Return Value		   : 	true or false

'Pre-requisite			:	Enter actual Value for Parameter Dialog should exist

'Examples				:   bReturn=Fn_SISW_Mech_EnterActualValueForBitParameterOperations("VerifyData","on","1~1~1~1~1~1~1~1~2~2~2~2~2~2~2~2","7 - A~6 - B~5 - C~4 - D~3 - E~2 - F~1 - G~0 - I~7 - J~6 - K~5 - L~4 - M~3 - N~2 - O~1 - P~0 - Q","1~0~1~0~0~0~0~0~0~0~0~0~0~0~0~0","~~~~~~~Off~~~~~~~","","")
'										bReturn=Fn_SISW_Mech_EnterActualValueForBitParameterOperations("VerifyData","","2","5 - K","0","","","")
'										bReturn=Fn_SISW_Mech_EnterActualValueForBitParameterOperations("SetData","","2","5 - K","1","","","")
'										bReturn=Fn_SISW_Mech_EnterActualValueForBitParameterOperations("GetAllColumNames","","1~2","","","","","")
'										bReturn=Fn_SISW_Mech_EnterActualValueForBitParameterOperations("VerifyCellToolTipText","","1","6 - B","Descriptor for '0'.:,Descriptor for '1'.:","","","")
'
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												01-June-2012							1.0																						 Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_SISW_Mech_EnterActualValueForBitParameterOperations(StrAction,bShowDescriptor,StrByte,StrBitName,StrValue,StrDescriptor,StrActualValueDialogButton,StrButtonName)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_Mech_EnterActualValueForBitParameterOperations"
 	'Declaring Variables
	Dim aByte,aBitName,aValue,aDescription,iCounter,bFlag,aBit,iCol,iRow,crrValue
	Dim objParaValueDialog
	Fn_SISW_Mech_EnterActualValueForBitParameterOperations=false
	'Checking existance of [ EnterActualValueForParameter ] dialog 
	If not JavaWindow("Mechatronics").JavaWindow("TcDefaultApplet").JavaDialog("EnterActualValueForParameter").Exist(5) Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail: [ EnterActualValueForParameter ] dialog not exist")
		Exit function
	else
		'Creating object of [ EnterActualValueForParameter ] dialog
		Set objParaValueDialog=JavaWindow("Mechatronics").JavaWindow("TcDefaultApplet").JavaDialog("EnterActualValueForParameter")
	End If
	'Setting [ Show Descriptor ] option
	If bShowDescriptor<>"" Then
		Call Fn_CheckBox_Set("Fn_SISW_Mech_EnterActualValueForBitParameterOperations", objParaValueDialog, "ShowDescriptor", bShowDescriptor)
	End If

	Select Case StrAction
		'Case to Verify Cell Tool Tip text
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "VerifyCellToolTipText"

			aByte=Split(StrByte,"~")
			aBitName=Split(StrBitName,"~")
			aValue=Split(StrValue,"~")
			For iCounter=0 to ubound(aByte)
				bFlag=false
				aBit=SPlit(aBitName(iCounter),"-")
				iCol=8-CInt(trim(aBit(0)))
				'Verifying Bit value
    			iRow=CInt(aByte(iCounter))*4-3
				crrValue=objParaValueDialog.JavaTable("CCDMBitValueDataPanel").Object.getCellRenderer(iRow,iCol).getToolTipText()
				crrValue=Replace(crrValue,"<html>","")
				crrValue=Replace(crrValue,"</html>","")
				crrValue=Split(crrValue,"<br>")
				If trim(aValue(iCounter))=trim(crrValue(0))+","+trim(crrValue(1)) Then
					bFlag=true
				else
					Exit for
				End If
			Next
			If bFlag=true Then
				Fn_SISW_Mech_EnterActualValueForBitParameterOperations=true
			End If
        'Case to Set Data
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "SetData"
			aByte=Split(StrByte,"~")
			aBitName=Split(StrBitName,"~")
			aValue=Split(StrValue,"~")
			For iCounter=0 to ubound(aByte)
				bFlag=false
				aBit=SPlit(aBitName(iCounter),"-")
				iCol=8-CInt(trim(aBit(0)))
				'Verifying Bit value
				If StrValue<>"" Then
					If aValue(iCounter)<>"" Then
						iRow=CInt(aByte(iCounter))*4-3
						crrValue=objParaValueDialog.JavaTable("CCDMBitValueDataPanel").GetCellData(iRow,iCol)
						crrValue=objParaValueDialog.JavaTable("CCDMBitValueDataPanel").Object.getValueAt(iRow,iCol).toString()
						If aValue(iCounter)=1 Then
							aValue(iCounter)="true"
						else
							aValue(iCounter)="false"
						End If
						If trim(aValue(iCounter))<>trim(crrValue) Then
							objParaValueDialog.JavaTable("CCDMBitValueDataPanel").ClickCell iRow,iCol
							objParaValueDialog.JavaTable("CCDMBitValueDataPanel").DoubleClickCell iRow,iCol
						End If
					End If
				End If
			Next

			If Err.Number < 0 Then
				Fn_SISW_Mech_EnterActualValueForBitParameterOperations=False
			Else
				Fn_SISW_Mech_EnterActualValueForBitParameterOperations=True
			End If
		'Case to verify Bit Defination Data
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "VerifyData"
			aByte=Split(StrByte,"~")
			aBitName=Split(StrBitName,"~")
			If StrValue<>"" Then
				aValue=Split(StrValue,"~")
			End If
			If StrDescriptor<>"" Then
				aDescription=Split(StrDescriptor,"~")
			End If
			For iCounter=0 to ubound(aByte)
				bFlag=false
				
				aBit=SPlit(aBitName(iCounter),"-")
				iCol=8-CInt(trim(aBit(0)))
				'Verifying Bit value
				If StrValue<>"" Then
					If aValue(iCounter)<>"" Then
						iRow=CInt(aByte(iCounter))*4-3
'                        crrValue=objParaValueDialog.JavaTable("CCDMBitValueDataPanel").GetCellData(iRow,iCol)
						crrValue=objParaValueDialog.JavaTable("CCDMBitValueDataPanel").Object.getValueAt(iRow,iCol).toString()
						If aValue(iCounter)=1 Then
							aValue(iCounter)="true"
						else
							aValue(iCounter)="false"
						End If
						If trim(aValue(iCounter))=trim(crrValue) Then
							bFlag=true
						else
							Exit for
						End If
					End If
				End If
				'Verifying Bit Description
				If StrDescriptor<>"" Then
					If aDescription(iCounter)<>"" Then
						iRow=CInt(aByte(iCounter))*4-2
'						crrValue=objParaValueDialog.JavaTable("CCDMBitValueDataPanel").GetCellData(iRow,iCol)
						crrValue=objParaValueDialog.JavaTable("CCDMBitValueDataPanel").Object.getValueAt(iRow,iCol).toString()
						If trim(aDescription(iCounter))=trim(crrValue) Then
							bFlag=true
						else
							Exit for
						End If
					End If
				End If
			Next
			If bFlag=true Then
				Fn_SISW_Mech_EnterActualValueForBitParameterOperations=true
			End If
		'Case to get All Column Names Byte wise
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "GetAllColumNames"
			aByte=Split(StrByte,"~")
			For iCounter=0 to ubound(aByte)
				iRow=CInt(aByte(iCounter))*4-4
				For iCol=1 to 8
					If iCounter=0 and iCol=1 Then
'						crrValue=objParaValueDialog.JavaTable("CCDMBitValueDataPanel").GetCellData(iRow,iCol)
						crrValue=objParaValueDialog.JavaTable("CCDMBitValueDataPanel").Object.getValueAt(iRow,iCol).toString()
					else
'						crrValue=crrValue+"~"+objParaValueDialog.JavaTable("CCDMBitValueDataPanel").GetCellData(iRow,iCol)
						crrValue=crrValue+"~"+objParaValueDialog.JavaTable("CCDMBitValueDataPanel").Object.getValueAt(iRow,iCol).toString()
					End If
				Next
			Next
                If Err.Number < 0 Then
					Fn_SISW_Mech_EnterActualValueForBitParameterOperations=False
				Else
					Fn_SISW_Mech_EnterActualValueForBitParameterOperations=crrValue
				End If
	End Select
	'Clicking on [ EnterActualValueForParameter ] dialogs Button
	If StrActualValueDialogButton<>"" Then
		Call Fn_Button_Click("Fn_SISW_Mech_EnterActualValueForBitParameterOperations", objParaValueDialog,StrActualValueDialogButton)
	End If
	'Clicking on Button
	If StrButtonName<>"" Then
		Call Fn_ReadyStatusSync(1)
		Call Fn_Button_Click( "Fn_SISW_Mech_EnterActualValueForBitParameterOperations", Window("MechatronicsWindow").JavaDialog("NewParameterValue"),StrButtonName)
		If lCase(StrButtonName)="finish" Then
			Call Fn_ReadyStatusSync(1)
			Call Fn_Button_Click( "Fn_SISW_Mech_EnterActualValueForBitParameterOperations", Window("MechatronicsWindow").JavaDialog("NewParameterValue"),"Close")
		End If
	End If
	Set objParaValueDialog=nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_SISW_Mech_ViewerTabOperations

'Description			 :	Function Used to perform operation on Viewer Tab

'Parameters			   :   1.StrAction: Action Name
'										2.dicViewerTabInfo: Viewer Tab information
'
'Return Value		   : 	true or false or menu name

'Pre-requisite			:	Should be exist on Viewer Tab

'Examples				:   bReturn=Fn_SISW_Mech_ViewerTabOperations("MaximumValues_PasteData",dicViewerTabInfo)
'										dicViewerTabInfo("Row")="1"
'										dicViewerTabInfo("ShowValueDescription")="off"
'										dicViewerTabInfo("Collapse")="off"
'										bReturn=Fn_SISW_Mech_ViewerTabOperations("InitialValues_PasteData",dicViewerTabInfo)
'										
'										dicViewerTabInfo("Row")="1"
'										dicViewerTabInfo("Column")="2"
'										bReturn=Fn_SISW_Mech_ViewerTabOperations("MinimumValues_PasteData",dicViewerTabInfo)
'										
'										dicViewerTabInfo("Row")="1"
'										dicViewerTabInfo("Column")="1"
'										dicViewerTabInfo("CellErrorMessage")="Property = Initial Values[1][1], Reason = Value is not within min-max limit [Maximum: 100, Minimum: 40]"
'										bReturn=Fn_SISW_Mech_ViewerTabOperations("InitialValues_VerifyCellErrorMessage",dicViewerTabInfo)
'										
'										bReturn=Fn_SISW_Mech_ViewerTabOperations("InitialValues_GetAllContextMenu",dicViewerTabInfo)
'										
'										dicViewerTabInfo("Row")="1"
'										bReturn=Fn_SISW_Mech_ViewerTabOperations("InitialValues_GetAllContextMenu",dicViewerTabInfo)
'										
'										bReturn=Fn_SISW_Mech_ViewerTabOperations("MaximumValues_GetAllContextMenu",dicViewerTabInfo)
'										
'										dicViewerTabInfo("Row")="1"
'										bReturn=Fn_SISW_Mech_ViewerTabOperations("MaximumValues_GetAllContextMenu",dicViewerTabInfo)
'										
'										bReturn=Fn_SISW_Mech_ViewerTabOperations("MinimumValues_GetAllContextMenu",dicViewerTabInfo)
'										
'										dicViewerTabInfo("Row")="1"
'										dicViewerTabInfo("Column")="2"
'										bReturn=Fn_SISW_Mech_ViewerTabOperations("MinimumValues_GetAllContextMenu",dicViewerTabInfo)
'										
'										dicViewerTabInfo("Row")="1~1~2"
'										dicViewerTabInfo("Column")="1~2~1"
'										dicViewerTabInfo("Value")="40~30~35:Third"
'										bReturn=Fn_SISW_Mech_ViewerTabOperations("MinimumValues_VerifyData",dicViewerTabInfo)
'
'										dicViewerTabInfo("PropertyName")="Rows~Columns"
'										dicViewerTabInfo("Value")="4~5"
'										bReturn=Fn_SISW_Mech_ViewerTabOperations("EditBox",dicViewerTabInfo)
'
'										dicViewerTabInfo("Byte")="1"
'										bReturn=Fn_SISW_Mech_ViewerTabOperations("BitDefination_CopyData",dicViewerTabInfo)
'										
'										dicViewerTabInfo("BitNumber")="7"
'										dicViewerTabInfo("Byte")="2"
'										bReturn=Fn_SISW_Mech_ViewerTabOperations("BitDefination_PasteData",dicViewerTabInfo)
'
'										dicViewerTabInfo("Value")="Byte number:Bit number:Bit name:Bit 0 meaning:Bit 1 meaning~Byte number:Bit number:Bit name:Bit 0 meaning:Bit 1 meaning"
'										dicViewerTabInfo("Value")="1:7:Name1:0:1~2:0:Name88:0:1"
'										bReturn=Fn_SISW_Mech_ViewerTabOperations("BitDefination_VerifyData",dicViewerTabInfo)
'
'									dicViewerTabInfo("Row")="1"
'									dicViewerTabInfo("Column")="2"
'									dicViewerTabInfo("ShowValueDescription")="off"
'									dicViewerTabInfo("Collapse")="off"
'									bReturn=Fn_SISW_Mech_ViewerTabOperations("InitialValues_CopyData",dicViewerTabInfo)
'									
'									dicViewerTabInfo("Row")="1"
'									bReturn=Fn_SISW_Mech_ViewerTabOperations("SEDValidValues_CopyData",dicViewerTabInfo)
'									
'									dicViewerTabInfo("ShowValueDescription")="off"
'									dicViewerTabInfo("Collapse")="off"
'									bReturn=Fn_SISW_Mech_ViewerTabOperations("MaximumValues_CopyData",dicViewerTabInfo)
'									
'									bReturn=Fn_SISW_Mech_ViewerTabOperations("SEDValidValues_CopyData",dicViewerTabInfo)
'
'									dicViewerTabInfo("ShowValueDescription")="on"
'									dicViewerTabInfo("Row")="1~2"
'									dicViewerTabInfo("Column")="1~2"
'									bReturn=Fn_SISW_Mech_ViewerTabOperations("MaximumValues_PasteRowsCols",dicViewerTabInfo)
'
'									dicViewerTabInfo("Row")="1"
'									bReturn=Fn_SISW_Mech_ViewerTabOperations("MaximumValues_GetRowName",dicViewerTabInfo)
'									dicViewerTabInfo("Column")="1"
'									bReturn=Fn_SISW_Mech_ViewerTabOperations("MaximumValues_GetColumnName",dicViewerTabInfo)
'
'								dicViewerTabInfo("Row")="1"
'								dicViewerTabInfo("Column")="1"
'								dicViewerTabInfo("Value")="21-Mar-2012"
'								bReturn=Fn_SISW_Mech_ViewerTabOperations("MaximumValues_VerifyDateData",dicViewerTabInfo)
'
'											dicViewerTabInfo("Column")="1"
'											dicViewerTabInfo("Color")="Red"
'											bReturn=Fn_SISW_Mech_ViewerTabOperations("InitialValues_VerifyHeaderForegroundColour",dicViewerTabInfo)
'
'								dicViewerTabInfo("PageLink")="All"
'								dicViewerTabInfo("PropertyName")="Checked-Out"
'								dicViewerTabInfo("PropertyState")="value"
'								bReturn=Fn_SISW_Mech_ViewerTabOperations("EditBox_GetPropertyState",dicViewerTabInfo)
'
'									dicViewerTabInfo("Row")="1~2~3~4"
'									bReturn=Fn_SISW_Mech_ViewerTabOperations("SEDValidValues_PasteData",dicViewerTabInfo)
'									
'									dicViewerTabInfo("DomainElementName")="Element 1~Element 2~Element 3"
'									dicViewerTabInfo("DomainElementValue")="1A~7F~10B"
'									dicViewerTabInfo("DomainElementDescription")="A new Description for Element 1~A new Description for Element 2~A new Description for Element 3"
'									bReturn=Fn_SISW_Mech_ViewerTabOperations("SEDValidValues_VerifyData",dicViewerTabInfo)
'
'										dicViewerTabInfo("DomainElementName")="Element 1"
'										dicViewerTabInfo("DomainElementValue")="0x1A"
'										dicViewerTabInfo("DomainElementDescription")="A new Description for Element 1"
'										bReturn=Fn_SISW_Mech_ViewerTabOperations("SEDInitialValues_VerifyData",dicViewerTabInfo)
'										
'										dicViewerTabInfo("DomainElementName")="Element 1~Element 1"
'										dicViewerTabInfo("DomainElementValue")="0x1A~0x7F"
'										bReturn=Fn_SISW_Mech_ViewerTabOperations("SEDInitialValues_VerifyListData",dicViewerTabInfo)
'
'										dicViewerTabInfo("ParameterName")="ParmDefHex12345"
'										bReturn=Fn_SISW_Mech_ViewerTabOperations("ParameterValuesCellDoubleClick",dicViewerTabInfo)
'
'										dicViewerTabInfo("DomainElementName")="Element 3"
'										bReturn=Fn_SISW_Mech_ViewerTabOperations("SEDInitialValuesSetData",dicViewerTabInfo)
'
'										dicViewerTabInfo("ParameterName")="ParmDefBitDef12345"
'										dicViewerTabInfo("Column")="Parameter Values"
'										dicViewerTabInfo("Value")="{00000000}"
'										bReturn= Fn_SISW_Mech_ViewerTabOperations("ParameterValuesVerifyData",dicViewerTabInfo)
'										
'										dicViewerTabInfo("ParameterName")="3"
'										dicViewerTabInfo("Column")="Type"
'										dicViewerTabInfo("Value")="ParmDefHexRevision"
'										bReturn= Fn_SISW_Mech_ViewerTabOperations("ParameterValuesVerifyData",dicViewerTabInfo)
'
'										dicViewerTabInfo("Row")="1~2"
'										bReturn=Fn_SISW_Mech_ViewerTabOperations("SEDValidValues_CopyData",dicViewerTabInfo)
'										
'										dicViewerTabInfo("Row")="6"
'										bReturn=Fn_SISW_Mech_ViewerTabOperations("SEDValidValues_IsNameAndValueCorrect",dicViewerTabInfo)
'										
'										dicViewerTabInfo("Row")="1~2~3"
'										dicViewerTabInfo("Value")="red~red~red"
'										bReturn=Fn_SISW_Mech_ViewerTabOperations("SEDValidValues_VerifyRowForegroundColor",dicViewerTabInfo)
'
'										dicViewerTabInfo("Row")="1"
'										dicViewerTabInfo("Column")="2"
'										bReturn=Fn_SISW_Mech_ViewerTabOperations("InitialValues_SelectCell",dicViewerTabInfo)
'
'
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												07-June-2012							1.0																						 Sunny R
'													Sandeep N												08-June-2012							1.1					Added Case :EditBox							  Sunny R
'													Sandeep N												08-June-2012							1.2					Added Case :BitDefination_CopyData,BitDefination_PasteData							  Sunny R
'													Sandeep N												12-June-2012							1.3					Added Case :BitDefination_VerifyData					Sunny R
'													Sandeep N												18-June-2012							1.4					Added Case :SEDValidValues_CopyData,MaximumValues_CopyData					Sunny R
'													Sandeep N												02-Jully-2012							1.5				Added Case :MaximumValues_PasteRowsCols					Sunny R
'													Sandeep N												30-Jully-2012							1.6				Added Case :MaximumValues_GetRowName & MaximumValues_GetColumnName					Sunny R
'													Sandeep N												01-Aug-2012							1.7				Added Case :MaximumValues_VerifyDateData					Sunny R
'													Sandeep N												06-Aug-2012							1.8				Added Case :InitialValues_VerifyHeaderForegroundColour					Sonal P
'													Sandeep N												07-Aug-2012							1.9				Added Case :EditBox_GetPropertyState					Sonal P
'													Sandeep N												10-Aug-2012							10.0				Added Case :SEDValidValues_PasteData & SEDValidValues_VerifyData					Anjali Mane
'													Sandeep N												10-Aug-2012							10.1				Added Case :SEDInitialValues_VerifyData & SEDInitialValues_VerifyListData					Anjali Mane
'													Sandeep N												21-Aug-2012							10.2				Added Case :ParameterValuesCellDoubleClick			Pranav I
'													Sandeep N												21-Aug-2012							10.3				Added Case :SEDInitialValuesSetData					Anjali M
'													Sandeep N												23-Aug-2012							10.4				Added Case :ParameterValuesVerifyData					Sachin J
'													Sandeep N												27-Aug-2012							10.5				Added Case :SEDValidValues_IsNameAndValueCorrect,SEDValidValues_VerifyRowForegroundColor					Anjali M
'													Sonal P														28-Aug-2012							10.6				Added Case :"InitialValues_SelectCell","MaximumValues_SelectCell","MinimumValues_SelectCell"									Sandeep N
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_SISW_Mech_ViewerTabOperations(StrAction,dicViewerTabInfo)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_Mech_ViewerTabOperations"
 	'Declaring Variables
	Dim objApplet,aAction,tableName,crrErrMsg
	Dim objMenu,crrMenu,iCounter,objChld,aProperty
	Dim aCol,aRow,aValue,bFlag,aValDesc,scrollMax
	Dim aByte,aBitNumber,iRow,StrLabel,cellval,aDate,sColour
	Dim aDomainEleName,aDomainEleValue,aDomainEleDesc,iCount

	'Creating Object of Java Applet [ JApplet ]
	Set objApplet=Window("MechatronicsWindow").JavaWindow("WEmbeddedFrame")
	'Clicking on page link [ General or All ]
	If StrAction<>"MaximumValues_VerifyCellErrorMessage" And StrAction<>"MinimumValues_VerifyCellErrorMessage" And StrAction<>"InitialValues_VerifyCellErrorMessage" And StrAction<>"InitialValues_VerifyHeaderForegroundColour" And StrAction<>"MaximumValues_VerifyHeaderForegroundColour" And StrAction <> "MinimumValues_VerifyHeaderForegroundColour" Then
		If dicViewerTabInfo("PageLink")<>"" Then
			objApplet.JavaStaticText("PageLink").SetTOProperty "label",dicViewerTabInfo("PageLink")
			objApplet.JavaStaticText("PageLink").Click 1,1
		End If
	End If

	If objApplet.JavaSlider("JScrollPane").Exist(3) Then
		'Scrolling till the mid of panel
		scrollMax=objApplet.JavaSlider("JScrollPane").GetROProperty("max")
		objApplet.JavaSlider("JScrollPane").Drag scrollMax
	End If
	Select Case StrAction
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to verify Data from [ InitialValues , Maximum Values , MinimumValues ] table
		Case "InitialValues_VerifyData","MaximumValues_VerifyData","MinimumValues_VerifyData"
				aAction=Split(StrAction,"_")
				tableName=aAction(0)
				'Setting [ ShowValueDescription ] option
				If dicViewerTabInfo("ShowValueDescription")<>"" Then
					Call Fn_CheckBox_Set("Fn_SISW_Mech_ViewerTabOperations", objApplet, tableName&"ValueDescription",dicViewerTabInfo("ShowValueDescription"))
				End If
				'Setting [ Collapse ] option
				If dicViewerTabInfo("Collapse")<>"" Then
					Call Fn_CheckBox_Set("Fn_SISW_Mech_ViewerTabOperations", objApplet, tableName&"Collapse",dicViewerTabInfo("Collapse"))
				End If
				'Spliting row numbers
				aCol=Split(dicViewerTabInfo("Column"),"~")
				aRow=Split(dicViewerTabInfo("Row"),"~")
				aValue=Split(dicViewerTabInfo("Value"),"~")
				For iCounter=0 to ubound(aValue)
					bFlag=False
					aValDesc=Split(aValue(iCounter),":")
					'Verifying cell value
					If trim(objApplet.JavaTable(tableName).Object.getCellData(Cint(aRow(iCounter))-1,Cint(aCol(iCounter))-1).get(3).get(1).toString())=aValDesc(0) Then
						If ubound(aValDesc)=1 Then
							'Verifying cell Description
							If trim(objApplet.JavaTable(tableName).Object.getCellData(Cint(aRow(iCounter))-1,Cint(aCol(iCounter))-1).get(2).toString())=aValDesc(1) Then
								bFlag=True		
							End If
						Else
							bFlag=True
						End If
					else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Value [ " & aValue(iCounter) & " ] is not exist in [ "& tableName &" ] table")
						Exit for
					End If
				Next
				If bFlag=True Then
					Fn_SISW_Mech_ViewerTabOperations=true
				else
					Fn_SISW_Mech_ViewerTabOperations=false
				End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to Get all availabe context menu in or on [ InitialValues , Maximum Values , MinimumValues ] table
		Case "InitialValues_GetAllContextMenu","MaximumValues_GetAllContextMenu","MinimumValues_GetAllContextMenu"
				aAction=Split(StrAction,"_")
				tableName=aAction(0)
				'Setting [ ShowValueDescription ] option
				If dicViewerTabInfo("ShowValueDescription")<>"" Then
					Call Fn_CheckBox_Set("Fn_SISW_Mech_ViewerTabOperations", objApplet, tableName&"ValueDescription",dicViewerTabInfo("ShowValueDescription"))
				End If
				'Setting [ Collapse ] option
				If dicViewerTabInfo("Collapse")<>"" Then
					Call Fn_CheckBox_Set("Fn_SISW_Mech_ViewerTabOperations", objApplet, tableName&"Collapse",dicViewerTabInfo("Collapse"))
				End If
				If dicViewerTabInfo("Row")<>"" and dicViewerTabInfo("Column")<>"" Then
					objApplet.JavaTable(tableName).ClickCell Cint(dicViewerTabInfo("Row"))-1,Cint(dicViewerTabInfo("Column"))-1,"RIGHT"
				elseIf dicViewerTabInfo("Row")<>"" Then
					objApplet.JavaTable(tableName&"Rows").SelectRow Cint(dicViewerTabInfo("Row"))-1
					objApplet.JavaTable(tableName&"Rows").ClickCell Cint(dicViewerTabInfo("Row"))-1,0,"RIGHT"
				else
					objApplet.JavaTable(tableName).ClickCell 0,0,"RIGHT"
				End If
				Set objMenu=Description.Create
				objMenu("Class Name").value="JavaMenu"
				Set objChld=objApplet.ChildObjects(objMenu)
				crrMenu=""
				For iCounter=0 to objChld.count-1
						If iCounter=0 Then
							crrMenu=objChld(0).GetROProperty("label")
						else
							crrMenu=crrMenu+"~"+objChld(iCounter).GetROProperty("label")
						End If
				Next
				If Err.Number < 0 Then
					Fn_SISW_Mech_ViewerTabOperations=False
				Else
					If crrMenu<>"" Then
						Fn_SISW_Mech_ViewerTabOperations=crrMenu
					else
						Fn_SISW_Mech_ViewerTabOperations=False
					End If
				End If
				If dicViewerTabInfo("Row")<>"" Then
					objApplet.JavaTable(tableName&"Rows").ClickCell Cint(dicViewerTabInfo("Row"))-1,0
					objApplet.JavaTable(tableName).ClickCell 0,0
				else
					objApplet.JavaTable(tableName).ClickCell 0,0
				End If
				Set objChld=Nothing
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to verify cell error message in [ InitialValues , Maximum Values , MinimumValues ] table
		Case "InitialValues_VerifyCellErrorMessage","MaximumValues_VerifyCellErrorMessage","MinimumValues_VerifyCellErrorMessage"
				aAction=Split(StrAction,"_")
				tableName=aAction(0)
				'Setting [ ShowValueDescription ] option
				If dicViewerTabInfo("ShowValueDescription")<>"" Then
					Call Fn_CheckBox_Set("Fn_SISW_Mech_ViewerTabOperations", objApplet, tableName&"ValueDescription",dicViewerTabInfo("ShowValueDescription"))
				End If
				'Setting [ Collapse ] option
				If dicViewerTabInfo("Collapse")<>"" Then
					Call Fn_CheckBox_Set("Fn_SISW_Mech_ViewerTabOperations", objApplet, tableName&"Collapse",dicViewerTabInfo("Collapse"))
				End If
				'Getting current error message from Specific cell
				crrErrMsg=objApplet.JavaTable(tableName).Object.getValueAt(Cint(dicViewerTabInfo("Row"))-1,Cint(dicViewerTabInfo("Column"))-1).getErrMsg()
				'Comparing user pass error message with actual error message
'				If InStr(1,crrErrMsg,dicViewerTabInfo("CellErrorMessage")) Then
'					Fn_SISW_Mech_ViewerTabOperations=true
'				Else
					If CBool(InStr(1,crrErrMsg,dicViewerTabInfo("Row"))) AND CBool(InStr(1,crrErrMsg,dicViewerTabInfo("Column"))) OR CBool(InStr(1,crrErrMsg,dicViewerTabInfo("CellErrorMessage"))) Then
						Fn_SISW_Mech_ViewerTabOperations=true
					Else
						Fn_SISW_Mech_ViewerTabOperations=false
					End If
				'End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to paste data in [ InitialValues , Maximum Values , MinimumValues ] table
		Case "InitialValues_PasteData","MaximumValues_PasteData","MinimumValues_PasteData"
				aAction=Split(StrAction,"_")
				tableName=aAction(0)
				'Setting [ ShowValueDescription ] option
				If dicViewerTabInfo("ShowValueDescription")<>"" Then
					Call Fn_CheckBox_Set("Fn_SISW_Mech_ViewerTabOperations", objApplet, tableName&"ValueDescription",dicViewerTabInfo("ShowValueDescription"))
				End If
				'Setting [ Collapse ] option
				If dicViewerTabInfo("Collapse")<>"" Then
					Call Fn_CheckBox_Set("Fn_SISW_Mech_ViewerTabOperations", objApplet, tableName&"Collapse",dicViewerTabInfo("Collapse"))
				End If
				' - - - - - Selecting Specified Cell or Row- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				If dicViewerTabInfo("Row")<>"" and dicViewerTabInfo("Column")<>"" Then
					objApplet.JavaTable(tableName).ClickCell Cint(dicViewerTabInfo("Row"))-1,Cint(dicViewerTabInfo("Column"))-1,"RIGHT"
				elseif dicViewerTabInfo("Row")<>"" then
					objApplet.JavaTable(tableName&"Rows").SelectRow Cint(dicViewerTabInfo("Row"))-1
					objApplet.JavaTable(tableName&"Rows").ClickCell Cint(dicViewerTabInfo("Row"))-1,0,"RIGHT"
				else
					objApplet.JavaTable(tableName).ClickCell 0,0,"RIGHT"
				End If
				'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				'Selecting menu [ Paste ]  to paste copied data
				objApplet.JavaMenu("index:=0","label:=Paste").Select
				If Err.Number < 0 Then
					Fn_SISW_Mech_ViewerTabOperations=False
				Else
					Fn_SISW_Mech_ViewerTabOperations=True
				End If
				If dicViewerTabInfo("Row")<>"" and dicViewerTabInfo("Column")="" Then
					objApplet.JavaTable(tableName).DeselectRow Cint(dicViewerTabInfo("Row"))-1
				End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to modify property value from edit boxes
		Case "EditBox"
				aProperty=Split(dicViewerTabInfo("PropertyName"),"~")
				aValue=Split(dicViewerTabInfo("Value"),"~")
				For iCounter=0 to UBound(aProperty)
					bFlag=False
					objApplet.JavaStaticText("Viewer_Text").SetTOProperty "label",aProperty(iCounter)+":"
					If objApplet.JavaEdit("Viewer_Edit").Exist(3) Then
						
						If aProperty(iCounter)="Rows" or aProperty(iCounter)="Columns" or aProperty(iCounter)="Size in Byte(s)" Then
							objApplet.JavaEdit("Viewer_Edit").Set aValue(iCounter)+ vbLf + "" 'changed by shweta
							wait 1
						else
							 objApplet.JavaEdit("Viewer_Edit").Set aValue(iCounter)
						End If
					else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Property [ " & aProperty(iCounter) & " ] is not exist on dialog")
						Exit function
					End If
				Next
				If Err.Number < 0 Then
					Fn_SISW_Mech_ViewerTabOperations=false
				else
					Fn_SISW_Mech_ViewerTabOperations=true
				End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to paste data in [ Bit Defination ] table
		Case "BitDefination_CopyData","BitDefination_PasteData","BitDefination_CopyData_Keyboard","BitDefination_PasteData_Keyboard"
			If dicViewerTabInfo("Byte")<>"" and dicViewerTabInfo("BitNumber")<>"" Then
				'Spliting Byte information in array
				aByte=Split(dicViewerTabInfo("Byte"),"~")
				aBitNumber=Split(dicViewerTabInfo("BitNumber"),"~")
				iRow=Cint(aByte(0))*8-CInt(aBitNumber(0))-1
				objApplet.JavaTable("CCDMBitDefTable").SelectRow iRow
				For iCounter=1 to ubound(aByte)
					iRow=Cint(aByte(iCounter))*8-CInt(aBitNumber(iCounter))-1
					objApplet.JavaTable("CCDMBitDefTable").ExtendRow iRow
				Next
				objApplet.JavaTable("CCDMBitDefTable").ClickCell iRow,"Bit Number","RIGHT"
			elseif dicViewerTabInfo("Byte")<>"" then
				iRow=CInt(dicViewerTabInfo("Byte"))-1
				objApplet.JavaTable("CCDMBitDefRowHeaderTable").SelectRow iRow
				objApplet.JavaTable("CCDMBitDefRowHeaderTable").ClickCell iRow,"Byte","RIGHT"
			else
				objApplet.JavaTable("CCDMBitDefTable").SelectCell 0,"Bit Number"
                objApplet.JavaTable("CCDMBitDefRowHeaderTable").ClickCell 0,0,"RIGHT"
			End If		
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 					
			Select Case StrAction
				' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				Case "BitDefination_CopyData"
					objApplet.JavaMenu("index:=0","label:=Copy").Select
				' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				Case "BitDefination_PasteData"
					objApplet.JavaMenu("index:=0","label:=Paste").Select
				' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			End Select
			If Err.Number < 0 Then
				Fn_SISW_Mech_ViewerTabOperations=False
			Else
				Fn_SISW_Mech_ViewerTabOperations=True
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to verify values from Bit Defination table from Viewer Tab
		Case "BitDefination_VerifyData"
			'dicViewerTabInfo("Value")="Byte number:Bit number:Bit name:Bit 0 meaning:Bit 1 meaning~Byte number:Bit number:Bit name:Bit 0 meaning:Bit 1 meaning"
			aValue=Split(dicViewerTabInfo("Value"),"~")
			If objApplet.JavaCheckBox("BitDefinitionTableCollapse").Exist(2) Then
				If objApplet.JavaCheckBox("BitDefinitionTableCollapse").GetROProperty("enabled")=1 Then
					objApplet.JavaCheckBox("BitDefinitionTableCollapse").Object.doClick
				End If
			End If
			For iCounter=0 to ubound(aValue)
				bFlag=false
				aValDesc=Split(aValue(iCounter),":")

				iRow=Cint(aValDesc(0))*8-CInt(aValDesc(1))-1
				'verifing Bit name
                If aValDesc(2)<>"" Then
                    If trim(objApplet.JavaTable("CCDMBitDefTable").GetCellData(iRow,"Name"))=aValDesc(2) then
						bFlag=true
					end if
				else
					bFlag=true
				End If
				If bFlag=false Then
					Exit for
				End If
				bFlag=false
				'verifing Bit 0 meaning
                If aValDesc(3)<>"" Then
					If trim(objApplet.JavaTable("CCDMBitDefTable").GetCellData(iRow,"""0"" Meaning"))=aValDesc(3) then
						bFlag=true
					end if
				else
					bFlag=true
				End If
				If bFlag=false Then
					Exit for
				End If
				bFlag=false
				'verifing Bit 1 meaning
                If aValDesc(4)<>"" Then
					If trim(objApplet.JavaTable("CCDMBitDefTable").GetCellData(iRow,"""1"" Meaning"))=aValDesc(4) then
						bFlag=true
					end if
				else
					bFlag=true
				End If
				If bFlag=false Then
					Exit for
				End If
			Next
			If objApplet.JavaCheckBox("BitDefinitionTableExpand").Exist(2) Then
				If objApplet.JavaCheckBox("BitDefinitionTableExpand").GetROProperty("enabled")=1 Then
					objApplet.JavaCheckBox("BitDefinitionTableExpand").Object.doClick
				End If
			End If
			If bFlag=true Then
				Fn_SISW_Mech_ViewerTabOperations=true
			Else
				Fn_SISW_Mech_ViewerTabOperations=False
			End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to Copy data from [ InitialValues , Maximum Values , MinimumValues ] table
		Case "InitialValues_CopyData","MaximumValues_CopyData","MinimumValues_CopyData"
				aAction=Split(StrAction,"_")
				tableName=aAction(0)
				'Setting [ ShowValueDescription ] option
				If dicViewerTabInfo("ShowValueDescription")<>"" Then
					Call Fn_CheckBox_Set("Fn_SISW_Mech_ViewerTabOperations", objApplet, tableName&"ValueDescription",dicViewerTabInfo("ShowValueDescription"))
				End If
				'Setting [ Collapse ] option
				If dicViewerTabInfo("Collapse")<>"" Then
					Call Fn_CheckBox_Set("Fn_SISW_Mech_ViewerTabOperations", objApplet, tableName&"Collapse",dicViewerTabInfo("Collapse"))
				End If
				' - - - - - Selecting Specified Cell or Row- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				If dicViewerTabInfo("Row")<>"" and dicViewerTabInfo("Column")<>"" Then
					objApplet.JavaTable(tableName).ClickCell Cint(dicViewerTabInfo("Row"))-1,Cint(dicViewerTabInfo("Column"))-1,"RIGHT"
				elseif dicViewerTabInfo("Row")<>"" then
					objApplet.JavaTable(tableName&"Rows").SelectRow Cint(dicViewerTabInfo("Row"))-1
					objApplet.JavaTable(tableName&"Rows").ClickCell Cint(dicViewerTabInfo("Row"))-1,0,"RIGHT"
				else
					objApplet.JavaTable(tableName).PressKey "A",micCtrl
					wait 1
					objApplet.JavaTable(tableName).ClickCell 0,0,"RIGHT"
				End If
				'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				'Selecting menu [ Paste ]  to paste copied data
				objApplet.JavaMenu("index:=0","label:=Copy").Select
				If Err.Number < 0 Then
					Fn_SISW_Mech_ViewerTabOperations=False
				Else
					Fn_SISW_Mech_ViewerTabOperations=True
				End If
				If dicViewerTabInfo("Row")<>"" and dicViewerTabInfo("Column")="" Then
					objApplet.JavaTable(tableName).DeselectRow Cint(dicViewerTabInfo("Row"))-1
				End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to Copy data from [ ISED Valid Values ] table
		Case "SEDValidValues_CopyData"
				aAction=Split(StrAction,"_")
				tableName=aAction(0)
				' - - - - - Selecting Specified Cell or Row- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				If dicViewerTabInfo("Row")<>"" Then
					aRow=split(dicViewerTabInfo("Row"),"~")
					objApplet.JavaTable(tableName).SelectRow Cint(aRow(0))-1
					For iCounter=1 to ubound(aRow)
						objApplet.JavaTable(tableName).ExtendRow Cint(aRow(iCounter))-1
					Next
					objApplet.JavaTable(tableName).ClickCell Cint(aRow(ubound(aRow)))-1,0,"RIGHT"
				else
					objApplet.JavaTable(tableName).PressKey "A",micCtrl
					wait 1
					objApplet.JavaTable(tableName).ClickCell 0,0,"RIGHT"
				End If
				'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				'Selecting menu [ Paste ]  to paste copied data
				objApplet.JavaMenu("index:=0","label:=Copy").Select
				If Err.Number < 0 Then
					Fn_SISW_Mech_ViewerTabOperations=False
				Else
					Fn_SISW_Mech_ViewerTabOperations=True
				End If
				If dicViewerTabInfo("Row")<>"" Then
					objApplet.JavaTable(tableName).DeselectRow Cint(aRow(ubound(aRow)))-1
				End If
'		- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to paste data in [ InitialValues , Maximum Values , MinimumValues ] table
		Case "InitialValues_PasteRowsCols","MaximumValues_PasteRowsCols","MinimumValues_PasteRowsCols"
				aAction=Split(StrAction,"_")
				tableName=aAction(0)
				'Setting [ ShowValueDescription ] option
				If dicViewerTabInfo("ShowValueDescription")<>"" Then
					Call Fn_CheckBox_Set("Fn_SISW_Mech_ViewerTabOperations", objApplet, tableName&"ValueDescription",dicViewerTabInfo("ShowValueDescription"))
				End If
				'Setting [ Collapse ] option
				If dicViewerTabInfo("Collapse")<>"" Then
					Call Fn_CheckBox_Set("Fn_SISW_Mech_ViewerTabOperations", objApplet, tableName&"Collapse",dicViewerTabInfo("Collapse"))
				End If
				' - - - - - Selecting Specified Cell or Row- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				If dicViewerTabInfo("Row")<>"" and dicViewerTabInfo("Column")<>"" Then
					aRow=Split(dicViewerTabInfo("Row"),"~")
					aCol=Split(dicViewerTabInfo("Column"),"~")
					objApplet.JavaTable(tableName).SelectCellsRange cint(aRow(0))-1,cint(aCol(0))-1,cint(aRow(1))-1,Cint(aCol(1))-1
					objApplet.JavaTable(tableName).ClickCell Cint(aRow(0))-1,Cint(aCol(0))-1,"RIGHT"
				Elseif dicViewerTabInfo("Row")<>"" then
					aRow=Split(dicViewerTabInfo("Row"),"~")
					objApplet.JavaTable(tableName&"Rows").SelectRow Cint(aRow(0))-1
					For iCounter=1 to ubound(aRow)
						objApplet.JavaTable(tableName&"Rows").ExtendRow CInt(aRow(iCounter))-1
					Next
					objApplet.JavaTable(tableName).ClickCell Cint(aRow(UBound(aRow)))-1,0,"RIGHT"
				End if
				'Selecting menu [ Paste ]  to paste copied data
				objApplet.JavaMenu("index:=0","label:=Paste").Select
				If Err.Number < 0 Then
					Fn_SISW_Mech_ViewerTabOperations=False
				Else
					Fn_SISW_Mech_ViewerTabOperations=True
				End If
            		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to paste data in [ InitialValues , Maximum Values , MinimumValues ] table
		Case "InitialValues_CopyRowsCols","MaximumValues_CopyRowsCols","MinimumValues_CopyRowsCols"
				aAction=Split(StrAction,"_")
				tableName=aAction(0)
				'Setting [ ShowValueDescription ] option
				If dicViewerTabInfo("ShowValueDescription")<>"" Then
					Call Fn_CheckBox_Set("Fn_SISW_Mech_ViewerTabOperations", objApplet, tableName&"ValueDescription",dicViewerTabInfo("ShowValueDescription"))
				End If
				'Setting [ Collapse ] option
				If dicViewerTabInfo("Collapse")<>"" Then
					Call Fn_CheckBox_Set("Fn_SISW_Mech_ViewerTabOperations", objApplet, tableName&"Collapse",dicViewerTabInfo("Collapse"))
				End If
				' - - - - - Selecting Specified Cell or Row- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				If dicViewerTabInfo("Row")<>"" and dicViewerTabInfo("Column")<>"" Then
					aRow=Split(dicViewerTabInfo("Row"),"~")
					aCol=Split(dicViewerTabInfo("Column"),"~")
					objApplet.JavaTable(tableName).SelectCellsRange cint(aRow(0))-1,cint(aCol(0))-1,cint(aRow(1))-1,Cint(aCol(1))-1
					objApplet.JavaTable(tableName).ClickCell Cint(aRow(0))-1,Cint(aCol(0))-1,"RIGHT"
				Elseif dicViewerTabInfo("Row")<>"" then
					aRow=Split(dicViewerTabInfo("Row"),"~")
					objApplet.JavaTable(tableName&"Rows").SelectRow Cint(aRow(0))-1
					For iCounter=1 to ubound(aRow)
						objApplet.JavaTable(tableName&"Rows").ExtendRow CInt(aRow(iCounter))-1
					Next
					objApplet.JavaTable(tableName).ClickCell Cint(aRow(UBound(aRow)))-1,0,"RIGHT"
				End if
				'Selecting menu [ Paste ]  to paste copied data
				objApplet.JavaMenu("index:=0","label:=Copy").Select
				If Err.Number < 0 Then
					Fn_SISW_Mech_ViewerTabOperations=False
				Else
					Fn_SISW_Mech_ViewerTabOperations=True
				End If
		   '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to get specific Column name Exist in Table
		Case "MaximumValues_GetColumnName","InitialValues_GetColumnName","MinimumValues_GetColumnName"
				aAction=Split(StrAction,"_")
				tableName=aAction(0)

				If dicViewerTabInfo("Column")<>"" Then
					StrLabel=objApplet.JavaObject(tableName&"TableHeader").Object.getColumnModel().getColumn(CInt(dicViewerTabInfo("Column"))-1).getHeaderRenderer().getColName()
					Fn_SISW_Mech_ViewerTabOperations=StrLabel
				Else
					Fn_SISW_Mech_ViewerTabOperations=false
				end if
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to get specific row name Exist in Table
		Case "MaximumValues_GetRowName","InitialValues_GetRowName","MinimumValues_GetRowName"
				aAction=Split(StrAction,"_")
				tableName=aAction(0)

				If dicViewerTabInfo("Row")<>"" Then
					StrLabel=objApplet.JavaTable(tableName&"Rows").GetCellData(Cint(dicViewerTabInfo("Row"))-1,0)
					Fn_SISW_Mech_ViewerTabOperations=StrLabel
				Else
					Fn_SISW_Mech_ViewerTabOperations=false
				end if
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "InitialValues_VerifyDateData","MaximumValues_VerifyDateData","MinimumValues_VerifyDateData"
				aAction=Split(StrAction,"_")
				tableName=aAction(0)
				'Setting [ ShowValueDescription ] option
				If dicViewerTabInfo("ShowValueDescription")<>"" Then
					Call Fn_CheckBox_Set("Fn_SISW_Mech_ViewerTabOperations", objApplet, tableName&"ValueDescription",dicViewerTabInfo("ShowValueDescription"))
				End If
				'Setting [ Collapse ] option
				If dicViewerTabInfo("Collapse")<>"" Then
					Call Fn_CheckBox_Set("Fn_SISW_Mech_ViewerTabOperations", objApplet, tableName&"Collapse",dicViewerTabInfo("Collapse"))
				End If
				'Spliting row numbers
				aCol=Split(dicViewerTabInfo("Column"),"~")
				aRow=Split(dicViewerTabInfo("Row"),"~")
				aValue=Split(dicViewerTabInfo("Value"),"~")
				For iCounter=0 to ubound(aValue)
					bFlag=False
					cellval=objApplet.JavaTable(tableName).Object.getCellData(Cint(aRow(iCounter))-1,Cint(aCol(iCounter))-1).get(3).toString()
					If instr(1,cellval,"[") Then cellval=Replace(cellval,"[","") end if
					If instr(1,cellval,"]") Then cellval=Replace(cellval,"]","") end if
					
					aValDesc=Split(aValue(iCounter),":")
					aDate=Split(cellval)

					If aDate(3)+"-"+aDate(2)+"-"+aDate(6)=aValDesc(0) or aDate(3)+"-0"+aDate(2)+"-"+aDate(6)=aValDesc(0) Then
						If ubound(aValDesc)=1 Then
							If trim(objApplet.JavaTable(tableName).Object.getCellData(Cint(aRow(iCounter))-1,Cint(aCol(iCounter))-1).get(2).toString())=aValDesc(1) Then
								bFlag=True		
							End If
						Else
							bFlag=True
						End If
					else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Value [ " & aValues(iCounter) & " ] is not exist in [ " & tableName & " ] table")
						Exit for
					End If
				Next
				If bFlag=True Then
					Fn_SISW_Mech_ViewerTabOperations=true
				End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to verify Foreground color of Table header
		Case "InitialValues_VerifyHeaderForegroundColour","MaximumValues_VerifyHeaderForegroundColour","MinimumValues_VerifyHeaderForegroundColour"
			aAction=Split(StrAction,"_")
			tableName=aAction(0)
			aCol=Split(dicViewerTabInfo("Column"),"~")
			For iCounter=0 to ubound(aCol)
				bFlag=false
				sColourCode=""
				sColour=objApplet.JavaObject(tableName&"TableHeader").Object.getColumnModel().getColumn(cint(aCol(iCounter))-1).getHeaderRenderer().getForeground().toString()
				sColour =  mid(sColour ,instr(sColour ,"[")  ,instr(sColour ,"]") )
				Select Case lcase(dicViewerTabInfo("Color"))
					Case "red"
						sColourCode="[r=255,g=0,b=0]"
				End Select
				If sColour=sColourCode Then
					bFlag=true
				else
					Exit for
				End if

			Next
			If bFlag=true Then
					Fn_SISW_Mech_ViewerTabOperations=true
			else
					Fn_SISW_Mech_ViewerTabOperations=false
			End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to get specific property state of specific Edit box : e.g { current value, editable state, enabled state }
		Case "EditBox_GetPropertyState"
				objApplet.JavaStaticText("Viewer_Text").SetTOProperty "label",dicViewerTabInfo("PropertyName")+":"
				If objApplet.JavaEdit("Viewer_Edit").Exist(3) Then
					Fn_SISW_Mech_ViewerTabOperations=objApplet.JavaEdit("Viewer_Edit").GetROProperty(dicViewerTabInfo("PropertyState"))
				else
					Fn_SISW_Mech_ViewerTabOperations=false
				End if
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to verify data from [ ISED Valid Values ] table
        Case "SEDValidValues_VerifyData"
			aDomainEleName=Split(dicViewerTabInfo("DomainElementName"),"~")
			aDomainEleValue=Split(dicViewerTabInfo("DomainElementValue"),"~")
			aDomainEleDesc=Split(dicViewerTabInfo("DomainElementDescription"),"~")

			For iCount=0 to ubound(aDomainEleName)
				bFlag=False
				For iCounter=0 to cint(objApplet.JavaTable("SEDValidValues").GetROProperty("rows"))-1
					If aDomainEleName(iCount)=objApplet.JavaTable("SEDValidValues").Object.getCellData(iCounter,0).get(2).toString() Then
						bFlag=true
						If aDomainEleValue(iCount)<>"" Then
							If aDomainEleValue(iCount)<>objApplet.JavaTable("SEDValidValues").Object.getCellData(iCounter,0).get(3).get(1).toString() Then
								bFlag=false
								Exit for
							end if
						End If
						If aDomainEleDesc(iCount)<>"" Then
							If aDomainEleDesc(iCount)<>objApplet.JavaTable("SEDValidValues").Object.getCellData(iCounter,0).get(4).get(1).toString() Then
								bFlag=false
								Exit for
							end if
						End If
					End If
				Next
				If bFlag=false Then
					Exit for
				End If
			Next
			If bFlag=false Then
				Fn_SISW_Mech_ViewerTabOperations=false
			else
				Fn_SISW_Mech_ViewerTabOperations=true
			End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to Paste data from [ ISED Valid Values ] table
		Case "SEDValidValues_PasteData"
				aAction=Split(StrAction,"_")
				tableName=aAction(0)
				' - - - - - Selecting Specified Cell or Row- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				If dicViewerTabInfo("Row")<>"" Then
					aRow=Split(dicViewerTabInfo("Row"),"~")

					objApplet.JavaTable(tableName).SelectRow Cint(aRow(0))-1
					For iCounter=1 to ubound(aRow)
						objApplet.JavaTable("SEDValidValues").ExtendRow Cint(aRow(iCounter))-1
					Next
					objApplet.JavaTable(tableName).ClickCell Cint(aRow(ubound(aRow)))-1,0,"RIGHT"
				else
					objApplet.JavaTable(tableName).PressKey "A",micCtrl
					wait 1
					objApplet.JavaTable(tableName).ClickCell 0,0,"RIGHT"
				End If
				'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				'Selecting menu [ Paste ]  to paste copied data
				objApplet.JavaMenu("index:=0","label:=Paste").Select
				If Err.Number < 0 Then
					Fn_SISW_Mech_ViewerTabOperations=False
				Else
					Fn_SISW_Mech_ViewerTabOperations=True
				End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to verify data from [ SED Initial Values ] List
        Case "SEDInitialValues_VerifyListData"
			Fn_SISW_Mech_ViewerTabOperations=False

			If dicViewerTabInfo("DomainElementName")<>"" Then
					aDomainEleName=Split(dicViewerTabInfo("DomainElementName"),"~")
					objApplet.JavaTable("SEDInitialValue").ClickCell 0,"Domain Element Name"
					wait 1
					For iCounter=0 to ubound(aDomainEleName)
						bFlag=Fn_UI_ListItemExist("Fn_SISW_Mech_ViewerTabOperations", objApplet, "InitialValuesList",aDomainEleName(iCounter))
						If bFlag=false Then
							objApplet.JavaTable("SEDInitialValue").ClickCell 0,"Description"
							Exit function
							'Releasing Object of Java Applet [ JApplet ]
							Set objApplet=Nothing
						End If
					Next
			End If
			objApplet.JavaTable("SEDInitialValue").ClickCell 0,"Description"
			wait 1
			If dicViewerTabInfo("DomainElementValue")<>"" Then
					aDomainEleValue=Split(dicViewerTabInfo("DomainElementValue"),"~")
					objApplet.JavaTable("SEDInitialValue").ClickCell 0,"Value"
					wait 1
					For iCounter=0 to ubound(aDomainEleValue)
						bFlag=Fn_UI_ListItemExist("Fn_SISW_Mech_ViewerTabOperations", objApplet, "InitialValuesList",aDomainEleValue(iCounter))
						If bFlag=false Then
							objApplet.JavaTable("SEDInitialValue").ClickCell 0,"Description"
							Exit function
							'Releasing Object of Java Applet [ JApplet ]
							Set objApplet=Nothing
						End If
					Next
			End if
			Fn_SISW_Mech_ViewerTabOperations=True
			objApplet.JavaTable("SEDInitialValue").ClickCell 0,"Description"
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to verify data from [ SED Initial Values ] table
        Case "SEDInitialValues_VerifyData"
			bFlag=False
			For iCounter=0 to 0
				If dicViewerTabInfo("DomainElementName")=objApplet.JavaTable("SEDInitialValue").GetCellData(0,"Domain Element Name") Then
					bFlag=true
					If dicViewerTabInfo("DomainElementValue")<>"" Then
						If dicViewerTabInfo("DomainElementValue")<>objApplet.JavaTable("SEDInitialValue").GetCellData(0,"Value") Then
							bFlag=false
							Exit for
						End if
					End If
					If dicViewerTabInfo("DomainElementDescription")<>"" Then
						If dicViewerTabInfo("DomainElementDescription")<>objApplet.JavaTable("SEDInitialValue").GetCellData(0,"Description") Then
							bFlag=false
							Exit for
						end if
					End If
				End If
			Next
				
			If bFlag=false Then
				Fn_SISW_Mech_ViewerTabOperations=false
			else
				Fn_SISW_Mech_ViewerTabOperations=true
			End If
         ' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to Double click on Parameter Values Cell Double Click
		Case "ParameterValuesCellDoubleClick"
			If dicViewerTabInfo("Column")="" Then
				dicViewerTabInfo("Column")="Parameter Values"
			End If
'			iRow=Window("MechatronicsWindow").JavaWindow("WEmbeddedFrame").JavaTable("ParameterValues").GetROProperty("rows")
			iRow=Fn_UI_Object_GetROProperty("Fn_SISW_Mech_ViewerTabOperations",Window("MechatronicsWindow").JavaWindow("WEmbeddedFrame").JavaTable("ParameterValues"), "rows")
			bFlag=false
			For iCounter=0 to iRow-1
				If dicViewerTabInfo("ParameterName")=objApplet.JavaTable("ParameterValues").GetCellData(iCounter,"Name") Then
					Window("MechatronicsWindow").JavaWindow("WEmbeddedFrame").JavaTable("ParameterValues").DoubleClickCell iCounter,dicViewerTabInfo("Column")
					wait 2
					bFlag=true
					Exit for
				End If
			Next
			If bFlag=true Then
				Fn_SISW_Mech_ViewerTabOperations=true
			End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
        'Case to set Valid values
		Case "SEDInitialValuesSetData"
				'Selecting Domain Element Name
				If dicViewerTabInfo("DomainElementName")<>"" Then
					objApplet.JavaTable("SEDInitialValue").SetCellData 0,"Domain Element Name",dicViewerTabInfo("DomainElementName")
					wait 1
				End If
                'Selecting Value
				If dicViewerTabInfo("DomainElementValue")<>"" Then
					objApplet.JavaTable("SEDInitialValue").SetCellData 0,"Value",dicViewerTabInfo("DomainElementValue")
					wait 1
				End If
				If Err.Number < 0 Then
					Fn_SISW_Mech_ViewerTabOperations=False
				Else
					Fn_SISW_Mech_ViewerTabOperations=True
				End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to Verify value from Parameter Values table
		Case "ParameterValuesVerifyData"
			If not isNumeric(dicViewerTabInfo("ParameterName")) Then
				bFlag=false
				iRow=Fn_UI_Object_GetROProperty("Fn_SISW_Mech_ViewerTabOperations",objApplet.JavaTable("ParameterValues"),"rows")
				For iCounter=0 to iRow-1
					cellval=objApplet.JavaTable("ParameterValues").GetCellData(iCounter,"Name")
					If trim(cellval)=trim(dicViewerTabInfo("ParameterName")) Then
						If trim(objApplet.JavaTable("ParameterValues").GetCellData(iCounter,dicViewerTabInfo("Column")))=trim(dicViewerTabInfo("Value")) Then
							bFlag=true
							Exit for
						End If
					End If
				Next
			else
				bFlag=false
				iCounter=cInt(dicViewerTabInfo("ParameterName"))-1
				If trim(objApplet.JavaTable("ParameterValues").GetCellData(iCounter,dicViewerTabInfo("Column")))=trim(dicViewerTabInfo("Value")) Then
					bFlag=true
				End if
			End If
			If bFlag=true Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Successfully verify value [ "+dicViewerTabInfo("Value")+ "] appear under column [ "+dicViewerTabInfo("Column")+" ] on row [ "+CStr(iCounter)+" ]")
				Fn_SISW_Mech_ViewerTabOperations=true
			else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail to verify value [ "+dicViewerTabInfo("Value")+ "] appear under column [ "+dicViewerTabInfo("Column")+" ] on row [ "+CStr(iCounter)+" ]")
				Fn_SISW_Mech_ViewerTabOperations=false
			End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to check current value and name is correct or not
		Case "SEDValidValues_IsNameAndValueCorrect"
			If objApplet.JavaTable("SEDValidValues").Object.getValueAt(Cint(dicViewerTabInfo("Row"))-1,0).isNameCorrect()="true" and objApplet.JavaTable("SEDValidValues").Object.getValueAt(Cint(dicViewerTabInfo("Row"))-1,0).isValueCorrect()="true" then
				Fn_SISW_Mech_ViewerTabOperations=true
			else
				Fn_SISW_Mech_ViewerTabOperations=false
			end if
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to verify Foreground color of specific row text
		Case "SEDValidValues_VerifyRowForegroundColor"
			aRow=Split(dicViewerTabInfo("Row"),"~")
			aValue=Split(dicViewerTabInfo("Value"),"~")
			For iCounter=0 to ubound(aRow)
				bFlag=false
				Dim objccdmTable,objccdmTableRenderer,objinnerValue,objinnerTable,objinnerRenderer,objcellComp

				Set objccdmTable=objApplet.JavaTable("SEDValidValues").Object
				Set objccdmTableRenderer=objccdmTable.getCCDMTableRenderer()
				Set objinnerValue=objccdmTable.getValueAt(1,0)
				Set objinnerTable=objccdmTableRenderer.getTableCellRendererComponent(objccdmTable,objinnerValue,False,False,Cint(aRow(iCounter))-1,0)
				Set objinnerRenderer=objccdmTable.getInnerTableCellRenderer()
				Set objcellComp=objinnerRenderer.getTableCellRendererComponent(objinnerTable,objinnerValue,False,False,Cint(aRow(iCounter))-1,0)
				If objcellComp.getForeground() is nothing Then
					bFlag=false
				Else
					sColour=objcellComp.getForeground().toString()
					sColour =  mid(sColour ,instr(sColour ,"[")  ,instr(sColour ,"]") )
				End If

				wait 1
				sColourCode=""
				Select Case lCase(aValue(iCounter))
					Case "red"
						sColourCode="[r=255,g=0,b=0]"
				End Select
				If sColour=sColourCode Then
					bFlag=true
				else
					Exit for
				End if
			Next
			 If bFlag=true Then
				Fn_SISW_Mech_ViewerTabOperations=true
			else
				Fn_SISW_Mech_ViewerTabOperations=false
			End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to select specific cell from [ InitialValues , MaximumValues  , MinimumValues ] tables
		Case "InitialValues_SelectCell","MaximumValues_SelectCell","MinimumValues_SelectCell"
				aAction=Split(StrAction,"_")
				tableName=aAction(0)
				'Setting [ ShowValueDescription ] option
				If dicViewerTabInfo("ShowValueDescription")<>"" Then
					Call Fn_CheckBox_Set("Fn_SISW_Mech_ViewerTabOperations", objApplet, tableName&"ValueDescription",dicViewerTabInfo("ShowValueDescription"))
				End If
				'Setting [ Collapse ] option
				If dicViewerTabInfo("Collapse")<>"" Then
					Call Fn_CheckBox_Set("Fn_SISW_Mech_ViewerTabOperations", objApplet, tableName&"Collapse",dicViewerTabInfo("Collapse"))
				End If
				' - - - - - Selecting Specified Cell or Row- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				If dicViewerTabInfo("Row")<>"" and dicViewerTabInfo("Column")<>"" Then
					objApplet.JavaTable(tableName).ClickCell Cint(dicViewerTabInfo("Row"))-1,Cint(dicViewerTabInfo("Column"))-1
				End If
				'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				If Err.Number < 0 Then
					Fn_SISW_Mech_ViewerTabOperations=False
				Else
					Fn_SISW_Mech_ViewerTabOperations=True
				End If
	End Select
	If dicViewerTabInfo("ButtonName")<>"" Then
		Call Fn_Button_Click( "Fn_SISW_Mech_ViewerTabOperations",objApplet,dicViewerTabInfo("ButtonName"))
	End If
	'Releasing Object of Java Applet [ JApplet ]
	Set objApplet=Nothing
End Function

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_SISW_Mech_ViewerTabErrorHandle

'Description			 :	Function Used to handle validation errors  From  Viewer tab And View properties Dialog

'Parameters			   :   1.StrAction: Action Name
'										2. StrTitle : Title Name
'									 	3.StrErrorMsg : Error Message
'										4. intMsgIndex = 1 If want to verify msg by clicking More Button else 0 
'									 	3.btnName: Button Name
'
'Return Value		   : 	True or False

'Pre-requisite			:	Validation dialog should be open

'Examples				:   
'										1. bReturn=Fn_SISW_Mech_ViewerTabErrorHandle("Modification_Error", "Modification error","Not all the cell values are valid or entered for the below properties", "1", "OK")

'										2. bReturn=Fn_SISW_Mech_ViewerTabErrorHandle("Modification_Error", "Modification error","Mandatory values not satisfied during modification", "0", "OK")

'										3. bReturn=Fn_SISW_Mech_ViewerTabErrorHandle("ValidationConfirmation_DoubleDialog", "Validation Confirmation","3 error(s) found", "2", "OK")

'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done															Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Pranav Ingle												06-Aug-2012								1.0																																														Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_SISW_Mech_ViewerTabErrorHandle(StrAction,StrTitle, StrErrorMsg, intMsgIndex,btnName)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_Mech_ViewerTabErrorHandle"
   Dim sDispMsg,bReturn,objErrorDialog,iCount
   Dim objDesc, objChild, DeviceReplay

	Select Case StrAction
			Case "Modification_Error"
			''  commented below code by Shweta R. as per design changes on tc11.1(20140529) discussed with Archana
'				If  intMsgIndex = 1 Then
'					Set objDesc=Description.Create()
'					objDesc("Class Name").value="JavaCheckBox"
'					objDesc("label").value="More..."
'					Set  objChild = Window("MechatronicsWindow").JavaWindow("WEmbeddedFrame").JavaDialog("label:="&StrTitle).ChildObjects(objDesc)
'					If  objChild.Count = 0 Then
'						Set  objChild = JavaWindow("Mechatronics").JavaWindow("TcDefaultApplet").JavaDialog("label:="&StrTitle).ChildObjects(objDesc)
'					End If
'					xCord=objChild(0).getROProperty("abs_x")
'					yCord=objChild(0).getROProperty("abs_y")
'					Set DeviceReplay = CreateObject("Mercury.DeviceReplay")
'					DeviceReplay.MouseClick xCord,yCord,LEFT_MOUSE_BUTTON
'					Set objChild= Nothing
'					Set objDesc = Nothing
'					Set DeviceReplay = Nothing
'					wait 1
'				End If
				
				Set objDesc=Description.Create()
				objDesc("Class Name").value="JavaEdit"
				Set  objChild = Window("MechatronicsWindow").JavaWindow("WEmbeddedFrame").JavaDialog("label:="&StrTitle).ChildObjects(objDesc)
				If  objChild.Count = 0 Then
					Set  objChild = JavaWindow("Mechatronics").JavaWindow("TcDefaultApplet").JavaDialog("label:="&StrTitle).ChildObjects(objDesc)
				End If
				sDispMsg=objChild(intMsgIndex).getROProperty("value")
				Set objChild= Nothing
				Set objDesc = Nothing
				wait 1
				
				Set objDesc=Description.Create()
				objDesc("Class Name").value="JavaButton"
				objDesc("label").value=btnName
				Set  objChild = Window("MechatronicsWindow").JavaWindow("WEmbeddedFrame").JavaDialog("label:="&StrTitle).ChildObjects(objDesc)
				If  objChild.Count = 0 Then
					Set  objChild = JavaWindow("Mechatronics").JavaWindow("TcDefaultApplet").JavaDialog("label:="&StrTitle).ChildObjects(objDesc)
				End If
				xCord=objChild(0).getROProperty("abs_x")
				yCord=objChild(0).getROProperty("abs_y")
				Set DeviceReplay = CreateObject("Mercury.DeviceReplay")
				DeviceReplay.MouseClick xCord,yCord,LEFT_MOUSE_BUTTON
				Set objChild= Nothing
				Set objDesc = Nothing
				Set DeviceReplay = Nothing
				wait 1

				If  Instr(1,sDispMsg,StrErrorMsg)>0 Then
					Fn_SISW_Mech_ViewerTabErrorHandle=True
				Else
					Fn_SISW_Mech_ViewerTabErrorHandle=False
				End If
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			 Case "ValidationConfirmation_DoubleDialog"
					Set objErrorDialog=JavaDialog("RemovalConfirmation")
					objErrorDialog.SetTOProperty "title",StrTitle
					
					For iCount=CInt(intMsgIndex) To 1 Step -1
						objErrorDialog.SetTOProperty "index",iCount-1

						' Get Error message
						If  iCount=Cint(intMsgIndex)  And StrErrorMsg <> "" Then
							Set objDesc=Description.Create()
							objDesc("Class Name").value="JavaEdit"
							Set  objChild = objErrorDialog.ChildObjects(objDesc)
							If  objChild.Count = 0 Then
								Set  objChild = Window("MechatronicsWindow").JavaWindow("WEmbeddedFrame").JavaDialog("label:="&StrTitle).ChildObjects(objDesc)
							End If

							If  objChild.Count <> 0 Then
								sDispMsg=objChild(0).getROProperty("value")
							Else
								Set objDesc=Description.Create()
								objDesc("tagname").value="MLabel"
								objDesc("Class Name").value="JavaObject"
								Set  objChild = objErrorDialog.ChildObjects(objDesc)
								If  objChild.Count = 0 Then
									Set  objChild = Window("MechatronicsWindow").JavaWindow("WEmbeddedFrame").JavaDialog("label:="&StrTitle).ChildObjects(objDesc)
								End If
								For iCounter = 0 To objChild.Count-1
									If objChild(iCounter).getROProperty("tagname") =  "MLabel" Then
										sDispMsg=objChild(iCounter).Object.getText
										Exit For
									End If
								Next
							End If
							
							Set objChild= Nothing
							Set objDesc = Nothing
							wait 1
						End If

						' Close Error Dialog
						Set objDesc=Description.Create()
						objDesc("Class Name").value="JavaButton"
						objDesc("label").value=btnName
						Set  objChild = objErrorDialog.ChildObjects(objDesc)
						If  objChild.Count = 0 Then
							Set  objChild = Window("MechatronicsWindow").JavaWindow("WEmbeddedFrame").JavaDialog("label:="&StrTitle).ChildObjects(objDesc)
						End If
						xCord=objChild(0).getROProperty("abs_x")
						yCord=objChild(0).getROProperty("abs_y")
						Set DeviceReplay = CreateObject("Mercury.DeviceReplay")
						DeviceReplay.MouseClick xCord,yCord,LEFT_MOUSE_BUTTON
						Set objChild= Nothing
						Set objDesc = Nothing
						Set DeviceReplay = Nothing
						wait 1
					Next

					' Compare Error Message
					If  StrErrorMsg <> "" Then
						If  Instr(1,sDispMsg,StrErrorMsg)>0 Then
							Fn_SISW_Mech_ViewerTabErrorHandle=True
						Else
							Fn_SISW_Mech_ViewerTabErrorHandle=False
						End If
					Else
						Fn_SISW_Mech_ViewerTabErrorHandle=True	
					End If
	End Select

	Set objErrorDialog=Nothing
End Function



'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'/$$$$
'/$$$$   FUNCTION NAME   :   Fn_SISW_Mechatronics_SPMNavigationTree(StrAction,StrNodeName,StrMenu,sInfo1.sInfo2)
'/$$$$
'/$$$$   DESCRIPTION        :  Perform Variour sOperations in the  Software Parameter manager Navigation tree
'/$$$$
'/$$$$    PARAMETERS      :   1.) StrAction : Action to be performed
'/$$$$                                     		   2.) StrNodeName : Path of the Node on which the Desired operation has to be performed
'/$$$$											  3.) StrMenu  : Context menu path in case context menu is invoked
'/$$$$											  4.) sInfo1	  :  For future use in case any new functionality is added
'/$$$$											  5.) sInfo2	  :  For future use in case any new functionality is added
'/$$$$
'/$$$$
'/$$$$
'/$$$$    Function Calls       :   Fn_WriteLogFile(),   Fn_UI_JavaTreeGetItemPathExt(),  Fn_ReadyStatusSync()
'/$$$$									  
'/$$$$
'/$$$$	 HISTORY           :   AUTHOR                 DATE        VERSION
'/$$$$
'/$$$$    CREATED BY     :   SHREYAS           31/10/2012         1.0
'/$$$$
'/$$$$    REVIWED BY     :   Shreyas
'/$$$$
'/$$$$
'/$$$$   EXAMPLE          : 	bReturn=Fn_SISW_Mechatronics_SPMNavigationTree("Expand","A-break1:Memory Layouts","","","")
'/$$$$ 										bReturn=Fn_SISW_Mechatronics_SPMNavigationTree("Select","A-break1:Memory Layouts","","","")
'/$$$$ 										bReturn=Fn_SISW_Mechatronics_SPMNavigationTree("Exist","A-break1:Memory Layouts","","","")
' 												bReturn=Fn_SISW_Mechatronics_SPMNavigationTree("PopupMenuSelect","A-break1:Memory Layouts","Add:Memory Block...","","","")
'/$$$$
'						Modified By 							Changes Done																Date									Version
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'					Pranav Ingle						Modified Case "GetChildrenList"									06-Nov-2012								1.1
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Public Function Fn_SISW_Mechatronics_SPMNavigationTree(StrAction,StrNodeName,StrMenu,sInfo1,sInfo2)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_Mechatronics_SPMNavigationTree"

	Dim NodeLists, intNodeCount, intCount, StrExist, aMenuList, sTreeItem,sCmpItm
	Dim objJavaWindowMyTc, objJavaTreeNav,ArrNodeName
	Dim ArrStrcomp, sArrStr1,sArrStr2, iCounter
	Dim iRows, colonCnt
	Dim iItemCount, aNodePath,  iInstance, instCount, aNodes
	Dim sPath, sEle ,iCnt, bFound
	Dim iLen,iIndex,iTotal,iCount,sReturn,iReturn,arr
	Dim iPath,iVal,iPath1,arrNode
	Dim arrStrNode,echStrNode,oCurrentNode
	Dim sParentPath

         'Variable Declaration
	Dim sItemPath,aStrNode, i
	Dim iInstanceCnt, iOccCnt

	Set objJavaWindowMyTc = JavaWindow("Mechatronics")

	Select Case StrAction

		'----------------------------------------------------------------------- For selecting single node -------------------------------------------------------------------------
		Case "Select"
				'Initial Item Path
				aStrNode = Split (StrNodeName, ":")
				For i =0 to UBound(aStrNode)-1
					If sParentPath = "" Then
							sParentPath  = aStrNode(i)
					Else
							sParentPath  = sParentPath + ":" + aStrNode(i)
					End If
				Next
            
				iPath = Fn_UI_JavaTreeGetItemPathExt("Fn_SISW_Mechatronics_SPMNavigationTree", JavaWindow("Mechatronics").JavaTree("SPMNavTree"), StrNodeName , ":", "@")
				If iPath=False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select Node [" + StrNodeName + "] of SPMNavTree")
					Fn_SISW_Mechatronics_SPMNavigationTree = False
				Else
					JavaWindow("Mechatronics").JavaTree("SPMNavTree").Select iPath
					Call Fn_ReadyStatusSync(1)
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully DeSelected Node [" + StrNodeName + "] of SPMNavTree")
					Fn_SISW_Mechatronics_SPMNavigationTree = True
				End If

		Case "Expand"
			iPath = Fn_UI_JavaTreeGetItemPathExt("Fn_SISW_Mechatronics_SPMNavigationTree", JavaWindow("Mechatronics").JavaTree("SPMNavTree"), StrNodeName , ":", "@")
			If iPath=False Then
				 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Expand Node [" + StrNodeName + "] of SPMNavTree")
				  Fn_SISW_Mechatronics_SPMNavigationTree = False
			Else
				JavaWindow("Mechatronics").JavaTree("SPMNavTree").Expand iPath
'				Call Fn_ReadyStatusSync(1)
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Expanded Node [" + StrNodeName + "] of SPMNavTree")
				Fn_SISW_Mechatronics_SPMNavigationTree = True
			End If

		' - - - - - - - - - - Collaplse Node
		Case "Collapse"
			iPath = Fn_UI_JavaTreeGetItemPathExt("Fn_SISW_Mechatronics_SPMNavigationTree", JavaWindow("Mechatronics").JavaTree("SPMNavTree"), StrNodeName , ":", "@")
			If iPath=False Then
				 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Collapse Node [" + StrNodeName + "] of SPMNavTree")
				  Fn_SISW_Mechatronics_SPMNavigationTree = False
			Else
				JavaWindow("Mechatronics").JavaTree("SPMNavTree").Collapse iPath
				Call Fn_ReadyStatusSync(1)
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Collapse Node [" + StrNodeName + "] of SPMNavTree")
				Fn_SISW_Mechatronics_SPMNavigationTree = True
			End If
		' - - - - - - - - - - Pop Up Menu Select
		Case "PopupMenuSelect"
			Set objJavaTreeNav = JavaWindow("Mechatronics").JavaTree("SPMNavTree")
					'Build the Popup menu to be selected
					aMenuList = split(StrMenu, ":",-1,1)
					intCount = Ubound(aMenuList)

					'Select node
                    iPath = Fn_UI_JavaTreeGetItemPathExt("Fn_SISW_Mechatronics_SPMNavigationTree", JavaWindow("Mechatronics").JavaTree("SPMNavTree"), StrNodeName , ":", "@")
					If iPath=False Then
						 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Select Node [" + StrNodeName + "] of SPMNavTree")
						  Fn_SISW_Mechatronics_SPMNavigationTree = False
						  Exit Function
					Else
						JavaWindow("Mechatronics").JavaTree("SPMNavTree").Select iPath
						Call Fn_ReadyStatusSync(1)
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Selected Node [" + StrNodeName + "] of SPMNavTree")
						Fn_SISW_Mechatronics_SPMNavigationTree = True
					End If

					'Open context menu
					Call Fn_UI_JavaTree_OpenContextMenu("Fn_SISW_Mechatronics_SPMNavigationTree",objJavaWindowMyTc,"SPMNavTree",iPath)
                    
					'Select Menu action
					Select Case intCount
						Case "0"
							 StrMenu = JavaWindow("Mechatronics").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0))
						Case "1"
							StrMenu = JavaWindow("Mechatronics").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1))
						Case "2"
							StrMenu = JavaWindow("Mechatronics").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1),aMenuList(2))
						Case Else
							Fn_SISW_Mechatronics_SPMNavigationTree = FALSE
							Exit Function
					End Select

					JavaWindow("Mechatronics").WinMenu("ContextMenu").Select StrMenu
					If Err.number < 0 Then
						Fn_SISW_Mechatronics_SPMNavigationTree = False
					Else
						Fn_SISW_Mechatronics_SPMNavigationTree = True
					End If
        ' - - - - - - - - - - PopUp Menu Existance on multi Select
		Case "MultiSelectContextMenuExist"
				NodeLists = split(StrNodeName,",",-1,1)
				Set objJavaTreeNav = JavaWindow("Mechatronics").JavaTree("SPMNavTree")
				Call Fn_SISW_Mechatronics_SPMNavigationTree("Multiselect",StrNodeName,"","","")
				iPath=Fn_UI_JavaTreeGetItemPath(JavaWindow("Mechatronics").JavaTree("SPMNavTree"),NodeLists(0))
				Call Fn_UI_JavaTree_OpenContextMenu("Fn_SISW_Mechatronics_SPMNavigationTree",objJavaWindowMyTc,"SPMNavTree",iPath,"","")
				If JavaWindow("Mechatronics").WinMenu("ContextMenu").GetItemProperty (StrMenu,"Exists") = True Then
					Fn_SISW_Mechatronics_SPMNavigationTree = True
				Else
					Fn_SISW_Mechatronics_SPMNavigationTree = False
			  	End If
		' - - - - - - - - - - Double Click on Node
		Case "DoubleClick"
			Dim intX, intY, intWidth, intHeight, strComputer, sOSName, objWMIService, oss, os
			intX = 0
			intY = 0
			intWidth = 0
			intHeight = 0
        			
			strComputer = "." 
			Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
			Set oss = objWMIService.ExecQuery ("Select * from Win32_OperatingSystem") 
			For Each os in oss 
			sOSName = os.Caption
			Next

			iPath = Fn_UI_JavaTreeGetItemPathExt("Fn_SISW_Mechatronics_SPMNavigationTree", JavaWindow("Mechatronics").JavaTree("SPMNavTree"), StrNodeName , ":", "@")
			
			If Instr(ucase(sOSName),"XP") > 0 then
			
             wait 1
			JavaWindow("Mechatronics").JavaTree("SPMNavTree").Activate iPath
			
			Set objWMIService = Nothing
			Set oss = Nothing

			Fn_SISW_Mechatronics_SPMNavigationTree = True
			
			Else
				'JavaWindow("Mechatronics").JavaTree("SPMNavTree").Activate iPath
				intX = objNodeBounds.x
				intY = objNodeBounds.y
				intWidth = objNodeBounds.width
				intHeight = objNodeBounds.height

				Set objNodeBounds = nothing

				JavaWindow("Mechatronics").JavaTree("SPMNavTree").DblClick intX + intWidth/2, intY + intHeight/2, "LEFT"
				If Err.number < 0 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to DoubleClick Node [" + StrNodeName + "] of SPMNavTree")
					Fn_SISW_Mechatronics_SPMNavigationTree = False
					Exit Function
				End If

				Call Fn_ReadyStatusSync(1)
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully DoubleClick Node [" + StrNodeName + "] of SPMNavTree")
				Fn_SISW_Mechatronics_SPMNavigationTree = True
			End If
		' - - - - - - - - - - Popup Menu operation on Multi Selected Nodes
		Case "MultiSelectContextMenu"
					NodeLists = split(StrNodeName,",",-1,1)
					aMenuList = split(StrMenu, ":",-1,1)
					intCount = Ubound(aMenuList)
					Set objJavaTreeNav = JavaWindow("Mechatronics").JavaTree("SPMNavTree")

					'Select multiple node
					Call Fn_SISW_Mechatronics_SPMNavigationTree("Multiselect", StrNodeName, "","","")
					'Open context menu
					iPath = Fn_UI_JavaTreeGetItemPathExt("Fn_SISW_Mechatronics_SPMNavigationTree", JavaWindow("Mechatronics").JavaTree("SPMNavTree"), NodeLists(0) , ":", "@")
					Call Fn_UI_JavaTree_OpenContextMenu("Fn_SISW_Mechatronics_SPMNavigationTree",objJavaWindowMyTc,"SPMNavTree",iPath)
					Select Case intCount
						Case "0"
							 StrMenu = JavaWindow("Mechatronics").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0))
						Case "1"
							StrMenu = JavaWindow("Mechatronics").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1))
						Case "2"
							StrMenu = JavaWindow("Mechatronics").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1),aMenuList(2))
						Case Else
							Fn_SISW_Mechatronics_SPMNavigationTree = False
							Exit Function
					End Select
					JavaWindow("Mechatronics").WinMenu("ContextMenu").Select StrMenu
					If Err.number < 0 Then
						Fn_SISW_Mechatronics_SPMNavigationTree = False
					else
						Fn_SISW_Mechatronics_SPMNavigationTree = True
					End If	
		' - - - - - - - - - - Existance of Node
		Case "Exist"
				Set objJavaTreeNav = JavaWindow("Mechatronics").JavaTree("SPMNavTree")
				iPath = Fn_UI_JavaTreeGetItemPathExt("Fn_SISW_Mechatronics_SPMNavigationTree", JavaWindow("Mechatronics").JavaTree("SPMNavTree"), StrNodeName , ":", "@")
				If iPath=False Then
				'iPath = Fn_UI_getJavaTreeIndex(JavaWindow("Mechatronics").JavaTree("SPMNavTree"),StrNodeName)
				'If iPath = -1 Then
					 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Node [" + StrNodeName + "] Not exist in SPMNavTree")
					  Fn_SISW_Mechatronics_SPMNavigationTree = False
				Else
					aNodePath = split(replace(iPath,"#",""),":")
					Fn_SISW_Mechatronics_SPMNavigationTree = true
					Set oCurrentNode = JavaWindow("Mechatronics").JavaTree("SPMNavTree").Object
					For iCnt = 0 to UBound(aNodePath) -1
						Set oCurrentNode = oCurrentNode.GetItem(aNodePath(iCnt))
						If cBool(oCurrentNode.getExpanded()) = False Then
							Fn_SISW_Mechatronics_SPMNavigationTree = false
							Exit for
						End If
					Next
					If Fn_SISW_Mechatronics_SPMNavigationTree Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Node [" + StrNodeName + "] Exist in SPMNavTree")
					End If
				End If
				Set oCurrentNode = Nothing
		' - - - - - - - - - - Existance of Popup Menu
		Case "PopupMenuExist"
					aMenuList = split(StrMenu, ":",-1,1)
					intCount = Ubound(aMenuList)
					Set objJavaTreeNav = JavaWindow("Mechatronics").JavaTree("SPMNavTree")
					'Open context menu
					iPath = Fn_UI_JavaTreeGetItemPathExt("Fn_SISW_Mechatronics_SPMNavigationTree", JavaWindow("Mechatronics").JavaTree("SPMNavTree"), StrNodeName , ":", "@")
					Call Fn_UI_JavaTree_OpenContextMenu("Fn_SISW_Mechatronics_SPMNavigationTree",objJavaWindowMyTc,"SPMNavTree",iPath)
					Select Case intCount
						Case "0"
							 StrMenu = JavaWindow("Mechatronics").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0))
                        Case "1"
							StrMenu = JavaWindow("Mechatronics").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1))
                        Case "2"
							StrMenu = JavaWindow("Mechatronics").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1),aMenuList(2))
                        Case Else
							Fn_SISW_Mechatronics_SPMNavigationTree = False
                        Exit Function
					End Select
					If JavaWindow("Mechatronics").WinMenu("ContextMenu").GetItemProperty (StrMenu,"Exists") = True Then
						Fn_SISW_Mechatronics_SPMNavigationTree = True
					Else
						Fn_SISW_Mechatronics_SPMNavigationTree = False
					End If
					Call Fn_KeyBoardOperation("SendKeys", "{ESC}")
		' - - - - - - - - - - Checking State of Popup Menu		
		Case "PopupMenuEnabled"
					aMenuList = split(StrMenu, ":",-1,1)
					intCount = Ubound(aMenuList)
					Set objJavaTreeNav = JavaWindow("Mechatronics").JavaTree("SPMNavTree")
					'Open context menu
					iPath=Fn_UI_JavaTreeGetItemPath(JavaWindow("Mechatronics").JavaTree("SPMNavTree"),StrNodeName)
                	Call Fn_UI_JavaTree_OpenContextMenu("Fn_SISW_Mechatronics_SPMNavigationTree",objJavaWindowMyTc,"SPMNavTree",iPath)
					Select Case intCount
						Case "0"
							 StrMenu = JavaWindow("Mechatronics").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0))
                        Case "1"
							StrMenu = JavaWindow("Mechatronics").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1))
                        Case "2"
							StrMenu = JavaWindow("Mechatronics").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1),aMenuList(2))
                        Case Else
							Fn_SISW_Mechatronics_SPMNavigationTree = FALSE
                        Exit Function
					End Select
				If JavaWindow("Mechatronics").WinMenu("ContextMenu").GetItemProperty (StrMenu,"Enabled") = True Then
					Fn_SISW_Mechatronics_SPMNavigationTree = True
				Else
					Fn_SISW_Mechatronics_SPMNavigationTree = False
			  	End If
		'------------------- Checks That item is inactively focused Or Not for single node OR Multiple Node(comma "," SEPERATED)---------------
		Case "GetSelected"
		wait 5
			Set objJavaTreeNav = JavaWindow("Mechatronics").JavaTree("SPMNavTree")
				
				ArrStrcomp = Split(objJavaTreeNav.GetROProperty("value") ,"",-1,1)
				sArrStr2 = ArrStrcomp(0)
				For iCounter = 1 To ubound(ArrStrcomp)
					sArrStr2 = sArrStr2 & "," & ArrStrcomp(iCounter)
				Next
				If sArrStr2 = StrNodeName Then
				   'Writing Log
				   Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Java Tree Multiple Node ["+StrNodeName+"] is Selected .")
				   Fn_SISW_Mechatronics_SPMNavigationTree = TRUE
				Else
				   'Writing Log
				   Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Java Tree Multiple  Node ["+StrNodeName+"] is Not Selected .")
				   Fn_SISW_Mechatronics_SPMNavigationTree = FALSE
			End If
		'------------------- Get Selected Node with Path---------------
		Case "GetSelectedNodePath"
			Set oCurrentNode = JavaWindow("Mechatronics").JavaTree("SPMNavTree").Object.getFocusItem()
			Fn_SISW_Mechatronics_SPMNavigationTree = oCurrentNode.getData().toString() 
			
			Do while lcase(typename(oCurrentNode.getParentItem())) <> "nothing"
				Set oCurrentNode = oCurrentNode.getParentItem()
				Fn_SISW_Mechatronics_SPMNavigationTree = oCurrentNode.getData().toString() & ":" & Fn_SISW_Mechatronics_SPMNavigationTree 
			Loop
		' - - - - - - - - - - Getting Index of Node
		Case "GetIndex"
			'Index of Item1
			 Fn_SISW_Mechatronics_SPMNavigationTree = Fn_UI_getJavaTreeIndex(JavaWindow("Mechatronics").JavaTree("SPMNavTree"), StrNodeName)
		' - - - - - - - - - - Getting Child Item Count
		Case "GetChildItemCount"
				If Fn_SISW_Mechatronics_SPMNavigationTree("Expand",StrNodeName,"","","")=True Then
					arrStrNode = Split (StrNodeName, ":")
					Set oCurrentNode = JavaWindow("Mechatronics").JavaTree("SPMNavTree").Object.getItem(0)
					intNodeCount=0
					For each echStrNode In arrStrNode
						iRows = oCurrentNode.getItemCount()
						For iCounter = 0 to iRows - 1
							If oCurrentNode.getItem(iCounter).getData().toString() = echStrNode Then
								intNodeCount = oCurrentNode.getItem(iCounter).getItemCount()
								Exit For
							End If
						Next
					Next 
					Fn_SISW_Mechatronics_SPMNavigationTree = intNodeCount
					Set oCurrentNode=Nothing
				Else
					Fn_SISW_Mechatronics_SPMNavigationTree = False
				End If
		' - - - - - - - - - - Selecting Range of Nodes
		Case "SelectRange"
					ReDim ArrNodeName(2)
					ArrNodeName = Split(StrNodeName,"|")
				
					iPath = Fn_UI_JavaTreeGetItemPathExt("Fn_SISW_Mechatronics_SPMNavigationTree", JavaWindow("Mechatronics").JavaTree("SPMNavTree"), ArrNodeName(0) , ":", "@")
					iPath1 = Fn_UI_JavaTreeGetItemPathExt("Fn_SISW_Mechatronics_SPMNavigationTree", JavaWindow("Mechatronics").JavaTree("SPMNavTree"), ArrNodeName(1) , ":", "@")
					If iPath=False Then
						 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Select Node [" + StrNodeName + "] of SPMNavTree")
						  Fn_SISW_Mechatronics_SPMNavigationTree = False
					Else
						JavaWindow("Mechatronics").JavaTree("SPMNavTree").SelectRange iPath,iPath1
						Call Fn_ReadyStatusSync(1)
						Fn_SISW_Mechatronics_SPMNavigationTree = True
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
				iCnt = Fn_SISW_Mechatronics_SPMNavigationTree( "GetIndex" , sPath , "","","")
				iItemCount = Fn_SISW_Mechatronics_SPMNavigationTree( "GetChildrenList" , sPath , "","","")
				For iCounter=0 To UBound(iItemCount)
					If Trim(iItemCount(iCounter))=Trim(aNodePath( UBound(aNodePath))) Then
						iInstance = iInstance+1
					End If
				Next
				Fn_SISW_Mechatronics_SPMNavigationTree = iInstance
			'- - - - - - - - - - - -  Retruns All Childs of any given Node in the tree in form of an array - - - - - - - - - - - - - - -
			Case "GetChildrenList"
					sReturn=""
					If Fn_SISW_Mechatronics_SPMNavigationTree("Expand",StrNodeName,"","","")=True Then
						arrStrNode = Split (StrNodeName, ":")
						If UBound(arrStrNode)=0 And  lCase(arrStrNode(0))="home" Then
								Set oCurrentNode = JavaWindow("Mechatronics").JavaTree("SPMNavTree").Object.getItem(0)
								intNodeCount = oCurrentNode.getItemCount()
								For iCount=0 To intNodeCount-1
									If iCount=0 Then
										sReturn=oCurrentNode.getItem(iCount).getData().getDisplayName()
									Else
										sReturn=sReturn+","+oCurrentNode.getItem(iCount).getData().getDisplayName()
									End If
								Next
									arr = Split(sReturn,",")
									Fn_SISW_Mechatronics_SPMNavigationTree = arr
									Set oCurrentNode=Nothing
									Exit Function
						Else
								Set oCurrentNode = JavaWindow("Mechatronics").JavaTree("SPMNavTree").Object.getItem(0)
								intNodeCount=0
								For each echStrNode In arrStrNode
									iRows = oCurrentNode.getItemCount()
									For iCounter = 0 to iRows - 1
										If oCurrentNode.getItem(iCounter).getData().getDisplayName() = echStrNode Then
											Set oCurrentNode=oCurrentNode.getItem(iCounter)
											intNodeCount = oCurrentNode.getItemCount()
											Exit For
										End If
									Next
								Next 
								For iCount=0 To intNodeCount-1
									If iCount=0 Then
										sReturn=oCurrentNode.getItem(iCount).getData().getDisplayName()
									Else
										sReturn=sReturn+","+oCurrentNode.getItem(iCount).getData().getDisplayName()
									End If
								Next
'								arr = Split(sReturn,",")
'								Fn_SISW_Mechatronics_SPMNavigationTree = arr
								Fn_SISW_Mechatronics_SPMNavigationTree = sReturn
								Set oCurrentNode=Nothing
						End If
					Else
						Fn_SISW_Mechatronics_SPMNavigationTree = False
					End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case Else
						Fn_SISW_Mechatronics_SPMNavigationTree = False
	End Select

	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), StrAction &" Sucessfully completed on Node [" + StrNodeName + "] of JavaTree of function Fn_SISW_Mechatronics_SPMNavigationTree")
	Set objJavaWindowMyTc = nothing
	Set objJavaTreeNav = nothing
End Function

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'/$$$$
'/$$$$   FUNCTION NAME   :   Fn_SISW_Mechatronics_CreateNewBlock(sNodeName,sBlockName,sButtons,sInfo1,sInfo2,sInfo3)
'/$$$$
'/$$$$   DESCRIPTION        :  Create a new Block under a node in Software Parameter manager
'/$$$$
'/$$$$    PARAMETERS      :   1.) sNodeName : Valid Node name under which new block has to be added
'/$$$$                                     		   2.) sBlockName : Valid Block name
'/$$$$											  3.) StrMenu  : Context menu path in case context menu is invoked
'/$$$$											  4.)sButtons : Buttons to be clicked on the New block dialog
'/$$$$											  5.) sInfo1	  :  For future use in case any new functionality is added
'/$$$$											 6.) sInfo2	  :  For future use in case any new functionality is added
'/$$$$											7.)sInfo3  :  For future Use in case any new functionality is added
'/$$$$
'/$$$$
'/$$$$
'/$$$$    Function Calls       :   Fn_WriteLogFile(),   Fn_SISW_Mechatronics_SPMNavigationTree()
'/$$$$									  
'/$$$$
'/$$$$	 HISTORY           :   AUTHOR                 DATE        VERSION
'/$$$$
'/$$$$    CREATED BY     :   SHREYAS           31/10/2012         1.0
'/$$$$
'/$$$$    REVIWED BY     :   Shreyas
'/$$$$
'/$$$$
'/$$$$   EXAMPLE          : 	bReturn=bReturn=Fn_SISW_Mechatronics_CreateNewBlock("A-break1:Memory Layouts:000034/A;1-lay1:New Block","New Block","","","","")
'/$$$$ 										
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

Public Function Fn_SISW_Mechatronics_CreateNewBlock(sNodeName,sBlockName,sButtons,sInfo1,sInfo2,sInfo3)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_Mechatronics_CreateNewBlock"
Dim objBlock,bReturn,StrMenu,sMsg
	Environment.Value("sPath") = Fn_GetEnvValue("User", "AutomationDir")

Set objBlock=Fn_SISW_Mechatronics_GetObject("NewBlock")
		
		If sNodeName<>"" Then
				StrMenu=Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\Mechatronics\Mechatronics_Menu.xml","NewBlock")
				bReturn=Fn_SISW_Mechatronics_SPMNavigationTree("PopupMenuSelect",sNodeName,StrMenu,"","")
				If bReturn=true Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Successfully called menu [ " & StrMenu & " ] on the Node ["+sNodeName+"]")
				Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to call menu [ " & StrMenu & " ] on the Node ["+sNodeName+"]")
						Fn_SISW_Mechatronics_CreateNewBlock=false
						Exit Function
				End If
		End If

'Check for the Checkout message & checkout if required

			If objBlock.JavaStaticText("Message").Exist(5) then
				If objBlock.JavaButton("Yes").Exist(5) Then
					objBlock.JavaButton("Yes").Click micLeftBtn
				End If
			End if

'Enter the Name for the new block

		If sBlockName<>"" Then
		
				'clear the contents
				objBlock.JavaEdit("BlockName").Set ""
				wait 1
				objBlock.JavaEdit("BlockName").Set sBlockName
				wait 1
		
				'Check if the error occurs

				If objBlock.JavaEdit("BlockNameError").Exist(5) then
						sMsg=objBlock.JavaEdit("BlockNameError").GetROProperty ("value")
						If instr(sMsg,"already exists")>0 Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Block name ["+sBlockName+"] entered is already used")
								Fn_SISW_Mechatronics_CreateNewBlock=false
								objBlock.Close
								Exit Function
						End If
				Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Successfully entered Block name as ["+sBlockName+"]")
				End if
		End if

			If sButtons<>"" Then
						'Click on the Required Buttons
						objBlock.JavaButton("OK").Click micLeftBtn
						wait 1
			End If

		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Successfully created new Block  ["+sBlockName+"]")
		Fn_SISW_Mechatronics_CreateNewBlock=true
		If 		objBlock.Exist(5) Then
			objBlock.Close
		End If
Set objBlock=nothing
End Function

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name			:	Fn_SISW_Mech_DeleteOverrideRecord
'@@
'@@    Description				:	Function Used to perform operations on  Delete Override Dialog
'@@
'@@    Parameters			   	:	1. sAction		: Action To Perform
'@@								:			  2. sMessage	  :  Message To Verify From Delete Override Dialog
'@@							 	:			  3. sButton 		: Column Name
'@@
'@@    Return Value		   	   	: 	True Or False
'@@
'@@    Pre-requisite			:	Delete Override Dialog should be open					
'@@
'@@    Examples					:	
'@@    									1.	Call Fn_SISW_Mech_DeleteOverrideRecord("Delete", "", "Yes")
'@@    									2.  Call Fn_SISW_Mech_DeleteOverrideRecord("VerifyMessage", "Do you want to delete the formula?", "Yes")
'@@	History				:	
'@@	Developer Name			Date			 Rev. No.	   Changes Done		 Reviewer
'@@------------------------------------------------------------------------------------------------------------------------------
'@@	Pranav Ingle			31-Oct-2012		  1.0			created
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_SISW_Mech_DeleteOverrideRecord(sAction, sMessage, sButton)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_Mech_DeleteOverrideRecord"
	Dim objDeleteOverrideRecord, sDispMsg
	
	Set objDeleteOverrideRecord = JavaWindow("Mechatronics").JavaWindow("DeleteOverrideRecord")
	Fn_SISW_Mech_DeleteOverrideRecord = False
	' add call to select Traceability Matrix Tab
    
	If Fn_UI_ObjectExist("Fn_SISW_Mech_DeleteOverrideRecord",objDeleteOverrideRecord ) = False  Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_Mech_DeleteOverrideRecord ] Nat Table object of [ Traceability Matrix ] is not visible.")
		Exit function
	End If

	Select Case sAction
		Case "Delete"
			Fn_SISW_Mech_DeleteOverrideRecord = Fn_Button_Click("Fn_SISW_Mech_DeleteOverrideRecord", objDeleteOverrideRecord,"Yes")

		Case "VerifyMessage"
        	sDispMsg = objDeleteOverrideRecord.JavaStaticText("sMsg").GetROProperty("value")
			If  Instr(1,sDispMsg,sMessage) > 0 Then
				Fn_SISW_Mech_DeleteOverrideRecord=True
			Else
				Fn_SISW_Mech_DeleteOverrideRecord=False
			End If

    		Call Fn_Button_Click("Fn_SISW_Mech_DeleteOverrideRecord", objDeleteOverrideRecord, sButton)

	' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_Mech_DeleteOverrideRecord ] Invalid case [ " & sAction & " ].")
	' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
	End Select
	If  Fn_SISW_Mech_DeleteOverrideRecord <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [ Fn_SISW_Mech_DeleteOverrideRecord ] executed successfuly with case [ " & sAction & " ].")
	End If
	Set objDeleteOverrideRecord = Nothing
End Function

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name			:	Fn_SISW_Mech_AvailableParam_BreakdownOperations
'@@
'@@    Description				:	Function Used to perform operations on  Show Traceability Matrix window in System Engineering
'@@
'@@    Parameters			   	:	1. sAction	: 		Action [Type of Attribute Group]
'@@												2. Dim dicAvailableParameters
'@@														Set dicAvailableParameters = CreateObject("Scripting.Dictionary")
'@@														With dicAvailableParameters
'@@															.Add  "TabName",""
'@@															.Add "TableHeader",""
'@@															.Add "Object",""
'@@															.Add "ColName",""
'@@															.Add "Value",""
'@@															.Add "ShowAllAvailableParameters",""      						--   Available Parameters Tab 
'@@															.Add "ShowAllAssignedParameters",""							   --  	Assigned Parameters  Tab		
'@@														End With
'@@	
'@@								:			  		dicAvailableParameters("TabName")	:     Available Parameters / Assigned Parameters
'@@								
'@@								:			  		dicAvailableParameters("TableHeader")	:    Parameter Definition / Override
'@@													 dicAvailableParameters("Object")			:  Object Column Row text [for Cases : Select, SelectRow, Assign ]
'@@								:			  																				ColName Row  [for Cases : CellVerify ] 
'@@							 	:			  	  dicAvailableParameters("ColName") 	  : Column Name        ----  [ Default Will be Object ]  -----
'@@							 	:			  	  dicAvailableParameters("Value") 			: 	Value To Verify 
'@@							 	:			 	  dicAvailableParameters("ShowAllAvailableParameters") 	: Show All Available Parameters Check Box Set [true: checked, false : Unchecked]
'@@							 	
'@@							 	:			 3. sPopupMenu 	: Popup Menu select
'@@
'@@    Return Value		   	   	: 	True Or False
'@@
'@@    Pre-requisite			:	Software Parameter Manager perspective should be activated and Available parameter Tab is Selected.						
'@@
'@@    Examples					:	
'@@										Case "SelectCell" , "Assign", "Unassign", "PopupMenuSelect"  , "DoubleClick"
'@@											dicAvailableParameters.RemoveAll
'@@											dicAvailableParameters("TabName") = "Available Parameters"
'@@											dicAvailableParameters("TableHeader") = "Parameter Definition"
'@@											dicAvailableParameters("Object") = "000025/A;1-bool"
'@@											dicAvailableParameters("ShowAllAvailableParameters") =  true
'@@											dicAvailableParameters("ColName")  =  "Object"
'@@
'@@									  	Call Fn_SISW_Mech_AvailableParam_BreakdownOperations("SelectCell", dicAvailableParameters,  "")
'@@									  	Call Fn_SISW_Mech_AvailableParam_BreakdownOperations("DoubleClick", dicAvailableParameters,  "")
'@@										Call Fn_SISW_Mech_AvailableParam_BreakdownOperations("Assign", dicAvailableParameters,  "")
'@@										Call Fn_SISW_Mech_AvailableParam_BreakdownOperations("Unassign", dicAvailableParameters,  "")
'@@									 	Call Fn_SISW_Mech_AvailableParam_BreakdownOperations("PopupMenuSelect", dicAvailableParameters,  "Delete Override")
'@@									 	Call Fn_SISW_Mech_AvailableParam_BreakdownOperations("PopupMenuExist", dicAvailableParameters,  "Delete Override")
'@@
'@@										Case "MultiSelect", "MultiPopupSelect", "MultiSelectAssign", "MultiSelectUnassign", "MultiSelectPopupMenuExist"
'@@											dicAvailableParameters.RemoveAll
'@@											dicAvailableParameters("TabName") = "Available Parameters"
'@@											dicAvailableParameters("TableHeader") = "Parameter Definition"
'@@											dicAvailableParameters("Object") =  "Hex45678/A;1-ParmDefHex12345~BCD45678/A;1-ParmDefBCD12345"
'@@											dicAvailableParameters("ColName")  =  "Object"
'@@
'@@									  	Call Fn_SISW_Mech_AvailableParam_BreakdownOperations("MultiSelect", dicAvailableParameters,  "")
'@@									  	Call Fn_SISW_Mech_AvailableParam_BreakdownOperations("MultiPopupSelect", dicAvailableParameters,  "Delete Override")
'@@										Call Fn_SISW_Mech_AvailableParam_BreakdownOperations("MultiSelectAssign", dicAvailableParameters, "")
'@@										Call Fn_SISW_Mech_AvailableParam_BreakdownOperations("MultiSelectUnassign", dicAvailableParameters, "")
'@@									 	Call Fn_SISW_Mech_AvailableParam_BreakdownOperations("MultiSelectPopupMenuExist", dicAvailableParameters,  "Delete Override")
'@@										
'@@										Case "SelectRow"
'@@											dicAvailableParameters.RemoveAll
'@@											dicAvailableParameters("TabName") = "Available Parameters"
'@@											dicAvailableParameters("Object") = "000025/A;1-bool"
'@@											dicAvailableParameters("ShowAllAvailableParameters") =  true
'@@	
'@@										Call Fn_SISW_Mech_AvailableParam_BreakdownOperations("SelectRow", dicAvailableParameters,  "")
'@@	
'@@										Case "GetColumnIndex"    								----   	Return -1  if Col Not Found
'@@											dicAvailableParameters.RemoveAll
'@@											dicAvailableParameters("TabName") = "Available Parameters"
'@@											dicAvailableParameters("TableHeader") = "Parameter Definition"
'@@											dicAvailableParameters("ColName")  =  "Conversion Rule"
'@@	
'@@										Call Fn_SISW_Mech_AvailableParam_BreakdownOperations("GetColumnIndex",dicAvailableParameters, "")
'@@										
'@@										Case "HeaderPopupMenuSelect" 
'@@											dicAvailableParameters.RemoveAll
'@@											dicAvailableParameters("TabName") = "Available Parameters"
'@@											dicAvailableParameters("TableHeader") = "Parameter Definition"
'@@											dicAvailableParameters("ColName")  =  "Conversion Rule"
'@@	
'@@										Call Fn_SISW_Mech_AvailableParam_BreakdownOperations("HeaderPopupMenuSelect",dicAvailableParameters, "Hide column")
'@@	
'@@										Case "SetCellData"
'@@											dicAvailableParameters.RemoveAll
'@@											dicAvailableParameters("TabName") = "Assigned Parameters"
'@@											dicAvailableParameters("TableHeader") = "Override"
'@@											dicAvailableParameters("Object") = "Hex45678/A;1-ParmDefHex12345"
'@@											dicAvailableParameters("ShowAllAvailableParameters") =  true
'@@											dicAvailableParameters("ColName")  =  "Conversion Rule"
'@@											dicAvailableParameters("Value")  =  "7*x*x+8*x+9"
'@@
'@@										Call Fn_SISW_Mech_AvailableParam_BreakdownOperations("SetCellData",dicAvailableParameters, "")
'@@										
'@@										Case "CellVerify"
'@@											dicAvailableParameters.RemoveAll
'@@											dicAvailableParameters("TabName") = "Available Parameters"
'@@											dicAvailableParameters("TableHeader") =  "Parameter Definition"
'@@											dicAvailableParameters("Object") = "Hex45678/A;1-ParmDefHex12345"
'@@											dicAvailableParameters("ColName")  =  "Minimum Values~Maximum Values~Initial Values"
'@@											dicAvailableParameters("Value")  =  "{0x00}~{0x00}~{0x00}"
'@@
'@@										Call Fn_SISW_Mech_AvailableParam_BreakdownOperations("CellVerify",dicAvailableParameters, "")
'@@										
'@@										Case "Ascending", "Descending"   - 
'@@											dicAvailableParameters.RemoveAll
'@@											dicAvailableParameters("TabName") = "Available Parameters"
'@@											dicAvailableParameters("TableHeader") =  "Parameter Definition"
'@@											dicAvailableParameters("ColName")  =  "Table Definition"
'@@
'@@										Call Fn_SISW_Mech_AvailableParam_BreakdownOperations("Ascending",dicAvailableParameters, "")
'@@										
'@@		History				:	
'@@		Developer Name					Date					Rev. No.	   Changes Done		 										Reviewer
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@		Pranav Ingle						31-Oct-2012				1.0					created
'@@		Sandeep Navghane		15-Nov-2012			   1.1				 Added case : PopupMenuExist
'@@		Pranav Ingle						28-Nov-2012			  1.2				Added Case  "Ascending", "Descending"
'@@		Koustubh Watwe				  30-Nov-2012			1.2				 Modified case "GetRowIndex"
'@@		Pranav Ingle						30-Nov-2012			  1.3				Modified case "GetColumnIndex"
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Dim aTableHeader
Public Function Fn_SISW_Mech_AvailableParam_BreakdownOperations(sAction, dicAvailableParameters, sPopupMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_Mech_AvailableParam_BreakdownOperations"
	Dim objNatTable, iRowCnt, iColCnt, iCnt, sCol, dicparameter
	Dim aBounds, aMenu, arrValues, iCount, arrCol
	Dim bReturn, startIndex, endIndex, sRectangle
	Dim WshShell,objPopupMenu
	Dim arrObj, myDeviceReplay

	Set objNatTable = JavaWindow("Mechatronics").JavaObject("NatTable")
	Fn_SISW_Mech_AvailableParam_BreakdownOperations = False
	' add call to select Traceability Matrix Tab
	If Fn_UI_ObjectExist("Fn_SISW_Mech_AvailableParam_BreakdownOperations",objNatTable ) = False  Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_Mech_AvailableParam_BreakdownOperations ] Nat Table object of [ Traceability Matrix ] is not visible.")
		Exit function
	End If

	If dicAvailableParameters("ShowAll"&Trim(Replace(dicAvailableParameters("TabName")," ",""))) <> "" Then
		If cBool(dicAvailableParameters("ShowAll"&Trim(Replace(dicAvailableParameters("TabName")," ","")))) Then
			Call Fn_CheckBox_Set("Fn_SISW_Mech_AvailableParam_BreakdownOperations",  JavaWindow("Mechatronics"), "ShowAll"&Trim(Replace(dicAvailableParameters("TabName")," ","")), "ON")
		Else
			Call Fn_CheckBox_Set("Fn_SISW_Mech_AvailableParam_BreakdownOperations",  JavaWindow("Mechatronics"),"ShowAll"&Trim(Replace(dicAvailableParameters("TabName")," ","")), "OFF")
		End If
	End If
	If dicAvailableParameters("TableHeader") <>""Then
		aTableHeader=dicAvailableParameters("TableHeader")
	End If
	If dicAvailableParameters("TabName") = "Available Parameters" AND dicAvailableParameters("TableHeader") = "Parameter Definition" Then dicAvailableParameters("TableHeader") = ""
	
	Select Case sAction
		Case "SelectCell", "DoubleClick", "PopupMenuSelect", "SetCellData","PopupMenuExist"
			iRowCnt = Fn_SISW_Mech_AvailableParam_BreakdownOperations("GetRowIndex",dicAvailableParameters, "")
			If CInt(iRowCnt) = -1 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_Mech_AvailableParam_BreakdownOperations ] Row Not Found[ " & dicAvailableParameters("Object") & " ].")
				Exit Function
			End If

			sCol = Fn_SISW_Mech_AvailableParam_BreakdownOperations("GetColumnIndex",dicAvailableParameters, "")
			If CInt(sCol) = -1 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_Mech_AvailableParam_BreakdownOperations ] Column Not Found[ " & dicAvailableParameters("ColName") & " ].")
				Exit Function
			End If

			sRectangle = cStr(objNatTable.Object.getCellByPosition(sCol,iRowCnt).getBounds().toString())
			sRectangle =  Right(sRectangle, (Len(sRectangle) - Instr(1, sRectangle, "{", 1)))	
			sRectangle = replace(sRectangle,"}","")
			sRectangle = replace(sRectangle," ","")
			aBounds = split(sRectangle,",")

            If sAction = "SelectCell" Then
				objNatTable.Click cInt(aBounds(0)) + (cInt(aBounds(2))/2),  cInt(aBounds(1)) + (cInt(aBounds(3))/2)  , "LEFT"
				Fn_SISW_Mech_AvailableParam_BreakdownOperations = True
			ElseIf sAction = "DoubleClick" Then
				objNatTable.DblClick cInt(aBounds(0)-10) + (cInt(aBounds(2))/2),  cInt(aBounds(1)) + (cInt(aBounds(3))/2)  , "LEFT"
				Fn_SISW_Mech_AvailableParam_BreakdownOperations = True

			ElseIf sAction = "PopupMenuSelect" Then
				objNatTable.Click cInt(aBounds(0)-10) + (cInt(aBounds(2))/2),  cInt(aBounds(1)) + (cInt(aBounds(3))/2)  , "LEFT"
                wait 1
				objNatTable.Click cInt(aBounds(0)-51) + (cInt(aBounds(2))/2),  cInt(aBounds(1)) + (cInt(aBounds(3))/2)  , "RIGHT"
				wait 2
				Fn_SISW_Mech_AvailableParam_BreakdownOperations = Fn_UI_JavaMenu_Select("Fn_SISW_Mech_AvailableParam_BreakdownOperations",JavaWindow("Mechatronics"),sPopupMenu)

			ElseIf sAction = "SetCellData" Then
				'added to set focus on the current cell  -pratap 
				objNatTable.MouseDrag aBounds(0),aBounds(1),aBounds(2),aBounds(3)
				objNatTable.Click cInt(aBounds(0)-25) + (cInt(aBounds(2))/2),  cInt(aBounds(1)) + (cInt(aBounds(3))/2)  , "LEFT"
				wait 2
				objNatTable.Object.setFocus
				wait 1
				call Fn_KeyBoardOperation("SendKey"," ")
				wait 1
				objNatTable.Type trim(dicAvailableParameters("Value")) + vblf

				Fn_SISW_Mech_AvailableParam_BreakdownOperations = True

            Elseif sAction = "PopupMenuExist" then
				objNatTable.Click cInt(aBounds(0)-10) + (cInt(aBounds(2))/2),  cInt(aBounds(1)) + (cInt(aBounds(3))/2)  , "LEFT"
                wait 1
				objNatTable.Click cInt(aBounds(0)-51) + (cInt(aBounds(2))/2),  cInt(aBounds(1)) + (cInt(aBounds(3))/2)  , "RIGHT"
				wait 2
				aMenu=Split(sPopupMenu,":")
				Set objPopupMenu=JavaWindow("Mechatronics").JavaMenu("label:="&aMenu(0)&"","index:=0")

				For iCount=1 to ubound(aMenu)
					Set objPopupMenu=objPopupMenu.JavaMenu("label:="&aMenu(iCount))
				Next
				Fn_SISW_Mech_AvailableParam_BreakdownOperations=objPopupMenu.Exist(5)
				
				wait 1
				Set WshShell = CreateObject("WScript.Shell")
				wait 1
				WshShell.SendKeys "{ESC}"
				Set WshShell =Nothing
			End If

		Case "MultiSelect", "MultiPopupSelect", "MultiSelectAssign", "MultiSelectUnassign", "MultiSelectPopupMenuExist"
			arrObj = Split(dicAvailableParameters("Object"), "~")
    		Set dicparameter = dicAvailableParameters
			Set myDeviceReplay = CreateObject("Mercury.DeviceReplay")

			For iCount = 0 To Ubound(arrObj)
				dicparameter("Object") = arrObj(iCount)

				sCol = 0
				iRowCnt = Fn_SISW_Mech_AvailableParam_BreakdownOperations("GetRowIndex",dicparameter, "")
				If CInt(iRowCnt) = -1 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_Mech_AvailableParam_BreakdownOperations ] Row Not Found [ " & dicAvailableParameters("Object") & " ].")
					Exit Function
				End If

				sRectangle = cStr(objNatTable.Object.getCellByPosition(sCol,iRowCnt).getBounds().toString())
				sRectangle =  Right(sRectangle, (Len(sRectangle) - Instr(1, sRectangle, "{", 1)))	
				sRectangle = replace(sRectangle,"}","")
				sRectangle = replace(sRectangle," ","")
				aBounds = split(sRectangle,",")
				If iCount = 1 Then
					myDeviceReplay.KeyDown 29
				End If
				objNatTable.Click cInt(aBounds(0)) + (cInt(aBounds(2))/2),  cInt(aBounds(1)) + (cInt(aBounds(3))/2)  , "LEFT"
			Next
			Fn_SISW_Mech_AvailableParam_BreakdownOperations = True
            myDeviceReplay.KeyUp 29

			If sAction = "MultiPopupSelect" Then
				sCol = 1
				If dicAvailableParameters("TabName") = "Assigned Parameters" Then
					sCol = 2
				End If

				sRectangle = cStr(objNatTable.Object.getCellByPosition(sCol,iRowCnt).getBounds().toString())
				sRectangle =  Right(sRectangle, (Len(sRectangle) - Instr(1, sRectangle, "{", 1)))	
				sRectangle = replace(sRectangle,"}","")
				sRectangle = replace(sRectangle," ","")
				aBounds = split(sRectangle,",")

				objNatTable.Click cInt(aBounds(0)) + (cInt(aBounds(2))/2),  cInt(aBounds(1)) + (cInt(aBounds(3))/2)  , "RIGHT"
				wait 2
				Fn_SISW_Mech_AvailableParam_BreakdownOperations = Fn_UI_JavaMenu_Select("Fn_SISW_Mech_AvailableParam_BreakdownOperations",JavaWindow("Mechatronics"),sPopupMenu)

			ElseIf sAction = "MultiSelectAssign" Then
				Fn_SISW_Mech_AvailableParam_BreakdownOperations = Fn_Button_Click("Fn_SISW_Mech_AvailableParam_BreakdownOperations", JavaWindow("Mechatronics"),"Assign")
			ElseIf sAction = "MultiSelectUnassign" Then
				Fn_SISW_Mech_AvailableParam_BreakdownOperations = Fn_Button_Click("Fn_SISW_Mech_AvailableParam_BreakdownOperations", JavaWindow("Mechatronics"),"Unassign")
			Elseif sAction = "MultiSelectPopupMenuExist" then
                sCol = 1
				If dicAvailableParameters("TabName") = "Assigned Parameters" Then
					sCol = 2
				End If

				sRectangle = cStr(objNatTable.Object.getCellByPosition(sCol,iRowCnt).getBounds().toString())
				sRectangle =  Right(sRectangle, (Len(sRectangle) - Instr(1, sRectangle, "{", 1)))	
				sRectangle = replace(sRectangle,"}","")
				sRectangle = replace(sRectangle," ","")
				aBounds = split(sRectangle,",")

				objNatTable.Click cInt(aBounds(0)) + (cInt(aBounds(2))/2),  cInt(aBounds(1)) + (cInt(aBounds(3))/2)  , "RIGHT"
				wait 2

				aMenu=Split(sPopupMenu,":")
				Set objPopupMenu=JavaWindow("Mechatronics").JavaMenu("label:="&aMenu(0))
				For iCount=1 to ubound(aMenu)
					Set objPopupMenu=objPopupMenu.JavaMenu("label:="&aMenu(iCount))
				Next
				Fn_SISW_Mech_AvailableParam_BreakdownOperations=objPopupMenu.Exist(5)
				
				wait 1
				Set WshShell = CreateObject("WScript.Shell")
				wait 1
				WshShell.SendKeys "{ESC}"
				Set WshShell =Nothing
			End If

		Case "CellVerify"
			On Error Resume Next
			arrCol = Split(dicAvailableParameters("ColName"), "~")
			arrValues  = Split(dicAvailableParameters("Value"), "~")

			iRowCnt = Fn_SISW_Mech_AvailableParam_BreakdownOperations("GetRowIndex",dicAvailableParameters, "")
			If CInt(iRowCnt) = -1 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_Mech_AvailableParam_BreakdownOperations ] Row Not Found[ " & dicAvailableParameters("Object") & " ].")
				Exit Function
			End If

			Set dicparameter = dicAvailableParameters

			' Verify Values
			For iCount = 0 To Ubound(arrCol)

				dicparameter("ColName") = arrCol(iCount)
        				
				sCol = Fn_SISW_Mech_AvailableParam_BreakdownOperations("GetColumnIndex",dicparameter, "")
				If CInt(sCol) = -1 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_Mech_AvailableParam_BreakdownOperations ] Column Not Found[ " & dicAvailableParameters("ColName") & " ].")
					Exit Function
				End If
				
				sRectangle = cStr(objNatTable.Object.getCellByPosition(sCol,iRowCnt).getBounds().toString())
				sRectangle =  Right(sRectangle, (Len(sRectangle) - Instr(1, sRectangle, "{", 1)))	
				sRectangle = replace(sRectangle,"}","")
				sRectangle = replace(sRectangle," ","")
				aBounds = split(sRectangle,",")

				bReturn = ""
                bReturn = objNatTable.Object.getCellByPosition(sCol, iRowCnt).getDataValue().toString()
				If bReturn = empty Then
					bReturn = ""
				End If

				If trim(arrValues(iCount)) <> trim(bReturn) then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_Mech_AvailableParam_BreakdownOperations ] Value ["& arrValues(iCount) &"] not exist [ " & dicparameter("ColName") & " ].")
					Exit Function
				End If
			Next
			Fn_SISW_Mech_AvailableParam_BreakdownOperations = True
	' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "Ascending", "Descending"
			On Error Resume Next
			sCol = Fn_SISW_Mech_AvailableParam_BreakdownOperations("GetColumnIndex",dicAvailableParameters, "")

			iRowCnt = cInt(objNatTable.Object.getRowCount())
			bReturn = True	
			For iCnt = 3 to iRowCnt-2
				If iCnt = 3 Then
					startIndex = ""
					startIndex = objNatTable.Object.getCellByPosition(sCol, iCnt).getDataValue().toString()
					If startIndex = empty Then
						startIndex = ""
					End If
				Else
					startIndex = endIndex
				End If

				endIndex = ""
				endIndex = objNatTable.Object.getCellByPosition(sCol,Cint(iCnt+1)).getDataValue().toString()
				If endIndex = empty Then
					endIndex = ""
				End If

				' Check Action Ascending OR Descending
				If sAction = "Ascending" Then
						If endIndex < startIndex then
							bReturn = False
							Exit For
						End If
				ElseIf sAction = "Descending" Then
						If endIndex > startIndex then
							bFlag = False
							Exit For
						End If
				End If

			Next
			If bReturn <> False Then
				Fn_SISW_Mech_AvailableParam_BreakdownOperations = True
			End If

		Case "SelectRow"
        	sCol = 0
    		iRowCnt = Fn_SISW_Mech_AvailableParam_BreakdownOperations("GetRowIndex",dicAvailableParameters, "")
			If CInt(iRowCnt) = -1 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_Mech_AvailableParam_BreakdownOperations ] Row Not Found[ " & dicAvailableParameters("Object") & " ].")
				Exit Function
			End If

			sRectangle = cStr(objNatTable.Object.getCellByPosition(sCol,iRowCnt).getBounds().toString())
			sRectangle =  Right(sRectangle, (Len(sRectangle) - Instr(1, sRectangle, "{", 1)))	
			sRectangle = replace(sRectangle,"}","")
			sRectangle = replace(sRectangle," ","")
			aBounds = split(sRectangle,",")
			objNatTable.Click cInt(aBounds(0)) + (cInt(aBounds(2))/2),  cInt(aBounds(1)) + (cInt(aBounds(3))/2)  , "LEFT"
			Fn_SISW_Mech_AvailableParam_BreakdownOperations = True

		Case "Assign", "Unassign"
			bReturn =Fn_SISW_Mech_AvailableParam_BreakdownOperations("SelectCell", dicAvailableParameters,  "")
			If bReturn = False Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_Mech_AvailableParam_BreakdownOperations ] Row not found [ " & dicAvailableParameters("Object") & " ].")
				Exit Function
			End If

			If sAction = "Assign" Then
				Fn_SISW_Mech_AvailableParam_BreakdownOperations = Fn_Button_Click("Fn_SISW_Mech_AvailableParam_BreakdownOperations", JavaWindow("Mechatronics"),"Assign")
			ElseIf  sAction = "Unassign" Then
				Fn_SISW_Mech_AvailableParam_BreakdownOperations = Fn_Button_Click("Fn_SISW_Mech_AvailableParam_BreakdownOperations", JavaWindow("Mechatronics"),"Unassign")
			End If

' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "HeaderPopupMenuSelect",  "HeaderSelect"
			sCol = Fn_SISW_Mech_AvailableParam_BreakdownOperations("GetColumnIndex",dicAvailableParameters, "")
			
			If CInt(sCol) <> -1 Then
				sRectangle = cStr(objNatTable.Object.getCellByPosition(sCol,1).getBounds().toString())
				sRectangle =  Right(sRectangle, (Len(sRectangle) - Instr(1, sRectangle, "{", 1)))	
				sRectangle = replace(sRectangle,"}","")
				sRectangle = replace(sRectangle," ","")
				aBounds = split(sRectangle,",")

				If sAction = "HeaderPopupMenuSelect" Then
					objNatTable.Click cInt(aBounds(0)-10) + (cInt(aBounds(2))/2),  cInt(aBounds(1)) + (cInt(aBounds(3))/2)  , "RIGHT"
					wait 2
					Fn_SISW_Mech_AvailableParam_BreakdownOperations = Fn_UI_JavaMenu_Select("Fn_SISW_Mech_AvailableParam_BreakdownOperations",JavaWindow("Mechatronics"),sPopupMenu)
				ElseIf  sAction = "HeaderSelect" Then
					objNatTable.Click cInt(aBounds(0)-10) + (cInt(aBounds(2))/2),  cInt(aBounds(1)-10) + (cInt(aBounds(3))/2)  , "LEFT"
					Fn_SISW_Mech_AvailableParam_BreakdownOperations = TRUE
				End If
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_Mech_AvailableParam_BreakdownOperations ] Row not found [ " & dicAvailableParameters("Object") & " ].")
			End If
    ' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "GetColumnIndex"
			On Error Resume Next
			Fn_SISW_Mech_AvailableParam_BreakdownOperations = -1
			Call Fn_SISW_RAC_NatTable_Init(objNatTable)
        
			'Resetting Horizontal scroll to default position
			objNatTable.Object.getHorizontalBar().setSelection(0)

			iColCnt = 1
			Do until False
'			For iColCnt =1 To 5
'				If iColCnt = 2 Then
'					' take scroll to left
'					Call Fn_SISW_RAC_NatTable_Init(objNatTable) ' Initialise nat table Positions
'				ElseIf iColCnt = 3 OR iColCnt = 4 OR iColCnt = 5 Then
'					' Take scroll to right step by step
'					objNatTable.Click NT_HScrollBarX + NT_HScrollBarW - 20, NT_HScrollBarY + 5, "LEFT"
'				End If
				iCount = ""
				iCount = cInt(objNatTable.Object.getColumnCount)
				If iCount = "" Then
					iCount=objNatTable.Object.getClientAreaProvider.getLayer().getColumnHeaderLayer.getColumnHeaderlayerStack.getDataLayer.getColumnCount()
				End If
	
				startIndex = -1 
				endIndex =  -1
				If dicAvailableParameters("TableHeader")<>"" Then
					For iCnt = 1 To iCount
						If trim(dicAvailableParameters("TableHeader")) = trim(objNatTable.Object.getCellByPosition(iCnt,0).getDataValue().toString()) Then
							If startIndex = -1 Then
								startIndex = iCnt
							Else
								endIndex = iCnt
							End If
						End If
					Next
				Else
					startIndex = 1 
					endIndex = iCount
				End If
				'[TC12-20171001-18_10_2017JotibaT]-- Addded code to handle coloumn count & index
				For iCnt = startIndex to endIndex-1
					sColName = ""
					sColName = objNatTable.Object.getCellByPosition(iCnt,1).getDataValue().toString()
					If trim(sColName) = "" Then
						sColName = objNatTable.Object.getClientAreaProvider.getLayer().getColumnHeaderLayer.getColumnHeaderlayerStack.getDataLayer.getDataValue(iCnt,0).toString()
					End If

					If trim(dicAvailableParameters("ColName")) = trim(sColName) Then
						Fn_SISW_Mech_AvailableParam_BreakdownOperations = iCnt
						exit for
					End If
				Next
				If CInt(Fn_SISW_Mech_AvailableParam_BreakdownOperations) <> -1 OR cdbl(NT_objHScrollBar.getSelection()) + NT_HScrollBarThumb = NT_HScrollBarMax Then
					Exit Do
'					Exit For
				End If
				iColCnt = iColCnt + 1
				If iColCnt = 2 Then
'					objNatTable.Click NT_HScrollBarX + NT_HScrollBarW - 20, NT_HScrollBarY + 5, "LEFT"
					objNatTable.Click NT_HScrollBarX + 20, NT_HScrollBarY+5 , "LEFT"
					Do Until cdbl(NT_objHScrollBar.getSelection()) < NT_HScrollBarX+100
						objNatTable.Click NT_HScrollBarX + 20, NT_HScrollBarY+5 , "LEFT"
					Loop
				ElseIf iColCnt > 2 Then 
					Call Fn_SISW_RAC_NatTable_MoveScrollBar(objNatTable, "Right")
					Call Fn_SISW_RAC_NatTable_MoveScrollBar(objNatTable, "Right")
				End If

'				If Cdbl(NT_HScrollBarMax-500)<cdbl(NT_objHScrollBar.getSelection()) + NT_HScrollBarThumb Then
'					Call Fn_SISW_RAC_NatTable_MoveScrollBar(objNatTable, "Right")
'					wait 0,100
'					Call Fn_SISW_RAC_NatTable_MoveScrollBar(objNatTable, "Right")
'				End If
			loop
    ' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "GetRowIndex"
			Dim iInstance, aObject, sObject
			sCol = 1
			If  dicAvailableParameters("TabName") = "Assigned Parameters" Then
				sCol = 2
			End If
			iInstance = 1
			aObject = split(dicAvailableParameters("Object") ,"@")
			sObject = trim(aObject(0))
			If instr(dicAvailableParameters("Object") ,"@") > 0 Then
				iInstance = cInt(trim(aObject(1)))
			End If
			Fn_SISW_Mech_AvailableParam_BreakdownOperations = -1
			Call Fn_SISW_RAC_NatTable_Init(objNatTable)
			objNatTable.Object.getHorizontalBar().setSelection(0)
			Call Fn_SISW_RAC_NatTable_MoveScrollBar(objNatTable, "Left")
			wait 1
			Call Fn_SISW_RAC_NatTable_MoveScrollBar(objNatTable, "Left")
			wait 1
'			objNatTable.Click NT_HScrollBarX + 20, NT_HScrollBarY+5 , "LEFT"
'			objNatTable.Click NT_HScrollBarX + 20, NT_HScrollBarY+5 , "LEFT"
			objNatTable.Object.getHorizontalBar().setSelection(0)
			iRowCnt = cInt(objNatTable.Object.getRowCount())
			For iCnt = 2 to iRowCnt -1
				If sObject = objNatTable.Object.getCellByPosition(sCol,iCnt).getDataValue().toString() then
					If iInstance = 1 Then
						Fn_SISW_Mech_AvailableParam_BreakdownOperations = iCnt
						Exit For
					Else
						iInstance = iInstance - 1
					End If
				End If
			Next
	' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_Mech_AvailableParam_BreakdownOperations ] Invalid case [ " & sAction & " ].")
	' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
	End Select
	If  Fn_SISW_Mech_AvailableParam_BreakdownOperations <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [ Fn_SISW_Mech_AvailableParam_BreakdownOperations ] executed successfuly with case [ " & sAction & " ].")
	End If
	Set objNatTable = Nothing
End Function 

'''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'''/$$$$
'''/$$$$   FUNCTION NAME   :   Fn_SISW_Mechatronics_DefineAdditionalMemoryLayoutInfo(sAction,aFieldName,aValues,sInfo1,sInfo2)
'''/$$$$
'''/$$$$   DESCRIPTION        :  Perform Variour sOperations in the  Software Parameter manager Navigation tree
'''/$$$$
'''/$$$$    PARAMETERS      :   1.) sAction : Valid Action Name
'''/$$$$                                     		2.) aFieldName : Array of Fields Name
'''/$$$$											  3.) aValues  : Array of Values to be inserted
'''/$$$$											  4.) sInfo1	  :  For future use in case any new functionality is added
'''/$$$$											 5.) sInfo2	  :  For future use in case any new functionality is added
'''/$$$$
'''/$$$$
'''/$$$$
'''/$$$$    Function Calls       :   Fn_WriteLogFile(),   Fn_SISW_Mechatronics_SPMNavigationTree()
'''/$$$$									  
'''/$$$$
'''/$$$$	 HISTORY           :   AUTHOR                 DATE        VERSION
'''/$$$$
'''/$$$$    CREATED BY     :   SHREYAS           31/10/2012         1.0
'''/$$$$
'''/$$$$    REVIWED BY     :   Shreyas
'''/$$$$
'''/$$$$
'''/$$$$   EXAMPLE          : 	bReturn=Fn_SISW_Mechatronics_DefineAdditionalMemoryLayoutInfo("EnterValues","Attribute~MemoryType~UserData_3","a~b~c","Finish","","")
'''/$$$$									bReturn=Fn_SISW_Mechatronics_DefineAdditionalMemoryLayoutInfo("VerifyValues","Attribute~MemoryType~UserData_3","a~b~c","Finish","","")
'''/$$$$ 										
'''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'
Public Function Fn_SISW_Mechatronics_DefineAdditionalMemoryLayoutInfo(sAction,aFieldName,aValues,sButton,sInfo1,sInfo2)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_Mechatronics_DefineAdditionalMemoryLayoutInfo"
Dim aFields,aValue,objLayout,Buttons
Set objLayout=Fn_SISW_Mechatronics_GetObject("NewMemoryLayout")
On error resume next

	Select Case sAction

		Case "EnterValues"
		
			aFields=split(aFieldName,"~",-1,1)
			aValue=split(aValues,"~",-1,1)
			   For iCount=0 to uBound(aFields)
					For jCount=0 to iCount
						If objLayout.JavaEdit(aFields(iCount)).Exist(5) then
							If err.number<0 Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"The Edit Box  ["+aFields(iCount)+"] Does Not Exist")
									Fn_SISW_Mechatronics_DefineAdditionalMemoryLayoutInfo=False
									Exit Function
							End If
							objLayout.JavaEdit(aFields(iCount)).Set aValue(iCount)
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Successfully Set the Value ["+aValue(iCount)+"] in the Edit Box  ["+aFields(iCount)+"] Does Not Exist")
							Fn_SISW_Mechatronics_DefineAdditionalMemoryLayoutInfo=True
							jCount=iCount
							Exit for
						End if
					Next
				Next

				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Completed Function [Fn_SISW_Mechatronics_DefineAdditionalMemoryLayoutInfo]")
				Fn_SISW_Mechatronics_DefineAdditionalMemoryLayoutInfo=True
		
		Case "VerifyValues"
		
			bFlag=false
			aFields=split(aFieldName,"~",-1,1)
			aValue=split(aValues,"~",-1,1)
			   For iCount=0 to uBound(aFields)
						For jCount=0 to iCount
							If objLayout.JavaEdit(aFields(iCount)).Exist(5) then
									sValue=objLayout.JavaEdit(aFields(iCount)).GetROProperty("value")
									If lCase(sValue)=lCase(aValue(iCount)) Then
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Successfully Verified the Value ["+aValue(iCount)+"] in the Edit Box  ["+aFields(iCount)+"] Does Not Exist")
											jCount=iCount
											Exit for
									Else
											Fn_SISW_Mechatronics_DefineAdditionalMemoryLayoutInfo=False
											Exit Function
									End If
							Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"The Edit Box  ["+aFields(iCount)+"] Does Not Exist")
                                Fn_SISW_Mechatronics_DefineAdditionalMemoryLayoutInfo=False
								Exit Function
							End if
						Next
				Next
				
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Completed Function [Fn_SISW_Mechatronics_DefineAdditionalMemoryLayoutInfo]")

	End Select

	If  sButton <> "" Then
		Call Fn_Button_Click("Fn_SISW_Mechatronics_DefineAdditionalMemoryLayoutInfo", objLayout,sButton)	
		If sButton = "Finish" Then
			Call Fn_Button_Click("Fn_SISW_Mechatronics_DefineAdditionalMemoryLayoutInfo", objLayout,"Close")
		End If
	End If
	Fn_SISW_Mechatronics_DefineAdditionalMemoryLayoutInfo=True
	Set objLayout=nothing
End Function

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name			:	Fn_SISW_Mech_OverrideConversionRule
'@@
'@@    Description				:	Function Used to perform operations on  Delete Override Dialog
'@@
'@@    Parameters			   	:	1. sAction	: 	Action To Perform Copy/ paste / Clear
'@@
'@@												2. Dim dicOverrideConversionRule
'@@												Set dicOverrideConversionRule = CreateObject("Scripting.Dictionary")
'@@												With dicOverrideConversionRule
'@@													.Add  "Action",""					' -  For Future use
'@@													.Add "Name",""							' Con Rule Name
'@@													.Add "Description",""					' Con Rule Desc
'@@													.Add "Type",""								Con Rule Type  	'Linear , Identical ,  Quadratic , Rational
'@@													.Add "Expression",""  				'- To Verify Expression After entring Constant Names and Values
'@@													.Add "ConstantsName",""			'  Constants names Seperated by ~
'@@													.Add "ConstantsValue",""		' 	Constants Values Sepereated by ~	
'@@												End With
'@@	
'@@							 	:			 3. sButton 	: Button Name  OK, Cancel 
'@@
'@@    Return Value		   	   	: 	True Or False
'@@
'@@    Pre-requisite			:	Override Conversion Rule Dialog should be open					
'@@
'@@    Examples					:	
'@@    									1.	Call Fn_SISW_Mech_OverrideConversionRule("Create", dicOverrideConversionRule,"OK")
'@@    									2.  Call Fn_SISW_Mech_OverrideConversionRule("Verify", dicOverrideConversionRule, "OK")
'@@    									
'@@    									3.   VerifyErrorMsg
'@@    									dicOverrideConversionRule("ShortErrMsg") = 		To Verify Short Error Msg
'@@    									dicOverrideConversionRule("DetailErrMsg") = 	To Verify Detail Error Message
'@@    									
'@@    						Call Fn_SISW_Mech_OverrideConversionRule("Verify", VerifyErrorMsg, "")   ' Send sButton only if you want close OverrideConversionRule dialog
'@@	History				:	
'@@			Developer Name			Date			 Rev. No.	   Changes Done		 Reviewer
'@@------------------------------------------------------------------------------------------------------------------------------
'@@			Pranav Ingle			02-Nov-2012		  1.0			created
'@@------------------------------------------------------------------------------------------------------------------------------
'@@			Pranav Ingle			26-Nov-2012		  1.1			Added Case ErrMsg
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_SISW_Mech_OverrideConversionRule(sAction, dicOverrideConversionRule, sButton)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_Mech_OverrideConversionRule"
	Dim objOverrideConversionRule, bFlag
	Dim arrConstantName, arrValue, iCounter, iCount, iRowCount, crrConstantValue, crrConstantName
    
	Set objOverrideConversionRule = JavaWindow("Mechatronics").JavaWindow("OverrideConversionRule")
	Fn_SISW_Mech_OverrideConversionRule = False
	' add call to select Traceability Matrix Tab
    
'	If Fn_UI_ObjectExist("Fn_SISW_Mech_OverrideConversionRule",objOverrideConversionRule ) = False  Then
	If Not(objOverrideConversionRule.Exist(5)) Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_Mech_OverrideConversionRule ] Nat Table object of [ Traceability Matrix ] is not visible.")
			Exit function
	End If
'	End If

	Select Case sAction
		Case "Create"
'			If  dicOverrideConversionRule("Action") <> "" Then   ' For Future use
'			End If

			If dicOverrideConversionRule("Name") <> "" Then
				Call Fn_Edit_Box("Fn_SISW_Mech_OverrideConversionRule",objOverrideConversionRule,"Name",dicOverrideConversionRule("Name"))
			End If

			If  dicOverrideConversionRule("Description") <> "" Then
				Call Fn_Edit_Box("Fn_SISW_Mech_OverrideConversionRule",objOverrideConversionRule,"Description",dicOverrideConversionRule("Description"))
			End If

			If  dicOverrideConversionRule("Type") <> "" Then
				Call Fn_List_Select("Fn_SISW_Mech_OverrideConversionRule", objOverrideConversionRule,"Type",dicOverrideConversionRule("Type"))
			End If
			
			If  dicOverrideConversionRule("ConstantsName") <> "" And dicOverrideConversionRule("ConstantsValue") <> "" Then
				
				arrConstantName=Split(dicOverrideConversionRule("ConstantsName"),"~")
				arrValue=Split(dicOverrideConversionRule("ConstantsValue"),"~")

				For iCounter=0 to ubound(arrConstantName)
					iRowCount = objOverrideConversionRule.JavaTable("Constants").GetROProperty("rows")

					For iCount=0 to iRowCount-1
						bFlag = false
						crrConstantName=objOverrideConversionRule.JavaTable("Constants").GetCellData(iCount,"Constant Name")
						If trim(crrConstantName)=trim(arrConstantName(iCounter)) Then
								objOverrideConversionRule.JavaTable("Constants").ClickCell iCount,"Constant Value"
		'						objOverrideConversionRule.JavaTable("Constants").SetCellData iCount,"Constant Value", arrValue(iCounter)
								objOverrideConversionRule.JavaTable("Constants").Type arrValue(iCounter)
								wait 1
								Set WshShell = CreateObject("WScript.Shell")
								WshShell.SendKeys "{ENTER}"
								Set WshShell = nothing
								wait 1
	
								bFlag=true
								Exit for
						End If
					Next
					If bFlag=false Then
						set objOverrideConversionRule=nothing
						Exit function
					End If
				Next
			End If
		
			Fn_SISW_Mech_OverrideConversionRule = True

		Case "Verify"

			If dicOverrideConversionRule("Name") <> "" Then
				If dicOverrideConversionRule("Name") <> Fn_Edit_Box_GetValue("Fn_SISW_Mech_OverrideConversionRule",objOverrideConversionRule,"Name") Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_Mech_OverrideConversionRule ] Failed To Verify [ " & dicOverrideConversionRule("Name") & " ].")
					Exit function
				End If
			End If

			If  dicOverrideConversionRule("Description") <> "" Then
				If dicOverrideConversionRule("Description") <> Fn_Edit_Box_GetValue("Fn_SISW_Mech_OverrideConversionRule",objOverrideConversionRule,"Description") Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_Mech_OverrideConversionRule ] Failed To Verify [ " & dicOverrideConversionRule("Description") & " ].")
					Exit function
				End If
			End If

			If  dicOverrideConversionRule("Type") <> "" Then
				If dicOverrideConversionRule("Type") <> objOverrideConversionRule.JavaList("Type").GetROProperty("value") Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_Mech_OverrideConversionRule ] Failed To Verify [ " & dicOverrideConversionRule("Type") & " ].")
					Exit function
				End If
			End If

			If  dicOverrideConversionRule("Expression") <> "" Then
				If dicOverrideConversionRule("Expression") <> Fn_Edit_Box_GetValue("Fn_SISW_Mech_OverrideConversionRule",objOverrideConversionRule,"Expression") Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_Mech_OverrideConversionRule ] Failed To Verify [ " & dicOverrideConversionRule("Expression") & " ].")
					Exit function
				End If
			End If	

			If  dicOverrideConversionRule("ConstantsName") <> "" And dicOverrideConversionRule("ConstantsValue") <> "" Then
				
				arrConstantName=Split(dicOverrideConversionRule("ConstantsName"),"~")
				arrValue=Split(dicOverrideConversionRule("ConstantsValue"),"~")

				For iCounter=0 to ubound(arrConstantName)
					iRowCount = objOverrideConversionRule.JavaTable("Constants").GetROProperty("rows")

					For iCount=0 to iRowCount-1
						bFlag=false
						crrConstantName	= objOverrideConversionRule.JavaTable("Constants").GetCellData(iCount,"Constant Name")
						If trim(crrConstantName)=trim(arrConstantName(iCounter)) Then
							crrConstantValue = objOverrideConversionRule.JavaTable("Constants").GetCellData(iCount,"Constant Value")
							If trim(crrConstantValue) = trim(arrValue(iCounter)) Then
								bFlag=true
								Exit for
							End If
						End If
					Next
					If bFlag=false Then
						set objOverrideConversionRule=nothing
						Exit function
					End If
				Next
			End If

			Fn_SISW_Mech_OverrideConversionRule = True
	' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "VerifyErrorMsg"

				If  dicOverrideConversionRule("ShortErrMsg") <> "" Then
					objOverrideConversionRule.JavaWindow("Error").JavaStaticText("ShortErrMsg").SetTOProperty "label", dicOverrideConversionRule("ShortErrMsg")
					If  objOverrideConversionRule.JavaWindow("Error").JavaStaticText("ShortErrMsg").Exist(5) Then
						Fn_SISW_Mech_OverrideConversionRule = True
					Else
						Fn_SISW_Mech_OverrideConversionRule = False
					End If
				End If

				If  dicOverrideConversionRule("DetailErrMsg") <> "" Then
					sDispMsg = objOverrideConversionRule.JavaWindow("Error").JavaEdit("ErrMsg").GetROProperty("value")
					If  Instr(1,sDispMsg,sMessage) > 0 Then
						Fn_SISW_Mech_OverrideConversionRule = True
					Else
						Fn_SISW_Mech_OverrideConversionRule = False
					End If
				End If

				Call Fn_Button_Click("Fn_SISW_Mech_OverrideConversionRule", objOverrideConversionRule.JavaWindow("Error"),"OK")
	' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_Mech_OverrideConversionRule ] Invalid case [ " & sAction & " ].")
	' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
	End Select

	If  Fn_SISW_Mech_OverrideConversionRule <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [ Fn_SISW_Mech_OverrideConversionRule ] executed successfuly with case [ " & sAction & " ].")
	End If

	If sButton <> "" Then
		Call Fn_Button_Click("Fn_SISW_Mech_OverrideConversionRule", objOverrideConversionRule,sButton)
	End If

	Set objOverrideConversionRule = Nothing
End Function

''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''''/$$$$
''''/$$$$   FUNCTION NAME   :   Fn_SISW_Mechatronics_MemoryLayoutBasicCreate(sPrimaryNode,StrItemType,StrConfItem,StrItemID,StrItemRevID,StrItemName,StrItemDesc,StrItemUOM,sInfo1,sInfo2)
''''/$$$$
''''/$$$$   DESCRIPTION        :  Create Basic Memory Layout
''''/$$$$
''''/$$$$    PARAMETERS      :   1.) sPrimaryNode : Valid Primary Node Name
''''/$$$$                                     		2.) StrItemType : Type Of Memory Layout
''''/$$$$											3.) StrConfItem  : Parameter to be set as On Or Off to check or uncheck "Configuration checkbox"
''''/$$$$											4.) StrItemID:      Memory Layout ID to be inserted
''''/$$$$											5.) StrItemRevID   : Memory Layout Revision ID to be Inserted
''''/$$$$											6.) StrItemName  :  Memory Layout Name
''''/$$$$											7.) StrItemDesc  :  Memory Layout Description
''''/$$$$											8.)	StrItemUOM
''''/$$$$											 9.) sInfo1	  :  For future use in case any new functionality is added
''''/$$$$											10.) sInfo2	  :  For future use in case any new functionality is added
''''/$$$$
''''/$$$$
''''/$$$$
''''/$$$$    Function Calls       :   Fn_WriteLogFile(),   Fn_SISW_Mechatronics_SPMNavigationTree()
''''/$$$$									  
''''/$$$$
''''/$$$$	 HISTORY           :   AUTHOR                 DATE        VERSION
''''/$$$$
''''/$$$$    CREATED BY     :   SHREYAS           02/11/2012         1.0
''''/$$$$
''''/$$$$    REVIWED BY     :   Shreyas
''''/$$$$
''''/$$$$
''''/$$$$   EXAMPLE          : 	bReturn= Fn_SISW_Mechatronics_MemoryLayoutBasicCreate("A-Breakdown","Parameter Memory Layout","","","","New Layout","Product","","","")
''''/$$$$ 										
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

Public Function Fn_SISW_Mechatronics_MemoryLayoutBasicCreate(sPrimaryNode,StrItemType,StrConfItem,StrItemID,StrItemRevID,StrItemName,StrItemDesc,StrItemUOM,sButton,sInfo2)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_Mechatronics_MemoryLayoutBasicCreate"
	Dim sItemId, sRevId
	Dim objLayout
	Dim hieght, width

	'Select menu [File -> New -> Item...]
	If Fn_UI_ObjectExist("Fn_SISW_Mechatronics_MemoryLayoutBasicCreate",Window("MechatronicsWindow").JavaDialog("NewMemoryLayout"))=False Then
		bReturn=Fn_SISW_Mechatronics_SPMNavigationTree("PopupMenuSelect",sPrimaryNode+":Memory Layouts","Add:Memory Layout...","","")
       Call  Fn_ReadyStatusSync(3)
	End If
	
	'Check the existence of "New Item " window
	Set objLayout = Fn_SISW_Mechatronics_GetObject("NewMemoryLayout")

		'Select Item Type
		Call Fn_List_Select("Fn_SISW_Mechatronics_MemoryLayoutBasicCreate", objLayout,"ItemType",StrItemType)
		'checked Configuration item or not
		If StrConfItem <> "" Then
		 Call Fn_CheckBox_Set("Fn_SISW_Mechatronics_MemoryLayoutBasicCreate", objLayout,"Configuration Item",StrConfItem)
		End If

	  Call  Fn_ReadyStatusSync(3)
		'Click on "Next" button
		objLayout.JavaButton("Next").Click micLeftBtn
	  ' Call Fn_Button_Click("Fn_SISW_Mechatronics_MemoryLayoutBasicCreate", objLayout,"Next")

		If StrItemID <> "" Then
			'Set  Item Id
			 Call Fn_Edit_Box("Fn_SISW_Mechatronics_MemoryLayoutBasicCreate",objLayout,"ItemID", StrItemID)
		End If

		If StrItemRevID <> "" Then
			'Set Revision ID
			Call Fn_Edit_Box("Fn_SISW_Mechatronics_MemoryLayoutBasicCreate",objLayout,"RevisionID", StrItemRevID)
		End If

		If  StrItemID = "" or StrItemRevID = "" Then
			'click on assign button
			If Not objLayout.JavaButton("Assign").GetROProperty("enabled")="0" Then
				Call Fn_Button_Click("Fn_SISW_Mechatronics_MemoryLayoutBasicCreate", objLayout, "Assign")
			End If
		End If

		wait(3)

		'Extract Creation data
		sItemId = Fn_Edit_Box_GetValue("Fn_SISW_Mechatronics_MemoryLayoutBasicCreate", objLayout,"ItemID")
		sRevId = Fn_Edit_Box_GetValue("Fn_SISW_Mechatronics_MemoryLayoutBasicCreate", objLayout,"RevisionID")
		
		'*****************************************************************
		If  sItemId = "" or sRevId = "" Then
			'click on assign button
			Call Fn_UpdateLogFiles(Time() & " - " & "WARNING - Assign button need to click again.", "")
			call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Item ID not shown in ItemId Textbox[" + CStr(sItemId) + "]")
			Call Fn_Button_Click("Fn_SISW_Mechatronics_MemoryLayoutBasicCreate", objLayout, "Assign")
			sItemId = Fn_Edit_Box_GetValue("Fn_SISW_Mechatronics_MemoryLayoutBasicCreate", objLayout,"ItemID")
			sRevId = Fn_Edit_Box_GetValue("Fn_SISW_Mechatronics_MemoryLayoutBasicCreate", objLayout,"RevisionID")
		End If
		'*****************************************************************
		
		'Set Item name
		 Call Fn_Edit_Box("Fn_SISW_Mechatronics_MemoryLayoutBasicCreate", objLayout,"ItemName",StrItemName)
		'Set description
		If StrItemDesc <> "" Then
			Call Fn_Edit_Box("Fn_SISW_Mechatronics_MemoryLayoutBasicCreate", objLayout,"Description",StrItemDesc)
		End If
		'Set UOM
		If StrItemUOM <> "" Then
		  Call Fn_Edit_Box("Fn_SISW_Mechatronics_MemoryLayoutBasicCreate", objLayout,"Unit of Measure",StrItemUOM)
		End If

		If  sButton <> "" Then
			Call Fn_Button_Click("Fn_SISW_Mechatronics_MemoryLayoutBasicCreate", objLayout,sButton)	
			If lcase(sButton)="next" Then
				wait 2
				' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
				'Resizing window
				hieght=JavaWindow("Mechatronics").GetROProperty("height")
				width=JavaWindow("Mechatronics").GetROProperty("width")
				objLayout.Move 0,0
				wait 2
				objLayout.Resize width-5,hieght-5
				wait 2
				' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
			End If
		
			If sButton = "Finish" Then
				Call Fn_Button_Click("Fn_SISW_Mechatronics_MemoryLayoutBasicCreate", objLayout,"Close")
			End If
		End If
		'= = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = 
		Fn_SISW_Mechatronics_MemoryLayoutBasicCreate = "'"&sItemId & "-" & sRevId
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Created an Item of ID [" + CStr(sItemId) + "]")

		Set objLayout=Nothing
End Function

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name			:	Fn_SISW_Mech_ColumnChooser
'@@
'@@    Description				:	Function Used to perform operations on  Delete Override Dialog
'@@
'@@    Parameters			   	:	1. sAction	: 	Action To Perform Copy/ paste / Clear
'@@ 										  2. sColPath
'@@												
'@@    Return Value		   	   	: 	True Or False
'@@
'@@    Pre-requisite			:	Column Chooser Dialog should be open					
'@@
'@@    Examples					:	
'@@    									1.	Call Fn_SISW_Mech_ColumnChooser("Add", "All:Memory Layout/Memory Block:Header Object")
'@@	History				:	
'@@	Developer Name			Date			 Rev. No.		 Reviewer			Changes Done
'@@------------------------------------------------------------------------------------------------------------------------------
'@@	Pranav Ingle			02-Nov-2012		  1.0			Self				created
'@@------------------------------------------------------------------------------------------------------------------------------
'@@	Swati Kuntullu			30-Nov-2012		  1.0			Koustubh Watwe		Added case Remove		
'@@------------------------------------------------------------------------------------------------------------------------------
'@@	Sandeep N			06-Dec-2012		  1.1			modified case : Add :- added code to add multiple column
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_SISW_Mech_ColumnChooser(sAction, sColumnPath, sInfo1)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_Mech_ColumnChooser"
	Dim objColumnChooser, sPath, arrColPath, iCount, bReturn
	Dim iCounter,arrPaths

	Set objColumnChooser = JavaWindow("Mechatronics").JavaWindow("ColumnChooser")
	Fn_SISW_Mech_ColumnChooser = False

	If Fn_UI_ObjectExist("Fn_SISW_Mech_ColumnChooser",objColumnChooser ) = False  Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_Mech_ColumnChooser ] Nat Table object of [ ColumnChooser ] is not visible.")
		Exit function
	End If

	Select Case sAction
		Case "Add"
			
			sPath = ""
			arrPaths=Split(sColumnPath,"~")

			For iCounter=0 to ubound(arrPaths)
				arrColPath = Split(arrPaths(iCounter),":")
				For iCount = 0 To  Ubound(arrColPath) - 1
					If iCount = 0 Then
						sPath = arrColPath(iCount)
					Else
						sPath = sPath &":"& arrColPath(iCount)
					End If
					Call Fn_UI_JavaTree_Expand("Fn_SISW_Mech_ColumnChooser",objColumnChooser,"AvailableColumns",sPath)
				Next
				bReturn = Fn_UI_JavaTree_NodeExist("Fn_SISW_Mech_ColumnChooser", objColumnChooser.JavaTree("AvailableColumns"),arrPaths(iCounter))
				If bReturn = True Then
					Call Fn_JavaTree_Select("Fn_SISW_Mech_ColumnChooser", objColumnChooser, "AvailableColumns", arrPaths(iCounter))
					wait 2
					Call Fn_Button_Click("Fn_SISW_Mech_ColumnChooser", objColumnChooser,"Right")
					wait 1
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [ Fn_SISW_Mech_ColumnChooser ] " & arrPaths(iCounter) & " Already exist in Selected Colimns.")
				End If
			Next			

			Call Fn_Button_Click("Fn_SISW_Mech_ColumnChooser", objColumnChooser,"Done")

			Fn_SISW_Mech_ColumnChooser = True
	' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 		
		Case "Remove"
			arrColPath = Split(sColumnPath,"~")
			For  iCount = 0 To  Ubound(arrColPath)
				bReturn = Fn_UI_ListItemExist("Fn_SISW_Mech_ColumnChooser", objColumnChooser,"SelectedColumns",arrColPath(iCount))
				If bReturn = True Then
					bReturn = Fn_List_Select("Fn_SISW_Mech_ColumnChooser", objColumnChooser, "SelectedColumns",arrColPath(iCount))
					If bReturn = True Then
						wait 2
						Call Fn_Button_Click("Fn_SISW_Mech_ColumnChooser", objColumnChooser,"Left")
						wait 1
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_Mech_ColumnChooser ] Failed to select " & arrColPath(iCount) & ".")
						Exit For
					End If
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_Mech_ColumnChooser ] " & arrColPath(iCount) & " Does Not exist in Selected Columns List.")
					Exit For
				End If
			Next
			Call Fn_Button_Click("Fn_SISW_Mech_ColumnChooser", objColumnChooser,"Done")
			Fn_SISW_Mech_ColumnChooser = bReturn
	' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_Mech_ColumnChooser ] Invalid case [ " & sAction & " ].")
	' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
	End Select
	If  Fn_SISW_Mech_ColumnChooser <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [ Fn_SISW_Mech_ColumnChooser ] executed successfuly with case [ " & sAction & " ].")
	End If
	Set objColumnChooser = Nothing
End Function

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name			:	Fn_SISW_Mechatronics_DefineMemoryLayoutRevisionInfo
'@@
'@@    Description				:	Function Used to perform operations on  Delete Override Dialog
'@@
'@@    Parameters			   	:	  1. sAction : Valid Action Name
'@@             		                     	2. aFieldName : Array of Fields Name
'@@    											3. aValues  : Array of Values to be inserted
'@@    											4. sInfo1	  :  For future use in case any new functionality is added
'@@   											5. sInfo2	  :  For future use in case any new functionality is added
'@@												
'@@    Return Value		   	   	: 	True Or False
'@@
'@@    Pre-requisite			:	MemoryLayoutRevisionInfo  panel should be open					
'@@
'@@    Examples					:	
'@@    									1.	Call Fn_SISW_Mechatronics_DefineMemoryLayoutRevisionInfo("EnterValues","Header Object~MirroredOffset~Size~StartAddress","Paste~20;30~2~10","Finish","","")
'@@    									2.	Call Fn_SISW_Mechatronics_DefineMemoryLayoutRevisionInfo("Verify","Header Object~MirroredOffset~Size~StartAddress","000123/A;1-Layout1~20;30~2~10","Finish","","")
'@@	History				:	
'@@	Developer Name			Date			 Rev. No.	   Changes Done		 Reviewer
'@@------------------------------------------------------------------------------------------------------------------------------
'@@	Pranav Ingle			05-Nov-2012		  1.0			created
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

Public Function Fn_SISW_Mechatronics_DefineMemoryLayoutRevisionInfo(sAction,aFieldName,aValues,sButton,sInfo1,sInfo2)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_Mechatronics_DefineMemoryLayoutRevisionInfo"
	Dim objPropetiesDialog
	Dim arrFields,arrValue,objLayout,Buttons
	Dim arrOffsetValues, sValue
	Dim iCount, jCounter, bFlag
	Dim iEleCount, jCount
	Set objPropetiesDialog=Window("MechatronicsWindow").JavaDialog("NewMemoryLayout")
	Set objLayout= Fn_SISW_Mechatronics_GetObject("NewMemoryLayout")
	Fn_SISW_Mechatronics_DefineMemoryLayoutRevisionInfo = False

	Select Case sAction
   
		Case "EnterValues"
		
			arrFields=split(aFieldName,"~",-1,1)
			arrValue=split(aValues,"~",-1,1)
			For iCount=0 to uBound(arrFields)
	
				Select Case arrFields(iCount)
					Case "Header Object", "Trailer Object"
						objLayout.JavaStaticText("HeaderObjectLabel").SetTOProperty "label", arrFields(iCount) & ":"
						If objPropetiesDialog.JavaStaticText("HeaderObjectValue").Exist(2) Then
							objPropetiesDialog.JavaStaticText("HeaderObjectValue").Click 1, 1
						Else
							objPropetiesDialog.JavaObject("HeaderObject").Click 1,1	
						End If
						wait 1
	
						objLayout.JavaMenu("label:=" & arrValue(iCount)).Click 1,1
							
					Case "Size", "StartAddress", "UserData_1", "UserData_2", "UserData_3", "ProjectID", "RevisionID", "SerialNumber", "ItemComment"
						Call Fn_Edit_Box("Fn_SISW_Mechatronics_DefineMemoryLayoutRevisionInfo",objLayout,arrFields(iCount),arrValue(iCount))
	
					Case "MirroredOffset"
						arrOffsetValues=Split(arrValue(iCount),";")
						Call Fn_CheckBox_Set("Fn_SISW_Mechatronics_DefineMemoryLayoutRevisionInfo", objLayout, "AddMirroredOffset","on")
						For jCounter=0 to ubound(arrOffsetValues)
							Call Fn_Edit_Box("Fn_SISW_Mechatronics_DefineMemoryLayoutRevisionInfo",objLayout,"AddMirrorOffset",arrOffsetValues(jCounter))
							Call Fn_Button_Click("Fn_SISW_Mechatronics_DefineMemoryLayoutRevisionInfo", objLayout, "AddMirrorOffset")
						Next
						Call Fn_CheckBox_Set("Fn_SISW_Mechatronics_DefineMemoryLayoutRevisionInfo", objLayout, "AddMirroredOffset","off")
				End Select
			Next
			Fn_SISW_Mechatronics_DefineMemoryLayoutRevisionInfo=True

		Case "Verify"
		
			arrFields=split(aFieldName,"~",-1,1)
			arrValue=split(aValues,"~",-1,1)
			For iCount=0 to uBound(arrFields)
	
				Select Case arrFields(iCount)
					Case "Header Object", "Trailer Object"
						objLayout.JavaStaticText("HeaderObjectLabel").SetTOProperty "label", arrFields(iCount) & ":"
						sValue = objLayout.JavaStaticText("HeaderObjectValue").GetROProperty("attached text")
						If  sValue <> arrValue(iCount) Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_Mechatronics_DefineMemoryLayoutRevisionInfo ] Failed To Verify [ " & arrFields(iCount) & " has value " & arrValue(iCount) & " ].")
							Exit function
						End If
    							
					Case "Size", "StartAddress", "UserData_1", "UserData_2", "UserData_3", "ProjectID", "RevisionID", "SerialNumber", "ItemComment"
						sValue = objLayout.JavaEdit(arrFields(iCount)).GetROProperty("value")
						If  sValue <> arrValue(iCount) Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_Mechatronics_DefineMemoryLayoutRevisionInfo ] Failed To Verify [ " & arrFields(iCount) & " has value " & arrValue(iCount) & " ].")
							Exit function
						End If
					
					Case "MirroredOffset"

						arrOffsetValues=Split(arrValue(iCount),";")
						iEleCount=Fn_UI_Object_GetROProperty("Fn_SISW_Mechatronics_DefineMemoryLayoutRevisionInfo", objLayout.JavaList("MirroredOffset"), "items count")
						
						For jCounter=0 to ubound(arrOffsetValues)
							bFlag = False
							For jCount=0 to iEleCount-1
								If Cstr(objLayout.JavaList("MirroredOffset").GetItem(jCount)) = Cstr(arrOffsetValues(jCounter)) Then
									bFlag=true
									Exit for
								End If
							Next
							If bFlag = False Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_Mechatronics_DefineMemoryLayoutRevisionInfo ] Failed To Verify [ " & arrFields(iCount) & " has value " & arrValue(iCount) & " ].")
								Exit function
							End If
						Next
						Call Fn_CheckBox_Set("Fn_SISW_Mechatronics_DefineMemoryLayoutRevisionInfo", objLayout, "AddMirroredOffset","off")
				End Select
			Next
			Fn_SISW_Mechatronics_DefineMemoryLayoutRevisionInfo=True

	End Select
         
	If sButton <> "" Then
		Call Fn_Button_Click("Fn_SISW_Mech_OverrideConversionRule", objLayout,sButton)	
		If sButton = "Finish" Then
			Call Fn_Button_Click("Fn_SISW_Mech_OverrideConversionRule", objLayout,"Close")
		End If
	End If
	 Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Completed Function [Fn_SISW_Mechatronics_DefineAdditionalMemoryLayoutInfo]")
  
   Set objLayout=nothing
End Function

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name			:	Fn_SISW_Mech_AvailableParam_DictionaryOperations
'@@
'@@    Description				:	Function Used to perform operations on  Show Traceability Matrix window in System Engineering
'@@
'@@    Parameters			   	:	1. sAction	: 		Action [Type of Attribute Group]
'@@												2. Dim dicAvailableParameters
'@@														Set dicAvailableParameters = CreateObject("Scripting.Dictionary")
'@@														With dicAvailableParameters
'@@															.Add  "TabName",""
'@@															.Add "Object",""
'@@															.Add "ColName",""
'@@															.Add "Value",""
'@@															.Add "ShowAllAssignedParameters",""							   --  	Assigned Parameters  Tab		
'@@														End With
'@@	
'@@								:			  		dicAvailableParameters("TabName")	:     Available Parameters / Assigned Parameters
'@@								
'@@													 dicAvailableParameters("Object")			:  Object Column Row text [for Cases : Select, SelectRow, Assign ]
'@@								:			  																				ColName Row  [for Cases : CellVerify ] 
'@@							 	:			  	  dicAvailableParameters("ColName") 	  : Column Name        ----  [ Default Will be Object ]  -----
'@@							 	:			  	  dicAvailableParameters("Value") 			: 	Value To Verify 
'@@							 	:			 	  dicAvailableParameters("ShowAllAssignedParameters") 	: Show All Assigned Parameters Check Box Set [true: checked, false : Unchecked]
'@@							 	
'@@							 	:			 3. sPopupMenu 	: Popup Menu select
'@@
'@@    Return Value		   	   	: 	True Or False
'@@
'@@    Pre-requisite			:	Software Parameter Manager perspective should be activated and Available parameter Tab is Selected.						
'@@
'@@    Examples					:	
'@@										Case "SelectCell" , "Assign", "Unassign", "PopupMenuSelect"  , "DoubleClick"
'@@											dicAvailableParameters.RemoveAll
'@@											dicAvailableParameters("TabName") = "Available Parameters"
'@@											dicAvailableParameters("Object") = "000025/A;1-bool"
'@@											dicAvailableParameters("ColName")  =  "Object"
'@@
'@@									  	Call Fn_SISW_Mech_AvailableParam_DictionaryOperations("SelectCell", dicAvailableParameters,  "")
'@@									  	Call Fn_SISW_Mech_AvailableParam_DictionaryOperations("DoubleClick", dicAvailableParameters,  "")
'@@										Call Fn_SISW_Mech_AvailableParam_DictionaryOperations("Assign", dicAvailableParameters,  "")
'@@										Call Fn_SISW_Mech_AvailableParam_DictionaryOperations("Unassign", dicAvailableParameters,  "")
'@@									 	Call Fn_SISW_Mech_AvailableParam_DictionaryOperations("PopupMenuSelect", dicAvailableParameters,  "Delete Override")
'@@
'@@										Case "MultiSelect", "MultiPopupSelect", "MultiSelectAssign", "MultiSelectUnassign", "MultiSelectPopupMenuExist"
'@@											dicAvailableParameters.RemoveAll
'@@											dicAvailableParameters("TabName") = "Available Parameters"
'@@											dicAvailableParameters("Object") =  "Hex45678/A;1-ParmDefHex12345~BCD45678/A;1-ParmDefBCD12345"
'@@											dicAvailableParameters("ColName")  =  "Object"
'@@
'@@									  	Call Fn_SISW_Mech_AvailableParam_DictionaryOperations("MultiSelect", dicAvailableParameters,  "")
'@@									  	Call Fn_SISW_Mech_AvailableParam_DictionaryOperations("MultiPopupSelect", dicAvailableParameters,  "Delete Override")
'@@										Call Fn_SISW_Mech_AvailableParam_DictionaryOperations("MultiSelectAssign", dicAvailableParameters, "")
'@@										Call Fn_SISW_Mech_AvailableParam_DictionaryOperations("MultiSelectUnassign", dicAvailableParameters, "")
'@@									 	Call Fn_SISW_Mech_AvailableParam_DictionaryOperations("MultiSelectPopupMenuExist", dicAvailableParameters,  "Delete Override")
'@@										
'@@										Case "SelectRow"
'@@											dicAvailableParameters.RemoveAll
'@@											dicAvailableParameters("TabName") = "Available Parameters"
'@@											dicAvailableParameters("Object") = "000025/A;1-bool"
'@@	
'@@										Call Fn_SISW_Mech_AvailableParam_DictionaryOperations("SelectRow", dicAvailableParameters,  "")
'@@	
'@@										Case "GetColumnIndex"    								----   	Return -1  if Col Not Found
'@@											dicAvailableParameters.RemoveAll
'@@											dicAvailableParameters("TabName") = "Available Parameters"
'@@											dicAvailableParameters("ColName")  =  "Conversion Rule"
'@@	
'@@										Call Fn_SISW_Mech_AvailableParam_DictionaryOperations("GetColumnIndex",dicAvailableParameters, "")
'@@										
'@@										Case "HeaderPopupMenuSelect" 
'@@											dicAvailableParameters.RemoveAll
'@@											dicAvailableParameters("TabName") = "Available Parameters"
'@@											dicAvailableParameters("ColName")  =  "Conversion Rule"
'@@	
'@@										Call Fn_SISW_Mech_AvailableParam_DictionaryOperations("HeaderPopupMenuSelect",dicAvailableParameters, "Hide column")
'@@	
'@@										Case "CellVerify"
'@@											dicAvailableParameters.RemoveAll
'@@											dicAvailableParameters("TabName") = "Assigned Parameters"
'@@											dicAvailableParameters("Object") = "Hex45678/A;1-ParmDefHex12345"
'@@											dicAvailableParameters("ColName")  =  "Conversion Rule~Minimum Values~Maximum Values~Initial Values"
'@@											dicAvailableParameters("Value")  =  "7*x*x+8*x+9~{0x00}~{0x00}~{0x00}"
'@@
'@@										Call Fn_SISW_Mech_AvailableParam_DictionaryOperations("CellVerify",dicAvailableParameters, "")
'@@	
'@@	History				:	
'@@				Developer Name				Date			 Rev. No.	   Changes Done															 Reviewer
'@@--------------------------------------------	--------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@					Pranav Ingle			06-Nov-2012		  1.0			created
'@@					Pranav Ingle			06-Nov-2012		  1.0			Added Cases		"Ascending", "Descending"
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_SISW_Mech_AvailableParam_DictionaryOperations(sAction, dicAvailableParameters, sPopupMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_Mech_AvailableParam_DictionaryOperations"
	Dim objNatTable, iRowCnt, iColCnt, iCnt, sCol, dicparameter
	Dim aBounds, aMenu, arrValues, iCount, arrCol
	Dim bReturn, startIndex, endIndex, sRectangle
	Dim WshShell,objPopupMenu
	Dim arrObj, myDeviceReplay

	Set objNatTable = JavaWindow("Mechatronics").JavaObject("NatTable")
	Fn_SISW_Mech_AvailableParam_DictionaryOperations = False
	' add call to select Traceability Matrix Tab
	If Fn_UI_ObjectExist("Fn_SISW_Mech_AvailableParam_DictionaryOperations",objNatTable ) = False  Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_Mech_AvailableParam_DictionaryOperations ] Nat Table object of [ Traceability Matrix ] is not visible.")
		Exit function
	End If

	If dicAvailableParameters("ShowAll"&Trim(Replace(dicAvailableParameters("TabName")," ",""))) <> "" Then
		If cBool(dicAvailableParameters("ShowAll"&Trim(Replace(dicAvailableParameters("TabName")," ","")))) Then
			Call Fn_CheckBox_Set("Fn_SISW_Mech_AvailableParam_DictionaryOperations",  JavaWindow("Mechatronics"), "ShowAll"&Trim(Replace(dicAvailableParameters("TabName")," ","")), "ON")
		Else
			Call Fn_CheckBox_Set("Fn_SISW_Mech_AvailableParam_DictionaryOperations",  JavaWindow("Mechatronics"),"ShowAll"&Trim(Replace(dicAvailableParameters("TabName")," ","")), "OFF")
		End If
	End If
	
	Select Case sAction
		Case "SelectCell", "DoubleClick", "PopupMenuSelect", "SetCellData"
			iRowCnt = Fn_SISW_Mech_AvailableParam_DictionaryOperations("GetRowIndex",dicAvailableParameters, "")
			If CInt(iRowCnt) = -1 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_Mech_AvailableParam_DictionaryOperations ] Row Not Found[ " & dicAvailableParameters("Object") & " ].")
				Exit Function
			End If

			sCol = Fn_SISW_Mech_AvailableParam_DictionaryOperations("GetColumnIndex",dicAvailableParameters, "")
			If CInt(sCol) = -1 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_Mech_AvailableParam_DictionaryOperations ] Column Not Found[ " & dicAvailableParameters("ColName") & " ].")
				Exit Function
			End If

			sRectangle = cStr(objNatTable.Object.getCellByPosition(sCol,iRowCnt).getBounds().toString())
			sRectangle =  Right(sRectangle, (Len(sRectangle) - Instr(1, sRectangle, "{", 1)))	
			sRectangle = replace(sRectangle,"}","")
			sRectangle = replace(sRectangle," ","")
			aBounds = split(sRectangle,",")

            If sAction = "SelectCell" Then
				objNatTable.Click cInt(aBounds(0)) + (cInt(aBounds(2))/2),  cInt(aBounds(1)) + (cInt(aBounds(3))/2)  , "LEFT"
				Fn_SISW_Mech_AvailableParam_DictionaryOperations = True
			ElseIf sAction = "DoubleClick" Then
				objNatTable.DblClick cInt(aBounds(0)-10) + (cInt(aBounds(2))/2),  cInt(aBounds(1)) + (cInt(aBounds(3))/2)  , "LEFT"
				Fn_SISW_Mech_AvailableParam_DictionaryOperations = True

			ElseIf sAction = "PopupMenuSelect" Then
				objNatTable.Click cInt(aBounds(0)) + (cInt(aBounds(2))/2),  cInt(aBounds(1)) + (cInt(aBounds(3))/2)  , "LEFT"
                wait 1
				objNatTable.Click cInt(aBounds(0)) + (cInt(aBounds(2))/2),  cInt(aBounds(1)) + (cInt(aBounds(3))/2)  , "RIGHT"
				wait 2
				Fn_SISW_Mech_AvailableParam_DictionaryOperations = Fn_UI_JavaMenu_Select("Fn_SISW_Mech_AvailableParam_DictionaryOperations",JavaWindow("Mechatronics"),sPopupMenu)

			ElseIf sAction = "SetCellData" Then
				objNatTable.Click cInt(aBounds(0)) + (cInt(aBounds(2))/2),  cInt(aBounds(1)) + (cInt(aBounds(3))/2)  , "LEFT"
				
				objNatTable.Type trim(dicAvailableParameters("Value")) + vblf

				Fn_SISW_Mech_AvailableParam_DictionaryOperations = True
			End If

		Case "CellVerify"
			On Error Resume Next
			arrCol = Split(dicAvailableParameters("ColName"), "~")
			arrValues  = Split(dicAvailableParameters("Value"), "~")

			iRowCnt = Fn_SISW_Mech_AvailableParam_DictionaryOperations("GetRowIndex",dicAvailableParameters, "")
			If CInt(iRowCnt) = -1 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_Mech_AvailableParam_DictionaryOperations ] Row Not Found[ " & dicAvailableParameters("Object") & " ].")
				Exit Function
			End If

			Set dicparameter = dicAvailableParameters

			' Verify Values
			For iCount = 0 To Ubound(arrCol)

				dicparameter("ColName") = arrCol(iCount)
        				
				sCol = Fn_SISW_Mech_AvailableParam_BreakdownOperations("GetColumnIndex",dicparameter, "")
				If CInt(sCol) = -1 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_Mech_AvailableParam_DictionaryOperations ] Column Not Found[ " & dicAvailableParameters("ColName") & " ].")
					Exit Function
				End If
				
				sRectangle = cStr(objNatTable.Object.getCellByPosition(sCol,iRowCnt).getBounds().toString())
				sRectangle =  Right(sRectangle, (Len(sRectangle) - Instr(1, sRectangle, "{", 1)))	
				sRectangle = replace(sRectangle,"}","")
				sRectangle = replace(sRectangle," ","")
				aBounds = split(sRectangle,",")

				bReturn = ""
                bReturn = objNatTable.Object.getCellByPosition(sCol, iRowCnt).getDataValue().toString()
				If bReturn = empty Then
					bReturn = ""
				End If

				If trim(arrValues(iCount)) <> trim(bReturn) then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_Mech_AvailableParam_DictionaryOperations ] Value ["& arrValues(iCount) &"] not exist [ " & dicparameter("ColName") & " ].")
					Exit Function
				End If
			Next
			Fn_SISW_Mech_AvailableParam_DictionaryOperations = True
	
' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "Ascending", "Descending"
			On Error Resume Next
			sCol = Fn_SISW_Mech_AvailableParam_DictionaryOperations("GetColumnIndex",dicAvailableParameters, "")

			iRowCnt = cInt(objNatTable.Object.getRowCount())
			bReturn = True	
			For iCnt = 3 to iRowCnt-2
				If iCnt = 3 Then
					startIndex = ""
					startIndex = objNatTable.Object.getCellByPosition(sCol, iCnt).getDataValue().toString()
					If startIndex = empty Then
						startIndex = ""
					End If
				Else
					startIndex = endIndex
				End If

				endIndex = ""
				endIndex = objNatTable.Object.getCellByPosition(sCol,Cint(iCnt+1)).getDataValue().toString()
				If endIndex = empty Then
					endIndex = ""
				End If

				' Check Action Ascending OR Descending
				If sAction = "Ascending" Then
						If endIndex < startIndex then
							bReturn = False
							Exit For
						End If
				ElseIf sAction = "Descending" Then
						If endIndex > startIndex then
							bReturn = False
							Exit For
						End If
				End If

			Next
			If bReturn <> False Then
				Fn_SISW_Mech_AvailableParam_DictionaryOperations = True
			End If

		Case "MultiSelect", "MultiPopupSelect", "MultiSelectAssign", "MultiSelectUnassign", "MultiSelectPopupMenuExist"
			arrObj = Split(dicAvailableParameters("Object"), "~")
    		Set dicparameter = dicAvailableParameters
			Set myDeviceReplay = CreateObject("Mercury.DeviceReplay")

			For iCount = 0 To Ubound(arrObj)
				dicparameter("Object") = arrObj(iCount)

				sCol = 0
				iRowCnt = Fn_SISW_Mech_AvailableParam_DictionaryOperations("GetRowIndex",dicparameter, "")
				If CInt(iRowCnt) = -1 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_Mech_AvailableParam_DictionaryOperations ] Row Not Found [ " & dicAvailableParameters("Object") & " ].")
					Exit Function
				End If

				sRectangle = cStr(objNatTable.Object.getCellByPosition(sCol,iRowCnt).getBounds().toString())
				sRectangle =  Right(sRectangle, (Len(sRectangle) - Instr(1, sRectangle, "{", 1)))	
				sRectangle = replace(sRectangle,"}","")
				sRectangle = replace(sRectangle," ","")
				aBounds = split(sRectangle,",")
				If iCount = 1 Then
					myDeviceReplay.KeyDown 29
				End If
				objNatTable.Click cInt(aBounds(0)) + (cInt(aBounds(2))/2),  cInt(aBounds(1)) + (cInt(aBounds(3))/2)  , "LEFT"
			Next
			Fn_SISW_Mech_AvailableParam_DictionaryOperations = True
            myDeviceReplay.KeyUp 29

			If sAction = "MultiPopupSelect" Then
				sCol = 1
				If dicAvailableParameters("TabName") = "Assigned Parameters" Then
					sCol = 2
				End If

				sRectangle = cStr(objNatTable.Object.getCellByPosition(sCol,iRowCnt).getBounds().toString())
				sRectangle =  Right(sRectangle, (Len(sRectangle) - Instr(1, sRectangle, "{", 1)))	
				sRectangle = replace(sRectangle,"}","")
				sRectangle = replace(sRectangle," ","")
				aBounds = split(sRectangle,",")

				objNatTable.Click cInt(aBounds(0)) + (cInt(aBounds(2))/2),  cInt(aBounds(1)) + (cInt(aBounds(3))/2)  , "RIGHT"
				wait 2
				Fn_SISW_Mech_AvailableParam_DictionaryOperations = Fn_UI_JavaMenu_Select("Fn_SISW_Mech_AvailableParam_DictionaryOperations",JavaWindow("Mechatronics"),sPopupMenu)

			ElseIf sAction = "MultiSelectAssign" Then
				Fn_SISW_Mech_AvailableParam_DictionaryOperations = Fn_Button_Click("Fn_SISW_Mech_AvailableParam_DictionaryOperations", JavaWindow("Mechatronics"),"Assign")
			ElseIf sAction = "MultiSelectUnassign" Then
				Fn_SISW_Mech_AvailableParam_DictionaryOperations = Fn_Button_Click("Fn_SISW_Mech_AvailableParam_DictionaryOperations", JavaWindow("Mechatronics"),"Unassign")
			Elseif sAction = "MultiSelectPopupMenuExist" then
                sCol = 1
				If dicAvailableParameters("TabName") = "Assigned Parameters" Then
					sCol = 2
				End If

				sRectangle = cStr(objNatTable.Object.getCellByPosition(sCol,iRowCnt).getBounds().toString())
				sRectangle =  Right(sRectangle, (Len(sRectangle) - Instr(1, sRectangle, "{", 1)))	
				sRectangle = replace(sRectangle,"}","")
				sRectangle = replace(sRectangle," ","")
				aBounds = split(sRectangle,",")

				objNatTable.Click cInt(aBounds(0)) + (cInt(aBounds(2))/2),  cInt(aBounds(1)) + (cInt(aBounds(3))/2)  , "RIGHT"
				wait 2

				aMenu=Split(sPopupMenu,":")
				Set objPopupMenu=JavaWindow("Mechatronics").JavaMenu("label:="&aMenu(0))
				For iCount=1 to ubound(aMenu)
					Set objPopupMenu=objPopupMenu.JavaMenu("label:="&aMenu(iCount))
				Next
				Fn_SISW_Mech_AvailableParam_DictionaryOperations=objPopupMenu.Exist(5)
				
				wait 1
				Set WshShell = CreateObject("WScript.Shell")
				wait 1
				WshShell.SendKeys "{ESC}"
				Set WshShell =Nothing
			End If

		Case "SelectRow"
        	sCol = 0
    		iRowCnt = Fn_SISW_Mech_AvailableParam_DictionaryOperations("GetRowIndex",dicAvailableParameters, "")
			If CInt(iRowCnt) = -1 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_Mech_AvailableParam_DictionaryOperations ] Row Not Found[ " & dicAvailableParameters("Object") & " ].")
				Exit Function
			End If

			sRectangle = cStr(objNatTable.Object.getCellByPosition(sCol,iRowCnt).getBounds().toString())
			sRectangle =  Right(sRectangle, (Len(sRectangle) - Instr(1, sRectangle, "{", 1)))	
			sRectangle = replace(sRectangle,"}","")
			sRectangle = replace(sRectangle," ","")
			aBounds = split(sRectangle,",")
			objNatTable.Click cInt(aBounds(0)) + (cInt(aBounds(2))/2),  cInt(aBounds(1)) + (cInt(aBounds(3))/2)  , "LEFT"
			Fn_SISW_Mech_AvailableParam_DictionaryOperations = True

		Case "Assign", "Unassign"
			bReturn =Fn_SISW_Mech_AvailableParam_DictionaryOperations("SelectCell", dicAvailableParameters,  "")
			If bReturn = False Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_Mech_AvailableParam_DictionaryOperations ] Row not found [ " & dicAvailableParameters("Object") & " ].")
				Exit Function
			End If

			If sAction = "Assign" Then
				Fn_SISW_Mech_AvailableParam_DictionaryOperations = Fn_Button_Click("Fn_SISW_Mech_AvailableParam_DictionaryOperations", JavaWindow("Mechatronics"),"Assign")
			ElseIf  sAction = "Unassign" Then
				Fn_SISW_Mech_AvailableParam_DictionaryOperations = Fn_Button_Click("Fn_SISW_Mech_AvailableParam_DictionaryOperations", JavaWindow("Mechatronics"),"Unassign")
			End If
	
	' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "HeaderPopupMenuSelect",  "HeaderSelect"
			sCol = Fn_SISW_Mech_AvailableParam_DictionaryOperations("GetColumnIndex",dicAvailableParameters, "")
			
			If CInt(sCol) <> -1 Then
				sRectangle = cStr(objNatTable.Object.getCellByPosition(sCol,2).getBounds().toString())
				sRectangle =  Right(sRectangle, (Len(sRectangle) - Instr(1, sRectangle, "{", 1)))	
				sRectangle = replace(sRectangle,"}","")
				sRectangle = replace(sRectangle," ","")
				aBounds = split(sRectangle,",")

				If sAction = "HeaderPopupMenuSelect" Then
					objNatTable.Click cInt(aBounds(0)-10) + (cInt(aBounds(2))/2),  cInt(aBounds(1)-15) + (cInt(aBounds(3))/2)  , "RIGHT"
					wait 2
					Fn_SISW_Mech_AvailableParam_DictionaryOperations = Fn_UI_JavaMenu_Select("Fn_SISW_Mech_AvailableParam_BreakdownOperations",JavaWindow("Mechatronics"),sPopupMenu)
				ElseIf  sAction = "HeaderSelect" Then
					objNatTable.Click cInt(aBounds(0)-10) + (cInt(aBounds(2))/2),  cInt(aBounds(1)) + (cInt(aBounds(3))/2)  , "LEFT"
					Fn_SISW_Mech_AvailableParam_DictionaryOperations = TRUE
				End If
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_Mech_AvailableParam_BreakdownOperations ] Row not found [ " & dicAvailableParameters("Object") & " ].")
			End If
    ' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "GetColumnIndex"
			On Error Resume Next
			Fn_SISW_Mech_AvailableParam_DictionaryOperations = -1
			
			For iColCnt =1 To 5
				If iColCnt = 2 Then
					' take scroll to left
					Call Fn_SISW_RAC_NatTable_Init(objNatTable) ' Initialise nat table Positions
					objNatTable.Click NT_HScrollBarX + 20, NT_HScrollBarY+5 , "LEFT"
					objNatTable.Click NT_HScrollBarX + 20, NT_HScrollBarY+5 , "LEFT"
				ElseIf iColCnt = 3 OR iColCnt = 4 OR iColCnt = 5 Then
					' Take scroll to right step by step
					objNatTable.Click NT_HScrollBarX + NT_HScrollBarW - 20, NT_HScrollBarY + 5, "LEFT"
				End If

				iCount = cInt(objNatTable.Object.getColumnCount)
    			For iCnt = 1 to iCount-1
					If trim(dicAvailableParameters("ColName")) = trim(objNatTable.Object.getCellByPosition(iCnt,1).getDataValue().toString()) Then
						Fn_SISW_Mech_AvailableParam_DictionaryOperations = iCnt
						exit for
					End If
				Next
				If CInt(Fn_SISW_Mech_AvailableParam_DictionaryOperations) <> -1 Then
					Exit For
				End If
			Next
    ' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "GetRowIndex"
			sCol = 1
			If  dicAvailableParameters("TabName") = "Assigned Parameters" Then
				sCol = 2
			End If

			Fn_SISW_Mech_AvailableParam_DictionaryOperations = -1
			Call Fn_SISW_RAC_NatTable_Init(objNatTable)
			objNatTable.Click NT_HScrollBarX + 20, NT_HScrollBarY+5 , "LEFT"
			objNatTable.Click NT_HScrollBarX + 20, NT_HScrollBarY+5 , "LEFT"

			iRowCnt = cInt(objNatTable.Object.getRowCount())
			For iCnt = 2 to iRowCnt -1
				If dicAvailableParameters("Object") = objNatTable.Object.getCellByPosition(sCol,iCnt).getDataValue().toString() then
					Fn_SISW_Mech_AvailableParam_DictionaryOperations = iCnt
					Exit For
				End If
			Next
	' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_Mech_AvailableParam_DictionaryOperations ] Invalid case [ " & sAction & " ].")
	' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
	End Select
	If  Fn_SISW_Mech_AvailableParam_DictionaryOperations <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [ Fn_SISW_Mech_AvailableParam_DictionaryOperations ] executed successfuly with case [ " & sAction & " ].")
	End If
	Set objNatTable = Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_SISW_Mech_ParameterDefRevisionInfoVerifyProperties

'Description			 :	Function Used to perform operations to verify properties on Parameter Defination 

'Parameters			   :   1.StrAction: Action Name
'										2.dicRevisionPropertyInfo: Revision properties information
'										3.StrButtonName: Button Name
'
'Return Value		   : 	True or False

'Pre-requisite			:	Parameter Defination Revision Information page should be appear on screen

'Examples				:   Dim dicRevisionPropertyInfo
'										Set dicRevisionPropertyInfo = CreateObject("Scripting.Dictionary")
'										
'										dicRevisionPropertyInfo("PropertyName")="Name~Description~Expression"
'										dicRevisionPropertyInfo("Value")="Identical1339~Identical Conversion Rule Created~x"
'										bReturn=Fn_SISW_Mech_ParameterDefRevisionInfoVerifyProperties("EditBox",dicRevisionPropertyInfo,"")
'										
'										dicRevisionPropertyInfo("PropertyName")="Type"
'										dicRevisionPropertyInfo("Value")="Identical"
'										dicRevisionPropertyInfo("CheckPropertyName")="value"
'										bReturn=Fn_SISW_Mech_ParameterDefRevisionInfoVerifyProperties("ListBoxCheckProperty",dicRevisionPropertyInfo,"")
'										
'										dicRevisionPropertyInfo("ConstantName")="A~B"
'										dicRevisionPropertyInfo("ConstantValue")="7~8"
'										bReturn=Fn_SISW_Mech_ParameterDefRevisionInfoVerifyProperties("ConstantsTable",dicRevisionPropertyInfo,"")
'
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												12-Dec-2012								1.0																						Priyanka B
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_SISW_Mech_ParameterDefRevisionInfoVerifyProperties(StrAction,dicRevisionPropertyInfo,StrButtonName)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_Mech_ParameterDefRevisionInfoVerifyProperties"
   'declaring variables
	Dim objParameterDefinitionDialog
    Dim aValues,iCounter,bFlag,cellval,aProperty,iRows,aConstantName,scrollMax,iCount

	Fn_SISW_Mech_ParameterDefRevisionInfoVerifyProperties=false
	'checking existance of [ NewParameterDefinition ] dialog
	if not Window("MechatronicsWindow").JavaDialog("NewParameterDefinition").Exist(6) then
		Exit function
	else
		'Creating object of [ NewParameterDefinition ] dialog
		set objParameterDefinitionDialog=Window("MechatronicsWindow").JavaDialog("NewParameterDefinition")
	end if
	If objParameterDefinitionDialog.JavaSlider("JScrollPane").Exist(3) Then
		'Scrolling till the end of panel
		scrollMax=objParameterDefinitionDialog.JavaSlider("JScrollPane").GetROProperty("max")
		objParameterDefinitionDialog.JavaSlider("JScrollPane").Drag scrollMax
	End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Select Case StrAction
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to verify property value from edit boxes
		Case "EditBox"
				aProperty=Split(dicRevisionPropertyInfo("PropertyName"),"~")
				aValues=Split(dicRevisionPropertyInfo("Value"),"~")
				For iCounter=0 to UBound(aProperty)
					bFlag=False
					objParameterDefinitionDialog.JavaStaticText("PropertyLabel").SetTOProperty "label",aProperty(iCounter)+":"
					If objParameterDefinitionDialog.JavaEdit("EditBox").Exist(2) Then
						If Fn_Edit_Box_GetValue("Fn_SISW_Mech_ParameterDefRevisionInfoVerifyProperties",objParameterDefinitionDialog,"EditBox")=aValues(iCounter) Then
							bFlag=True
						End If
					else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Property [ " & aProperty(iCounter) & " ] is not exist on dialog")
						Exit for
					End If
					If bFlag=False Then
						Exit for
					End If
				Next
				If bFlag=True Then
					Fn_SISW_Mech_ParameterDefRevisionInfoVerifyProperties=true
				End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to verify values from Constants Table
		Case "ConstantsTable"
				aConstantName=Split(dicRevisionPropertyInfo("ConstantName"),"~")
				aValues=Split(dicRevisionPropertyInfo("ConstantValue"),"~")
				iRows=objParameterDefinitionDialog.JavaTable("Constants").GetROProperty("rows")
				For iCounter=0 to ubound(aConstantName)
					bFlag=False
					For iCount=0 to iRows-1
						cellval=objParameterDefinitionDialog.JavaTable("Constants").GetCellData(iCount,"Constant Value")
						If trim(cellval)=aValues(iCounter) and objParameterDefinitionDialog.JavaTable("Constants").GetCellData(iCount,"Constant Name")=aConstantName(iCounter) Then
							bFlag=True
							Exit for
						End If
					Next
					If bFlag=false Then
						Exit for
					End If
				Next
				If bFlag=True Then
					Fn_SISW_Mech_ParameterDefRevisionInfoVerifyProperties=true
				End If
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			Case "ListBoxCheckProperty"
				objParameterDefinitionDialog.JavaStaticText("PropertyLabel").SetTOProperty "label",dicRevisionPropertyInfo("PropertyName")+":"
				If objParameterDefinitionDialog.JavaList("ListBox").Exist(2) Then
					Fn_SISW_Mech_ParameterDefRevisionInfoVerifyProperties = objParameterDefinitionDialog.JavaList("ListBox").CheckProperty(dicRevisionPropertyInfo("CheckPropertyName"),dicRevisionPropertyInfo("Value"))
				End IF
	End	Select
	'Clicking on button
	If StrButtonName<>"" Then
		Call Fn_Button_Click("Fn_SISW_Mech_ParameterDefRevisionInfoVerifyProperties", objParameterDefinitionDialog,StrButtonName)
	End If
	'Releasing object of [ Properties ] dialog
	Set objParameterDefinitionDialog=nothing
End Function

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 	Function to create New Item for Insert Level	 - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:				Fn_SISW_Mech_NewItemForInsertLevel

'Description			 :		 		 Creats New Item for Insert Level

'Parameters			   :	 			1.StrItemType: Type of the item.
'													2.StrItemID: ID of the item it should be unique.
'													3.StrItemRevID:Revision ID of the item.
'													4.StrItemName:Name of item.
'													5.StrItemDesc: Description of the item.
'													6:StrItemUOM: Unit of measure of item. ( not handling this part)

'Return Value		   : 				Item Id  -  Revision Id

'Pre-requisite			:		 		New Item For Insert Level dialog should be open

'Examples				:				 Fn_SISW_Mech_NewItemForInsertLevel("Item","121313","A","my","","")
'History					 :			
'										Developer Name							Date						Rev. No.				Changes Done											Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'										Pranav Ingle							  29-jan-2012				  1.0																							
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'										Swapna G								  18-Jul-2014					  1.0	added this code to handle new Insert Level dialog as per design Changes on TC11.1(20140618b)																				
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Public Function Fn_SISW_Mech_NewItemForInsertLevel(StrItemType,StrItemID,StrItemRevID,StrItemName,StrItemDesc,StrItemUOM)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_Mech_NewItemForInsertLevel"
	Dim sItemId, sRevId, sMenu
	Dim objDialogNewItem	', StrItemID
	Dim sTitle, iItemCount, iCount, crrItem,bFlag
	
	 Set  objDialogNewItem = Fn_SISW_Mechatronics_GetObject("NewItemForInsertLevel")

	sTitle = JavaWindow("DefaultWindow").GetROProperty("title")
	JavaWindow("DefaultWindow").JavaWindow("NewBusinessObject").SetTOProperty "title", "New Item For Insert Level"
	'Select menu [Edit -> Insert Level...]
	If (Not objDialogNewItem.Exist(5)) And (NOT JavaWindow("DefaultWindow").JavaWindow("NewBusinessObject").Exist(1)) Then
		bReturn = Fn_MenuOperation("Exist","Edit:Insert Level...")
		 If bReturn = False Then
				Call Fn_MenuOperation("Select","Edit:Insert Level")
			Else
				Call Fn_MenuOperation("Select","Edit:Insert Level...")		
		 End If
   
	End If
	wait(5)
		 
	If objDialogNewItem.Exist(5) And instr(sTitle, "Structure Manager") > 0 Then
	'Select Item Type
			Call Fn_List_Select("Fn_SISW_Mech_NewItemForInsertLevel", objDialogNewItem,"ItemType",StrItemType)
			
			' Wait till  Button is Enabled
			objDialogNewItem.JavaButton("Next").WaitProperty "enabled", 1, 60000
		    	
			'Click on "Next" button
			objDialogNewItem.JavaButton("Next").Click micLeftBtn
			Call  Fn_ReadyStatusSync(3)
		    	
			If StrItemID <> "" Then
				'Set  Item Id
				Call Fn_Edit_Box("Fn_SISW_Mech_NewItemForInsertLevel",objDialogNewItem,"ItemID", StrItemID)
			End If
			
			If StrItemRevID <> "" Then
				'Set Revision ID
				Call Fn_Edit_Box("Fn_SISW_Mech_NewItemForInsertLevel",objDialogNewItem,"RevisionID", StrItemRevID)
			End If
			
			If  StrItemID = "" or StrItemRevID = "" Then
				'click on assign button
				If Not objDialogNewItem.JavaButton("Assign").GetROProperty("enabled")="0" Then
					Call Fn_Button_Click("Fn_SISW_Mech_NewItemForInsertLevel", objDialogNewItem, "Assign")
				End If
			End If
			wait(3)
			
			'Extract Creation data
			sItemId = Fn_Edit_Box_GetValue("Fn_SISW_Mech_NewItemForInsertLevel", objDialogNewItem,"ItemID")
			sRevId = Fn_Edit_Box_GetValue("Fn_SISW_Mech_NewItemForInsertLevel", objDialogNewItem,"RevisionID")
			
			If  sItemId = "" or sRevId = "" Then
				'click on assign button
				Call Fn_UpdateLogFiles(Time() & " - " & "WARNING - Assign button need to click again.", "")
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Item ID not shown in ItemId Textbox[" + CStr(sItemId) + "]")
				Call Fn_Button_Click("Fn_SISW_Mech_NewItemForInsertLevel", objDialogNewItem, "Assign")
				sItemId = Fn_Edit_Box_GetValue("Fn_SISW_Mech_NewItemForInsertLevel", objDialogNewItem,"ItemID")
				sRevId = Fn_Edit_Box_GetValue("Fn_SISW_Mech_NewItemForInsertLevel", objDialogNewItem,"RevisionID")
			End If
			'*****************************************************************
			
			'Set Item name
			Call Fn_Edit_Box("Fn_SISW_Mech_NewItemForInsertLevel", objDialogNewItem,"ItemName",StrItemName)
		
			'Set description
			If StrItemDesc = "" Then
				StrItemDesc = "Test"
			End If
			Call Fn_Edit_Box("Fn_SISW_Mech_NewItemForInsertLevel", objDialogNewItem,"Description",StrItemDesc)
		
		'	'Set UOM
		'	If StrItemUOM <> "" Then
		'		Call Fn_Edit_Box("Fn_SISW_Mech_NewItemForInsertLevel", objDialogNewItem,"Unit of Measure",StrItemUOM)
		'	End If
			
		    objDialogNewItem.JavaButton("Finish").WaitProperty "enabled", 1, 20000
		    Call Fn_Button_Click("Fn_SISW_Mech_NewItemForInsertLevel", objDialogNewItem, "Finish")
		    '= = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = 
			Fn_SISW_Mech_NewItemForInsertLevel = sItemId & "-" & sRevId
			Call Fn_ReadyStatusSync(1)
			wait(2)
			'Click on Close button
			If Fn_UI_ObjectExist("Fn_SISW_Mech_NewItemForInsertLevel",objDialogNewItem)=True Then
				'Click on Close button
				Call Fn_Button_Click("Fn_SISW_Mech_NewItemForInsertLevel", objDialogNewItem, "Close") 
			End If
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Created an Item of ID [" + CStr(sItemId) + "]")
			
	ElseIf JavaWindow("DefaultWindow").JavaWindow("NewBusinessObject").exist(1) Then  '' added this code to handle new Insert Level dialog as per design Changes on TC11.1(20140618b)
      		Set objDialogNewItem = JavaWindow("DefaultWindow").JavaWindow("NewBusinessObject")
          		'Selecting "Business Object" from list
			iItemCount=Fn_UI_Object_GetROProperty("Fn_SISW_Mech_NewItemForInsertLevel",objDialogNewItem.JavaTree("BusinessObjectType"), "items count")
			For iCount=0 To iItemCount-1
				crrItem=objDialogNewItem.JavaTree("BusinessObjectType").GetItem(iCount)
				If Trim(crrItem)="Most Recently Used:"+Trim(StrItemType) Then
					bFlag=True
					Exit For
				ElseIf Trim(crrItem)="Complete List" Then
					Exit For
				End If
			Next
		
			If bFlag=True Then
				Call Fn_JavaTree_Select("Fn_SISW_Mech_NewItemForInsertLevel", objDialogNewItem, "BusinessObjectType","Most Recently Used")
				Call Fn_JavaTree_Select("Fn_SISW_Mech_NewItemForInsertLevel", objDialogNewItem, "BusinessObjectType","Most Recently Used:"+StrItemType)
			Else
				Call Fn_UI_JavaTree_Expand("Fn_SISW_Mech_NewItemForInsertLevel", objDialogNewItem, "BusinessObjectType","Complete List")
				Call Fn_JavaTree_Select("Fn_SISW_Mech_NewItemForInsertLevel", objDialogNewItem, "BusinessObjectType","Complete List")
				Call Fn_JavaTree_Select("Fn_SISW_Mech_NewItemForInsertLevel", objDialogNewItem, "BusinessObjectType","Complete List:"+StrItemType)	
			End If
			wait 2
		    
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			'Clicking On Next button
			objDialogNewItem.JavaButton("Next").WaitProperty "enabled", 1, 60000
		    Call Fn_Button_Click("Fn_SISW_Mech_NewItemForInsertLevel",objDialogNewItem, "Next")
      	
      	
			If StrItemID <> "" Then
				'Set  Item Id
				objDialogNewItem.JavaStaticText("ObjectName").SetToProperty "label", "ID:"
				Call Fn_SISW_UI_JavaEdit_Operations("Fn_StructureLevelInsert","Set",objDialogNewItem,"EditBox",StrItemID) 
			ElseIf StrItemID = "" Then
				'click on assign button
				objDialogNewItem.JavaStaticText("ObjectName").SetToProperty "label", "ID:"
				objDialogNewItem.JavaButton("AssignID").Click
			End If
			sItemId = objDialogNewItem.JavaEdit("EditBox").GetROProperty ("value")
			
			If StrItemRevID <> "" Then
				'Set Revision ID
				objDialogNewItem.JavaStaticText("ObjectName").SetToProperty "label", "Revision:"
				Call Fn_SISW_UI_JavaEdit_Operations("Fn_StructureLevelInsert","Set",objDialogNewItem,"EditBox",StrItemRevID) 
			ElseIf StrItemRevID = "" Then
				'click on assign button				
				objDialogNewItem.JavaStaticText("ObjectName").SetToProperty "label", "Revision:"
				objDialogNewItem.JavaButton("AssignRevID").Click
			End If
			sRevId = objDialogNewItem.JavaEdit("EditBox").GetROProperty ("value")

			'Extract Creation data
			Call Fn_ReadyStatusSync(2)

		'Set Item name
			Call Fn_SISW_UI_JavaEdit_Operations("Fn_SISW_Mech_NewItemForInsertLevel","Set",objDialogNewItem,"Name",StrItemName)   'changed by Abhisek U
		'Set description
			Call Fn_SISW_UI_JavaEdit_Operations("Fn_SISW_Mech_NewItemForInsertLevel","Set",objDialogNewItem,"Description",StrItemDesc) 
				
			'Click on "Finish" button
			Call Fn_ReadyStatusSync(2)
			objDialogNewItem.JavaButton("Finish").WaitProperty "enabled", "1" ,60000
			If Cint(objDialogNewItem.JavaButton("Finish").GetROProperty("enabled")) = 1 Then
				objDialogNewItem.JavaButton("Finish").Click
			Else
				Fn_SISW_Mech_NewItemForInsertLevel = FALSE				
               		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL:Finish Button not Enabled")
			End If
			
			'' verify existence of Error dialog
			If  JavaWindow("DefaultWindow").JavaWindow("ErrorJavaWindow").exist(1) Then
				Fn_SISW_Mech_NewItemForInsertLevel = FALSE
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),  "Fail : Failed to insert ["+ StrItemID +"/" + StrItemRevID +";1-" + StrItemName   +"] in the structure.")   
				Exit Function
			End if
			'' close Insert Level Dialog
			If objDialogNewItem.exist(1) Then
				objDialogNewItem.JavaButton("Cancel").Click
			End If
			Fn_SISW_Mech_NewItemForInsertLevel = sItemId & "-" & sRevId
            Call Fn_WriteLogFile(Environment.Value("TestLogFile"),  "Pass: Object ["+ StrItemID +"/" + StrItemRevID +";1-" + StrItemName   +"] inserted Successfully in structure.")
      
	Else
			Fn_SISW_Mech_NewItemForInsertLevel = FALSE
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),  "Fail : Failed to insert ["+ StrItemID +"/" + StrItemRevID +";1-" + StrItemName   +"] in the structure.")             
	
	End IF
	
	Set objDialogNewItem=Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_SISW_Mech_SoftwareDesignComponentBasicCreate

'Description			 :	Function Used to Create Software Design Component

'Parameters			   :   '1.StrSDCType: Software Design Component Type
'										 2.bConfigurationItem: Configuration Item option
'										 3.StrID: Software Design Component ID
'										 4.StrRevision: Software Design Component Revision
'										 5.StrName: Software Design Component Name
'										 6.StrDescription: Software Design Component Description
'										 7.UOM: Unit of measure										
'										 8.StrCloseFlag : Flag to close [ NewSoftwareDesignComponent ] dialog
'																		  If StrCloseFlag=false then function will not close the [ NewSoftwareDesignComponent ] dialog after creating SoftwareDesignComponent.
'																		  If StrCloseFlag = "" or true or yes then function will close the [ NewSoftwareDesignComponent ] dialog after creating SoftwareDesignComponent.
'
'Return Value		   : 	ID-Revision

'Pre-requisite			:	Should be log in RAC

'Examples				:   Call Fn_SISW_Mech_SoftwareDesignComponentBasicCreate("SwDesignComp","Off","","","SDC1","New Software Design Component","")
'                       
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												01-Mar-2012								1.0																						Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'If StrCloseFlag=false then function will not close the [ NewSoftwareDesignComponent ] dialog after creating SoftwareDesignComponent.
'If StrCloseFlag = "" or true or yes then function will close the [ NewSoftwareDesignComponent ] dialog after creating SoftwareDesignComponent.
'This flag is use to create multiple SoftwareDesignComponent without closing Dialog
Function Fn_SISW_Mech_SoftwareDesignComponentBasicCreate(StrSDCType,bConfigurationItem,StrID,StrRevision,StrName,StrDescription,UOM,StrCloseFlag)
   'Variable declaration
   Dim objSDCDialog,crrID,crrRevision

   Fn_SISW_Mech_SoftwareDesignComponentBasicCreate=False
   'Creating object of [ NewSoftwareDesignComponent ] dialog
'   Set objSDCDialog=JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("NewSoftwareDesignComponent")
   Set objSDCDialog=Window("TeamcenterWindow").JavaDialog("NewSoftwareDesignComponent")
   'Checking Existance of [ NewSoftwareDesignComponent ] dialog
   If Not objSDCDialog.Exist(6) Then
		'Calling Menu [ File -> New -> Software Design Component ]
		Call Fn_MenuOperation("Select","File:New:Software Design Component")
		Call  Fn_ReadyStatusSync(2)
   End If
   
    If StrSDCType = "SwDesignComp" Then    ' Modified by Chaitali R.
    	StrSDCType = "Software Design Component"
    End If
   
   'Selecting Software Design Component Type
	Call Fn_List_Select("Fn_SISW_Mech_SoftwareDesignComponentBasicCreate",objSDCDialog,"SoftwareDesignComponentType",StrSDCType)
	'Setting [ Configuration Item ] option
	If bConfigurationItem<>"" Then
		Call Fn_CheckBox_Set("Fn_SISW_Mech_SoftwareDesignComponentBasicCreate", objSDCDialog,"ConfigurationItem",bConfigurationItem)
	End If
	'Clicking [ Next ] button
	Call Fn_Button_Click("Fn_SISW_Mech_SoftwareDesignComponentBasicCreate", objSDCDialog,"Next")
	'Setting ID
	If StrID<>"" Then
		Call Fn_Edit_Box("Fn_SISW_Mech_SoftwareDesignComponentBasicCreate",objSDCDialog,"ID", StrID)
	End If
	'Setting Revision
	If StrRevision<>"" Then
		Call Fn_Edit_Box("Fn_SISW_Mech_SoftwareDesignComponentBasicCreate",objSDCDialog,"Revision", StrRevision)
	End If
	If StrID="" or StrRevision="" Then
		'click on assign button to Auto assing ID/Revision
	    Call Fn_Button_Click("Fn_SISW_Mech_SoftwareDesignComponentBasicCreate",objSDCDialog,"Assign")
	End If
	'Retriving [ ID ] and [ Revision ]
	crrID=Fn_Edit_Box_GetValue("Fn_SISW_Mech_SoftwareDesignComponentBasicCreate", objSDCDialog,"ID")
	crrRevision=Fn_Edit_Box_GetValue("Fn_SISW_Mech_SoftwareDesignComponentBasicCreate", objSDCDialog,"Revision")
	'Setting Name
	Call Fn_Edit_Box("Fn_SISW_Mech_SoftwareDesignComponentBasicCreate",objSDCDialog,"Name", StrName)
	'Setting Description
	If StrDescription<>"" Then
		Call Fn_Edit_Box("Fn_SISW_Mech_SoftwareDesignComponentBasicCreate",objSDCDialog,"Description", StrDescription)
	End If
	'Clicking [ Finish ] button
	Call Fn_Button_Click("Fn_SISW_Mech_SoftwareDesignComponentBasicCreate", objSDCDialog,"Finish")
	Call Fn_ReadyStatusSync(1)
	'Function Return ID-Revision
	 Fn_SISW_Mech_SoftwareDesignComponentBasicCreate="'"+crrID+"-"+crrRevision
	 If LCase(StrCloseFlag)<>"false" Then
		'Closing [ NewSoftwareDesignComponent ] dialog
		 If  objSDCDialog.Exist(5) Then
			 Call Fn_Button_Click("Fn_SISW_Mech_SoftwareDesignComponentBasicCreate", objSDCDialog,"Close")
		 End If
	 End If
	 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully created new software design component of ID [" + CStr(crrID) + "]")
	 'Releasing object of [ NewSoftwareDesignComponent ] dialog
	  Set objSDCDialog=Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_SISW_Mech_ParameterDefinitionBasicCreate

'Description			 :	Function Used to Create Parameter Definition

'Parameters			   :   '1.StrParameterDefinitionType: Parameter Definition Type
'										 2.bConfigurationItem: Configuration Item option
'										 3.StrID: Parameter Definition ID
'										 4.StrRevision: Parameter Definition Revision
'										 5.StrName: Parameter Definition Name
'										 6.StrDescription: Parameter Definition Description
'										 2.UOM: Unit of measure										
'
'Return Value		   : 	ID-Revision

'Pre-requisite			:	Should be log in RAC

'Examples				:   Call Fn_SISW_Mech_ParameterDefinitionBasicCreate("ParmDefBool","Off","","","PDB1","New Parameter Definition Bool","")
'                       
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												01-Mar-2012								1.0																						Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_SISW_Mech_ParameterDefinitionBasicCreate(StrParameterDefinitionType,bConfigurationItem,StrID,StrRevision,StrName,StrDescription,UOM)
   'Variable declaration
   Dim objPDDialog,crrID,crrRevision
	
   Fn_SISW_Mech_ParameterDefinitionBasicCreate=False
   'Creating object of [ NewParameterDefinition ] dialog
   Set objPDDialog=JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("NewParameterDefinition")
   'Checking Existance of [ NewParameterDefinition ] dialog
   If Not objPDDialog.Exist(6) Then
		'Calling Menu [ File -> New -> Parameter Definition ]
		Call Fn_MenuOperation("Select","File:New:Parameter Management:Parameter Definition...")
		Call  Fn_ReadyStatusSync(2)
   End If
   'Selecting Parameter Definition Type
	Call Fn_List_Select("Fn_SISW_Mech_ParameterDefinitionBasicCreate",objPDDialog,"ParameterDefinitionType",StrParameterDefinitionType)
	'Setting [ Configuration Item ] option
	If bConfigurationItem<>"" Then
		Call Fn_CheckBox_Set("Fn_SISW_Mech_ParameterDefinitionBasicCreate", objPDDialog,"ConfigurationItem",bConfigurationItem)
	End If
	'Clicking [ Next ] button
	Call Fn_Button_Click("Fn_SISW_Mech_ParameterDefinitionBasicCreate", objPDDialog,"Next")
	'Setting ID
	If StrID<>"" Then
		Call Fn_Edit_Box("Fn_SISW_Mech_ParameterDefinitionBasicCreate",objPDDialog,"ID", StrID)
	End If
	'Setting Revision
	If StrRevision<>"" Then
		Call Fn_Edit_Box("Fn_SISW_Mech_ParameterDefinitionBasicCreate",objPDDialog,"Revision", StrRevision)
	End If
	If StrID="" or StrRevision="" Then
		'click on assign button to Auto assing ID/Revision
	    Call Fn_Button_Click("Fn_SISW_Mech_ParameterDefinitionBasicCreate",objPDDialog,"Assign")
	End If
	'Retriving [ ID ] and [ Revision ]
	crrID=Fn_Edit_Box_GetValue("Fn_SISW_Mech_ParameterDefinitionBasicCreate", objPDDialog,"ID")
	crrRevision=Fn_Edit_Box_GetValue("Fn_SISW_Mech_ParameterDefinitionBasicCreate", objPDDialog,"Revision")
	'Setting Name
	Call Fn_Edit_Box("Fn_SISW_Mech_ParameterDefinitionBasicCreate",objPDDialog,"Name", StrName)
	'Setting Description
	If StrDescription<>"" Then
		Call Fn_Edit_Box("Fn_SISW_Mech_ParameterDefinitionBasicCreate",objPDDialog,"Description", StrDescription)
	End If
	'Clicking [ Finish ] button
	Call Fn_Button_Click("Fn_SISW_Mech_ParameterDefinitionBasicCreate", objPDDialog,"Finish")
	Call Fn_ReadyStatusSync(1)
	'Function Return ID-Revision
	 Fn_SISW_Mech_ParameterDefinitionBasicCreate=crrID+"-"+crrRevision
	'Closing [ NewParameterDefinition ] dialog
	 If  objPDDialog.Exist(5) Then
		 Call Fn_Button_Click("Fn_SISW_Mech_ParameterDefinitionBasicCreate", objPDDialog,"Close")
	 End If
	 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully created new Parameter Definition of ID [" + CStr(crrID) + "]")
	 'Releasing object of [ NewParameterDefinition ] dialog
	  Set objPDDialog=Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_SISW_Mech_ParameterDefinitionGroupDetailsCreate

'Description			 :	Function Used to Create Parameter Definition Group in Detail

'Parameters			   :   '1.StrPDGType: Parameter Definition Group Type
'										 2.dicParameterDefinitionGroupInfo: Parameter Definition Group Info
'
'Return Value		   : 	ID-Revision

'Pre-requisite			:	Should be log in RAC

'Examples				:   dicParameterDefinitionGroupInfo("ConfigurationItem")="off"
'										dicParameterDefinitionGroupInfo("Name")="PDG1"
'										dicParameterDefinitionGroupInfo("Description")="New ParameterDefinationGroup"
'										dicParameterDefinitionGroupInfo("GenericComponentID")="006213"
'										dicParameterDefinitionGroupInfo("Represents")="Packeted Parameter"
'										Call Fn_SISW_Mech_ParameterDefinitionGroupDetailsCreate("ParmGrpDef",dicParameterDefinitionGroupInfo)
'                       
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												01-Mar-2012								1.0																						Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_SISW_Mech_ParameterDefinitionGroupDetailsCreate(StrPDGType,dicParameterDefinitionGroupInfo)
   'Variable declaration
   Dim objPDGDialog,crrID,crrRevision,objStaticText,objChild
	
   Fn_SISW_Mech_ParameterDefinitionGroupDetailsCreate=False
   'Creating object of [ NewParameterDefinitionGroup ] dialog
   Set objPDGDialog=JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("NewParameterDefinitionGroup")
   'Checking Existance of [ NewParameterDefinitionGroup ] dialog
   If Not objPDGDialog.Exist(6) Then
		'Calling Menu [ File -> New -> Parameter Management -> Parameter Definition Group... ]
		Call Fn_MenuOperation("Select","File:New:Parameter Management:Parameter Definition Group...")
		Call  Fn_ReadyStatusSync(2)
   End If
   'Selecting Parameter Definition Type
	Call Fn_List_Select("Fn_SISW_Mech_ParameterDefinitionGroupDetailsCreate",objPDGDialog,"ParameterDefinitionGroupType",StrPDGType)
	'Setting [ Configuration Item ] option
	If dicParameterDefinitionGroupInfo("ConfigurationItem")<>"" Then
		Call Fn_CheckBox_Set("Fn_SISW_Mech_ParameterDefinitionGroupDetailsCreate", objPDGDialog,"ConfigurationItem",dicParameterDefinitionGroupInfo("ConfigurationItem"))
	End If
	'Clicking [ Next ] button
	Call Fn_Button_Click("Fn_SISW_Mech_ParameterDefinitionGroupDetailsCreate", objPDGDialog,"Next")
	'Setting ID
	If dicParameterDefinitionGroupInfo("ID")<>"" Then
		Call Fn_Edit_Box("Fn_SISW_Mech_ParameterDefinitionGroupDetailsCreate",objPDGDialog,"ID", dicParameterDefinitionGroupInfo("ID"))
	End If
	'Setting Revision
	If dicParameterDefinitionGroupInfo("Revision")<>"" Then
		Call Fn_Edit_Box("Fn_SISW_Mech_ParameterDefinitionGroupDetailsCreate",objPDGDialog,"Revision", dicParameterDefinitionGroupInfo("Revision"))
	End If
	If dicParameterDefinitionGroupInfo("ID")="" or dicParameterDefinitionGroupInfo("Revision")="" Then
		'click on assign button to Auto assing ID/Revision
	    Call Fn_Button_Click("Fn_SISW_Mech_ParameterDefinitionGroupDetailsCreate",objPDGDialog,"Assign")
	End If
	'Retriving [ ID ] and [ Revision ]
	crrID=Fn_Edit_Box_GetValue("Fn_SISW_Mech_ParameterDefinitionGroupDetailsCreate", objPDGDialog,"ID")
	crrRevision=Fn_Edit_Box_GetValue("Fn_SISW_Mech_ParameterDefinitionGroupDetailsCreate", objPDGDialog,"Revision")
	'Setting Name
	Call Fn_Edit_Box("Fn_SISW_Mech_ParameterDefinitionGroupDetailsCreate",objPDGDialog,"Name", dicParameterDefinitionGroupInfo("Name"))
	'Setting Description
	If dicParameterDefinitionGroupInfo("Description")<>"" Then
		Call Fn_Edit_Box("Fn_SISW_Mech_ParameterDefinitionGroupDetailsCreate",objPDGDialog,"Description", dicParameterDefinitionGroupInfo("Description"))
	End If
	'Setting GenericComponent ID
	Call Fn_Edit_Box("Fn_SISW_Mech_ParameterDefinitionGroupDetailsCreate",objPDGDialog,"GenericComponentID", dicParameterDefinitionGroupInfo("GenericComponentID"))
	'Clicking [ Next ] button
	Call Fn_Button_Click("Fn_SISW_Mech_ParameterDefinitionGroupDetailsCreate", objPDGDialog,"Next")
	'Selecting Represents
	If dicParameterDefinitionGroupInfo("Represents")<>"" Then
        Call Fn_Button_Click("Fn_SISW_Mech_ParameterDefinitionGroupDetailsCreate", objPDGDialog,"Represents")
		wait 1
		Set objStaticText=Description.Create()
		objStaticText("Class Name").value="JavaStaticText"
		objStaticText("label").value=dicParameterDefinitionGroupInfo("Represents")
		Set objChild=objPDGDialog.ChildObjects(objStaticText)
		objChild(0).Click 1,1
		wait 1
		Set objStaticText=Nothing
		Set objChild=Nothing
	End If
	'Clicking [ Finish ] button
	Call Fn_Button_Click("Fn_SISW_Mech_ParameterDefinitionGroupDetailsCreate", objPDGDialog,"Finish")
	Call Fn_ReadyStatusSync(1)
	'Function Return ID-Revision
	 Fn_SISW_Mech_ParameterDefinitionGroupDetailsCreate=crrID+"-"+crrRevision
	'Closing [ NewParameterDefinitionGroup ] dialog
	 If  objPDGDialog.Exist(5) Then
		 Call Fn_Button_Click("Fn_SISW_Mech_ParameterDefinitionGroupDetailsCreate", objPDGDialog,"Close")
	 End If
	 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully created new Parameter Definition of ID [" + CStr(crrID) + "]")
	 'Releasing object of [ NewParameterDefinitionGroup ] dialog
	  Set objPDGDialog=Nothing
End Function

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 	Function to create New Item for Insert Level	 - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:				Fn_SISW_MechCurrentobjName

'Description			 :		 		 Creats New Item for Insert Level

'Parameters			   :	 			1.StrItemType: Type of the item.

'Return Value		   : 				StrItemType
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Public Function Fn_SISW_MechCurrentobjName(sObjType)
   Select Case sObjType
            Case "ParmDefBCD"
                    Fn_SISW_MechCurrentobjName = "Parameter Definition Binary Coded Decimal"
		    Case "ParmDefInt"
					Fn_SISW_MechCurrentobjName = "Parameter Definition Integer"
			Case "ParmDefHex"
                  Fn_SISW_MechCurrentobjName = "Parameter Definition Hexadecimal"
		    Case "ParmDefDbl"
                  Fn_SISW_MechCurrentobjName = "Parameter Definition Double"
			Case "ParmDefBitDef"
                  Fn_SISW_MechCurrentobjName = "Parameter Definition Bit Definition"
			Case "ParmDefDate"
                  Fn_SISW_MechCurrentobjName = "Parameter Definition Date"
		    Case "ParmDefSED"
                  Fn_SISW_MechCurrentobjName = "Parameter Definition State Encoded"
			Case "ParmDefBitDef"
                  Fn_SISW_MechCurrentobjName = "Parameter Definition Bit Definition"
			Case "ParmDefBool"
				   Fn_SISW_MechCurrentobjName = "Parameter Definition Boolean"
			Case "ParmDefStr"
					Fn_SISW_MechCurrentobjName = "Parameter Definition String"
			Case "ParmGrpDef"
					Fn_SISW_MechCurrentobjName = "Parameter Group Definition"
			Case "ParmGrpVal"
					Fn_SISW_MechCurrentobjName = "Parameter Group Value"
			Case Else
					Fn_SISW_MechCurrentobjName =False
   End Select
End Function
