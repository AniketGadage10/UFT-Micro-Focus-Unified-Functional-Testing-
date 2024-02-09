Option Explicit
'*********************************************************	Function List		***********************************************************************
'0. Fn_SISW_BMIDERAC_GetObject()
'1. Fn_BMIDERAC_ObjectPropertyVerify()
'2. Fn_BMIDERAC_CreateBussinessObject()
'3. Fn_BMIDERAC_RemoteExportOptionsOperations()
'4. Fn_BMIDERAC_VerifyUnitOfMeasures()
'5. Fn_BMIDERAC_AlternateIDDetailsCreate()
'6. Fn_BMIDERAC_NavTreeNodeOperation()
'7. Fn_BMIDERAC_EditObjectProperties()
'8. Fn_BMIDERAC_VerifyToolTipText()
'9. Fn_BMIDERAC_ViewerTabOperation()
'10.Fn_BMIDERAC_CreateBussinessObjectExt()
'11.Fn_BMIDERAC_FormPropertyOperations()
'12.Fn_BMIDERAC_DialogMsgVerify()
'13.Fn_BMIDERAC_ObjectRevRevise()
'14.Fn_BMIDERAC_ItemRevSaveAs()
'15.Fn_BMIDERAC_IdentifierOptionsSettings()
'16.Fn_BMIDERAC_RequirementForDesignBasicCreate()
'17.Fn_BMIDERAC_OperationDataOptionOperation()
'18.Fn_BMIDERAC_PropertiesOperations()
'19.Fn_BMIDERAC_NewBusinessObjectSync()
'*********************************************************	Function List		***********************************************************************
'****************************************    Function to get Object hierarchy ***************************************
'
''Function Name		 	:	Fn_SISW_BMIDERAC_GetObject
'
''Description		    :  	Function to get Object hierarchy

''Parameters		    :	1. sObjectName : Object Handle name
								
''Return Value		    :  	Object \ Nothing
'
''Examples		     	:	Fn_SISW_BMIDERAC_GetObject("PSEApplet")

'History:
'	Developer Name			Date			Rev. No.		Reviewer		Changes Done	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Koustubh Watwe		 14-June-2012		1.0				
'-----------------------------------------------------------------------------------------------------------------------------------
'	Ashwini Kumar		 25-Sept-2013		2.0								Externalized object hierarchies
'-----------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_BMIDERAC_GetObject(sObjectName)
	Dim sObjectXMLPath
	sObjectXMLPath = Fn_GetEnvValue("User", "AutomationDir") & "\TestData\AutomationXML\ObjectXML\BMIDERAC.xml"
	Set Fn_SISW_BMIDERAC_GetObject = Fn_SISW_Setup_GetObjectFromXML(sObjectXMLPath, sObjectName)
End Function
'-------------------------------------------------------------------Function Used to Verify Objects Properties--------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDERAC_ObjectPropertyVerify

'Description			 :	Function Used to Verify Objects Properties

'Parameters			   :   '1.strControlName:Java Control Name Eg : - "EditBox"
										'2.strAction:Property Name
										'3.strExpectedPropertiesValue: Expected Properties Values

'Return Value		   : 	True Or False

'Pre-requisite			:	Properties Dialog Should Be Open

'Examples				: 	Call Fn_BMIDERAC_ObjectPropertyVerify("EditBox","Current Name:Current ID:Type","Test:000006:FunctItem79")
'										Call Fn_BMIDERAC_ObjectPropertyVerify("JavaList","Item Masters:Revisions","000026/A:000026/A;1-Item_121")
'										Call Fn_BMIDERAC_ObjectPropertyVerify("EditBox","Type","ItemRevision")
'										bReturn=Fn_BMIDERAC_ObjectPropertyVerify("RadioButton","p2_WSO_Bool1","False")
'										bReturn=Fn_BMIDERAC_ObjectPropertyVerify("DateCheckBox","p2_WSO_Date1","04-Sep-2012 14:27")


'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				09/12/2010			           1.0																				Sunny R
'													Sandeep N										   				28/09/2012			           1.1																				Priyanka
'													Sanjit K										   				31/01/2013			           1.2					Added call to check existance of [ Properties ]	 dialog														Sandeep N
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'strControlName :- Java Control Name of which value have to Verify
'									Eg:- "EditBox" , "JavaList"
'strAction : - Controls Attached Text (Means Name Of That Property Name)
'						Eg : "Type","Description","UserProperty","Revisions" ( Dont Pass : which appear next to it )
'strExpectedPropertiesValue : - Expected Property Values
'User Can verify Multiple properties of same Control.
'Pass the property Names And There Expected values separeted by Colan ( : )
'Eg :- "EditBox","Type:Name:Description","Item:DemoItem:My DemoItem"
Public Function Fn_BMIDERAC_ObjectPropertyVerify(strControlName,strAction,strExpectedPropertiesValue)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDERAC_ObjectPropertyVerify"
   'Declaring Variables
   Dim ObjPropDialog
   Dim arrAction,arrExpValue,bFlag,strCurrentValue,iCounter
   'Creating Object Of Propety Dialog
   Set ObjPropDialog=JavaWindow("BMIDERACDefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Properties")
   'Checking Existance of "Property Dialog"'@Modified below code by Shailendra on 1-Aug-2013
 	If Fn_UI_ObjectExist("Fn_BMIDERAC_ObjectPropertyVerify",JavaWindow("BMIDERACDefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Properties"))=True Then
		Set ObjPropDialog=JavaWindow("BMIDERACDefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Properties")
	Else
		Fn_BMIDERAC_ObjectPropertyVerify=False
		Set ObjPropDialog=Nothing
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Properties Dialog Existence Failed.")
		Exit Function
	End If
'	@End
	'Ckicking on "All" Proprty Link
    ObjPropDialog.JavaStaticText("BottomLinks").SetTOProperty "label","All"
	If Fn_Java_StaticText_Exist("", ObjPropDialog, "BottomLinks") = true then 
		ObjPropDialog.JavaStaticText("BottomLinks").Click 1,1
	End if
	If Fn_UI_ObjectExist("Fn_BMIDERAC_ObjectPropertyVerify",ObjPropDialog.JavaStaticText("ShowEmptyProperties"))=True Then
		ObjPropDialog.JavaStaticText("ShowEmptyProperties").Click 1,1
	End If
	
	Select Case strControlName
			Case "EditBox","EditBoxExt"
				'Spliting Properties
				arrAction=Split(strAction,":")
				arrExpValue=Split(strExpectedPropertiesValue,":")
				For iCounter=0 To Ubound(arrAction)
					bFlag=False
					'Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_BMIDERAC_ObjectPropertyVerify",ObjPropDialog.JavaEdit("Name"),"attached text",arrAction(iCounter)+".*")
					
				ObjPropDialog.JavaEdit("Name").SetToProperty "attached text",arrAction(iCounter)+".*"
                    If ObjPropDialog.JavaEdit("Name").exist(2) = False Then
                    	ObjPropDialog.JavaStaticText("Property_Name").SetToProperty"label","Responsible Partner Company Code:"
                    	
                          If ObjPropDialog.JavaEdit("Property_Text").exist(2) = True Then
                             strCurrentValue=Fn_Edit_Box_GetValue("Fn_BMIDERAC_ObjectPropertyVerify",ObjPropDialog,"Property_Text")
                    	End If
                    else
                    	strCurrentValue=Fn_Edit_Box_GetValue("Fn_BMIDERAC_ObjectPropertyVerify",ObjPropDialog,"Name")
                    End If

					If strControlName="EditBoxExt"Then
						If Instr(1,Trim(strCurrentValue),Trim(arrExpValue(iCounter)))>0 Then
							bFlag=True
						End If
					Else
						If Trim(strCurrentValue)=Trim(arrExpValue(iCounter)) Then
							bFlag=True
						End If
					End If
					If bFlag=False Then
						Call Fn_Button_Click("Fn_BMIDERAC_ObjectPropertyVerify",ObjPropDialog,"Cancel")
						'Releasing Object Of Property Dialog
                        Exit for
						Set ObjPropDialog=Nothing
					End If
				Next
				If bFlag=True Then
					Fn_BMIDERAC_ObjectPropertyVerify=True
                Else
					Fn_BMIDERAC_ObjectPropertyVerify=False
				End If
			Case "JavaList"
				'Spliting Properties
				arrAction=Split(strAction,":")
				arrExpValue=Split(strExpectedPropertiesValue,":")
				For iCounter=0 To Ubound(arrAction)
					bFlag=False
					Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_BMIDERAC_ObjectPropertyVerify",ObjPropDialog.JavaList("Revisions"),"attached text",arrAction(iCounter)+".*")
					bFlag=Fn_UI_ListItemExist("Fn_BMIDERAC_ObjectPropertyVerify", ObjPropDialog, "Revisions",arrExpValue(iCounter))	
					If bFlag=False Then
						Call Fn_Button_Click("Fn_BMIDERAC_ObjectPropertyVerify",ObjPropDialog,"Cancel")
						'Releasing Object Of Property Dialog
                        Exit for
						Set ObjPropDialog=Nothing
					End If
				Next
				If bFlag=True Then
					Fn_BMIDERAC_ObjectPropertyVerify=True
                Else
					Fn_BMIDERAC_ObjectPropertyVerify=False
				End If
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			'Case to set "Radio button" on
			Case "RadioButton"
				arrAction=Split(strAction,":")
				arrExpValue=Split(strExpectedPropertiesValue,":")
				For iCounter=0 To Ubound(arrAction)
					bFlag=False
					ObjPropDialog.JavaStaticText("Property_Name").SetTOProperty "label",arrAction(iCounter)+":"
					ObjPropDialog.JavaRadioButton("RadioButton").SetTOProperty "attached text",arrExpValue(iCounter)
					If ObjPropDialog.JavaRadioButton("RadioButton").Exist(5) Then
						If LCase(ObjPropDialog.JavaRadioButton("RadioButton").GetROProperty("value"))="1" Then
							bFlag=True
						End If
					End If
					If bFlag=False Then
						Exit for
					End If
				Next
				If bFlag=True Then
					Fn_BMIDERAC_ObjectPropertyVerify=true
				else
					Fn_BMIDERAC_ObjectPropertyVerify=False
				End If
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 		
			'Case to set date
			Case "DateCheckBox"
				arrAction=Split(strAction,":")
				arrExpValue=Split(strExpectedPropertiesValue,"~")
				For iCounter=0 To Ubound(arrAction)
					bFlag=False
					ObjPropDialog.JavaStaticText("Property_Name").SetTOProperty "label",arrAction(iCounter)+":"
					If ObjPropDialog.JavaCheckBox("DateCheckBox").Exist(5) Then
						If instr(1,ObjPropDialog.JavaCheckBox("DateCheckBox").GetROProperty("attached text"),arrExpValue(iCounter)) Then
							bFlag=True
						End If
					End If
				If bFlag=False Then
						Exit for
					End If
				Next
				If bFlag=True Then
					Fn_BMIDERAC_ObjectPropertyVerify=true
				else
					Fn_BMIDERAC_ObjectPropertyVerify=False
				End If

	End Select
	If Fn_UI_ObjectExist("Fn_BMIDERAC_ObjectPropertyVerify",ObjPropDialog)=True Then
		Call Fn_Button_Click("Fn_BMIDERAC_ObjectPropertyVerify",ObjPropDialog,"Cancel")
	End If
	'Releasing Object Of Property Dialog
	Set ObjPropDialog=Nothing
End Function

'-------------------------------------------------------------------Function Used to Create New Bussiness Object In RAC----------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDERAC_CreateBussinessObject
					'Note : Function has been Deprecated, Use Fn_SISW_CreateNewBusinessObject
					
'Description			 :	Function Used to Create New Bussiness Object In RAC

'Parameters			   :   '1.strType:Type Name Eg:- "Item","Folder"
										'2.sField1:First Type Field
										'3.sField2:2 Type Field
										'4.sField3:3 Type Field
										'5.sField4:4 Type Field
										'6.sField5:5 Type Field
										'7.sField6:6 Type Field
										''8.sField7:7 Type Field
										'9.sField8:8Type Field
										'10.sField9:9 Type Field
										'11.strUnitOfMeasure:- Unit Of Measure

'Return Value		   : 	True Or False

'Pre-requisite			:	Should Be Login In TeamCenter

'Examples				: 	Call Fn_BMIDERAC_CreateBussinessObject("Item","123458","A","TestItem","This Is Test Item","","","","","","Cm4780")

'History					 :			
'	Developer Name		Date			Rev. No.		Reviewer		Changes Done				
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Sandeep N			10/12/2010		1.0				Sunny R
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Koustubh Watwe		24/08/2012		1.0				'Note : Function has been Deprecated, Use Fn_SISW_CreateNewBusinessObject
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Public Function Fn_BMIDERAC_CreateBussinessObject(strType,sField1,sField2,sField3,sField4,sField5,sField6,sField7,sField8,sField9,strUnitOfMeasure)
'   Dim ObjObjectDialog
'   Dim bFlag,iItemCount,strItem,iCount
'	bFlag=False
'	Fn_BMIDERAC_CreateBussinessObject=False
'   If Fn_UI_ObjectExist("Fn_BMIDERAC_CreateBussinessObject",JavaWindow("BMIDERACDefaultWindow").JavaWindow("NewBusinessObject"))=True Then
'		Set ObjObjectDialog=JavaWindow("BMIDERACDefaultWindow").JavaWindow("NewBusinessObject")
'	Else
'	    Call Fn_MenuOperation("Select","File:New:Other...")
'		Set ObjObjectDialog=JavaWindow("BMIDERACDefaultWindow").JavaWindow("NewBusinessObject")
'   End If
' 	wait(8)
'	iItemCount=Fn_UI_Object_GetROProperty("Fn_BMIDERAC_CreateBussinessObject",ObjObjectDialog.JavaTree("BusinessObjectType"), "items count")
'	For iCount=0 To iItemCount-1
'		strItem=ObjObjectDialog.JavaTree("BusinessObjectType").GetItem(iCount)
'		If Trim(strItem)="Most Recently Used:"+Trim(strType) Then
'			bFlag=True
'			Exit For
'		ElseIf Trim(strItem)="Complete List" Then
'			Exit For
'		End If
'	Next
'	If bFlag=True Then
'		Call Fn_JavaTree_Select("Fn_BMIDERAC_CreateBussinessObject", ObjObjectDialog, "BusinessObjectType","Most Recently Used")
'		Call Fn_JavaTree_Select("Fn_BMIDERAC_CreateBussinessObject", ObjObjectDialog, "BusinessObjectType","Most Recently Used:"+strType)
'	Else
'		Call Fn_UI_JavaTree_Expand("Fn_BMIDERAC_CreateBussinessObject", ObjObjectDialog, "BusinessObjectType","Complete List")
'		Call Fn_JavaTree_Select("Fn_BMIDERAC_CreateBussinessObject", ObjObjectDialog, "BusinessObjectType","Complete List")
'		Call Fn_JavaTree_Select("Fn_BMIDERAC_CreateBussinessObject", ObjObjectDialog, "BusinessObjectType","Complete List:"+strType)	
'	End If
'	wait(5)
'    Call Fn_Button_Click("Fn_BMIDERAC_CreateBussinessObject", ObjObjectDialog, "Next")
'	wait(5)
'	If sField1<>"" Then
'		Call Fn_UI_EditBox_Type("Fn_BMIDERAC_CreateBussinessObject",ObjObjectDialog,"ID",sField1)
'	End If
'	wait(2)
'	If sField2<>"" Then
'		Call Fn_UI_EditBox_Type("Fn_BMIDERAC_CreateBussinessObject",ObjObjectDialog,"Revision",sField2)
'	End If
'	wait(2)
'	If sField3<>"" Then
'		Call Fn_UI_EditBox_Type("Fn_BMIDERAC_CreateBussinessObject",ObjObjectDialog,"Name",sField3)
'	End If
'	wait(2)
'	If sField4<>"" Then
'		Call Fn_UI_EditBox_Type("Fn_BMIDERAC_CreateBussinessObject",ObjObjectDialog,"Description",sField4)
'	End If
'	wait(2)
'	If sField5<>"" Then
'		Call Fn_UI_EditBox_Type("Fn_BMIDERAC_CreateBussinessObject",ObjObjectDialog,"Text1",sField5)
'	End If
'	wait(2)
'	If sField6<>"" Then
'		Call Fn_UI_EditBox_Type("Fn_BMIDERAC_CreateBussinessObject",ObjObjectDialog,"Text2",sField6)
'	End If
'	wait(2)
'	If sField7<>"" Then
'		Call Fn_UI_EditBox_Type("Fn_BMIDERAC_CreateBussinessObject",ObjObjectDialog,"Text3",sField7)
'	End If
'	wait(2)
'	If sField8<>"" Then
'		Call Fn_UI_EditBox_Type("Fn_BMIDERAC_CreateBussinessObject",ObjObjectDialog,"Text4",sField8)
'	End If
'	wait(2)
'	If sField9<>"" Then
'		Call Fn_UI_EditBox_Type("Fn_BMIDERAC_CreateBussinessObject",ObjObjectDialog,"Text5",sField9)
'	End If
'	wait(2)
'	If strUnitOfMeasure<>"" Then
'		ObjObjectDialog.Maximize
''		JavaWindow("BMIDERACDefaultWindow").JavaWindow("NewBusinessObject").JavaEdit("UnitOfMeasure").Activate
'		JavaWindow("BMIDERACDefaultWindow").JavaWindow("NewBusinessObject").JavaEdit("UnitOfMeasure").Type strUnitOfMeasure
'		ObjObjectDialog.Restore
'	End If
'	wait(5)
'	Call Fn_Button_Click("Fn_BMIDERAC_CreateBussinessObject", ObjObjectDialog, "Finish")
'	wait(2)
'	If Fn_UI_ObjectExist("Fn_BMIDERAC_CreateBussinessObject", ObjObjectDialog.JavaButton("Cancel"))=True Then 
'		Call Fn_Button_Click("Fn_BMIDERAC_CreateBussinessObject", ObjObjectDialog, "Cancel")
'	End If
'
'	Fn_BMIDERAC_CreateBussinessObject=True
'	Set ObjObjectDialog=Nothing
'End Function

'Note : Function has been Deprecated, Use Fn_SISW_CreateNewBusinessObject
Public Function Fn_BMIDERAC_CreateBussinessObject(strType,sField1,sField2,sField3,sField4,sField5,sField6,sField7,sField8,sField9,strUnitOfMeasure)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDERAC_CreateBussinessObject"
   Dim ObjObjectDialog,ObjEdit,ObjChild
   Dim bFlag,iItemCount,strItem,iCount,i
	bFlag=False
	Fn_BMIDERAC_CreateBussinessObject=False
   If not JavaWindow("BMIDERACDefaultWindow").JavaWindow("NewBusinessObject").Exist(6) Then
	    Call Fn_MenuOperation("Select","File:New:Other...")
   End If
   Set ObjObjectDialog=JavaWindow("BMIDERACDefaultWindow").JavaWindow("NewBusinessObject")
   Call  Fn_ReadyStatusSync(1)
	iItemCount=Fn_UI_Object_GetROProperty("Fn_BMIDERAC_CreateBussinessObject",ObjObjectDialog.JavaTree("BusinessObjectType"), "items count")
	For iCount=0 To iItemCount-1
		strItem=ObjObjectDialog.JavaTree("BusinessObjectType").GetItem(iCount)
		If Trim(strItem)="Most Recently Used:"+Trim(strType) Then
			bFlag=True
			Exit For
		ElseIf Trim(strItem)="Complete List" Then
			Exit For
		End If
	Next
	If bFlag=True Then
		Call Fn_JavaTree_Select("Fn_BMIDERAC_CreateBussinessObject", ObjObjectDialog, "BusinessObjectType","Most Recently Used")
		Call Fn_JavaTree_Select("Fn_BMIDERAC_CreateBussinessObject", ObjObjectDialog, "BusinessObjectType","Most Recently Used:"+strType)
	Else
		Call Fn_UI_JavaTree_Expand("Fn_BMIDERAC_CreateBussinessObject", ObjObjectDialog, "BusinessObjectType","Complete List")
		Call Fn_JavaTree_Select("Fn_BMIDERAC_CreateBussinessObject", ObjObjectDialog, "BusinessObjectType","Complete List")
		Call Fn_JavaTree_Select("Fn_BMIDERAC_CreateBussinessObject", ObjObjectDialog, "BusinessObjectType","Complete List:"+strType)	
	End If
	wait 2
    Call Fn_Button_Click("Fn_BMIDERAC_CreateBussinessObject", ObjObjectDialog, "Next")
	wait 1
	'Creating start index to set to Sataic text
	If ObjObjectDialog.JavaObject("Section").Exist(2) Then
		i=2
	else
		i=1
	End If

	If sField1<>"" Then
		ObjObjectDialog.JavaStaticText("StaticText").SetTOProperty "index",i
		If trim(ObjObjectDialog.JavaStaticText("StaticText").GetROProperty("label"))="*" Then
			i=i+1
			ObjObjectDialog.JavaStaticText("StaticText").SetTOProperty "index",i
		End If
		Call Fn_UI_EditBox_Type("Fn_BMIDERAC_CreateBussinessObject",ObjObjectDialog,"ID",sField1)
		wait 1
	End If
	i=i+1
	If sField2<>"" Then
		ObjObjectDialog.JavaStaticText("StaticText").SetTOProperty "index",i
		If trim(ObjObjectDialog.JavaStaticText("StaticText").GetROProperty("label"))="*" Then
			i=i+1
			ObjObjectDialog.JavaStaticText("StaticText").SetTOProperty "index",i
		End If
		Call Fn_UI_EditBox_Type("Fn_BMIDERAC_CreateBussinessObject",ObjObjectDialog,"Revision",sField2)
		wait 1
	End If
	i=i+1
	If sField3<>"" Then
		ObjObjectDialog.JavaStaticText("StaticText").SetTOProperty "index",i
		If trim(ObjObjectDialog.JavaStaticText("StaticText").GetROProperty("label"))="*" Then
			i=i+1
			ObjObjectDialog.JavaStaticText("StaticText").SetTOProperty "index",i
		End If
		Call Fn_UI_EditBox_Type("Fn_BMIDERAC_CreateBussinessObject",ObjObjectDialog,"Name",sField3)
		wait 1
	End If
	i=i+1
	If sField4<>"" Then
		ObjObjectDialog.JavaStaticText("StaticText").SetTOProperty "index",i
		If trim(ObjObjectDialog.JavaStaticText("StaticText").GetROProperty("label"))="*" Then
			i=i+1
			ObjObjectDialog.JavaStaticText("StaticText").SetTOProperty "index",i
		End If
		Call Fn_UI_EditBox_Type("Fn_BMIDERAC_CreateBussinessObject",ObjObjectDialog,"Description",sField4)
		wait 1
	End If
	i=i+1
	If sField5<>"" Then
		ObjObjectDialog.JavaStaticText("StaticText").SetTOProperty "index",i
		If trim(ObjObjectDialog.JavaStaticText("StaticText").GetROProperty("label"))="*" Then
			i=i+1
			ObjObjectDialog.JavaStaticText("StaticText").SetTOProperty "index",i
		End If
		Call Fn_UI_EditBox_Type("Fn_BMIDERAC_CreateBussinessObject",ObjObjectDialog,"Text1",sField5)
		wait 1
	End If
	i=i+1
	If sField6<>"" Then
		ObjObjectDialog.JavaStaticText("StaticText").SetTOProperty "index",i
		If trim(ObjObjectDialog.JavaStaticText("StaticText").GetROProperty("label"))="*" Then
			i=i+1
			ObjObjectDialog.JavaStaticText("StaticText").SetTOProperty "index",i
		End If
		Call Fn_UI_EditBox_Type("Fn_BMIDERAC_CreateBussinessObject",ObjObjectDialog,"Text2",sField6)
		wait 1
	End If
	i=i+1
	If sField7<>"" Then
		ObjObjectDialog.JavaStaticText("StaticText").SetTOProperty "index",i
		If trim(ObjObjectDialog.JavaStaticText("StaticText").GetROProperty("label"))="*" Then
			i=i+1
			ObjObjectDialog.JavaStaticText("StaticText").SetTOProperty "index",i
		End If
		Call Fn_UI_EditBox_Type("Fn_BMIDERAC_CreateBussinessObject",ObjObjectDialog,"Text3",sField7)
		wait 1
	End If
	i=i+1
	If sField8<>"" Then
		ObjObjectDialog.JavaStaticText("StaticText").SetTOProperty "index",i
		If trim(ObjObjectDialog.JavaStaticText("StaticText").GetROProperty("label"))="*" Then
			i=i+1
			ObjObjectDialog.JavaStaticText("StaticText").SetTOProperty "index",i
		End If
		Call Fn_UI_EditBox_Type("Fn_BMIDERAC_CreateBussinessObject",ObjObjectDialog,"Text4",sField8)
		wait 1
	End If
	i=i+1
	If sField9<>"" Then
		ObjObjectDialog.JavaStaticText("StaticText").SetTOProperty "index",i
		If trim(ObjObjectDialog.JavaStaticText("StaticText").GetROProperty("label"))="*" Then
			i=i+1
			ObjObjectDialog.JavaStaticText("StaticText").SetTOProperty "index",i
		End If
		Call Fn_UI_EditBox_Type("Fn_BMIDERAC_CreateBussinessObject",ObjObjectDialog,"Text5",sField9)
		wait 1
	End If

	If strUnitOfMeasure<>"" Then
		ObjObjectDialog.Maximize
		JavaWindow("BMIDERACDefaultWindow").JavaWindow("NewBusinessObject").JavaEdit("UnitOfMeasure").Type strUnitOfMeasure
		ObjObjectDialog.Restore
	End If
	wait 2
	Call Fn_Button_Click("Fn_BMIDERAC_CreateBussinessObject", ObjObjectDialog, "Finish")
	Call  Fn_ReadyStatusSync(1)
	If Fn_UI_ObjectExist("Fn_BMIDERAC_CreateBussinessObject", ObjObjectDialog.JavaButton("Cancel"))=True Then 
		Call Fn_Button_Click("Fn_BMIDERAC_CreateBussinessObject", ObjObjectDialog, "Cancel")
	End If

	Fn_BMIDERAC_CreateBussinessObject=True
	Set ObjObjectDialog=Nothing
End Function

'-------------------------------------------------------------------Function Used to Verify Values Of Remote Export Options Dialog--------------------------------------------------------------
'Function Name		:	Fn_BMIDERAC_RemoteExportOptionsOperations

'Description			 :	Function Used to Verify Values Of Remote Export Options Dialogs

'Parameters			   :   '1.strAction:Action Name
										'2.strExpectedValues:Expected Values

'Return Value		   : 	True Or False

'Pre-requisite			:	Should Be Log In TeamCenter

'Examples				: 	Call  Fn_BMIDERAC_RemoteExportOptionsOperations("IncludeReferenceVerify","Item Masters")

'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				10/12/2010			           1.0																				Sunny R
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_BMIDERAC_RemoteExportOptionsOperations(strAction,strExpectedValues)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDERAC_RemoteExportOptionsOperations"
   Dim ObjExportDialog,ObjExportOptDialog
   Dim arrExpValues,iCounter,bFlag
   Fn_BMIDERAC_RemoteExportOptionsOperations=False
   Set ObjExportDialog=JavaWindow("BMIDERACDefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("RemoteExport")
   Set ObjExportOptDialog=JavaWindow("BMIDERACDefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("RemoteExportOption")
   If Fn_UI_ObjectExist("Fn_BMIDERAC_RemoteExportOptionsOperations", ObjExportOptDialog)=True Then
	Else
		If Fn_UI_ObjectExist("Fn_BMIDERAC_RemoteExportOptionsOperations", ObjExportOptDialog)=True Then
			Call Fn_Button_Click("Fn_BMIDERAC_RemoteExportOptionsOperations", ObjExportDialog, "Advance")
		Else
			Call Fn_MenuOperation("Select","Tools:Multi-Site Collaboration:Send:Remote Export...")
			Call Fn_Button_Click("Fn_BMIDERAC_RemoteExportOptionsOperations", ObjExportDialog, "Advance")
		End If
   End If
   Select Case strAction
		 	Case "IncludeReferenceVerify"
				ObjExportOptDialog.JavaTab("JTabbedPane").Select "Advanced"
				arrExpValues=Split(strExpectedValues,":")
				For iCounter=0 To Ubound(arrExpValues)
						bFlag=False
						bFlag=Fn_UI_ListItemExist("Fn_BMIDERAC_RemoteExportOptionsOperations",ObjExportOptDialog, "IncludeReference",arrExpValues(iCounter))
						If bFlag=False Then
							Call Fn_Button_Click("Fn_BMIDERAC_RemoteExportOptionsOperations", ObjExportOptDialog, "Cancel")
							Call Fn_Button_Click("Fn_BMIDERAC_RemoteExportOptionsOperations", ObjExportDialog, "No")
							Set ObjExportDialog=Nothing
							Set ObjExportOptDialog=Nothing
							Exit Function
						End If
				Next
				Fn_BMIDERAC_RemoteExportOptionsOperations=True
   End Select
	Call Fn_Button_Click("Fn_BMIDERAC_RemoteExportOptionsOperations", ObjExportOptDialog, "Cancel")
	Call Fn_Button_Click("Fn_BMIDERAC_RemoteExportOptionsOperations", ObjExportDialog, "No")
	Set ObjExportDialog=Nothing
	Set ObjExportOptDialog=Nothing
End Function
'-------------------------------------------------------------------Function Used to Verify Unit Of Measure Values Of Item--------------------------------------------------------------------
'Function Name		:	Fn_BMIDERAC_VerifyUnitOfMeasures

'Description			 :	Function Used to Verify Unit Of Measure Values Of Item

'Parameters			   :   '1.strItemType:Item Type
										'2.strExpectedValues:Expected UOM Values

'Return Value		   : 	True Or False

'Pre-requisite			:	Should Be Log In TeamCenter

'Examples				: 	Call  Fn_BMIDERAC_VerifyUnitOfMeasures("Item","Ft1400:Cm4780")

'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				10/12/2010			           1.0																				Sunny R
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Pranav Ingle										   			06/3/2013			           	  1.1																			   Sunny R
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_BMIDERAC_VerifyUnitOfMeasures(strItemType,strExpectedValues)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDERAC_VerifyUnitOfMeasures"
   Dim arrExpValues,iCount,bFlag,iCounter
   Dim ObjStatText,ObjDilogChild,objNewItem
	Dim iItemCount,strItem
	Dim intNodeCount, sTreeItem

   Fn_BMIDERAC_VerifyUnitOfMeasures=False
   Set objNewItem=Fn_SISW_GetObject("New Item")

   'Select menu [File -> New -> Item...]
	If not objNewItem.Exist(5) Then
        Call Fn_MenuOperation("Select","File:New:Item...")
		Call Fn_ReadyStatusSync(3)
	End	If

	iItemCount=Fn_UI_Object_GetROProperty("Fn_BMIDERAC_VerifyUnitOfMeasures",objNewItem.JavaTree("ItemType"), "items count")
	For iCount=0 To iItemCount-1
		strItem=objNewItem.JavaTree("ItemType").GetItem(iCount)
		If Trim(strItem)="Most Recently Used:"+Trim(strItemType) Then
			bFlag=True
			Exit For
		ElseIf Trim(strItem)="Complete List" Then
			Exit For
		End If
	Next
	If bFlag=True Then
		Call Fn_JavaTree_Select("Fn_BMIDERAC_VerifyUnitOfMeasures", objNewItem, "ItemType","Most Recently Used")
		Call Fn_JavaTree_Select("Fn_BMIDERAC_VerifyUnitOfMeasures", objNewItem, "ItemType","Most Recently Used:"+strItemType)
	Else
		Call Fn_UI_JavaTree_Expand("Fn_BMIDERAC_VerifyUnitOfMeasures", objNewItem, "ItemType","Complete List")
		Call Fn_JavaTree_Select("Fn_BMIDERAC_VerifyUnitOfMeasures", objNewItem, "ItemType","Complete List")
		Call Fn_JavaTree_Select("Fn_BMIDERAC_VerifyUnitOfMeasures", objNewItem, "ItemType","Complete List:"+strItemType)	
	End If
	wait 2
    Call Fn_Button_Click("Fn_BMIDERAC_VerifyUnitOfMeasures", objNewItem, "Next")
	wait 1

    objNewItem.JavaButton("UnitOfMeasure").Click
	wait(3)
   arrExpValues=Split(strExpectedValues,":")
   For iCount=0 To UBound(arrExpValues)
		bFlag=False
        Set ObjStatText = Description.Create()
		ObjStatText("Class Name").value = "JavaTree"

		Set ObjDilogChild= objNewItem.ChildObjects(ObjStatText)
		intNodeCount = ObjDilogChild(0).GetROProperty ("items count") 

		For iCounter = 0 to   intNodeCount -1
			sTreeItem = ObjDilogChild(0).GetItem(iCounter)
			If sTreeItem=arrExpValues(iCount) Then
				bFlag = True
				Exit For
			End If
		Next
		If bFlag=False Then
			Set objNewItem=nothing
			Exit Function
		End If
		Set  ObjStatText=Nothing
		Set ObjDilogChild=Nothing
   Next
   Call Fn_Button_Click("Fn_BMIDERAC_VerifyUnitOfMeasures", objNewItem, "Cancel") 
   If Fn_UI_ObjectExist("Fn_BMIDERAC_VerifyUnitOfMeasures",objNewItem)=True Then
        Call Fn_Button_Click("Fn_BMIDERAC_VerifyUnitOfMeasures",objNewItem, "Cancel")
	End If
   Fn_BMIDERAC_VerifyUnitOfMeasures=True
   Set objNewItem=nothing
End Function

'-------------------------------------------------------------------Function Used to Create Detailed AlterNate ID----------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDERAC_AlternateIDDetailsCreate

'Description			 :	Function Used to Create Detailed AlterNate ID

'Parameters			   :   '1.strItemRevision:Item Revision Name
										'2.strContext:Context
										'3.strType:Type
										'4.strID:ID
										'5.strRev:Revision
										'6.strName:Name
										'7.strDesc:Description
										'8.strAdditionalIDInfo:Additional Alternate ID Information (Pass multiple values with : Separated)
										'9.strAdditionalRevInfo:Additional Alternate Revision Information (Pass multiple values with : Separated)
										'10.strDisplayOpt:Display Option (Pass multiple values with : Separated)

'Return Value		   : 	True Or False

'Pre-requisite			:	Should Be Log In TeamCenter

'Examples				: 	Call  Fn_BMIDERAC_AlternateIDDetailsCreate("000025/A;1-tEST","","","98227","A","Test","Test New ID","","","")

'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				10/12/2010			           1.0																				Sunny R
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_BMIDERAC_AlternateIDDetailsCreate(strItemRevision,strContext,strType,strID,strRev,strName,strDesc,strAdditionalIDInfo,strAdditionalRevInfo,strDisplayOpt)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDERAC_AlternateIDDetailsCreate"
   'Variable Declaration
   Dim arrAddIDInfo,arrDisplayOpt
   Dim ObjNewIdDialog,objSelectType,intNoOfObjects
   Fn_BMIDERAC_AlternateIDDetailsCreate=False
   'Creating Object Of "NewID" Dialog
   Set ObjNewIdDialog=JavaWindow("BMIDERACDefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("NewID")
   'Checking Existance Of "NewID" Dialog
	If  Fn_UI_ObjectExist("Fn_BMIDERAC_AlternateIDDetailsCreate",ObjNewIdDialog)=False Then
		'Invoking "NewID" Dialog
		Call Fn_MenuOperation("Select","File:New:ID...")
	End If
	'Selecting Item Revision
	If  strItemRevision<>"" Then
		Call Fn_Button_Click("Fn_BMIDERAC_AlternateIDDetailsCreate", ObjNewIdDialog, "Revisiondropdown")
        Set objSelectType=description.Create()
		objSelectType("Class Name").value = "JavaStaticText"
		objSelectType("label").value = strItemRevision
		Set  intNoOfObjects = ObjNewIdDialog.ChildObjects(objSelectType)
		intNoOfObjects(0).Click 1,1
		Set objSelectType=Nothing
		Set  intNoOfObjects = Nothing
	End If
	If CInt(Fn_UI_Object_GetROProperty("Fn_BMIDERAC_AlternateIDDetailsCreate",ObjNewIdDialog.JavaButton("Next"), "enabled"))=CInt(1) Then
		'Clicking On Next Button
		Call Fn_Button_Click("Fn_BMIDERAC_AlternateIDDetailsCreate", ObjNewIdDialog, "Next")
		'Selecting Context
		If strContext<>"" Then
			Call Fn_List_Select("Fn_BMIDERAC_AlternateIDDetailsCreate", ObjNewIdDialog, "Context",strContext)
		End If
		'Selecting Type
		If strType<>"" Then
			Call Fn_List_Select("Fn_BMIDERAC_AlternateIDDetailsCreate", ObjNewIdDialog, "Type",strType)
		End If
		'Setting ID
		If strID<>"" Then
			Call Fn_Edit_Box("Fn_BMIDERAC_AlternateIDDetailsCreate",ObjNewIdDialog,"ID",strID)
		End If
		'Setting Revision
		If strRev<>"" Then
			Call Fn_Edit_Box("Fn_BMIDERAC_AlternateIDDetailsCreate",ObjNewIdDialog,"Revision",strRev)
		End If
		'Setting Name
		If strName<>"" Then
			Call Fn_Edit_Box("Fn_BMIDERAC_AlternateIDDetailsCreate",ObjNewIdDialog,"Name",strName)
		End If
		If strDesc<>"" Then
			Call Fn_Edit_Box("Fn_BMIDERAC_AlternateIDDetailsCreate",ObjNewIdDialog,"Description",strDesc)
		End If
		'Clicking On Next Button
	End If
	If CInt(Fn_UI_Object_GetROProperty("Fn_BMIDERAC_AlternateIDDetailsCreate",ObjNewIdDialog.JavaButton("Next"), "enabled"))=CInt(1) Then
	Call Fn_Button_Click("Fn_BMIDERAC_AlternateIDDetailsCreate", ObjNewIdDialog, "Next")
	'Setting Additional ID Information
	If strAdditionalIDInfo<>"" Then
		arrAddIDInfo=Split(strAdditionalIDInfo,":")
		If arrAddIDInfo(0)<>"" Then
			Call Fn_Edit_Box("Fn_BMIDERAC_AlternateIDDetailsCreate",ObjNewIdDialog,"Field1",arrAddIDInfo(0))
		End If
	End If
	End If
	If CInt(Fn_UI_Object_GetROProperty("Fn_BMIDERAC_AlternateIDDetailsCreate",ObjNewIdDialog.JavaButton("Next"), "enabled"))=CInt(1) Then
	'Clicking On Next Button
	Call Fn_Button_Click("Fn_BMIDERAC_AlternateIDDetailsCreate", ObjNewIdDialog, "Next")
	'Setting Display Option
	If strDisplayOpt<>"" Then
		arrDisplayOpt=Split(strDisplayOpt,":")
		If arrDisplayOpt(0)<>"" Then
			Call Fn_CheckBox_Set("Fn_BMIDERAC_AlternateIDDetailsCreate", ObjNewIdDialog, "ItemIdentifier", arrDisplayOpt(0))
		End If
		If arrDisplayOpt(1)<>"" Then
			Call Fn_CheckBox_Set("Fn_BMIDERAC_AlternateIDDetailsCreate", ObjNewIdDialog, "RevisionIdentifier", arrDisplayOpt(1))
		End If
	End If
	End If
	Call Fn_Button_Click("Fn_BMIDERAC_AlternateIDDetailsCreate", ObjNewIdDialog, "Finish")
	wait(6)
	If  Fn_UI_ObjectExist("Fn_BMIDERAC_AlternateIDDetailsCreate",ObjNewIdDialog)=True Then
		'Invoking "NewID" Dialog
		Call Fn_Button_Click("Fn_BMIDERAC_AlternateIDDetailsCreate", ObjNewIdDialog, "Close")
	End If
	If  Fn_UI_ObjectExist("Fn_BMIDERAC_AlternateIDDetailsCreate",ObjNewIdDialog)=True Then
		'Invoking "NewID" Dialog
		Call Fn_Button_Click("Fn_BMIDERAC_AlternateIDDetailsCreate", ObjNewIdDialog, "Close")
	End If
	Fn_BMIDERAC_AlternateIDDetailsCreate=True
	Set ObjNewIdDialog=Nothing
End Function
'-------------------------------------------------------------------Function to action perform on NavTree----------------------------------------------------------------------------------------------
'Function Name		:				Fn_BMIDERAC_NavTreeNodeOperation

'Description			 :		 		 Actions performed in this function are:
'																	1. Node Select

'Parameters			   :			1. StrAction: Action to be performed
'												  2. StrNodeName: Fully qulified tree Path (delimiter as ':') [multiple node are separated by "," ] 
'												  3. StrMenu: Context menu to be selected

'Return Value		   : 				TRUE \ FALSE

'Pre-requisite			:		 		My Teamcenter module window should be displayed

'Examples				:		Call Fn_BMIDERAC_NavTreeNodeOperation("Select","Home:000025-tEST:23487@CustomI-Test","")
'											Call Fn_BMIDERAC_NavTreeNodeOperation("GetSelected","","")
'											Call Fn_BMIDERAC_NavTreeNodeOperation("Exist","Home:000025-tEST:23487@CustomI-Test","")
'
'History					 :		
'													Developer Name				Date						Rev. No.			Changes Done																		Reviewer
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'														Sandeep N					14-Dec-10				1.0																																Sunny R
'														Sandeep N					17-Dec-10				1.0																																Sunny R
'														Sandeep N					18-Nov-11				1.1				Added New Solution for Nav Tree Nade selection				Sunny R
'														Sandeep N					17-Dec-10				1.2				Added Case "Exist"																		    Sunny R
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_BMIDERAC_NavTreeNodeOperation(StrAction,StrNodeName,StrMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDERAC_NavTreeNodeOperation"
   Dim objJavaWindowMyTc
	Dim sItemPath,aStrNode,bFlag,i,iNodeItemsCount
	Dim oCurrentNode,eStrNode, iCount
	Fn_BMIDERAC_NavTreeNodeOperation=False

   Set objJavaWindowMyTc = JavaWindow("MyTeamcenter")
   Select Case StrAction
		Case "Select","Exist"
            'Initial Item Path
			sItemPath="#0"
			aStrNode = Split (StrNodeName, ":")
			bFlag=False
			
			Set oCurrentNode =JavaWindow("MyTeamcenter").JavaTree("NavTree").Object.getItem(0)
			'To handle the situation where operation needs to be performed on Root Node
			If UBound(aStrNode) = 0 Then
				bFlag=True
			Else
				'To Select first Occurance of Node
				For each eStrNode In aStrNode
					iNodeItemsCount = oCurrentNode.getItemCount()
					iCount=iCount+1
					bFlag=False
				
						For i = 0 to iNodeItemsCount - 1
							If Trim(oCurrentNode.getItem(i).getData().toString()) = Trim(eStrNode) Then
									Set oCurrentNode = oCurrentNode.getItem(i)
									sItemPath = sItemPath & ":#" & i
									bFlag=True
									Exit For
							End If
						Next
						If iCount=1 Then
							bFlag=True
						Else
							If bFlag=False Then
								Exit For
							End If
						End If
				Next 
			End If
			If bFlag=True Then
				If StrAction="Select" Then
					Call Fn_JavaTree_Select("Fn_BMIDERAC_NavTreeNodeOperation", objJavaWindowMyTc, "NavTree",sItemPath)
				End If
				Fn_BMIDERAC_NavTreeNodeOperation=True
			End If
			Set oCurrentNode =Nothing

		Case "GetSelected"
			Fn_BMIDERAC_NavTreeNodeOperation=objJavaWindowMyTc.JavaTree("NavTree").GetROProperty("value")
		Case Else
			Fn_BMIDERAC_NavTreeNodeOperation=False
   End Select
   Set objJavaWindowMyTc =Nothing
End Function
'-------------------------------------------------------------------Function Used to Modify(Edit) Objects Properties------------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDERAC_EditObjectProperties

'Description			 :	Function Used to Modify(Edit) Objects Properties

'Parameters			   :   '1.strControlName:Java Control Name Eg : - "DropDownList"
										'2.StrPropertyName:Property Name
										'3.StrNewValue: New Properties Values

'Return Value		   : 	True Or False

'Pre-requisite			:	Properties Dialog Should Be Open

'Examples				: 	Call Fn_BMIDERAC_EditObjectProperties("DropDownList","pDispName1","22 , LOVDesc11")
'										Call Fn_BMIDERAC_EditObjectProperties("EditBox","pDispName1","Work in progress")
'										For Case " AddClipboardObject " Pass index of edit_16 Check box separated by ~
'										Call Fn_BMIDERAC_EditObjectProperties("AddClipboardObject~3","","")
'										Call Fn_BMIDERAC_EditObjectProperties("DropDownTable~SaveAdnCheckIn","Name","BDisp")
'										bReturn=Fn_BMIDERAC_EditObjectProperties("RadioButton","p2_WSO_Bool1","False")
'										bReturn=Fn_BMIDERAC_EditObjectProperties("DateCheckBox","p2_WSO_Date1","04-Sep-2012 08:10")

'History					 :			
'	Developer Name			Date						Rev. No.						Changes Done										Reviewer
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Sandeep N				14/12/2010			           1.0																												Sunny R
'	Sandeep N				15/12/2010			           1.1																												Sunny R
'	Sandeep N				22/12/2011			           1.2						Added case 	"AddClipboardObject"				Veena G
'	Sandeep N				28/12/2011			           1.3						Added case 	"EditBoxSet"							Pranav I
'	Sandeep N				28/09/2012			           1.4						Added case 	"RadioButton","DateCheckBox"							Priyanka B
'	Sandeep N				31/12/2012			           1.5						Added case 	"DropDownTable"							Priyanka B
'	Koustubh Watwe			13/03/2013			           1.5						Modified Case "EditBox" to handle Date fields added ~ as a separator
'	Ganesh B				01/04/2014			           1.5						Modified Case "AddClipboardObject" to handle two hierarchy of Check in Dialog
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_BMIDERAC_EditObjectProperties(StrControlName,StrPropertyName,StrNewValue)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDERAC_EditObjectProperties"
	Dim ObjPropDialog
	Dim arrProp,arrNewValues,iCounter,arrControlName
	Dim objTable,objChild,bFlag
	Dim StrTitle, objCheckOut, objCheckIn
	Set objCheckOut=Fn_SISW_GetObject("Check-Out@2")
	If objCheckOut.Exist(1) Then
		call Fn_ObjectCheckOut("Menu CheckOut", "", "", "", "", "","", "")
	End If
	'Click on  'Check-Out and Edit' Button
	set ObjChkOut=JavaWindow("BMIDERACDefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Properties").JavaButton("CheckOutAndEdit")
    wait(5)
    If ObjChkOut.Exist(1) Then
    	Call Fn_Button_Click("Fn_BMIDERAC_EditObjectProperties", JavaWindow("BMIDERACDefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Properties"), "CheckOutAndEdit")
    End If   
    If JavaWindow("DefaultWindow").JavaWindow("Check-Out").Exist(2) Then
		Call Fn_Button_Click("Fn_BMIDERAC_EditObjectProperties", JavaWindow("DefaultWindow").JavaWindow("Check-Out"), "OK")
	End If
	'' get tiltle TC window
	StrTitle=JavaWindow("DefaultWindow").GetROProperty("title")
'	JavaWindow("BMIDERACDefaultWindow").Maximize
   Set ObjPropDialog=JavaWindow("BMIDERACDefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("EditProperties")
   Fn_BMIDERAC_EditObjectProperties=False
   'Checking Existance of "Property Dialog"
 	If Fn_UI_ObjectExist("Fn_BMIDERAC_EditObjectProperties",ObjPropDialog)=True Then
	ElseIf Fn_UI_ObjectExist("Fn_BMIDERAC_EditObjectProperties",JavaWindow("BMIDERACDefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("CheckOut"))=True Then
	    Call Fn_Button_Click("Fn_BMIDERAC_EditObjectProperties",JavaWindow("BMIDERACDefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("CheckOut"),"Yes")
	End If
	'Ckicking on "All" Proprty Link
    ObjPropDialog.JavaStaticText("BottomLink").SetTOProperty "label","All"
	If Fn_Java_StaticText_Exist("", ObjPropDialog, "BottomLink") = true then 
		ObjPropDialog.JavaStaticText("BottomLink").Click 1,1
	End if
	'Claicking On Show empty properties Link
	ObjPropDialog.JavaStaticText("ShowHide").SetTOProperty "label","Show empty properties..."
	If ObjPropDialog.JavaStaticText("ShowHide").Exist(5) Then
		ObjPropDialog.JavaStaticText("ShowHide").Click 1,1
	End If
	arrControlName=Split(StrControlName,"~")

	Select Case arrControlName(0)
			Case "DropDownList"
					ObjPropDialog.JavaStaticText("BottomLink").SetTOProperty "label","Hide empty properties..."
					ObjPropDialog.JavaStaticText("BottomLink").DblClick 1,1
'					Call Fn_Button_Click("Fn_BMIDERAC_EditObjectProperties",ObjPropDialog,"wizard_16")
'					ObjPropDialog.JavaButton("wizard_16").Click micLeftBtn
					ObjPropDialog.JavaStaticText("PropertyName").SetTOProperty "label",StrPropertyName+".*"
					ObjPropDialog.JavaButton("LOVdropdown_16").Click micLeftBtn
					wait 1
'					Call Fn_List_Select("Fn_BMIDERAC_EditObjectProperties", ObjPropDialog, "StepPanelManager",StrPropertyName)
'					ObjPropDialog.JavaList("StepPanelManager").Select StrPropertyName
'					Call Fn_List_Select("Fn_BMIDERAC_EditObjectProperties", ObjPropDialog, "iList",StrNewValue)
'					Call Fn_Button_Click("Fn_BMIDERAC_EditObjectProperties",ObjPropDialog,"Finish")
                    Set objTable=Description.Create()
					objTable("Class Name").value="JavaTable"
					'objTable("tagname").value="LOVTreeTable"
					objTable("toolkit class").value="com.teamcenter.rac.common.lov.view.components.LOVTreeTable"
					Set objChild=ObjPropDialog.ChildObjects(objTable)
					bFlag=False
					For iCounter=0 To objChild(0).GetROProperty("rows")-1
							If trim(objChild(0).Object.getValueAt(iCounter,0).getDisplayableValue())=trim(StrNewValue) Then
								objChild(0).ClickCell iCounter,0
								wait 1
								bFlag=True
								Exit For
							End If
					Next
					If bFlag=True  Then
						Fn_BMIDERAC_EditObjectProperties=True
					End If
			Case "EditBox"
					' ~ separater is added to habdle Date fields
					If instr(StrNewValue,"~") > 0 Then
						arrProp=Split(StrPropertyName,"~")
						arrNewValues=Split(StrNewValue,"~")
					Else
						arrProp=Split(StrPropertyName,":")
						arrNewValues=Split(StrNewValue,":")
					End If
					For iCounter=0 To Ubound(arrProp)
						If trim(arrProp(iCounter)) <> "" Then
							ObjPropDialog.JavaEdit("PropertyEditBox").SetTOProperty "attached text",arrProp(iCounter)+".*"
							If ObjPropDialog.JavaEdit("PropertyEditBox").Exist(5) Then
								Call Fn_UI_EditBox_Type("Fn_BMIDERAC_EditObjectProperties",ObjPropDialog,"PropertyEditBox",arrNewValues(iCounter))
								Call Fn_KeyBoardOperation("SendKeys", "{ENTER}")
							Else
								Set ObjPropDialog=Nothing
								Fn_BMIDERAC_EditObjectProperties=False
								Exit function
							End If
						End If
					Next

					Fn_BMIDERAC_EditObjectProperties=True
			Case "EditBoxSet"
					arrProp=Split(StrPropertyName,":")
					arrNewValues=Split(StrNewValue,":")
					For iCounter=0 To Ubound(arrProp)
						ObjPropDialog.JavaEdit("PropertyEditBox").SetTOProperty "attached text",arrProp(iCounter)+".*"
						Call Fn_Edit_Box("Fn_BMIDERAC_EditObjectProperties",ObjPropDialog,"PropertyEditBox",arrNewValues(iCounter))
					Next
					Fn_BMIDERAC_EditObjectProperties=True

			Case "AddClipboardObject"
'					 If UBound(arrControlName)=0 Then
'					Else
'						JavaWindow("BMIDERACDefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("EditProperties").JavaCheckBox("edit_16").SetTOProperty "index",arrControlName(1)
'					 End If
'					 If JavaWindow("BMIDERACDefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("EditProperties").JavaCheckBox("edit_16").Exist(4) Then
'						JavaWindow("BMIDERACDefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("EditProperties").JavaCheckBox("edit_16").Set "ON"
'						wait 1
'						Call Fn_Button_Click("Fn_BMIDERAC_EditObjectProperties",JavaWindow("BMIDERACDefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("EditProperties"),"add_16")
'						wait 1
'						JavaWindow("BMIDERACDefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("EditProperties").JavaCheckBox("edit_16").Set "OFF"
'						Fn_BMIDERAC_EditObjectProperties=True
'					Else
'						Fn_BMIDERAC_EditObjectProperties=False
'					 End If
'					JavaWindow("BMIDERACDefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("EditProperties").JavaCheckBox("edit_16").SetTOProperty "index",0
					If StrPropertyName<>"" Then ' [TC12-2017091400-5_10_2017-JotibaT-Porting]- Modified code to resolve index issue. 
						JavaWindow("BMIDERACDefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("EditProperties").JavaStaticText("PropertyLabel").SetTOProperty "label",StrPropertyName
						If JavaWindow("BMIDERACDefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("EditProperties").JavaCheckBox("Property_Edit16CheckBox").Exist(4) Then
							JavaWindow("BMIDERACDefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("EditProperties").JavaCheckBox("Property_Edit16CheckBox").Set "ON"
							wait 1
							Call Fn_Button_Click("Fn_BMIDERAC_EditObjectProperties",JavaWindow("BMIDERACDefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("EditProperties"),"add_16")
							wait 1
							JavaWindow("BMIDERACDefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("EditProperties").JavaCheckBox("Property_Edit16CheckBox").Set "OFF"
							Fn_BMIDERAC_EditObjectProperties=True
						Else
							Fn_BMIDERAC_EditObjectProperties=False
						End If
					End If
			Case "AddDate"
					 If UBound(arrControlName)=0 Then
					Else
						JavaWindow("BMIDERACDefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("EditProperties").JavaCheckBox("edit_16").SetTOProperty "index",arrControlName(1)
					 End If
					 If JavaWindow("BMIDERACDefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("EditProperties").JavaCheckBox("edit_16").Exist(4) Then
						JavaWindow("BMIDERACDefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("EditProperties").JavaCheckBox("edit_16").Set "ON"
						wait 1
						JavaWindow("BMIDERACDefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("EditProperties").JavaEdit("CommonEditProperties_JavaEdit").Set StrNewValue
						Call Fn_KeyBoardOperation("SendKey","{TAB}")
						Call Fn_Button_Click("Fn_BMIDERAC_EditObjectProperties",JavaWindow("BMIDERACDefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("EditProperties"),"add_16")
						wait 1
						JavaWindow("BMIDERACDefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("EditProperties").JavaCheckBox("edit_16").Set "OFF"
						Fn_BMIDERAC_EditObjectProperties=True
					Else
						Fn_BMIDERAC_EditObjectProperties=False
					 End If
					JavaWindow("BMIDERACDefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("EditProperties").JavaCheckBox("edit_16").SetTOProperty "index",0

            '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			'Case to set "Radio button" on
			Case "RadioButton"
				arrProp=Split(StrPropertyName,"~")
				arrNewValues=Split(StrNewValue,"~")
				For iCounter=0 To Ubound(arrProp)
					ObjPropDialog.JavaStaticText("PropertyName").SetTOProperty "label",arrProp(iCounter)+":"
					ObjPropDialog.JavaRadioButton("RadioButton").SetTOProperty "attached text",arrNewValues(iCounter)
					If ObjPropDialog.JavaRadioButton("RadioButton").Exist(5) Then
						Call Fn_UI_JavaRadioButton_SetON("Fn_BMIDERAC_EditObjectProperties",ObjPropDialog, "RadioButton")
					else
						Set ObjPropDialog=Nothing
						Fn_BMIDERAC_EditObjectProperties=False
						Exit function
					End If
				Next
				
				If Err.Number < 0 Then
					Fn_BMIDERAC_EditObjectProperties=False
				Else
					Fn_BMIDERAC_EditObjectProperties=True
				End If
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 		
			'Case to set date
			Case "DateCheckBox"
				arrProp=Split(StrPropertyName,"~")
				arrNewValues=Split(StrNewValue,"~")
				For iCounter=0 To Ubound(arrProp)
					ObjPropDialog.JavaStaticText("PropertyName").SetTOProperty "label",arrProp(iCounter)+":"
					If ObjPropDialog.JavaCheckBox("DateCheckBox").Exist(5) Then
						ObjPropDialog.JavaCheckBox("DateCheckBox").Object.setDate(arrNewValues(iCounter))
					else
						Set ObjPropDialog=Nothing
						Fn_BMIDERAC_EditObjectProperties=False
						Exit function
					End If
				Next
				If Err.Number < 0 Then
					Fn_BMIDERAC_EditObjectProperties=False
				Else
					Fn_BMIDERAC_EditObjectProperties=True
				End If

			Case "DropDownTable"
					ObjPropDialog.JavaStaticText("PropertyName").SetTOProperty "label",StrPropertyName+".*"
					ObjPropDialog.JavaButton("EditListDropDown").Click
					Wait 1
                    Set objTable=Description.Create()
					objTable("Class Name").value="JavaTable"
					'objTable("tagname").value="LOVTreeTable"
					objTable("toolkit class").value="com\.teamcenter\.rac\.common\.lov\.view\.components\.LOVTreeTable"
					
					Set objChild=ObjPropDialog.ChildObjects(objTable)
					bFlag=False
					For iCounter=0 To objChild(0).GetROProperty("rows")-1
							If trim(objChild(0).Object.getValueAt(iCounter,0).getDisplayableValue())=trim(StrNewValue) Then
								objChild(0).ClickCell iCounter,0
								Wait 1
								bFlag=True
								Exit For
							End If
					Next
					If bFlag=True  Then
						Fn_BMIDERAC_EditObjectProperties=True
					End If
					'provision to click on [SaveAndCheckIn  ] button
					If ubound(arrControlName)=1 Then
						Call Fn_Button_Click("Fn_BMIDERAC_EditObjectProperties",ObjPropDialog,"SaveAndCheckIn")
						Wait 3
						If Instr(1,StrTitle,"My Teamcenter") or Instr(1,StrTitle,"Change Manager") Then 	''Modified  to handle two hierarchy of Check in Dialog
							Set objCheckIn=Fn_SISW_GetObject("Check-In@2")
						Else
							Set objCheckIn=Fn_SISW_GetObject("Check-In")
						End If
						If  Fn_UI_ObjectExist("Fn_BMIDERAC_EditObjectProperties",objCheckIn)=True Then
							Call Fn_Button_Click("Fn_BMIDERAC_EditObjectProperties",objCheckIn,"Yes")
						 End If
					End If

	End Select
	
	Select Case arrControlName(0)
		Case "DropDownList", "DateCheckBox", "RadioButton", "AddClipboardObject", "EditBoxSet", "EditBox", "AddDate"
			Call Fn_Button_Click("Fn_BMIDERAC_EditObjectProperties",ObjPropDialog,"SaveAndCheckIn")
			if arrControlName(0) = "DropDownList" Then 
				wait 2
			End if
			If Instr(1,StrTitle,"My Teamcenter") or Instr(1,StrTitle,"Change Manager") Then 	''Modified  to handle two hierarchy of Check in Dialog
				Set objCheckIn=Fn_SISW_GetObject("Check-In@2")
			Else
				Set objCheckIn=Fn_SISW_GetObject("Check-In")
			End If
			If  Fn_UI_ObjectExist("Fn_BMIDERAC_EditObjectProperties",objCheckIn)=True Then
				Call Fn_Button_Click("Fn_BMIDERAC_EditObjectProperties",objCheckIn,"Yes")
			 End If
	End Select
	
	Set objCheckOut = Nothing
	Set objCheckIn = Nothing
	Set ObjPropDialog=Nothing
End Function
'-------------------------------------------------------------------Function Used to Verify ToolTip Text Of Any Object-----------------------------------------------------------------------------------
'Function Name		:	Fn_BMIDERAC_VerifyToolTipText

'Description			 :	Function Used to Verify ToolTip Text Of Any Object

'Parameters			   :   '1.ObjectPath:Full Object Path os which you have to verify the ToolTip Text
										'2.strExpectedToolTipText:Expected Tool Tip Text

'Return Value		   : 	True Or False

'Pre-requisite			:	Object Of which you have to verify ToolTip Text is Should appear on screen

'Examples				: 	Set ObjButton=JavaWindow("BMIDERACDefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Revise").JavaButton("referencecopy_16")
'										Call Fn_BMIDERAC_VerifyToolTipText(ObjButton,"Copy As Reference option set")

'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				17/12/2010			           1.0																				Sunny R
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_BMIDERAC_VerifyToolTipText(ObjectPath,strExpectedToolTipText)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDERAC_VerifyToolTipText"
   Dim strCurrentToolTip
	Fn_BMIDERAC_VerifyToolTipText=False
	strCurrentToolTip=ObjectPath.Object.getToolTipText()
	If LCase(Trim(strCurrentToolTip))=LCase(Trim(strExpectedToolTipText)) Then
		Fn_BMIDERAC_VerifyToolTipText=True
	End If
End Function 

'-------------------------------------------------------------------Function Used to Verify/Modify Objects Properties From Viewer Tab------------------------------------------------------------
'Function Name		:	Fn_BMIDERAC_ObjectPropertyVerify

'Description			 :	Function Used to Verify/Modify Objects Properties From Viewer Tab

'Parameters			   :   '1.strAction:Action Name
										'2.strPropertyName:Property Name
										'3.strValue: Value For Modification Or Verify

'Return Value		   : 	True Or False

'Pre-requisite			:	Should Be Log In Teamcenter

'Examples				: 	Call Fn_BMIDERAC_ViewerTabOperation("ModifyDropDownList","pDispName1 / pDispName2","BDisp , LOVDesc2")
'										Call Fn_BMIDERAC_ViewerTabOperation("VerifyEditBox","pDispName1:","CDisp")
'										Call Fn_BMIDERAC_ViewerTabOperation("VerifyEditBox","Name.*","testdoc")
'										Call Fn_BMIDERAC_ViewerTabOperation("ModifyEditBox","Name","NewName")
'										'User Can Use this Case From BOMTable Of PSE perspective for thet need to Activate "hierarchy_16" button
'										'JavaWindow("StructureManager").JavaApplet("PSEApplet").JavaTable("BOMTable").ClickCell 0,"Rev Description","LEFT"
'										Call Fn_BMIDERAC_ViewerTabOperation("ModifyLOVDescriptionTree","","D2MyLOV1_31985:B, DescriptionofB")
'										Msgbox Fn_BMIDERAC_ViewerTabOperation("ModifyList","","Change Analysts")
'										Msgbox Fn_BMIDERAC_ViewerTabOperation("ModifyRadioButton","True","ON")
'										Msgbox Fn_BMIDERAC_ViewerTabOperation("VerifyRadioButton","True","ON")
'History					 :			
'		Developer Name		Date	Rev. No.	Changes Done																Reviewer
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'		Sandeep N		17/12/2010	  1.0																					Sunny R
'		Sandeep N		20/12/2010	  1.0																					Sunny R
'		Sandeep N		27/12/2010	  1.0		Case "ModifyLOVDescriptionTree"				  								Sunny R
'		Sandeep N		03/1/2011	  1.0		Case "ModifyList"											   				Sunny R
'		Pranav Ingle	25/2/2013	  1.1		Case "ModifyEditBox"														Sunny R
'		Sandeep N		07/03/2013	  1.2		Added code to check existance of  Viewer innner tab
'		Sandeep N		26/03/2013	  1.3		Modified case : ModifyDropDownList
'		Nitish B		27/11/2015	  1.4		Modified Case "ModifyList"											[TC1121-2015110900-27_11_2015-VivekA-Maintenance]
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_BMIDERAC_ViewerTabOperation(strAction,strPropertyName,strValue)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDERAC_ViewerTabOperation"
   'Variable Declaration
    Dim bFlag,strCurrentPropVal,arrPropeNames,arrValues,iCounter,strCurrRdVal
	Dim ObjList,ObjStatText,ObjDilogChild,objApplet
   Fn_BMIDERAC_ViewerTabOperation=False
   bFlag=False
   strNode=Fn_BMIDERAC_NavTreeNodeOperation("GetSelected","","")
   'Activating Viewer Tab
   Call Fn_MyTc_TabOperation("Activate", "Viewer")
   'Checking Existance Of "BottomLink"
   If Fn_UI_ObjectExist("Fn_BMIDERAC_ViewerTabOperation", JavaWindow("BMIDERACDefaultWindow").JavaStaticText("ViewerTabBottomLink"))=True Then
	   'Ckicking On "All" Bottom Link
		JavaWindow("BMIDERACDefaultWindow").JavaStaticText("ViewerTabBottomLink").Click 1,1		
   Elseif JavaWindow("BMIDERACDefaultWindow").JavaTab("ViewerInnerTab").Exist(5) then
		JavaWindow("BMIDERACDefaultWindow").JavaTab("ViewerInnerTab").Select "All"
		wait 2
   End If
	'Checking Existance Of "ShowHideProperties" Link
	If  Fn_UI_ObjectExist("Fn_BMIDERAC_ViewerTabOperation",JavaWindow("BMIDERACDefaultWindow").JavaStaticText("ViewerTabShowHideLink"))=True Then
		'Clicking On "ShowHideProperties" Link
		JavaWindow("BMIDERACDefaultWindow").JavaStaticText("ViewerTabShowHideLink").Click 1,1	
		bFlag=True
	End If
	'Creating Object of JavaApplet
'	If JavaWindow("BMIDERACDefaultWindow").JavaWindow("JApplet").Exist(2) Then
'		Set objApplet=JavaWindow("BMIDERACDefaultWindow").JavaWindow("JApplet")
'	Else
		Set objApplet=JavaWindow("BMIDERACDefaultWindow")
'	End If

   Select Case strAction
   'Case to Verify Properties which Control is "EditBox"
		Case "VerifyEditBox"
                JavaWindow("BMIDERACDefaultWindow").JavaStaticText("ViewerTabControlNameText").SetTOProperty "label",strPropertyName
				If JavaWindow("BMIDERACDefaultWindow").JavaEdit("ViewTabPropertyEditBox").Exist(3) Then
					strCurrentPropVal=Fn_Edit_Box_GetValue("Fn_BMIDERAC_ViewerTabOperation",JavaWindow("BMIDERACDefaultWindow"),"ViewTabPropertyEditBox")
					If Trim(strCurrentPropVal)=Trim(strValue) Then
						Fn_BMIDERAC_ViewerTabOperation=True
					End If
				Else
					JavaWindow("BMIDERACDefaultWindow").JavaStaticText("ViewerTabControlNameText").SetTOProperty "label",strPropertyName+":"
					If JavaWindow("BMIDERACDefaultWindow").JavaEdit("ViewTabPropertyEditBox").Exist(3) Then
						strCurrentPropVal=Fn_Edit_Box_GetValue("Fn_BMIDERAC_ViewerTabOperation",JavaWindow("BMIDERACDefaultWindow"),"ViewTabPropertyEditBox")
						If Trim(strCurrentPropVal)=Trim(strValue) Then
							Fn_BMIDERAC_ViewerTabOperation=True
						End If
					End If
				End If
	'Case to Verify Properties which Control is "DropDownList"
		Case "ModifyDropDownList"
				If bFlag=True Then
					Set ObjList=JavaWindow("BMIDERACDefaultWindow").JavaWindow("JApplet")
				Else
					Set ObjList=JavaWindow("BMIDERACDefaultWindow")
				End If
				ObjList.JavaStaticText("ViewerTab_Text").SetTOProperty "label",strPropertyName+":"
				Call Fn_Button_Click("Fn_BMIDERAC_ViewerTabOperation",JavaWindow("BMIDERACDefaultWindow"), "ViewerTab_Button")
				If ObjList.JavaWindow("TreeShell").JavaTree("Tree").Exist(10) Then
				   ObjList.JavaWindow("TreeShell").JavaTree("Tree").Activate strValue
				   wait 2
				End If			
				If JavaWindow("BMIDERACDefaultWindow").JavaButton("Save").Exist(6) Then				
					Fn_BMIDERAC_ViewerTabOperation=Fn_Button_Click("Fn_BMIDERAC_ViewerTabOperation", JavaWindow("BMIDERACDefaultWindow"), "Save")
				Else
					Fn_BMIDERAC_ViewerTabOperation=Fn_ToolBarOperation("Click","Save and Keep Checked-Out", "" )
				End If
				Set ObjList=Nothing
		 'Case to Modify Properties which Control is "EditBox"
		Case "ModifyEditBox"
				arrPropeNames=Split(strPropertyName,"~")
				arrValues=Split(strValue,"~")
				For iCounter=0 To Ubound(arrPropeNames)
					Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_BMIDERAC_ViewerTabOperation",objApplet.JavaStaticText("ViewerTabControlNameText"),"label",arrPropeNames(iCounter)+":")
					Call Fn_Edit_Box("Fn_BMIDERAC_ViewerTabOperation",objApplet,"ViewTabPropertyEditBox",arrValues(iCounter))
				Next
				If JavaWindow("BMIDERACDefaultWindow").JavaButton("Save").Exist(6) Then				
					Call Fn_Button_Click("Fn_BMIDERAC_ViewerTabOperation", JavaWindow("BMIDERACDefaultWindow"), "Save")
				Else
					Call Fn_ToolBarOperation("Click","Save and Keep Checked-Out", "" )
				End If
				Call Fn_ErrorDialogHandle("Properties...","","OK")
                Fn_BMIDERAC_ViewerTabOperation=True
		'"hierarchy_16" Button Should be Displayed
		Case "ModifyLOVDescriptionTree"
                Call Fn_Button_Click("Fn_BMIDERAC_ViewerTabOperation", objApplet, "hierarchy_16")
                objApplet.JavaTree("LOVDescTree").Activate strValue
				Fn_BMIDERAC_ViewerTabOperation=True

		Case "ModifyList"	'[TC1121-2015110900-27_11_2015-VivekA-Maintenance] - Modified by Nitish B
				JavaWindow("BMIDERACDefaultWindow").JavaStaticText("ViewerTab_Text").SetTOProperty "label",strPropertyName
        		Call Fn_Button_Click("Fn_BMIDERAC_ViewerTabOperation", JavaWindow("BMIDERACDefaultWindow"), "ViewerTab_Button")
        		wait (3)
        		JavaWindow("BMIDERACDefaultWindow").JavaWindow("TreeShell").JavaTree("Tree_Prop").Select strValue
        		wait (3)
'               Call Fn_Button_Click("Fn_BMIDERAC_ViewerTabOperation", JavaWindow("BMIDERACDefaultWindow"), "ViewerDropdown")
'				wait(3)
'				Set ObjStatText=Description.Create
'				ObjStatText("Class Name").value="JavaStaticText"
'				ObjStatText("label").value=strValue
'				Set ObjDilogChild=JavaWindow("BMIDERACDefaultWindow").ChildObjects(ObjStatText)
'				ObjDilogChild(0).click 1,1
				If JavaWindow("BMIDERACDefaultWindow").JavaButton("Save").Exist(6) Then				
					Call Fn_Button_Click("Fn_BMIDERAC_ViewerTabOperation", JavaWindow("BMIDERACDefaultWindow"), "Save")
				Else
					Call Fn_ToolBarOperation("Click","Save and Keep Checked-Out", "" )
				End If
				Fn_BMIDERAC_ViewerTabOperation=True

		Case "ModifyRadioButton"
					JavaWindow("BMIDERACDefaultWindow").JavaRadioButton("ViewerOptionBttn").SetTOProperty "attached text",strPropertyName
					JavaWindow("BMIDERACDefaultWindow").JavaRadioButton("ViewerOptionBttn").Set strValue
					If JavaWindow("BMIDERACDefaultWindow").JavaButton("Save").Exist(6) Then				
						Call Fn_Button_Click("Fn_BMIDERAC_ViewerTabOperation", JavaWindow("BMIDERACDefaultWindow"), "Save")
					Else
						Call Fn_ToolBarOperation("Click","Save and Keep Checked-Out", "" )
					End If
					Fn_BMIDERAC_ViewerTabOperation=True

		Case "VerifyRadioButton"
					JavaWindow("BMIDERACDefaultWindow").JavaRadioButton("ViewerOptionBttn").SetTOProperty "attached text",strPropertyName
					strCurrRdVal=JavaWindow("BMIDERACDefaultWindow").JavaRadioButton("ViewerOptionBttn").GetROProperty("value")
					If Cint(strCurrRdVal)=0 Then
						strCurrRdVal="OFF"
					Else
						strCurrRdVal="ON"
					End If
					If strCurrRdVal=UCase(strValue) Then
						Fn_BMIDERAC_ViewerTabOperation=True
					End If

   End Select
   
   Call Fn_BMIDERAC_NavTreeNodeOperation("Select",strNode,"")
	If Fn_UI_ObjectExist("Fn_BMIDERAC_ViewerTabOperation",JavaDialog("SaveChanges"))=True Then
		Call Fn_Button_Click("Fn_BMIDERAC_ViewerTabOperation",JavaDialog("SaveChanges"), "Yes")
	End If
	Set objApplet=nothing
End Function

'-------------------------------------------------------------------Function Used to Create New Bussiness Object In RAC with LOV Values--------------------------------------------------------------------
'Function Name		:	Fn_BMIDERAC_CreateBussinessObjectExt
					'Note : Function has been Deprecated, Use Fn_SISW_CreateNewBusinessObject
'Description			 :	Function Used to Create New Bussiness Object In RAC with LOV Values

'Parameters			   :   '1.strType:Type Name Eg:- "Item","Folder"
										'strListValues: LOV Values
										'2.sField1:First Type Field
										'3.sField2:2 Type Field
										'4.sField3:3 Type Field
										'5.sField4:4 Type Field
										'6.sField5:5 Type Field
										'7.sField6:6 Type Field
										''8.sField7:7 Type Field
										'9.sField8:8Type Field
										'10.sField9:9 Type Field
										'11.strUnitOfMeasure:- Unit Of Measure

'Return Value		   : 	True Or False

'Pre-requisite			:	Should Be Login In TeamCenter

'Examples				: 	Call Fn_BMIDERAC_CreateBussinessObjectExt("Item90_2","","Prop1 , Property1ForLOV1","Test Description","1234","testItem","","","","","","","")
'							Call Fn_BMIDERAC_CreateBussinessObjectExt("Item90_2","LaborRate:PlantCode","40 , 40:10010 , 10010","Test Description","1234","testItem","","","","","","","")

'History					 :			
'	Developer Name		Date			Rev. No.		Reviewer						Changes Done				
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Sandeep N			20/12/2010		1.0				Sunny R
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Koustubh W			24/12/2012		1.0				'Note : Function has been Deprecated, Use Fn_SISW_CreateNewBusinessObject
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'strListValues : This Parameter Can be ( : ) colon separeted
'Note : Function has been Deprecated, Use Fn_SISW_CreateNewBusinessObject
Public Function Fn_BMIDERAC_CreateBussinessObjectExt(strType,strSteps,strListValues,sField1,sField2,sField3,sField4,sField5,sField6,sField7,sField8,sField9,strUnitOfMeasure)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDERAC_CreateBussinessObjectExt"
   Dim ObjObjectDialog
   Dim bFlag,iItemCount,strItem,iCount,arrListValues,iCounter,arrSteps
	bFlag=False
	Fn_BMIDERAC_CreateBussinessObjectExt=False
   If Fn_UI_ObjectExist("Fn_BMIDERAC_CreateBussinessObjectExt",JavaWindow("BMIDERACDefaultWindow").JavaWindow("NewBusinessObject"))=True Then
		Set ObjObjectDialog=JavaWindow("BMIDERACDefaultWindow").JavaWindow("NewBusinessObject")
	Else
	    Call Fn_MenuOperation("Select","File:New:Other...")
		Set ObjObjectDialog=JavaWindow("BMIDERACDefaultWindow").JavaWindow("NewBusinessObject")
   End If
 	wait(8)
	iItemCount=Fn_UI_Object_GetROProperty("Fn_BMIDERAC_CreateBussinessObjectExt",ObjObjectDialog.JavaTree("BusinessObjectType"), "items count")
	For iCount=0 To iItemCount-1
		strItem=ObjObjectDialog.JavaTree("BusinessObjectType").GetItem(iCount)
		If Trim(strItem)="Most Recently Used:"+Trim(strType) Then
			bFlag=True
			Exit For
		ElseIf Trim(strItem)="Complete List" Then
			Exit For
		End If
	Next
	If bFlag=True Then
		Call Fn_JavaTree_Select("Fn_BMIDERAC_CreateBussinessObjectExt", ObjObjectDialog, "BusinessObjectType","Most Recently Used")
		Call Fn_JavaTree_Select("Fn_BMIDERAC_CreateBussinessObjectExt", ObjObjectDialog, "BusinessObjectType","Most Recently Used:"+strType)
	Else
		Call Fn_UI_JavaTree_Expand("Fn_BMIDERAC_CreateBussinessObjectExt", ObjObjectDialog, "BusinessObjectType","Complete List")
		Call Fn_JavaTree_Select("Fn_BMIDERAC_CreateBussinessObjectExt", ObjObjectDialog, "BusinessObjectType","Complete List")
		Call Fn_JavaTree_Select("Fn_BMIDERAC_CreateBussinessObjectExt", ObjObjectDialog, "BusinessObjectType","Complete List:"+strType)	
	End If
	wait(5)
    Call Fn_Button_Click("Fn_BMIDERAC_CreateBussinessObjectExt", ObjObjectDialog, "Next")
	ObjObjectDialog.Maximize
	wait(5)

	If strListValues<>"" Then
		arrListValues=Split(strListValues,":")
		arrSteps=Split(strSteps,":")
		For iCounter=0 To UBound(arrListValues)
			If arrListValues(iCounter)<>"" Then
				ObjObjectDialog.JavaButton("wizard_16").SetTOProperty "index",iCounter
				Call Fn_Button_Click("Fn_BMIDERAC_CreateBussinessObjectExt", ObjObjectDialog, "wizard_16")
				If strSteps<>"" Then
					If arrSteps(iCounter)<>"" Then
						Call Fn_List_Select("Fn_BMIDERAC_CreateBussinessObjectExt", ObjObjectDialog, "StepPanelManager",arrSteps(iCounter))
					End If
				End If
				Call Fn_List_Select("Fn_BMIDERAC_CreateBussinessObjectExt", ObjObjectDialog, "iList",arrListValues(iCounter))
				Call Fn_Button_Click("Fn_BMIDERAC_CreateBussinessObjectExt", JavaWindow("BMIDERACDefaultWindow"), "Finish")
			End If
		Next
	End If
	If sField1<>"" Then
		Call Fn_UI_EditBox_Type("Fn_BMIDERAC_CreateBussinessObjectExt",ObjObjectDialog,"ID",sField1)
		wait(2)
	End If
	If sField2<>"" Then
		Call Fn_UI_EditBox_Type("Fn_BMIDERAC_CreateBussinessObjectExt",ObjObjectDialog,"Revision",sField2)
		wait(2)
	End If
	If sField3<>"" Then
		Call Fn_UI_EditBox_Type("Fn_BMIDERAC_CreateBussinessObjectExt",ObjObjectDialog,"Name",sField3)
		wait(2)
	End If
	If sField4<>"" Then
		Call Fn_UI_EditBox_Type("Fn_BMIDERAC_CreateBussinessObjectExt",ObjObjectDialog,"Description",sField4)
		wait(2)
	End If
	If sField5<>"" Then
		Call Fn_UI_EditBox_Type("Fn_BMIDERAC_CreateBussinessObjectExt",ObjObjectDialog,"Text1",sField5)
		wait(2)
	End If
	If sField6<>"" Then
		Call Fn_UI_EditBox_Type("Fn_BMIDERAC_CreateBussinessObjectExt",ObjObjectDialog,"Text2",sField6)
		wait(2)
	End If
	If sField7<>"" Then
		Call Fn_UI_EditBox_Type("Fn_BMIDERAC_CreateBussinessObjectExt",ObjObjectDialog,"Text3",sField7)
		wait(2)
	End If	
	If sField8<>"" Then
		Call Fn_UI_EditBox_Type("Fn_BMIDERAC_CreateBussinessObjectExt",ObjObjectDialog,"Text4",sField8)
		wait(2)
	End If
	If sField9<>"" Then
		Call Fn_UI_EditBox_Type("Fn_BMIDERAC_CreateBussinessObjectExt",ObjObjectDialog,"Text5",sField9)
	End If
	wait(2)
	If strUnitOfMeasure<>"" Then
		
'		JavaWindow("BMIDERACDefaultWindow").JavaWindow("NewBusinessObject").JavaEdit("UnitOfMeasure").Activate
		JavaWindow("BMIDERACDefaultWindow").JavaWindow("NewBusinessObject").JavaEdit("UnitOfMeasure").Type strUnitOfMeasure
	End If
	ObjObjectDialog.Restore
	wait(5)
	Call Fn_Button_Click("Fn_BMIDERAC_CreateBussinessObjectExt", ObjObjectDialog, "Finish")
	If ObjObjectDialog.Exist(7) Then
		Call Fn_Button_Click("Fn_BMIDERAC_CreateBussinessObjectExt", ObjObjectDialog, "Cancel")
	End If
	Fn_BMIDERAC_CreateBussinessObjectExt=True
	Set ObjObjectDialog=Nothing
End Function
'-------------------------------------------------------------------Function Used to Verify/Modify Form Properties From Property Dialog------------------------------------------------------------
'Function Name		:	Fn_BMIDERAC_FormPropertyOperations

'Description			 :	Function Used to Verify/Modify Form Properties From Property Dialog.
'										Which Appears After Double Clicking On Form

'Parameters			   :   '1.strFormName:Form Name
										'2.strAction:Action Name
										'3.strProperty: property name
										'4.strValue : Property Value

'Return Value		   : 	True Or False

'Pre-requisite			:	Form Property Dialog Should Be Open

'Examples				: 	Call Fn_BMIDERAC_FormPropertyOperations("000004/A","VerifyEditBox","Make/Buy","0")
'										Call  Fn_BMIDERAC_FormPropertyOperations("000004/A","VerifyEditBox","Make/Buy:Name","0:TestForm")
'										Call Fn_BMIDERAC_FormPropertyOperations("Form1","ModifyDropDown","d2CustProp1 / d2CustProp2","DE5345 , AfterGlow")
'										Call Fn_BMIDERAC_FormPropertyOperations("Form1","ModifyDropDown","","DE5345 , AfterGlow")
'										Call Fn_BMIDERAC_FormPropertyOperations("MustPass","ModifyDropDown","d2product1:d2frame1:d2component1","p1 , Prop1ForLOV1:f2 , Prop2ForLOV2:c3 , Prop3ForLOV5")
'										Call Fn_BMIDERAC_FormPropertyOperations("Form1","ModifyDropDownExt","d2product1:d2frame1:d2component1","p1 , Prop1ForLOV1:f2 , Prop2ForLOV2:c3 , Prop3ForLOV5")

'History					 :			
'							Developer Name												Date						Rev. No.						Changes Done												Reviewer
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'							Sandeep N										   				21/12/2010			           1.0																										Sunny R
'							Sandeep N										   				22/12/2010			           1.0					Case "ModifyDropDown"											Sunny R
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'							Pranav Ingle										   			15-May-2013		            1.1					Modified Case "ModifyDropDown"	as TC10.1  		Sunny R
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_BMIDERAC_FormPropertyOperations(strFormName,strAction,strProperty,strValue)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDERAC_FormPropertyOperations"
   'Variable Declaration
   Dim strCurrentVal,arrProperty,arrValue,iCounter,bFlag
   Dim ObjFormDialog
	Dim objTable, objChild, intCount

   'Function Returns False
   Fn_BMIDERAC_FormPropertyOperations=False
   'Changing "title" Property of "FormProperty" Dialog
   JavaWindow("BMIDERACDefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("FormProperty").SetTOProperty "title",strFormName
   'Checking Existance Of "FormProperty" Dialog
	If Fn_UI_ObjectExist("Fn_BMIDERAC_FormPropertyOperations",JavaWindow("BMIDERACDefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("FormProperty"))=True Then
		'Creating Object Of "FormProperty" Dialog
		 Set ObjFormDialog=JavaWindow("BMIDERACDefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("FormProperty")
	Else
		Exit Function
	End If
	Select Case strAction
	'Case To Verify "Property Value of Edit Box"
		Case "VerifyEditBox"
			'Splitting Properties
			arrProperty=Split(strProperty,":")
			'Spliting Values
			arrValue=Split(strValue,":")
			For iCounter=0 To Ubound(arrProperty)
				bFlag=False
				If arrProperty(iCounter)<>"" Then
					'Checking Existance of Property On Form
					Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_BMIDERAC_FormPropertyOperations",ObjFormDialog.JavaEdit("PropertyEdit"),"attached text",arrProperty(iCounter)+".*")
					'Taking Current Property Value
					strCurrentVal=Fn_Edit_Box_GetValue("Fn_BMIDERAC_FormPropertyOperations",ObjFormDialog,"PropertyEdit")
					If LCase(Trim(strCurrentVal))=LCase(Trim(arrValue(iCounter))) Then
						bFlag=True
					End If	
				End If
				If bFlag=False Then
					Exit For
				End If
			Next
			If bFlag=True Then
				Fn_BMIDERAC_FormPropertyOperations=True
			End If
		Case "ModifyDropDown"

				arrProperty=Split(strProperty,":")
				arrValue=Split(strValue,":")
				For iCounter=0 To UBound(arrProperty)

						ObjFormDialog.JavaStaticText("PropertyName").SetToProperty "label",arrProperty(iCounter)+":"
						Call Fn_Button_Click("Fn_BMIDERAC_FormPropertyOperations",ObjFormDialog, "dropdown_16")
						wait 1
	
						Set objTable=Description.Create()
						'objTable("Class Name").value="JavaTable"
						'objTable("tagname").value="LOVTreeTable"
						'Set objChild=ObjFormDialog.ChildObjects(objTable)
						Set objChild=Fn_SISW_UI_Object_GetChildObjects("", ObjFormDialog, "Class Name~toolkit class", "JavaTable~com.teamcenter.rac.common.lov.view.components.LOVTreeTable")
						For intCount=0 to objChild(0).GetROProperty("rows")
							If trim(arrValue(iCounter))=trim(objChild(0).Object.getValueAt(intCount,0).getDisplayableValue()) Then
								objChild(0).DoubleClickCell intCount,0
								Fn_BMIDERAC_FormPropertyOperations = true
								Exit for
							End If
						Next
	
						If intCount= objChild(0).GetROProperty("rows") Then
								Exit Function
						End If
			
				Next
							Call Fn_Button_Click("Fn_BMIDERAC_FormPropertyOperations", ObjFormDialog, "Save")
							wait 2
							Call Fn_Button_Click("Fn_BMIDERAC_FormPropertyOperations", ObjFormDialog, "Close")
							Fn_BMIDERAC_FormPropertyOperations=True

			Case "ModifyDropDownExt"
				Call Fn_Button_Click("Fn_BMIDERAC_FormPropertyOperations",ObjFormDialog, "wizard_16")
				arrProeprty=Split(strProperty,":")
				arrValue=Split(strValue,":")
				For iCounter=0 To UBound(arrValue)
							If strProperty<>"" Then
								Call Fn_List_Select("Fn_BMIDERAC_FormPropertyOperations", ObjFormDialog, "StepPanelManager",arrProeprty(iCounter))
							End If
							Call Fn_List_Select("Fn_BMIDERAC_FormPropertyOperations", ObjFormDialog, "iList",arrValue(iCounter))
				Next
							Call Fn_Button_Click("Fn_BMIDERAC_FormPropertyOperations", ObjFormDialog, "Finish")
							Call Fn_Button_Click("Fn_BMIDERAC_FormPropertyOperations", ObjFormDialog, "Close")
							If Fn_UI_ObjectExist("Fn_BMIDERAC_FormPropertyOperations", JavaWindow("BMIDERACDefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("UnSavedChanges"))=True Then
								Call Fn_Button_Click("Fn_BMIDERAC_FormPropertyOperations", JavaWindow("BMIDERACDefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("UnSavedChanges"), "Yes")	
							End If
							Fn_BMIDERAC_FormPropertyOperations=True

	End Select
	If Fn_UI_ObjectExist("Fn_BMIDERAC_FormPropertyOperations",ObjFormDialog.JavaButton("Cancel"))=True Then
		'Clicking On "Cancel" Button to Close "FormProperty" Dialog
		Call Fn_Button_Click("Fn_BMIDERAC_FormPropertyOperations", ObjFormDialog, "Cancel")
	ElseIf Fn_UI_ObjectExist("Fn_BMIDERAC_FormPropertyOperations",ObjFormDialog.JavaButton("Close"))=True Then 
		Call Fn_Button_Click("Fn_BMIDERAC_FormPropertyOperations", ObjFormDialog, "Close")
	End If
	Set ObjFormDialog=Nothing
End Function
'*********************************************************		Function to verify  dialog error message.	***********************************************************************
'Function Name		:					Fn_BMIDERAC_DialogMsgVerify

'Description			 :		 		  This function is used to get Schedule table Row Index.

'Parameters			   :	 			1.  sTitle:Title of dialog.
'													2. sMsg : Message to verify. (Optional)
'													3. sButton : Button Name.
											
'Return Value		   : 				True/False

'Pre-requisite			:		 		Error Msg should be opened

'Examples				:			  Msgbox Fn_BMIDERAC_DialogMsgVerify("Error","does not have any group member for the given group","OK") 
'											Msgbox Fn_BMIDERAC_DialogMsgVerify("No Group Member","no group member required by the selected work context","OK") 

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Sonal P				27-Dec-2010		1.0														Sandeep
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_BMIDERAC_DialogMsgVerify(sTitle,sMsg,sButton) 
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDERAC_DialogMsgVerify"
   On Error Resume Next
	Dim sResult, diaCreatePref, btnOK, lblMsg, tmp, sErrorMsg
	Fn_BMIDERAC_DialogMsgVerify = True
	' Create Object Description of  Dialog 
	Set diaCreatePref=description.Create()
	diaCreatePref("micclass").value="Dialog"
	diaCreatePref("regexpwndtitle").value = sTitle
	diaCreatePref("regexpwndclass").value = "#32770"
	'Description of  Button Object  on  dialog
	Set btnOK=description.Create()
	btnOK("micclass").value="WinButton"
	btnOK("nativeclass").value = "Button"
	btnOK("regexpwndtitle").value = sButton
	'General Object description to search all Objects
	Set lblMsg=description.Create()
	If Dialog(diaCreatePref).Exist(5) Then
			'Capture All runtime objects to find message text
			Set  tmp = Dialog(diaCreatePref).ChildObjects(lblMsg)
			'Set message text to variable 
			sErrorMsg = tmp(1). getroproperty("text")  
			'compare run time message to verify  the error message
			If sMsg <> "" Then
				'If (sMsg = sErrorMsg ) Then
				If Instr(1,sErrorMsg,sMsg) <> 0 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Error Message Verified Successfully")
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Error Message Verification Failed.")
					Fn_BMIDERAC_DialogMsgVerify = False
				End If
			End If
		' To Click "OK" Button after verification
			wait(2)
			Dialog(diaCreatePref).WinButton(btnOK).Click 10,10,micLeftBtn
			If Dialog(diaCreatePref).Exist(5) Then
				Dialog(diaCreatePref).Close()
			End If
		Else
            Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The " + sTitle + " Dialog does not Exist")
			Fn_BMIDERAC_DialogMsgVerify = False
        End If
	Set diaCreatePref=nothing
	Set btnOK=nothing
	Set lblMsg=nothing
	Set tmp=nothing
End Function
'-------------------------------------------------------------------Function Used to Detail Revise the Object Revision------------------------------------------------------------------
'Function Name		:	Fn_BMIDERAC_ObjectRevRevise

'Description			 :	Function Used to Detail Revise the Object Revision

'Parameters			   :   '1.strItemInfo:Item Information
										'2.strAddItemRevInfo:Additional Revision Information
										'3.strIdentifierInfo: Identifier Basic Information
										'4.strAddIDInfo : Additional ID Information
										'5.strAddRevInfo: Additional Revision Information
										'6.strAttachData : Attach Data
										'7.strAsgnProject: Assign Projects
										'8bDefineOpt: Define Options

'Return Value		   : 	True Or False

'Pre-requisite			:	Should be Log IN TeamCenter

'Examples				: 	Call Fn_BMIDERAC_ObjectRevRevise("RevItemName:ItemDescription","","T2_testsonal:Identifier:1234:A:test:rrrr","","","","","")

'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				29/12/2010			           1.0																				Sunny R
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'strItemInfo :- Item Name:Item Description
'strIdentifierInfo :- Revision:Type:Identifier Name:Identifier Revision:Identifier Name:Identifier Description
Public Function Fn_BMIDERAC_ObjectRevRevise(strItemInfo, strAddItemRevInfo,strIdentifierInfo,strAddIDInfo,strAddRevInfo,strAttachData,strAsgnProject,bDefineOpt)
		GBL_FAILED_FUNCTION_NAME="Fn_BMIDERAC_ObjectRevRevise"
		Dim ObjReviseDialog
		Dim sRevId,arrItemInfo,arrIdentifierInfo
		Fn_BMIDERAC_ObjectRevRevise=False
		Set ObjReviseDialog=JavaWindow("BMIDERACDefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Revise")
		'Select menu [File -> Revise...]
		If Not ObjReviseDialog.Exist(5) Then
				Call Fn_MenuOperation("Select","File:Revise...")
				Call Fn_ReadyStatusSync(2)
		End If
		If strItemInfo<>"" Then
			sRevId =Fn_Edit_Box_GetValue("Fn_BMIDERAC_ObjectRevRevise",ObjReviseDialog,"RevID")
			arrItemInfo=Split(strItemInfo,":")
			If  arrItemInfo(0)<>"" Then
				Call Fn_Edit_Box("Fn_BMIDERAC_ObjectRevRevise",ObjReviseDialog,"Name",arrItemInfo(0))
			End If
			If  arrItemInfo(1)<>"" Then
				Call Fn_Edit_Box("Fn_BMIDERAC_ObjectRevRevise",ObjReviseDialog,"Description",arrItemInfo(1))
			End If
		End If
		If strIdentifierInfo<>"" Then
			ObjReviseDialog.JavaStaticText("Steps").SetTOProperty "label","Enter Identifier Basic Information"
			ObjReviseDialog.JavaStaticText("Steps").Click 1,1,"LEFT"
			wait(2)
			arrIdentifierInfo=Split(strIdentifierInfo,":")
			If arrIdentifierInfo(0)<>"" Then
				Call Fn_List_Select("Fn_BMIDERAC_ObjectRevRevise", ObjReviseDialog, "Revision",arrIdentifierInfo(0))
			End If
			If arrIdentifierInfo(1)<>"" Then
				Call Fn_List_Select("Fn_BMIDERAC_ObjectRevRevise", ObjReviseDialog, "Type",arrIdentifierInfo(1))
			End If
			If arrIdentifierInfo(2)<>"" Then
				Call Fn_Edit_Box("Fn_BMIDERAC_ObjectRevRevise",ObjReviseDialog,"IdentifierID",arrIdentifierInfo(2))
			End If
			If arrIdentifierInfo(3)<>"" Then
				Call Fn_Edit_Box("Fn_BMIDERAC_ObjectRevRevise",ObjReviseDialog,"IdentifierRevision",arrIdentifierInfo(3))
			End If
			If arrIdentifierInfo(4)<>"" Then
				Call Fn_Edit_Box("Fn_BMIDERAC_ObjectRevRevise",ObjReviseDialog,"IdentifierName",arrIdentifierInfo(4))
			End If
			If arrIdentifierInfo(5)<>"" Then
				Call Fn_Edit_Box("Fn_BMIDERAC_ObjectRevRevise",ObjReviseDialog,"IdentifierDescription",arrIdentifierInfo(5))
			End If
		End If
		Call Fn_Button_Click("Fn_BMIDERAC_ObjectRevRevise", ObjReviseDialog, "Finish")
		ObjReviseDialog.JavaButton("Close").WaitProperty "enabled", True
		wait(2)
		Call Fn_Button_Click("Fn_BMIDERAC_ObjectRevRevise", ObjReviseDialog, "Close")
		Fn_BMIDERAC_ObjectRevRevise=sRevId
		Set ObjReviseDialog=Nothing
End Function

'-------------------------------------------------------------------Function Used to Detail Revise the Object Revision------------------------------------------------------------------
'Function Name		:	Fn_BMIDERAC_ItemRevSaveAs

'Description			 :	Function Used to Detail Revise the Object Revision

'Parameters			   :   '1.strItemInfo:Item Information
										'2.strAddItemInfo:Additional Item Information
										'3.strAddItemRevInfo:Additional Revision Information
										'4.strIdentifierInfo: Identifier Basic Information
										'5.strAddIDInfo : Additional ID Information
										'6.strAddRevInfo: Additional Revision Information
										'7.strAttachData : Attach Data
										'8.strAsgnProject: Assign Projects
										'9bDefineOpt: Define Options

'Return Value		   : 	True Or False

'Pre-requisite			:	Should be Log IN TeamCenter

'Examples				: 	Call Fn_BMIDERAC_ItemRevSaveAs("::TestID:TestID Desc:","","","T2_testsonal:Identifier:1873:A:Test:TestDesc","","","","","")

'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				29/12/2010			           1.0																				Sunny R
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'strItemInfo :-Item Id:Item Revision:Item Name:Item Description:UOM
'strIdentifierInfo :-Context:Type:Identifier ID:Identifier Revision:Identifier Name:Identifier Description
Public Function Fn_BMIDERAC_ItemRevSaveAs(strItemInfo,strAddItemInfo,strAddItemRevInfo,strIdentifierInfo,strAddIDInfo,strAddRevInfo,strAttachData,strAsgnProject,bDefineOpt)
			GBL_FAILED_FUNCTION_NAME="Fn_BMIDERAC_ItemRevSaveAs"
			Dim sItemId, sRevId,arrItemInfo,arrIdentifierInfo
			Dim ObjSaveAsDialog
			Set ObjSaveAsDialog=JavaWindow("BMIDERACDefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("SaveAs")
			If Not ObjSaveAsDialog.Exist(5) Then
					Call Fn_MenuOperation("Select","File:Save As")
		   End If
		  If strItemInfo<>"" Then
			 arrItemInfo=Split(strItemInfo,":")
			 If arrItemInfo(0)="" Or  arrItemInfo(1)="" Then
				Call Fn_Button_Click("Fn_BMIDERAC_ItemRevSaveAs", ObjSaveAsDialog,"Assign")
			Else
				Call Fn_Edit_Box("Fn_BMIDERAC_ItemRevSaveAs",ObjSaveAsDialog,"ItemID",arrItemInfo(0))
				Call Fn_Edit_Box("Fn_BMIDERAC_ItemRevSaveAs",ObjSaveAsDialog,"RevisionID",arrItemInfo(1))
			 End If
			 If arrItemInfo(2)<>"" Then
				Call Fn_Edit_Box("Fn_BMIDERAC_ItemRevSaveAs",ObjSaveAsDialog,"ItemName",arrItemInfo(2))
			 End If
			 If arrItemInfo(3)<>"" Then
				Call Fn_Edit_Box("Fn_BMIDERAC_ItemRevSaveAs",ObjSaveAsDialog,"Description",arrItemInfo(3))
			 End If
			If arrItemInfo(4)<>"" Then
				Call Fn_Edit_Box("Fn_BMIDERAC_ItemRevSaveAs",ObjSaveAsDialog,"Unit of Measure",arrItemInfo(4))
			 End If
		  End If
          sItemId =Fn_Edit_Box_GetValue("Fn_BMIDERAC_ItemRevSaveAs",ObjSaveAsDialog,"ItemID")
		  sRevId =Fn_Edit_Box_GetValue("Fn_BMIDERAC_ItemRevSaveAs",ObjSaveAsDialog,"RevisionID")
		  If strIdentifierInfo<>"" Then
			ObjSaveAsDialog.JavaStaticText("Steps").SetTOProperty "label","Enter Identifier Basic Information"
			ObjSaveAsDialog.JavaStaticText("Steps").Click 1,1,"LEFT"
			wait(2)
			arrIdentifierInfo=Split(strIdentifierInfo,":")
			If arrIdentifierInfo(0)<>"" Then
				Call Fn_List_Select("Fn_BMIDERAC_ItemRevSaveAs", ObjSaveAsDialog, "Context",arrIdentifierInfo(0))
			End If
			If arrIdentifierInfo(1)<>"" Then
				Call Fn_List_Select("Fn_BMIDERAC_ItemRevSaveAs", ObjSaveAsDialog, "Type",arrIdentifierInfo(1))
			End If
			If arrIdentifierInfo(2)<>"" Then
				Call Fn_Edit_Box("Fn_BMIDERAC_ItemRevSaveAs",ObjSaveAsDialog,"IdentifierID",arrIdentifierInfo(2))
			End If
			If arrIdentifierInfo(3)<>"" Then
				Call Fn_Edit_Box("Fn_BMIDERAC_ItemRevSaveAs",ObjSaveAsDialog,"IdentifierRevision",arrIdentifierInfo(3))
			End If
			If arrIdentifierInfo(4)<>"" Then
				Call Fn_Edit_Box("Fn_BMIDERAC_ItemRevSaveAs",ObjSaveAsDialog,"IdentifierName",arrIdentifierInfo(4))
			End If
			If arrIdentifierInfo(5)<>"" Then
				Call Fn_Edit_Box("Fn_BMIDERAC_ItemRevSaveAs",ObjSaveAsDialog,"IdentifierDescription",arrIdentifierInfo(5))
			End If
		  End If
		  wait(2)
		 ObjSaveAsDialog.JavaButton("Finish").WaitProperty "enabled", 1, 20000
		'Click on "Finish" button
		 If Cint(ObjSaveAsDialog.JavaButton("Finish").GetROProperty("enabled")) = 1 Then
			Call Fn_Button_Click("Fn_BMIDERAC_ItemRevSaveAs", ObjSaveAsDialog, "Finish")
		 Else
			Fn_BMIDERAC_ItemRevSaveAs = False
		 End If
		Call Fn_Button_Click("Fn_BMIDERAC_ItemRevSaveAs", ObjSaveAsDialog, "Close")
		Fn_BMIDERAC_ItemRevSaveAs = sItemId & "-" & sRevId
		Set ObjSaveAsDialog=Nothing
End Function
'-------------------------------------------------------------------Function to Use to Perform Operations On Identifier Option------------------------------------------------------------------------------------
'Function Name		:				Fn_BMIDERAC_IdentifierOptionsSettings

'Description			 :		 		 Function to Use to Perform Operations On Identifier Option

'Parameters			   :			1. StrAction: Action to be performed
'												  2. strContextLength: Context Length
'												  3. strAttachOptn: Context Separater

'Return Value		   : 				TRUE \ FALSE

'Pre-requisite			:		 		Should be Log In Team center

'Examples				:		Call Fn_BMIDERAC_IdentifierOptionsSettings("Modify","3","*")
'
'History					 :		
'													Developer Name				Date						Rev. No.			Changes Done						Reviewer
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'														Sandeep N					03-Jan-11				1.0																				Sunny R
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_BMIDERAC_IdentifierOptionsSettings(strAction,strContextLength,strAttachOptn)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDERAC_IdentifierOptionsSettings"
   Dim ObjOptnDialog
   Fn_BMIDERAC_IdentifierOptionsSettings=False
	Set ObjOptnDialog = JavaWindow("BMIDERACDefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Options")
	If Not ObjOptnDialog.Exist(10) Then
			Call Fn_MenuOperation("Select","Edit:Options...")
	End If    
	Call Fn_ReadyStatusSync(3)
	wait(3)
	Select Case strAction
		Case "Modify"
			Call Fn_JavaTree_Select("Fn_BMIDERAC_IdentifierOptionsSettings", ObjOptnDialog, "OptionsTree","Options:General:Identifier")	
			If strContextLength<>"" Then
				Call Fn_Edit_Box("Fn_BMIDERAC_IdentifierOptionsSettings",ObjOptnDialog,"ContextLength",strContextLength)
			End If
			If strAttachOptn<>"" Then
				Call Fn_Edit_Box("Fn_BMIDERAC_IdentifierOptionsSettings",ObjOptnDialog,"ContextAttach",strAttachOptn)
			End If
			Call Fn_Button_Click("Fn_BMIDERAC_IdentifierOptionsSettings", ObjOptnDialog, "OK")
			Fn_BMIDERAC_IdentifierOptionsSettings=True
	End Select
End Function

'-------------------------------------------------------------------Function to Use to Create Requirement For Design-------------------------------------------------------------------------
'Function Name		:				Fn_BMIDERAC_RequirementForDesignBasicCreate

'Description			 :		 		Function to Use to Create Requirement For Design

'Parameters			   :			1. strReqType: Requirement Design Type
'												  2. strID: ID
'												  3. strRev: Revision
'												4. strName:Name
'												5. strDesc:Description
'												6. strUOM:Unit Of Measure
'												7. strVariables :Variables

'Return Value		   : 				TRUE \ FALSE

'Pre-requisite			:		 		Should be Log In Team center 

'Examples				:		Call Fn_BMIDERAC_RequirementForDesignBasicCreate("S3DesignReq","","","TestDesign","TestDesign Requirement","","Test1:String:::TestDesc:123")
'
'History					 :		
'													Developer Name				Date						Rev. No.			Changes Done						Reviewer
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'														Sandeep N					05-Jan-11				1.0																				Sunny R
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'strVariables :- "Name:Type:Measure:Unit:Description:Value~Name:Type:Measure:Unit:Description:Value"
'Example :-"Test1:String:::TestDesc:123"
Public Function Fn_BMIDERAC_RequirementForDesignBasicCreate(strReqType,strID,strRev,strName,strDesc,strUOM,strVariables)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDERAC_RequirementForDesignBasicCreate"
   Dim bFlag,strCurrID,strCurrRev,arrSetVariable,iCounter,arrValueVariable
   Dim ObjDesignDialog,ObjDialogChild,ObjEdit,WshShell
   bFlag=False
   Fn_BMIDERAC_RequirementForDesignBasicCreate=False
	Set ObjDesignDialog=JavaWindow("BMIDERACDefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("NewRequirementForDesign")
	If Not ObjDesignDialog.Exist(10) Then
		'Select menu ["File -> New -> Requirement for Design..."]
		Call Fn_MenuOperation("Select","File:New:Requirement for Design...")
	End If
	bFlag=Fn_UI_ListItemExist("Fn_BMIDERAC_RequirementForDesignBasicCreate", ObjDesignDialog, "RequirementDesignList",strReqType)
	If bFlag=False Then
		Call Fn_Button_Click("Fn_BMIDERAC_RequirementForDesignBasicCreate", ObjDesignDialog, "Close")
		Set ObjDesignDialog=Nothing
		Exit Function
	End If
	Call Fn_List_Select("Fn_BMIDERAC_RequirementForDesignBasicCreate", ObjDesignDialog, "RequirementDesignList",strReqType)
	ObjDesignDialog.JavaButton("Next").WaitProperty "enabled",1,20000
	Call Fn_Button_Click("Fn_BMIDERAC_RequirementForDesignBasicCreate", ObjDesignDialog, "Next")
	If strID="" Or  strRev="" Then
		Call Fn_Button_Click("Fn_BMIDERAC_RequirementForDesignBasicCreate", ObjDesignDialog, "Assign")
	Else
		Call Fn_Edit_Box("Fn_BMIDERAC_RequirementForDesignBasicCreate",ObjDesignDialog,"ID",strID)
		Call Fn_Edit_Box("Fn_BMIDERAC_RequirementForDesignBasicCreate",ObjDesignDialog,"Revision",strRev)
	End If
	If strName<>"" Then
		Call Fn_Edit_Box("Fn_BMIDERAC_RequirementForDesignBasicCreate",ObjDesignDialog,"Name",strName)
	End If
	If strDesc<>"" Then
		Call Fn_Edit_Box("Fn_BMIDERAC_RequirementForDesignBasicCreate",ObjDesignDialog,"Description",strDesc)
	End If
	strCurrID=Fn_Edit_Box_GetValue("Fn_BMIDERAC_RequirementForDesignBasicCreate",ObjDesignDialog,"ID")
	strCurrRev=Fn_Edit_Box_GetValue("Fn_BMIDERAC_RequirementForDesignBasicCreate",ObjDesignDialog,"Revision")

	If strVariables<>"" Then
		ObjDesignDialog.JavaButton("Next").WaitProperty "enabled",1,20000
		Call Fn_Button_Click("Fn_BMIDERAC_RequirementForDesignBasicCreate", ObjDesignDialog, "Next")
		arrSetVariable=Split(strVariables,"~")
		For iCounter=0 To Ubound(arrSetVariable)
			Call Fn_Button_Click("Fn_BMIDERAC_RequirementForDesignBasicCreate", ObjDesignDialog, "Add")
			arrValueVariable=Split(arrSetVariable(iCounter),":")
			If arrValueVariable(0)<>"" Then
				ObjDesignDialog.JavaTable("DesignTable").SetCellData iCounter,"Name",arrValueVariable(0)
			End If
			If arrValueVariable(1)<>"" Then
				ObjDesignDialog.JavaTable("DesignTable").ClickCell iCounter,"Type","LEFT"
				ObjDesignDialog.JavaList("TypeList").Select arrValueVariable(1)
			End If
			If arrValueVariable(2)<>"" Then
				ObjDesignDialog.JavaTable("DesignTable").ClickCell iCounter,"Measure","LEFT"
				ObjDesignDialog.JavaList("TypeList").Select arrValueVariable(2)
			End If
			If arrValueVariable(3)<>"" Then
				
			End If
			If arrValueVariable(4)<>"" Then
				ObjDesignDialog.JavaTable("DesignTable").SetCellData iCounter,"Description",arrValueVariable(4)
			End If
			If arrValueVariable(5)<>"" Then
				ObjDesignDialog.JavaTable("DesignTable").SetCellData iCounter,"Value",arrValueVariable(5)			
			End If
		Next
	End If
	Call Fn_Button_Click("Fn_BMIDERAC_RequirementForDesignBasicCreate", ObjDesignDialog, "Finish")
	wait(5)
	If ObjDesignDialog.Exist(10) Then
		Call Fn_Button_Click("Fn_BMIDERAC_RequirementForDesignBasicCreate", ObjDesignDialog, "Close")
	End If
	Fn_BMIDERAC_RequirementForDesignBasicCreate=strCurrID+"/"+strCurrRev
	Set ObjDesignDialog=Nothing
End Function
'-------------------------------------------------------------------Function to Use to Perform Operations On Operational Data Option------------------------------------------------------------------
'Function Name		:				Fn_BMIDERAC_OperationDataOptionOperation

'Description			 :		 		Function to Use to Perform Operations On Operational Data Option

'Parameters			   :			1. StrAction: Action to be performed
'												  2. strOpsData: Operational Data Element
'												  3. strMsg: Operational Data Element Description

'Return Value		   : 				TRUE \ FALSE\Item Count

'Pre-requisite			:		 		Should be Log In Team center

'Examples				:		Msgbox Fn_BMIDERAC_OperationDataOptionOperation("Add","Alias ID Rule","")
										'Msgbox Fn_BMIDERAC_OperationDataOptionOperation("Add","Alias ID Rule:Change","")
										'Msgbox Fn_BMIDERAC_OperationDataOptionOperation("Remove","Alias ID Rule:Change","")
										'Msgbox Fn_BMIDERAC_OperationDataOptionOperation("RemoveAll","","")
										'Msgbox Fn_BMIDERAC_OperationDataOptionOperation("AddAll","","")
										'Msgbox Fn_BMIDERAC_OperationDataOptionOperation("VerifyMessage","Change","A change")
										'Msgbox Fn_BMIDERAC_OperationDataOptionOperation("VerifyAvailableElements","Change","")
										'Msgbox Fn_BMIDERAC_OperationDataOptionOperation("VerifyActiveElements","Change","")
										'Msgbox Fn_BMIDERAC_OperationDataOptionOperation("AvailableElementsCount","","")
										'Msgbox Fn_BMIDERAC_OperationDataOptionOperation("ActiveElementsCount","","")
'
'History					 :		
'													Developer Name				Date						Rev. No.			Changes Done						Reviewer
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'														Sandeep N					24-Jan-11				1.0																				Sunny R
'														Sandeep N					18-Nov-11				1.1			Modified Node from "Options:Oprations Data" to "Options:Live Update"																	Sunny R
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_BMIDERAC_OperationDataOptionOperation(strAction,strOpsData,strMsg)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDERAC_OperationDataOptionOperation"
   Dim ObjOptionDialog
   Dim arrOpsData,iCounter,strCurrValue,bFlag
   Fn_BMIDERAC_OperationDataOptionOperation=False
	Set ObjOptionDialog=JavaWindow("BMIDERACDefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Options")
	If Not ObjOptionDialog.Exist(8) Then
        'Select menu ["Edit-.Options...]
		Call Fn_MenuOperation("Select","Edit:Options...")
	End If
	Call Fn_ReadyStatusSync(3)
	wait(3)
    Call Fn_JavaTree_Select("Fn_BMIDERAC_OperationDataOptionOperation", ObjOptionDialog, "OptionsTree","Options:Live Update")	
	Select Case strAction
		Case "Add"
			arrOpsData=Split(strOpsData,":")
			For iCounter=0 To UBound(arrOpsData)
				'Call Fn_List_Select("Fn_BMIDERAC_OperationDataOptionOperation",ObjOptionDialog, "ActiveOpsDataList",arrOpsData(iCounter))
				ObjOptionDialog.JavaList("ActiveOpsDataList").Select arrOpsData(iCounter)
				wait 1
				Call Fn_Button_Click("Fn_BMIDERAC_OperationDataOptionOperation", ObjOptionDialog, "Add")
			Next
			Fn_BMIDERAC_OperationDataOptionOperation=True

		Case "Remove"
			arrOpsData=Split(strOpsData,":")
			For iCounter=0 To UBound(arrOpsData)
				'Call Fn_List_Select("Fn_BMIDERAC_OperationDataOptionOperation",ObjOptionDialog, "AvailableOpsDataList",arrOpsData(iCounter))
				ObjOptionDialog.JavaList("AvailableOpsDataList").Select arrOpsData(iCounter)
				wait 1
				Call Fn_Button_Click("Fn_BMIDERAC_OperationDataOptionOperation", ObjOptionDialog, "Remove")
			Next
			Fn_BMIDERAC_OperationDataOptionOperation=True

		Case "AddAll"
			Call Fn_Button_Click("Fn_BMIDERAC_OperationDataOptionOperation", ObjOptionDialog, "AddAll")
			Fn_BMIDERAC_OperationDataOptionOperation=True

		Case "RemoveAll"
			Call Fn_Button_Click("Fn_BMIDERAC_OperationDataOptionOperation", ObjOptionDialog, "RemoveAll")
			Fn_BMIDERAC_OperationDataOptionOperation=True

		Case "VerifyMessage"
				'Call Fn_List_Select("Fn_BMIDERAC_OperationDataOptionOperation",ObjOptionDialog, "ActiveOpsDataList",strOpsData)
				ObjOptionDialog.JavaList("ActiveOpsDataList").Select strOpsData
				wait 1
				strCurrValue=Fn_UI_Object_GetROProperty("Fn_BMIDERAC_OperationDataOptionOperation",ObjOptionDialog.JavaEdit("ActiveOpsDataDesc"), "value")
				If InStr(1,LCase(Trim(strCurrValue)),Lcase(Trim(strMsg)))>0 Then
					Fn_BMIDERAC_OperationDataOptionOperation=True
				End If
			
		Case "VerifyAvailableElements"
			arrOpsData=Split(strOpsData,":")
			For iCounter=0 To UBound(arrOpsData)
				bFlag=False
				bFlag=Fn_UI_ListItemExist("Fn_BMIDERAC_OperationDataOptionOperation",ObjOptionDialog, "AvailableOpsDataList",arrOpsData(iCounter))
				If bFlag=False Then
					Set ObjOptionDialog=Nothing
					Exit Function
				End If
			Next
			Fn_BMIDERAC_OperationDataOptionOperation=True

		Case "VerifyActiveElements"
			arrOpsData=Split(strOpsData,":")
			For iCounter=0 To UBound(arrOpsData)
				bFlag=False
				bFlag=Fn_UI_ListItemExist("Fn_BMIDERAC_OperationDataOptionOperation",ObjOptionDialog, "ActiveOpsDataList",arrOpsData(iCounter))
				If bFlag=False Then
					Set ObjOptionDialog=Nothing
					Exit Function
				End If
			Next
			Fn_BMIDERAC_OperationDataOptionOperation=True

		Case "AvailableElementsCount"
			Fn_BMIDERAC_OperationDataOptionOperation=Fn_UI_Object_GetROProperty("Fn_BMIDERAC_OperationDataOptionOperation",ObjOptionDialog.JavaList("AvailableOpsDataList"), "items count")

		Case "ActiveElementsCount"
			Fn_BMIDERAC_OperationDataOptionOperation=Fn_UI_Object_GetROProperty("Fn_BMIDERAC_OperationDataOptionOperation",ObjOptionDialog.JavaList("ActiveOpsDataList"), "items count")
	End Select

	Call Fn_Button_Click("Fn_BMIDERAC_OperationDataOptionOperation", ObjOptionDialog, "OK")
	Call Fn_ReadyStatusSync(2)
	Set ObjOptionDialog=Nothing
End Function

'-------------------------------------------------------------------Function Used to Perform Operations on Properties Dialog----------------------------------------------------------------------
'Function Name		:	Fn_BMIDERAC_PropertiesOperations

'Description			 :	Function Used to Verify Objects Properties

'Parameters			   :   '1.strAction:Action Name
										'2.strControlName:It could be Attached Text,Label ect
										'3.strValue: Expected Properties Values

'Return Value		   : 	True Or False Or Value

'Pre-requisite			:	Properties Dialog Should Be Open

'Examples				: 	Call Fn_BMIDERAC_PropertiesOperations("getEditBoxValue","Version Number:","")

'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				23/02/2010			           1.0																				Sunny R
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'strControlName:- It could be Attached Text,Label ect
Public Function Fn_BMIDERAC_PropertiesOperations(strAction,strControlName,strValue)
	GBL_FAILED_FUNCTION_NAME="Fn_BMIDERAC_PropertiesOperations"
	Fn_BMIDERAC_PropertiesOperations=False
   Set ObjPropDialog=JavaWindow("BMIDERACDefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Properties")
	If Not ObjPropDialog.Exist(5) Then
		Set ObjPropDialog=Nothing
		Exit Function
	End If
	Select Case strAction
	 	Case "getEditBoxValue"
			ObjPropDialog.JavaEdit("Name").SetTOProperty "attached text",strControlName
			Fn_BMIDERAC_PropertiesOperations=ObjPropDialog.JavaEdit("Name").GetROProperty("value")
            Call Fn_Button_Click("Fn_BMIDERAC_PropertiesOperations",ObjPropDialog,"Cancel")
	End Select
	Set ObjPropDialog=Nothing
End Function



''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''''/$$$$
''''/$$$$   FUNCTION NAME   :  Fn_BMIDERAC_NewBusinessObjectSync()
''''/$$$$
''''/$$$$   DESCRIPTION        :  This function will handle synchronization for creating Business Objects
''''/$$$$  
''''/$$$$   PRE-REQUISITES        :  New Business Object Window should be present after clicking the Finish Button
''''/$$$$
''''/$$$$  PARAMETERS   : 		No Parameters
''''/$$$$	
''''/$$$$		Return Value : 				True or False
''''/$$$$
''''/$$$$    Function Calls       :    No External Function Calls Used
''''/$$$$ 
''''/$$$$		HISTORY           :   		AUTHOR                 DATE        VERSION
''''/$$$$
''''/$$$$    CREATED BY     :   SHREYAS          02/04/2012         1.0
''''/$$$$
''''/$$$$    REVIWED BY     :  Shreyas			02/04/2012           1.0
''''/$$$$
''''/$$$$		How To Use :   bReturn=Fn_BMIDERAC_NewBusinessObjectSync()
''''/$$$$							
''''/$$$$	
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

Public Function Fn_BMIDERAC_NewBusinessObjectSync()
   Fn_BMIDERAC_NewBusinessObjectSync=false
		While JavaWindow("BMIDERACDefaultWindow").JavaWindow("NewBusinessObject").Exist=false
			wait 3
			If JavaWindow("BMIDERACDefaultWindow").JavaWindow("NewBusinessObject").Exist=false then
				wait 1
			End if
		Wend
		wait 1
		If JavaWindow("BMIDERACDefaultWindow").JavaWindow("NewBusinessObject").JavaButton("Cancel").Exist then
			JavaWindow("BMIDERACDefaultWindow").JavaWindow("NewBusinessObject").JavaButton("Cancel").Click micLeftBtn
		End if
		wait 1
		Fn_BMIDERAC_NewBusinessObjectSync=true
End Function
