Option Explicit
'*********************************************************	Function List		***********************************************************************
'1. Fn_SISW_Prop_GetObject
'2. Fn_SISW_Prop_ObjPropEdit
'3. Fn_SISW_Prop_ObjectPropertyVerify
'4. Fn_SISW_Prop_EditCancelCkOut
'5. Fn_SISW_Prop_CkOut_Edit_CkIn
'6. Fn_SISW_Prop_ObjectPropertyCkOut_Edit_CkIn
'7. Fn_SISW_Prop_ObjectProperty_Edit
'8. Fn_SISW_Prop_CkOut_Explore_Edit_CkIn
'9. Fn_SISW_Prop_HTML_PropertyVerify
'10. Fn_SISW_Prop_ObjectPropertyOperation
'11. Fn_SISW_Prop_ObjectPropertyCkOutEditSave
'12. Fn_SISW_Prop_HTML_PropertyRetrieve
'13. Fn_SISW_Prop_ObjectPropertyIsEditable
'14. Fn_SISW_Prop_CommonModifiableProperties_Operation
'15. Fn_SISW_Prop_PropertiesOnRelation_Operations
'16. Fn_SISW_Prop_Text_PropertyVerify
'17.  Fn_SISW_Prop_ObjPropertiesOperation
'18. Fn_SISW_Prop_VerifyProperties
'19. Fn_SISW_Prop_EditProperties
'*********************************************************	Function List		***********************************************************************


''****************************************    Function to get Object hierarchy ***************************************
'
''Function Name		 	:	Fn_SISW_Prop_GetObject
'
''Description		    :  	Function to get specified Object hierarchy.

''Parameters		    :	1. sObjectName : Object Handle name
								
''Return Value		    :  	Object \ Nothing
'
''Examples		     	:	Fn_SISW_Prop_GetObject("TcDefaultApplet")

'History:
'	Developer Name			Date			Rev. No.		Reviewer		Changes Done	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Sushma Pagare		 30-July-2013		1.0				
'   -----------------------------------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------------------------------
'	Anurag Khera		 26-Sept-2013		2.0								Externalized object hierarchies
'-----------------------------------------------------------------------------------------------------------------------------------
Function Fn_SISW_Prop_GetObject(sObjectName)
	Dim sObjectXMLPath
	sObjectXMLPath = Fn_GetEnvValue("User", "AutomationDir") & "\TestData\AutomationXML\ObjectXML\Property.xml"
	Set Fn_SISW_Prop_GetObject = Fn_SISW_Setup_GetObjectFromXML(sObjectXMLPath, sObjectName)
End Function
'*********************************************************		Edits some of the object properties		***********************************************************************
'Function Name		:				Fn_SISW_Prop_ObjPropEdit

'Description			 :		 		 This function edits some of the object properties (such as Configuration Item)

'Parameters			   :               				1.StrName :
'													2.StrDescription:
'													3.StrUrlname:
'													4.StrUOM:
'													5.StrVersionLimit:
'													6. BlnConfigItem:
'													7.StrProjectID:
'													8.StrSerialNo:


'Return Value		   : 				PASS/ FAIL

'Pre-requisite			:		 		Object to edit properties is selected.

'Examples				:                call Fn_SignalBasicCreate("Signal", "False", "", "", "Test Auto Signal 2", "Test Auto Desc 2", "")

'History:
'										Developer Name			Date			Rev. No.			Changes Done			Reviewer	 Revoew Date
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Sameer					29-Mar-2010		1.0											Santosh			29-Mar-10					
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Koustubh W				16-Aug-2011		1.0					Added code to check existence of Properties Dialog aftger performing menu opration
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Sandeep N				04-Jul-2012		1.1					Added code to handle New object Hierarchy : JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Properties")
'																													JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Edit Properties")
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function Fn_SISW_Prop_ObjPropEdit(sName, sDescription, sUrlname, sUOM, sVersionLimit, bConfigItem, sProjectID, sSerialNo)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_Prop_ObjPropEdit"
	On Error Resume Next
	Fn_SISW_Prop_ObjPropEdit = True
	Dim objEditProperties, objCheckIn, objCheckOutDialog, objPropertiesTC
	Dim bProprtyDialog,bEditProperties,bProprtyDialog1,bEditProperties1

	Dim StrTitle
	StrTitle=JavaWindow("DefaultWindow").GetROProperty("title")

	bProprtyDialog=JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Properties").Exist(SISW_MIN_TIMEOUT)
	If bProprtyDialog=False Then
		bEditProperties=JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Edit Properties").Exist(SISW_MIN_TIMEOUT)
	Else
		bEditProperties=False
	End If

	bProprtyDialog1=JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Properties").Exist(SISW_MICRO_TIMEOUT)
	If bProprtyDialog1=False Then
		bEditProperties1=JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Edit Properties").Exist(SISW_MICRO_TIMEOUT)
	Else
		bEditProperties1=False
	End If
 	' Object created for "Properties"  & "Edit Properties" Dialog

	'Checks whether the "Properties" or "Edit Property"Dialog is displayed
	If NOT (bProprtyDialog OR bProprtyDialog1)Then
		Call Fn_MenuOperation("Select","View:Properties")
		Call Fn_ReadyStatusSync(1)
		bProprtyDialog=JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Properties").Exist(5)
		bEditProperties=JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Edit Properties").Exist(1)

		bProprtyDialog1=JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Properties").Exist(1)
		bEditProperties1=JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Edit Properties").Exist(1)
	End If
 	
	'Checks whether the "Properties" or "Edit Property"Dialog is displayed
	If NOT (bProprtyDialog = True OR bEditProperties= True or bProprtyDialog1=True or bEditProperties1=true )Then
		Fn_SISW_Prop_ObjPropEdit = False
		'Call Fn_WriteLogFile("Fn_SISW_Prop_ObjPropEdit", 1, Err.Number,"FAIL: Could not open Property or Edit Property Dialog")
		Exit Function
	'Checks whether the "Properties" Dialog is displayed
	ElseIf bProprtyDialog = True Then
			Set objPropertiesTC = Fn_UI_ObjectCreate("Fn_SISW_Prop_ObjPropEdit",JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Properties"))
			If objPropertiesTC.JavaButton("Check-Out and Edit").Exist(1) Then
				Call Fn_Button_Click("Fn_SISW_Prop_ObjPropEdit",objPropertiesTC,"Check-Out and Edit")
			End If
	
			Set objCheckOutDialog = Fn_SISW_GetChkInChkOutObject("CheckOut")
			If TypeName(objCheckOutDialog) <> "Nothing" Then
				Call Fn_Button_Click("Fn_SISW_Prop_ObjPropEdit",objCheckOutDialog,"Yes")
			End If
			Call Fn_ReadyStatusSync(3)

	Elseif bProprtyDialog1 = True Then
			Set objPropertiesTC = Fn_UI_ObjectCreate("Fn_SISW_Prop_ObjPropEdit",JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Properties"))
			If objPropertiesTC.JavaButton("Check-Out and Edit").Exist(1) Then
				Call Fn_Button_Click("Fn_SISW_Prop_ObjPropEdit",objPropertiesTC,"Check-Out and Edit")
			End If

			Set objCheckOutDialog = Fn_SISW_GetChkInChkOutObject("CheckOut")
			If TypeName(objCheckOutDialog) <> "Nothing" Then
				Call Fn_Button_Click("Fn_SISW_Prop_ObjPropEdit",objCheckOutDialog,"Yes")
			End If
			Call Fn_ReadyStatusSync(3)
	Else			
			Fn_SISW_Prop_ObjPropEdit = False
			Call Fn_WriteLogFile("Fn_SISW_Prop_ObjPropEdit", 1, Err.Number,"FAIL: Properties Dialog is not displayed")
			Exit Function
	End If
	'Checking Existance of [ Edit Properties ] dialog
	If JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Edit Properties").Exist(5) Then
		Set objEditProperties = Fn_UI_ObjectCreate("Fn_SISW_Prop_ObjPropEdit",JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Edit Properties"))
	elseif JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Edit Properties").Exist(1) then
		Set objEditProperties = Fn_UI_ObjectCreate("Fn_SISW_Prop_ObjPropEdit",JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Edit Properties"))
	elseif objPropertiesTC.Exist(1) then
		Set objEditProperties = objPropertiesTC
	else
		Fn_SISW_Prop_ObjPropEdit = False
		'Call Fn_WriteLogFile("Fn_SISW_Prop_ObjPropEdit", 1, Err.Number,"FAIL: Check out Dialog is not displayed")
		Exit Function
	End If

	Call Fn_UI_JavaStaticText_SetTOProperty("Fn_SISW_Prop_ObjPropEdit",objEditProperties,"BottomLink","label", "All")
	'objEditProperties.JavaStaticText("BottomLink").Click 1,1,"LEFT"
	'If Fn_Java_StaticText_Exist("Fn_SISW_Prop_ObjPropEdit",objEditProperties, "BottomLink") Then
	If Fn_SISW_UI_Object_Operations("Fn_SISW_Prop_ObjPropEdit","Exist", objEditProperties.JavaStaticText("BottomLink"), SISW_MINLESS_TIMEOUT) then
		Call Fn_UI_JavaStaticText_Click("Fn_SISW_Prop_ObjPropEdit",objEditProperties,"BottomLink",1,1,"LEFT")
	End If
	Call Fn_ReadyStatusSync(1)

	'If Fn_Java_StaticText_Exist("Fn_SISW_Prop_ObjPropEdit",objEditProperties, "More...") Then
	If Fn_SISW_UI_Object_Operations("Fn_SISW_Prop_ObjPropEdit","Exist", objEditProperties.JavaStaticText("More..."), SISW_MICRO_TIMEOUT) then
    		Call Fn_UI_JavaStaticText_Click("Fn_SISW_Prop_ObjPropEdit",objEditProperties,"More...",1,1,  "LEFT")
	End If

	' Enters specified data
	If sName <> "" Then
	Call Fn_Edit_Box("Fn_SISW_Prop_ObjPropEdit",objEditProperties,"Name",sName)
	End if

	If sDescription <> "" Then
			Call Fn_Edit_Box("Fn_SISW_Prop_ObjPropEdit",objEditProperties,"Description",sDescription)
	End if
	If sUrlname <> "" Then
			Call Fn_Edit_Box("Fn_SISW_Prop_ObjPropEdit",objEditProperties,"URL",sUrlname)
	End if

	If sUOM <> "" Then
		Call Fn_Button_Click("Fn_SISW_Prop_ObjPropEdit",objEditProperties,"UOMDropDown")
		Call Fn_UI_JavaStaticText_Click("Fn_SISW_Prop_ObjPropEdit",objEditProperties,"label:=" & sUOM,1,1,"LEFT")
	End If

	If sVersionLimit <> "" Then
		'objEditProperties.JavaEdit("VersionLimit").Set sVersionLimit
		'Implementation
		Call Fn_Edit_Box("Fn_SISW_Prop_ObjPropEdit",objEditProperties,"VersionLimit",sVersionLimit)
	End if

	If bConfigItem = "True" Then
		Call Fn_UI_Object_SetTOProperty("Fn_SISW_Prop_ObjPropEdit",objEditProperties.JavaRadioButton("ConfigItem"),"label", "True")
        'objEditProperties.JavaRadioButton("ConfigItem").Set "ON"
		'Implementation
		Call  Fn_UI_JavaRadioButton_SetON("Fn_SISW_Prop_ObjPropEdit",objEditProperties,"ConfigItem","ON")
	ElseIf bConfigItem = "False" Then
        Call Fn_UI_Object_SetTOProperty("Fn_SISW_Prop_ObjPropEdit",objEditProperties.JavaRadioButton("ConfigItem"),"label", "False")
		'objEditProperties.JavaRadioButton("ConfigItem").Set "OFF"
		'Implementation
		Call Fn_UI_JavaRadioButtont_setOff("Fn_SISW_Prop_ObjPropEdit",objEditProperties,"ConfigItem")
	End If
	
	' Clicks on "Check In" button
	'objEditProperties.JavaButton("Check-In").Click
           'implemantation
	'Call Fn_Button_Click("Fn_SISW_Prop_ObjPropEdit",objEditProperties,"Check-In")
     If objEditProperties.JavaButton("SaveAndCheck-In").Exist(5) Then
			Call Fn_Button_Click("Fn_SISW_Prop_ObjPropEdit",objEditProperties,"SaveAndCheck-In")
	End If
	Call Fn_ReadyStatusSync(1)

	Set objCheckIn = Fn_SISW_GetChkInChkOutObject("CheckIn")
	If TypeName(objCheckIn) <> "Nothing" Then
		Call Fn_Button_Click("Fn_SISW_Prop_ObjPropEdit",objCheckIn,"Yes")
	End If


'	'Checks if any error has occured.
'	If Err.Number <> 0 Then
'		Fn_SISW_Prop_ObjPropEdit =False
'		'Call Fn_WriteLogFile("Fn_SISW_Prop_ObjPropEdit", 1, Err.Number,"FAIL: Coudn't edit object properties")
'		Err.clear
'	End If
'
'	On Error Goto 0
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Successfully Edited Object property")
	' Clear out memory allocated for objects
	Set objEditProperties = Nothing
	Set objCheckOutDialog  = Nothing
	Set objCheckIn = Nothing
	Set objPropertiesTC = Nothing
End Function

'#######################################################################################
'###    FUNCTION NAME   :   Fn_SISW_Prop_ObjectPropertyVerify(sPropertyName,sPropertyVal)
'###
'###    DESCRIPTION     :   Verify the Obejtc Properties usign PRINT Dialog
'###
'###    PARAMETERS      :   1.sPropertyName: (:) seperated list of properties to check.
'###                                             2.sPropertyVal: (:) Seperated  lsit of Values
'###                                             
'###    Return Value   :   The String which represents the result : "PASS" or "FAIL" with the reason
'###
'###    Pre-Requisites : Print Window is open
'###
'###    HISTORY         :   AUTHOR              	DATE        VERSION
'###
'###    CREATED BY      :   Amol    				30/03/2010   	1.0
'###
'###    REVIWED BY      :	Santosh				 30/03/2010 	1.0 	
'###
'###    MODIFIED BY     :
'###    EXAMPLE         :   Fn_SISW_Prop_ObjectPropertyVerify("Group ID:Last Modifying User","dba:infodba")
'#######################################################################################
Function Fn_SISW_Prop_ObjectPropertyVerify(sPropertyName,sPropertyVal)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_Prop_ObjectPropertyVerify"
On Error Resume Next
Dim arrPropertyName,arrPropertyVal,arrButtonName,ObjVerify,Rwcnt,objTable,i,j

'*********************************Split the number of parameters into an array**************************************
arrPropertyName = Split(sPropertyName,":",-1)
arrPropertyVal = Split(sPropertyVal,":",-1)
'*************************************Checking that the Browser is opened or not************************************
If  Browser("browser").Page("Page").Exist(SISW_MIN_TIMEOUT) Then
	'************************Here we count the Number of rows present in that perticular table**************************
	Rwcnt = Browser("Browser").Page("Page").WebTable("PropertyTable").RowCount
	Set objTable = Browser("Browser").Page("Page").WebTable("PropertyTable")
	'*****Here is the loop for checking the first colume value if matches then it goes to the second if  condition*******
	For i = 0 to Ubound(arrPropertyName)
		  For j = 0 to Rwcnt
				If Trim(objTable.GetCellData(j,1)) = arrPropertyName(i) Then
						If Trim(objTable.GetCellData(j,2)) =arrPropertyVal(i) Then
							'Call Fn_WriteLogFile("Fn_SISW_Prop_ObjectPropertyVerify", 1, Err.Number,"PASS:Object Property Verified for + [ + arrPropertyVal(i) + ]")        																																													
							  Fn_SISW_Prop_ObjectPropertyVerify = True
							   Exit For
						 End If
					 End If
			Next
		Next
End If							
					
'*****************************************************************For Error Log File*************************************************************************
If Err. Number <> 0 Then
		Fn_SISW_Prop_ObjectPropertyVerify = False
		'Call Fn_WriteLogFile("Fn_SISW_Prop_ObjectPropertyVerify", 1, Err.Number,"FAIL:Can't complete Property Verify")
		Err.clear
Else
		Fn_SISW_Prop_ObjectPropertyVerify = True
		'Call Fn_WriteLogFile("Fn_SISW_Prop_ObjectPropertyVerify", 1, Err.Number,"PASS:Property Verify complete successfully.")
End If
Set objTable = Nothing
Set ObjVerify = Nothing
End Function

'#=======================================================================================================================
'#
'# FUNCTION NAME :	Fn_SISW_Prop_EditCancelCkOut(sObjectProperty, sObjectPropertyValue)
'#
'# MODULE : My Teamcener
'#
'# DESIGN REQUIREMENT BY : 	Mallikarjun
'#
'# PROGRAMMED BY :	Pallavi Patil
'#
'# CREATION DATE : 		3/05/2010
'#
'# 
'# PRE-REQUISITE :	1) My TeamCenter Window should be Open.
'#										2) One Item,Item Revision,Dataset,Form should be created.
'#									   3) Node should be selected.
'#
'# DESCRIPTION : Edit the properties and Cancel Checkout Changes of a given business object  such as below:
'#									1. Item
'#									2. Item Revision
'#									3. Dataset
'#									4. Form
'#
'#
'#FUNCTIONS INTERNAL CALLS :  1) Fn_MenuOperation
'#																	2)Fn_Button_Click
'#																	3) Fn_Edit_Box
'#
'# PARAMETERS : sPropertyName : Name of the property to be edited
'#									sPropertyValue : Value of the property to be edited
'#		
'#	Example:	Call Fn_SISW_Prop_EditCancelCkOut("Name","Ths is my  New item.................")
'#
'#  Programmer	 |      Date			|  Version	    |	Description
'# SQS		            |  3-May-2010	|	 1.0		      |Created
'#====================================================================================================================================

Public Function Fn_SISW_Prop_EditCancelCkOut(sObjectProperty, sObjectPropertyValue)
		GBL_FAILED_FUNCTION_NAME="Fn_SISW_Prop_EditCancelCkOut"
		Dim bReturn,ObjPropertyWindow,ObjPropertyChange,ObjPropertyCancelCheckOut, sJavaList, sFunctionName

		'Checking Window Exist or not	
		bReturn=Fn_UI_ObjectExist("Fn_SISW_Prop_EditCancelCkOut",JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Properties")) 
 		If bReturn=false Then
				'Calling Menu Operation...
    			Call Fn_MenuOperation("KeyPress", "View:Properties")
		End If

		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "View Properties option get Selected " & sJavaList & " of Function " & sFunctionName)
		Fn_SISW_Prop_EditCancelCkOut=True

		'Pressing the Ckeck-Out and Edit Button
		Set ObjPropertyWindow =Fn_UI_ObjectCreate("Fn_SISW_Prop_EditCancelCkOut", JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Properties"))
		Call Fn_Button_Click(" Fn_SISW_Prop_EditCancelCkOut",ObjPropertyWindow,"Check-Out and Edit")
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Selected " & sJavaList & " of Function " & sFunctionName)
		Fn_SISW_Prop_EditCancelCkOut=True

		'Set the Check-Out JavaDialog object 
		Set ObjPropertyWindow = Fn_SISW_GetChkInChkOutObject("CheckOut")
		If TypeName(ObjPropertyWindow) <> "Nothing" Then
			Call Fn_Button_Click(" Fn_SISW_Prop_EditCancelCkOut",ObjPropertyWindow,"Yes")
		End If
		Set ObjPropertyWindow = Nothing

		'Checking Edit Properties Window exist or not
		Set ObjPropertyWindow =Fn_UI_ObjectCreate("Fn_SISW_Prop_EditCancelCkOut",JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Edit Properties"))
		Select Case sObjectProperty 
				Case "Name" 
						Call Fn_Edit_Box("Fn_SISW_Prop_EditCancelCkOut",ObjPropertyWindow,"Name",sObjectPropertyValue)
				Case "Description"
						Call Fn_Edit_Box("Fn_SISW_Prop_EditCancelCkOut",ObjPropertyWindow,"Description",sObjectPropertyValue)
		End Select
								
		'Pressing the Cancel Check-Out button
		Call Fn_Button_Click(" Fn_SISW_Prop_EditCancelCkOut",ObjPropertyWindow,"Cancel Check-Out")
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Selected " & sJavaList & " of Function " & sFunctionName)
		Fn_SISW_Prop_EditCancelCkOut=True
		Set ObjPropertyWindow = Nothing	
		
		Set ObjPropertyWindow = Fn_SISW_GetChkInChkOutObject("CancelCheckOut")
		If TypeName(ObjPropertyWindow) <> "Nothing" Then
			Call Fn_Button_Click(" Fn_SISW_Prop_EditCancelCkOut",ObjPropertyWindow,"Yes")
		End If
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Selected " & sJavaList & " of Function " & sFunctionName)
		Fn_SISW_Prop_EditCancelCkOut=True		

    	Set ObjPropertyWindow=Nothing
		Set ObjPropertyChange=Nothing
		Set ObjPropertyCancelCheckOut=Nothing

End Function

'######################################################################################################################################
'###
'###    FUNCTION NAME   :  Fn_SISW_Prop_CkOut_Edit_CkIn(sObjProperty,sObjPropertyValue)
'###
'###    DESCRIPTION     : Checkout Edit and Checkin operation for business object in My Teamcenter Navigator Tree.
'###
'###    PARAMETERS      : 1. sObjProperty-Valid Object Property Name
'###										2.sObjPropertyValue-Valid Object Property Value
'###                                         
'###    Function Calls  :  Fn_WriteLogFile()
'###									Fn_Button_Click()
'###									 Fn_UI_JavaList_ExtendSelect()
'###									 Fn_MenuOperation()
'###									Fn_UI_ObjectExist()
'###									 Fn_Edit_Box()
'###	 HISTORY         :   		AUTHOR                 DATE        VERSION
'###
'###    CREATED BY      :  Dhananjay       		  28/04/2010         1.0
'###
'###    REVIWED BY      :   							 		28/04/2010	    1.0
'###
'###    MODIFIED BY     :
'###    EXAMPLE         :   Fn_SISW_Prop_CkOut_Edit_CkIn("Description","Object Ready to Check-In")
'######################################################################################################################################
Function Fn_SISW_Prop_CkOut_Edit_CkIn(sObjProperty,sObjPropertyValue)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_Prop_CkOut_Edit_CkIn"
	Dim objEditProperty,objEditProperty1 , objCheckIn, objChkOut

	Set objEditProperty = Fn_SISW_GetObject("Edit Properties")
	Set objEditProperty1 = Fn_SISW_GetObject("Edit Properties@1")
    
	Set objChkOut = Fn_SISW_GetChkInChkOutObject("CheckOut")
	If Typename(objChkOut) <> "Nothing"Then		
			Call Fn_Button_Click("Fn_SISW_Prop_CkOut_Edit_CkIn",objChkOut,"Yes")
	end If
	Set objChkOut = Nothing
	'Check if Edit Properties Dialog Exist
	If not Fn_UI_ObjectExist("Fn_SISW_Prop_CkOut_Edit_CkIn",objEditProperty) AND not Fn_UI_ObjectExist("Fn_SISW_Prop_CkOut_Edit_CkIn",objEditProperty1) Then
			Call Fn_MenuOperation("Select","Edit:Properties")
			Call Fn_ReadyStatusSync(2)
			Set objChkOut = Fn_SISW_GetChkInChkOutObject("CheckOut")
			If Typename(objChkOut) <> "Nothing"Then
					Call Fn_Button_Click("Fn_SISW_Prop_CkOut_Edit_CkIn",objChkOut,"Yes")
			end If
			Set objChkOut = Nothing
	End If
		
	'Check Existence of Edit Property Dailog
		If objEditProperty.Exist(5) Then
			Set objEditProperty = Fn_SISW_GetObject("Edit Properties")
		ElseIf objEditProperty1.Exist(1) Then
			Set objEditProperty = Fn_SISW_GetObject("Edit Properties@1")
		Else
			Fn_SISW_Prop_CkOut_Edit_CkIn = False
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Edit Properties dialog do not exits" )
			Exit Function
		End If
		
		'Click on Static text
		'Call Fn_UI_JavaStaticText_Click(" Fn_TcObjectDelete", objEditProperty, "BottomLink", 1, 1, "LEFT")
        ' Set the Property Value
		Call Fn_Edit_Box("Fn_SISW_Prop_CkOut_Edit_CkIn",objEditProperty,sObjProperty,sObjPropertyValue)
		
		'Click on 'SaveAndCheck-In' Button
		Call Fn_Button_Click("Fn_SISW_Prop_CkOut_Edit_CkIn",objEditProperty,"SaveAndCheck-In")
		
		'Click on 'Yes' Button
		Set objCheckIn = Fn_SISW_GetChkInChkOutObject("CheckIn")
		If Typename(objCheckIn) <> "Nothing"Then	
				Call Fn_Button_Click("Fn_SISW_Prop_CkOut_Edit_CkIn",objCheckIn,"Yes")
		end If
		Fn_SISW_Prop_CkOut_Edit_CkIn = True
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS:Property '"&sObjProperty&"' Successfully Modified and changed to "&sObjPropertyValue&" of Function Fn_SISW_Prop_CkOut_Edit_CkIn" )
	
	Set objEditProperty = Nothing
	Set objEditProperty1=Nothing
	Set objCheckIn = Nothing

End Function

'#########################################################################################################
'###
'###    FUNCTION NAME   :   Fn_SISW_Prop_ObjectPropertyCkOut_Edit_CkIn(sObjectType, oItemPropDictonary)
'### 
'###    DESCRIPTION        :   Pre-requisite:
'###						
'###							1. Context My Teamcenter
'###							2. Object Should be selected 
'###						

'###
'###    PARAMETERS      :      1. sObjectType
'###						   2. oItemPropDictonary
'###									    
'###									  
'###
'###	 HISTORY       :   		 AUTHOR                 DATE        VERSION
'###
'###    CREATED BY     :   		Prasanna           		20/05/2010         1.0
'###
'###    REVIWED BY     :   	  Harshal Agrawal			20-May-10			1.0
'###
'###    MODIFIED BY   :  
'###    EXAMPLE          : Call Fn_SISW_Prop_ObjectPropertyCkOut_Edit_CkIn("Item",oItemPropDictonary)
'###    How to use in Test						
'		1. Data Dictionary Declaration 
'					Dim oItemPropDictonary
'					Set oItemPropDictonary  = CreateObject( "Scripting.Dictionary" )
'						With oItemPropDictonary
'         						.Add "Name", ""
''            					.Add "Description", ""     
''            					.Add "Configuration", ""
''								.Add "Revision","" 
'						End with

'		2. Use in Test :							
'							If oItemPropDictonary.Exists("Name") then 
'								oItemPropDictonary.Remove("Name")
'								oItemPropDictonary.Add "Name","Changed3"
'							Else
'								oItemPropDictonary.Add "Name","Changed3"
'							End if 
'							If oItemPropDictonary.Exists("Description") then 
'								oItemPropDictonary.Remove("Description")
'								oItemPropDictonary.Add "Description","description changed "
'							Else
'								oItemPropDictonary.Add "Description","Item REV description changed"
'							End if 
'
'
'							Call Fn_ObjectPropertyCkoutEditCkIn("BOMLine",oItemPropDictonary)            
'#############################################################################################################

'Modified by Pooja B -   Modified  for the  IPClassification         
'
'Date Modified   - 10-12-12

'Modified by Sandeep N -   Modified  Case  IPClassification  : modification made as per 10.1 design change
'#############################################################################################################
Function Fn_SISW_Prop_ObjectPropertyCkOut_Edit_CkIn(sObjectType, oItemPropDictonary)
		GBL_FAILED_FUNCTION_NAME="Fn_SISW_Prop_ObjectPropertyCkOut_Edit_CkIn"
		Dim objCheckOutDia,  oPropSelect, dictItems, dictKeys, i, sProperty, sObjElement,  ObjPropertyWindow, ObjPropertyYes,objStat,objDialog
		Dim bFlag,objTable,objChild,iCounter

		Set ObjPropertyWindow=JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Edit Properties")
  'Added by Nilesh on 12-Feb-2013
			If ObjPropertyWindow.Exist(SISW_MIN_TIMEOUT)=False Then
				oPropSelect = Fn_MenuOperation("Select", "Edit:Properties")
				Call Fn_ReadyStatusSync(3)				
			End If

		Fn_SISW_Prop_ObjectPropertyCkOut_Edit_CkIn = False
					Select Case sObjectType
							Case "Item" ,"ItemExt"
                                        'Use menu operations to open the Properties window.
'										oPropSelect = Fn_MenuOperation("Select", "Edit:Properties")                                        
										'Set the Check out Dialog Object
										Set objCheckOutDia = Fn_SISW_GetChkInChkOutObject("CheckOut")
										If TypeName(objCheckOutDia) <> "Nothing" Then
											Call Fn_Button_Click("Fn_SISW_Prop_ObjPropEdit",objCheckOutDia,"Yes")
											Call Fn_ReadyStatusSync(3)
										End If
										Set objCheckOutDia = Nothing
										
										'Set Object
										If JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Edit Properties").Exist(6) Then
											Set ObjPropertyWindow = JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Edit Properties")
										ElseIf JavaWindow("DefaultWindow").JavaWindow("DefaultEmbeddedFrame").JavaDialog("Edit Properties").Exist(1) Then
											Set ObjPropertyWindow = JavaWindow("DefaultWindow").JavaWindow("DefaultEmbeddedFrame").JavaDialog("Edit Properties")
										End If
										
										For i=0 to 10
											If ObjPropertyWindow.Exist Then
												Exit For
											Else
												wait(10)
											End If
										Next
    									if ObjPropertyWindow.Exist Then
    										If ObjPropertyWindow.JavaButton("Check-OutAndEdit").Exist Then
    											Call Fn_SISW_UI_JavaButton_Operations("Fn_SISW_Prop_ObjPropEdit", "Click", ObjPropertyWindow,"Check-OutAndEdit")
												Call Fn_ReadyStatusSync(3)
												Set objCheckOutDia = Fn_SISW_GetChkInChkOutObject("CheckOut")
												If TypeName(objCheckOutDia) <> "Nothing" Then
													Call Fn_Button_Click("Fn_SISW_Prop_ObjPropEdit",objCheckOutDia,"Yes")
													Call Fn_ReadyStatusSync(3)
												End If
    										End If
										'Get the keys & items count from data dictionary.	
										dictItems = oItemPropDictonary.Items
										dictKeys = oItemPropDictonary.Keys									

										For i = 0 to oItemPropDictonary.Count - 1
												   If IsNull(dictKeys(i))  Then
													Else
															If  dictitems(i) = "" Then										
															else
																		sObjElement = dictKeys(i)
																		' Set the value as per the data dictioanry key.
																		select case sObjElement
																				Case  "Name" , "Description","IsFastTrack","RecurringCost"
																							ObjPropertyWindow.JavaStaticText("BottomLink").SetTOProperty "label", "All"
																							Wait(1)
																							ObjPropertyWindow.JavaStaticText("BottomLink").Click 10, 10,"LEFT"
																							Wait(1)
																							ObjPropertyWindow.JavaStaticText("BottomLink").SetTOProperty "label", "Show empty properties..."
																							Wait(1)
																							'Code modified by Chandrakant Tyagi to check existence of BottomLink
																							If ObjPropertyWindow.JavaStaticText("BottomLink").exist(1) Then
																								ObjPropertyWindow.JavaStaticText("BottomLink").Click 10, 10,"LEFT"
																							End If
																							
																							If sObjElement = "Description" Then '[TC1121-2015101900-02_11_2015-VivekA-Maintenace] - Added to get focus on Description edit box
																								ObjPropertyWindow.JavaStaticText("BottomLink").SetTOProperty "label", "Hide empty properties..."
																								Wait 1
																								ObjPropertyWindow.JavaStaticText("BottomLink").Highlight
																								Wait 1
																								ObjPropertyWindow.JavaStaticText("BottomLink").SetTOProperty "label", "Show empty properties..."
																							End If
																							'[TC1123-20161010-24_10_2016-VivekA-Maintenace] - Property name changed from "Is Fast Track?" to "Fast Track:" - By Piyush P
																							If sObjElement = "IsFastTrack" Then
																								ObjPropertyWindow.JavaEdit(sObjElement).SetTOProperty "attached text", "Fast Track:"
																							End If
																							'--------------------------------------------------
																							If sObjectType = "ItemExt" Then
																							  Call Fn_SISW_UI_JavaEdit_Operations("Fn_SISW_Prop_ObjectPropertyCkOut_Edit_CkIn","Set",ObjPropertyWindow,sObjElement,dictItems(i))	
																							Else
																							 Call Fn_SISW_UI_JavaEdit_Operations("Fn_SISW_Prop_ObjectPropertyCkOut_Edit_CkIn", "SetExt",ObjPropertyWindow,sObjElement,dictItems(i))
																							End If
																							wait(1)
																							Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: The property "& sObjElement  & " changed successfully")   																	
																			'Added Case ID By Pallavi Jadhav
																				Case "ID"	
																							ObjPropertyWindow.JavaStaticText("BottomLink").SetTOProperty "label", "All"
																							Wait(1)
																							ObjPropertyWindow.JavaStaticText("BottomLink").Click 10, 10,"LEFT"
																							Wait(1)
																							ObjPropertyWindow.JavaEdit(sObjElement).set dictItems(i)
																				Case "Contract Pricing Model"
																							wait(2)
																							ObjPropertyWindow.JavaStaticText("BottomLink").Click 1,1
																							wait(2)
																							ObjPropertyWindow.JavaStaticText("More...").Click 1,1
																							wait(3)
																							'Call Fn_Edit_Box("Fn_SISW_Prop_ObjectPropertyCkOut_Edit_CkIn",ObjPropertyWindow,sObjElement,dictItems(i))
																							Call Fn_UI_EditBox_Type("Fn_SISW_Prop_ObjectPropertyCkOut_Edit_CkIn",ObjPropertyWindow,sObjElement,dictItems(i))
																							wait(2)
                                                                                            Call Fn_KeyBoardOperation("SendKeys", "{DOWN}~{ENTER}")
																							wait(2)
																							Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: The property "& sObjElement  & " changed successfully")
																				Case  "GovClassification", "Note Text"
																							ObjPropertyWindow.JavaStaticText("BottomLink").Click 1,1
																							If Fn_SISW_UI_Object_Operations("Fn_ObjectPropertyCkOut_Edit_CkIn","Exist", ObjPropertyWindow.JavaStaticText("More..."),"") Then			'By Ankit N_15July2015_Tc11.2_2015070100_Modified Library by checking the existance of "Show Empty Properties..." 
																								ObjPropertyWindow.JavaStaticText("More...").Click 1,1
																							End If 																							
																							wait(3)
																							Call Fn_Edit_Box("Fn_SISW_Prop_ObjectPropertyCkOut_Edit_CkIn",ObjPropertyWindow,sObjElement,dictItems(i))
																							Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: The property "& sObjElement  & " changed successfully")
																				Case "IPClassification"
																							ObjPropertyWindow.JavaStaticText("BottomLink").SetTOProperty "label", "All"
																							Wait(1)
																							ObjPropertyWindow.JavaStaticText("BottomLink").Click 10, 10,"LEFT"
																							Wait(1)
																							ObjPropertyWindow.JavaStaticText("BottomLink").SetTOProperty "label", "Show empty properties..."
																							wait 1																							
																							If Fn_SISW_UI_Object_Operations("Fn_ObjectPropertyCkOut_Edit_CkIn","Exist", ObjPropertyWindow.JavaStaticText("BottomLink"),"") Then				'By Ankit N_15July2015_Tc11.2_2015070100_Modified Library by checking the existance of "Show Empty Properties..." 
																								ObjPropertyWindow.JavaStaticText("BottomLink").Click 10, 10,"LEFT"
																							End If
																							Wait(1)
																							'ObjPropertyWindow.JavaStaticText("BottomLink").Click 10, 10,"LEFT"

																							ObjPropertyWindow.JavaStaticText("PropertyLabel").SetTOProperty "label","IP Classification:"
																							Call Fn_Button_Click(" Fn_SISW_Prop_ObjectPropertyCkOut_Edit_CkIn",ObjPropertyWindow,"DropdownButton")
																							Wait(1)
																							bFlag=False
																							Set objTable=Description.Create()					'	modification made as per TC Buiild-11.1 (20140402)  
																							objTable("Class Name").value="JavaTable"
																							'objTable("tagname").value="LOVTreeTable"
																							objTable("toolkit class").value="com.teamcenter.rac.common.lov.view.components.LOVTreeTable"
																							Set objChild=ObjPropertyWindow.ChildObjects(objTable)
																							For iCounter=0 to objChild(0).GetROProperty("rows")-1
																								If trim(objChild(0).Object.getValueAt(iCounter,0).getDisplayableValue())=trim(dictItems(i)) Then
																									objChild(0).DoubleClickCell iCounter,0
																									bFlag=true
																									Exit for
																								End If
																							Next
																							Set objTable=Nothing
																							Set objChild=Nothing
                                                                                Case "OriginalLocationCode"
																					      
																							ObjPropertyWindow.JavaStaticText("BottomLink").SetTOProperty "label", "All"
																							Wait(1)
																							ObjPropertyWindow.JavaStaticText("BottomLink").Click 10, 10,"LEFT"
																							Wait(1)
																							ObjPropertyWindow.JavaStaticText("BottomLink").SetTOProperty "label", "Show empty properties..."
																							Wait(1)
																							ObjPropertyWindow.JavaStaticText("BottomLink").Click 10, 10,"LEFT"

																							ObjPropertyWindow.JavaStaticText("PropertyLabel").SetTOProperty "label","Original Location Code:"
																							Call Fn_Edit_Box("Fn_SISW_Prop_ObjectPropertyCkOut_Edit_CkIn",ObjPropertyWindow,"Property_Edit",dictItems(i))
																							wait(1)
																							Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: The property "& sObjElement  & " changed successfully")
																							Wait(1)
																				End select
															End If
													End If                              
										Next
										'Click on Save and check-in button.
										Wait(3)
										Call Fn_Button_Click(" Fn_SISW_Prop_ObjectPropertyCkOut_Edit_CkIn",ObjPropertyWindow,"SaveAndCheck-In")
										wait(3)
										Call Fn_ReadyStatusSync(10)
'                                       
										Set ObjPropertyYes = Fn_SISW_GetChkInChkOutObject("CheckIn")
										If TypeName(ObjPropertyYes) <> "Nothing" Then
											Call Fn_Button_Click("Fn_SISW_Prop_ObjPropEdit",ObjPropertyYes,"Yes")
										End If
										Set ObjPropertyYes = Nothing
										wait(3)
										Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: The item Saved and checked in successfully")										
										Fn_SISW_Prop_ObjectPropertyCkOut_Edit_CkIn = True
									 end if
					Case "Item Revision"
								'Use menu operations to open the Properties window.
'                                oPropSelect = Fn_MenuOperation("Select", "Edit:Properties")  
								'Set the Check Out JavaDialog Object
								Set objCheckOutDia = Fn_SISW_GetChkInChkOutObject("CheckOut")
								If TypeName(objCheckOutDia) <> "Nothing" Then
									Call Fn_Button_Click("Fn_SISW_Prop_ObjPropEdit",objCheckOutDia,"Yes")
									Call Fn_WriteLogFile(Environment.Value("TestLogFile")  ,"PASS: The "& sObjectType  & " checked out  successfully")	
								End If
								Set objCheckOutDia = Nothing

								If JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Edit Properties").Exist Then
										If JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Edit Properties").JavaStaticText("BottomLink").Exist = True Then
											JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Edit Properties").JavaStaticText("BottomLink").Click 1,1,"LEFT"
										End If

										If JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Edit Properties").JavaStaticText("More...").Exist = True Then
											JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Edit Properties").JavaStaticText("More...").Click 1,1,"RIGHT"
										End If

                                        	'Get the keys & items count from data dictionary.	
										dictItems = oItemPropDictonary.Items
										dictKeys = oItemPropDictonary.Keys
									    For i = 0 to oItemPropDictonary.Count - 1
										If IsNull(dictKeys(i))  Then
										Else
												If  dictitems(i) = "" Then										
												else
                                                            sObjElement = dictKeys(i)
															Set ObjPropertyWindow = Fn_UI_ObjectCreate("Fn_SISW_Prop_ObjectPropertyCkOut_Edit_CkIn",JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Edit Properties"))
                                                            ' Set the value as per the data dictioanry key.
															select case sObjElement
																	Case  "Name" , "Description"             
																				Call Fn_Edit_Box("Fn_SISW_Prop_ObjectPropertyCkOut_Edit_CkIn",ObjPropertyWindow,sObjElement,dictItems(i))
																				Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: The property "& sObjElement  & " changed successfully")	
																			
															End select 
												End If
										End If										
										Next
											'Click on Save and check-in button.
											Call Fn_Button_Click(" Fn_SISW_Prop_ObjectPropertyCkOut_Edit_CkIn",ObjPropertyWindow,"SaveAndCheck-In")
											wait(2)
											Set ObjPropertyYes = Fn_SISW_GetChkInChkOutObject("CheckIn")
											If TypeName(ObjPropertyYes) <> "Nothing" Then
												Call Fn_Button_Click("Fn_SISW_Prop_ObjPropEdit",ObjPropertyYes,"Yes")
												Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: The Item Revision saved and checked in successfully")
											End If
											Set ObjPropertyYes = Nothing
											Fn_SISW_Prop_ObjectPropertyCkOut_Edit_CkIn = True
									end if
					Case "Dataset"
                            	 'Use menu operations to open the Properties window.
'								oPropSelect = Fn_MenuOperation("Select", "Edit:Properties")  
									
								Set objCheckOutDia = Fn_SISW_GetChkInChkOutObject("CheckOut")
								If TypeName(objCheckOutDia) <> "Nothing" Then
									Call Fn_Button_Click("Fn_SISW_Prop_ObjPropEdit",objCheckOutDia,"Yes")
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: The "& sObjectType  & " checked out  successfully")
									Call Fn_ReadyStatusSync(1)									
								End If
								Set objCheckOutDia = Nothing

								If JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Edit Properties").Exist Then
									 Set ObjPropertyWindow=JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Edit Properties")
									 If ObjPropertyWindow.JavaStaticText("More...").Exist(1) Then
									 	 ObjPropertyWindow.JavaStaticText("More...").Click 1,1
									 End If
								ElseIf JavaWindow("DefaultWindow").JavaWindow("DefaultEmbeddedFrame").JavaDialog("Edit Properties").Exist(1) Then
									 Set ObjPropertyWindow=JavaWindow("DefaultWindow").JavaWindow("DefaultEmbeddedFrame").JavaDialog("Edit Properties")
										 If ObjPropertyWindow.JavaStaticText("More...").Exist(1) Then
									 	 ObjPropertyWindow.JavaStaticText("More...").Click 1,1
									 End If
								 End If 
								 
								If ObjPropertyWindow.Exist(SISW_MIN_TIMEOUT)=True Then
										'Get the keys & items count from data dictionary.	
										dictItems = oItemPropDictonary.Items
										dictKeys = oItemPropDictonary.Keys
									    For i = 0 to oItemPropDictonary.Count - 1
										If IsNull(dictKeys(i))  Then
										Else
											If  dictitems(i) = "" Then										
												else
                                                            sObjElement = dictKeys(i)
'															Set ObjPropertyWindow = Fn_UI_ObjectCreate("Fn_SISW_Prop_ObjectPropertyCkOut_Edit_CkIn",JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Edit Properties"))

                                                            ' Set the value as per the data dictioanry key.

															select case sObjElement
																	Case  "Name" , "VersionLimit", "Description", "GovClassification"             
																				Call Fn_Edit_Box("Fn_SISW_Prop_ObjectPropertyCkOut_Edit_CkIn",ObjPropertyWindow,sObjElement,dictItems(i))
																				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: The property "& sObjElement  & " changed successfully")
																	Case "IPClassification"
																				ObjPropertyWindow.JavaStaticText("BottomLink").SetTOProperty "label", "Show empty properties..."
																				If ObjPropertyWindow.JavaStaticText("BottomLink").Exist(SISW_MICRO_TIMEOUT) Then
																					ObjPropertyWindow.JavaStaticText("BottomLink").Click 10, 10,"LEFT"
																				End If

																				Call Fn_Button_Click("Fn_SISW_Prop_ObjPropEdit",ObjPropertyWindow,"DropdownButton")
'																				ObjPropertyWindow.JavaEdit("GovClassification").SetTOProperty "attached text","IP Classification:"
'																				Wait(1)
'																				ObjPropertyWindow.JavaEdit("GovClassification").Type dictItems(i)	
																				Wait(1)
																				bFlag=False
																				Set objTable=Description.Create()
																				objTable("Class Name").value="JavaTable"
																				objTable("tagname").value="LOVTreeTable"
																				Set objChild=ObjPropertyWindow.ChildObjects(objTable)
																				For iCounter=0 to objChild(0).GetROProperty("rows")-1
																					If trim(objChild(0).Object.getValueAt(iCounter,0).getDisplayableValue())=trim(dictItems(i)) Then
																						objChild(0).DoubleClickCell iCounter,0
																						bFlag=true
																						Exit for
																					End If
																				Next
																				Set objTable=Nothing
																				Set objChild=Nothing
															End select 
												End If
										End If									
										Next
										'Click on Save and check-in button.
										Call Fn_Button_Click(" Fn_SISW_Prop_ObjectPropertyCkOut_Edit_CkIn",ObjPropertyWindow,"SaveAndCheck-In")
										Wait 2
										Set ObjPropertyYes = Fn_SISW_GetChkInChkOutObject("CheckIn")
										If TypeName(ObjPropertyYes) <> "Nothing" Then
											Call Fn_Button_Click("Fn_SISW_Prop_ObjPropEdit",ObjPropertyYes,"Yes")
											Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: The Item Revision saved and checked in successfully")
										End If
										Set ObjPropertyYes = Nothing
										Fn_SISW_Prop_ObjectPropertyCkOut_Edit_CkIn = True
									end if
					Case "Form"
                           	 '	Use menu operations to open the Properties window.
'								oPropSelect = Fn_MenuOperation("Select", "Edit:Properties") 
								Set objCheckOutDia = Fn_SISW_GetChkInChkOutObject("CheckOut")
								If TypeName(objCheckOutDia) <> "Nothing" Then
									Call Fn_Button_Click("Fn_SISW_Prop_ObjPropEdit",objCheckOutDia,"Yes")
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: The "& sObjectType  & " checked out  successfully")	
								End If
								Set objCheckOutDia = Nothing

								if JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Edit Properties").Exist Then
									'	Get the keys & items count from data dictionary.

										dictItems = oItemPropDictonary.Items
										dictKeys = oItemPropDictonary.Keys
									    For i = 0 to oItemPropDictonary.Count - 1
										If IsNull(dictKeys(i))  Then
										Else
												If  dictitems(i) = "" Then										
												else
                                                            sObjElement = dictKeys(i)
															Set ObjPropertyWindow = Fn_UI_ObjectCreate("Fn_SISW_Prop_ObjectPropertyCkOut_Edit_CkIn",JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Edit Properties"))

                                                            ' Set the value as per the data dictioanry key.

															select case sObjElement
																	Case  "Name", "Description"             
																				Call Fn_Edit_Box("Fn_SISW_Prop_ObjectPropertyCkOut_Edit_CkIn",ObjPropertyWindow,sObjElement,dictItems(i))
																				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: The property "& sObjElement  & " changed successfully")	
															End select 
												End If
										End If										
                                        Next
										'Click on Save and check-in button.
										Call Fn_ReadyStatusSync(1)
										Call Fn_Button_Click(" Fn_SISW_Prop_ObjectPropertyCkOut_Edit_CkIn",ObjPropertyWindow,"SaveAndCheck-In")
										Call Fn_ReadyStatusSync(1)

										Set ObjPropertyYes = Fn_SISW_GetChkInChkOutObject("CheckIn")
										If TypeName(ObjPropertyYes) <> "Nothing" Then
											Call Fn_Button_Click("Fn_SISW_Prop_ObjPropEdit",ObjPropertyYes,"Yes")
											Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: The Item Revision saved and checked in successfully")
										End If
										Set ObjPropertyYes = Nothing
										Fn_SISW_Prop_ObjectPropertyCkOut_Edit_CkIn = True
									end if
			Case "Schedule"
                            	'	Use menu operations to open the Properties window.
'								oPropSelect = Fn_MenuOperation("Select", "Edit:Properties")  
						
								Set objCheckOutDia = Fn_SISW_GetChkInChkOutObject("CheckOut")
								If TypeName(objCheckOutDia) <> "Nothing" Then
									Call Fn_Button_Click("Fn_SISW_Prop_ObjPropEdit",objCheckOutDia,"Yes")
									Call Fn_ReadyStatusSync(5)
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: The "& sObjectType  & " checked out  successfully")	
								End If
								Set objCheckOutDia = Nothing
								if JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Edit Properties").Exist Then
							  'Get the keys & items count from data dictionary.
										dictItems = oItemPropDictonary.Items
										dictKeys = oItemPropDictonary.Keys
									    For i = 0 to oItemPropDictonary.Count - 1
										If IsNull(dictKeys(i))  Then
										Else
												If  dictitems(i) = "" Then										
												else
                                                            sObjElement = dictKeys(i)
															Set ObjPropertyWindow = Fn_UI_ObjectCreate("Fn_SISW_Prop_ObjectPropertyCkOut_Edit_CkIn",JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Edit Properties"))
															' Set the value as per the data dictioanry key.
															select case sObjElement
																	Case  "Name" , "Description"             
																				Call Fn_Edit_Box("Fn_SISW_Prop_ObjectPropertyCkOut_Edit_CkIn",ObjPropertyWindow,sObjElement,dictItems(i))
																				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: The property "& sObjElement  & " changed successfully")	
															End select 
												End If
										End If									
                                        Next
										'Click on Save and check-in button.
										Call Fn_Button_Click(" Fn_SISW_Prop_ObjectPropertyCkOut_Edit_CkIn",ObjPropertyWindow,"SaveAndCheck-In")
										Call Fn_ReadyStatusSync(1)

										Set ObjPropertyYes = Fn_SISW_GetChkInChkOutObject("CheckIn")
										If TypeName(ObjPropertyYes) <> "Nothing" Then
											Call Fn_Button_Click("Fn_SISW_Prop_ObjPropEdit",ObjPropertyYes,"Yes")
											Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: The Item Revision saved and checked in successfully")
										End If
										Set ObjPropertyYes = Nothing
										Fn_SISW_Prop_ObjectPropertyCkOut_Edit_CkIn = True
									end if
        					Case "BOMLine"
								'Use menu operations to open the Properties window.
								call Fn_MenuOperation("KeyPress","View:Properties")		
								if JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Properties").Exist Then
										dictItems = oItemPropDictonary.Items
										dictKeys = oItemPropDictonary.Keys
									    For i = 0 to oItemPropDictonary.Count - 1
										'Check the keys and Items value from datadictionary.
										If IsNull(dictKeys(i))  Then
										Else
												If  dictitems(i) = "" Then										
												else
														sObjElement = dictKeys(i)
														Set ObjPropertyWindow = Fn_UI_ObjectCreate("Fn_SISW_Prop_ObjectPropertyCkOut_Edit_CkIn",JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Properties"))
														select case sObjElement
										' Set the value as per the data dictioanry key.																									
																		case "RevDescription"
																						Call Fn_Edit_Box("Fn_SISW_Prop_ObjectPropertyCkOut_Edit_CkIn", ObjPropertyWindow, sObjElement , dictItems(i))
																						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: The property "& sObjElement  & " changed successfully")	
                                                        End select 
												  End If
										End If	
										Next
									'Click the OK button to save the properties edited.
                                    Call Fn_Button_Click(" Fn_SISW_Prop_ObjectPropertyCkOut_Edit_CkIn",ObjPropertyWindow,"OK")
                                    Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: The BOMLine saved  successfully")
									JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Check-In").Activate
									JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Check-In").JavaButton("Yes").Click
									Fn_SISW_Prop_ObjectPropertyCkOut_Edit_CkIn = True
						end if
				Case "URL"
						 '	Use menu operations to open the Properties window.
						 Set objCheckOutDia = Fn_SISW_GetChkInChkOutObject("CheckOut")
						 If TypeName(objCheckOutDia) = "Nothing" Then
							bReturn = Fn_MenuOperation("Select", "Edit:Properties")
						 End If
						 Set objCheckOutDia = Nothing

						 Set objCheckOutDia = Fn_SISW_GetChkInChkOutObject("CheckOut")
						If TypeName(objCheckOutDia) <> "Nothing" Then
							Call Fn_Button_Click("Fn_SISW_Prop_ObjPropEdit",objCheckOutDia,"Yes")
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: The "& sObjectType  & " checked out  successfully")
						End If
						Set objCheckOutDia = Nothing
'					
						If JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Edit Properties").Exist(SISW_MAX_TIMEOUT) Then
									'Creating object of Edit Properties Dialog
									Set ObjPropertyWindow = Fn_UI_ObjectCreate("Fn_SISW_Prop_ObjectPropertyCkOut_Edit_CkIn",JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Edit Properties"))
									'Get the keys & items count from data dictionary.
									Call Fn_UI_JavaStaticText_Click(" Fn_TcObjectDelete", JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Edit Properties"), "BottomLink", 1, 1, "LEFT")

									'Clicking on "Show empty properties..." Static Text to show all properties
									Set objStat=description.Create()
										objStat("Class Name").value="JavaStaticText"
										objStat("label").value="Show empty properties..."
										wait(5)
									Set objDialog=JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Edit Properties").ChildObjects(objStat)
                                        objDialog(0).click 1,1
										wait(5)

									dictItems = oItemPropDictonary.Items
									dictKeys = oItemPropDictonary.Keys
									For i = 0 to oItemPropDictonary.Count - 1
									If IsNull(dictKeys(i))  Then
									Else
											If  dictitems(i) = "" Then						
											else
														sObjElement = dictKeys(i)
														
														' Set the value as per the data dictioanry key.

														Select case sObjElement
                                                        				Case  "Name", "Description", "URL"
																			If sObjElement = "URL" Then
																				ObjPropertyWindow.JavaEdit("URL").SetTOProperty "attached text", "URL:"
																				ObjPropertyWindow.JavaEdit("URL").Set  dictItems(i)
																			Else
																				Call Fn_Edit_Box("Fn_SISW_Prop_ObjectPropertyCkOut_Edit_CkIn",ObjPropertyWindow,sObjElement,dictItems(i))
																			End If

																			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: The property "& sObjElement  & " changed successfully")	
														End select 
											End If
									End If										
									Next
									'Click on Save and check-in button.
									Call Fn_ReadyStatusSync(1)
									Call Fn_Button_Click(" Fn_SISW_Prop_ObjectPropertyCkOut_Edit_CkIn",ObjPropertyWindow,"SaveAndCheck-In")
									Call Fn_ReadyStatusSync(1)
									Set ObjPropertyYes = Fn_SISW_GetChkInChkOutObject("CheckIn")
									If TypeName(ObjPropertyYes) <> "Nothing" Then
										Call Fn_Button_Click("Fn_SISW_Prop_ObjPropEdit",ObjPropertyYes,"Yes")
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: The URL Saved and checked in successfully")
									End If
									Set ObjPropertyYes = Nothing
									Fn_SISW_Prop_ObjectPropertyCkOut_Edit_CkIn = True
								end if
						
					End Select
SET objCheckOutDia = nothing
SET ObjPropertyWindow = nothing
SET ObjPropertyYes = nothing
SET objStat=Nothing
Set objDialog=Nothing
End Function

'#########################################################################################################
'###
'###    FUNCTION NAME   :   Fn_SISW_Prop_ObjectProperty_Edit(sObjectType, oObjetcTypeDictionary)
'###
'###    Description :
'###	Pre-requisite:					
'###							1. Context My Teamcenter
'###							2. Object Should be selected 
'###						
'###
'###    PARAMETERS      :      1. sObjectType - Which object like BOM line, Schedule, Tasks will be used for Property Edit 
'###						   2. oObjetcTypeDictionary -  Name of Property to be edited + Value will be passed using Data dictionary collection
'###									    
'###									  
'###
'###	 HISTORY       :   		 AUTHOR                 DATE        	VERSION
'###
'###    CREATED BY     :   		Mahendra Bhandarkar		21/05/2010      1.0
'###
'###    REVIWED BY     :   	    Mohit khare			    21-May-10		1.0
'###
'###    MODIFIED BY   :   Vrushali                            13-Dec-2011       
'###    MODIFIED BY   :   Swati									11-jun-2012       
'###    EXAMPLE          : Call Fn_SISW_Prop_ObjectProperty_Edit("BOMLine",oBOMLinePropDictonary)
'###    
	'How to use in Test						
'		1. Data Dictionary Declaration 
'					Dim oBOMLinePropDictonary
'					Set oBOMLinePropDictonary = CreateObject( "Scripting.Dictionary" )
'						With oBOMLinePropDictonary
'			            .Add "RevDescription", ""
'						End with

'		2. Use in Test :							
'					If oBOMLinePropDictonary.Exists("RevDescription") then 
'						oBOMLinePropDictonary.Remove("RevDescription")
'						oBOMLinePropDictonary.Add "RevDescription", "Changed3"
'					Else
'						oBOMLinePropDictonary.Add "RevDescription", "Changed3"
'					End if
'
'
'		Call Fn_SISW_Prop_ObjectProperty_Edit("BOMLine",oBOMLinePropDictonary)        
'---------------------------------------------------------------------------------------------------------------------------
'@@	   Developer Name		Date	  	Rev. No.					Changes Done										Reviewer
'@@-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@		Vivek Ahirrao	 24-Dec-2015	 1.0			   Added case "ReportInWhereUsed"					[TC1122-20151116d-24_12_2015-VivekA-NewDevelopment]	
'#########################################################################################################################################################################
Public Function Fn_SISW_Prop_ObjectProperty_Edit(sObjectType, oObjetcTypeDictionary)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_Prop_ObjectProperty_Edit"
	Dim objCheckOutDia,  oPropSelect, dictItems, dictKeys, iCounter, sProperty, sObjElement,  ObjPropertyWindow, ObjPropertyYes
	Dim objPropertyDialog1, objPropertyDialog

	Dim StrTitle
	StrTitle=JavaWindow("DefaultWindow").GetROProperty("title")    
	
	Select Case sObjectType
		Case "BOMLine" 
					'Menu change in TC 12's Structure Manager perspective
					If instr(1,StrTitle,"Structure Manager") Then
						If Not JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Properties").Exist(4) Then
						  Call Fn_MenuOperation("WinMenuSelect","View:View Properties	Alt+P")	 
				  	 	End If
				   Else 
				   		' # Code to handle Item Revision/BOMLine Property Edit
						'Use menu operations to open the Properties window.
					   	If Not JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Properties").Exist Then
							  Call Fn_MenuOperation("Select","View:Properties")	 
					    End If
					End If
					
					

					If JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Properties").Exist   Then
						Set ObjPropertyWindow = Fn_UI_ObjectCreate("Fn_SISW_Prop_ObjectProperty_Edit",JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Properties"))
					Elseif JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Properties").Exist Then
						Set ObjPropertyWindow = Fn_UI_ObjectCreate("Fn_SISW_Prop_ObjectProperty_Edit",JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Properties"))
					End If

					 If ObjPropertyWindow.Exist Then	
								dictItems = oObjetcTypeDictionary.Items
								dictKeys = oObjetcTypeDictionary.Keys
								For iCounter = 0 to oObjetcTypeDictionary.Count - 1
							  'Check the keys and Items value from datadictionary.
										If IsNull(dictKeys(iCounter))  Then
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: The Passed Value of Property in Dictionary is Null Empty ")
										Else
												If  dictitems(iCounter) = "" Then        
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: The Passed Value of Property in Dictionary is Empty ")
												else
														sObjElement = dictKeys(iCounter)
														Select case sObjElement
														' Set the value as per the data dictioanry key. 
														Case "RevName"
														 Call Fn_Edit_Box("Fn_SISW_Prop_ObjectProperty_Edit", ObjPropertyWindow, sObjElement , dictItems(iCounter))
														 Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: The property "+ sObjElement  +" changed with value [ "+dictItems(iCounter)+"] successfully")
														Case "RevDescription"
														 Call Fn_Edit_Box("Fn_SISW_Prop_ObjectProperty_Edit", ObjPropertyWindow, sObjElement , dictItems(iCounter))
														 Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: The property "+ sObjElement  +" changed with value [ "+dictItems(iCounter)+"] successfully")
														Case "Quantity"
														 Call Fn_Edit_Box("Fn_SISW_Prop_ObjectProperty_Edit", ObjPropertyWindow, sObjElement , dictItems(iCounter))
														 Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: The property "+ sObjElement  +" changed with value [ "+dictItems(iCounter)+"] successfully")
														End select 
												End If
										End If 
                    		  Next
					'Click the OK button to save the properties edited.
					Call Fn_Button_Click(" Fn_SISW_Prop_ObjectProperty_Edit",ObjPropertyWindow,"OK")
					Fn_SISW_Prop_ObjectProperty_Edit = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: The BOMLine saved  successfully")
					Else
					Fn_SISW_Prop_ObjectProperty_Edit = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: The BOMLine not saved  successfully")
					End If

		Case "Schedule"
					'# Code to handle Schedule Property Edit - Case not yet implemented
		Case "Task"
					'# Code to handle Task Property Edit - Case not yet implemented
		Case "Part"
					'Check if Edit Properties Dialog Exist
					If not Fn_UI_ObjectExist("Fn_SISW_Prop_ObjectProperty_Edit",JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Edit Properties"))Then
							Call Fn_MenuOperation("KeyPress","Edit:Properties")
					End If
		
					'Check If Dialog Check-Out Exist
					If not  Fn_UI_ObjectExist("Fn_SISW_Prop_ObjectProperty_Edit",JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Edit Properties")) Then
						'Check-Out the Item		
'						Call Fn_ObjectCheckOut("","", "", "False", "", "", "", "")	
						Set objCheckOutDia = Fn_SISW_GetChkInChkOutObject("CheckOut")
						 If TypeName(objCheckOutDia) <> "Nothing" Then
							Call Fn_Button_Click("Fn_SISW_Prop_ObjectProperty_Edit", objCheckOutDia,"Yes")
						 End If
						 Set objCheckOutDia = Nothing
					End If

					'Click on Static text
'					Call Fn_UI_JavaStaticText_Click(" Fn_SISW_Prop_ObjectProperty_Edit",  JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Edit Properties"), "BottomLink", 1, 1, "LEFT")
					JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Edit Properties").JavaStaticText("BottomLink").Click 1,1,"LEFT"
					Wait(3)
					JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Edit Properties").JavaStaticText("More...").Click 1,1,"LEFT"
					' Set the Property Value
					dictItems = oObjetcTypeDictionary.Items
					dictKeys = oObjetcTypeDictionary.Keys
					For iCounter = 0 to oObjetcTypeDictionary.Count - 1
						 'Check the keys and Items value from datadictionary.
							If IsNull(dictKeys(iCounter))  Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: The Passed Value of Property in Dictionary is Null Empty ")
							Else
								If  dictitems(iCounter) = "" Then        
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: The Passed Value of Property in Dictionary is Empty ")
								Else
									sObjElement = dictKeys(iCounter)
									Select case sObjElement
											' Set the value as per the data dictioanry key.                         
											case "DesignedRequired"
												JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Edit Properties").JavaRadioButton("DesignedRequired").SetTOProperty "attached text",Cstr(dictItems(iCounter))	
												JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Edit Properties").JavaRadioButton("DesignedRequired").Set "ON"
									End select 
								End If
							End If 
                      Next
		
					'Click on 'SaveAndCheck-In' Button
					Call Fn_Button_Click("Fn_SISW_Prop_ObjectProperty_Edit",JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Edit Properties"),"SaveAndCheck-In")
		
					'Click on 'Yes' Button
                  Set ObjPropertyYes = Fn_SISW_GetChkInChkOutObject("CheckIn")
				 If TypeName(ObjPropertyYes) <> "Nothing" Then
					Call Fn_Button_Click("Fn_SISW_Prop_ObjectProperty_Edit", ObjPropertyYes,"Yes")
				 End If
				 Set ObjPropertyYes = Nothing
				Fn_SISW_Prop_ObjectProperty_Edit = True

			Case "Item"
					'Check if Edit Properties Dialog Exist
					If not Fn_UI_ObjectExist("Fn_SISW_Prop_ObjectProperty_Edit",JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Edit Properties"))Then
							Call Fn_MenuOperation("KeyPress","Edit:Properties")
					End If
		
					'Check If Dialog Check-Out Exist
					If not  Fn_UI_ObjectExist("Fn_SISW_Prop_ObjectProperty_Edit",JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Edit Properties")) Then
						'Check-Out the Item		
'						Call Fn_ObjectCheckOut("","", "", "False", "", "", "", "")
						Set objCheckOutDia = Fn_SISW_GetChkInChkOutObject("CheckOut")
						 If TypeName(objCheckOutDia) <> "Nothing" Then
							Call Fn_Button_Click("Fn_SISW_Prop_ObjectProperty_Edit", objCheckOutDia,"Yes")
						 End If
						 Set objCheckOutDia = Nothing
					End If

					'Click on Static text
'					Call Fn_UI_JavaStaticText_Click(" Fn_SISW_Prop_ObjectProperty_Edit",  JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Edit Properties"), "BottomLink", 1, 1, "LEFT")
					JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Edit Properties").JavaStaticText("BottomLink").Click 1,1,"LEFT"
					Wait(3)

					If oObjetcTypeDictionary("ShowEmpty") = True Then
						JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Edit Properties").JavaStaticText("More...").Click 1,1,"LEFT"
					End If
					
					' Set the Property Value
					dictItems = oObjetcTypeDictionary.Items
					dictKeys = oObjetcTypeDictionary.Keys
					For iCounter = 0 to oObjetcTypeDictionary.Count - 1
						 'Check the keys and Items value from datadictionary.
							If IsNull(dictKeys(iCounter))  Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: The Passed Value of Property in Dictionary is Null Empty ")
							Else
								If  dictitems(iCounter) = "" Then        
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: The Passed Value of Property in Dictionary is Empty ")
								Else
									sObjElement = dictKeys(iCounter)
									Select case sObjElement
											' Set the value as per the data dictioanry key.                         
											case "ConfigurationItem"
												JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Edit Properties").JavaRadioButton("ConfigItem").SetTOProperty "attached text",Cstr(dictItems(iCounter))	
												JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Edit Properties").JavaRadioButton("ConfigItem").Set "ON"
									End select 
								End If
							End If 
                      Next
		
					'Click on 'SaveAndCheck-In' Button
					Call Fn_Button_Click("Fn_SISW_Prop_ObjectProperty_Edit",JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Edit Properties"),"SaveAndCheck-In")
		
					'Click on 'Yes' Button
                     Set ObjPropertyYes = Fn_SISW_GetChkInChkOutObject("CheckIn")
					 If TypeName(ObjPropertyYes) <> "Nothing" Then
						Call Fn_Button_Click("Fn_SISW_Prop_ObjectProperty_Edit", ObjPropertyYes,"Yes")
					 End If
					 Set ObjPropertyYes = Nothing
					Fn_SISW_Prop_ObjectProperty_Edit = True

	Case "Form"
					'Check if Edit Properties Dialog Exist
					 Set objPropertyDialog  = Fn_SISW_GetObject("Properties1")
					 Set objPropertyDialog1 = Fn_SISW_GetObject("Edit Properties")

					If  Fn_UI_ObjectExist("Fn_SISW_Prop_ObjectProperty_Edit",objPropertyDialog)Then
							Call Fn_Button_Click("Fn_SISW_Prop_ObjectProperty_Edit", objPropertyDialog, "Check-Out and Edit")
							wait 2
							   'Checking existance of [ Check-Out ] dialog
							If Instr(1,StrTitle,"My Teamcenter") or Instr(1,StrTitle,"Change Manager") Then
								Set objCheckOutDia = Fn_SISW_GetObject("PropCheck-Out")
							Else
								Set objCheckOutDia= Fn_SISW_GetChkInChkOutObject("CheckOut")
							End if
							 If TypeName(objCheckOutDia)<> "Nothing" Then
								Call Fn_Button_Click("Fn_SISW_Prop_ObjectProperty_Edit", objCheckOutDia,"Yes")
							 End If
							 Set objCheckOutDia = Nothing

							Set objPropertyDialog  = Fn_SISW_GetObject("Edit Properties_1")
							If  not Fn_UI_ObjectExist("Fn_SISW_Prop_ObjectProperty_Edit",objPropertyDialog)Then
								  Fn_SISW_Prop_ObjectProperty_Edit = False
								  Exit Function
							End If
					Else

								If not Fn_UI_ObjectExist("Fn_SISW_Prop_ObjectProperty_Edit",objPropertyDialog1)Then
										Call Fn_MenuOperation("KeyPress","Edit:Properties")
										Set objPropertyDialog = Fn_SISW_GetObject("Edit Properties")
								End If
		
								'Check If Dialog Check-Out Exist
								If not  Fn_UI_ObjectExist("Fn_SISW_Prop_ObjectProperty_Edit",objPropertyDialog1) Then
									'Check-Out the Item		
			'						Call Fn_ObjectCheckOut("","", "", "False", "", "", "", "")	
									
									'Checking existance of [ Check-Out ] dialog
									If Instr(1,StrTitle,"My Teamcenter") or Instr(1,StrTitle,"Change Manager") Then
										'[TC1121-2015101900-04_11_2015-VivekA-Maintenance] - Added by Priyanka K
										'Set objCheckOutDia = Fn_SISW_GetObject("PropCheck-Out")
										If JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("CheckOut").Exist(1) Then
											Set objCheckOutDia = Fn_SISW_GetObject("PropCheck-Out")
										ElseIf JavaWindow("DefaultWindow").JavaWindow("Check-Out").Exist(1) Then											
											Set objCheckOutDia = Fn_SISW_GetObject("Check-Out@2")
										End If
									Else
										Set objCheckOutDia= Fn_SISW_GetChkInChkOutObject("CheckOut")
									End if
									 If TypeName(objCheckOutDia)<> "Nothing" Then
										Call Fn_Button_Click("Fn_SISW_Prop_ObjectProperty_Edit", objCheckOutDia,"Yes")
									 End If
									 Set objCheckOutDia = Nothing
								End If
					End If
					
					If Fn_UI_ObjectExist("Fn_SISW_Prop_ObjectProperty_Edit",objPropertyDialog1)Then
						Set objPropertyDialog = objPropertyDialog1
					End if 

					'Click on Static text
'					Call Fn_UI_JavaStaticText_Click(" Fn_SISW_Prop_ObjectProperty_Edit",  JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Edit Properties"), "BottomLink", 1, 1, "LEFT")
					objPropertyDialog.JavaStaticText("BottomLink").Click 1,1,"LEFT"
					Wait(3)
					objPropertyDialog.JavaStaticText("More...").Click 1,1,"LEFT"
					' Set the Property Value
					dictItems = oObjetcTypeDictionary.Items
					dictKeys = oObjetcTypeDictionary.Keys
					For iCounter = 0 to oObjetcTypeDictionary.Count - 1
						 'Check the keys and Items value from datadictionary.
							If IsNull(dictKeys(iCounter))  Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: The Passed Value of Property in Dictionary is Null Empty ")
							Else
								If  dictitems(iCounter) = "" Then        
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: The Passed Value of Property in Dictionary is Empty ")
								Else
									sObjElement = dictKeys(iCounter)
									objPropertyDialog.JavaStaticText("ObjStaticText").setToProperty "attached text", sObjElement & ":"	
									call Fn_SISW_UI_JavaEdit_Operations("Fn_SISW_Prop_ObjectProperty_Edit", "Set", objPropertyDialog, "ObjEditbox", dictItems(iCounter))
									End If
							End If
					Next
		
					'Click on 'SaveAndCheck-In' Button
			Call Fn_Button_Click("Fn_SISW_Prop_ObjectProperty_Edit",objPropertyDialog,"SaveAndCheck-In")
			If Instr(1,StrTitle,"My Teamcenter") or Instr(1,StrTitle,"Change Manager") Then
				Set ObjPropertyYes = Fn_SISW_GetObject("Check-In@2")
			Else
				Set ObjPropertyYes = Fn_SISW_GetChkInChkOutObject("CheckIn")
			End if
			 If TypeName(ObjPropertyYes)<> "Nothing" Then
				Call Fn_Button_Click("Fn_SISW_Prop_ObjectProperty_Edit", ObjPropertyYes,"Yes")
				Fn_SISW_Prop_ObjectProperty_Edit = True
			 End If
			 Set ObjPropertyYes = Nothing
		'[TC1121-20161116a-09_12_2015-VivekA-NewDevelopment] - Added by Reema W
		' # Code to handle 4GD Workset Property Edit
		Case "Workset"
			'Use menu operations to open the Properties window.
			If JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Properties").Exist   Then
				Set ObjPropertyWindow = Fn_UI_ObjectCreate("Fn_SISW_Prop_ObjectProperty_Edit",JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Properties"))
			Elseif JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Properties").Exist Then
				Set ObjPropertyWindow = Fn_UI_ObjectCreate("Fn_SISW_Prop_ObjectProperty_Edit",JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Properties"))
			Else
				Fn_SISW_Prop_ObjectProperty_Edit = False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: The Property Dialog does not exist")
				Exit Function
			End If

			If ObjPropertyWindow.Exist Then	
				dictItems = oObjetcTypeDictionary.Items
				dictKeys = oObjetcTypeDictionary.Keys
				For iCounter = 0 to oObjetcTypeDictionary.Count - 1
					'Check the keys and Items value from datadictionary.
					If IsNull(dictKeys(iCounter))  Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: The Passed Value of Property in Dictionary is Null Empty ")
					Else
						If dictitems(iCounter) = "" Then        
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: The Passed Value of Property in Dictionary is Empty ")
						Else
							sObjElement = dictKeys(iCounter)
							Select case sObjElement
								' Set the value as per the data dictioanry key. 
								Case "IncludeInPartsList"
									ObjPropertyWindow.JavaStaticText("Property_label").SetTOProperty "label","Include In Parts List:"
									If dictItems(iCounter) = "ON" Then
										ObjPropertyWindow.JavaRadioButton("Property_True").Set "ON"
									Else
										ObjPropertyWindow.JavaRadioButton("Property_False").Set "ON"
									End If
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: The property "+sObjElement+" changed with value [ "+dictItems(iCounter)+"] successfully")
								' Set the value as per the data dictioanry key. 
								Case "ReportInWhereUsed"
									ObjPropertyWindow.JavaStaticText("Property_label").SetTOProperty "label","Report In Where Used:"
									If dictItems(iCounter) = "ON" Then
										ObjPropertyWindow.JavaRadioButton("Property_True").Set "ON"
									Else
										ObjPropertyWindow.JavaRadioButton("Property_False").Set "ON"
									End If
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: The property "+sObjElement+" changed with value [ "+dictItems(iCounter)+"] successfully")
							End select 
						End If
					End If 
            	Next
				'Click the OK button to save the properties edited.
				Call Fn_Button_Click(" Fn_SISW_Prop_ObjectProperty_Edit",ObjPropertyWindow,"OK")
				Fn_SISW_Prop_ObjectProperty_Edit = True
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: The Workset saved  successfully")
			Else
				Fn_SISW_Prop_ObjectProperty_Edit = False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: The Workset not saved  successfully")
			End If
	End Select

	Set objCheckOutDia = Nothing
	Set oPropSelect = Nothing
	Set dictItems = Nothing
	Set  dictKeys = Nothing
	Set iCounter = Nothing
	Set  sProperty = Nothing
	Set  sObjElement=Nothing
	Set  ObjPropertyWindow = Nothing
	Set ObjPropertyYes = Nothing

End Function

'*********************************************************	Function to Check out th property of object	,Explore and edit check in	******************************************************

'Function Name		:				Fn_SISW_Prop_CkOut_Explore_Edit_CkIn  

'Description			 :		 		 Note: This is the function which is made for special Case and  used in very few TestCases.

'Parameters			   :	 			1)sModifiedName:	ObjectName to be modified.
'													 2)sModifiedDesc:	 Description of the Object to be modified.
'													3)sExploreSelectAction: 	"SelectAll" or Objects to be Checked ':' Seperated String.If blank Donot Click Explore
'													4)sChkOutChildObjectVerify: 	Verify the Child object is Displayed After Explore operation in Edit property  Dialog.
'																												if Blank Donot Verify.
'													5)sEndAction: 	Action to be Taken at the end of the Function Like Save and Checkin/ Cancel Check Out/ Close/ Save
'																					if Multiple EndAction to be Passed it should be ":" seperated String.

'Return Value		   : 				True/False 

'Pre-requisite			:		 		Item should be selected

'Examples				:				 Fn_SISW_Prop_CkOut_Explore_Edit_CkIn ("Modifiednameshouldbegreaterthan32","Description","SelectAll","004171/A;1-new~004171/A~004171","Save:Close")

'History					 :		
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'												------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Amol/Pranav										   			20/05/2010			              1.0									Created
'												------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'**************************************************************************************************************************************************************************************************

Public Function Fn_SISW_Prop_CkOut_Explore_Edit_CkIn(sModifiedName,sModifiedDesc,sExploreSelectAction,sChkOutChildObjectVerify,sEndAction)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_Prop_CkOut_Explore_Edit_CkIn"
	Dim sFunctionName, bFlag, sactionItem, sfetched_value, sBottomPaneName
	Dim aItem,obj_statictxt,objchild1,iCount, objChkOut, objCancelChkOut, objChkIn

	Dim StrTitle
	StrTitle=JavaWindow("DefaultWindow").GetROProperty("title")

	  If Not Fn_UI_ObjectExist("Fn_SISW_Prop_CkOut_Explore_Edit_CkIn", JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Properties")) Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Properties Dialog  is  not presented of function " & sFunctionName)
			Fn_SISW_Prop_CkOut_Explore_Edit_CkIn=True
			Call Fn_MenuOperation("KeyPress","View:Properties")
			Call Fn_ReadyStatusSync(2)
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Sucessfully Clicked on view:Properties of Function " & sFunctionName)
	Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Properties Dialog  is presented of function " & sFunctionName)
			Fn_SISW_Prop_CkOut_Explore_Edit_CkIn=False
	 End If
	If  Fn_UI_ObjectExist("Fn_SISW_Prop_CkOut_Explore_Edit_CkIn",JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Properties")) Then
		Call Fn_Button_Click("Fn_SISW_Prop_CkOut_Explore_Edit_CkIn", JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Properties"),"Check-Out and Edit")
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "  Sucessfully Clicked on Check-Out and Edit Button of Function " & sFunctionName)
	End If

		Set objChkOut = Fn_SISW_GetChkInChkOutObject("CheckOut")
		If Typename(objChkOut) <> "Nothing" Then
			Call Fn_Button_Click("Fn_SISW_Prop_CkOut_Explore_Edit_CkIn",objChkOut,"Yes")
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Check-Out  Dialog is presented of Function " & sFunctionName)
		End If
		Set objChkOut = Nothing

		If  Fn_UI_ObjectExist("Fn_SISW_Prop_CkOut_Explore_Edit_CkIn",JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Edit Properties")) Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "   Edit Properties Dialog  Exists of Function " & sFunctionName)
			Fn_SISW_Prop_CkOut_Explore_Edit_CkIn=True

            	If  Fn_UI_ObjectExist("Fn_SISW_Prop_CkOut_Explore_Edit_CkIn",JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Edit Properties").JavaButton("SaveAndCheck-In")) Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "   Save and Check-InButton Exists of Function " & sFunctionName)
					Fn_SISW_Prop_CkOut_Explore_Edit_CkIn=True
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "  Save and Check-In Button does not exists of Function " & sFunctionName)
					Fn_SISW_Prop_CkOut_Explore_Edit_CkIn=False
				End if

				If  Fn_UI_ObjectExist("Fn_SISW_Prop_CkOut_Explore_Edit_CkIn",JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Edit Properties").JavaButton("Cancel Check-Out")) Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "   Cancel Check-Out Button exists of Function " & sFunctionName)
					Fn_SISW_Prop_CkOut_Explore_Edit_CkIn=True
				 Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "    Cancel Check-Out Button does not exists of Function " & sFunctionName)
					Fn_SISW_Prop_CkOut_Explore_Edit_CkIn=False
				End if

				If  Fn_UI_ObjectExist("Fn_SISW_Prop_CkOut_Explore_Edit_CkIn",JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Edit Properties").JavaButton("Close")) Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "   Close Button exists of Function " & sFunctionName)
					Fn_SISW_Prop_CkOut_Explore_Edit_CkIn=True
				 Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "  Close Button does not exists of Function " & sFunctionName)
					Fn_SISW_Prop_CkOut_Explore_Edit_CkIn=False
				End if

				If  Fn_UI_ObjectExist("Fn_SISW_Prop_CkOut_Explore_Edit_CkIn",JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Edit Properties").JavaButton("Save")) Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "  Save Button exists of Function " & sFunctionName)
					Fn_SISW_Prop_CkOut_Explore_Edit_CkIn=True
				 Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "  Save Button does not  exists of Function " & sFunctionName)
					Fn_SISW_Prop_CkOut_Explore_Edit_CkIn=False
				End if
				
				Call Fn_Button_Click("Fn_SISW_Prop_CkOut_Explore_Edit_CkIn",JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Edit Properties"),"Explore")

				If  Fn_UI_ObjectExist("Fn_SISW_Prop_CkOut_Explore_Edit_CkIn",JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Explore")) Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "   Explore Window is presented of function" & sFunctionName)
					Fn_SISW_Prop_CkOut_Explore_Edit_CkIn=True
					Call Fn_Button_Click("Fn_SISW_Prop_CkOut_Explore_Edit_CkIn",JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Explore"),"SelectAll")
					Call Fn_Button_Click("Fn_SISW_Prop_CkOut_Explore_Edit_CkIn",JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Explore"),"OK")
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "  Explore Window is not presented of function " & sFunctionName)
					Fn_SISW_Prop_CkOut_Explore_Edit_CkIn=False
				End if

	Set objChkOut = Fn_SISW_GetChkInChkOutObject("CheckOut")
	If Typename(objChkOut) <> "Nothing" Then
		Do 
			Call Fn_Button_Click("Fn_SISW_Prop_CkOut_Explore_Edit_CkIn",objChkOut,"Yes")
			Call Fn_ReadyStatusSync(2)
		Loop While objChkOut.Exist(1)
	End If
	Set objChkOut = Nothing

	Set obj_statictxt=description.Create()
	obj_statictxt("Class Name").value="JavaStaticText"
    Set objchild1=JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Edit Properties").ChildObjects(obj_statictxt)
	For iCount=1 to  objchild1.count -14
		aItem=split(sChkOutChildObjectVerify,"~")
		sBottomPaneName=objchild1(iCount).GetROProperty("attached text")
        If aItem(iCount-1)= sBottomPaneName Then 
			If iCount= (ubound (aItem))+1Then
				Exit For
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[ " +sBottomPaneName +"]  present in the bottom pane of the Edit Properties dialog of function " & sFunctionName)
				Fn_SISW_Prop_CkOut_Explore_Edit_CkIn=True
			End If
		Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[" +sBottomPaneName +"]  not present in the bottom pane of the Edit Properties dialog of function " & sFunctionName)
			Fn_SISW_Prop_CkOut_Explore_Edit_CkIn=False
		End If
	Next
		If  sModifiedDesc <>"" Then
			Call Fn_Edit_Box("Fn_SISW_Prop_CkOut_Explore_Edit_CkIn",JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Edit Properties"),"Description",sModifiedDesc)
			'Call Fn_Button_Click("Fn_SISW_Prop_CkOut_Explore_Edit_CkIn",JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Edit Properties"),"Save")
		End If
		If  Fn_UI_ObjectExist("Fn_SISW_Prop_CkOut_Explore_Edit_CkIn",JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Edit Properties")) Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Edit Properties Dialog is opened of function" & sFunctionName)
				Fn_SISW_Prop_CkOut_Explore_Edit_CkIn=True
				If len(sModifiedName)>32 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The length of modified name is greater than 32 characters of function" & sFunctionName)
                    Call Fn_Edit_Box("Fn_SISW_Prop_CkOut_Explore_Edit_CkIn",JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Edit Properties"),"Name",sModifiedName)
					sfetched_value=JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Edit Properties").JavaEdit("Name").GetROProperty("value")
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The Accepted name [" + sfetched_value + "]is equal or less than 32 characters " & sFunctionName)
				Else
					 Call Fn_Edit_Box("Fn_SISW_Prop_CkOut_Explore_Edit_CkIn",JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Edit Properties"),"Name",sModifiedName)
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The length of modified name [" + sModifiedName + "]is equal or less than 32 characters " & sFunctionName)
				End If
				sactionItem=split(sEndAction,":")
		        Call Fn_Button_Click("Fn_SISW_Prop_CkOut_Explore_Edit_CkIn",JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Edit Properties"),sactionItem(0))			

				Set objChkIn = Fn_SISW_GetChkInChkOutObject("CheckIn")
				If Typename(objChkIn) <> "Nothing" Then
						Call Fn_Button_Click("Fn_SISW_Prop_CkOut_Explore_Edit_CkIn",objChkIn,"Yes")
				End If
				Set objChkIn = Nothing

				Set objCancelChkOut = Fn_SISW_GetChkInChkOutObject("CancelCheckOut")
				If Typename(objCancelChkOut) <> "Nothing" Then
						Call Fn_Button_Click("Fn_SISW_Prop_CkOut_Explore_Edit_CkIn",objCancelChkOut,"Yes")
				End If
				Set objCancelChkOut = Nothing

				If  Fn_UI_ObjectExist("Fn_SISW_Prop_CkOut_Explore_Edit_CkIn",JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Edit Properties")) Then
					Call Fn_Button_Click("Fn_SISW_Prop_CkOut_Explore_Edit_CkIn",JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Edit Properties"),sactionItem(1))
				End If
        Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Edit Properties Dialog is not opened of function" & sFunctionName)
				Fn_SISW_Prop_CkOut_Explore_Edit_CkIn=False
		End if
	Else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Edit Properties Dialog Does Not  Exists of Function " & sFunctionName)
		Fn_SISW_Prop_CkOut_Explore_Edit_CkIn=False
	End If
End Function


'#######################################################################################
'###     FUNCTION NAME   :   Fn_SISW_Prop_HTML_PropertyVerify()
'###
'###    DESCRIPTION     :   This function verifies the object property through HTML page
 
'###   During the pilot function Fn_ObjectPropertyVerify got coded, which needs to be renamed and modified as per the current requirement.
 
'### The Following Function Are Club Together In this function
'### Fn_HTML_PropertyInvoke
'### Fn_SISW_Prop_HTML_PropertyVerify
'### Fn_HTML_PropertyClose
'###
'###    PARAMETERS      :   sPropertyName,sPropertyVal
'###
'###    Return Value  :   True/False  
'###
'###    HISTORY         :   			AUTHOR              		DATE        		VERSION
'###
'###    CREATED BY      :       Prasanna 					25/05/2010   			1.0
'###
'###    REVIWED BY      :		
'###
'###    MODIFIED BY     :
'###    EXAMPLE         :   Fn_SISW_Prop_HTML_PropertyVerify("current_name:creation_date" ,"AutoTestsData:20-May-2010 12:36:")
'#######################################################################################

Function Fn_SISW_Prop_HTML_PropertyVerify(sPropertyName,sPropertyVal)
		GBL_FAILED_FUNCTION_NAME="Fn_SISW_Prop_HTML_PropertyVerify"
		Dim ObjInv, arrPropertyName, arrPropertyVal, Rwcnt, objTable, iCounter, jRowConter, ObjClose, arrResult,  bCheckProp, iValueCounter , iArrSize        
		Dim arrSpecificString,iCountSpecific

Set ObjInv = JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet")
' ***********************Checking That  if both the windows are not already exists then call to Menu operation function***************
If  Fn_UI_ObjectExist("Fn_SISW_Prop_HTML_PropertyVerify",ObjInv.JavaDialog("Properties"))=False AND Fn_UI_ObjectExist("Fn_SISW_Prop_HTML_PropertyVerify",ObjInv.JavaDialog("Edit Properties"))=False Then
        Call Fn_MenuOperation("KeyPress","View:Properties")
		Call Fn_ReadyStatusSync(3)		
End If
'******************************************Here checking that  properties window *****************************************************
		If  Fn_UI_ObjectExist("Fn_SISW_Prop_HTML_PropertyVerify",ObjInv.JavaDialog("Properties"))=True Then
				'If Fn_UI_ObjectExist("Fn_SISW_Prop_HTML_PropertyVerify",ObjInv.JavaDialog("Properties").JavaButton("Print"))=True Then
						Call Fn_Button_Click("Fn_SISW_Prop_HTML_PropertyVerify",ObjInv.JavaDialog("Properties"),"Print")
						Call Fn_ReadyStatusSync(3)
				'End if
		End If

		If Fn_UI_ObjectExist("Fn_SISW_Prop_HTML_PropertyVerify",ObjInv.JavaDialog("Edit Properties"))=True Then		
						Call Fn_Button_Click("Fn_SISW_Prop_HTML_PropertyVerify",ObjInv.JavaDialog("Edit Properties"),"Print")
						Call Fn_ReadyStatusSync(3)
		End If

				'******************************************Here checking that  Print window  Exists*****************************************************
				If Fn_UI_ObjectExist("Fn_SISW_Prop_HTML_PropertyVerify",ObjInv.JavaDialog("Print"))=True Then
                        Call Fn_CheckBox_Set("Fn_SISW_Prop_HTML_PropertyVerify", ObjInv.JavaDialog("Print"), "HTML","ON")
						Call Fn_ReadyStatusSync(3)
                        Call Fn_Button_Click("Fn_SISW_Prop_HTML_PropertyVerify",ObjInv.JavaDialog("Print"),"OpenInBrowser")
						wait(3)
								 '********************************************Add Here browser check function************************************************************								
									If Browser("browser").page("Page").Exist(SISW_DEFAULT_TIMEOUT) Then							
'**********************************	**************************Property Verification start here -20-05-2010   ********************************************													
													'*********************************Split the number of parameters into an array**************************************
													arrPropertyName = Split(sPropertyName,":",-1)
													arrPropertyVal = Split(sPropertyVal,":",-1)										
													'*************************************Checking that the Browser is opened or not************************************
													If  Browser("browser").Page("Page").Exist(SISW_MIN_TIMEOUT) Then
														'************************Here we count the Number of rows present in that perticular table**************************
														Rwcnt = Browser("Browser").Page("Page").WebTable("PropertyTable").RowCount
														Set objTable = Browser("Browser").Page("Page").WebTable("PropertyTable")														
														'*****Here is the loop for checking the first colume value if matches then it goes to the second if  condition*******
														
														iValueCounter  = 0
														iCounter = 0														
														iArrSize = Ubound(arrPropertyName)													
														arrResult =arrPropertyName
														bCheckProp = false			  
														Do
																  For jRowConter = 1 to Rwcnt
																		'bCheckProp = false			  
																		If Trim(objTable.GetCellData(jRowConter,1)) = arrPropertyName(iCounter) Then
																			    bCheckProp = true
																				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Object Property Found for " + arrPropertyName(iCounter))      		
																				If Trim(objTable.GetCellData(jRowConter,2)) =arrPropertyVal(iValueCounter) Then                                                                          																							
																							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Object Value Verified for " +  arrPropertyName(iCounter))																							
																							iValueCounter = iValueCounter + 1
																							arrResult(iCounter) = true
																							Exit For
																				else 
															'*****Exceptional case in Case of date fields*******
																							If iArrSize >1 Then
																							
																										If  Trim(objTable.GetCellData(jRowConter,2)) = arrPropertyVal(iValueCounter)  + ":" + arrPropertyVal(iValueCounter+1)Then																									
																												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Object Value Verified for " + arrPropertyName(iCounter) )
																												iValueCounter = iValueCounter +2
																												arrResult(iCounter) = true
																												Exit For
																										else		
																												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Object Value Not Verified for" +  arrPropertyName(iCounter))
																												arrResult(iCounter) = false             'if  values in the table does not match return false
																										End If
																							Else
																											'Added to check participants
																										   If instr(1,arrPropertyVal(iValueCounter),",")  Then
																														 arrSpecificString = split(arrPropertyVal(iValueCounter),",",-1,1)
																														 For iCountSpecific = 0 to Ubound(arrSpecificString)
																																	If instr(1,Trim(objTable.GetCellData(jRowConter,2)),arrSpecificString(iCountSpecific)) Then
																																				arrResult(iCounter) = true
																																	Else
																																				  arrResult(iCounter) = false
																																					Exit for
																																	End If      
																														 Next
																											Else
																															arrResult(iCounter) = false
																											 End If               
																							End If                                                                      																						
																				 End If		
																	       End If																	
																	Next				
																			If bCheckProp = false Then
																				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Object Property Not Found for " + arrPropertyName(iCounter))      		
																				arrResult(iCounter) = false
																			End If
																	iCounter  = iCounter +1												
														Loop while iCounter <= iArrSize
												Else
															Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Browser window does not exist. ")      		

												End If                          				

												For iCounter = 0 to iArrSize
															if arrResult(iCounter) = true then
																	Fn_SISW_Prop_HTML_PropertyVerify = true
															else
																	Fn_SISW_Prop_HTML_PropertyVerify = false
																	Exit for
															End If
												Next
																											
	'*******************************************'Property Verification Ends  here -20-05-2010************************************************************************************                                    									
									Else
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Print window does not exist. ")      		
									End If

				Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Properties window does not exist. ")      		
							Exit Function
				 End If
'********************************************** Close All the window code start  here -20-05-2010***********************************
                                        									
											Set ObjClose = JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet")
										'**************************************Check that browser is already open or not*************************************
											If browser("Browser").Page("Page").Exist(SISW_MIN_TIMEOUT) Then
														Browser("Browser").Close
														If Fn_UI_ObjectExist("Fn_SISW_Prop_HTML_PropertyVerify", ObjClose.JavaDialog("Print"))=True Then
															Call Fn_Button_Click("Fn_SISW_Prop_HTML_PropertyVerify", ObjClose.JavaDialog("Print"), "Close")
															wait(3)
														End if
														'**************************************Check that  Edit Properties Dialog box is exists or not*************************
														If  Fn_UI_ObjectExist("Fn_SISW_Prop_HTML_PropertyVerify", ObjClose.JavaDialog("Edit Properties"))=True Then
																	Call Fn_Button_Click("Fn_SISW_Prop_HTML_PropertyVerify",ObjClose.JavaDialog("Edit Properties"), "Close")
																	wait(3)
												
																	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS:Closed Edit Properties  Dialog successfully.")
														'**********************************Check that  Properties Dialog box is exists or not**********************************
														Elseif Fn_UI_ObjectExist("Fn_SISW_Prop_HTML_PropertyVerify", ObjClose.JavaDialog("Properties"))=True Then
																	Call Fn_Button_Click("Fn_SISW_Prop_HTML_PropertyVerify",ObjClose.JavaDialog("Properties"), "Cancel")                                                                	 																
																	wait(3)
												
																   Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS:Closed Properties  Dialog successfully.")
														Else 
														'I************************that window is not  exist then it giver the error message**************************************															
												
																	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL:Couldn't find the Edit Properties or Properties dialog")
																	Exit Function 
														End If
											Else														
																	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL:Can't Find Browser")
											End If								
	'End If

	Set ObjInv = Nothing
	Set ObjClose = Nothing
	Set objTable = Nothing

End Function


'######################################################################################################################################################################
'#########################################################################################################
'###
'###    FUNCTION NAME   :   Fn_SISW_Prop_ObjectPropertyOperation(sObjectName, sAction, oObjetcTypeDictionary)
'###
'###    DESCRIPTION        :   Update / Validate that the Object Property/Value field associated to Teamcenter Business Object
'###
'###    PARAMETERS      :   sObjectName: Name of the Teamcenter Business Object operated upon. This value is used to dynamically set the window caption
'###											 sAction: This is action to be performed on the Teamcenter Business Object 
'###											oObjetcTypeDictionary: Dictionary Object holding [Key-Value] pair of the properties to be evaluated for Read-Only State on [Object Property] dialog
'###
'###    RETURNS     		:   True / False
'###
'###	 HISTORY            	 :   AUTHOR                 				DATE        VERSION
'###
'###    CREATED BY     :    Mahendra Bhandarkar          				 26/05/2010         1.0
'###
'###    Modified BY     :   Koustubh watwe           	  				14/03/2012         1.0		
'###
'###    REVIWED BY     :    Mohit Khare									26/05/2010		 	1.0
'###
'###  Modified By : 		Nilesh Gadekar                  11-October-2012   :  Build : Teamcenter 10 (20120919.00)
'###  Modified By :         Dipali Karwande					8 Feb 2013			Added case "PropertyNameVerify" 	
'###
'###   EXAMPLES 		:  Call Fn_SISW_Prop_ObjectPropertyOperation("A_1", "Form Property Is Editable", oFormPropDictonary)
'###						EXAMPLES 		:  Call Fn_SISW_Prop_ObjectPropertyOperation("All", "Form Property Is Editable", oFormPropDictonary)
'###    How to use in Test						
'		1. Data Dictionary Declaration 
'					Dim oFormPropDictonary
'					Set oFormPropDictonary = CreateObject( "Scripting.Dictionary" )
' 					With oFormPropDictonary     
'           			 .Add "FormName", ""
'            			.Add "FormDescription", ""     
'					End with
'		2. Use in Test :							
'							If oFormPropDictonary.Exists("Name") then 
'								oFormPropDictonary.Remove("Name")
'								oFormPropDictonary.Add "FormName","verified3"
'							Else
'								oFormPropDictonary.Add "FormName","verified"3
'							End if 
'							If oFormPropDictonary.Exists("FormDescription") then 
'								oFormPropDictonary.Add "FormDescription","description verified "
'							Else
'								oFormPropDictonary.Add "FormDescription","description verified"
'							End if 
'
							'Dim oObjetcTypeDictionary
							'					Set oObjetcTypeDictionary = CreateObject( "Scripting.Dictionary" )
							' 					With oObjetcTypeDictionary     
							'	'					.Add "PropertyNames", "Object:Name"
							'           			.Add "PropertyValues", "000041-Item:Item"     
							'					End with
							'msgbox  Fn_SISW_Prop_ObjectPropertyOperation("General", "PropertyValueVerify", oObjetcTypeDictionary)
'#############################################################################################################
Public Function Fn_SISW_Prop_ObjectPropertyOperation(sObjectName, sAction, oObjetcTypeDictionary)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_Prop_ObjectPropertyOperation"
	Dim objCheckOutDia,  oPropSelect, dictItems, dictKeys, iCounter, sProperty, sObjElement,  objPropertyWindow, ObjPropertyYes, bFlag
	Dim sValues,oss,sOSName,os
	Dim objWMIService, colItems, iXpos, iYpos, objItem

    Dim StrTitle
	StrTitle=JavaWindow("DefaultWindow").GetROProperty("title")

	Set  objPropertyWindow = JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Form Properties")
			
	'# Code till Step [1]
	Select Case sAction
		'# Code to handle Form Property Evaluation
		Case "Form Property Is Editable"
			'	Use menu operations to open the Properties window.
			objPropertyWindow.SetTOProperty "title", sObjectName

			Set objCheckOutDia = Fn_SISW_GetChkInChkOutObject("CheckOut")
			 If TypeName(objCheckOutDia) <> "Nothing" Then
				Call Fn_Button_Click("Fn_SISW_Prop_ObjectProperty_Edit", objCheckOutDia,"Yes")
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: The Element ["+sObjectName+"]checked out  successfully")	
			 End If
			 Set objCheckOutDia = Nothing

			If objPropertyWindow.Exist Then

				'	Get the keys & items count from data dictionary.

					dictItems = oObjetcTypeDictionary.Items
					dictKeys = oObjetcTypeDictionary.Keys

					For iCounter = 0 to oObjetcTypeDictionary.Count - 1

					If IsNull(dictKeys(iCounter))  Then
								bFlag = False
					Else
							If  dictitems(iCounter) = "" Then
									bFlag = False
							else
										sObjElement = dictKeys(iCounter)

										' Set the value as per the data dictioanry key.
										select case sObjElement
												Case  "ProjectID", "PreviousID", "SerialNumber", "UserData1", "UserData2", "UserData3", "ItemComment"
													 If CBool(Fn_UI_Object_GetROProperty("Fn_SISW_Prop_ObjectPropertyOperation",objPropertyWindow.JavaEdit(sObjElement), "editable"))  = CBool(dictItems(iCounter)) Then
																bFlag = True
													Else
																bFlag = False
													End If
										End select 
							End If
					End If

					If bFlag = True Then
						Fn_SISW_Prop_ObjectPropertyOperation = True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: The property "& sObjElement  & " is verified in readonly state successfully")
					Else
						Fn_SISW_Prop_ObjectPropertyOperation = False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: The property "& sObjElement  & " is verified in readonly state successfully")
					End If

					Next
					'Harshal Agrawal: Changes Button From Cancel to Close [15DEC2010]
					Call Fn_Button_Click(" Fn_SISW_Prop_ObjectPropertyOperation",objPropertyWindow,"Close")
				
				End If

			Case "Form Property Edit And Save"'					# Code to handle Form Property Edit and Save

'				#Use menu operations to open the Properties window.
				objPropertyWindow.SetTOProperty "title", sObjectName
				If objPropertyWindow.Exist(SISW_MIN_TIMEOUT) = False Then
					 Set objPropertyWindow = JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Form Properties")
					objPropertyWindow.SetTOProperty "title", sObjectName
					If objPropertyWindow.Exist(SISW_MIN_TIMEOUT) = False Then
						Fn_SISW_Prop_ObjectPropertyOperation = False
					End If
				End If
				If Fn_UI_ObjectExist("Fn_SISW_Prop_ObjectPropertyOperation", objPropertyWindow.JavaButton("Check-OutAndEdit")) Then
					Call Fn_Button_Click(" Fn_SISW_Prop_ObjectPropertyOperation",objPropertyWindow ,"Check-OutAndEdit")
				End If
				Wait 3

				Set objCheckOutDia = Fn_SISW_GetChkInChkOutObject("CheckOut")
				If TypeName(objCheckOutDia) <> "Nothing" Then
					Call Fn_Button_Click("Fn_SISW_Prop_ObjectProperty_Edit", objCheckOutDia,"Yes")
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: The Element ["+sObjectName+"]checked out  successfully")	
				End If
				Set objCheckOutDia = Nothing

				If objPropertyWindow.Exist Then
				'	Get the keys & items count from data dictionary.
					dictItems = oObjetcTypeDictionary.Items
					dictKeys = oObjetcTypeDictionary.Keys
					For iCounter = 0 to oObjetcTypeDictionary.Count - 1
						If IsNull(dictKeys(iCounter))  Then
							bFlag = False
						Else
							If  dictitems(iCounter) = "" Then
								bFlag = False
							Else
								sObjElement = dictKeys(iCounter)
								' Set the value as per the data dictioanry key.
								select case sObjElement
										Case "PreviousID","UserData1", "UserData2", "UserData3", "ItemComment"
											If Fn_Edit_Box("Fn_SISW_Prop_ObjectPropertyOperation",objPropertyWindow,sObjElement,dictItems(iCounter)) = True Then
												bFlag = True
											Else
												bFlag = False
											End If
										Case  "SerialNumber"
												objPropertyWindow.JavaEdit("SerialNumber").highlight
												objPropertyWindow.JavaEdit("SerialNumber").Type dictItems(iCounter)
												bFlag = True
										Case  "ProjectID"
												objPropertyWindow.JavaEdit("ProjectID").highlight
												objPropertyWindow.JavaEdit("ProjectID").Type dictItems(iCounter)
												bFlag = True
										Case Else
												bFlag = False
								End select 
							End If
					End If
					If bFlag = True Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: The property "& sObjElement  & " is modified with value ["+dictItems(iCounter)+"] successfully")
					Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: The property "& sObjElement  & " was not modified with value ["+dictItems(iCounter)+"].")
					End If
				Next
				Call Fn_Button_Click(" Fn_SISW_Prop_ObjectPropertyOperation",JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Form Properties"),"Save")
				If bFlag = True Then
					Fn_SISW_Prop_ObjectPropertyOperation = True
				Else
					Fn_SISW_Prop_ObjectPropertyOperation = False
				End If						
			End If

		Case "Form Property Save And Checkin"
			'# Code to handle Form Property Edit Save and Check-In
			'Use menu operations to open the Properties window.
			objPropertyWindow.SetTOProperty "title", sObjectName
			If objPropertyWindow.Exist(SISW_MIN_TIMEOUT) = False Then
				Set objPropertyWindow = JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Form Properties")
				objPropertyWindow.SetTOProperty "title", sObjectName
				If objPropertyWindow.Exist(SISW_MICRO_TIMEOUT) = False Then
					Fn_SISW_Prop_ObjectPropertyOperation = False
				End If
			End If
			If Fn_UI_ObjectExist("Fn_SISW_Prop_ObjectPropertyOperation", objPropertyWindow.JavaButton("Check-OutAndEdit")) Then
				Call Fn_Button_Click(" Fn_SISW_Prop_ObjectPropertyOperation",objPropertyWindow ,"Check-OutAndEdit")
				gLastMenuCall = "CheckOutAndEdit"
			End If
			Wait 3

			Set objCheckOutDia = Fn_SISW_GetChkInChkOutObject("CheckOut")
			If TypeName(objCheckOutDia) <> "Nothing" Then
				Call Fn_Button_Click("Fn_SISW_Prop_ObjectProperty_Edit", objCheckOutDia,"Yes")
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: The Element ["+sObjectName+"]checked out  successfully")	
			End If
			Set objCheckOutDia = Nothing

			If objPropertyWindow.Exist Then
			'	Get the keys & items count from data dictionary.
				dictItems = oObjetcTypeDictionary.Items
				dictKeys = oObjetcTypeDictionary.Keys

				For iCounter = 0 to oObjetcTypeDictionary.Count - 1

				If IsNull(dictKeys(iCounter))  Then
					bFlag = False
				Else
					If  dictitems(iCounter) = "" Then
						bFlag = False
					else
						sObjElement = dictKeys(iCounter)
						' Set the value as per the data dictioanry key.
						select case sObjElement
							Case  "ProjectID", "PreviousID", "SerialNumber", "UserData1", "UserData2", "UserData3", "ItemComment"
								wait 2
								If Fn_Edit_Box("Fn_SISW_Prop_ObjectPropertyOperation",objPropertyWindow,sObjElement,dictItems(iCounter)) = True Then
									bFlag = True
								Else
									bFlag = False
								End If
						End select 
					End If
				End If

				If bFlag = True Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: The property "& sObjElement  & " is modified with value ["+dictItems(iCounter)+"] successfully")
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: The property "& sObjElement  & " was not modified with value ["+dictItems(iCounter)+"].")
				End If
				Next

				Call Fn_Button_Click(" Fn_SISW_Prop_ObjectPropertyOperation",objPropertyWindow,"SaveAndCheck-In")
				gLastMenuCall = "SaveAndCheckIn"
				Set objCheckOutDia = Fn_SISW_GetChkInChkOutObject("CheckIn")
				 If TypeName(objCheckOutDia) <> "Nothing" Then
					Call Fn_Button_Click("Fn_SISW_Prop_ObjectProperty_Edit", objCheckOutDia,"Yes")
				 End If
				 Set objCheckOutDia = Nothing

'						Call Fn_Button_Click(" Fn_SISW_Prop_ObjectPropertyOperation",JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Form Properties"),"SaveAndCheck-In")
'						If Fn_UI_ObjectExist("Fn_SISW_Prop_ObjectPropertyOperation", JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Check-In")) Then
'									Set objCheckOutDia = Fn_UI_ObjectCreate("Fn_SISW_Prop_ObjectPropertyOperation",JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Check-In"))
'									Call Fn_Button_Click(" Fn_SISW_Prop_ObjectPropertyOperation",objCheckOutDia,"Yes")
'									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: The Element ["+sObjectName+"] checked in  successfully")	
'						End If
				If bFlag = True Then
					Fn_SISW_Prop_ObjectPropertyOperation = True
				Else
					Fn_SISW_Prop_ObjectPropertyOperation = False
				End If						
			End If

		Case "Form Property Cancel Check-Out"
'			# Code to handle Form Property Cancel Check-Out

		Case "URL Property Is Editable"
'		# Code to handle URL Property Evaluation

		Case "URL Property Edit And Save"
'			# Code to handle URL Property Edit and Save

		Case "URL Property Save And Checkin"
'		# Code to handle URL Property Save And Checkin

			JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Form Properties").SetTOProperty "title", sObjectName
			Set  objPropertyWindow = JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Form Properties")
'			
			Set objCheckOutDia = Fn_SISW_GetChkInChkOutObject("CheckOut")
			If TypeName(objCheckOutDia) <> "Nothing" Then
				Call Fn_Button_Click("Fn_SISW_Prop_ObjectProperty_Edit", objCheckOutDia,"Yes")
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: The Element ["+sObjectName+"] checked out  successfully")	
			End If
			Set objCheckOutDia = Nothing

			If objPropertyWindow.Exist Then

				'	Get the keys & items count from data dictionary.

					dictItems = oObjetcTypeDictionary.Items
					dictKeys = oObjetcTypeDictionary.Keys

					For iCounter = 0 to oObjetcTypeDictionary.Count - 1

					If IsNull(dictKeys(iCounter))  Then
								bFlag = False
					Else
							If  dictitems(iCounter) = "" Then
									bFlag = False
							else
										sObjElement = dictKeys(iCounter)

										' Set the value as per the data dictioanry key.
										select case sObjElement
												Case  "URL"
													objPropertyWindow.JavaEdit("SerialNumber").SetTOProperty "attached text", "URL:"
													objPropertyWindow.JavaEdit("SerialNumber").Set dictItems(iCounter)
													 bFlag = True
												Case  "Disposition"
													objPropertyWindow.JavaEdit("SerialNumber").SetTOProperty "attached text", "Disposition:"
													objPropertyWindow.JavaEdit("SerialNumber").Set dictItems(iCounter)
													 bFlag = True
										End select 
							End If
					End If

					If bFlag = True Then
						Fn_SISW_Prop_ObjectPropertyOperation = True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: The property "& sObjElement  & " is verified in readonly state successfully")
					Else
						Fn_SISW_Prop_ObjectPropertyOperation = False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: The property "& sObjElement  & " is verified in readonly state successfully")
					End If

					Next
					
					Call Fn_Button_Click(" Fn_SISW_Prop_ObjectPropertyOperation",objPropertyWindow,"SaveAndCheck-In")
					gLastMenuCall = "SaveAndCheckIn"
					Set objCheckOutDia = Fn_SISW_GetChkInChkOutObject("CheckIn")
					 If TypeName(objCheckOutDia) <> "Nothing" Then
						Call Fn_Button_Click("Fn_SISW_Prop_ObjectProperty_Edit", objCheckOutDia,"Yes")
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: The Element ["+sObjectName+"] checked in  successfully")	
					 End If
					 Set objCheckOutDia = Nothing
				
				End If

		Case "URL Property Cancel Check-Out"
'			# Code to handle URL Cancel Check-Out

		Case "Form Property Exist"
'		# Code to check the existence of Property

								'Use menu operations to open the Properties window.

								JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Form Properties").SetTOProperty "title", sObjectName
								Set  objPropertyWindow = Fn_UI_ObjectCreate("Fn_SISW_Prop_ObjectPropertyOperation", JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Form Properties"))

								If objPropertyWindow.Exist Then

								'	Get the keys & items count from data dictionary.

										dictItems = oObjetcTypeDictionary.Items
										dictKeys = oObjetcTypeDictionary.Keys

									    For iCounter = 0 to oObjetcTypeDictionary.Count - 1

											If IsNull(dictKeys(iCounter))  Then
														bFlag = False
											Else
													If  dictitems(iCounter) = "" Then
															bFlag = False
													else
																sObjElement = dictKeys(iCounter)
																' Set the value as per the data dictioanry key.
																select case sObjElement
																		Case   "SaveAndCheck-In", "Save", "Close", "Cancel Check-Out", "Check-OutAndEdit","Cancel"', "FormName", "FormDescription"
																				If Fn_UI_ObjectExist("", JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Form Properties").JavaButton(sObjElement))  = dictItems(iCounter) Then
																							bFlag = True
																							Exit for
																				Else
																							bFlag = False
																				End If
																		Case  "ProjectID", "PreviousID", "SerialNumber", "UserData1", "UserData2", "UserData3", "ItemComment"
																				If Fn_UI_ObjectExist("", JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Form Properties").JavaEdit(sObjElement))  = dictItems(iCounter) Then
																							bFlag = True
																				Else
																							bFlag = False
																				End If

																End select 
													End If
											End If
										Next
										
										If bFlag = True Then
											Fn_SISW_Prop_ObjectPropertyOperation = True
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: The property "& sObjElement  & " Exist ")
										Else
											Fn_SISW_Prop_ObjectPropertyOperation = False
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: The property "& sObjElement  & " does not Exist ")
										End If

                                        If objPropertyWindow.JavaButton("Cancel").Exist = True Then
												Call Fn_Button_Click(" Fn_SISW_Prop_ObjectPropertyOperation",objPropertyWindow,"Cancel")
										End If
										
										If objPropertyWindow.JavaButton("Close").Exist = True Then
												Call Fn_Button_Click(" Fn_SISW_Prop_ObjectPropertyOperation",objPropertyWindow,"Close")
										End If
									
								End If
								
					Case "PropertyNameVerify"
						if Fn_SISW_GetObject("Properties").Exist(SISW_MIN_TIMEOUT) = False AND Fn_SISW_GetObject("Properties1").Exist(SISW_MIN_TIMEOUT) = False Then
							call Fn_MenuOperation("KeyPress","View~Properties")	
						End If
						If Fn_SISW_GetObject("Properties").Exist(SISW_MICRO_TIMEOUT) Then
							Set objPropertyWindow = Fn_SISW_GetObject("Properties")
						ElseIf Fn_SISW_GetObject("Properties1").Exist(SISW_MICRO_TIMEOUT) Then
							Set objPropertyWindow = Fn_SISW_GetObject("Properties1")
						Else
							Set objPropertyWindow = Nothing
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail~ [ Fn_SISW_Prop_ObjectPropertyOperation ] Failed to open Proeprty Winodw.")
							Exit function
						End IF
					
						' fetching screen resolution
						Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
						Set oss = objWMIService.ExecQuery ("Select * from Win32_OperatingSystem") 
						For Each os in oss 
							sOSName = os.Caption
						Next
						Set objWMIService = Nothing
						If instr(1,lcase(sOSName),"xp") Then
							' fetching screen resolution
							Set objWMIService = GetObject("winmgmts~\\" & "." & "\root\cimv2")
							Set colItems = objWMIService.ExecQuery("Select * From Win32_DisplayConfiguration")
						Else
							Set objWMIService = GetObject("Winmgmts:\\.\root\cimv2")
							Set colItems = objWMIService.ExecQuery("Select * From Win32_DisplayConfiguration")
						End If

						For Each objItem in colItems
							iXpos = objItem.PelsWidth       'Horizontal Resolution
							iYpos = objItem.PelsHeight      'Vertical Resolution
						Next
						' resizing property window to screen resolution
						objPropertyWindow.Resize iXpos, iYpos
							
						' second click on given link
						If sObjectName = "" Then 
							sObjectName = "All"
						End If
'                        
						objPropertyWindow.JavaStaticText("BottomLink").setTOProperty "label", sObjectName
						Call Fn_UI_JavaStaticText_Click("Fn_SISW_Prop_ObjectPropertyOperation", objPropertyWindow,"BottomLink",10, 10,"LEFT")

						sProperty = Split(oObjetcTypeDictionary("PropertyNames"),"~")
						For iCounter = 0 to UBound(sProperty)
							objPropertyWindow.JavaStaticText("Property_label").setTOProperty "label", trim(sProperty(iCounter)) & ":"
							If  objPropertyWindow.JavaStaticText("Property_label").exist(SISW_MIN_TIMEOUT) Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Pass: [ Fn_SISW_Prop_ObjectPropertyOperation ] Successfully to verify [ " & trim(sProperty(iCounter)) &" ].")
								Fn_SISW_Prop_ObjectPropertyOperation = True
							Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail: [ Fn_SISW_Prop_ObjectPropertyOperation ] Failed to verify [ " & trim(sProperty(iCounter))  & " ].")
								Fn_SISW_Prop_ObjectPropertyOperation = False
								Exit for
							End if 
						Next
					''Close property Dialog
						If objPropertyWindow.JavaButton("Cancel").Exist = True Then
							Call Fn_Button_Click(" Fn_SISW_Prop_ObjectPropertyOperation",objPropertyWindow,"Cancel")
						End If				
				Case "PropertyValueVerify"
					If Fn_SISW_GetObject("Properties").Exist(SISW_MIN_TIMEOUT) = False AND Fn_SISW_GetObject("Properties1").Exist(SISW_MIN_TIMEOUT) = False Then
						call Fn_MenuOperation("KeyPress","View:Properties")	
					End If
					If Fn_SISW_GetObject("Properties").Exist(SISW_MICRO_TIMEOUT) Then
						Set objPropertyWindow = Fn_SISW_GetObject("Properties")
					ElseIf Fn_SISW_GetObject("Properties1").Exist(SISW_MICRO_TIMEOUT) Then
						Set objPropertyWindow = Fn_SISW_GetObject("Properties1")
					Else
						Set objPropertyWindow = Nothing
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail: [ Fn_SISW_Prop_ObjectPropertyOperation ] Failed to open Proeprty Winodw.")
						Exit function
					End IF
					
					' fetching screen resolution
					Set objWMIService = GetObject("winmgmts:\\" & "." & "\root\cimv2")
					Set colItems = objWMIService.ExecQuery("Select * From Win32_DisplayConfiguration")

					For Each objItem in colItems
						iXpos = objItem.PelsWidth       'Horizontal Resolution
						iYpos = objItem.PelsHeight      'Vertical Resolution
					Next
					' resizing property window to screen resolution
					objPropertyWindow.Resize iXpos, iYpos
						
					' second click on given link
					If sObjectName = "" Then 
						sObjectName = "All"
					End If
                        
					objPropertyWindow.JavaStaticText("BottomLink").setTOProperty "label", sObjectName
					Call Fn_UI_JavaStaticText_Click("Fn_SISW_Prop_ObjectPropertyOperation", objPropertyWindow,"BottomLink",10, 10,"LEFT")
					' oObjetcTypeDictionary
					' if empty click on All
					sProperty = Split(oObjetcTypeDictionary("PropertyNames"),":")
					sValues =  Split(oObjetcTypeDictionary("PropertyValues"),":")
					Fn_SISW_Prop_ObjectPropertyOperation = false
					' loop - verify fieilds and values
					For iCounter = 0 to uBound(sProperty)
								Fn_SISW_Prop_ObjectPropertyOperation = true
								objPropertyWindow.JavaStaticText("Property_label").setTOProperty "label", trim(sProperty(iCounter)) & ":"
								' edit
								bFlag=False
								' Added This Code to Varify object_string and Logged Date in which its Value has ":" , hence Instr is used
								' Modified By : Nilesh Gadekar : 11-October-2012 : Teamcenter 10 (20120919.00)
								If sProperty(iCounter)="object_string" OR sProperty(iCounter)="Logged Date" Then
									bFlag=True
										If objPropertyWindow.JavaEdit("Property_field").Exist(SISW_MIN_TIMEOUT) Then
											If Instr(trim(objPropertyWindow.JavaEdit("Property_field").getROProperty("value")) , trim(sValues(iCounter)))<=0 Then
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail: [ Fn_SISW_Prop_ObjectPropertyOperation ] Failed to verify [ " & trim(sProperty(iCounter)) & " = " & trim(sValues(iCounter)) & " ].")
													Fn_SISW_Prop_ObjectPropertyOperation = False
													Exit For
											End If
										ElseIF objPropertyWindow.JavaStaticText("Property_label_value").exist(SISW_MIN_TIMEOUT) Then
												If Instr(trim(objPropertyWindow.JavaStaticText("Property_label_value").getROProperty("label")) , trim(sValues(iCounter)))<=0 Then
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail: [ Fn_SISW_Prop_ObjectPropertyOperation ] Failed to verify [ " & trim(sProperty(iCounter)) & " = " & trim(sValues(iCounter)) & " ].")
														Fn_SISW_Prop_ObjectPropertyOperation = False
														Exit For
												End If
										ElseIF objPropertyWindow.JavaCheckBox("Property_field").exist(SISW_MIN_TIMEOUT) Then
												If Instr(trim(objPropertyWindow.JavaCheckBox("Property_field").getROProperty("label")) , trim(sValues(iCounter)))<=0 Then
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail: [ Fn_SISW_Prop_ObjectPropertyOperation ] Failed to verify [ " & trim(sProperty(iCounter)) & " = " & trim(sValues(iCounter)) & " ].")
														Fn_SISW_Prop_ObjectPropertyOperation = False
														Exit For
												End If
										End If
									End If
									If bFlag=False Then
                                    			If objPropertyWindow.JavaEdit("Property_field").Exist(SISW_MICRO_TIMEOUT) Then
													If trim(objPropertyWindow.JavaEdit("Property_field").getROProperty("value")) <> trim(sValues(iCounter)) Then
															Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail: [ Fn_SISW_Prop_ObjectPropertyOperation ] Failed to verify [ " & trim(sProperty(iCounter)) & " = " & trim(sValues(iCounter)) & " ].")
															Fn_SISW_Prop_ObjectPropertyOperation = False
															Exit For
													End If
											ElseIF objPropertyWindow.JavaStaticText("Property_label_value").exist(SISW_MICRO_TIMEOUT) Then
													If trim(objPropertyWindow.JavaStaticText("Property_label_value").getROProperty("label")) <> trim(sValues(iCounter)) Then
															Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail: [ Fn_SISW_Prop_ObjectPropertyOperation ] Failed to verify [ " & trim(sProperty(iCounter)) & " = " & trim(sValues(iCounter)) & " ].")
															Fn_SISW_Prop_ObjectPropertyOperation = False
															Exit For
	
													End If
											ElseIF objPropertyWindow.JavaCheckBox("Property_field").exist(SISW_MICRO_TIMEOUT)Then
													If instr(trim(objPropertyWindow.JavaCheckBox("Property_field").getROProperty("label")), ":") > 0 Then
														If instr(trim(objPropertyWindow.JavaCheckBox("Property_field").getROProperty("label")), trim(sValues(iCounter))) <= 0 Then
																Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail: [ Fn_SISW_Prop_ObjectPropertyOperation ] Failed to verify [ " & trim(sProperty(iCounter)) & " = " & trim(sValues(iCounter)) & " ].")
																Fn_SISW_Prop_ObjectPropertyOperation = False
																Exit For
														End If
													Else
														If trim(objPropertyWindow.JavaCheckBox("Property_field").getROProperty("label")) <> trim(sValues(iCounter)) Then
																Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail: [ Fn_SISW_Prop_ObjectPropertyOperation ] Failed to verify [ " & trim(sProperty(iCounter)) & " = " & trim(sValues(iCounter)) & " ].")
																Fn_SISW_Prop_ObjectPropertyOperation = False
																Exit For
														End If
													End If
													
											End If
'								End If
'								Else
'								' list
'								' radio
'								' checkbox
'									' not yet implemented
'									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail: [ Fn_SISW_Prop_ObjectPropertyOperation ] Failed to verify [ " & trim(sProperty(iCounter)) & " = " & trim(sValues(iCounter)) & " ].")
'									Fn_SISW_Prop_ObjectPropertyOperation = False
'									Exit for
								End If
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [ Fn_SISW_Prop_ObjectPropertyOperation ] Successfully verified [ " & trim(sProperty(iCounter)) & " = " & trim(sValues(iCounter)) & " ].")
					Next
					' close dialog
					If objPropertyWindow.JavaButton("Cancel").Exist = True Then
							Call Fn_Button_Click(" Fn_SISW_Prop_ObjectPropertyOperation",objPropertyWindow,"Cancel")
					End If
				Case "ClickOnPropertyValue"
					If Fn_SISW_GetObject("Properties").Exist(SISW_MIN_TIMEOUT) = False AND Fn_SISW_GetObject("Properties1").Exist(SISW_MIN_TIMEOUT) = False Then
						call Fn_MenuOperation("KeyPress","View:Properties")	
					End If
					If Fn_SISW_GetObject("Properties").Exist(SISW_MICRO_TIMEOUT) Then
						Set objPropertyWindow = Fn_SISW_GetObject("Properties")
					ElseIf Fn_SISW_GetObject("Properties1").Exist(SISW_MICRO_TIMEOUT) Then
						Set objPropertyWindow = Fn_SISW_GetObject("Properties1")
					Else
						Set objPropertyWindow = Nothing
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail: [ Fn_SISW_Prop_ObjectPropertyOperation ] Failed to open Proeprty Winodw.")
						Exit function
					End IF
					
					' fetching screen resolution
					Set objWMIService = GetObject("winmgmts:\\" & "." & "\root\cimv2")
					Set colItems = objWMIService.ExecQuery("Select * From Win32_DisplayConfiguration")

					For Each objItem in colItems
						iXpos = objItem.PelsWidth       'Horizontal Resolution
						iYpos = objItem.PelsHeight      'Vertical Resolution
					Next
					' resizing property window to screen resolution
					objPropertyWindow.Resize iXpos, iYpos
						
					' second click on given link
					If sObjectName = "" Then 
						sObjectName = "All"
					End If
         
					objPropertyWindow.JavaStaticText("BottomLink").setTOProperty "label", sObjectName
					Call Fn_UI_JavaStaticText_Click("Fn_SISW_Prop_ObjectPropertyOperation", objPropertyWindow,"BottomLink",10, 10,"LEFT")
					' oObjetcTypeDictionary
					' if empty click on All
					sProperty = Split(oObjetcTypeDictionary("PropertyNames"),":")
					sValues =  Split(oObjetcTypeDictionary("PropertyValues"),":")
					Fn_SISW_Prop_ObjectPropertyOperation = false
					' loop - verify fieilds and values
					For iCounter = 0 to uBound(sProperty)
						objPropertyWindow.JavaStaticText("Property_label").setTOProperty "label", trim(sProperty(iCounter)) & ":"
						IF objPropertyWindow.JavaStaticText("Property_label_value").exist(SISW_MICRO_TIMEOUT) Then
							If trim(objPropertyWindow.JavaStaticText("Property_label_value").getROProperty("label")) = trim(sValues(iCounter)) Then
									objPropertyWindow.JavaStaticText("Property_label_value").Click 1, 1, "LEFT"
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail: [ Fn_SISW_Prop_ObjectPropertyOperation ] Failed to verify [ " & trim(sProperty(iCounter)) & " = " & trim(sValues(iCounter)) & " ].")
									Fn_SISW_Prop_ObjectPropertyOperation = True
									Exit For
							End If
						End If								
					Next
		
		'[TC1123(20161205c00)_PoonamC_NewDevelopment_13Mar2017:Added Cases "DblClickForm_Property_Edit_And_Save" & "DblClickForm_Property_Edit_And_SaveAndCheckIn" ]
		Case "DblClickForm_Property_Edit_And_Save","DblClickForm_Property_Edit_And_SaveAndCheckIn"'					# Code to handle Form Property Edit and Save
		
			   Set  objPropertyWindow = JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Form Properties")

'				#Use menu operations to open the Properties window.
				objPropertyWindow.SetTOProperty "title", sObjectName
				If objPropertyWindow.Exist(SISW_MIN_TIMEOUT) = False Then
					Fn_SISW_Prop_ObjectPropertyOperation = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail: [ Fn_SISW_Prop_ObjectPropertyOperation ] Winodw is not Exist.")
					Exit Function
				End If
				' Handle checkout dialog
				If Fn_UI_ObjectExist("Fn_SISW_Prop_ObjectPropertyOperation", objPropertyWindow.JavaButton("Check-OutAndEdit")) Then
					Call Fn_Button_Click(" Fn_SISW_Prop_ObjectPropertyOperation",objPropertyWindow ,"Check-OutAndEdit")
				End If
				Call Fn_ReadyStatusSync(1)
				Set objCheckOutDia = Fn_SISW_GetChkInChkOutObject("CheckOut")
				If TypeName(objCheckOutDia) <> "Nothing" Then
					Call Fn_Button_Click("Fn_SISW_Prop_ObjectProperty_Edit", objCheckOutDia,"Yes")
					Call Fn_ReadyStatusSync(1)
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: The Element ["+sObjectName+"]checked out  successfully")	
				End If
				Set objCheckOutDia = Nothing

				'	Get the keys & items count from data dictionary.
					dictItems = oObjetcTypeDictionary.Items
					dictKeys = oObjetcTypeDictionary.Keys
					For iCounter = 0 to oObjetcTypeDictionary.Count - 1
					   	 sObjElement = dictKeys(iCounter)
					   	 objPropertyWindow.JavaEdit("EditField").SetTOProperty "attached text",dictKeys(iCounter)+":"
						 If objPropertyWindow.JavaEdit("EditField").Exist(SISW_MIN_TIMEOUT) = False Then
							  	Fn_SISW_Prop_ObjectPropertyOperation = False
							  	Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail: [ Fn_SISW_Prop_ObjectPropertyOperation ] Proeprty not Exists on Winodw.")
							  	bFlag = False
						 Else
							  	bFlag = Fn_Edit_Box("Fn_SISW_Prop_ObjectPropertyOperation",objPropertyWindow,"EditField",dictItems(iCounter))
							  	Call Fn_ReadyStatusSync(1)
						 End IF
						  
						If bFlag = True Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: The property "& sObjElement  & " is modified with value ["+dictItems(iCounter)+"] successfully")
						Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: The property "& sObjElement  & " was not modified with value ["+dictItems(iCounter)+"].")
						End If
				Next
				
				If bFlag = True Then
						Fn_SISW_Prop_ObjectPropertyOperation = True
						If sAction = "DblClickForm_Property_Edit_And_SaveAndCheckIn" Then
								Call Fn_Button_Click(" Fn_SISW_Prop_ObjectPropertyOperation",objPropertyWindow,"Save and Check-In")
								Call Fn_ReadyStatusSync(1)
								Set objCheckOutDia = Fn_SISW_GetChkInChkOutObject("CheckIn")
								If TypeName(objCheckOutDia) <> "Nothing" Then
									Call Fn_Button_Click("Fn_SISW_Prop_ObjectProperty_Edit", objCheckOutDia,"Yes")
									Call Fn_ReadyStatusSync(1)
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: The Element ["+sObjectName+"] checked In  successfully")	
								End If
								Set objCheckOutDia = Nothing
						Else
								Call Fn_Button_Click(" Fn_SISW_Prop_ObjectPropertyOperation",objPropertyWindow,"Save")
								Call Fn_ReadyStatusSync(1)
								Call Fn_Button_Click(" Fn_SISW_Prop_ObjectPropertyOperation",objPropertyWindow,"Close")
								Call Fn_ReadyStatusSync(1)						
						End If	
				Else
					Fn_SISW_Prop_ObjectPropertyOperation = False
				End If
		'[TC1123(20161205c00)_PoonamC_NewDevelopment_13Mar2017:Added Cases "EditBox_IsEditable" to verify property is editable or not ]			
		Case "EditBox_IsEditable"
		
				Set  objPropertyWindow = JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Form Properties")

'				#Use menu operations to open the Properties window.
				objPropertyWindow.SetTOProperty "title", sObjectName
				If objPropertyWindow.Exist(SISW_MIN_TIMEOUT) = False Then
					Fn_SISW_Prop_ObjectPropertyOperation = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail: [ Fn_SISW_Prop_ObjectPropertyOperation ] Winodw is not Exist.")
					Exit Function
				End If
				' Handle checkout dialog
				If Fn_UI_ObjectExist("Fn_SISW_Prop_ObjectPropertyOperation", objPropertyWindow.JavaButton("Check-OutAndEdit")) Then
					Call Fn_Button_Click(" Fn_SISW_Prop_ObjectPropertyOperation",objPropertyWindow ,"Check-OutAndEdit")
				End If
				Call Fn_ReadyStatusSync(1)
				Set objCheckOutDia = Fn_SISW_GetChkInChkOutObject("CheckOut")
				If TypeName(objCheckOutDia) <> "Nothing" Then
					Call Fn_Button_Click("Fn_SISW_Prop_ObjectProperty_Edit", objCheckOutDia,"Yes")
					Call Fn_ReadyStatusSync(1)
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: The Element ["+sObjectName+"]checked out  successfully")	
				End If
				Set objCheckOutDia = Nothing

				'	Get the keys & items count from data dictionary.
				If vartype(oObjetcTypeDictionary) = "9" Then
					 dictItems = oObjetcTypeDictionary.Items
					 dictKeys = oObjetcTypeDictionary.Keys
					 For iCounter = 0 to oObjetcTypeDictionary.Count - 1
						   	 objPropertyWindow.JavaEdit("EditField").SetTOProperty "attached text",dictKeys(iCounter)+":"
							 If objPropertyWindow.JavaEdit("EditField").Exist(SISW_MIN_TIMEOUT) = False Then
								  	Fn_SISW_Prop_ObjectPropertyOperation = False
								  	Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail: [ Fn_SISW_Prop_ObjectPropertyOperation ] Proeprty not Exists on Winodw.")
								  	Exit For
							 Else
							 	    If cbool(dictItems(iCounter)) = cbool(objPropertyWindow.JavaEdit("EditField").GetROProperty("enabled")) Then
							 	    	Fn_SISW_Prop_ObjectPropertyOperation = True
							 	    Else
										Fn_SISW_Prop_ObjectPropertyOperation = False
										Exit For										
							 	    End If    
							 End IF	 
					Next	
					 Call Fn_Button_Click(" Fn_SISW_Prop_ObjectPropertyOperation",objPropertyWindow,"Close")	
					 Call Fn_ReadyStatusSync(1)
			  End If
			  
		End Select

Set objCheckOutDia = nothing
Set oPropSelect = Nothing
Set dictItems = Nothing
Set dictKeys = Nothing
Set iCounter = Nothing
Set sProperty = Nothing
Set sObjElement = Nothing
Set bFlag = Nothing
Set objPropertyWindow = nothing
Set ObjPropertyYes = nothing

End Function

'######################################################################################################################################
'###
'###    FUNCTION NAME   :  Fn_SISW_Prop_ObjectPropertyCkOutEditSave(sObjProperty,sObjPropertyValue)
'###
'###    DESCRIPTION     : Checkout Edit and Checkin operation for business object in My Teamcenter Navigator Tree.
'###
'###    PARAMETERS      : 1. sObjProperty-Valid Object Property Name
'###										2.sObjPropertyValue-Valid Object Property Value
'###                                         
'###    Function Calls  :  
'###
'###
'###	 HISTORY         :   		AUTHOR                 DATE        VERSION
'###
'###    CREATED BY      :     Ketan       		  		   29/05/10      1.0
'###
'###    REVIWED BY      :   							 	
'###
'###    MODIFIED BY     :
'###    EXAMPLE         :   Call Fn_SISW_Prop_ObjectPropertyCkOutEditSave("Description","Object Ready to Save")
'######################################################################################################################################

Function Fn_SISW_Prop_ObjectPropertyCkOutEditSave(sObjProperty,sObjPropertyValue)
		GBL_FAILED_FUNCTION_NAME="Fn_SISW_Prop_ObjectPropertyCkOutEditSave"
		'Check if Edit Properties Dialog Exist
		If not Fn_UI_ObjectExist("Fn_SISW_Prop_ObjectPropertyCkOutEditSave",JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Edit Properties"))Then
        		Call Fn_MenuOperation("Select","Edit:Properties")
		End If
		
		'Check If Dialog Check-Out Exist
		If not  Fn_UI_ObjectExist("Fn_SISW_Prop_ObjectPropertyCkOutEditSave",JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Edit Properties")) Then
				'Check-Out the Item	
				 Call Fn_ObjectCheckOut ("Menu CheckOut","", "", "","","","","" )		
        end If

		If Fn_UI_ObjectExist("Fn_SISW_Prop_ObjectPropertyCkOutEditSave", JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Edit Properties").JavaStaticText("BottomLink")) = True Then
			JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Edit Properties").JavaStaticText("BottomLink").Click 1,1,"LEFT"
		End If
		'Click on Static text
		'Call Fn_UI_JavaStaticText_Click(" Fn_TcObjectDelete",  JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Edit Properties"), "BottomLink", 1, 1, "LEFT")
        ' Set the Property Value
		Call Fn_Edit_Box("Fn_SISW_Prop_ObjectPropertyCkOutEditSave",JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Edit Properties"),sObjProperty,sObjPropertyValue)
		
		'Click on 'SaveAndCheck-In' Button
		Call Fn_Button_Click("Fn_SISW_Prop_ObjectPropertyCkOutEditSave",JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Edit Properties"),"Save")
		
		'Click on Close Button
		Call Fn_Button_Click("Fn_SISW_Prop_ObjectPropertyCkOutEditSave",JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Edit Properties"),"Close")
		Fn_SISW_Prop_ObjectPropertyCkOutEditSave = True

		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS:Property '"&sObjProperty&"' Successfully Modified and changed to "&sObjPropertyValue&" of Function Fn_SISW_Prop_ObjectPropertyCkOutEditSave" )

End Function

'#######################################################################################
'###     FUNCTION NAME   :   Fn_SISW_Prop_HTML_PropertyRetrieve(sPropertyName)
'###
'###    DESCRIPTION     :   This function retrieves the object property through HTML page
'### The Following Function Are Club Together In this function
'### 					Fn_HTML_PropertyInvoke
'### 					Fn_SISW_Prop_HTML_PropertyRetrieve
'### 					Fn_HTML_PropertyClose
'###
'###    PARAMETERS      :   sPropertyName
'###
'###    Return Value  :   PropertyValue/False  
'###
'###    HISTORY         :   			AUTHOR              		DATE        		VERSION
'###
'###    CREATED BY      :       Sunny 					07/06/2010   			1.0
'###
'###    REVIWED BY      :		Rizwan					09/06/2010
'###
'###    MODIFIED BY     :
'###    EXAMPLE         :   Fn_SISW_Prop_HTML_PropertyRetrieve("Created Date")
'#############################################################################################

Function Fn_SISW_Prop_HTML_PropertyRetrieve(sPropertyName)
GBL_FAILED_FUNCTION_NAME="Fn_SISW_Prop_HTML_PropertyRetrieve"

Dim ObjInv, arrPropertyName, arrPropertyVal, Rwcnt, objTable, iCounter, jRowConter, ObjClose, arrResult,  bCheckProp, iValueCounter , iArrSize        
    
Set ObjInv = JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet")
' ***********************Checking That  if both the windows are not already exists then call to Menu operation function***************
If  Fn_UI_ObjectExist("Fn_SISW_Prop_HTML_PropertyRetrieve",ObjInv.JavaDialog("Properties"))=False AND Fn_UI_ObjectExist("Fn_SISW_Prop_HTML_PropertyRetrieve",ObjInv.JavaDialog("Edit Properties"))=False Then
        Call Fn_MenuOperation("KeyPress","View:Properties") 				
End If
'******************************************Here checking that  properties window *****************************************************
		If  Fn_UI_ObjectExist("Fn_SISW_Prop_HTML_PropertyRetrieve",ObjInv.JavaDialog("Properties"))=True Then
				'If Fn_UI_ObjectExist("Fn_SISW_Prop_HTML_PropertyRetrieve",ObjInv.JavaDialog("Properties").JavaButton("Print"))=True Then
						Call Fn_Button_Click("Fn_SISW_Prop_HTML_PropertyRetrieve",ObjInv.JavaDialog("Properties"),"Print")
				'End if
		End If

		If Fn_UI_ObjectExist("Fn_SISW_Prop_HTML_PropertyRetrieve",ObjInv.JavaDialog("Edit Properties"))=True Then		
						Call Fn_Button_Click("Fn_SISW_Prop_HTML_PropertyRetrieve",ObjInv.JavaDialog("Edit Properties"),"Print")
		End If

				'******************************************Here checking that  Print window  Exists*****************************************************
				If Fn_UI_ObjectExist("Fn_SISW_Prop_HTML_PropertyRetrieve",ObjInv.JavaDialog("Print"))=True Then
                        Call Fn_CheckBox_Set("Fn_SISW_Prop_HTML_PropertyRetrieve", ObjInv.JavaDialog("Print"), "HTML","ON")
                        Call Fn_Button_Click("Fn_SISW_Prop_HTML_PropertyRetrieve",ObjInv.JavaDialog("Print"),"OpenInBrowser")
								 '********************************************Add Here browser check function************************************************************								
									If Browser("browser").page("Page").Exist(SISW_DEFAULT_TIMEOUT)Then							
'**********************************	**************************Property Verification start here -20-05-2010   ********************************************													
													'*********************************Split the number of parameters into an array**************************************
													arrPropertyName = Split(sPropertyName,":",-1)
													'arrPropertyVal = Split(sPropertyVal,":",-1)										
													'*************************************Checking that the Browser is opened or not************************************
													If  Browser("browser").Page("Page").Exist(SISW_MIN_TIMEOUT) Then
														'************************Here we count the Number of rows present in that perticular table**************************
														Rwcnt = Browser("Browser").Page("Page").WebTable("PropertyTable").RowCount
														Set objTable = Browser("Browser").Page("Page").WebTable("PropertyTable")														
														'*****Here is the loop for checking the first colume value if matches then it goes to the second if  condition*******
														
														iValueCounter  = 0
														iCounter = 0														
														iArrSize = Ubound(arrPropertyName)													
														arrResult =arrPropertyName
														
														Do
																  For jRowConter = 1 to Rwcnt
																		bCheckProp = false			  
																		If Trim(objTable.GetCellData(jRowConter,1)) = arrPropertyName(iCounter) Then
																			    bCheckProp = true
																				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Object Property Found for " + arrPropertyName(iCounter))      		
																				 Fn_SISW_Prop_HTML_PropertyRetrieve = Trim(objTable.GetCellData(jRowConter,2))                                                                          																							
																							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Object Value Verified for " +  arrPropertyName(iCounter))																							
																							iValueCounter = iValueCounter + 1
																							arrResult(iCounter) = true
																							Exit For
																	       End If																	
																	Next				
																			If bCheckProp = false Then
																				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Object Property Not Found for " + arrPropertyName(iCounter))      		
																				arrResult(iCounter) = false
																			End If
																	iCounter  = iCounter +1												
														Loop while iCounter <= iArrSize
												Else
															Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Browser window does not exist. ")      		

												End If                          				

												For iCounter = 0 to iArrSize
															if arrResult(iCounter) = false then
																	Fn_SISW_Prop_HTML_PropertyRetrieve = false
																	Exit for
															End If
												Next
																											
	'*******************************************'Property Verification Ends  here -20-05-2010************************************************************************************                                    									
									Else
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Print window does not exist. ")      		
									End If

				Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Properties window does not exist. ")      		
							Exit Function
				 End If
'********************************************** Close All the window code start  here -20-05-2010***********************************
                                        									
											Set ObjClose = JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet")
										'**************************************Check that browser is already open or not*************************************
											If browser("Browser").Page("Page").Exist(SISW_MIN_TIMEOUT) Then
														Browser("Browser").Close
														If Fn_UI_ObjectExist("Fn_SISW_Prop_HTML_PropertyRetrieve", ObjClose.JavaDialog("Print"))=True Then
															Call Fn_Button_Click("Fn_SISW_Prop_HTML_PropertyRetrieve", ObjClose.JavaDialog("Print"), "Close")
														End if
														'**************************************Check that  Edit Properties Dialog box is exists or not*************************
														If  Fn_UI_ObjectExist("Fn_SISW_Prop_HTML_PropertyRetrieve", ObjClose.JavaDialog("Edit Properties"))=True Then
																	Call Fn_Button_Click("Fn_SISW_Prop_HTML_PropertyRetrieve",ObjClose.JavaDialog("Edit Properties"), "Close")
												
																	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS:Closed Edit Properties  Dialog successfully.")
														'**********************************Check that  Properties Dialog box is exists or not**********************************
														Elseif Fn_UI_ObjectExist("Fn_SISW_Prop_HTML_PropertyRetrieve", ObjClose.JavaDialog("Properties"))=True Then
																	Call Fn_Button_Click("Fn_SISW_Prop_HTML_PropertyRetrieve",ObjClose.JavaDialog("Properties"), "Cancel")                                                                	 																
												
																   Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS:Closed Properties  Dialog successfully.")
														Else 
														'I************************that window is not  exist then it giver the error message**************************************															
												
																	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL:Couldn't find the Edit Properties or Properties dialog")
																	Exit Function 
														End If
											Else														
																	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL:Can't Find Browser")
											End If								
	'End If

	Set ObjInv = Nothing
	Set ObjClose = Nothing
	Set objTable = Nothing

End Function


'###########################################     (Function to Check Property Editable)      ###############################################
'#
'# 	Function Name		:				Fn_SISW_Prop_ObjectPropertyIsEditable()

'#	Description			 :		 		     Validate that the Object Property value filed associated to Property is editable 
'#
'#	Parameters			   :	 		   1) sPropertyName:Name of the Property
'#											
'#	Return Value		   : 				TRUE (if property is enabled)\ FALSE (if property is disabled)
'#
'#	Pre-requisite			:		 		  Object Properties window is open.
'#
'#	Examples				:			     Fn_SISW_Prop_ObjectPropertyIsEditable("Current Name") 
'#
'#	History:
'#	Developer Name			Date			Rev. No.			Changes Done			Reviewer	
'###############################################################################################################################
'#	Sunny Ruparel					15_06_10	1										Mohit Khare
''--------------------------------------------------------------------------------------------------------------------------------------------------------------
'#	Ashok kakade				17-May-2012			1.0				Modified Hierarchy of Dialog intNoOfObjects				Koustubh Watwe
'--------------------------------------------------------------------------------------------------------------------------------------------------------------
'#	Koustubh W					21-May-2012			1.0				Modified function
'###############################################################################################################################
Public Function Fn_SISW_Prop_ObjectPropertyIsEditable(sPropertyName)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_Prop_ObjectPropertyIsEditable"
	Dim objPropertyDiag,objDesc, intNoOfObjects, sPropName, arrPropName, i
	Set objDesc = Description.Create()
	objDesc("Class Name").value = "JavaEdit"
	Fn_SISW_Prop_ObjectPropertyIsEditable = False

	If JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Properties").Exist Then
		Set  objPropertyDiag = JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Properties")
	ElseIf JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Properties").Exist Then
		Set  objPropertyDiag = JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Properties") 
	Else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Faile : Fn_SISW_Prop_ObjectPropertyIsEditable : Property dialog does not exists.")
		Exit function
	End If

	Set  intNoOfObjects =  objPropertyDiag.ChildObjects(objDesc)
	If Fn_UI_ObjectExist("Fn_SISW_Prop_ObjectPropertyIsEditable",objPropertyDiag.JavaStaticText("BottomLink")) = True Then
			objPropertyDiag.JavaStaticText("BottomLink").Click 1,1,"LEFT"
	End If

	For i = 0 to intNoOfObjects.count-1
		sPropName = intNoOfObjects(i).getROProperty("attached text")
		arrPropName = Split(sPropName,":")
		If arrPropName(0) = sPropertyName Then
			If cInt(intNoOfObjects(i).getROProperty("editable")) = 1 Then
				Fn_SISW_Prop_ObjectPropertyIsEditable = True
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sPropertyName + " Property is editable")
			Else
				Fn_SISW_Prop_ObjectPropertyIsEditable = False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sPropertyName + " Property is not editable")
				Exit for
			End If
    	End If
  Next
  Set objDesc = Nothing
  Set intNoOfObjects = Nothing
  Set objPropertyDiag = Nothing
End Function

'#######################################################################################################################################################
'###    FUNCTION NAME   :   Fn_SISW_Prop_CommonModifiableProperties_Operation(sAction, aGlobalDictionary, sText)
'###
'###    DESCRIPTION     :   Addition and verification of Records in Common Modifiable Properties Dialog
'###
'###    PARAMETERS      :   sAction: Action string to navigate to appropriate case
'###											aGlobalDictionary: Array of Datadictionary (Collection of Properties)
'###											sText: For Future Use
'###
'###    Return Value  	:   True/False  
'###
'###    HISTORY         :   			AUTHOR              DATE        	VERSION
'###
'###    CREATED BY      :       Mahendra Bhandarkar 	14/06/2010   		 1.0
'###
'###    REVIWED BY      :		Mohit Khare				14/06/2010			 1.0
'###
'###    MODIFIED BY     :
'###    EXAMPLE         :   Dim aGlobalDictionary(2)
'###									Dim subDictionary1, subDictionary2
'###									Set subDictionary1 = CreateObject("Scripting.Dictionary")
'###									Set subDictionary2 = CreateObject("Scripting.Dictionary")
'###									With subDictionary1
'###											.Add "JavaEdit:Description", "Item Description 1"
'###											.Add "JavaEdit:Name", "IBC1"
'###									End with
'###									
'###									Set aGlobalDictionary(0) = subDictionary1
'###									
'###									With subDictionary2
'###										.Add "JavaEdit:Description", "Item Description 2"
'###										.Add "JavaEdit:Name", "IBC2"
'###									End with
'###									
'###									Set aGlobalDictionary(1) = subDictionary2
'###									
'###						Call	Fn_SISW_Prop_CommonModifiableProperties_Operation("Verify", aGlobalDictionary, "")
'###			
'###		Added two new cases : "SelectColumnEdit","SelectColumnAppendEdit"	Added By : Ketan On 12-Jan-2011
'###		Example Case "SelectColumnEdit" : Fn_SISW_Prop_CommonModifiableProperties_Operation("SelectColumnEdit", aGlobalDictionary, "Test111:A")
'###		Example Case "SelectColumnAppendEdit" : Fn_SISW_Prop_CommonModifiableProperties_Operation("SelectColumnAppendEdit", aGlobalDictionary, "ing:B")
'###		Case "ChangeOwningUser" added on 15-Jan-2011 by ketan
'###		Example Case "ChangeOwningUser" : Fn_SISW_Prop_CommonModifiableProperties_Operation("ChangeOwningUser", "", "Organization:dba:DBA:AutoTest7 (autotest7)")	
'###		Case "VerifyColumnData" added on 16-Jan-2011 by ketan
'###		Example Case "VerifyColumnData" : Msgbox Fn_SISW_Prop_CommonModifiableProperties_Operation("VerifyColumnData", "Owner", "AutoTest7 (autotest7)")
'###		Case "VerifyOpenedInEditMode" : Call Fn_SISW_Prop_CommonModifiableProperties_Operation("VerifyOpenedInEditMode", "", "")
'#######################################################################################################################################################
'History :	Developer Name		Date	Rev. No.	Changes Done														Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'			Koustubh W		24-Feb-11	  1.0		modified cases  SelectColumnEdit , SelectColumnAppendEdit
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'			Koustubh W		15-Jun-12	  1.0		modified object hierarchy
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'			Nilima Pandit	04-Feb-16	  1.0		Added New Case "VerifyOpenedInEditMode"								[Tc1122:2016011300:04Feb2016:AnkitN:NewDevelopment]
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'			Vivek Ahirrao	18-May-16	  1.1		Added New Case "EditColumnValues"									[TC1122-20160427-18_05_2016-VivekA-NewDevelopment]
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_Prop_CommonModifiableProperties_Operation(sAction, aGlobalDictionary, sText)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_Prop_CommonModifiableProperties_Operation"
	Dim objDialog, iTotalRow,  iTotalCol, iCounter1, iCounter4 ,iCounter2, iCounter3, objUser
	Dim sSubDictionary, aSplitKeyField, aKeyName, sSplitValueField, bFlag, iCountChkRow, iCountChkCol, sFunctionName
	Dim sData, iTextCnt
	Dim objDiag1, objDiag2, objCheckOutDia
	Dim StrTitle
	Dim dicCount, dicItems, dicKeys, aColumn, aButton
	Dim iCount1, sSubAction, sColumnField, iRowNum, iCount, sAppValue
	Dim aList , aArrayList
	Dim iCounter, iIterator

	StrTitle=JavaWindow("DefaultWindow").GetROProperty("title")

	Set objDiag1 = JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Common Modifiable Properties")
	Set objDiag2 = JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Common Modifiable Properties")
	Fn_SISW_Prop_CommonModifiableProperties_Operation = False
	If Fn_UI_ObjectExist("Fn_SISW_Prop_CommonModifiableProperties_Operation", objDiag2) = False AND Fn_UI_ObjectExist("Fn_SISW_Prop_CommonModifiableProperties_Operation", objDiag1) = False Then
		Call Fn_MenuOperation("Select", "View:Properties")
	End If
	
	If Fn_UI_ObjectExist("Fn_SISW_Prop_CommonModifiableProperties_Operation", objDiag2) = True Then
		Set objDialog = objDiag2
	ElseIf Fn_UI_ObjectExist("Fn_SISW_Prop_CommonModifiableProperties_Operation", objDiag1) = True Then
		Set objDialog = objDiag1
	Else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail : Fn_SISW_Prop_CommonModifiableProperties_Operation :Failed to find [ Common Modifiable Properties ] window")
		Exit function
	End If
	
	Wait 5
	
	'iTotalRow = objDialog.JavaTable("PropertyTable").GetROProperty("rows")
	'iTotalCol = objDialog.JavaTable("PropertyTable").GetROProperty("cols")
	 iTotalRow = Fn_UI_Object_GetROProperty("Fn_SISW_Prop_CommonModifiableProperties_Operation",objDialog.JavaTable("PropertyTable"), "rows")
	 iTotalCol = Fn_UI_Object_GetROProperty("Fn_SISW_Prop_CommonModifiableProperties_Operation",objDialog.JavaTable("PropertyTable"), "cols")
	Select Case sAction
		Case "Verify"
				bFlag = False
				iCountChkRow  = 0
				' Block for Table Rows
				For iCounter1 = 0 To iTotalRow - 1
					' Block for fetching records from an array
					'	For iCounter2 = 0 To iTotalCol - 1
					Set sSubDictionary = aGlobalDictionary(iCounter1)
					' Block for matching the records with the column fields
					For iCounter3 = 0 To iTotalCol - 1
						iCountChkCol = 0
						aKeyName = sSubDictionary.Keys
						sSplitValueField = sSubDictionary.Items
						iCounter2 = UBound(sSplitValueField) + 1
						For iCounter4 = 0 To UBound(aKeyName)
							aSplitKeyField = Split(aKeyName(iCounter4), ":")
							If objDialog.JavaTable("PropertyTable").GetColumnName(iCounter3) =  aSplitKeyField(1)Then
								objDialog.JavaTable("PropertyTable").SelectCell iCounter1, iCounter3
								If aSplitKeyField(0) = "JavaEdit" Then
									objDialog.JavaEdit("PropertyValue").SetTOProperty "attached text", objDialog.JavaTable("PropertyTable").GetColumnName(iCounter3)
									If objDialog.JavaEdit("PropertyValue").GetROProperty("value") = sSplitValueField(iCounter4) Then
										iCountChkCol = iCountChkCol + 1
										Call Fn_WriteLogFile(Environment.Value("TestLogFile")," Property ["+CStr(aSplitKeyField(1))+"] value ["+CStr(sSplitValueField(iCounter4))+"] Verification done successfully in Function " & sFunctionName)
									End If	 
								Else 
									If aSplitKeyField(0) = "JavaRadioButton" Then
										objDialog.JavaRadioButton("BlnPropertyValue").SetTOProperty "attached text", sSplitValueField(iCounter4)
										If objDialog.JavaRadioButton("BlnPropertyValue").GetROProperty("label") =  sSplitValueField(iCounter4) Then
											iCountChkCol = iCountChkCol + 1
											Call Fn_WriteLogFile(Environment.Value("TestLogFile")," Property ["+CStr(aSplitKeyField(1))+"] value ["+CStr(sSplitValueField(iCounter4))+"] Verification done successfully in Function " & sFunctionName)
										End If
								   	Else					
								   	End If
								End If
	
							   	' Block for Comparing the Keys and  the Column Header
								'				If objDialog.JavaTable("PropertyTable").GetColumnName(iCounter3) = aSplitKeyField(1)  Then
								'               ' Block for Comparing the Value and  the Dictionary Value
								'						If objDialog.JavaTable("PropertyTable").Object.getValueAt(iCounter1, iCounter3) = sSplitValueField(iCounter4) Then
								'								iCountChkCol = iCountChkCol + 1
								'						End If
								'				End If
							End if
						Next
						If iCountChkCol = UBound(sSplitValueField) Then
							iCountChkRow = iCountChkRow + 1
						End If
					Next
					'Next
				Next
				
				If Cint(iCountChkRow/iCounter2) = UBound(aGlobalDictionary) Then
					bFlag = True
				End If
	
				If bFlag = True Then
					Fn_SISW_Prop_CommonModifiableProperties_Operation = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile")," All Properties Verified successfully in Function " & sFunctionName)
				Else
					Fn_SISW_Prop_CommonModifiableProperties_Operation = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile")," All Properties are not Verified in Function " & sFunctionName)		
				End If
	
		Case "Add"
				iCountChkRow  = 0
				' Block for Table Rows
				For iCounter1 = 0 To iTotalRow - 1
					' Block for fetching records from an array
					'	For iCounter2 = 0 To iTotalCol - 1
					Set	sSubDictionary = aGlobalDictionary(iCounter1)
					' Block for matching the records with the column fields
					For iCounter3 = 0 To iTotalCol - 1
						iCountChkCol = 1
						aKeyName = sSubDictionary.Keys
						sSplitValueField = sSubDictionary.Items
						For iCounter4 = 0 To UBound(aKeyName)
							aSplitKeyField = Split(aKeyName(iCounter4), ":")
							If objDialog.JavaTable("PropertyTable").GetColumnName(iCounter3) =  aSplitKeyField(1)Then
								objDialog.JavaTable("PropertyTable").SelectCell iCounter1, iCounter3
							   	' Block for Comparing the Keys and  the Column Header
								If aSplitKeyField(0) = "JavaEdit" Then
									objDialog.JavaEdit("PropertyValue").SetTOProperty "attached text", objDialog.JavaTable("PropertyTable").GetColumnName(iCounter3)
									Call Fn_Edit_Box("Fn_Common_Modifiable_Properties_Operation",objDialog,"PropertyValue", sSplitValueField(iCounter4))
									Call Fn_Button_Click("Fn_Common_Modifiable_Properties_Operation", objDialog, "SubmitChanges")
									Call Fn_WriteLogFile(Environment.Value("TestLogFile")," Property ["+CStr(aSplitKeyField(1))+"] value ["+CStr(sSplitValueField(iCounter4))+"] is updated successfully in Function " & sFunctionName)
									iCountChkRow = iCountChkRow + 1 
							   	Else 
							   		If aSplitKeyField(0) = "JavaRadioButton" Then
										objDialog.JavaRadioButton("BlnPropertyValue").SetTOProperty "attached text", sSplitValueField(iCounter4)
										If sSplitValueField(iCounter4) = "True" Then
											objDialog.JavaRadioButton("BlnPropertyValue").Set "ON"
										Else
											objDialog.JavaRadioButton("BlnPropertyValue").Set "OFF"
										End If
										Call Fn_Button_Click("Fn_Common_Modifiable_Properties_Operation", objDialog, "SubmitChanges")
										Call Fn_WriteLogFile(Environment.Value("TestLogFile")," Property ["+CStr(aSplitKeyField(1))+"] value ["+CStr(sSplitValueField(iCounter4))+"] is updated successfully in Function " & sFunctionName)
							   		Else
							   		End If
							   	End If
							End If ' If clause end
						Next
					Next
					'Next
				Next
	
				If iCountChkRow = UBound(aGlobalDictionary) Then
					bFlag = True
				End If
	
				If bFlag = True Then
					Fn_SISW_Prop_CommonModifiableProperties_Operation = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile")," All Properties are Updated successfully in Function " & sFunctionName)		
				Else
					Fn_SISW_Prop_CommonModifiableProperties_Operation = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile")," All Properties are not updated successfully in Function " & sFunctionName)		
				End If
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "SelectColumnEdit","SelectColumnAppendEdit"
				aSplitKeyField = aGlobalDictionary(0).Keys
				sSplitValueField = aGlobalDictionary(0).Items
				aKeyName = Split(sText,":",-1,1)
				iTextCnt = 0
				For iCounter1 = 0 to Ubound(aSplitKeyField)
					'If Trim(Lcase(sSplitValueField(iCounter1))) = "javaedit" Then
					'	objDialog.JavaEdit("PropertyValue").SetTOProperty "attached text",aSplitKeyField(iCounter1)
					'	If sAction="SelectColumnEdit" Then
					'		'Set property value
					'		Call Fn_UI_EditBox_Type("Fn_SISW_Prop_CommonModifiableProperties_Operation",objDialog,"PropertyValue",aKeyName(iCounter1))
					'		'Click on Submit Changes button
					'		Call Fn_Button_Click("Fn_Common_Modifiable_Properties_Operation", objDialog, "SubmitChanges")
					'	ElseIf sAction = "SelectColumnAppendEdit" Then
					'		'Get value from property value editbox
					'		iCounter2 = Fn_Edit_Box_GetValue("Fn_Common_Modifiable_Properties_Operation",objDialog,"PropertyValue")
					'		'Set property value
					'		Call Fn_UI_EditBox_Type("Fn_SISW_Prop_CommonModifiableProperties_Operation",objDialog,"PropertyValue",iCounter2+aKeyName(iCounter1))
					'		'Click on Submit Changes button
					'		Call Fn_Button_Click("Fn_Common_Modifiable_Properties_Operation", objDialog, "SubmitChanges")
					'	End If
					'End If
					
					Select Case Trim(Lcase(sSplitValueField(iCounter1)))
						Case "javaedit"
								bFlag = Fn_UI_JavaTable_ClickColumnHeader("Fn_SISW_Prop_CommonModifiableProperties_Operation", objDialog, "PropertyTable",aSplitKeyField(iCounter1),"LEFT","")
								If bFlag = True Then
									objDialog.JavaEdit("PropertyValue").SetTOProperty "attached text",aSplitKeyField(iCounter1)
									If sAction="SelectColumnEdit" Then
										'Set property value
										Call Fn_UI_EditBox_Type("Fn_SISW_Prop_CommonModifiableProperties_Operation",objDialog,"PropertyValue",aKeyName(iTextCnt))
										'Click on Submit Changes button
										Call Fn_Button_Click("Fn_Common_Modifiable_Properties_Operation", objDialog, "SubmitChanges")
									ElseIf sAction = "SelectColumnAppendEdit" Then
										'Get value from property value editbox
										iCounter2 = Fn_Edit_Box_GetValue("Fn_Common_Modifiable_Properties_Operation",objDialog,"PropertyValue")
										'Set property value
										Call Fn_UI_EditBox_Type("Fn_SISW_Prop_CommonModifiableProperties_Operation",objDialog,"PropertyValue",iCounter2+aKeyName(iTextCnt))
										'Click on Submit Changes button
										Call Fn_Button_Click("Fn_Common_Modifiable_Properties_Operation", objDialog, "SubmitChanges")
									End If
									iTextCnt = iTextCnt + 1
								Else 
									Fn_SISW_Prop_CommonModifiableProperties_Operation = False
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail : Fn_SISW_Prop_CommonModifiableProperties_Operation function failed as "& aSplitKeyField(iCounter1) &" column not found.")
									Exit Function
								End If
						Case "javabutton"
								Call Fn_Button_Click("Fn_Common_Modifiable_Properties_Operation", objDialog, aSplitKeyField(iCounter1))
								If Trim(Lcase(aSplitKeyField(iCounter1))) = "check-out and edit" Then
									Set objCheckOutDia = Fn_SISW_GetChkInChkOutObject("CheckOut")
									If TypeName(objCheckOutDia) <> "Nothing" Then
										Call Fn_Button_Click("Fn_Common_Modifiable_Properties_Operation", objCheckOutDia,"Yes")
										wait 5
									End If
									Set objCheckOutDia = Nothing
								end if
								' click on table header
					End Select
				Next
				'Click on Buttons
				aSplitKeyField = aGlobalDictionary(1).Keys
				sSplitValueField = aGlobalDictionary(1).Items
				For iCounter1 = 0 to Ubound(aSplitKeyField)
					Call Fn_Button_Click("Fn_Common_Modifiable_Properties_Operation", objDialog, sSplitValueField(iCounter1))
					Select Case Trim(Lcase(sSplitValueField(iCounter1))) 
						Case "save and check-in"
		                		Set objCheckOutDia = Fn_SISW_GetChkInChkOutObject("CheckIn")
								If TypeName(objCheckOutDia) <> "Nothing" Then
									Call Fn_Button_Click("Fn_Common_Modifiable_Properties_Operation", objCheckOutDia,"Yes")
								End If
								Set objCheckOutDia = Nothing
						Case "cancel check-out"
								Set objCheckOutDia = Fn_SISW_GetChkInChkOutObject("CancelCheckOut")
								If TypeName(objCheckOutDia) <> "Nothing" Then
									Call Fn_Button_Click("Fn_Common_Modifiable_Properties_Operation", objCheckOutDia,"Yes")
								End If
								Set objCheckOutDia = Nothing
					End Select
				Next
				Fn_SISW_Prop_CommonModifiableProperties_Operation = True
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Pass : Fn_SISW_Prop_CommonModifiableProperties_Operation completed successfully with action"& sAction)
	
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "SelectColumnEditCancelCheckOut"
				aSplitKeyField = aGlobalDictionary(0).Keys
				sSplitValueField = aGlobalDictionary(0).Items
				
				aKeyName = Split(sText,":",-1,1)				
				iTextCnt = 0
				For iCounter1 = 0 to Ubound(aSplitKeyField)					
					Select Case Trim(Lcase(sSplitValueField(iCounter1)))
						Case "javaedit"
								bFlag = Fn_UI_JavaTable_ClickColumnHeader("Fn_SISW_Prop_CommonModifiableProperties_Operation", objDialog, "PropertyTable",aSplitKeyField(iCounter1),"LEFT","")
								If bFlag = True Then
									objDialog.JavaEdit("PropertyValue").SetTOProperty "attached text",aSplitKeyField(iCounter1)
									'Set property value
									Call Fn_UI_EditBox_Type("Fn_SISW_Prop_CommonModifiableProperties_Operation",objDialog,"PropertyValue",aKeyName(iTextCnt))
									'Click on Submit Changes button
									Call Fn_Button_Click("Fn_Common_Modifiable_Properties_Operation", objDialog, "SubmitChanges")
									iTextCnt = iTextCnt + 1
								Else 
									Fn_SISW_Prop_CommonModifiableProperties_Operation = False
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail : Fn_SISW_Prop_CommonModifiableProperties_Operation function failed as "& aSplitKeyField(iCounter1) &" column not found.")
									Exit Function
								End If
						Case "javabutton"
								Call Fn_Button_Click("Fn_Common_Modifiable_Properties_Operation", objDialog, aSplitKeyField(iCounter1))
								If Trim(Lcase(aSplitKeyField(iCounter1))) = "check-out and edit" Then
									Set objCheckOutDia = Fn_SISW_GetChkInChkOutObject("CheckOut")
									If TypeName(objCheckOutDia) <> "Nothing" Then
										Call Fn_Button_Click("Fn_Common_Modifiable_Properties_Operation", objCheckOutDia,"Yes")
										wait 5
										Set objCheckOutDia = Nothing
									end if
								end if								
					End Select
				Next
				Wait 3
								
				'Click on Buttons
				aSplitKeyField = aGlobalDictionary(1).Keys
				sSplitValueField = aGlobalDictionary(1).Items
				
				For iCounter1 = 0 to Ubound(aSplitKeyField)
					If instr(sSplitValueField(iCounter1),"~") Then
						aList = Split(sSplitValueField(iCounter1),"~")
						aArrayList = aList(1)
						sSplitValueField(iCounter1) = aList(0)
					End If
					Select Case Trim(Lcase(sSplitValueField(iCounter1))) 
						Case "cancel check-out"
								Call Fn_Button_Click("Fn_Common_Modifiable_Properties_Operation", objDialog, sSplitValueField(iCounter1))
								Set objCheckOutDia = Fn_SISW_GetChkInChkOutObject("CancelCheckOut")
								Call Fn_Button_Click("Fn_Common_Modifiable_Properties_Operation", objCheckOutDia,"No")
								Set objCheckOutDia = Nothing
						Case "cancel check-out and verifyobjectlist"
								Call Fn_Button_Click("Fn_Common_Modifiable_Properties_Operation", objDialog, "cancel check-out")
								Set objCheckOutDia = Fn_SISW_GetChkInChkOutObject("CancelCheckOut")
								aArrayList = Split(aArrayList,",")
								For iCounter = 0 to Ubound(aArrayList)
									bFlag = False
									For iIterator = 0 to Cint(objCheckOutDia.JavaTable("ObjectInfo").GetROProperty("rows"))-1
										If Trim(Fn_SISW_UI_JavaTable_Operations("", "GetCellData", objCheckOutDia, "ObjectInfo", "Object.GetItem", "", iIterator, "", "", "", "")) = Trim(aArrayList(iCounter)) Then
											bFlag = True
											Exit For
										End If										
									Next
									If bFlag = False Then
										Exit For
										Exit Function
									End If
								Next								
								Set objCheckOutDia = Nothing					   
					End Select
				Next
				
				'Click on Buttons				
				Fn_SISW_Prop_CommonModifiableProperties_Operation = True
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Pass : Fn_SISW_Prop_CommonModifiableProperties_Operation completed successfully with action"& sAction)
		
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "VerifyButtonEnableAndClick"
				Dim objBtDialog
				'Verify button is enabled
				aSplitKeyField = aGlobalDictionary(0).Keys
				sSplitValueField = aGlobalDictionary(0).Items
				For iCounter1 = 0 to Ubound(aSplitKeyField)					
					Select Case Trim(Lcase(aSplitKeyField(iCounter1)))
						Case "checkout"
						Case "checkin"
						Case "cancelcheckout"
							Set objBtDialog = Fn_SISW_GetChkInChkOutObject("CancelCheckOut")
						Case "commonmodifiableproperties"
							Set objBtDialog = objDialog
					End Select	
					sText = Split(sSplitValueField(iCounter1),"~")			
					For iCounter = 0 to Ubound(sText)
						bFlag = False
						If sText(iCounter) = "OK" Then
							sText(iCounter) = "Yes"
						ElseIf sText(iCounter) = "Cancel" Then
							sText(iCounter) = "No"
						End If
						If Fn_UI_Object_GetROProperty("Fn_SISW_Prop_CommonModifiableProperties_Operation",objBtDialog.JavaButton(sText(iCounter)), "enabled") = "1" Then
							bFlag = True	
						Else
							Exit For
						End If	
					Next
					If bFlag = False Then
						Exit For
						Exit Function
					End If
					'Click on Button
					aSplitKeyField = aGlobalDictionary(1).Keys
					sSplitValueField = aGlobalDictionary(1).Items
					Select Case Trim(Lcase(sSplitValueField(iCounter1)))
						Case "ok"
							Fn_SISW_Prop_CommonModifiableProperties_Operation = Fn_Button_Click("Fn_Common_Modifiable_Properties_Operation", objBtDialog,"Yes")				
						Case "cancel"
							Fn_SISW_Prop_CommonModifiableProperties_Operation = Fn_Button_Click("Fn_Common_Modifiable_Properties_Operation", objBtDialog,"No")				
						Case "close"
							Fn_SISW_Prop_CommonModifiableProperties_Operation = Fn_Button_Click("Fn_Common_Modifiable_Properties_Operation", objBtDialog,"Close")				
					End Select
				Next
				Set objBtDialog = Nothing
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "ChangeOwningUser"
				'Click on Column header "Owner".
				Call Fn_UI_JavaTable_ClickColumnHeader("Fn_SISW_Prop_CommonModifiableProperties_Operation", objDialog, "PropertyTable","Owner","LEFT","")
				'Click on Additional Options
				Call Fn_Button_Click("Fn_Common_Modifiable_Properties_Operation", objDialog, "AdditionalOptions")
				Set objUser = Fn_UI_ObjectCreate("Fn_SISW_Prop_CommonModifiableProperties_Operation", JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("ChangeOwnerGroup"))
				'Split sText  as User name which is to be given the ownership.
				aSplitKeyField = Split(sText,"(")
				aKeyName = Split(aSplitKeyField(1),")")
				'Set the username in the search criteria Editbox.
				Call Fn_Edit_Box("Fn_Common_Modifiable_Properties_Operation",objUser,"SearchCriteria", aKeyName(0))
				'Click on Find Users
				Call Fn_Button_Click("Fn_Common_Modifiable_Properties_Operation", objUser, "FindUser")
				'Select the User from Organization Tree.
				Call Fn_JavaTree_Select("Fn_Common_Modifiable_Properties_Operation", objUser, "OrganizationTree",sText)
				'Click on OK button
				Call Fn_Button_Click("Fn_Common_Modifiable_Properties_Operation", objUser, "OK")
				'Click on Save and Check-In button
				Call Fn_Button_Click("Fn_Common_Modifiable_Properties_Operation", objDialog, "save and check-in")
				'Check-In the changes
				Set objCheckOutDia = Fn_SISW_GetChkInChkOutObject("CheckIn")
				 If TypeName(objCheckOutDia) <> "Nothing" Then
					Call Fn_Button_Click("Fn_Common_Modifiable_Properties_Operation", objCheckOutDia,"Yes")
				 End If
				 Set objCheckOutDia = Nothing
				Fn_SISW_Prop_CommonModifiableProperties_Operation = True
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Pass : Fn_SISW_Prop_CommonModifiableProperties_Operation completed successfully with action"& sAction)
	
		Case "VerifyColumnData"
				bFlag = 0
				For iCounter1 = 0 to iTotalCol-1
					If Trim(Lcase(objDialog.JavaTable("PropertyTable").GetColumnName(iCounter1))) = Trim(Lcase(aGlobalDictionary)) Then
						Exit For
					End If
				Next
				For iCounter2 = 0 to iTotalRow-1
					sData = Lcase(objDialog.JavaTable("PropertyTable").GetCellData(iCounter2,iCounter1))
					If sData = "" Then
						sData = Lcase(objDialog.JavaTable("PropertyTable").Object.getValueAt(iCounter2,iCounter1).toString)
					End If
					If Trim(sData) = Trim(Lcase(sText)) Then
						bFlag = bFlag+1
					End If
				Next
				If Cint(bFlag) = Cint(iTotalRow) Then
					Fn_SISW_Prop_CommonModifiableProperties_Operation = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),sText&" found succesfully in all rows in column "&aGlobalDictionary)
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Pass : Fn_SISW_Prop_CommonModifiableProperties_Operation completed successfully with action"& sAction)
				Else
					Fn_SISW_Prop_CommonModifiableProperties_Operation = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),sText&" not found in all rows in column "&aGlobalDictionary)
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail : Fn_SISW_Prop_CommonModifiableProperties_Operation failed")
				End If
				'Close the Common properties dialog
				objDialog.Close
				
		'[Tc1122:2016011300:04Feb2016:AnkitN:NewDevelopment] - Added Case to Properties Dialog opened in Edit mode 
		Case "VerifyOpenedInEditMode"			
				bFlag = 0	
				objDialog.JavaTable("PropertyTable").SelectCell 1,1
				bFlag=objDialog.JavaButton("SubmitChanges").GetROProperty("enabled")
				If bFlag = 1 Then
					Fn_SISW_Prop_CommonModifiableProperties_Operation = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Pass : Fn_SISW_Prop_CommonModifiableProperties_Operation completed successfully with action"& sAction)
				Else
					Fn_SISW_Prop_CommonModifiableProperties_Operation = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail : Fn_SISW_Prop_CommonModifiableProperties_Operation failed")
				End If	
		
		'--------------------------------------------------------------------------------------------------------------------------------------------------
		'[TC1122-20160427-13_05_2016-VivekA-NewDevelopment] - Added for RM Testcases
		'Added case to edit perticular row and column value
		'Example for 2 selected rows
		'	Set aGlobalDictionary = CreateObject("Scripting.Dictionary")
		'	With aGlobalDictionary
		'		.Add "SelectTab", "BOMLines"						'Select Tab
		'		.Add "GetRowNumber1", "Item Id:000132"			'This case is mandatory to decide which row u want to update, it depends on Item Id or any other column u want
		'		.Add "EditBox1", "Rev Description:Description1"	'These are the column names which u want to update, on row return by case "GetRowNumber1"
		'		.Add "EditBox2", "Rev Name:Test1"
		'		.Add "GetRowNumber2", "Item Id:000133"			'This case is mandatory to decide which row u want to update, it depends on Item Id or any other column u want
		'		.Add "EditBox3", "Rev Description:Description2"	'These are the column names which u want to update, on row return by case "GetRowNumber2"
		'		.Add "EditBox4", "Rev Name:Test2"
		'		.Add "Button", "Apply:OK"						'This is mandatory case to apply the changes done in column value and then OK button to set
		'	End with
		' It works only for Mutliple selected Rows in BOM Table
		Case "EditColumnValues"	
				If varType(aGlobalDictionary) <> "9" Then
					Set objDialog = Nothing
					Fn_SISW_Prop_CommonModifiableProperties_Operation = False
					Exit Function
				End If
				
				dicCount = aGlobalDictionary.Count
				dicItems = aGlobalDictionary.Items
				dicKeys = aGlobalDictionary.Keys
				
				iTotalRow = objDialog.JavaTable("PropertyTable").GetROProperty("rows")
				iTotalCol = objDialog.JavaTable("PropertyTable").GetROProperty("cols")
				
				For iCount1 = 0 To dicCount - 1
					If Instr(dicKeys(iCount1),"EditBox")>0 Then
						sSubAction = "EditBox"
					ElseIf Instr(dicKeys(iCount1),"GetRowNumber")>0 Then
						sSubAction = "GetRowNumber"
					ElseIf Instr(dicKeys(iCount1),"Button")>0 Then
						sSubAction = "Button"
					Else
						sSubAction = dicKeys(iCount1)
					End If
					
					sColumnField = dicItems(iCount1)
					
					Select Case sSubAction
						'Case to Select Tab BOMLines or Item Revisions
						Case "SelectTab"
								If sColumnField<>"" Then
									objDialog.JavaTab("PropertyNameTab").Select sColumnField
									If Err.Number < 0 Then
										Set objDialog = Nothing
										Fn_SISW_Prop_CommonModifiableProperties_Operation = False
										Exit Function
									End If
									Fn_SISW_Prop_CommonModifiableProperties_Operation = True
								End If
						'Case to get Row number of value present in column provided,
						'This row number will be used in EditBox case to edit the column value
						Case "GetRowNumber"
								If sColumnField <> "" Then
									iRowNum = ""
									aColumn = Split(sColumnField,":")
									bFlag = False
									For iCount = 0 To iTotalRow
										objDialog.JavaTable("PropertyTable").SelectCell iCount,aColumn(0)
										Wait 1
										sAppValue = objDialog.JavaTable("PropertyTable").GetCellData(iCount,aColumn(0))
										If Instr(sAppValue,aColumn(1))>0 Then
											iRowNum = iCount
											bFlag = True
											Exit For
										End If
									Next
									If bFlag = False Then
										Set objDialog = Nothing
										Fn_SISW_Prop_CommonModifiableProperties_Operation = False
										Exit Function
									End If
									Fn_SISW_Prop_CommonModifiableProperties_Operation = iRowNum
								End If
						Case "EditBox"
								If sColumnField<>"" Then
									aColumn = Split(sColumnField,":")
									'Select Cell which u want to edit
									objDialog.JavaTable("PropertyTable").SelectCell iRowNum,aColumn(0)
									Wait 1
									objDialog.JavaEdit("PropertyValue").SetTOProperty "attached text",aColumn(0)
									Wait 0,500
									bFlag = Fn_UI_Object_GetROProperty("Fn_SISW_Prop_CommonModifiableProperties_Operation",objDialog.JavaEdit("PropertyValue"),"enabled")
									If bFlag <> "1" Then
										Set objDialog = Nothing
										Fn_SISW_Prop_CommonModifiableProperties_Operation = False
										Exit Function
									End If
									Call Fn_UI_EditBox_Type("Fn_SISW_Prop_CommonModifiableProperties_Operation",objDialog,"PropertyValue",aColumn(1))
									Wait 1
									'Click on Submit Changes button to update the value
									Call Fn_Button_Click("Fn_SISW_Prop_CommonModifiableProperties_Operation",objDialog,"SubmitChanges")
									Fn_SISW_Prop_CommonModifiableProperties_Operation = True
								End If
						Case "Button"
								aButton = Split(sColumnField,":")
								For iCount = 0 To UBound(aButton)
									bFlag = Fn_Button_Click("Fn_SISW_Prop_CommonModifiableProperties_Operation",objDialog,aButton(iCount))
									If bFlag = False Then
										Set objDialog = Nothing
										Fn_SISW_Prop_CommonModifiableProperties_Operation = False
										Exit Function
									End If
									Wait 1
								Next
								Fn_SISW_Prop_CommonModifiableProperties_Operation = True
					End Select
				Next
				Set objDialog = Nothing
	'--------------------------------------------------------------------------------------------------------------------------------------------------			
		'[TC11.4_2017080100_NewDevelopment_PoonamC_22Aug2017_Added Case to verify cell value for corresponding row - Added case for RM TestCases]		
		Case "VerifyColumnValues"	
				If varType(aGlobalDictionary) <> "9" Then
					Set objDialog = Nothing
					Fn_SISW_Prop_CommonModifiableProperties_Operation = False
					Exit Function
				End If
				
				dicCount = aGlobalDictionary.Count
				dicItems = aGlobalDictionary.Items
				dicKeys = aGlobalDictionary.Keys
				
				iTotalRow = objDialog.JavaTable("PropertyTable").GetROProperty("rows")
				iTotalCol = objDialog.JavaTable("PropertyTable").GetROProperty("cols")
				
				For iCount1 = 0 To dicCount - 1
					If Instr(dicKeys(iCount1),"EditBox")>0 Then
						sSubAction = "EditBox"
					ElseIf Instr(dicKeys(iCount1),"GetRowNumber")>0 Then
						sSubAction = "GetRowNumber"
					ElseIf Instr(dicKeys(iCount1),"Button")>0 Then
						sSubAction = "Button"
					Else
						sSubAction = dicKeys(iCount1)
					End If
					
					sColumnField = dicItems(iCount1)
					
					Select Case sSubAction
						'Case to Select Tab BOMLines or Item Revisions
						Case "SelectTab"
								If sColumnField<>"" Then
									objDialog.JavaTab("PropertyNameTab").Select sColumnField
									If Err.Number < 0 Then
										Set objDialog = Nothing
										Fn_SISW_Prop_CommonModifiableProperties_Operation = False
										Exit Function
									End If
									Fn_SISW_Prop_CommonModifiableProperties_Operation = True
								End If
						'Case to get Row number of value present in column provided,
						'This row number will be used in EditBox case to edit the column value
						Case "GetRowNumber"
								If sColumnField <> "" Then
									iRowNum = ""
									aColumn = Split(sColumnField,":")
									bFlag = False
									For iCount = 0 To iTotalRow
										objDialog.JavaTable("PropertyTable").SelectCell iCount,aColumn(0)
										Wait 1
										sAppValue = objDialog.JavaTable("PropertyTable").GetCellData(iCount,aColumn(0))
										If Instr(sAppValue,aColumn(1))>0 Then
											iRowNum = iCount
											bFlag = True
											Exit For
										End If
									Next
									If bFlag = False Then
										Set objDialog = Nothing
										Fn_SISW_Prop_CommonModifiableProperties_Operation = False
										Exit Function
									End If
									Fn_SISW_Prop_CommonModifiableProperties_Operation = iRowNum
								End If
						Case "EditBox"
								If sColumnField<>"" Then
									aColumn = Split(sColumnField,":")
									'Select Cell which u want to edit
									objDialog.JavaTable("PropertyTable").SelectCell iRowNum,aColumn(0)
									Wait 1
									objDialog.JavaEdit("PropertyValue").SetTOProperty "attached text",aColumn(0)
									Wait 0,500
									sData = Fn_UI_Object_GetROProperty("Fn_SISW_Prop_CommonModifiableProperties_Operation",objDialog.JavaEdit("PropertyValue"),"value")
									If trim(sData) <> trim(aColumn(1)) Then
										Set objDialog = Nothing
										Fn_SISW_Prop_CommonModifiableProperties_Operation = False
										Exit Function
									End If
									Fn_SISW_Prop_CommonModifiableProperties_Operation = True
								End If
						Case "Button"
								aButton = Split(sColumnField,":")
								For iCount = 0 To UBound(aButton)
									bFlag = Fn_Button_Click("Fn_SISW_Prop_CommonModifiableProperties_Operation",objDialog,aButton(iCount))
									If bFlag = False Then
										Set objDialog = Nothing
										Fn_SISW_Prop_CommonModifiableProperties_Operation = False
										Exit Function
									End If
									Wait 1
								Next
								Fn_SISW_Prop_CommonModifiableProperties_Operation = True
					End Select
				Next
				Set objDialog = Nothing			 
		'--------------------------------------------------------------------------------------------------------------------------------------------------
	End Select
	Set objDiag1 = Nothing
	Set objDiag2 = Nothing
End Function

'######################################################################################################################################
'###
'###    FUNCTION NAME   :  Fn_SISW_Prop_PropertiesOnRelation_Operations(sAction, sObjProperty, sObjPropertyValue, sButtons)
'###
'###    DESCRIPTION     : To verifty and Set the Relation 
'###                                         
'###    Function Calls  :  
'###
'###	 HISTORY         :   		AUTHOR                 DATE        VERSION		BUILD
'###
'###    CREATED BY      :     Harshal		  		   08/09/10      1.0					818
'###
'###    REVIWED BY      :   Ketan			 	
'###
'###    MODIFIED BY     :	Ketan Raje 				14/09/2010							902
'###    MODIFIED BY     :	Sandeep N 				23/05/2013			Added Case : Incorporation Status				
'###    EXAMPLE         :   Call Fn_SISW_Prop_PropertiesOnRelation_Operations("Verify", "RelationType_StaticText", "TC_WorkContext_Relation", "Cancel")
'###								Case "Name_Link" = Msgbox Fn_SISW_Prop_PropertiesOnRelation_Operations("Verify", "Name_Link", "AutoTestDBA (autotestdba)", "Cancel")
'###								Case "Verify": Call Fn_SISW_Prop_PropertiesOnRelation_Operations("Verify", "Notes", "sagar:sagar1", "OK")
'###								Case "Set": Call Fn_SISW_Prop_PropertiesOnRelation_Operations("Set", "Notes", "ABC:XYZ", "OK")'
'###								bReturn=Fn_SISW_Prop_PropertiesOnRelation_Operations("Set", "Incorporation Status", "Cancelled", "OK")
'######################################################################################################################################
Public Function Fn_SISW_Prop_PropertiesOnRelation_Operations(sAction, sObjProperty, sObjPropertyValue, sButtons)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_Prop_PropertiesOnRelation_Operations"
	Dim objWrkCtxt, aButtons, iCounter, iCount, sAppValue, aObjPropertyValue, iRows, iFlag
	Dim objTable,objChild,bFlag
	
	If JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Properties").Exist(2) Then
		Set objWrkCtxt = Fn_UI_ObjectCreate("Fn_SISW_Prop_PropertiesOnRelation_Operations", JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Properties"))
		'Click on Static text
		If Fn_UI_ObjectExist("Fn_SISW_Prop_PropertiesOnRelation_Operations", objWrkCtxt.JavaStaticText("BottomLink")) = True Then
			objWrkCtxt.JavaStaticText("BottomLink").Click 1,1,"LEFT"
			Wait 1
		End If
	Else 
		Set objWrkCtxt=JavaWindow("DefaultWindow").JavaWindow("EnterthevaluesforProperties")
	End If
	
	Select Case sAction
		Case "Verify"
			'The Above part is to be coded as required.
			Select Case sObjProperty
				Case "RelationType_StaticText"
					objWrkCtxt.JavaStaticText("Property_label").setTOProperty "label", "Relation Type:"
					sAppValue  = objWrkCtxt.JavaStaticText("Property_label_value").GetROProperty("label")
					If trim(lcase(sAppValue)) = trim(lcase(sObjPropertyValue)) Then
						Fn_SISW_Prop_PropertiesOnRelation_Operations = True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS : Function [ Fn_SISW_Prop_PropertiesOnRelation_Operations ] executed successfully with case [ " & sObjProperty & " ] ")   	
					Else
						Fn_SISW_Prop_PropertiesOnRelation_Operations = False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to execute Function [ Fn_SISW_Prop_PropertiesOnRelation_Operations ] Invalid case [ " & sObjProperty & " ] ")
					End If
				Case "Name_Link"
					sAppValue  = objWrkCtxt.JavaStaticText("Name").GetROProperty("label")
					If trim(lcase(sAppValue)) = trim(lcase(sObjPropertyValue)) Then
						Fn_SISW_Prop_PropertiesOnRelation_Operations = True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS : Function [ Fn_SISW_Prop_PropertiesOnRelation_Operations ] executed successfully with case [ " & sObjProperty & " ] ")   	
					Else
						Fn_SISW_Prop_PropertiesOnRelation_Operations = False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to execute Function [ Fn_SISW_Prop_PropertiesOnRelation_Operations ] Invalid case [ " & sObjProperty & " ] ")
					End If
                 Case "RelationType_EditBox"
					   Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_SISW_Prop_PropertiesOnRelation_Operations",objWrkCtxt.JavaEdit("Name"),"attached text","Relation Type:")
					  sAppValue= Fn_Edit_Box_GetValue("Fn_SISW_Prop_PropertiesOnRelation_Operations",objWrkCtxt,"Name")
						If trim(lcase(sAppValue)) = trim(lcase(sObjPropertyValue)) Then
							Fn_SISW_Prop_PropertiesOnRelation_Operations = True
							Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS : Function [ Fn_SISW_Prop_PropertiesOnRelation_Operations ] executed successfully with case [ " & sObjProperty & " ] ")   	
						Else
							Fn_SISW_Prop_PropertiesOnRelation_Operations = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to execute Function [ Fn_SISW_Prop_PropertiesOnRelation_Operations ] Invalid case [ " & sObjProperty & " ] ")
						End If
				 Case "RelationTypeName_EditBox"
					   Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_SISW_Prop_PropertiesOnRelation_Operations",objWrkCtxt.JavaEdit("Name"),"attached text","Relation Type Name:")
					  sAppValue= Fn_Edit_Box_GetValue("Fn_SISW_Prop_PropertiesOnRelation_Operations",objWrkCtxt,"Name")
						If trim(lcase(sAppValue)) = trim(lcase(sObjPropertyValue)) Then
							Fn_SISW_Prop_PropertiesOnRelation_Operations = True
							Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS : Function [ Fn_SISW_Prop_PropertiesOnRelation_Operations ] executed successfully with case [ " & sObjProperty & " ] ")   	
						Else
							Fn_SISW_Prop_PropertiesOnRelation_Operations = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to execute Function [ Fn_SISW_Prop_PropertiesOnRelation_Operations ] Invalid case [ " & sObjProperty & " ] ")
						End If
				Case "Notes"
					iFlag = 0
					aObjPropertyValue = split(sObjPropertyValue, ":",-1, 1)
					iRows = objWrkCtxt.JavaList("NotesList").GetROProperty("items count")
					For iCount=0 to Ubound(aObjPropertyValue)
						For iCounter=0 to iRows-1
							If Trim(Lcase(aObjPropertyValue(iCount))) = Trim(Lcase(objWrkCtxt.JavaList("NotesList").GetItem(iCounter))) Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile") , aObjPropertyValue(iCount)&" found successfully in NotesList.")   	
								iFlag = iFlag + 1
								Exit For 
							End If
						Next
					Next
					If iFlag = Ubound(aObjPropertyValue)+1 Then
						Fn_SISW_Prop_PropertiesOnRelation_Operations = True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS : Function [ Fn_SISW_Prop_PropertiesOnRelation_Operations ] executed successfully with case [ " & sObjProperty & " ] ")   	
					Else
						Fn_SISW_Prop_PropertiesOnRelation_Operations = False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to execute Function [ Fn_SISW_Prop_PropertiesOnRelation_Operations ] Invalid case [ " & sObjProperty & " ] ")
					End If
				Case "Incorporation Status"
					   Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_SISW_Prop_PropertiesOnRelation_Operations",objWrkCtxt.JavaEdit("Name"),"attached text","Incorporation Status:")
					    sAppValue= Fn_Edit_Box_GetValue("Fn_SISW_Prop_PropertiesOnRelation_Operations",objWrkCtxt,"Name")
						If trim(lcase(sAppValue)) = trim(lcase(sObjPropertyValue)) Then
							Fn_SISW_Prop_PropertiesOnRelation_Operations = True
							Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS : Function [ Fn_SISW_Prop_PropertiesOnRelation_Operations ] executed successfully with case [ " & sObjProperty & " ] ")   	
						Else
							Fn_SISW_Prop_PropertiesOnRelation_Operations = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to execute Function [ Fn_SISW_Prop_PropertiesOnRelation_Operations ] Invalid case [ " & sObjProperty & " ] ")
						End If
				Case Else
					Fn_SISW_Prop_PropertiesOnRelation_Operations = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to execute Function [ Fn_SISW_Prop_PropertiesOnRelation_Operations ] Invalid case [ " & sObjProperty & " ] ")
			End Select
		Case "Set"
			Select Case sObjProperty
				Case "Notes"
					aObjPropertyValue = split(sObjPropertyValue, ":",-1, 1)
					'Set the Edit Notes Check box to "ON"
					Call Fn_CheckBox_Set("Fn_SISW_Prop_PropertiesOnRelation_Operations", objWrkCtxt, "EditNote", "ON")
					For iCount=0 to Ubound(aObjPropertyValue)
						'Set the value in "Notes" Edit Box.
						Call Fn_Edit_Box("Fn_SISW_Prop_PropertiesOnRelation_Operations",objWrkCtxt,"NoteText",aObjPropertyValue(iCount))
						'Click on Add Notes button
						Call Fn_Button_Click("Fn_SISW_Prop_PropertiesOnRelation_Operations", objWrkCtxt, "AddNote")
					Next
						Fn_SISW_Prop_PropertiesOnRelation_Operations = True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS : Function [ Fn_SISW_Prop_PropertiesOnRelation_Operations ] executed successfully with case [ " & sObjProperty & " ] ")
				Case "Incorporation Status"
					Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_SISW_Prop_PropertiesOnRelation_Operations",objWrkCtxt.JavaEdit("Name"),"attached text",sObjProperty&":")
					objWrkCtxt.JavaEdit("Name").Set ""
					wait 2
					objWrkCtxt.JavaStaticText("Property_label").SetTOProperty "label",sObjProperty+":"
					objWrkCtxt.JavaButton("LOVdropdown_16").Click micLeftBtn
					wait 2
					Set objTable=Description.Create()
					objTable("Class Name").value="JavaTable"
					'objTable("tagname").value="LOVTreeTable"
					objTable("toolkit class").value="com.teamcenter.rac.common.lov.view.components.LOVTreeTable"
					Set objChild=objWrkCtxt.ChildObjects(objTable)
					bFlag=False
					Wait 1
					For iCounter=0 To objChild(0).GetROProperty("rows")-1
						If trim(objChild(0).Object.getValueAt(iCounter,0).getDisplayableValue())=trim(sObjPropertyValue) Then
							objChild(0).ClickCell iCounter,0
							bFlag=True
							Exit For
						End If
					Next
					If bFlag=False Then
						Fn_SISW_Prop_PropertiesOnRelation_Operations=False
'						objWrkCtxt.JavaEdit("Name").Set ""
						Exit function
					Else
						Fn_SISW_Prop_PropertiesOnRelation_Operations=True
					End If
					
				'===========================================================================
				Case "SVTL_objective"	
					If sObjPropertyValue<>"" Then
						objWrkCtxt.JavaStaticText("SVTL_objective").SetTOProperty "label",sObjProperty+":"
						Call Fn_Edit_Box("Fn_SISW_Prop_PropertiesOnRelation_Operations",objWrkCtxt,"SVTL_Obj",sObjPropertyValue)
					End If
					If Fn_SISW_UI_Object_Operations("Fn_SISW_Prop_PropertiesOnRelation_Operations","Exist", objWrkCtxt.JavaButton("Apply All"), "")=True  Then
						Call Fn_Button_Click("Fn_SISW_Prop_PropertiesOnRelation_Operations", objWrkCtxt, "Apply All")
						Call Fn_ReadyStatusSync(1)
					End If
					If Fn_SISW_UI_Object_Operations("Fn_SISW_Prop_PropertiesOnRelation_Operations","Enabled", objWrkCtxt.JavaButton(sButtons), "")=True Then
						Fn_SISW_Prop_PropertiesOnRelation_Operations = True
					Else
						Fn_SISW_Prop_PropertiesOnRelation_Operations = False					
					End If	
					'===========================================================================
				Case Else
					Fn_SISW_Prop_PropertiesOnRelation_Operations = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to execute Function [ Fn_SISW_Prop_PropertiesOnRelation_Operations ] Invalid case [ " & sObjProperty & " ] ")
			End Select
		Case Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to execute Function [ Fn_SISW_Prop_PropertiesOnRelation_Operations ] Invalid case [ " & sAction & " ] ")
				Fn_SISW_Prop_PropertiesOnRelation_Operations = False
	End Select
	'Click on Buttons
	If sButtons<>"" Then
			aButtons = split(sButtons, ":",-1,1)
			iCounter = Ubound(aButtons)
			For iCount=0 to iCounter
				'Click on Add Button
				Call Fn_Button_Click("Fn_SISW_Prop_PropertiesOnRelation_Operations", objWrkCtxt, aButtons(iCount))
			Next
	End If
Set objWrkCtxt = Nothing
End Function

''*********************************************************		Function used to Perform Operations On Viewer Tab Controls**********************************
'Function Name	:			Fn_SISW_Prop_Text_PropertyVerify
'
'Description 		:			Checking the property from Textfield

'Parameters		:	 		sProperty, sValue

'Return Value	: 			True|False

'Pre-requisite	:		 	Property Window should Exist

'Examples		:			MsgBox Fn_SISW_Prop_Text_PropertyVerify("Contract Category","CONTRACT")

 
'History		:		
'							Developer Name				Reviewer Name					Date						Rev. No.						Changes Done						Reviewer
'						---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'							Harshal Agrawal			Harshal Agrawal				21/10/2010
'						----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_Prop_Text_PropertyVerify(sProperty,sValue)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_Prop_Text_PropertyVerify"
   Dim ObjPV,sAppValue,iCount,aPropValue,aValue
   Set ObjPV = JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Print")
   If NOT ObjPV.Exist(SISW_MIN_TIMEOUT) Then
		JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Properties").JavaButton("Print").Click micLeftBtn
   End If
    Wait(5)
	ObjPV.JavaCheckBox("Text").Set "ON"
	Wait(5)
	sAppValue = ObjPV.JavaEdit("TextPane").GetROProperty("value")
	Wait(5)
	aPropValue= Split(sAppValue,vblf)
	For iCount = 0 to ubound(aPropValue)
    	If aPropValue(iCount) <> "" Then
			aValue = Split(aPropValue(iCount),".")
			If lcase(trim(aValue(0))) = lcase(trim(sProperty)) Then
				If instr(1,(lcase(trim(aValue(1)))),(lcase(trim(sValue))))<>0 Then
					Fn_SISW_Prop_Text_PropertyVerify = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Fn_SISW_Prop_Text_PropertyVerify ")
					Exit For
				Else
					Fn_SISW_Prop_Text_PropertyVerify = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Fn_SISW_Prop_Text_PropertyVerify ")
				End If
			Else
				Fn_SISW_Prop_Text_PropertyVerify =False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Fn_SISW_Prop_Text_PropertyVerify ")
			End If
		End If
	Next
	ObjPV.JavaCheckBox("HTML").Set "ON"
	ObjPV.JavaButton("Close").Click micLeftBtn
	Set ObjPV = Nothing 
End Function

'*********************************************************		Edit / Verify object properties		***********************************************************************
'Function Name		:				Fn_SISW_Prop_ObjPropertiesOperation

'Description			 :		 		 This function edits or verifies object properties

'Return Value		   : 				PASS/ FAIL

'Pre-requisite			:		 		Object to edit properties is selected.

'Examples				:               Call Fn_SISW_Prop_ObjPropertiesOperation("Modify", "JavaList", "Excel Template Rules", "disable_outline~disable_outline") 
'												MsgBox  Fn_SISW_Prop_ObjPropertiesOperation("Verify","SetAndVerifyDateControltype", "02/20/12","February 20, 2012") 
'History:
'										Developer Name			Date			Rev. No.			Changes Done			Reviewer	
'----------------------------------------------------------------------------------------------------------------------------------------------
'										Ketan Raje			12-Apr-2011			1.0														Harshal					
'----------------------------------------------------------------------------------------------------------------------------------------------
'										Koustubh W			26-May-2011			1.0					Added code for case Modify				
'----------------------------------------------------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------------------------------------------------------
'										Avinash J.		10-Aug-2012		1.0					Added Cases 	[" SetAndVerifyDateControltype"] and [VerifyDateControltype]
'----------------------------------------------------------------------------------------------------------------------------------------------
 
Public Function Fn_SISW_Prop_ObjPropertiesOperation(sAction, sObjType, sObjName, sObjValue)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_Prop_ObjPropertiesOperation"
	On Error Resume Next
	Fn_SISW_Prop_ObjPropertiesOperation = False
	Dim objEditProperties, objProperties, objCheckIn, objCheckOutDialog, arrValues, iCounter
	Dim objAllOuterObjects, objChildObj,objChildButton, objChildEditbox, objCheckbox
	Dim objInnerObjects, bFlag,sDateText
	Dim iCnt


    Set objCheckOutDialog = Fn_SISW_GetChkInChkOutObject("CheckOut")
	If TypeName(objCheckOutDialog) <> "Nothing" Then
		Call Fn_Button_Click("Fn_SISW_Prop_ObjPropertiesOperation", objCheckOutDialog,"Yes")
	End If
	Set objCheckOutDialog = Nothing

    Call Fn_ReadyStatusSync(2)
    ' 	Object created for "Edit Properties" Dialog
    Set objEditProperties = Fn_SISW_GetObject("Edit Properties")
    ' 	Object created for "Properties" Dialog
    Set objProperties = Fn_SISW_GetObject("Properties")
    
     'Checks whether the "Edit Property" or "Properties" Dialog is displayed
   If objEditProperties.Exist(SISW_MIN_TIMEOUT) = False AND objProperties.Exist(SISW_MIN_TIMEOUT) = False Then
        Call Fn_MenuOperation("Select","Edit:Properties")
		Call Fn_ReadyStatusSync(1)
		If objProperties.Exist(5) Then
			Call Fn_Button_Click("",objProperties,"Check-Out and Edit")
		End If
		Set objCheckOutDialog = Fn_SISW_GetChkInChkOutObject("CheckOut")
		If TypeName(objCheckOutDialog) <> "Nothing" Then
				Call Fn_Button_Click("Fn_SISW_Prop_ObjPropertiesOperation", objCheckOutDialog,"Yes")
		End If
		Set objCheckOutDialog = Nothing
		  
	End If

'	Fn_UI_Object_SetTOProperty_ExistCheck("Fn_SISW_Prop_ObjPropertiesOperation",objEditProperties.JavaStaticText("BottomLink"),"label","All")
   objEditProperties.JavaStaticText("BottomLink").SetTOProperty "label","All"
   wait 1
   If objEditProperties.JavaStaticText("BottomLink").Exist(SISW_MIN_TIMEOUT) Then
	   Call Fn_UI_JavaStaticText_Click("Fn_SISW_Prop_ObjPropertiesOperation",objEditProperties,"BottomLink",1,1,"LEFT")
   End If
   Call Fn_ReadyStatusSync(1)
   'If Fn_Java_StaticText_Exist("Fn_SISW_Prop_ObjPropertiesOperation",objEditProperties, "More...") Then
   if Fn_SISW_UI_Object_Operations("Fn_SISW_Prop_ObjPropertiesOperation","Exist", objEditProperties.JavaStaticText("More..."), SISW_MICRO_TIMEOUT) then
		Call Fn_UI_JavaStaticText_Click("Fn_SISW_Prop_ObjPropertiesOperation",objEditProperties,"More...",1,1,  "LEFT")
   End If

   Select Case sAction
		Case "Verify","VerifyClose","VerifyWithoutClose"
			 Select Case sObjType
			 Case "JavaEdit"
					  If sObjName <> "" Then
							objEditProperties.JavaStaticText("ObjStaticText").SetTOProperty "attached text",sObjName & ":"
						   'objEditProperties.JavaEdit("ObjEditbox").SetTOProperty "attached text",sObjName & ":"
						   If Trim(Lcase(Fn_Edit_Box_GetValue("Fn_SISW_Prop_ObjPropertiesOperation",objEditProperties,"ObjEditbox"))) = Trim(Lcase(sObjValue)) Then
								Fn_SISW_Prop_ObjPropertiesOperation = True
						   End If
					  End If
			 Case "JavaStaticText"
						   objEditProperties.JavaStaticText("ObjStaticText").SetTOProperty "attached text",sObjValue
						   If objEditProperties.JavaStaticText("ObjStaticText").Exist(1) Then
								Fn_SISW_Prop_ObjPropertiesOperation = True
						   End If
			 Case "JavaList"
					  If sObjName <> "" Then
							arrValues = split(sObjValue, "~")
							objEditProperties.JavaList("ReleaseStatus").SetTOProperty "attached text", sObjName & ":"
							for iCounter = 0 to uBound(arrValues)
								Fn_SISW_Prop_ObjPropertiesOperation = Fn_UI_ListItemExist("Fn_SISW_Prop_ObjPropertiesOperation", objEditProperties, "ReleaseStatus",arrValues(iCounter))
								if Fn_SISW_Prop_ObjPropertiesOperation = False then exit for
							next
					  End If
		 Case  "VerifyDateControltype"                    ''Added by Avinash J. 10-Aug-2012
         			 If sObjValue <> "" Then
							 sDateText= objEditProperties.JavaCheckBox("Date").GetROProperty("attached text")
							sDateText=Split(sDateText," ")
                		   If Trim(Lcase(sDateText(0)) )= Trim(Lcase(sObjValue)) Then
									Fn_SISW_Prop_ObjPropertiesOperation = True
							 End If
					  End If

			Case  "SetAndVerifyDateControltype"       ''Added by Avinash J. 10-Aug-2012
        		Call Fn_CheckBox_Set("Fn_BOMViewRev_SaveAs" ,objEditProperties,"Date", "ON") 
					   Wait 2
                    	Call  Fn_Edit_Box("Fn_SISW_Prop_ObjPropertiesOperation",objEditProperties,"Date",sObjName)
								Wait 2
								objEditProperties.Type micTab
								Wait 2
						If objEditProperties.JavaObject("CalendarPanel").Exist (SISW_MIN_TIMEOUT)= True   Then   ''Check the Existance of CalenderPanel
							
							 If Trim(Lcase(Fn_Edit_Box_GetValue("Fn_SISW_Prop_ObjPropertiesOperation",objEditProperties,"Date"))) = Trim(LCase(sObjValue)) Then
	   									Wait 3
										objEditProperties.JavaButton("OK").Click 
										Fn_SISW_Prop_ObjPropertiesOperation=True	
							 End If
					End If	

    End Select
		Case "Modify"
			'Is to be coded as required.
			Select Case sObjType
				Case "JavaList"
					If sObjName <> "" Then
						'read all objects under Edit Properties
						Set objChildObj =Description.Create
						objChildObj("toolkit class").value = "com.teamcenter.rac.stylesheet.PropertyArray"
						objChildObj("Class Name").value ="JavaObject"
						Set objAllOuterObjects = objEditProperties.ChildObjects(objChildObj)
						'read each JavaObject
						For iCounter = 0 to cInt(objAllOuterObjects.count) -1
							Set objChildObj =Description.Create
							objChildObj("Class Name").value ="JavaObject"
							objChildObj("attached text").value= sObjName & ":"
							'select object which has child JPanel object having attached text Excel template Rules:
							Set objInnerObjects = objAllOuterObjects(iCounter).ChildObjects(objChildObj)
							If objInnerObjects.count > 0 Then
								bFlag = True
								Exit for
							End If
						Next

						' set checkbox on
						If iCounter > 0 and iCounter < cInt(objAllOuterObjects.count) Then
							Set objChildObj =Description.Create
							objChildObj("Class Name").value ="JavaCheckBox"
							objChildObj("class description").value ="check_button"
							objChildObj("attached text").value="edit_16"
							Set objCheckbox = objAllOuterObjects(iCounter).ChildObjects(objChildObj)
							If objCheckbox.count > 0  Then
								objCheckbox(0).set "ON"
								objCheckbox(0).Type "ON"
								wait(3)
								'taking editbox object
								Set objChildObj = Description.Create
								objChildObj("Class Name").value ="JavaEdit"
								'objChildObj("tagname").value="iComboBox\$9"
								objChildObj("toolkit class").value="com\.teamcenter\.rac\.common\.lov\.view\.components\.LOVSelectionDisplayView\$1"
								Set objChildEditbox = objAllOuterObjects(iCounter).ChildObjects(objChildObj)
								If objChildEditbox.count > 0 Then
									'taking button object
									Set objChildObj = Description.Create
									objChildObj("Class Name").value ="JavaButton"
									objChildObj("attached text").value="add_16"
							
									Set objChildButton = objAllOuterObjects(iCounter).ChildObjects(objChildObj)
							
									arrValues = split(sObjValue, "~")
									For iCounter = 0 to Ubound(arrValues)
'										objChildEditbox(0).Type arrValues(iCounter)
										'objChildEditbox(0).Object.SetText arrValues(iCounter)
										objChildEditbox(0).Type arrValues(iCounter)                        ''Added by Avinash J. on 20130130 build        -13-feb-13
										wait 2
										objChildEditbox(0).Type micReturn  
                                        wait(3)
										objChildButton(0).click 
										wait 1
									Next
								Else
									bFlag = False
								End If
								objCheckbox(0).set "OFF"
							Else
								bFlag = False
							End If
						Else
							bFlag = False
						End If
					Else
						bFlag = False
					End If
					Fn_SISW_Prop_ObjPropertiesOperation = bFlag
			End Select
End Select 

If  sAction <> "VerifyClose" And sAction<>"VerifyWithoutClose"Then 'Added by Nilesh on 14-Sep-2012
		Call Fn_Button_Click("Fn_SISW_Prop_ObjPropertiesOperation",objEditProperties,"SaveAndCheck-In")
		Call Fn_ReadyStatusSync(1)
		Set objCheckIn = Fn_SISW_GetChkInChkOutObject("CheckIn")
		 If TypeName(objCheckIn) <> "Nothing" Then
			Call Fn_Button_Click("Fn_SISW_Prop_ObjPropertiesOperation", objCheckIn,"Yes")
		 End If
		 Set objCheckIn = Nothing
Else
		If sAction<>"VerifyWithoutClose" Then
			Call Fn_Button_Click("Fn_SISW_Prop_ObjPropertiesOperation",objEditProperties,"Close")
		End If      	
End If

   'Write Log
   Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction&" successfully executed of function Fn_SISW_Prop_ObjPropertiesOperation.")
   ' Clear out memory allocated for objects
   Set objEditProperties = nothing
   Set objCheckOutDialog  = nothing
   Set objCheckIn = nothing
   Set objAllOuterObjects = nothing
   Set objChildObj = nothing
   Set objChildButton = nothing
   Set objChildEditbox = nothing
   Set objCheckbox = nothing
   Set objInnerObjects = nothing
End Function

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Function Name	:	Fn_SISW_VerifyProperties

'	Description		:	Function Used to verify object properties

'	Parameters		:	1. StrAction		: Action Name
'						2. StrLink			: Property link [ All | General ]
'						3. dicProperties	: Properties information
'						4. StrButtonName	: Button Name
'
'	Return Value	:	True / False

'	Pre-requisite	:	Properties dialog should be appear

'	Examples		:   Dim dicProperties
'						Set dicProperties = CreateObject( "Scripting.Dictionary" )
'										
'						dicProperties("PropertyName")="Variant Option Families"
'						dicProperties("Value")="Family"
'						bReturn= Fn_SISW_VerifyProperties("ListBox","All",dicProperties,"Cancel")
'										
'						dicProperties("PropertyName")="Type"
'						dicProperties("Value")="Option Family Group"
'						bReturn= Fn_SISW_VerifyProperties("EditBox","All",dicProperties,"Cancel")
'										
'						dicProperties("PropertyName")="User Can Unmanage"
'						dicProperties("Value")="False"
'						bReturn=Fn_SISW_VerifyProperties("RadioButton","All",dicProperties,"Cancel")
'									
'						dicProperties("PropertyName")="Type"
'						dicProperties("PropertyState")="enabled"
'						bReturn=Fn_SISW_VerifyProperties("EditBox_GetPropertyState","All",dicProperties,"Cancel")
'
'						dicProperties("PropertyName")="Owner~Group ID"
'						dicProperties("Value")="AutoTestDBA (autotestdba)~dba"
'						bReturn=Fn_SISW_VerifyProperties("Link","",dicProperties,"")
'										
'						dicProperties("PropertyName")="Last Modified Date"
'						dicProperties("PropertyState")="attached text"
'						bReturn=Fn_SISW_VerifyProperties("CheckBox_GetPropertyState","",dicProperties,"")
'
'						dicProperties("PropertyName")="Date Created"
'						dicProperties("PropertyState")="text"
'						bReturn=Fn_SISW_Prop_VerifyProperties("DateTime_GetPropertyState","",dicProperties,"Cancel")
'
'	History			:
'
'	Developer Name				Date	  	  Rev. No.						Changes Done																									Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'	Sandeep N				15-May-2013			1.0								Created																							  			Sunny R
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'	Sandeep N				15-May-2013			1.1				Added case : Link,CheckBox_GetPropertyState																					Anjali M
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'	Sandeep N				15-May-2013			1.1				Added case : VerifyPropertyLabels														 									Veena G
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'	Pranav Ingle			04-Dec-2013			1.2				Modified Function To handle Properties dialog under JavaApplet	
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'	Pranav Ingle			04-Dec-2013			1.3				Added Cases "ModifyEditBox" & "ModifyDateCheckBox"
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'	Vivek Ahirrao			06-July-2015						Added Case "ModifyListValue" to set Finifh Time value in Properties dialog, as per design change
'	Vivek Ahirrao			22-Jan-2016			1.4				Added Case "ListBox_PopupMenuSelect" & "RelationType_StaticText" 										[TC1122-20151116d-22_01_2016-VivekA-NewDevelopment]
'	shweta rathod			04-Mar-2016			1.4				Added code to click on General -> showempty 															[TC1122:2016021000:10Mar2016:ShwetaR:NewDevelopment]
'	shweta rathod			09-Mar-2016         1.0             Added case "RadioButtonMakeBuy"																			[TC1122:2016021000:10Mar2016:ShwetaR:NewDevelopment]
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'	Ankit Nigam				26-May-2016		 	1.0				Added Case "DateTime_GetPropertyState" to get Date Time text value 											[TC1122-2016050400-26_05_2016-AnkitN-NewDevelopment]	
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Function Fn_SISW_Prop_VerifyProperties(StrAction,StrLink,dicProperties,StrButtonName)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_Prop_VerifyProperties"
 	'Declaring variables
	Dim objPropetiesDialog
	Dim aValues,iCounter,bFlag,iCount,aProperty,iEleCount
    Dim sAppValue
    
	Fn_SISW_Prop_VerifyProperties=False
	'Setting window title
	If dicProperties("PropertyDialogTitle")<>"" Then
		 JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Properties").SetTOProperty "title",dicProperties("PropertyDialogTitle")
		 JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Properties").SetTOProperty "title",dicProperties("PropertyDialogTitle")
		 wait 1
	End If

 	'Checking existance of [ Properties ] dialog
	If not(JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Properties").Exist(SISW_MIN_TIMEOUT)) Then
		If not(JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Properties").Exist(SISW_MIN_TIMEOUT)) Then
			Call Fn_MenuOperation("Select","View:Properties")
		End If
	End If
	Call Fn_ReadyStatusSync(1)
	'Creating object of [ Properties ] dialog
	If JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Properties").Exist(SISW_MIN_TIMEOUT) Then
		Set objPropetiesDialog=JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Properties")
	ElseIf JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Properties").Exist(SISW_MIN_TIMEOUT) Then
		Set objPropetiesDialog=JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Properties")
	End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Clicking on page link
	If StrLink="" Then
		objPropetiesDialog.JavaStaticText("BottomLink").SetTOProperty "label","General"
		objPropetiesDialog.JavaStaticText("BottomLink").Click 1,1,"LEFT"
		wait 1
	ElseIf StrLink="GeneralShowEmpty" Then
		objPropetiesDialog.JavaStaticText("BottomLink").SetTOProperty "label","General"
		objPropetiesDialog.JavaStaticText("BottomLink").Click 1,1,"LEFT"
		wait 1
		bFlag = Fn_UI_Object_SetTOProperty_ExistCheck("Fn_SISW_Prop_VerifyProperties",objPropetiesDialog.JavaStaticText("BottomLink"),"label","Show empty properties...")
		If bFlag = True Then
			objPropetiesDialog.JavaStaticText("BottomLink").Click 1,1,"LEFT"
		End If
	Else
		objPropetiesDialog.JavaStaticText("BottomLink").SetTOProperty "label",StrLink
		objPropetiesDialog.JavaStaticText("BottomLink").Click 1,1,"LEFT"
		If StrLink="All" Then
			'to click on [ Show empty properties... ] link
			wait 1
			objPropetiesDialog.JavaStaticText("BottomLink").SetTOProperty "label","Show empty properties..."
			If objPropetiesDialog.JavaStaticText("BottomLink").Exist(SISW_MICRO_TIMEOUT) Then
				objPropetiesDialog.JavaStaticText("BottomLink").Click 1,1,"LEFT"
			End If
		End If
	End If

	Select Case StrAction
    	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to verify values from list Box
		'to verify values of List box list box should be enabled so if want to verify values from List box object should be check out
		Case "ListBox"
				objPropetiesDialog.JavaStaticText("Property_label").SetTOProperty "label",dicProperties("PropertyName")+":"
				If objPropetiesDialog.JavaList("Property_field").Exist(SISW_MICRO_TIMEOUT) Then
					aValues=Split(dicProperties("Value"),"~")
					For iCounter=0 to uBound(aValues)
						bFlag=false
						'Verifying value exist in list or not
						'taking item count from list
						iEleCount=Fn_UI_Object_GetROProperty("Fn_SISW_Prop_VerifyProperties",objPropetiesDialog.JavaList("Property_field"), "items count")
						For iCount=0 to iEleCount-1
							If objPropetiesDialog.JavaList("Property_field").GetItem(iCount)=aValues(iCounter) Then
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
						Fn_SISW_Prop_VerifyProperties=true
					End If
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"List box [ " & dicProperties("PropertyName") & " ] is not exist on dialog")
				End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to verify property value from edit boxes
		Case "EditBox","EditBoxExt"
				aProperty=Split(dicProperties("PropertyName"),"~")
				aValues=Split(dicProperties("Value"),"~")
				For iCounter=0 to UBound(aProperty)
					bFlag=False
					objPropetiesDialog.JavaStaticText("Property_label").SetTOProperty "label",aProperty(iCounter)+":"
					If objPropetiesDialog.JavaEdit("Property_field").Exist(SISW_MIN_TIMEOUT) Then
						If lcase(aValues(iCounter)) = "blank" Then
							If Fn_Edit_Box_GetValue("Fn_SISW_Prop_VerifyProperties",objPropetiesDialog,"Property_field")="" Then
								bFlag=True
							End If
						Else
							If True Then
								If StrAction ="EditBoxExt" Then									
									If Instr(1,Fn_Edit_Box_GetValue("Fn_SISW_Prop_VerifyProperties",objPropetiesDialog,"Property_field"),aValues(iCounter))>0 Then
										bFlag=True
									End If									
								Else
									If Fn_Edit_Box_GetValue("Fn_SISW_Prop_VerifyProperties",objPropetiesDialog,"Property_field")=aValues(iCounter) Then
										bFlag=True
									End If
								End If							
							End If							
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
					Fn_SISW_Prop_VerifyProperties=true
				End If
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to modify property value from edit boxes
		Case "ModifyEditBox"
				aProperty=Split(dicProperties("PropertyName"),"~")
				aValues=Split(dicProperties("Value"),"~")
				For iCounter=0 to ubound(aProperty)
					objPropetiesDialog.JavaStaticText("Property_label").SetTOProperty "label",aProperty(iCounter)+":"
					If objPropetiesDialog.JavaEdit("Property_field").Exist(SISW_MICRO_TIMEOUT) Then
						Call Fn_Edit_Box("Fn_SISW_Prop_VerifyProperties",objPropetiesDialog,"Property_field",aValues(iCounter))
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Property [ " & aProperty(iCounter) & " ] is not exist on dialog")
						Set objPropetiesDialog=Nothing
						Exit function
					End If
				Next
				Fn_SISW_Prop_VerifyProperties=True		
		'Case to modify property value (Finish time) from List boxes			
		Case "ModifyListValue"
				aProperty=Split(dicProperties("PropertyName"),"~")
				aValues=Split(dicProperties("Value"),"~")
				For iCounter=0 to ubound(aProperty)
					objPropetiesDialog.JavaStaticText("Property_label").SetTOProperty "label",dicProperties("PropertyName")+":"
					If objPropetiesDialog.JavaEdit("FinishDate").Exist(SISW_MICRO_TIMEOUT) Then
						'JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Properties").JavaList("Property_field").
						objPropetiesDialog.JavaEdit("FinishDate").Set ""
						objPropetiesDialog.JavaEdit("FinishDate").Type aValues(iCounter)
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Property [ " & aProperty(iCounter) & " ] is not exist on dialog")
						Set objPropetiesDialog=Nothing
						Exit function
					End If
				Next
				Fn_SISW_Prop_VerifyProperties=True	
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 		
		'Case to set date
		Case "ModifyDateCheckBox"
			aProperty=Split(dicProperties("PropertyName"),"~")
			aValues=Split(dicProperties("Value"),"~")
			For iCounter=0 to ubound(aProperty)
				objPropetiesDialog.JavaStaticText("DateReceived").SetTOProperty "label",aProperty(iCounter)+":"
				If objPropetiesDialog.JavaCheckBox("Date").Exist(SISW_MIN_TIMEOUT) Then
					objPropetiesDialog.JavaCheckBox("Date").Object.setDate(aValues(iCounter))
				else
					Set objPropetiesDialog=Nothing
					Exit function
				End If
			Next
			Fn_SISW_Prop_VerifyProperties=True
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to verify Radio Buttons
		Case "RadioButton"
				aProperty=Split(dicProperties("PropertyName"),"~")
				aValues=Split(dicProperties("Value"),"~")
				For iCounter=0 to UBound(aProperty)
					bFlag=False
					objPropetiesDialog.JavaStaticText("Property_label").SetTOProperty "label",aProperty(iCounter)+":"
					objPropetiesDialog.JavaRadioButton("Property_True").SetTOProperty "attached text",aValues(iCounter)

					If objPropetiesDialog.JavaRadioButton("Property_True").Exist(SISW_MIN_TIMEOUT) Then
						If Fn_UI_Object_GetROProperty("Fn_SISW_Prop_VerifyProperties",objPropetiesDialog.JavaRadioButton("Property_True"), "value")=1 Then
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
					Fn_SISW_Prop_VerifyProperties=true
				End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "RadioButtonMakeBuy"					'[TC1122:2016021000:10Mar2016:ShwetambriR:NewDevelopment] - Added to Verify Make/Buy radio button
				aProperty=Split(dicProperties("PropertyName"),"~")
				aValues=Split(dicProperties("Value"),"~")
				For iCounter=0 to UBound(aProperty)
					bFlag=False
					objPropetiesDialog.JavaRadioButton("Make/Buy").SetTOProperty "attached text",aValues(iCounter)
					If objPropetiesDialog.JavaRadioButton("Make/Buy").Exist(SISW_MIN_TIMEOUT) Then
						If Fn_UI_Object_GetROProperty("Fn_SISW_Prop_VerifyProperties",objPropetiesDialog.JavaRadioButton("Make/Buy"), "value")=1 Then
							bFlag=True
						End If
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Property [ " & aProperty(iCounter) & " ] is not exist on dialog")
						Exit for
					End If
					If bFlag=False Then
						Exit for
					End If
				Next
				If bFlag=True Then
					Fn_SISW_Prop_VerifyProperties=true
				End If
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get specific property state of specific Edit box : e.g { current value, editable state, enabled state }
		Case "EditBox_GetPropertyState"
				objPropetiesDialog.JavaStaticText("Property_label").SetTOProperty "label",dicProperties("PropertyName")+":"
				If objPropetiesDialog.JavaEdit("Property_field").Exist(SISW_MIN_TIMEOUT) Then
					Fn_SISW_Prop_VerifyProperties=objPropetiesDialog.JavaEdit("Property_field").GetROProperty(dicProperties("PropertyState"))
				else
					Fn_SISW_Prop_VerifyProperties=false
				End if
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get specific property state of specific Date Time : e.g { Date Created:, Last Modified Date: }
		'Added Case "DateTime_GetPropertyState" to get Date Time value 			-[TC1122-2016050400-26_05_2016-AnkitN-NewDevelopment]	
		Case "DateTime_GetPropertyState"
				objPropetiesDialog.JavaStaticText("Property_label").SetTOProperty "label", dicProperties("PropertyName") + ":"
				If objPropetiesDialog.JavaEdit("Property_field").Exist(SISW_MIN_TIMEOUT) AND objPropetiesDialog.JavaList("Property_field").Exist(SISW_MIN_TIMEOUT) Then
					Fn_SISW_Prop_VerifyProperties = objPropetiesDialog.JavaEdit("Property_field").GetROProperty(dicProperties("PropertyState")) & " " & objPropetiesDialog.JavaList("Property_field").GetROProperty(dicProperties("PropertyState"))
				Else
					Fn_SISW_Prop_VerifyProperties=False
				End if
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "ListBox_CheckProperty"
			objPropetiesDialog.JavaStaticText("Property_label").SetTOProperty "label",dicProperties("PropertyName")+":"
			If objPropetiesDialog.JavaList("Property_field").Exist(SISW_MIN_TIMEOUT) Then
				Fn_SISW_Prop_VerifyProperties = objPropetiesDialog.JavaList("Property_field").CheckProperty(dicProperties("CheckPropertyName"),dicProperties("Value"))
			End IF
            '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to get specific property state of specific Check Box : e.g { current value, attached text }
		Case "CheckBox_GetPropertyState"
				objPropetiesDialog.JavaStaticText("Property_label").SetTOProperty "label",dicProperties("PropertyName")+":"
				If objPropetiesDialog.JavaCheckBox("Property_field").Exist(SISW_MIN_TIMEOUT) Then
					Fn_SISW_Prop_VerifyProperties=objPropetiesDialog.JavaCheckBox("Property_field").GetROProperty(dicProperties("PropertyState"))
				else
					Fn_SISW_Prop_VerifyProperties=false
				End if
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to verify lproperty links
		Case "Link"
			aProperty=Split(dicProperties("PropertyName"),"~")
				aValues=Split(dicProperties("Value"),"~")
				For iCounter=0 to UBound(aProperty)
					bFlag=False
					objPropetiesDialog.JavaStaticText("Property_label").SetTOProperty "label",aProperty(iCounter)+":"

					If objPropetiesDialog.JavaStaticText("Property_label_value").Exist(SISW_MIN_TIMEOUT) Then
						If aValues(iCounter) = "blank" Then
							If Fn_UI_Object_GetROProperty("Fn_SISW_Prop_VerifyProperties",objPropetiesDialog.JavaStaticText("Property_label_value"), "label")="" Then
								bFlag=True
							End If
						Else
							If Fn_UI_Object_GetROProperty("Fn_SISW_Prop_VerifyProperties",objPropetiesDialog.JavaStaticText("Property_label_value"), "label")=aValues(iCounter) Then
								bFlag=True
							End If
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
					Fn_SISW_Prop_VerifyProperties=true
				End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "VerifyPropertyLabels"
			aProperty=Split(dicProperties("PropertyName"),"~")
			For iCounter=0 to ubound(aProperty)
				bFlag=False
				objPropetiesDialog.JavaStaticText("Property_label").SetTOProperty "label",aProperty(iCounter)+":"
				If objPropetiesDialog.JavaStaticText("Property_label").Exist(SISW_MICRO_TIMEOUT) Then
					bFlag=True
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Property [ " & aProperty(iCounter) & " ] not found")
					Exit for
				End If
			Next
			If bFlag=True Then
				Fn_SISW_Prop_VerifyProperties=true
			End If
		'[TC1122-20151116d-22_01_2016-VivekA-NewDevelopment]
		Case "ListBox_PopupMenuSelect"
			objPropetiesDialog.JavaStaticText("Property_label").SetTOProperty "label",dicProperties("PropertyName")+":"
			If objPropetiesDialog.JavaList("Property_field").Exist(SISW_MICRO_TIMEOUT) Then
				objPropetiesDialog.JavaList("Property_field").Select "0"
				Wait 1
				objPropetiesDialog.JavaList("Property_field").Click 5,5,"RIGHT"
				Wait 1
				Select Case dicProperties("PopupMenuName")
					Case "Properties..."
						bFlag = Fn_KeyBoardOperation("SendKeys","(P)")
						If bFlag = True Then
							Fn_SISW_Prop_VerifyProperties = True
						End If
						Wait 5
					Case "Open"
						'Future Use
					Case "Copy"
						'Future Use
				End Select
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"List box [ " & dicProperties("PropertyName") & " ] is not exist on dialog")
			End If
		'[TC1122-20151116d-22_01_2016-VivekA-NewDevelopment]
		Case "RelationType_StaticText"
			objPropetiesDialog.JavaStaticText("Property_label").setTOProperty "label", "Relation Type:"
			'sAppValue = objPropetiesDialog.JavaStaticText("Property_label_value").GetROProperty("label")
			
			sAppValue = Fn_UI_Object_GetROProperty("Fn_SISW_Prop_VerifyProperties",objPropetiesDialog.JavaStaticText("Property_label_value"), "label")
			If Trim(LCase(sAppValue)) = Trim(LCase(dicProperties("Value"))) Then
				Fn_SISW_Prop_VerifyProperties = True
				Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS : Function [ Fn_SISW_Prop_VerifyProperties ] executed successfully with case [ " & dicProperties("Value") & " ] ")   	
			Else
				Fn_SISW_Prop_VerifyProperties = False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to execute Function [ Fn_SISW_Prop_VerifyProperties ] Invalid case [ " & dicProperties("Value") & " ] ")
			End If
	   '--------------------------------------------------------------------------------------------------
	   '[TC11.3-20170509d-02_June_2017-PoonamC-NewDevelopment]
		Case "VerifyListBoxContent"
				objPropetiesDialog.JavaStaticText("Property_label").SetTOProperty "label",dicProperties("PropertyName")+":"
				If objPropetiesDialog.JavaList("Property_field").Exist(SISW_MICRO_TIMEOUT) Then
					aValues=Split(dicProperties("Value"),"~")
					aProperty = Fn_UI_Object_GetROProperty("Fn_SISW_Prop_VerifyProperties",objPropetiesDialog.JavaList("Property_field"), "list_content")
					aProperty = Split(aProperty,"")
					For iCounter=0 to uBound(aValues)
						bFlag=false
						For iCount=0 to UBound(aProperty)
							If aProperty(iCount)=aValues(iCounter) Then
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
						Fn_SISW_Prop_VerifyProperties=true
					End If
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"List box [ " & dicProperties("PropertyName") & " ] is not exist on dialog")
				End If			
	End	Select
	'Clicking on button
	If StrButtonName<>"" Then
		Call Fn_Button_Click("Fn_SISW_Prop_VerifyProperties", objPropetiesDialog,StrButtonName)
	End If
	'Releasing object of [ Properties ] dialog
	Set objPropetiesDialog=nothing
End Function

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_SISW_Prop_EditProperties

'Description			 :	Function Used to verify object properties

'Parameters			   :   1.StrAction: Action Name
'										2.StrLink: Property link [ All | General ]
'										3.dicEditProperties: Properties information
'										4.StrButtonName: Button Name
'
'Return Value		   : 	True / False
'
'Pre-requisite			:	Object should be selected
'
'Examples				:   Dim dicEditProperties
'										Set dicEditProperties=CreateObject("Scripting.Dictionary")
'
'										dicEditProperties("PropertyName")="Finish Items"
'										bReturn=Fn_SISW_Prop_EditProperties("AddClipboardObjects","",dicEditProperties,"")
'										
'										dicEditProperties("PropertyName")="Description~Name"
'										dicEditProperties("Value")="ModifiedDescription~FinishGroup02_2"
'										bReturn=Fn_SISW_Prop_EditProperties("EditBox","",dicEditProperties,"SaveAndCheck-In")
'
'										dicEditProperties("PropertyName")="Custom1"
'										dicEditProperties("PrimaryColumnName")="ID"
'										dicEditProperties("ColumnName")="Name"
'										dicEditProperties("PrimaryColumnValue")="000037~000038"
'										dicEditProperties("Value")="TestItem3~TestItem4"
'										bReturn= Fn_SISW_Prop_EditProperties("LOVTableCellExist","All",dicEditProperties,"")
'										
'										dicEditProperties("PropertyName")="Custom1"
'										dicEditProperties("PrimaryColumnName")="ID"
'										dicEditProperties("PrimaryColumnValue")="000037"
'										bReturn= Fn_SISW_Prop_EditProperties("LOVTableSelect","All",dicEditProperties,"")
'										
'										dicEditProperties("PropertyName")="Test3"
'										dicEditProperties("Value")="True"
'										bReturn= Fn_SISW_Prop_EditProperties("RadioButton","All",dicEditProperties,"")
'										
'										dicEditProperties("PropertyName")="Test2"
'										dicEditProperties("Value")="04-Sep-2012 08:10"
'										bReturn= Fn_SISW_Prop_EditProperties("DateCheckBox","All",dicEditProperties,"")
'
'										dicEditProperties("PropertyName")="strArray5"
'										dicEditProperties("Value")="Test1~Test2~Test3"
'										bReturn=Fn_SISW_Prop_EditProperties("EditBox_AddToList","All",dicEditProperties,"")
'
'										dicEditProperties("PropertyName")="Excel Template Rules"
'										dicEditProperties("Value")="apply_packing"
'										bReturn=Fn_SISW_Prop_EditProperties("LOVTable_AddToList","All",dicEditProperties,"SaveAndCheck-In")
'
'History					 :			
'		Developer Name		Date		Rev. No.	Changes Done																	Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'		Sandeep N		07-June-2013	1.0																							Veena G
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'		Sandeep N		08-Jul-2013		1.1			Added Case : LOVTableCellExist,LOVTableSelect,RadioButton,DateCheckBox			Preeti S
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'		Sandeep N		10-Jul-2013		1.2			Added Case : EditBox_AddToList													Sonal P
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'		Sonal P			13-Jul-2013		1.3			Added Case : LOVTable_AddToList													Sonal P
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'		Vivek A			16-Dec-2015		1.4			Added Case : EditBox_ModifyandAddToList											[TC1122-20151116b-16_12_2015-VivekA-NewDevelopment]
'		Poonam C		10-Feb-2016		1.5			Added Cases : SortContentList, VerifySortedContentList							[TC1122-20160113-10_02_2016-VivekA-NewDevelopment]
'       Shweta Rathod   09-Mar-2016		1.0			Added Case : Case "RadioButtonMakeBuy"											[TC1122:2016021000:10Mar2016:ShwetaR:NewDevelopment]
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_SISW_Prop_EditProperties(StrAction,StrLink,dicEditProperties,StrButtonName)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_Prop_EditProperties"
	'Declaring variables
	Dim aProperty,aValues,iCounter,bFlag,aPrimaryValues,iCount,aCase
	Dim objPropetiesDialog,objCheckOut, iCnt
    Dim StrTitle, arrValue, iEleCount, sAppValues, aTime, sDate, sTime
	'StrTitle=JavaWindow("DefaultWindow").GetROProperty("title")
	
	if dicEditProperties("PropertyName")<>"" and dicEditProperties("PropertyName")="Gov Classification" then
		dicEditProperties("PropertyName")="Government Classification"
	End if

   Fn_SISW_Prop_EditProperties=False
   'Checking existance of [ Properties ] dialog
'	If Instr(1,StrTitle,"My Teamcenter") or Instr(1,StrTitle,"Change Manager") Then
'		Set objCheckOut=JavaWindow("DefaultWindow").JavaWindow("Check-Out")
'	Else
'		Set objCheckOut=JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Check-Out")
'	End if
	If not JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Edit Properties").Exist(6) And Not JavaWindow("DefaultWindow").JavaWindow("DefaultEmbeddedFrame").JavaDialog("Edit Properties").Exist(1) Then
		'checking existance of Check Out dialog
			Set objCheckOut= Fn_SISW_GetChkInChkOutObject("CheckOut")
			If typename(objCheckOut) = "Nothing" Then
				'calling menu [ Edit:Properties ]
				Call Fn_MenuOperation("Select","Edit:Properties")
				Call Fn_ReadyStatusSync(1)
			End If
			Set objCheckOut= Nothing
			'checking existance of Check Out dialog
'		else
'			'ckicking on Yes button to checkout object
'			Call Fn_Button_Click("Fn_SISW_EditProperties",objCheckOut,"Yes")
	End If
	   Set objCheckOut= Fn_SISW_GetChkInChkOutObject("CheckOut")
		If typename(objCheckOut) <> "Nothing" Then
			'ckicking on Yes button to checkout object
			Call Fn_Button_Click("Fn_SISW_Prop_EditProperties",objCheckOut,"Yes")
		End If
		Set objCheckOut= Nothing
	Call Fn_ReadyStatusSync(1)
	'Creating object of [ EditProperties ] dialog
	If JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Edit Properties").Exist(1) Then
		Set objPropetiesDialog=JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Edit Properties")
	ElseIf JavaWindow("DefaultWindow").JavaWindow("DefaultEmbeddedFrame").JavaDialog("Edit Properties").Exist(1) Then
		Set objPropetiesDialog=JavaWindow("DefaultWindow").JavaWindow("DefaultEmbeddedFrame").JavaDialog("Edit Properties")
	End If

	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Clicking on link
	If StrLink="" Then
		objPropetiesDialog.JavaStaticText("BottomLink").SetTOProperty "label","General"
		objPropetiesDialog.JavaStaticText("BottomLink").Click 1,1,"LEFT"
	Else
		objPropetiesDialog.JavaStaticText("BottomLink").SetTOProperty "label",StrLink
		objPropetiesDialog.JavaStaticText("BottomLink").Click 1,1,"LEFT"
	End if	
	wait 1
	objPropetiesDialog.JavaStaticText("BottomLink").SetTOProperty "label","Show empty properties..."
	If objPropetiesDialog.JavaStaticText("BottomLink").Exist(SISW_MICRO_TIMEOUT) Then
		objPropetiesDialog.JavaStaticText("BottomLink").Click 1,1,"LEFT"
	End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Select Case StrAction
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to modify property value from edit boxes
		Case "EditBox", "EditBoxDate","BlankEditBoxDate","EditBoxList"
				aProperty=Split(dicEditProperties("PropertyName"),"~")
				aValues=Split(dicEditProperties("Value"),"~")
				For iCounter=0 to ubound(aProperty)
					objPropetiesDialog.JavaStaticText("PropertyLabel").SetTOProperty "label",aProperty(iCounter)+":"
					If objPropetiesDialog.JavaEdit("Property_Edit").Exist(SISW_MICRO_TIMEOUT) Then
						wait 2
						If StrAction = "EditBoxDate" Then
							objPropetiesDialog.JavaEdit("Property_Edit").RefreshObject
							wait 1
							objPropetiesDialog.JavaEdit("Property_Edit").Click 5,5,"LEFT"
							wait 7
							Set WshShell = CreateObject("WScript.Shell")
'							WshShell.SendKeys ("^{a}")
'							wait 1
'							WshShell.SendKeys ("{DELETE}")
'							wait 1
'							Call Fn_UI_EditBox_Type("Fn_SISW_Prop_EditProperties",objPropetiesDialog,"Property_Edit",aValues(iCounter)) ' Added by Jotiba T to type CustomDate 
'							wait 2
'							WshShell.SendKeys "{ENTER}"
							objPropetiesDialog.JavaEdit("Property_Edit").Set aValues(iCounter)
							wait 1
							objPropetiesDialog.JavaEdit("Property_Edit").Activate
							wait 1
							objPropetiesDialog.JavaEdit("Property_Edit").RefreshObject
							wait 1
							WshShell.SendKeys "{TAB}"
							Set WshShell =nothing
							Wait 1
							Call Fn_ReadyStatusSync(1)	
						ElseIf StrAction = "BlankEditBoxDate" Then '[TC11.3(20170509d)_NewDevelopment_PoonamC_06June2017 : Added case to unset value]
							objPropetiesDialog.JavaEdit("Property_Edit").RefreshObject
							wait 1
							objPropetiesDialog.JavaEdit("Property_Edit").Click 5,5,"LEFT"
							wait 7
							Set WshShell = CreateObject("WScript.Shell")
							objPropetiesDialog.JavaEdit("Property_Edit").Set aValues(iCounter)
							wait 1
							objPropetiesDialog.JavaEdit("Property_Edit").Activate
							wait 1
							objPropetiesDialog.JavaEdit("Property_Edit").RefreshObject
							wait 1
							 WshShell.SendKeys "{TAB}"
'							WshShell.SendKeys ("^{a}")
'							wait 1
'							WshShell.SendKeys ("{DELETE}")
'							wait 1
							'WshShell.SendKeys "{ENTER}"
							Set WshShell =nothing
							Wait 1
							Call Fn_ReadyStatusSync(1)
						Else
							Call Fn_Edit_Box("Fn_SISW_Prop_EditProperties",objPropetiesDialog,"Property_Edit",aValues(iCounter))
							wait 1
								If StrAction="EditBoxList" Then
									Set WshShell = CreateObject("WScript.Shell")
								    WshShell.SendKeys "{TAB}"
								    Wait 2
									Set WshShell =nothing
								End If
							Wait 3
						End If
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Property [ " & aProperty(iCounter) & " ] is not exist on dialog")
						Set objPropetiesDialog=Nothing
						Exit function
					End If
				Next
				Fn_SISW_Prop_EditProperties=True
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to modify property value from edit boxes
		Case "AddClipboardObjects"
				objPropetiesDialog.JavaStaticText("PropertyLabel").SetTOProperty "label",dicEditProperties("PropertyName")+":"
				If objPropetiesDialog.JavaCheckBox("Property_Edit16CheckBox").Exist(SISW_MICRO_TIMEOUT) then
					Call Fn_CheckBox_Set("Fn_SISW_Prop_EditProperties", objPropetiesDialog, "Property_Edit16CheckBox", "ON")
					If objPropetiesDialog.JavaButton("Property_Add16Button").Exist(SISW_MIN_TIMEOUT) Then
						Call Fn_Button_Click("Fn_SISW_Prop_EditProperties", objPropetiesDialog,"Property_Add16Button")
						Fn_SISW_Prop_EditProperties=Fn_CheckBox_Set("Fn_SISW_Prop_EditProperties", objPropetiesDialog, "Property_Edit16CheckBox", "Off")
					End If
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Property [ " & aProperty(iCounter) & " ] is not exist on dialog")
					Set objPropetiesDialog=Nothing
					Exit function
				End if
        ' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to modify property value of Radio Button
		Case "RadioButton"
			aProperty=Split(dicEditProperties("PropertyName"),"~")
			aValues=Split(dicEditProperties("Value"),"~")
			For iCounter=0 to ubound(aProperty)
				objPropetiesDialog.JavaStaticText("PropertyLabel").SetTOProperty "label",aProperty(iCounter)+":"
				objPropetiesDialog.JavaRadioButton("Property_RadioButton").SetTOProperty "attached text",aValues(iCounter)
				If objPropetiesDialog.JavaRadioButton("Property_RadioButton").Exist(SISW_MIN_TIMEOUT) Then
					Call Fn_UI_JavaRadioButton_SetON("Fn_SISW_Prop_EditProperties",objPropetiesDialog, "Property_RadioButton")
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Property [ " & aProperty(iCounter) & " ] is not exist on dialog")
					Set objPropetiesDialog=Nothing
					Exit function
				End If
			Next
			Fn_SISW_Prop_EditProperties=True
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 		
		 'Case to modify property value of Radio Button for Make/Buy
		Case "RadioButtonMakeBuy"					'[TC1122:2016021000:10Mar2016:ShwetambriR:NewDevelopment]
				aProperty=Split(dicEditProperties("PropertyName"),"~")
				aValues=Split(dicEditProperties("Value"),"~")
				For iCounter=0 to ubound(aProperty)
					objPropetiesDialog.JavaRadioButton("Make/Buy").SetTOProperty "attached text",aValues(iCounter)
					If objPropetiesDialog.JavaRadioButton("Make/Buy").Exist(SISW_MIN_TIMEOUT) Then
						Call Fn_UI_JavaRadioButton_SetON("Fn_SISW_Prop_EditProperties",objPropetiesDialog, "Make/Buy")
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Property [ " & aProperty(iCounter) & " ] is not exist on dialog")
						Set objPropetiesDialog=Nothing
						Exit function
					End If
				Next
				Fn_SISW_Prop_EditProperties=True
			
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 		
		'Case to set date
		Case "DateCheckBox"
			aProperty=Split(dicEditProperties("PropertyName"),"~")
			aValues=Split(dicEditProperties("Value"),"~")
			For iCounter=0 to ubound(aProperty)
				objPropetiesDialog.JavaStaticText("PropertyLabel").SetTOProperty "label",aProperty(iCounter)+":"
				objPropetiesDialog.JavaEdit("Property_Edit").Set aValues(iCounter)

				' Set Time
				aTime=Split(aValues(iCounter))
				If UBound(aTime) > 0 Then
					Call Fn_Edit_Box("Fn_SISW_Prop_EditProperties",objPropetiesDialog,"CheckOutTime",aTime(1))
				End If
'				If objPropetiesDialog.JavaCheckBox("Date").Exist(SISW_MIN_TIMEOUT) Then
'					objPropetiesDialog.JavaCheckBox("Date").Object.setDate(aValues(iCounter))
'				else
'					Set objPropetiesDialog=Nothing
'					Exit function
'				End If
			Next
			Fn_SISW_Prop_EditProperties=True
	    ' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'case for custom checkbox from bmide - Suraj Mayande
		Case "CustDateCheckBox"
			aProperty=Split(dicEditProperties("PropertyName"),"~")
			aValues=Split(dicEditProperties("Value"),"~")
			aCase=dicEditProperties("Case")
			Select Case aCase
				Case "Verify"
					objPropetiesDialog.JavaStaticText("PropertyLabel").SetTOProperty "label",aProperty(iCounter)+":"
					If objPropetiesDialog.JavaEdit("CustomTime").Exist(1) Then
						Fn_SISW_Prop_EditProperties=True
					Else
						Fn_SISW_Prop_EditProperties=False
					End If
				Case "Edit"
					For iCounter=0 to ubound(aProperty)
						objPropetiesDialog.JavaStaticText("PropertyLabel").SetTOProperty "label",aProperty(iCounter)+":"
						objPropetiesDialog.JavaEdit("Property_Edit").Set aValues(iCounter)
						' Set Time
						aTime=Split(aValues(iCounter))
						If UBound(aTime) > 0 Then
							If objPropetiesDialog.JavaEdit("CustomTime").Exist(1) Then
								Call Fn_Edit_Box("Fn_SISW_Prop_EditProperties",objPropetiesDialog,"CustomTime",aTime(1))
								sDate=objPropetiesDialog.JavaEdit("Property_Edit").GetROProperty ("value")
								sTime=objPropetiesDialog.JavaEdit("CustomTime").GetROProperty ("value")
								Fn_SISW_Prop_EditProperties = sDate&" "&sTime
							Else
								Fn_SISW_Prop_EditProperties=False
							End If
						End If
					Next
				End select
		
		 ' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		
		Case "LOVTableCellExist","LOVTableSelect"
			'Setting property name
			objPropetiesDialog.JavaStaticText("PropertyLabel").SetTOProperty "label",dicEditProperties("PropertyName")+":"
			'Clicking on LOV Dropdown Button
			Call Fn_Button_Click("Fn_SISW_Prop_EditProperties",objPropetiesDialog,"DropdownButton")
			wait 2
			'Checking Existance of LOV Table
			If objPropetiesDialog.JavaTable("LovTreeTable").Exist(SISW_MIN_TIMEOUT) Then
				bFlag=True
				If dicEditProperties("PrimaryColumnName")="" then
					dicEditProperties("PrimaryColumnName")=0
				Else
					If iSNumeric(dicEditProperties("PrimaryColumnName")) Then
					Else
						bFlag=False
						For iCounter=0 to Cint(objPropetiesDialog.JavaTable("LovTreeTable").GetROProperty("cols"))-1
							If objPropetiesDialog.JavaTable("LovTreeTable").GetColumnName(iCounter)=dicEditProperties("PrimaryColumnName") Then
								dicEditProperties("PrimaryColumnName")=iCounter
								bFlag=True
								Exit for
							End If
						Next
					End If
				End if
				If bFlag=True Then
					If dicEditProperties("ColumnName")="" Then
						dicEditProperties("ColumnName")=0
					Else
						If iSNumeric(dicEditProperties("ColumnName")) Then
						Else
							bFlag=False
							For iCounter=0 to Cint(objPropetiesDialog.JavaTable("LovTreeTable").GetROProperty("cols"))-1
								If objPropetiesDialog.JavaTable("LovTreeTable").GetColumnName(iCounter)=dicEditProperties("ColumnName") Then
									dicEditProperties("ColumnName")=iCounter
									bFlag=True
									Exit for
								End If
							Next
						End If
					End If
				End If
				If bFlag=True Then
					aPrimaryValues=Split(dicEditProperties("PrimaryColumnValue"),"~")
					aValues=Split(dicEditProperties("Value"),"~")
					For iCounter=0 to ubound(aPrimaryValues)
						bFlag=False
						For iCount=0 to Cint(objPropetiesDialog.JavaTable("LovTreeTable").GetROProperty("rows"))-1
							If objPropetiesDialog.JavaTable("LovTreeTable").Object.getValueAt(iCount,dicEditProperties("PrimaryColumnName")).getDisplayableValue()=Trim(aPrimaryValues(iCounter)) Then
								If StrAction="LOVTableSelect" Then
									objPropetiesDialog.JavaTable("LovTreeTable").ClickCell iCount,0
									bFlag=True
									Exit for
								Else
									If objPropetiesDialog.JavaTable("LovTreeTable").Object.getValueAt(iCount,dicEditProperties("ColumnName")).getDisplayableValue()=Trim(aValues(iCounter)) Then
										bFlag=True
										Exit for
									End If
								End If
							End If
						Next
						If bFlag=False Then
							Exit for
						End If
					Next
					If bFlag=True Then
						Fn_SISW_Prop_EditProperties=True
					End If
				End If
			End If
			If StrAction<>"LOVTableSelect" Then
				Dim WshShell
				Set WshShell = CreateObject("WScript.Shell")
				wait(1)
				WshShell.SendKeys "{ESC}"
				wait(2)
				Set WshShell =Nothing 
'				Call Fn_Button_Click("Fn_SISW_Prop_EditProperties",objPropetiesDialog,"DropdownButton")
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "LOVShellTableSelect"
			'Setting property name
			objPropetiesDialog.JavaStaticText("PropertyLabel").SetTOProperty "label", dicEditProperties("PropertyName") + ":"
			'Clicking on LOV Dropdown Button
			Call Fn_Button_Click("Fn_SISW_Prop_EditProperties", objPropetiesDialog, "DropdownButton")
			wait 2
			For iCounter = 1 To 30
				objPropetiesDialog.JavaWindow("Shell").SetTOProperty "index", iCounter
				If objPropetiesDialog.JavaWindow("Shell").JavaTable("LOVProp").Exist(1) Then
					bFlag = True
					Exit For
				Else
					bFlag = False				
				End If
			Next

			If bFlag = True Then
				bFlag = False
				For iCnt = 0 To objPropetiesDialog.JavaWindow("Shell").JavaTable("LOVProp").GetROProperty("cols") - 1
					If objPropetiesDialog.JavaWindow("Shell").JavaTable("LOVProp").GetColumnName(iCnt) = dicEditProperties("PrimaryColumnName") Then
						bFlag = True
						Exit For
					End If			
				Next
				If bFlag = False Then
					'Clicking on button
					Call Fn_Button_Click("Fn_SISW_Prop_EditProperties", objPropetiesDialog, "Close")
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"LOVProp table Column does not exist in dialog")
					Set objPropetiesDialog=Nothing			
					Exit Function					
				End If
				bFlag = False
				For iCount = 0 To objPropetiesDialog.JavaWindow("Shell").JavaTable("LOVProp").GetROProperty("rows")
					If objPropetiesDialog.JavaWindow("Shell").JavaTable("LOVProp").GetCellData(iCount, iCnt + 1) = dicEditProperties("PrimaryColumnValue") Then
						objPropetiesDialog.JavaWindow("Shell").JavaTable("LOVProp").ClickCell iCount, iCnt + 1
						wait 1
						bFlag = True
						Exit For
					End If
				Next
				If bFlag = False Then
					'Clicking on button
					Call Fn_Button_Click("Fn_SISW_Prop_EditProperties", objPropetiesDialog, "Close")
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"LOVProp table row does not exist in dialog")
					Set objPropetiesDialog=Nothing			
					Exit Function					
				End If				
			Else
				'Clicking on button
				Call Fn_Button_Click("Fn_SISW_Prop_EditProperties", objPropetiesDialog, "Close")
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"LOVProp table does not exist in dialog")
				Set objPropetiesDialog=Nothing			
				Exit Function
			End If
			
			If bFlag=True Then
						Fn_SISW_Prop_EditProperties=True
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to add values in List box using edit boxe
		Case "EditBox_AddToList"
			aValues=Split(dicEditProperties("Value"),"~")
			objPropetiesDialog.JavaStaticText("PropertyLabel").SetTOProperty "label",dicEditProperties("PropertyName")+":"
			If objPropetiesDialog.JavaList("Property_List").Exist(SISW_MICRO_TIMEOUT) Then
				Call Fn_CheckBox_Set("Fn_SISW_Prop_EditProperties", objPropetiesDialog, "Property_Edit16CheckBox", "ON")
				For iCounter=0 to ubound(aValues)
					objPropetiesDialog.JavaEdit("Property_ListEdit").SetFocus
					Call Fn_Edit_Box("Fn_SISW_Prop_EditProperties",objPropetiesDialog,"Property_ListEdit",aValues(iCounter))
					Call Fn_Button_Click("Fn_SISW_Prop_EditProperties", objPropetiesDialog,"Property_Add16Button")
				Next
				Call Fn_CheckBox_Set("Fn_SISW_Prop_EditProperties", objPropetiesDialog, "Property_Edit16CheckBox", "OFF")
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Property [ " & aProperty(iCounter) & " ] is not exist on dialog")
				Set objPropetiesDialog=Nothing
				Exit function
			End If
			Fn_SISW_Prop_EditProperties=True
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "PasteLink"
	        objPropetiesDialog.JavaStaticText("PropertyLabel").SetTOProperty "label",dicEditProperties("PropertyName")+":"
'			objPropetiesDialog.JavaObject("LinkOptionDropDown").Click 1,1
			objPropetiesDialog.JavaStaticText("StaticLinkOptionDropDown").Click 1,1
			Select Case StrAction
					' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
					Case "PasteLink"
						objPropetiesDialog.JavaMenu("index:=0","label:=Paste").Select
			End Select
			If Err.Number < 0 Then
				Fn_SISW_Prop_EditProperties=False
			Else
				Fn_SISW_Prop_EditProperties=True
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to add values in List box using LOV Table
		Case "LOVTable_AddToList"
			aValues=Split(dicEditProperties("Value"),"~")
			objPropetiesDialog.JavaStaticText("PropertyLabel").SetTOProperty "label",dicEditProperties("PropertyName")+":"
			If objPropetiesDialog.JavaList("Property_List").Exist(SISW_MICRO_TIMEOUT) Then
				Call Fn_CheckBox_Set("Fn_SISW_Prop_EditProperties", objPropetiesDialog, "Property_Edit16CheckBox", "ON")
				For iCounter=0 to ubound(aValues)
					bFlag=False
					Call Fn_Button_Click("Fn_SISW_Prop_EditProperties",objPropetiesDialog,"Property_ListLOVDropDownButton")
					wait 2
					For iCount=0 to Cint(objPropetiesDialog.JavaTable("LovTreeTable").GetROProperty("rows"))-1
						If objPropetiesDialog.JavaTable("LovTreeTable").Object.getValueAt(iCount,0).getDisplayableValue()=Trim(aValues(iCounter)) Then
								objPropetiesDialog.JavaTable("LovTreeTable").ClickCell iCount,0
								bFlag=True
								Exit for
							End If
					Next
					If bFlag=False Then
						Exit for
					End If
					Call Fn_Button_Click("Fn_SISW_Prop_EditProperties", objPropetiesDialog,"Property_Add16Button")
				Next
				Call Fn_CheckBox_Set("Fn_SISW_Prop_EditProperties", objPropetiesDialog, "Property_Edit16CheckBox", "OFF")
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Property [ " & aProperty(iCounter) & " ] is not exist on dialog")
				Set objPropetiesDialog=Nothing
				Exit function
			End If
			If bFlag=True Then
				Fn_SISW_Prop_EditProperties=True
			End If
		'[TC1122-20151116b-16_12_2015-VivekA-NewDevelopment] - Case to modify existing values in List box using edit box
		Case "EditBox_ModifyandAddToList"
			aValues=Split(dicEditProperties("Value"),"~")
			objPropetiesDialog.JavaStaticText("PropertyLabel").SetTOProperty "label",dicEditProperties("PropertyName")+":"
			If objPropetiesDialog.JavaList("Property_List").Exist(SISW_MICRO_TIMEOUT) Then
				Call Fn_CheckBox_Set("Fn_SISW_Prop_EditProperties", objPropetiesDialog, "Property_Edit16CheckBox", "ON")
				For iCounter=0 to ubound(aValues)
					arrValue = Split(aValues(iCounter),":")
					'Select value from list to modify
					objPropetiesDialog.JavaList("Property_List").select "#"+arrValue(0)
					'Modify the selected value
					objPropetiesDialog.JavaEdit("Property_ListEdit").SetFocus
					Call Fn_Edit_Box("Fn_SISW_Prop_EditProperties",objPropetiesDialog,"Property_ListEdit",arrValue(1))
					Call Fn_Button_Click("Fn_SISW_Prop_EditProperties", objPropetiesDialog,"Property_Modify16Button")
				Next
				Call Fn_CheckBox_Set("Fn_SISW_Prop_EditProperties", objPropetiesDialog, "Property_Edit16CheckBox", "OFF")
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Property [ " & aProperty(iCounter) & " ] is not exist on dialog")
				Set objPropetiesDialog=Nothing
				Exit function
			End If
			Fn_SISW_Prop_EditProperties=True
		'Case to sort Values in Content List - Added by Poonam Chopade
		Case "SortContentList"
				objPropetiesDialog.JavaStaticText("PropertyLabel").SetTOProperty "label",dicEditProperties("PropertyName")+":"
				If objPropetiesDialog.JavaList("Property_List").Exist(SISW_MICRO_TIMEOUT) Then
					Call Fn_CheckBox_Set("Fn_SISW_Prop_EditProperties", objPropetiesDialog, "Property_Edit16CheckBox", "ON")
					wait 1
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Property [ " & aProperty(iCounter) & " ] is not exist on dialog")
					Set objPropetiesDialog=Nothing
					Exit Function
				End If
				If objPropetiesDialog.JavaButton("sorting_16").Exist(SISW_MICRO_TIMEOUT) Then
				     Call Fn_Button_Click("Fn_SISW_Prop_EditProperties",objPropetiesDialog,"sorting_16")
					 wait 1
				Else
				 	Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Button [ sorting_16 ] is not exist on dialog")
					Set objPropetiesDialog=Nothing
					Exit Function
				End If
				Fn_SISW_Prop_EditProperties=True 
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 	
		'Case to verify the sorted Values in Content List - Added by Poonam Chopade
		Case "VerifySortedContentList"
				objPropetiesDialog.JavaStaticText("PropertyLabel").SetTOProperty "label",dicEditProperties("PropertyName")+":"
				If objPropetiesDialog.JavaList("Property_List").Exist(SISW_MICRO_TIMEOUT) Then
                    iEleCount=Fn_UI_Object_GetROProperty("Fn_SISW_Prop_EditProperties",objPropetiesDialog.JavaList("Property_List"), "items count")
					For iCount=0 To iEleCount-1
						If iCount = 0 Then
							sAppValues = objPropetiesDialog.JavaList("Property_List").GetItem(iCount)
						Else	
							sAppValues = sAppValues & "~" & objPropetiesDialog.JavaList("Property_List").GetItem(iCount)
						End If
					Next
					If Instr(1,Trim(sAppValues),Trim(dicEditProperties("Value"))) > 0 Then
						Fn_SISW_Prop_EditProperties=True
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Value [ " & aValues(iCounter) & " ] is not exist in [ " & dicProperties("PropertyName") & " ] List")
						Fn_SISW_Prop_EditProperties=False
					End If
				End If	
	End Select
	'Clicking on button
	If StrButtonName<>"" Then
		Call Fn_Button_Click("Fn_SISW_Prop_EditProperties", objPropetiesDialog,StrButtonName)
		Call Fn_ReadyStatusSync(2)
		'saving changes and checking in
		If StrButtonName="SaveAndCheck-In" Then
				Set objCheckOut= Fn_SISW_GetChkInChkOutObject("CheckIn")
				If typename(objCheckOut) <> "Nothing" Then
					'ckicking on Yes button to checkout object
					Call Fn_Button_Click("Fn_SISW_Prop_EditProperties",objCheckOut,"Yes")
				End If
				Set objCheckOut= Nothing
			Call Fn_ReadyStatusSync(1)
		End If
	End If
	Set objPropetiesDialog=Nothing
End Function
