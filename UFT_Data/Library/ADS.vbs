Option Explicit

'*********************************************************	Function List		***********************************************************************
'0. Fn_ADS_SISW_GetObject()
'1. Fn_ADS_ChangeOperations()
'2.Fn_ADS_ObjectPropertyCkOutEditCkIn()
'3.Fn_ADS_ProgramAssociation()
'4.Fn_ADS_MeasureTextLength()
'5.Fn_ADS_PartDetailCreate()
'6.Fn_ADS_SessionInfoVerify()
'7.Fn_ADS_SummaryTabOperation()
'8.Fn_ADS_DialogHandle()
'9.Fn_ADS_ObjectPropertyChkOut()
'10.Fn_ADS_ChangeOperationsExtn()
'11.Fn_ADS_VariableLenRandomNoGen()
'12.Fn_ContractEventSchedule()
'13.Fn_ADS_GenerateSubmittalDelivery()
'14.Fn_ADS_ScheduleDetailCreate()
'15.Fn_ADS_ViewerTabOperation()
'16.Fn_ADS_SetDateViewerTab()
'17.Fn_ADS_PropertiesOperations()
'18.Fn_ADS_ChangeOperationsDic()
'19.Fn_ADS_VerifyChangeObjects()
'20.Fn_ADS_ObjectROPropertyCheck()
'21.Fn_ADS_TechDocOperations()
'22.Fn_ADS_ParametricValues()
'23.Fn_ADS_SubmittalsTableContentOperation()
'24.Fn_ADS_PropertyRetrive()
'25.Fn_ADS_ItemListVerify()
'26.Fn_ADS_GenerateSubmittalDeliverySync()
'27.Fn_ADS_DeliverableTableContentOperation()
'28.Fn_ADS_DialogMsgVerify()
'29.Fn_ADS_UIOperations()
'30.Fn_ADS_ItemDetailCreate()
'31.Fn_ADS_PartDetailCreateDic()
'32.Fn_ADS_DesignDetailCreateDic()
'33.Fn_ADS_CustomNoteCreate()
'34. Fn_SISW_ADS_JavaTable_GetCellData()
'35. Fn_ADS_AssignCompanyLocation()
'36 Fn_SISW_ADS_ErrorVerify()
'*********************************************************	Function List		***********************************************************************

'****************************************    Function to get Object hierarchy ***************************************
'
''Function Name		 	:	Fn_ADS_SISW_GetObject
'
''Description		    :  	Function to get specified Object hierarchy.

''Parameters		    :	1. sObjectName : Object Handle name
								
''Return Value		    :  	Object \ Nothing
'
''Examples		     	:	Fn_ADS_SISW_GetObject("New Part")

'History:
'	Developer Name			Date			Rev. No.		Reviewer		Changes Done	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Swapna Ghatge		 23-July-2012		1.0				
'-----------------------------------------------------------------------------------------------------------------------------------
'	Ashwini Kumar		 25-Sept-2013		2.0								Externalized object hierarchies
'-----------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_ADS_SISW_GetObject(sObjectName)
	Dim sObjectXMLPath
	sObjectXMLPath = Fn_GetEnvValue("User", "AutomationDir") & "\TestData\AutomationXML\ObjectXML\ADS.xml"
	Set Fn_ADS_SISW_GetObject = Fn_SISW_Setup_GetObjectFromXML(sObjectXMLPath, sObjectName)
End Function

'-------------------------------------------------------------------Function Used to Create Proplem Report,Change Notice,Change Request,Deviation Request----------------------------------------------------------
'Function Name  : Fn_ADS_ChangeOperations
'Description    : Function Used to Create Proplem Report,Change Notice,Change Request,Deviation Request
'Return Value     :  True Or False
'Examples    :  Case "Set" : Call Fn_ADS_ChangeOperations("Set", "", "Change Notice", "ECN-654321", "A", "Test", "Testing", "Type", "", "Finish:Cancel")
'         			  Case "Verify" : Call Fn_ADS_ChangeOperations("Verify", "", "Problem Report", "", "", "", "", "", "SelectionAddition:Problem Items", "Cancel")
             
'History      :   
'             Developer Name            Date      Rev. No.      Changes Done      Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'             Ketan Raje.                 04/10/2010              1.0                   Harshal
'             Sandeep N                 09-Aug-2012              1.1       modified case : Verify -> Case : RelationshipTable ( Added native method to get cell data )          	Swapna
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_ADS_ChangeOperations(sAction, strFilterText, sNodeName, sChangeID, sChangeRev, sChangeSynopsis, sChangeDesc, sChangeType, aDiffFields, sButtons)
	GBL_FAILED_FUNCTION_NAME="Fn_ADS_ChangeOperations"
   'Declaring Variables
    Dim aButtons, iCount, iCounter, intCount, bFlag, iRows, arrDiffFields
    Dim ObjChangeWnd, intNodeCount, sTreeItem, sDiffFields
 Fn_ADS_ChangeOperations=False
  For iCount=0 to 0
     JavaWindow("DefaultWindow").JavaWindow("New Change").SetTOProperty "title","New Change"
     If JavaWindow("DefaultWindow").JavaWindow("New Change").Exist(5) Then
      Exit For
     End If
     JavaWindow("DefaultWindow").JavaWindow("New Change").SetTOProperty "title","New Change in context"
     If JavaWindow("DefaultWindow").JavaWindow("New Change").Exist(5) Then
      Exit For
     End If
     JavaWindow("DefaultWindow").JavaWindow("New Change").SetTOProperty "title","Derive Change"
     If JavaWindow("DefaultWindow").JavaWindow("New Change").Exist(5) Then      
      Exit For
     End If
  Next
  Set ObjChangeWnd=Fn_UI_ObjectCreate("Fn_ADS_ChangeOperations",JavaWindow("DefaultWindow").JavaWindow("New Change"))
 Select Case sAction
  Case "Set"
   'Set FilterText
   strFilterText = sNodeName
   
   If strFilterText <> "" Then
   
     If Fn_UI_ObjectExist("Fn_ADS_ChangeOperations", ObjChangeWnd.JavaTree("ChangeTypeTree"))=False Then
		  Call Fn_Button_Click("Fn_ADS_ChangeOperations",ObjChangeWnd,"Back")
		  Wait(1)
	 End If
     
    Call Fn_UI_EditBox_Type("Fn_ADS_ChangeOperations",ObjChangeWnd,"FilterText",strFilterText)
    
   End If
   
   If sNodeName<>"" Then
    Wait(10)
    'Selecting Node from tree
    Call Fn_JavaTree_Select("Fn_ADS_ChangeOperations", ObjChangeWnd, "ChangeTypeTree","Complete List")
    strNodePath="Complete List:"+sNodeName
     'Verifying Node is present in Tree
     intNodeCount =Fn_UI_Object_GetROProperty("Fn_ADS_ChangeOperations",ObjChangeWnd.JavaTree("ChangeTypeTree"),"items count")    
     For intCount = 0 to intNodeCount - 1
      sTreeItem = ObjChangeWnd.JavaTree("ChangeTypeTree").GetItem(intCount)
      If Trim(lcase(sTreeItem)) = Trim(Lcase(strNodePath)) Then
       Fn_ADS_ChangeOperations = True
       Exit For
      End If
     Next
     If Cint(intCount) =Cint( intNodeCount) Then
      Call Fn_Button_Click("Fn_ADS_ChangeOperations",ObjChangeWnd,"Cancel")
      Fn_ADS_ChangeOperations =False
      Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Fail: "& strNodePath &"Node is not found in View Tree")
      Set ObjChangeWnd=Nothing
      Exit Function
     End If
    Call Fn_JavaTree_Select("Fn_ADS_ChangeOperations", ObjChangeWnd, "ChangeTypeTree","Complete List:"+sNodeName)
    'Clicking on Next button to proceed 
    Call Fn_Button_Click("Fn_ADS_ChangeOperations",ObjChangeWnd,"Next")
   End If
   'Set Change Id
   wait 5
   If sChangeID<>"" Then
   	Call Fn_SISW_UI_JavaEdit_Operations("Fn_ADS_ChangeOperations", "Set",  ObjChangeWnd, "ID", sChangeID )
    'Call Fn_UI_EditBox_Type("Fn_ADS_ChangeOperations",ObjChangeWnd,"ID",sChangeID)
   End If
   wait 2
   'Set Change Revision
   If sChangeRev<>"" Then
    Call Fn_UI_EditBox_Type("Fn_ADS_ChangeOperations",ObjChangeWnd,"Revision",sChangeRev)
   End If
   wait 2
   'Set Synopsis
   If sChangeSynopsis<>"" Then
    Call Fn_UI_EditBox_Type("Fn_ADS_ChangeOperations",ObjChangeWnd,"Name",sChangeSynopsis)
   End If
   wait 2
   'Set Description
   If sChangeDesc<>"" Then
    Call Fn_UI_EditBox_Type("Fn_ADS_ChangeOperations",ObjChangeWnd,"Description",sChangeDesc)
   End If
   wait 2
   'Set Change Type
   If sChangeType<>"" Then
		If strFilterText="Deviation Request" Then
			Call Fn_UI_Object_SetTOProperty_ExistCheck("RACUpdateActionItem",JavaWindow("DefaultWindow").JavaWindow("New Change").JavaStaticText("Change Type"),"label","Deviation Type:")
		End If
		Call Fn_UI_EditBox_Type("Fn_ADS_ChangeOperations",ObjChangeWnd,"Change Type",sChangeType)		
   End If
   'Code for Fields differing in Case of CN,CR,PR,DR is yet to be coded
   'Set Change Type
   If aDiffFields<>"" Then
    sDiffFields = Split(aDiffFields,"|",-1,1)
    For iCount = 0 to ubound(sDiffFields)
     arrDiffFields = Split(sDiffFields(iCount),":",-1,1)
     Call Fn_UI_EditBox_Type("Fn_ADS_ChangeOperations",ObjChangeWnd,arrDiffFields(0),arrDiffFields(1))
    Next
   End If
  Case "Verify"
   'Set FilterText
    strFilterText = sNodeName
    
    If strFilterText<>"" Then
        
       If Fn_UI_ObjectExist("Fn_ADS_ChangeOperations", ObjChangeWnd.JavaTree("ChangeTypeTree"))=False Then
		  Call Fn_Button_Click("Fn_ADS_ChangeOperations",ObjChangeWnd,"Back")
		  wait(1)
	   End If
        
		Call Fn_UI_EditBox_Type("Fn_ADS_ChangeOperations",ObjChangeWnd,"FilterText",strFilterText)
		
    End If
    
    If sNodeName<>"" Then
     Call Fn_JavaTree_Select("Fn_ADS_ChangeOperations", ObjChangeWnd, "ChangeTypeTree","Complete List")
     Call Fn_JavaTree_Select("Fn_ADS_ChangeOperations", ObjChangeWnd, "ChangeTypeTree","Complete List:"+sNodeName)
    End If
    Wait(4)
    bFlag = False
    sDiffFields = Split(aDiffFields,":",-1,1)
    Select Case sDiffFields(0)
     Case "RelationshipTable"
				Call Fn_Button_Click("Fn_ADS_ChangeOperations",ObjChangeWnd,"Next")
			   iRows = ObjChangeWnd.JavaTable("RelationshipTable").GetROProperty("rows")
			   For iCount=0 to iRows-1
				If Trim(Lcase(ObjChangeWnd.JavaTable("RelationshipTable").Object.getItem(iCount).getData().toString())) = Trim(Lcase(sDiffFields(1))) Then
				 Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: "& sDiffFields(1) &"found successfully in Relationship Table.")
				 bFlag = True
				 Exit For
				End If
			   Next
     Case "SelectionAddition"
			   If Trim(Lcase(Fn_Edit_Box_GetValue("Fn_ADS_ChangeOperations",ObjChangeWnd,"SelectionAddition"))) = Trim(Lcase(sDiffFields(1))) Then
				 Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: "& sDiffFields(1) &"found successfully in Selection addition editbox.")
				 bFlag = True
			   End If
	 End Select
				If bFlag=False Then
					  Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: "& sDiffFields(1) &" failed in verification")
					  'Click on Buttons
					  If sButtons<>"" Then
						aButtons = split(sButtons, ":",-1,1)
						iCounter = Ubound(aButtons)
						For iCount=0 to iCounter
						 'Click on Add Button
						 Call Fn_Button_Click("Fn_ADS_ChangeOperations", ObjChangeWnd, aButtons(iCount))						        
						Next
					  End If
					  Fn_ADS_ChangeOperations=False
					  Set ObjChangeWnd=Nothing
					  Exit Function
				Else
					  Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: "& sDiffFields(1) &" value successfully verified.")
				End If
	Case "CompleteListTreeExist"
	
				If Fn_UI_ObjectExist("Fn_ADS_ChangeOperations", ObjChangeWnd.JavaTree("ChangeTypeTree"))=False Then
					Call Fn_Button_Click("Fn_ADS_ChangeOperations",ObjChangeWnd,"Back")
				End If	
	
				 If sNodeName<>"" Then
					strNodePath="Complete List:"+sNodeName			
					  intNodeCount =Fn_UI_Object_GetROProperty("Fn_ADS_ChangeOperations",ObjChangeWnd.JavaTree("ChangeTypeTree"),"items count")    
							For intCount = 0 to intNodeCount -1
								sTreeItem =ObjChangeWnd.JavaTree("ChangeTypeTree").GetItem(intCount)
								If Trim(lcase(sTreeItem)) = Trim(Lcase(strNodePath)) Then									
									Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: "& sNodeName &" Is Exist")
								End If
							Next
							If Cint(intCount) = Cint(intNodeCount) Then
								  Call Fn_Button_Click("Fn_ADS_ChangeOperations",ObjChangeWnd,"Cancel")
								  Fn_ADS_ChangeOperations =False
								  Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: "& strNodePath &"Node is not found in View Tree")
								  Set ObjChangeWnd=Nothing
								  Exit Function
							End If
				End If
			 End Select
 'Click on Buttons
 If sButtons<>"" Then
   aButtons = split(sButtons, ":",-1,1)
   iCounter = Ubound(aButtons)
   For iCount=0 to iCounter
    'Click on Add Button
    Call Fn_Button_Click("Fn_ADS_ChangeOperations", ObjChangeWnd, aButtons(iCount))
    Call Fn_ReadyStatusSync(2)
   Next
 End If
 Fn_ADS_ChangeOperations = TRUE
 Set ObjChangeWnd=Nothing
End Function
'######################################################################################################################################
'###    FUNCTION NAME   :  Fn_ADS_ObjectPropertyCkOutEditCkIn(sAction, sObjProperty,sObjPropertyValue)
'###
'###    DESCRIPTION     : Checkout Edit and Checkin operation for business object in Viewer Tab.
'###
'###	 HISTORY         :   		AUTHOR                 DATE        VERSION		BUILD
'###
'###    CREATED BY      :     Ketan       		  		   06/10/10      1.0				916
'###
'###    REVIWED BY      :   Harshal							 	
'###
'###    EXAMPLE         :   
'######################################################################################################################################
Function Fn_ADS_ObjectPropertyCkOutEditCkIn(sAction, sObjProperty, sObjPropertyValue)
	GBL_FAILED_FUNCTION_NAME="Fn_ADS_ObjectPropertyCkOutEditCkIn"
   Dim ObjChkOut
	Set ObjChkOut = JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame")
	'Click on 'Check-Out ' Button
	'Call Fn_ToolbatButtonClick("Check Out...")
	 If Fn_UI_ObjectExist("Fn_ObjectCheckOut", JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Check-Out")) = False Then
		'Selecting Check Out  and edit buttoin from Summary Toolbar
		If JavaWindow("MyTeamcenter").JavaWindow("MyTcJApplet").JavaButton("Check-Out and Edit").Exist(5) Then
			Call Fn_Button_Click("Fn_ObjectCheckOut", JavaWindow("MyTeamcenter").JavaWindow("MyTcJApplet"), "Check-Out and Edit")
		ElseIf  JavaWindow("MyTeamcenter").JavaToolbar("CheckInCheckOutToolbar").Exist(5) Then
			' buttons are replaced with JavaToolbar after deploying ADS template - Ashok Kakade, Koustubh Watwe
			Call Fn_UI_JavaToolbar_Press("Fn_ObjectCheckOut", JavaWindow("MyTeamcenter"), "CheckInCheckOutToolbar","Check-Out and Edit")
		ElseIf 	Fn_ToolbatButtonClick("Check-Out and Edit") =True then
				objCheckOut=True
		ElseIf 	Fn_ToolbatButtonClick("Check Out...") =True then
				objCheckOut=True		
		Else							
			 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_ObjectCheckOut : failed to find Check out... button / toolbar.")
			exit function
		End If
	End If

	If JavaWindow("DefaultWindow").JavaWindow("Check-Out").Exist(2) Then
		Call Fn_Button_Click("Fn_ObjectCheckOut", JavaWindow("DefaultWindow").JavaWindow("Check-Out"), "OK")
	End If
	
	If JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Check-Out").Exist(2) Then 
		Call Fn_Button_Click("Fn_ObjectCheckOut", JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Check-Out"), "Yes")
	End If
	
	Wait(5)
	'Click on Static text
	If JavaWindow("DefaultWindow").JavaTab("GeneralTab").Exist(3) Then
		JavaWindow("DefaultWindow").JavaTab("GeneralTab").Select "All"
	Else
	   ObjChkOut.JavaStaticText("BottomLink").SetTOProperty "label","All"
		wait 1
		If ObjChkOut.JavaStaticText("BottomLink").Exist(3) Then
			ObjChkOut.JavaStaticText("BottomLink").Click 1,1,"LEFT"
			wait 2
		End If
	End If

		Select Case sAction
			Case "Set"
					'Set the Property Value
					JavaWindow("DefaultWindow").JavaStaticText("ViewerTab_Text").SetTOProperty "label",sObjProperty & ":"
					Call Fn_Edit_Box("Fn_ADS_ObjectPropertyCkOutEditCkIn",JavaWindow("DefaultWindow"),"ViewerTab_Edit",sObjPropertyValue)
						If Err.Number <0 Then
									Fn_ADS_ObjectPropertyCkOutEditCkIn = False
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL:Property verification failed of Function Fn_ADS_ObjectPropertyCkOutEditCkIn" )
									Call Fn_ToolbatButtonClick("Check In...")
									Wait(2)
									If JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Check-In").Exist(2) Then 
										Call Fn_Button_Click("Fn_ObjectCheckOut", JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Check-In"), "Yes")
									End If
									Exit Function
						End If
					Wait(2)
			Case "Verify"
					JavaWindow("DefaultWindow").JavaStaticText("ViewerTab_Text").SetTOProperty "label",sObjProperty & ":"
					'Verify the Property Value
					If Trim(Lcase(Fn_Edit_Box_GetValue("Fn_ADS_ObjectPropertyCkOutEditCkIn",JavaWindow("DefaultWindow"),"ViewerTab_Edit"))) = Trim(Lcase(sObjPropertyValue)) Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS:Property '"&sObjProperty&"' Successfully Verified to "&sObjPropertyValue&" of Function Fn_ADS_ObjectPropertyCkOutEditCkIn" )												
					Else
						Fn_ADS_ObjectPropertyCkOutEditCkIn = False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL:Property verification failed of Function Fn_ADS_ObjectPropertyCkOutEditCkIn" )
						Call Fn_ToolbatButtonClick("Check In...")
						Wait(2)
						If JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Check-In").Exist(2) Then 
							Call Fn_Button_Click("Fn_ADS_ObjectPropertyCkOutEditCkIn", JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Check-In"), "Yes")
						End If
						Exit Function
					End If
		End Select		
		Call Fn_ToolbatButtonClick("Check In...")
		
		If JavaWindow("DefaultWindow").JavaWindow("Check-In").Exist(2) Then
			Call Fn_Button_Click("Fn_ADS_ObjectPropertyCkOutEditCkIn", JavaWindow("DefaultWindow").JavaWindow("Check-In"), "OK")
		End If 
		
		Wait(2)
		If JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Check-In").Exist(2) Then 
			Call Fn_Button_Click("Fn_ADS_ObjectPropertyCkOutEditCkIn", JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Check-In"), "Yes")
		End If
		Wait(2)
		Fn_ADS_ObjectPropertyCkOutEditCkIn = True
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &"Case executed successfully of function Fn_ADS_ObjectPropertyCkOutEditCkIn")
End Function
'#########################################################################################################
'###
'###    FUNCTION NAME   :   Fn_ADS_ProgramAssociation(sAction,sProgram,sParent,sChild,sButtons)
'###
'###    DESCRIPTION        :   Add / Remove Programs
'###
'###    PARAMETERS      :   1. sAction: Show / Hide
'###											 2.	sProgram : Program Name
'###											 2.	sParent : 
'###											 2.	sChild : 
'###											 2.	sButtons : Buttons to be clicked.
'###                                         
'###    Function Calls       :   Fn_WriteLogFile(), Fn_UI_ObjectCreate(), Fn_Button_Click()
'###
'###	 HISTORY             :   AUTHOR                 DATE        VERSION
'###
'###    CREATED BY     :   Ketan Raje           12/10/2010         1.0
'###
'###    REVIWED BY     :   Harshal	             12/10/2010         1.0				
'###
'###    MODIFIED BY   :  
'###
'###    EXAMPLE          : 		Case "Add" : Call Fn_ADS_ProgramAssociation("Add","AutoProj_1_1182010134651:AutoProj_1_11820101511","Item Revision","0","Apply:Cancel")
'###										 Case "Remove" : Call Fn_ADS_ProgramAssociation("Remove","AutoProj_1_1182010134651:AutoProj_1_11820101511","","","Apply:Cancel")
'###										 Case "AvailableListisEmpty" : Call Fn_ADS_ProgramAssociation("AvailableListisEmpty","","","","Cancel")
'###										 Case "SelectedListisEmpty" : Call Fn_ADS_ProgramAssociation("SelectedListisEmpty","","","","Cancel")
'###										 Case "AvailableListVerify" : Call Fn_ADS_ProgramAssociation("AvailableListVerify","AutoProj_1682010123421","","","Cancel")
'###										 Case "SelectedListVerify" : Call Fn_ADS_ProgramAssociation("SelectedListVerify","AutoProj_1682010124841","","","Cancel")
'#############################################################################################################
Public Function Fn_ADS_ProgramAssociation(sAction,sProgram,sParent,sChild,sButtons)
	GBL_FAILED_FUNCTION_NAME="Fn_ADS_ProgramAssociation"
	Dim objProgram, iCounter, bReturn, aColname, iCount, iRowData, aButtons, iCnt
						For iCount=0 to 0
								JavaWindow("ADS-TeamCenter").JavaWindow("TcDefaultApplet").JavaDialog("AssignObjectToProgram").SetTOProperty "title","Assign an Object to Program"
								If JavaWindow("ADS-TeamCenter").JavaWindow("TcDefaultApplet").JavaDialog("AssignObjectToProgram").Exist(3) Then
									Set objProgram = Fn_UI_ObjectCreate("Fn_ADS_ProgramAssociation", JavaWindow("ADS-TeamCenter").JavaWindow("TcDefaultApplet").JavaDialog("AssignObjectToProgram"))
									Exit For
								End If
								JavaWindow("ADS-TeamCenter").JavaWindow("TcDefaultApplet").JavaDialog("AssignObjectToProgram").SetTOProperty "title","Assign Objects to Program"
								If JavaWindow("ADS-TeamCenter").JavaWindow("TcDefaultApplet").JavaDialog("AssignObjectToProgram").Exist(3) Then
									Set objProgram = Fn_UI_ObjectCreate("Fn_ADS_ProgramAssociation", JavaWindow("ADS-TeamCenter").JavaWindow("TcDefaultApplet").JavaDialog("AssignObjectToProgram"))
									Exit For
								End If
								JavaWindow("ADS-TeamCenter").JavaWindow("TcDefaultApplet").JavaDialog("AssignObjectToProgram").SetTOProperty "title","Remove an Object from Program"
								If JavaWindow("ADS-TeamCenter").JavaWindow("TcDefaultApplet").JavaDialog("AssignObjectToProgram").Exist(3) Then
									Set objProgram = Fn_UI_ObjectCreate("Fn_ADS_ProgramAssociation", JavaWindow("ADS-TeamCenter").JavaWindow("TcDefaultApplet").JavaDialog("AssignObjectToProgram"))
									Exit For
								End If
								JavaWindow("ADS-TeamCenter").JavaWindow("TcDefaultApplet").JavaDialog("AssignObjectToProgram").SetTOProperty "title","Remove Objects from Program"
								If JavaWindow("ADS-TeamCenter").JavaWindow("TcDefaultApplet").JavaDialog("AssignObjectToProgram").Exist(3) Then
									Set objProgram = Fn_UI_ObjectCreate("Fn_ADS_ProgramAssociation", JavaWindow("ADS-TeamCenter").JavaWindow("TcDefaultApplet").JavaDialog("AssignObjectToProgram"))
									Exit For
								End If
								JavaWindow("ADS-TeamCenter").JavaWindow("ADSWebEmbeddedFrame").JavaDialog("AssignObjectToProgram").SetTOProperty "title","Assign an Object to Program"
								If JavaWindow("ADS-TeamCenter").JavaWindow("ADSWebEmbeddedFrame").JavaDialog("AssignObjectToProgram").Exist(3) Then
									Set objProgram = Fn_UI_ObjectCreate("Fn_ADS_ProgramAssociation", JavaWindow("ADS-TeamCenter").JavaWindow("ADSWebEmbeddedFrame").JavaDialog("AssignObjectToProgram"))
									Exit For
								End If
								JavaWindow("ADS-TeamCenter").JavaWindow("ADSWebEmbeddedFrame").JavaDialog("AssignObjectToProgram").SetTOProperty "title","Remove Objects from Program"
								If JavaWindow("ADS-TeamCenter").JavaWindow("ADSWebEmbeddedFrame").JavaDialog("AssignObjectToProgram").Exist(3) Then
									Set objProgram = Fn_UI_ObjectCreate("Fn_ADS_ProgramAssociation", JavaWindow("ADS-TeamCenter").JavaWindow("ADSWebEmbeddedFrame").JavaDialog("AssignObjectToProgram"))
									Exit For
								End If
						Next
		Select Case sAction
				Case "Add"						
						If sProgram<>"" Then
								bReturn = objProgram.JavaList("OwningProgram").GetROProperty("items count")
								'Extract the index of row at which the object exist.
								aColname = split(sProgram, ":",-1,1)
								iCount = Ubound(aColname)
								For iRowData=0 to iCount
									For iCounter=0 to bReturn-1
										If Trim(lcase(objProgram.JavaList("OwningProgram").GetItem(iCounter))) = Trim(lcase(aColname(iRowData))) then
											objProgram.JavaList("OwningProgram").Select aColname(iRowData)
											'Click on Add Button
											Call Fn_Button_Click("Fn_ADS_ProgramAssociation", objProgram, "AddProgram")
											Exit For 
										End If
									Next
								Next
						End If
				Case "Remove" 
						If sProgram<>"" Then
								bReturn = objProgram.JavaList("OwningProgram").GetROProperty("items count")
								'Extract the index of row at which the object exist.
								aColname = split(sProgram, ":",-1,1)
								iCount = Ubound(aColname)
								For iRowData=0 to iCount
									For iCounter=0 to bReturn-1
										If Trim(lcase(objProgram.JavaList("OwningProgram").GetItem(iCounter))) = Trim(lcase(aColname(iRowData))) then
											objProgram.JavaList("OwningProgram").Select aColname(iRowData)
											'Click on Remove Button
											Call Fn_Button_Click("Fn_ADS_ProgramAssociation", objProgram, "AddProgram")											
											Exit For 
										End If
									Next
								Next
						End If
			Case "AvailableListisEmpty"
						If objProgram.JavaList("OwningProgram").GetROProperty("items count") = 0 Then
									'Click on Buttons
									If sButtons<>"" Then
											aButtons = split(sButtons, ":",-1,1)
											iCount = Ubound(aButtons)
											For iRowData=0 to iCount
												'Click on Add Button
												Call Fn_Button_Click("Fn_ADS_ProgramAssociation", objProgram, aButtons(iRowData))
											Next
									End If		
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" Sucessfully completed of function Fn_ADS_ProgramAssociation")
									Set objProgram = nothing 
									Fn_ADS_ProgramAssociation = True
									Exit Function
						Else
									'Click on Buttons
									If sButtons<>"" Then
											aButtons = split(sButtons, ":",-1,1)
											iCount = Ubound(aButtons)
											For iRowData=0 to iCount
												'Click on Add Button
												Call Fn_Button_Click("Fn_ADS_ProgramAssociation", objProgram, aButtons(iRowData))
											Next
									End If		
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" Sucessfully completed of function Fn_ADS_ProgramAssociation")
									Set objProgram = nothing 
									Fn_ADS_ProgramAssociation = False
									Exit Function
						End If
			Case "SelectedListisEmpty"
						If objProgram.JavaList("SelectedProgram").GetROProperty("items count") = 0 Then
									'Click on Buttons
									If sButtons<>"" Then
											aButtons = split(sButtons, ":",-1,1)
											iCount = Ubound(aButtons)
											For iRowData=0 to iCount
												'Click on Add Button
												Call Fn_Button_Click("Fn_ADS_ProgramAssociation", objProgram, aButtons(iRowData))
											Next
									End If		
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" Sucessfully completed of function Fn_ADS_ProgramAssociation")
									Set objProgram = nothing 
									Fn_ADS_ProgramAssociation = True
									Exit Function
						Else
									'Click on Buttons
									If sButtons<>"" Then
											aButtons = split(sButtons, ":",-1,1)
											iCount = Ubound(aButtons)
											For iRowData=0 to iCount
												'Click on Add Button
												Call Fn_Button_Click("Fn_ADS_ProgramAssociation", objProgram, aButtons(iRowData))
											Next
									End If		
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" Sucessfully completed of function Fn_ADS_ProgramAssociation")
									Set objProgram = nothing 
									Fn_ADS_ProgramAssociation = False
									Exit Function
						End If
				Case "AvailableListVerify"						
						bReturn = objProgram.JavaList("OwningProgram").GetROProperty("items count")
						iCnt = 0
						If sProgram<>"" AND Cint(bReturn)<> 0 Then
									'Extract the index of row at which the object exist.
									aColname = split(sProgram, ":",-1,1)
									iCount = Ubound(aColname)
									For iRowData=0 to iCount
										For iCounter=0 to Cint(bReturn)-1
											If Trim(lcase(objProgram.JavaList("OwningProgram").GetItem(iCounter))) = Trim(lcase(aColname(iRowData))) then
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), aColname(iRowData) &" Sucessfully found in Available List")
												iCnt = iCnt + 1
												Exit For
											ElseIf iCounter = Cint(bReturn)-1 Then
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), aColname(iRowData) &" not found in Available List")
											End If
										Next
									Next
									If iCnt=iCount+1 Then
										'Click on Buttons
										If sButtons<>"" Then
												aButtons = split(sButtons, ":",-1,1)
												iCount = Ubound(aButtons)
												For iRowData=0 to iCount
													'Click on Add Button
													Call Fn_Button_Click("Fn_ADS_ProgramAssociation", objProgram, aButtons(iRowData))
												Next
										End If		
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" Sucessfully completed of function Fn_ADS_ProgramAssociation")
										Fn_ADS_ProgramAssociation = TRUE
										Set objProgram = nothing 
										Exit Function
									Else
										'Click on Buttons
										If sButtons<>"" Then
												aButtons = split(sButtons, ":",-1,1)
												iCount = Ubound(aButtons)
												For iRowData=0 to iCount
													'Click on Add Button
													Call Fn_Button_Click("Fn_ADS_ProgramAssociation", objProgram, aButtons(iRowData))
												Next
										End If		
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" Sucessfully completed of function Fn_ADS_ProgramAssociation")
									Fn_ADS_ProgramAssociation = False
									Set objProgram = nothing
									Exit Function
									End If
						Else
								'Click on Buttons
								If sButtons<>"" Then
										aButtons = split(sButtons, ":",-1,1)
										iCount = Ubound(aButtons)
										For iRowData=0 to iCount
											'Click on Add Button
											Call Fn_Button_Click("Fn_ADS_ProgramAssociation", objProgram, aButtons(iRowData))
										Next
								End If		
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" Sucessfully completed of function Fn_ADS_ProgramAssociation")
							Fn_ADS_ProgramAssociation = False
							Set objProgram = nothing
							Exit Function
						End If
				Case "SelectedListVerify"						
						bReturn = objProgram.JavaList("SelectedProgram").GetROProperty("items count")
						iCnt = 0
						If sProgram<>"" AND Cint(bReturn)<> 0 Then
									'Extract the index of row at which the object exist.
									aColname = split(sProgram, ":",-1,1)
									iCount = Ubound(aColname)
									For iRowData=0 to iCount
										For iCounter=0 to Cint(bReturn)-1
											If Trim(lcase(objProgram.JavaList("SelectedProgram").GetItem(iCounter))) = Trim(lcase(aColname(iRowData))) then
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), aColname(iRowData) &" Sucessfully found in Selected List")
												iCnt = iCnt + 1
												Exit For
											ElseIf iCounter = Cint(bReturn)-1 Then
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), aColname(iRowData) &" not found in Selected List")
											End If
										Next
									Next
									If iCnt=iCount+1 Then
										'Click on Buttons
										If sButtons<>"" Then
												aButtons = split(sButtons, ":",-1,1)
												iCount = Ubound(aButtons)
												For iRowData=0 to iCount
													'Click on Add Button
													Call Fn_Button_Click("Fn_ADS_ProgramAssociation", objProgram, aButtons(iRowData))
												Next
										End If		
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" Sucessfully completed of function Fn_ADS_ProgramAssociation")
										Fn_ADS_ProgramAssociation = TRUE
										Set objProgram = nothing 
										Exit Function
									Else
										'Click on Buttons
										If sButtons<>"" Then
												aButtons = split(sButtons, ":",-1,1)
												iCount = Ubound(aButtons)
												For iRowData=0 to iCount
													'Click on Add Button
													Call Fn_Button_Click("Fn_ADS_ProgramAssociation", objProgram, aButtons(iRowData))
												Next
										End If		
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" Sucessfully completed of function Fn_ADS_ProgramAssociation")
									Fn_ADS_ProgramAssociation = False
									Set objProgram = nothing
									Exit Function
									End If
						Else
								'Click on Buttons
								If sButtons<>"" Then
										aButtons = split(sButtons, ":",-1,1)
										iCount = Ubound(aButtons)
										For iRowData=0 to iCount
											'Click on Add Button
											Call Fn_Button_Click("Fn_ADS_ProgramAssociation", objProgram, aButtons(iRowData))
										Next
								End If		
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" Sucessfully completed of function Fn_ADS_ProgramAssociation")
							Fn_ADS_ProgramAssociation = False
							Set objProgram = nothing
							Exit Function
						End If
			Case Else						
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_ADS_ProgramAssociation function failed")
						Fn_ADS_ProgramAssociation = FALSE
						Exit Function						
		End Select
		'Select Parent To Assign
		If sParent<>"" Then
			Call Fn_List_Select("Fn_ADS_ProgramAssociation", objProgram, "ParentToAssign",sParent)
		End If
		'Select Parent To Assign
		If sChild<>"" Then
			Call Fn_List_Select("Fn_ADS_ProgramAssociation", objProgram, "ChildrenLevels",sChild)
		End If
		'Click on Buttons
		If sButtons<>"" Then
				aButtons = split(sButtons, ":",-1,1)
				iCount = Ubound(aButtons)
				For iRowData=0 to iCount
					'Click on Add Button
					Call Fn_Button_Click("Fn_ADS_ProgramAssociation", objProgram, aButtons(iRowData))
                    Call Fn_ReadyStatusSync(2)
				Next
		End If		
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" Sucessfully completed of function Fn_ADS_ProgramAssociation")
	Fn_ADS_ProgramAssociation = TRUE
    Set objProgram = nothing 	
End Function
'#######################################################################################
'###     FUNCTION NAME   :   Fn_ADS_MeasureTextLength(sAction, sObject, sButtons)
'###
'###    DESCRIPTION     :   Get the length of Text written inside any Java control.
'###
'###    Return Value  	:   	Length of the TextString.
'###
'###    HISTORY         :   		AUTHOR              	DATE        		VERSION
'###
'###    CREATED BY      :     Ketan Raje			  12/10/2010   			1.0
'###
'###    REVIWED BY      :		Harshal 				12/10/2010			1.0
'###
'###    MODIFIED BY     :
'###    EXAMPLE         :  	Fn_ADS_MeasureTextLength("JavaEdit|Name", JavaWindow("DefaultWindow").JavaWindow("New Change"), "Cancel") 
'#############################################################################################
Public Function Fn_ADS_MeasureTextLength(sAction, sObject, sButtons)
	GBL_FAILED_FUNCTION_NAME="Fn_ADS_MeasureTextLength"
Dim objDialog, sText, aButtons, iCount, iRowData
	Set objDialog = sObject
	aAction = split(sAction, "|",-1, 1)
	Select Case aAction(0)
	Case "JavaEdit"
				sText = objDialog.JavaEdit(aAction(1)).GetROProperty("value")
				Fn_ADS_MeasureTextLength = Len(sText)
	End Select
	'Click on Buttons
	If sButtons<>"" Then
			aButtons = split(sButtons, ":",-1,1)
			iCount = Ubound(aButtons)
			For iRowData=0 to iCount
				'Click on Add Button
				Call Fn_Button_Click("Fn_ADS_MeasureTextLength", objDialog, aButtons(iRowData))
                Call Fn_ReadyStatusSync(2)
			Next
	End If		
Set objDialog = Nothing
End Function
'#########################################################################################################
'###    FUNCTION NAME   :   Fn_ADS_PartDetailCreate
'###
'###    DESCRIPTION     :  Create New Part
'###
'###    Return      			:  	The PartDtl String formatted (SPartID-SPartRevID)
'###
'###	 HISTORY        	:   AUTHOR                 DATE     	   VERSION
'###
'###    CREATED BY     	:   Ketan					13/10/2010         1.0
'###
'###    REVIWED BY     	:   Harshal				  13/10/2010         1.0
'###
'###    MODIFIED BY   	:  Shrikant N             07-06-2012         Changes : modified New Part Hierarchy.
'###
'###    EXAMPLE         : 	Call Fn_ADS_PartDetailCreate("Part", "", "None:None:NewPart:Testing:None", "INFO:None:None:None:None:None:None", "", "", "", "", "", "", "", "", "Finish:Close")
'#############################################################################################################
Function Fn_ADS_PartDetailCreate(sSelectType, bConfItem, sPartInfo, sAddPartInfo, sAddPartRevInfo, sAttachFileInfo, sWorkFlowInfo, sIdentifierBasicInfo, sAddIDInfo, sAddRevInfo, sAssignProj, sDefineOptions, sButtons)
	GBL_FAILED_FUNCTION_NAME="Fn_ADS_PartDetailCreate"
     on error Resume Next
	Dim ObjStaticText, objDialogNewPart, aPartInfo, sPartId, sRevId, aAddPartInfo, aProgramName, iRowData, iCount, iCounter, sOptions, aButtons, objSelectType, intNoOfObjects
	Dim sNewItemMenu
		'Check the existence of "New Item " window
'	Set objDialogNewPart=Fn_UI_ObjectCreate("Fn_ADS_PartDetailCreate",Window("ADSWindow").JavaDialog("New Part"))
'	Set objDialogNewPart=Window("ADSWindow").JavaDialog("New Part")
	Set objDialogNewPart = Fn_ADS_SISW_GetObject("New Part")
	sNewItemMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("RAC_Menu"),"FileNewPart")
	'Select menu [File -> New -> Item...]
	'If Fn_UI_ObjectExist("Fn_ADS_PartDetailCreate", objDialogNewPart)=False Then
	If objDialogNewPart.Exist(SISW_MIN_TIMEOUT)=False Then
        Call Fn_MenuOperation("Select",sNewItemMenu)
		Call Fn_ReadyStatusSync(2)
	End If
	
	'Creating Object of links on the left side of the window
	Set ObjStaticText = objDialogNewPart.JavaStaticText("Steps")
	'Select Item Type
	If sSelectType <> "" Then
		Call Fn_UI_JavaList_ExtendSelect("Fn_ADS_PartDetailCreate", objDialogNewPart,"SelectedProgram",sSelectType)
	End If
	''checked Configuration item or not
	'If bConfItem <> "" Then
	'	Call Fn_CheckBox_Set("Fn_ADS_PartDetailCreate", objDialogNewPart,"Configuration Item",bConfItem)
	'End If
	'Click on "Next" button
	 Call Fn_Button_Click("Fn_ADS_PartDetailCreate", objDialogNewPart,"Next")
	'Enter Item Information
		If sPartInfo<>"" Then
				aPartInfo = split(sPartInfo, ":",-1,1)
				'click on assign button
				If  aPartInfo(0) = "None" or aPartInfo(1) = "None" Then	
					Call Fn_Button_Click("Fn_ADS_PartDetailCreate", objDialogNewPart,"Assign")
				Else
					'Set Item ID
					Call Fn_Edit_Box("Fn_ADS_PartDetailCreate",objDialogNewPart,"ID", aPartInfo(0))
					'Set Revision ID
					Call Fn_Edit_Box("Fn_ADS_PartDetailCreate",objDialogNewPart,"RevisionID", aPartInfo(1))
				End If
				'Extract Creation data
				sPartId = Fn_Edit_Box_GetValue("Fn_ADS_PartDetailCreate", objDialogNewPart,"ID")
				sRevId = Fn_Edit_Box_GetValue("Fn_ADS_PartDetailCreate", objDialogNewPart,"RevisionID")
				'Set Item name
				 Call Fn_Edit_Box("Fn_ADS_PartDetailCreate", objDialogNewPart,"Name",aPartInfo(2))
				'Set description
				If aPartInfo(3)<>"None" Then
					Call Fn_Edit_Box("Fn_ADS_PartDetailCreate", objDialogNewPart,"Description",aPartInfo(3))
				End If
				'Set UOM
				If aPartInfo(4) <> "None" Then
				  Call Fn_Edit_Box("Fn_ADS_PartDetailCreate", objDialogNewPart,"Unit of Measure",aPartInfo(4))
				End If 
		End If
		'Entering Additional Item Information
			If sAddPartInfo<>"" Then				
				' Click on Next Button
				ObjStaticText.SetTOProperty "label", "Enter Additional Part Information"
				ObjStaticText.Click 1, 1
				aAddPartInfo = split(sAddPartInfo, ":",-1,1)	
					If aAddPartInfo(0) <>"None" Then
						 Call Fn_Edit_Box("Fn_ADS_PartDetailCreate",objDialogNewPart,"PartCategory", aAddPartInfo(0))
					End If
					ObjStaticText.Click 1, 1
					If aAddPartInfo(1) <>"None" Then
						 Call Fn_Edit_Box("Fn_ADS_PartDetailCreate",objDialogNewPart,"SourceDocCategory", aAddPartInfo(1))
					End If
					If aAddPartInfo(2) <>"None" Then
						 Call Fn_Edit_Box("Fn_ADS_PartDetailCreate",objDialogNewPart,"SourceDocID", aAddPartInfo(2))
					End If
					If aAddPartInfo(3) <>"None" Then
						 Call Fn_Button_Click("Fn_ADS_PartDetailCreate", objDialogNewPart,"SrcDocOrg")
								Set objSelectType=description.Create()
								objSelectType("Class Name").value = "JavaStaticText"
								Set  intNoOfObjects = objNewForm.ChildObjects(objSelectType)
								for  iCounter = 0 to intNoOfObjects.count-1
								   If  intNoOfObjects(iCounter).getROProperty("label") = aAddPartInfo(3) Then
										intNoOfObjects(iCounter).Click 1,1
										wait 3
										Exit for
								   End If
								Next
					End If
					If aAddPartInfo(4) <>"None" Then
						 Call Fn_Edit_Box("Fn_ADS_PartDetailCreate",objDialogNewPart,"SourceDocRevision", aAddPartInfo(4))
					End If
					If aAddPartInfo(5) <>"None" Then
						 Call Fn_Edit_Box("Fn_ADS_PartDetailCreate",objDialogNewPart,"SrcTechDocCategory", aAddPartInfo(5))
					End If
					'This Design Required part is yet to be coded
					If aAddPartInfo(6) <>"None" Then
						 
					End If
			End If
			'Enter Additional Item Revision Information
			If sAddPartRevInfo<>"" Then
				' Click on Next Button		
				ObjStaticText.SetTOProperty "label", "Enter Additional Part Revision Information"
				ObjStaticText.Click 1, 1
				aAddItemRevInfo = split(sAddPartRevInfo, ":",-1,1)	
					If aAddItemRevInfo(0) <>"None" Then
						 'Call Fn_Edit_Box("Fn_ADS_PartDetailCreate",objDialogNewPart,"ItemProjectID", aAddItemRevInfo(0))
					End If
					If aAddItemRevInfo(1) <>"None" Then
						 'Call Fn_Edit_Box("Fn_ADS_PartDetailCreate",objDialogNewPart,"ItemRevPreviousID", aAddItemRevInfo(1))
					End If
					If aAddItemRevInfo(2) <>"None" Then
						 'Call Fn_Edit_Box("Fn_ADS_PartDetailCreate",objDialogNewPart,"ItemSerialNumber", aAddItemRevInfo(2))
					End If
					If aAddItemRevInfo(3) <>"None" Then
						 'Call Fn_Edit_Box("Fn_ADS_PartDetailCreate",objDialogNewPart,"ItemComment", aAddItemRevInfo(3))
					End If
					If aAddItemRevInfo(4) <>"None" Then
						 'Call Fn_Edit_Box("Fn_ADS_PartDetailCreate",objDialogNewPart,"ItemUserData1", aAddItemRevInfo(4))
					End If
					If aAddItemRevInfo(5) <>"None" Then
						 'Call Fn_Edit_Box("Fn_ADS_PartDetailCreate",objDialogNewPart,"ItemUserData2", aAddItemRevInfo(5))
					End If
					If aAddItemRevInfo(6) <>"None" Then
						 'Call Fn_Edit_Box("Fn_ADS_PartDetailCreate",objDialogNewPart,"ItemUserData3", aAddItemRevInfo(6))
					End If
			End If
			'Enter Identifier Basic Information
			If sIdentifierBasicInfo<>"" Then
				' Click on Next Button		
				ObjStaticText.SetTOProperty "label", "Enter Identifier Basic Information"
				ObjStaticText.Click 1, 1
				Wait(2)
				If sIdentifierBasicInfo(0)<>"None" Then
					'Set TOProperties of Dialog Box
					JavaWindow("DefaultWindow").JavaWindow("Shell").JavaWindow("No Assign Privilege").SetTOProperty "title","New Part ..."
					'Set the "Don't show this message" Status
					Call Fn_CheckBox_Set("Fn_ADS_PartDetailCreate", JavaWindow("DefaultWindow").JavaWindow("Shell").JavaWindow("No Assign Privilege"), "Don't show this message", sIdentifierBasicInfo(0))
					'Click on OK button
					Call Fn_Button_Click("Fn_ADS_PartDetailCreate", JavaWindow("DefaultWindow").JavaWindow("Shell").JavaWindow("No Assign Privilege"), "OK")
				End If
					If sIdentifierBasicInfo(1) <>"None" Then
						 'Call Fn_Edit_Box("Fn_ADS_PartDetailCreate",objDialogNewPart,"ItemProjectID", sIdentifierBasicInfo(1))
					End If
					If sIdentifierBasicInfo(2) <>"None" Then
						 'Call Fn_Edit_Box("Fn_ADS_PartDetailCreate",objDialogNewPart,"ItemRevPreviousID", sIdentifierBasicInfo(2))
					End If
					If sIdentifierBasicInfo(3) <>"None" Then
						 'Call Fn_Edit_Box("Fn_ADS_PartDetailCreate",objDialogNewPart,"ItemSerialNumber", sIdentifierBasicInfo(3))
					End If
					If sIdentifierBasicInfo(4) <>"None" Then
						 'Call Fn_Edit_Box("Fn_ADS_PartDetailCreate",objDialogNewPart,"ItemComment", sIdentifierBasicInfo(4))
					End If
					If sIdentifierBasicInfo(5) <>"None" Then
						 'Call Fn_Edit_Box("Fn_ADS_PartDetailCreate",objDialogNewPart,"ItemUserData1", sIdentifierBasicInfo(5))
					End If
					If sIdentifierBasicInfo(6) <>"None" Then
						 'Call Fn_Edit_Box("Fn_ADS_PartDetailCreate",objDialogNewPart,"ItemUserData2", sIdentifierBasicInfo(6))
					End If
					If sIdentifierBasicInfo(7) <>"None" Then
						 'Call Fn_Edit_Box("Fn_ADS_PartDetailCreate",objDialogNewPart,"ItemUserData3", sIdentifierBasicInfo(7))
					End If
			End If
			'Assign to Project
			If sAssignProj<>"" Then
				' Click on Next Button
				ObjStaticText.SetTOProperty "label", "Assign to Program"
				ObjStaticText.Click 1, 1
				Call Fn_ReadyStatusSync(3)
				bReturn = objDialogNewPart.JavaList("ProgramForSelect").GetROProperty("items count")
				'Extract the index of row at which the object exist.
				aProgramName = split(sAssignProj, ":",-1,1)
				iCount = Ubound(aProgramName)
				For iRowData=0 to iCount
					For iCounter=0 to bReturn-1
						If Trim(lcase(objDialogNewPart.JavaList("ProgramForSelect").GetItem(iCounter))) = Trim(lcase(aProgramName(iRowData))) then
							objDialogNewPart.JavaList("ProgramForSelect").Select aProgramName(iRowData)
							'Click on Remove Button
							Call Fn_Button_Click("Fn_ADS_PartDetailCreate", objDialogNewPart, "AddProject")											
							Exit For 
						End If
					Next
				Next
			End If
			If sDefineOptions<>"" Then
				' Click on Next Button
					ObjStaticText.SetTOProperty "label", "Define Options"
					ObjStaticText.Click 1, 1	
						sOptions = split(sDefineOptions, ":",-1,1)					
							If sOptions(0) <> "" Then
								'Call Fn_CheckBox_Set("Fn_ADS_PartDetailCreate" ,objDialogNewPart,"ShowAsNwRt", sOptions(0)) 
							End If
							If sOptions(1) <> "" Then
								'Call Fn_CheckBox_Set("Fn_ADS_PartDetailCreate" ,objDialogNewPart,"UsItIdentifierAs", sOptions(1)) 
							End If
							If sOptions(2) <> "" Then
								'Call Fn_CheckBox_Set("Fn_ADS_PartDetailCreate" ,objDialogNewPart,"UsRevIdentifier", sOptions(2)) 
							End If
							If sOptions(3) <> "" Then
								'Call Fn_CheckBox_Set("Fn_ADS_PartDetailCreate" ,objDialogNewPart,"ChkOutItmRevOnCr", sOptions(3)) 
							End If
			End If
			objDialogNewPart.JavaButton("Next").WaitProperty "enabled", 1, 20000
			'Click on Buttons
			If sButtons<>"" Then
					aButtons = split(sButtons, ":",-1,1)
					iCounter = Ubound(aButtons)
					For iCount=0 to iCounter
						'Click on Add Button
						Call Fn_Button_Click("Fn_ADS_PartDetailCreate", objDialogNewPart, aButtons(iCount))
						Call Fn_ReadyStatusSync(1)
					Next
			End If
			'Function Returns Item ID and True
			Fn_ADS_PartDetailCreate = "'"&sPartId & "-" & sRevId
			'Write Log
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Completed the function Fn_ADS_PartDetailCreate")
	Set ObjStaticText = Nothing
	Set objDialogNewPart = Nothing
	Set objSelectType = Nothing
	Set intNoOfObjects = Nothing
 End Function
'#########################################################################################################
'###    FUNCTION NAME   :   Fn_ADS_SessionInfoVerify
'###
'###    DESCRIPTION     :  Verify Session information
'###
'###    Return      			:  	True/False
'###
'###	 HISTORY        	:   AUTHOR                 DATE     	   VERSION 		BUILD
'###
'###    CREATED BY     	:   Harshal					14/10/2010         1.0				916a
'###
'###    REVIWED BY     	:   Harshal				  14/10/2010         1.0				916a
'###
'###    MODIFIED BY   	:  
'###
'###    EXAMPLE         : 	Call Fn_ADS_SessionInfoVerify("Training")
'#############################################################################################################
Public Function Fn_ADS_SessionInfoVerify(sInfo)
	GBL_FAILED_FUNCTION_NAME="Fn_ADS_SessionInfoVerify"
	Dim objDialog
	Dim sAppInfo
	Set objDialog = JavaWindow("DefaultWindow").JavaStaticText("SessionInfo")
	sAppInfo = objDialog.GetROProperty("label")
	If instr(1,sAppInfo,sInfo)<>0 Then
		Fn_ADS_SessionInfoVerify = True
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Completed the function Fn_ADS_SessionInfoVerify")
	Else
		Fn_ADS_SessionInfoVerify = False
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail to Complete the function Fn_ADS_SessionInfoVerify")
	End If
	Set objDialog = Nothing
End Function
'#########################################################################################################
'###    FUNCTION NAME   :   Fn_ADS_SummaryTabOperation
'###
'###    DESCRIPTION     :  operation related to Summary Tab
'###
'###    Return      			:  	True/False
'###
'###	 HISTORY        	:   AUTHOR                 DATE     	   VERSION 		BUILD
'###
'###    CREATED BY     	:   Harshal					14/10/2010         1.0				916a
'###
'###    REVIWED BY     	:   Harshal				  14/10/2010         1.0				916a
'###
'###    MODIFIED BY   	:  
'###
'###    EXAMPLE         : 	'MSGBOX Fn_ADS_SummaryTabOperation("Verify","Closure:Disposition:Maturity","Open:None:Elaborating")
'###								'MSGBOX Fn_ADS_SummaryTabOperation("Verify","Closure","Open")
'###								'Msgbox Fn_ADS_SummaryTabOperation("GetLicenseSummary", "GroupName:ObjectId", "")	Added By Ketan on 04/10/2011.
'#############################################################################################################
Function Fn_ADS_SummaryTabOperation(sAction,sProperty,sValue)
	GBL_FAILED_FUNCTION_NAME="Fn_ADS_SummaryTabOperation"
	Dim sAppValue,iCount,aProperty,aValue,jcount, aLicSummary
	Dim objJavaTable, objLicElement, iNoOfLicObjs, objTab
	Select Case sAction
		Case "Verify"
			If sProperty<>"" Then
            	aProperty = Split(sProperty,":")
				aValue = Split(sValue,":")
				For iCount = 0 to ubound(aProperty)
					Select Case aProperty(iCount)
						Case "Closure","Disposition","Maturity"
							sAppValue = JavaWindow("ADS-TeamCenter").JavaLink(aProperty(iCount)).GetROProperty("value")
							If sAppValue <> aValue(iCount) Then
								Fn_ADS_SummaryTabOperation = False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail to Complete the function Fn_ADS_SummaryTabOperation")
								Exit Function
							Else
								Fn_ADS_SummaryTabOperation = True
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Completed the function Fn_ADS_SummaryTabOperation")
							End If
						Case Else
							Fn_ADS_SummaryTabOperation = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail to Complete the function Fn_ADS_SummaryTabOperation")
								Exit Function
						End Select
				Next
			End If
			
	Case "AttachesFile_AddNew","ActionItemTab_AddNew"
        	Dim arrToolbarButtons, objToolBarButton
			Fn_ADS_SummaryTabOperation  =False
			If sAction =  "AttachesFile_AddNew" Then
				JavaWindow("MyTeamcenter").JavaTab("SummaryTab").Select "Related Datasets"
				set objTab = JavaWindow("MyTeamcenter").JavaToolbar("RelatedDatasetsTableToolBar")
			else
				JavaWindow("MyTeamcenter").JavaTab("SummaryTab").Select "Action Items"
				set objTab = JavaWindow("DefaultWindow").JavaToolbar("ActionItemsToolbar")
			End If
			If objTab.Exist Then
				Set arrToolbarButtons = Fn_SISW_UI_Object_GetChildObjects("Fn_ADS_SummaryTabOperation", objTab, "Class Name~label", "JavaButton~Add New")
				If TypeName(arrToolbarButtons) <> "Nothing"  Then
					Set objToolBarButton =  arrToolbarButtons(0)
					If Environment.Value("ProductName") = sUFTProductName Then
       					objToolBarButton.Click 1,1
						Fn_ADS_SummaryTabOperation = True
					Else					
						Fn_ADS_SummaryTabOperation = Fn_SISW_UI_JavaButton_Operations("Fn_ADS_SummaryTabOperation", "Click", objToolBarButton, "")
					End If
					
				Else
					Set arrToolbarButtons = Fn_SISW_UI_Object_GetChildObjects("Fn_ADS_SummaryTabOperation",objTab, "Class Name~label", "JavaButton~Add New")
					If TypeName(arrToolbarButtons) <> "Nothing"  Then
						Set objToolBarButton =  arrToolbarButtons(0)
						Fn_ADS_SummaryTabOperation  = Fn_SISW_UI_JavaButton_Operations("Fn_ADS_SummaryTabOperation", "Click", objToolBarButton, "")
					End If
				End If
			End If
		
	Case "Schedule_AddNewCntEventSchedule"
		If sAction="Schedule_AddNewCntEventSchedule"  Then
			JavaWindow("MyTeamcenter").JavaTab("SummaryTab").Select "Schedule"
			If JavaWindow("ADS-TeamCenter").JavaButton("Add Contract Event Schedule").Exist Then
				JavaWindow("ADS-TeamCenter").JavaButton("Add Contract Event Schedule").Click micLeftBtn
				Fn_ADS_SummaryTabOperation  = True
			Else
				Fn_ADS_SummaryTabOperation  = False
			End If
		End If
	Case "SubmittalsVerify","CorrespondencesVerify","Contracts","DIDVerify","DRIVerify","SubmittalScheduleVerify"
		If sAction="CorrespondencesVerify"  Then
			JavaWindow("MyTeamcenter").JavaTab("SummaryTab").Select "Correspondences"
		Elseif sAction = "SubmittalsVerify"  Then
			JavaWindow("MyTeamcenter").JavaTab("SummaryTab").Select "Submittals"
		Elseif sAction = "SubmittalScheduleVerify"  Then
			JavaWindow("MyTeamcenter").JavaTab("SummaryTab").Select "Submittal Schedule"
		Elseif sAction = "Contracts"  Then
			JavaWindow("MyTeamcenter").JavaTab("SummaryTab").Select "Contracts"
		Elseif sAction = "DIDVerify"  Then
			JavaWindow("MyTeamcenter").JavaTab("SummaryTab").Select "Data Item Descriptions"
		Elseif sAction = "DRIVerify"  Then
			JavaWindow("MyTeamcenter").JavaTab("SummaryTab").Select "Data Requirement Items"
		End If
		JavaWindow("MyTeamcenter").JavaButton("ItmAttachFileTblView").SetTOProperty "Index","0"
		Wait(2)
        JavaWindow("MyTeamcenter").JavaButton("ItmAttachFileTblView").Click micLeftBtn
		If sProperty<>"" Then
            	aProperty = Split(sProperty,":")
				aValue = Split(sValue,":")
				For iCount = 0 to ubound(aProperty)
					Select Case aProperty(iCount)
					Case"Object"
							If JavaWindow("MyTeamcenter").JavaTable("ScheduleTable").Exist(2) Then
								Set objJavaTable = JavaWindow("MyTeamcenter").JavaTable("ScheduleTable")
							ElseIf JavaWindow("MyTeamcenter").JavaTable("SummaryTabTable").Exist(2) Then
								Set objJavaTable = JavaWindow("MyTeamcenter").JavaTable("SummaryTabTable")
							Else 
								Fn_ADS_SummaryTabOperation  = False
								Exit Function
							End If
							For jcount = 0 to cint(objJavaTable.GetROProperty("rows")) - 1
	'							sAppValue = objJavaTable.GetCellData(jcount,0)
								sAppValue = Fn_SISW_ADS_JavaTable_GetCellData(objJavaTable,jcount , "Object")
								If trim(lcase(sAppValue)) = trim(lcase(aValue(iCount))) Then
									Fn_ADS_SummaryTabOperation  = True
									Exit for
								End If
								Fn_ADS_SummaryTabOperation  = False
							Next
					Case Else
							Fn_ADS_SummaryTabOperation = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail to Complete the function Fn_ADS_SummaryTabOperation")
							Exit Function
					End Select
				Next
		End if
	Case "Submittal Schedule"
				JavaWindow("MyTeamcenter").JavaTab("SummaryTab").Select "Submittal Schedule"
				Fn_ADS_SummaryTabOperation  = True
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully selected Submittal Schedule Tab")
	Case "GetLicenseSummary"
				If sProperty <> "" Then
						aProperty = Split(sProperty,":")
						ReDim aLicSummary(Ubound(aProperty))
						For jcount = 0 to Ubound(aProperty)
								'Locate the given property Label.
								Set objLicElement = Description.Create()
								objLicElement("Class Name").value = "JavaStaticText"
								Set iNoOfLicObjs = JavaWindow("DefaultWindow").JavaObject("LicenseSummary").ChildObjects(objLicElement)
								For iCount = 0 to iNoOfLicObjs.count-1
									If iNoOfLicObjs(iCount).getROProperty("label") = aProperty(jcount)&":" Then
										Exit For
									End If															
								Next
								'Extract values from EditBoxex.		
								objLicElement("Class Name").value = "JavaEdit"
								Set iNoOfLicObjs = JavaWindow("DefaultWindow").JavaObject("LicenseSummary").ChildObjects(objLicElement)
								aLicSummary(jcount) = iNoOfLicObjs(iCount).getROProperty("value")
						Next
					
						Set objLicElement = Nothing
						Set iNoOfLicObjs = Nothing
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully extracted values for following properties "&sProperty)
						Fn_ADS_SummaryTabOperation = aLicSummary
				End If		
	Case Else
		Fn_ADS_SummaryTabOperation = False
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail to Complete the function Fn_ADS_SummaryTabOperation")
		Exit Function
	End Select
	Set objJavaTable = nothing
End Function
'#######################################################################################
'###     FUNCTION NAME   :   Fn_ADS_DialogHandle(sObject,sTitle,sButtons)
'###
'###    DESCRIPTION     :   Handle a window in ADS
'###
'###    PARAMETERS      :   sObject, sTitle, sButtons
'###
'###    Return Value  	:   	True/False 
'###
'###    HISTORY         :   		AUTHOR              	DATE        		VERSION
'###
'###    CREATED BY      :     Shreyas					14/10/2010   			1.0
'###
'###    REVIWED BY      :	  Ketan Raje			  14/10/2010
'###
'###    MODIFIED BY     :
'###    EXAMPLE         :   msgbox Fn_ADS_DialogHandle(JavaWindow("ADS-TeamCenter").JavaWindow("ADSDialog"),"Enter the values for Properties on Relation","Finish")
'###								msgbox Fn_ADS_DialogHandle(JavaWindow("ADS-TeamCenter").JavaWindow("ADSDialog"),"Paste...","Close")
'#############################################################################################
Public Function Fn_ADS_DialogHandle(sObject,sTitle,sButtons)

	Dim dicErrorInfo
	 Set dicErrorInfo = CreateObject("Scripting.Dictionary")
	 With dicErrorInfo 
	  .Add "Object" , sObject
	  .Add "Title", sTitle
	  .Add "Button", sButtons
	  .Add "Action", "DialogHandle" 	  
	 End with
   Fn_ADS_DialogHandle = Fn_SISW_ADS_ErrorVerify(dicErrorInfo)
   Set dicErrorInfo = Nothing

End Function
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$   	FUNCTION NAME   	:  Fn_ADS_ObjectPropertyChkOut()
'$$$$
'$$$$  		DESCRIPTION     		: 	CheckOut Operation for the Object/Item .
'$$$$
'$$$$		FUNCTION CALLS		:   Fn_WriteLogFile(), Fn_Button_Click(), Fn_MenuOperation()
'$$$$
'$$$$    	PRE-REQUISITES  	:  Item/Object  to be Selected.
'$$$$
'$$$$	 	HISTORY         			:   		AUTHOR                 DATE   	     VERSION			BUILD
'$$$$
'$$$$      CREATED BY      		:     			Pranav					14/10/10   			   1.0					  916a
'$$$$
'$$$$      REVIWED BY      		:   			Harshal					14/10/10   			   1.0					
'$$$$
'$$$$      EXAMPLE         			:   
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Function Fn_ADS_ObjectPropertyChkOut()
	GBL_FAILED_FUNCTION_NAME="Fn_ADS_ObjectPropertyChkOut"
	Dim ObjChkOut, sGetTxt

	Dim StrTitle
	StrTitle=JavaWindow("DefaultWindow").GetROProperty("title")

		'Menu Operation......
		Call Fn_MenuOperation("Select","Tools:Check-In/Out:Check Out...")
        If Instr(1,StrTitle,"My Teamcenter") or Instr(1,StrTitle,"Change Manager") Then
			Set ObjChkOut =Fn_UI_ObjectCreate( "Fn_ADS_ObjectPropertyChkOut", JavaWindow("DefaultWindow").JavaWindow("Check-Out"))
		Else
			Set ObjChkOut = Fn_UI_ObjectCreate( "Fn_ADS_ObjectPropertyChkOut",JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Check-Out"))
		End if
			If ObjChkOut.Exist Then
					'Click on YES
					Call Fn_Button_Click("Fn_ObjectCheckOut", ObjChkOut ,"Yes")
					Wait(15)
					If ObjChkOut.JavaObject("CheckOutErrorMessage").Exist(2) Then
                		sGetTxt = ObjChkOut.JavaButton("CheckOutErrorBtn").GetROProperty ("tool_tip_text")
						'Return the text
						Fn_ADS_ObjectPropertyChkOut = sGetTxt
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Check-Out Item of Function Fn_ADS_ObjectPropertyChkOut")
						Call Fn_Button_Click("Fn_ObjectCheckOut", ObjChkOut ,"OK")
						Exit Function
					ElseIf ObjChkOut.JavaStaticText("CheckOutErrorMessage").Exist(2) Then
						Fn_ADS_ObjectPropertyChkOut=ObjChkOut.JavaStaticText("CheckOutErrorMessage").GetROProperty ("label")
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Check-Out Item of Function Fn_ADS_ObjectPropertyChkOut")
						Call Fn_Button_Click("Fn_ObjectCheckOut", ObjChkOut ,"OK")
						Exit Function
					End If
			End If
		Fn_ADS_ObjectPropertyChkOut = True
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Checked-Out Item of Function Fn_ADS_ObjectPropertyChkOut")
End Function
'-------------------------------------------------------------------Function Used to Create Problem Report,Change Notice,Change Request,Deviation Request----------------------------------------------------------
'Function Name  : Fn_ADS_ChangeOperationsExtn
'Description    : Function Used to Create Problem Report,Change Notice,Change Request,Deviation Request
'Return Value     :  True Or False
'Examples    :  ...................................................................................................................Cases for Set......................................................................................................................
'Msgbox Fn_ADS_ChangeOperationsExtn("SetCNEdit", "", "Change Notice", "ECN-111234", "A", "Test", "Testing", "DCN", "", "", "Finish:Cancel")
'Msgbox Fn_ADS_ChangeOperationsExtn("SetCREdit", "", "Change Request", "ECR-111234", "A", "Test", "Testing", "DCN", "", "", "Finish:Cancel")
'Msgbox Fn_ADS_ChangeOperationsExtn("SetDREdit", "", "Deviation Request", "EDR-111234", "A", "Test", "Testing", "DCN", "", "", "Finish:Cancel")
'Msgbox Fn_ADS_ChangeOperationsExtn("SetPREdit", "", "Problem Report", "01-9998", "A", "Test", "Testing", "", "", "", "Finish:Cancel")
'Msgbox Fn_ADS_ChangeOperationsExtn("SetCNDropDown", "", "Change Notice", "01-9998", "A", "Test", "Testing", "DCN", "", "", "Finish:Cancel")
'Msgbox Fn_ADS_ChangeOperationsExtn("SetDRDropDown", "", "Deviation Request", "01-9998", "A", "Test", "Testing", "RFW", "", "", "Finish:Cancel")
'Msgbox Fn_ADS_ChangeOperationsExtn("SetCRDropDown", "", "Change Request", "01-9998", "A", "Test", "Testing", "CA", "", "", "Finish:Cancel")
'..............................................................................................................................................................................................................................................................................................             
'History      :   
'             Developer Name            Date      Rev. No.      Changes Done      Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'             Ketan Raje.              14/10/2010              		                   			Harshal
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'             Koustubh Watwe.          29/02/2012    1.0		Modified code to set ECN No.
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Public Function Fn_ADS_ChangeOperationsExtn(sAction, strFilterText, sNodeName, sChangeID, sChangeRev, sChangeSynopsis, sChangeDesc, sChangeType, sVerifyCombos, dicChangeParam, sButtons)
'   'Declaring Variables
'    Dim ObjChangeWnd, aButtons, iCount, iCounter, objSelectType, intNoOfObjects, sChangeNo, WshShell, iLen, aVerifyCombos
'	Fn_ADS_ChangeOperationsExtn=False
'	  For iCount=0 to 0
'		 JavaWindow("ADS-TeamCenter").JavaWindow("New Change").SetTOProperty "title","New Change"
'		 If JavaWindow("ADS-TeamCenter").JavaWindow("New Change").Exist(5) Then
'		  Exit For
'		 End If
'		 JavaWindow("ADS-TeamCenter").JavaWindow("New Change").SetTOProperty "title","New Change in context"
'		 If JavaWindow("ADS-TeamCenter").JavaWindow("New Change").Exist(5) Then
'		  Exit For
'		 End If
'		 JavaWindow("ADS-TeamCenter").JavaWindow("New Change").SetTOProperty "title","Derive Change"
'		 If JavaWindow("ADS-TeamCenter").JavaWindow("New Change").Exist(5) Then      
'		  Exit For
'		 End If
'	  Next
' Set ObjChangeWnd=Fn_UI_ObjectCreate("Fn_ADS_ChangeOperationsExtn",JavaWindow("ADS-TeamCenter").JavaWindow("New Change"))
' strFilterText = sNodeName
' Select Case sAction
' Case "SetCNEdit","SetCNDropDown"		'Changed the index values for DropDownBtn for ChangeNotice due to design change.
'			   'Set FilterText				
'			   If strFilterText<>"" Then
'				'Call Fn_Edit_Box("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd,"Filter",strFilterText)
'				Call Fn_UI_EditBox_Type("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd,"Filter",strFilterText)
'			   End If
'			   If sNodeName<>"" Then
'				Wait(5)
'				JavaWindow("ADS-TeamCenter").JavaWindow("New Change").JavaTree("ChangeTypeTree").Select "Complete List"
'				JavaWindow("ADS-TeamCenter").JavaWindow("New Change").JavaTree("ChangeTypeTree").Select "Complete List:"+sNodeName
'				'Selecting Node from tree
'				'Call Fn_JavaTree_Select("Fn_ADS_ChangeOperationsExtn", ObjChangeWnd, "ChangeTypeTree","Complete List")
'				'Call Fn_JavaTree_Select("Fn_ADS_ChangeOperationsExtn", ObjChangeWnd, "ChangeTypeTree","Complete List:"+sNodeName)
'				'Clicking on Next button to proceed 
'				Call Fn_Button_Click("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd,"Next")
'			   End If
'			   'Set Change Id
'			   If sChangeID<>"" Then
''			      JavaWindow("ADS-TeamCenter").JavaWindow("New Change").JavaEdit("ECNNo_CN").Activate
'			      JavaWindow("ADS-TeamCenter").JavaWindow("New Change").JavaEdit("ECNNo_CN").Set  sChangeID
''							iLen = len(sChangeID)
''								Set WshShell = CreateObject("WScript.Shell")
''							For iCount = 1 to iLen
''								WshShell.SendKeys mid(sChangeID,iCount,1)
''							Next
''								Set WshShell = Nothing
'			   End If
'			   'Set Change Revision
'			   If sChangeRev<>"" Then
'				Call Fn_UI_EditBox_Type("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd,"Revision_CN",sChangeRev)
'			   End If
'			   'Set Synopsis
'			   If sChangeSynopsis<>"" Then
'				Call Fn_UI_EditBox_Type("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd,"Synopsis_CN",sChangeSynopsis)
'			   End If
'			   'Set Description
'			   If sChangeDesc<>"" Then
'				Call Fn_UI_EditBox_Type("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd,"Description_CN",sChangeDesc)
'			   End If
'			   ObjChangeWnd.Maximize
'				Wait(3)
'			   'Set Change Type
'			   If sAction = "SetCNEdit" Then
'					If sChangeType<>"" Then
'						Call Fn_UI_EditBox_Type("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd,"ChangeType_CN",sChangeType)
'					End If
'			   ElseIf sAction = "SetCNDropDown" Then
'					If sChangeType<>"" Then
''						ObjChangeWnd.JavaButton("DropDownBtn").SetTOProperty "attached text","*"
'						ObjChangeWnd.JavaButton("DropDownBtn").SetTOProperty "index",0
'						Call Fn_Button_Click("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd,"DropDownBtn")
'						Wait(3)
'						iRows = JavaWindow("ADS-TeamCenter").JavaWindow("New Change").JavaWindow("Shell").JavaTable("ChangeTable").GetROProperty("rows")
'						For iCount = 0 to iRows-1
'							If Trim(Lcase(JavaWindow("ADS-TeamCenter").JavaWindow("New Change").JavaWindow("Shell").JavaTable("ChangeTable").GetCellData(iCount,0))) = Trim(Lcase(sChangeType)) Then
'								JavaWindow("ADS-TeamCenter").JavaWindow("New Change").JavaWindow("Shell").JavaTable("ChangeTable").ClickCell iCount,0,"LEFT","NONE"
'								Exit For
'							End If
'						Next
'					End If
'			   End If
'			   'Code for Fields differing in Case of CN,CR,PR,DR is yet to be coded
'			   'Set Change Item Affected.
'				If dicChangeParam("PaperChange")<>"" Then
'					ObjChangeWnd.JavaRadioButton("False").SetTOProperty "Index",0
'					ObjChangeWnd.JavaRadioButton("True").SetTOProperty "Index",0
'					If Trim(Lcase(dicChangeParam("PaperChange"))) = "true" Then
'						Call Fn_UI_JavaRadioButton_SetON("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd, "True")
'					Else
'						Call Fn_UI_JavaRadioButton_SetON("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd, "False")
'					End If
'				End If
'			   'Set Change Class
'			   If sAction = "SetCNEdit" Then
'					If dicChangeParam("ChangeClass")<>"" Then
'						Call Fn_UI_EditBox_Type("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd,"Change Class_CN",dicChangeParam("ChangeClass"))
'					End If
'			   ElseIf sAction = "SetCNDropDown" Then
'					If dicChangeParam("ChangeClass")<>"" Then
''						ObjChangeWnd.JavaButton("DropDownBtn").SetTOProperty "attached text",""
'						ObjChangeWnd.JavaButton("DropDownBtn").SetTOProperty "index",0
'						Call Fn_Button_Click("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd,"DropDownBtn")
'						Wait(3)
'						JavaWindow("ADS-TeamCenter").JavaWindow("New Change").JavaWindow("Shell").SetToProperty "index",0
'						iRows = JavaWindow("ADS-TeamCenter").JavaWindow("New Change").JavaWindow("Shell").JavaTable("ChangeTable").GetROProperty("rows")
'						For iCount = 0 to iRows-1
'							If Trim(Lcase(JavaWindow("ADS-TeamCenter").JavaWindow("New Change").JavaWindow("Shell").JavaTable("ChangeTable").GetCellData(iCount,0))) = Trim(Lcase(dicChangeParam("ChangeClass"))) Then
'								JavaWindow("ADS-TeamCenter").JavaWindow("New Change").JavaWindow("Shell").JavaTable("ChangeTable").ClickCell iCount,0,"LEFT","NONE"
'								Exit For
'							End If
'						Next
'					End If
'			   End If
'			   'Set Change Category
'			   If sAction = "SetCNEdit" Then
'					If dicChangeParam("ChangeCategory")<>"" Then
'						Call Fn_UI_EditBox_Type("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd,"Category_CN",dicChangeParam("ChangeCategory"))
'					End If
'			   ElseIf sAction = "SetCNDropDown" Then
'					If dicChangeParam("ChangeCategory")<>"" Then
'						ObjChangeWnd.JavaButton("DropDownBtn").SetTOProperty "index",1
'						Call Fn_Button_Click("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd,"DropDownBtn")
'						ObjChangeWnd.JavaButton("DropDownBtn").Click 
'						Wait(3)
'						JavaWindow("ADS-TeamCenter").JavaWindow("New Change").JavaWindow("Shell").SetToProperty "index",1
'						iRows = JavaWindow("ADS-TeamCenter").JavaWindow("New Change").JavaWindow("Shell").JavaTable("ChangeTable").GetROProperty("rows")
'						For iCount = 0 to iRows-1
'							If Trim(Lcase(JavaWindow("ADS-TeamCenter").JavaWindow("New Change").JavaWindow("Shell").JavaTable("ChangeTable").GetCellData(iCount,0))) = Trim(Lcase(dicChangeParam("ChangeCategory"))) Then
'								JavaWindow("ADS-TeamCenter").JavaWindow("New Change").JavaWindow("Shell").JavaTable("ChangeTable").ClickCell iCount,0,"LEFT","NONE"
'								Exit For
'							End If
'						Next
'					End If
'			   End If
'  Case "SetCREdit","SetCRDropDown"
'			   'Set FilterText
'			   If strFilterText<>"" Then
'				'Call Fn_Edit_Box("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd,"Filter",strFilterText)
'				Call Fn_UI_EditBox_Type("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd,"Filter",strFilterText)
'			   End If
'			   If sNodeName<>"" Then
'				Wait(1)
'				'Selecting Node from tree
'				Call Fn_JavaTree_Select("Fn_ADS_ChangeOperationsExtn", ObjChangeWnd, "ChangeTypeTree","Complete List")
'				Call Fn_JavaTree_Select("Fn_ADS_ChangeOperationsExtn", ObjChangeWnd, "ChangeTypeTree","Complete List:"+sNodeName)
'				'Clicking on Next button to proceed 
'				Call Fn_Button_Click("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd,"Next")
'			   End If
'			   'Set Change Id
'			   If sChangeID<>"" Then				
'			      JavaWindow("ADS-TeamCenter").JavaWindow("New Change").JavaEdit("ECRNo_CR").Activate
'							iLen = len(sChangeID)
'								Set WshShell = CreateObject("WScript.Shell")
'							For iCount = 1 to iLen
'								WshShell.SendKeys mid(sChangeID,iCount,1)
'							Next
'								Set WshShell = Nothing
'			   End If
'			   'Set Change Revision
'			   If sChangeRev<>"" Then
'				Call Fn_UI_EditBox_Type("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd,"Revision_CR",sChangeRev)
'			   End If
'			   'Set Synopsis
'			   If sChangeSynopsis<>"" Then
'				Call Fn_UI_EditBox_Type("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd,"Synopsis_CR",sChangeSynopsis)
'			   End If
'			   'Set Description
'			   If sChangeDesc<>"" Then
'				Call Fn_UI_EditBox_Type("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd,"Desc_CR",sChangeDesc)
'			   End If
'			   ObjChangeWnd.Maximize
'			   Wait(3)
'			   'Set Change Type
'			   If sAction = "SetCREdit" Then
'					If sChangeType<>"" Then
'						Call Fn_UI_EditBox_Type("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd,"ChangeType_CR",sChangeType)
'					End If
'			   ElseIf sAction = "SetCRDropDown" Then
'					If sChangeType<>"" Then
'						ObjChangeWnd.JavaButton("DropDownBtn").SetTOProperty "index",0
'						Call Fn_Button_Click("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd,"DropDownBtn")
'						Wait(3)
'						iRows = JavaWindow("ADS-TeamCenter").JavaWindow("New Change").JavaWindow("Shell").JavaTable("ChangeTable").GetROProperty("rows")
'						For iCount = 0 to iRows-1
'							If Trim(Lcase(JavaWindow("ADS-TeamCenter").JavaWindow("New Change").JavaWindow("Shell").JavaTable("ChangeTable").GetCellData(iCount,0))) = Trim(Lcase(sChangeType)) Then
'								JavaWindow("ADS-TeamCenter").JavaWindow("New Change").JavaWindow("Shell").JavaTable("ChangeTable").ClickCell iCount,0,"LEFT","NONE"
'								Exit For
'							End If
'						Next
'					End If
'			   End If
'			   'Code for Fields differing in Case of CN,CR,PR,DR is yet to be coded
'			   'Set Change Class
'			   If sAction = "SetCREdit" Then
'					If dicChangeParam("ChangeClass")<>"" Then
'						Call Fn_UI_EditBox_Type("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd,"ChangeClass_CR",dicChangeParam("ChangeClass"))
'					End If
'			   ElseIf sAction = "SetCRDropDown" Then
'					If dicChangeParam("ChangeClass")<>"" Then
'						ObjChangeWnd.JavaButton("DropDownBtn").SetTOProperty "index",1
'						Call Fn_Button_Click("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd,"DropDownBtn")
'						Wait(3)
'						iRows = JavaWindow("ADS-TeamCenter").JavaWindow("New Change").JavaWindow("Shell").JavaTable("ChangeTable").GetROProperty("rows")
'						For iCount = 0 to iRows-1
'							If Trim(Lcase(JavaWindow("ADS-TeamCenter").JavaWindow("New Change").JavaWindow("Shell").JavaTable("ChangeTable").GetCellData(iCount,0))) = Trim(Lcase(dicChangeParam("ChangeClass"))) Then
'								JavaWindow("ADS-TeamCenter").JavaWindow("New Change").JavaWindow("Shell").JavaTable("ChangeTable").ClickCell iCount,0,"LEFT","NONE"
'								Exit For
'							End If
'						Next
'					End If
'			   End If
'			   'Set Change Category
'			   If sAction = "SetCREdit" Then
'					If dicChangeParam("ChangeCategory")<>"" Then
'						Call Fn_UI_EditBox_Type("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd,"ChangeCategory_CR",dicChangeParam("ChangeCategory"))
'					End If
'			   ElseIf sAction = "SetCRDropDown" Then
'					If dicChangeParam("ChangeCategory")<>"" Then
'						ObjChangeWnd.JavaButton("DropDownBtn").SetTOProperty "index",2
'						Call Fn_Button_Click("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd,"DropDownBtn")
'						Wait(3)
'						iRows = JavaWindow("ADS-TeamCenter").JavaWindow("New Change").JavaWindow("Shell").JavaTable("ChangeTable").GetROProperty("rows")
'						For iCount = 0 to iRows-1
'							If Trim(Lcase(JavaWindow("ADS-TeamCenter").JavaWindow("New Change").JavaWindow("Shell").JavaTable("ChangeTable").GetCellData(iCount,0))) = Trim(Lcase(dicChangeParam("ChangeCategory"))) Then
'								JavaWindow("ADS-TeamCenter").JavaWindow("New Change").JavaWindow("Shell").JavaTable("ChangeTable").ClickCell iCount,0,"LEFT","NONE"
'								Exit For
'							End If
'						Next
'					End If
'			   End If
'			   'Set Change Item Affected.
'				If dicChangeParam("ChangeItemAffected")<>"" Then
'					ObjChangeWnd.JavaRadioButton("False").SetTOProperty "Index",0
'					ObjChangeWnd.JavaRadioButton("True").SetTOProperty "Index",0
'					If Trim(Lcase(dicChangeParam("ChangeItemAffected"))) = "true" Then
'						Call Fn_UI_JavaRadioButton_SetON("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd, "True")
'					Else
'						Call Fn_UI_JavaRadioButton_SetON("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd, "False")
'					End If
'				End If
'			   'Set Warranty Affected.
'				If dicChangeParam("WarrantyAffected")<>"" Then
'					ObjChangeWnd.JavaRadioButton("False").SetTOProperty "Index",1
'					ObjChangeWnd.JavaRadioButton("True").SetTOProperty "Index",1
'					If Trim(Lcase(dicChangeParam("WarrantyAffected"))) = "true" Then
'						Call Fn_UI_JavaRadioButton_SetON("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd, "True")
'					Else
'						Call Fn_UI_JavaRadioButton_SetON("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd, "False")
'					End If
'				End If
'			   'Set In Production.
'				If dicChangeParam("InProduction")<>"" Then
'					ObjChangeWnd.JavaRadioButton("False").SetTOProperty "Index",2
'					ObjChangeWnd.JavaRadioButton("True").SetTOProperty "Index",2
'					If Trim(Lcase(dicChangeParam("InProduction"))) = "true" Then
'						Call Fn_UI_JavaRadioButton_SetON("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd, "True")
'					Else
'						Call Fn_UI_JavaRadioButton_SetON("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd, "False")
'					End If
'				End If
'			   'Set Is Primary Change.
'				If dicChangeParam("IsPrimaryChange")<>"" Then
'					ObjChangeWnd.JavaRadioButton("False").SetTOProperty "Index",3
'					ObjChangeWnd.JavaRadioButton("True").SetTOProperty "Index",3
'					If Trim(Lcase(dicChangeParam("IsPrimaryChange"))) = "true" Then
'						Call Fn_UI_JavaRadioButton_SetON("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd, "True")
'					Else
'						Call Fn_UI_JavaRadioButton_SetON("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd, "False")
'					End If
'				End If
'			   'Set Retrofit Required.
'				If dicChangeParam("RetrofitRequired")<>"" Then
'					ObjChangeWnd.JavaRadioButton("False").SetTOProperty "Index",4
'					ObjChangeWnd.JavaRadioButton("True").SetTOProperty "Index",4
'					If Trim(Lcase(dicChangeParam("RetrofitRequired"))) = "true" Then
'						Call Fn_UI_JavaRadioButton_SetON("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd, "True")
'					Else
'						Call Fn_UI_JavaRadioButton_SetON("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd, "False")
'					End If
'				End If
'  Case "SetDREdit","SetDRDropDown"
'			   'Set FilterText
'			   If strFilterText<>"" Then
'				'Call Fn_Edit_Box("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd,"Filter",strFilterText)
'				Call Fn_UI_EditBox_Type("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd,"Filter",strFilterText)
'			   End If
'			   If sNodeName<>"" Then
'				Wait(1)
'				'Selecting Node from tree
'				Call Fn_JavaTree_Select("Fn_ADS_ChangeOperationsExtn", ObjChangeWnd, "ChangeTypeTree","Complete List")
'				Call Fn_JavaTree_Select("Fn_ADS_ChangeOperationsExtn", ObjChangeWnd, "ChangeTypeTree","Complete List:"+sNodeName)
'				'Clicking on Next button to proceed 
'				Call Fn_Button_Click("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd,"Next")
'			   End If
'			   'Set Change Id
'			   If sChangeID<>"" Then
'			      JavaWindow("ADS-TeamCenter").JavaWindow("New Change").JavaEdit("ECRNo_DR").Activate
'							iLen = len(sChangeID)
'								Set WshShell = CreateObject("WScript.Shell")
'							For iCount = 1 to iLen
'								WshShell.SendKeys mid(sChangeID,iCount,1)
'							Next
'								Set WshShell = Nothing
'			   End If
'			   'Set Change Revision
'			   If sChangeRev<>"" Then
'				Call Fn_UI_EditBox_Type("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd,"Revision_DR",sChangeRev)
'			   End If
'			   'Set Synopsis
'			   If sChangeSynopsis<>"" Then
'				Call Fn_UI_EditBox_Type("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd,"Synopsis_DR",sChangeSynopsis)
'			   End If
'			   'Set Description
'			   If sChangeDesc<>"" Then
'				Call Fn_UI_EditBox_Type("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd,"Desc_DR",sChangeDesc)
'			   End If
'			   ObjChangeWnd.Maximize
'			   Wait(3)
'			   'Set Change Type
'			   If sAction = "SetDREdit" Then
'					If sChangeType<>"" Then
'						Call Fn_UI_EditBox_Type("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd,"ChangeType_DR",sChangeType)
'					End If
'			   ElseIf sAction = "SetDRDropDown" Then
'					If sChangeType<>"" Then
'						ObjChangeWnd.JavaButton("DropDownBtn").SetTOProperty "index",0
'						Call Fn_Button_Click("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd,"DropDownBtn")
'						Wait(3)
'						iRows = JavaWindow("ADS-TeamCenter").JavaWindow("New Change").JavaWindow("Shell").JavaTable("ChangeTable").GetROProperty("rows")
'						For iCount = 0 to iRows-1
'							If Trim(Lcase(JavaWindow("ADS-TeamCenter").JavaWindow("New Change").JavaWindow("Shell").JavaTable("ChangeTable").GetCellData(iCount,0))) = Trim(Lcase(sChangeType)) Then
'								JavaWindow("ADS-TeamCenter").JavaWindow("New Change").JavaWindow("Shell").JavaTable("ChangeTable").ClickCell iCount,0,"LEFT","NONE"
'								Exit For
'							End If
'						Next
'					End If
'			   End If
'			   'Code for Fields differing in Case of CN,CR,PR,DR is yet to be coded
'  Case "SetPREdit","SetPRDropDown"
'			   'Set FilterText
'			   If strFilterText<>"" Then
'				'Call Fn_Edit_Box("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd,"Filter",strFilterText)
'				Call Fn_UI_EditBox_Type("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd,"Filter",strFilterText)
'			   End If
'			   If sNodeName<>"" Then
'				Wait(1)
'				'Selecting Node from tree
'				Call Fn_JavaTree_Select("Fn_ADS_ChangeOperationsExtn", ObjChangeWnd, "ChangeTypeTree","Complete List")
'				Call Fn_JavaTree_Select("Fn_ADS_ChangeOperationsExtn", ObjChangeWnd, "ChangeTypeTree","Complete List:"+sNodeName)
'				'Clicking on Next button to proceed 
'				Call Fn_Button_Click("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd,"Next")
'			   End If
'			   'Set Change Id
'			   If sChangeID<>"" Then
'			      JavaWindow("ADS-TeamCenter").JavaWindow("New Change").JavaEdit("PRNo_PR").Activate
'							iLen = len(sChangeID)
'								Set WshShell = CreateObject("WScript.Shell")
'							For iCount = 1 to iLen
'								WshShell.SendKeys mid(sChangeID,iCount,1)
'							Next
'								Set WshShell = Nothing
'			   End If
'			   'Set Change Revision
'			   If sChangeRev<>"" Then
'				Call Fn_UI_EditBox_Type("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd,"Revision_PR",sChangeRev)
'			   End If
'			   'Set Synopsis
'			   If sChangeSynopsis<>"" Then
'				Call Fn_UI_EditBox_Type("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd,"Synopsis_PR",sChangeSynopsis)
'			   End If
'			   'Set Description
'			   If sChangeDesc<>"" Then
'				Call Fn_UI_EditBox_Type("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd,"Desc_PR",sChangeDesc)
'			   End If
'			   'Code for Fields differing in Case of CN,CR,PR,DR is yet to be coded
' End Select
'	 'Click on Buttons
'	 If sButtons<>"" Then
'	   aButtons = split(sButtons, ":",-1,1)
'	   iCounter = Ubound(aButtons)
'	   For iCount=0 to iCounter
'		'Click on Add Button
'		Call Fn_Button_Click("Fn_ADS_ChangeOperationsExtn", ObjChangeWnd, aButtons(iCount))
'        Call Fn_ReadyStatusSync(2)
'	   Next
'	 End If
'	 'function Return True
'	 Fn_ADS_ChangeOperationsExtn=True
'	 'Releasing "New Change" window's object
'	 Set ObjChangeWnd=Nothing
'End Function


Public Function Fn_ADS_ChangeOperationsExtn(sAction, strFilterText, sNodeName, sChangeID, sChangeRev, sChangeSynopsis, sChangeDesc, sChangeType, sVerifyCombos, dicChangeParam, sButtons)
	GBL_FAILED_FUNCTION_NAME="Fn_ADS_ChangeOperationsExtn"
   'Declaring Variables
    Dim ObjChangeWnd, aButtons, iCount, iCounter, objSelectType, intNoOfObjects, sChangeNo, WshShell, iLen, aVerifyCombos
	Fn_ADS_ChangeOperationsExtn=False
	  For iCount=0 to 0
		 JavaWindow("ADS-TeamCenter").JavaWindow("New Change").SetTOProperty "title","New Change"
		 If JavaWindow("ADS-TeamCenter").JavaWindow("New Change").Exist(5) Then
		  Exit For
		 End If
		 JavaWindow("ADS-TeamCenter").JavaWindow("New Change").SetTOProperty "title","New Change in context"
		 If JavaWindow("ADS-TeamCenter").JavaWindow("New Change").Exist(5) Then
		  Exit For
		 End If
		 JavaWindow("ADS-TeamCenter").JavaWindow("New Change").SetTOProperty "title","Derive Change"
		 If JavaWindow("ADS-TeamCenter").JavaWindow("New Change").Exist(5) Then      
		  Exit For
		 End If
	  Next
 Set ObjChangeWnd=Fn_UI_ObjectCreate("Fn_ADS_ChangeOperationsExtn",JavaWindow("ADS-TeamCenter").JavaWindow("New Change"))
 strFilterText = sNodeName
 Select Case sAction
 Case "SetCNEdit","SetCNDropDown"		'Changed the index values for DropDownBtn for ChangeNotice due to design change.
			   'Set FilterText				
			   If strFilterText<>"" Then
				'Call Fn_Edit_Box("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd,"Filter",strFilterText)
				Call Fn_UI_EditBox_Type("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd,"Filter",strFilterText)
			   End If
			   If sNodeName<>"" Then
				Wait(5)
				JavaWindow("ADS-TeamCenter").JavaWindow("New Change").JavaTree("ChangeTypeTree").Select "Complete List"
				JavaWindow("ADS-TeamCenter").JavaWindow("New Change").JavaTree("ChangeTypeTree").Select "Complete List:"+sNodeName
				'Selecting Node from tree
				'Call Fn_JavaTree_Select("Fn_ADS_ChangeOperationsExtn", ObjChangeWnd, "ChangeTypeTree","Complete List")
				'Call Fn_JavaTree_Select("Fn_ADS_ChangeOperationsExtn", ObjChangeWnd, "ChangeTypeTree","Complete List:"+sNodeName)
				'Clicking on Next button to proceed 
				Call Fn_Button_Click("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd,"Next")
				wait 1,500
				
				ObjChangeWnd.Maximize
				wait 3
				
			   End If
			   'Set Change Id
			   If sChangeID<>"" Then
'			      JavaWindow("ADS-TeamCenter").JavaWindow("New Change").JavaEdit("ECNNo_CN").Activate
			      JavaWindow("ADS-TeamCenter").JavaWindow("New Change").JavaEdit("ECNNo_CN").Set  sChangeID
'							iLen = len(sChangeID)
'								Set WshShell = CreateObject("WScript.Shell")
'							For iCount = 1 to iLen
'								WshShell.SendKeys mid(sChangeID,iCount,1)
'							Next
'								Set WshShell = Nothing
			   End If
			   'Set Change Revision
			   If sChangeRev<>"" Then
				Call Fn_UI_EditBox_Type("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd,"Revision_CN",sChangeRev)
			   End If
			   'Set Synopsis
			   If sChangeSynopsis<>"" Then
				Call Fn_UI_EditBox_Type("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd,"Synopsis_CN",sChangeSynopsis)
			   End If
			   'Set Description
			   If sChangeDesc<>"" Then
				Call Fn_UI_EditBox_Type("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd,"Description_CN",sChangeDesc)
			   End If
			   
				Wait(3)
			   'Set Change Type
			   If sAction = "SetCNEdit" Then
					If sChangeType<>"" Then
						Call Fn_UI_EditBox_Type("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd,"ChangeType_CN",sChangeType)
					End If
			   ElseIf sAction = "SetCNDropDown" Then
					If sChangeType<>"" Then
'						ObjChangeWnd.JavaButton("DropDownBtn").SetTOProperty "attached text","*"
						ObjChangeWnd.JavaButton("DropDownBtn").SetTOProperty "index",0
						Call Fn_Button_Click("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd,"DropDownBtn")
						Wait(3)
						'*Modified By Nilesh on 15-Feb-2013 
						If  JavaWindow("ADS-TeamCenter").JavaWindow("New Change").JavaWindow("Shell").JavaTable("ChangeTable").Exist(5)=True Then
							iRows = JavaWindow("ADS-TeamCenter").JavaWindow("New Change").JavaWindow("Shell").JavaTable("ChangeTable").GetROProperty("rows")
							For iCount = 0 to iRows-1
								If Trim(Lcase(JavaWindow("ADS-TeamCenter").JavaWindow("New Change").JavaWindow("Shell").JavaTable("ChangeTable").GetCellData(iCount,0))) = Trim(Lcase(sChangeType)) Then
									JavaWindow("ADS-TeamCenter").JavaWindow("New Change").JavaWindow("Shell").JavaTable("ChangeTable").ClickCell iCount,0,"LEFT","NONE"
									Exit For
								End If
							Next
						Else 
								ObjChangeWnd.JavaWindow("Shell").JavaTree("ChangeTree").Activate sChangeType
						End If
						'*End
					End If
			   End If
			   'Code for Fields differing in Case of CN,CR,PR,DR is yet to be coded
			   'Set Change Item Affected.
				If dicChangeParam("PaperChange")<>"" Then
					ObjChangeWnd.JavaRadioButton("False").SetTOProperty "Index",0
					ObjChangeWnd.JavaRadioButton("True").SetTOProperty "Index",0
					If Trim(Lcase(dicChangeParam("PaperChange"))) = "true" Then
						Call Fn_UI_JavaRadioButton_SetON("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd, "True")
					Else
						Call Fn_UI_JavaRadioButton_SetON("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd, "False")
					End If
				End If
			   'Set Change Class
			   If sAction = "SetCNEdit" Then
					If dicChangeParam("ChangeClass")<>"" Then
						Call Fn_UI_EditBox_Type("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd,"Change Class_CN",dicChangeParam("ChangeClass"))
					End If
			   ElseIf sAction = "SetCNDropDown" Then
					If dicChangeParam("ChangeClass")<>"" Then
'						ObjChangeWnd.JavaButton("DropDownBtn").SetTOProperty "attached text",""
						ObjChangeWnd.JavaButton("DropDownBtn").SetTOProperty "index",0
						Call Fn_Button_Click("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd,"DropDownBtn")
						Wait(3)
						JavaWindow("ADS-TeamCenter").JavaWindow("New Change").JavaWindow("Shell").SetToProperty "index",0
						'*Modified By Nilesh on 15-Feb-2013 
						If  JavaWindow("ADS-TeamCenter").JavaWindow("New Change").JavaWindow("Shell").JavaTable("ChangeTable").Exist(5)=True Then
							iRows = JavaWindow("ADS-TeamCenter").JavaWindow("New Change").JavaWindow("Shell").JavaTable("ChangeTable").GetROProperty("rows")
							For iCount = 0 to iRows-1
								If Trim(Lcase(JavaWindow("ADS-TeamCenter").JavaWindow("New Change").JavaWindow("Shell").JavaTable("ChangeTable").GetCellData(iCount,0))) = Trim(Lcase(dicChangeParam("ChangeClass"))) Then
									JavaWindow("ADS-TeamCenter").JavaWindow("New Change").JavaWindow("Shell").JavaTable("ChangeTable").ClickCell iCount,0,"LEFT","NONE"
									Exit For
								End If
							Next
						Else
							ObjChangeWnd.JavaWindow("Shell").JavaTree("ChangeTree").Activate dicChangeParam("ChangeClass")
						End If
						'*End
					End If
			   End If
			   'Set Change Category
			   If sAction = "SetCNEdit" Then
					If dicChangeParam("ChangeCategory")<>"" Then
						Call Fn_UI_EditBox_Type("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd,"Category_CN",dicChangeParam("ChangeCategory"))
					End If
			   ElseIf sAction = "SetCNDropDown" Then
					If dicChangeParam("ChangeCategory")<>"" Then
						ObjChangeWnd.JavaButton("DropDownBtn").SetTOProperty "index",1
						Call Fn_Button_Click("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd,"DropDownBtn")
						ObjChangeWnd.JavaButton("DropDownBtn").Click 
						Wait(3)
						JavaWindow("ADS-TeamCenter").JavaWindow("New Change").JavaWindow("Shell").SetToProperty "index",1
						'*Modified By Nilesh on 15-Feb-2013 
						If   JavaWindow("ADS-TeamCenter").JavaWindow("New Change").JavaWindow("Shell").JavaTable("ChangeTable").Exist(5)=True Then
							iRows = JavaWindow("ADS-TeamCenter").JavaWindow("New Change").JavaWindow("Shell").JavaTable("ChangeTable").GetROProperty("rows")
							For iCount = 0 to iRows-1
								If Trim(Lcase(JavaWindow("ADS-TeamCenter").JavaWindow("New Change").JavaWindow("Shell").JavaTable("ChangeTable").GetCellData(iCount,0))) = Trim(Lcase(dicChangeParam("ChangeCategory"))) Then
									JavaWindow("ADS-TeamCenter").JavaWindow("New Change").JavaWindow("Shell").JavaTable("ChangeTable").ClickCell iCount,0,"LEFT","NONE"
									Exit For
								End If
							Next
						Else
							ObjChangeWnd.JavaWindow("Shell").JavaTree("ChangeTree").Activate dicChangeParam("ChangeCategory")
						End If
						'*End
					End If
			   End If
  Case "SetCREdit","SetCRDropDown"
			   'Set FilterText
			   If strFilterText<>"" Then
				'Call Fn_Edit_Box("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd,"Filter",strFilterText)
				Call Fn_UI_EditBox_Type("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd,"Filter",strFilterText)
			   End If
			   If sNodeName<>"" Then
				Wait(1)
				'Selecting Node from tree
				Call Fn_JavaTree_Select("Fn_ADS_ChangeOperationsExtn", ObjChangeWnd, "ChangeTypeTree","Complete List")
				Call Fn_JavaTree_Select("Fn_ADS_ChangeOperationsExtn", ObjChangeWnd, "ChangeTypeTree","Complete List:"+sNodeName)
				'Clicking on Next button to proceed 
				Call Fn_Button_Click("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd,"Next")
				wait 1
			   End If
			   'Set Change Id
			   If sChangeID<>"" Then				
			      JavaWindow("ADS-TeamCenter").JavaWindow("New Change").JavaEdit("ECRNo_CR").Activate
							iLen = len(sChangeID)
								Set WshShell = CreateObject("WScript.Shell")
							For iCount = 1 to iLen
								WshShell.SendKeys mid(sChangeID,iCount,1)
								wait 0,200
							Next
								Set WshShell = Nothing
			   End If
			   'Set Change Revision
			   If sChangeRev<>"" Then
				Call Fn_UI_EditBox_Type("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd,"Revision_CR",sChangeRev)
			   End If
			   'Set Synopsis
			   If sChangeSynopsis<>"" Then
				Call Fn_UI_EditBox_Type("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd,"Synopsis_CR",sChangeSynopsis)
			   End If
			   'Set Description
			   If sChangeDesc<>"" Then
				Call Fn_UI_EditBox_Type("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd,"Desc_CR",sChangeDesc)
			   End If
			   Wait(3)
			   'Set Change Type
			   If sAction = "SetCREdit" Then
					If sChangeType<>"" Then
						If strFilterText="Deviation Request" Then
							Call Fn_UI_Object_SetTOProperty_ExistCheck("RACUpdateActionItem",JavaWindow("ADS-TeamCenter").JavaWindow("New Change").JavaStaticText("ChangeType"),"label","Deviation Type:")
						End If
						Call Fn_UI_EditBox_Type("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd,"ChangeType_CR",sChangeType)
					End If
			   ElseIf sAction = "SetCRDropDown" Then
					If sChangeType<>"" Then
						ObjChangeWnd.JavaButton("DropDownBtn").SetTOProperty "index",0
						Call Fn_Button_Click("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd,"DropDownBtn")
						Wait(3)
						'*Modified By Nilesh on 15-Feb-2013 
						If JavaWindow("ADS-TeamCenter").JavaWindow("New Change").JavaWindow("Shell").JavaTable("ChangeTable").Exist(5)=True Then
							iRows = JavaWindow("ADS-TeamCenter").JavaWindow("New Change").JavaWindow("Shell").JavaTable("ChangeTable").GetROProperty("rows")
							For iCount = 0 to iRows-1
								If Trim(Lcase(JavaWindow("ADS-TeamCenter").JavaWindow("New Change").JavaWindow("Shell").JavaTable("ChangeTable").GetCellData(iCount,0))) = Trim(Lcase(sChangeType)) Then
									JavaWindow("ADS-TeamCenter").JavaWindow("New Change").JavaWindow("Shell").JavaTable("ChangeTable").ClickCell iCount,0,"LEFT","NONE"
									Exit For
								End If
							Next
						Else
							ObjChangeWnd.JavaWindow("Shell").JavaTree("ChangeTree").Activate sChangeType
						End If
						'*End
					End If
			   End If
			   'Code for Fields differing in Case of CN,CR,PR,DR is yet to be coded
			   'Set Change Class
			If vartype(dicChangeParam) = 9 Then
				If sAction = "SetCREdit" Then
					If dicChangeParam("ChangeClass")<>"" Then
						Call Fn_UI_EditBox_Type("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd,"ChangeClass_CR",dicChangeParam("ChangeClass"))
					End If
				ElseIf sAction = "SetCRDropDown" Then
					If dicChangeParam("ChangeClass")<>"" Then
						ObjChangeWnd.JavaButton("DropDownBtn").SetTOProperty "index",1
						Call Fn_Button_Click("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd,"DropDownBtn")
						Wait(3)
						'*Modified By Nilesh on 15-Feb-2013 
						If  JavaWindow("ADS-TeamCenter").JavaWindow("New Change").JavaWindow("Shell").JavaTable("ChangeTable").Exist(5)=True Then
							iRows = JavaWindow("ADS-TeamCenter").JavaWindow("New Change").JavaWindow("Shell").JavaTable("ChangeTable").GetROProperty("rows")
							For iCount = 0 to iRows-1
								If Trim(Lcase(JavaWindow("ADS-TeamCenter").JavaWindow("New Change").JavaWindow("Shell").JavaTable("ChangeTable").GetCellData(iCount,0))) = Trim(Lcase(dicChangeParam("ChangeClass"))) Then
									JavaWindow("ADS-TeamCenter").JavaWindow("New Change").JavaWindow("Shell").JavaTable("ChangeTable").ClickCell iCount,0,"LEFT","NONE"
									Exit For
								End If
							Next
						Else
								ObjChangeWnd.JavaWindow("Shell").JavaTree("ChangeTree").Activate dicChangeParam("ChangeClass")
						End If
						'*End
					End If
			   End If
			   'Set Change Category
			   If sAction = "SetCREdit" Then
					If dicChangeParam("ChangeCategory")<>"" Then
						Call Fn_UI_EditBox_Type("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd,"ChangeCategory_CR",dicChangeParam("ChangeCategory"))
					End If
			   ElseIf sAction = "SetCRDropDown" Then
					If dicChangeParam("ChangeCategory")<>"" Then
						ObjChangeWnd.JavaButton("DropDownBtn").SetTOProperty "index",2
						Call Fn_Button_Click("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd,"DropDownBtn")
						Wait(3)
						'*Modified By Nilesh on 15-Feb-2013 
						If  JavaWindow("ADS-TeamCenter").JavaWindow("New Change").JavaWindow("Shell").JavaTable("ChangeTable").Exist(5)=True Then
							iRows = JavaWindow("ADS-TeamCenter").JavaWindow("New Change").JavaWindow("Shell").JavaTable("ChangeTable").GetROProperty("rows")
							For iCount = 0 to iRows-1
								If Trim(Lcase(JavaWindow("ADS-TeamCenter").JavaWindow("New Change").JavaWindow("Shell").JavaTable("ChangeTable").GetCellData(iCount,0))) = Trim(Lcase(dicChangeParam("ChangeCategory"))) Then
									JavaWindow("ADS-TeamCenter").JavaWindow("New Change").JavaWindow("Shell").JavaTable("ChangeTable").ClickCell iCount,0,"LEFT","NONE"
									Exit For
								End If
							Next
						Else
	                        ObjChangeWnd.JavaWindow("Shell").JavaTree("ChangeTree").Activate dicChangeParam("ChangeCategory")
						End If
						'*End
					End If
			   End If
			   'Set Change Item Affected.
				If dicChangeParam("ChangeItemAffected")<>"" Then
					ObjChangeWnd.JavaRadioButton("False").SetTOProperty "Index",0
					ObjChangeWnd.JavaRadioButton("True").SetTOProperty "Index",0
					If Trim(Lcase(dicChangeParam("ChangeItemAffected"))) = "true" Then
						Call Fn_UI_JavaRadioButton_SetON("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd, "True")
					Else
						Call Fn_UI_JavaRadioButton_SetON("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd, "False")
					End If
				End If
			   'Set Warranty Affected.
				If dicChangeParam("WarrantyAffected")<>"" Then
					ObjChangeWnd.JavaRadioButton("False").SetTOProperty "Index",1
					ObjChangeWnd.JavaRadioButton("True").SetTOProperty "Index",1
					If Trim(Lcase(dicChangeParam("WarrantyAffected"))) = "true" Then
						Call Fn_UI_JavaRadioButton_SetON("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd, "True")
					Else
						Call Fn_UI_JavaRadioButton_SetON("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd, "False")
					End If
				End If
			   'Set In Production.
				If dicChangeParam("InProduction")<>"" Then
					ObjChangeWnd.JavaRadioButton("False").SetTOProperty "Index",2
					ObjChangeWnd.JavaRadioButton("True").SetTOProperty "Index",2
					If Trim(Lcase(dicChangeParam("InProduction"))) = "true" Then
						Call Fn_UI_JavaRadioButton_SetON("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd, "True")
					Else
						Call Fn_UI_JavaRadioButton_SetON("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd, "False")
					End If
				End If
			   'Set Is Primary Change.
				If dicChangeParam("IsPrimaryChange")<>"" Then
					ObjChangeWnd.JavaRadioButton("False").SetTOProperty "Index",3
					ObjChangeWnd.JavaRadioButton("True").SetTOProperty "Index",3
					If Trim(Lcase(dicChangeParam("IsPrimaryChange"))) = "true" Then
						Call Fn_UI_JavaRadioButton_SetON("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd, "True")
					Else
						Call Fn_UI_JavaRadioButton_SetON("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd, "False")
					End If
				End If
			   'Set Retrofit Required.
				If dicChangeParam("RetrofitRequired")<>"" Then
					ObjChangeWnd.JavaRadioButton("False").SetTOProperty "Index",4
					ObjChangeWnd.JavaRadioButton("True").SetTOProperty "Index",4
					If Trim(Lcase(dicChangeParam("RetrofitRequired"))) = "true" Then
						Call Fn_UI_JavaRadioButton_SetON("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd, "True")
					Else
						Call Fn_UI_JavaRadioButton_SetON("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd, "False")
					End If
				End If
			End If
  Case "SetDREdit","SetDRDropDown"
			   'Set FilterText
			   If strFilterText<>"" Then
				'Call Fn_Edit_Box("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd,"Filter",strFilterText)
				Call Fn_UI_EditBox_Type("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd,"Filter",strFilterText)
			   End If
			   If sNodeName<>"" Then
				Wait(1)
				'Selecting Node from tree
				Call Fn_JavaTree_Select("Fn_ADS_ChangeOperationsExtn", ObjChangeWnd, "ChangeTypeTree","Complete List")
				Call Fn_JavaTree_Select("Fn_ADS_ChangeOperationsExtn", ObjChangeWnd, "ChangeTypeTree","Complete List:"+sNodeName)
				'Clicking on Next button to proceed 
				Call Fn_Button_Click("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd,"Next")
				Wait(3)
			   End If
			   
			   'Set Change Id
			   If sChangeID<>"" Then
			      JavaWindow("ADS-TeamCenter").JavaWindow("New Change").JavaEdit("ECRNo_DR").Activate
				  Wait(3)
  				  Call Fn_SISW_UI_JavaEdit_Operations("Fn_ADS_ChangeOperationsExtn", "Set",  JavaWindow("ADS-TeamCenter").JavaWindow("New Change"), "ECRNo_DR", sChangeID )
				  Wait(3)

							'iLen = len(sChangeID)
							'	Set WshShell = CreateObject("WScript.Shell")
							'For iCount = 1 to iLen
							'	WshShell.SendKeys mid(sChangeID,iCount,1)
							'Next
							'	Set WshShell = Nothing
			   End If
			   'Set Change Revision
			   If sChangeRev<>"" Then
				Call Fn_UI_EditBox_Type("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd,"Revision_DR",sChangeRev)
				Wait(3)
			   End If
			   'Set Synopsis
			   If sChangeSynopsis<>"" Then
				Call Fn_UI_EditBox_Type("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd,"Synopsis_DR",sChangeSynopsis)
				Wait(3)
			   End If
			   'Set Description
			   If sChangeDesc<>"" Then
				Call Fn_UI_EditBox_Type("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd,"Desc_DR",sChangeDesc)
				Wait(3)
			   End If
			   'ObjChangeWnd.Maximize
			   Wait(3)
			   'Set Change Type
			   If sAction = "SetDREdit" Then
					If sChangeType<>"" Then
						Call Fn_UI_EditBox_Type("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd,"ChangeType_DR",sChangeType)
					End If
			   ElseIf sAction = "SetDRDropDown" Then
					If sChangeType<>"" Then
						ObjChangeWnd.JavaButton("DropDownBtn").SetTOProperty "index",0
						Call Fn_Button_Click("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd,"DropDownBtn")
						Wait(3)
						'*Modified By Nilesh on 15-Feb-2013 
						If JavaWindow("ADS-TeamCenter").JavaWindow("New Change").JavaWindow("Shell").JavaTable("ChangeTable").Exist(5)=True	 Then
							iRows = JavaWindow("ADS-TeamCenter").JavaWindow("New Change").JavaWindow("Shell").JavaTable("ChangeTable").GetROProperty("rows")
							For iCount = 0 to iRows-1
								If Trim(Lcase(JavaWindow("ADS-TeamCenter").JavaWindow("New Change").JavaWindow("Shell").JavaTable("ChangeTable").GetCellData(iCount,0))) = Trim(Lcase(sChangeType)) Then
									JavaWindow("ADS-TeamCenter").JavaWindow("New Change").JavaWindow("Shell").JavaTable("ChangeTable").ClickCell iCount,0,"LEFT","NONE"
									Exit For
								End If
							Next
						Else
							ObjChangeWnd.JavaWindow("Shell").JavaTree("ChangeTree").Activate sChangeType
						End If
						'*End
					End If
			   End If
			   'Code for Fields differing in Case of CN,CR,PR,DR is yet to be coded
  Case "SetPREdit","SetPRDropDown"
			   'Set FilterText
			   If strFilterText<>"" Then
				'Call Fn_Edit_Box("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd,"Filter",strFilterText)
				Call Fn_UI_EditBox_Type("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd,"Filter",strFilterText)
			   End If
			   If sNodeName<>"" Then
				Wait(1)
				'Selecting Node from tree
				Call Fn_JavaTree_Select("Fn_ADS_ChangeOperationsExtn", ObjChangeWnd, "ChangeTypeTree","Complete List")
				Call Fn_JavaTree_Select("Fn_ADS_ChangeOperationsExtn", ObjChangeWnd, "ChangeTypeTree","Complete List:"+sNodeName)
				'Clicking on Next button to proceed 
				Call Fn_Button_Click("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd,"Next")
			   End If
			   'Set Change Id
			   If sChangeID<>"" Then
			      JavaWindow("ADS-TeamCenter").JavaWindow("New Change").JavaEdit("PRNo_PR").Activate
							iLen = len(sChangeID)
								Set WshShell = CreateObject("WScript.Shell")
							For iCount = 1 to iLen
								WshShell.SendKeys mid(sChangeID,iCount,1)
							Next
								Set WshShell = Nothing
			   End If
			   'Set Change Revision
			   If sChangeRev<>"" Then
				Call Fn_UI_EditBox_Type("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd,"Revision_PR",sChangeRev)
			   End If
			   'Set Synopsis
			   If sChangeSynopsis<>"" Then
				Call Fn_UI_EditBox_Type("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd,"Synopsis_PR",sChangeSynopsis)
			   End If
			   'Set Description
			   If sChangeDesc<>"" Then
				Call Fn_UI_EditBox_Type("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd,"Desc_PR",sChangeDesc)
			   End If
			   'Code for Fields differing in Case of CN,CR,PR,DR is yet to be coded
 End Select
	 'Click on Buttons
	 If sButtons<>"" Then
	   aButtons = split(sButtons, ":",-1,1)
	   iCounter = Ubound(aButtons)
	   For iCount=0 to iCounter
		'Click on Add Button
		Call Fn_Button_Click("Fn_ADS_ChangeOperationsExtn", ObjChangeWnd, aButtons(iCount))
        Call Fn_ReadyStatusSync(2)
	   Next
	 End If
	 'function Return True
	 Fn_ADS_ChangeOperationsExtn=True
	 'Releasing "New Change" window's object
	 Set ObjChangeWnd=Nothing
End Function

''#######################################################################################
''###     FUNCTION NAME   :   Fn_ADS_VariableLenRandomNoGen(iLength)
'###    DESCRIPTION     :   	Generates a variable Length random number
'###    PARAMETERS      :     iLength
'###    Return Value  	:   	  Random number of given length
'###    HISTORY         :   	 AUTHOR              	DATE        		VERSION
'###    CREATED BY      :    Ketan Raje			   15/10/2010   			1.0
'###    REVIWED BY      :	 Harshal	 			15/10/2010
'###    EXAMPLE         :   
'#############################################################################################
Public Function Fn_ADS_VariableLenRandomNoGen(iLength)
Dim num
Randomize
On Error Resume Next
    num = Right("00000" & Int(Rnd * 1000000), iLength)
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), num& " number generated of length " &iLength)
	Fn_ADS_VariableLenRandomNoGen = num
End Function
'#######################################################################################
'###     FUNCTION NAME   :   Fn_ContractEventSchedule(sAction, sProperty,sValue)
'###    DESCRIPTION     :   	for Schedule templete
'###    PARAMETERS      :     sAction, sProperty,sValue
'###    Return Value  	:   	  True/False
'###    HISTORY         :   	 AUTHOR              	DATE        		VERSION
'###    CREATED BY      :    Harshal			   18/10/2010   			1.0
'###    REVIWED BY      :	 Harshal	 			18/10/2010
'###    EXAMPLE         :   Case "Set" : MsgBox Fn_ContractEventSchedule("Set","ScheduleTemplate","She1_000592")
'###    				         :   Case "Verify" : Msgbox Fn_ContractEventSchedule("Verify", "ScheduleTemplate","ADSSchedule7913_000816")
'#############################################################################################
Function Fn_ContractEventSchedule(sAction, sProperty,sValue)
	GBL_FAILED_FUNCTION_NAME="Fn_ContractEventSchedule"
	Dim aValue, iFlag, iCount, iCounter, intCount
	If JavaWindow("ADS-TeamCenter").JavaWindow("Contract Event Schedule").Exist Then
		Select Case sAction
			Case "Set"
				Select Case sProperty
					Case "ScheduleTemplate"
						JavaWindow("ADS-TeamCenter").JavaWindow("Contract Event Schedule").JavaList("Schedule Template").Select sValue
						Fn_ContractEventSchedule = True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_ContractEventSchedule executed Sucessful")
                    Case Else
						Fn_ContractEventSchedule = False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_ContractEventSchedule execution Failed")
				End Select
				JavaWindow("ADS-TeamCenter").JavaWindow("Contract Event Schedule").JavaButton("OK").Click micLeftBtn
			Case "Verify"
				Select Case sProperty
					Case "ScheduleTemplate"							
							iFlag = 0
							'For Schedule Template
							If sValue <>  "" then
								aValue = Split(sValue,":",-1,1)
								iCount = cint(JavaWindow("ADS-TeamCenter").JavaWindow("Contract Event Schedule").JavaList("Schedule Template").GetROProperty("items count"))
								For iCounter=0 to Ubound(aValue)
									For intCount=0 to iCount-1
										If Trim(Lcase(aValue(iCounter))) = Trim(Lcase(JavaWindow("ADS-TeamCenter").JavaWindow("Contract Event Schedule").JavaList("Schedule Template").GetItem(intCount))) Then
											iFlag = iFlag+1
											Exit For
										End If
									Next
								Next
								If iFlag = Ubound(aValue)+1 Then
									Fn_ContractEventSchedule = True
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_ContractEventSchedule :Value found in List.")
								Else
									Fn_ContractEventSchedule = False
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_ContractEventSchedule :Value not found in List.")
								End If
							End If	
                    Case Else
						Fn_ContractEventSchedule = False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_ContractEventSchedule execution Failed")
				End Select
				JavaWindow("ADS-TeamCenter").JavaWindow("Contract Event Schedule").JavaButton("Cancel").Click micLeftBtn
			Case Else
					Fn_ContractEventSchedule = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_ContractEventSchedule execution Failed")
		End Select
	Else
		Fn_ContractEventSchedule = False
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_ContractEventSchedule execution Failed")
	End If
End Function
'#######################################################################################
'###     FUNCTION NAME   :   Fn_ADS_GenerateSubmittalDelivery(sAction,sProcessTemplate,sTaskDuration,sAlignTask,sButton)
'###    DESCRIPTION     :   	for Schedule templete
'###    PARAMETERS      :     sAction,sProcessTemplate,sTaskDuration,sAlignTask,sButton
'###    Return Value  	:   	  True/False
'###    HISTORY         :   	 AUTHOR              	DATE        		VERSION
'###    CREATED BY      :    Harshal			   18/10/2010   			1.0
'###    REVIWED BY      :	 Harshal	 			18/10/2010
'###    EXAMPLE         :   Case "Set" : MsgBox Fn_ADS_GenerateSubmittalDelivery("Set","AAU2","1:1:1:1:1","StartDate","Cancel")
'###								 Case "UIExist" : Msgbox Fn_ADS_GenerateSubmittalDelivery("UIExist","","","","Cancel")
'#############################################################################################
Function Fn_ADS_GenerateSubmittalDelivery(sAction,sProcessTemplate,sTaskDuration,sAlignTask,sButton)
	GBL_FAILED_FUNCTION_NAME="Fn_ADS_GenerateSubmittalDelivery"
	Dim objGSD,aTaskDuration,iCounter
	iCounter = 0
	Set objGSD =JavaWindow("ADS-TeamCenter").JavaWindow("Generate Submittal Delivery")
	Select Case sAction
	Case "Set"
			If objGSD.Exist Then
				If sProcessTemplate<>"" Then
					objGSD.JavaList("Process Template").Select sProcessTemplate
				End If
				If sTaskDuration<>""  Then
					objGSD.JavaButton("Set").Click micLeftBtn
					If objGSD.JavaWindow("Set Duration").Exist Then
						objGSD.JavaWindow("Set Duration").JavaButton("Clear").Click
						aTaskDuration = Split(sTaskDuration,":")
						objGSD.JavaWindow("Set Duration").JavaEdit("Years").Set aTaskDuration(0)
						objGSD.JavaWindow("Set Duration").JavaEdit("Weeks").Set aTaskDuration(1)
						objGSD.JavaWindow("Set Duration").JavaEdit("Days").Set aTaskDuration(2)
						objGSD.JavaWindow("Set Duration").JavaEdit("Hours").Set aTaskDuration(3)
						objGSD.JavaWindow("Set Duration").JavaEdit("Minutes").Set aTaskDuration(4)
						objGSD.JavaWindow("Set Duration").JavaButton("OK").Click micLeftBtn
					Else
						Fn_ADS_GenerateSubmittalDelivery = False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_ADS_GenerateSubmittalDelivery execution Failed")
					End If
				End If
						If sAlignTask<>"" Then
							If sAlignTask = "EndDate" Then
								objGSD.JavaRadioButton("End Date with Submittal").Set "ON"
							ElseIf sAlignTask = "StartDate" Then
								objGSD.JavaRadioButton("Start Date with Submittal").Set "ON"
							End If
						End If
						objGSD.JavaButton(sButton).Click micLeftBtn
						Fn_ADS_GenerateSubmittalDelivery = True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_ADS_GenerateSubmittalDelivery execution Sucessful")
			Else
				Fn_ADS_GenerateSubmittalDelivery = False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_ADS_GenerateSubmittalDelivery execution Failed")
			End If
	Case "UIExist"
			'Check if Process Template List exist
			If objGSD.JavaList("Process Template").Exist Then
				iCounter = iCounter + 1
			End If
			'Check if Task Duration Editbox exist
			If objGSD.JavaEdit("Task Duration").Exist Then
				iCounter = iCounter + 1
			End If
			'Check if Radio button for Start Date exist
			If objGSD.JavaRadioButton("Start Date with Submittal").Exist Then
				iCounter = iCounter + 1
			End If
			'Check if Radio button for End Date exist
			If objGSD.JavaRadioButton("End Date with Submittal").Exist Then
				iCounter = iCounter + 1
			End If
			'Check if OK button exist
			If objGSD.JavaButton("OK").Exist Then
				iCounter = iCounter + 1
			End If
			'Check if Cancel button exist
			If objGSD.JavaButton("Cancel").Exist Then
				iCounter = iCounter + 1
			End If
			'Check if Set button exist
			If objGSD.JavaButton("Set").Exist Then
				iCounter = iCounter + 1
				objGSD.JavaButton("Set").Click micLeftBtn
				If objGSD.JavaWindow("Set Duration").Exist Then
					iCounter = iCounter + 1
					objGSD.JavaWindow("Set Duration").JavaButton("Cancel").Click micLeftBtn
				End If
			End If
			'Click on sbutton 
			objGSD.JavaButton(sButton).Click micLeftBtn
			If iCounter=8 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "UI exist successfully verified")
				Fn_ADS_GenerateSubmittalDelivery = TRUE
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "UI exist failed to verify")
				Fn_ADS_GenerateSubmittalDelivery = FALSE
			End If
	Case Else						
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_ADS_GenerateSubmittalDelivery function failed")
			Fn_ADS_GenerateSubmittalDelivery = FALSE
	End Select
	Set objGSD = nothing
End Function
'*********************************************************		Function to create detail Schedule		***********************************************************************
'Function Name		:        Fn_ADS_ScheduleDetailCreate  

'Description	    	:        Creates an Schedule with detail information

'Parameters		     :    		sSchedulType: Schedule type to be selected
'			                         		 sSchedulID: Unique ID for the Schedule [if non-empty, then enter]
'							          		sSchedulRevID: Revision of the Schedule [if non-empty, then enter] - if any one of the fields (id/rev) are blank then click Assign button
'									 		sSchedulName: Name of the Schedule
'									  		sSchedulDesc: Description of the Schedule
' 											dicSchParam : Dictionary paramter  for detail creation

'Return Value		: 			SchedulID-SchedulRevID 

'Pre-requisite	    :		 	Should be logged in

'Examples		    :			Call Fn_ADS_ScheduleDetailCreate("Schedule","","","TesSch","sfdwtrtyjghjjg",dicScheduleInfo)

'History		    :		
'													Developer Name				Date						Rev. No.						Changes Done						Reviewer
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Harshal 							     27/05/2010			              1.0								
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_ADS_ScheduleDetailCreate(sSchedulType ,sSchedulID,sSchedulRevID,sSchedulName ,sSchedulDesc,dicScheduleInfo)
   Fn_ADS_ScheduleDetailCreate = Fn_ScheduleDetailCreate(sSchedulType ,sSchedulID,sSchedulRevID,sSchedulName ,sSchedulDesc,dicScheduleInfo)
End Function
'***********************************************************************************************************************************************************************************
'Function Name		         :	       Fn_ADS_ViewerTabOperation()

'Description			   :		   This function is used to perform operation on ViewerTab Dialog (Event List)

'Parameters			  :	 				1.   sAction - Action need to perform
'
'											  		2.  dicViewerTab - Dictionary Object  
'											    										
'Return Value		            :		True / False

'Pre-requisite			   :		  Viewer Tab should be Diaplayed
'
'Example with Dictionary Calls:
'												'dicViewerTab("EventName")="Test"
'												'dicViewerTab("StartDate")="01-Jun-2011 08:00:00"
'												'dicViewerTab("EndDate")="30-Jun-2011 17:00:00"
'												'dicViewerTab("Offset")="1"
'												'dicViewerTab("RelativeTo")="End Date"
'												'dicViewerTab("Recurrence")="Monthly"
'												'dicViewerTab("RecurrenceEndDate")="30-Jun-2011 17:00:00"
'												'msgbox Fn_ADS_ViewerTabOperation("Add",dicViewerTab)
'												
'												'dicViewerTab("ButtonName")="Save"
'												'msgbox Fn_ADS_ViewerTabOperation("Save",dicViewerTab)
'												
'												'msgbox Fn_ADS_ViewerTabOperation("Delete","")
'												
'												'dicViewerTab("EventName")="Test"
'												'msgbox Fn_ADS_ViewerTabOperation("Verify",dicViewerTab)
'												
'												'dicViewerTab("UIColumns")="Event Name:Start Date:End Date:Offset:Relative To:Recurrence:Recurrence End Date"
'												'dicViewerTab("ColsData")="Test|01-Jun-2011 08:00|30-Jun-2011 17:00|1|End Date|Monthly|30-Jun-2011 17:00"
'												'msgbox Fn_ADS_ViewerTabOperation("VerifyData",dicViewerTab)
'												
'												'dicViewerTab("UIColumns")="Event Name:Start Date:End Date:Offset:Relative To:Recurrence:Recurrence End Date"
'												'msgbox Fn_ADS_ViewerTabOperation("UIExist",dicViewerTab)
'												
'												'dicViewerTab("ColsData")="1"
'												'msgbox Fn_ADS_ViewerTabOperation("GetDataList",dicViewerTab)
'												
'History:
'						Developer Name					Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'					   			Govind Singh		20-Oct-2010															Harshal Agarwal
'
'	Modified By: 	Pranav Shirode	  23-Feb-2012
'***********************************************************************************************************************************************************************************						
'Public Function Fn_ADS_ViewerTabOperation(sAction,dicViewerTab)
'   On Error Resume Next
'	Dim iCount, iRows,ObjEventTable,ObjWin,bReturn,WshShell,aDate,objDateControl
'	Dim iFlag, aEventName,iCounter, aColumns, iCols ,aColsData, iRowCnt, intCount, sReturn
'
'	Set ObjEventTable = JavaWindow("ADS-TeamCenter").JavaTable("Event List:")
'	Set ObjWin = JavaWindow("ADS-TeamCenter")
'
'	'Maximize Viewer Tab
'	Call Fn_TabFolder_Operation("DoubleClickTab", "Viewer", "")
'	Call Fn_ReadyStatusSync(1)
'
'	If JavaWindow("ADS-TeamCenter").Exist Then
'		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_ADS_ViewerTabOperation :Viewer Tab Exists")
'			Select Case sAction
'			Case "Add"
'					Fn_ADS_ViewerTabOperation = False									
'					Call Fn_Button_Click("Fn_ADS_ViewerTabOperation",ObjWin,"Add")
'					iRows =  cint(ObjEventTable.GetROProperty("rows"))-1
'					ObjEventTable.SelectRow iRows 
'					ObjWin.JavaList("EventName").SetToProperty "Index",Cint (iRows+0)
'					ObjWin.JavaList("RelativeTo").SetToProperty "Index", Cint(iRows*2)
'					ObjWin.JavaList("Recurrence").SetToProperty "Index",CInt( iRows*2+1)
'							'For Event Name list
'					If dicViewerTab("EventName") <>  "" Then
'						ObjEventTable.SelectRow iRows 
'						ObjEventTable.ClickCell iRows ,"Event Name"
'						ObjWin.JavaList("EventName").Select dicViewerTab("EventName")
'						Fn_ADS_ViewerTabOperation = True
'						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_ADS_ViewerTabOperation :Event Name "+dicViewerTab("EventName")+" set successfully")
'					 End If
'					Call Fn_ReadyStatusSync(1)
'
'				'Start Date of Event
'					If dicViewerTab("StartDate") <>  "" then
'							ObjEventTable.SelectRow iRows 
'							ObjEventTable.ActivateCell iRows ,"Start Date"
'							aDate=Split(dicViewerTab("StartDate") ," ",-1,1)
'							Set objDateControl=JavaWindow("ADS-TeamCenter").JavaWindow("Select Date")
'							objDateControl.JavaCalendar("Date").SetDate aDate(0)
'							objDateControl.JavaCalendar("Time").SetTime aDate(1)							
'							objDateControl.JavaButton("OK").Click
'							Fn_ADS_ViewerTabOperation = True
'							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_ADS_ViewerTabOperation :Start Date "+dicViewerTab("StartDate")+" set successfully")
'					End If
'					Call Fn_ReadyStatusSync(1)
'
'				'End Date of Event
'					If dicViewerTab("EndDate") <>  "" then
'							ObjEventTable.SelectRow iRows 
'							ObjEventTable.ActivateCell iRows ,"End Date"
'							aDate=Split(dicViewerTab("EndDate") ," ",-1,1)
'							Set objDateControl=JavaWindow("ADS-TeamCenter").JavaWindow("Select Date")
'							objDateControl.JavaCalendar("Date").SetDate aDate(0)
'							objDateControl.JavaCalendar("Time").SetTime aDate(1)							
'							objDateControl.JavaButton("OK").Click
'							Fn_ADS_ViewerTabOperation = True
'							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_ADS_ViewerTabOperation :End Date "+dicViewerTab("EndDate")+" set successfully")
'					End If
'					Call Fn_ReadyStatusSync(1)
'
'					'Set Offset
'					If dicViewerTab("Offset") <>  "" then
'							Set WshShell =CreateObject("WScript.Shell")
'							JavaWindow("ADS-TeamCenter").JavaTable("Event List:").ActivateCell iRows,"Offset"
'							WshShell.SendKeys( dicViewerTab("Offset"))
'							Fn_ADS_ViewerTabOperation = True
'							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_ADS_ViewerTabOperation :Offset "+dicViewerTab("Offset") +"set successfully")
'					End If
'					Call Fn_ReadyStatusSync(1)
'
'					'Set Relative To
'					If dicViewerTab("RelativeTo") <>  "" then
'						ObjEventTable.ClickCell iRows ,"Relative To"
'                    	ObjWin.JavaList("RelativeTo").Select dicViewerTab("RelativeTo") 
'						Fn_ADS_ViewerTabOperation = True
'						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_ADS_ViewerTabOperation :Relative To value "+dicViewerTab("RelativeTo") +"set successfully")
'					End If
'					Call Fn_ReadyStatusSync(1)
'
'						'For Recurrence
'					If dicViewerTab("Recurrence") <>  "" then
'						ObjEventTable.SelectRow iRows 
'						ObjEventTable.ClickCell iRows ,"Recurrence"
'						ObjWin.JavaList("Recurrence").Select dicViewerTab("Recurrence") 
'                        Fn_ADS_ViewerTabOperation = True
'						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_ADS_ViewerTabOperation : Recurrence value "+dicViewerTab("Recurrence") +" set successfully")
'					End If
'						Call Fn_ReadyStatusSync(1)
'
'					'For Recurrence Date
'					If dicViewerTab("RecurrenceEndDate") <>  "" then
'							ObjEventTable.SelectRow iRows 
'							ObjEventTable.ActivateCell iRows ,"Recurrence End Date"
'							aDate=Split(dicViewerTab("RecurrenceEndDate") ," ",-1,1)
'							Set objDateControl=JavaWindow("ADS-TeamCenter").JavaWindow("Select Date")
'							objDateControl.JavaCalendar("Date").SetDate aDate(0)
'							objDateControl.JavaCalendar("Time").SetTime aDate(1)							
'							objDateControl.JavaButton("OK").Click
'							Fn_ADS_ViewerTabOperation = True
'							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_ADS_ViewerTabOperation :Recurrence End Date "+dicViewerTab("RecurrenceEndDate") +" set successfully")
'					End If
'							Call Fn_ReadyStatusSync(1)
'
'		 Case "Save"
'            		'To save the values in the Event list Table
'					If dicViewerTab("ButtonName") <> "" Then
'							Call Fn_MenuOperation("Select","File:Save")
'							Fn_ADS_ViewerTabOperation=True
'							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_ADS_ViewerTabOperation :Save "+dicViewerTab("ButtonName") +"done successfully")
'					End If
'						Call Fn_ReadyStatusSync(1)
'
'		 Case "Delete"
'					iRows =  cint(ObjEventTable.GetROProperty("rows"))-1
'                    ObjEventTable.ActivateRow iRows
'					ObjEventTable.SelectRow iRows
'					'Click on Delete button
'					Call Fn_Button_Click("Fn_ADS_ViewerTabOperation",ObjWin,"Delete")
'					Call Fn_ReadyStatusSync(1)
'					Fn_ADS_ViewerTabOperation = True
'					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_ADS_ViewerTabOperation :Successfully Deleted the selected row.")
'					Call Fn_ReadyStatusSync(1)
'
'		 Case "VerifyAutoFill"
'							'Clicking on Add Button
'							Call Fn_Button_Click("Fn_ADS_ViewerTabOperation",ObjWin,"Add")
'							Call Fn_ReadyStatusSync(1)
'							iRows =  cint(ObjEventTable.GetROProperty("rows"))-1
'							ObjEventTable.SelectRow iRows 
'							ObjWin.JavaList("EventName").SetToProperty "Index",Cint (iRows+0)
'							'Selecting Event Name
'							If dicViewerTab("EventName") <>  "" Then
'								ObjEventTable.SelectRow iRows 
'								ObjEventTable.ClickCell iRows ,"Event Name"
'								ObjWin.JavaList("EventName").Select dicViewerTab("EventName")
'								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_ADS_ViewerTabOperation :Event Name "+dicViewerTab("EventName")+" set successfully")
'							End If
'
'							'For Verifying Dates gets AutoFilled, after Selecting Event Name
'							ObjEventTable.SelectRow iRows 
'							If ObjEventTable.GetCellData(iRows,"Start Date")<> ""  AND  ObjEventTable.GetCellData(iRows,"End Date") <> ""  AND  ObjEventTable.GetCellData(iRows,"Recurrence End Date") <> "" then
'								Fn_ADS_ViewerTabOperation = True
'								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_ADS_ViewerTabOperation :Successfully Verified the dates present for Start Date, End Date, Recurrence End Date")
'							Else
'								Fn_ADS_ViewerTabOperation = False
'								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_ADS_ViewerTabOperation :Failed to Verify the dates for Start Date, End Date ,Recurrence End Date")
'							End If
'								Call Fn_ReadyStatusSync(1)
'
'			Case "Verify"
'					iFlag = 0
'							'For Event Name
'							If dicViewerTab("EventName") <>  "" then
'								aEventName = Split(dicViewerTab("EventName"),":",-1,1)
'								iRows =  cint(ObjEventTable.GetROProperty("rows"))-1
'								For iCounter=0 to Ubound(aEventName)
'									For iCount=0 to iRows
'										If Trim(Lcase(aEventName(iCounter))) = Trim(Lcase(ObjEventTable.GetCellData(iCount,0))) Then
'											iFlag = iFlag+1
'											Exit For
'										End If
'									Next
'								Next
'								If iFlag = Ubound(aEventName)+1 Then
'									Fn_ADS_ViewerTabOperation = True
'									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_ADS_ViewerTabOperation :Event value found in table successfully")
'								Else
'									Fn_ADS_ViewerTabOperation = False
'									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_ADS_ViewerTabOperation :Event value not found in table.")
'								End If
'							End If
'								Call Fn_ReadyStatusSync(1)
'
'			Case "UIExist"
'						iCount = 0
'						'Check the existance of Event List Table
'						If JavaWindow("ADS-TeamCenter").JavaTable("Event List:").Exist Then
'							iCount = iCount + 1
'						End If
'
'						'Check the existance Add button
'						If JavaWindow("ADS-TeamCenter").JavaButton("Add").Exist Then
'							iCount = iCount + 1
'						End If
'
'						'Check the existance Delete button
'						If JavaWindow("ADS-TeamCenter").JavaButton("Delete").Exist Then
'							iCount = iCount + 1
'						End If
'
'						' Check the existance for all UI columns
'						If dicViewerTab("UIColumns")<>"" Then
'							aColumns = Split(dicViewerTab("UIColumns"),":",-1,1)
'							iCols = JavaWindow("ADS-TeamCenter").JavaTable("Event List:").GetROProperty("cols")
'							For iCounter=0 to Ubound(aColumns)
'								For intCount=0 to iCols-1
'									If Trim(Lcase(JavaWindow("ADS-TeamCenter").JavaTable("Event List:").GetColumnName(intCount)))= Trim(Lcase(aColumns(iCounter))) Then
'										iCount = iCount + 1
'										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), aColumns(iCounter)&"Successfully found column name")
'										Exit For
'									End If
'								Next
'							Next
'						End If 
'						If iCount = Ubound(aColumns)+4 Then
'							Fn_ADS_ViewerTabOperation = True
'							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: UI Existance successfully verified.")
'						Else
'							Fn_ADS_ViewerTabOperation = False
'							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: UI Existance failed to verify.")
'						End If
'							Call Fn_ReadyStatusSync(1)
'									
'			Case "VerifyData"					
'					If dicViewerTab("UIColumns")<>"" Then
'							aColumns = Split(dicViewerTab("UIColumns"),":",-1,1)
'							aColsData = Split(dicViewerTab("ColsData"),"|",-1,1)
'							iCols = JavaWindow("ADS-TeamCenter").JavaTable("Event List:").GetROProperty("cols")
'							iRows = JavaWindow("ADS-TeamCenter").JavaTable("Event List:").GetROProperty("rows")
'							For iRowCnt=0 to iRows-1
'								iFlag = 0
'								For iCounter=0 to Ubound(aColumns)
'									For intCount=0 to iCols-1
'										If Trim(Lcase(JavaWindow("ADS-TeamCenter").JavaTable("Event List:").GetColumnName(intCount)))= Trim(Lcase(aColumns(iCounter))) Then
'											If Trim(Lcase(JavaWindow("ADS-TeamCenter").JavaTable("Event List:").GetCellData(iRowCnt,intCount))) = Trim(Lcase(aColsData(iCounter))) Then
'													iFlag = iFlag +1
'													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully found row value as "&aColsData(iCounter)&"of column name "&aColumns(iCounter))
'													Exit For
'													If iFlag=Ubound(aColumns)+1 Then
'															Call Fn_WriteLogFile(Environment.Value("TestLogFile"), aColumns(iCounter)&"column does'nt found")
'															Exit For
'													End If									
'											Else
'												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), aColumns(iCounter)&"column does not match with the given data"&aColsData(iCounter))
'												Exit For
'											End If
'										End If
'									Next
'									If iFlag=Ubound(aColumns)+1 OR Trim(Lcase(JavaWindow("ADS-TeamCenter").JavaTable("Event List:").GetCellData(iRowCnt,intCount))) <> Trim(Lcase(aColsData(iCounter))) Then
'										Exit For
'									End If									
'								Next
'								If iFlag=Ubound(aColumns)+1 Then
'									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), aColumns(iCounter)&"column does not exist")
'									Exit For
'								End If									
'							Next
'							If iFlag=Ubound(aColumns)+1 Then
'								Fn_ADS_ViewerTabOperation = iRowCnt
'							Else
'								Fn_ADS_ViewerTabOperation = False
'							End If
'					End If
'						Call Fn_ReadyStatusSync(1)
'							
'		 Case "GetDataList"
'					iCols = JavaWindow("ADS-TeamCenter").JavaTable("Event List:").GetROProperty("cols")
'					For iCount = 0 to iCols-1
'                        aColumns(iCount) = JavaWindow("ADS-TeamCenter").JavaTable("Event List:").GetCellData (dicViewerTab("ColsData"),iCount)
'						sReturn = sReturn & aColumns(iCount) & "|"
'					Next
'					sReturn = Mid(sReturn,1,Len(sReturn)-1)
'					Fn_ADS_ViewerTabOperation = sReturn
'					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_ADS_ViewerTabOperation :Successfully retrived Data")
'					Call Fn_ReadyStatusSync(1)
'
'		End Select
'
'	Else
'		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_ADS_ViewerTabOperation :Viewer Tab Does not Exists")
'		Fn_ADS_ViewerTabOperation = False
'	End If
'
'	'Again restore Viewer tab to default 
'	Call Fn_TabFolder_Operation("DoubleClickTab", "Viewer", "")
'	Call Fn_ReadyStatusSync(1)
'
'	Set ObjEventTable = Nothing
'	Set ObjWin = Nothing
'	Set WshShell =Nothing
'	Set objDateControl=Nothing
'End Function


'Modified Few Case entirely By Shreyas & Avinash 14/08/2012

'----------------------------------------------------------------------------------------------------------------------------------------
'		Function					case				Release				changed done
'Fn_ADS_ViewerTabOperation		Add, VerifyAutoFill		TC 11.2				modified column name to Contract Schedule Tasks, added code to set process template 
'																			modified index property for lists
'----------------------------------------------------------------------------------------------------------------------------------------

Public Function Fn_ADS_ViewerTabOperation(sAction,dicViewerTab)
	GBL_FAILED_FUNCTION_NAME="Fn_ADS_ViewerTabOperation"
   On Error Resume Next
	Dim iCount, iRows,ObjEventTable,ObjWin,bReturn,WshShell,aDate,objDateControl
	Dim iFlag, aEventName,iCounter, aColumns, iCols ,aColsData, iRowCnt, intCount, sReturn
	Dim objSelectType,intNoOfObjects,iIndex
	Set ObjEventTable = JavaWindow("ADS-TeamCenter").JavaTable("Event List:")
	Set ObjWin = JavaWindow("ADS-TeamCenter")
	 
	'Maximize Viewer Tab
	Call Fn_TabFolder_Operation("DoubleClickTab", "Viewer", "")
	Call Fn_ReadyStatusSync(1)

	If JavaWindow("ADS-TeamCenter").Exist Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_ADS_ViewerTabOperation :Viewer Tab Exists")
			Select Case sAction
			Case "Add", "Modify"
					Fn_ADS_ViewerTabOperation = False	

					If sAction <> "Modify"	Then
						Call Fn_Button_Click("Fn_ADS_ViewerTabOperation",ObjWin,"Add")
					End If
					iRows =  cint(ObjEventTable.GetROProperty("rows"))-1
					ObjEventTable.SelectRow iRows 
					'ObjWin.JavaList("EventName").SetToProperty "Index",Cint ((iRows+1)*3)-1
					'ObjWin.JavaList("RelativeTo").SetToProperty "Index", Cint((iRows*2)+Cint(iRows))
					'ObjWin.JavaList("Recurrence").SetToProperty "Index",Cint((iRows*2)+Cint(iRows))+1
					ObjWin.JavaList("EventName").SetToProperty "Index",Cint((iRows+1)*3)+Cint(iRows)
					ObjWin.JavaList("RelativeTo").SetToProperty "Index", Cint(((iRows*3))+Cint(iRows))
					ObjWin.JavaList("Recurrence").SetToProperty "Index",Cint((iRows*3)+Cint(iRows))+1

							'For Event Name list
					If dicViewerTab("EventName") <>  "" Then
'						ObjEventTable.SelectRow iRows 
'						ObjEventTable.ClickCell iRows ,"Contract Schedule Tasks"
'						wait 1
'							 If  ObjWin.JavaList("EventName").Exist(5) Then
'								wait 1
'								ObjWin.JavaList("EventName").Select dicViewerTab("EventName")
'								Fn_ADS_ViewerTabOperation = True
'								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_ADS_ViewerTabOperation :Event Name "+dicViewerTab("EventName")+" set successfully")
'							Else   '*Added by Nilesh 13-Feb-2013
'									Wait 1
'									Call Fn_KeyBoardOperation("SendKeys",dicViewerTab("EventName"))
'									Fn_ADS_ViewerTabOperation = True
'									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_ADS_ViewerTabOperation :Event Name "+dicViewerTab("EventName")+" set successfully")
'							 End If '* End

						ObjEventTable.ActivateCell iRows,"Contract Schedule Tasks"
						Wait 2
						'-------- Tc12.4_2020012703a_Maintenance_PoonamC_04Feb2020 - Get Shell Index -------------------
						Set objSelectType = description.Create()
							objSelectType("Class Name").value = "JavaWindow"
							objSelectType("tagname").value = "Shell"
						Set  intNoOfObjects = JavaWindow("ADS-TeamCenter").ChildObjects(objSelectType)
						iIndex = cint(intNoOfObjects.count())-1
						JavaWindow("ADS-TeamCenter").JavaWindow("Shell").SetTOProperty "index",iIndex
						'-----------------------------------------------------------------------------------------------
						Set objSelectType=description.Create()
							objSelectType("Class Name").value = "JavaList"
							objSelectType("to_class").value = "JavaList"
						Set  intNoOfObjects = JavaWindow("ADS-TeamCenter").JavaWindow("Shell").ChildObjects(objSelectType)
						If intNoOfObjects.count <> 0 Then
							intNoOfObjects(0).Select dicViewerTab("EventName")
							Wait 1
							Call Fn_KeyBoardOperation("SendKey","{TAB}")
						End If
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_ADS_ViewerTabOperation :Event Name "+dicViewerTab("EventName")+" set successfully")
					
					End If
					Call Fn_ReadyStatusSync(1)
					
					'Added code to set value in Summary Task Name column of event table
					If dicViewerTab("SummaryTaskName") <>  "" Then
						ObjEventTable.SelectRow iRows 
						wait 1
'						ObjEventTable.ClickCell iRows ,"Summary Task Name"
						JavaWindow("ADS-TeamCenter").JavaTable("Event List:").ActivateCell iRows,"Summary Task Name"
						wait 1
						Call Fn_KeyBoardOperation("SendKeys",dicViewerTab("SummaryTaskName"))
						wait 1
						Fn_ADS_ViewerTabOperation = True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_ADS_ViewerTabOperation :Summary Task Name "+dicViewerTab("SummaryTaskName")+" set successfully")
			 		End If '* End
					Call Fn_ReadyStatusSync(1)
					

				'Start Date of Event
					If dicViewerTab("StartDate") <>  "" then
							ObjEventTable.SelectRow iRows 
							ObjEventTable.ActivateCell iRows ,"Start Date"
							aDate=Split(dicViewerTab("StartDate") ," ",-1,1)
							If ubound(aDate)=0 Then
								Call Fn_UI_SetDateAndTime("Fn_ADS_ViewerTabOperation",aDate(0), "")
							Else
								Call Fn_UI_SetDateAndTime("Fn_ADS_ViewerTabOperation",aDate(0), aDate(1))
							End If
							Fn_ADS_ViewerTabOperation = True
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_ADS_ViewerTabOperation :Start Date "+dicViewerTab("StartDate")+" set successfully")
					End If
					Call Fn_ReadyStatusSync(1)

				'End Date of Event
					If dicViewerTab("EndDate") <>  "" then
							ObjEventTable.SelectRow iRows 
							ObjEventTable.ActivateCell iRows ,"End Date"
							aDate=Split(dicViewerTab("EndDate") ," ",-1,1)
							If ubound(aDate)=0 Then
								Call Fn_UI_SetDateAndTime("Fn_ADS_ViewerTabOperation",aDate(0), "")
							Else
								Call Fn_UI_SetDateAndTime("Fn_ADS_ViewerTabOperation",aDate(0), aDate(1))
							End If							
							Fn_ADS_ViewerTabOperation = True
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_ADS_ViewerTabOperation :End Date "+dicViewerTab("EndDate")+" set successfully")
					End If
					Call Fn_ReadyStatusSync(1)

					'Set Offset
					If dicViewerTab("Offset") <>  "" then
							Set WshShell =CreateObject("WScript.Shell")
							JavaWindow("ADS-TeamCenter").JavaTable("Event List:").ActivateCell iRows,"Offset"
							WshShell.SendKeys( dicViewerTab("Offset"))
							Fn_ADS_ViewerTabOperation = True
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_ADS_ViewerTabOperation :Offset "+dicViewerTab("Offset") +"set successfully")
					End If
					Call Fn_ReadyStatusSync(1)

					'Set Relative To
					If dicViewerTab("RelativeTo") <>  "" then
'						ObjEventTable.ClickCell iRows ,"Relative To"
'                    	ObjWin.JavaList("RelativeTo").Select dicViewerTab("RelativeTo") 
'						Fn_ADS_ViewerTabOperation = True

						ObjEventTable.ActivateCell iRows,"Relative To"
						Wait 2
						'-------- Tc12.4_2020012703a_Maintenance_PoonamC_04Feb2020 - Get Shell Index -------------------
						Set objSelectType = description.Create()
							objSelectType("Class Name").value = "JavaWindow"
							objSelectType("tagname").value = "Shell"
						Set  intNoOfObjects = JavaWindow("ADS-TeamCenter").ChildObjects(objSelectType)
						iIndex = cint(intNoOfObjects.count())-1
						JavaWindow("ADS-TeamCenter").JavaWindow("Shell").SetTOProperty "index",iIndex
						'-----------------------------------------------------------------------------------------------
						Set objSelectType=description.Create()
							objSelectType("Class Name").value = "JavaList"
							objSelectType("to_class").value = "JavaList"	
						Set  intNoOfObjects = JavaWindow("ADS-TeamCenter").JavaWindow("Shell").ChildObjects(objSelectType)
						If intNoOfObjects.count <> 0 Then
							intNoOfObjects(0).Select dicViewerTab("RelativeTo")
							Wait 1
							Call Fn_KeyBoardOperation("SendKey","{TAB}")
						End If
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_ADS_ViewerTabOperation :Relative To value "+dicViewerTab("RelativeTo") +"set successfully")
					End If
					Call Fn_ReadyStatusSync(1)

						'For Recurrence
					If dicViewerTab("Recurrence") <>  "" then
'						ObjEventTable.SelectRow iRows 
'						ObjEventTable.ClickCell iRows ,"Recurrence"
'						ObjWin.JavaList("Recurrence").Select dicViewerTab("Recurrence") 
'                        Fn_ADS_ViewerTabOperation = True

						ObjEventTable.ActivateCell iRows,"Recurrence"
						Wait 2
						'-------- Tc12.4_2020012703a_Maintenance_PoonamC_04Feb2020 - Get Shell Index -------------------
						Set objSelectType = description.Create()
							objSelectType("Class Name").value = "JavaWindow"
							objSelectType("tagname").value = "Shell"
						Set  intNoOfObjects = JavaWindow("ADS-TeamCenter").ChildObjects(objSelectType)
						iIndex = cint(intNoOfObjects.count())-1
						JavaWindow("ADS-TeamCenter").JavaWindow("Shell").SetTOProperty "index",iIndex
						'-----------------------------------------------------------------------------------------------
						Set objSelectType=description.Create()
							objSelectType("Class Name").value = "JavaList"
							objSelectType("to_class").value = "JavaList"	
						Set  intNoOfObjects = JavaWindow("ADS-TeamCenter").JavaWindow("Shell").ChildObjects(objSelectType)
						If intNoOfObjects.count <> 0 Then
							intNoOfObjects(0).Select dicViewerTab("Recurrence")
							Wait 1
							Call Fn_KeyBoardOperation("SendKey","{TAB}")
						End If
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_ADS_ViewerTabOperation : Recurrence value "+dicViewerTab("Recurrence") +" set successfully")
					End If
						Call Fn_ReadyStatusSync(1)

					'For Recurrence Date
					If dicViewerTab("RecurrenceEndDate") <>  "" then
							ObjEventTable.SelectRow iRows 
							ObjEventTable.ActivateCell iRows ,"Recurrence End Date"
							aDate=Split(dicViewerTab("RecurrenceEndDate") ," ",-1,1)
							If ubound(aDate)=0 Then
								Call Fn_UI_SetDateAndTime("Fn_ADS_ViewerTabOperation",aDate(0), "")
							Else
								Call Fn_UI_SetDateAndTime("Fn_ADS_ViewerTabOperation",aDate(0), aDate(1))
							End If
							Fn_ADS_ViewerTabOperation = True
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_ADS_ViewerTabOperation :Recurrence End Date "+dicViewerTab("RecurrenceEndDate") +" set successfully")
					End If
							Call Fn_ReadyStatusSync(1)
							
					'To set Process template, Task Duration and Align Task as per designed change 
					If dicViewerTab("ProcessTemplate") <> "" Then
						JavaWindow("MyTeamcenter").JavaStaticText("Viewer_Field").SetTOProperty "label","Process Template:"
						JavaWindow("MyTeamcenter").JavaEdit("Viewer_Field").SetTOProperty "attached text","Process Template:"
						JavaWindow("MyTeamcenter").JavaEdit("Viewer_Field").Set dicViewerTab("ProcessTemplate")
						wait 3
						Set WshShell = CreateObject("WScript.Shell")
						WshShell.SendKeys "{ENTER}"
						Set WshShell =nothing
						Fn_ADS_ViewerTabOperation = True
					End If
					
					If dicViewerTab("TaskDurationHour") <> "" Then
						JavaWindow("MyTeamcenter").JavaStaticText("Viewer_Field").SetTOProperty "label","Task Duration Hours:"
						JavaWindow("MyTeamcenter").JavaEdit("Viewer_Field").SetTOProperty "attached text","Task Duration Hours:"
						JavaWindow("MyTeamcenter").JavaEdit("Viewer_Field").Set dicViewerTab("TaskDurationHour")
						wait 1
						Fn_ADS_ViewerTabOperation = True
					End If
					
					If dicViewerTab("AlignTask") <> "" Then
						JavaWindow("MyTeamcenter").JavaStaticText("Viewer_Field").SetTOProperty "label","Align Task:"
						JavaWindow("MyTeamcenter").JavaEdit("Viewer_Field").SetTOProperty "attached text","Align Task:"
						JavaWindow("MyTeamcenter").JavaEdit("Viewer_Field").Set dicViewerTab("AlignTask")
						wait 2
				        Set WshShell = CreateObject("WScript.Shell")
						WshShell.SendKeys "{ENTER}"
						Set WshShell =nothing
						Fn_ADS_ViewerTabOperation = True
					End If		

		 Case "Save"
            		'To save the values in the Event list Table
					If dicViewerTab("ButtonName") <> "" Then
							Call Fn_MenuOperation("Select","File:Save")
							Fn_ADS_ViewerTabOperation=True
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_ADS_ViewerTabOperation :Save "+dicViewerTab("ButtonName") +"done successfully")
					End If
						Call Fn_ReadyStatusSync(1)

		 Case "Delete"
		            iRows =  cint(ObjEventTable.GetROProperty("rows"))-1
		            ObjEventTable.ActivateCell iRows ,"Summary Task Name"
		            wait 1
'				    ObjEventTable.ActivateRow iRows
'					ObjEventTable.SelectRow iRows
					'Click on Delete button
					Call Fn_Button_Click("Fn_ADS_ViewerTabOperation",ObjWin,"Delete")
					Call Fn_ReadyStatusSync(1)
					Fn_ADS_ViewerTabOperation = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_ADS_ViewerTabOperation :Successfully Deleted the selected row.")
					Call Fn_ReadyStatusSync(1)

		 Case "VerifyAutoFill"
							'Clicking on Add Button
							Call Fn_Button_Click("Fn_ADS_ViewerTabOperation",ObjWin,"Add")
							Call Fn_ReadyStatusSync(1)
							iRows =  cint(ObjEventTable.GetROProperty("rows"))-1
							ObjEventTable.SelectRow iRows 
							'ObjWin.JavaList("EventName").SetToProperty "Index",Cint ((iRows+1)*3)-1
							ObjWin.JavaList("EventName").SetToProperty "Index",Cint ((iRows+1)*3)+iRows
							If dicViewerTab("EventName") <>  "" Then
								'ObjEventTable.SelectRow iRows 
								'ObjEventTable.ClickCell iRows ,"Contract Schedule Tasks"
								'wait 1
								'ObjWin.JavaList("EventName").Select dicViewerTab("EventName")
								
								ObjEventTable.ActivateCell iRows,"Contract Schedule Tasks"
								Wait 2
								Set objSelectType=description.Create()
									objSelectType("Class Name").value = "JavaList"
									objSelectType("to_class").value = "JavaList"
								Set  intNoOfObjects = JavaWindow("MyTeamcenter").JavaWindow("Shell").ChildObjects(objSelectType)
								If intNoOfObjects.count <> 0 Then
									intNoOfObjects(0).Select dicViewerTab("EventName")
									Wait 1
									Call Fn_KeyBoardOperation("SendKey","{TAB}")
								Else
									Set intNoOfObjects = JavaWindow("ADS-TeamCenter").JavaWindow("Shell1").ChildObjects(objSelectType)
									intNoOfObjects(0).Select dicViewerTab("EventName")
									Wait 1
									Call Fn_KeyBoardOperation("SendKey","{TAB}")
								End If
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_ADS_ViewerTabOperation :Event Name "+dicViewerTab("EventName")+" set successfully")
							End If

							'For Verifying Dates gets AutoFilled, after Selecting Event Name
							ObjEventTable.SelectRow iRows 
							If ObjEventTable.GetCellData(iRows,"Start Date")<> ""  AND  ObjEventTable.GetCellData(iRows,"End Date") <> ""  AND  ObjEventTable.GetCellData(iRows,"Recurrence End Date") <> "" then
								Fn_ADS_ViewerTabOperation = True
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_ADS_ViewerTabOperation :Successfully Verified the dates present for Start Date, End Date, Recurrence End Date")
							Else
								Fn_ADS_ViewerTabOperation = False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_ADS_ViewerTabOperation :Failed to Verify the dates for Start Date, End Date ,Recurrence End Date")
							End If
								Call Fn_ReadyStatusSync(1)

			Case "Verify"
					iFlag = 0
							'For Event Name
							If dicViewerTab("EventName") <>  "" then
								aEventName = Split(dicViewerTab("EventName"),":",-1,1)
								iRows =  cint(ObjEventTable.GetROProperty("rows"))-1
								For iCounter=0 to Ubound(aEventName)
									For iCount=0 to iRows
										If Trim(Lcase(aEventName(iCounter))) = Trim(Lcase(ObjEventTable.GetCellData(iCount,1))) Then
											iFlag = iFlag+1
											Exit For
										End If
									Next
								Next
								If iFlag = Ubound(aEventName)+1 Then
									Fn_ADS_ViewerTabOperation = True
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_ADS_ViewerTabOperation :Event value found in table successfully")
								Else
									Fn_ADS_ViewerTabOperation = False
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_ADS_ViewerTabOperation :Event value not found in table.")
								End If
							End If
								Call Fn_ReadyStatusSync(1)

			Case "UIExist"
						iCount = 0
						'Check the existance of Event List Table
						If JavaWindow("ADS-TeamCenter").JavaTable("Event List:").Exist Then
							iCount = iCount + 1
						End If

						'Check the existance Add button
						If JavaWindow("ADS-TeamCenter").JavaButton("Add").Exist Then
							iCount = iCount + 1
						End If

						'Check the existance Delete button
						If JavaWindow("ADS-TeamCenter").JavaButton("Delete").Exist Then
							iCount = iCount + 1
						End If

						' Check the existance for all UI columns
						If dicViewerTab("UIColumns")<>"" Then
							aColumns = Split(dicViewerTab("UIColumns"),":",-1,1)
							iCols = JavaWindow("ADS-TeamCenter").JavaTable("Event List:").GetROProperty("cols")
							For iCounter=0 to Ubound(aColumns)
								For intCount=0 to iCols-1
									If Trim(Lcase(JavaWindow("ADS-TeamCenter").JavaTable("Event List:").GetColumnName(intCount)))= Trim(Lcase(aColumns(iCounter))) Then
										iCount = iCount + 1
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), aColumns(iCounter)&"Successfully found column name")
										Exit For
									End If
								Next
							Next
						End If 
						If iCount = Ubound(aColumns)+4 Then
							Fn_ADS_ViewerTabOperation = True
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: UI Existance successfully verified.")
						Else
							Fn_ADS_ViewerTabOperation = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: UI Existance failed to verify.")
						End If
							Call Fn_ReadyStatusSync(1)
									
			Case "VerifyData"					
					If dicViewerTab("UIColumns")<>"" Then
							aColumns = Split(dicViewerTab("UIColumns"),":",-1,1)
							aColsData = Split(dicViewerTab("ColsData"),"|",-1,1)
							iCols = JavaWindow("ADS-TeamCenter").JavaTable("Event List:").GetROProperty("cols")
							iRows = JavaWindow("ADS-TeamCenter").JavaTable("Event List:").GetROProperty("rows")
							For iRowCnt=0 to iRows-1
								iFlag = 0
								For iCounter=0 to Ubound(aColumns)
									For intCount=0 to iCols-1
										If Trim(Lcase(JavaWindow("ADS-TeamCenter").JavaTable("Event List:").GetColumnName(intCount)))= Trim(Lcase(aColumns(iCounter))) Then
											If Trim(Lcase(JavaWindow("ADS-TeamCenter").JavaTable("Event List:").GetCellData(iRowCnt,intCount))) = Trim(Lcase(aColsData(iCounter))) Then
													iFlag = iFlag +1
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully found row value as "&aColsData(iCounter)&"of column name "&aColumns(iCounter))
													Exit For
													If iFlag=Ubound(aColumns)+1 Then
															Call Fn_WriteLogFile(Environment.Value("TestLogFile"), aColumns(iCounter)&"column does'nt found")
															Exit For
													End If									
											Else
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), aColumns(iCounter)&"column does not match with the given data"&aColsData(iCounter))
												Exit For
											End If
										End If
									Next
									If iFlag=Ubound(aColumns)+1 OR Trim(Lcase(JavaWindow("ADS-TeamCenter").JavaTable("Event List:").GetCellData(iRowCnt,intCount))) <> Trim(Lcase(aColsData(iCounter))) Then
										Exit For
									End If									
								Next
								If iFlag=Ubound(aColumns)+1 Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), aColumns(iCounter)&"column does not exist")
									Exit For
								End If									
							Next
							If iFlag=Ubound(aColumns)+1 Then
								Fn_ADS_ViewerTabOperation = iRowCnt
							Else
								Fn_ADS_ViewerTabOperation = False
							End If
					End If
						Call Fn_ReadyStatusSync(1)
							
		 Case "GetDataList"
					iCols = JavaWindow("ADS-TeamCenter").JavaTable("Event List:").GetROProperty("cols")
					For iCount = 0 to iCols-1
                        aColumns(iCount) = JavaWindow("ADS-TeamCenter").JavaTable("Event List:").GetCellData (dicViewerTab("ColsData"),iCount)
						sReturn = sReturn & aColumns(iCount) & "|"
					Next
					sReturn = Mid(sReturn,1,Len(sReturn)-1)
					Fn_ADS_ViewerTabOperation = sReturn
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_ADS_ViewerTabOperation :Successfully retrived Data")
					Call Fn_ReadyStatusSync(1)

		End Select

	Else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_ADS_ViewerTabOperation :Viewer Tab Does not Exists")
		Fn_ADS_ViewerTabOperation = False
	End If

	'Again restore Viewer tab to default 
	Call Fn_TabFolder_Operation("DoubleClickTab", "Viewer", "")
	Call Fn_ReadyStatusSync(1)

	Set ObjEventTable = Nothing
	Set ObjWin = Nothing
	Set WshShell =Nothing
	Set objDateControl=Nothing
End Function

'***********************************************************************************************************************************************************************************
'Function Name		         :	       Fn_ADS_SetDateViewerTab()

'Description			   :		   This function is used to Set Date underViewerTab Dialog 

'Parameters			  :	 				1.   sAction - Action need to perform
'														"SetDate",
'
'											  		2.  iYear  : To set Year 3. strMonth : To Set Month  4. strDate : To Set Date  5.sButtonName : To Click on the Button of Select Date Dialog
'											    										
'Return Value		            :		True / False

'Pre-requisite			   :		  Viewer Tab should be Diaplayed
'
'Examples				   :	Call Fn_ADS_SetDateViewerTab("SetStartDate","2010","October","25"."OK")	
'											 
'History:
'						Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'					   	    					    Govind Singh	20-Oct-2010															Harshal Agarwal
'***********************************************************************************************************************************************************************************		
Public Function Fn_ADS_SetDateViewerTab(sAction, iYear, strMonth, strDate, sButtonName)
	GBL_FAILED_FUNCTION_NAME="Fn_ADS_SetDateViewerTab"
   'Declaring Variables
    Dim iRows, iCount, ObjEventTable, ObjButton
Set ObjEventTable = JavaWindow("ADS-TeamCenter").JavaWindow("ADSWebEmbeddedFrame").JavaTable("Event List:")
Set ObjButton = JavaDialog("Select Date")	
If JavaWindow("ADS-TeamCenter").JavaWindow("ADSWebEmbeddedFrame").Exist Then
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_ADS_SetDateViewerTab :Viewer Tab Exists")

		Select Case sAction
			Case "SetStartDate"
				iRows =  cint(ObjEventTable.GetROProperty("rows"))-1
				ObjEventTable.SelectRow iRows 
				'Clicking onto the Start Date Cell
				ObjEventTable.ClickCell iRows ,"Start Date"
				JavaDialog("Select Date").Activate
				'Setting the value for month.
				JavaDialog("Select Date").JavaStaticText("MonthName").SetTOProperty "label",strMonth
				'Writing Success Log
				Call Fn_WriteLogFile (Environment.Value("TestLogFile") ,"PASS: Successfully Set [" + Cstr(strMonth) + "] from the list")
				For iCount=0 To 11
						If JavaDialog("Select Date").JavaStaticText("MonthName").Exist(2) Then
							Exit For
						End If	
					JavaDialog("Select Date").JavaButton("JavaButton").Click					
				Next
				'Setting the value for year.
				JavaDialog("Select Date").JavaEdit("Year").SetTOProperty "attached text",strMonth
				JavaDialog("Select Date").JavaEdit("Year").Set iYear
				JavaDialog("Select Date").JavaEdit("Year").Activate
				'Writing Success Log
				Call Fn_WriteLogFile (Environment.Value("TestLogFile") ,"PASS: Successfully Set [" + Cstr(iYear) + "] from the list")				
				'Setting the value for Date.
				JavaDialog("Select Date").JavaCheckBox("CheckboxDate").SetTOProperty "attached text",strDate
				JavaDialog("Select Date").JavaCheckBox("CheckboxDate").Set "ON"
				wait(3)
				Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Successfully Set [" + Cstr(strDate) + "] from the list")
				'Clicking on Button from the list
				Call Fn_Button_Click("Fn_ADS_Set _ViewerTabDate",ObjButton,sButtonName)
				Wait 2
				If JavaWindow("ADS-TeamCenter").Dialog("Eventlist").Exist(5) Then
						JavaWindow("ADS-TeamCenter").Dialog("Eventlist").WinButton("OK").Click
				End If
				Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Successfully Clicked on  [" + Cstr(sButtonName) + "] from the list")				
				Fn_ADS_SetDateViewerTab = True
				 
			Case "SetEndDate"
				iRows =  cint(ObjEventTable.GetROProperty("rows"))-1
				ObjEventTable.SelectRow iRows 
				'Clicking onto the Start Date Cell
				ObjEventTable.ClickCell iRows ,"End Date"
				JavaDialog("Select Date").Activate
				'Setting the value for month.
				JavaDialog("Select Date").JavaStaticText("MonthName").SetTOProperty "label",strMonth
				'Writing Success Log
				Call Fn_WriteLogFile (Environment.Value("TestLogFile") ,"PASS: Successfully Set [" + Cstr(strMonth) + "] from the list")
				For iCount=0 To 11
						If JavaDialog("Select Date").JavaStaticText("MonthName").Exist(2) Then
							Exit For
						End If	
					JavaDialog("Select Date").JavaButton("JavaButton").Click					
				Next
				'Setting the value for year.
				JavaDialog("Select Date").JavaEdit("Year").SetTOProperty "attached text",strMonth
				JavaDialog("Select Date").JavaEdit("Year").Set iYear
				JavaDialog("Select Date").JavaEdit("Year").Activate
				'Writing Success Log
				Call Fn_WriteLogFile (Environment.Value("TestLogFile") ,"PASS: Successfully Set [" + Cstr(iYear) + "] from the list")				
				'Setting the value for Date.
				JavaDialog("Select Date").JavaCheckBox("CheckboxDate").SetTOProperty "attached text",strDate
				JavaDialog("Select Date").JavaCheckBox("CheckboxDate").Set "ON"
				wait(3)
				Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Successfully Set [" + Cstr(strDate) + "] from the list")
				'Clicking on Button from the list
				Call Fn_Button_Click("Fn_ADS_Set _ViewerTabDate",ObjButton,sButtonName)
				Wait 2
				If JavaWindow("ADS-TeamCenter").Dialog("Eventlist").Exist(5) Then
						JavaWindow("ADS-TeamCenter").Dialog("Eventlist").WinButton("OK").Click
				End If
				Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Successfully Clicked on  [" + Cstr(sButtonName) + "] from the list")				
				Fn_ADS_SetDateViewerTab = True
								
			Case "RecurrenceEndDate"
				iRows =  cint(ObjEventTable.GetROProperty("rows"))-1
				ObjEventTable.SelectRow iRows 
				'Clicking onto the Start Date Cell
				ObjEventTable.ClickCell iRows ,"Recurrence End Date"
				JavaDialog("Select Date").Activate
				'Setting the value for month.
				JavaDialog("Select Date").JavaStaticText("MonthName").SetTOProperty "label",strMonth
				'Writing Success Log
				Call Fn_WriteLogFile (Environment.Value("TestLogFile") ,"PASS: Successfully Set [" + Cstr(strMonth) + "] from the list")
				For iCount=0 To 11
						If JavaDialog("Select Date").JavaStaticText("MonthName").Exist(2) Then
							Exit For
						End If	
					JavaDialog("Select Date").JavaButton("JavaButton").Click					
				Next
				'Setting the value for year.
				JavaDialog("Select Date").JavaEdit("Year").SetTOProperty "attached text",strMonth
				JavaDialog("Select Date").JavaEdit("Year").Set iYear
				JavaDialog("Select Date").JavaEdit("Year").Activate
				'Writing Success Log
				Call Fn_WriteLogFile (Environment.Value("TestLogFile") ,"PASS: Successfully Set [" + Cstr(iYear) + "] from the list")				
				'Setting the value for Date.
				JavaDialog("Select Date").JavaCheckBox("CheckboxDate").SetTOProperty "attached text",strDate
				JavaDialog("Select Date").JavaCheckBox("CheckboxDate").Set "ON"
				wait(3)
				Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Successfully Set [" + Cstr(strDate) + "] from the list")
				'Clicking on Button from the list
				Call Fn_Button_Click("Fn_ADS_Set _ViewerTabDate",ObjButton,sButtonName)
				Wait 2
				If JavaWindow("ADS-TeamCenter").Dialog("Eventlist").Exist(5) Then
						JavaWindow("ADS-TeamCenter").Dialog("Eventlist").WinButton("OK").Click
				End If
				Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Successfully Clicked on  [" + Cstr(sButtonName) + "] from the list")				
				Fn_ADS_SetDateViewerTab = True
End Select	
	Else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_ADS_ViewerTabOperation :Viewer TabDoes not Exists")
				Fn_ADS_SetDateViewerTab = False
			End If
		Set ObjEventTable = Nothing
		Set ObjButton = Nothing
End Function
''===================================Modified function for Date Controls in Viewer panel ====================================
'dicViewerTab("StartDate") ="")="01-Jun-2011 08:00:00"
'Example: Fn_ADS_SetDateViewerTab_Ext("SetStartDate", dicViewerTab)

Public Function Fn_ADS_SetDateViewerTab_Ext(sAction, dicViewerTab)
	GBL_FAILED_FUNCTION_NAME="Fn_ADS_SetDateViewerTab_Ext"
   'Declaring Variables
    Dim iRows, ObjEventTable, aDate, objDateControl
Set ObjEventTable = JavaWindow("ADS-TeamCenter").JavaTable("Event List:")
	If ObjEventTable.Exist Then
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_ADS_SetDateViewerTab :Viewer Tab Exists")

	'Maximize Viewer Tab
	Call Fn_TabFolder_Operation("DoubleClickTab", "Viewer", "")
	Call Fn_ReadyStatusSync(1)

		Select Case sAction
			Case "SetStartDate"
				If dicViewerTab("StartDate") <>  "" then
							iRows =  cint(ObjEventTable.GetROProperty("rows"))-1
							ObjEventTable.SelectRow iRows 
							ObjEventTable.ActivateCell iRows ,"Start Date"
							aDate=Split(dicViewerTab("StartDate") ," ",-1,1)
							'Modified by Anumol P on 22-Feb-2013
							'Set objDateControl=JavaWindow("ADS-TeamCenter").JavaWindow("Select Date")
							'objDateControl.JavaCalendar("Date").SetDate aDate(0)
							'objDateControl.JavaCalendar("Time").SetTime aDate(1)							
							'objDateControl.JavaButton("OK").Click
                            If ubound(aDate)=0 Then
								Call Fn_UI_SetDateAndTime("Fn_ADS_ViewerTabOperation",aDate(0), "")
							Else
								Call Fn_UI_SetDateAndTime("Fn_ADS_ViewerTabOperation",aDate(0), aDate(1))
							End If
							Fn_ADS_SetDateViewerTab_Ext = True
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_ADS_SetDateViewerTab_Ext :Start Date "+dicViewerTab("StartDate")+" set successfully")
					End If
					Call Fn_ReadyStatusSync(1)

			Case "SetEndDate"
				
				If dicViewerTab("EndDate") <>  "" then
							iRows =  cint(ObjEventTable.GetROProperty("rows"))-1
							ObjEventTable.SelectRow iRows 
							ObjEventTable.ActivateCell iRows ,"End Date"
							aDate=Split(dicViewerTab("EndDate") ," ",-1,1)
'							Modified by Anumol P on 22-Feb-2013
							'Set objDateControl=JavaWindow("ADS-TeamCenter").JavaWindow("Select Date")
							'objDateControl.JavaCalendar("Date").SetDate aDate(0)
							'objDateControl.JavaCalendar("Time").SetTime aDate(1)							
							'objDateControl.JavaButton("OK").Click
                            If ubound(aDate)=0 Then
								Call Fn_UI_SetDateAndTime("Fn_ADS_ViewerTabOperation",aDate(0), "")
							Else
								Call Fn_UI_SetDateAndTime("Fn_ADS_ViewerTabOperation",aDate(0), aDate(1))
							End If
							Fn_ADS_SetDateViewerTab_Ext = True
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_ADS_SetDateViewerTab_Ext :End Date "+dicViewerTab("EndDate")+" set successfully")
					End If
					Call Fn_ReadyStatusSync(1)

								
			Case "RecurrenceEndDate"
				If dicViewerTab("RecurrenceEndDate") <>  "" then
							iRows =  cint(ObjEventTable.GetROProperty("rows"))-1
							ObjEventTable.SelectRow iRows 
							ObjEventTable.ActivateCell iRows ,"Recurrence End Date"
							aDate=Split(dicViewerTab("RecurrenceEndDate") ," ",-1,1)
'							Modified by Anumol P on 22-Feb-2013
'							Set objDateControl=JavaWindow("ADS-TeamCenter").JavaWindow("Select Date")
'							objDateControl.JavaCalendar("Date").SetDate aDate(0)
'							objDateControl.JavaCalendar("Time").SetTime aDate(1)							
'							objDateControl.JavaButton("OK").Click
							 If ubound(aDate)=0 Then
								Call Fn_UI_SetDateAndTime("Fn_ADS_ViewerTabOperation",aDate(0), "")
							Else
								Call Fn_UI_SetDateAndTime("Fn_ADS_ViewerTabOperation",aDate(0), aDate(1))
							End If
							Fn_ADS_SetDateViewerTab_Ext = True
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_ADS_SetDateViewerTab_Ext :Recurrence End Date "+dicViewerTab("RecurrenceEndDate") +" set successfully")
					End If
							Call Fn_ReadyStatusSync(1)

		End Select	
	Else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_ADS_SetDateViewerTab_Ext :Viewer TabDoes not Exists")
		Fn_ADS_SetDateViewerTab_Ext = False
	End If

'	'Again restore Viewer tab to default 
'	Call Fn_TabFolder_Operation("DoubleClickTab", "Viewer", "")
'	Call Fn_ReadyStatusSync(1)

	Set ObjEventTable = Nothing
End Function

'***********************************************************************************************************************************************************************************
'Function Name		         :	       Fn_ADS_PropertiesOperations()

'Description			   :		   This function is used to Set Date underViewerTab Dialog 

'Parameters			  :	 			sAction,sProperty,sValue,sButtons
'											    										
'Return Value		            :		True / False

'Pre-requisite			   :		  Properties Dialog Should Exist
'
'Examples				   :	Call Fn_ADS_PropertiesOperations("Verify","ProductSystemsList","Fuel","")	
'									Call Fn_ADS_PropertiesOperations("ClickAll","","All","")
'									Call Fn_ADS_PropertiesOperations("Verify", "Correspondences", "000022/A;1-Corr1~000023/A;1-Corr2~000024/A;1-Corr3","Cancel")
'									Call Fn_ADS_PropertiesOperations("Verify", "Contract Has Correspondence", "000022/A;1-Corr1~000023/A;1-Corr2~000024/A;1-Corr3","Cancel")
'									Call Fn_ADS_PropertiesOperations("Verify", "Contracts", "000022/A;1-Con1","Cancel")
'									Call Fn_ADS_PropertiesOperations("RetrieveValue_JavaList", "Schedule Deliverables", "","Cancel")
'									Call Fn_ADS_PropertiesOperations("Verify_JavaEdit", "Note Text", "","Cancel")		
'									Call Fn_ADS_PropertiesOperations("Edit_JavaEdit", "Note Text", "xyz","Save and Check-In")		
'											 
'History:
'						Developer Name			Date				Rev. No.			Changes Done					Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'					   	Harshal Agrawal		21-Oct-2010																	Harshal Agarwal
'						Sushma Pagare		16-May-2011					Added Case "Correspondences", 					Harshal Agarwal
'																		"Contract Has Correspondence", "Contracts"	
'						Sushma Pagare		21-May-2011					Added Case "RetrieveValue_JavaList", 			Harshal Agarwal
'						Sushma Pagare		21-May-2011					Added Case "Verify_JavaEdit", 					Harshal Agarwal
'						Sushma Pagare		26-May-2011					Added Case "Edit_JavaEdit", 											Harshal Agarwal
'***********************************************************************************************************************************************************************************
'Public Function Fn_ADS_PropertiesOperations(sAction,sProperty,sValue,sButtons)
'  Dim OjProp,iCounter,aButtons, ObjJavaList , aValues, iItemsCount, iCounter2, ObjJavaEdit
'   Select Case sAction
' 	Case "Verify"
'		Select Case sProperty
'			Case "ProductSystemsList"
'				Set OjProp = JavaWindow("ADS-TeamCenter").JavaWindow("TcDefaultApplet").JavaDialog("Properties")
'				If OjProp.Exist Then
'					OjProp.JavaCheckBox("EditProductSystem").Set "ON"
'					OjProp.JavaButton("ListDropDown").Click micLeftBtn
'					Wait(3)
''					OjProp.JavaStaticText("ListItems").SetTOProperty "label",sValue
'Dim objSelectType,intNoOfObjects,i,bFlag
'bFlag=false
'											Set objSelectType=description.Create()
'						objSelectType("Class Name").value = "JavaStaticText"
'						Set  intNoOfObjects = OjProp.ChildObjects(objSelectType)
'						  For i = 0 to intNoOfObjects.count-1
'							   If  intNoOfObjects(i).getROProperty("label") = sValue Then
''										intNoOfObjects(i).Click 1,1
'bFlag=True
'										Exit for
'							   End If
'						  Next
''					If OjProp.JavaStaticText("ListItems").Exist Then
''						Fn_ADS_PropertiesOperations = True
''						OjProp.JavaList("ProductSystemsList").Click 5,5,"LEFT"
''						OjProp.JavaCheckBox("EditProductSystem").Set "OFF"
''					Else
''						Fn_ADS_PropertiesOperations = False
''						OjProp.JavaList("ProductSystemsList").Click 5,5,"LEFT"
''						OjProp.JavaCheckBox("EditProductSystem").Set "OFF"
''						Exit Function
''					End If
'				If bFlag=True then
'							Fn_ADS_PropertiesOperations = True
''						OjProp.JavaList("ProductSystemsList").Click 5,5,"LEFT"
'						OjProp.JavaCheckBox("EditProductSystem").Set "OFF"
'					Else
'						Fn_ADS_PropertiesOperations = False
'						OjProp.JavaList("ProductSystemsList").Click 5,5,"LEFT"
'						OjProp.JavaCheckBox("EditProductSystem").Set "OFF"
'						Exit Function
'					End If
'				Else
'						Fn_ADS_PropertiesOperations = False
'						OjProp.JavaList("ProductSystemsList").Click 5,5,"LEFT"
'						OjProp.JavaCheckBox("EditProductSystem").Set "OFF"
'						Exit Function
'					End If
'						'Case Added By Sushma 
'			Case "Correspondences", "Contract Has Correspondence", "Contracts"
'				Set OjProp = JavaWindow("ADS-TeamCenter").JavaWindow("TcDefaultApplet").JavaDialog("Properties")
'				If OjProp.Exist  Then
'						If OjProp.JavaStaticText("BottomLink").Exist Then
'								Call Fn_UI_JavaStaticText_Click("Fn_ADS_PropertiesOperations", OjProp,"BottomLink",0,0,"LEFT")
'						End If
'						If OjProp.JavaStaticText("More...").Exist Then
'							Call Fn_UI_JavaStaticText_Click("Fn_ADS_PropertiesOperations", OjProp,"More...",0,0,"LEFT")
'						End If
'						Set ObjJavaList = OjProp.JavaList("ProductSystemsList")
'						ObjJavaList.SetTOProperty "attached text", sProperty& ":"
'						aValues=Split(sValue,"~",-1,1)
'						For iCounter=0 to ubound(aValues)
'								bFlag=FALSE
'								 iItemsCount = ObjJavaList.GetROProperty("items count")
'								 For iCounter2 = 0 to iItemsCount-1
'										If 	ObjJavaList.GetItem(iCounter2) <> ""	 Then
'											If  Trim(cstr(ObjJavaList.GetItem(iCounter2)))=Trim(cstr(aValues(iCounter)))         Then
'														bFlag=True
'													    Exit For
'											End If
'									    End If
'								 Next
'								 If bFlag = False Then
'										Fn_ADS_PropertiesOperations = False
'										Exit For
'								 End If
'						Next
'						If iCounter >  ubound(aValues) Then
'								Fn_ADS_PropertiesOperations = True
'						Else
'								Fn_ADS_PropertiesOperations = False
'						End If
'				Else
'					Fn_ADS_PropertiesOperations = False
'					Exit Function
'				End If
'		End Select
'	Case "Edit_JavaEdit"  ''Added By Sushma
'				Set OjProp = JavaWindow("ADS-TeamCenter").JavaWindow("TcDefaultApplet").JavaDialog("Properties")
'				If OjProp.Exist  Then
'						If OjProp.JavaStaticText("BottomLink").Exist Then
'								Call Fn_UI_JavaStaticText_Click("Fn_ADS_PropertiesOperations", OjProp,"BottomLink",0,0,"LEFT")
'						End If
'						If OjProp.JavaStaticText("More...").Exist Then
'							Call Fn_UI_JavaStaticText_Click("Fn_ADS_PropertiesOperations", OjProp,"More...",0,0,"LEFT")
'						End If
'						'Set ObjJavaEdit = OjProp.JavaEdit("Name")
'						OjProp.JavaEdit("Name").SetTOProperty "attached text", sProperty & ":"
'						Call Fn_Edit_Box("Fn_ADS_PropertiesOperations",OjProp,"Name",sValue)	
'						Fn_ADS_PropertiesOperations = True
'						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Function [ Fn_ADS_PropertiesOperations ] Successfully Edited ["+sProperty+"] To ["+sValue+"]")
'				Else
'						Fn_ADS_PropertiesOperations = False
'						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Function [ Fn_ADS_PropertiesOperations ] Properties Dialog Does not exist.")
'				End If
'	Case "ClickAll"
'		Set ObjProp = JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Edit Properties")
'			ObjProp.SetTOProperty "title","Properties"
'			If ObjProp.Exist Then				
'				If ObjProp.JavaStaticText("BottomLink").Exist Then
'					Call Fn_UI_JavaStaticText_Click("Fn_ADS_PropertiesOperations", ObjProp,"BottomLink",0,0,"LEFT")
'				End If
'			End If		
'			Fn_ADS_PropertiesOperations = True
'			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Function [ Fn_ADS_PropertiesOperations ] Successfully Clicked on [All].")
'	Case"VerifyPropContractRef"
'		Set ObjProp = JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Edit Properties")
'			If  Fn_UI_ObjectExist("Fn_ADS_PropertiesOperations",ObjProp) = True Then
'				Call Fn_UI_JavaStaticText_Click("Fn_ADS_PropertiesOperations", ObjProp,"BottomLink",0,0,"LEFT")
'			End If
'			If  Fn_UI_ObjectExist("Fn_ADS_PropertiesOperations",ObjProp.JavaEdit("Contract Reference")) = True Then
'				sStatus = ObjProp.JavaEdit("Contract Reference").GetROProperty("editable")
'				If sStatus = "0" Then
'					Fn_ADS_PropertiesOperations = True
'					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Function [ Fn_ADS_PropertiesOperations ] Successfully Verified that Contract Reference editbox is uneditable.")
'					Call Fn_Button_Click("Fn_ADS_PropertiesOperations", ObjProp, "Close")
'				Else
'					Fn_ADS_PropertiesOperations = False
'					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Function [ Fn_ADS_PropertiesOperations ]  Contract Reference is editbox editable.")
'					Exit Function 
'				End If
'			End If
'		Select Case sProperty
'			Case "ProductSystems"
'				'Yet to be Develop
'		End Select
'	Case "RetrieveValue_JavaList"	
'				Set OjProp = JavaWindow("ADS-TeamCenter").JavaWindow("TcDefaultApplet").JavaDialog("Properties")
'				If OjProp.Exist  Then
'						If OjProp.JavaStaticText("BottomLink").Exist Then
'								Call Fn_UI_JavaStaticText_Click("Fn_ADS_PropertiesOperations", OjProp,"BottomLink",0,0,"LEFT")
'						End If
'						If OjProp.JavaStaticText("More...").Exist Then
'							Call Fn_UI_JavaStaticText_Click("Fn_ADS_PropertiesOperations", OjProp,"More...",0,0,"LEFT")
'						End If
'						Set ObjJavaList = OjProp.JavaList("ProductSystemsList")
'						ObjJavaList.SetTOProperty "attached text", sProperty& ":"
'						 iItemsCount = ObjJavaList.GetROProperty("items count")
'						 sValue = ""
'						For iCounter = 0 to iItemsCount-1
'								sValue = sValue & ObjJavaList.GetItem(iCounter)
'								If  iCounter <  iItemsCount-1 Then
'										sValue = sValue & ","
'								 End If					
'						Next
'						Fn_ADS_PropertiesOperations = sValue
'				Else
'					Fn_ADS_PropertiesOperations = False
'					Exit Function
'				End If
'	Case "Verify_JavaEdit"	  ''Case Added  By Sushma
'				Set OjProp = JavaWindow("ADS-TeamCenter").JavaWindow("TcDefaultApplet").JavaDialog("Properties")
'				If OjProp.Exist  Then
'						If OjProp.JavaStaticText("BottomLink").Exist Then
'								Call Fn_UI_JavaStaticText_Click("Fn_ADS_PropertiesOperations", OjProp,"BottomLink",0,0,"LEFT")
'						End If
'						If OjProp.JavaStaticText("More...").Exist Then
'							Call Fn_UI_JavaStaticText_Click("Fn_ADS_PropertiesOperations", OjProp,"More...",0,0,"LEFT")
'						End If
'						Set ObjJavaEdit = OjProp.JavaEdit("Name")
'						ObjJavaEdit.SetTOProperty "attached text", sProperty& ":"
'
'						If cstr(trim(ObjJavaEdit.GetROProperty("value"))) =  cstr(trim(sValue)) Then
'								Fn_ADS_PropertiesOperations = True
'						Else
'								Fn_ADS_PropertiesOperations = False
'						End If
'				Else
'					Fn_ADS_PropertiesOperations = False
'					Exit Function
'				End If	
'   End Select
'   'Click on Buttons
'	 If sButtons<>"" Then
'	   aButtons = split(sButtons, ":",-1,1)
'	   iCounter = Ubound(aButtons)
'	   For iCount=0 to iCounter
'		'Click on Add Button
'		Call Fn_Button_Click("Fn_ADS_PropertiesOperations", OjProp, aButtons(iCount))
'        Call Fn_ReadyStatusSync(2)
'	   Next
'	End If
'   Set OjProp = Nothing
'End Function


Public Function Fn_ADS_PropertiesOperations(sAction,sProperty,sValue,sButtons)
	GBL_FAILED_FUNCTION_NAME="Fn_ADS_PropertiesOperations"
  Dim OjProp,iCounter,aButtons, ObjJavaList , aValues, iItemsCount, iCounter2, ObjJavaEdit
  Dim objSelectType,intNoOfObjects,i,bFlag
  If JavaWindow("ADS-TeamCenter").JavaWindow("TcDefaultApplet").JavaDialog("Properties").Exist(5) Then
	  Set OjProp = JavaWindow("ADS-TeamCenter").JavaWindow("TcDefaultApplet").JavaDialog("Properties")
  Else
  	  Set OjProp = JavaWindow("ADS-TeamCenter").JavaWindow("ADSWebEmbeddedFrame").JavaDialog("Properties")
  End If
   Select Case sAction
 	Case "Verify"
		Select Case sProperty
			Case "ProductSystemsList"
'				Set OjProp = JavaWindow("ADS-TeamCenter").JavaWindow("TcDefaultApplet").JavaDialog("Properties")
				If OjProp.Exist Then
					OjProp.JavaCheckBox("EditProductSystem").Set "ON"
					OjProp.JavaButton("ListDropDown").Click micLeftBtn
					Wait(3)
					bFlag=false
					For i = 0 to Int(OjProp.JavaTable("LOVTreeTable").GetROProperty("rows"))-1
					   If  OjProp.JavaTable("LOVTreeTable").Object.getValueAt(i,0).getDisplayableValue() = sValue Then
							bFlag=True
							Exit for
					   End If
					  Next

				If bFlag=True then
							Fn_ADS_PropertiesOperations = True
'						OjProp.JavaList("ProductSystemsList").Click 5,5,"LEFT"
						OjProp.JavaCheckBox("EditProductSystem").Set "OFF"
					Else
						Fn_ADS_PropertiesOperations = False
						OjProp.JavaList("ProductSystemsList").Click 5,5,"LEFT"
						OjProp.JavaCheckBox("EditProductSystem").Set "OFF"
						Exit Function
					End If
				Else
						Fn_ADS_PropertiesOperations = False
						OjProp.JavaList("ProductSystemsList").Click 5,5,"LEFT"
						OjProp.JavaCheckBox("EditProductSystem").Set "OFF"
						Exit Function
					End If
						'Case Added By Sushma 
			Case "Correspondences", "Contract Has Correspondence", "Contracts"
				'Set OjProp = JavaWindow("ADS-TeamCenter").JavaWindow("TcDefaultApplet").JavaDialog("Properties")
				If OjProp.Exist  Then
						If OjProp.JavaStaticText("BottomLink").Exist Then
								Call Fn_UI_JavaStaticText_Click("Fn_ADS_PropertiesOperations", OjProp,"BottomLink",0,0,"LEFT")
						End If
						If OjProp.JavaStaticText("More...").Exist Then
							Call Fn_UI_JavaStaticText_Click("Fn_ADS_PropertiesOperations", OjProp,"More...",0,0,"LEFT")
						End If
						Set ObjJavaList = OjProp.JavaList("ProductSystemsList")
						ObjJavaList.SetTOProperty "attached text", sProperty& ":"
						aValues=Split(sValue,"~",-1,1)
						For iCounter=0 to ubound(aValues)
								bFlag=FALSE
								 iItemsCount = ObjJavaList.GetROProperty("items count")
								 For iCounter2 = 0 to iItemsCount-1
										If 	ObjJavaList.GetItem(iCounter2) <> ""	 Then
											If  Trim(cstr(ObjJavaList.GetItem(iCounter2)))=Trim(cstr(aValues(iCounter)))         Then
														bFlag=True
													    Exit For
											End If
									    End If
								 Next
								 If bFlag = False Then
										Fn_ADS_PropertiesOperations = False
										Exit For
								 End If
						Next
						If iCounter >  ubound(aValues) Then
								Fn_ADS_PropertiesOperations = True
						Else
								Fn_ADS_PropertiesOperations = False
						End If
				Else
					Fn_ADS_PropertiesOperations = False
					Exit Function
				End If
		End Select
	Case "Edit_JavaEdit","Edit_JavaEdit_Ext"  ''Added By Sushma
				'Set OjProp = JavaWindow("ADS-TeamCenter").JavaWindow("TcDefaultApplet").JavaDialog("Properties")
				If OjProp.Exist  Then
						If OjProp.JavaStaticText("BottomLink").Exist Then
								Call Fn_UI_JavaStaticText_Click("Fn_ADS_PropertiesOperations", OjProp,"BottomLink",0,0,"LEFT")
						End If
						If OjProp.JavaStaticText("More...").Exist Then
							Call Fn_UI_JavaStaticText_Click("Fn_ADS_PropertiesOperations", OjProp,"More...",0,0,"LEFT")
						End If
						'Set ObjJavaEdit = OjProp.JavaEdit("Name")
						OjProp.JavaEdit("Name").SetTOProperty "attached text", sProperty & ":"
						Call Fn_Edit_Box("Fn_ADS_PropertiesOperations",OjProp,"Name",sValue)
						If sAction = "Edit_JavaEdit_Ext" Then
							wait 3
							Call Fn_KeyBoardOperation("SendKey","{TAB}")
						End If
						Fn_ADS_PropertiesOperations = True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Function [ Fn_ADS_PropertiesOperations ] Successfully Edited ["+sProperty+"] To ["+sValue+"]")
				Else
						Fn_ADS_PropertiesOperations = False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Function [ Fn_ADS_PropertiesOperations ] Properties Dialog Does not exist.")
				End If
	Case "ClickAll"
		Set ObjProp = JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Edit Properties")
			ObjProp.SetTOProperty "title","Properties"
			If ObjProp.Exist Then				
				If ObjProp.JavaStaticText("BottomLink").Exist Then
					Call Fn_UI_JavaStaticText_Click("Fn_ADS_PropertiesOperations", ObjProp,"BottomLink",0,0,"LEFT")
				End If
			End If		
			Fn_ADS_PropertiesOperations = True
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Function [ Fn_ADS_PropertiesOperations ] Successfully Clicked on [All].")
	Case"VerifyPropContractRef"
		Set ObjProp = JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Edit Properties")
			If  Fn_UI_ObjectExist("Fn_ADS_PropertiesOperations",ObjProp) = True Then
				Call Fn_UI_JavaStaticText_Click("Fn_ADS_PropertiesOperations", ObjProp,"BottomLink",0,0,"LEFT")
			End If
			If  Fn_UI_ObjectExist("Fn_ADS_PropertiesOperations",ObjProp.JavaEdit("Contract Reference")) = True Then
				sStatus = ObjProp.JavaEdit("Contract Reference").GetROProperty("editable")
				If sStatus = "0" Then
					Fn_ADS_PropertiesOperations = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Function [ Fn_ADS_PropertiesOperations ] Successfully Verified that Contract Reference editbox is uneditable.")
					Call Fn_Button_Click("Fn_ADS_PropertiesOperations", ObjProp, "Close")
				Else
					Fn_ADS_PropertiesOperations = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Function [ Fn_ADS_PropertiesOperations ]  Contract Reference is editbox editable.")
					Exit Function 
				End If
			End If
		Select Case sProperty
			Case "ProductSystems"
				'Yet to be Develop
		End Select
	Case "RetrieveValue_JavaList"	
				'Set OjProp = JavaWindow("ADS-TeamCenter").JavaWindow("TcDefaultApplet").JavaDialog("Properties")
				If OjProp.Exist  Then
						If OjProp.JavaStaticText("BottomLink").Exist Then
								Call Fn_UI_JavaStaticText_Click("Fn_ADS_PropertiesOperations", OjProp,"BottomLink",0,0,"LEFT")
						End If
						If OjProp.JavaStaticText("More...").Exist Then
							Call Fn_UI_JavaStaticText_Click("Fn_ADS_PropertiesOperations", OjProp,"More...",0,0,"LEFT")
						End If
						If sProperty= "Data Requirement Item Has Submittal" Then
							Set ObjJavaList= JavaWindow("ADS-TeamCenter").JavaWindow("TcDefaultApplet").JavaDialog("Properties").JavaList("DataRequirementItemHasSubmittal")
						Else 
							Set ObjJavaList = OjProp.JavaList("ProductSystemsList")
							ObjJavaList.SetTOProperty "attached text", sProperty& ":"				
						End If
						
						 iItemsCount = ObjJavaList.GetROProperty("items count")
						 sValue = ""
						For iCounter = 0 to iItemsCount-1
								sValue = sValue & ObjJavaList.GetItem(iCounter)
								If  iCounter <  iItemsCount-1 Then
										sValue = sValue & ","
								 End If					
						Next
						Fn_ADS_PropertiesOperations = sValue
				Else
					Fn_ADS_PropertiesOperations = False
					Exit Function
				End If
	Case "Verify_JavaEdit"	  ''Case Added  By Sushma
				'Set OjProp = JavaWindow("ADS-TeamCenter").JavaWindow("TcDefaultApplet").JavaDialog("Properties")
				If OjProp.Exist  Then
						If OjProp.JavaStaticText("BottomLink").Exist Then
								Call Fn_UI_JavaStaticText_Click("Fn_ADS_PropertiesOperations", OjProp,"BottomLink",0,0,"LEFT")
						End If
						If OjProp.JavaStaticText("More...").Exist Then
							Call Fn_UI_JavaStaticText_Click("Fn_ADS_PropertiesOperations", OjProp,"More...",0,0,"LEFT")
						End If
						Set ObjJavaEdit = OjProp.JavaEdit("Name")
						ObjJavaEdit.SetTOProperty "attached text", sProperty& ":"

						If cstr(trim(ObjJavaEdit.GetROProperty("value"))) =  cstr(trim(sValue)) Then
								Fn_ADS_PropertiesOperations = True
						Else
								Fn_ADS_PropertiesOperations = False
						End If
				Else
					Fn_ADS_PropertiesOperations = False
					Exit Function
				End If	
   End Select
   'Click on Buttons
	 If sButtons<>"" Then
	   aButtons = split(sButtons, ":",-1,1)
	   iCounter = Ubound(aButtons)
	   For iCount=0 to iCounter
		'Click on Add Button
		Call Fn_Button_Click("Fn_ADS_PropertiesOperations", OjProp, aButtons(iCount))
        Call Fn_ReadyStatusSync(2)
	   Next
	End If
   Set OjProp = Nothing
End Function


'***********************************************************************************************************************************************************************************
'Function Name		         :	       Fn_ADS_ChangeOperationsDic()

'Description			 		  :		   This function is used to Set Edit Box for Change Request  Dialog 

'Parameters					  :	 				1.   sAction - Action need to perform
'											:							"SetCREdit"
'											:	  		2.  dicNewChange
'											    										
'Return Value		     	    :		True / False

'Pre-requisite			  		 :		  
'Examples						:	'dicNewChange("NodeName") = "Change Request"
												'dicNewChange("ECRNo") = "0000012"
												'dicNewChange("Revision")="A"
												'dicNewChange("Synopsis")="ASDFGH"
												'dicNewChange("Desc")= "Test Desc"
												'dicNewChange("ChangeType")="Seesaw"
												'dicNewChange("ButtonName")="Finish"
												'dicNewChange("ButtonName")="Cancel"

		'			   :						Call Fn_ADS_ChangeOperationsDic ("SetCREdit",dicNewChange)
'											 
'History:
'						Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'					   	    					    Govind Singh	21-Oct-2010															Harshal Agarwal
'***********************************************************************************************************************************************************************************	
 Public Function Fn_ADS_ChangeOperationsDic(sAction,dicNewChange)
 	GBL_FAILED_FUNCTION_NAME="Fn_ADS_ChangeOperationsDic"
	Dim ObjDialog, aButtons, iCount, sNewItemMenu
	' Configure variant window.
	Set ObjDialog = Fn_ADS_SISW_GetObject("New Change")
	sNewItemMenu = Fn_GetXMLNodeValue( Fn_LogUtil_GetXMLPath("RAC_Menu"),"NewChange")
	
	'Set ObjDialog = JavaWindow("ADS-TeamCenter").JavaWindow("New Change")
	If Fn_UI_ObjectExist("Fn_ADS_ChangeOperationsDic", ObjDialog) = False Then
		'Operate File:New:Change...menu to invoke required dialog
		Call Fn_MenuOperation("Select", sNewItemMenu)
		Call Fn_ReadyStatusSync(2)
	End If

	If Fn_UI_ObjectExist("Fn_ADS_ChangeOperationsDic", ObjDialog) = True Then
		'Set FilterText
		dicNewChange("Filter") = dicNewChange("NodeName")
		 If  Trim(dicNewChange("Filter")) <> "" then
			Call Fn_UI_EditBox_Type("Fn_ADS_ChangeOperationsDic",ObjDialog,"Filter",dicNewChange("Filter"))
			wait(3)
		   End If
		If  Trim(dicNewChange("NodeName")) <> "" then
				Wait(1)
				'Selecting Node from tree
				Call Fn_JavaTree_Select("Fn_ADS_ChangeOperationsDic", ObjDialog, "ChangeTypeTree","Complete List")
				wait(3)
				Call Fn_JavaTree_Select("Fn_ADS_ChangeOperationsDic", ObjDialog, "ChangeTypeTree","Complete List:"+dicNewChange("NodeName"))
				wait(3)
				'Clicking on Next button to proceed 
				Call Fn_Button_Click("Fn_ADS_ChangeOperationsDic",ObjDialog,"Next")
				wait(3)
		 End If		
Select Case sAction
 Case "SetCREdit"
	 	If  Trim(dicNewChange("ECRNo")) <> "" then
			'Call Fn_UI_EditBox_Type("Fn_ADS_ChangeOperationsDic",ObjDialog,"ECNNo_CN",dicNewChange("ECRNo"))
			JavaWindow("ADS-TeamCenter").JavaWindow("New Change").JavaEdit("ECNNo_CN").Activate
			Call Fn_SISW_UI_JavaEdit_Operations("Fn_ADS_ChangeOperationsDic", "Set",  ObjDialog, "ECNNo_CN", dicNewChange("ECRNo") )
			wait 2
		End If

		If  Trim(dicNewChange("Revision")) <> "" then
			Call Fn_UI_EditBox_Type("Fn_ADS_ChangeOperationsDic",ObjDialog,"Revision_CN",dicNewChange("Revision"))
			wait 1,500
		End If

		If  Trim(dicNewChange("Synopsis")) <> "" then
			Call Fn_UI_EditBox_Type("Fn_ADS_ChangeOperationsDic",ObjDialog,"Synopsis_CN",dicNewChange("Synopsis"))
			wait 1,500
		End If

		If  Trim(dicNewChange("Desc")) <> "" then
			Call Fn_UI_EditBox_Type("Fn_ADS_ChangeOperationsDic",ObjDialog,"Description_CN",dicNewChange("Desc"))
			wait 1,500
		End If

		If  Trim(dicNewChange("ChangeType")) <> "" then
			If dicNewChange("Filter")="Change Notice" or dicNewChange("Filter")="Change Request" Then
				Call Fn_UI_Object_SetTOProperty_ExistCheck("RACUpdateActionItem",JavaWindow("ADS-TeamCenter").JavaWindow("New Change").JavaStaticText("ChangeType"),"label","Change Type:")
			End If
			Call Fn_UI_EditBox_Type("Fn_ADS_ChangeOperationsDic",ObjDialog,"ChangeType_CN",dicNewChange("ChangeType"))
			wait 1,500
		End If

		If  Trim(dicNewChange("ButtonName")) <> "" then
			aButtons = Split(dicNewChange("ButtonName"),":",-1,1)
			For iCount = 0 to Ubound(aButtons)
				Call Fn_Button_Click("Fn_ADS_ChangeOperationsDic", ObjDialog,aButtons(iCount))
				Call Fn_ReadyStatusSync(2)
			Next
		End If
		'Return value.
		Fn_ADS_ChangeOperationsDic = True
End Select
Else
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"New Change Dialog Does not Exists")	
	Fn_ADS_ChangeOperationsDic = False
End If
Set ObjDialog = Nothing
End Function
'***********************************************************************************************************************************************************************************
'Function Name		         :	       Fn_ADS_VerifyChangeObjects()

'Description			   :		   This function is used to Verify Objects Present in Dropdown of the particular Objects.

'Return Value		            :		True / False
'
'Examples				   :	Case "VerifyChangeType" : Call Fn_ADS_VerifyChangeObjects("VerifyChangeType","DCN","Cancel")
'											Case "VerifyChangeType" : Call Fn_ADS_VerifyChangeObjects("VerifyChangeClass","I","Cancel") 
'											Case "VerifyCategory" : Call Fn_ADS_VerifyChangeObjects("VerifyCategory","A","Cancel") 
'History:
'						Developer Name				Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'					   	   Govind Singh	 				20-Oct-2010															Harshal Agarwal
'***********************************************************************************************************************************************************************************	
'Public Function Fn_ADS_VerifyChangeObjects(sAction,sChangeType,sButtons)
'   Dim ObjChangeWnd,iRows,iCount,iCounter,aChangeType, aButtons
'   Fn_ADS_VerifyChangeObjects = False
'	  For iCount=0 to 0
'		 JavaWindow("ADS-TeamCenter").JavaWindow("New Change").SetTOProperty "title","New Change"
'		 If JavaWindow("ADS-TeamCenter").JavaWindow("New Change").Exist(5) Then
'		  Exit For
'		 End If
'		 JavaWindow("ADS-TeamCenter").JavaWindow("New Change").SetTOProperty "title","New Change in context"
'		 If JavaWindow("ADS-TeamCenter").JavaWindow("New Change").Exist(5) Then
'		  Exit For
'		 End If
'		 JavaWindow("ADS-TeamCenter").JavaWindow("New Change").SetTOProperty "title","Derive Change"
'		 If JavaWindow("ADS-TeamCenter").JavaWindow("New Change").Exist(5) Then      
'		  Exit For
'		 End If
'	  Next
' Set ObjChangeWnd=Fn_UI_ObjectCreate("Fn_ADS_ChangeOperationsExtn",JavaWindow("ADS-TeamCenter").JavaWindow("New Change"))
'	If  Fn_UI_ObjectExist("Fn_ADS_VerifyChangeObjects",ObjChangeWnd)=True Then
'		ObjChangeWnd.Maximize
'Select Case sAction
'	 Case "VerifyChangeType","VerifyChangeClass","VerifyCategory"
'		 If sAction ="VerifyChangeType"  Then
'			If sChangeType<>"" Then
'					aChangeType = Split(sChangeType,":","-1","1")	
'					'Clicking on Change Type DropDown Button
''					ObjChangeWnd.JavaButton("DropDownBtn").SetTOProperty "attached text",".*"				
'					ObjChangeWnd.JavaButton("DropDownBtn").SetTOProperty "index",0
'					Call Fn_Button_Click("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd,"DropDownBtn")
'					JavaWindow("ADS-TeamCenter").JavaWindow("New Change").JavaWindow("Shell").SetToProperty "index",0
'				End If
'		 End If
'		If sAction = "VerifyChangeClass" Then
'			 If sChangeType<>"" Then
'					aChangeType = Split(sChangeType,":","-1","1")	
'					'Clicking on Change Class DropDown Button
''					ObjChangeWnd.JavaButton("DropDownBtn").SetTOProperty "attached text",""
'					ObjChangeWnd.JavaButton("DropDownBtn").SetTOProperty "index",1
'					Call Fn_Button_Click("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd,"DropDownBtn")
'					JavaWindow("ADS-TeamCenter").JavaWindow("New Change").JavaWindow("Shell").SetToProperty "index",0
'			End If
'		End If
'			If sAction = "VerifyCategory"  Then
'					aChangeType = Split(sChangeType,":","-1","1")	
'				If sChangeType<>"" Then
'					'Clicking on Category DropDown Button
'					ObjChangeWnd.JavaButton("DropDownBtn").SetTOProperty "index",2
'					Call Fn_Button_Click("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd,"DropDownBtn")
'					JavaWindow("ADS-TeamCenter").JavaWindow("New Change").JavaWindow("Shell").SetToProperty "index",0
'			End If	
'		End If	
'			Wait(3)
''			iRows = JavaWindow("ADS-TeamCenter").JavaWindow("New Change").JavaWindow("Shell").JavaTable("ChangeTable").GetROProperty("rows")
'			iRows = JavaWindow("ADS-TeamCenter").JavaWindow("New Change").JavaWindow("Shell").JavaTable("ChangeTable").GetROProperty("rows")
'			For iCount = 0 to Ubound(aChangeType)	
'				For iCounter= 0 to iRows-1
'				sChange = JavaWindow("ADS-TeamCenter").JavaWindow("New Change").JavaWindow("Shell").JavaTable("ChangeTable").GetCellData(iCounter,0)
'				If Trim(sChange) = Trim(aChangeType(iCount)) Then
'					JavaWindow("ADS-TeamCenter").JavaWindow("New Change").JavaWindow("Shell").JavaTable("ChangeTable").SelectRow iCounter
'					Fn_ADS_VerifyChangeObjects  =True					
'					Exit For
'				End If
'			Next			
'		If Cstr(iCounter) =Cstr( iRows) Then
'			Fn_ADS_VerifyChangeObjects = False
'			Exit Function 
'		 End If
'		 If Cstr(iCounter) =Cstr( iRows) Then
'			Fn_ADS_VerifyChangeObjects = False
'			Exit Function 
'		 End If
'		Next
'      'Click on Buttons
'      If sButtons<>"" Then
'        aButtons = split(sButtons, ":",-1,1)
'        iCounter = Ubound(aButtons)
'        For iCount=0 to iCounter
'         'Click on Buttons
'         Call Fn_Button_Click("Fn_ADS_VerifyChangeObjects", ObjChangeWnd, aButtons(iCount))    
'		 Call Fn_ReadyStatusSync(2)        
'        Next
'      End If
'End Select
'Else 
'	Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"New Change Dialog Does not Exists")
'	Fn_ADS_VerifyChangeObjects = False
'End If
'Set ObjChangeWnd = Nothing
'End Function	


Public Function Fn_ADS_VerifyChangeObjects(sAction,sChangeType,sButtons)
	GBL_FAILED_FUNCTION_NAME="Fn_ADS_VerifyChangeObjects"
   Dim ObjChangeWnd,iRows,iCount,iCounter,aChangeType, aButtons
   Fn_ADS_VerifyChangeObjects = False
	  For iCount=0 to 0
		 JavaWindow("ADS-TeamCenter").JavaWindow("New Change").SetTOProperty "title","New Change"
		 If JavaWindow("ADS-TeamCenter").JavaWindow("New Change").Exist(5) Then
		  Exit For
		 End If
		 JavaWindow("ADS-TeamCenter").JavaWindow("New Change").SetTOProperty "title","New Change in context"
		 If JavaWindow("ADS-TeamCenter").JavaWindow("New Change").Exist(5) Then
		  Exit For
		 End If
		 JavaWindow("ADS-TeamCenter").JavaWindow("New Change").SetTOProperty "title","Derive Change"
		 If JavaWindow("ADS-TeamCenter").JavaWindow("New Change").Exist(5) Then      
		  Exit For
		 End If
	  Next
 Set ObjChangeWnd=Fn_UI_ObjectCreate("Fn_ADS_ChangeOperationsExtn",JavaWindow("ADS-TeamCenter").JavaWindow("New Change"))
	If  Fn_UI_ObjectExist("Fn_ADS_VerifyChangeObjects",ObjChangeWnd)=True Then
		ObjChangeWnd.Maximize
Select Case sAction
	 Case "VerifyChangeType","VerifyChangeClass","VerifyCategory"
		 If sAction ="VerifyChangeType"  Then
			If sChangeType<>"" Then
					aChangeType = Split(sChangeType,":","-1","1")	
					'Clicking on Change Type DropDown Button
'					ObjChangeWnd.JavaButton("DropDownBtn").SetTOProperty "attached text",".*"				
					ObjChangeWnd.JavaButton("DropDownBtn").SetTOProperty "index",0
					Call Fn_Button_Click("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd,"DropDownBtn")
					JavaWindow("ADS-TeamCenter").JavaWindow("New Change").JavaWindow("Shell").SetToProperty "index",0
				End If
		 End If
		If sAction = "VerifyChangeClass" Then
			 If sChangeType<>"" Then
					aChangeType = Split(sChangeType,":","-1","1")	
					'Clicking on Change Class DropDown Button
'					ObjChangeWnd.JavaButton("DropDownBtn").SetTOProperty "attached text",""
					ObjChangeWnd.JavaButton("DropDownBtn").SetTOProperty "index",1
					Call Fn_Button_Click("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd,"DropDownBtn")
					JavaWindow("ADS-TeamCenter").JavaWindow("New Change").JavaWindow("Shell").SetToProperty "index",0
			End If
		End If
			If sAction = "VerifyCategory"  Then
					aChangeType = Split(sChangeType,":","-1","1")	
				If sChangeType<>"" Then
					'Clicking on Category DropDown Button
					ObjChangeWnd.JavaButton("DropDownBtn").SetTOProperty "index",2
					Call Fn_Button_Click("Fn_ADS_ChangeOperationsExtn",ObjChangeWnd,"DropDownBtn")
					JavaWindow("ADS-TeamCenter").JavaWindow("New Change").JavaWindow("Shell").SetToProperty "index",0
			End If	
		End If	
			Wait(3)
'			iRows = JavaWindow("ADS-TeamCenter").JavaWindow("New Change").JavaWindow("Shell").JavaTable("ChangeTable").GetROProperty("rows")
'			iRows = JavaWindow("ADS-TeamCenter").JavaWindow("New Change").JavaWindow("Shell").JavaTable("ChangeTable").GetROProperty("rows")
            iRows = JavaWindow("ADS-TeamCenter").JavaWindow("New Change").JavaWindow("Shell").JavaTree("ChangeTree").GetROProperty("items count")
			For iCount = 0 to Ubound(aChangeType)	
				For iCounter= 0 to iRows-1
'				sChange = JavaWindow("ADS-TeamCenter").JavaWindow("New Change").JavaWindow("Shell").JavaTable("ChangeTable").GetCellData(iCounter,0)
'				If Trim(sChange) = Trim(aChangeType(iCount)) Then
'					JavaWindow("ADS-TeamCenter").JavaWindow("New Change").JavaWindow("Shell").JavaTable("ChangeTable").SelectRow iCounter
'					Fn_ADS_VerifyChangeObjects  =True					
'					Exit For
'				End If
						'*Added by Nilesh on 27-Feb-2013
						sChange = JavaWindow("ADS-TeamCenter").JavaWindow("New Change").JavaWindow("Shell").JavaTree("ChangeTree").GetItem(iCounter)
						If Trim(sChange) = Trim(aChangeType(iCount)) Then
							Fn_ADS_VerifyChangeObjects  =True					
							Exit For
						End If
						'*End
				Next			
			If Cstr(iCounter) =Cstr( iRows) Then
				Fn_ADS_VerifyChangeObjects = False
				Exit Function 
			 End If
			 If Cstr(iCounter) =Cstr( iRows) Then
				Fn_ADS_VerifyChangeObjects = False
				Exit Function 
			 End If
		Next
      'Click on Buttons
      If sButtons<>"" Then
        aButtons = split(sButtons, ":",-1,1)
        iCounter = Ubound(aButtons)
        For iCount=0 to iCounter
         'Click on Buttons
         Call Fn_Button_Click("Fn_ADS_VerifyChangeObjects", ObjChangeWnd, aButtons(iCount))    
		 Call Fn_ReadyStatusSync(2)        
        Next
      End If
End Select
Else 
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"New Change Dialog Does not Exists")
	Fn_ADS_VerifyChangeObjects = False
End If
Set ObjChangeWnd = Nothing
End Function	

''***********************************************************************************************************************************************************************************
''Function Name		         :	       Fn_ADS_ObjectROPropertyCheck()

'Description			   :		   This function is used to Verify Objects Existence

'Return Value		            :		True / False
'
'Examples				   :	Case "DIDTypeEditBoxExists" :Call Fn_ADS_ObjectROPropertyCheck("DIDTypeEditBoxExists","","")
'											Case "ProgarmPhasesEditBoxExists" : Call Fn_ADS_ObjectROPropertyCheck("ProgarmPhasesEditBoxExists",""',"")
'											Case "ProvideContractDeliverablesStatus" : Fn_ADS_ObjectROPropertyCheck("ProvideContractDeliverablesStatus",False,"")
'History:
'						Developer Name				Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'					   	   Govind Singh	 				25-Oct-2010																		Harshal Agarwal
'***********************************************************************************************************************************************************************************	
Public Function Fn_ADS_ObjectROPropertyCheck (sProperty,sValue,sButtons)
	GBL_FAILED_FUNCTION_NAME="Fn_ADS_ObjectROPropertyCheck"
Dim ObjProp,aButtons,iCount,iCounter,sStatus
Set ObjProp =Fn_SISW_GetObject("New Item")
Select Case sProperty
Case "DIDTypeEditBoxExists"
	If  Fn_UI_ObjectExist("Fn_ADS_ObjectROPropertyCheck",ObjProp.JavaEdit("DID Type")) = True Then
				sStatus = ObjProp.JavaEdit("DID Type").GetROProperty("editable")
				If sStatus = "1" Then
					Fn_ADS_ObjectROPropertyCheck = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Function [ Fn_ADS_ObjectROPropertyCheck ] Successfully Verified that Contract Reference editbox is uneditable.")
				Else
					Fn_ADS_ObjectROPropertyCheck = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Function [ Fn_ADS_ObjectROPropertyCheck ]  Contract Reference is editbox editable.")
					Exit Function 
				End If
		End If
	Case "ProgarmPhasesEditBoxExists"
		If  Fn_UI_ObjectExist("Fn_ADS_ObjectROPropertyCheck",ObjProp.JavaEdit("Program Phases")) = True Then
				sStatus = ObjProp.JavaEdit("Program Phases").GetROProperty("editable")
				If sStatus = "1" Then
					Fn_ADS_ObjectROPropertyCheck = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Function [ Fn_ADS_ObjectROPropertyCheck ] Successfully Verified that Contract Reference editbox is uneditable.")
				Else
					Fn_ADS_ObjectROPropertyCheck = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Function [ Fn_ADS_ObjectROPropertyCheck ]  Contract Reference is editbox editable.")
					Exit Function 
				End If
			End If
    Case "Recurring Cost","SOW Affected"    'Case Added by Pallavi J  on 19-Feb-13
			ObjProp.JavaStaticText("ItemLabel").SetTOProperty "label",sProperty+":"
			ObjProp.JavaRadioButton("ItemRadioButton").SetTOProperty "attached text",sValue
			If ObjProp.JavaRadioButton("ItemRadioButton").GetROProperty("value") = 1 Then
				Fn_ADS_ObjectROPropertyCheck = True
			Else
				Fn_ADS_ObjectROPropertyCheck = False
				Exit Function
			End If
	Case "ProvideContractDeliverablesStatus","RecurringCost","SOWEffected"'RadioButton
		If sProperty ="RecurringCost"  Then
			ObjProp.JavaRadioButton("True").SetTOProperty "index" ,0
		ElseIf  sProperty ="SOWEffected"  Then
			ObjProp.JavaRadioButton("True").SetTOProperty "index", 1
		End If
		If sValue = True Then
			If  Fn_UI_ObjectExist("Fn_ADS_ObjectROPropertyCheck",ObjProp.JavaRadioButton("True")) = True Then
				sStatus = ObjProp.JavaRadioButton("True").GetROProperty("value")
				If sStatus = "1" Then
					Fn_ADS_ObjectROPropertyCheck = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Function [ Fn_ADS_ObjectROPropertyCheck ] Successfully Verified that Radio Button is Enabled.")
				Else
					Fn_ADS_ObjectROPropertyCheck = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Function [ Fn_ADS_ObjectROPropertyCheck ]  Radio Button is NOT Enabled")
					Exit Function 
				End If
			End If
			ElseIf sValue =False Then
			If  Fn_UI_ObjectExist("Fn_ADS_ObjectROPropertyCheck",ObjProp.JavaRadioButton("True")) = True Then
				sStatus = ObjProp.JavaRadioButton("True").GetROProperty("value")
				If sStatus = "0" Then
					Fn_ADS_ObjectROPropertyCheck = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Function [ Fn_ADS_ObjectROPropertyCheck ] Successfully Verified that Radio Button is Disabled.")
				Else
					Fn_ADS_ObjectROPropertyCheck = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Function [ Fn_ADS_ObjectROPropertyCheck ]  Radio Button is NOT Disabled")
					Exit Function 
				End If
			End If				
		End If
End Select
	'Click on Buttons
      If sButtons<>"" Then
        aButtons = split(sButtons, ":",-1,1)
        iCounter = Ubound(aButtons)
        For iCount=0 to iCounter
         'Click on Buttons
         Call Fn_Button_Click("Fn_ADS_ObjectROPropertyCheck", ObjProp, aButtons(iCount))
         Call Fn_ReadyStatusSync(2)
		  Fn_ADS_ObjectROPropertyCheck = True         
        Next
      End If
Set ObjProp = Nothing
End Function
'***********************************************************************************************************************************************************************************
'Function Name		         :	       Fn_ADS_TechDocOperations()

'Description			 		  :		   This function is used to Set Edit Box for Change Request  Dialog 

'Return Value		     	    :		ID-RevID / False

'Pre-requisite			  		 :		  
'Examples						:	
												'dicTechDoc("TechType") = "Technical Document"
												'dicTechDoc("TechIDPattern") = +"""01-"""+"NNNN"
												'''dicTechDoc("TechIDPattern") = +"""01-"""+"NNNN"+":"+"""02-"""+"NNNN"       -------------->	(To Verify Multiple Values seperated by " : " {Same For CategoryList and TechDocList To Verify} )
												'dicTechDoc("TechRevID") = "1"  
												'''dicTechDoc("TechRevID") ="-, Secondary Revision."+":"+"1, Initial Revision." -------------> (To Verify Multiple Values seperated by " : " {Same For CategoryList and TechDocList To Verify} )
												'dicTechDoc("TechName") = "Sagar"
												'dicTechDoc("CategoryList") =  "SPEC"
												'dicTechDoc("TechDocList") = "None"
												'dicTechDoc("ButtonName") = "Finish" 
												'dicTechDoc("ButtonName") = "Close"

		'			   :					Call Fn_ADS_TechDocOperations("Create",dicTechDoc,"")
'											 
'History:
'						Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'					   	   Govind Singh				21-Oct-2010																		Harshal Agarwal
'					   	   Sushma Pagare			06-Mar-2013																		
'***********************************************************************************************************************************************************************************	
Public Function Fn_ADS_TechDocOperations(sAction,dicTechDoc,sCategoryVerify)
	GBL_FAILED_FUNCTION_NAME="Fn_ADS_TechDocOperations"
	Dim objDialogNewItem,aPattern,iCount,iCounter,intNodeCount,sTreeItem,objSelectType,intNoOfObjects
	Dim aSrcDocCategory,aSrcTecDocCategory,iFlag,sTechRevID,sTechID,sStatus,aButtons,iReturn
	Dim WshShell, bExit, sList
	Dim dicCount,dicItems,dicKeys,sField,sSubAction	
	
	Set WshShell = CreateObject("WScript.Shell")
	Fn_ADS_TechDocOperations = False
	If Not JavaWindow("DefaultWindow").JavaWindow("NewItem").Exist(2)  Then
			 For iCount=0 to 0
				Window("ADSWindow").JavaDialog("TechnicalDocument").SetTOProperty "title","New Item"
				 If Window("ADSWindow").JavaDialog("TechnicalDocument").Exist(5) Then
				  Exit For
				 End If
				 Window("ADSWindow").JavaDialog("TechnicalDocument").SetTOProperty "title","New Part"
				 If Window("ADSWindow").JavaDialog("TechnicalDocument").Exist(5) Then
				  Exit For
				 End If
				Window("ADSWindow").JavaDialog("TechnicalDocument").SetTOProperty "title","New Design"
				 If Window("ADSWindow").JavaDialog("TechnicalDocument").Exist(5) Then
				  Exit For
				 End If
			Next
     End If
     If Fn_UI_ObjectExist("Fn_ADS_TechDocOperations",Window("ADSWindow").JavaDialog("TechnicalDocument"))=True OR Fn_UI_ObjectExist("Fn_ADS_TechDocOperations", JavaWindow("DefaultWindow").JavaWindow("NewItem"))=True Then
		'Check the existence of "New Item " window
			If Window("ADSWindow").JavaDialog("TechnicalDocument").Exist(2) Then
				Set objDialogNewItem=Fn_UI_ObjectCreate("Fn_ADS_TechDocOperations",Window("ADSWindow").JavaDialog("TechnicalDocument"))
			Else
				Set objDialogNewItem=Fn_UI_ObjectCreate("Fn_ADS_TechDocOperations",JavaWindow("DefaultWindow").JavaWindow("NewItem"))
			End If

			Select Case sAction
			Case "Create"
				'Select Item Type
				Call Fn_List_Select("Fn_ADS_TechDocOperations", objDialogNewItem,"ItemType",dicTechDoc("TechType"))
				'Checked Configuration item or not
				If dicTechDoc("ConfItem") <> "" Then
				 Call Fn_CheckBox_Set("Fn_ADS_TechDocOperations", objDialogNewItem,"ConfigurationItem",dicTechDoc("ConfItem"))
				End If
				'Click on "Next" button
				 Call Fn_Button_Click("Fn_ADS_TechDocOperations", objDialogNewItem,"Next")
				'Set   ID Pattern
				If dicTechDoc("TechIDPattern") <> "" Then
					Call Fn_List_Select("Fn_ADS_TechDocOperations", objDialogNewItem,"IDPattern",dicTechDoc("TechIDPattern"))
				End If
				'Set Technical  Id
				If dicTechDoc("TechID") <> "" Then	
					 Call Fn_Edit_Box("Fn_ADS_TechDocOperations",objDialogNewItem,"TechID", dicTechDoc("TechID"))
				End If
				'Set Technical Revision ID
				If dicTechDoc ("TechRevID") = ""  Then
						 Call Fn_Button_Click("Fn_ADS_TechDocOperations", objDialogNewItem,"Next")
						Call Fn_Edit_Box("Fn_ADS_TechDocOperations", objDialogNewItem,"Revision",dicTechDoc("TechRevID"))
				End If
				If  dicTechDoc("TechID") = "" or dicTechDoc ("TechRevID") = "" Then
					'Click on assign button if  ID and RevID is Blank
					  Call Fn_Button_Click("Fn_ADS_TechDocOperations", objDialogNewItem, "Assign")
				End If
				'Set Technical Revision ID
				If dicTechDoc ("TechRevID") <> ""  Then
						Call Fn_Edit_Box("Fn_ADS_TechDocOperations", objDialogNewItem,"Revision",dicTechDoc("TechRevID"))
				End If
				'Extract Creation data
				dicTechDoc("TechID") = Fn_Edit_Box_GetValue("Fn_ADS_TechDocOperations", objDialogNewItem,"TechID")
				dicTechDoc ("TechRevID") = Fn_Edit_Box_GetValue("Fn_ADS_TechDocOperations", objDialogNewItem,"Revision")
				'Set Item name
				 Call Fn_Edit_Box("Fn_ADS_TechDocOperations", objDialogNewItem,"TechName",dicTechDoc("TechName"))
				'Set description
				Call Fn_Edit_Box("Fn_ADS_TechDocOperations", objDialogNewItem,"Description",dicTechDoc("TechDesc"))
				'Set UOM
				If dicTechDoc("TechUOM") <> "" Then
				  Call Fn_Edit_Box("Fn_ADS_TechDocOperations", objDialogNewItem,"UnitOfMeasure",dicTechDoc("TechUOM"))
				End If
				wait(2)
				'Clicking on Next Button			
				Call Fn_Button_Click("Fn_ADS_TechDocOperations", objDialogNewItem, "Next") 
				Fn_ADS_TechDocOperations = 	dicTechDoc("TechID") & "-" & 	dicTechDoc("TechRevID")
				Wait(2)
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Created an Item of ID [" + CStr(dicTechDoc("TechID")) + "]")
				If dicTechDoc("TechType")="Technical Document" OR dicTechDoc("TechType")="ADS Tec Document" Then
						'Setting Category from the Category list
						If dicTechDoc("CategoryList")<> "" then
							Call Fn_Button_Click("Fn_ADS_TechDocOperations", objDialogNewItem, "Category")  
							Call Fn_List_Select("Fn_ADS_TechDocOperations", objDialogNewItem,"CategoryList",dicTechDoc("CategoryList"))
						End If
						'Clicking on the Next Button of Category List Dialog 
						Call Fn_Button_Click("Fn_ADS_TechDocOperations", objDialogNewItem, "Next_TD")  	
						If dicTechDoc("TechDocList")<> "" then					
							Call Fn_List_Select("Fn_ADS_TechDocOperations", objDialogNewItem,"TechnicalDocList",dicTechDoc("TechDocList"))
						End If
						'Clicking on the Finish Button of Category List Dialog 
						Call Fn_Button_Click("Fn_ADS_TechDocOperations", objDialogNewItem, "Finish_TD")
				ElseIf dicTechDoc("TechType")="ADS Part Sub" OR dicTechDoc("TechType")="Part" OR dicTechDoc("TechType")="Drawing" OR dicTechDoc("TechType")="Design" Then
					If dicTechDoc("TechType")="ADS Part Sub" OR dicTechDoc("TechType")="Part" Then
							'Setting Part Category
							If dicTechDoc("Categories")<> "" then
								bFlag=False 
								objDialogNewItem.JavaStaticText("Label").SetTOProperty "label","Part Category:"
								objDialogNewItem.JavaStaticText("Label").Click 1,1,"LEFT"
								Call Fn_Button_Click("Fn_ADS_TechDocOperations",objDialogNewItem, "LOVDropDown")							
								wait 1
								WshShell.SendKeys "{TAB}"
								wait 1
								WshShell.SendKeys "{DOWN}"
								wait 1
								If objDialogNewItem.JavaTable("LOVTreeTable").Exist(1) Then
									For iCount = 0 to CInt(objDialogNewItem.JavaTable("LOVTreeTable").getROProperty("rows"))-1
										If  trim(objDialogNewItem.JavaTable("LOVTreeTable").Object.getValueAt(iCount,0).getDisplayableValue())= trim(dicTechDoc("Categories")) Then
												objDialogNewItem.JavaTable("LOVTreeTable").DoubleClickCell iCount,0
												bFlag=True
												If objDialogNewItem.JavaTable("LOVTreeTable").Exist(1) Then
													bFlag = False
												End If
												Exit For
										End If
									Next
								End If
								If bFlag = False Then
									Set WshShell = Nothing
									Fn_ADS_TechDocOperations = False
									Exit Function
								End If
							End If
					ElseIf dicTechDoc("TechType")="Design" Then
							'Setting Design Category
							If dicTechDoc("Categories")<> "" then
								objDialogNewItem.JavaEdit("Part Category").SetTOProperty "attached text","Design Category:"
								Call Fn_Edit_Box("Fn_ADS_TechDocOperations", objDialogNewItem,"Part Category",dicTechDoc("Categories"))
							End If
					End If
						'Setting Source Document ID
						If dicTechDoc("SourceDocID")<> "" then
							Call Fn_Edit_Box("Fn_ADS_TechDocOperations", objDialogNewItem,"SourceDocID",dicTechDoc("SourceDocID"))
						End If 
						'Setting Category from the Category list
						If dicTechDoc("CategoryList")<> "" then
								bFlag=False 
								objDialogNewItem.JavaStaticText("Label").SetTOProperty "label","Source Document Category:"
								Call Fn_Button_Click("Fn_ADS_TechDocOperations",objDialogNewItem, "LOVDropDown")							
								wait 1
								WshShell.SendKeys "{TAB}"
								wait 1
								WshShell.SendKeys "{DOWN}"
								wait 1
								If objDialogNewItem.JavaTable("LOVTreeTable").Exist(1) Then
									For iCount = 0 to CInt(objDialogNewItem.JavaTable("LOVTreeTable").getROProperty("rows"))-1
										If  trim(objDialogNewItem.JavaTable("LOVTreeTable").Object.getValueAt(iCount,0).getDisplayableValue())= trim(dicTechDoc("CategoryList")) Then
												objDialogNewItem.JavaTable("LOVTreeTable").DoubleClickCell iCount,0
												bFlag=True
												If objDialogNewItem.JavaTable("LOVTreeTable").Exist(1) Then
													bFlag = False
												End If
												Exit For
										End If
									Next
								End If
								If bFlag = False Then
									Set WshShell = Nothing
									Fn_ADS_TechDocOperations = False
									Exit Function
								End If
						End If
						'Clicking on the Next Button of Category List Dialog 
						If dicTechDoc("TechDocList")<> ""  then	
								bFlag=False 
								objDialogNewItem.JavaStaticText("Label").SetTOProperty "label","Source Technical Document Category:"
								Call Fn_Button_Click("Fn_ADS_TechDocOperations",objDialogNewItem, "LOVDropDown")							
								wait 1
								WshShell.SendKeys "{TAB}"
								wait 1
								WshShell.SendKeys "{DOWN}"
								wait 1
								If objDialogNewItem.JavaTable("LOVTreeTable").Exist(1) Then
									For iCount = 0 to CInt(objDialogNewItem.JavaTable("LOVTreeTable").getROProperty("rows"))-1
										If  trim(objDialogNewItem.JavaTable("LOVTreeTable").Object.getValueAt(iCount,0).getDisplayableValue())= trim(dicTechDoc("TechDocList")) Then
												objDialogNewItem.JavaTable("LOVTreeTable").DoubleClickCell iCount,0
												bFlag=True
												If objDialogNewItem.JavaTable("LOVTreeTable").Exist(1) Then
													bFlag = False
												End If
												Exit For
										End If
									Next
								End If
								If bFlag = False Then
									Set WshShell = Nothing
									Fn_ADS_TechDocOperations = False
									Exit Function
								End If
						End If
						'Clicking on the Finish Button of Category List Dialog 
						Call Fn_Button_Click("Fn_ADS_TechDocOperations", objDialogNewItem, "Finish_TD")
				End If 
			Case "VerifyNamingParttern"
				'Select Item Type
				If dicTechDoc("TechType")="Technical Document" OR dicTechDoc("TechType")="ADS Tec Document" Then
						If dicTechDoc("TechIDPattern") <> "" Then
							aPattern = Split (dicTechDoc("TechIDPattern"),":")
						Else
							Fn_ADS_TechDocOperations = False 
						End If
						iFlag=False						
						For iCount=0 To CInt(objDialogNewItem.JavaTree("ItemType").getROProperty("items count"))-1
							If Trim(objDialogNewItem.JavaTree("ItemType").GetItem(iCount))="Most Recently Used:"+Trim( dicTechDoc("TechType")) Then
								iFlag=True
								Exit For
							ElseIf Trim(objDialogNewItem.JavaTree("ItemType").GetItem(iCount))="Complete List" Then
								Exit For
							End If
						Next
						If iFlag=True Then
							Call Fn_JavaTree_Select(Environment.value("TestName"), objDialogNewItem, "ItemType","Most Recently Used")
							Call Fn_JavaTree_Select(Environment.value("TestName"), objDialogNewItem, "ItemType","Most Recently Used:"+ dicTechDoc("TechType"))
						Else
							Call Fn_UI_JavaTree_Expand(Environment.value("TestName"), objDialogNewItem, "ItemType","Complete List")
							Call Fn_JavaTree_Select(Environment.value("TestName"), objDialogNewItem, "ItemType","Complete List")
							Call Fn_JavaTree_Select(Environment.value("TestName"), objDialogNewItem, "ItemType","Complete List:"+ dicTechDoc("TechType"))	
						End If
						wait 3
						Call Fn_Button_Click(Environment.value("TestName"),objDialogNewItem, "Next")
						objDialogNewItem.JavaStaticText("ItemLabel").SetTOProperty "label","ID:"
						Call Fn_Button_Click(Environment.value("TestName"),objDialogNewItem, "LOVDropDownButton")
						wait 1
						WshShell.SendKeys "{TAB}"
						wait 1
						WshShell.SendKeys "{DOWN}"
						wait 1
						intNodeCount =Fn_UI_Object_GetROProperty("Fn_ADS_TechDocOperations",objDialogNewItem.JavaWindow("TreeShell").JavaTable("LOVTable"),"rows")  
						For iCount = 0 to Ubound(aPattern)		    
								For iCounter = 0 to intNodeCount -1
									sTreeItem =objDialogNewItem.JavaWindow("TreeShell").JavaTable("LOVTable").getCellData(iCounter,0)
									If Trim(lcase(sTreeItem)) = Trim(lcase(aPattern(iCount)))Then
										Fn_ADS_TechDocOperations = TRUE
										Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: "& dicTechDoc("TechIDPattern") &" Exists")
										Exit For
									End If
								Next
								If Cstr (iCounter) = Cstr(intNodeCount) Then
									Set WshShell = Nothing
									Fn_ADS_TechDocOperations = FALSE
									Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: "& dicTechDoc("TechIDPattern") &" Not Exist")
								End If
						Next
						Call Fn_Button_Click(Environment.value("TestName"),objDialogNewItem, "LOVDropDownButton")		
	
				Else
					If dicTechDoc("TechType") <> "" Then
						Call Fn_List_Select("Fn_ADS_TechDocOperations", objDialogNewItem,"ItemType",dicTechDoc("TechType"))
					End If
					'checked Configuration item or not
					If dicTechDoc("ConfItem") <> "" Then
						Call Fn_CheckBox_Set("Fn_ADS_TechDocOperations", objDialogNewItem,"ConfigurationItem",dicTechDoc("ConfItem"))
					End If
					'Click on "Next" button
					 Call Fn_Button_Click("Fn_ADS_TechDocOperations", objDialogNewItem,"Next")			
					If dicTechDoc("TechIDPattern") <> "" Then
						aPattern = Split (dicTechDoc("TechIDPattern"),":")
					Else
						Fn_ADS_TechDocOperations = False 
					End If
					  intNodeCount =Fn_UI_Object_GetROProperty("Fn_ADS_TechDocOperations",objDialogNewItem.JavaList("IDPattern"),"items count")  
						For iCount = 0 to Ubound(aPattern)		    
							For iCounter = 0 to intNodeCount -1
								sTreeItem =objDialogNewItem.JavaList("IDPattern").GetItem(iCounter)
								If Trim(lcase(sTreeItem)) = Trim(lcase(aPattern(iCount)))Then
									Fn_ADS_TechDocOperations = TRUE
									Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: "& dicTechDoc("TechIDPattern") &" Is Exist")
									Exit For
								End If
							Next
							If Cstr (iCounter) = Cstr(intNodeCount) Then
								Fn_ADS_TechDocOperations = FALSE
								Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: "& dicTechDoc("TechIDPattern") &" Not Exist")
							End If
					Next
				End If
			Case "VerifyRevPattern","VerifyDesignLevel", "VerifyCategory", "VerifyTechDocCategory"
					iFlag = 0
					If dicTechDoc("TechRevID")  <> "" Then
						aPattern = Split (dicTechDoc("TechRevID") ,":")
					Else
						Fn_ADS_TechDocOperations = False 
					End If
					If dicTechDoc("TechType")="Technical Document" OR dicTechDoc("TechType")="ADS Tec Document" Then						
							If Trim(Lcase(sAction)) = Trim(Lcase("VerifyRevPattern")) Then
								objDialogNewItem.JavaStaticText("ItemLabel").SetTOProperty "label","Revision:"
							ElseIf Trim(Lcase(sAction)) = Trim(Lcase("VerifyDesignLevel")) Then
								objDialogNewItem.JavaStaticText("ItemLabel").SetTOProperty "label","Design Level:"
							ElseIf  Trim(Lcase(sAction)) = Trim(Lcase("VerifyCategory")) Then
								objDialogNewItem.JavaStaticText("ItemLabel").SetTOProperty "label","Category:"
							ElseIf Trim(Lcase(sAction)) = Trim(Lcase("VerifyTechDocCategory")) Then
								objDialogNewItem.JavaStaticText("ItemLabel").SetTOProperty "label","Technical Document Category:"
							End If                            
							Call Fn_Button_Click(Environment.value("TestName"),objDialogNewItem, "LOVDropDownButton")
							wait 1
							WshShell.SendKeys "{TAB}"
							wait 1
							WshShell.SendKeys "{DOWN}"
							wait 1	
							If objDialogNewItem.JavaWindow("TreeShell").JavaTree("Tree").Exist(1) Then
								  intNodeCount =Fn_UI_Object_GetROProperty("Fn_ADS_TechDocOperations",objDialogNewItem.JavaWindow("TreeShell").JavaTree("Tree"),"items count")  
									For iCount = 0 to Ubound(aPattern)		    
											For iCounter = 0 to intNodeCount -1
												sTreeItem =objDialogNewItem.JavaWindow("TreeShell").JavaTree("Tree").getItem(iCounter)
												If Trim(lcase(aPattern(iCount))) = Trim(lcase(sTreeItem)) Then
													iFlag = iFlag + 1
													Exit For
												End If
											Next
									Next		
									wait 1
									objDialogNewItem.JavaWindow("TreeShell").JavaTree("Tree").Activate 0            									
							ElseIf objDialogNewItem.JavaWindow("TreeShell").JavaTable("LOVTable").Exist(1) Then
									intNodeCount =Fn_UI_Object_GetROProperty("Fn_ADS_TechDocOperations",objDialogNewItem.JavaWindow("TreeShell").JavaTable("LOVTable"),"rows")  
									For iCount = 0 to Ubound(aPattern)		    
											For iCounter = 0 to intNodeCount -1
												sTreeItem =objDialogNewItem.JavaWindow("TreeShell").JavaTable("LOVTable").getCellData(iCounter,0)
												If Instr(1, Trim(lcase(aPattern(iCount))), Trim(lcase(sTreeItem)))>0 Then
													iFlag = iFlag + 1
													Exit For
												End If
											Next
									Next	
									wait 1
									objDialogNewItem.JavaWindow("TreeShell").JavaTable("LOVTable").ActivateCell 0,0						
						   End If			
						   wait 1
						   If objDialogNewItem.JavaWindow("TreeShell").JavaTree("Tree").Exist(1) OR objDialogNewItem.JavaWindow("TreeShell").JavaTable("LOVTable").Exist(1) Then 				
								Call Fn_Button_Click(Environment.value("TestName"),objDialogNewItem, "LOVDropDownButton")
						   End If
							If iFlag = Ubound(aPattern)+1 Then
								Fn_ADS_TechDocOperations = TRUE
								Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: "& dicTechDoc("TechRevID") &" values exist in the List.")
							Else
								Set WshShell = Nothing
								Fn_ADS_TechDocOperations = False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: "& intNodeCount &" does not exist in the List")
							End If							
						
					Else 
		
						Call Fn_Button_Click("Fn_ADS_TechDocOperations", objDialogNewItem, "RevisionID")
						Wait(3)
						Set objSelectType=description.Create()
						objSelectType("Class Name").value = "JavaStaticText"
						Set  intNoOfObjects = objDialogNewItem.ChildObjects(objSelectType)
						For iCount= 0 To Ubound(aPattern)
							For  iCounter = 0 to intNoOfObjects.count-1
							   If intNoOfObjects(iCounter).getROProperty("label") = aPattern(iCount) Then
									iFlag = iFlag + 1
									wait 1
									Exit for
								End If
							Next
						Next			
						If Trim(Lcase(sAction)) = "verifydesignlevel" Then
							Window("ADSWindow").JavaDialog("TechnicalDocument").JavaEdit("Design Level").Click 1,1,"LEFT"
						End If
						If iFlag = Ubound(aPattern)+1 Then
							Fn_ADS_TechDocOperations = TRUE
							Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: "& dicTechDoc("TechRevID") &" values exist in the List.")
						Else
							Fn_ADS_TechDocOperations = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: "& intNodeCount &" does not exist in the List")
						End If
					End If
			Case "VerifyCategoryList","VerifySrcDocCategoryList", "VerifySrcTechDocCategoryList"
					iFlag = 0
					bExit=False					
					Select Case sAction
						Case "VerifyCategoryList"
								If dicTechDoc("CategoryList")  <> "" Then
									sList = dicTechDoc("CategoryList")									
									If dicTechDoc("TechType")="ADS Part Sub" OR dicTechDoc("TechType")="Part" Then 
										objDialogNewItem.JavaStaticText("Label").SetTOProperty "label","Part Category:"
									ElseIf dicTechDoc("TechType")="Design" OR dicTechDoc("TechType")="SubDesign" Then 
										objDialogNewItem.JavaStaticText("Label").SetTOProperty "label","Design Category:"
									End If
								Else
									Fn_ADS_TechDocOperations = False 
									bExit= True		
								End If
						Case "VerifySrcDocCategoryList"
								If dicTechDoc("SrcDocCategories")  <> "" Then
									sList = dicTechDoc("SrcDocCategories")
									objDialogNewItem.JavaStaticText("Label").SetTOProperty "label","Source Document Category:"
								Else
									Fn_ADS_TechDocOperations = False 
									bExit= True		
								End If
						Case "VerifySrcTechDocCategoryList"
								If dicTechDoc("SrcTecDocCategories")  <> "" Then
									sList = dicTechDoc("SrcTecDocCategories")
									objDialogNewItem.JavaStaticText("Label").SetTOProperty "label","Source Technical Document Category:"
								Else
									Fn_ADS_TechDocOperations = False 
									bExit= True		
								End If
					End Select

					If bExit = False Then												
								
							aPattern = Split (sList ,":")
							iFlag=0 
							objDialogNewItem.JavaStaticText("Label").Click 1,1,"LEFT"
							Call Fn_Button_Click("Fn_ADS_TechDocOperations",objDialogNewItem, "LOVDropDown")							
							wait 1
							WshShell.SendKeys "{TAB}"
							wait 1
							WshShell.SendKeys "{DOWN}"
							wait 1								
							If objDialogNewItem.JavaTable("LOVTreeTable").Exist(1) Then
										intNodeCount =Fn_UI_Object_GetROProperty("Fn_ADS_TechDocOperations",objDialogNewItem.JavaTable("LOVTreeTable"),"rows")  
										For iCount = 0 to Ubound(aPattern)		    
												For iCounter = 0 to intNodeCount -1
													If trim(aPattern(iCount))=trim(objDialogNewItem.JavaTable("LOVTreeTable").Object.getValueAt(iCounter,0).getDisplayableValue()) Then
														iFlag = iFlag + 1
														Exit For
													End If
												Next
										Next	
										wait 1		
							   End If   
							If iFlag = Ubound(aPattern)+1 Then
								Fn_ADS_TechDocOperations = TRUE
								Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: "& sList &" values exist in the List.")
							Else
								Fn_ADS_TechDocOperations = False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: "& sList  &" does not exist in the List")
							End If
					End If
			'=========================================================================================================================
			'[TC11.4_20170605_NewDevelopment_PoonamC_16Jan2017] : Added Case to create Tech doc as per sent parameter.
			Case "CreateExt"
			
					dicCount = dicTechDoc.Count
					dicItems = dicTechDoc.Items
					dicKeys = dicTechDoc.Keys
					
					For iCount = 0 To dicCount - 1
						If Instr(dicKeys(iCount),"Button") > 0 and dicKeys(iCount) <> "ButtonName" Then
							sSubAction = "Button"
						Else
							sSubAction = dicKeys(iCount)
						End if
						sField = dicItems(iCount)
					
						Select Case sSubAction
							Case "TechType" ''Select Type
								intNodeCount = Fn_UI_Object_GetROProperty("Fn_ADS_TechDocOperations",objDialogNewItem.JavaTree("ItemType"), "items count")
								For iCounter=0 To intNodeCount-1
									sTreeItem=objDialogNewItem.JavaTree("ItemType").GetItem(iCounter)
									If Trim(sTreeItem)="Most Recently Used:"+Trim(sField) Then
										iFlag=True
										Exit For
									ElseIf Trim(sTreeItem)="Complete List" Then
										Exit For
									End If
								Next
							
								If iFlag=True Then
									Call Fn_JavaTree_Select("Fn_ADS_TechDocOperations", objDialogNewItem, "ItemType","Most Recently Used")
									Call Fn_JavaTree_Select("Fn_ADS_TechDocOperations", objDialogNewItem, "ItemType","Most Recently Used:"+sField)
								Else
									Call Fn_UI_JavaTree_Expand("Fn_ADS_TechDocOperations", objDialogNewItem, "ItemType","Complete List")
									Call Fn_JavaTree_Select("Fn_ADS_TechDocOperations", objDialogNewItem, "ItemType","Complete List")
									Call Fn_JavaTree_Select("Fn_ADS_TechDocOperations", objDialogNewItem, "ItemType","Complete List:"+sField)	
								End If
								Call Fn_ReadyStatusSync(1)		
							Case "Button" 'Click On Button
								Call Fn_Button_Click("Fn_ADS_TechDocOperations", objDialogNewItem,sField)
								Call Fn_ReadyStatusSync(1)	
							Case "TechID" ' Assign Tech ID
								If sField <> "" Then
									 Call Fn_Edit_Box("Fn_ADS_TechDocOperations",objDialogNewItem,"ItemID", sField)
								Else
									Call Fn_Button_Click("Fn_ADS_TechDocOperations", objDialogNewItem, "AssignID")
									Call Fn_ReadyStatusSync(1)
								End If
								sTechID = Fn_Edit_Box_GetValue("Fn_ADS_TechDocOperations", objDialogNewItem,"ItemID")
							Case "TechRevID" 'Assign Tech Rev ID
								Call Fn_Edit_Box("Fn_ADS_TechDocOperations",objDialogNewItem,"RevisionID",sField)
								sTechRevID = Fn_Edit_Box_GetValue("Fn_ADS_TechDocOperations", objDialogNewItem,"RevisionID")
							Case "TechRevIDPattern" ' Assign Text Rev pattern ID
								 objDialogNewItem.JavaButton("AssignRevID").SetTOProperty "label",""
								 objDialogNewItem.JavaButton("AssignRevID").Click micLeftBtn
								 wait 1
								 Set objSelectType=Description.Create()
								 objSelectType("Class Name").value = "JavaTable"
								 Set intNoOfObjects=objDialogNewItem.ChildObjects(objSelectType)
								 For iCounter = 0 To intNoOfObjects(0).GetROProperty("rows") - 1
										If trim(intNoOfObjects(0).GetCellData(iCounter,0)) = trim(sField) Then
											intNoOfObjects(0).SelectCell iCounter,0
											Wait 1
										End if
								 Next 
								 sTechRevID = Fn_Edit_Box_GetValue("Fn_ADS_TechDocOperations", objDialogNewItem,"RevisionID")
							Case "TechName"
									Call Fn_Edit_Box("Fn_ADS_TechDocOperations", objDialogNewItem,"ItemName",sField)	 
							Case "CategoryList" ' Select Category
								objDialogNewItem.JavaEdit("TechCategory").SetTOProperty "index",0
								Call Fn_Edit_Box("Fn_ADS_TechDocOperations", objDialogNewItem,"TechCategory",sField)
								Call Fn_KeyBoardOperation("SendKeys", "{TAB}")
							Case "TechDocList" ' Select Tech Doc Category
								objDialogNewItem.JavaEdit("TechCategory").SetTOProperty "index",1
								Call Fn_Edit_Box("Fn_ADS_TechDocOperations",objDialogNewItem,"TechCategory", sField)
								Call Fn_KeyBoardOperation("SendKeys", "{TAB}")					
						End Select
					Next
					Fn_ADS_TechDocOperations = sTechID & "-" & sTechRevID
			'=========================================================================================================================
			End Select

				'Click on Buttons
				If dicTechDoc("ButtonName")<>"" Then
						aButtons = split(dicTechDoc("ButtonName"), ":",-1,1)
						iCounter = Ubound(aButtons)
						For iCount=0 to iCounter
							'Click on aButtons			
							Wait(2)			
							Call Fn_Button_Click("Fn_ADS_TechDocOperations", objDialogNewItem, aButtons(iCount))  
							Wait(5)
							Call Fn_ReadyStatusSync(3)
							If Trim(Lcase(aButtons(iCount))) = "finish" OR Trim(Lcase(aButtons(iCount))) = "close" Then
								If objDialogNewItem.Exist(7) Then
									If objDialogNewItem.JavaButton(aButtons(iCount)).GetROProperty("enabled") = 1 Then									
										Call Fn_Button_Click("Fn_ADS_TechDocOperations", objDialogNewItem, aButtons(iCount))  
									End If
								End If
							End If
						Next
				End If
	Else 
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: New Item Dialog Does not Exists")	
	End If
		Set objDialogNewItem=Nothing
End Function
'#######################################################################################
'###     FUNCTION NAME   :   Fn_ADS_ParametricValues(sAction, sParameter, sValue, sNoteText, bFlagNote, sPartsListNote ,sButtons)
'###
'###    DESCRIPTION     :   Set / Verify values in Input Parametric Input values table.
'###
'###    Return Value  	:   	True / False
'###
'###    HISTORY         :   		AUTHOR              	DATE        		VERSION
'###
'###    CREATED BY      :     Ketan Raje			   29/10/2010   			1.0
'###
'###    REVIWED BY      :		Harshal 				29/10/2010			1.0
'###
'###    MODIFIED BY     :
'###    EXAMPLE         :  	Case "Set" : Call Fn_ADS_ParametricValues("Set", "", "30:90", "", "", "STDNOTE-000018","OK")
'###    							Case "Verify" : Call Fn_ADS_ParametricValues("Verify", "", "", "", "", "STDNOTE-000018" ,"")
'#############################################################################################
Public Function Fn_ADS_ParametricValues(sAction, sParameter, sValue, sNoteText, bFlagNote, sPartsListNote ,sButtons)
	GBL_FAILED_FUNCTION_NAME="Fn_ADS_ParametricValues"
Dim objDialog, i, iLen, WshShell, aValues, iCount, aButtons
Fn_ADS_ParametricValues = FALSE
	Set objDialog = Fn_UI_ObjectCreate( "Fn_ADS_ParametricValues",JavaWindow("ADS-TeamCenter").JavaWindow("Input Parametric Values"))
	If objDialog.Exist Then
		Select Case sAction
				Case "Set"
						'Set Parameters
						'Set Values
						If sValue<>"" Then
							aValues = split(sValue, ":",-1, 1)
							For iCount=0 to ubound(aValues)
								JavaWindow("ADS-TeamCenter").JavaWindow("Input Parametric Values").JavaTable("ParametricValueTable").ActivateCell iCount,"Value"
									iLen = len(aValues(iCount))
										Set WshShell = CreateObject("WScript.Shell")
									For i = 1 to iLen
										WshShell.SendKeys mid(aValues(iCount),i,1)
									Next
										Set WshShell = Nothing						
							Next
						End If
						'Set NoteText
						'Set FlagNote
						If bFlagNote<>"" Then
							If Trim(Lcase(bFlagNote)) = "off" Then
								'Set to OFF
								Call Fn_CheckBox_Set("Fn_ADS_ParametricValues", objDialog, "FlagNote", "OFF")
							ElseIf Trim(Lcase(bFlagNote)) = "on" Then
								'Set to ON
								Call Fn_CheckBox_Set("Fn_ADS_ParametricValues", objDialog, "FlagNote", "ON")
							End If
						End If
						'Set Part List Note
						If sPartsListNote<>"" Then
							Call Fn_Edit_Box("Fn_ADS_ParametricValues",objDialog,"PartsListNode",sPartsListNote)
						End If
						Fn_ADS_ParametricValues = TRUE
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_ADS_ParametricValues successfully completed with Set case")
				Case "Verify"
						'Verify Parts List Note.
						If sPartsListNote<>"" Then
							If Trim(Lcase(sPartsListNote)) = Trim(Lcase(Fn_Edit_Box_GetValue("Fn_ADS_ParametricValues",objDialog,"PartsListNode"))) Then
								Fn_ADS_ParametricValues = TRUE
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_ADS_ParametricValues successfully completed with Verify case")
							End If
						End If
		End Select
	Else
		Fn_ADS_ParametricValues = FALSE
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_ADS_ParametricValues function failed")
	End If
	'Click on Buttons
	If sButtons<>"" Then
			aButtons = split(sButtons, ":",-1,1)
			iCount = Ubound(aButtons)
			For i=0 to iCount
				'Click on Add Button
				Call Fn_Button_Click("Fn_ADS_ParametricValues", objDialog, aButtons(i))
                Call Fn_ReadyStatusSync(2)
			Next
	End If		
Set objDialog = Nothing
Set WshShell = Nothing
End Function
'#########################################################################################################
'###
'###    FUNCTION NAME   :   Fn_ADS_SubmittalsTableContentOperation(sAction, sObjectName, sPropertyName, sExpectedValue)
'###
'###    DESCRIPTION        :   General utility function which applied to Submittals Table Content Operation
'###
'###    Function Calls       :   
'###
'###	 HISTORY             :   AUTHOR                 DATE        VERSION
'###
'###    CREATED BY     :   Ketan Raje           02-Nov-10         1.0
'###
'###    REVIWED BY     :  Harshal 				02-Nov-10         1.0
'###    
'###    EXAMPLE         : 		1.) Case "Rowexist"	:	Msgbox Fn_ADS_SubmittalsTableContentOperation("Rowexist", "000242-000219-000242", "", "")
'###    EXAMPLE          : 		2.) Case "Rowselect"	:	Msgbox Fn_ADS_SubmittalsTableContentOperation("Rowselect", "000242-000219-000242", "", "")
'###    EXAMPLE          : 		3.) Case "Rowmultiselect"	:	Msgbox Fn_ADS_SubmittalsTableContentOperation("Rowmultiselect", "000242-000219-000242:000244-000219-000244", "", "")
'###    EXAMPLE          : 		4.) Case "Rowcellexist"	:	Msgbox Fn_ADS_SubmittalsTableContentOperation("Rowcellexist", "000242-000219-000242", "ID", 242)
'###    																		Msgbox Fn_ADS_SubmittalsTableContentOperation("Rowcellexist", "000242-000219-000242", "Owner", "AutoTestDBA (autotestdba)")
'###									5.)	Case "RowCount":Msgbox Fn_ADS_SubmittalsTableContentOperation("RowCount", "", "", "")
'###									6.) Case "GetRowData" : Msgbox Fn_ADS_SubmittalsTableContentOperation("GetRowData", 0, "", "")
'#############################################################################################################
Public Function Fn_ADS_SubmittalsTableContentOperation(sAction, sObjectName, sPropertyName, sExpectedValue)	
		GBL_FAILED_FUNCTION_NAME="Fn_ADS_SubmittalsTableContentOperation"
		Dim objSubmittalsTable,bReturn,bDoubleClickReturn,iCounter,aObjList,intItemCount,oCounter, rowIndex, bFlag, aMenuList, intCount, sMenu, sText, aMenuList1()
		Dim colCount, i, tab, textArr, columnNumber, columnFoundFlag, intObjectColumnNumber
		Dim colNameArr, bHeaderFoundFlag
		columnFoundFlag = False
		bHeaderFoundFlag = False
		intObjectColumnNumber = -1
		Fn_ADS_SubmittalsTableContentOperation = False	
		' create an object of the table
		Set objSubmittalsTable =JavaWindow("DefaultWindow").JavaTable("SubmittalsTable")				
		colCount =  objSubmittalsTable.GetROProperty("cols")
		' Mapping Object column to column number.		
		For i = 0 to colCount - 1
			If objSubmittalsTable.GetColumnName(i) = "Object" then
						intObjectColumnNumber = i
						bHeaderFoundFlag = True
						Exit for
			end if
		next
		If  bHeaderFoundFlag = False Then
				For i = 0 to colCount - 1 
						textArr = split(objSubmittalsTable.GetColumnName(i),"text=")
						colNameArr = split(textArr(1),",")
						If colNameArr(0) = "Object" then
									intObjectColumnNumber = i
									Exit for
						end if
				 Next
		End If	
		 If intObjectColumnNumber = -1 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Fn_ADS_SubmittalsTableContentOperation : Object column does not exist")						
		 End If	
		bHeaderFoundFlag = False
		columnNumber = -1
	' Mapping column name to number
		If  sPropertyName<>"" Then
				For i = 0 to colCount - 1
					If objSubmittalsTable.GetColumnName(i) = sPropertyName then
							columnNumber = i
							bHeaderFoundFlag = True
							Exit for
					end if
				next
				If  bHeaderFoundFlag = False Then
					 For i = 0 to colCount - 1 
							textArr = split(objSubmittalsTable.GetColumnName(i),"text=")
							colNameArr = split(textArr(1),",")
							If colNameArr(0) = sPropertyName then
										columnNumber = i
										columnFoundFlag = true
										Exit for
							end if
					 Next
					If columnFoundFlag = true Then
								columnNumber = i
					else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Fn_ADS_SubmittalsTableContentOperation : Column " + sPropertyName + " does not exist")	
								Exit function
					 End If
				 End If
		End If
		Select Case sAction
			 Case "Rowexist"
					bFlag = false
					'Count number of rows of Table
					bReturn = objSubmittalsTable.GetROProperty("rows")	
					'Extract the index of row at which the object exist.
					For iCounter=0 to bReturn - 1
'						sText = objSubmittalsTable.GetCellData(iCounter, intObjectColumnNumber )'	Object  column				
                        sText=objSubmittalsTable.object.getItem(iCounter).getData().toString() 'Added by Nilesh on 22-Feb-2013
							If IsNumeric(sObjectName) Then
								 If cstr(sText) = cstr(cint(sObjectName))  Then
									 bFlag = true
									 Exit for
								End If
							elseIf cstr(sText) = cstr(sObjectName)  Then
									 bFlag = true
									 Exit for
							End If									
					Next
					If bFlag = false Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Fn_ADS_SubmittalsTableContentOperation : Row with Object "&sObjectName&" does not exist")	
							Exit function
					Else 
							Fn_ADS_SubmittalsTableContentOperation = True
					End If	
			Case "Rowselect"									 
					'Count number of rows of Table
					bReturn = objSubmittalsTable.GetROProperty("rows")	
					'Extract the index of row at which the object exist.
					For iCounter=0 to bReturn - 1
'						sText = objSubmittalsTable.GetCellData(iCounter, intObjectColumnNumber) ' Object Column			
						 sText=objSubmittalsTable.object.getItem(iCounter).getData().toString() 'Added by Nilesh on 22-Feb-2013			
						If IsNumeric(sObjectName) Then
							 If cstr(sText) = cstr(cint(sObjectName))  Then
								 objSubmittalsTable.ClickCell iCounter,intObjectColumnNumber,"LEFT"
								 Fn_ADS_SubmittalsTableContentOperation = True				 
								 Exit for
							End If
						elseIf cstr(sText) = cstr(sObjectName)  Then
								 objSubmittalsTable.ClickCell iCounter, intObjectColumnNumber,"LEFT"
								 Fn_ADS_SubmittalsTableContentOperation = True				 
								 Exit for
						End If									
					Next
			 Case "Rowmultiselect"
					Dim bSelectFlag
					'Split the string where " : " exist
					aObjList = Split(sObjectName,":")
					intItemCount =ubound(aObjList)
					'Count number of rows of Table
					bReturn = objSubmittalsTable.GetROProperty("rows")	
					'Extract the index of row at which the object exist.
					For oCounter=0 to intItemCount
							bSelectFlag = False
							For iCounter=0 to bReturn-1
									sText = objSubmittalsTable.GetCellData(iCounter, intObjectColumnNumber)	' Object column					
									If IsNumeric(aObjList(oCounter)) Then
											If cstr(sText) = cstr(cint(aObjList(oCounter)))  Then
														objSubmittalsTable.ClickCell iCounter, intObjectColumnNumber ,"LEFT","CONTROL"
														bSelectFlag =True
														Exit for
											End If
									elseIf cstr(sText) = cstr(aObjList(oCounter))  Then
											 objSubmittalsTable.ClickCell iCounter, intObjectColumnNumber ,"LEFT","CONTROL"
											 bSelectFlag =True
											 Exit for
									End If
							Next
							If bSelectFlag = False Then
									 Fn_ADS_SubmittalsTableContentOperation = False
									 Exit function
							End If
					Next
					Fn_ADS_SubmittalsTableContentOperation = True					
			Case "Rowcellexist"
					If  sObjectName = "" or sExpectedValue = ""  Then
							Fn_ADS_SubmittalsTableContentOperation = FALSE	 
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_ADS_SubmittalsTableContentOperation: Rowcellexist : Incorrect input parameters")
							Exit function
					End If
					'Count number of rows of Table
					bReturn = objSubmittalsTable.GetROProperty("rows")
					intCount = 0
					For iCounter=0 to bReturn - 1
							sText = objSubmittalsTable.GetCellData(iCounter, intObjectColumnNumber) ' Object column
							If IsNumeric(sObjectName) Then
								 If cstr(sText) = cstr(cint(sObjectName))  Then
									ReDim Preserve aMenuList1(intCount+1)
									aMenuList1( intCount) = iCounter
									intCount = intCount + 1
								End If
							elseIf cstr(sText) = cstr(sObjectName)  Then
								ReDim Preserve aMenuList1(intCount+1)
									aMenuList1(intCount) = iCounter
									intCount = intCount + 1
							End If
					Next
					If intCount <> 0 Then
							For iCounter = 0 To UBound(aMenuList1) - 1
									   If objSubmittalsTable.GetCellData(aMenuList1(iCounter), columnNumber ) = sExpectedValue  Then
												 intCount = aMenuList1(iCounter)
												 Exit For
									   End If
							Next
 						   If objSubmittalsTable.GetCellData(intCount, columnNumber) <> sExpectedValue Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_ADS_SubmittalsTableContentOperation: Rowcellexist : Expected value is not present")
									 Fn_ADS_SubmittalsTableContentOperation = False  
									Exit function
							End If  
					End If
					Fn_ADS_SubmittalsTableContentOperation = True 
			Case "RowCount"
					bReturn = objSubmittalsTable.GetROProperty("rows")	
					Fn_ADS_SubmittalsTableContentOperation = bReturn
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_ADS_SubmittalsTableContentOperation: Case: RowCount")
			Case "GetRowData"
					 intCount = objSubmittalsTable.GetROProperty("rows")
					 colCount = objSubmittalsTable.GetROProperty("cols")
					 ReDim colNameArr(colCount-1)
					 If sObjectName>intCount-1 Then
						  Fn_ADS_SubmittalsTableContentOperation = False
						  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_ADS_SubmittalsTableContentOperation failed with case "&sAction) 
						  Exit Function
					 Else      
						  For iCounter = 0 to colCount-1
								colNameArr(iCounter) = objSubmittalsTable.GetCellData(sObjectName,iCounter)
						  Next
						  Fn_ADS_SubmittalsTableContentOperation = colNameArr
						  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_ADS_SubmittalsTableContentOperation passed with case "&sAction&" on Row number "&sObjectName)
						  Exit Function
					 End If
		End Select
					If Fn_ADS_SubmittalsTableContentOperation = True Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_ADS_SubmittalsTableContentOperation passed with case "&sAction&" on Object "&sObjectName)	
					else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_ADS_SubmittalsTableContentOperation failed with case "&sAction&" on Object "&sObjectName)	
					End If
					Set objSubmittalsTable = Nothing 
End Function
'***********************************************************************************************************************************************************************************
'Function Name		         :	       Fn_ADS_PropertyRetrive()

'Description			   :		   This function is used to retrive the property value form property panel.

'Parameters			  :	 			sAction,sProperty,sButtons
'											    										
'Return Value		            :		True / False

'Pre-requisite			   :		  Properties Tree Should Exist
'
'Examples				   :	Msgbox Fn_ADS_PropertyRetrive("Get","Schedule Deliverables","")
	 
'History:
'						Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'					   	Ketan Raje				21-Oct-2010																Harshal Agarwal
'***********************************************************************************************************************************************************************************
Public Function Fn_ADS_PropertyRetrive(sAction,sProperty,sButtons)
	GBL_FAILED_FUNCTION_NAME="Fn_ADS_PropertyRetrive"
	Dim objPropertyTree,iCount,bFlag,aButtons   
	Call Fn_SetView("General:Properties")
	wait 2
	Call Fn_ToolbarOperation("Click", "Show Advanced Properties","")
	'Call Fn_ReadyStatusSync(3)
	wait 2
	bFlag = False
	Set objPropertyTree = JavaWindow("DefaultWindow").JavaTree("PropertiesTree")
	   Select Case sAction
		Case "Get"
					sProperty = "Properties:"+sProperty
					If objPropertyTree.Exist Then
						For iCount=0 to objPropertyTree.GetROProperty("count_all_items")-1
							If sProperty = objPropertyTree.GetItem(iCount) then
								bFlag = True
								Exit For
							End If
						Next
						If bFlag = True Then
							Fn_ADS_PropertyRetrive = objPropertyTree.GetColumnValue(sProperty,"Value")
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_ADS_PropertyRetrive compeleted successfully")							
						ElseIf bFlag = False Then
							Fn_ADS_PropertyRetrive = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Property Name is Invalid")							
						End If
					Else
						Fn_ADS_PropertyRetrive = False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Property Tree does not exist")
					End If
	   End Select
   Set objPropertyTree = Nothing
End Function

'***********************************************************************************************************************************************************************************
'Function Name		         :	       Fn_ADS_ItemListVerify()

'Description			 		  :		   This function is used to Verify List Box and label for Correspondence, Category,Correspondence Direction etc

'Return Value		     	    :		True/False

'Pre-requisite			  		 :		  
'Examples						:	
'										-> Case "Correspondence"
		'			   :					Call Fn_ADS_ItemListVerify("Correspondence","Correspondence", "OFF","None:None:Ketan:Testing:None", "Correspondence Type:|Wire List:Assembly Drawing", "", "", "", "", "", "", "", "", "")
'										-> Case "Category"
'											Call  Fn_ADS_ItemListVerify("Category","", "","", "None", "Category:|ICM:MEMO", "", "", "", "", "", "", "", "")
'										-> Case "VerifyLabel"
'											Call Fn_ADS_ItemListVerify("VerifyLabel","", "","", "None", "Recieved Org Name:~References:", "", "", "", "", "", "", "", "")
'History:
'						Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'					   	   Govind Singh				02-Nov-2010																	Ketan R
'***********************************************************************************************************************************************************************************	
Function Fn_ADS_ItemListVerify(sAction,sSelectType, bConfItem, sItemInfo, sAddItemInfo, sAddItemRevInfo, sAttachFileInfo, sWorkFlowInfo, sIdentifierBasicInfo, sAddIDInfo, sAddRevInfo, sAssignProj, sDefineOptions, sButtons)
	GBL_FAILED_FUNCTION_NAME="Fn_ADS_ItemListVerify"
   on error Resume Next
	Dim ObjStaticText, objDialogNewItem, aItemInfo, sItemId, sRevId, aAddItemInfo, aProjectName, iRowData, icount, iCounter, sOptions, aButtons,WshShell,iLen,i,j,sStatus,arrA,arrListItem
	Dim objSelectType,sText

	Set objDialogNewItem = Fn_SISW_GetObject("New Item")
	'Select menu [File -> New -> Item...]
	If Fn_UI_ObjectExist("Fn_ADS_ItemListVerify",objDialogNewItem)=False Then
        Call Fn_MenuOperation("Select","File:New:Item...")
		Call Fn_ReadyStatusSync(2)
	End If
	'Creating Object of links on the left side of the window
	Set ObjStaticText =objDialogNewItem.JavaStaticText("Stpes")
'	Set objDialogNewItem = JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("New Item")
'	'Check the existence of "New Item " window
	Set objDialogNewItem=Fn_UI_ObjectCreate("Fn_ADS_ItemListVerify",objDialogNewItem)
			'Select Item Type
			If sSelectType <> "" Then
				Call Fn_List_Select("Fn_ADS_ItemListVerify", objDialogNewItem,"SelectedProject",sSelectType)
			End If
			'checked Configuration item or not
			If bConfItem <> "" Then
             Call Fn_CheckBox_Set("Fn_ADS_ItemListVerify", objDialogNewItem,"Configuration Item",bConfItem)
			End If
			'Click on "Next" button
			If sSelectType <> "" Then
             Call Fn_Button_Click("Fn_ADS_ItemListVerify", objDialogNewItem,"Next")
			End If
		'Enter Item Information
		If sItemInfo<>"" Then
				aItemInfo = split(sItemInfo, ":",-1,1)
				'click on assign button
				If  aItemInfo(0) = "None" or aItemInfo(1) = "None" Then	
					Call Fn_Button_Click("Fn_ADS_ItemListVerify", objDialogNewItem,"Assign")
				Else
					'Set Item ID
					Call Fn_Edit_Box("Fn_ADS_ItemListVerify",objDialogNewItem,"ItemID", aItemInfo(0))
					If sSelectType="P4_AAUItem1" Then
						objDialogNewItem.JavaButton("UnitOfMeasure").SetTOProperty "index",1
						If objDialogNewItem.JavaButton("UnitOfMeasure").Exist Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verified that Revision is a ListBox")
						End If
					End If
					'Set Revision ID
					Call Fn_Edit_Box("Fn_ADS_ItemListVerify",objDialogNewItem,"RevisionID", aItemInfo(1))
				End If				
				'Set Item name
				 Call Fn_Edit_Box("Fn_ADS_ItemListVerify", objDialogNewItem,"ItemName",aItemInfo(2))
				'Set description
				If aItemInfo(3)<>"None" Then
					Call Fn_Edit_Box("Fn_ADS_ItemListVerify", objDialogNewItem,"Description",aItemInfo(3))
				End If
				'Set UOM
				If aItemInfo(4) <> "None" Then
				  Call Fn_Edit_Box("Fn_ADS_ItemListVerify", objDialogNewItem,"Unit of Measure",aItemInfo(4))
				End If 
		End If
		'Entering Additional Item Information
			If sAddItemInfo<>"None" Then				
				' Click on Next Button
				ObjStaticText.SetTOProperty "label", "Enter Additional Item Information"
				ObjStaticText.Click 1, 1
				aAddItemInfo = split(sAddItemInfo, "~",-1,1)	
                    	If sSelectType="Contract" OR sSelectType="P4_SubContract"  Then
									objDialogNewItem.JavaButton("UnitOfMeasure").Click micLeftBtn
									objDialogNewItem.JavaStaticText("Stpes").SetTOProperty "label",aAddItemInfo(0)
									Wait(3)
									objDialogNewItem.JavaStaticText("Stpes").Click 5,5,"LEFT"
						ElseIf sSelectType = "ADS Tec Document" OR sSelectType = "Technical Document" Then
									Call Fn_Edit_Box("Fn_ADS_ItemListVerify",objDialogNewItem,"Category", aAddItemInfo(0))
								  Call Fn_Edit_Box("Fn_ADS_ItemListVerify",objDialogNewItem,"Technical Document Category", aAddItemInfo(1))
						ElseIf sSelectType = "Data Requirement Item" OR sSelectType = "P4_SubDRI" Then
							If aAddItemInfo(0)<>"None" Then
									Call Fn_Edit_Box("Fn_ADS_ItemListVerify",objDialogNewItem,"Contract Line Item Number", aAddItemInfo(0))
							End If
							If aAddItemInfo(1)<>"None" Then
									Call Fn_Edit_Box("Fn_ADS_ItemListVerify",objDialogNewItem,"Contract Reference", aAddItemInfo(1))
							End If
					ElseIf sSelectType = "Data Item Description" OR sSelectType = "P4_SubDID" Then
							If aAddItemInfo(0)<>"None" Then
									Call Fn_Edit_Box("Fn_ADS_ItemListVerify",objDialogNewItem,"DID Type", aAddItemInfo(0))
							End If
							If aAddItemInfo(1)<>"None" Then
									Call Fn_Edit_Box("Fn_ADS_ItemListVerify",objDialogNewItem,"Program Phases", aAddItemInfo(1))
							End If
					ElseIf sSelectType = "Standard Note" Then
							If aAddItemInfo(0)<>"" Then
									Call Fn_Edit_Box("Fn_ADS_ItemListVerify",objDialogNewItem,"Note Category", aAddItemInfo(0))
							End If
					ElseIf  sSelectType = "Correspondence" Then
				End If
			End If
Select Case sAction
		Case "Correspondence"						
							aAddItemInfo = split(sAddItemInfo, "~",-1,1)	
							sStatus =objDialogNewItem.JavaStaticText("Correspondence Type").GetROProperty("label")
								arrA  = Split(aAddItemInfo(0),"|")
					For i = 0 to uBound (arrA)
								If sStatus = arrA(0) Then
									objDialogNewItem.JavaButton("UnitOfMeasure").SetTOProperty "index",0
									objDialogNewItem.JavaButton("UnitOfMeasure").Click micLeftBtn
								End If
										arrListItem  =Split(arrA(1),":")
							For icount = 0 to ubound(arrListItem)						
									Set objSelectType=description.Create()
								objSelectType("Class Name").value = "JavaStaticText"
								Set  intNoOfObjects =objDialogNewItem.ChildObjects(objSelectType)
								For  iCounter = 0 to intNoOfObjects.count-1
									   If  intNoOfObjects(iCounter).getROProperty("label") = arrListItem(icount) Then
											Fn_ADS_ItemListVerify = True
											Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS:Successfully Verified the values")
											wait 1
											Exit for
										Else
											Fn_ADS_ItemListVerify = False
											Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: "& arrListItem  &" Is NOT Exist in the List")
										End If
								Next
							Next
						Next
	Case "SentDateExistence"
		If objDialogNewItem.JavaCheckBox("SentDate").Exist then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Successfully verified the Existence of Sent Date Chechbox")
			Fn_ADS_ItemListVerify = True
		Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Failed to verify the Existence of Sent Date Chechbox")
				Fn_ADS_ItemListVerify = False
		End If
Case "RequestedtDateExistence"
		If objDialogNewItem.JavaCheckBox("RequestedDate").Exist then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Successfully verified the Existence of Requested Date Chechbox")
			Fn_ADS_ItemListVerify = True
		Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Failed to verify the Existence of Requested Date Chechbox")
				Fn_ADS_ItemListVerify = False
		End If
'		 Call Fn_Button_Click("Fn_ADS_ItemListVerify", objDialogNewItem,"Next")	
	Case "Category"
		If sAddItemRevInfo<>"None" Then
		' Click on Next Button		
		ObjStaticText.SetTOProperty "label", "Enter Additional Item Revision Information"
		ObjStaticText.Click 1, 1
							aAddItemInfo = split(sAddItemRevInfo, "~",-1,1)	
							objDialogNewItem.JavaStaticText("Correspondence Type").SetTOProperty "label","Category:"
							sStatus = objDialogNewItem.JavaStaticText("Correspondence Type").GetROProperty("label")
								arrA  = Split(aAddItemInfo(0),"|")
					For i = 0 to uBound (arrA)
								If sStatus = arrA(0) Then
									objDialogNewItem.JavaButton("UnitOfMeasure").SetTOProperty "index",0
									objDialogNewItem.JavaButton("UnitOfMeasure").Click micLeftBtn
								End If
										arrListItem  =Split(arrA(1),":")
							For icount = 0 to ubound(arrListItem)						
									Set objSelectType=description.Create()
								objSelectType("Class Name").value = "JavaStaticText"
								Set  intNoOfObjects = objDialogNewItem.ChildObjects(objSelectType)
								For  iCounter = 0 to intNoOfObjects.count-1
									   If  intNoOfObjects(iCounter).getROProperty("label") = arrListItem(icount) Then
											Fn_ADS_ItemListVerify = True
											Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Successfully verified the Listbox for Category")
											wait 1
											Exit for
										Else
											Fn_ADS_ItemListVerify = False
											Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Failed to verify the Listbox for Category")
										End If
								Next
							Next
						Next
				End If
		Case "CorrespondenceDir"	
					If sAddItemRevInfo<>"None" Then
				' Click on Next Button		
				ObjStaticText.SetTOProperty "label", "Enter Additional Item Revision Information"
				ObjStaticText.Click 1, 1
							aAddItemInfo = split(sAddItemRevInfo, "~",-1,1)	
							objDialogNewItem.JavaStaticText("Correspondence Type").SetTOProperty "label","Correspondence Direction:"
							sStatus =objDialogNewItem.JavaStaticText("Correspondence Type").GetROProperty("label")
								arrA  = Split(aAddItemInfo(0),"|")
					For i = 0 to uBound (arrA)
								If sStatus = arrA(0) Then
									objDialogNewItem.JavaButton("UnitOfMeasure").SetTOProperty "index",1
									objDialogNewItem.JavaButton("UnitOfMeasure").Click micLeftBtn
								End If
										arrListItem  =Split(arrA(1),":")
							For icount = 0 to ubound(arrListItem)						
									Set objSelectType=description.Create()
								objSelectType("Class Name").value = "JavaStaticText"
								Set  intNoOfObjects =objDialogNewItem.ChildObjects(objSelectType)
								For  iCounter = 0 to intNoOfObjects.count-1
									   If  intNoOfObjects(iCounter).getROProperty("label") = arrListItem(icount) Then
											Fn_ADS_ItemListVerify = True
											Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Successfully verified the Listbox for Correspondence Directory")
											wait 1
											Exit for
										Else
											Fn_ADS_ItemListVerify = False
											Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Failed to verify the Listbox for Correspondence Directory")
										End If
								Next
							Next
						Next	
					End If
	Case "VerifyLabel"	 	
		aAddItemInfo = split(sAddItemRevInfo, "~",-1,1)					
			For i = 0 to Ubound(aAddItemInfo)
			objDialogNewItem.JavaStaticText("Correspondence Type").SetTOProperty "label",aAddItemInfo(i)
'			sStatus = JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("New Item").JavaStaticText("Correspondence Type").GetROProperty("label")
'				If sStatus =aAddItemInfo(i)  Then
			If objDialogNewItem.JavaStaticText("Correspondence Type").Exist(2) = True Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Successfully Verified the Label"+ aAddItemInfo+ "")
						Fn_ADS_ItemListVerify = True
				Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Failed to Verify the Label"+ aAddItemInfo +"")
						Fn_ADS_ItemListVerify = False
				End If
			Next
				If sButtons<>"" Then
					aButtons = split(sButtons, ":",-1,1)
					iCounter = Ubound(aButtons)
					For icount=0 to iCounter
						'Click on Add Button
						Call Fn_Button_Click("Fn_ADS_ItemListVerify", objDialogNewItem, aButtons(iCount))
						Call Fn_ReadyStatusSync(2)
					Next
			End If
			Fn_ADS_ItemListVerify = True
			Call Fn_ReadyStatusSync(1)
			'Write Log
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Completed the function Fn_PWC_ItemDetailsCreate")
 Case "Priority"
		  If sAddItemRevInfo<>"None" Then
		  ' Click on Next Button  
		  ObjStaticText.SetTOProperty "label", "Enter Additional Item Revision Information"
		  ObjStaticText.Click 1, 1
			   aAddItemInfo = split(sAddItemRevInfo, "~",-1,1) 
				objDialogNewItem.JavaStaticText("Correspondence Type").SetTOProperty "label","Priority:"
			   sStatus =objDialogNewItem.JavaStaticText("Correspondence Type").GetROProperty("label")
				arrA  = Split(aAddItemInfo,"|")
			 For i = 0 to uBound (arrA)
				If sStatus = arrA Then
				objDialogNewItem.JavaButton("UnitOfMeasure").SetTOProperty "index",0
				 objDialogNewItem.JavaButton("UnitOfMeasure").Click micLeftBtn
				End If
				  arrListItem  =Split(arrA(1),":")
			   For icount = 0 to ubound(arrListItem)      
				 Set objSelectType=description.Create()
				objSelectType("Class Name").value = "JavaStaticText"
				Set  intNoOfObjects =objDialogNewItem.ChildObjects(objSelectType)
				For  iCounter = 0 to intNoOfObjects.count-1
					If  intNoOfObjects(iCounter).getROProperty("label") = arrListItem(icount) Then
				   Fn_ADS_ItemListVerify = True
				   Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Successfully verified the Listbox for Category")
				   wait 1
				   Exit for
				  Else
				   Fn_ADS_ItemListVerify = False
				   Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Failed to verify the Listbox for Category")
				  End If
				Next
			   Next
			  Next
			End If
Case "VerifyText"
				'Assign to Project
			If sAssignProj<>"" Then
				' Click on Next Button
					ObjStaticText.SetTOProperty "label", "Assign to Program"
					ObjStaticText.Click 1, 1
					Call Fn_ReadyStatusSync(2)	
				sStatus =  	objDialogNewItem.JavaStaticText("SelectProgram").GetROProperty ("label")
				If 	sStatus = sAssignProj Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Successfully Verified the Label for Project"+ sAssignProj+ "")
							Fn_ADS_ItemListVerify = True
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Failed to Verify the Label for Project"+ sAssignProj +"")
						Fn_ADS_ItemListVerify = False
				End If
		End If			
Case "VerifyLengthOfficePrimary"
		If sAddItemRevInfo<>"None" Then
			' Click on Next Button		
			ObjStaticText.SetTOProperty "label", "Enter Additional Item Revision Information"
			ObjStaticText.Click 1, 1
			sText =objDialogNewItem.JavaEdit("OfficePrimaryResp").GetROProperty("value")
			Fn_ADS_ItemListVerify = Len(sText)
		End If
Case "VerifyLengthComments"
		If sAddItemRevInfo<>"None" Then
			' Click on Next Button		
			ObjStaticText.SetTOProperty "label", "Enter Additional Item Revision Information"
			ObjStaticText.Click 1, 1
			 objDialogNewItem.JavaEdit("OfficePrimaryResp").SetTOProperty"attached text","Comments:"
			sText =objDialogNewItem.JavaEdit("OfficePrimaryResp").GetROProperty("value")
			Fn_ADS_ItemListVerify = Len(sText)
		End If		
	End Select
Set objDialogNewItem = Nothing
Set objSelectType = Nothing
End Function
'***********************************************************************************************************************************************************************************
'Function Name		         :	       Fn_ADS_GenerateSubmittalDeliverySync()

'Description			   :		   This function is used to Sync the Progress Information Dialog

'Return Value		            :		True

'Pre-requisite			   :		  Progress Information Dialog should Exist
'
'Examples				   :	Msgbox Fn_ADS_GenerateSubmittalDeliverySync()
	 
'History:
'						Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'					   	Harshal Agrawal			08-Nov-2010																Harshal Agarwal
'**************************************************************************************************************************************************************************************
Public Function Fn_ADS_GenerateSubmittalDeliverySync()
	Wait(3)
While JavaWindow("ADS-TeamCenter").JavaWindow("Generate Submittal Delivery").JavaWindow("Progress Information").Exist
	Wait(5)
Wend
Fn_ADS_GenerateSubmittalDeliverySync = True
End Function
'#########################################################################################################
'###
'###    FUNCTION NAME   :   Fn_ADS_DeliverableTableContentOperation(sAction, sObjectName, sPropertyName, sExpectedValue)
'###
'###    DESCRIPTION        :   General utility function which applied to Deliverable Table Content Operation
'###
'###    Function Calls       :   
'###
'###	 HISTORY             :   AUTHOR                 DATE        VERSION
'###
'###    CREATED BY     :   Ketan Raje           02-Nov-10         1.0
'###
'###    REVIWED BY     :  Harshal 				02-Nov-10         1.0
'###    
'###    EXAMPLE         : 		1.) Case "Rowexist"	:	Msgbox Fn_ADS_DeliverableTableContentOperation("Rowexist", "000242-000219-000242", "", "")
'###    EXAMPLE          : 		2.) Case "Rowselect"	:	Msgbox Fn_ADS_DeliverableTableContentOperation("Rowselect", "000242-000219-000242", "", "")
'###    EXAMPLE          : 		3.) Case "Rowmultiselect"	:	Msgbox Fn_ADS_DeliverableTableContentOperation("Rowmultiselect", "000242-000219-000242:000244-000219-000244", "", "")
'###    EXAMPLE          : 		4.) Case "Rowcellexist"	:	Msgbox Fn_ADS_DeliverableTableContentOperation("Rowcellexist", "000242-000219-000242", "ID", 242)
'###    																		Msgbox Fn_ADS_DeliverableTableContentOperation("Rowcellexist", "000242-000219-000242", "Owner", "AutoTestDBA (autotestdba)")
'###									5.)	Case "RowCount":Msgbox Fn_ADS_DeliverableTableContentOperation("RowCount", "", "", "")
'###									6.) Case "GetRowData" : Msgbox Fn_ADS_DeliverableTableContentOperation("GetRowData", 0, "", "")
'#############################################################################################################
Public Function Fn_ADS_DeliverableTableContentOperation(sAction, sObjectName, sPropertyName, sExpectedValue)	
		GBL_FAILED_FUNCTION_NAME="Fn_ADS_DeliverableTableContentOperation"
		Dim objDeliverablesTable,bReturn,bDoubleClickReturn,iCounter,aObjList,intItemCount,oCounter, rowIndex, bFlag, aMenuList, intCount, sMenu, sText, aMenuList1()
		Dim colCount, i, tab, textArr, columnNumber, columnFoundFlag, intObjectColumnNumber
		Dim colNameArr, bHeaderFoundFlag
		columnFoundFlag = False
		bHeaderFoundFlag = False
		intObjectColumnNumber = -1
		Fn_ADS_DeliverableTableContentOperation = False	
		' create an object of the table
		Set objDeliverablesTable =JavaWindow("DefaultWindow").JavaTable("Deliverables")				
		colCount =  objDeliverablesTable.GetROProperty("cols")
		' Mapping Object column to column number.		
		For i = 0 to colCount - 1
			If objDeliverablesTable.GetColumnName(i) = "Name" then
						intObjectColumnNumber = i
						bHeaderFoundFlag = True
						Exit for
			end if
		next
		If  bHeaderFoundFlag = False Then
				For i = 0 to colCount - 1 
						textArr = split(objDeliverablesTable.GetColumnName(i),"text=")
						colNameArr = split(textArr(1),",")
						If colNameArr(0) = "Name" then
									intObjectColumnNumber = i
									Exit for
						end if
				 Next
		End If	
		 If intObjectColumnNumber = -1 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Fn_ADS_DeliverableTableContentOperation : Object column does not exist")						
		 End If	
		bHeaderFoundFlag = False
		columnNumber = -1
	' Mapping column name to number
		If  sPropertyName<>"" Then
				For i = 0 to colCount - 1
					If objDeliverablesTable.GetColumnName(i) = sPropertyName then
							columnNumber = i
							bHeaderFoundFlag = True
							Exit for
					end if
				next
				If  bHeaderFoundFlag = False Then
					 For i = 0 to colCount - 1 
							textArr = split(objDeliverablesTable.GetColumnName(i),"text=")
							colNameArr = split(textArr(1),",")
							If colNameArr(0) = sPropertyName then
										columnNumber = i
										columnFoundFlag = true
										Exit for
							end if
					 Next
					If columnFoundFlag = true Then
								columnNumber = i
					else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Fn_ADS_DeliverableTableContentOperation : Column " + sPropertyName + " does not exist")	
								Exit function
					 End If
				 End If
		End If
		Select Case sAction
			 Case "Rowexist"
					bFlag = false
					'Count number of rows of Table
					bReturn = objDeliverablesTable.GetROProperty("rows")	
					'Extract the index of row at which the object exist.
					For iCounter=0 to bReturn - 1
						sText = objDeliverablesTable.GetCellData(iCounter, intObjectColumnNumber )'	Object  column				
							If IsNumeric(sObjectName) Then
								 If cstr(sText) = cstr(cint(sObjectName))  Then
									 bFlag = true
									 Exit for
								End If
							elseIf cstr(sText) = cstr(sObjectName)  Then
									 bFlag = true
									 Exit for
							End If									
					Next
					If bFlag = false Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Fn_ADS_DeliverableTableContentOperation : Row with Object "&sObjectName&" does not exist")	
							Exit function
					Else 
							Fn_ADS_DeliverableTableContentOperation = True
					End If	
			Case "Rowselect"									 
					'Count number of rows of Table
					bReturn = objDeliverablesTable.GetROProperty("rows")	
					'Extract the index of row at which the object exist.
					For iCounter=0 to bReturn - 1
						sText = objDeliverablesTable.GetCellData(iCounter, intObjectColumnNumber) ' Object Column						
						If IsNumeric(sObjectName) Then
							 If cstr(sText) = cstr(cint(sObjectName))  Then
								 objDeliverablesTable.ClickCell iCounter,intObjectColumnNumber,"LEFT"
								 Fn_ADS_DeliverableTableContentOperation = True				 
								 Exit for
							End If
						elseIf cstr(sText) = cstr(sObjectName)  Then
								 objDeliverablesTable.ClickCell iCounter, intObjectColumnNumber,"LEFT"
								 Fn_ADS_DeliverableTableContentOperation = True				 
								 Exit for
						End If									
					Next
			 Case "Rowmultiselect"
					Dim bSelectFlag
					'Split the string where " : " exist
					aObjList = Split(sObjectName,":")
					intItemCount =ubound(aObjList)
					'Count number of rows of Table
					bReturn = objDeliverablesTable.GetROProperty("rows")	
					'Extract the index of row at which the object exist.
					For oCounter=0 to intItemCount
							bSelectFlag = False
							For iCounter=0 to bReturn-1
									sText = objDeliverablesTable.GetCellData(iCounter, intObjectColumnNumber)	' Object column					
									If IsNumeric(aObjList(oCounter)) Then
											If cstr(sText) = cstr(cint(aObjList(oCounter)))  Then
														objDeliverablesTable.ClickCell iCounter, intObjectColumnNumber ,"LEFT","CONTROL"
														bSelectFlag =True
														Exit for
											End If
									elseIf cstr(sText) = cstr(aObjList(oCounter))  Then
											 objDeliverablesTable.ClickCell iCounter, intObjectColumnNumber ,"LEFT","CONTROL"
											 bSelectFlag =True
											 Exit for
									End If
							Next
							If bSelectFlag = False Then
									 Fn_ADS_DeliverableTableContentOperation = False
									 Exit function
							End If
					Next
					Fn_ADS_DeliverableTableContentOperation = True					
			Case "Rowcellexist"
					If  sObjectName = "" or sExpectedValue = ""  Then
							Fn_ADS_DeliverableTableContentOperation = FALSE	 
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_ADS_DeliverableTableContentOperation: Rowcellexist : Incorrect input parameters")
							Exit function
					End If
					'Count number of rows of Table
					bReturn = objDeliverablesTable.GetROProperty("rows")
					intCount = 0
					For iCounter=0 to bReturn - 1
							sText = objDeliverablesTable.GetCellData(iCounter, intObjectColumnNumber) ' Object column
							If IsNumeric(sObjectName) Then
								 If cstr(sText) = cstr(cint(sObjectName))  Then
									ReDim Preserve aMenuList1(intCount+1)
									aMenuList1( intCount) = iCounter
									intCount = intCount + 1
								End If
							elseIf cstr(sText) = cstr(sObjectName)  Then
								ReDim Preserve aMenuList1(intCount+1)
									aMenuList1(intCount) = iCounter
									intCount = intCount + 1
							End If
					Next
					If intCount <> 0 Then
							For iCounter = 0 To UBound(aMenuList1) - 1
									   If objDeliverablesTable.GetCellData(aMenuList1(iCounter), columnNumber ) = sExpectedValue  Then
												 intCount = aMenuList1(iCounter)
												 Exit For
									   End If
							Next
 						   If objDeliverablesTable.GetCellData(intCount, columnNumber) <> sExpectedValue Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_ADS_DeliverableTableContentOperation: Rowcellexist : Expected value is not present")
									 Fn_ADS_DeliverableTableContentOperation = False  
									Exit function
							End If  
					End If
					Fn_ADS_DeliverableTableContentOperation = True 
			Case "RowCount"
					bReturn = objDeliverablesTable.GetROProperty("rows")	
					Fn_ADS_DeliverableTableContentOperation = bReturn
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_ADS_DeliverableTableContentOperation: Case: RowCount")
			Case "GetRowData"
					 intCount = objDeliverablesTable.GetROProperty("rows")
					 colCount = objDeliverablesTable.GetROProperty("cols")
					 ReDim colNameArr(colCount-1)
					 If sObjectName>intCount-1 Then
						  Fn_ADS_DeliverableTableContentOperation = False
						  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_ADS_DeliverableTableContentOperation failed with case "&sAction) 
						  Exit Function
					 Else      
						  For iCounter = 0 to colCount-1
								colNameArr(iCounter) = objDeliverablesTable.GetCellData(sObjectName,iCounter)
						  Next
						  Fn_ADS_DeliverableTableContentOperation = colNameArr
						  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_ADS_DeliverableTableContentOperation passed with case "&sAction&" on Row number "&sObjectName)
						  Exit Function
					 End If
		End Select
					If Fn_ADS_DeliverableTableContentOperation = True Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_ADS_DeliverableTableContentOperation passed with case "&sAction&" on Object "&sObjectName)	
					else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_ADS_DeliverableTableContentOperation failed with case "&sAction&" on Object "&sObjectName)	
					End If
					Set objDeliverablesTable = Nothing 
End Function
'#################################################################################################################
'###    FUNCTION NAME   :    Fn_ADS_DialogMsgVerify(sErrMsg,sDialogTitle,sButton)
'###
'###    DESCRIPTION     :   This function used to verify the Dialog messages
'###
'###    PARAMETERS      :   sErrMsg,sDialogTitle,sButton
'###                        
'###    Function Calls  :   Fn_WriteLogFile(), Fn_Button_Click()
'###
'###    HISTORY         :   AUTHOR                   DATE        VERSION
'###
'###    CREATED BY      :   Ketan Raje      	16/11/2010	  1.0
'###
'###    REVIWED BY      :   Harshal		   		16/11/2010	  1.0          
'###
'###    MODIFIED BY     :
'###    EXAMPLE         :   Msgbox Fn_ADS_DialogMsgVerify("Do you want to save your modifications to the Rich Content?","Teamcenter","Yes")
'################################################################################################################
Function Fn_ADS_DialogMsgVerify(sErrMsg,sDialogTitle,sButton)

	 Dim dicErrorInfo
	 Set dicErrorInfo = CreateObject("Scripting.Dictionary")
	 With dicErrorInfo 
	  .Add "Message" , sErrMsg
	  .Add "Title", sDialogTitle
	  .Add "Button", sButton
	  .Add "Action", "DialogMsgVerify" 	  
	 End with
   Fn_ADS_DialogMsgVerify = Fn_SISW_ADS_ErrorVerify(dicErrorInfo)
   Set dicErrorInfo = Nothing

End Function
'#################################################################################################################
'###    FUNCTION NAME   :    Fn_ADS_UIOperations(sAction,sSettings,sValues, sButtons)
'###
'###    DESCRIPTION     :   	Verify the existance of UI Objects and the values.
'###                        
'###    Function Calls  :   		Fn_Button_Click, Fn_WriteLogFile
'###
'###    HISTORY         :   	AUTHOR                   DATE        	VERSION  		Build
'###
'###    CREATED BY      :   Ketan Raje      		08/06/2011	  		1.0				20110504	
'###
'###	MODIFIED BY     :   Sushma Pagare     08/06/2011                            20110504     Added Case "LocationCodeLOV"
'###	
'###	MODIFIED BY     :   Sandeep Navghane   04/01/2012                       modified case : LocationType & CageCode . Added case to set label property of static text [ ObjectName ] to recognize DrpDwnButton button as its visualy attached with it
'###
'###    EXAMPLE          :   Msgbox Fn_ADS_UIOperations("Verify", "LocationType", "CAGE, Commercial and Government Entity", "")
'###    							Msgbox Fn_ADS_UIOperations("Verify", "OrgCageCode", "CAGE1:CAGE2:CAGE3:CAGE4", "Close")	
'###    							Msgbox Fn_ADS_UIOperations("Verify", "CageCode", "CAGE1:CAGE2:CAGE3:CAGE4", "Cancel")	
'################################################################################################################
Function Fn_ADS_UIOperations(sAction, sObject, sValues, sButtons)
	GBL_FAILED_FUNCTION_NAME="Fn_ADS_UIOperations"
	Dim iRows, iCount, aValues, intCount, iCounter
	Fn_ADS_UIOperations = False
	Select Case sAction
	Case "Verify"
			Select Case sObject
						Case "LocationType"
								'Click on the dropdownButton.			
                                JavaWindow("DefaultWindow").JavaWindow("NewBusinessObject").JavaStaticText("ObjectName").SetTOProperty "label","Location Type:"
								Call Fn_Button_Click("Fn_ADS_UIOperations", JavaWindow("DefaultWindow").JavaWindow("NewBusinessObject"), "DrpDwnButton")
								If JavaWindow("DefaultWindow").JavaWindow("NewBusinessObject").JavaWindow("Shell").JavaTable("LocationTypeTable").Exist Then
									iRows = JavaWindow("DefaultWindow").JavaWindow("NewBusinessObject").JavaWindow("Shell").JavaTable("LocationTypeTable").GetROProperty("rows")
									For iCount = 0 to iRows-1
										If Trim(Lcase(JavaWindow("DefaultWindow").JavaWindow("NewBusinessObject").JavaWindow("Shell").JavaTable("LocationTypeTable").GetCellData(iCount,0))) = Trim(Lcase(sValues)) Then
											Fn_ADS_UIOperations = True
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verified value "&sValues&" in LocationType List.")
											Exit For
										End If
									Next
								Else
									'Added by siddhi
                                    'Object change to java tree in TC 10.1
											aValues =split(sValues,",")											
											If JavaWindow("DefaultWindow").JavaWindow("NewBusinessObject").JavaWindow("Shell").JavaTree("Tree").Exist Then
													iRows = JavaWindow("DefaultWindow").JavaWindow("NewBusinessObject").JavaWindow("Shell").JavaTree("Tree").GetROProperty("items count")
														For iCount = 0 to iRows-1
															If Trim(Lcase(JavaWindow("DefaultWindow").JavaWindow("NewBusinessObject").JavaWindow("Shell").JavaTree("Tree").GetColumnValue(aValues(0),"Description"))) = Trim(Lcase(aValues(1))) Then
																Fn_ADS_UIOperations = True
																Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verified value "& aValues(0) &" - "& aValues(1)&" in LocationType Tree.")
																Exit For
															End If
														Next
											End If
								End If
								'Click on sButtons button.			
								If sButtons <> "" Then									
									Call Fn_Button_Click("Fn_ADS_UIOperations", JavaWindow("DefaultWindow").JavaWindow("NewBusinessObject"), sButtons)
								End If			
						Case "OrgCageCode"'-------------------------Modified code as per the changed UI 10.1_0213				By Pranav(01-Mar-2013)
								intCount = 0 
								iCounter = 0
								aValues = Split(sValues,":")
								'Click on DropDown button for OrgCageCode.
'								Call Fn_Button_Click("Fn_ADS_UIOperations", JavaWindow("ADS-TeamCenter").JavaWindow("TcDefaultApplet").JavaDialog("New Item"), "UnitOfMeasure")
								JavaWindow("DefaultWindow").JavaWindow("NewBusinessObject").JavaStaticText("ObjectName").SetTOProperty "label","Original CAGE Code:"
								JavaWindow("DefaultWindow").JavaWindow("NewBusinessObject").JavaButton("DrpDwnButton").Click micLeftBtn
								For iCount = 0 to Ubound(aValues)
									intCount = intCount + 1
									'SetTOProperty for CAGECode.
'									JavaWindow("ADS-TeamCenter").JavaWindow("TcDefaultApplet").JavaDialog("New Item").JavaStaticText("Steps").SetTOProperty "Label",aValues(iCount)
'									If JavaWindow("ADS-TeamCenter").JavaWindow("TcDefaultApplet").JavaDialog("New Item").JavaStaticText("Steps").Exist(5) Then
'										iCounter = iCounter + 1
'									End If
									If JavaWindow("DefaultWindow").JavaWindow("NewBusinessObject").JavaWindow("Shell").JavaTree("Tree").GetItem(iCount)=aValues(iCount) Then
										iCounter = iCounter + 1
									End If
								Next
'								JavaWindow("DefaultWindow").JavaWindow("NewBusinessObject").JavaButton("DrpDwnButton").Click micLeftBtn
								If intCount = iCounter Then
										Fn_ADS_UIOperations = True
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verified values "&sValues&" in CAGECode List.")
								End If
								'Click on sButtons button.			
								If sButtons <> "" Then							
									Call Fn_Button_Click("Fn_ADS_UIOperations", JavaWindow("DefaultWindow").JavaWindow("NewBusinessObject"), sButtons)
'								JavaWindow("DefaultWindow").JavaWindow("NewBusinessObject").JavaButton("Cancel").Click micLeftBtn
								End If
						Case "CageCode"  ''Added By Sushma'-------------------------Modified code as per the changed UI 10.1_0213				By Pranav(11-Mar-2013)
								aValues = Split(sValues,":")
								JavaWindow("DefaultWindow").JavaWindow("NewBusinessObject").JavaStaticText("ObjectName").SetTOProperty "label","CAGE Code:"
								Call Fn_Button_Click("Fn_ADS_UIOperations", JavaWindow("DefaultWindow").JavaWindow("NewBusinessObject"), "DrpDwnButton")
'								If JavaWindow("DefaultWindow").JavaWindow("NewBusinessObject").JavaWindow("Shell").JavaTable("LocationTypeTable").Exist Then
'									iRows = JavaWindow("DefaultWindow").JavaWindow("NewBusinessObject").JavaWindow("Shell").JavaTable("LocationTypeTable").GetROProperty("rows")
'									For iCounter = 0 to Ubound(aValues)
'										Fn_ADS_UIOperations = False
'										For iCount = 0 to iRows-1
'											If Trim(Lcase(JavaWindow("DefaultWindow").JavaWindow("NewBusinessObject").JavaWindow("Shell").JavaTable("LocationTypeTable").GetCellData(iCount,0))) = Trim(Lcase(aValues(iCounter))) Then
'												Fn_ADS_UIOperations = True
'												Exit For
'											End If
'										Next
									For iCount = 0 to Ubound(aValues)
									intCount = intCount + 1
										If JavaWindow("DefaultWindow").JavaWindow("NewBusinessObject").JavaWindow("Shell").JavaTree("Tree").GetItem(iCount)=aValues(iCount) Then
											iCounter = iCounter + 1
										End If
									Next
									If intCount = iCounter Then
										Fn_ADS_UIOperations = True
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verified values "&sValues&" in CAGECode List.")
									End If
'										If Fn_ADS_UIOperations = False Then
'											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to find value "&aValues(iCounter)&" in Location Code Drop down List")
'											Exit For
'										End If
'									Next
'								End If
								If Fn_ADS_UIOperations = True Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verified values "&sValues&" in  Location Code Drop down List.")
								End If
								'Click on sButtons button.          
								If sButtons <> "" Then									
									Call Fn_Button_Click(Environment.Value("TestName"), JavaWindow("DefaultWindow").JavaWindow("NewBusinessObject"), sButtons)
								End If																											
			End Select
	End Select
End Function
'*********************************************************		Function to create detail Item		***********************************************************************
'Function Name		:        Fn_ADS_ItemDetailCreate  

'Description	    	:        Creates an Item with detail information

'Parameters		     :    		sItemType: Item type to be selected
'			                         	sItemID: Unique ID for the Item [if non-empty, then enter]
'							          	sItemRevID: Revision of the Item [if non-empty, then enter] - if any one of the fields (id/rev) are blank then click Assign button
'									 	sItemName: Name of the Item
'									  	sItemDesc: Description of the Item
' 										dicItemDetailsCreate : Dictionary paramter  for detail creation
'										Example 						dicItemDetailsCreate("ItemAddInfo") = "yes"
'																			dicItemDetailsCreate("OrgCageCode") = "pune123456"
'Return Value		: 			ItemID-ItemRevID 

'Pre-requisite	    :		 	Should be logged in

'Examples		    :			Call Fn_ADS_ItemDetailCreate("Item","","","TesItem","Testing Item creation.",dicItemDetailsCreate)

'History		    :		
'													Developer Name				Date						Rev. No.			
'--------------------------------------------------------------------------------------------------------------------------
'													Ketan Raje				     06/06/2011			           1.0								
'													Sandeep N				    25/07/2011			           1.1        Remove Old Item dialog Hierarchy and added new							
'--------------------------------------------------------------------------------------------------------------------------
Public Function Fn_ADS_ItemDetailCreate(sItemType, sItemID, sItemRevID, sItemName, sItemDesc, dicItemDetailsCreate)
		GBL_FAILED_FUNCTION_NAME="Fn_ADS_ItemDetailCreate"
		Dim sAssignId, sAssignRevId, aProjectName, objDialogNewItem, ObjStaticText, bReturn
		Dim sNewItemMenu
		On Error Resume Next
		'Creating Object for New Item window.
		Set objDialogNewItem =Fn_ADS_SISW_GetObject("New Item")	
		 'Creating Object of links on the left side of the window
		Set ObjStaticText = objDialogNewItem.JavaStaticText("Steps")
		sNewItemMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("RAC_Menu"),"NewItem")
		
		If not objDialogNewItem.Exist (SISW_MIN_TIMEOUT)  Then
			'Select menu [File -> New -> Item...]
			bReturn = Fn_MenuOperation("Select",sNewItemMenu)
			Call Fn_ReadyStatusSync(3)
			If bReturn = False Then
					Fn_ADS_ItemDetailCreate = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Select Menu [File:New:Item]")
					Set objDialogNewItem = Nothing
					Exit Function
			Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Selected Menu [File:New:Item]")
			End If
		End If
		
	   'Check the existence of the "NewItem" Window
		'If objDialogNewItem.Exist (20)  Then
		If Fn_SISW_UI_Object_Operations("Fn_ADS_ItemDetailCreate","Exist",objDialogNewItem,SISW_MIN_TIMEOUT)  Then
					'Select  "Item Type"
					objDialogNewItem.JavaList("ItemType").Select sItemType
					If Err.Number < 0 Then
							Fn_ADS_ItemDetailCreate = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Select Item Type [" + sItemType + "]")
							Set objDialogNewItem = Nothing
							Exit Function
					End If
					'Click on "Next" button
					objDialogNewItem.JavaButton("Next").WaitProperty "enabled", 1, 200000
					objDialogNewItem.JavaButton("Next").Click
					If Err.Number < 0 Then
							Fn_ADS_ItemDetailCreate = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Click [Next] Button")
							Set objDialogNewItem = Nothing
							Exit Function
					End If
					'Enter Item ID
					If sItemID <> "" Then
						 objDialogNewItem.JavaEdit("ItemID").Set sItemID
					End If
					'Enter Revision ID
					If sItemRevID <> "" Then
						objDialogNewItem.JavaEdit("RevisionID").Set sItemRevID
					End If
					'Check  "Item Id and Revision ID"
					If sItemID = "" or sItemRevID = "" Then
							'Click on "Assign" button
							objDialogNewItem.JavaButton("Assign").WaitProperty "enabled", 1, 20000
							objDialogNewItem.JavaButton("Assign").Click
					End If
					'Extract Item Id and Rev Id
					sAssignId = objDialogNewItem.JavaEdit("ItemID").GetROProperty("value")
					sAssignRevId = objDialogNewItem.JavaEdit("RevisionID").GetROProperty("value")
					'Set the Item Name
					If sItemName <> "" Then
						objDialogNewItem.JavaEdit("ItemName").Set sItemName
					End If
					If sItemDesc <> "" Then
						objDialogNewItem.JavaEdit("Description").Set sItemDesc
					End If
		Else
					Fn_ADS_ItemDetailCreate = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[New Item] Dialog Not Found")
		End If

		Select Case sItemType
                    Case "Item","CageCodeItem","Document", "CageDrawing", "CageCustNote", "CageStdNote"
								If Trim(Lcase(dicItemDetailsCreate("ItemAddInfo"))) = "yes" Then
										'Click on link Enter Item Options and Details
										ObjStaticText.SetTOProperty "label","Enter Additional Item Information"
										ObjStaticText.WaitProperty "enabled" , 1, 20000
										ObjStaticText.Click 1, 1
										'Enter Original Cage Code.
										If dicItemDetailsCreate("OrgCageCode") <> "" Then
											objDialogNewItem.JavaEdit("CageCode").Set dicItemDetailsCreate("OrgCageCode")
										End If
										'Enter Note Category.
										If dicItemDetailsCreate("NoteCategory") <> "" Then
											objDialogNewItem.JavaEdit("Note Category").Set dicItemDetailsCreate("NoteCategory")
										End If
										'Enter Source Document ID.
										If dicItemDetailsCreate("SrcDocID") <> "" Then
											objDialogNewItem.JavaEdit("SourceDocID").Set dicItemDetailsCreate("SrcDocID")
										End If
								End If
                     Case "Technical Document","ADS Tec Document", "CageTechDoc"
								If Trim(Lcase(dicItemDetailsCreate("ItemAddInfo"))) = "yes" Then
										'Click on link Enter Item Options and Details
										ObjStaticText.SetTOProperty "label","Enter Additional Item Information"
										ObjStaticText.WaitProperty "enabled" , 1, 20000
										ObjStaticText.Click 1, 1
										'Enter Category
										If dicItemDetailsCreate("Category") <> "" Then
											objDialogNewItem.JavaEdit("Category").Set dicItemDetailsCreate("Category")
										End If
										'Enter Original Cage Code.
										If dicItemDetailsCreate("OrgCageCode") <> "" Then
											objDialogNewItem.JavaEdit("CageCode").Set dicItemDetailsCreate("OrgCageCode")
										End If
										'Enter Technical Document Category.
										If dicItemDetailsCreate("TecDocCategory") <> "" Then
											objDialogNewItem.JavaEdit("Technical Document Category").Set dicItemDetailsCreate("TecDocCategory")
										End If
								End If
		End Select

		'Click on "Finish" button
		objDialogNewItem.JavaButton("Finish").WaitProperty "enabled" , 1, 20000
		objDialogNewItem.JavaButton("Finish").Click
		'Click on "Close" button
		objDialogNewItem.JavaButton("Close").WaitProperty "enabled" , 1, 20000
		objDialogNewItem.JavaButton("Close").Click
		Call Fn_ReadyStatusSync(2)
		Fn_ADS_ItemDetailCreate = sAssignId+"-"+sAssignRevId
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Item [" + sItemName + "] Created Successfully")

End Function
'*********************************************************		Function to create detail Part		***********************************************************************
'Function Name		:        Fn_ADS_PartDetailCreateDic  

'Description	    	:        Creates an Part with detail information

'Parameters		     :    		sPartType: Part type to be selected
'			                         	sPartID: Unique ID for the Part [if non-empty, then enter]
'							          	sPartRevID: Revision of the Part [if non-empty, then enter] - if any one of the fields (id/rev) are blank then click Assign button
'									 	sPartName: Name of the Part
'									  	sPartDesc: Description of the Part
' 										dicPartDetailsCreate : Dictionary paramter  for detail creation
'	Example	 		:						  Set dicPartDetailsCreate = CreateObject( "Scripting.Dictionary" )
'													dicPartDetailsCreate.RemoveAll
'													dicPartDetailsCreate("PartAddInfo") = "yes"
'													dicPartDetailsCreate("OrgCageCode") = "pune123456"
'													Msgbox Fn_ADS_PartDetailCreateDic("Commercial Part","","","TesPart","Testing Part creation.",dicPartDetailsCreate)
'													Set dicPartDetailsCreate = Nothing
'Return Value		: 			PartID-PartRevID 

'Pre-requisite	    :		 	Should be logged in

'Examples		    :			Call Fn_ADS_PartDetailCreateDic("Part","","","TesPart","Testing Part creation.",dicPartDetailsCreate)

'History		    :		
'													Developer Name				Date						Rev. No.			
'--------------------------------------------------------------------------------------------------------------------------
'													Ketan Raje				     15/06/2011			           1.0								
'--------------------------------------------------------------------------------------------------------------------------
Public Function Fn_ADS_PartDetailCreateDic(sPartType, sPartID, sPartRevID, sPartName, sPartDesc, dicPartDetailsCreate)
		GBL_FAILED_FUNCTION_NAME="Fn_ADS_PartDetailCreateDic"
		Dim sAssignId, sAssignRevId, aProjectName, objDialogNewPart, ObjStaticText, bReturn

		On Error Resume Next
		'Creating Object for New Part window.
		'Coomented by Siddhi
'		set objDialogNewPart = JavaWindow("ADS-TeamCenter").JavaWindow("TcDefaultApplet").JavaDialog("New Part")		
	'	Update New Hierarchy
		set objDialogNewPart = Window("ADSWindow").JavaDialog("New Part")
		 'Creating Object of links on the left side of the window
		Set ObjStaticText =Window("ADSWindow").JavaDialog("New Part").JavaStaticText("Steps")

		If not objDialogNewPart.Exist (5)  Then
			'Select menu [File -> New -> Part...]
			bReturn = Fn_MenuOperation("Select","File:New:Part...")
			Wait(10)
			If bReturn = False Then
					Fn_ADS_PartDetailCreateDic = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Select Menu [File:New:Part]")
					Set objDialogNewPart = Nothing
					Exit Function
			Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Selected Menu [File:New:Part]")
			End If
		End If
	   'Check the existence of the "NewPart" Window
		If objDialogNewPart.Exist (20)  Then
					'Select  "Part Type"
					objDialogNewPart.JavaList("PartType").Select sPartType
					If Err.Number < 0 Then
							Fn_ADS_PartDetailCreateDic = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Select Part Type [" + sPartType + "]")
							Set objDialogNewPart = Nothing
							Exit Function
					End If
					'Click on "Next" button
					objDialogNewPart.JavaButton("Next").WaitProperty "enabled", 1, 200000
					objDialogNewPart.JavaButton("Next").Click
					If Err.Number < 0 Then
							Fn_ADS_PartDetailCreateDic = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Click [Next] Button")
							Set objDialogNewPart = Nothing
							Exit Function
					End If
					'Enter Part ID
					If sPartID <> "" Then
						 objDialogNewPart.JavaEdit("ID").Set sPartID
					End If
					'Enter Revision ID
					If sPartRevID <> "" Then
						objDialogNewPart.JavaEdit("RevisionID").Set sPartRevID
					End If
					'Check  "Part Id and Revision ID"
					If sPartID = "" or sPartRevID = "" Then
							'Click on "Assign" button
							objDialogNewPart.JavaButton("Assign").WaitProperty "enabled", 1, 20000
							objDialogNewPart.JavaButton("Assign").Click
					End If
					'Extract Part Id and Rev Id
					sAssignId = objDialogNewPart.JavaEdit("ID").GetROProperty("value")
					sAssignRevId = objDialogNewPart.JavaEdit("RevisionID").GetROProperty("value")
					'Set the Part Name
					If sPartName <> "" Then
						objDialogNewPart.JavaEdit("Name").Set sPartName
					End If
					If sPartDesc <> "" Then
						objDialogNewPart.JavaEdit("Description").Set sPartDesc
					End If
		Else
					Fn_ADS_PartDetailCreateDic = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[New Part] Dialog Not Found")
		End If

		Select Case sPartType
					Case "Commercial Part", "CagePart"
								If Trim(Lcase(dicPartDetailsCreate("PartAddInfo"))) = "yes" Then
										'Click on link Enter Part Options and Details
										ObjStaticText.SetTOProperty "label","Enter Additional Part Information"
										ObjStaticText.WaitProperty "enabled" , 1, 20000
										ObjStaticText.Click 1, 1
										'Enter Original Cage Code.
										If dicPartDetailsCreate("OrgCageCode") <> "" Then
											objDialogNewPart.JavaEdit("CageCode").Set dicPartDetailsCreate("OrgCageCode")
										End If
										'Enter Part Category.
										If dicPartDetailsCreate("PartCategory") <> "" Then
											objDialogNewPart.JavaEdit("PartCategory").Set dicPartDetailsCreate("PartCategory")
									    End If
										'Enter Source Document ID.
										If dicPartDetailsCreate("SrcDocID") <> "" Then
											objDialogNewPart.JavaEdit("SourceDocID").Set dicPartDetailsCreate("SrcDocID")
										End If
								End If
		End Select

		'Click on "Finish" button
		objDialogNewPart.JavaButton("Finish").WaitProperty "enabled" , 1, 20000
		objDialogNewPart.JavaButton("Finish").Click
		'Click on "Close" button
		objDialogNewPart.JavaButton("Close").WaitProperty "enabled" , 1, 20000
		objDialogNewPart.JavaButton("Close").Click
		Call Fn_ReadyStatusSync(2)
		Fn_ADS_PartDetailCreateDic = sAssignId+"-"+sAssignRevId
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Part [" + sPartName + "] Created Successfully")

End Function

'*********************************************************		Function to create detail Design		***********************************************************************
'Function Name		:        Fn_ADS_DesignDetailCreateDic  

'Description	    	:        Creates an Design with detail information

'Parameters		     :    		sDesignType: Design type to be selected
'			                         	sDesignID: Unique ID for the Design [if non-empty, then enter]
'							          	sDesignRevID: Revision of the Design [if non-empty, then enter] - if any one of the fields (id/rev) are blank then click Assign button
'									 	sDesignName: Name of the Design
'									  	sDesignDesc: Description of the Design
' 										dicDesignDetailsCreate : Dictionary paramter  for detail creation
'	Example						Set dicDesignDetailsCreate = CreateObject( "Scripting.Dictionary" )
'										dicDesignDetailsCreate.RemoveAll
'											dicDesignDetailsCreate("DesignAddInfo") = "yes"
'											dicDesignDetailsCreate("DesignCategory") = "Open123"
'											dicDesignDetailsCreate("OrgCageCode") = "pune123456"
'											dicDesignDetailsCreate("SrcDocID") = "000056"
'										Msgbox Fn_ADS_DesignDetailCreateDic("CageDesign","","","TestDesign","Testing Design creation.",dicDesignDetailsCreate)
'										Set dicDesignDetailsCreate = Nothing
'Return Value		: 			DesignID-DesignRevID 

'Pre-requisite	    :		 	Should be logged in

'Examples		    :			Call Fn_ADS_DesignDetailCreateDic("Design","","","TesDesign","Testing Design creation.",dicDesignDetailsCreate)

'History		    :		
'													Developer Name				Date						Rev. No.			
'--------------------------------------------------------------------------------------------------------------------------
'													Ketan Raje				     23/06/2011			           1.0								
'--------------------------------------------------------------------------------------------------------------------------
Public Function Fn_ADS_DesignDetailCreateDic(sDesignType, sDesignID, sDesignRevID, sDesignName, sDesignDesc, dicDesignDetailsCreate)
		GBL_FAILED_FUNCTION_NAME="Fn_ADS_DesignDetailCreateDic"
		Dim sAssignId, sAssignRevId, aProjectName, objDialogNewDesign, ObjStaticText, bReturn

		On Error Resume Next
		'Creating Object for New Design window.
		JavaWindow("ADS-TeamCenter").JavaWindow("TcDefaultApplet").JavaDialog("New Item").SetTOProperty "title","New Design"
		set objDialogNewDesign = JavaWindow("ADS-TeamCenter").JavaWindow("TcDefaultApplet").JavaDialog("New Item")
		 'Creating Object of links on the left side of the window
		Set ObjStaticText =JavaWindow("ADS-TeamCenter").JavaWindow("TcDefaultApplet").JavaDialog("New Item").JavaStaticText("Steps")

		If not objDialogNewDesign.Exist (5)  Then
			'Select menu [File -> New -> Design...]
			bReturn = Fn_MenuOperation("Select","File:New:Design...")
			Wait(10)
			If bReturn = False Then
					Fn_ADS_DesignDetailCreateDic = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Select Menu [File:New:Design]")
					Set objDialogNewDesign = Nothing
					Exit Function
			Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Selected Menu [File:New:Design]")
			End If
		End If
	   'Check the existence of the "NewDesign" Window
		If objDialogNewDesign.Exist (20)  Then
					'Select  "Design Type"
					objDialogNewDesign.JavaList("ItemType").Select sDesignType
					If Err.Number < 0 Then
							Fn_ADS_DesignDetailCreateDic = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Select Design Type [" + sDesignType + "]")
							Set objDialogNewDesign = Nothing
							Exit Function
					End If
					'Click on "Next" button
					objDialogNewDesign.JavaButton("Next").WaitProperty "enabled", 1, 200000
					objDialogNewDesign.JavaButton("Next").Click
					If Err.Number < 0 Then
							Fn_ADS_DesignDetailCreateDic = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Click [Next] Button")
							Set objDialogNewDesign = Nothing
							Exit Function
					End If
					'Enter Design ID
					If sDesignID <> "" Then
						 objDialogNewDesign.JavaEdit("ItemID").Set sDesignID
					End If
					'Enter Revision ID
					If sDesignRevID <> "" Then
						objDialogNewDesign.JavaEdit("RevisionID").Set sDesignRevID
					End If
					'Check  "Design Id and Revision ID"
					If sDesignID = "" or sDesignRevID = "" Then
							'Click on "Assign" button
							objDialogNewDesign.JavaButton("Assign").WaitProperty "enabled", 1, 20000
							objDialogNewDesign.JavaButton("Assign").Click
					End If
					'Extract Design Id and Rev Id
					sAssignId = objDialogNewDesign.JavaEdit("ItemID").GetROProperty("value")
					sAssignRevId = objDialogNewDesign.JavaEdit("RevisionID").GetROProperty("value")
					'Set the Design Name
					If sDesignName <> "" Then
						objDialogNewDesign.JavaEdit("ItemName").Set sDesignName
					End If
					If sDesignDesc <> "" Then
						objDialogNewDesign.JavaEdit("Description").Set sDesignDesc
					End If
		Else
					Fn_ADS_DesignDetailCreateDic = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[New Design] Dialog Not Found")
		End If

		Select Case sDesignType
					Case "CageDesign"
								If Trim(Lcase(dicDesignDetailsCreate("DesignAddInfo"))) = "yes" Then
										'Click on link Enter Design Options and Details
										ObjStaticText.SetTOProperty "label","Enter Additional Design Information"
										ObjStaticText.WaitProperty "enabled" , 1, 20000
										ObjStaticText.Click 1, 1
										'Enter Design Category.
										If dicDesignDetailsCreate("DesignCategory") <> "" Then
											objDialogNewDesign.JavaEdit("Design Category").Set dicDesignDetailsCreate("DesignCategory")
										End If
										'Enter Original Cage Code.
										If dicDesignDetailsCreate("OrgCageCode") <> "" Then
											objDialogNewDesign.JavaEdit("CageCode").Set dicDesignDetailsCreate("OrgCageCode")
										End If
										'Enter Source Document ID.
										If dicDesignDetailsCreate("SrcDocID") <> "" Then
											objDialogNewDesign.JavaEdit("SourceDocID").Set dicDesignDetailsCreate("SrcDocID")
										End If
								End If
		End Select

		'Click on "Finish" button
		objDialogNewDesign.JavaButton("Finish").WaitProperty "enabled" , 1, 20000
		objDialogNewDesign.JavaButton("Finish").Click
		'Click on "Close" button
		objDialogNewDesign.JavaButton("Close").WaitProperty "enabled" , 1, 20000
		objDialogNewDesign.JavaButton("Close").Click
		Call Fn_ReadyStatusSync(2)
		Fn_ADS_DesignDetailCreateDic = sAssignId+"-"+sAssignRevId
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Design [" + sDesignName + "] Created Successfully")

End Function
 '*********************************************************		Function to Create Specification in SE ***********************************************************************
'Function Name		:				Fn_ADS_CustomNoteCreate

'Description			 :		 		 This function is used to Create the Custom Note in System Engineering

'Parameters			   :	 			1. sNodeName: Select the Note Spec
'													2. sCustDesc: ID of the Specification
'												   3. sCustID: Revision of the Spec
'												  4. strCusrName: Name of the Spec
'												  5.sCustUOM: Description of the Spec

'Return Value		   : 				TRUE \ FALSE

'Pre-requisite			:		 		should be in SE Prespective

'Examples				:				 Msgbox Fn_ADS_CustomNoteCreate("Set","Custom Note","TestingCustNote","CSTMNOTE-000001","CustName","", "Finish:Close","")

'History:
'										Developer Name			Date			Rev. No.			Reviewer			Build			Changes Done
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'						 				Ketan Raje					24/10/2011		1.0					Harshal A.			20110928		Handled JavaEdit by using descriptive code.	
'						 				Sandeep N					09/08/2012		1.1				Swapna											remove descriptive coding to enter value in edit box and added UI call to enter value in edit box
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function Fn_ADS_CustomNoteCreate(sAction, sNodeName,sCustDesc,sCustID,sCustName,sCustUOM, sButtons, dicCustNote)
	GBL_FAILED_FUNCTION_NAME="Fn_ADS_CustomNoteCreate"
	Dim sNodePath, sNodePathC, ObjCustNote, objElement, intNoOfObjects, aButtons

	Fn_ADS_CustomNoteCreate=False

	'Verifying "New CustomNote window" window's existance
	If Fn_UI_ObjectExist("Fn_ADS_CustomNoteCreate",JavaWindow("ADS-TeamCenter").JavaWindow("NewCustomNote"))=False Then
		'Invoking "New CustomNote" Window
		Call Fn_MenuOperation("Select","File:New:Custom Note")
	End If
	'Creating Object of "New CustomNote" window
	Set ObjCustNote=Fn_UI_ObjectCreate("Fn_ADS_CustomNoteCreate",JavaWindow("ADS-TeamCenter").JavaWindow("NewCustomNote"))
	Call Fn_UI_JavaTree_Expand("Fn_ADS_CustomNoteCreate", ObjCustNote, "CustomNoteTree","Complete List")
	JavaWindow("ADS-TeamCenter").JavaWindow("NewCustomNote").JavaTree("CustomNoteTree").WaitProperty "items count" , micGreaterThan(1)
	If Fn_UI_JavaTree_NodeExist("Fn_SE_NoteSpecCreate",ObjCustNote.JavaTree("CustomNoteTree"),"Complete List:"+sNodeName) Then
			sNodePathC="Complete List:"+sNodeName
	Else
			sNodePathC="Most Recently Used:"+sNodeName
	End If
		   Call Fn_JavaTree_Select("Fn_ADS_CustomNoteCreate", ObjCustNote, "CustomNoteTree",sNodePathC)
		   Call Fn_JavaTree_Select("Fn_ADS_CustomNoteCreate", ObjCustNote, "CustomNoteTree","Complete List")
		   Call Fn_JavaTree_Select("Fn_ADS_CustomNoteCreate", ObjCustNote, "CustomNoteTree",sNodePathC)
		   Call Fn_Button_Click("Fn_ADS_CustomNoteCreate",ObjCustNote,"Next")
			Wait(2)
		Select Case sAction
						Case "Set"
									'Setting Description.
									If sCustDesc<>"" Then
	                                        ObjCustNote.JavaStaticText("CustomNote_Text").SetTOProperty "label","Description:"
											Call Fn_Edit_Box("Fn_ADS_CustomNoteCreate",ObjCustNote,"CustomNote_Edit",sCustDesc)
									End If
									'Setting ID
									If sCustID<>"" Then
											ObjCustNote.JavaStaticText("CustomNote_Text").SetTOProperty "label","ID:"
											Call Fn_Edit_Box("Fn_ADS_CustomNoteCreate",ObjCustNote,"CustomNote_Edit",sCustID)
									End If
									'Setting Name.
									If sCustName <> "" Then
											ObjCustNote.JavaStaticText("CustomNote_Text").SetTOProperty "label","Name:"
											Call Fn_Edit_Box("Fn_ADS_CustomNoteCreate",ObjCustNote,"CustomNote_Edit",sCustName)
									End If
									'Setting Unit Of Measure.
									If sCustUOM <> "" Then
											ObjCustNote.JavaStaticText("CustomNote_Text").SetTOProperty "label","Unit of Measure:"
											Call Fn_Edit_Box("Fn_ADS_CustomNoteCreate",ObjCustNote,"CustomNote_Edit",sCustUOM)
									End If
									'Function Return True
									Fn_ADS_CustomNoteCreate=True
		End Select
	 'Click on Buttons
	 If sButtons<>"" Then
	   aButtons = split(sButtons, ":",-1,1)	  
	   For iCount=0 to Ubound(aButtons)
		Call Fn_Button_Click("Fn_ADS_CustomNoteCreate", ObjCustNote, aButtons(iCount))
		Call Fn_ReadyStatusSync(2)
	   Next
	 End If

	'Releasing "New CustomNote" window's object
	Set ObjChangeWnd=Nothing
	Set objElement=Nothing
	Set intNoOfObjects=Nothing
End Function
''*********************************************************		Function to perform action on Library Tree	***********************************************************************
'Function Name		:				Fn_SISW_ADS_JavaTable_GetCellData()
'Description		:		 		 For library tree in Project
'Parameters			:	 			objTable, iRow, iCol
'Return Value		: 				true/false.
'Pre-requisite		:		 		Project Prespective is Open.
'Examples			:				
'													bReturn = Fn_SISW_ADS_JavaTable_GetCellData(objTable, 1, "Object")
'History			:		
'		Developer Name			Date						Rev. No.	
'		-----------------------------------------------------------------------------------------------------------------
'		Koustubh W					08-08-2012			1.1						
'		-----------------------------------------------------------------------------------------------------------------
'*******************************************************************************************************************************************************************************************
Public Function Fn_SISW_ADS_JavaTable_GetCellData(objTable, iRow, iCol)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_ADS_JavaTable_GetCellData"
	Dim sPropName, sColName
	Dim sOutVal

	Fn_SISW_ADS_JavaTable_GetCellData = False

	'Code Commented by Sachin on 23-may-2012 as discussed with Vallari.
	'Reason : this code returns the old value from table, and sOutVal is not empty. therefore, the function returns old value. 
	'due to this Fn_MyTc_DetailTableContentOperation() case "Rowcellexist" returns old value and test fails.

	'sOutVal = objTable.GetCellData(iRow, iCol)

	'If trim(sOutVal) = "" Then
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
			Case "Name"
                sPropName = "object_name"	
		End Select
	
		If trim(sColName) = "Relation" Then
			sOutVal = objTable.Object.getItem(iRow).getData().getContext().toString()
			Select Case trim(sOutVal)
				Case "Fnd0ListsParamReqments"
					Fn_SISW_ADS_JavaTable_GetCellData = "Standard Notes Lists"
				Case "CMHasProblemItem"
					Fn_SISW_ADS_JavaTable_GetCellData = "Problem Items"
				Case "CMHasImpactedItem"
					Fn_SISW_ADS_JavaTable_GetCellData = "Impacted Items"
				Case "CMReferences"
					Fn_SISW_ADS_JavaTable_GetCellData = "Reference Items"
				Case "IMAN_reference"
					Fn_SISW_ADS_JavaTable_GetCellData = "References"
				Case "IMAN_specification"
					Fn_SISW_ADS_JavaTable_GetCellData = "Specifications"
				Case "contents"
					Fn_SISW_ADS_JavaTable_GetCellData = "Contents"
				Case "revision_list"
					Fn_SISW_ADS_JavaTable_GetCellData = "Revisions"
				Case "IMAN_master_form"
					' aaded s at the end - snehal salunkhe - 3-Apr-12 
					Fn_SISW_ADS_JavaTable_GetCellData = "Item Masters"
				Case "IMAN_classification"
					Fn_SISW_ADS_JavaTable_GetCellData = "Classification"
				Case "TC_Attaches"
					Fn_SISW_ADS_JavaTable_GetCellData = "Attaches"
                Case "IMAN_Rendering"
					Fn_SISW_ADS_JavaTable_GetCellData = "Rendering"
				Case "IMAN_manifestation"
					Fn_SISW_ADS_JavaTable_GetCellData = "Manifestations"
				Case "IMAN_aliasid"
					Fn_SISW_ADS_JavaTable_GetCellData = "Alias IDs"
				Case "release_status_list"
					Fn_SISW_ADS_JavaTable_GetCellData = "Release Status"
				Case "Fnd0ListsCustomNotes"
					Fn_SISW_ADS_JavaTable_GetCellData = "Custom Notes Lists"
				Case Else
					Fn_SISW_ADS_JavaTable_GetCellData = sOutVal
			End Select
		Else
			Fn_SISW_ADS_JavaTable_GetCellData = objTable.Object.getItem(iRow).getData().getComponent().getProperty(sPropName)
		End If
'	Else
'		Fn_SISW_ADS_JavaTable_GetCellData = sOutVal	
'	End If
End Function

''*********************************************************		Function to perform operations on Assign Company Location  *****************************************************
'Function Name		:				Fn_ADS_AssignCompanyLocation()
'Description		:		 		 Perform Operations on Assign Company Location Window
'Return Value		: 				true/false.
'Examples			:               Call Fn_ADS_AssignCompanyLocation("Add","Engineering:Designer:AutoTest5 (autotest5)","Newyork [  ]","","","OK")
'Examples			:               Call Fn_ADS_AssignCompanyLocation("MultiSelect","AutoAdminGrp~AutoGrp1","Mumbai [  ]~Pune [  ]","","","")
'Examples			:               Call Fn_ADS_AssignCompanyLocation("Deselect","AutoAdminGrp~AutoGrp1","Mumbai [  ]~Pune [  ]","","","")
'Examples			:               Call Fn_ADS_AssignCompanyLocation("VerifyMessage", "",  "","Multiple users or groups are selected." & vblf & " As a consequence, only one location may be chosen.", "", "Cancel")

'History			:				
'		Developer Name			Date						
'		-----------------------------------------------------------
'		Pranav S					09-11-2012										
'		-----------------------------------------------------------
'		Koustubh W					17-06-2013		Added cases MultiSelect, Deselect, VerifyMessage
'		-----------------------------------------------------------
Public Function Fn_ADS_AssignCompanyLocation(sAction,sOrgTree,sCompLocTree,aRelType,sCombo,sButtons)
	GBL_FAILED_FUNCTION_NAME="Fn_ADS_AssignCompanyLocation"
	Fn_ADS_AssignCompanyLocation = False
	Dim ObjAssCompLoc,ObjSelRelType,sOrgUser,sCompLoc,sNodeName,iCounter,sExpand,sNodeName1
	Dim arrOrgTreePath, arrCompLocTreePath, iCnt
	'Set Hierarchy
	Set ObjAssCompLoc=JavaWindow("ADS-TeamCenter").JavaWindow("Assign Company Location")
	Set ObjSelRelType=JavaWindow("ADS-TeamCenter").JavaWindow("Assign Company Location").JavaWindow("Select A RelationType")
	Set sOrgUser=JavaWindow("ADS-TeamCenter").JavaWindow("Assign Company Location").JavaTree("OrganizationTree")
	Set sCompLoc=JavaWindow("ADS-TeamCenter").JavaWindow("Assign Company Location").JavaTree("CompanyLocationTree")

	'Check Existance of Assign Company Location Window
	IF Fn_UI_ObjectExist("Fn_ADS_AssignCompanyLocation",ObjAssCompLoc)=False Then
        Call Fn_MenuOperation("Select","Tools:Assign Company Location")
		Call Fn_ReadyStatusSync(2)
	End If

	Select Case sAction
		Case "Deselect"
			If sOrgTree <> "" Then
				arrOrgTreePath = split(sOrgTree, "~")
				For iCnt = 0 to uBound(arrOrgTreePath)
					sNodeName = split(arrOrgTreePath(iCnt),":",-1,1)
					sExpand = ""
					For iCounter=0 to ubound(sNodeName)-1
						If iCounter=0 Then
							sExpand=sNodeName(0)
						else
							sExpand = sExpand+":"+sNodeName(iCounter)
						End If
						Call Fn_ReadyStatusSync(2)
						Call Fn_UI_JavaTree_Expand("Fn_ADS_AssignCompanyLocation",ObjAssCompLoc,"OrganizationTree",sExpand)
						Call Fn_ReadyStatusSync(2)
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Expanded the given node "+sExpand)
					Next
					'Check the Selected Item from Organization Tree
					sOrgUser.SetItemState arrOrgTreePath(iCnt), micUnchecked
				Next
			End If
			
			If sCompLocTree <> "" Then
				arrCompLocTreePath = split(sCompLocTree, "~")
				For iCnt = 0 to uBound(arrCompLocTreePath)
					sNodeName = split(arrCompLocTreePath(iCnt),":",-1,1)
					sExpand = ""
					For iCounter=0 to ubound(sNodeName)-1
						If iCounter=0 Then
							sExpand=sNodeName(0)
						else
							sExpand = sExpand+":"+sNodeName(iCounter)
						End If
						Call Fn_ReadyStatusSync(2)
						Call Fn_UI_JavaTree_Expand("Fn_ADS_AssignCompanyLocation",ObjAssCompLoc,"CompanyLocationTree",sExpand)
						Call Fn_ReadyStatusSync(2)
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Expanded the given node "+sExpand)
					Next
					'Check the Selected Item from Organization Tree
					sCompLoc.SetItemState arrCompLocTreePath(iCnt), micUnchecked
				Next
			End If
			Fn_ADS_AssignCompanyLocation = True	
		Case "MultiSelect"
			If sOrgTree <> "" Then
				arrOrgTreePath = split(sOrgTree, "~")
				For iCnt = 0 to uBound(arrOrgTreePath)
					sNodeName = split(arrOrgTreePath(iCnt),":",-1,1)
					sExpand = ""
					For iCounter=0 to ubound(sNodeName)-1
						If iCounter=0 Then
							sExpand=sNodeName(0)
						else
							sExpand = sExpand+":"+sNodeName(iCounter)
						End If
						Call Fn_ReadyStatusSync(2)
						Call Fn_UI_JavaTree_Expand("Fn_ADS_AssignCompanyLocation",ObjAssCompLoc,"OrganizationTree",sExpand)
						Call Fn_ReadyStatusSync(2)
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Expanded the given node "+sExpand)
					Next
					'Check the Selected Item from Organization Tree
					sOrgUser.SetItemState arrOrgTreePath(iCnt), micChecked
				Next
			End If
			
			If sCompLocTree <> "" Then
				arrCompLocTreePath = split(sCompLocTree, "~")
				For iCnt = 0 to uBound(arrCompLocTreePath)
					sNodeName = split(arrCompLocTreePath(iCnt),":",-1,1)
					sExpand = ""
					For iCounter=0 to ubound(sNodeName)-1
						If iCounter=0 Then
							sExpand=sNodeName(0)
						else
							sExpand = sExpand+":"+sNodeName(iCounter)
						End If
						Call Fn_ReadyStatusSync(2)
						Call Fn_UI_JavaTree_Expand("Fn_ADS_AssignCompanyLocation",ObjAssCompLoc,"CompanyLocationTree",sExpand)
						Call Fn_ReadyStatusSync(2)
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Expanded the given node "+sExpand)
					Next
					'Check the Selected Item from Organization Tree
					sCompLoc.SetItemState arrCompLocTreePath(iCnt), micChecked
				Next
			End If
			Fn_ADS_AssignCompanyLocation = True
		Case "VerifyMessage"
			ObjAssCompLoc.JavaStaticText("Message").setTOProperty "label", aRelType
			Fn_ADS_AssignCompanyLocation = Fn_UI_ObjectExist("Fn_ADS_AssignCompanyLocation", ObjAssCompLoc.JavaStaticText("Message"))			

		Case "Add"
			sNodeName = split(sOrgTree,":",-1,1)
			sExpand = ""
			For iCounter=0 to ubound(sNodeName)-1
				If iCounter=0 Then
					sExpand=sNodeName(0)
				else
					sExpand = sExpand+":"+sNodeName(iCounter)
				End If
				Call Fn_ReadyStatusSync(2)
				Call Fn_UI_JavaTree_Expand("Fn_ADS_AssignCompanyLocation",ObjAssCompLoc,"OrganizationTree",sExpand)
				Call Fn_ReadyStatusSync(2)
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Expanded the given node "+sExpand)
			Next
		
			'Check the Selected Item from Organization Tree
			sOrgUser.SetItemState sOrgTree, micChecked

			sNodeName1 = split(sCompLocTree,":",-1,1)
			For iCounter=0 to ubound(sNodeName1)-1
				If iCounter=0 Then
					sExpand=sNodeName1(0)
				else
					sExpand=sExpand+":"+sNodeName1(iCounter)
				End If
				Call Fn_UI_JavaTree_Expand("Fn_ADS_AssignCompanyLocation",ObjAssCompLoc,"CompanyLocationTree",sExpand)
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Expanded the given node "+sExpand)
			Next

			'Check the Selected Item from Company Location Tree
			sCompLoc.SetItemState sCompLocTree,micChecked 
			Call Fn_ReadyStatusSync(2)
			'Click on Add button
			Call Fn_Button_Click("Fn_ADS_AssignCompanyLocation", ObjAssCompLoc, "Add")
			Call Fn_ReadyStatusSync(2)
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Clicked on Add button of Relation Type Window")

					'Select the Relation Type
					If aRelType="DesignAuthorityAffiliation" Then
						Call Fn_UI_JavaRadioButton_SetON("Fn_ADS_AssignCompanyLocation",ObjSelRelType, "DesignAuthorityAffiliation")
					Else
						Call Fn_UI_JavaRadioButton_SetON("Fn_ADS_AssignCompanyLocation",ObjSelRelType, "TrueCompanyAffiliation")
					End If
		
					'Click on OK button of Relation Type
					Call Fn_Button_Click("Fn_ADS_AssignCompanyLocation", ObjSelRelType, "OK")
					Call Fn_ReadyStatusSync(2)
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Clicked on OK button of Relation Type Window")
					Fn_ADS_AssignCompanyLocation = TRUE

			'Click on OK button on Assign Company Location Window
			Call Fn_Button_Click("Fn_ADS_AssignCompanyLocation", ObjAssCompLoc, "OK")
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Clicked on OK button of Asign Company Location Window")
			Fn_ADS_AssignCompanyLocation = TRUE
			
			
	Case "Remove"

			sNodeName1 = split(sCompLocTree,":",-1,1)
			sExpand = ""
			For iCounter=0 to ubound(sNodeName1)-1
				If iCounter=0 Then
					sExpand=sNodeName1(0)
				else
					sExpand=sExpand+":"+sNodeName1(iCounter)
				End If
				Call Fn_UI_JavaTree_Expand("Fn_ADS_AssignCompanyLocation",ObjAssCompLoc,"CompanyLocationTree",sExpand)
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Expanded the given node "+sExpand)
			Next
			wait 3		
			'sCompLocTree = sCompLocTree&" [ "+aRelType+" ]"		
			'Check the Selected Item from Company Location Tree
			sCompLoc.SetItemState sCompLocTree,micChecked 
			Call Fn_ReadyStatusSync(2)
			'Click on Remove button
			Call Fn_Button_Click("Fn_ADS_AssignCompanyLocation", ObjAssCompLoc, "Remove")
			Call Fn_ReadyStatusSync(2)
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Clicked on Remove button of Relation Type Window")				

			'Click on OK button on Assign Company Location Window
			Call Fn_Button_Click("Fn_ADS_AssignCompanyLocation", ObjAssCompLoc, "OK")
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Clicked on OK button of Asign Company Location Window")
			Fn_ADS_AssignCompanyLocation = TRUE		

	Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_ADS_AssignCompanyLocation function failed")
			Fn_ADS_AssignCompanyLocation = FALSE
			Exit Function		
	End Select
	If sButtons <> "" Then
		Call Fn_Button_Click("Fn_ADS_AssignCompanyLocation", ObjAssCompLoc, sButtons)
	End If
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" Sucessfully completed function Fn_ADS_AssignCompanyLocation")
	Set ObjAssCompLoc = nothing 
	Set ObjAssCompLoc = nothing 
	Set sOrgUser=nothing
	Set sCompLoc=nothing
End Function

'*********************************************************	Generic function to handle Error dialogs in ADS Module  	***********************************************************************
'Function Name		:		Fn_SISW_ADS_ErrorVerify()

'Description		:	The function is generic function to handle error dialogs. It is created after combining error dialog functions from ADS.vbs
'							Fn_ADS_DialogHandle
'							Fn_ADS_DialogMsgVerify

'Parameters			 :	 			1.  dicErrorInfo
											
'Return Value		 : 				True/False

'Pre-requisite		 :		 		NA.

'Examples			 :				 Dim dicErrorInfo
'												 Set dicErrorInfo = CreateObject("Scripting.Dictionary")
'												 With dicErrorInfo 
'												  .Add "Object" , JavaWindow("ADS-TeamCenter").JavaWindow("ADSDialog")
'												  .Add "Title", "Enter the values for Properties on Relation"
'												  .Add "Button", "Finish:Close"
'												  .Add "Action", "DialogHandle" 	  
'												 End with
'											   bReturn = Fn_SISW_ADS_ErrorVerify(dicErrorInfo)
'History:
'										Developer Name			Date			Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Sushma Pagare          5-Jul-2013
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
public Function Fn_SISW_ADS_ErrorVerify(dicErrorInfo)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_ADS_ErrorVerify"
	Dim dicKeys, dicItems, iCounter
	Dim sAction, sTitle, sErrorMsg,sButton, sAppMsg
	Dim objErrorDialog, aButtons,  sObject

	On Error Resume Next
	Fn_SISW_ADS_ErrorVerify = False

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
			Case "Object"
					Set sObject =  dicItems(iCounter)				
			Case "Button"
					sButton = dicItems(iCounter)
		End Select
	Next
	
	Select Case sAction

		''  This covers Fn_ADS_DialogHandle(sObject,sTitle,sButtons)
		Case "DialogHandle"
            			
				 Set objErrorDialog = sObject
				 objErrorDialog.SetTOProperty "title",sTitle
				 If objErrorDialog.Exist = True Then
					   Fn_SISW_ADS_ErrorVerify=true
					   Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,sTitle+" Dialog Exist")
				 Else
					   Fn_SISW_ADS_ErrorVerify=false
					   Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,sTitle+" Dialog does not Exist") 
				 End If
				 'Click on Buttons
				 If sButton<>"" Then
					   aButtons = split(sButton, ":",-1,1)				   
					   For iCounter=0 to Ubound(aButtons)
							Call Fn_Button_Click("Fn_SISW_ADS_ErrorVerify", objErrorDialog, aButtons(iCount))
							Call Fn_ReadyStatusSync(2)
					   Next
				 End If
				Set objErrorDialog = Nothing
				Exit Function
                
		'' Case covers Fn_ADS_DialogMsgVerify(sErrMsg,sDialogTitle,sButton)
		Case "DialogMsgVerify"			

			Set objErrorDialog = JavaWindow("ADS-TeamCenter").JavaWindow("MsgDialog")					
			objErrorDialog.SetTOProperty "title", sTitle

			If  objErrorDialog.Exist(2) Then
				If sErrorMsg <> ""  Then
					sAppMsg = objErrorDialog.JavaStaticText("DialogMsg").GetROProperty("label")					
					If instr(1,sAppMsg,sErrMsg)<> 0 Then       '' Check Actual Message with expected Message
							Call Fn_Button_Click("Fn_SISW_ADS_ErrorVerify",objErrorDialog,sButton)
							Call Fn_WriteLogFile(Environment.Value("TestLogFile")," Sucessfully Verified Message"+sErrorMsg)
							Fn_SISW_ADS_ErrorVerify = True
						  Else
						  	GBL_ACTUAL_MESSAGE=sAppMsg
							Call Fn_WriteLogFile(Environment.Value("TestLogFile")," Message Verification Failed"+sErrorMsg)
							Fn_SISW_ADS_ErrorVerify = False
						End if
				Else
						Call Fn_Button_Click("Fn_SISW_ADS_ErrorVerify",objErrorDialog,sButton)
						Call Fn_WriteLogFile(Environment.Value("TestLogFile")," Sucessfully Verified Message"+sErrorMsg)
						Fn_SISW_ADS_ErrorVerify = True
				End If
			Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile")," Message Verification Failed"+sErrorMsg)
					Fn_SISW_ADS_ErrorVerify = False
			End If 
			Set objErrorDialog = nothing
			Exit Function			
			
	End Select

End Function
'********************************************************************************************************************************
'Function Name		:	Fn_ADS_ChangeOwningProgram_Ops()
'
'Description		:	The function is for performing operations on Change Owning Program
'
'Parameters			:	 1.sAction : Action Name
'						 2.dicInfo : Dictionary Object
'
'Return Value		 : 	True/False
'
'Examples			 :	Set dicInfo = CreateObject("Scripting.Dictionary")
'						dicInfo("ProgramName") = "PC_DocProj2"
'						bReturn = Fn_ADS_ChangeOwningProgram_Ops("ChangeOwningProgram",dicInfo)
'											   
'History			: 	Developer Name				Date				Rev. No.			Changes Done			Reviewer	
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'						Poonam Chopade          	14-June-2017		1.0					Created					TC11.4_(20170605.00)_NewDevelopment_PoonamC_14Jun2017
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
public Function Fn_ADS_ChangeOwningProgram_Ops(sAction,dicInfo)
	
	GBL_FAILED_FUNCTION_NAME="Fn_ADS_ChangeOwningProgram_Ops"
	Dim objChangeDialog,sMenu
	
	Fn_ADS_ChangeOwningProgram_Ops = False
	Set objChangeDialog = Fn_ADS_SISW_GetObject("ChangeOwningProgram")
	'Check existence of dialog
	 If Fn_UI_ObjectExist("Fn_ADS_ChangeOwningProgram_Ops", objChangeDialog) = False Then
		sMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("RAC_Menu"),"ToolsChangeOwningProgram")
		Call Fn_MenuOperation("Select",sMenu)
		Call Fn_ReadyStatusSync(1)
		
		If Fn_UI_ObjectExist("Fn_ADS_ChangeOwningProgram_Ops", objChangeDialog) = False  Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_ADS_ChangeOwningProgram_Ops ] Failed to open Change Owning Program window.")
				Set objChangeDialog = Nothing
				Exit function
		End if	
	End If
	
	Select Case sAction
	
			Case "ChangeOwningProgram"
			
				If dicInfo("ProgramName") <> "" Then
						'Select Program from list
						Call Fn_List_Select("Fn_ADS_ChangeOwningProgram_Ops",objChangeDialog,"CCombo",dicInfo("ProgramName"))
						Call Fn_ReadyStatusSync(1)
						
						'Click OK
						 Call Fn_Button_Click("Fn_ADS_ChangeOwningProgram_Ops",objChangeDialog,"OK") 
						 Call Fn_ReadyStatusSync(1)
						 Wait(7) 'Added wait for Inprogress dialog
						 
						 'Click OK on Message : Change owning program operation is successfully performed for the selected objects.
						 If Fn_UI_ObjectExist("Fn_ADS_ChangeOwningProgram_Ops", objChangeDialog) = False Then
						 	Fn_ADS_ChangeOwningProgram_Ops = Fn_Button_Click("Fn_ADS_ChangeOwningProgram_Ops",objChangeDialog.JavaWindow("ProgressInformation").JavaWindow("Change owning program"),"OK") 
						 	Call Fn_ReadyStatusSync(1)
						 End If							
			    End If				
	End Select

	Set objChangeDialog = Nothing

End Function

