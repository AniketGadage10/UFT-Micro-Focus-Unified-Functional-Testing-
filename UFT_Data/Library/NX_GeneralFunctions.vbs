Option Explicit
Dim iNXASMTreeStepCounter
iNXASMTreeStepCounter = 0
'===============================================================================================================
' Function List
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'000. Fn_SISW_NX_GetObject
'001. Fn_SISW_NX_General_JournalPlayBack
'002. Fn_SISW_NX_General_MenuOperation
'003. Fn_SISW_NX_General_ItemCreate
'004. Fn_SISW_NX_General_InsertGeometry
'005. Fn_SISW_NX_General_AssemblyAddComponent	
'006. Fn_SISW_NX_General_NavigatorTab_Operation
'007. Fn_SISW_NX_General_ANT_AssemblyVerify
'008. Fn_SISW_NX_General_OpenInNXOperation
'009. Fn_SISW_NX_General_FitToResolution
'010. Fn_SISW_NX_General_TC_NavTeeOperation
'011. Fn_SISW_NX_General_ExtraComponents 
'012. Fn_SISW_NX_General_PartDialog
'013. Fn_SISW_NX_General_SaveAs
'014. Fn_SISW_NX_General_CustomerDefault
'--------Functions for Briefcase Browser Tc's development with NX-----------
'015. Fn_SISW_NX_General_AssemblyCreate				(vivek.ahirrao.ext@siemens.com)
'016. Fn_SISW_NX_General_ModelCreate				(vivek.ahirrao.ext@siemens.com)
'017. Fn_NX_ApplicationLaunch						(vivek.ahirrao.ext@siemens.com)
'018. Fn_NX_ApplicationExit							(vivek.ahirrao.ext@siemens.com)
'019. Fn_SISW_NX_SaveOptions						(vivek.ahirrao.ext@siemens.com)
'020. Fn_SISW_NX_General_AddExistingComponent		(vivek.ahirrao.ext@siemens.com)
'021. Fn_SISW_NX_General_ComponentProperties_Ops	(vivek.ahirrao.ext@siemens.com)
'-------- Functions for DIPRO NX Tc's development with NX -----------
'022. Fn_SISW_NX_Find_Component_Operation
'023. Fn_SISW_NX_Find_Object_Operation
'024. Fn_SISW_NX_General_Model_ItemCreate
'025. Fn_SISW_NX_VisualReportingOperations
'026. Fn_SISW_NX_HandleMessageDialog
'027. Fn_SISW_NX_General_SaveAsExt
'028. Fn_SISW_NX_ClosePartOperations
'029. Fn_SISW_NX_Name_Parts_For_Save
'030. Fn_SISW_NX_NewItem_ItemType_Operation
'031. Fn_SISW_NX_OutputPartNumber_Operation
'032. Fn_SISW_NX_SaveAsNonMasterParts_Operation
'033. Fn_SISW_NX_AssemblyLoadOptions
'---------------------------------------------------------------------------
'****************************************    Function to get Object hierarchy ***************************************
'
''Function Name	:	Fn_SISW_NX_GetObject
'
''Description		:  	Function to get specified Object hierarchy.

''Parameters		:	1. sObjectName : Object Handle name
								
''Return Value		:  	Object \ Nothing
'
''Examples		:	Fn_SISW_NX_GetObject("Remove")

'History			:	Developer Name			Date				Rev. No.		Reviewer		Changes Done	
'------------------------------------------------------------------------------------------------------------------------------------
'Created By 		:	Snehal Salunkhe	   		06-Aug-2014	
'-----------------------------------------------------------------------------------------------------------------------------------
Function Fn_SISW_NX_GetObject(sObjectName)
	Dim sObjectXMLPath
	sObjectXMLPath = Fn_GetEnvValue("User", "AutomationDir") & "\TestData\AutomationXML\ObjectXML\NX.xml"
	Set Fn_SISW_NX_GetObject = Fn_SISW_Setup_GetObjectFromXML(sObjectXMLPath, sObjectName)
End Function

'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Function Name     			: Fn_SISW_NX_General_JournalPlayBack
' Function Description 			 : Function used to Run journal file in NX environment
' Parameters			: 			sFilePath: Journal file path
'											 sReserve:  Reserve for Future
' Return Value           :        True/False
' 
' Examples		    	: 		Call Fn_SISW_NX_General_JournalPlayBack("C:\mainline\TestData\NX\Journal\Journal_Item_Create.vb",sReserve1,sReserve2,sReserve3,sReserve4)

' History               :  
'			Developer Name			 Date	  					Rev. No. 				Changes					Reviewer
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'			Nilesh Gadekar  	    27-Aug-2013					1.0							
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'			Pranav Ingle  	    	  8-Dec-2013					1.1							
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_NX_General_JournalPlayBack(sFilePath,sReserve1,sReserve2,sReserve3,sReserve4)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_NX_General_JournalPlayBack"
   Dim objJournalDlg,bResult
	Set objJournalDlg=Window("NXWindow").Dialog("Journal Manager")
	Fn_SISW_NX_General_JournalPlayBack=False

	' Check Existance of 
	If Fn_SISW_UI_Object_Operations("Fn_SISW_NX_General_JournalPlayBack", "Exist", objJournalDlg, 5) =False Then
		'Call to Invoke Journal Play Dialog
		Call Fn_SISW_NX_General_MenuOperation("Select","Play Journal ","")
		Wait 2
	End If

    If  Fn_SISW_UI_Object_Operations("Fn_SISW_NX_General_JournalPlayBack", "Exist", objJournalDlg, 5) = False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Get Existence of Journal Manager Dialog")
		Exit Function
	End If

	Wait 1
	bResult=Fn_SISW_NX_UI_EditBoxOperation("Fn_SISW_NX_General_JournalPlayBack","Set",objJournalDlg,"FileName",sFilePath)
	If bResult=False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Set File path of Journal in FIle Name Editbox of Journal Manager Dialog")
		Exit Function
	Else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Set File path of Journal in FIle Name Editbox of Journal Manager Dialog")
	End If
	
	bResult=Fn_SISW_NX_UI_ButtonOperation("Fn_SISW_NX_General_JournalPlayBack", "Click", objJournalDlg, "Run")
	If bResult=False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Click on Run button of Journal Manager Dialog")
		Exit Function
	Else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Clicked on Run button of Journal Manager Dialog")
	End If
	Call Fn_SISW_NX_Setup_ReadyStatusSync(1)

	bResult= Fn_SISW_UI_Object_Operations("Fn_SISW_NX_General_JournalPlayBack", "Exist", objJournalDlg, 5) 
	If bResult=False Then
		Fn_SISW_NX_General_JournalPlayBack=True
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Ran Journal file ["& sFilePath &"] in NX")
	Else
		objJournalDlg.Type micEsc 
		Wait 2
		objJournalDlg.Close()
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Run Journal file ["&sFilePath&"] in NX please validate Journal File is present at specified path " & sFilePath)
	End If

	Set objJournalDlg=Nothing
End Function 

'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Function Name     			: Fn_SISW_NX_General_MenuOperation
' Function Description 			 : Function used t Perform Menu operations on NX application
' Parameters			: 			   sAction: Action
'												sMenu: Menu command
'												sReserve: Reserved for future use
'
' Return Value           :        True/False
'	
' Examples		    	: 		Call 	 Fn_SISW_NX_General_MenuOperation("Select","Journal Play ","")

' History               :  Developer Name			 	Date	  				Rev. No. 				Changes					Reviewer
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'								Pranav Ingle		 	    28-Aug-2013					1.0							
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_NX_General_MenuOperation(sAction,sMenu,sReserve)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_NX_General_MenuOperation"
   Dim objNXWindow
   Set objNXWindow=Window("NXWindow")

   Select Case sAction
	 	Case "Select"
					If sMenu= "macrorun" Then
							Window("NXWindow").Type  micLCtrlDwn +micShiftDwn +"P"+micLCtrlUp +micShiftUp
							wait 2
							Fn_SISW_NX_General_MenuOperation=True
					Else
							bReturn=Fn_SISW_NX_Setup_CmdFinderOperation(sMenu,"Start","")
							If bReturn=True Then
									Fn_SISW_NX_General_MenuOperation=True
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully performed Menu operation for ["&sMenu &"]")
							Else
									Fn_SISW_NX_General_MenuOperation=False
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to perform Menu operation for ["&sMenu &"]")
							End If
					End If
					wait 1
   End Select
End Function

'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Function Name     	  : 	  Fn_SISW_NX_General_ItemCreate
' Function Description 	: 		Function used to Create Item in NX
' Parameters			     : 		 sItemName: Item Name
'											  sFolderPath: Folder path separated by :
'											  sReserve: Reserved for future use
' Return Value           :          True/False
'
' Examples		    	: 		Call 	Fn_SISW_NX_General_ItemCreate(DataTable.Value("ItemName",dtGlobalSheet),Environment.Value("NXTestFolderName"),"","")
'
' History               :  
'			Developer Name	 		 	Date	  						Rev. No. 				Changes					Reviewer
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'			Nilesh Gadekar  	     28-Aug-2013					 1.0							
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'			Pranav Ingle		  	    28-Dec-2013						1.1
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_NX_General_ItemCreate(sItemName,sFolderPath,sReserve1,sReserve2)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_NX_General_ItemCreate"
'	Update the Item Information Journal folder of TestData
	Dim bResult,sItemId,sItemRev
	Fn_SISW_NX_General_ItemCreate=False

	bResult= Fn_UpdateEnvXMLNode(Environment.Value("sPath")&"\TestData\NX\Journal\ItemCreate_Input.xml", "ItemName",sItemName)
	If bResult=False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Update Item Name Information in ItemCreate_Input.xml")
			Exit Function
	End If

	bResult= Fn_UpdateEnvXMLNode(Environment.Value("sPath")&"\TestData\NX\Journal\ItemCreate_Input.xml", "Folder",sFolderPath)
	If bResult=False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to UpdateFolder Path in ItemCreate_Input.xml")
		Exit Function
	End If

    bResult=Fn_SISW_NX_General_JournalPlayBack(Environment.Value("sPath") &"\TestData\NX\Journal\Journal_Item_Create.vb","","","","")
	If  bResult=False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Ran Journal File to create Item in NX application")
		Exit Function
	End If
	Call Fn_SISW_NX_Setup_ReadyStatusSync(2)

	sItemId=Fn_GetXMLNodeValue(Environment.Value("sPath")&"\TestData\NX\Journal\ItemCreate_Output.xml","ItemID")
	sItemRev=Fn_GetXMLNodeValue(Environment.Value("sPath")&"\TestData\NX\Journal\ItemCreate_Output.xml","ItemRevision")
	sItemName=Fn_GetXMLNodeValue(Environment.Value("sPath")&"\TestData\NX\Journal\ItemCreate_Output.xml","ItemName")
	Fn_SISW_NX_General_ItemCreate="'"&sItemId&"~"&sItemRev &"~"& sItemName
End Function

'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Function Name     			: Fn_SISW_NX_General_InsertGeometry
' Function Description 			 : Function used to Insert Geometry in Item/Part
' Parameters			: 			sFeatureType: Modeling feature type i.e Block,Cylinder
'												dicFeatureDetails: Geometry Information
'												sReserve: Reserved for future use
' Return Value           :        True/False
' 
' Examples		    	: 		With dicFeatureDetails
'											.Add "XCord", 0
'											.Add "YCord", 0
'											.Add "ZCord", 0
'											.Add "Diameter", 10
'											.Add "Height", 50
'										End With
'										bReturn=Fn_SISW_NX_General_InsertGeometry("Cylinder",dicFeatureDetails,"")
' History               :  Developer Name			 Date	  					Rev. No. 				Changes					Reviewer
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'								Nilesh Gadekar  	    29-Aug-2013					1.0							
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'								Pranav Ingle		  	   28-Dec-2013				 1.1
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'                             Ashwini Patil					30-Jan-2014					1.2			Added Cases "Cone", "Sphere"
'__________________________________________________________________________________________________________________
Public Function Fn_SISW_NX_General_InsertGeometry(sFeatureType,dicFeatureDetails,sReserve)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_NX_General_InsertGeometry"
	Dim iDicCnt,DictItems,DictKeys,iCount,sAttName,sAttValue,bResult
	Fn_SISW_NX_General_InsertGeometry=False

	Select Case sFeatureType
			Case "Block"
						iDicCnt=dicFeatureDetails.Count
                        DictItems = dicFeatureDetails.Items
						DictKeys = dicFeatureDetails.Keys
						For iCount=0 to iDicCnt-1
							sAttName=DictKeys(iCount)
							sAttValue=DictItems(iCount)
							Call Fn_UpdateEnvXMLNode(Environment.Value("sPath")&"\TestData\NX\Journal\Modeling\BlockCreate_Input.xml", sAttName,sAttValue)
						Next
						
						bResult=Fn_SISW_NX_General_JournalPlayBack(Environment.Value("sPath") &"\TestData\NX\Journal\Modeling\Journal_BlockCreate.vb","","","","")
						If  bResult=False Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Ran Journal File to create Block in NX application")
							Exit Function
						Else
							Fn_SISW_NX_General_InsertGeometry=True
						End If
						Call Fn_SISW_NX_Setup_ReadyStatusSync(1)

			Case "Cylinder"
						iDicCnt=dicFeatureDetails.Count
                        DictItems = dicFeatureDetails.Items
						DictKeys = dicFeatureDetails.Keys
						For iCount=0 to iDicCnt-1
							sAttName=DictKeys(iCount)
							sAttValue=DictItems(iCount)
							Call Fn_UpdateEnvXMLNode(Environment.Value("sPath")&"\TestData\NX\Journal\Modeling\CylinderCreate_Input.xml", sAttName,sAttValue)
						Next

						bResult=Fn_SISW_NX_General_JournalPlayBack(Environment.Value("sPath") &"\TestData\NX\Journal\Modeling\Journal_CylinderCreate.vb","","","","")
						If  bResult=False Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Ran Journal File to create Cylinder in NX application")
							Exit Function
						Else
							Fn_SISW_NX_General_InsertGeometry=True
						End If
						Call Fn_SISW_NX_Setup_ReadyStatusSync(1)

			Case "Cone"
						iDicCnt=dicFeatureDetails.Count
                        DictItems = dicFeatureDetails.Items
						DictKeys = dicFeatureDetails.Keys
						For iCount=0 to iDicCnt-1
							sAttName=DictKeys(iCount)
							sAttValue=DictItems(iCount)
							Call Fn_UpdateEnvXMLNode(Environment.Value("sPath")&"\TestData\NX\Journal\Modeling\ConeCreate_Input.xml", sAttName,sAttValue)
						Next
						
						bResult=Fn_SISW_NX_General_JournalPlayBack(Environment.Value("sPath") &"\TestData\NX\Journal\Modeling\Journal_ConeCreate.vb","","","","")
						If  bResult=False Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Ran Journal File to create Cone in NX application")
							Exit Function
						Else
							Fn_SISW_NX_General_InsertGeometry=True
						End If
						Call Fn_SISW_NX_Setup_ReadyStatusSync(1)

			  Case "Sphere"
                        iDicCnt=dicFeatureDetails.Count
                        DictItems = dicFeatureDetails.Items
						DictKeys = dicFeatureDetails.Keys
						For iCount=0 to iDicCnt-1
							sAttName=DictKeys(iCount)
							sAttValue=DictItems(iCount)
							Call Fn_UpdateEnvXMLNode(Environment.Value("sPath")&"\TestData\NX\Journal\Modeling\SphereCreate_Input.xml", sAttName,sAttValue)
						Next
						
						bResult=Fn_SISW_NX_General_JournalPlayBack(Environment.Value("sPath") &"\TestData\NX\Journal\Modeling\Journal_SphereCreate.vb","","","","")
						If  bResult=False Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Ran Journal File to create Sphere in NX application")
							Exit Function
						Else
							Fn_SISW_NX_General_InsertGeometry=True
						End If
						Call Fn_SISW_NX_Setup_ReadyStatusSync(1)
	End Select
End Function

'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Function Name     			: Fn_SISW_NX_General_AssemblyAddComponent
' Function Description 			 : TO add components in Assembly
' Parameters			: 			dicCompInformaton: Component Information 
'									sReserve: For future use
' Return Value           :        True/False
' 
' Examples		    	: 			 Set dicCompInformaton=CreateObject("Scripting.Dictionary")
'												With dicCompInformaton
'													.Add "ComponentsIDs",sItemID1&"~" & sItemID2
'													.Add "ComponentRevisions", sRevId1 &"~" & sRevId2
'													.Add "3DPoints", "0.0,0.0,0.0~0.0,0.0,0.0"
'													.Add "Orientations", "1.0,0.0,0.0,0.0,1.0,0.0,0.0,0.0,1.0~1.0,0.0,0.0,0.0,1.0,0.0,0.0,0.0,1.0"
'												End With
'												bReturn= Fn_SISW_NX_General_AssemblyAddComponent(dicCompInformaton,"")
' History               :  
'			Developer Name			 	   Date	  					Rev. No. 				Changes					Reviewer
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'			Nilesh Gadekar  			 2-Sep-2013					1.0							
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'			Pranav Ingle		  	    	28-Dec-2013				 1.1
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_NX_General_AssemblyAddComponent(dicCompInformaton,sReserve1)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_NX_General_AssemblyAddComponent"
		Dim iDicCnt,DictItems,DictKeys,iCount,sAttName,sAttValue,bResult

		iDicCnt=dicCompInformaton.Count
		DictItems = dicCompInformaton.Items
		DictKeys = dicCompInformaton.Keys
		For iCount=0 to iDicCnt-1
			sAttName=DictKeys(iCount)
			sAttValue=DictItems(iCount)
			Call Fn_UpdateEnvXMLNode(Environment.Value("sPath")&"\TestData\NX\Journal\AddComponent.xml", sAttName,sAttValue)
		Next

		bResult=Fn_SISW_NX_General_JournalPlayBack(Environment.Value("sPath") &"\TestData\NX\Journal\Journal_AddComponent.vb","","","","")
		If bResult=False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Ran Journal File to create Cylinder in NX application")
			Fn_SISW_NX_General_AssemblyAddComponent=False
			Exit Function
		Else
			Fn_SISW_NX_General_AssemblyAddComponent=True
		End If
		Call Fn_SISW_NX_Setup_ReadyStatusSync(1)
End Function
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Function Name     			: Fn_SISW_NX_General_NavigatorTab_Operation
' Function Description 			 : To select Tab from navigator pane
' Parameters			: 			sTabName: Tab name to be selected
'											 sReserve: For future use
' Return Value           :        True/False
' 
' Examples		    	: 			 Fn_SISW_NX_General_NavigatorTab_Operation("HD3D Tools","")
' History               :  
'		Developer Name				 Date	  				Rev. No. 				Changes					Reviewer
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'		Nilesh Gadekar  			2-Sep-2013				1.0							
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'		Pranav Ingle		  	  	 28-Dec-2013			 1.1
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_NX_General_NavigatorTab_Operation(sTabName,sReserve)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_NX_General_NavigatorTab_Operation"

	Dim objTab,objTree,width,height,x,y,yfact
	Dim iCounter,bResult,bFlag
	Dim sFirstTab,sFlag
	Fn_SISW_NX_General_NavigatorTab_Operation=False

	Call Fn_SISW_NX_Setup_Restore()
	Set objTab=Window("NXWindow").WinTab("NavigatorTab")
	Set objTree=Window("NXWindow").WinObject("NavigatorPane")
	Fn_SISW_NX_General_NavigatorTab_Operation=False

	If objTree.Exist(5)=False Then
		bResult=Fn_SISW_NX_Setup_DisplayResourceBar("On Left")
		If  bResult=False Then
			Set objTab=Nothing
			Set objTree=Nothing
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to  set preference Display Resource Bar on Left of NX window")
			Exit Function
		Else
			Fn_SISW_NX_General_NavigatorTab_Operation=True
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully set preference Display Resource Bar on Left of NX window")
		End If
	End If

	If Fn_SISW_UI_Object_Operations("Fn_SISW_NX_General_NavigatorTab_Operation", "Exist", Window("NXWindow").WinObject("NavigatorTabTitle"), 5) =True Then
		If Window("NXWindow").WinObject("NavigatorTabTitle").GetROProperty("regexpwndtitle")=sTabName Then
			Fn_SISW_NX_General_NavigatorTab_Operation=True
			Set objTab=Nothing
			Set objTree=Nothing
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Tab "+sTabName+" is already activated")
			Exit Function
		End If
	End If

	 width=objTab.GetROProperty("width")
	 x=Cint(width/2)
	 If Window("NXWindow").Window("WinDrawingArea").Exist(5) Then
		 height=Window("NXWindow").Window("WinDrawingArea").GetROProperty("height")
	Else 
		height=Window("NXWindow").WinObject("NavigatorPane").GetROProperty("height")
	 End If
	 yfact=Cint(height/100)
	 bFlag=False
	 sFlag=False

	For iCounter=1 To 100
		y=iCounter*yfact
		Window("NXWindow").WinTab("NavigatorTab").MouseMove x+5,y

		If Window("nativeclass:=tooltips_class32").Exist(1) Then
			If sFlag=False Then
				sFirstTab=Window("nativeclass:=tooltips_class32").GetROProperty("text")
				sFlag=True
			End If
			bResult= Window("nativeclass:=tooltips_class32").GetROProperty("text")
			If bResult=sTabName Then
				Window("NXWindow").WinTab("NavigatorTab").Click  x+5,y
				If  Window("NXWindow").WinObject("NavigatorTabTitle").GetROProperty("regexpwndtitle")=sTabName Then
					Fn_SISW_NX_General_NavigatorTab_Operation=True
					bFlag=True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully selected "+sTabName+ "from navigator")
				Else
					Fn_SISW_NX_General_NavigatorTab_Operation=False
					Set objTab=Nothing
					Set objTree=Nothing
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to select "+sTabName+ "from navigator")
					Exit Function
				End If
				Exit For 
			End If
		End If
	Next

	If bFlag=False Then
		For iCounter=1 To 100
				y=iCounter*yfact
				Window("NXWindow").WinTab("NavigatorTab").MouseMove x,y
		
				If Window("nativeclass:=tooltips_class32").Exist(1) Then
					bResult= Window("nativeclass:=tooltips_class32").GetROProperty("text")
					If bResult=sTabName Then
						Window("NXWindow").WinTab("NavigatorTab").Click  x,y
						If  Window("NXWindow").WinObject("NavigatorTabTitle").GetROProperty("regexpwndtitle")=sTabName Then
							Fn_SISW_NX_General_NavigatorTab_Operation=True
							bFlag=False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully selected "+sTabName+ "from navigator")
						Else
							Fn_SISW_NX_General_NavigatorTab_Operation=False
							Set objTab=Nothing
							Set objTree=Nothing
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to select "+sTabName+ "from navigator")
							Exit Function
						End If
						Exit For 
					End If
					If bResult=sFirstTab Then
							Window("NXWindow").WinTab("NavigatorTab").Click Cint(x-1),Cint(height-15)
					End If
				End If
		Next
	End If
	Set objTab=Nothing
	Set objTree=Nothing

End Function

'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Function Name     			: Fn_SISW_NX_General_ANT_AssemblyVerify
' Function Description 			 : To verify Assembly in Assembly Navigator Tree
' Parameters			: 			sAction: Action Name
'											 sAssemblyStruct: Assembly Structure sepearated with :
' Return Value           :        True/False
' 	
' Examples		    	: 			 Fn_SISW_NX_General_ANT_AssemblyVerify("Assembly","000018:000017","","")
'									sNode = "000118 (Order:Chronological)~000121~000130~000153~000154~000156"
'									bReturn = Fn_SISW_NX_General_ANT_AssemblyVerify("Activate",sNode,"","3")
' History               :  
'		Developer Name			Date	  	Rev. No. 	Changes					Reviewer
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'		Nilesh Gadekar  	2-Sep-2013		1.0							
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'		Pranav Ingle		28-Dec-2013		1.1
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'		Vivek Ahirrao		19-Aug-2016		1.2			Added new cases "SelectAssemblyModelNode", "ActivateAssemblyModelNode","DeselectAssemblyModelNode" 
'														Added for Briefcase Browser New Dev TC's with NX		[TC1123-20160608-19_08_2016-VivekA-NewDevelopment]
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function  Fn_SISW_NX_General_ANT_AssemblyVerify(sAction,sAssemblyStruct,sColName,intCnt)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_NX_General_ANT_AssemblyVerify"
	Dim sData,aActAssembly,aExpAssembly,arrNode,sTreePath,sFilePath,aNode,aNode1,iRow, sASMNodeName
	Dim jCount,iStart,iCount,bResult,sCol,objTable,iColCount,iCol,bFlag,sRowContent,iRowCount
	
	Call Fn_SISW_NX_General_NavigatorTab_Operation("Assembly Navigator","")
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully select Assembly Navigator in NX")
	Call Fn_SISW_NX_Setup_ReadyStatusSync(1)
	
	Fn_SISW_NX_General_ANT_AssemblyVerify = False
	
	'---------------------------------------------   Header of Macro -Start   -------------------------------------------------------------------------------
	If sAction<>"Exist" AND sAction<>"Assembly" Then
		Set objFSO = CreateObject("Scripting.FileSystemObject")
		If lcase(Environment.Value("BBFlag")) = "true" Then
			sASMNodeName = sAssemblyStruct
	        iNXASMTreeStepCounter = iNXASMTreeStepCounter + 1
	        sDirPath = Environment.Value("BBReportFolderPath")
	        sASMNodeName = Replace(sASMNodeName,"(","_")
	        sASMNodeName = Replace(sASMNodeName,")","")
	        sASMNodeName = Replace(sASMNodeName,":","")
	        sASMNodeName = Replace(sASMNodeName," ","")
	        sFileName = CStr(iNXASMTreeStepCounter)+"-"+sAction+"-"+sASMNodeName+".macro"
		Else
	        sDirPath = Environment.Value("BatchFldName")
	        sFileName = sAction+"_"+Fn_RandNoGenerate+".macro"
		End If
		sFilePath = sDirPath&"\"&sFileName
		Set objFile = objFSO.CreateTextFile(sFilePath,True)
		objFile.WriteLine(Environment.Value("NXRelease"))
		objFile.WriteLine("Macro File: " &sFilePath)
		objFile.WriteLine("Macro Version "&Environment.Value("MacroVersion"))
		objFile.WriteLine("Macro List Language and Codeset: "&Environment.Value("MacroLanguage"))
		objFile.WriteLine("Created by "& "infodba" & " on "& Cstr(now))
		objFile.WriteLine("Part Name Display Style: $FILENAME")
		objFile.WriteLine("Selection Parameters 1 2 0.229167 1")
		objFile.WriteLine("Display Parameters 1.000000 9.385417 8.197917 -1.000000 -0.873474 1.000000 0.873474")
		objFile.WriteLine("*****************")
		objFile.WriteLine("RESET")
		objFile.WriteLine("CUSTOM HEADER 25 ""UGTL_macro"" 0")
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully create Macro File To execute "&sFileName)
	End If
	'--------------------------------------------------  Header of Macro -End  --------------------------------------------------------------------------

	Set objTable = Browser("AssemblyExportBrowser").Page("Page").WebTable("AssemblyTable")
	Call Fn_SISW_MakeIEDefaultBrowser()
	Call Fn_WindowsApplications("TerminateAll","iexplore.EXE")
	wait 1
	If sAction = "SelectAssemblyModelNode" OR sAction = "ActivateAssemblyModelNode" OR sAction = "DeselectAssemblyModelNode" Then
		If sColName="" Then
			sColName="Descriptive Part Name"
		End If
	Else
		If sAction = "Assembly" Then
			sAction = "Exist"
		End If
		If sColName="" Then
			sColName="Number"
		End If
	End If
    
	bResult = Fn_SISW_NX_Setup_LoadRunMacro("Set",Environment.Value("sPath")+"\TestData\NX\Macro\ANT_Export_IE.macro")
	If bResult=False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Run Macro to Export Assembly Navigator content into IE")
		Fn_SISW_NX_General_ANT_AssemblyVerify=False
		Exit Function
	End If
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Export Assembly in Browser")
	wait 5	
  
	Select Case sAction
		'Case for Briefcase Browser testcases
		Case "SelectAssemblyModelNode", "ActivateAssemblyModelNode","DeselectAssemblyModelNode"
				aNode = Split(sAssemblyStruct,"~") 
				aNode1 = Split(aNode(0)," ", -1, 1)
				sTreePath = aNode1(0)&".prt%"
				
				iColCount=objTable.GetROProperty("cols")
				bResult=False
				For iCount=1 to iColCount
					sCol=objTable.GetCellData(1,iCount)
					If Trim(sCol)=sColName Then
						iCol=iCount
						bResult=True
						Exit For
					End If
				Next 
				bSetFlag = False
				iRowCount=objTable.GetROProperty("rows")
				iSetFlagCount = 0
				For iCount = Ubound(aNode)  to 0  Step -1
					bResult = False
					iStart=1
                	For jCount = iRowCount  To intCnt Step -1
						sRowContent=RTrim(objTable.GetCellData(jCount,iCol))
						
						'Calculate total Space chars on left side in nodes in application
						'Only for 2nd Level nodes it is 4 and after that if we add child then 1 space is added to that
						'----------------------------
						'ANode
						'    BNode
						'     CNode
						'      DNode
						'       ENode
						'        FNode
						'    GNode
						'----------------------------
'						If bSetFlag = False Then
							iTotalSpaceCharInApp = Len(sRowContent)-Len(LTrim(sRowContent))
							iParamSpaceChar = (UBound(aNode)-1) + 4
							If iTotalSpaceCharInApp = iParamSpaceChar Then
								If Instr(Trim(sNode),Trim(sRowContent))>0 Then
									iSetFlagCount = iSetFlagCount
								Else
									iSetFlagCount = iSetFlagCount + 1
								End If
							ElseIf iTotalSpaceCharInApp < iParamSpaceChar Then
								'Used this if block to reset the value of iSetFlagCount to 0 if there is 2nd occurance of node on same level
								If iSetFlagCount>0 Then
									iSetFlagCount = 0
								End If
							End If
'						End If

						If Trim(sRowContent)= Trim(aNode(iCount)) Then
							'Verify Total hierarchy of node here i.e. Parent nodes
							'Verified using Spaces on Left side of node
							If UBound(aNode)>1 Then
								iLeftSpaceChars = Len(sRowContent) - Len(LTrim(sRowContent)) - 4
								If UBound(aNode)-1 = iLeftSpaceChars Then
									bFlag = True
								Else
									If iLeftSpaceChars=0 Then
										iStart=iStart+1
									End If
									bFlag = False
								End If
							Else
								bFlag = True
							End If
							
							If bFlag = True Then
'								bSetFlag = True
								If sAction = "SelectAssemblyModelNode" OR sAction = "DeselectAssemblyModelNode" Then
									If UBound(aNode)>1 Then
										sStart = ""
										For iBound = 3 To UBound(aNode) Step 1
											sStart = sStart + " 1%"
										Next
										sTreePath = sTreePath& " " & iStart & "%" & sStart & " " & iSetFlagCount & " * "
									Else
										sTreePath = sTreePath& " " & iStart & " * "
									End If
								ElseIf sAction = "ActivateAssemblyModelNode" Then
									If UBound(aNode)>1 Then
										sStart = ""
										For iBound = 3 To UBound(aNode) Step 1
											sStart = sStart + " 1%"
										Next
										sTreePath = sTreePath& " " & iStart & "%" & sStart & " " & iSetFlagCount & " * "
									Else
										sTreePath = sTreePath& " " & iStart & " * "
									End If
								End If
								bResult = True
								jCount = jCount - 1
	                        	Exit For
							End If
						Else
							'Checking Left Space charactore for 2nd Level Childs only
							If (Len(sRowContent)-Len(LTrim(sRowContent)))>4 Then
								'Do Nothing
							Else
								iStart=iStart+1
							End If
                        End If
                    Next
					intCnt = jCount+1
                Next
                
                If UBound(aNode) <> "0"  Then
                	If sAction = "SelectAssemblyModelNode" Then
						sTreePath = sTreePath & "0 * 1 " & "! "& aNode(UBound(aNode))
					ElseIf sAction = "ActivateAssemblyModelNode" Then
						sTreePath = sTreePath & "0 ! "& aNode(UBound(aNode))
					ElseIf sAction = "DeselectAssemblyModelNode" Then
						sTreePath = sTreePath & "0 * 0 " & "! "& aNode(UBound(aNode)) 
					End If
				Else
					If sAction = "SelectAssemblyModelNode" Then
						sTreePath = aNode1(0)&".prt * 0 * 1 ! "	& "<" & aNode1(LBound(aNode1)) & ">"
					ElseIf sAction = "ActivateAssemblyModelNode" Then
						sTreePath = aNode1(0)&".prt * 0 ! "	& "<" & aNode1(LBound(aNode1)) & ">"
					ElseIf sAction = "DeselectAssemblyModelNode" Then
						sTreePath = aNode1(0)&".prt * 0 * 0 ! "	& "<" & aNode1(LBound(aNode1)) & ">"
					End If
				End If
				
				If sAction = "SelectAssemblyModelNode" Then
					objFile.WriteLine("CUSTOM 25 ANT * TL_SELECT * RP"&sTreePath)
				ElseIf sAction = "ActivateAssemblyModelNode" Then
					objFile.WriteLine("CUSTOM 25 ANT * TL_DEFAULT_ACTION * RP"&sTreePath)
				ElseIf sAction = "DeselectAssemblyModelNode" Then
					objFile.WriteLine("CUSTOM 25 ANT * TL_SELECT * RP"&sTreePath)
				End If
				
				objFile.Close
				wait 1
				bResult= Fn_SISW_NX_Setup_LoadRunMacro("Set",sFilePath)
				If  bResult=False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to run Macro to Select Node["&sColName &"] from Assembly Tree in NX")
					Exit Function
				End If
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Select Node ["& sNode &"] from Assembly Tree in NX")
		Case "Select","Deselect","Activate","Exist","PopupMenuSelect"
				sAssemblyStruct = Replace(sAssemblyStruct," (Order: Chronological)"," (Order:Chronological)")
				sAssemblyStruct = Replace(sAssemblyStruct,"Order:Chronological","Order$Chronological")
				sAssemblyStruct = Replace(sAssemblyStruct,":","~")
				sAssemblyStruct = Replace(sAssemblyStruct,"Order$Chronological","Order:Chronological")
				aNode = Split(sAssemblyStruct,"~")
				aNodeSlashTop = Split(aNode(0),"/", -1, 1)
				aNodeTop = Split(aNodeSlashTop(0)," ", -1, 1)
				aNodeSlashBottom = Split(aNode(UBound(aNode)),"/", -1, 1)
				aNodeBottom = Split(aNodeSlashBottom(0)," ", -1, 1)
				'Get Column number which u want
				iColCount = objTable.GetROProperty("cols")
				For iCount=1 to iColCount
					sCol = objTable.GetCellData(1,iCount)
					If Trim(sCol)=sColName Then
						iCol = iCount
						Exit For
					End If
				Next
				'Get Total Row Count
				iRowCount = objTable.GetROProperty("rows")
				
				iStartFromRow = 3
				iEndOfRow = iRowCount
				iPreviousChild = iEndOfRow
				'sParentNode = ""
				'sParentNode1 = ""
				bSetFlag = False
				iTreePathMiddle = 0
				sTreePathMiddle = ""
				
				For iCount = 0 To UBound(aNode)
					bFlag = False
					iTreePathMiddle = 0
					Select Case CStr(iCount)
						Case "0"
							sRowContent = Trim(objTable.GetCellData(3,iCol))
							sFirstColContent = RTrim(objTable.GetCellData(3,1))
							sCompareNode = Replace(aNode(iCount)," (Order:Chronological)","")
							sRowContent = Replace(sRowContent," (Order:Chronological)","")
							sRowContent = Replace(sRowContent," (Order: Chronological)","")
							'Space chars before node
							iSpaceChars = (Len(sFirstColContent)-3)-Len(LTrim(sFirstColContent))
							If iSpaceChars < 0 Then iSpaceChars = 0
							
							If iSpaceChars = iCount Then
								If sRowContent = sCompareNode Then
									bFlag = True
									iStartFromRow = 1
									iParentSpaceChars = iSpaceChars
									'sParentNode = aNode(iCount)
									'sParentNode1 = aNode(iCount)
									If UBound(aNode)>iCount Then
										sTreePathStart = aNodeTop(0)&"/A%"
										If sAction = "Select" Then
											sTreePathEnd = " * 0 * 1 ! "&aNodeBottom(0)
										ElseIf sAction = "Deselect" Then
											sTreePathEnd = " * 0 * 0 ! "&aNodeBottom(0)
										ElseIf sAction = "Activate" Then
											sTreePathEnd = " * 0 ! "&aNodeBottom(0)
										ElseIf sAction = "PopupMenuSelect" Then
											sPopupMenu = intCnt
											sTreePathEnd = " * 0 * "&sPopupMenu&" ! "&aNodeBottom(0)
										End If
									Else
										sTreePathStart = aNodeTop(0)&"/A"
										If sAction = "Select" Then
											sTreePathEnd = " * 0 * 1 ! <"&aNodeTop(0)&">"
										ElseIf sAction = "Deselect" Then
											sTreePathEnd = " * 0 * 0 ! <"&aNodeTop(0)&">"
										ElseIf sAction = "Activate" Then
											sTreePathEnd = " * 0 ! <"&aNodeTop(0)&">"
										ElseIf sAction = "PopupMenuSelect" Then
											sPopupMenu = intCnt
											sTreePathEnd = " * 0 * "&sPopupMenu&" ! <"&aNodeTop(0)&">"
										End If
									End If
								End If
							End If
						Case "1"
							For iCount1 = iEndOfRow To iStartFromRow Step -1
								sRowContent = Trim(objTable.GetCellData(iCount1,iCol))
								sFirstColContent = RTrim(objTable.GetCellData(iCount1,1))
								sCompareNode = aNode(iCount)
								'Space chars before node
								iSpaceChars = (Len(sFirstColContent)-3)-Len(LTrim(sFirstColContent))
									If iSpaceChars < 0 Then iSpaceChars = 1
								If iSpaceChars = iCount Then
									If sRowContent = sCompareNode Then
										bFlag = True
										iTreePathMiddle = iTreePathMiddle+1
										iCurrentChild = iCount1+1
										'sParentNode1 = sParentNode1 + "~" + aNode(iCount)
'										If Instr(sAssemblyStruct,sParentNode1)>0 Then
'											sParentNode = sParentNode1
'										Else
'											sParentNode1 = sParentNode
'										End If
										If UBound(aNode)>iCount Then
											sTreePathMiddle = sTreePathMiddle+" "+CStr(iTreePathMiddle)+"%"
										Else
											sTreePathMiddle = sTreePathMiddle+" "+Cstr(iTreePathMiddle)
										End If
										Exit For
									Else
										iTreePathMiddle = iTreePathMiddle+1
										iPreviousChild = iCount1-1
									End If
								End If
							Next
							iStartFromRow = iCurrentChild
							iEndOfRow = iPreviousChild
						Case "2","3","4","5","6"
							iChildCount = 0
							For iCount1 = iStartFromRow To iEndOfRow Step -1
							'For iCount1 = iStartFromRow To iEndOfRow
								sFirstColContent = RTrim(objTable.GetCellData(iCount1,1))
								'Space chars before node
								iSpaceChars = (Len(sFirstColContent)-3)-Len(LTrim(sFirstColContent))
								
								If iSpaceChars = iCount Then
									iChildCount = iChildCount+1
								End If
							Next
							For iCount1 = iEndOfRow+1 To iStartFromRow+1
							'For iCount1 = iEndOfRow To iStartFromRow Step -1
								sRowContent = Trim(objTable.GetCellData(iCount1,iCol))
								sFirstColContent = RTrim(objTable.GetCellData(iCount1,1))
								sCompareNode = aNode(iCount)
								'Space chars before node
								iSpaceChars = (Len(sFirstColContent)-3)-Len(LTrim(sFirstColContent))
								
								If iSpaceChars < 0 Then iSpaceChars = iCount
								If iSpaceChars = iCount Then
									If sRowContent = sCompareNode Then
										bFlag = True
										iTreePathMiddle = iTreePathMiddle+1
										iTreePathMiddle = iChildCount-iTreePathMiddle+1
										iCurrentChild = iCount1+1
										'sParentNode1 = sParentNode1 + "~" + aNode(iCount)
'										If Instr(sAssemblyStruct,sParentNode1)>0 Then
'											sParentNode = sParentNode1
'										Else
'											sParentNode1 = sParentNode
'										End If
										If UBound(aNode)>iCount Then
											sTreePathMiddle = sTreePathMiddle+" "+CStr(iTreePathMiddle)+"%"
										Else
											sTreePathMiddle = sTreePathMiddle+" "+CStr(iTreePathMiddle)
										End If
										Exit For
									Else
										iTreePathMiddle = iTreePathMiddle+1
										iPreviousChild = iCount1-1
									End If
								End If
							Next
							iStartFromRow = iCurrentChild
							iEndOfRow = iPreviousChild
					End Select
					If bFlag = False Then
						Exit For
					End If
				Next
				
				If bFlag = False Then
					If Browser("AssemblyExportBrowser").Exist(5) = True Then
						Browser("AssemblyExportBrowser").Close()
						Call Fn_WindowsApplications("TerminateAll","iexplore.EXE")
					End If
				
					Fn_SISW_NX_General_ANT_AssemblyVerify = False
					Set objTable = Nothing
					Exit Function
				End If
				
				If sAction <> "Exist" Then
					'Create Macro last line
					sTreePath = sTreePathStart + sTreePathMiddle + sTreePathEnd
					If sAction = "Select" Then
						objFile.WriteLine("CUSTOM 25 ANT * TL_SELECT * RP@DB/"&sTreePath)
					ElseIf sAction = "Activate" Then
						objFile.WriteLine("CUSTOM 25 ANT * TL_DEFAULT_ACTION * RP@DB/"&sTreePath)
					ElseIf sAction = "Deselect" Then
						objFile.WriteLine("CUSTOM 25 ANT * TL_SELECT * RP@DB/"&sTreePath)
					ElseIf sAction = "PopupMenuSelect" Then
						Select Case sPopupMenu
							Case "Open:Assembly"
								sTreePath = Replace(sTreePath,sPopupMenu,"OPEN_ASSEMBLY")
							Case "Select Assembly"
								sTreePath = Replace(sTreePath,sPopupMenu,"SELECT_ASSEMBLY")
						End Select
						
						objFile.WriteLine("CUSTOM 25 ANT * POPUP * RP@DB/"&sTreePath)
					End If
					
					objFile.Close
					wait 1
					bResult = Fn_SISW_NX_Setup_LoadRunMacro("Set",sFilePath)
					If bResult = False Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to run Macro to "+sAction+" Node [ "+sNode+" ] from Assembly Tree in NX")
						Exit Function
					End If
				End If
				
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully "+sAction+" Node [ "+sNode+" ] from Assembly Tree in NX")
		Case "Quantity"
				aNode= Split(sAssemblyStruct,"~") 
			
			    iColCount=objTable.GetROProperty("cols")
				bResult=False
				For iCount=1 to iColCount
					sCol=objTable.GetCellData(1,iCount)
					If Trim(sCol)=sColName Then
						iCol=iCount
						bResult=True
						Exit For
					End If
				Next 
	
				iRowCount=objTable.GetROProperty("rows")
								
				For iCount = Ubound(aNode)  to 0  Step -1
                   	For jCount = 1 To iRowCount 
						sRowContent=RTrim(objTable.GetCellData(jCount,iCol))
						If Trim(sRowContent)= Trim(aNode(UBound(aNode)))  Then
							iRow = jCount
							Exit For
                       End If
					   	jCount = jCount + 1
                    Next
                Next

				For iCount=1 to iColCount
					sCol=objTable.GetCellData(1,iCount)
					If Trim(sCol)="Quantity" Then
                        iCol=iCount
						sRowContent=RTrim(objTable.GetCellData(iRow,iCol))
						If Trim(sRowContent) = intCnt Then
							bResult=True
							Exit For
					   End If
                    End If
				Next
	End Select
	
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Function Fn_SISW_NX_General_ANT_AssemblyVerify pass Successfully")
	If Browser("AssemblyExportBrowser").Exist(5) = True Then
		Browser("AssemblyExportBrowser").Close()
		Call Fn_WindowsApplications("TerminateAll","iexplore.EXE")
	End If

	Fn_SISW_NX_General_ANT_AssemblyVerify = True
	Set objTable = Nothing
End Function

'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Function Name     			: Fn_SISW_NX_General_OpenInNXOperation
' Function Description 			 : To open Assembly or Part in NX
' Parameters			: 			sAction: Action Name
'									sItemDetails: Assembly /Item Details
' Return Value           :        True/False
' 
' Examples		    	: 			Fn_SISW_NX_General_OpenInNXOperation("Item","000017/A","","")
' History               :  
'		Developer Name			 		Date	  				Rev. No. 				Changes					Reviewer
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'		Nilesh Gadekar  			5-Sep-2013					1.0							
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'		Pranav Ingle		  	   		28-Dec-2013				 1.1
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_NX_General_OpenInNXOperation(sAction,sItemDetails,sReserve1,sReserve2)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_NX_General_OpenInNXOperation"
   Dim iRan,sMasterFile,sNewFile,bResult,sJornalFile,sJornalFilePath
   Select Case sAction
	 	Case "Item"
			iRan=Fn_RandNoGenerate()
			sMasterFile=Environment.Value("sPath")&"\TestData\NX\Journal\Journal_OpenItem.txt"
			sNewFile=Environment.Value("BatchFldName")&"\Journal_OpenItem_"& iRan &".txt"
			sJornalFile="Journal_OpenItem_"& iRan &".vb"
			sJornalFilePath=Environment.Value("BatchFldName") & "\" & sJornalFile
			'Copy the Journal File at New LocationJournal_OpenItem
			bResult=Fn_Local_File_Operations("CopyFile",sMasterFile,sNewFile)
			If bResult=False Then
				Fn_SISW_NX_General_OpenInNXOperation=False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Copy Journal file ["& sMasterFile &"] to Location ["& sNewFile &"]")
				Exit Function
			Else
				Fn_SISW_NX_General_OpenInNXOperation=True
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Copied Journal file ["& sMasterFile &"] to Location ["& sNewFile &"]")
			End If
			'Replace Existing Item Details with current Item Details
			bResult=Fn_Local_File_Operations("FindAndReplace",sNewFile,"<ItemID>/<Item Rev>|"& sItemDetails)
			If bResult=False Then
				Fn_SISW_NX_General_OpenInNXOperation=False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Insert Item Detaiols in Journal File located at ["& sNewFile &"]")
				Exit Function
			Else
				Fn_SISW_NX_General_OpenInNXOperation=True
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Inserted Item Detaiols in Journal File located at ["& sNewFile &"]")
			End If

			''Rename txt file to .vb File
			Call Fn_Local_File_Operations("Rename",sNewFile,sJornalFile)
			wait 1
			'Play Journal File in NX Environment
			bResult=Fn_SISW_NX_General_JournalPlayBack(sJornalFilePath,"","","","")
			If bResult=False Then
				Fn_SISW_NX_General_OpenInNXOperation=False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to run Journal ["& sJornalFilePath &"] in NX Application")
				Exit Function
			Else
				Fn_SISW_NX_General_OpenInNXOperation=True
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully ran Journal ["& sJornalFilePath &"] in NX Application")
			End If
			Call Fn_SISW_NX_Setup_ReadyStatusSync(2)
			'Delete Temporary Journal File
			Call Fn_Local_File_Operations("DeleteFile",sJornalFilePath,"")
			wait 1
   End Select
   
End Function

'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Function Name     			: Fn_SISW_NX_General_TC_NavTeeOperation
' Function Description 			 : To Perform different Navtree operations in TC Navigator in NX
' Parameters			: 			sAction: Action Name
'									sNode: Node path 
' Return Value           :        True/False
' Pre-requisite		    : 			Nothing
' Function Call          :  			
' Examples		    	: 			Call Fn_SISW_NX_General_TC_NavTeeOperation ("Open","Clipboard:000111","","","")

' History               :  
'		Developer Name			 Date	  				Rev. No. 				Changes					Reviewer
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'		Pranav Ingle  			23-Sep-2013					1.0																Self
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'		Ashwini Patil			03-Mar-2014						1.1						Handeled Template dialog for "Double Click" Case
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_NX_General_TC_NavTeeOperation(sAction,sNode,sReserve1,sReserve2,sReserve3)
		GBL_FAILED_FUNCTION_NAME="Fn_SISW_NX_General_TC_NavTeeOperation"
		Dim objFSO,sDirPath,sFileName,sFilePath,objFile,bResult,sTreePath
		Dim sSpaceBefore, intCnt, sPath, objTable,objDialog
		Fn_SISW_NX_General_TC_NavTeeOperation=False

		Call Fn_SISW_NX_General_NavigatorTab_Operation("Teamcenter Navigator","")
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully select Teamcenter Navigator in NX")
		Call Fn_SISW_NX_Setup_ReadyStatusSync(1)
		
		'---------------------------------------------   Header of Macro -Start   -------------------------------------------------------------------------------
		Set objFSO = CreateObject("Scripting.FileSystemObject")	
		sDirPath = Environment.Value("BatchFldName")
		sFileName = sAction+"_"+Fn_RandNoGenerate+".macro"
		sFilePath =sDirPath&"\"&sFileName
		Set objFile=objFSO.CreateTextFile(sFilePath,True)
        objFile.WriteLine(Environment.Value("NXRelease"))
		objFile.WriteLine("Macro File: " &sFilePath)
		objFile.WriteLine("Macro Version "&Environment.Value("MacroVersion"))
		objFile.WriteLine("Macro List Language and Codeset: "&Environment.Value("MacroLanguage"))
		objFile.WriteLine("Created by "& "infodba" & " on "& Cstr(now))
		objFile.WriteLine("Part Name Display Style: $FILENAME")
		objFile.WriteLine("Selection Parameters 1 2 0.229167 1")
		objFile.WriteLine("Display Parameters 1.000000 9.437500 8.197917 -1.000000 -0.868653 1.000000 0.868653")
		objFile.WriteLine("*****************")
		objFile.WriteLine("RESET")
		objFile.WriteLine("CUSTOM HEADER 29 ""HD3D_Node"" 0")
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully create Macro File To execute "&sFileName)
		'--------------------------------------------------  Header of Macro -End  --------------------------------------------------------------------------

		'	----------------     Open  Nav Tree in Browser  and get Path ---------------------
        'Insert  First two Node
		If Instr(sNode,"Object:Teamcenter:")<=0 Then
			sNode="Object:Teamcenter:"&sNode
		End If
'		sSpaceBefore="  "
		sTreePath=Fn_SISW_NX_UI_GetNodePath(sNode, "Object",3,"0 0", sSpaceBefore, "NX_TCTree_Export_Browser.macro")
		If  sTreePath=False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to get Path for Node["& sNode &"] from Teamcenter Tree in NX")
			Exit Function
		End If
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Get Tree path for Node ["& sNode &"] from Teamcenter Tree in NX")
		'----------------------------------------------------------------------------------------

   Select Case sAction
	 	Case "Select"
				objFile.WriteLine("CUSTOM 29 TCNAV * TREE_TYPE * NODE_SELECT"&sTreePath)
				objFile.Close
				wait 1
				bResult= Fn_SISW_NX_Setup_LoadRunMacro("Set",sFilePath)
				If  bResult=False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to run Macro to Select Node["& sNode &"] from Teamcenter Tree in NX")
					Exit Function
				End If
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Select Node ["& sNode &"] from Teamcenter Tree in NX")

		Case "Expand"
				objFile.WriteLine("CUSTOM 29 TCNAV * TREE_TYPE * NODE_EXPAND"&sTreePath)
				objFile.Close
				wait 1
				bResult= Fn_SISW_NX_Setup_LoadRunMacro("Set",sFilePath)
				If  bResult=False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to run Macro to Expand Node["& sNode &"] from Teamcenter Tree in NX")
					Exit Function
				End If
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Expand Node ["& sNode &"] from Teamcenter Tree in NX")

		Case "Collapse"
				objFile.WriteLine("CUSTOM 29 TCNAV * TREE_TYPE * NODE_COLLAPSE"&sTreePath)
				objFile.Close
				wait 1
				bResult= Fn_SISW_NX_Setup_LoadRunMacro("Set",sFilePath)
				If  bResult=False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to run Macro to Collapse Node ["& sNode &"] from Teamcenter Tree in NX")
					Exit Function
				End If
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Collapse Node ["& sNode &"] from Teamcenter Tree in NX")

		Case "Exist"
				Fn_SISW_NX_General_TC_NavTeeOperation=True
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Verify Node ["& sNode &"] from Teamcenter Tree in NX")

		Case "DoubleClick","Open"
				objFile.WriteLine("CUSTOM 29 TCNAV * TREE_TYPE * NODE_MENUPOPUP"&sTreePath)
				objFile.WriteLine("DYN_POP_TOGGLE 2 1 ! Open")
				objFile.Close
				wait 1
				bResult= Fn_SISW_NX_Setup_LoadRunMacro("Set",sFilePath)
				If  bResult=False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to run Macro to Open/Double Click Node ["& sNode &"] from Teamcenter Tree in NX")
					Exit Function
				End If
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Open Node ["& sNode &"] from Teamcenter Tree in NX")
   End Select
   Wait 2
   Fn_SISW_NX_General_TC_NavTeeOperation=True
End Function

'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Function Name     			: Fn_SISW_NX_General_FitToResolution
' Function Description 			 : To Fit Drawing To Resolution
' Parameters			:           sReserve: For future use
' Return Value           :        True/False
' Pre-requisite		    : 			Nothing

' Examples		    	:            bReturn= Fn_SISW_NX_General_FitToResolution("")
'
' History               :  
		'Developer Name			 	Date	  					Rev. No. 				Changes					Reviewer
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'		Pranav Ingle	 			23-Sep-2013					1.0																Self
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_NX_General_FitToResolution(sReserve1)
		GBL_FAILED_FUNCTION_NAME="Fn_SISW_NX_General_FitToResolution"
		bResult=Fn_SISW_NX_General_JournalPlayBack(Environment.Value("sPath") &"\TestData\NX\Journal\Journal_FitDrawing.vb","","","","")
		If  bResult=False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Ran Journal Fit Drawing to Resolution in NX application")
			Fn_SISW_NX_General_FitToResolution=False
		Else
			Fn_SISW_NX_General_FitToResolution=True
		End If
		Call Fn_SISW_NX_Setup_ReadyStatusSync(1)
End Function

'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Function Name     	    :  Fn_SISW_NX_General_ExtraComponents()

' Function Description  :  	Function used to check existence of extra component addition box and click on it.

'Parameter						:   	sAction :- Action name
'													sMsg:- To be verify
' Return Value		    	:   	True/False
'
' Examples		    		:   Fn_SISW_NX_General_ExtraComponents("Click","The Structure of...")

'History					 :		

'	Developer Name					Date						Rev. No.			Changes								Reviewer
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Ashwini Patil						23-Jan-2014 				1.0																	Pranav Ingle
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_NX_General_ExtraComponents(sAction,sMsg)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_NX_General_ExtraComponents"
	 Dim objDialog,bReturn,sErrMsg

	Set objDialog = Dialog("ExtraComponents")
	Fn_SISW_NX_General_ExtraComponents = False

     If objDialog.Exist = False Then
			Fn_SISW_NX_General_ExtraComponents = False
			Call Fn_UpdateLogFiles("FAIL : Extra Componenet dialog does Not  Exist", "FAIL: Extra Componenet dialog does Not  Exist")
			Exit Function
	End If

	Select Case sAction
		Case "Click"
				If sMsg <> ""  Then
					sErrMsg=Dialog("ExtraComponents").Static("ErrMsg").GetROProperty("text")
					If Trim(sMsg) <> Trim(sErrMsg) Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to verify the message of Extra Component dialog box ")
						Set ObjDialog=Nothing
						Exit Function
					End If
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verified the message of Extra Component dialog box ")
				End If
				Wait 1
				bReturn = Fn_SISW_NX_UI_ButtonOperation("Fn_SISW_NX_General_ExtraComponents","Click",objDialog,"OK")
				If bReturn=False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Click on OK button of Extra Component Dialog")
					Set ObjDialog=Nothing
					Exit Function
				End If
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Successfully Clicked on OK button of Extra Component Dialog")
	End Select
	Set ObjDialog=Nothing
	Fn_SISW_NX_General_ExtraComponents = True
End Function
'
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Function Name     	    :  	Fn_SISW_NX_General_PartDialog()

' Function Description  :  	Function used to check existence of Part Dialog box and click on it.
'Parameter						:   	sAction :- Action name
'													sMsg:- To be verify

' Return Value		    	  :   True/False
'
' Examples		    			:   Fn_SISW_NX_General_PartDialog("ChangeDisplayedPart","You choose to...")
'                                               Fn_SISW_NX_General_PartDialog("Cancel","You choose to...")

'History					 :		

'	Developer Name						Date						Rev. No.			Changes							Reviewer
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Ashwini Patil							23-Jan-2014 				1.0																Pranav Ingle														
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_NX_General_PartDialog(sAction,sMsg)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_NX_General_PartDialog"
	 Dim objDialog,bReturn,sErrMsg

	Set objDialog = Dialog("OpenPart")
    
     If objDialog.Exist = False Then
			Call Fn_UpdateLogFiles("FAIL : Open Part dialog does Not  Exist", "FAIL: Open Part dialog does Not  Exist")
			Fn_SISW_NX_General_PartDialog = True
			Exit Function
	End If

	Fn_SISW_NX_General_PartDialog = False

	Select Case sAction
		Case "ChangeDisplayedPart"

				If sMsg <> ""  Then
					sErrMsg=Dialog("OpenPart").Static("ErrMsg").GetROProperty("text")
					If Trim(sMsg) <> Trim(sErrMsg) Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to verify the message of Open Part dialog box ")
						Set ObjDialog=Nothing
						Exit Function
					End If
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verified the message of Open Part dialog box ")
				End If

				bReturn = Fn_SISW_NX_UI_ButtonOperation("Fn_SISW_NX_General_PartDialog","Click",objDialog,"ChangeDisplayedPart")
				If bReturn=False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Click on ChangeDisplayedPart button of Part Dialog")
					Set ObjDialog=Nothing
					Exit Function
				End If
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Successfully Clicked on ChangeDisplayedPart button of Part Dialog")

		Case "Cancel"
				bReturn = Fn_SISW_NX_UI_ButtonOperation("Fn_SISW_NX_General_PartDialog","Click",objDialog,"Cancel")
				If bReturn=False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Click on Cancel button of Part Dialog")
					Set ObjDialog=Nothing
					Exit Function
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Successfully Clicked on Cancel button of Part Dialog")
				End If
	End Select
	Set ObjDialog=Nothing
	Fn_SISW_NX_General_PartDialog = True
End Function
'
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Function Name     	    	 :  Fn_SISW_NX_General_SaveAs()
' Function Description 	    :  	Function used to check existence of extra component addition box and click on it.
'Parameter						   :   	sAction :- Action name
'													Dim dicSaveAsInfo
'													Set dicSaveAsInfo = CreateObject("Scripting.Dictionary")
'													with dicSaveAsInfo
'	 													.Add "Action", sSaveAsAction
'													   .Add "LoadedPart", sLoadedPart
'													  .Add "ItemId", sPartID
'													 .Add "ItemRev", sPartRev
'													 .Add "ItemName", sPartName
'												End with 
'												sButton:- Button to be clicked	
'												sReserve2- For future use			
' Return Value		    	  :   True/False/Values
'
' Examples		    	      :   Fn_SISW_NX_General_SaveAs("WorkPart",dicSaveAsInfo,"","")

'History						  :		

'	Developer Name											Date						Rev. No.			Changes											Reviewer					Tc Release
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Ashwini Patil											23-Jan-2014 				1.0																	Pranav Ingle		
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Ankit Nigam												27-Oct-2015 				1.0				Modified function as per the design change for 		Vivek A.					Tc1121_2015100700
'																										ID , Revision and name input dialog 						
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_NX_General_SaveAs(sSaveAsScope,dicSaveAsInfo,sButton,sReserve2)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_NX_General_SaveAs"
	 Dim objDialog,bReturn,sItemId,sRev,objArr
	 Dim objLocation,L,T,R,B,X,Y

	Set objDialog = Window("NXWindow").Dialog("SavePartsAs")
	Set objSaveEdit = Window("NXWindow").Dialog("SavePartAsEdit")
	 Fn_SISW_NX_General_SaveAs = False

	If objDialog.Exist(2) = False Then
			Call Fn_UpdateLogFiles("FAIL : Save As dialog does Not  Exist", "FAIL: Save As dialog does Not  Exist")
			Exit Function
	End If

	If sSaveAsScope<>"" Then
		bResult=Fn_SISW_NX_UI_ComboBoxOperation( "Set",objDialog, "SaveAsScope","", sSaveAsScope,"")
		If  bResult=False Then
			  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select the Save As Scope Option of ComboBox ")
			  Exit Function
		End If
	End If

	If  dicSaveAsInfo("Action") = "Selected Parts" Then 
		If dicSaveAsInfo("LoadedPart") <> "" Then
			objLocation = Window("NXWindow").Dialog("SavePartsAs").WinObject("LoadedParts").GetTextLocation(dicSaveAsInfo("LoadedPart"), L, T, R, B)
			X = (L+R)/2
			Y = (T+B)/2	
			
			Window("NXWindow").Dialog("SavePartsAs").WinObject("LoadedParts").MouseMove X,Y
			Window("NXWindow").Dialog("SavePartsAs").WinObject("LoadedParts").Click X,Y
      		L=0
			R=0
			T=0
			B=0
			X=0
			Y=0       	
		Else
			Exit Function
		End If
    End If

	If dicSaveAsInfo("Action") <> "" Then
		bResult=Fn_SISW_NX_UI_ComboBoxOperation( "Set",objDialog, "SaveAs","", dicSaveAsInfo("Action"),"")
		If  bResult=False Then
			 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select the Save As Option of ComboBox ")
			 Exit Function
		End If
	End If

	If dicSaveAsInfo("ItemId") <> "" OR dicSaveAsInfo("ItemRev") <> "" Then						'Tc1121-2015100700-28_Oct_2015-AnkitN-Regression-Modified function as per the design change for ID , Revision and name input dialog 
		If dicSaveAsInfo("ItemId") <> "" Then
			call objDialog.WinObject("QWidget").GetTextLocation("ID",L, T, R, B,True)
			X = L + 10
			Y = B + 10
			objDialog.WinObject("QWidget").Click X ,Y ,micLeftBtn
			objDialog.WinObject("QWidget").Click X ,Y ,micRightBtn
      		L=0
			R=0
			T=0
			B=0
			X=0
			Y=0
			call objDialog.Window("PopUpMenu").GetTextLocation("Edit...",L, T, R, B,True)
			X = L + 10
			Y = T + 10
			objDialog.Window("PopUpMenu").Click X, Y ,micLeftBtn
			wait 2			
			objSaveEdit.WinEdit("SavaAsID").SetTOProperty "attached text" ,"ID"
			objSaveEdit.WinEdit("SavaAsID").Type dicSaveAsInfo("ItemId")
			objSaveEdit.WinButton("OK").Click
			If objSaveEdit.Exist(2) Then
				objSaveEdit.Close()	
			End If
			L=0
			R=0
			T=0
			B=0
			X=0
			Y=0		
		End If
		If dicSaveAsInfo("ItemRev") <> "" Then
			call objDialog.WinObject("QWidget").GetTextLocation("Revision",L, T, R, B,True)
			X = L + 10
			Y = B + 10
			objDialog.WinObject("QWidget").Click X ,Y ,micLeftBtn
			objDialog.WinObject("QWidget").Click X ,Y ,micRightBtn
      		L=0
			R=0
			T=0
			B=0
			X=0
			Y=0
			call objDialog.Window("PopUpMenu").GetTextLocation("Edit...",L, T, R, B,True)
			X = L + 10
			Y = T + 10
			objDialog.Window("PopUpMenu").Click X, Y ,micLeftBtn
			wait 2			
			objSaveEdit.WinEdit("SavaAsID").SetTOProperty "attached text" ,"Revision"
			objSaveEdit.WinEdit("SavaAsID").Type dicSaveAsInfo("ItemRev")
			objSaveEdit.WinButton("OK").Click
			If objSaveEdit.Exist(2) Then
				objSaveEdit.Close()	
			End If
			L=0
			R=0
			T=0
			B=0
			X=0
			Y=0		
		End If
	Else 	
		call objDialog.WinObject("QWidget").GetTextLocation("ID",L, T, R, B,True)
		X = L + 10
		Y = B + 10
		objDialog.WinObject("QWidget").Click X, Y, micLeftBtn
		objDialog.WinObject("QWidget").DblClick X, Y, micLeftBtn	
			L=0
			R=0
			T=0
			B=0
			X=0
			Y=0		
	
    End If

    If dicSaveAsInfo("ItemName") <> "" Then
			call objDialog.WinObject("QWidget").GetTextLocation("Revision",L, T, R, B,True)
			X = R + 10
			Y = B + 10
			objDialog.WinObject("QWidget").Click X + 100,Y ,micLeftBtn
			objDialog.WinObject("QWidget").Click X + 100,Y ,micRightBtn
      		L=0
			R=0
			T=0
			B=0
			X=0
			Y=0
			call objDialog.Window("PopUpMenu").GetTextLocation("Edit...",L, T, R, B,True)
			X = L + 10
			Y = T + 10
			objDialog.Window("PopUpMenu").Click X, Y ,micLeftBtn
			wait 2			
			objSaveEdit.WinEdit("SavaAsID").SetTOProperty "attached text" ,"Name"
			objSaveEdit.WinEdit("SavaAsID").Type dicSaveAsInfo("ItemName")
			objSaveEdit.WinButton("OK").Click
			If objSaveEdit.Exist(2) Then
				objSaveEdit.Close()	
			End If
			L=0
			R=0
			T=0
			B=0
			X=0
			Y=0

			call objDialog.WinObject("QWidget").GetTextLocation("ID",L, T, R, B,True)
			X = L + 10
			Y = B + 10
			objDialog.WinObject("QWidget").Click X ,Y ,micLeftBtn
			objDialog.WinObject("QWidget").Click X ,Y ,micRightBtn
      		L=0
			R=0
			T=0
			B=0
			X=0
			Y=0
			call objDialog.Window("PopUpMenu").GetTextLocation("Edit...",L, T, R, B,True)
			X = L + 10
			Y = T + 10
			objDialog.Window("PopUpMenu").Click X, Y ,micLeftBtn
			wait 2	
			objSaveEdit.WinEdit("SavaAsID").SetTOProperty "attached text" ,"ID"			
			sItemID = objSaveEdit.WinEdit("SavaAsID").GetROProperty("text")
			objSaveEdit.WinButton("OK").Click
			If objSaveEdit.Exist(2) Then
				objSaveEdit.Close()	
			End If
			L=0
			R=0
			T=0
			B=0
			X=0
			Y=0	

			call objDialog.WinObject("QWidget").GetTextLocation("Revision",L, T, R, B,True)
			X = L + 10
			Y = B + 10
			objDialog.WinObject("QWidget").Click X ,Y ,micLeftBtn
			objDialog.WinObject("QWidget").Click X ,Y ,micRightBtn
      		L=0
			R=0
			T=0
			B=0
			X=0
			Y=0
			call objDialog.Window("PopUpMenu").GetTextLocation("Edit...",L, T, R, B,True)
			X = L + 10
			Y = T + 10
			objDialog.Window("PopUpMenu").Click X, Y ,micLeftBtn
			wait 2	
			objSaveEdit.WinEdit("SavaAsID").SetTOProperty "attached text" ,"Revision"			
			sRev = objSaveEdit.WinEdit("SavaAsID").GetROProperty("text")
			objSaveEdit.WinButton("OK").Click
			If objSaveEdit.Exist(2) Then
				objSaveEdit.Close()	
			End If
			L=0
			R=0
			T=0
			B=0
			X=0
			Y=0	
	End If
	bResult=Fn_SISW_NX_UI_ComboBoxOperation( "Set",objDialog, "DependentFilesSaveAs","", "Save All","")
	If bResult=False Then
			Fn_SISW_NX_Setup_CmdFinderOperation=False
			Set ObjCmdDialog=Nothing
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to set value Save All in combobox")
			Exit Function
	End If	

	bReturn = Fn_SISW_NX_UI_ButtonOperation("Fn_SISW_NX_General_SaveAs","Click",objDialog,sButton)
	If bReturn=False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Click on ["+sButton+"] button of Save As Dialog")
		Set ObjDialog=Nothing
		Exit Function
	Else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Successfully Clicked on  ["+sButton+"] button of Save As Dialog")
	End If	
	Fn_SISW_NX_General_SaveAs = True
	Fn_SISW_NX_General_SaveAs = sItemId & "-" & sRev
	Set ObjDialog=Nothing		
End Function

'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Function Name     	    	 :  Fn_SISW_NX_General_CustomerDefault()
' Function Description 	    :  	Function used to set the User defined preference through customer default dialog box.
'Parameter						   :   	Dim dicCustomerDefaultInfo
'													Set dicCustomerDefaultInfo = CreateObject("Scripting.Dictionary")
'													with  dicCustomerDefaultInfo
'	 													.Add "Application", sApplication
'													   .Add "Category", sCategory
'													  .Add "Tab", sTab
'													 .Add "PreferenceName", sPreferenceName
'													 .Add "Button", sButton
'												End with 	
'												sButton1 - OK button for Find Dilaog						
'												sReserve1- For future use       								
' Return Value		    	  :   True/False/Values
'
' Examples		    	      :   Fn_SISW_NX_General_CustomerDefault(dicCustomerDefaultInfo,"OK","")
'History						  :		

'	Developer Name											Date						Rev. No.			Changes							Reviewer
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Ashwini Patil											12-Feb-2014 				1.0															Pranav Ingle		
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_NX_General_CustomerDefault(dicCustomerDefaultInfo,sButton1,sReserve1)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_NX_General_CustomerDefault"
	 Dim objDialog,objDialog1,bReturn,sItemId,sRev,objArr
	 Dim objLocation,L,T,R,B,X,Y

	Set objDialog = Window("NXWindow").Dialog("CustomerDefaults")
	 Fn_SISW_NX_General_CustomerDefault = False

	If objDialog.Exist(2) = False Then
			Call Fn_UpdateLogFiles("FAIL : Customer Default dialog does Not  Exist", "FAIL:  Customer Default dialog does Not  Exist")
			Exit Function
	End If

	If dicCustomerDefaultInfo("Button")<>"" Then
		bReturn = Fn_SISW_NX_UI_ButtonOperation("Fn_SISW_NX_General_CustomerDefault","Click",objDialog,"Find")
		If bReturn=False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Click on Find button of Customer Default Dialog")
			Set ObjDialog=Nothing
			Exit Function
		End If
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Successfully Clicked on Find button of Customer Default Dialog")
  	End If

	Set objDialog1 = Window("NXWindow").Dialog("FindDefault")
	Fn_SISW_NX_General_CustomerDefault = False

	If  dicCustomerDefaultInfo("PreferenceName") <> "" Then 
		bResult=Fn_SISW_NX_UI_EditBoxOperation("Fn_SISW_NX_General_CustomerDefault","Set",objDialog1,"SearchDefault", dicCustomerDefaultInfo("PreferenceName"))
		If bResult=False Then
            Set objDialog1=Nothing
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to set value ["+dicCustomerDefaultInfo("PreferenceName")+"] inSearch Default Edit box ")
			Exit Function
		End If
    End If

	If dicCustomerDefaultInfo("Button") <> "" Then
		bReturn = Fn_SISW_NX_UI_ButtonOperation("Fn_SISW_NX_General_CustomerDefault","Click",objDialog1,"Find")
		If bReturn=False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Click on Find button of Find Default Dialog")
			Set ObjDialog1=Nothing
			Exit Function
		End If
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Successfully Clicked on Find button of Find Default Dialog")
	End If

	If dicCustomerDefaultInfo("Application") <> "" Then
			objLocation = Window("NXWindow").Dialog("FindDefault").WinObject("DefaultsFound").GetTextLocation(dicCustomerDefaultInfo("Application"), L, T, R, B)

			X = (L+R)/2
			Y = (T+B)/2	
			
			'Window("NXWindow").Dialog("FindDefault").WinObject("DefaultsFound").MouseMove X,Y
			'wait 5
			Window("NXWindow").Dialog("FindDefault").WinObject("DefaultsFound").Click X,Y
                   	
	Else
			Exit Function
	End If

	If sButton1 <> "" Then
		bReturn = Fn_SISW_NX_UI_ButtonOperation("Fn_SISW_NX_General_CustomerDefault","Click",objDialog1,"OK")
		If bReturn=False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Click on OK button of Find Default Dialog")
			Set ObjDialog1=Nothing
			Exit Function
		End If
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Successfully Clicked on OK button of Find Default Dialog")
	End If		
	'==========================================================================================
	'Added Field to Set ON/OFF Display Messsage When Modifying Read only parts in General tab of Assemblies
	'TC11.5_2018022600_DOPRO_NX_NewDevelopment_PoonamC_06Apr2018
	Wait 2
	If dicCustomerDefaultInfo("DisplayMsgWhenReadOnlyParts") <> "" Then
		Select Case dicCustomerDefaultInfo("DisplayMsgWhenReadOnlyParts")
			Case "ON"
				Window("NXWindow").Dialog("CustomerDefaults").WinCheckBox("DisplayMessagewhenModifying").Set "ON"	
			Case "OFF"
				Window("NXWindow").Dialog("CustomerDefaults").WinCheckBox("DisplayMessagewhenModifying").Set "OFF"	
		End Select		
	End If
	
	If dicCustomerDefaultInfo("AllowSaveAsDifferentItemType") <> "" Then
		Select Case dicCustomerDefaultInfo("AllowSaveAsDifferentItemType")
			Case "ON"
				Window("NXWindow").Dialog("CustomerDefaults").WinCheckBox("AllowSaveAsDifferentItemType").Set "ON"	
			Case "OFF"
				Window("NXWindow").Dialog("CustomerDefaults").WinCheckBox("AllowSaveAsDifferentItemType").Set "OFF"	
		End Select		
	End If
	'==========================================================================================

    If Window("NXWindow").Dialog("CustomerDefaults").WinButton("OK").GetROProperty("enabled") <> False Then
		bReturn = Fn_SISW_NX_UI_ButtonOperation("Fn_SISW_NX_General_CustomerDefault","Click",objDialog,"OK")
		If bReturn=False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Click on OK button of Customer Default Dialog")
			Set ObjDialog=Nothing
			Exit Function
		End If
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Successfully Clicked on OK button of Customer Default Dialog")
	Else
		bReturn = Fn_SISW_NX_UI_ButtonOperation("Fn_SISW_NX_General_CustomerDefault","Click",objDialog,"Cancel")
		If bReturn=False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Click on Cancel button of Customer Default Dialog")
			Set ObjDialog=Nothing
			Exit Function
		End If
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Successfully Clicked on Cancel button of Customer Default Dialog")
	End If
	
	If objDialog.Exist(2) = True Then
		bReturn = Fn_SISW_NX_UI_ButtonOperation("Fn_SISW_NX_General_CustomerDefault","Click",objDialog,"OK_2")
		If bReturn=False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Click on OK button of Customer Default Dialog")
			Set ObjDialog1=Nothing
			Exit Function
		End If
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Successfully Clicked on OK button of Customer Default Dialog")
	End If	
		
    Fn_SISW_NX_General_CustomerDefault = True
    Set ObjDialog=Nothing	
	Set ObjDialog1=Nothing		
End Function

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name	:	Fn_SISW_NX_General_AssemblyCreate
'@@
'@@    Description		:	Function Used to Create Assembly in NX for Briefcase Browser testcases
'@@
'@@    Parameters		:	1. sAssemblyName	: Name of assembly node
'@@						:	2. sAssemblyPath	: Path for node to create
'@@						:	3. sReserve1 		: Future use
'@@
'@@    Return Value		: 	True Or False
'@@
'@@    Examples			:   Fn_SISW_NX_General_AssemblyCreate("Asm1_A","C:/Temp/NX","","")
'@@    							
'@@	   History			:	
'@@			Developer Name		Date	  		Rev. No.	Changes Done										Reviewer
'@@-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@			Vivek Ahirrao		19-Aug-2016		1.0		  	Created												[TC1123-20160608-19_08_2016-VivekA-NewDevelopment]
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_SISW_NX_General_AssemblyCreate(sAssemblyName,sAssemblyPath,sReserve1,sReserve2)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_NX_General_AssemblyCreate"
'	Update the Item Information Journal folder of TestData
	Dim bResult,sItemId,sItemRev
	Fn_SISW_NX_General_AssemblyCreate=False

	If Instr(sAssemblyName,"SubASM@@")>0 Then
		aAssemblyName = Split(sAssemblyName,"@@")
		sXMLFileName = "Journal_SubAssembly_Create.vb"
		sAssemblyName = aAssemblyName(1)
	Else
		sXMLFileName = "Journal_Assembley_Create.vb"
		sAssemblyName = sAssemblyName
	End If
	
	bResult= Fn_UpdateEnvXMLNode(Environment.Value("sPath")&"\TestData\NX\Journal\BriefcaseBrowser\AssemblyCreate_Input.xml", "AssemblyName",sAssemblyName)
	If bResult=False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Update Item Name Information in ItemCreate_Input.xml")
		Exit Function
	End If

	bResult= Fn_UpdateEnvXMLNode(Environment.Value("sPath")&"\TestData\NX\Journal\BriefcaseBrowser\AssemblyCreate_Input.xml", "AssemblyPath",sAssemblyPath)
	If bResult=False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to UpdateFolder Path in ItemCreate_Input.xml")
		Exit Function
	End If
	
    bResult=Fn_SISW_NX_General_JournalPlayBack(Environment.Value("sPath") &"\TestData\NX\Journal\BriefcaseBrowser\"&sXMLFileName,"","","","")
	If  bResult=False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Ran Journal File to create Item in NX application")
		Exit Function
	End If
	Call Fn_SISW_NX_Setup_ReadyStatusSync(2)

	Fn_SISW_NX_General_AssemblyCreate=True
End Function

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name	:	Fn_SISW_NX_General_ModelCreate
'@@
'@@    Description		:	Function Used to Create Model in NX for Briefcase Browser testcases
'@@
'@@    Parameters		:	1. sModelName	: Name of Model node
'@@						:	2. sModelPath	: Path for node to create
'@@						:	3. sReserve1 	: Future use
'@@
'@@    Return Value		: 	True Or False
'@@
'@@    Examples			:   Fn_SISW_NX_General_ModelCreate("Mod1_A","C:/Temp/NX","","")
'@@    							
'@@	   History			:	
'@@			Developer Name		Date	  		Rev. No.	Changes Done										Reviewer
'@@-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@			Vivek Ahirrao		19-Aug-2016		1.0		  	Created												[TC1123-20160608-19_08_2016-VivekA-NewDevelopment]
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_SISW_NX_General_ModelCreate(sModelName,sModelPath,sReserve1,sReserve2)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_NX_General_ModelCreate"
'	Update the Item Information Journal folder of TestData
	Dim bResult,sItemId,sItemRev
	Fn_SISW_NX_General_ModelCreate=False

	bResult= Fn_UpdateEnvXMLNode(Environment.Value("sPath")&"\TestData\NX\Journal\BriefcaseBrowser\ModelCreate_Input.xml", "ModelName",sModelName)
	If bResult=False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Update Item Name Information in ItemCreate_Input.xml")
		Exit Function
	End If

	bResult= Fn_UpdateEnvXMLNode(Environment.Value("sPath")&"\TestData\NX\Journal\BriefcaseBrowser\ModelCreate_Input.xml", "ModelPath",sModelPath)
	If bResult=False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to UpdateFolder Path in ItemCreate_Input.xml")
		Exit Function
	End If

    bResult=Fn_SISW_NX_General_JournalPlayBack(Environment.Value("sPath") &"\TestData\NX\Journal\BriefcaseBrowser\Journal_Model_Create.vb","","","","")
	If  bResult=False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Ran Journal File to create Item in NX application")
		Exit Function
	End If
	Call Fn_SISW_NX_Setup_ReadyStatusSync(2)

	Fn_SISW_NX_General_ModelCreate=True
End Function

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name	:	Fn_NX_ApplicationLaunch
'@@
'@@    Description		:	Function Used to Launch NX application
'@@
'@@    Parameters		:	1. sAction		: Action to launch NX application
'@@						:	2. sReserve1 	: Future use
'@@
'@@    Return Value		: 	True Or False
'@@
'@@    Examples			:   Fn_NX_ApplicationLaunch("ugrafexe","","","","")
'@@    							
'@@	   History			:	
'@@			Developer Name		Date	  		Rev. No.	Changes Done										Reviewer
'@@-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@			Vivek Ahirrao		19-Aug-2016		1.0		  	Created												[TC1123-20160608-19_08_2016-VivekA-NewDevelopment]
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_NX_ApplicationLaunch(sAction,sReserve1,sReserve2,sReserve3,sReserve4)
	GBL_FAILED_FUNCTION_NAME="Fn_NX_ApplicationLaunch"
	Dim objNXWindow, objWSH, objSystemVariables
	
	Fn_NX_ApplicationLaunch = False
	
	Set objNXWindow = Window("NXWindow")
	If objNXWindow.Exist(2) Then
		Call Fn_NX_ApplicationExit("KillProcess","ugraf.exe","")
		Wait 2
	End If
	
	Select Case sAction
		Case "ugrafexe"
				Set objWSH =  CreateObject("WScript.Shell")
				Set objSystemVariables = objWSH.Environment("System")
				sNXSetupPath = (objSystemVariables("UGII_ROOT_DIR"))
				'run ugraf.exe
				SystemUtil.Run sNXSetupPath&"/ugraf.exe"
				Wait 10
				If objNXWindow.Exist(SISW_MAX_TIMEOUT) Then						  							
					Fn_NX_ApplicationLaunch = TRUE  																						
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Invoked NX Application from ["+sNXSetupPath+"]")
				Else
					Fn_NX_ApplicationLaunch = FALSE
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Invoke NX Application from ["+sNXSetupPath+"]")
					Exit Function
				End If
				
		Case "InvokeFromTeamcenter"
				'Future USe
	End Select
	Fn_NX_ApplicationLaunch = true
End Function

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name	:	Fn_NX_ApplicationExit
'@@
'@@    Description		:	Function Used to Exit from NX application
'@@
'@@    Parameters		:	1. sAction		: Action to Exit from NX application
'@@						:	2. sNXProcess 	: Future use
'@@
'@@    Return Value		: 	True Or False
'@@
'@@    Examples			:   Fn_NX_ApplicationLaunch("KillProcess","ugraf.exe","","","")
'@@    							
'@@	   History			:	
'@@			Developer Name		Date	  		Rev. No.	Changes Done										Reviewer
'@@-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@			Vivek Ahirrao		19-Aug-2016		1.0		  	Created												[TC1123-20160608-19_08_2016-VivekA-NewDevelopment]
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_NX_ApplicationExit(sAction,sNXProcess,sReserve1)
	GBL_FAILED_FUNCTION_NAME="Fn_NX_ApplicationExit"
	Fn_NX_ApplicationExit = False
	
	Select Case sAction
		Case "KillProcess"
				sArrData = split(sNXProcess, ":",-1,1)
				strComputer = "." 
				Set objWMIService = GetObject("winmgmts:"& "{impersonationLevel=impersonate}!\\"& strComputer & "\root\cimv2") 
				For iCount = 0 to ubound(sArrData)
					Set colProcess = objWMIService.ExecQuery("Select * from Win32_Process Where Name ='"+sArrData(iCount)+"'")  
					'For Each objProcess in colProcess 
					For Each objProcess in colProcess 
						objProcess.Terminate() 
					Next 
				Next
				Fn_NX_ApplicationExit = True
				
		Case "ExitApplication"
				'Future Use
	End Select
End Function

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name	:	Fn_SISW_NX_SaveOptions
'@@
'@@    Description		:	Function Used to perform operations "Save Options" dialog to seva assembly in proper format.
'@@
'@@    Parameters		:	1. sAction		: Action to Exit from NX application
'@@						:	2. sNXProcess 	: Future use
'@@
'@@    Return Value		: 	True Or False
'@@
'@@    Examples			:   Fn_SISW_NX_SaveOptions()
'@@    							
'@@	   History			:	
'@@			Developer Name		Date	  		Rev. No.	Changes Done										Reviewer
'@@-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@			Vivek Ahirrao		19-Aug-2016		1.0		  	Created												[TC1123-20160608-19_08_2016-VivekA-NewDevelopment]
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_SISW_NX_SaveOptions()
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_NX_SaveOptions"
	Dim objNXWindow
	Set objNXSaveWindow = Window("NXWindow").Dialog("SaveOptions")
	
	bReturn = Fn_SISW_NX_Setup_CmdFinderOperation("Save Options","Start","")
	If bReturn=True Then
		Fn_SISW_NX_SaveOptions = True
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully performed Menu operation for [ Save Options ]")
	Else
		Fn_SISW_NX_SaveOptions = False
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to perform Menu operation for [ Save Options ]")
	End If
	
	If objNXSaveWindow.Exist(5) then
		Err.Clear
		objNXSaveWindow.WinCheckBox("SaveJTData").Set "ON"
		If Err.Number < 0 Then
			Fn_SISW_NX_SaveOptions = False
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to click on check box [ Save JT Data ]")
		Else
			Fn_SISW_NX_SaveOptions = True
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully clicked on check box [ Save JT Data ]")
		End If
		Wait 1
		
		Err.Clear
		objNXSaveWindow.WinButton("Apply").Click
		If Err.Number < 0 Then
			Fn_SISW_NX_SaveOptions = False
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to click on button [ Apply ]")
		Else
			Fn_SISW_NX_SaveOptions = True
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully clicked on button [ Apply ]")
		End If
		Wait 1
		
		Err.Clear
		objNXSaveWindow.WinButton("OK").Click
		If Err.Number < 0 Then
			Fn_SISW_NX_SaveOptions = False
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to click on button [ OK ]")
		Else
			Fn_SISW_NX_SaveOptions = True
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully clicked on button [ OK ]")
		End If
		Wait 1											
	Else
		Fn_SISW_NX_SaveOptions = False
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to invoke [ Save Options ] window")
	End If
	Fn_SISW_NX_SaveOptions = True
End Function

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name	:	Fn_SISW_NX_General_AddExistingComponent
'@@
'@@    Description		:	Function Used to Add Existing Component under Top Node in NX Application
'@@
'@@    Parameters		:	1. sAction			: Action Name
'@@						:	2. sComponentPath	: Path for Existing component
'@@						:	3. sComponentName 	: Existing component name
'@@						:	3. s3DPoints 		: x,y,z points    		[0.0,0.0,0.0]
'@@						:	3. sOrientations 	: Orientations points	[1.0,0.0,0.0,0.0,1.0,0.0,0.0,0.0,1.0]	
'@@						:	3. sReserve1 		: Future use
'@@
'@@    Return Value		: 	True Or False
'@@
'@@    Examples			:   sComponentPath = "C:\mainline\Reports\BEBriefcaseEditorApplication_OpenBriefcaseViewStructure\NX\Top16845_A65185_A.prt"
'@@    						sComponentName = "TOP16845_A65185_A"
'@@    						bReturn = Fn_SISW_NX_General_AddExistingComponent("Add",sComponentPath,sComponentName,"","","")
'@@    							
'@@	   History			:	
'@@			Developer Name		Date	  		Rev. No.	Changes Done										Reviewer
'@@-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@			Vivek Ahirrao		28-Sep-2016		1.0		  	Created												[BB1123-20160608-28_09_2016-VivekA-NewDevelopment]
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_SISW_NX_General_AddExistingComponent(sAction,sComponentPath,sComponentName,s3DPoints,sOrientations,sReserve1)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_NX_General_AddExistingComponent"
	
	Fn_SISW_NX_General_AddExistingComponent = False
	
	Select Case sAction
		Case "Add"
			'Update ComponentPath in XML
			bResult = Fn_UpdateEnvXMLNode(Environment.Value("sPath")&"\TestData\NX\Journal\BriefcaseBrowser\AddExistingComponent.xml", "ComponentPath",sComponentPath)
			If bResult = False Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Update ComponentPath Information in AddExistingComponent.xml")
				Exit Function
			End If
			'Update ComponentPath in XML
			bResult = Fn_UpdateEnvXMLNode(Environment.Value("sPath")&"\TestData\NX\Journal\BriefcaseBrowser\AddExistingComponent.xml", "ComponentName",sComponentName)
			If bResult = False Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Update ComponentName Information in AddExistingComponent.xml")
				Exit Function
			End If
			
			'Update ComponentPath in XML
			If s3DPoints<>"" Then
				bResult = Fn_UpdateEnvXMLNode(Environment.Value("sPath")&"\TestData\NX\Journal\BriefcaseBrowser\AddExistingComponent.xml", "3DPoints",s3DPoints)
				If bResult = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Update 3DPoints Information in AddExistingComponent.xml")
					Exit Function
				End If
			End If
			
			'Update ComponentPath in XML
			If sOrientations<>"" Then
				bResult = Fn_UpdateEnvXMLNode(Environment.Value("sPath")&"\TestData\NX\Journal\BriefcaseBrowser\AddExistingComponent.xml", "Orientations",sOrientations)
				If bResult = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Update Orientations Information in AddExistingComponent.xml")
					Exit Function
				End If
			End If
			
			bResult = Fn_SISW_NX_General_JournalPlayBack(Environment.Value("sPath") &"\TestData\NX\Journal\BriefcaseBrowser\Journal_AddExistingComponent.vb","","","","")
			If bResult = False Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Run Journal File [Journal_AddExistingComponent.vb] to Add Existing Component in NX application")
				Exit Function
			End If
			Call Fn_SISW_NX_Setup_ReadyStatusSync(5)
			Fn_SISW_NX_General_AddExistingComponent = True
		Case Else
			'Do Nothing
			Fn_SISW_NX_General_AddExistingComponent = False
	End Select

End Function

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name	:	Fn_SISW_NX_General_ComponentProperties_Ops
'@@
'@@    Description		:	Function Used to Perform operations on Component Properties dialog
'@@
'@@    Parameters		:	1. sAction				: Action Name
'@@						:	2. dicCompPropDetails	: Dictionary object
'@@						:	3. sButton	 			: OK / Apply / Cancel
'@@						:	4. sReserve 			: Future Use
'@@
'@@    Return Value		: 	True Or False
'@@
'@@    Examples			:   sComponentPath = "C:\mainline\Reports\BEBriefcaseEditorApplication_OpenBriefcaseViewStructure\NX\Top16845_A65185_A.prt"
'@@    						sComponentName = "TOP16845_A65185_A"
'@@    						bReturn = Fn_SISW_NX_General_AddExistingComponent("Add",sComponentPath,sComponentName,"","","")
'@@    							
'@@	   History			:	
'@@			Developer Name		Date	  		Rev. No.	Changes Done					Reviewer
'@@-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@			Vivek Ahirrao		07-Oct-2016		1.0		  	Created							[TC1123-20160915-07_10_2016-VivekA-NewDevelopment]
'@@															Added for PSM-NX integration testcases
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_SISW_NX_General_ComponentProperties_Ops(sAction,dicCompPropDetails,sButton,sReserve)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_NX_General_ComponentProperties_Ops"
	Dim sFuncName, sSubAction, sProperty
	Dim objCompPropDlg
	Dim bFlag, iCounter
	Dim dicCount, dicItems, dicKeys, aProperty
	
	Fn_SISW_NX_General_ComponentProperties_Ops = False
	sFuncName = "Fn_SISW_NX_General_ComponentProperties_Ops"
	
	Set objCompPropDlg = Fn_SISW_NX_GetObject("ComponentProperties")
	If objCompPropDlg.Exist(1) = False Then
		sMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("NX_Menu"),"Properties")
		bFlag = Fn_SISW_NX_General_MenuOperation("Select","Properties","")
		If bFlag = False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: ["&sFuncName&"] Failed to perform ["&sMenu&"] Menu/RMB operation on selected Node.")
			Set objCompPropDlg = Nothing
			Exit Function
		End If
		Call Fn_SISW_NX_Setup_ReadyStatusSync(1)
		If objCompPropDlg.Exist(1) = False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: ["&sFuncName&"] Failed as [Component Properties] dialog does not exist.")
			Set objCompPropDlg = Nothing
			Exit Function
		End If
	End If
	
	Select Case sAction
		Case "Set"
				dicCount = dicCompPropDetails.Count
				dicItems = dicCompPropDetails.Items
				dicKeys = dicCompPropDetails.Keys
				For iCounter = 0 To dicCount - 1
					If Instr(dicKeys(iCounter),"TabName")>0 Then
						sSubAction = "TabName"
					ElseIf Instr(dicKeys(iCounter),"CheckBox")>0 Then
						sSubAction = "CheckBox"
					ElseIf Instr(dicKeys(iCounter),"StaticText")>0 Then
						sSubAction = "StaticText"
					Else
						sSubAction = dicKeys(iCounter)
					End If
					sProperty = dicItems(iCounter)
					bFlag = False
					Select Case sSubAction
						Case "TabName"
							If sProperty<>"" Then
								bFlag = Fn_SISW_NX_UI_WinTabOperation(sFuncName, "Select", objCompPropDlg, "InternalTab", sProperty)
								If bFlag = False Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [Function-"&sFuncName&"] [Case-"&sAction&"] [SubCase-"&sSubAction&"] - Failed to Select Tab ["+sProperty+"] in [Component Properties] dialog.")
									Set objCompPropDlg = Nothing
									Exit Function
								End If
								Wait SISW_NX_MICRO_TIMEOUT
								bFlag = True
							End If
						Case "CheckBox"
							If sProperty<>"" Then
								aProperty = Split(sProperty,"~")
								objCompPropDlg.WinCheckBox("PropertyCheckBox").SetTOProperty "text", aProperty(0)
								bFlag = Fn_SISW_NX_UI_CheckBoxOperation(sFuncName, "Set", objCompPropDlg, "PropertyCheckBox", aProperty(1))
								If bFlag = False Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [Function-"&sFuncName&"] [Case-"&sAction&"] [SubCase-"&sSubAction&"] - Failed to Set CheckBox ["&aProperty(0)&"] value as ["&aProperty(1)&"] in [Component Properties] dialog.")
									Set objCompPropDlg = Nothing
									Exit Function
								End If
								Wait SISW_NX_MICRO_TIMEOUT
								bFlag = True
							End If
						Case "StaticText"
							If sProperty<>"" Then
								objCompPropDlg.Static("StatusText").SetTOProperty "text", sProperty
								If objCompPropDlg.Static("StatusText").Exist(2) Then
									bFlag = True
								Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [Function-"&sFuncName&"] [Case-"&sAction&"] [SubCase-"&sSubAction&"] - Failed to Set CheckBox ["&aProperty(0)&"] value as ["&aProperty(1)&"] in [Component Properties] dialog.")
									Set objCompPropDlg = Nothing
									Exit Function
								End If
								Wait SISW_NX_MICRO_TIMEOUT
							End If
					End Select
				Next
				
				If sButton <> "" Then
					bFlag = Fn_SISW_NX_UI_ButtonOperation(sFuncName, "Click", objCompPropDlg, sButton)
					If bFlag = False Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [Function-"&sFuncName&"] [Case-"&sAction&"] - Failed to Click on WinButton ["&sButton&"] value as ["&aProperty(1)&"] in [Component Properties] dialog.")
						Set objCompPropDlg = Nothing
						Exit Function
					End If
				End If
				Fn_SISW_NX_General_ComponentProperties_Ops = True
	End Select
	Set objCompPropDlg = Nothing
End Function
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Function Name     			: Fn_SISW_NX_Find_Component_Operation
' Function Description 			 : Function used to  on Find Component dialog in NX application
' Parameters			: 			sAction: Action
'												sMenu: Menu command
'												sReserve: Reserved for future use
'
' Return Value           :        True/False
'	
' Examples		    	: 		Call 	 Fn_SISW_NX_Find_Component_Operation("Find","000231/A;1-Comp1","OK","")
'							Call 	 Fn_SISW_NX_Find_Component_Operation("Find","Comp1","OK","")

' History               :  Developer Name			 	Date	  				Rev. No. 				Changes					Reviewer
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'				 Jotiba Takkekar	 	    		21-March-2018			1.0							
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_NX_Find_Component_Operation(sAction,sItemName,sButton,sReserve)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_NX_Find_Component_Operation"
   Dim objCommandFinder,WshShell
   Set objCommandFinder=Window("NXWindow").Dialog("FindComponent")
   Fn_SISW_NX_Find_Component_Operation=False
   
   If not objCommandFinder.Exist(2) Then
   	bReturn = Fn_SISW_NX_General_MenuOperation("Select","Find Component","")
	If bReturn=True Then
		Fn_SISW_NX_Find_Component_Operation=True
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully performed Menu operation for [Select->Find Component]")
	Else
		Fn_SISW_NX_Find_Component_Operation=False
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to perform Menu operation for [Select->Find Component]")
		Set objCommandFinder=Nothing
		Exit Function
	End If
   End If
   
   Select Case sAction
   	Case "Find","ComponentExists"
		If sItemName<>"" Then
		   	'set Item ID, Item Name
		   	bReturn=Fn_SISW_NX_UI_EditBoxOperation("Fn_SISW_NX_Find_Component_Operation","Set",objCommandFinder,"Name",sItemName)
		   	If bReturn=True Then
				Fn_SISW_NX_Find_Component_Operation=True
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully set value to EditBox [ Name ].")
			Else
				Fn_SISW_NX_Find_Component_Operation=False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to set value to EditBox [ Name ]")
				Set objCommandFinder=Nothing
				Exit Function
			End If
		  	 'Hit Enter
			 Set WshShell = CreateObject("WScript.Shell")
				WshShell.SendKeys "{ENTER}"
				Set WshShell =nothing
				wait 1
			Set WshShell =Nothing
			'------------------------------------------------
			If sAction = "ComponentExists" Then
				If Instr(Window("NXWindow").WinObject("StatusBar").GetVisibleText(),"No matches found") > 0 Then
					Fn_SISW_NX_Find_Component_Operation=False
				ElseIf Instr(Window("NXWindow").WinObject("StatusBar").GetVisibleText(),"Match i out erf 1") > 0 Then
					Fn_SISW_NX_Find_Component_Operation=True
				End IF  
			End If
			'-------------------------------------------------
		End If
   	End Select
   	
   	  If sButton<>"" Then
   	  		bReturn = Fn_SISW_NX_UI_ButtonOperation("Fn_SISW_NX_Find_Component_Operation","Click",objCommandFinder,"OK")
			If bReturn=False Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Click on OK button of Find Component Dialog")
				Set objCommandFinder=Nothing
				Exit Function
			End If
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Successfully Clicked on OK button of Find Component Dialog")
   	  End If
   Set objCommandFinder=Nothing
End Function
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Function Name     		: 	Fn_SISW_NX_Find_Object_Operation
' Function Description 	    : 	Function used to Find Object dialog in NX application
' Parameters				: 	sAction: Action
'								sItemName: object to serach
'								sFindIn : Find Scope
'								sReserve: Reserved for future use
'
' Return Value           	:  True/False
'	
' Examples		    		: Call Fn_SISW_NX_Find_Object_Operation("Find","Block","Object Name Only","OK","")
'							  Call Fn_SISW_NX_Find_Object_Operation("Find","Cylinder","Object Name and Attributes","OK","")
'
' History               	:  Developer Name			 Date	  			Rev. No. 				Changes					Reviewer
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'							 Poonam Chopade	 	    03-April-2018			1.0					   Created			TC11.5(20180226.00)_DIPRO_NewDevelopment					
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_NX_Find_Object_Operation(sAction,sItemName,sFindIn,sButton,sReserve)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_NX_Find_Object_Operation"
   Dim objCommandFinder,WshShell
   Set objCommandFinder=Window("NXWindow").Window("Find Object(to be retired)")
   Fn_SISW_NX_Find_Object_Operation=False
   
   If not objCommandFinder.Exist(2) Then
   	bReturn = Fn_SISW_NX_General_MenuOperation("Select","Find Object (to be retired)","")
	If bReturn=True Then
		Fn_SISW_NX_Find_Object_Operation=True
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully performed Menu operation for [Select->Find Object]")
	Else
		Fn_SISW_NX_Find_Object_Operation=False
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to perform Menu operation for [Select->Find Object]")
		Set objCommandFinder=Nothing
		Exit Function
	End If
   End If
   
   Select Case sAction
   	Case "Find"
   		If sFindIn <> "" Then
   			objCommandFinder.WinRadioButton("Object Name Only").SetTOProperty "text",sFindIn
   			objCommandFinder.WinRadioButton("Object Name Only").Set
   			Wait 1
		   	If Err.Number < 0 Then
				Fn_SISW_NX_Find_Object_Operation=False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to set Find In option as [ "&sFindIn&" ]")
				Set objCommandFinder=Nothing
				Exit Function
			Else
				Fn_SISW_NX_Find_Object_Operation=True
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully set Find In option as [ "&sFindIn&" ].")
			End If
   		End If
   	
		If sItemName<>"" Then
		   	'Node name
		   	bReturn=Fn_SISW_NX_UI_EditBoxOperation("Fn_SISW_NX_Find_Object_Operation","Set",objCommandFinder,"SearchString",sItemName)
		   	If bReturn=True Then
				Fn_SISW_NX_Find_Object_Operation=True
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully set value to EditBox [ Search String ].")
			Else
				Fn_SISW_NX_Find_Object_Operation=False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to set value to EditBox [ Search String ]")
				Set objCommandFinder=Nothing
				Exit Function
			End If
		  	 'Hit Enter
			 Set WshShell = CreateObject("WScript.Shell")
				WshShell.SendKeys "{ENTER}"
				Set WshShell =nothing
				wait 1
			Set WshShell =Nothing
		End If
		
   	End Select
   	
   	  If sButton<>"" Then
   	  		bReturn = Fn_SISW_NX_UI_ButtonOperation("Fn_SISW_NX_Find_Object_Operation","Click",objCommandFinder,"OK")
			If bReturn=False Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Click on OK button of Find Object Dialog")
				Set objCommandFinder=Nothing
				Exit Function
			End If
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Successfully Clicked on OK button of Find Object Dialog")
   	  End If
   Set objCommandFinder=Nothing
End Function
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Function Name     	  	: 	  Fn_SISW_NX_General_Model_ItemCreate
' Function Description 		: 	  Function used to Create Model Item in NX
' Parameters			    : 	  sItemName: Item Name
'								  sFolderPath: Folder path separated by :
'								  sReserve: Reserved for future use
' Return Value           :          True/False
'
' Examples		    	: 		Call Fn_SISW_NX_General_Model_ItemCreate(DataTable.Value("ItemName",dtGlobalSheet),Environment.Value("NXTestFolderName"),"","")
'
' History               :  
'			Developer Name	 		 	Date	  			Rev. No. 		Changes					Reviewer
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------
'			Poonam Chopade	 	    04-April-2018			1.0			 	Created			TC11.5_20180226.00_DIPRO_NewDevelopment				
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_NX_General_Model_ItemCreate(sItemName,sFolderPath,sReserve1,sReserve2)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_NX_General_Model_ItemCreate"
'	Update the Item Information Journal folder of TestData
	Dim bResult,sItemId,sItemRev
	Fn_SISW_NX_General_Model_ItemCreate=False
	bResult= Fn_UpdateEnvXMLNode(Environment.Value("sPath")&"\TestData\NX\Journal\ItemCreate_Input.xml", "ItemName",sItemName)
	If bResult=False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Update Item Name Information in ItemCreate_Input.xml")
			Exit Function
	End If
	bResult= Fn_UpdateEnvXMLNode(Environment.Value("sPath")&"\TestData\NX\Journal\ItemCreate_Input.xml", "Folder",sFolderPath)
	If bResult=False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to UpdateFolder Path in ItemCreate_Input.xml")
		Exit Function
	End If
    bResult=Fn_SISW_NX_General_JournalPlayBack(Environment.Value("sPath") &"\TestData\NX\Journal\Journal_ModelItem_Create.vb","","","","")
	If  bResult=False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Ran Journal File to create Item in NX application")
		Exit Function
	End If
	Call Fn_SISW_NX_Setup_ReadyStatusSync(2)
	sItemId=Fn_GetXMLNodeValue(Environment.Value("sPath")&"\TestData\NX\Journal\ItemCreate_Output.xml","ItemID")
	sItemRev=Fn_GetXMLNodeValue(Environment.Value("sPath")&"\TestData\NX\Journal\ItemCreate_Output.xml","ItemRevision")
	sItemName=Fn_GetXMLNodeValue(Environment.Value("sPath")&"\TestData\NX\Journal\ItemCreate_Output.xml","ItemName")
	Fn_SISW_NX_General_Model_ItemCreate="'"&sItemId&"~"&sItemRev &"~"& sItemName
End Function

'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Function Name     		: 	Fn_SISW_NX_VisualReportingOperations
'
' Function Description 	    : 	Function used to perform operation on Visual Reporting dialog in NX
'
' Parameters				: 	sAction: Action to be performed
'								dicReportDetails: Dictionary object
'								sButton : Button name
'
' Return Value           	:  True/False
'	
' Examples		    		: Set dicReportDetails = CreateObject("Scripting.Dictionary")
'								  dicReportDetails("EditBox1") = "ReportName~Report1"
'								  dicReportDetails("ComboBox1") = "Property~Number"
'							 bReturn = Fn_SISW_NX_VisualReportingOperations("Create",dicReportDetails,"OK")
'
' History               	:  Developer Name			 Date	  			Rev. No. 				Changes					Reviewer
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'					Jotiba Takkekar	 	    09-April-2018				1.0					   Created			TC11.5(20180226.00)_DIPRO_NewDevelopment					
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_NX_VisualReportingOperations(sAction,dicReportDetails,sButton)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_NX_VisualReportingOperations"
	
	Dim ObjVisualReport,ObjVisualReportDefination,dicCount,dicItems,dicKeys,iCounter,sSubAction,bReturn,aField,bFlag, iCnt

	Set ObjVisualReport=Window("NXWindow").WinButton("DefineNewReport")
	Set ObjVisualReportDefination=Fn_SISW_NX_GetObject("VisualReportDefinition")
	bFlag=False
	Fn_SISW_NX_VisualReportingOperations=False
	
'	 Check Existance of 
	If Fn_SISW_UI_Object_Operations("Fn_SISW_NX_VisualReportingOperations", "Exist", ObjVisualReport, 2) =False Then
		'Call to Invoke Visual Reporting Dialog
		Call Fn_SISW_NX_General_MenuOperation("Select","Start Visual Reporting","")
		Call Fn_SISW_NX_Setup_ReadyStatusSync(2)
		Wait 2
	End If
	
	If Fn_SISW_UI_Object_Operations("Fn_SISW_NX_VisualReportingOperations", "Exist", ObjVisualReport, 2) = False Then
		Fn_SISW_NX_VisualReportingOperations = False
    		Set ObjVisualReport = Nothing
    		Exit Function
	End If
	
	Select Case sAction
		Case "Create"
				If Fn_SISW_UI_Object_Operations("Fn_SISW_NX_VisualReportingOperations", "Exist", ObjVisualReportDefination, 2) = False Then
					bReturn=Fn_SISW_NX_UI_ButtonOperation("Fn_SISW_NX_VisualReportingOperations", "Click", Window("NXWindow"), "DefineNewReport")
					If bReturn=True Then
						Fn_SISW_NX_VisualReportingOperations=True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully Clicked on [ DefineNewReport ] button.")
					Else
						Fn_SISW_NX_VisualReportingOperations=False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed Click on [ DefineNewReport ] button.")
						Set ObjVisualReport=Nothing
						Set ObjVisualReportDefination=Nothing
						Exit Function
					End If
				End If
				
				' To open Action Tab
				If Fn_SISW_UI_Object_Operations("Fn_SISW_NX_VisualReportingOperations", "Exist", Window("NXWindow").Dialog("VisualReportDefinition").WinObject("ActionTab"), 2) = False Then
					Call Fn_SISW_NX_UI_CheckBoxOperation("Fn_SISW_NX_VisualReportingOperations", "Set", ObjVisualReportDefination, "ArrowButton", "ON")
					wait 1
					ObjVisualReportDefination.WinObject("ActionTab").Click
					wait 1
				Else
'					Click on New Report button
					Call Fn_SISW_NX_UI_ButtonOperation("Fn_SISW_NX_VisualReportingOperations", "Click", ObjVisualReportDefination, "NewReport")
					wait 1					
				End If 

			dicCount = dicReportDetails.Count
			dicItems = dicReportDetails.Items
			dicKeys = dicReportDetails.Keys
				
			For iCounter = 0 To dicCount - 1
				If  Instr(dicKeys(iCounter),"EditBox") > 0	Then
					sSubAction = "EditBox"
				ElseIf Instr(dicKeys(iCounter),"ComboBox") > 0  Then 
					sSubAction = "ComboBox"
				End If
					sField = dicItems(iCounter)
					
				Select Case sSubAction
							Case "EditBox"
									If sField<>"" Then
										aField=split(sField,"~")
										ObjVisualReportDefination.WinEdit("ReportName").SetTOProperty "attached text", aField(0)
										bReturn=Fn_SISW_NX_UI_EditBoxOperation("Fn_SISW_NX_VisualReportingOperations","Set",ObjVisualReportDefination,"ReportName",aField(1))
										If bReturn=True Then
											Fn_SISW_NX_VisualReportingOperations = True
										Else
											Fn_SISW_NX_VisualReportingOperations=False
											Set ObjVisualReport=Nothing
											Set ObjVisualReportDefination=Nothing
											Exit Function
										End If
									End If 
									
								Case "ComboBox"
									If sField<>"" Then
										aField=split(sField,"~")
										'set attached text property 
										ObjVisualReportDefination.WinComboBox("Property").SetTOProperty "attached text",aField(0)
										bReturn = Fn_SISW_NX_UI_ComboBoxOperation("Set",ObjVisualReportDefination,"Property","",aField(1),"")
										If bReturn=True Then
											Fn_SISW_NX_VisualReportingOperations = True
										Else
											Fn_SISW_NX_VisualReportingOperations=False
											Set ObjVisualReport=Nothing
											Set ObjVisualReportDefination=Nothing
											Exit Function
										End If
									End If
					 End Select
				Next 
				
				If sButton<>"" Then
					bReturn= Fn_SISW_NX_UI_ButtonOperation("Fn_SISW_NX_VisualReportingOperations", "Click", ObjVisualReportDefination, sButton)
					If bReturn= False Then
						Fn_SISW_NX_VisualReportingOperations=False
						Set ObjVisualReport=Nothing
						Set ObjVisualReportDefination=Nothing
						Exit Function
					End If
				End If
	End Select
End Function
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Function Name     		: 	Fn_SISW_NX_HandleMessageDialog
'
' Function Description 	    : 	Function used to handle Message Windows in NX
'
' Parameters				: 	sAction: Action
'								dicMsgDetails: msg details
'								sButton : Button name
'
' Return Value           	:  True/False
'	
' Examples		    		: Set dicMsgDetails = CreateObject("Scripting.Dictionary")
'								  dicMsgDetails("title") = "Check-out Parts"
'								  dicMsgDetails("Message") = "The part is read-only  000345/A;1-Item1 "
'							 bReturn = Fn_SISW_NX_HandleMessageDialog("VerifyStaticMessage",dicMsgDetails,"OK")
'
' History               	:  Developer Name			 Date	  			Rev. No. 				Changes					Reviewer
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'							 Poonam Chopade	 	    06-April-2018			1.0					   Created			TC11.5(20180226.00)_DIPRO_NewDevelopment					
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_NX_HandleMessageDialog(sAction,dicMsgDetails,sButton)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_NX_HandleMessageDialog"
    Dim objMsgWindow,sAppMsg
   
   Set objMsgWindow = Window("NXWindow").Dialog("MessageDialog")
   Fn_SISW_NX_HandleMessageDialog=False
   If dicMsgDetails("title") <> "" Then
   		objMsgWindow.SetTOProperty "text",dicMsgDetails("title")
   End If
   'Check Existence of dialog
    If not objMsgWindow.Exist(2) Then
    	Fn_SISW_NX_HandleMessageDialog = False
    	Set objMsgWindow = Nothing
    	Exit Function
    End If
   
   Select Case sAction
   		Case "VerifyStaticMessage"
	   		If dicMsgDetails("Message") <> "" Then
	   			sAppMsg = objMsgWindow.Static("Msg").GetROProperty("text")
			   	If Instr(sAppMsg,dicMsgDetails("Message")) > 0 Then
					Fn_SISW_NX_HandleMessageDialog = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully verified message [ "+dicMsgDetails("Message")+" ].")
				Else
					Fn_SISW_NX_HandleMessageDialog = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to verify message [ "+dicMsgDetails("Message")+" ].")
				End If
	   		End If	
   		Case "VerifyWinListMessage"
			sAppMsg = objMsgWindow.WinList("WinList").GetContent()
			If Instr(sAppMsg,dicMsgDetails("Message")) > 0 Then
					Fn_SISW_NX_HandleMessageDialog = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully verified message [ "+dicMsgDetails("Message")+" ].")
			Else
				Fn_SISW_NX_HandleMessageDialog = False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to verify message [ "+dicMsgDetails("Message")+" ].")
			End If
		Case "VerifyWinObjectMessage", "VerifyWinObjectMessageExt"
			sAppMsg = objMsgWindow.WinObject("Message").GetVisibleText()
			If sAction="VerifyWinObjectMessageExt" Then
				sAppMsg=Replace(sAppMsg," ","")
			End If
			If Instr(sAppMsg,dicMsgDetails("Message")) > 0 Then
					Fn_SISW_NX_HandleMessageDialog = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully verified message [ "+dicMsgDetails("Message")+" ].")
			Else
				Fn_SISW_NX_HandleMessageDialog = False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to verify message [ "+dicMsgDetails("Message")+" ].")
			End If	
		Case "Exists"
			If objMsgWindow.Exist(2) = True Then
				Fn_SISW_NX_HandleMessageDialog = True
			End If
		Case "ButtonClick"
			  Fn_SISW_NX_HandleMessageDialog = Fn_SISW_NX_UI_ButtonOperation("Fn_SISW_NX_HandleMessageDialog","Click",objMsgWindow,dicMsgDetails("Button"))
   	End Select
   	
   	 'Close code for some dialog 
   	  If dicMsgDetails("Close") = "Yes" Then
   	  		objMsgWindow.Close()
   	  End If	
   	
   	  If sButton<>"" Then
	  		'Click on button
   			Call Fn_SISW_NX_UI_ButtonOperation("Fn_SISW_NX_HandleMessageDialog","Click",objMsgWindow,sButton)
   			Wait 1
   	  End If
   	  
   Set objMsgWindow=Nothing
End Function
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Function Name     	   	:  Fn_SISW_NX_General_SaveAsExt()
' Function Description 	    :  	Function used to Save As on opened item / Part.
' Parameter					:   	sAction :- Action name
'									Dim dicSaveAsInfo
'									Set dicSaveAsInfo = CreateObject("Scripting.Dictionary")
'									with dicSaveAsInfo
'											.Add "SaveAsScope","Work Part"	
'	 										.Add "Action", "New Item"
'											.Add "ItemId", "Assembly0"
'											.Add "ItemRev", "A"
'										 .Add "ItemName", "Assembly0"
'									End with 
'								  sButton:- Button to be clicked	
'								  sReserve2- For future use	
'
' Return Value		    	  :   True/False/Values
'
' Examples		    	      :   Fn_SISW_NX_General_SaveAsExt("SaveAs",dicSaveAsInfo,"OK","")
'
' History					  :		
'
' Developer Name					Date				Rev. No.			Changes				Reviewer					
'----------------------------------------------------------------------------------------------------------------------------------------------
' Poonam Chopade				10-Apr-2018 			1.0					Created				TC11.5_20180329.00_DIPRO_NX_NewDevelopment		
'----------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_NX_General_SaveAsExt(sAction,dicSaveAsInfo,sButton,sReserve2)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_NX_General_SaveAsExt"
	 Dim objDialog,bReturn,sItemId,sRev,objArr,WshShell
	 Dim objLocation,L,T,R,B,X,Y

	Set objDialog = Window("NXWindow").Dialog("SavePartsAs")
	Set objSaveEdit = Window("NXWindow").Dialog("SavePartsAs").Dialog("SavePartAsEdit")
	Set WshShell = CreateObject("WScript.Shell")
	Set ObjFolderSelect=Fn_SISW_NX_GetObject("FolderSelect")
	 Fn_SISW_NX_General_SaveAsExt = False

	If objDialog.Exist(2) = False Then
			Call Fn_UpdateLogFiles("FAIL : Save As dialog does Not  Exist", "FAIL: Save As dialog does Not  Exist")
			Exit Function
	End If

Select Case sAction
		Case "SaveAs"
			If dicSaveAsInfo("SaveAsScope")<>"" Then
				bResult=Fn_SISW_NX_UI_ComboBoxOperation( "Set",objDialog, "SaveAsScope","", dicSaveAsInfo("SaveAsScope"),"")
				If  bResult=False Then
					  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select the Save As Scope Option of ComboBox ")
					  Exit Function
				End If
			End If
		
			If  dicSaveAsInfo("Action") = "Selected Parts" Then 
				If dicSaveAsInfo("LoadedPart") <> "" Then
					objLocation = Window("NXWindow").Dialog("SavePartsAs").WinObject("LoadedParts").GetTextLocation(dicSaveAsInfo("LoadedPart"), L, T, R, B)
					X = (L+R)/2
					Y = (T+B)/2	
					
					Window("NXWindow").Dialog("SavePartsAs").WinObject("LoadedParts").MouseMove X,Y
					Window("NXWindow").Dialog("SavePartsAs").WinObject("LoadedParts").Click X,Y
		      		L=0
					R=0
					T=0
					B=0
					X=0
					Y=0       	
				Else
					Exit Function
				End If
		    End If
		
			If dicSaveAsInfo("Action") <> "" Then
				bResult=Fn_SISW_NX_UI_ComboBoxOperation( "Set",objDialog, "SaveAs","", dicSaveAsInfo("Action"),"")
				If  bResult=False Then
					 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select the Save As Option of ComboBox ")
					 Exit Function
				End If
			End If
		
			If dicSaveAsInfo("ItemId") <> "" OR dicSaveAsInfo("ItemRev") <> "" Then						
				If dicSaveAsInfo("ItemId") <> "" Then
					call objDialog.WinObject("QWidget").GetTextLocation("ID",L, T, R, B,True)
					X = L + 10
					Y = B + 10
					objDialog.WinObject("QWidget").Click X ,Y ,micLeftBtn
					objDialog.WinObject("QWidget").Click X ,Y ,micRightBtn
		      		L=0
					R=0
					T=0
					B=0
					X=0
					Y=0
					call objDialog.Window("PopUpMenu").GetTextLocation("Edit...",L, T, R, B,True)
					X = L + 10
					Y = T + 10
					objDialog.Window("PopUpMenu").Click X, Y ,micLeftBtn
					wait 2			
					objSaveEdit.WinEdit("SavaAsID").SetTOProperty "attached text" ,"ID"
					objSaveEdit.WinEdit("SavaAsID").Type dicSaveAsInfo("ItemId")
					WshShell.SendKeys "{ENTER}"
					Wait 1
					objSaveEdit.WinButton("OK").Click
					If objSaveEdit.Exist(2) Then
						objSaveEdit.Close()	
					End If
					L=0
					R=0
					T=0
					B=0
					X=0
					Y=0		
				End If
				Wait 1
				If dicSaveAsInfo("ItemRev") <> "" Then
					call objDialog.WinObject("QWidget").GetTextLocation("Revision",L, T, R, B,True)
					X = L + 40
					Y = B + 40
					objDialog.WinObject("QWidget").Click X ,Y ,micLeftBtn
					objDialog.WinObject("QWidget").Click X ,Y ,micRightBtn
		      		L=0
					R=0
					T=0
					B=0
					X=0
					Y=0
					call objDialog.Window("PopUpMenu").GetTextLocation("Edit...",L, T, R, B,True)
					X = L + 10
					Y = T + 10
					objDialog.Window("PopUpMenu").Click X, Y ,micLeftBtn
					wait 2			
					objSaveEdit.WinEdit("SavaAsID").SetTOProperty "attached text" ,"Revision"
					objSaveEdit.WinEdit("SavaAsID").Type dicSaveAsInfo("ItemRev")
					WshShell.SendKeys "{ENTER}"
					Wait 1
					objSaveEdit.WinButton("OK").Click
					If objSaveEdit.Exist(2) Then
						objSaveEdit.Close()	
					End If
					L=0
					R=0
					T=0
					B=0
					X=0
					Y=0		
				End If
				Wait 1
			Else 	
				call objDialog.WinObject("QWidget").GetTextLocation("ID",L, T, R, B,True)
				X = L + 10
				Y = B + 10
				objDialog.WinObject("QWidget").Click X, Y, micLeftBtn
				objDialog.WinObject("QWidget").DblClick X, Y, micLeftBtn	
					L=0
					R=0
					T=0
					B=0
					X=0
					Y=0		
		    End If
		
		    If dicSaveAsInfo("ItemName") <> "" Then
					call objDialog.WinObject("QWidget").GetTextLocation("Revision",L, T, R, B,True)
					X = R + 40
					Y = B + 40
					objDialog.WinObject("QWidget").Click X - 200,Y ,micLeftBtn
					objDialog.WinObject("QWidget").Click X - 200,Y ,micRightBtn
		      		L=0
					R=0
					T=0
					B=0
					X=0
					Y=0
					call objDialog.Window("PopUpMenu").GetTextLocation("Edit...",L, T, R, B,True)
					X = L + 10
					Y = T + 10
					objDialog.Window("PopUpMenu").Click X, Y ,micLeftBtn
					wait 2			
					objSaveEdit.WinEdit("SavaAsID").SetTOProperty "attached text" ,"Name"
					objSaveEdit.WinEdit("SavaAsID").Type dicSaveAsInfo("ItemName")
					WshShell.SendKeys "{ENTER}"
					Wait 1
					objSaveEdit.WinButton("OK").Click
					If objSaveEdit.Exist(2) Then
						objSaveEdit.Close()	
					End If
					L=0
					R=0
					T=0
					B=0
					X=0
					Y=0	    
			End If
		
					call objDialog.WinObject("QWidget").GetTextLocation("ID",L, T, R, B,True)
					X = L + 10
					Y = B + 10
					objDialog.WinObject("QWidget").Click X ,Y ,micLeftBtn
					objDialog.WinObject("QWidget").Click X ,Y ,micRightBtn
		      		L=0
					R=0
					T=0
					B=0
					X=0
					Y=0
					call objDialog.Window("PopUpMenu").GetTextLocation("Edit...",L, T, R, B,True)
					X = L + 10
					Y = T + 10
					objDialog.Window("PopUpMenu").Click X, Y ,micLeftBtn
					wait 2	
					objSaveEdit.WinEdit("SavaAsID").SetTOProperty "attached text" ,"ID"			
					sItemID = objSaveEdit.WinEdit("SavaAsID").GetROProperty("text")
					objSaveEdit.WinButton("OK").Click
					If objSaveEdit.Exist(2) Then
						objSaveEdit.Close()	
					End If
					L=0
					R=0
					T=0
					B=0
					X=0
					Y=0	
		
					call objDialog.WinObject("QWidget").GetTextLocation("Revision",L, T, R, B,True)
					X = L + 40
					Y = B + 40
					objDialog.WinObject("QWidget").Click X ,Y ,micLeftBtn
					objDialog.WinObject("QWidget").Click X ,Y ,micRightBtn
		      		L=0
					R=0
					T=0
					B=0
					X=0
					Y=0
					call objDialog.Window("PopUpMenu").GetTextLocation("Edit...",L, T, R, B,True)
					X = L + 10
					Y = T + 10
					objDialog.Window("PopUpMenu").Click X, Y ,micLeftBtn
					wait 2	
					objSaveEdit.WinEdit("SavaAsID").SetTOProperty "attached text" ,"Revision"			
					sRev = objSaveEdit.WinEdit("SavaAsID").GetROProperty("text")
					objSaveEdit.WinButton("OK").Click
					If objSaveEdit.Exist(2) Then
						objSaveEdit.Close()	
					End If
					L=0
					R=0
					T=0
					B=0
					X=0
					Y=0	
					
			If dicSaveAsInfo("Folder")="Newstuff" Then
						bReturn = Fn_SISW_NX_UI_ButtonOperation("Fn_SISW_NX_General_SaveAsExt","Click",objDialog,"SelectFolder")
						If bReturn=False Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Click on [ FolderButton ] button of [ Save As Non Master Parts ] Dialog")
							Fn_SISW_NX_General_SaveAsExt=False
							Exit Function
						End If
					
						Call Fn_SISW_NX_UI_ComboBoxOperation("Set",ObjFolderSelect,"LookIn","","Home","")
						wait 2
						
						'Select Newstuff folder
						Call ObjFolderSelect.WinObject("NodeList").GetTextLocation("Newstuff",L, T, R, B,True)
                                  ObjFolderSelect.WinObject("NodeList").DblClick L,T,micLeftBtn
						wait 2
						
						bReturn=Fn_SISW_NX_UI_ButtonOperation("", "Click", ObjFolderSelect, "OK")
						If bReturn=False Then
							Call Fn_UpdateLogFiles("FAIL : Fail to click on  [ Ok ] button", "FAIL: Fail to click on  [ Ok ] button")
							Fn_SISW_NX_General_SaveAsExt=false
							Set ObjSave=Nothing
							Set ObjFolderSelect=Nothing
							Exit Function
						Else
							Fn_SISW_NX_General_SaveAsExt=True
							wait 1
						End If
			End If
			bResult=Fn_SISW_NX_UI_ComboBoxOperation( "Set",objDialog, "DependentFilesSaveAs","", "Save All","")
			If bResult=False Then
					Fn_SISW_NX_Setup_CmdFinderOperation=False
					Set ObjCmdDialog=Nothing
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to set value Save All in combobox")
					Exit Function
			End If
			Fn_SISW_NX_General_SaveAsExt = sItemId & "-" & sRev
			
			
	End Select
	
	If sButton <> "" Then
		bReturn = Fn_SISW_NX_UI_ButtonOperation("Fn_SISW_NX_General_SaveAsExt","Click",objDialog,sButton)
		If bReturn=False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Click on ["+sButton+"] button of Save As Dialog")
			Set ObjDialog=Nothing
			Exit Function
		Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Successfully Clicked on  ["+sButton+"] button of Save As Dialog")
		End If
	End If
	Set ObjDialog=Nothing
	Set objSaveEdit = Nothing
	Set WshShell = Nothing
End Function
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Function Name     		: 	Fn_SISW_NX_ClosePartOperations()
'
' Function Description 	    : 	Function used to perform operations on Close Part dialog Windows in NX
'
' Parameters				: 	sAction: Action
'								dicDetails: Close Part details
'								sButton : Button name
'
' Return Value           	:  True/False
'	
' Examples		    		: Set dicDetails = CreateObject("Scripting.Dictionary")
'								  dicDetails("Filter") = "All Parts in Session:ON"
'								  dicDetails("Part") = "000345/A;1-Item1 "
'								  dicDetails("CloseType") = "Part and Components:ON"	
'							 bReturn = Fn_SISW_NX_ClosePartOperations("ClosePart",dicDetails,"OK")
'
' History               	:  Developer Name			 Date	  			Rev. No. 				Changes					Reviewer
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'							 Poonam Chopade	 	    16-April-2018			1.0					   Created			TC11.5(20180226.00)_DIPRO_NX_NewDevelopment					
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_NX_ClosePartOperations(sAction,dicDetails,sButton)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_NX_ClosePartOperations"
    Dim objClosePartWindow,sField,iCounter
   
   Set objClosePartWindow = Window("NXWindow").Dialog("ClosePart")
   Fn_SISW_NX_ClosePartOperations=False
   'Check Existence of Window
   If Fn_SISW_UI_Object_Operations("Fn_SISW_NX_ClosePartOperations", "Exist", objClosePartWindow, 2) = False Then
		'Call to Invoke Close Part Dialog
'		Call Fn_SISW_NX_General_MenuOperation("Select","Close Selected Part","")
		Call Fn_SISW_NX_Setup_CmdFinderOperation("Close Selected Part","Start2","")
		Call Fn_SISW_NX_Setup_ReadyStatusSync(2)
		Wait 2
		'Check Existence
		If Fn_SISW_UI_Object_Operations("Fn_SISW_NX_ClosePartOperations", "Exist", objClosePartWindow, 2) = False Then
			Fn_SISW_NX_ClosePartOperations = False
			Set objClosePartWindow = Nothing
			Exit Function
		End If	
	End If
	
    Select Case sAction
   		Case "ClosePart"
			'Set Filter type 
			If dicDetails("Filter") <> "" Then	 
				sField = Split(dicDetails("Filter"),":")
				objClosePartWindow.WinRadioButton("Filter").SetTOProperty "text",sField(0)
				objClosePartWindow.WinRadioButton("Filter").Set
				Wait 1
				If Err.Number < 0 Then
					Fn_SISW_NX_ClosePartOperations = False
					Set objClosePartWindow = Nothing
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed select Filter [ "&sField(0)&" ].")
					Exit Function
				End If
			End If
			'Select Part to close
			If dicDetails("Part") <> "" Then	 
				Call Fn_SISW_NX_UI_EditBoxOperation("Fn_SISW_NX_ClosePartOperations","Set",objClosePartWindow,"Search",dicDetails("Part"))
				Wait 1
				objClosePartWindow.WinObject("ViewStyle").Click 5,5,micLeftBtn
				Wait 1
				If instr(objClosePartWindow.WinEdit("PartName").GetROProperty("text"),dicDetails("Part")) = 0 Then
					Fn_SISW_NX_ClosePartOperations = False
					Set objClosePartWindow = Nothing
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed select Part [ "&dicDetails("Part")&" ].")
					Exit Function
				End If
			End If
			'Select Close Type
			If dicDetails("CloseType") <> "" Then	 
				sField = Split(dicDetails("CloseType"),":")
				objClosePartWindow.WinRadioButton("CloseType").SetTOProperty "text",sField(0)
				objClosePartWindow.WinRadioButton("CloseType").Set
				Wait 1
				If Err.Number < 0 Then
					Fn_SISW_NX_ClosePartOperations = False
					Set objClosePartWindow = Nothing
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed select Type [ "&sField(0)&" ].")
					Exit Function
				End If
			End If
   	End Select
    
   	 If sButton<>"" Then
	  		'Click on button
			sButton = Split(sButton,":")
			For iCounter = 0 To UBound(sButton)
				Call Fn_SISW_NX_UI_ButtonOperation("Fn_SISW_NX_ClosePartOperations","Click",objClosePartWindow,sButton(iCounter))
				Wait 1
			Next	
   	  End If
   	 
   Fn_SISW_NX_ClosePartOperations = True	   	 
   Set objClosePartWindow=Nothing
End Function

'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Function Name     	   	:  Fn_SISW_NX_Name_Parts_For_Save()
' Function Description 	    :  	Function used to Save As operation on Name Parts For Save
' Parameter					:   	sAction :- Action name
'									Dim dicSaveAsInfo
'									Set dicSaveAsInfo = CreateObject("Scripting.Dictionary")
'									with dicSaveAsInfo
'											.Add "ItemId", "Assembly0"
'											.Add "ItemRev", "A"
'										 .Add "ItemName", "Assembly0"
'									End with 
'								  sButton:- Button to be clicked	
'								  sReserve2- For future use	
'
' Return Value		    	  :   True/False/Values
'
' Examples		    	      :   Fn_SISW_NX_Name_Parts_For_Save("SaveAs",dicSaveAsInfo,"OK","")
'
' History					  :		
'
' Developer Name					Date				Rev. No.			Changes				Reviewer					
'----------------------------------------------------------------------------------------------------------------------------------------------
' Jotiba Takkekar				13-Apr-2018 			1.0					Created				TC11.5_20180329.00_DIPRO_NX_NewDevelopment		
'----------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_NX_Name_Parts_For_Save(sAction,dicSaveAsInfo,sButton,sReserve2)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_NX_Name_Parts_For_Save"
	 Dim objDialog,bReturn,sItemId,sRev,WshShell
	 Dim objLocation,L,T,R,B,X,Y

	Set objDialog = Window("NXWindow").Dialog("NamePartsForSave")
	Set objSaveEdit = Window("NXWindow").Dialog("NamePartsForSave").Dialog("EditDialog")
	Set WshShell = CreateObject("WScript.Shell")
	 Fn_SISW_NX_Name_Parts_For_Save = False


	If objDialog.Exist(2) = False Then
		bReturn=Fn_SISW_NX_General_MenuOperation("Select","Save As","")
		If bReturn=False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to open  [ Name Parts For Save ] dialog.")
			Set ObjDialog=Nothing
			Set objSaveEdit=Nothing
			Exit Function
		End If
	End If

	If objDialog.Exist(2) = False Then
			Call Fn_UpdateLogFiles("FAIL : [ Name Parts For Save ] dialog does Not  Exist", "FAIL: [ Name Parts For Save ] dialog does Not  Exist")
			Exit Function
	End If

Select Case sAction
		Case "SaveAs"
		
			If dicSaveAsInfo("ItemId") <> "" OR dicSaveAsInfo("ItemRev") <> "" Then						
				If dicSaveAsInfo("ItemId") <> "" Then
					call objDialog.WinObject("QWidget").GetTextLocation("ID",L, T, R, B,True)
					X = L + 10
					Y = B + 10
					objDialog.WinObject("QWidget").Click X ,Y ,micLeftBtn
					objDialog.WinObject("QWidget").Click X ,Y ,micRightBtn
		      		L=0
					R=0
					T=0
					B=0
					X=0
					Y=0
					call objDialog.Window("PopUpMenu").GetTextLocation("Edit...",L, T, R, B,True)
					X = L + 10
					Y = T + 10
					objDialog.Window("PopUpMenu").Click X, Y ,micLeftBtn
					wait 2			
					objSaveEdit.WinEdit("SavaAsID").SetTOProperty "attached text" ,"ID"
					objSaveEdit.WinEdit("SavaAsID").Type dicSaveAsInfo("ItemId")
					WshShell.SendKeys "{ENTER}"
					Wait 1
					objSaveEdit.WinButton("OK").Click
					If objSaveEdit.Exist(2) Then
						objSaveEdit.Close()	
					End If
					L=0
					R=0
					T=0
					B=0
					X=0
					Y=0		
				End If
				Wait 1
				If dicSaveAsInfo("ItemRev") <> "" Then
					call objDialog.WinObject("QWidget").GetTextLocation("Revision",L, T, R, B,True)
					X = L + 40
					Y = B + 40
					objDialog.WinObject("QWidget").Click X + 50 ,Y ,micLeftBtn
					objDialog.WinObject("QWidget").Click X + 50 ,Y ,micRightBtn
		      		L=0
					R=0
					T=0
					B=0
					X=0
					Y=0
					call objDialog.Window("PopUpMenu").GetTextLocation("Edit...",L, T, R, B,True)
					X = L + 10
					Y = T + 10
					objDialog.Window("PopUpMenu").Click X, Y ,micLeftBtn
					wait 2			
					objSaveEdit.WinEdit("SavaAsID").SetTOProperty "attached text" ,"Revision"
					objSaveEdit.WinEdit("SavaAsID").Type dicSaveAsInfo("ItemRev")
					WshShell.SendKeys "{ENTER}"
					Wait 1
					objSaveEdit.WinButton("OK").Click
					If objSaveEdit.Exist(2) Then
						objSaveEdit.Close()	
					End If
					L=0
					R=0
					T=0
					B=0
					X=0
					Y=0		
				End If
				Wait 1
			Else 	
				call objDialog.WinObject("QWidget").GetTextLocation("ID",L, T, R, B,True)
				X = L + 10
				Y = B + 10
				objDialog.WinObject("QWidget").Click X, Y, micLeftBtn
				objDialog.WinObject("QWidget").DblClick X, Y, micLeftBtn	
					L=0
					R=0
					T=0
					B=0
					X=0
					Y=0		
		    End If
		
		    If dicSaveAsInfo("ItemName") <> "" Then
					call objDialog.WinObject("QWidget").GetTextLocation("Revision",L, T, R, B,True)
					X = R + 40
					Y = B + 40
					objDialog.WinObject("QWidget").Click X - 200,Y ,micLeftBtn
					objDialog.WinObject("QWidget").Click X - 200,Y ,micRightBtn
		      		L=0
					R=0
					T=0
					B=0
					X=0
					Y=0
					call objDialog.Window("PopUpMenu").GetTextLocation("Edit...",L, T, R, B,True)
					X = L + 10
					Y = T + 10
					objDialog.Window("PopUpMenu").Click X, Y ,micLeftBtn
					wait 2			
					objSaveEdit.WinEdit("SavaAsID").SetTOProperty "attached text" ,"Name"
					objSaveEdit.WinEdit("SavaAsID").Type dicSaveAsInfo("ItemName")
					WshShell.SendKeys "{ENTER}"
					Wait 1
					objSaveEdit.WinButton("OK").Click
					If objSaveEdit.Exist(2) Then
						objSaveEdit.Close()	
					End If
					L=0
					R=0
					T=0
					B=0
					X=0
					Y=0	    
			End If
		
					call objDialog.WinObject("QWidget").GetTextLocation("ID",L, T, R, B,True)
					X = L + 10
					Y = B + 10
					objDialog.WinObject("QWidget").Click X ,Y ,micLeftBtn
					objDialog.WinObject("QWidget").Click X ,Y ,micRightBtn
		      		L=0
					R=0
					T=0
					B=0
					X=0
					Y=0
					call objDialog.Window("PopUpMenu").GetTextLocation("Edit...",L, T, R, B,True)
					X = L + 10
					Y = T + 10
					objDialog.Window("PopUpMenu").Click X, Y ,micLeftBtn
					wait 2	
					objSaveEdit.WinEdit("SavaAsID").SetTOProperty "attached text" ,"ID"			
					sItemID = objSaveEdit.WinEdit("SavaAsID").GetROProperty("text")
					objSaveEdit.WinButton("OK").Click
					If objSaveEdit.Exist(2) Then
						objSaveEdit.Close()	
					End If
					L=0
					R=0
					T=0
					B=0
					X=0
					Y=0	
		
					call objDialog.WinObject("QWidget").GetTextLocation("Revision",L, T, R, B,True)
					X = L + 40
					Y = B + 40
					objDialog.WinObject("QWidget").Click X + 50 ,Y ,micLeftBtn
					objDialog.WinObject("QWidget").Click X + 50 ,Y ,micRightBtn
		      		L=0
					R=0
					T=0
					B=0
					X=0
					Y=0
					call objDialog.Window("PopUpMenu").GetTextLocation("Edit...",L, T, R, B,True)
					X = L + 10
					Y = T + 10
					objDialog.Window("PopUpMenu").Click X, Y ,micLeftBtn
					wait 2	
					objSaveEdit.WinEdit("SavaAsID").SetTOProperty "attached text" ,"Revision"			
					sRev = objSaveEdit.WinEdit("SavaAsID").GetROProperty("text")
					objSaveEdit.WinButton("OK").Click
					If objSaveEdit.Exist(2) Then
						objSaveEdit.Close()	
					End If
					L=0
					R=0
					T=0
					B=0
					X=0
					Y=0	

				Fn_SISW_NX_Name_Parts_For_Save = sItemId & "-" & sRev
	End Select
	
	If sButton <> "" Then
		bReturn = Fn_SISW_NX_UI_ButtonOperation("Fn_SISW_NX_Name_Parts_For_Save","Click",objDialog,sButton)
		If bReturn=False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Click on ["+sButton+"] button of Name Parts for Save Dialog")
			Set ObjDialog=Nothing
			Exit Function
		Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Successfully Clicked on  ["+sButton+"] button of Name Parts for Save Dialog")
		End If
	End If
	Set ObjDialog=Nothing
	Set objSaveEdit = Nothing
	Set WshShell = Nothing
End Function
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Function Name     	   	:  Fn_SISW_NX_NewItem_ItemType_Operation()
' Function Description 	    :  	Function used to Save As on opened item / Part.
' Parameter					:   	sAction :- Action name
'									Dim dicItemInfo
'									Set dicItemInfo = CreateObject("Scripting.Dictionary")
'									with dicItemInfo
'											.Add "TabName", Blank"
'											.Add "ItemType", "Item"
'											.Add "TemplateName", "Blank"
'									End with 
'								  sButton:- Button to be clicked	
'								  sReserve2- For future use	
'
' Return Value		    	  :   True/False/Values
'
' Examples		    	      :   Fn_SISW_NX_Name_Parts_For_Save("VerifyItemType",dicItemInfo,"OK","")
'
' History					  :		
'
' Developer Name					Date				Rev. No.			Changes				Reviewer					
'----------------------------------------------------------------------------------------------------------------------------------------------
' Jotiba Takkekar				13-Apr-2018 			1.0					Created				TC11.5_20180329.00_DIPRO_NX_NewDevelopment		
'----------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_NX_NewItem_ItemType_Operation(sAction,dicItemInfo,sButton,sReserve2)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_NX_NewItem_ItemType_Operation"
	 Dim ObjNewItem,bReturn
	 Set ObjNewItem=Fn_SISW_NX_GetObject("NewItem")
	 Fn_SISW_NX_NewItem_ItemType_Operation=False
	 
	 If ObjNewItem.Exist(2)=False Then
	 	bReturn=Fn_SISW_NX_General_MenuOperation("Select","File New Item","")
	 	Call Fn_SISW_NX_Setup_ReadyStatusSync(2)
	 	If bReturn=False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to open  [ New Item ] dialog.")
			Set ObjNewItem=Nothing
			Exit Function
		End If
	End If

	If ObjNewItem.Exist(2) = False Then
			Call Fn_UpdateLogFiles("FAIL : [ New Item ] dialog does Not  Exist", "FAIL: [ New Item ] dialog does Not  Exist")
			Set ObjNewItem=Nothing
			Exit Function
	End If
	
	
	Select Case sAction
		Case "SelectBlank", "VerifyItemType","ItemTypeCount"
				If dicItemInfo("TemplateName")="Blank" Then
					If dicItemInfo("TabName")  <> "" Then
						bReturn= Fn_SISW_NX_UI_WinTabOperation("", "Select", Window("NXWindow").Dialog("NewItem"), "TabName", dicItemInfo("TabName"))
						If bReturn=True Then
							Fn_SISW_NX_NewItem_ItemType_Operation=True
						End If
						wait 1
						For iCnt = 1 To 2
							Call Fn_KeyBoardOperation("SendKeys", "{TAB}") 
						Next
						
						For iCnt = 1 To 3
							bReturn= Fn_KeyBoardOperation("SendKeys", "{PGDN}") 
						Next
					End If
				End If
				
				If sAction ="VerifyItemType" Then 
					 bReturn=Fn_SISW_NX_UI_ComboBoxOperation("Exist",ObjNewItem,"ItemType","",dicItemInfo("ItemType"),"")
					If bReturn=False Then
						Call Fn_UpdateLogFiles("FAIL : Item Type - [ "+dicItemInfo("ItemName")+" ] does Not  Exist", "FAIL: Item Type - [ "+dicItemInfo("ItemName")+" ] does Not  Exist")
						Fn_SISW_NX_NewItem_ItemType_Operation=False
						Set ObjNewItem=Nothing
						Exit Function
					Else
						Fn_SISW_NX_NewItem_ItemType_Operation=True
					End If
				End If
				
				If sAction ="ItemTypeCount" Then 
					Fn_SISW_NX_NewItem_ItemType_Operation=ObjNewItem.WinComboBox("ItemType").GetROProperty("items count")
				End If
		End Select 
		
		If sButton <> "" Then
			bReturn = Fn_SISW_NX_UI_ButtonOperation("Fn_SISW_NX_NewItem_ItemType_Operation","Click",ObjNewItem,sButton)
			If bReturn=False Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Click on ["+sButton+"] button of [ New Item ] Dialog")
				Fn_SISW_NX_NewItem_ItemType_Operation=False
				Set ObjNewItem=Nothing
				Exit Function
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Successfully Clicked on  ["+sButton+"] button of [ New Item ] Dialog")
			End If
		End If
End Function

'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Function Name     	   	:  Fn_SISW_NX_OutputPartNumber_Operation()
' Function Description 	    :  	Function used to perform operation on Output Part Number
' Parameter					:   	sAction :- Action name
'									Dim dicOutputInfo
'									Set dicOutputInfo = CreateObject("Scripting.Dictionary")
'									with dicOutputInfo
'											.Add "PartType", Item"
'											.Add "PartNumber", "TCQA_TEST_105_023"
'											.Add "PartRevision", "A"
'									End with 
'								  sButton:- Button to be clicked	
'								  sReserve2- For future use	
'
' Return Value		    	  :   True/False/Values
'
' Examples		    	      :   Fn_SISW_NX_OutputPartNumber_Operation("Set",dicOutputInfo,"OK","")
'
' History					  :		
'
' Developer Name					Date				Rev. No.			Changes				Reviewer					
'----------------------------------------------------------------------------------------------------------------------------------------------
' Jotiba Takkekar				16-Apr-2018 			1.0					Created				TC11.5_20180329.00_DIPRO_NX_NewDevelopment		
'----------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_NX_OutputPartNumber_Operation(sAction,dicOutputInfo,sButton,sReserve2)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_NX_OutputPartNumber_Operation"
	 Dim ObjOutput,bReturn
	 Set ObjOutput=Fn_SISW_NX_GetObject("OutputPartNumber")
	 Fn_SISW_NX_OutputPartNumber_Operation=False
	 
	If ObjOutput.Exist(2) = False Then
			Call Fn_UpdateLogFiles("FAIL : [ Output Part Number ] dialog does Not  Exist", "FAIL: [ Output Part Number ] dialog does Not  Exist")
			Set ObjOutput=Nothing
			Exit Function
	Else
			ObjOutput.Click 0,105,micLeftBtn
	End If
	
	
	Select Case sAction
		Case "Set"
				If dicOutputInfo("PartType")<>"" Then
					bReturn=Fn_SISW_NX_UI_ComboBoxOperation("Set",ObjOutput,"PartType","",dicOutputInfo("PartType"),"")
					If bReturn=False Then
						Call Fn_UpdateLogFiles("FAIL : Fail to set- "&dicOutputInfo("PartType"), "FAIL: Fail to set - "&dicOutputInfo("PartType"))
						Fn_SISW_NX_OutputPartNumber_Operation=false
						Set ObjOutput=Nothing
						Exit Function
					Else
						wait 1
						Fn_SISW_NX_OutputPartNumber_Operation=True
					End If
				End If
				
				If dicOutputInfo("PartNumber")<>"" Then
					bReturn=Fn_SISW_NX_UI_EditBoxOperation("Fn_SISW_NX_OutputPartNumber_Operation","Set",ObjOutput,"PartNumber", dicOutputInfo("PartNumber"))
					If bReturn=False Then
						Call Fn_UpdateLogFiles("FAIL : Fail to set - "&dicOutputInfo("PartNumber"), "FAIL: Fail to set - "&dicOutputInfo("PartNumber"))
						Fn_SISW_NX_OutputPartNumber_Operation=false
						Set ObjOutput=Nothing
						Exit Function
					Else
						wait 1
						Fn_SISW_NX_OutputPartNumber_Operation=True
					End If
				End If
				
				If dicOutputInfo("PartRevision")<>"" Then
					bReturn=Fn_SISW_NX_UI_EditBoxOperation("Fn_SISW_NX_OutputPartNumber_Operation","Set",ObjOutput,"PartRevision", dicOutputInfo("PartRevision"))
					If bReturn=False Then
						Call Fn_UpdateLogFiles("FAIL : Fail to set - "&dicOutputInfo("PartRevision"), "FAIL: Fail to set - "&dicOutputInfo("PartRevision"))
						Fn_SISW_NX_OutputPartNumber_Operation=false
						Set ObjOutput=Nothing
						Exit Function
					Else
						Fn_SISW_NX_OutputPartNumber_Operation=True
					End If
				End If
				
		Case "Verify"
				If dicOutputInfo("PartType")<>"" Then
					bReturn=Fn_SISW_NX_UI_ComboBoxOperation("Exist",ObjOutput,"PartType","",dicOutputInfo("PartType"),"")
					If bReturn=False Then
						Call Fn_UpdateLogFiles("FAIL : Fail to verify- "&dicOutputInfo("PartType"), "FAIL: Fail to verify - "&dicOutputInfo("PartType"))
						Fn_SISW_NX_OutputPartNumber_Operation=false
						Set ObjOutput=Nothing
						Exit Function
					Else
						Fn_SISW_NX_OutputPartNumber_Operation=True
					End If
				End If
				
		Case "PartTypeCount"
					Fn_SISW_NX_OutputPartNumber_Operation=ObjOutput.WinComboBox("PartType").GetROProperty("items count")
					
		End Select 
		
		If sButton <> "" Then
			bReturn = Fn_SISW_NX_UI_ButtonOperation("Fn_SISW_NX_OutputPartNumber_Operation","Click",ObjOutput,sButton)
			If bReturn=False Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Click on ["+sButton+"] button of [ Output Part Number ] Dialog")
				Fn_SISW_NX_OutputPartNumber_Operation=False
				Set ObjNewItem=Nothing
				Exit Function
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Successfully Clicked on  ["+sButton+"] button of [ Output Part Number ] Dialog")
			End If
		End If
End Function
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Function Name     	   	:  Fn_SISW_NX_SaveAsNonMasterParts_Operation()
' Function Description 	    :  	Function used to perform operation on Output Part Number
' Parameter					:   	sAction :- Action name
'									Dim dicSaveAsInfo
'									Set dicSaveAsInfo = CreateObject("Scripting.Dictionary")
'									with dicSaveAsInfo
'											.Add "Number", "1234"
'											.Add "Revision", "TCQA_TEST_105_023"
'											.Add "ItemName", "A"
'									End with 
'								  sButton:- Button to be clicked	
'								  sReserve2- For future use	
'
' Return Value		    	  :   True/False/Values
'
' Examples		    	      :   Fn_SISW_NX_SaveAsNonMasterParts_Operation("SaveAs",dicSaveAsInfo,"OK","")
'
' History					  :		
'
' Developer Name					Date				Rev. No.			Changes				Reviewer					
'----------------------------------------------------------------------------------------------------------------------------------------------
' Jotiba Takkekar				16-Apr-2018 			1.0					Created				TC11.5_20180329.00_DIPRO_NX_NewDevelopment		
'----------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_NX_SaveAsNonMasterParts_Operation(sAction,dicSaveAsInfo,sButton,sReserve2)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_NX_SaveAsNonMasterParts_Operation"
	 Dim ObjSave,bReturn
	 Dim L,T,R,B,X,Y
	 Set ObjSave=Fn_SISW_NX_GetObject("SaveAsNonMasterParts")
	 Set ObjFolderSelect=Fn_SISW_NX_GetObject("FolderSelect")
	 
	 Fn_SISW_NX_SaveAsNonMasterParts_Operation=False
	 
	  If ObjSave.Exist(2)=False Then
	 	bReturn=Fn_SISW_NX_General_MenuOperation("Select","Save As Non-master Parts","")
	 	Call Fn_SISW_NX_Setup_ReadyStatusSync(2)
	 	If bReturn=False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to open  [ New Item ] dialog.")
			Set ObjNewItem=Nothing
			Exit Function
		End If
	End If
	
	If ObjSave.Exist(2) = False Then
			Call Fn_UpdateLogFiles("FAIL : [ Save As Non Master Parts] dialog does Not  Exist", "FAIL: [ Save As Non Master Parts] dialog does Not  Exist")
			Set ObjOutput=Nothing
			Exit Function
	End If
	
	Select Case sAction
		Case "SaveAs"
				If dicSaveAsInfo("Number")<>"" Then
					bReturn=Fn_SISW_NX_UI_EditBoxOperation("Fn_SISW_NX_SaveAsNonMasterParts_Operation","Set",ObjSave,"Number", dicSaveAsInfo("Number"))
					If bReturn=False Then
						Call Fn_UpdateLogFiles("FAIL : Fail to set- "&dicSaveAsInfo("Number"), "FAIL: Fail to set - "&dicSaveAsInfo("Number"))
						Fn_SISW_NX_SaveAsNonMasterParts_Operation=false
						Set ObjSave=Nothing
						Set ObjFolderSelect=Nothing
						Exit Function
					Else
						wait 1
						Fn_SISW_NX_SaveAsNonMasterParts_Operation=True
					End If
				End If
				
				If dicSaveAsInfo("Revision")<>"" Then
					bReturn=Fn_SISW_NX_UI_EditBoxOperation("Fn_SISW_NX_SaveAsNonMasterParts_Operation","Set",ObjSave,"Revision", dicSaveAsInfo("Revision"))
					If bReturn=False Then
						Call Fn_UpdateLogFiles("FAIL : Fail to set- "&dicSaveAsInfo("Revision"), "FAIL: Fail to set - "&dicSaveAsInfo("Revision"))
						Fn_SISW_NX_SaveAsNonMasterParts_Operation=false
						Set ObjSave=Nothing
						Set ObjFolderSelect=Nothing
						Exit Function
					Else
						wait 1
						Fn_SISW_NX_SaveAsNonMasterParts_Operation=True
					End If
				End If
				
				If dicSaveAsInfo("ItemName")<>"" Then
					bReturn=Fn_SISW_NX_UI_EditBoxOperation("Fn_SISW_NX_SaveAsNonMasterParts_Operation","Set",ObjSave,"ItemName", dicSaveAsInfo("ItemName"))
					If bReturn=False Then
						Call Fn_UpdateLogFiles("FAIL : Fail to set- "&dicSaveAsInfo("ItemName"), "FAIL: Fail to set - "&dicSaveAsInfo("ItemName"))
						Fn_SISW_NX_SaveAsNonMasterParts_Operation=false
						Set ObjSave=Nothing
						Set ObjFolderSelect=Nothing
						Exit Function
					Else
						wait 2
						bReturn=Fn_SISW_NX_UI_EditBoxOperation("Fn_SISW_NX_SaveAsNonMasterParts_Operation","GetValue",ObjSave,"Revision", dicSaveAsInfo("ItemName"))
						If trim(bReturn)<> trim(dicSaveAsInfo("ItemName"))Then
							Call Fn_SISW_NX_UI_EditBoxOperation("Fn_SISW_NX_SaveAsNonMasterParts_Operation","Set",ObjSave,"ItemName", dicSaveAsInfo("ItemName"))
							wait 2
						End If
						Fn_SISW_NX_SaveAsNonMasterParts_Operation=True
					End If
				End If
				
				If dicSaveAsInfo("Folder")="Newstuff" Then
						bReturn = Fn_SISW_NX_UI_ButtonOperation("Fn_SISW_NX_SaveAsNonMasterParts_Operation","Click",ObjSave,"FolderButton")
						If bReturn=False Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Click on [ FolderButton ] button of [ Save As Non Master Parts ] Dialog")
							Fn_SISW_NX_SaveAsNonMasterParts_Operation=False
							Set ObjSave=Nothing
							Set ObjFolderSelect=Nothing
							Exit Function
						End If
					
						Call Fn_SISW_NX_UI_ComboBoxOperation("Set",ObjFolderSelect,"LookIn","","Home","")
						wait 2
						
						'Select Newstuff folder
						Call ObjFolderSelect.WinObject("NodeList").GetTextLocation("Newstuff",L, T, R, B,True)
                                  ObjFolderSelect.WinObject("NodeList").DblClick L,T,micLeftBtn
						wait 2
						
						bReturn=Fn_SISW_NX_UI_ButtonOperation("", "Click", ObjFolderSelect, "OK")
						If bReturn=False Then
							Call Fn_UpdateLogFiles("FAIL : Fail to click on  [ Ok ] button", "FAIL: Fail to click on  [ Ok ] button")
							Fn_SISW_NX_SaveAsNonMasterParts_Operation=false
							Set ObjSave=Nothing
							Set ObjFolderSelect=Nothing
							Exit Function
						Else
							Fn_SISW_NX_SaveAsNonMasterParts_Operation=True
							wait 1
						End If
					End If 
			End Select
				
			If sButton <> "" Then
				bReturn = Fn_SISW_NX_UI_ButtonOperation("Fn_SISW_NX_SaveAsNonMasterParts_Operation","Click",ObjSave,sButton)
				If bReturn=False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Click on ["+sButton+"] button in [ Save As Non Master Parts ] Dialog")
					Fn_SISW_NX_SaveAsNonMasterParts_Operation=False
					Set ObjSave=Nothing
					Set ObjFolderSelect=Nothing
					Exit Function
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Successfully Clicked on  ["+sButton+"] button of [ Save As Non Master Parts ] Dialog")
				End If
			End If
End Function
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Function Name     		: 	Fn_SISW_NX_AssemblyLoadOptions()
'
' Function Description 	    : 	Function used to perform operations on Assembly Load Options dialog in NX
'
' Parameters				: 	sAction: Action
'								dicDetails: Assembly Load Options details
'								sButton : Button name
'
' Return Value           	:  True/False
'	
' Examples		    		: Set dicDetails = CreateObject("Scripting.Dictionary")
'								  dicDetails("ConfigDetailsComboBox") = "As Saved"	
'							 bReturn = Fn_SISW_NX_ClosePartOperations("SetConfigurationDetails",dicDetails,"OK")
'
' History               	:  Developer Name			 Date	  			Rev. No. 				Changes					Reviewer
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'							 Poonam Chopade	 	    19-April-2018			1.0					   Created			TC11.5(20180402.00)_DIPRO_NX_NewDevelopment					
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_NX_AssemblyLoadOptions(sAction,dicDetails,sButton)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_NX_AssemblyLoadOptions"
    Dim objAssmLoadOptionsWindow,bFlag
   
   Set objAssmLoadOptionsWindow = Window("NXWindow").Dialog("AssemblyLoadOptions")
   Fn_SISW_NX_AssemblyLoadOptions=False
   
   'Check Existence of Window
   If Fn_SISW_UI_Object_Operations("Fn_SISW_NX_AssemblyLoadOptions", "Exist", objAssmLoadOptionsWindow, 2) = False Then
		'Call to Invoke Close Part Dialog
		Call Fn_SISW_NX_General_MenuOperation("Select","Assembly Load Options","")
		Wait 4
		'Check Existence
		If Fn_SISW_UI_Object_Operations("Fn_SISW_NX_AssemblyLoadOptions", "Exist", objAssmLoadOptionsWindow, 2) = False Then
			Fn_SISW_NX_AssemblyLoadOptions = False
			Set objAssmLoadOptionsWindow = Nothing
			Exit Function
		End If	
	End If
	
    Select Case sAction
   		Case "SetConfigurationDetails"
			
			If dicDetails("ConfigDetailsComboBox") <> "" Then	'Set Configuration Details combo Box  
				bFlag = Fn_SISW_NX_UI_ComboBoxOperation("Set",objAssmLoadOptionsWindow,"ConfigurationDetails","",dicDetails("ConfigDetailsComboBox"),"")
				If bFlag = False Then
					Fn_SISW_NX_AssemblyLoadOptions = False
					Set objAssmLoadOptionsWindow = Nothing
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed select Configuration Details ComboBox value [ "&dicDetails("ConfigDetailsComboBox")&" ].")
					Exit Function
				End If
			End If
   	End Select
    
	'Click OK/Cancel button
   	If sButton <>"" Then
		Call Fn_SISW_NX_UI_ButtonOperation("Fn_SISW_NX_AssemblyLoadOptions","Click",objAssmLoadOptionsWindow,sButton)
		Wait 1
   	End If
	  
   Fn_SISW_NX_AssemblyLoadOptions = True	   	 
   Set objAssmLoadOptionsWindow=Nothing
End Function
