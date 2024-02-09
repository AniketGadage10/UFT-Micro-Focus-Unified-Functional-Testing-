Dim sUIFail
' global variables for recursive function Fn_UI_getTreeIndex - Koustubh
Private Fn_UI_getTreeIndex_CompareString, Fn_UI_getTreeIndex_iGblCnt, Fn_UI_getTreeIndex_bFound
Public objNodeBounds
Public SISW_DEFAULT_TIMEOUT, SISW_MICRO_TIMEOUT, SISW_MIN_TIMEOUT, SISW_MAX_TIMEOUT

SISW_MAX_TIMEOUT = 240 'time in seconds
SISW_DEFAULT_TIMEOUT = 10 'time in seconds
SISW_MIN_TIMEOUT = 5 'time in seconds
SISW_MINLESS_TIMEOUT = 3
SISW_MICROLESS_TIMEOUT = 2
SISW_MICRO_TIMEOUT = 1 'time in seconds

'*********************************************************	UI Function List		***********************************************************************
'1	Fn_List_Click									- Depricated Function :	Note: Use Fn_SISW_UI_JavaList_Operations
'2	Fn_Menu_Select													                                            Actions performed in this function are:1. Menu Select                  2. Sub menu select
'3	Fn_CheckBox_Select								- Depricated Function :	Note: Use Fn_SISW_UI_JavaCheckBox_Operations
'4	Fn_Table_Select_Cell										                                          Select cell of the table
'5	Fn_List_Select									- Depricated Function :	Note: Use Fn_SISW_UI_JavaList_Operations
'6	Fn_MenuItem_Select																					   Actions performed in this function are    1. Menu Select            2. Menu multi-select"
'7	Fn_Window_Maximize									  												Actions performed in this function are:1. Maximizing Window
'8	Fn_ToolBar_ShowDropdown							- Depricated Function :	Note: Use Fn_SISW_UI_JavaToolbar_Operations
'9	Fn_JavaTree_Select																					  This Function is used to select Node/Element from the JavaTree.
'10	Fn_List_Activate								- Depricated Function :	Note: Use Fn_SISW_UI_JavaList_Operations
'11	Fn_CheckBox_Set									- Depricated Function :	Note: Use Fn_SISW_UI_JavaCheckBox_Operations
'12	Fn_Edit_Box										- Depricated Function :	Note: Use Fn_SISW_UI_JavaEdit_Operations
'13	Fn_Button_Click 								- Depricated Function :	Note: Use Fn_SISW_UI_JavaButton_Operations
'14	Fn_UI_JavaTable_SelectRow																   Select row of the table
'15	Fn_UI_JavaToolbar_Press							- Depricated Function :	Note: Use Fn_SISW_UI_JavaToolbar_Operations
'16	Fn_UI_JavaWindow_Minimize																   This function  is used to Minimize Java Window
'17	Fn_UI_JavaRadioButtont_setOff					- Depricated Function :	Note: Use Fn_SISW_UI_JavaRadioButton_Operations
'18	Fn_UI_JavaWindow_Presskey											                     This function  is used to set the press key function
'19	Fn_UI_JavaStaticText_SetTOProperty													This function  is used to set property to static text object
'20	Fn_UI_JavaRadioButton_SetON						- Depricated Function :	Note: Use Fn_SISW_UI_JavaRadioButton_Operations
'21	Fn_UI_JavaTable_SelectCell																	This function  is used to select the cell from the javatable
'22	Fn_UI_JavaTable_GetCellData																  Read Data of  The Selected Cell
'23	Fn_UI_JavaTable_ClickCell																	  This function is used to Click the Cell in the given Table.
'24	Fn_UI_JavaTable_SetCellData	 															 Actions performed in this function are: 1. Set the value for a cell of the table
'25	Fn_UI_JavaTab_Select							- Depricated Function :	Note: Use Fn_SISW_UI_JavaTab_Operations
'26	Fn_UI_JavaTree_Activate_Select_Node											  This function  is used to Activate and Select the Node in JavaTree
'27	Fn_UI_JavaList_ExtendSelect						- Depricated Function :	Note: Use Fn_SISW_UI_JavaList_Operations
'28	Fn_UI_JavaTree_ExtendSelect	 															Select multiple nodes of a tree.
'29	Fn_UI_JavaObject_Click																			This function is used to click Java Object
'30	Fn_UI_JavaDialog_ChildObjects														   Display JavaChild Objects Name
'31	Fn_UI_JavaTree_Expand																	     Function is used to Expand Node/Element from the JavaTree.
'32	Fn_UI_WinButton_Click																			   Function to verify WinButton is enabled and to Click the mouse button at X,Y Co-ordinates. 
'33	Fn_UI_JavaTree_Collapse																		 Function is used to Collapse Node/Element from the JavaTree.
'34	Fn_UI_JavaTree_OpenContextMenu													 This function is used open the context menu.
'35	Fn_UI_JavaStaticText_Click																	  This function  is used to click item on java static text
'36	Fn_UI_JavaTable_DoubleClickCell	                                                    Function is used to Double-click the specified cell in the table.
'37	Fn_UI_JavaTable_ExtendRow	          													Actions performed in this function are: 1.  Extend the Table upto iRows.
'38	Fn_UI_WinMenu_BuildMenuPath															  This function  is used to reach at the given menu item
'39	Fn_UI_ObjectCreate								- Depricated Function :	Note: Use Fn_SISW_UI_Object_Operations
'40	Fn_UI_ObjectExist								- Depricated Function :	Note: Use Fn_SISW_UI_Object_Operations
'41	Fn_Edit_Box_GetValue	 						- Depricated Function :	Note: Use Fn_SISW_UI_JavaEdit_Operations
'42	Fn_Java_StaticText_Exist						- Depricated Function :	Note: Use Fn_SISW_UI_Object_Operations, special function for StaticText is not required.
'43	Fn_JavaTable_Type																				  This function  is used to click item on java static text
'44	Fn_UI_Object_GetROProperty																 This function  is used to click item on java static text
'45	Fn_UI_Object_SetTOProperty																  This function  is used to click item on java static text
'46	Fn_Table_GetRowCount																		  This function  is used to click item on java static text
'47	Fn_UI_JavaStaticText_SearchAndClick												This function  is used to set property to static text object
'48	Fn_UI_JavaMenu_Select																			Actions performed in this function are:       1. Menu Select                          2. Menu multi-select"
'49	Fn_UI_SetDateAndTime																			Select row of the table
'50	Fn_UI_JavaMenu_SearchAndSelect													This function  is used to Search and Click The Static Text
'51	 Fn_UI_ObjectPressKey																			This function  is used PressKey Method
'52	 Fn_UI_EditBox_Type	 							- Depricated Function :	Note: Use Fn_SISW_UI_JavaEdit_Operations
'53	Fn_JavaTree_Node_Activate																This Function is used to Activate Node/Element from the JavaTree.
'54	Fn_UI_WebTable_GetCellData															 Read Data of  The Selected Cell From Web Table
'55	Fn_UI_JavaTable_RightClickCell														 This function is used to right click on JavaTable
'57	Fn_JavaTree_NodeIndex																	   This Function is used to Retrieve The index of Node
'58	Fn_ExitFromUI																							 This function  is used to  exit from test case after doing kill  active process mentioned in parameter .
'59 Fn_UI_Object_SetTOProperty_ExistCheck									  This function  is used to Set TO PropertyFor given Object BEFORE cheking existance of  given Object
'60 Fn_UI_ListItemExist								- Depricated Function :	Note: Use Fn_SISW_UI_JavaList_Operations
'61 Fn_JavaTree_NodeIndexExt()
'62 Fn_UI_JavaMenu_Exist()
'63 Fn_UI_JavaTable_CheckColumnExists()											This function checks wheather the given column name exist in the table.
'64 Fn_UI_JavaTable_ClickColumnHeader()											This function performs Right & Left click on column header of any java table
'65 Fn_UI_JavaTree_NodeExist																This function  is used to  check the Existance of the node in Tree
'66 Fn_UI_SwfButtonClick														This function is used to click the Swf Button.
'67 Fn_UI_WpfButtonClick														This function is used to click the Wpf Button.
'68 Fn_UI_TableOperations
'69 Fn_UI_JavaTreeGetItemPath
'70 Fn_UI_getTreeIndex()
'71 Fn_UI_getJavaTreeIndex()
'72 Fn_UI_JavaTreeGetItemPathExt()
'73 Fn_UI_ClickJavaTreeCell()
'74 Fn_UI_JavaStatictextOperations()
'75. Fn_SISW_UI_Spin_Edit
'76. Fn_SISW_UI_JavaTree_GetSanitizedNodeName()
'77. Fn_SISW_UI_RACTabFolderWidget_Operation()
'78. Fn_SISW_UI_JavaTableGetCellData()
'79. Fn_SISW_UI_GetRealPropertyName()
'80. Fn_SISW_UI_GetDisplayedRelation()
'81. Fn_SISW_UI_Twistie_Operations()
'82. Fn_SISW_UI_Object_Operations
'83. Fn_SISW_UI_CustomComboBox_Operations()	 						- Depricated Function :	Note: Use Fn_SISW_UI_CustomizedComboBox_Operations
'84. Fn_SISW_UI_Object_GetChildObjects()
'85. Fn_SISW_UI_JavaTab_Operations()
'86. Fn_SISW_UI_JavaButton_Operations()
'87. Fn_SISW_UI_JavaList_Operations()
'88. Fn_SISW_UI_JavaEdit_Operations()
'89. Fn_SISW_UI_JavaCheckBox_Operations()
'90. Fn_SISW_UI_JavaRadioButton_Operations()
'91. Fn_SISW_UI_JavaToolbar_Operations()
'92. Fn_SISW_UI_JavaTable_Operations()
'93. Fn_SISW_UI_DeviceReplayObjectClick
'94. Fn_SISW_UI_CustomizedComboBox_Operations
'95. Fn_SISW_UI_InsightObject_Operations
'96. Fn_SISW_UI_WinListView_Operations()
'97. Fn_UI_SetDateAndTimeExt(sFunctionName,sDate,sTime,objDate,objTime)	:	This function is use to Set Date in edit box and Time in list
'98. Fn_UI_ResizeObject()
'99. Fn_UI_VerifyHorizontalVerticalBar()
'100. Fn_SISW_UI_WinButton_Operations()                                 : This function is used to perform operations on WinButton object.
'101. Fn_SISW_UI_WinEdit_Operations()									: This function is use to perform opeartions on WinEdit object.
'*********************************************************	UI Function List		***********************************************************************
'#################################################################################################################
'###    FUNCTION NAME   :    Fn_ExitFromUI(sProcessToKill)
'###
'###    DESCRIPTION     :   This function  is used to  exit from test case after doing kill  active process mentioned in parameter .
'###
'###    PARAMETERS      :   sProcessToKill - list of processes seperated by colon (:) 
'###             
'###    Function Calls  :   Fn_KillProcess(sProcessToKill)
'###
'###    HISTORY         :   
'###
'###    CREATED BY      :   Sagar Shivade
'###
'###    REVIWED BY      :      Sameer Chitnis
'###
'###    MODIFIED BY     :
'###    EXAMPLE         :    ExitFromUI("Teamcenter.exe:java.exe:jucheck.exe:jusched.exe:javaw.exe")
'###												ExitFromUI("")
'################################################################################################################
Function ExitFromUI(sMsg)
	' call function killProcess .
	Call Fn_KillProcess("")
	'exit from current test.
	'Code added by Archana
	Call Fn_UpdateLogFiles(Environment.Value("ActionName") & " > " & " FAIL | "+ sMsg, "FAIL:"+ sMsg)
	ExitTestIteration 
End Function
'#################################################################################################################
'###    FUNCTION NAME   :    Fn_List_Click(sFunctionName,objJavaDialog,sJavaList,iRow,iCol)
'###    Note : Depricated Function :	Note: Use Fn_SISW_UI_JavaList_Operations
'###    EXAMPLE         :   call Fn_List_Click("Fn_OptionSettings_ProdStrc", "Options", "DefaultViewType", 1, 1)
'################################################################################################################
Function Fn_List_Click(sFunctionName, objJavaDialog, sJavaList, iRow, iCol)
	bGblFailedFunctionName = sFunctionName
	'log on success
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Depricated function Fn_List_Click : Calling function Fn_SISW_UI_JavaList_Operations")
	Fn_List_Click = Fn_SISW_UI_JavaList_Operations(sFunctionName, "Click", objJavaDialog, sJavaList, iRow, iCol, "")
End Function
'#####################################################################################################################################################
'##Function Name	:	Fn_Menu_Select( )
'##
'##Description	:	Actions performed in this function are:1. Menu Select
'##							       2. Sub menu select
'##				
'##
'##Parameters	:	1. sFunctionName: Valid Function Name
'##			2. objJavaWindow: Valid Java Dialog 
'##			3. sWinMenu: Valid menu item name
'##
'##
'##Return Value	: 	TRUE \ FALSE
'##
'##HISTORY       :  	AUTHOR                  DATE         VERSION
'##
'##CREATED BY    :     Sandeep & Amol          16/04/2010      1.0
'##											   28/04/2010	   Upadated
'##REVIWED BY    :     Rajesh 	              16/04/2010      2.0           
'##
'##Examples	:    Fn_Menu_Select(Fn_QuickSearch(sSrchType, sSrchText),JavaWindow("DefaultWindow").WinMenu,"ContextMenu","File;New")
'## 
'####################################################################################################################################################
Function Fn_Menu_Select(sFunctionName,objJavaDialog,sWinMenu,sWinMenuPath)
	Dim objWinMenu
	sUIFail = sFunctionName + ">> Fn_Menu_Select >> " +  objJavaDialog.toString +">> " +  sWinMenu
 	bGblFailedFunctionName = sFunctionName
		'Setting the Object 
		Set objWinMenu = objJavaDialog.WinMenu(sWinMenu)

		'Check  existence of object
		If objWinMenu.Exist Then

		'Apply wait property on Object
		objWinMenu.WaitProperty "enabled", "1"

			'Checking wait property Enable or not
			If objWinMenu.GetROProperty("enabled") = "1"  Then
				' objWinMenu.Select
				'sWinMenuPath is "Colon (:)"Seperated
				objWinMenu.Select(sWinMenuPath)
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Sucessfully selected SubMenu " & sWinMenuPath &" of WinMenu " &sWinMenu & " of Function" &sFunctionName)
				Fn_Menu_Select=true
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sWinMenuPath & " is disable" &" of WinMenu " &sWinMenu &" in Function " &sFunctionName)
				Fn_Menu_Select=False				
				Call ExitFromUI(sUIFail)
			End If
		Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sWinMenuPath & " does not exist "&" of Function " &sFunctionName)
			Fn_Menu_Select=False
			Call ExitFromUI(sUIFail)
		End If

	 Set objWinMenu= Nothing 
End Function
'################################################################################
'#### FUNCTION NAME :Fn_CheckBox_Select(sFunctionName, objJavaDialog, sJavaCheckBox)
'#### EXAMPLE :  Fn_CheckBox_Select(Fn_ObjectCheckOut (sChngID, sComment, bExportDateSet, bOverwriteFiles), objNewSignalDialog.JavaCheckBox, "Configuration Item")
'#################################################################################
Function Fn_CheckBox_Select(sFunctionName, objJavaDialog, sJavaCheckBox)
	bGblFailedFunctionName = sFunctionName
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Depricated function Fn_CheckBox_Select : Calling function Fn_SISW_UI_JavaCheckBox_Operations")
	Fn_CheckBox_Select = Fn_SISW_UI_JavaCheckBox_Operations(sFunctionName, "Set", objJavaDialog, sJavaCheckBox, "ON")
End Function
'#################################################################################################################
'###    FUNCTION NAME   :   Fn_Table_Select_Cell(sFunctionName, objJavaDialog, sJavaTable,iRow,iColumn)
'###	Note : Depricated Function :	Note: Use Fn_SISW_UI_JavaTable_Operations
'###    EXAMPLE         :   call Fn_Table_Select_Cell("Fn_itemBasicCreate", ObjJavaDialog, JavaTable, 1,2)
'################################################################################################################
Function Fn_Table_Select_Cell(sFunctionName, objJavaDialog, sJavaTable,iRow,iColumn)
	bGblFailedFunctionName = sFunctionName
    Call Fn_WriteLogFile(Environment.Value("TestLogFile"),  "Depricated function : Fn_Table_Select_Cell > Calling Function Fn_SISW_UI_JavaTable_Operations")
	
	If isNumeric(iRow) Then 
		iRow = cInt(iRow)
	End If

	If isNumeric(iColumn) Then 
		iColumn = cInt(iColumn)
	End If
	
	Fn_Table_Select_Cell = Fn_SISW_UI_JavaTable_Operations(sFunctionName, "SelectCell", objJavaDialog , sJavaTable, "", "", iRow, iColumn, "", "", "")
End Function
'##############################################################################################################################################
'###    FUNCTION NAME   :   Fn_List_Select()
'###	Note : Depricated Function :	Note: Use Fn_SISW_UI_JavaList_Operations
'###    EXAMPLE         :   Fn_List_Select("Fn_ItemBasicCreate",JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("New Item"),"ItemType","000110/A;1-Part1")
'###############################################################################################################################################
Function Fn_List_Select(sFunctionName, objJavaDialog, sJavaList,sElementToSelect)
	bGblFailedFunctionName = sFunctionName
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"),  "Depricated function : Fn_List_Select > Calling Function Fn_SISW_UI_JavaList_Operations ")
	Fn_List_Select = Fn_SISW_UI_JavaList_Operations(sFunctionName, "Select", objJavaDialog, sJavaList, sElementToSelect, "", "")
End Function
'###################################################################################################################################################################
'Function Name	   :	Fn_MenuItem_Select( )
'
'Description	   :	 Actions performed in this function are:
'	   		1. Menu Select
'			2. Menu multi-select
'				
'
'Parameters	   :	    1. sFunctionName: Valid Function Name
'		 	    2. objJavaDialog: Valid Java Dialog 
'			    3. sJavaMenu: Valid menu  name
'			    4.sJavaMenuItem: Valid menu item name
'
'Return Value	   : 	TRUE \ FALSE
'
'HISTORY       	   :  	            AUTHOR                      DATE    	    		 VERSION
'
'CREATED BY        :           Pranav  Shirode               19/04/2010  		   	   1.0
'
'REVIWED BY    	   :         Rajesh							         
'"
'Examples	   :    Fn_MenuItem_Select("Fn_MenuItem_select_Operation",JavaWindow("StructureManager").JavaMenu("Text menu")."File","New" )
' 											
'#################################################################################################################################################################
Function Fn_MenuItem_Select(sFunctionName,objJavaDialog,sJavaMenu,sJavaMenuItem)

Dim objJavaMenuItem
	bGblFailedFunctionName = sFunctionName
 sUIFail = sFunctionName + ">> Fn_Menu_Select >> " +  objJavaDialog.toString +">> " +  sJavaMenu

'Setting the Object 
		Set objJavaMenuItem = objJavaDialog.JavaMenu(sJavaMenu).JavaMenuItem(sJavaMenuItem)
'Check  existence of object
		If objJavaMenuItem.Exist Then
'Apply wait property on Object
			objJavaMenuItem.WaitProperty "enabled", "1"
'Checking wait property Enable or not
				If objJavaMenuItem.GetROProperty("enabled") = "1"  Then

						objJavaMenuItem.Select(sJavaMenuItem)
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Selected element  " & sJavaMenuItem &"of Menu " &sJavaMenu & " of Function" &sFunctionName)
						Fn_Menu_Select=true
				Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Selected element  " & sJavaMenuItem &"of Menu " &sJavaMenu &" is Disable " &"of Function " &sFunctionName)
						Fn_Menu_Select=False				
						Call ExitFromUI(sUIFail)
				End If
		Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Selected element  " & sJavaMenuItem &"of Menu " &sJavaMenu &" doest not exist "&"of Function " &sFunctionName)
					Fn_Menu_Select=False
					Call ExitFromUI(sUIFail)
		End If

Set objJavaMenuItem=Nothing
End Function 
'#####################################################################################################################################################
'##Function Name	:	Fn_Window_Maximize( )
'##
'##Description	:	Actions performed in this function are:1. Maximizing Window
'##				
'##
'##Parameters	:	1. sFunctionName: Valid Function Name
'##			2. objJavaWindow: Valid Java Dialog 
'##			
'##
'##Return Value	: 	TRUE \ FALSE
'##
'##HISTORY       :  	AUTHOR                  DATE         VERSION
'##
'##CREATED BY    :      Sandeep                 19/04/2010      1.0(2.0 on 20/04/2010)
'##
'##REVIWED BY    :  	
'##
'##Examples	:    Fn_Window_Maximize("Fn_Create_New_Opt()",JavaDialog("Create new option"))
'## 
'####################################################################################################################################################

Function Fn_Window_Maximize(sFunctionName,objJavaDialog)

	Dim objWin
	bGblFailedFunctionName = sFunctionName
		sUIFail = sFunctionName + ">> Fn_Window_Maximize >> " +  objJavaDialog.toString
	
	'Setting Object
	Set objWin=objJavaDialog
	'Check Existance Of Object

	If objWin.exist Then
'Apply wait property on Object
	objWin.WaitProperty  "enabled", "1"
'Checking wait property Enable or not
			If objWin.GetROProperty("enabled")="1" or objWin.GetROProperty("enabled")=true Then
'maximizing objWin
'Check Whether the window is maximizable 
				If objWin.GetROProperty("Maximizable")=True Then
'If Maximizable then checking Window Is already Maximized or not
					If objWin.GetROProperty("Maximized")=True Then
						'objWin.Activate
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Sucessfully Maximized the Window/Dialog "  & objJavaDialog.toString & "of Function" &sFunctionName)
						Fn_Window_Maximize=True
          			Else
						objWin.maximize
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Sucessfully Maximized the Window/Dialog "  & objJavaDialog.toString & "of Function" &sFunctionName)
						Fn_Window_Maximize=True
                	End If
				Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), objJavaDialog.toString & " Window/Dialog is not Maximizable of Function" &sFunctionName)
						Fn_Window_Maximize=False
						Call ExitFromUI(sUIFail)
				End If
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Window/Dialog " & objJavaDialog.toString &" is Disable  of Function" &sFunctionName)
				Fn_Window_Maximize=False
				Call ExitFromUI(sUIFail)
			End If
    Else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile")," Window/Dialog" & objJavaDialog.toString & "does not Exist  of Function" &sFunctionName)
		Fn_Window_Maximize=False
		Call ExitFromUI(sUIFail)
   End If
Set objWin = Nothing
End Function
'#####################################################################################################################################################
'##Function Name	:	Fn_ToolBar_ShowDropdown( )
'##
'###	Note : Depricated Function :	Note: Use Fn_SISW_UI_JavaToolbar_Operations
'####################################################################################################################################################
Public Function Fn_ToolBar_ShowDropdown(sFunctionName,objJavaDialog,sJavaToolBar,sIndexName)
	bGblFailedFunctionName = sFunctionName
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Depricated function Fn_ToolBar_ShowDropdown : Calling function Fn_SISW_UI_JavaToolbar_Operations")
	Fn_ToolBar_ShowDropdown = Fn_SISW_UI_JavaToolbar_Operations(sFunctionName, "OpenDropdownMenu", objJavaDialog, sJavaToolBar, sIndexName, "", "", "")
End Function
'############################################################################################################################################################
'###    FUNCTION NAME   :   Fn_JavaTree_Select()
'###
'###    DESCRIPTION     :   This Function is used to select Node/Element from the JavaTree.
'###
'###    PARAMETERS      :   sFunctionName	: Valid Function name,
'### 			    objJavaDialog	: Valid Dialog Path,
'### 			    sJavaTree		: Valid Javatree Name,
'### 			    sElementToSelect 	: Valid Note/Element to be selected ( Element seperated by :)
'###
'###    Function Calls  :   Fn_WriteLogFile()
'###
'###    HISTORY         :   AUTHOR      DATE        VERSION
'###	
'###    CREATED BY      :   Manisha    	19/04/2010   1.0
'###
'###    REVIWED BY      :   Rajesh	19/04/2010   1.0	
'###
'###    MODIFIED BY     :   NA
'###    EXAMPLE         :   Fn_JavaTree_Select("Fn_OptionSettings_ProdStrc", JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Options"), "OptionsTree","Options:Product Structure")
'###############################################################################################################################################################	
Function Fn_JavaTree_Select(sFunctionName, objJavaDialog, sJavaTree,sElementToSelect)
	Dim objJavaTree
	bGblFailedFunctionName = sFunctionName
  	sUIFail = sFunctionName + ">> Fn_JavaTree_Select >> " +  objJavaDialog.toString +">> " +  sJavaTree

'Object Creation
	Set objJavaTree = objJavaDialog.JavaTree(sJavaTree)

' Verify  JavaTree object exists
   If objJavaTree.Exist Then

' Synchronization Point for an Java Tree Object 
	objJavaTree.WaitProperty "enabled","1"

' Verify JavaTree Object is enabled 
	If objJavaTree.GetROProperty("enabled") = "1"  Then

' Select the element/Node from JavaTree.
           objJavaTree.Select sElementToSelect

'  Report message when selected the element from the tree sucessfully.
           Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Selected Element/Node  " & sElementToSelect &" from JavaTree" &sjavaTree &" of Function " &sFunctionName)
		   Fn_JavaTree_Select = True

' Report error when JavaTree object is disable.
	Else
   	   Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sjavaTree & "Tree is Disable  of Function " &sFunctionName)
  	   Fn_JavaTree_Select = False
	   Call ExitFromUI(sUIFail)
	End If

' Report error when JavaTree object does not exists.
	Else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "JavaTree " & sJavaTree & " does not exists of Function " &sFunctionName)
		Fn_JavaTree_Select = False
		Call ExitFromUI(sUIFail)
	End If
	Set objJavaTree=Nothing
' End of function 
End Function
'#####################################################################################################################################################
'##Function Name:	Fn_List_Activate( )
'## 	Note: Use Fn_SISW_UI_JavaList_Operations
'##Examples	:    Fn_List_Activate("Fn_MyTc_GeneralSearch",objWindow.JavaWindow("Type:").JavaWindow("PopupWindow"),"Value","hotdog") 
'####################################################################################################################################################
Public function Fn_List_Activate(sFunctionName,objJavaDialog,sListName,sItemName)
	bGblFailedFunctionName = sFunctionName
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Depricated function : Fn_List_Activate > Calling function Fn_SISW_UI_JavaList_Operations")
	Fn_List_Activate = Fn_SISW_UI_JavaList_Operations(sFunctionName, "Activate", objJavaDialog, sListName, sItemName, "", "")
End Function
'##########################################################################################################################################################################
'###    FUNCTION NAME   :   Fn_CheckBox_Set(sFunctionName, objJavaDialog, sJavaCheckBox, sStatus)
'###    Depricated Function :	Note: Use Fn_SISW_UI_JavaCheckBox_Operations
'###########################################################################################################################################################################
Function Fn_CheckBox_Set(sFunctionName, objJavaDialog, sJavaCheckBox, sStatus)
	bGblFailedFunctionName = sFunctionName
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Depricated function : Fn_CheckBox_Set > Calling Function Fn_SISW_UI_JavaCheckBox_Operations")
	Fn_CheckBox_Set = Fn_SISW_UI_JavaCheckBox_Operations(sFunctionName, "Set", objJavaDialog, sJavaCheckBox, sStatus)
End Function
'#########################################################################################################
'###    FUNCTION NAME   :   Function Fn_Edit_Box(sFunctionName,objJavaDialog,sJavaEdit,sText)
'###    Depricated Function :	Note: Use Fn_SISW_UI_JavaEdit_Operations
'###    EXAMPLE         : Fn_Edit_Box("Fn_TeamcenterLogin",JavaWindow("Teamcenter Login"),"User ID:","admin")
'#############################################################################################################
Function Fn_Edit_Box(sFunctionName,objJavaDialog,sJavaEdit,sText)
	bGblFailedFunctionName = sFunctionName
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Depricated function : Fn_Edit_Box > Calling Function Fn_SISW_UI_JavaEdit_Operations")
	Fn_Edit_Box = Fn_SISW_UI_JavaEdit_Operations(sFunctionName, "Set", objJavaDialog, sJavaEdit, sText)
End Function
'############################################################################################################
'###    FUNCTION NAME   :   Fn_Button_Click() 
'###    Depricated Function :	Note: Use Fn_SISW_UI_JavaButton_Operations
'#######################################################################################################
Function Fn_Button_Click(sFunctionName, objJavaDialog, sJavaButton)
	bGblFailedFunctionName = sFunctionName
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Depricated function : Fn_Button_Click > Calling Function Fn_SISW_UI_JavaButton_Operations")
	Fn_Button_Click =  Fn_SISW_UI_JavaButton_Operations(sFunctionName, "Fn_Button_Click", objJavaDialog, sJavaButton)
	If Fn_Button_Click = False Then
		Call ExitFromUI("Failed to Click on " & sJavaButton )
	End If
End Function
'#######################################################################################################################
'###    FUNCTION NAME   :   Fn_UI_JavaTable_SelectRow(sFunctionName, objJavaDialog, sJavaTableName,iRow)
'###	Note : Depricated Function :	Note: Use Fn_SISW_UI_JavaTable_Operations
'###
'###    EXAMPLE         :    Call Fn_UI_JavaTable_SelectRow("Fn_itemBasicCreate", objEffectivity1,"EffectivityTable",1)
'########################################################################################################################
Function Fn_UI_JavaTable_SelectRow(sFunctionName, objJavaDialog, sJavaTableName,iRow)
	bGblFailedFunctionName = sFunctionName
   If isNumeric(iRow) Then
	   iRow = cInt(iRow)
   End If
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"),  "Depricated function : Fn_UI_JavaTable_SelectRow > Calling Function Fn_SISW_UI_JavaTable_Operations")
	Fn_UI_JavaTable_SelectRow = Fn_SISW_UI_JavaTable_Operations(sFunctionName, "SelectRow", objJavaDialog , sJavaTableName, "", "", iRow, "", "", "", "")
End Function
'#######################################################################################################################
'###    FUNCTION NAME   :   Fn_UI_JavaToolbar_Press
'###    Depricated Function :	Note: Use Fn_SISW_UI_JavaToolbar_Operations
'########################################################################################################################
Public Function Fn_UI_JavaToolbar_Press(sFunctionName, objJavaDialog, sJavaToolbarName, sbuttonName)
	bGblFailedFunctionName = sFunctionName
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Depricated function Fn_UI_JavaToolbar_Press > Calling function Fn_SISW_UI_JavaToolbar_Operations")
	Fn_UI_JavaToolbar_Press = Fn_SISW_UI_JavaToolbar_Operations(sFunctionName, "Click", objJavaDialog, sJavaToolbarName, sbuttonName, "", "", "")
End Function
'################################################################################################################
'###    FUNCTION NAME   :   Fn_UI_JavaWindow_Minimize(sFunctionName,objJavaDialog)
'###
'###    DESCRIPTION     :   This function  is used to Minimize Java Window
'###
'###    PARAMETERS      :   sFunctionName - Valid function name
'###                        objJavaDialog -  Valid Dialog/Window Path
'###                        
'###    Function Calls  :   Fn_WriteLogFile()
'###
'###    HISTORY         :   AUTHOR               DATE              VERSION
'###
'###    CREATED BY      :   Dhananjay      	20/04/2010           1.0
'###
'###    REVIWED BY      :   Rajesh			20/04/2010			 1.0  
'###
'###    MODIFIED BY     :
'###    EXAMPLE         :  Fn_UI_JavaWindow_Minimize("Fn_RunRCAFScript(sScriptPath)",JavaWindow("RCAF Console"))
'################################################################################################################

Function Fn_UI_JavaWindow_Minimize(sFunctionName,objJavaDialog)
Dim objWindow
	bGblFailedFunctionName = sFunctionName
    sUIFail = sFunctionName + ">> Fn_UI_JavaWindow_Minimize >> " +  objJavaDialog.toString

	'Create Object of JavaWindow
	Set objWindow=objJavaDialog

	'Check JavaWindow Object 
	If objWindow.Exist Then
		' Apply  Wait Property
		objWindow.WaitProperty "enabled","1"
					
		'Check if Window is enabled
			If   objWindow.GetROProperty("enabled")="1"  Then
			'Check If Window is Maximized
				If  objWindow.GetROProperty("maximized")=true Then
					'Activate JavaWindow
					objWindow.Activate
					'Check if Window Minimizable
						If  objWindow.GetROProperty("minimizable")=true Then
							'Minimize Window
							objWindow.minimize
							'Log the Result						
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Sucessfully Minimized the " & objJavaDialog.toString &"Window/Dialog " &sFunctionName)
							Fn_UI_JavaWindow_Minimize=true
						else
							'Log the Result
							Call Fn_WriteLogFile(Environment.Value("TestLogFile")," Not able to Minimize  " & objJavaDialog.toString &"Window/Dialog  of Function  " &sFunctionName)
							Fn_UI_JavaWindow_Minimize=false
							Call ExitFromUI(sUIFail)
				             End If
										
				Else
				'Log the Result
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), objJavaDialog.toString &" Window/Dialog is already Minimized of Function " &sFunctionName)
				Fn_UI_JavaWindow_Minimize=true
                End If
			Else
				'Log the Result
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), objJavaDialog.toString &" Window/Dialog is Disabled of Function " & sFunctionName)
				Fn_UI_JavaWindow_Minimize=false
				Call ExitFromUI(sUIFail)
   		 	End If

	Else 

	'Log the Result
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), objJavaDialog.toString &" Window/Dialog does not Exist  of Function " & sFunctionName)
	Fn_UI_JavaWindow_Minimize=false
	Call ExitFromUI(sUIFail)
	End If
	Set objWindow = Nothing 
End Function
'#################################################################################################################
'###    FUNCTION NAME   :   Fn_UI_JavaRadioButtont_setOff(sFunctionName,objJavaDialog,sJavaRadioButton)
'###
'###    EXAMPLE         :   call Fn_UI_JavaRadioButtont_setOff("Fn_itemBasicCreate", objEditProperties , "ConfigItem")
'################################################################################################################
Function Fn_UI_JavaRadioButtont_setOff(sFunctionName,objJavaDialog,sJavaRadioButton)
	bGblFailedFunctionName = sFunctionName
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Depricated function Fn_UI_JavaRadioButtont_setOff : Calling function Fn_SISW_UI_JavaRadioButton_Operations")
	Fn_UI_JavaRadioButtont_setOff = Fn_SISW_UI_JavaRadioButton_Operations(sFunctionName, "Set", objJavaDialog, sJavaRadioButton, "OFF")
End Function
'#################################################################################################################
'###    FUNCTION NAME   :    Fn_UI_JavaWindow_Presskey(fn_presskey,objJavaDialog,skey,scntrl)
'###
'###    DESCRIPTION     :   This function  is used to set the press key function
'###
'###    PARAMETERS      :   sFunctionName - Valid function name
'###                        objJavaDialog -  Valid Dialog/Window Path
'###                        skey          - Vaild Key  name entered by  the user
'###                        scntrl        - valid control characterpress with valid key
'###                        
'###    Function Calls  :   Fn_WriteLogFile ()
'###
'###    HISTORY         :   AUTHOR                        DATE                   VERSION
'###
'###    CREATED BY      :   Vidya and Rakesh              20/04/2010               1.0
'###
'###    REVIWED BY      :  Rajeev						 20/04/2010					1.0
'###
'###    MODIFIED BY     :   
'###  
'###   EXAMPLE          :   Fn_UI_JavaWindow_Presskey("Fn_itemBasic_Create",JavaDialog("New Item"),"F1","Ctrl")

'################################################################################################################

 Function Fn_UI_JavaWindow_Presskey(sFunctionName,objJavaDialog,skey,sCntrl)
	bGblFailedFunctionName = sFunctionName
    sUIFail = sFunctionName + ">> Fn_UI_JavaWindow_Presskey >> " +  objJavaDialog.toString +">> " +  skey
        
     'Set objPressKey = objJavaDialog
     If  objJavaDialog.Exist  Then
          objJavaDialog.WaitProperty "enabled", "1"

	       	If objJavaDialog.GetROProperty("enabled")="1" Then
				 If objJavaDialog.Activate Then
				 	If sCntrl <>"" Then
					 objJavaDialog.PressKey sKey, sCntrl
					 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Sucessfully pressed Key "& sCntrl &" " &skey &" of Window/Dialog " & objJavaDialog.toString & " of Function " &sFunctionName)
					 Else 
					  objJavaDialog.PressKey sKey 
					  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Sucessfully pressed Key " &skey &" of Window/Dialog " & objJavaDialog.toString & " of Function " &sFunctionName)
  			         End If
				 	'Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Sucessfully pressed "& skey & " Key of Function " &sFunctionName)
				 	Fn_UI_JavaWindow_Presskey = True
				
			  	Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),  objJavaDialog.toString &"Window/Dialog " &" is not Activated to perform PressKey operation of Function " &sFunctionName)
					Fn_UI_JavaWindow_Presskey = False
					Call ExitFromUI(sUIFail) 
			  End If 
		    Else
			  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), objJavaDialog.toString &"Window/Dialog is Disabled to perform PressKey operation of Function " &sFunctionName)
			  Fn_UI_JavaWindow_Presskey = False
			  Call ExitFromUI(sUIFail)
		    End If

     Else 
	   Call Fn_WriteLogFile(Environment.Value("TestLogFile"),  objJavaDialog.toString &"Window/Dialog does not Exist to perform PressKey operation of Function " &sFunctionName)
	   Fn_UI_JavaWindow_Presskey = False
	   Call ExitFromUI(sUIFail) 
     End If

End Function


'########################################################################################

'#################################################################################################################
'###    FUNCTION NAME   :     Fn_UI_JavaStaticText_SetTOProperty(sFunctionName, objJavaDialog,sStaticText , sProperty, sPropValue)
'###
'###    DESCRIPTION     :   This function  is used to set property to static text object
'###
'###    PARAMETERS      :   sFunctionName - Valid function name
'###                        objJavaDialog -  Valid Dialog/Window Path
'###                        sStaticText - Vaild static text name
'###                       sProperty  -  valid text
'###                       sPropValue - Valid Property value
'###                        
'###    Function Calls  :   Fn_WriteLogFile
'###
'###    HISTORY         :   AUTHOR            DATE        VERSION
'###
'###    CREATED BY      :   Deepak kumar   20/04/2010  		1.0
'###
'###    REVIWED BY      :    Rajesh			20/04/2010     1.0          
'###
'###    MODIFIED BY     :
'###    EXAMPLE         :  Fn_UI_JavaStaticText_SetTOProperty("Function Fn_ReadyStatusSync",JavaWindow("DefaultWindow"),"Ready","label","Ready")
'################################################################################################################
Function Fn_UI_JavaStaticText_SetTOProperty(sFunctionName, objJavaDialog,sJavaStaticText , sProperty, sPropValue) 

Dim objJavaStaticText
	'Set an object on variable 
  	bGblFailedFunctionName = sFunctionName
  sUIFail = sFunctionName + ">> Fn_UI_JavaStaticText_SetTOProperty >> " +  objJavaDialog.toString +">> " +  sJavaStaticText
         
	Set objJavaStaticText = objJavaDialog.JavaStaticText(sJavaStaticText)
                
        'checking StaticText object exist or not
	If objJavaStaticText.Exist Then  

				 objJavaStaticText.WaitProperty "displayed", "1"

		 'Syncronization point
                 If objJavaStaticText.GetROProperty("displayed")  = "1" Then

				 'set TOProperty  with the value
                         objJavaStaticText.SetTOProperty  sProperty, sPropValue
                                 
			 	'log the success
                          Call Fn_WriteLogFile( Environment.Value("TestLogFile"), sProperty &"Property is Set to  " &sPropValue &" value of  " & sJavaStaticText  &" JavaStaticText of Function " & sFunctionName)

			       'Return True from Function
                                Fn_UI_JavaStaticText_SetTOProperty= True

				Else

				 'log the failure when text not Displayed
                                 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sJavaStaticText  &" JavaStaticText is not Displayed of Function " & sFunctionName)
                                 Fn_UI_JavaStaticText_SetTOProperty= False
                                 Call ExitFromUI(sUIFail)

			 End If

		Else
                                     
			'log the failure when text not enabled
                        Call Fn_WriteLogFile(Environment.Value("TestLogFile"),  sJavaStaticText  &" JavaStaticText does not Exist of Function " & sFunctionName)
                 
		'Return False from function
                Fn_UI_JavaStaticText_SetTOProperty= False

	        Call ExitFromUI(sUIFail)

        End If
	Set objJavaStaticText = Nothing 
End Function
'################################################################################################################
'###    FUNCTION NAME   :    Fn_UI_JavaRadioButton_SetON(strFunctionName,objJavaDialog, sJavaRadioButton)
'###	Note : Depricated Function :	Note: Use Fn_SISW_UI_JavaRadioButton_Operations
'################################################################################################################
Function Fn_UI_JavaRadioButton_SetON(sFunctionName,objJavaDialog, sJavaRadioButton)
	'Log the Result		
	bGblFailedFunctionName = sFunctionName	
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Depricated function Fn_UI_JavaRadioButton_SetON : Calling function Fn_SISW_UI_JavaRadioButton_Operations")
	Fn_UI_JavaRadioButton_SetON = Fn_SISW_UI_JavaRadioButton_Operations(sFunctionName, "Set", objJavaDialog, sJavaRadioButton, "ON")
End Function
'################################################################################################################
'###    FUNCTION NAME   :    Fn_UI_JavaTable_SelectCell(sFunctionName,objJavaDialog,sJavaTable,iRow,varCol)  - Duplicate Function
'###	Note : Depricated Function :	Note: Use Fn_SISW_UI_JavaTable_Operations
'################################################################################################################
Function Fn_UI_JavaTable_SelectCell(sFunctionName, objJavaDialog, sJavaTable,iRow,varCol)
	bGblFailedFunctionName = sFunctionName
    Call Fn_WriteLogFile(Environment.Value("TestLogFile"),  "Depricated function : Fn_UI_JavaTable_SelectCell > Calling Function Fn_SISW_UI_JavaTable_Operations")
	If isNumeric(iRow) Then iRow = cInt(iRow)
	Fn_UI_JavaTable_SelectCell = Fn_SISW_UI_JavaTable_Operations(sFunctionName, "SelectCell", objJavaDialog , sJavaTable, "", "", iRow, varCol, "", "", "")
End Function
'################################################################################################################
'###    FUNCTION NAME   :   Fn_UI_JavaTable_GetCellData(sFunctionName, objJavaDialog, sJavaTable,sRow,sColumn)
'###	Note : Depricated Function :	Note: Use Fn_SISW_UI_JavaTable_Operations
'################################################################################################################
Function Fn_UI_JavaTable_GetCellData(sFunctionName, objJavaDialog, sJavaTable,sRow,sColumn)
	bGblFailedFunctionName = sFunctionName
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"),  "Depricated function : Fn_UI_JavaTable_GetCellData > Calling Function Fn_SISW_UI_JavaTable_Operations")
	Fn_UI_JavaTable_GetCellData = Fn_SISW_UI_JavaTable_Operations(sFunctionName, "GetCellData", objJavaDialog , sJavaTable, "", "", sRow, sColumn, "", "", "")
End Function
'################################################################################################################
'###    FUNCTION NAME   :   Fn_UI_JavaTable_ClickCell(sFunctionName,objJavaDialog,sJavaTable,sRow, sCol)
'###	Note : Depricated Function :	Note: Use Fn_SISW_UI_JavaTable_Operations
'################################################################################################################
Function Fn_UI_JavaTable_ClickCell(sFunctionName,objJavaDialog,sJavaTable,sRow, sCol)
	bGblFailedFunctionName = sFunctionName
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"),  "Depricated function : Fn_UI_JavaTable_ClickCell > Calling Function Fn_SISW_UI_JavaTable_Operations")
	Fn_UI_JavaTable_ClickCell = Fn_SISW_UI_JavaTable_Operations(sFunctionName, "ClickCell", objJavaDialog , sJavaTable, "", "", sRow, sCol, "", "", "")
End Function
'################################################################################################################
'###	Function Name	   	:	Fn_UI_JavaTable_SetCellData(sFunctionName,objJavaDialog,sJavaTable,iRow,iCol,sData)
'###	Note : Depricated Function :	Note: Use Fn_SISW_UI_JavaTable_Operations
'################################################################################################################
Function Fn_UI_JavaTable_SetCellData(sFunctionName,objJavaDialog,sJavaTable,iRow,iCol,sData)
	bGblFailedFunctionName = sFunctionName
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"),  "Depricated function : Fn_UI_JavaTable_SetCellData > Calling Function Fn_SISW_UI_JavaTable_Operations")
	If isNumeric(iRow) Then iRow = cInt(iRow)
	If isNumeric(iCol) Then iCol = cInt(iCol)
	Fn_UI_JavaTable_SetCellData = Fn_SISW_UI_JavaTable_Operations(sFunctionName, "SetCellData", objJavaDialog , sJavaTable, "", "", iRow, iCol, sData, "", "")
End Function 
'################################################################################################################
'###    FUNCTION NAME   :   Fn_UI_JavaTab_Select - Depricated Function
'###	Note : Please use function Fn_SISW_UI_JavaTab_Operations
'################################################################################################################
Function Fn_UI_JavaTab_Select(sFunctionName,objJavaDialog,sJavaTab, sSelect)
	bGblFailedFunctionName = sFunctionName
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & "Depricated function  : Fn_UI_JavaTab_Select :: Calling new function Fn_SISW_UI_JavaTab_Operations.")
	Fn_UI_JavaTab_Select = Fn_SISW_UI_JavaTab_Operations(sFunctionName, "Select", objJavaDialog, sJavaTab, sSelect)
End Function
'################################################################################################################
'###    FUNCTION NAME   : Fn_UI_JavaTree_Activate_Select_Node(sFunctionName,objJavaWindow,sTreeName,sNodeName)
'###
'###    DESCRIPTION     : This function  is used to Activate and Select the Node in JavaTree
'###
'###    PARAMETERS      : sFunctionName - Valid function name
'###                      objJavaWindow - Valid Java Window Object & it's Name
'###			  		  sTreeName -Valid Java Tree  name
'###                      sNodeName - Valid Java Node name
'###    Function Calls  : Fn_WriteLogFile()
'###
'###    HISTORY         : AUTHOR         DATE      VERSION
'###
'###    CREATED BY      : Dhananjay  19/04/2010    1.0
'###
'###    REVIWED BY      : Rizwan  
'###
'###    MODIFIED BY     :
'###    EXAMPLE         :    Fn_UI_JavaTree_Activate_Select_Node("Fn_MyTc_NavTree_NodeOperation",JavaWindow("MyTeamcenter"),"NavTree","Item")
'################################################################################################################
Function Fn_UI_JavaTree_Activate_Select_Node(sFunctionName,objJavaWindow,sTreeName,sNodeName)
	Dim iNode,iCount,bFlag,sNode
	Dim objTree
	bGblFailedFunctionName = sFunctionName
  sUIFail = sFunctionName + ">> Fn_UI_JavaTree_Activate_Select_Node >> " +  objJavaWindow.toString +">> " +  sTreeName

   	Set objTree=objJavaWindow.JavaTree(sTreeName)
	bFlag=False

   	'Check if Tree Exist
   	If objTree.Exist Then
		'Apply WaitProperty
		objTree.WaitProperty "enabled","1"
		'Check if Tree Enabled
		
		If  objTree.GetROProperty("enabled")="1" Then
			'Count the Number of Nodes in the tree
			iNode=objTree.GetROProperty("items count")

			'Check if the Node is present in the Tree
			For  iCount=0 to iNode-1
			sNode = objTree.GetItem(iCount)
			
				If  sNode=sNodeName Then
					bFlag=true
					Exit For
				End If
			Next
				'If Node Present in the tree
			If bFlag=true Then
				'Select the Node
				objTree.Select sNodeName
				'Log the Result						
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Sucessfully selected Node " &sNodeName &" of JavaTree " &sJavaTreeName& " of Function " &sFunctionName)
				Fn_UI_JavaTree_Activate_Select_Node=true
			else
				'Log the Result						
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),sNodeName & "Node is not present in the JavaTree" &sJavaTreeName& " of Function " &sFunctionName)
				Fn_UI_JavaTree_Activate_Select_Node=false
				Call ExitFromUI(sUIFail)
				End If
	
		else	
		'Log the Result						
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sJavaTreeName &"JavaTree is Disabled of Function " &sFunctionName)
		Fn_UI_JavaTree_Activate_Select_Node=False
		Call ExitFromUI(sUIFail)
		End If 
	else
	'Log the Result						
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sJavaTreeName &"JavaTree does not Exist of Function " &sFunctionName)
	Fn_UI_JavaTree_Activate_Select_Node=false
	Call ExitFromUI(sUIFail)
	
 
 	End If
	Set objTree = Nothing 
End Function
'#################################################################################################################
'###    FUNCTION NAME   :   Fn_UI_JavaList_ExtendSelect(sFunctionName, objJavaDialog, sJavaList, sItemsList)
'###	Note : Depricated Function :	Note: Use Fn_SISW_UI_JavaList_Operations
'###    EXAMPLE         :   call Fn_UI_JavaList_ExtendSelect("Fn_MyTc_Item_SaveAs", JavaWindow("MyTeamcenter").JavaWindow("MyTcApplet").JavaDialog("Save Item As"), "AvailableProjects", "sAvailedProj")
'################################################################################################################
Function Fn_UI_JavaList_ExtendSelect(sFunctionName, objJavaDialog, sJavaList, sItemsList)
	Dim sItems
	bGblFailedFunctionName = sFunctionName
	sItems = replace(sItemsList, ":", "~")
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Depricated function : Fn_UI_JavaList_ExtendSelect > Calling Function Fn_SISW_UI_JavaList_Operations")
	Fn_UI_JavaList_ExtendSelect =  Fn_SISW_UI_JavaList_Operations(sFunctionName, "ExtendSelect", objJavaDialog, sJavaList, sItems, "", "")
End Function
'#################################################################################################################
'###    FUNCTION NAME   :   Fn_UI_JavaTree_ExtendSelect(sFunctionName, objJavaDialog, sJavaTree, sNodeName)
'###
'###    DESCRIPTION     :   Select multiple nodes of a tree.
'###
'###    PARAMETERS      :   1. sFunctionNamee - Valid function name
'###                        2. objJavaDialog - Valid Java Dialog Name
'###                        3. sJavaTree - Vaild JavaTree name
'###                        4. sNodeName - Miltiple nodes which are to be selected
'###                        					
'###                        
'###    Function Calls  :   Fn_WriteLogFile ()
'###
'###    HISTORY         :   AUTHOR                   DATE        VERSION
'###
'###    CREATED BY      :   Ketan Raje           19/04/2010        1.0
'###
'###    REVIWED BY      :   Rizwan
'###
'###    MODIFIED BY     :
'###    EXAMPLE         :   call Fn_UI_JavaTree_ExtendSelect(Fn_MyTc_NavTree_NodeOperation, objJavaDialog, sJavaTree, sNodeName)
'################################################################################################################

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function Fn_UI_JavaTree_ExtendSelect(sFunctionName, objJavaDialog, sJavaTree, sNodeNames)

   Dim intNodeCount, NodeLists, iCounter
   Dim objJavaTree
	bGblFailedFunctionName = sFunctionName
     sUIFail = sFunctionName + ">> Fn_UI_JavaTree_ExtendSelect >> " +  objJavaDialog.toString +">> " +  sJavaTree
     
	Set objJavaTree = objJavaDialog.JavaTree(sJavaTree)
 	
   	If objJavaTree.Exist Then
		objJavaTree.WaitProperty "enabled", "1"
		If objJavaTree.GetROProperty("enabled") = "1"  Then
				'Split the string where "'," exist
				NodeLists = Split(sNodeNames,",")
				intNodeCount = Ubound(NodeLists)
				For iCounter = 0 to intNodeCount				
					objJavaTree.ExtendSelect NodeLists(iCounter)
					
					'Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Selected multiple Tree Nodes [" + NodeLists(iCounter) + "]")
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Selected multiple Tree Nodes [" + NodeLists(iCounter) + "]" &"of JavaTree " &sJavaTree &"of Function " &sFunctionName)
				Next
				Fn_UI_JavaTree_ExtendSelect = TRUE
			Else
					
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sJavaTree &" JavaTree is Disable of Function " &sFunctionName)
					Fn_UI_JavaTree_ExtendSelect = FALSE
                    Call ExitFromUI(sUIFail)
			End If	
            Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sJavaTree & " does not  exist of Function " &sFunctionName)
			Fn_UI_JavaTree_ExtendSelect = False
			Call ExitFromUI(sUIFail)
	End If

Set objJavaTree = Nothing 
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'#################################################################################################################
'###    FUNCTION NAME   :   Fn_UI_JavaObject_Click(sFunctionName, objJavaDialog, sJavaObjecName,iCordx,iCordy,sMouseButton)
'###
'###    DESCRIPTION     :   This function is used to click Java Object
'###
'###    PARAMETERS      :   sFunctionNamee - Valid function name
'###                        objJavaDialog - Valid Java Dialog Name
'###                        sJavaObjecName - Vaild Java Object name
'###                      	iCordx- Valid x cordinates				 
'###                        iCordy - Valid y cordinates
'###						sMouseButton- Valid mouse button (LEFT/RIGHT)- Optional
'###
'###    Function Calls  :   Fn_WriteLogFile ()
'###
'###    HISTORY       	:   AUTHOR          DATE        VERSION
'###
'###    CREATED BY      :   Rajeev          21/04/2010    1.0
'###
'###    REVIWED BY      :   Rajesh	    21/04/2010	  1.0
'###
'###    MODIFIED BY     :
'###    EXAMPLE         :   call Fn_UI_JavaObject_Click("Fn_itemBasicCreate", JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("BOMViewSaveAs"), "TypeButton",56,3,"LEFT")
'################################################################################################################


Function Fn_UI_JavaObject_Click(sFunctionName, objJavaDialog, sJavaObjecName,iCordx,iCordy,sMouseButton)
Dim objJavaObject
	bGblFailedFunctionName = sFunctionName
 sUIFail = sFunctionName + ">> Fn_UI_JavaObject_Click >> " +  objJavaDialog.toString +">> " +  sJavaObjecName

	Set objJavaObject = objJavaDialog.JavaObject(sJavaObjecName)

	If objJavaObject.Exist Then
		objJavaObject.WaitProperty "enabled", "1"
		     '*************************************Checking that the JavaObject is Exists or not**********************************
			If objJavaObject.GetROProperty("enabled") = "1" Then
				'*************************************Checking the JavaObject is clicked*********************************
				If sMouseButton<>"" Then
					objJavaObject.Click iCordx,iCordy,sMouseButton
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Sucessfully " &sMouseButton &"Clicked at " &iCordx &"," &iCordy &"Co-ordinates on JavaObject " &sJavaObjecName &"of Function " &sFunctionName)
					Fn_UI_JavaObject_Click = True
				Else
					objJavaObject.Click iCordx,iCordy
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Sucessfully Clicked at " &iCordx &"," &iCordy &"Co-ordinates on JavaObject " &sJavaObjecName &"of Function " &sFunctionName)
					Fn_UI_JavaObject_Click = True
				End If
			 Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sJavaObjecName & " JavaObject is not enabled  of Function " &sFunctionName)
                		Fn_UI_JavaObject_Click = False
						Call ExitFromUI(sUIFail)
			End If
	Else
		     '************************************* Checking if JavaObject not exist **********************************
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sJavaObjecName & " JavaObject does not exist of Function " &sFunctionName)
                Fn_UI_JavaObject_Click = False
				Call ExitFromUI(sUIFail)
	End If

Set objJavaObject = Nothing 
End Function

'#######################################################################################

'#################################################################################################################
'###    FUNCTION NAME   :   Fn_UI_JavaDialog_ChildObjects(sFunctionName, objJavaDialog, sJavaObjecName)
'###
'###    DESCRIPTION     :   Display JavaChild Objects Name
'###
'###    PARAMETERS      :   sFunctionName - Valid function name
'###                        objJavaDialog - Valid Java Dialog/Window 
'###                        sJavaObjecName - Vaild JavaChild bject name
'###                      					 
'###                        
'###    Function Calls  :   Fn_WriteLogFile ()
'###
'###    HISTORY         :   AUTHOR          DATE        VERSION
'###
'###    CREATED BY      :   Rajeev          21/04/2010        1.0
'###
'###    REVIWED BY      :   Rajesh			21/04/2010		  1.0
'###
'###    MODIFIED BY     :
'###    EXAMPLE         :   Call Fn_UI_JavaDialog_ChildObjects("Fn_itemBasicCreate", "JavaWindow("TcDefaultApplet").JavaDialog("BOMViewSaveAs")", "List")
'################################################################################################################

Function Fn_UI_JavaDialog_ChildObjects(sFunctionName, objJavaDialog, sChildObjects)

sUIFail = sFunctionName + ">> Fn_UI_JavaDialog_ChildObjects >> " +  objJavaDialog.toString +">> " +  sChildObjects
	bGblFailedFunctionName = sFunctionName
	If objJavaDialog.Exist Then
		objJavaDialog.WaitProperty "enabled", "1"
				 '*************************************Checking ChildObject is enable or not**********************************
			If objJavaDialog.GetROProperty("enabled") = "1"  Then
				'*************************************Listing  The JavaChildObjects*********************************
                          	Set Fn_UI_JavaDialog_ChildObjects =objJavaDialog.ChildObjects(sChildObjects)
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),sChildObjects & " are Child Objects  of Function " &sFunctionName)
						
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), objJavaDialog.toString & " JavaDialog/Window is not enabled to Check ChildObject " &sChildObjects &" of Function " &sFunctionName)
				Fn_UI_JavaDialog_ChildObjects = False
				Call ExitFromUI(sUIFail)
			End If
	Else
		     '************************************* Checking if JavaChildObject not exist **********************************
			 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), objJavaDialog.toString & " JavaDialog/Window does not exist to check ChildObject " &sChildObjects &"of Function " &sFunctionName)
        	 Fn_UI_JavaDialog_ChildObjects = False
			 Call ExitFromUI(sUIFail)
	End If

End Function


'#######################################################################################

'#############################################################################################################################################
'###    FUNCTION NAME   :   Fn_UI_JavaTree_Expand()
'###
'###    DESCRIPTION     :   Function is used to Expand Node/Element from the JavaTree.
'###
'###    PARAMETERS      :   1. sFunctionName : Valid Function name,
'###                                     2. objJavaDialog : Valid Dialog Path,
'###                                    3.  sJavaTree  : Valid Javatree Name,
'###                                    4. sElementToExpand  : Valid Note/Element to be Expanded ( Element seperated by :)
'###
'###
'###    HISTORY         :   AUTHOR      DATE        VERSION
'### 
'###    CREATED BY      :   Amol     19/04/2010   1.0
'###
'###    REVIWED BY      : 
'###
'###    MODIFIED BY     :   NA
'###    EXAMPLE         :   Call Fn_UI_JavaTree_Expand("Fn_UI_JavaTree_Expand",JavaWindow("MyTeamcenter"),"NavTree","Home:AutomatedTests")
'#############################################################################################################################################




Function Fn_UI_JavaTree_Expand(sFunctionName, objJavaDialog, sJavaTree,sElementToExpand)


Dim objJavaTree
	bGblFailedFunctionName = sFunctionName
  sUIFail = sFunctionName + ">> Fn_UI_JavaTree_Expand >> " +  objJavaDialog.toString +">> " +  sJavaTree

'Object Creation
Set objJavaTree = objJavaDialog.JavaTree(sJavaTree)     
       
' Verify  JavaTree object exists
  If objJavaTree.Exist Then

 objJavaTree.WaitProperty "enabled", "1"               
' Verify JavaTree Object is enabled 

		   If objJavaTree.GetROProperty("enabled") = "1"  Then 
			    
				objJavaTree.Expand sElementToExpand
			   ' Report message when Expanded the element from the tree sucessfully.
				 Call Fn_WriteLogFile(Environment.Value("TestLogFile"),  "Successfully Expanded/Node " &sElementToExpand &"is Present in" &sJavaTree &" of Function " &sFunctionName)
				 Fn_UI_JavaTree_Expand= True
		
		          ' Report error when JavaTree object is disable.
			 Else
				 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sJavaTree &"JavaTree is Disabled  of Function " &sFunctionName)
				 Fn_UI_JavaTree_Expand= False
				 Call ExitFromUI(sUIFail)
		  End If
' Report error when JavaTree object does not exists.
  Else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sJavaTree & " JavaTree does not exist of Function " &sFunctionName)
		Fn_UI_JavaTree_Expand= False
		Call ExitFromUI(sUIFail)
End If
'Disassociate an object variable from any actual object
Set objJavaTree=Nothing
' End of function 
End Function    
'#######################################################################################



'#######################################################################################
'##################################################################################################################################################
'###    FUNCTION NAME   :   Fn_UI_WinButton_Click(sFunctionName, objJavaDialog, sWinButton,iXValue,iYValue,coMicButton) 
'###
'###    DESCRIPTION     :   Function to verify WinButton is enabled and to Click the mouse button at X,Y Co-ordinates. 
'###
'###    PARAMETERS      :   sFunctionName : Valid Function Name, 
'###			    		objJavaDialog : Valid Dialog Path,
'###					    sWinButton    : Valid Button Name,
'###					    iXValue	  : Valid X Co-ordiate value,
'###			    		iYValue	  : Valid Y Co-ordinate Value,
'###			    		coMicButtonToClick : Valid Mouse button to be Clicked 
'###                        
'###    Function Calls  :   Fn_WriteLogFile()
'###
'###    HISTORY         :   AUTHOR      DATE        VERSION
'###
'###    CREATED BY      :   Manisha    	21/04/2010   1.0
'###
'###    REVIWED BY      :	Rajesh		21/04/2010	 1.0
'###
'###    MODIFIED BY     :   NA		
'###    EXAMPLE         :   Fn_UI_WinButton_Click(Fn_RunRCAFScript,JavaWindow("RCAF Console").Dialog("Open"),"Open",5,5,micLeftBtn)
'##################################################################################################################################################


Function Fn_UI_WinButton_Click(sFunctionName, objJavaDialog, sWinButton,iXValue,iYValue,coMicButton)
	Dim objWinButton
	bGblFailedFunctionName = sFunctionName
		sUIFail = sFunctionName + ">> Fn_UI_WinButton_Click >> " +  objJavaDialog.toString +">> " +  sWinButton
	
'Object Creation
	Set objWinButton = objJavaDialog.WinButton(sWinButton)

'Verify  WinButton object exists
	If objWinButton.Exist Then

'Synchronization Point for an WinButton Object
		objWinButton.WaitProperty "enabled", "1"

'Verify Object is enabled 
	    If objWinButton.GetROProperty("enabled") = "1" OR objWinButton.GetROProperty("enabled") = True Then


		      If  iXValue <> "" AND iYValue <> "" Then
						'Click the mouse button at X,Y Co-ordinates
						objWinButton.Click iXValue,iYValue,coMicButton
                        'log on success
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),  "Sucessfully clicked on WinButton" & sWinButton &" at Co-ordinates " &  iXValue &"," &iYValue & " of Function " & sFunctionName)
              Else
						objWinButton.Click
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Clicked on " & sWinButton & "WinButton of Function " & sFunctionName)
				 End If

          Fn_UI_WinButton_Click = True


'Report Error when WinButton object is disable.
        Else
	       Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sWinButton & "WinButton  is disabled of Function " &sFunctionName)
	       Fn_UI_WinButton_Click = False
        End If

'Report Error when WinButton object does not exists.
	Else
	      Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sWinButton & "WinButton does not exist of Function " &sFunctionName)
		  Fn_UI_WinButton_Click = False
	End If

Set objWinButton = Nothing

'End of function 
End Function

'#############################################################################################################################################


'#############################################################################################################################################
'###    FUNCTION NAME   :   Fn_UI_JavaTree_Collapse(sFunctionName, objJavaDialog, sJavaTree,sElementToCollapse)
'###
'###    DESCRIPTION     :   Function is used to Collapse Node/Element from the JavaTree.
'###
'###    PARAMETERS      :   1. sFunctionName	: Valid Function name,
'### 			                                 2. objJavaDialog	: Valid Dialog Path,
'### 			                                 3. sJavaTree		: Valid Javatree Name,
'### 			                                 4.  sElementToCollapse 	: Valid Note/Element to be Collapseed ( Element seperated by :)
'###
'###
'###    HISTORY         :   AUTHOR      DATE        VERSION
'###	
'###    CREATED BY      :   Amol    	19/04/2010   1.0
'###
'###    REVIWED BY      :	Rizwan
'###
'###    MODIFIED BY     :   NA
'###    EXAMPLE         :   Call Fn_UI_JavaTree_Collapse("Fn_UI_JavaTree_Collapse",JavaWindow("MyTeamcenter"),"NavTree","Home:AutomatedTests")
'#############################################################################################################################################

Function Fn_UI_JavaTree_Collapse(sFunctionName, objJavaDialog, sJavaTree,sElementToCollapse)
Dim objJavaTree
	bGblFailedFunctionName = sFunctionName
  sUIFail = sFunctionName + ">> Fn_UI_JavaTree_Collapse >> " +  objJavaDialog.toString +">> " +  sJavaTree

'Object Creation
Set objJavaTree = objJavaDialog.JavaTree(sJavaTree)												



' Verify  JavaTree object exists
   If objJavaTree.Exist Then

' Synchronization Point for an Java Tree Object 
	objJavaTree.WaitProperty "enabled", "1"															

' Verify JavaTree Object is enabled 
	      If objJavaTree.GetROProperty("enabled") = "1"  Then											

			       
    	           objJavaTree.Collapse sElementToCollapse
				'  Report message when Collapsed the element from the tree sucessfully.
                    Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Collapse Node "  & sElementToCollapse &"is Present on " &sjavaTree &" JavaTree of Function " &sFunctionName)
	                Fn_UI_JavaTree_Collapse= True
                           
' Report error when JavaTree object is disable.
	         Else
   	                 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sJavaTree &"JavaTree is Disabled of Function " &sFunctionName)
  	                 Fn_UI_JavaTree_Collapse = False
	                 Call ExitFromUI(sUIFail)
 	       End If
' Report error when JavaTree object does not exists.
	Else
	   Call Fn_WriteLogFile(Environment.Value("TestLogFile"),sJavaTree & " JavaTree does not exist of Function " &sFunctionName)
	   Fn_UI_JavaTree_Collapse = False
	   Call ExitFromUI(sUIFail)
 End If
'Disassociate an object variable from any actual object
Set objJavaTree=Nothing
' End of function 
End Function  


'############################################################################################################################
'###    FUNCTION NAME    :   Fn_UI_JavaTree_OpenContextMenu(sFunctionName,objJavaDialog,sJavaTreeName,sMenuName)
'###
'###    DESCRIPTION      :   This function is used open the context menu.
'###
'###    PARAMETERS       :   sFunctionName   :  Valid Function name,
'###  			     objJavaDialog   :  Valid Object name,
'###  			     objJavaTree     :  Valid name for Tree,
'###  			     sMenuNme        :  Valid Menu Name 
'###
'###    Function Calls   :   Fn_WriteLogFile ()
'###
'###    HISTORY          :   AUTHOR           DATE               VERSION
'###
'###    CREATED BY       :   Manish           21/04/2010         1.0
'###
'###    REVIWED BY       :	Rajesh			  21/04/2010
'###
'###    MODIFIED BY      :   NA
'###    EXAMPLE          :  Fn_UI_JavaTree_OpenContextMenu("Fn_MyTc_NavTree_NodeOperation",JavaWindow("MyTeamcenter"),"NavTree","Exam")
'##############################################################################################################################

Function Fn_UI_JavaTree_OpenContextMenu(sFunctionName,objJavaDialog,sJavaTreeName,sMenuName)
Dim objOpenContext
	bGblFailedFunctionName = sFunctionName
     sUIFail = sFunctionName + ">> Fn_UI_JavaTree_OpenContextMenu >> " +  objJavaDialog.toString +">> " +  sJavaTreeName

'Object Creation
   Set objOpenContext= objJavaDialog.JavaTree(sJavaTreeName)

'Verify  OpenContext object exists
   If objOpenContext.Exist Then

 'Synchronization Point for context menu Object
			objOpenContext.WaitProperty "enabled", "1"

			 If objOpenContext.GetROProperty("enabled") = "1"  Then
			    objOpenContext.OpenContextMenu(sMenuName)

 'Report message Open context menu  is successful.
                             Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Sucessfully opened the Context Menu " & sMenuName &" of JavaTree " &sJavaTreeName & " of Function " &sFunctionName)
			     			Fn_UI_JavaTree_OpenContextMenu=True

             Else

'Report error/message when open Context menu  is disable.
	                    Call Fn_WriteLogFile(Environment.Value("TestLogFile"),sJavaTreeName &" JavaTree is Disable of Function " &sFunctionName)
	                    Fn_UI_JavaTree_OpenContextMenu = False
	                    Call ExitFromUI(sUIFail)    
             End If

'Report error/message when  open Context  menu object does not exists.
 Else
       Call Fn_WriteLogFile(Environment.Value("TestLogFile"),sJavaTreeName & "JavaTree does not exist of Function " &sFunctionName)
       Fn_UI_JavaTree_OpenContextMenu = False
        Call ExitFromUI(sUIFail)
End If

'Clear memory of Context Menu
set objOpenContext = Nothing 

End Function

'#################################################################################################################
'###    FUNCTION NAME   :    Fn_UI_JavaStaticText_Click(sFunctionName,objJavaDialog,sJavaStaticText,iRow,iCol, sPosition)
'###
'###    DESCRIPTION     :   This function  is used to click item on java static text
'###
'###    PARAMETERS      :   sFunctionName - Valid function name
'###                        objJavaDialog - Valid Java Dialog Name
'###                        sJavaStaticText- Vaild Static Text Name
'###                        iRow  - Valid Row No.
'###                        iCol  - Valid Column No. 
'###                        sPosition - Valid Position (LEFT, RIGHT)
'###                        
'###    Function Calls  :   Fn_UI_JavaStaticText_Click(sFunctionName.objJavaDialog,sJavaStaticText, iRow,iCol, sPosition)
'###
'###    HISTORY         :   AUTHOR                   DATE        VERSION
'###
'###    CREATED BY      :   Mahendra      			15/04/2010		1.0  
'###
'###    REVIWED BY      :     Rizwan               
'###
'###    MODIFIED BY     :
'###    EXAMPLE         :   call Fn_UI_JavaStaticText_Click(" Fn_TcObjectDelete", objDialog, "More...", 1, 1, "LEFT")
'################################################################################################################
 
Function Fn_UI_JavaStaticText_Click(sFunctionName, objJavaDialog, sJavaStaticText, iRow, iCol, sPosition)
Dim objJavaStaticText
	bGblFailedFunctionName = sFunctionName
  sUIFail = sFunctionName + ">> Fn_UI_JavaStaticText_Click >> " +  objJavaDialog.toString +">> " +  sJavaStaticText

	'Set an list object on variable        
	Set objJavaStaticText = objJavaDialog.JavaStaticText(sJavaStaticText)
	
	'checking List object exist or not
	If objJavaStaticText.Exist Then  

		'Syncronization point
		objJavaStaticText.WaitProperty "displayed", "1"
										
		'if its enabled click on the java list of iRow and iCol
		If sPosition<>"" Then
					objJavaStaticText.Click iRow, iCol, sPosition 
					'log the success
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Sucessfully " &sPosition &" Mouse button Clicked on " & sJavaStaticText &"at " & iRow  &"," & iCol &"Coordinates of Function " & sFunctionName)
    	 Else
					objJavaStaticText.Click iRow, iCol			 
					'log the success
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Sucessfully Mouse button Clicked on " & sJavaStaticText &"at " & iRow  &"," & iCol &"Coordinates of Function " & sFunctionName)
		End If
					' Return True from Function
					Fn_UI_JavaStaticText_Click = True
	Else
                                     
		'log the failure when list not enabled
		 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sJavaStaticText & " JavaStaticText does not exist of Function " & sFunctionName)
		 'Return False from function
		 Fn_UI_JavaStaticText_Click = False
		 Call ExitFromUI(sUIFail)
 
	End If
Set objJavaStaticText = Nothing 
End Function
'#######################################################################################
'###    FUNCTION NAME   :   Fn_UI_JavaTable_DoubleClickCell(sFunctionName, objJavaDialog, sJavaTable,varRowValue,varColValue,varMicButton,varModifier)
'###
'###    DESCRIPTION     :   Function is used to Double-click the specified cell in the table.
'###
'###    PARAMETERS      :   sFunctionName	: Valid Function name,
'### 						objJavaDialog	: Valid Dialog Path,
'### 						sJavaTable		: Valid JavaTable Name,
'### 						varRowValue		: Valid row number or row header label to be selected
'###						varColValue		: Valid Column number or column header label to be selected
'###						varMicButton	: Left or right mouse button to be double clicked (Optional)
'###						varModifier		: keyboard keys used to perform the operation (Optional)
'###    Function Calls  :   Fn_WriteLogFile()
'###
'###    HISTORY         :   AUTHOR            DATE        VERSION
'###
'###    CREATED BY      :   Manisha    19/04/2010   1.0
'###
'###    REVIWED BY      :	
'###
'###    MODIFIED BY     :   NA
'###    EXAMPLE         :   Fn_JavaTable_DoubleClickCell("Fn_PSE_BOMTable_NodeOperation",JavaWindow("StructureManager").JavaApplet("PSEApplet"),"BOMTable",StrIndex,StrColName, "LEFT", "NONE")
'#######################################################################################
Function Fn_UI_JavaTable_DoubleClickCell(sFunctionName, objJavaDialog, sJavaTable,varRowValue,varColValue,varMicButton,varModifier)
Dim objJavaTable
	bGblFailedFunctionName = sFunctionName
		sUIFail = sFunctionName + ">> Fn_UI_JavaTable_DoubleClickCell >> " +  objJavaDialog.toString +">> " +  sJavaTable
		
		
'Object Creation
	Set objJavaTable = objJavaDialog.JavaTable(sJavaTable)												

'Verify  JavaTable object exists
	If objJavaTable.Exist Then

'Synchronization Point for an JavaTable Object 
		objJavaTable.WaitProperty "enabled", "1"

'Verify Object is enabled 
		If objJavaTable.GetROProperty("enabled") = "1"  Then											
		
'Verify Specified Row and column exists in the JavaTable
		If Cint(objJavaTable.GetROProperty("rows")) >= Cint(varRowValue) AND Cint(objJavaTable.GetROProperty("cols")) >= Cint(varColValue) Then
			If  varMicButton<>"" AND varModifier<>"" Then
				' Double-click the specified cell in the table.
				objJavaTable.DoubleClickCell varRowValue,varColValue,varMicButton,varModifier
				' Report message that specified Row and Column exists in the JavaTable
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),  " Table Cell " &varRowValue &"," &varColValue &" is Double Click on Table  " & sJavaTable & " of Function " &sFunctionName)
				Fn_UI_JavaTable_DoubleClickCell = True
				Else
				objJavaTable.DoubleClickCell varRowValue,varColValue
				' Report message that specified Row and Column exists in the JavaTable
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),  " Table Cell " &varRowValue &"," &varColValue &" is Double Click on Table  " & sJavaTable & " of Function " &sFunctionName)
				Fn_UI_JavaTable_DoubleClickCell = True
			End If
		Else

			' Report error that specified Row and Column does not exists in the JavaTable
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),  " Table Cell " &varRowValue &"," &varColValue &"does not exists in  " & sJavaTable & " of Function " &sFunctionName)
			Fn_UI_JavaTable_DoubleClickCell= False
			Call ExitFromUI(sUIFail)
		End If

		Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "JavaTable " &sJavaTable & " is disable of Function  " &sFunctionName)
			Fn_UI_JavaTable_DoubleClickCell= False
			Call ExitFromUI(sUIFail)
		End If

'Report error when JavaTable object does not exists.
	Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sJavaTable & " JavaTable does not exists " &" of Function  " &sFunctionName)
			Fn_UI_JavaTable_DoubleClickCell = False
			Call ExitFromUI(sUIFail)
	End If

Set objJavaTable=Nothing

End Function
'###############################################################################################################
'##### 		Function Name	:	Fn_UI_JavaTable_ExtendRow(sFunctionName,objJavaDialog,sJavaTable,iRow)
'#####
'#####  	Examples	:  Fn_UI_JavaTable_ExtendRow("Fn_PSE_BOMTable_NodeOperation",JavaWindow("StructureManager").JavaApplet("PSEApplet"),"BOMTable",3 )
'###############################################################################################################
Function Fn_UI_JavaTable_ExtendRow(sFunctionName,objJavaDialog,sJavaTable,iRow)
	bGblFailedFunctionName = sFunctionName
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"),  "Depricated function : Fn_UI_JavaTable_ExtendRow > Calling Function Fn_SISW_UI_JavaTable_Operations")
	Fn_UI_JavaTable_ExtendRow = Fn_SISW_UI_JavaTable_Operations(sFunctionName, "ExtendRow", objJavaDialog , sJavaTable, "", "", iRow, "", "", "", "")
End Function
'#################################################################################################################
'###    FUNCTION NAME   :    Fn_UI_WinMenu_BuildMenuPath(sFunctionName, objJavaDialog,sWinMenu,sWinMenuPath)
'###
'###    DESCRIPTION     :   This function  is used to reach at the given menu item
'###
'###    PARAMETERS      :   sFunctionName - Valid function name
'###                        objJavaDialog -  Valid Dialog/Window Path
'###                        sWinMenu  -  Valid WinMenu Name or  Index
'###                        
'###    Function Calls  :     Fn_WriteLogFile()
'###
'###    HISTORY         :   AUTHOR                   DATE        VERSION
'###
'###    CREATED BY      :   Deepak kumar                                                   20/04/2010    1.0
'###
'###    REVIWED BY      :    Rajesh                                                                 20/04/2010           1.0                
'###
'###    MODIFIED BY     :
'###    EXAMPLE         :  Fn_UI_WinMenu_BuildMenuPath("Fn_InvokeTeamCenter" ,JavaWindow("DefaultWindow"), "ContextMenu",2:3:1)
'################################################################################################################

Function Fn_UI_WinMenu_BuildMenuPath(sFunctionName, objJavaDialog,sWinMenuType,sWinMenuPath)
Dim objWinMenu
  	bGblFailedFunctionName = sFunctionName
  	sUIFail = sFunctionName + ">> Fn_UI_WinMenu_BuildMenuPath >> " +  objJavaDialog.toString +">> " +  sWinMenuPath
  
           'Setting the Object 
           Set objWinMenu = objJavaDialog.WinMenu(sWinMenuType)

           'Check  existence of object
           If objWinMenu.Exist Then

                 'Apply wait property on Object
                 objWinMenu.WaitProperty "enabled", "1"

                 'Checking wait property Enable or not
                 If objWinMenu.GetROProperty("enabled") = "1"  Then

                     'objWinMenu.Select
                      aMenuArray = Split(sWinMenuPath , ",")

                     'reaching the desired menu  through Build MenuPath function
                     ItemPath = objWinMenu.BuildMenuPath(aMenuArray)

                     'Checking  for reching  on the perticular  menu
                     If ItemPath <> Null Then
                                                                                                
                         'log the success
                         Call Fn_WriteLogFile( Environment.Value("TestLogFile"), sWinMenuPath &" Menu Path is Set for WinMenu " &sWinMenuType  & " of Function " & sFunctionName)
										 
						 'Return true from function
                          Fn_UI_WinMenu_BuildMenuPath= True
                     Else
                     	'log the failure when menu path not found
                     	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sWinMenuPath & " Menu item doesn't exists of WinMenu " &sWinMenuType  & " of Function " & sFunctionName)
                                                                                                
												 
						'Return False from function
                        Fn_UI_WinMenu_BuildMenuPath= False
						Call ExitFromUI(sUIFail)
                     End If

                  Else
                 	 Call Fn_WriteLogFile(Environment.Value("TestLogFile"),  sWinMenuType & " WinMenu is not Enable "&" of Function " &sFunctionName)
                                                               
					'Return False from function
                    Fn_UI_WinMenu_BuildMenuPath= False
					Call ExitFromUI(sUIFail)
                  End if

           Else
               Call Fn_WriteLogFile(Environment.Value("TestLogFile"),sWinMenuType & "  WinMenu does not Exist "&"of Function " &sFunctionName)
                                
				'Return False from function
				Fn_UI_WinMenu_BuildMenuPath= False
				Call ExitFromUI(sUIFail)
           End if
Set objWinMenu = Nothing 
End Function
'############################################################################################################
'###
'###    FUNCTION NAME   :  Fn_UI_ObjectCreate(sFunctionName, sReferencePath) - Depricated Function
'###
'###	Note : Please use function Fn_SISW_UI_Object_Operations
'#############################################################################################################
Function Fn_UI_ObjectCreate(sFunctionName, sReferencePath)
	bGblFailedFunctionName = sFunctionName
	Set Fn_UI_ObjectCreate = Fn_SISW_UI_Object_Operations(sFunctionName,"Create", sReferencePath,"")
End Function
'#########################################################################################################
'###
'###    FUNCTION NAME   :  Fn_UI_ObjectExist(sFunctionName, sReferencePath) - Depricated Function
'###
'###	Note : Please use function Fn_SISW_UI_Object_Operations
'#############################################################################################################
Function Fn_UI_ObjectExist(sFunctionName, sReferencePath)
	bGblFailedFunctionName = sFunctionName
	Fn_UI_ObjectExist = Fn_SISW_UI_Object_Operations(sFunctionName,"Enabled", sReferencePath,"")
End Function
'#########################################################################################################
'###    FUNCTION NAME   :   Fn_Edit_Box_GetValue(sFunctionName,objJavaDialog,sJavaEditBoxName)
'###
'###    EXAMPLE         : Fn_Edit_Box_GetValue("Fn_TeamcenterLogin",JavaWindow("Teamcenter Login"),"User ID")
'#############################################################################################################
Function Fn_Edit_Box_GetValue(sFunctionName,objJavaDialog,sJavaEditBoxName)
	bGblFailedFunctionName = sFunctionName
	Fn_Edit_Box_GetValue= Fn_SISW_UI_JavaEdit_Operations(sFunctionName, "GetText", objJavaDialog, sJavaEditBoxName, "")
End Function
'#################################################################################################################
'###    FUNCTION NAME   :    Fn_Java_StaticText_Exist(sFunctionName,objJavaDialog,sJavaStaticTextExist) - Depricated Function
'###	Note : Please use function Fn_SISW_UI_Object_Operations, special function for StaticText is not required.
'#############################################################################################################
Function Fn_Java_StaticText_Exist(sFunctionName,objJavaDialog, sJavaStaticTextExist)
	bGblFailedFunctionName = sFunctionName
	Fn_Java_StaticText_Exist = Fn_SISW_UI_Object_Operations(sFunctionName & ">> Fn_Java_StaticText_Exist","Exist", objJavaDialog.JavaStaticText(sJavaStaticTextExist),"")
End Function
'#################################################################################################################
'###    FUNCTION NAME   :    Fn_JavaTable_Type(sFunctionName,objJavaDialog, sTableName,sTypeString)
'###	Note 			: Depricated Function :	Note: Use Fn_SISW_UI_JavaTable_Operations
'###    EXAMPLE         :   Fn_JavaTable_Type("Fn_SetPerspective",JavaWindow("DefaultWindow").JavaWindow("Open Perspective"),"Table,"ABC")
'################################################################################################################
Function Fn_JavaTable_Type(sFunctionName,objJavaDialog, sTableName,sTypeString)
	bGblFailedFunctionName = sFunctionName
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"),  "Depricated function : Fn_JavaTable_Type > Calling Function Fn_SISW_UI_JavaTable_Operations")
	Fn_JavaTable_Type = Fn_SISW_UI_JavaTable_Operations(sFunctionName, "Type", objJavaDialog , sTableName, "", "", "", "", sTypeString, "", "")
End function
'#################################################################################################################
'###    FUNCTION NAME   :    Fn_UI_Object_GetROProperty(sFunctionName,objJavaDialog, sPropertyName)
'###
'###    DESCRIPTION     :   This function  is used to Return RO Property of given Object
'###
'###    PARAMETERS      :   sFunctionName - Valid function name
'###                        					objJavaDialog - Valid Java Dialog Name
'###                      					  	sPropertyName- Vaild Property name
'###                              				
'###                        
'###                        
'###    Function Calls  :   Fn_WriteLogFile()
'###
'###    HISTORY         :   AUTHOR                   DATE      	  VERSION
'###
'###    CREATED BY      :   Sandeep      		30/04/2010		1.0  
'###
'###    REVIWED BY      :   Sammer              
'###
'###    MODIFIED BY     :
'###    EXAMPLE         :   Fn_UI_Object_GetROProperty("Fn_SetPerspective",JavaWindow("DefaultWindow").JavaTable("table"),"rows")
'################################################################################################################
Function Fn_UI_Object_GetROProperty(sFunctionName,objJavaDialog, sPropertyName)
	Dim objGetROProperty
	bGblFailedFunctionName = sFunctionName
		sUIFail = sFunctionName + ">> Fn_UI_Object_GetROProperty >> " +  objJavaDialog.toString +">> " +  sPropertyName
	
		'Setting Java Table Object
		Set objGetROProperty = objJavaDialog
		If objGetROProperty.Exist Then
				Fn_UI_Object_GetROProperty=objGetROProperty.GetROProperty(sPropertyName)
            	Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"GetROProperty" & sPropertyName&" for object "& objGetROProperty.toString &" is returned successfully of Function " &sFunctionName)
		Else
				'Checking if table not exist
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), objGetROProperty.toString & " Property does not exist of Function " &sFunctionName)
				Fn_UI_Object_GetROProperty = False
				'Call ExitFromUI(sUIFail)
	End If
	Set objGetROProperty=Nothing
End Function

'#############################################################################################################################################################
'#################################################################################################################
'###    FUNCTION NAME   :    Fn_UI_Object_SetTOProperty(sFunctionName,objJavaDialog,sProperty,sPropValue)s
'###
'###    DESCRIPTION     :   This function  is used to Set TO Property For given Object
'###
'###    PARAMETERS      :   sFunctionName - Valid function name
'###                        					objJavaDialog - Valid Java Dialog Name
'###                      					  	sButtonName- Vaild Button  Name
'###                              				sProperty-Valid Property
'###                        					sPropValue-New value
'###                        
'###    Function Calls  :   Fn_WriteLogFile()
'###
'###    HISTORY         :   AUTHOR                   DATE        VERSION
'###
'###    CREATED BY      :   Sandeep      			3004/2010		1.0  
'###
'###    REVIWED BY      :   Sameer
'###
'###    MODIFIED BY     :
'###    EXAMPLE         :Fn_UI_Object_SetTOProperty("Fn_SetPerspective",JavaWindow("DefaultWindow").JavaButton("Click"),"Label","NewLabel")
'###										
'################################################################################################################

Function Fn_UI_Object_SetTOProperty(sFunctionName,objJavaDialog,sProperty,sPropValue)
	Dim objSetTOProperty
	bGblFailedFunctionName = sFunctionName
	sUIFail = sFunctionName + ">> Fn_UI_Object_SetTOProperty >> " +  objJavaDialog.toString +">> " +  sProperty
	
	'Setting object to javabutton
	Set objSetTOProperty=objJavaDialog

	'Checking existance of object
	If objSetTOProperty.Exist Then

					objSetTOProperty.SetTOProperty sProperty,sPropValue
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Set the TO PRoperty"&sProperty &" as " & sPropValue &" for " & objJavaDialog.toString & "  of Function " &sFunctionName)
					Fn_UI_Object_SetTOProperty=True

	Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),objJavaDialog.toString & " Property does not Exist of function" &sFunctionName)
				Fn_UI_Object_SetTOProperty=False
		
	End If
	Set objSetTOProperty = Nothing 
End Function
'#################################################################################################################
'###    FUNCTION NAME   :    Fn_Table_GetRowCount(sFunctionName,objJavaDialog, sTableName)
'###	Note : Depricated Function :	Note: Use Fn_SISW_UI_JavaTable_Operations
'################################################################################################################
Function Fn_Table_GetRowCount(sFunctionName,objJavaDialog, sTableName)
	bGblFailedFunctionName = sFunctionName
    Call Fn_WriteLogFile(Environment.Value("TestLogFile"),  "Depricated function : Fn_Table_GetRowCount > Calling Function Fn_SISW_UI_JavaTable_Operations")
	Fn_Table_GetRowCount = Fn_SISW_UI_JavaTable_Operations(sFunctionName, "GetRowCount", objJavaDialog , sTableName, "", "", "", "", "", "", "")
End Function
'#################################################################################################################
'###    FUNCTION NAME   :     Fn_UI_JavaStaticText_SearchAndClick(sFunctionName, objJavaDialog,sStaticText)

'###    DESCRIPTION     :   This function  is used to Search and Click The Static Text

'###    PARAMETERS      :   sFunctionName - Valid function name

'###                        objJavaDialog -  Valid Dialog/Window Path

'###                        sStaticText - Vaild static text name                     

'###    Function Calls  :   Fn_WriteLogFile

'###    HISTORY         :   AUTHOR            DATE        VERSION

'###    CREATED BY      :   Prasanna        30/04/2010     1.0

'###    REVIWED BY      :    Sameer			18/05/2010

'###    MODIFIED BY     : 	Sandeep		18/05/2010

'###    EXAMPLE         :  call Fn_Java_StaticText_Exist(" Fn_TcObjectDelete", objDialog, "More...",)
'################################################################################################################
Function Fn_UI_JavaStaticText_SearchAndClick(sFunctionName, objJavaDialog,sFormType) 

    Dim  intNoOfObjects, iCounter,bFlag
	Dim objSelectType, objJavaStaticText
	bGblFailedFunctionName = sFunctionName
	 sUIFail = sFunctionName + ">> Fn_UI_JavaStaticText_SearchAndClick >> " +  objJavaDialog.toString +">> " +  sFormType
	
	bFlag=False
		Set objSelectType=description.Create()
		objSelectType("Class Name").value = "JavaStaticText"
		Set intNoOfObjects = objJavaDialog.ChildObjects(objSelectType)
		For  iCounter = 0 to intNoOfObjects.count-1
				If  intNoOfObjects(iCounter).getROProperty("label") = sFormType Then
						intNoOfObjects(iCounter).Click 1,1
	   					bFlag=True
						Exit For
				End If
		 Next

				If bFlag=True Then
                	'log the success
					Call Fn_WriteLogFile( Environment.Value("TestLogFile"),"Successfully Searched and Clicked on JavaStaticText " & sFormType & " of Function " & sFunctionName)
					'Return True from Function
					Fn_UI_JavaStaticText_SearchAndClick = True
				Else
                    'log the failure when text not visible
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " JavaStaticText " &sFormType  & " is not Visible " & "of Function " & sFunctionName)
					Fn_UI_JavaStaticText_SearchAndClick = False
					Call ExitFromUI(sUIFail)
				End If

	'Object set to nothing
	Set objSelectType = Nothing
	Set objJavaStaticText = Nothing                        

End Function
'#################################################################################################################

'###    FUNCTION NAME   :     Fn_UI_JavaMenu_SearchAndSelect(sFunctionName, objJavaDialog,sFormType) 

'###    DESCRIPTION     :   This function  is used to Search and Click The Static Text

'###    PARAMETERS      :   sFunctionName - Valid function name

'###                        objJavaDialog -  Valid Dialog/Window Path
'###                        sStaticText - Valid java menu name                     

'###    Function Calls  :   Fn_WriteLogFile

'###    HISTORY         :   AUTHOR            DATE        VERSION

'###    CREATED BY      :   Sandeep        18/05/2010     1.0

'###    REVIWED BY      :     Sameer 			18/05/2010
'###    MODIFIED BY     :
'###    EXAMPLE         :  Call Fn_UI_JavaMenu_SearchAndSelect("Fn_BOMViewRev_SaveAs", JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("BOMViewSaveAs"),"CAEAnalysis") 
'################################################################################################################
Function Fn_UI_JavaMenu_SearchAndSelect(sFunctionName, objJavaDialog,sFormType) 

    Dim  intNoOfObjects, iCounter,i,bFlag
	Dim objSelectType
	bGblFailedFunctionName = sFunctionName
	 sUIFail = sFunctionName + ">> Fn_UI_JavaMenu_SearchAndSelect >> " +  objJavaDialog.toString +">> " +  sFormType
	
	bFlag=False
	Set objSelectTypeMenu=Description.Create()
					objSelectTypeMenu("Class Name").value = "JavaMenu"
					 Set intNoOfObjects = objJavaDialog.ChildObjects(objSelectTypeMenu)

					  For i = 0 to intNoOfObjects.count-1
						   If  intNoOfObjects(i).getROProperty("label") = sFormType Then
									intNoOfObjects(i).Select
									bFlag=True
									Exit for
							  End If
							Next

							If bFlag=True Then
									Call Fn_WriteLogFile( Environment.Value("TestLogFile"),"Successfully searched and selected  JavaMenu " & sFormType & " of Function " & sFunctionName)
									'Return True from Function
									Fn_UI_JavaMenu_SearchAndSelect = True
                            Else
									'log the failure when text not visible
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " JavaMenu " &sFormType  & " is not Visible " & "of Function " & sFunctionName)
									Fn_UI_JavaMenu_SearchAndSelect = False
									Call ExitFromUI(sUIFail)
							End If
	'Object set to nothing
	Set objSelectType = Nothing
	Set objJavaStaticText = Nothing                        

End Function

'##########################################################################################################################################################################
'Function Name	   :	Fn_UI_JavaMenu_Select(sFunctionName,objJavaDialog,sMenuPath)
'
'Description	   :	 Actions performed in this function are:
'	   						1. Menu Select
'							2. Menu multi-select
'				
'
'Parameters	   	   :	1. sFunctionName: Valid Function Name
'		 	    		2. objJavaDialog: Valid Java Dialog 
'			    		3. sMenuPath: Valid menu  name
'			    		
'
'Return Value	   : 	TRUE \ FALSE
'
'HISTORY       	   :  	     AUTHOR                      DATE    	    		 VERSION
'
'CREATED BY        :         Sandeep               	  30/04/2010  		   	       1.0
'
'MODIFIED BY        :   	 Koustubh               	  17/12/2010  		   	       1.0						Added case 
'
'REVIWED BY    	   :         Sameer							         
'"
'Examples	   	   :    Fn_UI_JavaMenu_Select("Fn_MenuItem_select_Operation",JavaWindow("StructureManager"),"File:New")
' 						
'##########################################################################################################################################################################

Function Fn_UI_JavaMenu_Select(sFunctionName,objJavaDialog,sMenuPath)
    Dim ArrMenu, NumObjects,iCounter
	Dim objJavaMenuSelect
	'Setting Object
	Set objJavaMenuSelect=objJavaDialog
	bGblFailedFunctionName = sFunctionName
	'Checking Object Exist Or Not
'		If 	objJavaMenuSelect.Exist Then
			
'				objJavaMenuSelect.WaitProperty "enabled", "1"
'					'Checking Object Enable or not
'					If objJavaMenuSelect.GetROProperty("enabled") = "1"  Then
								
						ArrMenu=Split(sMenuPath,":") 
						NumObjects = ubound(ArrMenu)
						
						Select Case NumObjects
								'	added by Koustubh
								Case "0"				
										objJavaMenuSelect.JavaMenu("label:="&ArrMenu(0)&"","index:=0").Select
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Menu selected succefully of Function" & sFunctionName)
										Fn_UI_JavaMenu_Select = TRUE		
								Case "1"				
										objJavaMenuSelect.JavaMenu("label:="&ArrMenu(0)&"","index:=0").JavaMenu("label:="&ArrMenu(1)&"","index:=0").Select
														
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Menu selected succefully of Function" & sFunctionName)
										Fn_UI_JavaMenu_Select = TRUE		
								Case "2"
										objJavaMenuSelect.JavaMenu("label:="&ArrMenu(0)&"","index:=0").JavaMenu("label:="&ArrMenu(1)&"","index:=0").JavaMenu("label:="&ArrMenu(2)&"","index:=0").Select
															
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Menu selected succefully of Function" & sFunctionName)
										Fn_UI_JavaMenu_Select = TRUE		
								 Case "3"
										objJavaMenuSelect.JavaMenu("label:="&ArrMenu(0)&"","index:=0").JavaMenu("label:="&ArrMenu(1)&"","index:=0").JavaMenu("label:="&ArrMenu(2)&"","index:=0").JavaMenu("label:="&ArrMenu(3)&"","index:=0").Select
															
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Menu selected succefully of Function" & sFunctionName)
										Fn_UI_JavaMenu_Select = TRUE
								Case "4"
										objJavaMenuSelect.JavaMenu("label:="&ArrMenu(0)&"","index:=0").JavaMenu("label:="&ArrMenu(1)&"","index:=0").JavaMenu("label:="&ArrMenu(2)&"","index:=0").JavaMenu("label:="&ArrMenu(3)&"","index:=0").JavaMenu("label:="&ArrMenu(4)&"","index:=0").Select
																	
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Menu selected succefully of Function" & sFunctionName)
										Fn_UI_JavaMenu_Select = TRUE
								Case Else
										
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Invalid Menu operation of Function" & sFunctionName)
										Fn_UI_JavaMenu_Select = FALSE
										Call ExitFromUI(sUIFail)
						End Select
'					Else
'						Fn_UI_JavaMenu_Select = FALSE
'						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), objJavaMenuSelect.toString & "FAIL:Window is not Enable of Function" & sFunctionName)
'						Call ExitFromUI(sUIFail)
'					End If
'		Else
'			Fn_UI_JavaMenu_Select = FALSE
'			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), objJavaMenuSelect.toString & "FAIL:Window is not Exist of Function" & sFunctionName)
'			Call ExitFromUI(sUIFail)
'		End If
	Set objJavaMenuSelect=Nothing
End Function


 

'################################################################################################################################################
 

'#######################################################################################################################
'###    FUNCTION NAME   :	Fn_UI_SetDateAndTime(sFunctionName,sDate,sTime)
'###
'###    DESCRIPTION     :   This function is use to Set Date and Time from Calender into Shell
'###
'###    PARAMETERS      :   1.sFunctionName: Valid Function Name
'###		 	   								2.sDate:Valid Date
'###			   								3.sTime:Valid Time
'###			   
'###			    
'###    RETURNVALUE   :   True/False
'###
'###    PRE-REQUISITES  :  Window should be open
'###
'###    HISTORY         :	AUTHOR			|	DATE		|  VERSION
'###
'###    CREATED BY      :	Sandeep 		|	03/05/2010	| 	1.0
'###
'###    REVIWED BY      :	Sameer 
'###
'###    MODIFIED BY     :	Koustubh Watwe |	12-09-2012	| Added code to click on Today Button
'###    EXAMPLE         :    Call Fn_UI_SetDateAndTime("Fn_MyTc_SearchDateSet","28_Feb_2008", "5:00:00 PM")
'###    EXAMPLE         :    Call Fn_UI_SetDateAndTime("Fn_MyTc_SearchDateSet","Today", "")
'###    EXAMPLE         :    Call Fn_UI_SetDateAndTime("Fn_MyTc_SearchDateSet","OK", "")
'########################################################################################################################
Function Fn_UI_SetDateAndTime(sFunctionName,sDate,sTime)
   	Dim  WshShell, iCnt
	Dim objDateControl
	'On Error Resume Next
	bGblFailedFunctionName = sFunctionName
	sUIFail = sFunctionName + ">> Fn_UI_SetDateAndTime >> " +  sDate +">> " +  sTime

	set WshShell = CreateObject("WScript.Shell")
    WshShell.SendKeys " "
	Set WshShell = nothing

	For iCnt = 1 to 30
		JavaWindow("MyTcShell").SetTOProperty "index", iCnt
		If JavaWindow("MyTcShell").JavaWindow("Date Control").Exist(1) Then
			Set objDateControl = JavaWindow("MyTcShell").JavaWindow("Date Control")
			Exit For
		End If
	Next
	
	'Checking Existance of object
	If objDateControl.Exist Then
		objDateControl.WaitProperty "enabled","1"
		'Checking object is enable or not
		If objDateControl.GetROProperty("enabled")="1" Then
			'Setting Date
			If lcase(sDate) = "today" Then
				objDateControl.JavaButton("Today").Click
			ElseIf  lcase(sDate) = "ok" Then
				objDateControl.JavaButton("OK").Click
			Else
				If sDate <> "" Then
					objDateControl.JavaCalendar("Date").SetDate sDate
				End If
				'Setting Time
				If sTime <> "" Then
					objDateControl.JavaCalendar("Time").SetTime sTime							
				End If
				objDateControl.JavaButton("OK").Click
			End If
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Date & Time is set  to"&sDate&sTime&"of Function"& sFunctionName)
			Fn_UI_SetDateAndTime=True
		Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), objDateControl.ToString &"is not Enable of Function " & sFunctionName)
			Fn_UI_SetDateAndTime=False
			Call ExitFromUI(sUIFail)
		End If
	Else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),  objDateControl.ToString&"is not Exist of Function " & sFunctionName)
		Fn_UI_SetDateAndTime=False 
		Call ExitFromUI(sUIFail)
	End If
	Set objDateControl = Nothing
End Function
'#############################################################################################################################################################
'###    FUNCTION NAME   :    Function Fn_UI_ObjectPressKey(sFunctionName,objDialog,sKey,sModifier)
'###
'###    DESCRIPTION     :   This function  is used PressKey Method
'###
'###    PARAMETERS      :   sFunctionName - Valid function name
'###                        objJavaDialog - Valid Java Dialog Name
'###                        vKey		  - Vaild Key Name(This Parameter is Varient Either variable or String )
'###                        vModifier     - Valid Modifier (Optional Parameter IN Function)&&(This Parameter is Varient Either variable or String )
'###                        
'###                        
'###    Function Calls  :   Fn_WriteLogFile ()
'###
'###    HISTORY         :   AUTHOR                   DATE        VERSION
'###
'###    CREATED BY      :   Sandeep				   05/05/2010	  1.0
'###
'###    REVIWED BY      :   Sameer					10/05/2010	
'###
'###    MODIFIED BY     :
'###    EXAMPLE         :1.Call Function Fn_UI_ObjectPressKey("Fn_MyTc_DatasetSearch",JavaWindow("MyTeamcenter").JavaEdit("CreatedAfterDt"),vbTab,"")
'###					 2.Call Function Fn_UI_ObjectPressKey("Fn_MyTc_DatasetSearch",JavaWindow("MyTeamcenter").JavaEdit("CreatedAfterDt"),vbTab,micAlt)
'#############################################################################################################################################################
''vKey :-This parameter is Compulsory
''vModifier :-This parameter is "Optional" parameter
Function Fn_UI_ObjectPressKey(sFunctionName,objDialog,vKey,vModifier)

sUIFail = sFunctionName + ">> Fn_UI_ObjectPressKey >> " +  objDialog.toString +">> " +  vKey
	bGblFailedFunctionName = sFunctionName
	'Checking Existance of Object
	If objDialog.Exist Then
		objDialog.WaitProperty "enabled","1"
		'Checking Whether object is Enable OR Disable
		If objDialog.GetROProperty("enabled")="1" OR objDialog.GetROProperty("enabled")=True Then
				If vModifier<>"" Then
					objDialog.PressKey vKey,vModifier
				 Else
					objDialog.PressKey vKey
				End If
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Sucessfully Press Key" & Cstr(vKey) &" On " & objDialog.toString &"of Function " & sFunctionName)
			Fn_UI_ObjectPressKey=True
		Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), objDialog.toString & " Dialog is Disable of Function " & sFunctionName)
			Fn_UI_ObjectPressKey=False
			Call ExitFromUI(sUIFail)
		End If
	Else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), objDialog.toString & " Dialog does not Exist of Function " & sFunctionName)
		Fn_UI_ObjectPressKey=False
	End If
End Function
'####################################################################################################################################################
'###    FUNCTION NAME   :   Function Fn_UI_EditBox_Type(sFunctionName,objJavaDialog,sEditBoxName,sStringToType)
'###
'###    EXAMPLE         : Call Fn_UI_EditBox_Type("Fn_TeamcenterLogin",JavaWindow("Teamcenter Login"),"User ID:","admin")
'#########################################################################################################################################################
Function Fn_UI_EditBox_Type(sFunctionName,objJavaDialog,sEditBoxName,sStringToType)
	bGblFailedFunctionName = sFunctionName
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Depricated function Fn_UI_EditBox_Type : Calling function Fn_SISW_UI_JavaEdit_Operations")
	Fn_UI_EditBox_Type = Fn_SISW_UI_JavaEdit_Operations(sFunctionName, "Type", objJavaDialog, sEditBoxName, sStringToType)
End Function
'############################################################################################################################################################
'###    FUNCTION NAME   :   Fn_JavaTree_Node_Activate()
'###
'###    DESCRIPTION     :   This Function is used to Activate Node/Element from the JavaTree.
'###
'###    PARAMETERS      :   sFunctionName	: Valid Function name,
'### 			    							objJavaDialog	: Valid Dialog Path,
'### 			    							sTreeName		: Valid Javatree Name,
'### 			   							 sItemNameToActivate 	: Valid Node/Element to be Activated ( Element seperated by :)
'###
'###    Function Calls  :   Fn_WriteLogFile()
'###
'###    HISTORY         :   AUTHOR      DATE        VERSION
'###	
'###    CREATED BY      :   sandeep  10/05/10		1.0
'###
'###    REVIWED BY      :  
'###
'###    MODIFIED BY     :   NA
'###    EXAMPLE         :   Call Fn_JavaTree_Node_Activate("Test",JavaWindow("MyTeamcenter"),"NavTree","Home:Automation")
'###############################################################################################################################################################	
Function Fn_JavaTree_Node_Activate(sFunctionName,objJavaDialog,sTreeName,sItemNameToActivate)
	Dim iNodeCount,iCounter,aItemName,sItemName,bFlag
	Dim objJavaTree
	bGblFailedFunctionName = sFunctionName
	sUIFail = sFunctionName + ">> Fn_JavaTree_Node_Activate >> " +  objJavaDialog.toString +">> " +  sTreeName
	
'Split Item Name into array
	aItemName=Split(sItemNameToActivate,":")
	bFlag=False
'Setting java tree object
   	Set objJavaTree=objJavaDialog.JavaTree(sTreeName)
'Checking Existance of java tree
		If objJavaTree.Exist Then
'Apllying Wait property
			objJavaTree.WaitProperty "enabled","1"
'Checking Java tree is Enable or Not 
				If objJavaTree.GetROProperty("enabled")="1" Then
'Retriving Number of Items present in tree 
					iNodeCount=objJavaTree.GetROProperty("items count")
						For iCounter=0 To Cint(iNodeCount)-1
							sItemName=objJavaTree.GetItem(iCounter)
'Checking passe item present in tree or not					
								If  sItemName=sItemNameToActivate Then
									bFlag=True
									Exit For
								ElseIf replace(sItemName," ","")=sItemNameToActivate Then	
									bFlag=True
									Exit For
                                End If
						Next		
								If bFlag=True Then
									objJavaTree.Activate sItemNameToActivate
	                                Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Sucessfully Activated"& sItemNameToActivate & " Node of Function " & sFunctionName)
									Fn_JavaTree_Node_Activate=True
								Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),sItemNameToActivate &  "Node not found of Function " & sFunctionName)
									Fn_JavaTree_Node_Activate=False
									Call ExitFromUI(sUIFail)
								End If
				
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),sTreeName & "tree Disable of Function " & sFunctionName)
					Fn_JavaTree_Node_Activate=False
					Call ExitFromUI(sUIFail)
				End If
				
		Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),sTreeName &  "tree is nit Exist of Function " & sFunctionName)
			Fn_JavaTree_Node_Activate=False
			Call ExitFromUI(sUIFail)
		End If
   Set onjJavaTree=Nothing
End Function
'######################################################################################################################################################



'#################################################################################################################
'###    FUNCTION NAME   :   Fn_UI_WebTable_GetCellData(sFunctionName,objJavaDialog,sTableName,iRow,iCol)
'###
'###    DESCRIPTION     :   Read Data of  The Selected Cell From Web Table
'###
'###    PARAMETERS      :   sFunctionName - Valid function name
'###                        objJavaDialog - Valid Java Dialog Name
'###                        sTableName - Vaild JavaTable name
'###                       	iRow  - Row no. from which data is to be selected
'###                        iCol  - Column no. from which data is to be selected
'###                        
'###    Function Calls  :   Fn_WriteLogFile ()
'###
'###    HISTORY   	    :   AUTHOR            DATE          VERSION
'###
'###    CREATED BY      :   Sandeep         10/05/10		  1.0
'###
'###    REVIWED BY      :   
'###
'###    MODIFIED BY     :
'###    EXAMPLE         :   Call Fn_UI_WebTable_GetCellData("Fn_itemBasicCreate", Browser("Mercury Tours").Page("Search Results"),"OutboundFlight", 1,1)
'################################################################################################################
'objDialogTable is may be web path also(E.g:-Browser("Mercury Tours").Page("Search Results"))
Function Fn_UI_WebTable_GetCellData(sFunctionName,objJavaDialog,sTableName,iRow,iCol)
	Dim iRowCount,iColCount
	Dim objDialogTable
	bGblFailedFunctionName = sFunctionName
	sUIFail = sFunctionName + ">> Fn_UI_WebTable_GetCellData >> " +  objJavaDialog.toString +">> " +  sTableName

	
'Setting Java Table object
	Set objDialogTable=objJavaDialog.WebTable(sTableName)


'Checking Existance of web Table	
			If objDialogTable.Exist Then
					iRowCount=objDialogTable.RowCount
					iColCount=objDialogTable.ColumnCount(1)
					If iRow <= iRowCount  And  iCol <= iColCount Then
								'Returnig the specific cell data from web table	
								Fn_UI_WebTable_GetCellData=objDialogTable.GetCellData(iRow,iCol)
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Sucessfully Selected Celldata " &  iRow &"," & iCol &"of the WebTable" & sTableName & " of Function " &sFunctionName)
						 Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Incorrect Column Number  " &  iCol &" or row number " & iRow & "of the WebTable" & sTableName & " of Function " &sFunctionName)
								Fn_UI_WebTable_GetCellData=False
								Call ExitFromUI(sUIFail)
						End If
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),sTableName & "Web table is not Exist  of Function " &sFunctionName)
				Fn_UI_WebTable_GetCellData=False
				Call ExitFromUI(sUIFail)
			End If
'Setting Java table object to nothing			
	Set objDialogTable=Nothing
End Function
'#################################################################################################################
'###    FUNCTION NAME   :   Fn_UI_JavaTable_RightClickCell(sFunctionName,objJavaDialog,sTableName,iRow,vCol,sMoseButton,sModifier)
'###
'###    DESCRIPTION     :   This function is used to right click on JavaTable
'###
'###    PARAMETERS      :   sFunctionName - Valid function name
'###                        objJavaDialog - Valid Java Dialog Name
'###                        sTableName - Vaild JavaTable name
'###                       	iRow  - Row no. from which data is to be selected
'###                        vCol  - Column no. from which data is to be selected or Column Name
'###                        sMoseButton-Valid Mouse button Name
'###						sModifier-Valid modifier name
'###    Function Calls  :   Fn_WriteLogFile ()
'###
'###    HISTORY   	    :   AUTHOR            DATE          VERSION
'###
'###    CREATED BY      :   Sandeep       11/05/10		  1.0
'###
'###    REVIWED BY      :   
'###
'###    MODIFIED BY     :
'###    EXAMPLE         :   Call Fn_UI_JavaTable_RightClickCell("Fn_itemBasicCreate", JavaWindow("DefaultWindow"),"Details Table", 1,"Options","RIGHT")
'################################################################################################################

'***vCol:-This parameter is Compulsory and varient Either Column Name or Column Number
Function Fn_UI_JavaTable_CellRightClick(sFunctionName,objJavaDialog,sTableName,iRow,vCol,sMouseButton,sModifier)
Dim objJavaTableRightClickCell
	bGblFailedFunctionName = sFunctionName
'Setting Java Table Object
Set objJavaTableRightClickCell=objJavaDialog.JavaTable(sTableName)
		'Checking Java table exist or Not
		If objJavaTableRightClickCell.Exist Then
			'Applying Wait Property on object
			objJavaTableRightClickCell.WaitProperty "enabled","1"
				'Checking Object is enable or not
				If objJavaTableRightClickCell.GetROProperty("enabled")="1" Then
					'Checking iRow parameter is correctly pass or not
					If  Cint(iRow) <= Cint(objJavaTableRightClickCell.GetROProperty("rows")) Then
						If  sMouseButton<>"" AND sModifier<>"" Then
							objJavaTableRightClickCell.ClickCell iRow,vCol,sMouseButton,sModifier
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Succesfully Click on Cell"& iRow &" for Function " & sFunctionName)
							Fn_UI_JavaTable_CellRightClick=True
						Elseif sMouseButton<>"" AND sModifier="" Then
							objJavaTableRightClickCell.ClickCell iRow,vCol,sMouseButton
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Succesfully Click on Cell"& iRow &" for Function " & sFunctionName)
							Fn_UI_JavaTable_CellRightClick=True
						Else
							objJavaTableRightClickCell.ClickCell iRow,vCol
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Succesfully Click on Cell"& iRow &" for Function " & sFunctionName)
							Fn_UI_JavaTable_CellRightClick=True
						End If
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), iRow &"row does not exist in"& sTableName &" table for Function " & sFunctionName)
						Fn_UI_JavaTable_CellRightClick=False
						Call ExitFromUI(sUIFail)
					End If
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sTableName &" table is disable for Function " & sFunctionName)
					Fn_UI_JavaTable_CellRightClick=False
					Call ExitFromUI(sUIFail)
				End If
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sTableName &" table does not exist for Function " & sFunctionName)
				Fn_UI_JavaTable_CellRightClick=False
				Call ExitFromUI(sUIFail)
		End If
'Setting object to nothing		
Set objJavaTableRightClickCell=Nothing
End Function
'#######################################################################################################################################
'############################################################################################################################################################
'###    FUNCTION NAME   :   Fn_JavaTree_NodeIndex(sFunctionName,objJavaDialog,sTreeName,sNodeForIndex)
'###
'###    DESCRIPTION     :   This Function is used to Retrieve The index of Node
'###
'###    PARAMETERS      :   sFunctionName	: Valid Function name,
'### 			    							objJavaDialog	: Valid Dialog Path,
'### 			    							sTreeName		: Valid Javatree Name,
'### 			   							 sNodeForIndex 	: Valid Node/Element for which we need Index
'###
'###    Function Calls  :   Fn_WriteLogFile()
'###
'###    HISTORY         :   AUTHOR      DATE        VERSION
'###	
'###    CREATED BY      :   Sandeep  20/05/10		1.0
'###
'###    REVIWED BY      :  	Sameer	20/05/10		
'###
'###    MODIFIED BY     :   NA
'###    EXAMPLE         :   Call Fn_JavaTree_NodeIndex("Test",JavaWindow("MyTeamcenter"),"NavTree","Home:AutomatedTests")
'###############################################################################################################################################################	
'
'
'		NOTE : USE Fn_JavaTree_NodeIndexExt INSTEAD OF THIS FUNCTION
'		BY : KOUSTUBH WATWE
'
'###############################################################################################################################################################	

Function Fn_JavaTree_NodeIndex(sFunctionName,objJavaDialog,sTreeName,sNodeForIndex)
  Dim iItemCount,iCnt,bFlag,iTempCnt
  Dim objJavaTreeIndexNode
  	bGblFailedFunctionName = sFunctionName
  	sUIFail = sFunctionName + ">> Fn_JavaTree_NodeIndex >> " +  objJavaDialog.toString +">> " +  sTreeName
  
  bFlag=False
 'Setting object of java tree
  Set objJavaTreeIndexNode=objJavaDialog.JavaTree(sTreeName)
		 'Checking existance of object
		  If objJavaTreeIndexNode.Exist Then
			  'Applying wait property
			  objJavaTreeIndexNode.WaitProperty "enabled","1"
					If objJavaTreeIndexNode.GetROProperty("enabled")="1" Then
							'Retriving  itemcount of java tree
							iItemCount=objJavaTreeIndexNode.GetROProperty("items count")
									For iCnt=0 to iItemCount-1
										'Checking node is present in tree or not
										If objJavaTreeIndexNode.GetItem(iCnt)=sNodeForIndex  Then
												iTempCnt=iCnt
												bFlag=True
												Exit For
										End If
									Next
													
							If  bFlag=True Then
										Fn_JavaTree_NodeIndex=iTempCnt
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sNodeForIndex & "nodes Index is" & iTempCnt & "  For Function " & sFunctionName)
							Else
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sNodeForIndex & "node is not present in " & sTreeName & "Tree  For Function " & sFunctionName)
										Fn_JavaTree_NodeIndex=False
										Call ExitFromUI(sUIFail)
							End If
					Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sTreeName & "Tree  is disable For Function " & sFunctionName)
							Fn_JavaTree_NodeIndex=False
							Call ExitFromUI(sUIFail)
					End If
		  Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sTreeName & "Tree  is not Exist For Function " & sFunctionName)
					Fn_JavaTree_NodeIndex=False
					Call ExitFromUI(sUIFail)
		  End If
	Set objJavaTreeIndexNode=Nothing
End Function   

'#############################################################################################################################################################
'#################################################################################################################
'###    FUNCTION NAME   :    Fn_UI_Object_SetTOProperty_ExistCheck(sFunctionName,objJavaDialog,sProperty,sPropValue)
'###
'###    DESCRIPTION     :   This function  is used to Set TO Property For given Object BEFORE cheking existance of  given Object
'###
'###    PARAMETERS      :   sFunctionName - Valid function name
'###                        					objJavaDialog - Valid Java Dialog Name
'###                      					  	sButtonName- Vaild Button  Name
'###                              				sProperty-Valid Property
'###                        					sPropValue-New value
'###                        
'###    Function Calls  :   Fn_WriteLogFile()
'###
'###    HISTORY         :   AUTHOR                   DATE        VERSION
'###
'###    CREATED BY      : Sagar Shivade				10/06/2010
'###
'###    REVIWED BY      : Sameer					10/06/2010
'###
'###    MODIFIED BY     :
'###    EXAMPLE         :Fn_UI_Object_SetTOProperty_ExistCheck("Fn_SetPerspective",JavaWindow("DefaultWindow").JavaButton("Click"),"Label","NewLabel")
'###										
'################################################################################################################

Function Fn_UI_Object_SetTOProperty_ExistCheck(sFunctionName,objJavaDialog,sProperty,sPropValue)
	Dim objSetTOProperty
	bGblFailedFunctionName = sFunctionName
		sUIFail = sFunctionName + ">> Fn_UI_Object_SetTOProperty_ExistCheck >> " +  objJavaDialog.toString +">> " +  sProperty
	
	'Setting object to Object  variable 
	Set objSetTOProperty=objJavaDialog

 ' set properties  of objects before checking  existance
	objSetTOProperty.SetTOProperty sProperty,sPropValue

	'Checking existance of object
	If objSetTOProperty.Exist Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Set the TO Property "&sProperty &" as " & sPropValue &" for " & objJavaDialog.toString & " of Function " &sFunctionName)
			Fn_UI_Object_SetTOProperty_ExistCheck=True
	Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),objJavaDialog.toString & " Does not Exist after setting Property  " &sProperty & " to value " &  sPropValue & " of Functiopn "& sFunctionName)
			Fn_UI_Object_SetTOProperty_ExistCheck=False
			Call ExitFromUI(sUIFail)
	End If

	Set objSetTOProperty = Nothing 
End Function
'#########################################################################################################
'###
'###    FUNCTION NAME   :  Fn_UI_ListItemExist(sFunctionName, objJavaDialog, sJavaList,sElementToSelect)
'###	Note : Depricated Function :	Note: Use Fn_SISW_UI_JavaList_Operations
'###    EXAMPLE         : Fn_UI_ListItemExist("", JavaWindow("StructureManager").JavaApplet("PSEApplet").JavaDialog("SearchReferenceDesignators"), "ReferenceDesignatorsList","000110/A;1-Part1")
'#############################################################################################################
Function Fn_UI_ListItemExist(sFunctionName, objJavaDialog, sJavaList,sElementToSelect)		
	bGblFailedFunctionName = sFunctionName
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Depricated function : Fn_UI_ListItemExist > Calling Function Fn_SISW_UI_JavaList_Operations")
	Fn_UI_ListItemExist =  Fn_SISW_UI_JavaList_Operations(sFunctionName, "Exist", objJavaDialog, sJavaList, sElementToSelect, "", "@")
End Function						
'*********************************************************	Generic function to get tree index ***********************************************************************

'Function Name		        :   Fn_JavaTree_NodeIndexExt

'Description			    :	Generic function to get tree index.

'Parameters			   :	1. sFunctionName - caller function's name
'							2. objJavaDialog - parent object
'							3. sTreeName - Tree name
'							4. sNode - Node to get index
'							5. sSeparator - Node Separator,  Default is ':'
'								       We should use seperator other than ':' in case if node text contains ':'
'							6. sInstanceHandler - Instance operator, Default is '@'
'								       We should use instance handler other than '@' in case if node text contains '@'
											
'Return Value		    : 	Node index / -1 (in case of False)

'Pre-requisite			:	Tree object should be visible.

'Examples				:         'sNode = "Home~KouTest~test2 $2"
							        'msgbox Fn_JavaTree_NodeIndexExt("Caller Function Name",JavaWindow("MyTeamcenter"),"NavTree", sNode,"~", "$")
							        'sNode = "Home:KouTest:test1 @2:test2 @1"
							        'msgbox Fn_JavaTree_NodeIndexExt("Caller Function Name",JavaWindow("MyTeamcenter"),"NavTree", sNode, "", "")
							        'sNode =  "-995:?????~???:????? @3"
							        'msgbox Fn_JavaTree_NodeIndexExt("Fn_ClassAdmin_KeyLOVTreeOperations",JavaWindow("ClassAdminMainWin").JavaApplet("ClassAdminApplet"),"LOVKeyTree", sNode, "~", "")
'History:
'			Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'			Koustubh 				16-Dec-2010			   1.0
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public function Fn_JavaTree_NodeIndexExt(sFunctionName, objJavaDialog, sTreeName, sNode, sSeparator, sInstanceHandler)
	Dim iRowCounter, iRows, aNodes, sPath, iCnt, jCnt, iInstance, aNodeArr
	Dim iInstanceCnt, aDummyArr, objTree
	bGblFailedFunctionName = sFunctionName
    Set objTree = objJavaDialog.JavaTree(sTreeName)
	Fn_JavaTree_NodeIndexExt = -1
	If sSeparator = ""  Then sSeparator = ":"
	If sInstanceHandler = "" Then sInstanceHandler = "@"
	iInstanceCnt = 1
	If objTree.exist(10) = False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ " & sFunctionName & " ] Specified Tree object [ " & sTreeName & " ] does not exist.")
		Exit function
	End If
	If cInt(objTree.GetROProperty("enabled")) <> 1 Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ " & sFunctionName & " ] Specified Tree object [ " & sTreeName & " ] is not visible.")
		Exit function
	End If
	iRows = cInt(objTree.GetROProperty("items count"))
	iRowCounter = 0
	If instr(sNode, sInstanceHandler) > 0 Then
		' multiple instances
                    aNodes = split(sNode, sSeparator)

		If UBound( aNodes ) <> 0 Then
			'path with multiple instance handler
			For iCnt = 0 to ubound(aNodes)
				If iRowCounter = iRows Then
					' node not found
					Exit FOR
				End If
				aNodeArr = split(aNodes(iCnt), sInstanceHandler)
				iInstanceCnt = 1
				If UBound(aNodeArr) = 0 Then
					iInstance = 1
				Else
					iInstance = cInt(trim(aNodeArr(1)))
				End If
				sPath = ""
				'generating node path
                                        For jCnt = 0 to iCnt
                                                  aDummyArr = split(aNodes(jCnt), sInstanceHandler)
					If sPath = "" Then
						sPath = trim(aDummyArr(0))
					Else
						sPath = sPath & ":" & trim(aDummyArr(0))
					End If
				Next
				'verifiyig path
                                        Do While iRowCounter < iRows
					If objTree.GetItem(iRowCounter) = sPath Then
						If iInstanceCnt = iInstance Then
							iRowCounter = iRowCounter +1
							Exit do
						End If
						iInstanceCnt = iInstanceCnt + 1
					End If
					iRowCounter = iRowCounter +1
				loop
			Next
			If iRowCounter < iRows Then
				Fn_JavaTree_NodeIndexExt = iRowCounter - 1
			End If
		Else
			'With instance with no child items
			aNodeArr = split(aNodes(0), sInstanceHandler)
			sPath = trim(aNodeArr(0))
			iInstance = cInt(trim(aNodeArr(1)))
			For iRowCounter = 0 to iRows - 1
				If objTree.GetItem(iRowCounter) = sPath Then
					If iInstanceCnt = iInstance Then
						Fn_JavaTree_NodeIndexExt = iRowCounter
						Exit for
					End If
					iInstanceCnt = iInstanceCnt + 1
				End If
			Next
		End If
	Else
		' normal path without instance handler
		sPath = replace(sNode,sSeparator,":")
		For iRowCounter = 0 to iRows - 1
			If objTree.GetItem(iRowCounter) = sPath Then
				Fn_JavaTree_NodeIndexExt = iRowCounter
				Exit for
			End If
		Next
	End If
          Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: [ " & sFunctionName & " ] Function Fn_JavaTree_NodeIndexExt executed successfully.")
End Function
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'##########################################################################################################################################################################
'Function Name	   :	Fn_UI_JavaMenu_Exist(sFunctionName,objJavaDialog,sMenuPath)
'
'Description	   :	 Actions performed in this function are:
'	   						1. Menu Select
'							2. Menu multi-select
'				
'
'Parameters	   	   :	1. sFunctionName: Valid Function Name
'		 	    		2. objJavaDialog: Valid Java Dialog 
'			    		3. sMenuPath: Valid menu  name
'			    		
'
'Return Value	   : 	TRUE \ FALSE
'
'HISTORY       	   :  	     AUTHOR                      DATE    	    		 VERSION
'
'CREATED BY        :         Koustubh               	 17/12/2010  		   	       1.0
'
'REVIWED BY    	   :						         
'"
'Examples	   	   :    Fn_UI_JavaMenu_Exist("Fn_MenuItem_select_Operation",JavaWindow("StructureManager"),"File:New")
' 						
'##########################################################################################################################################################################

Function Fn_UI_JavaMenu_Exist(sFunctionName,objJavaDialog,sMenuPath)
    Dim ArrMenu, NumObjects,iCounter
	Dim objJavaMenuSelect
	bGblFailedFunctionName = sFunctionName
	Fn_UI_JavaMenu_Exist = False
	'Setting Object
	Set objJavaMenuSelect=objJavaDialog
	ArrMenu=Split(sMenuPath,":") 
	NumObjects = ubound(ArrMenu)
	Select Case NumObjects
			Case "0"
					Fn_UI_JavaMenu_Exist = objJavaMenuSelect.JavaMenu("label:="&ArrMenu(0)&"","index:=0").exist(5)
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Menu verified succefully of Function" & sFunctionName)
			Case "1"
					if objJavaMenuSelect.JavaMenu("label:="&ArrMenu(0)&"","index:=0").exist(5) then
						objJavaMenuSelect.JavaMenu("label:="&ArrMenu(0)&"","index:=0").Select
					Else
						Set objJavaMenuSelect=Nothing
						Exit function
					End if 
					Fn_UI_JavaMenu_Exist = objJavaMenuSelect.JavaMenu("label:="&ArrMenu(0)&"","index:=0").JavaMenu("label:="&ArrMenu(1)&"","index:=0").exist(5)
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Menu verified succefully of Function" & sFunctionName)
			Case "2"
					if objJavaMenuSelect.JavaMenu("label:="&ArrMenu(0)&"","index:=0").exist(5) then
						objJavaMenuSelect.JavaMenu("label:="&ArrMenu(0)&"","index:=0").Select
					Else
						Set objJavaMenuSelect=Nothing
						Exit function
					End if 
					if objJavaMenuSelect.JavaMenu("label:="&ArrMenu(0)&"","index:=0").JavaMenu("label:="&ArrMenu(1)&"","index:=0").exist(5) then
						objJavaMenuSelect.JavaMenu("label:="&ArrMenu(0)&"","index:=0").JavaMenu("label:="&ArrMenu(1)&"","index:=0").Select
					Else
						Set objJavaMenuSelect=Nothing
						Exit function
					End if 
					Fn_UI_JavaMenu_Exist = objJavaMenuSelect.JavaMenu("label:="&ArrMenu(0)&"","index:=0").JavaMenu("label:="&ArrMenu(1)&"","index:=0").JavaMenu("label:="&ArrMenu(2)&"","index:=0").exist(5)
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Menu verified succefully of Function" & sFunctionName)
			 Case "3"
					if objJavaMenuSelect.JavaMenu("label:="&ArrMenu(0)&"","index:=0").exist(5) then
						objJavaMenuSelect.JavaMenu("label:="&ArrMenu(0)&"","index:=0").Select
					Else
						Set objJavaMenuSelect=Nothing
						Exit function
					End if 
					if objJavaMenuSelect.JavaMenu("label:="&ArrMenu(0)&"","index:=0").JavaMenu("label:="&ArrMenu(1)&"","index:=0").exist(5) then
						objJavaMenuSelect.JavaMenu("label:="&ArrMenu(0)&"","index:=0").JavaMenu("label:="&ArrMenu(1)&"","index:=0").Select
					Else
						Set objJavaMenuSelect=Nothing
						Exit function
					End if 
					if objJavaMenuSelect.JavaMenu("label:="&ArrMenu(0)&"","index:=0").JavaMenu("label:="&ArrMenu(1)&"","index:=0").JavaMenu("label:="&ArrMenu(2)&"","index:=0").exist(5) then
						objJavaMenuSelect.JavaMenu("label:="&ArrMenu(0)&"","index:=0").JavaMenu("label:="&ArrMenu(1)&"","index:=0").JavaMenu("label:="&ArrMenu(2)&"","index:=0").Select
					Else
						Set objJavaMenuSelect=Nothing
						Exit function
					End If
					Fn_UI_JavaMenu_Exist = objJavaMenuSelect.JavaMenu("label:="&ArrMenu(0)&"","index:=0").JavaMenu("label:="&ArrMenu(1)&"","index:=0").JavaMenu("label:="&ArrMenu(2)&"","index:=0").JavaMenu("label:="&ArrMenu(3)&"","index:=0").Exist(5)
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Menu verified succefully of Function" & sFunctionName)
			Case "4"
					if objJavaMenuSelect.JavaMenu("label:="&ArrMenu(0)&"","index:=0").exist(5) then
						objJavaMenuSelect.JavaMenu("label:="&ArrMenu(0)&"","index:=0").Select
					Else
						Set objJavaMenuSelect=Nothing
						Exit function
					End if 
					if objJavaMenuSelect.JavaMenu("label:="&ArrMenu(0)&"","index:=0").JavaMenu("label:="&ArrMenu(1)&"","index:=0").exist(5) then
						objJavaMenuSelect.JavaMenu("label:="&ArrMenu(0)&"","index:=0").JavaMenu("label:="&ArrMenu(1)&"","index:=0").Select
					Else
						Set objJavaMenuSelect=Nothing
						Exit function
					End if 
					if objJavaMenuSelect.JavaMenu("label:="&ArrMenu(0)&"","index:=0").JavaMenu("label:="&ArrMenu(1)&"","index:=0").JavaMenu("label:="&ArrMenu(2)&"","index:=0").exist(5) then
						objJavaMenuSelect.JavaMenu("label:="&ArrMenu(0)&"","index:=0").JavaMenu("label:="&ArrMenu(1)&"","index:=0").JavaMenu("label:="&ArrMenu(2)&"","index:=0").Select
					Else
						Set objJavaMenuSelect=Nothing
						Exit function
					End If
					if objJavaMenuSelect.JavaMenu("label:="&ArrMenu(0)&"","index:=0").JavaMenu("label:="&ArrMenu(1)&"","index:=0").JavaMenu("label:="&ArrMenu(2)&"","index:=0").JavaMenu("label:="&ArrMenu(3)&"","index:=0").exist(5) then
						objJavaMenuSelect.JavaMenu("label:="&ArrMenu(0)&"","index:=0").JavaMenu("label:="&ArrMenu(1)&"","index:=0").JavaMenu("label:="&ArrMenu(2)&"","index:=0").JavaMenu("label:="&ArrMenu(3)&"","index:=0").Select
					Else
						Set objJavaMenuSelect=Nothing
						Exit function
					End If
					Fn_UI_JavaMenu_Exist = objJavaMenuSelect.JavaMenu("label:="&ArrMenu(0)&"","index:=0").JavaMenu("label:="&ArrMenu(1)&"","index:=0").JavaMenu("label:="&ArrMenu(2)&"","index:=0").JavaMenu("label:="&ArrMenu(3)&"","index:=0").JavaMenu("label:="&ArrMenu(4)&"","index:=0").Exist(5)
												
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Menu verified succefully of Function" & sFunctionName)
			Case Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Invalid Menu operation of Function" & sFunctionName)
					Call ExitFromUI(sUIFail)
	End Select
	Set objJavaMenuSelect=Nothing
End Function
''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'###   FUNCTION NAME   :   Fn_UI_JavaTable_CheckColumnExists(sFunctionName, objJavaDialog, sJavaTableName,sColName)
'###	Note : Depricated Function :	Note: Use Fn_SISW_UI_JavaTable_Operations
'###    EXAMPLE          : 	Fn_UI_JavaTable_CheckColumnExists("Fn_UI_JavaTable_CheckColumnExists", JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Common Modifiable Properties"), "PropertyTable","Description")
''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Function Fn_UI_JavaTable_CheckColumnExists(sFunctionName, objJavaDialog, sJavaTableName,sColName)
	bGblFailedFunctionName = sFunctionName
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"),  "Depricated function : Fn_UI_JavaTable_CheckColumnExists > Calling Function Fn_SISW_UI_JavaTable_Operations")
	Fn_UI_JavaTable_CheckColumnExists = Fn_SISW_UI_JavaTable_Operations(sFunctionName, "GetColumnIndex", objJavaDialog , sJavaTableName, "", sColName, "", "", "", "","")
End function
''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''/$$$$   FUNCTION NAME   :   Fn_UI_JavaTable_ClickColumnHeader(sFunctionName, objJavaDialog, sJavaTableName,sColName,sAction,sMenu)
''/$$$$   Note : Depricated Function :	Note: Use Fn_SISW_UI_JavaTable_Operations
''/$$$$    EXAMPLE          : 	bReturn=	Fn_UI_JavaTable_ClickColumnHeader("Fn_UI_JavaTable_ClickColumnHeader", JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Common Modifiable Properties"), "PropertyTable","Name","LEFT","")
''/$$$$										bReturn=	Fn_UI_JavaTable_ClickColumnHeader("Fn_UI_JavaTable_ClickColumnHeader", JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Audit Log"), "LogTable","Object Type","RIGHT","Print Table:Graphics")
''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Function Fn_UI_JavaTable_ClickColumnHeader(sFunctionName, objJavaDialog, sJavaTableName,sColName,sAction,sMenu)
	bGblFailedFunctionName = sFunctionName
	If lcase(sAction) = "right" Then
		Fn_UI_JavaTable_ClickColumnHeader = Fn_SISW_UI_JavaTable_Operations(sFunctionName, "RMB_MenuSelect_On_ColumnHeader", objJavaDialog , sJavaTableName, sTableType, sColName, "", "", "", sMenu, "")
	Else
		Fn_UI_JavaTable_ClickColumnHeader = Fn_SISW_UI_JavaTable_Operations(sFunctionName, "SelectColumnHeader", objJavaDialog , sJavaTableName, sTableType, sColName, "", "", "", sMenu, "")
	End If
End function
'#################################################################################################################
'###    FUNCTION NAME   :    Fn_UI_JavaTree_NodeExist(sFunctionName,sTreeObject,sNodeNameWithPath)
'###
'###    DESCRIPTION     :   This function  is used to  check the Existance of the node in Tree
'###
'###    PARAMETERS      :   sFunctionName:				 Name of the Function In which it is implemented
'###											objJavaTree:					 The Object of the Tree
'###											sNodeNameWithPath: Full Path till the node from the Parent
'###             
'###    Function Calls  :   Fn_KillProcess(sProcessToKill)
'###
'###    CREATED BY      :   Harshal Agrawal			14 Feb 2011
'###
'###	REViEWED BY		: Ketan Raje
'###
'###   PRE - REQUISITE:  If we need  to Check the Existance of  'Nth' Node then till N -1  level  node should be Expanded.
'###
'###    MODIFIED BY     :
'###
'### 	 EXAMPLE         :    MsgBox Fn_UI_JavaTree_NodeExist("Fn_UI_JavaTree_NodeExist",JavaWindow("MyTeamcenter").JavaTree("NavTree"),"Home:001864-001864_a1:001864/A;1-001864_a1:View:001866-001866_a2")
'################################################################################################################
Function Fn_UI_JavaTree_NodeExist(sFunctionName,objJavaTree,sNodeNameWithPath)
	Dim intNodeCount,intCount,sTreeItem
	bGblFailedFunctionName = sFunctionName
    sUIFail = sFunctionName + ">> Fn_UI_JavaTree_NodeExist >> " +  objJavaTree.toString +">> " +  sNodeNameWithPath
	If objJavaTree.Exist Then
			intNodeCount = objJavaTree.GetROProperty ("items count")
			For intCount = 0 to intNodeCount - 1
					sTreeItem = objJavaTree.GetItem(intCount)
					If Trim(Lcase(sTreeItem)) = Trim(Lcase(sNodeNameWithPath)) Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sNodeNameWithPath &"Not Exist Function " &sFunctionName)
							Fn_UI_JavaTree_NodeExist = True
							Exit For
					Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sNodeNameWithPath &"Not Exist For Function " &sFunctionName)
							Fn_UI_JavaTree_NodeExist = False
					End If
			Next
	End If
End Function
'#################################################################################################################
'###    FUNCTION NAME   :    Fn_UI_SwfButtonClick(sFunctionName, objSwfDialog, sSwfButton)
'###
'###    DESCRIPTION     :   This function  is used to  click on SwfButton
'###
'###    PARAMETERS      :   	  sFunctionName : Valid Function name, 
'###                        				objSwfDialog : Valid object name,
'###                        				sSwfButton   : Valid Button name
'###                        
'###
'###    CREATED BY       :   Koustubh Watwe 	17 June 2011
'###
'###	REViEWED BY		: 
'###
'###    MODIFIED BY     :  Nilesh                      1 June 2012
'###
'### 	 EXAMPLE         	   :    MsgBox Fn_UI_SwfButtonClick("Test Function", SwfWindow("Select Sheet"), "OK")
'################################################################################################################
Public Function Fn_UI_SwfButtonClick(sFunctionName, objSwfDialog, sSwfButton)
	Dim objSwfButton
	'Object Creation
	bGblFailedFunctionName = sFunctionName
	Fn_UI_SwfButtonClick = False
	sUIFail = sFunctionName & " >> Fn_UI_SwfButtonClick >> " &  objSwfDialog.toString & " >> " &  sSwfButton
	Set objSwfButton = objSwfDialog.SwfButton(sSwfButton)
	If objSwfDialog.exist(10) Then
			If objSwfButton.exist(10) Then
					objSwfButton.WaitProperty "enabled", True
					If objSwfButton.GetROProperty("enabled") = True  Then
							Call Fn_UI_WpfObjectClick("Fn_MSO_TeamcenterLogin",objSwfDialog.SwfButton(sSwfButton)) 'Added by Nilesh on 1st June 2012
'							objSwfButton.Click 1,1, micLeftBtn
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Sucessfully clicked on SwfButton " & sSwfButton & " of Function " & sFunctionName)
							Fn_UI_SwfButtonClick = True
					Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "SwfButton " & sSwfButton & " is desabled of Function " &sFunctionName )
							Call ExitFromUI(sUIFail)
					End If
			Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "SwfButton " & sSwfButton & " does not exist of Function " &sFunctionName )
					Call ExitFromUI(sUIFail)
			End If
	Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Dialog "& objSwfDialog.toString & " does not exist of Function " &sFunctionName )
			Call ExitFromUI(sUIFail)
	End If
End Function
'#################################################################################################################
'###    FUNCTION NAME   :    Fn_UI_WpfButtonClick(sFunctionName, objSwfDialog, sSwfButton)
'###
'###    DESCRIPTION     :   This function  is used to  click on SwfButton
'###
'###    PARAMETERS      :   	  sFunctionName : Valid Function name, 
'###                        	  objWpfDialog : Valid object name,
'###                        	  sWpfButton   : Valid Button name
'###                        
'###
'###    CREATED BY       :   Koustubh Watwe 	17 June 2011
'###
'###	REViEWED BY		: 
'###
'###    MODIFIED BY     :  Nilesh           1st June 2012
'###
'### 	 EXAMPLE         	   :    MsgBox Fn_UI_WpfButtonClick("Test Function", WpfWindow("Teamcenter Login"), "Clear")
'################################################################################################################
Public Function Fn_UI_WpfButtonClick(sFunctionName, objSwfDialog, sWpfButton)
	Dim objWpfButton
    Dim x,y,dr,w,h
	bGblFailedFunctionName = sFunctionName
	'Object Creation
	Fn_UI_WpfButtonClick = False
	sUIFail = sFunctionName & " >> Fn_UI_WpfButtonClick >> " &  objSwfDialog.toString & " >> " &  sWpfButton
	Set objWpfButton = objSwfDialog.WpfButton(sWpfButton)
	If Fn_UI_ObjectExist("Fn_UI_WpfButtonClick", objSwfDialog) Then
			If Fn_UI_ObjectExist("Fn_UI_WpfButtonClick", objWpfButton) Then
					objWpfButton.WaitProperty "enabled", True
					If objWpfButton.GetROProperty("enabled") = True  Then
'							objWpfButton.Click 1,1, micLeftBtn
'							Added by Nilesh on 1st June 2012
                            x=objWpfButton.GetRoProperty("abs_x")
							y=objWpfButton.GetRoProperty("abs_y")
							w=Cint(objWpfButton.GetRoProperty("width")/2)
							h=Cint(objWpfButton.GetRoProperty("height")/2)
							Set dr = CreateObject("Mercury.DeviceReplay")
							dr.MouseClick  Cint(x+w),Cint(y+h),LEFT_MOUSE_BUTTON
							Set dr=Nothing
							'End
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Sucessfully clicked on WpfButton " & sWpfButton & " of Function " & sFunctionName)
							Fn_UI_WpfButtonClick = True
					Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "WpfButton " & sWpfButton & " is desabled of Function " &sFunctionName )
							Call ExitFromUI(sUIFail)
					End If
			Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "WpfButton " & sWpfButton & " does not exist of Function " &sFunctionName )
					Call ExitFromUI(sUIFail)
			End If
	Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Dialog "& objSwfDialog.toString & " does not exist of Function " &sFunctionName )
			Call ExitFromUI(sUIFail)
	End If
End Function
'#################################################################################################################
'###    FUNCTION NAME   :    Fn_UI_TableOperations(sFunctionName,sAction,sJavaTable,iRow,iColumn)
'###	Note : Depricated Function :	Note: Use Fn_SISW_UI_JavaTable_Operations
'################################################################################################################
Function Fn_UI_TableOperations(sFunctionName,sAction,sJavaTable,iRow,iColumn)
	bGblFailedFunctionName = sFunctionName
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"),  "Depricated function : Fn_UI_TableOperations > Calling Function Fn_SISW_UI_JavaTable_Operations")
	Fn_UI_TableOperations = Fn_SISW_UI_JavaTable_Operations(sFunctionName, sAction, sJavaTable , "", "", "", iRow, iColumn, "", "", "")
End Function
'@@------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@
'@@    Function Name		:	Fn_UI_JavaTreeGetItemPath
'@@
'@@    Description				 :	Function Used to Return Index Of Tree Node
'@@
'@@    Parameters			   :	1.ObjTree: Tree Object
'@@											  2.StrNode: Full Node 
'@@
'@@    Return Value		   	   : 	Node Index Path / False
'@@
'@@    Pre-requisite			:	Tree Should Exist							
'@@
'@@    Examples					:	Call Fn_UI_JavaTreeGetItemPath(JavaWindow("MyTeamcenter").JavaTree("NavTree"),"Home:000021-sdasda:000021/A;1-sdasda")
'@@
'@@	   History					 	:	
'@@													Developer Name												Date						Rev. No.						Changes Done								Reviewer
'@@------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@												Yann G/ Sandeep N/Sunny R							11-11-11						1.0																						
'@@------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@												Koustubh Watwe										16-11-11						1.0										Modified code to handle multiple occurrences												
'@@------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@												Koustubh Watwe										19-12-11						1.0										Removed code to handle multiple occurrences
'@@------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_UI_JavaTreeGetItemPath(ObjTree,StrNode)
   'Variable Declaration
	Dim sItemPath,aStrNode,bFlag,i,iNodeItemsCount
	Dim oCurrentNode,eStrNode, iCount
	'bGblFailedFunctionName = "Fn_UI_JavaTreeGetItemPath"
	'Initial Item Path
	sItemPath="#0"
	aStrNode = Split (StrNode, ":")
	bFlag=False
	Set oCurrentNode = ObjTree.Object.getItem(0)
	'To handle the situation where operation needs to be performed on Root Node
	If UBound(aStrNode) = 0 Then
		Fn_UI_JavaTreeGetItemPath = sItemPath
		Exit Function
	End If
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
	If bFlag=True Then
		'Function Returns Item Path
		Fn_UI_JavaTreeGetItemPath = sItemPath
	Else
		Fn_UI_JavaTreeGetItemPath = False
	End If
	Set oCurrentNode =Nothing
End Function
'@@------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@
'@@    Function Name		:	Fn_UI_getTreeIndex
'@@
'@@    Description			:	Private Recursive Function Used to calculate Index Of Tree Node
'@@
'@@    Parameters			:	1. objCurrent: Tree Item Node object - generally root node of the tree
'@@								2. StrNode: Full Node 
'@@								
'@@								Global variables must be set
'@@								Fn_UI_getTreeIndex_CompareString = emty string must be set to ""
'@@								Fn_UI_getTreeIndex_iGblCnt = Global Item node counter
'@@								Fn_UI_getTreeIndex_bFound  = Falg which indicates when to terminate function execution
'@@								
'@@
'@@    Return Value		   	: 	None
'@@
'@@	   Note 				:  This function is used in Fn_UI_getJavaTreeIndex
'@@
'@@    Pre-requisite		:	Tree Should Exist							
'@@
'@@    Examples				:	Call Fn_UI_JavaTreeGetItemPath(objTree.Object.getItem(0),"Home:000021-sdasda:000021/A;1-sdasda")
'@@
'@@	   History					 	:	
'@@					Developer Name			Date				Rev. No.		Changes Done								Reviewer
'@@------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@					Koustubh Watwe			08-12-2011			1.0				Created
'@@------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@					Koustubh Watwe			08-02-2012			1.0
'@@------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@					Koustubh Watwe			04-Apr-2012			2.0
'@@------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private function Fn_UI_getTreeIndex(objCurrent, sNode)
	Dim iCnt, iItemCount
	'bGblFailedFunctionName = "Fn_UI_getTreeIndex"
	'Print objCurrent.getData().toString()
	If cBool(objCurrent.getExpanded()) = false AND cInt(objCurrent.getItemCount()) = 0 Then
		Fn_UI_getTreeIndex_iGblCnt = Fn_UI_getTreeIndex_iGblCnt + 1
		exit function
	Else
		If cBool(objCurrent.getExpanded())  then
			iItemCount =  cInt(objCurrent.getItemCount())
			for iCnt = 0 to iItemCount -1 
				Call Fn_UI_getTreeIndex(objCurrent.getItem(iCnt),"")
			next
		End IF
	End If
	Fn_UI_getTreeIndex_iGblCnt = Fn_UI_getTreeIndex_iGblCnt + 1	
End Function

'@@------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@
'@@    Function Name		:	Fn_UI_getJavaTreeIndex
'@@
'@@    Description			:	Function Used to retrieve Index Of Java Tree Node
'@@
'@@    Parameters			:	1. objTree: Tree object
'@@								2. StrNode: Full Node 
'@@								
'@@    Return Value		   	: 	Node index / -1
'@@
'@@    Pre-requisite		:	Tree Should Exist							
'@@
'@@    Examples				:	Call Fn_UI_getJavaTreeIndex(JavaWindow("MyTeamcenter").JavaTree("NavTree"), "Home:Newstuff:000136-new")
'@@								Call Fn_UI_getJavaTreeIndex(JavaWindow("MyTeamcenter").JavaTree("SearchResultTree"),  "Item... (1):000108-Top")
'@@
'@@	   History					 	:	
'@@					Developer Name			Date				Rev. No.		Changes Done								Reviewer
'@@------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@					Koustubh Watwe			08-12-2011			1.0				Created
'@@------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@					Koustubh Watwe			04-Apr-2012			2.0				Created
'@@------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public function Fn_UI_getJavaTreeIndex(objTree, sNode)
	Dim iPath, arriPath, iArrCnt, iNodeCnt, iLimit, obj
	Fn_UI_getTreeIndex_iGblCnt = -1
	Fn_UI_getTreeIndex_CompareString = ""
	Fn_UI_getTreeIndex_bFound = False
	'bGblFailedFunctionName = "Fn_UI_getJavaTreeIndex"
	iPath = Fn_UI_JavaTreeGetItemPathExt("FunctionName", objTree, sNode ,"","")
	If iPath = False Then
		Fn_UI_getJavaTreeIndex = Fn_UI_getTreeIndex_iGblCnt
		exit function
	End If
	Fn_UI_getTreeIndex_iGblCnt = 0
	iPath=Replace(iPath,"#","")
	arriPath=Split(iPath,":")
	If UBound(arriPath) <> 0 Then
		For iArrCnt = 0 to UBound(arriPath)-1
			arriPath(iArrCnt) = cInt(arriPath(iArrCnt))
			If iArrCnt = 0 Then
				Set obj = objTree.Object.getItem(arriPath(iArrCnt))
			Else
				Set obj = obj.getItem(arriPath(iArrCnt)) 
			End If
			iLimit = cInt(arriPath(iArrCnt + 1))
			If iArrCnt <> uBound(arriPath) Then
				iLimit = iLimit - 1 
			End If
			For iNodeCnt = 0 to iLimit
				Call Fn_UI_getTreeIndex(obj.getItem(iNodeCnt), "")
			Next
			Fn_UI_getTreeIndex_iGblCnt = Fn_UI_getTreeIndex_iGblCnt + 1
		Next
	Else
		Fn_UI_getTreeIndex_iGblCnt = cInt(arriPath(0))
	End If
	Fn_UI_getJavaTreeIndex = Fn_UI_getTreeIndex_iGblCnt
End Function
'@@------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@
'@@    Function Name		:	Fn_UI_JavaTreeGetItemPathExt
'@@
'@@    Description				 :	Function Used to Return Index Of Tree Node
'@@
'@@    Parameters			   :	1.ObjTree: Tree Object
'@@											  2.StrNode: Full Node 
'@@
'@@    Return Value		   	   : 	Node Index Path / False
'@@
'@@    Pre-requisite			:	Tree Should Exist							
'@@
'@@    Examples					:	Call Fn_UI_JavaTreeGetItemPathExt("FunctionName", JavaWindow("MyTeamcenter").JavaTree("NavTree"),"Home:000021-sdasda:000021/A;1-sdasda",":","@")
'@@
'@@	   History					:	
'@@					Developer Name						Date				Rev. No.			Changes Done								Reviewer
'@@------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@					Koustubh Watwe						28-Feb-2012			  1.0				Modified code to handle multiple occurrences												
'@@------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_UI_JavaTreeGetItemPathExt(sFunctionName, ObjTree, StrNode, sDelimiter, sInstanceHandler)
   'Variable Declaration
	Dim sItemPath,aStrNode,bFlag,i,iNodeItemsCount
	Dim oCurrentNode,eStrNode, iCount, iNodecnt
	Dim iInstanceCnt, aNode,iOccCnt
	Dim sTreeNodeStr
	bGblFailedFunctionName = sFunctionName
	If sDelimiter = "" Then sDelimiter = ":"
	If sInstanceHandler = "" Then sInstanceHandler = "@"
	Fn_UI_JavaTreeGetItemPathExt = False
	Set objNodeBounds = nothing

	'Initial Item Path
	sItemPath= False
	aStrNode = Split (StrNode, sDelimiter)
	bFlag=False
	
	'To handle the situation where operation needs to be performed on Root Node
	iOccCnt = 1
	For iCount = 0 to ObjTree.Object.getItemCount() - 1
		If Instr(aStrNode(0), sInstanceHandler) > 0 Then
			aNode = split(aStrNode(0),sInstanceHandler)
			eStrNode = trim(aNode(0))
			iInstanceCnt = cInt(aNode(1) )
		Else
			eStrNode = trim(aStrNode(0))
			iInstanceCnt = 1
		End If
		If ObjTree.Object.getItem(iCount).getText() = eStrNode Then
			If  iOccCnt = iInstanceCnt Then
				Set oCurrentNode = ObjTree.Object.getItem(iCount)
				sItemPath = "#" & iCount
				bFlag = True
				Exit For
			else
				iOccCnt = iOccCnt + 1
			End If
		Else
			sTreeNodeStr = Fn_SISW_UI_JavaTree_GetSanitizedNodeName(ObjTree.Object.getItem(iCount))
			If sTreeNodeStr = eStrNode Then
				If  iOccCnt = iInstanceCnt Then
					Set oCurrentNode = ObjTree.Object.getItem(iCount)
					sItemPath = "#" & iCount
					bFlag = True
					Exit For
				else
					iOccCnt = iOccCnt + 1
				End If
			End If
		End If
	Next
	If UBound(aStrNode) = 0 Then
		Fn_UI_JavaTreeGetItemPathExt = sItemPath
		If sItemPath=False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Function " & sFunctionName & " Failed to find item [ " & StrNode & " ]"  )
		Else
			Set objNodeBounds = oCurrentNode.getBounds()
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Function " & sFunctionName & " executed successfully for item [ " & StrNode & " ]"  )
		End If
		Exit Function
	End If
	If bFlag Then
		bFlag = False
	Else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Function " & sFunctionName & " Failed to find item [ " & StrNode & " ]"  )
		Exit function
	End If
		'To Select first Occurance of Node
'		For each eStrNode In 
		For iNodecnt = 1 to UBound(aStrNode)
			eStrNode = aStrNode(iNodecnt)
			iNodeItemsCount = oCurrentNode.getItemCount()
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
				If Trim(oCurrentNode.getItem(i).getText()) = Trim(eStrNode) Then
					If  iOccCnt = iInstanceCnt Then
						Set oCurrentNode = oCurrentNode.getItem(i)
						sItemPath = sItemPath & ":#" & i
						bFlag=True
						Exit For
					else
						iOccCnt = iOccCnt + 1
					End If
				Else
					sTreeNodeStr = Fn_SISW_UI_JavaTree_GetSanitizedNodeName(oCurrentNode.getItem(i))
					If sTreeNodeStr = Trim(eStrNode) Then
						If  iOccCnt = iInstanceCnt Then
							Set oCurrentNode = oCurrentNode.getItem(i)
							sItemPath = sItemPath & ":#" & i
							bFlag=True
							Exit For
						else
							iOccCnt = iOccCnt + 1
						End If
					End If
				End If
			Next
			If bFlag=False Then
				Exit For
			End If
		Next 
	If bFlag=True Then
		'Function Returns Item Path
		Fn_UI_JavaTreeGetItemPathExt = sItemPath
		Set objNodeBounds = oCurrentNode.getBounds()
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Function " & sFunctionName & " executed successfully for item [ " & StrNode & " ]"  )
	Else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Function " & sFunctionName & " Failed to find item [ " & StrNode & " ]"  )
		Fn_UI_JavaTreeGetItemPathExt = False
	End If
	Set oCurrentNode =Nothing
End Function
'@@------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@
'@@    Function Name		:	Fn_UI_ClickJavaTreeCell
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
'@@    Examples				:	Call Fn_UI_ClickJavaTreeCell("Fn_CPD_RecipeOperations",JavaWindow("Collaborative Product"), "ComponentTree","DE000176/001;1-d2","Logic","LEFT")
'@@    Examples				:	Call Fn_UI_ClickJavaTreeCell("Fn_CPD_RecipeOperations",JavaWindow("Collaborative Product"), "ComponentTree","DE000176/001;1-d2","Logic","RIGHT")
'@@
'@@	   History				:	
'@@					Developer Name					Date				Rev. No.			Changes Done								Reviewer
'@@----------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@					Valalri, Koustubh				11-Apr-2012			  1.0				Created											
'@@----------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_UI_ClickJavaTreeCell(sCallerFunctionName, objDialog, sTree, sNode, sColumn, sButton)
	Dim intX, intY, iCnt, iIterate, objTree
	Dim sColName, sItemName, sUIFail
	Dim iColWidth, iItmHeight
	bGblFailedFunctionName = sCallerFunctionName
	sUIFail = sCallerFunctionName + ">> Fn_Menu_Select >> " +  objDialog.toString + " >>  Tree[ " & sTree & " ] "
	Fn_UI_ClickJavaTreeCell = False
	intY = 0
	intX = 0
	Set objTree = objDialog.JavaTree(sTree)
	If Fn_UI_ObjectExist("Fn_UI_ClickJavaTreeCell", objTree) = False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_UI_ClickJavaTreeCell : FAIL : " & sUIFail & " Doesn not exist.")	
		Exit function
	End If
	If isNumeric(trim(sColumn)) = Flase Then
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
		iCnt = Fn_UI_getJavaTreeIndex(objTree, sNode)
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
			objTree.Click intX, intY,"LEFT"
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
	Fn_UI_ClickJavaTreeCell = True
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_UI_ClickJavaTreeCell : PASS : Executed successfully with [ " & sButton & " Click].")	
End Function

''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''''/$$$$
''''/$$$$   FUNCTION NAME   :  Fn_UI_WpfObjectClick(sFunctionName, objSwfObject)
''''/$$$$
''''/$$$$   DESCRIPTION        : To click on SWFobject
''''/$$$$ 
''''/$$$$
''''/$$$$   PARAMETERS      :   sFunctionNamee - Valid function name
''''/$$$$                    objSwfObject - Object hierachy
''''/$$$$	
''''/$$$$		Return Value : 				Nothing
''''/$$$$
''''/$$$$    Function Calls       :   NA
''''/$$$$										
''''/$$$$
''''/$$$$		HISTORY           :   		AUTHOR                 DATE        VERSION
''''/$$$$
''''/$$$$    CREATED BY     :   Nilesh     			    01/05/2012         1.0
''''/$$$$
''''/$$$$    REVIWED BY     :  
''''/$$$$
''''/$$$$		How To Use :    bReturn=Fn_UI_WpfObjectClick("Test" ,WpfWindow("Save"))
''''/$$$$							
''''/$$$$	
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

Function Fn_UI_WpfObjectClick(sFunctionName, objSwfObject)
	Dim objDeviceRply,objSwf,x,y,h,w
	bGblFailedFunctionName = sFunctionName
	Set objSwf=objSwfObject
	Set objDeviceRply=CreateObject("Mercury.DeviceReplay")

	x=objSwf.GetRoProperty("abs_x")
	y=objSwf.GetRoProperty("abs_y")
	w=Cint(objSwf.GetRoProperty("width")/2)
	h=Cint(objSwf.GetRoProperty("height")/2)
	objDeviceRply.MouseClick Cint(x+w),Cint(y+h),LEFT_MOUSE_BUTTON
	Set objDeviceRply=Nothing
	Set objSwf=Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_UI_JavaStatictextOperations

'Description			 :	Function Used to perform operation on JavaStaticText

'Parameters			   :   1.StrFunctionName: Function Name
'										2.StrAction: Action Name
'										3.objDialog: Dialog path
'										4.StrDialogTitle: Dialog Title optional
'										5.StrStaticText: Static texts
'
'Return Value		   : 	True or False

'Pre-requisite			:	

'Examples				:  	bReturn=Fn_UI_JavaStatictextOperations("","Verify",JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("NewLink"),"New Link","InterpartEquation~InterpartLink~MatingConstraint~ProductionRelation~TC_Link")
'										bReturn=Fn_UI_JavaStatictextOperations("","Click",JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Interface Definition"),"",aStaticText)
'
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												12-Apr-2012								1.0																					Priyanka B
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Public Function Fn_UI_JavaStatictextOperations(StrFunctionName,StrAction,objDialog,StrDialogTitle,StrStaticText)
   Dim objJavaDialog,arrStaticText,bFlag,objStaticText,objChild,iCounter,iCount
   	bGblFailedFunctionName = StrFunctionName
	Set objJavaDialog=objDialog
   If StrDialogTitle<>"" Then
		objJavaDialog.SetTOProperty "title",StrDialogTitle
   End If
   If objJavaDialog.Exist(6) Then
		Select Case StrAction
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			Case "Verify"
					arrStaticText=Split(StrStaticText,"~")
					For iCounter=0 To UBound(arrStaticText)
						bFlag=False
						Set objStaticText=Description.Create()
						objStaticText("Class Name").value="JavaStaticText"
						objStaticText("label").value=arrStaticText(iCounter)
						Set objChild=objJavaDialog.ChildObjects(objStaticText)
						For iCount=0 to objChild.count-1
								If objChild(iCount).GetROProperty("label")=arrStaticText(iCounter) Then
									bFlag=True
								End If
						Next
						Set objStaticText=Nothing
						Set objChild=nothing
						If bFlag=False Then
							Fn_UI_JavaStatictextOperations=False
							Exit For
						End If
					Next
					If bFlag=True Then
						Fn_UI_JavaStatictextOperations=True
					End If
				' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				Case "Click"
                	Set objStaticText=Description.Create()
					objStaticText("Class Name").value="JavaStaticText"
					objStaticText("label").value=StrStaticText
					Set objChild=objJavaDialog.ChildObjects(objStaticText)
					If objChild(0).Exist(5) Then
						objChild(0).Click 1,1
						bFlag=True
					End If
					Set objStaticText=Nothing
					Set objChild=nothing
                    If bFlag=True Then
						Fn_UI_JavaStatictextOperations=True
					End If
				Case Else
					Fn_UI_JavaStatictextOperations=False
		End Select
	Else
		Fn_UI_JavaStatictextOperations=False
   End If
   Set objJavaDialog=nothing
End Function

'#########################################################################################################
'###
'###    FUNCTION NAME   :   Function Fn_Spin_Edit(sFunctionName,objJavaDialog,sJavaEdit,sText)
'###
'###    DESCRIPTION     :   This function is to insert or set the text
'###
'###    PARAMETERS      :   1.sFunctionName: Valid Function Name
'###			    		2.objJavaDialog: Valid Java Dialog	
'###			    		3.sJavaSpin:Valid Spin box name
'###			   			4.sText: Valid text to be set in the edit box
'###                                         
'###    Function Calls  :  Fn_WriteLogFile()
'### 
'###	 HISTORY         :   AUTHOR                 DATE             VERSION
'###
'###    CREATED BY      :   Shrikant N            29-06-12          1.0
'###
'###    EXAMPLE         : Fn_SISW_UI_Spin_Edit("Fn_RDV_SpatialCriteria", JavaWindow("RDV_StructureManager").JavaWindow("Spatial Filter"),"Xmin","-544")
'#############################################################################################################

 
Function Fn_SISW_UI_Spin_Edit(sFunctionName,objJavaDialog,sJavaSpin,sText)

Dim objJavaSpin
	bGblFailedFunctionName = sFunctionName
  sUIFail = sFunctionName + ">> Fn_SISW_UI_Spin_Edit >> " +  objJavaDialog.toString +">> " +  sJavaSpin
  
	'Set an Edit Object on variable
   Set objJavaSpin= objJavaDialog.JavaSpin(sJavaSpin)

		
		'Checking the editbox is exists or not
 		 If objJavaSpin.Exist Then

			'Syncronization point
				objJavaSpin.WaitProperty "enabled", "1"
		
				
				'Checking It is Enabled or not	
			  If  objJavaSpin.GetROProperty("enabled") = "1"  Then
			
				'Setting the JavaSpin
				objJavaSpin.Type sText

				'Log the success
            			 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Text " &sText  &"is Set/ Entered in JavaSpinBox" & sJavaSpin & " of Function " &sFunctionName)
			
				'Return true from Function
				 Fn_SISW_UI_Spin_Edit= True
			 Else
			
				'Log the failure when it not enabled	
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"JavaSpinEdit " & sJavaSpin &"is Disable of Function " &sFunctionName)
			
				'Return False from function
				Fn_SISW_UI_Spin_Edit= False
				Call ExitFromUI(sUIFail)
			End If
		 Else

			'Log the failure when it does not exists
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sJavaSpin & "JavaSpinEdit " &sJavaSpin &" does Not Exists" &" of Function  " &sFunctionName)

			'Return False from function
			Fn_SISW_UI_Spin_Edit= False
			Call ExitFromUI(sUIFail)
	End If

	Set objJavaSpin = Nothing 
End Function
'#########################################################################################################
'###
'###	FUNCTION NAME	:	Function Fn_SISW_UI_JavaTree_GetSanitizedNodeName
'###
'###	DESCRIPTION		:	This function is to remove unnecessary text from Java Tree Node name
'###
'###	PARAMETERS		:	1.objTreeNode: Tree Node Object
'###
'###	HISTORY			:	AUTHOR					DATE			VERSION
'###
'###	Return Value	:	Sanitized Node Text
'###
'###	CREATED BY		:	Koustubh W				16-07-12		1.0
'###
'###	EXAMPLE			:	Fn_SISW_UI_JavaTree_GetSanitizedNodeName(objTree.Object.getItem(0))
'###	
'###	MODIFIED BY		:	Reema W					20-11-15		1.1			[TC1121-2015102600-20_11_2015-VivekA-NewDevelopment]
'###						Added Case "class com.teamcenter.rac.partition.nodes.PartitionableBaselineNode"
'###						to select Baseline nodes & nodes under baseline in CPD perspective as Baseline node contains strings "( Open )" or "( Closed )"
'#############################################################################################################
Public Function Fn_SISW_UI_JavaTree_GetSanitizedNodeName(objTreeNode)
	Fn_SISW_UI_JavaTree_GetSanitizedNodeName = objTreeNode.getData().toString()
	'bGblFailedFunctionName = "Fn_SISW_UI_JavaTree_GetSanitizedNodeName"
	Select Case objTreeNode.getData().getClass().toString()
		Case "class com.teamcenter.rac.cme.framework.application.model.impl.OccurrenceGroupNode"
			Fn_SISW_UI_JavaTree_GetSanitizedNodeName = replace(Fn_SISW_UI_JavaTree_GetSanitizedNodeName," (OccurrenceGroupNode)","")

		Case "class com.teamcenter.rac.cme.framework.application.model.impl.StructureContextNode"
			Fn_SISW_UI_JavaTree_GetSanitizedNodeName = replace(Fn_SISW_UI_JavaTree_GetSanitizedNodeName," (StructureContextNode)","")

		Case "class com.teamcenter.rac.cme.framework.application.model.impl.EndItemNode"
			Fn_SISW_UI_JavaTree_GetSanitizedNodeName = replace(Fn_SISW_UI_JavaTree_GetSanitizedNodeName," (EndItemNode)","")

		Case "class com.teamcenter.rac.cme.framework.application.model.impl.ConfigurationContextNode"
			Fn_SISW_UI_JavaTree_GetSanitizedNodeName = replace(Fn_SISW_UI_JavaTree_GetSanitizedNodeName," (ConfigurationContextNode)","")

		Case "class com.teamcenter.rac.cme.framework.application.model.impl.CCObjectRootNode"
			Fn_SISW_UI_JavaTree_GetSanitizedNodeName = replace(Fn_SISW_UI_JavaTree_GetSanitizedNodeName," (CCObjectRootNode)","")

		Case "class com.teamcenter.rac.cm.ui.changehome.ChangeHomePseudoFolder",_
			 "class com.teamcenter.rac.cm.ui.changehome.ChangeHomeQueryPseudoFolder",_
			 "class com.teamcenter.rac.mechatronics.ccdm.common.model.ParmDicOrPrjModel",_
			 "class com.teamcenter.rac.mechatronics.ccdm.common.model.ParmMemLayoutsModel",_
			 "class com.teamcenter.rac.mechatronics.ccdm.common.model.OverrideContPseudoModel"
			Fn_SISW_UI_JavaTree_GetSanitizedNodeName = objTreeNode.getData().getDisplayName()
		Case "class com.teamcenter.rac.partition.nodes.PartitionableBaselineNode"
			Fn_SISW_UI_JavaTree_GetSanitizedNodeName = replace(Fn_SISW_UI_JavaTree_GetSanitizedNodeName,"( Closed )","")
			Fn_SISW_UI_JavaTree_GetSanitizedNodeName = replace(Fn_SISW_UI_JavaTree_GetSanitizedNodeName,"( Open )","")
		Case Else
			'Do Nothing
	End Select
	Fn_SISW_UI_JavaTree_GetSanitizedNodeName = Trim(Fn_SISW_UI_JavaTree_GetSanitizedNodeName)
End Function 
'#########################################################################################################
'###
'###	FUNCTION NAME	:	Function Fn_SISW_UI_RACTabFolderWidget_Operation
'###
'###	DESCRIPTION		:	This function is to perform operations on all  RACTabFolderWidget objects in Default Winodw
'###
'###	PARAMETERS		:	1. sAction: Tree Node Object
'###						2. sItem = Tab to be selected
'###						3. sMenu = Menu 
'###	HISTORY			:	AUTHOR					DATE			VERSION
'###
'###	Return Value	:	Sanitized Node Text
'###
'###	CREATED BY		:	Koustubh W				07-09-12		1.0
'###
'###	EXAMPLE			:	Fn_SISW_UI_RACTabFolderWidget_Operation("Select", "Structure Search", "")
'###
'###	Modified BY		:	Vivek A					09-09-16		1.1				[TC1123-20160810-09_09_2016-VivekA-NewDevelopment] - 4GD New Tc's
'###								Added new Case "IsMaximized" to check whether tab is maximized or not
'#############################################################################################################
Public Function Fn_SISW_UI_RACTabFolderWidget_Operation(sAction, sItem, sMenu)
'	Dim strMenu, sxLen, syLen
'	Dim i, iCnt, aBounds, sBounds
'	Dim objTabFld, objItem,objSelectType
'	Dim objIntNoOfObjects,iCount
'	Dim arrItem,iCounter,bFlag, iIndex
'	Dim X, Y, W, H 
'	'bGblFailedFunctionName = "Fn_SISW_UI_RACTabFolderWidget_Operation"
'	bFlag = False
'	Fn_SISW_UI_RACTabFolderWidget_Operation = False
'	
'	arrItem = Split(sItem,"@")
'	If arrItem(0) = "Requirements" Then
'		arrItem(0) = "Requirement"
'	End If
'
'	' finding object of tab in all RACTabFolderWidget objects
'	Set objTabFld = JavaWindow("DefaultWindow").JavaObject("RACTabFolderWidget")
'	objTabFld.SetTOProperty "Index", 0
'	iCnt = objTabFld.Object.getTabItemCount
'	For iIndexCnt = 0 to 5
'		objTabFld.SetTOProperty "Index", iIndexCnt
'		sxLen = 0
'		syLen = 0
'		If uBound(arrItem) > 0 Then
'			iIndex = cInt(arrItem(1)) - 1
'		Else
'			iIndex = 0
'		End If
'		If objTabFld.Exist(2) Then
'			iCnt = objTabFld.Object.getTabItemCount
'			For i = 0 to iCnt-1
'				set objItem = objTabFld.Object.getItem(i)
'				'print objItem.text
'				sxLen = sxLen + objItem.getWidth
'				If trim(objItem.text) = trim(arrItem(0)) Then
'					If iIndex = 0 Then
'						bFlag = True
'						Exit for
'					Else
'						iIndex = iIndex - 1
'					End If
'				End If
'			Next
'			If bFlag = True Then
'				Exit for
'			End If
'		End If
'	Next
'	If bFlag = False Then
'		objTabFld.SetTOProperty "Index", 1
'		Set objItem = Nothing
'		Exit Function
'	End If
'	If sAction="Select" Then
'		bFlag=objItem.isShowing()
'		If Cbool(bFlag)=false Then
'			objTabFld.DblClick 1,1,"LEFT"
'			wait 2
'		End If
'		bFlag1=objItem.isShowing()
'		If CBool(bFlag1)=false Then
'			objTabFld.Object.setSelection(i)
'			Fn_SISW_UI_RACTabFolderWidget_Operation = True
'			wait 2
'		End If
'	End If
''	sxLen = sxLen - (objItem.getWidth/2)
'	sBounds = objItem.getBounds().toString()
'	sBounds = mid(sBounds,instr(sBounds,"{")+1, len(sBounds) -instr(sBounds,"{")-1)
'	aBounds = split(sBounds,",")
'	X = cInt(trim(aBounds(0)))
''	Y = cInt(trim(aBounds(1)))
''	W = cInt(trim(aBounds(2)))
''	H = cInt(trim(aBounds(3)))
'
'	sxLen = X + 15
'	syLen = (objItem.getHeight/2)
'	
'	Select Case sAction
'		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'		Case "Select"	
'			
'			If cBool(bFlag1) =True Then
'				objTabFld.Click sxLen, syLen, "LEFT"
'				Fn_SISW_UI_RACTabFolderWidget_Operation = True
'			End If
'			
'			If cBool(bFlag)=false Then
'                wait 2
'				objTabFld.DblClick sxLen, syLen, "LEFT"
''				objTabFld.Click sxLen, syLen, "LEFT"
'				wait 2
'			End if
'		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'		Case "DoubleClick" 
'			objTabFld.DblClick sxLen, syLen, "LEFT"
'			Fn_SISW_UI_RACTabFolderWidget_Operation = True
'		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'		Case "RMBMenuSelect"
'			objTabFld.Click sxLen, syLen, "RIGHT"
'			wait 2
'			aMenuList = split(sMenu, ":",-1,1)
'			iTabCount = Ubound(aMenuList)
'			'Select Menu action
'			Select Case iTabCount
'				Case "0"
'					 StrMenu = JavaWindow("DefaultWindow").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0))
'				Case "1"
'					StrMenu = JavaWindow("DefaultWindow").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1))
'				Case "2"
'					StrMenu = JavaWindow("DefaultWindow").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1),aMenuList(2))
'				Case Else
'					Fn_SISW_UI_RACTabFolderWidget_Operation = FALSE
'					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_SISW_UI_RACTabFolderWidget_Operation : Invalid PopupMenu  [" + sMenu + "] is Requested.")
'					objTabFld.SetTOProperty "Index", 1
'					Exit Function
'			End Select
'			On error resume next
'			JavaWindow("DefaultWindow").WinMenu("ContextMenu").Select StrMenu
'			If Err.Number < 0 Then
'				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to select [" & sMenu & "] PopupMenu.")
'				Fn_SISW_UI_RACTabFolderWidget_Operation = False
'			Else
'				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully selected [" & sMenu & "] PopupMenu.")
'				Fn_SISW_UI_RACTabFolderWidget_Operation = TRUE
'			End If    
'		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'		Case "VerifyActivate"
'            i = objTabFld.Object.getSelectedTabIndex
'			set objItem = objTabFld.Object.getItem(i)
'			If trim(objItem.text) = trim(sItem) Then
'					Fn_SISW_UI_RACTabFolderWidget_Operation = True
'					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_SISW_UI_RACTabFolderWidget_Operation :Successfully Verified that Tab ["+sItem+"] is Activated.")
'			Else
'					Fn_SISW_UI_RACTabFolderWidget_Operation = False
'					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_SISW_UI_RACTabFolderWidget_Operation : Failed to Verify that Tab ["+sItem+"] is Activated.")
'			End If
'		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'		Case "Verify"
'			Fn_SISW_UI_RACTabFolderWidget_Operation = True
'			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_SISW_UI_RACTabFolderWidget_Operation : Successfully Verified that Tab ["+sItem+"] is Activated.")
'		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'		Case "Close"
'			objTabFld.Click sxLen, syLen, "LEFT"
'			Call Fn_ReadyStatusSync(2)
'			sBounds = objItem.getCloseButtonBounds.toString()
'			sBounds = right(sBounds, Len(sBounds)-instr(sBounds, "{"))
'			aBounds = split(sBounds, ",", -1, 1)
'			sxLen = Cint(trim(aBounds(0))) + 5
'			syLen = Cint(trim(aBounds(1))) + 5
'			objTabFld.Click sxLen, syLen, "LEFT"
'			Fn_SISW_UI_RACTabFolderWidget_Operation = True
'		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'		Case "IsMaximized"
'			iWidth = CInt(objTabFld.GetROProperty("width"))
'			If iWidth>900 Then
'				Fn_SISW_UI_RACTabFolderWidget_Operation = True
'			End If
'		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'		Case "GetTabToolTipText"
'			Fn_SISW_UI_RACTabFolderWidget_Operation = objItem.getToolTipText()
'		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'		Case Else
'			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_SISW_UI_RACTabFolderWidget_Operation : Invalid case called.")
'		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'	End Select
'objTabFld.SetTOProperty "Index", 1
'Set objItem = Nothing
'Set objTabFld = Nothing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''	
	' [TC12 - 201706300 -18_7_2017-JotibaT-HC_Maintenance] Modified code by Jotiba T. (Object changed from JavaObject to JavaTab)
	Dim objSelectType,objIntNoOfObjects,objItem
	Dim icount,iItemCount,iCounter,iIndex,iWidth,iTabCount
	Dim bFlag, bFlag1
	Dim sBounds,aBounds,X,H,sxLen,syLen
	Dim aMenuList, StrMenu,arrItem
	sxLen=0
	syLen=0
	bFlag=False
	
'	arrItem = Split(sItem,"@")
'	If arrItem(0) = "Requirements" Then
'		sItem = "Requirement"
'	End If

	
	Set objSelectType = description.Create()
	objSelectType("Class Name").value = "JavaTab"
	objSelectType("toolkit class").value = "org.eclipse.swt.custom.CTabFolder"						
	Set  objIntNoOfObjects = JavaWindow("DefaultWindow").ChildObjects(objSelectType)
	
	Select Case sAction
		'--------------------------------------------------------------------------------------------------------------------------------
			Case "Select"	
				For icount = 0 To objIntNoOfObjects.Count-1 Step 1			
					iItemCount = cInt(objIntNoOfObjects(icount).Object.getItemCount())
					For iCounter = 0 To iItemCount- 1 Step 1
						If trim(sItem) = trim(objIntNoOfObjects(icount).Object.getItems().mic_arr_get(iCounter).getText()) Then
							objIntNoOfObjects(icount).Select sItem
							bFlag=True
							Exit For 
						End IF
					Next
					If bFlag=True Then Exit For 
				Next
		'--------------------------------------------------------------------------------------------------------------------------------		
			Case "DoubleClick"
					For icount = 0 To objIntNoOfObjects.Count-1 Step 1			
						iItemCount = cInt(objIntNoOfObjects(icount).Object.getItemCount())
						For iCounter = 0 To iItemCount- 1 Step 1
							If trim(sItem) = trim(objIntNoOfObjects(icount).Object.getItems().mic_arr_get(iCounter).getText()) Then
								objIntNoOfObjects(icount).Select sItem
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
		'--------------------------------------------------------------------------------------------------------------------------------		
			Case "VerifyActivate"
				For icount = 0 To objIntNoOfObjects.Count-1 Step 1			
					iItemCount = cInt(objIntNoOfObjects(icount).Object.getItemCount())
					For iCounter = 0 To iItemCount- 1 Step 1
						If trim(sItem) = trim(objIntNoOfObjects(icount).Object.getItems().mic_arr_get(iCounter).getText()) Then
							iIndex=objIntNoOfObjects(icount).Object.getSelectionIndex
							Set objItem=objIntNoOfObjects(icount).Object.getItem(iIndex)
							  If trim(objItem.text)= trim(sItem)Then
							  	bFlag=True
							  	Exit For 
							  End If
						End IF
					Next
					If bFlag=True Then Exit For 
				Next
		'--------------------------------------------------------------------------------------------------------------------------------			
			Case "Verify", "Exist"
				bFlag1 = 0
				If instr(1,sItem,"@")>0 Then
					aItem = split(sItem,"@") 
					For icount = 0 To objIntNoOfObjects.Count-1 Step 1			
						iItemCount = cInt(objIntNoOfObjects(icount).Object.getItemCount())
						For iCounter = 0 To iItemCount- 1 Step 1
							If trim(aItem(0)) = trim(objIntNoOfObjects(icount).Object.getItems().mic_arr_get(iCounter).getText()) Then
							  	bFlag1 = bFlag1+1 
							End IF
						Next
						If bFlag1 > 1 Then 
							bFlag=True
							Exit For
						End If
						'Exit For 
					Next
				Else
					For icount = 0 To objIntNoOfObjects.Count-1 Step 1			
						iItemCount = cInt(objIntNoOfObjects(icount).Object.getItemCount())
						For iCounter = 0 To iItemCount- 1 Step 1
							If trim(sItem) = trim(objIntNoOfObjects(icount).Object.getItems().mic_arr_get(iCounter).getText()) Then
							  	bFlag=True
							  	Exit For 
							End IF
						Next
						If bFlag=True Then Exit For 
					Next
				End If
		'--------------------------------------------------------------------------------------------------------------------------------			
			Case "Close"
				For icount = 0 To objIntNoOfObjects.Count-1 Step 1			
					iItemCount = cInt(objIntNoOfObjects(icount).Object.getItemCount())
					For iCounter = 0 To iItemCount- 1 Step 1
						If trim(sItem) = trim(objIntNoOfObjects(icount).Object.getItems().mic_arr_get(iCounter).getText()) Then
							objIntNoOfObjects(icount).Select sItem
							wait 1
							objIntNoOfObjects(icount).CloseTab sItem
							bFlag=True
							Exit For 
						End IF
					Next
					If bFlag=True Then Exit For 
				Next
		'--------------------------------------------------------------------------------------------------------------------------------			
			Case "GetTabToolTipText"
				For icount = 0 To objIntNoOfObjects.Count-1 Step 1			
					iItemCount = cInt(objIntNoOfObjects(icount).Object.getItemCount())
					For iCounter = 0 To iItemCount- 1 Step 1
						If trim(sItem) = trim(objIntNoOfObjects(icount).Object.getItems().mic_arr_get(iCounter).getText()) Then
							Fn_SISW_UI_RACTabFolderWidget_Operation=objIntNoOfObjects(icount).Object.getItems().mic_arr_get(iCounter).getToolTipText()
							bFlag=True
							Exit Function 
						End IF
					Next
					If bFlag=True Then Exit For 
				Next
		'--------------------------------------------------------------------------------------------------------------------------------			
			Case "IsMaximized"
				bFlag1=False
				For icount = 0 To objIntNoOfObjects.Count-1 Step 1			
					iItemCount = cInt(objIntNoOfObjects(icount).Object.getItemCount())
					For iCounter = 0 To iItemCount- 1 Step 1
						If trim(sItem) = trim(objIntNoOfObjects(icount).Object.getItems().mic_arr_get(iCounter).getText()) Then
							iWidth=CInt(objIntNoOfObjects(icount).GetRoProperty("width"))
							If iWidth>900 Then
								bFlag=True
								Exit For 
							Else
								bFlag1=True
								Exit For 
							End If
						End IF
					Next
					If bFlag=True Then Exit For 
					If bFlag1=True Then Exit For
				Next
		'--------------------------------------------------------------------------------------------------------------------------------
			Case "RMBMenuSelect"
				For icount = 0 To objIntNoOfObjects.Count-1 Step 1			
						iItemCount = cInt(objIntNoOfObjects(icount).Object.getItemCount())
						For iCounter = 0 To iItemCount- 1 Step 1
							If trim(sItem) = trim(objIntNoOfObjects(icount).Object.getItems().mic_arr_get(iCounter).getText()) Then
								objIntNoOfObjects(icount).Select sItem
								iIndex=objIntNoOfObjects(icount).Object.getSelectionIndex
								Set objItem=objIntNoOfObjects(icount).Object.getItem(iIndex)
									sBounds = objItem.getBounds().toString()
									sBounds = mid(sBounds,instr(sBounds,"{")+1, len(sBounds) -instr(sBounds,"{")-1)
									aBounds = split(sBounds,",")
									X = cInt(trim(aBounds(0)))
									H = cInt(trim(aBounds(3)))
									sxLen = X + 15
									syLen = (H/2)
								objIntNoOfObjects(icount).Click sxLen,syLen,"RIGHT"
								wait 2
								bFlag=True
								Exit For 
						End IF
					Next
					If bFlag=True Then Exit For 
				Next
				
				aMenuList = split(sMenu, ":",-1,1)
				iTabCount = Ubound(aMenuList)
				'Select Menu action
				Select Case iTabCount
					Case "0"
						 StrMenu = JavaWindow("DefaultWindow").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0))
					Case "1"
						StrMenu = JavaWindow("DefaultWindow").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1))
					Case "2"
						StrMenu = JavaWindow("DefaultWindow").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1),aMenuList(2))
					Case Else
						bFlag = FALSE
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_SISW_UI_RACTabFolderWidget_Operation : Invalid PopupMenu  [" + sMenu + "] is Requested.")
						Exit Function
				End Select
				On error resume next
				JavaWindow("DefaultWindow").WinMenu("ContextMenu").Select StrMenu
				If Err.Number < 0 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to select [" & sMenu & "] PopupMenu.")
					bFlag=False
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully selected [" & sMenu & "] PopupMenu.")
					bFlag = TRUE
				End If   
		'--------------------------------------------------------------------------------------------------------------------------------
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_SISW_UI_RACTabFolderWidget_Operation : Invalid case called.")				
		'--------------------------------------------------------------------------------------------------------------------------------
	End Select
	
	If bFlag=False Then
		Fn_SISW_UI_RACTabFolderWidget_Operation = False
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to perform ["+sAction+"] operation on ["+sItem+"] Tab.")
	Else
		Fn_SISW_UI_RACTabFolderWidget_Operation = True
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully perform ["+sAction+"] operation on ["+sItem+"] Tab.")
	End If
	
	Set  objItem=Nothing
	Set  objIntNoOfObjects=Nothing
	Set objSelectType=Nothing
	
End Function
'#########################################################################################################
'###	FUNCTION NAME	:	Function Fn_SISW_UI_JavaTableGetCellData
'###	Note : Depricated Function :	Note: Use Fn_SISW_UI_JavaTable_Operations
'###	EXAMPLE			:	 Fn_SISW_UI_JavaTableGetCellData("Fn_SISW_RDV_MSM_SearchCriteriaOperations", JavaWindow("RDV_StructureManager").JavaTable("MSM_ScopesSearchCriteriaTable"), 0, "Item Type")
'#############################################################################################################
Public Function Fn_SISW_UI_JavaTableGetCellData(sCallerFunction, objJavaTable, iRow, iCol)
	bGblFailedFunctionName = sCallerFunction
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"),  "Depricated function : Fn_SISW_UI_JavaTableGetCellData > Calling Function Fn_SISW_UI_JavaTable_Operations")
	Fn_SISW_UI_JavaTableGetCellData = Fn_SISW_UI_JavaTable_Operations(sCallerFunction, "GetCellData", objJavaTable , "", "GetProperty", "", iRow, iCol, "", "", "")
End Function
'#########################################################################################################
'###
'###	FUNCTION NAME	:	Function Fn_SISW_UI_GetRealPropertyName
'###
'###	DESCRIPTION		:	This function is to return Real property name for diplay name. 
'###
'###	PARAMETERS		:	1. sProeprtyDisplayName : Proeprty Name
'###
'###	HISTORY			:	AUTHOR					DATE			VERSION
'###
'###	Return Value	:	Text value / False
'###
'###	CREATED BY		:	Koustubh W				12-09-12		1.0
'###
'###	EXAMPLE			:	msgbox Fn_SISW_UI_GetRealPropertyName("Item Type")
'#############################################################################################################
Private Function Fn_SISW_UI_GetRealPropertyName(sProeprtyDisplayName)
	Fn_SISW_UI_GetRealPropertyName = False
	'bGblFailedFunctionName = "Fn_SISW_UI_GetRealPropertyName"
	Select Case trim(sProeprtyDisplayName)
		Case "Type", "Object Type"
			Fn_SISW_UI_GetRealPropertyName = "object_type"
		Case "Owner"
			Fn_SISW_UI_GetRealPropertyName = "owning_user"
		Case "Object"
			Fn_SISW_UI_GetRealPropertyName = "object_string"
		Case "Group ID"
			Fn_SISW_UI_GetRealPropertyName = "owning_group"
		Case "Last Modified Date", "Date Modified"
			Fn_SISW_UI_GetRealPropertyName = "last_mod_date"
		Case "Checked-Out"
			Fn_SISW_UI_GetRealPropertyName = "checked_out"
		Case "Release Status"
			Fn_SISW_UI_GetRealPropertyName = "release_status_list"
		Case "Checked-Out By"
			Fn_SISW_UI_GetRealPropertyName = "checked_out_user"
		Case "Project IDs"
			Fn_SISW_UI_GetRealPropertyName = "project_ids"
		Case "Checked-Out Date"
			Fn_SISW_UI_GetRealPropertyName = "checked_out_date"			
		Case "Classified"				
			Fn_SISW_UI_GetRealPropertyName = "ics_classified"			
		Case "Classified in"				
			Fn_SISW_UI_GetRealPropertyName = "ics_subclass_name"	
		Case "Description"
			Fn_SISW_UI_GetRealPropertyName = "object_desc"	
		Case "Name"
			Fn_SISW_UI_GetRealPropertyName = "object_name"	
		Case "License List"
			Fn_SISW_UI_GetRealPropertyName = "license_list"
		Case "BOM Line"
			Fn_SISW_UI_GetRealPropertyName = "bl_indented_title"
		Case "BOM Line Name"
			Fn_SISW_UI_GetRealPropertyName = "bl_line_name"			
		Case "Item Type"
			Fn_SISW_UI_GetRealPropertyName = "bl_item_object_type"
		Case "Item Description"
			Fn_SISW_UI_GetRealPropertyName = "bl_item_object_desc"
		Case "Item Name"
			Fn_SISW_UI_GetRealPropertyName = "bl_item_object_name"
		'Shreyas[13/09/2012] Added new cases
		Case "Object Name"
			Fn_SISW_UI_GetRealPropertyName="object_name"
		Case "Logged Date"
			Fn_SISW_UI_GetRealPropertyName="fnd0LoggedDate"
		Case "Event Type Name"
			Fn_SISW_UI_GetRealPropertyName="fnd0EventTypeName"
		Case "Primary Object ID"
			Fn_SISW_UI_GetRealPropertyName="fnd0PrimaryObjectID"
		Case "Primary Object Revision ID"
			Fn_SISW_UI_GetRealPropertyName="fnd0PrimaryObjectRevID"
		Case "User ID"
			Fn_SISW_UI_GetRealPropertyName="fnd0UserId"
		Case "Group Name"
			Fn_SISW_UI_GetRealPropertyName="fnd0GroupName"
		Case "Role Name"
			Fn_SISW_UI_GetRealPropertyName="fnd0RoleName"
		Case "Sequence ID"
			Fn_SISW_UI_GetRealPropertyName="sequence_id"
		Case "Change ID"
			Fn_SISW_UI_GetRealPropertyName="fnd0ChangeID"
		Case "Reason"
			Fn_SISW_UI_GetRealPropertyName="fnd0Reason"
		Case "Secondary Object ID"
			Fn_SISW_UI_GetRealPropertyName="fnd0SecondaryObjectID"
		Case "Complying Objects"
			Fn_SISW_UI_GetRealPropertyName="fnd0complying_objects"
		Case "Defining Objects"
			Fn_SISW_UI_GetRealPropertyName="fnd0defining_objects"
		Case "Trace Link"
			Fn_SISW_UI_GetRealPropertyName="FND_TraceLink"
		Case Else
			Fn_SISW_UI_GetRealPropertyName = sProeprtyDisplayName
	End Select
End Function
'#########################################################################################################
'###
'###	FUNCTION NAME	:	Function Fn_SISW_UI_GetDisplayedRelation
'###
'###	DESCRIPTION		:	This function is to return Relation display value for real value. 
'###
'###	PARAMETERS		:	1. sRelationRealValue : Relation Value
'###
'###	HISTORY			:	AUTHOR					DATE			VERSION
'###
'###	Return Value	:	Text value / False
'###
'###	CREATED BY		:	Koustubh W				12-09-12		1.0
'###
'###	EXAMPLE			:	msgbox Fn_SISW_UI_GetDisplayedRelation("Fnd0ListsParamReqments")
'#############################################################################################################
Private Function Fn_SISW_UI_GetDisplayedRelation(sRelationRealValue)
	'bGblFailedFunctionName = "Fn_SISW_UI_GetDisplayedRelation"
	Select Case trim(sRelationRealValue)
			Case "Fnd0ListsParamReqments"
				Fn_SISW_UI_GetDisplayedRelation = "Standard Notes Lists"
			Case "CMHasProblemItem"
				Fn_SISW_UI_GetDisplayedRelation = "Problem Items"
			Case "CMHasImpactedItem"
				Fn_SISW_UI_GetDisplayedRelation = "Impacted Items"
			Case "CMReferences"
				Fn_SISW_UI_GetDisplayedRelation = "Reference Items"
			Case "IMAN_reference"
				Fn_SISW_UI_GetDisplayedRelation = "References"
			Case "IMAN_specification"
				Fn_SISW_UI_GetDisplayedRelation = "Specifications"
			Case "contents"
				Fn_SISW_UI_GetDisplayedRelation = "Contents"
			Case "revision_list"
				Fn_SISW_UI_GetDisplayedRelation = "Revisions"
			Case "IMAN_master_form"
				' aaded s at the end - snehal salunkhe - 3-Apr-12 
				Fn_SISW_UI_GetDisplayedRelation = "Item Masters"
			Case "IMAN_classification"
				Fn_SISW_UI_GetDisplayedRelation = "Classification"
			Case "TC_Attaches"
				Fn_SISW_UI_GetDisplayedRelation = "Attaches"
			Case "IMAN_Rendering"
				Fn_SISW_UI_GetDisplayedRelation = "Rendering"
			Case "IMAN_manifestation"
				Fn_SISW_UI_GetDisplayedRelation = "Manifestations"
			Case "IMAN_aliasid"
				Fn_SISW_UI_GetDisplayedRelation = "Alias IDs"
			Case "release_status_list"
				Fn_SISW_UI_GetDisplayedRelation = "Release Status"
			''Changed By DhananjayN from "Custom Notes Lists"
			Case "Fnd0ListsCustomNotes"
				Fn_SISW_UI_GetDisplayedRelation = "Custom Requirements Lists"
			Case "Cdm0ListsCorspRefItems"
				Fn_SISW_UI_GetDisplayedRelation = "Contracts"
			Case "IMAN_external_object_link" 'Added by Nilesh on 10-Aug-2012
				Fn_SISW_UI_GetDisplayedRelation="External Proxy Relation"
			Case Else
				Fn_SISW_UI_GetDisplayedRelation = sRelationRealValue
		End Select
End Function
'-----------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_SISW_UI_Twistie_Operations

'Description		:	This function to perform operations on Twistie control.

'Parameters		    :	1. sFunctionName:
'								2. sAction		:
'								3. objContainer : Prent UI Component or JavaTable object
'								4. sTwistie 	: Twistie Control name
'								5. sTwistieText	: Twistie Control text
'								6. sTwistieStatic : Static Text Against Twistie

'Returns        : 	TRUE \ FALSE

'Examples			:	Call Fn_SISW_UI_Twistie_Operations("", "Expand", JavaWindow("ServicePlanner"), "Twistie", "Fault","StaticText")
'								Call Fn_SISW_UI_Twistie_Operations("", "Collapse", JavaWindow("ServicePlanner"), "Twistie", "References","StaticText") 

'History			:

'	Developer Name			Date		 	     Rev. No.	 	Reviewer					 	Changes Done	
'	Sachin 			  		 		19-Nov-2012	 	1.0															Created
'-----------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_UI_Twistie_Operations(sFunctionName, sAction, objContainer, sTwistie, sTwistieText,sTwistieStatic)
	Dim sFuncLog, objTwistie, objTwistieText
	bGblFailedFunctionName = sFunctionName
	Set objTwistie = objContainer.JavaObject(sTwistie)
	Set objTwistieText = objContainer.JavaStaticText(sTwistieStatic)
	Fn_SISW_UI_Twistie_Operations = False

	objTwistieText.SetTOProperty "label", sTwistieText
	sFuncLog = sFunctionName & " > Fn_SISW_UI_Twistie_Operations : [ " & sTwistieText & " Twistie Control ] : Action = " & sAction & " : "
	If objTwistie.Exist(5)= False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & " : FAIL : Does not exist.")
		Call ExitFromUI(sFuncLog)
		Set objTwistie = Nothing 
		Set objTwistieText = Nothing
		Exit Function
	End If

	Select Case sAction
		Case "Expand"
			If cBool(objTwistie.Object.isExpanded()) = False Then
				objTwistie.Click 1, 1,"LEFT"
				wait 2
				If cBool(objTwistie.Object.isExpanded()) = True Then
					Fn_SISW_UI_Twistie_Operations = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & "PASS : Successfully expanded." )
				Else
					Fn_SISW_UI_Twistie_Operations = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & "FAIL : Failed to expand." )
				End If
			Else
				Fn_SISW_UI_Twistie_Operations = True
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & "PASS : Successfully expanded." )
			End If
		
		Case "Collapse"
			If cBool(objTwistie.Object.isExpanded()) = True Then
				objTwistie.Click 1, 1,"LEFT"
				wait 2
				If cBool(objTwistie.Object.isExpanded()) = False Then
					Fn_SISW_UI_Twistie_Operations = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & "PASS : Successfully collapsed." )
				Else
					Fn_SISW_UI_Twistie_Operations = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & "FAIL : Failed to collapse." )
				End If
			Else
				Fn_SISW_UI_Twistie_Operations = True
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & "PASS : Successfully collapsed." )
			End If
	End Select
	Set objTwistie = Nothing 
	Set objTwistieText = Nothing
End Function
'*******************************************************************************************************************
''Function Name		 	:	Fn_SISW_UI_Object_Operations
'
''Description		    :  	This function is Use to check Existance of given Object.

''Parameters		    :	1. sFunctionName: Valid Function Name
''				    		2. sAction
''				    		3. objReferencePath: Valid Java Dialog	
''				    		4. iTimeOut: Time out time in seconds.
								
''Return Value		    :  	True \ false
'
''Examples		     	:	Fn_SISW_UI_Object_Operations("Fn_SISW_RAC_UI_DateControl_SetDate","Create", objDateControl,"") - Depricated
''Examples		     	:	Fn_SISW_UI_Object_Operations("Fn_SISW_RAC_UI_DateControl_SetDate","Enabled", objDateControl, SISW_MAX_TIMEOUT)
''Examples		     	:	Fn_SISW_UI_Object_Operations("Fn_SISW_RAC_UI_DateControl_SetDate","Exist", objDateControl,"")

'History:
'	Developer Name				Date			Rev. No.	Reviewer		Changes Done	
'*******************************************************************************************************************
'	Koustubh Watwe		        18-Oct-2012	    1.0			Self	
'*******************************************************************************************************************
Public Function Fn_SISW_UI_Object_Operations(sFunctionName, sAction, objReferencePath, iTimeOut)
	Dim sFuncLog
	bGblFailedFunctionName = sFunctionName
	sFuncLog = sFunctionName + " > Fn_SISW_UI_Object_Operations : [ " &  objReferencePath.toString & " ] : Action = " & sAction & " : "
	Fn_SISW_UI_Object_Operations = False
	If iTimeOut = "" Then
		iTimeOut = SISW_DEFAULT_TIMEOUT
	Else
		iTimeOut = cInt(iTimeOut)
	End If
	
	Select Case sAction
		Case "Create"
			If Fn_SISW_UI_Object_Operations(sFunctionName, "Enabled", objReferencePath, iTimeOut) Then
				Set Fn_SISW_UI_Object_Operations = objReferencePath
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),sFuncLog & "New Object created for " & objReferencePath.toString & " in Function ")
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),sFuncLog & "Object " & objReferencePath.toString & " is not enable of Function ")
				Call ExitFromUI(sFuncLog & "Object " & objReferencePath.toString & " is not enable of Function ") 
			End If
		Case "Exist"
			Fn_SISW_UI_Object_Operations = objReferencePath.Exist(iTimeOut)
			If Fn_SISW_UI_Object_Operations Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),sFuncLog & " object is exist.")
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),sFuncLog & "object does not exist.")
			End If
		Case "Enabled"
			If Fn_SISW_UI_Object_Operations(sFunctionName, "Exist", objReferencePath, iTimeOut) Then
				If objReferencePath.GetROProperty("enabled") = "1"  OR objReferencePath.GetROProperty("enabled") = True Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & "object is exists and enabled.")
					Fn_SISW_UI_Object_Operations = True
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),sFuncLog & "object is not enabled.")
				End If
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),sFuncLog & "does not exist.")
			End If
		Case "Activate"
			objReferencePath.Activate
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),sFuncLog & " object activated.")
			Fn_SISW_UI_Object_Operations = True
		Case "Close"
			objReferencePath.Close
			Wait SISW_MICROLESS_TIMEOUT
			Fn_SISW_UI_Object_Operations = True
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),sFuncLog & " object is closed.")
	End Select
End Function
'*******************************************************************************************************************
''Function Name		 	:	Fn_SISW_UI_CustomizedComboBox_Operations
'
''Description		    :  	This function is used to select row from TableCombo JavaObject.

''Parameters		    :	1. sFunctionName - caller function name
''							2. ComboboxType - Custom combo box type
''							3. sAction
''							4. objContainer : Parent UI Component or JavaTable object
''							5. sTableCombo : TableCombo Object's Name
''							6. sColumn : Column Name.
''							7. sValue : Cell value.

''Return Value		    :  	True \ False
'
''Examples		     	:	Call Fn_SISW_UI_CustomizedComboBox_Operations("", "LovDisplayer","Select", objTableCombo, "DRD", "", "LD")
''Examples		     	:	Call Fn_SISW_UI_CustomizedComboBox_Operations("", "TableCombo","Select", objTableCombo, "DRD", "", "*")
''Examples		     	:	Call Fn_SISW_UI_CustomizedComboBox_Operations("", "LovDisplayer", "Select", objTableCombo, "DRD", "Corp Group", "LD")
''Examples		     	:	Call Fn_SISW_UI_CustomizedComboBox_Operations("", "TableCombo", "Select", objTableCombo, "", "Corp Group", "LD")
''Examples		     	:	Call Fn_SISW_UI_CustomizedComboBox_Operations("", "TableCombo", "Select", objTableCombo, "", "", "")
''Examples		     	:	Call Fn_SISW_UI_CustomizedComboBox_Operations("", "LovDisplayer", "GetText", objTableCombo, "", "", "")
''Examples		     	:	Call Fn_SISW_UI_CustomizedComboBox_Operations("", "TableCombo", "GetText", objTableCombo, "DRD", "", "")
''Examples		     	:	Call Fn_SISW_UI_CustomizedComboBox_Operations("", "TableCombo",  "VerifyValue",  JavaWindow("ProductMasterManagerAdmin"),  "CODPWOState", "", "PCMP")
''Examples		     	:	Call Fn_SISW_UI_CustomizedComboBox_Operations("", "LovDisplayer", "VerifyValue",  objTableCombo,  "", "", "PCMP~PRT1")
''Examples		     	:	Call Fn_SISW_UI_CustomizedComboBox_Operations("", "TableCombo", "GetContents", JavaWindow("ProductMasterManager"), "ProductType", "", "")
''Examples		     	:	Call Fn_SISW_UI_CustomizedComboBox_Operations("", "LovDisplayer", "GetContents", objTableCombo, "", "Vehicle Line", "")
''Examples		     	:	Call Fn_SISW_UI_CustomizedComboBox_Operations("", "TableCombo", "GetContents", JavaWindow("ProductMasterManager"), "ProductType", "Product Type~Vehicle Line", "")
''Examples		     	:	Call Fn_SISW_UI_CustomizedComboBox_Operations("", "TableCombo", "TypeAndSelect", objTableCombo, "", "", "AutoTest1")

'History			:
'-------------------------------+---------------+---------------+-------------------------------+-----------------------------------
'	Developer Name				|	Date		|	Rev. No.	|	Reviewer					|	Changes Done
'-------------------------------+---------------+---------------+-------------------------------+-----------------------------------
'	Koustubh Watwe		 		|22-Oct-2012	|	1.0			|	Koustubh Watwe		 		| Created Fn_SISW_UI_CustomComboBox_Operations
'-------------------------------+---------------+---------------+-------------------------------+-----------------------------------
'	Koustubh Watwe		 		|19-Jun-2014	|	1.0			|	Koustubh Watwe		 		| Deprecated function Fn_SISW_UI_CustomComboBox_Operations, added new function Fn_SISW_UI_CustomizedComboBox_Operations
'-------------------------------+---------------+---------------+-------------------------------+-----------------------------------
'	Koustubh Watwe		 		|19-Nov-2015	|	1.0			|	Koustubh Watwe		 		| Added case TypeAndSelect
'-------------------------------+---------------+---------------+-------------------------------+-----------------------------------
Public Function Fn_SISW_UI_CustomComboBox_Operations(sFunctionName, sAction, objContainer, sTableCombo, sColumn, sValue)
	' Deprecated function
	bGblFailedFunctionName = sFunctionName
 	Fn_SISW_UI_CustomComboBox_Operations = Fn_SISW_UI_CustomizedComboBox_Operations(sFunctionName, "TableCombo", sAction, objContainer, sTableCombo, sColumn, sValue)
End Function

Public Function Fn_SISW_UI_CustomizedComboBox_Operations(sFunctionName, ComboboxType,  sAction, objContainer, sTableCombo, sColumn, sValue)
	Dim sFuncLog, objTableCombo, objTables, iColumnsCount, sProperties, sPropertyValues
	Dim iTreeItemCount, iTreeCounter, arrItems, arrColumns, iColCounter, arrValues
	bGblFailedFunctionName = sFunctionName
	Fn_SISW_UI_CustomizedComboBox_Operations = False
	If typeName(sTableCombo) = "String" Then
		Set objTableCombo = objContainer.javaObject(sTableCombo)
		sFuncLog = sFunctionName & " > Fn_SISW_UI_CustomizedComboBox_Operations : [ " & objContainer.toString() & " ] : [ TableCombo : " & sTableCombo & " ] : Action = " & sAction & " : "
	Else
		Set objTableCombo = sTableCombo
		Set objContainer = JavaWindow("DefaultWindow")
		sFuncLog = sFunctionName & " > Fn_SISW_UI_CustomizedComboBox_Operations : [ TableCombo : " & sTableCombo.toString() & " ] : Action = " & sAction & " : "
	End If
	If Fn_SISW_UI_Object_Operations("Fn_SISW_UI_CustomizedComboBox_Operations","Exist", objTableCombo, SISW_MIN_TIMEOUT) Then
		Select Case sAction
			Case "Select", "VerifyValue","TypeAndSelect"
				
				Select case sValue
					Case ""
						objTableCombo.Object.clearSelection
						Fn_SISW_UI_CustomizedComboBox_Operations = true
					Case "%", "*"
						objTableCombo.Object.clearSelection
						'objTableCombo.Object.setSelectedValue sValue
						objTableCombo.type sValue
						Fn_SISW_UI_CustomizedComboBox_Operations = true
					Case Else
						If sAction = "TypeAndSelect" Then
							objTableCombo.Object.clearSelection
							objTableCombo.type sValue
						Else
							objTableCombo.Click cInt(objTableCombo.GetROProperty("width")) - 10,  cInt(objTableCombo.GetROProperty("height")) - 10, "LEFT"
						End If
						wait 5
						Select Case ComboboxType
							Case "LovDisplayer"
								'Set objTables =  Fn_SISW_UI_Object_GetChildObjects("Fn_SISW_UI_CustomizedComboBox_Operations", objContainer , "Class Name~path$RegularExpression", "JavaTable~Table;Shell;Shell;.*")
								sProperties = "Class Name~path$RegularExpression"
								sPropertyValues =  "JavaTree~(Tree;)?.*Shell;Shell;.*"
								Set objTables =  Fn_SISW_UI_Object_GetChildObjects("Fn_SISW_UI_CustomizedComboBox_Operations", objContainer , sProperties, sPropertyValues )
								If TypeName(objTables) <> "Nothing"  Then
									Set objTables =  objTables(0)
									'bResult = Fn_SISW_RAC_UI_JavaTable_Operations("Fn_SISW_UI_CustomizedComboBox_Operations", "RowSelect", objTables(0) ,"", "CompositeColumn", sColumn, sValue, "", "", "")
		                            iTreeItemCount = cInt(objTables.Object.getItemCount())
									For iTreeCounter = 0 to iTreeItemCount - 1
		                                If cstr(objTables.GetColumnValue( "#" & iTreeCounter , 0)) = sValue Then
	                                		objTables.Activate "#" & iTreeCounter
	                                		Fn_SISW_UI_CustomizedComboBox_Operations = True
											exit for
										End If
									Next
								Else
									sFuncLog = sFuncLog & " Failed to display dropdown tree : "
								End If
							Case "TableCombo"
								'Set objTables =  Fn_SISW_UI_Object_GetChildObjects("Fn_SISW_UI_CustomizedComboBox_Operations", objContainer , "Class Name~path$RegularExpression", "JavaTable~Table;Shell;Shell;.*")
								sProperties = "Class Name~path$RegularExpression"
								sPropertyValues =  "JavaTable~Table;.*Shell;Shell;.*"
								Set objTables =  Fn_SISW_UI_Object_GetChildObjects("Fn_SISW_UI_CustomizedComboBox_Operations", objContainer , sProperties, sPropertyValues )
								If TypeName(objTables) <> "Nothing"  Then
									Set objTables =  objTables(0)
									'bResult = Fn_SISW_RAC_UI_JavaTable_Operations("Fn_SISW_UI_CustomizedComboBox_Operations", "RowSelect", objTables(0) ,"", "CompositeColumn", sColumn, sValue, "", "", "")
		                            iTreeItemCount = Fn_UI_Object_GetROProperty("Fn_SISW_UI_CustomizedComboBox_Operations", objTables,"rows")
									For iTreeCounter = 0 to iTreeItemCount - 1
		                                If objTables.GetCellData(iTreeCounter,0) = sValue Then
		                                    objTables.ClickCell iTreeCounter, 0
											Fn_SISW_UI_CustomizedComboBox_Operations = True
											Exit for
										End If
									Next
								Else
									sFuncLog = sFuncLog & " Failed to display dropdown tree : "
								End If
						End Select
				End Select
				sFuncLog = sFuncLog & " Set value = " & sValue & " : "
			Case "GetText"
				Fn_SISW_UI_CustomizedComboBox_Operations = cstr(objTableCombo.Object.getSelectedDisplayValue())
			Case Else
		End Select
	End If
	If Fn_SISW_UI_CustomizedComboBox_Operations <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & "PASS : Executed successfully.")		
	Else
		If sAction <> "VerifyValue" Then
			Call ExitFromUI(sFuncLog & "FAIL : Execution Failed.")
		End If
	End If
	Set objTables =  Nothing
	Set arrItems = Nothing
	Set objTableCombo = Nothing
End Function
'*******************************************************************************************************************
''Function Name		 	:	Fn_SISW_UI_Object_GetChildObjects
'
''Description		    :  	This function is Use to get child objects of specified component descriptively.

''Parameters		    :	1. sFunctionName
''							2. objComponent
''							3. sProperties : ~ separated list of property names
''							4. sValues : ~ separated list of property values

''Return Value		    :  	Nothing \ Array of Objects
'
''Examples		     	:	Set objDiag =  Fn_SISW_UI_Object_GetChildObjects("", objDialog, "Class Name~tagname", "JavaObject~ImageHyperlink")
''Examples		     	:	Set objDiag =  Fn_SISW_UI_Object_GetChildObjects("", objDialog, "Class Name$RegularExpression~tagname", "Java.*~ImageHyperlink")

'History:
'	Developer Name				Date			Rev. No.	Reviewer		Changes Done	
'*******************************************************************************************************************
'	Koustubh Watwe		        22-Oct-2012	    1.0			Self	
'*******************************************************************************************************************
Public Function Fn_SISW_UI_Object_GetChildObjects(sFunctionName, objComponent, sProperties, sValues)
	Dim objSelectType, intNoOfObjects, sFuncLog
	Dim arrProperties, arrValues, arrCounter
	bGblFailedFunctionName = sFunctionName
	Set Fn_SISW_UI_Object_GetChildObjects = Nothing
	Set objSelectType = Description.Create()
	arrProperties = split(sProperties, "~")
	arrValues = split(sValues, "~") 
	sFuncLog = sFunctionName & " > Fn_SISW_UI_Object_GetChildObjects : [ " & objComponent.toString() & " ] : "

	For arrCounter = 0 to UBound(arrProperties)
		If instr(arrProperties(arrCounter),"$RegularExpression") > 0 Then
			objSelectType(replace(arrProperties(arrCounter), "$RegularExpression", "")).RegularExpression = True
			objSelectType(replace(arrProperties(arrCounter), "$RegularExpression", "")).value = arrValues(arrCounter)
		Else
			objSelectType(arrProperties(arrCounter)).value = arrValues(arrCounter)
		End If
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & "Setting Properties [" + arrProperties(arrCounter) & " = " & arrValues(arrCounter) & " ].")		
	Next
	If Fn_SISW_UI_Object_Operations("Fn_SISW_UI_Object_GetChildObjects","Exist", objComponent, SISW_MIN_TIMEOUT) Then
		Set  intNoOfObjects = objComponent.ChildObjects(objSelectType)
		If intNoOfObjects.count <> 0 Then
			Set Fn_SISW_UI_Object_GetChildObjects = intNoOfObjects
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & "PASS: Successfully found [" & intNoOfObjects.count & "] child objects.")		
		Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & "FAIL: No child objects found.")		
		End If
	End If
	Set intNoOfObjects = Nothing
	Set objSelectType = Nothing
End Function
'-----------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_SISW_UI_JavaTab_Operations

'Description		:	This function to perform operations on JavaTab component.

'Parameters		    :	1. sFunctionName	: Caller function's name
'						2. sAction			: Action to be performed
'						3. objJavaDialog 	: Prent UI Component or JavaCheckBox object
'						4. sTabObjectName 	: JavaTab Control name
'						5. sItem			: Tab Text

'Return Value		: 	TRUE \ FALSE

'Pre-requisite		:	JavaTab must be displayed in Object repository

'Examples			:	Call Fn_SISW_UI_JavaTab_Operations("", "Select", JavaWindow("ProductMasterManager"), "AuthorizationTab",  "EW11397")
'					:	Call Fn_SISW_UI_JavaTab_Operations("", "Exist", JavaWindow("ProductMasterManager").JavaTab("AuthorizationTab"), "", "EW11397")

'History			:
'-------------------------------+---------------+---------------+-------------------------------+-----------------------------------
'	Developer Name				|	Date		|	Rev. No.	|	Reviewer					|	Changes Done	
'-------------------------------+---------------+---------------+-------------------------------+-----------------------------------
'	Koustubh Watwe		 		|10-Jul-2013	|	1.0			|	Koustubh Watwe		 		| 	Created
'-------------------------------+---------------+---------------+-------------------------------+-----------------------------------
'	Shweta Rathod		 		|25-Aug-2016	|	1.0			|	Koustubh Watwe		 		| 	Added New case "close"
'-------------------------------+---------------+---------------+-------------------------------+-----------------------------------
Public Function Fn_SISW_UI_JavaTab_Operations(sFunctionName, sAction, objJavaDialog, sTabObjectName, sItem)
	Dim sFuncLog, objTab, sitems, iItemCount, iCounter
	'Object Creation
	bGblFailedFunctionName = sFunctionName
	Fn_SISW_UI_JavaTab_Operations = False
	If sTabObjectName <> "" Then
		Set objTab = objJavaDialog.JavaTab(sTabObjectName)
		sFuncLog = sFunctionName + " > Fn_SISW_UI_JavaTab_Operations  : [ " &  objJavaDialog.toString & " ] : [ " +  sTabObjectName + " ] : Action = " & sAction & " : "
	Else
		Set objTab = objJavaDialog
		sFuncLog = sFunctionName + " > Fn_SISW_UI_JavaTab_Operations  : [ " +  objTab.toString + " ] : Action = " & sAction & " : "
	ENd IF
	
	'Verify JavaTab object exists
	If Fn_SISW_UI_Object_Operations("Fn_SISW_UI_JavaTab_Operations","Exist", objTab, SISW_MIN_TIMEOUT) = False Then	
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & " : FAIL : Does not exist.")
		Set objTab = Nothing 
		Exit Function
	End If
	
	Select case sAction
		Case "Select"
			If Fn_SISW_UI_JavaTab_Operations(sFunctionName, "Exist", objJavaDialog, sTabObjectName, sItem) = True Then
				On error resume next
				objTab.Select sItem
				If Err.Number <> 0 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & "FAIL : Failed to find JavaTab [ " & sItem & " ].")
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & "FAIL : Error Description - " & Err.Description )
                    On Error GoTo 0
				Else
					Fn_SISW_UI_JavaTab_Operations = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & "PASS : Sucessfully set JavaTab [ " & sItem & " ].")
				End If				
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & "FAIL : Failed to find JavaTab [ " & sItem & " ].")
			End If
		Case "Exist"
				Select Case objTab.Object.getClass().toString()
					Case "class javax.swing.JTabbedPane",_
						 "class com.teamcenter.rac.pse.search.PSESearchResultDialog$PSESearchTabbedPane",_
						 "class com.teamcenter.rac.classification.common.G4MTabbedPane"
							iItemCount = cInt(objTab.Object.getTabCount())
							For iCounter = 0 to iItemCount - 1
								If sItem = objTab.Object.getTitleAt(iCounter) Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & "PASS : Sucessfully verified JavaTab [ " & sItem & " ].")
									Fn_SISW_UI_JavaTab_Operations = True
									Exit for
								End If
							Next

					Case Else
							Set sitems = objTab.Object.getItems()
							'- - - - - - - - This condition put to handle BMIDE Tab
							If instr(1,lCase(objTab.Object.toString()),"wrong thread") then
								Fn_SISW_UI_JavaTab_Operations = True
								Exit function
							End if
							'- - - - - - 
							iItemCount = cInt(objTab.Object.getItemCount())
							For iCounter = 0 to iItemCount - 1
								If sItem = sitems.mic_arr_get(iCounter).getText() Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & "PASS : Sucessfully verified JavaTab [ " & sItem & " ].")
									Fn_SISW_UI_JavaTab_Operations = True
									Exit for
								End If
							Next
				End Select
		Case "close"
			If Fn_SISW_UI_JavaTab_Operations(sFunctionName, "Exist", objJavaDialog, sTabObjectName, sItem) = True Then
				objTab.CloseTab sItem
				Fn_SISW_UI_JavaTab_Operations = True
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & "PASS : Sucessfully close JavaTab [ " & sItem & " ].")
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & "FAIL : Failed to find JavaTab [ " & sItem & " ].")
			End If
			
	End Select
	'Clear memory of JavaTab object.
	Set objTab = Nothing 
End Function
'*******************************************************************************************************************
'Function Name		 	:	Fn_SISW_UI_JavaButton_Operations
'
'Description		    :  	This function is used to perform operations on JavaButton object.

'Parameters			    :	1. sFunctionName : Valid Function name, 
'							2. sAction
'							3. objJavaDialog
'							4. sJavaButton   : Valid Button name

'Return Value		    :  	True \ False
'
'Examples		     	:	Call Fn_SISW_UI_JavaButton_Operations("Fn_TeamcenterLogin", "Click", JavaWindow("Teamcenter Login"),"Login")
'Examples		     	:	Call Fn_SISW_UI_JavaButton_Operations("Fn_TeamcenterLogin", "Click", JavaWindow("Teamcenter Login").JavaButton("Login"),"")

'History:
'	Developer Name				Date			Rev. No.	Reviewer		Changes Done	
'*******************************************************************************************************************
'	Koustubh Watwe		        18-Oct-2012	    1.0			Self
'	Koustubh Watwe		        29-Feb-2016	    1.0			Self			Added case for depricated function call Fn_Button_Click
'*******************************************************************************************************************
Public Function Fn_SISW_UI_JavaButton_Operations(sFunctionName, sAction, objJavaDialog, sJavaButton)
	Dim objJavaButton, sFuncLog, objDeviceReplay
	bGblFailedFunctionName = sFunctionName
	Fn_SISW_UI_JavaButton_Operations = False
	'Object Creation
	If sJavaButton <> "" Then
		Set objJavaButton = objJavaDialog.JavaButton(sJavaButton)
		sFuncLog = sFunctionName + " > Fn_SISW_UI_JavaButton_Operations  : [ " &  objJavaDialog.toString & " ] : [ " +  objJavaButton.toString + " ] : Action = " & sAction & " : "
	Else
		Set objJavaButton = objJavaDialog
		sFuncLog = sFunctionName + " > Fn_SISW_UI_JavaButton_Operations  : [ " +  objJavaButton.toString + " ] : Action = " & sAction & " : "
	ENd IF
	
	'Verify JavaButton object exists
	If Fn_SISW_UI_Object_Operations("Fn_SISW_UI_JavaButton_Operations", "Exist", objJavaButton,"") = False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & " : FAIL : Does not exist.")
		If sAction = "Fn_Button_Click" Then
			Call ExitFromUI(sFuncLog & " : FAIL : Does not exist.")
		End If
		Set objJavaButton = Nothing 
		Exit Function
	End If
	
	Select Case sAction
		Case "Click", "Fn_Button_Click" ' Depricated case Fn_Button_Click special case for function Fn_Button_Click
			objJavaButton.Click	
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & "PASS : Sucessfully clicked on JavaButton.")
			Fn_SISW_UI_JavaButton_Operations = True
		Case "DeviceReplay.Click"
			If sJavaButton <> "" Then
				objJavaDialog.Click 1, 1,"LEFT"
				wait SISW_MICRO_TIMEOUT
			End If
			Set objDeviceReplay = CreateObject("Mercury.DeviceReplay")
			objDeviceReplay.MouseMove (objJavaButton.GetROProperty("abs_x") + 5), (objJavaButton.GetROProperty("abs_y") + 5)
			objDeviceReplay.MouseClick  (objJavaButton.GetROProperty("abs_x") + 5), (objJavaButton.GetROProperty("abs_y") + 5), 0
			Fn_SISW_UI_JavaButton_Operations = True
		Case "Object.click"
			If objJavaButton.GetROProperty("enabled") = 1 Then
			   objJavaButton.Object.click
			   Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & "PASS : Sucessfully clicked on JavaButton.")
			   Fn_SISW_UI_JavaButton_Operations = True 
			End If
		Case "Object.setFocus"
			objJavaButton.Object.setFocus
			Set objDeviceReplay = CreateObject("Mercury.DeviceReplay")
			objDeviceReplay.PressKey (28)
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & "PASS : Sucessfully clicked on JavaButton.")
			Fn_SISW_UI_JavaButton_Operations = True 
		Case Else
	End Select
	'Clear memory of JavaButton object.
	Set objDeviceReplay = Nothing
	Set objJavaButton = Nothing 
End Function
'*******************************************************************************************************************
'Function Name		 	:	Fn_SISW_UI_JavaList_Operations
'
'Description		    :  	This function is used to perform operations on JavaButton object.

'Parameters			    :	1. sFunctionName 	: Valid Function name, 
'							2. sAction			: Action to be performed
'							3. objJavaDialog	: Parent dialog / javalist Object
'							4. sJavaList   		: JavaList Name
'							5. sValues   		: Value to be selected / verified
'							6. sColumns   		: --
'							7. sInstanceHandler : Instance Handler

'Return Value		    :  	True \ False
'
'Examples		     	:	Call Fn_SISW_UI_JavaList_Operations("Fn_TeamcenterLogin", "Select", JavaWindow("Teamcenter Login"),"Login","AutoTest", "", "")
'Examples		     	:	Call Fn_SISW_UI_JavaList_Operations("Fn_TeamcenterLogin", "Click", JavaWindow("Teamcenter Login"),"Login","AutoTest", "", "")
'Examples		     	:	Call Fn_SISW_UI_JavaList_Operations("Fn_TeamcenterLogin", "Activate", JavaWindow("Teamcenter Login"),"Login","AutoTest", "", "")
'Examples		     	:	Call Fn_SISW_UI_JavaList_Operations("Fn_TeamcenterLogin", "Exist", JavaWindow("Teamcenter Login"),"Login","AutoTest", "", "")
'Examples		     	:	Call Fn_SISW_UI_JavaList_Operations("Fn_TeamcenterLogin", "ExtendSelect", JavaWindow("Teamcenter Login"),"Login","AutoTest", "", "")
'Examples		     	:	Call Fn_SISW_UI_JavaList_Operations("Fn_TeamcenterLogin", "GetText", JavaWindow("Teamcenter Login"),"Login","", "", "")
'Examples		     	:	Call Fn_SISW_UI_JavaList_Operations("Fn_TeamcenterLogin", "GetContents", JavaWindow("Teamcenter Login"),"Login","", "", "")
'Examples		     	:	Call Fn_SISW_UI_JavaList_Operations("Fn_TeamcenterLogin", "GetIndex", JavaWindow("Teamcenter Login"),"Login","AutoTest", "", "")

'History:
'	Developer Name				Date			Rev. No.	Reviewer		Changes Done	
'*******************************************************************************************************************
'	Koustubh Watwe		        18-Jul-2013	    1.0			Self	
'*******************************************************************************************************************
'	Ganesh B		        	23-Apr-2014	    1.1							Added case "VerifyAscendingOrder"
'*******************************************************************************************************************
Public Function Fn_SISW_UI_JavaList_Operations(sFunctionName, sAction, objJavaDialog, sJavaList, sValues, sColumns, sInstanceHandler)
   Dim objJavaList, sFuncLog, arrSelectList, iCounter, iElecount, iInstanceCnt
   Dim bFlag, itr
   	bGblFailedFunctionName = sFunctionName
	Fn_SISW_UI_JavaList_Operations = False
	If sInstanceHandler = "" Then
		sInstanceHandler = "@"
	End If
	'Object Creation
	If sJavaList <> "" Then
		Set objJavaList = objJavaDialog.JavaList(sJavaList)
		sFuncLog = sFunctionName + " > Fn_SISW_UI_JavaList_Operations  : [ " &  objJavaDialog.toString & " ] : [ " +  objJavaList.toString + " ] : Action = " & sAction & " : "
	Else
		Set objJavaList = objJavaDialog
		sFuncLog = sFunctionName + " > Fn_SISW_UI_JavaList_Operations  : [ " +  objJavaList.toString + " ] : Action = " & sAction & " : "
	ENd IF
	
	'Verify JavaButton object exists
	If Fn_SISW_UI_Object_Operations("Fn_SISW_UI_JavaList_Operations", "Exist", objJavaList,"") = False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & " : FAIL : List does not exist.")
		Set objJavaList = Nothing 
		Exit Function
	End If
	
	Select Case sAction
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "GetIndex"
			Fn_SISW_UI_JavaList_Operations = -1
			' get total items from list
			iEelecount = objJavaList.GetROProperty("items count")
			iInstanceCnt = 1
			arrSelectList = split(sValues, sInstanceHandler)
			If uBound(arrSelectList) > 0 Then
				iInstanceCnt = arrSelectList(1)
			End If
			arrSelectList(0) = trim(arrSelectList(0))
			For iCounter = 0 To iEelecount - 1
				If objJavaList.GetItem(iCounter) <> "" Then
					If Trim(cstr(objJavaList.GetItem(iCounter))) = Trim(arrSelectList(0)) Then
						iF iInstanceCnt = 1 Then
							Fn_SISW_UI_JavaList_Operations = iCounter
							Exit For
						End If
						iInstanceCnt = iInstanceCnt - 1 
					End If
				End If
			Next
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "GetContents"
			iEelecount = objJavaList.GetROProperty("items count")
			For iCounter = 0 To iEelecount - 1
				If iCounter = 0 Then
					Fn_SISW_UI_JavaList_Operations = Trim(cstr(objJavaList.GetItem(iCounter)))
				Else
					Fn_SISW_UI_JavaList_Operations = Fn_SISW_UI_JavaList_Operations & "~" & Trim(cstr(objJavaList.GetItem(iCounter)))
				End If
			Next
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "GetText"
			Fn_SISW_UI_JavaList_Operations = objJavaList.GetROProperty("value")
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "Click"
			objJavaList.Click sValues, sColumns
			Fn_SISW_UI_JavaList_Operations = true
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "Activate"
			objJavaList.Activate sValues
			Fn_SISW_UI_JavaList_Operations = true
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "Exist"
			' get total items from list
			iEelecount = objJavaList.GetROProperty("items count")
			iInstanceCnt = 1
			arrSelectList = split(sValues, sInstanceHandler)
			If uBound(arrSelectList) > 0 Then
				iInstanceCnt = arrSelectList(1)
			End If
			arrSelectList(0) = trim(arrSelectList(0))
			For iCounter = 0 To iEelecount - 1
				If objJavaList.GetItem(iCounter) <> "" Then
					If Trim(cstr(objJavaList.GetItem(iCounter))) = Trim(arrSelectList(0)) Then
						iF iInstanceCnt = 1 Then
							Fn_SISW_UI_JavaList_Operations = True
							Exit For
						End If
						iInstanceCnt = iInstanceCnt - 1 
					End If
				End If
			Next
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "ExtendSelect"
			arrSelectList=Split(sValues,"~")
			' Select the element from list  
			For iCounter = 0 To Ubound(arrSelectList)
				objJavaList.ExtendSelect arrSelectList(iCounter)
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),sFuncLog & "Successfully Selected [" + arrSelectList(iCounter) + "]")
			Next
			Fn_SISW_UI_JavaList_Operations = true
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "VerifyAscendingOrder"
			arrSelectList=Split(sValues,"~")
			' Select the element from list  
			For iCounter = 0 To Ubound(arrSelectList)
			
				itr  = objJavaList.GetItemIndex(arrSelectList(iCounter))
				If iCounter <> 0 Then
					If itr > objJavaList.GetItemIndex(arrSelectList(iCounter-1)) Then
						bFlag = True
					Else
						bFlag = false
					End If
				End If
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),sFuncLog & "Successfully Selected [" + arrSelectList(iCounter) + "]")
			Next
			If bFlag = True AND iCounter - 1 = Ubound(arrSelectList) Then
				Fn_SISW_UI_JavaList_Operations = True
			Else
				Fn_SISW_UI_JavaList_Operations = False
			End If
			
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "Select"
			arrSelectList=Split(sValues,"~")
			' Select the element from list  
			For iCounter = 0 To Ubound(arrSelectList)
				if iCounter = 0 then
					objJavaList.Select arrSelectList(iCounter)												
				Else
					objJavaList.ExtendSelect arrSelectList(iCounter)
				End If 
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),sFuncLog & "Successfully Selected [" + arrSelectList(iCounter) + "]")
			Next
			Fn_SISW_UI_JavaList_Operations = true
	End Select
	If Fn_SISW_UI_JavaList_Operations <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),sFuncLog & "Executed successfully")
	End If
End Function
'*******************************************************************************************************************
'Function Name		 	:	Fn_SISW_UI_JavaEdit_Operations
'
'Description		    :  	This function is use to check existance of given object.

'Parameters			    :	1. sFunctionName : Valid Function name, 
'							2. sAction
'							3. objJavaDialog
'							4. sJavaEdit
'							5. sText

'Return Value		    :  	True \ false
'
'Examples		     	:	Call Fn_SISW_UI_JavaEdit_Operations("Fn_SISW_PMM_UserContextSettings", "Set",  objDialog, "EWO", "EWO_Name" )
'Examples		     	:	Call Fn_SISW_UI_JavaEdit_Operations("Fn_SISW_PMM_UserContextSettings", "Type",  objDialog.JavaEdit("EWO"),"", "EWO_Name" )

'History:
'	Developer Name				Date			Rev. No.	Reviewer		Changes Done	
'*******************************************************************************************************************
'	Koustubh Watwe		        18-Oct-2012	    1.0			Self	
'*******************************************************************************************************************
Function Fn_SISW_UI_JavaEdit_Operations(sFunctionName, sAction, objJavaDialog, sJavaEdit, sText)
	Dim objJavaEdit, sFuncLog
	bGblFailedFunctionName = sFunctionName
	Fn_SISW_UI_JavaEdit_Operations = False
	'Set an Edit Object on variable
	If sJavaEdit <> "" Then
		Set objJavaEdit= objJavaDialog.JavaEdit(sJavaEdit)
		sFuncLog = sFunctionName + "> Fn_SISW_UI_JavaEdit_Operations : [ " &  objJavaDialog.toString & " ] : [ " & objJavaEdit.toString & " ] : Action = " & sAction & " : "
	Else
		Set objJavaEdit= objJavaDialog
		sFuncLog = sFunctionName + "> Fn_SISW_UI_JavaEdit_Operations : [ " & objJavaEdit.toString & " ] : Action = " & sAction & " : "
	End If

	If Fn_SISW_UI_Object_Operations("Fn_SISW_UI_JavaEdit_Operations", "Exist", objJavaEdit,"") = False Then
		Fn_SISW_UI_JavaEdit_Operations= False
		Call ExitFromUI(sFuncLog & " : FAIL : Object does not exist.")
		Set objJavaEdit = Nothing 
		Exit Function
	End If
	Select Case sAction
		Case "Set"
			'Setting the editbox
			objJavaEdit.Set sText
			If objJavaEdit.GetROProperty("value") <> sText and Environment.Value("ProductName") = sQTPProductName OR Environment.Value("ProductName") = sUFTProductName Then  '' added by Ganesh to handle set method in UFT 
				objJavaEdit.Set ""
				objJavaEdit.Type sText
			End If
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & "PASS : Text " &sText  &"is Set/ Entered in JavaEditBox.")
			Fn_SISW_UI_JavaEdit_Operations= True
		Case "setText"
		'Setting the editbox
		objJavaEdit.Object.setText sText
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & "PASS : Text " &sText  &"is Set/ Entered in JavaEditBox.")
		Fn_SISW_UI_JavaEdit_Operations= True

		Case "Type"
			'Setting the editbox
			objJavaEdit.Type sText
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & "PASS : Text " &sText  &" is Set/ Entered in JavaEditBox.")
			Fn_SISW_UI_JavaEdit_Operations= True
		Case "GetText"
			Fn_SISW_UI_JavaEdit_Operations = objJavaEdit.getROProperty("value")
		Case "SetExt"
			'Setting the editbox
			objJavaEdit.Set sText
			If objJavaEdit.GetROProperty("value") <> sText and Environment.Value("ProductName") = sQTPProductName OR Environment.Value("ProductName") = sUFTProductName Then  '' added by Ganesh to handle set method in UFT 
				objJavaEdit.Set ""
				objJavaEdit.Set sText
			End If
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & "PASS : Text " &sText  &"is Set/ Entered in JavaEditBox.")
			Fn_SISW_UI_JavaEdit_Operations= True
	End Select
	Set objJavaEdit = Nothing 
End Function
'-----------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_SISW_UI_JavaCheckBox_Operations

'Description		:	This function to perform operations on JavaCheckBox component.

'Parameters		    :	1. sFunctionName: Caller function's name
'						2. sAction		: Action to be performed
'						3. objJavaDialog : Prent UI Component or JavaCheckBox object
'						4. sCheckBoxName 	: JavaCheckBox Control name
'						5. sValue	: value

'Return Value		: 	TRUE \ FALSE

'Pre-requisite		:	objTwistieText must be in format TwistieControlName_Label present in Object repository

'Examples			:	Fn_SISW_UI_JavaCheckBox_Operations("", "Set", objDialog, "Version", "ON")
'					:	Fn_SISW_UI_JavaCheckBox_Operations("", "Set", objDialog.JavaCheckBox("Precise"), "" , "OFF")

'History			:
'-------------------------------+---------------+---------------+-------------------------------+-----------------------------------
'	Developer Name				|	Date		|	Rev. No.	|	Reviewer					|	Changes Done	
'-------------------------------+---------------+---------------+-------------------------------+-----------------------------------
'	Koustubh Watwe		 		|15-Nov-2012	|	1.0			|	Koustubh Watwe		 		| 	Created
'-------------------------------+---------------+---------------+-------------------------------+-----------------------------------
Public Function Fn_SISW_UI_JavaCheckBox_Operations(sFunctionName, sAction, objJavaDialog, sCheckBoxName, sValue)
	Dim sFuncLog, objCheckBox
	bGblFailedFunctionName = sFunctionName
	'Object Creation
	If sCheckBoxName <> "" Then
		Set objCheckBox = objJavaDialog.JavaCheckBox(sCheckBoxName)
		sFuncLog = sFunctionName + " > Fn_SISW_UI_JavaCheckBox_Operations  : [ " &  objJavaDialog.toString & " ] : [ " +  objCheckBox.toString + " ] : Action = " & sAction & " : "
	Else
		Set objCheckBox = objJavaDialog
		sFuncLog = sFunctionName + " > Fn_SISW_UI_JavaCheckBox_Operations  : [ " +  objCheckBox.toString + " ] : Action = " & sAction & " : "
	ENd IF
	
	'Verify JavaCheckBox object exists
	If Fn_SISW_UI_Object_Operations("Fn_SISW_UI_JavaCheckBox_Operations", "Exist", objCheckBox,"") = False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & " : FAIL : Does not exist.")
		Fn_SISW_UI_JavaCheckBox_Operations = False
		Call ExitFromUI(sFuncLog & " : FAIL : Object does not exist.")
		Set objCheckBox = Nothing 
		Exit Function
	End If
	
	Select case sAction
		Case "Set"
			If uCase(trim(cstr(sValue))) = "TRUE" OR uCase(trim(cstr(sValue))) = "ON" Then
				objCheckBox.Set "ON" 
			Else
				objCheckBox.Set "OFF" 			
			End If
			Fn_SISW_UI_JavaCheckBox_Operations = True
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & "PASS : Sucessfully set JavaCheckBox [ " & sValue & " ].")
		Case Else
	End Select
	'Clear memory of JavaCheckBox object.
	Set objCheckBox = Nothing 
End Function
'-----------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_SISW_UI_JavaRadioButton_Operations

'Description		:	This function to perform operations on JavaRadioButton component.

'Parameters		    :	1. sFunctionName: Caller function's name
'						2. sAction		: Action to be performed
'						3. objJavaDialog : Prent UI Component or JavaRadioButton object
'						4. sRadioButtonName 	: JavaRadioButton Control name
'						5. sValue	: value

'Return Value		: 	TRUE \ FALSE

'Pre-requisite		:	objTwistieText must be in format TwistieControlName_Label present in Object repository

'Examples			:	Fn_SISW_UI_JavaRadioButton_Operations("Fn_SISW_RAC_PMM_UserContextSettings", "Set", objDialog, "Version", "ON")
'					:	Fn_SISW_UI_JavaRadioButton_Operations("Fn_SISW_RAC_PMM_UserContextSettings", "Set", objDialog.JavaRadioButton("Precise"), "" , "OFF")

'History			:
'-------------------------------+---------------+---------------+-------------------------------+-----------------------------------
'	Developer Name				|	Date		|	Rev. No.	|	Reviewer					|	Changes Done	
'-------------------------------+---------------+---------------+-------------------------------+-----------------------------------
'	Koustubh Watwe		 		|05-Nov-2012	|	1.0			|	Koustubh Watwe		 		| 	Created
'-------------------------------+---------------+---------------+-------------------------------+-----------------------------------
Public Function Fn_SISW_UI_JavaRadioButton_Operations(sFunctionName, sAction, objJavaDialog, sRadioButtonName, sValue)
	Dim sFuncLog, objRadioButton
	bGblFailedFunctionName = sFunctionName
	'Object Creation
	If sRadioButtonName <> "" Then
		Set objRadioButton = objJavaDialog.JavaRadioButton(sRadioButtonName)
		sFuncLog = sFunctionName + " > Fn_SISW_UI_JavaRadioButton_Operations  : [ " &  objJavaDialog.toString & " ] : [ " +  objRadioButton.toString + " ] : Action = " & sAction & " : "
	Else
		Set objRadioButton = objJavaDialog
		sFuncLog = sFunctionName + " > Fn_SISW_UI_JavaRadioButton_Operations  : [ " +  objRadioButton.toString + " ] : Action = " & sAction & " : "
	ENd IF
	
	'Verify JavaRadioButton object exists
	If Fn_SISW_UI_Object_Operations("Fn_SISW_UI_JavaRadioButton_Operations", "Exist", objRadioButton,"") = False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & " : FAIL : Does not exist.")
		Fn_SISW_UI_JavaRadioButton_Operations = False
		Call ExitFromUI(sFuncLog & " : FAIL : Object does not exist.")
		Set objRadioButton = Nothing 
		Exit Function
	End If
	
	Select case sAction
		Case "Set"
			objRadioButton.Set sValue
			Fn_SISW_UI_JavaRadioButton_Operations = True
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & "PASS : Sucessfully set JavaRadioButton [ " & sValue & " ].")
		Case "DeviceReplay.Set"
			If sJavaButton <> "" Then
				objJavaDialog.Click 1, 1,"LEFT"
				wait SISW_MICRO_TIMEOUT
			End If
			Set objDeviceReplay = CreateObject("Mercury.DeviceReplay")
			objDeviceReplay.MouseMove (objRadioButton.GetROProperty("abs_x") + 5), (objRadioButton.GetROProperty("abs_y") + 5)
			objDeviceReplay.MouseClick  (objRadioButton.GetROProperty("abs_x") + 5), (objRadioButton.GetROProperty("abs_y") + 5), 0
			Fn_SISW_UI_JavaRadioButton_Operations = True
		Case Else
	End Select
	'Clear memory of JavaRadioButton object.
	Set objRadioButton = Nothing 
End Function
'*******************************************************************************************************************
''Function Name		 	:	Fn_SISW_UI_JavaToolbar_Operations
'
''Description		    :  	This function is use to perform operations on javaToolbar object.

''Parameters		    :	1. sFunctionName : Caller function's Name
''							2. sAction	: Action to be performed
''							3. objContainer
''							4. sToolbar : ~ Toolbar's logical name from OR.
''							5. sToobarButtonName : ~ toolbar button Name
''							6. sValue : for future use
''							7. sMenu : popup / dropdown menu
''							8. iIndex : instance number.

''Return Value		    :  	true \ false
'
''Examples		     	:	Call Fn_SISW_UI_JavaToolbar_Operations("", "Click", JavaWindow("ProductMasterManager"), "", "Clear", "", "", 2)
'						:	Call Fn_SISW_UI_JavaToolbar_Operations("", "DropdownMenuSelect", JavaWindow("ProductMasterManager"), "", "Part Type", "", "Standard Part", "")
'						:	Call Fn_SISW_UI_JavaToolbar_Operations("", "DropdownMenuSelect", JavaWindow("ProductMasterManager"), "TabFolderWidgetToolBar", "Part Type", "", "Standard Part", "")
'						:	Call Fn_SISW_UI_JavaToolbar_Operations("", "OpenDropdownMenu", JavaWindow("ProductMasterManager"), "TabFolderWidgetToolBar", "Part Type", "", "", "")
'						:   Call Fn_SISW_UI_JavaToolbar_Operations("","isenabled",JavaWindow("BriefcaseBrowser"),"","Save Briefcase","","","")

'History:
'	Developer Name				Date			Rev. No.	Reviewer		Changes Done	
'*******************************************************************************************************************
'	Koustubh Watwe		        22-Oct-2012	    1.0			Self	
'-------------------------------------------------------------------------------------------------------------------
'	Shweta Rathod				25-Aug-2016		1.0        Koustubh watwe	added case "isenabled"	,"isselected" 
'*******************************************************************************************************************
Public Function Fn_SISW_UI_JavaToolbar_Operations(sFunctionName, sAction, objContainer, sToolbar, sToobarButtonName, sValue, sMenu, iIndex)
	Dim objJavaToolbar, objObjects, iCnt, sContents, bFound, sFuncLog
	Dim iX, iY
	bGblFailedFunctionName = sFunctionName
	sFuncLog = sFunctionName & " > Fn_SISW_UI_JavaToolbar_Operations : [ " & objContainer.toString() & " ] : Action = " & sAction & " : Toolbar Button = " & sToobarButtonName & " : "
	Fn_SISW_UI_JavaToolbar_Operations = False
	If sToolbar = "" Then
		If iIndex <> "" Then
			iIndex = cInt(iIndex)
		Else
			iIndex = 1
		End If
		bFound = False
		Set objObjects = Fn_SISW_UI_Object_GetChildObjects( "Fn_SISW_UI_JavaToolbar_Operations", objContainer, "Class Name~enabled", "JavaToolbar~1")
		If typename(objObjects ) <> "Nothing" Then
			For iCnt = 0 to objObjects.Count - 1
                sContents = objObjects(iCnt).GetContent()
				If  Fn_SISW_Setup_ArrayStringContains(sContents, sToobarButtonName, ";") Then
					If iIndex = 1 Then
						Set objJavaToolbar = objObjects(iCnt)
						bFound = True
						Exit for
					Else
						iIndex = iIndex - 1
					End If
				End If
			Next
		End If
		If bFound = False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & "FAIL: Failed to find toolbar [ " & sToobarButtonName & " ].")
			Exit function
		End If
	Else
		Set objJavaToolbar = objContainer.JavaToolbar(sToolbar)
	End If
	If Fn_SISW_UI_Object_Operations("Fn_SISW_UI_JavaToolbar_Operations", "Exist", objJavaToolbar,"") = False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & "FAIL: Failed to find toolbar button [ " & sValue & " ].")
		Exit Function
	End If
	Select Case sAction
	
			Case "ClickExt","ClickPackBtn"
			For iCounter=0 to objJavaToolbar.Object.getItemCount-1
				If objJavaToolbar.Object.getItem(iCounter).getToolTipText()=sToobarButtonName Then
					If sAction = "ClickPackBtn" Then
					      objJavaToolbar.Object.Pack(False) 
					      wait 1
				     End If
					objJavaToolbar.Object.getItem(iCounter).click(True)	
					Fn_SISW_UI_JavaToolbar_Operations = True
					Exit For
				ElseIF instr(objJavaToolbar.Object.getItem(iCounter).getToolTipText(), sToobarButtonName) > 0 Then
					objJavaToolbar.Object.getItem(iCounter).click(True)	
					Fn_SISW_UI_JavaToolbar_Operations = True
					Exit For
				End If
			Next
		Case "Exist"
			For iCounter=0 to objJavaToolbar.Object.getItemCount-1
				If objJavaToolbar.Object.getItem(iCounter).getToolTipText()=sToobarButtonName Then
					Fn_SISW_UI_JavaToolbar_Operations = True
					Exit For
				ElseIF instr(objJavaToolbar.Object.getItem(iCounter).getToolTipText(), sToobarButtonName) > 0 Then	
					Fn_SISW_UI_JavaToolbar_Operations = True
					Exit For
				End If
			Next	
		Case "Click"
			objJavaToolbar.Press sToobarButtonName
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & "PASS: Successfully clicked on [ " & sToobarButtonName & " ].")
			Fn_SISW_UI_JavaToolbar_Operations = True
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "DeviceReplay.Click"
			If sToolbar <> "" Then
				objContainer.Click 1, 1,"LEFT"
				wait SISW_MICRO_TIMEOUT
			End If
			Set objDeviceReplay = CreateObject("Mercury.DeviceReplay")
			Select Case objJavaToolbar.Object.getClass().toString()
				Case "class javax.swing.JToolBar"
					For iCnt = 0 to objJavaToolbar.Object.getComponentCount - 1
						If sToobarButtonName = objJavaToolbar.Object.getComponent(iCnt).getText() Then
							iX = cInt(objJavaToolbar.Object.getComponent(iCnt).getBounds().getX())
							iY = cInt(objJavaToolbar.Object.getComponent(iCnt).getBounds().getY())
							objDeviceReplay.MouseMove (cInt(objJavaToolbar.GetROProperty("abs_x")) + iX + 5 ), (cInt(objJavaToolbar.GetROProperty("abs_y")) + iY + 5 )
							wait SISW_MICRO_TIMEOUT
							objDeviceReplay.MouseClick (cInt(objJavaToolbar.GetROProperty("abs_x")) + iX + 5 ), (cInt(objJavaToolbar.GetROProperty("abs_y")) + iY + 5 ), 0
							Fn_SISW_UI_JavaToolbar_Operations = True
							exit for
						End If
					Next
			End Select
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "DropdownMenuSelect"
			wait SISW_MICRO_TIMEOUT
			objJavaToolbar.ShowDropdown sToobarButtonName
			wait SISW_MICRO_TIMEOUT
			Fn_SISW_UI_JavaToolbar_Operations = Fn_SISW_RAC_UI_Menu_Operations("Fn_SISW_UI_JavaToolbar_Operations",  "Select", objContainer, sMenu)
			If Fn_SISW_UI_JavaToolbar_Operations Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & "PASS: Successfully clicked on [ " & sToobarButtonName & " ] and [ " & sMenu & " ] selected.")
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & "FAIL: Failed to click on [ " & sToobarButtonName & " ] and select [ " & sMenu & " ].")
			End If				
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "OpenDropdownMenu"
			wait SISW_MICRO_TIMEOUT
			objJavaToolbar.ShowDropdown sToobarButtonName
			wait SISW_MICRO_TIMEOUT
			Fn_SISW_UI_JavaToolbar_Operations = True
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "isenabled"	'		
			For iCounter = 0 to objObjects.Count - 1
				sContents = objObjects(iCounter).GetContent()
				If instr(sContents, sToobarButtonName) > 0 Then
					If  "1" = objObjects(iCounter).GetItemProperty(sToobarButtonName, "enabled")  Then
			            Fn_SISW_UI_JavaToolbar_Operations = TRUE
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Object Java Button["+sButtonName+"] Is Enabled.")
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Object Java Button["+sButtonName+"] Is Not Enabled.")
					End if
					Exit For
				End If
			Next	
		Case "isselected"
			For iCounter = 0 to objObjects.Count - 1
				sContents = objObjects(iCounter).GetContent()
				If instr(sContents, sToobarButtonName) > 0 Then
					If  "1" = objObjects(iCounter).GetItemProperty (sToobarButtonName, "selected")  Then
                        Fn_SISW_UI_JavaToolbar_Operations = TRUE
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Object Java Button["+sButtonName+"] Is Selected.")
					Else
						Fn_SISW_UI_JavaToolbar_Operations = FALSE
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Object Java Button["+sButtonName+"] Is Not Selected.")
					End if
					Exit For
				End If
			Next				
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & "FAIL: invalid case.")		
	End Select
	If Fn_SISW_UI_JavaToolbar_Operations = True Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & "PASS: Executed successfully [ " & sToobarButtonName & " ].")				
	Else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & "FAIL: Execution Failed on [ " & sToobarButtonName & " ].")		
	End If
End Function
'*******************************************************************************************************************
''Function Name		 	:	Fn_SISW_UI_JavaTable_Operations
'
''Description		    :  	This function is used to perform operations on JavaTable object

''Parameters		    :	01. sFunctionName	: Cller function's name
''							02. sAction			: Action to be performed
''							03. objContainer 	: Prent UI Component or JavaTable object
''							04. sJavaTable 		: JavaTable Name / ""
''							05. sTableType 		: JavaTable Type ( eg. "", "GetCellData","","GetValueAt.getDisplayableValue", "Object.GetItem", "GetProperty" )
''							06. sPrimaryColumn	: Primary Column through which Row to be identiefied ( Column Names / index)
''							07. sRow			: Row text / Row number
''							08. sColumn 		: Column Name / Column Number
''							09. sCellData 		: Cell value.
''							10. sValue 			: Cell New value.
''							11. sPopupMenu 		: Popup Menu.
''							10. sInstanceHandler : Instance Handler .


''Return Value		    :  	For Case GetColumnIndex : -1 / Index of the Column
''Return Value		    :  	For Case RowSelect : True / False

''Examples
''							Call Fn_SISW_UI_JavaTable_Operations("CallerFunctionName", "Type", objApplet , "ListSubstitutes", "", "", "", "", "value", "", "")
''							Call Fn_SISW_UI_JavaTable_Operations("CallerFunctionName", "GetFirstRowData", objApplet , "ListSubstitutes", "", "", "", "", "", "", "")
''							Call Fn_SISW_UI_JavaTable_Operations("CallerFunctionName", "GetLastRowData", objApplet , "ListSubstitutes", "", "", "", "", "", "", "")
''							Call Fn_SISW_UI_JavaTable_Operations("CallerFunctionName", "GetRowCount", objApplet , "ListSubstitutes", "", "", "", "", "", "", "")
''							Call Fn_SISW_UI_JavaTable_Operations("CallerFunctionName", "GetColCount", objApplet , "ListSubstitutes", "", "", "", "", "", "", "")
''							Call Fn_SISW_UI_JavaTable_Operations("CallerFunctionName", "GetAllColumnNames", objApplet , "ListSubstitutes", "", "", "", "", "", "", "")
''							Call Fn_SISW_UI_JavaTable_Operations("CallerFunctionName", "GetColumnIndex", objReportTable , "", "", "BOM Line", "", "", "", "", "")
''							Call Fn_SISW_UI_JavaTable_Operations("CallerFunctionName", "GetRowIndex", objApplet , "ListSubstitutes", "", 0, sSubstituteNm, "", "", "", "")
''							Call Fn_SISW_UI_JavaTable_Operations("CallerFunctionName", "GetColumnContents", objReportTable , "", "", "BOM Line", "", "", "", "", "")
''							Call Fn_SISW_UI_JavaTable_Operations("CallerFunctionName", "Exist", objApplet , "ListSubstitutes", "", 0, sSubstituteNm, "", "", "", "")
''							Call Fn_SISW_UI_JavaTable_Operations("CallerFunctionName", "SelectRow", objApplet , "ListSubstitutes", "", 0, "Engine", "", "", "", "")
''							Call Fn_SISW_UI_JavaTable_Operations("CallerFunctionName", "SelectRow", objApplet , "ListSubstitutes", "", "Option", "Engine", "", "", "", "")
''							Call Fn_SISW_UI_JavaTable_Operations("CallerFunctionName", "Object.GetItem.ClickRow", objApplet , "ListSubstitutes", "", "Option", "Engine", "", "", "", "")
''	Example of Composite Column	-	Call Fn_SISW_UI_JavaTable_Operations("CallerFunctionName", "SelectRow", objApplet , "ListSubstitutes", "", "Item Id~Item Name", "000050~topAsm", "", "", "", "")
''							Call Fn_SISW_UI_JavaTable_Operations("CallerFunctionName", "DeselectRow", objReportTable , "", "", "Option", "Engine", "", "", "", "")
''							Call Fn_SISW_UI_JavaTable_Operations("CallerFunctionName", "ActivateRow", objReportTable , "", "", "Option", "Engine", "", "", "", "")
''							Call Fn_SISW_UI_JavaTable_Operations("CallerFunctionName", "ExtendRow", objReportTable , "", "", "Option", "Engine", "", "", "", "")
''							Call Fn_SISW_UI_JavaTable_Operations("CallerFunctionName", "ClickCell", objReportTable , "", "", "Option", "Engine", "Value", "", "", "")
''							Call Fn_SISW_UI_JavaTable_Operations("CallerFunctionName", "ActivateCell", objReportTable , "", "", "Option", "Engine", "Value", "", "", "")
''							Call Fn_SISW_UI_JavaTable_Operations("CallerFunctionName", "DoubleClickCell", objReportTable , "", "", "Option", "Engine", "Value", "", "", "")
''							Call Fn_SISW_UI_JavaTable_Operations("CallerFunctionName", "SetCellData", objReportTable , "", "", "Option", "Engine", "Value", "1600", "", "")
''							Call Fn_SISW_UI_JavaTable_Operations("CallerFunctionName", "SelectCell", objReportTable , "", "", "Option", "Engine", "Value", "", "", "")
''							Call Fn_SISW_UI_JavaTable_Operations("CallerFunctionName", "GetCellData", objReportTable , "", "", "Option", "Engine", "Value", "", "", "")
''							Call Fn_SISW_UI_JavaTable_Operations("CallerFunctionName", "GetCellData", objReportTable , "", "GetValueAt", "Option", "Engine", "Value", "", "", "")
''							Call Fn_SISW_UI_JavaTable_Operations("CallerFunctionName", "GetCellData", objReportTable , "", "GetValueAt.getDisplayableValue", "Option", "Engine", "Value", "", "", "")
''							Call Fn_SISW_UI_JavaTable_Operations("CallerFunctionName", "GetCellData", objReportTable , "", "Object.GetItem", "Option", "Engine", "Value", "", "", "")
''							Call Fn_SISW_UI_JavaTable_Operations("CallerFunctionName", "GetCellData", objReportTable , "", "GetProperty", "Option", "Engine", "Value", "", "", "")
''							Call Fn_SISW_UI_JavaTable_Operations("CallerFunctionName", "VerifyCellData", objReportTable , "", "", "Option", "Engine", "Value", "1900", "", "")
''							Call Fn_SISW_UI_JavaTable_Operations("CallerFunctionName", "VerifyCellData", objReportTable , "", "GetValueAt", "Option", "Engine", "Value", "1900", "", "")
''							Call Fn_SISW_UI_JavaTable_Operations("CallerFunctionName", "VerifyCellData", objReportTable , "", "GetValueAt.getDisplayableValue", "Option", "Engine", "Value", "1900", "", "")
''							Call Fn_SISW_UI_JavaTable_Operations("CallerFunctionName", "VerifyCellData", objReportTable , "", "Object.GetItem", "Option", "Engine", "Value", "1900", "", "")
''							Call Fn_SISW_UI_JavaTable_Operations("CallerFunctionName", "VerifyCellData", objReportTable , "", "GetProperty", "Option", "Engine", "Value", "1900", "", "")
''							Call Fn_SISW_UI_JavaTable_Operations("CallerFunctionName", "SelectColumnHeader", objReportTable , "", "", "BOM Line", "", "", "", "", "")
''							Call Fn_SISW_UI_JavaTable_Operations("CallerFunctionName", "RMB_MenuSelect_On_ColumnHeader", objContainer , sJavaTable, sTableType, sPrimaryColumn, sRow, sColumn, sValue, "", "")
''							Call Fn_SISW_UI_JavaTable_Operations("CallerFunctionName", "SelectRange", objContainer , sJavaTable, "Property~Type","", "Current Name~Folder$Description~Folder", "", "", "", "")
''							Call Fn_SISW_UI_JavaTable_Operations("CallerFunctionName", "SelectRowsRange", objContainer , sJavaTable, "", "", "1$5", "", "", "", "")
''							Call Fn_SISW_UI_JavaTable_Operations("CallerFunctionName", "SelectAllRows", objContainer , sJavaTable, "", "", "", "", "", "", "")
'History :
'-------------------------------+---------------+---------------+-------------------------------+-----------------------------------
'	Developer Name				|	Date		|	Rev. No.	|	Reviewer					|	Changes Done	
'-------------------------------+---------------+---------------+-------------------------------+-----------------------------------
'	Koustubh Watwe			 	|23-Aug-2013	|	1.0			|	Koustubh Watwe		 		| 	Created
'-------------------------------+---------------+---------------+-------------------------------+-----------------------------------
'	Ashwini Kumar			 	|01-Oct-2013	|	1.0			|	Koustubh Watwe		 		| 	Added cases SelectRange, SelectRowsRange, SelectAllRows
'-------------------------------+---------------+---------------+-------------------------------+-----------------------------------
'	Koustubh Watwe			 	|07-Jul-2014	|	1.0			|	Koustubh Watwe		 		| 	Added case "Object.GetItem.ClickRow"
'-------------------------------+---------------+---------------+-------------------------------+-----------------------------------
'	Vivek Ahirrao			 	|29-March-2016	|	1.1			|	Vivek Ahirrao		 		| 	Added case "Object.GetItem.ClickColumn"
'-------------------------------+---------------+---------------+-------------------------------+-----------------------------------
Public Function Fn_SISW_UI_JavaTable_Operations(sFunctionName, sAction, objContainer , sJavaTable, sTableType, sPrimaryColumn, sRow, sColumn, sValue, sPopupMenu, sInstanceHandler)
   Dim objTable, iCnt, iCols, iRowCount, sFuncLog, sData, iInstanceCounter, iRowIndex, iColCounter
	Dim iCol, aCols, iArrCnt, bFound, aRows, iArrRowCnt, aRowInstance, aRowData, var
	bGblFailedFunctionName = sFunctionName
	If sAction = "GetColumnIndex" OR sAction = "GetRowIndex" Then
		Fn_SISW_UI_JavaTable_Operations = -1
	Else
		Fn_SISW_UI_JavaTable_Operations = False
	End If
	If sInstanceHandler = "" Then sInstanceHandler = "@"
	If sJavaTable <> "" Then
		Set objTable = objContainer.JavaTable(sJavaTable)
		sFuncLog = sFunctionName & " > Fn_SISW_UI_JavaTable_Operations : [ " & objContainer.toString() & " ] : [ " & objTable.toString() & " ] : Action = " & sAction & " : "
	Else
		Set objTable = objContainer
		sFuncLog = sFunctionName & " > Fn_SISW_UI_JavaTable_Operations : [ " & objTable.toString() & " ] : Action = " & sAction & " : "
	End If
	
	If Fn_SISW_UI_Object_Operations("Fn_SISW_UI_JavaTable_Operations", "Exist", objTable, SISW_MICRO_TIMEOUT) = False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & "FAIL: Failed to find JavaTable.")
		Call ExitFromUI(sFuncLog)		
		Exit Function
	End If
	
	Select Case sAction
		Case "Type" ' Depricated
			objTable.Type sValue
			Fn_SISW_UI_JavaTable_Operations = True
			
		Case "GetFirstRowData" ' Depricated
			Fn_SISW_UI_JavaTable_Operations = objTable.GetCellData(0,0)
			
		Case "GetLastRowData" ' Depricated
			iRowCount = cInt(objTable.getROProperty("rows"))
			Fn_SISW_UI_JavaTable_Operations = objTable.GetCellData(CInt(iRowCount-1),0)
			
		Case "GetRowCount"
			Fn_SISW_UI_JavaTable_Operations = cInt(objTable.getROProperty("rows"))
			
		Case "GetColCount"
			Fn_SISW_UI_JavaTable_Operations = cInt(objTable.getROProperty("cols"))
				
		Case "GetAllColumnNames"
			iCols = cInt(objTable.getROProperty("cols"))
			For iCnt = 0 to iCols -1
				If iCnt = 0 then
					Fn_SISW_UI_JavaTable_Operations = ObjTable.GetColumnName(iCnt)
				Else
					Fn_SISW_UI_JavaTable_Operations = Fn_SISW_UI_JavaTable_Operations & "~" & ObjTable.GetColumnName(iCnt)
				End If
			Next
			
		Case "GetColumnIndex"
			' Multiple Columns
			If sPrimaryColumn = "" Then 
				Fn_SISW_UI_JavaTable_Operations = 0
			Else
				aCols = split(sPrimaryColumn,"~")
				iCols = cInt(objTable.getROProperty("cols"))
				Fn_SISW_UI_JavaTable_Operations = ""
				For iArrCnt = 0 to UBound(aCols)
					bFound = False
					For iCnt = 0 to iCols -1
						If lcase(typeName(aCols(iArrCnt))) = "integer" Then
							If Fn_SISW_UI_JavaTable_Operations = "" Then
								Fn_SISW_UI_JavaTable_Operations = aCols(iArrCnt)
							Else
								Fn_SISW_UI_JavaTable_Operations = Fn_SISW_UI_JavaTable_Operations & "~" & aCols(iArrCnt)
							End IF
							bFound = True
							Exit for
						Else
							If instr(aCols(iArrCnt),"#") > 0 Then
								If Fn_SISW_UI_JavaTable_Operations = "" Then
									Fn_SISW_UI_JavaTable_Operations = replace(aCols(iArrCnt), "#","")
								Else
									Fn_SISW_UI_JavaTable_Operations = Fn_SISW_UI_JavaTable_Operations & "~" & replace(aCols(iArrCnt), "#","")
								End IF
								bFound = True
								Exit for
							Else
								If Trim(Lcase(objTable.getColumnName(iCnt))) = Trim(Lcase(aCols(iArrCnt))) Then
									bFound = True
									If Fn_SISW_UI_JavaTable_Operations = "" Then
										Fn_SISW_UI_JavaTable_Operations = iCnt
									Else
										Fn_SISW_UI_JavaTable_Operations = Fn_SISW_UI_JavaTable_Operations & "~" & iCnt
									End IF
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),sFuncLog & "PASS: The Column [" & aCols(iArrCnt) & "] is present at [ "+ Cstr(iCnt) +" ] index.")
									Exit for
								End If
							End If
						End If
					Next
					If bFound = False Then
						Fn_SISW_UI_JavaTable_Operations = -1 
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),sFuncLog & "PASS: The Column [" & aCols(iArrCnt) & "] is not present.")
						Exit For
					End If
				Next
			End If
			If Fn_SISW_UI_JavaTable_Operations = -1 Then 
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & "FAIL: Column  [" & sPrimaryColumn & "] is not present.")
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "GetRowIndex"
			If sPrimaryColumn = "" Then
				sPrimaryColumn = 0
			End If
			If lcase(typeName(sPrimaryColumn)) = "integer" OR lcase(typeName(sPrimaryColumn)) = "double" OR instr(sPrimaryColumn,"#") > 0 Then
				iCol = sPrimaryColumn
			Else
				If sTableType = "GetProperty" Then
					iCol = sPrimaryColumn ' properties
				Else
					If isNumeric(sPrimaryColumn) Then
						iCol = cInt(sPrimaryColumn)
					Else
						iCol = Fn_SISW_UI_JavaTable_Operations(sFunctionName, "GetColumnIndex", objContainer , sJavaTable, sTableType, sPrimaryColumn, "", "", "", "","")
					End If
					If iCol = -1 Then
						' Column Not found
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & "Fail: Column  [" & sPrimaryColumn & "] is not present.")
						Call ExitFromUI(sFuncLog)
						Exit function
					End If
				End If
			End If
			aCols = split(iCol,"~")
			aRows = split(sRow, sInstanceHandler)
			iInstanceCounter = 1
			If uBound(aRows) > 0 Then
				iInstanceCounter = cInt(aRows(1))
			End If
			aRowData = split(trim(aRows(0)), "~")
			
'			iArrRowCnt
			iRowCount = cInt(objTable.getROProperty("rows"))
			If uBound(aCols) = 0 Then
				bFound = True
			Else
				bFound = False
			End If
			bFound = False
			For iCnt = 0 to iRowCount -1
				For iColCounter = 0 to uBound(aCols)
					Select Case sTableType
						Case "", "GetCellData"
							sData = trim(cstr(objTable.getCellData(iCnt, aCols(iColCounter))))
						Case "GetValueAt"
							sData = trim(cstr(objTable.Object.getValueAt(iCnt, aCols(iColCounter)).toString()))
						Case "GetValueAt.getDisplayableValue"
							sData = trim(cstr(objTable.Object.getValueAt(iCnt, aCols(iColCounter)).getDisplayableValue()))
						Case "Object.GetItem"
							sData = trim(cstr(objTable.Object.getItem(iCnt).getData().toString()))
						Case "GetProperty"
							If trim(aCols(iColCounter)) = "Relation" Then
								sData = objTable.Object.getItem(iCnt).getData().getContext().toString()
								Fn_SISW_UI_JavaTable_Operations = Fn_SISW_UI_GetDisplayedRelation(sData)
							Else
								On Error Resume Next
								sData = ""
								iCol = Fn_SISW_UI_GetRealPropertyName(aCols(iColCounter))
								sData = objTable.Object.getItem(iCnt).getData().getComponent().getProperty(iCol)

								If sData = "" Then
									sData = objTable.Object.getItem(iCnt).getData().getProperty(iCol)
								End If

								If sData = "" Then
									sData = objTable.Object.getItem(iCnt).getData().getStringProperty(iCol)
								End If
							End If
					End Select
					If isNumeric(sData) Then
						sData = cstr(cint(sData))
					End If
					If isNumeric(aRowData(iColCounter)) Then
						aRowData(iColCounter) = cstr(cint(aRowData(iColCounter)))
					End If

					If sData = aRowData(iColCounter) Then
						If iColCounter = uBound(aCols) Then
							If iInstanceCounter = 1 Then
								bFound = True
								exit for
							End If
							iInstanceCounter = iInstanceCounter - 1
						End If
					Else
						If iColCounter = 0 Then exit for
					End If
				Next
				If bFound = True Then exit for
			Next
			If bFound = True Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & "PASS : Successfully Clicked on Cell matching values [ " & replace(sRow,"~",", ") & " ] in columns [ " & replace(sPrimaryColumn,"~",", ") &" ]" )
				Fn_SISW_UI_JavaTable_Operations = iCnt
			End If
			If Fn_SISW_UI_JavaTable_Operations <> -1 Then 
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & "PASS : Successfully executed with JavaTable Type [ " & sTableType & " ]")
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & "FAIL : Failed to select row with JavaTable Type [ " & sTableType & " ]")
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "GetColumnContents"
			If lcase(typeName(sPrimaryColumn)) = "integer" OR lcase(typeName(sPrimaryColumn)) = "double" OR instr(sPrimaryColumn,"#") > 0 Then
				iCol = sPrimaryColumn
			Else
				If sTableType = "GetProperty" Then
					iCol = sPrimaryColumn ' properties
				Else
					If isNumeric(sPrimaryColumn) Then
						iCol = cInt(sPrimaryColumn)
					Else
						iCol = cInt(Fn_SISW_UI_JavaTable_Operations(sFunctionName, "GetColumnIndex", objContainer , sJavaTable, sTableType,  sPrimaryColumn, "", "", "", "", sInstanceHandler))
					End If
				End If
			End If
			If iCol <> -1 Then
				Fn_SISW_UI_JavaTable_Operations = ""
				iRowCount = cInt(objTable.getROProperty("rows"))
				For iCnt = 0 to iRowCount -1
					Select Case sTableType
						Case "", "GetCellData"
							sData = trim(cstr(objTable.getCellData(iCnt, iCol)))
						Case "GetValueAt"
							sData = trim(cstr(objTable.Object.getValueAt(iCnt, iCol).toString()))
						Case "GetValueAt.getDisplayableValue"
							sData = trim(cstr(objTable.Object.getValueAt(iCnt, iCol).getDisplayableValue()))
						Case "Object.GetItem"
							sData = trim(cstr(objTable.Object.getItem(iCnt).getData().toString()))
						Case "GetProperty"
							If trim(sPrimaryColumn) = "Relation" Then
								sData = objTable.Object.getItem(iCnt).getData().getContext().toString()
								Fn_SISW_UI_JavaTable_Operations = Fn_SISW_UI_GetDisplayedRelation(sData)
							Else
								On Error Resume Next
								sData = ""
								iCol = Fn_SISW_UI_GetRealPropertyName(sPrimaryColumn)
								sData = objTable.Object.getItem(iCnt).getData().getComponent().getProperty(iCol)

								If sData = "" Then
									sData = objTable.Object.getItem(iCnt).getData().getProperty(iCol)
								End If

								If sData = "" Then
									sData = objTable.Object.getItem(iCnt).getData().getStringProperty(iCol)
								End If
							End If
					End Select
					If iCnt = 0 Then
						Fn_SISW_UI_JavaTable_Operations = sData
					Else
						Fn_SISW_UI_JavaTable_Operations = Fn_SISW_UI_JavaTable_Operations & "~" & sData
					End If
				Next
				If Fn_SISW_UI_JavaTable_Operations = "" Then Fn_SISW_UI_JavaTable_Operations = False
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Exist","SelectRow", "DeselectRow", "ActivateRow", "ExtendRow", "Object.GetItem.ClickRow", "Object.GetItem.ClickColumn"
			If lcase(typeName(sRow)) = "integer" OR lcase(typeName(sRow)) = "double" OR instr(sRow,"#") > 0 Then
				iRowIndex = sRow
			Else
				iRowIndex = Fn_SISW_UI_JavaTable_Operations(sFunctionName, "GetRowIndex", objContainer , sJavaTable, sTableType, sPrimaryColumn, sRow, "", "", "", sInstanceHandler)
			End If
			If iRowIndex <> -1 Then
				Select Case sAction
					Case "Exist"
						' Do Nothing
					Case "ExtendRow"
						objTable.ExtendRow iRowIndex
					Case "SelectRow"
						objTable.SelectRow iRowIndex
					Case "DeselectRow"
						objTable.DeselectRow iRowIndex
					Case "ActivateRow"
						objTable.ActivateRow iRowIndex
					Case "Object.GetItem.ClickRow"
						Dim objRect, aBounds, sxLen, syLen, X, Y ,W, H, sBounds 
						Set objRect = objTable.Object.getItem(iRowIndex).getBounds()
						sBounds = objRect.toString()
						sBounds = mid(sBounds,instr(sBounds,"{")+1, len(sBounds) -instr(sBounds,"{")-1)
						aBounds = split(sBounds,",")
						X = cInt(trim(aBounds(0)))
						Y = cInt(trim(aBounds(1)))
						W = cInt(trim(aBounds(2)))
						H = cInt(trim(aBounds(3)))
						sxLen = X + 15
						syLen = Y + (H/2)
						objTable.click sxLen, syLen,"LEFT"
					'[TC1122-20160316-29_03_2016-VivekA-NewDevelopment] - Added case to click on Cell (Single click), as click cell was not working properly
					Case "Object.GetItem.ClickColumn"
						If IsNumeric(sPrimaryColumn) Then
							iCol = cInt(sPrimaryColumn)
						Else
							iCol = cInt(Fn_SISW_UI_JavaTable_Operations(sFunctionName, "GetColumnIndex", objContainer , sJavaTable, sTableType,  sPrimaryColumn, "", sColumn, "", "", sInstanceHandler))
						End If
						'get Total width of previous columns
						For iCount = 0 To iCol-1
							iTotalWidth = CInt(iTotalWidth) + CInt(objTable.Object.getColumn(iCount).getWidth)
						Next
						
						Set objRect = objTable.Object.getItem(iRowIndex).getBounds()
						sBounds = objRect.toString()
						sBounds = mid(sBounds,instr(sBounds,"{")+1, len(sBounds) -instr(sBounds,"{")-1)
						aBounds = split(sBounds,",")
						X = cInt(trim(aBounds(0)))
						Y = cInt(trim(aBounds(1)))
						W = cInt(trim(aBounds(2)))
						H = cInt(trim(aBounds(3)))
						sxLen = CInt(iTotalWidth)+ X + 15
						syLen = Y + (H/2)
						objTable.click sxLen, syLen,"LEFT"
				End Select
				Fn_SISW_UI_JavaTable_Operations = True
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "ClickCell", "ActivateCell", "DoubleClickCell", "SetCellData", "SelectCell"
			If lcase(typeName(sRow)) = "integer" OR lcase(typeName(sRow)) = "double"  OR instr(sRow,"#") > 0 Then
				iRowIndex = sRow
			Else
				iRowIndex = Fn_SISW_UI_JavaTable_Operations(sFunctionName, "GetRowIndex", objContainer , sJavaTable, sTableType, sPrimaryColumn, sRow, "", "", "", sInstanceHandler)
			End If

			If lcase(typeName(sColumn)) = "integer" OR lcase(typeName(sColumn)) = "double" OR instr(sColumn,"#") > 0 Then
				iCol = sColumn
			Else
				If isNumeric(sColumn) Then
					iCol = cInt(sColumn)
				Else
					iCol = cInt(Fn_SISW_UI_JavaTable_Operations(sFunctionName, "GetColumnIndex", objContainer , sJavaTable, sTableType,  sColumn, "", "", "", "", sInstanceHandler))
				End If
			End If
			If iRowIndex <> -1 AND iCol <> -1 Then
				Select Case sAction
					Case "ClickCell"
						objTable.ClickCell iRowIndex, iCol
					Case "ActivateCell"
						objTable.ActivateCell iRowIndex, iCol
					Case "DoubleClickCell"
						objTable.DoubleClickCell iRowIndex, iCol
					Case "SetCellData"
						objTable.SetCellData iRowIndex, iCol, sValue
					Case "SelectCell"
						objTable.SelectCell iRowIndex, iCol
				End Select
				Fn_SISW_UI_JavaTable_Operations = True
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "GetCellData"
			If sPrimaryColumn = "" Then
				sPrimaryColumn = 0
			End If
			If lcase(typeName(sRow)) = "integer" OR lcase(typeName(sRow)) = "double" OR instr(sRow,"#") > 0 Then
				iRowIndex = sRow
			Else
				If isNumeric(sRow) Then
					iRowIndex = cInt(sRow)
				Else
					iRowIndex = Fn_SISW_UI_JavaTable_Operations(sFunctionName, "GetRowIndex", objContainer , sJavaTable, sTableType, sPrimaryColumn, sRow, "", "", "", sInstanceHandler)
				End IF
			End If

			If sColumn = "" Then sColumn = 0
			If lcase(typeName(sColumn)) = "integer" OR lcase(typeName(sColumn)) = "double" OR instr(sColumn,"#") > 0 Then
				iCol = sColumn
			Else
				If isNumeric(sColumn) Then
					iCol = cInt(sColumn)
				Else
					If sTableType = "GetProperty" Then
						' assign any value other than -1
						iCol = 0
					Else
						iCol = cInt(Fn_SISW_UI_JavaTable_Operations(sFunctionName, "GetColumnIndex", objContainer , sJavaTable, sTableType,  sColumn, "", "", "", "", sInstanceHandler))
					End If
				End If
			End If

			If iRowIndex <> -1 AND iCol <> -1 Then
				Select Case sTableType
					Case "", "GetCellData"
						sData = trim(cstr(objTable.GetCellData(iRowIndex, iCol)))
					Case "GetValueAt"
						sData = trim(cstr(objTable.Object.getValueAt(iRowIndex, iCol).toString()))
					Case "GetValueAt.getDisplayableValue"
						sData = trim(cstr(objTable.Object.getValueAt(iRowIndex, iCol).getDisplayableValue()))
					Case "Object.GetItem"
						sData = trim(cstr(objTable.Object.getItem(iRowIndex).getData().toString()))
					Case "GetProperty"
						If trim(sColumn) = "Relation" Then
							sData = objTable.Object.getItem(iRowIndex).getData().getContext().toString()
							sData = Fn_SISW_UI_GetDisplayedRelation(sData)
						Else
							On Error Resume Next
							sData = ""
							iCol = Fn_SISW_UI_GetRealPropertyName(sColumn)
							sData = objTable.Object.getItem(iRowIndex).getData().getComponent().getProperty(iCol)

							If sData = "" Then
								sData = objTable.Object.getItem(iRowIndex).getData().getProperty(iCol)
							End If

							If sData = "" Then
								sData = objTable.Object.getItem(iRowIndex).getData().getStringProperty(iCol)
							End If
						End If
				End Select
				Fn_SISW_UI_JavaTable_Operations = sData
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "VerifyCellData"
			If sPrimaryColumn = "" Then
				sPrimaryColumn = 0
			End If
			If lcase(typeName(sRow)) = "integer" OR lcase(typeName(sRow)) = "double" OR instr(sRow,"#") > 0 Then
				iRowIndex = sRow
			Else
'				If isNumeric(sRow) Then
'					iRowIndex = cInt(sRow)
'				Else
					iRowIndex = Fn_SISW_UI_JavaTable_Operations(sFunctionName, "GetRowIndex", objContainer , sJavaTable, sTableType, sPrimaryColumn, sRow, "", "", "", sInstanceHandler)
'				End If
			end If

			If sColumn = "" Then sColumn = 0
			If lcase(typeName(sColumn)) = "integer" OR lcase(typeName(sColumn)) = "double" OR instr(sColumn,"#") > 0 Then
				iCol = sColumn
			Else
				If isNumeric(sColumn) Then
					iCol = cInt(sColumn)
				Else
					If sTableType = "GetProperty" Then
						' assign any value other than -1
						iCol = 0
					Else
						iCol = cInt(Fn_SISW_UI_JavaTable_Operations(sFunctionName, "GetColumnIndex", objContainer , sJavaTable, sTableType,  sColumn, "", "", "", "", sInstanceHandler))
					End If
				End If
			End If

			If iRowIndex <> -1 AND iCol <> -1 Then
				Select Case sTableType
					Case "", "GetCellData"
						sData = trim(cstr(objTable.GetCellData(iRowIndex, iCol)))
					Case "Object.GetItem"
						sData = trim(cstr(objTable.Object.getItem(iRowIndex).getData().toString()))
					Case "GetProperty"
						If trim(sColumn) = "Relation" Then
							sData = objTable.Object.getItem(iRowIndex).getData().getContext().toString()
							sData = Fn_SISW_UI_GetDisplayedRelation(sData)
						Else
							On Error Resume Next
							sData = ""
							iCol = Fn_SISW_UI_GetRealPropertyName(sColumn)
							sData = objTable.Object.getItem(iRowIndex).getData().getComponent().getProperty(iCol)

							If sData = "" Then
								sData = objTable.Object.getItem(iRowIndex).getData().getProperty(iCol)
							End If

							If sData = "" Then
								sData = objTable.Object.getItem(iRowIndex).getData().getStringProperty(iCol)
							End If
						End If
				End Select
				Fn_SISW_UI_JavaTable_Operations = (sData = sValue)
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "SelectColumnHeader"
			If lcase(typeName(sPrimaryColumn)) = "integer" OR lcase(typeName(sPrimaryColumn)) = "double" OR instr(sPrimaryColumn,"#") > 0 Then
				iCol = sPrimaryColumn
			Else
				If sTableType = "GetProperty" Then
					iCol = sPrimaryColumn ' properties
				Else
					If isNumeric(sPrimaryColumn) Then
						iCol = cInt(sPrimaryColumn)
					Else
						iCol = cInt(Fn_SISW_UI_JavaTable_Operations(sFunctionName, "GetColumnIndex", objContainer , sJavaTable, sTableType,  sPrimaryColumn, "", "", "", "", ""))
					End If
				End If
			End If
			If iCol <> -1 Then
				objTable.SelectColumnHeader iCol, "LEFT"
				Fn_SISW_UI_JavaTable_Operations = True
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "RMB_MenuSelect_On_ColumnHeader"
			If lcase(typeName(sPrimaryColumn)) = "integer" OR lcase(typeName(sPrimaryColumn)) = "double" OR instr(sPrimaryColumn,"#") > 0 Then
				iCol = sPrimaryColumn
			Else
				If sTableType = "GetProperty" Then
					iCol = sPrimaryColumn ' properties
				Else
					If isNumeric(sPrimaryColumn) Then
						iCol = cInt(sPrimaryColumn)
					Else
						iCol = cInt(Fn_SISW_UI_JavaTable_Operations(sFunctionName, "GetColumnIndex", objContainer , sJavaTable, sTableType,  sPrimaryColumn, "", "", "", "", ""))
					End If
				End If
			End If
			If iCol <> -1 Then
				objTable.SelectColumnHeader iCol, "RIGHT"
				wait SISW_MICRO_TIMEOUT
				set var = Description.Create()
				var("Class Name").value = "JavaMenu"
				set childObjects = objTable.ChildObjects(var)
				If childObjects.count <> 0 then
					iArrCnt = split(sPopupMenu, ":",-1,1)
					Select Case Ubound(iArrCnt)
						Case 0
							objTable.JavaMenu("label:="&iArrCnt(0),"index:=0").Select
							bFound = True
						Case 1
							objTable.JavaMenu("label:="&iArrCnt(0),"index:=0").JavaMenu("label:="&iArrCnt(1),"index:=1").Select
							bFound = True
						Case 2
							objTable.JavaMenu("label:="&iArrCnt(0),"index:=0").JavaMenu("label:="&iArrCnt(1),"index:=1").JavaMenu("label:="&iArrCnt(2),"index:=2").Select
							bFound = True
						Case Else
							bFound = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_SISW_UI_JavaTable_Operations Failed to select Menu " & sPopupMenu)						
					End Select								
				End If
				Fn_SISW_UI_JavaTable_Operations = bFound
			End If				
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "SelectRowsRange"
			aRows = split(sRow, "$")
			If CInt(aRows(0))> -1 AND CInt(aRows(1))>-1 Then
				objTable.SelectRowsRange CInt(aRows(0)),CInt(aRows(1))
				Fn_SISW_UI_JavaTable_Operations = True
			 Else
				Fn_SISW_UI_JavaTable_Operations = False
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "SelectRange"
			aRows = split(sRow, "$")
			aRows(0)= Fn_SISW_UI_JavaTable_Operations(sFunctionName, "GetRowIndex", objTable, "", "",sPrimaryColumn, aRows(0), "", "", "", "")
			aRows(1)= Fn_SISW_UI_JavaTable_Operations(sFunctionName, "GetRowIndex", objTable, "", "",sPrimaryColumn, aRows(1), "", "", "", "")		
			Fn_SISW_UI_JavaTable_Operations= Fn_SISW_UI_JavaTable_Operations(sFunctionName, "SelectRowsRange",objTable, "", "", "", CStr(aRows(0))&"$"&CStr(aRows(1)), "", "", "", "")
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "SelectAllRows"
			Fn_SISW_UI_JavaTable_Operations= Fn_SISW_UI_JavaTable_Operations(sFunctionName, "SelectRowsRange",objTable, "", "", "", "0$"&(objTable.GetROProperty("rows") - 1), "", "", "", "")
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & "FAIL: Invalid Case.")		
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	End Select
	
	If sAction = "GetColumnIndex" OR sAction = "GetRowIndex" Then
		If Fn_SISW_UI_JavaTable_Operations <> -1 Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & "PASS: Executed successfully.")				
		Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & "FAIL: Execution Failed.")		
		End If
	Else
		If Fn_SISW_UI_JavaTable_Operations <> False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & "PASS: Executed successfully.")				
		Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & "FAIL: Execution Failed.")		
		End If
	End If
	
End Function

''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''''/$$$$
''''/$$$$   FUNCTION NAME   :  Fn_SISW_UI_DeviceReplayObjectClick
''''/$$$$
''''/$$$$   DESCRIPTION     : To click on any Objects using device replay
''''/$$$$ 
''''/$$$$
''''/$$$$   PARAMETERS      :   sFunctionNamee - Valid function name
''''/$$$$                   	 						objObject - Object hierachy
''''/$$$$	
''''/$$$$	Return Value 			: 		Nothing
''''/$$$$
''''/$$$$    Function Calls       :  		 NA
''''/$$$$										
''''/$$$$
''''/$$$$		HISTORY           :   		AUTHOR                 		DATE        	  VERSION
''''/$$$$
''''/$$$$    CREATED BY     :  		Pranav Ingle			    02-Dec-2013  	       		1.0
''''/$$$$
''''/$$$$    MODIFIED BY     :  	Koustubh W			    	17-Jul-2014  	       		1.0				Added code to add half height and width in both coordinates
''''/$$$$
''''/$$$$		How To Use :    Call Fn_SISW_UI_DeviceReplayObjectClick("Test" ,WpfWindow("Save"))
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Function Fn_SISW_UI_DeviceReplayObjectClick(sFunctionName, objObject)
	Fn_SISW_UI_DeviceReplayObjectClick = False
	Dim objDeviceRply,objSwf,xCord,yCord, height, width
	bGblFailedFunctionName = sFunctionName
	Set objSwf=objObject
	Set objDeviceRply=CreateObject("Mercury.DeviceReplay")
	xCord=objSwf.GetRoProperty("abs_x")
	yCord=objSwf.GetRoProperty("abs_y")
	height=objSwf.GetRoProperty("height")
	width=objSwf.GetRoProperty("width")
	objDeviceRply.MouseClick (xCord + width/2 ),(yCord + height/2),LEFT_MOUSE_BUTTON
	If Err.Number = 0 then
		Fn_SISW_UI_DeviceReplayObjectClick = True
	End if 
	Set objDeviceRply=Nothing
	Set objSwf=Nothing
End Function



'*******************************************************************************************************************
'Function Name		 	:	Fn_SISW_UI_InsightObject_Operations
'
'Description		    :  	This function is used to perform operations on JavaButton object.

'Parameters			    :	1. sFunctionName : Valid Function name, 
'							2. sAction: Action to be performed
'							3. objJavaDialog : Hiearchy of Parent object
'							4. sInsightObject   : Valid sInsight Object name

'Return Value		    :  	True \ False
'
'Examples		     	:	Call Fn_SISW_UI_InsightObject_Operations("Fn_TeamcenterLogin", "Click", JavaWindow("TcVizMainWin"),"Snapshot2")
'Examples		     	:	Call Fn_SISW_UI_InsightObject_Operations("Fn_TeamcenterLogin", "DoubleClick", JavaWindow("TcVizMainWin"),"Snapshot2Image")

'History:
'	Developer Name				Date			Rev. No.	Reviewer		Changes Done	
'*******************************************************************************************************************
'	swapna G		        04-Nov-2014	    1.0			Self	
'*******************************************************************************************************************
Public Function Fn_SISW_UI_InsightObject_Operations(sFunctionName, sAction, objJavaDialog, sInsightObject)
	Dim objInsightObject, sFuncLog
	bGblFailedFunctionName = sFunctionName
	Fn_SISW_UI_InsightObject_Operations = False
	'Object Creation
	If sInsightObject <> "" Then
		Set objInsightObject = objJavaDialog.InsightObject(sInsightObject)
		sFuncLog = sFunctionName + " > Fn_SISW_UI_InsightObject_Operations  : [ " &  objJavaDialog.toString & " ] : [ " +  objInsightObject.toString + " ] : Action = " & sAction & " : "
	Else
		Set objInsightObject = objJavaDialog
		sFuncLog = sFunctionName + " > Fn_SISW_UI_InsightObject_Operations  : [ " +  objInsightObject.toString + " ] : Action = " & sAction & " : "
	ENd IF
	
	'Verify InsightObject object exists
	If Fn_SISW_UI_Object_Operations("Fn_SISW_UI_InsightObject_Operations", "Exist", objInsightObject,"") = False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & " : FAIL : Does not exist.")
		'Call ExitFromUI(sFuncLog & " : FAIL : Does not exist.")
		Fn_SISW_UI_InsightObject_Operations = False
		Set objInsightObject = Nothing 
		Exit Function
	End If
	
	Select Case sAction
		Case "Click"
			objInsightObject.Click	
			If Err.Number < 0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_UI_ClickJavaTreeCell : FAIL : Failed to click on InsightObject")
				Exit Function
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & "PASS : Sucessfully clicked on InsightObject.")
				Fn_SISW_UI_InsightObject_Operations = True
			End If
		Case "DoubleClick"
			objInsightObject.DblClick	
			If Err.Number < 0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_UI_ClickJavaTreeCell : FAIL : Failed to double click on InsightObject")
				Exit Function
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & "PASS : Sucessfully double clicked on InsightObject.")
				Fn_SISW_UI_InsightObject_Operations = True
			End if
		Case Else
	End Select
	'Clear memory of InsightObject object.
	Set objInsightObject = Nothing 
End Function

'*******************************************************************************************************************
'Function Name		 	:	Fn_SISW_UI_WinListView_Operations
'
'Description		    :  	This function is used to perform operations on WinListView object.

'Parameters			    :	1. sFunctionName 	: Valid Function name, 
'							2. sAction			: Action to be performed
'							3. objJavaDialog	: Parent dialog / WinListView Object
'							4. sWinListView   	: WinListView Name
'							5. sValue   		: Value to be selected / verified
'							6. sColumn   		: --

'Return Value		    :  	True \ False\ Index
'
'Examples		     	:	Call Fn_SISW_UI_WinListView_Operations("Fn_DocumentSearchOrder", "Select", Window("LifeViewWin").Dialog("AutoFileSearchPreferences") ,"DocSearchOrderList","Examples", "Directory Set")
'							Call Fn_SISW_UI_WinListView_Operations("Fn_DocumentSearchOrder", "GetIndex", Window("LifeViewWin").Dialog("AutoFileSearchPreferences") ,"DocSearchOrderList","Examples", "Directory Set")

'History:
'	Developer Name				Date			Rev. No.	Reviewer		Changes Done	
'*******************************************************************************************************************
'	Reema Wadhwa	        12-Nov-2014	    1.0			Paresh 	
'*******************************************************************************************************************
Public Function Fn_SISW_UI_WinListView_Operations(sFunctionName, sAction, objJavaDialog, sWinListView, sValue, sColumn)
   Dim objWinListView, sFuncLog, iCounter, bFlag,iColCnt
   	bGblFailedFunctionName = sFunctionName
	bFlag = False
	Fn_SISW_UI_WinListView_Operations = False
	'Object Creation
	If sWinListView <> "" Then
		Set objWinListView = objJavaDialog.WinListView(sWinListView)
		sFuncLog = sFunctionName + " > Fn_SISW_UI_WinListView_Operations  : [ " &  objJavaDialog.toString & " ] : [ " +  objWinListView.toString + " ] : Action = " & sAction & " : "
	Else
		Set objWinListView = objJavaDialog
		sFuncLog = sFunctionName + " > Fn_SISW_UI_WinListView_Operations  : [ " +  objWinListView.toString + " ] : Action = " & sAction & " : "
	ENd IF
	
	'Verify ListView object exists
	If Fn_SISW_UI_Object_Operations("Fn_SISW_UI_WinListView_Operations", "Exist", objWinListView,"") = False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & " : FAIL : WinListView does not exist.")
		Call ExitFromUI(sFuncLog & " : FAIL : WinListView does not exist.")
		Set objWinListView = Nothing 
		Exit Function
	End If
	
	Select Case sAction
		'------------------------------------------
		Case "Select"

			For iColCnt = 0 To objJavaDialog.WinListView(sWinListView).ColumnCount-1 Step 1
			 If Trim(objJavaDialog.WinListView(sWinListView).GetColumnHeader(iColCnt)) = Trim(sColumn)  Then
			 	' Select the element from list  
				For iCounter=0 to objJavaDialog.WinListView(sWinListView).GetItemsCount()-1
				    If Trim(objJavaDialog.WinListView(sWinListView).GetSubItem(iCounter, iColCnt))=Trim(sValue) Then
		                objJavaDialog.WinListView(sWinListView).Select iCounter
		                Call Fn_WriteLogFile(Environment.Value("TestLogFile"),sFuncLog & "Successfully Selected [" + sValue + "]")
		                bFlag=True
		                Exit For
				    End If
				Next
			 End If	
			 If bFlag = True Then
			 	Exit For
			 End If
			Next
		'------------------------------------------
		Case "GetRowIndex"

			iColCnt = objJavaDialog.WinListView(sWinListView).ColumnCount
			For iCount = 0 To iColCnt-1 Step 1
			 If Trim(objJavaDialog.WinListView(sWinListView).GetColumnHeader(iCount)) = Trim(sColumn)  Then
			 	' Select the element from list  
				For iCounter=0 to objJavaDialog.WinListView(sWinListView).GetItemsCount()-1
				    If Trim(objJavaDialog.WinListView(sWinListView).GetSubItem(iCounter, iCount))=Trim(sValue) Then
				    	Call Fn_WriteLogFile(Environment.Value("TestLogFile"),sFuncLog & "Successfully Get Index [" + cstr(iCounter) + "] for Value [" + sValue + "]")
		                bFlag = iCounter
		                Exit For
				    End If
				Next
			 End If	
			  If bFlag <> False Then
			 	Exit For
			 End If
			Next
			
	End Select
	
	Fn_SISW_UI_WinListView_Operations = bFlag 
	
	If Fn_SISW_UI_WinListView_Operations <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),sFuncLog & "Executed successfully")
	End If
End Function

'#######################################################################################################################
'###    FUNCTION NAME   :	Fn_UI_SetDateAndTimeExt(sFunctionName,sDate,sTime,objDate,objTime)
'###
'###    DESCRIPTION     :   This function is use to Set Date in edit box and Time in list
'###
'###    PARAMETERS      :   1. sFunctionName	:	Valid Function Name
'###		 	   			2. sDate			:	Valid Date
'###			   			3. sTime			:	Valid Time
'###			   			4. objDate			:	Edit box object for Date
'###			   			5. objTime			:	List box object for Time
'###			    
'###    RETURNVALUE   	:   True/False
'###
'###    PRE-REQUISITES  :  	Date Edit box and Time List box should be displayed and enabled
'###
'###    HISTORY         :	AUTHOR		|	   DATE		|  VERSION
'###
'###    CREATED BY      :	Vivek A 	|	03/05/2010	| 	1.0
'###
'########################################################################################################################
Public Function Fn_UI_SetDateAndTimeExt(sFunctionName,sDate,sTime,objDate,objTime)
   	Dim WshShell, sUIFail
   	bGblFailedFunctionName = sFunctionName
	Fn_UI_SetDateAndTimeExt = False
	sUIFail = sFunctionName + " >> Fn_UI_SetDateAndTimeExt >> "+sDate+" >> "+sTime

	If sDate<>"" Then
		objDate.Set sDate
		Wait 1
		Set WshShell = CreateObject("WScript.Shell")
		WshShell.SendKeys "{ESC}"
		Set WshShell = Nothing
	End If
	Call Fn_SyncTCObjects()
	If sTime<>"" Then
		objTime.Select ""
		Wait 1
		objTime.Select sTime
	End If
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Date and Time selected as "+sDate+" "+sTime)
	Fn_UI_SetDateAndTimeExt = True
End Function

'#######################################################################################################################
'###    FUNCTION NAME   :	Fn_UI_ResizeObject(sFunctionName, sObject, sLength, sBreadth)
'### 
'###    DESCRIPTION     :   This function is use to Set Object Length and Breath
'###
'###    PARAMETERS      :   1. sFunctionName	:	Valid Function Name
'###		 	   			2. objObject		:	Valid Object
'###			   			3. sLength			:	Valid Length
'###			   			4. sBreadth			:	Valid Breadth
'###			    
'###    RETURNVALUE   	:   True/False
'###
'###    PRE-REQUISITES  :   Object should be opened
'###
'###    HISTORY         :	AUTHOR			|	   DATE		|  VERSION
'###
'###    CREATED BY      :	Ankit Nigam 	|	08/02/2016	| 	1.0
'########################################################################################################################
Public Function Fn_UI_ResizeObject(sFunctionName, objObject, sLength, sBreadth)
	bGblFailedFunctionName = sFunctionName
   	Fn_UI_ResizeObject = False
   	
   	If Fn_SISW_UI_Object_Operations("Fn_UI_ResizeObject","Exist", objObject,"") = False Then
   		Fn_UI_ResizeObject = False
   		Exit Function
   	Else
		objObject.Object.setSize sLength, sBreadth   	
		Fn_UI_ResizeObject = True   	
   	End If
  	
   	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Object size has been set to Length [" & sLength & "] and Breadth [" & sBreadth)
End Function

'####################################### Function to Verify Existance of Vertical and Horizontal Bar ####################################
'###    FUNCTION NAME   :	Fn_UI_VerifyHorizontalVerticalBar(sFunctionName, sBar, objJavaDialog)
'###
'###    DESCRIPTION     :   This function is use to Verify Existance of Vertical and Horizontal Bar in given Dialog
'###
'###    PARAMETERS      :   1. sFunctionName	:	Valid Function Name
'###		 	   			2. sBar				:	Valid Bar Name (i.e. Horizontal/Vertical)
'###			   			3. objJavaDialog	:	Valid Object Dialog (i.e Tree/Table)
'###			    
'###    RETURNVALUE   	:   true/false
'###
'###    PRE-REQUISITES  :  	Dialog should be enabled
'###
'###    HISTORY         :	  AUTHOR		|	   DATE		|  VERSION
'###
'###    CREATED BY      :	Ankit Nigam 	|	02/02/2016	| 	 1.0
'###
'#########################################################################################################################################
Function Fn_UI_VerifyHorizontalVerticalBar(sFunctionName, sBar, objJavaDialog)
	Fn_UI_VerifyHorizontalVerticalBar = False
	bGblFailedFunctionName = sFunctionName
	Select Case sBar
		Case "HorizontalBar"
				Fn_UI_VerifyHorizontalVerticalBar = objJavaDialog.Object.GetHorizontalBar().isVisible()	
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Verified Existance of HorizontalBar under [" & objJavaDialog.toString & "]")
		Case "VerticalBar"
				Fn_UI_VerifyHorizontalVerticalBar = objJavaDialog.Object.GetVerticalBar().isVisible()
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Verified Existance of VerticalBar under [" & objJavaDialog.toString & "]")
	End Select
End Function

'*******************************************************************************************************************
'Function Name		 	:	Fn_SISW_UI_WinButton_Operations
'
'Description		    :  	This function is used to perform operations on WinButton object.

'Parameters			    :	1. sFunctionName : Valid Function name, 
'							2. sAction
'							3. objWinDialog
'							4. sWinButton   : Valid Button name
'							5. iXValue : x-coordinate
'							6. iYValue : Y- coordinate	
'							7. coMicButton : button value (micLeftBtn,micRightBtn)

'Return Value		    :  	True \ False
'
'Examples		     	:	Call Fn_SISW_UI_WinButton_Operations("Fn_TeamcenterLogin", "Click", JavaWindow("Teamcenter Login"),"Login","","","")
'Examples		     	:	call Fn_SISW_UI_WinButton_Operations("Fn_RunRCAFScript",JavaWindow("RCAF Console").Dialog("Open"),"Open",5,5,micLeftBtn)

'History:
'	Developer Name				Date			Rev. No.	Reviewer			Changes Done	
'*******************************************************************************************************************
'	Shweta Rathod		        31-Aug-2016	    1.0			Koustubh watwe
'*******************************************************************************************************************
Public Function Fn_SISW_UI_WinButton_Operations(sFunctionName, sAction, objWinDialog, sWinButton,iXValue,iYValue,coMicButton)
	Dim objWinButton, sFuncLog
	bGblFailedFunctionName = sFunctionName
	Fn_SISW_UI_WinButton_Operations = False
	'Object Creation
	If sWinButton <> "" Then
		Set objWinButton = objWinDialog.WinButton(sWinButton)
		sFuncLog = sFunctionName + " > Fn_SISW_UI_WinButton_Operations  : [ " &  objWinDialog.toString & " ] : [ " +  objWinButton.toString + " ] : Action = " & sAction & " : "
	Else
		Set objWinButton = objWinDialog
		sFuncLog = sFunctionName + " > Fn_SISW_UI_WinButton_Operations  : [ " +  objWinButton.toString + " ] : Action = " & sAction & " : "
	End If
	
	'Verify WinButton object exists
	If Fn_SISW_UI_Object_Operations("Fn_SISW_UI_WinButton_Operations", "Exist", objWinButton,SISW_MICRO_TIMEOUT) = False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & " : FAIL : Does not exist.")
		Exit Function
	End If
	
	'Verify WinButton object enabled
	If objWinButton.GetROProperty("enabled") = False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & "FAIL : [ "+sWinButton+" ] WinButton  is disabled of Function " &sFunctionName)
		Exit Function
	end if
	
	Select Case lcase(sAction)
		Case "click", "fn_ui_winbutton_click" ' Depricated case Fn_UI_WinButton_Click special case for function Fn_UI_WinButton_Click
			If  iXValue <> "" AND iYValue <> "" Then
				objWinButton.Click iXValue,iYValue,coMicButton
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),  "Sucessfully clicked on WinButton" & sWinButton &" at Co-ordinates " &  iXValue &"," &iYValue & " of Function " & sFunctionName)
			else
				objWinButton.Click
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Clicked on " & sWinButton & "WinButton of Function " & sFunctionName)
			end if			
	End Select
	'Clear memory of WinButton object.
	Set objWinButton = Nothing 
	Fn_SISW_UI_WinButton_Operations = True
End Function

'*******************************************************************************************************************
'Function Name		 	:	Fn_SISW_UI_WinEdit_Operations
'
'Description		    :  	This function is use to perform opeartions on WinEdit object.

'Parameters			    :	1. sFunctionName : Valid Function name, 
'							2. sAction
'							3. objWinDialog
'							4. sWinEdit
'							5. sText

'Return Value		    :  	True \ false
'
'Examples		     	:	Call Fn_SISW_UI_WinEdit_Operations("Fn_SISW_PMM_UserContextSettings", "Set",  objDialog, "EWO", "EWO_Name" )
'Examples		     	:	Call Fn_SISW_UI_WinEdit_Operations("Fn_SISW_PMM_UserContextSettings", "Type",  objDialog.WinEdit("EWO"),"", "EWO_Name" )

'History:
'	Developer Name				Date			Rev. No.	Reviewer		Changes Done	
'*******************************************************************************************************************
'	Shweta Rathod		        31-Aug-2016	    1.0			Koustubh Watwe	
'*******************************************************************************************************************
Function Fn_SISW_UI_WinEdit_Operations(sFunctionName, sAction, objWinDialog, sWinEdit, sText)
	Dim objWinEdit, sFuncLog
	bGblFailedFunctionName = sFunctionName
	Fn_SISW_UI_WinEdit_Operations = False
	'Set an Edit Object on variable
	If sWinEdit <> "" Then
		Set objWinEdit= objWinDialog.WinEdit(sWinEdit)
		sFuncLog = sFunctionName + "> Fn_SISW_UI_WinEdit_Operations : [ " &  objWinDialog.toString & " ] : [ " & objWinEdit.toString & " ] : Action = " & sAction & " : "
	Else
		Set objWinEdit= objWinDialog
		sFuncLog = sFunctionName + "> Fn_SISW_UI_WinEdit_Operations : [ " & objWinEdit.toString & " ] : Action = " & sAction & " : "
	End If

	If Fn_SISW_UI_Object_Operations("Fn_SISW_UI_WinEdit_Operations", "Exist", objWinEdit,"") = False Then
		Set objWinEdit = Nothing 
		Exit Function
	End If
	Select Case lcase(sAction)
		Case "set"
			'Setting the editbox
			objWinEdit.Set sText
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & "PASS : Text " &sText  &"is Set/ Entered in WinEditBox.")
			Fn_SISW_UI_WinEdit_Operations= True

		Case "type"
			'Setting the editbox
			objWinEdit.Type sText
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sFuncLog & "PASS : Text " &sText  &" is Set/ Entered in WinEditBox.")
			Fn_SISW_UI_WinEdit_Operations= True
		Case "gettext"
			Fn_SISW_UI_WinEdit_Operations = objWinEdit.getROProperty("value")
	End Select
	Set objWinEdit = Nothing 
End Function
