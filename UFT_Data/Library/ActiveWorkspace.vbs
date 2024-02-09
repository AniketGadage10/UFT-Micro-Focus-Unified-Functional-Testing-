Option Explicit

' Function List
'================================================================================================================================================================
'0. Fn_AWSummary_TabButton_Operation()                         This function is used to check existance of Button in summary tab. 
'1. Fn_AWSummary_HederTabText_Exist()                          This function is used to check existance of Text in Heder tab of AW summary 
'2. Fn_AWSummary_Popupwindow_Button_Operation()                This function is used perform button operation in AW sumary tab
'3. Fn_AW_BOM_TreeOperation()								   Function Used to perform operation on BOMtable Node from Active workspace tab
'4. Fn_AW_Tab_Operation()									   Function Used to select tab in Active workspace 
'5. Fn_AW_PropertiesTab_Operation()							   Function Used to verify text in properties tab of in Active workspace tab
'6. Fn_AW_TabButton_Operation()								   Function Used to button operation in AW tab
'7. Fn_AW_HederTabText_Exist()                                 This function is used to check existance of Text in Heder tab
'8. Fn_AW_ProfilePanel_Operation()                             Function Used to Perform operation on Profile panel in Active workspace tab
'9. Fn_AWSummary_DatasetDownload_Operation()                   This function is used to Perform dataset download operation in summary tab
'10. Fn_AW_RefTable_Operation()                      		   This Function Used to perform operation on All Reference table
'11. Fn_AW_AddPreference_Operation()						   This Function Used to add Reference
'12. Fn_AW_AddTarget_Operation()							   This Function Used to add Target
'13. Fn_AW_TargetTable_Operation()                             This Function Used to perform operation on All Target table
'14. Fn_AW_WorklistTaskOperation()                             This Function Used to perform operation on Worklist Task
'15. Fn_AW_Popupwindow_Button_Operation()                      This function Used to perform button operation in AW  tab
'16. Fn_AW_TaskOperation()                                     This function Used to perform task operations in AW tab
' =========================================================End Of List ==========================================================================================


'#######################################################################################################
'###    FUNCTION NAME   :   Fn_AWSummary_TabButton_Operation()
'###
'###    DESCRIPTION     :   This function is used to Perform button operation in summary tab
'###
'###    PARAMETERS      :   sFunctionName : Valid Function name, 
'###                       						 ObjTab : Valid Dialog box object name,
'###                        					sAwsButton   : Valid Button name
'###
'###    HISTORY         :   AUTHOR         DATE          
'###
'###    CREATED BY      :   Amruta         16/07/2021    
'###    MODIFIED BY     :   
'###    EXAMPLE         :   
'#######################################################################################################

Function Fn_AWSummary_TabButton_Operation(sAction,ObjTab, sAwsButton)

	Dim objAwsButton, bReturn,sbuttonArray,iCount,tabobj,tabsummary,objDesc
	'Object Creation
		GBL_FAILED_FUNCTION_NAME="Fn_AWSummary_TabButton_Operation"
		Fn_AWSummary_TabButton_Operation=False
		Set tabsummary = UIAWindow("Teamcenter RAC (Eclipse)").UIATab("Summary")
		set objDesc =  Description.Create
		objDesc("controltype").Value="Group"
		Set tabobj = tabsummary.ChildObjects(objDesc)
		If Instr(ObjTab.ToString(),"LeftTab") Then
			For iCount = 1 To tabobj.Count - 1
				If tabobj(iCount).GetROProperty("name")="Primary" Then
					ObjTab.SetTOProperty "path",tabobj(iCount).GetROProperty("path")
					ObjTab.SetTOProperty "hwnd",tabobj(iCount).GetROProperty("hwnd")
					Exit For
				End If
			Next
		End If 
		If Instr(ObjTab.ToString(),"Righthandtoolbar") Then
			For iCount = 1 To tabobj.Count - 1
			If tabobj(iCount).GetROProperty("name")="Right hand toolbar" Then
				ObjTab.SetTOProperty "path",tabobj(iCount).GetROProperty("path")
				ObjTab.SetTOProperty "hwnd",tabobj(iCount).GetROProperty("hwnd")
				ObjTab.SetTOProperty "name",tabobj(iCount).GetROProperty("name")
				Exit For
			End If
			Next
		End If 
		Select Case sAction
		
				Case "ButtonExist"
						sbuttonArray = Split(sAwsButton,"~")
						For iCount = 0 To UBound(sbuttonArray)

									ObjTab.UIAButton("TabUIButton").SetTOProperty "name",sbuttonArray(iCount)
									Set objAwsButton = ObjTab.UIAButton("TabUIButton")
							
									bReturn= Fn_AWS_UI_ObjectExist("Fn_AWSummary_TabButton_Operation", objAwsButton)
									If  bReturn=True Then
									  	  Fn_AWSummary_TabButton_Operation = True
									  	  Exit Function
									Else			
										  Fn_AWSummary_TabButton_Operation = False
									  	  Exit Function
									End If
									Set objAwsButton = Nothing 
							Next
		
				Case "ButtonClick"
						
						sbuttonArray = Split(sAwsButton,"~")
						For iCount = 0 To UBound(sbuttonArray)

									ObjTab.UIAButton("TabUIButton").SetTOProperty "name",sbuttonArray(iCount)
									Set objAwsButton = ObjTab.UIAButton("TabUIButton")
							
									bReturn = Fn_UIAButton_Operations("Fn_AWSummary_TabButton_Operation", "Click", ObjTab,"TabUIButton")
									If  bReturn=True Then
											
									  	  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "UI_AWS : Sucessfully clicked on Button '" & sAwsButton & "' of Function " & sFunctionName)
									  	  Fn_AWSummary_TabButton_Operation = True
									  	  Exit Function
									  	  
									Else			
										  Fn_AWSummary_TabButton_Operation = False
									  	  Exit Function
									End If
									Set objAwsButton = Nothing 
							Next
		End Select

End Function

'#######################################################################################################
'###    FUNCTION NAME   :   Fn_AWSummary_HederTabText_Exist()
'###
'###    DESCRIPTION     :   This function is used to check existance of Text in Heder tab
'###
'###    PARAMETERS      :   sFunctionName : Valid Function name, 
'###                       						 ObjTab : Valid Dialog box object name,
'###                        					sAwsButton   : Valid Text
'###
'###    CREATED BY      :   Amruta         16/07/2021    
'###    MODIFIED BY     :   
'###    EXAMPLE         :   
'#######################################################################################################

Function Fn_AWSummary_HederTabText_Exist(ObjTab, sAwsText)

	Dim objAwsButton, bReturn,sbuttonArray,iCount,objAwsText
	'Object Creation
	GBL_FAILED_FUNCTION_NAME="Fn_AWSummary_HederTabText_Exist"
	Fn_AWSummary_HederTabText_Exist=False

					If Instr(ObjTab.toString,"Tab") > 0 Then
					
						sTextArray = Split(sAwsText,"~")
						For iCount = 0 To UBound(sTextArray)

									ObjTab.UIAObject("HederText").SetTOProperty "name",sTextArray(iCount)
									Set objAwsText = ObjTab.UIAObject("HederText")
							
									bReturn= Fn_AWS_UI_ObjectExist("Fn_AWSummary_HederTabText_Exist", objAwsText)
									If  bReturn=True Then
									  	  Fn_AWSummary_HederTabText_Exist = True
									  	  Exit Function
									  	  
									Else			
										  Fn_AWSummary_HederTabText_Exist = False
									  	  Exit Function
									End If
									Set objAwsText = Nothing 
							Next
					End If
End Function

'#######################################################################################################
'###    FUNCTION NAME   :   Fn_AWSummary_Popupwindow_Button_Operation()
'###
'###    DESCRIPTION     :   This function is used perform button operation in AW sumary tab
'###
'###    PARAMETERS      :   sFunctionName : Valid Function name, 
'###                       						 sAction : action name
'###                        					sAwsButton   : Valid Button name
'###
'###    HISTORY         :   AUTHOR         DATE          
'###
'###    CREATED BY      :   Amruta         16/07/2021    
'###    MODIFIED BY     :   
'###    EXAMPLE         :   
'#######################################################################################################

Function Fn_AWSummary_Popupwindow_Button_Operation(sAction,sAwsButton)

	Dim objAwsButton, bReturn,sbuttonArray,iCount,buttonobj,tabsummary,objDesc,tabobj,ObjPopupwin
	'Object Creation
	
		Fn_AWSummary_Popupwindow_Button_Operation=False
		Set tabsummary = UIAWindow("Teamcenter RAC (Eclipse)").UIATab("Summary")
		Set ObjPopupwin = UIAWindow("Teamcenter RAC (Eclipse)").UIATab("Summary").UIAObject("popupmenuwin")
		set objDesc =  Description.Create
		objDesc("nativeclass").Value="ng-scope"
		Set tabobj = tabsummary.ChildObjects(objDesc)
		If tabobj.Count = 0 Then
			Fn_AWSummary_Popupwindow_Button_Operation = True
			Exit Function
		End If
		For iCount = 1 To tabobj.Count - 1
			If Instr(tabobj(iCount).GetROProperty("name"),"ui-id") Then
				ObjPopupwin.SetTOProperty "path",tabobj(iCount).GetROProperty("path") 
				ObjPopupwin.SetTOProperty "name",tabobj(iCount).GetROProperty("name") 
				ObjPopupwin.SetTOProperty "hwnd",tabobj(iCount).GetROProperty("hwnd") 
			End If
		Next
		
		set objDesc =  Description.Create
		objDesc("controltype").Value="button"
		Set tabobj = ObjPopupwin.ChildObjects(objDesc)
		If tabobj.Count = 0 Then
			Fn_AWSummary_Popupwindow_Button_Operation = True
			Exit Function
		End If
		For iCount = 1 To tabobj.Count - 1
			If tabobj(iCount).GetROProperty("name") = sAwsButton Then
				ObjPopupwin.UIAButton("UIButton").SetTOProperty "path",tabobj(iCount).GetROProperty("path") 
				ObjPopupwin.UIAButton("UIButton").SetTOProperty "name",tabobj(iCount).GetROProperty("name") 
				ObjPopupwin.UIAButton("UIButton").SetTOProperty "hwnd",tabobj(iCount).GetROProperty("hwnd") 
			End If
		Next
		
		
		Select Case sAction		
				Case "ButtonClick"
						bReturn = Fn_UIAButton_Operations("Fn_AWSummary_TabButton_Operation", "Click",ObjPopupwin,"UIButton")
						If  bReturn=True Then
						  	  Fn_AWSummary_Popupwindow_Button_Operation = True
						  	  Exit Function
						Else			
							  Fn_AWSummary_Popupwindow_Button_Operation = False
						  	  Exit Function
						End If
						Set ObjPopupwin = Nothing 
		End Select

End Function
'#######################################################################################################
'###
'###    Function Name	:	Fn_AW_Summary_ProfilePanel_Operation
'###
'###    Description		:	Function Used to Perform operation on Profile panel
'###
'###    Parameters		:	1.strAction : Action Name
'###					:	 2: strNode Name
'###
'###    Return Value	: 	True Or False
'###                            
'###
'###    CREATED BY      :   Pratiksha         16/07/2021    
'###    MODIFIED BY     :   
'###    EXAMPLE         :   call Fn_AW_Summary_ProfilePanel_Operation("linkExist","Designer","")
'#######################################################################################################
'
Function Fn_AW_Summary_ProfilePanel_Operation(sAction,text,StrPopupMenu)
	
	Dim objAwsButton, bReturn,stextArray,iCount,tabobj,tabsummary,objDesc,ObjTab
	'Object Creation
	GBL_FAILED_FUNCTION_NAME="Fn_AW_Summary_ProfilePanel_Operation"
	Fn_AW_Summary_ProfilePanel_Operation=False
		Set ObjTab = UIAWindow("Teamcenter RAC (Eclipse)").UIATab("Summary").UIAObject("ProfileNavigation Panel")
		Set tabsummary = UIAWindow("Teamcenter RAC (Eclipse)").UIATab("Summary")
		tabsummary.Highlight
		set objDesc =  Description.Create
		objDesc("controltype").Value="Group"
		Set tabobj = tabsummary.ChildObjects(objDesc)
		For iCount = 1 To tabobj.Count - 1
			If tabobj(iCount).GetROProperty("name")="Navigation Panel" Then
				ObjTab.SetTOProperty "path",tabobj(iCount).GetROProperty("path")
				ObjTab.SetTOProperty "hwnd",tabobj(iCount).GetROProperty("hwnd")
	        Exit For
			End If
		Next
					
		Select Case sAction
		
				Case "linkExist"
						stextArray = Split(text,"~")
						For iCount = 0 To UBound(stextArray)

									ObjTab.UIAHyperlink("HyperLink").SetTOProperty "name",stextArray(iCount)
									Set objAwslink = ObjTab.UIAHyperlink("HyperLink")
							
									bReturn= Fn_AWS_UI_ObjectExist("Fn_AW_Summary_ProfilePanel_Operation", objAwslink)
									If  bReturn=True Then
									  	  Fn_AW_Summary_ProfilePanel_Operation = True
										  Exit Function
									Else			
											'Report error/message when UI Button object is disable.
										  Fn_AW_Summary_ProfilePanel_Operation = False
										  Exit Function
									End If
									'Clear memory of WebButton object.
									Set objAwsButton = Nothing 
							Next
							
End Select
Set ObjTab = Nothing
Set tabsummary = Nothing
End Function

'#######################################################################################################
'###    FUNCTION NAME   :   Fn_Tab_Popupwindow_Button_Operation()
'###
'###    DESCRIPTION     :   This function is used perform RMB operation on Lefttab ot righttab in Summry and AW 
'###
'###    PARAMETERS      :   sFunctionName : Valid Function name, 
'###                       						 sAction : action name
'###                        					sAwsButton   : Valid Button name
'###
'###    HISTORY         :   AUTHOR         DATE          
'###
'###    CREATED BY      :   Amruta         25/11/2021    
'###    MODIFIED BY     :   
'###    EXAMPLE         :   
'#######################################################################################################
Function Fn_Tab_Popupwindow_Button_Operation(sAction,sAwsButton)

	Dim objAwsButton, bReturn,sbuttonArray,iCount,buttonobj,tabsummary,objDesc,tabobj,ObjPopupwin,ObjPopupmenu
	'Object Creation
	
		Fn_Tab_Popupwindow_Button_Operation=False
		Set tabsummary = UIAWindow("Teamcenter RAC (Eclipse)")
		Set ObjPopupwin = UIAWindow("Teamcenter RAC (Eclipse)").UIAMenu("TabPopup")
		Set ObjPopupmenu = UIAWindow("Teamcenter RAC (Eclipse)").UIAMenu("TabPopup").UIAMenu("TabPopupMenu")
		
		set objDesc =  Description.Create
		objDesc("name").Value="Web context"
		
		Set tabobj = tabsummary.ChildObjects(objDesc)
		If tabobj.Count = 0 Then
			Fn_Tab_Popupwindow_Button_Operation = True
			Exit Function
		Else
				ObjPopupwin.SetTOProperty "path",tabobj(iCount).GetROProperty("path") 
				ObjPopupwin.SetTOProperty "name",tabobj(iCount).GetROProperty("name") 
				ObjPopupwin.SetTOProperty "hwnd",tabobj(iCount).GetROProperty("hwnd") 
		End If
		
		
		
		set objDesc =  Description.Create
		objDesc("name").Value="Web context"
		
		Set tabobj = ObjPopupwin.ChildObjects(objDesc)
		If tabobj.Count = 0 Then
			Fn_Tab_Popupwindow_Button_Operation = True
			Exit Function
		Else
				ObjPopupmenu.SetTOProperty "path",tabobj(iCount).GetROProperty("path") 
				ObjPopupmenu.SetTOProperty "name",tabobj(iCount).GetROProperty("name") 
				ObjPopupmenu.SetTOProperty "hwnd",tabobj(iCount).GetROProperty("hwnd") 
		End If
		set objDesc =  Description.Create
		objDesc("name").Value= sAwsButton
		Set tabobj = ObjPopupmenu.ChildObjects(objDesc)
		If tabobj.Count = 0 Then
			Fn_Tab_Popupwindow_Button_Operation = False
			Exit Function
		Else
				ObjPopupmenu.UIAObject("Item").SetTOProperty "path",tabobj(0).GetROProperty("path") 
				ObjPopupmenu.UIAObject("Item").SetTOProperty "name",tabobj(0).GetROProperty("name") 
				ObjPopupmenu.UIAObject("Item").SetTOProperty "hwnd",tabobj(0).GetROProperty("hwnd") 
		End If
		ObjPopupmenu.UIAObject("Item").Highlight
		
		Select Case sAction		
				Case "Select"
						ObjPopupmenu.UIAObject("Item").Click
						If  ObjPopupmenu.Exist(5) Then
						  	  Fn_Tab_Popupwindow_Button_Operation = False
						  	  Exit Function
						Else			
							  Fn_Tab_Popupwindow_Button_Operation = True
						  	  Exit Function
						End If
						Set ObjPopupwin = Nothing 
		End Select

End Function


'#######################################################################################################
'###
'###    Function Name		:	  Fn_AW_BOM_TreeOperation()
'###
'###    Description			:	Function Used to perform operation on BOMtable Node from Active workspace tab
'###
'###    Parameters			   :	1.strAction : Action Name
'###											:	 2: strNode Name
'###
'###    Return Value		   	   : 	True Or False
'###                            
'###
'###    CREATED BY      :   Amruta         16/07/2021    
'###    MODIFIED BY     :   
'###    EXAMPLE         :   call Fn_AW_BOM_TreeOperation("Select","Top","")
'###						call Fn_AW_BOM_TreeOperation("Select","Top","Expand Below")
'###						
'#######################################################################################################

Public Function Fn_AW_BOM_TreeOperation(strAction,strNode,sMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_AW_BOM_TreeOperation"
   	Fn_AW_BOM_TreeOperation = FALSE
   	Dim objDesc,TabActWrkspc,TableElmt,Colnames,Arrcolname,ColNum,iCount,RowNum,objNavtree
   	Dim ObjPopup,Objbtn,Objwindec,Objpopupwin,iCounter,bReturn
		
		Set TabActWrkspc = UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace")
		Set TableElmt = UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIATable("ElementTable")
			
		Set objDesc =  Description.Create
		objDesc("nativeclass").Value = "aw-splm-tableContainer"
		UIAWindow("Teamcenter RAC (Eclipse)").Activate()
		wait 2
		  UIAWindow("Teamcenter RAC (Eclipse)").Maximize()
		  wait 2
		Set objNavtree =UIAWindow("Teamcenter RAC (Eclipse)").ChildObjects(objDesc)
		If objNavtree.Count = 0 Then
			UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").Highlight
			Set objNavtree =UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").ChildObjects(objDesc)
			If objNavtree.Count = 0  Then
				 UIAWindow("Teamcenter RAC (Eclipse)").Highlight
				 wait 2
				Set objNavtree =UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").ChildObjects(objDesc)
			End If
		End If
		If objNavtree.Count = 0  Then
			Fn_AW_BOM_TreeOperation = False	
			Exit Function
		Else
   		wait 2
			For icount = 0 To objNavtree.Count-1
				If InStr(objNavtree(iCount).GetROProperty("path"),"occTreeTable") Then
					TableElmt.SetTOProperty "path",objNavtree(iCount).GetROProperty("path")
					TableElmt.SetTOProperty "hwnd",objNavtree(iCount).GetROProperty("hwnd")
					If TableElmt.Exist(5) Then
						TableElmt.Highlight
						Exit For
					Else
						Fn_AW_BOM_TreeOperation = FALSE
					End If
			
				End If
			Next
	
	End If		
			
			
	Select Case strAction

		Case "Select"
		
			If TableElmt.Exist(5) Then
				TableElmt.Highlight
			Else
				Fn_AW_BOM_TreeOperation = FALSE
			End If
			wait 2			
			Colnames = TableElmt.ColumnHeaders()
			Arrcolname = Split(Colnames,VbLf)
			For iCount = 0 To UBound(Arrcolname)-1
				If Arrcolname(iCount) = "Element" Then
					ColNum = iCount
					Exit For
				End If
			Next
			RowNum =TableElmt.GetROProperty("rowcount")
			
			For iCount = 0 To RowNum - 1
				RowNames = TableElmt.GetCellName(iCount,ColNum)
				RowNames = Split(RowNames," ")
				If RowNames(UBound(RowNames)) = strNode Then
					TableElmt.ClickCell iCount,ColNum
					Fn_AW_BOM_TreeOperation = True
					Exit Function
				Else
					If iCount = RowNum - 1 Then
						Fn_AW_BOM_TreeOperation = False
						Exit Function
					End If
				End If
			Next
			
		Case "Popupmenu"	
			
			Set ObjPopup = UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("popupmenuwin")
			Set Objbtn = UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("popupmenuwin").UIAButton("UIButton")

			If TableElmt.Exist(5) Then
					TableElmt.Highlight
				Else
					Fn_AW_BOM_TreeOperation = FALSE
				End If
				wait 2			
				Colnames = TableElmt.ColumnHeaders()
				Arrcolname = Split(Colnames,VbLf)
				For iCount = 0 To UBound(Arrcolname)-1
					If Arrcolname(iCount) = "Element" Then
						ColNum = iCount
						Exit For
					End If
				Next
				RowNum =TableElmt.GetROProperty("rowcount")
				For iCount = 0 To RowNum - 1 
					RowNames = TableElmt.GetCellName(iCount,ColNum)
					RowNames = Split(RowNames," ")
					If RowNames(UBound(RowNames)) = strNode Then
						TableElmt.ClickCell iCount,ColNum, micNoCoordinate,micNoCoordinate,micRightBtn 
						wait 2
							Set Objwindec = Description.Create
							Objwindec("nativeclass").Value = "ng-scope"
							Set Objpopupwin =TabActWrkspc.ChildObjects(Objwindec)
							If Objpopupwin.Count <> 0 Then
								For iCounter = 1 To Objpopupwin.Count - 1
									If Instr(Objpopupwin(iCounter).GetROProperty("name"),"ui-id-") Then
										ObjPopup.SetTOProperty "path",Objpopupwin(iCounter).GetROProperty("path") 
										ObjPopup.SetTOProperty "name",Objpopupwin(iCounter).GetROProperty("name") 
										ObjPopup.SetTOProperty "hwnd",Objpopupwin(iCounter).GetROProperty("hwnd") 
										Exit For
									End If
								Next
								set Objbtndesc =  Description.Create
								objDesc("controltype").Value="button"
								Set Objbtn = ObjPopup.ChildObjects(Objbtndesc)
								For iCounter = 1 To Objbtn.Count - 1
									If Objbtn(iCounter).GetROProperty("name") = sMenu Then
										UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("popupmenuwin").UIAButton("UIButton").SetTOProperty "path",Objbtn(iCounter).GetROProperty("path") 
										UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("popupmenuwin").UIAButton("UIButton").SetTOProperty "name",Objbtn(iCounter).GetROProperty("name") 
										UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("popupmenuwin").UIAButton("UIButton").SetTOProperty "hwnd",Objbtn(iCounter).GetROProperty("hwnd") 
										Exit For
									End If
								Next
								bReturn = Fn_UIAButton_Operations("Fn_AW_BOM_TreeOperation", "Click",ObjPopup, "UIButton")
								If bReturn = True Then
									Fn_AW_BOM_TreeOperation = True
									Exit Function
								Else
									Fn_AW_BOM_TreeOperation = False
									Exit Function
								End If
								
							Else
								Fn_AW_BOM_TreeOperation = False
								Exit Function
							End If
					Else
						If iCount = RowNum - 1 Then
							Fn_AW_BOM_TreeOperation = False
							Exit Function
						End If
						
					End If
				Next

	End Select
	Set TabActWrkspc = Nothing
	Set TableElmt = Nothing
End Function

'#######################################################################################################
'###
'###    Function Name		:	  Fn_AW_Tab_Operation()
'###
'###    Description			:	Function Used to select tab in Active workspace 
'###
'###    Parameters			   :	1.strAction : Action Name
'###											:	 2: strTab Name
'###
'###    Return Value		   	   : 	True Or False
'###                            
'###
'###    CREATED BY      :   Amruta         16/07/2021    
'###    MODIFIED BY     :   
'###    EXAMPLE         :   call Fn_AW_Tab_Operation("Select","Overview")
'#######################################################################################################
Public Function Fn_AW_Tab_Operation(strAction,strTab)
	GBL_FAILED_FUNCTION_NAME="Fn_AW_Tab_Operation"
   	Fn_AW_Tab_Operation = FALSE
   	Dim objDesc,TabActWrkspc,ObjTablnk,ObjTablinks,iCount,bflag
   	
   	Set TabActWrkspc = UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace")
   	Set ObjTablnk = UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAHyperlink("TabLink")
	Select Case strAction

		Case "Select"
			UIAWindow("Teamcenter RAC (Eclipse)").Activate
			UIAWindow("Teamcenter RAC (Eclipse)").Maximize()
			TabActWrkspc.Highlight
			wait 5
			Set objDesc =  Description.Create
			objDesc("controltype").Value = "Hyperlink"
			wait 2
			TabActWrkspc.Highlight
			wait 2
			Set ObjTablinks = TabActWrkspc.ChildObjects(objDesc)
			If ObjTablinks.Count = 0 Then
				UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").Highlight
				Set ObjTablinks =UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").ChildObjects(objDesc)
				If ObjTablinks.Count = 0  Then
					 UIAWindow("Teamcenter RAC (Eclipse)").Highlight
					 wait 2
					 Set objNavtree =UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").ChildObjects(objDesc)
				End If
			End If
			bflag = False
			
			If ObjTablinks.Count = 0 Then
					Fn_AW_Tab_Operation = False
					Exit Function
			Else
					For icount = 0 To ObjTablinks.Count-1
						If ObjTablinks(iCount).GetROProperty("name")= strTab Then
							ObjTablnk.SetTOProperty "path", ObjTablinks(iCount).GetROProperty("path")
							ObjTablnk.SetTOProperty "hwnd", ObjTablinks(iCount).GetROProperty("hwnd")
							ObjTablnk.SetTOProperty "name", ObjTablinks(iCount).GetROProperty("name")
							ObjTablnk.Click()
							wait 2
							bflag = True
							Fn_AW_Tab_Operation = True
							Exit Function
						End If
					Next
					If bflag = False Then
							Set ObjListitm = UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAList("MyTasklist").UIAObject("ListItem")
							
							Set ObjDesc = Description.Create
							ObjDesc("controltype").Value = "Image"
							Set Objlist = TabActWrkspc.ChildObjects(ObjDesc)
							For iCount  = 0 To Objlist.Count - 1
								If Instr(Objlist(iCount).GetROProperty("path"),"Secondary") Then
									TabActWrkspc.UIAObject("ArrowImage").SetTOProperty "path",Objlist(iCount).GetROProperty("path")
									TabActWrkspc.UIAObject("ArrowImage").SetTOProperty "hwnd",Objlist(iCount).GetROProperty("hwnd")
									TabActWrkspc.UIAObject("ArrowImage").SetTOProperty "nativeclass",Objlist(iCount).GetROProperty("nativeclass")
									TabActWrkspc.UIAObject("ArrowImage").SetTOProperty "name",Objlist(iCount).GetROProperty("name")
									wait 2
									TabActWrkspc.UIAObject("ArrowImage").Click
									wait 2
									Exit For
								End If
							Next
							Set ObjDesc = Description.Create
							ObjDesc("controltype").Value = "List"
							Set Objlist = TabActWrkspc.ChildObjects(ObjDesc)
							For iCount  = 0 To Objlist.Count - 1
								If Instr(Objlist(iCount).GetROProperty("path"),"Secondary") Then
									TabActWrkspc.UIAList("MyTasklist").SetTOProperty "path",Objlist(iCount).GetROProperty("path")
									TabActWrkspc.UIAList("MyTasklist").SetTOProperty "hwnd",Objlist(iCount).GetROProperty("hwnd")
									TabActWrkspc.UIAList("MyTasklist").SetTOProperty "nativeclass",Objlist(iCount).GetROProperty("nativeclass")
									TabActWrkspc.UIAList("MyTasklist").SetTOProperty "name",Objlist(iCount).GetROProperty("name")
									Exit For
								End If
							Next
			
							Set ObjDesc = Description.Create
							ObjDesc("controltype").Value = "ListItem"
							Set Objlistitem = TabActWrkspc.UIAList("MyTasklist").ChildObjects(ObjDesc)
							For iCount  = 0 To Objlistitem.Count - 1
								If Instr(Objlistitem(iCount).GetROProperty("name"),"Workflow") Then
									TabActWrkspc.UIAList("MyTasklist").UIAObject("ListItem").SetTOProperty "name",Objlistitem(iCount).GetROProperty("name")
									TabActWrkspc.UIAList("MyTasklist").UIAObject("ListItem").SetTOProperty "nativeclass",Objlistitem(iCount).GetROProperty("nativeclass")
									TabActWrkspc.UIAList("MyTasklist").UIAObject("ListItem").SetTOProperty "path",Objlistitem(iCount).GetROProperty("path")
									TabActWrkspc.UIAList("MyTasklist").UIAObject("ListItem").SetTOProperty "hwnd",Objlistitem(iCount).GetROProperty("hwnd")
									wait 5
									ObjListitm.Click micNoCoordinate,micNoCoordinate,micLefttBtn
									Fn_AW_Tab_Operation = True
								Exit For
								End If
							Next

					End If
			End If

	End Select
	Set TabActWrkspc = Nothing
	Set ObjTablnk = Nothing
End Function

'#######################################################################################################
'###
'###    Function Name		:	  Fn_AW_PropertiesTab_Operation()
'###
'###    Description			:	Function Used to verify text in properties tab of in Active workspace tab
'###
'###    Parameters			   :	1.strAction : Action Name
'###											:	 2: strText Name
'###
'###    Return Value		   	   : 	True Or False
'###                            
'###
'###    CREATED BY      :   Amruta         16/07/2021    
'###    MODIFIED BY     :   
'###    EXAMPLE         :   call Fn_AW_PropertiesTab_Operation("TextVerify","250")
'#######################################################################################################

Public Function Fn_AW_PropertiesTab_Operation(strAction,strText)
	GBL_FAILED_FUNCTION_NAME="Fn_AW_PropertiesTab_Operation"
   	Fn_AW_PropertiesTab_Operation = FALSE
	Dim TabActWrkspc,Objtab,Objtabtext,objDesc,ObjPropTab,Objdesctext,iCount,iCounter,flag

	Set TabActWrkspc = UIAWindow("Teamcenter RAC (Eclipse)")
   	Set Objtab = UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("TabProperties")
	Select Case strAction
		Case "TextVerify"
			If TabActWrkspc.Exist(5) Then
				TabActWrkspc.Highlight
				wait 5
			Else	
				Fn_AW_PropertiesTab_Operation = False
			End  If
			
			
			Set objDesc =  Description.Create
			objDesc("controltype").Value = "Group"
			TabActWrkspc.Highlight
			wait 5
			Set ObjPropTab = TabActWrkspc.ChildObjects(objDesc)
			If ObjPropTab.Count <> 0 Then
					For icount = 0 To ObjPropTab.Count-1
					If Instr(ObjPropTab(iCount).GetROProperty("path"),"Properties") Then
							Objtab.SetTOProperty "path", ObjPropTab(iCount).GetROProperty("path")
							Objtab.SetTOProperty "name", ObjPropTab(iCount).GetROProperty("name")
							Objtab.SetTOProperty "hwnd", ObjPropTab(iCount).GetROProperty("hwnd")
'							If Objtab.Exist(5) Then
'								Objtab.Highlight
'								wait 2
'							Else
'								Fn_AW_PropertiesTab_Operation = False
'								Exit Function
'							End If
'							
							
							Set Objdesctext = Description.Create
							Objdesctext("controltype").Value = "Text"
							 
							Set ObjtText = Objtab.ChildObjects(Objdesctext)
							
							If ObjtText.Count <> 0 Then
									For iCounter = 0 To ObjtText.Count - 1
										If ObjtText(iCounter).GetROProperty("name") = strText Then
											flag = True
											Exit For
										End If
									Next		
							Else
								Fn_AW_PropertiesTab_Operation = False
								Exit Function
							End If
					End If
					
					If flag = True Then
						Fn_AW_PropertiesTab_Operation = True
						Exit Function
					End  If 
				Next
			Else
				Fn_AW_PropertiesTab_Operation = False
				Exit Function
			End If

	End Select
	Set TabActWrkspc = Nothing
	Set Objtab = Nothing
End  Function	


'#######################################################################################################
'###
'###    Function Name		:	  Fn_AW_TabButton_Operation
'###
'###    Description			:	Function Used to button operation in AW tab
'###
'###    Parameters			   :	1.strAction : Action Name
'###								2. ObjTab	:	Tab Object (Leffttab ,RightTab) 
'###								3.sAwsButton : Button name
'###    Return Value		   	   : 	True Or False
'###                            
'###
'###    CREATED BY      :   Amruta         16/07/2021    
'###    MODIFIED BY     :   
'###    EXAMPLE         :   call Fn_AW_TabButton_Operation("ButtonExist","UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("Righthandtoolbar")","Inbox")
'#######################################################################################################

Function Fn_AW_TabButton_Operation(sAction,ObjTab, sAwsButton)

	Dim objAwsButton, bReturn,sbuttonArray,iCount,tabobj,tabAW,objDesc,objViewOption
	'Object Creation
		GBL_FAILED_FUNCTION_NAME = "Fn_AW_TabButton_Operation"
		Fn_AW_TabButton_Operation=False
		Set tabAW = UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace")
		set objDesc =  Description.Create
		objDesc("controltype").Value="Group"
		tabAW.Highlight
		wait 2
		Set tabobj = UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").ChildObjects(objDesc)
		Set objViewOption = UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("Teamcenter_Inbox")
		wait 2
		If tabobj.Count = 0 Then
			Set tabobj = UIAWindow("Teamcenter RAC (Eclipse)").ChildObjects(objDesc)
		End If
		Set tabobj = UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").ChildObjects(objDesc)
		If Instr(ObjTab.ToString(),"LeftTab") Then
			For iCount = 1 To tabobj.Count - 1
				If tabobj(iCount).GetROProperty("name")="Primary" Then
					ObjTab.SetTOProperty "path",tabobj(iCount).GetROProperty("path")
					ObjTab.SetTOProperty "hwnd",tabobj(iCount).GetROProperty("hwnd")
					Exit For
				End If
			Next
		End If 
		If Instr(ObjTab.ToString(),"Righthandtoolbar") Then
			For iCount = 1 To tabobj.Count - 1
			If tabobj(iCount).GetROProperty("name")="Right hand toolbar" Then
				ObjTab.SetTOProperty "path",tabobj(iCount).GetROProperty("path")
				ObjTab.SetTOProperty "hwnd",tabobj(iCount).GetROProperty("hwnd")
				ObjTab.SetTOProperty "name",tabobj(iCount).GetROProperty("name")
				Exit For
			End If
			Next
		End If 
		Select Case sAction
		
				Case "ButtonExist"
					If Instr(ObjTab.toString,"Tab") > 0 Then
					
						sbuttonArray = Split(sAwsButton,"~")
						For iCount = 0 To UBound(sbuttonArray)

									ObjTab.UIAButton("TabUIButton").SetTOProperty "name",sbuttonArray(iCount)
									Set objAwsButton = ObjTab.UIAButton("TabUIButton")
							
									bReturn= Fn_AWS_UI_ObjectExist("Fn_AW_TabButton_Operation", objAwsButton)
									If  bReturn=True Then
									  	  Fn_AW_TabButton_Operation = True
									  	  Exit Function
									Else			
										  Fn_AW_TabButton_Operation = False
									  	  Exit Function
									End If
									Set objAwsButton = Nothing 
							Next
					End If
		
				Case "ButtonClick"
						
									ObjTab.UIAButton("TabUIButton").SetTOProperty "name",sAwsButton
									Set objAwsButton = ObjTab.UIAButton("TabUIButton")
									wait 2					
									bReturn = Fn_UIAButton_Operations("Fn_AW_TabButton_Operation", "Click", ObjTab,"TabUIButton")
									If  bReturn=True Then
											
									  	  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "UI_AWS : Sucessfully clicked on Button '" & sAwsButton & "' of Function " & sFunctionName)
									  	  Fn_AW_TabButton_Operation = True
									  	  Exit Function
									  	  
									Else			
										  Fn_AW_TabButton_Operation = False
									  	  Exit Function
										  
									End If
									Set objAwsButton = Nothing 
									
				Case "ClickViewOption" 						   
						   
						         If Fn_UIAButton_Operations("Fn_AW_TabButton_Operation", "Click", objViewOption,"SelectionViewOption")= True Then		              
						           objViewOption.UIAButton("SelectionViewOption").SetTOProperty "name",sAwsButton
						           If Fn_AWS_UI_ObjectExist("Fn_AW_TabButton_Operation", objViewOption.UIAButton("SelectionViewOption"))= True Then
						             If Fn_UIAButton_Operations("Fn_AW_TabButton_Operation", "Click", objViewOption,"SelectionViewOption")= True Then
						               Fn_AW_TabButton_Operation = True
						             Else
						                Fn_AW_TabButton_Operation = False
										Exit Function
						             End If
						           Else
						             Fn_AW_TabButton_Operation = False
									 Exit Function									 
						           End If	
						         Else
					                Fn_AW_TabButton_Operation = False
					                Exit Function
						         End If
						         
		End Select

End Function

'#######################################################################################################
'###    FUNCTION NAME   :   Fn_AW_HederTabText_Exist()
'###
'###    DESCRIPTION     :   This function is used to check existance of Text in Heder tab
'###
'###    PARAMETERS      :   sFunctionName : Valid Function name, 
'###                       						 ObjTab : Valid Dialog box object,
'###                        					sAwsText   : Valid Text
'###
'###    CREATED BY      :   Amruta         16/07/2021    
'###    MODIFIED BY     :   
'###    EXAMPLE         :   
'#######################################################################################################

Function Fn_AW_HederTabText_Exist(ObjTab, sAwsText)

	Dim objAwsButton, bReturn,sbuttonArray,iCount,sText,ObjDesc,ObjText,ObjNodepath
	'Object Creation
	GBL_FAILED_FUNCTION_NAME = "Fn_AW_HederTabText_Exist"
	Fn_AW_HederTabText_Exist=False

					If Instr(ObjTab.toString,"Tab") > 0 Then
					Set ObjNodepath = UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("NodepathText")
						sTextArray = Split(sAwsText,"~")
						For iCount = 0 To UBound(sTextArray)
									
								If Instr(sTextArray(iCount),">") Then
										sTextArray(iCount) = Replace(sTextArray(iCount),">"," Breadcrumb ")
										
										Set ObjDesc = Description.Create
										ObjDesc("controltype").Value = "Text"
										
										Set ObjText = ObjTab.ChildObjects(ObjDesc)
										For iCounter = 0 To ObjText.Count - 1
											If ObjText(iCounter).GetROProperty("name") = sTextArray(iCount) Then
												ObjNodepath.SetTOProperty "hwnd",ObjText(iCounter).GetROProperty("hwnd")
												ObjNodepath.SetTOProperty "path",ObjText(iCounter).GetROProperty("path")
												ObjNodepath.SetTOProperty "name",sTextArray(iCount)
												bReturn= Fn_AWS_UI_ObjectExist("Fn_AW_HederTabText_Exist", ObjNodepath)
												Exit For
											End If
										Next
									Else
												ObjTab.UIAObject("HederText").SetTOProperty "name",sTextArray(iCount)
												Set objAwsText = ObjTab.UIAObject("HederText")	
												bReturn= Fn_AWS_UI_ObjectExist("Fn_AW_HederTabText_Exist", objAwsText)
									End If
									
									If  bReturn=True Then
									  	  Fn_AW_HederTabText_Exist = True
									  	  Exit Function
									  	  
									Else			
										  Fn_AW_HederTabText_Exist = False
									  	  Exit Function
									End If
									Set objAwsText = Nothing 
						Next
					End If
					
End Function

'#######################################################################################################
'###
'###    Function Name	:	Fn_AW_ProfilePanel_Operation
'###
'###    Description		:	Function Used to Perform operation on Profile panel
'###
'###    Parameters		:	1.strAction : Action Name
'###					:	 2: strNode Name
'###
'###    Return Value	: 	True Or False
'###                            
'###
'###    CREATED BY      :   Pratiksha         16/07/2021    
'###    MODIFIED BY     :   
'###    EXAMPLE         :   call Fn_AW_ProfilePanel_Operation("linkExist","Designer","")
'#######################################################################################################
'
Function Fn_AW_ProfilePanel_Operation(sAction,text,StrPopupMenu)
	
	Dim objAwsButton, bReturn,stextArray,iCount,tabobj,tabsummary,objDesc,ObjTab
	'Object Creation
	GBL_FAILED_FUNCTION_NAME="Fn_AW_ProfilePanel_Operation"
	Fn_AW_ProfilePanel_Operation=False
		Set ObjTab = UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("ProfileNavigation Panel")
		Set tabsummary = UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace")
		tabsummary.Highlight
		set objDesc =  Description.Create
		objDesc("controltype").Value="Group"
		Set tabobj = tabsummary.ChildObjects(objDesc)
		For iCount = 1 To tabobj.Count - 1
			If tabobj(iCount).GetROProperty("name")="Navigation Panel" Then
				ObjTab.SetTOProperty "path",tabobj(iCount).GetROProperty("path")
				ObjTab.SetTOProperty "hwnd",tabobj(iCount).GetROProperty("hwnd")
	        Exit For
			End If
		Next
					
		Select Case sAction
		
				Case "linkExist"
						stextArray = Split(text,"~")
						For iCount = 0 To UBound(stextArray)

									ObjTab.UIAHyperlink("HyperLink").SetTOProperty "name",stextArray(iCount)
									Set objAwslink = ObjTab.UIAHyperlink("HyperLink")
							
									bReturn= Fn_AWS_UI_ObjectExist("Fn_AW_ProfilePanel_Operation", objAwslink)
									If  bReturn=True Then
									  	  Fn_AW_ProfilePanel_Operation = True
										  Exit Function
									Else			
											'Report error/message when UI Button object is disable.
										  Fn_AW_ProfilePanel_Operation = False
										  Exit Function
									End If
									'Clear memory of WebButton object.
									Set objAwsButton = Nothing 
							Next
							
							
							
					Case "linkpopup"

									ObjTab.UIAHyperlink("HyperLink").SetTOProperty "name",text
									Set objAwslink = ObjTab.UIAHyperlink("HyperLink")
									aMenu = split(StrPopupMenu,":",-1,1)
									objAwslink.Highlight
									objAwslink.Click micNoCoordinate,micNoCoordinate,micRightBtn
									wait 2
									
									Set UImenu = UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("ProfileNavigation Panel").UIAMenu("Webcontext")
									UImenu.Select "Refresh"
								
								If  bReturn=True Then
								  	  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "UI_AWS : Sucessfully verified existance of AWSButton '" & text & "' of Function " & sFunctionName)
								  	  Fn_AW_ProfilePanel_Operation = True
								Else			
										'Report error/message when UI Button object is disable.
									  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "UI_AWS : AWSButton " &text &"  is not exist of Function " & sFunctionName)
									  Fn_AW_ProfilePanel_Operation = False
								End If
								'Clear memory of WebButton object.
								Set objAwsButton = Nothing 
End Select
Set ObjTab = Nothing
Set tabsummary = Nothing
End Function

'#######################################################################################################
'###    FUNCTION NAME   :   Fn_AWSummary_DatasetDownload_Operation()
'###
'###    DESCRIPTION     :   This function is used to Perform dataset download operation in summary tab
'###
'###    PARAMETERS      :   sFunctionName : Valid Function name, 
'###                       	ObjTab : Valid Dialog box object name,
'###                        sAwsButton   : Valid Button name
'###
'###    HISTORY         :   AUTHOR         DATE          
'###
'###    CREATED BY      :  Pratiksha      12/08/2021    
'###    MODIFIED BY     :   
'###    EXAMPLE         :  call Fn_AWSummary_DatasetDownload_Operation("ButtonClick","Open file","") 
'#######################################################################################################

Function Fn_AWSummary_DatasetDownload_Operation(sAction,ObjTab, sAwsButton)

	Dim objAwsButton, bReturn,sbuttonArray,iCount,tabobj,tabdataset,objDesc
	'Object Creation
		GBL_FAILED_FUNCTION_NAME="Fn_AWSummary_DatasetDownload_Operation"
		Fn_AWSummary_DatasetDownload_Operation=False
		Set tabdataset = UIAWindow("DatasetDownload Panel")
		set objDesc =  Description.Create
		objDesc("controltype").Value="ToolBar"
		tabdataset.Highlight
		wait 2
		Set tabobj = tabdataset.ChildObjects(objDesc)
		wait 2
		If Instr(ObjTab.ToString(),"Downloadsbar") Then
			For iCount = 1 To tabobj.Count - 1
			If tabobj(iCount).GetROProperty("name")="Downloads bar" Then
				ObjTab.SetTOProperty "path",tabobj(iCount).GetROProperty("path")
				ObjTab.SetTOProperty "hwnd",tabobj(iCount).GetROProperty("hwnd")
				ObjTab.SetTOProperty "name",tabobj(iCount).GetROProperty("name")
				Exit For
			End If
			Next
		End If 
		Select Case sAction
		
				Case "ButtonClick"
						
						sbuttonArray = Split(sAwsButton,"~")
						For iCount = 0 To UBound(sbuttonArray)

									ObjTab.UIAButton("TabUIButton").SetTOProperty "name",sbuttonArray(iCount)
									Set objAwsButton = ObjTab.UIAButton("TabUIButton")
							
									bReturn = Fn_UIAButton_Operations("Fn_AWSummary_DatasetDownload_Operation", "Click", ObjTab,"TabUIButton")
									If  bReturn=True Then
											
									  	  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "UI_AWS : Sucessfully clicked on Button '" & sAwsButton & "' of Function " & sFunctionName)
									  	  Fn_AWSummary_DatasetDownload_Operation = True
									  	  Exit Function
									  	  
									Else			
										  Fn_AWSummary_DatasetDownload_Operation = False
									  	  Exit Function
									End If
									Set objAwsButton = Nothing 
							Next
		End Select

End Function


'#######################################################################################################
'###
'###    Function Name		:	  Fn_AW_MyTasks_Operation()
'###
'###    Description			:	Function Used to perform operation on Mytasks tab Node in Active workspace 
'###
'###    Parameters			   :	1.strAction : Action Name
'###											: 2: node Name
'###
'###    Return Value		   	   : 	True Or False
'###                            
'###
'###    CREATED BY      :   Amruta          30/08/2021    
'###    MODIFIED BY     :   
'###    EXAMPLE         :   call Fn_AW_MyTasks_Operation("Select","029776/A;1-Item1")
'###						call Fn_AW_MyTasks_Operation("VerifyPropertyInList","Targets: 029776/A;1-Item1")
'###						call Fn_AW_MyTasks_Operation("VerifyPropertyInList","Targets: 029776/A;1-Item1~Assignee: Engineering/Designer/*")
'#######################################################################################################
Public Function Fn_AW_MyTasks_Operation(strAction,nodename)
	Dim nodeitem,ObjAW,ObjDesc,Objlistitem,iCount,ObjListitm
	Dim aProperties,iPropcnt,aPropValue,iPropVal,bFlag
	GBL_FAILED_FUNCTION_NAME="Fn_AW_MyTasks_Operation"
   	Fn_AW_MyTasks_Operation = FALSE
		Select Case strAction
		Case "Select"
				Set ObjAW = UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace")
				Set ObjListitm =UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAList("MyTasklist").UIAObject("ListItem")
				ObjAW.Highlight()
				Set ObjDesc = Description.Create
				ObjDesc("controltype").Value = "List"
				Set Objlist = ObjAW.ChildObjects(ObjDesc)
				For iCount  = 0 To Objlist.Count - 1
					If Instr(Objlist(iCount).GetROProperty("path"),"list") Then
						ObjAW.UIAList("MyTasklist").SetTOProperty "path",Objlist(iCount).GetROProperty("path")
						ObjAW.UIAList("MyTasklist").SetTOProperty "hwnd",Objlist(iCount).GetROProperty("hwnd")
						Exit For
					End If
				Next
						
				Set ObjDesc = Description.Create
				ObjDesc("controltype").Value = "ListItem"
				Set Objlistitem = ObjAW.UIAList("MyTasklist").ChildObjects(ObjDesc)
				For iCount  = 0 To Objlistitem.Count - 1
					If Instr(Objlistitem(iCount).GetROProperty("name"),nodename) Then
						ObjAW.UIAList("MyTasklist").UIAObject("ListItem").SetTOProperty "name",Objlistitem(iCount).GetROProperty("name")
						ObjAW.UIAList("MyTasklist").UIAObject("ListItem").SetTOProperty "path",Objlistitem(iCount).GetROProperty("path")
						ObjAW.UIAList("MyTasklist").UIAObject("ListItem").SetTOProperty "hwnd",Objlistitem(iCount).GetROProperty("hwnd")
						wait 5
						ObjListitm.Click micNoCoordinate,micNoCoordinate,micLefttBtn 
						Fn_AW_MyTasks_Operation = True
					Exit For
					End If
				Next
				Set ObjAW = Nothing
				Set ObjDesc = Nothing
		
			Case "VerifyPropertyInList"			''Added by Radha Mane
				Set ObjAW = UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace")
				Set ObjListitm =UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAList("MyTasklist").UIAObject("ListItem")
				ObjAW.Highlight()
				Set ObjDesc = Description.Create
				ObjDesc("controltype").Value = "List"
				Set Objlist = ObjAW.ChildObjects(ObjDesc)
				For iCount  = 0 To Objlist.Count - 1
					If Instr(Objlist(iCount).GetROProperty("path"),"list") Then
						ObjAW.UIAList("MyTasklist").SetTOProperty "path",Objlist(iCount).GetROProperty("path")
						ObjAW.UIAList("MyTasklist").SetTOProperty "hwnd",Objlist(iCount).GetROProperty("hwnd")
						Exit For
					End If
				Next
						
				Set ObjDesc = Description.Create
				ObjDesc("controltype").Value = "Text"
				Set Objlistitem = ObjAW.UIAList("MyTasklist").ChildObjects(ObjDesc)
				aProperties = Split(nodename,"~")		'Split Properties list by ~
				For iPropcnt = 0 To UBound(aProperties) Step 1
					aPropValue = Split(aProperties(iPropcnt)," ")		'Split Property and value by space
					For iPropVal = 0 To UBound(aPropValue) Step 1
						bFlag = False
						For iCount  = 0 To Objlistitem.Count - 1
							If Instr(Objlistitem(iCount).GetROProperty("name"),aPropValue(iPropVal)) Then
								ObjAW.UIAList("MyTasklist").UIAObject("PropertyInList").SetTOProperty "name",Objlistitem(iCount).GetROProperty("name")
								ObjAW.UIAList("MyTasklist").UIAObject("PropertyInList").SetTOProperty "path",Objlistitem(iCount).GetROProperty("path")
								ObjAW.UIAList("MyTasklist").UIAObject("PropertyInList").SetTOProperty "hwnd",Objlistitem(iCount).GetROProperty("hwnd")
								bFlag = True
								Fn_AW_MyTasks_Operation = TRUE
								Exit For
							End If
						Next
						If bFlag = False Then
							Fn_AW_MyTasks_Operation = FALSE
							Exit For
						End If
					Next
				Next
				
				Set ObjAW = Nothing
				Set ObjDesc = Nothing
				
			Case "VerifyAndSelect"'Case is used to verify e.g Simple Do TAsk 2and select task with name  Exmaple call Fn_AW_MyTasks_Operation("VerifyAndSelect","Simple Do 2: 029776/A;1-Item1")
				Set ObjAW = UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace")
				Set ObjListitm =UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAList("MyTasklist").UIAObject("ListItem")
				ObjAW.Highlight()
				Set ObjDesc = Description.Create
				ObjDesc("controltype").Value = "List"
				Set Objlist = ObjAW.ChildObjects(ObjDesc)
				For iCount  = 0 To Objlist.Count - 1
					If Instr(Objlist(iCount).GetROProperty("path"),"list") Then
						ObjAW.UIAList("MyTasklist").SetTOProperty "path",Objlist(iCount).GetROProperty("path")
						ObjAW.UIAList("MyTasklist").SetTOProperty "hwnd",Objlist(iCount).GetROProperty("hwnd")
						Exit For
					End If
				Next
						
				Set ObjDesc = Description.Create
				ObjDesc("controltype").Value = "ListItem"
				Set Objlistitem = ObjAW.UIAList("MyTasklist").ChildObjects(ObjDesc)
				nodename = Split(nodename,":")		'Split nodename  by :
				For iCount  = 0 To Objlistitem.Count - 1
					If Instr(Objlistitem(iCount).GetROProperty("name"),nodename(1))<>0 AND Instr(Objlistitem(iCount).GetROProperty("path"),nodename(0))<>0 Then
						ObjAW.UIAList("MyTasklist").UIAObject("ListItem").SetTOProperty "name",Objlistitem(iCount).GetROProperty("name")
						ObjAW.UIAList("MyTasklist").UIAObject("ListItem").SetTOProperty "path",Objlistitem(iCount).GetROProperty("path")
						ObjAW.UIAList("MyTasklist").UIAObject("ListItem").SetTOProperty "hwnd",Objlistitem(iCount).GetROProperty("hwnd")
						wait 5
						ObjListitm.Click micNoCoordinate,micNoCoordinate,micLefttBtn 
						Fn_AW_MyTasks_Operation = True
					Exit For
					End If
				Next
				Set ObjAW = Nothing
				Set ObjDesc = Nothing	
				
			
		End Select
	
End Function
'#######################################################################################################
'###
'###    Function Name		:	  Fn_AW_AddTarget_Operation()
'###
'###    Description			:	Function Used to ADD target
'###
'###    Parameters			   :	1.strAction : Action Name
'###											: 2: node Name
'###
'###    Return Value		   	   : 	True Or False
'###                            
'###
'###    CREATED BY      :   Amruta          03/09/2021    
'###    MODIFIED BY     :   
'###    EXAMPLE         :   call Fn_AW_MyTasks_Operation("Drawing",dicDetails)
'#######################################################################################################
Public Function Fn_AW_AddTarget_Operation(Targettype,dicDetails)
	Dim ObjDesc,iCount,Objgrp
	GBL_FAILED_FUNCTION_NAME="Fn_AW_AddTarget_Operation"
   	Fn_AW_AddTarget_Operation = FALSE
		Set objDesc = Description.Create
		objDesc("controltype").Value = "Group"
		Set Objgrp = UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").ChildObjects(objDesc)
		If Objgrp.Count <> 0 Then
			For iCount = 0 To Objgrp.Count - 1
				If Instr(Objgrp(iCount).getROProperty("name"),"All Targets") Then
					UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("All Targets").SetTOProperty "path",Objgrp(iCount).GetROProperty("path")
					UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("All Targets").SetTOProperty "hwnd",Objgrp(iCount).GetROProperty("hwnd")
					Exit For
				End If
			Next
		Else
		   	Fn_AW_AddTarget_Operation = FALSE
		End If
		
		Set objDesc = Description.Create
		objDesc("controltype").Value = "Button"
		Set Objgrp = UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("All Targets").ChildObjects(objDesc)
		If Objgrp.Count <> 0 Then
			For iCount = 0 To Objgrp.Count - 1
				If Instr(Objgrp(iCount).getROProperty("name"),"Add") Then
					UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("All Targets").UIAButton("Add to").SetTOProperty "path",Objgrp(iCount).GetROProperty("path")
					UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("All Targets").UIAButton("Add to").SetTOProperty "hwnd",Objgrp(iCount).GetROProperty("hwnd")
					UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("All Targets").UIAButton("Add to").Click
					Exit For
				End If
			Next	
		Else
		   	Fn_AW_AddTarget_Operation = FALSE
		End If
		wait 5
		
		Set objDesc = Description.Create
		objDesc("controltype").Value = "Edit"
		Set Objgrp = UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").ChildObjects(objDesc)
		If Objgrp.Count <> 0 Then
			For iCount = 0 To Objgrp.Count - 1
					If Instr(Objgrp(iCount).getROProperty("name"),"Filter") Then
						UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAEdit("Filter").SetTOProperty "name" ,Objgrp(iCount).getROProperty("name")
						UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAEdit("Filter").SetTOProperty "path" ,Objgrp(iCount).getROProperty("path")
						UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAEdit("Filter").SetTOProperty "hwnd" ,Objgrp(iCount).getROProperty("hwnd")
						wait 2
						UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAEdit("Filter").Click
						UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAEdit("Filter").Type Targettype
						wait 2
						call Fn_KeyBoardOperation("SendKey", "{TAB}{Enter}")
						Exit For
					End If
				Next	
		Else
			Fn_AW_AddTarget_Operation = FALSE
		End If
		
			Set objDesc = Description.Create
			objDesc("controltype").Value = "Group"
			Set Objgrp = UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").ChildObjects(objDesc)
			If Objgrp.Count <> 0 Then
				For iCount = 0 To Objgrp.Count - 1
					If Instr(Objgrp(iCount).getROProperty("name"),"Task Panel") Then
						UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("Task Panel").UIAObject("Inboxproperties").SetTOProperty "path" ,Objgrp(iCount).getROProperty("path")
						UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("Task Panel").UIAObject("Inboxproperties").SetTOProperty "hwnd" ,Objgrp(iCount).getROProperty("hwnd")
						wait 2
						Exit For
					End If
				Next
			Else
				Fn_AW_AddTarget_Operation = FALSE
			End If
			Set objDesc = Description.Create
			objDesc("controltype").Value = "Group"
			Set Objgrp = UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("Task Panel").ChildObjects(objDesc)
			If Objgrp.Count <> 0 Then
				For iCount = 0 To Objgrp.Count - 1
					If Instr(Objgrp(iCount).getROProperty("path"),"Properties") Then
						UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("Task Panel").UIAObject("Inboxproperties").SetTOProperty "path" ,Objgrp(iCount).getROProperty("path")
						UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("Task Panel").UIAObject("Inboxproperties").SetTOProperty "hwnd" ,Objgrp(iCount).getROProperty("hwnd")
						wait 2
						UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("Task Panel").UIAObject("Inboxproperties").Highlight
						Exit For
					End If
				Next
			Else
				Fn_AW_AddTarget_Operation = FALSE
			End If
			Set objDesc = Description.Create
			objDesc("controltype").Value = "Edit"
			Set Objgrp = UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("Task Panel").UIAObject("Inboxproperties").ChildObjects(objDesc)
			If Objgrp.Count <> 0 Then
				For iCount = 0 To Objgrp.Count - 1
					If Instr(Objgrp(iCount).getROProperty("name"),"ID") Then
						UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("Task Panel").UIAObject("Inboxproperties").UIAEdit("TargetName").SetTOProperty "path" ,Objgrp(iCount).getROProperty("path")
						UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("Task Panel").UIAObject("Inboxproperties").UIAEdit("TargetName").SetTOProperty "hwnd" ,Objgrp(iCount).getROProperty("hwnd")
						UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("Task Panel").UIAObject("Inboxproperties").UIAEdit("TargetName").SetTOProperty "name" ,Objgrp(iCount).getROProperty("name")
						
						wait 2
						UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("Task Panel").UIAObject("Inboxproperties").UIAEdit("TargetName").Click
						wait 2
						Call Fn_KeyBoardOperation("SendKey", "^(a)")
						 If dicDetails("ID") <> "" Then
							UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("Task Panel").UIAObject("Inboxproperties").UIAEdit("TargetName").Type  dicDetails("ID")
						 End If	
						Exit For
					End If
				Next	
			Else
				Fn_AW_AddTarget_Operation = FALSE
			End If
			Set objDesc = Description.Create
			objDesc("controltype").Value = "Edit"
			Set Objgrp = UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("Task Panel").UIAObject("Inboxproperties").ChildObjects(objDesc)
			If Objgrp.Count <> 0 Then
				For iCount = 0 To Objgrp.Count - 1
					If Instr(Objgrp(iCount).getROProperty("name"),"Name") Then
						UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("Task Panel").UIAObject("Inboxproperties").UIAEdit("TargetName").SetTOProperty "path" ,Objgrp(iCount).getROProperty("path")
						UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("Task Panel").UIAObject("Inboxproperties").UIAEdit("TargetName").SetTOProperty "hwnd" ,Objgrp(iCount).getROProperty("hwnd")
						UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("Task Panel").UIAObject("Inboxproperties").UIAEdit("TargetName").SetTOProperty "name" ,Objgrp(iCount).getROProperty("name")
						wait 2
						UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("Task Panel").UIAObject("Inboxproperties").UIAEdit("TargetName").Click
						 If dicDetails("Name") <> "" Then
							UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("Task Panel").UIAObject("Inboxproperties").UIAEdit("TargetName").Type  dicDetails("Name")
						 End If	
						Exit For
					End If
				Next	
			Else
				Fn_AW_AddTarget_Operation = FALSE
			End If
			Set objDesc = Description.Create
			objDesc("controltype").Value = "Button"
			Set Objgrp = UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("Task Panel").ChildObjects(objDesc)
			If Objgrp.Count <> 0 Then
				For iCount = 0 To Objgrp.Count - 1
					If Instr(Objgrp(iCount).getROProperty("name"),"Add") Then
						UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("Task Panel").UIAButton("AddTarget").SetTOProperty "path" ,Objgrp(iCount).getROProperty("path")
						UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("Task Panel").UIAButton("AddTarget").SetTOProperty "hwnd" ,Objgrp(iCount).getROProperty("hwnd")
						wait 2
						UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("Task Panel").UIAButton("AddTarget").Click
						Fn_AW_AddTarget_Operation = TRUE
						Exit For
					End If
				Next
			Else
				Fn_AW_AddTarget_Operation = FALSE
			End If
End Function

'#######################################################################################################
'###
'###    Function Name		:	  Fn_AW_TargetTable_Operation()
'###
'###    Description			:	Function Used to perform operation on target table
'###
'###    Parameters			   :	1.strAction : Action Name
'###											: 2: node Name
'###                                              3:Column name
'###
'###    Return Value		   	   : 	True Or False
'###                            
'###
'###    CREATED BY      :   Amruta          03/09/2021    
'###    MODIFIED BY     :   
'###    EXAMPLE         :   call Fn_AW_TargetTable_Operation("Exist","029776/A;1-Item1","Object")
'#######################################################################################################
Public Function Fn_AW_TargetTable_Operation(strAction,nodename,Columnname)
	Dim nodeitem,ObjAW,ObjDesc,Objlistitem,iCount,ObjListitm
	GBL_FAILED_FUNCTION_NAME="Fn_AW_TargetTable_Operation"
   	Fn_AW_TargetTable_Operation = FALSE
		Set objDesc = Description.Create
		objDesc("controltype").Value = "Group"
		Set Objgrp = UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").ChildObjects(objDesc)
		If Objgrp.Count <> 0 Then
			For iCount = 0 To Objgrp.Count - 1
				If Instr(Objgrp(iCount).getROProperty("name"),"All Targets") Then
					UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("All Targets").SetTOProperty "path",Objgrp(iCount).GetROProperty("path")
					UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("All Targets").SetTOProperty "hwnd",Objgrp(iCount).GetROProperty("hwnd")
					Exit For
				End If
			Next
		Else
		   	Fn_AW_TargetTable_Operation = FALSE
		End If
		Set objDesc = Description.Create
		objDesc("controltype").Value = "DataGrid"
		Set Objgrp = UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("All Targets").ChildObjects(objDesc)
		If Objgrp.Count <> 0 Then
				For iCount = 0 To Objgrp.Count - 1
					If Instr(Objgrp(iCount).getROProperty("path"),"All Targets") Then
						UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("All Targets").UIATable("Tarrgettable").SetTOProperty "path",Objgrp(iCount).GetROProperty("path")
						UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("All Targets").UIATable("Tarrgettable").SetTOProperty "hwnd",Objgrp(iCount).GetROProperty("hwnd")
						Exit For
					End If
				Next
		Else
			   	Fn_AW_TargetTable_Operation = FALSE
		End  If
		
		Select Case strAction
		Case "Exist"
				Colnames = UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("All Targets").UIATable("Tarrgettable").ColumnHeaders()
				Arrcolname = Split(Colnames,VbLf)
				For iCount = 0 To UBound(Arrcolname)-1
					If Arrcolname(iCount) = Columnname Then
						ColNum = iCount - 1
						Exit For
					End If
				Next
			
				RowNum = UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("All Targets").UIATable("Tarrgettable").GetROProperty("rowcount")
				
				For iCount = 0 To RowNum - 1
					RowNames = UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("All Targets").UIATable("Tarrgettable").GetCellName(iCount,ColNum)
					If RowNames <> "" Then
						If RowNames = nodename Then
						Fn_AW_TargetTable_Operation = True
						Exit Function
						Else
							If iCount = RowNum - 1 Then
								Fn_AW_TargetTable_Operation = False
								Exit Function
							End If
						End If
					End If
				Next
End Select

End  Function

'#######################################################################################################
'###
'###    Function Name		:	  Fn_AW_AddPreference_Operation()
'###
'###    Description			:	Function Used to add Reference
'###
'###    Parameters			   :	1.Reftype : Ref type
'###											: 2: dicDetails -Name ,ID
'###
'###    Return Value		   	   : 	True Or False
'###                            
'###
'###    CREATED BY      :   Amruta          03/09/2021    
'###    MODIFIED BY     :   
'###    EXAMPLE         :   call Fn_AW_RefTable_Operation("Drawing",dicobject)
'#######################################################################################################
Public Function Fn_AW_AddPreference_Operation(Reftype,dicDetails)
	Dim ObjDesc,iCount,Objgrp
	GBL_FAILED_FUNCTION_NAME="Fn_AW_AddPreference_Operation"
   	Fn_AW_AddPreference_Operation = FALSE
   	
		Set objDesc = Description.Create
		objDesc("controltype").Value = "Group"
		Set Objgrp = UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").ChildObjects(objDesc)
		If Objgrp.Count <> 0 Then
			For iCount = 0 To Objgrp.Count - 1
				If Instr(Objgrp(iCount).getROProperty("name"),"All References") Then
					UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("All References").SetTOProperty "path",Objgrp(iCount).GetROProperty("path")
					UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("All References").SetTOProperty "hwnd",Objgrp(iCount).GetROProperty("hwnd")
					Exit For
				End If
			Next	
		Else
   			Fn_AW_AddPreference_Operation = FALSE
		End If
		
		Set objDesc = Description.Create
		objDesc("controltype").Value = "Button"
		Set Objgrp = UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("All References").ChildObjects(objDesc)
		If Objgrp.Count <> 0 Then
			For iCount = 0 To Objgrp.Count - 1
				If Instr(Objgrp(iCount).getROProperty("name"),"Add") Then
					UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("All References").UIAButton("Add to").SetTOProperty "path",Objgrp(iCount).GetROProperty("path")
					UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("All References").UIAButton("Add to").SetTOProperty "hwnd",Objgrp(iCount).GetROProperty("hwnd")
					UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("All References").UIAButton("Add to").Click
					Exit For
				End If
			Next
			wait 5
		Else
   			Fn_AW_AddPreference_Operation = FALSE
		End If
		
			Set objDesc = Description.Create
			objDesc("controltype").Value = "Edit"
			Set Objgrp = UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").ChildObjects(objDesc)
			If Objgrp.Count <> 0 Then
				For iCount = 0 To Objgrp.Count - 1
					If Instr(Objgrp(iCount).getROProperty("name"),"Filter") Then
						UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAEdit("Filter").SetTOProperty "name" ,Objgrp(iCount).getROProperty("name")
						UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAEdit("Filter").SetTOProperty "path" ,Objgrp(iCount).getROProperty("path")
						UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAEdit("Filter").SetTOProperty "hwnd" ,Objgrp(iCount).getROProperty("hwnd")
						wait 2
						UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAEdit("Filter").Click
						UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAEdit("Filter").Type Reftype
						wait 2
						call Fn_KeyBoardOperation("SendKey", "{TAB}{Enter}")
						Exit For
					End If
				Next	
			Else
   				Fn_AW_AddPreference_Operation = FALSE
			End If
			
			Set objDesc = Description.Create
			objDesc("controltype").Value = "Group"
			Set Objgrp = UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").ChildObjects(objDesc)
			If Objgrp.Count <> 0 Then
				For iCount = 0 To Objgrp.Count - 1
					If Instr(Objgrp(iCount).getROProperty("name"),"Task Panel") Then
						UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("Task Panel").UIAObject("Inboxproperties").SetTOProperty "path" ,Objgrp(iCount).getROProperty("path")
						UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("Task Panel").UIAObject("Inboxproperties").SetTOProperty "hwnd" ,Objgrp(iCount).getROProperty("hwnd")
						wait 2
						Exit For
					End If
				Next	
			Else
   				Fn_AW_AddPreference_Operation = FALSE
			End If
			
			
			Set objDesc = Description.Create
			objDesc("controltype").Value = "Group"
			Set Objgrp = UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("Task Panel").ChildObjects(objDesc)
			If Objgrp.Count <> 0 Then
				For iCount = 0 To Objgrp.Count - 1
					If Instr(Objgrp(iCount).getROProperty("path"),"Properties") Then
						UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("Task Panel").UIAObject("Inboxproperties").SetTOProperty "path" ,Objgrp(iCount).getROProperty("path")
						UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("Task Panel").UIAObject("Inboxproperties").SetTOProperty "hwnd" ,Objgrp(iCount).getROProperty("hwnd")
						wait 2
						UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("Task Panel").UIAObject("Inboxproperties").Highlight
						Exit For
					End If
				Next
			Else
   				Fn_AW_AddPreference_Operation = FALSE
			End If
			
			Set objDesc = Description.Create
			objDesc("controltype").Value = "Edit"
			Set Objgrp = UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("Task Panel").UIAObject("Inboxproperties").ChildObjects(objDesc)
			If Objgrp.Count <> 0 Then
				For iCount = 0 To Objgrp.Count - 1
					If Instr(Objgrp(iCount).getROProperty("name"),"ID") Then
						UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("Task Panel").UIAObject("Inboxproperties").UIAEdit("TargetName").SetTOProperty "path" ,Objgrp(iCount).getROProperty("path")
						UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("Task Panel").UIAObject("Inboxproperties").UIAEdit("TargetName").SetTOProperty "hwnd" ,Objgrp(iCount).getROProperty("hwnd")
						UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("Task Panel").UIAObject("Inboxproperties").UIAEdit("TargetName").SetTOProperty "name" ,Objgrp(iCount).getROProperty("name")
						
						wait 2
						UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("Task Panel").UIAObject("Inboxproperties").UIAEdit("TargetName").Click
						wait 2
						Call Fn_KeyBoardOperation("SendKey", "^(a)")
						 If dicDetails("ID") <> "" Then
							UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("Task Panel").UIAObject("Inboxproperties").UIAEdit("TargetName").Type  dicDetails("ID")
						 End If	
						Exit For
					End If
				Next	
			Else
   				Fn_AW_AddPreference_Operation = FALSE
			End If
			
			Set objDesc = Description.Create
			objDesc("controltype").Value = "Edit"
			Set Objgrp = UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("Task Panel").UIAObject("Inboxproperties").ChildObjects(objDesc)
			If Objgrp.Count <> 0 Then
				For iCount = 0 To Objgrp.Count - 1
					If Instr(Objgrp(iCount).getROProperty("name"),"Name") Then
						UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("Task Panel").UIAObject("Inboxproperties").UIAEdit("TargetName").SetTOProperty "path" ,Objgrp(iCount).getROProperty("path")
						UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("Task Panel").UIAObject("Inboxproperties").UIAEdit("TargetName").SetTOProperty "hwnd" ,Objgrp(iCount).getROProperty("hwnd")
						UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("Task Panel").UIAObject("Inboxproperties").UIAEdit("TargetName").SetTOProperty "name" ,Objgrp(iCount).getROProperty("name")
						
						wait 2
						UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("Task Panel").UIAObject("Inboxproperties").UIAEdit("TargetName").Click
						 If dicDetails("Name") <> "" Then
							UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("Task Panel").UIAObject("Inboxproperties").UIAEdit("TargetName").Type  dicDetails("Name")
						 End If	
						Exit For
					End If
				Next
			Else
   				Fn_AW_AddPreference_Operation = FALSE
			End If
			
			
			Set objDesc = Description.Create
			objDesc("controltype").Value = "Button"
			Set Objgrp = UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("Task Panel").ChildObjects(objDesc)
			If Objgrp.Count <> 0 Then
				For iCount = 0 To Objgrp.Count - 1
					If Instr(Objgrp(iCount).getROProperty("name"),"Add") Then
						UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("Task Panel").UIAButton("AddTarget").SetTOProperty "path" ,Objgrp(iCount).getROProperty("path")
						UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("Task Panel").UIAButton("AddTarget").SetTOProperty "hwnd" ,Objgrp(iCount).getROProperty("hwnd")
						wait 2
						UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("Task Panel").UIAButton("AddTarget").Click
   						Fn_AW_AddPreference_Operation = TRUE
						Exit For
					End If
				Next
			Else
   				Fn_AW_AddPreference_Operation = FALSE
			End If
			
End Function

'#######################################################################################################
'###
'###    Function Name		:	 Fn_AW_RefTable_Operation()
'###
'###    Description			:	Function Used to perform operation on All Reference table
'###
'###    Parameters			   :	1.strAction : Action Name
'###											: 2: node Name
'###                                              3:Column name    
'###
'###    Return Value		   	   : 	True Or False
'###                            
'###
'###    CREATED BY      :   Amruta          03/09/2021    
'###    MODIFIED BY     :   
'###    EXAMPLE         :   call Fn_AW_RefTable_Operation("Exist","029776/A;1-Item1","Object")
'#######################################################################################################
Public Function Fn_AW_RefTable_Operation(strAction,nodename,Columnname)
	Dim ObjDesc,iCount,Objgrp,Colnames,Arrcolname,RowNum,RowNames
	GBL_FAILED_FUNCTION_NAME="Fn_AW_RefTable_Operation"
   	Fn_AW_RefTable_Operation = FALSE
		Set objDesc = Description.Create
		objDesc("controltype").Value = "Group"
		Set Objgrp = UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").ChildObjects(objDesc)
		If  Objgrp.Count <> 0 Then
			For iCount = 0 To Objgrp.Count - 1
				If Instr(Objgrp(iCount).getROProperty("name"),"All References") Then
					UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("All References").SetTOProperty "path",Objgrp(iCount).GetROProperty("path")
					UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("All References").SetTOProperty "hwnd",Objgrp(iCount).GetROProperty("hwnd")
					Exit For
				End If
			Next
		Else
   			Fn_AW_RefTable_Operation = FALSE
		End If
		
		Set objDesc = Description.Create
		objDesc("controltype").Value = "DataGrid"
		Set Objgrp = UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("All References").ChildObjects(objDesc)
		If  Objgrp.Count <> 0 Then
			For iCount = 0 To Objgrp.Count - 1
				If Instr(Objgrp(iCount).getROProperty("path"),"All References") Then
					UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("All References").UIATable("reftable").SetTOProperty "path",Objgrp(iCount).GetROProperty("path")
					UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("All References").UIATable("reftable").SetTOProperty "hwnd",Objgrp(iCount).GetROProperty("hwnd")
					Exit For
				End If
			Next
		Else
   			Fn_AW_RefTable_Operation = FALSE
		End If
		Select Case strAction
		Case "Exist"
				Colnames = UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("All References").UIATable("reftable").ColumnHeaders()
				Arrcolname = Split(Colnames,VbLf)
				For iCount = 0 To UBound(Arrcolname)-1
					If Arrcolname(iCount) = Columnname Then
						ColNum = iCount - 1
						Exit For
					End If
				Next
			
				RowNum = UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("All References").UIATable("reftable").GetROProperty("rowcount")
				
				For iCount = 0 To RowNum - 1
					RowNames = UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("All References").UIATable("reftable").GetCellName(iCount,ColNum)
					If RowNames <> "" Then
						If RowNames = nodename Then
							Fn_AW_RefTable_Operation = True
							Exit Function
						Else
							If iCount = RowNum - 1 Then
								Fn_AW_RefTable_Operation = False
								Exit Function
							End If
						End If
					End If
				Next
End Select
End  Function

'#######################################################################################################
'###
'###    Function Name		:	 Fn_AW_WorklistTaskOperation()
'###
'###    Description			:	Function Used to perform operation on Worklist Task
'###
'###    Parameters			   :	1.strAction : Action Name
'###							 				: 2: sComment : Comment
'###
'###    Return Value		   	   : 	True Or False
'###                            
'###
'###    CREATED BY      :   Radha Mane          07/09/2021    
'###    MODIFIED BY     :   
'###    EXAMPLE         :   call Fn_AW_WorklistTaskOperation("Complete","Commented task here")
'#######################################################################################################

Function Fn_AW_WorklistTaskOperation(sButtonName,sDescription)
	Dim objDesc,Objgrp,iCount,objButton
	
	Set objDesc = Description.Create
	objDesc("controltype").Value = "Edit"
	Set Objgrp = UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").ChildObjects(objDesc)
	If Objgrp.Count <> 0 Then
		For iCount = 0 To Objgrp.Count - 1
			If Instr(Objgrp(iCount).getROProperty("name"),"Comments") Then
				UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAEdit("Filter").SetTOProperty "name" ,Objgrp(iCount).getROProperty("name")
				UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAEdit("Filter").SetTOProperty "path" ,Objgrp(iCount).getROProperty("path")
				UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAEdit("Filter").SetTOProperty "hwnd" ,Objgrp(iCount).getROProperty("hwnd")
				wait 2
				UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAEdit("Filter").Click
				UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAEdit("Filter").Type sDescription
				wait 2
				call Fn_KeyBoardOperation("SendKey", "{TAB}")
				Exit For
			End If
		Next	
	Else
		Fn_AW_WorklistTaskOperation = FALSE
	End If

	If sButtonName <> "" Then
		Set objButton = UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAButton("Button")
		objButton.SetTOProperty "name",sButtonName
		objButton.Click
		Fn_AW_WorklistTaskOperation = True
	Else
		Fn_AW_WorklistTaskOperation = False
	End If
	Set objDesc = Nothing
	Set Objgrp = Nothing
End Function

'#######################################################################################################
'###    FUNCTION NAME   :   Fn_AW_Popupwindow_Button_Operation()
'###
'###    DESCRIPTION     :   This function is used perform button operation in AW  tab
'###
'###    PARAMETERS      :   sFunctionName : Valid Function name, 
'###                       	sAction : action name
'###                        sAwsButton   : Valid Button name
'###
'###    HISTORY         :   AUTHOR         DATE          
'###
'###    CREATED BY      :   Pratiksha      16/09/2021    
'###    MODIFIED BY     :   
'###    EXAMPLE         :   Fn_AW_Popupwindow_Button_Operation("ButtonClick","Claim Workflow Task")
'###                        Fn_AW_Popupwindow_Button_Operation("ButtonClick","Suspend")
'###					    Fn_AW_Popupwindow_Button_Operation("ButtonClick","Resume")
'#######################################################################################################

Function Fn_AW_Popupwindow_Button_Operation(sAction,sAwsButton)

	Dim objAwsButton, bReturn,sbuttonArray,iCount,buttonobj,tabAW,objDesc,tabobj,ObjPopupwin
	'Object Creation
	
		Fn_AW_Popupwindow_Button_Operation=False
		Set tabAW = UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace")
		Set ObjPopupwin = UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("popupmenuwin")
		set objDesc =  Description.Create
		objDesc("nativeclass").Value="ng-scope"
		Set tabobj = tabAW.ChildObjects(objDesc)
		If tabobj.Count = 0 Then
			Fn_AW_Popupwindow_Button_Operation = True
			Exit Function
		End If
		For iCount = 1 To tabobj.Count - 1
			If Instr(tabobj(iCount).GetROProperty("name"),"ui-id") Then
				ObjPopupwin.SetTOProperty "path",tabobj(iCount).GetROProperty("path") 
				ObjPopupwin.SetTOProperty "name",tabobj(iCount).GetROProperty("name") 
				ObjPopupwin.SetTOProperty "hwnd",tabobj(iCount).GetROProperty("hwnd") 
			End If
		Next
		
		set objDesc =  Description.Create
		objDesc("controltype").Value="button"
		Set tabobj = ObjPopupwin.ChildObjects(objDesc)
		If tabobj.Count = 0 Then
			Fn_AW_Popupwindow_Button_Operation = True
			Exit Function
		End If
		For iCount = 1 To tabobj.Count - 1
			If tabobj(iCount).GetROProperty("name") = sAwsButton Then
				ObjPopupwin.UIAButton("UIButton").SetTOProperty "path",tabobj(iCount).GetROProperty("path") 
				ObjPopupwin.UIAButton("UIButton").SetTOProperty "name",tabobj(iCount).GetROProperty("name") 
				ObjPopupwin.UIAButton("UIButton").SetTOProperty "hwnd",tabobj(iCount).GetROProperty("hwnd") 
			End If
		Next
		
		
		Select Case sAction		
				Case "ButtonClick"
						bReturn = Fn_UIAButton_Operations("Fn_AW_Popupwindow_Button_Operation", "Click",ObjPopupwin,"UIButton")
						If  bReturn=True Then
						  	  Fn_AW_Popupwindow_Button_Operation = True
						  	  Exit Function
						Else			
							  Fn_AW_Popupwindow_Button_Operation = False
						  	  Exit Function
						End If
						Set ObjPopupwin = Nothing 
		End Select

End Function

'#######################################################################################################
'###
'###    Function Name		:	 Fn_AW_TaskOperation()
'###
'###    Description			:	Function Used to perform operations on Task 
'###
'###    Parameters		   :	1:strAction : Action Name
'###					   :                  sButtonName   : Valid Button name
'###                                          sDescription : Comment
'###                                          sStatus : Task Status
'###    Return Value	   : 	True Or False
'###                            
'###
'###    CREATED BY      :   Pratiksha           17/09/2021    
'###    MODIFIED BY     :   
'###    EXAMPLE         :   Fn_AW_TaskOperation("PanelExist","","","")
'###                            Fn_AW_TaskOperation("ButtonExist","Suspend","","")
'###							Fn_AW_TaskOperation("ButtonExist","Resume","","")
'###							Fn_AW_TaskOperation("EnterComments","Suspend", "Suspend task here","")
'###							Fn_AW_TaskOperation("Verify","", "","Suspended")
'###							Fn_AW_TaskOperation("Verify","", "","Started")
'#######################################################################################################

Function Fn_AW_TaskOperation(sAction,sButtonName,sDescription,sStatus)
	Dim objDesc,Objgrp,iCount,objButton,objPanel,objstatus,rowcount,colcount,cellval,rowNum
	GBL_FAILED_FUNCTION_NAME="Fn_AW_TaskOperation"
   	Fn_AW_TaskOperation = FALSE
   	Select Case sAction
		 Case "PanelExist"
	           Set objPanel = UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("Task Panel")
               Set objDesc = Description.Create
               objDesc("controltype").Value = "Group"
               Set Objgrp = UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").ChildObjects(objDesc)
               If Objgrp.Count <> 0 Then
                  For iCount = 0 To Objgrp.Count - 1
		             If Instr(Objgrp(iCount).getROProperty("name"),"Task Panel") Then
		                objPanel.SetTOProperty "path",Objgrp(iCount).GetROProperty("path")
			            objPanel.SetTOProperty "hwnd",Objgrp(iCount).GetROProperty("hwnd")
			            Exit For
		             End If
                  Next
               Else
		            Fn_AW_TaskOperation = FALSE
	           End If
	           bReturn = Fn_AWS_UI_ObjectExist("Fn_AW_TaskOperation", objPanel)
			   If  bReturn=True Then
			       Fn_AW_TaskOperation = True
			   Else			
	               Fn_AW_TaskOperation = False
			   End If
	
	     Case "ButtonExist"
	           Set objButton = UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAButton("Button")
		       objButton.SetTOProperty "name",sButtonName
		       bReturn = Fn_AWS_UI_ObjectExist("Fn_AW_TaskOperation", objButton)
			   If  bReturn=True Then
			       Fn_AW_TaskOperation = True
			   Else			
	               Fn_AW_TaskOperation = False
			   End If
			   
	     Case "EnterComments"
	           Set objDesc = Description.Create
	           objDesc("controltype").Value = "Edit"
	           Set Objgrp = UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").ChildObjects(objDesc)
	           If Objgrp.Count <> 0 Then
		         For iCount = 0 To Objgrp.Count - 1
			        If Instr(Objgrp(iCount).getROProperty("name"),"Comments") Then
				      UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("Task Panel").UIAEdit("Comments").SetTOProperty "name" ,Objgrp(iCount).getROProperty("name")
				      UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("Task Panel").UIAEdit("Comments").SetTOProperty "path" ,Objgrp(iCount).getROProperty("path")
				      UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("Task Panel").UIAEdit("Comments").SetTOProperty "hwnd" ,Objgrp(iCount).getROProperty("hwnd")
				      wait 2
				      UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("Task Panel").UIAEdit("Comments").Click
				      UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("Task Panel").UIAEdit("Comments").Type sDescription
				      wait 2
				      call Fn_KeyBoardOperation("SendKey", "{TAB}")
				      Exit For
			        End If
		         Next	
	           Else
		         Fn_AW_TaskOperation = FALSE
	           End If

	           If sButtonName <> "" Then
		          Set objButton = UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAButton("Button")
		          objButton.SetTOProperty "name",sButtonName
		          objButton.Click
		          Fn_AW_TaskOperation = True
	           Else
		         Fn_AW_TaskOperation = False
	           End If
	           
	     Case "Verify" 
	           Set objDesc = Description.Create
	           objDesc("controltype").Value = "DataGrid"
	           Set Objgrp = UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").ChildObjects(objDesc)
	           If Objgrp.Count <> 0 Then
		         For iCount = 0 To Objgrp.Count - 1
			        If Instr(Objgrp(iCount).getROProperty("path"),"Current and Completed Tasks") Then
				      UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIATable("DataGrid").SetTOProperty "name" ,Objgrp(iCount).getROProperty("name")
				      UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIATable("DataGrid").SetTOProperty "path" ,Objgrp(iCount).getROProperty("path")
				      UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIATable("DataGrid").SetTOProperty "hwnd" ,Objgrp(iCount).getROProperty("hwnd")
				      Exit For
			        End If
		         Next	
	           Else
		         Fn_AW_TaskOperation = FALSE
	           End If
	           
               Set objDesc = Description.Create
	           objDesc("controltype").Value = "Group"
	           Set Objgrp = UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").ChildObjects(objDesc)
	           If Objgrp.Count <> 0 Then
		         For iCount = 0 To Objgrp.Count - 1
			        If Instr(Objgrp(iCount).getROProperty("name"),sStatus) Then
				      UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIATable("DataGrid").UIAObject("Status").SetTOProperty "name" ,Objgrp(iCount).getROProperty("name")
				      UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIATable("DataGrid").UIAObject("Status").SetTOProperty "path" ,Objgrp(iCount).getROProperty("path")
				      UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIATable("DataGrid").UIAObject("Status").SetTOProperty "hwnd" ,Objgrp(iCount).getROProperty("hwnd")
				      Exit For
			        End If
		         Next	
	           Else
		         Fn_AW_TaskOperation = FALSE
	           End If
	           
	           Set objstatus = UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIATable("DataGrid").UIAObject("Status")
	           bReturn = Fn_AWS_UI_ObjectExist("Fn_AW_TaskOperation", objstatus)
			   If  bReturn=True Then
			       Fn_AW_TaskOperation = True
			   Else			
	               Fn_AW_TaskOperation = False
			   End If
		
	End  Select 
    Set objDesc = Nothing
	Set Objgrp = Nothing
	
End Function


'#######################################################################################################
'###
'###    Function Name		:	Fn_AW_MyTaskTable_Operation()
'###
'###    Description			:	Function Used to perform operation on Mytask tab table
'###
'###    Parameters		   :	1.strAction : Action Name
'###						    2. node Name
'###                            3.Column name
'###                            4.Dictionary parameter
'###
'###    Return Value	 : 	True Or False
'###                            
'###
'###    CREATED BY      :   Neha          16/09/2021    
'###    MODIFIED BY     :   
'###    EXAMPLE         :   Fn_AW_MyTaskTable_Operation("ClickCellElement","10.1.2 Simple Review No Profile : 030841/A;1-Item1","","")
'###############################################################################################################################################
Public Function Fn_AW_MyTaskTable_Operation(strAction,nodename,Columnname,DicInfo)
	Dim ObjMyTaskTblEle,RowNum
	
	GBL_FAILED_FUNCTION_NAME="Fn_AW_MyTaskTable_Operation"
   	Fn_AW_MyTaskTable_Operation = FALSE
   	Set ObjMyTaskTblEle = UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("Teamcenter_Inbox").UIATable("MyTaskTable").UIAObject("TableElement")
	
	Select Case strAction
		Case "VerifyCellDataExist"
			
			RowNum = UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("Teamcenter_Inbox").UIATable("MyTaskTable").GetROProperty("rowcount")			
			If RowNum <> "0" Then
			    ObjMyTaskTblEle.SetTOProperty "Name",nodename
			     Wait 2
			   	 If ObjMyTaskTblEle.Exist = True Then
			   	 	Fn_AW_MyTaskTable_Operation = True				
				 Else
				    Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " FAIL: Fail to check existance of node :"& nodename)
			        Fn_AW_MyTaskTable_Operation = False
			        Exit Function
			   	 End If
			Else
			   Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " FAIL: Row does not exist in table") 
			End If
			    
	  Case "ClickCellElement"				
			
			RowNum = UIAWindow("Teamcenter RAC (Eclipse)").UIATab("ActiveWorkspace").UIAObject("Teamcenter_Inbox").UIATable("MyTaskTable").GetROProperty("rowcount")			
			If RowNum <> "0" Then
			   ObjMyTaskTblEle.SetTOProperty "Name",nodename
			     Wait 2
			   	 If ObjMyTaskTblEle.Exist = True Then
			   	     ObjMyTaskTblEle.Click 
			   	     Wait 1
			   	     Fn_AW_MyTaskTable_Operation = True
				 Else
				    Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " FAIL: Fail to check existance of node :"& nodename)
			        Fn_AW_MyTaskTable_Operation = False
			        Exit Function
			   	 End If
			Else
			   Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " FAIL: Row does not exist in table")
               Fn_AW_MyTaskTable_Operation = False
      		   Exit Function				   
			End If
		
End Select
 
 Set ObjMyTaskTblEle = Nothing
End  Function
