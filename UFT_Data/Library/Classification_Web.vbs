Option Explicit

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~List~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'0.) Fn_SISW_ClassificationWeb_GetObject()
'1.) Fn_ClassificationWEb_SearchClassTreeTabOperations()
'2.) Fn_ClassificationWEb_HeirarchyTreeOperations()
'3.) Fn_ClassificationWEb_TabOpeartions()
'4.) Fn_ClassificationWEb_ClassifyObject()
'5.) Fn_ClassificationWEb_QuerryTabOperations()
'6.) Fn_ClassificationWEb_TableTabOperations()
'7.) Fn_ClassificationWEb_ClassifyWhileSettingAttributeValues()
'8.) Fn_ClassificationWEb_CreateNewICO()
'9.) Fn_ClassificationWEb_PropertyTabOperations()
'10.) Fn_ClassificationWEb_BrowseICO()
'11.) Fn_ClassificationWEb_BookmarkClass()
'12.) Fn_ClassificationWEb_SelectBookmarkedClass()
'13) Fn_SISW_ClassificationWeb_GetObject()
''**>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 			Function to get Object hierarchy  		- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''Function Name		:	Fn_SISW_ClassificationWeb_GetObject
'
''Description		  	 :  	Function to get Object hierarchy

''Parameters		   :	1. sObjectName : Object Handle name
								
''Return Value		   :  	Object \ Nothing
'
''Examples		     	:	Fn_SISW_Classification_GetObject("ClassificationApplet")

'History:                
'								Developer Name							Date				Rev. No.		Reviewer		Changes Done	
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'									Dipali k		 				6-July-2012				1.0					Prasanna
'-----------------------------------------------------------------------------------------------------------------------------------
'	Ashwini Kumar		 25-Sept-2013		2.0								Externalized object hierarchies
'-----------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_ClassificationWeb_GetObject(sObjectName)
	Dim sObjectXMLPath
	sObjectXMLPath = Fn_GetEnvValue("User", "AutomationDir") & "\TestData\AutomationXML\ObjectXML\Classification_Web.xml"
	Set Fn_SISW_ClassificationWeb_GetObject = Fn_SISW_Setup_GetObjectFromXML(sObjectXMLPath, sObjectName)
End Function
''***************************************************************************************************************************************************************************************
'***  
'***			Function Name : Fn_ClassificationWEb_SearchClassTreeTabOperations(sAction,sSearchType,sSearchClass,sVerifyClass,bDisplayResults,sButton,sInfo)
'***
'***			ParaMeters  : sAction -->> Valid Action name
'***						   		  sSearchType-->> Search Type to be selected
'***						  		  sSearchClass-->> Class To Be Searched
'***                          		  sVerifyClass-->> Class To be verified with ICM name after search is returned
'***						  		 bDisplayResults-->> Boolean Parameter to display results
'***                          	 	sButton-->> Button To Be clicked
'***                           		sInfo-->> For Future Use
'***   
'***   		  Return Value  : True / False
'***
'***  		 Function Calls : Fn_WriteLogFile()
'*** 
'***  		Developer : 	SHREYAS
'***
'***       Reviewer : Prasanna
'***
'***      Date : 10/05/2011
'***  
'***     How To Use : bReturn=Fn_ClassificationWEb_SearchClassTreeTabOperations("Search&Verify","Class Name","Storage_12019","ICM04; Storage_12019","true","Close","")
'***
'**>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>


Public function Fn_ClassificationWEb_SearchClassTreeTabOperations(sAction,sSearchType,sSearchClass,sVerifyClass,bDisplayResults,sButton,sInfo)
	GBL_FAILED_FUNCTION_NAME="Fn_ClassificationWEb_SearchClassTreeTabOperations"
   Dim sRows,sValue,iCount,WshShell,bFlag
   Set objSearch=Browser("Classification").Page("Classification")
   Set WshShell = CreateObject("WScript.Shell")
   Fn_ClassificationWEb_SearchClassTreeTabOperations=false

'first clear the Search type edit box

'			objSearch.WebEdit("ClassSearchType").click
'			WshShell.SendKeys "{DEL}"
'			wait(3)

			Select Case sAction


					 Case "Search&Verify"
					
					'enter the search type in the Edit box
								If sSearchType<>"" Then
									bFlag=false
'									bReturn=Fn_Web_UI_WebEdit_Set("Fn_ClassificationWEb_SearchClassTreeTabOperations", objSearch, "ClassSearchType", sValue)
										'objSearch.WebEdit("ClassSearchType").set sSearchType
										'wait(3)
'										
'										objSearch.WebEdit("ClassSearchType").Click
'										objSearch.WebEdit("ClassSearchType").Object.innertext = ""
'										
'										WshShell.SendKeys (sSearchType)
										objSearch.WebButton("CriteriaSelect").Click                                 '										
										'wait 3
										if Fn_Web_UI_ObjectExist("Fn_ClassificationWEb_SearchClassTreeTabOperations", objSearch.WebElement("Class NameClass IDAttribute")) then 
										'If  objSearch.WebElement("Class NameClass IDAttribute").Exist(5) Then											
													objSearch.WebElement("Attribute Name").SetTOProperty "innertext",sSearchType
													'wait 1
													objSearch.WebElement("Attribute Name").Click												
										End If										
										'wait 3
										bFlag=true
										If bFlag=true Then
											Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS : Successfully set the value " +sSearchType+ " in the ClassSearchType edit box")	
											Fn_ClassificationWEb_SearchClassTreeTabOperations=true
										Else
											Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL : Failed to set the value " +sSearchType+ " in the ClassSearchType edit box")	
											Fn_ClassificationWEb_SearchClassTreeTabOperations=False
											Set objSearch=nothing
											Exit function
										End If
								End If
									
									'enter the search class in the SearchText box
									If sSearchClass<>"" Then
										bFlag=false
'										bReturn=Fn_Web_UI_WebEdit_Set("Fn_ClassificationWEb_SearchClassTreeTabOperations", objSearch, "SearchText", sValue)
'										objSearch.WebEdit("SearchText").set ""
'										objSearch.WebEdit("SearchText").Click
'										objSearch.WebEdit("SearchText").Set sSearchClass
										objSearch.WebEdit("SearchText").Click									
										wait 1
										'objSearch.WebEdit("SearchText").Object.innertext = ""
										'objSearch.WebEdit("SearchText").set ""
										Call Fn_Web_UI_WebEdit_SetExt("Fn_ClassificationWEb_SearchClassTreeTabOperations", "Set",objSearch, "SearchText", "")
										'wait 1
										'WshShell.SendKeys (sSearchClass)
										objSearch.WebEdit("SearchText").Set sSearchClass
										Set WshShell = nothing
										bFlag=True
										If bFlag=true Then
											Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS : Successfully set the value " +sSearchClass+ " in the SearchText edit box")	
											Fn_ClassificationWEb_SearchClassTreeTabOperations=true
										Else
											Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL : Failed to set the value " +sSearchClass+ " in the SearchText edit box")	
											Fn_ClassificationWEb_SearchClassTreeTabOperations=False
											Set objSearch=nothing
											Exit function
										End If
									End If
									
									'now search for the class
									objSearch.Image("ClassSearch").Click
									
									
									'check the existence of "No class Found" Dialog
									If  Browser("Classification").Dialog("Information").Exist(10) Then
										Browser("Classification").Dialog("Information").Close
										Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL : The Class "+sSearchClass+" is not found")	
										Fn_ClassificationWEb_SearchClassTreeTabOperations=False
										Set objSearch=nothing
										Exit function
									Else
										Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS : Successfully set the value " +sSearchClass+ " in the SearchText edit box")	
										Fn_ClassificationWEb_SearchClassTreeTabOperations=true
									End If


								If bDisplayResults<>"" Then

										If cbool(bDisplayResults)=true Then
											If objSearch.WebButton("DisplayResult").Exist(3) Then
												objSearch.WebButton("DisplayResult").Click
												'wait(3)
											Else
													Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL : The Display Results Button is not present")		
													Fn_ClassificationWEb_SearchClassTreeTabOperations=False
													Set objSearch=nothing
													Exit function
											End If

												If Browser("Classification").Page("Classification").WebElement("ClassSearchResults").Exist(10) then
														Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS : Successfully invoked the DisplayResults dialog")	
														Fn_ClassificationWEb_SearchClassTreeTabOperations=true
												Else
														Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL : The Display Results Dialog is not present")		
														Fn_ClassificationWEb_SearchClassTreeTabOperations=False
														Set objSearch=nothing
														Exit function
													
												End If
											Else
													Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"Display Results dialog is not required")
										End If
								End If

									If sVerifyClass<>"" Then
											sRows=objSearch.WebElement("ClassSearchResults").WebList("ResultList").GetROProperty("items count")
											For iCounter=1 to sRows
														sValue=objSearch.WebElement("ClassSearchResults").WebList("ResultList").GetItem(iCounter)
														If lCase(sValue)=lCase(sVerifyClass) Then
																	Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS : Successfully verified that the class "+sVerifyClass+" is returned after the search")
																	Fn_ClassificationWEb_SearchClassTreeTabOperations=true
																	Exit For
														End If
														If cInt(sRows)=1 and cint(iCounter)=1 and lCase(sValue)=lCase(sVerifyClass) Then
																	Fn_ClassificationWEb_SearchClassTreeTabOperations=true
														elseIf cInt(sRows)>1 and cInt(sRows) =cInt(iCounter) Then
																	Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL : Failed to verify that the class "+sVerifyClass+" is returned after the search")		
																	Fn_ClassificationWEb_SearchClassTreeTabOperations=False
																	Set objSearch=nothing
																	Exit function
														End If
											Next
								End if

								If sButton<>"" then
									bFlag=false
									objSearch.WebElement("ClassSearchResults").WebButton(sButton).Click
									bFlag=True
										If bFlag=True Then
											Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS : Successfully clicked on the "+sButton+" button")	
											Fn_ClassificationWEb_SearchClassTreeTabOperations=true
										Else
											Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL : Failed to click on the "+sButton+" button")	
											Fn_ClassificationWEb_SearchClassTreeTabOperations=False
											Set objSearch=nothing
											Exit function
										End If
								End if

				End select 
				Set objSearch=nothing
End function



''**>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'***  
'***			Function Name : Fn_ClassificationWEb_HeirarchyTreeOperations(sAction,sNode,sCount,sID,sType,sInfo1,sInfo2)
'***
'***			ParaMeters  : 		 sAction -->> Valid Action name
'***						   		  sNode-->> Valid node name
'***						  		  sCount-->> Count to be verified
'***                          		  sID-->> Class ID to be verified
'***						  		 sType-->> Class Type to be verified
'***                          	 	sInfo1-->> For Future Use
'***                           		sInfo2-->> For Future Use
'***   
'***   		  Return Value  : True / False
'***
'***  		 Function Calls : Fn_WriteLogFile()
'*** 
'***  		Developer : 	SHREYAS
'***
'***       Reviewer : 	Prasanna
'***
'***      Date : 10/05/2011
'***  
'***     How To Use : bReturn=Fn_ClassificationWEb_HeirarchyTreeOperations("Exist","Storage_20217","","","","","")
'***				  bReturn=Fn_ClassificationWEb_HeirarchyTreeOperations("VerifyTableEntries","Unit Definition Class","19","","","","")
'***
'**>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Public function Fn_ClassificationWEb_HeirarchyTreeOperations(sAction,sNode,sCount,sID,sType,sInfo1,sInfo2)
	
	GBL_FAILED_FUNCTION_NAME="Fn_ClassificationWEb_HeirarchyTreeOperations"	
   Dim sRows,sValue,objTree,iCounter,sDetails,objImg,bFlagFound
   Set objTree=Browser("Classification").Page("Classification").WebTable("ClassTable")
   '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
   'Added synchronization to check existance of  webtable [ ClassTable ]
	For iCounter=0 to 2
		If objTree.Exist(10) Then
			'wait 2
			bFlagFound=True
			Exit for
		Else
			'Wait 2
			bFlagFound=False
		End If
	Next
	
	''""""Added Code to Handle Expanded "Classification Root" node--Dhananjay Niwal
		sValue=objTree.GetRowWithCellText("Classification Root")
   		Set objImg =objTree.ChildItem(sValue,2,"Image",0)
   		If objImg.GetROProperty("file name") = "plus.png" Then
			objImg.Click 1,1, micLeftBtn
		End if
   	
	If bFlagFound=False Then
		Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL : Class table not found")	
		Exit function
	End If
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
   Fn_ClassificationWEb_HeirarchyTreeOperations=false

	 	Select Case sAction
				'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				Case "Select"
						sValue=objTree.GetRowWithCellText(sNode)
						If sValue="" Then
								Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL : Failed to retrieve the row number of the node "+sNode)	
								Fn_ClassificationWEb_HeirarchyTreeOperations=False
								Set objTree=nothing
								Exit function
						Else
								Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS : Successfully retrieved the row number of the node "+sNode+" as "+cstr(sValue))
								Fn_ClassificationWEb_HeirarchyTreeOperations=true
						End If
						'objTree.ChildItem(sValue,1,"WebElement",0).Click
						objTree.ChildItem(sValue,0,"WebCheckBox",0).Set "ON"
				'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				Case "Deselect"
						sValue=objTree.GetRowWithCellText(sNode)
						If sValue="" Then
								Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL : Failed to retrieve the row number of the node "+sNode)	
								Fn_ClassificationWEb_HeirarchyTreeOperations=False
								Set objTree=nothing
								Exit function
						Else
								Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS : Successfully retrieved the row number of the node "+sNode+" as "+cstr(sValue))
								Fn_ClassificationWEb_HeirarchyTreeOperations=true
						End If
						objTree.ChildItem(sValue,0,"WebCheckBox",0).Set "OFF"
				'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				Case "Activate"

					'when we click (activate) the node,it by default opens in the "QuerryTab"
						sValue=objTree.GetRowWithCellText(sNode)
						If sValue="" Then
								Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL : Failed to retrieve the row number of the node "+sNode)	
								Fn_ClassificationWEb_HeirarchyTreeOperations=False
								Set objTree=nothing
								Exit function
						Else
								Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS : Successfully retrieved the row number of the node "+sNode+" as "+cstr(sValue))
								Fn_ClassificationWEb_HeirarchyTreeOperations=true
						End If
						objTree.ChildItem(sValue,2,"WebElement",2).Click

				Case "Exist"
						bFlagFound = false
						sRows=objTree.GetROProperty("rows")
						For iCounter=1 to sRows
							sValue=objTree.GetCellData(iCounter,2)
							If lCase(trim(sValue))=lCase(trim(sNode)) Then
								Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS : Successfully verified that the node "+sNode+" exists in the Table")
								Fn_ClassificationWEb_HeirarchyTreeOperations=true
								bFlagFound = true
								Exit for
							End If
						Next

						If bFlagFound = false Then
								Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL : Failed to verify that the node "+sNode+" exists in the Table")
								Fn_ClassificationWEb_HeirarchyTreeOperations=False
								Set objTree=nothing
								Exit function
						End If
						
						
			Case "VerifySelection"  'this case is used to check if a particular class has its adjecent checkbox checked after searching it

					If sNode<>"" Then
'						sValue= objTree.GetRowWithCellText(sNode)
'						Browser("Classification").Page("Classification").WebCheckBox("WebCheckBox").SetTOProperty "index",sValue-2  'sValue-2 is implemented as the checkbox index is differemt than class name adjecent to it and was observed as 2 less than the node index in several cases,hence implemented
'
'						'check if the checkbox is checked
'						sValue=Browser("Classification").Page("Classification").WebCheckBox("WebCheckBox").object.checked
	
'							sRowNo=Browser("Classification").Page("Classification").WebTable("ClassTable").GetRowWithCellText(sNode,2,1)
'							sRowNo=sRowNo-1
'							sColorValue=Browser("Classification").Page("Classification").WebTable("ClassTable").object.rows(sRowNo).currentStyle.backgroundcolor 
'	
'							If lCase(trim(scolorValue))=lCase(trim("#bdbebd")) Then 
'									sValue=true
'							End If
							sRowNo=Browser("Classification").Page("Classification").WebTable("ClassTable").GetRowWithCellText(sNode,2,1)
'							sRowNo=sRowNo-1
							Set scolorValue=Browser("Classification").Page("Classification").WebTable("ClassTable").ChildItem(sRowNo,0,"WebCheckBox",0)
							If Typename(scolorValue)<>"Nothing" Then
								If lcase(scolorValue.GetROProperty("checked"))="1" Then
									sValue=true
								Else
									sValue=false
								End If
							End If

							If cBool(sValue)=true Then
									Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS : Successfully verified that the checkbox is checked for the class ["+sNode+" ]")
									Fn_ClassificationWEb_HeirarchyTreeOperations=true
							Else
									Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL : Failed to verify that the checkbox is checked for the class ["+sNode+" ]")
									Fn_ClassificationWEb_HeirarchyTreeOperations=False
									Set objTree=nothing
									Exit function
							End If

					End If
					
				Case "VerifyTableEntries"  'this case will verify the values occuring in count,ID & Type Columns in the Class Table

				If sCount<>"" Then
					sValue=objTree.GetRowWithCellText(sNode)
					If sValue="" Then
							Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL : Failed to retrieve the row number of the node "+sNode)	
							Fn_ClassificationWEb_HeirarchyTreeOperations=False
							Set objTree=nothing
							Exit function
					Else
							Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS : Successfully retrieved the row number of the node "+sNode+" as "+cstr(sValue))
							Fn_ClassificationWEb_HeirarchyTreeOperations=true
					End If
					objTree.ChildItem(sValue,1,"WebElement",0).Click
					sDetails=Browser("Classification").Page("Classification").WebTable("ClassTable").GetCellData(sValue,3)
					If lCase(sCount)= lCase(sDetails) Then
							Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS : Successfully verified that the value ["+sCount+"] is present in the COunt Column Of the class table")
							Fn_ClassificationWEb_HeirarchyTreeOperations=true
					Else
							Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL : Failed to verify that the value ["+sCount+"] is present in the COunt Column Of the class table")
							Fn_ClassificationWEb_HeirarchyTreeOperations=False
							Set objTree=nothing
							Exit function
					End If

				End If

				'==================================================================================================================
				Case "Expand"

						sValue=objTree.GetRowWithCellText(sNode)
						If sValue="" Then
								Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL : Failed to retrieve the row number of the node "+sNode)	
								Fn_ClassificationWEb_HeirarchyTreeOperations=False
								Set objTree=nothing
								Exit function
						Else
								Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS : Successfully retrieved the row number of the node "+sNode+" as "+cstr(sValue))
								Fn_ClassificationWEb_HeirarchyTreeOperations=true
						End If

						Set objImg =objTree.ChildItem(sValue,2,"Image",0)

						If TypeName(objImg) <> "Nothing" Then
									If objImg.GetROProperty("file name") = "plus.png" Then
													objImg.Click 1,1, micLeftBtn
													Fn_ClassificationWEb_HeirarchyTreeOperations = True
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_ClassificationWEb_HeirarchyTreeOperations: node  ["+CStr(sNode)+"] expanded.")
									ElseIf objImg.GetROProperty("file name") = "minus.png" Then
													Fn_ClassificationWEb_HeirarchyTreeOperations = True
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_ClassificationWEb_HeirarchyTreeOperations: node  ["+CStr(sNode)+"] was already expanded.")
									Else
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_ClassificationWEb_HeirarchyTreeOperations: can not expand node  ["+CStr(sNode)+"].")
									End If				
	
						End If
			'=================================================================================================================

		End Select

		Set objTree=nothing

End function


'**>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'***  
'***			Function Name :  Fn_ClassificationWEb_TabOpeartions(sAction,sTabName,sDetails)
'***
'***			ParaMeters  : sAction -->> Valid Action name
'***						   		  sTabName-->> Tab To Be Activated
'***						  		  sDetails-->> For Future Use
'***   
'***   		   Return Value  : True / False
'***
'***  		  Function Calls : Fn_WriteLogFile()
'*** 
'***  		  Developer : 	SHREYAS
'***
'***         Reviewer : 		Prasanna
'*** 
'***         Date : 			10/05/2011
'***  
'***        How To Use : bReturn=Fn_ClassificationWEb_TabOpeartions("Set","Table","")
'***
'**>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Public function Fn_ClassificationWEb_TabOpeartions(sAction,sTabName,sDetails)

	GBL_FAILED_FUNCTION_NAME="Fn_ClassificationWEb_TabOpeartions"
	Dim objTab,bReturn
		Set objTab=Browser("Classification").Page("Classification")


		Fn_ClassificationWEb_TabOpeartions=false

			Select Case sAction
			
					Case "Set"
			
						If sTabName<>"" Then
							objTab.WebElement("Tree").SetTOProperty "innertext",sTabName
							bReturn=Fn_Web_UI_WebElement_Click("Fn_ClassificationWEb_TabOpeartions",objTab, "Tree", "", "", "")
							If bReturn=true Then
								Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS : Successfully activated the tab ["+sTabName+"]")
								Fn_ClassificationWEb_TabOpeartions=true
							Else
								Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL : Failed to activate the tab ["+sTabName+"]")
								Fn_ClassificationWEb_TabOpeartions=False
								Set objTab=nothing
								Exit function
							End If
						End If
				
			End Select

			Set objTab=nothing

End Function


''**>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'***  
'***			Function Name : Fn_ClassificationWEb_ClassifyObject(sAction,sNavTreeNode,sClassName,sInfo1,sInfo2)
'***
'***			ParaMeters  : 		 sAction -->> Valid Action name
'***						   		  sNavTreeNode-->> Valid NavTree node name
'***						  		  sClassName-->> Class name in which Item / Item Revision has to be classified
'***                          	 	sButton-->> To Click on Yes or No Button on the classify dialog
'***                           		sInfo-->> For Future Use
'***   
'***   		  Return Value  : True / False
'***
'***  		 Function Calls : Fn_WriteLogFile()
'*** 
'***  		Developer : 	SHREYAS
'***
'***       Reviewer : 	Prasanna
'***
'***      Date : 10/05/2011
'***  
'***     How To Use : bReturn=Fn_ClassificationWEb_ClassifyObject("Classify","Home:Newstuff:000024-itm","strClass","yes:save","Value1") (to click on save button)
'***				  bReturn = Fn_ClassificationWEb_ClassifyObject("Classify","Home:AutomatedTests:ClsWebClientSearchICOwith_64406:000147-ClsItm","Storage_22509","Yes:nosave","str&") (to NOT click on save button)
'***
'**>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Public function Fn_ClassificationWEb_ClassifyObject(sAction,sNavTreeNode,sClassName,sButton,sInfo)

   GBL_FAILED_FUNCTION_NAME="Fn_ClassificationWEb_ClassifyObject"	
   Dim objClassify,aValues,bReturn
   Dim sExpandNode,iCount

   Set objClassify=Browser("Classification").Page("Classification")


   Fn_ClassificationWEb_ClassifyObject=false

	Select Case sAction

		Case "Classify"

					'Select the NavTree Node
		
					bReturn= Fn_Web_NavTreeOperation("Select",sNavTreeNode)
					If bReturn=false Then
							Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL : Failed to select the NavTree node ["+sNavTreeNode+"]")
							Fn_ClassificationWEb_ClassifyObject=False
							Set objTab=nothing
							Exit function
					Else
							Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS : Successfully selected the NavTree node ["+sNavTreeNode+"]")
							Fn_ClassificationWEb_ClassifyObject=true
					End If
					wait(3)
		
					'perform menuOperation "Edit:Copy"
		
					bReturn=Fn_Web_MenuOperation("Select","Edit:Copy")
					If bReturn=false Then
							Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL : Failed to perform the operation [Edit:Copy]")
							Fn_ClassificationWEb_ClassifyObject=False
							Set objTab=nothing
							Exit function
					Else
							Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS : Successfully performed the operation [Edit:Copy]")
							Fn_ClassificationWEb_ClassifyObject=true
					End If
					wait(3)
		
					'load the Classification Perspective
		
					bReturn=Fn_Web_SetPerspective("Classification")
					If bReturn=false Then
							Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL : Failed to load the Classification Perspective")
							Fn_ClassificationWEb_ClassifyObject=False
							Set objTab=nothing
							Exit function
					Else
							Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS : Successfully loaded the Classification Perspective")
							Fn_ClassificationWEb_ClassifyObject=true
					End If
					wait(3)
		
					'select the class in classification
			'==============================Added to select Collapsed Class  =======================================================================
					
					If  instr(1, sClassName,":") > 0 Then
								sExpandNode=split(sClassName,":",-1,1) 
		
								For iCount=0 to Ubound(sExpandNode)
										bReturn=Fn_ClassificationWEb_HeirarchyTreeOperations("Expand",sExpandNode(iCount),"","","","","")
										If bReturn=false Then
												Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL : Failed to Expand the class ["+sExpandNode(iCount)+"]")
												Fn_ClassificationWEb_ClassifyObject=False
												Set objTab=nothing
												Exit function
										End If
										sClassName=sExpandNode(iCount)
								Next
		
								bReturn=Fn_ClassificationWEb_HeirarchyTreeOperations("Select",sClassName,"","","","","")
								If bReturn=false Then
										Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL : Failed to Select the class ["+sClassName+"]")
										Fn_ClassificationWEb_ClassifyObject=False
										Set objTab=nothing
										Exit function
								Else
										Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS : Successfully selected the class ["+sClassName+"]")
										Fn_ClassificationWEb_ClassifyObject=true
								End If		

					Else
		
								bReturn=Fn_ClassificationWEb_HeirarchyTreeOperations("Select",sClassName,"","","","","")
								If bReturn=false Then
										Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL : Failed to select the class ["+sClassName+"]")
										Fn_ClassificationWEb_ClassifyObject=False
										Set objTab=nothing
										Exit function
								Else
										Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS : Successfully selected the class ["+sClassName+"]")
										Fn_ClassificationWEb_ClassifyObject=true
								End If
								wait(3)

					End If

			'===========================================================================================================
	
					'perform menuOperation "Edit:Paste"
		
					bReturn=Fn_Web_MenuOperation("Select","Edit:Paste")
					If bReturn=false Then
							Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL : Failed to perform the operation [Edit:Paste]")
							Fn_ClassificationWEb_ClassifyObject=False
							Set objTab=nothing
							Exit function
					Else
							Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS : Successfully performed the operation [Edit:Paste]")
							Fn_ClassificationWEb_ClassifyObject=true
					End If
					wait(3)

					If sButton<>"" Then
							If instr(1,sButton,":")>0 Then
										
									aValues=split(sButton,":",-1,1)
				
									'click on yes or button of the "Classify object" dialog
									
										If lCase(aValues(0))="yes" Then
											objClassify.WebElement("ClassifyObject").WebButton("Yes").Click
											wait(3)
'											If sInfo<>"" Then
'												 If objClassify.WebList("AttributeID").Exist  Then  '' Added By Vidya 28/3/2012
'													 Wait(5)
'												objClassify.WebList("AttributeID").Select sInfo
'											Else 
'													objClassify.WebTable("PropertiesTable2").WebEdit("AttributeValue").Set sInfo
'												End If
'												
'											End If


											If sInfo<>"" Then
												 If objClassify.WebTable("PropertiesTable2").WebEdit("AttributeValue").Exist  Then  '' Added By Vidya 28/3/2012
													 objClassify.WebTable("PropertiesTable2").WebEdit("AttributeValue").Set sInfo
													 Wait(5)
											Else 
													If objClassify.WebList("AttributeID").Exist  Then  '' Added By Vidya 28/3/2012
												 Wait(5)
												objClassify.WebList("AttributeID").Select sInfo
												End If
											End If
										End If
										'click on the "Save" button in the properties tab
										If lCase(aValues(1))="save" Then
										objClassify.WebButton("Save").Click
										End If
										Fn_ClassificationWEb_ClassifyObject=true
									End if
							
								Else
										If lCase(sButton)="yes" Then
														objClassify.WebElement("ClassifyObject").WebButton("Yes").Click
														If sInfo<>"" Then
															objClassify.WebTable("PropertiesTable2").WebEdit("AttributeValue").Set sInfo
														End If
														'click on the "Save" button in the properties tab
														objClassify.WebButton("Save").Click
														Fn_ClassificationWEb_ClassifyObject=true
														wait(3)
														Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS : Successfully classified the Item Path ["+sNavTreeNode+"] in the Class [ "+sClassName+"]")
														Exit Function
										Elseif lCase(sButton)="no" Then
										objClassify.WebElement("ClassifyObject").WebButton("No").Click
										End If
													
												
													
				End If
		End if

					wait(3)

					Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS : Successfully classified the Item Path ["+sNavTreeNode+"] in the Class [ "+sClassName+"]")
									
		End Select

End Function


''**>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'***  
'***			Function Name : Fn_ClassificationWEb_QuerryTabOperations(sAction,sClass,bQueryTab,bOidSearch,sObjectId,bRevisionRule,sSetValue,bLockUnlockValue,sInfo1,sInfo2)
'***
'***			ParaMeters  : sAction -->> Valid Action name
'***						   		  sClass-->> Class To Be Searched
'***						  		  bQueryTab-->> boolean parameter to activate or to not activate the query tab
'***                          		  bOidSearch-->>  boolean parameter to perform OidSearch
'***						  		 sObjectId-->> Object Id to be verified
'***                          	 	bRevisionRule-->> boolean parameter to deploy revision rule
'***                           		sSetValue-->> Set the LOV value
'***     							bLockUnlockValue-->> boolean parameter to lock or unlock the LOV
'***     							sInfo1-->> For Future Use
'***     							sInfo1-->> For Future Use
'***     
'***     
'***   
'***   		  Return Value  : True / False
'***
'***  		 Function Calls : Fn_WriteLogFile()
'*** 
'***  		Developer : 	SHREYAS
'***
'***       Reviewer : Prasanna
'***
'***      Date : 12/05/2011
'***  
'***     How To Use : bReturn=Fn_ClassificationWEb_QuerryTabOperations("VerifyObjectId","Reamer30367_2_40905","true","","Integer_42637","","","","","")
'***
'**>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Public function Fn_ClassificationWEb_QuerryTabOperations(sAction,sClass,bQueryTab,bOidSearch,sObjectId,bRevisionRule,sSetValue,bLockUnlockValue,sInfo1,sButton)

	GBL_FAILED_FUNCTION_NAME="Fn_ClassificationWEb_QuerryTabOperations"
   Fn_ClassificationWEb_QuerryTabOperations=false 

   Dim objQuery,bReturn,aValues

   Set objQuery=Browser("Classification").Page("Classification").WebTable("QuerryTable")



			If cbool(bQueryTab)=true Then
			   'select the class
			   If sClass<>"" Then
'				   Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL : Class to be selected is not mentioned")
'					Fn_ClassificationWEb_QuerryTabOperations=False
'					Set objQuery=nothing
'					Exit function
					
					bReturn=Fn_ClassificationWEb_HeirarchyTreeOperations("VerifySelection",sClass,"","","","","") 		'------------Added by Snehal S. to check selection of storage class	-	18-Oct-2013
					If bReturn=false Then

						bReturn=Fn_ClassificationWEb_HeirarchyTreeOperations("Select",sClass,"","","","","")
						If bReturn=false Then
								Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL : Failed to select the class ["+sClass+"]")
								Fn_ClassificationWEb_QuerryTabOperations=False
								Set objQuery=nothing
								Exit function
						Else
								Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS : Successfully selected the class ["+sClass+"]")
								Fn_ClassificationWEb_QuerryTabOperations=True
						End If

					End If
			   End If
								   	
		   'activate the querry tab
				 bReturn=Fn_ClassificationWEb_TabOpeartions("Set","Query","")
				If bReturn=false Then
					Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL : Failed to activate the query tab")
					Fn_ClassificationWEb_QuerryTabOperations=False
					Set objQuery=nothing
					Exit function
				Else
					Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS : Successfully activated the query tab")
					Fn_ClassificationWEb_QuerryTabOperations=True
				End If
			Else
				Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"Query tab is already active")
			End if

			Fn_Web_ReadyStatusSync(2)
					Select Case sAction
					
						Case "VerifyObjectId"
									
					
							If sObjectId<>"" Then
								objQuery.WebElement("Value").SetTOProperty "innertext",sObjectId
								If  objQuery.WebElement("Value").Exist(5) Then
										Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS : Successfully verified that the object ["+sObjectId+"] exists in the Querry Table")
										Fn_ClassificationWEb_QuerryTabOperations=True
								Else
										Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL : Failed to verify that the object ["+sObjectId+"] exists in the Querry Table")
										Fn_ClassificationWEb_QuerryTabOperations=False
										Set objQuery=nothing
										Exit function
								End If
							End If

					Case "ObjectIdSearch"

						If sObjectId<>"" Then

								'click on clear button
								Browser("Classification").Page("Classification").WebButton("Clear").Click
								'Browser("Classification").Page("Classification").WebElement("Clear").Click
								Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS : Successfully cleared all the fields")
								wait(2)

								'set the search value in the OID Search Field
								'objQuery.WebEdit("ObjectId").set sObjectId
								Call Fn_Web_UI_WebEdit_SetExt("Fn_ClassificationWEb_QuerryTabOperations", "Set",objQuery, "ObjectId", sObjectId)
								Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS : Successfully set the value ["+sObjectId+"] in the ObjectId edit box")
								wait(2)

								If sInfo1<>"" OR sInfo1="ObjectID_SearchIcon" Then
											'click on the search icon on left of Object Id field
											Browser("Classification").Page("Classification").Image("ClassSearch").SetTOProperty "alt", "OID Search"
											Browser("Classification").Page("Classification").Image("ClassSearch").Click
											Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS : Successfully clicked on the search icon on left of Object Id field")
											Fn_ClassificationWEb_QuerryTabOperations=True
											wait(1)
								Else
											'click on the search button
											Browser("Classification").Page("Classification").WebButton("Search").Click
											'Browser("Classification").Page("Classification").WebElement("Search").Click
											Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS : Successfully clicked on the search button")
											Fn_ClassificationWEb_QuerryTabOperations=True
								End If
					End If

				'===================================================================================================
					Case "AttrValueSearch"    ' Addded by Pooja 8-June 2011

						If sObjectId<>"" Then

								'click on clear button
								Browser("Classification").Page("Classification").WebButton("Clear").Click
								'Browser("Classification").Page("Classification").WebElement("Clear").Click
								Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS : Successfully cleared all the fields")
								wait(2)

								'set the search value in the OID Search Field
								objQuery.WebEdit("SetLOV").set sObjectId
								Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS : Successfully set the value ["+sObjectId+"] in the Attr Value edit box")
								wait(2)

								'click on the search button
								Browser("Classification").Page("Classification").WebButton("Search").Click
								'Browser("Classification").Page("Classification").WebElement("Search").Click
								Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS : Successfully clicked on the search button")
								Fn_ClassificationWEb_QuerryTabOperations=True

								'the remaining verifications can be done using the Table Operations Function in the table tab
					End If


			Case "SetValues"

	                	'click on the button to select the value to be set
					If sInfo1<>""  Then
						aValues=split(sInfo1,":",-1,1)
						Browser("Classification").Page("Classification").WebButton("AddLOV").SetTOProperty "index",aValues(0)
						Browser("Classification").Page("Classification").WebButton("AddLOV").Click 0,0,micLeftBtn
						Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS : Successfully clicked on the AddLOV button")
							Browser("Classification").Page("Classification").WebElement("LOVName").SetTOProperty "innertext",aValues(1)
						Browser("Classification").Page("Classification").WebElement("LOVName").Click 0,0,micLeftBtn
						Fn_ClassificationWEb_QuerryTabOperations=True 
						Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS : Successfully selected the value ["+aValues(1)+"]")
					End If


					'now select the value
					If sButton<>"" Then
						Browser("Classification").Page("Classification").WebButton(sButton).Click 0,0,micLeftBtn
						Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS : Successfully clicked on the ["+ sButton+"] button")
					End If

					

			
			End Select

   Set objQuery=nothing
 
End Function


''**>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'***  
'***			Function Name :  Fn_ClassificationWEb_TableTabOperations(sAction,bTableTab,sObjectId,sButton,sInfo1,sInfo2)
'***
'***			ParaMeters  : sAction -->> Valid Action name
'***						  		  bTableTab-->> boolean parameter to activate or to not activate the Table tab
'***						  		 sObjectId-->> Object Id to be verified
'***								 sButton-->> Valid Button Name to be clicked
'***     							sInfo1-->> For Future Use
'***     							sInfo1-->> For Future Use
'***     
'***     
'***   
'***   		  Return Value  : True / False
'***
'***  		 Function Calls : Fn_WriteLogFile()
'*** 
'***  		Developer : 	SHREYAS
'***
'***       Reviewer : Prasanna
'***
'***      Date : 19/05/2011
'***  
'***     How To Use : bReturn=Fn_ClassificationWEb_TableTabOperations("VerifyObjectId","true","000024","","","")
'***								msgbox bReturn,vbinformation,"Function Return Value"
'***
'**>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Public Function Fn_ClassificationWEb_TableTabOperations(sAction,bTableTab,sObjectId,sButton,sInfo1,sInfo2)

	GBL_FAILED_FUNCTION_NAME="Fn_ClassificationWEb_TableTabOperations"
   Fn_ClassificationWEb_TableTabOperations=false 

   Dim objTable,bReturn

   Set objTable=Browser("Classification").Page("Classification").WebTable("TableResults")



			If cbool(bTableTab)=true Then
		
		   'activate the Table tab if required
				 bReturn=Fn_ClassificationWEb_TabOpeartions("Set","Table","")
				If bReturn=false Then
					Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL : Failed to activate the Table tab")
					  Fn_ClassificationWEb_TableTabOperations=False
					Set objTable=nothing
					Exit function
				Else
					Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS : Successfully activated the Table tab")
					  Fn_ClassificationWEb_TableTabOperations=True
					  wait(3)
				End If
			Else
				Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"Table tab is already active")
			End if

			Fn_Web_ReadyStatusSync(3)
			wait(3)

					Select Case sAction
					
						Case "VerifyObjectId"
							
					
							If sObjectId<>"" Then
								objTable.WebElement("Value").SetTOProperty "innertext",sObjectId
								If  objTable.WebElement("Value").Exist(5) Then
										Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS : Successfully verified that the object ["+sObjectId+"] exists in the TableResults Table")
										  Fn_ClassificationWEb_TableTabOperations=True
								Else
										Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL : Failed to verify that the object ["+sObjectId+"] exists in the TableResults Table")
										  Fn_ClassificationWEb_TableTabOperations=False
										Set objTable=nothing
										Exit function
								End If
							End If

			
					End Select

   Set objTable=nothing
 
End Function




'''**>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
''***  
''***			Function Name :  Fn_ClassificationWEb_ClassifyWhileSettingAttributeValues(sAction,sObjectClassName,sIndex,sValue,sButton,sInfo1,sInfo2)
''***
''***			ParaMeters  : 	 sAction -->> Valid Action name
''***						   		  		  sObjectClassName-->> Class Name of the object
''***						  		 	      sIndex-->> Index of the Object
''***									     sValue-->> IValue to be Set / Selected in the object
''***                          	 	         sButton-->> To Click on Save or Cancel Button
''***                           		    sInfo1-->> For Future Use
''***                           		    sInfo1-->> For Future Use
''***   
''***   		  Return Value  : True / False
''***
''***  		 Function Calls : Fn_WriteLogFile()
''*** 
''***  		Developer : 	SHREYAS
''***
''***       Reviewer : 	Prasanna
''***
''***      Date : 19/05/2011
''***  
''***     How To Use : bReturn= Fn_ClassificationWEb_ClassifyWhileSettingAttributeValues("Set","WebList","0","1 ch1","","","")
''***
''**>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Public function Fn_ClassificationWEb_ClassifyWhileSettingAttributeValues(sAction,sObjectClassName,sIndex,sValue,sButton,sInfo1,sInfo2)

	GBL_FAILED_FUNCTION_NAME="Fn_ClassificationWEb_ClassifyWhileSettingAttributeValues"
	Dim objVal, bReturn, iCount, arrValue, iNoOfEditBox, objPropertyName
	
	Fn_ClassificationWEb_ClassifyWhileSettingAttributeValues=false
	
	Select Case sAction
			
				Case "Set"

								If  sObjectClassName="" Then
										Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL : The  Object Class Name should not be blank")
										Fn_ClassificationWEb_ClassifyWhileSettingAttributeValues=False
										Exit function
								End If
			
								If lcase(sObjectClassName)="webedit" Then
												Set objVal=Browser("Classification").Page("Classification").WebTable("PropertiesTable2").WebEdit("AttributeValue")
												objVal.SetTOProperty "index",sIndex
												objVal.highlight
												Wait(1)
												'now set the value in the Attribute edit box of choice
												If sValue<>"" Then
														objVal.Set sValue
														Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS : Successfully set the value  ["+sValue+"] in the Attribute Edit Box")
														Fn_ClassificationWEb_ClassifyWhileSettingAttributeValues=true
												Else
														Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL : The  Attribute value to be set should not be blank")
														Fn_ClassificationWEb_ClassifyWhileSettingAttributeValues=False
														Set objVal=nothing
														Exit function
												End If

								Elseif lcase(sObjectClassName)="webeditmultiple" Then  'Added by Prasanna for TC ClsMultiUnitWeb_CreateClsItemUsingEnhAttriInputFeat
												
												Set objSelectType1=description.Create()
												objSelectType1("micClass").value = "WebEdit"
												Set  intNoOfObjects1 = Browser("Classification").Page("Classification").WebTable("PropertiesTable2").ChildObjects(objSelectType1)
												For iCounter = 0 to intNoOfObjects1.Count
													If cint(iCounter) = cint(sIndex) Then
																If sValue<>"" Then
																			intNoOfObjects1(iCounter).Set sValue
																			Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS : Successfully set the value  ["+sValue+"] in the Attribute Edit Box")
																			Fn_ClassificationWEb_ClassifyWhileSettingAttributeValues=true
																Else
																			Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL : The  Attribute value to be set should not be blank")
																			Fn_ClassificationWEb_ClassifyWhileSettingAttributeValues=False
																			Set objVal=nothing
																			Exit function
																End If
																
													End If

												Next
								Elseif lcase(sObjectClassName)="weblist" Then
												Set objVal=Browser("Classification").Page("Classification").WebTable("PropertiesTable2").WebList("KeyLOV")
												objVal.SetTOProperty "index",sIndex
												objVal.highlight
			
													'now set the value in the Attribute list of choice
												If sValue<>"" Then
														objVal.Select sValue
														Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS : Successfully set the value  ["+sValue+"] from the Attribute List")
														Fn_ClassificationWEb_ClassifyWhileSettingAttributeValues=true
												Else
														Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL : The  Attribute value to be set should not be blank")
														Fn_ClassificationWEb_ClassifyWhileSettingAttributeValues=False
														Set objVal=nothing
														Exit function
												End If
			
								End If
				'[TC1122-2016010600-14_Jan_2016-VivekA-Maintenance] - Added Case to Verify Array Size of Attributes - By Snehal S - Copied from TC1015
				Case "VerifyAttributeArraySize"
								Select Case sObjectClassName
										Case "WebEdit"
												For iCount = 0 To UBound(sValue)
													arrValue = Split(sValue(iCount),":")
													iNoOfEditBox = Browser("Classification").Page("Classification").WebTable("PropertiesTable2").ChildItemCount(iCount+2,0,"WebEdit")
													objPropertyName = Browser("Classification").Page("Classification").WebTable("PropertiesTable2").ChildItem(iCount+2,0,"WebElement",1).getroproperty("innertext")
													If Trim(arrValue(0)) = Trim(objPropertyName) AND CInt(arrValue(1)) =  CInt(iNoOfEditBox) Then
														Fn_ClassificationWEb_ClassifyWhileSettingAttributeValues = True
													Else 
														Fn_ClassificationWEb_ClassifyWhileSettingAttributeValues = False
														Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL : The  Attribute value for WebEdit is not Verified ")
														Exit For
													End If
												Next
								End Select
								
		End Select
		wait(3)
		If sButton<>"" Then
			Browser("Classification").Page("Classification").WebButton(sButton).Click
			Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS : Successfully Clicked on the button  ["+sButton+"] ")
		Else
			Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"No Button Needs to be clicked")
		End If

		Set objVal=nothing

End Function


''**>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
''***  
''***			Function Name :  Fn_ClassificationWEb_CreateNewICO(sAction,sClass,sICOId,sButton,sInfo1,sInfo2)
''***
''***			ParaMeters  : 		 sAction -->> Valid Action name
''***						   		  			sClass-->> Class name in which new ICO has to be created
''***						  		 			 sICOId-->> Valid ICO Id
''***                          	 				sButton-->> To Click on OK or Cancel Button
''***                           				sInfo1-->> For Future Use
''***                           				sInfo1-->> For Future Use
''***   
''***   		  Return Value  : True / False
''***
''***  		 Function Calls : Fn_WriteLogFile()
''*** 
''***  		Developer : 	SHREYAS
''***
''***       Reviewer : 	Prasanna
''***
''***      Date : 20/05/2011
''***  
''***     How To Use : bReturn= Fn_ClassificationWEb_CreateNewICO("Create","Storage_31042","05042011","OK","","")
''***  								msgbox bReturn,vbinformation,"Function Return Value"
''***
''**>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Public function Fn_ClassificationWEb_CreateNewICO(sAction,sClass,sICOId,sButton,sInfo1,sInfo2)

	GBL_FAILED_FUNCTION_NAME="Fn_ClassificationWEb_CreateNewICO"
	Fn_ClassificationWEb_CreateNewICO=false

   Dim bReturn,objICO

   Set objICO=Browser("Classification").Page("Classification")

	 	Select Case sAction
	
			Case "Create"

				If sICOId="" Then
					Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL : ICOId should not be blank")
					Fn_ClassificationWEb_CreateNewICO=False
					Set objICO=nothing
					Exit function
			 End if

				If sClass="" Then
					Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL : Clas Name should not be blank")
					Fn_ClassificationWEb_CreateNewICO=False
					Set objICO=nothing
					Exit function
			 End if

	   				'select the class in classification
		
					bReturn=Fn_ClassificationWEb_HeirarchyTreeOperations("Select",sClass,"","","","","")
					If bReturn=false Then
							Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL : Failed to select the class ["+sClass+"]")
							Fn_ClassificationWEb_CreateNewICO=False
							Set objICO=nothing
							Exit function
					Else
							Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS : Successfully selected the class ["+sClass+"]")
							Fn_ClassificationWEb_CreateNewICO=true
					End If
					objICO.Sync

				 'invoke the new classification object dialog
					bReturn=Fn_Web_MenuOperation("Select","New:Classification Object")
					If bReturn=false Then
							Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL : Failed to perform the operation [New:Classification Object]")
							Fn_ClassificationWEb_CreateNewICO=False
							Set objICO=nothing
							Exit function
					Else
							Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS : Successfully performed the operation [New:Classification Object]")
							Fn_ClassificationWEb_CreateNewICO=true
					End If
	
					objICO.Sync
	
	
					If objICO.WebElement("NewICO").Exist Then
	
							'set the value in the ID field
							objICO.WebEdit("ID").set sICOId
							Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS : Successfully set the value  ["+sICOId+"] in the ICOId field")
		
							'click on the required button
							If sButton<>"" Then
								objICO.WebButton(sButton).Click
								Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS : Successfully clicked on the button  ["+sButton+"] ")
								Fn_ClassificationWEb_CreateNewICO=true
								objICO.Sync
							Else
								Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"No Button needs to be clicked")
								Fn_ClassificationWEb_CreateNewICO=true
							End If
							Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS : Successfully Created New ICO with ID generated as  ["+sICOId+"] ")
				End If
			
		End Select
		Set objICO=nothing
				
End Function


'**>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
''***  
''***			Function Name :   Fn_ClassificationWEb_PropertyTabOperations(sAction,sClass,bPropertyTab,sObjectId,sInfo1,sInfo2,sButton)
''***
''***			ParaMeters  : 		 sAction -->> Valid Action name
''***						   		  			sClass-->> Class name in which new ICO has to be created
''***						  		 			 bPropertyTab-->> Boolean parameter to activate or to not activate the property tab
''***                          	 				sObjectId-->> Object Id to be verified
''***                           				sInfo1-->> For Future Use
''***                           				sInfo1-->> For Future Use
''***                           				sButton-->> Button To Be clicked
''***   
''***   		  Return Value  : True / False
''***
''***  		 Function Calls : Fn_WriteLogFile()
''*** 
''***  		Developer : 	SHREYAS
''***
''***       Reviewer : 	Prasanna
''***
''***      Date : 26/05/2011
''***  
''***     How To Use : bReturn=Fn_ClassificationWEb_PropertyTabOperations("VerifyObjectId","","true",Environment.Value("KeyLovAttrName"),"","","")
''***  								msgbox bReturn,vbinformation,"Function Return Value"
''***
''**>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Public function Fn_ClassificationWEb_PropertyTabOperations(sAction,sClass,bPropertyTab,sObjectId,sInfo1,sInfo2,sButton)

	GBL_FAILED_FUNCTION_NAME="Fn_ClassificationWEb_PropertyTabOperations"
   Fn_ClassificationWEb_PropertyTabOperations=false 

   Dim objProperty,bReturn

   Set objProperty=Browser("Classification").Page("Classification").WebTable("PropertiesTable2")



			If cbool(bPropertyTab)=true Then
		
		   'activate the Property tab
				 bReturn=Fn_ClassificationWEb_TabOpeartions("Set","Properties","")
				If bReturn=false Then
					Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL : Failed to activate the Properties tab")
					Fn_ClassificationWEb_PropertyTabOperations=False
					Set objProperty=nothing
					Exit function
				Else
					Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS : Successfully activated the Properties tab")
					Fn_ClassificationWEb_PropertyTabOperations=True
				End If
			Else
				Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"Properties tab is already active")
			End if

					Select Case sAction
					
						Case "VerifyObjectId"
									   'select the class
								   If sClass<>"" Then
									
										bReturn=Fn_ClassificationWEb_HeirarchyTreeOperations("Select",sClass,"","","","","")
										If bReturn=false Then
											Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL : Failed to select the class ["+sClass+"]")
											Fn_ClassificationWEb_PropertyTabOperations=False
											Set objProperty=nothing
											Exit function
										Else
											Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS : Successfully selected the class ["+sClass+"]")
											Fn_ClassificationWEb_PropertyTabOperations=True
										End If
								   End If
					
							If sObjectId<>"" Then
								objProperty.WebElement("Value").SetTOProperty "innertext",sObjectId
								If  objProperty.WebElement("Value").Exist(5) = false Then
									objProperty.WebElement("Value").SetTOProperty "html tag","SPAN"	
								End if		
								If  objProperty.WebElement("Value").Exist(5) Then
										If objProperty.WebElement("Value").getROProperty("height") <= 0 then
												Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL : Failed to verify that the object ["+sObjectId+"] exists in the Querry Table")
												Fn_ClassificationWEb_PropertyTabOperations=False												
												Set objProperty=nothing
												Exit function	
										end if
										Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS : Successfully verified that the object ["+sObjectId+"] exists in the Querry Table")
										Fn_ClassificationWEb_PropertyTabOperations=True
								Else
										Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL : Failed to verify that the object ["+sObjectId+"] exists in the Querry Table")
										Fn_ClassificationWEb_PropertyTabOperations=False
										Set objProperty=nothing
										Exit function
								End If
							End If

							' Code added to verify the Attribute Value after saving.(Added by Sneha) for TestCase (ClsMultiUnitWeb_ClsItemInputUnitByTyping).
							If sInfo1<>"" Then
								objProperty.WebElement("AttrValue").SetTOProperty "innertext",sInfo1
								If  objProperty.WebElement("AttrValue").Exist(5) = false Then
									objProperty.WebElement("AttrValue").SetTOProperty "html tag","SPAN"
								End if	
								If  objProperty.WebElement("AttrValue").Exist(5) Then
										Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS : Successfully verified that the object ["+sInfo1+"] exists in the Querry Table")
										Fn_ClassificationWEb_PropertyTabOperations=True
								Else
										Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL : Failed to verify that the object ["+sInfo1+"] exists in the Querry Table")
										Fn_ClassificationWEb_PropertyTabOperations=False
										Set objProperty=nothing
										Exit function
								End If
						End If
						
						Case "VerifyItemId"

								If sObjectId<>"" Then
								objProperty.WebElement("IdValue").SetTOProperty "innertext",sObjectId

								If  objProperty.WebElement("IdValue").Exist(5) Then
										Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS : Successfully verified that the ItemID ["+sObjectId+"] exists in the Property Table")
										Fn_ClassificationWEb_PropertyTabOperations=True
								Else

									Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL : Failed to verify that the ItemID ["+sObjectId+"] exists in the Property Table")
									Fn_ClassificationWEb_PropertyTabOperations=False
									Set objProperty=nothing
									Exit function
	
								End If
							End If

                         Case "Verifydefaultvalue"					'' Added By Vidya.

										If sObjectId<>"" Then
										objProperty.WebElement("IdValue").SetTOProperty "innertext",sObjectId
		
										If  objProperty.WebEdit("AttributeValue").Exist(5) Then
												If objProperty.WebEdit("AttributeValue").CheckProperty("value",sInfo1)= True Then
													Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS : Successfully verified the default value of the attribute ["+sObjectId+"]  in the Property Table")
													Fn_ClassificationWEb_PropertyTabOperations=True
										Else
													Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL : Failed to verify the default value of the attribute ["+sObjectId+"]  in the Property Table")
													Fn_ClassificationWEb_PropertyTabOperations=False
													Set objProperty=nothing
													Exit function

												End If
										End If
									End If
							
						Case "VerifyLOVListValue"

						sValue= objProperty.WebList("KeyLOV").GetROProperty("items Count")

						For iCount=1 to sValue
							sDetails=objProperty.WebList("KeyLOV").GetItem(iCount)
							If lCase(sObjectId)=lCase(sDetails) Then
								Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS : Successfully verified that the ItemID ["+sObjectId+"] exists in the KeyLOV List")
								Fn_ClassificationWEb_PropertyTabOperations=True
								Exit for
							End If
						
							If cint(iCount)=cint(sValue) Then
									Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL : Failed to verify that the ItemID ["+sObjectId+"] exists in the KeyLOV List")
									Fn_ClassificationWEb_PropertyTabOperations=False
									Set objProperty=nothing
									Exit function
							End If
						
						Next

					End Select
			End Function

''''**>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'''***  
'''***			Function Name :  Fn_ClassificationWEb_BrowseICO's(sAction,sBrowseCount,sTab,sInfo1,sInfo2)
'''***
'''***			ParaMeters  : 		  sAction -->> Valid Action name
'''***						   		  			   sBrowseCount-->> Number of ICO's to be browsed
'''***						  		 			  sTab-->> Valid Tab Name
'''***                           				 sInfo1-->> For Future Use
'''***                           				sInfo1-->> For Future Use
'''***   
'''***   		  Return Value  : True / False
'''***
'''***  		 Function Calls : Fn_WriteLogFile()
'''*** 
'''***  		Developer : 	SHREYAS
'''***
'''***       Reviewer : 	Prasanna
'''***
'''***      Date : 		13/06/2011
'''***  
'''***     How To Use : bReturn= Fn_ClassificationWEb_BrowseICOs("Next","3","Properties","","")
'''***  								msgbox bReturn,vbinformation,"Function Return Value"
'''***
'''**>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Public function Fn_ClassificationWEb_BrowseICOs(sAction,sBrowseCount,sTab,sInfo1,sInfo2)

	GBL_FAILED_FUNCTION_NAME="Fn_ClassificationWEb_BrowseICOs"
	Dim sRows,sValue,iCount,objBrowse,bReturn
	 Set objBrowse=Browser("Classification").Page("Classification")
	Fn_ClassificationWEb_BrowseICOs=false


		'Set the tab in which Ico's Have to be browsed
		If sTab<>"" Then
				bReturn=Fn_ClassificationWEb_TabOpeartions("Set",sTab,"")
							If bReturn=false Then
									Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL : Failed to set the tab ["+sTab+"]")
									Fn_ClassificationWEb_BrowseICOs=False
									Set objBrowse=nothing
									Exit function
							Else
		                          Call  Fn_Web_ReadyStatusSync(2)
								  Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS : Successfully set the tab ["+sTab+"]")
									Fn_ClassificationWEb_BrowseICOs=true
							End If
		End If


		Select Case sAction

				Case "Next"
					If  sBrowseCount="" Then
						Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL : The Browse Count cannot be blank")
						Fn_ClassificationWEb_BrowseICOs=False
						Set objBrowse=nothing
						Exit function
					Else
						For iCount=1 to sBrowseCount
							objBrowse.Image("Next").Click
							Wait(1)
						Next
						
							Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS : Successfully browsed through the ICO's "+sBrowseCount+" times Using button"+sAction )
							Fn_ClassificationWEb_BrowseICOs=true
						
					End If

					Case "Previous"
					If  sBrowseCount="" Then
						Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL : The Browse Count cannot be blank")
						Fn_ClassificationWEb_BrowseICOs=False
						Set objBrowse=nothing
						Exit function
					Else
						For iCount=1 to sBrowseCount
							objBrowse.Image("Previous").Click
							Wait(1)
						Next
						
							Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS : Successfully browsed through the ICO's "+sBrowseCount+" times Using button"+sAction )
							Fn_ClassificationWEb_BrowseICOs=true
					
					End If

					Case "First"
					If  sBrowseCount="" Then
						Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL : The Browse Count cannot be blank")
						Fn_ClassificationWEb_BrowseICOs=False
						Set objBrowse=nothing
						Exit function
					Else
						For iCount=1 to sBrowseCount
							objBrowse.Image("First").Click
							Wait(1)
						Next
						
							Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS : Successfully browsed through the ICO's "+sBrowseCount+" times Using button"+sAction )
							Fn_ClassificationWEb_BrowseICOs=true
						
					End If

					Case "Last"
					If  sBrowseCount="" Then
						Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL : The Browse Count cannot be blank")
						Fn_ClassificationWEb_BrowseICOs=False
						Set objBrowse=nothing
						Exit function
					Else
						For iCount=1 to sBrowseCount
							objBrowse.Image("Last").Click
							Wait(1)
						Next
						
							Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS : Successfully browsed through the ICO's "+sBrowseCount+" times Using button"+sAction )
							Fn_ClassificationWEb_BrowseICOs=true
						
					End If

		End Select

End Function

'''**>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
''***  
''***			Function Name : Fn_ClassificationWEb_BookmarkClass(sAction,sClassName,sBookmarkName,sFolder,sInfo1,snfo2)
''***
''***			ParaMeters  : 		 sAction -->> Valid Action name
''***						   		  sClassName-->> Class name in which Item / Item Revision has to be classified
''***						  		  sBookmarkName-->> Valid Bookmark Name
''***						  		  sFolder-->> Valid Folder Name in Which the bookmark has to be added
''***                          	 	sInfo-->> For Future Use
''***                           		sInfo-->> For Future Use
''***   
''***   		  Return Value  : True / False
''***
''***  		 Function Calls : Fn_WriteLogFile()
''*** 
''***  		Developer : 	SHREYAS
''***
''***       Reviewer : 	Prasanna
''***
''***      Date : 16/06/2011
''***  
''***     How To Use : bReturn =Fn_ClassificationWEb_BookmarkClass("Bookmark","Shields_51089","Shields_51089","","","")
''***  
''**>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Public function Fn_ClassificationWEb_BookmarkClass(sAction,sClassName,sBookmarkName,sFolder,sInfo1,snfo2)
	GBL_FAILED_FUNCTION_NAME="Fn_ClassificationWEb_BookmarkClass"
   Dim bReturn,objBookmark
   Dim locX,locY,WshShell,obj1
   Set objBookmark=Browser("Classification").Dialog("Add a Favorite")

   Select Case sAction
	 
		Case "Bookmark"

			'first select the class in the heirarchy tree
			If sClassName<>"" Then
						bReturn=Fn_ClassificationWEb_HeirarchyTreeOperations("Select",sClassName,"","","","","")
						If bReturn=false Then
								Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL : Failed to select the class ["+sClassName+"]")
								Fn_ClassificationWEb_BookmarkClass=False
								Set objBookmark=nothing
								Exit function
						Else
								Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS : Successfully selected the class ["+sClassName+"]")
								Fn_ClassificationWEb_BookmarkClass=true
						End If
						wait(2)
		End If

		'invoke the "Add a Favorite" dialog
'			bReturn=Fn_Web_MenuOperation("Select","Tools:Bookmark Class")
'					If bReturn=false Then
'							Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL : Failed to perform the operation [Tools:Bookmark Class]")
'							Fn_ClassificationWEb_BookmarkClass=False
'							Set objBookmark=nothing
'							Exit function
'					Else
'							Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS : Successfully performed the operation [Tools:Bookmark Class")
'							Fn_ClassificationWEb_BookmarkClass=true
'					End If

				'Added by Prasanna 05 July 2011
					
				  Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_MenuOperation",Browser("TeamcenterWeb").Link("MenuLink"),"text","Tools")'
				  Call Fn_Web_UI_Link_Click("Fn_Web_MenuOperation", Browser("TeamcenterWeb"), "MenuLink", "","","")
				   wait 2
					locX=Browser("Classification").Page("Classification").Link("BookmarkClass").GetROProperty("abs_x")
					locY=Browser("Classification").Page("Classification").Link("BookmarkClass").GetROProperty("abs_y")
					set obj1 = CreateObject("Mercury.DeviceReplay")
					obj1.MouseMove locX,locY
					wait 2

					set WshShell = CreateObject("WScript.Shell")
					WshShell.SendKeys "{DOWN}"
					wait 2
					WshShell.SendKeys "{DOWN}"
					wait 2
					'WshShell.SendKeys "~"
					WshShell.SendKeys "{ENTER}"

					wait 2
					

					If objBookmark.Exist Then

						'set the name of the bookmark
							If sBookmarkName<>"" Then
								Browser("Classification").Dialog("Add a Favorite").WinEdit("Name:").Set sBookmarkName
								Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS : Successfully set the bookmark name as  ["+sBookmarkName+"]")
							End If

							If sFolder<>"" Then
'								to be implemented
							End If

				Else
							Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL : The 'Add a Favorite' dialog does not exist")
							Fn_ClassificationWEb_BookmarkClass=False
							Set objBookmark=nothing
							Exit function
	
				End If

					'click on "add" button
					Browser("Classification").Dialog("Add a Favorite").WinButton("Add").Click 0,0,micLeftBtn
					Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS : Successfully clicked on the  'Add' Button")
					Fn_ClassificationWEb_BookmarkClass = true
   End Select
   Set objBookmark=nothing

End Function

''''**>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
''***  
''***			Function Name : Fn_ClassificationWEb_SelectBookmarkedClass(sAction,sBookmarkName,sInfo1,snfo2)
''***
''***			ParaMeters  : 		 sAction -->> Valid Action name
''***						  		  			 sBookmarkName-->> Valid Bookmark Name to be selected
''***                          	 				sInfo1-->> For Future Use
''***                           				sInfo1-->> For Future Use
''***   
''***   		  Return Value  : True / False
''***
''***  		 Function Calls : Fn_WriteLogFile()
''*** 
''***  		Developer : 	SHREYAS
''***
''***       Reviewer : 	Prasanna
''***
''***      Date : 20/06/2011
''***  
''***     How To Use : bReturn =Fn_ClassificationWEb_SelectBookmarkedClass("VerifyBookmark","classification","","")
''***   							 bReturn =Fn_ClassificationWEb_SelectBookmarkedClass("SelectBookmark","classification","","")
''***  
''**>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Public function Fn_ClassificationWEb_SelectBookmarkedClass(sAction,sBookmarkName,sInfo1,snfo2)
	GBL_FAILED_FUNCTION_NAME="Fn_ClassificationWEb_SelectBookmarkedClass"
   Dim bReturn,objBookmark,sValue,i,sDetails
  
  If Window("Windows Internet Explorer").WinToolbar("FavoritesToolBar").Exist(5) Then
  	 Set objBookmark=Window("Windows Internet Explorer")
  Else 
	 Set objBookmark=Browser("Teamcenter_Web") 
  End If
 
   Fn_ClassificationWEb_SelectBookmarkedClass=false

   Select Case sAction

	 	Case "SelectBookmark"

					If sBookmarkName<>"" Then

						'click on the favorites toolbar & select the bookmark
                           'objBookmark.WinToolbar("FavoritesToolBar").Click 0,0,micLeftBtn
						   objBookmark.WinToolbar("FavoritesToolBar").Press("Favorites")
						If objBookmark.WinTreeView("FavoritesTreeView").Exist then
							objBookmark.WinTreeView("FavoritesTreeView").Expand "Favorites Bar"
							sValue=objBookmark.WinTreeView("FavoritesTreeView").GetItemsCount
							For i=0 to sValue-1
											sDetails=objBookmark.WinTreeView("FavoritesTreeView").GetItem(i)
											If lCase(sBookmarkName)=lCase(sDetails) Then
												objBookmark.WinTreeView("FavoritesTreeView").Select i
												Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS : Successfully verified that  the bookmark ["+sBookmarkName+"] exists in Favorites")
												Fn_ClassificationWEb_SelectBookmarkedClass=true
												Exit for
											End If
											If i=sValue-1 Then
												Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL :Failed to verify that  the bookmark ["+sBookmarkName+"] exists in Favorites")
												Fn_ClassificationWEb_SelectBookmarkedClass=False
												Set objBookmark=nothing
												Exit function
											End If
							Next
					End if
				End If

				Set objBookmark=nothing
			Case "VerifyBookmark"

					If sBookmarkName<>"" Then

						'click on the favorites toolbar & select the bookmark
                           'objBookmark.WinToolbar("FavoritesToolBar").Click 0,0,micLeftBtn
						   objBookmark.WinToolbar("FavoritesToolBar").Press("Favorites")
						If objBookmark.WinTreeView("FavoritesTreeView").Exist then
							objBookmark.WinTreeView("FavoritesTreeView").Expand "Favorites Bar"
							sValue=objBookmark.WinTreeView("FavoritesTreeView").GetItemsCount
							For i=0 to sValue-1
											sDetails=objBookmark.WinTreeView("FavoritesTreeView").GetItem(i)
											If lCase(sBookmarkName)=lCase(sDetails) Then
												Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS : Successfully verified that  the bookmark ["+sBookmarkName+"] exists in Favorites")
												Fn_ClassificationWEb_SelectBookmarkedClass=true
												Exit for
											End If
											If i=sValue-1 Then
												Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL :Failed to verify that  the bookmark ["+sBookmarkName+"] exists in Favorites")
												Fn_ClassificationWEb_SelectBookmarkedClass=False
												Set objBookmark=nothing
												Exit function
											End If
							Next
					End if
							'objBookmark.WinToolbar("FavoritesToolBar").Click 0,0,micLeftBtn
							 objBookmark.WinToolbar("FavoritesToolBar").Press("Favorites")
			End If

   End Select
   Set objBookmark=nothing
End Function
