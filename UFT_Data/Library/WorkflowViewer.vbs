'*****'*********************************************************	Function List		***********************************************************************
'1.  Fn_SISW_WorkflowViewer_GetObject()
'2.  Fn_WrkflwViewr_ProcessTree_NodeOperations()
'3.  Fn_WrkflwViewr_TskAttachment_PanelOperations()
'4.  Fn_WorkflowViewer_Attributes() 
'5.  Fn_WrkflwViewr_PanelOperations() 
'6.	 Fn_WrkflwViewr_ChangeToDesignModeConf()
'7.  Fn_WrkflwViewr_AuditLogOperations()
'8. Fn_WrkflwViewr_TaskActions()		'The function performs Actions on Workflow Viewer Nodes. Action can be Complete,Start,Resume...
'9. Fn_WrkflwViewr_VerifyToolTipText()	'Function used to verify tool tip text of Process in Workflow Viewer
'****************************************    Function to get Object hierarchy ***************************************
'
''Function Name		 	:	Fn_SISW_WorkflowViewer_GetObject
'
''Description		    :  	Function to get Object hierarchy

''Parameters		    :	1. sObjectName : Object Handle name
								
''Return Value		    :  	Object \ Nothing
'
''Examples		     	:	Fn_SISW_WorkflowViewer_GetObject("Promote Action Comments")

'History:
'	Developer Name			Date			Rev. No.		Reviewer		      Changes Done	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Pritam Shikare		   22-June-2012		1.0				Nilesh Gadekar
'	Ashok kakade		   07-June-2012		1.0				Prasanna B.
'-----------------------------------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------------------------------
'	Anurag Khera		 25-Sept-2013		2.0								Externalized object hierarchies
'-----------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_WorkflowViewer_GetObject(sObjectName)
	Dim sObjectXMLPath
	sObjectXMLPath = Fn_GetEnvValue("User", "AutomationDir") & "\TestData\AutomationXML\ObjectXML\WorkflowViewer.xml"
	Set Fn_SISW_WorkflowViewer_GetObject = Fn_SISW_Setup_GetObjectFromXML(sObjectXMLPath, sObjectName)
End Function


'*********************************************************  Function do Operation on Workflow Viewer Process Tree *********************************************************************

'Function Name		:					Fn_WrkflwViewr_ProcessTree_NodeOperations

'Description			 :		 		    Action  performed :-
'																	1. Node Select																	
'																	2. Node Expand
'																	3. Node Collapse																	
'																    4.Exist

'Parameters			   :	 			1. sAction: Action to be performed
'												2.sNodeName: Fully qulified tree Path (delimiter as ':') [multiple node are separated by "," ] 
' 												   

'Return Value		   : 			 True/False

'Pre-requisite			:		 	 Process tree should be displayed.

'Examples				:			 Fn_WrkflwViewr_ProcessTree_NodeOperations("Select","AutoValidateWrkFlw:New Do Task 1")

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										     Prasanna				12-Aug-2010	       1.0
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_WrkflwViewr_ProcessTree_NodeOperations(sAction,sNodeName)
	GBL_FAILED_FUNCTION_NAME="Fn_WrkflwViewr_ProcessTree_NodeOperations"
   On Error Resume Next

   Dim arrNodeList,iItemCount,iCounter,sTreeItem,arrNode,iOuterCount,aMenuList
   Dim objProcessTree, objContext, intCount, StrMenu

   Set objProcessTree =  JavaWindow("WorkflowViewerWindow").JavaWindow("WEmbeddedFrame").JavaTree("ProcessTree")

	If InStr(1, sAction,":", 1) > 0 Then
		arrNode = Split(sAction, ":", -1, 1)
		sAction = arrNode(0)
		StrMenu = arrNode(1)
	End If

   If objProcessTree.Exist(5) Then
	        Select Case sAction
						Case "Select"                   		                            
										objProcessTree.Select sNodeName
										If Err.Number < 0 Then
												Fn_WrkflwViewr_ProcessTree_NodeOperations = False
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to select Node  " + sNodeName + " of Process Tree." )	
												Set objProcessTree = Nothing
												Exit Function 
										Else
												Fn_WrkflwViewr_ProcessTree_NodeOperations = True
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully selected  Node  " + sNodeName + " of Process Tree.")	
										End If  

						Case  "Expand"									
									objProcessTree.Expand sNodeName
										If Err.Number < 0 Then
												Fn_WrkflwViewr_ProcessTree_NodeOperations = False
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to expand node   " + sNodeName + " of Process Tree." )	
												Set objProcessTree = Nothing
												Exit Function 
										Else
												Fn_WrkflwViewr_ProcessTree_NodeOperations = True
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully expanded node  " + sNodeName  + " of Process Tree.")	
										End If
						Case  "ExpandBelow"									
									objProcessTree.Select sNodeName
										If Err.Number < 0 Then
												Fn_WrkflwViewr_ProcessTree_NodeOperations = False
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to expand node   " + sNodeName + " of Process Tree." )	
												Set objProcessTree = Nothing
												Exit Function 
										Else
												Call Fn_menuOperation("Select",Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("PSE_Menu"), "ViewExpandBelow"))
												Call Fn_Button_Click("Fn_WrkflwViewr_ProcessTree_NodeOperations", JavaWindow("WorkflowViewerWindow").JavaWindow("WEmbeddedFrame").JavaDialog("ExpandBelow"), "Yes")
												Fn_WrkflwViewr_ProcessTree_NodeOperations = True
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully expanded node  " + sNodeName  + " of Process Tree.")	
												Call Fn_ReadyStatusSync(1)
										End If				

						Case  "Collapse"										
										objProcessTree.Collapse sNodeName
										If Err.Number < 0 Then
											Fn_WrkflwViewr_ProcessTree_NodeOperations = False
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Collapse node   " + sNodeName + " of Process Tree." )	
											Set objProcessTree = Nothing
											Exit Function 
										Else
											Fn_WrkflwViewr_ProcessTree_NodeOperations = True
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Collapsed node  " + sNodeName  + " of Process Tree.")	
										End If

						Case "Exist"
			                            iItemCount = objProcessTree.GetROProperty( "items count")										
										For iCounter=0 To (iItemCount-1)
											sTreeItem = objProcessTree.GetItem(iCounter)
											
											If Trim (Lcase(sTreeItem)) = Trim(Lcase(sNodeName)) Then
												Fn_WrkflwViewr_ProcessTree_NodeOperations = True
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully found node " + sNodeName + " of Process Tree." )	
												Exit For
											End If
										Next 	
							
										If  Cint(iCounter) = Cint (iItemCount) Then
											Fn_WrkflwViewr_ProcessTree_NodeOperations = False
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to  Find node " + sNodeName + " of Process Tree." )	
											Set objProcessTree = Nothing
											Exit Function 
										End If

							Case "PopupMenuSelect"
										Set objContext = Fn_UI_ObjectCreate( "Fn_WrkflwViewr_ProcessTree_NodeOperations", JavaWindow("WorkflowViewerWindow"))
										'Build the Popup menu to be selected
										aMenuList = split(StrMenu, ":",-1,1)
										intCount = Ubound(aMenuList)
										'Select node
										Call Fn_ReadyStatusSync(3)
										Call Fn_WrkflwViewr_ProcessTree_NodeOperations("Select", sNodeName)
										Call Fn_ReadyStatusSync(3)
										'Open context menu
										Call Fn_UI_JavaTree_OpenContextMenu("Fn_WrkflwViewr_ProcessTree_NodeOperations",objContext.JavaWindow("WEmbeddedFrame"),"ProcessTree",sNodeName)
										Call Fn_ReadyStatusSync(3)
										
										'Select Menu action
										Select Case intCount
												Case "0"
													 StrMenu = JavaWindow("WorkflowViewerWindow").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0))
												Case "1"
													StrMenu = JavaWindow("WorkflowViewerWindow").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1))
												Case "2"
													StrMenu = JavaWindow("WorkflowViewerWindow").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1),aMenuList(2))
												Case Else
													Fn_WrkflwViewr_ProcessTree_NodeOperations = FALSE
													Exit Function
										End Select
										JavaWindow("WorkflowViewerWindow").WinMenu("ContextMenu").Select StrMenu	
										Call Fn_ReadyStatusSync(3)										
										Fn_WrkflwViewr_ProcessTree_NodeOperations = TRUE
						Case else
										Fn_WrkflwViewr_ProcessTree_NodeOperations = False
			End Select
   End if

Set objContext = Nothing
Set objProcessTree = Nothing

End Function 

'*********************************************************  Function do Operation on Workflow Viewer Attachment Panel TC Tree *********************************************************************

'Function Name		:					Fn_WrkflwViewr_TskAttachment_PanelOperations

'Description			 :		 		    Action  performed :-
'																	1. Node Select																	
'																	2. Node Expand
'																	3. Node Collapse																	
'																    4.Exist

'Parameters			   :	 			1. sAction: Action to be performed
'												2.sNodeName: Fully qulified tree Path (delimiter as ':') [multiple node are separated by "," ] 
' 												   

'Return Value		   : 			 True/False

'Pre-requisite			:		 	  Process tree node is selected from the Workflow Viewer module.

'Examples				:			 Fn_WrkflwViewr_TskAttachment_PanelOperations("Exist","New Do Task 1:Targets:004337/A;1-mmm")

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										     Prasanna				13-Aug-2010	       1.0
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function Fn_WrkflwViewr_TskAttachment_PanelOperations(sAction,sNodeName)
			GBL_FAILED_FUNCTION_NAME="Fn_WrkflwViewr_TskAttachment_PanelOperations"
		  On Error Resume Next
		
		  Dim arrNodeList,iItemCount,iCounter,sTreeItem,arrNode,iOuterCount,aMenuList, iCounter2, objDialog
		  Dim objTCTree,objAttachmentPanel,bReturn

		 If   JavaWindow("WorkflowViewerWindow").JavaWindow("WEmbeddedFrame").Exist(2) = True Then
			 Set objDialog = JavaWindow("WorkflowViewerWindow").JavaWindow("WEmbeddedFrame")
						objDialog.javacheckbox("attachmentpanel").Click 1,1,"LEFT"
                        If objDialog.JavaDialog("Attachments").Exist(5)=False Then   'Added by Nilesh Gadekat  for build change Tc 10 0606Build
								bReturn=objDialog.JavaStaticText("Attachments").GetROProperty("displayed")
								If bReturn=0 Then
											objDialog.JavaStaticText("Attachments").SetTOProperty "Index",1
											bReturn=objDialog.JavaStaticText("Attachments").GetROProperty("displayed")
											If  bReturn=0 Then
													Fn_WrkflwViewr_TskAttachment_PanelOperations = False
													Exit Function			
											End If			
								End If
						End If
		Else
			Set objDialog = JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow")
						objDialog.javacheckbox("AttachmentPanel").set "ON"
		 End If
		  
		  If Err.Number < 0 Then
					Fn_WrkflwViewr_TskAttachment_PanelOperations = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Click on Attachment Checkbox.")
					Set objDialog = Nothing
					Exit Function 
		  Else
					Fn_WrkflwViewr_TskAttachment_PanelOperations = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Clicked on Attachment Checkbox.")	
		 End If
   		Wait 2
		If objDialog.JavaDialog("Attachments").Exist(2)=False	Then		'Added by Nilesh Gadekat  for build change Tc 10 0606Build
			 objDialog.JavaStaticText("Attachments").DblClick 0,0
			  If Err.Number < 0 Then
						Fn_WrkflwViewr_TskAttachment_PanelOperations = False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Click on Attachment Panel header.")
						Set objDialog = Nothing
						Exit Function 
			  Else
						Fn_WrkflwViewr_TskAttachment_PanelOperations = True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Clicked on Attachment Panel header.")	
			 End If
		End If

   			Wait 3
			objDialog.RefreshObject
			Set objAttachmentPanel =  objDialog.JavaDialog("Attachments")
			If objDialog.JavaDialog("Attachments").Exist(5) then
		
				 Set objTCTree = objDialog.JavaDialog("Attachments").JavaTree("TCTree")	
				 Select Case sAction
							 Case "Copy"
										objTCTree.Select sNodeName
										If Err.Number < 0 Then
												Fn_WrkflwViewr_TskAttachment_PanelOperations = False
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to select Node  " + sNodeName + " of TC Tree." )	
												Set objTCTree = Nothing
												Set objAttachmentPanel = Nothing
												Set objDialog = Nothing
												Exit Function 
										Else
												Fn_WrkflwViewr_TskAttachment_PanelOperations = True
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully selected  Node  " + sNodeName + " of TC Tree.")	
										End If  
		
										'Click on Copy Button of Attachment Panel
										objAttachmentPanel.JavaButton("Copy").Click micLeftBtn
										If Err.Number < 0 Then
												Fn_WrkflwViewr_TskAttachment_PanelOperations = False
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Click on Copy Button of Attachments Panel." )	
												Set objTCTree = Nothing
												Set objAttachmentPanel = Nothing
												Set objDialog = Nothing
												Exit Function 
										Else
												Fn_WrkflwViewr_TskAttachment_PanelOperations = True
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Clicked on Copy Button of Attachments Panel")	
										End If  		
											
						   Case "Cut"
										objTCTree.Select sNodeName
										If Err.Number < 0 Then
												Fn_WrkflwViewr_TskAttachment_PanelOperations = False
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to select Node  " + sNodeName + " of TC Tree." )	
												Set objTCTree = Nothing
												Set objAttachmentPanel = Nothing
												Set objDialog = Nothing
												Exit Function 
										Else
												Fn_WrkflwViewr_TskAttachment_PanelOperations = True
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully selected  Node  " + sNodeName + " of TC Tree.")	
										End If  
		
										'Click on Cut Button of Attachment Panel
										objAttachmentPanel.JavaButton("Cut").Click micLeftBtn
										If Err.Number < 0 Then
												Fn_WrkflwViewr_TskAttachment_PanelOperations = False
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Click on Cut Button of Attachments Panel." )	
												Set objTCTree = Nothing
												Set objAttachmentPanel = Nothing
												Set objDialog = Nothing
												Exit Function 
										Else
												Fn_WrkflwViewr_TskAttachment_PanelOperations = True
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Clicked on Cut Button of Attachments Panel")	
										End If  			
		
						 Case "Paste"
										objTCTree.Select sNodeName
										If Err.Number < 0 Then
												Fn_WrkflwViewr_TskAttachment_PanelOperations = False
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to select Node  " + sNodeName + " of TC Tree." )	
												Set objTCTree = Nothing
												Set objAttachmentPanel = Nothing
												Set objDialog = Nothing
												Exit Function 
										Else
												Fn_WrkflwViewr_TskAttachment_PanelOperations = True
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully selected  Node  " + sNodeName + " of TC Tree.")	
										End If  
		
										'Click on Paste Button of Attachment Panel
										objAttachmentPanel.JavaButton("Paste").Click micLeftBtn
										If Err.Number < 0 Then
												Fn_WrkflwViewr_TskAttachment_PanelOperations = False
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Click on Paste Button of Attachments Panel." )	
												Set objTCTree = Nothing
												Set objAttachmentPanel = Nothing
												Set objDialog = Nothing
												Exit Function 
										Else
												Fn_WrkflwViewr_TskAttachment_PanelOperations = True
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Clicked on Paste Button of Attachments Panel")	
										End If  
													
						 Case "Open"
										objTCTree.Select sNodeName
										If Err.Number < 0 Then
												Fn_WrkflwViewr_TskAttachment_PanelOperations = False
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to select Node  " + sNodeName + " of TC Tree." )	
												Set objTCTree = Nothing
												Set objAttachmentPanel = Nothing
												Set objDialog = Nothing
												Exit Function 
										Else
												Fn_WrkflwViewr_TskAttachment_PanelOperations = True
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully selected  Node  " + sNodeName + " of TC Tree.")	
										End If  
		
										'Click on Open Button of Attachment Panel
										objAttachmentPanel.JavaButton("Open").Click micLeftBtn
										If Err.Number < 0 Then
												Fn_WrkflwViewr_TskAttachment_PanelOperations = False
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Click on Open Button of Attachments Panel." )	
												Set objTCTree = Nothing
												Set objAttachmentPanel = Nothing
												Set objDialog = Nothing
												Exit Function 
										Else
												Fn_WrkflwViewr_TskAttachment_PanelOperations = True
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Clicked on Open Button of Attachments Panel")	
										End If 
		
						Case  "Exist"
										 iItemCount = objTCTree.GetROProperty("items count")
										 For iCounter=0 To (iItemCount-1)
												sTreeItem = objTCTree.GetItem(iCounter)
												aMenuList = Split(sNodeName, ":", -1, 1)
												iOuterCount = 0
												For iCounter2 = 0 To (iItemCount - iCounter - 1)
													If Trim (Lcase(sTreeItem)) =  aMenuList(iOuterCount) Then
															objTCTree.Expand Trim(Lcase(sTreeItem))
															If Err.Number < 0 Then
																  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Expand Node ["+aMenuList(iOuterCount)+"] in TC Tree.")
																  Fn_WrkflwViewr_TskAttachment_PanelOperations = False
																   Set objDialog = Nothing
																  Exit Function
															Else
																  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Expand Node ["+aMenuList(iOuterCount)+"] in TC Tree.")
																  Call Fn_ReadyStatusSync(1)
																   Wait(1)
																  iOuterCount = iOuterCount + 1
															End If
													End If
												Next
												If Trim (Lcase(sTreeItem)) = Trim(Lcase(sNodeName)) Then
													Fn_WrkflwViewr_TskAttachment_PanelOperations = True
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully found node " + sNodeName + " of TC Tree." )	
													Exit For
												End If
										Next 	
							
										If  Cint(iCounter) = Cint (iItemCount) Then
											Fn_WrkflwViewr_TskAttachment_PanelOperations = False
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to  Find node " + sNodeName + " of TC Tree." )	
											objAttachmentPanel.Close
											Set objTCTree = Nothing
											Set objAttachmentPanel = Nothing
											Set objDialog = Nothing
											Exit Function 
										End If

						Case  "Expand"
										objTCTree.Expand sNodeName
										If Err.Number < 0 Then
											Fn_WrkflwViewr_TskAttachment_PanelOperations = False
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to expand node   " + sNodeName + "of TC Tree." )	
											Set objTCTree = Nothing
											Set objDialog = Nothing
											Exit Function 
										Else
											Fn_WrkflwViewr_TskAttachment_PanelOperations = True
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully expand node  " + sNodeName  + "of TC Tree.")	
										End If
	
						Case Else
										Fn_WrkflwViewr_TskAttachment_PanelOperations = False
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Perform Action " + sAction )	
										Set objTCTree = Nothing
										Set objAttachmentPanel = Nothing
										Set objDialog = Nothing
										Exit Function 
						
				 End Select
				 objAttachmentPanel.Close
				 If Err.Number < 0 Then
						Fn_WrkflwViewr_TskAttachment_PanelOperations = False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Close Attachments Panel." )	
						Set objTCTree = Nothing
						Set objAttachmentPanel = Nothing
						Set objDialog = Nothing
						Exit Function 
				 Else
						Fn_WrkflwViewr_TskAttachment_PanelOperations = True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Closed Attachments Panel.")	
				 End If 				 
		 Else
				 Fn_WrkflwViewr_TskAttachment_PanelOperations = False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Perform Action")
		 End if
		 Set objDialog = Nothing

End Function

'*********************************************************  Function perform the worklist process view attributes operation *********************************************************************
'Function Name  :   Fn_WorkflowViewer_Attributes
'
'Description    :        Workflow Viewer Attributes Operation
' 
'Parameters      :     sAction: 							Add/Remove/Modify/Verify
'           				 dicProcessViewAttributes: 	 Refer DictionaryDeclaration.vbs for the defination & keys included
' 
'Return Value     :   True/False
'
'Examples    :      
'								dicProcessViewAttributes.RemoveAll								
'								dicProcessViewAttributes.Add("ProcessTree") = "AutoRevFailPath:New Review Task 1:select-signoff-team"	---- Mandatory Field							
'								dicProcessViewAttributes.Add("State") = "Started"
'								dicProcessViewAttributes.Add("ResParty") = "AutoTestDBA (autotestdba)"
'								
'           					Call Fn_WorkflowViewer_Attributes(sAction, dicProcessViewAttributes)
' 
'History:
'          Developer Name   Date    Rev. No.   Changes Done   Reviewer 
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'           Prasanna    		11-Oct-2010   1.0               
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function Fn_WorkflowViewer_Attributes(sAction, dicProcessViewAttributes)
	GBL_FAILED_FUNCTION_NAME="Fn_WorkflowViewer_Attributes"
	 On Error Resume Next
	 Dim dicCount , dicKeys , dicItems
	 Dim iCounter, bReturn
	 Dim arrNodeList, iNodeCounter, arrNode, sExpnadNode
	 Dim iCounter1, sActionSelect
	 Dim objSelectType, intNoOfObjects, objDialog, iCounter2, sListHierarchy, arrQuorum
	
	 dicCount  = dicProcessViewAttributes.Count
	 dicItems = dicProcessViewAttributes.Items
	 dicKeys = dicProcessViewAttributes.Keys

   Select Case sAction         

   Case "Verify"

    For iCounter = 0 to dicCount - 1
          If  dicItems(iCounter) <> "" Then
			   Select Case dicKeys(iCounter)

				  Case "ProcessTree"
							arrNode = Split(dicItems(iCounter), ":", -1, 1)
							For iNodeCounter = 0 To UBound(arrNode)
							  If iNodeCounter = 0 Then
								   sExpnadNode = arrNode(iNodeCounter)
							  Else
								   sExpnadNode = sExpnadNode+":"+arrNode(iNodeCounter)
							  End If
							If iNodeCounter <>  UBound(arrNode) Then
								JavaWindow("WorkflowViewerWindow").JavaWindow("WEmbeddedFrame").JavaTree("ProcessTree").Expand sExpnadNode
								If Err.Number < 0 Then
									  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Expand Node ["+sExpnadNode+"] in Process Tree.")
									  Fn_WorkflowViewer_Attributes = False         
									  Exit Function
								Else
									  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Expand Node ["+sExpnadNode+"] in Process Tree.")
									  Call Fn_ReadyStatusSync(1)
									   Wait(1)
								End If
							End If
							Next

							'Select the Node    							
							bReturn = Fn_WrkflwViewr_ProcessTree_NodeOperations("Select",sExpnadNode)
							If bReturn = false Then
								  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Node ["+sExpnadNode+"] in Process Tree.")
								  Fn_WorkflowViewer_Attributes = False         
								  Exit Function
							Else                      
								  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Select Node ["+sExpnadNode+"] in Process Tree.")
								  Call Fn_ReadyStatusSync(1)
								   Wait(1)
							End If
		 			
							' Select the Attributes Dialog
							If JavaWindow("WorkflowViewerWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Attributes").Exist(2) = False  Then
								JavaWindow("WorkflowViewerWindow").JavaWindow("WEmbeddedFrame").JavaCheckBox("AttributesBtn").Set "ON"
									 If Err.Number < 0 Then
										   Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Click on Attributes Checkbox.")
										   Fn_WorkflowViewer_Attributes = False
										   Exit Function
									 Else                      
										   Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Clicked on Attributes Checkbox.")
										   Call Fn_ReadyStatusSync(2)
											Wait(2)
									 End If
							End If
		
							'Activate the Attributes Dialog
							If JavaWindow("WorkflowViewerWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Attributes").Exist(2) = True Then
								 JavaWindow("WorkflowViewerWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Attributes").Activate
								  If Err.Number < 0 Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Attribute Dialog Does not exist.")
										Fn_WorkflowViewer_Attributes = False
										Exit Function
								  Else
										wait(3)	
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully verified the existence of Attributes Dialog.")
								  End If
							End If

							'Click on Attributes text  
                            If JavaWindow("WorkflowViewerWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Attributes").Exist(3) = False Then 'Added by Nilesh on 13_Jun_12
								JavaWindow("WorkflowViewerWindow").JavaWindow("WEmbeddedFrame").JavaStaticText("Attributes").DblClick 1, 1, "LEFT"
								If Err.Number < 0 Then
										  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Click on Attributes Text.")
										  Fn_WorkflowViewer_Attributes = False         
										  Exit Function
								Else                      
										  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Clicked on Attributes Text.")
										  Wait(2)					
								End If
					         End If

				 Case "State"
						   Set objComp = JavaWindow("WorkflowViewerWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Attributes").JavaObject("State").Object.getComponent(0)
						   If objComp.getText() = dicItems(iCounter) Then
									 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : State Selection ["+dicItems(iCounter)+"] verified in Attributes Dialog.")
						   Else
									 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to verify the State Selected Text ["+dicItems(iCounter)+"] in Attributes Dialog. ")
									 Fn_WorkflowViewer_Attributes = False       
									 Exit Function
						   End If
						   Set objComp = Nothing
		
				 Case "ResParty"
		
						   Set objComp = JavaWindow("WorkflowViewerWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Attributes").JavaObject("ResponsibleParty").Object.getComponent(0)
						   If objComp.getText() = dicItems(iCounter) Then
									 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Responsible Party Selection ["+dicItems(iCounter)+"] verified in Attributes Dialog.")
						   Else
									 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to verify the Responsible Party Selected Text ["+dicItems(iCounter)+"] in Attributes Dialog. ")
									 Fn_WorkflowViewer_Attributes = False       
									 Exit Function
						   End If
						   Set objComp = Nothing
		
				 Case "NameACL"

				 Case "SignOffsQuorum"

				 Case "DueDate"

				 Case "Duration"

						   Set objComp = JavaWindow("WorkflowViewerWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Attributes").JavaEdit("Duration")
						   If objComp.GetROProperty("text") = dicItems(iCounter) Then
									 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Duration Text ["+dicItems(iCounter)+"] verified in Attributes Dialog.")
						   Else
									 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to verify the Duration Text ["+dicItems(iCounter)+"] in Attributes Dialog.")
									 Fn_WorkflowViewer_Attributes = False       
									 Exit Function
						   End If
						   Set objComp = Nothing
		
                 		 
       End Select

    End if

  Next

 End Select

' Close the Attributes Dialog
If JavaWindow("WorkflowViewerWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Attributes").Exist(2) = True  Then
	JavaWindow("WorkflowViewerWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Attributes").Close
		 If Err.Number < 0 Then
			   Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Close Attributes Dialog.")
			   Fn_WorkflowViewer_Attributes = False
			   Exit Function
		 Else                      
			   Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Closed Attributes Dialog.")
			   Call Fn_ReadyStatusSync(2)
				Wait(2)
		 End If
End If

Fn_WorkflowViewer_Attributes = True  
Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully verified Attributes Dialog.")
End Function


'*********************************************************  Function do Operation on Workflow Viewer Process Tree *********************************************************************

'Function Name		:	Fn_WrkflwViewr_PanelOperations

'Description		:	Action  performed :-
'						1. ProcessNodeDoubleClick																	
'						2. ProcessSelect
'                                                                   
'Parameters			:	1. sAction: Action to be performed
'						2.sProcessName: Process to select from Context Menu
' 												   

'Return Value		: 	True/False

'Pre-requisite		:	Process tree node should be selected 

'Examples			:	Fn_WrkflwViewr_PanelOperations("ProcessNodeDoubleClick","")
'						Fn_WrkflwViewr_PanelOperations("ProcessSelect","000380/A;1-ps1")

'History:
'					Developer Name		Date		Rev. No.	Changes Done																Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'					  Prasanna		13-Oct-2010	     1.0
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Modified 		:	  Madhura		7-July-2015		 1.1		Added new Case 	"OpenSubProcess" to select SubProcess in Workflow Viewer Perspective.
'Modified 		:	  Shweta R		12-Jan-2016		 1.2		Added new Case 	"ProcessSelectExt" 										[TC1122-20151116d-12_01_2016-VivekA-NewDevelopment]
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_WrkflwViewr_PanelOperations(sAction,sProcessName)
	GBL_FAILED_FUNCTION_NAME="Fn_WrkflwViewr_PanelOperations"
	On Error Resume Next
    Dim objDrpDwn, intNoOfObjects, objJavaObject, objWin, objDeviceReplay, objButton
    Dim h,w
    Dim iWidth, iHeight

    Select Case sAction

				Case "ProcessNodeDoubleClick"
						  JavaWindow("WorkflowViewerWindow").JavaWindow("WEmbeddedFrame").JavaObject("ProcessStart").DblClick 0,0
					   	   wait(2)
						   If Err.Number < 0  Then
							   Fn_WrkflwViewr_PanelOperations = false
							   Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Double Click on Process Node." )
							   Exit Function
						   Else
								Call Fn_ReadyStatusSync(3)   	
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Clicked on Process Node." )	
						   End If
				'[TC1122-20151116d-12_01_2016-VivekA-NewDevelopment] - Added bew Shweta R
				Case "ProcessSelect","ProcessSelectExt"		'Modified by Nilesh on 27 -Aug-2012
						If sAction = "ProcessSelectExt" Then
							JavaWindow("WorkflowViewerWindow").JavaWindow("WEmbeddedFrame").JavaStaticText("ProcessNameSText").SetTOProperty "label",sProcessName
							Wait 1
							Set objButton = JavaWindow("WorkflowViewerWindow").JavaWindow("WEmbeddedFrame").JavaButton("ProcessButtonSignOff")
						Else
							Set objButton = JavaWindow("WorkflowViewerWindow").JavaWindow("WEmbeddedFrame").JavaButton("ProcessButton")
						End If
						objButton.highlight
						Wait 1
						If objButton.Exist(5)=True Then	
							  Set objWin = JavaWindow("WorkflowViewerWindow").JavaWindow("WEmbeddedFrame")
							  If sAction = "ProcessSelectExt" Then
							  		objButton.Type micReturn
							  Else
							  		objButton.Type micReturn
							  End If						
							  Wait 5 'Added by Pritam Shikare on 26_Jun_2012
							  
							  Set objDrpDwn=description.Create()
							  
							  objDrpDwn("Class Name").value = "JavaStaticText"
							  objDrpDwn("label").value = sProcessName
							  objDrpDwn("displayed").value = 1
							  Set  intNoOfObjects =  objWin.ChildObjects(objDrpDwn)
							'Added by Nilesh on 29-Aug-12
                              h=intNoOfObjects(0).GetRoProperty("height")
							  w=intNoOfObjects(0).GetRoProperty("width")
							  intNoOfObjects(0).Click Cint(w/2),Cint(h/2)
							  
	
							  If Err.Number <> 0 Then
									Fn_WrkflwViewr_PanelOperations = False
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail selected the Process " + sProcessName) 
									Set objButton = Nothing
									Set objWin = nothing
									Set objDrpDwn = nothing
									Exit Function
							Else								
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Sucessfully selected the Process " + sProcessName) 
									Set objButton = Nothing
									Set objWin = nothing
									Set objDrpDwn = nothing
							End If
						Else
							Fn_WrkflwViewr_PanelOperations = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail get existence of process button")
							Set objButton = Nothing
							Set objWin = nothing
							Set objDrpDwn = nothing
							Exit Function
						End If	
				Case "OpenSubProcess"
						Set objJavaObject = JavaWindow("WorkflowViewerWindow").JavaWindow("WEmbeddedFrame").JavaObject("ProcessStart")				
						If objJavaObject.Exist(5) Then						
							Set objWin = JavaWindow("WorkflowViewerWindow").JavaWindow("WEmbeddedFrame")
			   				Set objDeviceReplay = CreateObject("Mercury.DeviceReplay")
	
							objDeviceReplay.MouseClick(objJavaObject.GetROProperty("abs_x") + 5), (objJavaObject.GetROProperty("abs_y") + 5), 0
							Wait 2 
								  
							Set objDrpDwn=description.Create()
								  
							objDrpDwn("Class Name").value = "JavaStaticText"
							objDrpDwn("label").value = sProcessName
	
							Set intNoOfObjects =  objWin.ChildObjects(objDrpDwn)
	
	                        iHeight=intNoOfObjects(0).GetRoProperty("height")
							iWidth=intNoOfObjects(0).GetRoProperty("width")
							intNoOfObjects(0).Click Cint(iWidth/2),Cint(iHeight/2)
									
						    If Err.Number <> 0 Then
								Fn_WrkflwViewr_PanelOperations = False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail to Select and open Subprocess " + sProcessName) 
								Set objWin = nothing
								Set objDrpDwn = nothing
								Set objDeviceReplay = nothing
								Set objJavaObject = nothing
								Set intNoOfObjects = nothing 
								Exit Function
						    Else								
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Sucessfully selected and opened SubProcess " + sProcessName) 
								Set objWin = nothing
								Set objDrpDwn = nothing
								Set objJavaObject = nothing
								Set objDeviceReplay = nothing
								Set intNoOfObjects = nothing 
							End If
						Else
							Fn_WrkflwViewr_PanelOperations = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail get existence of process Start Object.")
							Set objJavaObject = nothing
							Exit Function
						End If												
	End Select

	Fn_WrkflwViewr_PanelOperations = True

End Function

'*********************************************************  Function do Operation on Workflow Viewer Process Tree *********************************************************************

'Function Name		:					Fn_WrkflwViewr_ChangeToDesignModeConf

'Description			 :		 		    Verify Error Message
'                                                                   
'Parameters			   :	 			1. sObjName: Action to be performed
'												2.sErrMsg: Error Message
'												3.sBtnName: Button Name
' 
'Return Value		   : 			 True/False
'
'Pre-requisite			:		 	 Change to Design Mode Conf

'Examples				:			 Fn_WrkflwViewr_ChangeToDesignModeConf("Workflow Viewer - Change To Design Mode Confirmation","Message","Yes")
'
'History:
'										Developer Name				Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Mahendra Bhandarkar		17-Nov-2010	       1.0
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function Fn_WrkflwViewr_ChangeToDesignModeConf(sObjName, sErrMsg, sBtnName)
	GBL_FAILED_FUNCTION_NAME="Fn_WrkflwViewr_ChangeToDesignModeConf"
	Dim objDialog, sMsg

	Set objDialog = JavaDialog("WorkflowViewerChange")

	If Trim(sObjName) <> "" Then
			objDialog.SetToProperty "title", sObjName
	End If

	If objDialog.Exist(2) = True Then
			If Trim(sErrMsg) <> "" Then
				sMsg = objDialog.JavaObject("MLabel").Object.GetText
				If InStr(1, sMsg, sErrMsg, 1) > 0 Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Pass :  Message ["+CStr(sErrMsg)+"] Verified with ["+CStr(sMsg)+"].")
				Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),  "Fail : Message ["+CStr(sErrMsg)+"] not matched with ["+CStr(sMsg)+"].") 
						Fn_WrkflwViewr_ChangeToDesignModeConf = False
						Exit Function
				End If
			End If
			Wait(5)
			If Trim(sBtnName) <> "" Then
				objDialog.JavaButton(sBtnName).Click micLeftBtn
				Wait(5)
				If Err.Number < 0 Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),  "Fail : To Click on Button ["+CStr(sBtnName)+"].") 
						Fn_WrkflwViewr_ChangeToDesignModeConf = False
						Exit Function
				Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Pass :  Successfully Clicked on Button ["+CStr(sBtnName)+"].")
				End If
			End If
			Fn_WrkflwViewr_ChangeToDesignModeConf = True
	Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail :  ["+CStr(sObjName)+"] Dialog does not exist.")
						Fn_WrkflwViewr_ChangeToDesignModeConf = False						
	End If
	Wait(5)
	Set objDialog = Nothing

End Function


'*********************************************************		Function to Find / Verify the Audit Log Of Workflow Viewer		**********************************************************************
'Function Name		:				Fn_WrkflwViewr_AuditLogOperations(sAction,sObjID,sObjName,sObjRev,sObjTypeName,sObjSeqNum,sObjSecID,sObjSecName,sObjSecRev,sObjSecType,sObjSecSeqNum,sErrCode,sGroupName,sEventTypeName,sUserID,sDateCreBefore,sDateCreAfter,sClose)

'Description			 :		 		 To View the Audit Log of Workflow Viewer

'Parameters			   :	 			sAction,sObjID,sObjName,sObjRev,sObjTypeName,sObjSeqNum,sObjSecID,sObjSecName,sObjSecRev,sObjSecType,sObjSecSeqNum,sErrCode,sGroupName,sEventTypeName,sUserID,sDateCreBefore,sDateCreAfter,sClose
											
'Return Value		   : 				TRUE \ FALSE

'Pre-requisite			:		 		ADALicense prespective should be displayed.

'Examples				:				Case "Find" : Call Fn_WrkflwViewr_AuditLogOperations("Find","","samir123","","","","","","","","","","","Add Users/Groups","","","","no")
'													Case "Verify" : Call Fn_WrkflwViewr_AuditLogOperations("Verify","UserData","thosars","","","","","","","","","","","","","","","no")'													
'													Case "LoadAllVerify" : Call Fn_WrkflwViewr_AuditLogOperations("LoadAllVerify","","","","","","","","","","","","","","","","","yes")
'													Case "VerifyObject" : Call Fn_WrkflwViewr_AuditLogOperations("VerifyObject","PR-000001","ChgName","A","","","","","","","","","","","","","","")
'													Case "AdvancedFind" : Call Fn_WrkflwViewr_AuditLogOperations("AdvancedFind","","Test","","EPMTask","","","","","Type","","","","Update Process","","","","")

'History:
'										Developer Name			Date			Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Vidya Kulkarni
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_WrkflwViewr_AuditLogOperations(sAction,sObjID,sObjName,sObjRev,sObjTypeName,sObjSeqNum,sObjSecID,sObjSecName,sObjSecRev,sObjSecType,sObjSecSeqNum,sErrCode,sGroupName,sEventTypeName,sUserID,sDateCreBefore,sDateCreAfter,sClose)
	GBL_FAILED_FUNCTION_NAME="Fn_WrkflwViewr_AuditLogOperations"
	Dim objLicense, DateTime, iRowCount, iRow, sReturn, iColCount, sCol, iCol, ObjExport, aTableRecord, objExcel, iCount, sExceldata, iReturn, iCounter, bFlag
	ReDim aTableRecord (18)
    		Select Case sAction
				
			Case "Verify"
						Set objLicense = Fn_UI_ObjectCreate("Fn_WrkflwViewr__AuditLogOperations", JavaWindow("WorkflowViewerWindow").JavaWindow("QuickLinks").JavaDialog("Audit Log"))
						iRowCount = objLicense.JavaTable("LogTable").GetROProperty("rows")
						If iRowCount < 1 Then
							Fn_WrkflwViewr_AuditLogOperations = False
							Set objLicense = Nothing
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Audit Log Table does not exist")
							Exit Function
						End If
						iColCount = objLicense.JavaTable("LogTable").GetROProperty("cols")
						For iCol=0 to iColCount-1
							sCol = objLicense.JavaTable("LogTable").Object.getColumnPropertyName(iCol)
							If Trim(sCol)=Trim(sObjID) Then
								Exit For
							End If
						Next
						iRowCount = Fn_Table_GetRowCount("Fn_WrkflwViewr_AuditLogOperations",objLicense, "LogTable")
						For iRow=0 to iRowCount-1
							objLicense.JavaTable("LogTable").SelectRow iRow
							sReturn = objLicense.JavaTable("LogTable").GetCellData(iRow,iCol)
							If instr(1,sReturn,sObjName)<>0 Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sObjName &" Sucessfully found in row "&iRow)							
								Fn_WrkflwViewr_AuditLogOperations = TRUE
								If Trim(Lcase(sClose)) = "yes" Then
									objLicense.Close 
								End If
								Set objLicense = nothing 		
								Exit Function
							End If
							If iRow=iRowCount-1 Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sObjName &" not found in the Log Table")								
								Fn_WrkflwViewr_AuditLogOperations = False
								If Trim(Lcase(sClose)) = "yes" Then
									objLicense.Close 
								End If
								Set objLicense = nothing 		
								Exit Function
							End If
						Next
'			
'			
			Case "VerifyObject"
						'Click on View Audit Log button
						If Fn_UI_ObjectExist("Fn_WrkflwViewr_AuditLogOperations", JavaWindow("WorkflowViewerWindow").JavaWindow("QuickLinks").JavaDialog("Audit Log")) = False Then
									Call Fn_MenuOperation("Select","View:Audit:View Audit Logs")
									Call Fn_ReadyStatusSync(1)
						End If
						Set objLicense = Fn_UI_ObjectCreate("Fn_WrkflwViewr_AuditLogOperations", JavaWindow("WorkflowViewerWindow").JavaWindow("QuickLinks").JavaDialog("Audit Log"))
						bFlag = True
						'Set Object ID
						If Trim(sObjID) <> "" Then
							If Fn_UI_Object_GetROProperty("Fn_WrkflwViewr_AuditLogOperations",objLicense.JavaEdit("ObjectID"),"text") <> Trim(sObjID) Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Failed to verify Selected Object ID Matches with value ["+Trim(sObjID)+"] ")
								Fn_WrkflwViewr_AuditLogOperations = False
								bFlag = False
							Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Selected Object ID Matches with value ["+Trim(sObjID)+"] Verified Successfully.")
							End If
						End If
						'Set Object Name
						If sObjName<>"" Then
							If Fn_UI_ObjectExist("Fn_WrkflwViewr_AuditLogOperations", objLicense.JavaEdit("ObjectName")) = True Then
								If Fn_UI_Object_GetROProperty("Fn_WrkflwViewr_AuditLogOperations",objLicense.JavaEdit("ObjectName"),"text") <> Trim(sObjName) Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Failed to verify Selected Object Name Matches with value ["+Trim(sObjName)+"] ")
									Fn_WrkflwViewr_AuditLogOperations = False
									bFlag = False
								Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Selected Object Name Matches with value ["+Trim(sObjName)+"] Verified Successfully.")
								End If
							Else
								If Fn_UI_Object_GetROProperty("Fn_WrkflwViewr_AuditLogOperations",objLicense.JavaEdit("JobName"),"text") <> Trim(sObjName) Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Failed to verify Selected Job Name Matches with value ["+Trim(sObjName)+"] ")
									Fn_WrkflwViewr_AuditLogOperations = False
									bFlag = False
								Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Selected Job Name Matches with value ["+Trim(sObjName)+"] Verified Successfully.")
								End If
							End If
						End If
						'Set Object Revision
						If sObjRev<>"" Then
							If Fn_UI_Object_GetROProperty("Fn_WrkflwViewr_AuditLogOperations",objLicense.JavaEdit("ObjectRevision"),"text") <> Trim(sObjRev) Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Failed to verify Selected Object Revision Matches with value ["+Trim(sObjRev)+"] ")
								Fn_WrkflwViewr_AuditLogOperations = False
								bFlag = False
							Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Selected Object Revision Matches with value ["+Trim(sObjRev)+"] Verified Successfully.")
							End If
						End If
						'Set Object Type Name
						If sObjTypeName<>"" Then
							If Fn_UI_Object_GetROProperty("Fn_WrkflwViewr_AuditLogOperations",objLicense.JavaEdit("ObjectTypName"),"text") <> Trim(sObjTypeName) Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Failed to verify Selected Object Type Name Matches with value ["+Trim(sObjTypeName)+"] ")
								Fn_WrkflwViewr_AuditLogOperations = False
								bFlag = False
							Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Selected Object Type Name Matches with value ["+Trim(sObjTypeName)+"] Verified Successfully.")
							End If
						End If
						'Set Object Sequence Number
						If sObjSeqNum<>"" Then
							If Fn_UI_Object_GetROProperty("Fn_WrkflwViewr_AuditLogOperations",objLicense.JavaEdit("ObjectSeqNo"),"text") <> Trim(sObjSeqNum) Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Failed to verify Selected Object Sequence Number Matches with value ["+Trim(sObjSeqNum)+"] ")
								Fn_WrkflwViewr_AuditLogOperations = False
								bFlag = False
							Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Selected Object Sequence Number Matches with value ["+Trim(sObjSeqNum)+"] Verified Successfully.")
							End If
						End If
						'Set Secondary Object ID
						If sObjSecID<>"" Then
							If Fn_UI_Object_GetROProperty("Fn_WrkflwViewr_AuditLogOperations",objLicense.JavaEdit("SecondaryObjID"),"text") <> Trim(sObjSecID) Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Failed to verify Selected Secondary Object ID Matches with value ["+Trim(sObjSecID)+"] ")
								Fn_WrkflwViewr_AuditLogOperations = False
								bFlag = False
							Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Selected Secondary Object ID Matches with value ["+Trim(sObjSecID)+"] Verified Successfully.")
							End If
						End If
						'Set Secondary Object Name
						If sObjSecName<>"" Then
							If Fn_UI_Object_GetROProperty("Fn_WrkflwViewr_AuditLogOperations",objLicense.JavaEdit("SecondaryObjName"),"text") <> Trim(sObjSecName) Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Failed to verify Selected Secondary Object Name Matches with value ["+Trim(sObjSecName)+"] ")
								Fn_WrkflwViewr_AuditLogOperations = False
								bFlag = False
							Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Selected Secondary Object Name Matches with value ["+Trim(sObjSecName)+"] Verified Successfully.")
							End If
						End If
						'Set Secondary Object Revision
						If sObjSecRev<>"" Then
							If Fn_UI_Object_GetROProperty("Fn_WrkflwViewr_AuditLogOperations",objLicense.JavaEdit("SecondaryObjRev"),"text") <> Trim(sObjSecRev) Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Failed to verify Selected Secondary Object Revision Matches with value ["+Trim(sObjSecRev)+"] ")
								Fn_WrkflwViewr_AuditLogOperations = False
								bFlag = False
							Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Selected Secondary Object Revision Matches with value ["+Trim(sObjSecRev)+"] Verified Successfully.")
							End If
						End If
						'Set Secondary Object Type
						If sObjSecType<>"" Then
							If Fn_UI_Object_GetROProperty("Fn_WrkflwViewr_AuditLogOperations",objLicense.JavaEdit("SecondaryObjType"),"text") <> Trim(sObjSecType) Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Failed to verify Selected Secondary Object Type Matches with value ["+Trim(sObjSecType)+"] ")
								Fn_WrkflwViewr_AuditLogOperations = False
								bFlag = False
							Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Selected Secondary Object Type Matches with value ["+Trim(sObjSecType)+"] Verified Successfully.")
							End If
						End If
						'Set Secondary Object Sequence Number
						If sObjSecSeqNum<>"" Then
							If Fn_UI_Object_GetROProperty("Fn_WrkflwViewr_AuditLogOperations",objLicense.JavaEdit("SecondaryObjSeqNo"),"text") <> Trim(sObjSecSeqNum) Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Failed to verify Selected Secondary Object Sequence Number Matches with value ["+Trim(sObjSecSeqNum)+"] ")
								Fn_WrkflwViewr_AuditLogOperations = False
								bFlag = False
							Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Selected Secondary Object Sequence Number Matches with value ["+Trim(sObjSecSeqNum)+"] Verified Successfully.")
							End If
						End If
						'Set Error Code
						If sErrCode<>"" Then
							If Fn_UI_Object_GetROProperty("Fn_WrkflwViewr_AuditLogOperations",objLicense.JavaEdit("ErrorCode"),"text") <> Trim(sErrCode) Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Failed to verify Error Code Matches with value ["+Trim(sErrCode)+"] ")
								Fn_WrkflwViewr_AuditLogOperations = False
								bFlag = False
							Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Selected Error Code Matches with value ["+Trim(sErrCode)+"] Verified Successfully.")
							End If
						End If
						'Set Group Name
						If sGroupName<>"" Then
							If Fn_UI_Object_GetROProperty("Fn_WrkflwViewr_AuditLogOperations",objLicense.JavaEdit("GroupName"),"text") <> Trim(sGroupName) Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Failed to verify Group Name Matches with value ["+Trim(sGroupName)+"] ")
								Fn_WrkflwViewr_AuditLogOperations = False
								bFlag = False
							Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Selected Group Name Matches with value ["+Trim(sGroupName)+"] Verified Successfully.")
							End If
						End If
						'Select Event Type Name
						If sEventTypeName<>"" Then
							If Fn_UI_Object_GetROProperty("Fn_WrkflwViewr_AuditLogOperations",objLicense.JavaEdit("EventTypeName"),"text") <> Trim(sEventTypeName) Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Failed to verify Event Type Name Matches with value ["+Trim(sEventTypeName)+"] ")
								Fn_WrkflwViewr_AuditLogOperations = False
								bFlag = False
							Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Selected Event Type Name Matches with value ["+Trim(sEventTypeName)+"] Verified Successfully.")
							End If
						End If
						'Set User ID
						If sUserID<>"" Then
							If Fn_UI_Object_GetROProperty("Fn_WrkflwViewr_AuditLogOperations",objLicense.JavaEdit("UserID"),"text") <> Trim(sUserID) Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Failed to verify User ID Matches with value ["+Trim(sUserID)+"] ")
								FFn_WrkflwViewr_AuditLogOperations = False
								bFlag = False
							Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Selected User ID Matches with value ["+Trim(sUserID)+"] Verified Successfully.")
							End If
						End If
						'Set Date Created Before
						If sDateCreBefore<>"" Then
							If InStr(1, objLicense.JavaCheckBox("DateCreatedBefore").GetROProperty("label"), Trim(sDateCreBefore), 1) > 0 Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verified Date Created Before ["+CStr(sDateCreBefore)+"]")
								Call Fn_ReadyStatusSync(2)
							Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to verify Date Created Before ["+CStr(sDateCreBefore)+"]") 		
								Fn_WrkflwViewr_AuditLogOperations = False
								bFlag = False
							End If
						End If
						'Set Date Created After
						If sDateCreAfter<>"" Then
							If InStr(1, objLicense.JavaCheckBox("DateCreatedAfter").GetROProperty("label"), Trim(sDateCreAfter), 1) > 0 Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verified Date Created Before ["+CStr(sDateCreAfter)+"]")
								Call Fn_ReadyStatusSync(2)
							Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to verify Date Created Before ["+CStr(sDateCreAfter)+"]") 		
								Fn_WrkflwViewr_AuditLogOperations = False
								bFlag = False
							End If
						End If
						'Click on Find button
						Call Fn_Button_Click("Fn_WrkflwViewr_AuditLogOperations", objLicense, "Find")
						If Trim(Lcase(sClose)) = "yes" Then
							objLicense.Close 
						End If
'			
			Case "AdvancedFind"
						'Case for Advanced Tab Operation
						Set objLicense = Fn_UI_ObjectCreate("Fn_WrkflwViewr_AuditLogOperations", JavaWindow("WorkflowViewerWindow").JavaWindow("QuickLinks").JavaDialog("Audit Log"))
						objLicense.JavaTab("Tab").Select "Advanced"
						objLicense.RefreshObject
						'sObjID,sObjName,sObjRev,sObjTypeName,sObjSeqNum,sObjSecID,sObjSecName,sObjSecRev,sObjSecType,sObjSecSeqNum,sErrCode,sGroupName,sEventTypeName,sUserID,sDateCreBefore,sDateCreAfter,sClose
						If  Trim(sObjTypeName) <> "" Then
							Call Fn_List_Select("Fn_WrkflwViewr_AuditLogOperations", objLicense, "ObjTypeName",sObjTypeName)							
						End If
						If Trim(sEventTypeName) <> "" Then
							Call Fn_List_Select("Fn_WrkflwViewr_AuditLogOperations", objLicense, "EventTypeName",sEventTypeName)
						End If
						objLicense.RefreshObject
						If Trim(sObjName) <> "" Then
							Call Fn_Edit_Box("Fn_WrkflwViewr_AuditLogOperations", objLicense, "Name",sObjName)
						End If
						If Trim(sObjSecType) <> "" Then
							Call Fn_Edit_Box("Fn_WrkflwViewr_AuditLogOperations", objLicense, "TaskType",sObjSecType)
						End If
						'Click on Find button
						Call Fn_Button_Click("Fn_WrkflwViewr_AuditLogOperations", objLicense, "Find")
						' for Closing Dialog
						If Trim(Lcase(sClose)) = "yes" Then
							objLicense.Close 
						End If
'							
		End Select
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" Sucessfully completed of function Fn_WrkflwViewr_AuditLogOperations")
    Set objLicense = nothing 
	Fn_WrkflwViewr_AuditLogOperations = TRUE		
End Function

'*********************************************  Fn_WrkflwViewr_TaskActions **************************************************************

'Function Name		:		Fn_WrkflwViewr_TaskActions(sAction, sOption, sComment, sButton)

'Description		:		The function performs Actions on Workflow Viewer Nodes

'Parameters			:		sAction - Add/ Verify
'							sOption - Perform [Start/Complete/Suspend/Resume/Promote/Demote]
'							sComment - Entered Comment here
'							sButton - OK/Cancel/Clear
'
'Return Value		:		True/False

'Pre-requisite		:		WorkflowViewer Tree Opened and Node should be selected

'Examples			:		bReturn = Fn_WrkflwViewr_TaskActions("Add", "Complete", DataTable("Comment",dtGlobalSheet), "OK")
'
'History:
'								Developer Name			Date				Version			Build			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------
'								Radha Mane				3-Aug-2021	   		001				20210715
'-------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_WrkflwViewr_TaskActions(sAction, sOption, sComment, sButton)
	GBL_FAILED_FUNCTION_NAME="Fn_WrkflwViewr_TaskActions"
	Dim objDialog, bReturn, sMenu
	Select Case Trim(sOption)
		case "Start","Complete","Suspend","Resume","Promote","Demote"
			Set objDialog = JavaWindow("WorkflowViewerWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Demote Action Comments")
			objDialog.SetTOProperty "title",sOption+" Action Comments.*"
			sMenu = "Actions:"+sOption
	End Select
		
	If objDialog.Exist(2) = False Then
		bReturn = Fn_MenuOperation("Select", sMenu)
		Call Fn_ReadyStatusSync(3)
		If bReturn = False Then
			Fn_WrkflwViewr_TaskActions = False
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: To Invoke  Menu [" + sMenu+"]." )	
			Exit Function
		Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully Invoked  Menu [" + sMenu+"]." )	
		End If
	End If
		
	If objDialog.Exist Then
		Select Case Trim(sAction)
			Case "Add"
				If Trim(sComment) <> "" Then
					'Set the Comments
					objDialog.JavaEdit("Comments").Set Trim(sComment)
					If Err.Number < 0 Then
						Fn_WrkflwViewr_TaskActions = False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Set Comment [" + sComment+"].")	
						objDialog.JavaButton("Cancel").Click micLeftBtn
						Set objDialog = Nothing
						Exit Function
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Entered Comment [" + sComment+"].")
						Call Fn_ReadyStatusSync(2)
					End If
				End If

				'Click on button
				If Trim(sButton) <> "" Then
					objDialog.JavaButton(sButton).Click micLeftBtn
					If Err.Number < 0 Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Click on ["+CStr(sButton)+"] Button ") 		
						Fn_WrkflwViewr_TaskActions = False
						objDialog.JavaButton("Cancel").Click micLeftBtn
						Set objDialog = Nothing
						Exit Function		
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Clicked on ["+CStr(sButton)+"] Button")
						Call Fn_ReadyStatusSync(2)
					End If
				End If

				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Performed ["+CStr(sAction)+"] for Task [" + sWorkListNode+"]")

			Case "Verify"						' For feature use
			
		End Select
	Else
		Fn_WrkflwViewr_TaskActions = False
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: To Find ["+CStr(sAction)+" Action Comments] Dialog" )	
		Exit Function
	End If
	
	Fn_WrkflwViewr_TaskActions = True
	Set objDialog = Nothing
End Function

'-------------------------------------------------------------------Function Used to Verify ToolTip Text Of JavaStaticText-----------------------------------------------------------------------------------
'Function Name		:	Fn_WrkflwViewr_VerifyToolTipText

'Description		:	Function Used to Verify ToolTip Text Of JavaStaticText in Workflow Viewer flow

'Parameters			:	1.sTaskName:Task Name of Process
'						2.sExpectedToolTipText:Expected Tool Tip Text

'Return Value		: 	True Or False

'Pre-requisite		:	Workflow Viewer Process Flow should be opened

'Examples			:	bReturn = Fn_WrkflwViewr_VerifyToolTipText("New Do Task 1","Task State: Started")
'						bReturn = Fn_WrkflwViewr_VerifyToolTipText("New Do Task 1","Responsible Party: AutoTest3 (autotest3)")

'History			:			
'						Developer Name			Date		Version		Build			Reviewer
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'						Radha Mane			17/12/2010		001			20210715
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_WrkflwViewr_VerifyToolTipText(sTaskName,sExpectedToolTipText)
	GBL_FAILED_FUNCTION_NAME="Fn_WrkflwViewr_VerifyToolTipText"
    Dim sCurrentToolTip,objJavaWindow,aPropName,bFlag
    bFlag = False
	Fn_WrkflwViewr_VerifyToolTipText=False
	
	Set objJavaWindow = JavaWindow("WorkflowViewerWindow").JavaWindow("WEmbeddedFrame")
	Set objJWin = JavaWindow("WorkflowViewerWindow").JavaWindow("WEmbeddedFrame").JavaStaticText("ToolTipIcon")
	objJWin.SetTOProperty "tagname",sTaskName+"(st)"
		
	aPropName = Split(sExpectedToolTipText,":")
	If aPropName(0) = "Task State" Then
		bFlag = Fn_UI_JavaStaticText_SetTOProperty("Fn_WrkflwViewr_VerifyToolTipText",objJavaWindow,"ToolTipIcon","index","0")
	Else	''This is for Responsible Party:
		bFlag = Fn_UI_JavaStaticText_SetTOProperty("Fn_WrkflwViewr_VerifyToolTipText",objJavaWindow,"ToolTipIcon","index","1")
	End If
	
	If bFlag = True Then
		objJWin.Object.focusable(True)
		sCurrentToolTip=objJWin.Object.getToolTipText()
		If LCase(Trim(sCurrentToolTip))=LCase(Trim(sExpectedToolTipText)) Then
			Fn_WrkflwViewr_VerifyToolTipText=True
		Else
			Fn_WrkflwViewr_VerifyToolTipText=False
		End If
	Else 
		Fn_WrkflwViewr_VerifyToolTipText=False
	End If
	Set objJavaWindow = Nothing
End Function
