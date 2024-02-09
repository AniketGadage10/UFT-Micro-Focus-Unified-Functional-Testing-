Option Explicit
'*********************************************************	Function List		***********************************************************************
'1.  Fn_CAEManager_AnalysisTree_Operation
'****************************************************************************************************************************************************
''*********************************************************		Function to action perform on NavTree	***********************************************************************
'Function Name		:				Fn_CAEManager_AnalysisTree_Operation

'Description			 :		 		 Actions performed in this function are:
'																	1. Node Select

'Parameters			   :	 			1. StrAction: Action to be performed
'													2. StrNodeName: Fully qulified tree Path (delimiter as ':') [multiple node are separated by "," ] 
'												   3. StrMenu: Context menu to be selected

'Return Value		   : 				TRUE \ FALSE

'Pre-requisite			:		 		CAEManager prespective should be open with Analysis Tab activated.

'Examples				:			Call  Fn_CAEManager_AnalysisTree_Operation("NodeExist","CAE Analysis Item Revisions:000023/A;1-CAEItem","")

'History					 :		
'													Developer Name				Date						Rev. No.			Changes Done						Reviewer
'												------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'												Harshal Agrawal				06July2011					1.0						Developed							Amol Lanke
'												------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'MODIFIED BY    : 		       Pallavi Patil                     09 July 2012               
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'MODIFIED BY    : 		       Ankit Nigam									26 June 2015 				Modified case "NodeExist" as per design change.
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Function Fn_CAEManager_AnalysisTree_Operation(StrAction,sNodeNameWithPath,StrMenu)
		GBL_FAILED_FUNCTION_NAME="Fn_CAEManager_AnalysisTree_Operation"
		Dim objAnalysisTree, objAnalysisTree1
		Dim sNode, iCount, iCount1, iCount2, sTreeItem, sTreeItem1, intNodeCount
			Set objAnalysisTree = Fn_UI_ObjectCreate("Fn_CAEManager_AnalysisTree_Operation",JavaWindow("CAE Manager").JavaTree("AnalysisTree"))
		Select Case StrAction
			Case "NodeExist"
				sNode = Split(sNodeNameWithPath,":")
				
				For iCount = 0 To UBound(sNode)

					If iCount = 0 Then
						sTreeItem = objAnalysisTree.Object.GetItem(0).getData().getValue().tostring()
						
						If Trim(Lcase(sTreeItem)) = Trim(Lcase(sNodeNameWithPath)) Then
							Fn_CAEManager_AnalysisTree_Operation = True
							Exit Function
						Else
							Set objAnalysisTree1 = objAnalysisTree.Object.GetItem(0)
						End If
					Else 
						intNodeCount = objAnalysisTree1.getItemCount()
						iCount2=0
						For iCount1 = 0 To intNodeCount - 1
						
							sTreeItem1 =  objAnalysisTree1.GetItem(iCount2).getData().tostring()
							
							If Trim(Lcase(sTreeItem1)) = Trim(Lcase(sNode(iCount))) Then
								sTreeItem = sTreeItem + ":" + sTreeItem1
								If Trim(Lcase(sTreeItem)) = Trim(Lcase(sNodeNameWithPath)) Then
									Fn_CAEManager_AnalysisTree_Operation = True
									Exit Function
									
								Else
									Set objAnalysisTree1 = objAnalysisTree1.GetItem(iCount2)
									Fn_CAEManager_AnalysisTree_Operation = False
									iCount2=0
									Exit For
								End If	
							Else
								iCount2=iCount2+1
							End If
						Next
					End If
			 	Next
	 	End Select	
		'Fn_CAEManager_AnalysisTree_Operation = Fn_UI_JavaTree_NodeExist("Fn_CAEManager_AnalysisTree_Operation",objAnalysisTree,sNodeNameWithPath)
		Set objAnalysisTree = Nothing
End Function
