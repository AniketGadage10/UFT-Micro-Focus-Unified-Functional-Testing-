Option  Explicit


Dim FirstNum, SecondNum
Dim Action
Dim Custom_Error
Public Function Calculator( Action, FirstNum, SecondNum ) 
	FirstNum=Int(FirstNum)
	SecondNum=Int(SecondNum)	
	On Error Resume Next
	Dim Answer
	
	Select Case Lcase(Action)
		Case  "add"
			Answer=FirstNum+SecondNum
		Case "sub"
			Answer=FirstNum-SecondNum
		Case  "mult"
			Answer=FirstNum*SecondNum
		Case "div"
			Answer=(FirstNum/SecondNum)
		Case Else
			Answer=0
	
	End Select
	
	If Err.number =0  Then
		Calculator=Answer
	Else
		 Custom_Error="Error Number : "&Err.Number & " Error Description : "& Err.Description
		Calculator =Custom_Error
	End If
	
	
	On Error Goto 0
	
End  Function

