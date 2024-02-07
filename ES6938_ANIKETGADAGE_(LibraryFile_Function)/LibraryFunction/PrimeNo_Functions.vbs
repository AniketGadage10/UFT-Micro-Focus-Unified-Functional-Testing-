
Public Function GetPrimeNumber_In_Range(Start_No,End_No)
	
	Dim Iterator_i
	Dim Prime_Arr()
	Dim Counter:Counter=0
	For Iterator_i = Start_No To End_No Step 1
		ReDim Preserve Prime_Arr(Counter)
		IF PrimeOrNot(Iterator_i) Then
			Prime_Arr(Counter)=Iterator_i
			Counter=Counter+1
		End If
	Next
	GetPrimeNumber_In_Range=Join(Prime_Arr," , ")
End Function

Public Function PrimeOrNot( INum )
	Dim Iterator_j
	Dim Flag:Flag=True
	
	For Iterator_j = 2 To INum-1 Step 1
		If (INum Mod Iterator_j) = 0 Then
			Flag=False
		End If
	Next
	PrimeOrNot=Flag
End Function

