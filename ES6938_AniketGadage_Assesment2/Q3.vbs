'3. (10 Marks)
'arrFirstArray = Array(1,2,3,4,5)
'arrSecondArray = Array(3,4,5,6,7)
'write a vb script to get common values from both the arrays 
'Expected O/p - 3,4,5

Option Explicit

Dim First_Arr,Second_Arr,Result_Arr()
Dim int_i,int_j,int_k

int_k=0

First_Arr=Array(1,2,3,4,5)
Second_Arr=Array(3,4,5,6,7)

For int_i=0 to UBound(First_Arr)
	For int_j=0 to UBound(Second_Arr)
			If First_Arr(int_i)=Second_Arr(int_j) Then
				Redim Preserve Result_Arr(int_k)
				Result_Arr(int_k)=First_Arr(int_i)
				int_k=int_k+1
			end If
	Next
Next

MsgBox Join(Result_Arr,",")
