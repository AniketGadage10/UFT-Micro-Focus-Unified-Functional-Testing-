'4. (10 Marks)
'Write a program  to print below pattern with help of mininmum possible loops 
			
'			03
'		04	06	08
'	04	08	12	16	20



Option Explicit


Dim Row_count,col_count

Dim Int_i,Int_j,space1,ele_count
Dim Str


Str=""

Row_count=3
col_count=Row_count+2
space1=0
ele_count=1

For Int_i =1 TO Row_count
	Space1=Row_count-Int_i
	
	str=Str+Space(Space1)

	space1=Space1+1
	
	For int_j=1 to ele_count
			
		str=str+cstr(Space1*Int_i)+" "
		Space1 =Space1+1
	Next
	Str=Str&vbNewLine
	ele_count=ele_count+2	

Next

MsgBox Str