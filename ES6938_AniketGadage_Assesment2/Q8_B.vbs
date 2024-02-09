'5 Marks'
'B. Calculate length of the string without using inbuilt function



Option Explicit

Dim STR
Dim int_i,length_count
Dim ch
STR="Expleo office Pune."
length_count=0
	for int_i=1 to 100
	
		ch=Mid(STR,int_i,1)
		
		if ch="" Then
			Exit for
		end if
		length_count=length_count+1
	Next
	
	MsgBox "length Of String = "&length_count