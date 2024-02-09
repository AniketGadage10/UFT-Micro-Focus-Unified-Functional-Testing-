'10. (10 Marks)
' a program to count the occurence of perticular sentence after 35th line in given text file 



Dim File_Obj,Read_Obj
Dim src
Dim ForRead,Search_Count,Position
Dim Str,Search_Text
Dim int_i:int_i=0

str=""

ForRead=1
Search_Count=0
Position=1
src="c:\MyFolder\sample.txt"

Set File_Obj=CreateObject("scripting.FileSystemObject")


IF Not(File_Obj.FileExists(src)) Then
	File_Obj.CreateTextFile src
end IF

Set Read_Obj=File_Obj.OpenTextFile(src,ForRead)

str=Read_Obj.Readall

Search_Text=InputBox("Enter the Search_Text")

For int_i=10 to Len(str)
	ch=Mid(str,int_i,Len(Search_Text))
	
	IF ch=Search_Text Then
		Search_Count=Search_Count+1
	end if	
Next

MsgBox Search_Count

set File_Obj=Nothing
Set Write_Obj=Nothing
sET Read_Obj=Nothing





