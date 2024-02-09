'2. (10 Marks)
'Prerquisite - "c:\MyFolder\MyDetails.txt" file is present
'Line 1 : Hi am Sandeep 
'Line 2 : I have total 4+ Years of experience in Manual testing
'Line 3 : I have handson experience of Retail and PLM domain
'Line 4 : I am associated with Expleo since past 2 years
'Line 5 : I am learning VB scripting
'write a vb script which will read above file and paste the content of this file in to other file but order of sentence will be reversed
'(last sentence should occure as first sentence in 2nd file,second last sentence at 2nd line... )
'Expected O/p "c:\MyFolder\CopyMyDetails.txt" with below lines
'Line 1 : I am learning VB scripting
'Line 2 : I am associated with Expleo since past 2 years
'Line 3 : I have handson experience of Retail and PLM domain
'Line 4 : I have total 4+ Years of experience in Manual testing
'Line 5 : Hi am Sandeep 

Option Explicit

Dim File_Obj,Write_Obj,Read_Obj
Dim src,dest
Dim ForRead,ForWrite
Dim Str
Dim arr(4)
Dim int_i:int_i=0

str=""

ForRead=1
ForWrite=2

src="c:\MyFolder\MyDetails.txt"
dest="c:\MyFolder\CopyMyDetails.txt"

Set File_Obj=CreateObject("scripting.FileSystemObject")


IF Not(File_Obj.FileExists(src)) Then
	File_Obj.CreateTextFile src
end IF

Set Read_Obj=File_Obj.OpenTextFile(src,ForRead)

IF Not(File_Obj.FileExists(dest)) Then
	File_Obj.CreateTextFile dest
end IF

Set Write_Obj=File_Obj.OpenTextFile(dest,ForWrite)

Do Until Read_Obj.AtEndOfStream

	arr(int_i)=Read_Obj.ReadLine()
	int_i=int_i+1
Loop

For int_i=UBound(arr) to 0 step -1
	Write_Obj.WriteLine(arr(int_i))
Next

set File_Obj=Nothing
Set Write_Obj=Nothing
sET Read_Obj=Nothing





