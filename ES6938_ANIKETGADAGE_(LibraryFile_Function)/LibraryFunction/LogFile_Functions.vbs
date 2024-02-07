Option Explicit

Dim  File_Obj,Write_Obj
Dim Src
Const For_Append=8

Public Function CreateLogFile( Src)	
	
	Set File_Obj=CreateObject("Scripting.FileSystemObject")
	print Src
	If Not(File_Obj.FileExists(Src)) Then
		File_Obj.CreateTextFile Src
	End If
	
	Set Write_Obj= File_Obj.OpenTextFile(Src,For_Append)
	
	Write_Obj.WriteLine(" UserName = "&Environment.Value("UserName"))
	Write_Obj.WriteLine("Test Case Name = "&Environment.Value("TestName"))
	Write_Obj.WriteLine("Test Case Directory = "&Environment.Value("TestDir"))

End Function
