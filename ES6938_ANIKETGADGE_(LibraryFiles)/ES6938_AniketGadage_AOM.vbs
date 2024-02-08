
Option Explicit

	Dim File_obj,Read_obj,Write_obj
	Dim src,result
	Dim Tst_name
	Dim uft_obj
	
	Const iRead=1
	Const iWrite=8
	
	src="C:\Users\agadage\Desktop\UFT-Micro-Focus-Unified-Functional-Testing-\ES6938_ANIKETGADGE_(LibraryFiles)\Test_Script_List.txt"
	result="C:\Users\agadage\Desktop\UFT-Micro-Focus-Unified-Functional-Testing-\ES6938_ANIKETGADGE_(LibraryFiles)\result.txt"
	
	'Creating File Object
	set File_obj=CreateObject("Scripting.FileSystemObject")
	
	'Checking Existance And Creating result File
	If Not(File_obj.FileExists(result)) Then
		File_Obj.CreateTextFile result
	End If
	
	'Creating Write Object
	
	Set Write_obj=File_obj.OpenTextFile(result,iWrite)
	
	'Checking Existance of Test_Script_List
	If File_obj.FileExists(src) Then
	'Creating Read Object
		
		Set Read_obj=File_obj.OpenTextFile(src,iRead)
		Write_obj.WriteLine("----------------------------------"&Now()&"----------------------------------")
		Do Until Read_obj.AtEndOfStream
		
		Tst_name=Read_obj.ReadLine()
		
		
		Set uft_obj=CreateObject("QuickTest.Application")
		
		If uft_obj.Launched=False Then
			uft_obj.Launch=True
		End If
		
		uft_obj.Visible=True
		
		LauchTest(Tst_name)
		
		Loop
	Else
		Write_obj.WriteLine("----------------------------------"&Now()&"----------------------------------")
		Write_obj.WriteLine("Fail To Found File "&src)
	End If
	
	uft_obj.Quit
	'Write_obj.close
	'Read_obj.close
	
	Set uft_obj=Nothing
	Set File_Obj=Nothing
	Set Write_obj=Nothing
	Set Read_obj=Nothing
	
	
	
	
	
	
	Public Function LauchTest(Tst_name)
		uft_obj.Open Tst_name,True,False
		uft_obj.Test.save
		uft_obj.Test.Run
		Write_obj.WriteLine(Mid(Tst_name,114,Len(Tst_name))&" : "&uft_obj.Test.LastRunResults.Status)		
	End Function

