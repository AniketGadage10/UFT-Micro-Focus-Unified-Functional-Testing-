
Option Explicit

Dim uftobj,libobj,repoobj

Set uftobj=CreateObject("QuickTest.Application")

If uftobj.Launched=False Then
	uftobj.Launch
	uftobj.visible=True
	
	uftobj.open "C:\Vrushali\GUITest1",True,False
	

	
	set libobj=uftobj.Test.Settings.Resources.Libraries
	msgbox libobj.Find("C:\Users\agadage\Desktop\UFT-Micro-Focus-Unified-Functional-Testing-\ES6938_ANIKETGADAGE_(LibraryFile_Function)\LibraryFunction\Arithmetic_Calculator_Function.vbs")
	if libobj.Find("C:\Users\agadage\Desktop\UFT-Micro-Focus-Unified-Functional-Testing-\ES6938_ANIKETGADAGE_(LibraryFile_Function)\LibraryFunction\Arithmetic_Calculator_Function.vbs")=-1 Then
		libobj.add "C:\Users\agadage\Desktop\UFT-Micro-Focus-Unified-Functional-Testing-\ES6938_ANIKETGADAGE_(LibraryFile_Function)\LibraryFunction\Arithmetic_Calculator_Function.vbs",1
	end if
	
	Set repoobj=uftobj.test.Actions("Action1").ObjectRepositories
	
	msgbox repoobj.find("C:\Users\agadage\Desktop\UFT-Micro-Focus-Unified-Functional-Testing-\ES6938_ANIKETGADAGE_(ObjectRepo)\SharedRepo\Shared_Repo1.tsr")
	
	if repoobj.find("C:\Users\agadage\Desktop\UFT-Micro-Focus-Unified-Functional-Testing-\ES6938_ANIKETGADAGE_(ObjectRepo)\SharedRepo\Shared_Repo1.tsr")=-1 Then
	repoobj.add "C:\Users\agadage\Desktop\UFT-Micro-Focus-Unified-Functional-Testing-\ES6938_ANIKETGADAGE_(ObjectRepo)\SharedRepo\Shared_Repo1.tsr",1
	end if
	
	msgbox "dada"
		uftobj.test.save
		
	uftobj.test.run
	
	MsgBox uftobj.test.lastrunresults.status
Else

	MsgBox "UFT Launched Fail"

End If

uftobj.quit
